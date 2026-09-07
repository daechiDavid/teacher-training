import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { registerHooks } from 'node:module';
import test from 'node:test';
import ts from 'typescript';

// Load the production TypeScript functions without a separate test dependency.
registerHooks({
  resolve(specifier, context, nextResolve) {
    if (context.parentURL?.endsWith('.ts') && specifier.startsWith('.') && !/\.[a-z]+$/i.test(specifier)) {
      return nextResolve(`${specifier}.ts`, context);
    }
    return nextResolve(specifier, context);
  },
  load(url, context, nextLoad) {
    if (!url.endsWith('.ts')) return nextLoad(url, context);
    return {
      format: 'module', shortCircuit: true,
      source: ts.transpileModule(readFileSync(new URL(url), 'utf8'), {
        compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.ESNext },
      }).outputText,
    };
  },
});
const {
  buildDefaultCaptureRows, updateCaptureResult, applyZoomChatAttendanceText,
  applyRecognizedZoomRowsToCaptureRows: applyZoom, buildSummaryRows,
  calculateEffectiveZoomMinutes,
} = await import('../src/lib/attendanceDocuments.ts');
const people = ['김가람', '이보람', '박다솜'].map((name, index) => ({
  sequence: String(index + 1), name, niceNumber: '없음', schoolName: '테스트학교', source: {},
}));
const zoom = (results) => people.map((person, index) => ({ ...person, result: results[index] }));
const marks = (rows) => rows.map(({ period1, period2, result }) => [period1, period2, result]);

test('2번 인정만 자동 O/O, 미인정은 채팅 결과 유지', async () => {
  const chat = new File([
    '00:00:00\t보조강사\t19:00\n00:05:00\t이보람\t출석\n01:50:00\t박다솜\t출석\n02:00:00\t보조강사\t종료',
  ], 'chat.txt');
  const applied = await applyZoomChatAttendanceText(chat, buildDefaultCaptureRows(people), {
    startTime: '19:00', endTime: '21:00',
  });
  const before = structuredClone(applied.rows);
  const zoomRows = zoom(['인정', '미인정', '미인정']);
  assert.deepEqual(marks(applyZoom(applied.rows, zoomRows)), [
    ['O', 'O', '인정'], ['O', 'X', '미인정'], ['X', 'O', '미인정'],
  ]);
  assert.deepEqual(applied.rows, before);
  assert.deepEqual(buildSummaryRows(applied.rows, zoomRows).map((r) => [r.result1, r.result2, r.result3]), [
    ['인정', '인정', '이수'], ['미인정', '미인정', '미이수'], ['미인정', '미인정', '미이수'],
  ]);
});

test('인정 후 미인정으로 수정하거나 파일을 교체하면 원래 출결 복원', () => {
  const base = buildDefaultCaptureRows(people);
  base[1] = updateCaptureResult({ ...base[1], period1: 'O' });
  base[2] = updateCaptureResult({ ...base[2], period1: 'O', period2: 'O' });
  const original = structuredClone(base);
  for (let visit = 0; visit < 3; visit++) {
    assert.ok(applyZoom(base, zoom(['인정', '인정', '인정'])).every((row) => row.result === '인정'));
    assert.deepEqual(marks(applyZoom(base, zoom(['미인정', '미인정', '미인정']))), marks(original));
  }
  assert.deepEqual(applyZoom(base, []), original);
  assert.deepEqual(base, original);
});

test('2번 인정 중 입력한 채팅도 보존하고 미인정으로 바뀌면 반영', async () => {
  const base = buildDefaultCaptureRows(people);
  applyZoom(base, zoom(['인정', '인정', '인정']));
  const chat = new File(['00:00:00\t보조강사\t19:00\n00:05:00\t김가람\t출석\n02:00:00\t보조강사\t종료'], 'chat.txt');
  const applied = await applyZoomChatAttendanceText(chat, base, { startTime: '19:00', endTime: '21:00' });
  assert.equal(applied.startMatches, 1);
  assert.equal(applied.endMatches, 0);
  assert.deepEqual(marks(applyZoom(applied.rows, zoom(['미인정', '인정', '미인정']))), [
    ['O', 'X', '미인정'], ['O', 'O', '인정'], ['X', 'X', '미인정'],
  ]);
});

test('접속시간은 쉬는시간 10분만 제외하고 연수 구간 안에서 계산', () => {
  const start = 20 * 60;
  const end = 21 * 60 + 50;
  assert.equal(calculateEffectiveZoomMinutes(start, start + 50, start, end), 50);
  assert.equal(calculateEffectiveZoomMinutes(start, end, start, end), 100);
  assert.equal(calculateEffectiveZoomMinutes(start + 50, start + 60, start, end), 0);
  assert.equal(calculateEffectiveZoomMinutes(start + 55, start + 80, start, end), 20);
  assert.equal(calculateEffectiveZoomMinutes(start - 30, start + 20, start, end), 20);
});
