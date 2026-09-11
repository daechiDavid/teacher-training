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
  calculateEffectiveZoomMinutes, calculateMergedEffectiveZoomMinutes,
  calculateEvidenceCellLayout,
  createCaptureHwpx,
  parseZoomAttendanceWorkbook,
} = await import('../src/lib/attendanceDocuments.ts');
const people = ['김가람', '이보람', '박다솜'].map((name, index) => ({
  sequence: String(index + 1), name, niceNumber: '없음', schoolName: '테스트학교', source: {},
}));
const zoom = (results) => people.map((person, index) => ({ ...person, result: results[index] }));
const marks = (rows) => rows.map(({ period1, period2, result }) => [period1, period2, result]);

test('캠 이미지 수에 맞춰 증빙 칸을 균등 분할', () => {
  for (const count of [1, 2, 3, 4]) {
    const layout = calculateEvidenceCellLayout(count);
    assert.equal(layout.length, count);
    assert.equal(layout.reduce((sum, cell) => sum + cell.colSpan, 0), 11);
    assert.equal(layout.reduce((sum, cell) => sum + cell.cellWidth, 0), 32113);
    assert.ok(Math.max(...layout.map((cell) => cell.cellWidth)) - Math.min(...layout.map((cell) => cell.cellWidth)) <= 1);
    layout.forEach((cell, index) => {
      if (index > 0) {
        assert.equal(cell.colAddr, layout[index - 1].colAddr + layout[index - 1].colSpan);
      }
    });
  }
});

test('캠 이미지 4장이 포함된 문서1 HWPX 생성', async () => {
  const originalFetch = globalThis.fetch;
  const originalImage = globalThis.Image;
  const template = readFileSync(new URL('../public/templates/attendance_capture.hwpx', import.meta.url));
  globalThis.fetch = async () => new Response(template);
  globalThis.Image = class MockImage {
    naturalWidth = 1600;
    naturalHeight = 900;
    onload;
    set src(_value) {
      queueMicrotask(() => this.onload?.());
    }
  };
  try {
    const image = (index) => ({
      name: `camera-${index}.png`,
      dataUrl: 'data:image/png;base64,iVBORw0KGgo=',
    });
    const document = await createCaptureHwpx({
      trainingName: '테스트 연수',
      institute: '테스트 기관',
      trainingDate: '2026-09-11',
      startTime: '09:00',
      endTime: '10:50',
      period1Label: '1교시',
      period2Label: '2교시',
      instructorName: '테스트 강사',
    }, [], [{
      id: 'camera-row',
      period: 1,
      mode: 'camera',
      cameraImages: [1, 2, 3, 4].map(image),
      chatImages: [],
    }]);
    assert.equal(document.type, 'application/x-hwpml-package');
    assert.ok(document.size > 0);
  } finally {
    globalThis.fetch = originalFetch;
    globalThis.Image = originalImage;
  }
});

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

test('다중 기기의 겹치는 접속 구간은 한 번만 계산', () => {
  const start = 8 * 60;
  const end = 9 * 60 + 50;
  assert.equal(calculateMergedEffectiveZoomMinutes([
    [start, start + 30],
    [start + 20, start + 50],
  ], start, end), 50);
  assert.equal(calculateMergedEffectiveZoomMinutes([
    [start, start + 50],
    [start + 10, start + 20],
  ], start, end), 50);
  assert.equal(calculateMergedEffectiveZoomMinutes([
    [start, start + 20],
    [start + 30, start + 50],
  ], start, end), 40);
});

test('병합된 접속 구간에서도 쉬는시간은 한 번만 제외', () => {
  const start = 20 * 60;
  const end = 21 * 60 + 50;
  assert.equal(calculateMergedEffectiveZoomMinutes([
    [start + 40, start + 70],
    [start + 45, start + 65],
  ], start, end), 20);
});

test('자정을 넘는 다중 접속 구간도 중복 없이 계산', () => {
  const start = 23 * 60 + 30;
  const end = 1 * 60 + 20;
  assert.equal(calculateMergedEffectiveZoomMinutes([
    [23 * 60 + 40, 20],
    [23 * 60 + 50, 40],
  ], start, end), 50);
});

test('CSV 처리 경로에서 겹치는 기기 접속을 병합하고 경고 표시', async () => {
  const headers = Array.from({ length: 24 }, (_, index) => `열${index + 1}`);
  headers[16] = '이름(원래 이름)';
  headers[18] = '참가 시간';
  headers[19] = '나간 시간';
  headers[20] = '기간(분)';
  const makeRow = (entry, exit) => {
    const row = Array(24).fill('');
    row[16] = '홍길동_테스트학교';
    row[18] = entry;
    row[19] = exit;
    row[20] = '30';
    return row;
  };
  const csv = [
    headers,
    makeRow('2026/09/10 08:00:00 AM', '2026/09/10 08:30:00 AM'),
    makeRow('2026/09/10 08:20:00 AM', '2026/09/10 08:50:00 AM'),
  ].map((row) => row.join(',')).join('\n');
  const file = new File([csv], 'zoom.csv', { type: 'text/csv' });
  const [row] = await parseZoomAttendanceWorkbook(file, [{
    sequence: '1',
    name: '홍길동',
    niceNumber: '없음',
    schoolName: '테스트학교',
    source: { school: '테스트학교' },
  }], { startTime: '08:00', endTime: '09:50' });

  assert.equal(row.rawMinutes, 60);
  assert.equal(row.effectiveMinutes, 50);
  assert.equal(row.result, '미인정');
  assert.equal(row.warning, '중복 접속 10분 제외 · 인정시간 80분 미만 확인 필요');
});

test('CSV 학교명 축약형과 병설유치원 표기를 정식 학교명으로 자동 매칭', async () => {
  const headers = Array.from({ length: 24 }, (_, index) => `열${index + 1}`);
  headers[16] = '이름(원래 이름)';
  headers[18] = '참가 시간';
  headers[19] = '나간 시간';
  headers[20] = '기간(분)';
  const zoomNames = [
    '배새하_세종캠고',
    '안현수_서울화계병유',
    '이문숙_현민초병설유',
  ];
  const peopleWithOfficialSchools = [
    ['배새하', '세종컴퍼스고등학교'],
    ['안현수', '서울화계초등학교병설유치원'],
    ['이문숙', '현민초등학교병설유치원'],
  ].map(([name, school], index) => ({
    sequence: String(index + 1),
    name,
    niceNumber: '없음',
    schoolName: school,
    source: { school },
  }));
  const makeRow = (zoomName) => {
    const row = Array(24).fill('');
    row[16] = zoomName;
    row[18] = '2026/09/10 08:00:00 AM';
    row[19] = '2026/09/10 09:50:00 AM';
    row[20] = '110';
    return row;
  };
  const csv = [headers, ...zoomNames.map(makeRow)].map((row) => row.join(',')).join('\n');
  const rows = await parseZoomAttendanceWorkbook(
    new File([csv], 'zoom.csv', { type: 'text/csv' }),
    peopleWithOfficialSchools,
    { startTime: '08:00', endTime: '09:50' },
  );

  assert.deepEqual(rows.map((row) => row.result), ['인정', '인정', '인정']);
  assert.deepEqual(rows.map((row) => row.entryTime), ['08:00', '08:00', '08:00']);

  const abbreviatedKindergartenCsv = [headers, makeRow('안현수_화계병설유')]
    .map((row) => row.join(','))
    .join('\n');
  const [abbreviatedKindergartenRow] = await parseZoomAttendanceWorkbook(
    new File([abbreviatedKindergartenCsv], 'zoom.csv', { type: 'text/csv' }),
    [peopleWithOfficialSchools[1]],
    { startTime: '08:00', endTime: '09:50' },
  );
  assert.equal(abbreviatedKindergartenRow.result, '인정');
});
