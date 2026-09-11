import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { registerHooks } from 'node:module';
import test from 'node:test';
import ts from 'typescript';

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

const { validateRosterIntegrity } = await import('../src/lib/workflow.ts');

function roster(overrides = {}) {
  return {
    rowNumber: 2,
    sequence: '1',
    semester: '1학기',
    name: '홍길동',
    school: '테스트학교',
    phone: '01012345678',
    issueCount: 1,
    maxIssueCount: 2,
    course1: '테스트 연수',
    issuedAt1: '2026-09-11',
    link1: '',
    course2: '',
    issuedAt2: '',
    link2: '',
    completedTrainings: [],
    ...overrides,
  };
}

test('양식 없음 기록은 링크가 비어 있어도 정상으로 본다', () => {
  assert.deepEqual(validateRosterIntegrity([roster()]), []);
});

test('양식 없음 기록도 과정명과 발급날짜는 필수다', () => {
  const issues = validateRosterIntegrity([roster({ issuedAt1: '' })]);
  assert.equal(issues.length, 1);
  assert.match(issues[0].message, /과정명과 발급날짜/);
});
