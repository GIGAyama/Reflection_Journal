import test from 'node:test';
import assert from 'node:assert/strict';
import {
  appendJournal,
  computeClassId,
  createClassRecord,
  createPortfolio,
  decodeInvite,
  driveQueryValue,
  encodeInvite,
  inviteUrl,
  normalizeClassCode
} from '../docs/drive-core.js';

test('クラスIDは先生メールと正規化済みコードから決定的に作られる', async () => {
  const a = await computeClassId('Teacher@Example.ED.JP ', 'ab01-cd23');
  const b = await computeClassId('teacher@example.ed.jp', 'ABCD23');
  assert.equal(a, b);
  assert.match(a, /^[0-9a-f]{64}$/);
  assert.equal(normalizeClassCode('ab01-cd23'), 'ABCD23');
});

test('日本語を含む招待情報をURL安全に往復できる', async () => {
  const encoded = encodeInvite({
    classCode: 'ABC23456', className: '５年１組',
    teacherEmail: 'sensei@example.ed.jp', teacherName: '山田 先生'
  });
  assert.doesNotMatch(encoded, /[+/=]/);
  const decoded = await decodeInvite(encoded);
  assert.equal(decoded.className, '５年１組');
  assert.equal(decoded.teacherName, '山田 先生');
  assert.match(decoded.classId, /^[0-9a-f]{64}$/);
  const url = inviteUrl('https://example.github.io/app/?backend=drive', decoded);
  assert.match(url, /^https:\/\/example\.github\.io\/app\/#join=/);
});

test('ポートフォリオへの追記は元データを変更せず、必要な項目だけを保持する', () => {
  const klass = createClassRecord({
    classId: 'class-id', classCode: 'ABC23456', className: '6年2組',
    teacher: { email: 'teacher@example.ed.jp', name: '先生' }, now: '2026-01-01T00:00:00.000Z'
  });
  const original = createPortfolio({
    invite: { classId: klass.classId, classCode: klass.classCode, className: klass.className, teacherEmail: klass.teacher.email, teacherName: klass.teacher.name },
    student: { email: 'student@example.ed.jp', name: '児童A' }, now: '2026-01-02T00:00:00.000Z'
  });
  const updated = appendJournal(original, {
    id: 'journal-1', theme: '算数', content: '分数が分かった。', emotion: '😊'
  }, '2026-01-03T00:00:00.000Z');
  assert.equal(original.journals.length, 0);
  assert.equal(updated.journals.length, 1);
  assert.equal(updated.journals[0].content, '分数が分かった。');
  assert.equal(updated.updatedAt, '2026-01-03T00:00:00.000Z');
});

test('Drive検索値の引用符とバックスラッシュをエスケープする', () => {
  assert.equal(driveQueryValue("a'b\\c"), "a\\'b\\\\c");
});

test('壊れた招待情報を拒否する', async () => {
  await assert.rejects(() => decodeInvite('not-an-invite'), /招待情報/);
});

