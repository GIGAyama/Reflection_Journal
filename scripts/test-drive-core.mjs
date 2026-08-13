import test from 'node:test';
import assert from 'node:assert/strict';
import {
  appendJournal,
  analyzeClass,
  computeClassId,
  createChannel,
  createClassRecord,
  createPortfolio,
  currentTheme,
  decodeInvite,
  driveQueryValue,
  encodeInvite,
  exportCsv,
  inviteUrl,
  mergePortfoliosIntoMembers,
  normalizeClassCode,
  setFeedback,
  syncChannel
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
  const url = inviteUrl('https://example.github.io/app/', decoded);
  assert.match(url, /^https:\/\/example\.github\.io\/app\/#join=/);
});

test('教師所有チャンネルへテーマと返却を保存し、児童ポートフォリオと分離する', () => {
  const klass = createClassRecord({
    classId: 'class-id', classCode: 'ABC23456', className: '5年1組',
    teacher: { email: 'teacher@example.ed.jp', name: '先生' }, now: '2026-01-01T00:00:00.000Z'
  });
  klass.settings.todayTheme = { date: '2026-01-06', text: '今日できたこと' };
  const member = { email: 'student@example.ed.jp', name: '児童A', status: 'active' };
  const channel = createChannel({ classRecord: klass, member, now: '2026-01-02T00:00:00.000Z' });
  const returned = setFeedback(channel, 'journal-1', { comment: '具体的に書けましたね。', stamp: '👏' }, '2026-01-03T00:00:00.000Z');
  assert.equal(returned.feedback['journal-1'].returned, true);
  assert.equal(returned.feedback['journal-1'].stamp, '👏');
  klass.settings.weeklyThemes[2] = '火曜日のテーマ';
  const synced = syncChannel(returned, klass, member, '2026-01-04T00:00:00.000Z');
  assert.equal(synced.themes.weeklyThemes[2], '火曜日のテーマ');
  assert.equal(currentTheme(synced.themes, new Date(2026, 0, 6)), '今日できたこと');
});

test('共有ポートフォリオを名簿へ統合し、分析とCSVを生成する', () => {
  const klass = createClassRecord({ classId: 'class-id', classCode: 'ABC23456', className: '6年1組', teacher: { email: 'teacher@example.ed.jp', name: '先生' } });
  klass.settings.approvalRequired = false;
  const portfolio = createPortfolio({
    invite: { classId: 'class-id', classCode: 'ABC23456', className: '6年1組', teacherEmail: 'teacher@example.ed.jp', teacherName: '先生', approvalRequired: false },
    student: { email: 'student@example.ed.jp', name: '児童A' }, now: '2026-01-05T00:00:00.000Z'
  });
  const updated = appendJournal(portfolio, { id: 'j1', content: '学びを具体的に書いた。', emotion: '😊', createdAt: '2026-01-06T09:00:00.000Z' }, '2026-01-06T09:00:00.000Z');
  const items = [{ file: { id: 'portfolio-file' }, record: updated }];
  const merged = mergePortfoliosIntoMembers(klass, items);
  assert.equal(merged.members[0].status, 'active');
  assert.equal(merged.members[0].portfolioFileId, 'portfolio-file');
  const channel = setFeedback(createChannel({ classRecord: klass, member: merged.members[0] }), 'j1', { comment: 'いいですね', returned: true });
  const channels = new Map([['student@example.ed.jp', channel]]);
  const stats = analyzeClass(items, channels, new Date('2026-01-06T12:00:00.000Z'));
  assert.equal(stats.submittedToday, 1);
  assert.equal(stats.returned, 1);
  const csv = exportCsv(items, channels);
  assert.ok(csv.startsWith('\ufeff'));
  assert.match(csv, /いいですね/);
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
