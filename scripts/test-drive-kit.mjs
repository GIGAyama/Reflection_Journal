/**
 * Driveネイティブ・キットの試験。
 *
 * 目的は2つある。
 *   1. キット単体が、別のアプリ設定でも正しく動くこと（横展開したときの安全性）
 *   2. キットの出力が、ふりかえりジャーナル本体の値と一致し続けること（＝作り直しても
 *      既存の児童・先生のファイルが読めなくなる事故を起こさないこと）
 * 2 を落とさない限り、キットと本体は分岐しない。
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  RecordCache,
  ScopeGrant,
  SessionPolicy,
  createDriveNativeApp,
  defineAppNamespace,
  driveQueryValue,
  isEmail,
  mergeSharedIntoRoster,
  normalizeEmail,
  ownerEmailOf,
  tokenExpiryFrom,
  validateSharedRecord
} from '../docs/kit/index.js';
import {
  createInviteKey, createdWithinInvite, decodeInvite, encodeSignedInvite,
  inviteKeyUsable, inviteTokenFromUrl, inviteUrl, matchesIssuedKey
} from '../docs/kit/invite.js';
import { KitDriveClient } from '../docs/kit/drive-client.js';
import {
  RJ, classAppProperties, computeClassId, decodeInvite as decodeAppInvite,
  encodeSignedInvite as signAppInvite, normalizeClassCode, studentKey
} from '../docs/drive-core.js';
import { DriveClient } from '../docs/drive-api.js';

// 横展開先の例。用語・接頭辞・版だけを差し替える。
const drill = defineAppNamespace({
  appId: 'kanji-drill',
  propertyPrefix: 'kd',
  schemaVersion: 3,
  terms: { tenant: 'Room', member: 'Learner' }
});

function jsonResponse(value, status = 200) {
  return new Response(JSON.stringify(value), { status, headers: { 'Content-Type': 'application/json' } });
}

test('名前空間ごとに、ID・appProperties・招待の版が分かれる', async () => {
  assert.deepEqual(drill.appProperties('room-1', 'portfolio'), {
    kdSchema: '3', kdType: 'portfolio', kdRoomId: 'room-1'
  });
  assert.equal(drill.properties.member, 'kdLearner');
  assert.equal(drill.tokenPrefix, 'kd3');
  const sameInput = await drill.tenantId('teacher@example.ed.jp', 'ABC23456');
  const otherApp = await RJ.tenantId('teacher@example.ed.jp', 'ABC23456');
  assert.match(sameInput, /^[0-9a-f]{64}$/);
  assert.notEqual(sameInput, otherApp, '同じ入力でもアプリが違えばテナントIDは別になる');
});

test('参加コードは英数字の紛らわしい文字を落として正規化する', () => {
  assert.equal(drill.normalizeCode('ab01-cd23'), 'ABCD23', '0・1・I・O は読み違えるので使わない');
  assert.equal(drill.normalizeCode('234567892345678'), '2345678923', '10文字で打ち切る');
  assert.equal(drill.randomCode(8).length, 8);
  assert.match(drill.randomCode(8), /^[23456789A-Z]{8}$/);
});

test('検索式は所有関係とテナントで絞り、引用符をエスケープする', () => {
  assert.equal(
    drill.query({ type: 'portfolio', owner: 'shared', tenantId: "room'1" }),
    "trashed = false and sharedWithMe and appProperties has { key='kdType' and value='portfolio' } and appProperties has { key='kdRoomId' and value='room\\'1' }"
  );
  assert.equal(
    drill.query({ type: 'channel', owner: 'me' }),
    "trashed = false and 'me' in owners and appProperties has { key='kdType' and value='channel' }"
  );
  assert.equal(driveQueryValue("a'b\\c"), "a\\'b\\\\c");
  assert.equal(isEmail('teacher@example.ed.jp'), true);
  assert.equal(normalizeEmail(' Teacher@Example.ED.JP '), 'teacher@example.ed.jp');
});

test('署名付き招待は改ざん・期限切れ・別アプリの版を拒否する', async () => {
  const security = await createInviteKey(drill, { now: new Date('2099-01-01T00:00:00.000Z') });
  assert.equal(inviteKeyUsable(security, new Date('2099-01-02T00:00:00.000Z')), true);
  assert.equal(inviteKeyUsable(security, new Date('2100-01-01T00:00:00.000Z')), false);

  const token = await encodeSignedInvite(drill, {
    code: 'ABC23456', name: '3年2組', hostEmail: 'sensei@example.ed.jp', hostName: '山田'
  }, security);
  assert.ok(token.startsWith('kd3.'));

  const invite = await decodeInvite(drill, token);
  assert.equal(invite.signed, true);
  assert.equal(invite.keyId, security.keyId);
  assert.equal(invite.tenantId, await drill.tenantId('sensei@example.ed.jp', 'ABC23456'));

  const parts = token.split('.');
  const tampered = `${parts[0]}.${parts[1].slice(0, -1)}${parts[1].endsWith('A') ? 'B' : 'A'}.${parts[2]}`;
  await assert.rejects(() => decodeInvite(drill, tampered), (error) => error.code === 'unreadable');
  // 別アプリの名前空間では、そもそも署名付きとして読まない（＝形式不一致で落ちる）
  await assert.rejects(() => decodeInvite(RJ, token), (error) => error.code === 'unreadable');

  const expiredKey = await createInviteKey(drill, { now: new Date('2000-01-01T00:00:00.000Z'), validityDays: 1 });
  const expired = await encodeSignedInvite(drill, {
    code: 'ABC23456', name: '3年2組', hostEmail: 'sensei@example.ed.jp', hostName: '山田'
  }, expiredKey);
  await assert.rejects(() => decodeInvite(drill, expired), (error) => error.code === 'expired');
  assert.equal((await decodeInvite(drill, expired, { allowExpired: true })).signed, true);
});

test('招待は発行世代と鍵IDが一致したときだけ取り込みを許す', async () => {
  const first = await createInviteKey(drill, { generation: 1, now: new Date('2099-01-01T00:00:00.000Z') });
  const rotated = await createInviteKey(drill, { generation: 2, now: new Date('2099-01-01T00:00:00.000Z') });
  const invite = await decodeInvite(drill, await encodeSignedInvite(drill, {
    code: 'ABC23456', name: '3年2組', hostEmail: 'sensei@example.ed.jp', hostName: '山田'
  }, first));

  assert.equal(matchesIssuedKey(invite, first, { tenantId: invite.tenantId }).ok, true);
  assert.equal(matchesIssuedKey(invite, rotated).reason, 'key_mismatch');
  assert.equal(matchesIssuedKey(invite, first, { tenantId: 'other-room' }).reason, 'tenant_mismatch');
  assert.equal(matchesIssuedKey({ ...invite, signed: false }, first).reason, 'unsigned');
  assert.equal(createdWithinInvite('2099-01-05T00:00:00.000Z', invite), true);
  assert.equal(createdWithinInvite('2100-01-05T00:00:00.000Z', invite), false);
});

test('招待URLはフラグメントに載り、そこから読み戻せる', () => {
  const url = inviteUrl('https://example.github.io/app/?utm=1', 'kd3.payload.signature');
  assert.equal(url, 'https://example.github.io/app/#join=kd3.payload.signature');
  assert.equal(inviteTokenFromUrl(url), 'kd3.payload.signature');
  assert.equal(inviteTokenFromUrl('https://example.github.io/app/'), '');
});

test('共有記録は種別・テナント・Drive所有者・宛先の4点で検証する', () => {
  const record = {
    kind: 'kanji-drill-channel',
    tenant: { id: 'room-1' },
    host: { email: 'teacher@example.ed.jp' },
    learner: { email: 'child@example.ed.jp' }
  };
  const file = { id: 'f1', owners: [{ emailAddress: 'teacher@example.ed.jp' }] };
  const options = {
    kind: 'kanji-drill-channel',
    tenantId: 'room-1',
    ownerMustBe: (value) => value?.host?.email,
    subjectEmail: 'child@example.ed.jp',
    subjectEmailOf: (value) => value?.learner?.email
  };
  assert.equal(validateSharedRecord(file, record, options).ok, true);
  assert.equal(ownerEmailOf(file), 'teacher@example.ed.jp');
  assert.equal(validateSharedRecord(file, { ...record, kind: 'other' }, options).reason, 'kind');
  assert.equal(validateSharedRecord(file, { ...record, tenant: { id: 'room-2' } }, options).reason, 'tenant');
  // 中身の自己申告だけが正しくても、Driveの所有者が違えば通さない
  assert.equal(validateSharedRecord({ ...file, owners: [{ emailAddress: 'attacker@example.ed.jp' }] }, record, options).reason, 'owner');
  assert.equal(validateSharedRecord(file, { ...record, learner: { email: 'other@example.ed.jp' } }, options).reason, 'subject');
});

test('版が同じ記録はキャッシュから返し、更新されたら取り直す', () => {
  const cache = new RecordCache();
  const file = { id: 'f1', modifiedTime: '2026-01-01T00:00:00.000Z', version: '3' };
  cache.remember(file, { value: 1 });
  assert.deepEqual(cache.read(file), { value: 1 });
  assert.equal(cache.read({ ...file, version: '4' }), null);
  assert.equal(cache.read({ ...file, modifiedTime: '2026-01-02T00:00:00.000Z' }), null);
  cache.forget('f1');
  assert.equal(cache.read(file), null);
});

test('名簿統合は既存の行を壊さず、変化が無ければ変更なしを返す', () => {
  const items = [{ file: { id: 'file-1' }, record: { learner: { email: 'Child@Example.ed.jp', name: '児童A' }, createdAt: '2026-01-05T00:00:00.000Z' } }];
  const options = {
    subjectOf: (record) => record.learner,
    fileIdField: 'recordFileId',
    defaults: { role: 'learner' },
    status: 'pending',
    now: '2026-02-01T00:00:00.000Z'
  };
  const first = mergeSharedIntoRoster([], items, options);
  assert.equal(first.changed, true);
  assert.deepEqual(first.members[0], {
    email: 'child@example.ed.jp', name: '児童A', role: 'learner',
    status: 'pending', recordFileId: 'file-1', joinedAt: '2026-01-05T00:00:00.000Z'
  });

  const again = mergeSharedIntoRoster(first.members, items, options);
  assert.equal(again.changed, false, '同じ内容の再同期でファイルを書き換えない');

  const invited = mergeSharedIntoRoster(
    [{ email: 'child@example.ed.jp', name: '名簿の名前', status: 'invited', recordFileId: '' }], items, options
  );
  assert.equal(invited.changed, true);
  assert.equal(invited.members[0].status, 'active', '招待済みの行は記録が届いた時点で参加済みになる');
  assert.equal(invited.members[0].name, '名簿の名前', '先生が入れた名前を児童の自己申告で上書きしない');
});

test('トークンは既定で保存せず、保存を選んだときだけ復元する', () => {
  const store = new Map();
  const storage = {
    getItem: (key) => (store.has(key) ? store.get(key) : null),
    setItem: (key, value) => store.set(key, value),
    removeItem: (key) => store.delete(key)
  };
  const base = { clientId: 'client-1', storageKey: 'kd_session', storage };
  const session = { accessToken: 'token', expiresAt: Date.now() + 3600_000, scopes: ['a'], user: { email: 'child@example.ed.jp', name: '児童' } };

  const memoryOnly = new SessionPolicy({ ...base, persist: false });
  assert.equal(memoryOnly.save(session), false);
  assert.equal(store.size, 0, '既定ではアクセストークンをどこにも書かない');

  const persisted = new SessionPolicy({ ...base, persist: true });
  assert.equal(persisted.save(session), true);
  assert.equal(persisted.restore().user.email, 'child@example.ed.jp');

  // 保存しない設定で開き直したら、前の保存分は消す
  assert.equal(new SessionPolicy({ ...base, persist: false }).restore(), null);
  assert.equal(store.size, 0);

  persisted.save(session);
  assert.equal(new SessionPolicy({ ...base, persist: true, clientId: 'client-2' }).restore(), null, 'クライアントIDが変わった保存分は使わない');
  persisted.save({ ...session, expiresAt: Date.now() + 10_000 });
  assert.equal(persisted.restore(), null, '期限間際のトークンは使わない');
  persisted.save(session);
  assert.equal(new SessionPolicy({ ...base, persist: true, allowedDomains: ['school.ed.jp'] }).restore(), null, '許可ドメイン外の保存分は使わない');
});

test('公開元とドメインの許可リストを判定する', () => {
  const policy = new SessionPolicy({
    allowedOrigins: ['https://gigayama.github.io'],
    allowedDomains: ['school.ed.jp']
  });
  assert.equal(policy.originAllowed('https://gigayama.github.io', 'gigayama.github.io'), true);
  assert.equal(policy.originAllowed('https://copy.example.com', 'copy.example.com'), false);
  assert.equal(policy.originAllowed('http://localhost:8080', 'localhost'), true, '手元の確認はできるようにする');
  assert.equal(policy.domainAllowed('teacher@school.ed.jp'), true);
  assert.equal(policy.domainAllowed('teacher@other.ed.jp'), false);
  assert.equal(new SessionPolicy({}).domainAllowed('anyone@example.com'), true, '未設定なら制限しない');
});

test('granular consentで外された権限を許可済みとして扱わない', () => {
  const grant = new ScopeGrant();
  grant.remember({ scope: 'openid email' }, ['https://www.googleapis.com/auth/drive.file']);
  assert.equal(grant.has('openid'), true);
  assert.equal(grant.has('https://www.googleapis.com/auth/drive.file'), false);
  grant.remember({ scope: '' }, ['https://www.googleapis.com/auth/drive.readonly'], {
    hasGrantedAllScopes: (response, scope) => scope.endsWith('drive.readonly')
  });
  assert.equal(grant.has('https://www.googleapis.com/auth/drive.readonly'), true);
  assert.deepEqual(grant.list().sort(), ['https://www.googleapis.com/auth/drive.readonly', 'email', 'openid'].sort());
  assert.ok(tokenExpiryFrom({ expires_in: 3600 }, 1000) === 1000 + 3600_000);
  assert.ok(tokenExpiryFrom({}, 0) === 3600_000, '期限が無い応答でも既定値で扱う');
});

test('キットのDriveクライアントは種別検索と競合検知をそのまま提供する', async () => {
  const calls = [];
  const client = new KitDriveClient({
    accessToken: 'token', namespace: drill,
    fetchImpl: async (url, options) => {
      calls.push({ url: String(url), options });
      return jsonResponse({ files: [], id: 'file-id', version: '9' });
    }
  });
  await client.listByType({ type: 'portfolio', owner: 'shared', tenantId: 'room-1' });
  assert.match(new URL(calls[0].url).searchParams.get('q'), /sharedWithMe.*kdType.*kdRoomId/);
  assert.equal(calls[0].options.headers.get('Authorization'), 'Bearer token');
  await assert.rejects(
    () => client.updateJson('file-id', { a: 1 }, { expectedVersion: '8' }),
    (error) => error.code === 'conflict' && error.status === 409
  );
});

test('Driveのエラー文言はアプリ側で差し替えられる', async () => {
  const client = new KitDriveClient({
    accessToken: 'token', namespace: drill,
    messages: { forbidden: 'この教室では共有が止められています。' },
    fetchImpl: async () => jsonResponse({ error: { message: 'policy' } }, 403)
  });
  await assert.rejects(() => client.listByType({ type: 'portfolio' }), (error) => {
    assert.equal(error.code, 'forbidden');
    assert.equal(error.message, 'この教室では共有が止められています。');
    assert.equal(error.detail, 'policy');
    return true;
  });
});

test('createDriveNativeApp は設定1か所からアプリ一式を組み立てる', async () => {
  const application = createDriveNativeApp({
    appId: 'kanji-drill',
    propertyPrefix: 'kd',
    schemaVersion: 3,
    terms: { tenant: 'Room', member: 'Learner' },
    clientId: 'client-1',
    allowedOrigins: ['https://example.github.io'],
    entryUrl: 'https://example.github.io/drill/'
  });
  assert.equal(application.scopes.base, 'openid email profile https://www.googleapis.com/auth/drive.file');
  assert.equal(application.scopes.sharedRead, 'https://www.googleapis.com/auth/drive.readonly');
  assert.equal(application.session.originAllowed('https://example.github.io', 'example.github.io'), true);

  const security = await application.invites.createKey({ now: new Date('2099-01-01T00:00:00.000Z') });
  const token = await application.invites.sign({
    code: 'ABC23456', name: '3年2組', hostEmail: 'sensei@example.ed.jp', hostName: '山田'
  }, security);
  const invite = await application.invites.decode(token);
  assert.equal(invite.tenantId, await application.namespace.tenantId('sensei@example.ed.jp', 'ABC23456'));
  assert.equal(application.invites.url('', token), `https://example.github.io/drill/#join=${token}`);
  assert.ok(application.client('token') instanceof KitDriveClient);
});

// ── ここから下は「本体と一致し続けること」の確認 ──

test('本体のID・appPropertiesはキットの計算と一致する', async () => {
  assert.equal(await computeClassId('Teacher@Example.ED.JP ', 'ab01-cd23'), await RJ.tenantId('teacher@example.ed.jp', 'ABCD23'));
  assert.equal(normalizeClassCode('ab01-cd23'), 'ABCD23');
  assert.equal(await studentKey('Child@Example.ed.jp'), await RJ.memberKey('child@example.ed.jp'));
  assert.deepEqual(classAppProperties('class-1', 'portfolio'), {
    rjSchema: '2', rjType: 'portfolio', rjClassId: 'class-1'
  });
  assert.equal(RJ.properties.member, 'rjStudent');
  assert.equal(RJ.tokenPrefix, 'rj2');
});

test('本体の招待トークンはキットの復号でもそのまま読める', async () => {
  const security = await createInviteKey(RJ, { now: new Date('2099-01-01T00:00:00.000Z') });
  const token = await signAppInvite({
    classCode: 'ABC23456', className: '５年１組',
    teacherEmail: 'sensei@example.ed.jp', teacherName: '山田 先生'
  }, security);
  const viaKit = await decodeInvite(RJ, token);
  const viaApp = await decodeAppInvite(token);
  assert.equal(viaKit.signed, true);
  assert.equal(viaKit.name, viaApp.className);
  assert.equal(viaKit.hostEmail, viaApp.teacherEmail);
  assert.equal(viaKit.tenantId, viaApp.classId);
  assert.equal(viaApp.className, '５年１組');
});

test('本体のDriveクライアントはキットの検索式を使っている', async () => {
  let query = '';
  const client = new DriveClient('token', async (url) => {
    query = new URL(url).searchParams.get('q');
    return jsonResponse({ files: [] });
  });
  await client.listSharedPortfolios('class-1');
  assert.equal(query, RJ.query({ type: 'portfolio', owner: 'shared', tenantId: 'class-1' }));
});
