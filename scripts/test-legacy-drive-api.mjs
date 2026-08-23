import test from 'node:test';
import assert from 'node:assert/strict';
import { DriveApiError, DriveClient } from '../legacy/drive-native/drive-api.js';

function jsonResponse(value, status = 200) {
  return new Response(JSON.stringify(value), { status, headers: { 'Content-Type': 'application/json' } });
}

test('ブラウザ標準fetchを正しいglobalThisレシーバーで呼び出す', async () => {
  const originalFetch = globalThis.fetch;
  const calls = [];
  globalThis.fetch = function (url, options) {
    assert.equal(this, globalThis);
    calls.push({ url: String(url), options });
    return Promise.resolve(jsonResponse({ files: [] }));
  };
  try {
    const client = new DriveClient('token-value');
    await client.listClasses();
    assert.equal(calls.length, 1);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('クラス一覧はdrive.file向けappProperties検索を使う', async () => {
  const calls = [];
  const client = new DriveClient('token-value', async (url, options) => {
    calls.push({ url: String(url), options });
    return jsonResponse({ files: [{ id: 'class-file' }] });
  });
  const files = await client.listClasses();
  assert.equal(files[0].id, 'class-file');
  const url = new URL(calls[0].url);
  assert.match(url.searchParams.get('q'), /rjType.*class/);
  assert.equal(calls[0].options.headers.get('Authorization'), 'Bearer token-value');
});

test('教師一覧はsharedWithMeとクラスIDの両方で絞る', async () => {
  let query = '';
  const client = new DriveClient('token', async (url) => {
    query = new URL(url).searchParams.get('q');
    return jsonResponse({ files: [] });
  });
  await client.listSharedPortfolios("class'id");
  assert.match(query, /sharedWithMe/);
  assert.match(query, /rjType.*portfolio/);
  assert.match(query, /class\\'id/);
});

test('児童のおへんじ一覧もsharedWithMeとチャンネル種別で絞る', async () => {
  let query = '';
  const client = new DriveClient('token', async (url) => {
    query = new URL(url).searchParams.get('q');
    return jsonResponse({ files: [] });
  });
  await client.listSharedChannels('class-id');
  assert.match(query, /sharedWithMe/);
  assert.match(query, /rjType.*channel/);
  assert.match(query, /rjClassId.*class-id/);
});

test('教師所有チャンネルは所有者とクラスIDで絞る', async () => {
  let query = '';
  const client = new DriveClient('token', async (url) => {
    query = new URL(url).searchParams.get('q');
    return jsonResponse({ files: [] });
  });
  await client.listOwnChannels('class-id');
  assert.match(query, /'me' in owners/);
  assert.match(query, /rjType.*channel/);
  assert.match(query, /rjClassId.*class-id/);
});

test('JSON作成はメタデータと本文をmultipartで一度に送る', async () => {
  let captured;
  const client = new DriveClient('token', async (url, options) => {
    captured = { url: String(url), options, body: await options.body.text() };
    return jsonResponse({ id: 'new-file', name: 'file.json' });
  });
  const result = await client.createJson('file.json', { hello: '世界' }, { rjType: 'portfolio' });
  assert.equal(result.id, 'new-file');
  assert.match(captured.url, /uploadType=multipart/);
  assert.match(captured.body, /"rjType":"portfolio"/);
  assert.match(captured.body, /"hello": "世界"/);
});

test('バックアップJSONは指定フォルダの子として作成する', async () => {
  let body = '';
  const client = new DriveClient('token', async (url, options) => {
    body = await options.body.text();
    return jsonResponse({ id: 'archive-file' });
  });
  await client.createJson('manifest.json', { ok: true }, { rjType: 'archive-item' }, ['folder-id']);
  assert.match(body, /"parents":\["folder-id"\]/);
});

test('更新前にDriveバージョンを確認し、古い画面からの上書きを拒否する', async () => {
  const calls = [];
  const client = new DriveClient('token', async (url, options = {}) => {
    calls.push({ url: String(url), options });
    return jsonResponse({ id: 'file-id', version: '8' });
  });
  await assert.rejects(
    () => client.updateJson('file-id', { updated: true }, { expectedVersion: '7' }),
    /別の端末でデータが更新/
  );
  assert.equal(calls.length, 1);
  assert.match(calls[0].url, /fields=/);
});

test('共有通知を送らず特定ユーザーへreader権限を作る', async () => {
  let captured;
  const client = new DriveClient('token', async (url, options) => {
    captured = { url: String(url), options };
    return jsonResponse({ id: 'permission' });
  });
  await client.shareWithUser('file-id', 'teacher@example.ed.jp');
  assert.equal(new URL(captured.url).searchParams.get('sendNotificationEmail'), 'false');
  assert.deepEqual(JSON.parse(captured.options.body), {
    type: 'user', role: 'reader', emailAddress: 'teacher@example.ed.jp'
  });
});

test('Workspaceの403を利用者向けメッセージへ変換する', async () => {
  const client = new DriveClient('token', async () => jsonResponse({ error: { message: 'domain policy' } }, 403));
  await assert.rejects(() => client.listClasses(), (error) => {
    assert.ok(error instanceof DriveApiError);
    assert.equal(error.status, 403);
    assert.match(error.message, /学校のGoogle Workspace設定/);
    assert.equal(error.detail, 'domain policy');
    return true;
  });
});
