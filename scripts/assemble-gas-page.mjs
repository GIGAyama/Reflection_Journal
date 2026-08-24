/**
 * GAS が組み立てる 1 枚の HTML を、手元で作り直す。
 *
 * doGet は index.html を「テンプレート」として評価し、その中の
 * `<?!= include('vendor') ?>` などが app / css / vendor / qr を差し込む。
 * つまり**ブラウザが受け取るのは 5 ファイルを貼り合わせた 1 枚**である。
 *
 * ファイル単位で構文を見ても、この 1 枚が壊れていないことは分からない。
 * 実際 2026-08-24 に、貼り合わせた側だけが壊れて画面が出なかった。
 * tests/gas-page.spec.mjs がこれを本物のブラウザに読ませる。
 */
import { readFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const read = (p) => readFileSync(join(ROOT, p), 'utf8');

/** @param {'owner'|'member'} mode  @param {object} boot doGet が渡すブート情報 */
export function assembleGasPage(mode = 'owner', boot = {}) {
  const bootObj = {
    mode,
    className: '3年2組',
    signedIn: true,
    setupDone: true,
    webAppUrl: 'https://script.google.com/macros/s/AAA/exec',
    ...boot
  };
  // Main.gs と同じ形で渡す（< を潰してからテンプレートへ入れる）
  const bootJson = JSON.stringify(bootObj).replace(/</g, '\\u003c');

  // ⚠️ 置き換えは必ず関数で渡すこと。文字列で渡すと、差し込む中身に含まれる
  //    `$&` や `$'` が「置換の特殊記号」として解釈され、別のものが入る。
  //    React の minify 済みコードには実際に `$'` が入っている。
  const put = (src, needle, value) => src.replace(needle, () => value);

  let out = read('index.html');
  out = put(out, '<?!= bootJson ?>', bootJson);
  // 先生の画面のときだけ QR を読む、という分岐
  out = put(out, "<? if (bootMode === 'owner') { ?><?!= include('qr'); ?><? } ?>",
    mode === 'owner' ? read('qr.html') : '');
  for (const name of ['vendor', 'css', 'app']) {
    out = put(out, `<?!= include('${name}'); ?>`, read(`${name}.html`));
  }

  // 差し込みが 1 つでも取りこぼされていたら、そこで気づく
  const left = out.match(/<\?[^\n]{0,40}/g);
  if (left) throw new Error('組み立てに失敗しました（スクリプトレットが残っています）: ' + left.join(' / '));
  return out;
}
