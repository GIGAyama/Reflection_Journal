/** Tailwind の設定。使うクラスだけを先に作るため、原本を content に並べる。
 *  cdn.tailwindcss.com（ブラウザ内で CSS を生成する版）は使わない。 */
module.exports = {
  content: ['./src/**/*.jsx', './index.html'],
  theme: { extend: {} },
  plugins: []
};
