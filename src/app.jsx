const { useState, useEffect, useRef, useMemo } = React;

    // 端末の「動きを減らす」設定。紙吹雪など、止められるべき演出の判定に使う
    const prefersReducedMotion = () =>
      typeof window.matchMedia === 'function' && window.matchMedia('(prefers-reduced-motion: reduce)').matches;

    // モーダル共通: Esc で閉じる。開いている間は背後がスクロールしないようにする
    const useDismissable = (isOpen, onClose) => {
      useEffect(() => {
        if (!isOpen) return;
        const onKey = (e) => { if (e.key === 'Escape') { e.stopPropagation(); onClose(); } };
        window.addEventListener('keydown', onKey);
        return () => window.removeEventListener('keydown', onKey);
      }, [isOpen, onClose]);
    };

    // XSS対策: dangerouslySetInnerHTML に渡す前に必ずエスケープする
    const escapeHtml = (s) => String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');

    // highlights セルが壊れていてもUI全体を落とさない安全なパーサ
    const parseHighlights = (raw) => {
      if (!raw) return [];
      try {
        const arr = JSON.parse(raw);
        return Array.isArray(arr) ? arr : [];
      } catch (e) { return []; }
    };

    // 本文 + ハイライトを安全な HTML に変換する共通処理
    const buildHighlightedHtml = (rawContent, highlights, markBuilder) => {
      let html = escapeHtml(rawContent);
      highlights.forEach(h => {
        if (!h || !h.textToHighlight) return;
        const escapedTarget = escapeHtml(h.textToHighlight);
        if (html.includes(escapedTarget)) {
          html = html.replace(escapedTarget, markBuilder(escapedTarget, h));
        }
      });
      return html;
    };

    // 画像はサーバーへ送る前にブラウザ側で縮小・圧縮する
    // （シート内チャンク保存のため Data URL 上限 ≒ 700,000 文字に収める）
    const compressImage = (file, maxSize = 1024) => new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new Error('画像を読み込めませんでした'));
      reader.onload = (ev) => {
        const img = new Image();
        img.onerror = () => reject(new Error('画像を読み込めませんでした'));
        img.onload = () => {
          const scale = Math.min(1, maxSize / Math.max(img.width, img.height));
          const canvas = document.createElement('canvas');
          canvas.width = Math.max(1, Math.round(img.width * scale));
          canvas.height = Math.max(1, Math.round(img.height * scale));
          canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
          let quality = 0.85;
          let dataUrl = canvas.toDataURL('image/jpeg', quality);
          while (dataUrl.length > 700000 && quality > 0.3) {
            quality -= 0.15;
            dataUrl = canvas.toDataURL('image/jpeg', quality);
          }
          if (dataUrl.length > 700000) reject(new Error('画像が大きすぎます。小さい画像でためしてね'));
          else resolve(dataUrl);
        };
        img.src = ev.target.result;
      };
      reader.readAsDataURL(file);
    });

    // ふりがなの色は CSS 側（tools/extra.css）で決める。
    // ここで text-gray-500 と決め打ちしていたため、色のついたボタンや帯の上で
    // 比 1.08〜1.36 になり、いちばん読めなくて困る低学年がいちばん読めなかった。
    // <rp> は読み上げでふりがなが二重に読まれるのを防ぐ。
    const RubyText = ({ text, kana }) => (
      <ruby className="align-baseline">
        {text}
        <rp>(</rp>
        <rt className="text-[0.6em] font-medium select-none leading-none">{kana}</rt>
        <rp>)</rp>
      </ruby>
    );

    // className を渡すと既定値ごと置き換わるため、大きさを書き忘れた箇所で
    // SVG が親いっぱいに広がっていた（時計アイコンが巨大化していた）。
    // 大きさの指定が無いときだけ w-5 h-5 を足す。装飾なので読み上げからは外す。
    const Icon = ({ path, className = "" }) => {
      const hasSize = /(^|\s)(w-|h-|size-)/.test(className);
      return (
        <svg aria-hidden="true" focusable="false"
             className={`${hasSize ? '' : 'w-5 h-5 '}${className}`.trim()}
             fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d={path} /></svg>
      );
    };
    const Icons = {
      Book: "M4 19.5A2.5 2.5 0 0 1 6.5 17H20V4H6.5A2.5 2.5 0 0 0 4 6.5v13z M4 19.5v-13", Send: "M22 2L11 13 M22 2L15 22L11 13L2 9L22 2z",
      Image: "M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4 M17 8l-5-5-5 5 M12 3v12", Check: "M20 6L9 17l-5-5", User: "M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2 M12 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8z", Close: "M18 6L6 18 M6 6l12 12",
      Trash: "M3 6h18 M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2", Sparkles: "M12 3v3 M12 18v3 M3 12h3 M18 12h3 M5.6 5.6l2.1 2.1 M16.3 16.3l2.1 2.1 M5.6 18.4l2.1-2.1 M16.3 5.6l2.1 2.1", Info: "M12 22c5.523 0 10-4.477 10-10S17.523 2 12 2 2 6.477 2 12s4.477 10 10 10z M12 16v-4 M12 8h.01", Settings: "M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6z M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z", Search: "M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z", Pulse: "M4.318 16.318a4.5 4.5 0 000 6.364L12 20.364l7.682 2.318a4.5 4.5 0 000-6.364M12 7.636l-1.318-1.318a4.5 4.5 0 00-6.364 0M12 7.636l1.318-1.318a4.5 4.5 0 016.364 0M12 7.636v-6",
      Pen: "M15.232 5.232l3.536 3.536m-2.036-5.036a2.5 2.5 0 113.536 3.536L6.5 21.036H3v-3.572L16.732 3.732z", Save: "M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4",
      Plus: "M12 4v16m8-8H4", Download: "M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4",
      Link: "M13.828 10.172a4 4 0 00-5.656 0l-4 4a4 4 0 105.656 5.656l1.102-1.101m-.758-4.899a4 4 0 005.656 0l4-4a4 4 0 00-5.656-5.656l-1.1 1.1",
      Copy: "M8 16H6a2 2 0 01-2-2V6a2 2 0 012-2h8a2 2 0 012 2v2m-6 12h8a2 2 0 002-2v-8a2 2 0 00-2-2h-8a2 2 0 00-2 2v8a2 2 0 002 2z",
      Refresh: "M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"
    };

    // 児童用URLを第三者のQR生成サービスへ送らず、このブラウザ内だけでQR化する。
    // qr.html は先生ポータルにだけ読み込まれるため、児童端末の初回JSは増えない。
    const makeQrModel = (value) => {
      if (!value || typeof window.qrcode !== 'function') throw new Error('QR生成機能を読み込めませんでした');
      const qr = window.qrcode(0, 'M');
      qr.addData(String(value), 'Byte');
      qr.make();
      const count = qr.getModuleCount();
      const path = [];
      for (let row = 0; row < count; row += 1) {
        for (let col = 0; col < count; col += 1) {
          if (qr.isDark(row, col)) path.push(`M${col} ${row}h1v1H${col}z`);
        }
      }
      return { count, path: path.join('') };
    };

    const QrCode = ({ model, size = 180 }) => {
      if (!model) {
        return <div role="status" className="w-[180px] h-[180px] rounded-2xl border border-red-200 bg-red-50 text-red-700 text-xs font-bold p-4 flex items-center justify-center text-center">QRコードを生成できませんでした。URLをコピーして配布してください。</div>;
      }
      const quietZone = 4;
      const viewSize = model.count + quietZone * 2;
      return (
        <svg role="img" aria-label="児童用URLのQRコード" width={size} height={size}
          viewBox={`${-quietZone} ${-quietZone} ${viewSize} ${viewSize}`}
          className="rounded-2xl border border-gray-200 shadow-sm bg-white" shapeRendering="crispEdges">
          <rect x={-quietZone} y={-quietZone} width={viewSize} height={viewSize} fill="white" />
          <path d={model.path} fill="#111827" />
        </svg>
      );
    };

    const Toast = ({ message, type, onClose }) => {
      useEffect(() => { const timer = setTimeout(onClose, 3000); return () => clearTimeout(timer); }, []);
      // 保存できた・できなかったを画面で見ていない人にも伝える。
      // エラーは割り込んで読ませたいので role="alert"、それ以外は polite。
      const bg = type === 'error' ? 'bg-red-600' : 'bg-emerald-700';
      return <div role={type === 'error' ? 'alert' : 'status'} aria-live={type === 'error' ? 'assertive' : 'polite'}
        className={`no-print fixed bottom-6 left-1/2 transform -translate-x-1/2 ${bg} text-white px-6 py-3.5 rounded-full shadow-2xl flex items-center gap-2 z-50 animate-bounce font-bold tracking-wide`}><Icon path={type === 'error' ? Icons.Close : Icons.Check} className="w-5 h-5 shrink-0" /> <span>{message}</span></div>;
    };

    const ConfirmDialog = ({ isOpen, title, text, onConfirm, onCancel, confirmText="はい", cancelText="キャンセル", isDanger=false }) => {
      useDismissable(isOpen, onCancel || (() => {}));
      if (!isOpen) return null;
      return (
        <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-50 flex items-center justify-center p-4 transition-opacity">
          <div role="dialog" aria-modal="true" aria-label={title}
               className="bg-white rounded-[2rem] shadow-2xl shadow-black/20 p-8 max-w-sm w-full animate-fade-in text-center transform transition-all">
            <div className={`w-20 h-20 mx-auto rounded-full flex items-center justify-center mb-5 ${isDanger ? 'bg-red-50 text-red-500' : 'bg-blue-50 text-blue-500'}`}><Icon path={isDanger ? Icons.Trash : Icons.Info} className="w-10 h-10" /></div>
            <h3 className="text-2xl font-black text-gray-800 mb-3">{title}</h3>
            <p className="text-gray-500 mb-8 text-sm whitespace-pre-wrap leading-relaxed">{text}</p>
            <div className="flex gap-3">
              <button onClick={onCancel} className="flex-1 py-3.5 rounded-2xl font-bold bg-gray-50 text-gray-600 hover:bg-gray-100 transition-colors">{cancelText}</button>
              <button onClick={onConfirm} className={`flex-1 py-3.5 rounded-2xl font-bold text-white shadow-lg transition-all active:scale-95 ${isDanger ? 'bg-red-500 hover:bg-red-600 shadow-red-500/30' : 'bg-blue-600 hover:bg-blue-700 shadow-blue-500/30'}`}>{confirmText}</button>
            </div>
          </div>
        </div>
      );
    };

    // API 呼び出しの共通フック。owner / member を明示して呼び分ける。
    // 戻り値は常に { success, ... }。通信自体の失敗はトーストを出して success:false を返す。
    const useServer = () => {
      const [isLoading, setIsLoading] = useState(false);
      const [toast, setToast] = useState(null);
      const wrap = async (promise) => {
        setIsLoading(true);
        try { return await promise; }
        catch (e) { setToast({ message: (e && e.message) || '通信エラー', type: 'error' }); return { success: false, error: (e && e.message) || '通信エラー' }; }
        finally { setIsLoading(false); }
      };
      const owner = (fn, ...args) => wrap(window.callOwnerApi(fn, ...args));
      const member = (fn, ...args) => wrap(window.callMemberApi(fn, ...args));
      const memberWrite = (fn, ...args) => wrap(window.callMemberApiWithBackoff(fn, ...args));
      return { owner, member, memberWrite, isLoading, toast, setToast, showToast: (msg, type="success") => setToast({ message: msg, type }) };
    };

    const failMsg = (res, fallback) => (res && (res.message || res.error)) || fallback;

    // 画像の遅延読み込み（画像本体は getImage API から Data URL で取得する）
    const LazyImage = ({ imageId, fetcher, className, alt }) => {
      const [src, setSrc] = useState(null);
      const [failed, setFailed] = useState(false);
      useEffect(() => {
        let alive = true;
        setSrc(null); setFailed(false);
        if (!imageId) return;
        fetcher(imageId).then(res => {
          if (!alive) return;
          if (res && res.success && res.dataUrl) setSrc(res.dataUrl);
          else setFailed(true);
        }).catch(() => { if (alive) setFailed(true); });
        return () => { alive = false; };
      }, [imageId]);
      if (!imageId || failed) return null;
      if (!src) return <div className="mt-4 w-40 h-24 rounded-xl bg-gray-100 animate-pulse" />;
      return <img src={src} alt={alt || ''} className={className} />;
    };

    const FullScreenSpinner = ({ color = 'border-orange-500' }) => (
      <div className="h-screen flex items-center justify-center bg-gray-50"><div className={`w-16 h-16 border-4 ${color} border-t-transparent rounded-full animate-spin`}></div></div>
    );

    const CenterCard = ({ icon, children }) => (
      <div className="h-screen w-full flex items-center justify-center bg-gradient-to-br from-orange-50 via-amber-50 to-blue-50 p-4 overflow-y-auto">
        <div className="bg-white rounded-[2rem] shadow-premium border border-gray-100 p-8 md:p-10 max-w-xl w-full animate-fade-in my-8 text-center">
          <div className="w-20 h-20 mx-auto rounded-3xl bg-gradient-to-br from-orange-400 to-orange-500 text-white flex items-center justify-center mb-4 shadow-lg shadow-orange-500/30 text-3xl">{icon}</div>
          {children}
        </div>
      </div>
    );

    // ==========================================
    // 🎒 児童用画面 (Student App)
    // ==========================================
    const StudentApp = ({ data, server, refresh }) => {
      const { member, memberWrite, isLoading, toast, setToast, showToast } = server;
      const [content, setContent] = useState("");
      const [emotion, setEmotion] = useState("");
      const [imageFile, setImageFile] = useState(null);
      const [showHints, setShowHints] = useState(false);
      const [activePastJournal, setActivePastJournal] = useState(null);
      const [pastCommentInput, setPastCommentInput] = useState("");
      const [currentMonth, setCurrentMonth] = useState(new Date());
      const textareaRef = useRef(null);
      useDismissable(!!activePastJournal, () => setActivePastJournal(null));

      const [readJournals, setReadJournals] = useState([]);
      const [unreadJournals, setUnreadJournals] = useState([]);

      // localStorage のキーは email ではなく匿名 uid を使う（端末に email を残さない）
      const storeKey = (prefix) => `${prefix}_${data.user.uid}`;

      useEffect(() => {
        const savedDraft = localStorage.getItem(storeKey('draft'));
        if (savedDraft && !content) setContent(savedDraft);
        const savedRead = JSON.parse(localStorage.getItem(storeKey('read_journals')) || "[]");
        setReadJournals(savedRead);
        const unreads = data.journals.filter(j => j.status === '返却済み' && !savedRead.includes(j.journalId));
        setUnreadJournals(unreads);
      }, []);

      const handleContentChange = (e) => {
        const val = e.target.value;
        setContent(val);
        localStorage.setItem(storeKey('draft'), val);
      };

      const emotions = ["😊 うれしい", "😠 くやしい", "💡 なるほど", "🤔 もやもや"];
      const templates = { ywt: "【Y: やったこと】\n\n【W: わかったこと】\n\n【T: つぎにやること】\n", kpt: "【K: Keep よかったこと・つづけること】\n\n【P: Problem こまったこと・やめること】\n\n【T: Try つぎにためすこと】\n", '5w1h': "【いつ】\n\n【どこで】\n\n【だれが】\n\n【なにを】\n\n【なぜ】\n\n【どのように】\n", '3notice': "【わかったこと】\n\n【おどろいたこと】\n\n【もっとしりたいこと】\n" };
      const starters = ["今日いちばん心にのこったのは、", "今日、はじめてできたことは、", "今日うれしかったのは、", "明日がんばりたいことは、"];

      const hintDictionary = [
        { icon: "😊", title: <><RubyText text="気持" kana="きも"/>ち</>, items: [
          { text: "今日、いちばんうれしかったことは、", display: <><RubyText text="今日" kana="きょう"/>、いちばんうれしかったことは、</> },
          { text: "くやしかったこと・かなしかったことは、", display: <>くやしかったこと・かなしかったことは、</> },
          { text: "ドキドキ・ワクワクした瞬間は、", display: <>ドキドキ・ワクワクした<RubyText text="瞬間" kana="しゅんかん"/>は、</> }
        ]},
        { icon: "📚", title: <><RubyText text="学" kana="まな"/>び</>, items: [
          { text: "新しくわかったこと・できたことは、", display: <><RubyText text="新" kana="あたら"/>しくわかったこと・できたことは、</> },
          { text: "「なぜだろう？」と思ったことは、", display: <>「なぜだろう？」と<RubyText text="思" kana="おも"/>ったことは、</> },
          { text: "もっと調べたい・知りたいことは、", display: <>もっと<RubyText text="調" kana="しら"/>べたい・<RubyText text="知" kana="し"/>りたいことは、</> }
        ]},
        { icon: "🤝", title: "つながり", items: [
          { text: "友だちに「ありがとう」と言いたいことは、", display: <><RubyText text="友" kana="とも"/>だちに「ありがとう」と<RubyText text="言" kana="い"/>いたいことは、</> },
          { text: "だれかの「すごいな」と思ったところは、", display: <>だれかの「すごいな」と<RubyText text="思" kana="おも"/>ったところは、</> },
          { text: "みんなで協力してできたことは、", display: <>みんなで<RubyText text="協力" kana="きょうりょく"/>してできたことは、</> }
        ]},
        { icon: "🚀", title: <><RubyText text="次" kana="つぎ"/>へ</>, items: [
          { text: "明日がんばりたいことは、", display: <><RubyText text="明日" kana="あした"/>がんばりたいことは、</> },
          { text: "もう一度やるなら、どうする？", display: <>もう<RubyText text="一度" kana="いちど"/>やるなら、どうする？</> },
          { text: "今日の自分にひとこと言うなら、", display: <><RubyText text="今日" kana="きょう"/>の<RubyText text="自分" kana="じぶん"/>にひとこと<RubyText text="言" kana="い"/>うなら、</> }
        ]}
      ];

      const insertTextAtCursor = (textToInsert) => {
        const ta = textareaRef.current;
        if (ta) {
          const startPos = ta.selectionStart;
          const endPos = ta.selectionEnd;
          const newContent = content.substring(0, startPos) + textToInsert + content.substring(endPos);
          setContent(newContent);
          localStorage.setItem(storeKey('draft'), newContent);
          setTimeout(() => { ta.focus(); ta.setSelectionRange(startPos + textToInsert.length, startPos + textToInsert.length); }, 0);
        } else {
          setContent(prev => prev + textToInsert);
        }
        setShowHints(false);
      };

      const handleImage = (e) => {
        const file = e.target.files[0];
        if (!file) return;
        compressImage(file).then(dataUrl => {
          setImageFile({ name: file.name, dataUrl: dataUrl });
        }).catch(err => showToast(err.message, "error"));
      };

      const handleSubmit = async () => {
        if (!content.trim()) return showToast("本文を書いてね！", "error");
        const res = await memberWrite('mbSaveJournal', {
          content, emotion, imageDataUrl: imageFile ? imageFile.dataUrl : ''
        });
        if (!res || !res.success) return showToast(failMsg(res, "提出できませんでした。もう一度ためしてね。"), "error");
        localStorage.removeItem(storeKey('draft'));
        // 感覚過敏の児童に配慮し、「動きを減らす」設定のときは紙吹雪を出さない
        if (typeof confetti === 'function' && !prefersReducedMotion()) {
          confetti({ particleCount: 200, spread: 90, origin: { y: 0.6 }, colors: ['#f97316', '#fbbf24', '#3b82f6', '#10b981'] });
        }
        showToast("先生に提出できたよ！");
        setContent(""); setEmotion(""); setImageFile(null);
        setTimeout(refresh, 2000);
      };

      const openJournalModal = (journal) => {
        setActivePastJournal(journal);
        setPastCommentInput(journal.pastComment || "");
        if (journal.status === '返却済み' && !readJournals.includes(journal.journalId)) {
          const newRead = [...readJournals, journal.journalId];
          setReadJournals(newRead);
          localStorage.setItem(storeKey('read_journals'), JSON.stringify(newRead));
          setUnreadJournals(prev => prev.filter(j => j.journalId !== journal.journalId));
        }
      };

      const showRandomPastJournal = () => {
        const oneMonthAgo = new Date(); oneMonthAgo.setMonth(oneMonthAgo.getMonth() - 1);
        const pasts = data.journals.filter(j => new Date(j.timestamp) < oneMonthAgo);
        if (pasts.length === 0) return showToast("1ヶ月以上前の記録がまだないよ", "error");
        openJournalModal(pasts[Math.floor(Math.random() * pasts.length)]);
      };

      const handleSavePastComment = async () => {
        if(!pastCommentInput.trim()) return;
        const res = await memberWrite('mbAddPastComment', activePastJournal.journalId, pastCommentInput);
        if(res && res.success) { showToast("過去の自分にメッセージを送りました！"); setActivePastJournal(null); }
        else showToast(failMsg(res, "保存できませんでした"), "error");
      };

      const renderCalendarDays = () => {
        const y = currentMonth.getFullYear(), m = currentMonth.getMonth();
        const firstDay = new Date(y, m, 1).getDay(), lastDate = new Date(y, m + 1, 0).getDate();

        let days = [];
        for (let i = 0; i < firstDay; i++) days.push(<div key={`empty-${i}`} className="p-2"></div>);
        for (let i = 1; i <= lastDate; i++) {
          const ds = new Date(y, m, i).toLocaleDateString('ja-JP');
          const matchedJournals = data.journals.filter(j => new Date(j.timestamp).toLocaleDateString('ja-JP') === ds);
          const hasJournal = matchedJournals.length > 0;
          const isReturned = matchedJournals.some(j => j.status === '返却済み');

          days.push(
            <div key={i} className={`p-2 text-center rounded-xl text-sm font-bold transition-all relative ${hasJournal ? 'bg-orange-100 text-orange-700 shadow-sm border border-orange-200 cursor-pointer hover:bg-orange-200 hover:-translate-y-0.5' : 'text-gray-500'}`}
                 onClick={() => { if(hasJournal) { openJournalModal(matchedJournals[0]); } }}>
              {i}
              {isReturned && <div className="absolute -top-1 -right-1 w-4 h-4 bg-blue-500 rounded-full border-2 border-white shadow-sm flex items-center justify-center text-[8px] text-white animate-pulse">💬</div>}
            </div>
          );
        }
        return days;
      };

      return (
        <div className="flex flex-col h-full gap-6 relative animate-fade-in">
          {isLoading && <div className="absolute inset-0 bg-white/50 backdrop-blur-sm z-40 flex items-center justify-center"><div className="w-16 h-16 border-4 border-orange-500 border-t-transparent rounded-full animate-spin shadow-lg"></div></div>}

          {unreadJournals.length > 0 && (
            <div className="w-full bg-gradient-to-r from-blue-600 to-indigo-700 p-4 rounded-[2rem] shadow-lg shadow-blue-600/30 flex flex-col sm:flex-row gap-3 justify-between items-stretch sm:items-center animate-notification border border-blue-400">
              <div className="flex items-center gap-3 md:gap-4 text-white min-w-0">
                <div className="text-3xl animate-bounce shrink-0">💌</div>
                <div className="min-w-0">
                  <h3 className="font-black text-base md:text-lg"><RubyText text="先生" kana="せんせい"/>からのおへんじが<RubyText text="届" kana="とど"/>いてるよ！</h3>
                  <p className="text-sm font-bold opacity-90">{unreadJournals.length}件のあたらしいおへんじ</p>
                </div>
              </div>
              <button onClick={() => openJournalModal(unreadJournals[0])} className="bg-white text-blue-600 hover:bg-blue-50 font-black px-6 py-2.5 rounded-xl shadow-sm active:scale-95 transition-all shrink-0">
                <RubyText text="読" kana="よ"/>む
              </button>
            </div>
          )}

          {/* スマホ縦持ちでは「書く」領域を最上部に置き、カレンダー等は下へ回す */}
          <div className="flex flex-col lg:flex-row h-full gap-6">
            <div className="w-full lg:w-[280px] flex flex-col gap-5 order-2 lg:order-1">
               <div className="bg-white rounded-[2rem] shadow-premium border border-gray-100 p-5">
                  <div className="flex justify-between items-center mb-5">
                    <button onClick={()=>setCurrentMonth(new Date(currentMonth.setMonth(currentMonth.getMonth()-1)))} aria-label="前の月" className="tap-44 p-2 bg-gray-50 rounded-full text-gray-500 hover:text-orange-500 hover:bg-orange-50 transition-colors"><Icon path="M15 18l-6-6 6-6" className="w-4 h-4" /></button>
                    <h3 className="font-black text-gray-700">{currentMonth.getFullYear()}<RubyText text="年" kana="ねん"/> {currentMonth.getMonth()+1}<RubyText text="月" kana="がつ"/></h3>
                    <button onClick={()=>setCurrentMonth(new Date(currentMonth.setMonth(currentMonth.getMonth()+1)))} aria-label="次の月" className="tap-44 p-2 bg-gray-50 rounded-full text-gray-500 hover:text-orange-500 hover:bg-orange-50 transition-colors"><Icon path="M9 18l6-6-6-6" className="w-4 h-4"/></button>
                  </div>
                  <div className="grid grid-cols-7 gap-1 text-xs text-center text-gray-500 font-bold mb-3"><div className="text-red-600">日</div><div>月</div><div>火</div><div>水</div><div>木</div><div>金</div><div className="text-blue-600">土</div></div>
                  <div className="grid grid-cols-7 gap-1.5">{renderCalendarDays()}</div>
               </div>
               <button onClick={showRandomPastJournal} className="bg-gradient-to-br from-indigo-50 to-blue-50 hover:from-indigo-100 hover:to-blue-100 text-indigo-600 border border-indigo-100 font-black py-4 px-4 rounded-[2rem] shadow-sm flex items-center justify-center gap-2 transition-all active:scale-95 group">
                 <Icon path="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" className="group-hover:rotate-12 transition-transform" /> <span><RubyText text="過去" kana="かこ"/>の<RubyText text="自分" kana="じぶん"/>と<RubyText text="対話" kana="たいわ"/>する</span>
               </button>
            </div>

            <div className="flex-1 flex flex-col relative transition-all order-1 lg:order-2 min-h-[60vh] lg:min-h-0">

              <div className="flex-1 bg-white rounded-[2rem] shadow-premium border border-orange-100 flex flex-col overflow-hidden hover:shadow-lg transition-shadow relative">
                <div className="bg-gradient-to-r from-orange-50 to-amber-50 p-5 border-b border-orange-100 flex items-center gap-4">
                  <span className="bg-gradient-to-br from-orange-400 to-orange-600 text-white p-3 rounded-2xl shadow-sm shadow-orange-500/30"><Icon path={Icons.Book} className="w-6 h-6"/></span>
                  <div className="flex-1">
                    <p className="text-xs text-orange-700 font-black tracking-wider mb-1"><span><RubyText text="今日" kana="きょう" />のテーマ</span></p>
                    <h2 className="text-xl font-black text-gray-800">{data.todayTheme}</h2>
                  </div>
                  <div className="flex flex-col items-end gap-1">
                    {content.length > 0 && <span className="text-xs text-emerald-500 font-bold flex items-center gap-1"><Icon path={Icons.Save} className="w-3 h-3"/> <RubyText text="自動" kana="じどう"/><RubyText text="保存済" kana="ほぞんず"/>み</span>}
                  </div>
                </div>

                <textarea
                  ref={textareaRef}
                  value={content}
                  onChange={handleContentChange}
                  placeholder="ここをタップしてかきはじめよう..."
                  className="flex-1 w-full px-8 pt-[40px] pb-8 text-lg lined-paper resize-none outline-none text-gray-800 font-medium tracking-wide"
                />

                {showHints && (
                  <div className="absolute inset-x-0 bottom-0 bg-white/95 backdrop-blur-md border-t-2 border-yellow-200 p-6 shadow-[0_-10px_30px_rgba(0,0,0,0.05)] rounded-t-[2rem] z-20 animate-[fadeInScale_0.3s_ease-out]">
                    <div className="flex justify-between items-center mb-4">
                      <h3 className="font-black text-yellow-800 text-lg flex items-center gap-2"><Icon path={Icons.Sparkles} className="w-5 h-5"/> <RubyText text="書" kana="か"/>くことを<RubyText text="見" kana="み"/>つける<RubyText text="魔法" kana="まほう"/>のヒント</h3>
                      <button onClick={() => setShowHints(false)} aria-label="ヒントを閉じる" className="tap-44 bg-gray-100 hover:bg-gray-200 text-gray-600 p-2 rounded-full font-bold transition-colors"><Icon path={Icons.Close} className="w-5 h-5"/></button>
                    </div>
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                      {hintDictionary.map((cat, idx) => (
                        <div key={idx} className="bg-yellow-50/50 rounded-2xl p-3 border border-yellow-100">
                          <h4 className="font-bold text-yellow-700 text-sm mb-2">{cat.icon} {cat.title}</h4>
                          <div className="space-y-2">
                            {cat.items.map((item, i) => (
                              <button key={i} onClick={() => insertTextAtCursor(item.text)} className="w-full min-h-[44px] text-left text-xs bg-white hover:bg-yellow-100 border border-yellow-300 p-3 rounded-xl text-gray-700 font-medium transition-colors shadow-sm hover:shadow active:scale-95 leading-snug">
                                {item.display} <span className="text-yellow-700 float-right">＋</span>
                              </button>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                    <p className="text-center text-xs text-gray-500 mt-4">ボタンを<RubyText text="押" kana="お"/>すと、そのままノートに<RubyText text="入" kana="はい"/>るよ！</p>
                  </div>
                )}
              </div>
            </div>

            <div className="w-full lg:w-[280px] flex flex-col gap-5 order-3">
              <div className="bg-white rounded-[2rem] shadow-premium border border-gray-100 p-5">
                <p className="font-black text-gray-700 mb-3 text-sm flex items-center gap-2"><span className="p-1.5 bg-gray-100 rounded-lg"><Icon path={Icons.Pen} className="w-4 h-4 text-gray-600"/></span> <span>サポートツール</span></p>
                <div className="grid grid-cols-1 gap-2 mb-4">
                   <button onClick={() => setShowHints(true)} className="w-full py-3.5 bg-gradient-to-r from-yellow-50 to-amber-50 text-yellow-700 border border-yellow-200 rounded-xl hover:shadow-md transition-all font-black active:scale-95 flex items-center justify-center gap-2">
                     <Icon path={Icons.Sparkles} className="w-5 h-5"/> 💡 ヒントを<RubyText text="開" kana="ひら"/>く
                   </button>
                </div>
                <p className="text-xs font-bold text-gray-500 mb-2">▼ テンプレート（<RubyText text="形" kana="かたち"/>）を<RubyText text="使" kana="つか"/>う</p>
                <div className="grid grid-cols-2 gap-2 mb-4">
                  {Object.keys(templates).map(k => <button key={k} onClick={() => insertTextAtCursor(templates[k])} className="text-xs min-h-[44px] py-3 px-2 bg-gray-50 text-gray-700 border border-gray-200 rounded-xl hover:bg-white hover:border-gray-400 hover:shadow-sm transition-all font-bold active:scale-95">{k === 'ywt' ? <><RubyText text="YWT法" kana="ほう"/></> : k === 'kpt' ? <><RubyText text="KPT法" kana="ほう"/></> : k === '5w1h' ? '5W1H' : '3つのきづき'}</button>)}
                </div>
                <p className="text-xs font-bold text-gray-500 mb-2">▼ ランダムな<RubyText text="書" kana="か"/>き<RubyText text="出" kana="だ"/>し</p>
                <div className="flex gap-2">
                   <button onClick={() => insertTextAtCursor(starters[Math.floor(Math.random() * starters.length)])} className="flex-1 text-xs min-h-[44px] py-3 bg-blue-50 text-blue-700 rounded-xl hover:bg-blue-100 transition-colors font-bold active:scale-95">🎯 <RubyText text="書" kana="か"/>き<RubyText text="出" kana="だ"/>しを追加(ついか)</button>
                </div>
              </div>

              <div className="bg-white rounded-[2rem] shadow-premium border border-gray-100 p-5">
                <p className="font-black text-gray-700 mb-3 text-sm flex items-center gap-2"><span className="p-1.5 bg-gray-100 rounded-lg text-base leading-none">😊</span> <span><RubyText text="気持" kana="きも" />ちをえらぼう</span></p>
                <div className="grid grid-cols-2 gap-2">
                  {emotions.map(e => <button key={e} onClick={() => setEmotion(e)} className={`py-3 px-2 text-sm rounded-xl border-2 transition-all active:scale-95 ${emotion === e ? 'border-orange-500 bg-orange-50 font-black text-orange-600 shadow-sm shadow-orange-500/20' : 'border-gray-100 hover:border-orange-200 text-gray-600 hover:bg-orange-50/50'}`}>{e}</button>)}
                </div>
              </div>

              <div className="bg-white rounded-[2rem] shadow-premium border border-gray-100 p-5">
                <label className="flex items-center justify-center gap-2 w-full py-4 border-2 border-dashed border-gray-300 rounded-xl cursor-pointer hover:bg-gray-50 hover:border-gray-400 transition-colors text-gray-500 font-bold text-sm">
                  <Icon path={Icons.Image} /> <span>{imageFile ? imageFile.name : <><RubyText text="画像" kana="がぞう" />をのせる</>}</span>
                  <input type="file" accept="image/*" className="hidden" onChange={handleImage} />
                </label>
                {imageFile && <img src={imageFile.dataUrl} alt="のせる画像のプレビュー" className="mt-3 rounded-xl max-h-32 mx-auto object-cover shadow-sm"/>}
              </div>

              <button onClick={handleSubmit} className="mt-auto bg-gradient-to-b from-orange-700 to-orange-800 hover:from-orange-800 hover:to-orange-900 text-white font-black py-4 rounded-[2rem] shadow-lg shadow-orange-500/40 flex justify-center items-center gap-2 text-lg active:scale-95 transition-all">
                <Icon path={Icons.Send} /> <span><RubyText text="提出" kana="ていしゅつ" />する</span>
              </button>
            </div>
          </div>

          {activePastJournal && (
             <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
               <div role="dialog" aria-modal="true" aria-label="むかしの記録"
                    className="bg-white rounded-[2rem] shadow-2xl w-full max-w-3xl max-h-[90vh] flex flex-col overflow-hidden animate-fade-in border border-gray-100">

                 <div className="bg-gradient-to-r from-blue-50 to-indigo-50 p-5 flex justify-between items-center border-b border-indigo-100 shrink-0">
                    <h3 className="font-black text-indigo-800 text-lg flex items-center gap-2"><Icon path="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" className="text-indigo-500"/> {activePastJournal.date} <RubyText text="の記録" kana="のきろく"/></h3>
                    <button onClick={()=>setActivePastJournal(null)} aria-label="閉じる" className="tap-44 p-2 bg-white rounded-full text-gray-500 hover:text-indigo-600 shadow-sm transition-colors"><Icon path={Icons.Close}/></button>
                 </div>

                 <div className="p-8 overflow-y-auto flex-1 bg-white">
                    {activePastJournal.status === '返却済み' && (
                      <div className="bg-gradient-to-r from-blue-50 to-cyan-50 border-2 border-blue-200 p-6 rounded-3xl mb-8 shadow-sm relative">
                        <div className="absolute -top-3 -right-3 text-4xl rotate-12">{activePastJournal.teacherStamp}</div>
                        <h4 className="font-black text-blue-700 mb-3 flex items-center gap-2 text-lg"><span className="text-2xl">👩‍🏫</span> <RubyText text="先生" kana="せんせい"/>からのおへんじ</h4>
                        {activePastJournal.teacherComment && <p className="text-gray-800 whitespace-pre-wrap font-bold leading-relaxed">{activePastJournal.teacherComment}</p>}
                      </div>
                    )}

                    <div className="text-gray-700 whitespace-pre-wrap lined-paper text-lg leading-[40px] font-medium mb-6">
                      {(() => {
                        const hls = parseHighlights(activePastJournal.highlights);
                        if (hls.length === 0) return activePastJournal.content;
                        const html = buildHighlightedHtml(activePastJournal.content, hls,
                          (escapedText) => `<mark class="bg-yellow-200/60 border-b-2 border-yellow-400 px-1 py-0.5 rounded">${escapedText}</mark>`);
                        return <div dangerouslySetInnerHTML={{__html: html}} />;
                      })()}
                    </div>
                    <LazyImage imageId={activePastJournal.imageId} fetcher={(id) => window.callMemberApi('mbGetImage', id)} alt="この日にのせた画像" className="mt-4 rounded-xl border max-w-xs shadow-sm"/>
                 </div>

                 <div className="bg-gray-50 p-6 border-t border-gray-100 shrink-0">
                    <p className="text-sm font-black text-indigo-600 mb-3 flex items-center gap-2"><Icon path="M8 12h.01M12 12h.01M16 12h.01M21 12c0 4.418-4.03 8-9 8a9.863 9.863 0 01-4.255-.949L3 20l1.395-3.72C3.512 15.042 3 13.574 3 12c0-4.418 4.03-8 9-8s9 3.582 9 8z" className="w-4 h-4"/> <span><RubyText text="今" kana="いま"/>の<RubyText text="自分" kana="じぶん"/>から、<RubyText text="過去" kana="かこ"/>の<RubyText text="自分" kana="じぶん"/>へメッセージ</span></p>
                    <textarea value={pastCommentInput} onChange={e=>setPastCommentInput(e.target.value)} className="w-full p-4 bg-white border border-gray-200 rounded-2xl resize-none outline-none focus:border-indigo-400 focus:ring-4 focus:ring-indigo-50 transition-all font-medium" rows="2" placeholder="1ヶ月前の自分、がんばってたね！"></textarea>
                    <div className="text-right mt-3"><button onClick={handleSavePastComment} className="bg-indigo-600 hover:bg-indigo-700 text-white px-8 py-3 rounded-2xl font-bold shadow-lg shadow-indigo-500/30 active:scale-95 transition-all"><RubyText text="保存" kana="ほぞん"/>する</button></div>
                 </div>
               </div>
             </div>
          )}
          {toast && <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />}
        </div>
      );
    };

    // ==========================================
    // 🧠 心のバイタルサイン解析 (先生用)
    // ==========================================
    const analyzeVitals = (journals, roster) => {
      const studentMap = {};
      roster.forEach(s => { studentMap[s.email] = { name: s.name, email: s.email, journals: [], alerts: [] }; });
      journals.forEach(j => { if(studentMap[j.email]) studentMap[j.email].journals.push(j); });

      const today = new Date(); today.setHours(0,0,0,0);
      const past14Days = Array.from({length: 14}, (_, i) => {
         const d = new Date(today); d.setDate(d.getDate() - (13 - i)); return d.toLocaleDateString('ja-JP');
      });

      Object.values(studentMap).forEach(s => {
         s.journals.sort((a,b) => new Date(a.timestamp) - new Date(b.timestamp));
         s.heatMap = past14Days.map(dateStr => {
            const j = s.journals.find(jj => new Date(jj.timestamp).toLocaleDateString('ja-JP') === dateStr);
            if(!j) return { date: dateStr, status: 'none', color: 'bg-gray-100', content: '未提出' };
            const emotion = String(j.emotion ?? '');
            let color = 'bg-gray-200';
            if(emotion.includes('うれしい')) color = 'bg-rose-400 shadow-rose-400/40 shadow-sm';
            else if(emotion.includes('なるほど')) color = 'bg-amber-400 shadow-amber-400/40 shadow-sm';
            else if(emotion.includes('もやもや')) color = 'bg-indigo-400 shadow-indigo-400/40 shadow-sm';
            else if(emotion.includes('くやしい')) color = 'bg-violet-600 shadow-violet-600/40 shadow-sm';
            return { date: dateStr, status: 'submitted', color, content: j.content, emotion: emotion };
         });

         const recent = s.journals.slice(-3);
         if (recent.length >= 2) {
            let nV = 0; recent.forEach(j => { const em = String(j.emotion ?? ''); if(em.includes('もやもや') || em.includes('くやしい')) nV++; });
            if (nV >= 2) s.alerts.push("ネガティブな感情が続いています");
         }
         if (s.journals.length >= 3) {
            const lenOf = (j) => String(j.content ?? '').length;
            const avgLen = s.journals.slice(0, -1).reduce((acc, j) => acc + lenOf(j), 0) / (s.journals.length - 1);
            const lastLen = lenOf(s.journals[s.journals.length - 1]);
            if (avgLen > 20 && lastLen < avgLen * 0.3) s.alerts.push("記述量が急激に減少しています");
         }
      });
      return { students: Object.values(studentMap).sort((a,b) => b.alerts.length - a.alerts.length), dates: past14Days };
    };

    // 学期末の記録用: ブラウザの印刷ダイアログから PDF 保存できる帳票を開く
    const openPrintView = (journals, tenantName) => {
      const grouped = {};
      journals.slice().sort((a,b) => (a.studentName||'').localeCompare(b.studentName||'') || new Date(a.timestamp) - new Date(b.timestamp))
        .forEach(j => { const k = j.studentName || '不明'; (grouped[k] = grouped[k] || []).push(j); });
      const pages = Object.keys(grouped).map(name => `
        <section class="student">
          <h2>${escapeHtml(name)}</h2>
          ${grouped[name].map(j => `
            <article>
              <h3>📅 ${escapeHtml(j.date || '')}</h3>
              ${j.theme ? `<p class="theme">テーマ: ${escapeHtml(j.theme)}</p>` : ''}
              <p class="content">${escapeHtml(j.content || '').replace(/\n/g, '<br>')}</p>
              ${j.teacherComment ? `<p class="comment">💬 先生より<br>${escapeHtml(j.teacherComment).replace(/\n/g, '<br>')}</p>` : ''}
            </article>`).join('')}
        </section>`).join('');
      const w = window.open('', '_blank');
      if (!w) return false;
      w.document.write(`<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8"><title>ふりかえりジャーナル_${escapeHtml(tenantName || '')}</title>
        <style>
          body { font-family: 'Hiragino Sans', 'Yu Gothic', sans-serif; color: #1f2937; margin: 24px; }
          section.student { page-break-after: always; }
          h2 { text-align: center; border-bottom: 2px solid #f97316; padding-bottom: 8px; }
          article { border-bottom: 1px dashed #cbd5e1; padding: 12px 0; }
          h3 { margin: 0 0 4px; font-size: 14px; }
          .theme { color: #6b7280; font-size: 12px; margin: 0 0 8px; }
          .content { line-height: 1.9; white-space: normal; }
          .comment { background: #eff6ff; border-radius: 8px; padding: 10px; }
        </style></head><body>${pages}<script>window.onload = function(){ window.print(); };<\/script></body></html>`);
      w.document.close();
      return true;
    };

    // ==========================================
    // 👩‍🏫 教員用画面 (Teacher App)
    // ==========================================
    const TeacherApp = ({ data, tenants, server, refresh, onSwitchTenant, onCreateNew }) => {
      const { owner, isLoading, toast, setToast, showToast } = server;
      const code = data.tenant.tenantCode;
      const call = (fn, ...args) => owner(fn, code, ...args);

      const [journals, setJournals] = useState(data.journals);
      const [activeTab, setActiveTab] = useState('dashboard');
      const [activeJournal, setActiveJournal] = useState(null);
      const [feedback, setFeedback] = useState("");

      const [filterStudent, setFilterStudent] = useState('all');
      const [filterStatus, setFilterStatus] = useState('all');
      const [searchQuery, setSearchQuery] = useState('');

      const [weeklyThemes, setWeeklyThemes] = useState(data.weeklyThemes || { mon: '', tue: '', wed: '', thu: '', fri: '' });
      const [todayThemeInput, setTodayThemeInput] = useState(data.todayTheme);
      const [confirmConfig, setConfirmConfig] = useState({ isOpen: false });

      // 名簿管理用ステート（状態列 = active/pending を含む）
      const [rosterData, setRosterData] = useState(data.rosterAll || []);
      const [apiKeyInput, setApiKeyInput] = useState("");
      const [copiedItem, setCopiedItem] = useState('');
      useDismissable(!!activeJournal, () => setActiveJournal(null));

      useEffect(() => {
        setJournals(data.journals);
        setRosterData(data.rosterAll || []);
        setWeeklyThemes(data.weeklyThemes || { mon: '', tue: '', wed: '', thu: '', fri: '' });
        setTodayThemeInput(data.todayTheme);
      }, [data]);

      const vitalData = useMemo(() => analyzeVitals(journals, data.classRoster), [journals, data.classRoster]);
      const pendingMembers = (data.pendingMembers || []);
      const inviteQr = useMemo(() => {
        try { return makeQrModel(data.tenant.memberUrl); }
        catch (e) { return null; }
      }, [data.tenant.memberUrl]);

      const filteredJournals = useMemo(() => {
        return journals.filter(j => {
          const mStudent = filterStudent === 'all' || j.email === filterStudent;
          const mStatus = filterStatus === 'all' || j.status === filterStatus;
          // 本文が数値のみ等の場合シートから数値型で返るため、必ず文字列化してから検索する
          const mSearch = !searchQuery || String(j.content ?? '').includes(searchQuery) || String(j.studentName ?? '').includes(searchQuery) || String(j.theme ?? '').includes(searchQuery);
          return mStudent && mStatus && mSearch;
        }).sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp));
      }, [journals, filterStudent, filterStatus, searchQuery]);

      const handleTextSelect = () => {
         const selection = window.getSelection();
         if (!selection.isCollapsed && selection.toString().trim() !== "" && activeJournal) {
            const text = selection.toString();
            const newHl = { id: 'hl-'+Date.now(), textToHighlight: text, suggestedComment: '', suggestedStamp: '' };
            const currentHighlights = parseHighlights(activeJournal.highlights);
            currentHighlights.push(newHl);
            setActiveJournal({...activeJournal, highlights: JSON.stringify(currentHighlights)});
            selection.removeAllRanges();
         }
      };

      const updateHighlight = (id, field, value) => {
         const updated = parseHighlights(activeJournal.highlights).map(h => h.id === id ? {...h, [field]: value} : h);
         setActiveJournal({...activeJournal, highlights: JSON.stringify(updated)});
      };

      const removeHighlight = (id) => {
         const updated = parseHighlights(activeJournal.highlights).filter(h => h.id !== id);
         setActiveJournal({...activeJournal, highlights: JSON.stringify(updated)});
      };

      const handleSetTheme = async () => {
        const res = await call('opSetTodayTheme', todayThemeInput);
        if(res && res.success) showToast("今日のテーマを設定しました");
        else showToast(failMsg(res, "テーマを設定できませんでした"), "error");
      };

      const handleSaveWeekly = async () => {
        const res = await call('opSaveWeeklyThemes', weeklyThemes);
        if(res && res.success) showToast("週間スケジュールを保存しました");
        else showToast(failMsg(res, "保存できませんでした"), "error");
      };

      const executeAiAction = async (type) => {
        setConfirmConfig({ isOpen: false });
        const res = await call(type === 'simple' ? 'opAiSimple' : 'opAiFull');
        if (res && res.success) { showToast(res.message); setTimeout(refresh, 1500); }
        else showToast(failMsg(res, "AI処理に失敗しました"), "error");
      };

      const handleSaveFeedback = async () => {
        const res = await call('opSaveFeedback', { journalId: activeJournal.journalId, comment: feedback, highlights: activeJournal.highlights });
        if (res && res.success) {
          showToast("フィードバックを返却しました");
          setJournals(journals.map(j => j.journalId === activeJournal.journalId ? {...j, status: '返却済み', teacherComment: feedback, highlights: activeJournal.highlights} : j));
          setActiveJournal(null);
        } else showToast(failMsg(res, "保存できませんでした"), "error");
      };

      const handleQuickAction = async (jId, action, value) => {
        if (action === 'stamp') {
          const res = await call('opQuickReturn', jId, value);
          if(res && res.success){ setJournals(journals.map(j => j.journalId === jId ? {...j, status: '返却済み', teacherStamp: value} : j)); showToast(`${value} を押しました！`); }
          else showToast(failMsg(res, "返却できませんでした"), "error");
        } else if (action === 'revert') {
          const res = await call('opRevertStatus', jId);
          if(res && res.success){ setJournals(journals.map(j => j.journalId === jId ? {...j, status: '未返却'} : j)); showToast("返却を取り消しました"); }
          else showToast(failMsg(res, "取り消せませんでした"), "error");
        }
      };

      const executeDataReset = async () => {
        setConfirmConfig({ isOpen: false });
        const res = await call('opResetData');
        if(res && res.success) { showToast("全データを削除しました"); setTimeout(refresh, 1500); }
        else showToast(failMsg(res, "削除できませんでした"), "error");
      };

      const executeBatchReturn = async () => {
        setConfirmConfig({ isOpen: false });
        const res = await call('opBatchReturnAll');
        if (res && res.success) { showToast(res.message); setTimeout(refresh, 1500); }
        else showToast(failMsg(res, "一括返却できませんでした"), "error");
      };

      const handleRosterChange = (index, field, value) => {
        const newRoster = [...rosterData];
        newRoster[index] = {...newRoster[index], [field]: value};
        setRosterData(newRoster);
      };

      const handleAddRoster = () => {
        setRosterData([...rosterData, { role: '児童', name: '', email: '', status: 'active' }]);
      };

      const handleRemoveRoster = (index) => {
        const newRoster = [...rosterData];
        newRoster.splice(index, 1);
        setRosterData(newRoster);
      };

      const handleSaveRoster = async () => {
        const validRows = rosterData.filter(r => String(r.name||'').trim() && String(r.email||'').trim());
        if (validRows.length === 0) return showToast('有効なデータがありません', 'error');
        const res = await call('opSaveRoster', validRows);
        if (res && res.success) {
          showToast('名簿を更新しました');
          setTimeout(refresh, 1200);
        } else showToast(failMsg(res, '名簿を保存できませんでした'), 'error');
      };

      const handleApprove = async (email, ok) => {
        const res = await call(ok ? 'opApproveMember' : 'opRejectMember', email);
        if (res && res.success) { showToast(ok ? '参加を承認しました' : '申請を却下しました'); refresh(); }
        else showToast(failMsg(res, '処理できませんでした'), 'error');
      };

      const handleExport = async (type) => {
         if (type === 'pdf') {
            if (journals.length === 0) return showToast('データがありません', 'error');
            if (!openPrintView(journals, data.tenant.tenantName)) showToast('ポップアップがブロックされました。許可してください', 'error');
            return;
         }
         const res = await call('opExportCsv', { startDate: null, endDate: null, email: 'all' });
         if(res && res.success) {
            const blob = new Blob([res.csv], { type: 'text/csv;charset=utf-8' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = res.fileName;
            document.body.appendChild(a); a.click(); a.remove();
            setTimeout(() => URL.revokeObjectURL(a.href), 5000);
            showToast('CSVをダウンロードしました');
         } else {
            showToast(failMsg(res, '出力できませんでした'), "error");
         }
      };

      const handleSaveApiKey = async () => {
        if (!apiKeyInput.trim()) return showToast('APIキーを入力してください', 'error');
        const res = await call('opSaveSettings', { apiKey: apiKeyInput.trim() });
        if (res && res.success) { showToast('APIキーを保存しました'); setApiKeyInput(''); refresh(); }
        else showToast(failMsg(res, '保存できませんでした'), 'error');
      };

      const copyInviteValue = async (value, item, label) => {
        try {
          if (!navigator.clipboard || !window.isSecureContext) throw new Error('clipboard unavailable');
          await navigator.clipboard.writeText(value);
          setCopiedItem(item);
          showToast(`${label}をコピーしました`);
          setTimeout(() => setCopiedItem(current => current === item ? '' : current), 2000);
        } catch (e) {
          window.prompt(`この${label}をコピーしてください`, value);
        }
      };

      const handleCopyUrl = () => copyInviteValue(data.tenant.memberUrl, 'url', 'URL');
      const handleCopyCode = () => copyInviteValue(data.tenant.tenantCode, 'code', 'クラスコード');
      const handleCopyInvitation = () => copyInviteValue(
        `「${data.tenant.tenantName}」のふりかえりジャーナルに参加してください。\nクラスコード: ${data.tenant.tenantCode}\n児童用URL: ${data.tenant.memberUrl}`,
        'invitation',
        '招待文'
      );

      const handleDownloadQr = () => {
        if (!inviteQr) return showToast('QRコードを生成できませんでした', 'error');
        const quietZone = 4;
        const viewSize = inviteQr.count + quietZone * 2;
        const svg = `<?xml version="1.0" encoding="UTF-8"?>\n<svg xmlns="http://www.w3.org/2000/svg" viewBox="${-quietZone} ${-quietZone} ${viewSize} ${viewSize}" shape-rendering="crispEdges"><rect x="${-quietZone}" y="${-quietZone}" width="${viewSize}" height="${viewSize}" fill="white"/><path d="${inviteQr.path}" fill="#111827"/></svg>`;
        const url = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml;charset=utf-8' }));
        const a = document.createElement('a');
        const safeName = String(data.tenant.tenantName || 'クラス').replace(/[\\/:*?"<>|]/g, '_');
        a.href = url;
        a.download = `ふりかえりジャーナル_${safeName}_${data.tenant.tenantCode}_QR.svg`;
        document.body.appendChild(a); a.click(); a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 5000);
        showToast('QR画像を保存しました');
      };

      const handleRegenerateCode = async () => {
        setConfirmConfig({ isOpen: false });
        const res = await call('opRegenerateCode');
        if (res && res.success) { showToast('クラスコードを再発行しました。旧URLは使えなくなります'); refresh(); }
        else showToast(failMsg(res, '再発行できませんでした'), 'error');
      };

      const handleJoinPolicy = async (joinOpen, requireApproval) => {
        const res = await call('opUpdateJoinPolicy', joinOpen, requireApproval);
        if (res && res.success) { showToast('設定を変更しました'); refresh(); }
        else showToast(failMsg(res, '変更できませんでした'), 'error');
      };

      // 同じ児童が複数回提出しても100%を超えないよう、ユニークな児童数でカウントする
      const pct = useMemo(() => {
        if (!data.classRoster.length) return 0;
        const todayStr = new Date().toLocaleDateString('ja-JP');
        const submitted = new Set(journals.filter(j => new Date(j.timestamp).toLocaleDateString('ja-JP') === todayStr).map(j => j.email));
        return Math.min(100, Math.round(submitted.size / data.classRoster.length * 100));
      }, [journals, data.classRoster]);

      return (
        <div className="h-full flex flex-col gap-6 relative">
          {isLoading && <div className="absolute inset-0 bg-white/50 backdrop-blur-sm z-40 flex items-center justify-center"><div className="w-16 h-16 border-4 border-blue-600 border-t-transparent rounded-full animate-spin shadow-lg"></div></div>}
          <ConfirmDialog {...confirmConfig} />

          <div className="flex flex-col lg:flex-row gap-3 lg:items-center lg:justify-between">
            <div className="flex gap-2 bg-white/60 backdrop-blur-md p-1.5 rounded-2xl shadow-sm border border-gray-200/50 self-start max-w-full overflow-x-auto whitespace-nowrap">
               <button onClick={()=>setActiveTab('dashboard')} className={`shrink-0 min-h-[44px] px-4 md:px-6 py-2.5 rounded-xl font-bold flex items-center gap-2 transition-all text-sm md:text-base ${activeTab==='dashboard'?'bg-white text-blue-600 shadow-sm':'text-gray-500 hover:text-gray-700 hover:bg-white/50'}`}><Icon path="M3 3v18h18 M18 17V9 M13 17V5 M8 17v-3" className="w-5 h-5"/> ジャーナル管理</button>
               <button onClick={()=>setActiveTab('vitals')} className={`shrink-0 min-h-[44px] px-4 md:px-6 py-2.5 rounded-xl font-bold flex items-center gap-2 transition-all text-sm md:text-base ${activeTab==='vitals'?'bg-gradient-to-r from-rose-600 to-pink-700 text-white shadow-md shadow-pink-600/20':'text-gray-500 hover:text-gray-700 hover:bg-white/50'}`}><Icon path={Icons.Pulse} className="w-5 h-5"/> 心のバイタル</button>
               <button onClick={()=>setActiveTab('admin')} className={`shrink-0 min-h-[44px] px-4 md:px-6 py-2.5 rounded-xl font-bold flex items-center gap-2 transition-all text-sm md:text-base ${activeTab==='admin'?'bg-white text-gray-800 shadow-sm':'text-gray-500 hover:text-gray-700 hover:bg-white/50'}`}>
                 <Icon path={Icons.Settings} className="w-5 h-5"/> クラス設定
                 {pendingMembers.length > 0 && <span className="bg-rose-600 text-white text-xs font-black px-2 py-0.5 rounded-full">{pendingMembers.length}</span>}
               </button>
            </div>

            {/* クラス切り替え */}
            <div className="flex items-center gap-2 self-start lg:self-auto">
              <select value={code} onChange={e => onSwitchTenant(e.target.value)} className="min-h-[44px] bg-white border border-gray-200 text-sm font-bold text-gray-700 outline-none rounded-xl px-3 py-2.5 cursor-pointer shadow-sm max-w-[220px]">
                {tenants.map(t => <option key={t.tenantCode} value={t.tenantCode}>{t.tenantName}</option>)}
              </select>
              <button onClick={onCreateNew} className="min-h-[44px] bg-white border border-gray-200 hover:border-blue-500 hover:text-blue-700 text-gray-600 rounded-xl px-3 py-2.5 shadow-sm font-bold text-sm flex items-center gap-1 transition-colors"><Icon path={Icons.Plus} className="w-4 h-4"/> クラス追加</button>
            </div>
          </div>

          {activeTab === 'dashboard' && (
            <div className="animate-fade-in flex flex-col gap-6 h-full">
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                 {/* サマリードーナツ */}
                 <div className="col-span-1 bg-white p-7 rounded-[2rem] shadow-premium border border-gray-100 flex flex-col relative overflow-hidden">
                    <div className="absolute -right-10 -top-10 w-40 h-40 bg-blue-50 rounded-full blur-3xl opacity-60"></div>
                    <h3 className="font-black text-gray-800 mb-6 flex items-center gap-2 text-lg z-10"><div className="p-1.5 bg-blue-100 text-blue-600 rounded-lg"><Icon path="M12 20v-6 M6 20V10 M18 20V4" className="w-5 h-5"/></div> 今日の提出率</h3>
                    <div className="flex items-center justify-center z-10 mb-5">
                      <div className="relative w-28 h-28">
                        <svg viewBox="0 0 100 100" className="w-full h-full transform -rotate-90 drop-shadow-sm">
                          <circle cx="50" cy="50" r="45" fill="none" stroke="#f1f5f9" strokeWidth="10"></circle>
                          <circle cx="50" cy="50" r="45" fill="none" stroke="#3b82f6" strokeWidth="10" strokeLinecap="round" strokeDasharray="283" strokeDashoffset={283 - (283 * pct) / 100} className="circle-chart-circle"></circle>
                        </svg>
                        <div className="absolute inset-0 flex flex-col items-center justify-center"><span className="text-2xl font-black text-gray-800">{pct}%</span></div>
                      </div>
                    </div>
                 </div>

                 {/* テーマ設定エリア */}
                 <div className="col-span-1 lg:col-span-2 bg-white p-7 rounded-[2rem] shadow-premium border border-gray-100 flex flex-col">
                    <h3 className="font-black text-gray-800 mb-6 flex items-center gap-2 text-lg"><div className="p-1.5 bg-orange-100 text-orange-600 rounded-lg"><Icon path="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7 M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" className="w-5 h-5"/></div> テーマ・スケジュール設定</h3>
                    <div className="flex flex-col sm:flex-row gap-3 mb-6">
                       <input type="text" value={todayThemeInput} onChange={e=>setTodayThemeInput(e.target.value)} className="flex-1 bg-gray-50 border border-gray-200 rounded-2xl px-5 py-3 font-bold outline-none focus:border-blue-400 focus:bg-white transition-colors text-base md:text-lg" placeholder="今日のテーマを入力"/>
                       <button onClick={handleSetTheme} className="bg-blue-600 hover:bg-blue-700 text-white font-bold px-8 py-3 rounded-2xl shadow-lg shadow-blue-500/30 active:scale-95 transition-all shrink-0">今すぐ適用</button>
                    </div>
                    <div className="mt-auto bg-gray-50/50 rounded-2xl p-5 border border-gray-100">
                      <details className="group">
                         <summary className="cursor-pointer text-sm font-bold text-gray-600 hover:text-blue-600 list-none flex items-center justify-between select-none">
                           <span className="flex items-center gap-2"><Icon path="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" className="w-5 h-5"/> 曜日ごとの自動テーマ（週間スケジュール）</span>
                           <Icon path="M19 9l-7 7-7-7" className="w-5 h-5 transition-transform duration-300 group-open:rotate-180"/>
                         </summary>
                         <div className="mt-4 pt-4 border-t border-gray-200">
                           <div className="grid grid-cols-2 lg:grid-cols-5 gap-3">
                              {['mon','tue','wed','thu','fri'].map((d,i) => (
                                 <div key={d}>
                                   {/* 入力欄をラベルの中に入れる。ラベル部分を押しても入力に入れるようになり、
                                       当たり判定も 44px を超える（input は疑似要素を持てないため） */}
                                   <label className="block">
                                     <span className="text-xs font-bold text-gray-600 block mb-1.5 px-1">{['月','火','水','木','金'][i]}曜日</span>
                                     <input type="text" value={weeklyThemes[d]} onChange={e=>setWeeklyThemes({...weeklyThemes, [d]: e.target.value})} className="w-full min-h-[44px] text-sm p-2.5 bg-white border border-gray-200 rounded-xl focus:border-blue-500 outline-none transition-colors" placeholder="テーマ"/>
                                   </label>
                                 </div>
                              ))}
                           </div>
                           <div className="mt-4 flex justify-end">
                             <button onClick={handleSaveWeekly} className="text-sm min-h-[44px] bg-gray-800 hover:bg-gray-900 text-white font-bold py-2.5 px-6 rounded-xl shadow-md active:scale-95 transition-all">スケジュールを保存</button>
                           </div>
                         </div>
                      </details>
                    </div>
                 </div>
              </div>

              {/* リストとツールバー */}
              <div className="flex-1 bg-white rounded-[2rem] shadow-premium border border-gray-100 overflow-hidden flex flex-col min-h-[500px]">
                 <div className="p-5 border-b border-gray-100 flex flex-col lg:flex-row gap-4 justify-between items-center bg-gray-50/50">
                    <div className="flex w-full lg:w-auto items-center gap-3">
                      <div className="relative flex-1 lg:w-64">
                        <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none text-gray-500"><Icon path={Icons.Search} className="w-5 h-5" /></div>
                        <input type="text" value={searchQuery} onChange={e=>setSearchQuery(e.target.value)} placeholder="名前や内容で検索..." className="w-full min-h-[44px] pl-11 pr-4 py-2.5 bg-white border border-gray-200 rounded-xl outline-none focus:border-blue-500 transition-colors font-medium text-sm"/>
                      </div>
                      <select value={filterStatus} onChange={e=>setFilterStatus(e.target.value)} className="min-h-[44px] bg-white border border-gray-200 text-sm font-bold text-gray-600 outline-none rounded-xl px-3 py-2.5 cursor-pointer">
                         <option value="all">全状況</option><option value="未返却">未返却</option><option value="返却済み">返却済</option>
                      </select>
                    </div>

                    <div className="flex w-full lg:w-auto items-center gap-2 overflow-x-auto pb-1 lg:pb-0">
                       <button onClick={()=>setConfirmConfig({isOpen:true, title:"AIコメント作成", text:"未返却のジャーナルに温かいコメント案を作ります。", onConfirm: ()=>executeAiAction('simple'), onCancel: ()=>setConfirmConfig({isOpen:false})})} className="shrink-0 min-h-[44px] bg-purple-50 text-purple-700 hover:bg-purple-100 font-bold py-2.5 px-4 rounded-xl text-sm transition-colors flex items-center gap-2"><Icon path={Icons.Sparkles} className="w-4 h-4"/> AIシンプル</button>
                       <button onClick={()=>setConfirmConfig({isOpen:true, title:"AI高度分析", text:"未返却のジャーナルを分析し、ハイライトとコメント案を作ります。", onConfirm: ()=>executeAiAction('full'), onCancel: ()=>setConfirmConfig({isOpen:false})})} className="shrink-0 min-h-[44px] bg-indigo-50 text-indigo-700 hover:bg-indigo-100 font-bold py-2.5 px-4 rounded-xl text-sm transition-colors flex items-center gap-2"><Icon path="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" className="w-4 h-4"/> AI高度分析</button>
                       <div className="w-px h-6 bg-gray-300 mx-1"></div>
                       <button onClick={()=>setConfirmConfig({isOpen:true, title:"一括返却", text:"コメントが付いている未返却ジャーナルを全て返却済みにしますか？", onConfirm: executeBatchReturn, onCancel: ()=>setConfirmConfig({isOpen:false})})} className="shrink-0 min-h-[44px] bg-emerald-700 hover:bg-emerald-800 text-white font-bold py-2.5 px-4 rounded-xl text-sm shadow-sm transition-colors flex items-center gap-2 active:scale-95"><Icon path="M5 13l4 4L19 7" className="w-4 h-4"/> 一括返却</button>
                    </div>
                 </div>

                 <div className="overflow-auto flex-1 p-0">
                    <table className="w-full min-w-[640px] text-left border-collapse">
                      <thead className="bg-gray-50/80 backdrop-blur sticky top-0 z-10 shadow-sm border-b border-gray-100">
                        <tr><th className="p-4 pl-6 font-black text-gray-500 text-xs tracking-wider uppercase">日付</th><th className="p-4 font-black text-gray-500 text-xs tracking-wider uppercase">児童名</th><th className="p-4 font-black text-gray-500 text-xs tracking-wider uppercase">状況</th><th className="p-4 pr-6 font-black text-gray-500 text-xs tracking-wider uppercase text-right">アクション</th></tr>
                      </thead>
                      <tbody className="divide-y divide-gray-100">
                        {filteredJournals.map(j => (
                          <tr key={j.journalId} className="hover:bg-blue-50/40 transition-colors group">
                            <td className="p-4 pl-6 text-sm font-bold text-gray-500 whitespace-nowrap">{j.date}</td>
                            <td className="p-4 font-black text-gray-800">{j.studentName}</td>
                            <td className="p-4">
                              <span className={`inline-flex items-center px-3 py-1.5 rounded-full text-xs font-bold tracking-wide ${j.status === '返却済み' ? 'bg-emerald-100 text-emerald-700' : 'bg-orange-100 text-orange-700 ring-1 ring-orange-400/30'}`}>
                                {j.status === '未返却' && <span className="w-1.5 h-1.5 rounded-full bg-orange-500 mr-1.5 animate-pulse"></span>}{j.status}
                              </span>
                            </td>
                            {/* タッチ端末ではホバーが使えないため、lg未満では常時表示にする */}
                            <td className="p-4 pr-6 flex gap-2 justify-end items-center">
                              {j.status === '未返却' ? (
                                <div className="flex bg-gray-50 border border-gray-100 rounded-xl p-1 mr-2 opacity-100 lg:opacity-0 lg:group-hover:opacity-100 transition-opacity">
                                  {['😊','👏','💪','⭐'].map(s => <button key={s} onClick={()=>handleQuickAction(j.journalId, 'stamp', s)} aria-label={`${s} のスタンプで返却する`} className="tap-44 hover:scale-125 hover:bg-white hover:shadow-sm rounded transition-all px-2 py-0.5 text-lg" title={`${s}で即返却`}>{s}</button>)}
                                </div>
                              ) : <button onClick={()=>handleQuickAction(j.journalId, 'revert')} className="tap-44 text-xs text-red-600 hover:bg-red-50 px-3 py-1.5 rounded-xl font-bold mr-2 transition-colors opacity-100 lg:opacity-0 lg:group-hover:opacity-100 focus:opacity-100">取消</button>}
                              <button onClick={() => {setActiveJournal({...j}); setFeedback(j.teacherComment || "");}} className="min-h-[44px] bg-white border-2 border-gray-200 hover:border-blue-500 hover:text-blue-700 text-gray-700 px-4 md:px-5 py-2 rounded-xl text-sm font-bold shadow-sm active:scale-95 transition-all whitespace-nowrap">詳細を開く</button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                 </div>
              </div>
            </div>
          )}

          {activeTab === 'vitals' && (
             <div className="flex flex-col gap-6 animate-fade-in h-full">
                <div className="bg-white rounded-[2rem] shadow-premium border border-gray-100 p-6">
                   <h3 className="font-black text-gray-800 mb-4 flex items-center gap-2 text-lg"><div className="p-1.5 bg-rose-100 text-rose-600 rounded-lg"><Icon path="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" className="w-5 h-5"/></div> 気にかけたい児童 (AI解析)</h3>
                   <div className="flex flex-wrap gap-4">
                     {vitalData.students.filter(s => s.alerts.length > 0).length === 0 ? (
                        <p className="text-gray-500 font-bold text-sm p-4 bg-gray-50 rounded-2xl border border-dashed border-gray-200 w-full text-center">現在、特筆すべきアラートはありません。クラスの心は安定しています✨</p>
                     ) : (
                        vitalData.students.filter(s => s.alerts.length > 0).map(s => (
                           <div key={s.email} className="flex-1 min-w-[250px] bg-rose-50/50 border border-rose-100 p-4 rounded-2xl">
                             <div className="flex justify-between items-center mb-2"><span className="font-black text-gray-800 text-lg">{s.name}</span><span className="text-xs font-bold text-rose-600 bg-rose-100 px-2 py-1 rounded-lg">{s.alerts.length}つの予兆</span></div>
                             <ul className="text-sm text-gray-600 space-y-1">{s.alerts.map((a,i) => <li key={i} className="flex items-center gap-1.5"><span className="w-1.5 h-1.5 rounded-full bg-rose-400"></span>{a}</li>)}</ul>
                           </div>
                        ))
                     )}
                   </div>
                </div>

                <div className="flex-1 bg-white rounded-[2rem] shadow-premium border border-gray-100 p-6 flex flex-col overflow-hidden">
                   <div className="flex flex-col lg:flex-row lg:justify-between lg:items-center gap-3 mb-6">
                     <h3 className="font-black text-gray-800 flex items-center gap-2 text-lg"><div className="p-1.5 bg-indigo-100 text-indigo-600 rounded-lg"><Icon path="M14 10l-2 1m0 0l-2-1m2 1v2.5M20 7l-2 1m2-1l-2-1m2 1v2.5M14 4l-2-1-2 1M4 7l2-1M4 7l2 1M4 7v2.5M12 21l-2-1m2 1l2-1m-2 1v-2.5M6 18l-2-1v-2.5M18 18l2-1v-2.5" className="w-5 h-5"/></div> クラスの心の波 (過去14日間)</h3>
                     <div className="flex flex-wrap gap-x-4 gap-y-1 text-xs font-bold text-gray-500 bg-gray-50 px-4 py-2 rounded-xl self-start lg:self-auto">
                        <div className="flex items-center gap-1"><span className="w-3 h-3 rounded-full bg-rose-400"></span>うれしい</div>
                        <div className="flex items-center gap-1"><span className="w-3 h-3 rounded-full bg-amber-400"></span>なるほど</div>
                        <div className="flex items-center gap-1"><span className="w-3 h-3 rounded-full bg-indigo-400"></span>もやもや</div>
                        <div className="flex items-center gap-1"><span className="w-3 h-3 rounded-full bg-violet-600"></span>くやしい</div>
                     </div>
                   </div>

                   <div className="overflow-x-auto pb-4 custom-scrollbar">
                     <div className="min-w-[800px]">
                       <div className="flex mb-3 pl-32">
                         {vitalData.dates.map((d, i) => <div key={i} className="flex-1 text-center text-[10px] font-bold text-gray-500 transform -rotate-45 origin-bottom-left whitespace-nowrap">{d.split('/')[1] + '/' + d.split('/')[2]}</div>)}
                       </div>
                       <div className="space-y-3">
                         {vitalData.students.map(s => (
                            <div key={s.email} className="flex items-center group/row">
                              <div className="w-32 pr-4 text-sm font-bold text-gray-700 truncate text-right group-hover/row:text-indigo-600 transition-colors">{s.name}</div>
                              <div className="flex flex-1 gap-2">
                                 {s.heatMap.map((h, i) => (
                                    <div key={i} className="flex-1 relative group flex justify-center">
                                       <div className={`w-6 h-6 rounded-full transition-all duration-300 ${h.status === 'none' ? 'bg-gray-100 hover:bg-gray-200' : h.color + ' hover:scale-125 cursor-pointer'}`}></div>
                                       {h.status !== 'none' && (
                                          <div className="tooltip-content absolute bottom-full left-1/2 -translate-x-1/2 mb-3 w-48 bg-gray-900/95 backdrop-blur text-white text-xs p-3 rounded-2xl shadow-xl z-50 pointer-events-none border border-gray-700">
                                            <div className="font-bold text-gray-300 mb-1 flex justify-between"><span>{h.date}</span><span>{h.emotion}</span></div>
                                            <p className="line-clamp-3 leading-relaxed">{h.content}</p>
                                            <div className="absolute -bottom-2 left-1/2 -translate-x-1/2 border-4 border-transparent border-t-gray-900/95"></div>
                                          </div>
                                       )}
                                    </div>
                                 ))}
                              </div>
                            </div>
                         ))}
                       </div>
                     </div>
                   </div>
                </div>
             </div>
          )}

          {activeTab === 'admin' && (
             <div className="flex flex-col gap-6 animate-fade-in h-full">

                {/* 🎫 児童の招待 */}
                <div className="bg-white p-8 rounded-[2rem] shadow-premium border border-gray-100">
                  <div className="flex flex-col md:flex-row gap-6 items-start">
                    <div className="flex-1 w-full">
                      <h3 className="font-black text-gray-800 flex items-center gap-2 text-lg mb-2">
                        <div className="p-1.5 bg-orange-100 text-orange-600 rounded-lg"><Icon path={Icons.Link} className="w-5 h-5"/></div> 児童をクラスに招待
                      </h3>
                      <p className="text-sm text-gray-500 mb-4 leading-relaxed">クラスコード、専用URL、QRコードのどれでも招待できます。児童はGoogleでサインインするだけで参加でき、スプレッドシートの権限は付与されません。</p>
                      <div className="bg-orange-50 border border-orange-200 rounded-2xl p-4 mb-4 flex flex-col sm:flex-row sm:items-center gap-3">
                        <div className="flex-1">
                          <p className="text-xs font-bold text-orange-700 mb-1">クラスコード</p>
                          <p className="font-mono text-2xl font-black tracking-[0.2em] text-gray-900">{data.tenant.tenantCode}</p>
                        </div>
                        <button onClick={handleCopyCode} className={`min-h-[44px] shrink-0 font-bold px-4 py-2.5 rounded-xl shadow-sm active:scale-95 transition-all flex items-center justify-center gap-2 ${copiedItem === 'code' ? 'bg-emerald-600 text-white' : 'bg-white hover:bg-orange-100 text-orange-800 border border-orange-200'}`}><Icon path={copiedItem === 'code' ? Icons.Check : Icons.Copy} className="w-4 h-4"/> {copiedItem === 'code' ? 'コピーしました' : 'コードをコピー'}</button>
                      </div>
                      <div className="flex flex-col sm:flex-row gap-2 mb-4">
                        <label className="flex-1"><span className="sr-only">児童用URL</span><input type="text" readOnly value={data.tenant.memberUrl} onFocus={e=>e.target.select()} className="w-full bg-gray-50 border border-gray-200 rounded-xl px-4 py-3 text-sm font-mono text-gray-600 outline-none focus:ring-2 focus:ring-blue-500"/></label>
                        <button onClick={handleCopyUrl} className={`shrink-0 font-bold px-6 py-3 rounded-xl shadow-sm active:scale-95 transition-all flex items-center justify-center gap-2 ${copiedItem === 'url' ? 'bg-emerald-600 text-white' : 'bg-gray-800 hover:bg-black text-white'}`}><Icon path={copiedItem === 'url' ? Icons.Check : Icons.Copy} className="w-4 h-4"/> {copiedItem === 'url' ? 'コピーしました' : 'URLをコピー'}</button>
                      </div>
                      <div className="flex flex-wrap gap-2 mb-4">
                        <button onClick={handleCopyInvitation} className={`min-h-[44px] text-sm font-bold rounded-xl px-4 py-2.5 transition-colors flex items-center gap-2 ${copiedItem === 'invitation' ? 'bg-emerald-600 text-white' : 'bg-blue-50 hover:bg-blue-100 text-blue-800 border border-blue-200'}`}><Icon path={copiedItem === 'invitation' ? Icons.Check : Icons.Copy} className="w-4 h-4"/> {copiedItem === 'invitation' ? 'コピーしました' : '招待文をコピー'}</button>
                        <a href={data.tenant.memberUrl} target="_blank" rel="noopener noreferrer" className="min-h-[44px] text-sm font-bold text-gray-700 hover:text-blue-800 bg-gray-50 hover:bg-blue-50 border border-gray-200 rounded-xl px-4 py-2.5 transition-colors flex items-center gap-2"><Icon path={Icons.Link} className="w-4 h-4"/> 児童画面を確認</a>
                      </div>
                      <div className="flex flex-wrap gap-3 items-center">
                        <label className="flex items-center gap-2 min-h-[44px] text-sm font-bold text-gray-600 cursor-pointer bg-gray-50 border border-gray-200 rounded-xl px-4 py-2.5">
                          <input type="checkbox" checked={!!data.tenant.joinOpen} onChange={e=>handleJoinPolicy(e.target.checked, data.tenant.requireApproval)} className="accent-blue-600 w-5 h-5"/> 参加を受け付ける
                        </label>
                        <label className="flex items-center gap-2 min-h-[44px] text-sm font-bold text-gray-600 cursor-pointer bg-gray-50 border border-gray-200 rounded-xl px-4 py-2.5">
                          <input type="checkbox" checked={!!data.tenant.requireApproval} onChange={e=>handleJoinPolicy(data.tenant.joinOpen, e.target.checked)} className="accent-blue-600 w-5 h-5"/> 参加は先生の承認制にする
                        </label>
                        <button onClick={()=>setConfirmConfig({isOpen:true, title:"コードの再発行", text:"クラスコードを作り直します。\n今までのURLは使えなくなります。よろしいですか？", confirmText:"再発行する", onConfirm: handleRegenerateCode, onCancel: ()=>setConfirmConfig({isOpen:false})})} className="min-h-[44px] text-sm font-bold text-gray-600 hover:text-red-600 bg-gray-50 hover:bg-red-50 border border-gray-200 rounded-xl px-4 py-2.5 transition-colors flex items-center gap-2"><Icon path={Icons.Refresh} className="w-4 h-4"/> コード再発行</button>
                      </div>
                    </div>
                    <div className="shrink-0 mx-auto md:mx-0 text-center flex flex-col items-center">
                      <QrCode model={inviteQr} />
                      <p className="text-xs font-bold text-gray-500 mt-2 mb-2">読み取ると児童用URLが開きます</p>
                      <button onClick={handleDownloadQr} disabled={!inviteQr} className="min-h-[44px] text-sm font-bold text-gray-700 hover:text-blue-800 bg-white hover:bg-blue-50 border border-gray-200 rounded-xl px-4 py-2.5 transition-colors flex items-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed"><Icon path={Icons.Download} className="w-4 h-4"/> QR画像を保存</button>
                    </div>
                  </div>
                </div>

                {/* ✋ 参加申請 */}
                {pendingMembers.length > 0 && (
                  <div className="bg-amber-50 p-8 rounded-[2rem] shadow-premium border border-amber-200">
                    <h3 className="font-black text-amber-800 flex items-center gap-2 text-lg mb-4">
                      <span className="text-2xl">✋</span> 参加申請（{pendingMembers.length}件）
                    </h3>
                    <div className="flex flex-col gap-3">
                      {pendingMembers.map(m => (
                        <div key={m.email} className="bg-white rounded-2xl border border-amber-100 p-4 flex flex-col sm:flex-row sm:items-center gap-3">
                          <div className="flex-1 min-w-0">
                            <p className="font-black text-gray-800">{m.name}</p>
                            <p className="text-xs text-gray-500 font-mono truncate">{m.email}</p>
                          </div>
                          <div className="flex gap-2 shrink-0">
                            <button onClick={()=>handleApprove(m.email, true)} className="bg-emerald-700 hover:bg-emerald-800 text-white font-bold px-5 py-2.5 rounded-xl shadow-sm active:scale-95 transition-all flex items-center gap-1.5"><Icon path={Icons.Check} className="w-4 h-4"/> 承認</button>
                            <button onClick={()=>handleApprove(m.email, false)} className="bg-white hover:bg-red-50 text-red-600 border border-red-300 font-bold px-5 py-2.5 rounded-xl active:scale-95 transition-all">却下</button>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {/* 👥 名簿管理エディタ */}
                <div className="bg-white p-8 rounded-[2rem] shadow-premium border border-gray-100 flex flex-col">
                  <div className="flex justify-between items-center mb-6">
                    <h3 className="font-black text-gray-800 flex items-center gap-2 text-lg">
                      <div className="p-1.5 bg-blue-100 text-blue-600 rounded-lg"><Icon path={Icons.User} className="w-5 h-5"/></div> 児童名簿の管理
                    </h3>
                    <button onClick={handleSaveRoster} className="min-h-[44px] bg-blue-600 hover:bg-blue-700 text-white font-bold py-2.5 px-6 rounded-xl shadow-md active:scale-95 transition-all">保存して更新</button>
                  </div>

                  <div className="overflow-auto max-h-[400px] bg-gray-50/50 rounded-2xl border border-gray-200 p-2 custom-scrollbar">
                    <table className="w-full min-w-[640px] text-left">
                      <thead className="text-xs text-gray-600 font-bold bg-gray-100/80 rounded-xl sticky top-0 z-10">
                        <tr>
                           <th className="p-3 pl-4 rounded-l-xl w-24">役割</th>
                           <th className="p-3">氏名</th>
                           <th className="p-3">メールアドレス <span className="text-[10px] font-normal">(Googleアカウント)</span></th>
                           <th className="p-3 w-24">状態</th>
                           <th className="p-3 pr-4 rounded-r-xl text-right">操作</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-100">
                        {rosterData.map((r, i) => (
                           <tr key={i} className="hover:bg-white transition-colors group">
                             <td className="p-2 pl-4">
                               <select value={r.role} onChange={(e)=>handleRosterChange(i, 'role', e.target.value)} className="w-full min-h-[44px] bg-transparent border-none outline-none font-bold text-gray-700 cursor-pointer">
                                 <option value="児童">児童</option>
                                 <option value="担任">担任</option>
                               </select>
                             </td>
                             <td className="p-2">
                               <input type="text" value={r.name} onChange={(e)=>handleRosterChange(i, 'name', e.target.value)} placeholder="例: 山田 太郎" className="w-full min-h-[44px] bg-transparent border-none outline-none font-bold text-gray-800 focus:ring-2 focus:ring-blue-200 rounded px-2 py-1.5 transition-all"/>
                             </td>
                             <td className="p-2">
                               <input type="text" value={r.email} onChange={(e)=>handleRosterChange(i, 'email', e.target.value)} placeholder="例: yamada@school.ed.jp" className="w-full min-h-[44px] bg-transparent border-none outline-none text-gray-700 text-sm focus:ring-2 focus:ring-blue-200 rounded px-2 py-1.5 transition-all font-mono"/>
                             </td>
                             <td className="p-2">
                               <span className={`inline-flex px-2.5 py-1 rounded-full text-xs font-bold ${r.status === 'pending' ? 'bg-amber-100 text-amber-700' : 'bg-emerald-100 text-emerald-700'}`}>{r.status === 'pending' ? '承認待ち' : '参加中'}</span>
                             </td>
                             <td className="p-2 pr-4 text-right">
                               <button onClick={()=>handleRemoveRoster(i)} aria-label={`${r.name || "この行"} を名簿から削除する`} className="tap-44 p-2 text-gray-500 hover:text-red-500 hover:bg-red-50 rounded-xl transition-colors opacity-100 lg:opacity-0 lg:group-hover:opacity-100 focus:opacity-100"><Icon path={Icons.Trash} className="w-4 h-4"/></button>
                             </td>
                           </tr>
                        ))}
                      </tbody>
                    </table>
                    <div className="p-2 mt-2">
                      <button onClick={handleAddRoster} className="w-full py-3 border-2 border-dashed border-blue-200 text-blue-600 font-bold rounded-xl hover:bg-blue-50 hover:border-blue-400 transition-all flex items-center justify-center gap-2 active:scale-[0.99]">
                        <Icon path={Icons.Plus} className="w-5 h-5"/> あたらしい行を追加する
                      </button>
                    </div>
                  </div>
                  <p className="text-xs text-gray-500 mt-3 leading-relaxed">💡 名簿に登録した児童は、児童用URLからサインインするだけで参加できます（承認不要）。名簿に無いアカウントは「参加申請」として届きます。</p>
               </div>

               <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-6">
                 {/* AI設定 */}
                 <div className="bg-white p-8 rounded-[2rem] shadow-premium border border-gray-100 flex flex-col">
                    <div className="w-16 h-16 bg-purple-50 rounded-full flex items-center justify-center mb-4 text-purple-500 mx-auto"><Icon path={Icons.Sparkles} className="w-8 h-8"/></div>
                    <h3 className="font-black text-gray-800 mb-2 text-center">Gemini AI 設定</h3>
                    <p className="text-sm text-gray-500 mb-4 text-center">{data.settings && data.settings.hasApiKey ? `設定済み: ${data.settings.apiKeyMasked}` : 'AIコメント機能にはAPIキーが必要です。'}</p>
                    <input type="password" value={apiKeyInput} onChange={e=>setApiKeyInput(e.target.value)} placeholder="Gemini APIキーを貼り付け" className="w-full bg-gray-50 border border-gray-200 rounded-xl px-4 py-3 text-sm font-mono outline-none focus:border-purple-400 mb-3"/>
                    <button onClick={handleSaveApiKey} className="w-full bg-purple-600 hover:bg-purple-700 text-white font-bold py-3 rounded-xl transition-all active:scale-95 mt-auto">保存</button>
                 </div>

                 {/* データベース */}
                 <div className="bg-white p-8 rounded-[2rem] shadow-premium border border-gray-100 flex flex-col justify-center items-center text-center">
                    <div className="w-16 h-16 bg-emerald-50 rounded-full flex items-center justify-center mb-4 text-emerald-500"><Icon path={Icons.Link} className="w-8 h-8"/></div>
                    <h3 className="font-black text-gray-800 mb-2">データベース</h3>
                    <p className="text-sm text-gray-500 mb-6">このクラスの記録は、あなたのDrive内のスプレッドシートに保存されています。</p>
                    <a href={data.spreadsheetUrl} target="_blank" rel="noopener noreferrer" className="w-full bg-emerald-50 hover:bg-emerald-100 text-emerald-700 font-bold py-3 rounded-xl transition-all active:scale-95">シートを開く</a>
                 </div>

                 {/* エクスポート */}
                 <div className="bg-white p-8 rounded-[2rem] shadow-premium border border-gray-100 flex flex-col justify-center items-center text-center">
                    <div className="w-16 h-16 bg-gray-100 rounded-full flex items-center justify-center mb-4 text-gray-600"><Icon path={Icons.Download} className="w-8 h-8"/></div>
                    <h3 className="font-black text-gray-800 mb-2">データの一括出力</h3>
                    <p className="text-sm text-gray-500 mb-6">CSVをダウンロード、または印刷用の帳票（PDF保存可）を開きます。</p>
                    <div className="flex gap-2 w-full">
                       <button onClick={()=>handleExport('csv')} className="flex-1 bg-gray-800 hover:bg-black text-white font-bold py-3 rounded-xl transition-all active:scale-95">CSV</button>
                       <button onClick={()=>handleExport('pdf')} className="flex-1 bg-red-50 hover:bg-red-100 text-red-700 font-bold py-3 rounded-xl transition-all active:scale-95">印刷/PDF</button>
                    </div>
                 </div>

                 {/* データ削除 */}
                 <div className="bg-red-50 p-8 rounded-[2rem] shadow-premium border border-red-100 flex flex-col justify-center items-center text-center">
                    <div className="w-16 h-16 bg-red-100 rounded-full flex items-center justify-center mb-4 text-red-500"><Icon path={Icons.Trash} className="w-8 h-8"/></div>
                    <h3 className="font-black text-red-800 mb-2">危険な操作</h3>
                    <p className="text-sm text-red-700 mb-6">このクラスの全ジャーナルデータを削除します。年度替わりなどに使用してください。</p>
                    <button onClick={()=>setConfirmConfig({isOpen:true, isDanger:true, title:"データ全削除", text:"本当にこのクラスのすべてのジャーナルデータを削除しますか？\nこの操作は元に戻せません。", confirmText:"削除する", onConfirm: executeDataReset, onCancel: ()=>setConfirmConfig({isOpen:false})})} className="w-full bg-red-600 hover:bg-red-700 text-white font-bold py-3 rounded-xl shadow-lg shadow-red-600/30 active:scale-95 transition-all">全データを削除</button>
                 </div>
               </div>
             </div>
          )}

          {activeJournal && (
            <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-50 flex items-center justify-center p-2 md:p-8 transition-opacity">
              <div role="dialog" aria-modal="true" aria-label={`${activeJournal.studentName} さんのジャーナル`}
                   className="bg-white rounded-[2rem] shadow-2xl w-full max-w-5xl h-full max-h-[92vh] md:max-h-[85vh] flex flex-col md:flex-row overflow-hidden animate-fade-in border border-gray-100">
                <div className="flex-1 min-h-[35%] bg-gradient-to-br from-orange-50/50 to-amber-50/50 p-4 md:p-8 border-b md:border-b-0 md:border-r overflow-y-auto relative" onMouseUp={handleTextSelect} onTouchEnd={handleTextSelect}>
                  <div className="flex justify-between items-start mb-6">
                    <div><p className="text-sm font-bold text-gray-500 mb-1">{activeJournal.date}</p><h2 className="text-3xl font-black text-gray-800">{activeJournal.studentName} <span className="text-xl font-medium text-gray-500">さん</span></h2></div>
                    <span className="text-3xl bg-white px-4 py-2 rounded-2xl shadow-sm border border-orange-100">{activeJournal.emotion || '📝'}</span>
                  </div>
                  <div className="bg-white p-6 md:p-8 rounded-3xl shadow-sm border border-orange-100 min-h-[250px] whitespace-pre-wrap text-lg lined-paper text-gray-800 leading-[40px] font-medium select-text relative">
                     {(() => {
                        const hls = parseHighlights(activeJournal.highlights);
                        if (hls.length === 0) return activeJournal.content;
                        const html = buildHighlightedHtml(activeJournal.content, hls,
                          (escapedText, h) => `<mark class="highlight-marker group relative" data-id="${escapeHtml(h.id)}">${escapedText}<span class="absolute bottom-full left-1/2 -translate-x-1/2 mb-2 hidden group-hover:block bg-gray-800 text-white text-xs px-3 py-1.5 rounded-lg whitespace-nowrap z-10 font-bold shadow-lg">${escapeHtml(h.suggestedStamp||'')} ${escapeHtml(h.suggestedComment||'編集可')}</span></mark>`);
                        return <div dangerouslySetInnerHTML={{__html: html}} />;
                     })()}
                  </div>
                  <LazyImage imageId={activeJournal.imageId} fetcher={(id) => window.callOwnerApi('opGetImage', code, id)} alt={`${activeJournal.studentName} さんがのせた画像`} className="mt-4 rounded-xl border max-w-xs shadow-sm"/>
                </div>
                <div className="w-full md:w-[420px] shrink-0 max-h-[55%] md:max-h-none p-5 md:p-8 bg-white flex flex-col relative overflow-y-auto">
                  <button onClick={() => setActiveJournal(null)} aria-label="閉じる" className="tap-44 absolute top-6 right-6 p-2 text-gray-500 hover:text-gray-700 bg-gray-50 rounded-full hover:bg-gray-100 transition-colors"><Icon path={Icons.Close} /></button>
                  <h3 className="text-lg font-black text-blue-600 mb-4 flex items-center gap-2"><div className="p-1.5 bg-blue-50 rounded-lg"><Icon path={Icons.Pen} className="w-5 h-5"/></div> 全体コメント</h3>
                  <textarea value={feedback} onChange={e => setFeedback(e.target.value)} placeholder="温かいコメントを入力..." className="w-full border-2 border-gray-100 rounded-2xl p-4 text-gray-700 outline-none focus:border-blue-400 focus:bg-blue-50/30 transition-colors resize-none mb-6 min-h-[160px] font-medium" />

                  <h3 className="text-sm font-black text-gray-700 mb-3 flex items-center gap-2"><Icon path={Icons.Link} className="w-4 h-4 text-yellow-500" /> 抽出したハイライト</h3>
                  <div className="flex-1 overflow-y-auto space-y-3 mb-6 pr-1 custom-scrollbar">
                     {parseHighlights(activeJournal.highlights).map(h => (
                        <div key={h.id} className="bg-yellow-50/50 border border-yellow-200 p-4 rounded-2xl relative group hover:bg-yellow-50 transition-colors">
                           <button onClick={()=>removeHighlight(h.id)} aria-label="このハイライトを消す" className="tap-44 absolute top-3 right-3 text-yellow-600 hover:text-red-600 bg-white rounded-full p-1 opacity-100 lg:opacity-0 lg:group-hover:opacity-100 focus:opacity-100 transition-opacity shadow-sm"><Icon path={Icons.Close} className="w-3 h-3"/></button>
                           <p className="text-sm font-bold text-gray-800 mb-3 border-l-4 border-yellow-400 pl-3 leading-relaxed">"{h.textToHighlight}"</p>
                           <input type="text" placeholder="コメント (任意)" value={h.suggestedComment||''} onChange={e=>updateHighlight(h.id, 'suggestedComment', e.target.value)} className="w-full text-xs p-2.5 rounded-xl border border-yellow-200 outline-none focus:border-yellow-400 mb-3 font-medium bg-white" />
                           <div className="flex gap-1.5 overflow-x-auto pb-1">
                              {['😄','👏','🤔','😭','👍','✨','💪','🌱'].map(s=>(<button key={s} onClick={()=>updateHighlight(h.id, 'suggestedStamp', s)} aria-label={`スタンプ ${s} を選ぶ`} className={`tap-44 shrink-0 text-base p-1.5 rounded-lg hover:bg-yellow-100 transition-colors ${h.suggestedStamp===s?'bg-yellow-200 border-yellow-400 shadow-sm transform scale-110':''}`}>{s}</button>))}
                           </div>
                        </div>
                     ))}
                     {parseHighlights(activeJournal.highlights).length === 0 && <div className="text-xs font-bold text-gray-500 text-center py-6 bg-gray-50 rounded-2xl border border-gray-100 border-dashed">ハイライトはありません</div>}
                  </div>

                  <button onClick={handleSaveFeedback} className="mt-auto bg-blue-600 hover:bg-blue-700 text-white font-black py-4 rounded-[2rem] shadow-lg shadow-blue-500/30 text-lg active:scale-95 transition-all w-full flex justify-center items-center gap-2"><Icon path={Icons.Send} /> 保存して返却する</button>
                </div>
              </div>
            </div>
          )}
          {toast && <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />}
        </div>
      );
    };

    // ==========================================
    // 🏫 クラス作成（先生の初回オンボーディング / クラス追加）
    // ==========================================
    const CreateClassScreen = ({ server, onCreated, onCancel, ownerEmail }) => {
      const { owner, isLoading, toast, setToast, showToast } = server;
      const [className, setClassName] = useState("");
      const [adoptUrl, setAdoptUrl] = useState("");
      const [done, setDone] = useState(null);

      const handleCreate = async () => {
        if (!className.trim()) return showToast('クラス名を入力してください', 'error');
        const res = await owner('opCreateTenant', className.trim());
        if (res && res.success) { setDone(res); setTimeout(() => onCreated(res.tenantCode), 2500); }
        else showToast(failMsg(res, 'クラスを作成できませんでした'), 'error');
      };

      const handleAdopt = async () => {
        if (!className.trim()) return showToast('クラス名を入力してください', 'error');
        if (!adoptUrl.trim()) return showToast('スプレッドシートのURLを入力してください', 'error');
        const res = await owner('opAdoptTenant', adoptUrl.trim(), className.trim());
        if (res && res.success) { setDone(res); setTimeout(() => onCreated(res.tenantCode), 2500); }
        else showToast(failMsg(res, '取り込みできませんでした'), 'error');
      };

      return (
        <div className="h-screen w-full flex items-center justify-center bg-gradient-to-br from-orange-50 via-amber-50 to-blue-50 p-4 overflow-y-auto">
          {isLoading && <div className="fixed inset-0 bg-white/50 backdrop-blur-sm z-40 flex items-center justify-center"><div className="w-16 h-16 border-4 border-orange-500 border-t-transparent rounded-full animate-spin shadow-lg"></div></div>}
          <div className="bg-white rounded-[2rem] shadow-premium border border-gray-100 p-8 md:p-10 max-w-xl w-full animate-fade-in my-8">
            {done ? (
              <div className="text-center py-4">
                <div className="w-16 h-16 mx-auto rounded-full bg-emerald-50 text-emerald-500 flex items-center justify-center mb-4"><Icon path={Icons.Check} className="w-8 h-8"/></div>
                <p className="font-black text-gray-700 text-lg mb-2">クラス「{done.tenantName}」を作成しました！</p>
                <p className="text-sm text-gray-500 font-medium mb-4">クラスコード: <span className="font-mono font-black tracking-widest text-orange-600">{done.tenantCode}</span></p>
                <p className="text-xs text-gray-500">ダッシュボードを読み込んでいます…</p>
              </div>
            ) : (
              <>
                <div className="text-center mb-8">
                  <div className="w-20 h-20 mx-auto rounded-3xl bg-gradient-to-br from-orange-400 to-orange-500 text-white flex items-center justify-center mb-4 shadow-lg shadow-orange-500/30"><Icon path={Icons.Book} className="w-10 h-10"/></div>
                  <h1 className="text-3xl font-black text-gray-800 mb-2">クラスをつくろう</h1>
                  <p className="text-gray-500 font-medium leading-relaxed">クラス専用のデータベースが<br/>あなたの Google Drive に自動作成されます。</p>
                  {ownerEmail && <p className="mt-3 text-xs font-bold text-gray-500 bg-gray-50 inline-block px-4 py-1.5 rounded-full">ログイン中: {ownerEmail}</p>}
                </div>

                <div className="border-2 border-orange-100 bg-orange-50/40 rounded-3xl p-6 mb-5">
                  <h2 className="font-black text-gray-800 mb-1 flex items-center gap-2"><Icon path={Icons.Sparkles} className="w-5 h-5 text-orange-500"/> 新しいクラスを作成</h2>
                  <p className="text-sm text-gray-500 mb-4 leading-relaxed">作成すると「児童用URL」が発行されます。URLを配るだけで児童が参加できます。</p>
                  <input type="text" value={className} onChange={e=>setClassName(e.target.value)} placeholder="クラス名（例: 6年1組）" className="w-full border-2 border-orange-100 rounded-2xl p-3.5 text-sm outline-none focus:border-orange-400 bg-white font-bold mb-3 transition-colors"/>
                  <button onClick={handleCreate} disabled={isLoading} className="w-full bg-gradient-to-r from-orange-500 to-amber-500 hover:from-orange-600 hover:to-amber-600 text-white font-black py-4 rounded-2xl shadow-lg shadow-orange-500/30 active:scale-95 transition-all text-lg disabled:opacity-50">🎒 クラスを作成する</button>
                </div>

                <details className="border-2 border-blue-100 bg-blue-50/40 rounded-3xl p-6">
                  <summary className="font-black text-gray-800 cursor-pointer flex items-center gap-2 list-none"><Icon path={Icons.Link} className="w-5 h-5 text-blue-500"/> 以前のスプレッドシートを取り込む（引っ越し）</summary>
                  <p className="text-sm text-gray-500 my-3 leading-relaxed">旧バージョンで使っていたDBスプレッドシート（あなたが編集権限を持つもの）をクラスとして登録できます。</p>
                  <input type="text" value={adoptUrl} onChange={e=>setAdoptUrl(e.target.value)} placeholder="https://docs.google.com/spreadsheets/d/..." className="w-full border-2 border-blue-100 rounded-2xl p-3.5 text-sm outline-none focus:border-blue-400 bg-white font-mono mb-3 transition-colors"/>
                  <button onClick={handleAdopt} disabled={isLoading} className="w-full bg-blue-600 hover:bg-blue-700 text-white font-black py-3.5 rounded-2xl shadow-lg shadow-blue-500/30 active:scale-95 transition-all disabled:opacity-50">🔗 このスプレッドシートを取り込む</button>
                </details>

                {onCancel && <button onClick={onCancel} className="w-full mt-5 py-3 rounded-2xl font-bold bg-gray-50 text-gray-500 hover:bg-gray-100 transition-colors">もどる</button>}
              </>
            )}
          </div>
          {toast && <Toast message={toast.message} type={toast.type} onClose={() => setToast(null)} />}
        </div>
      );
    };

    // ==========================================
    // 🧑‍🏫 先生ポータルのルート（デプロイ A）
    // ==========================================
    const OwnerRoot = () => {
      const server = useServer();
      const [phase, setPhase] = useState('loading');   // loading | create | app | error
      const [tenants, setTenants] = useState([]);
      const [ownerEmail, setOwnerEmail] = useState('');
      const [selectedCode, setSelectedCode] = useState(null);
      const [tenantData, setTenantData] = useState(null);
      const [errorMsg, setErrorMsg] = useState('');

      const loadTenants = async (preferCode) => {
        const res = await window.callOwnerApi('opListTenants').catch(e => ({ success: false, error: e.message }));
        if (!res || !res.success) {
          setErrorMsg(res && res.error || 'ポータルに接続できませんでした');
          setPhase('error');
          return;
        }
        setOwnerEmail(res.ownerEmail || '');
        setTenants(res.tenants || []);
        if (!res.tenants || res.tenants.length === 0) { setPhase('create'); return; }
        let code = preferCode || localStorage.getItem('rj_selectedClass');
        if (!res.tenants.some(t => t.tenantCode === code)) code = res.tenants[0].tenantCode;
        // 同じクラスを選び直した場合は useEffect が発火しないため直接読み込む
        if (code === selectedCode) loadTenantData(code);
        else setSelectedCode(code);
      };

      const loadTenantData = async (code) => {
        setPhase('loading');
        const res = await server.owner('opGetTenantData', code);
        if (res && res.success) {
          setTenantData(res);
          localStorage.setItem('rj_selectedClass', code);
          setPhase('app');
        } else {
          setErrorMsg(failMsg(res, 'クラスのデータを読み込めませんでした'));
          setPhase('error');
        }
      };

      useEffect(() => { loadTenants(); }, []);
      useEffect(() => { if (selectedCode) loadTenantData(selectedCode); }, [selectedCode]);

      if (phase === 'loading') return <FullScreenSpinner color="border-blue-500" />;

      if (phase === 'error') return (
        <CenterCard icon="⚠️">
          <h1 className="text-2xl font-black text-gray-800 mb-3">読み込めませんでした</h1>
          <p className="text-sm text-gray-500 font-medium leading-relaxed mb-6 whitespace-pre-wrap">{errorMsg}</p>
          <button onClick={()=>{ setPhase('loading'); loadTenants(); }} className="w-full bg-blue-600 hover:bg-blue-700 text-white font-black py-3.5 rounded-2xl shadow-lg active:scale-95 transition-all">もう一度読み込む</button>
        </CenterCard>
      );

      if (phase === 'create') return (
        <CreateClassScreen
          server={server}
          ownerEmail={ownerEmail}
          onCreated={(code) => { loadTenants(code); }}
          onCancel={tenants.length > 0 ? () => setSelectedCode(tenants[0].tenantCode) : null}
        />
      );

      return (
        <div className="h-screen w-full flex flex-col bg-gray-50 text-gray-800 font-sans selection:bg-orange-200">
          <header className="flex-none bg-white/80 backdrop-blur-md z-30 px-3 py-2.5 md:p-4 md:px-6 flex justify-between items-center gap-2 border-b border-blue-500 shadow-[0_2px_10px_rgba(0,0,0,0.02)]">
            <div className="flex items-center gap-2 md:gap-3 min-w-0">
              <div className="p-2 md:p-2.5 rounded-2xl text-white shadow-sm shrink-0 bg-gradient-to-br from-blue-500 to-blue-600"><Icon path={Icons.Book} className="w-5 h-5 md:w-6 md:h-6"/></div>
              <h1 className="text-lg md:text-2xl font-black tracking-tight truncate text-blue-600">教員ダッシュボード <span aria-hidden="true" className="hidden sm:inline-block w-px h-5 bg-gray-300 align-middle"></span> <span className="text-gray-600 text-base md:text-xl hidden sm:inline">{tenantData.tenant.tenantName}</span></h1>
            </div>
            <div className="flex items-center gap-1.5 md:gap-2 bg-gray-100/80 px-3 md:px-4 py-1.5 md:py-2 rounded-2xl font-bold text-xs md:text-sm text-gray-600 shrink-0 max-w-[45%]"><Icon path={Icons.User} className="w-4 h-4 text-gray-500 shrink-0" /> <span className="truncate">{ownerEmail} 先生</span></div>
          </header>
          <main className="flex-1 overflow-y-auto p-3 md:p-6 lg:p-8"><div className="max-w-[1400px] mx-auto h-full">
            <TeacherApp
              data={tenantData}
              tenants={tenants}
              server={server}
              refresh={() => loadTenantData(selectedCode)}
              onSwitchTenant={(code) => setSelectedCode(code)}
              onCreateNew={() => setPhase('create')}
            />
          </div></main>
          <footer className="flex-none w-full text-center text-gray-500 py-3 bg-white border-t border-gray-100 text-sm">
            <span>© 2026 ふりかえりジャーナル <a href="https://giga-school.com" target="_blank" rel="noopener noreferrer" className="tap-44 inline-block no-underline text-inherit hover:opacity-80 transition-opacity">GIGA山</a></span>
          </footer>
        </div>
      );
    };

    // ==========================================
    // 🎒 児童アプリのルート（デプロイ B）
    //    シェルから ID トークン + クラスコードを受け取り、
    //    参加状態に応じて 参加申請 / 承認待ち / 本体 を出し分ける
    // ==========================================
    const MemberRoot = () => {
      const server = useServer();
      const [phase, setPhase] = useState('connecting');  // connecting | join | pending | app | error
      const [statusInfo, setStatusInfo] = useState({});
      const [data, setData] = useState(null);
      const [joinName, setJoinName] = useState('');
      const [errorMsg, setErrorMsg] = useState('');

      const sync = async () => {
        const res = await window.callMemberApi('mbSync').catch(e => ({ success: false, error: e.message }));
        if (res && res.success) {
          setData({ user: res.me, todayTheme: res.todayTheme, journals: res.journals, tenantName: res.tenantName });
          setPhase('app');
          return;
        }
        if (res && (res.code === 'NOT_MEMBER' || res.code === 'MEMBER_PENDING')) { await checkStatus(); return; }
        setErrorMsg(failMsg(res, 'つながりませんでした。もう一度ためしてね'));
        setPhase('error');
      };

      const checkStatus = async () => {
        const res = await window.callMemberApi('mbGetStatus').catch(e => ({ success: false, error: e.message }));
        if (!res || !res.success) {
          setErrorMsg(failMsg(res, 'つながりませんでした。もう一度ためしてね'));
          setPhase('error');
          return;
        }
        setStatusInfo(res);
        if (res.status === 'active') { await sync(); }
        else if (res.status === 'pending') setPhase('pending');
        else setPhase('join');
      };

      useEffect(() => { checkStatus(); }, []);

      const handleJoin = async () => {
        if (!joinName.trim()) return server.showToast('名前を入力してね', 'error');
        const res = await server.memberWrite('mbRequestJoin', joinName.trim());
        if (res && res.success) {
          if (res.status === 'active') { server.showToast('参加できたよ！'); await sync(); }
          else setPhase('pending');
        } else server.showToast(failMsg(res, '参加できませんでした'), 'error');
      };

      if (phase === 'connecting') return <FullScreenSpinner />;

      if (phase === 'error') return (
        <CenterCard icon="😢">
          <h1 className="text-2xl font-black text-gray-800 mb-3">つながりませんでした</h1>
          <p className="text-sm text-gray-500 font-medium leading-relaxed mb-6 whitespace-pre-wrap">{errorMsg}</p>
          <button onClick={()=>window.appReload()} className="w-full bg-orange-500 hover:bg-orange-600 text-white font-black py-3.5 rounded-2xl shadow-lg active:scale-95 transition-all">もう<ruby>一度<rp>(</rp><rt className="text-[0.6em]">いちど</rt><rp>)</rp></ruby>ひらく</button>
        </CenterCard>
      );

      if (phase === 'join') return (
        <CenterCard icon="🎒">
          <h1 className="text-2xl font-black text-gray-800 mb-2">{statusInfo.tenantName || 'クラス'} に<ruby>参加<rp>(</rp><rt className="text-[0.5em]">さんか</rt><rp>)</rp></ruby>する</h1>
          <p className="text-sm text-gray-500 font-medium leading-relaxed mb-6"><ruby>名前<rp>(</rp><rt className="text-[0.6em]">なまえ</rt><rp>)</rp></ruby>を<ruby>入力<rp>(</rp><rt className="text-[0.6em]">にゅうりょく</rt><rp>)</rp></ruby>して、<ruby>参加<rp>(</rp><rt className="text-[0.6em]">さんか</rt><rp>)</rp></ruby>ボタンをおしてね。</p>
          {statusInfo.joinOpen ? (
            <>
              <input type="text" value={joinName} onChange={e=>setJoinName(e.target.value)} placeholder="なまえ（例: やまだ たろう）" className="w-full border-2 border-orange-100 rounded-2xl p-4 text-base outline-none focus:border-orange-400 bg-white font-bold mb-4 transition-colors text-center"/>
              <button onClick={handleJoin} disabled={server.isLoading} className="w-full bg-gradient-to-r from-orange-500 to-amber-500 text-white font-black py-4 rounded-2xl shadow-lg shadow-orange-500/30 active:scale-95 transition-all text-lg disabled:opacity-50">✋ クラスに<ruby>参加<rp>(</rp><rt className="text-[0.5em]">さんか</rt><rp>)</rp></ruby>する</button>
            </>
          ) : (
            <p className="text-sm font-bold text-red-500 bg-red-50 px-4 py-3 rounded-xl">いまは参加の受付が止まっています。先生に確認してね。</p>
          )}
          {server.toast && <Toast message={server.toast.message} type={server.toast.type} onClose={() => server.setToast(null)} />}
        </CenterCard>
      );

      if (phase === 'pending') return (
        <CenterCard icon="⏳">
          <h1 className="text-2xl font-black text-gray-800 mb-3"><ruby>先生<rp>(</rp><rt className="text-[0.5em]">せんせい</rt><rp>)</rp></ruby>の<ruby>承認<rp>(</rp><rt className="text-[0.5em]">しょうにん</rt><rp>)</rp></ruby>をまっています</h1>
          <p className="text-sm text-gray-500 font-medium leading-relaxed mb-6"><ruby>先生<rp>(</rp><rt className="text-[0.6em]">せんせい</rt><rp>)</rp></ruby>がOKしたら<ruby>使<rp>(</rp><rt className="text-[0.6em]">つか</rt><rp>)</rp></ruby>えるようになるよ。<br/>しばらくしてから、もう<ruby>一度<rp>(</rp><rt className="text-[0.6em]">いちど</rt><rp>)</rp></ruby>ためしてね。</p>
          <button onClick={()=>{ setPhase('connecting'); checkStatus(); }} className="w-full bg-blue-600 hover:bg-blue-700 text-white font-black py-3.5 rounded-2xl shadow-lg active:scale-95 transition-all">もう一度たしかめる</button>
        </CenterCard>
      );

      return (
        <div className="h-screen w-full flex flex-col bg-gray-50 text-gray-800 font-sans selection:bg-orange-200">
          <header className="flex-none bg-white/80 backdrop-blur-md z-30 px-3 py-2.5 md:p-4 md:px-6 flex justify-between items-center gap-2 border-b border-orange-500 shadow-[0_2px_10px_rgba(0,0,0,0.02)]">
            <div className="flex items-center gap-2 md:gap-3 min-w-0">
              <div className="p-2 md:p-2.5 rounded-2xl text-white shadow-sm shrink-0 bg-gradient-to-br from-orange-400 to-orange-500"><Icon path={Icons.Book} className="w-5 h-5 md:w-6 md:h-6"/></div>
              <h1 className="text-lg md:text-2xl font-black tracking-tight truncate text-orange-700">ふりかえりジャーナル <span aria-hidden="true" className="hidden sm:inline-block w-px h-5 bg-gray-300 align-middle"></span> <span className="text-gray-500 text-sm md:text-lg hidden sm:inline">{data.tenantName}</span></h1>
            </div>
            <div className="flex items-center gap-1.5 md:gap-2 bg-gray-100/80 px-3 md:px-4 py-1.5 md:py-2 rounded-2xl font-bold text-xs md:text-sm text-gray-600 shrink-0 max-w-[45%]"><Icon path={Icons.User} className="w-4 h-4 text-gray-500 shrink-0" /> <span className="truncate">{data.user.name} さん</span></div>
          </header>
          <main className="flex-1 overflow-y-auto p-3 md:p-6 lg:p-8"><div className="max-w-[1400px] mx-auto h-full">
            <StudentApp data={data} server={server} refresh={sync} />
          </div></main>
          <footer className="flex-none w-full text-center text-gray-500 py-3 bg-white border-t border-gray-100 text-sm">
            <span>© 2026 ふりかえりジャーナル <a href="https://giga-school.com" target="_blank" rel="noopener noreferrer" className="tap-44 inline-block no-underline text-inherit hover:opacity-80 transition-opacity">GIGA山</a></span>
          </footer>
        </div>
      );
    };

    const App = () => {
      // 描画完了をPWAシェルへ通知（シェル外では無害）
      useEffect(() => { if (window.notifyShellReady) window.notifyShellReady(); }, []);
      return window.BOOT.mode === 'member' ? <MemberRoot /> : <OwnerRoot />;
    };

    const root = ReactDOM.createRoot(document.getElementById('root'));
    root.render(<App />);
