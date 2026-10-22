// 散らばった過去の見積を集めて、一枚の実績表に整え、似た案件を数える。
// 業種に依存するのは pack（費目の辞書・条件の読み方・似た案件の決め方）だけ＝差し替えれば他業種でも動く。
// 金額はAIに出させない。AIの仕事は「辞書で読めなかった品目名を、どの費目かに振り分ける」ことだけ（確定は人）。
(function () {
  const TODAY = new Date('2026-10-17');
  const N = s => String(s == null ? '' : s).normalize('NFKC').trim();
  const num = v => { if (typeof v === 'number') return v; const m = N(v).replace(/[¥,円\s]/g, ''); return m && !isNaN(+m) ? +m : null; };
  const ymd = d => d.toISOString().slice(0, 10);
  const daysBetween = (a, b) => (new Date(b) - new Date(a)) / 864e5;

  // ── 辞書で品目を振り分ける ──
  function compile(pack) {
    pack._re = pack.categories.map(c => ({ key: c.key, re: new RegExp(c.re.map(x => N(x).replace(/[.*+?()[\]{}|\\]/g, m => m === '^' ? m : '\\' + m)).join('|')) }));
    pack._label = Object.fromEntries(pack.categories.map(c => [c.key, c.label]));
    return pack;
  }
  function ruleCat(pack, desc) { const s = N(desc); for (const c of pack._re) if (c.re.test(s)) return c.key; return null; }

  // ── 置き場1：freee（見積書・請求書のAPIの返り値） ──
  function fromFreee(q) {
    return q.quotations.map(x => ({
      id: 'F' + x.id, src: 'freee', srcLabel: 'freee見積書 ' + x.quotation_number, date: x.issue_date, partner: x.partner_name,
      title: x.title, memo: x.memo || '', total: x.total_amount,
      lines: x.quotation_lines.map(l => ({ desc: l.description, qty: l.quantity, unit: l.unit || '', price: l.unit_price, amount: Math.round(l.quantity * l.unit_price) })),
    }));
  }

  // ── 置き場2：freee前のExcel（担当ごとに様式が違う） ──
  function parseDate(v) {
    if (v instanceof Date) return ymd(new Date(v.getTime() - v.getTimezoneOffset() * 6e4));
    if (typeof v === 'number') return v > 40000 && v < 50000 ? ymd(new Date(Date.UTC(1899, 11, 30) + v * 864e5)) : null;
    const s = N(v); let m;
    if ((m = s.match(/令和(\d+)年(\d+)月(\d+)日/))) return ymd(new Date(Date.UTC(2018 + +m[1], m[2] - 1, +m[3])));
    if ((m = s.match(/平成(\d+)年(\d+)月(\d+)日/))) return ymd(new Date(Date.UTC(1988 + +m[1], m[2] - 1, +m[3])));
    if ((m = s.match(/(20\d\d)[年\/-](\d{1,2})[月\/-](\d{1,2})/))) return ymd(new Date(Date.UTC(+m[1], m[2] - 1, +m[3])));
    return null;
  }
  function parseSheet(rows, fname, sheetName, multi) {
    let h = -1;
    for (let i = 0; i < rows.length; i++) { const r = (rows[i] || []).map(N); if (r.includes('金額') && r.some(c => ['品名', '項目', '内容'].includes(c))) { h = i; break; } }
    if (h < 0) return null;
    const hr = rows[h].map(N), col = k => hr.indexOf(k);
    const descCol = col('品名') >= 0 ? col('品名') : col('項目') >= 0 ? col('項目') : col('内容');
    const noteCol = col('項目') >= 0 ? col('内容') : -1;
    const qtyCol = col('数量'), unitCol = col('単位'), priceCol = col('単価'), amtCol = col('金額');
    let partner = '', date = null, title = '', memo = '', total = null, pRow = -1;
    rows.forEach((r, i) => (r || []).forEach(c => {
      const s = N(c);
      if (i < h && !partner && /(御中|様)$/.test(s)) { partner = s.replace(/\s*(御中|様)$/, ''); pRow = i; }
      if (i < h && !date) { const d = parseDate(c); if (d) date = d; }
      if (/^件名/.test(s)) title = s.replace(/^件名[：:\s]*/, '');
      if (/^備考/.test(s)) memo = s.replace(/^備考[：:\s]*/, '');
    }));
    if (!title && pRow >= 0) for (let i = pRow + 1; i < h && !title; i++) { const s = N((rows[i] || []).find(c => N(c))); if (s && !/見積|下記|アオバ|担当/.test(s) && !parseDate(s)) title = s; }
    const lines = [];
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i] || [], d = N(r[descCol]);
      if (r.some(c => /^(小計|合計)/.test(N(c)))) { total = num(r[amtCol]); break; }
      if (!d) continue;
      const amount = num(r[amtCol]); if (amount == null) continue;
      let qty = qtyCol >= 0 ? num(r[qtyCol]) : null, unit = unitCol >= 0 ? N(r[unitCol]) : '';
      if (qty == null && noteCol >= 0) { const m = N(r[noteCol]).match(/^([\d.]+)(\D+)$/); if (m) { qty = +m[1]; unit = m[2]; } }
      if (qty == null) qty = 1;
      const price = priceCol >= 0 && num(r[priceCol]) != null ? num(r[priceCol]) : Math.round(amount / qty);
      lines.push({ desc: d, qty, unit, price, amount });
    }
    if (!lines.length) return null;
    return {
      id: 'X' + fname + '#' + sheetName, src: 'excel', srcLabel: 'Excel ' + fname + (multi ? '（' + sheetName + '）' : ''), file: fname, sheet: sheetName, multi,
      date, partner, title: title || fname.replace(/\.xlsx$/, ''), memo, total: total != null ? total : lines.reduce((t, l) => t + l.amount, 0), lines,
    };
  }
  // xlsxの中身（zipの中のXML）を直接読む。値と文字だけ取れればよい
  const dec = new TextDecoder();
  const unxml = t => t.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, d) => String.fromCharCode(+d)).replace(/&amp;/g, '&');
  function colIdx(ref) { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; }
  function readXlsx(u8) {
    const z = fflate.unzipSync(u8), txt = n => z[n] ? dec.decode(z[n]) : '';
    const ss = [...txt('xl/sharedStrings.xml').matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m => unxml([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(x => x[1]).join('')));
    const attr = (tag, a) => (tag.match(new RegExp('\\s' + a + '="([^"]*)"')) || [])[1];
    const rels = Object.fromEntries([...txt('xl/_rels/workbook.xml.rels').matchAll(/<Relationship\b[^>]*>/g)].map(m => [attr(m[0], 'Id'), (attr(m[0], 'Target') || '').replace(/^\/?(xl\/)?/, 'xl/')]));
    const sheets = [...txt('xl/workbook.xml').matchAll(/<sheet\b[^>]*>/g)].map(m => ({ name: unxml(attr(m[0], 'name') || ''), path: rels[attr(m[0], 'r:id')] }));
    return sheets.map(sh => {
      const rows = [];
      for (const m of txt(sh.path).matchAll(/<c r="([A-Z]+)(\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const t = (m[3].match(/t="(\w+)"/) || [])[1], body = m[4] || '';
        const v = (body.match(/<v>([\s\S]*?)<\/v>/) || [])[1], is = (body.match(/<t[^>]*>([\s\S]*?)<\/t>/) || [])[1];
        let val = t === 's' ? ss[+v] : t === 'inlineStr' ? unxml(is || '') : t === 'str' ? unxml(v || '') : v != null ? +v : '';
        const r = +m[2] - 1, c = colIdx(m[1]);
        (rows[r] = rows[r] || [])[c] = val;
      }
      for (let i = 0; i < rows.length; i++) rows[i] = Array.from(rows[i] || [], x => x === undefined ? '' : x);
      return { name: sh.name, rows };
    });
  }
  async function fromExcelZip(buf, onFile) {
    const files = fflate.unzipSync(new Uint8Array(buf)), out = [], names = Object.keys(files).filter(n => /\.xlsx$/i.test(n)).sort();
    let k = 0;
    for (const n of names) {
      const shs = readXlsx(files[n]), fname = n.split('/').pop(), multi = shs.length > 1;
      shs.forEach(sh => { const c = parseSheet(sh.rows, fname, sh.name, multi); if (c) out.push(c); });
      if (++k % 40 === 0) { if (onFile) onFile(k, names.length); await new Promise(r => setTimeout(r, 0)); }
    }
    return { cases: out, files: names.length };
  }

  // ── 置き場3：外注の支払い（freeeの支出をCSVで） ──
  function parseCsv(t) {
    return t.replace(/^﻿/, '').trim().split(/\r?\n/).slice(1).map(l => { const c = l.split(','); return { date: c[0], vendor: c[1], kamoku: c[2], amount: +c[3], memo: c.slice(4).join(',') }; });
  }

  // ── 整える：各案件の条件を品目から読む ──
  function secOf(s) { s = N(s); let m; if ((m = s.match(/(\d+)分(\d+)秒/))) return +m[1] * 60 + +m[2]; if ((m = s.match(/(\d+)分/))) return +m[1] * 60; if ((m = s.match(/(\d+)秒/))) return +m[1]; return null; }
  function derive(c) {
    const has = k => c.lines.some(l => l.cat === k), L = k => c.lines.filter(l => l.cat === k), txt = k => L(k).map(l => N(l.desc)).join(' ');
    const d = {};
    d.jissha = has('satsuei'); d.anime = has('anime'); d.cg3d = has('cg3d'); d.cad = /CAD/.test(txt('cg3d'));
    d.days = null; d.crew = null;
    for (const l of L('satsuei')) {
      const s = N(l.desc); let m;
      if (/半日/.test(s)) d.days = (d.days || 0) + 0.5;
      else if ((m = s.match(/(\d+(?:\.\d+)?)\s*(日|day)/i))) d.days = (d.days || 0) + +m[1];
      else if (l.unit === '日') d.days = (d.days || 0) + l.qty;
      if ((m = s.match(/(\d+)\s*名/) || s.match(/x\s*(\d+)/i))) d.crew = +m[1];
    }
    d.places = d.jissha ? 1 + L('ido').reduce((t, l) => t + (l.unit === '箇所' ? l.qty : 1), 0) : 0;
    d.cast = has('cast') ? 'pro' : 'nashi'; d.cast_n = L('cast').reduce((t, l) => t + (l.unit === '名' ? l.qty : 1), 0);
    d.narration = has('narration') ? 'ari' : 'nashi';
    const b = txt('bgm'); d.bgm = !b ? null : /か月|ヶ月|期間/.test(b) ? 'kigen' : /買取|ライセンス\)?$|ライセンス）/.test(b) ? 'mukigen' : null;
    d.shaku_sec = null;
    for (const s of [txt('henshu'), txt('anime'), N(c.title)]) { const v = secOf(s); if (v) { d.shaku_sec = v; break; } }
    let hn = null; const m = txt('henshu').match(/×\s*(\d+)\s*本/); if (m) hn = +m[1];
    if (!hn) for (const l of [...L('jimaku'), ...L('resize')]) if (l.unit === '本') hn = l.qty;
    d.honsu = hn || 1;
    d.tokkyu = has('tokkyu');
    c.cond = d;
    c.sub = c.lines.filter(l => !['tokkyu', 'shokeihi', 'nebiki'].includes(l.cat)).reduce((t, l) => t + l.amount, 0);
  }

  // ── 突き合わせ：受注したか（請求書・ファイル名・外注の支払い）と、持ち出し ──
  function tokens(c) {
    const p = N(c.partner).replace(/株式会社/g, ''), t = N(c.title), base = t.split('(')[0].split('（')[0];
    return [...new Set([p, p.slice(0, 4) + '様分', base.length >= 4 ? t.slice(0, 6) : ''].filter(x => x && x.length >= 3))];
  }
  function link(cases, invoices, pays) {
    const byPartner = {};
    invoices.forEach(v => (byPartner[v.partner_name] = byPartner[v.partner_name] || []).push(v));
    for (const c of cases) {
      c.pays = []; c.mochidashi = false;
      if (c.src === 'freee') {
        const iv = (byPartner[c.partner] || []).filter(v => N(v.title).startsWith(N(c.title)) && v.issue_date >= c.date);
        if (iv.length) { c.status = '受注'; c.statusWhy = 'freeeの請求書 ' + iv.map(v => v.invoice_number).join('・'); }
      } else if (/【受注】/.test(c.file) && !c.multi) { c.status = '受注'; c.statusWhy = 'ファイル名に【受注】'; }
      else if (/【受注】/.test(c.file) && c.multi) { c.status = '不明'; c.statusWhy = 'ファイル名に【受注】。ただしA案・B案のどちらで決まったかは書かれていない'; }
    }
    let ambiguous = 0;
    for (const p of pays) {
      const memo = N(p.memo);
      const cand = cases.filter(c => c.date && c.date <= p.date && daysBetween(c.date, p.date) <= 400 && tokens(c).some(t => memo.includes(t)));
      if (!cand.length) { p.case = null; continue; }
      if (cand.length > 1) ambiguous++;
      cand.sort((a, b) => daysBetween(a.date, p.date) - daysBetween(b.date, p.date));
      const best = cand.find(c => c.status === '受注') || cand[0];
      p.case = best.id; p.sure = cand.length === 1; best.pays.push(p);
      if (/掲載延長|追加ライセンス/.test(memo)) best.mochidashi = true;
    }
    for (const c of cases) {
      if (c.status) continue;
      if (c.pays.length) { c.status = '受注（推定）'; c.statusWhy = '外注の支払いが紐づく（' + c.pays.length + '件）'; }
      else if (c.src === 'freee') { c.status = daysBetween(c.date, TODAY) > 90 ? '失注' : '検討中'; c.statusWhy = c.status === '失注' ? '請求書なし・90日以上経過' : '請求書なし・90日以内'; }
      else { c.status = '不明'; c.statusWhy = 'freee前のExcelで、請求の記録も外注の支払いもない'; }
    }
    return { ambiguous };
  }

  // ── AIに回す前に、取引先の名前を伏せる ──
  function masker(cases) {
    const names = new Set();
    cases.forEach(c => { const p = N(c.partner); if (p) { names.add(p); const s = p.replace(/株式会社/g, ''); if (s.length >= 3) { names.add(s); names.add(s.slice(0, 6)); } } });
    const list = [...names].filter(n => n.length >= 3).sort((a, b) => b.length - a.length);
    return s => { let t = N(s), hit = false; for (const n of list) if (t.includes(n)) { t = t.split(n).join('［取引先］'); hit = true; } return { text: t, hit }; };
  }

  async function ingest(base, onStep) {
    const step = (k, msg) => onStep && onStep(k, msg);
    const get = async (p, kind) => { const r = await fetch(base + p); if (!r.ok) throw new Error(p + ' ' + r.status); return kind === 'buf' ? r.arrayBuffer() : kind === 'text' ? r.text() : r.json(); };
    const pack = compile(await get('pack_eizo.json'));
    step('pack', '業種の定義を読みました：' + pack.industry + '（費目' + pack.categories.length + '種）');
    const fq = await get('freee_quotations.json'), fi = await get('freee_invoices.json');
    const freee = fromFreee(fq);
    step('freee', 'freee：見積書 ' + freee.length + '件・請求書 ' + fi.invoices.length + '件');
    const ex = await fromExcelZip(await get('kyu_mitsumori_2019-2021.zip', 'buf'), (k, n) => step('excel-p', '共有フォルダ「旧見積」：' + k + '／' + n + 'ファイル'));
    step('excel', '共有フォルダ「旧見積」：' + ex.files + 'ファイル → 見積 ' + ex.cases.length + '件（1ファイルにA案・B案があるものは別々に）');
    const pays = parseCsv(await get('gaichu_shiharai.csv', 'text'));
    step('pays', 'freeeの支出（外注費）：' + pays.length + '件');
    const cases = [...ex.cases, ...freee].sort((a, b) => (a.date || '').localeCompare(b.date || ''));
    let lines = 0, byRule = 0;
    cases.forEach(c => c.lines.forEach(l => { lines++; l.cat = ruleCat(pack, l.desc); l.how = l.cat ? 'rule' : null; if (l.cat) byRule++; }));
    const mask = masker(cases);
    const odd = {};
    cases.forEach(c => c.lines.forEach(l => { if (!l.cat) { const m = mask(l.desc); (odd[m.text] = odd[m.text] || { text: m.text, masked: m.hit, n: 0, ex: c.srcLabel }).n++; } }));
    cases.forEach(derive);
    const lk = link(cases, fi.invoices, pays);
    step('done', '整えました');
    return { pack, cases, pays, invoices: fi.invoices, stats: { lines, byRule, odd: Object.values(odd).sort((a, b) => b.n - a.n), ambiguous: lk.ambiguous, excelFiles: ex.files, freeeN: freee.length, excelN: ex.cases.length }, mask };
  }

  // AI（または人）が決めた費目を、同じ書き方の全行に当てて、条件を読み直す
  function applyCats(db, map, how) {
    db.cases.forEach(c => { let ch = false; c.lines.forEach(l => { if (l.how === 'rule') return; const k = db.mask(l.desc).text; if (map[k] !== undefined) { l.cat = map[k] || null; l.how = map[k] ? how[k] : null; ch = true; } }); if (ch) derive(c); });
  }

  // ── 数える：似た案件の実績から、費目ごとの幅を出す（決まった計算＝同じ条件なら同じ金額） ──
  const q = (a, p) => { if (!a.length) return null; const s = [...a].sort((x, y) => x - y), i = (s.length - 1) * p, lo = Math.floor(i), hi = Math.ceil(i); return s[lo] + (s[hi] - s[lo]) * (i - lo); };
  const r1k = v => Math.round(v / 1000) * 1000;
  function estimate(db, f) {
    const P = db.pack.similar, since = new Date(TODAY); since.setFullYear(since.getFullYear() - P.since_years);
    const pool = db.cases.filter(c => (c.status === '受注' || c.status === '受注（推定）') && c.date && new Date(c.date) >= since);
    const req = { jissha: !!f.jissha, anime: !!f.anime, cg3d: !!f.cg3d, shaku: f.shaku_sec > 0 ? f.shaku_sec : 0, honsu: f.honsu > 0 ? f.honsu : 1 };
    const notes = [];
    let lo = P.ratio.shaku_sec[0], hi = P.ratio.shaku_sec[1], useHonsu = true;
    const pick = () => pool.filter(c => P.same.every(k => !!c.cond[k] === req[k]) &&
      (!req.shaku || (c.cond.shaku_sec && c.cond.shaku_sec / req.shaku >= lo && c.cond.shaku_sec / req.shaku <= hi)) &&
      (!useHonsu || Math.abs(c.cond.honsu - req.honsu) <= P.near.honsu));
    let sim = pick();
    if (sim.length < P.min_n) { lo = 1 / 3; hi = 3; useHonsu = false; sim = pick(); notes.push('似た案件が少ないので、尺の幅を広げ、本数では絞っていません'); }
    if (!req.shaku) notes.push('尺が未確認なので、尺では絞っていません');
    const rows = [];
    const row = (cat, name, opt = {}) => {
      let src = sim.filter(c => c.lines.some(l => l.cat === cat)); let wide = false;
      if (opt.filter) { const s2 = src.filter(opt.filter); if (s2.length >= 3) src = s2; }
      if (src.length < 3) { src = pool.filter(c => c.lines.some(l => l.cat === cat) && (!opt.filter || opt.filter(c))); wide = true; }
      const vals = src.map(c => { let a = c.lines.filter(l => l.cat === cat).reduce((t, l) => t + l.amount, 0); if (opt.per) { const u = opt.per(c); if (!u) return null; a = a / u * opt.mult; } return { c, a }; }).filter(x => x && x.a > 0);
      if (vals.length < 3) { rows.push({ cat, name, none: true, n: vals.length, why: '実績が' + vals.length + '件しかない＝滝川さんに確認' }); return; }
      const a = vals.map(v => v.a);
      rows.push({ cat, name, lo: r1k(q(a, .25)), mid: r1k(q(a, .5)), hi: r1k(q(a, .75)), n: vals.length, wide, ev: vals.sort((x, y) => y.c.date.localeCompare(x.c.date)), note: opt.note || '' });
    };
    const share = cat => sim.length ? sim.filter(c => c.lines.some(l => l.cat === cat)).length / sim.length : 0;
    const L = db.pack._label;
    row('kikaku', L.kikaku);
    if (f.jissha) { const d = f.days > 0 ? f.days : 1; row('satsuei', L.satsuei + '（' + (d === 0.5 ? '半日' : d + '日') + '）', { per: c => c.cond.days, mult: d, filter: c => d === 0.5 ? c.cond.days === 0.5 : c.cond.days >= 1, note: (d === 0.5 ? '半日の撮影だけを数えた実績' : '1日以上の撮影だけを数え、1日あたりの実績×日数') + (f.days > 0 ? '' : '（日数が未確認＝1日で仮置き）') }); }
    if (f.jissha && f.places > 1) row('ido', L.ido + '（' + (f.places - 1) + '箇所）');
    if (f.cast === 'pro') row('cast', L.cast, { per: c => c.cond.cast_n, mult: 1, note: '1名あたり。掲載期間（契約期間）で変わる' });
    if (f.cast === 'mikakunin' && f.jissha) rows.push({ cat: 'cast', name: '出演', none: true, why: '社員かプロかが未確認＝金額に入れていません' });
    row('henshu', L.henshu);
    if (share('jimaku') >= .5) row('jimaku', L.jimaku, { note: '似た案件の' + Math.round(share('jimaku') * 10) + '割に入っている' });
    if (f.anime) row('anime', L.anime);
    if (f.cg3d) row('cg3d', L.cg3d + (f.cad ? '（CAD支給あり）' : '（CAD支給なし）'), { filter: c => c.cond.cad === !!f.cad });
    if (f.narration === 'ari') { row('narration', L.narration); row('ma', L.ma); }
    if (f.narration === 'mikakunin') rows.push({ cat: 'narration', name: 'ナレーション', none: true, why: '有無が未確認＝金額に入れていません' });
    if (f.bgm === 'mukigen' || f.bgm === 'kigen') row('bgm', L.bgm + (f.bgm === 'mukigen' ? '（無期限）' : '（期間限定）'), { filter: c => c.cond.bgm === f.bgm });
    else row('bgm', L.bgm + '（形は未確認）', { note: '掲載期間が未確認＝無期限と期間限定の両方から' });
    if (req.honsu > 1 && share('resize') >= .5) row('resize', L.resize);
    let sub = rows.filter(r => !r.none), tlo = sub.reduce((t, r) => t + r.lo, 0), tmid = sub.reduce((t, r) => t + r.mid, 0), thi = sub.reduce((t, r) => t + r.hi, 0);
    if (f.tokkyu === 'ari') {
      const rs = pool.filter(c => c.cond.tokkyu && c.sub > 0).map(c => c.lines.filter(l => l.cat === 'tokkyu').reduce((t, l) => t + l.amount, 0) / c.sub);
      const k = rs.length >= 3 ? q(rs, .5) : 0.2;
      rows.push({ cat: 'tokkyu', name: L.tokkyu, lo: r1k(tlo * k), mid: r1k(tmid * k), hi: r1k(thi * k), n: rs.length, note: '過去' + rs.length + '件の特急の割合（中央値' + Math.round(k * 100) + '%）', ev: [] });
      tlo += r1k(tlo * k); tmid += r1k(tmid * k); thi += r1k(thi * k);
    }
    return { rows, sim, pool: pool.length, notes, total: { lo: tlo, mid: tmid, hi: thi } };
  }

  // ── 押さえどころ：いちばん似た過去の案件と、どこが違うか ──
  // 見積が外れるのは、前回の似た案件と「一つだけ条件が違う」とき。その違いを先に並べる（決まった計算）
  function compare(db, f) {
    const since = new Date(TODAY); since.setFullYear(since.getFullYear() - db.pack.similar.since_years);
    const pool = db.cases.filter(c => (c.status === '受注' || c.status === '受注（推定）') && c.date && new Date(c.date) >= since);
    const pro = x => x === 'pro', known = v => v !== undefined && v !== null && v !== 'mikakunin';
    const dist = c => {
      const d = c.cond; let x = 0;
      for (const k of ['jissha', 'anime', 'cg3d']) if (!!d[k] !== !!f[k]) x += 10;
      if (f.shaku_sec > 0) x += d.shaku_sec ? Math.abs(Math.log2(d.shaku_sec / f.shaku_sec)) * 3 : 1;
      x += Math.abs((d.honsu || 1) - (f.honsu > 0 ? f.honsu : 1)) * 1.5;
      if (f.jissha && f.days > 0) x += d.days ? Math.abs(d.days - f.days) * 2 : 1;
      if (f.jissha && f.places > 0) x += Math.abs((d.places || 1) - f.places);
      if (known(f.cast)) x += pro(f.cast) !== pro(d.cast) ? 3 : 0;
      if (known(f.narration)) x += (f.narration === 'ari') !== (d.narration === 'ari') ? 2 : 0;
      if (known(f.bgm)) x += d.bgm && d.bgm !== f.bgm ? 3 : d.bgm ? 0 : 1;
      if (f.cg3d) x += !!d.cad !== !!f.cad ? 3 : 0;
      if (known(f.tokkyu)) x += (f.tokkyu === 'ari') !== !!d.tokkyu ? 1 : 0;
      if (f.kw && f.kw.length) x += f.kw.some(k => N(c.title).includes(k)) ? -4 : 3;  // 問い合わせの言葉（会社紹介・採用…）と件名が合うものを優先
      return x;
    };
    const ranked = pool.map(c => ({ c, x: dist(c) })).sort((a, b) => a.x - b.x || b.c.date.localeCompare(a.c.date));
    if (!ranked.length) return null;
    const c = ranked[0].c, d = c.cond, diffs = [], todo = [];
    const sec = s => s >= 60 ? Math.floor(s / 60) + '分' + (s % 60 ? s % 60 + '秒' : '') : s + '秒';
    const hy = o => [o.jissha ? '実写' : '', o.anime ? 'アニメ' : '', o.cg3d ? '3DCG' : ''].filter(Boolean).join('＋') || '—';
    const add = (label, now, before, why) => diffs.push({ label, now, before, why });
    if (hy(f) !== hy(d)) add('表現', hy(f), hy(d), '撮影費・制作費の組み立てが変わる');
    if (f.shaku_sec > 0 && d.shaku_sec && (f.shaku_sec / d.shaku_sec > 1.25 || f.shaku_sec / d.shaku_sec < 0.8)) add('尺', sec(f.shaku_sec), sec(d.shaku_sec), '編集の工数が変わる');
    if ((f.honsu > 0 ? f.honsu : 1) !== (d.honsu || 1)) add('本数', (f.honsu || 1) + '本', (d.honsu || 1) + '本', '企画・編集・サイズ展開が増える');
    if (f.jissha && f.days > 0 && d.days && f.days !== d.days) add('撮影日数', f.days + '日', d.days + '日', '撮影費は日数で決まる');
    if (f.jissha && f.places > 0 && d.places && f.places !== d.places) add('撮影場所', f.places + '箇所', d.places + '箇所', '移動と追加ロケの費用');
    if (known(f.cast) && pro(f.cast) !== pro(d.cast)) add('出演', pro(f.cast) ? 'プロ' : '社員・なし', pro(d.cast) ? 'プロ' : '社員・なし', '出演料と、掲載期間に合わせた契約が要る');
    if (known(f.narration) && (f.narration === 'ari') !== (d.narration === 'ari')) add('ナレーション', f.narration === 'ari' ? 'あり' : 'なし', d.narration === 'ari' ? 'あり' : 'なし', '収録とMAの費用');
    const mochi = db.cases.filter(x => x.mochidashi).length;
    const bgmWhy = 'BGMのライセンスの形が変わる' + (mochi ? '。過去' + mochi + '件は、後から掲載を延ばされて追加ライセンス料を払っている（持ち出し）' : '');
    const bl = b => b === 'mukigen' ? '無期限' : b === 'kigen' ? '期間限定' : '不明';
    if (known(f.bgm) && d.bgm && f.bgm !== d.bgm) add('掲載期間（BGM）', bl(f.bgm), bl(d.bgm), bgmWhy);
    if (f.cg3d && !!f.cad !== !!d.cad) add('CADデータ', f.cad ? '支給あり' : 'なし', d.cad ? '支給あり' : 'なし', 'CADが無いとモデリングから＝3DCGの工数が大きく変わる');
    if (known(f.tokkyu) && (f.tokkyu === 'ari') !== !!d.tokkyu) add('納期', f.tokkyu === 'ari' ? '特急' : '通常', d.tokkyu ? '特急' : '通常', '特急の割増が乗る');
    if (!known(f.bgm)) todo.push({ label: '掲載期間（BGM）', why: bgmWhy });
    if (f.jissha && !known(f.cast)) todo.push({ label: '出演（社員かプロか）', why: '出演料と契約期間が変わる' });
    if (!known(f.narration)) todo.push({ label: 'ナレーション', why: '収録とMAの費用' });
    if (f.jissha && !(f.days > 0)) todo.push({ label: '撮影日数', why: '撮影費は日数で決まる' });
    if (!(f.shaku_sec > 0)) todo.push({ label: '尺', why: '編集の工数が変わる' });
    if (!known(f.tokkyu)) todo.push({ label: '納期', why: '特急なら割増が乗る' });
    return { c, diffs, todo, n: pool.length, near: ranked.slice(1, 4).map(r => r.c) };
  }

  window.SEIRI = { ingest, applyCats, estimate, compare, N };
})();
