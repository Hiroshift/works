// 受注管理（架空データの見本・板金工場）。サーバーは使わず、ブラウザの中だけで動く。
// 流れは社内で動いている本番と同じ：営業ごとのExcel → 案件一覧 → 担当と図面確定日で絞る → チェック → 並べる → 帳票3種。
// 画面で決めたこと（機械・状況・注意事項）は、この端末のブラウザにだけ残す（営業のExcelは書き換えない）。
(function () {
  "use strict";
  const D = window.DEMO_DATA;
  // 機械の顔ぶれ（切断・抜きの機械。曲げ・溶接・表面処理は後工程で、ここでは割り当てない）
  const MACHINES = {
    "レーザーA": "ファイバー 4kW・鉄 〜6mm",
    "レーザーB": "ファイバー 6kW・厚板 〜12mm・ステンレス/アルミ",
    "タレパン": "穴の多い薄板 〜3.2mm",
    "複合機": "パンチ＋レーザー・成形穴（タップ/バーリング）",
  };
  const MACHINE_ORDER = Object.keys(MACHINES);
  // 図面/データの状況（営業が画面で選ぶ。Excelには無い）
  const STATUS = ["", "図面待ち", "図面受領", "図面確定", "加工データ済"];
  const STATUS_CLS = { "図面待ち": "s1", "図面受領": "s2", "図面確定": "s3", "加工データ済": "s4" };
  const COLORS = ["#2563eb", "#16a34a", "#dc2626", "#7c3aed", "#ea580c"];
  const STORE_KEY = "juchu_demo_bankin_v1";

  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => [...el.querySelectorAll(s)];
  const esc = (v) => String(v ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  // ── 日付（データは「今日から何営業日目か」で持っている） ──
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const isWork = (d) => d.getDay() >= 1 && d.getDay() <= 5;
  function addWd(n) {
    const d = new Date(today);
    while (n > 0) { d.setDate(d.getDate() + 1); if (isWork(d)) n--; }
    while (n < 0) { d.setDate(d.getDate() - 1); if (isWork(d)) n++; }
    return d;
  }
  const pad = (n) => String(n).padStart(2, "0");
  const iso = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const md = (d) => (d ? `${d.getMonth() + 1}/${d.getDate()}` : "");
  const wdName = (d) => "日月火水木金土"[d.getDay()];
  const stamp = (d) => `${md(d)} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
  const nextWd = addWd(1);

  // ── 保存（この端末だけ。使えないブラウザでも画面は動く） ──
  let S = null;
  function save() { try { localStorage.setItem(STORE_KEY, JSON.stringify(S)); } catch (e) { /* 残せなくても動かす */ } }
  function load() { try { const v = JSON.parse(localStorage.getItem(STORE_KEY) || "null"); return v && v.day === iso(today) ? v : null; } catch (e) { return null; } }
  function initState() { S = { day: iso(today), ui: {} }; save(); }

  // ── 受注（営業のExcelを読んだ結果にあたる） ──
  function materialize(o) {
    const fix = addWd(o.fix), nouki = addWd(o.nouki);
    return Object.assign({}, o, {
      fixDate: fix, fixIso: iso(fix), noukiDate: nouki, ukeDate: o.uke === null ? null : addWd(o.uke),
      kakouDate: o.kakouN === null || o.kakouN === undefined ? null : addWd(o.kakouN),
    });
  }
  const ORDERS = D.orders.map(materialize);
  const byNo = new Map(ORDERS.map((o) => [o.no, o]));
  const people = D.people.map((p) => p.name);
  const colorOf = (name) => { const i = people.indexOf(name); return i < 0 ? "#6b7280" : COLORS[i % COLORS.length]; };

  // 画面で決めたこと（台帳）。機械＝台帳の確定/希望があればそれ、無ければExcelの機械列。
  const machineOf = (o) => { const v = (S.ui[o.no] || {}).machine; return v === undefined ? (o.machine || "") : (v === "-" ? "" : v); };
  const confirmed = (o) => (S.ui[o.no] || {}).machine !== undefined;
  const statusOf = (o) => (S.ui[o.no] || {}).status || "";
  const noteOf = (o) => (S.ui[o.no] || {}).note || "";
  function setUi(no, field, v) {
    const h = S.ui[no] || (S.ui[no] = {});
    if (v === "" && field !== "machine") delete h[field]; else h[field] = v;
    if (!Object.keys(h).length) delete S.ui[no];
    save();
  }
  // 紙に出す納期＝加工納期を優先し、空なら納期（本番と同じ決まり）
  const paperNouki = (o) => md(o.kakouDate || o.noukiDate);

  function toast(msg) { const t = $("#toast"); t.textContent = msg; t.classList.add("on"); clearTimeout(toast.h); toast.h = setTimeout(() => t.classList.remove("on"), 1800); }

  // ── 画面の切り替え ──
  const route = () => (location.hash.replace("#", "") || "list");
  const VIEWS = { list: viewList, assign: viewAssign };
  let picked = null;
  function render() {
    const r = VIEWS[route()] ? route() : "list";
    $$("#nav a").forEach((a) => a.classList.toggle("on", a.dataset.r === r));
    const m = $("#main"); m.onclick = m.ondragstart = m.ondragend = null; picked = null;
    VIEWS[r](m);
  }

  // ── 案件一覧 ──
  const F = { eigyo: "すべて", day: "", month: false };
  let listOrder = null; // ドラッグ／自動整列で並べたときの並び（画面を開き直しても保つ）
  function viewList(el) {
    const counts = {}; ORDERS.forEach((o) => { counts[o.eigyo] = (counts[o.eigyo] || 0) + 1; });
    const saved = new Date(today); saved.setHours(8, 40);
    const chips = D.people.map((p) => `<span class="fchip"><b style="background:${colorOf(p.name)}"></b>${esc(p.file)}<em class="num">${counts[p.name] || 0}件・保存 ${stamp(saved)}${p.header !== "数量" ? `・見出し「${esc(p.header)}」でも読める` : ""}</em></span>`).join("");
    el.innerHTML = `
      <h1>案件一覧 ― 伝票を出したい案件をえらぶ</h1>
      <p class="lead">営業${people.length}人が、それぞれ自分のExcelに1行1案件で書く。この画面は開くたびに受注フォルダの${D.people.length}本を全部読んで、1つの一覧にまとめます（合わせて <b class="num">${ORDERS.length}</b> 件）。</p>
      <div class="keep box"><b>変えないもの</b><ul>
        <li>営業は今までどおり、自分のExcelに書くだけ（書き方も置き場所も同じ）</li>
        <li>営業のPCには何も入れない。ブラウザで開くだけ</li>
        <li>アプリはExcelを読むだけで、書き換えない（正本はExcelのまま）</li>
        <li>見出しの書き方が人によってずれていても、そのまま読む（例：深沢さんの「数量（個）」）</li>
      </ul></div>
      <div class="files">${chips}</div>
      <div class="filter box">
        <label>担当営業<select id="fE">${["すべて"].concat(people).map((x) => `<option${x === F.eigyo ? " selected" : ""}>${esc(x)}</option>`).join("")}</select></label>
        <label>図面確定日<input type="date" id="fD" value="${F.day}"></label>
        <label class="ck"><input type="checkbox" id="fM"${F.month ? " checked" : ""}>その月で絞る</label>
        <button type="button" class="btn go" id="fGo">絞り込む</button>
        <button type="button" class="btn sub" id="fAll">すべて表示</button>
        <button type="button" class="btn sub" id="fNext" title="図面確定日を翌営業日にして絞ります">翌営業日（${md(nextWd)}${wdName(nextWd)}）</button>
        <span class="count num" id="count"></span>
      </div>
      <div class="hint">↑ 担当営業＝自分・図面確定日でしぼり込み → 案件をチェック → 下のボタンで帳票。「機械予定入れ（紙）」はプリンターで出して、予定入れのクリアファイルへ。<br>
      ↑↓ 左端の <span class="h">⠿</span> をドラッグで行を並べ替え／「自動整列」で <b>材質（鉄→ステンレス→アルミ）→ 板厚 → 機械</b> の順にまとめます。並べた順がそのまま機械予定入れの並びになります。</div>
      <div class="tbl box"><table>
        <colgroup><col style="width:22px"><col style="width:30px"><col style="width:80px"><col style="width:116px"><col><col style="width:74px"><col style="width:44px"><col style="width:44px"><col style="width:112px"><col style="width:66px"><col style="width:108px"><col style="width:170px"><col style="width:58px"><col style="width:52px"></colgroup>
        <thead><tr>
          <th class="drag"></th>
          <th class="chk"><input type="checkbox" id="all" aria-label="表示中を全部選ぶ"></th>
          <th>発注番号</th><th>得意先</th><th>品名<small>図番</small></th>
          <th class="cen">材質</th><th class="cen">板厚</th><th class="cen">数量</th>
          <th class="cen">機械<small>希望</small></th>
          <th class="hot">図面確定</th>
          <th class="cen">図面/データ<small>状況</small></th>
          <th>注意事項<small>立ち合い/客要望</small></th>
          <th class="hot">納期</th><th>担当</th>
        </tr></thead>
        <tbody id="tb"></tbody>
      </table></div>
      <div class="bar">
        <button type="button" class="btn sub" id="arrange" title="材質→板厚→機械の順に並べ替え">⇅ 自動整列</button>
        <span class="muted hint" style="font-size:12.5px;margin:0">☑ チェックした案件に対して →</span>
        <button type="button" class="btn hot" data-sheet="wo">作業伝票を出す</button>
        <button type="button" class="btn blue" data-sheet="ct">切断伝票を出す</button>
        <button type="button" class="btn main" data-sheet="yi">機械予定入れ（紙）を出す</button>
        <span class="muted sel" id="sel"></span>
      </div>`;
    const tb = $("#tb");
    // 本番と同じく、Excelの下の行（新しい受付）から上に出す
    let list = ORDERS.slice().sort((a, b) => (b.uke - a.uke) || (b.no < a.no ? -1 : 1));
    if (listOrder) { const pos = new Map(listOrder.map((no, i) => [no, i])); list.sort((a, b) => (pos.get(a.no) ?? 1e9) - (pos.get(b.no) ?? 1e9)); }
    tb.innerHTML = list.map(rowHtml).join("");
    applyFilter();
    countSel();
    initDnD(tb);

    const readF = () => { F.eigyo = $("#fE").value; F.day = $("#fD").value; F.month = $("#fM").checked; };
    $("#fGo").onclick = () => { readF(); applyFilter(); };
    $("#fE").onchange = $("#fD").onchange = $("#fM").onchange = () => { readF(); applyFilter(); };
    $("#fNext").onclick = () => { F.day = iso(nextWd); F.month = false; $("#fD").value = F.day; $("#fM").checked = false; applyFilter(); };
    $("#fAll").onclick = () => { Object.assign(F, { eigyo: "すべて", day: "", month: false }); render(); };
    $("#all").onchange = (e) => { $$("#tb tr").forEach((tr) => { if (tr.style.display !== "none") $("input[type=checkbox]", tr).checked = e.target.checked; }); countSel(); };
    tb.addEventListener("change", (e) => {
      const t = e.target, no = t.dataset.no;
      if (t.name === "pick") return countSel();
      if (!no) return;
      const tr = t.closest("tr");
      if (t.classList.contains("gsel")) {
        setUi(no, "machine", t.value || "-"); t.className = "gsel " + (t.value ? "set" : "empty"); tr.dataset.machine = t.value;
        if (t.value) { $("input[name=pick]", tr).checked = true; countSel(); } // 機械を指定したら、その行を自動でチェック＝出し忘れを防ぐ
      }
      if (t.classList.contains("ssel")) { setUi(no, "status", t.value); t.className = "ssel " + (STATUS_CLS[t.value] || "none"); }
      if (t.classList.contains("nsel")) {
        const v = t.value.trim(); setUi(no, "note", v); t.className = "nsel " + (v ? "set" : "empty");
        if (v) { $("input[name=pick]", tr).checked = true; countSel(); } // 注意事項を打ったら、その行を自動でチェック
      }
      toast("保存しました（この端末のブラウザの中だけ）");
    });
    $("#arrange").onclick = () => { autoArrange(tb); listOrder = $$("#tb tr").map((r) => r.dataset.no); toast("材質 → 板厚 → 機械 の順に並べました"); };
    $$(".bar [data-sheet]").forEach((b) => b.onclick = () => {
      const nos = $$("#tb input[name=pick]:checked").map((c) => c.value); // チェックはDOM順＝画面で並べた順のまま
      const recs = nos.map((no) => byNo.get(no)).filter(Boolean);
      if (!recs.length) return toast("先に案件をチェックしてください");
      openSheet(b.dataset.sheet, recs);
    });
  }
  function rowHtml(o) {
    const mc = machineOf(o), st = statusOf(o), note = noteOf(o);
    const mcOpts = [""].concat(MACHINE_ORDER).map((g) => `<option value="${g}"${g === mc ? " selected" : ""}>${g || "—"}</option>`).join("");
    const stOpts = STATUS.map((s) => `<option value="${s}"${s === st ? " selected" : ""}>${s || "—"}</option>`).join("");
    return `<tr data-no="${esc(o.no)}" data-eigyo="${esc(o.eigyo)}" data-day="${o.fixIso}" data-zai="${esc(o.zai)}" data-ita="${esc(o.ita)}" data-machine="${esc(mc)}" class="${o.done ? "done" : ""}">
      <td class="drag" title="ドラッグで並べ替え">⠿</td>
      <td class="chk"><input type="checkbox" name="pick" value="${esc(o.no)}" aria-label="選ぶ"></td>
      <td class="num">${esc(o.no)}</td><td>${esc(o.tokui)}</td>
      <td class="hin">${esc(o.hinmei)}<span>${esc(o.zuban)}${o.done ? "・済" : ""}</span></td>
      <td class="cen">${esc(o.zai)}</td><td class="cen num">${esc(o.ita)}</td><td class="cen num">${esc(o.qty)}</td>
      <td><select class="gsel ${mc ? "set" : "empty"}" data-no="${esc(o.no)}" aria-label="機械（希望）">${mcOpts}</select></td>
      <td class="hot num">${md(o.fixDate)}</td>
      <td><select class="ssel ${STATUS_CLS[st] || "none"}" data-no="${esc(o.no)}" aria-label="図面/データの状況">${stOpts}</select></td>
      <td><input type="text" class="nsel ${note ? "set" : "empty"}" data-no="${esc(o.no)}" value="${esc(note)}" placeholder="立ち合い/客要望…" aria-label="注意事項"></td>
      <td class="hot num">${md(o.noukiDate)}</td>
      <td><span class="who" style="border-color:${colorOf(o.eigyo)}">${esc(o.eigyo)}</span></td></tr>`;
  }
  function applyFilter() {
    let n = 0;
    $$("#tb tr").forEach((tr) => {
      let ok = true; const d = tr.dataset.day;
      if (F.eigyo !== "すべて" && tr.dataset.eigyo !== F.eigyo) ok = false;
      if (F.day) { if (F.month) { if (d.slice(0, 7) !== F.day.slice(0, 7)) ok = false; } else if (d !== F.day) ok = false; }
      tr.style.display = ok ? "" : "none"; if (ok) n++;
    });
    $("#count").textContent = `表示中：${n} 件`;
  }
  function countSel() { const el = $("#sel"); if (el) el.textContent = `選択 ${$$("#tb input[name=pick]:checked").length} 件`; }

  // 自動整列：材質（鉄→ステンレス→アルミ）→板厚（薄→厚）→機械。同じ材質・板厚が隣に並ぶと機械の段取り替えが減る。
  const zaiRank = (z) => (/SUS/.test(z) ? 1 : /^A\d/.test(z) ? 2 : 0);
  const zaiGroup = (z) => ["鉄", "ステンレス", "アルミ"][zaiRank(z)];
  const mcRank = (m) => { const i = MACHINE_ORDER.indexOf(m); return i < 0 ? 99 : i; };
  function autoArrange(tb) {
    const all = $$("tr", tb), shown = all.filter((r) => r.style.display !== "none"), hidden = all.filter((r) => r.style.display === "none");
    shown.sort((a, b) => (zaiRank(a.dataset.zai) - zaiRank(b.dataset.zai)) || (parseFloat(a.dataset.ita) - parseFloat(b.dataset.ita)) || (mcRank(a.dataset.machine) - mcRank(b.dataset.machine)));
    shown.concat(hidden).forEach((r) => tb.appendChild(r));
    // 材質の境目に区切り線と札
    $$("tr.grp", tb).forEach((r) => { r.classList.remove("grp"); const t = $(".grp-tag", r); if (t) t.remove(); });
    let last = null;
    shown.forEach((r) => { const g = zaiGroup(r.dataset.zai); if (g !== last) { r.classList.add("grp"); $("td.cen", r).insertAdjacentHTML("beforeend", `<i class="grp-tag">${g}</i>`); last = g; } });
  }
  // ⠿ からだけドラッグ開始＝チェックや選択を邪魔しない
  function initDnD(tb) {
    let dragRow = null;
    $$("tr", tb).forEach((tr) => {
      const h = $(".drag", tr);
      h.addEventListener("mousedown", () => tr.setAttribute("draggable", "true"));
      h.addEventListener("mouseup", () => tr.removeAttribute("draggable"));
      tr.addEventListener("dragstart", (e) => { dragRow = tr; tr.classList.add("dragging"); e.dataTransfer.effectAllowed = "move"; });
      tr.addEventListener("dragend", () => { tr.classList.remove("dragging"); tr.removeAttribute("draggable"); dragRow = null; listOrder = $$("tr", tb).map((r) => r.dataset.no); });
      tr.addEventListener("dragover", (e) => {
        e.preventDefault(); if (!dragRow || tr === dragRow) return;
        const rc = tr.getBoundingClientRect(); const after = (e.clientY - rc.top) / rc.height > 0.5;
        tb.insertBefore(dragRow, after ? tr.nextSibling : tr);
      });
    });
  }

  // ── 機械わりあて盤（次段階・本番でも次段階の扱い） ──
  let assignDay = null, assignEigyo = "すべて";
  function viewAssign(el) {
    if (assignDay === null) assignDay = iso(nextWd);
    const recs = ORDERS.filter((o) => !o.done).sort((a, b) => a.fixIso.localeCompare(b.fixIso));
    const lanes = [["", "未割当", "ここから各機械へドラッグ"]].concat(MACHINE_ORDER.map((k) => [k, k, MACHINES[k]]));
    el.innerHTML = `
      <h1>機械わりあて盤 <span class="muted" style="font-size:12px;font-weight:400">― 次段階（本番では工場長が紙を切って台紙に貼る。その置き換え案）</span></h1>
      <p class="lead">カードを機械の列へドラッグ＝ハサミ＋糊の代わり。放した瞬間に確定が保存されます（営業のExcelは書き換えません）。スマホでは、カードを押してから列の見出しを押します。「未割当」を空にするのが目標。</p>
      <div class="filter box"><label>担当営業<select id="aE">${["すべて"].concat(people).map((x) => `<option${x === assignEigyo ? " selected" : ""}>${esc(x)}</option>`).join("")}</select></label>
        <label>図面確定日<input type="date" id="aD" value="${assignDay}"></label><button type="button" class="btn sub" id="aClear">すべて表示</button>
        <button type="button" class="btn main" id="aPdf" style="margin-left:auto">表示中の案件で 機械予定入れ（紙）を出す</button></div>
      <div class="board">${lanes.map(([g, t, sub]) => `<div class="lane${g ? "" : " un"}" data-g="${esc(g)}"><div class="lh"><b>${esc(t)}</b><span class="lc"></span></div><div class="ls">${esc(sub)}</div><div class="cards">${
        recs.filter((o) => machineOf(o) === g).map((o) => {
          const badge = confirmed(o) ? '<i class="b conf">確定</i>' : (o.machine ? '<i class="b kibou">営業希望</i>' : '<i class="b none">希望なし</i>');
          return `<div class="cj" draggable="true" data-no="${esc(o.no)}" data-eigyo="${esc(o.eigyo)}" data-day="${o.fixIso}" style="border-left-color:${colorOf(o.eigyo)}"><div class="ct"><span class="num">${esc(o.no)}</span>${badge}</div><div class="cc">${esc(o.tokui)}</div><div class="ch">${esc(o.hinmei)}</div><div class="cm num">図面確定 ${md(o.fixDate)}　納 ${md(o.noukiDate)}　${esc(o.qty)}個　${esc(o.zai)} t${esc(o.ita)}　${esc(o.eigyo)}</div></div>`;
        }).join("")}</div></div>`).join("")}</div>`;
    const apply = () => {
      $$(".cj").forEach((c) => { let ok = true; if (assignEigyo !== "すべて" && c.dataset.eigyo !== assignEigyo) ok = false; if (assignDay && c.dataset.day !== assignDay) ok = false; c.style.display = ok ? "" : "none"; });
      $$(".lane").forEach((l) => { $(".lc", l).textContent = $$(".cj", l).filter((c) => c.style.display !== "none").length; });
    };
    apply();
    $("#aE").onchange = (e) => { assignEigyo = e.target.value; apply(); };
    $("#aD").onchange = (e) => { assignDay = e.target.value; apply(); };
    $("#aClear").onclick = () => { assignDay = ""; assignEigyo = "すべて"; $("#aD").value = ""; $("#aE").value = "すべて"; apply(); };
    $("#aPdf").onclick = () => {
      const recs2 = $$(".cj").filter((c) => c.style.display !== "none").map((c) => byNo.get(c.dataset.no)).filter(Boolean);
      if (!recs2.length) return toast("表示中の案件がありません");
      openSheet("yi", recs2);
    };
    const moveTo = (card, lane) => {
      const g = lane.dataset.g;
      $(".cards", lane).prepend(card);
      setUi(card.dataset.no, "machine", g || "-");
      const b = $(".b", card); b.className = "b " + (g ? "conf" : "none"); b.textContent = g ? "確定" : "希望なし";
      card.classList.remove("picked"); picked = null; $$(".lane").forEach((l) => l.classList.remove("target"));
      apply(); toast(`${g || "未割当"}へ移しました`);
    };
    let drag = null;
    el.ondragstart = (e) => { const c = e.target.closest && e.target.closest(".cj"); if (c) { drag = c; c.classList.add("dg"); } };
    el.ondragend = () => { if (drag) drag.classList.remove("dg"); drag = null; $$(".over").forEach((x) => x.classList.remove("over")); };
    $$(".cards").forEach((z) => {
      z.addEventListener("dragover", (e) => { e.preventDefault(); z.classList.add("over"); });
      z.addEventListener("dragleave", () => z.classList.remove("over"));
      z.addEventListener("drop", (e) => { e.preventDefault(); z.classList.remove("over"); if (drag) moveTo(drag, z.closest(".lane")); });
    });
    // 押して選ぶ → 列を押して移す（スマホ向け）
    el.onclick = (e) => {
      const card = e.target.closest(".cj");
      if (card) { if (picked === card) { card.classList.remove("picked"); picked = null; $$(".lane").forEach((l) => l.classList.remove("target")); return; }
        if (picked) picked.classList.remove("picked"); picked = card; card.classList.add("picked"); $$(".lane").forEach((l) => l.classList.toggle("target", l !== card.closest(".lane"))); return; }
      const lane = e.target.closest(".lane");
      if (lane && picked && lane !== picked.closest(".lane")) moveTo(picked, lane);
    };
  }

  // ── 帳票（PDFの代わりに、画面で見てブラウザからプリンターへ出す） ──
  const SHEET_NAME = { wo: "作業伝票", ct: "切断伝票", yi: "機械予定入れ" };
  const outStamp = () => stamp(new Date());
  const or = (v, alt) => (String(v || "").trim() ? esc(v) : `<span style="color:#9ca3af;font-weight:400">${alt || ""}</span>`);
  const hinmeiFull = (o) => `${esc(o.tokui)}　${esc(o.hinmei)}${o.zuban ? `　<span style="font-size:.8em;color:#4b5563">${esc(o.zuban)}</span>` : ""}`;

  // 作業伝票（A4縦・1案件1枚）。工程の順＝共通 → 図面・データ → 切断・抜き → 曲げ・仕上げ。渡された部署が読む欄を固める。
  function pageWo(o) {
    const note = noteOf(o), mc = machineOf(o);
    const tokki = (label, v) => `<tr class="${v ? "tokki" : ""} big"><td class="k">${label}</td><td colspan="3">${esc(v)}</td></tr>`;
    return `<div class="page"><span class="pg-tag">架空データの見本</span>
      <div class="wo-title">作業伝票</div>
      <table class="wo">
        <tr><td class="k">発注番号</td><td class="num">${esc(o.no)}</td><td class="k">図面確定</td><td class="num">${md(o.fixDate)}</td></tr>
        <tr><td class="k">品　名</td><td colspan="3">${hinmeiFull(o)}</td></tr>
        <tr class="${note ? "tokki" : ""}"><td class="k">注意事項</td><td colspan="3">${esc(note)}</td></tr>
        <tr><td class="k">担　当</td><td>${esc(o.eigyo)}</td><td class="k">納　期</td><td class="num">${md(o.noukiDate)}</td></tr>
        <tr class="blk"><td colspan="4">■ 図面・データ</td></tr>
        <tr><td class="k">材　質</td><td>${esc(o.zai)}</td><td class="k">板　厚</td><td class="num">${esc(o.ita)} mm</td></tr>
        <tr><td class="k">展開寸法</td><td colspan="3" class="num">${esc(o.tenkai)}${o.atsuen ? `　<span style="font-size:.85em">圧延方向：${esc(o.atsuen)}</span>` : ""}</td></tr>
        <tr><td class="k">数　量</td><td class="num">${esc(o.qty)} 個</td><td class="k">材料サイズ</td><td class="num">${esc(o.zairyo)}</td></tr>
        <tr class="fig"><td class="k">展開図</td><td colspan="3" style="color:#9ca3af;font-size:11px">（手で描く・板取りの指示）</td></tr>
        ${tokki("データ特記", o.tdata)}
        <tr class="blk"><td colspan="4">■ 切断・抜き</td></tr>
        <tr><td class="k">材　質</td><td>${esc(o.zai)}</td><td class="k">板　厚</td><td class="num">${esc(o.ita)} mm</td></tr>
        <tr><td class="k">材料サイズ</td><td class="num">${esc(o.zairyo)}</td><td class="k">切断寸法</td><td class="num">${or(o.cut, "（未記入）")}</td></tr>
        <tr><td class="k">穴・抜き</td><td colspan="3">${esc(o.ana)}</td></tr>
        <tr><td class="k">数　量</td><td class="num">${esc(o.qty)} 個</td><td class="k">機　械</td><td>${mc ? `${esc(mc)}<span style="font-size:.75em;color:#6b7280">　${confirmed(o) ? "確定" : "希望"}</span>` : '<span style="color:#9ca3af;font-weight:400">（工場長が決める）</span>'}</td></tr>
        ${tokki("加工特記", o.tkako)}
        <tr class="blk"><td colspan="4">■ 曲げ・仕上げ</td></tr>
        <tr><td class="k">曲　げ</td><td>${esc(o.mage)}</td><td class="k">溶接・仕上げ</td><td>${or(o.yosetsu, "―")}</td></tr>
        <tr><td class="k">表面処理</td><td>${or(o.hyomen, "なし")}</td><td class="k">外注先</td><td>${or(o.gaichu, "社内")}</td></tr>
        <tr><td class="k">加工納期</td><td class="num">${or(md(o.kakouDate), "―")}</td><td class="k">納入先</td><td>${esc(o.saki)}</td></tr>
      </table>
      <div class="pg-foot">受注管理（架空データの見本）　出力 ${outStamp()}</div></div>`;
  }

  // 切断伝票（現品票）。切る指示であり、材料に付いて回る札。機械は工場長が後で決めるので載せない（本番と同じ）。
  const parseWH = (s) => { const m = String(s || "").match(/(\d+)\s*[×x]\s*(\d+)/); return m ? [Number(m[1]), Number(m[2])] : null; };
  function cutCount(o) {
    const cut = String(o.cut || "").trim();
    if (!cut) return "（切断寸法 記入後に自動算出）";
    if (cut.includes("定尺")) return "定尺のまま（切断不要・機械でネスティング）";
    const z = parseWH(o.zairyo), c = parseWH(cut), q = Number(o.qty) || 0;
    if (!z || !c) return "（材料サイズ・切断寸法 記入後に自動算出）";
    const k = Math.max(Math.floor(z[0] / c[0]) * Math.floor(z[1] / c[1]), Math.floor(z[0] / c[1]) * Math.floor(z[1] / c[0]));
    if (!k) return `切断寸法 ${cut} が材料 ${o.zairyo} に入らない`;
    const sheets = Math.ceil(q / k);
    return `材料 ${sheets}枚 × 取り数 ${k}（${cut}）＝ ${sheets * k}枚　（数量 ${q}個）`;
  }
  // 整合の確認。止めない・確定は営業の責任（本番と同じ考え）。
  function cutWarnings(o) {
    const w = [], t = parseFloat(o.ita) || 0, mc = machineOf(o);
    if (!String(o.cut || "").trim()) w.push("切断寸法 未記入");
    if (mc === "タレパン" && t > 3.2) w.push(`板厚 ${o.ita} はタレパンで抜けない（〜3.2mm）`);
    if (mc === "レーザーA" && (t > 6 || /SUS|^A\d/.test(o.zai))) w.push("レーザーAは鉄・〜6mm：ステンレス/アルミ/厚板はレーザーBへ");
    if (/定尺/.test(o.cut || "") && mc === "" && t > 6) w.push("厚板を定尺のまま＝機械の確定を先に");
    return w;
  }
  function pageCt(o) {
    const warns = cutWarnings(o);
    return `<div class="page ctp"><span class="pg-tag">架空データの見本</span>
      <div class="ct-top"><div class="ct-no num">${esc(o.no)}</div><div class="ct-who">${esc(o.eigyo)}</div></div>
      <table class="cts">
        <tr><td class="k">得意先名</td><td colspan="3">${esc(o.tokui)}</td></tr>
        <tr><td class="k">品　名</td><td colspan="3">${esc(o.hinmei)}${o.zuban ? `<small>${esc(o.zuban)}</small>` : ""}</td></tr>
        <tr><td class="k">材　質</td><td>${esc(o.zai)}</td><td class="k">板　厚</td><td class="num">${esc(o.ita)} mm</td></tr>
        <tr><td class="k">材料サイズ</td><td class="num">${esc(o.zairyo)}</td><td class="k">数　量</td><td class="num">${esc(o.qty)} 個</td></tr>
        <tr><td class="k">穴・抜き</td><td>${esc(o.ana)}</td><td class="k">曲　げ</td><td>${esc(o.mage)}</td></tr>
        <tr><td class="k">圧延方向</td><td colspan="3" class="blue">${or(o.atsuen, "（指定なし）")}</td></tr>
        <tr><td class="k">切断指定</td><td colspan="3" class="xl num">材料： ${esc(o.zairyo)}　　切断： ${or(o.cut, "（未記入）")}</td></tr>
        <tr><td class="k">切断枚数</td><td colspan="3" class="l num">${esc(cutCount(o))}</td></tr>
        <tr><td class="k">表面処理</td><td>${or(o.hyomen, "なし")}</td><td class="k">外注先</td><td>${or(o.gaichu, "社内")}</td></tr>
        <tr><td class="k">納　期</td><td colspan="3" class="num">${md(o.noukiDate)}${o.kakouDate ? `<small>加工納期 ${md(o.kakouDate)}</small>` : ""}</td></tr>
        ${warns.length ? `<tr class="warn"><td colspan="4">【要確認】 ${esc(warns.join("　／　"))}</td></tr>` : ""}
      </table>
      <div class="pg-foot">受注管理（架空データの見本）　この札は材料から離さない　出力 ${outStamp()}</div></div>`;
  }

  // 機械予定入れ（A4横・モノクロ）。画面で並べた順のまま。工場長が行を切り取り、各機械の台紙（B4）に貼る。
  const YI_HEAD = ["図面/データ", "注意事項", "得意先", "品名", "数量", "材質", "板厚", "切断寸法", "穴・抜き", "曲げ", "表面処理", "納期", "担当", "発注番号", "機械"];
  const YI_W = [64, 84, 84, 118, 40, 54, 36, 72, 88, 62, 78, 44, 34, 56, 56];
  function pagesYi(recs) {
    const tantou = [...new Set(recs.map((o) => o.eigyo))];
    const tantouDisp = tantou.length === 1 ? tantou[0] : `複数（${tantou.join("・")}）`;
    const per = 8, pages = [];
    for (let i = 0; i < recs.length; i += per) pages.push(recs.slice(i, i + per));
    const head = `<tr>${YI_HEAD.map((h, i) => `<th${i === YI_HEAD.length - 1 ? ' class="mc"' : ""}>${h}</th>`).join("")}</tr>`;
    const row = (o) => `<tr>
      <td class="st">${esc(statusOf(o))}</td><td class="l">${esc(noteOf(o))}</td><td class="l">${esc(o.tokui)}</td><td class="l">${esc(o.hinmei)}</td>
      <td class="num">${esc(o.qty)}</td><td>${esc(o.zai)}</td><td class="num">${esc(o.ita)}</td><td class="num">${esc(o.cut)}</td><td class="l">${esc(o.ana)}</td><td>${esc(o.mage)}</td><td class="l">${esc(o.hyomen)}</td>
      <td class="num">${paperNouki(o)}</td><td>${esc(o.eigyo)}</td><td class="small num">${esc(o.no)}</td><td class="mc">${esc(machineOf(o))}</td></tr>`;
    return pages.map((p, i) => `<div class="page land"><span class="pg-tag">架空データの見本</span>
      <h2 class="yi-h">機械予定入れ${pages.length > 1 ? `　<span style="font-size:12px;font-weight:400">（${i + 1}/${pages.length}）</span>` : ""}</h2>
      <p class="yi-sub">担当：${esc(tantouDisp)}　／　${recs.length} 件　／　発行 ${md(new Date())}　　（切り取り → 各機械の台紙へ。機械列＝割り当て先・希望）</p>
      <table class="yi"><colgroup>${YI_W.map((w) => `<col style="width:${w}px">`).join("")}</colgroup>${head}${p.map(row).join("")}</table>
      <div class="yi-foot" style="width:968px">受注管理（架空データの見本）／ 機械予定入れ（A4横で出す → 切って各機械の台紙[B4]へ）</div></div>`).join("");
  }

  function openSheet(kind, recs) {
    let html = "", note = "";
    if (kind === "wo") { html = recs.map(pageWo).join(""); note = `${recs.length} 枚（1案件1枚・A4縦）`; }
    if (kind === "ct") { html = recs.map(pageCt).join(""); note = `${recs.length} 枚（1案件1枚・A4縦・上の余白は材料に挟む分）`; }
    if (kind === "yi") { html = pagesYi(recs); note = `${recs.length} 件（画面で並べた順・A4横・切って各機械の台紙へ）`; }
    $("#sheetTitle").textContent = SHEET_NAME[kind];
    $("#sheetNote").textContent = note + "　PDFの代わりに、ブラウザからプリンターへ出します";
    $("#sheets").innerHTML = html;
    $("#sheet").hidden = false;
    $("#sheets").scrollTop = 0;
  }
  $("#closeSheet").onclick = () => { $("#sheet").hidden = true; $("#sheets").innerHTML = ""; };
  $("#doPrint").onclick = () => window.print();
  document.addEventListener("keydown", (e) => { if (e.key === "Escape" && !$("#sheet").hidden) $("#closeSheet").click(); });

  $("#reset").onclick = () => {
    try { localStorage.removeItem(STORE_KEY); } catch (e) { /* なし */ }
    listOrder = null; assignDay = null; assignEigyo = "すべて"; Object.assign(F, { eigyo: "すべて", day: "", month: false });
    initState(); render(); toast("最初の場面に戻しました");
  };

  S = load();
  if (!S) initState();
  window.addEventListener("hashchange", render);
  render();
})();
