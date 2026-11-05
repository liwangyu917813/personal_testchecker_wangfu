/* =====================================================================
 * 看板交互逻辑 (app.js) —— 依赖 core.js 暴露的 window.ExamCore / 全局 XLSX
 *
 * 架构（v5 整理）：
 *   状态      parsed / workbook / lastResult / source（模块内闭包持有）
 *   数据入口  A: tryFetchRemote（自动拉取 data/exam.xlsx）
 *             B: 手动上传（拖拽/点选，备用路径，file:// 打开时为唯一路径）
 *             两入口汇聚到 applyParsed(name, parsed, tag)
 *   查询      doQuery → ExamCore.findExamsForStudent → renderBoard
 *   渲染      renderBoard（分组/倒计时/徽标/动效） + buildMeta（卡片元信息）
 *   工具      showMsg / escapeHtml / 日期格式化 / 天数差
 * ===================================================================== */
'use strict';

(function () {
  const $ = (id) => document.getElementById(id);
  let parsed = null;      // parseWorkbook 的结果
  let workbook = null;    // SheetJS workbook
  let lastResult = null;  // 最近一次查询结果
  let source = 'none';    // 数据来源：'remote'（在线数据文件）| 'manual'（手动上传）| 'none'

  /* ---------- 数据应用：解析 → 界面（自动/手动两个入口共用） ---------- */
  function applyParsed(name, p, tag) {
    parsed = p;
    source = tag;
    $('fileName').textContent = name + '（' + p.exams.length + ' 场考试）';
    if (tag === 'manual') $('fileChosen').classList.remove('hidden');
    $('board').classList.add('hidden');
    let extra = '';
    if (p.meta.warnings.length) {
      extra = '<br><b>注意：有 ' + p.meta.warnings.length + ' 行数据格式异常，已跳过：</b><br>' +
        p.meta.warnings.slice(0, 8).map(escapeHtml).join('<br>') +
        (p.meta.warnings.length > 8 ? '<br>… 以及另外 ' + (p.meta.warnings.length - 8) + ' 条' : '');
    }
    if (p.meta.repairs && p.meta.repairs.length) {
      // 日期自动修正透明展示（冗余：让用户知道哪些日期被智能修正过）
      extra += '<br><b>🛠 已自动修正 ' + p.meta.repairs.length + ' 处日期：</b><br>' +
        p.meta.repairs.slice(0, 6).map((rp) =>
          escapeHtml('"' + rp.raw + '"：' + rp.from + ' → ' + rp.to +
            (rp.reason ? '（' + rp.reason + '）' : ''))).join('<br>') +
        (p.meta.repairs.length > 6 ? '<br>… 以及另外 ' + (p.meta.repairs.length - 6) + ' 处' : '');
    }
    if (tag === 'manual') {
      showMsg('parseMsg', 'info', '✅ 解析成功：共 ' + p.exams.length + ' 场考试，' +
        countStudents(p) + ' 名学生。请在下方输入姓名查询。' + extra);
    } else {
      showMsg('parseMsg', '', '');   // 在线加载成功时结果已显示在 fetchMsg，不重复展示
    }
    $('queryCard').classList.remove('hidden');
    if (tag === 'remote') {
      // 自动加载数据成功后聚焦姓名框（键盘直接可输入）
      try { $('nameInput').focus(); } catch (e) { /* 忽略 */ }
    }
  }

  /** v7：候选名单构建（doQuery 的 partial 分支与输入框实时联想共用）。
   *  每个候选按钮显示 姓名（N 场）+ 匹配徽标（前缀/包含/同字）；点击即查该同学。 */
  function buildCandidateList(candidates, typedName, titleHtml) {
    const wrap = document.createElement('div');
    if (titleHtml) {
      const t = document.createElement('div');
      t.className = 'title';
      t.innerHTML = titleHtml;
      wrap.appendChild(t);
    }
    // 候选过多时只显示前 60 个（极端输入如单字"王"可能命中上百人），提示可用更长的输入缩小范围
    const MAX_SHOW = 60;
    candidates.slice(0, MAX_SHOW).forEach((c) => {
      const b = document.createElement('button');
      b.type = 'button';
      const tag = c.score >= 100 ? '前缀' : (c.score >= 10 ? '同字' : '包含');
      b.innerHTML = escapeHtml(c.student) + '<span class="candCount">' + c.examCount + ' 场</span>' +
        '<span class="candTag">' + tag + '</span>';
      b.addEventListener('click', () => {
        $('nameInput').value = c.student;
        try { localStorage.setItem(LS_KEY, c.student); } catch (e) { /* 忽略 */ }
        const r2 = ExamCore.findExamsForStudent(parsed, c.student);
        if (r2.status === 'found') { showMsg('queryMsg', '', ''); renderBoard(c.student, r2.hits.map((h) => h.exam)); }
        else if (r2.status === 'partial') { showMsg('queryMsg', 'info', '「' + escapeHtml(c.student) + '」仍有多个匹配：'); doQuery(); }
      });
      wrap.appendChild(b);
    });
    if (candidates.length > MAX_SHOW) {
      const more = document.createElement('div');
      more.className = 'candMore';
      more.textContent = '共 ' + candidates.length + ' 位候选人，已显示前 ' + MAX_SHOW + ' 位——输入更多字可缩小范围';
      wrap.appendChild(more);
    }
    return wrap;
  }

  /** v7：输入框实时联想（debounce 160ms）。输入 ≥1 个字符即列出候选，完整姓名直接查。 */
  let suggestTimer = null;
  function onNameInputLive() {
    if (suggestTimer) clearTimeout(suggestTimer);
    suggestTimer = setTimeout(() => {
      if (!guardParsed()) return;
      const name = $('nameInput').value.trim();
      const box = $('candList');
      if (!name) { box.innerHTML = ''; showMsg('queryMsg', '', ''); return; }
      let res;
      try { res = ExamCore.findExamsForStudent(parsed, name); } catch (e) { return; }
      if (res.status === 'partial') {
        showMsg('queryMsg', 'info', '找到 ' + res.candidates.length + ' 位含「' + escapeHtml(name) + '」的同学（点击选择）：');
        box.innerHTML = '';
        box.appendChild(buildCandidateList(res.candidates, name, ''));
      } else if (res.status === 'found') {
        // 完整命中：清掉联想（用户按回车或点查询按钮时才渲染看板，避免打字过程中页面跳动）
        box.innerHTML = '';
        showMsg('queryMsg', '', '');
      } else if (res.status === 'not-found') {
        box.innerHTML = '';
      }
    }, 160);
  }

  function showParseError(err) {
    parsed = null;
    $('fileChosen').classList.add('hidden');
    $('queryCard').classList.add('hidden');
    $('board').classList.add('hidden');
    showMsg('parseMsg', 'error', '❌ 解析失败：' + escapeHtml(err && err.message ? err.message : String(err)) +
      '<br><br>请确认提交的是<b>最新版监考表原件</b>（需包含 日期/时间/课程名称 等列）。');
  }

  /* ---------- 入口 A：打开页面自动加载仓库内的在线数据文件 ---------- */
  const DATA_URL = 'data/exam.xlsx';
  /** 入口 A：fetch 仓库内 data/exam.xlsx（no-cache，带 2 次重试）；失败时中文提示并进入手动模式 */
  async function tryFetchRemote() {
    const intro = $('dataIntro'), fm = $('fetchMsg');
    // 冗余：瞬时网络抖动重试 2 次（指数退避 300ms/900ms），只对网络/5xx 类失败重试
    let resp = null, lastErr = null;
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        if (attempt > 0) await new Promise((r) => setTimeout(r, attempt === 1 ? 300 : 900));
        resp = await fetch(DATA_URL, { cache: 'no-cache' });
        if (resp.ok) break;
        if (resp.status >= 500 || resp.status === 0) { lastErr = new Error('HTTP ' + resp.status); resp = null; continue; }
        lastErr = new Error('HTTP ' + resp.status);   // 404 等客户端错误：重试无意义
        resp = null;                                   // 交给下方 throw lastErr 统一报错
        break;
      } catch (e) { lastErr = e; resp = null; }       // TypeError（断网/DNS）→ 重试
    }
    try {
      if (!resp) throw lastErr || new Error('网络不可用');
      const buf = await resp.arrayBuffer();
      const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: false });
      const p = ExamCore.parseWorkbook(wb, { year: new Date().getFullYear() });
      if (!p.exams.length) throw new Error('数据文件中没有解析出任何考试');
      applyParsed('在线数据（exam.xlsx）', p, 'remote');
      intro.innerHTML = '本页数据由管理员统一更新：' +
        '<b>共 ' + p.exams.length + ' 场考试 · ' + countStudents(p) + ' 名学生</b>。直接在下方输入姓名即可查询。';
      showMsg('fetchMsg', 'ok', '✅ 考试数据已自动加载（' + p.exams.length + ' 场考试，' +
        countStudents(p) + ' 名学生）。' + buildExtra(p));
      exitManualMode();   // 在线加载成功 → 保持手动上传区与切换链接隐藏，页面最简
    } catch (err) {
      source = 'none';
      showMsg('fetchMsg', 'error',
        '⚠️ 在线数据暂时无法加载（' + escapeHtml(String(err && err.message || err)) + '）。' +
        '可以点击下方「手动选择本地文件」临时使用本地的监考表 Excel 查询。');
      enterManualMode();
    }
  }
  function buildExtra(p) {
    let extra = '';
    if (p.meta.warnings.length) {
      extra += '<br><b>注意：有 ' + p.meta.warnings.length + ' 行数据格式异常，已跳过：</b><br>' +
        p.meta.warnings.slice(0, 8).map(escapeHtml).join('<br>') +
        (p.meta.warnings.length > 8 ? '<br>… 以及另外 ' + (p.meta.warnings.length - 8) + ' 条' : '');
    }
    if (p.meta.repairs && p.meta.repairs.length) {
      extra += '<br><b>🛠 已自动修正 ' + p.meta.repairs.length + ' 处日期：</b><br>' +
        p.meta.repairs.slice(0, 6).map((rp) =>
          escapeHtml('"' + rp.raw + '"：' + rp.from + ' → ' + rp.to +
            (rp.reason ? '（' + rp.reason + '）' : ''))).join('<br>') +
        (p.meta.repairs.length > 6 ? '<br>… 以及另外 ' + (p.meta.repairs.length - 6) + ' 处' : '');
    }
    return extra;
  }

  /* ---------- 入口 B：手动上传本地文件（备用） ---------- */
  const drop = $('drop'), fileInput = $('fileInput');

  drop.addEventListener('click', () => fileInput.click());
  drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('over'));
  drop.addEventListener('drop', (e) => {
    e.preventDefault(); drop.classList.remove('over');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change', () => {
    if (fileInput.files && fileInput.files[0]) handleFile(fileInput.files[0]);
  });
  $('clearFile').addEventListener('click', () => {
    fileInput.value = ''; parsed = null; workbook = null; source = 'none';
    $('fileChosen').classList.add('hidden');
    $('queryCard').classList.add('hidden');
    $('board').classList.add('hidden');
    $('parseMsg').innerHTML = '';
  });
  // 展开/收起手动上传区（在线加载成功时默认收起，失败或本地打开时自动展开）
  $('toggleManual').addEventListener('click', () => {
    const hidden = $('manualZone').classList.contains('hidden');
    if (hidden) enterManualMode();
    else exitManualMode();
  });
  function enterManualMode() {
    $('manualZone').classList.remove('hidden');
    $('toggleRow').classList.remove('hidden');   // 手动模式里"收起"入口要可见
    $('toggleManual').textContent = '收起手动上传 ↑';
  }
  function exitManualMode() {
    $('manualZone').classList.add('hidden');
    $('toggleManual').textContent = '手动选择本地文件 →';
  }

  /** 入口 B：读取用户选择的本地文件 → XLSX.read → parseWorkbook → applyParsed。
   *  双保险：.xls 老格式先直接尝试解析（很多 WPS 导出的 .xls 实际是 xlsx 内容），
   *  失败才提示转存；非 Excel 扩展名也先尝试，给出内容嗅探的机会。 */
  function handleFile(f) {
    const name = f.name || '';
    if (/\.(xlsx|xlsm|xls)$/i.test(name)) {
      readFileAsWorkbook(f, name);
      return;
    }
    // 未知扩展名：先试内容，不行再报错（冗余：有些系统导出的文件扩展名不规范）
    readFileAsWorkbook(f, name, () => showMsg('parseMsg', 'error',
      '❌ 只支持 .xlsx 格式的 Excel 文件（"' + escapeHtml(name) + '" 无法按 Excel 解析）。' +
      '老版 .xls 请先在 Excel/WPS 中另存为 .xlsx 再提交。'));
  }

  function readFileAsWorkbook(f, name, onFail) {
    const reader = new FileReader();
    reader.onerror = () => showMsg('parseMsg', 'error', '❌ 文件读取失败，请重新选择文件再试。');
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        let workbook;
        try {
          workbook = XLSX.read(data, { type: 'array', cellDates: false });
        } catch (readErr) {
          // 双保险第二路：xlsx 内容可能被包装成 CFB（.xls 容器）——XLSX.read 通常都能处理；
          // 走到这里说明确实不是 Excel，交给 onFail 或统一错误
          throw readErr;
        }
        const p = ExamCore.parseWorkbook(workbook, { year: new Date().getFullYear() });
        applyParsed(name, p, 'manual');
      } catch (err) {
        if (typeof onFail === 'function') { onFail(); return; }
        showParseError(err);
      }
    };
    reader.readAsArrayBuffer(f);
  }

  /** 统计去重学生总数（跨场次合并） */
  function countStudents(p) {
    const s = new Set();
    p.exams.forEach((ex) => ex.students.forEach((st) => s.add(ExamCore.normalizeName(st))));
    return s.size;
  }

  /* ---------- 步骤 2：查询 ---------- */
  $('goBtn').addEventListener('click', doQuery);
  $('nameInput').addEventListener('keydown', (e) => { if (e.key === 'Enter') doQuery(); });
  // v7：输入即联想（防抖 160ms），写"王"立刻列出所有含"王"的同学
  $('nameInput').addEventListener('input', onNameInputLive);

  // 记住上次查询的姓名（localStorage，仅本机）
  const LS_KEY = 'examBoard.lastName';
  try {
    const saved = localStorage.getItem(LS_KEY);
    if (saved) $('nameInput').value = saved;
  } catch (e) { /* 隐私模式等场景忽略 */ }

  /** 查询主流程：空名校验 → 精确/部分匹配 → 同名提示 → 记住姓名 → 渲染看板 */
  function doQuery() {
    if (!guardParsed()) { showMsg('queryMsg', 'error', '考试数据还没加载好：请等在线数据加载完成，或在上方「手动选择本地文件」提交监考表。'); return; }
    const name = $('nameInput').value;
    try { localStorage.setItem(LS_KEY, name.trim()); } catch (e) { /* 忽略 */ }
    const res = ExamCore.findExamsForStudent(parsed, name);
    lastResult = res;
    $('candList').innerHTML = '';
    if (res.status === 'empty') {
      showMsg('queryMsg', 'error', '请输入姓名后再查询。'); return;
    }
    if (res.status === 'found') {
      let msg = '';
      if (res.ambiguous) {
        msg = '⚠️ <b>检测到可能的同名情况：' + escapeHtml(res.ambiguityReason || '') + '。</b>下列结果可能混合了同名同学的考试，请核对班级与考场。<br>';
      }
      showMsg('queryMsg', res.ambiguous ? 'warn' : '', msg);
      renderBoard(name, res.hits.map((h) => h.exam));
      return;
    }
    if (res.status === 'partial') {
      const box = $('candList');
      box.innerHTML = '';
      box.appendChild(buildCandidateList(res.candidates, name.trim(),
        '没有找到与「' + escapeHtml(name.trim()) + '」完全一致的姓名，以下是相近的候选（点击选择）：'));
      return;
    }
    showMsg('queryMsg', 'error',
      '❌ 未找到「' + escapeHtml(name.trim()) + '」的任何考试记录。<br>' +
      '请检查：① 姓名是否输入正确（不要有空格）；② 当前数据是否是<b>包含你这个年级</b>的最新版监考表' +
      (source === 'manual' ? '（当前使用的是手动提交的文件）' : '') + '。');
  }

  /** 冗余防护：解析结果对象被意外破坏（如扩展冲突/手工篡改）时兜底，避免 UI 崩溃白屏 */
  function guardParsed() {
    if (!parsed) return false;
    if (!parsed.exams || !Array.isArray(parsed.exams) ||
        !parsed.meta || !Array.isArray(parsed.meta.warnings)) {
      parsed = null;
      showMsg('queryMsg', 'error', '内部数据结构异常，请重新提交监考表。');
      return false;
    }
    return true;
  }

  /* ---------- 看板渲染 ---------- */
  function renderBoard(name, exams) {
    const board = $('board');
    const now = new Date();
    const todayStr = fmtLocalDate(now);

    // 排序：按日期 + 开始时间
    exams = exams.slice().sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time));

    // 状态计算
    const nowMin = now.getHours() * 60 + now.getMinutes();
    const toMin = (t) => {
      const m = String(t).match(/(\d{1,2}):(\d{2})/);
      return m ? (+m[1]) * 60 + (+m[2]) : null;
    };
    let nextIdx = -1;
    exams.forEach((ex, i) => {
      const st = toMin(ex.time);
      const en = st != null && ex.minutes ? st + ex.minutes : (st != null ? st + 60 : null);
      ex._past = ex.date < todayStr || (ex.date === todayStr && en != null && nowMin >= en);
      ex._today = ex.date === todayStr;
      // 正在进行的考试：今天、已开始、未结束（结束时间未知按开始后 60 分钟估算）
      ex._now = ex.date === todayStr && st != null && en != null && nowMin >= st && nowMin < en;
      if (ex._now && nextIdx < 0) nextIdx = i;          // 正在进行的优先标为"当前"
      if (!ex._past && nextIdx < 0) nextIdx = i;
    });

    // 倒计时 + 全天进度条
    let cd = '', cdBar = '';
    if (nextIdx >= 0) {
      const nx = exams[nextIdx];
      const d = daysBetween(todayStr, nx.date);
      if (nx._now) {
        const en = toMin(nx.time) + (nx.minutes || 60);
        const remainPct = Math.max(0, Math.min(100, ((en - nowMin) / (nx.minutes || 60)) * 100));
        cd = '<div class="countdown cd-today">🔥 ' + escapeHtml(nx.subject) + ' 正在进行，预计 ' +
          Math.max(0, en - nowMin) + ' 分钟后结束</div>';
        cdBar = '<div class="cdBar hot"><i style="width:' + remainPct.toFixed(1) + '%"></i></div>';
      } else if (d > 1) {
        cd = '<div class="countdown cd-before">⏳ 距离下一场考试（' + escapeHtml(nx.subject) + '）还有 <b>' + d + '</b> 天</div>';
        cdBar = '<div class="cdBar"><i style="width:' + Math.max(3, Math.min(100, (1 - d / 30) * 100)).toFixed(1) + '%"></i></div>';
      } else if (d === 1) {
        cd = '<div class="countdown cd-before">⏳ 下一场考试（' + escapeHtml(nx.subject) + '）就在<b>明天</b>！</div>';
        cdBar = '<div class="cdBar"><i style="width:88%"></i></div>';
      } else {
        const st = toMin(nx.time);
        const pct = st != null && st > 0 ? (nowMin / st) * 100 : 3;   // 今天距开考的昼夜进度
        cd = '<div class="countdown cd-today">🔥 今天有考试，下一场是 ' + escapeHtml(nx.time) + ' 的 ' + escapeHtml(nx.subject) + '</div>';
        cdBar = '<div class="cdBar hot"><i style="width:' + Math.min(97, Math.max(3, pct)).toFixed(1) + '%"></i></div>';
      }
    } else {
      cd = '<div class="countdown cd-over">🎉 所有考试已结束，辛苦了！</div>';
    }
    // 按日期分组
    const groups = [];
    exams.forEach((ex) => {
      let g = groups.find((x) => x.date === ex.date);
      if (!g) { g = { date: ex.date, items: [] }; groups.push(g); }
      g.items.push(ex);
    });

    let html = '<div class="card"><div class="boardHead"><h2>📋 ' + escapeHtml(name) + ' 的考试看板</h2>' +
      '<span class="cnt">共 ' + exams.length + ' 场 · ' + groups.length + ' 个考试日</span></div>' + cd + cdBar;

    for (const g of groups) {
      const md = g.date.slice(5).replace('-', ' 月 ') + ' 日';
      const wd = ExamCore.weekdayCN(g.date);
      const isToday = g.date === todayStr;
      html += '<div class="dayGroup"><div class="dayTitle' + (isToday ? ' todayRow' : '') + '"><span>' + md + '（' + wd + '）</span>' +
        (isToday ? '<span class="todayBadge">今天</span>' : '') + '</div>';
      for (const ex of g.items) {
        const idx = (typeof ex._row === 'number') ? ex._row : exams.indexOf(ex);   // v7.1：优先用全表 _row，杜绝个人数组下标与全表下标错位
        const cls = 'examCard clickable' + (ex._past ? ' past' : '') + (ex._now ? ' isNow' : '') + (idx === nextIdx && !ex._now ? ' next' : '');
        const dur = ex.minutes ? '<span class="dur">时长 ' + ex.minutes + ' 分钟</span>' : '';
        const nowB = ex._now ? '<span class="nowBadge">● 进行中</span>' : '';
        const nextB = (idx === nextIdx && !ex._now) ? '<span class="nextBadge">← 下一场</span>' : '';
        const fixB = ex.repairedFrom ? '<span class="fixBadge" title="原表日期为 ' + escapeHtml(ex.repairedFrom) + '，因脱离考试日期连续段且符合日期尾零丢失特征已自动修正">🛠 日期已修正</span>' : '';
        // v7：点击卡片弹出本场详情（监考老师 + 同场学生）。data-idx 关联 parsed.exams 下标。
        html += '<div class="' + cls + '" data-exidx="' + idx + '" title="点击查看本场详情（监考老师 / 同场考生）">' +
          '<div class="time">' + escapeHtml(ex.time) + dur + '</div>' +
          '<div><div class="subj">' + escapeHtml(ex.subject) + nowB + nextB + fixB + '</div>' +
          '<div class="meta">' + buildMeta(ex) + '</div></div>' +
          '<div class="roomTag">📍 ' + escapeHtml(ex.room) + '<span class="detailHint">详情 ›</span></div>' +
          '</div>';
      }
      html += '</div>';
    }

    html += '<div class="actions no-print">' +
      '<button class="btn" onclick="window.print()">🖨️ 打印 / 存为 PDF</button>' +
      '<button class="btn ghost" id="copyBtn">📋 复制为文本</button>' +
      '<button class="btn ghost" id="icsBtn">📅 导出日历 (.ics)</button>' +
      '<button class="btn ghost" onclick="window.__examQueryNew()">🔍 换个人查询</button>' +
      '<span class="copyTip hidden" id="copyTip">已复制 ✓</span>' +
      '</div>';
    html += '<div class="footer">生成于 ' + fmtLocalDateTime(now) + ' · 数据来自' +
      (source === 'manual' ? '手动提交的监考表' : '管理员发布的在线数据') + ' · 共 ' +
      parsed.meta.totalRows + ' 场考试参与匹配</div>';

    board.innerHTML = html;
    // v7：事件委托绑定卡片点击 → 详情弹层（比逐卡绑定更省，且打印前重渲染也安全）
    board.addEventListener('click', (ev) => {
      const cardEl = ev.target.closest ? ev.target.closest('.examCard.clickable') : null;
      if (!cardEl) return;
      const idx = parseInt(cardEl.getAttribute('data-exidx'), 10);
      if (!Number.isInteger(idx)) return;
      openExamDetail(idx, name);
    });
    // 复制为纯文本（冗余出口：打印之外再给一个分享渠道）
    const copyBtn = document.getElementById('copyBtn');
    if (copyBtn) {
      copyBtn.addEventListener('click', () => {
        const text = exams.map((ex) =>
          ex.date + '（' + ExamCore.weekdayCN(ex.date) + '）' + ex.time +
          (ex.minutes ? ' 共' + ex.minutes + '分钟' : '') + '  ' + ex.subject +
          '  @' + ex.room).join('\n');
        const done = () => {
          const tip = document.getElementById('copyTip');
          if (tip) { tip.textContent = '已复制 ✓'; tip.classList.remove('hidden'); setTimeout(() => tip.classList.add('hidden'), 2000); }
        };
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(name + ' 的考试安排：\n' + text).then(done).catch(() => {});
        } else {
          // 兼容非安全上下文（file:// 打开时部分浏览器无 clipboard API）
          const ta = document.createElement('textarea');
          ta.value = name + ' 的考试安排：\n' + text;
          document.body.appendChild(ta); ta.select();
          try { document.execCommand('copy'); done(); } catch (e) { /* 忽略 */ }
          document.body.removeChild(ta);
        }
      });
    }
    // 导出 .ics 日历（冗余出口：可导入手机/电脑日历 App 提前提醒）
    const icsBtn = document.getElementById('icsBtn');
    if (icsBtn) {
      icsBtn.addEventListener('click', () => {
        try {
          const pad = (n) => String(n).padStart(2, '0');
          const toIcsDT = (dateStr, hm) => {
            const p = dateStr.split('-');
            const m = String(hm || '08:00').match(/(\d{1,2}):(\d{2})/);
            return p[0] + p[1] + p[2] + 'T' + pad(+m[1]) + m[2] + '00';
          };
          const endHM = (ex) => {
            const m = String(ex.time).match(/(\d{1,2}):(\d{2})/g);
            if (m && m.length >= 2) return m[1];
            const st = m ? (+m[0].slice(0, 2)) * 60 + (+m[0].slice(3)) : 480;
            const en = st + (ex.minutes || 60);
            return pad(Math.floor(en / 60)) + ':' + pad(en % 60);
          };
          const esc = (s) => String(s).replace(/\\/g, '\\\\').replace(/[,;]/g, (c) => '\\' + c).replace(/\n/g, '\\n');
          const nowStamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
          let ics = 'BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//ExamBoard//CN//\r\nCALSCALE:GREGORIAN\r\n';
          exams.forEach((ex, i) => {
            ics += 'BEGIN:VEVENT\r\n' +
              'UID:exam-' + nowStamp + '-' + i + '@examboard.local\r\n' +
              'DTSTAMP:' + nowStamp + '\r\n' +
              'DTSTART:' + toIcsDT(ex.date, ex.time) + '\r\n' +
              'DTEND:' + toIcsDT(ex.date, endHM(ex)) + '\r\n' +
              'SUMMARY:' + esc(ex.subject + (ex.room && ex.room !== '（未标注）' ? ' @' + ex.room : '')) + '\r\n' +
              'DESCRIPTION:' + esc(name + ' 的考试 · ' + ex.time + (ex.minutes ? ' 共' + ex.minutes + '分钟' : '') +
                (ex.grade ? ' · 年级' + ex.grade : '') + (ex.klass ? ' · ' + ex.klass : '')) + '\r\n' +
              (ex.room && ex.room !== '（未标注）' ? 'LOCATION:' + esc(ex.room) + '\r\n' : '') +
              'BEGIN:VALARM\r\nTRIGGER:-PT30M\r\nACTION:DISPLAY\r\nDESCRIPTION:30分钟后开考：' + esc(ex.subject) + '\r\nEND:VALARM\r\n' +
              'END:VEVENT\r\n';
          });
          ics += 'END:VCALENDAR\r\n';
          const blob = new Blob([ics], { type: 'text/calendar;charset=utf-8' });
          const a = document.createElement('a');
          a.href = URL.createObjectURL(blob);
          a.download = name + '的考试安排.ics';
          document.body.appendChild(a); a.click();
          setTimeout(() => { URL.revokeObjectURL(a.href); document.body.removeChild(a); }, 800);
          const tip = document.getElementById('copyTip');
          if (tip) { tip.textContent = '日历已导出 ✓'; tip.classList.remove('hidden'); setTimeout(() => tip.classList.add('hidden'), 2200); }
        } catch (e) { /* 导出失败不影响主流程 */ }
      });
    }
    board.classList.remove('hidden');
    board.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  /** v7：详情弹层。点击考试卡片 → 展示监考老师 + 同场考生 + 完整信息。
   *  渲染在独立遮罩层，Esc / 点遮罩 / 关闭按钮均可退出；学生名单点击可直接查该同学。 */
  function openExamDetail(idx, queryName) {
    // v7.1 修复：idx 现在传入的是考试对象的 _row（全表源行号）。按 _row 查详情；
    // 若查不到（旧数据兼容）再退回下标语义。另加防线：详情里必须包含查询人本人，否则视为串场拒绝展示。
    let d = (typeof ExamCore.getExamDetailByRow === 'function') ? ExamCore.getExamDetailByRow(parsed, idx) : null;
    if (!d) d = ExamCore.getExamDetail(parsed, idx);
    if (!d) return;   // 冗余防护：数据结构异常时静默退出，不影响看板
    // 遮罩层只建一次，重复调用复用
    let overlay = document.getElementById('detailOverlay');
    if (!overlay) {
      overlay = document.createElement('div');
      overlay.id = 'detailOverlay';
      overlay.innerHTML =
        '<div class="detailMask"></div>' +
        '<div class="detailModal" role="dialog" aria-modal="true">' +
        '  <button class="detailClose" id="detailClose" title="关闭 (Esc)">✕</button>' +
        '  <div id="detailBody"></div>' +
        '</div>';
      document.body.appendChild(overlay);
      // v7.1 修复：改用 closest 判定（✕ 内部文本/遮罩/按钮任意命中都关闭），并给按钮直绑双保险
      overlay.addEventListener('click', (ev) => {
        const t = ev.target;
        if (!t || !t.closest) return;
        if (t.closest('.detailClose') || t.closest('.detailMask')) closeExamDetail();
      });
      document.addEventListener('keydown', (ev) => { if (ev.key === 'Escape') closeExamDetail(); });
    }
    const closeBtn = overlay.querySelector('#detailClose');
    if (closeBtn) closeBtn.addEventListener('click', closeExamDetail);   // 双保险：直绑一次
    const body = overlay.querySelector('#detailBody');
    const headRow = '<div class="dtDate">' + escapeHtml(d.date) + '（' + escapeHtml(d.weekday) + '）' +
      '<span class="dtTime">' + escapeHtml(d.time) +
      (d.minutes ? ' · 共 ' + d.minutes + ' 分钟' : '') + '</span>' +
      (d.repairedFrom ? '<span class="fixBadge" title="原表日期为 ' + escapeHtml(d.repairedFrom) + '">🛠 已修正（原 ' + escapeHtml(d.repairedFrom) + '）</span>' : '') +
      '</div>';
    const metaRow = '<div class="dtMeta">' +
      '<span class="dtSubject">' + escapeHtml(d.subject) + '</span>' +
      '<span class="dtRoom">📍 ' + escapeHtml(d.room) + '</span>' +
      (d.grade ? '<span class="dtTag">年级 ' + escapeHtml(d.grade) + '</span>' : '') +
      (d.klass ? '<span class="dtTag">' + escapeHtml(d.klass) + '</span>' : '') +
      '</div>';
    const proctorsHtml = '<div class="dtSection"><div class="dtSecTitle">👩‍🏫 监考老师（' + d.proctors.length + '）</div>' +
      (d.proctors.length
        ? '<div class="dtChips">' + d.proctors.map((p) => '<span class="dtChip proctor">' + escapeHtml(p) + '</span>').join('') + '</div>'
        : '<div class="dtNone">原表未标注本场监考老师</div>') +
      '</div>';
    const studentsHtml = '<div class="dtSection"><div class="dtSecTitle">🧑‍🎓 同场考生（' + d.students.length + ' 人）' +
      '<span class="dtSecHint">点击姓名可查询该同学的考试</span></div>' +
      '<div class="dtChips">' + d.students.map((st) =>
        '<span class="dtChip student" data-name="' + escapeHtml(st) + '">' + escapeHtml(st) + '</span>').join('') + '</div>' +
      '</div>';
    body.innerHTML = headRow + metaRow + proctorsHtml + studentsHtml;
    // 学生名单可点击 → 关弹层 + 直接查该同学（复用主查询流程，记住姓名逻辑同步生效）
    body.querySelectorAll('.dtChip.student').forEach((chip) => {
      chip.addEventListener('click', () => {
        const nm = chip.getAttribute('data-name') || '';
        closeExamDetail();
        $('nameInput').value = nm;
        doQuery();
      });
    });
    overlay.classList.add('open');
    document.body.classList.add('modalOpen');   // 锁定背景滚动
    if (closeBtn) closeBtn.focus();
  }

  function closeExamDetail() {
    const overlay = document.getElementById('detailOverlay');
    if (overlay) overlay.classList.remove('open');
    document.body.classList.remove('modalOpen');
  }
  // 供外部（如控制台调试）安全关闭
  window.__examDetailClose = closeExamDetail;

  /** 单场考试卡片右侧元信息（星期/时长/倒计时文案） */
  function buildMeta(ex) {
    const parts = [];
    if (ex.grade) parts.push('年级 ' + escapeHtml(ex.grade));
    if (ex.klass) parts.push(escapeHtml(ex.klass));
    if (ex.minutes) parts.push(ex.minutes + ' 分钟');
    return parts.length ? parts.join(' · ') : '&nbsp;';
  }

  window.__examQueryNew = function () {
    $('nameInput').value = '';
    $('board').classList.add('hidden');
    $('candList').innerHTML = '';
    showMsg('queryMsg', '', '');
    $('nameInput').focus();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  /* ---------- 小工具 ---------- */
  function showMsg(id, type, html) {
    const el = $(id);
    if (!html) { el.innerHTML = ''; return; }
    el.innerHTML = '<div class="msg ' + type + '">' + html + '</div>';
  }
  function escapeHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function fmtLocalDate(d) {
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  function fmtLocalDateTime(d) {
    return fmtLocalDate(d) + ' ' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
  }
  function daysBetween(a, b) {
    const da = new Date(a + 'T00:00:00'), db = new Date(b + 'T00:00:00');
    return Math.round((db - da) / 86400000);
  }

  /* ---------- 启动：优先自动加载在线数据；双击本地打开（file://）时直接进手动模式 ---------- */
  if (location.protocol === 'file:') {
    $('dataIntro').innerHTML = '当前以<b>本地文件方式</b>打开：请提交监考表 Excel 生成你的个人看板（文件只在本机解析）。';
    enterManualMode();
  } else {
    tryFetchRemote();
  }
})();
