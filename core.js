/* =====================================================================
 * 考试查询核心逻辑 (core.js)  v2 —— 增强容错版
 * 职责：解析监考表 XLSX -> 结构化考试数组 -> 按姓名查找
 * 设计原则：任何异常都给出明确中文提示，绝不静默出错；
 *           对日期、时间、表头、姓名均做多重冗余识别。
 * ===================================================================== */

'use strict';

/* ---------- 基础工具 ---------- */

// 规范化姓名：去空格/全角空格/间隔号，全角转半角，转小写（英文姓名大小写不敏感）
function normalizeName(s) {
  let t = String(s == null ? '' : s);
  t = t.replace(/[\uFF01-\uFF5E\u3000]/g, (ch) =>
    ch === '\u3000' ? ' ' : String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
  return t.replace(/[\s\u00B7\u2022\u2027·''''"""]/g, '').trim().toLowerCase();
}

function fullWidthToAscii(s) {
  return String(s).replace(/[\uFF01-\uFF5E\u3000]/g, (ch) =>
    ch === '\u3000' ? ' ' : String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/[\uFF0D\u2010\u2011\u2012\u2013\u2014\u2015\u2212\uFF0F]/g, '-'); // 全角横线/各类破折号/全角斜杠 -> '-'
}

function isEmptyDataRow(row) {
  for (let c = 0; c < Math.min(11, row.length); c++) {
    const v = row[c];
    if (v !== undefined && v !== null && String(v).trim() !== '') return false;
  }
  return true;
}

/* ---------- 时间解析（多重格式冗余） ---------- */
/* 支持：800-930 / 8:00-9:30 / 8．00—9：30 / 800~930 / 8-10 / 8时00分-9时30分 /
 *       上午8:00-9:30 / 8am-9:30am / 下午2点-4点 等写法 */
function fmtTimeRange(s) {
  if (s == null) return null;
  let raw = fullWidthToAscii(s).trim()
    .replace(/[—–＿~～﹣]/g, '-')
    .replace(/[至到]/g, '-')
    .replace(/\s+/g, '');
  // 时段词（上午/下午/am/pm...）：pm 类 shift+12，am 类直接剔除
  let pmShift = false;
  raw = raw.replace(/下午|午后|傍晚|晚上|夜里|夜间|夜晚|pm/gi, () => { pmShift = true; return ''; });
  raw = raw.replace(/上午|早上|凌晨|清晨|中午|午间|am/gi, '');
  // 中文时分："8时00分" / "8时" / "8点30分" -> "800"/"830"
  raw = raw.replace(/点/g, '时')
    .replace(/(\d{1,2})时(\d{1,2})分?/g, (x, h, mi) => h + String(mi).padStart(2, '0'))
    .replace(/(\d{1,2})时/g, (x, h) => h + '00')
    .replace(/分/g, '');
  // 冒号/点分时间 "8:30" / "8。30" -> "830"（统一为无分隔 hhmm）
  raw = raw.replace(/(\d{1,2})[:.。：](\d{2})/g, (x, h, mi) => h + mi);
  if (!raw) return null;
  const mt = raw.match(/^(\d{1,4})[-](\d{1,4})$/);
  if (!mt) return null;
  const fmt = (tok) => {
    let h, mi;
    if (tok.length <= 2) { h = tok; mi = '00'; }
    else if (tok.length === 3) { h = tok.slice(0, 1); mi = tok.slice(1); }
    else { h = tok.slice(0, 2); mi = tok.slice(2); }
    h = parseInt(h, 10); mi = parseInt(mi, 10);
    if (pmShift && h >= 1 && h <= 11) h += 12;   // 12小时制：下午1-11点 +12（12点不变）
    if (!(h >= 0 && h < 24 && mi >= 0 && mi < 60)) return null;
    return String(h).padStart(2, '0') + ':' + String(mi).padStart(2, '0');
  };
  const a = fmt(mt[1]), b = fmt(mt[2]);
  if (!a || !b) return null;
  return a + '–' + b;
}

// 单一时间（无范围）："800" / "8:00" / "14:30" / "下午2点" -> "08:00"/"14:00"（结束时间未知）
function fmtSingleTime(s) {
  if (s == null) return null;
  const raw = fullWidthToAscii(s).trim();
  if (!raw) return null;
  if (raw.includes('-')) return fmtTimeRange(raw);   // 已是范围则直接走范围解析
  // 借范围解析器校验小时/分钟合法性（补哨兵结束时刻 2359，随后取开始段）
  const r = fmtTimeRange(raw + '-2359');
  if (!r) return null;
  return r.slice(0, 5);
}

// 由时间范围与时长（分钟）推导结束时间（冗余：时长列缺失时也能显示时长）
function minutesFromRange(rangeStr) {
  const m = String(rangeStr).match(/(\d{2}):(\d{2})–(\d{2}):(\d{2})/);
  if (!m) return null;
  return (+m[3] * 60 + +m[4]) - (+m[1] * 60 + +m[2]);
}

/* ---------- 日期解析（核心难点） ---------- */
/* 监考表常用数字日期 11.9 / 11.1 / 11.11。
 * 致命坑：Excel 把 "11.10" 存成数字 11.1（尾零丢失），数值上无法区分 11月1日/10日。
 * 解决：为每个日期值生成候选集（字面义 + 尾零恢复义），再按"表中日期先后顺序"
 *       贪心修正——考试表总是按时间先后排列的，顺序矛盾时启用候选恢复。
 * 另支持：字符串 11月10日 / 2026.11.10 / 11-10 / 11/10、Excel 日期序列号、Date 对象。
 */

function toSerial(dt) { return Date.UTC(dt.y, dt.m - 1, dt.d) / 86400000; }
function describeDate(dt) { return dt.m + '月' + dt.d + '日'; }

// 生成候选（不含年份，之后统一补）：[{m, d, literal, y?}]，按可信度排序
function candsFromParts(a, b, baseMonth, allowZeroRecover) {
  const out = [];
  const push = (m, d, literal) => {
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31 && !out.some((c) => c.m === m && c.d === d)) {
      out.push({ m: m, d: d, literal: literal });
    }
  };
  push(a, b, true);                                   // 字面 M.D
  if (baseMonth != null && a !== b && b === baseMonth && a !== baseMonth) {
    push(b, a, true);                                 // 交换 D.M 后命中基准月（如 "9.11"=11月9日）
  }
  if (allowZeroRecover && b >= 1 && b <= 3) {
    push(a, b * 10, false);                           // 尾零丢失恢复：11.1 → 可能是 11.10
  }
  if (a > 12 && b <= 12 && out.length === 0) {
    push(b, a, true);                                 // 首数>12 必为 D.M
  }
  if (!out.length) return null;
  // 排序：基准月优先 > 字面义优先（稳定）
  out.sort((x, y) => {
    const bx = x.m === baseMonth ? 0 : 1, by = y.m === baseMonth ? 0 : 1;
    if (bx !== by) return bx - by;
    if (x.literal !== y.literal) return x.literal ? -1 : 1;
    return 0;
  });
  return out;
}

// 单元格值 -> 候集数组；无法识别返回 null
function rawDateCands(v, baseMonth) {
  if (v == null) return null;
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return null;
    return [{ m: v.getMonth() + 1, d: v.getDate(), y: v.getFullYear(), literal: true }];
  }
  if (typeof v === 'number') {
    if (!isFinite(v)) return null;
    if (Number.isInteger(v)) {
      if (v >= 20000 && v <= 80000) {                 // Excel 日期序列号
        const dt = new Date(Math.round((v - 25569) * 86400000));
        return [{ m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), y: dt.getUTCFullYear(), literal: true }];
      }
      // MMDD 整数写法：1109 -> 11月9日（月*100+日；月 1-12，日 1-31，且非序列号范围）
      if (v >= 101 && v <= 1231 && baseMonth != null) {
        const mm = Math.floor(v / 100), dd = v % 100;
        if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) {
          return [{ m: mm, d: dd, literal: true }];
        }
      }
      if (baseMonth != null && v >= 1 && v <= 31) {   // 裸日数：借基准月
        return [{ m: baseMonth, d: v, literal: true }];
      }
      return null;
    }
    const s = v.toFixed(6).replace(/0+$/, '').replace(/\.$/, '');
    const mt = s.match(/^(\d{1,2})\.(\d{1,6})$/);
    if (!mt) return null;
    return candsFromParts(+mt[1], +mt[2], baseMonth, true);  // 浮点必查尾零
  }
  if (typeof v === 'string') {
    let s = fullWidthToAscii(v).trim();
    if (!s) return null;
    let mt = s.match(/^(\d{4})[.\/\-年](\d{1,2})[.\/\-月](\d{1,2})日?$/);
    if (mt) return [{ m: +mt[2], d: +mt[3], y: +mt[1], literal: true }];  // 完整年月日
    s = s.replace(/^[（(【\[]+/, '').replace(/[）)】\]]+$/, '');            // 去首尾括号包裹
    s = s.replace(/^[（(]?(星期|周)[一二三四五六日天][）)]?/, '');           // 去星期/周前缀（"周三 11.9"）
    s = s.replace(/^\d{1,2}[:：]\d{2}(?=[\s.\/\-月])/, '');                // 去时间前缀（"8:00 11.9"）
    // 英文月名：Nov 9 / 9-Nov / Nov-09 / November 9, 2026
    const MON = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, sept: 9, oct: 10, nov: 11, dec: 12 };
    const monOf = (w) => { const k = w.toLowerCase(); return MON[k] != null ? MON[k] : MON[k.slice(0, 3)] != null ? MON[k.slice(0, 3)] : null; };
    const me1 = s.match(/^([A-Za-z]{3,9})\s*[-.\/,]?\s*(\d{1,2})(?:st|nd|rd|th)?\s*,?\s*(\d{4})?$/);
    if (me1 && monOf(me1[1])) return [{ m: monOf(me1[1]), d: +me1[2], y: me1[3] ? +me1[3] : undefined, literal: true }];
    const me2 = s.match(/^(\d{1,2})(?:st|nd|rd|th)?\s*[-.\/,]?\s*([A-Za-z]{3,9})\s*,?\s*(\d{4})?$/);
    if (me2 && monOf(me2[2])) return [{ m: monOf(me2[2]), d: +me2[1], y: me2[3] ? +me2[3] : undefined, literal: true }];
    // 中文数字日期：十一月九日 / 11月初三 / 十月十日
    const CN = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9 };
    const cnNum = (str) => {
      str = String(str).replace(/^初/, '');
      if (/^\d+$/.test(str)) return +str;
      if (str === '十') return 10;
      if (str.length === 1 && CN[str] != null) return CN[str];
      let m3 = str.match(/^十([一二三四五六七八九])$/); if (m3) return 10 + CN[m3[1]];
      let m1 = str.match(/^([一二三四五六七八九])十$/); if (m1) return CN[m1[1]] * 10;
      let m2 = str.match(/^([一二三四五六七八九])十([一二三四五六七八九])$/); if (m2) return CN[m2[1]] * 10 + CN[m2[2]];
      return null;
    };
    const cz = s.match(/^([一二三四五六七八九十\d]{1,3})月([一二三四五六七八九十初\d]{1,3})[日号]?$/);
    if (cz) {
      const cm = cnNum(cz[1]), cd = cnNum(cz[2]);
      if (cm != null && cd != null) return [{ m: cm, d: cd, literal: true }];
    }
    s = s.replace(/[月日号]/g, '.').replace(/[.\/\-\s、,]+/g, '.').replace(/^\.+|\.+$/g, '');
    const parts = s.split('.').filter(Boolean).map(Number);
    if (parts.length === 2 && parts.every(Number.isFinite)) {
      // 字符串 "11.10" 保留了尾零（day=10 唯一）；"11.1" 则存在 1/10 两义
      const explicit = parts[1] >= 10 || String(parts[1]).length >= 2;
      return candsFromParts(parts[0], parts[1], baseMonth, !explicit);
    }
    if (parts.length === 1 && baseMonth != null && parts[0] >= 1 && parts[0] <= 31) {
      return [{ m: baseMonth, d: parts[0], literal: true }];
    }
    // "1109" 四位数字符串 = MMDD（有基准月时启用）
    if (parts.length === 1 && /^\d{4}$/.test(s) && baseMonth != null) {
      const mm = +s.slice(0, 2), dd = +s.slice(2);
      if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) return [{ m: mm, d: dd, literal: true }];
    }
    return null;
  }
  return null;
}

function resolveYear(month, baseMonth, inputYear) {
  // 学期年份按"该日期自身的月份"判断（而不是全表基准月）：
  //   8月~12月 -> 当年（秋季学期）；1月~7月 -> 学期跨年，次年。
  // 这对混排学期（12月 + 次年1月）中每个日期都成立。
  if (month >= 8) return inputYear;
  if (month <= 7) return inputYear + 1;
  return inputYear;
}

function withYear(cand, baseMonth, inputYear) {
  return { y: cand.y != null ? cand.y : resolveYear(cand.m, baseMonth, inputYear), m: cand.m, d: cand.d };
}

function fmtDate(y, m, d) {
  return y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');
}

const WEEKDAYS = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'];
function weekdayCN(dateStr) {
  const dt = new Date(dateStr + 'T00:00:00');
  return WEEKDAYS[dt.getDay()];
}

/* ---------- SheetJS 访问辅助 ---------- */
function decode_range(ref) { return XLSX.utils.decode_range(ref); }
function encode_cell(addr) { return XLSX.utils.encode_cell(addr); }

function buildMergedMap(ws) {
  const map = new Map();
  const merges = ws['!merges'] || [];
  for (const m of merges) {
    const tlAddr = XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c });
    const tlVal = ws[tlAddr] ? ws[tlAddr].v : undefined;
    for (let r = m.s.r; r <= m.e.r; r++) {
      for (let c = m.s.c; c <= m.e.c; c++) {
        map.set(XLSX.utils.encode_cell({ r: r, c: c }), tlVal);
      }
    }
  }
  return map;
}

/* ---------- 主流程：解析工作簿 ---------- */

function parseWorkbook(wb, options) {
  options = options || {};
  const inputYear = options.year || new Date().getFullYear();
  const warnings = [];

  // 1. 选数据表：优先命中表头关键词的表，否则取行数最多的表
  let sheetName = null, ws = null, bestRows = -1;
  for (const name of wb.SheetNames) {
    const w = wb.Sheets[name];
    if (!w || !w['!ref']) continue;
    const ref = decode_range(w['!ref']);
    let headerHit = false;
    for (let R = ref.s.r; R <= Math.min(ref.e.r, ref.s.r + 9) && !headerHit; R++) {
      for (let C = ref.s.c; C <= ref.e.c; C++) {
        const cell = w[XLSX.utils.encode_cell({ r: R, c: C })];
        if (cell && typeof cell.v === 'string' && /课程名称|Subject|监考|考场|学生/.test(cell.v)) { headerHit = true; break; }
      }
    }
    if (headerHit) { sheetName = name; ws = w; break; }
    const nrows = ref.e.r + 1;
    if (nrows > bestRows) { bestRows = nrows; sheetName = name; ws = w; }
  }
  if (!ws) throw new Error('未找到包含数据的工作表（Sheet），请确认提交的是监考表原件。');

  const mergedMap = buildMergedMap(ws);
  const range = decode_range(ws['!ref']);
  const grid = [];
  for (let R = range.s.r; R <= range.e.r; R++) {
    const row = [];
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = encode_cell({ r: R, c: C });
      let v = ws[addr] ? ws[addr].v : undefined;
      if ((v === undefined || v === null) && mergedMap.has(addr)) v = mergedMap.get(addr);
      row.push(v);
    }
    grid.push(row);
  }

  // 2. 表头行：前 30 行中命中关键词最多（>=2）的行（表头可能不在文件顶部，
  //    例如前面有标题行、空行、说明行、多级表头）
  const KEYWORDS = [/日期|Date/i, /时间|Time/i, /课程|科目|Subject/i, /考场|教室|地点|Room|Venue/i, /监考|Proctor/i, /学生|考生|Student/i];
  let headerRow = -1, headerScore = 0;
  for (let i = 0; i < Math.min(30, grid.length); i++) {
    const joined = grid[i].map((c) => String(c == null ? '' : c)).join('|');
    let score = 0;
    for (const kw of KEYWORDS) if (kw.test(joined)) score++;
    if (score >= 2 && score > headerScore) { headerRow = i; headerScore = score; }
  }
  if (headerRow < 0) throw new Error('未找到表头行（需要同时包含"日期、时间、课程"等字样）。请确认提交的是标准监考表。');

  // 3. 列识别（宽松匹配 + 兜底推断）
  const hdr = grid[headerRow];
  const htext = hdr.map((c) => String(c == null ? '' : c));
  const findCol = (res) => {
    for (let c = 0; c < hdr.length; c++) if (res.test(htext[c])) return c;
    return -1;
  };
  const colIdx = {
    date: findCol(/日期|Date/i),
    time: findCol(/时间|Time|Start|Begin|开始/i),
    subject: findCol(/课程|科目|Subject|考试|Exam(?!inee)/i),
    minutes: findCol(/时长|分钟|Minutes|Duration/i),
    grade: findCol(/年级|Grade|Year/i),
    klass: findCol(/班级|Class|Form/i),
    room: findCol(/考场|教室|地点|Room|Venue|Location|Site/i),
    count: findCol(/人数|Numbers|Count|Num(?!ber)/i),
    students: findCol(/学生|考生|Student|Candidate|名单/i),
  };
  // 监考列：匹配"监考/Proctor"的第 1、2 列
  const proctorCols = [];
  for (let c = 0; c < hdr.length; c++) if (/监考|Proctor/i.test(htext[c])) proctorCols.push(c);
  colIdx.proctor1 = proctorCols[0] != null ? proctorCols[0] : -1;
  colIdx.proctor2 = proctorCols[1] != null ? proctorCols[1] : -1;

  const missing = [];
  for (const k of ['date', 'time', 'subject']) if (colIdx[k] < 0) missing.push(k);
  if (missing.length) {
    // 兜底：按标准监考表列序推断（A=日期 B=时间 C=课程）
    if (colIdx.date < 0 && hdr.length > 0) colIdx.date = 0;
    if (colIdx.time < 0 && hdr.length > 1) colIdx.time = 1;
    if (colIdx.subject < 0 && hdr.length > 2) colIdx.subject = 2;
    const still = [];
    for (const k of ['date', 'time', 'subject']) if (colIdx[k] < 0 || colIdx[k] == null) still.push(k);
    if (still.length) throw new Error('表头缺少关键列：' + still.join('、') + '。请确认提交的是标准监考表（需含 日期/时间/课程名称 列）。');
    warnings.push('表头部分列未能按名称识别，已按标准列序（日期/时间/课程）推断解析。');
  }
  if (colIdx.students < 0) {
    // 学生列兜底：人数列之后；再不行取已识别列的最大列号 + 1
    let start = colIdx.count >= 0 ? colIdx.count + 1 : -1;
    if (start < 0) {
      let mx = -1;
      for (const k of Object.keys(colIdx)) if (colIdx[k] > mx) mx = colIdx[k];
      start = mx + 1;
    }
    colIdx.students = start;
    warnings.push('未找到"学生"列表头，已按第 ' + (start + 1) + ' 列起作为学生名单解析。');
  }

  // 4. 第一遍：收集日期单元格 -> 基准月份
  const dateRows = [];   // {row(0-based), raw, key}
  for (let r = headerRow + 1; r < grid.length; r++) {
    const dv = grid[r][colIdx.date];
    if (dv === undefined || dv === null || String(dv).trim() === '') continue;
    dateRows.push({ row: r, raw: dv });
  }
  const monthTally = {};
  for (const dr of dateRows) {
    const cands = rawDateCands(dr.raw, null);
    if (cands) {
      const m = cands[0].m;
      monthTally[m] = (monthTally[m] || 0) + 1;
    } else if (typeof dr.raw === 'number' && Number.isInteger(dr.raw) && dr.raw >= 101 && dr.raw <= 1231) {
      // MMDD 整数（1109=11月9日）：基准月尚未确定也能投票（首位是月）
      const mm = Math.floor(dr.raw / 100), dd = dr.raw % 100;
      if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) monthTally[mm] = (monthTally[mm] || 0) + 1;
    } else if (typeof dr.raw === 'string') {
      // "1109" / "1110" 字符串形式的 MMDD
      const sM = fullWidthToAscii(dr.raw).trim();
      if (/^\d{4}$/.test(sM)) {
        const mm = +sM.slice(0, 2), dd = +sM.slice(2);
        if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) monthTally[mm] = (monthTally[mm] || 0) + 1;
      }
    }
  }
  let baseMonth = null, best = 0;
  for (const m of Object.keys(monthTally)) {
    if (monthTally[m] > best) { best = monthTally[m]; baseMonth = +m; }
  }

  // 5. 日期顺序修正：同一原始值共享同一解析结果；按出现顺序贪心修正
  const distinct = [];   // {key, raw, label, cands(含年份)}
  const byKey = new Map();
  for (const dr of dateRows) {
    const key = (typeof dr.raw === 'number' ? 'n:' + dr.raw :
      dr.raw instanceof Date ? 'd:' + dr.raw.getTime() : 's:' + String(dr.raw).trim());
    if (!byKey.has(key)) {
      const cands = rawDateCands(dr.raw, baseMonth);
      if (!cands) {
        warnings.push('第 ' + (dr.row + 1) + ' 行：无法识别日期 "' + String(dr.raw).trim() + '"，该行将被跳过。');
        byKey.set(key, null);
      } else {
        const full = cands.map((c) => withYear(c, baseMonth, inputYear));
        const label = (typeof dr.raw === 'number' ? String(dr.raw) :
          dr.raw instanceof Date ? fmtDate(dr.raw.getFullYear(), dr.raw.getMonth() + 1, dr.raw.getDate()) : String(dr.raw).trim());
        const entry = { key: key, raw: dr.raw, label: label, cands: full, chosen: full[0] };
        byKey.set(key, entry);
        distinct.push(entry);
      }
    }
    dr.key = key;
  }
  const repairs = [];
  let prevSerial = null, prevDesc = null;
  // 段末上界：全表其他日期的最大 serial。顺序修正的候选不得超过它（+跨年容差 45 天），
  // 防止把 "3.1" 之类改到远超段末的 3.10（该情形应由连贯性分析按"保留原值+警告"处理）
  let boundSerial = -1;
  for (const entry of distinct) {
    if (!entry) continue;
    for (const c of entry.cands) {
      const s = toSerial(c);
      if (s != null && s > boundSerial) boundSerial = s;
    }
  }
  for (const entry of distinct) {
    if (!entry) continue;
    // 跨年容忍：若与前序日期差接近一年（学期跨年，如 12月 -> 1月），视为正常推进，不触发顺序修正
    let pick = entry.cands[0];
    if (prevSerial != null && toSerial(pick) < prevSerial) {
      const gap = prevSerial - toSerial(pick);
      if (gap >= 300 && gap <= 370) {
        // 跨年（前序 12月末 -> 当前 1月）：保持字面值，年份由 withYear 已按"次月<=2 => 次年"处理
        entry.chosen = pick;
        prevSerial = toSerial(pick);
        prevDesc = describeDate(pick);
        continue;
      }
      const alt = entry.cands.find((c) => {
        const s = toSerial(c);
        return s >= prevSerial && s <= boundSerial + 45;   // 上界保护：不越过全表最大日期（容差 45 天）
      });
      if (alt) {
        entry.dateRepaired = { from: describeDate(entry.cands[0]) };
        repairs.push({
          raw: entry.label,
          from: describeDate(entry.cands[0]),
          to: describeDate(alt),
          reason: '位于 ' + prevDesc + ' 之后，按表中日期先后顺序自动修正',
        });
        pick = alt;
      } else {
        warnings.push('日期 "' + entry.label + '"（' + describeDate(entry.cands[0]) + '）与前面日期顺序矛盾且无法自动修正，请人工核对原表。');
      }
    }
    entry.chosen = pick;
    prevSerial = toSerial(pick);
    prevDesc = describeDate(pick);
  }

  /* 5.5 考试日期连贯性分析（v4）：考试日期通常连续几天（如 11.9/11.10/11.11）。
   * 若某日期脱离连续段，且符合 Excel 吞尾零特征（11.1 的 ×10=11.10 恰落在连续段范围内），
   * 判定为录入错误并自动修正；无法确定目标时保留原值并警告，绝不猜测。
   * 对任意月份通用（3月、5月……），跨年连续段（12月→1月）同样支持。
   * 链接判定支持"隔日桥"：11.9 与 11.11（中间缺 11.10）也视为互相衔接。 */
  if (distinct.length >= 2) {
    const dayGap = (a, b) => {          // a、b 为 serial；返回是否可视为"连续/近邻"
      const d = Math.abs(a - b);
      return d === 1 || d === 2 || (d >= 300 && d <= 370);
    };
    const chains = [];
    {
      const items = distinct
        .map((en) => ({ en: en, s: toSerial(en.chosen) }))
        .filter((x) => x.s != null)
        .sort((a, b) => a.s - b.s);
      let cur = [];
      for (const it of items) {
        if (cur.length && !dayGap(it.s, cur[cur.length - 1].s)) { chains.push(cur); cur = []; }
        cur.push(it);
      }
      if (cur.length) chains.push(cur);
    }
    const longest = chains.reduce((a, b) => (b.length > a.length ? b : a), []);
    const active = longest.length >= 2 ? longest : null;   // 门槛：至少 2 个互相连续(含隔日桥)的日期才启用
    if (active) {
      const inChain = new Set(longest.map((x) => x.s));
      const minS = longest[0].s, maxS = longest[longest.length - 1].s;
      const spanDesc = describeDate(longest[0].en.chosen) + ' 至 ' + describeDate(longest[longest.length - 1].en.chosen);
      // 脱离连续段的异常项：不在最长链内，且与任何其他日期都不近邻
      const outliers = distinct.filter((en) => {
        const s = toSerial(en.chosen);
        if (s == null || inChain.has(s)) return false;
        return !distinct.some((o) => {
          if (o === en) return false;
          const os = toSerial(o.chosen);
          return os != null && dayGap(s, os);
        });
      });
      // ×10 目标日认领：多个异常项映射到同一天视为歧义，全部不修（宁可不改也不错改）
      const claims = new Map();   // targetSerial -> [{en, tgt, origD}]
      for (const en of outliers) {
        const c = en.chosen;
        const dd = c.d * 10;
        if (dd > 31) continue;   // 11.4×10=40 不是合法日，无尾零嫌疑
        const tgt = { y: c.y, m: c.m, d: dd };
        const ts = toSerial(tgt);
        if (ts == null || ts < minS || ts > maxS || c.m !== longest[0].en.chosen.m) continue;
        if (!claims.has(ts)) claims.set(ts, []);
        claims.get(ts).push({ en: en, tgt: tgt, origD: c.d });
      }
      const claimed = new Set();
      for (const [ts, list] of claims) {
        if (list.length !== 1) {
          for (const it of list) {
            claimed.add(it.en);
            warnings.push('日期 "' + it.en.label + '"（' + describeDate(it.en.chosen) + '）脱离考试日期连续段（' + spanDesc + '），且修正目标存在歧义，已保留原值，请人工核对。');
          }
          continue;
        }
        const it = list[0];
        const fromDesc = describeDate(it.en.chosen);
        const toDesc = describeDate(it.tgt);
        it.en.cands = [it.tgt].concat(it.en.cands);   // 修正候选提到最前
        it.en.chosen = it.tgt;
        it.en.dateRepaired = { from: fromDesc };
        claimed.add(it.en);
        repairs.push({
          raw: it.en.label,
          from: fromDesc,
          to: toDesc,
          reason: '该日期脱离考试日期连续段（' + spanDesc + '），且符合 Excel 日期尾零丢失特征（' + it.tgt.m + '月' + it.origD + '日 → ' + it.tgt.m + '月' + it.tgt.d + '日）',
        });
      }
      // 回滚裁决：异常项若曾被"顺序修正"改过（dateRepaired 已有值）但 ×10 目标不在段内缺失位，
      // 说明顺序修正改错了方向（如 3.1→3.10 越过段末）——回滚到原始字面值并警告，宁可不改也不错改
      for (const en of outliers) {
        if (claimed.has(en)) continue;
        if (en.chosen && en.cands.length && en.cands[0] !== en.chosen) {
          // cands[0] 是原始字面值解析；若与 chosen 不同，说明被顺序/其他修正改过
          const orig = en.cands[0];
          const rollbackOk = !claims.size || true;   // 只要 ×10 无有效认领即回滚到字面值
          if (rollbackOk && toSerial(orig) != null) {
            repairs.splice(repairs.findIndex((rp) => rp.raw === en.label && rp.to === describeDate(en.chosen)), 1);
            warnings.push('日期 "' + en.label + '"（' + describeDate(orig) + '）脱离考试日期连续段（' + spanDesc + '），无法确定正确日期，已保留原值，请人工核对。');
            en.chosen = orig;
            en.dateRepaired = null;
            continue;
          }
        }
        warnings.push('日期 "' + en.label + '"（' + describeDate(en.chosen) + '）脱离考试日期连续段（' + spanDesc + '），无法确定正确日期，已保留原值，请人工核对。');
      }
    }
  }

  // 6. 第二遍：构建考试记录
  const exams = [];
  const seenKeys = new Set();
  for (let r = headerRow + 1; r < grid.length; r++) {
    const row = grid[r];
    if (isEmptyDataRow(row)) continue;

    // 跳过正文中重复出现的表头行
    const dvRaw = row[colIdx.date];
    if (typeof dvRaw === 'string' && /日期|Date/i.test(dvRaw) &&
      /时间|Time/i.test(String(row[colIdx.time] == null ? '' : row[colIdx.time]))) continue;

    // 学生名单（学生列起，直到表头词为止）；同格多名单（顿号/逗号/分号/空格/换行分隔）自动拆分
    const students = [];
    for (let c = colIdx.students; c < row.length; c++) {
      const v = row[c];
      if (v === undefined || v === null) continue;
      const s = String(v).trim();
      if (!s) continue;
      if (/^(学生|考生|Students?)$/i.test(s)) break;
      // 同一格多个姓名：中文顿号/逗号/分号、英文逗号/分号、换行拆分；
      // 空格仅当格内含 2+ 个 2-4 字中文段时才拆（避免英文名 "John Smith" 被误拆）
      const pieces = s.split(/[、，；;,\n\r]+/)
        .map((p) => p.trim())
        .filter(Boolean)
        .flatMap((p) => {
          if (/^[\u4e00-\u9fa5·\s]+$/.test(p) && /[\u4e00-\u9fa5]\s+[\u4e00-\u9fa5]/.test(p)) {
            const segs = p.split(/\s+/).filter((x) => /^[\u4e00-\u9fa5·]{2,4}$/.test(x));
            if (segs.length >= 2) return segs;   // "张三 李四 王五" 视为多人
          }
          return [p];
        });
      students.push(...pieces);
    }
    if (!students.length) continue;   // 无学生的行（备注、空行）静默跳过

    const entry = byKey.get(typeof dvRaw === 'number' ? 'n:' + dvRaw :
      dvRaw instanceof Date ? 'd:' + dvRaw.getTime() :
        's:' + (dvRaw == null ? '' : String(dvRaw).trim()));
    if (!entry || !entry.chosen) continue;   // 已在第一遍告警
    const dateStr = fmtDate(entry.chosen.y, entry.chosen.m, entry.chosen.d);

    let timeStr = fmtTimeRange(row[colIdx.time]);
    if (!timeStr) {
      const single = fmtSingleTime(row[colIdx.time]);   // 冗余：单一 "800" 也识别
      if (single) {
        timeStr = single;
        warnings.push('第 ' + (r + 1) + ' 行：时间 "' + String(row[colIdx.time]).trim() + '" 只有开始时刻，无结束时间。');
      }
    }
    if (!timeStr) {
      warnings.push('第 ' + (r + 1) + ' 行：无法识别时间 "' + String(row[colIdx.time] == null ? '' : row[colIdx.time]).trim() + '"（科目：' + String(row[colIdx.subject] == null ? '?' : row[colIdx.subject]).trim() + '）');
      timeStr = String(row[colIdx.time] == null ? '' : row[colIdx.time]).trim() || '（未标注）';
    }

    let minutes = null;
    const mv = row[colIdx.minutes];
    if (mv != null && String(mv).trim() !== '' && Number(mv) > 0) minutes = Number(mv);
    if (minutes == null) minutes = minutesFromRange(timeStr);   // 冗余：由时间范围推导

    const subject = String(row[colIdx.subject] == null ? '（未命名科目）' : row[colIdx.subject]).trim() || '（未命名科目）';
    let room = row[colIdx.room] != null ? String(row[colIdx.room]).trim() : '';
    // 冗余：考场列缺失时，从科目文本提取括号内的考场号（如 "数学(3103)" -> room=3103）
    let subjClean = subject;
    if (!room) {
      const rm = subject.match(/[（(]([A-Za-z0-9\-—－]{2,12})[）)]\s*$/);
      if (rm) { room = rm[1]; subjClean = subject.slice(0, rm.index).trim(); }
    }
    const grade = row[colIdx.grade] != null ? String(row[colIdx.grade]).trim() : '';
    const klass = row[colIdx.klass] != null ? String(row[colIdx.klass]).trim() : '';

    // 冗余去重：完全相同的行（含相同学名单）只保留一条
    const sig = [dateStr, timeStr, subject, grade, klass, room, students.slice().sort().join('§')].join('¦');
    if (seenKeys.has(sig)) continue;
    seenKeys.add(sig);

    exams.push({
      date: dateStr,
      weekday: weekdayCN(dateStr),
      time: timeStr,
      minutes: minutes,
      subject: subjClean || subject,
      grade: grade,
      klass: klass,
      room: room || '（未标注）',
      students: students,
      repairedFrom: entry.dateRepaired ? entry.dateRepaired.from : null,
      _row: r + 1,
    });
  }

  if (!exams.length) throw new Error('解析完成但没有找到任何有效的考试记录。请确认提交的是最新版监考表（需包含学生名单）。');

  exams.sort((a, b) => (a.date + ' ' + a.time).localeCompare(b.date + ' ' + b.time));

  // 汇总去重后的考试日期（供 UI 展示核对）
  const dateList = [];
  for (const ex of exams) {
    if (!dateList.some((x) => x.date === ex.date)) {
      dateList.push({ date: ex.date, label: describeDate(entryOf(ex.date)) , weekday: ex.weekday });
    }
  }
  function entryOf(dateStr) {
    const p = dateStr.split('-');
    return { y: +p[0], m: +p[1], d: +p[2] };
  }

  return {
    exams: exams,
    meta: {
      sheetName: sheetName,
      baseMonth: baseMonth,
      inputYear: inputYear,
      totalRows: exams.length,
      warnings: warnings,
      repairs: repairs,
      dateList: dateList,
    },
  };
}

/* ---------- 按姓名查找 ---------- */

function findExamsForStudent(parsed, rawName) {
  const name = normalizeName(rawName);
  if (!name) return { status: 'empty' };

  const hits = [];
  for (const ex of parsed.exams) {
    for (const st of ex.students) {
      if (normalizeName(st) === name) { hits.push({ exam: ex, matchType: 'exact' }); break; }
    }
  }
  if (hits.length) {
    hits.sort((a, b) => (a.exam.date + ' ' + a.exam.time).localeCompare(b.exam.date + ' ' + b.exam.time));
    // 同名检测（两条可靠信号，避免合考场误报）：
    //  信号A：仅统计"非合考场"行（年级无逗号、班级为单一班），班级基础名（去括号后缀）不同 => 真同名
    //  信号B：同一天两场时间区间重叠 => 单人无法同时参加，必为同名两人或数据错误
    const classBase = (k) => String(k || '').replace(/[（(【\[][^）)】\]]*[）)】\]]/g, '').replace(/\s+/g, '');
    const isSingleClassRow = (ex) => ex && !(String(ex.grade || '').includes(',')) && !(String(ex.klass || '').includes(','));
    const bases = new Set();
    for (const h of hits) {
      if (isSingleClassRow(h.exam)) bases.add((h.exam.grade || '') + '|' + classBase(h.exam.klass));
    }
    let timeOverlap = false;
    const toMin = (t) => { const m = String(t || '').match(/(\d{1,2}):(\d{2})/g); return m && m.length >= 1 ? (+m[0].slice(0, 2)) * 60 + (+m[0].slice(3)) : null; };
    const endMin = (t) => { const m = String(t || '').match(/(\d{1,2}):(\d{2})/g); return m && m.length >= 2 ? (+m[1].slice(0, 2)) * 60 + (+m[1].slice(3)) : null; };
    for (let i = 0; i < hits.length && !timeOverlap; i++) {
      for (let j = i + 1; j < hits.length; j++) {
        const a = hits[i].exam, b = hits[j].exam;
        if (a.date !== b.date) continue;
        const as = toMin(a.time), ae = endMin(a.time), bs = toMin(b.time), be = endMin(b.time);
        if (as == null || bs == null) continue;
        const aEnd = ae == null ? as : ae, bEnd = be == null ? bs : be;   // 无结束时间按开始时刻判
        if (as < bEnd && bs < aEnd) { timeOverlap = true; break; }
      }
    }
    // 展示用组合列表（含合考场，仅供 UI 参考）
    const combos = new Map();
    for (const ex of parsed.exams) {
      const has = ex.students.some((st) => normalizeName(st) === name);
      if (!has) continue;
      const key = (ex.grade || '') + '|' + (ex.klass || '');
      if (!combos.has(key)) combos.set(key, { grade: ex.grade || '', klass: ex.klass || '', count: 0 });
      combos.get(key).count++;
    }
    const comboList = Array.from(combos.values());
    return {
      status: 'found',
      hits: hits,
      ambiguous: bases.size > 1 || timeOverlap,
      ambiguityReason: bases.size > 1 ? '同名同学出现在不同班级' : (timeOverlap ? '存在时间重叠的两场考试' : null),
      combos: comboList,
    };
  }

  // 部分匹配（含大小写/全半角容错）：输入为名单子串，或名单为输入子串
  const partial = [];
  const seen = new Set();
  for (const ex of parsed.exams) {
    for (const st of ex.students) {
      const ns = normalizeName(st);
      if (!ns) continue;
      if (ns.includes(name) || name.includes(ns)) {
        if (!seen.has(ns)) { seen.add(ns); partial.push({ student: st, norm: ns, examCount: 0 }); }
      }
    }
  }
  for (const p of partial) {
    p.examCount = parsed.exams.filter((ex) => ex.students.some((st) => normalizeName(st) === p.norm)).length;
  }
  partial.sort((a, b) => a.norm.localeCompare(b.norm));
  if (partial.length) return { status: 'partial', candidates: partial };

  return { status: 'not-found' };
}

/* ---------- 导出 ---------- */
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { parseWorkbook, findExamsForStudent, normalizeName, fmtTimeRange, fmtSingleTime, rawDateCands, weekdayCN };
} else if (typeof window !== 'undefined') {
  window.ExamCore = { parseWorkbook, findExamsForStudent, normalizeName, fmtTimeRange, fmtSingleTime, rawDateCands, weekdayCN };
}
