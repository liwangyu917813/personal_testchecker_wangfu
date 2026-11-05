/* =====================================================================
 * 考试查询核心逻辑 (core.js)  v5 —— 流水线架构版
 * 职责：解析监考表 XLSX -> 结构化考试数组 -> 按姓名查找
 * 设计原则：任何异常都给出明确中文提示，绝不静默出错；
 *           对日期、时间、表头、姓名均做多重冗余识别。
 *
 * 架构（v5 重构）：parseWorkbook 是一条 7 阶段流水线，
 *   _P1_selectSheet   选数据表 + 展开合并单元格 -> grid
 *   _P2_findHeader    表头行定位（前 30 行关键词打分）
 *   _P3_mapColumns    列识别（宽松匹配 + 标准列序兜底）
 *   _P4_collectDates  日期第一遍收集 + 投票定基准月
 *   _P5_repairDates   日期同值共享 + 出现顺序贪心修正
 *   _P55_coherence    连贯性分析（×10 尾零认领 + 歧义回滚）
 *   _P6_buildRecords  记录构建（名单拆分/时间/分钟/考场/去重）
 *   _P7_finalize      排序 + 汇总 + 组装返回值
 * 阶段间通过 ctx 对象传递：数组/对象字段共享引用，标量经 Object.assign 写回。
 * 每个阶段职责单一，可独立测试；阶段函数以 _P 前缀标识内部实现。
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

/** 全角字符转半角（数字、字母、冒号、波浪线等），提高各种输入法的容错 */
function fullWidthToAscii(s) {
  return String(s).replace(/[\uFF01-\uFF5E\u3000]/g, (ch) =>
    ch === '\u3000' ? ' ' : String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/[\uFF0D\u2010\u2011\u2012\u2013\u2014\u2015\u2212\uFF0F]/g, '-'); // 全角横线/各类破折号/全角斜杠 -> '-'
}

/** 判断 grid 一行是否为空行（前 11 列全空）——用于静默跳过，不产生警告 */
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
/** 日期对象 -> 中文短描述（"11月10日"，警告与修正记录用） */
function describeDate(dt) { return dt.m + '月' + dt.d + '日'; }

// 生成候选（不含年份，之后统一补）：[{m, d, literal, y?}]，按可信度排序
function candsFromParts(a, b, baseMonth, allowZeroRecover) {
  // 冗余校验：字面月/日越界（m>12 或 m<1 或 d>31 或 d<1）直接判为无法识别的日期。
  // 之前 "13.5" 会被静默交换成 5月13日——这对监考表是危险的猜测（真实含义更可能是录错的月份），
  // 现在返回 null → 上层统一给"无法识别日期，该行将被跳过"警告。
  if (a < 1 || a > 12 || b < 1 || b > 31) return null;
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
  if (baseMonth != null && a > 12 && b >= 1 && b <= 12) {
    push(b, a, true);                                 // 首数>12：必为 D.M（"13.5"=5月13日），即使无基准月也给出候选并标记越界月
  }
  if (allowZeroRecover && b >= 1 && b <= 3) {
    push(a, b * 10, false);                           // 尾零丢失恢复：11.1 → 可能是 11.10
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

/**
 * 跨年学年推断：8 月及以后属 inputYear 年（上学期），7 月及以前属 inputYear+1（下学期）。
 * 例：inputYear=2026 时 12.20 → 2026-12-20，1.5 → 2027-01-05。
 */
function resolveYear(month, baseMonth, inputYear) {
  // 学期年份按"该日期自身的月份"判断（而不是全表基准月）：
  //   8月~12月 -> 当年（秋季学期）；1月~7月 -> 学期跨年，次年。
  // 这对混排学期（12月 + 次年1月）中每个日期都成立。
  if (month >= 8) return inputYear;
  if (month <= 7) return inputYear + 1;
  return inputYear;
}

/** 给不含年份的候选补年份（跨年规则见 resolveYear；已带 y 的候选原样保留） */
/** 数字年月日 -> "M月D日"（警告文案用） */
function label0(m, d) { return m + '.' + d; }

function withYear(cand, baseMonth, inputYear) {
  // 冗余校验：字面日期必须是真实存在的日历日（2.30/4.31 这类不存在）。
  // 解析内部序列号计算不会崩溃，但输出的 "2026-02-30" 是幻觉日期，标记供警告。
  const DIM = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  const maxD = cand.m === 2 ? 29 : DIM[cand.m - 1];
  return {
    y: cand.y != null ? cand.y : resolveYear(cand.m, baseMonth, inputYear),
    m: cand.m, d: cand.d,
    invalidDay: cand.d > maxD,
  };
}

//** 数字年月日 -> "YYYY-MM-DD"（补零） */
function fmtDate(y, m, d) {
  return y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');
}

const WEEKDAYS = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'];
/** "YYYY-MM-DD" -> "周X"（本地时区安全解析） */
function weekdayCN(dateStr) {
  const dt = new Date(dateStr + 'T00:00:00');
  return WEEKDAYS[dt.getDay()];
}

/* ---------- SheetJS 访问辅助 ---------- */
function decode_range(ref) { return XLSX.utils.decode_range(ref); }
function encode_cell(addr) { return XLSX.utils.encode_cell(addr); }

/**
 * 展开工作表合并单元格：返回 Map<"r,c" 地址, 左上角值>。
 * 合并区域内的每个格子都能取到日期/年级等合并值（监考表日期列大量使用合并）。
 */
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

/* ---------- 主流程：解析工作簿（7 阶段流水线） ----------
 * 阶段间数据通过 ctx 传递；数组/对象字段为共享引用，标量字段经 Object.assign 写回。
 * ctx 字段清单：
 *   输入：wb, options, inputYear
 *   累积：warnings, repairs
 *   P1 产物：sheetName, ws, mergedMap, grid
 *   P2/P3 产物：headerRow, colIdx
 *   P4 产物：dateRows, baseMonth
 *   P5 产物：distinct, byKey（repairs 同累积）
 *   P6 产物：exams, seenKeys
 *   P7 产物：dateList（并返回 meta 对象） */
function parseWorkbook(wb, options) {
  // 输入防御：非对象/null/无 SheetNames 的输入给出中文错误，而不是裸的内部异常
  if (!wb || typeof wb !== 'object' || !Array.isArray(wb.SheetNames) || !wb.Sheets || typeof wb.Sheets !== 'object') {
    throw new Error('传入的数据不是有效的 Excel 工作簿（缺少工作表），请重新选择正确的 .xlsx 文件。');
  }
  const ctx = {
    wb: wb,
    options: options || {},
    inputYear: (options && options.year) || new Date().getFullYear(),
    warnings: [],
    repairs: [],
    exams: [],
    seenKeys: new Set(),
    byKey: new Map(),
    distinct: [],
    dateRows: [],
    colIdx: null,
    headerRow: -1,
    baseMonth: null,
    sheetName: null,
    ws: null,
    grid: null,
    mergedMap: null,
    dateList: [],
  };
  _P1_selectSheet(ctx);
  _P2_findHeader(ctx);
  _P3_mapColumns(ctx);
  _P4_collectDates(ctx);
  _P5_repairDates(ctx);
  _P55_coherence(ctx);
  _P6_buildRecords(ctx);
  const meta = _P7_finalize(ctx);
  _verifyParse(ctx, meta);          // 解析后自检（双保险）：异常结果给出明确警告
  return meta;
}

/* ---------- 解析后自检（双保险冗余） ---------- */
/**
 * 流水线结束后的健全性检查：不改变解析结果，只对"可疑但未触发异常"的输出
 * 追加警告，让 UI 有机会提示用户核对——宁可多提醒，不可静默出错。
 * 检查项：
 *   1. 有 grid 有表头却 0 场考试 → 疑似学生列识别错误
 *   2. 全部考试都没有学生 → 疑似名单列错位
 *   3. 警告数超过考试数的一半 → 数据质量差，提示人工核对
 *   4. 单场考试学生数 > 200 → 疑似名单列拼接异常
 * @param {object} ctx   流水线上下文
 * @param {{exams:Array,meta:object}} meta  _P7_finalize 返回值
 */
function _verifyParse(ctx, meta) {
  try {
    const exams = meta.exams, grid = ctx.grid, headerRow = ctx.headerRow, warnings = meta.meta.warnings;
    // 1) 找到了表头和网格，却没有任何有效考试行 → 最可能是学生列错位或名单全空
    if (exams.length === 0 && grid && headerRow >= 0 && grid.length - headerRow > 3) {
      warnings.push('自检：找到了表头（第 ' + (headerRow + 1) + ' 行）但没解析出任何考试记录，可能是学生名单列不在此表中，请人工核对列位置。');
      return;
    }
    if (exams.length === 0) return;   // 其余 0 场场景已由主流程抛错/警告，不重复
    // 2) 抽查：统计"没有任何学生"的占比（正常监考表每场都应有学生）
    let noStudents = 0;
    for (const ex of exams) if (!ex.students || !ex.students.length) noStudents++;
    if (noStudents === exams.length) {
      warnings.push('自检：所有 ' + exams.length + ' 场考试都没有学生名单，请确认学生列位置是否正确。');
    }
    // 3) 警告数量异常多 → 提示整体数据质量
    if (warnings.length > Math.max(20, exams.length / 2)) {
      warnings.push('自检：数据警告多达 ' + warnings.length + ' 条，表格格式可能差异较大，建议对照原件人工核对。');
    }
    // 4) 单场学生数异常（>200 视为名单被挤成一格）
    const big = exams.filter((ex) => ex.students && ex.students.length > 200);
    if (big.length) {
      warnings.push('自检：有 ' + big.length + ' 场考试的学生数超过 200 人（如 "' +
        String(big[0].subject) + '" 有 ' + big[0].students.length + ' 人），疑似多名学生的名单被挤在同一格且未按预期拆分，请核对。');
    }
  } catch (e) {
    // 自检本身绝不能影响解析结果
    try { ctx.warnings.push('自检过程出错（不影响解析结果）：' + e.message); } catch (e2) { /* 忽略 */ }
  }
}

/* P1 · 选数据表：优先命中表头关键词的表，否则取行数最多的表；随后展开合并单元格构建 grid */
function _P1_selectSheet(ctx) {
    const wb = ctx.wb;
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


    Object.assign(ctx, { sheetName, ws, mergedMap, grid });
}
/* P2 · 表头行定位：前 30 行中命中关键词最多（>=2）的行 */
function _P2_findHeader(ctx) {
    const grid = ctx.grid;
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


    Object.assign(ctx, { headerRow });
}
/* P3 · 列识别：宽松匹配 + 标准列序兜底 + 学生列位置推断 */
function _P3_mapColumns(ctx) {
    const grid = ctx.grid, headerRow = ctx.headerRow, warnings = ctx.warnings;
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


    Object.assign(ctx, { colIdx });
}
/* P4 · 日期第一遍：收集日期单元格 → 投票定基准月 baseMonth */
function _P4_collectDates(ctx) {
    const grid = ctx.grid, headerRow = ctx.headerRow, colIdx = ctx.colIdx, warnings = ctx.warnings;
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
  // 双保险兜底：投票失败（全表日期都无法识别月份）时，用第一行的时间/科目旁证推断不可行，
  // 则取第一个能解析出候选的日期的月份作为基准月——"全表大概率同月"是监考表的普遍形态。
  if (baseMonth == null) {
    for (const dr of dateRows) {
      const cands = rawDateCands(dr.raw, null);
      if (cands && cands[0] && cands[0].m) { baseMonth = cands[0].m; break; }
      // MMDD 整数兜底
      if (typeof dr.raw === 'number' && Number.isInteger(dr.raw) && dr.raw >= 101 && dr.raw <= 1231) {
        const mm = Math.floor(dr.raw / 100), dd = dr.raw % 100;
        if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) { baseMonth = mm; break; }
      }
    }
    if (baseMonth != null) warnings.push('日期月份投票未能确定基准月，已按首个可识别日期的月份（' + baseMonth + ' 月）兜底解析。');
  }

  Object.assign(ctx, { dateRows, baseMonth });
}
/* P5 · 日期顺序修正：同值共享解析 + 按出现顺序贪心修正（含段末上界与跨年容忍） */
function _P5_repairDates(ctx) {
    const dateRows = ctx.dateRows, baseMonth = ctx.baseMonth, inputYear = ctx.inputYear, warnings = ctx.warnings;
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
        // 冗余校验：字面日期必须是真实存在的日历日（2.30/4.31/11.31 这类不存在）
        // 解析器内部用"日 序列号"计算不会崩溃，但对外输出的 "2026-11-31" 是幻觉日期，
        // 必须警告提示人工核对（保留原值，不擅自改动）
        const full = cands.map((c) => withYear(c, baseMonth, inputYear));
        const first = full[0];
        if (first.invalidDay) {
          const DIM = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
          const maxD = first.m === 2 ? 29 : DIM[first.m - 1];
          warnings.push('第 ' + (dr.row + 1) + ' 行：日期 "' + label0(first.m, first.d) + '" 不存在（' + first.m + '月只有 ' + maxD + ' 天），请人工核对原表。');
        }
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


    Object.assign(ctx, { distinct, byKey, repairs });
}
/* P5.5 · 连贯性分析：考试日期连续段 + Excel 尾零丢失(×10)认领 + 歧义回滚 */
function _P55_coherence(ctx) {
    const distinct = ctx.distinct, repairs = ctx.repairs, warnings = ctx.warnings;
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


    Object.assign(ctx, { distinct, repairs });
}
/* P6 · 记录构建：名单拆分/时间解析/分钟推导/考场提取/去重 */
function _P6_buildRecords(ctx) {
    const grid = ctx.grid, headerRow = ctx.headerRow, colIdx = ctx.colIdx, byKey = ctx.byKey, distinct = ctx.distinct, warnings = ctx.warnings;
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

    // 监考老师提取（v7）：proctor1/proctor2 列；同格多人（换行/顿号/逗号分隔）拆分。
    // 合并单元格已由 P1 的 mergedMap 展开到 grid，监考列的纵向合并天然生效。
    const proctors = [];
    for (const pc of [colIdx.proctor1, colIdx.proctor2]) {
      if (pc == null || pc < 0) continue;
      const pv = row[pc];
      if (pv === undefined || pv === null) continue;
      const sRaw = String(pv).trim();
      if (!sRaw || /^(监考|Proctor)(1|2)?$/i.test(sRaw)) continue;   // 表头行残留
      for (const piece of sRaw.split(/[、，;；,\n\r]+/).map((x) => x.trim()).filter(Boolean)) {
        if (!proctors.includes(piece)) proctors.push(piece);
      }
    }

    // 冗余交叉验证：报名人数列 vs 实际学生名单数 —— 两者不一致提示可能漏名单
    if (colIdx.count >= 0) {
      const cv = row[colIdx.count];
      const cnt = Number(cv);
      if (cv != null && String(cv).trim() !== '' && Number.isFinite(cnt) && cnt > 0 &&
          Math.abs(cnt - students.length) > Math.max(2, cnt * 0.15)) {
        warnings.push('第 ' + (r + 1) + ' 行：报名人数(' + cnt + ')与实际学生名单数(' + students.length +
          ')相差较大，可能名单列有遗漏，请人工核对该科目学生名单。');
      }
    }

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
      proctors: proctors,
      repairedFrom: entry.dateRepaired ? entry.dateRepaired.from : null,
      _row: r + 1,
    });
  }

  Object.assign(ctx, { exams, seenKeys });
}

/* P7 · 收尾：排序 + 汇总日期列表 + 组装返回值 */
function _P7_finalize(ctx) {
    const exams = ctx.exams, warnings = ctx.warnings, repairs = ctx.repairs, sheetName = ctx.sheetName, baseMonth = ctx.baseMonth, inputYear = ctx.inputYear;
  exams.sort((a, b) => (a.date + ' ' + a.time).localeCompare(b.date + ' ' + b.time));

  // 汇总去重后的考试日期（供 UI 展示核对）
  const dateList = [];
  function entryOf(dateStr) {
    const p = dateStr.split('-');
    return { y: +p[0], m: +p[1], d: +p[2] };
  }
  for (const ex of exams) {
    if (!dateList.some((x) => x.date === ex.date)) {
      dateList.push({ date: ex.date, label: describeDate(entryOf(ex.date)), weekday: ex.weekday });
    }
  }

  return {
    exams: exams,
    meta: {
      sheetName: sheetName, baseMonth: baseMonth, inputYear: inputYear,
      totalRows: exams.length, warnings: warnings, repairs: repairs, dateList: dateList,
    },
  };
}




/* ---------- 按姓名查找 ---------- */

/**
 * 模糊匹配打分（v7）：输入 query 相对姓名 name 的贴合程度，分数越高越靠前。
 *   100  姓名以输入开头（"王艺" vs "王艺睿"——前缀是最强信号）
 *    80  输入的字符在姓名中连续出现（非开头，如 "艺睿"）
 *     80 - 起始位置  姓名中段命中（越靠前分越高）
 *     10  乱序全字符命中（每个字都在，顺序不同）
 *     0   其他（双向子串但未覆盖上述情形的兜底）
 * @param {string} name  归一化后的名单姓名
 * @param {string} query 归一化后的用户输入
 * @returns {number} 分数
 */
function scoreMatch(name, query) {
  if (!query) return 0;
  if (name === query) return 200;
  if (name.startsWith(query)) return 100;
  const idx = name.indexOf(query);
  if (idx > 0) return Math.max(1, 80 - idx);
  // 乱序：每个字符都出现（仅在调用方已确认无子串命中时作为兜底，这里给出低分）
  let all = true;
  for (const ch of query) if (!name.includes(ch)) { all = false; break; }
  return all ? 10 : 0;
}

/**
 * 单场考试详情（v7，公开 API，供 UI「点击查看详情」弹层调用）。
 * 返回该场的完整信息：时间/地点/年级班级 + 监考老师（proctors）+ 同场全部学生。
 * @param {{exams:Array}} parsed parseWorkbook 的返回值
 * @param {number|string} indexOrRef 考试在 parsed.exams 中的下标，或带 _row 的考试对象
 * @returns {object|null} 详情对象；无效输入返回 null（调用方自行兜底）
 */
function getExamDetail(parsed, indexOrRef) {
  if (!parsed || !Array.isArray(parsed.exams)) return null;
  let ex = null;
  if (typeof indexOrRef === 'number' && parsed.exams[indexOrRef]) ex = parsed.exams[indexOrRef];
  else if (indexOrRef && typeof indexOrRef === 'object' && typeof indexOrRef._row === 'number') {
    ex = parsed.exams.find((e) => e._row === indexOrRef._row) || null;
  }
  if (!ex) return null;
  return {
    date: ex.date, weekday: ex.weekday, time: ex.time, minutes: ex.minutes,
    subject: ex.subject, grade: ex.grade, klass: ex.klass, room: ex.room,
    proctors: Array.isArray(ex.proctors) ? ex.proctors.slice() : [],
    students: Array.isArray(ex.students) ? ex.students.slice() : [],
    repairedFrom: ex.repairedFrom || null,
    sourceRow: ex._row,
  };
}

/**
 * v7.1：按全表源行号 _row 取详情（供 UI 弹层使用，杜绝“个人数组下标 ≠ 全表下标”的错位）。
 * _row 是 parseWorkbook 写入每场考试的唯一稳定标识（= 源表行号，从 1 起）。
 * @param {{exams:Array}} parsed parseWorkbook 的返回值
 * @param {number} row 源行号（考试对象的 _row 字段）
 * @returns {object|null} 详情对象；无效输入返回 null
 */
function getExamDetailByRow(parsed, row) {
  if (!parsed || !Array.isArray(parsed.exams) || typeof row !== 'number') return null;
  const ex = parsed.exams.find((e) => e && e._row === row) || null;
  if (!ex) return null;
  return getExamDetail(parsed, ex);   // 复用对象引用分支
}

/**
 * 按姓名查找考试（公开 API，供 UI 层调用）。
 * 匹配顺序：精确匹配（归一化后相等）→ 部分匹配（互为子串，返回候选名单）。
 * 同名检测双信号（避免合考场误报）：
 *   信号A：非合考场行的"年级|班级基础名"出现多种 → 真同名
 *   信号B：同一天两场时间区间重叠 → 单人无法同时参加
 * @param {{exams:Array}} parsed  parseWorkbook 的返回值
 * @param {string} rawName        用户输入的姓名
 * @returns {{status:'found'|'partial'|'not-found'|'empty', ...}} 查找结果
 */
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

  // 部分匹配（含大小写/全半角容错）：输入为名单子串，或名单为输入子串。
  // v7 强化：任意子串命中即列出（"王"→所有含王的姓名）；再叠加逐字连续匹配打分，
  // 输入的字符在姓名中越靠前、越连续，排序越靠前（"王艺"排在"王宇航"前面）。
  const partial = [];
  const seen = new Set();
  for (const ex of parsed.exams) {
    for (const st of ex.students) {
      const ns = normalizeName(st);
      if (!ns) continue;
      if (ns.includes(name) || name.includes(ns)) {
        if (!seen.has(ns)) { seen.add(ns); partial.push({ student: st, norm: ns, examCount: 0, score: scoreMatch(ns, name) }); }
      }
    }
  }
  // 冗余兜底：子串双向匹配无结果时，退化为"输入的每个字符都出现在姓名中"（乱序命中，
  // 如 "王睿" 命中 "王艺睿"），score 更低排在后面，确保能找到但排序合理。
  if (!partial.length && name.length >= 2) {
    for (const ex of parsed.exams) {
      for (const st of ex.students) {
        const ns = normalizeName(st);
        if (!ns || seen.has(ns)) continue;
        let ok = true;
        for (const ch of name) if (!ns.includes(ch)) { ok = false; break; }
        if (ok) { seen.add(ns); partial.push({ student: st, norm: ns, examCount: 0, score: scoreMatch(ns, name) }); }
      }
    }
  }
  for (const p of partial) {
    p.examCount = parsed.exams.filter((ex) => ex.students.some((st) => normalizeName(st) === p.norm)).length;
  }
  partial.sort((a, b) => (b.score - a.score) || a.norm.localeCompare(b.norm));
  if (partial.length) return { status: 'partial', candidates: partial };

  return { status: 'not-found' };
}

/* ---------- 导出 ---------- */
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { parseWorkbook, findExamsForStudent, getExamDetail, getExamDetailByRow, scoreMatch, normalizeName, fmtTimeRange, fmtSingleTime, rawDateCands, weekdayCN };
} else if (typeof window !== 'undefined') {
  window.ExamCore = { parseWorkbook, findExamsForStudent, getExamDetail, getExamDetailByRow, scoreMatch, normalizeName, fmtTimeRange, fmtSingleTime, rawDateCands, weekdayCN };
}
