const EXCLUDE_SHEETS = ['汇总'];
const GROUP_SHEET = '分组';
const DAY_RE = /(?:星期|周)[一二三四五六日天]/;

function ensureExcelJS(){
  if (typeof ExcelJS === 'undefined') {
    alert('表格组件未加载：请确认 exceljs.min.js 与本页面在同一目录');
    return false;
  }
  return true;
}

const dz = document.getElementById('dz');
const fileInput = document.getElementById('file');

dz.addEventListener('click', () => fileInput.click());
dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag'); });
dz.addEventListener('dragleave', () => dz.classList.remove('drag'));
dz.addEventListener('drop', e => {
  e.preventDefault(); dz.classList.remove('drag');
  if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', e => {
  if (e.target.files[0]) handleFile(e.target.files[0]);
  e.target.value = ''; // 重置，允许重复选择同一文件
});

let weekRows = null, rankRows = null, groupRows = null, groupDetail = null;
let state = null; // { students:Map, groups, dayLabels, fileName, catSheets }

function resetAll() {
  weekRows = null; rankRows = null; groupRows = null; groupDetail = null; state = null;
  ['podiumCard','rankCard','groupPodiumCard','groupRankCard','weekCard','reuploadBar','initScoreCard'].forEach(id => {
    document.getElementById(id).classList.add('hidden');
  });
  document.getElementById('uploadCard').classList.remove('hidden');
  document.getElementById('emptyCard').classList.remove('hidden');
  document.getElementById('personalInit').value = 0;
  document.getElementById('groupInit').value = 0;
  fileInput.value = '';
}

function normCell(v){
  if(v==null) return null;
  if(v instanceof Date) return v.getTime();
  if(typeof v==='object'){
    if(v.result!=null) return v.result;
    if(v.richText) return v.richText.map(t=>t.text).join('');
    if(v.text!=null) return v.text;
    return null;
  }
  return v;
}
function toRowMap(wb){
  const names = [];
  const sheets = new Map();
  wb.worksheets.forEach(ws=>{
    const rows = [];
    const rowNums = []; // 与 rows 平行，记录每行在 Excel 中的真实（物理）行号
    ws.eachRow((row, rowNumber)=>{
      const vals = row.values;
      const arr = [];
      for(let c=1;c<vals.length;c++) arr.push(normCell(vals[c]));
      rows.push(arr);
      rowNums.push(rowNumber);
    });
    names.push(ws.name);
    sheets.set(ws.name, {rows, rowNums});
  });
  return {names, sheets};
}

function handleFile(file) {
  if (!ensureExcelJS()) return;
  const reader = new FileReader();
  reader.onload = async e => {
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(e.target.result);
      const data = toRowMap(wb);
      console.log('识别到的工作表：', data.names);
      processWorkbook(data, file.name);
    } catch (err) { alert('解析失败：' + err.message); console.error(err); }
  };
  reader.readAsArrayBuffer(file);
}

function toNum(v){ if(v==null||v==='')return null; const n=Number(v); return Number.isFinite(n)?n:null; }
function trim(v){ return v==null?'':String(v).trim(); }

function processWorkbook(data, fileName) {
  const students = new Map();
  const dayLabels = [];
  const catSheets = []; // 按出现顺序记录所有有效分类表（迟到情况/默写/…），供明细导出
  const usedSheets = [], skippedSheets = [];
  const errors = []; // 记录分数列中无法识别的特殊符号/文字
  const groups = parseGroupSheet(data);

  data.names.forEach(sn => {
    const t = sn.trim();
    if (t === GROUP_SHEET) return;
    if (EXCLUDE_SHEETS.includes(t)) { skippedSheets.push(sn + '（已忽略）'); return; }
    const {rows, rowNums} = data.sheets.get(sn) || {rows: [], rowNums: []};
    if (rows.length < 2) { skippedSheets.push(sn + '（空表）'); return; }

    let headIdx = -1, sidIdx = -1, nameIdx = -1;
    for (let i = 0; i < Math.min(rows.length, 6); i++) {
      const h = (rows[i] || []).map(trim);
      const si = h.indexOf('学号'), ni = h.indexOf('姓名');
      if (si >= 0 && ni >= 0) { headIdx = i; sidIdx = si; nameIdx = ni; break; }
    }
    if (headIdx < 0) { skippedSheets.push(sn + '（无学号/姓名列）'); return; }

    const header = (rows[headIdx] || []).map(trim);
    const dayCols = [];
    header.forEach((h, idx) => {
      if (DAY_RE.test(h)) { dayCols.push({ label: h, idx }); if (!dayLabels.includes(h)) dayLabels.push(h); }
    });
    if (dayCols.length === 0) { skippedSheets.push(sn + '（无星期X分数列）'); return; }
    catSheets.push(sn);

    let counted = 0;
    for (let i = headIdx + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const sidKey = trim(r[sidIdx]);
      if (!sidKey) continue;
      if (toNum(sidKey) === null) continue;
      const name = trim(r[nameIdx]);
      if (!students.has(sidKey)) {
        const d = {}; dayLabels.forEach(l => d[l] = 0);
        // pos=全部正分之和（得分）；neg=全部负分之和（失分，保留负数）；cats=按分类表记净分
        students.set(sidKey, { sid: sidKey, name: name, days: d, pos: 0, neg: 0, cats: new Map() });
      }
      const stu = students.get(sidKey);
      if (!stu.name && name) stu.name = name;
      dayCols.forEach(({ label, idx }) => {
        const raw = r[idx];
        // 空单元格视为当天无分数，正常跳过，不算错误
        if (raw == null || trim(raw) === '') return;
        const n = toNum(raw);
        if (n === null) {
          // 识别出特殊符号/文字：跳过不累加，但记录位置用于事后提示
          errors.push({ sheet: sn, row: rowNums[i], sid: sidKey, name: name || '(未命名)', day: label, value: String(raw) });
        } else {
          stu.days[label] = (stu.days[label] || 0) + n;
          if (n > 0) stu.pos += n;
          else if (n < 0) stu.neg += n;
          stu.cats.set(sn, (stu.cats.get(sn) || 0) + n);
        }
      });
      counted++;
    }
    usedSheets.push(`${sn}（${counted} 行）`);
  });

  if (students.size === 0) { alert('未解析到任何学生数据，请检查表格格式。'); return; }

  // 按星期顺序排序列（周一→周日），避免因 sheet 遍历顺序导致列错乱
  const DAY_ORDER = {'星期一':1,'星期二':2,'星期三':3,'星期四':4,'星期五':5,'星期六':6,'星期日':7,'星期天':7,'周一':1,'周二':2,'周三':3,'周四':4,'周五':5,'周六':6,'周日':7,'周天':7};
  dayLabels.sort((a, b) => (DAY_ORDER[a] || 99) - (DAY_ORDER[b] || 99));

  // 缓存原始解析结果，初始分修改后可直接重算
  state = { students, groups, dayLabels, fileName, catSheets };
  computeAndRender();

  document.getElementById('uploadCard').classList.add('hidden');
  document.getElementById('reuploadBar').classList.remove('hidden');
  document.getElementById('initScoreCard').classList.remove('hidden');
  document.getElementById('reuploadFileName').textContent = '当前文件：' + fileName;

  // 小组数据校验：成员不在任何积分表中时提示
  const groupWarnings = [...(groups.warnings || [])];
  groups.forEach(g => {
    const miss = [...new Set(g.members.filter(sid => !students.has(sid)))];
    if (miss.length) groupWarnings.push(`${g.name} 有 ${miss.length} 个学号不在积分表中：${miss.join('、')}（不计入组总分）`);
  });

  // 计算完成后，若发现分数列中有无法识别的特殊符号/文字，弹窗提示
  if (errors.length > 0 || groupWarnings.length > 0) {
    let msg = '';
    if (errors.length > 0) {
      const preview = errors.slice(0, 10).map(e => `· [${e.sheet}] 第${e.row}行 ${e.name}(${e.sid}) ${e.day}：${e.value}`);
      msg += `⚠ 表格有误：共发现 ${errors.length} 处无法识别的内容（已自动跳过，未计入分数）：\n\n` + preview.join('\n');
      if (errors.length > 10) msg += `\n…等共 ${errors.length} 处`;
      msg += '\n\n请检查原始 Excel 对应单元格是否含特殊符号或文字（如“缺”“/”“事假”等）。';
    }
    if (groupWarnings.length > 0) {
      if (msg) msg += '\n\n';
      msg += '⚠ 分组问题：\n' + groupWarnings.join('\n');
    }
    alert(msg);
  }
}

function computeAndRender() {
  if (!state) return;
  const { students, groups, dayLabels } = state;
  const personalInit = Number(document.getElementById('personalInit').value) || 0;
  const groupInit = Number(document.getElementById('groupInit').value) || 0;

  const sidSorted = [...students.keys()].sort((a, b) => Number(a) - Number(b));

  // 积分周表：每日明细不变，周总分 = 每日合计 + 个人初始分
  const wHead = ['学号', '姓名', '初始分', ...dayLabels, '周总分'];
  const wRows = [wHead];
  const dayTotals = {}; dayLabels.forEach(l => dayTotals[l] = 0);
  let grandTotal = 0, initTotal = 0;
  const studentDaily = {};  // 仅每日合计（不含个人初始分）——供小组汇总使用，保证两组互不干扰
  const studentTotal = {};  // 每日合计 + 个人初始分
  sidSorted.forEach(sid => {
    const s = students.get(sid);
    let week = 0; const line = [s.sid, s.name, personalInit];
    dayLabels.forEach(l => { const v = s.days[l] || 0; line.push(v); week += v; dayTotals[l] += v; });
    const total = week + personalInit;
    line.push(total); wRows.push(line);
    grandTotal += total; initTotal += personalInit;
    studentDaily[sid] = week; studentTotal[sid] = total;
  });
  const sumLine = ['合计', '', initTotal];
  dayLabels.forEach(l => sumLine.push(dayTotals[l]));
  sumLine.push(grandTotal); wRows.push(sumLine);
  weekRows = wRows;

  // 个人排名
  const list = sidSorted.map(sid => ({ sid: students.get(sid).sid, name: students.get(sid).name, init: personalInit, total: studentTotal[sid] }));
  list.sort((a, b) => b.total - a.total || Number(a.sid) - Number(b.sid));
  let rank = 0, prev = null;
  list.forEach((it, i) => { if (prev === null || it.total !== prev) { rank = i + 1; prev = it.total; } it.rank = rank; });
  rankRows = [['名次', '学号', '姓名', '初始分', '周总分']];
  list.forEach(it => rankRows.push([it.rank, it.sid, it.name, it.init, it.total]));

  // 小组排名：组总分 = 成员每日合计之和 + 小组初始分（不含个人初始分，互不干扰）
  // 同时组装成员贡献明细：每人 得分(pos) / 失分(loss) / 净贡献(net) / 分类净分(cats)
  const groupList = groups.map(g => {
    const members = g.members.map(sid => {
      const s = students.get(sid);
      if (!s) return null;
      const net = studentDaily[sid] != null ? studentDaily[sid] : 0;
      return {
        sid: s.sid, name: s.name,
        pos: s.pos || 0, loss: -(s.neg || 0), net: net, cats: s.cats
      };
    }).filter(Boolean);
    // 组内按净贡献降序，平分比得分，再按学号
    members.sort((a, b) => b.net - a.net || b.pos - a.pos || Number(a.sid) - Number(b.sid));
    const netSum = members.reduce((t, m) => t + m.net, 0);
    const posSum = members.reduce((t, m) => t + m.pos, 0);
    const lossSum = members.reduce((t, m) => t + m.loss, 0);
    const total = netSum + groupInit;
    return {
      name: g.name, count: members.length, init: groupInit, total: total,
      avg: members.length ? total / members.length : 0,
      members: members, posSum: posSum, lossSum: lossSum, netSum: netSum
    };
  });
  groupList.sort((a, b) => b.total - a.total || a.name.localeCompare(b.name, 'zh'));
  let grank = 0, gprev = null;
  groupList.forEach((it, i) => { if (gprev === null || it.total !== gprev) { grank = i + 1; gprev = it.total; } it.rank = grank; });
  groupDetail = groupList;
  groupRows = [['名次', '小组', '人数', '初始分', '组总分', '人均分']];
  groupList.forEach(it => groupRows.push([it.rank, it.name, it.count, it.init, it.total, Math.round(it.avg * 100) / 100]));

  renderWeek(wRows, dayLabels.length);
  renderPodium(list);
  renderRank(list);
  renderGroupPodium(groupList);
  renderGroupRank(groupList);
}

function parseGroupSheet(data) {
  const groups = [];
  groups.warnings = [];
  // 工作表名允许带首尾空格
  const realName = data.names.find(n => n.trim() === GROUP_SHEET);
  if (!realName) return groups;
  const {rows} = data.sheets.get(realName) || {rows: []};
  if (rows.length < 2) return groups;
  let nameRowIdx = -1;
  for (let i = 0; i < Math.min(rows.length, 6); i++) {
    if ((rows[i] || []).some(v => /第.+小组/.test(trim(v)))) { nameRowIdx = i; break; }
  }
  if (nameRowIdx < 0) return groups;
  const nameRow = (rows[nameRowIdx] || []).map(trim);
  nameRow.forEach((gname, colIdx) => {
    if (!/第.+小组/.test(gname)) return;
    const members = [];
    const seen = new Set();
    for (let r = nameRowIdx + 1; r < rows.length; r++) {
      const v = (rows[r] || [])[colIdx];
      const sid = trim(v);
      if (!sid) continue;
      if (toNum(sid) === null) continue;
      if (seen.has(sid)) {
        groups.warnings.push(`${gname} 中学号 ${sid} 重复出现，已自动去重`);
        continue;
      }
      seen.add(sid);
      members.push(sid);
    }
    groups.push({ name: gname, members });
  });
  return groups;
}

function fmt(v){ if(typeof v!=='number')return v; return Number.isInteger(v)?v:Math.round(v*10)/10; }
// 人均分固定保留两位小数
function fmt2(v){ if(typeof v!=='number')return v; return v.toFixed(2); }

function renderWeek(rows, dayCount) {
  document.getElementById('weekCard').classList.remove('hidden');
  document.getElementById('emptyCard').classList.add('hidden');
  const thead = document.getElementById('weekHead'), tbody = document.getElementById('weekBody');
  thead.innerHTML = ''; tbody.innerHTML = '';
  rows[0].forEach(h => { const th = document.createElement('th'); th.textContent = h; thead.appendChild(th); });
  const n = rows.length - 2;
  for (let i = 1; i < rows.length; i++) {
    const tr = document.createElement('tr');
    if (i === rows.length - 1) tr.className = 'total-row';
    rows[i].forEach((v, j) => {
      const td = document.createElement('td'); td.textContent = fmt(v);
      if (typeof v === 'number' && i !== rows.length - 1 && j >= 3) { if (v < 0) td.className = 'neg'; else if (v > 0) td.className = 'pos'; }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }
  document.getElementById('weekStat').textContent = `共 ${n} 名学生，${dayCount} 天`;
}

function renderPodium(list) {
  const card = document.getElementById('podiumCard'), box = document.getElementById('podium');
  box.innerHTML = '';
  if (list.length === 0) { card.classList.add('hidden'); return; }
  card.classList.remove('hidden');
  const order = [];
  if (list[1]) order.push({ who: list[1], cls: 's2' });
  if (list[0]) order.push({ who: list[0], cls: 's1' });
  if (list[2]) order.push({ who: list[2], cls: 's3' });
  order.forEach(({ who, cls }) => {
    const seat = document.createElement('div'); seat.className = 'seat ' + cls;
    seat.innerHTML = `
      <div class="person">
        <div class="medal">${who.rank}</div>
        <div class="sid"></div><div class="sname"></div><div class="score"></div>
      </div>
      <div class="block">第${who.rank}名</div>`;
    seat.querySelector('.sid').textContent = who.sid;
    seat.querySelector('.sname').textContent = who.name;
    seat.querySelector('.score').textContent = fmt(who.total) + ' 分';
    box.appendChild(seat);
  });
}

function renderRank(list) {
  document.getElementById('rankCard').classList.remove('hidden');
  const thead = document.getElementById('rankHead'), tbody = document.getElementById('rankBody');
  thead.innerHTML = ''; tbody.innerHTML = '';
  ['名次', '学号', '姓名', '初始分', '周总分'].forEach(h => { const th = document.createElement('th'); th.textContent = h; thead.appendChild(th); });
  list.forEach(it => {
    const tr = document.createElement('tr');
    if (it.rank === 1) tr.className = 'r1'; else if (it.rank === 2) tr.className = 'r2'; else if (it.rank === 3) tr.className = 'r3';
    [it.rank, it.sid, it.name, it.init, fmt(it.total)].forEach((v, j) => {
      const td = document.createElement('td'); td.textContent = v;
      if (j === 4 && typeof v === 'number') { if (v < 0) td.className = 'neg'; else if (v > 0) td.className = 'pos'; }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  document.getElementById('rankStat').textContent = `共 ${list.length} 人`;
}

function renderGroupPodium(glist) {
  const card = document.getElementById('groupPodiumCard'), box = document.getElementById('groupPodium');
  box.innerHTML = '';
  if (glist.length === 0) { card.classList.add('hidden'); return; }
  card.classList.remove('hidden');
  const order = [];
  if (glist[1]) order.push({ who: glist[1], cls: 'g2' });
  if (glist[0]) order.push({ who: glist[0], cls: 'g1' });
  if (glist[2]) order.push({ who: glist[2], cls: 'g3' });
  order.forEach(({ who, cls }) => {
    const seat = document.createElement('div'); seat.className = 'gseat ' + cls;
    seat.innerHTML = `
      <div class="gcard">
        <div class="gname"></div><div class="gscore"></div><div class="gcount"></div>
      </div>
      <div class="gblock">第${who.rank}名</div>`;
    seat.querySelector('.gname').textContent = who.name;
    seat.querySelector('.gscore').textContent = fmt(who.total) + ' 分';
    seat.querySelector('.gcount').textContent = who.count + ' 人 · 人均 ' + fmt2(who.avg);
    box.appendChild(seat);
  });
}

function renderGroupRank(glist) {
  document.getElementById('groupRankCard').classList.remove('hidden');
  const thead = document.getElementById('groupHead'), tbody = document.getElementById('groupBody');
  thead.innerHTML = ''; tbody.innerHTML = '';
  ['名次', '小组', '人数', '初始分', '组总分', '人均分'].forEach(h => { const th = document.createElement('th'); th.textContent = h; thead.appendChild(th); });
  glist.forEach(it => {
    const tr = document.createElement('tr');
    if (it.rank === 1) tr.className = 'r1'; else if (it.rank === 2) tr.className = 'r2'; else if (it.rank === 3) tr.className = 'r3';
    [it.rank, it.name, it.count, it.init, fmt(it.total), fmt2(it.avg)].forEach((v, j) => {
      const td = document.createElement('td'); td.textContent = v;
      const num = j === 4 ? it.total : (j === 5 ? it.avg : null);
      if (num !== null) { if (num < 0) td.className = 'neg'; else if (num > 0) td.className = 'pos'; }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  document.getElementById('groupStat').textContent = `共 ${glist.length} 个小组`;
}

function datedName(prefix){ const t=new Date(),p=n=>String(n).padStart(2,'0'); return `${prefix}_${t.getFullYear()}${p(t.getMonth()+1)}${p(t.getDate())}.xlsx`; }

async function exportAoa(aoa, sheetName, fileName, colWidths){
  if (!ensureExcelJS()) return;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName);
  if(colWidths) ws.columns = colWidths.map(w=>({width:w}));
  aoa.forEach(r=>ws.addRow(r));
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = fileName;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

document.getElementById('downloadWeekBtn').addEventListener('click', () => {
  if (!weekRows) return;
  const widths = weekRows[0].map((h,i)=> i<2 ? 12 : 10);
  exportAoa(weekRows, '积分周表', datedName('积分周表'), widths);
});
document.getElementById('downloadRankBtn').addEventListener('click', () => {
  if (!rankRows) return;
  exportAoa(rankRows, '个人排名', datedName('个人排名表'), [8,12,14,10,10]);
});
// 小组排名导出：Sheet1 汇总（得分/失分拆分）+ Sheet2 成员得分失分贡献明细，带样式
async function exportGroupWorkbook(glist, catSheets, groupInit) {
  if (!ensureExcelJS()) return;
  const C = {
    green: 'FF059669', greenDeep: 'FF047857', red: 'FFDC2626', ink: 'FF16181F',
    muted: 'FF6B7280', headerBg: 'FF059669', bandBg: 'FFECFDF5',
    goldBg: 'FFFFFBEB', silverBg: 'FFF8FAFC', bronzeBg: 'FFFFF7ED'
  };
  const thin = { style: 'thin', color: { argb: 'FFE5E7EB' } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };
  const center = { horizontal: 'center', vertical: 'middle' };

  const colorSigned = (cell, val, color) => {
    if (typeof val === 'number' && val !== 0) cell.font = { bold: true, color: { argb: color }, size: 11 };
  };
  const colorNet = (cell, val) => {
    if (typeof val !== 'number') return;
    if (val > 0) cell.font = { bold: true, color: { argb: C.green }, size: 11 };
    else if (val < 0) cell.font = { bold: true, color: { argb: C.red }, size: 11 };
  };
  const shadeRow = (ws, rowIdx, ncol, bg) => {
    for (let j = 1; j <= ncol; j++) {
      ws.getCell(rowIdx, j).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
    }
  };

  const wb = new ExcelJS.Workbook();

  // ===== Sheet1 小组排名汇总 =====
  const ws1 = wb.addWorksheet('小组排名');
  const h1 = ['名次', '小组', '人数', '初始分', '得分合计', '失分合计', '组总分', '人均分'];
  ws1.columns = [8, 14, 8, 10, 12, 12, 12, 10].map(w => ({ width: w }));
  ws1.mergeCells(1, 1, 1, h1.length);
  const t1 = ws1.getCell(1, 1);
  t1.value = '小组排名汇总';
  t1.font = { size: 15, bold: true, color: { argb: C.greenDeep } };
  t1.alignment = center;
  ws1.getRow(1).height = 30;
  ws1.mergeCells(2, 1, 2, h1.length);
  const s1 = ws1.getCell(2, 1);
  s1.value = `小组初始分 ${fmt(groupInit)} 分/组；组总分 = 得分合计 − 失分合计 + 初始分`;
  s1.font = { size: 10.5, color: { argb: C.muted } };
  s1.alignment = center;
  h1.forEach((h, j) => {
    const c = ws1.getCell(3, j + 1);
    c.value = h;
    c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.headerBg } };
    c.alignment = center; c.border = border;
  });
  let r = 4, gPos = 0, gLoss = 0, gNet = 0, gCnt = 0;
  glist.forEach(it => {
    gPos += it.posSum; gLoss += it.lossSum; gNet += it.netSum; gCnt += it.count;
    const avgNum = Math.round(it.avg * 100) / 100;
    const vals = [it.rank, it.name, it.count, it.init, fmt(it.posSum), fmt(it.lossSum), fmt(it.total), avgNum];
    vals.forEach((v, j) => {
      const c = ws1.getCell(r, j + 1);
      c.value = v; c.border = border; c.alignment = center;
      c.font = { color: { argb: C.ink }, size: 11 };
    });
    ws1.getCell(r, 8).numFmt = '0.00';
    colorSigned(ws1.getCell(r, 5), it.posSum, C.green);
    colorSigned(ws1.getCell(r, 6), it.lossSum, C.red);
    colorNet(ws1.getCell(r, 7), it.total);
    if (it.rank === 1) shadeRow(ws1, r, h1.length, C.goldBg);
    else if (it.rank === 2) shadeRow(ws1, r, h1.length, C.silverBg);
    else if (it.rank === 3) shadeRow(ws1, r, h1.length, C.bronzeBg);
    r++;
  });
  const initAll = fmt(groupInit * glist.length);
  const tot = ['合计', '', gCnt, initAll, fmt(gPos), fmt(gLoss), fmt(gNet + groupInit * glist.length), ''];
  tot.forEach((v, j) => {
    const c = ws1.getCell(r, j + 1);
    c.value = v; c.border = border; c.alignment = center;
    c.font = { bold: true, color: { argb: C.greenDeep }, size: 11 };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bandBg } };
  });
  ws1.views = [{ state: 'frozen', ySplit: 3 }];

  // ===== Sheet2 成员贡献明细 =====
  const ws2 = wb.addWorksheet('成员贡献明细');
  const baseCols = ['组内名次', '学号', '姓名', '得分', '失分', '净贡献'];
  // 只保留至少出现过非零记录的分类表，避免空列
  const activeCats = catSheets.filter(cn =>
    glist.some(g => g.members.some(m => (m.cats.get(cn) || 0) !== 0))
  );
  const allCols = [...baseCols, ...activeCats];
  ws2.columns = [10, 12, 12, 10, 10, 10, ...activeCats.map(() => 13)].map(w => ({ width: w }));
  ws2.mergeCells(1, 1, 1, allCols.length);
  const t2 = ws2.getCell(1, 1);
  t2.value = '小组成员得分 / 失分贡献明细';
  t2.font = { size: 15, bold: true, color: { argb: C.greenDeep } };
  t2.alignment = center;
  ws2.getRow(1).height = 30;
  ws2.mergeCells(2, 1, 2, allCols.length);
  const s2 = ws2.getCell(2, 1);
  s2.value = '得分 = 全部正分之和；失分 = 全部负分绝对值之和；净贡献 = 得分 − 失分（不含小组初始分）';
  s2.font = { size: 10.5, color: { argb: C.muted } };
  s2.alignment = center;

  let rr = 3;
  glist.forEach(it => {
    ws2.mergeCells(rr, 1, rr, allCols.length);
    const band = ws2.getCell(rr, 1);
    band.value = `${it.name} · 小组第${it.rank}名 · ${it.count}人 · 得分 +${fmt(it.posSum)} · 失分 −${fmt(it.lossSum)} · 净贡献 ${fmt(it.netSum)}`;
    band.font = { bold: true, size: 12, color: { argb: C.greenDeep } };
    band.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.bandBg } };
    band.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
    for (let j = 1; j <= allCols.length; j++) ws2.getCell(rr, j).border = border;
    rr++;
    allCols.forEach((h, j) => {
      const c = ws2.getCell(rr, j + 1);
      c.value = h;
      c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10.5 };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.headerBg } };
      c.alignment = center; c.border = border;
    });
    rr++;
    it.members.forEach((m, idx) => {
      const vals = [idx + 1, m.sid, m.name, fmt(m.pos), fmt(m.loss), fmt(m.net)];
      activeCats.forEach(cn => {
        const v = m.cats.get(cn) || 0;
        vals.push(v === 0 ? '' : fmt(v));
      });
      vals.forEach((v, j) => {
        const c = ws2.getCell(rr, j + 1);
        c.value = v; c.border = border; c.alignment = center;
        c.font = { color: { argb: C.ink }, size: 11 };
      });
      colorSigned(ws2.getCell(rr, 4), m.pos, C.green);
      colorSigned(ws2.getCell(rr, 5), m.loss, C.red);
      colorNet(ws2.getCell(rr, 6), m.net);
      activeCats.forEach((cn, k) => {
        const v = m.cats.get(cn) || 0;
        if (v !== 0) colorNet(ws2.getCell(rr, 7 + k), v);
      });
      rr++;
    });
    rr++; // 组间空一行
  });

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = datedName('小组排名表');
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

document.getElementById('downloadGroupBtn').addEventListener('click', () => {
  if (!groupDetail || !state) return;
  const groupInit = Number(document.getElementById('groupInit').value) || 0;
  exportGroupWorkbook(groupDetail, state.catSheets, groupInit);
});

// 初始分修改后实时重算（二者独立读取，互不干扰）
document.getElementById('personalInit').addEventListener('input', computeAndRender);
document.getElementById('groupInit').addEventListener('input', computeAndRender);
