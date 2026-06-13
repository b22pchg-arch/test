'use strict';

const $ = sel => document.querySelector(sel);
const els = {
  file: $('#excelFile'), sheet: $('#sheetSelect'), btnParse: $('#btnParse'), btnSave: $('#btnSaveBank'), btnLoad: $('#btnLoadBank'), btnClear: $('#btnClearBank'),
  btnStart: $('#btnStartQuiz'), btnSubmit: $('#btnSubmitQuiz'), btnSubmitSticky: $('#btnSubmitSticky'), btnPrint: $('#btnPrint'),
  count: $('#quizCount'), shuffleQ: $('#shuffleQuestions'), shuffleO: $('#shuffleOptions'), explain: $('#showAutoExplain'), seed: $('#seedInput'),
  status: $('#status'), statPills: $('#statPills'), preview: $('#preview'), quizInfo: $('#quizInfo'), quizList: $('#quizList'),
  resultSummary: $('#resultSummary'), resultList: $('#resultList'), progressText: $('#progressText'), progressBar: $('#progressBar')
};

const state = {
  workbook: null,
  fileName: '',
  bank: [],
  errors: [],
  mapping: null,
  quiz: [],
  submitted: false,
  lastResult: null
};

const DB_NAME = 'excel_quiz_offline_db_v1';
const STORE = 'kv';
const BANK_KEY = 'question_bank';

const VI_STOPWORDS = new Set(`
la cua va hoac de den duoc bi trong ngoai mot cac nhung nhung ma thi voi cho ve theo tai tu khi luc nao gi do nay kia ay tren duoi vao ra bang nhu neu thi hon kem da dang se can phai duoc khong chua cung moi sau truoc phan noi dung cau hoi phuong an lua chon dap an dung sai don vi linh vuc he thong
là của và hoặc để đến được bị trong ngoài một các những mà thì với cho về theo tại từ khi lúc nào gì đó này kia ấy trên dưới vào ra bằng như nếu hơn kém đã đang sẽ cần phải không chưa cùng mỗi sau trước phần nội dung câu hỏi phương án lựa chọn đáp án đúng sai đơn vị lĩnh vực hệ thống
`.split(/\s+/).filter(Boolean));

function setStatus(msg, type='') {
  els.status.textContent = msg;
  els.status.className = 'status ' + type;
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function normalizeVN(str) {
  return String(str ?? '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd').replace(/Đ/g, 'D')
    .toLowerCase()
    .replace(/[“”‘’]/g, ' ')
    .replace(/[_.,;:!?()[\]{}<>/\\|+*=~`^"'-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function visibleText(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/\r\n/g, '\n').replace(/\s+/g, ' ').trim();
}

function tokenize(str) {
  const norm = normalizeVN(str);
  if (!norm) return [];
  return norm.split(/\s+/).filter(t => t.length > 1 && !VI_STOPWORDS.has(t));
}

function unique(arr) { return [...new Set(arr.filter(Boolean))]; }
function intersection(a, b) { const bs = new Set(b); return a.filter(x => bs.has(x)); }
function difference(a, b) { const bs = new Set(b); return a.filter(x => !bs.has(x)); }
function top(arr, n=10) { return unique(arr).slice(0, n); }

function extractNumbers(str) {
  const src = normalizeVN(str);
  const re = /(?:tren|duoi|den|tu|khong qua|toi thieu|toi da|lon hon|nho hon)?\s*[-+]?\d+(?:[,.]\d+)?\s*(?:kv|v|a|ka|kw|mw|mva|kva|hz|%|phan tram|phut|gio|ngay|thang|nam|m|km)?/g;
  return unique((src.match(re) || []).map(x => x.replace(/\s+/g, ' ').trim()).filter(Boolean));
}

function extractRelationWords(str) {
  const src = normalizeVN(str);
  const keys = [
    'tren','duoi','den','tu','khong qua','toi thieu','toi da','lon hon','nho hon','bang','khong','chua','cam','phai','duoc','truoc','sau','trong','ngoai','dong','cat','mo','cap','ha','cao','sieu cao'
  ];
  return keys.filter(k => src.includes(k));
}

function phraseSet(tokens) {
  const out = [];
  for (let n = 2; n <= 3; n++) {
    for (let i = 0; i <= tokens.length - n; i++) out.push(tokens.slice(i, i+n).join(' '));
  }
  return unique(out);
}

function jaccard(a, b) {
  const A = new Set(a), B = new Set(b);
  const uni = new Set([...A, ...B]);
  if (!uni.size) return 0;
  let inter = 0; A.forEach(x => { if (B.has(x)) inter++; });
  return inter / uni.size;
}

function explainDifference(correct, wrong) {
  const ct = tokenize(correct), wt = tokenize(wrong);
  const cp = phraseSet(ct), wp = phraseSet(wt);
  const common = top([...intersection(cp, wp), ...intersection(ct, wt)], 12);
  const correctOnly = top([...difference(cp, wp), ...difference(ct, wt)], 12);
  const wrongOnly = top([...difference(wp, cp), ...difference(wt, ct)], 12);
  const cNums = extractNumbers(correct), wNums = extractNumbers(wrong);
  const cRel = extractRelationWords(correct), wRel = extractRelationWords(wrong);
  const sim = jaccard(unique([...ct, ...cp]), unique([...wt, ...wp]));

  const parts = [];
  if (common.length) parts.push(`<b>Giống:</b> ${common.map(x => `<span class="tag">${escapeHtml(x)}</span>`).join(' ')}`);
  else parts.push('<b>Giống:</b> rất ít từ khóa trùng nhau.');

  const numsDiff = cNums.join('|') !== wNums.join('|');
  if (numsDiff && (cNums.length || wNums.length)) {
    parts.push(`<b>Khác số liệu/mốc:</b> đúng có ${tagList(cNums) || 'không rõ'}; sai có ${tagList(wNums) || 'không rõ'}.`);
  }
  const relDiff = cRel.join('|') !== wRel.join('|');
  if (relDiff && (cRel.length || wRel.length)) {
    parts.push(`<b>Khác quan hệ/điều kiện:</b> đúng có ${tagList(cRel) || 'không rõ'}; sai có ${tagList(wRel) || 'không rõ'}.`);
  }
  if (correctOnly.length) parts.push(`<b>Đáp án đúng nhấn mạnh:</b> ${tagList(correctOnly)}.`);
  if (wrongOnly.length) parts.push(`<b>Phương án sai lệch ở:</b> ${tagList(wrongOnly)}.`);

  let note = '';
  if (sim >= .78) note = 'Hai phương án rất giống nhau; cần chú ý các cụm khóa hoặc số liệu nhỏ khác nhau.';
  else if (sim >= .45) note = 'Hai phương án cùng chủ đề nhưng khác một số điều kiện/từ khóa quan trọng.';
  else note = 'Hai phương án khác khá nhiều về đối tượng hoặc nội dung chính.';
  parts.push(`<b>Mức gần nhau:</b> ${(sim*100).toFixed(0)}%. ${escapeHtml(note)}`);
  return parts.join('<br>');
}

function tagList(items) {
  return unique(items).slice(0, 12).map(x => `<span class="tag">${escapeHtml(x)}</span>`).join(' ');
}

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function idbPut(key, value) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).put(value, key);
    tx.oncomplete = () => resolve(); tx.onerror = () => reject(tx.error);
  });
}
async function idbGet(key) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readonly');
    const req = tx.objectStore(STORE).get(key);
    req.onsuccess = () => resolve(req.result); req.onerror = () => reject(req.error);
  });
}
async function idbDel(key) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).delete(key);
    tx.oncomplete = () => resolve(); tx.onerror = () => reject(tx.error);
  });
}

async function readWorkbookFromFile(file) {
  const buf = await file.arrayBuffer();
  return XLSX.read(buf, {type:'array', cellDates:true, raw:false});
}

function detectHeaderRow(rows) {
  let best = {idx:0, score:-1};
  const maxRows = Math.min(rows.length, 40);
  for (let r = 0; r < maxRows; r++) {
    const cells = rows[r] || [];
    const joined = cells.map(normalizeVN).join(' | ');
    let score = 0;
    if (/cau hoi|noi dung|question/.test(joined)) score += 4;
    if (/dap an|correct|answer|dung/.test(joined)) score += 3;
    const optMatches = joined.match(/phuong an|lua chon|option|choice/g);
    if (optMatches) score += optMatches.length * 2;
    cells.forEach(c => {
      const n = normalizeVN(c);
      if (/^(a|b|c|d|e)$/.test(n)) score += 1;
      if (/^phuong an lua chon \d+$/.test(n)) score += 3;
    });
    if (score > best.score) best = {idx:r, score};
  }
  return best.idx;
}

function headerScore(cell, kind) {
  const h = normalizeVN(cell);
  if (!h) return 0;
  const has = (...keys) => keys.some(k => h.includes(k));
  if (kind === 'id') return has('stt','so thu tu','id','ma cau','ma so') ? 5 : 0;
  if (kind === 'question') {
    let s = 0;
    if (has('cau hoi','question')) s += 8;
    if (has('noi dung')) s += 5;
    if (has('phuong an','lua chon','dap an')) s -= 8;
    return s;
  }
  if (kind === 'answer') {
    let s = 0;
    if (has('dap an','answer','correct','dung')) s += 8;
    if (has('phuong an','lua chon','option','choice')) s -= 6;
    return s;
  }
  if (kind === 'option') {
    let s = 0;
    if (has('phuong an','lua chon','option','choice')) s += 7;
    if (/^(a|b|c|d|e|f)$/.test(h)) s += 8;
    if (/^pa\s*[a-f0-9]?/.test(h)) s += 5;
    if (has('dap an dung')) s -= 6;
    return s;
  }
  if (kind === 'source') return has('can cu','phap ly','giai thich','explain','reference','ghi chu','nguon') ? 6 : 0;
  return 0;
}

function bestColumn(headers, kind, exclude = new Set()) {
  let best = {idx:-1, score:0};
  headers.forEach((h, i) => {
    if (exclude.has(i)) return;
    const s = headerScore(h, kind);
    if (s > best.score) best = {idx:i, score:s};
  });
  return best.idx;
}

function detectMapping(rows, headerIdx) {
  const headers = rows[headerIdx] || [];
  const exclude = new Set();
  const idCol = bestColumn(headers, 'id'); if (idCol >= 0) exclude.add(idCol);
  const questionCol = bestColumn(headers, 'question', exclude); if (questionCol >= 0) exclude.add(questionCol);
  const answerCol = bestColumn(headers, 'answer', exclude); if (answerCol >= 0) exclude.add(answerCol);
  const sourceCol = bestColumn(headers, 'source', exclude); if (sourceCol >= 0) exclude.add(sourceCol);

  let optionCols = [];
  headers.forEach((h, i) => {
    if ([idCol, questionCol, answerCol, sourceCol].includes(i)) return;
    if (headerScore(h, 'option') >= 5) optionCols.push(i);
  });
  optionCols.sort((a, b) => a - b);

  if (optionCols.length < 2 && questionCol >= 0) {
    const candidates = [];
    for (let c = 0; c < Math.max(...rows.slice(headerIdx, Math.min(rows.length, headerIdx+25)).map(r => (r || []).length)); c++) {
      if ([idCol, questionCol, answerCol, sourceCol].includes(c)) continue;
      let filled = 0, avgLen = 0;
      for (let r = headerIdx + 1; r < Math.min(rows.length, headerIdx + 21); r++) {
        const txt = visibleText((rows[r] || [])[c]);
        if (txt) { filled++; avgLen += txt.length; }
      }
      if (filled >= 3) candidates.push({c, filled, avgLen});
    }
    optionCols = candidates.sort((a,b) => b.filled - a.filled || b.avgLen - a.avgLen).slice(0, 6).map(x => x.c).sort((a,b)=>a-b);
  }
  return {headerIdx, headers, idCol, questionCol, answerCol, sourceCol, optionCols};
}

function stripOptionMarker(txt) {
  return visibleText(txt).replace(/^\s*(\*|✓|✔|✅|\[x\]|\(x\)|x\.)\s*/i, '').trim();
}
function hasCorrectMarker(txt) {
  return /^\s*(\*|✓|✔|✅|\[x\]|\(x\)|x\.)\s*/i.test(visibleText(txt));
}

function resolveCorrectIndex(answerRaw, options) {
  const ans = visibleText(answerRaw);
  if (!ans) {
    const marked = options.findIndex(o => o.marked);
    return marked >= 0 ? marked : -1;
  }
  const aNorm = normalizeVN(ans).trim();
  const num = aNorm.match(/(?:^|\D)([1-9])(?:\D|$)/);
  if (num) {
    const idx = Number(num[1]) - 1;
    if (idx >= 0 && idx < options.length) return idx;
  }
  const letter = aNorm.match(/\b([a-f])\b/);
  if (letter) {
    const idx = 'abcdef'.indexOf(letter[1]);
    if (idx >= 0 && idx < options.length) return idx;
  }
  const byChoice = aNorm.match(/lua chon\s*([1-9])|phuong an\s*([1-9])/);
  if (byChoice) {
    const idx = Number(byChoice[1] || byChoice[2]) - 1;
    if (idx >= 0 && idx < options.length) return idx;
  }
  const ansClean = normalizeVN(ans);
  let best = {idx:-1, score:0};
  options.forEach((o, idx) => {
    const optNorm = normalizeVN(o.text);
    let score = optNorm === ansClean ? 1 : jaccard(tokenize(ansClean), tokenize(optNorm));
    if (optNorm.includes(ansClean) || ansClean.includes(optNorm)) score = Math.max(score, .85);
    if (score > best.score) best = {idx, score};
  });
  return best.score >= .55 ? best.idx : -1;
}

function parseRows(rows) {
  const headerIdx = detectHeaderRow(rows);
  const mapping = detectMapping(rows, headerIdx);
  const bank = [], errors = [];
  if (mapping.questionCol < 0 || mapping.optionCols.length < 2) {
    throw new Error('Không nhận diện được cột câu hỏi hoặc các cột phương án. Hãy kiểm tra hàng tiêu đề của Excel.');
  }
  for (let r = headerIdx + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const question = visibleText(row[mapping.questionCol]);
    if (!question) continue;
    const options = mapping.optionCols.map((c, idx) => {
      const raw = visibleText(row[c]);
      return {originalIndex: idx, text: stripOptionMarker(raw), raw, marked: hasCorrectMarker(raw)};
    }).filter(o => o.text);
    if (options.length < 2) {
      errors.push({row:r+1, reason:'Thiếu phương án lựa chọn', question});
      continue;
    }
    const correctIndex = resolveCorrectIndex(mapping.answerCol >= 0 ? row[mapping.answerCol] : '', options);
    if (correctIndex < 0 || correctIndex >= options.length) {
      errors.push({row:r+1, reason:'Không xác định được đáp án đúng', question});
      continue;
    }
    bank.push({
      id: visibleText(mapping.idCol >= 0 ? row[mapping.idCol] : '') || String(bank.length + 1),
      sourceRow: r + 1,
      question,
      options: options.map(o => o.text),
      correctIndex,
      correctText: options[correctIndex].text,
      source: visibleText(mapping.sourceCol >= 0 ? row[mapping.sourceCol] : ''),
      answerRaw: visibleText(mapping.answerCol >= 0 ? row[mapping.answerCol] : ''),
      sheetHeaderRow: headerIdx + 1
    });
  }
  return {bank, errors, mapping};
}

function renderStats() {
  const sheetName = els.sheet.value || 'Chưa chọn sheet';
  els.statPills.innerHTML = `
    <span class="pill">${state.bank.length} câu hợp lệ</span>
    <span class="pill">${state.errors.length} dòng lỗi</span>
    <span class="pill">${escapeHtml(sheetName)}</span>
  `;
  els.btnStart.disabled = state.bank.length === 0;
  els.btnSave.disabled = state.bank.length === 0;
  if (state.bank.length) els.count.max = String(state.bank.length);
}

function renderPreview() {
  if (!state.bank.length) {
    els.preview.innerHTML = '<table><tbody><tr><td class="muted">Chưa có dữ liệu hợp lệ.</td></tr></tbody></table>';
    return;
  }
  const rows = state.bank.slice(0, 30).map((q, i) => `
    <tr>
      <td class="nowrap">${i + 1}</td>
      <td>${escapeHtml(q.question)}</td>
      <td>${escapeHtml(q.correctText)}</td>
      <td>${escapeHtml(q.options.join(' | '))}</td>
      <td>${escapeHtml(q.source || '')}</td>
    </tr>`).join('');
  els.preview.innerHTML = `<table><thead><tr><th>STT</th><th>Câu hỏi</th><th>Đáp án đúng</th><th>Phương án</th><th>Căn cứ/giải thích</th></tr></thead><tbody>${rows}</tbody></table>`;
}

async function handleFileSelected() {
  const file = els.file.files && els.file.files[0];
  if (!file) return;
  state.fileName = file.name;
  setStatus('Đang đọc workbook: ' + file.name + ' ...');
  try {
    await idle();
    state.workbook = await readWorkbookFromFile(file);
    els.sheet.innerHTML = state.workbook.SheetNames.map(n => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join('');
    parseSelectedSheet();
  } catch (err) {
    console.error(err);
    setStatus('❌ Không đọc được Excel: ' + (err.message || err), 'bad');
  }
}

function parseSelectedSheet() {
  if (!state.workbook) return;
  const sheetName = els.sheet.value || state.workbook.SheetNames[0];
  const ws = state.workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:false, defval:''});
  try {
    const {bank, errors, mapping} = parseRows(rows);
    state.bank = bank; state.errors = errors; state.mapping = mapping;
    renderStats(); renderPreview();
    const msg = [
      `✅ Đã nạp ${bank.length} câu hỏi hợp lệ từ “${state.fileName || 'workbook'}” / sheet “${sheetName}”.`,
      `Hàng tiêu đề phát hiện: ${mapping.headerIdx + 1}.`,
      `Cột câu hỏi: ${colName(mapping.questionCol)}; cột đáp án: ${mapping.answerCol >= 0 ? colName(mapping.answerCol) : 'không có'}; cột phương án: ${mapping.optionCols.map(colName).join(', ')}; cột căn cứ/giải thích: ${mapping.sourceCol >= 0 ? colName(mapping.sourceCol) : 'không có'}.`,
      errors.length ? `⚠️ Có ${errors.length} dòng chưa nạp được. Ví dụ: dòng ${errors[0].row} - ${errors[0].reason}.` : 'Không có dòng lỗi.'
    ].join('\n');
    setStatus(msg);
  } catch (err) {
    console.error(err);
    setStatus('❌ ' + (err.message || err), 'bad');
  }
}

function colName(idx) {
  if (idx < 0) return '';
  let s = '', n = idx + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function idle() { return new Promise(res => setTimeout(res, 20)); }

function mulberry32(seed) {
  let a = seed >>> 0;
  return function() {
    a += 0x6D2B79F5;
    let t = a;
    t = Math.imul(t ^ t >>> 15, t | 1);
    t ^= t + Math.imul(t ^ t >>> 7, t | 61);
    return ((t ^ t >>> 14) >>> 0) / 4294967296;
  };
}
function hashSeed(str) {
  const s = String(str || Date.now() + ':' + Math.random());
  let h = 2166136261;
  for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); }
  return h >>> 0;
}
function shuffled(arr, rand) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(rand() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; }
  return a;
}

function startQuiz() {
  if (!state.bank.length) return;
  const wanted = Math.max(1, Math.min(Number(els.count.value || 1), state.bank.length));
  const seedText = els.seed.value.trim() || String(Date.now());
  const rand = mulberry32(hashSeed(seedText));
  let qs = els.shuffleQ.checked ? shuffled(state.bank, rand) : state.bank.slice();
  qs = qs.slice(0, wanted);
  state.quiz = qs.map((q, qi) => {
    const opts = q.options.map((text, idx) => ({text, sourceIndex: idx, isCorrect: idx === q.correctIndex}));
    const finalOpts = els.shuffleO.checked ? shuffled(opts, rand) : opts;
    return {no: qi + 1, item: q, options: finalOpts, userChoice: null};
  });
  state.submitted = false; state.lastResult = null;
  renderQuiz();
  els.btnSubmit.disabled = false; els.btnSubmitSticky.disabled = false;
  els.resultSummary.textContent = 'Chưa nộp bài.';
  els.resultList.innerHTML = '';
  setStatus(`✅ Đã tạo đề ${wanted} câu. Seed mã đề: ${seedText}.`);
  location.hash = '#quizSection';
}

function renderQuiz() {
  if (!state.quiz.length) {
    els.quizInfo.textContent = 'Chưa tạo đề.'; els.quizList.innerHTML = ''; updateProgress(); return;
  }
  els.quizInfo.innerHTML = `Đề gồm <b>${state.quiz.length}</b> câu. ${els.shuffleQ.checked ? 'Có đảo câu hỏi.' : 'Không đảo câu hỏi.'} ${els.shuffleO.checked ? 'Có đảo phương án.' : 'Không đảo phương án.'}`;
  els.quizList.innerHTML = state.quiz.map((q, qi) => `
    <article class="question-card" data-q="${qi}">
      <div class="question-title">Câu ${q.no}. ${escapeHtml(q.item.question)}</div>
      ${q.options.map((op, oi) => {
        const letter = 'ABCDEF'[oi] || String(oi + 1);
        let cls = 'option';
        if (state.submitted) {
          if (op.isCorrect) cls += ' right';
          if (q.userChoice === oi && !op.isCorrect) cls += ' wrong';
        } else if (q.userChoice === oi) cls += ' chosen';
        return `<label class="${cls}"><input type="radio" name="q${qi}" value="${oi}" ${q.userChoice===oi?'checked':''} ${state.submitted?'disabled':''}><span class="letter">${letter}</span><span>${escapeHtml(op.text)}</span></label>`;
      }).join('')}
      <div class="muted small">Nguồn dòng Excel: ${escapeHtml(q.item.sourceRow)}${q.item.source ? ' | ' + escapeHtml(q.item.source) : ''}</div>
    </article>`).join('');
  els.quizList.querySelectorAll('input[type=radio]').forEach(inp => {
    inp.addEventListener('change', ev => {
      const card = ev.target.closest('.question-card');
      state.quiz[Number(card.dataset.q)].userChoice = Number(ev.target.value);
      renderQuiz(); updateProgress();
    });
  });
  updateProgress();
}

function updateProgress() {
  const total = state.quiz.length, done = state.quiz.filter(q => q.userChoice !== null).length;
  els.progressText.textContent = `${done}/${total}`;
  els.progressBar.style.width = total ? `${done / total * 100}%` : '0%';
}

function submitQuiz() {
  if (!state.quiz.length) return;
  const unanswered = state.quiz.filter(q => q.userChoice === null).length;
  if (unanswered && !confirm(`Còn ${unanswered} câu chưa chọn. Vẫn nộp bài?`)) return;
  let correct = 0;
  state.quiz.forEach(q => { if (q.userChoice !== null && q.options[q.userChoice]?.isCorrect) correct++; });
  const total = state.quiz.length;
  state.submitted = true;
  state.lastResult = {correct, total, score10: total ? correct / total * 10 : 0, time: new Date().toLocaleString('vi-VN')};
  renderQuiz(); renderResult();
  els.btnSubmit.disabled = true; els.btnSubmitSticky.disabled = true;
  location.hash = '#resultSection';
}

function renderResult() {
  const r = state.lastResult;
  if (!r) return;
  els.resultSummary.innerHTML = `<span class="pill">Đúng ${r.correct}/${r.total}</span> <span class="pill">Điểm ${(r.score10).toFixed(2)}/10</span> <span class="pill">${escapeHtml(r.time)}</span>`;
  els.resultList.innerHTML = state.quiz.map((q, qi) => {
    const chosen = q.userChoice === null ? null : q.options[q.userChoice];
    const correctOpt = q.options.find(o => o.isCorrect);
    const ok = chosen && chosen.isCorrect;
    const optionRows = q.options.map((op, oi) => {
      const letter = 'ABCDEF'[oi] || String(oi + 1);
      const status = op.isCorrect ? '<span class="ok">Đáp án đúng</span>' : (q.userChoice === oi ? '<span class="bad">Bạn đã chọn</span>' : '<span class="muted">Phương án sai</span>');
      const explain = op.isCorrect
        ? `<div class="analysis-row source"><b>Vì sao đúng:</b> Đây là phương án được cột đáp án của Excel xác định là đúng.${q.item.source ? '<br><b>Căn cứ:</b> ' + escapeHtml(q.item.source) : ''}</div>`
        : `<div class="analysis-row">${els.explain.checked ? explainDifference(correctOpt.text, op.text) : 'Đã tắt phân tích tự động.'}</div>`;
      return `<tr><td class="nowrap"><b>${letter}</b></td><td>${escapeHtml(op.text)}</td><td>${status}</td><td>${explain}</td></tr>`;
    }).join('');
    return `<article class="question-card">
      <h3>Câu ${qi + 1}. ${ok ? '<span class="ok">Đúng</span>' : '<span class="bad">Sai / chưa chọn</span>'}</h3>
      <p><b>${escapeHtml(q.item.question)}</b></p>
      <p>Chọn của bạn: <b>${chosen ? escapeHtml(chosen.text) : 'Chưa chọn'}</b><br>Đáp án đúng: <b class="ok">${escapeHtml(correctOpt.text)}</b></p>
      <div class="table-wrap"><table><thead><tr><th>PA</th><th>Nội dung</th><th>Trạng thái</th><th>Giải thích khác nhau so với đáp án đúng</th></tr></thead><tbody>${optionRows}</tbody></table></div>
    </article>`;
  }).join('');
}

async function saveBank() {
  if (!state.bank.length) return;
  await idbPut(BANK_KEY, {bank: state.bank, savedAt: new Date().toISOString(), fileName: state.fileName || '', mapping: state.mapping});
  setStatus(`✅ Đã lưu ${state.bank.length} câu hỏi vào bộ nhớ trình duyệt. Lần sau có thể bấm “Dùng bộ nhớ đã lưu”.`);
}
async function loadBank() {
  const data = await idbGet(BANK_KEY);
  if (!data || !data.bank) { setStatus('Chưa có ngân hàng đã lưu trong bộ nhớ trình duyệt.'); return; }
  state.bank = data.bank; state.errors = []; state.mapping = data.mapping || null; state.fileName = data.fileName || 'Bộ nhớ đã lưu';
  renderStats(); renderPreview();
  setStatus(`✅ Đã nạp ${state.bank.length} câu hỏi từ bộ nhớ trình duyệt. Thời điểm lưu: ${data.savedAt ? new Date(data.savedAt).toLocaleString('vi-VN') : 'không rõ'}.`);
}
async function clearBank() {
  if (!confirm('Xóa ngân hàng câu hỏi đã lưu trong bộ nhớ trình duyệt?')) return;
  await idbDel(BANK_KEY);
  setStatus('Đã xóa ngân hàng đã lưu. Dữ liệu đang mở trên màn hình không bị xóa cho đến khi tải lại trang.');
}

function registerSW() {
  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('./sw.js').catch(err => console.warn('SW register failed:', err));
  }
}

els.file.addEventListener('change', handleFileSelected);
els.btnParse.addEventListener('click', () => state.workbook ? parseSelectedSheet() : handleFileSelected());
els.sheet.addEventListener('change', parseSelectedSheet);
els.btnStart.addEventListener('click', startQuiz);
els.btnSubmit.addEventListener('click', submitQuiz);
els.btnSubmitSticky.addEventListener('click', submitQuiz);
els.btnSave.addEventListener('click', () => saveBank().catch(e => setStatus('❌ Không lưu được: ' + e.message, 'bad')));
els.btnLoad.addEventListener('click', () => loadBank().catch(e => setStatus('❌ Không đọc bộ nhớ: ' + e.message, 'bad')));
els.btnClear.addEventListener('click', () => clearBank().catch(e => setStatus('❌ Không xóa bộ nhớ: ' + e.message, 'bad')));
els.btnPrint.addEventListener('click', () => window.print());
registerSW();
renderStats();
