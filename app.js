'use strict';

const APP_VERSION = 'V17.0-20260613-legal-reference-from-source-column';
const DB_NAME = 'excel_quiz_offline_v3_fixed';
const STORE_NAME = 'kv';
const BANK_KEY = 'active_question_bank';
const STUDY_PROGRESS_KEY = 'excel_quiz_study_progress_v13';
const STUDY_ACTIVE_KEY = 'excel_quiz_study_active_ids_v13';
const STUDY_CONFIG_KEY = 'excel_quiz_study_config_v13';
const $ = (id) => document.getElementById(id);

const els = {};
const state = {
  bank: [],
  errors: [],
  meta: {},
  workbook: null,
  fileName: '',
  currentSheet: '',
  currentRows: [],
  currentMapping: null,
  quiz: [],
  submitted: false,
  lastResult: null,
  lastQuizConfig: null,
  mode: 'exam',
  study: { progress: {}, activeIds: [], config: { batchSize: 10, threshold: 10, shuffleQuestions: true, shuffleOptions: true } }
};

const VI_STOPWORDS = new Set(`
la là cua của va và hoac hoặc de để den đến duoc được bi bị trong ngoai ngoài mot một cac các nhung những ma mà thi thì voi với cho ve về theo tai tại tu từ khi luc lúc nao nào gi gì do đó nay này kia ay ấy tren trên duoi dưới vao vào ra bang bằng nhu như neu nếu hon hơn kem kém da đã dang đang se sẽ can cần phai phải khong không chua chưa cung cùng moi mỗi sau truoc trước phan phần noi nội dung cau câu hoi hỏi phuong phương an án lua lựa chon chọn dap đáp dung đúng sai don đơn vi vị linh lĩnh vuc vực he hệ thong thống hay boi bởi viec việc yeu yêu cau cầu quy quy dinh định
`.split(/\s+/).filter(Boolean));
VI_STOPWORDS.delete('phan'); // giữ được cụm kỹ thuật như "phân phối điện"

function initElements(){
  ['embeddedInfo','status','bankStats','bankPreview','excelFile','sheetSelect','btnReadSheet','btnReadAll','btnUseEmbedded','btnSaveEmbedded','btnSaveBank','btnLoadBank','btnClearBank','manualBox','manualHeaderRow','manualQuestionCol','manualAnswerCol','manualSourceCol','manualOptionCols','btnRefreshMapping','btnApplyMapping','internetQuestionSelect','internetSearchTemplate','internetProposedSource','internetStatus','internetApiUrl','btnInternetRefreshList','btnInternetAutoBuildSource','btnInternetEnrichAllByLegalRef','btnInternetOpenSearch','btnInternetOpenOfficial','btnInternetFetchApi','btnInternetSuggestSource','btnInternetApplySource','btnInternetAppendSource','btnInternetSaveDb','btnInternetExportXlsx','btnInternetExportLinks','quizCount','shuffleQuestions','shuffleOptions','showAutoExplain','seedInput','btnStartQuiz','quizInfo','quizList','btnSubmitQuiz','btnSubmitSticky','btnExitFocus','btnExitFocus2','btnSubmitQuizTop','btnScrollTopQuiz','btnScrollTopSticky','btnExitSticky','btnBackToSetup','btnBackToSetupSticky','btnNewQuizResult','btnNewQuizSticky','resultSummary','resultList','progressText','progressBar','btnPrint','btnClearOldCache','btnForceUpdatePWA','studyBatchSize','studyMasterThreshold','studyShuffleQuestions','studyShuffleOptions','btnStartStudy','btnNextStudy','btnResetStudy','btnMarkStudyAllLearned','btnSaveStudyProgress','btnSaveStudyProgressSticky','btnExportStudyStats','btnExportStudyWrongs','studyStats','btnExportResultHtml'].forEach(id => els[id] = $(id));
}

function setStatus(message, type='info'){
  if(!els.status) return;
  els.status.className = 'status ' + type;
  els.status.textContent = message;
}
function escapeHtml(v){ return String(v ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch])); }
function visibleText(v){ return String(v ?? '').replace(/\r\n/g,'\n').replace(/\s+/g,' ').trim(); }
function norm(v){
  return String(v ?? '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/đ/g,'d').replace(/Đ/g,'D')
    .toLowerCase().replace(/[“”‘’]/g,' ')
    .replace(/[_.,;:!?()[\]{}<>/\\|+*=~`^"']/g,' ')
    .replace(/-/g,' ')
    .replace(/\s+/g,' ').trim();
}
function tokens(v){ return norm(v).split(/\s+/).filter(t => t.length > 1 && !VI_STOPWORDS.has(t)); }
function unique(a){ return [...new Set(a.filter(Boolean))]; }
function intersection(a,b){ const B=new Set(b); return a.filter(x=>B.has(x)); }
function difference(a,b){ const B=new Set(b); return a.filter(x=>!B.has(x)); }
function tagList(items){ return unique(items).slice(0,14).map(x=>`<span class="tag">${escapeHtml(x)}</span>`).join(' '); }
function phrases(tks){ const out=[]; for(let n=2;n<=3;n++){ for(let i=0;i<=tks.length-n;i++) out.push(tks.slice(i,i+n).join(' ')); } return unique(out); }
function jaccard(a,b){ const A=new Set(a), B=new Set(b), U=new Set([...A,...B]); if(!U.size) return 0; let c=0; A.forEach(x=>{ if(B.has(x)) c++; }); return c/U.size; }
function extractNumbers(v){ const src=norm(v); const re=/(?:tren|duoi|den|tu|khong qua|toi thieu|toi da|lon hon|nho hon|bang|±)?\s*[-+]?\d+(?:[,.]\d+)?\s*(?:kv|v|a|ka|kw|mw|mva|kva|hz|%|phan tram|phut|gio|ngay|thang|nam|m|km)?/g; return unique((src.match(re)||[]).map(s=>s.replace(/\s+/g,' ').trim())); }
function extractRelations(v){ const src=norm(v); return ['tren','duoi','den','tu','khong qua','toi thieu','toi da','lon hon','nho hon','bang','khong','chua','cam','phai','duoc','cho phep','khong cho phep','truoc','sau','trong','ngoai','dong','cat','mo','cap','ha','cao','sieu cao','ngung','giam','tang'].filter(k => src.includes(k)); }
function textWords(v, keepStop=false){
  return String(v ?? '')
    .replace(/[“”‘’]/g,' ')
    .replace(/[_.,;:!?()[\]{}<>/\\|+*=~`^"']/g,' ')
    .replace(/-/g,' ')
    .split(/\s+/)
    .map(raw => ({raw: raw.trim(), key: norm(raw)}))
    .filter(w => w.raw && w.key && w.key.length > 1 && !/^\d/.test(w.key) && (keepStop || !VI_STOPWORDS.has(w.key)));
}
function keyTerms(v){
  const full = textWords(v, true);
  const filtered = textWords(v, false);
  const out = [];
  for(let n=5;n>=2;n--){
    for(let i=0;i<=full.length-n;i++){
      const part = full.slice(i,i+n);
      if(part.every(x => VI_STOPWORDS.has(x.key))) continue;
      const useful = part.filter(x => !VI_STOPWORDS.has(x.key));
      if(!useful.length) continue;
      const key = part.map(x=>x.key).join(' ');
      const raw = part.map(x=>x.raw).join(' ');
      if(key.length >= 8) out.push({key, raw});
    }
  }
  filtered.forEach(x => { if(x.key.length >= 3) out.push({key:x.key, raw:x.raw}); });
  const seen = new Set();
  return out.filter(t => {
    if(seen.has(t.key)) return false;
    seen.add(t.key); return true;
  });
}
function termListHtml(items, cls=''){
  return unique(items.map(x => typeof x === 'string' ? x : x.raw)).slice(0,10).map(x=>`<span class="tag ${cls}">${escapeHtml(x)}</span>`).join(' ');
}
function extractFacts(v){
  const raw = String(v ?? '');
  const src = raw.toLowerCase();
  const relationWords = '(không\\s+vượt\\s+quá|không\\s+quá|vượt\\s+quá|tối\\s+thiểu|tối\\s+đa|lớn\\s+hơn|nhỏ\\s+hơn|trên|dưới|đến|từ|bằng|±)';
  const unitWords = '(kv|v|ka|a|mw|kw|mva|kva|hz|%|phần\\s*trăm|phút|giờ|ngày|tháng|năm|km|m)';
  const re = new RegExp(`(?:${relationWords}\\s*)?[-+]?\\d+(?:[,.]\\d+)?\\s*(?:${unitWords})?`, 'gi');
  const out = [];
  let m;
  while((m = re.exec(src))){
    let s = m[0].replace(/\s+/g,' ').trim();
    if(!s) continue;
    const num = (s.match(/[-+]?\d+(?:[,.]\d+)?/)||[''])[0].replace(',', '.');
    const unit = (s.match(new RegExp(unitWords,'i'))||[''])[0].replace(/\s+/g,' ').trim();
    const rel = (s.match(new RegExp(relationWords,'i'))||[''])[0].replace(/\s+/g,' ').trim();
    const key = [rel, String(parseFloat(num)), unit.toLowerCase()].filter(Boolean).join('|');
    const looseKey = [String(parseFloat(num)), unit.toLowerCase()].filter(Boolean).join('|');
    out.push({raw:s, key, looseKey, num:parseFloat(num), unit:unit.toLowerCase(), rel});
  }
  const seen = new Set();
  return out.filter(f => { if(seen.has(f.key)) return false; seen.add(f.key); return true; });
}
function extractPolarity(v){
  const src = norm(v);
  const pairs = [
    ['cho phep','khong cho phep'], ['duoc','khong duoc'], ['phai','khong phai'], ['co','khong co'],
    ['dong','cat'], ['mo','dong'], ['tang','giam'], ['truoc','sau'], ['trong','ngoai'],
    ['truc tiep','cach ly'], ['so lech','qua dong'], ['van hanh','ngung']
  ];
  const found=[];
  pairs.forEach(p => p.forEach(x => { if(src.includes(x)) found.push(x); }));
  return unique(found);
}
function chooseShort(items, max=6){
  const arr = items.slice().sort((a,b) => (b.raw||b).length - (a.raw||a).length);
  const chosen=[];
  for(const it of arr){
    const key = typeof it === 'string' ? it : it.key;
    if(chosen.some(c => key.includes(typeof c === 'string' ? c : c.key) || (typeof c !== 'string' && c.key.includes(key)))) continue;
    chosen.push(it); if(chosen.length >= max) break;
  }
  return chosen;
}
function chooseContrastTerms(terms, otherTerms, max=6){
  const otherWords = new Set();
  otherTerms.forEach(t => String(t.key || '').split(' ').forEach(w => otherWords.add(w)));
  const scored = terms.map(t => {
    const words = String(t.key || '').split(' ').filter(Boolean);
    const uniqueWords = words.filter(w => !otherWords.has(w) && !VI_STOPWORDS.has(w));
    return {term:t, uniqueCount: uniqueWords.length, len:String(t.raw||'').length};
  }).filter(x => x.uniqueCount > 0);
  scored.sort((a,b) => (b.uniqueCount-a.uniqueCount) || (b.len-a.len));
  const chosen=[];
  for(const x of scored){
    const key = x.term.key;
    if(chosen.some(c => key.includes(c.key) || c.key.includes(key))) continue;
    chosen.push(x.term);
    if(chosen.length >= max) break;
  }
  return chosen;
}
function normSeqToken(raw){
  let s = String(raw ?? '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/đ/g,'d').replace(/Đ/g,'D')
    .toLowerCase().trim();
  s = s.replace(/,/g,'.').replace(/\s+/g,'');
  s = s.replace(/[^a-z0-9.%+\-]/g,'');
  return s;
}
function sequenceTokens(v, keepStopwords=true){
  const src = String(v ?? '');
  const re = /[-+]?\d+(?:[,.]\d+)?\s*(?:mva|kva|kv|ka|kw|mw|hz|v|a|%|phút|giờ|giờ|ngày|tháng|năm|km|m)?|[A-Za-zÀ-ỹĐđ]+/gi;
  const out = [];
  let m;
  while((m = re.exec(src))){
    const raw = m[0].replace(/\s+/g,' ').trim();
    const key = normSeqToken(raw);
    if(!key) continue;
    if(!keepStopwords && VI_STOPWORDS.has(key)) continue;
    if(key.length <= 1 && !/\d/.test(key)) continue;
    out.push({raw, key});
  }
  return out;
}
function htmlTermTokenList(items, cls=''){
  return (items || []).slice(0,12).map(x => `<span class="tag ${cls}">${escapeHtml(typeof x === 'string' ? x : x.raw)}</span>`).join(' ');
}
function seqText(items, max=16){
  const arr = (items || []).map(x => typeof x === 'string' ? x : x.raw).filter(Boolean);
  if(!arr.length) return '';
  const head = arr.slice(0, max).join(' ');
  return arr.length > max ? head + ' …' : head;
}
function lcsSequence(correctSeq, wrongSeq){
  const n = correctSeq.length, m = wrongSeq.length;
  const dp = Array.from({length:n+1}, () => Array(m+1).fill(0));
  for(let i=n-1;i>=0;i--){
    for(let j=m-1;j>=0;j--){
      dp[i][j] = correctSeq[i].key === wrongSeq[j].key ? dp[i+1][j+1] + 1 : Math.max(dp[i+1][j], dp[i][j+1]);
    }
  }
  const ops = [];
  let i=0, j=0;
  while(i<n || j<m){
    if(i<n && j<m && correctSeq[i].key === wrongSeq[j].key){
      ops.push({type:'match', c:correctSeq[i], w:wrongSeq[j], cIndex:i, wIndex:j}); i++; j++;
    } else if(j>=m || (i<n && dp[i+1][j] >= dp[i][j+1])){
      ops.push({type:'del', c:correctSeq[i], cIndex:i, wIndex:j}); i++;
    } else {
      ops.push({type:'add', w:wrongSeq[j], cIndex:i, wIndex:j}); j++;
    }
  }
  const blocks = [];
  let cur = null;
  function flush(){ if(cur && (cur.del.length || cur.add.length)){ blocks.push(cur); } cur = null; }
  for(const op of ops){
    if(op.type === 'match'){ flush(); continue; }
    if(!cur) cur = {cStart: op.cIndex, wStart: op.wIndex, del:[], add:[]};
    if(op.type === 'del') cur.del.push(op.c);
    if(op.type === 'add') cur.add.push(op.w);
  }
  flush();
  return {lcsLen: dp[0][0], ops, blocks};
}
function prefixCount(a,b){ let i=0; while(i<a.length && i<b.length && a[i].key === b[i].key) i++; return i; }
function suffixCount(a,b,prefix=0){
  let i=a.length-1, j=b.length-1, c=0;
  while(i>=prefix && j>=prefix && a[i].key === b[j].key){ c++; i--; j--; }
  return c;
}
function sequenceExplainHtml(correct, wrong){
  const cSeq = sequenceTokens(correct, true);
  const wSeq = sequenceTokens(wrong, true);
  if(!cSeq.length || !wSeq.length) return '';
  const al = lcsSequence(cSeq, wSeq);
  const pctWrongMatched = Math.round((al.lcsLen / Math.max(1, wSeq.length)) * 100);
  const pctCorrectCovered = Math.round((al.lcsLen / Math.max(1, cSeq.length)) * 100);
  const pctBalanced = Math.round((2 * al.lcsLen / Math.max(1, cSeq.length + wSeq.length)) * 100);
  const pref = prefixCount(cSeq, wSeq);
  const cNext = cSeq[pref];
  const wNext = wSeq[pref];
  const correctStart = cSeq.slice(0, Math.min(4, cSeq.length));
  const wrongStart = wSeq.slice(0, Math.min(4, wSeq.length));
  const samePrefix = cSeq.slice(0, Math.min(pref, 10));
  const cRest = cSeq.slice(pref, Math.min(pref + 10, cSeq.length));
  const wRest = wSeq.slice(pref, Math.min(pref + 10, wSeq.length));

  const lines = [];
  lines.push(`<div class="diff-block sequence"><b>So khớp từ đầu đến cuối:</b> phương án sai giống <b>${pctBalanced}%</b> so với đáp án đúng. Nó trùng <b>${pctWrongMatched}%</b> chữ của chính nó và bao phủ <b>${pctCorrectCovered}%</b> ý của đáp án đúng.</div>`);

  if(pref === 0){
    lines.push(`<div class="diff-block focus-sequence"><b>Sai ngay từ đầu chuỗi:</b><br>Đáp án đúng bắt đầu bằng: ${htmlTermTokenList(correctStart, 'need')}.<br>Phương án sai không có phần bắt đầu này, mà bắt đầu bằng: ${htmlTermTokenList(wrongStart, 'wrongterm') || '<span class="muted">không có nội dung</span>'}.<br><b>Cách nhận biết nhanh:</b> chỉ cần nhìn từ/cụm mở đầu đã thấy phương án sai không đi theo đúng đáp án.</div>`);
  } else if(cNext && wNext){
    lines.push(`<div class="diff-block focus-sequence"><b>Điểm sai đầu tiên khi đọc từ đầu đến cuối:</b><br>Hai phương án giống nhau đến trước từ/cụm số <b>${pref + 1}</b> ${samePrefix.length ? '(' + htmlTermTokenList(samePrefix, 'same') + ')' : ''}.<br>Đến vị trí này, đáp án đúng phải có: ${htmlTermTokenList([cNext], 'need')}<br>Phương án sai lại ghi: ${htmlTermTokenList([wNext], 'wrongterm')}<br><b>Kết luận nhanh:</b> từ/cụm sai đầu tiên này là điểm tách phương án sai khỏi đáp án đúng.</div>`);
  } else if(!wNext && cNext){
    lines.push(`<div class="diff-block focus-sequence"><b>Phương án sai bị thiếu phần cuối:</b> sau ${pref} từ/cụm đầu đã giống nhau, phương án sai dừng lại hoặc thiếu nội dung mà đáp án đúng còn yêu cầu: ${htmlTermTokenList(cRest, 'need')}.</div>`);
  } else if(wNext && !cNext){
    lines.push(`<div class="diff-block focus-sequence"><b>Phương án sai thêm phần không cần có:</b> đáp án đúng đã kết thúc, nhưng phương án sai còn thêm: ${htmlTermTokenList(wRest, 'wrongterm')}.</div>`);
  } else {
    lines.push(`<div class="diff-block focus-sequence"><b>Chuỗi chữ gần như giống hoàn toàn:</b> nếu vẫn bị xác định sai, hãy kiểm tra dấu câu, khoảng trắng, ký hiệu hoặc cột đáp án đúng trong Excel.</div>`);
  }

  if(al.blocks.length){
    const ordered = al.blocks.slice(0, 10).map((b,idx) => {
      const a = seqText(b.add, 10) || '∅';
      const d = seqText(b.del, 10) || '∅';
      return `<div class="sequence-point"><b>${idx+1}.</b> Phương án sai ghi: <span class="tag wrongterm">${escapeHtml(a)}</span><br>Đáp án đúng cần có: <span class="tag need">${escapeHtml(d)}</span></div>`;
    }).join('');
    lines.push(`<div class="diff-block ordered-diff"><b>Các điểm sai theo đúng thứ tự đọc từ đầu đến cuối:</b>${ordered}</div>`);
  }
  return lines.join('');
}
function explainDifference(correct, wrong){
  const correctFacts = extractFacts(correct), wrongFacts = extractFacts(wrong);
  const sameFacts = correctFacts.filter(c => wrongFacts.some(w => w.looseKey === c.looseKey));
  const missingFacts = correctFacts.filter(c => !wrongFacts.some(w => w.looseKey === c.looseKey));
  const addedFacts = wrongFacts.filter(w => !correctFacts.some(c => c.looseKey === w.looseKey));

  const correctTerms = keyTerms(correct), wrongTerms = keyTerms(wrong);
  const sameTermKeys = new Set(wrongTerms.map(t=>t.key));
  const correctTermKeys = new Set(correctTerms.map(t=>t.key));
  const sameTerms = chooseShort(correctTerms.filter(t => sameTermKeys.has(t.key)), 5);
  const missingTerms = chooseContrastTerms(correctTerms.filter(t => !sameTermKeys.has(t.key)), wrongTerms, 7);
  const addedTerms = chooseContrastTerms(wrongTerms.filter(t => !correctTermKeys.has(t.key)), correctTerms, 7);

  const correctPol = extractPolarity(correct), wrongPol = extractPolarity(wrong);
  const missingPol = correctPol.filter(x => !wrongPol.includes(x));
  const addedPol = wrongPol.filter(x => !correctPol.includes(x));

  const proofItems = [...missingFacts.map(f=>f.raw), ...missingTerms, ...missingPol];
  const wrongItems = [...addedFacts.map(f=>f.raw), ...addedTerms, ...addedPol];
  let conclusion = '';
  if(missingFacts.length || addedFacts.length){
    conclusion = `<b>Nhìn từ phương án sai:</b> sai chủ yếu vì <b>số liệu/mốc/phạm vi không khớp</b>. Đáp án đúng có ${termListHtml(missingFacts.map(f=>f.raw), 'need') || 'mốc đúng'}, còn phương án sai dùng ${termListHtml(addedFacts.map(f=>f.raw), 'wrongterm') || 'mốc khác/thiếu mốc đúng'}.`;
  } else if(missingPol.length || addedPol.length){
    conclusion = `<b>Nhìn từ phương án sai:</b> sai do <b>đổi điều kiện hoặc tính chất</b>. Đáp án đúng giữ điều kiện ${termListHtml(missingPol,'need')||'cần có'}, còn phương án sai thể hiện ${termListHtml(addedPol,'wrongterm')||'điều kiện khác/thiếu điều kiện đó'}.`;
  } else if(missingTerms.length || addedTerms.length){
    conclusion = `<b>Nhìn từ phương án sai:</b> sai do thiếu hoặc đổi <b>từ khóa/đối tượng quyết định</b>. Đáp án đúng có ${termListHtml(missingTerms,'need')||'ý chính cần có'}, còn phương án sai chuyển sang ${termListHtml(addedTerms,'wrongterm')||'ý khác hoặc thiếu ý chính đó'}.`;
  } else {
    conclusion = '<b>Nhìn từ phương án sai:</b> hai phương án rất gần nhau; điểm đúng/sai có thể nằm ở một từ nhỏ, dấu phủ định, ký hiệu hoặc căn cứ pháp lý.';
  }

  const visible = sequenceExplainHtml(correct, wrong);
  const extra = [];
  extra.push(`<div class="diff-conclusion">${conclusion}</div>`);
  if(proofItems.length){
    extra.push(`<div class="diff-block"><b>Đáp án đúng có gì mà phương án sai không có:</b> ${termListHtml(proofItems, 'need')}.</div>`);
  }
  if(wrongItems.length){
    extra.push(`<div class="diff-block"><b>Phương án sai dùng/thêm nội dung nào:</b> ${termListHtml(wrongItems, 'wrongterm')}.</div>`);
  }
  if(sameFacts.length || sameTerms.length){
    extra.push(`<div class="diff-block"><b>Phần giống nhau dễ gây nhầm:</b> ${termListHtml([...sameFacts.map(f=>f.raw), ...sameTerms], 'same') || 'rất ít nội dung trùng nhau'}.</div>`);
  }
  const details = `<details class="more-analysis"><summary>Chạm để đọc thêm phân tích số liệu, từ khóa và phần giống nhau</summary>${extra.join('')}</details>`;
  return [visible, details].filter(Boolean).join('');
}

function cloneData(obj){ return JSON.parse(JSON.stringify(obj)); }
function parseEmbedded(){
  const node = $('embedded-bank');
  if(!node) throw new Error('Không tìm thấy khối dữ liệu nhúng trong HTML.');
  const payload = JSON.parse(node.textContent || '{}');
  const bank = Array.isArray(payload.bank) ? payload.bank : [];
  state.meta = payload.meta || {};
  state.bank = normalizeBank(bank);
  state.errors = [];
  state.fileName = state.meta.sourceFile || 'Dữ liệu nhúng';
  state.currentSheet = state.meta.sheetName || 'Dữ liệu nhúng';
  state.currentRows = [];
  state.currentMapping = {source:'embedded'};
  if(els.embeddedInfo){
    els.embeddedInfo.innerHTML = `<b>${state.bank.length}</b> câu hỏi đã nhúng từ <b>${escapeHtml(state.fileName)}</b>. Mở file là dùng ngay, không cần tải Excel.`;
  }
  renderAll();
}
function normalizeBank(bank){
  return bank.map((q,i) => {
    const options = (q.options || []).map(visibleText).filter(Boolean);
    let idx = Number.isInteger(q.correctIndex) ? q.correctIndex : inferAnswerIndex(q.answerRaw, options);
    if(idx < 0 || idx >= options.length) idx = 0;
    return {
      id: visibleText(q.id) || String(i+1),
      sourceRow: q.sourceRow || '',
      sheetName: q.sheetName || '',
      question: visibleText(q.question),
      options,
      correctIndex: idx,
      correctText: options[idx] || visibleText(q.correctText),
      source: visibleText(q.source),
      answerRaw: visibleText(q.answerRaw)
    };
  }).filter(q => q.question && q.options.length >= 2 && q.correctText);
}

function renderAll(){ renderStats(); renderPreview(); renderInternetTools(); resetQuiz(); }
function renderStats(){
  if(!els.bankStats) return;
  const src = state.currentSheet || state.fileName || 'Dữ liệu';
  els.bankStats.innerHTML = `<span class="pill okp">${state.bank.length} câu hợp lệ</span><span class="pill ${state.errors.length?'warnp':''}">${state.errors.length} dòng lỗi</span><span class="pill">${escapeHtml(src)}</span><span class="pill">${APP_VERSION}</span>`;
  if(els.btnStartQuiz) els.btnStartQuiz.disabled = state.bank.length === 0;
  if(els.btnSaveBank) els.btnSaveBank.disabled = state.bank.length === 0;
  if(els.quizCount){ els.quizCount.max = String(Math.max(1,state.bank.length)); if(Number(els.quizCount.value || 0) > state.bank.length) els.quizCount.value = String(state.bank.length || 1); }
  renderStudyStats();
}

function getQuestionKey(q){
  return String(q?.id || norm(q?.question || '')).trim();
}
function loadStudyState(){
  try { state.study.progress = JSON.parse(localStorage.getItem(STUDY_PROGRESS_KEY) || '{}') || {}; } catch(e){ state.study.progress = {}; }
  try { state.study.activeIds = JSON.parse(localStorage.getItem(STUDY_ACTIVE_KEY) || '[]') || []; } catch(e){ state.study.activeIds = []; }
  try {
    const cfg = JSON.parse(localStorage.getItem(STUDY_CONFIG_KEY) || '{}') || {};
    state.study.config = Object.assign({batchSize:10, threshold:10, shuffleQuestions:true, shuffleOptions:true}, cfg);
  } catch(e){}
}
function saveStudyState(){
  try {
    localStorage.setItem(STUDY_PROGRESS_KEY, JSON.stringify(state.study.progress || {}));
    localStorage.setItem(STUDY_ACTIVE_KEY, JSON.stringify(state.study.activeIds || []));
    localStorage.setItem(STUDY_CONFIG_KEY, JSON.stringify(state.study.config || {}));
  } catch(e){ console.warn('Không lưu được tiến độ ôn tập', e); }
}
function studyThreshold(){ return Math.max(1, Number(els.studyMasterThreshold?.value || state.study.config.threshold || 10)); }
function studyBatchSize(){ return Math.max(1, Math.min(Number(els.studyBatchSize?.value || state.study.config.batchSize || 10), Math.max(1,state.bank.length))); }
function isStudyLearned(key){
  const p = state.study.progress[key] || {};
  return !!p.learned || Number(p.correctCount || 0) >= studyThreshold();
}
function syncStudyConfigFromUI(){
  state.study.config = {
    batchSize: studyBatchSize(),
    threshold: studyThreshold(),
    shuffleQuestions: !!els.studyShuffleQuestions?.checked,
    shuffleOptions: !!els.studyShuffleOptions?.checked
  };
  saveStudyState();
}

function getStudyRows(){
  return state.bank.map((q, idx) => {
    const key = getQuestionKey(q);
    const p = state.study.progress[key] || {};
    const correctCount = Number(p.correctCount || 0);
    const wrongCount = Number(p.wrongCount || 0);
    const attempts = Number(p.attempts || 0);
    const learned = !!p.learned || correctCount >= studyThreshold();
    return {idx, key, q, p, correctCount, wrongCount, attempts, streak:Number(p.streak || 0), learned, manualLearned:!!p.manualLearned, lastAt:p.lastAt || '', learnedAt:p.learnedAt || ''};
  });
}
function saveStudyProgressManual(){
  syncStudyConfigFromUI();
  saveStudyState();
  renderStudyStats();
  setStatus('✅ Đã lưu tiến độ ôn tập vào bộ nhớ trình duyệt. Lần sau mở lại PWA/HTML có thể bấm “Lượt ôn tiếp” để học tiếp.', 'good');
}
function renderStudyStats(){
  if(!els.studyStats) return;
  if(els.studyBatchSize) els.studyBatchSize.value = String(state.study.config.batchSize || 10);
  if(els.studyMasterThreshold) els.studyMasterThreshold.value = String(state.study.config.threshold || 10);
  if(els.studyShuffleQuestions) els.studyShuffleQuestions.checked = !!state.study.config.shuffleQuestions;
  if(els.studyShuffleOptions) els.studyShuffleOptions.checked = !!state.study.config.shuffleOptions;
  const total = state.bank.length;
  let learned = 0, practicing = 0;
  const active = new Set(state.study.activeIds || []);
  for(const q of state.bank){ const k=getQuestionKey(q); if(isStudyLearned(k)) learned++; if(active.has(k) && !isStudyLearned(k)) practicing++; }
  const remain = Math.max(0,total-learned);
  const rows = getStudyRows();
  const attempts = rows.reduce((s,r)=>s+r.attempts,0);
  const wrongs = rows.reduce((s,r)=>s+r.wrongCount,0);
  els.studyStats.innerHTML = `<span class="pill okp">Đã học: ${learned}/${total}</span><span class="pill ${remain?'warnp':'okp'}">Còn lại: ${remain}</span><span class="pill">Đang ôn: ${practicing || Math.min(studyBatchSize(), remain)}</span><span class="pill">Đạt khi đúng ≥ ${studyThreshold()} lần</span><span class="pill">Tổng lượt làm: ${attempts}</span><span class="pill ${wrongs?'warnp':'okp'}">Tổng lượt sai: ${wrongs}</span>`;
}
function refreshStudyActiveIds(){
  const bankKeys = new Set(state.bank.map(getQuestionKey));
  const batch = studyBatchSize();
  let active = (state.study.activeIds || []).filter(k => bankKeys.has(k) && !isStudyLearned(k));
  const candidates = state.bank
    .map(q => ({q, key:getQuestionKey(q), p:state.study.progress[getQuestionKey(q)] || {}}))
    .filter(x => !active.includes(x.key) && !isStudyLearned(x.key))
    .sort((a,b) => (Number(a.p.correctCount||0)-Number(b.p.correctCount||0)) || (Number(a.p.attempts||0)-Number(b.p.attempts||0)));
  for(const c of candidates){ if(active.length >= batch) break; active.push(c.key); }
  state.study.activeIds = active.slice(0,batch);
  saveStudyState();
  renderStudyStats();
  return state.study.activeIds;
}
function updateStickyLabels(){
  document.body.classList.toggle('study-mode', state.mode === 'study');
  if(els.btnNewQuizSticky) els.btnNewQuizSticky.textContent = state.mode === 'study' ? 'Lượt tiếp' : 'Đề khác';
  if(els.btnBackToSetupSticky) els.btnBackToSetupSticky.textContent = state.mode === 'study' ? 'Thiết lập' : 'Thiết lập';
  if(els.btnSaveStudyProgressSticky) els.btnSaveStudyProgressSticky.title = state.mode === 'study' ? 'Lưu tiến độ ôn tập' : 'Lưu tiến độ';
}
function renderPreview(){
  if(!els.bankPreview) return;
  if(!state.bank.length){ els.bankPreview.innerHTML = '<table><tbody><tr><td class="muted">Chưa có dữ liệu hợp lệ.</td></tr></tbody></table>'; return; }
  const rows = state.bank.slice(0,50).map((q,i)=>`<tr><td class="nowrap">${i+1}</td><td>${escapeHtml(q.question)}</td><td>${escapeHtml(q.correctText)}</td><td>${escapeHtml(q.options.join(' | '))}</td><td>${escapeHtml(q.source||'')}</td></tr>`).join('');
  els.bankPreview.innerHTML = `<table><thead><tr><th>STT</th><th>Câu hỏi</th><th>Đáp án đúng</th><th>Các phương án</th><th>Căn cứ / giải thích</th></tr></thead><tbody>${rows}</tbody></table>`;
}


const LEGAL_DOC_META = {
  '05/2025/TT-BCT': {
    code:'05/2025/TT-BCT',
    title:'Thông tư 05/2025/TT-BCT - Quy định hệ thống truyền tải điện, phân phối điện và đo đếm điện năng',
    shortTitle:'Quy định hệ thống truyền tải điện, phân phối điện và đo đếm điện năng',
    issued:'01/02/2025',
    effective:'01/02/2025',
    official:'https://chinhphu.vn/?classid=1&docid=212774&orggroupid=4&pageid=27160',
    fullText:'https://thuvienphapluat.vn/van-ban/Thuong-mai/Thong-tu-05-2025-TT-BCT-he-thong-truyen-tai-dien-phan-phoi-dien-do-dem-dien-nang-642994.aspx',
    note:'Lưu ý: Thông tư 46/2025/TT-BCT có sửa đổi, bổ sung một số điều liên quan Thông tư 05/2025/TT-BCT từ 22/9/2025; khi dùng chính thức cần đối chiếu văn bản hợp nhất/cập nhật.'
  },
  '06/2025/TT-BCT': {
    code:'06/2025/TT-BCT',
    title:'Thông tư 06/2025/TT-BCT - Quy định điều độ, vận hành, thao tác, xử lý sự cố, khởi động đen và khôi phục hệ thống điện quốc gia',
    shortTitle:'Quy định điều độ, vận hành, thao tác, xử lý sự cố, khởi động đen và khôi phục hệ thống điện quốc gia',
    issued:'01/02/2025',
    effective:'01/02/2025',
    official:'https://vanban.chinhphu.vn/?classid=1&docid=212775&orggroupid=4&pageid=27160',
    fullText:'https://luatvietnam.vn/cong-nghiep/thong-tu-06-2025-tt-bct-van-hanh-thao-tac-xu-ly-su-co-khoi-dong-den-khoi-phuc-he-thong-dien-quoc-gia-388847-d1.html',
    note:'Lưu ý: Thông tư 46/2025/TT-BCT có sửa đổi, bổ sung một số điều liên quan Thông tư 06/2025/TT-BCT từ 22/9/2025; khi dùng chính thức cần đối chiếu văn bản hợp nhất/cập nhật.'
  }
};

function normalizeLegalCode(text){
  const m = String(text||'').match(/(\d{1,3})\s*\/\s*(\d{4})\s*\/\s*TT\s*-\s*BCT/i);
  return m ? `${String(m[1]).padStart(2,'0')}/${m[2]}/TT-BCT` : '';
}
function parseLegalReference(text){
  const src = visibleText(text || '');
  const code = normalizeLegalCode(src);
  const article = (src.match(/điều\s*(\d+)/i) || [,''])[1];
  const clause = (src.match(/khoản\s*(\d+)/i) || [,''])[1];
  const point = (src.match(/điểm\s*([a-zA-Zà-ỹ])/i) || [,''])[1];
  return {raw:src, code, article, clause, point, meta:LEGAL_DOC_META[code] || null};
}
function compactOriginalSource(src){
  return visibleText(src).replace(/\s*—\s*Nội dung trọng tâm:[\s\S]*$/i,'').replace(/\s*\|\s*Nội dung trọng tâm:[\s\S]*$/i,'').trim();
}
function buildLegalSourceFromColumn(q, includeNote=false){
  const original = compactOriginalSource(q?.sourceOriginal || q?.source || '');
  const parsed = parseLegalReference(original);
  const meta = parsed.meta;
  const base = original || 'Chưa có căn cứ pháp lý';
  const parts = [base];
  if(q && q.correctText) parts.push('Nội dung trọng tâm theo đáp án đúng: ' + visibleText(q.correctText));
  if(meta){
    parts.push('Văn bản: ' + meta.title);
    parts.push('Ban hành/hiệu lực: ' + meta.issued + ' / ' + meta.effective);
    parts.push('Nguồn chính thức: ' + meta.official);
    parts.push('Toàn văn tra cứu: ' + meta.fullText);
    if(includeNote) parts.push(meta.note);
  } else if(parsed.code){
    parts.push('Văn bản: ' + parsed.code + ' (chưa có metadata nhúng, cần tra cứu bổ sung)');
  }
  return parts.join(' — ');
}
function legalSearchPhrase(q){
  const parsed = parseLegalReference(q?.sourceOriginal || q?.source || '');
  const pieces = [];
  if(parsed.clause) pieces.push('khoản ' + parsed.clause);
  if(parsed.article) pieces.push('Điều ' + parsed.article);
  if(parsed.code) pieces.push('Thông tư ' + parsed.code);
  const base = pieces.join(' ');
  return (base || visibleText(q?.source || '') || visibleText(q?.question || '')).trim();
}
function autoBuildLegalSourceForSelected(){
  const q = currentInternetQuestion();
  if(!q){ setInternetStatus('Chưa chọn câu để tạo căn cứ.', 'bad'); return; }
  const val = buildLegalSourceFromColumn(q, true);
  if(els.internetProposedSource) els.internetProposedSource.value = val;
  setInternetStatus('✅ Đã tạo nội dung căn cứ từ chính cột căn cứ pháp lý của câu đang chọn. Kiểm tra rồi bấm “Cập nhật căn cứ” hoặc “Ghép thêm”.', 'good');
}
async function enrichAllSourcesFromLegalColumn(){
  if(!state.bank.length){ setInternetStatus('Chưa có ngân hàng câu hỏi để cập nhật.', 'bad'); return; }
  let changed = 0;
  state.bank = state.bank.map(q => {
    const original = compactOriginalSource(q.sourceOriginal || q.source || '');
    const enriched = buildLegalSourceFromColumn(Object.assign({}, q, {source: original, sourceOriginal: original}), false);
    if(!q.sourceOriginal) q.sourceOriginal = original;
    if(q.source !== enriched){ q.source = enriched; changed++; }
    return q;
  });
  state.meta = Object.assign({}, state.meta || {}, {updatedSourceAt:new Date().toISOString(), updatedSourceBy:'legal_source_column_v17', legalSourceMode:'from_last_excel_column'});
  renderStats(); renderPreview(); renderInternetTools();
  try { await saveBank('ngân hàng đã tự cập nhật căn cứ theo cột pháp lý', true); } catch(e){}
  setInternetStatus(`✅ Đã tự cập nhật căn cứ cho ${changed} câu dựa trên cột căn cứ pháp lý hiện có, đáp án đúng và metadata văn bản 05/2025/TT-BCT, 06/2025/TT-BCT.`, 'good');
}

function setInternetStatus(message, type='info'){
  if(!els.internetStatus) return;
  els.internetStatus.className = 'status ' + type;
  els.internetStatus.textContent = message;
}
function shortLabel(v, n=90){
  const t = visibleText(v);
  return t.length > n ? t.slice(0,n-1) + '…' : t;
}
function extractLawRefs(text){
  const src = visibleText(text);
  const refs = [];
  const re = /(?:khoản\s*\d+\s*)?(?:điều\s*\d+\s*)?(?:Thông\s*tư\s*)?\d{1,3}\s*\/\s*\d{4}\s*\/\s*TT\s*-\s*BCT/gi;
  let m;
  while((m = re.exec(src))) refs.push(m[0].replace(/\s+/g,' ').trim());
  const re2 = /(?:khoản\s*\d+\s*)?điều\s*\d+/gi;
  while((m = re2.exec(src))) refs.push(m[0].replace(/\s+/g,' ').trim());
  return unique(refs).slice(0,8);
}
function buildInternetQuery(q){
  if(!q) return '';
  const template = els.internetSearchTemplate?.value || '"{source}"';
  const lawRefs = extractLawRefs((q.sourceOriginal || q.source) || '').join(' ');
  const legalPhrase = legalSearchPhrase(q);
  return template
    .replaceAll('{question}', q.question || '')
    .replaceAll('{answer}', q.correctText || '')
    .replaceAll('{source}', q.sourceOriginal || compactOriginalSource(q.source) || '')
    .replaceAll('{law}', lawRefs).replaceAll('{legal}', legalPhrase)
    .replace(/\s+/g,' ').trim();
}
function currentInternetQuestion(){
  const idx = Number(els.internetQuestionSelect?.value || 0);
  if(!state.bank.length) return null;
  return state.bank[Math.max(0, Math.min(idx, state.bank.length-1))] || null;
}
function renderInternetTools(){
  if(!els.internetQuestionSelect) return;
  const previous = els.internetQuestionSelect.value;
  els.internetQuestionSelect.innerHTML = state.bank.map((q,i)=>`<option value="${i}">${i+1}. ${escapeHtml(shortLabel(q.sourceOriginal || q.source || 'Chưa có căn cứ',76))} — ${escapeHtml(shortLabel(q.question,48))}</option>`).join('') || '<option value="0">Chưa có câu hỏi</option>';
  if(previous && Number(previous) < state.bank.length) els.internetQuestionSelect.value = previous;
  updateInternetQuestionView();
}
function updateInternetQuestionView(){
  if(!els.internetStatus) return;
  const q = currentInternetQuestion();
  if(!q){ setInternetStatus('Chưa có ngân hàng câu hỏi để tra cứu căn cứ.', 'bad'); return; }
  const parsed = parseLegalReference(q.sourceOriginal || q.source || '');
  const lawRefs = extractLawRefs(q.sourceOriginal || q.source || '').join(', ') || 'chưa tách được ký hiệu văn bản';
  const docName = parsed.meta ? parsed.meta.shortTitle : (parsed.code || 'chưa nhận diện văn bản');
  setInternetStatus(`Câu đang chọn: ${q.id || ''} — căn cứ gốc: ${shortLabel(q.sourceOriginal || compactOriginalSource(q.source) || 'chưa có',120)} — văn bản: ${docName} — đáp án đúng: ${shortLabel(q.correctText,90)} — mốc tra cứu: ${lawRefs}`, 'info');
  if(els.internetProposedSource && !els.internetProposedSource.value.trim()){
    els.internetProposedSource.placeholder = `Căn cứ hiện có: ${q.sourceOriginal || compactOriginalSource(q.source) || 'Theo khoản ... điều ... Thông tư ...'}. Có thể bấm “Tạo căn cứ từ cột cuối” để sinh nội dung đề xuất, hoặc dán đoạn tra cứu được từ văn bản pháp lý.`;
  }
}
function internetSearchUrl(q, official=false){
  const query = buildInternetQuery(q);
  const officialFilter = official ? ' site:vanban.chinhphu.vn OR site:chinhphu.vn OR site:vbpl.vn OR site:eav.gov.vn OR site:thuvienphapluat.vn' : '';
  return 'https://www.google.com/search?q=' + encodeURIComponent(query + officialFilter);
}
function openInternetSearch(official=false){
  const q = currentInternetQuestion();
  if(!q){ setInternetStatus('Chưa chọn được câu hỏi để tra cứu.', 'bad'); return; }
  window.open(internetSearchUrl(q, official), '_blank', 'noopener,noreferrer');
  setInternetStatus(official ? 'Đã mở tra cứu trên các nguồn pháp lý/chính thống. Sau khi tìm được đoạn phù hợp, dán vào ô nội dung đề xuất rồi bấm cập nhật.' : 'Đã mở trang tìm kiếm Internet. Sau khi tìm được đoạn phù hợp, dán vào ô nội dung đề xuất rồi bấm cập nhật.', 'info');
}
function stripHtmlToText(s){
  return visibleText(String(s || '').replace(/<script[\s\S]*?<\/script>/gi,' ').replace(/<style[\s\S]*?<\/style>/gi,' ').replace(/<[^>]+>/g,' '));
}
function suggestSourceFromPasted(){
  const q = currentInternetQuestion();
  const pasted = visibleText(els.internetProposedSource?.value || '');
  if(!q){ setInternetStatus('Chưa chọn câu hỏi.', 'bad'); return; }
  if(!pasted){ setInternetStatus('Hãy dán đoạn văn bản/link tìm được rồi mới rút gọn căn cứ.', 'bad'); return; }
  const text = stripHtmlToText(pasted);
  const refs = extractLawRefs(text + ' ' + q.source);
  const qTerms = unique([...tokens(q.sourceOriginal || q.source), ...tokens(q.correctText)]).filter(t => t.length > 2);
  const sentences = text.split(/(?<=[.!?。\n])\s+|\n+/).map(visibleText).filter(Boolean);
  let scored = sentences.map((sen, idx) => {
    const st = norm(sen);
    let score = 0;
    qTerms.forEach(t => { if(st.includes(t)) score += 2; });
    refs.forEach(r => { if(norm(sen).includes(norm(r))) score += 4; });
    if(/khoản|điều|thông tư|tt-bct|quy định/i.test(sen)) score += 3;
    if(sen.length > 40 && sen.length < 800) score += 1;
    return {sen, score, idx};
  }).sort((a,b)=>b.score-a.score || a.idx-b.idx);
  const picked = scored.filter(x => x.score > 0).slice(0,3).map(x => x.sen);
  const prefix = refs.length ? refs.join('; ') : (q.source || 'Căn cứ tra cứu');
  const compact = picked.length ? `${prefix}. Nội dung liên quan: ${picked.join(' ')}` : `${prefix}. ${text.slice(0,900)}`;
  els.internetProposedSource.value = compact.replace(/\s+/g,' ').trim();
  setInternetStatus('✅ Đã rút gọn nội dung dán thành căn cứ đề xuất. Hãy kiểm tra lại trước khi cập nhật vào bộ đề.', 'good');
}
async function fetchInternetApi(){
  const q = currentInternetQuestion();
  const raw = visibleText(els.internetApiUrl?.value || '');
  if(!q){ setInternetStatus('Chưa chọn câu hỏi.', 'bad'); return; }
  if(!raw){ setInternetStatus('Chưa nhập URL API. Có thể dùng API/proxy nội bộ trả về text hoặc JSON có trường text/snippet/content.', 'bad'); return; }
  const url = raw.replaceAll('{q}', encodeURIComponent(buildInternetQuery(q))).replaceAll('{question}', encodeURIComponent(q.question || '')).replaceAll('{answer}', encodeURIComponent(q.correctText || '')).replaceAll('{source}', encodeURIComponent(q.sourceOriginal || q.source || '')).replaceAll('{legal}', encodeURIComponent(legalSearchPhrase(q)));
  try{
    setInternetStatus('Đang gọi API tra cứu...', 'info');
    const res = await fetch(url, {cache:'no-store'});
    if(!res.ok) throw new Error('HTTP ' + res.status);
    const ct = res.headers.get('content-type') || '';
    let text = '';
    if(ct.includes('application/json')){
      const data = await res.json();
      text = data.text || data.snippet || data.content || data.result || data.answer || JSON.stringify(data);
    } else text = await res.text();
    if(els.internetProposedSource) els.internetProposedSource.value = stripHtmlToText(text).slice(0,5000);
    suggestSourceFromPasted();
  } catch(e){
    setInternetStatus('❌ Không gọi được API/nguồn Internet. Nguyên nhân thường gặp: không có mạng, CORS, hoặc API cần khóa truy cập. Có thể dùng nút “Mở tra cứu” rồi dán thủ công.', 'bad');
  }
}
async function applyInternetSource(append=false){
  const q = currentInternetQuestion();
  const idx = Number(els.internetQuestionSelect?.value || 0);
  const val = visibleText(els.internetProposedSource?.value || '');
  if(!q){ setInternetStatus('Chưa chọn câu hỏi để cập nhật.', 'bad'); return; }
  if(!val){ setInternetStatus('Chưa có nội dung căn cứ đề xuất để cập nhật.', 'bad'); return; }
  const old = visibleText(q.source || '');
  if(!q.sourceOriginal) q.sourceOriginal = compactOriginalSource(old);
  q.source = append && old ? unique([old, val]).join(' | ') : val;
  state.bank[idx] = q;
  state.meta = Object.assign({}, state.meta || {}, {updatedSourceAt:new Date().toISOString(), updatedSourceBy:'legal_source_lookup_module_v17'});
  renderStats(); renderPreview(); updateInternetQuestionView();
  try { await saveBank('ngân hàng đã cập nhật căn cứ', true); } catch(e){}
  setInternetStatus(`✅ Đã ${append?'ghép thêm':'cập nhật'} căn cứ cho câu ${idx+1} và lưu vào database offline.`, 'good');
}
async function saveInternetUpdatedBank(){
  await saveBank('ngân hàng đã cập nhật căn cứ từ Internet');
  setInternetStatus('✅ Đã lưu ngân hàng đã cập nhật căn cứ vào database offline.', 'good');
}
function exportUpdatedBankXlsx(){
  if(!window.XLSX){ setInternetStatus('Không tìm thấy SheetJS để xuất Excel.', 'bad'); return; }
  if(!state.bank.length){ setInternetStatus('Chưa có ngân hàng câu hỏi để xuất.', 'bad'); return; }
  const header = ['STT','Câu hỏi','Đáp án đúng','Phương án lựa chọn 1','Phương án lựa chọn 2','Phương án lựa chọn 3','Phương án lựa chọn 4','Căn cứ / giải thích','ID gốc','Sheet','Dòng nguồn'];
  const rows = state.bank.map((q,i)=>[
    i+1, q.question || '', q.correctText || '', q.options[0] || '', q.options[1] || '', q.options[2] || '', q.options[3] || '', q.source || '', q.id || '', q.sheetName || '', q.sourceRow || ''
  ]);
  const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
  ws['!cols'] = [{wch:6},{wch:48},{wch:34},{wch:34},{wch:34},{wch:34},{wch:34},{wch:70},{wch:10},{wch:14},{wch:10}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Bo de cap nhat can cu');
  XLSX.writeFile(wb, 'bo_de_cap_nhat_can_cu_' + fileStamp() + '.xlsx');
  setInternetStatus('✅ Đã xuất Excel bộ đề có cột căn cứ đã cập nhật.', 'good');
}
function exportInternetLinksHtml(){
  if(!state.bank.length){ setInternetStatus('Chưa có ngân hàng câu hỏi để xuất danh sách tra cứu.', 'bad'); return; }
  const rows = state.bank.map((q,i)=>{ const parsed=parseLegalReference(q.sourceOriginal || q.source || ''); return `<tr><td>${i+1}</td><td>${escapeHtml(q.sourceOriginal || compactOriginalSource(q.source) || '')}</td><td>${escapeHtml(parsed.meta ? parsed.meta.shortTitle : (parsed.code || ''))}</td><td>${escapeHtml(q.question)}</td><td>${escapeHtml(q.correctText||'')}</td><td>${escapeHtml(q.source||'')}</td><td><a href="${internetSearchUrl(q,false)}" target="_blank">Tìm Internet</a></td><td><a href="${internetSearchUrl(q,true)}" target="_blank">Tìm nguồn pháp lý</a></td></tr>`; }).join('');
  const html = makeStandaloneHtml('Danh sách link tra cứu căn cứ', `<div class="card report-card-full"><h2>Danh sách tra cứu căn cứ</h2><p class="muted">Bấm link để tra cứu, sau đó quay lại PWA và cập nhật cột căn cứ.</p><div class="table-wrap"><table><thead><tr><th>STT</th><th>Căn cứ gốc</th><th>Văn bản</th><th>Câu hỏi</th><th>Đáp án đúng</th><th>Căn cứ cập nhật</th><th>Internet</th><th>Nguồn pháp lý</th></tr></thead><tbody>${rows}</tbody></table></div></div>`);
  downloadTextFile('danh_sach_tra_cuu_can_cu_' + fileStamp() + '.html', html);
  setInternetStatus('✅ Đã xuất HTML danh sách link tra cứu căn cứ cho toàn bộ ngân hàng.', 'good');
}

function openDB(){
  return new Promise((resolve,reject) => {
    if(!('indexedDB' in window)) return reject(new Error('Trình duyệt không hỗ trợ IndexedDB'));
    const req = indexedDB.open(DB_NAME,1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE_NAME);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error('Không mở được IndexedDB'));
  });
}
async function dbPut(key, value){ const db=await openDB(); return new Promise((res,rej)=>{ const tx=db.transaction(STORE_NAME,'readwrite'); tx.objectStore(STORE_NAME).put(value,key); tx.oncomplete=()=>res(); tx.onerror=()=>rej(tx.error); }); }
async function dbGet(key){ const db=await openDB(); return new Promise((res,rej)=>{ const tx=db.transaction(STORE_NAME,'readonly'); const req=tx.objectStore(STORE_NAME).get(key); req.onsuccess=()=>res(req.result); req.onerror=()=>rej(req.error); }); }
async function dbDel(key){ const db=await openDB(); return new Promise((res,rej)=>{ const tx=db.transaction(STORE_NAME,'readwrite'); tx.objectStore(STORE_NAME).delete(key); tx.oncomplete=()=>res(); tx.onerror=()=>rej(tx.error); }); }
async function saveBank(label='ngân hàng hiện tại', silent=false){
  const payload = {bank: state.bank, errors: state.errors, meta: state.meta, fileName: state.fileName, sheetName: state.currentSheet, savedAt: new Date().toISOString(), version: APP_VERSION};
  try { await dbPut(BANK_KEY, payload); }
  catch(e){ localStorage.setItem(DB_NAME + ':' + BANK_KEY, JSON.stringify(payload)); }
  if(!silent) setStatus(`✅ Đã lưu ${state.bank.length} câu hỏi của ${label} vào database offline.`, 'good');
}
async function loadBank(){
  let payload = null;
  try { payload = await dbGet(BANK_KEY); } catch(e){ try { payload = JSON.parse(localStorage.getItem(DB_NAME + ':' + BANK_KEY) || 'null'); } catch(_){} }
  if(!payload || !Array.isArray(payload.bank)){ setStatus('Chưa có ngân hàng đã lưu. Dữ liệu nhúng trong HTML vẫn đang sẵn sàng để dùng.', 'info'); return; }
  state.bank = normalizeBank(payload.bank); state.errors = payload.errors || []; state.meta = payload.meta || {}; state.fileName = payload.fileName || 'Bộ nhớ đã lưu'; state.currentSheet = payload.sheetName || state.fileName; state.currentRows = []; state.currentMapping = {source:'db'};
  renderAll();
  setStatus(`✅ Đã nạp ${state.bank.length} câu hỏi từ database offline.`, 'good');
}
async function clearBank(){
  if(!confirm('Xóa ngân hàng đã lưu trong database offline? Dữ liệu nhúng trong HTML không bị xóa.')) return;
  try { await dbDel(BANK_KEY); } catch(e){}
  localStorage.removeItem(DB_NAME + ':' + BANK_KEY);
  setStatus('✅ Đã xóa database đã lưu. Dữ liệu nhúng trong HTML vẫn còn, bấm “Dùng dữ liệu nhúng” để nạp lại.', 'good');
}

async function fileToArrayBuffer(file){
  if(file.arrayBuffer) return await file.arrayBuffer();
  return new Promise((resolve,reject)=>{ const fr=new FileReader(); fr.onload=()=>resolve(fr.result); fr.onerror=()=>reject(fr.error || new Error('FileReader lỗi')); fr.readAsArrayBuffer(file); });
}
async function readWorkbook(file){
  if(!window.XLSX) throw new Error('Không tìm thấy thư viện SheetJS XLSX. Hãy đặt xlsx.full.min.js cùng thư mục index.html hoặc dùng bản HTML đơn đã nhúng thư viện.');
  const buf = await fileToArrayBuffer(file);
  const bytes = new Uint8Array(buf);
  try { return XLSX.read(bytes, {type:'array', cellDates:false, raw:false}); }
  catch(e1){
    let binary = ''; const chunk = 0x8000;
    for(let i=0;i<bytes.length;i+=chunk) binary += String.fromCharCode.apply(null, bytes.subarray(i,i+chunk));
    return XLSX.read(binary, {type:'binary', cellDates:false, raw:false});
  }
}
async function handleExcelFile(){
  const file = els.excelFile && els.excelFile.files && els.excelFile.files[0];
  if(!file){ setStatus('Chưa chọn file Excel.', 'bad'); return; }
  setStatus(`Đang đọc file: ${file.name} ...`, 'info');
  try {
    state.workbook = await readWorkbook(file);
    state.fileName = file.name;
    if(!state.workbook.SheetNames || !state.workbook.SheetNames.length) throw new Error('Workbook không có sheet nào.');
    els.sheetSelect.innerHTML = state.workbook.SheetNames.map(s=>`<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');
    parseSelectedSheet(false);
  } catch(err){
    console.error(err);
    setStatus(`❌ Không đọc được file Excel: ${err.message || err}\nDữ liệu nhúng vẫn được giữ nguyên, bạn vẫn có thể tạo đề ngay.`, 'bad');
  }
}
function rowsFromSheet(sheet){ return XLSX.utils.sheet_to_json(sheet, {header:1, raw:false, defval:''}); }
function parseSelectedSheet(manual=false){
  if(!state.workbook){ handleExcelFile(); return; }
  const sheetName = els.sheetSelect.value || state.workbook.SheetNames[0];
  const rows = rowsFromSheet(state.workbook.Sheets[sheetName]);
  state.currentRows = rows;
  state.currentSheet = sheetName;
  let mapping = null;
  if(manual) mapping = manualMappingFromUI(rows);
  try {
    const parsed = parseRows(rows, sheetName, mapping);
    renderMapping(rows, parsed.mapping);
    if(!parsed.bank.length){
      state.errors = parsed.errors;
      setStatus(`❌ Đọc được sheet nhưng chưa nhận diện được câu hỏi hợp lệ. Dữ liệu nhúng cũ vẫn được giữ nguyên.\nHãy mở “Nhận diện thủ công” để chọn đúng cột.`, 'bad');
      return;
    }
    state.bank = parsed.bank; state.errors = parsed.errors; state.currentMapping = parsed.mapping; state.meta = {sourceFile: state.fileName, sheetName, total: parsed.bank.length, version: APP_VERSION};
    renderStats(); renderPreview(); resetQuiz();
    setStatus(`✅ Đã nạp ${parsed.bank.length} câu hỏi từ file “${state.fileName}”, sheet “${sheetName}”.\nHàng tiêu đề: ${parsed.mapping.headerIdx+1}; cột câu hỏi: ${colName(parsed.mapping.questionCol)}; cột đáp án: ${parsed.mapping.answerCol>=0?colName(parsed.mapping.answerCol):'không có'}; cột phương án: ${parsed.mapping.optionCols.map(colName).join(', ')}. ${parsed.errors.length?'Có '+parsed.errors.length+' dòng bỏ qua.':''}`, 'good');
  } catch(err){
    console.error(err); renderMapping(rows, null);
    setStatus(`❌ Lỗi nhận diện sheet: ${err.message || err}\nDữ liệu nhúng cũ vẫn được giữ nguyên.`, 'bad');
  }
}
function parseAllSheets(){
  if(!state.workbook){ handleExcelFile(); return; }
  let all=[], errs=[], logs=[];
  for(const s of state.workbook.SheetNames){
    try { const p = parseRows(rowsFromSheet(state.workbook.Sheets[s]), s, null); all = all.concat(p.bank); errs = errs.concat(p.errors.map(e => ({...e, sheetName:s}))); logs.push(`${s}: ${p.bank.length} câu`); }
    catch(e){ logs.push(`${s}: lỗi (${e.message || e})`); }
  }
  if(!all.length){ setStatus('❌ Không nhận diện được câu hỏi hợp lệ ở bất kỳ sheet nào. Dữ liệu nhúng vẫn được giữ nguyên.\n' + logs.join('\n'), 'bad'); return; }
  state.bank = all; state.errors = errs; state.currentSheet = 'Tất cả sheet'; state.currentMapping = {source:'all_sheets'}; state.meta = {sourceFile: state.fileName, sheetName:'Tất cả sheet', total: all.length, version: APP_VERSION};
  renderStats(); renderPreview(); resetQuiz();
  setStatus(`✅ Đã đọc tất cả sheet. Tổng ${all.length} câu hỏi hợp lệ.\n${logs.join('\n')}${errs.length?'\n⚠️ Bỏ qua '+errs.length+' dòng.':''}`, 'good');
}

function scoreHeader(row){
  const text = row.map(visibleText).join(' | '); const n = norm(text);
  let score = 0;
  if(/cau hoi|noi dung cau hoi|question/.test(n)) score += 30;
  if(/dap an|answer|correct|dung/.test(n)) score += 15;
  const optHits = (n.match(/phuong an|lua chon|option|choice/g) || []).length;
  score += Math.min(40, optHits * 12);
  if(/can cu|phap ly|giai thich|nguon/.test(n)) score += 6;
  const nonEmpty = row.filter(x=>visibleText(x)).length;
  if(nonEmpty >= 4) score += 8;
  return score;
}
function detectHeaderRow(rows){
  let best=0, bestScore=-1; const limit=Math.min(25, rows.length);
  for(let i=0;i<limit;i++){ const s=scoreHeader(rows[i]||[]); if(s>bestScore){ bestScore=s; best=i; } }
  return bestScore > 10 ? best : 0;
}
function findBestCol(headers, patterns, exclude=[]){
  let best=-1, score=-1;
  headers.forEach((h,i)=>{ if(exclude.includes(i)) return; const n=norm(h); let s=0; patterns.forEach(p=>{ if(p.test(n)) s += 10; }); if(s>score){ score=s; best=i; } });
  return score > 0 ? best : -1;
}
function detectMapping(rows){
  const headerIdx = detectHeaderRow(rows), headers = rows[headerIdx] || [], sampleRows = rows.slice(headerIdx+1, Math.min(rows.length, headerIdx+15));
  const questionCol = findBestCol(headers, [/noi dung cau hoi/, /^cau hoi$/, /cau hoi/, /question/, /noi dung/]);
  let answerCol = findBestCol(headers, [/^dap an$/, /dap an(?!.*phuong)/, /answer/, /correct/, /dap an dung/], questionCol>=0?[questionCol]:[]);
  const sourceCol = findBestCol(headers, [/can cu/, /phap ly/, /giai thich/, /nguon/, /reference/, /explain/], [questionCol, answerCol]);
  const idCol = findBestCol(headers, [/^stt$/, /^tt$/, /^id$/, /so thu tu/], [questionCol, answerCol, sourceCol]);
  let optionCols = [];
  headers.forEach((h,i)=>{
    if([questionCol,answerCol,sourceCol,idCol].includes(i)) return;
    const n=norm(h);
    if(/phuong an|lua chon|option|choice/.test(n) || /^[abcd]$/.test(n) || /^pa\s*[1-6a-f]$/.test(n)) optionCols.push(i);
  });
  const maxCols = Math.max(...rows.slice(0, Math.min(rows.length, 20)).map(r=>r.length), 0);
  if(optionCols.length < 2){
    const candidates=[];
    for(let c=0;c<maxCols;c++){
      if([questionCol,answerCol,sourceCol,idCol].includes(c)) continue;
      let filled=0, avgLen=0;
      sampleRows.forEach(r=>{ const t=visibleText(r[c]); if(t){ filled++; avgLen += t.length; }});
      if(filled >= Math.max(2, Math.floor(sampleRows.length*0.35))) candidates.push({c, filled, avgLen: avgLen/(filled||1)});
    }
    candidates.sort((a,b)=> b.filled-a.filled || a.c-b.c);
    optionCols = candidates.map(x=>x.c).slice(0,6);
  }
  if(questionCol < 0){
    let best=-1, len=-1;
    for(let c=0;c<maxCols;c++){
      if(optionCols.includes(c) || [answerCol,sourceCol,idCol].includes(c)) continue;
      const total = sampleRows.reduce((s,r)=>s+visibleText(r[c]).length,0);
      if(total>len){ len=total; best=c; }
    }
    return {headerIdx, headers, idCol, questionCol:best, answerCol, sourceCol, optionCols};
  }
  return {headerIdx, headers, idCol, questionCol, answerCol, sourceCol, optionCols};
}
function parseRows(rows, sheetName, manualMapping){
  const mapping = manualMapping || detectMapping(rows);
  if(mapping.questionCol < 0) throw new Error('Không tìm thấy cột câu hỏi.');
  if(mapping.optionCols.length < 2) throw new Error('Không tìm thấy đủ cột phương án lựa chọn.');
  const bank=[], errors=[];
  for(let r=mapping.headerIdx+1; r<rows.length; r++){
    const row=rows[r] || [];
    const question = cleanQuestion(row[mapping.questionCol]);
    const rawOptions = mapping.optionCols.map(c => ({col:c, raw: visibleText(row[c])}));
    let markedCorrect = -1;
    let options = rawOptions.map((o,idx) => {
      const marked = /^\s*(?:\[x\]|\(x\)|✓|✔|\*)\s*/i.test(o.raw) || /\s*(?:\[dung\]|\[đúng\])\s*$/i.test(o.raw);
      if(marked && markedCorrect < 0) markedCorrect = idx;
      return cleanOption(o.raw);
    }).filter(Boolean);
    const id = mapping.idCol >= 0 ? visibleText(row[mapping.idCol]) : String(bank.length + 1);
    const source = mapping.sourceCol >= 0 ? visibleText(row[mapping.sourceCol]) : '';
    const answerRaw = mapping.answerCol >= 0 ? visibleText(row[mapping.answerCol]) : '';
    let correctIndex = inferAnswerIndex(answerRaw, options);
    if(correctIndex < 0 && markedCorrect >= 0) correctIndex = markedCorrect;
    if(!question){ if(row.some(x=>visibleText(x))) errors.push({row:r+1, reason:'Thiếu câu hỏi'}); continue; }
    if(options.length < 2){ errors.push({row:r+1, reason:'Thiếu phương án lựa chọn', question:question.slice(0,80)}); continue; }
    if(correctIndex < 0 || correctIndex >= options.length){ errors.push({row:r+1, reason:'Không xác định được đáp án đúng', answerRaw, question:question.slice(0,80)}); continue; }
    bank.push({id:id || String(bank.length+1), sourceRow:r+1, sheetName, question, options, correctIndex, correctText:options[correctIndex], source, answerRaw});
  }
  return {bank, errors, mapping};
}
function cleanQuestion(v){ return visibleText(v).replace(/^\s*(?:câu\s*)?\d+[.)\-:]\s*/i,'').trim(); }
function cleanOption(v){ return visibleText(v).replace(/^\s*(?:\[x\]|\(x\)|✓|✔|\*)\s*/i,'').replace(/\s*(?:\[dung\]|\[đúng\])\s*$/i,'').replace(/^\s*(?:[A-Fa-f]|[1-6])\s*[.)\-:]\s*/,'').trim(); }
function inferAnswerIndex(answerRaw, options){
  const a = visibleText(answerRaw); if(!a) return -1;
  const an = norm(a);
  if(/^\d+$/.test(an)){ const i=parseInt(an,10)-1; return i>=0&&i<options.length?i:-1; }
  if(/^[a-f]$/.test(an)){ const i=an.charCodeAt(0)-97; return i>=0&&i<options.length?i:-1; }
  let m = an.match(/(?:phuong an|lua chon|option|choice|pa)\s*([1-6a-f])/);
  if(m){ const x=m[1]; const i=/\d/.test(x)?parseInt(x,10)-1:x.charCodeAt(0)-97; return i>=0&&i<options.length?i:-1; }
  m = an.match(/(?:dap an|answer)\s*[:\-]?\s*([1-6a-f])/);
  if(m){ const x=m[1]; const i=/\d/.test(x)?parseInt(x,10)-1:x.charCodeAt(0)-97; return i>=0&&i<options.length?i:-1; }
  const exact = options.findIndex(o => norm(o) === an); if(exact >= 0) return exact;
  if(an.length > 8){ const contains = options.findIndex(o => norm(o).includes(an) || an.includes(norm(o))); if(contains >= 0) return contains; }
  return -1;
}
function colName(idx){ if(idx<0) return ''; let s='', n=idx+1; while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s; }
function renderMapping(rows, mapping){
  if(!rows || !rows.length || !els.manualQuestionCol) return;
  const m = mapping || detectMapping(rows);
  const headerIdx = Math.max(0, Number(els.manualHeaderRow.value || m.headerIdx+1)-1);
  const headers = rows[headerIdx] || rows[m.headerIdx] || [];
  const sample = rows[headerIdx+1] || [];
  const maxCol = Math.max(headers.length, sample.length, 8);
  els.manualHeaderRow.value = String(headerIdx + 1);
  const none = '<option value="-1">-- Không dùng --</option>';
  let opts = '';
  for(let c=0;c<maxCol;c++){
    const head = visibleText(headers[c]) || `(cột ${colName(c)})`;
    const ex = visibleText(sample[c]);
    opts += `<option value="${c}">${colName(c)} - ${escapeHtml(head)}${ex?' | VD: '+escapeHtml(ex.slice(0,65)):''}</option>`;
  }
  els.manualQuestionCol.innerHTML = none + opts;
  els.manualAnswerCol.innerHTML = none + opts;
  els.manualSourceCol.innerHTML = none + opts;
  els.manualOptionCols.innerHTML = opts;
  const use = mapping || m;
  els.manualQuestionCol.value = String(use.questionCol);
  els.manualAnswerCol.value = String(use.answerCol);
  els.manualSourceCol.value = String(use.sourceCol);
  [...els.manualOptionCols.options].forEach(o => o.selected = use.optionCols.includes(Number(o.value)));
}
function manualMappingFromUI(rows){
  const headerIdx = Math.max(0, Number(els.manualHeaderRow.value || 1)-1);
  return {headerIdx, headers: rows[headerIdx] || [], idCol:-1, questionCol:Number(els.manualQuestionCol.value), answerCol:Number(els.manualAnswerCol.value), sourceCol:Number(els.manualSourceCol.value), optionCols:[...els.manualOptionCols.selectedOptions].map(o=>Number(o.value)).filter(v=>v>=0), manual:true};
}

function hashSeed(text){ let h=2166136261; const s=String(text || Date.now()); for(let i=0;i<s.length;i++){ h ^= s.charCodeAt(i); h = Math.imul(h,16777619); } return h>>>0; }
function rng(seed){ let a=seed>>>0; return function(){ a+=0x6D2B79F5; let t=a; t=Math.imul(t^t>>>15,t|1); t^=t+Math.imul(t^t>>>7,t|61); return ((t^t>>>14)>>>0)/4294967296; }; }
function shuffled(arr, rand){ const a=arr.slice(); for(let i=a.length-1;i>0;i--){ const j=Math.floor(rand()*(i+1)); [a[i],a[j]]=[a[j],a[i]]; } return a; }
function startQuiz(){
  state.mode = 'exam';
  updateStickyLabels();
  if(!state.bank.length){ setStatus('Chưa có ngân hàng câu hỏi.', 'bad'); return; }
  const count = Math.max(1, Math.min(Number(els.quizCount.value || 1), state.bank.length));
  const seedText = els.seedInput.value.trim() || new Date().toISOString();
  state.lastQuizConfig = {
    count,
    shuffleQuestions: !!els.shuffleQuestions.checked,
    shuffleOptions: !!els.shuffleOptions.checked,
    showAutoExplain: !!els.showAutoExplain.checked
  };
  const rand = rng(hashSeed(seedText));
  let questions = els.shuffleQuestions.checked ? shuffled(state.bank, rand) : state.bank.slice();
  questions = questions.slice(0, count);
  state.quiz = questions.map((q, qi) => {
    const opts = q.options.map((text, idx) => ({text, originalIndex:idx, isCorrect:idx===q.correctIndex}));
    const finalOpts = els.shuffleOptions.checked ? shuffled(opts, rand) : opts;
    return {no: qi+1, item: q, options: finalOpts, userChoice: null};
  });
  state.submitted = false; state.lastResult = null;
  renderQuiz();
  els.btnSubmitQuiz.disabled = false; els.btnSubmitSticky.disabled = false; if(els.btnSubmitQuizTop) els.btnSubmitQuizTop.disabled = false;
  els.resultSummary.textContent = 'Chưa nộp bài.'; els.resultList.innerHTML = '';
  setStatus(`✅ Đã tạo đề ${count} câu. Mã đảo đề/seed: ${seedText}`, 'good');
  enterQuizFocus();
  location.hash = '#quizSection';
}

function startStudyRound(){
  if(!state.bank.length){ setStatus('Chưa có ngân hàng câu hỏi để ôn tập.', 'bad'); return; }
  state.mode = 'study';
  syncStudyConfigFromUI();
  updateStickyLabels();
  const activeIds = refreshStudyActiveIds();
  if(!activeIds.length){
    setStatus('✅ Bạn đã học đủ toàn bộ câu hỏi trong ngân hàng. Có thể bấm “Xuất HTML thống kê ôn tập” hoặc “Xuất câu sai nhiều” để lưu lại quá trình học.', 'good');
    renderStudyStats();
    return;
  }
  const idSet = new Set(activeIds);
  let questions = state.bank.filter(q => idSet.has(getQuestionKey(q)));
  const seedText = 'on-tap-' + new Date().toISOString();
  const rand = rng(hashSeed(seedText));
  if(state.study.config.shuffleQuestions) questions = shuffled(questions, rand);
  state.quiz = questions.map((q, qi) => {
    const opts = q.options.map((text, idx) => ({text, originalIndex:idx, isCorrect:idx===q.correctIndex}));
    const finalOpts = state.study.config.shuffleOptions ? shuffled(opts, rand) : opts;
    return {no: qi+1, item:q, options:finalOpts, userChoice:null};
  });
  state.submitted = false; state.lastResult = null;
  state.lastQuizConfig = {mode:'study', count:questions.length, shuffleQuestions:state.study.config.shuffleQuestions, shuffleOptions:state.study.config.shuffleOptions, showAutoExplain:!!els.showAutoExplain?.checked};
  renderQuiz();
  els.btnSubmitQuiz.disabled = false; els.btnSubmitSticky.disabled = false; if(els.btnSubmitQuizTop) els.btnSubmitQuizTop.disabled = false;
  els.resultSummary.textContent = 'Chưa nộp bài.'; els.resultList.innerHTML = '';
  setStatus(`📚 Đã tạo lượt ôn tập ${questions.length} câu. Câu nào đúng đủ ${studyThreshold()} lần hoặc được xác nhận đã học sẽ được thay bằng câu mới ở lượt tiếp.`, 'good');
  enterQuizFocus();
  location.hash = '#quizSection';
}
function resetStudyProgress(){
  if(!confirm('Xóa toàn bộ tiến độ ôn tập đã lưu?')) return;
  state.study.progress = {}; state.study.activeIds = [];
  saveStudyState(); renderStudyStats();
  setStatus('✅ Đã xóa tiến độ ôn tập. Có thể bắt đầu học lại từ đầu.', 'good');
}
function markCurrentStudyRoundLearned(){
  if(!state.quiz.length){ setStatus('Chưa có lượt ôn tập để xác nhận đã học.', 'bad'); return; }
  state.quiz.forEach(q => markStudyLearnedByKey(getQuestionKey(q.item), false));
  refreshStudyActiveIds(); renderResult();
  setStatus('✅ Đã đánh dấu các câu trong lượt hiện tại là đã học. Bấm “Lượt tiếp” để nạp câu mới.', 'good');
}
function markStudyLearnedByKey(key, rerender=true){
  const p = state.study.progress[key] || {};
  p.learned = true; p.learnedAt = new Date().toISOString(); p.manualLearned = true;
  state.study.progress[key] = p;
  state.study.activeIds = (state.study.activeIds || []).filter(x => x !== key);
  saveStudyState(); renderStudyStats();
  if(rerender && state.lastResult) renderResult();
}
function applyStudySubmission(){
  const threshold = studyThreshold();
  state.quiz.forEach(q => {
    const key = getQuestionKey(q.item);
    const p = state.study.progress[key] || {correctCount:0, wrongCount:0, attempts:0, streak:0, learned:false};
    const ok = q.userChoice !== null && q.options[q.userChoice]?.isCorrect;
    p.attempts = Number(p.attempts || 0) + 1;
    p.lastAt = new Date().toISOString();
    p.lastCorrect = !!ok;
    if(ok){ p.correctCount = Number(p.correctCount || 0) + 1; p.streak = Number(p.streak || 0) + 1; }
    else { p.wrongCount = Number(p.wrongCount || 0) + 1; p.streak = 0; }
    if(Number(p.correctCount || 0) >= threshold){ p.learned = true; p.learnedAt = p.learnedAt || new Date().toISOString(); }
    state.study.progress[key] = p;
  });
  state.study.activeIds = (state.study.activeIds || []).filter(k => !isStudyLearned(k));
  saveStudyState();
  renderStudyStats();
}
function studyQuestionMeta(item){
  const key = getQuestionKey(item);
  const p = state.study.progress[key] || {};
  const learned = isStudyLearned(key);
  const correct = Number(p.correctCount || 0), attempts = Number(p.attempts || 0), wrong = Number(p.wrongCount || 0);
  const btn = learned ? '<span class="pill okp">Đã học</span>' : `<button type="button" class="light btn-mark-learned" data-study-key="${escapeHtml(key)}">Đã học câu này</button>`;
  return `<div class="study-result-meta"><span class="pill">Ôn tập: đúng ${correct}/${studyThreshold()} lần</span><span class="pill">Đã làm ${attempts} lượt</span><span class="pill ${wrong?'warnp':'okp'}">Sai ${wrong}</span>${btn}</div>`;
}
function bindResultStudyButtons(){
  if(!els.resultList) return;
  els.resultList.querySelectorAll('.btn-mark-learned').forEach(btn => btn.addEventListener('click', () => {
    markStudyLearnedByKey(btn.dataset.studyKey);
    setStatus('✅ Đã xác nhận câu này đã học. Bấm “Lượt tiếp” để thay bằng câu mới.', 'good');
  }));
}
function resetQuiz(){
  state.quiz=[]; state.submitted=false; state.lastResult=null;
  if(els.quizInfo) els.quizInfo.textContent='';
  if(els.quizList) els.quizList.innerHTML='';
  if(els.resultSummary) els.resultSummary.textContent='Chưa nộp bài.';
  if(els.resultList) els.resultList.innerHTML='';
  if(els.btnSubmitQuiz) els.btnSubmitQuiz.disabled=true;
  if(els.btnSubmitSticky) els.btnSubmitSticky.disabled=true;
  if(els.btnSubmitQuizTop) els.btnSubmitQuizTop.disabled=true;
  updateProgress();
}
function renderQuiz(){
  if(!state.quiz.length){ if(els.quizInfo) els.quizInfo.textContent=''; els.quizList.innerHTML=''; updateProgress(); return; }
  if(els.quizInfo) els.quizInfo.textContent = '';
  els.quizList.innerHTML = state.quiz.map((q, qi) => {
    const opts = q.options.map((op, oi) => {
      const letter = 'ABCDEF'[oi] || String(oi+1);
      let cls = 'option';
      if(state.submitted){ if(op.isCorrect) cls += ' right'; if(q.userChoice === oi && !op.isCorrect) cls += ' wrong'; }
      else if(q.userChoice === oi) cls += ' chosen';
      return `<label class="${cls}"><input type="radio" name="q${qi}" value="${oi}" ${q.userChoice===oi?'checked':''} ${state.submitted?'disabled':''}><span class="letter">${letter}</span><span>${escapeHtml(op.text)}</span></label>`;
    }).join('');
    return `<article class="question-card" data-q="${qi}"><div class="question-title">Câu ${q.no}. ${escapeHtml(q.item.question)}</div>${opts}<div class="muted small">Nguồn: ${escapeHtml(q.item.sheetName||'')} dòng ${escapeHtml(q.item.sourceRow||'')}${q.item.source?' | '+escapeHtml(q.item.source):''}</div></article>`;
  }).join('');
  els.quizList.querySelectorAll('input[type=radio]').forEach(input => input.addEventListener('change', ev => {
    const card = ev.target.closest('.question-card'); state.quiz[Number(card.dataset.q)].userChoice = Number(ev.target.value); renderQuiz(); updateProgress();
  }));
  updateProgress();
}
function updateProgress(){
  const total=state.quiz.length, done=state.quiz.filter(q=>q.userChoice !== null).length;
  if(els.progressText) els.progressText.textContent = `${done}/${total}`;
  if(els.progressBar) els.progressBar.style.width = total ? `${done/total*100}%` : '0%';
}
function submitQuiz(){
  if(!state.quiz.length) return;
  const unanswered = state.quiz.filter(q=>q.userChoice === null).length;
  if(unanswered && !confirm(`Còn ${unanswered} câu chưa chọn. Vẫn nộp bài?`)) return;
  let correct=0; state.quiz.forEach(q => { if(q.userChoice !== null && q.options[q.userChoice]?.isCorrect) correct++; });
  state.submitted=true; state.lastResult={correct,total:state.quiz.length,score10:state.quiz.length?correct/state.quiz.length*10:0,time:new Date().toLocaleString('vi-VN')};
  if(state.mode === 'study') applyStudySubmission();
  renderQuiz(); renderResult();
  els.btnSubmitQuiz.disabled = true; els.btnSubmitSticky.disabled = true; if(els.btnSubmitQuizTop) els.btnSubmitQuizTop.disabled = true;
  enterResultFocus();
  location.hash = '#resultSection';
}

function resultModeFor(q){
  const chosen = q.userChoice === null ? null : q.options[q.userChoice];
  if(!chosen) return 'unanswered';
  return chosen.isCorrect ? 'correct' : 'wrong';
}
function resultRank(mode){
  if(mode === 'wrong') return 0;
  if(mode === 'unanswered') return 1;
  return 2;
}
function sortedResultEntries(){
  return state.quiz.map((q, idx) => ({q, idx, mode: resultModeFor(q)}))
    .sort((a,b) => resultRank(a.mode) - resultRank(b.mode) || a.idx - b.idx);
}
function renderResult(){
  const r=state.lastResult; if(!r) return;
  const wrongCount = state.quiz.filter(q => resultModeFor(q) === 'wrong').length;
  const unansweredCount = state.quiz.filter(q => resultModeFor(q) === 'unanswered').length;
  els.resultSummary.innerHTML = `<span class="pill okp">Đúng ${r.correct}/${r.total}</span><span class="pill ${wrongCount?'warnp':'okp'}">Sai ${wrongCount}</span><span class="pill ${unansweredCount?'warnp':'okp'}">Chưa chọn ${unansweredCount}</span><span class="pill">Điểm ${r.score10.toFixed(2)}/10</span><span class="pill">${escapeHtml(r.time)}</span><span class="pill">Đã xếp câu sai lên trước</span>`;

  function correctReason(q, correctOpt){
    return `<div class="analysis-row source"><b>Vì sao đây là đáp án đúng:</b><br>Phương án này là nội dung được cột đáp án trong Excel xác định là đúng. Khi so với các phương án sai, đáp án đúng giữ đúng từ khóa, điều kiện, mốc số liệu hoặc phạm vi bắt buộc.${q.item.source?'<br><b>Căn cứ:</b> '+escapeHtml(q.item.source):''}</div>`;
  }
  function wrongExplainBlock(correctOpt, op, oi, title){
    const letter = 'ABCDEF'[oi] || String(oi+1);
    return `<div class="analysis-row"><b>${escapeHtml(title || ('Phương án ' + letter + ' sai ở đâu:'))}</b><br>${els.showAutoExplain.checked ? explainDifference(correctOpt.text, op.text) : 'Đã tắt phân tích tự động.'}</div>`;
  }
  function optionExplanation(q, correctOpt, op, oi, mode){
    if(op.isCorrect) return correctReason(q, correctOpt);
    const letter = 'ABCDEF'[oi] || String(oi+1);
    if(mode === 'wrong' && oi === q.userChoice){
      return wrongExplainBlock(correctOpt, op, oi, 'Bạn chọn phương án ' + letter + ' nên sai ở đâu so với đáp án đúng:');
    }
    return wrongExplainBlock(correctOpt, op, oi, 'Phương án ' + letter + ' sai ở đâu so với đáp án đúng:');
  }
  function optionStatus(q, op, oi){
    if(op.isCorrect && q.userChoice === oi) return '<span class="mini-status ok">Bạn chọn đúng</span>';
    if(op.isCorrect) return '<span class="mini-status ok">Đáp án đúng</span>';
    if(q.userChoice === oi) return '<span class="mini-status bad">Bạn chọn sai</span>';
    return '<span class="mini-status muted">Phương án sai</span>';
  }
  function resultOption(q, correctOpt, op, oi, mode){
    const letter = 'ABCDEF'[oi] || String(oi+1);
    const isUserWrongChoice = mode === 'wrong' && oi === q.userChoice && !op.isCorrect;
    const open = isUserWrongChoice ? ' open' : '';
    const cls = ['result-option','option'];
    if(op.isCorrect) cls.push('right','result-correct-option');
    if(q.userChoice === oi && !op.isCorrect) cls.push('wrong','result-user-wrong');
    if(q.userChoice === oi && op.isCorrect) cls.push('chosen','result-user-correct');
    const textCls = op.isCorrect ? 'option-text correct-answer-text' : 'option-text';
    const explanation = optionExplanation(q, correctOpt, op, oi, mode);
    const arrowTitle = op.isCorrect ? 'Xem vì sao đây là đáp án đúng' : 'Xem vì sao phương án này sai';
    return `<div class="${cls.join(' ')}${open}">
      <div class="result-option-main">
        <span class="letter">${letter}</span>
        <span class="${textCls}">${escapeHtml(op.text)} ${optionStatus(q, op, oi)}</span>
        <button type="button" class="option-arrow" title="${arrowTitle}" aria-label="${arrowTitle}" aria-expanded="${open?'true':'false'}" onclick="const p=this.closest('.result-option');p.classList.toggle('open');this.setAttribute('aria-expanded',p.classList.contains('open')?'true':'false');this.textContent=p.classList.contains('open')?'⌃':'⌄';">${open?'⌃':'⌄'}</button>
      </div>
      <div class="option-explain-panel">${explanation}</div>
    </div>`;
  }

  const ordered = sortedResultEntries();
  els.resultList.innerHTML = ordered.map((entry) => {
    const q = entry.q;
    const originalNo = entry.idx + 1;
    const chosen = q.userChoice === null ? null : q.options[q.userChoice];
    const correctOpt = q.options.find(o => o.isCorrect) || q.options[0];
    const ok = !!(chosen && chosen.isCorrect);
    const mode = entry.mode;
    const resultText = mode === 'unanswered' ? 'Chưa chọn' : (ok ? 'Đúng' : 'Sai');
    const resultCls = mode === 'unanswered' ? 'warn' : (ok ? 'ok' : 'bad');
    const resultPill = mode === 'unanswered' ? 'warnp' : (ok ? 'okp' : 'warnp');
    const studyMeta = state.mode === 'study' ? studyQuestionMeta(q.item) : '';
    const shortNote = mode === 'wrong'
      ? '<div class="analysis-row focus-diff compact-result-note"><b>Bạn trả lời sai.</b> Phương án bạn chọn đã được mở sẵn phần giải thích. Các câu sai được xếp lên trước để kiểm tra nhanh.</div>'
      : (mode === 'correct'
        ? '<div class="analysis-row source compact-result-note"><b>Bạn trả lời đúng.</b> Các giải thích được ẩn dưới nút mũi tên ở cuối từng phương án.</div>'
        : '<div class="analysis-row compact-unanswered"><b>Bạn chưa chọn câu này.</b> Đáp án đúng đã được tô chữ xanh. Bấm mũi tên cuối từng phương án để xem giải thích.</div>');
    const optionsHtml = q.options.map((op, oi) => resultOption(q, correctOpt, op, oi, mode)).join('');
    return `<article class="question-card result-question-card result-order-${mode}"><h3>Câu ${q.no}. <span class="${resultCls}">${resultText}</span> <span class="muted small">(thứ tự gốc trong đề: ${originalNo})</span></h3><div class="question-title result-question-title">${escapeHtml(q.item.question)}</div><div class="result-question-meta"><span class="pill ${resultPill}">Kết quả: ${resultText}</span><span class="pill">Đáp án đúng: ${escapeHtml(correctOpt.text)}</span>${chosen?`<span class="pill">Bạn chọn: ${escapeHtml(chosen.text)}</span>`:'<span class="pill warnp">Bạn chưa chọn</span>'}</div>${studyMeta}${shortNote}<div class="result-options-list">${optionsHtml}</div></article>`;
  }).join('');
  bindResultStudyButtons();
}

function makeStandaloneHtml(title, contentHtml){
  const css = Array.from(document.querySelectorAll('style')).map(s => s.textContent || '').join('\n');
  const stamp = new Date().toLocaleString('vi-VN');
  return `<!doctype html><html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escapeHtml(title)}</title><style>${css}\nbody{overflow:auto!important}.app{max-width:none!important;width:100%!important;margin:0!important;padding:12px!important}.card{max-width:none!important;width:100%!important;margin-left:0!important;margin-right:0!important}.no-print,.stickybar{display:none!important}.table-wrap table{min-width:1280px}</style></head><body class="standalone-export"><main class="app"><div class="card report-card-full"><h1>${escapeHtml(title)}</h1><p class="muted">Xuất lúc: ${escapeHtml(stamp)}. File này xem offline được, không cần PWA.</p></div>${contentHtml}</main></body></html>`;
}
function downloadTextFile(filename, content, mime='text/html;charset=utf-8'){
  const blob = new Blob([content], {type:mime});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url), 1000);
}
function fileStamp(){
  const d = new Date();
  const pad = n => String(n).padStart(2,'0');
  return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}
function exportCurrentResultHtml(){
  if(!state.lastResult || !els.resultList || !els.resultList.innerHTML.trim()){
    setStatus('Chưa có kết quả để xuất HTML. Hãy nộp bài trước.', 'bad'); return;
  }
  const title = state.mode === 'study' ? 'Kết quả lượt ôn tập' : 'Kết quả bài thi trắc nghiệm';
  const body = `<div class="card report-card-full"><h2>Tóm tắt</h2><div class="status">${els.resultSummary.innerHTML}</div><p class="muted">Các câu sai/chưa chọn đã được xếp lên trước câu đúng.</p></div><div class="card report-card-full"><h2>Chi tiết</h2>${els.resultList.innerHTML}</div>`;
  downloadTextFile(`${state.mode === 'study' ? 'ket_qua_on_tap' : 'ket_qua_bai_thi'}_${fileStamp()}.html`, makeStandaloneHtml(title, body));
  setStatus('✅ Đã xuất HTML kết quả. File có thể lưu lại và mở offline để xem.', 'good');
}
function buildStudyStatsHtml(onlyWrong=false){
  const rows = getStudyRows();
  const total = rows.length;
  const learned = rows.filter(r=>r.learned).length;
  const attempts = rows.reduce((s,r)=>s+r.attempts,0);
  const correct = rows.reduce((s,r)=>s+r.correctCount,0);
  const wrong = rows.reduce((s,r)=>s+r.wrongCount,0);
  let list = rows;
  if(onlyWrong) list = list.filter(r=>r.wrongCount > 0).sort((a,b)=>b.wrongCount-a.wrongCount || b.attempts-a.attempts || a.idx-b.idx);
  else list = list.sort((a,b)=>Number(a.learned)-Number(b.learned) || b.wrongCount-a.wrongCount || a.idx-b.idx);
  const tableRows = list.map((r,i)=>`<tr><td class="nowrap">${i+1}</td><td class="nowrap">${r.idx+1}</td><td>${escapeHtml(r.q.question)}</td><td>${escapeHtml(r.q.correctText || '')}</td><td class="nowrap">${r.correctCount}</td><td class="nowrap">${r.wrongCount}</td><td class="nowrap">${r.attempts}</td><td>${r.learned?'Đã học':'Đang học'}</td><td>${escapeHtml(r.q.source || '')}</td></tr>`).join('') || '<tr><td colspan="9" class="muted">Chưa có câu sai nào được ghi nhận.</td></tr>';
  return `<div class="card wide-report study-report"><h2>${onlyWrong?'Thống kê câu sai nhiều':'Thống kê quá trình ôn tập'}</h2><div class="row"><span class="pill okp">Đã học ${learned}/${total}</span><span class="pill">Tổng lượt làm ${attempts}</span><span class="pill okp">Tổng lượt đúng ${correct}</span><span class="pill ${wrong?'warnp':'okp'}">Tổng lượt sai ${wrong}</span><span class="pill">Ngưỡng đạt ${studyThreshold()} lần đúng</span></div><p class="muted">${onlyWrong?'Danh sách được sắp xếp theo số lần sai giảm dần.':'Danh sách ưu tiên câu chưa học và câu sai nhiều để tiếp tục ôn.'}</p><div class="table-wrap"><table><thead><tr><th>STT</th><th>Câu trong NH</th><th>Câu hỏi</th><th>Đáp án đúng</th><th>Đúng</th><th>Sai</th><th>Lượt</th><th>Trạng thái</th><th>Căn cứ</th></tr></thead><tbody>${tableRows}</tbody></table></div></div>`;
}
function exportStudyStatsHtml(){
  if(!state.bank.length){ setStatus('Chưa có ngân hàng câu hỏi để xuất thống kê.', 'bad'); return; }
  downloadTextFile(`thong_ke_on_tap_${fileStamp()}.html`, makeStandaloneHtml('Thống kê quá trình ôn tập', buildStudyStatsHtml(false)));
  setStatus('✅ Đã xuất HTML thống kê quá trình ôn tập.', 'good');
}
function exportStudyWrongsHtml(){
  if(!state.bank.length){ setStatus('Chưa có ngân hàng câu hỏi để xuất thống kê.', 'bad'); return; }
  downloadTextFile(`cau_sai_nhieu_${fileStamp()}.html`, makeStandaloneHtml('Thống kê câu sai nhiều', buildStudyStatsHtml(true)));
  setStatus('✅ Đã xuất HTML danh sách câu sai, sắp xếp sai nhiều lên trên cùng.', 'good');
}


function enterQuizFocus(){
  updateStickyLabels();
  document.body.classList.remove('result-fullscreen');
  document.body.classList.add('quiz-fullscreen');
  setTimeout(() => { const q = $('quizSection'); if(q) q.scrollTo({top:0, behavior:'smooth'}); }, 0);
}
function enterResultFocus(){
  updateStickyLabels();
  document.body.classList.remove('quiz-fullscreen');
  document.body.classList.add('result-fullscreen');
  setTimeout(() => { const r = $('resultSection'); if(r) r.scrollTo({top:0, behavior:'smooth'}); }, 0);
}
function exitFocus(){
  document.body.classList.remove('quiz-fullscreen','result-fullscreen','study-mode');
}
function scrollQuizTop(){
  const q = $('quizSection'); if(q) q.scrollTo({top:0, behavior:'smooth'});
}
function scrollCurrentTop(){
  const target = document.body.classList.contains('result-fullscreen') ? $('resultSection') : $('quizSection');
  if(target) target.scrollTo({top:0, behavior:'smooth'});
}
function makeFreshSeed(){
  return 'de-khac-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
}
function startNewQuizSameConfig(){
  if(state.mode === 'study'){ startStudyRound(); return; }
  if(!state.bank.length){ setStatus('Chưa có ngân hàng câu hỏi để tạo đề.', 'bad'); return; }
  const cfg = state.lastQuizConfig;
  if(cfg){
    if(els.quizCount) els.quizCount.value = String(Math.min(cfg.count || state.bank.length, state.bank.length));
    if(els.shuffleQuestions) els.shuffleQuestions.checked = !!cfg.shuffleQuestions;
    if(els.shuffleOptions) els.shuffleOptions.checked = !!cfg.shuffleOptions;
    if(els.showAutoExplain) els.showAutoExplain.checked = !!cfg.showAutoExplain;
  }
  if(els.seedInput) els.seedInput.value = makeFreshSeed();
  startQuiz();
}
function backToSetup(){
  exitFocus();
  const btn = $('btnStartQuiz'); if(btn) btn.scrollIntoView({behavior:'smooth', block:'center'});
}

async function clearOldCache(){
  try {
    if('serviceWorker' in navigator){ const regs = await navigator.serviceWorker.getRegistrations(); await Promise.all(regs.map(r=>r.unregister())); }
    if('caches' in window){ const names = await caches.keys(); await Promise.all(names.map(n=>caches.delete(n))); }
    setStatus('✅ Đã xóa cache/service worker cũ của trình duyệt. Hãy tải lại trang nếu cần.', 'good');
  } catch(e){ setStatus('Không xóa được cache cũ: ' + (e.message || e), 'bad'); }
}

async function forceUpdatePWA(){
  try {
    setStatus('⏳ Đang ép cập nhật PWA: xóa cache cũ, gỡ service worker cũ và tải lại trang...', 'info');
    if('caches' in window){
      const names = await caches.keys();
      await Promise.all(names.map(n => caches.delete(n)));
    }
    if('serviceWorker' in navigator){
      const regs = await navigator.serviceWorker.getRegistrations();
      await Promise.all(regs.map(async r => {
        try { await r.update(); } catch(e) {}
        try { if(r.waiting) r.waiting.postMessage({type:'SKIP_WAITING'}); } catch(e) {}
        try { await r.unregister(); } catch(e) {}
      }));
    }
    const url = new URL(location.href);
    url.searchParams.set('pwa_update', Date.now().toString());
    location.replace(url.toString());
  } catch(e){
    setStatus('❌ Không ép cập nhật được PWA: ' + (e.message || e), 'bad');
  }
}

function registerSW(){
  if('serviceWorker' in navigator && location.protocol !== 'file:'){
    navigator.serviceWorker.register('./sw.js?v=' + encodeURIComponent(APP_VERSION)).catch(err => console.warn('SW register failed:', err));
  }
}
function bindEvents(){
  els.excelFile.addEventListener('change', handleExcelFile);
  els.btnReadSheet.addEventListener('click', () => state.workbook ? parseSelectedSheet(false) : handleExcelFile());
  els.btnReadAll.addEventListener('click', parseAllSheets);
  els.sheetSelect.addEventListener('change', () => parseSelectedSheet(false));
  els.btnRefreshMapping.addEventListener('click', () => state.currentRows.length ? renderMapping(state.currentRows, null) : setStatus('Chưa có sheet để hiển thị cột.', 'bad'));
  els.btnApplyMapping.addEventListener('click', () => parseSelectedSheet(true));
  els.btnUseEmbedded.addEventListener('click', () => { parseEmbedded(); setStatus(`✅ Đã nạp lại ${state.bank.length} câu hỏi nhúng trong HTML.`, 'good'); });
  els.btnSaveEmbedded.addEventListener('click', () => saveBank('dữ liệu nhúng'));
  els.btnSaveBank.addEventListener('click', () => saveBank('ngân hàng hiện tại'));
  els.btnLoadBank.addEventListener('click', loadBank);
  els.btnClearBank.addEventListener('click', clearBank);
  els.btnStartQuiz.addEventListener('click', startQuiz);
  if(els.btnStartStudy) els.btnStartStudy.addEventListener('click', startStudyRound);
  if(els.btnNextStudy) els.btnNextStudy.addEventListener('click', startStudyRound);
  if(els.btnResetStudy) els.btnResetStudy.addEventListener('click', resetStudyProgress);
  if(els.btnMarkStudyAllLearned) els.btnMarkStudyAllLearned.addEventListener('click', markCurrentStudyRoundLearned);
  if(els.btnSaveStudyProgress) els.btnSaveStudyProgress.addEventListener('click', saveStudyProgressManual);
  if(els.btnSaveStudyProgressSticky) els.btnSaveStudyProgressSticky.addEventListener('click', saveStudyProgressManual);
  if(els.btnExportStudyStats) els.btnExportStudyStats.addEventListener('click', exportStudyStatsHtml);
  if(els.btnExportStudyWrongs) els.btnExportStudyWrongs.addEventListener('click', exportStudyWrongsHtml);
  els.btnSubmitQuiz.addEventListener('click', submitQuiz);
  els.btnSubmitSticky.addEventListener('click', submitQuiz);
  els.btnPrint.addEventListener('click', () => window.print());
  if(els.btnExportResultHtml) els.btnExportResultHtml.addEventListener('click', exportCurrentResultHtml);
  if(els.btnSubmitQuizTop) els.btnSubmitQuizTop.addEventListener('click', submitQuiz);
  if(els.btnExitFocus) els.btnExitFocus.addEventListener('click', exitFocus);
  if(els.btnExitFocus2) els.btnExitFocus2.addEventListener('click', exitFocus);
  if(els.btnScrollTopQuiz) els.btnScrollTopQuiz.addEventListener('click', scrollQuizTop);
  if(els.btnScrollTopSticky) els.btnScrollTopSticky.addEventListener('click', scrollCurrentTop);
  if(els.btnExitSticky) els.btnExitSticky.addEventListener('click', exitFocus);
  if(els.btnBackToSetup) els.btnBackToSetup.addEventListener('click', backToSetup);
  if(els.btnBackToSetupSticky) els.btnBackToSetupSticky.addEventListener('click', backToSetup);
  if(els.btnNewQuizResult) els.btnNewQuizResult.addEventListener('click', startNewQuizSameConfig);
  if(els.btnNewQuizSticky) els.btnNewQuizSticky.addEventListener('click', startNewQuizSameConfig);
  if(els.btnClearOldCache) els.btnClearOldCache.addEventListener('click', clearOldCache);
  if(els.btnForceUpdatePWA) els.btnForceUpdatePWA.addEventListener('click', forceUpdatePWA);
  if(els.internetQuestionSelect) els.internetQuestionSelect.addEventListener('change', () => { if(els.internetProposedSource) els.internetProposedSource.value=''; updateInternetQuestionView(); });
  if(els.btnInternetRefreshList) els.btnInternetRefreshList.addEventListener('click', renderInternetTools);
  if(els.btnInternetAutoBuildSource) els.btnInternetAutoBuildSource.addEventListener('click', autoBuildLegalSourceForSelected);
  if(els.btnInternetEnrichAllByLegalRef) els.btnInternetEnrichAllByLegalRef.addEventListener('click', enrichAllSourcesFromLegalColumn);
  if(els.btnInternetOpenSearch) els.btnInternetOpenSearch.addEventListener('click', () => openInternetSearch(false));
  if(els.btnInternetOpenOfficial) els.btnInternetOpenOfficial.addEventListener('click', () => openInternetSearch(true));
  if(els.btnInternetFetchApi) els.btnInternetFetchApi.addEventListener('click', fetchInternetApi);
  if(els.btnInternetSuggestSource) els.btnInternetSuggestSource.addEventListener('click', suggestSourceFromPasted);
  if(els.btnInternetApplySource) els.btnInternetApplySource.addEventListener('click', () => applyInternetSource(false));
  if(els.btnInternetAppendSource) els.btnInternetAppendSource.addEventListener('click', () => applyInternetSource(true));
  if(els.btnInternetSaveDb) els.btnInternetSaveDb.addEventListener('click', saveInternetUpdatedBank);
  if(els.btnInternetExportXlsx) els.btnInternetExportXlsx.addEventListener('click', exportUpdatedBankXlsx);
  if(els.btnInternetExportLinks) els.btnInternetExportLinks.addEventListener('click', exportInternetLinksHtml);
}
function boot(){
  initElements();
  loadStudyState();
  try { parseEmbedded(); }
  catch(e){ console.error(e); setStatus('❌ Lỗi nạp dữ liệu nhúng: ' + (e.message || e), 'bad'); }
  bindEvents();
  registerSW();
  saveBank('dữ liệu nhúng', true).catch(()=>{});
  setStatus(`✅ Đã tự nạp ${state.bank.length} câu hỏi nhúng từ Excel mới. Có thể bấm “Tạo đề” ngay, không cần chọn file Excel.`, 'good');
}
window.addEventListener('DOMContentLoaded', boot);
