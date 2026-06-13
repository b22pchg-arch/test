
'use strict';

const APP_VERSION = 'V7.0-20260613-prefix-sequence-diff';
const DB_NAME = 'excel_quiz_offline_v3_fixed';
const STORE_NAME = 'kv';
const BANK_KEY = 'active_question_bank';
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
  lastResult: null
};

const VI_STOPWORDS = new Set(`
la là cua của va và hoac hoặc de để den đến duoc được bi bị trong ngoai ngoài mot một cac các nhung những ma mà thi thì voi với cho ve về theo tai tại tu từ khi luc lúc nao nào gi gì do đó nay này kia ay ấy tren trên duoi dưới vao vào ra bang bằng nhu như neu nếu hon hơn kem kém da đã dang đang se sẽ can cần phai phải khong không chua chưa cung cùng moi mỗi sau truoc trước phan phần noi nội dung cau câu hoi hỏi phuong phương an án lua lựa chon chọn dap đáp dung đúng sai don đơn vi vị linh lĩnh vuc vực he hệ thong thống hay boi bởi viec việc yeu yêu cau cầu quy quy dinh định
`.split(/\s+/).filter(Boolean));
VI_STOPWORDS.delete('phan'); // giữ được cụm kỹ thuật như "phân phối điện"

function initElements(){
  ['embeddedInfo','status','bankStats','bankPreview','excelFile','sheetSelect','btnReadSheet','btnReadAll','btnUseEmbedded','btnSaveEmbedded','btnSaveBank','btnLoadBank','btnClearBank','manualBox','manualHeaderRow','manualQuestionCol','manualAnswerCol','manualSourceCol','manualOptionCols','btnRefreshMapping','btnApplyMapping','quizCount','shuffleQuestions','shuffleOptions','showAutoExplain','seedInput','btnStartQuiz','quizInfo','quizList','btnSubmitQuiz','btnSubmitSticky','btnExitFocus','btnExitFocus2','btnSubmitQuizTop','btnScrollTopQuiz','btnBackToSetup','btnNewQuizResult','resultSummary','resultList','progressText','progressBar','btnPrint','btnClearOldCache'].forEach(id => els[id] = $(id));
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
  const suff = suffixCount(cSeq, wSeq, pref);
  const prefixPct = Math.round((pref / Math.max(1, cSeq.length)) * 100);
  const first = al.blocks[0];
  const biggest = al.blocks.slice().sort((x,y) => (y.del.length + y.add.length) - (x.del.length + x.add.length))[0];
  const correctMissingInWrong = al.blocks.flatMap(b => b.del || []);
  const wrongOnly = al.blocks.flatMap(b => b.add || []);
  const lines = [];

  const correctStart = cSeq.slice(0, Math.min(3, cSeq.length));
  const wrongStart = wSeq.slice(0, Math.min(3, wSeq.length));
  const samePrefix = cSeq.slice(0, Math.min(pref, 10));
  const cNext = cSeq[pref];
  const wNext = wSeq[pref];
  const cRest = cSeq.slice(pref, Math.min(pref + 8, cSeq.length));
  const wRest = wSeq.slice(pref, Math.min(pref + 8, wSeq.length));

  lines.push(`<div class="diff-block sequence"><b>So khớp chuỗi từ đầu đến cuối:</b> phương án sai trùng <b>${pctWrongMatched}%</b> nội dung của đáp án đúng; bao phủ <b>${pctCorrectCovered}%</b> ý của đáp án đúng. Độ giống cân bằng: <b>${pctBalanced}%</b>. Riêng phần giống liên tiếp từ đầu là <b>${prefixPct}%</b> (${pref}/${cSeq.length} từ/cụm).</div>`);

  if(pref === 0){
    lines.push(`<div class="diff-block"><b>Khác ngay từ đầu:</b><br>Đáp án đúng bắt đầu bằng: ${htmlTermTokenList(correctStart, 'need')}.<br>Phương án sai không có phần bắt đầu này, mà bắt đầu bằng: ${htmlTermTokenList(wrongStart, 'wrongterm') || '<span class="muted">không có nội dung</span>'}.<br><b>Kết luận nhanh:</b> phương án sai lệch ngay từ đối tượng/mốc mở đầu nên không thể thay cho đáp án đúng.</div>`);
  } else if(cNext && wNext){
    lines.push(`<div class="diff-block"><b>Đọc từ đầu đến cuối:</b> hai phương án giống nhau đến trước từ/cụm số <b>${pref + 1}</b> ${samePrefix.length ? '(' + htmlTermTokenList(samePrefix, 'same') + ')' : ''}.<br><b>Từ/cụm bắt đầu khác:</b><br>Đáp án đúng phải có: ${htmlTermTokenList([cNext], 'need')}<br>Phương án sai lại ghi: ${htmlTermTokenList([wNext], 'wrongterm')}<br><b>Điểm cần nhớ:</b> từ/cụm sai đầu tiên này là dấu hiệu tách phương án sai khỏi đáp án đúng.</div>`);
  } else if(!wNext && cNext){
    lines.push(`<div class="diff-block"><b>Phương án sai bị thiếu phần cuối:</b> sau ${pref} từ/cụm đầu đã giống nhau, phương án sai dừng lại hoặc thiếu tiếp nội dung mà đáp án đúng còn yêu cầu: ${htmlTermTokenList(cRest, 'need')}.</div>`);
  } else if(wNext && !cNext){
    lines.push(`<div class="diff-block"><b>Phương án sai thêm phần không cần có:</b> nội dung đáp án đúng đã kết thúc, nhưng phương án sai còn thêm: ${htmlTermTokenList(wRest, 'wrongterm')}.</div>`);
  } else {
    lines.push(`<div class="diff-block"><b>Chuỗi chữ giống hoàn toàn:</b> nếu vẫn bị xác định sai, hãy kiểm tra dấu câu, khoảng trắng, cột đáp án đúng trong Excel hoặc cách chuẩn hóa đáp án.</div>`);
  }

  if(correctMissingInWrong.length){
    lines.push(`<div class="diff-block"><b>Đáp án đúng có nhưng phương án sai thiếu/đổi sai:</b> ${htmlTermTokenList(correctMissingInWrong, 'need')}</div>`);
  }
  if(wrongOnly.length){
    lines.push(`<div class="diff-block"><b>Phương án sai đang dùng thay thế/thêm vào:</b> ${htmlTermTokenList(wrongOnly, 'wrongterm')}</div>`);
  }
  if(biggest && biggest !== first){
    lines.push(`<div class="diff-block"><b>Đoạn lệch lớn nhất:</b><br>Phương án sai ghi: ${htmlTermTokenList(biggest.add, 'wrongterm') || '<span class="muted">thiếu đoạn tương ứng</span>'}<br>Đáp án đúng yêu cầu: ${htmlTermTokenList(biggest.del, 'need') || '<span class="muted">không yêu cầu đoạn này</span>'}</div>`);
  }
  if(suff){
    lines.push(`<div class="diff-block"><b>Phần cuối vẫn giống nhau:</b> ${suff} từ/cụm cuối ${htmlTermTokenList(cSeq.slice(Math.max(cSeq.length-suff,0)).slice(0,10), 'same')}. Phần cuối giống nhau chưa đủ, vì điểm sai đã xuất hiện trước đó.</div>`);
  }
  if(al.blocks.length > 1){
    const more = al.blocks.slice(0, 6).map((b,idx) => {
      const a = seqText(b.add, 8) || '∅';
      const d = seqText(b.del, 8) || '∅';
      return `<div class="small">${idx+1}) Sai ghi: <b>${escapeHtml(a)}</b> → Đúng phải là/cần có: <b>${escapeHtml(d)}</b></div>`;
    }).join('');
    lines.push(`<details><summary>Các điểm sai theo đúng thứ tự đọc từ đầu đến cuối</summary>${more}</details>`);
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
    conclusion = `<b>Nhìn từ phương án sai:</b> phương án này sai chủ yếu vì <b>số liệu/mốc/phạm vi không khớp</b>. Đáp án đúng là đúng vì có ${termListHtml(missingFacts.map(f=>f.raw), 'need') || 'mốc đúng'}, trong khi phương án sai dùng ${termListHtml(addedFacts.map(f=>f.raw), 'wrongterm') || 'mốc khác/thiếu mốc đúng'}.`;
  } else if(missingPol.length || addedPol.length){
    conclusion = `<b>Nhìn từ phương án sai:</b> phương án này sai do <b>đổi điều kiện hoặc tính chất</b>. Đáp án đúng giữ điều kiện ${termListHtml(missingPol,'need')||'cần có'}, còn phương án sai lại thể hiện ${termListHtml(addedPol,'wrongterm')||'điều kiện khác/thiếu điều kiện đó'}.`;
  } else if(missingTerms.length || addedTerms.length){
    conclusion = `<b>Nhìn từ phương án sai:</b> phương án này thiếu hoặc đổi <b>từ khóa/đối tượng quyết định</b>. Đáp án đúng là đúng vì có ${termListHtml(missingTerms,'need')||'ý chính cần có'}, còn phương án sai chuyển sang ${termListHtml(addedTerms,'wrongterm')||'ý khác hoặc thiếu ý chính đó'}.`;
  } else {
    conclusion = '<b>Nhìn từ phương án sai:</b> hai phương án rất gần nhau; điểm đúng/sai có thể nằm ở một từ nhỏ, dấu phủ định, hoặc căn cứ pháp lý. Cần đối chiếu lại cột đáp án đúng của Excel.';
  }

  const parts = [`<div class="diff-conclusion">${conclusion}</div>`];
  parts.push(sequenceExplainHtml(correct, wrong));
  if(proofItems.length){
    parts.push(`<div class="diff-block"><b>Vì sao đáp án đúng đúng hơn:</b> nó có ${termListHtml(proofItems, 'need')}. Đây là phần phương án sai không có, thiếu, hoặc đã đổi sai.</div>`);
  }
  if(wrongItems.length){
    parts.push(`<div class="diff-block"><b>Vì sao phương án này không được chọn:</b> nó dùng ${termListHtml(wrongItems, 'wrongterm')}, không trùng với ý/mốc bắt buộc của đáp án đúng.</div>`);
  }
  if(sameFacts.length || sameTerms.length){
    parts.push(`<div class="diff-block"><b>Phần giống nhau chỉ để gây nhầm:</b> ${termListHtml([...sameFacts.map(f=>f.raw), ...sameTerms], 'same') || 'rất ít nội dung trùng nhau'}. Phần giống này chưa đủ làm phương án sai trở thành đúng.</div>`);
  }
  return parts.filter(Boolean).join('');
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

function renderAll(){ renderStats(); renderPreview(); resetQuiz(); }
function renderStats(){
  if(!els.bankStats) return;
  const src = state.currentSheet || state.fileName || 'Dữ liệu';
  els.bankStats.innerHTML = `<span class="pill okp">${state.bank.length} câu hợp lệ</span><span class="pill ${state.errors.length?'warnp':''}">${state.errors.length} dòng lỗi</span><span class="pill">${escapeHtml(src)}</span><span class="pill">${APP_VERSION}</span>`;
  if(els.btnStartQuiz) els.btnStartQuiz.disabled = state.bank.length === 0;
  if(els.btnSaveBank) els.btnSaveBank.disabled = state.bank.length === 0;
  if(els.quizCount){ els.quizCount.max = String(Math.max(1,state.bank.length)); if(Number(els.quizCount.value || 0) > state.bank.length) els.quizCount.value = String(state.bank.length || 1); }
}
function renderPreview(){
  if(!els.bankPreview) return;
  if(!state.bank.length){ els.bankPreview.innerHTML = '<table><tbody><tr><td class="muted">Chưa có dữ liệu hợp lệ.</td></tr></tbody></table>'; return; }
  const rows = state.bank.slice(0,50).map((q,i)=>`<tr><td class="nowrap">${i+1}</td><td>${escapeHtml(q.question)}</td><td>${escapeHtml(q.correctText)}</td><td>${escapeHtml(q.options.join(' | '))}</td><td>${escapeHtml(q.source||'')}</td></tr>`).join('');
  els.bankPreview.innerHTML = `<table><thead><tr><th>STT</th><th>Câu hỏi</th><th>Đáp án đúng</th><th>Các phương án</th><th>Căn cứ / giải thích</th></tr></thead><tbody>${rows}</tbody></table>`;
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
  if(!state.bank.length){ setStatus('Chưa có ngân hàng câu hỏi.', 'bad'); return; }
  const count = Math.max(1, Math.min(Number(els.quizCount.value || 1), state.bank.length));
  const seedText = els.seedInput.value.trim() || new Date().toISOString();
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
function resetQuiz(){
  state.quiz=[]; state.submitted=false; state.lastResult=null;
  if(els.quizInfo) els.quizInfo.textContent='Chưa tạo đề.';
  if(els.quizList) els.quizList.innerHTML='';
  if(els.resultSummary) els.resultSummary.textContent='Chưa nộp bài.';
  if(els.resultList) els.resultList.innerHTML='';
  if(els.btnSubmitQuiz) els.btnSubmitQuiz.disabled=true;
  if(els.btnSubmitSticky) els.btnSubmitSticky.disabled=true;
  if(els.btnSubmitQuizTop) els.btnSubmitQuizTop.disabled=true;
  updateProgress();
}
function renderQuiz(){
  if(!state.quiz.length){ els.quizInfo.textContent='Chưa tạo đề.'; els.quizList.innerHTML=''; updateProgress(); return; }
  els.quizInfo.innerHTML = `Đề gồm <b>${state.quiz.length}</b> câu. ${els.shuffleQuestions.checked?'Có đảo câu hỏi.':'Không đảo câu hỏi.'} ${els.shuffleOptions.checked?'Có đảo phương án.':'Không đảo phương án.'}`;
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
  renderQuiz(); renderResult();
  els.btnSubmitQuiz.disabled = true; els.btnSubmitSticky.disabled = true; if(els.btnSubmitQuizTop) els.btnSubmitQuizTop.disabled = true;
  enterResultFocus();
  location.hash = '#resultSection';
}
function renderResult(){
  const r=state.lastResult; if(!r) return;
  els.resultSummary.innerHTML = `<span class="pill okp">Đúng ${r.correct}/${r.total}</span><span class="pill">Điểm ${r.score10.toFixed(2)}/10</span><span class="pill">${escapeHtml(r.time)}</span><br><span class="muted">Bảng giải thích mới tập trung vào điểm khác biệt cơ bản: số liệu/phạm vi, điều kiện/tính chất, hoặc từ khóa/đối tượng chính.</span>`;
  els.resultList.innerHTML = state.quiz.map((q, qi) => {
    const chosen = q.userChoice === null ? null : q.options[q.userChoice];
    const correctOpt = q.options.find(o => o.isCorrect) || q.options[0];
    const ok = !!(chosen && chosen.isCorrect);
    const chosenDiff = (!ok && chosen) ? `<div class="analysis-row focus-diff"><b>Phân tích từ lựa chọn sai của bạn:</b><br>${els.showAutoExplain.checked ? explainDifference(correctOpt.text, chosen.text) : 'Đã tắt phân tích tự động.'}</div>` : '';
    const rows = q.options.map((op, oi) => {
      const letter = 'ABCDEF'[oi] || String(oi+1);
      const status = op.isCorrect ? '<span class="ok">Đáp án đúng</span>' : (q.userChoice===oi ? '<span class="bad">Bạn đã chọn</span>' : '<span class="muted">Phương án sai</span>');
      const explain = op.isCorrect ? `<div class="analysis-row source"><b>Vì sao đúng:</b> Đây là phương án được cột đáp án trong Excel xác định là đúng; các phương án sai bị loại vì thiếu, đổi sai hoặc thêm ý không khớp với phương án này.${q.item.source?'<br><b>Căn cứ:</b> '+escapeHtml(q.item.source):''}</div>` : `<div class="analysis-row">${els.showAutoExplain.checked ? explainDifference(correctOpt.text, op.text) : 'Đã tắt phân tích tự động.'}</div>`;
      const trClass = op.isCorrect ? 'result-status-ok' : (q.userChoice===oi ? 'result-status-bad' : '');
      return `<tr class="${trClass}"><td class="nowrap"><b>${letter}</b></td><td>${escapeHtml(op.text)}</td><td>${status}</td><td>${explain}</td></tr>`;
    }).join('');
    return `<article class="question-card"><h3>Câu ${qi+1}. ${ok?'<span class="ok">Đúng</span>':'<span class="bad">Sai / chưa chọn</span>'}</h3><p><b>${escapeHtml(q.item.question)}</b></p><div class="result-question-meta"><span class="pill ${ok?'okp':'warnp'}">${ok?'Bạn chọn đúng':'Cần xem lại'}</span><span class="pill">Đáp án đúng: ${escapeHtml(correctOpt.text)}</span>${chosen?`<span class="pill">Bạn chọn: ${escapeHtml(chosen.text)}</span>`:'<span class="pill warnp">Bạn chưa chọn</span>'}</div>${chosenDiff}<div class="table-wrap"><table><thead><tr><th>PA</th><th>Nội dung</th><th>Trạng thái</th><th>Giải thích trọng tâm</th></tr></thead><tbody>${rows}</tbody></table></div></article>`;
  }).join('');
}

function enterQuizFocus(){
  document.body.classList.remove('result-fullscreen');
  document.body.classList.add('quiz-fullscreen');
  setTimeout(() => { const q = $('quizSection'); if(q) q.scrollTo({top:0, behavior:'smooth'}); }, 0);
}
function enterResultFocus(){
  document.body.classList.remove('quiz-fullscreen');
  document.body.classList.add('result-fullscreen');
  setTimeout(() => { const r = $('resultSection'); if(r) r.scrollTo({top:0, behavior:'smooth'}); }, 0);
}
function exitFocus(){
  document.body.classList.remove('quiz-fullscreen','result-fullscreen');
}
function scrollQuizTop(){
  const q = $('quizSection'); if(q) q.scrollTo({top:0, behavior:'smooth'});
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
  els.btnSubmitQuiz.addEventListener('click', submitQuiz);
  els.btnSubmitSticky.addEventListener('click', submitQuiz);
  els.btnPrint.addEventListener('click', () => window.print());
  if(els.btnSubmitQuizTop) els.btnSubmitQuizTop.addEventListener('click', submitQuiz);
  if(els.btnExitFocus) els.btnExitFocus.addEventListener('click', exitFocus);
  if(els.btnExitFocus2) els.btnExitFocus2.addEventListener('click', exitFocus);
  if(els.btnScrollTopQuiz) els.btnScrollTopQuiz.addEventListener('click', scrollQuizTop);
  if(els.btnBackToSetup) els.btnBackToSetup.addEventListener('click', backToSetup);
  if(els.btnNewQuizResult) els.btnNewQuizResult.addEventListener('click', backToSetup);
  els.btnClearOldCache.addEventListener('click', clearOldCache);
}
function boot(){
  initElements();
  try { parseEmbedded(); }
  catch(e){ console.error(e); setStatus('❌ Lỗi nạp dữ liệu nhúng: ' + (e.message || e), 'bad'); }
  bindEvents();
  registerSW();
  saveBank('dữ liệu nhúng', true).catch(()=>{});
  setStatus(`✅ Đã tự nạp ${state.bank.length} câu hỏi nhúng từ Excel mới. Có thể bấm “Tạo đề” ngay, không cần chọn file Excel.`, 'good');
}
window.addEventListener('DOMContentLoaded', boot);
