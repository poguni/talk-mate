/* 토크메이트(TalkMate): 상황극으로 연습하는 AI 외국어 회화
 * - 서버·로그인·DB 없음. OpenRouter를 브라우저에서 직접 호출합니다.
 * - API 키: 기본은 sessionStorage(탭을 닫으면 삭제), 선택 시 localStorage.
 * - 설정은 localStorage에 기억하고, 진행 중인 대화는 sessionStorage에만 임시 보관합니다.
 */
'use strict';

/* ───────────── 상수 ───────────── */
const OR_BASE = 'https://openrouter.ai/api/v1';
const STORE = {
  settings: 'talkmate:settings',
  key: 'talkmate:apikey',       // sessionStorage 또는 localStorage
  session: 'talkmate:session',  // sessionStorage
};
const CDN = {
  xlsx: 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  docx: 'https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js',
};

const MODELS = {
  sonnet: 'anthropic/claude-sonnet-5.5',
  luna: 'openai/gpt-6-luna',
  flash: 'google/gemini-3.8-flash',
};
const PRESETS = {
  standard: { label: '기본', desc: '대화 Gemini 3.8 Flash · 교정 GPT-6 Luna · 리포트 Claude Sonnet 5.5',
    partner: MODELS.flash, coach: MODELS.luna, report: MODELS.sonnet },
  quality: { label: '고품질', desc: '대화 Gemini 3.8 Flash · 교정·리포트 Claude Sonnet 5.5',
    partner: MODELS.flash, coach: MODELS.sonnet, report: MODELS.sonnet },
  economy: { label: '절약', desc: '모든 역할 GPT-6 Luna',
    partner: MODELS.luna, coach: MODELS.luna, report: MODELS.luna },
};
// 100만 토큰당 USD. 연결 테스트 때 OpenRouter 모델 목록의 실제 단가로 갱신합니다.
const DEFAULT_PRICES = {
  [MODELS.sonnet]: { in: 2, out: 10 },
  [MODELS.luna]: { in: 0.1, out: 0.5 },
  [MODELS.flash]: { in: 0.75, out: 3.75 },
};

const LANGS = {
  en: { name: '영어', native: 'English', en: 'English', flag: '🇺🇸', stt: 'en-US', tts: 'en-US', level: 'A2+',
    tones: { auto: '상황에 맞게', polite: '정중하게', casual: '친근하게' } },
  ja: { name: '일본어', native: '日本語', en: 'Japanese', flag: '🇯🇵', stt: 'ja-JP', tts: 'ja-JP', level: 'B1',
    tones: { auto: '상황에 맞게', polite: '존댓말 (です・ます)', casual: '반말 (タメ口)' } },
  es: { name: '스페인어', native: 'Español', en: 'Spanish', flag: '🇪🇸', stt: 'es-ES', tts: 'es-ES', level: 'A1',
    tones: { auto: '상황에 맞게', polite: '격식 (usted)', casual: '친근 (tú)' } },
};
const TONE_PROMPT = {
  en: { polite: 'Use polite, formal English.', casual: 'Use friendly, casual English.' },
  ja: { polite: 'Use polite Japanese (です・ます form).', casual: 'Use casual Japanese (タメ口, plain form).' },
  es: { polite: 'Address the learner formally with "usted".', casual: 'Address the learner informally with "tú".' },
};
const LEVELS = [
  { id: 'A1', ko: '입문·초급', jlpt: 'N5', guide: 'very basic words and short, simple present-tense sentences; speak slowly and clearly' },
  { id: 'A2', ko: '초급', jlpt: 'N4', guide: 'simple everyday vocabulary and short sentences; basic past and future' },
  { id: 'A2+', ko: '초중급', jlpt: 'N4~N3', guide: 'everyday vocabulary, mostly simple sentences with some connectors (between A2 and B1)' },
  { id: 'B1', ko: '중급', jlpt: 'N3', guide: 'common everyday and travel vocabulary, connected sentences, a few common idioms' },
  { id: 'B2', ko: '중상급', jlpt: 'N2', guide: 'natural pace, varied grammar and vocabulary, some idiomatic expressions' },
  { id: 'C1', ko: '고급', jlpt: 'N1', guide: 'fully natural, idiomatic native-like speech' },
];
const LENGTHS = {
  short: '1-2 short sentences',
  normal: '2-3 sentences',
  long: '3-5 sentences',
};
const ERROR_TYPES = ['문법', '시제', '활용', '어순', '어휘 선택', '성·수 일치', '관사·한정사', '전치사', '조사', '높임말·말투', '철자·표기', '기타'];

const SCENARIOS = [
  { id: 'airport', cat: '여행', icon: '✈️', title: '공항 체크인', desc: '좌석·수하물·탑승구',
    en: 'checking in for a flight at an airline counter', ai: 'an airline check-in agent', user: 'a traveler',
    missions: ['창가 또는 통로 좌석을 요청하기', '수하물 개수와 무게 제한을 묻기', '탑승구와 탑승 시간을 확인하기'], travel: true },
  { id: 'immigration', cat: '여행', icon: '🛂', title: '입국 심사', desc: '방문 목적·체류 기간',
    en: 'passing through immigration control after landing', ai: 'an immigration officer', user: 'a tourist arriving in the country',
    missions: ['방문 목적을 말하기', '체류 기간을 말하기', '숙소 주소(호텔 이름)를 말하기'], travel: true },
  { id: 'hotel', cat: '여행', icon: '🏨', title: '호텔 체크인·문제 해결', desc: '예약 확인·조식·고장 신고',
    en: 'checking in at a hotel front desk and later reporting a problem with the room', ai: 'a hotel front desk clerk', user: 'a hotel guest',
    missions: ['이름으로 예약을 확인하기', '조식 시간과 장소를 묻기', '방의 문제(예: 에어컨 고장)를 알리고 해결을 요청하기'], travel: true },
  { id: 'restaurant', cat: '여행', icon: '🍽️', title: '식당 주문·계산', desc: '추천 메뉴·알레르기·계산',
    en: 'ordering food at a local restaurant and paying the bill', ai: 'a friendly waiter', user: 'a customer',
    missions: ['추천 메뉴를 묻기', '음식 알레르기나 못 먹는 재료를 말하기', '음료를 포함해 주문하기', '계산서를 요청하고 카드로 결제하기'], travel: true },
  { id: 'directions', cat: '여행', icon: '🚇', title: '길 묻기·대중교통', desc: '길 찾기·표 사기',
    en: 'asking a local for directions and how to use public transport', ai: 'a helpful local passer-by', user: 'a lost tourist',
    missions: ['목적지까지 가는 길을 묻기', '어떤 교통수단을 타야 하는지 묻기', '걸리는 시간이나 요금을 묻기'], travel: true },
  { id: 'shopping', cat: '여행', icon: '🛍️', title: '쇼핑·교환·환불', desc: '사이즈·가격·환불',
    en: 'shopping for clothes and then asking to exchange or refund an item', ai: 'a shop assistant', user: 'a customer',
    missions: ['다른 사이즈나 색상이 있는지 묻기', '가격이나 할인 여부를 묻기', '교환 또는 환불을 요청하기'], travel: true },
  { id: 'pharmacy', cat: '여행', icon: '💊', title: '약국·병원', desc: '증상 설명·복용법',
    en: 'explaining symptoms at a pharmacy while traveling', ai: 'a pharmacist', user: 'a traveler who feels sick',
    missions: ['증상과 시작된 시점을 설명하기', '약 복용법(횟수·시간)을 묻기', '부작용이나 주의사항을 확인하기'], travel: true },
  { id: 'smalltalk', cat: '일상', icon: '👋', title: '자기소개·스몰토크', desc: '처음 만난 사람과 대화',
    en: 'meeting a friendly local at a cafe and making small talk', ai: 'a friendly local person of a similar age', user: 'a visitor from Korea',
    missions: ['자기소개하기(이름·사는 곳·직업)', '여행이나 방문 이유를 말하기', '상대에게 질문을 두 번 이상 하기'], travel: true },
  { id: 'hobby', cat: '일상', icon: '📷', title: '취미 이야기', desc: '좋아하는 것·이유·추천',
    en: 'chatting about hobbies (for example photography, reading, travel, music)', ai: 'a new friend who is curious about the learner', user: 'themself',
    missions: ['내 취미를 소개하기', '그 취미를 좋아하는 이유를 설명하기', '상대에게 추천을 부탁하기'] },
  { id: 'phone', cat: '일상', icon: '📞', title: '전화 예약', desc: '날짜·시간·인원',
    en: 'calling a popular restaurant on the phone to make a reservation', ai: 'a restaurant staff member answering the phone', user: 'a customer calling',
    missions: ['날짜·시간·인원을 말하기', '이름과 연락처를 전달하기', '특별 요청(창가 자리 등)을 하고 예약 내용을 확인하기'], travel: true },
  { id: 'free', cat: '학습', icon: '💬', title: '자유 대화', desc: '주제를 정해 대화',
    en: 'a relaxed free conversation about a chosen topic', ai: 'a friendly conversation partner', user: 'themself',
    missions: ['주제에 대한 내 경험을 말하기', '내 의견과 이유를 말하기', '상대에게 질문하기'], topic: true },
  { id: 'debate', cat: '학습', icon: '⚖️', title: '찬반 토론', desc: '의견·근거·반론',
    en: 'a friendly debate on a chosen topic; the AI takes the opposite side', ai: 'a debate partner who politely takes the opposite position', user: 'a debater',
    missions: ['내 입장을 분명히 말하기', '근거를 두 가지 이상 말하기', '상대의 반론에 대답하기'], topic: true },
  { id: 'custom', cat: '학습', icon: '✏️', title: '직접 만들기', desc: '상황·역할·미션 직접 입력', custom: true },
];

const DEFAULT_SETTINGS = {
  keyPersist: false,
  lang: 'en',
  levels: { en: 'A2+', ja: 'B1', es: 'A1' },
  tones: { en: 'auto', ja: 'auto', es: 'auto' },
  length: 'short',
  scenario: 'restaurant',
  city: '',
  topic: '',
  custom: { place: '', ai: '', user: '', missions: '' },
  preset: 'standard',
  models: { partner: PRESETS.standard.partner, coach: PRESETS.standard.coach, report: PRESETS.standard.report },
  prices: { ...DEFAULT_PRICES },
  fx: 1400,
  autoSpeak: true,
  autoSend: false,
  rate: 0.95,
  voices: { en: '', ja: '', es: '' },
  turns: 12,
};

/* ───────────── 상태 ───────────── */
let settings = loadSettings();
let session = loadSession();
let partnerBusy = false;
let modelCatalog = null; // OpenRouter 모델 목록 캐시

/* ───────────── 유틸 ───────────── */
const $ = (sel) => document.querySelector(sel);
function el(tag, attrs, ...children) {
  const node = document.createElement(tag);
  if (attrs) {
    for (const [k, v] of Object.entries(attrs)) {
      if (v == null || v === false) continue;
      if (k === 'class') node.className = v;
      else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.slice(2), v);
      else if (k === 'text') node.textContent = v;
      else node.setAttribute(k, v === true ? '' : v);
    }
  }
  for (const c of children.flat()) {
    if (c == null || c === false) continue;
    node.append(c instanceof Node ? c : document.createTextNode(String(c)));
  }
  return node;
}
function uid() { return Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4); }
function deepMerge(base, extra) {
  const out = Array.isArray(base) ? [...base] : { ...base };
  if (!extra || typeof extra !== 'object') return out;
  for (const [k, v] of Object.entries(extra)) {
    if (v && typeof v === 'object' && !Array.isArray(v) && base[k] && typeof base[k] === 'object') out[k] = deepMerge(base[k], v);
    else out[k] = v;
  }
  return out;
}
function safeGet(storage, key) { try { return storage.getItem(key); } catch { return null; } }
function safeSet(storage, key, val) {
  try { storage.setItem(key, val); return true; } catch (e) { console.warn('저장 실패', e); return false; }
}
function safeDel(storage, key) { try { storage.removeItem(key); } catch { /* 무시 */ } }

let toastTimer;
function toast(msg, isError) {
  const t = $('#toast');
  t.textContent = msg;
  t.className = 'show' + (isError ? ' error' : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.className = ''; }, isError ? 5000 : 2600);
}
function fmtUSD(v) {
  if (!v) return '$0';
  return v < 0.01 ? '$' + v.toFixed(4) : '$' + v.toFixed(3);
}
function fmtKRW(v) { return Math.round(v * (settings.fx || 1400)).toLocaleString('ko-KR') + '원'; }
function pad(n) { return String(n).padStart(2, '0'); }
function fmtDate(ts, withTime = true) {
  const d = new Date(ts);
  const s = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  return withTime ? `${s} ${pad(d.getHours())}:${pad(d.getMinutes())}` : s;
}
function levelInfo(id) { return LEVELS.find((l) => l.id === id) || LEVELS[1]; }
function levelLabel(lang, id) {
  const l = levelInfo(id);
  return `${l.id} ${l.ko}` + (lang === 'ja' ? ` (JLPT ${l.jlpt})` : '');
}

/* ───────────── 저장소 ───────────── */
function loadSettings() {
  let saved = null;
  try { saved = JSON.parse(safeGet(localStorage, STORE.settings) || 'null'); } catch { saved = null; }
  return deepMerge(DEFAULT_SETTINGS, saved || {});
}
function saveSettings() { safeSet(localStorage, STORE.settings, JSON.stringify(settings)); }

function getKey() { return safeGet(sessionStorage, STORE.key) || safeGet(localStorage, STORE.key) || ''; }
function setKey(key, persist) {
  safeSet(sessionStorage, STORE.key, key);
  if (persist) safeSet(localStorage, STORE.key, key);
  else safeDel(localStorage, STORE.key);
  settings.keyPersist = !!persist;
  saveSettings();
}
function clearKey() { safeDel(sessionStorage, STORE.key); safeDel(localStorage, STORE.key); }

function loadSession() {
  try { return JSON.parse(safeGet(sessionStorage, STORE.session) || 'null'); } catch { return null; }
}
function saveSession() {
  if (!session) { safeDel(sessionStorage, STORE.session); return; }
  safeSet(sessionStorage, STORE.session, JSON.stringify(session));
}

/* ───────────── 화면 전환 ───────────── */
function showView(name) {
  for (const v of document.querySelectorAll('.view')) v.classList.add('hidden');
  $('#view-' + name).classList.remove('hidden');
  window.scrollTo(0, 0);
  updateCostBadge();
}

/* ───────────── OpenRouter 호출 ───────────── */
class ApiError extends Error {
  constructor(status, message) { super(message); this.status = status; }
}
function friendlyError(e) {
  if (e instanceof ApiError) {
    if (e.status === 401) return 'API 키가 올바르지 않습니다. 설정에서 키를 확인해 주세요.';
    if (e.status === 402) return 'OpenRouter 크레딧이 부족합니다. 충전하거나 키의 사용 한도를 확인해 주세요.';
    if (e.status === 429) return '요청이 너무 많습니다. 잠시 후 다시 시도해 주세요.';
    if (e.status === 400 || e.status === 404) return `요청 오류: ${e.message || '모델 ID를 확인해 주세요.'}`;
    return `OpenRouter 오류(${e.status}): ${e.message || '잠시 후 다시 시도해 주세요.'}`;
  }
  if (e && e.name === 'TypeError') return '네트워크에 연결할 수 없습니다. 인터넷 연결을 확인해 주세요.';
  return (e && e.message) || '알 수 없는 오류가 발생했습니다.';
}
function headers() {
  const h = { 'Authorization': 'Bearer ' + getKey(), 'Content-Type': 'application/json', 'X-Title': 'TalkMate' };
  if (location.protocol.startsWith('http')) h['HTTP-Referer'] = location.origin + location.pathname;
  return h;
}

/**
 * Chat Completions 호출. stream=true면 onDelta로 조각을 넘깁니다.
 * role은 비용 기록용(대화/교정/힌트/번역/리포트/미션).
 */
async function chat({ model, messages, stream = false, onDelta, maxTokens = 1000, temperature = 0.7, role = '기타', retries = 1 }) {
  if (!getKey()) throw new Error('API 키가 없습니다. 설정에서 키를 입력해 주세요.');
  const body = { model, messages, stream, max_tokens: maxTokens, temperature, usage: { include: true } };
  let res;
  for (let attempt = 0; ; attempt++) {
    try {
      res = await fetch(OR_BASE + '/chat/completions', { method: 'POST', headers: headers(), body: JSON.stringify(body) });
      if (res.ok) break;
      let msg = '';
      try { const j = await res.json(); msg = (j.error && (j.error.message || j.error.code)) || ''; } catch { /* 무시 */ }
      const err = new ApiError(res.status, msg);
      if (attempt < retries && (res.status === 429 || res.status >= 500)) { await sleep(1500); continue; }
      throw err;
    } catch (e) {
      if (e instanceof ApiError) throw e;
      if (attempt < retries) { await sleep(1200); continue; }
      throw e;
    }
  }

  let text = '';
  let usage = null;
  if (!stream) {
    const j = await res.json();
    if (j.error) throw new ApiError(j.error.code || 500, j.error.message);
    text = (j.choices && j.choices[0] && j.choices[0].message && j.choices[0].message.content) || '';
    usage = j.usage || null;
  } else {
    const reader = res.body.getReader();
    const dec = new TextDecoder();
    let buf = '';
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += dec.decode(value, { stream: true });
      let idx;
      while ((idx = buf.indexOf('\n')) >= 0) {
        const line = buf.slice(0, idx).trim();
        buf = buf.slice(idx + 1);
        if (!line.startsWith('data:')) continue; // ": OPENROUTER PROCESSING" 같은 주석 무시
        const data = line.slice(5).trim();
        if (data === '[DONE]') continue;
        let j;
        try { j = JSON.parse(data); } catch { continue; }
        if (j.error) throw new ApiError(j.error.code || 500, j.error.message);
        const d = j.choices && j.choices[0] && j.choices[0].delta && j.choices[0].delta.content;
        if (d) { text += d; if (onDelta) onDelta(d); }
        if (j.usage) usage = j.usage;
      }
    }
  }
  recordCost(role, model, usage, messages, text);
  return text;
}
function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

function priceOf(model) { return settings.prices[model] || null; }
function roughTokens(str) { return Math.ceil((str || '').length / 3); }
function recordCost(role, model, usage, messages, text) {
  if (!session) return;
  const p = priceOf(model);
  const inT = usage && usage.prompt_tokens != null ? usage.prompt_tokens : roughTokens(messages.map((m) => m.content).join(' '));
  const outT = usage && usage.completion_tokens != null ? usage.completion_tokens : roughTokens(text);
  let cost = usage && typeof usage.cost === 'number' ? usage.cost : null;
  const estimated = cost == null;
  if (cost == null) cost = p ? (inT * p.in + outT * p.out) / 1e6 : 0;
  session.costs.push({ role, model, in: inT, out: outT, cost, estimated, ts: Date.now() });
  saveSession();
  updateCostBadge();
  renderLiveCost();
}
function totalCost() { return session ? session.costs.reduce((a, c) => a + (c.cost || 0), 0) : 0; }
function updateCostBadge() {
  const b = $('#costBadge');
  const onChat = !$('#view-chat').classList.contains('hidden') || !$('#view-report').classList.contains('hidden');
  if (!session || !onChat) { b.classList.add('hidden'); return; }
  const t = totalCost();
  b.textContent = `💰 ${fmtUSD(t)} · ${fmtKRW(t)}`;
  b.classList.remove('hidden');
}

/** 응답 텍스트에서 JSON 객체를 꺼냅니다(코드 블록·앞뒤 설명 허용). */
function parseJSON(text) {
  if (!text) return null;
  let t = text.trim().replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '');
  try { return JSON.parse(t); } catch { /* 다음 시도 */ }
  const s = t.indexOf('{');
  const e = t.lastIndexOf('}');
  if (s >= 0 && e > s) { try { return JSON.parse(t.slice(s, e + 1)); } catch { /* 실패 */ } }
  return null;
}

/* ───────────── 프롬프트 ───────────── */
function scenarioOf(s) {
  if (s.scenario.custom) return s.scenario;
  return SCENARIOS.find((x) => x.id === s.scenario.id) || s.scenario;
}
function partnerSystemPrompt(s) {
  const L = LANGS[s.lang];
  const lv = levelInfo(s.level);
  const sc = s.scenario;
  const place = s.city ? ` The scene takes place in ${s.city}; use realistic local details (place names, currency, customs).` : '';
  const topic = s.topic ? ` Topic: "${s.topic}".` : '';
  const tone = s.tone !== 'auto' && TONE_PROMPT[s.lang][s.tone] ? TONE_PROMPT[s.lang][s.tone] : 'Use the register a real person would use in this situation.';
  return [
    `You are a role-play conversation partner for a Korean learner practicing ${L.en}.`,
    `Situation: ${sc.en}.${place}${topic}`,
    `You play: ${sc.ai}. The learner plays: ${sc.user}.`,
    `Learner level: CEFR ${lv.id}${s.lang === 'ja' ? ` (about JLPT ${lv.jlpt})` : ''}. Use ${lv.guide}.`,
    'Rules:',
    `- Speak ONLY in ${L.en}. Stay in character. Do not mention that you are an AI unless asked.`,
    `- Each turn: ${LENGTHS[s.length] || LENGTHS.short}. End most turns with a question or a prompt so the learner keeps talking.`,
    '- If the learner makes a mistake, do NOT explain or point it out. Naturally reuse the correct form in your reply (recast) and continue.',
    `- If the learner writes in Korean or seems stuck, help briefly with a simpler ${L.en} question.`,
    `- ${tone}`,
    `- Gently steer the conversation so the learner can accomplish these goals (written in Korean): ${s.missions.map((m, i) => `${i + 1}) ${m.text}`).join('; ')}.`,
    '- When all goals seem done, wrap up the scene naturally but keep answering if the learner continues.',
    '- Output only what your character says: plain text, no translations, no romanization, no markdown, no stage directions, no emojis.',
  ].join('\n');
}
function partnerMessages(s) {
  const msgs = [{ role: 'system', content: partnerSystemPrompt(s) },
    { role: 'user', content: '(Start the role-play now. Greet me in character and open the conversation.)' }];
  for (const m of s.messages) {
    if (m.pending || m.error || !m.text) continue;
    msgs.push({ role: m.role === 'ai' ? 'assistant' : 'user', content: m.text });
  }
  return msgs;
}
function recentContext(s, beforeId, n = 4) {
  const idx = s.messages.findIndex((m) => m.id === beforeId);
  const list = s.messages.slice(0, idx < 0 ? s.messages.length : idx).filter((m) => m.text && !m.error).slice(-n);
  return list.map((m) => `${m.role === 'ai' ? 'Partner' : 'Learner'}: ${m.text}`).join('\n') || '(none)';
}
function coachPrompt(s, msg) {
  const L = LANGS[s.lang];
  const lv = levelInfo(s.level);
  return [
    `You are a precise, encouraging ${L.en} teacher for a Korean learner at CEFR ${lv.id}${s.lang === 'ja' ? ` (JLPT ${lv.jlpt})` : ''}.`,
    `Role-play situation: ${s.scenario.en}. The learner plays ${s.scenario.user}.`,
    'Analyze ONLY the learner\'s latest message below. Return ONLY one JSON object (no markdown) with this shape:',
    '{"corrected": "full corrected message (same as original if nothing is wrong)",',
    ' "errors": [{"wrong": "exact substring copied from the original", "fix": "replacement", "type": "one of the types", "explain_ko": "한국어로 한 문장 설명"}],',
    ' "natural": ["up to 2 more natural ways a native speaker would say it, suitable for the learner\'s level"],',
    ' "note_ko": "짧은 한국어 팁 (없으면 빈 문자열)",',
    ' "praise_ko": "잘한 점이 있으면 짧은 한국어 칭찬 (없으면 빈 문자열)",',
    ' "missions_done": [0-based indices of the goals below that this message clearly accomplished]}',
    `Error types (use exactly one of these Korean labels): ${ERROR_TYPES.join(', ')}.`,
    'Rules:',
    '- The message may come from speech recognition: ignore missing punctuation, capitalization, and obvious recognition artifacts.',
    '- Only real errors go in "errors". Correct but unnatural phrasing goes in "natural" instead.',
    '- If there are no errors, use "errors": [] and "corrected" equal to the original.',
    `- If the learner wrote in Korean (or mixed Korean), use "errors": [] and put how to say it in ${L.en} in "natural".`,
    '- Do not suggest expressions far above the learner\'s level.',
    `Goals: ${s.missions.map((m, i) => `${i}. ${m.text}`).join(' / ')}`,
    'Recent conversation:',
    recentContext(s, msg.id),
    `Learner's latest message: """${msg.text}"""`,
  ].join('\n');
}
function hintPrompt(s) {
  const L = LANGS[s.lang];
  const lv = levelInfo(s.level);
  const open = s.missions.filter((m) => !m.done).map((m) => m.text);
  return [
    `A Korean learner (CEFR ${lv.id}) is in a ${L.en} role-play: ${s.scenario.en}. The learner plays ${s.scenario.user}.`,
    'Recent conversation:',
    recentContext(s, null, 6),
    `Remaining goals (Korean): ${open.join(' / ') || '(all done)'}`,
    `Suggest 3 different things the learner could say next in ${L.en}, at their level, that fit the conversation (and help with a remaining goal if possible).`,
    'Return ONLY JSON: {"hints":[{"text":"...", "reading":"' + (s.lang === 'ja' ? 'hiragana reading of the whole sentence' : '') + '", "ko":"한국어 뜻과 짧은 사용 팁"}]}',
    s.lang === 'ja' ? '' : 'Use an empty string for "reading".',
  ].join('\n');
}
function reportPrompt(s) {
  const L = LANGS[s.lang];
  const lv = levelInfo(s.level);
  const lines = s.messages.filter((m) => m.text && !m.error).map((m) => {
    if (m.role === 'ai') return `Partner: ${m.text}`;
    const c = m.correction;
    const fix = c && c.errors && c.errors.length ? `  [corrected: ${c.corrected}]` : '';
    return `Learner: ${m.text}${fix}`;
  });
  return [
    `You are a supportive ${L.en} teacher writing a feedback report in Korean for a Korean learner (CEFR ${lv.id}).`,
    `Role-play: ${s.scenario.en}. The learner played ${s.scenario.user}.`,
    `Goals (status tracked during the chat): ${s.missions.map((m, i) => `${i}. ${m.text} [${m.done ? 'done' : 'not marked'}]`).join(' / ')}`,
    'Transcript:',
    lines.join('\n'),
    'Return ONLY one JSON object (no markdown):',
    '{"summary_ko": "2-3문장 총평",',
    ' "level_comment_ko": "현재 수준에 대한 한 문장 코멘트",',
    ' "missions": [{"goal": "미션 내용", "done": true, "comment_ko": "짧은 코멘트"}],',
    ' "strengths_ko": ["잘한 점 1~2개"],',
    ' "improve_ko": ["다음에 연습할 점 1~2개 (구체적으로)"],',
    ` "expressions": [{"text": "${L.en} expression worth memorizing (include corrected forms of the learner's mistakes)", "meaning_ko": "뜻", "example": "short example sentence in ${L.en}"}],`,
    ' "next": {"scenario_ko": "다음에 추천하는 상황", "tip_ko": "추천 이유나 팁"}}',
    'Give 5-8 expressions. Judge each goal from the transcript itself.',
  ].join('\n');
}

/* ───────────── 음성 합성(TTS) ───────────── */
let voices = [];
function loadVoices() {
  if (!('speechSynthesis' in window)) return;
  voices = speechSynthesis.getVoices() || [];
}
if ('speechSynthesis' in window) {
  loadVoices();
  speechSynthesis.addEventListener && speechSynthesis.addEventListener('voiceschanged', loadVoices);
}
function voicesFor(lang) {
  return voices.filter((v) => (v.lang || '').replace('_', '-').toLowerCase().startsWith(lang));
}
function pickVoice(lang) {
  const name = settings.voices[lang];
  if (name) { const v = voices.find((x) => x.name === name); if (v) return v; }
  const list = voicesFor(lang);
  if (!list.length) return null;
  const want = LANGS[lang].tts.toLowerCase();
  const score = (v) => {
    let sc = 0;
    if ((v.lang || '').replace('_', '-').toLowerCase() === want) sc += 2;
    if (/natural|online|neural|premium|enhanced/i.test(v.name)) sc += 3;
    if (/google/i.test(v.name)) sc += 1;
    return sc;
  };
  return list.slice().sort((a, b) => score(b) - score(a))[0];
}
function speak(text, lang) {
  if (!('speechSynthesis' in window) || !text) return;
  stopMic();
  speechSynthesis.cancel();
  const u = new SpeechSynthesisUtterance(text);
  u.lang = LANGS[lang].tts;
  const v = pickVoice(lang);
  if (v) u.voice = v;
  u.rate = Number(settings.rate) || 1;
  speechSynthesis.speak(u);
}
function listenBtn(text, lang, label = '🔊') {
  return el('button', { class: 'listen', type: 'button', title: '듣기', 'aria-label': '듣기', onclick: () => speak(text, lang) }, label);
}

/* ───────────── 음성 인식(STT) ───────────── */
const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
let rec = null;
let recGotResult = false;
function setMicStatus(t) { $('#micStatus').textContent = t; }
function stopMic() { if (rec) { try { rec.stop(); } catch { /* 무시 */ } } }
function toggleMic() {
  if (!SR) { toast('이 브라우저는 음성 인식을 지원하지 않습니다. Edge·삼성 인터넷·크롬을 사용하거나 입력창에 타자로 입력해 주세요.', true); return; }
  if (rec) { stopMic(); return; }
  if ('speechSynthesis' in window) speechSynthesis.cancel();
  const input = $('#input');
  const sep = session.lang === 'ja' ? '' : ' ';
  const base = input.value.trim() ? input.value.trim() + sep : '';
  recGotResult = false;
  rec = new SR();
  rec.lang = LANGS[session.lang].stt;
  rec.interimResults = true;
  rec.continuous = false;
  rec.maxAlternatives = 1;
  rec.onstart = () => { $('#micBtn').classList.add('on'); setMicStatus('듣고 있어요… 말을 마치면 잠시 기다리거나 다시 누르세요.'); };
  rec.onresult = (e) => {
    let finalText = '';
    let interim = '';
    for (let i = 0; i < e.results.length; i++) {
      const r = e.results[i];
      if (r.isFinal) finalText += r[0].transcript; else interim += r[0].transcript;
    }
    input.value = base + (finalText + interim).trim();
    if (finalText.trim()) recGotResult = true;
    autoResize();
  };
  rec.onerror = (e) => {
    const map = {
      'not-allowed': '마이크 권한이 거부되었습니다. 주소창의 자물쇠 아이콘에서 마이크를 허용해 주세요.',
      'service-not-allowed': '음성 인식 서비스를 사용할 수 없습니다. HTTPS 주소에서 열었는지 확인해 주세요.',
      'no-speech': '음성이 감지되지 않았습니다. 다시 눌러 말해 보세요.',
      'audio-capture': '마이크를 찾을 수 없습니다.',
      'network': '음성 인식 서버에 연결할 수 없습니다. 인터넷 연결을 확인해 주세요.',
      'language-not-supported': '이 브라우저에서 해당 언어의 음성 인식을 지원하지 않습니다.',
    };
    if (e.error !== 'aborted') toast(map[e.error] || `음성 인식 오류: ${e.error}`, e.error !== 'no-speech');
  };
  rec.onend = () => {
    rec = null;
    $('#micBtn').classList.remove('on');
    setMicStatus(recGotResult ? (settings.autoSend ? '보내는 중…' : '인식된 문장을 확인한 뒤 보내기를 누르세요.') : '마이크를 눌러 말해 보세요.');
    if (recGotResult && settings.autoSend && input.value.trim()) sendUser();
    else input.focus();
  };
  try { rec.start(); } catch (e) { rec = null; toast('음성 인식을 시작하지 못했습니다: ' + e.message, true); }
}

/* ───────────── 0. 키 화면 ───────────── */
async function testKey(key, out) {
  out.textContent = '확인 중…';
  try {
    const res = await fetch(OR_BASE + '/key', { headers: { Authorization: 'Bearer ' + key } });
    if (!res.ok) throw new ApiError(res.status, '');
    const j = await res.json();
    const d = j.data || {};
    const usage = typeof d.usage === 'number' ? `사용액 $${d.usage.toFixed(2)}` : '';
    const limit = d.limit != null ? ` / 한도 $${Number(d.limit).toFixed(2)}` : ' / 한도 없음';
    let modelNote = '';
    try { modelNote = await refreshCatalog(); } catch { modelNote = ' (모델 목록은 불러오지 못했습니다)'; }
    out.textContent = `✅ 연결되었습니다. ${usage}${limit}${modelNote}`;
    return true;
  } catch (e) {
    out.textContent = '❌ ' + (e instanceof ApiError && e.status === 401 ? 'API 키가 올바르지 않습니다.' : friendlyError(e));
    return false;
  }
}
/** OpenRouter 모델 목록으로 모델 ID를 확인하고 단가를 갱신합니다. */
async function refreshCatalog() {
  const res = await fetch(OR_BASE + '/models');
  if (!res.ok) throw new Error('models');
  const j = await res.json();
  modelCatalog = j.data || [];
  const used = new Set([...Object.values(settings.models), ...Object.values(MODELS)]);
  const missing = [];
  for (const id of used) {
    const m = modelCatalog.find((x) => x.id === id);
    if (!m) { missing.push(id); continue; }
    const pin = Number(m.pricing && m.pricing.prompt);
    const pout = Number(m.pricing && m.pricing.completion);
    if (isFinite(pin) && isFinite(pout)) settings.prices[id] = { in: +(pin * 1e6).toFixed(4), out: +(pout * 1e6).toFixed(4) };
  }
  saveSettings();
  fillModelList();
  const inUse = Object.values(settings.models).filter((id) => missing.includes(id));
  return inUse.length ? ` ⚠️ 모델 목록에 없는 ID: ${inUse.join(', ')} (설정에서 확인해 주세요)` : ' · 모델 ID와 단가 확인 완료';
}
function fillModelList() {
  const dl = $('#modelList');
  dl.textContent = '';
  const ids = modelCatalog ? modelCatalog.map((m) => m.id) : Object.values(MODELS);
  for (const id of ids) dl.append(el('option', { value: id }));
}
function initKeyView() {
  const input = $('#keyInput');
  $('#keyShow').onclick = () => {
    input.type = input.type === 'password' ? 'text' : 'password';
    $('#keyShow').textContent = input.type === 'password' ? '보기' : '숨기기';
  };
  for (const r of document.querySelectorAll('input[name=keyMode]')) r.checked = (r.value === 'local') === !!settings.keyPersist;
  $('#keyTest').onclick = () => {
    const k = input.value.trim();
    if (!k) { $('#keyResult').textContent = '키를 입력해 주세요.'; return; }
    testKey(k, $('#keyResult'));
  };
  $('#keySave').onclick = () => {
    const k = input.value.trim();
    if (!k) { $('#keyResult').textContent = '키를 입력해 주세요.'; input.focus(); return; }
    const persist = document.querySelector('input[name=keyMode]:checked').value === 'local';
    setKey(k, persist);
    input.value = '';
    toast(persist ? '키를 이 기기에 저장했습니다.' : '키를 이번 세션에만 보관합니다.');
    goSetup();
  };
  input.addEventListener('keydown', (e) => { if (e.key === 'Enter') $('#keySave').click(); });
}

/* ───────────── 1. 연습 설정 ───────────── */
function goSetup() {
  if (!getKey()) { showView('key'); return; }
  renderSetup();
  showView('setup');
}
function currentScenario() {
  const sc = SCENARIOS.find((x) => x.id === settings.scenario) || SCENARIOS[3];
  return sc;
}
function setupMissions() {
  const sc = currentScenario();
  if (sc.custom) {
    return settings.custom.missions.split('\n').map((t) => t.trim()).filter(Boolean);
  }
  return sc.missions;
}
function renderSetup() {
  // 이어서 하기 배너
  const hasOpen = session && session.messages && session.messages.length > 0;
  $('#resumeBanner').classList.toggle('hidden', !hasOpen);
  if (hasOpen) {
    $('#resumeBanner').firstElementChild.textContent = session.endedAt
      ? `저장하지 않은 대화 리포트가 있습니다. (${LANGS[session.lang].name} · ${session.scenario.title})`
      : `진행 중이던 대화가 있습니다. (${LANGS[session.lang].name} · ${session.scenario.title})`;
    $('#resumeBtn').textContent = session.endedAt ? '리포트 보기' : '이어서 하기';
  }
  $('#fileNotice').classList.toggle('hidden', location.protocol !== 'file:');

  // 언어
  const lc = $('#langChoices');
  lc.textContent = '';
  for (const [id, L] of Object.entries(LANGS)) {
    lc.append(el('button', {
      type: 'button', class: 'choice' + (settings.lang === id ? ' selected' : ''), 'aria-pressed': String(settings.lang === id),
      onclick: () => { settings.lang = id; saveSettings(); renderSetup(); },
    }, el('span', { class: 'ic' }, L.flag), el('span', { class: 't' }, `${L.name} · ${L.native}`),
    el('span', { class: 'd' }, levelLabel(id, settings.levels[id]))));
  }
  // 수준
  const ls = $('#levelSel');
  ls.textContent = '';
  for (const l of LEVELS) ls.append(el('option', { value: l.id }, levelLabel(settings.lang, l.id)));
  ls.value = settings.levels[settings.lang];
  ls.onchange = () => { settings.levels[settings.lang] = ls.value; saveSettings(); renderSetup(); };
  // 말투
  const ts = $('#toneSel');
  ts.textContent = '';
  for (const [k, v] of Object.entries(LANGS[settings.lang].tones)) ts.append(el('option', { value: k }, v));
  ts.value = settings.tones[settings.lang] || 'auto';
  ts.onchange = () => { settings.tones[settings.lang] = ts.value; saveSettings(); };
  // 길이
  $('#lengthSel').value = settings.length;
  $('#lengthSel').onchange = () => { settings.length = $('#lengthSel').value; saveSettings(); renderEstimate(); };

  // 상황
  const sc = $('#scenChoices');
  sc.textContent = '';
  for (const cat of ['여행', '일상', '학습']) {
    sc.append(el('div', { class: 'cat-label' }, cat));
    const grid = el('div', { class: 'choices scen' });
    for (const s of SCENARIOS.filter((x) => x.cat === cat)) {
      grid.append(el('button', {
        type: 'button', class: 'choice' + (settings.scenario === s.id ? ' selected' : ''), 'aria-pressed': String(settings.scenario === s.id),
        onclick: () => { settings.scenario = s.id; saveSettings(); renderSetup(); },
      }, el('span', { class: 'ic' }, s.icon), el('span', { class: 't' }, s.title), el('span', { class: 'd' }, s.desc)));
    }
    sc.append(grid);
  }
  renderScenarioExtra();
  renderMissionPreview();

  // 모델 구성
  const seg = $('#presetSeg');
  seg.textContent = '';
  for (const [k, p] of Object.entries(PRESETS)) {
    seg.append(el('button', { type: 'button', class: settings.preset === k ? 'on' : '', onclick: () => applyPreset(k) }, p.label));
  }
  if (settings.preset === 'custom') seg.append(el('button', { type: 'button', class: 'on' }, '사용자 지정'));
  $('#presetDesc').textContent = settings.preset === 'custom'
    ? `대화 ${settings.models.partner} · 교정 ${settings.models.coach} · 리포트 ${settings.models.report}`
    : PRESETS[settings.preset].desc;
  renderEstimate();
}
function applyPreset(k) {
  settings.preset = k;
  const p = PRESETS[k];
  settings.models = { partner: p.partner, coach: p.coach, report: p.report };
  saveSettings();
  renderSetup();
}
function renderScenarioExtra() {
  const box = $('#scenExtra');
  box.textContent = '';
  const sc = currentScenario();
  const bind = (input, path) => {
    input.addEventListener('input', () => {
      const [a, b] = path;
      if (b) settings[a][b] = input.value; else settings[a] = input.value;
      saveSettings();
      if (a === 'custom') renderMissionPreview();
    });
    return input;
  };
  if (sc.travel) {
    box.append(el('label', { class: 'field' }, el('span', null, '여행지 (선택) — 예: 마드리드, 오사카, 뉴욕'),
      bind(el('input', { type: 'text', value: settings.city, placeholder: '입력하면 현지 상황에 맞춰 대화합니다' }), ['city'])));
  }
  if (sc.topic) {
    box.append(el('label', { class: 'field' }, el('span', null, '주제'),
      bind(el('input', { type: 'text', value: settings.topic, placeholder: sc.id === 'debate' ? '예: 재택근무가 사무실 근무보다 낫다' : '예: 최근에 읽은 책, 좋아하는 사진 장소' }), ['topic'])));
  }
  if (sc.custom) {
    const c = settings.custom;
    box.append(
      el('label', { class: 'field' }, el('span', null, '상황 (한국어 가능)'), bind(el('input', { type: 'text', value: c.place, placeholder: '예: 박물관 매표소에서 학생 할인을 문의하는 상황' }), ['custom', 'place'])),
      el('div', { class: 'grid-2' },
        el('label', { class: 'field' }, el('span', null, 'AI 역할'), bind(el('input', { type: 'text', value: c.ai, placeholder: '예: 매표소 직원' }), ['custom', 'ai'])),
        el('label', { class: 'field' }, el('span', null, '내 역할'), bind(el('input', { type: 'text', value: c.user, placeholder: '예: 관광객' }), ['custom', 'user']))),
      el('label', { class: 'field' }, el('span', null, '미션 (한 줄에 하나, 비워 두면 AI가 만듭니다)'),
        bind(el('textarea', { rows: '3', placeholder: '학생 할인이 되는지 묻기\n오디오 가이드 대여하기' }, c.missions), ['custom', 'missions'])),
    );
  }
}
function renderMissionPreview() {
  const ol = $('#missionPreview');
  ol.textContent = '';
  const ms = setupMissions();
  if (!ms.length) ol.append(el('li', { class: 'muted' }, '대화를 시작하면 AI가 미션을 만들어 드립니다.'));
  for (const m of ms) ol.append(el('li', null, m));
}
/** 목표 턴 수 기준의 대략적인 예상 비용 */
function estimateCost(turns = settings.turns) {
  const P = (id) => priceOf(id) || { in: 1, out: 5 };
  const outPer = { short: 70, normal: 110, long: 180 }[settings.length] || 80;
  let partnerIn = 0;
  for (let i = 0; i <= turns; i++) partnerIn += 550 + i * (outPer + 40);
  const partnerOut = (turns + 1) * outPer;
  const coachIn = turns * 750;
  const coachOut = turns * 260;
  const repIn = 900 + turns * (outPer + 70);
  const repOut = 1300;
  const m = settings.models;
  const c = (id, i, o) => (i * P(id).in + o * P(id).out) / 1e6;
  return c(m.partner, partnerIn, partnerOut) + c(m.coach, coachIn, coachOut) + c(m.report, repIn, repOut);
}
function renderEstimate() {
  const v = estimateCost();
  $('#estimate').textContent = '';
  $('#estimate').append('예상 비용 (', String(settings.turns), '턴 기준): ', el('b', null, `약 ${fmtUSD(v)} (${fmtKRW(v)})`),
    el('br'), el('span', { class: 'small' }, '힌트·번역을 쓰면 조금 늘어납니다.'));
}

/* ───────────── 2. 회화 ───────────── */
async function startSession() {
  const sc = currentScenario();
  let missions = setupMissions();
  let scenario;
  if (sc.custom) {
    const c = settings.custom;
    if (!c.place.trim()) { toast('상황을 입력해 주세요.', true); return; }
    scenario = { id: 'custom', custom: true, icon: sc.icon, title: c.place.trim().slice(0, 30), en: c.place.trim(),
      ai: c.ai.trim() || 'a person who naturally fits the situation', user: c.user.trim() || 'the learner' };
  } else {
    scenario = { id: sc.id, icon: sc.icon, title: sc.title, en: sc.en, ai: sc.ai, user: sc.user };
  }
  if (sc.topic && !settings.topic.trim()) { toast('주제를 입력해 주세요.', true); return; }
  if (session && session.messages.length && !session.saved) {
    if (!confirm('저장하지 않은 이전 대화가 있습니다. 새 대화를 시작하면 이전 대화는 사라집니다. 계속할까요?')) return;
  }
  session = {
    id: uid(), lang: settings.lang, level: settings.levels[settings.lang], tone: settings.tones[settings.lang] || 'auto',
    length: settings.length, scenario, city: sc.travel ? settings.city.trim() : '', topic: sc.topic ? settings.topic.trim() : '',
    models: { ...settings.models }, missions: missions.map((t) => ({ text: t, done: false })),
    messages: [], costs: [], startedAt: Date.now(), endedAt: null, report: null, saved: false, turnNoticeShown: false,
  };
  saveSession();
  $('#startBtn').disabled = true;
  try {
    if (!session.missions.length) {
      $('#startBtn').textContent = '미션 만드는 중…';
      await generateMissions();
    }
  } finally {
    $('#startBtn').disabled = false;
    $('#startBtn').textContent = '대화 시작하기 ▶';
  }
  openChat();
  partnerTurn();
}
async function generateMissions() {
  const L = LANGS[session.lang];
  try {
    const text = await chat({
      model: session.models.coach, role: '미션', maxTokens: 600, temperature: 0.5,
      messages: [{ role: 'user', content: `Create 3 short, concrete speaking goals in Korean for a ${L.en} learner (CEFR ${session.level}) in this role-play: "${session.scenario.en}" (learner: ${session.scenario.user}, partner: ${session.scenario.ai}). Return ONLY JSON: {"missions":["...","...","..."]}` }],
    });
    const j = parseJSON(text);
    if (j && Array.isArray(j.missions) && j.missions.length) session.missions = j.missions.slice(0, 4).map((t) => ({ text: String(t), done: false }));
  } catch (e) {
    toast(friendlyError(e), true);
  }
  if (!session.missions.length) session.missions = [{ text: '상황에 맞게 자연스럽게 대화하기', done: false }];
  saveSession();
}
function openChat() {
  const L = LANGS[session.lang];
  $('#chatTitle').textContent = `${session.scenario.icon || ''} ${session.scenario.title}${session.city ? ' · ' + session.city : ''}`;
  renderChatSub();
  $('#hints').classList.add('hidden');
  $('#turnNotice').classList.add('hidden');
  $('#input').placeholder = `${L.name}로 말하거나 입력하세요 (Enter: 보내기, Shift+Enter: 줄바꿈)`;
  $('#micBtn').disabled = !SR;
  setMicStatus(SR ? '마이크를 눌러 말해 보세요.' : '이 브라우저는 음성 인식을 지원하지 않아 타자로 입력합니다.');
  renderMessages();
  renderMissions();
  renderLiveStats();
  renderLiveCost();
  showView('chat');
}
function userTurns() { return session.messages.filter((m) => m.role === 'user').length; }
function renderChatSub() {
  const L = LANGS[session.lang];
  const done = session.missions.filter((m) => m.done).length;
  $('#chatSub').textContent = `${L.flag} ${L.name} · ${levelLabel(session.lang, session.level)} · 턴 ${userTurns()}/${settings.turns} · 미션 ${done}/${session.missions.length}`;
}

function renderMessages() {
  const box = $('#messages');
  box.textContent = '';
  for (const m of session.messages) box.append(messageNode(m));
  scrollChat();
}
function updateMessage(m) {
  const old = document.querySelector(`[data-mid="${m.id}"]`);
  const node = messageNode(m);
  if (old) old.replaceWith(node); else $('#messages').append(node);
}
function scrollChat() {
  const box = $('#messages');
  box.scrollTop = box.scrollHeight;
  const last = box.lastElementChild;
  if (last && window.innerWidth <= 860) last.scrollIntoView({ block: 'end', behavior: 'smooth' });
}
function messageNode(m) {
  const lang = session.lang;
  const wrap = el('div', { class: 'msg ' + m.role, 'data-mid': m.id });
  if (m.role === 'ai') {
    const bubble = el('div', { class: 'bubble' + (m.pending ? ' typing' : ''), lang: LANGS[lang].tts }, m.text || '');
    wrap.append(bubble);
    if (m.error) {
      wrap.append(el('div', { class: 'corr fix' }, `⚠️ ${m.error} `, el('button', { class: 'btn sm', type: 'button', onclick: () => retryPartner(m) }, '다시 시도')));
    } else if (!m.pending) {
      const tools = el('div', { class: 'tools' },
        el('button', { class: 'link-btn', type: 'button', onclick: () => speak(m.text, lang) }, '🔊 듣기'),
        el('button', { class: 'link-btn', type: 'button', onclick: () => toggleTranslation(m) }, m.showTr ? '번역 숨기기' : '🇰🇷 번역'));
      wrap.append(tools);
      if (m.showTr) wrap.append(el('div', { class: 'translation' }, m.translation || '번역 중…'));
    }
  } else {
    wrap.append(el('div', { class: 'bubble', lang: LANGS[lang].tts }, m.text));
    const c = correctionNode(m);
    if (c) wrap.append(c);
  }
  return wrap;
}
function highlightOriginal(text, errors) {
  const ranges = [];
  let cursor = 0;
  for (const er of errors || []) {
    if (!er.wrong) continue;
    let i = text.indexOf(er.wrong, cursor);
    if (i < 0) i = text.indexOf(er.wrong);
    if (i < 0) continue;
    ranges.push([i, i + er.wrong.length]);
    cursor = i + er.wrong.length;
  }
  ranges.sort((a, b) => a[0] - b[0]);
  const out = [];
  let pos = 0;
  for (const [s, e] of ranges) {
    if (s < pos) continue;
    out.push(text.slice(pos, s), el('mark', { class: 'wrong' }, text.slice(s, e)));
    pos = e;
  }
  out.push(text.slice(pos));
  return out;
}
function correctionNode(m) {
  const lang = session.lang;
  if (m.correcting) return el('div', { class: 'corr loading' }, el('span', { class: 'spinner' }), ' 교정 중…');
  if (m.correctionError) {
    return el('div', { class: 'corr' }, '교정을 불러오지 못했습니다. ', el('button', { class: 'btn sm', type: 'button', onclick: () => correct(m) }, '다시 시도'));
  }
  const c = m.correction;
  if (!c) return null;
  const errs = Array.isArray(c.errors) ? c.errors : [];
  const card = el('div', { class: 'corr ' + (errs.length ? 'fix' : 'ok') });
  if (!errs.length) {
    card.append(el('div', { class: 'head' }, '✅ 좋아요! 틀린 곳이 없습니다.'));
    if (c.praise_ko) card.append(el('div', null, c.praise_ko));
  } else {
    card.append(el('div', { class: 'head' }, `✏️ 고칠 부분 ${errs.length}개`));
    card.append(el('div', { class: 'orig', lang: LANGS[lang].tts }, ...highlightOriginal(m.text, errs)));
    card.append(el('div', { class: 'say' }, '→ ', el('span', { class: 'fixed', lang: LANGS[lang].tts }, c.corrected || ''), listenBtn(c.corrected, lang)));
    const ul = el('ul');
    for (const er of errs) {
      ul.append(el('li', null, el('span', { class: 'chip' }, er.type || '기타'),
        el('b', null, er.wrong || ''), ' → ', el('b', null, er.fix || ''), er.explain_ko ? ` : ${er.explain_ko}` : ''));
    }
    card.append(ul);
  }
  const nat = (Array.isArray(c.natural) ? c.natural : []).filter(Boolean);
  if (nat.length) {
    card.append(el('div', { class: 'sec' }, '💬 더 자연스러운 표현'));
    const ul = el('ul');
    for (const n of nat) ul.append(el('li', { class: 'say' }, el('span', { lang: LANGS[lang].tts }, n), listenBtn(n, lang)));
    card.append(ul);
  }
  if (c.note_ko) card.append(el('div', { class: 'small', style: 'margin-top:6px' }, '💡 ' + c.note_ko));
  return card;
}

async function partnerTurn(existing) {
  const m = existing || { id: uid(), role: 'ai', text: '', pending: true };
  m.text = ''; m.pending = true; m.error = null;
  if (!existing) session.messages.push(m);
  partnerBusy = true;
  updateSendState();
  updateMessage(m);
  scrollChat();
  let lastPaint = 0;
  try {
    await chat({
      model: session.models.partner, role: '대화', stream: true, maxTokens: 900, temperature: 0.8,
      messages: partnerMessages(session),
      onDelta: (d) => {
        m.text += d;
        const now = performance.now();
        if (now - lastPaint > 60) {
          lastPaint = now;
          const b = document.querySelector(`[data-mid="${m.id}"] .bubble`);
          if (b) b.textContent = m.text;
          scrollChat();
        }
      },
    });
    m.text = m.text.trim();
    if (!m.text) throw new Error('AI 응답이 비어 있습니다.');
    m.pending = false;
    updateMessage(m);
    scrollChat();
    if (settings.autoSpeak) speak(m.text, session.lang);
  } catch (e) {
    m.pending = false;
    m.error = friendlyError(e);
    updateMessage(m);
  } finally {
    partnerBusy = false;
    updateSendState();
    saveSession();
  }
}
function retryPartner(m) { partnerTurn(m); }

function updateSendState() {
  $('#sendBtn').disabled = partnerBusy;
  $('#sendBtn').textContent = partnerBusy ? 'AI 응답 중…' : '보내기';
}
function sendUser() {
  if (!session || partnerBusy) return;
  const input = $('#input');
  const text = input.value.trim();
  if (!text) { input.focus(); return; }
  stopMic();
  input.value = '';
  autoResize();
  $('#hints').classList.add('hidden');
  setMicStatus(SR ? '마이크를 눌러 말해 보세요.' : '');
  const m = { id: uid(), role: 'user', text, correction: null, correcting: true };
  session.messages.push(m);
  session.saved = false;
  session.report = null; // 리포트 이후 대화를 이어 가면 다시 생성
  updateMessage(m);
  scrollChat();
  renderChatSub();
  saveSession();
  correct(m);
  partnerTurn();
  if (userTurns() >= settings.turns && !session.turnNoticeShown) {
    session.turnNoticeShown = true;
    $('#turnNotice').classList.remove('hidden');
  }
}
async function correct(m) {
  m.correcting = true; m.correctionError = false;
  updateMessage(m);
  try {
    const text = await chat({
      model: session.models.coach, role: '교정', maxTokens: 1500, temperature: 0.2,
      messages: [{ role: 'user', content: coachPrompt(session, m) }],
    });
    const j = parseJSON(text);
    if (!j) throw new Error('교정 결과를 해석하지 못했습니다.');
    if (!Array.isArray(j.errors)) j.errors = [];
    j.errors = j.errors.filter((er) => er && (er.wrong || er.fix)).map((er) => ({ ...er, type: ERROR_TYPES.includes(er.type) ? er.type : '기타' }));
    m.correction = j;
    if (Array.isArray(j.missions_done)) {
      for (const i of j.missions_done) {
        const idx = Number(i);
        if (session.missions[idx]) session.missions[idx].done = true;
      }
      renderMissions();
    }
  } catch (e) {
    console.warn(e);
    m.correctionError = true;
  } finally {
    m.correcting = false;
    updateMessage(m);
    renderLiveStats();
    saveSession();
  }
}
async function toggleTranslation(m) {
  m.showTr = !m.showTr;
  updateMessage(m);
  if (!m.showTr || m.translation) return;
  try {
    const t = await chat({
      model: session.models.coach, role: '번역', maxTokens: 500, temperature: 0.2,
      messages: [{ role: 'user', content: `Translate the following ${LANGS[session.lang].en} text into natural Korean. Output only the translation.\n\n${m.text}` }],
    });
    m.translation = t.trim();
  } catch (e) {
    m.translation = '번역 실패: ' + friendlyError(e);
  }
  updateMessage(m);
  saveSession();
}
async function showHints() {
  const box = $('#hints');
  box.classList.remove('hidden');
  box.textContent = '';
  box.append(el('div', { class: 'h-head' }, el('span', null, '💡 이렇게 말해 보세요'),
    el('button', { class: 'link-btn', type: 'button', onclick: () => box.classList.add('hidden') }, '닫기')),
  el('div', { class: 'muted small' }, el('span', { class: 'spinner' }), ' 힌트를 만드는 중…'));
  try {
    const text = await chat({
      model: session.models.coach, role: '힌트', maxTokens: 800, temperature: 0.7,
      messages: [{ role: 'user', content: hintPrompt(session) }],
    });
    const j = parseJSON(text);
    const hints = j && Array.isArray(j.hints) ? j.hints : [];
    if (!hints.length) throw new Error('힌트를 만들지 못했습니다.');
    box.lastElementChild.remove();
    for (const h of hints.slice(0, 3)) {
      box.append(el('div', { class: 'hint-item' },
        el('div', { class: 'ht say', lang: LANGS[session.lang].tts }, h.text, listenBtn(h.text, session.lang)),
        h.reading ? el('div', { class: 'hr' }, h.reading) : null,
        el('div', { class: 'hr' }, h.ko || ''),
        el('button', { class: 'link-btn', type: 'button', onclick: () => { $('#input').value = h.text; autoResize(); $('#input').focus(); } }, '입력창에 넣기')));
    }
  } catch (e) {
    box.lastElementChild.textContent = '힌트를 불러오지 못했습니다: ' + friendlyError(e);
  }
}

function renderMissions() {
  const ul = $('#missionList');
  ul.textContent = '';
  session.missions.forEach((ms, i) => {
    const cb = el('input', { type: 'checkbox', id: 'ms' + i });
    cb.checked = ms.done;
    cb.onchange = () => { ms.done = cb.checked; saveSession(); renderMissions(); };
    ul.append(el('li', { class: ms.done ? 'done' : '' }, cb, el('label', { for: 'ms' + i }, el('span', null, ms.text))));
  });
  const done = session.missions.filter((m) => m.done).length;
  $('#missionCard h2').textContent = `🎯 미션 ${done}/${session.missions.length}`;
  renderChatSub();
}
function errorStats(s = session) {
  const counts = {};
  for (const m of s.messages) {
    if (m.role !== 'user' || !m.correction) continue;
    for (const er of m.correction.errors || []) counts[er.type || '기타'] = (counts[er.type || '기타'] || 0) + 1;
  }
  return Object.entries(counts).sort((a, b) => b[1] - a[1]);
}
function barsNode(stats) {
  const box = el('div', { class: 'bars' });
  if (!stats.length) { box.append(el('div', { class: 'small muted' }, '아직 기록된 오류가 없습니다.')); return box; }
  const max = Math.max(...stats.map((x) => x[1]));
  for (const [type, n] of stats) {
    box.append(el('div', { class: 'bar-row' }, el('span', null, type),
      el('div', { class: 'bar-track' }, el('div', { class: 'bar-fill', style: `width:${Math.round((n / max) * 100)}%` })),
      el('span', { class: 'n' }, n)));
  }
  return box;
}
function renderLiveStats() {
  const box = $('#liveStats');
  box.replaceWith(Object.assign(barsNode(errorStats()), { id: 'liveStats' }));
}
function renderLiveCost() {
  const box = $('#liveCost');
  if (!box || !session) return;
  box.textContent = '';
  const t = totalCost();
  box.append(el('div', null, '누적: ', el('b', null, `${fmtUSD(t)} (${fmtKRW(t)})`)),
    el('div', { class: 'muted' }, `AI 호출 ${session.costs.length}회`),
    el('div', { class: 'muted' }, `대화: ${shortModel(session.models.partner)} · 교정: ${shortModel(session.models.coach)}`));
}
function shortModel(id) { return (id || '').split('/').pop(); }

/* ───────────── 3. 리포트 ───────────── */
async function endSession() {
  if (!session) return;
  if (userTurns() === 0) { toast('한 번 이상 말한 뒤에 리포트를 받을 수 있습니다.', true); return; }
  stopMic();
  if ('speechSynthesis' in window) speechSynthesis.cancel();
  if (!session.endedAt) session.endedAt = Date.now();
  saveSession();
  showView('report');
  renderReport();
  if (!session.report) await generateReport();
}
async function generateReport() {
  session.reportError = null;
  session.reportLoading = true;
  renderReport();
  try {
    const text = await chat({
      model: session.models.report, role: '리포트', maxTokens: 4000, temperature: 0.4,
      messages: [{ role: 'user', content: reportPrompt(session) }],
    });
    const j = parseJSON(text);
    if (!j) throw new Error('리포트 결과를 해석하지 못했습니다.');
    session.report = j;
    if (Array.isArray(j.missions)) {
      j.missions.forEach((rm, i) => { if (session.missions[i] && rm && rm.done === true) session.missions[i].done = true; });
    }
  } catch (e) {
    session.reportError = friendlyError(e);
  } finally {
    session.reportLoading = false;
    saveSession();
    renderReport();
  }
}
function sessionStats(s = session) {
  const users = s.messages.filter((m) => m.role === 'user');
  const units = s.lang === 'ja'
    ? users.reduce((a, m) => a + m.text.replace(/\s/g, '').length, 0)
    : users.reduce((a, m) => a + m.text.split(/\s+/).filter(Boolean).length, 0);
  const corrected = users.filter((m) => m.correction);
  const perfect = corrected.filter((m) => !(m.correction.errors || []).length).length;
  const mins = Math.max(1, Math.round(((s.endedAt || Date.now()) - s.startedAt) / 60000));
  return { turns: users.length, units, unitLabel: s.lang === 'ja' ? '글자 수' : '단어 수', perfect, corrected: corrected.length, mins };
}
function corrections(s = session) {
  return s.messages.filter((m) => m.role === 'user' && m.correction && (m.correction.errors || []).length);
}
function renderReport() {
  const s = session;
  const L = LANGS[s.lang];
  const r = s.report || {};
  const st = sessionStats(s);
  const box = $('#reportBody');
  box.textContent = '';
  box.append(
    el('h1', null, '토크메이트 회화 리포트'),
    el('div', { class: 'meta' }, `${fmtDate(s.startedAt)} · ${L.name} · ${levelLabel(s.lang, s.level)} · ${s.scenario.title}${s.city ? ' (' + s.city + ')' : ''}${s.topic ? ' · 주제: ' + s.topic : ''}`),
    el('div', { class: 'stats' },
      stat(st.turns, '내가 말한 횟수'), stat(st.units, st.unitLabel),
      stat(`${st.perfect}/${st.corrected}`, '틀린 곳 없는 문장'), stat(`${st.mins}분`, '연습 시간')),
  );
  if (s.reportLoading) box.append(el('p', null, el('span', { class: 'spinner' }), ' AI가 리포트를 작성하고 있습니다… (', shortModel(s.models.report), ')'));
  if (s.reportError) {
    box.append(el('div', { class: 'banner no-print' }, el('span', null, `리포트 생성 실패: ${s.reportError}`),
      el('button', { class: 'btn sm', type: 'button', onclick: generateReport }, '다시 시도')));
  }
  if (r.summary_ko) {
    box.append(el('section', null, el('h2', null, '총평'), el('p', null, r.summary_ko), r.level_comment_ko ? el('p', { class: 'muted' }, r.level_comment_ko) : null));
  }
  // 미션
  const mRows = s.missions.map((m, i) => {
    const rm = Array.isArray(r.missions) ? r.missions[i] : null;
    return [m.done ? '✅' : '⬜', m.text, (rm && rm.comment_ko) || ''];
  });
  box.append(el('section', null, el('h2', null, '미션'), table(['달성', '미션', '코멘트'], mRows)));
  // 잘한 점 / 연습할 점
  if ((r.strengths_ko || []).length || (r.improve_ko || []).length) {
    box.append(el('section', null, el('div', { class: 'two-col' },
      el('div', null, el('h2', null, '👍 잘한 점'), el('ul', null, ...(r.strengths_ko || []).map((t) => el('li', null, t)))),
      el('div', null, el('h2', null, '🎯 다음에 연습할 점'), el('ul', null, ...(r.improve_ko || []).map((t) => el('li', null, t)))))));
  }
  // 오류 통계
  box.append(el('section', null, el('h2', null, '오류 유형 통계'), barsNode(errorStats(s))));
  // 표현
  if ((r.expressions || []).length) {
    box.append(el('section', null, el('h2', null, '📌 꼭 외울 표현'),
      table(['표현', '뜻', '예문'], r.expressions.map((x) => [x.text, x.meaning_ko, x.example]), true)));
  }
  // 교정 노트
  const cs = corrections(s);
  if (cs.length) {
    box.append(el('section', { class: 'flow' }, el('h2', null, '✏️ 교정 노트'),
      table(['내 문장', '고친 문장', '설명'], cs.map((m) => [m.text, m.correction.corrected,
        (m.correction.errors || []).map((er) => `[${er.type}] ${er.wrong} → ${er.fix}: ${er.explain_ko || ''}`).join('\n')]), true)));
  }
  // 다음 추천
  if (r.next && (r.next.scenario_ko || r.next.tip_ko)) {
    box.append(el('section', null, el('h2', null, '다음 연습 추천'), el('p', null, el('b', null, r.next.scenario_ko || ''), r.next.tip_ko ? ' — ' + r.next.tip_ko : '')));
  }
  // 대화 전문
  const tr = el('div', { class: 'transcript' });
  for (const m of s.messages) {
    if (!m.text || m.error) continue;
    const line = el('div', { class: 'line' }, el('span', { class: 'who' }, m.role === 'ai' ? '🤖 AI' : '🙋 나'), el('span', { lang: L.tts }, m.text));
    if (m.role === 'user' && m.correction && (m.correction.errors || []).length) line.append(el('div', { class: 'fixline' }, '→ ' + m.correction.corrected));
    tr.append(line);
  }
  box.append(el('section', { class: 'flow' }, el('h2', null, '대화 전문'), tr));
  // 비용
  box.append(el('section', null, el('h2', null, 'AI 사용 비용'), table(['역할', '모델', '호출', '입력 토큰', '출력 토큰', '비용'], costRows(s).map((c) => [c.role, c.model, c.calls, c.in, c.out, `${fmtUSD(c.cost)} (${fmtKRW(c.cost)})`])),
    el('p', { class: 'small muted' }, `합계 ${fmtUSD(totalCost())} (${fmtKRW(totalCost())}), 환율 1달러 = ${settings.fx}원 기준. AI 교정과 리포트는 참고용이며 틀릴 수 있습니다.`)));
  updateCostBadge();
}
function stat(v, l) { return el('div', { class: 'stat' }, el('div', { class: 'v' }, String(v)), el('div', { class: 'l' }, l)); }
function table(head, rows, pre) {
  return el('div', { class: 'tbl-wrap' }, el('table', { class: 'tbl' },
    el('thead', null, el('tr', null, ...head.map((h) => el('th', null, h)))),
    el('tbody', null, ...rows.map((r) => el('tr', null, ...r.map((c) => el('td', pre ? { style: 'white-space:pre-wrap' } : null, c == null ? '' : String(c))))))));
}
function costRows(s = session) {
  const map = new Map();
  for (const c of s.costs) {
    const k = c.role + '|' + c.model;
    const v = map.get(k) || { role: c.role, model: c.model, calls: 0, in: 0, out: 0, cost: 0 };
    v.calls++; v.in += c.in || 0; v.out += c.out || 0; v.cost += c.cost || 0;
    map.set(k, v);
  }
  return [...map.values()];
}

/* ───────────── 내보내기 ───────────── */
function fileBase() {
  const s = session;
  const d = new Date(s.startedAt);
  const stamp = `${fmtDate(s.startedAt, false)}_${pad(d.getHours())}${pad(d.getMinutes())}`;
  // 브라우저에 따라 한글 파일 이름이 깨질 수 있어 영문으로 만듭니다. 예: talkmate_es_restaurant_2026-10-16_1530
  return `talkmate_${s.lang}_${s.scenario.id}_${stamp}`;
}
function download(blob, name) {
  const a = el('a', { href: URL.createObjectURL(blob), download: name });
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}
function loadScript(src, globalName) {
  if (window[globalName]) return Promise.resolve(window[globalName]);
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src;
    s.crossOrigin = 'anonymous';
    s.onload = () => (window[globalName] ? resolve(window[globalName]) : reject(new Error(globalName + ' 로드 실패')));
    s.onerror = () => reject(new Error('라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인해 주세요.'));
    document.head.append(s);
  });
}
function markSaved() { session.saved = true; saveSession(); }

function exportPdf() {
  const prevTitle = document.title;
  document.title = fileBase();
  const restore = () => { document.title = prevTitle; window.removeEventListener('afterprint', restore); };
  window.addEventListener('afterprint', restore);
  markSaved();
  toast('인쇄 창에서 대상(프린터)을 "PDF로 저장"으로 선택해 주세요.');
  setTimeout(() => window.print(), 300);
}

async function exportXlsx() {
  let XLSX;
  try { XLSX = await loadScript(CDN.xlsx, 'XLSX'); } catch (e) { toast(e.message, true); return; }
  const s = session;
  const L = LANGS[s.lang];
  const r = s.report || {};
  const st = sessionStats(s);
  const wb = XLSX.utils.book_new();
  const add = (name, aoa, widths) => {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    if (widths) ws['!cols'] = widths.map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, name);
  };
  add('요약', [
    ['항목', '내용'],
    ['날짜', fmtDate(s.startedAt)], ['언어', L.name], ['수준', levelLabel(s.lang, s.level)],
    ['상황', s.scenario.title], ['여행지', s.city || ''], ['주제', s.topic || ''],
    ['내가 말한 횟수', st.turns], [st.unitLabel, st.units], ['틀린 곳 없는 문장', `${st.perfect}/${st.corrected}`], ['연습 시간(분)', st.mins],
    ['총평', r.summary_ko || ''], ['수준 코멘트', r.level_comment_ko || ''],
    ['잘한 점', (r.strengths_ko || []).join('\n')], ['연습할 점', (r.improve_ko || []).join('\n')],
    ['다음 추천', r.next ? `${r.next.scenario_ko || ''} ${r.next.tip_ko || ''}`.trim() : ''],
    ['총비용(USD)', +totalCost().toFixed(5)],
  ], [18, 90]);
  add('대화 기록', [['번호', '화자', '문장', '한국어 번역']].concat(
    s.messages.filter((m) => m.text && !m.error).map((m, i) => [i + 1, m.role === 'ai' ? 'AI' : '나', m.text, m.translation || ''])), [6, 6, 70, 50]);
  const corrRows = [['번호', '내 문장', '고친 문장', '오류 유형', '틀린 부분', '고침', '설명', '더 자연스러운 표현', '팁']];
  let n = 0;
  for (const m of s.messages.filter((x) => x.role === 'user' && x.correction)) {
    n++;
    const c = m.correction;
    const errs = c.errors && c.errors.length ? c.errors : [null];
    errs.forEach((er, i) => {
      corrRows.push([n, i === 0 ? m.text : '', i === 0 ? c.corrected : '', er ? er.type : '(오류 없음)', er ? er.wrong : '', er ? er.fix : '',
        er ? er.explain_ko || '' : '', i === 0 ? (c.natural || []).join('\n') : '', i === 0 ? c.note_ko || '' : '']);
    });
  }
  add('교정 노트', corrRows, [6, 45, 45, 12, 20, 20, 45, 45, 35]);
  add('표현 목록', [['표현', '뜻', '예문', '언어', '날짜']].concat((r.expressions || []).map((x) => [x.text, x.meaning_ko, x.example, L.name, fmtDate(s.startedAt, false)])), [40, 30, 50, 8, 12]);
  add('오류 통계', [['오류 유형', '횟수']].concat(errorStats(s)), [16, 8]);
  add('미션', [['미션', '달성', '코멘트']].concat(s.missions.map((m, i) => [m.text, m.done ? 'O' : 'X', (r.missions && r.missions[i] && r.missions[i].comment_ko) || ''])), [45, 6, 50]);
  add('비용', [['역할', '모델', '호출 수', '입력 토큰', '출력 토큰', '비용(USD)', '비용(원)']].concat(
    costRows(s).map((c) => [c.role, c.model, c.calls, c.in, c.out, +c.cost.toFixed(5), Math.round(c.cost * settings.fx)])), [10, 32, 8, 10, 10, 10, 10]);
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  download(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), fileBase() + '.xlsx');
  markSaved();
  toast('Excel 파일을 저장했습니다.');
}

async function exportDocx() {
  let D;
  try { D = await loadScript(CDN.docx, 'docx'); } catch (e) { toast(e.message, true); return; }
  const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType } = D;
  const s = session;
  const L = LANGS[s.lang];
  const r = s.report || {};
  const st = sessionStats(s);
  const p = (text, opts = {}) => new Paragraph({ children: [new TextRun({ text: String(text ?? ''), ...opts })], spacing: { after: 80 } });
  const h = (text, level = HeadingLevel.HEADING_2) => new Paragraph({ text, heading: level, spacing: { before: 240, after: 120 } });
  const bullet = (text) => new Paragraph({ text: String(text ?? ''), bullet: { level: 0 } });
  const cell = (text, head) => new TableCell({
    children: String(text ?? '').split('\n').map((line) => new Paragraph({ children: [new TextRun({ text: line, bold: !!head })] })),
    shading: head ? { type: ShadingType.CLEAR, color: 'auto', fill: 'E7ECFF' } : undefined,
  });
  const tbl = (head, rows) => new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [new TableRow({ tableHeader: true, children: head.map((x) => cell(x, true)) }),
      ...rows.map((row) => new TableRow({ children: row.map((x) => cell(x)) }))],
  });
  const children = [
    new Paragraph({ text: '토크메이트 회화 리포트', heading: HeadingLevel.TITLE }),
    p(`${fmtDate(s.startedAt)} · ${L.name} · ${levelLabel(s.lang, s.level)} · ${s.scenario.title}${s.city ? ' (' + s.city + ')' : ''}${s.topic ? ' · 주제: ' + s.topic : ''}`, { color: '555555' }),
    p(`내가 말한 횟수 ${st.turns}회 · ${st.unitLabel} ${st.units} · 틀린 곳 없는 문장 ${st.perfect}/${st.corrected} · 연습 시간 ${st.mins}분`),
  ];
  if (r.summary_ko) children.push(h('총평'), p(r.summary_ko), ...(r.level_comment_ko ? [p(r.level_comment_ko, { italics: true })] : []));
  children.push(h('미션'), tbl(['달성', '미션', '코멘트'], s.missions.map((m, i) => [m.done ? '✔' : '✘', m.text, (r.missions && r.missions[i] && r.missions[i].comment_ko) || ''])));
  if ((r.strengths_ko || []).length) children.push(h('잘한 점'), ...r.strengths_ko.map(bullet));
  if ((r.improve_ko || []).length) children.push(h('다음에 연습할 점'), ...r.improve_ko.map(bullet));
  const es = errorStats(s);
  children.push(h('오류 유형 통계'), es.length ? tbl(['오류 유형', '횟수'], es) : p('기록된 오류가 없습니다.'));
  if ((r.expressions || []).length) children.push(h('꼭 외울 표현'), tbl(['표현', '뜻', '예문'], r.expressions.map((x) => [x.text, x.meaning_ko, x.example])));
  const cs = corrections(s);
  if (cs.length) {
    children.push(h('교정 노트'));
    for (const m of cs) {
      const c = m.correction;
      children.push(new Paragraph({ children: [new TextRun({ text: '나: ', bold: true }), new TextRun({ text: m.text, strike: true, color: 'C92A2A' })] }));
      children.push(new Paragraph({ children: [new TextRun({ text: '→ ', bold: true }), new TextRun({ text: c.corrected || '', bold: true, color: '2B8A3E' })] }));
      for (const er of c.errors) children.push(bullet(`[${er.type}] ${er.wrong} → ${er.fix}: ${er.explain_ko || ''}`));
      if ((c.natural || []).length) children.push(p('더 자연스러운 표현: ' + c.natural.join(' / '), { color: '555555' }));
    }
  }
  if (r.next && (r.next.scenario_ko || r.next.tip_ko)) children.push(h('다음 연습 추천'), p(`${r.next.scenario_ko || ''}${r.next.tip_ko ? ' — ' + r.next.tip_ko : ''}`));
  children.push(h('대화 전문'));
  for (const m of s.messages.filter((x) => x.text && !x.error)) {
    children.push(new Paragraph({ children: [new TextRun({ text: m.role === 'ai' ? 'AI: ' : '나: ', bold: true }), new TextRun(m.text)], spacing: { after: 60 } }));
    if (m.role === 'user' && m.correction && (m.correction.errors || []).length) children.push(p('   → ' + m.correction.corrected, { color: '2B8A3E' }));
  }
  children.push(h('AI 사용 비용'), tbl(['역할', '모델', '호출', '입력 토큰', '출력 토큰', '비용'], costRows(s).map((c) => [c.role, c.model, c.calls, c.in, c.out, `${fmtUSD(c.cost)} (${fmtKRW(c.cost)})`])));
  children.push(p('AI 교정과 리포트는 참고용이며 틀릴 수 있습니다.', { color: '777777', size: 18 }));
  const doc = new Document({
    styles: { default: { document: { run: { font: 'Malgun Gothic', size: 21 } } } },
    sections: [{ children }],
  });
  const blob = await Packer.toBlob(doc);
  download(blob, fileBase() + '.docx');
  markSaved();
  toast('Word 파일을 저장했습니다.');
}

/* ───────────── 설정 대화 상자 ───────────── */
function openSettings() {
  const dlg = $('#settingsDlg');
  const key = getKey();
  const where = safeGet(localStorage, STORE.key) ? '이 기기(localStorage)' : (safeGet(sessionStorage, STORE.key) ? '이번 세션(sessionStorage)' : '');
  $('#keyStatus').textContent = key ? `키 입력됨: ${key.slice(0, 8)}…${key.slice(-4)} · 보관 위치: ${where}` : '입력된 키가 없습니다.';
  $('#setKeyPersist').checked = !!safeGet(localStorage, STORE.key) || (!key && settings.keyPersist);
  $('#setKeyResult').textContent = '';
  $('#setPartner').value = settings.models.partner;
  $('#setCoach').value = settings.models.coach;
  $('#setReport').value = settings.models.report;
  renderPriceRows();
  $('#setFx').value = settings.fx;
  $('#setAutoSpeak').checked = settings.autoSpeak;
  $('#setAutoSend').checked = settings.autoSend;
  $('#setRate').value = settings.rate;
  $('#rateVal').textContent = `${Number(settings.rate).toFixed(2)}배`;
  $('#setTurns').value = settings.turns;
  renderVoiceRows();
  fillModelList();
  if (dlg.showModal) dlg.showModal(); else dlg.setAttribute('open', '');
}
function renderPriceRows() {
  const box = $('#priceRows');
  box.textContent = '';
  const ids = [...new Set([$('#setPartner').value.trim(), $('#setCoach').value.trim(), $('#setReport').value.trim()].filter(Boolean))];
  for (const id of ids) {
    const pr = settings.prices[id] || { in: '', out: '' };
    box.append(el('div', { class: 'price-row', 'data-model': id }, el('code', null, id),
      el('input', { type: 'number', step: '0.01', min: '0', value: pr.in, 'aria-label': id + ' 입력 단가' }),
      el('input', { type: 'number', step: '0.01', min: '0', value: pr.out, 'aria-label': id + ' 출력 단가' })));
  }
}
function renderVoiceRows() {
  loadVoices();
  const box = $('#voiceRows');
  box.textContent = '';
  for (const [id, L] of Object.entries(LANGS)) {
    const sel = el('select', { 'data-lang': id }, el('option', { value: '' }, '자동 선택'));
    for (const v of voicesFor(id)) sel.append(el('option', { value: v.name }, `${v.name} (${v.lang})`));
    sel.value = settings.voices[id] || '';
    const test = el('button', { type: 'button', class: 'btn sm', onclick: () => {
      const prevV = settings.voices[id]; const prevR = settings.rate;
      settings.voices[id] = sel.value; settings.rate = Number($('#setRate').value);
      speak({ en: 'Hello! Nice to meet you.', ja: 'こんにちは。はじめまして。', es: '¡Hola! Mucho gusto.' }[id], id);
      settings.voices[id] = prevV; settings.rate = prevR;
    } }, '▶ 들어 보기');
    box.append(el('label', { class: 'field' }, el('span', null, `${L.flag} ${L.name} 목소리`), el('div', { class: 'row' }, el('div', { class: 'grow' }, sel), test)));
  }
  if (!voices.length) box.append(el('p', { class: 'small muted' }, '이 브라우저에서 사용할 수 있는 음성 목록을 아직 불러오지 못했습니다. 잠시 후 다시 열어 주세요.'));
}
function saveSettingsDialog() {
  const persist = $('#setKeyPersist').checked;
  const key = getKey();
  if (key) setKey(key, persist); else settings.keyPersist = persist;
  const before = { ...settings.models };
  settings.models = { partner: $('#setPartner').value.trim() || before.partner, coach: $('#setCoach').value.trim() || before.coach, report: $('#setReport').value.trim() || before.report };
  const preset = Object.entries(PRESETS).find(([, p]) => p.partner === settings.models.partner && p.coach === settings.models.coach && p.report === settings.models.report);
  settings.preset = preset ? preset[0] : 'custom';
  for (const row of document.querySelectorAll('#priceRows .price-row')) {
    const [i, o] = row.querySelectorAll('input');
    const vi = parseFloat(i.value); const vo = parseFloat(o.value);
    if (isFinite(vi) && isFinite(vo)) settings.prices[row.dataset.model] = { in: vi, out: vo };
  }
  settings.fx = Math.max(1, parseFloat($('#setFx').value) || 1400);
  settings.autoSpeak = $('#setAutoSpeak').checked;
  settings.autoSend = $('#setAutoSend').checked;
  settings.rate = Number($('#setRate').value) || 1;
  settings.turns = Math.min(40, Math.max(4, parseInt($('#setTurns').value, 10) || 12));
  for (const sel of document.querySelectorAll('#voiceRows select')) settings.voices[sel.dataset.lang] = sel.value;
  saveSettings();
  if (session && !session.endedAt) { session.models = { ...settings.models }; saveSession(); }
  $('#settingsDlg').close();
  toast('설정을 저장했습니다.');
  if (!$('#view-setup').classList.contains('hidden')) renderSetup();
  if (!$('#view-chat').classList.contains('hidden')) { renderChatSub(); renderLiveCost(); }
}
function initSettings() {
  $('#btnSettings').onclick = openSettings;
  $('#setSave').onclick = saveSettingsDialog;
  $('#setRate').oninput = () => { $('#rateVal').textContent = `${Number($('#setRate').value).toFixed(2)}배`; };
  for (const id of ['#setPartner', '#setCoach', '#setReport']) $(id).addEventListener('change', renderPriceRows);
  $('#setKeyChange').onclick = () => { $('#settingsDlg').close(); showView('key'); $('#keyInput').focus(); };
  $('#setKeyClear').onclick = () => {
    if (!confirm('저장된 API 키를 지울까요?')) return;
    clearKey();
    $('#settingsDlg').close();
    toast('API 키를 지웠습니다.');
    showView('key');
  };
  $('#setKeyTest').onclick = () => {
    const k = getKey();
    if (!k) { $('#setKeyResult').textContent = '먼저 키를 입력해 주세요.'; return; }
    testKey(k, $('#setKeyResult')).then(renderPriceRows);
  };
  $('#setReset').onclick = () => {
    if (!confirm('모든 설정을 기본값으로 되돌릴까요? (API 키는 그대로 둡니다)')) return;
    const persist = settings.keyPersist;
    settings = deepMerge(DEFAULT_SETTINGS, { keyPersist: persist });
    saveSettings();
    $('#settingsDlg').close();
    toast('설정을 초기화했습니다.');
    if (!$('#view-setup').classList.contains('hidden')) renderSetup();
  };
}

/* ───────────── 입력창 ───────────── */
function autoResize() {
  const t = $('#input');
  t.style.height = 'auto';
  t.style.height = Math.min(160, t.scrollHeight + 2) + 'px';
}

/* ───────────── 시작 ───────────── */
function init() {
  initKeyView();
  initSettings();
  $('#startBtn').onclick = startSession;
  $('#resumeBtn').onclick = () => {
    if (!session) return;
    if (session.endedAt) { showView('report'); renderReport(); } else { openChat(); if (!session.messages.length) partnerTurn(); }
  };
  $('#discardBtn').onclick = () => {
    if (!confirm('이전 대화를 버릴까요? 저장하지 않은 내용은 복구할 수 없습니다.')) return;
    session = null; saveSession(); renderSetup();
  };
  $('#sendBtn').onclick = sendUser;
  $('#micBtn').onclick = toggleMic;
  $('#hintBtn').onclick = showHints;
  $('#endBtn').onclick = endSession;
  $('#turnEnd').onclick = endSession;
  $('#turnMore').onclick = () => $('#turnNotice').classList.add('hidden');
  $('#sideToggle').onclick = () => {
    const open = $('#chatSide').classList.toggle('open');
    $('#sideToggle').textContent = open ? '미션 접기' : '미션 보기';
  };
  $('#backChat').onclick = () => {
    session.endedAt = null; session.report = session.report || null; saveSession(); openChat();
  };
  $('#newBtn').onclick = () => {
    if (!session.saved && !confirm('아직 파일로 저장하지 않았습니다. 저장하지 않고 새 연습을 시작할까요?')) return;
    session = null; saveSession(); goSetup();
  };
  $('#expPdf').onclick = exportPdf;
  $('#expXlsx').onclick = exportXlsx;
  $('#expDocx').onclick = () => exportDocx().catch((e) => toast('Word 저장 실패: ' + e.message, true));
  const input = $('#input');
  input.addEventListener('input', autoResize);
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) { e.preventDefault(); sendUser(); }
  });
  window.addEventListener('beforeunload', (e) => {
    if (session && session.messages.some((m) => m.role === 'user') && !session.saved) { e.preventDefault(); e.returnValue = ''; }
  });
  goSetup();
}
init();
