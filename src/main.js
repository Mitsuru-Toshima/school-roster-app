import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { auth, googleProvider, db, storage } from './firebase-config.js';
import { 
  signInWithPopup, 
  signInWithRedirect, 
  getRedirectResult, 
  signOut, 
  onAuthStateChanged 
} from 'firebase/auth';
import { collection, doc, writeBatch, getDocs, setDoc, getDoc } from 'firebase/firestore';
import { ref, uploadBytes } from 'firebase/storage';

const SUPER_ADMIN_EMAIL = "mitsuru.toshima@hitachi-tech.ac.jp";

// 部活動名 フルネームマッピング
const CLUB_FULL_NAMES = {
  '野': '野球', '野球': '野球',
  'サ': 'サッカー', 'サッカー': 'サッカー',
  'レ': 'バレー', 'バレー': 'バレー', 'バレーボール': 'バレー',
  'ス': 'バスケ', 'バスケ': 'バスケ', 'バスケットボール': 'バスケ',
  'ソ': 'テニス', 'テニス': 'テニス', 'ソフトテニス': 'テニス',
  '卓': '卓球', '卓球': '卓球',
  '剣': '剣道', '剣道': '剣道',
  '陸': '陸上', '陸上': '陸上', '陸上長距離': '陸上',
  '吹': '吹奏楽', '吹奏楽': '吹奏楽',
  '五': '技能五輪', 'モ': '技能五輪', '技能五輪': '技能五輪'
};

function getFullClubName(abbr) {
  if (!abbr) return '-';
  const clean = abbr.trim().replace(/[（\(].*?[）\)]/g, '');
  return CLUB_FULL_NAMES[clean] || CLUB_FULL_NAMES[abbr.trim()] || abbr.trim();
}

function getPrefCity(addr) {
  if (!addr) return '-';
  const m = addr.match(/^(.*?[都道府県])(.*?市|.*?区|.*?町|.*?村|.*?郡.*?町|.*?郡.*?村)/);
  if (m) return m[1] + m[2];
  const m2 = addr.match(/^(.*?市|.*?区|.*?町|.*?村)/);
  if (m2) return m2[1];
  return addr.substring(0, 10);
}

function getShortAddress(addr) {
  if (!addr) return '-';
  const m1 = addr.match(/^(.*?\d+丁目)/);
  if (m1) return m1[1];
  const m2 = addr.match(/^(.*?[一二三四五六七八九十]+丁目)/);
  if (m2) return m2[1];
  const m3 = addr.match(/^(.*?[^\d\-ー番号地]+)/);
  if (m3) return m3[1].replace(/[-ー番地号 ]+$/, '');
  return addr;
}

const PHOTO_SIZES = [
  { name: "証明写真 4.0 × 3.0 cm（一般・履歴書）", w: 30, h: 40 },
  { name: "証明写真 4.5 × 3.5 cm（パスポート）", w: 35, h: 45 },
  { name: "証明写真 3.0 × 2.4 cm（免許証等）", w: 24, h: 30 },
  { name: "証明写真 5.0 × 4.0 cm", w: 40, h: 50 },
  { name: "証明写真 5.0 × 5.0 cm（ビザ等）", w: 50, h: 50 },
  { name: "証明写真 6.0 × 4.0 cm", w: 40, h: 60 }
];

const LOCKER_PLATE_SIZES = [
  { name: "ロッカー用 3.0 × 6.4 cm", w: 64, h: 30, font: "22pt" },
  { name: "ロッカー用 1.8 × 4.4 cm", w: 44, h: 18, font: "14pt" },
  { name: "ロッカー用 2.2 × 5.8 cm", w: 58, h: 22, font: "17pt" },
  { name: "ロッカー用 2.3 × 6.3 cm", w: 63, h: 23, font: "18pt" },
  { name: "ロッカー用 3.5 × 6.8 cm", w: 68, h: 35, font: "26pt" }
];

const PRESET_CLASSES = ["1-1", "1-2", "1-3", "1E", "1M", "1W", "2E", "2M", "2W", "3E1", "3E2", "3MW"];

const PRESET_CLASSROOMS = [
  { name: "23番教室（1-1・1E）", defaultClass: "1-1", cols: 5, rows: 5 },
  { name: "22番教室（1-2・1M）", defaultClass: "1-2", cols: 5, rows: 5 },
  { name: "21番教室（1-3・1W）", defaultClass: "1-3", cols: 5, rows: 5 },
  { name: "36番教室（2E）", defaultClass: "2E", cols: 5, rows: 5 },
  { name: "35番教室（2M）", defaultClass: "2M", cols: 5, rows: 5 },
  { name: "34番教室（2W）", defaultClass: "2W", cols: 5, rows: 5 },
  { name: "33番教室（3E1）", defaultClass: "3E1", cols: 5, rows: 5 },
  { name: "32番教室（3E2）", defaultClass: "3E2", cols: 5, rows: 5 },
  { name: "31番教室（3MW）", defaultClass: "3MW", cols: 5, rows: 5 },
  { name: "特別教室（講座）", defaultClass: "", cols: 6, rows: 11, isSpecialLecture: true },
  { name: "特別教室（グループ学習）", defaultClass: "", cols: 3, rows: 3, isGroupStudy: true },
  { name: "11番教室（定員16名）", defaultClass: "", cols: 4, rows: 4 },
  { name: "12番教室（歴史展示室・定員8名）", defaultClass: "", cols: 4, rows: 2 },
  { name: "会議室（定員12名）", defaultClass: "", cols: 4, rows: 3 },
  { name: "13番教室（考動アクション室）", defaultClass: "", cols: 5, rows: 5 },
  { name: "14番教室（安全教室）", defaultClass: "", cols: 5, rows: 5 },
  { name: "15番教室（ICT室）", defaultClass: "", cols: 6, rows: 5 },
  { name: "16番教室（倉庫）", defaultClass: "", cols: 4, rows: 4 },
  { name: "多目的室", defaultClass: "", cols: 6, rows: 5 },
  { name: "職員室", defaultClass: "", cols: 6, rows: 5 },
  { name: "印刷室", defaultClass: "", cols: 4, rows: 3 },
  { name: "校長室", defaultClass: "", cols: 3, rows: 2 },
  { name: "放送室", defaultClass: "", cols: 3, rows: 2 },
  { name: "資料室", defaultClass: "", cols: 4, rows: 4 }
];

let students = [];
let staffList = [];
let photoStore = {};
let idPhotoStore = {};
let currentTab = "seat";
let currentDiaryType = "day";

let seatAssignments = [];
let selectedSlotForSwap = null;
let dragSourceIndex = null;
let currentRollCallDate = "2026-11-06";
let rollCallRecords = {};

let savedDiaries = { day: {}, night: {}, health: {} };

let currentDiary = {
  day: {
    date: "2026-11-08",
    dutyPerson: "村上",
    weather: "晴れ",
    schoolType: "休日",
    dormPatrols: ["8:45～9:45", "", "", "", ""],
    schoolPatrols: ["9:45～10:15", "12:45～13:45", "", "", ""],
    club: "なし",
    notice: "特になし",
    special: "英検三級試験（茨城キリスト）希望者のみ",
    isolationStudentId: "",
    isolationTemp: "",
    isolationReason: "足首捻挫のため、松葉杖使用。"
  },
  night: {
    date: "2026-11-06",
    dutyPerson: "荻野目",
    weather: "晴れ",
    schoolType: "登校日",
    arriveTime: "16:20",
    lastLeave: "19:00（相馬）",
    firstArrive: "6:05（豊田）",
    rollcallStatus: "（夜）20:45 異常なし ・ （朝）6:45 異常なし",
    patrols: ["20:00～20:30", "21:45～22:10", "", "", ""],
    club: "有り",
    officeDuty: "3-401（机）（橋本）（興野）",
    emergencyStay: "無し",
    outReport: "20名帰寮報告有り",
    repair: "無し",
    leftFood: "無し",
    bus: "無し",
    sickList: [
      { studentId: "", temp: "38.2", condition: "発熱、頭痛のため自宅療養" }
    ],
    other: "生徒間の挨拶が全然、出来ていない。部屋の整理整頓、こちらも殆どだめ。"
  },
  health: {
    date: "2026-11-06",
    stopCount: 0,
    absentCount: 2,
    lateCount: 1,
    earlyCount: 2,
    homeCareCount: 2,
    dormCareCount: 0,
    absenceItems: [
      { studentId: "", type: "欠席", temp: "37.8", symptom: "昨夜は解熱せず自宅で経過観察" },
      { studentId: "", type: "早退", temp: "36.5", symptom: "アキレス腱の痛みのため受診（MRI検査）" }
    ],
    visitorItems: [
      { studentId: "", temp: "36.6", reason: "除草作業中にふらつきがあり、保健室で水分補給と休養後復学" }
    ],
    clinicResult: "風邪症状軽減。抗原検査で陰性を確認して登校。"
  }
};

savedDiaries.day[currentDiary.day.date] = JSON.parse(JSON.stringify(currentDiary.day));
savedDiaries.night[currentDiary.night.date] = JSON.parse(JSON.stringify(currentDiary.night));
savedDiaries.health[currentDiary.health.date] = JSON.parse(JSON.stringify(currentDiary.health));

// ================= 名簿パース（指定列: Q, V, AG, CG, CJ, CW 完全準拠） =================
document.getElementById('rosterFile').addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'csv') {
    Papa.parse(file, {
      header: false,
      skipEmptyLines: false,
      complete: (res) => {
        students = parseStrict209Students(res.data);
        onRosterLoaded();
      }
    });
  } else {
    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(new Uint8Array(evt.target.result), { type: 'array' });
      const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
      students = parseStrict209Students(rawRows);
      onRosterLoaded();
    };
    reader.readAsArrayBuffer(file);
  }
});

function parseStrict209Students(rawRows) {
  if (!rawRows || rawRows.length === 0) return [];
  const cleanRows = rawRows.filter(r => r && r.length > 0 && r.some(c => String(c || "").trim() !== ""));
  const studentMap = new Map();

  // Excel列番号 (0始まりインデックス):
  // Q=16, V=21, AG=32, CG=84, CJ=87, CW=100
  const COL_Q  = 16;  // ふりがな
  const COL_V  = 21;  // 出身中学
  const COL_AG = 32;  // 帰省先住所
  const COL_CG = 84;  // 委員会
  const COL_CJ = 87;  // 寮生判定
  const COL_CW = 100; // 部屋番号

  cleanRows.forEach(r => {
    let code = null;
    for (let c of r) {
      const s = String(c || "").trim();
      if (/^202[4-6]\d{3}$/.test(s)) {
        code = s;
        break;
      }
    }
    if (!code) return; // ヘッダー・物品申請行・空行を除外

    if (!studentMap.has(code)) {
      studentMap.set(code, {
        id: "",
        studentCode: code,
        name: "",
        cleanName: "",
        no: 0,
        cls: "",
        kana: "",
        club: "",
        committee: "",
        address: "",
        juniorHigh: "",
        room: ""
      });
    }
    const item = studentMap.get(code);

    // 【最優先】リスト行の列数が十分にある場合、指定列（Q, V, AG, CG, CJ, CW）から直接引用
    if (r.length > COL_AG) {
      if (r.length > COL_Q && r[COL_Q]) item.kana = String(r[COL_Q]).trim();
      if (r.length > COL_V && r[COL_V]) item.juniorHigh = String(r[COL_V]).trim();
      if (r.length > COL_AG && r[COL_AG]) item.address = String(r[COL_AG]).trim();
      if (r.length > COL_CG && r[COL_CG]) item.committee = String(r[COL_CG]).trim();
      if (r.length > COL_CW && r[COL_CW]) {
        const rm = String(r[COL_CW]).trim();
        if (rm && rm !== "-" && rm !== "0") item.room = rm;
      }
    }

    // 氏名・クラス・出席番号・部活動の補完（表2または大名簿から）
    if (r.length >= 9 && String(r[3] || "").trim() === code) {
      let clsVal = String(r[4] || "").trim().replace("1G1", "1-1").replace("1G2", "1-2").replace("1G3", "1-3")
                                            .replace("2G1", "2-1").replace("2G2", "2-2").replace("2G3", "2-3");
      if (clsVal && !item.cls) item.cls = clsVal;
      const noVal = parseInt(r[5]);
      if (!isNaN(noVal) && !item.no) item.no = noVal;

      const nameVal = String(r[8] || "").trim();
      if (nameVal && !item.name) {
        item.name = nameVal.replace(/[\s　]+/g, " ");
        item.cleanName = nameVal.replace(/[\s　]+/g, "");
      }
      if (r[7] && !item.club) item.club = getFullClubName(String(r[7]).trim());
    }

    // 表1形式（先頭がコードの行）のフォールバック
    if (r.length >= 12 && String(r[0] || "").trim() === code) {
      const nameVal = String(r[1] || "").trim();
      if (nameVal && !item.name) {
        item.name = nameVal.replace(/[\s　]+/g, " ");
        item.cleanName = nameVal.replace(/[\s　]+/g, "");
      }
      if (!item.kana && r[2]) item.kana = String(r[2]).trim();
      if (!item.juniorHigh && r[8]) item.juniorHigh = String(r[8]).trim();
      if (!item.address && r[11]) item.address = String(r[11]).trim();
    }
  });

  const resultList = [];
  studentMap.forEach(item => {
    if (!item.cls) {
      if (item.studentCode.startsWith("2026")) item.cls = "1-1";
      else if (item.studentCode.startsWith("2025")) item.cls = "2E";
      else if (item.studentCode.startsWith("2024")) item.cls = "3E1";
    }
    if (!item.no) item.no = resultList.length + 1;
    item.id = `${item.cls}_${String(item.no).padStart(2, '0')}`;
    item.club = getFullClubName(item.club);
    if (item.name) resultList.push(item);
  });

  return resultList;
}

function onRosterLoaded() {
  document.getElementById('rosterCount').innerText = `生徒: ${students.length}名`;
  document.getElementById('rosterCount').classList.add('badge-ok');
  updateFilterSelectors();
  initSeatAssignments();
  renderPreview();
}

// ================= 写真登録・照合 =================
function handlePhotoFiles(files, targetStore, countElId) {
  let count = 0;
  Array.from(files).forEach(f => {
    if (!f.type.startsWith('image/') || f.name.startsWith('._')) return;
    const reader = new FileReader();
    const rawName = f.name.replace(/\.[^/.]+$/, "");
    reader.onload = (evt) => {
      const digits = rawName.match(/\d{4,}/);
      if (digits) targetStore[digits[0]] = evt.target.result;
      targetStore[rawName.replace(/[\s_\-　]/g, "")] = evt.target.result;
      targetStore[rawName] = evt.target.result;
      count++;
      document.getElementById(countElId).innerText = `写真: ${Object.keys(targetStore).length}枚 登録済`;
      document.getElementById(countElId).classList.add('badge-ok');
      renderPreview();
    };
    reader.readAsDataURL(f);
  });
}

function handleZipPhotos(file, targetStore, countElId) {
  JSZip.loadAsync(file).then(zip => {
    let count = 0;
    const promises = [];
    zip.forEach((_, entry) => {
      if (entry.dir || entry.name.includes('__MACOSX')) return;
      const fileName = entry.name.split(/[\/\\]/).pop();
      if (fileName.startsWith('._') || !/\.(jpe?g|png|webp|bmp)$/i.test(fileName)) return;

      const p = entry.async("base64").then((base64) => {
        const ext = fileName.split('.').pop().toLowerCase();
        const mime = ext === 'png' ? 'image/png' : 'image/jpeg';
        const dataUrl = `data:${mime};base64,${base64}`;
        const rawName = fileName.replace(/\.[^/.]+$/, "");
        const cleanKey = rawName.replace(/[\s_\-　]/g, "");

        const digits = rawName.match(/\d{4,}/);
        if (digits) targetStore[digits[0]] = dataUrl;
        targetStore[cleanKey] = dataUrl;
        targetStore[rawName] = dataUrl;
        count++;
      });
      promises.push(p);
    });
    Promise.all(promises).then(() => {
      document.getElementById(countElId).innerText = `写真: ${Object.keys(targetStore).length}枚 登録済`;
      document.getElementById(countElId).classList.add('badge-ok');
      renderPreview();
    });
  });
}

document.getElementById('photoZipFile').addEventListener('change', (e) => handleZipPhotos(e.target.files[0], photoStore, 'photoCount'));
document.getElementById('photoDir').addEventListener('change', (e) => handlePhotoFiles(e.target.files, photoStore, 'photoCount'));

document.getElementById('idPhotoZipFile').addEventListener('change', (e) => handleZipPhotos(e.target.files[0], idPhotoStore, 'idPhotoCount'));
document.getElementById('idPhotoDir').addEventListener('change', (e) => handlePhotoFiles(e.target.files, idPhotoStore, 'idPhotoCount'));

function getPhoto(student, isIdPhoto = false) {
  if (!student) return "";
  const store = isIdPhoto ? (Object.keys(idPhotoStore).length > 0 ? idPhotoStore : photoStore) : photoStore;

  const code = (student.studentCode || "").trim();
  const name = (student.cleanName || "").trim();
  const numPad = String(student.no).padStart(2, '0');
  const numRaw = String(student.no);
  const cls = (student.cls || "").replace(/[\s_\-　]/g, "");

  if (code && store[code]) return store[code];
  if (code) {
    for (let k in store) if (k.includes(code)) return store[k];
  }
  if (name && store[name]) return store[name];
  if (name) {
    for (let k in store) if (k.includes(name)) return store[k];
  }
  for (let k in store) {
    const ck = k.replace(/[\s_\-　]/g, "");
    if (ck.includes(`${cls}${numPad}`) || ck.includes(`${cls}${numRaw}`)) return store[k];
    if (cls.includes("-")) {
      const parts = cls.split("-");
      if (ck.includes(`${parts[0]}${parts[1]}${numPad}`) || ck.includes(`${parts[1]}${numPad}`)) return store[k];
    }
  }
  return "";
}

// ================= 座席表配置（右端教卓側から縦順昇順・特別教室2種） =================
function arrangeSeatSlots(studentList, pattern, cols, rows) {
  const grid = Array.from({ length: rows }, () => Array(cols).fill(null));
  const n = studentList.length;

  if (pattern === 'podium_col') {
    const colCounts = Array(cols).fill(Math.floor(n / cols));
    for (let i = 0; i < (n % cols); i++) colCounts[cols - 1 - i] += 1;

    let curr = 0;
    for (let c = cols - 1; c >= 0; c--) {
      const count = Math.min(colCounts[c], rows);
      for (let rOffset = 0; rOffset < count; rOffset++) {
        const r = (rows - 1) - rOffset;
        if (curr < n) grid[r][c] = studentList[curr++];
      }
    }
  } else if (pattern === 'podium_row_lr') {
    let curr = 0;
    for (let r = rows - 1; r >= 0; r--) {
      for (let c = 0; c < cols; c++) if (curr < n) grid[r][c] = studentList[curr++];
    }
  } else if (pattern === 'podium_row_rl') {
    let curr = 0;
    for (let r = rows - 1; r >= 0; r--) {
      for (let c = cols - 1; c >= 0; c--) if (curr < n) grid[r][c] = studentList[curr++];
    }
  } else if (pattern === 'desc') {
    let curr = 0;
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) if (curr < n) grid[r][c] = studentList[curr++];
    }
  } else {
    let curr = 0;
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) if (curr < n) grid[r][c] = studentList[curr++];
    }
  }

  const flattened = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) flattened.push({ slot: r * cols + c, student: grid[r][c] });
  }
  return flattened;
}

function initSeatAssignments(customList = null, customPattern = null) {
  const roomName = document.getElementById('classroomPresetSelect')?.value || "";
  const roomObj = PRESET_CLASSROOMS.find(r => r.name === roomName);

  if (roomObj && roomObj.isGroupStudy) {
    const total = 54;
    document.getElementById('totalSeatsLabel').innerText = `(計54席・9グループ)`;
    const list = customList ? [...customList] : getTargetStudentsForSeating();
    list.sort((a, b) => a.no - b.no);
    seatAssignments = Array.from({ length: total }, (_, i) => ({ slot: i, student: list[i] || null }));
    return;
  }

  const cols = parseInt(document.getElementById('seatCols').value) || 5;
  const rows = parseInt(document.getElementById('seatRows').value) || 5;
  document.getElementById('totalSeatsLabel').innerText = `(計${cols * rows}席)`;

  const pattern = customPattern || (document.getElementById('seatOrderPattern')?.value || 'podium_col');
  let list = customList ? [...customList] : getTargetStudentsForSeating();

  if (pattern === 'desc') list.sort((a, b) => b.no - a.no);
  else if (pattern !== 'random') list.sort((a, b) => a.no - b.no);

  if (pattern === 'random') {
    list.sort(() => Math.random() - 0.5);
    seatAssignments = Array.from({ length: cols * rows }, (_, i) => ({ slot: i, student: list[i] || null }));
  } else {
    seatAssignments = arrangeSeatSlots(list, pattern, cols, rows);
  }
}

function getTargetStudentsForSeating() {
  const mode = document.getElementById('seatGroupingMode').value;
  if (mode === 'single') {
    const targetCls = (document.getElementById('seatClassSelect').value || '').trim().toUpperCase();
    return students.filter(s => (s.cls || '').trim().toUpperCase() === targetCls);
  } else if (mode === 'combined') {
    const selected = Array.from(document.getElementById('seatCombinedSelect').selectedOptions).map(o => o.value.trim().toUpperCase());
    return students.filter(s => selected.includes((s.cls || '').trim().toUpperCase()));
  } else if (mode === 'club') {
    const club = document.getElementById('seatClubSelect').value;
    return students.filter(s => s.club === club);
  } else if (mode === 'committee') {
    const comm = document.getElementById('seatCommSelect').value;
    return students.filter(s => s.committee === comm);
  }
  return [...students];
}

function updateFilterSelectors() {
  const seatCls = document.getElementById('seatClassSelect');
  const combCls = document.getElementById('seatCombinedSelect');
  const genCls = document.getElementById('generalClassFilter');
  const clubSel = document.getElementById('seatClubSelect');
  const commSel = document.getElementById('seatCommSelect');

  const classes = PRESET_CLASSES;
  if (seatCls) seatCls.innerHTML = classes.map(c => `<option value="${c}">${c}</option>`).join('');
  if (combCls) combCls.innerHTML = classes.map(c => `<option value="${c}">${c}</option>`).join('');
  if (genCls) genCls.innerHTML = '<option value="ALL">全クラス</option>' + classes.map(c => `<option value="${c}">${c}</option>`).join('');

  const clubs = ['野球', 'サッカー', 'バレー', 'バスケ', 'テニス', '卓球', '剣道', '陸上', '吹奏楽', '技能五輪'];
  if (clubSel) clubSel.innerHTML = clubs.map(c => `<option value="${c}">${c}</option>`).join('');

  const comms = ['環境整備委員', '体育委員', '文化委員', '図書委員', '広報委員', 'ICT委員', '生活委員', '保健委員'];
  if (commSel) commSel.innerHTML = comms.map(c => `<option value="${c}">${c}</option>`).join('');
}

function initPresets() {
  const roomSel = document.getElementById('classroomPresetSelect');
  if (roomSel) {
    roomSel.innerHTML = PRESET_CLASSROOMS.map(r => `<option value="${r.name}">${r.name}</option>`).join('');
    roomSel.value = "23番教室（1-1・1E）";

    roomSel.addEventListener('change', (e) => {
      const selectedRoom = PRESET_CLASSROOMS.find(r => r.name === e.target.value);
      if (!selectedRoom) return;

      const colsRowsBox = document.getElementById('seatColsRowsBox');
      if (selectedRoom.isGroupStudy) {
        colsRowsBox.style.display = 'none';
      } else {
        colsRowsBox.style.display = 'flex';
        document.getElementById('seatCols').value = selectedRoom.cols;
        document.getElementById('seatRows').value = selectedRoom.rows;
      }

      if (selectedRoom.defaultClass) {
        document.getElementById('seatClassSelect').value = selectedRoom.defaultClass;
      }
      initSeatAssignments();
      renderPreview();
    });
  }

  updateFilterSelectors();
  initSeatAssignments();
}

// ================= コントロールイベント =================
document.getElementById('btnRefreshSeat')?.addEventListener('click', () => { initSeatAssignments(); renderPreview(); });
document.getElementById('btnRefreshGeneral')?.addEventListener('click', renderPreview);

document.getElementById('seatGroupingMode').addEventListener('change', (e) => {
  const m = e.target.value;
  document.getElementById('singleClassBox').style.display = (m === 'single') ? 'flex' : 'none';
  document.getElementById('combinedClassBox').style.display = (m === 'combined') ? 'flex' : 'none';
  document.getElementById('clubSelectBox').style.display = (m === 'club') ? 'flex' : 'none';
  document.getElementById('commSelectBox').style.display = (m === 'committee') ? 'flex' : 'none';
  initSeatAssignments();
  renderPreview();
});

document.getElementById('seatClassSelect').addEventListener('change', () => { initSeatAssignments(); renderPreview(); });
document.getElementById('seatCombinedSelect').addEventListener('change', () => { initSeatAssignments(); renderPreview(); });
document.getElementById('seatClubSelect').addEventListener('change', () => { initSeatAssignments(); renderPreview(); });
document.getElementById('seatCommSelect').addEventListener('change', () => { initSeatAssignments(); renderPreview(); });
document.getElementById('seatOrderPattern').addEventListener('change', () => { initSeatAssignments(); renderPreview(); });
document.getElementById('seatCols').addEventListener('change', () => { initSeatAssignments(); renderPreview(); });
document.getElementById('seatRows').addEventListener('change', () => { initSeatAssignments(); renderPreview(); });

document.getElementById('btnSortRandom')?.addEventListener('click', () => { initSeatAssignments(null, 'random'); renderPreview(); });
document.getElementById('btnClearSeats')?.addEventListener('click', () => {
  const cols = parseInt(document.getElementById('seatCols').value) || 5;
  const rows = parseInt(document.getElementById('seatRows').value) || 5;
  seatAssignments = Array.from({ length: cols * rows }, (_, i) => ({ slot: i, student: null }));
  renderPreview();
});

document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', (e) => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    e.target.classList.add('active');
    currentTab = e.target.dataset.tab;

    document.getElementById('seatControls').style.display = (currentTab === 'seat') ? 'flex' : 'none';
    document.getElementById('rollcallControls').style.display = (currentTab === 'rollcall') ? 'flex' : 'none';
    document.getElementById('diaryControls').style.display = (currentTab === 'diaries') ? 'flex' : 'none';

    const genControls = document.getElementById('generalControls');
    const gradeBox = document.getElementById('gradeFilterBox');
    const classBox = document.getElementById('classFilterBox');
    const sizeBox = document.getElementById('sizeSelectorBox');
    const sizeSel = document.getElementById('sizeSelector');

    if (currentTab === 'photo') {
      genControls.style.display = 'flex';
      gradeBox.style.display = 'flex';
      classBox.style.display = 'flex';
      sizeBox.style.display = 'flex';
      sizeSel.innerHTML = PHOTO_SIZES.map((s, idx) => `<option value="${idx}">${s.name}</option>`).join('');
    } else if (currentTab === 'plate_locker') {
      genControls.style.display = 'flex';
      gradeBox.style.display = 'flex';
      classBox.style.display = 'flex';
      sizeBox.style.display = 'flex';
      sizeSel.innerHTML = LOCKER_PLATE_SIZES.map((s, idx) => `<option value="${idx}">${s.name}</option>`).join('');
    } else if (currentTab === 'attendance' || currentTab === 'plate_dorm' || currentTab === 'origin') {
      genControls.style.display = 'flex';
      gradeBox.style.display = 'flex';
      classBox.style.display = 'flex';
      sizeBox.style.display = 'none';
    } else {
      genControls.style.display = 'none';
    }

    renderPreview();
  });
});

document.getElementById('generalGradeFilter').addEventListener('change', renderPreview);
document.getElementById('generalClassFilter').addEventListener('change', renderPreview);
document.getElementById('sizeSelector').addEventListener('change', renderPreview);

// ================= プレビュー描画 =================
function renderPreview() {
  const container = document.getElementById('preview-area');
  if (students.length === 0 && currentTab !== 'diaries') {
    container.innerHTML = `<div class="empty-state"><p>生徒名簿ファイル（CSV / Excel）をアップロードしてください。</p></div>`;
    return;
  }

  switch (currentTab) {
    case 'seat':
      renderSeatingChart(container);
      break;
    case 'attendance':
      renderAttendanceSheet(container);
      break;
    case 'rollcall':
      renderRollCallSystem(container);
      break;
    case 'plate_locker':
      renderLockerPlates(container);
      break;
    case 'plate_dorm':
      renderDormPhotoPlates(container);
      break;
    case 'photo':
      renderPhotoSheet(container);
      break;
    case 'diaries':
      renderDiariesTab(container);
      break;
    case 'origin':
      renderOriginSheet(container);
      break;
  }
}

// 1. 座席表（写真下枠極小・空席非表示・A4縦1枚収容）
function renderSeatingChart(container) {
  const roomName = document.getElementById('classroomPresetSelect').value || "座席表";
  const roomObj = PRESET_CLASSROOMS.find(r => r.name === roomName);

  if (roomObj && roomObj.isGroupStudy) {
    renderGroupStudySeating(container, roomName);
    return;
  }

  const cols = parseInt(document.getElementById('seatCols').value) || 5;
  const rows = parseInt(document.getElementById('seatRows').value) || 5;

  let html = `
    <div class="seating-chart">
      <div class="seating-header-block">
        <div class="seating-room-title">${roomName}</div>
        <div class="seating-sub-info">座席表（出席簿用）</div>
      </div>
      <div class="seating-grid-container" style="grid-template-columns: repeat(${cols}, 1fr);">
  `;

  seatAssignments.forEach((slotData, idx) => {
    const s = slotData.student;
    const isSelected = selectedSlotForSwap === idx ? "selected-swap" : "";

    if (!s) {
      // 空席は非表示（透明化・枠なし）
      html += `<div class="seat-slot empty-seat" style="visibility:hidden; border:none;"></div>`;
    } else {
      const photo = getPhoto(s);
      html += `
        <div class="seat-slot ${isSelected}" draggable="true" onclick="handleSeatClick(${idx})" ondragstart="handleDragStart(event, ${idx})" ondragover="handleDragOver(event)" ondrop="handleDrop(event, ${idx})">
          <div class="seat-photo-box">
            ${photo ? `<img class="seat-photo" src="${photo}">` : `<div class="no-photo">写真</div>`}
          </div>
          <!-- 写真下枠極小表示（クラス名なし・番号＋氏名のみ） -->
          <div class="seat-info-tiny">
            <span class="seat-name-compact">${s.no}. ${s.name}</span>
          </div>
        </div>
      `;
    }
  });

  html += `
      </div>
      <div class="podium-box">教 卓</div>
    </div>
  `;
  container.innerHTML = html;
}

function renderGroupStudySeating(container, roomName) {
  let html = `
    <div class="seating-chart">
      <div class="seating-header-block">
        <div class="seating-room-title">${roomName}</div>
        <div class="seating-sub-info">座席表（3×3グループ学習 / 各島6名対面）</div>
      </div>
      <div class="group-islands-grid">
  `;

  for (let g = 0; g < 9; g++) {
    html += `
      <div class="group-island-box">
        <div class="group-island-header">第${g + 1}グループ</div>
        <div class="group-island-seats">
    `;
    for (let sIdx = 0; sIdx < 6; sIdx++) {
      const slotIndex = g * 6 + sIdx;
      const slotData = seatAssignments[slotIndex];
      const s = slotData ? slotData.student : null;

      if (!s) {
        html += `<div class="group-seat-slot empty-seat" style="visibility:hidden;"></div>`;
      } else {
        const photo = getPhoto(s);
        html += `
          <div class="group-seat-slot">
            <div class="seat-photo-box-group">
              ${photo ? `<img src="${photo}">` : `<div class="no-photo">写真</div>`}
            </div>
            <div class="seat-info-tiny">
              <span class="seat-name-compact">${s.name}</span>
            </div>
          </div>
        `;
      }
    }
    html += `</div></div>`;
  }

  html += `
      </div>
      <div class="podium-box">教 卓</div>
    </div>
  `;
  container.innerHTML = html;
}

// 2. 生徒名簿（出席簿原本書式準拠・Q, V, AG, CG列完全引用・番地省略）
function renderAttendanceSheet(container) {
  const list = getFilteredStudents();

  let html = `
    <div class="attendance-sheet-wrapper">
      <div class="attendance-header">
        <div style="font-size:14pt; font-weight:bold;">生徒名簿（出席簿）</div>
        <div>対象人数: ${list.length}名</div>
      </div>
      <table class="attendance-table">
        <thead>
          <tr>
            <th style="width:40px;">番号</th>
            <th style="width:110px;">氏名</th>
            <th style="width:130px;">ふりがな</th>
            <th style="width:100px;">出身中学</th>
            <th>帰省先住所</th>
            <th style="width:90px;">部活動</th>
            <th style="width:100px;">委員会</th>
          </tr>
        </thead>
        <tbody>
  `;

  list.forEach(s => {
    html += `
      <tr>
        <td>${s.no}</td>
        <td style="font-weight:bold; font-size:10.5pt;">${s.name}</td>
        <td style="font-size:8.5pt;">${s.kana || '-'}</td>
        <td>${s.juniorHigh || '-'}</td>
        <td style="text-align:left; font-size:8.5pt;">${getShortAddress(s.address)}</td>
        <td>${s.club || '-'}</td>
        <td>${s.committee || '-'}</td>
      </tr>
    `;
  });

  html += `</tbody></table></div>`;
  container.innerHTML = html;
}

// 3. 寮生点呼簿（週間記録・CJ寮生抽出・CW部屋番号・フロア別ソート＆A4印刷）
function renderRollCallSystem(container) {
  const filter = document.getElementById('dormFilterSelect').value;

  let list = students.filter(s => s.room && s.room !== '-');
  list.sort((a, b) => (a.room || '').localeCompare(b.room || '', undefined, { numeric: true }));

  if (filter !== 'ALL') {
    list = list.filter(s => (s.room || '').startsWith(filter));
  }

  const absents = [];
  list.forEach(s => {
    const rec = rollCallRecords[s.id];
    if (rec && rec.status && rec.status !== "在寮") {
      absents.push({ student: s, status: rec.status, note: rec.note || "" });
    }
  });

  document.getElementById('absentSummaryBadge').innerText = `不在者: ${absents.length}名`;
  const chipContainer = document.getElementById('absentListContainer');
  if (absents.length === 0) {
    chipContainer.innerHTML = `<span style="color:#666; font-size:12px;">現在、不在者として登録された生徒はいません。</span>`;
  } else {
    chipContainer.innerHTML = absents.map(a => `
      <span class="absent-chip">
        <strong>${a.student.name}</strong> (${a.student.cls} / ${a.student.room}) - <span style="color:#d32f2f; font-weight:bold;">${a.status}</span> ${a.note ? `[${a.note}]` : ''}
      </span>
    `).join('');
  }

  const baseDate = new Date(currentRollCallDate);
  const dayOfWeek = baseDate.getDay();
  const monday = new Date(baseDate);
  monday.setDate(baseDate.getDate() - ((dayOfWeek + 6) % 7));

  const weekDays = [];
  const dayNames = ['月', '火', '水', '木', '金', '土', '日'];
  for (let i = 0; i < 7; i++) {
    const d = new Date(monday);
    d.setDate(monday.getDate() + i);
    weekDays.push({
      dateStr: `${d.getMonth() + 1}/${d.getDate()}`,
      dayName: dayNames[i]
    });
  }

  let html = `
    <div class="rollcall-wrapper">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
        <h3>寮生点呼簿（週間記録: ${weekDays[0].dateStr}(月) ～ ${weekDays[6].dateStr}(日)）</h3>
        <div>フロア: <strong>${filter === 'ALL' ? '全寮生' : filter + '階'}</strong> (${list.length}名)</div>
      </div>
      <table class="rollcall-table">
        <thead>
          <tr>
            <th rowspan="2" style="width:50px;">部屋</th>
            <th rowspan="2" style="width:55px;">クラス</th>
            <th rowspan="2" style="width:105px;">氏名</th>
            <th rowspan="2" style="width:75px;">部活動</th>
            ${weekDays.map(w => `<th colspan="2">${w.dateStr} (${w.dayName})</th>`).join('')}
            <th rowspan="2" style="width:75px;">状態設定</th>
            <th rowspan="2">不在理由・特記メモ</th>
          </tr>
          <tr>
            <th>朝</th><th>夜</th><th>朝</th><th>夜</th><th>朝</th><th>夜</th><th>朝</th><th>夜</th><th>朝</th><th>夜</th><th>朝</th><th>夜</th><th>朝</th><th>夜</th>
          </tr>
        </thead>
        <tbody>
  `;

  list.forEach(s => {
    const rec = rollCallRecords[s.id] || { status: "在寮", note: "" };
    html += `
      <tr>
        <td style="font-weight:bold;">${s.room || '-'}</td>
        <td>${s.cls}</td>
        <td style="font-weight:bold;">${s.name}</td>
        <td>${s.club || '-'}</td>
        <td><input type="checkbox"></td><td><input type="checkbox"></td>
        <td><input type="checkbox"></td><td><input type="checkbox"></td>
        <td><input type="checkbox"></td><td><input type="checkbox"></td>
        <td><input type="checkbox"></td><td><input type="checkbox"></td>
        <td><input type="checkbox"></td><td><input type="checkbox"></td>
        <td><input type="checkbox"></td><td><input type="checkbox"></td>
        <td><input type="checkbox"></td><td><input type="checkbox"></td>
        <td>
          <select onchange="updateRollCallStatus('${s.id}', this.value)" style="font-size:10px;">
            <option value="在寮" ${rec.status === '在寮' ? 'selected' : ''}>在寮</option>
            <option value="帰省" ${rec.status === '帰省' ? 'selected' : ''}>帰省</option>
            <option value="公欠" ${rec.status === '公欠' ? 'selected' : ''}>公欠</option>
            <option value="自宅療養" ${rec.status === '自宅療養' ? 'selected' : ''}>自宅療養</option>
            <option value="寮隔離" ${rec.status === '寮隔離' ? 'selected' : ''}>寮隔離</option>
            <option value="外出中" ${rec.status === '外出中' ? 'selected' : ''}>外出中</option>
          </select>
        </td>
        <td><input type="text" value="${rec.note || ''}" onchange="updateRollCallNote('${s.id}', this.value)" style="width:95%; font-size:10px;"></td>
      </tr>
    `;
  });

  html += `</tbody></table></div>`;
  container.innerHTML = html;
}

window.updateRollCallStatus = function(id, val) {
  if (!rollCallRecords[id]) rollCallRecords[id] = {};
  rollCallRecords[id].status = val;
  renderPreview();
};

window.updateRollCallNote = function(id, val) {
  if (!rollCallRecords[id]) rollCallRecords[id] = {};
  rollCallRecords[id].note = val;
};

document.getElementById('btnPrevWeek')?.addEventListener('click', () => {
  const d = new Date(currentRollCallDate);
  d.setDate(d.getDate() - 7);
  currentRollCallDate = d.toISOString().slice(0, 10);
  document.getElementById('rollcallWeekDate').value = currentRollCallDate;
  renderPreview();
});

document.getElementById('btnNextWeek')?.addEventListener('click', () => {
  const d = new Date(currentRollCallDate);
  d.setDate(d.getDate() + 7);
  currentRollCallDate = d.toISOString().slice(0, 10);
  document.getElementById('rollcallWeekDate').value = currentRollCallDate;
  renderPreview();
});

document.getElementById('rollcallWeekDate')?.addEventListener('change', (e) => {
  currentRollCallDate = e.target.value;
  renderPreview();
});

document.getElementById('dormFilterSelect')?.addEventListener('change', renderPreview);
document.getElementById('btnPrintFloorRollCall')?.addEventListener('click', () => window.print());

// 4. ロッカー用ネームプレート（名前のみ枠いっぱいに表示・正確なA4実寸）
function renderLockerPlates(container) {
  const spec = LOCKER_PLATE_SIZES[document.getElementById('sizeSelector').value || 0];
  const list = getFilteredStudents();

  let html = `<div class="plate-sheet">`;
  list.forEach(s => {
    html += `
      <div class="name-plate" style="width:${spec.w}mm; height:${spec.h}mm;">
        <span class="plate-pure-name" style="font-size:${spec.font};">${s.name}</span>
      </div>
    `;
  });
  html += `</div>`;
  container.innerHTML = html;
}

// 5. ネームプレート（寮 4×3cm：写真面積8割以上、縦書き氏名、下部横並び）
function renderDormPhotoPlates(container) {
  const list = getFilteredStudents();

  let html = `<div class="plate-sheet">`;
  list.forEach(s => {
    const photo = getPhoto(s);
    html += `
      <div class="name-plate-dorm">
        <div class="dorm-plate-top">
          <div class="dorm-plate-photo-box">
            ${photo ? `<img src="${photo}">` : `<div class="no-photo">写真</div>`}
          </div>
          <div class="dorm-plate-vertical-name">${s.name}</div>
        </div>
        <div class="dorm-plate-bottom">
          <span>${s.room || '-'}</span>
          <span>${s.club || '-'}</span>
        </div>
      </div>
    `;
  });
  html += `</div>`;
  container.innerHTML = html;
}

// 6. 証明写真印刷（名前を写真枠外に配置しカット時正確寸法維持）
function renderPhotoSheet(container) {
  const spec = PHOTO_SIZES[document.getElementById('sizeSelector').value || 0];
  const list = getFilteredStudents();

  let html = `<div class="photo-sheet">`;
  list.forEach(s => {
    const photo = getPhoto(s, true);
    html += `
      <div class="photo-item-container">
        <div class="photo-card" style="width:${spec.w}mm; height:${spec.h}mm;">
          ${photo ? `<img src="${photo}">` : `<div class="no-photo">写真なし</div>`}
        </div>
        <div class="photo-external-label" style="width:${spec.w}mm;">${s.cls}-${s.no} ${s.name}</div>
      </div>
    `;
  });
  html += `</div>`;
  container.innerHTML = html;
}

// 7. 各種日誌（過去閲覧機能付き）
window.switchDiaryType = function(type) {
  currentDiaryType = type;
  document.querySelectorAll('.diary-subtab').forEach(b => b.classList.remove('active'));
  event.target.classList.add('active');
  updateDiaryHistorySelect();
  renderPreview();
};

function getWeekday(dateStr) {
  if (!dateStr) return '';
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  const d = new Date(dateStr);
  return days[d.getDay()] || '';
}

function renderStudentOptionList() {
  return `<option value="">-- 生徒を選択 --</option>` +
    students.map(s => `<option value="${s.id}">${s.name} (${s.cls} / ${s.room || '寮室無'})</option>`).join('');
}

function updateDiaryHistorySelect() {
  const historySel = document.getElementById('diaryHistorySelect');
  if (!historySel) return;
  const store = savedDiaries[currentDiaryType] || {};
  const dates = Object.keys(store).sort().reverse();
  historySel.innerHTML = '<option value="">-- 保存済み履歴から選択 --</option>' +
    dates.map(dt => `<option value="${dt}">${dt} (${getWeekday(dt)}) の日誌</option>`).join('');
}

window.loadPastDiary = function(dateStr) {
  if (!dateStr) return;
  const store = savedDiaries[currentDiaryType];
  if (store && store[dateStr]) {
    currentDiary[currentDiaryType] = JSON.parse(JSON.stringify(store[dateStr]));
    renderPreview();
  }
};

document.getElementById('btnSaveDiary')?.addEventListener('click', () => {
  const dt = currentDiary[currentDiaryType].date;
  if (!dt) return;
  savedDiaries[currentDiaryType][dt] = JSON.parse(JSON.stringify(currentDiary[currentDiaryType]));
  updateDiaryHistorySelect();
  alert(`【${currentDiaryType === 'day' ? '日直' : currentDiaryType === 'night' ? '宿直' : '保健室'}日誌】${dt} 分を正常に保存しました。`);
});

function renderDiariesTab(container) {
  const studentOptions = renderStudentOptionList();
  const period = document.getElementById('diarySummaryPeriod')?.value || 'single';

  if (period !== 'single') {
    renderDiarySummary(container, period);
    return;
  }

  if (currentDiaryType === 'day') {
    const d = currentDiary.day;
    const weekday = getWeekday(d.date);
    const selStudent = students.find(s => s.id === d.isolationStudentId);

    container.innerHTML = `
      <div class="diary-paper">
        <h2 style="text-align:center; border-bottom:2px solid #333; padding-bottom:6px; margin-bottom:12px;">日直日誌（学校休日用）</h2>
        <table class="diary-table">
          <tr>
            <th style="width:110px;">日時</th>
            <td><input type="date" value="${d.date}" onchange="currentDiary.day.date=this.value; renderPreview();"> (${weekday})</td>
            <th style="width:100px;">天候</th>
            <td>
              <select onchange="currentDiary.day.weather=this.value">
                <option value="晴れ" ${d.weather==='晴れ'?'selected':''}>晴れ</option>
                <option value="曇り" ${d.weather==='曇り'?'selected':''}>曇り</option>
                <option value="雨" ${d.weather==='雨'?'selected':''}>雨</option>
                <option value="雪" ${d.weather==='雪'?'selected':''}>雪</option>
              </select>
            </td>
          </tr>
          <tr>
            <th>日直者</th>
            <td><input type="text" value="${d.dutyPerson}" onchange="currentDiary.day.dutyPerson=this.value" placeholder="日直者氏名"></td>
            <th>学校</th>
            <td>
              <select onchange="currentDiary.day.schoolType=this.value">
                <option value="休日" selected>休日</option>
                <option value="登校日">登校日</option>
                <option value="寮閉鎖">寮閉鎖</option>
              </select>
            </td>
          </tr>
          <tr>
            <th>寮内巡回 (5回)</th>
            <td colspan="3">
              <div style="display:flex; gap:6px; flex-wrap:wrap;">
                ${[0,1,2,3,4].map(i => `
                  <span>${['①','②','③','④','⑤'][i]} <input type="text" value="${d.dormPatrols[i]}" onchange="currentDiary.day.dormPatrols[${i}]=this.value" style="width:100px;" placeholder="例: 8:45～9:45"></span>
                `).join('')}
              </div>
            </td>
          </tr>
          <tr>
            <th>校内巡回 (5回)</th>
            <td colspan="3">
              <div style="display:flex; gap:6px; flex-wrap:wrap;">
                ${[0,1,2,3,4].map(i => `
                  <span>${['①','②','③','④','⑤'][i]} <input type="text" value="${d.schoolPatrols[i]}" onchange="currentDiary.day.schoolPatrols[${i}]=this.value" style="width:100px;" placeholder="例: 9:45～10:15"></span>
                `).join('')}
              </div>
            </td>
          </tr>
          <tr>
            <th>部活</th>
            <td colspan="3"><input type="text" value="${d.club}" onchange="currentDiary.day.club=this.value" placeholder="活動部活動や活動なし等"></td>
          </tr>
          <tr>
            <th>巡回連絡事項</th>
            <td colspan="3"><textarea rows="2" onchange="currentDiary.day.notice=this.value" placeholder="巡回時の連絡事項・確認事項">${d.notice}</textarea></td>
          </tr>
          <tr>
            <th>特記事項</th>
            <td colspan="3"><textarea rows="2" onchange="currentDiary.day.special=this.value" placeholder="検定試験・行事等">${d.special}</textarea></td>
          </tr>
          <tr>
            <th>隔離・療養</th>
            <td colspan="3">
              <div style="margin-bottom:6px;">
                <select onchange="currentDiary.day.isolationStudentId=this.value; renderPreview();">${studentOptions}</select>
                <span style="margin-left:12px;">体温: <input type="text" value="${d.isolationTemp}" onchange="currentDiary.day.isolationTemp=this.value" style="width:60px;" placeholder="36.5"> ℃</span>
              </div>
              ${selStudent ? `
                <div style="background:#f8fafc; padding:6px; border:1px solid #e2e8f0; margin-bottom:6px; font-size:9pt;">
                  <strong>${selStudent.name}</strong> ｜ クラス: ${selStudent.cls} ｜ 部屋: ${selStudent.room || '-'} ｜ 部活: ${selStudent.club || '-'} ｜ 住所: ${getPrefCity(selStudent.address)}
                </div>
              ` : ''}
              <textarea rows="2" onchange="currentDiary.day.isolationReason=this.value" placeholder="病状・経過・通院状況を記入">${d.isolationReason}</textarea>
            </td>
          </tr>
        </table>
      </div>
    `;
  } else if (currentDiaryType === 'night') {
    const d = currentDiary.night;
    const weekday = getWeekday(d.date);

    container.innerHTML = `
      <div class="diary-paper">
        <h2 style="text-align:center; border-bottom:2px solid #333; padding-bottom:6px; margin-bottom:12px;">宿直日誌（寮毎日用）</h2>
        <table class="diary-table">
          <tr>
            <th style="width:110px;">日時</th>
            <td><input type="date" value="${d.date}" onchange="currentDiary.night.date=this.value; renderPreview();"> (${weekday})</td>
            <th style="width:100px;">天候</th>
            <td>
              <select onchange="currentDiary.night.weather=this.value">
                <option value="晴れ" ${d.weather==='晴れ'?'selected':''}>晴れ</option>
                <option value="曇り" ${d.weather==='曇り'?'selected':''}>曇り</option>
                <option value="雨" ${d.weather==='雨'?'selected':''}>雨</option>
                <option value="雪" ${d.weather==='雪'?'selected':''}>雪</option>
              </select>
            </td>
          </tr>
          <tr>
            <th>宿直者</th>
            <td><input type="text" value="${d.dutyPerson}" onchange="currentDiary.night.dutyPerson=this.value"></td>
            <th>学校</th>
            <td>
              <select onchange="currentDiary.night.schoolType=this.value">
                <option value="登校日" ${d.schoolType==='登校日'?'selected':''}>登校日</option>
                <option value="休日" ${d.schoolType==='休日'?'selected':''}>休日</option>
                <option value="寮閉鎖" ${d.schoolType==='寮閉鎖'?'selected':''}>寮閉鎖</option>
              </select>
            </td>
          </tr>
          <tr>
            <th>寮着時間</th>
            <td><input type="text" value="${d.arriveTime}" onchange="currentDiary.night.arriveTime=this.value" placeholder="16:20"></td>
            <th>最終退勤/最初出勤</th>
            <td>退勤: <input type="text" value="${d.lastLeave}" onchange="currentDiary.night.lastLeave=this.value" style="width:90px;"> / 出勤: <input type="text" value="${d.firstArrive}" onchange="currentDiary.night.firstArrive=this.value" style="width:90px;"></td>
          </tr>
          <tr>
            <th>点呼確認</th>
            <td colspan="3"><input type="text" value="${d.rollcallStatus}" onchange="currentDiary.night.rollcallStatus=this.value" style="width:90%;"></td>
          </tr>
          <tr>
            <th>寮内巡回 (5回)</th>
            <td colspan="3">
              <div style="display:flex; gap:6px; flex-wrap:wrap;">
                ${[0,1,2,3,4].map(i => `
                  <span>${['①','②','③','④','⑤'][i]} <input type="text" value="${d.patrols[i]}" onchange="currentDiary.night.patrols[${i}]=this.value" style="width:100px;"></span>
                `).join('')}
              </div>
            </td>
          </tr>
          <tr>
            <th>事務当番 / 外出</th>
            <td colspan="3">事務当番: <input type="text" value="${d.officeDuty}" onchange="currentDiary.night.officeDuty=this.value" style="width:140px;"> &nbsp; 外出報告: <input type="text" value="${d.outReport}" onchange="currentDiary.night.outReport=this.value" style="width:140px;"></td>
          </tr>
          <tr>
            <th>緊急在寮 / 残食 / 検食 / バス</th>
            <td colspan="3">在寮: <input type="text" value="${d.emergencyStay}" onchange="currentDiary.night.emergencyStay=this.value" style="width:60px;"> 残食: <input type="text" value="${d.leftFood}" onchange="currentDiary.night.leftFood=this.value" style="width:60px;"> 検食: <input type="text" value="${d.repair}" onchange="currentDiary.night.repair=this.value" style="width:60px;"> バス: <input type="text" value="${d.bus}" onchange="currentDiary.night.bus=this.value" style="width:60px;"></td>
          </tr>
          <tr>
            <th>体調不良者</th>
            <td colspan="3">
              ${d.sickList.map((item, idx) => {
                const sel = students.find(s => s.id === item.studentId);
                return `
                  <div style="margin-bottom:6px; border-bottom:1px dashed #ccc; padding-bottom:4px;">
                    <select onchange="currentDiary.night.sickList[${idx}].studentId=this.value; renderPreview();">
                      <option value="">-- 生徒を選択 --</option>
                      ${students.map(s => `<option value="${s.id}" ${s.id===item.studentId?'selected':''}>${s.name} (${s.cls} / ${s.room||'-'})</option>`).join('')}
                    </select>
                    体温: <input type="text" value="${item.temp}" onchange="currentDiary.night.sickList[${idx}].temp=this.value" style="width:50px;"> ℃
                    ${sel ? `<span>[部屋: ${sel.room||'-'} / 部活: ${sel.club||'-'} / 住所: ${getPrefCity(sel.address)}]</span>` : ''}
                    <input type="text" value="${item.condition}" onchange="currentDiary.night.sickList[${idx}].condition=this.value" style="width:40%;" placeholder="症状・自宅療養/隔離">
                  </div>
                `;
              }).join('')}
              <button class="btn btn-sm btn-light" onclick="currentDiary.night.sickList.push({studentId:'', temp:'', condition:''}); renderPreview();">＋ 体調不良者を追加</button>
            </td>
          </tr>
          <tr>
            <th>その他特記</th>
            <td colspan="3"><textarea rows="2" onchange="currentDiary.night.other=this.value">${d.other}</textarea></td>
          </tr>
        </table>
      </div>
    `;
  } else if (currentDiaryType === 'health') {
    const d = currentDiary.health;
    const weekday = getWeekday(d.date);

    container.innerHTML = `
      <div class="diary-paper">
        <h2 style="text-align:center; border-bottom:2px solid #333; padding-bottom:6px; margin-bottom:12px;">保健室日誌（平日用）</h2>
        <table class="diary-table">
          <tr>
            <th style="width:110px;">日時</th>
            <td colspan="3"><input type="date" value="${d.date}" onchange="currentDiary.health.date=this.value; renderPreview();"> (${weekday})</td>
          </tr>
          <tr>
            <th>概況内訳</th>
            <td colspan="3">
              欠席: <input type="number" value="${d.absentCount}" onchange="currentDiary.health.absentCount=this.value" style="width:40px;"> 名 ｜
              遅刻: <input type="number" value="${d.lateCount}" onchange="currentDiary.health.lateCount=this.value" style="width:40px;"> 名 ｜
              早退: <input type="number" value="${d.earlyCount}" onchange="currentDiary.health.earlyCount=this.value" style="width:40px;"> 名 ｜
              出停: <input type="number" value="${d.stopCount}" onchange="currentDiary.health.stopCount=this.value" style="width:40px;"> 名 （内、自宅療養: <input type="number" value="${d.homeCareCount}" onchange="currentDiary.health.homeCareCount=this.value" style="width:40px;"> 名、寮隔離: <input type="number" value="${d.dormCareCount}" onchange="currentDiary.health.dormCareCount=this.value" style="width:40px;"> 名）
            </td>
          </tr>
          <tr>
            <th>欠席・欠課者詳細</th>
            <td colspan="3">
              ${d.absenceItems.map((item, idx) => {
                const sel = students.find(s => s.id === item.studentId);
                return `
                  <div style="margin-bottom:6px; border-bottom:1px dashed #ccc; padding-bottom:4px;">
                    区分: <select onchange="currentDiary.health.absenceItems[${idx}].type=this.value"><option value="欠席" ${item.type==='欠席'?'selected':''}>欠席</option><option value="遅刻" ${item.type==='遅刻'?'selected':''}>遅刻</option><option value="早退" ${item.type==='早退'?'selected':''}>早退</option></select>
                    生徒: <select onchange="currentDiary.health.absenceItems[${idx}].studentId=this.value; renderPreview();"><option value="">-- 選択 --</option>${students.map(s => `<option value="${s.id}" ${s.id===item.studentId?'selected':''}>${s.name} (${s.cls})</option>`).join('')}</select>
                    体温: <input type="text" value="${item.temp}" onchange="currentDiary.health.absenceItems[${idx}].temp=this.value" style="width:50px;"> ℃
                    ${sel ? `<span>[部屋: ${sel.room||'-'} / 部活: ${sel.club||'-'} / 住所: ${getPrefCity(sel.address)}]</span>` : ''}
                    <input type="text" value="${item.symptom}" onchange="currentDiary.health.absenceItems[${idx}].symptom=this.value" style="width:35%;" placeholder="症状・経過">
                  </div>
                `;
              }).join('')}
              <button class="btn btn-sm btn-light" onclick="currentDiary.health.absenceItems.push({studentId:'', type:'欠席', temp:'', symptom:''}); renderPreview();">＋ 欠席・欠課者を追加</button>
            </td>
          </tr>
          <tr>
            <th>来室者記録</th>
            <td colspan="3">
              ${d.visitorItems.map((item, idx) => {
                const sel = students.find(s => s.id === item.studentId);
                return `
                  <div style="margin-bottom:6px; border-bottom:1px dashed #ccc; padding-bottom:4px;">
                    生徒: <select onchange="currentDiary.health.visitorItems[${idx}].studentId=this.value; renderPreview();"><option value="">-- 選択 --</option>${students.map(s => `<option value="${s.id}" ${s.id===item.studentId?'selected':''}>${s.name} (${s.cls})</option>`).join('')}</select>
                    体温: <input type="text" value="${item.temp}" onchange="diaryData.health.visitorItems[${idx}].temp=this.value" style="width:50px;"> ℃
                    ${sel ? `<span>[部屋: ${sel.room||'-'} / 部活: ${sel.club||'-'}]</span>` : ''}
                    <input type="text" value="${item.reason}" onchange="diaryData.health.visitorItems[${idx}].reason=this.value" style="width:40%;" placeholder="処置・理由">
                  </div>
                `;
              }).join('')}
              <button class="btn btn-sm btn-light" onclick="currentDiary.health.visitorItems.push({studentId:'', temp:'', reason:''}); renderPreview();">＋ 来室者を追加</button>
            </td>
          </tr>
          <tr>
            <th>その他（受診等）</th>
            <td colspan="3"><textarea rows="3" onchange="currentDiary.health.clinicResult=this.value">${d.clinicResult}</textarea></td>
          </tr>
        </table>
      </div>
    `;
  }
}

function renderDiarySummary(container, period) {
  const store = savedDiaries[currentDiaryType] || {};
  const dates = Object.keys(store).sort();

  let html = `
    <div class="diary-paper">
      <h2 style="text-align:center; border-bottom:2px solid #333; padding-bottom:6px; margin-bottom:12px;">
        ${currentDiaryType === 'day' ? '日直日誌' : currentDiaryType === 'night' ? '宿直日誌' : '保健室日誌'}（${period === 'week' ? '週間' : period === 'month' ? '月間' : period === 'quarter' ? '4半期' : period === 'half' ? '半期' : '年間'}まとめ）
      </h2>
      <table class="diary-table">
        <thead>
          <tr>
            <th style="width:110px;">日付</th>
            <th style="width:90px;">担当</th>
            <th style="width:80px;">天候/学校</th>
            <th>要点・巡回・特記事項</th>
          </tr>
        </thead>
        <tbody>
  `;

  if (dates.length === 0) {
    html += `<tr><td colspan="4" style="padding:20px; color:#888;">保存された過去日誌データがありません。「日誌を保存」ボタンで記録してください。</td></tr>`;
  } else {
    dates.forEach(dt => {
      const rec = store[dt];
      html += `
        <tr>
          <td><strong>${dt}</strong> (${getWeekday(dt)})</td>
          <td>${rec.dutyPerson || '-'}</td>
          <td>${rec.weather || '-'} / ${rec.schoolType || '-'}</td>
          <td style="text-align:left; font-size:9pt;">
            ${rec.notice ? `連絡: ${rec.notice} <br>` : ''}
            ${rec.special ? `特記: ${rec.special} <br>` : ''}
            ${rec.sick ? `体調不良: ${rec.sick} <br>` : ''}
            ${rec.clinicResult ? `受診等: ${rec.clinicResult}` : ''}
          </td>
        </tr>
      `;
    });
  }

  html += `</tbody></table></div>`;
  container.innerHTML = html;
}

// 8. 出身地別名簿
function renderOriginSheet(container) {
  const groups = {};
  students.forEach(s => {
    let reg = getPrefCity(s.address);
    if (!groups[reg]) groups[reg] = [];
    groups[reg].push(s);
  });

  let html = `<div class="attendance-sheet-wrapper"><div class="attendance-header"><div style="font-size:14pt; font-weight:bold;">出身地別名簿</div></div>`;
  for (const reg in groups) {
    html += `
      <h4 style="margin:12px 0 4px;border-left:4px solid #1e3a8a;padding-left:8px;">${reg} (${groups[reg].length}名)</h4>
      <table class="attendance-table">
        <thead><tr><th style="width:60px;">組-番</th><th style="width:120px;">氏名</th><th>帰省先住所</th><th style="width:90px;">部活動</th></tr></thead>
        <tbody>
          ${groups[reg].map(s => `<tr><td>${s.cls}-${s.no}</td><td style="font-weight:bold;">${s.name}</td><td style="text-align:left;">${getShortAddress(s.address)}</td><td>${s.club || '-'}</td></tr>`).join('')}
        </tbody>
      </table>
    `;
  }
  html += `</div>`;
  container.innerHTML = html;
}

function getFilteredStudents() {
  const grade = document.getElementById('generalGradeFilter').value;
  const cls = document.getElementById('generalClassFilter').value;

  return students.filter(s => {
    const matchGrade = (grade === 'ALL') || (s.cls && s.cls.startsWith(grade));
    const matchClass = (cls === 'ALL') || (s.cls === cls);
    return matchGrade && matchClass;
  });
}

document.getElementById('btnExportRollCallCSV')?.addEventListener('click', () => {
  const exportData = students.map(s => {
    const rec = rollCallRecords[s.id] || { status: '在寮', note: '' };
    return {
      '部屋': s.room || '',
      'クラス': s.cls,
      '出席番号': s.no,
      '氏名': s.name,
      '部活動': s.club,
      '点呼状態': rec.status,
      '不在理由メモ': rec.note
    };
  });
  downloadBomCsv(`寮生点呼簿_${currentRollCallDate}.csv`, Papa.unparse(exportData));
});

document.getElementById('btnExportDiaryCSV')?.addEventListener('click', () => {
  const rows = [];
  ['day', 'night', 'health'].forEach(t => {
    const st = savedDiaries[t] || {};
    Object.keys(st).forEach(dt => {
      rows.push({ タイプ: t, 日付: dt, データ: JSON.stringify(st[dt]) });
    });
  });
  downloadBomCsv(`学校日誌データ一括_${new Date().toISOString().slice(0,10)}.csv`, Papa.unparse(rows));
});

function downloadBomCsv(filename, content) {
  const bom = new Uint8Array([0xEF, 0xBB, 0xBF]);
  const blob = new Blob([bom, content], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
}

window.handleSeatClick = function(index) {
  if (selectedSlotForSwap === null) {
    selectedSlotForSwap = index;
    renderPreview();
  } else {
    const temp = seatAssignments[selectedSlotForSwap].student;
    seatAssignments[selectedSlotForSwap].student = seatAssignments[index].student;
    seatAssignments[index].student = temp;
    selectedSlotForSwap = null;
    renderPreview();
  }
};

window.handleDragStart = function(e, index) { dragSourceIndex = index; };
window.handleDragOver = function(e) { e.preventDefault(); };
window.handleDrop = function(e, targetIndex) {
  e.preventDefault();
  if (dragSourceIndex !== null && dragSourceIndex !== targetIndex) {
    const temp = seatAssignments[dragSourceIndex].student;
    seatAssignments[dragSourceIndex].student = seatAssignments[targetIndex].student;
    seatAssignments[targetIndex].student = temp;
    dragSourceIndex = null;
    renderPreview();
  }
};

// ================= イマーシブ・全画面モード制御 =================
let currentZoom = 1.0;

function toggleImmersiveMode() {
  document.body.classList.toggle('immersive-mode');
  
  if (document.body.classList.contains('immersive-mode')) {
    if (document.documentElement.requestFullscreen) {
      document.documentElement.requestFullscreen().catch(() => {});
    }
  } else {
    if (document.exitFullscreen && document.fullscreenElement) {
      document.exitFullscreen().catch(() => {});
    }
  }
}

document.getElementById('immersiveBtn')?.addEventListener('click', toggleImmersiveMode);
document.getElementById('floatingImmersiveExitBtn')?.addEventListener('click', toggleImmersiveMode);

document.getElementById('toggleControlsBtn')?.addEventListener('click', () => {
  const panel = document.getElementById('controlsContainer');
  if (panel) {
    panel.style.display = (panel.style.display === 'none') ? 'block' : 'none';
  }
});

window.toggleSection = function(id) {
  const el = document.getElementById(id);
  if (el) {
    el.style.display = (el.style.display === 'none') ? 'flex' : 'none';
  }
};

function applyZoom(delta) {
  if (delta === 0) currentZoom = 1.0;
  else currentZoom = Math.min(2.0, Math.max(0.5, currentZoom + delta));

  const wrapper = document.getElementById('previewZoomWrapper');
  const label = document.getElementById('zoomLevelLabel');
  if (wrapper) wrapper.style.transform = `scale(${currentZoom})`;
  if (label) label.innerText = `${Math.round(currentZoom * 100)}%`;
}

document.getElementById('btnZoomIn')?.addEventListener('click', () => applyZoom(0.1));
document.getElementById('btnZoomOut')?.addEventListener('click', () => applyZoom(-0.1));
document.getElementById('btnZoomReset')?.addEventListener('click', () => applyZoom(0));

window.addEventListener('DOMContentLoaded', () => {
  initPresets();
  document.getElementById('rollcallWeekDate').value = currentRollCallDate;
  updateDiaryHistorySelect();
});
