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


// 茨城県庁（https://www.pref.ibaraki.jp/towns/）公式掲載順マスター（44市町村）
const IBARAKI_MUNICIPALITIES_ORDER = [
  // 県北地域
  "日立市", "常陸太田市", "高萩市", "北茨城市", "常陸大宮市", "那珂市", "大子町", "東海村",
  // 県央地域
  "水戸市", "笠間市", "ひたちなか市", "茨城町", "大洗町", "城里町", "小美玉市",
  // 鹿行地域
  "鹿嶋市", "潮来市", "神栖市", "行方市", "鉾田市",
  // 県南地域
  "土浦市", "石岡市", "龍ケ崎市", "取手市", "牛久市", "つくば市", "美浦村", "阿見町", "河内町", "かすみがうら市", "稲敷市", "つくばみらい市", "守谷市", "利根町",
  // 県西地域
  "古河市", "結城市", "下妻市", "常総市", "筑西市", "坂東市", "桜川市", "八千代町", "五霞町", "境町"
];

// JIS都道府県コード順（茨城県を先頭に固定）
const PREFECTURES_JIS_ORDER = [
  "茨城県",
  "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
  "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県", "新潟県", "富山県",
  "石川県", "福井県", "山梨県", "長野県", "岐阜県", "静岡県", "愛知県", "三重県",
  "滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山県", "鳥取県", "島根県",
  "岡山県", "広島県", "山口県", "徳島県", "香川県", "愛媛県", "高知県", "福岡県",
  "佐賀県", "長崎県", "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県"
];


function downloadBomCsv(filename, content) {
  const bom = new Uint8Array([0xEF, 0xBB, 0xBF]);
  const blob = new Blob([bom, content], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(a.href);
}

const SUPER_ADMIN_EMAIL = "mitsuru.toshima@hitachi-tech.ac.jp";
let photoScale = 1.0;
let photoOffsetY = 15;
let adminTeachers = new Set([SUPER_ADMIN_EMAIL.toLowerCase()]);

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

const CLASS_TEACHERS = {
  "1-1": "遠島　充",
  "1-2": "小室　尚子",
  "1-3": "竹井　勇佑",
  "1E": "遠島　充",
  "1M": "小室　尚子",
  "1W": "竹井　勇佑",
  "2E": "佐久間　健",
  "2M": "工藤　雄大",
  "2W": "川崎　清勝",
  "3E1": "齊藤　慎",
  "3E2": "目良　優也",
  "3M": "金澤　章",
  "3W": "金澤　章",
  "3MW": "金澤　章"
};

const CLASS_DISPLAY_NAMES = {
  "1-1": "１年総合科１組 (1-1)",
  "1-2": "１年総合科２組 (1-2)",
  "1-3": "１年総合科３組 (1-3)",
  "1E": "１年電気科 (1E)",
  "1M": "１年機械科 (1M)",
  "1W": "１年溶接科 (1W)",
  "2E": "２年電気科 (2E)",
  "2M": "２年機械科 (2M)",
  "2W": "２年溶接科 (2W)",
  "3E1": "３年電気科１組 (3E1)",
  "3E2": "３年電気科２組 (3E2)",
  "3M": "３年機械科 (3M)",
  "3W": "３年溶接科 (3W)",
  "3MW": "３年機械・溶接科合同 (3MW)"
};

const PRESET_CLASSES = ["1-1", "1-2", "1-3", "2E", "2M", "2W", "3E1", "3E2", "3MW", "3M", "3W"];

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

// ================= クラウド（Firestore）自動保存＆自動復元 =================
async function saveRosterToCloud() {
  if (!db || students.length === 0) return;
  try {
    const docRef = doc(db, 'system_data', 'roster_master');
    await setDoc(docRef, {
      students: students,
      seatAssignments: seatAssignments,
      updatedAt: new Date().toISOString()
    });
    console.log("生徒名簿および座席データをクラウドに自動保存しました。");
  } catch (err) {
    console.warn("クラウド自動保存エラー:", err);
  }
}

async function loadRosterFromCloud() {
  if (!db) return;
  try {
    const docRef = doc(db, 'system_data', 'roster_master');
    const snap = await getDoc(docRef);
    if (snap.exists()) {
      const data = snap.data();
      if (data.students && data.students.length > 0) {
        students = data.students;
        if (data.seatAssignments && data.seatAssignments.length > 0) {
          seatAssignments = data.seatAssignments;
        } else {
          initSeatAssignments();
        }
        const badge = document.getElementById('rosterCount');
        if (badge) {
          badge.innerText = `生徒: ${students.length}名 (保存済)`;
          badge.classList.add('badge-ok');
        }
        updateFilterSelectors();
        renderPreview();
        console.log("クラウドから名簿データを復元しました（生徒: " + students.length + "名）");
      }
    }
  } catch (err) {
    console.warn("クラウド復元エラー:", err);
  }
}

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
    isolatedStudents: [
      { studentId: "", location: "隔離", temp: "37.2", reason: "足首捻挫のため、松葉杖使用。経過観察。" }
    ]
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
    officeRoom: "3-401",
    officeDuty: "（机）（橋本）（興野）",
    emergencyStay: "無し",
    outCount: 20,
    outReport: "特記事項なし。時間通り全員帰寮完了。",
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
        students = parseFlexibleRoster(res.data);
        onRosterLoaded();
      }
    });
  } else {
    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(new Uint8Array(evt.target.result), { type: 'array' });
      const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
      students = parseFlexibleRoster(rawRows);
      onRosterLoaded();
    };
    reader.readAsArrayBuffer(file);
  }
});


// ================= ヘボン式ローマ字変換（英字フォールバック用） =================

function normalizeClassName(raw) {
  if (!raw) return "1-1";
  let s = String(raw).trim().toUpperCase();
  if (s === "1G1" || s === "1-1" || s === "1E") return "1-1";
  if (s === "1G2" || s === "1-2" || s === "1M") return "1-2";
  if (s === "1G3" || s === "1-3" || s === "1W") return "1-3";
  if (s === "2E") return "2E";
  if (s === "2M") return "2M";
  if (s === "2W") return "2W";
  if (s.includes("3.00E+01") || s.includes("3E1") || s.includes("3.0E+01") || s === "3E-1") return "3E1";
  if (s.includes("3.00E+02") || s.includes("3E2") || s.includes("3.0E+02") || s === "3E-2") return "3E2";
  if (s === "3M") return "3M";
  if (s === "3W") return "3W";
  if (s === "3MW") return "3MW";
  return s;
}

function parseFlexibleRoster(rawRows) {
  if (!rawRows || rawRows.length === 0) return [];
  const cleanRows = rawRows.filter(r => r && r.length > 0 && r.some(c => String(c || "").trim() !== ""));
  
  // Find header row if exists
  let headerRowIdx = -1;
  const colIndexMap = {};

  for (let i = 0; i < Math.min(cleanRows.length, 10); i++) {
    const row = cleanRows[i].map(c => String(c || "").replace(/[\r\n]+/g, "").trim());
    if (row.some(c => c.includes("氏名") || c.includes("氏名番号") || c.includes("学籍番号") || c.includes("生徒番号"))) {
      headerRowIdx = i;
      row.forEach((h, colIdx) => {
        if (/氏名番号|学籍番号|生徒番号|コード/.test(h)) colIndexMap.code = colIdx;
        else if (h === "氏名" || (h.includes("氏名") && !h.includes("ふりがな") && !h.includes("ローマ字") && !h.includes("保護者"))) colIndexMap.name = colIdx;
        else if (/ふりがな|フリガナ|かな/.test(h) && !h.includes("保護者")) colIndexMap.kana = colIdx;
        else if (/クラス|組/.test(h)) colIndexMap.cls = colIdx;
        else if (/出席番号|番号/.test(h) && !h.includes("氏名番号") && !h.includes("携帯") && !h.includes("郵便")) colIndexMap.no = colIdx;
        else if (/部活/.test(h) && !h.includes("顧問")) colIndexMap.club = colIdx;
        else if (/役職|委員会/.test(h)) colIndexMap.committee = colIdx;
        else if (/ローマ字|英語|英字/.test(h) && !h.includes("名（ローマ字）")) colIndexMap.romaji = colIdx;
        else if (h.includes("名（ローマ字）") && colIndexMap.romajiFirst === undefined) colIndexMap.romajiFirst = colIdx;
        else if (/寮部屋|部屋/.test(h)) colIndexMap.room = colIdx;
        else if (/住所（統合）|住所|帰省先/.test(h)) colIndexMap.address = colIdx;
        else if (/出身中学校.*市町村|出身.*市町村/.test(h)) colIndexMap.jhCity = colIdx;
        else if (/出身.*学校名|出身中学|中学名/.test(h)) colIndexMap.jhName = colIdx;
      });
      break;
    }
  }

  const resultList = [];
  const rowsToProcess = (headerRowIdx !== -1) ? cleanRows.slice(headerRowIdx + 1) : cleanRows;

  rowsToProcess.forEach((r, idx) => {
    // Check if row has valid student code or name
    let code = colIndexMap.code !== undefined ? String(r[colIndexMap.code] || "").trim() : "";
    if (!code) {
      for (let cell of r) {
        const val = String(cell || "").trim();
        if (/^202[4-6]\d{3}$/.test(val)) { code = val; break; }
      }
    }
    if (!code) code = `2026${String(idx + 1).padStart(3, '0')}`;

    let name = colIndexMap.name !== undefined ? String(r[colIndexMap.name] || "").trim() : "";
    if (!name) {
      for (let cell of r) {
        const val = String(cell || "").trim();
        if (val.length >= 2 && !/^\d+$/.test(val) && !val.includes("@") && !val.includes("組") && !val.includes("科")) {
          name = val;
          break;
        }
      }
    }
    if (!name || name === "氏名" || name === "氏名番号") return;

    let kana = colIndexMap.kana !== undefined ? String(r[colIndexMap.kana] || "").trim() : "";
    let rawCls = colIndexMap.cls !== undefined ? String(r[colIndexMap.cls] || "").trim() : "";
    let cls = normalizeClassName(rawCls);

    let noVal = colIndexMap.no !== undefined ? parseInt(r[colIndexMap.no]) : (idx + 1);
    if (isNaN(noVal) || noVal <= 0) noVal = idx + 1;

    let club = colIndexMap.club !== undefined ? getFullClubName(String(r[colIndexMap.club] || "").trim()) : "-";
    let committee = colIndexMap.committee !== undefined ? String(r[colIndexMap.committee] || "").trim() : "-";
    let address = colIndexMap.address !== undefined ? String(r[colIndexMap.address] || "").trim() : "";
    let room = colIndexMap.room !== undefined ? String(r[colIndexMap.room] || "").trim() : "-";

    // Combine junior high school name
    let jhCity = colIndexMap.jhCity !== undefined ? String(r[colIndexMap.jhCity] || "").trim() : "";
    let jhName = colIndexMap.jhName !== undefined ? String(r[colIndexMap.jhName] || "").trim() : "";
    let juniorHigh = "";
    if (jhCity && jhName) {
      juniorHigh = `${jhCity}${jhName.endsWith("中") || jhName.endsWith("中学校") || jhName.endsWith("校") ? jhName : jhName + "中"}`;
    } else if (jhName) {
      juniorHigh = jhName;
    } else if (jhCity) {
      juniorHigh = jhCity;
    }

    // Romaji
    let romaji = colIndexMap.romaji !== undefined ? String(r[colIndexMap.romaji] || "").trim() : "";
    if (!romaji && kana) {
      romaji = kanaToRomaji(kana);
    }

    resultList.push({
      id: `${cls}_${String(noVal).padStart(2, '0')}`,
      studentCode: code,
      name: name.replace(/[\s　]+/g, " "),
      cleanName: name.replace(/[\s　]+/g, ""),
      kana: kana,
      romaji: romaji,
      cls: cls,
      no: noVal,
      club: club,
      committee: committee,
      address: address,
      juniorHigh: juniorHigh,
      room: room
    });
  });

  return resultList.length > 0 ? resultList : parseStrict209Students(rawRows);
}


function kanaToRomaji(kana) {
  if (!kana) return '';
  const map = {
    'あ': 'a', 'い': 'i', 'う': 'u', 'え': 'e', 'お': 'o',
    'か': 'ka', 'き': 'ki', 'く': 'ku', 'け': 'ke', 'こ': 'ko',
    'さ': 'sa', 'し': 'shi', 'す': 'su', 'せ': 'se', 'そ': 'so',
    'た': 'ta', 'ち': 'chi', 'つ': 'tsu', 'て': 'te', 'と': 'to',
    'な': 'na', 'に': 'ni', 'ぬ': 'nu', 'ね': 'ne', 'の': 'no',
    'は': 'ha', 'ひ': 'hi', 'ふ': 'fu', 'へ': 'he', 'ほ': 'ho',
    'ま': 'ma', 'み': 'mi', 'む': 'mu', 'め': 'me', 'も': 'mo',
    'や': 'ya', 'ゆ': 'yu', 'よ': 'yo',
    'ら': 'ra', 'り': 'ri', 'る': 'ru', 'れ': 're', 'ろ': 'ro',
    'わ': 'wa', 'を': 'o', 'ん': 'n',
    'が': 'ga', 'ぎ': 'gi', 'ぐ': 'gu', 'げ': 'ge', 'ご': 'go',
    'ざ': 'za', 'じ': 'ji', 'ず': 'zu', 'ぜ': 'ze', 'ぞ': 'zo',
    'だ': 'da', 'ぢ': 'ji', 'づ': 'zu', 'で': 'de', 'ど': 'do',
    'ば': 'ba', 'び': 'bi', 'ぶ': 'bu', 'べ': 'be', 'ぼ': 'bo',
    'ぱ': 'pa', 'ぴ': 'pi', 'ぷ': 'pu', 'ぺ': 'pe', 'ぽ': 'po',
    'きゃ': 'kya', 'きゅ': 'kyu', 'きょ': 'kyo',
    'しゃ': 'sha', 'しゅ': 'shu', 'しょ': 'sho',
    'ちゃ': 'cha', 'ちゅ': 'chu', 'ちょ': 'cho',
    'にゃ': 'nya', 'にゅ': 'nyu', 'にょ': 'nyo',
    'ひゃ': 'hya', 'ひゅ': 'hyu', 'ひょ': 'hyo',
    'みゃ': 'mya', 'みゅ': 'myu', 'みょ': 'myo',
    'りゃ': 'rya', 'りゅ': 'ryu', 'りょ': 'ryo',
    'ぎゃ': 'gya', 'ぎゅ': 'gyu', 'ぎょ': 'gyo',
    'じゃ': 'ja', 'じゅ': 'ju', 'じょ': 'jo',
    'びゃ': 'bya', 'びゅ': 'byu', 'びょ': 'byo',
    'ぴゃ': 'pya', 'ぴゅ': 'pyu', 'ぴょ': 'pyo'
  };

  const parts = kana.trim().split(/[\s　]+/);
  return parts.map(part => {
    let res = '';
    let i = 0;
    while (i < part.length) {
      if (i + 1 < part.length && map[part.slice(i, i + 2)]) {
        res += map[part.slice(i, i + 2)];
        i += 2;
      } else if (part[i] === 'っ' && i + 1 < part.length) {
        const next = map[part.slice(i + 1, i + 3)] || map[part[i + 1]] || '';
        if (next) res += next[0];
        i += 1;
      } else if (map[part[i]]) {
        res += map[part[i]];
        i += 1;
      } else if (part[i] === 'ー') {
        i += 1;
      } else {
        res += part[i];
        i += 1;
      }
    }
    return res.charAt(0).toUpperCase() + res.slice(1);
  }).join('\n');
}

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
        romaji: "",
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

    // 表1形式（氏名、ふりがな、英字ローマ字の取得）
    if (r.length >= 12 && String(r[0] || "").trim() === code) {
      const nVal = String(r[1] || "").trim();
      if (nVal && !item.name) {
        item.name = nVal.replace(/[\s　]+/g, " ");
        item.cleanName = nVal.replace(/[\s　]+/g, "");
      }
      if (!item.kana && r[2]) item.kana = String(r[2]).trim();
      if (!item.juniorHigh && r[8]) item.juniorHigh = String(r[8]).trim();
      if (!item.address && r[11]) item.address = String(r[11]).trim();

      // ローマ字（姓・名 2段形式: 例 "Ito\nMasato"）
      const lastEn = String(r[5] || "").trim();
      const firstEn = String(r[6] || "").trim();
      if (lastEn && firstEn) {
        item.romaji = `${lastEn}\n${firstEn}`;
      } else if (r[4]) {
        const parts = String(r[4]).trim().split(/\s+/);
        item.romaji = parts.length >= 2 ? `${parts[1]}\n${parts[0]}` : parts[0];
      }
    }

    // 表2形式（クラス、出席番号、部活）
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

    // 表3形式（大名簿からのローマ字・部屋番号補完）
    if (r.length > 55 && String(r[27] || "").trim() === code) {
      if (!item.romaji) {
        const lastEn = String(r[55] || "").trim();
        const firstEn = String(r[53] || "").trim();
        if (lastEn && firstEn) item.romaji = `${lastEn}\n${firstEn}`;
      }
      const rm = String(r[92] || "").trim();
      if (!item.room && rm && rm !== "-") item.room = rm;
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
    // ローマ字が空の場合はふりがなからヘボン式ローマ字を生成（ひらがな表示防止）
    if (!item.romaji && item.kana) {
      item.romaji = kanaToRomaji(item.kana);
    }
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
  saveRosterToCloud();
}

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
    if (targetCls === '3MW') {
      const list3M = students.filter(s => (s.cls || '').trim().toUpperCase() === '3M').sort((a, b) => a.no - b.no);
      const list3W = students.filter(s => (s.cls || '').trim().toUpperCase() === '3W').sort((a, b) => a.no - b.no);
      return [...list3M, ...list3W].map((s, idx) => ({
        ...s,
        displayNo: idx + 1
      }));
    }
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
    const originCtl = document.getElementById('originControls');
    if (originCtl) originCtl.style.display = (currentTab === 'origin') ? 'flex' : 'none';

    const genControls = document.getElementById('generalControls');
    const gradeBox = document.getElementById('gradeFilterBox');
    const classBox = document.getElementById('classFilterBox');
    const sizeBox = document.getElementById('sizeSelectorBox');
    const sizeSel = document.getElementById('sizeSelector');
    const photoAdjustBox = document.getElementById('photoAdjustBox');

    if (currentTab === 'photo') {
      genControls.style.display = 'flex';
      gradeBox.style.display = 'flex';
      classBox.style.display = 'flex';
      sizeBox.style.display = 'flex';
      if (photoAdjustBox) photoAdjustBox.style.display = 'flex';
      sizeSel.innerHTML = PHOTO_SIZES.map((s, idx) => `<option value="${idx}">${s.name}</option>`).join('');
    } else if (currentTab === 'plate_locker') {
      genControls.style.display = 'flex';
      gradeBox.style.display = 'flex';
      classBox.style.display = 'flex';
      sizeBox.style.display = 'flex';
      if (photoAdjustBox) photoAdjustBox.style.display = 'none';
      sizeSel.innerHTML = LOCKER_PLATE_SIZES.map((s, idx) => `<option value="${idx}">${s.name}</option>`).join('');
    } else if (currentTab === 'attendance' || currentTab === 'plate_dorm') {
      genControls.style.display = 'flex';
      gradeBox.style.display = 'flex';
      classBox.style.display = 'flex';
      sizeBox.style.display = 'none';
      if (photoAdjustBox) photoAdjustBox.style.display = 'none';
    } else {
      genControls.style.display = 'none';
      if (photoAdjustBox) photoAdjustBox.style.display = 'none';
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
    container.innerHTML = `<div class="empty-state"><p>生徒名簿ファイル（CSV / Excel）を上部からアップロードしてください。</p></div>`;
    return;
  }

  if (['plate_locker', 'plate_dorm', 'photo', 'attendance', 'origin'].includes(currentTab)) {
    document.body.classList.add('multi-page-mode');
  } else {
    document.body.classList.remove('multi-page-mode');
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

// 1. 座席表（原本PDF完全準拠：写真3:4比率、縦書き漢字氏名、横一列英字氏名＋出席番号、フッター）
function renderSeatingChart(container) {
  const roomName = document.getElementById('classroomPresetSelect')?.value || "座席表";
  const roomObj = PRESET_CLASSROOMS.find(r => r.name === roomName);

  if (roomObj && roomObj.isGroupStudy) {
    renderGroupStudySeating(container, roomName);
    return;
  }

  const cols = parseInt(document.getElementById('seatCols')?.value) || 5;
  const currentClass = document.getElementById('seatClassSelect')?.value || "1-1";
  const classDisplayName = CLASS_DISPLAY_NAMES[currentClass] || currentClass;
  const teacherName = CLASS_TEACHERS[currentClass] || "　　　";

  let html = `
    <div class="pdf-style-seating-wrapper a4-fit-container">
      <div class="pdf-header-bar">
        <div class="pdf-header-year">2026年度</div>
        <div class="pdf-header-title">座席表</div>
      </div>
      <div class="pdf-seating-grid" style="grid-template-columns: repeat(${cols}, 1fr);">
  `;

  seatAssignments.forEach((slotData, idx) => {
    const s = slotData.student;
    const isSelected = selectedSlotForSwap === idx ? "selected-swap" : "";

    if (!s) {
      html += `<div class="pdf-seat-card empty-seat" style="visibility:hidden; border:none;"></div>`;
    } else {
      const photo = getPhoto(s);
      const romaji = (s.romaji || kanaToRomaji(s.kana)).replace(/[\n\r\s　]+/g, ' ').trim();

      html += `
        <div class="pdf-seat-card ${isSelected}" 
             draggable="true" 
             onclick="handleSeatClick(${idx})" 
             ondragstart="handleDragStart(event, ${idx})" 
             ondragover="handleDragOver(event)" 
             drop="handleDrop(event, ${idx})">
          
          <div class="pdf-seat-top">
            <div class="pdf-seat-photo-box">
              ${photo ? `<img class="pdf-seat-photo" src="${photo}">` : `<div class="pdf-no-photo">写真</div>`}
            </div>
            <div class="pdf-seat-vertical-name">
              <span>${s.name}</span>
            </div>
          </div>

          <div class="pdf-seat-bottom">
            <div class="pdf-seat-romaji">
              <span>${romaji}</span>
            </div>
            <div class="pdf-seat-number">
              ${s.displayNo !== undefined ? s.displayNo : s.no}
            </div>
          </div>

        </div>
      `;
    }
  });

  html += `
      </div>
      <div class="pdf-footer-bar">
        <div class="pdf-footer-class">${classDisplayName}</div>
        <div class="pdf-podium-box">教　　卓</div>
        <div class="pdf-footer-teacher">担任: ${teacherName}</div>
      </div>
    </div>
  `;
  container.innerHTML = html;
}

function renderGroupStudySeating(container, roomName) {
  let html = `
    <div class="pdf-style-seating-wrapper a4-fit-container">
      <div class="pdf-header-bar">
        <div class="pdf-header-year">2026年度</div>
        <div class="pdf-header-title">座席表</div>
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
      <div class="pdf-footer-bar">
        <div class="pdf-footer-class">特別教室（グループ学習）</div>
        <div class="pdf-podium-box">教　　卓</div>
        <div class="pdf-footer-teacher"></div>
      </div>
    </div>
  `;
  container.innerHTML = html;
}

// 2. 生徒名簿（表題「生徒名簿」、クラス名・担任名を併記、長い中学校名は縮小）
function renderAttendanceSheet(container) {
  const targetClass = document.getElementById('generalClassFilter')?.value || "1-1";
  const classDisplayName = CLASS_DISPLAY_NAMES[targetClass] || (targetClass === "ALL" ? "全校生徒" : targetClass);
  const teacherName = CLASS_TEACHERS[targetClass] || "　　　";
  const list = getFilteredStudents();

  let html = `
    <div class="attendance-sheet-wrapper">
      <div class="attendance-header" style="display:flex; justify-content:space-between; align-items:baseline; margin-bottom:8px; border-bottom:2px solid #333; padding-bottom:4px;">
        <div style="display:flex; align-items:baseline; gap:16px;">
          <div style="font-size:16pt; font-weight:bold;">生徒名簿</div>
          <div style="font-size:12pt; font-weight:bold; color:#1e3a8a;">${classDisplayName}</div>
          <div style="font-size:11pt;">担任: <strong>${teacherName}</strong></div>
        </div>
        <div style="font-size:10pt;">対象人数: ${list.length}名</div>
      </div>
      <table class="attendance-table">
        <thead>
          <tr>
            <th style="width:40px;">番号</th>
            <th style="width:110px;">氏名</th>
            <th style="width:120px;">ふりがな</th>
            <th style="width:130px;">出身中学</th>
            <th>帰省先住所</th>
            <th style="width:90px;">部活動</th>
            <th style="width:100px;">委員会</th>
          </tr>
        </thead>
        <tbody>
  `;

  list.forEach(s => {
    const isLong = (s.juniorHigh || '').length >= 9;
    html += `
      <tr>
        <td>${s.displayNo !== undefined ? s.displayNo : s.no}</td>
        <td style="font-weight:bold; font-size:10pt;">${s.name}</td>
        <td style="font-size:8pt;">${s.kana || '-'}</td>
        <td class="${isLong ? 'school-long-text' : ''}">${s.juniorHigh || '-'}</td>
        <td style="text-align:left; font-size:8.5pt;">${getShortAddress(s.address)}</td>
        <td>${s.club || '-'}</td>
        <td>${s.committee || '-'}</td>
      </tr>
    `;
  });

  html += `</tbody></table></div>`;
  container.innerHTML = html;
}

// 3. 寮生点呼簿（1-301, 1-302の管理棟非表示、ダッシュボードに部活名追加）
function renderRollCallSystem(container) {
  const filter = document.getElementById('dormFilterSelect').value;

  let list = students.filter(s => {
    if (!s.room || s.room === '-') return false;
    const rm = s.room.trim();
    if (rm === '1-301' || rm === '1-302' || rm.includes('管理')) return false;
    return true;
  });
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
        <strong>${a.student.name}</strong> (${a.student.cls} / ${a.student.room}) 
        <span style="color:#0284c7; font-weight:bold;">[${a.student.club || '部活無'}]</span> - 
        <span style="color:#d32f2f; font-weight:bold;">${a.status}</span> ${a.note ? `[${a.note}]` : ''}
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

// 4. ロッカー用ネームプレート
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

// 5. ネームプレート（寮 4×3cm）
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

// 6. 証明写真印刷（微調整スライダー連動）
function renderPhotoSheet(container) {
  const spec = PHOTO_SIZES[document.getElementById('sizeSelector').value || 0];
  const list = getFilteredStudents();

  let html = `<div class="photo-sheet">`;
  list.forEach(s => {
    const photo = getPhoto(s, true);
    html += `
      <div class="photo-item-container">
        <div class="photo-card" style="width:${spec.w}mm; height:${spec.h}mm; overflow:hidden;">
          ${photo ? `<img src="${photo}" style="transform: scale(${photoScale}); object-position: center ${photoOffsetY}%;">` : `<div class="no-photo">写真なし</div>`}
        </div>
        <div class="photo-external-label" style="width:${spec.w}mm;">${s.cls}-${s.displayNo !== undefined ? s.displayNo : s.no} ${s.name}</div>
      </div>
    `;
  });
  html += `</div>`;
  container.innerHTML = html;
}

// 7. 各種日誌（日直・宿直・保健室）
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

function renderStudentOptionList(selectedId = "") {
  return `<option value="">-- 生徒を選択 --</option>` +
    students.map(s => `<option value="${s.id}" ${s.id === selectedId ? 'selected' : ''}>${s.name} (${s.cls} / ${s.room || '寮室無'} / ${s.club || '部活無'})</option>`).join('');
}

function getStudentsInRoom(roomNo) {
  if (!roomNo) return "生徒なし";
  const matched = students.filter(s => s.room === roomNo).map(s => s.name);
  return matched.length > 0 ? matched.join('・') : "該当生徒なし";
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
  const period = document.getElementById('diarySummaryPeriod')?.value || 'single';

  if (period !== 'single') {
    renderDiarySummary(container, period);
    return;
  }

  if (currentDiaryType === 'day') {
    const d = currentDiary.day;
    const weekday = getWeekday(d.date);

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
                <option value="休日" ${d.schoolType==='休日'?'selected':''}>休日</option>
                <option value="登校日" ${d.schoolType==='登校日'?'selected':''}>登校日</option>
                <option value="寮閉鎖" ${d.schoolType==='寮閉鎖'?'selected':''}>寮閉鎖</option>
              </select>
            </td>
          </tr>
          <tr>
            <th>寮内巡回</th>
            <td colspan="3">
              <div style="display:flex; gap:6px; flex-wrap:wrap;">
                ${[0,1,2,3,4].map(i => `
                  <span>${['①','②','③','④','⑤'][i]} <input type="text" value="${d.dormPatrols[i]}" onchange="currentDiary.day.dormPatrols[${i}]=this.value" style="width:100px;" placeholder="例: 8:45～9:45"></span>
                `).join('')}
              </div>
            </td>
          </tr>
          <tr>
            <th>校内巡回</th>
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
            <th>隔離・療養生徒</th>
            <td colspan="3">
              ${(d.isolatedStudents || []).map((item, idx) => {
                const sel = students.find(s => s.id === item.studentId);
                return `
                  <div style="margin-bottom:8px; border-bottom:1px dashed #cbd5e1; padding-bottom:6px;">
                    <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap;">
                      生徒: <select onchange="currentDiary.day.isolatedStudents[${idx}].studentId=this.value; renderPreview();">${renderStudentOptionList(item.studentId)}</select>
                      所在: <select onchange="currentDiary.day.isolatedStudents[${idx}].location=this.value;">
                        <option value="自室" ${item.location==='自室'?'selected':''}>自室</option>
                        <option value="保健室" ${item.location==='保健室'?'selected':''}>保健室</option>
                        <option value="隔離" ${item.location==='隔離'?'selected':''}>隔離</option>
                        <option value="自宅" ${item.location==='自宅'?'selected':''}>自宅</option>
                      </select>
                      体温: <input type="text" value="${item.temp}" onchange="currentDiary.day.isolatedStudents[${idx}].temp=this.value;" style="width:50px;"> ℃
                      理由・病状: <input type="text" value="${item.reason}" onchange="currentDiary.day.isolatedStudents[${idx}].reason=this.value;" style="width:40%;">
                      <button class="btn btn-sm btn-danger-outline" onclick="currentDiary.day.isolatedStudents.splice(${idx}, 1); renderPreview();">削除</button>
                    </div>
                    ${sel ? `
                      <div style="background:#f8fafc; padding:4px 8px; border-radius:4px; margin-top:4px; font-size:9pt; color:#475569;">
                        <strong>${sel.name}</strong> ｜ クラス: ${sel.cls} ｜ 部屋: ${sel.room || '-'} ｜ 部活: ${sel.club || '-'} ｜ 住所: ${getPrefCity(sel.address)}
                      </div>
                    ` : ''}
                  </div>
                `;
              }).join('')}
              <button class="btn btn-sm btn-light" onclick="currentDiary.day.isolatedStudents.push({studentId:'', location:'隔離', temp:'', reason:''}); renderPreview();">＋ 隔離生徒を追加</button>
            </td>
          </tr>
        </table>
      </div>
    `;
  } else if (currentDiaryType === 'night') {
    const d = currentDiary.night;
    const weekday = getWeekday(d.date);

    // 部屋リスト（ユニーク）
    const roomList = Array.from(new Set(students.map(s => s.room).filter(r => r && r !== '-'))).sort();

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
            <td>
              <select onchange="currentDiary.night.arriveTime=this.value" style="font-size:11px;">
                ${["15:45", "16:00", "16:15", "16:20", "16:30", "16:45", "17:00"].map(t => `<option value="${t}" ${d.arriveTime===t?'selected':''}>${t}</option>`).join('')}
              </select>
            </td>
            <th>最終退勤/最初出勤</th>
            <td>退勤: <input type="text" value="${d.lastLeave}" onchange="currentDiary.night.lastLeave=this.value" style="width:90px;"> / 出勤: <input type="text" value="${d.firstArrive}" onchange="currentDiary.night.firstArrive=this.value" style="width:90px;"></td>
          </tr>
          <tr>
            <th>点呼確認</th>
            <td colspan="3"><input type="text" value="${d.rollcallStatus}" onchange="currentDiary.night.rollcallStatus=this.value" style="width:90%;"></td>
          </tr>
          <tr>
            <th>寮内巡回</th>
            <td colspan="3">
              <div style="display:flex; gap:8px; flex-wrap:wrap;">
                ${[0,1,2,3,4].map(i => `
                  <span>${['①','②','③','④','⑤'][i]} 
                    <select onchange="currentDiary.night.patrols[${i}]=this.value" style="font-size:11px;">
                      <option value="">--</option>
                      ${["20:00～20:30", "21:00～21:30", "21:45～22:10", "22:30～23:00", "23:30～24:00", "6:00～6:30"].map(t => `<option value="${t}" ${d.patrols[i]===t?'selected':''}>${t}</option>`).join('')}
                    </select>
                  </span>
                `).join('')}
              </div>
            </td>
          </tr>
          <tr>
            <th>事務当番 / 外出報告</th>
            <td colspan="3">
              <div style="display:flex; flex-direction:column; gap:6px;">
                <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                  <span>事務当番（部屋選択）:</span>
                  <select onchange="currentDiary.night.officeRoom=this.value; renderPreview();">
                    <option value="">-- 部屋選択 --</option>
                    ${roomList.map(r => `<option value="${r}" ${d.officeRoom===r?'selected':''}>${r}</option>`).join('')}
                  </select>
                  <span>担当生徒: <strong>${getStudentsInRoom(d.officeRoom)}</strong></span>
                  <input type="text" value="${d.officeDuty || ''}" onchange="currentDiary.night.officeDuty=this.value" placeholder="特記事項（机当番等）" style="width:140px;">
                </div>
                <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap; margin-top:4px;">
                  <span>外出報告:</span>
                  <input type="number" value="${d.outCount || 0}" style="width:55px;" onchange="currentDiary.night.outCount=this.value">
                  <span>名の申告・帰寮報告あり</span>
                  <input type="text" value="${d.outReport || ''}" onchange="currentDiary.night.outReport=this.value" placeholder="外出に関する備考・特記" style="flex:1;">
                </div>
              </div>
            </td>
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
                    <select onchange="currentDiary.night.sickList[${idx}].studentId=this.value; renderPreview();">${renderStudentOptionList(item.studentId)}</select>
                    体温: <input type="text" value="${item.temp}" onchange="currentDiary.night.sickList[${idx}].temp=this.value" style="width:50px;"> ℃
                    ${sel ? `<span>[部屋: ${sel.room||'-'} / 部活: ${sel.club||'-'} / 住所: ${getPrefCity(sel.address)}]</span>` : ''}
                    <input type="text" value="${item.condition}" onchange="currentDiary.night.sickList[${idx}].condition=this.value" style="width:35%;" placeholder="症状・自宅療養/隔離">
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
                const badgeClass = item.type === '欠席' ? 'health-type-absent' : (item.type === '遅刻' ? 'health-type-late' : 'health-type-early');
                return `
                  <div class="health-entry-card">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:4px;">
                      <span class="health-type-badge ${badgeClass}">${item.type}</span>
                      <button class="btn btn-sm btn-danger-outline" onclick="currentDiary.health.absenceItems.splice(${idx}, 1); renderPreview();">削除</button>
                    </div>
                    <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap;">
                      区分: <select onchange="currentDiary.health.absenceItems[${idx}].type=this.value; renderPreview();"><option value="欠席" ${item.type==='欠席'?'selected':''}>欠席</option><option value="遅刻" ${item.type==='遅刻'?'selected':''}>遅刻</option><option value="早退" ${item.type==='早退'?'selected':''}>早退</option></select>
                      生徒: <select onchange="currentDiary.health.absenceItems[${idx}].studentId=this.value; renderPreview();">${renderStudentOptionList(item.studentId)}</select>
                      体温: <input type="text" value="${item.temp}" onchange="currentDiary.health.absenceItems[${idx}].temp=this.value" style="width:50px;"> ℃
                      症状・経過: <input type="text" value="${item.symptom}" onchange="currentDiary.health.absenceItems[${idx}].symptom=this.value" style="width:40%;">
                    </div>
                    ${sel ? `
                      <div style="font-size:9pt; color:#64748b; margin-top:4px;">
                        ${sel.cls} | 部屋: ${sel.room || '-'} | 部活: ${sel.club || '-'} | 出身: ${getPrefCity(sel.address)}
                      </div>
                    ` : ''}
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
                  <div class="health-entry-card">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:4px;">
                      <strong>来室 #${idx + 1}</strong>
                      <button class="btn btn-sm btn-danger-outline" onclick="currentDiary.health.visitorItems.splice(${idx}, 1); renderPreview();">削除</button>
                    </div>
                    <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap;">
                      生徒: <select onchange="currentDiary.health.visitorItems[${idx}].studentId=this.value; renderPreview();">${renderStudentOptionList(item.studentId)}</select>
                      体温: <input type="text" value="${item.temp}" onchange="currentDiary.health.visitorItems[${idx}].temp=this.value" style="width:50px;"> ℃
                      処置・理由: <input type="text" value="${item.reason}" onchange="currentDiary.health.visitorItems[${idx}].reason=this.value" style="width:45%;">
                    </div>
                    ${sel ? `
                      <div style="font-size:9pt; color:#64748b; margin-top:4px;">
                        ${sel.cls} | 部屋: ${sel.room || '-'} | 部活: ${sel.club || '-'}
                      </div>
                    ` : ''}
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

// 8. 出身地別名簿（委員会列追加、茨城県公式44市町村順最優先、他県JIS順、CSV出力＆印刷）
function renderOriginSheet(container) {
  const groups = {};
  students.forEach(s => {
    let reg = getPrefCity(s.address);
    if (!groups[reg]) groups[reg] = [];
    groups[reg].push(s);
  });

  const sortedKeys = Object.keys(groups).sort((a, b) => {
    const isAIbaraki = a.includes("茨城") || IBARAKI_MUNICIPALITIES_ORDER.some(m => a.includes(m));
    const isBIbaraki = b.includes("茨城") || IBARAKI_MUNICIPALITIES_ORDER.some(m => b.includes(m));

    if (isAIbaraki && !isBIbaraki) return -1;
    if (!isAIbaraki && isBIbaraki) return 1;

    if (isAIbaraki && isBIbaraki) {
      let idxA = IBARAKI_MUNICIPALITIES_ORDER.findIndex(m => a.includes(m));
      let idxB = IBARAKI_MUNICIPALITIES_ORDER.findIndex(m => b.includes(m));
      if (idxA === -1) idxA = 999;
      if (idxB === -1) idxB = 999;
      return idxA - idxB;
    }

    let prefA = PREFECTURES_JIS_ORDER.findIndex(p => a.startsWith(p));
    let prefB = PREFECTURES_JIS_ORDER.findIndex(p => b.startsWith(p));
    return prefA - prefB;
  });

  let html = `<div class="attendance-sheet-wrapper"><div class="attendance-header"><div style="font-size:16pt; font-weight:bold;">出身地別名簿</div><div>対象人数: ${students.length}名</div></div>`;
  sortedKeys.forEach(reg => {
    html += `
      <h4 style="margin:14px 0 4px; border-left:4px solid #1e3a8a; padding-left:8px;">${reg} (${groups[reg].length}名)</h4>
      <table class="attendance-table">
        <thead>
          <tr>
            <th style="width:60px;">組-番</th>
            <th style="width:110px;">氏名</th>
            <th>帰省先住所</th>
            <th style="width:90px;">部活動</th>
            <th style="width:100px;">委員会</th>
          </tr>
        </thead>
        <tbody>
          ${groups[reg].map(s => `<tr><td>${s.cls}-${s.displayNo !== undefined ? s.displayNo : s.no}</td><td style="font-weight:bold;">${s.name}</td><td style="text-align:left;">${getShortAddress(s.address)}</td><td>${s.club || '-'}</td><td>${s.committee || '-'}</td></tr>`).join('')}
        </tbody>
      </table>
    `;
  });
  html += `</div>`;
  container.innerHTML = html;
}

// 出身地別名簿 スプレッドシート用CSV出力
document.getElementById('btnExportOriginCSV')?.addEventListener('click', () => {
  const exportData = [];
  students.forEach(s => {
    exportData.push({
      "地域・市町村": getPrefCity(s.address),
      "クラス": s.cls,
      "出席番号": s.no,
      "氏名": s.name,
      "帰省先住所": s.address,
      "部活動": s.club,
      "委員会": s.committee
    });
  });
  const today = new Date().toISOString().slice(0, 10).replace(/-/g, "");
  downloadBomCsv(`出身地別名簿_スプレッドシート用_${today}.csv`, Papa.unparse(exportData));
});

// 写真微調整スライダー
document.getElementById('photoScaleSlider')?.addEventListener('input', (e) => {
  photoScale = parseFloat(e.target.value);
  document.getElementById('photoScaleVal').innerText = `${Math.round(photoScale * 100)}%`;
  renderPreview();
});

document.getElementById('photoOffsetYSlider')?.addEventListener('input', (e) => {
  photoOffsetY = parseInt(e.target.value);
  document.getElementById('photoOffsetYVal').innerText = `${photoOffsetY}%`;
  renderPreview();
});

function getFilteredStudents() {
  const grade = document.getElementById('generalGradeFilter').value;
  const cls = document.getElementById('generalClassFilter').value;

  return students.filter(s => {
    const matchGrade = (grade === 'ALL') || (s.cls && s.cls.startsWith(grade));
    const matchClass = (cls === 'ALL') || (s.cls === cls);
    return matchGrade && matchClass;
  });
}

// ================= 教員マスター & Google SSO アクセスガード =================
let allowedTeachers = new Set([SUPER_ADMIN_EMAIL.toLowerCase()]);

// 1. 教員名簿（CSV / Excel）の取り込み処理
document.getElementById('teacherFile')?.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const ext = file.name.split('.').pop().toLowerCase();
  const handleRows = async (rows) => {
    let count = 0;
    const teacherList = [];

    rows.forEach(r => {
      if (!r || r.length === 0) return;
      let email = null;
      let name = "";
      r.forEach((cell, idx) => {
        const val = String(cell || "").trim();
        if (/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(val)) {
          email = val.toLowerCase();
        } else if (!name && val.length >= 2 && !/^\d+$/.test(val) && idx < 5) {
          name = val;
        }
      });

      if (email) {
        allowedTeachers.add(email);
        teacherList.push({ email, name });
        count++;
      }
    });

    try {
      const batch = writeBatch(db);
      teacherList.forEach(t => {
        const docId = t.email.replace(/[.@]/g, '_');
        const refDoc = doc(db, 'teachers', docId);
        batch.set(refDoc, t);
      });
      await batch.commit();
    } catch (err) {
      console.warn("Firestore教員保存通知:", err);
    }

    localStorage.setItem('allowed_teachers', JSON.stringify(Array.from(allowedTeachers)));
    
    const badge = document.getElementById('teacherCount');
    if (badge) {
      badge.innerText = `教員: ${allowedTeachers.size}名 登録済`;
      badge.classList.add('badge-ok');
    }
    alert(`教員 ${count} 名のメールアドレスを登録・承認しました。`);
    checkUserAccess(auth.currentUser);
  };

  if (ext === 'csv') {
    Papa.parse(file, { header: false, complete: (res) => handleRows(res.data) });
  } else {
    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(new Uint8Array(evt.target.result), { type: 'array' });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
      handleRows(rows);
    };
    reader.readAsArrayBuffer(file);
  }
});

// 2. 登録済み教員リストの初期ロード
async function loadAllowedTeachers() {
  const cached = localStorage.getItem('allowed_teachers');
  if (cached) {
    try {
      JSON.parse(cached).forEach(em => allowedTeachers.add(em.toLowerCase()));
    } catch (e) {}
  }

  try {
    const snap = await getDocs(collection(db, 'teachers'));
    snap.forEach(d => {
      const data = d.data();
      if (data.email) allowedTeachers.add(data.email.toLowerCase());
    });
    localStorage.setItem('allowed_teachers', JSON.stringify(Array.from(allowedTeachers)));
  } catch (err) {
    console.log("Firestore教員一覧取得通知:", err);
  }

  const badge = document.getElementById('teacherCount');
  if (badge) {
    badge.innerText = `教員: ${allowedTeachers.size}名 登録済`;
    if (allowedTeachers.size > 1) badge.classList.add('badge-ok');
  }
}

// 3. SSOログイン状態 & 利用資格チェック
function checkUserAccess(user) {
  const overlay = document.getElementById('authLockOverlay');
  const lockMsg = document.getElementById('authLockMsg');
  const userStatus = document.getElementById('userStatus');
  const authBtn = document.getElementById('authBtn');

  if (!user) {
    if (overlay) overlay.style.display = 'flex';
    if (lockMsg) lockMsg.innerHTML = `本システムは登録された教員アカウント（Google SSO）のみご利用いただけます。<br>学校のGoogleアカウントでログインしてください。`;
    if (userStatus) userStatus.innerText = "未ログイン";
    if (authBtn) authBtn.innerText = "ログイン";
    return;
  }

  const userEmail = (user.email || "").toLowerCase();
  if (userStatus) userStatus.innerText = userEmail;
  if (authBtn) authBtn.innerText = "ログアウト";

  const isAdmin = (userEmail === SUPER_ADMIN_EMAIL.toLowerCase() || adminTeachers.has(userEmail));
  const adminBar = document.getElementById('adminBar');
  if (adminBar) adminBar.style.display = isAdmin ? 'block' : 'none';

  const assignBtn = document.getElementById('btnAssignAdmin');
  if (assignBtn) {
    assignBtn.style.display = (userEmail === SUPER_ADMIN_EMAIL.toLowerCase()) ? 'inline-flex' : 'none';
  }

  if (allowedTeachers.has(userEmail) || isAdmin) {
    if (overlay) overlay.style.display = 'none';
  } else {
    if (overlay) overlay.style.display = 'flex';
    if (lockMsg) lockMsg.innerHTML = `<span style="color:#ef4444; font-weight:bold;">【アクセス拒否】</span><br>アカウント <strong>${userEmail}</strong> は教員名簿に登録されていません。<br>管理者に登録を依頼するか、承認された教員アカウントでログインし直してください。`;
  }
}

// 4. 認証イベントの監視
onAuthStateChanged(auth, (user) => {
  checkUserAccess(user);
});

// ログイン・ログアウトボタン動作
document.getElementById('authBtn')?.addEventListener('click', () => {
  if (auth.currentUser) {
    signOut(auth).then(() => checkUserAccess(null));
  } else {
    signInWithPopup(auth, googleProvider).catch((e) => alert("ログインエラー: " + e.message));
  }
});

document.getElementById('lockLoginBtn')?.addEventListener('click', () => {
  signInWithPopup(auth, googleProvider).catch((e) => alert("ログインエラー: " + e.message));
});


// ================= 名簿クリア & 雛形/ユーザーリスト ダウンロード =================
document.getElementById('btnClearStudent')?.addEventListener('click', () => {
  if (!confirm("読み込まれている生徒名簿データをクリアしますか？\n※座席表や名簿の表示がリセットされます。")) return;
  students = [];
  seatAssignments = [];
  const fileInput = document.getElementById('rosterFile');
  if (fileInput) fileInput.value = '';
  const badge = document.getElementById('rosterCount');
  if (badge) {
    badge.innerText = '生徒: 0名';
    badge.classList.remove('badge-ok');
  }
  renderPreview();
  try {
    const docRef = doc(db, 'system_data', 'roster_master');
    setDoc(docRef, { students: [], seatAssignments: [], updatedAt: new Date().toISOString() });
  } catch (e) {}
  alert("生徒名簿データをクリアしました。");
});

document.getElementById('btnClearTeacher')?.addEventListener('click', () => {
  if (!confirm("登録済みの教員名簿をクリアしますか？\n※管理者アカウントのみ有効な状態に戻ります。")) return;
  allowedTeachers = new Set([SUPER_ADMIN_EMAIL.toLowerCase()]);
  localStorage.removeItem('allowed_teachers');
  const fileInput = document.getElementById('teacherFile');
  if (fileInput) fileInput.value = '';
  const badge = document.getElementById('teacherCount');
  if (badge) {
    badge.innerText = `教員: 1名 (管理者)`;
    badge.classList.remove('badge-ok');
  }
  alert("教員名簿をクリアしました（管理者のみ有効）。");
  checkUserAccess(auth.currentUser);
});

document.getElementById('btnDownloadStudentTpl')?.addEventListener('click', () => {
  const sample = [
    {
      "学籍番号": "2026001",
      "氏名": "日専 太郎",
      "ふりがな": "にっせん たろう",
      "英字氏名": "Nissen\nTaro",
      "クラス": "1-1",
      "出席番号": "1",
      "出身中学": "日立第一中学校",
      "帰省先住所": "茨城県日立市幸町1丁目",
      "部活動": "野球",
      "委員会": "体育委員",
      "部屋番号": "2-301"
    },
    {
      "学籍番号": "2026002",
      "氏名": "技術 花子",
      "ふりがな": "ぎじゅつ はなこ",
      "英字氏名": "Gijutsu\nHanako",
      "クラス": "1-1",
      "出席番号": "2",
      "出身中学": "助川中学校",
      "帰省先住所": "茨城県日立市助川町2丁目",
      "部活動": "吹奏楽",
      "委員会": "図書委員",
      "部屋番号": "1-102"
    }
  ];
  downloadBomCsv("生徒名簿_雛形.csv", Papa.unparse(sample));
});

document.getElementById('btnDownloadStudentList')?.addEventListener('click', () => {
  if (students.length === 0) {
    alert("現在読み込まれている生徒名簿データがありません。");
    return;
  }
  const exportData = students.map(s => ({
    "学籍番号": s.studentCode || "",
    "クラス": s.cls || "",
    "出席番号": s.no || "",
    "氏名": s.name || "",
    "ふりがな": s.kana || "",
    "英字氏名": (s.romaji || "").replace(/\n/g, " "),
    "出身中学": s.juniorHigh || "",
    "帰省先住所": s.address || "",
    "部活動": s.club || "",
    "委員会": s.committee || "",
    "部屋番号": s.room || ""
  }));
  const today = new Date().toISOString().slice(0, 10).replace(/-/g, "");
  downloadBomCsv(`生徒名簿_ユーザーリスト_${today}.csv`, Papa.unparse(exportData));
});

document.getElementById('btnDownloadTeacherTpl')?.addEventListener('click', () => {
  const sample = [
    { "氏名": "遠島 充", "メールアドレス": "mitsuru.toshima@hitachi-tech.ac.jp", "役職": "管理者" },
    { "氏名": "小室 尚子", "メールアドレス": "komuro@example.ac.jp", "役職": "教員" },
    { "氏名": "竹井 勇佑", "メールアドレス": "takei@example.ac.jp", "役職": "教員" }
  ];
  downloadBomCsv("教員名簿_雛形.csv", Papa.unparse(sample));
});

document.getElementById('btnDownloadTeacherList')?.addEventListener('click', () => {
  const exportData = Array.from(allowedTeachers).map(email => ({
    "メールアドレス": email,
    "種別": email === SUPER_ADMIN_EMAIL.toLowerCase() ? "管理者" : "教員"
  }));
  const today = new Date().toISOString().slice(0, 10).replace(/-/g, "");
  downloadBomCsv(`登録教員リスト_${today}.csv`, Papa.unparse(exportData));
});

document.getElementById('btnAssignAdmin')?.addEventListener('click', () => {
  const email = prompt("管理者に指名する教員のメールアドレスを入力してください：");
  if (!email) return;
  adminTeachers.add(email.trim().toLowerCase());
  localStorage.setItem('admin_teachers', JSON.stringify(Array.from(adminTeachers)));
  alert(`教員「${email}」を管理者に指名しました。`);
});



// ================= 表示拡大率 (ズーム) コントロール =================
let currentZoom = 1.0;

function applyZoom(delta) {
  if (delta === 0) currentZoom = 1.0;
  else currentZoom = Math.min(2.0, Math.max(0.5, currentZoom + delta));

  const wrapper = document.getElementById('previewZoomWrapper');
  const label = document.getElementById('zoomLevelLabel');
  if (wrapper) {
    wrapper.style.transform = `scale(${currentZoom})`;
    wrapper.style.transformOrigin = 'top center';
  }
  if (label) label.innerText = `${Math.round(currentZoom * 100)}%`;
}

document.getElementById('btnZoomIn')?.addEventListener('click', () => applyZoom(0.1));
document.getElementById('btnZoomOut')?.addEventListener('click', () => applyZoom(-0.1));
document.getElementById('btnZoomReset')?.addEventListener('click', () => applyZoom(0));


// 教員数バッジクリックで登録済み教員一覧を表示
document.getElementById('teacherCount')?.addEventListener('click', () => {
  const teacherArr = Array.from(allowedTeachers);
  if (teacherArr.length === 0) {
    alert("登録されている教員はありません。");
    return;
  }
  const lines = teacherArr.map((em, idx) => {
    const isSuper = em === SUPER_ADMIN_EMAIL.toLowerCase();
    const isAdmin = adminTeachers.has(em);
    const role = isSuper ? "【特権管理者】" : (isAdmin ? "【管理者】" : "【一般教員】");
    return `${idx + 1}. ${em} ${role}`;
  }).join('\n');
  alert(`【登録済み教員一覧（計 ${teacherArr.length} 名）】\n\n` + lines);
});


// ================= 写真データのローカルキャッシュ永続化 =================
function savePhotosToCache() {
  try {
    const keys = Object.keys(photoStore);
    if (keys.length === 0) return;
    localStorage.setItem('cached_photo_keys', JSON.stringify(keys.slice(0, 50)));
    // save each photo
    keys.slice(0, 50).forEach(k => {
      try { localStorage.setItem('photo_' + k, photoStore[k]); } catch(e){}
    });
  } catch(e) {
    console.warn("写真キャッシュ容量超過通知:", e);
  }
}

function loadPhotosFromCache() {
  try {
    const keys = JSON.parse(localStorage.getItem('cached_photo_keys') || '[]');
    let loaded = 0;
    keys.forEach(k => {
      const data = localStorage.getItem('photo_' + k);
      if (data) { photoStore[k] = data; loaded++; }
    });
    if (loaded > 0) {
      const pCount = document.getElementById('photoCount');
      if (pCount) {
        pCount.innerText = `写真: ${loaded}枚 (保存済)`;
        pCount.classList.add('badge-ok');
      }
    }
  } catch(e){}
}

function initializeApp() {
  initPresets();
  updateFilterSelectors();
  const rollcallDateEl = document.getElementById('rollcallWeekDate');
  if (rollcallDateEl) rollcallDateEl.value = currentRollCallDate;
  updateDiaryHistorySelect();
  loadAllowedTeachers();
  loadRosterFromCloud();
  loadPhotosFromCache();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initializeApp);
} else {
  initializeApp();
}
