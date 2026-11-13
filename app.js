/**
 * HỆ THỐNG THỐNG KÊ NỤ CƯỜI TTC - DASHBOARD & ANALYTICS
 * Quy tắc: Hình ảnh = 1 nụ cười | Video = 2 nụ cười
 * Phân cấp chuẩn: 99 Chi nhánh thuộc 15 Khu vực (theo file Chi nhánh.xlsx)
 */

const SPREADSHEET_ID = '1QQwYkkOkJWxR36KamV8jRSE98I8pLNRtkeXXe7MOcas';
const GID = '0';

// Danh mục chuẩn 15 Khu vực & 78 Chi nhánh (Đã bỏ tiền tố TIN / Phuong Nam và gộp các phòng đối tác -01, -02 của tỉnh về 1 Chi nhánh)
const REGION_BRANCH_HIERARCHY = [
  {
    region: 'HNI',
    regionName: 'Hà Nội (HNI)',
    branches: [
      'HNI-01', 'HNI-02', 'HNI-03', 'HNI-04',
      'HNI-05', 'HNI-06', 'HNI-07', 'HNI-09'
    ]
  },
  {
    region: 'HCM',
    regionName: 'TP. Hồ Chí Minh (HCM)',
    branches: [
      'HCM-01', 'HCM-02', 'HCM-03', 'HCM-04',
      'HCM-06', 'HCM-07', 'HCM-08', 'HCM-09',
      'HCM-10', 'HCM-11', 'HCM-16'
    ]
  },
  {
    region: 'TBB',
    regionName: 'Tây Bắc Bộ (TBB)',
    branches: [
      'BGG', 'BNH', 'HBH', 'CBG', 'LCI', 'LSN',
      'PTO', 'TNN', 'TQG', 'VPC', 'YBI', 'HAG'
    ]
  },
  {
    region: 'ĐBB',
    regionName: 'Đông Bắc Bộ (ĐBB)',
    branches: [
      'DBN', 'HNM', 'HTH', 'HYN',
      'NAN', 'NBH', 'NDH', 'SLA', 'TBH', 'THA'
    ]
  },
  {
    region: 'MT&TN',
    regionName: 'Miền Trung & Tây Nguyên (MT&TN)',
    branches: [
      'BDH', 'DLK', 'DKG', 'GLI',
      'KTM', 'PYN', 'QNI', 'QNM',
      'HUE', 'QBH', 'QTI'
    ]
  },
  {
    region: 'ĐNB',
    regionName: 'Đông Nam Bộ (ĐNB)',
    branches: [
      'BTN', 'LAN',
      'LDG', 'NTN', 'TNH'
    ]
  },
  {
    region: 'TNB',
    regionName: 'Tây Nam Bộ (TNB)',
    branches: [
      'AGG', 'BLU', 'BTE', 'CMU',
      'CTO', 'DTP', 'HGG', 'KGG',
      'STG', 'TGG', 'TVH', 'VLG'
    ]
  },
  {
    region: 'QNH',
    regionName: 'Quảng Ninh (QNH)',
    branches: ['QNH']
  },
  {
    region: 'HDG',
    regionName: 'Hải Dương (HDG)',
    branches: ['HDG']
  },
  {
    region: 'HPG',
    regionName: 'Hải Phòng (HPG)',
    branches: ['HPG']
  },
  {
    region: 'DNG',
    regionName: 'Đà Nẵng (DNG)',
    branches: ['DNG']
  },
  {
    region: 'KHA',
    regionName: 'Khánh Hòa (KHA)',
    branches: ['KHA']
  },
  {
    region: 'BDG',
    regionName: 'Bình Dương (BDG)',
    branches: ['BDG']
  },
  {
    region: 'DNI',
    regionName: 'Đồng Nai (DNI)',
    branches: ['DNI', 'BPC']
  },
  {
    region: 'VTU',
    regionName: 'Vũng Tàu (VTU)',
    branches: ['VTU']
  },
  {
    region: 'SOC',
    regionName: 'SOC',
    branches: ['SOC']
  }
];

// Xây dựng bảng tra cứu nhanh Chi nhánh -> Khu vực & STT chuẩn (1 -> 78)
const BRANCH_TO_REGION = {};
const BRANCH_STT_MAP = {};
const ALL_BRANCHES = [];
const ALL_REGIONS = REGION_BRANCH_HIERARCHY.map(g => g.region);

let sttCounter = 1;
REGION_BRANCH_HIERARCHY.forEach(group => {
  group.branches.forEach(branch => {
    BRANCH_TO_REGION[branch] = group.region;
    BRANCH_STT_MAP[branch] = sttCounter++;
    ALL_BRANCHES.push(branch);
  });
});

// Chuẩn hóa chuỗi để map tên Chi nhánh không phụ thuộc dấu / khoảng trắng thừa
function normalizeKey(str) {
  if (!str) return '';
  return String(str)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .toUpperCase()
    .replace(/\s+/g, ' ')
    .replace(/\s*-\s*/g, '-')
    .trim();
}

const NORMALIZED_BRANCH_LOOKUP = {};
ALL_BRANCHES.forEach(b => {
  NORMALIZED_BRANCH_LOOKUP[normalizeKey(b)] = b;
});

// Ánh xạ tự động mọi cách viết (kể cả có TIN / Phuong Nam hay -01, -02 của tỉnh, SG01, HN4, HNI3, HPG02, DNG01...) về tên Chi nhánh chuẩn
const CUSTOM_BRANCH_ALIASES = {
  'HNI': 'HNI-01',
  'HN1': 'HNI-01', 'HN01': 'HNI-01', 'HNI1': 'HNI-01', 'HNI01': 'HNI-01', 'TIN-01': 'HNI-01', 'TIN-01.1': 'HNI-01', 'TIN-01.2': 'HNI-01',
  'HN2': 'HNI-02', 'HN02': 'HNI-02', 'HNI2': 'HNI-02', 'HNI02': 'HNI-02', 'TIN-02': 'HNI-02', 'TIN-02.1': 'HNI-02', 'TIN-02.2': 'HNI-02',
  'HN3': 'HNI-03', 'HN03': 'HNI-03', 'HNI3': 'HNI-03', 'HNI03': 'HNI-03', 'TIN-03': 'HNI-03', 'TIN-03.1': 'HNI-03', 'TIN-03.2': 'HNI-03',
  'HN4': 'HNI-04', 'HN04': 'HNI-04', 'HNI4': 'HNI-04', 'HNI04': 'HNI-04', 'TIN-04': 'HNI-04', 'TIN-04.1': 'HNI-04', 'TIN-04.2': 'HNI-04', 'TIN-04.3': 'HNI-04',
  'HN5': 'HNI-05', 'HN05': 'HNI-05', 'HNI5': 'HNI-05', 'HNI05': 'HNI-05', 'TIN-05': 'HNI-05', 'TIN-05.1': 'HNI-05', 'TIN-05.2': 'HNI-05', 'TIN-05.3': 'HNI-05',
  'HN6': 'HNI-06', 'HN06': 'HNI-06', 'HNI6': 'HNI-06', 'HNI06': 'HNI-06', 'TIN-06': 'HNI-06', 'TIN-06.1': 'HNI-06', 'TIN-06.2': 'HNI-06',
  'HN7': 'HNI-07', 'HN07': 'HNI-07', 'HNI7': 'HNI-07', 'HNI07': 'HNI-07', 'TIN-07': 'HNI-07', 'TIN-07.1': 'HNI-07', 'TIN-07.2': 'HNI-07',
  'HN9': 'HNI-09', 'HN09': 'HNI-09', 'HNI9': 'HNI-09', 'HNI09': 'HNI-09', 'TIN-09': 'HNI-09', 'TIN-09.1': 'HNI-09', 'TIN-09.2': 'HNI-09',
  'HCM': 'HCM-01', 'SG': 'HCM-01',
  'SG1': 'HCM-01', 'SG01': 'HCM-01', 'HCM1': 'HCM-01', 'HCM01': 'HCM-01', 'PHUONG NAM-01': 'HCM-01',
  'SG2': 'HCM-02', 'SG02': 'HCM-02', 'HCM2': 'HCM-02', 'HCM02': 'HCM-02', 'PHUONG NAM-02': 'HCM-02',
  'SG3': 'HCM-03', 'SG03': 'HCM-03', 'HCM3': 'HCM-03', 'HCM03': 'HCM-03', 'PHUONG NAM-03': 'HCM-03',
  'SG4': 'HCM-04', 'SG04': 'HCM-04', 'HCM4': 'HCM-04', 'HCM04': 'HCM-04', 'PHUONG NAM-04': 'HCM-04',
  'SG6': 'HCM-06', 'SG06': 'HCM-06', 'HCM6': 'HCM-06', 'HCM06': 'HCM-06', 'PHUONG NAM-06': 'HCM-06',
  'SG7': 'HCM-07', 'SG07': 'HCM-07', 'HCM7': 'HCM-07', 'HCM07': 'HCM-07', 'PHUONG NAM-07': 'HCM-07',
  'SG8': 'HCM-08', 'SG08': 'HCM-08', 'HCM8': 'HCM-08', 'HCM08': 'HCM-08', 'PHUONG NAM-08': 'HCM-08',
  'SG9': 'HCM-09', 'SG09': 'HCM-09', 'HCM9': 'HCM-09', 'HCM09': 'HCM-09', 'PHUONG NAM-09': 'HCM-09',
  'SG10': 'HCM-10', 'HCM10': 'HCM-10', 'PHUONG NAM-10': 'HCM-10',
  'SG11': 'HCM-11', 'HCM11': 'HCM-11', 'PHUONG NAM-11': 'HCM-11',
  'SG16': 'HCM-16', 'HCM16': 'HCM-16', 'PHUONG NAM-16': 'HCM-16',
  'QNH-01': 'QNH', 'QNH-02': 'QNH', 'QNH-03': 'QNH', 'QNH01': 'QNH', 'QNH02': 'QNH', 'QNH03': 'QNH',
  'HDG-01': 'HDG', 'HDG-02': 'HDG', 'HDG01': 'HDG', 'HDG02': 'HDG',
  'HPG-01': 'HPG', 'HPG-02': 'HPG', 'HPG01': 'HPG', 'HPG02': 'HPG',
  'DNG-01': 'DNG', 'DNG-02': 'DNG', 'DNG01': 'DNG', 'DNG02': 'DNG',
  'KHA-01': 'KHA', 'KHA-02': 'KHA', 'KHA01': 'KHA', 'KHA02': 'KHA',
  'BDG-01': 'BDG', 'BDG-02': 'BDG', 'BDG01': 'BDG', 'BDG02': 'BDG',
  'DNI-01': 'DNI', 'DNI-02': 'DNI', 'DNI-03': 'DNI', 'DNI-04': 'DNI', 'DNI01': 'DNI', 'DNI02': 'DNI', 'DNI03': 'DNI', 'DNI04': 'DNI',
  'HYN-01': 'HYN', 'HYN-02': 'HYN', 'HYN01': 'HYN', 'HYN02': 'HYN',
  'SOC-01': 'SOC', 'SOC-02': 'SOC', 'SOC01': 'SOC', 'SOC02': 'SOC'
};

function resolveCanonicalBranch(rawBranch) {
  if (!rawBranch || !String(rawBranch).trim()) {
    return { branch: 'Chưa điền CN', region: 'KHÁC', matched: false };
  }
  const trimmed = String(rawBranch).trim();
  if (BRANCH_TO_REGION[trimmed]) {
    return { branch: trimmed, region: BRANCH_TO_REGION[trimmed], matched: true };
  }
  // Bỏ tiền tố TIN / Phuong Nam nếu có
  const stripped = normalizeKey(trimmed).replace(/^(TIN|PHUONG NAM)[\s-]+/, '');
  if (NORMALIZED_BRANCH_LOOKUP[stripped]) {
    const canon = NORMALIZED_BRANCH_LOOKUP[stripped];
    return { branch: canon, region: BRANCH_TO_REGION[canon], matched: true };
  }
  if (CUSTOM_BRANCH_ALIASES[stripped]) {
    const canon = CUSTOM_BRANCH_ALIASES[stripped];
    return { branch: canon, region: BRANCH_TO_REGION[canon], matched: true };
  }
  const norm = normalizeKey(trimmed);
  if (CUSTOM_BRANCH_ALIASES[norm]) {
    const canon = CUSTOM_BRANCH_ALIASES[norm];
    return { branch: canon, region: BRANCH_TO_REGION[canon], matched: true };
  }
  // Tự động nhận diện HNI/HN/HCM/SG + số hoặc Mã tỉnh + số (HNI02, HNI3, HPG02, DNG01...)
  const compact = stripped.replace(/[\s\-_.]+/g, '');
  const mHn = compact.match(/^(?:HNI|HN)0?(\d{1,2})$/);
  if (mHn) {
    const cand = `HNI-${mHn[1].padStart(2, '0')}`;
    if (BRANCH_TO_REGION[cand]) {
      return { branch: cand, region: BRANCH_TO_REGION[cand], matched: true };
    }
  }
  const mSg = compact.match(/^(?:HCM|SG)0?(\d{1,2})$/);
  if (mSg) {
    const cand = `HCM-${mSg[1].padStart(2, '0')}`;
    if (BRANCH_TO_REGION[cand]) {
      return { branch: cand, region: BRANCH_TO_REGION[cand], matched: true };
    }
  }
  const mProv = compact.match(/^([A-Z]{3})0?\d{1,2}$/);
  if (mProv && BRANCH_TO_REGION[mProv[1]]) {
    return { branch: mProv[1], region: BRANCH_TO_REGION[mProv[1]], matched: true };
  }
  return { branch: trimmed, region: 'KHÁC', matched: false };
}

function normalizeDepartment(rawDept) {
  if (!rawDept || !String(rawDept).trim()) return 'Chưa phân loại';
  const raw = String(rawDept).trim();
  const up = raw.toUpperCase().replace(/\s+/g, ' ');
  if (up === 'TPDVKH' || up === 'TP DVKH') return 'TP DVKH';
  if (up === 'TPKD') return 'TPKD';
  if (up === 'PPKT') return 'PPKT';
  return up;
}

// Quy tắc tính điểm Nụ cười từ cột "Loại"
function parseMediaType(rawType) {
  const norm = normalizeKey(rawType);
  if (norm.includes('VIDEO') || norm.includes('CLIP') || norm.includes('PHIM') || norm === '2' || norm === '5') {
    return { code: 'VIDEO', label: 'Video', smiles: 5, images: 0, videos: 1 };
  }
  // Mặc định Hình ảnh = 3 nụ cười
  return { code: 'IMAGE', label: 'Hình ảnh', smiles: 3, images: 1, videos: 0 };
}

// Parse cột "Thời gian đăng" đa định dạng (DD/MM/YYYY, YYYY-MM-DD, Google Date(), Excel serial)
function parsePostDate(rawVal) {
  if (!rawVal && rawVal !== 0) return null;
  if (rawVal instanceof Date && !isNaN(rawVal.getTime())) return rawVal;

  // Excel serial number
  if (typeof rawVal === 'number' && rawVal > 25000 && rawVal < 80000) {
    const utcDays = Math.floor(rawVal - 25569);
    const utcValue = utcDays * 86400;
    const d = new Date(utcValue * 1000);
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), 12, 0, 0);
  }

  const str = String(rawVal).trim();
  if (!str) return null;

  // Google Visualization Date(2026, 10, 5, 8, 30, 0)
  const gvizMatch = str.match(/^Date\((\d+),(\d+),(\d+)(?:,(\d+),(\d+),(\d+))?\)/);
  if (gvizMatch) {
    return new Date(
      Number(gvizMatch[1]),
      Number(gvizMatch[2]),
      Number(gvizMatch[3]),
      Number(gvizMatch[4] || 0),
      Number(gvizMatch[5] || 0),
      Number(gvizMatch[6] || 0)
    );
  }

  // DD/MM/YYYY or DD-MM-YYYY (with optional HH:mm:ss)
  const dmyMatch = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if (dmyMatch) {
    const day = Number(dmyMatch[1]);
    const month = Number(dmyMatch[2]);
    const yr = Number(dmyMatch[3]);
    const hr = Number(dmyMatch[4] || 0);
    const mn = Number(dmyMatch[5] || 0);
    const sc = Number(dmyMatch[6] || 0);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return new Date(yr, month - 1, day, hr, mn, sc);
    }
  }

  // HH:mm(:ss) DD/MM/YYYY or DD-MM-YYYY (e.g., "20:24 06/11/2026")
  const timeFirstMatch = str.match(/^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s+(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (timeFirstMatch) {
    const hr = Number(timeFirstMatch[1]);
    const mn = Number(timeFirstMatch[2]);
    const sc = Number(timeFirstMatch[3] || 0);
    const day = Number(timeFirstMatch[4]);
    const month = Number(timeFirstMatch[5]);
    const yr = Number(timeFirstMatch[6]);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return new Date(yr, month - 1, day, hr, mn, sc);
    }
  }

  // YYYY-MM-DD or ISO
  const ymdMatch = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:[T\s](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if (ymdMatch) {
    return new Date(
      Number(ymdMatch[1]),
      Number(ymdMatch[2]) - 1,
      Number(ymdMatch[3]),
      Number(ymdMatch[4] || 0),
      Number(ymdMatch[5] || 0),
      Number(ymdMatch[6] || 0)
    );
  }

  return null;
}

function toYMD(dateObj) {
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, '0');
  const d = String(dateObj.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function formatDMY(ymdStr) {
  if (!ymdStr) return '';
  const [y, m, d] = ymdStr.split('-');
  return `${d}/${m}/${y}`;
}

function formatShortDM(ymdStr) {
  if (!ymdStr) return '';
  const [, m, d] = ymdStr.split('-');
  return `${d}/${m}`;
}

// Tính Tuần ISO (Thứ 2 -> Chủ Nhật) & nhãn tuần
function getWeekInfo(dateObj) {
  const d = new Date(Date.UTC(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  const isoYear = d.getUTCFullYear();

  // Ngày Thứ 2 đầu tuần
  const monday = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
  const localDay = monday.getDay() || 7;
  monday.setDate(monday.getDate() - localDay + 1);
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);

  const key = `${isoYear}-W${String(weekNo).padStart(2, '0')}`;
  const label = `Tuần ${String(weekNo).padStart(2, '0')} (${String(monday.getDate()).padStart(2, '0')}/${String(monday.getMonth() + 1).padStart(2, '0')}-${String(sunday.getDate()).padStart(2, '0')}/${String(sunday.getMonth() + 1).padStart(2, '0')})`;
  return { key, label, weekNo, year: isoYear };
}

function getMonthInfo(dateObj) {
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, '0');
  return {
    key: `${y}-${m}`,
    label: `Tháng ${m}/${y}`
  };
}

// Trạng thái ứng dụng
const state = {
  rawRecords: [],
  filteredRecords: [],
  timeGrain: 'day', // 'day' | 'week' | 'month'
  startDate: '',
  endDate: '',
  region: 'ALL',
  branch: 'ALL',
  dept: 'ALL',
  mediaType: 'ALL',
  keyword: '',
  showZeroBranches: true,
  showTimeColumns: true,
  branchSort: 'hierarchy',
  regionSort: 'smiles_desc',
  deptMatrixLevel: 'branch_and_region',
  isDemoData: false,
  charts: {}
};

// Tạo dữ liệu mô phỏng (Demo) bám sát đúng 99 Chi nhánh, 15 Khu vực & các Bộ phận thực tế của TTC
function generateRealisticDemoData() {
  const departments = [
    'Kỹ thuật (KT)',
    'Chăm sóc khách hàng (CSKH)',
    'Kinh doanh (KD)',
    'Hành chính - Nhân sự (HCNS)',
    'Kho vận & Điều phối',
    'Quản lý chất lượng (QA)'
  ];

  // Tạo dữ liệu trải đều trong 45 ngày gần nhất quanh ngày hiện tại của máy
  const now = new Date();
  const records = [];
  let idCounter = 1;

  // Deterministic pseudo-random để dữ liệu demo ổn định và chân thực
  let seed = 20261105;
  function nextRand() {
    seed = (seed * 1664525 + 1013904223) % 4294967296;
    return seed / 4294967296;
  }

  for (let dayOffset = 38; dayOffset >= 0; dayOffset--) {
    const d = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dayOffset);
    // Số bài đăng mỗi ngày dao động 18 - 42 bài trên toàn hệ thống
    const postsToday = 18 + Math.floor(nextRand() * 24);

    for (let i = 0; i < postsToday; i++) {
      const branchIdx = Math.floor(nextRand() * ALL_BRANCHES.length);
      const branch = ALL_BRANCHES[branchIdx];
      const dept = departments[Math.floor(nextRand() * departments.length)];
      const isVideo = nextRand() < 0.34; // ~34% là Video (= 2 nụ cười), 66% là Hình ảnh (= 1 nụ cười)
      const hour = 7 + Math.floor(nextRand() * 13);
      const minute = Math.floor(nextRand() * 60);
      const postDate = new Date(d.getFullYear(), d.getMonth(), d.getDate(), hour, minute, 0);

      records.push(buildRecord({
        id: idCounter++,
        rawTimestamp: `${String(postDate.getDate()).padStart(2, '0')}/${String(postDate.getMonth() + 1).padStart(2, '0')}/${postDate.getFullYear()} ${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`,
        dateObj: postDate,
        rawBranch: branch,
        rawDept: dept,
        rawType: isVideo ? 'Video' : 'Hình ảnh'
      }));
    }
  }

  return records;
}

function buildRecord({ id, rawTimestamp, dateObj, rawBranch, rawDept, rawType, authorName, authorEmail, postTitle }) {
  const resolved = resolveCanonicalBranch(rawBranch);
  const media = parseMediaType(rawType);
  const ymd = toYMD(dateObj);
  const weekInfo = getWeekInfo(dateObj);
  const monthInfo = getMonthInfo(dateObj);
  const deptClean = normalizeDepartment(rawDept);

  return {
    id,
    rawTimestamp: rawTimestamp || formatDMY(ymd),
    dateObj,
    ymd,
    dmy: formatDMY(ymd),
    weekKey: weekInfo.key,
    weekLabel: weekInfo.label,
    monthKey: monthInfo.key,
    monthLabel: monthInfo.label,
    rawBranch: String(rawBranch || '').trim() || '(Trống)',
    branch: resolved.branch,
    region: resolved.region,
    branchMatched: resolved.matched,
    dept: deptClean,
    rawType: String(rawType || 'Hình ảnh').trim(),
    mediaCode: media.code,
    mediaLabel: media.label,
    images: media.images,
    videos: media.videos,
    smiles: media.smiles,
    authorName: String(authorName || '').trim(),
    authorEmail: String(authorEmail || '').trim(),
    postTitle: String(postTitle || '').trim()
  };
}

// Chuyển đổi mảng dữ liệu thô (từ Google Sheet / Excel / CSV / JSON) sang danh sách chuẩn
function ingestTabularRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return [];

  // Nếu là mảng các object (ví dụ từ JSON hoặc SheetJS sheet_to_json)
  const sample = rows[0];
  const keys = Object.keys(sample);

  function findColKey(patterns) {
    for (const pat of patterns) {
      const found = keys.find(k => normalizeKey(k).includes(pat));
      if (found) return found;
    }
    return null;
  }

  let colTime = findColKey(['THOI GIAN DANG', 'THOI GIAN', 'NGAY DANG', 'TIMESTAMP', 'NGAY', 'DATE', 'TIME']) || keys[0];
  let colBranch = findColKey(['CHI NHANH', 'DON VI', 'BRANCH', 'TRUNG TAM']) || keys[1];
  let colType = findColKey(['LOAI', 'HINH THUC', 'HINH ANH', 'VIDEO', 'TYPE', 'MEDIA']) || keys[2];
  let colDept = findColKey(['BO PHAN', 'PHONG BAN', 'DEPARTMENT', 'BP', 'NHOM']) || keys[3];
  let colAuthorName = findColKey(['HO TEN NGUOI DANG', 'HO TEN', 'NGUOI DANG']);
  let colAuthorEmail = findColKey(['EMAIL NGUOI DANG', 'EMAIL']);
  let colTitle = findColKey(['TIEU DE BAI DANG', 'TIEU DE']);
  let colStt = findColKey(['STT']);

  // Tự động phát hiện nếu file Excel/CSV tải lên bị lệch dòng tiêu đề sang phải 1 cột (như trên Google Sheet)
  if (colStt && sample[colStt] && String(sample[colStt]).includes('@') && !parsePostDate(sample[colTime])) {
    const contentKey = findColKey(['NOI DUNG BAI DANG']);
    const approverEmailKey = findColKey(['EMAIL NGUOI DUYET']);
    const emptyKey = keys.find(k => k === '' || k.startsWith('__EMPTY'));
    if (contentKey && parsePostDate(sample[contentKey])) {
      colTime = contentKey;
      if (approverEmailKey) colType = approverEmailKey;
      if (colAuthorName) colTitle = colAuthorName;
      if (colAuthorEmail) colAuthorName = colAuthorEmail;
      colAuthorEmail = colStt;
      if (emptyKey) colStt = emptyKey;
    }
  }

  const parsed = [];
  let idx = 1;
  for (const row of rows) {
    const rawTime = row[colTime];
    const rawBranch = row[colBranch];
    const rawType = colType ? row[colType] : 'Hình ảnh';
    const rawDept = colDept ? row[colDept] : '';

    if (!rawTime && !rawBranch) continue;
    const dateObj = parsePostDate(rawTime);
    if (!dateObj) continue;

    parsed.push(buildRecord({
      id: (colStt && row[colStt]) ? row[colStt] : idx++,
      rawTimestamp: String(rawTime),
      dateObj,
      rawBranch,
      rawDept,
      rawType,
      authorName: colAuthorName ? row[colAuthorName] : '',
      authorEmail: colAuthorEmail ? row[colAuthorEmail] : '',
      postTitle: colTitle ? row[colTitle] : ''
    }));
  }

  parsed.sort((a, b) => b.dateObj - a.dateObj);
  return parsed;
}

// Đồng bộ dữ liệu từ file data.json (do script Python tạo) hoặc trực tiếp từ Google Sheet (JSONP gviz)
async function trySyncDataSource(showAlertOnFail = false) {
  const btn = document.getElementById('btnSyncSheet');
  if (btn) {
    btn.disabled = true;
    btn.textContent = '⏳ Đang đồng bộ...';
  }

  try {
    // 0. Nếu đang chạy qua server.py (http://localhost:8080), gọi /api/sync để kéo trực tiếp từ Google Sheet qua gsheets CLI
    if (showAlertOnFail && window.location.protocol.startsWith('http')) {
      try {
        const apiResp = await fetch(`/api/sync?t=${Date.now()}`);
        if (apiResp.ok) {
          const payload = await apiResp.json();
          if (payload && Array.isArray(payload.records) && payload.records.length > 0) {
            const parsed = ingestTabularRows(payload.records);
            if (parsed.length > 0) {
              applyNewRecords(parsed, false, `Google Sheet Live (${parsed.length} bài đăng)`);
              return;
            }
          }
        }
      } catch (_) {}
    }

    // 1. Ưu tiên đọc từ file data.json nếu script Python/gsheets đã kéo về (chống cache hoàn toàn)
    try {
      const resp = await fetch(`data.json?t=${Date.now()}`, { cache: 'no-store' });
      if (resp.ok) {
        const payload = await resp.json();
        if (payload && Array.isArray(payload.records) && payload.records.length > 0) {
          if (payload.updated_at) {
            window.LIVE_SHEET_DATA = payload;
          }
          const parsed = ingestTabularRows(payload.records);
          if (parsed.length > 0) {
            applyNewRecords(parsed, false, `Dữ liệu thực Google Sheet (${parsed.length} bài đăng)`);
            return;
          }
        }
      }
    } catch (_) {
      // Bỏ qua nếu đang mở file:// trực tiếp hoặc chưa có data.json
    }

    // 2. Thử kết nối trực tiếp Google Sheet qua JSONP (gviz/tq)
    const gvizRows = await fetchGoogleSheetViaJsonp(SPREADSHEET_ID, GID);
    if (gvizRows && gvizRows.length > 0) {
      const parsed = ingestTabularRows(gvizRows);
      if (parsed.length > 0) {
        applyNewRecords(parsed, false, `Google Sheet Live (${parsed.length} dòng)`);
        return;
      }
    }

    // 3. Nếu đang mở trực tiếp file:// và đã có gói dữ liệu thực tế trong window.LIVE_SHEET_DATA
    if (window.LIVE_SHEET_DATA && Array.isArray(window.LIVE_SHEET_DATA.records) && window.LIVE_SHEET_DATA.records.length > 0) {
      const parsed = ingestTabularRows(window.LIVE_SHEET_DATA.records);
      if (parsed.length > 0) {
        applyNewRecords(parsed, false, `Dữ liệu thực Google Sheet (${parsed.length} bài đăng)`);
        return;
      }
    }

    if (showAlertOnFail) {
      alert(
        'Hiện tại file Google Sheet chưa mở quyền xem công khai hoặc chưa chạy lệnh đăng nhập gsheets.\n\n' +
        'Anh có thể:\n' +
        '1. Mở quyền "Anyone with the link = Viewer" trên Google Sheet rồi bấm Đồng bộ lại.\n' +
        '2. Hoặc bấm nút "📂 Nhập File (.xlsx / .csv)" để tải file từ máy tính lên ngay lập tức.\n' +
        '3. Hoặc chạy `python sync_and_report.py` để đồng bộ tự động.'
      );
    }
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = '🔄 Đồng bộ Google Sheet';
    }
  }
}

function fetchGoogleSheetViaJsonp(sheetId, gid) {
  return new Promise((resolve) => {
    const callbackName = 'gvizCallback_' + Math.floor(Math.random() * 1000000);
    const script = document.createElement('script');
    let settled = false;

    const timer = setTimeout(() => {
      if (!settled) {
        settled = true;
        cleanup();
        resolve(null);
      }
    }, 6000);

    function cleanup() {
      clearTimeout(timer);
      delete window[callbackName];
      if (script.parentNode) script.parentNode.removeChild(script);
    }

    window[callbackName] = function (response) {
      if (settled) return;
      settled = true;
      cleanup();
      try {
        if (!response || response.status !== 'ok' || !response.table) {
          resolve(null);
          return;
        }
        const cols = response.table.cols.map((c, i) => (c.label || c.id || `Col_${i}`).trim());
        const rows = (response.table.rows || []).map(r => {
          const obj = {};
          cols.forEach((colName, idx) => {
            const cell = r.c ? r.c[idx] : null;
            obj[colName] = cell ? (cell.f !== undefined && cell.f !== null ? cell.f : cell.v) : '';
          });
          return obj;
        });
        resolve(rows);
      } catch (e) {
        resolve(null);
      }
    };

    script.onerror = function () {
      if (!settled) {
        settled = true;
        cleanup();
        resolve(null);
      }
    };

    script.src = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?gid=${gid}&tqx=responseHandler:${callbackName}`;
    document.body.appendChild(script);
  });
}

function applyNewRecords(records, isDemo, sourceLabel) {
  state.rawRecords = records;
  state.isDemoData = isDemo;

  const pill = document.getElementById('dataSourcePill');
  const text = document.getElementById('dataSourceText');
  const syncText = document.getElementById('lastSyncText');

  if (pill && text) {
    pill.className = isDemo ? 'status-pill demo' : 'status-pill';
    text.textContent = sourceLabel;
  }
  if (syncText) {
    const syncTime = (window.LIVE_SHEET_DATA && window.LIVE_SHEET_DATA.updated_at)
      ? window.LIVE_SHEET_DATA.updated_at
      : `${new Date().toLocaleTimeString('vi-VN')} ${new Date().toLocaleDateString('vi-VN')}`;
    syncText.textContent = `Cập nhật: ${syncTime} (Tự động 1h/lần)`;
  }

  populateFilterDropdowns();
  applyQuickRange('this_month');
}

function populateFilterDropdowns() {
  // Region dropdown
  const regSelect = document.getElementById('filterRegion');
  regSelect.innerHTML = `<option value="ALL">-- Tất cả Khu vực (${REGION_BRANCH_HIERARCHY.length}) --</option>`;
  REGION_BRANCH_HIERARCHY.forEach(g => {
    const opt = document.createElement('option');
    opt.value = g.region;
    opt.textContent = `${g.region} - ${g.regionName} (${g.branches.length} CN)`;
    regSelect.appendChild(opt);
  });

  updateBranchDropdown('ALL');

  // Department dropdown
  const deptSelect = document.getElementById('filterDept');
  const depts = Array.from(new Set(state.rawRecords.map(r => r.dept))).sort((a, b) => a.localeCompare(b, 'vi'));
  deptSelect.innerHTML = `<option value="ALL">-- Tất cả Bộ phận (${depts.length}) --</option>`;
  depts.forEach(d => {
    const opt = document.createElement('option');
    opt.value = d;
    opt.textContent = d;
    deptSelect.appendChild(opt);
  });
}

function updateBranchDropdown(selectedRegion) {
  const brSelect = document.getElementById('filterBranch');
  const prevVal = brSelect.value;
  const branches = selectedRegion === 'ALL'
    ? ALL_BRANCHES
    : (REGION_BRANCH_HIERARCHY.find(g => g.region === selectedRegion)?.branches || []);

  brSelect.innerHTML = `<option value="ALL">-- Tất cả Chi nhánh (${branches.length}) --</option>`;
  branches.forEach(b => {
    const opt = document.createElement('option');
    opt.value = b;
    opt.textContent = `${BRANCH_STT_MAP[b]}. ${b} (${BRANCH_TO_REGION[b]})`;
    brSelect.appendChild(opt);
  });

  if (branches.includes(prevVal)) {
    brSelect.value = prevVal;
  } else {
    brSelect.value = 'ALL';
    state.branch = 'ALL';
  }
}

function applyQuickRange(rangeCode) {
  document.querySelectorAll('#quickRangeGroup .chip-btn').forEach(b => {
    b.classList.toggle('active', b.dataset.range === rangeCode);
  });

  const now = new Date();
  const minDateObj = state.rawRecords.length > 0
    ? state.rawRecords.reduce((min, r) => (r.dateObj < min ? r.dateObj : min), state.rawRecords[0].dateObj)
    : now;

  // Neo theo ngày thực tế hiện tại (Hôm nay)
  const ref = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  let start = new Date(ref);
  let end = new Date(ref);

  if (rangeCode === 'today') {
    // Hôm nay (ngày thực tế hiện tại)
  } else if (rangeCode === 'yesterday') {
    start.setDate(ref.getDate() - 1);
    end = new Date(start);
  } else if (rangeCode === 'this_week') {
    const day = ref.getDay() || 7;
    start.setDate(ref.getDate() - day + 1);
  } else if (rangeCode === 'last_7d') {
    start.setDate(ref.getDate() - 6);
  } else if (rangeCode === 'this_month') {
    start = new Date(ref.getFullYear(), ref.getMonth(), 1);
  } else if (rangeCode === 'last_month') {
    start = new Date(ref.getFullYear(), ref.getMonth() - 1, 1);
    end = new Date(ref.getFullYear(), ref.getMonth(), 0);
  } else if (rangeCode === 'all') {
    start = new Date(minDateObj.getFullYear(), minDateObj.getMonth(), minDateObj.getDate());
  }

  state.startDate = toYMD(start);
  state.endDate = toYMD(end);
  document.getElementById('filterStartDate').value = state.startDate;
  document.getElementById('filterEndDate').value = state.endDate;

  refreshDashboard();
}

// Lọc dữ liệu theo các điều kiện trên thanh Filter
function filterRecords() {
  const kw = normalizeKey(state.keyword);
  state.filteredRecords = state.rawRecords.filter(r => {
    if (state.startDate && r.ymd < state.startDate) return false;
    if (state.endDate && r.ymd > state.endDate) return false;
    if (state.region !== 'ALL' && r.region !== state.region) return false;
    if (state.branch !== 'ALL' && r.branch !== state.branch) return false;
    if (state.dept !== 'ALL' && r.dept !== state.dept) return false;
    if (state.mediaType !== 'ALL' && r.mediaCode !== state.mediaType) return false;
    if (kw) {
      const hay = normalizeKey(`${r.branch} ${r.region} ${r.dept} ${r.rawBranch} ${r.authorEmail} ${r.authorName} ${r.postTitle}`);
      if (!hay.includes(kw)) return false;
    }
    return true;
  });
}

// Lấy danh sách các cột thời gian động theo Chiều thống kê (Ngày / Tuần / Tháng)
function getTimeBuckets(records) {
  const map = new Map();
  records.forEach(r => {
    let key, label;
    if (state.timeGrain === 'day') {
      key = r.ymd;
      label = formatShortDM(r.ymd);
    } else if (state.timeGrain === 'week') {
      key = r.weekKey;
      label = r.weekLabel;
    } else {
      key = r.monthKey;
      label = r.monthLabel;
    }
    if (!map.has(key)) {
      map.set(key, { key, label });
    }
  });
  return Array.from(map.values()).sort((a, b) => a.key.localeCompare(b.key));
}

function getRecordBucketKey(r) {
  if (state.timeGrain === 'day') return r.ymd;
  if (state.timeGrain === 'week') return r.weekKey;
  return r.monthKey;
}

// Tính các mốc thời gian chuẩn theo ngày thực tế hiện tại: Hôm nay, Hôm qua, Hôm kia, Tuần này, Tuần trước, Tháng này
function getTemporalMilestones() {
  const now = new Date();
  const todayObj = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const yestObj = new Date(todayObj.getFullYear(), todayObj.getMonth(), todayObj.getDate() - 1);
  const prevObj = new Date(todayObj.getFullYear(), todayObj.getMonth(), todayObj.getDate() - 2);
  const prevWeekObj = new Date(todayObj.getFullYear(), todayObj.getMonth(), todayObj.getDate() - 7);
  return {
    latestYmd: toYMD(todayObj),
    yesterdayYmd: toYMD(yestObj),
    prevDayYmd: toYMD(prevObj),
    currentWeekKey: getWeekInfo(todayObj).key,
    prevWeekKey: getWeekInfo(prevWeekObj).key,
    currentMonthKey: getMonthInfo(todayObj).key
  };
}

function formatDeltaHtml(diff) {
  if (diff > 0) return `<span class="delta-up">▲ +${diff}</span>`;
  if (diff < 0) return `<span class="delta-down">▼ ${diff}</span>`;
  return `<span class="delta-flat">= 0</span>`;
}

function formatRankBadge(rank) {
  if (!rank || rank <= 0) return `<span class="rank-pill">-</span>`;
  if (rank === 1) return `<span class="rank-pill gold">🥇 #1</span>`;
  if (rank === 2) return `<span class="rank-pill silver">🥈 #2</span>`;
  if (rank === 3) return `<span class="rank-pill bronze">🥉 #3</span>`;
  return `<span class="rank-pill">#${rank}</span>`;
}

// Lọc theo Đơn vị / Khu vực / Bộ phận / Loại (không khóa cứng khoảng ngày để tính đúng Hôm nay / Hôm qua / Tuần / Tháng)
function getUnitScopedRecords() {
  const kw = normalizeKey(state.keyword);
  return state.rawRecords.filter(r => {
    if (state.region !== 'ALL' && r.region !== state.region) return false;
    if (state.branch !== 'ALL' && r.branch !== state.branch) return false;
    if (state.dept !== 'ALL' && r.dept !== state.dept) return false;
    if (state.mediaType !== 'ALL' && r.mediaCode !== state.mediaType) return false;
    if (kw) {
      const hay = normalizeKey(`${r.branch} ${r.region} ${r.dept} ${r.rawBranch} ${r.authorEmail} ${r.authorName} ${r.postTitle}`);
      if (!hay.includes(kw)) return false;
    }
    return true;
  });
}

// Cập nhật toàn bộ Dashboard
function refreshDashboard() {
  filterRecords();
  const timeBuckets = getTimeBuckets(state.filteredRecords);

  document.querySelectorAll('#regionHorizontalBar [data-region-chip]').forEach(btn => {
    btn.classList.toggle('active', btn.getAttribute('data-region-chip') === state.region);
  });

  const grainLabelMap = { day: 'Ngày', week: 'Tuần', month: 'Tháng' };
  const lblMode = document.getElementById('lblTimeColumnMode');
  if (lblMode) lblMode.textContent = grainLabelMap[state.timeGrain];

  renderKpis();
  renderUnitSpotlight();
  renderLeaderboardStrip();
  renderBranchAndRegionTable(timeBuckets);
  renderRegionSummaryTable(timeBuckets);
  renderDepartmentTables(timeBuckets);
  renderCharts(timeBuckets);
  renderRawDataTable();
}

function renderKpis() {
  const recs = state.filteredRecords;
  const unitRecs = getUnitScopedRecords();
  const m = getTemporalMilestones();

  let todaySmiles = 0, todayImages = 0, todayVideos = 0;
  let yestSmiles = 0, yestImages = 0, yestVideos = 0;
  let prevDaySmiles = 0;
  let weekSmiles = 0, weekImages = 0, weekVideos = 0;
  let prevWeekSmiles = 0;

  unitRecs.forEach(r => {
    if (r.ymd === m.latestYmd) {
      todaySmiles += r.smiles;
      todayImages += r.images;
      todayVideos += r.videos;
    }
    if (r.ymd === m.yesterdayYmd) {
      yestSmiles += r.smiles;
      yestImages += r.images;
      yestVideos += r.videos;
    }
    if (r.ymd === m.prevDayYmd) {
      prevDaySmiles += r.smiles;
    }
    if (r.weekKey === m.currentWeekKey) {
      weekSmiles += r.smiles;
      weekImages += r.images;
      weekVideos += r.videos;
    }
    if (r.weekKey === m.prevWeekKey) {
      prevWeekSmiles += r.smiles;
    }
  });

  let totalSmiles = 0;
  let totalImages = 0;
  let totalVideos = 0;
  const activeBranchesSet = new Set();
  const activeRegionsSet = new Set();
  const activeDaysSet = new Set();
  const regionSmiles = {};
  const deptSmiles = {};

  recs.forEach(r => {
    totalSmiles += r.smiles;
    totalImages += r.images;
    totalVideos += r.videos;
    if (BRANCH_TO_REGION[r.branch]) {
      activeBranchesSet.add(r.branch);
      activeRegionsSet.add(r.region);
    }
    activeDaysSet.add(r.ymd);
    regionSmiles[r.region] = (regionSmiles[r.region] || 0) + r.smiles;
    deptSmiles[r.dept] = (deptSmiles[r.dept] || 0) + r.smiles;
  });

  const numDays = Math.max(1, activeDaysSet.size);
  const avgPerDay = (totalSmiles / numDays).toFixed(1);

  const totalTargetBranches = state.region === 'ALL'
    ? ALL_BRANCHES.length
    : (REGION_BRANCH_HIERARCHY.find(g => g.region === state.region)?.branches.length || ALL_BRANCHES.length);

  const branchRate = totalTargetBranches > 0
    ? ((activeBranchesSet.size / totalTargetBranches) * 100).toFixed(1)
    : '0.0';

  // Thẻ 1: Hôm nay
  const elTodayTitle = document.getElementById('kpiTodayTitle');
  if (elTodayTitle) {
    elTodayTitle.textContent = `Hôm Nay (${formatShortDM(m.latestYmd)})`;
    document.getElementById('kpiTodaySmiles').textContent = `+${todaySmiles.toLocaleString('vi-VN')} 😊`;
    document.getElementById('kpiTodayDetail').textContent = `${todayImages} Ảnh (×3) • ${todayVideos} Video (×5)`;
    const diffToday = todaySmiles - yestSmiles;
    const signToday = diffToday >= 0 ? `+${diffToday}` : `${diffToday}`;
    document.getElementById('kpiTodayVsYesterday').textContent = `Hôm qua (${formatShortDM(m.yesterdayYmd)}): ${yestSmiles} 😊 (${signToday})`;
  }

  // Thẻ 2: Hôm qua & So với hôm kia
  const elYestTitle = document.getElementById('kpiYesterdayTitle');
  if (elYestTitle) {
    elYestTitle.textContent = `Hôm Qua (${formatShortDM(m.yesterdayYmd)})`;
    document.getElementById('kpiYesterdaySmiles').textContent = `+${yestSmiles.toLocaleString('vi-VN')} 😊`;
    document.getElementById('kpiYesterdayDetail').textContent = `${yestImages} Ảnh (×3) • ${yestVideos} Video (×5)`;
    const diffYest = yestSmiles - prevDaySmiles;
    const arrowYest = diffYest > 0 ? `▲ +${diffYest}` : (diffYest < 0 ? `▼ ${diffYest}` : `= 0`);
    document.getElementById('kpiYesterdayVsPrev').textContent = `So với ${formatShortDM(m.prevDayYmd)}: ${arrowYest} 😊`;
  }

  // Thẻ 3: Lũy kế Tuần này
  const elWeekTitle = document.getElementById('kpiWeekTitle');
  if (elWeekTitle) {
    elWeekTitle.textContent = `Tuần Này (${m.currentWeekKey.split('-')[1]})`;
    document.getElementById('kpiWeekSmiles').textContent = `${weekSmiles.toLocaleString('vi-VN')} 😊`;
    document.getElementById('kpiWeekDetail').textContent = `${weekImages} Ảnh (×3) • ${weekVideos} Video (×5)`;
    const diffWeek = weekSmiles - prevWeekSmiles;
    const arrowWeek = diffWeek > 0 ? `▲ +${diffWeek}` : (diffWeek < 0 ? `▼ ${diffWeek}` : `= 0`);
    document.getElementById('kpiWeekVsPrev').textContent = `Tuần trước: ${prevWeekSmiles} 😊 (${arrowWeek})`;
  }

  // Thẻ 4: Lũy kế Tháng / Kỳ đang xem
  const elMonthTitle = document.getElementById('kpiMonthTitle');
  if (elMonthTitle) {
    elMonthTitle.textContent = `Lũy Kế Tháng (${m.currentMonthKey})`;
  }
  document.getElementById('kpiTotalSmiles').textContent = `${totalSmiles.toLocaleString('vi-VN')} 😊`;
  document.getElementById('kpiTotalPostsText').textContent = `${totalImages.toLocaleString('vi-VN')} Ảnh (×3) + ${totalVideos.toLocaleString('vi-VN')} Video (×5)`;
  document.getElementById('kpiAvgSmilesPerDay').textContent = `TB: ${avgPerDay} nụ cười/ngày`;

  // Thẻ 5: Độ phủ hoặc Thứ hạng Chi nhánh đang chọn
  const elCovTitle = document.getElementById('kpiCoverageTitle');
  if (state.branch !== 'ALL') {
    // Tính thứ hạng toàn quốc và trong khu vực cho chi nhánh đang chọn
    const allBrScores = {};
    ALL_BRANCHES.forEach(b => { allBrScores[b] = 0; });
    state.rawRecords.forEach(r => {
      if (allBrScores[r.branch] !== undefined) allBrScores[r.branch] += r.smiles;
    });
    const sortedNational = Object.entries(allBrScores).sort((a, b) => b[1] - a[1]);
    const natRank = sortedNational.findIndex(x => x[0] === state.branch) + 1;
    const regCode = BRANCH_TO_REGION[state.branch];
    const regBranches = REGION_BRANCH_HIERARCHY.find(g => g.region === regCode)?.branches || [];
    const sortedReg = regBranches.map(b => [b, allBrScores[b] || 0]).sort((a, b) => b[1] - a[1]);
    const regRank = sortedReg.findIndex(x => x[0] === state.branch) + 1;

    if (elCovTitle) elCovTitle.textContent = `Xếp Hạng Đơn Vị (${state.branch})`;
    document.getElementById('kpiActiveBranches').textContent = `Hạng #${natRank} / ${ALL_BRANCHES.length}`;
    document.getElementById('kpiActiveRegionsText').textContent = `Hạng #${regRank} / ${regBranches.length} tại KV ${regCode}`;
    document.getElementById('kpiBranchRate').textContent = `${totalSmiles} điểm lũy kế`;
  } else {
    if (elCovTitle) elCovTitle.textContent = state.region === 'ALL' ? 'Độ Phủ Toàn Hệ Thống' : `Độ Phủ KV ${state.region}`;
    document.getElementById('kpiActiveBranches').textContent = `${activeBranchesSet.size} / ${totalTargetBranches} CN`;
    document.getElementById('kpiActiveRegionsText').textContent = `${activeRegionsSet.size} / ${REGION_BRANCH_HIERARCHY.length} Khu vực có bài`;
    document.getElementById('kpiBranchRate').textContent = `${branchRate}% CN đã tham gia`;
  }

  document.getElementById('badgeDeptCount').textContent = `${Object.keys(deptSmiles).length} BP`;
  document.getElementById('badgeRawCount').textContent = `${recs.length} bài`;
}

// Hiển thị Banner thông tin nhanh khi người dùng chọn 1 Khu vực hoặc Chi nhánh cụ thể
function renderUnitSpotlight() {
  const banner = document.getElementById('unitSpotlightBanner');
  if (!banner) return;

  if (state.region === 'ALL' && state.branch === 'ALL') {
    banner.style.display = 'none';
    return;
  }

  const m = getTemporalMilestones();
  const unitRecs = getUnitScopedRecords();
  let todayS = 0, yestS = 0, prevS = 0, weekS = 0, mtdS = 0;
  unitRecs.forEach(r => {
    if (r.ymd === m.latestYmd) todayS += r.smiles;
    if (r.ymd === m.yesterdayYmd) yestS += r.smiles;
    if (r.ymd === m.prevDayYmd) prevS += r.smiles;
    if (r.weekKey === m.currentWeekKey) weekS += r.smiles;
    if (r.monthKey === m.currentMonthKey) mtdS += r.smiles;
  });

  const diffYest = yestS - prevS;
  const deltaStr = diffYest > 0 ? `+${diffYest}` : `${diffYest}`;
  const titleLabel = state.branch !== 'ALL'
    ? `🏢 Chi nhánh <strong>${state.branch}</strong> (Khu vực ${BRANCH_TO_REGION[state.branch]})`
    : `🗺️ Khu vực <strong>${state.region}</strong>`;

  banner.style.display = 'flex';
  banner.innerHTML = `
    <div>
      📍 Đang xem kết quả của: ${titleLabel}
    </div>
    <div class="unit-spotlight-metrics">
      <span class="spotlight-chip">🔥 Hôm nay (${formatShortDM(m.latestYmd)}): <strong>+${todayS} 😊</strong></span>
      <span class="spotlight-chip">⚡ Hôm qua (${formatShortDM(m.yesterdayYmd)}): <strong>+${yestS} 😊</strong> (${deltaStr} so với ${formatShortDM(m.prevDayYmd)})</span>
      <span class="spotlight-chip">📅 Tuần này: <strong>${weekS} 😊</strong></span>
      <span class="spotlight-chip">🏆 Lũy kế tháng: <strong>${mtdS} 😊</strong></span>
      <button type="button" id="btnResetUnitFilter" style="background:#1e3a8a;color:#fff;border:none;border-radius:6px;padding:4px 10px;font-size:12px;font-weight:600;cursor:pointer;">
        ✕ Xem lại toàn quốc
      </button>
    </div>
  `;

  const resetBtn = document.getElementById('btnResetUnitFilter');
  if (resetBtn) {
    resetBtn.addEventListener('click', () => {
      state.region = 'ALL';
      state.branch = 'ALL';
      document.getElementById('filterRegion').value = 'ALL';
      updateBranchDropdown('ALL');
      refreshDashboard();
    });
  }
}

// Render 3 hộp Bảng Vàng Thi Đua (Top 5 Chi nhánh Lũy kế / Top 5 Bứt phá Hôm qua & Hôm nay / Top 5 Khu vực)
function renderLeaderboardStrip() {
  const mtdEl = document.getElementById('topBranchesMtdList');
  const recentEl = document.getElementById('topBranchesRecentList');
  const regEl = document.getElementById('topRegionsMtdList');
  if (!mtdEl || !recentEl || !regEl) return;

  const m = getTemporalMilestones();
  const lblRecent = document.getElementById('lblRecentDates');
  if (lblRecent) lblRecent.textContent = `${formatShortDM(m.yesterdayYmd)} & ${formatShortDM(m.latestYmd)}`;
  const lblScope = document.getElementById('lblLeaderScope');
  if (lblScope) lblScope.textContent = state.region === 'ALL' ? 'Toàn quốc (78 CN)' : `Trong KV ${state.region}`;

  const brMap = {};
  ALL_BRANCHES.forEach(b => {
    if (state.region !== 'ALL' && BRANCH_TO_REGION[b] !== state.region) return;
    brMap[b] = { branch: b, region: BRANCH_TO_REGION[b], mtd: 0, images: 0, videos: 0, today: 0, yest: 0, recent: 0 };
  });

  const regMap = {};
  REGION_BRANCH_HIERARCHY.forEach(g => {
    regMap[g.region] = { region: g.region, name: g.regionName.split('(')[0].trim(), smiles: 0, activeCN: new Set(), totalCN: g.branches.length };
  });

  state.rawRecords.forEach(r => {
    if (brMap[r.branch]) {
      brMap[r.branch].mtd += r.smiles;
      brMap[r.branch].images += r.images;
      brMap[r.branch].videos += r.videos;
      if (r.ymd === m.latestYmd) brMap[r.branch].today += r.smiles;
      if (r.ymd === m.yesterdayYmd) brMap[r.branch].yest += r.smiles;
      if (r.ymd === m.latestYmd || r.ymd === m.yesterdayYmd) brMap[r.branch].recent += r.smiles;
    }
    if (regMap[r.region]) {
      regMap[r.region].smiles += r.smiles;
      if (BRANCH_TO_REGION[r.branch]) regMap[r.region].activeCN.add(r.branch);
    }
  });

  const topMtd = Object.values(brMap).sort((a, b) => b.mtd - a.mtd || b.videos - a.videos).slice(0, 5);
  const topRecent = Object.values(brMap).sort((a, b) => b.recent - a.recent || b.today - a.today || b.mtd - a.mtd).slice(0, 5);
  const topRegs = Object.values(regMap).sort((a, b) => b.smiles - a.smiles).slice(0, 5);

  const medals = ['🥇', '🥈', '🥉', '4.', '5.'];

  mtdEl.innerHTML = topMtd.map((item, idx) => `
    <div class="leader-item" data-select-branch="${item.branch}" title="Bấm để xem chi tiết ${item.branch}">
      <div>
        <span class="leader-rank top${idx + 1}">${medals[idx]}</span>
        <strong>${item.branch}</strong>
        <span class="region-tag" style="margin-left:4px;">${item.region}</span>
      </div>
      <div>
        <span style="color:#64748b;font-size:11.5px;margin-right:6px;">${item.images} ảnh • ${item.videos} video</span>
        <span class="smile-pill">${item.mtd} 😊</span>
      </div>
    </div>
  `).join('');

  recentEl.innerHTML = topRecent.map((item, idx) => `
    <div class="leader-item" data-select-branch="${item.branch}" title="Bấm để xem chi tiết ${item.branch}">
      <div>
        <span class="leader-rank top${idx + 1}">${medals[idx]}</span>
        <strong>${item.branch}</strong>
        <span class="region-tag" style="margin-left:4px;">${item.region}</span>
      </div>
      <div>
        <span style="color:#64748b;font-size:11.5px;margin-right:6px;">H.qua: +${item.yest} | H.nay: +${item.today}</span>
        <span class="delta-up">+${item.recent} 😊</span>
      </div>
    </div>
  `).join('');

  regEl.innerHTML = topRegs.map((item, idx) => {
    const rate = Math.round((item.activeCN.size / item.totalCN) * 100);
    return `
      <div class="leader-item" data-select-region="${item.region}" title="Bấm để lọc Khu vực ${item.region}">
        <div>
          <span class="leader-rank top${idx + 1}">${medals[idx]}</span>
          <strong>${item.region}</strong>
          <span style="color:#64748b;font-size:11.5px;margin-left:4px;">(${item.activeCN.size}/${item.totalCN} CN - ${rate}%)</span>
        </div>
        <div>
          <span class="smile-pill">${item.smiles} 😊</span>
        </div>
      </div>
    `;
  }).join('');

  // Cho phép bấm vào bất kỳ dòng nào trên Bảng Vàng để lọc nhanh đơn vị đó
  document.querySelectorAll('[data-select-branch]').forEach(el => {
    el.addEventListener('click', () => {
      const br = el.getAttribute('data-select-branch');
      const reg = BRANCH_TO_REGION[br] || 'ALL';
      state.region = reg;
      document.getElementById('filterRegion').value = reg;
      updateBranchDropdown(reg);
      state.branch = br;
      document.getElementById('filterBranch').value = br;
      refreshDashboard();
    });
  });

  document.querySelectorAll('[data-select-region]').forEach(el => {
    el.addEventListener('click', () => {
      const reg = el.getAttribute('data-select-region');
      state.region = reg;
      state.branch = 'ALL';
      document.getElementById('filterRegion').value = reg;
      updateBranchDropdown(reg);
      refreshDashboard();
    });
  });
}

// TAB 1: Render Bảng Chi nhánh & Khu vực có đầy đủ Hạng, Hôm kia, Hôm qua, So sánh, Hôm nay, Tuần này, Lũy kế & từng ngày
function renderBranchAndRegionTable(timeBuckets) {
  const headEl = document.getElementById('branchTableHead');
  const bodyEl = document.getElementById('branchTableBody');
  const showDynamic = state.showTimeColumns && timeBuckets.length > 0;
  const m = getTemporalMilestones();

  const grainTitle = state.timeGrain === 'day' ? 'Chi Tiết Từng Ngày (Nụ Cười)' :
                     state.timeGrain === 'week' ? 'Chi Tiết Từng Tuần (Nụ Cười)' :
                     'Chi Tiết Từng Tháng (Nụ Cười)';

  let headHtml = `
    <tr>
      <th rowspan="2" class="sticky-col-1">STT</th>
      <th rowspan="2" class="sticky-col-2">Đơn vị / Chi nhánh</th>
      <th rowspan="2">Khu vực</th>
      <th rowspan="2">Hạng</th>
      <th colspan="5" style="background:#1e3a8a;color:#fff;">Tiến Độ Hôm Qua • Hôm Nay • Tuần Này</th>
      <th colspan="4">Lũy Kế Trong Kỳ Đang Chọn</th>
      <th rowspan="2">Tỷ trọng</th>
      ${showDynamic ? `<th colspan="${timeBuckets.length}">${grainTitle}</th>` : ''}
    </tr>
    <tr class="sub-header">
      <th style="background:#eff6ff;color:#1e3a8a;">Hôm kia (${formatShortDM(m.prevDayYmd)})</th>
      <th style="background:#eff6ff;color:#1e3a8a;">Hôm qua (${formatShortDM(m.yesterdayYmd)})</th>
      <th style="background:#eff6ff;color:#1e3a8a;">So với ${formatShortDM(m.prevDayYmd)}</th>
      <th style="background:#ecfdf5;color:#047857;">Hôm nay (${formatShortDM(m.latestYmd)})</th>
      <th style="background:#f5f3ff;color:#5b21b6;">Tuần này (${m.currentWeekKey.split('-')[1]})</th>
      <th>Hình ảnh (×3)</th>
      <th>Video (×5)</th>
      <th>Tổng bài</th>
      <th>Tổng Nụ Cười</th>
      ${showDynamic ? timeBuckets.map(b => `<th>${b.label}</th>`).join('') : ''}
    </tr>
  `;
  headEl.innerHTML = headHtml;

  // Khởi tạo bộ đếm cho 78 Chi nhánh
  const branchStats = {};
  ALL_BRANCHES.forEach(b => {
    branchStats[b] = {
      branch: b,
      region: BRANCH_TO_REGION[b],
      prevDay: 0,
      yesterday: 0,
      today: 0,
      thisWeek: 0,
      images: 0,
      videos: 0,
      posts: 0,
      smiles: 0,
      byBucket: {}
    };
  });

  // Tính các cột Hôm kia / Hôm qua / Hôm nay / Tuần này từ dữ liệu gốc (lọc theo Bộ phận / Loại nếu có)
  state.rawRecords.forEach(r => {
    if (state.dept !== 'ALL' && r.dept !== state.dept) return;
    if (state.mediaType !== 'ALL' && r.mediaCode !== state.mediaType) return;
    if (!branchStats[r.branch]) return;
    const st = branchStats[r.branch];
    if (r.ymd === m.prevDayYmd) st.prevDay += r.smiles;
    if (r.ymd === m.yesterdayYmd) st.yesterday += r.smiles;
    if (r.ymd === m.latestYmd) st.today += r.smiles;
    if (r.weekKey === m.currentWeekKey) st.thisWeek += r.smiles;
  });

  let grandImages = 0, grandVideos = 0, grandPosts = 0, grandSmiles = 0;
  let grandPrevDay = 0, grandYesterday = 0, grandToday = 0, grandThisWeek = 0;
  const grandByBucket = {};

  state.filteredRecords.forEach(r => {
    if (!branchStats[r.branch]) {
      branchStats[r.branch] = {
        branch: r.branch,
        region: r.region,
        prevDay: 0, yesterday: 0, today: 0, thisWeek: 0,
        images: 0, videos: 0, posts: 0, smiles: 0, byBucket: {}
      };
    }
    const st = branchStats[r.branch];
    const bKey = getRecordBucketKey(r);

    st.images += r.images;
    st.videos += r.videos;
    st.posts += 1;
    st.smiles += r.smiles;
    st.byBucket[bKey] = (st.byBucket[bKey] || 0) + r.smiles;

    grandImages += r.images;
    grandVideos += r.videos;
    grandPosts += 1;
    grandSmiles += r.smiles;
    grandByBucket[bKey] = (grandByBucket[bKey] || 0) + r.smiles;
  });

  // Tính thứ hạng (Rank 1 -> 78) dựa trên Tổng Nụ Cười trong kỳ
  const rankedList = ALL_BRANCHES.slice().sort((a, b) => {
    const diff = branchStats[b].smiles - branchStats[a].smiles;
    if (diff !== 0) return diff;
    return (branchStats[b].today + branchStats[b].yesterday) - (branchStats[a].today + branchStats[a].yesterday);
  });
  const rankMap = {};
  rankedList.forEach((b, idx) => {
    rankMap[b] = branchStats[b].smiles > 0 ? idx + 1 : 0;
  });

  const kw = normalizeKey(state.keyword);
  let rowsHtml = '';

  function renderSingleBranchRow(branch) {
    const st = branchStats[branch];
    const groupRegion = BRANCH_TO_REGION[branch] || st.region;
    const sharePct = grandSmiles > 0 ? ((st.smiles / grandSmiles) * 100).toFixed(1) : '0.0';
    const stt = BRANCH_STT_MAP[branch] || '-';
    const diffYest = st.yesterday - st.prevDay;

    return `
      <tr>
        <td class="sticky-col-1">${stt}</td>
        <td class="sticky-col-2"><strong>${branch}</strong></td>
        <td class="center-col"><span class="region-tag">${groupRegion}</span></td>
        <td class="center-col">${formatRankBadge(rankMap[branch])}</td>
        <td class="num-col ${st.prevDay === 0 ? 'zero-val' : ''}">${st.prevDay}</td>
        <td class="num-col ${st.yesterday === 0 ? 'zero-val' : ''}" style="font-weight:700;color:#1e40af;">${st.yesterday > 0 ? '+' + st.yesterday : '0'}</td>
        <td class="center-col">${formatDeltaHtml(diffYest)}</td>
        <td class="num-col ${st.today === 0 ? 'zero-val' : ''}" style="font-weight:700;color:#047857;">${st.today > 0 ? '+' + st.today : '0'}</td>
        <td class="num-col ${st.thisWeek === 0 ? 'zero-val' : ''}" style="font-weight:600;">${st.thisWeek}</td>
        <td class="num-col ${st.images === 0 ? 'zero-val' : ''}">${st.images}</td>
        <td class="num-col ${st.videos === 0 ? 'zero-val' : ''}">${st.videos}</td>
        <td class="num-col ${st.posts === 0 ? 'zero-val' : ''}">${st.posts}</td>
        <td class="num-col"><span class="smile-pill">${st.smiles}</span></td>
        <td class="num-col">${sharePct}%</td>
        ${showDynamic ? timeBuckets.map(b => {
          const v = st.byBucket[b.key] || 0;
          return `<td class="num-col ${v === 0 ? 'zero-val' : ''}">${v === 0 ? '-' : v}</td>`;
        }).join('') : ''}
      </tr>
    `;
  }

  if (state.branchSort === 'rank_desc' || state.branchSort === 'today_desc') {
    const flatBranches = ALL_BRANCHES.filter(branch => {
      const reg = BRANCH_TO_REGION[branch];
      if (state.region !== 'ALL' && reg !== state.region) return false;
      if (state.branch !== 'ALL' && branch !== state.branch) return false;
      if (kw && !normalizeKey(`${branch} ${reg}`).includes(kw)) return false;
      const st = branchStats[branch];
      if (!state.showZeroBranches && st.smiles === 0) return false;
      return true;
    });

    if (state.branchSort === 'rank_desc') {
      flatBranches.sort((a, b) => branchStats[b].smiles - branchStats[a].smiles || branchStats[b].today - branchStats[a].today);
    } else {
      flatBranches.sort((a, b) => (branchStats[b].today + branchStats[b].yesterday) - (branchStats[a].today + branchStats[a].yesterday) || branchStats[b].smiles - branchStats[a].smiles);
    }

    flatBranches.forEach(branch => {
      const st = branchStats[branch];
      grandPrevDay += st.prevDay;
      grandYesterday += st.yesterday;
      grandToday += st.today;
      grandThisWeek += st.thisWeek;
      rowsHtml += renderSingleBranchRow(branch);
    });
  } else {
    REGION_BRANCH_HIERARCHY.forEach(group => {
      if (state.region !== 'ALL' && group.region !== state.region) return;

      let regImages = 0, regVideos = 0, regPosts = 0, regSmiles = 0;
      let regPrevDay = 0, regYesterday = 0, regToday = 0, regThisWeek = 0;
      let regActiveBranches = 0;
      const regByBucket = {};
      let groupRowsHtml = '';

      group.branches.forEach(branch => {
        if (state.branch !== 'ALL' && branch !== state.branch) return;
        if (kw && !normalizeKey(`${branch} ${group.region}`).includes(kw)) return;

        const st = branchStats[branch];
        regPrevDay += st.prevDay;
        regYesterday += st.yesterday;
        regToday += st.today;
        regThisWeek += st.thisWeek;
        regImages += st.images;
        regVideos += st.videos;
        regPosts += st.posts;
        regSmiles += st.smiles;
        if (st.smiles > 0) regActiveBranches++;

        grandPrevDay += st.prevDay;
        grandYesterday += st.yesterday;
        grandToday += st.today;
        grandThisWeek += st.thisWeek;

        timeBuckets.forEach(b => {
          const val = st.byBucket[b.key] || 0;
          regByBucket[b.key] = (regByBucket[b.key] || 0) + val;
        });

        if (!state.showZeroBranches && st.smiles === 0) return;
        groupRowsHtml += renderSingleBranchRow(branch);
      });

      if (groupRowsHtml || (state.showZeroBranches && state.branch === 'ALL' && !kw)) {
        const regShare = grandSmiles > 0 ? ((regSmiles / grandSmiles) * 100).toFixed(1) : '0.0';
        const regDiffYest = regYesterday - regPrevDay;
        rowsHtml += groupRowsHtml;
        rowsHtml += `
          <tr class="region-summary-row">
            <td class="sticky-col-1" style="background:var(--region-row-bg)">KV</td>
            <td class="sticky-col-2" style="background:var(--region-row-bg)">${group.region} (${regActiveBranches}/${group.branches.length} CN)</td>
            <td class="center-col">${group.region}</td>
            <td class="center-col">--</td>
            <td class="num-col">${regPrevDay}</td>
            <td class="num-col">+${regYesterday}</td>
            <td class="center-col">${formatDeltaHtml(regDiffYest)}</td>
            <td class="num-col">+${regToday}</td>
            <td class="num-col">${regThisWeek}</td>
            <td class="num-col">${regImages.toLocaleString('vi-VN')}</td>
            <td class="num-col">${regVideos.toLocaleString('vi-VN')}</td>
            <td class="num-col">${regPosts.toLocaleString('vi-VN')}</td>
            <td class="num-col"><span class="smile-pill">${regSmiles.toLocaleString('vi-VN')}</span></td>
            <td class="num-col">${regShare}%</td>
            ${showDynamic ? timeBuckets.map(b => {
              const v = regByBucket[b.key] || 0;
              return `<td class="num-col">${v}</td>`;
            }).join('') : ''}
          </tr>
        `;
      }
    });
  }

  // Dòng Tổng Cộng Toàn Hệ Thống
  const grandDiffYest = grandYesterday - grandPrevDay;
  rowsHtml += `
    <tr class="grand-total-row">
      <td class="sticky-col-1" style="background:var(--total-row-bg);color:#fff">∑</td>
      <td class="sticky-col-2" style="background:var(--total-row-bg);color:#fff">TỔNG TOÀN HỆ THỐNG</td>
      <td class="center-col">${REGION_BRANCH_HIERARCHY.length} KV</td>
      <td class="center-col">--</td>
      <td class="num-col">${grandPrevDay.toLocaleString('vi-VN')}</td>
      <td class="num-col">+${grandYesterday.toLocaleString('vi-VN')}</td>
      <td class="center-col">${grandDiffYest >= 0 ? '+' + grandDiffYest : grandDiffYest}</td>
      <td class="num-col">+${grandToday.toLocaleString('vi-VN')}</td>
      <td class="num-col">${grandThisWeek.toLocaleString('vi-VN')}</td>
      <td class="num-col">${grandImages.toLocaleString('vi-VN')}</td>
      <td class="num-col">${grandVideos.toLocaleString('vi-VN')}</td>
      <td class="num-col">${grandPosts.toLocaleString('vi-VN')}</td>
      <td class="num-col">${grandSmiles.toLocaleString('vi-VN')}</td>
      <td class="num-col">100%</td>
      ${showDynamic ? timeBuckets.map(b => {
        const v = grandByBucket[b.key] || 0;
        return `<td class="num-col">${v.toLocaleString('vi-VN')}</td>`;
      }).join('') : ''}
    </tr>
  `;

  bodyEl.innerHTML = rowsHtml;
}

// TAB 2: Render Bảng Tổng hợp 15 Khu vực (Có Hạng, Hôm qua, Hôm nay, Tuần này, Lũy kế)
function renderRegionSummaryTable(timeBuckets) {
  const headEl = document.getElementById('regionTableHead');
  const bodyEl = document.getElementById('regionTableBody');
  const m = getTemporalMilestones();

  headEl.innerHTML = `
    <tr>
      <th class="sticky-col-1">Hạng</th>
      <th class="sticky-col-2">Mã Khu Vực</th>
      <th>Tên Khu Vực</th>
      <th>Số CN Có Bài</th>
      <th>Tỷ Lệ Tham Gia</th>
      <th style="background:#eff6ff;color:#1e3a8a;">Hôm qua (${formatShortDM(m.yesterdayYmd)})</th>
      <th style="background:#ecfdf5;color:#047857;">Hôm nay (${formatShortDM(m.latestYmd)})</th>
      <th style="background:#f5f3ff;color:#5b21b6;">Tuần này (${m.currentWeekKey.split('-')[1]})</th>
      <th>Hình Ảnh (×3)</th>
      <th>Video (×5)</th>
      <th>Tổng Bài Đăng</th>
      <th>Tổng Nụ Cười</th>
      <th>TB / CN</th>
      <th>Tỷ Trọng</th>
      ${timeBuckets.map(b => `<th>${b.label}</th>`).join('')}
    </tr>
  `;

  let grandSmiles = 0, grandImages = 0, grandVideos = 0, grandPosts = 0;
  let grandYesterday = 0, grandToday = 0, grandThisWeek = 0;
  const grandActiveBranches = new Set();
  const grandByBucket = {};

  const regionRows = REGION_BRANCH_HIERARCHY.map((g, idx) => {
    let images = 0, videos = 0, posts = 0, smiles = 0;
    let yesterday = 0, today = 0, thisWeek = 0;
    const activeSet = new Set();
    const byBucket = {};

    state.rawRecords.forEach(r => {
      if (r.region !== g.region) return;
      if (state.dept !== 'ALL' && r.dept !== state.dept) return;
      if (state.mediaType !== 'ALL' && r.mediaCode !== state.mediaType) return;
      if (r.ymd === m.yesterdayYmd) yesterday += r.smiles;
      if (r.ymd === m.latestYmd) today += r.smiles;
      if (r.weekKey === m.currentWeekKey) thisWeek += r.smiles;
    });

    state.filteredRecords.forEach(r => {
      if (r.region === g.region) {
        images += r.images;
        videos += r.videos;
        posts += 1;
        smiles += r.smiles;
        activeSet.add(r.branch);
        grandActiveBranches.add(r.branch);
        const bKey = getRecordBucketKey(r);
        byBucket[bKey] = (byBucket[bKey] || 0) + r.smiles;
        grandByBucket[bKey] = (grandByBucket[bKey] || 0) + r.smiles;
      }
    });

    grandSmiles += smiles;
    grandImages += images;
    grandVideos += videos;
    grandPosts += posts;
    grandYesterday += yesterday;
    grandToday += today;
    grandThisWeek += thisWeek;

    const totalBranches = g.branches.length;
    const rate = totalBranches > 0 ? (activeSet.size / totalBranches) * 100 : 0;
    const avgPerBranch = totalBranches > 0 ? smiles / totalBranches : 0;

    return {
      excelOrder: idx + 1,
      region: g.region,
      regionName: g.regionName,
      activeCount: activeSet.size,
      totalBranches,
      rate,
      yesterday,
      today,
      thisWeek,
      images,
      videos,
      posts,
      smiles,
      avgPerBranch,
      byBucket
    };
  });

  if (state.regionSort === 'smiles_desc') {
    regionRows.sort((a, b) => b.smiles - a.smiles);
  } else if (state.regionSort === 'rate_desc') {
    regionRows.sort((a, b) => b.rate - a.rate || b.smiles - a.smiles);
  }

  let html = '';
  regionRows.forEach((row, i) => {
    const share = grandSmiles > 0 ? ((row.smiles / grandSmiles) * 100).toFixed(1) : '0.0';
    html += `
      <tr>
        <td class="sticky-col-1">${formatRankBadge(i + 1)}</td>
        <td class="sticky-col-2"><strong>${row.region}</strong></td>
        <td>${row.regionName}</td>
        <td class="center-col"><strong>${row.activeCount}</strong> / ${row.totalBranches}</td>
        <td>
          <div class="progress-inline">
            <div class="progress-bar-bg">
              <div class="progress-bar-fill" style="width:${Math.min(100, row.rate).toFixed(0)}%"></div>
            </div>
            <span style="font-weight:600;font-size:12px">${row.rate.toFixed(0)}%</span>
          </div>
        </td>
        <td class="num-col" style="font-weight:700;color:#1e40af;">${row.yesterday > 0 ? '+' + row.yesterday : '0'}</td>
        <td class="num-col" style="font-weight:700;color:#047857;">${row.today > 0 ? '+' + row.today : '0'}</td>
        <td class="num-col" style="font-weight:600;">${row.thisWeek}</td>
        <td class="num-col">${row.images.toLocaleString('vi-VN')}</td>
        <td class="num-col">${row.videos.toLocaleString('vi-VN')}</td>
        <td class="num-col">${row.posts.toLocaleString('vi-VN')}</td>
        <td class="num-col"><span class="smile-pill">${row.smiles.toLocaleString('vi-VN')}</span></td>
        <td class="num-col">${row.avgPerBranch.toFixed(1)}</td>
        <td class="num-col">${share}%</td>
        ${timeBuckets.map(b => {
          const v = row.byBucket[b.key] || 0;
          return `<td class="num-col ${v === 0 ? 'zero-val' : ''}">${v === 0 ? '-' : v}</td>`;
        }).join('')}
      </tr>
    `;
  });

  const totalRate = ((grandActiveBranches.size / ALL_BRANCHES.length) * 100).toFixed(1);
  html += `
    <tr class="grand-total-row">
      <td class="sticky-col-1" style="background:var(--total-row-bg);color:#fff">∑</td>
      <td class="sticky-col-2" style="background:var(--total-row-bg);color:#fff">TỔNG ${REGION_BRANCH_HIERARCHY.length} KHU VỰC</td>
      <td>Toàn hệ thống</td>
      <td class="center-col">${grandActiveBranches.size} / ${ALL_BRANCHES.length}</td>
      <td class="center-col">${totalRate}%</td>
      <td class="num-col">+${grandYesterday.toLocaleString('vi-VN')}</td>
      <td class="num-col">+${grandToday.toLocaleString('vi-VN')}</td>
      <td class="num-col">${grandThisWeek.toLocaleString('vi-VN')}</td>
      <td class="num-col">${grandImages.toLocaleString('vi-VN')}</td>
      <td class="num-col">${grandVideos.toLocaleString('vi-VN')}</td>
      <td class="num-col">${grandPosts.toLocaleString('vi-VN')}</td>
      <td class="num-col">${grandSmiles.toLocaleString('vi-VN')}</td>
      <td class="num-col">${(grandSmiles / ALL_BRANCHES.length).toFixed(1)}</td>
      <td class="num-col">100%</td>
      ${timeBuckets.map(b => `<td class="num-col">${(grandByBucket[b.key] || 0).toLocaleString('vi-VN')}</td>`).join('')}
    </tr>
  `;

  bodyEl.innerHTML = html;
}

// TAB 3: Thống kê chuyên sâu theo "Bộ phận" (Bảng 3A: Xếp hạng Bộ phận + Bảng 3B: Ma trận Đơn vị x Bộ phận)
function renderDepartmentTables(timeBuckets) {
  const deptHead = document.getElementById('deptTableHead');
  const deptBody = document.getElementById('deptTableBody');

  const deptMap = {};
  let totalSmiles = 0, totalImages = 0, totalVideos = 0, totalPosts = 0;
  const grandByBucket = {};

  state.filteredRecords.forEach(r => {
    if (!deptMap[r.dept]) {
      deptMap[r.dept] = {
        dept: r.dept,
        images: 0,
        videos: 0,
        posts: 0,
        smiles: 0,
        branches: new Set(),
        regions: new Set(),
        byBucket: {}
      };
    }
    const d = deptMap[r.dept];
    d.images += r.images;
    d.videos += r.videos;
    d.posts += 1;
    d.smiles += r.smiles;
    d.branches.add(r.branch);
    d.regions.add(r.region);
    const bKey = getRecordBucketKey(r);
    d.byBucket[bKey] = (d.byBucket[bKey] || 0) + r.smiles;
    grandByBucket[bKey] = (grandByBucket[bKey] || 0) + r.smiles;

    totalSmiles += r.smiles;
    totalImages += r.images;
    totalVideos += r.videos;
    totalPosts += 1;
  });

  const deptList = Object.values(deptMap).sort((a, b) => b.smiles - a.smiles);

  deptHead.innerHTML = `
    <tr>
      <th class="sticky-col-1">STT</th>
      <th class="sticky-col-2">Bộ Phận (Cột: Bộ phận)</th>
      <th>Số CN Tham Gia</th>
      <th>Số KV Tham Gia</th>
      <th>Hình Ảnh (×3)</th>
      <th>Video (×5)</th>
      <th>Tổng Bài Đăng</th>
      <th>Tổng Nụ Cười</th>
      <th>Tỷ Trọng (%)</th>
      ${timeBuckets.map(b => `<th>${b.label}</th>`).join('')}
    </tr>
  `;

  let htmlA = '';
  deptList.forEach((d, idx) => {
    const share = totalSmiles > 0 ? ((d.smiles / totalSmiles) * 100).toFixed(1) : '0.0';
    htmlA += `
      <tr>
        <td class="sticky-col-1">${idx + 1}</td>
        <td class="sticky-col-2"><strong>${d.dept}</strong></td>
        <td class="center-col">${d.branches.size} CN</td>
        <td class="center-col">${d.regions.size} KV</td>
        <td class="num-col">${d.images.toLocaleString('vi-VN')}</td>
        <td class="num-col">${d.videos.toLocaleString('vi-VN')}</td>
        <td class="num-col">${d.posts.toLocaleString('vi-VN')}</td>
        <td class="num-col"><span class="smile-pill">${d.smiles.toLocaleString('vi-VN')}</span></td>
        <td class="num-col">${share}%</td>
        ${timeBuckets.map(b => {
          const v = d.byBucket[b.key] || 0;
          return `<td class="num-col ${v === 0 ? 'zero-val' : ''}">${v === 0 ? '-' : v}</td>`;
        }).join('')}
      </tr>
    `;
  });

  htmlA += `
    <tr class="grand-total-row">
      <td class="sticky-col-1" style="background:var(--total-row-bg);color:#fff">∑</td>
      <td class="sticky-col-2" style="background:var(--total-row-bg);color:#fff">TỔNG TẤT CẢ BỘ PHẬN (${deptList.length})</td>
      <td class="center-col">--</td>
      <td class="center-col">--</td>
      <td class="num-col">${totalImages.toLocaleString('vi-VN')}</td>
      <td class="num-col">${totalVideos.toLocaleString('vi-VN')}</td>
      <td class="num-col">${totalPosts.toLocaleString('vi-VN')}</td>
      <td class="num-col">${totalSmiles.toLocaleString('vi-VN')}</td>
      <td class="num-col">100%</td>
      ${timeBuckets.map(b => `<td class="num-col">${(grandByBucket[b.key] || 0).toLocaleString('vi-VN')}</td>`).join('')}
    </tr>
  `;
  deptBody.innerHTML = htmlA;

  // Bảng 3B: Ma trận Đơn vị x Bộ phận
  const matrixHead = document.getElementById('deptMatrixHead');
  const matrixBody = document.getElementById('deptMatrixBody');
  const deptNames = deptList.map(d => d.dept);

  matrixHead.innerHTML = `
    <tr>
      <th class="sticky-col-1">STT</th>
      <th class="sticky-col-2">Đơn Vị / Khu Vực</th>
      <th>Khu Vực</th>
      <th>Tổng Nụ Cười</th>
      ${deptNames.map(name => `<th>${name}</th>`).join('')}
    </tr>
  `;

  // Tính ma trận branch -> dept -> smiles
  const branchDeptSmiles = {};
  state.filteredRecords.forEach(r => {
    if (!branchDeptSmiles[r.branch]) branchDeptSmiles[r.branch] = {};
    branchDeptSmiles[r.branch][r.dept] = (branchDeptSmiles[r.branch][r.dept] || 0) + r.smiles;
  });

  let htmlB = '';
  let regionStt = 1;
  REGION_BRANCH_HIERARCHY.forEach(group => {
    if (state.region !== 'ALL' && group.region !== state.region) return;

    const regDeptTotals = {};
    let regTotalSmiles = 0;

    group.branches.forEach(branch => {
      const bMap = branchDeptSmiles[branch] || {};
      let bTotal = 0;
      deptNames.forEach(dn => {
        const val = bMap[dn] || 0;
        bTotal += val;
        regDeptTotals[dn] = (regDeptTotals[dn] || 0) + val;
      });
      regTotalSmiles += bTotal;

      if (state.deptMatrixLevel === 'branch_and_region') {
        if (!state.showZeroBranches && bTotal === 0) return;
        htmlB += `
          <tr>
            <td class="sticky-col-1">${BRANCH_STT_MAP[branch]}</td>
            <td class="sticky-col-2">${branch}</td>
            <td class="center-col"><span class="region-tag">${group.region}</span></td>
            <td class="num-col"><span class="smile-pill">${bTotal}</span></td>
            ${deptNames.map(dn => {
              const v = bMap[dn] || 0;
              return `<td class="num-col ${v === 0 ? 'zero-val' : ''}">${v === 0 ? '-' : v}</td>`;
            }).join('')}
          </tr>
        `;
      }
    });

    if (state.deptMatrixLevel === 'region_only') {
      htmlB += `
        <tr>
          <td class="sticky-col-1">${regionStt++}</td>
          <td class="sticky-col-2"><strong>${group.region} - ${group.regionName}</strong></td>
          <td class="center-col"><span class="region-tag">${group.region}</span></td>
          <td class="num-col"><span class="smile-pill">${regTotalSmiles}</span></td>
          ${deptNames.map(dn => {
            const v = regDeptTotals[dn] || 0;
            return `<td class="num-col ${v === 0 ? 'zero-val' : ''}">${v === 0 ? '-' : v}</td>`;
          }).join('')}
        </tr>
      `;
    } else {
      htmlB += `
        <tr class="region-summary-row">
          <td class="sticky-col-1" style="background:var(--region-row-bg)">KV</td>
          <td class="sticky-col-2" style="background:var(--region-row-bg)">Tổng ${group.region}</td>
          <td class="center-col">${group.region}</td>
          <td class="num-col"><span class="smile-pill">${regTotalSmiles}</span></td>
          ${deptNames.map(dn => `<td class="num-col">${regDeptTotals[dn] || 0}</td>`).join('')}
        </tr>
      `;
    }
  });

  matrixBody.innerHTML = htmlB;
}

// TAB 4: Render 4 Biểu đồ phân tích
function renderCharts(timeBuckets) {
  if (typeof Chart === 'undefined') return;

  // 1. Time Trend Chart
  const labels = timeBuckets.map(b => b.label);
  const imgSmilesArr = [];
  const vidSmilesArr = [];
  const totalSmilesArr = [];

  timeBuckets.forEach(b => {
    let imgS = 0, vidS = 0;
    state.filteredRecords.forEach(r => {
      if (getRecordBucketKey(r) === b.key) {
        imgS += r.images * 3;
        vidS += r.videos * 5;
      }
    });
    imgSmilesArr.push(imgS);
    vidSmilesArr.push(vidS);
    totalSmilesArr.push(imgS + vidS);
  });

  if (state.charts.timeTrend) state.charts.timeTrend.destroy();
  const ctxTrend = document.getElementById('chartTimeTrend').getContext('2d');
  state.charts.timeTrend = new Chart(ctxTrend, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {
          type: 'line',
          label: 'Tổng Nụ Cười',
          data: totalSmilesArr,
          borderColor: '#2563eb',
          backgroundColor: '#2563eb',
          borderWidth: 2.5,
          tension: 0.25,
          yAxisID: 'y'
        },
        {
          type: 'bar',
          label: 'Nụ cười từ Hình ảnh (×3)',
          data: imgSmilesArr,
          backgroundColor: '#10b981',
          stack: 'smiles'
        },
        {
          type: 'bar',
          label: 'Nụ cười từ Video (×5)',
          data: vidSmilesArr,
          backgroundColor: '#f59e0b',
          stack: 'smiles'
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'top' } },
      scales: {
        x: { stacked: true },
        y: { stacked: true, beginAtZero: true }
      }
    }
  });

  // 2. Region Chart (16 Khu vực)
  const regionLabels = REGION_BRANCH_HIERARCHY.map(g => g.region);
  const regImgSmiles = [];
  const regVidSmiles = [];

  REGION_BRANCH_HIERARCHY.forEach(g => {
    let iS = 0, vS = 0;
    state.filteredRecords.forEach(r => {
      if (r.region === g.region) {
        iS += r.images * 3;
        vS += r.videos * 5;
      }
    });
    regImgSmiles.push(iS);
    regVidSmiles.push(vS);
  });

  if (state.charts.regionBar) state.charts.regionBar.destroy();
  const ctxReg = document.getElementById('chartRegionBar').getContext('2d');
  state.charts.regionBar = new Chart(ctxReg, {
    type: 'bar',
    data: {
      labels: regionLabels,
      datasets: [
        { label: 'Hình ảnh (×3)', data: regImgSmiles, backgroundColor: '#3b82f6', stack: 'reg' },
        { label: 'Video (×5)', data: regVidSmiles, backgroundColor: '#f59e0b', stack: 'reg' }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'top' } },
      scales: {
        x: { stacked: true },
        y: { stacked: true, beginAtZero: true }
      }
    }
  });

  // 3. Department Chart
  const deptTotals = {};
  state.filteredRecords.forEach(r => {
    deptTotals[r.dept] = (deptTotals[r.dept] || 0) + r.smiles;
  });
  const sortedDepts = Object.entries(deptTotals).sort((a, b) => b[1] - a[1]);

  if (state.charts.deptBar) state.charts.deptBar.destroy();
  const ctxDept = document.getElementById('chartDeptBar').getContext('2d');
  state.charts.deptBar = new Chart(ctxDept, {
    type: 'bar',
    data: {
      labels: sortedDepts.map(d => d[0]),
      datasets: [{
        label: 'Tổng Nụ Cười',
        data: sortedDepts.map(d => d[1]),
        backgroundColor: ['#2563eb', '#10b981', '#f59e0b', '#7c3aed', '#e11d48', '#06b6d4', '#64748b']
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { x: { beginAtZero: true } }
    }
  });

  // 4. Top 15 Branches Chart
  const branchTotals = {};
  state.filteredRecords.forEach(r => {
    branchTotals[r.branch] = (branchTotals[r.branch] || 0) + r.smiles;
  });
  const top15Branches = Object.entries(branchTotals).sort((a, b) => b[1] - a[1]).slice(0, 15);

  if (state.charts.topBranches) state.charts.topBranches.destroy();
  const ctxTop = document.getElementById('chartTopBranches').getContext('2d');
  state.charts.topBranches = new Chart(ctxTop, {
    type: 'bar',
    data: {
      labels: top15Branches.map(b => `${b[0]} (${BRANCH_TO_REGION[b[0]] || ''})`),
      datasets: [{
        label: 'Điểm Nụ Cười',
        data: top15Branches.map(b => b[1]),
        backgroundColor: '#7c3aed'
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { x: { beginAtZero: true } }
    }
  });
}

// TAB 5: Danh sách bài đăng chi tiết
function renderRawDataTable() {
  const bodyEl = document.getElementById('rawTableBody');
  const rows = state.filteredRecords.slice(0, 500);

  let html = '';
  rows.forEach((r, idx) => {
    const typeClass = r.mediaCode === 'VIDEO' ? 'video' : 'image';
    const typeIcon = r.mediaCode === 'VIDEO' ? '🎬 Video (×5)' : '📸 Hình ảnh (×3)';
    const branchDisplay = `<strong>${r.branch}</strong>`;
    const emailDisplay = r.authorEmail
      ? `<a href="mailto:${r.authorEmail}" style="color:#2563eb;text-decoration:none;font-weight:600;">${r.authorEmail}</a>`
      : '--';
    const authorDisplay = r.authorName ? `<strong>${r.authorName}</strong>` : '--';
    const titleDisplay = r.postTitle
      ? `<div style="max-width:320px;overflow:hidden;text-overflow:ellipsis;" title="${r.postTitle.replace(/"/g, '&quot;')}">${r.postTitle}</div>`
      : '--';

    html += `
      <tr>
        <td class="sticky-col-1">${idx + 1}</td>
        <td>${r.rawTimestamp}</td>
        <td class="center-col">${r.dmy}</td>
        <td class="center-col">${r.weekKey}</td>
        <td>${authorDisplay}</td>
        <td>${emailDisplay}</td>
        <td>${titleDisplay}</td>
        <td>${branchDisplay}</td>
        <td class="center-col"><span class="region-tag">${r.region}</span></td>
        <td><strong>${r.dept}</strong></td>
        <td class="center-col"><span class="type-badge ${typeClass}">${typeIcon}</span></td>
        <td class="num-col"><span class="smile-pill">+${r.smiles}</span></td>
      </tr>
    `;
  });

  bodyEl.innerHTML = html;
}

// Xuất báo cáo Excel (.xlsx) nhiều Sheet chuẩn quản trị
function exportExcelReport() {
  if (typeof XLSX === 'undefined') {
    alert('Thư viện xuất Excel đang tải, vui lòng thử lại sau giây lát.');
    return;
  }

  const wb = XLSX.utils.book_new();
  const timeBuckets = getTimeBuckets(state.filteredRecords);

  // Sheet 1: Chi nhánh & Khu vực (Chuẩn Chi nhánh.xlsx, có STT)
  const ws1Data = [];
  const header1 = [
    'STT', 'Đơn vị / Chi nhánh', 'Khu vực',
    'Hình ảnh (x3)', 'Video (x5)', 'Tổng bài đăng', 'Tổng Nụ Cười',
    ...timeBuckets.map(b => b.label)
  ];
  ws1Data.push(header1);

  const branchStats = {};
  ALL_BRANCHES.forEach(b => {
    branchStats[b] = { images: 0, videos: 0, posts: 0, smiles: 0, byBucket: {} };
  });
  state.filteredRecords.forEach(r => {
    if (!branchStats[r.branch]) {
      branchStats[r.branch] = { images: 0, videos: 0, posts: 0, smiles: 0, byBucket: {} };
    }
    const st = branchStats[r.branch];
    const bKey = getRecordBucketKey(r);
    st.images += r.images;
    st.videos += r.videos;
    st.posts += 1;
    st.smiles += r.smiles;
    st.byBucket[bKey] = (st.byBucket[bKey] || 0) + r.smiles;
  });

  REGION_BRANCH_HIERARCHY.forEach(group => {
    let regImg = 0, regVid = 0, regPosts = 0, regSmiles = 0;
    const regBuckets = {};

    group.branches.forEach(branch => {
      const st = branchStats[branch];
      regImg += st.images;
      regVid += st.videos;
      regPosts += st.posts;
      regSmiles += st.smiles;
      timeBuckets.forEach(b => {
        regBuckets[b.key] = (regBuckets[b.key] || 0) + (st.byBucket[b.key] || 0);
      });

      ws1Data.push([
        BRANCH_STT_MAP[branch],
        branch,
        group.region,
        st.images,
        st.videos,
        st.posts,
        st.smiles,
        ...timeBuckets.map(b => st.byBucket[b.key] || 0)
      ]);
    });

    // Dòng tổng Khu vực giống Chi nhánh.xlsx
    ws1Data.push([
      'KV',
      group.region,
      group.region,
      regImg,
      regVid,
      regPosts,
      regSmiles,
      ...timeBuckets.map(b => regBuckets[b.key] || 0)
    ]);
  });

  const ws1 = XLSX.utils.aoa_to_sheet(ws1Data);
  XLSX.utils.book_append_sheet(wb, ws1, '1_ChiNhanh_KhuVuc');

  // Sheet 2: Tổng hợp Bộ phận
  const ws2Data = [['STT', 'Bộ phận', 'Hình ảnh (x3)', 'Video (x5)', 'Tổng bài đăng', 'Tổng Nụ Cười']];
  const deptMap = {};
  state.filteredRecords.forEach(r => {
    if (!deptMap[r.dept]) deptMap[r.dept] = { images: 0, videos: 0, posts: 0, smiles: 0 };
    deptMap[r.dept].images += r.images;
    deptMap[r.dept].videos += r.videos;
    deptMap[r.dept].posts += 1;
    deptMap[r.dept].smiles += r.smiles;
  });
  Object.entries(deptMap)
    .sort((a, b) => b[1].smiles - a[1].smiles)
    .forEach(([dept, st], idx) => {
      ws2Data.push([idx + 1, dept, st.images, st.videos, st.posts, st.smiles]);
    });
  const ws2 = XLSX.utils.aoa_to_sheet(ws2Data);
  XLSX.utils.book_append_sheet(wb, ws2, '2_TheoBoPhan');

  // Sheet 3: Dữ liệu gốc đã chuẩn hóa (Có cột Email)
  const ws3Data = [['STT', 'Thời gian đăng', 'Ngày', 'Tuần', 'Tháng', 'Email', 'Họ tên người đăng', 'Tiêu đề bài đăng', 'Mã CN gốc', 'Chi nhánh chuẩn', 'Khu vực', 'Bộ phận', 'Loại', 'Điểm Nụ Cười']];
  state.filteredRecords.forEach((r, i) => {
    ws3Data.push([i + 1, r.rawTimestamp, r.dmy, r.weekKey, r.monthKey, r.authorEmail, r.authorName, r.postTitle, r.rawBranch, r.branch, r.region, r.dept, r.mediaLabel, r.smiles]);
  });
  const ws3 = XLSX.utils.aoa_to_sheet(ws3Data);
  XLSX.utils.book_append_sheet(wb, ws3, '3_DuLieuChiTiet');

  XLSX.writeFile(wb, `Bao_Cao_Nu_Cuoi_TTC_${state.startDate}_den_${state.endDate}.xlsx`);
}

// Tạo HTML mẫu Email Báo cáo hàng ngày (Chuẩn Executive Summary gửi lúc 09:15 sáng T+1)
function generateDailyEmailHtml() {
  const allRecs = state.rawRecords.length > 0 ? state.rawRecords : state.filteredRecords;
  if (allRecs.length === 0) return '<p>Chưa có dữ liệu.</p>';

  // Xác định ngày thực tế hôm nay (T+1) và Ngày hôm qua (Ngày T chốt 09:00 sáng)
  const maxDateObj = new Date();
  const latestYmd = toYMD(maxDateObj);
  const yesterdayObj = new Date(maxDateObj.getFullYear(), maxDateObj.getMonth(), maxDateObj.getDate() - 1);
  const yesterdayYmd = toYMD(yesterdayObj);
  const currentWeekKey = getWeekInfo(maxDateObj).key;
  const currentMonthKey = getMonthInfo(maxDateObj).key;

  let mtdSmiles = 0, mtdImages = 0, mtdVideos = 0;
  let weekSmiles = 0;
  let dayTSmiles = 0, dayTImages = 0, dayTVideos = 0;
  let todaySmiles = 0;

  const activeBranchesMtd = new Set();
  const activeRegionsMtd = new Set();
  const regionStats = {};
  const deptStats = {};
  const branchStats = {};
  const missingBranchRows = [];

  REGION_BRANCH_HIERARCHY.forEach(g => {
    regionStats[g.region] = {
      region: g.region,
      name: g.regionName,
      totalCN: g.branches.length,
      activeCN: new Set(),
      dayTSmiles: 0,
      todaySmiles: 0,
      weekSmiles: 0,
      mtdImages: 0,
      mtdVideos: 0,
      mtdSmiles: 0
    };
  });

  allRecs.forEach(r => {
    if (!r.branchMatched) {
      missingBranchRows.push(r);
    }
    if (r.monthKey === currentMonthKey) {
      mtdSmiles += r.smiles;
      mtdImages += r.images;
      mtdVideos += r.videos;
      if (BRANCH_TO_REGION[r.branch]) {
        activeBranchesMtd.add(r.branch);
        activeRegionsMtd.add(r.region);
      }
    }
    if (r.weekKey === currentWeekKey) {
      weekSmiles += r.smiles;
    }
    if (r.ymd === yesterdayYmd) {
      dayTSmiles += r.smiles;
      dayTImages += r.images;
      dayTVideos += r.videos;
    }
    if (r.ymd === latestYmd) {
      todaySmiles += r.smiles;
    }

    if (regionStats[r.region]) {
      const st = regionStats[r.region];
      if (r.monthKey === currentMonthKey) {
        st.activeCN.add(r.branch);
        st.mtdImages += r.images;
        st.mtdVideos += r.videos;
        st.mtdSmiles += r.smiles;
      }
      if (r.weekKey === currentWeekKey) st.weekSmiles += r.smiles;
      if (r.ymd === yesterdayYmd) st.dayTSmiles += r.smiles;
      if (r.ymd === latestYmd) st.todaySmiles += r.smiles;
    }

    if (!deptStats[r.dept]) {
      deptStats[r.dept] = { dayTSmiles: 0, weekSmiles: 0, mtdImages: 0, mtdVideos: 0, mtdSmiles: 0 };
    }
    if (r.monthKey === currentMonthKey) {
      deptStats[r.dept].mtdImages += r.images;
      deptStats[r.dept].mtdVideos += r.videos;
      deptStats[r.dept].mtdSmiles += r.smiles;
    }
    if (r.weekKey === currentWeekKey) deptStats[r.dept].weekSmiles += r.smiles;
    if (r.ymd === yesterdayYmd) deptStats[r.dept].dayTSmiles += r.smiles;

    if (BRANCH_TO_REGION[r.branch]) {
      if (!branchStats[r.branch]) {
        branchStats[r.branch] = { branch: r.branch, region: r.region, dayTSmiles: 0, mtdImages: 0, mtdVideos: 0, mtdSmiles: 0 };
      }
      if (r.monthKey === currentMonthKey) {
        branchStats[r.branch].mtdImages += r.images;
        branchStats[r.branch].mtdVideos += r.videos;
        branchStats[r.branch].mtdSmiles += r.smiles;
      }
      if (r.ymd === yesterdayYmd) branchStats[r.branch].dayTSmiles += r.smiles;
    }
  });

  const topBranches = Object.values(branchStats).sort((a, b) => b.mtdSmiles - a.mtdSmiles).slice(0, 7);
  const zeroRegions = REGION_BRANCH_HIERARCHY.filter(g => regionStats[g.region].mtdSmiles === 0).map(g => g.region);

  const regionRowsHtml = REGION_BRANCH_HIERARCHY.map((g, i) => {
    const st = regionStats[g.region];
    const rate = Math.round((st.activeCN.size / st.totalCN) * 100);
    const dayBadge = st.dayTSmiles > 0
      ? `<span style="background:#dcfce7;color:#166534;padding:2px 7px;border-radius:99px;font-weight:700;">+${st.dayTSmiles}</span>`
      : `<span style="color:#94a3b8;">0</span>`;
    return `
      <tr style="border-bottom:1px solid #e2e8f0; background:${st.mtdSmiles === 0 ? '#fff1f2' : (i % 2 === 0 ? '#ffffff' : '#f8fafc')}">
        <td style="padding:7px 10px;text-align:center;color:#64748b;font-weight:600;">${i + 1}</td>
        <td style="padding:7px 10px;font-weight:700;color:#1e3a8a;">${st.region} <span style="font-weight:400;color:#64748b;font-size:11.5px;">(${st.name.split('(')[0].trim()})</span></td>
        <td style="padding:7px 10px;text-align:center;"><strong>${st.activeCN.size}</strong>/${st.totalCN} <span style="color:#64748b;font-size:11px;">(${rate}%)</span></td>
        <td style="padding:7px 10px;text-align:right;">${dayBadge}</td>
        <td style="padding:7px 10px;text-align:right;font-weight:600;">${st.weekSmiles}</td>
        <td style="padding:7px 10px;text-align:right;">${st.mtdImages}</td>
        <td style="padding:7px 10px;text-align:right;">${st.mtdVideos}</td>
        <td style="padding:7px 10px;text-align:right;font-weight:800;color:#1d4ed8;background:#eff6ff;">${st.mtdSmiles}</td>
      </tr>
    `;
  }).join('');

  const deptRowsHtml = Object.entries(deptStats)
    .sort((a, b) => b[1].mtdSmiles - a[1].mtdSmiles)
    .map(([dept, st], i) => {
      const share = mtdSmiles > 0 ? ((st.mtdSmiles / mtdSmiles) * 100).toFixed(1) : '0.0';
      return `
        <tr style="border-bottom:1px solid #e2e8f0;background:${i % 2 === 0 ? '#ffffff' : '#f8fafc'}">
          <td style="padding:7px 10px;text-align:center;color:#64748b;font-weight:600;">${i + 1}</td>
          <td style="padding:7px 10px;font-weight:700;color:#0f172a;">${dept}</td>
          <td style="padding:7px 10px;text-align:right;color:#166534;font-weight:600;">${st.dayTSmiles > 0 ? '+' + st.dayTSmiles : '0'}</td>
          <td style="padding:7px 10px;text-align:right;">${st.weekSmiles}</td>
          <td style="padding:7px 10px;text-align:right;">${st.mtdImages}</td>
          <td style="padding:7px 10px;text-align:right;">${st.mtdVideos}</td>
          <td style="padding:7px 10px;text-align:right;font-weight:800;color:#047857;">${st.mtdSmiles}</td>
          <td style="padding:7px 10px;text-align:right;color:#64748b;">${share}%</td>
        </tr>
      `;
    }).join('');

  const topBranchesHtml = topBranches.map((b, i) => `
    <tr style="border-bottom:1px solid #e2e8f0;">
      <td style="padding:6px 10px;text-align:center;font-weight:700;color:#b45309;">#${i + 1}</td>
      <td style="padding:6px 10px;font-weight:700;">${b.branch}</td>
      <td style="padding:6px 10px;text-align:center;"><span style="background:#dbeafe;color:#1e40af;padding:2px 6px;border-radius:4px;font-size:11px;font-weight:700;">${b.region}</span></td>
      <td style="padding:6px 10px;text-align:right;">+${b.dayTSmiles}</td>
      <td style="padding:6px 10px;text-align:right;">${b.mtdImages} ảnh / ${b.mtdVideos} video</td>
      <td style="padding:6px 10px;text-align:right;font-weight:800;color:#7c3aed;">${b.mtdSmiles} 😊</td>
    </tr>
  `).join('');

  return `
    <div style="font-family:'Segoe UI',Arial,sans-serif;color:#0f172a;max-width:840px;margin:0 auto;border:1px solid #cbd5e1;border-radius:12px;overflow:hidden;box-shadow:0 4px 12px rgba(15,23,42,0.08);">
      <!-- Khung mô phỏng Header Email gửi tự động -->
      <div style="background:#f8fafc;padding:12px 20px;border-bottom:1px solid #e2e8f0;font-size:12.5px;color:#334155;">
        <div style="margin-bottom:4px;"><strong>Tiêu đề (Subject):</strong> <span style="color:#0f172a;font-weight:700;">(Báo cáo tay GG sheet) Báo cáo Thống kê Nụ Cười hàng ngày (Chốt 09:00 sáng ${formatDMY(latestYmd)})</span></div>
        <div style="margin-bottom:4px;"><strong>Người nhận (To):</strong> <span style="color:#1d4ed8;font-weight:600;">VuongNT1@vienthongtin.com</span></div>
        <div style="margin-bottom:4px;"><strong>Lịch gửi tự động:</strong> 09:30 AM hàng ngày (Sau giờ chốt duyệt bài ngày T lúc 09:00 sáng T+1)</div>
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:6px;">
          <strong>File đính kèm (2 file):</strong>
          <a href="Bao_Cao_Dashboard_TTC.html" download style="display:inline-flex;align-items:center;gap:4px;background:#eff6ff;border:1px solid #93c5fd;color:#1d4ed8;padding:3px 10px;border-radius:6px;text-decoration:none;font-weight:600;font-size:12px;">
            📎 Bao_Cao_Dashboard_TTC.html (Mở Dashboard tương tác độc lập)
          </a>
          <span style="display:inline-flex;align-items:center;gap:4px;background:#ecfdf5;border:1px solid #6ee7b7;color:#047857;padding:3px 10px;border-radius:6px;font-weight:600;font-size:12px;">
            📊 Bao_Cao_Nu_Cuoi_TTC_${latestYmd}.xlsx (3 Sheet có STT)
          </span>
        </div>
      </div>

      <!-- Banner Chính (Chuẩn Outlook Desktop: dùng table + bgcolor="#1e3a8a" + background-color:#1e3a8a) -->
      <table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#1e3a8a" style="width:100%;border-collapse:collapse;background-color:#1e3a8a;">
        <tr>
          <td bgcolor="#1e3a8a" style="background-color:#1e3a8a;color:#ffffff;padding:20px 24px;">
            <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.08em;color:#fde68a;font-weight:700;">CHƯƠNG TRÌNH FTEL30 - TRIỆU NỤ CƯỜI TTC</div>
            <div style="margin:6px 0;font-size:20px;font-weight:800;color:#ffffff;">😊 BÁO CÁO THỐNG KÊ NỤ CƯỜI HÀNG NGÀY</div>
            <div style="margin:0;color:#e2e8f0;font-size:13px;">
              Ngày chốt số liệu: <strong style="color:#ffffff;">${formatDMY(yesterdayYmd)} (Ngày T)</strong> &amp; sáng <strong style="color:#ffffff;">${formatDMY(latestYmd)}</strong> &nbsp;|&nbsp;
              Quy tắc: <strong style="color:#fde68a;">Hình ảnh = 3 nụ cười, Video = 5 nụ cười</strong>
            </div>
          </td>
        </tr>
      </table>

      <div style="padding:20px 24px;background:#ffffff;">
        <p style="margin:0 0 14px;font-size:13.5px;color:#334155;">
          Kính gửi <strong>anh Vương</strong>,<br/>
          Hệ thống gửi báo cáo tổng hợp kết quả chương trình <strong>Triệu Nụ Cười</strong> theo Ngày, Tuần và Lũy kế Tháng (kèm file <code>.html</code> Dashboard tương tác đính kèm):
        </p>

        <!-- 4 Ô KPI Tổng quan -->
        <table style="width:100%;border-collapse:collapse;margin-bottom:20px;">
          <tr>
            <td style="padding:12px;background:#ecfdf5;border:1px solid #a7f3d0;border-radius:8px;text-align:center;width:24%;">
              <div style="font-size:11px;color:#065f46;font-weight:700;text-transform:uppercase;">Phát Sinh Ngày T (${formatShortDM(yesterdayYmd)})</div>
              <div style="font-size:23px;font-weight:800;color:#047857;margin:3px 0;">+${dayTSmiles} 😊</div>
              <div style="font-size:11.5px;color:#065f46;">${dayTImages} Ảnh (×3) • ${dayTVideos} Video (×5)</div>
            </td>
            <td style="width:1.3%;"></td>
            <td style="padding:12px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:8px;text-align:center;width:24%;">
              <div style="font-size:11px;color:#1e40af;font-weight:700;text-transform:uppercase;">Lũy Kế Tuần (${currentWeekKey})</div>
              <div style="font-size:23px;font-weight:800;color:#1d4ed8;margin:3px 0;">${weekSmiles} 😊</div>
              <div style="font-size:11.5px;color:#1e40af;">Sáng ${formatShortDM(latestYmd)} thêm: +${todaySmiles} 😊</div>
            </td>
            <td style="width:1.3%;"></td>
            <td style="padding:12px;background:#fffbeb;border:1px solid #fde68a;border-radius:8px;text-align:center;width:24%;">
              <div style="font-size:11px;color:#92400e;font-weight:700;text-transform:uppercase;">Lũy Kế Tháng (${currentMonthKey})</div>
              <div style="font-size:23px;font-weight:800;color:#b45309;margin:3px 0;">${mtdSmiles} 😊</div>
              <div style="font-size:11.5px;color:#92400e;">${mtdImages} Ảnh (×3) + ${mtdVideos} Video (×5)</div>
            </td>
            <td style="width:1.3%;"></td>
            <td style="padding:12px;background:#f5f3ff;border:1px solid #ddd6fe;border-radius:8px;text-align:center;width:24%;">
              <div style="font-size:11px;color:#5b21b6;font-weight:700;text-transform:uppercase;">Độ Phủ Đơn Vị</div>
              <div style="font-size:23px;font-weight:800;color:#6d28d9;margin:3px 0;">${activeBranchesMtd.size}/${ALL_BRANCHES.length} CN</div>
              <div style="font-size:11.5px;color:#5b21b6;">${activeRegionsMtd.size}/${REGION_BRANCH_HIERARCHY.length} Khu vực có bài</div>
            </td>
          </tr>
        </table>

        <!-- Bảng 1: 16 Khu vực -->
        <h3 style="font-size:14px;margin:18px 0 8px;color:#0f172a;border-left:4px solid #2563eb;padding-left:8px;">
          1. Thống Kê Theo ${REGION_BRANCH_HIERARCHY.length} Khu Vực (Có STT - Ngày / Tuần / Tháng)
        </h3>
        <table style="width:100%;border-collapse:collapse;font-size:12.5px;border:1px solid #cbd5e1;">
          <thead>
            <tr style="background:#1e293b;color:#ffffff;">
              <th style="padding:8px 8px;">STT</th>
              <th style="padding:8px 10px;text-align:left;">Khu Vực</th>
              <th style="padding:8px 8px;">Độ Phủ CN</th>
              <th style="padding:8px 10px;text-align:right;">Ngày T (${formatShortDM(yesterdayYmd)})</th>
              <th style="padding:8px 10px;text-align:right;">Tuần Này</th>
              <th style="padding:8px 10px;text-align:right;">Ảnh Tháng (×3)</th>
              <th style="padding:8px 10px;text-align:right;">Video Tháng (×5)</th>
              <th style="padding:8px 10px;text-align:right;background:#1e40af;">Tổng Nụ Cười Tháng</th>
            </tr>
          </thead>
          <tbody>
            ${regionRowsHtml}
            <tr style="background:#0f172a;color:#ffffff;font-weight:800;">
              <td style="padding:8px;text-align:center;">∑</td>
              <td style="padding:8px 10px;">TỔNG TOÀN HỆ THỐNG (${REGION_BRANCH_HIERARCHY.length} KV)</td>
              <td style="padding:8px;text-align:center;">${activeBranchesMtd.size}/${ALL_BRANCHES.length} CN</td>
              <td style="padding:8px 10px;text-align:right;color:#6ee7b7;">+${dayTSmiles}</td>
              <td style="padding:8px 10px;text-align:right;">${weekSmiles}</td>
              <td style="padding:8px 10px;text-align:right;">${mtdImages}</td>
              <td style="padding:8px 10px;text-align:right;">${mtdVideos}</td>
              <td style="padding:8px 10px;text-align:right;color:#fde68a;font-size:13.5px;">${mtdSmiles}</td>
            </tr>
          </tbody>
        </table>

        <!-- Bảng 2: Bộ phận -->
        <h3 style="font-size:14px;margin:20px 0 8px;color:#0f172a;border-left:4px solid #10b981;padding-left:8px;">
          2. Thống Kê Theo Bộ Phận (Cột: Bộ phận - Có STT)
        </h3>
        <table style="width:100%;border-collapse:collapse;font-size:12.5px;border:1px solid #cbd5e1;">
          <thead>
            <tr style="background:#334155;color:#ffffff;">
              <th style="padding:8px;">STT</th>
              <th style="padding:8px 10px;text-align:left;">Bộ Phận</th>
              <th style="padding:8px 10px;text-align:right;">Ngày T (${formatShortDM(yesterdayYmd)})</th>
              <th style="padding:8px 10px;text-align:right;">Tuần Này</th>
              <th style="padding:8px 10px;text-align:right;">Hình Ảnh (×3)</th>
              <th style="padding:8px 10px;text-align:right;">Video (×5)</th>
              <th style="padding:8px 10px;text-align:right;">Tổng Nụ Cười</th>
              <th style="padding:8px 10px;text-align:right;">Tỷ Trọng</th>
            </tr>
          </thead>
          <tbody>${deptRowsHtml}</tbody>
        </table>

        <!-- Bảng 3: Vinh danh Top Chi nhánh & Lưu ý đôn đốc -->
        <h3 style="font-size:14px;margin:20px 0 8px;color:#0f172a;border-left:4px solid #7c3aed;padding-left:8px;">
          3. Top Chi Nhánh Dẫn Đầu &amp; Nhắc Nhở Đơn Vị
        </h3>
        <table style="width:100%;border-collapse:collapse;font-size:12.5px;border:1px solid #cbd5e1;margin-bottom:14px;">
          <thead>
            <tr style="background:#4c1d95;color:#ffffff;">
              <th style="padding:7px 10px;">Hạng</th>
              <th style="padding:7px 10px;text-align:left;">Chi Nhánh</th>
              <th style="padding:7px 10px;">Khu Vực</th>
              <th style="padding:7px 10px;text-align:right;">Phát Sinh Ngày T</th>
              <th style="padding:7px 10px;text-align:right;">Chi Tiết Lũy Kế</th>
              <th style="padding:7px 10px;text-align:right;">Tổng Điểm Nụ Cười</th>
            </tr>
          </thead>
          <tbody>${topBranchesHtml}</tbody>
        </table>

        <div style="background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:12px 16px;font-size:12.5px;color:#92400e;">
          <div style="margin-bottom:4px;"><strong>⚠️ Khu vực chưa phát sinh bài đăng trong tháng (${zeroRegions.length} KV):</strong> ${zeroRegions.length > 0 ? zeroRegions.join(', ') : 'Tất cả 15 Khu vực đều đã có bài đăng!'}</div>
          ${missingBranchRows.length > 0 ? `<div><strong>🔎 Cần rà soát trên Google Sheet (${missingBranchRows.length} dòng chưa điền Chi nhánh/Bộ phận):</strong> STT ${missingBranchRows.map(r => `${r.id} (${r.authorEmail})`).join(', ')}.</div>` : ''}
        </div>

        <p style="margin:16px 0 0;font-size:12px;color:#64748b;text-align:center;">
          💡 <em>Để xem chi tiết toàn bộ ${ALL_BRANCHES.length} Chi nhánh (STT 1 → ${ALL_BRANCHES.length}) và lọc tương tác theo Ngày/Tuần/Tháng, vui lòng mở file đính kèm <strong>Bao_Cao_Dashboard_TTC.html</strong> trên trình duyệt.</em>
        </p>
      </div>
    </div>
  `;
}

// Gắn sự kiện tương tác trên giao diện
function bindEvents() {
  // Chuyển đổi Chiều thống kê: Ngày / Tuần / Tháng
  document.querySelectorAll('#timeGrainControl .seg-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#timeGrainControl .seg-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.timeGrain = btn.dataset.grain;
      refreshDashboard();
    });
  });

  // Chọn nhanh khoảng thời gian
  document.querySelectorAll('#quickRangeGroup .chip-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      applyQuickRange(btn.dataset.range);
    });
  });

  // Thay đổi ngày tùy chọn
  document.getElementById('filterStartDate').addEventListener('change', e => {
    state.startDate = e.target.value;
    document.querySelectorAll('#quickRangeGroup .chip-btn').forEach(b => b.classList.remove('active'));
    refreshDashboard();
  });
  document.getElementById('filterEndDate').addEventListener('change', e => {
    state.endDate = e.target.value;
    document.querySelectorAll('#quickRangeGroup .chip-btn').forEach(b => b.classList.remove('active'));
    refreshDashboard();
  });

  // Lọc Khu vực / Chi nhánh / Bộ phận / Loại
  document.querySelectorAll('#regionHorizontalBar [data-region-chip]').forEach(btn => {
    btn.addEventListener('click', () => {
      const reg = btn.getAttribute('data-region-chip');
      state.region = reg;
      state.branch = 'ALL';
      document.getElementById('filterRegion').value = reg;
      updateBranchDropdown(reg);
      refreshDashboard();
    });
  });

  document.getElementById('filterRegion').addEventListener('change', e => {
    state.region = e.target.value;
    updateBranchDropdown(state.region);
    refreshDashboard();
  });
  document.getElementById('filterBranch').addEventListener('change', e => {
    state.branch = e.target.value;
    refreshDashboard();
  });
  document.getElementById('filterDept').addEventListener('change', e => {
    state.dept = e.target.value;
    refreshDashboard();
  });
  document.getElementById('filterMediaType').addEventListener('change', e => {
    state.mediaType = e.target.value;
    refreshDashboard();
  });
  document.getElementById('searchKeyword').addEventListener('input', e => {
    state.keyword = e.target.value;
    refreshDashboard();
  });

  // Tùy chọn hiển thị bảng
  const sortBranchEl = document.getElementById('sortBranchSelect');
  if (sortBranchEl) {
    sortBranchEl.addEventListener('change', e => {
      state.branchSort = e.target.value;
      refreshDashboard();
    });
  }
  document.getElementById('chkShowZeroBranches').addEventListener('change', e => {
    state.showZeroBranches = e.target.checked;
    refreshDashboard();
  });
  document.getElementById('chkShowTimeColumns').addEventListener('change', e => {
    state.showTimeColumns = e.target.checked;
    refreshDashboard();
  });
  document.getElementById('sortRegionSelect').addEventListener('change', e => {
    state.regionSort = e.target.value;
    refreshDashboard();
  });
  document.getElementById('deptMatrixLevel').addEventListener('change', e => {
    state.deptMatrixLevel = e.target.value;
    refreshDashboard();
  });

  // Chuyển Tab
  document.querySelectorAll('#mainNavTabs .tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#mainNavTabs .tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      const pane = document.getElementById(btn.dataset.tab);
      if (pane) pane.classList.add('active');
      if (btn.dataset.tab === 'tab-charts') {
        refreshDashboard();
      }
    });
  });

  // Đồng bộ Google Sheet (nếu có nút trên giao diện Admin)
  const btnSync = document.getElementById('btnSyncSheet');
  if (btnSync) {
    btnSync.addEventListener('click', () => {
      trySyncDataSource(true);
    });
  }

  // Nhập file Excel / CSV (nếu có trên giao diện Admin)
  const fileUpload = document.getElementById('fileUploadInput');
  if (fileUpload) {
    fileUpload.addEventListener('change', e => {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = evt => {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });
        const parsed = ingestTabularRows(rows);
        if (parsed.length > 0) {
          applyNewRecords(parsed, false, `File: ${file.name} (${parsed.length} dòng)`);
        } else {
          alert('Không tìm thấy dữ liệu hợp lệ trong file. Vui lòng kiểm tra các cột: Thời gian đăng, Chi nhánh, Loại, Bộ phận.');
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  // Xuất Excel
  const btnExport = document.getElementById('btnExportExcel');
  if (btnExport) {
    btnExport.addEventListener('click', exportExcelReport);
  }

  // Modal Email (nếu có trên giao diện Admin)
  const emailModal = document.getElementById('emailModal');
  const btnPreviewEmail = document.getElementById('btnPreviewEmail');
  if (emailModal && btnPreviewEmail) {
    btnPreviewEmail.addEventListener('click', () => {
      document.getElementById('emailPreviewBody').innerHTML = generateDailyEmailHtml();
      emailModal.classList.add('open');
    });
    document.getElementById('btnCloseEmailModal')?.addEventListener('click', () => emailModal.classList.remove('open'));
    document.getElementById('btnCloseEmailFooter')?.addEventListener('click', () => emailModal.classList.remove('open'));
    document.getElementById('btnCopyEmailHtml')?.addEventListener('click', () => {
      const html = generateDailyEmailHtml();
      navigator.clipboard.writeText(html).then(() => {
        alert('Đã copy mã HTML báo cáo Email vào clipboard!');
      });
    });
  }
}

// Khởi chạy ứng dụng
document.addEventListener('DOMContentLoaded', () => {
  bindEvents();

  // Nếu mở qua web (GitHub Pages / localhost), tự động kiểm tra bản data.json mới nhất định kỳ mỗi 3 phút
  if (window.location.protocol.startsWith('http')) {
    setInterval(async () => {
      try {
        const resp = await fetch(`data.json?t=${Date.now()}`, { cache: 'no-store' });
        if (resp.ok) {
          const payload = await resp.json();
          if (payload && payload.updated_at && (!window.LIVE_SHEET_DATA || payload.updated_at !== window.LIVE_SHEET_DATA.updated_at)) {
            window.LIVE_SHEET_DATA = payload;
            const parsed = ingestTabularRows(payload.records);
            if (parsed.length > 0) {
              applyNewRecords(parsed, false, `Dữ liệu thực Google Sheet (${parsed.length} bài đăng)`);
            }
          }
        }
      } catch (_) {}
    }, 180000);
  }

  // Ưu tiên nạp trực tiếp từ sheet_data_bundle.js (đã đồng bộ từ Google Sheet thực tế)
  if (window.LIVE_SHEET_DATA && Array.isArray(window.LIVE_SHEET_DATA.records) && window.LIVE_SHEET_DATA.records.length > 0) {
    const parsed = ingestTabularRows(window.LIVE_SHEET_DATA.records);
    if (parsed.length > 0) {
      applyNewRecords(parsed, false, `Dữ liệu thực Google Sheet (${parsed.length} bài đăng)`);
      if (window.location.search.includes('preview_email=1')) {
        document.getElementById('emailPreviewBody').innerHTML = generateDailyEmailHtml();
        document.getElementById('emailModal').classList.add('open');
      }
      // Nếu mở qua web (GitHub Pages / localhost), tự động kiểm tra bản data.json mới nhất chống cache trình duyệt
      if (window.location.protocol.startsWith('http')) {
        trySyncDataSource(false);
      }
      return;
    }
  }
  // Không nạp dữ liệu giả lập (demo) nữa để tránh làm lệch số báo cáo; tự động đồng bộ từ Google Sheet / data.json
  applyNewRecords([], false, 'Đang tải dữ liệu thực Google Sheet...');
  trySyncDataSource(false);
});

