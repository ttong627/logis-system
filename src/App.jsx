import React, { useState, useMemo, useEffect, useRef } from 'react';
import { Settings, Truck, FileText, Plus, Trash2, CheckCircle2, AlertCircle, MapPin, Users, PhoneCall, Building2, BookOpen, Lock, Save, FolderOpen, Calendar, CreditCard, Download, ShieldCheck, Eye, PieChart, RefreshCw, Upload, FileDown, Sparkles, Copy, X } from 'lucide-react';
import { initializeApp } from 'firebase/app';
import { getAuth, signInWithCustomToken, signInAnonymously, onAuthStateChanged } from 'firebase/auth';
import { getFirestore, doc, setDoc, getDoc, collection, query, onSnapshot } from 'firebase/firestore';

// --- 🌟 데이터 보존을 위한 영구 고정 식별자 (데이터 유실 방지) 🌟 ---
const WELLSHARE_PERMANENT_ID = 'wellshare-logis-v1-production-stable';

// --- Firebase 초기화 (환경 변수 적용 완료!) ---
const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
  measurementId: import.meta.env.VITE_FIREBASE_MEASUREMENT_ID
};

let app, auth, db, appId;
try {
  app = initializeApp(firebaseConfig);
  auth = getAuth(app);
  db = getFirestore(app);
  // 환경변수가 바뀌어도 고정된 ID를 사용하여 데이터를 안전하게 유지합니다.
  appId = typeof __app_id !== 'undefined' ? __app_id : WELLSHARE_PERMANENT_ID;
} catch (error) {
  console.warn('Firebase 초기화 실패:', error);
}

// --- 🌟 Gemini API 연동 (AI 기능용) 🌟 ---
const generateGeminiContent = async (prompt) => {
  const apiKey = ""; // 실제 환경에서는 자동으로 주입됩니다.
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent?key=${apiKey}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    systemInstruction: {
      parts: [{ text: "당신은 물류 배송 정산 관리 시스템에 내장된 유능하고 전문적인 업무 보조 AI 비서입니다. 답변은 항상 한국어로 간결하고 명확하게 작성하세요." }]
    }
  };

  let retries = 0;
  const delays = [1000, 2000, 4000, 8000, 16000];

  while (retries < 5) {
    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      
      if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
      const data = await response.json();
      return data.candidates?.[0]?.content?.parts?.[0]?.text || "결과를 생성하지 못했습니다.";
    } catch (error) {
      if (retries === 4) return `AI 생성 중 오류가 발생했습니다: ${error.message}`;
      await new Promise(resolve => setTimeout(resolve, delays[retries]));
      retries++;
    }
  }
};

// --- 지자체 출력 순서 절대 고정 ---
const REGION_ORDER = [
  '부천시 소사구', '부천시 원미구', '부천시 오정구', '시흥시', '여주시', '중구', '종로구', '용산구', '동대문구'
];

const SEOUL_REGIONS = ['중구', '종로구', '용산구', '동대문구'];
const GYEONGGI_REGIONS = ['부천시 소사구', '부천시 원미구', '부천시 오정구', '시흥시', '여주시'];

// 지역명 포맷 변환기 (결제서 엑셀 양식용)
const getFullRegionName = (region) => {
  if (SEOUL_REGIONS.includes(region)) return `서울 ${region}`;
  if (GYEONGGI_REGIONS.includes(region)) return `경기 ${region}`;
  return region;
};

// 회원사 목록
const MEMBERS = ['사회적협동조합 행복나눔', '참자연', '미소 협동조합', '웰쉐어 사회적협동조합', '부천희망나르미', '(주)한울'];

// --- 유틸리티 ---
const parseNumber = (val) => {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  return Number(String(val).replace(/[^0-9]/g, '')) || 0;
};
const formatNumber = (num) => (num || num === 0) ? num.toLocaleString('ko-KR') : '0';
const formatCur = (num) => new Intl.NumberFormat('ko-KR').format(num) + '원';

const copyToClipboard = (text) => {
  const textArea = document.createElement("textarea");
  textArea.value = text;
  document.body.appendChild(textArea);
  textArea.select();
  try { document.execCommand('copy'); } catch (err) { console.error('복사 실패', err); }
  document.body.removeChild(textArea);
};

// --- 초기 기본 설정 ---
const INITIAL_ZONES = {
  '1급지': { billing: 2780 }, '2급지': { billing: 2810 }, '3급지': { billing: 3000 },
  '4급지': { billing: 3200 }, '5급지': { billing: 3490 }, '6급지': { billing: 3820 }, '7급지': { billing: 4170 },
};

const INITIAL_REGIONS_DATA = {
  '부천시 소사구': '1급지', '부천시 원미구': '1급지', '부천시 오정구': '1급지',
  '시흥시': '2급지', '여주시': '4급지', '중구': '2급지', '종로구': '2급지', '용산구': '2급지', '동대문구': '2급지',
};

const FIXED_MAPPING = {
  '부천시 원미구': '사회적협동조합 행복나눔', '부천시 오정구': '사회적협동조합 행복나눔',
  '부천시 소사구': '부천희망나르미', '중구': '부천희망나르미', '종로구': '부천희망나르미',
  '용산구': '부천희망나르미', '여주시': '미소 협동조합'
};

const CONTACTS = [
  { agency: '사회적협동조합 행복나눔', region: '시흥시', detail: '능곡동 제외 전체', manager: '전재형 이사장', phone: '010-4710-7460' },
  { agency: '사회적협동조합 행복나눔', region: '부천시 오정구', detail: '', manager: '조유라 팀장', phone: '010-4726-0437' },
  { agency: '사회적협동조합 행복나눔', region: '부천시 원미구', detail: '', manager: '사무실', phone: '070-7518-7362' },
  { agency: '사회적협동조합 행복나눔', region: '서울 동대문구', detail: '청량리동, 이문2동, 제기동, 회기동', manager: '', phone: '' },
  { agency: '참자연', region: '경기도 시흥시', detail: '능곡동', manager: '박경선 대표', phone: '010-7424-2477' },
  { agency: '미소 협동조합', region: '경기도 여주시', detail: '전체', manager: '김해승 대표 / 강성임 실장', phone: '010-7537-3447 / 010-6530-4211' },
  { agency: '웰쉐어 사회적협동조합', region: '서울 동대문구', detail: '답십리1동, 전농1동, 휘경1동', manager: '이진만 차장', phone: '010-6381-9205' },
  { agency: '부천희망나르미', region: '부천시 소사구', detail: '대산동, 범안동, 소사본동(소사구)', manager: '박한울 과장', phone: '010-3110-9426' },
  { agency: '부천희망나르미', region: '서울 중구', detail: '전체', manager: '김은선 과장', phone: '010-7152-8729' },
  { agency: '부천희망나르미', region: '서울 종로구', detail: '전체', manager: '사무실', phone: '032-713-4644' },
  { agency: '부천희망나르미', region: '서울 용산구', detail: '전체', manager: '', phone: '' },
  { agency: '(주)한울', region: '서울 동대문구', detail: '용신동, 전농2동, 휘경2동, 답십리2동, 이문1동, 장안1동, 장안2동', manager: '장영수 대표', phone: '010-9457-1617' },
];

const GOV_CONTACTS = [
  { region: '종로구', order: '1차', manager: '-', phone: '02-2148-2553' },
  { region: '동대문구', order: '2차', manager: '심정영', phone: '02-2127-4569' },
  { region: '용산구', order: '2차', manager: '-', phone: '02-2199-7094' },
  { region: '여주시', order: '2차', manager: '강호연', phone: '031-887-2278' },
  { region: '중구', order: '1차', manager: '-', phone: '02-3396-5351' },
  { region: '부천시', order: '1차', manager: '-', phone: '032-625-2859' },
  { region: '시흥시', order: '1차', manager: '이채원', phone: '031-310-2269' },
];

// 엑셀 내보내기용 인라인 스타일 (디자인 유지)
const excelStyles = {
  table: { borderCollapse: 'collapse', width: '100%', fontFamily: '"Malgun Gothic", sans-serif', fontSize: '10pt', textAlign: 'center' },
  th: { border: '1pt solid #000000', padding: '6px', backgroundColor: '#d9d9d9', fontWeight: 'bold', textAlign: 'center', verticalAlign: 'middle' },
  td: { border: '1pt solid #000000', padding: '6px', verticalAlign: 'middle' },
  tdLeft: { border: '1pt solid #000000', padding: '6px', verticalAlign: 'middle', textAlign: 'left' },
  tdRight: { border: '1pt solid #000000', padding: '6px', verticalAlign: 'middle', textAlign: 'right' },
  titleRow: { backgroundColor: '#f2f2f2', fontWeight: 'bold' },
  subTotalRow: { backgroundColor: '#ccff99', fontWeight: 'bold' },
  regionSumRow: { backgroundColor: '#ffff00', fontWeight: 'bold' },
  grandSumRow: { backgroundColor: '#00b0f0', color: '#ffffff', fontWeight: 'bold', fontSize: '12pt' },
  emptyRow: { border: 'none', height: '20px' }
};

export default function App() {
  const [activeTab, setActiveTab] = useState('orders'); 
  const [isAdmin, setIsAdmin] = useState(true);
  const [user, setUser] = useState(null);
  const fileInputRef = useRef(null);

  // --- 상태 관리 ---
  const [zonePrices, setZonePrices] = useState(INITIAL_ZONES);
  const [regions, setRegions] = useState(INITIAL_REGIONS_DATA);
  const [orders, setOrders] = useState({});
  const [allocations, setAllocations] = useState({}); 
  const [savedMonths, setSavedMonths] = useState([]);
  const [isSaving, setIsSaving] = useState(false);
  const [isXlsxReady, setIsXlsxReady] = useState(false);
  const [toastMessage, setToastMessage] = useState(null);
  const [currentMonth, setCurrentMonth] = useState(() => {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  });

  const [aiAnalysis, setAiAnalysis] = useState("");
  const [isAiLoading, setIsAiLoading] = useState(false);
  const [draftModal, setDraftModal] = useState({ isOpen: false, data: null, text: "", isLoading: false });

  const showToast = (msg) => {
    setToastMessage(msg);
    setTimeout(() => setToastMessage(null), 3000);
  };

  // 1. 인증 및 라이브러리 준비
  useEffect(() => {
    const initAuth = async () => {
      try {
        if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) {
          await signInWithCustomToken(auth, __initial_auth_token);
        } else {
          await signInAnonymously(auth);
        }
      } catch (e) { console.error("인증 실패:", e); }
    };
    initAuth();
    const unsubscribe = onAuthStateChanged(auth, setUser);

    const script = document.createElement("script");
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    script.onload = () => setIsXlsxReady(true);
    document.head.appendChild(script);

    return () => unsubscribe();
  }, []);

  // 2. 기록 목록 로드
  useEffect(() => {
    if (!user) return;
    const q = query(collection(db, 'artifacts', appId, 'public', 'data', 'billing_records'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const months = snapshot.docs.map(doc => doc.id);
      setSavedMonths(months.sort().reverse());
    }, (err) => console.error("데이터 목록 로드 실패:", err));
    return () => unsubscribe();
  }, [user]);

  // 3. 데이터 로딩 (월 변경/새로고침 시)
  const forceLoadData = async (month) => {
    if (!user || !db || !month) return;
    try {
      const docRef = doc(db, 'artifacts', appId, 'public', 'data', 'billing_records', month);
      const docSnap = await getDoc(docRef);
      if (docSnap.exists()) {
        const data = docSnap.data();
        setZonePrices(data.zonePrices || INITIAL_ZONES);
        setRegions(data.regions || INITIAL_REGIONS_DATA);
        setOrders(data.orders || {});
        setAllocations(data.allocations || {});
      } else {
        setZonePrices(INITIAL_ZONES);
        setRegions(INITIAL_REGIONS_DATA);
        setOrders({});
        setAllocations({});
      }
      setAiAnalysis(""); 
    } catch (e) { console.error("데이터 로드 오류:", e); }
  };

  useEffect(() => {
    if (user) forceLoadData(currentMonth);
  }, [user, currentMonth]);

  // 4. 데이터 영구 저장
  const handleSaveData = async () => {
    if (!isAdmin) return showToast("읽기 전용 모드입니다.");
    if (!user) return showToast("사용자 인증 대기 중...");
    setIsSaving(true);
    try {
      const docRef = doc(db, 'artifacts', appId, 'public', 'data', 'billing_records', currentMonth);
      await setDoc(docRef, {
        zonePrices, regions, orders, allocations,
        updatedAt: new Date().toISOString()
      });
      showToast(`${currentMonth} 정산 내역이 클라우드에 영구 보존되었습니다.`);
    } catch (e) { showToast("저장 오류: " + e.message); }
    finally { setIsSaving(false); }
  };

  // --- 통합 엑셀 기능 ---
  const downloadUploadTemplate = () => {
    if (!window.XLSX) return showToast("라이브러리 로딩 중...");
    const headers = ["지역명", "차상위수량", "수급자수량", ...MEMBERS];
    const data = [headers, ...REGION_ORDER.map(r => [r, 0, 0, ...MEMBERS.map(() => 0)])];
    const ws = window.XLSX.utils.aoa_to_sheet(data);
    const wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, ws, "정산통합양식");
    window.XLSX.writeFile(wb, `수량_및_배분_통합양식_희망나르미.xlsx`);
  };

  const handleExcelUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = window.XLSX.read(bstr, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = window.XLSX.utils.sheet_to_json(ws);
      const nextOrders = { ...orders };
      const nextAllocations = { ...allocations };
      data.forEach(row => {
        const rName = row["지역명"];
        if (REGION_ORDER.includes(rName)) {
          nextOrders[rName] = {
            povertyQty: parseNumber(row["차상위수량"] || 0),
            basicQty: parseNumber(row["수급자수량"] || 0)
          };
          if (!nextAllocations[rName]) nextAllocations[rName] = {};
          MEMBERS.forEach(m => {
            if (row[m] !== undefined) nextAllocations[rName][m] = parseNumber(row[m]);
          });
        }
      });
      setOrders(nextOrders);
      setAllocations(nextAllocations);
      showToast("엑셀의 지자체 수량 및 회원사 할당 데이터가 모두 반영되었습니다.");
      e.target.value = null;
    };
    reader.readAsBinaryString(file);
  };

  const formattedMonthStr = useMemo(() => {
    if (!currentMonth) return '';
    const [year, month] = currentMonth.split('-');
    return `${year}년 ${parseInt(month, 10)}월`;
  }, [currentMonth]);

  // 🌟 (완벽 호환) 엑셀 다운로드 🌟
  const handleDownloadExcel = (tableId, fileName) => {
    const table = document.getElementById(tableId);
    if (!table) return;
    const excelFile = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">
      </head>
      <body>${table.outerHTML}</body>
      </html>
    `;
    const blob = new Blob([excelFile], { type: 'application/vnd.ms-excel' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${currentMonth}_${fileName}.xls`;
    link.click();
  };

  const handleZonePriceChange = (zone, field, value) => {
    if(!isAdmin) return;
    setZonePrices(prev => ({ ...prev, [zone]: { ...prev[zone], [field]: parseNumber(value) } }));
  };
  const handleRegionZoneChange = (r, z) => { if(!isAdmin) return; setRegions(prev => ({ ...prev, [r]: z })); };
  const handleOrderChange = (r, f, v) => {
    if(!isAdmin) return;
    setOrders(prev => ({ ...prev, [r]: { ...(prev[r] || { basicQty: 0, povertyQty: 0 }), [f]: parseNumber(v) } }));
  };
  const handleKeyDown = (e, currentId) => {
    if(!isAdmin) return;
    const seq = [];
    REGION_ORDER.forEach(r => {
      seq.push(`poverty-${r}`); seq.push(`basic-${r}`);
      if (r === '시흥시') seq.push(`alloc-시흥시-참자연`);
      else if (r === '동대문구') {
        seq.push(`alloc-동대문구-(주)한울`); seq.push(`alloc-동대문구-사회적협동조합 행복나눔`);
      }
    });
    const idx = seq.indexOf(currentId);
    if (idx === -1) return;
    if (e.key === 'Enter' || e.key === 'ArrowDown') {
      e.preventDefault();
      const next = seq[idx + 1];
      if (next) document.getElementById(next)?.focus();
      else if (e.key === 'Enter') e.target.blur();
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      const prev = seq[idx - 1];
      if (prev) document.getElementById(prev)?.focus();
    }
  };
  const handleAllocationChange = (r, m, v) => {
    if(!isAdmin) return;
    setAllocations(prev => ({ ...prev, [r]: { ...prev[r], [m]: parseNumber(v) } }));
  };

  // --- 핵심 연산 로직 ---
  const computedAllocations = useMemo(() => {
    const computed = {};
    REGION_ORDER.forEach(r => {
      const order = orders[r] || { basicQty: 0, povertyQty: 0 };
      const total = order.basicQty + order.povertyQty;
      computed[r] = {};
      if (FIXED_MAPPING[r]) {
        computed[r][FIXED_MAPPING[r]] = total;
      } else if (r === '시흥시') {
        const cham = allocations[r]?.['참자연'] || 0;
        computed[r]['참자연'] = cham;
        computed[r]['사회적협동조합 행복나눔'] = Math.max(0, total - cham);
      } else if (r === '동대문구') {
        const hanul = allocations[r]?.['(주)한울'] || 0;
        const haeng = allocations[r]?.['사회적협동조합 행복나눔'] || 0;
        computed[r]['(주)한울'] = hanul;
        computed[r]['사회적협동조합 행복나눔'] = haeng;
        computed[r]['웰쉐어 사회적협동조합'] = Math.max(0, total - hanul - haeng);
      } else {
        Object.entries(allocations[r] || {}).forEach(([m, q]) => { computed[r][m] = q || 0; });
      }
    });
    return computed;
  }, [orders, allocations]);

  const billingReport = useMemo(() => {
    let report = [];
    let gTotal = { qty: 0, supply: 0, vat: 0, amount: 0 };
    let sTotal = { qty: 0, supply: 0, vat: 0, amount: 0 };
    let grandTotal = { qty: 0, supply: 0, vat: 0, amount: 0 };

    REGION_ORDER.forEach(r => {
      const z = regions[r] || '2급지';
      const p = zonePrices[z]?.billing || 0;
      const o = orders[r] || { basicQty: 0, povertyQty: 0 };
      const tr = o.basicQty + o.povertyQty;

      if (tr > 0) {
        const isSeoul = SEOUL_REGIONS.includes(r);
        const city = isSeoul ? '서울시' : '경기도';
        const calc = (q) => {
          const tot = q * p;
          const sup = Math.round(tot / 1.1);
          const vat = tot - sup;
          return { tot, sup, vat };
        };
        const pov = calc(o.povertyQty);
        const bas = calc(o.basicQty);
        const sum = { qty: tr, amount: pov.tot + bas.tot, supply: pov.sup + bas.sup, vat: pov.vat + bas.vat };
        
        report.push({ city, region: r, poverty: pov, basic: bas, sum });
        const t = isSeoul ? sTotal : gTotal;
        t.qty += tr; t.amount += sum.amount; t.supply += sum.supply; t.vat += sum.vat;
        grandTotal.qty += tr; grandTotal.amount += sum.amount; grandTotal.supply += sum.supply; grandTotal.vat += sum.vat;
      }
    });
    return { report, gTotal, sTotal, grandTotal };
  }, [zonePrices, regions, orders]);

  const billingSummary = useMemo(() => {
    let gtQty = 0, gtSup = 0, gtVat = 0, gtAmt = 0;
    const pDetails = {};
    REGION_ORDER.forEach(r => {
      const z = regions[r] || '2급지';
      const price = zonePrices[z]?.billing || 0;
      Object.entries(computedAllocations[r] || {}).forEach(([m, qty]) => {
        if (qty > 0) {
          const total = Math.floor((qty * price * 0.975) / 100) * 100;
          const supply = Math.round(total / 1.1);
          const vat = total - supply;
          if (!pDetails[m]) pDetails[m] = { totalQty: 0, totalSupply: 0, totalVat: 0, totalAmount: 0, regions: [] };
          pDetails[m].totalQty += qty; pDetails[m].totalSupply += supply; pDetails[m].totalVat += vat; pDetails[m].totalAmount += total;
          pDetails[m].regions.push({ region: r, qty, supplyValue: supply, vatValue: vat, finalRowTotal: total });
          gtQty += qty; gtSup += supply; gtVat += vat; gtAmt += total;
        }
      });
    });
    const sorted = Object.entries(pDetails).sort((a,b)=>a[0].localeCompare(b[0])).map(([m, d]) => ({ member: m, ...d }));
    return { sorted, grandTotalQty: gtQty, grandTotalSupply: gtSup, grandTotalVat: gtVat, grandTotalAmount: gtAmt };
  }, [zonePrices, regions, computedAllocations]);

  const orderSummaries = useMemo(() => {
    let s = 0, g = 0, o = 0, b = 0, p = 0;
    REGION_ORDER.forEach(r => {
      const basic = orders[r]?.basicQty || 0;
      const poverty = orders[r]?.povertyQty || 0;
      o += (basic + poverty); b += basic; p += poverty;
      if (SEOUL_REGIONS.includes(r)) s += (basic + poverty);
      else if (GYEONGGI_REGIONS.includes(r)) g += (basic + poverty);
    });
    return { seoulTotal: s, gyeonggiTotal: g, overallTotal: o, basicTotal: b, povertyTotal: p };
  }, [orders]);

  const memberSummaries = useMemo(() => {
    const totals = {};
    Object.values(computedAllocations).forEach(regionAllocs => {
      Object.entries(regionAllocs).forEach(([member, qty]) => {
        totals[member] = (totals[member] || 0) + qty;
      });
    });
    return Object.entries(totals).sort((a,b)=>a[0].localeCompare(b[0])).map(([m, q]) => ({ member: m, qty: q }));
  }, [computedAllocations]);

  // --- AI 동작 함수 ---
  const handleGenerateAnalysis = async () => {
    setIsAiLoading(true);
    setAiAnalysis("");
    const prompt = `당신은 정부양곡 배송 정산 데이터를 분석하는 AI 경영 컨설턴트입니다. 다음 데이터를 바탕으로 경영진 보고용 '이번 달 정산 요약 브리핑'을 3문단으로 작성해주세요. 어조는 정중하고 전문적인 보고서 형식이어야 하며, 불필요한 인사말은 생략하세요. 주요 수치를 강조해 주세요. 
      [데이터 요약] - 기준월: ${currentMonth} - 총 배송량: ${orderSummaries.overallTotal}포 - 총 정산 금액: ${formatCur(billingReport.grandTotal.amount)}`;
    const result = await generateGeminiContent(prompt);
    setAiAnalysis(result);
    setIsAiLoading(false);
  };

  const handleDraftMessage = async (contactInfo) => {
    setDraftModal({ isOpen: true, data: contactInfo, text: "", isLoading: true });
    const prompt = `당신은 물류 회사의 사무 비서 AI입니다. 다음 담당자에게 '이번 달 정부양곡 배송 정산 내역이 시스템에 등록 및 확정되었으니 확인을 부탁드린다'는 내용의 정중한 업무용 메시지를 작성해주세요. 문구는 간결해야 하며, 끝에 (주)웰쉐어 로지스 드림 이라고 명시해주세요.
      [수신자] - 소속: ${contactInfo.agency || contactInfo.region} - 담당자: ${contactInfo.manager}`;
    const result = await generateGeminiContent(prompt);
    setDraftModal(prev => ({ ...prev, text: result, isLoading: false }));
  };

  if (!user) {
    return (
      <div className="flex flex-col items-center justify-center min-h-screen bg-white">
        <RefreshCw className="w-12 h-12 text-fuchsia-600 animate-spin mb-4" />
        <h2 className="text-xl font-black text-gray-800 tracking-tight">시스템 보안 연결 중...</h2>
        <p className="text-gray-400 text-sm mt-2">잠시만 기다려주세요.</p>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-gray-800 p-4 md:p-6 font-sans relative selection:bg-fuchsia-200 pb-20">
      {toastMessage && (
        <div className="fixed bottom-10 left-1/2 -translate-x-1/2 bg-gray-900 text-white px-8 py-4 rounded-2xl shadow-2xl z-50 flex items-center gap-3 animate-in fade-in slide-in-from-bottom-5 duration-300">
          <CheckCircle2 className="w-6 h-6 text-cyan-400" /> <span className="font-black text-lg">{toastMessage}</span>
        </div>
      )}

      {/* AI 모달 */}
      {draftModal.isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 backdrop-blur-sm p-4 animate-in fade-in">
          <div className="bg-white rounded-[2rem] shadow-2xl w-full max-w-md overflow-hidden flex flex-col border border-gray-100">
            <div className="p-6 bg-gradient-to-r from-fuchsia-600 to-cyan-500 text-white flex justify-between items-center">
              <h3 className="font-black text-lg flex items-center gap-2"><Sparkles className="w-5 h-5" /> AI 메시지 자동 생성</h3>
              <button onClick={() => setDraftModal({ isOpen: false, data: null, text: "", isLoading: false })} className="hover:bg-white/20 p-1.5 rounded-full transition-colors"><X className="w-5 h-5" /></button>
            </div>
            <div className="p-6 bg-gray-50 flex-1">
              <div className="mb-4">
                <p className="text-xs font-bold text-gray-500 mb-1">수신자</p>
                <p className="font-black text-gray-800">{draftModal.data?.agency || draftModal.data?.region} {draftModal.data?.manager}</p>
              </div>
              <div className="relative">
                {draftModal.isLoading ? (
                  <div className="flex flex-col items-center justify-center py-10 gap-3">
                    <RefreshCw className="w-8 h-8 text-cyan-500 animate-spin" />
                    <span className="text-sm font-bold text-gray-500">메시지 작성 중...</span>
                  </div>
                ) : (
                  <textarea value={draftModal.text} readOnly className="w-full h-48 p-4 bg-white border border-gray-200 rounded-xl resize-none outline-none focus:ring-2 focus:ring-cyan-500 font-medium text-gray-700 leading-relaxed" />
                )}
              </div>
            </div>
            <div className="p-4 bg-white border-t border-gray-100 flex gap-3">
              <button onClick={() => { copyToClipboard(draftModal.text); showToast("메시지가 복사되었습니다!"); setDraftModal({ isOpen: false, data: null, text: "", isLoading: false }); }} disabled={draftModal.isLoading} className="flex-1 bg-gray-900 hover:bg-black text-white font-black py-3 rounded-xl transition-all flex items-center justify-center gap-2 disabled:bg-gray-400">
                <Copy className="w-4 h-4" /> 텍스트 복사하기
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="max-w-[1400px] mx-auto">
        {/* 헤더 섹션 (로고 적용 & 디자인 개편) */}
        <header className="mb-6 flex flex-col lg:flex-row lg:items-center justify-between gap-6 bg-white p-6 rounded-[2rem] shadow-sm border border-gray-100">
          <div className="flex items-center gap-6">
            <div className="bg-white p-2 rounded-2xl border border-gray-100 shadow-sm flex items-center justify-center w-20 h-20">
              <img src="" alt="logo" className="w-16 h-16 object-contain" onError={(e) => e.target.style.display='none'} />
            </div>
            <div>
              <div className="text-sm font-black text-gray-400 tracking-widest uppercase mb-1">(주)웰쉐어 로지스</div>
              <h1 className="text-3xl font-black text-transparent bg-clip-text bg-gradient-to-r from-fuchsia-600 to-cyan-500 tracking-tighter drop-shadow-sm">희망나르미 정산시스템</h1>
            </div>
          </div>
          
          <div className="flex flex-col items-end gap-3">
            <div className="flex items-center gap-2 bg-gray-50 p-1.5 rounded-full border border-gray-200 shadow-inner">
               <button onClick={() => setIsAdmin(true)} className={`px-4 py-1.5 rounded-full text-xs font-black transition-all ${isAdmin ? 'bg-gray-900 text-white shadow-md' : 'text-gray-400 hover:text-gray-600'}`}>관리자 모드</button>
               <button onClick={() => setIsAdmin(false)} className={`px-4 py-1.5 rounded-full text-xs font-black transition-all ${!isAdmin ? 'bg-fuchsia-600 text-white shadow-md' : 'text-gray-400 hover:text-gray-600'}`}>열람 전용</button>
            </div>
            <div className="flex items-center gap-3 bg-white p-2.5 rounded-2xl border border-gray-200 shadow-sm">
              <div className="flex items-center gap-1.5 bg-gray-50 px-4 py-2.5 rounded-xl border border-gray-200">
                <Calendar className="w-5 h-5 text-cyan-600" />
                <input type="month" value={currentMonth} onChange={(e) => setCurrentMonth(e.target.value)} disabled={!isAdmin} className="border-none bg-transparent font-black text-gray-800 outline-none cursor-pointer text-base" />
              </div>
              <button onClick={handleSaveData} disabled={isSaving || !isAdmin} className="flex items-center gap-2 px-8 py-2.5 bg-gray-900 text-white rounded-xl hover:bg-black transition-all font-black text-sm shadow-md disabled:bg-gray-300"><Save className="w-4 h-4" /> 전체 저장</button>
              <div className="h-8 w-px bg-gray-200 mx-1"></div>
              <div className="flex items-center gap-2 px-2">
                <FolderOpen className="w-5 h-5 text-fuchsia-500" />
                <select value="" onChange={(e) => setCurrentMonth(e.target.value)} className="border-none bg-transparent outline-none cursor-pointer text-gray-700 text-sm font-black focus:ring-0 hover:text-fuchsia-600 transition-colors">
                  <option value="" disabled>기록 불러오기</option>
                  {savedMonths.map(m => <option key={m} value={m}>{m} 정산분</option>)}
                </select>
                <button onClick={() => forceLoadData(currentMonth)} className="p-2 text-gray-400 hover:text-cyan-600 transition-colors" title="데이터 새로고침"><RefreshCw className="w-5 h-5" /></button>
              </div>
            </div>
          </div>
        </header>

        {/* 탭 네비게이션 */}
        <nav className="flex space-x-2 bg-white p-2 rounded-2xl shadow-sm mb-8 border border-gray-100 overflow-x-auto whitespace-nowrap">
          {[
            { id: 'prices', label: '1. 단가설정', icon: Settings },
            { id: 'orders', label: '2. 수량할당', icon: Truck },
            { id: 'billing', label: '3. 청구서(본사)', icon: FileText },
            { id: 'payment', label: '4. 결제서(회원사)', icon: CreditCard },
            { id: 'contacts', label: '5. 주소록', icon: BookOpen },
          ].map(t => (
            <button key={t.id} onClick={() => setActiveTab(t.id)} className={`flex-1 min-w-max flex items-center justify-center gap-2 py-4 px-4 rounded-xl font-black text-[15px] transition-all duration-300 ${activeTab === t.id ? 'bg-gradient-to-r from-fuchsia-600 to-cyan-500 text-white shadow-lg shadow-cyan-500/20 scale-[1.02] z-10' : 'text-gray-500 hover:bg-gray-50 hover:text-gray-800'}`}>
              <t.icon className="w-5 h-5" /> {t.label}
            </button>
          ))}
        </nav>

        <main className="bg-white rounded-[2rem] shadow-xl shadow-gray-200/50 border border-gray-100 p-8 min-h-[600px]">
          {/* TAB 1: 단가설정 */}
          {activeTab === 'prices' && (
            <div className="animate-in fade-in duration-500 space-y-12">
              <div>
                <h2 className="text-2xl font-black text-gray-900 flex items-center gap-3 mb-8 pb-4 border-b border-gray-100"><MapPin className="w-7 h-7 text-cyan-500" /> 급지별 10Kg 청구 단가 설정</h2>
                <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-7 gap-5">
                  {Object.entries(zonePrices).map(([zone, p]) => (
                    <div key={zone} className="border border-gray-100 rounded-3xl p-6 bg-white shadow-sm flex flex-col justify-center transition-all hover:shadow-md hover:border-cyan-200 group">
                      <div className="font-black text-center text-gray-400 mb-3 text-xs uppercase tracking-widest group-hover:text-cyan-600 transition-colors">{zone}</div>
                      <div className="relative">
                        <input type="text" disabled={!isAdmin} value={formatNumber(p.billing)} onChange={(e) => handleZonePriceChange(zone, 'billing', e.target.value)} onFocus={(e) => e.target.select()} className="w-full p-3 bg-gray-50 rounded-2xl text-center focus:bg-white focus:border-cyan-500 focus:ring-4 focus:ring-cyan-50 outline-none font-black text-2xl text-gray-900 disabled:bg-gray-50 transition-all border border-transparent" />
                        <span className="absolute -right-2 -bottom-2 text-[10px] font-black text-gray-300">원</span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
              <div>
                <h2 className="text-2xl font-black text-gray-900 flex items-center gap-3 mb-8 pb-4 border-b border-gray-100"><Settings className="w-7 h-7 text-fuchsia-500" /> 지역별 배송 급지 지정</h2>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
                  {REGION_ORDER.map(r => (
                    <div key={r} className="flex items-center justify-between p-5 border border-gray-100 rounded-3xl bg-white shadow-sm hover:shadow-md hover:border-fuchsia-200 transition-all group">
                      <span className="font-black text-gray-800 text-lg group-hover:text-fuchsia-700 transition-colors">{r}</span>
                      <div className="flex items-center gap-2">
                        <select disabled={!isAdmin} value={regions[r] || '2급지'} onChange={(e) => handleRegionZoneChange(r, e.target.value)} className="bg-fuchsia-50 text-fuchsia-700 border-none px-4 py-2 rounded-xl font-black text-sm focus:ring-4 focus:ring-fuchsia-100 outline-none cursor-pointer transition-all">
                          {Object.keys(zonePrices).map(k => <option key={k} value={k}>{k}</option>)}
                        </select>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* TAB 2: 수량할당 (통합 업로드 유지 및 회원사 스크롤 제거) */}
          {activeTab === 'orders' && (
            <div className="animate-in fade-in duration-500 space-y-8">
              <div className="flex flex-col xl:flex-row gap-6 mb-10">
                <div className="flex-1 bg-white border border-gray-200 rounded-3xl p-6 shadow-sm grid grid-cols-1 md:grid-cols-3 gap-6">
                  <div className="flex flex-col items-center justify-center p-5 bg-gradient-to-b from-fuchsia-50 to-white rounded-2xl border border-fuchsia-100 shadow-inner">
                    <span className="text-xs font-black text-fuchsia-600 uppercase mb-2 tracking-widest">차상위 배송 총합</span>
                    <span className="text-4xl font-black text-gray-900">{formatNumber(orderSummaries.povertyTotal)} <small className="text-sm font-bold text-gray-400">포</small></span>
                  </div>
                  <div className="flex flex-col items-center justify-center p-5 bg-gradient-to-b from-cyan-50 to-white rounded-2xl border border-cyan-100 shadow-inner">
                    <span className="text-xs font-black text-cyan-600 uppercase mb-2 tracking-widest">수급자 배송 총합</span>
                    <span className="text-4xl font-black text-gray-900">{formatNumber(orderSummaries.basicTotal)} <small className="text-sm font-bold text-gray-400">포</small></span>
                  </div>
                  <div className="flex flex-col items-center justify-center p-5 bg-gradient-to-br from-fuchsia-600 to-cyan-500 rounded-2xl shadow-lg text-white">
                    <span className="text-xs font-black text-white/80 uppercase mb-2 tracking-widest">전체 배송 총량</span>
                    <span className="text-5xl font-black drop-shadow-md">{formatNumber(orderSummaries.overallTotal)} <small className="text-base font-bold text-white/70">포</small></span>
                  </div>
                </div>

                <div className="xl:w-80 bg-gray-900 rounded-3xl p-8 shadow-xl flex flex-col justify-center gap-5">
                  <h4 className="text-white font-black text-base flex items-center gap-2"><Upload className="w-5 h-5 text-cyan-400" /> 통합 엑셀 업로드</h4>
                  <div className="grid grid-cols-1 gap-3">
                    <button onClick={downloadUploadTemplate} className="flex items-center justify-center gap-2 bg-white/10 hover:bg-white/20 text-white py-3 rounded-2xl text-sm font-black transition-all border border-white/10">
                      <FileDown className="w-4 h-4" /> 업로드 양식 받기
                    </button>
                    <button onClick={() => fileInputRef.current.click()} disabled={!isXlsxReady || !isAdmin} className="flex items-center justify-center gap-2 bg-gradient-to-r from-cyan-500 to-blue-500 hover:opacity-90 text-white py-3.5 rounded-2xl text-sm font-black shadow-lg shadow-cyan-500/30 transition-all disabled:opacity-50">
                      <Upload className="w-5 h-5" /> 엑셀 파일 올리기
                    </button>
                  </div>
                  <input type="file" ref={fileInputRef} onChange={handleExcelUpload} className="hidden" accept=".xlsx, .xls" />
                </div>
              </div>

              <div className="grid grid-cols-1 gap-8">
                {REGION_ORDER.map(r => {
                  const o = orders[r] || { basicQty: 0, povertyQty: 0 };
                  const tr = o.basicQty + o.povertyQty;
                  const cA = computedAllocations[r] || {};
                  const ta = Object.values(cA).reduce((a, b) => a + b, 0);
                  const balanced = tr === ta;
                  const z = regions[r] || '2급지';
                  return (
                    <div key={r} className={`border rounded-3xl overflow-hidden transition-all shadow-sm ${!balanced && tr > 0 ? 'border-fuchsia-300 shadow-fuchsia-100' : 'border-gray-200'}`}>
                      <div className="bg-gray-50/50 px-8 py-5 border-b border-gray-100 flex justify-between items-center">
                        <div className="flex items-center gap-4">
                          <h3 className="font-black text-gray-900 text-xl tracking-tight">{r}</h3>
                          <span className="text-xs font-black bg-white border border-gray-200 text-gray-600 px-3 py-1 rounded-full shadow-sm">{z}</span>
                        </div>
                        <div className="flex items-center gap-6 text-right">
                          <div className="flex flex-col items-end">
                            <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">지자체 합계</span>
                            <span className="text-2xl font-black text-gray-900">{formatNumber(tr)} <small className="text-xs font-bold text-gray-500">포</small></span>
                          </div>
                          <div className={`flex items-center gap-3 font-black px-6 py-2.5 rounded-full text-sm border shadow-sm ${tr === 0 ? 'bg-white text-gray-400 border-gray-200' : balanced ? 'bg-cyan-50 text-cyan-700 border-cyan-200' : 'bg-fuchsia-50 text-fuchsia-700 border-fuchsia-200'}`}>
                            배정: {formatNumber(ta)} / {formatNumber(tr)}
                          </div>
                        </div>
                      </div>
                      <div className="p-8 grid grid-cols-1 md:grid-cols-12 gap-10 bg-white">
                        <div className="md:col-span-5 space-y-5">
                           <div className="flex items-center justify-between gap-4 p-2">
                             <span className="text-sm font-black text-gray-400 uppercase tracking-widest">차상위용</span>
                             <input id={`poverty-${r}`} type="text" disabled={!isAdmin} value={formatNumber(o.povertyQty)} onChange={(e) => handleOrderChange(r, 'povertyQty', e.target.value)} onKeyDown={(e) => handleKeyDown(e, `poverty-${r}`)} onFocus={(e) => e.target.select()} className="w-40 p-3 bg-gray-50 border border-transparent rounded-2xl text-right text-2xl font-black text-fuchsia-700 focus:bg-white focus:border-fuchsia-400 focus:ring-4 focus:ring-fuchsia-50 outline-none disabled:opacity-50 transition-all" />
                           </div>
                           <div className="flex items-center justify-between gap-4 p-2">
                             <span className="text-sm font-black text-gray-400 uppercase tracking-widest">수급자용</span>
                             <input id={`basic-${r}`} type="text" disabled={!isAdmin} value={formatNumber(o.basicQty)} onChange={(e) => handleOrderChange(r, 'basicQty', e.target.value)} onKeyDown={(e) => handleKeyDown(e, `basic-${r}`)} onFocus={(e) => e.target.select()} className="w-40 p-3 bg-gray-50 border border-transparent rounded-2xl text-right text-2xl font-black text-cyan-700 focus:bg-white focus:border-cyan-400 focus:ring-4 focus:ring-cyan-50 outline-none disabled:opacity-50 transition-all" />
                           </div>
                        </div>
                        <div className="md:col-span-7 space-y-3">
                           {Object.entries(cA).map(([m, q]) => {
                             const isAuto = FIXED_MAPPING[r] || (r==='시흥시' && m==='사회적협동조합 행복나눔') || (r==='동대문구' && m==='웰쉐어 사회적협동조합');
                             return (
                               <div key={m} className={`flex items-center justify-between px-6 py-3 rounded-2xl border ${isAuto ? 'bg-gray-50 border-gray-100' : 'bg-white border-gray-200 shadow-sm hover:border-cyan-300 transition-colors'}`}>
                                 <div className="flex items-center gap-3 font-black text-sm text-gray-700">{isAuto ? <Lock className="w-4 h-4 text-gray-300"/> : <Truck className="w-4 h-4 text-cyan-500"/>} {m}</div>
                                 <div className="w-36">
                                   {isAuto ? <span className="block text-right font-black text-gray-900 text-2xl pr-4">{formatNumber(q)}</span> : (
                                     <input id={`alloc-${r}-${m}`} type="text" disabled={!isAdmin} value={formatNumber(q)} onChange={(e) => handleAllocationChange(r, m, e.target.value)} onKeyDown={(e) => handleKeyDown(e, `alloc-${r}-${m}`)} onFocus={(e) => e.target.select()} className="w-full p-2 border-b-2 border-gray-200 bg-transparent text-right text-2xl font-black text-gray-900 outline-none focus:border-gray-800 transition-all" />
                                   )}
                                 </div>
                               </div>
                             );
                           })}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* 하단 요약 (스크롤 제거) */}
              <div className="mt-16 border-t border-gray-200 pt-12 grid grid-cols-1 lg:grid-cols-2 gap-10">
                <div className="bg-white border border-gray-100 rounded-3xl p-8 shadow-md flex flex-col h-full">
                  <h4 className="font-black text-lg text-gray-800 mb-8 pb-4 border-b border-gray-100 flex items-center gap-2"><PieChart className="w-5 h-5 text-fuchsia-500" /> 권역별 합계 요약</h4>
                  <div className="space-y-4 flex-1 flex flex-col justify-center">
                    <div className="flex justify-between items-center p-5 rounded-2xl bg-gray-50 border border-gray-100">
                      <span className="font-black text-gray-600">서울특별시 합계</span>
                      <span className="text-3xl font-black text-gray-900">{formatNumber(orderSummaries.seoulTotal)} <small className="text-sm font-bold text-gray-400">포</small></span>
                    </div>
                    <div className="flex justify-between items-center p-5 rounded-2xl bg-gray-50 border border-gray-100">
                      <span className="font-black text-gray-600">경기도 합계</span>
                      <span className="text-3xl font-black text-gray-900">{formatNumber(orderSummaries.gyeonggiTotal)} <small className="text-sm font-bold text-gray-400">포</small></span>
                    </div>
                    <div className="flex justify-between items-center p-6 rounded-2xl bg-gradient-to-r from-gray-900 to-gray-800 shadow-xl mt-6">
                      <span className="font-black text-white text-lg tracking-tight">전체 지자체 물량 총합</span>
                      <span className="text-4xl font-black text-cyan-400">{formatNumber(orderSummaries.overallTotal)} <small className="text-sm font-bold text-white/50">포</small></span>
                    </div>
                  </div>
                </div>

                <div className="bg-white border border-gray-100 rounded-3xl p-8 shadow-md flex flex-col h-full">
                  <h4 className="font-black text-lg text-gray-800 mb-8 pb-4 border-b border-gray-100 flex justify-between items-center">
                    <div className="flex items-center gap-2"><Users className="w-5 h-5 text-cyan-500" /> 회원사별 배정 합계</div>
                    <span className="text-[10px] text-gray-500 bg-gray-100 px-3 py-1.5 rounded-full tracking-widest">가나다순</span>
                  </h4>
                  <div className="overflow-visible flex-1 flex flex-col justify-between">
                    <table className="w-full">
                      <tbody className="divide-y divide-gray-100">
                        {memberSummaries.map(m => (
                          <tr key={m.member} className="hover:bg-gray-50 transition-colors group">
                            <td className="py-4 px-2 font-black text-gray-700 text-base group-hover:text-cyan-700 transition-colors">{m.member}</td>
                            <td className="py-4 px-2 text-right font-black text-gray-900 text-xl">{formatNumber(m.qty)} <small className="text-xs font-bold text-gray-400">포</small></td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* TAB 3: 청구서 (엑셀 인라인 스타일 강제 적용) */}
          {activeTab === 'billing' && (
            <div className="animate-in fade-in duration-500 space-y-8">
              <div className="flex justify-between items-center bg-gray-900 p-10 rounded-[2rem] text-white shadow-xl relative overflow-hidden">
                <div className="absolute -right-10 -top-10 w-64 h-64 bg-gradient-to-br from-fuchsia-600/20 to-cyan-500/20 rounded-full blur-3xl"></div>
                <div className="relative z-10">
                  <h3 className="text-gray-400 font-black text-sm uppercase tracking-widest mb-2">Monthly Billing Report</h3>
                  <div className="text-6xl font-black tracking-tighter">{formatCur(billingReport.grandTotal.amount)}</div>
                </div>
                <div className="relative z-10 flex gap-3">
                  <button onClick={handleGenerateAnalysis} disabled={isAiLoading} className="flex items-center gap-2 bg-gradient-to-r from-fuchsia-600 to-purple-600 text-white px-8 py-4 rounded-2xl shadow-lg hover:scale-105 transition-all font-black text-sm disabled:opacity-50">
                    {isAiLoading ? <RefreshCw className="w-5 h-5 animate-spin" /> : <Sparkles className="w-5 h-5 text-yellow-300" />} ✨ AI 월간 브리핑
                  </button>
                  <button onClick={() => handleDownloadExcel('billing-table-export', '청구서(본사)')} className="flex items-center gap-2 bg-white text-gray-900 px-8 py-4 rounded-2xl shadow-lg hover:scale-105 transition-all font-black text-sm">
                    <Download className="w-5 h-5" /> 엑셀 다운로드
                  </button>
                </div>
              </div>

              {aiAnalysis && (
                <div className="bg-fuchsia-50/50 border border-fuchsia-100 rounded-[2rem] p-8 shadow-sm animate-in zoom-in duration-300 relative">
                  <button onClick={() => setAiAnalysis("")} className="absolute top-6 right-6 text-gray-400 hover:text-gray-800"><X className="w-5 h-5" /></button>
                  <h3 className="font-black text-fuchsia-800 flex items-center gap-2 mb-4"><Sparkles className="w-5 h-5" /> AI 경영 보고 요약</h3>
                  <div className="text-gray-700 font-medium leading-loose whitespace-pre-wrap">{aiAnalysis}</div>
                </div>
              )}
              
              <div className="overflow-x-auto rounded-[2rem] border-2 border-gray-200 shadow-sm bg-white">
                <table id="billing-table-export" style={excelStyles.table}>
                   <thead>
                      <tr>
                        <td colSpan="8" style={excelStyles.titleRow} className="py-4 px-8 text-sm">희망나르미 본사 발행 청구 내역</td>
                        <td colSpan="2" style={{...excelStyles.titleRow, textAlign: 'right'}} className="py-4 px-8 text-sm">{formattedMonthStr}</td>
                      </tr>
                      <tr>
                        <th style={excelStyles.th}>행정시</th>
                        <th style={excelStyles.th}>행정구</th>
                        <th style={excelStyles.th}>구분</th>
                        <th style={excelStyles.th}>Kg</th>
                        <th style={excelStyles.th}>품 명</th>
                        <th style={excelStyles.th}>포수</th>
                        <th style={excelStyles.th}>금액</th>
                        <th style={excelStyles.th}>VAT</th>
                        <th style={excelStyles.th}>합계</th>
                        <th style={excelStyles.th}>비고</th>
                      </tr>
                   </thead>
                   <tbody>
                      {billingReport.report.map((item, idx) => (
                        <React.Fragment key={idx}>
                           <tr>
                              <td rowSpan="3" style={excelStyles.td}>{item.city}</td>
                              <td rowSpan="3" style={excelStyles.td}>{item.region}</td>
                              <td style={excelStyles.td}>차상위</td>
                              <td style={excelStyles.td}>10Kg</td>
                              <td style={excelStyles.tdLeft}>차상위 배송비</td>
                              <td style={{...excelStyles.tdRight, fontWeight: 'bold'}}>{formatNumber(item.poverty.qty)}</td>
                              <td style={excelStyles.tdRight}>{formatNumber(item.poverty.sup)}</td>
                              <td style={excelStyles.tdRight}>{formatNumber(item.poverty.vat)}</td>
                              <td style={{...excelStyles.tdRight, fontWeight: 'bold'}}>{formatNumber(item.poverty.tot)}</td>
                              <td style={excelStyles.td}></td>
                           </tr>
                           <tr>
                              <td style={excelStyles.td}>수급자</td>
                              <td style={excelStyles.td}>10Kg</td>
                              <td style={excelStyles.tdLeft}>수급자 배송비</td>
                              <td style={{...excelStyles.tdRight, fontWeight: 'bold'}}>{formatNumber(item.basic.qty)}</td>
                              <td style={excelStyles.tdRight}>{formatNumber(item.basic.sup)}</td>
                              <td style={excelStyles.tdRight}>{formatNumber(item.basic.vat)}</td>
                              <td style={{...excelStyles.tdRight, fontWeight: 'bold'}}>{formatNumber(item.basic.tot)}</td>
                              <td style={excelStyles.td}></td>
                           </tr>
                           <tr>
                              <td colSpan="3" style={{...excelStyles.subTotalRow, border: '1pt solid #000'}}>합 계</td>
                              <td style={{...excelStyles.subTotalRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(item.sum.qty)}</td>
                              <td style={{...excelStyles.subTotalRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(item.sum.supply)}</td>
                              <td style={{...excelStyles.subTotalRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(item.sum.vat)}</td>
                              <td style={{...excelStyles.subTotalRow, border: '1pt solid #000', textAlign: 'right', color: '#000080'}}>{formatNumber(item.sum.amount)}</td>
                              <td style={{...excelStyles.subTotalRow, border: '1pt solid #000'}}></td>
                           </tr>
                           {idx !== billingReport.report.length - 1 && (
                             <tr><td colSpan="10" style={excelStyles.emptyRow}></td></tr>
                           )}
                        </React.Fragment>
                      ))}
                   </tbody>
                   <tfoot>
                      <tr>
                        <td colSpan="2" style={{...excelStyles.regionSumRow, border: '1pt solid #000', padding: '10px'}}>경기도 합계</td>
                        <td colSpan="3" style={{...excelStyles.regionSumRow, border: '1pt solid #000'}}>10Kg / 배송비</td>
                        <td style={{...excelStyles.regionSumRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(billingReport.gTotal.qty)}</td>
                        <td style={{...excelStyles.regionSumRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(billingReport.gTotal.supply)}</td>
                        <td style={{...excelStyles.regionSumRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(billingReport.gTotal.vat)}</td>
                        <td style={{...excelStyles.regionSumRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(billingReport.gTotal.amount)}</td>
                        <td style={{...excelStyles.regionSumRow, border: '1pt solid #000'}}></td>
                      </tr>
                      <tr>
                        <td colSpan="2" style={{...excelStyles.regionSumRow, border: '1pt solid #000', padding: '10px'}}>서울시 합계</td>
                        <td colSpan="3" style={{...excelStyles.regionSumRow, border: '1pt solid #000'}}>10Kg / 배송비</td>
                        <td style={{...excelStyles.regionSumRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(billingReport.sTotal.qty)}</td>
                        <td style={{...excelStyles.regionSumRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(billingReport.sTotal.supply)}</td>
                        <td style={{...excelStyles.regionSumRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(billingReport.sTotal.vat)}</td>
                        <td style={{...excelStyles.regionSumRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(billingReport.sTotal.amount)}</td>
                        <td style={{...excelStyles.regionSumRow, border: '1pt solid #000'}}></td>
                      </tr>
                      <tr>
                        <td colSpan="2" style={{...excelStyles.grandSumRow, border: '1pt solid #000', padding: '15px'}}>전 체 합 계</td>
                        <td colSpan="3" style={{...excelStyles.grandSumRow, border: '1pt solid #000'}}>10Kg / 배송비</td>
                        <td style={{...excelStyles.grandSumRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(billingReport.grandTotal.qty)}</td>
                        <td style={{...excelStyles.grandSumRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(billingReport.grandTotal.supply)}</td>
                        <td style={{...excelStyles.grandSumRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(billingReport.grandTotal.vat)}</td>
                        <td style={{...excelStyles.grandSumRow, border: '1pt solid #000', textAlign: 'right', fontSize: '14pt'}}>{formatNumber(billingReport.grandTotal.amount)}</td>
                        <td style={{...excelStyles.grandSumRow, border: '1pt solid #000'}}></td>
                      </tr>
                   </tfoot>
                </table>
              </div>
            </div>
          )}

          {/* TAB 4: 결제서 (엑셀 인라인 스타일 강제 적용) */}
          {activeTab === 'payment' && (
            <div className="animate-in fade-in duration-500 space-y-6">
              <div className="flex justify-between items-center mb-6">
                <h2 className="text-2xl font-black text-gray-900 flex items-center gap-3"><CreditCard className="w-8 h-8 text-cyan-600" /> 결제 명세서 <small className="text-gray-400 font-bold ml-2">(조합사 발행용)</small></h2>
                <button onClick={() => handleDownloadExcel('payment-table-export', '결제서(회원사)')} className="flex items-center gap-2 bg-gray-900 text-white px-8 py-3 rounded-2xl shadow-xl hover:scale-105 transition-all font-black text-sm"><Download className="w-5 h-5 text-cyan-400" /> 엑셀 다운로드</button>
              </div>
              <div className="overflow-x-auto rounded-[2rem] border-2 border-gray-200 shadow-sm bg-white">
                <table id="payment-table-export" style={excelStyles.table}>
                  {billingSummary.sorted.map((m, mIdx) => (
                    <tbody key={m.member}>
                      {mIdx > 0 && <tr><td colSpan="9" style={excelStyles.emptyRow}></td></tr>}
                      <tr>
                        <td colSpan="7" style={{...excelStyles.titleRow, textAlign: 'left'}} className="py-3.5 px-8 text-[13px]">조합사 -&gt; 웰쉐어 발행 내역</td>
                        <td colSpan="2" style={{...excelStyles.titleRow, textAlign: 'right', color: '#666666'}} className="py-3.5 px-8 text-[12px]">{formattedMonthStr}</td>
                      </tr>
                      <tr>
                        <th style={excelStyles.th} className="w-40">조합사</th>
                        <th style={excelStyles.th} className="w-48">행정구</th>
                        <th style={excelStyles.th} className="w-16">Kg</th>
                        <th style={excelStyles.th} className="w-24">품명</th>
                        <th style={excelStyles.th} className="w-24">포수</th>
                        <th style={excelStyles.th} className="w-36">금액</th>
                        <th style={excelStyles.th} className="w-36">VAT</th>
                        <th style={excelStyles.th} className="w-40">합계</th>
                        <th style={excelStyles.th} className="w-24">비고</th>
                      </tr>
                      {m.regions.map((r, i) => (
                        <tr key={r.region}>
                          {i === 0 && <td rowSpan={m.regions.length} style={{...excelStyles.td, fontWeight: 'bold'}}>{m.member}</td>}
                          <td style={{...excelStyles.tdLeft, fontWeight: 'bold'}}>{getFullRegionName(r.region)}</td>
                          <td style={excelStyles.td}>10Kg</td>
                          <td style={excelStyles.td}>배송비</td>
                          <td style={{...excelStyles.tdRight, fontWeight: 'bold'}}>{formatNumber(r.qty)}</td>
                          <td style={excelStyles.tdRight}>{formatNumber(r.supplyValue)}</td>
                          <td style={excelStyles.tdRight}>{formatNumber(r.vatValue)}</td>
                          <td style={{...excelStyles.tdRight, color: '#000080', fontWeight: 'bold'}}>{formatNumber(r.finalRowTotal)}</td>
                          <td style={excelStyles.td}></td>
                        </tr>
                      ))}
                      <tr>
                        <td colSpan="2" style={{...excelStyles.regionSumRow, border: '1pt solid #000', padding: '10px'}}>합 계</td>
                        <td style={{...excelStyles.regionSumRow, border: '1pt solid #000'}}>10Kg</td>
                        <td style={{...excelStyles.regionSumRow, border: '1pt solid #000'}}>배송비</td>
                        <td style={{...excelStyles.regionSumRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(m.totalQty)}</td>
                        <td style={{...excelStyles.regionSumRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(m.totalSupply)}</td>
                        <td style={{...excelStyles.regionSumRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(m.totalVat)}</td>
                        <td style={{...excelStyles.regionSumRow, border: '1pt solid #000', textAlign: 'right', color: '#000080', fontSize: '12pt'}}>{formatNumber(m.totalAmount)}</td>
                        <td style={{...excelStyles.regionSumRow, border: '1pt solid #000'}}></td>
                      </tr>
                    </tbody>
                  ))}
                  <tfoot>
                    <tr>
                      <td colSpan="4" style={{...excelStyles.grandSumRow, border: '1pt solid #000', padding: '15px'}}>전 체 합 계</td>
                      <td style={{...excelStyles.grandSumRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(billingSummary.grandTotalQty)}</td>
                      <td style={{...excelStyles.grandSumRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(billingSummary.grandTotalSupply)}</td>
                      <td style={{...excelStyles.grandSumRow, border: '1pt solid #000', textAlign: 'right'}}>{formatNumber(billingSummary.grandTotalVat)}</td>
                      <td style={{...excelStyles.grandSumRow, border: '1pt solid #000', textAlign: 'right', fontSize: '15pt'}}>{formatNumber(billingSummary.grandTotalAmount)}</td>
                      <td style={{...excelStyles.grandSumRow, border: '1pt solid #000'}}></td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>
          )}

          {/* TAB 5: 주소록 */}
          {activeTab === 'contacts' && (
            <div className="animate-in fade-in duration-500 space-y-10">
               <div className="bg-gray-900 rounded-[2rem] p-10 text-white flex flex-col md:flex-row items-center justify-between gap-10 border-b-8 border-cyan-500 shadow-2xl relative overflow-hidden">
                 <div className="absolute -left-20 -bottom-20 w-64 h-64 bg-cyan-500/20 rounded-full blur-3xl"></div>
                 <div className="flex items-center gap-8 relative z-10">
                   <div className="p-5 bg-white/10 rounded-3xl backdrop-blur-sm border border-white/10"><PhoneCall className="w-12 h-12 text-cyan-400" /></div>
                   <div><h3 className="text-2xl font-black mb-1 tracking-tighter text-cyan-100">희망나르미 본사</h3><div className="font-black text-4xl text-white tracking-tighter">02-324-2155</div></div>
                 </div>
                 <div className="bg-white/5 p-8 rounded-3xl border border-white/10 text-center md:text-right min-w-[300px] relative z-10">
                   <div className="text-xs text-fuchsia-400 font-black uppercase mb-2 tracking-widest">General Manager</div>
                   <div className="font-black text-3xl">김애영 과장</div>
                   <div className="text-cyan-400 font-black text-xl mt-2 tracking-tighter">010-4511-4469</div>
                 </div>
               </div>
               
               <div className="grid grid-cols-1 gap-12">
                  <div className="border border-gray-100 rounded-3xl overflow-hidden shadow-md bg-white overflow-x-auto">
                    <div className="bg-fuchsia-50 p-6 border-b border-fuchsia-100 font-black text-xl text-fuchsia-900 flex items-center gap-3"><Building2 className="w-6 h-6" /> 지자체 담당 공무원 연락처</div>
                    <table className="w-full text-left text-sm">
                      <thead className="bg-gray-50 border-b-2"><tr><th className="p-5 font-black text-gray-400">지역</th><th className="p-5 text-center font-black text-gray-400">차수</th><th className="p-5 font-black text-gray-400">담당자</th><th className="p-5 font-black text-gray-400">연락처</th><th className="p-5 font-black text-gray-400 text-center">동작</th></tr></thead>
                      <tbody>{GOV_CONTACTS.map(g => (<tr key={g.region} className="border-b border-gray-100 last:border-0 hover:bg-fuchsia-50/30 transition-colors"><td className="p-5 font-black text-gray-800 text-lg">{g.region}</td><td className="p-5 text-center"><span className={`px-4 py-1.5 rounded-full font-black text-[10px] border ${g.order==='1차'?'bg-cyan-50 text-cyan-700 border-cyan-200':'bg-fuchsia-50 text-fuchsia-700 border-fuchsia-200'}`}>{g.order}</span></td><td className="p-5 font-black text-gray-600 text-lg">{g.manager}</td><td className="p-5 font-black text-fuchsia-600 text-xl tracking-tighter">{g.phone}</td><td className="p-5 text-center"><button onClick={() => handleDraftMessage(g)} className="bg-white border border-gray-200 text-gray-600 hover:text-fuchsia-600 hover:border-fuchsia-300 px-3 py-1.5 rounded-lg text-xs font-black transition-colors flex items-center justify-center gap-1 mx-auto"><Sparkles className="w-3 h-3" /> ✨ 메시지 작성</button></td></tr>))}</tbody>
                    </table>
                  </div>

                  <div className="border border-gray-100 rounded-3xl overflow-hidden shadow-md bg-white overflow-x-auto">
                    <div className="bg-cyan-50 p-6 border-b border-cyan-100 font-black text-xl text-cyan-900 flex items-center gap-3"><Users className="w-6 h-6" /> 회원사(조합사) 담당자 연락망</div>
                    <table className="w-full text-left min-w-[1000px]">
                      <thead className="bg-gray-50 border-b-2"><tr><th className="p-5 font-black text-gray-400 w-56 font-black">기관명</th><th className="p-5 font-black text-gray-400 font-black">배송 담당 구역</th><th className="p-5 font-black text-gray-400 font-black">담당자</th><th className="p-5 font-black text-gray-400 font-black">연락처</th><th className="p-5 font-black text-gray-400 font-black text-center">동작</th></tr></thead>
                      <tbody>{CONTACTS.map((c, idx) => { 
                        const show = idx === 0 || CONTACTS[idx-1].agency !== c.agency; 
                        const rs = CONTACTS.filter(x => x.agency === c.agency).length; 
                        return (
                          <tr key={idx} className="border-b border-gray-100 last:border-0 hover:bg-cyan-50/30 transition-colors">
                            {show ? <td rowSpan={rs} className="p-5 bg-white border-r border-gray-100 font-black text-center align-middle text-cyan-900 text-lg">{c.agency}</td> : null}
                            <td className="p-5 font-black text-gray-800 text-base">{c.region} {c.detail && <span className="ml-3 text-[11px] bg-gray-50 border border-gray-200 text-gray-500 px-3 py-1 rounded-lg font-normal tracking-wide">구역: {c.detail}</span>}</td>
                            <td className="p-5 font-black text-gray-600 text-base">{c.manager || '-'}</td>
                            <td className="p-5 font-black text-cyan-600 text-lg tracking-tighter">{c.phone || '-'}</td>
                            <td className="p-5 text-center"><button onClick={() => handleDraftMessage(c)} disabled={!c.manager} className="bg-white border border-gray-200 text-gray-600 hover:text-cyan-600 hover:border-cyan-300 disabled:opacity-30 px-3 py-1.5 rounded-lg text-xs font-black transition-colors flex items-center justify-center gap-1 mx-auto"><Sparkles className="w-3 h-3" /> ✨ 메시지 작성</button></td>
                          </tr>
                        );
                      })}</tbody>
                    </table>
                  </div>
               </div>
            </div>
          )}
        </main>
      </div>
    </div>
  );
}