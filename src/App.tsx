import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  Zap, 
  AlertTriangle, 
  Download, 
  Search, 
  Filter, 
  BarChart3, 
  FileSpreadsheet, 
  Layers, 
  Building2, 
  CheckCircle2, 
  Bell, 
  RefreshCw, 
  Sliders, 
  ArrowRight,
  TrendingDown,
  Info,
  ShieldCheck,
  ExternalLink,
  PlusCircle,
  X,
  FileUp,
  ClipboardPaste,
  LogIn,
  LogOut,
  Database,
  CheckCircle,
  Calculator,
  Shuffle
} from 'lucide-react';
import * as maplibregl from 'maplibre-gl';
import 'maplibre-gl/dist/maplibre-gl.css';
import maplibreWorkerUrl from 'maplibre-gl/dist/maplibre-gl-worker.mjs?worker&url';

maplibregl.setWorkerUrl(maplibreWorkerUrl);
import { TransformerRecord, PEA_OFFICES, generateInitialTransformers, evaluateTransformerRightsizing, formatPeaTransformerCode } from './data/mockTransformers';
import { exportTransformersToExcel, exportTRMatchingToExcel } from './utils/exportExcel';
import { parseCsvRows } from './utils/csvParser';
import { 
  initAuth, 
  googleSignIn, 
  logout, 
  getAccessToken 
} from './services/firebaseAuth';
import { 
  TARGET_SPREADSHEET_ID, 
  fetchSpreadsheetInfo, 
  fetchSheetValues, 
  parseSheetDataToTransformers, 
  findMatchingSheetTitles,
  SpreadsheetInfo 
} from './services/googleSheets';
import type { User } from 'firebase/auth';

export const format2Decimals = (val: number | string | null | undefined): string => {
  if (val === null || val === undefined || val === '') return '-';
  const num = Number(val);
  if (isNaN(num)) return '-';
  return num.toFixed(2);
};

interface MapLibreContainerProps {
  center: { lat: number; lng: number };
  zoom: number;
  transformers: TransformerRecord[];
  mapOffice: string;
  mapRating: string;
  mapSearchQuery: string;
  onSelectTransformer: (tr: TransformerRecord) => void;
}

function MapLibreContainer({
  center,
  zoom,
  transformers,
  mapOffice,
  mapRating,
  mapSearchQuery,
  onSelectTransformer
}: MapLibreContainerProps) {
  const mapContainerRef = useRef<HTMLDivElement>(null);
  const mapRef = useRef<maplibregl.Map | null>(null);
  const markersRef = useRef<maplibregl.Marker[]>([]);

  // Initialize Map
  useEffect(() => {
    if (!mapContainerRef.current) return;

    const map = new maplibregl.Map({
      container: mapContainerRef.current,
      style: 'https://basemaps.cartocdn.com/gl/positron-gl-style/style.json',
      center: [center.lng, center.lat],
      zoom: zoom,
    });

    map.addControl(new maplibregl.NavigationControl(), 'top-right');
    mapRef.current = map;

    return () => {
      map.remove();
      mapRef.current = null;
    };
  }, []);

  // Update Center & Zoom when search or office selection changes
  useEffect(() => {
    if (mapRef.current) {
      mapRef.current.flyTo({
        center: [center.lng, center.lat],
        zoom: zoom,
        essential: true,
      });
    }
  }, [center.lat, center.lng, zoom]);

  // Update Markers
  useEffect(() => {
    const map = mapRef.current;
    if (!map) return;

    // Remove existing markers
    markersRef.current.forEach(m => m.remove());
    markersRef.current = [];

    // Add markers for eligible transformers
    transformers.forEach(tr => {
      if (mapOffice !== 'all' && tr.feaOffice !== mapOffice) return;
      if (mapRating !== 'all' && tr.ratedKva !== Number(mapRating)) return;
      if (mapSearchQuery.trim()) {
        const q = mapSearchQuery.toLowerCase().trim();
        const matchCode = tr.trCode.toLowerCase().includes(q);
        const matchOffice = tr.feaOffice.toLowerCase().includes(q);
        const matchLoc = (tr.location || '').toLowerCase().includes(q);
        const matchRaw = tr.trCode.replace(/-/g, '').includes(q.replace(/-/g, ''));
        if (!matchCode && !matchOffice && !matchLoc && !matchRaw) return;
      }

      const lat = (tr as any).latitude;
      const lng = (tr as any).longitude;
      if (!lat || !lng) return;

      const evalRes = evaluateTransformerRightsizing(tr);
      if (!evalRes.isEligible) return;

      // Custom Pin Marker Element (separated container from visual pin to prevent transform overrides)
      const el = document.createElement('div');
      el.className = 'custom-map-marker-container';

      const pinInner = document.createElement('div');
      pinInner.className = 'custom-map-marker-pin';
      pinInner.style.width = '34px';
      pinInner.style.height = '34px';
      pinInner.style.borderRadius = '50%';
      pinInner.style.backgroundColor = '#059669';
      pinInner.style.border = '2px solid #047857';
      pinInner.style.color = '#ffffff';
      pinInner.style.display = 'flex';
      pinInner.style.alignItems = 'center';
      pinInner.style.justifyContent = 'center';
      pinInner.style.fontSize = '11px';
      pinInner.style.fontWeight = 'bold';
      pinInner.style.fontFamily = 'monospace';
      pinInner.style.boxShadow = '0 4px 10px rgba(0,0,0,0.25)';
      pinInner.style.cursor = 'pointer';
      pinInner.style.transition = 'transform 0.15s ease-in-out';
      pinInner.title = `PEA: ${tr.trCode} (${tr.ratedKva} kVA) - ${tr.feaOffice}`;
      pinInner.innerText = `${tr.ratedKva}`;

      pinInner.addEventListener('mouseenter', () => {
        pinInner.style.transform = 'scale(1.2)';
      });
      pinInner.addEventListener('mouseleave', () => {
        pinInner.style.transform = 'scale(1)';
      });

      el.appendChild(pinInner);

      // Construct Popup Content Element
      const recommendedKva = evalRes.recommendedKva ?? '-';
      const currentLoad = format2Decimals(tr.year3LoadPct);
      const newLoad = evalRes.newLoadPct != null ? format2Decimals(evalRes.newLoadPct) : '-';

      const popupNode = document.createElement('div');
      popupNode.className = 'p-1 text-slate-800 font-sans min-w-[240px]';
      popupNode.innerHTML = `
        <div class="border-b border-slate-100 pb-2 mb-2">
          <div class="flex items-center justify-between gap-2">
            <span class="text-[11px] font-bold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
              เข้าเกณฑ์ Rightsizing
            </span>
            <span class="text-[10px] text-slate-500 font-mono font-semibold">${tr.feaOffice}</span>
          </div>
          <h4 class="text-sm font-bold text-slate-900 mt-1 flex items-center gap-1.5">
            <span>PEA: ${tr.trCode}</span>
          </h4>
          <p class="text-[11px] text-slate-600 mt-0.5">
            สังกัด: <strong class="text-slate-800">${tr.feaOffice}</strong> (${tr.feeder || 'F-1'})
          </p>
        </div>

        <div class="space-y-1.5 text-xs">
          <div class="flex items-center justify-between bg-slate-50 p-2 rounded-lg border border-slate-200">
            <span class="text-slate-600 text-[11px]">ขนาดพิกัด (kVA):</span>
            <span class="font-bold text-slate-800 font-mono flex items-center gap-1">
              <span class="text-slate-400 line-through">${tr.ratedKva} kVA</span>
              <span class="text-emerald-700 font-extrabold">➔ ${recommendedKva} kVA</span>
            </span>
          </div>

          <div class="grid grid-cols-2 gap-1.5">
            <div class="bg-amber-50 p-2 rounded-lg border border-amber-200">
              <div class="text-[10px] text-amber-800 font-semibold">โหลดปัจจุบัน</div>
              <div class="text-xs font-bold text-amber-900 font-mono mt-0.5">${currentLoad}%</div>
            </div>
            <div class="bg-emerald-50 p-2 rounded-lg border border-emerald-200">
              <div class="text-[10px] text-emerald-800 font-semibold">โหลดใหม่ (&lt;50%)</div>
              <div class="text-xs font-bold text-emerald-900 font-mono mt-0.5">${newLoad}%</div>
            </div>
          </div>

          ${tr.location ? `<div class="text-[10px] text-slate-500 truncate pt-0.5">📍 ${tr.location}</div>` : ''}
        </div>

        <button type="button" class="mt-2.5 w-full py-1.5 px-3 bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white font-medium text-xs rounded-md shadow-xs transition-colors flex items-center justify-center gap-1.5 cursor-pointer select-btn">
          <span>ดูรายละเอียด / จำลองการลดขนาด</span>
          <span>➔</span>
        </button>
      `;

      const selectBtn = popupNode.querySelector('.select-btn');
      if (selectBtn) {
        selectBtn.addEventListener('click', (ev) => {
          ev.stopPropagation();
          onSelectTransformer(tr);
        });
      }

      const popup = new maplibregl.Popup({
        offset: 25,
        closeButton: true,
        closeOnClick: false,
        maxWidth: '300px'
      }).setDOMContent(popupNode);

      const marker = new maplibregl.Marker({ element: el })
        .setLngLat([lng, lat])
        .setPopup(popup)
        .addTo(map);

      markersRef.current.push(marker);
    });
  }, [transformers, mapOffice, mapRating, mapSearchQuery, onSelectTransformer]);

  return <div ref={mapContainerRef} style={{ width: '100%', height: '100%' }} className="rounded-xl overflow-hidden shadow-inner border border-slate-200" />;
}

export default function App() {
  const [transformers, setTransformers] = useState<TransformerRecord[]>([]);
  const [peaOffices, setPeaOffices] = useState<string[]>(PEA_OFFICES);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'database' | 'target' | 'pea34' | 'alerts' | 'trmatching' | 'map'>('dashboard');

  const [selectedOffice, setSelectedOffice] = useState<string>('all');
  const [selectedRating, setSelectedRating] = useState<string>('all');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [filterEligibleOnly, setFilterEligibleOnly] = useState<boolean>(false);
  const [isSyncing, setIsSyncing] = useState<boolean>(false);
  const [syncStatus, setSyncStatus] = useState<string>('กำลังเริ่มต้นฐานข้อมูลจาก Google Sheets ID: 1dKqK6soiSHHnaqemJbr8GCHCv7or4Ro8Htess1g-ovw');
  const [selectedTransformer, setSelectedTransformer] = useState<TransformerRecord | null>(null);
  const [showAddModal, setShowAddModal] = useState<boolean>(false);
  const [showAlertModal, setShowAlertModal] = useState<boolean>(false);
  const [showImportModal, setShowImportModal] = useState<boolean>(false);
  const [pastedText, setPastedText] = useState<string>('');
  const [importStatus, setImportStatus] = useState<string>('');

  // Selected Sheet Tab Names for Google Sheets
  const [selectedBaseTab, setSelectedBaseTab] = useState<string>('ข้อมูลปีปัจจุบัน');
  const [selectedPrior1Tab, setSelectedPrior1Tab] = useState<string>('ข้อมูล 1 ปีก่อน');
  const [selectedPrior2Tab, setSelectedPrior2Tab] = useState<string>('ข้อมูล 2 ปีก่อน');

  // Google Workspace OAuth State
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [accessToken, setAccessToken] = useState<string | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState<boolean>(false);
  const [spreadsheetMeta, setSpreadsheetMeta] = useState<SpreadsheetInfo | null>(null);
  const [lastSyncTime, setLastSyncTime] = useState<string>('');

  // New transformer form state
  const [newTr, setNewTr] = useState({
    trCode: '52-001928',
    feaOffice: PEA_OFFICES[0],
    feeder: 'F-1 (ยะลา)',
    ratedKva: 100,
    year1LoadPct: 22.5,
    year2LoadPct: 24.1,
    year3LoadPct: 21.0,
    location: 'จุดติดตั้งตัวอย่าง'
  });

  useEffect(() => {
    // 1. Initial base data
    const data = generateInitialTransformers();
    setTransformers(data);

    // 2. Setup Google Workspace Auth listener
    const unsubscribe = initAuth(
      (user, token) => {
        setCurrentUser(user);
        setAccessToken(token);
        // Automatically sync from Google Sheets ID 1dKqK6soiSHHnaqemJbr8GCHCv7or4Ro8Htess1g-ovw
        handleSyncGoogleSheet(token);
      },
      () => {
        setCurrentUser(null);
        setAccessToken(null);
      }
    );

    // Attempt initial sync
    handleSyncGoogleSheet();

    return () => {
      if (typeof unsubscribe === 'function') unsubscribe();
    };
  }, []);

  // Compute evaluation statistics
  const stats = useMemo(() => {
    let total = transformers.length;
    let target100 = 0;
    let target160 = 0;
    let target250 = 0;
    let eligibleList: { tr: TransformerRecord; evalRes: any }[] = [];

    transformers.forEach(tr => {
      const evalRes = evaluateTransformerRightsizing(tr);
      if (evalRes.isEligible) {
        eligibleList.push({ tr, evalRes });
        if (tr.ratedKva === 100) target100++;
        if (tr.ratedKva === 160) target160++;
        if (tr.ratedKva === 250) target250++;
      }
    });

    // Sort eligibleList by newLoadPct ascending (lowest to highest)
    eligibleList.sort((a, b) => (a.evalRes.newLoadPct ?? 0) - (b.evalRes.newLoadPct ?? 0));

    // Breakdown by 34 PEA offices
    const officeMap: { [key: string]: number } = {};
    PEA_OFFICES.forEach(off => { officeMap[off] = 0; });
    eligibleList.forEach(item => {
      officeMap[item.tr.feaOffice] = (officeMap[item.tr.feaOffice] || 0) + 1;
    });

    return {
      total,
      totalEligible: eligibleList.length,
      target100,
      target160,
      target250,
      eligibleList,
      officeMap
    };
  }, [transformers]);

  // Filtered transformers for table view
  const filteredTransformers = useMemo(() => {
    const list = transformers.filter(tr => {
      if (selectedOffice !== 'all' && tr.feaOffice !== selectedOffice) return false;
      if (selectedRating !== 'all' && tr.ratedKva.toString() !== selectedRating) return false;
      if (searchQuery.trim() !== '') {
        const q = searchQuery.toLowerCase();
        const matchCode = tr.trCode.toLowerCase().includes(q);
        const matchOffice = tr.feaOffice.toLowerCase().includes(q);
        const matchFeeder = tr.feeder.toLowerCase().includes(q);
        if (!matchCode && !matchOffice && !matchFeeder) return false;
      }
      if (filterEligibleOnly) {
        const res = evaluateTransformerRightsizing(tr);
        if (!res.isEligible) return false;
      }
      return true;
    });

    // Sort by year3LoadPct (โหลดปีปัจจุบัน (%)) ascending (lowest to highest)
    list.sort((a, b) => (a.year3LoadPct ?? 0) - (b.year3LoadPct ?? 0));

    return list;
  }, [transformers, selectedOffice, selectedRating, searchQuery, filterEligibleOnly]);

  // States for Rightsizing Targets Search & Filter
  const [targetSearchQuery, setTargetSearchQuery] = useState<string>('');
  const [targetOffice, setTargetOffice] = useState<string>('all');
  const [targetRating, setTargetRating] = useState<string>('all');
  const [targetPage, setTargetPage] = useState<number>(1);
  const TARGET_ITEMS_PER_PAGE = 20;

  // States for Map Tab Search & Filter
  const [mapSearchQuery, setMapSearchQuery] = useState<string>('');
  const [mapOffice, setMapOffice] = useState<string>('all');
  const [mapRating, setMapRating] = useState<string>('all');
  const [mapCenter, setMapCenter] = useState<{ lat: number; lng: number }>({ lat: 6.845, lng: 101.233 });
  const [mapZoom, setMapZoom] = useState<number>(9);

  const filteredEligibleList = useMemo(() => {
    return stats.eligibleList.filter(({ tr }) => {
      if (targetOffice !== 'all' && tr.feaOffice !== targetOffice) return false;
      if (targetRating !== 'all' && tr.ratedKva !== Number(targetRating)) return false;
      if (targetSearchQuery.trim()) {
        const q = targetSearchQuery.toLowerCase().trim();
        const matchCode = tr.trCode.toLowerCase().includes(q);
        const matchOffice = tr.feaOffice.toLowerCase().includes(q);
        const matchFeeder = tr.feeder.toLowerCase().includes(q);
        const matchLoc = (tr.location || '').toLowerCase().includes(q);
        if (!matchCode && !matchOffice && !matchFeeder && !matchLoc) return false;
      }
      return true;
    });
  }, [stats.eligibleList, targetOffice, targetRating, targetSearchQuery]);

  const totalTargetPages = Math.ceil(filteredEligibleList.length / TARGET_ITEMS_PER_PAGE) || 1;
  const paginatedEligibleList = useMemo(() => {
    const start = (targetPage - 1) * TARGET_ITEMS_PER_PAGE;
    return filteredEligibleList.slice(start, start + TARGET_ITEMS_PER_PAGE);
  }, [filteredEligibleList, targetPage]);

  useEffect(() => {
    setTargetPage(1);
  }, [targetSearchQuery, targetOffice, targetRating]);

  // States for Database Table Pagination (20 items per page)
  const [dbPage, setDbPage] = useState<number>(1);
  const DB_ITEMS_PER_PAGE = 20;
  const totalDbPages = Math.ceil(filteredTransformers.length / DB_ITEMS_PER_PAGE) || 1;
  const paginatedFilteredTransformers = useMemo(() => {
    const start = (dbPage - 1) * DB_ITEMS_PER_PAGE;
    return filteredTransformers.slice(start, start + DB_ITEMS_PER_PAGE);
  }, [filteredTransformers, dbPage]);

  // States for Alert Notification Center Pagination (20 items per page)
  const [alertPage, setAlertPage] = useState<number>(1);
  const ALERT_ITEMS_PER_PAGE = 20;
  const totalAlertPages = Math.ceil(stats.eligibleList.length / ALERT_ITEMS_PER_PAGE) || 1;
  const paginatedAlertList = useMemo(() => {
    const start = (alertPage - 1) * ALERT_ITEMS_PER_PAGE;
    return stats.eligibleList.slice(start, start + ALERT_ITEMS_PER_PAGE);
  }, [stats.eligibleList, alertPage]);

  // Compute TR Matching pairs: matching eligible rightsizing transformers with overloaded transformers (>80% load) in the same office
  // Conditions:
  // 1. targetTr.ratedKva MUST equal sourceEval.recommendedKva (พิกัดเดิมของปลายทางตรงกับพิกัดที่แนะนำของต้นทาง)
  // 2. เมื่อจับคู่กันแล้ว % โหลดใหม่ทั้ง 2 เครื่อง ต้องน้อยกว่า 50% (< 50%):
  //    - เครื่องต้นทาง: sourceEval.newLoadPct < 50
  //    - เครื่องปลายทาง: targetNewLoadPct = (targetActualKva / sourceTr.ratedKva) * 100 < 50
  // 3. จัดลำดับเลือกหม้อแปลงปลายทางที่มีโหลดมากสุดเป็นอันดับแรก
  const trMatches = useMemo(() => {
    const matches: {
      office: string;
      sourceTr: TransformerRecord;
      sourceEval: any;
      targetTr: TransformerRecord;
      targetNewLoadPct: number;
      distanceKm: number | null;
    }[] = [];

    const officeMap = new Map<string, { eligible: { tr: TransformerRecord; evalRes: any }[]; overloaded: TransformerRecord[] }>();

    transformers.forEach(tr => {
      const evalRes = evaluateTransformerRightsizing(tr);
      if (!officeMap.has(tr.feaOffice)) {
        officeMap.set(tr.feaOffice, { eligible: [], overloaded: [] });
      }
      const entry = officeMap.get(tr.feaOffice)!;
      if (evalRes.isEligible && (evalRes.newLoadPct ?? 100) < 50) {
        entry.eligible.push({ tr, evalRes });
      }
      if (tr.year3LoadPct > 80) {
        entry.overloaded.push(tr);
      }
    });

    // Haversine formula to calculate distance in km between two lat/lng points
    const calculateDistanceKm = (lat1?: number, lon1?: number, lat2?: number, lon2?: number): number | null => {
      if (lat1 === undefined || lon1 === undefined || lat2 === undefined || lon2 === undefined) return null;
      if (isNaN(lat1) || isNaN(lon1) || isNaN(lat2) || isNaN(lon2)) return null;
      const R = 6371; // Radius of the earth in km
      const dLat = (lat2 - lat1) * (Math.PI / 180);
      const dLon = (lon2 - lon1) * (Math.PI / 180);
      const a =
        Math.sin(dLat / 2) * Math.sin(dLat / 2) +
        Math.cos(lat1 * (Math.PI / 180)) * Math.cos(lat2 * (Math.PI / 180)) *
        Math.sin(dLon / 2) * Math.sin(dLon / 2);
      const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
      const d = R * c; // Distance in km
      return Number(d.toFixed(2));
    };

    officeMap.forEach((data, office) => {
      const usedTargetIds = new Set<string>();

      data.eligible.forEach(({ tr: sourceTr, evalRes: sourceEval }) => {
        const recommendedKva = sourceEval.recommendedKva;
        if (!recommendedKva) return;

        // Filter and map candidate target transformers matching criteria:
        // 1. targetTr.ratedKva === recommendedKva
        // 2. not used yet
        // 3. targetTr.id !== sourceTr.id
        // 4. when paired, % new load of target machine < 50%
        const candidates = data.overloaded
          .filter(t => {
            if (t.ratedKva !== recommendedKva) return false;
            if (usedTargetIds.has(t.id)) return false;
            if (t.id === sourceTr.id) return false;

            const targetActualKva = (t.year3LoadPct / 100) * t.ratedKva;
            const newTargetLoad = (targetActualKva / sourceTr.ratedKva) * 100;
            return newTargetLoad < 50;
          })
          .map(t => {
            const dist = calculateDistanceKm(
              (sourceTr as any).latitude,
              (sourceTr as any).longitude,
              (t as any).latitude,
              (t as any).longitude
            );
            return {
              targetTr: t,
              dist: dist !== null ? dist : Infinity
            };
          });

        // Sort candidate target transformers by shortest distance first, then highest load %
        candidates.sort((a, b) => {
          if (a.dist !== b.dist) {
            return a.dist - b.dist; // Closest distance first
          }
          return b.targetTr.year3LoadPct - a.targetTr.year3LoadPct; // Highest load % tie breaker
        });

        if (candidates.length > 0) {
          const matchedTarget = candidates[0].targetTr;
          usedTargetIds.add(matchedTarget.id);
          const targetActualKva = (matchedTarget.year3LoadPct / 100) * matchedTarget.ratedKva;
          const targetNewLoadPct = Number(((targetActualKva / sourceTr.ratedKva) * 100).toFixed(2));
          const distanceKm = candidates[0].dist !== Infinity ? candidates[0].dist : null;

          matches.push({
            office,
            sourceTr,
            sourceEval,
            targetTr: matchedTarget,
            targetNewLoadPct,
            distanceKm
          });
        }
      });
    });

    return matches;
  }, [transformers]);

  // Breakdown of TR matches by rating
  const matchStatsByRating = useMemo(() => {
    let m100 = 0;
    let m160 = 0;
    let m250 = 0;
    trMatches.forEach(m => {
      if (m.sourceTr.ratedKva === 100) m100++;
      else if (m.sourceTr.ratedKva === 160) m160++;
      else if (m.sourceTr.ratedKva === 250) m250++;
    });
    return { m100, m160, m250, total: trMatches.length };
  }, [trMatches]);

  // States for TR Matching Search & Filter
  const [matchSearchQuery, setMatchSearchQuery] = useState<string>('');
  const [matchOffice, setMatchOffice] = useState<string>('all');
  const [matchRating, setMatchRating] = useState<string>('all');
  const [matchPage, setMatchPage] = useState<number>(1);
  const MATCH_ITEMS_PER_PAGE = 20;

  const filteredMatches = useMemo(() => {
    return trMatches.filter(({ office, sourceTr, targetTr }) => {
      if (matchOffice !== 'all' && office !== matchOffice) return false;
      if (matchRating !== 'all') {
        const r = Number(matchRating);
        if (sourceTr.ratedKva !== r && targetTr.ratedKva !== r) return false;
      }
      if (matchSearchQuery.trim()) {
        const q = matchSearchQuery.toLowerCase().trim();
        const matchSourceCode = sourceTr.trCode.toLowerCase().includes(q);
        const matchTargetCode = targetTr.trCode.toLowerCase().includes(q);
        const matchOfficeName = office.toLowerCase().includes(q);
        if (!matchSourceCode && !matchTargetCode && !matchOfficeName) return false;
      }
      return true;
    });
  }, [trMatches, matchOffice, matchRating, matchSearchQuery]);

  const totalMatchPages = Math.ceil(filteredMatches.length / MATCH_ITEMS_PER_PAGE) || 1;
  const paginatedMatches = useMemo(() => {
    const start = (matchPage - 1) * MATCH_ITEMS_PER_PAGE;
    return filteredMatches.slice(start, start + MATCH_ITEMS_PER_PAGE);
  }, [filteredMatches, matchPage]);

  useEffect(() => {
    setMatchPage(1);
  }, [matchSearchQuery, matchOffice, matchRating]);

  useEffect(() => {
    setDbPage(1);
  }, [selectedOffice, selectedRating, searchQuery, filterEligibleOnly]);

  // States for Multi-sheet import
  const [importTab, setImportTab] = useState<'current' | 'year1' | 'year2' | 'all'>('current');
  const [currentYearInput, setCurrentYearInput] = useState<string>('');
  const [year1Input, setYear1Input] = useState<string>('');
  const [year2Input, setYear2Input] = useState<string>('');

  // Helper to extract [PEA หม้อแปลง => % load] from secondary sheets (ข้อมูล 1 ปีก่อน, ข้อมูล 2 ปีก่อน)
  const extractLoadMapFromSheet = (text: string): Map<string, number> => {
    const map = new Map<string, number>();
    if (!text || !text.trim()) return map;
    const rows = parseCsvRows(text);
    if (rows.length < 1) return map;
    
    // Find header row (0, 1, or 2)
    let headerRowIdx = 0;
    for (let r = 0; r < Math.min(3, rows.length); r++) {
      const rowStr = rows[r].join(' ');
      if (/pea|หม้อแปลง|รหัส|tr|%|load|โหลด|ลำดับ/i.test(rowStr)) {
        headerRowIdx = r;
        break;
      }
    }

    const headerCols = rows[headerRowIdx] || [];
    let codeIdx = headerCols.findIndex(c => /pea.*หม้อแปลง|รหัส.*หม้อแปลง|หม้อแปลง|tr.*code/i.test(c));
    if (codeIdx === -1) codeIdx = 0;

    // Column E is index 4: ข้อมูล 1 ปีก่อน (%) และ ข้อมูล 2 ปีก่อน (%) ให้ใช้ข้อมูลจากชีท คอลัมน์ E
    const COL_E_INDEX = 4;
    let loadIdx = COL_E_INDEX;
    if (headerCols.length <= COL_E_INDEX) {
      loadIdx = headerCols.findIndex(c => /%|load|โหลด|peak|สูงสุด/i.test(c));
      if (loadIdx === -1) {
        loadIdx = headerCols.length > 1 ? 1 : 0;
      }
    }

    for (let i = headerRowIdx + 1; i < rows.length; i++) {
      const cols = rows[i];
      if (!cols || cols.length === 0) continue;
      const code = (cols[codeIdx] || cols[0])?.trim();
      if (!code || /pea.*หม้อแปลง|ลำดับ|หัวข้อ|รวม/i.test(code)) continue;
      const formattedCode = formatPeaTransformerCode(code);
      const digitsOnly = code.replace(/\D/g, '');

      // ดึง % โหลด จาก คอลัมน์ E (index 4) เป็นอันดับแรก ตามข้อกำหนด
      let val = NaN;
      if (cols.length > COL_E_INDEX && cols[COL_E_INDEX] !== undefined && cols[COL_E_INDEX].trim() !== '') {
        const numStr = cols[COL_E_INDEX].replace(/[^0-9.]/g, '');
        const parsed = Number(numStr);
        if (!isNaN(parsed) && numStr.length > 0) {
          val = parsed;
        }
      }
      if (isNaN(val) && loadIdx !== -1 && loadIdx < cols.length && cols[loadIdx] !== undefined) {
        const numStr = cols[loadIdx].replace(/[^0-9.]/g, '');
        const parsed = Number(numStr);
        if (!isNaN(parsed) && numStr.length > 0) {
          val = parsed;
        }
      }

      if (!isNaN(val)) {
        map.set(code, val);
        map.set(formattedCode, val);
        if (digitsOnly) map.set(digitsOnly, val);
      }
    }
    return map;
  };

  // Main Parser & Processor:
  // 1. Uses PEA หม้อแปลง from sheet "ข้อมูลปีปัจจุบัน" (Column A) as the baseline starting dataset.
  // 2. Uses กฟฟ. from sheet "ข้อมูลปีปัจจุบัน" (Column AG, Index 32)
  // 3. Analyzes backward 2 years into "ข้อมูล 1 ปีก่อน" and "ข้อมูล 2 ปีก่อน".
  const processBaseAndHistoricalData = (
    baseText: string,
    prior1Text?: string,
    prior2Text?: string
  ): boolean => {
    const rows = parseCsvRows(baseText);
    if (rows.length < 2) return false;

    // Find header row (0, 1, or 2)
    let headerRowIdx = 0;
    for (let r = 0; r < Math.min(4, rows.length); r++) {
      const rowStr = rows[r].join(' ');
      if (/pea|หม้อแปลง|รหัส|tr|ลำดับ|กฟฟ|การไฟฟ้า|kva|โหลด|feeder|สถานที่/i.test(rowStr) || rows[r].length > 5) {
        headerRowIdx = r;
        break;
      }
    }

    const headerCols = rows[headerRowIdx] || [];

    // Column A is index 0: PEA หม้อแปลง (ข้อมูลตั้งต้น)
    const colAIndex = 0;

    // Column B is index 1: สถานที่ติดตั้ง (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ B)
    let colLocationIndex = 1;
    const foundLocIdx = headerCols.findIndex(c => /สถานที่|ติดตั้ง|พิกัดที่ตั้ง|location|site/i.test(c));
    if (foundLocIdx !== -1 && foundLocIdx !== 0) {
      colLocationIndex = foundLocIdx;
    } else if (headerCols.length > 1) {
      colLocationIndex = 1;
    }

    // Column AG is index 32: กฟฟ. (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ AG)
    const COL_AG_INDEX = 32;
    let colOfficeIndex = COL_AG_INDEX;
    if (headerCols.length > COL_AG_INDEX) {
      colOfficeIndex = COL_AG_INDEX;
    } else {
      const foundOfficeIdx = headerCols.findIndex(c => /^(กฟฟ\.?|การไฟฟ้า|สังกัด.*กฟฟ)/i.test(c.trim()) || /กฟ/i.test(c));
      if (foundOfficeIdx !== -1) {
        colOfficeIndex = foundOfficeIdx;
      }
    }

    // Column E is index 4: ข้อมูลปีปัจจุบัน (%) [ตั้งต้น] (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ E)
    const COL_E_INDEX = 4;
    let colY3CurrentIndex = COL_E_INDEX;
    if (headerCols.length > COL_E_INDEX) {
      colY3CurrentIndex = COL_E_INDEX;
    } else {
      const foundYCurrent = headerCols.findIndex(c => /ปัจจุบัน|current|y3|ปี.*3|load.*3|โหลด.*3/i.test(c));
      if (foundYCurrent !== -1) {
        colY3CurrentIndex = foundYCurrent;
      }
    }

    // Column D is index 3: พิกัดเดิม (kVA) (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ D)
    const COL_D_INDEX = 3;
    let colKvaIndex = COL_D_INDEX;
    if (headerCols.length > COL_D_INDEX) {
      colKvaIndex = COL_D_INDEX;
    } else {
      const foundKva = headerCols.findIndex((c, idx) => idx !== COL_E_INDEX && /kva|พิกัด|ขนาด/i.test(c));
      if (foundKva !== -1) {
        colKvaIndex = foundKva;
      }
    }

    // Feeder column (avoid Column D: พิกัดเดิม and Column E: โหลดปีปัจจุบัน)
    let colFeederIndex = headerCols.findIndex((c, idx) => idx !== COL_D_INDEX && idx !== COL_E_INDEX && /feeder|ฟีดเดอร์|วงจร/i.test(c));
    if (colFeederIndex === -1) colFeederIndex = headerCols.length > 2 && colKvaIndex !== 2 ? 2 : -1;

    let colY2Index = headerCols.findIndex(c => /1.*ปี|y2|load.*2|โหลด.*2|ปี.*2/i.test(c));
    if (colY2Index === -1) colY2Index = 6;

    let colY1Index = headerCols.findIndex(c => /2.*ปี|y1|load.*1|โหลด.*1|ปี.*1/i.test(c));
    if (colY1Index === -1) colY1Index = 5;

    // Extract lookup maps from historical sheets if provided
    const prior1Map = prior1Text ? extractLoadMapFromSheet(prior1Text) : new Map<string, number>();
    const prior2Map = prior2Text ? extractLoadMapFromSheet(prior2Text) : new Map<string, number>();

    const getHistoricalVal = (map: Map<string, number>, trCode: string, rawCode: string): number | undefined => {
      if (map.has(trCode)) return map.get(trCode);
      if (map.has(rawCode)) return map.get(rawCode);
      const digits = rawCode.replace(/\D/g, '');
      if (digits && map.has(digits)) return map.get(digits);
      const trDigits = trCode.replace(/\D/g, '');
      if (trDigits && map.has(trDigits)) return map.get(trDigits);
      return undefined;
    };

    const parsedTransformers: TransformerRecord[] = [];

    for (let i = headerRowIdx + 1; i < rows.length; i++) {
      const cols = rows[i];
      if (!cols || cols.length === 0) continue;

      // Column A: PEA หม้อแปลง (ข้อมูลตั้งต้น - ใส่เลขให้ครบทุกหลัก)
      const rawCode = (cols[colAIndex] || cols[0])?.trim();
      if (!rawCode || /pea.*หม้อแปลง|ลำดับ|หัวข้อ/i.test(rawCode)) continue;
      const trCode = formatPeaTransformerCode(rawCode);

      // Column B: สถานที่ติดตั้ง (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ B)
      const rawLocation = (colLocationIndex < cols.length && cols[colLocationIndex] !== undefined) ? cols[colLocationIndex].trim() : '';
      const locationVal = rawLocation || (cols[1] ? cols[1].trim() : `จุดติดตั้งแถวที่ ${i + 1}`);

      // Column AG: กฟฟ. (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ AG)
      // Strictly prioritize Column AG (index 32)
      let cleanOffice = '';
      if (cols.length > COL_AG_INDEX && cols[COL_AG_INDEX] !== undefined && cols[COL_AG_INDEX].trim() !== '') {
        cleanOffice = cols[COL_AG_INDEX].trim();
      } else if (colOfficeIndex !== -1 && colOfficeIndex < cols.length && cols[colOfficeIndex] && cols[colOfficeIndex].trim() !== '') {
        cleanOffice = cols[colOfficeIndex].trim();
      }
      if (!cleanOffice) {
        // Check if any column contains known PEA office
        const foundKnown = cols.find(c => peaOffices.some(po => c.includes(po)));
        if (foundKnown) {
          cleanOffice = foundKnown.trim();
        } else {
          cleanOffice = peaOffices[i % peaOffices.length];
        }
      }
      cleanOffice = cleanOffice.replace(/^"|"$/g, '').trim();

      // พิกัดเดิม (kVA) (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ D)
      let kvaVal = NaN;
      if (cols.length > COL_D_INDEX && cols[COL_D_INDEX] !== undefined && cols[COL_D_INDEX].trim() !== '') {
        const cleanVal = cols[COL_D_INDEX].replace(/[^0-9.]/g, '');
        const parsed = Number(cleanVal);
        if (!isNaN(parsed) && cleanVal.length > 0 && parsed > 0) {
          kvaVal = parsed;
        }
      }
      if (isNaN(kvaVal) && colKvaIndex !== -1 && colKvaIndex < cols.length && cols[colKvaIndex] !== undefined) {
        const cleanVal = cols[colKvaIndex].replace(/[^0-9.]/g, '');
        const parsed = Number(cleanVal);
        if (!isNaN(parsed) && cleanVal.length > 0 && parsed > 0) {
          kvaVal = parsed;
        }
      }
      if (isNaN(kvaVal) || !kvaVal) {
        kvaVal = [100, 160, 250][i % 3];
      }

      // 1. Current Year load % (ข้อมูลตั้งต้น จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ E)
      let yCurrent = NaN;
      if (cols.length > COL_E_INDEX && cols[COL_E_INDEX] !== undefined && cols[COL_E_INDEX].trim() !== '') {
        const numStr = cols[COL_E_INDEX].replace(/[^0-9.]/g, '');
        const val = Number(numStr);
        if (!isNaN(val) && numStr.length > 0) {
          yCurrent = val;
        }
      }
      if (isNaN(yCurrent) && colY3CurrentIndex < cols.length && cols[colY3CurrentIndex] !== undefined) {
        const numStr = cols[colY3CurrentIndex].replace(/[^0-9.]/g, '');
        const val = Number(numStr);
        if (!isNaN(val) && numStr.length > 0) {
          yCurrent = val;
        }
      }
      if (isNaN(yCurrent) || !yCurrent) {
        yCurrent = Number((16 + ((i + 4) % 13)).toFixed(2));
      }

      // 2. Backward Analysis 1 year prior (วิเคราะห์ย้อนกลับ: จากชีท ข้อมูล 1 ปีก่อน คอลัมน์ E แมปด้วย PEA หม้อแปลง)
      let y1Ago = getHistoricalVal(prior1Map, trCode, rawCode);
      if (y1Ago === undefined) {
        const fromBase = colY2Index < cols.length ? Number(cols[colY2Index]?.replace(/[^0-9.]/g, '')) : NaN;
        y1Ago = !isNaN(fromBase) && fromBase > 0 ? Number(fromBase.toFixed(2)) : Number((14 + ((i + 2) % 15)).toFixed(2));
      }

      // 3. Backward Analysis 2 years prior (วิเคราะห์ย้อนกลับ: จากชีท ข้อมูล 2 ปีก่อน คอลัมน์ E แมปด้วย PEA หม้อแปลง)
      let y2Ago = getHistoricalVal(prior2Map, trCode, rawCode);
      if (y2Ago === undefined) {
        const fromBase = colY1Index < cols.length ? Number(cols[colY1Index]?.replace(/[^0-9.]/g, '')) : NaN;
        y2Ago = !isNaN(fromBase) && fromBase > 0 ? Number(fromBase.toFixed(2)) : Number((15 + (i % 14)).toFixed(2));
      }

      // AH is index 33, AI is index 34
      const lat = cols.length > 33 ? Number(cols[33]) : undefined;
      const lng = cols.length > 34 ? Number(cols[34]) : undefined;

      parsedTransformers.push({
        id: `row_${i}_${Date.now()}`,
        trCode, // PEA หม้อแปลง จากคอลัมน์ A (ข้อมูลตั้งต้น)
        feaOffice: cleanOffice, // กฟฟ. (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ AG)
        feeder: (colFeederIndex !== -1 && colFeederIndex < cols.length && cols[colFeederIndex]) ? cols[colFeederIndex].trim() : `F-${(i % 5) + 1}`,
        ratedKva: [50, 100, 160, 250, 315, 400, 500].includes(kvaVal) ? kvaVal : (kvaVal > 0 ? kvaVal : 100),
        year1LoadPct: y2Ago, // ข้อมูล 2 ปีก่อน (วิเคราะห์ย้อนหลัง 2 ปี)
        year2LoadPct: y1Ago, // ข้อมูล 1 ปีก่อน (วิเคราะห์ย้อนหลัง 1 ปี)
        year3LoadPct: yCurrent, // ข้อมูลปีปัจจุบัน (ข้อมูลตั้งต้น)
        location: locationVal, // สถานที่ติดตั้ง (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ B)
        latitude: !isNaN(lat!) ? lat : undefined,
        longitude: !isNaN(lng!) ? lng : undefined,
        installationDate: '2022-01-01',
        status: (!locationVal.includes('จัดสรร') && !locationVal.includes('หมู่บ้าน') && !locationVal.includes('โครงการ') && !locationVal.includes('สถานีอัดประจุ') && !locationVal.toUpperCase().includes('EV') && y2Ago < 30 && y1Ago < 30 && yCurrent < 30 && [100, 160, 250].includes(kvaVal)) ? 'Target Rightsizing' : 'Normal'
      });
    }

    if (parsedTransformers.length > 0) {
      setTransformers(parsedTransformers);
      const uniqueOffices = Array.from(new Set(parsedTransformers.map(t => t.feaOffice).filter(Boolean)));
      if (uniqueOffices.length > 0) {
        setPeaOffices(uniqueOffices);
      }
      return true;
    }
    return false;
  };

  // Google Sheets OAuth Sign-In handler
  const handleGoogleLogin = async () => {
    setIsLoggingIn(true);
    try {
      const res = await googleSignIn();
      if (res) {
        setCurrentUser(res.user);
        setAccessToken(res.accessToken);
        await handleSyncGoogleSheet(res.accessToken);
      }
    } catch (err: any) {
      console.error('Google Sign-In failed', err);
      alert(`การเข้าสู่ระบบ Google ขัดข้อง: ${err?.message || err}`);
    } finally {
      setIsLoggingIn(false);
    }
  };

  // Google Sign-Out handler
  const handleGoogleLogout = async () => {
    await logout();
    setCurrentUser(null);
    setAccessToken(null);
    setSpreadsheetMeta(null);
    setSyncStatus('ออกจากระบบ Google เรียบร้อยแล้ว');
  };

  // Google Sheets Sync handler:
  // Strictly queries Google Sheet ID 1dKqK6soiSHHnaqemJbr8GCHCv7or4Ro8Htess1g-ovw
  // Uses "ข้อมูลปีปัจจุบัน" Column A (PEA หม้อแปลง) as baseline, Column B as สถานที่ติดตั้ง,
  // and retroactively analyzes backward 2 years against "ข้อมูล 1 ปีก่อน" and "ข้อมูล 2 ปีก่อน"
  const handleSyncGoogleSheet = async (
    overrideToken?: string,
    customBase?: string,
    customPrior1?: string,
    customPrior2?: string
  ) => {
    setIsSyncing(true);
    const token = overrideToken || accessToken || (await getAccessToken());
    const sheetId = TARGET_SPREADSHEET_ID;

    // 1. If Google OAuth token is available, use Google Sheets API v4
    if (token) {
      setSyncStatus(`กำลังเข้าถึง Google Sheet ID: ${sheetId} ผ่าน Google Sheets API...`);
      try {
        const meta = await fetchSpreadsheetInfo(sheetId, token);
        setSpreadsheetMeta(meta);

        const tabTitles = meta.sheets.map(s => s.title);
        const autoMatch = findMatchingSheetTitles(tabTitles);
        
        // Find base sheet (ข้อมูลปีปัจจุบัน)
        const baseTitle = customBase || selectedBaseTab || autoMatch.baseTitle;
        // Find historical 1 (ข้อมูล 1 ปีก่อน)
        const prior1Title = customPrior1 || selectedPrior1Tab || autoMatch.prior1Title || '';
        // Find historical 2 (ข้อมูล 2 ปีก่อน)
        const prior2Title = customPrior2 || selectedPrior2Tab || autoMatch.prior2Title || '';

        // Keep local state in sync
        if (baseTitle) setSelectedBaseTab(baseTitle);
        if (prior1Title) setSelectedPrior1Tab(prior1Title);
        if (prior2Title) setSelectedPrior2Tab(prior2Title);

        setSyncStatus(`กำลังดึงข้อมูลชีท "${baseTitle}" (PEA หม้อแปลง คอลัมน์ A ตั้งต้น) และวิเคราะห์ย้อนหลัง 2 ปี (${prior1Title}, ${prior2Title})...`);

        const baseRows = await fetchSheetValues(sheetId, baseTitle, token);
        let prior1Rows: string[][] = [];
        let prior2Rows: string[][] = [];

        if (prior1Title && tabTitles.includes(prior1Title)) {
          try {
            prior1Rows = await fetchSheetValues(sheetId, prior1Title, token);
          } catch (e) {
            console.warn('Could not fetch prior 1 sheet', e);
          }
        }

        if (prior2Title && tabTitles.includes(prior2Title)) {
          try {
            prior2Rows = await fetchSheetValues(sheetId, prior2Title, token);
          } catch (e) {
            console.warn('Could not fetch prior 2 sheet', e);
          }
        }

        const parsed = parseSheetDataToTransformers(baseRows, prior1Rows, prior2Rows);
        if (parsed.transformers.length > 0) {
          setTransformers(parsed.transformers);
          if (parsed.officesFound.length > 0) {
            setPeaOffices(parsed.officesFound);
          }
          const nowStr = new Date().toLocaleTimeString('th-TH');
          setLastSyncTime(nowStr);
          setSyncStatus(`ซิงค์ข้อมูลสำเร็จ: ฐานข้อมูลจาก Google Sheets ID: ${sheetId} ชีท "${baseTitle}" (คอลัมน์ A เป็น PEA หม้อแปลง) วิเคราะห์ย้อนหลัง 2 ปี (${prior1Title ? `"${prior1Title}", ` : ''}${prior2Title ? `"${prior2Title}", ` : ''}) โหลดแล้ว ${parsed.transformers.length.toLocaleString()} รายการ (เวลา ${nowStr})`);
          setIsSyncing(false);
          return;
        }
      } catch (apiErr: any) {
        console.error('Sheets API Error:', apiErr);
        if (apiErr.message && apiErr.message.includes('401')) {
          setAccessToken(null);
          sessionStorage.removeItem('g_sheets_access_token');
          setSyncStatus('Session Google หมดอายุ กรุณากด "เข้าสู่ระบบ Google" เพื่อเชื่อมต่อใหม่');
        } else {
          setSyncStatus(`Google Sheets API แจ้งเตือน: ${apiErr.message || 'ไม่สามารถอ่านข้อมูล'} - กำลังตรวจสอบช่องทางสำรอง...`);
        }
      }
    }

    // 2. Fallback / Public export attempt
    try {
      const baseSheetName = encodeURIComponent('ข้อมูลปีปัจจุบัน');
      const prior1SheetName = encodeURIComponent('ข้อมูล 1 ปีก่อน');
      const prior2SheetName = encodeURIComponent('ข้อมูล 2 ปีก่อน');

      // Fetch base sheet (ข้อมูลปีปัจจุบัน)
      let baseCsv = '';
      try {
        const respBase = await fetch(`https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${baseSheetName}`);
        if (respBase.ok) {
          const text = await respBase.text();
          if (!text.startsWith('<!DOCTYPE') && !text.includes('ServiceLogin')) {
            baseCsv = text;
          }
        }
      } catch {
        // continue
      }

      // If specific sheet name failed, try default export
      if (!baseCsv) {
        try {
          const respDefault = await fetch(`https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv`);
          if (respDefault.ok) {
            const text = await respDefault.text();
            if (!text.startsWith('<!DOCTYPE') && !text.includes('ServiceLogin')) {
              baseCsv = text;
            }
          }
        } catch {
          // continue
        }
      }

      // Fetch historical sheet 1 (ข้อมูล 1 ปีก่อน)
      let prior1Csv = '';
      try {
        const resp1 = await fetch(`https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${prior1SheetName}`);
        if (resp1.ok) {
          const text = await resp1.text();
          if (!text.startsWith('<!DOCTYPE') && !text.includes('ServiceLogin')) {
            prior1Csv = text;
          }
        }
      } catch {
        // continue
      }

      // Fetch historical sheet 2 (ข้อมูล 2 ปีก่อน)
      let prior2Csv = '';
      try {
        const resp2 = await fetch(`https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${prior2SheetName}`);
        if (resp2.ok) {
          const text = await resp2.text();
          if (!text.startsWith('<!DOCTYPE') && !text.includes('ServiceLogin')) {
            prior2Csv = text;
          }
        }
      } catch {
        // continue
      }

      let success = false;
      if (baseCsv) {
        success = processBaseAndHistoricalData(baseCsv, prior1Csv, prior2Csv);
      }

      if (success) {
        setSyncStatus('ซิงค์ข้อมูลสำเร็จ: ใช้ข้อมูล PEA หม้อแปลง จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ A เป็นข้อมูลตั้งต้น และวิเคราะห์ย้อนกลับ 2 ปี (ข้อมูล 1 ปีก่อน & 2 ปีก่อน) เรียบร้อยแล้ว');
      } else {
        setSyncStatus(`ฐานข้อมูล: Google Sheet ID 1dKqK6soiSHHnaqemJbr8GCHCv7or4Ro8Htess1g-ovw (กด "เชื่อมต่อ Google Account" ด้านบนเพื่อดึงข้อมูลสด)`);
      }
    } catch {
      setSyncStatus(`เชื่อมต่อฐานข้อมูล Google Sheet: 1dKqK6soiSHHnaqemJbr8GCHCv7or4Ro8Htess1g-ovw`);
    } finally {
      setIsSyncing(false);
    }
  };

  const handleApprovePlan = (trId: string) => {
    setTransformers(prev => prev.map(t => t.id === trId ? { ...t, status: 'Action Planned' } : t));
    alert('บันทึกอนุมัติแผนการปรับลดขนาดหม้อแปลงเรียบร้อยแล้ว');
  };

  const handleAddTransformer = (e: React.FormEvent) => {
    e.preventDefault();
    const formattedCode = formatPeaTransformerCode(newTr.trCode);
    const newRecord: TransformerRecord = {
      id: `tr_custom_${Date.now()}`,
      trCode: formattedCode,
      feaOffice: newTr.feaOffice,
      feeder: newTr.feeder,
      ratedKva: Number(newTr.ratedKva),
      year1LoadPct: Number(newTr.year1LoadPct),
      year2LoadPct: Number(newTr.year2LoadPct),
      year3LoadPct: Number(newTr.year3LoadPct),
      location: newTr.location,
      installationDate: new Date().toISOString().split('T')[0],
      status: 'Normal'
    };
    const evalRes = evaluateTransformerRightsizing(newRecord);
    if (evalRes.isEligible) {
      newRecord.status = 'Target Rightsizing';
    }

    setTransformers([newRecord, ...transformers]);
    setShowAddModal(false);
    alert('เพิ่มข้อมูลหม้อแปลงสำเร็จ');
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col font-['Sarabun',sans-serif]">
      {/* Top Header */}
      <header className="bg-slate-900 text-white border-b border-slate-800 sticky top-0 z-30 shadow-md">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center space-x-3">
            <div className="w-10 h-10 rounded-lg bg-emerald-500 flex items-center justify-center shadow">
              <Zap className="w-6 h-6 text-white" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="font-bold text-lg tracking-tight text-white">PEA Transformer Rightsizing</h1>
                <span className="text-xs bg-emerald-500/20 text-emerald-300 font-medium px-2 py-0.5 rounded border border-emerald-500/30">
                  34 กฟฟ. Analytics
                </span>
              </div>
              <p className="text-xs text-slate-400">ระบบวิเคราะห์และวางแผนปรับลดขนาดหม้อแปลง (100, 160, 250 kVA)</p>
            </div>
          </div>

          <div className="flex items-center space-x-2.5">
            {/* Google Account Login / Status */}
            {currentUser ? (
              <div className="flex items-center gap-2 bg-slate-800/90 px-3 py-1.5 rounded-lg border border-slate-700">
                <div className="w-6 h-6 rounded-full bg-emerald-500/20 text-emerald-300 font-bold flex items-center justify-center text-[10px] border border-emerald-500/40">
                  {currentUser.displayName ? currentUser.displayName[0] : (currentUser.email ? currentUser.email[0].toUpperCase() : 'G')}
                </div>
                <div className="text-[11px] leading-tight text-slate-300 hidden md:block max-w-[140px] truncate" title={currentUser.email || ''}>
                  {currentUser.email}
                </div>
                <button
                  onClick={handleGoogleLogout}
                  className="text-slate-400 hover:text-rose-400 text-[11px] px-1 py-0.5 rounded cursor-pointer transition-colors"
                  title="ออกจากระบบ Google"
                >
                  <LogOut className="w-3.5 h-3.5" />
                </button>
              </div>
            ) : (
              <button
                onClick={handleGoogleLogin}
                disabled={isLoggingIn}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-blue-600 hover:bg-blue-500 text-white shadow-xs transition-all cursor-pointer disabled:opacity-50"
                title="เข้าสู่ระบบ Google เพื่อเชื่อมต่อ Google Sheet ID: 1dKqK6soiSHHnaqemJbr8GCHCv7or4Ro8Htess1g-ovw"
              >
                <LogIn className="w-3.5 h-3.5" />
                <span>{isLoggingIn ? 'กำลังเข้าสู่ระบบ...' : 'เข้าสู่ระบบ Google'}</span>
              </button>
            )}

            <button
              onClick={() => handleSyncGoogleSheet()}
              disabled={isSyncing}
              className="inline-flex items-center gap-2 px-3 py-1.5 text-xs font-medium rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition-all cursor-pointer disabled:opacity-50"
              title="ซิงค์ข้อมูลจาก Google Sheet ID: 1dKqK6soiSHHnaqemJbr8GCHCv7or4Ro8Htess1g-ovw"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${isSyncing ? 'animate-spin text-emerald-400' : ''}`} />
              <span>{isSyncing ? 'กำลังซิงค์...' : 'ซิงค์ชีท'}</span>
            </button>

            <button
              onClick={() => {
                setShowImportModal(true);
                setImportStatus('');
              }}
              className="inline-flex items-center gap-2 px-3 py-1.5 text-xs font-semibold rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white shadow-xs transition-all cursor-pointer"
              title="วางข้อมูลหรืออัปโหลดไฟล์จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ A ถึง AG"
            >
              <FileUp className="w-3.5 h-3.5" />
              <span>นำเข้าชีท</span>
            </button>

            <button
              onClick={() => setShowAlertModal(true)}
              className="relative p-2 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition-all cursor-pointer"
              title="แจ้งเตือนหม้อแปลงเข้าเกณฑ์"
            >
              <Bell className="w-4 h-4" />
              {stats.totalEligible > 0 && (
                <span className="absolute -top-1 -right-1 w-5 h-5 bg-rose-500 text-white text-[10px] font-bold rounded-full flex items-center justify-center animate-pulse">
                  {stats.totalEligible}
                </span>
              )}
            </button>

            <button
              onClick={() => exportTransformersToExcel(transformers)}
              className="inline-flex items-center gap-2 px-3.5 py-1.5 text-xs font-semibold rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white shadow-sm transition-all cursor-pointer"
            >
              <Download className="w-4 h-4" />
              <span>Excel</span>
            </button>
          </div>
        </div>
      </header>

      {/* Main Navigation Tabs */}
      <div className="bg-white border-b border-slate-200 shadow-xs">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <nav className="flex space-x-1 sm:space-x-8 overflow-x-auto py-3">
            {[
              { id: 'dashboard', label: 'แดชบอร์ดภาพรวม', icon: BarChart3, count: null },
              { id: 'target', label: 'หม้อแปลงเข้าเกณฑ์ปรับลด', icon: AlertTriangle, count: stats.totalEligible, badgeColor: 'bg-rose-100 text-rose-700' },
              { id: 'map', label: 'แผนที่หม้อแปลง', icon: Layers, count: null },
              { id: 'trmatching', label: 'TR Matching', icon: Shuffle, count: trMatches.length, badgeColor: 'bg-indigo-100 text-indigo-700' },
              { id: 'database', label: 'รายการหม้อแปลงทั้งหมด', icon: Layers, count: stats.total },
              { id: 'pea34', label: 'วิเคราะห์ราย 34 กฟฟ.', icon: Building2, count: 34 },
              { id: 'alerts', label: 'ศูนย์แจ้งเตือนอัตโนมัติ', icon: Bell, count: stats.totalEligible }
            ].map(tab => {
              const Icon = tab.icon;
              const isActive = activeTab === tab.id;
              return (
                <button
                  key={tab.id}
                  onClick={() => setActiveTab(tab.id as any)}
                  className={`flex items-center gap-2 py-2 px-3 text-sm font-medium rounded-lg transition-all whitespace-nowrap cursor-pointer ${
                    isActive
                      ? 'bg-emerald-50 text-emerald-700 border border-emerald-200 shadow-2xs'
                      : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
                  }`}
                >
                  <Icon className={`w-4 h-4 ${isActive ? 'text-emerald-600' : 'text-slate-400'}`} />
                  <span>{tab.label}</span>
                  {tab.count !== null && (
                    <span className={`px-2 py-0.5 text-xs font-bold rounded-full ${tab.badgeColor || 'bg-slate-100 text-slate-700'}`}>
                      {tab.count}
                    </span>
                  )}
                </button>
              );
            })}
          </nav>
        </div>
      </div>

      {/* Content Area */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 flex-1 w-full">
        {/* ================= TAB 1: DASHBOARD ================= */}
        {activeTab === 'dashboard' && (
          <div className="space-y-6">
            {/* KPI Cards */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5">
              <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-xs flex items-center justify-between">
                <div>
                  <p className="text-xs font-semibold uppercase tracking-wider text-slate-500">หม้อแปลงในระบบทั้งหมด</p>
                  <h3 className="text-2xl font-bold text-slate-900 mt-1 tabular-nums">{stats.total.toLocaleString()} <span className="text-xs font-normal text-slate-500">เครื่อง</span></h3>
                  <p className="text-xs text-emerald-600 mt-1 flex items-center gap-1 font-medium">
                    <Building2 className="w-3.5 h-3.5" /> ครอบคลุม 34 กฟฟ.
                  </p>
                </div>
                <div className="w-12 h-12 rounded-xl bg-slate-100 flex items-center justify-center text-slate-700">
                  <Layers className="w-6 h-6" />
                </div>
              </div>

              <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-xs flex items-center justify-between">
                <div>
                  <p className="text-xs font-semibold uppercase tracking-wider text-slate-500">หม้อแปลงเข้าเกณฑ์ลดขนาด</p>
                  <h3 className="text-2xl font-bold text-rose-600 mt-1 tabular-nums">{stats.totalEligible.toLocaleString()} <span className="text-xs font-normal text-slate-500">เครื่อง</span></h3>
                  <p className="text-xs text-slate-600 mt-1.5 font-medium flex flex-wrap items-center gap-x-2">
                    <span>100 kVA: <strong className="text-slate-900 font-mono">{stats.target100.toLocaleString()}</strong></span>
                    <span className="text-slate-300">|</span>
                    <span>160 kVA: <strong className="text-slate-900 font-mono">{stats.target160.toLocaleString()}</strong></span>
                    <span className="text-slate-300">|</span>
                    <span>250 kVA: <strong className="text-slate-900 font-mono">{stats.target250.toLocaleString()}</strong></span>
                  </p>
                  <p className="text-[11px] text-rose-600 mt-1">
                    โหลด &lt; 30% ติดต่อกัน 3 ปี
                  </p>
                </div>
                <div className="w-12 h-12 rounded-xl bg-rose-50 flex items-center justify-center text-rose-600">
                  <AlertTriangle className="w-6 h-6" />
                </div>
              </div>

              <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-xs flex items-center justify-between">
                <div>
                  <p className="text-xs font-semibold uppercase tracking-wider text-slate-500">TR Matching</p>
                  <h3 className="text-2xl font-bold text-indigo-600 mt-1 tabular-nums">
                    {trMatches.length.toLocaleString()} <span className="text-xs font-normal text-slate-500">คู่</span>
                  </h3>
                  <p className="text-xs text-slate-600 mt-1.5 font-medium flex flex-wrap items-center gap-x-2">
                    <span>100 kVA: <strong className="text-slate-900 font-mono">{matchStatsByRating.m100.toLocaleString()}</strong></span>
                    <span className="text-slate-300">|</span>
                    <span>160 kVA: <strong className="text-slate-900 font-mono">{matchStatsByRating.m160.toLocaleString()}</strong></span>
                    <span className="text-slate-300">|</span>
                    <span>250 kVA: <strong className="text-slate-900 font-mono">{matchStatsByRating.m250.toLocaleString()}</strong></span>
                  </p>
                  <p className="text-[11px] text-emerald-600 mt-1 font-medium">
                    จับคู่ภายใน กฟฟ. เดียวกัน (% โหลดใหม่ &lt; 50%)
                  </p>
                </div>
                <div className="w-12 h-12 rounded-xl bg-indigo-50 flex items-center justify-center text-indigo-600">
                  <Shuffle className="w-6 h-6" />
                </div>
              </div>

              <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-xs flex items-center justify-between">
                <div>
                  <p className="text-xs font-semibold uppercase tracking-wider text-slate-500">ประมาณการลดพลังงานสูญเสีย</p>
                  <h3 className="text-2xl font-bold text-emerald-600 mt-1 tabular-nums">
                    {(stats.totalEligible * 1240).toLocaleString()} <span className="text-xs font-normal text-slate-500">kWh/ปี</span>
                  </h3>
                  <p className="text-xs text-emerald-600 mt-1 font-medium">
                    ลด No-load loss ประหยัดงบลงทุน
                  </p>
                </div>
                <div className="w-12 h-12 rounded-xl bg-emerald-50 flex items-center justify-center text-emerald-600">
                  <TrendingDown className="w-6 h-6" />
                </div>
              </div>
            </div>

            {/* Quick Action & Rules Banner */}
            <div className="bg-gradient-to-r from-slate-900 to-indigo-950 rounded-2xl p-6 text-white shadow-lg flex flex-col md:flex-row items-center justify-between gap-6">
              <div className="space-y-2">
                <h2 className="text-xl font-bold">เงื่อนไขการลดขนาดหม้อแปลงระบบจำหน่ายไฟฟ้า</h2>
                <ul className="text-sm text-slate-300 space-y-1.5 list-disc list-inside">
                  <li><strong>การวิเคราะห์ย้อนกลับไป 2 ปี:</strong> ตรวจสอบประวัติโหลด 3 ปีติดต่อกัน (โหลด 2 ปีก่อน, โหลด 1 ปีก่อน, โหลดปีปัจจุบัน) ต้อง &lt; 30% ทั้งหมด</li>
                  <li>เมื่อลดขนาดแล้ว % โหลดใหม่ต้องน้อยกว่า 50%</li>
                  <li><strong>พิกัดหม้อแปลงเป้าหมาย:</strong> 100 kVA, 160 kVA, และ 250 kVA</li>
                  <li><strong>เงื่อนไขยกเว้นสถานที่ติดตั้ง:</strong> ไม่นำหม้อแปลงที่สถานที่ติดตั้งมีคำว่า <strong>"จัดสรร"</strong>, <strong>"หมู่บ้าน"</strong>, <strong>"โครงการ"</strong>, <strong>"สถานีอัดประจุ"</strong> หรือ <strong>"EV"</strong> เข้ามาคำนวณในระบบนี้</li>
                </ul>
              </div>
              <div className="flex flex-col sm:flex-row gap-3 shrink-0">
                <button
                  onClick={() => setActiveTab('target')}
                  className="px-5 py-2.5 bg-emerald-500 hover:bg-emerald-400 text-white font-semibold rounded-xl text-sm transition-all shadow cursor-pointer flex items-center justify-center gap-2"
                >
                  <span>ดูรายการหม้อแปลงเข้าเกณฑ์</span>
                  <ArrowRight className="w-4 h-4" />
                </button>
                <button
                  onClick={() => setShowAddModal(true)}
                  className="px-5 py-2.5 bg-white/10 hover:bg-white/20 text-white font-semibold rounded-xl text-sm transition-all border border-white/20 cursor-pointer flex items-center justify-center gap-2"
                >
                  <PlusCircle className="w-4 h-4" />
                  <span>เพิ่มหม้อแปลงทดสอบ</span>
                </button>
              </div>
            </div>

            {/* Grid Charts / Office Summary */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              {/* Left 2 Cols: Top Offices with Oversized Transformers */}
              <div className="lg:col-span-2 bg-white rounded-xl p-6 border border-slate-200 shadow-xs">
                <div className="flex items-center justify-between mb-4">
                  <div>
                    <h3 className="font-bold text-slate-900 text-base">กฟฟ. ที่พบหม้อแปลงเข้าเกณฑ์ปรับลดขนาดสูงสุด</h3>
                    <p className="text-xs text-slate-500">จัดอันดับตามจำนวนหม้อแปลงพิกัด 100, 160, 250 kVA ที่มีโหลด &lt; 30% 3 ปีต่อเนื่อง</p>
                  </div>
                  <button
                    onClick={() => setActiveTab('pea34')}
                    className="text-xs font-medium text-emerald-600 hover:text-emerald-700 hover:underline"
                  >
                    ดูครบทั้ง 34 กฟฟ. →
                  </button>
                </div>

                <div className="space-y-3">
                  {Object.entries(stats.officeMap)
                    .sort((a, b) => b[1] - a[1])
                    .slice(0, 10)
                    .map(([office, count], idx) => {
                      const maxCount = Math.max(...Object.values(stats.officeMap)) || 1;
                      const pct = Math.round((count / maxCount) * 100);
                      return (
                        <div key={office} className="space-y-1">
                          <div className="flex justify-between text-xs font-medium">
                            <span className="text-slate-700 flex items-center gap-2">
                              <span className="w-5 h-5 rounded-full bg-slate-100 text-slate-600 flex items-center justify-center font-bold text-[10px]">
                                {idx + 1}
                              </span>
                              {office}
                            </span>
                            <span className="font-bold text-slate-900 tabular-nums">{count.toLocaleString()} เครื่อง</span>
                          </div>
                          <div className="w-full bg-slate-100 rounded-full h-2 overflow-hidden">
                            <div
                              className="bg-emerald-600 h-2 rounded-full transition-all duration-500"
                              style={{ width: `${Math.max(pct, 5)}%` }}
                            ></div>
                          </div>
                        </div>
                      );
                    })}
                </div>
              </div>

              {/* Right Col: Rating Summary Breakdown */}
              <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-xs flex flex-col justify-between">
                <div>
                  <h3 className="font-bold text-slate-900 text-base mb-1">สัดส่วนพิกัดหม้อแปลงที่เข้าเกณฑ์</h3>
                  <p className="text-xs text-slate-500 mb-6">แยกตามขนาดพิกัดเดิมและการจับคู่ลดขนาด</p>

                  <div className="space-y-4">
                    <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200">
                      <span className="text-xs font-semibold text-slate-500 uppercase">พิกัด 100 kVA</span>
                      <h4 className="text-lg font-bold text-slate-900">{stats.target100} <span className="text-xs font-normal text-slate-500">เครื่อง</span></h4>
                      <p className="text-[11px] text-emerald-600 font-medium">แนะนำลดเป็น: 50 kVA</p>
                    </div>

                    <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200">
                      <span className="text-xs font-semibold text-slate-500 uppercase">พิกัด 160 kVA</span>
                      <h4 className="text-lg font-bold text-slate-900">{stats.target160} <span className="text-xs font-normal text-slate-500">เครื่อง</span></h4>
                      <p className="text-[11px] text-emerald-600 font-medium">แนะนำลดเป็น: 100 หรือ 50 kVA</p>
                    </div>

                    <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200">
                      <span className="text-xs font-semibold text-slate-500 uppercase">พิกัด 250 kVA</span>
                      <h4 className="text-lg font-bold text-slate-900">{stats.target250} <span className="text-xs font-normal text-slate-500">เครื่อง</span></h4>
                      <p className="text-[11px] text-emerald-600 font-medium">แนะนำลดเป็น: 160, 100 หรือ 50 kVA</p>
                    </div>
                  </div>
                </div>

                <button
                  onClick={() => setActiveTab('target')}
                  className="w-full mt-6 py-2.5 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-medium transition-all cursor-pointer"
                >
                  ตรวจสอบรายการแนะนำทั้งหมด
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ================= TAB 2: TARGET RIGHTSIZING ================= */}
        {activeTab === 'target' && (
          <div className="space-y-6">
            <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-xs flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
              <div>
                <h2 className="text-lg font-bold text-slate-900">หม้อแปลงเข้าเกณฑ์ลดขนาด (Rightsizing Targets)</h2>
                <p className="text-xs text-slate-500">
                  เงื่อนไข: % โหลด &lt; 30% ติดต่อกัน 3 ปี, เมื่อลดขนาดแล้ว % โหลดใหม่ต้องน้อยกว่า 50%, และ<strong>ไม่เอาหม้อแปลงที่สถานที่ติดตั้งมีคำว่า "จัดสรร", "หมู่บ้าน", "โครงการ", "สถานีอัดประจุ" หรือ "EV"</strong> เข้ามาคำนวณในระบบนี้
                </p>
                <div className="flex flex-wrap items-center gap-2.5 mt-3">
                  <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-slate-100 text-slate-800 border border-slate-200 shadow-2xs">
                    <span className="w-2 h-2 rounded-full bg-slate-500"></span>
                    100 kVA: <span className="font-mono text-slate-900">{stats.target100.toLocaleString()}</span> เครื่อง
                  </span>
                  <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-indigo-50 text-indigo-700 border border-indigo-200 shadow-2xs">
                    <span className="w-2 h-2 rounded-full bg-indigo-500"></span>
                    160 kVA: <span className="font-mono text-indigo-900">{stats.target160.toLocaleString()}</span> เครื่อง
                  </span>
                  <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-purple-50 text-purple-700 border border-purple-200 shadow-2xs">
                    <span className="w-2 h-2 rounded-full bg-purple-500"></span>
                    250 kVA: <span className="font-mono text-purple-900">{stats.target250.toLocaleString()}</span> เครื่อง
                  </span>
                </div>
              </div>

              {/* Filters Header for Rightsizing Targets */}
            </div>
            
            <div className="bg-white rounded-xl p-4 border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-3">
              <div className="flex flex-wrap items-center gap-3 w-full sm:w-auto flex-1">
                <div className="relative flex-1 sm:w-64 max-w-xs">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-2.5" />
                  <input
                    type="text"
                    placeholder="ค้นหา PEA หม้อแปลง, กฟฟ., สถานที่..."
                    value={targetSearchQuery}
                    onChange={(e) => setTargetSearchQuery(e.target.value)}
                    className="w-full pl-9 pr-4 py-1.5 text-xs border border-slate-200 rounded-lg focus:outline-none focus:border-emerald-500"
                  />
                </div>

                <select
                  value={targetOffice}
                  onChange={(e) => setTargetOffice(e.target.value)}
                  className="py-1.5 px-3 text-xs border border-slate-200 rounded-lg bg-white text-slate-700 focus:outline-none focus:border-emerald-500"
                >
                  <option value="all">ทุก กฟฟ. (34 กฟฟ.)</option>
                  {peaOffices.map(off => (
                    <option key={off} value={off}>{off}</option>
                  ))}
                </select>

                <select
                  value={targetRating}
                  onChange={(e) => setTargetRating(e.target.value)}
                  className="py-1.5 px-3 text-xs border border-slate-200 rounded-lg bg-white text-slate-700 focus:outline-none focus:border-emerald-500"
                >
                  <option value="all">ทุกพิกัด kVA</option>
                  <option value="100">100 kVA</option>
                  <option value="160">160 kVA</option>
                  <option value="250">250 kVA</option>
                </select>

                {(targetSearchQuery || targetOffice !== 'all' || targetRating !== 'all') && (
                  <button
                    onClick={() => {
                      setTargetSearchQuery('');
                      setTargetOffice('all');
                      setTargetRating('all');
                    }}
                    className="text-xs text-rose-600 hover:text-rose-700 font-medium cursor-pointer"
                  >
                    ล้างตัวกรอง
                  </button>
                )}
              </div>

              <div className="flex items-center gap-3">
                <button
                  onClick={() => exportTransformersToExcel(filteredEligibleList.map(i => i.tr))}
                  className="inline-flex items-center gap-2 px-4 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded-lg shadow-sm transition-all cursor-pointer"
                >
                  <Download className="w-4 h-4" />
                  <span>ส่งออกข้อมูล (Excel)</span>
                </button>
                <div className="text-xs text-slate-500 font-medium">
                  พบ <strong className="text-emerald-700 font-mono">{filteredEligibleList.length.toLocaleString()}</strong> จากทั้งหมด {stats.eligibleList.length.toLocaleString()} เครื่อง
                </div>
              </div>
            </div>


            <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-100 text-slate-700 text-xs font-bold tracking-wider border-b border-slate-200">
                      <th className="py-2.5 px-2 text-center w-10">ที่</th>
                      <th className="py-2.5 px-2 text-left whitespace-nowrap leading-tight">PEA<br />หม้อแปลง</th>
                      <th className="py-2.5 px-2 text-left whitespace-nowrap">กฟฟ.</th>
                      <th className="py-2.5 px-2 text-left whitespace-nowrap">สถานที่ติดตั้ง</th>
                      <th className="py-2.5 px-2 text-center whitespace-nowrap leading-tight">พิกัดเดิม<br />(<span className="lowercase" style={{ textTransform: 'lowercase' }}>k</span>VA)</th>
                      <th className="py-2.5 px-2 text-center whitespace-nowrap leading-tight">โหลด 2 ปีก่อน<br />(%)</th>
                      <th className="py-2.5 px-2 text-center whitespace-nowrap leading-tight">โหลด 1 ปีก่อน<br />(%)</th>
                      <th className="py-2.5 px-2 text-center bg-indigo-50 text-indigo-900 whitespace-nowrap leading-tight">โหลดปีปัจจุบัน<br />(%)</th>
                      <th className="py-2.5 px-2 text-center whitespace-nowrap leading-tight">พิกัดแนะนำ<br />(<span className="lowercase" style={{ textTransform: 'lowercase' }}>k</span>VA)</th>
                      <th className="py-2.5 px-2 text-center whitespace-nowrap leading-tight">% โหลดใหม่<br />(&lt;50%)</th>
                      <th className="py-2.5 px-2 text-center whitespace-nowrap leading-tight">สถานะ<br />แผนงาน</th>
                      <th className="py-2.5 px-2 text-center whitespace-nowrap">จัดการ</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 text-sm">
                    {paginatedEligibleList.length === 0 ? (
                      <tr>
                        <td colSpan={12} className="py-8 text-center text-slate-500 text-sm">
                          ไม่พบหม้อแปลงที่เข้าเกณฑ์
                        </td>
                      </tr>
                    ) : (
                      paginatedEligibleList.map(({ tr, evalRes }, idx) => (
                        <tr key={tr.id} className="hover:bg-slate-50 transition-colors">
                          <td className="py-2.5 px-2 text-center font-mono text-xs text-slate-500">
                            {(targetPage - 1) * TARGET_ITEMS_PER_PAGE + idx + 1}
                          </td>
                          <td className="py-2.5 px-2 font-bold text-slate-900 font-mono whitespace-nowrap">{tr.trCode}</td>
                          <td className="py-2.5 px-2 text-slate-800 font-medium whitespace-nowrap">{tr.feaOffice}</td>
                          <td className="py-2.5 px-2 text-slate-700 text-xs">{tr.location || '-'}</td>
                          <td className="py-2.5 px-2 text-center font-bold text-indigo-600 font-mono">{tr.ratedKva}</td>
                          <td className="py-2.5 px-2 text-center font-mono text-xs font-medium text-slate-700">{format2Decimals(tr.year1LoadPct)}</td>
                          <td className="py-2.5 px-2 text-center font-mono text-xs font-medium text-slate-700">{format2Decimals(tr.year2LoadPct)}</td>
                          <td className="py-2.5 px-2 text-center font-mono text-xs font-bold text-indigo-700 bg-indigo-50/50">{format2Decimals(tr.year3LoadPct)}</td>
                          <td className="py-2.5 px-2 text-center">
                            <span className="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-bold bg-emerald-100 text-emerald-800 font-mono whitespace-nowrap">
                              {evalRes.recommendedKva}
                            </span>
                          </td>
                          <td className="py-2.5 px-2 text-center font-mono font-bold text-slate-800">
                            {format2Decimals(evalRes.newLoadPct)}
                          </td>
                          <td className="py-2.5 px-2 text-center">
                            <span className={`inline-flex items-center px-2 py-0.5 rounded text-xs font-medium whitespace-nowrap ${
                              tr.status === 'Action Planned' 
                                ? 'bg-indigo-100 text-indigo-800' 
                                : 'bg-amber-100 text-amber-800'
                            }`}>
                              {tr.status === 'Action Planned' ? 'อนุมัติแผนแล้ว' : 'รอพิจารณา'}
                            </span>
                          </td>
                          <td className="py-2.5 px-2 text-center space-x-1 whitespace-nowrap">
                            <button
                              onClick={() => setSelectedTransformer(tr)}
                              className="px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-medium rounded transition-colors cursor-pointer"
                            >
                              รายละเอียด
                            </button>
                            {tr.status !== 'Action Planned' && (
                              <button
                                onClick={() => handleApprovePlan(tr.id)}
                                className="px-2 py-1 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-medium rounded transition-colors cursor-pointer"
                              >
                                อนุมัติแผน
                              </button>
                            )}
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Pagination Controls */}
            {filteredEligibleList.length > 0 && (
              <div className="bg-white rounded-xl px-5 py-3 border border-slate-200 shadow-xs flex items-center justify-between text-xs">
                <span className="text-slate-500">
                  แสดงรายการที่ {((targetPage - 1) * TARGET_ITEMS_PER_PAGE) + 1} ถึง {Math.min(targetPage * TARGET_ITEMS_PER_PAGE, filteredEligibleList.length)} จากทั้งหมด {filteredEligibleList.length.toLocaleString()} รายการ (แสดงหน้าละ {TARGET_ITEMS_PER_PAGE} รายการ)
                </span>
                <div className="flex items-center gap-1.5">
                  <button
                    onClick={() => setTargetPage(prev => Math.max(prev - 1, 1))}
                    disabled={targetPage === 1}
                    className="px-3 py-1.5 rounded-lg border border-slate-200 bg-white text-slate-700 hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed font-medium cursor-pointer"
                  >
                    ก่อนหน้า
                  </button>
                  <span className="px-3 py-1.5 font-medium text-slate-700">
                    หน้า {targetPage} จาก {totalTargetPages}
                  </span>
                  <button
                    onClick={() => setTargetPage(prev => Math.min(prev + 1, totalTargetPages))}
                    disabled={targetPage === totalTargetPages}
                    className="px-3 py-1.5 rounded-lg border border-slate-200 bg-white text-slate-700 hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed font-medium cursor-pointer"
                  >
                    ถัดไป
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ================= TAB 3: TRANSFORMER DATABASE ================= */}
        {activeTab === 'database' && (
          <div className="space-y-6">

            {/* Filters Header */}
            <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-xs flex flex-col md:flex-row items-center justify-between gap-4">
              <div className="flex flex-wrap items-center gap-3 w-full md:w-auto">
                <div className="relative flex-1 sm:w-64">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-3" />
                  <input
                    type="text"
                    placeholder="ค้นหา PEA หม้อแปลง, กฟฟ., ฟีดเดอร์..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="w-full pl-9 pr-4 py-2 text-xs border border-slate-200 rounded-lg focus:outline-none focus:border-emerald-500"
                  />
                </div>

                <select
                  value={selectedOffice}
                  onChange={(e) => setSelectedOffice(e.target.value)}
                  className="py-2 px-3 text-xs border border-slate-200 rounded-lg bg-white text-slate-700 focus:outline-none focus:border-emerald-500"
                >
                  <option value="all">ทุก กฟฟ. (34 กฟฟ.)</option>
                  {peaOffices.map(off => (
                    <option key={off} value={off}>{off}</option>
                  ))}
                </select>

                <select
                  value={selectedRating}
                  onChange={(e) => setSelectedRating(e.target.value)}
                  className="py-2 px-3 text-xs border border-slate-200 rounded-lg bg-white text-slate-700 focus:outline-none focus:border-emerald-500"
                >
                  <option value="all">ทุกพิกัด kVA</option>
                  <option value="50">50 kVA</option>
                  <option value="100">100 kVA</option>
                  <option value="160">160 kVA</option>
                  <option value="250">250 kVA</option>
                  <option value="315">315 kVA</option>
                  <option value="400">400 kVA</option>
                  <option value="500">500 kVA</option>
                </select>

                <label className="flex items-center gap-2 text-xs font-medium text-slate-700 cursor-pointer select-none">
                  <input
                    type="checkbox"
                    checked={filterEligibleOnly}
                    onChange={(e) => setFilterEligibleOnly(e.target.checked)}
                    className="rounded border-slate-300 text-emerald-600 focus:ring-emerald-500 w-4 h-4"
                  />
                  <span>แสดงเฉพาะเข้าเกณฑ์ Rightsizing</span>
                </label>
              </div>

              <div className="flex items-center gap-2 shrink-0">
                <button
                  onClick={() => setShowAddModal(true)}
                  className="inline-flex items-center gap-2 px-4 py-2 bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold rounded-lg shadow-sm transition-all cursor-pointer"
                >
                  <PlusCircle className="w-4 h-4" />
                  <span>เพิ่มหม้อแปลง</span>
                </button>
                <button
                  onClick={() => exportTransformersToExcel(filteredTransformers)}
                  className="inline-flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded-lg shadow-sm transition-all cursor-pointer"
                >
                  <Download className="w-4 h-4" />
                  <span>ส่งออกข้อมูล (Excel)</span>
                </button>
              </div>
            </div>

            {/* Table */}
            <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
              <div className="px-5 py-3 border-b border-slate-200 bg-slate-50 flex items-center justify-between text-xs text-slate-500">
                <span>แสดงผล {filteredTransformers.length > 0 ? ((dbPage - 1) * DB_ITEMS_PER_PAGE) + 1 : 0} ถึง {Math.min(dbPage * DB_ITEMS_PER_PAGE, filteredTransformers.length)} จากทั้งหมด {filteredTransformers.length.toLocaleString()} รายการ (จากทั้งหมด {transformers.length.toLocaleString()} รายการ) (แสดงหน้าละ {DB_ITEMS_PER_PAGE} รายการ)</span>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-100 text-slate-700 text-xs font-bold tracking-wider border-b border-slate-200">
                      <th className="py-2.5 px-2 text-center w-10">ที่</th>
                      <th className="py-2.5 px-2 text-left whitespace-nowrap leading-tight">PEA<br />หม้อแปลง</th>
                      <th className="py-2.5 px-2 text-left whitespace-nowrap">กฟฟ.</th>
                      <th className="py-2.5 px-2 text-left whitespace-nowrap">สถานที่ติดตั้ง</th>
                      <th className="py-2.5 px-2 text-left whitespace-nowrap">ฟีดเดอร์</th>
                      <th className="py-2.5 px-2 text-center whitespace-nowrap leading-tight">พิกัดเดิม<br />(<span className="lowercase" style={{ textTransform: 'lowercase' }}>k</span>VA)</th>
                      <th className="py-2.5 px-2 text-center whitespace-nowrap leading-tight">โหลด 2 ปีก่อน<br />(%)</th>
                      <th className="py-2.5 px-2 text-center whitespace-nowrap leading-tight">โหลด 1 ปีก่อน<br />(%)</th>
                      <th className="py-2.5 px-2 text-center bg-indigo-50 text-indigo-900 whitespace-nowrap leading-tight">โหลดปีปัจจุบัน<br />(%)</th>
                      <th className="py-2.5 px-2 text-center whitespace-nowrap leading-tight">สถานะ<br />Rightsizing</th>
                      <th className="py-2.5 px-2 text-center whitespace-nowrap">จัดการ</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 text-sm">
                    {paginatedFilteredTransformers.length === 0 ? (
                      <tr>
                        <td colSpan={11} className="py-8 text-center text-slate-500 text-sm">
                          ไม่พบข้อมูลหม้อแปลงตามเงื่อนไขที่ค้นหา
                        </td>
                      </tr>
                    ) : (
                      paginatedFilteredTransformers.map((tr, idx) => {
                        const evalRes = evaluateTransformerRightsizing(tr);
                        return (
                          <tr key={tr.id} className="hover:bg-slate-50 transition-colors">
                            <td className="py-2.5 px-2 text-center font-mono text-xs text-slate-500">
                              {(dbPage - 1) * DB_ITEMS_PER_PAGE + idx + 1}
                            </td>
                            <td className="py-2.5 px-2 font-bold text-slate-900 font-mono whitespace-nowrap">{tr.trCode}</td>
                            <td className="py-2.5 px-2 text-slate-800 font-medium whitespace-nowrap">{tr.feaOffice}</td>
                            <td className="py-2.5 px-2 text-slate-700 text-xs">{tr.location || '-'}</td>
                            <td className="py-2.5 px-2 text-slate-600 text-xs whitespace-nowrap">{tr.feeder}</td>
                            <td className="py-2.5 px-2 text-center font-bold text-slate-900 font-mono">{tr.ratedKva}</td>
                            <td className={`py-2.5 px-2 text-center font-mono text-xs ${tr.year1LoadPct < 30 ? 'text-emerald-600 font-semibold' : 'text-slate-600'}`}>{format2Decimals(tr.year1LoadPct)}</td>
                            <td className={`py-2.5 px-2 text-center font-mono text-xs ${tr.year2LoadPct < 30 ? 'text-emerald-600 font-semibold' : 'text-slate-600'}`}>{format2Decimals(tr.year2LoadPct)}</td>
                            <td className={`py-2.5 px-2 text-center font-mono text-xs bg-indigo-50/50 ${tr.year3LoadPct < 30 ? 'text-indigo-700 font-bold' : 'text-slate-700 font-semibold'}`}>{format2Decimals(tr.year3LoadPct)}</td>
                            <td className="py-2.5 px-2 text-center">
                              {evalRes.isEligible ? (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-bold bg-emerald-100 text-emerald-800 whitespace-nowrap">
                                  <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                                  ลดเป็น {evalRes.recommendedKva} kVA
                                </span>
                              ) : (
                                <span className="text-xs text-slate-400">ปกติ / ไม่เข้าเกณฑ์</span>
                              )}
                            </td>
                            <td className="py-2.5 px-2 text-center whitespace-nowrap">
                              <button
                                onClick={() => setSelectedTransformer(tr)}
                                className="px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-medium rounded transition-colors cursor-pointer"
                              >
                                รายละเอียด
                              </button>
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Pagination Controls for Database Table */}
            {filteredTransformers.length > 0 && (
              <div className="bg-white rounded-xl px-5 py-3 border border-slate-200 shadow-xs flex items-center justify-between text-xs">
                <span className="text-slate-500">
                  แสดงรายการที่ {((dbPage - 1) * DB_ITEMS_PER_PAGE) + 1} ถึง {Math.min(dbPage * DB_ITEMS_PER_PAGE, filteredTransformers.length)} จากทั้งหมด {filteredTransformers.length.toLocaleString()} รายการ (แสดงหน้าละ {DB_ITEMS_PER_PAGE} รายการ)
                </span>
                <div className="flex items-center gap-1.5">
                  <button
                    onClick={() => setDbPage(prev => Math.max(prev - 1, 1))}
                    disabled={dbPage === 1}
                    className="px-3 py-1.5 rounded-lg border border-slate-200 bg-white text-slate-700 hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed font-medium cursor-pointer"
                  >
                    ก่อนหน้า
                  </button>
                  <span className="px-3 py-1.5 font-medium text-slate-700">
                    หน้า {dbPage} จาก {totalDbPages}
                  </span>
                  <button
                    onClick={() => setDbPage(prev => Math.min(prev + 1, totalDbPages))}
                    disabled={dbPage === totalDbPages}
                    className="px-3 py-1.5 rounded-lg border border-slate-200 bg-white text-slate-700 hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed font-medium cursor-pointer"
                  >
                    ถัดไป
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ================= TAB 4: 34 PEA OFFICES ================= */}
        {activeTab === 'pea34' && (
          <div className="space-y-6">
            <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-xs">
              <h2 className="text-lg font-bold text-slate-900">สรุปผลการวิเคราะห์จำแนกราย 34 กฟฟ. (การไฟฟ้าส่วนภูมิภาค)</h2>
              <p className="text-xs text-slate-500">
                ภาพรวมสถิติจำนวนหม้อแปลง 100, 160, 250 kVA ที่เข้าเกณฑ์ปรับลดขนาดในแต่ละสังกัด กฟฟ.
              </p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
              {peaOffices.map((office) => {
                const officeTrs = transformers.filter(t => t.feaOffice === office);
                const eligibleOfficeTrs = officeTrs.filter(t => evaluateTransformerRightsizing(t).isEligible);
                const pct = officeTrs.length > 0 ? Math.round((eligibleOfficeTrs.length / officeTrs.length) * 100) : 0;

                return (
                  <div key={office} className="bg-white rounded-xl p-5 border border-slate-200 shadow-xs flex flex-col justify-between">
                    <div>
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-xs font-semibold uppercase text-slate-400">กฟฟ. สังกัด</span>
                        <span className={`px-2 py-0.5 rounded text-xs font-bold ${eligibleOfficeTrs.length > 0 ? 'bg-rose-100 text-rose-700' : 'bg-slate-100 text-slate-600'}`}>
                          เป้าหมาย: {eligibleOfficeTrs.length.toLocaleString()} เครื่อง
                        </span>
                      </div>
                      <h3 className="font-bold text-slate-900 text-base mb-3">{office}</h3>
                      <div className="space-y-2 text-xs text-slate-600">
                        <div className="flex justify-between">
                          <span>หม้อแปลงทั้งหมดในสังกัด:</span>
                          <span className="font-bold text-slate-900">{officeTrs.length.toLocaleString()} เครื่อง</span>
                        </div>
                        <div className="flex justify-between">
                          <span>เข้าเกณฑ์ Rightsizing:</span>
                          <span className="font-bold text-emerald-600">{eligibleOfficeTrs.length.toLocaleString()} เครื่อง</span>
                        </div>
                      </div>
                    </div>

                    <div className="mt-5 pt-4 border-t border-slate-100 flex items-center justify-between">
                      <button
                        onClick={() => {
                          setMapOffice(office);
                          setMapRating('all');
                          setMapSearchQuery('');

                          const officeEligible = stats.eligibleList.filter(({ tr }) => tr.feaOffice === office && (tr as any).latitude && (tr as any).longitude);
                          if (officeEligible.length > 0) {
                            let sumLat = 0;
                            let sumLng = 0;
                            officeEligible.forEach(({ tr }) => {
                              sumLat += Number((tr as any).latitude);
                              sumLng += Number((tr as any).longitude);
                            });
                            setMapCenter({
                              lat: sumLat / officeEligible.length,
                              lng: sumLng / officeEligible.length
                            });
                            setMapZoom(officeEligible.length === 1 ? 15 : 12);
                          } else {
                            setMapCenter({ lat: 6.845, lng: 101.233 });
                            setMapZoom(9);
                          }

                          setActiveTab('map');
                        }}
                        className="text-xs font-semibold text-emerald-600 hover:text-emerald-700 hover:underline cursor-pointer flex items-center gap-1"
                      >
                        <span>ดูรายการใน กฟฟ. นี้ (แผนที่)</span>
                        <ArrowRight className="w-3.5 h-3.5" />
                      </button>
                      <span className="text-[11px] font-mono text-slate-400">สัดส่วน {pct}%</span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* ================= TAB 5: ALERTS & NOTIFICATIONS ================= */}
        {activeTab === 'alerts' && (
          <div className="space-y-6">
            <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-xs flex items-center justify-between">
              <div>
                <h2 className="text-lg font-bold text-slate-900">ศูนย์แจ้งเตือนอัตโนมัติ (Automated Notification Center)</h2>
                <p className="text-xs text-slate-500">ระบบตรวจสอบและแจ้งเตือนหม้อแปลงที่เข้าเกณฑ์ปรับลดขนาดอัตโนมัติจากฐานข้อมูล</p>
              </div>
              <div className="flex items-center gap-3">
                <span className="px-3 py-1 bg-rose-100 text-rose-800 text-xs font-bold rounded-full">
                  พบหม้อแปลงต้องปรับลด {stats.totalEligible.toLocaleString()} รายการ
                </span>
              </div>
            </div>

            <div className="space-y-4">
              {paginatedAlertList.length === 0 ? (
                <div className="bg-white rounded-xl p-8 text-center text-slate-500 text-sm border border-slate-200">
                  ไม่พบรายการแจ้งเตือน
                </div>
              ) : (
                paginatedAlertList.map(({ tr, evalRes }, idx) => (
                  <div key={tr.id} className="bg-white rounded-xl p-5 border border-slate-200 shadow-xs flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
                    <div className="flex items-start gap-4">
                      <div className="w-10 h-10 rounded-xl bg-rose-50 text-rose-600 flex items-center justify-center shrink-0">
                        <Bell className="w-5 h-5 animate-bounce" />
                      </div>
                      <div>
                        <div className="flex items-center gap-2">
                          <h4 className="font-bold text-slate-900 font-mono text-base">{tr.trCode}</h4>
                          <span className="text-xs bg-indigo-50 text-indigo-700 px-2 py-0.5 rounded font-medium border border-indigo-200">
                            {tr.feaOffice}
                          </span>
                        </div>
                        <p className="text-xs text-slate-600 mt-1">
                          พิกัดเดิม: <strong className="text-slate-900">{tr.ratedKva} kVA</strong> | โหลด 3 ปีย้อนหลัง: <span className="font-mono">{tr.year1LoadPct}%, {tr.year2LoadPct}%, {tr.year3LoadPct}%</span>
                        </p>
                        <p className="text-xs text-emerald-700 font-medium mt-1">
                          👉 คำแนะนำ: ลดขนาดเป็น <strong className="underline">{evalRes.recommendedKva} kVA</strong> (% โหลดใหม่ {evalRes.newLoadPct}%)
                        </p>
                      </div>
                    </div>

                    <div className="flex items-center gap-2 self-end md:self-center">
                      <button
                        onClick={() => setSelectedTransformer(tr)}
                        className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-medium rounded-lg transition-colors cursor-pointer"
                      >
                        ดูรายละเอียด
                      </button>
                      {tr.status !== 'Action Planned' ? (
                        <button
                          onClick={() => handleApprovePlan(tr.id)}
                          className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded-lg shadow-sm transition-all cursor-pointer"
                        >
                          อนุมัติแผนปรับลด
                        </button>
                      ) : (
                        <span className="px-3 py-1.5 bg-indigo-50 text-indigo-700 text-xs font-semibold rounded-lg border border-indigo-200">
                          อนุมัติแล้ว
                        </span>
                      )}
                    </div>
                  </div>
                ))
              )}
            </div>

            {/* Pagination Controls for Alerts */}
            {stats.eligibleList.length > 0 && (
              <div className="bg-white rounded-xl px-5 py-3 border border-slate-200 shadow-xs flex items-center justify-between text-xs">
                <span className="text-slate-500">
                  แสดงรายการที่ {((alertPage - 1) * ALERT_ITEMS_PER_PAGE) + 1} ถึง {Math.min(alertPage * ALERT_ITEMS_PER_PAGE, stats.eligibleList.length)} จากทั้งหมด {stats.eligibleList.length.toLocaleString()} รายการ (แสดงหน้าละ {ALERT_ITEMS_PER_PAGE} รายการ)
                </span>
                <div className="flex items-center gap-1.5">
                  <button
                    onClick={() => setAlertPage(prev => Math.max(prev - 1, 1))}
                    disabled={alertPage === 1}
                    className="px-3 py-1.5 rounded-lg border border-slate-200 bg-white text-slate-700 hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed font-medium cursor-pointer"
                  >
                    ก่อนหน้า
                  </button>
                  <span className="px-3 py-1.5 font-medium text-slate-700">
                    หน้า {alertPage} จาก {totalAlertPages}
                  </span>
                  <button
                    onClick={() => setAlertPage(prev => Math.min(prev + 1, totalAlertPages))}
                    disabled={alertPage === totalAlertPages}
                    className="px-3 py-1.5 rounded-lg border border-slate-200 bg-white text-slate-700 hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed font-medium cursor-pointer"
                  >
                    ถัดไป
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ================= TAB 6: TR MATCHING ================= */}
        {activeTab === 'trmatching' && (
          <div className="space-y-6">
            <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-xs flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
              <div>
                <h2 className="text-lg font-bold text-slate-900">การจับคู่หม้อแปลงเข้าเกณฑ์ลดขนาดกับหม้อแปลงโหลดเกิน (&gt;80%)</h2>
                <p className="text-xs text-slate-500">
                  นำหม้อแปลงที่เข้าเกณฑ์ปรับลดขนาด (Rightsizing) ไปจับคู่กับหม้อแปลงที่มีภาระโหลดปีปัจจุบันเกิน 80% ภายในสังกัด กฟฟ. เดียวกัน โดยคัดเลือกหม้อแปลงปลายทางที่อยู่ใกล้ที่สุดเป็นอันดับแรก และเมื่อจับคู่กันแล้ว % โหลดใหม่ทั้ง 2 เครื่องต้องน้อยกว่า 50%
                </p>
              </div>
            </div>

            <div className="bg-white rounded-xl p-4 border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-3">
              <div className="flex flex-wrap items-center gap-3 w-full sm:w-auto flex-1">
                <div className="relative flex-1 sm:w-64 max-w-xs">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-2.5" />
                  <input
                    type="text"
                    placeholder="ค้นหา PEA หม้อแปลง (ต้นทาง/ปลายทาง), กฟฟ...."
                    value={matchSearchQuery}
                    onChange={(e) => setMatchSearchQuery(e.target.value)}
                    className="w-full pl-9 pr-4 py-1.5 text-xs border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-500"
                  />
                </div>

                <select
                  value={matchOffice}
                  onChange={(e) => setMatchOffice(e.target.value)}
                  className="py-1.5 px-3 text-xs border border-slate-200 rounded-lg bg-white text-slate-700 focus:outline-none focus:border-indigo-500"
                >
                  <option value="all">ทุก กฟฟ. (34 กฟฟ.)</option>
                  {peaOffices.map(off => (
                    <option key={off} value={off}>{off}</option>
                  ))}
                </select>

                <select
                  value={matchRating}
                  onChange={(e) => setMatchRating(e.target.value)}
                  className="py-1.5 px-3 text-xs border border-slate-200 rounded-lg bg-white text-slate-700 focus:outline-none focus:border-indigo-500"
                >
                  <option value="all">ทุกพิกัด kVA</option>
                  <option value="50">พิกัด 50 kVA</option>
                  <option value="100">พิกัด 100 kVA</option>
                  <option value="160">พิกัด 160 kVA</option>
                  <option value="250">พิกัด 250 kVA</option>
                </select>

                {(matchSearchQuery || matchOffice !== 'all' || matchRating !== 'all') && (
                  <button
                    onClick={() => {
                      setMatchSearchQuery('');
                      setMatchOffice('all');
                      setMatchRating('all');
                    }}
                    className="text-xs text-rose-600 hover:text-rose-700 font-medium cursor-pointer"
                  >
                    ล้างตัวกรอง
                  </button>
                )}
                <button
                  onClick={() => exportTRMatchingToExcel(filteredMatches)}
                  className="inline-flex items-center gap-2 px-4 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded-lg shadow-sm transition-all cursor-pointer"
                >
                  <Download className="w-4 h-4" />
                  <span>ส่งออกข้อมูล (Excel)</span>
                </button>
              </div>

              <div className="text-xs text-slate-500 font-medium">
                พบ <strong className="text-indigo-700 font-mono">{filteredMatches.length.toLocaleString()}</strong> จากทั้งหมด {trMatches.length.toLocaleString()} คู่
              </div>
            </div>

            <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-100 text-slate-700 text-xs font-bold tracking-wider border-b border-slate-200">
                      <th className="py-2.5 px-3 text-center w-12">ที่</th>
                      <th className="py-2.5 px-3 text-left whitespace-nowrap">สังกัด กฟฟ.</th>
                      <th className="py-2.5 px-3 text-left whitespace-nowrap leading-tight bg-emerald-50/70 text-emerald-900">
                        หม้อแปลงต้นทาง<br />(ลดขนาด / Rightsizing)
                      </th>
                      <th className="py-2.5 px-3 text-center whitespace-nowrap bg-emerald-50/70 text-emerald-900">พิกัดเดิม / แนะนำ (kVA)</th>
                      <th className="py-2.5 px-3 text-center whitespace-nowrap bg-emerald-50/70 text-emerald-900">โหลดปัจจุบัน (%)</th>
                      <th className="py-2.5 px-3 text-left whitespace-nowrap leading-tight bg-rose-50/70 text-rose-900">
                        หม้อแปลงปลายทาง<br />(โหลดเกิน &gt;80%)
                      </th>
                      <th className="py-2.5 px-3 text-center whitespace-nowrap bg-rose-50/70 text-rose-900">พิกัดเดิม (kVA)</th>
                      <th className="py-2.5 px-3 text-center whitespace-nowrap bg-rose-50/70 text-rose-900">โหลดปัจจุบัน (%)</th>
                      <th className="py-2.5 px-3 text-center whitespace-nowrap bg-indigo-50/70 text-indigo-900">ระยะห่าง (km)</th>
                      <th className="py-2.5 px-3 text-center whitespace-nowrap">จัดการ</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 text-sm">
                    {paginatedMatches.length === 0 ? (
                      <tr>
                        <td colSpan={10} className="py-8 text-center text-slate-500 text-sm">
                          ไม่พบการจับคู่หม้อแปลงตามเงื่อนไขในขณะนี้
                        </td>
                      </tr>
                    ) : (
                      paginatedMatches.map(({ office, sourceTr, sourceEval, targetTr, targetNewLoadPct, distanceKm }, idx) => (
                        <tr key={`${sourceTr.id}_${targetTr.id}`} className="hover:bg-slate-50 transition-colors">
                          <td className="py-3 px-3 text-center font-mono text-xs text-slate-500">
                            {(matchPage - 1) * MATCH_ITEMS_PER_PAGE + idx + 1}
                          </td>
                          <td className="py-3 px-3 font-semibold text-slate-800 text-xs whitespace-nowrap">{office}</td>
                          <td className="py-3 px-3 font-mono font-bold text-emerald-700 text-xs whitespace-nowrap bg-emerald-50/30">
                            {sourceTr.trCode}
                          </td>
                          <td className="py-3 px-3 text-center font-mono text-xs whitespace-nowrap bg-emerald-50/30">
                            <strong className="text-emerald-600">{sourceTr.ratedKva}</strong> → <strong className="text-rose-600">{sourceEval.recommendedKva}</strong>
                          </td>
                          <td className="py-3 px-3 text-center font-mono text-xs font-bold text-emerald-800 bg-emerald-50/30">
                            <div>{format2Decimals(sourceTr.year3LoadPct)}</div>
                            <div className="text-[10px] text-rose-600 font-sans font-medium">โหลดใหม่ {format2Decimals(sourceEval.newLoadPct)}%</div>
                          </td>
                          <td className="py-3 px-3 font-mono font-bold text-rose-700 text-xs whitespace-nowrap bg-rose-50/30">
                            {targetTr.trCode}
                          </td>
                          <td className="py-3 px-3 text-center font-mono text-xs font-bold text-rose-600 whitespace-nowrap bg-rose-50/30">
                            {targetTr.ratedKva}
                          </td>
                          <td className="py-3 px-3 text-center font-mono text-xs font-bold text-rose-800 bg-rose-50/30">
                            <div>{format2Decimals(targetTr.year3LoadPct)}</div>
                            <div className="text-[10px] text-emerald-600 font-sans font-medium">โหลดใหม่ {format2Decimals(targetNewLoadPct)}%</div>
                          </td>
                          <td className="py-3 px-3 text-center font-mono text-xs font-bold text-indigo-700 bg-indigo-50/30 whitespace-nowrap">
                            {distanceKm !== null ? `${distanceKm} km` : '-'}
                          </td>
                          <td className="py-3 px-3 text-center whitespace-nowrap space-x-1">
                            <button
                              onClick={() => setSelectedTransformer(sourceTr)}
                              className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-medium rounded transition-colors cursor-pointer"
                            >
                              รายละเอียด
                            </button>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Pagination Controls for TR Matching */}
            {filteredMatches.length > 0 && (
              <div className="bg-white rounded-xl px-5 py-3 border border-slate-200 shadow-xs flex items-center justify-between text-xs">
                <span className="text-slate-500">
                  แสดงรายการที่ {((matchPage - 1) * MATCH_ITEMS_PER_PAGE) + 1} ถึง {Math.min(matchPage * MATCH_ITEMS_PER_PAGE, filteredMatches.length)} จากทั้งหมด {filteredMatches.length.toLocaleString()} คู่ (แสดงหน้าละ {MATCH_ITEMS_PER_PAGE} รายการ)
                </span>
                <div className="flex items-center gap-1.5">
                  <button
                    onClick={() => setMatchPage(prev => Math.max(prev - 1, 1))}
                    disabled={matchPage === 1}
                    className="px-3 py-1.5 rounded-lg border border-slate-200 bg-white text-slate-700 hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed font-medium cursor-pointer"
                  >
                    ก่อนหน้า
                  </button>
                  <span className="px-3 py-1.5 font-medium text-slate-700">
                    หน้า {matchPage} จาก {totalMatchPages}
                  </span>
                  <button
                    onClick={() => setMatchPage(prev => Math.min(prev + 1, totalMatchPages))}
                    disabled={matchPage === totalMatchPages}
                    className="px-3 py-1.5 rounded-lg border border-slate-200 bg-white text-slate-700 hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed font-medium cursor-pointer"
                  >
                    ถัดไป
                  </button>
                </div>
              </div>
            )}
          </div>
        )}
        {activeTab === 'map' && (
          <div className="space-y-6">
            <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-xs flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
              <div>
                <h2 className="text-lg font-bold text-slate-900">แผนที่แสดงตำแหน่งหม้อแปลงเข้าเกณฑ์ (Map View)</h2>
                <p className="text-xs text-slate-500">แสดงตำแหน่งพิกัด GPS (ละติจูด/ลองจิจูด) ของหม้อแปลงที่เข้าเกณฑ์ Rightsizing ทั่ว 34 กฟฟ.</p>
              </div>
            </div>

            {/* Filter Toolbar */}
            <div className="bg-white rounded-xl p-4 border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-3">
              <div className="flex flex-wrap items-center gap-3 w-full sm:w-auto flex-1">
                <div className="relative flex-1 sm:w-64 max-w-xs">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-2.5" />
                  <input
                    type="text"
                    placeholder="ค้นหา PEA หม้อแปลง, กฟฟ., สถานที่..."
                    value={mapSearchQuery}
                    onChange={(e) => setMapSearchQuery(e.target.value)}
                    className="w-full pl-9 pr-4 py-1.5 text-xs border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-500"
                  />
                </div>

                <select
                  value={mapOffice}
                  onChange={(e) => setMapOffice(e.target.value)}
                  className="py-1.5 px-3 text-xs border border-slate-200 rounded-lg bg-white text-slate-700 focus:outline-none focus:border-indigo-500"
                >
                  <option value="all">ทุก กฟฟ. (34 กฟฟ.)</option>
                  {peaOffices.map(off => (
                    <option key={off} value={off}>{off}</option>
                  ))}
                </select>

                <select
                  value={mapRating}
                  onChange={(e) => setMapRating(e.target.value)}
                  className="py-1.5 px-3 text-xs border border-slate-200 rounded-lg bg-white text-slate-700 focus:outline-none focus:border-indigo-500"
                >
                  <option value="all">ทุกพิกัด kVA</option>
                  <option value="50">พิกัด 50 kVA</option>
                  <option value="100">พิกัด 100 kVA</option>
                  <option value="160">พิกัด 160 kVA</option>
                  <option value="250">พิกัด 250 kVA</option>
                </select>

                {(mapSearchQuery || mapOffice !== 'all' || mapRating !== 'all') && (
                  <button
                    onClick={() => {
                      setMapSearchQuery('');
                      setMapOffice('all');
                      setMapRating('all');
                      setMapCenter({ lat: 6.845, lng: 101.233 });
                      setMapZoom(9);
                    }}
                    className="text-xs text-rose-600 hover:text-rose-700 font-medium cursor-pointer"
                  >
                    ล้างตัวกรอง
                  </button>
                )}

                <button
                  onClick={() => {
                    const q = mapSearchQuery.toLowerCase().trim();
                    const officeFilter = mapOffice !== 'all' ? mapOffice : null;

                    // Find eligible transformers matching query or office
                    const matches = stats.eligibleList.filter(({ tr }) => {
                      if (officeFilter && tr.feaOffice !== officeFilter) return false;
                      if (mapRating !== 'all' && tr.ratedKva !== Number(mapRating)) return false;
                      if (q) {
                        const matchCode = tr.trCode.toLowerCase().includes(q);
                        const matchOffice = tr.feaOffice.toLowerCase().includes(q);
                        const matchLoc = (tr.location || '').toLowerCase().includes(q);
                        const matchRaw = tr.trCode.replace(/-/g, '').includes(q.replace(/-/g, ''));
                        if (!matchCode && !matchOffice && !matchLoc && !matchRaw) return false;
                      }
                      return (tr as any).latitude && (tr as any).longitude;
                    });

                    if (matches.length > 0) {
                      // Compute average lat/lng of matches to center zoom
                      let sumLat = 0;
                      let sumLng = 0;
                      matches.forEach(({ tr }) => {
                        sumLat += Number((tr as any).latitude);
                        sumLng += Number((tr as any).longitude);
                      });
                      const avgLat = sumLat / matches.length;
                      const avgLng = sumLng / matches.length;

                      setMapCenter({ lat: avgLat, lng: avgLng });
                      // If specific office or single match, zoom closer (12), otherwise 11
                      setMapZoom(matches.length === 1 ? 15 : 12);
                    } else {
                      alert(`ไม่พบพิกัดตำแหน่ง GPS ของหม้อแปลงเข้าเกณฑ์ที่ตรงกับเงื่อนไขการค้นหา`);
                    }
                  }}
                  className="inline-flex items-center gap-1.5 px-3.5 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold rounded-lg shadow-sm transition-all cursor-pointer"
                >
                  <Search className="w-3.5 h-3.5" />
                  <span>ค้นหา (Zoom)</span>
                </button>
              </div>

              {(() => {
                const filteredMapList = stats.eligibleList.filter(({ tr }) => {
                  if (mapOffice !== 'all' && tr.feaOffice !== mapOffice) return false;
                  if (mapRating !== 'all' && tr.ratedKva !== Number(mapRating)) return false;
                  if (mapSearchQuery.trim()) {
                    const q = mapSearchQuery.toLowerCase().trim();
                    const matchCode = tr.trCode.toLowerCase().includes(q);
                    const matchOffice = tr.feaOffice.toLowerCase().includes(q);
                    const matchLoc = (tr.location || '').toLowerCase().includes(q);
                    if (!matchCode && !matchOffice && !matchLoc) return false;
                  }
                  return true;
                });
                const withCoords = filteredMapList.filter(({ tr }) => (tr as any).latitude && (tr as any).longitude);

                return (
                  <div className="text-xs text-slate-500 font-medium">
                    แสดงหม้อแปลงเข้าเกณฑ์ <strong className="text-emerald-700 font-mono">{filteredMapList.length.toLocaleString()}</strong> เครื่อง (มีพิกัด GPS บนแผนที่ <strong className="text-indigo-700 font-mono">{withCoords.length.toLocaleString()}</strong> จุด)
                  </div>
                );
              })()}
            </div>

            <div className="bg-white rounded-xl p-6 border border-slate-200 shadow-xs">
              <div style={{ height: '650px', width: '100%' }}>
                <MapLibreContainer
                  center={mapCenter}
                  zoom={mapZoom}
                  transformers={transformers}
                  mapOffice={mapOffice}
                  mapRating={mapRating}
                  mapSearchQuery={mapSearchQuery}
                  onSelectTransformer={setSelectedTransformer}
                />
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Modal: Transformer Detail & Simulation */}
      {selectedTransformer && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-xl w-full p-6 shadow-2xl border border-slate-200 space-y-6">
            <div className="flex items-center justify-between border-b border-slate-100 pb-4">
              <div>
                <span className="text-xs font-semibold text-emerald-600 uppercase tracking-wider">รายละเอียดและจำลองการปรับลด</span>
                <h3 className="text-lg font-bold text-slate-900 font-mono">PEA หม้อแปลง: {selectedTransformer.trCode}</h3>
              </div>
              <button
                onClick={() => setSelectedTransformer(null)}
                className="w-8 h-8 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-600 flex items-center justify-center cursor-pointer transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {(() => {
              const evalRes = evaluateTransformerRightsizing(selectedTransformer);
              return (
                <div className="space-y-4 text-sm">
                  <div className="grid grid-cols-2 gap-4 bg-slate-50 p-4 rounded-xl border border-slate-200">
                    <div>
                      <span className="text-xs text-slate-500 block">PEA หม้อแปลง</span>
                      <strong className="text-slate-900 font-mono text-base">{selectedTransformer.trCode}</strong>
                    </div>
                    <div>
                      <span className="text-xs text-slate-500 block">สังกัด กฟฟ. (คอลัมน์ AG)</span>
                      <strong className="text-slate-900">{selectedTransformer.feaOffice}</strong>
                    </div>
                    <div>
                      <span className="text-xs text-slate-500 block">ฟีดเดอร์</span>
                      <strong className="text-slate-900">{selectedTransformer.feeder}</strong>
                    </div>
                    <div>
                      <span className="text-xs text-slate-500 block">พิกัดปัจจุบัน</span>
                      <strong className="text-indigo-600 font-mono text-base">{selectedTransformer.ratedKva} kVA</strong>
                      {evalRes.isEligible && evalRes.recommendedKva && (
                        <div className="mt-2 pt-1 border-t border-slate-200/60">
                          <span className="text-xs text-emerald-700 font-semibold block">พิกัดที่แนะนำให้ลด</span>
                          <strong className="text-emerald-600 font-mono text-base">{evalRes.recommendedKva} kVA</strong>
                        </div>
                      )}
                    </div>
                    <div>
                      <span className="text-xs text-slate-500 block">สถานที่ติดตั้ง</span>
                      <span className="text-slate-900 text-xs">{selectedTransformer.location}</span>
                    </div>
                  </div>

                  <div>
                    <h4 className="font-bold text-slate-900 text-xs uppercase tracking-wider mb-2">ประวัติ % โหลด 3 ปีย้อนหลัง (วิเคราะห์ย้อนกลับ 2 ปี)</h4>
                    <div className="grid grid-cols-3 gap-3 text-center">
                      <div className="bg-emerald-50 border border-emerald-200 p-3 rounded-xl">
                        <span className="text-[11px] text-slate-500 block font-medium">โหลด 2 ปีก่อน (%)</span>
                        <strong className="text-emerald-700 text-base font-mono">{format2Decimals(selectedTransformer.year1LoadPct)}</strong>
                      </div>
                      <div className="bg-emerald-50 border border-emerald-200 p-3 rounded-xl">
                        <span className="text-[11px] text-slate-500 block font-medium">โหลด 1 ปีก่อน (%)</span>
                        <strong className="text-emerald-700 text-base font-mono">{format2Decimals(selectedTransformer.year2LoadPct)}</strong>
                      </div>
                      <div className="bg-indigo-50 border border-indigo-200 p-3 rounded-xl">
                        <span className="text-[11px] text-indigo-700 block font-semibold">โหลดปีปัจจุบัน (%)</span>
                        <strong className="text-indigo-800 text-base font-mono">{format2Decimals(selectedTransformer.year3LoadPct)}</strong>
                      </div>
                    </div>
                  </div>

                  <div className={`p-4 rounded-xl border ${evalRes.isEligible ? 'bg-emerald-50 border-emerald-200 text-emerald-900' : 'bg-amber-50 border-amber-200 text-amber-900'}`}>
                    <h4 className="font-bold text-xs uppercase tracking-wider mb-1 flex items-center gap-1.5">
                      {evalRes.isEligible ? <CheckCircle2 className="w-4 h-4 text-emerald-600" /> : <AlertTriangle className="w-4 h-4 text-amber-600" />}
                      ผลการวิเคราะห์ Rightsizing
                    </h4>
                    <p className="text-xs leading-relaxed mt-1">{evalRes.reason}</p>
                    {evalRes.isEligible && (
                      <div className="mt-3 pt-3 border-t border-emerald-200/70 space-y-2.5">
                        <div className="flex items-center justify-between text-xs font-bold text-emerald-900">
                          <span>พิกัดใหม่ที่แนะนำ: {evalRes.recommendedKva} kVA</span>
                          <span>% โหลดใหม่: {format2Decimals(evalRes.newLoadPct)}% (&lt; 50%)</span>
                        </div>
                        <div className="bg-white/80 rounded-lg p-3 text-[11px] font-mono text-slate-700 border border-emerald-200 space-y-1.5">
                          <div className="font-bold text-emerald-800 font-sans text-xs flex items-center gap-1">
                            <Calculator className="w-3.5 h-3.5 text-emerald-600" />
                            ขั้นตอนการคำนวณ % โหลดใหม่ (&lt;50%):
                          </div>
                          <div className="text-slate-600">
                            • คำนวณโหลด: ({format2Decimals(selectedTransformer.year3LoadPct)}% ÷ 100) × {selectedTransformer.ratedKva} = <strong>{format2Decimals(evalRes.currentLoadKva)} kVA</strong>
                          </div>
                          <div className="text-slate-600">
                            • คำนวณ % โหลดใหม่: ({format2Decimals(evalRes.currentLoadKva)} ÷ {evalRes.recommendedKva}) × 100 = <strong className="text-emerald-700 text-xs">{format2Decimals(evalRes.newLoadPct)}%</strong>
                          </div>
                          <div className="text-emerald-700 text-[10px] font-sans font-semibold pt-1 border-t border-emerald-100 flex items-center gap-1">
                            <CheckCircle2 className="w-3 h-3 text-emerald-600 shrink-0" />
                            <span>ผลลัพธ์: % โหลดใหม่ {format2Decimals(evalRes.newLoadPct)}% น้อยกว่า 50 เป็นไปตามเงื่อนไขของระบบนี้</span>
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              );
            })()}

            <div className="flex justify-end gap-3 pt-2 border-t border-slate-100">
              <button
                onClick={() => setSelectedTransformer(null)}
                className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-lg transition-colors cursor-pointer"
              >
                ปิดหน้าต่าง
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal: Add Transformer */}
      {showAddModal && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-lg w-full p-6 shadow-2xl border border-slate-200 space-y-6">
            <div className="flex items-center justify-between border-b border-slate-100 pb-4">
              <h3 className="text-lg font-bold text-slate-900">เพิ่มข้อมูลหม้อแปลงเพื่อทดสอบ Rightsizing</h3>
              <button
                onClick={() => setShowAddModal(false)}
                className="w-8 h-8 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-600 flex items-center justify-center cursor-pointer transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <form onSubmit={handleAddTransformer} className="space-y-4 text-xs">
              <div>
                <label className="block font-semibold text-slate-700 mb-1">PEA หม้อแปลง</label>
                <input
                  type="text"
                  required
                  placeholder="เช่น 52-001928 หรือ 50-001294 (ระบุเลข PEA หม้อแปลงครบทุกหลัก)"
                  value={newTr.trCode}
                  onChange={(e) => setNewTr({ ...newTr, trCode: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:border-emerald-500 font-mono"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">สังกัด กฟฟ. (คอลัมน์ AG)</label>
                  <select
                    value={newTr.feaOffice}
                    onChange={(e) => setNewTr({ ...newTr, feaOffice: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-200 rounded-lg bg-white text-slate-700 focus:outline-none focus:border-emerald-500"
                  >
                    {peaOffices.map(off => (
                      <option key={off} value={off}>{off}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">พิกัดหม้อแปลงเดิม (kVA)</label>
                  <select
                    value={newTr.ratedKva}
                    onChange={(e) => setNewTr({ ...newTr, ratedKva: Number(e.target.value) })}
                    className="w-full px-3 py-2 border border-slate-200 rounded-lg bg-white text-slate-700 focus:outline-none focus:border-emerald-500 font-mono"
                  >
                    <option value={100}>100 kVA (ลดเหลือ 50)</option>
                    <option value={160}>160 kVA (ลดเหลือ 100/50)</option>
                    <option value={250}>250 kVA (ลดเหลือ 160/100/50)</option>
                    <option value={50}>50 kVA</option>
                    <option value={315}>315 kVA</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-3 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">% โหลด ข้อมูล 2 ปีก่อน</label>
                  <input
                    type="number"
                    step="0.1"
                    required
                    value={newTr.year1LoadPct}
                    onChange={(e) => setNewTr({ ...newTr, year1LoadPct: Number(e.target.value) })}
                    className="w-full px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:border-emerald-500 font-mono"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">% โหลด ข้อมูล 1 ปีก่อน</label>
                  <input
                    type="number"
                    step="0.1"
                    required
                    value={newTr.year2LoadPct}
                    onChange={(e) => setNewTr({ ...newTr, year2LoadPct: Number(e.target.value) })}
                    className="w-full px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:border-emerald-500 font-mono"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">% โหลด ข้อมูลปีปัจจุบัน</label>
                  <input
                    type="number"
                    step="0.1"
                    required
                    value={newTr.year3LoadPct}
                    onChange={(e) => setNewTr({ ...newTr, year3LoadPct: Number(e.target.value) })}
                    className="w-full px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:border-emerald-500 font-mono"
                  />
                </div>
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">สถานที่ติดตั้ง (ชีท ข้อมูลปีปัจจุบัน คอลัมน์ B)</label>
                <input
                  type="text"
                  value={newTr.location}
                  onChange={(e) => setNewTr({ ...newTr, location: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-200 rounded-lg focus:outline-none focus:border-emerald-500"
                />
              </div>

              <div className="flex justify-end gap-3 pt-3">
                <button
                  type="button"
                  onClick={() => setShowAddModal(false)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold rounded-lg transition-colors cursor-pointer"
                >
                  ยกเลิก
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-semibold rounded-lg shadow-sm transition-all cursor-pointer"
                >
                  บันทึกข้อมูล
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Modal: Import Data from Sheet */}
      {showImportModal && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-2xl w-full p-6 shadow-2xl border border-slate-200 space-y-5">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div>
                <span className="text-xs font-semibold text-indigo-600 uppercase tracking-wider">นำเข้าฐานข้อมูล Google Sheets</span>
                <h3 className="text-lg font-bold text-slate-900">ฐานข้อมูลชีท "ข้อมูลปีปัจจุบัน" (ตั้งต้น) และ ย้อนหลัง 2 ปี</h3>
              </div>
              <button
                onClick={() => setShowImportModal(false)}
                className="w-8 h-8 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-600 flex items-center justify-center cursor-pointer transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Sheet Selector Tabs */}
            <div className="flex border-b border-slate-200 gap-1 overflow-x-auto text-xs font-medium">
              <button
                type="button"
                onClick={() => setImportTab('current')}
                className={`py-2 px-3 border-b-2 rounded-t-lg transition-colors cursor-pointer ${
                  importTab === 'current'
                    ? 'border-indigo-600 text-indigo-700 bg-indigo-50 font-bold'
                    : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-50'
                }`}
              >
                1. ข้อมูลปีปัจจุบัน (ข้อมูลตั้งต้น) {currentYearInput ? '✓' : ''}
              </button>
              <button
                type="button"
                onClick={() => setImportTab('year1')}
                className={`py-2 px-3 border-b-2 rounded-t-lg transition-colors cursor-pointer ${
                  importTab === 'year1'
                    ? 'border-indigo-600 text-indigo-700 bg-indigo-50 font-bold'
                    : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-50'
                }`}
              >
                2. ข้อมูล 1 ปีก่อน (ย้อนหลัง 1 ปี) {year1Input ? '✓' : ''}
              </button>
              <button
                type="button"
                onClick={() => setImportTab('year2')}
                className={`py-2 px-3 border-b-2 rounded-t-lg transition-colors cursor-pointer ${
                  importTab === 'year2'
                    ? 'border-indigo-600 text-indigo-700 bg-indigo-50 font-bold'
                    : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-50'
                }`}
              >
                3. ข้อมูล 2 ปีก่อน (ย้อนหลัง 2 ปี) {year2Input ? '✓' : ''}
              </button>
              <button
                type="button"
                onClick={() => setImportTab('all')}
                className={`py-2 px-3 border-b-2 rounded-t-lg transition-colors cursor-pointer ${
                  importTab === 'all'
                    ? 'border-indigo-600 text-indigo-700 bg-indigo-50 font-bold'
                    : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-50'
                }`}
              >
                4. วางแบบรวม (All-in-One)
              </button>
            </div>

            <div className="text-xs text-slate-600 space-y-1.5 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
              {importTab === 'current' && (
                <div>
                  <p className="font-bold text-indigo-900">📋 ชีท "ข้อมูลปีปัจจุบัน" (ฐานข้อมูลตั้งต้นหลัก):</p>
                  <p className="text-slate-600">
                    ระบบจะใช้ <strong>PEA หม้อแปลง จากคอลัมน์ A</strong> เป็นรายการหม้อแปลงตั้งต้นทั้งหมด, <strong>สถานที่ติดตั้ง จากคอลัมน์ B</strong>, <strong>พิกัดเดิม (kVA) จากคอลัมน์ D</strong>, <strong>โหลดปีปัจจุบัน (%) จากคอลัมน์ E</strong>, และ <strong>กฟฟ. จากคอลัมน์ AG</strong>
                  </p>
                </div>
              )}
              {importTab === 'year1' && (
                <div>
                  <p className="font-bold text-indigo-900">📋 ชีท "ข้อมูล 1 ปีก่อน" (วิเคราะห์ย้อนกลับ 1 ปี):</p>
                  <p className="text-slate-600">
                    ระบบจะใช้ <strong>PEA หม้อแปลง จากคอลัมน์ A</strong> ในการจับคู่ และดึง <strong>โหลด 1 ปีก่อน (%) จากคอลัมน์ E</strong>
                  </p>
                </div>
              )}
              {importTab === 'year2' && (
                <div>
                  <p className="font-bold text-indigo-900">📋 ชีท "ข้อมูล 2 ปีก่อน" (วิเคราะห์ย้อนกลับ 2 ปี):</p>
                  <p className="text-slate-600">
                    ระบบจะใช้ <strong>PEA หม้อแปลง จากคอลัมน์ A</strong> ในการจับคู่ และดึง <strong>โหลด 2 ปีก่อน (%) จากคอลัมน์ E</strong>
                  </p>
                </div>
              )}
              {importTab === 'all' && (
                <div>
                  <p className="font-bold text-indigo-900">📋 วางข้อมูลแบบรวม 3 ปี ในครั้งเดียว:</p>
                  <p className="text-slate-600">
                    วางข้อมูลที่รวมคอลัมน์ A: PEA หม้อแปลง, AG: กฟฟ. และประวัติโหลดทั้ง 3 ปี
                  </p>
                </div>
              )}
            </div>

            <div className="space-y-2">
              <div className="flex items-center justify-between text-xs">
                <label className="font-semibold text-slate-700">
                  {importTab === 'current' && 'วางข้อความชีท "ข้อมูลปีปัจจุบัน" (TSV / CSV):'}
                  {importTab === 'year1' && 'วางข้อความชีท "ข้อมูล 1 ปีก่อน" (TSV / CSV):'}
                  {importTab === 'year2' && 'วางข้อความชีท "ข้อมูล 2 ปีก่อน" (TSV / CSV):'}
                  {importTab === 'all' && 'วางข้อมูลรวมทุกชีท (TSV / CSV):'}
                </label>
                <label className="text-indigo-600 font-semibold cursor-pointer hover:underline flex items-center gap-1">
                  <FileUp className="w-3.5 h-3.5" />
                  <span>เลือกไฟล์ CSV</span>
                  <input
                    type="file"
                    accept=".csv,.tsv,.txt"
                    className="hidden"
                    onChange={(e) => {
                      const file = e.target.files?.[0];
                      if (file) {
                        const reader = new FileReader();
                        reader.onload = (evt) => {
                          const content = evt.target?.result as string;
                          if (content) {
                            if (importTab === 'current') setCurrentYearInput(content);
                            else if (importTab === 'year1') setYear1Input(content);
                            else if (importTab === 'year2') setYear2Input(content);
                            else setPastedText(content);
                          }
                        };
                        reader.readAsText(file);
                      }
                    }}
                  />
                </label>
              </div>

              {importTab === 'current' && (
                <textarea
                  rows={8}
                  value={currentYearInput}
                  onChange={(e) => setCurrentYearInput(e.target.value)}
                  placeholder="วางข้อมูลจากชีท ข้อมูลปีปัจจุบัน ที่นี่ (คอลัมน์ A: PEA หม้อแปลง, คอลัมน์ AG: กฟฟ.)..."
                  className="w-full p-3 text-xs font-mono border border-slate-300 rounded-xl focus:outline-none focus:border-indigo-500 bg-slate-50"
                />
              )}
              {importTab === 'year1' && (
                <textarea
                  rows={8}
                  value={year1Input}
                  onChange={(e) => setYear1Input(e.target.value)}
                  placeholder="วางข้อมูลจากชีท ข้อมูล 1 ปีก่อน ที่นี่ (คอลัมน์ A: PEA หม้อแปลง, % โหลด)..."
                  className="w-full p-3 text-xs font-mono border border-slate-300 rounded-xl focus:outline-none focus:border-indigo-500 bg-slate-50"
                />
              )}
              {importTab === 'year2' && (
                <textarea
                  rows={8}
                  value={year2Input}
                  onChange={(e) => setYear2Input(e.target.value)}
                  placeholder="วางข้อมูลจากชีท ข้อมูล 2 ปีก่อน ที่นี่ (คอลัมน์ A: PEA หม้อแปลง, % โหลด)..."
                  className="w-full p-3 text-xs font-mono border border-slate-300 rounded-xl focus:outline-none focus:border-indigo-500 bg-slate-50"
                />
              )}
              {importTab === 'all' && (
                <textarea
                  rows={8}
                  value={pastedText}
                  onChange={(e) => setPastedText(e.target.value)}
                  placeholder="วางข้อมูลแบบรวมทุกชีทที่นี่ (ระบบจะอ่านคอลัมน์ A เป็น PEA หม้อแปลง และ AG เป็น 34 กฟฟ.)..."
                  className="w-full p-3 text-xs font-mono border border-slate-300 rounded-xl focus:outline-none focus:border-indigo-500 bg-slate-50"
                />
              )}
            </div>

            {/* Status overview badges */}
            <div className="flex flex-wrap gap-2 text-[11px]">
              <span className={`px-2 py-0.5 rounded border ${currentYearInput ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : 'bg-slate-100 text-slate-500 border-slate-200'}`}>
                ชีท ข้อมูลปีปัจจุบัน: {currentYearInput ? '✓ พร้อมประมวลผล' : 'ใช้ชุดข้อมูลหลัก'}
              </span>
              <span className={`px-2 py-0.5 rounded border ${year1Input ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : 'bg-slate-100 text-slate-500 border-slate-200'}`}>
                ชีท ข้อมูล 1 ปีก่อน: {year1Input ? '✓ พร้อมเชื่อมโยง' : 'ค่าตั้งต้น'}
              </span>
              <span className={`px-2 py-0.5 rounded border ${year2Input ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : 'bg-slate-100 text-slate-500 border-slate-200'}`}>
                ชีท ข้อมูล 2 ปีก่อน: {year2Input ? '✓ พร้อมเชื่อมโยง' : 'ค่าตั้งต้น'}
              </span>
            </div>

            {importStatus && (
              <div className={`p-3 rounded-lg text-xs font-medium ${importStatus.includes('สำเร็จ') ? 'bg-emerald-50 text-emerald-800 border border-emerald-200' : 'bg-rose-50 text-rose-800 border border-rose-200'}`}>
                {importStatus}
              </div>
            )}

            <div className="flex items-center justify-between pt-3 border-t border-slate-100">
              <span className="text-[11px] text-slate-400">
                * คอลัมน์ A: PEA หม้อแปลง | คอลัมน์ AG: 34 กฟฟ.
              </span>
              <div className="flex items-center gap-3">
                <button
                  type="button"
                  onClick={() => setShowImportModal(false)}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-lg transition-colors cursor-pointer"
                >
                  ยกเลิก
                </button>
                <button
                  type="button"
                  onClick={() => {
                    const baseContent = currentYearInput.trim() || pastedText.trim();
                    if (!baseContent) {
                      setImportStatus('กรุณาวางข้อมูลชีท "ข้อมูลปีปัจจุบัน" (คอลัมน์ A: PEA หม้อแปลง) ก่อนกดประมวลผล');
                      return;
                    }
                    const ok = processBaseAndHistoricalData(baseContent, year1Input, year2Input);
                    if (ok) {
                      setImportStatus('นำเข้าสำเร็จ: ประมวลผลโดยใช้ PEA หม้อแปลง จากชีท ข้อมูลปีปัจจุบัน (คอลัมน์ A) เป็นข้อมูลตั้งต้น และวิเคราะห์ย้อนกลับ 2 ปี เรียบร้อยแล้ว!');
                      setTimeout(() => {
                        setShowImportModal(false);
                        setImportStatus('');
                      }, 1200);
                    } else {
                      setImportStatus('ไม่สามารถอ่านข้อมูลได้ กรุณาตรวจสอบว่ามีข้อมูลคอลัมน์ A (PEA หม้อแปลง)');
                    }
                  }}
                  className="px-4 py-2 bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold rounded-lg shadow-sm transition-all cursor-pointer flex items-center gap-2"
                >
                  <CheckCircle2 className="w-4 h-4" />
                  <span>ประมวลผลวิเคราะห์ย้อนกลับ 2 ปี</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Modal: Alert Notification Center */}
      {showAlertModal && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-lg w-full p-6 shadow-2xl border border-slate-200 space-y-6">
            <div className="flex items-center justify-between border-b border-slate-100 pb-4">
              <div className="flex items-center gap-2">
                <Bell className="w-5 h-5 text-rose-500" />
                <h3 className="text-lg font-bold text-slate-900">การแจ้งเตือนหม้อแปลงเข้าเกณฑ์ Rightsizing</h3>
              </div>
              <button
                onClick={() => setShowAlertModal(false)}
                className="w-8 h-8 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-600 flex items-center justify-center cursor-pointer transition-colors"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="space-y-3 max-h-80 overflow-y-auto pr-2">
              <div className="p-3 bg-rose-50 border border-rose-200 rounded-xl text-xs text-rose-900">
                <strong>ตรวจพบหม้อแปลงขนาดพิกัด 100, 160, 250 kVA</strong> ที่มี % โหลดเฉลี่ยต่ำกว่า 30% ติดต่อกัน 3 ปี และเมื่อลดขนาดแล้วมี % โหลดใหม่ &lt; 50% จำนวนทั้งสิ้น <strong>{stats.totalEligible} เครื่อง</strong> ทั่ว 34 กฟฟ.
              </div>
              {stats.eligibleList.slice(0, 5).map(({ tr, evalRes }) => (
                <div key={tr.id} className="p-3 bg-slate-50 border border-slate-200 rounded-xl text-xs space-y-1">
                  <div className="flex justify-between font-bold text-slate-900">
                    <span className="font-mono">{tr.trCode} ({tr.ratedKva} kVA)</span>
                    <span className="text-emerald-600">ลดเป็น {evalRes.recommendedKva} kVA</span>
                  </div>
                  <p className="text-slate-600">{tr.feaOffice}</p>
                  <p className="text-[11px] text-slate-500">โหลด 3 ปีย้อนหลัง: {tr.year1LoadPct}%, {tr.year2LoadPct}%, {tr.year3LoadPct}%</p>
                </div>
              ))}
            </div>

            <div className="flex justify-end gap-3 pt-2 border-t border-slate-100">
              <button
                onClick={() => {
                  setShowAlertModal(false);
                  setActiveTab('target');
                }}
                className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded-lg shadow-sm transition-all cursor-pointer"
              >
                ดูรายการทั้งหมด
              </button>
            </div>
          </div>
        </div>
      )}




      <footer className="bg-white border-t border-slate-200 py-4 text-center text-xs text-slate-500">
        <div className="max-w-7xl mx-auto px-4 space-y-1">
          <div>© 2026 PEA Transformer Rightsizing System — PI InnoTech</div>
          <div className="text-[11px] text-slate-400">Provincial Electricity Authority</div>
        </div>
      </footer>
    </div>
  );
}
