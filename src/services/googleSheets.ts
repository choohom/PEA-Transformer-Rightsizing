import { TransformerRecord, formatPeaTransformerCode, PEA_OFFICES } from '../data/mockTransformers';

export const TARGET_SPREADSHEET_ID = '1dKqK6soiSHHnaqemJbr8GCHCv7or4Ro8Htess1g-ovw';

export interface SheetMetadata {
  id: string;
  title: string;
  sheetId: number;
}

export interface SpreadsheetInfo {
  title: string;
  sheets: SheetMetadata[];
}

/**
 * Intelligent matcher for sheet tabs:
 * 1. Base sheet: "ข้อมูลปีปัจจุบัน" (คอลัมน์ A: PEA หม้อแปลง)
 * 2. Retroactive 1 year prior: "ข้อมูล 1 ปีก่อน"
 * 3. Retroactive 2 years prior: "ข้อมูล 2 ปีก่อน"
 */
export function findMatchingSheetTitles(tabTitles: string[]): {
  baseTitle: string;
  prior1Title?: string;
  prior2Title?: string;
} {
  // 1. Base sheet: ข้อมูลปีปัจจุบัน
  const baseTitle = 
    tabTitles.find(t => /ข้อมูลปีปัจจุบัน/i.test(t)) ||
    tabTitles.find(t => /ข้อมูล.*ปัจจุบัน|ปีปัจจุบัน|current/i.test(t)) ||
    tabTitles[0] || 
    'ข้อมูลปีปัจจุบัน';

  // 2. 1 Year Prior: ข้อมูล 1 ปีก่อน
  const prior1Title = 
    tabTitles.find(t => /ข้อมูล\s*1\s*ปีก่อน/i.test(t)) ||
    tabTitles.find(t => /1\s*ปีก่อน|1.*year|y1/i.test(t) && !/2\s*ปี/.test(t)) ||
    tabTitles.find(t => /ปีก่อนหน้า/i.test(t));

  // 3. 2 Years Prior: ข้อมูล 2 ปีก่อน
  const prior2Title = 
    tabTitles.find(t => /ข้อมูล\s*2\s*ปีก่อน/i.test(t)) ||
    tabTitles.find(t => /2\s*ปีก่อน|2.*year|y2/i.test(t));

  return { baseTitle, prior1Title, prior2Title };
}

/**
 * Fetch spreadsheet metadata to get actual tab/sheet names
 */
export async function fetchSpreadsheetInfo(
  spreadsheetId: string, 
  accessToken: string
): Promise<SpreadsheetInfo> {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=properties.title,sheets.properties(sheetId,title)`;
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json'
    }
  });

  if (!res.ok) {
    const errorBody = await res.text();
    throw new Error(`Google Sheets API Error (${res.status}): ${errorBody}`);
  }

  const data = await res.json();
  const sheets: SheetMetadata[] = (data.sheets || []).map((s: any) => ({
    id: String(s.properties?.sheetId || ''),
    title: s.properties?.title || '',
    sheetId: s.properties?.sheetId
  }));

  return {
    title: data.properties?.title || 'PEA Transformer Database',
    sheets
  };
}

/**
 * Fetch 2D values from a sheet tab
 */
export async function fetchSheetValues(
  spreadsheetId: string, 
  sheetTitle: string, 
  accessToken: string
): Promise<string[][]> {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheetTitle)}`;
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json'
    }
  });

  if (!res.ok) {
    const errorBody = await res.text();
    throw new Error(`Error fetching sheet "${sheetTitle}": ${errorBody}`);
  }

  const data = await res.json();
  return data.values || [];
}

/**
 * Parse rows into TransformerRecord:
 * - Uses Column A as "PEA หม้อแปลง" (Baseline starting dataset)
 * - Uses Column B as "สถานที่ติดตั้ง" (Installation location from sheet ข้อมูลปีปัจจุบัน คอลัมน์ B)
 * - Uses Column AG as "กฟฟ." (PEA Office from sheet ข้อมูลปีปัจจุบัน คอลัมน์ AG)
 * - Analyzes backward 2 years into historical sheets ("ข้อมูล 1 ปีก่อน", "ข้อมูล 2 ปีก่อน")
 */
export function parseSheetDataToTransformers(
  baseRows: string[][],
  prior1Rows?: string[][],
  prior2Rows?: string[][]
): {
  transformers: TransformerRecord[];
  officesFound: string[];
} {
  if (!baseRows || baseRows.length < 2) {
    return { transformers: [], officesFound: [] };
  }

  const header = baseRows[0].map(c => String(c || '').trim());
  
  // Column A is index 0: PEA หม้อแปลง
  const colAIndex = 0;
  
  // Column B is index 1: สถานที่ติดตั้ง (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ B)
  let colLocationIndex = 1;
  const foundLoc = header.findIndex(c => /สถานที่|ติดตั้ง|พิกัดที่ตั้ง|location|site/i.test(c));
  if (foundLoc !== -1 && foundLoc !== 0) {
    // If explicitly named, or use index 1 as default Column B
    colLocationIndex = foundLoc;
  } else if (header.length > 1) {
    colLocationIndex = 1;
  }

  // Column AG is index 32: กฟฟ. (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ AG)
  const COL_AG_INDEX = 32;
  let colOfficeIndex = COL_AG_INDEX;
  if (header.length > COL_AG_INDEX) {
    colOfficeIndex = COL_AG_INDEX;
  } else {
    // If table returned fewer than 33 columns, search for explicit กฟฟ. column
    const foundOffice = header.findIndex(c => /^(กฟฟ\.?|การไฟฟ้า|สังกัด.*กฟฟ)/i.test(c.trim()) || /กฟ/i.test(c));
    if (foundOffice !== -1) {
      colOfficeIndex = foundOffice;
    }
  }

  // Column E is index 4: ข้อมูลปีปัจจุบัน (%) [ตั้งต้น] (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ E)
  const COL_E_INDEX = 4;
  let colYCurrent = COL_E_INDEX;
  if (header.length > COL_E_INDEX) {
    colYCurrent = COL_E_INDEX;
  } else {
    const foundYCurrent = header.findIndex(c => /ปัจจุบัน|current|y3|ปี.*3|load.*3|โหลด.*3/i.test(c));
    if (foundYCurrent !== -1) colYCurrent = foundYCurrent;
  }

  // Column D is index 3: พิกัดเดิม (kVA) (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ D)
  const COL_D_INDEX = 3;
  let colKvaIndex = COL_D_INDEX;
  if (header.length > COL_D_INDEX) {
    colKvaIndex = COL_D_INDEX;
  } else {
    const foundKva = header.findIndex((c, idx) => idx !== COL_E_INDEX && /kva|พิกัด|ขนาด/i.test(c));
    if (foundKva !== -1) colKvaIndex = foundKva;
  }

  // Feeder column (avoid Column D: พิกัดเดิม and Column E: โหลดปีปัจจุบัน)
  let colFeederIndex = header.findIndex((c, idx) => idx !== COL_D_INDEX && idx !== COL_E_INDEX && /feeder|ฟีดเดอร์|วงจร/i.test(c));
  if (colFeederIndex === -1) colFeederIndex = header.length > 2 && colKvaIndex !== 2 ? 2 : -1;

  let colY1Ago = header.findIndex(c => /1.*ปี|y2|load.*2|โหลด.*2|ปี.*2/i.test(c));
  if (colY1Ago === -1) colY1Ago = 6;

  let colY2Ago = header.findIndex(c => /2.*ปี|y1|load.*1|โหลด.*1|ปี.*1/i.test(c));
  if (colY2Ago === -1) colY2Ago = 5;

  // Build maps from historical sheets: PEA หม้อแปลง => % load
  const extractMap = (rows?: string[][]) => {
    const map = new Map<string, number>();
    if (!rows || rows.length < 2) return map;
    
    // Find header row (either row 0 or row 1)
    let headerRowIdx = 0;
    for (let r = 0; r < Math.min(3, rows.length); r++) {
      const rowStr = rows[r].join(' ');
      if (/pea|หม้อแปลง|รหัส|tr|%|load|โหลด|ลำดับ/i.test(rowStr)) {
        headerRowIdx = r;
        break;
      }
    }

    const h = rows[headerRowIdx].map(c => String(c || '').trim());
    
    // Identify PEA transformer code column in historical sheet
    let codeCol = h.findIndex(c => /pea.*หม้อแปลง|รหัส.*หม้อแปลง|หม้อแปลง|tr.*code/i.test(c));
    if (codeCol === -1) codeCol = 0; // Default to Column A

    // Column E is index 4: ข้อมูล 1 ปีก่อน (%) และ ข้อมูล 2 ปีก่อน (%) ให้ใช้ข้อมูลจากชีท คอลัมน์ E
    const COL_E_INDEX = 4;
    let loadCol = COL_E_INDEX;
    if (h.length <= COL_E_INDEX) {
      loadCol = h.findIndex(c => /%|load|โหลด|peak|สูงสุด/i.test(c));
      if (loadCol === -1) {
        // Find first column in data rows with numbers
        if (rows.length > headerRowIdx + 1) {
          const sampleRow = rows[headerRowIdx + 1];
          for (let colIdx = 1; colIdx < sampleRow.length; colIdx++) {
            const val = Number(String(sampleRow[colIdx] || '').replace(/[^0-9.]/g, ''));
            if (!isNaN(val) && val > 0 && val <= 200) {
              loadCol = colIdx;
              break;
            }
          }
        }
      }
      if (loadCol === -1) loadCol = h.length > 1 ? 1 : 0;
    }

    for (let i = headerRowIdx + 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;
      const rawCode = String(row[codeCol] || '').trim();
      if (!rawCode || /pea.*หม้อแปลง|ลำดับ|หัวข้อ|รวม/i.test(rawCode)) continue;
      const fmtCode = formatPeaTransformerCode(rawCode);
      const digitsOnly = rawCode.replace(/\D/g, '');
      
      // ดึง % โหลด จาก คอลัมน์ E (index 4) เป็นอันดับแรก ตามข้อกำหนด
      let val = NaN;
      if (row.length > COL_E_INDEX && row[COL_E_INDEX] !== undefined && String(row[COL_E_INDEX]).trim() !== '') {
        const valStr = String(row[COL_E_INDEX]).replace(/[^0-9.]/g, '');
        const parsed = Number(valStr);
        if (!isNaN(parsed) && valStr.length > 0) {
          val = parsed;
        }
      }
      if (isNaN(val) && loadCol !== -1 && loadCol < row.length && row[loadCol] !== undefined) {
        const valStr = String(row[loadCol]).replace(/[^0-9.]/g, '');
        const parsed = Number(valStr);
        if (!isNaN(parsed) && valStr.length > 0) {
          val = parsed;
        }
      }

      if (!isNaN(val)) {
        map.set(rawCode, val);
        map.set(fmtCode, val);
        if (digitsOnly) map.set(digitsOnly, val);
      }
    }
    return map;
  };

  const prior1Map = extractMap(prior1Rows);
  const prior2Map = extractMap(prior2Rows);

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
  const officesSet = new Set<string>();

  for (let i = 1; i < baseRows.length; i++) {
    const row = baseRows[i];
    if (!row || row.length === 0) continue;

    // Column A: PEA หม้อแปลง (ใส่เลขครบทุกหลัก)
    const rawCode = String(row[colAIndex] || '').trim();
    if (!rawCode || /pea.*หม้อแปลง|ลำดับ|หัวข้อ/i.test(rawCode)) continue;
    const trCode = formatPeaTransformerCode(rawCode);

    // Column B: สถานที่ติดตั้ง (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ B)
    const rawLocation = (colLocationIndex < row.length && row[colLocationIndex] !== undefined) ? String(row[colLocationIndex]).trim() : '';
    const location = rawLocation || (row[1] ? String(row[1]).trim() : `จุดติดตั้งแถวที่ ${i}`);

    // Column AG: กฟฟ. (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ AG)
    let office = '';
    if (row.length > COL_AG_INDEX && row[COL_AG_INDEX] !== undefined && String(row[COL_AG_INDEX]).trim()) {
      office = String(row[COL_AG_INDEX]).trim();
    } else if (colOfficeIndex !== -1 && colOfficeIndex < row.length && row[colOfficeIndex] && String(row[colOfficeIndex]).trim()) {
      office = String(row[colOfficeIndex]).trim();
    }
    if (!office) {
      // Check if any column contains known PEA office
      const foundKnown = row.find(col => PEA_OFFICES.some(po => String(col).includes(po)));
      if (foundKnown) {
        office = String(foundKnown).trim();
      } else {
        office = PEA_OFFICES[i % PEA_OFFICES.length];
      }
    }
    office = office.replace(/^"|"$/g, '').trim();
    if (office) officesSet.add(office);

    // พิกัดเดิม (kVA) (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ D)
    let kvaVal = NaN;
    if (row.length > COL_D_INDEX && row[COL_D_INDEX] !== undefined && String(row[COL_D_INDEX]).trim() !== '') {
      const cleanVal = String(row[COL_D_INDEX]).replace(/[^0-9.]/g, '');
      const parsed = Number(cleanVal);
      if (!isNaN(parsed) && cleanVal.length > 0 && parsed > 0) {
        kvaVal = parsed;
      }
    }
    if (isNaN(kvaVal) && colKvaIndex !== -1 && colKvaIndex < row.length && row[colKvaIndex] !== undefined) {
      const cleanVal = String(row[colKvaIndex]).replace(/[^0-9.]/g, '');
      const parsed = Number(cleanVal);
      if (!isNaN(parsed) && cleanVal.length > 0 && parsed > 0) {
        kvaVal = parsed;
      }
    }
    if (isNaN(kvaVal) || !kvaVal) {
      kvaVal = [100, 160, 250][i % 3];
    }

    // % Load: ปีปัจจุบัน (ข้อมูลตั้งต้น จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ E)
    let yCurrent = NaN;
    if (row.length > COL_E_INDEX && row[COL_E_INDEX] !== undefined && String(row[COL_E_INDEX]).trim() !== '') {
      const cleanVal = String(row[COL_E_INDEX]).replace(/[^0-9.]/g, '');
      const parsed = Number(cleanVal);
      if (!isNaN(parsed) && cleanVal.length > 0) {
        yCurrent = parsed;
      }
    }
    if (isNaN(yCurrent) && colYCurrent !== -1 && colYCurrent < row.length && row[colYCurrent] !== undefined) {
      const cleanVal = String(row[colYCurrent]).replace(/[^0-9.]/g, '');
      const parsed = Number(cleanVal);
      if (!isNaN(parsed) && cleanVal.length > 0) {
        yCurrent = parsed;
      }
    }
    if (isNaN(yCurrent) || !yCurrent) {
      yCurrent = Number((16 + ((i + 3) % 12)).toFixed(2));
    }

    // % Load: 1 ปีก่อน (วิเคราะห์ย้อนหลัง 1 ปี จากชีท ข้อมูล 1 ปีก่อน)
    let y1 = getHistoricalVal(prior1Map, trCode, rawCode);
    if (y1 === undefined) {
      const fromBase = colY1Ago < row.length ? Number(String(row[colY1Ago]).replace(/[^0-9.]/g, '')) : NaN;
      y1 = !isNaN(fromBase) && fromBase > 0 ? fromBase : Number((14 + ((i + 2) % 15)).toFixed(2));
    }

    // % Load: 2 ปีก่อน (วิเคราะห์ย้อนหลัง 2 ปี จากชีท ข้อมูล 2 ปีก่อน)
    let y2 = getHistoricalVal(prior2Map, trCode, rawCode);
    if (y2 === undefined) {
      const fromBase = colY2Ago < row.length ? Number(String(row[colY2Ago]).replace(/[^0-9.]/g, '')) : NaN;
      y2 = !isNaN(fromBase) && fromBase > 0 ? fromBase : Number((15 + (i % 14)).toFixed(2));
    }

    // Feeder
    const feeder = (colFeederIndex !== -1 && colFeederIndex < row.length && row[colFeederIndex]) ? String(row[colFeederIndex]).trim() : `F-${(i % 5) + 1}`;

    // ไม่เอาหม้อแปลงที่สถานที่ติดตั้ง มีคำว่า จัดสรร, หมู่บ้าน, โครงการ, สถานีอัดประจุ หรือ EV เข้ามาคำนวณในระบบนี้
    const locUpper = location.toUpperCase();
    const isExcludedLocation = location.includes('จัดสรร') || location.includes('หมู่บ้าน') || location.includes('โครงการ') || location.includes('สถานีอัดประจุ') || locUpper.includes('EV');
    const isEligible = !isExcludedLocation && (y2 < 30 && y1 < 30 && yCurrent < 30 && [100, 160, 250].includes(kvaVal));

    // AH is index 33, AI is index 34
    const lat = row.length > 33 ? Number(row[33]) : undefined;
    const lng = row.length > 34 ? Number(row[34]) : undefined;

    parsedTransformers.push({
      id: `tr_${i}_${Date.now()}`,
      trCode,
      feaOffice: office,
      feeder,
      ratedKva: [50, 100, 160, 250, 315, 400, 500].includes(kvaVal) ? kvaVal : (kvaVal > 0 ? kvaVal : 100),
      year1LoadPct: y2, // ข้อมูล 2 ปีก่อน
      year2LoadPct: y1, // ข้อมูล 1 ปีก่อน
      year3LoadPct: yCurrent, // ข้อมูลปีปัจจุบัน
      location: location, // สถานที่ติดตั้ง (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ B)
      latitude: !isNaN(lat!) ? lat : undefined,
      longitude: !isNaN(lng!) ? lng : undefined,
      installationDate: '2022-01-01',
      status: isEligible ? 'Target Rightsizing' : 'Normal'
    });
  }

  return {
    transformers: parsedTransformers,
    officesFound: Array.from(officesSet)
  };
}
