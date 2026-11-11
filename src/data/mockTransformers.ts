export interface TransformerRecord {
  id: string;
  trCode: string; // PEA หม้อแปลง (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ A เป็นข้อมูลตั้งต้น)
  feaOffice: string; // กฟฟ. (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ AG)
  feeder: string;
  ratedKva: number; // พิกัดเดิม (kVA) (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ D)
  year1LoadPct: number; // โหลด 2 ปีก่อน (%) จากชีท "ข้อมูล 2 ปีก่อน" คอลัมน์ E (วิเคราะห์ย้อนหลัง 2 ปี)
  year2LoadPct: number; // โหลด 1 ปีก่อน (%) จากชีท "ข้อมูล 1 ปีก่อน" คอลัมน์ E (วิเคราะห์ย้อนหลัง 1 ปี)
  year3LoadPct: number; // โหลดปีปัจจุบัน (%) จากชีท "ข้อมูลปีปัจจุบัน" คอลัมน์ E (ข้อมูลตั้งต้น)
  location: string; // สถานที่ติดตั้ง (จากชีท ข้อมูลปีปัจจุบัน คอลัมน์ B)
  latitude?: number; // จากชีท คอลัมน์ AH
  longitude?: number; // จากชีท คอลัมน์ AI
  installationDate: string;
  status: 'Normal' | 'Target Rightsizing' | 'Action Planned' | 'Completed';
}

export const PEA_OFFICES = [
  "กฟจ.ยะลา",
  "กฟส.บันนังสตา",
  "กฟส.รามัน",
  "กฟจ.นราธิวาส",
  "กฟส.ระแงะ",
  "กฟส.รือเสาะ",
  "กฟส.ตากใบ",
  "กฟจ.ปัตตานี",
  "กฟส.โคกโพธิ์",
  "กฟส.หนองจิก",
  "กฟจ.สงขลา",
  "กฟส.สิงหนคร",
  "กฟจ.สตูล",
  "กฟส.ละงู",
  "กฟส.ควนกาหลง",
  "กฟจ.พัทลุง",
  "กฟส.ตะโหมด",
  "กฟส.ควนขนุน",
  "กฟส.ปากพะยูน",
  "กฟส.หาดใหญ่",
  "กฟส.รัตภูมิ",
  "กฟส.นาหม่อม",
  "กฟส.สุไหงโก-ลก",
  "กฟส.เบตง",
  "กฟส.สายบุรี",
  "กฟส.มายอ",
  "กฟส.ระโนด",
  "กฟส.สทิงพระ",
  "กฟส.สะเดา",
  "กฟส.พังลา",
  "กฟส.จะนะ",
  "กฟส.นาทวี",
  "กฟส.สะบ้าย้อย",
  "กฟส.เทพา"
];

// Format PEA transformer number to ensure all digits are complete (YY-XXXXXX, 6 digits serial)
export function formatPeaTransformerCode(raw: string | number | undefined | null): string {
  if (!raw) return '';
  const str = String(raw).trim();
  
  // Format XX-XXXXXX (2-digit year, hyphen, 6-digit serial - complete 8 digits)
  if (/^\d{2}-\d{6}$/.test(str)) {
    return str;
  }
  
  // Hyphen format with fewer than 6 digits serial: XX-XXXXX -> pad serial to 6 digits
  const dashMatch = str.match(/^(\d{2})-(\d{1,5})$/);
  if (dashMatch) {
    return `${dashMatch[1]}-${dashMatch[2].padStart(6, '0')}`;
  }
  
  // Pure digits: 7 digits (2 digits year + 5 digits serial, e.g. 5001294) -> pad serial to 6 digits: XX-XXXXXX
  if (/^\d{7}$/.test(str)) {
    return `${str.slice(0, 2)}-${str.slice(2).padStart(6, '0')}`;
  }
  
  // Pure digits: 8 digits (2 digits year + 6 digits serial, e.g. 50001294, 63003398) -> XX-XXXXXX
  if (/^\d{8}$/.test(str)) {
    return `${str.slice(0, 2)}-${str.slice(2)}`;
  }

  // Pure digits: 1 to 6 digits serial without year prefix -> pad to 6 digits with default prefix
  if (/^\d{1,6}$/.test(str)) {
    return `50-${str.padStart(6, '0')}`;
  }

  return str;
}

// Generate robust realistic transformer dataset across 34 offices
export function generateInitialTransformers(): TransformerRecord[] {
  const list: TransformerRecord[] = [];
  let idCounter = 1;

  // หม้อแปลงตัวอย่างตามเกณฑ์: PEA หม้อแปลง 56-014418 พิกัดเดิม 100 kVA โหลดปีปัจจุบัน 15.08%
  list.push({
    id: 'tr_sample_56_014418',
    trCode: '56-014418',
    feaOffice: 'กฟจ.สงขลา',
    feeder: 'F-1 (สงขลา)',
    ratedKva: 100,
    year1LoadPct: 16.5, // โหลด 2 ปีก่อน (%)
    year2LoadPct: 18.2, // โหลด 1 ปีก่อน (%)
    year3LoadPct: 15.08, // โหลดปีปัจจุบัน (%)
    location: 'จุดติดตั้ง ถ.เพชรเกษม หน้า สนง. กฟจ.สงขลา',
    installationDate: '2020-03-15',
    status: 'Target Rightsizing'
  });

  PEA_OFFICES.forEach((office, officeIdx) => {
    // Generate 15-25 transformers per office
    const count = 18;
    for (let i = 1; i <= count; i++) {
      const ratings = [50, 100, 160, 250, 315, 400, 500];
      // Weighted towards 100, 160, 250 for testing
      const ratingWeights = [10, 30, 25, 20, 5, 5, 5];
      let randW = Math.random() * 100;
      let cumulative = 0;
      let ratedKva = 100;
      for (let w = 0; w < ratings.length; w++) {
        cumulative += ratingWeights[w];
        if (randW <= cumulative) {
          ratedKva = ratings[w];
          break;
        }
      }

      // Determine if this one should meet criteria (% load < 30% for 3 consecutive years for 100, 160, 250)
      const isTargetCandidate = [100, 160, 250].includes(ratedKva) && Math.random() < 0.38;

      let y1: number, y2: number, y3: number;
      if (isTargetCandidate) {
        y1 = Number((12 + Math.random() * 16).toFixed(2)); // 12% - 28%
        y2 = Number((10 + Math.random() * 18).toFixed(2)); // 10% - 28%
        y3 = Number((12 + Math.random() * 16).toFixed(2)); // 12% - 28%
      } else {
        // Generate some transformers with load >80% (81% - 94%) for TR Matching
        const isOverloadedCandidate = (i === 3 || i === 9 || i === 12);
        if (isOverloadedCandidate) {
          y1 = Number((65 + Math.random() * 20).toFixed(2));
          y2 = Number((72 + Math.random() * 18).toFixed(2));
          y3 = Number((81 + Math.random() * 13).toFixed(2)); // 81% - 94%
        } else {
          y1 = Number((25 + Math.random() * 45).toFixed(2));
          y2 = Number((28 + Math.random() * 45).toFixed(2));
          y3 = Number((30 + Math.random() * 45).toFixed(2));
        }
      }

      // PEA หม้อแปลง: รูปแบบหมายเลขทรัพย์สินหม้อแปลง PEA ให้ครบทุกหลัก (เช่น 50-001294, 52-004821, 63-003398 รวม 8 หลัก: ปี 2 หลัก - เลขลำดับ 6 หลัก)
      const peaYear = 48 + ((officeIdx * 7 + i) % 18);
      const peaSerial = String(1000 + ((officeIdx * 29 + i * 17) % 89999)).padStart(6, '0');
      const trCode = `${peaYear}-${peaSerial}`;

      // Approximate coordinates around Southern Thailand (Yala, Pattani, Narathiwat, Songkhla, Satun, Phatthalung)
      const baseLat = 6.5 + (officeIdx % 6) * 0.3 + (i * 0.02) % 0.4;
      const baseLng = 100.5 + (officeIdx % 6) * 0.35 + (i * 0.025) % 0.5;
      
      list.push({
        id: `tr_${idCounter}`,
        trCode,
        feaOffice: office,
        feeder: `F-${(idCounter % 5) + 1} (${office.replace('กฟฟ. ', '')})`,
        ratedKva,
        year1LoadPct: y1,
        year2LoadPct: y2,
        year3LoadPct: y3,
        location: (i === 6)
          ? `โครงการเคหะชุมชน ซอย ${(officeIdx % 5) + 1}`
          : (i === 7)
            ? `ที่ดินจัดสรรสินทวี ซอย ${(officeIdx % 5) + 1}`
            : (i === 8
              ? `หมู่บ้านมณฑาทิพย์ ซอย ${(officeIdx % 4) + 1}`
              : (i === 9
                ? `สถานีอัดประจุ ELEXa PEA จุดที่ ${(officeIdx % 3) + 1}`
                : (i === 10
                  ? `EV Charging Station ถ.เพชรเกษม`
                  : `จุดติดตั้ง หมู่ ${(idCounter % 10) + 1} ถ.เพชรเกษม/สายหลัก`))),
        latitude: Number(baseLat.toFixed(6)),
        longitude: Number(baseLng.toFixed(6)),
        installationDate: `20${15 + (idCounter % 7)}-0${(idCounter % 9) + 1}-15`,
        status: (isTargetCandidate && i !== 6 && i !== 7 && i !== 8 && i !== 9 && i !== 10) ? 'Target Rightsizing' : 'Normal'
      });
      idCounter++;
    }
  });

  return list;
}

// Evaluation helper for Rightsizing rules
export interface RightsizingEvaluation {
  isEligible: boolean;
  currentKva: number;
  recommendedKva: number | null;
  newLoadPct: number | null;
  currentLoadKva: number | null;
  calculationStep?: string;
  reason: string;
}

export function evaluateTransformerRightsizing(tr: TransformerRecord): RightsizingEvaluation {
  // เงื่อนไขยกเว้น: ไม่เอาหม้อแปลงที่สถานที่ติดตั้ง มีคำว่า "จัดสรร", "หมู่บ้าน", "โครงการ", "สถานีอัดประจุ" หรือ "EV" เข้ามาคำนวณในระบบนี้
  if (tr.location) {
    const locUpper = tr.location.toUpperCase();
    const isExcluded = tr.location.includes('จัดสรร') ||
                       tr.location.includes('หมู่บ้าน') ||
                       tr.location.includes('โครงการ') ||
                       tr.location.includes('สถานีอัดประจุ') ||
                       locUpper.includes('EV');

    if (isExcluded) {
      const matchedTerms: string[] = [];
      if (tr.location.includes('จัดสรร')) matchedTerms.push('จัดสรร');
      if (tr.location.includes('หมู่บ้าน')) matchedTerms.push('หมู่บ้าน');
      if (tr.location.includes('โครงการ')) matchedTerms.push('โครงการ');
      if (tr.location.includes('สถานีอัดประจุ')) matchedTerms.push('สถานีอัดประจุ');
      if (locUpper.includes('EV')) matchedTerms.push('EV');
      return {
        isEligible: false,
        currentKva: tr.ratedKva,
        recommendedKva: null,
        newLoadPct: null,
        currentLoadKva: null,
        reason: `ไม่นำมาคำนวณในระบบ: สถานที่ติดตั้งมีคำว่า "${matchedTerms.join(' / ')}" (${tr.location}) ได้รับการยกเว้น`
      };
    }
  }

  const targetRatings = [100, 160, 250];
  if (!targetRatings.includes(tr.ratedKva)) {
    return {
      isEligible: false,
      currentKva: tr.ratedKva,
      recommendedKva: null,
      newLoadPct: null,
      currentLoadKva: null,
      reason: `พิกัด ${tr.ratedKva} kVA ไม่อยู่ในเงื่อนไขการพิจารณา (เฉพาะ 100, 160, 250 kVA)`
    };
  }

  // Condition: % load < 30% for 3 consecutive years
  if (tr.year1LoadPct >= 30 || tr.year2LoadPct >= 30 || tr.year3LoadPct >= 30) {
    return {
      isEligible: false,
      currentKva: tr.ratedKva,
      recommendedKva: null,
      newLoadPct: null,
      currentLoadKva: null,
      reason: `ไม่เข้าเกณฑ์: % โหลดบางปี >= 30% (โหลด 2 ปีก่อน: ${tr.year1LoadPct.toFixed(2)}%, โหลด 1 ปีก่อน: ${tr.year2LoadPct.toFixed(2)}%, โหลดปีปัจจุบัน: ${tr.year3LoadPct.toFixed(2)}%)`
    };
  }

  // คำนวณโหลดของปีปัจจุบัน (kVA):
  // ของหม้อแปลง 100 kVA ให้เอาโหลดปีปัจจุบัน หารด้วย 100 คูณด้วย 100 จะได้เป็นโหลด
  // สูตรทั่วไป: (โหลดปีปัจจุบัน / 100) * พิกัดเดิม
  // เช่น PEA หม้อแปลง 56-014418 พิกัดเดิม 100 kVA โหลดปีปัจจุบัน 15.08%
  // -> (15.08 / 100) * 100 = 15.08 kVA
  const currentLoadKva = Number(((tr.year3LoadPct / 100) * tr.ratedKva).toFixed(2));

  // กำหนดพิกัดหม้อแปลงที่แนะนำสำหรับลดขนาด:
  // - 100 kVA -> แนะนำลดเป็น 50 kVA
  // - 160 kVA -> แนะนำลดเป็น 50 kVA หรือ 100 kVA (ขนาดที่ได้ % โหลดใหม่ < 50%)
  // - 250 kVA -> แนะนำลดเป็น 50 kVA, 100 kVA หรือ 160 kVA (ขนาดที่ได้ % โหลดใหม่ < 50%)
  let candidateSizes: number[] = [];
  if (tr.ratedKva === 100) {
    candidateSizes = [50];
  } else if (tr.ratedKva === 160) {
    candidateSizes = [50, 100];
  } else if (tr.ratedKva === 250) {
    candidateSizes = [50, 100, 160];
  }

  let bestRecommendedKva: number | null = null;
  let bestNewLoadPct: number | null = null;

  for (const candidate of candidateSizes) {
    // เอาโหลดหารด้วยพิกัดหม้อแปลงที่แนะนำ คูณด้วย 100
    // เช่น 15.08 / 50 * 100 = 30.16%
    const newLoad = (currentLoadKva / candidate) * 100;
    if (newLoad < 50) {
      bestRecommendedKva = candidate;
      bestNewLoadPct = Number(newLoad.toFixed(2));
      break;
    }
  }

  if (bestRecommendedKva === null || bestNewLoadPct === null) {
    return {
      isEligible: false,
      currentKva: tr.ratedKva,
      recommendedKva: null,
      newLoadPct: null,
      currentLoadKva,
      reason: `โหลด 3 ปีย้อนหลัง < 30% แต่เมื่อคำนวณ % โหลดใหม่จากโหลดปีปัจจุบัน (${tr.year3LoadPct.toFixed(2)}%) แล้วมีค่า >= 50%`
    };
  }

  const calcExplanation = tr.ratedKva === 100
    ? `(${tr.year3LoadPct.toFixed(2)} ÷ 100) × 100 = ${currentLoadKva.toFixed(2)} kVA จากนั้นนำ (${currentLoadKva.toFixed(2)} ÷ ${bestRecommendedKva}) × 100 = ${bestNewLoadPct.toFixed(2)}% (< 50%)`
    : `(${tr.year3LoadPct.toFixed(2)} ÷ 100) × ${tr.ratedKva} = ${currentLoadKva.toFixed(2)} kVA จากนั้นนำ (${currentLoadKva.toFixed(2)} ÷ ${bestRecommendedKva}) × 100 = ${bestNewLoadPct.toFixed(2)}% (< 50%)`;

  return {
    isEligible: true,
    currentKva: tr.ratedKva,
    recommendedKva: bestRecommendedKva,
    newLoadPct: bestNewLoadPct,
    currentLoadKva,
    calculationStep: calcExplanation,
    reason: `ผ่านเกณฑ์: โหลด 3 ปีย้อนหลัง < 30% ต่อเนื่อง และ % โหลดใหม่เมื่อลดเป็น ${bestRecommendedKva} kVA คำนวณได้ ${bestNewLoadPct.toFixed(2)}% (< 50%) จากโหลดปีปัจจุบัน (${tr.year3LoadPct.toFixed(2)}%)`
  };
}
