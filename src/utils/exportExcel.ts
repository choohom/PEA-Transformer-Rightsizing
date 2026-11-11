import { TransformerRecord, evaluateTransformerRightsizing } from '../data/mockTransformers';

export function exportTransformersToExcel(transformers: TransformerRecord[], filename = 'PEA_Transformer_Rightsizing_Report.csv') {
  const headers = [
    'ลำดับ',
    'PEA หม้อแปลง',
    'กฟฟ. (คอลัมน์ AG)',
    'สถานที่ติดตั้ง (คอลัมน์ B)',
    'ฟีดเดอร์ (Feeder)',
    'พิกัดเดิม (kVA) (คอลัมน์ D)',
    'โหลด 2 ปีก่อน (%) (คอลัมน์ E)',
    'โหลด 1 ปีก่อน (%) (คอลัมน์ E)',
    'โหลดปีปัจจุบัน (%) (คอลัมน์ E)',
    'สถานะเกณฑ์ (<30% ทั้ง 3 ปีต่อเนื่อง)',
    'พิกัดแนะนำปรับลด (kVA)',
    '% โหลดใหม่ (< 50%)',
    'สถานะแผนงาน',
    'ผลการวิเคราะห์ Rightsizing'
  ];

  const rows = transformers.map((tr, idx) => {
    const evalRes = evaluateTransformerRightsizing(tr);
    return [
      idx + 1,
      `"${tr.trCode}"`,
      `"${tr.feaOffice}"`,
      `"${tr.location || tr.feaOffice}"`,
      `"${tr.feeder}"`,
      tr.ratedKva,
      tr.year1LoadPct.toFixed(2),
      tr.year2LoadPct.toFixed(2),
      tr.year3LoadPct.toFixed(2),
      `"${evalRes.isEligible ? 'ผ่านเกณฑ์ชี้เป้า' : 'ไม่เข้าเกณฑ์'}"`,
      evalRes.recommendedKva || '-',
      evalRes.newLoadPct !== null ? evalRes.newLoadPct.toFixed(2) : '-',
      `"${tr.status}"`,
      `"${evalRes.reason.replace(/"/g, '""')}"`
    ];
  });

  const csvContent = '\uFEFF' + [headers.join(','), ...rows.map(e => e.join(','))].join('\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', filename);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

export function exportTRMatchingToExcel(matches: any[], filename = 'TR_Matching_Report.csv') {
  const headers = [
    'ลำดับ',
    'สังกัด กฟฟ.',
    'PEA หม้อแปลง ต้นทาง (Rightsizing)',
    'พิกัดเดิมต้นทาง (kVA)',
    'โหลดปัจจุบันต้นทาง (%)',
    'PEA หม้อแปลง ปลายทาง (โหลดเกิน)',
    'พิกัดเดิมปลายทาง (kVA)',
    'โหลดปัจจุบันปลายทาง (%)',
    '% โหลดใหม่หลังจับคู่',
    'ระยะห่าง (km)'
  ];

  const rows = matches.map((m, idx) => {
    return [
      idx + 1,
      `"${m.office}"`,
      `"${m.sourceTr.trCode}"`,
      m.sourceTr.ratedKva,
      m.sourceTr.year3LoadPct.toFixed(2),
      `"${m.targetTr.trCode}"`,
      m.targetTr.ratedKva,
      m.targetTr.year3LoadPct.toFixed(2),
      m.targetNewLoadPct.toFixed(2),
      m.distanceKm !== null ? m.distanceKm : '-'
    ];
  });

  const csvContent = '\uFEFF' + [headers.join(','), ...rows.map(e => e.join(','))].join('\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', filename);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
