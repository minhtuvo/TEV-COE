/**
 * PMP (Preventive Maintenance Program / Precision Maintenance Procedures)
 * Evaluation & Real-Time Color Highlighting Service for Electrical Field Data Entry.
 * 
 * Standards referenced:
 * - ANSI/NETA ATS-2025 (Standard for Acceptance Testing Specifications)
 * - ANSI/NETA MTS-2023 (Standard for Maintenance Testing Specifications)
 * - IEEE Std 43, IEEE Std C57.104, IEEE Std C57.12.90, ASTM D1816, ASTM D1533
 */

export interface PmpStandardRule {
  key: string;
  nameVi: string;
  nameEn: string;
  rule: 'min' | 'max'; // 'min' means value must be >= threshold (smaller is FAIL); 'max' means value must be <= threshold (larger is FAIL)
  threshold: number;
  warningMarginPct?: number; // e.g. 15% margin before failing
  unit: string;
  standardCode: string;
  standardText: string;
  criticalDesc: string;
  healthyDesc: string;
}

export const PMP_STANDARD_RULES: Record<string, PmpStandardRule> = {
  // 1. Dielectric Breakdown Voltage (Dầu biến áp)
  dielectricBreakdown: {
    key: 'dielectricBreakdown',
    nameVi: 'Điện áp đánh thủng dầu D1816 (1mm)',
    nameEn: 'Dielectric Breakdown Voltage (ASTM D1816 1mm)',
    rule: 'min',
    threshold: 25.0,
    warningMarginPct: 15, // 25.0 to 28.75 is warning
    unit: 'kV',
    standardCode: 'NETA ATS Table 100.4.1 / PMP Oil',
    standardText: '≥ 25.0 kV (Bắt buộc theo chuẩn PMP & NETA)',
    criticalDesc: 'NHỎ HƠN TIÊU CHUẨN PMP: Dầu suy giảm độ bền điện môi, nguy cơ phóng điện vầng quang và chọc thủng cách điện.',
    healthyDesc: 'ĐẠT TIÊU CHUẨN PMP: Độ bền điện môi dầu cách điện đạt ngưỡng an toàn vận hành.'
  },

  // 2. Insulation Resistance - Medium Voltage (Cao thế / Trung thế)
  insulationResistanceMv: {
    key: 'insulationResistanceMv',
    nameVi: 'Điện trở cách điện IR Trung/Cao thế (1 min)',
    nameEn: 'Insulation Resistance MV/HV (1 min)',
    rule: 'min',
    threshold: 1000.0,
    warningMarginPct: 20, // 1000 to 1200 is warning
    unit: 'MΩ',
    standardCode: 'NETA ATS Table 100.5 / Table 100.1',
    standardText: '≥ 1,000 MΩ (1 GΩ)',
    criticalDesc: 'NHỎ HƠN TIÊU CHUẨN PMP: Điện trở cách điện suy giảm nặng nề, nguy cơ phóng điện đánh thủng xuống đất.',
    healthyDesc: 'ĐẠT TIÊU CHUẨN PMP: Cách điện khô ráo, an toàn đóng điện.'
  },

  // 3. Insulation Resistance - Low Voltage (Hạ thế 400V/480V)
  insulationResistanceLv: {
    key: 'insulationResistanceLv',
    nameVi: 'Điện trở cách điện IR Hạ thế (1 min)',
    nameEn: 'Insulation Resistance LV (1 min)',
    rule: 'min',
    threshold: 100.0,
    warningMarginPct: 20,
    unit: 'MΩ',
    standardCode: 'NETA ATS Table 100.1 / PMP LV',
    standardText: '≥ 100.0 MΩ (Tối thiểu ≥ 25 MΩ)',
    criticalDesc: 'NHỎ HƠN TIÊU CHUẨN PMP: Cách điện hạ thế bị ẩm ướt hoặc xuống cấp.',
    healthyDesc: 'ĐẠT TIÊU CHUẨN PMP: Cách điện hạ thế an toàn.'
  },

  // 4. Control Circuit Insulation Resistance (Mạch điều khiển nhị thứ)
  controlWiringIr: {
    key: 'controlWiringIr',
    nameVi: 'Điện trở cách điện mạch điều khiển nhị thứ',
    nameEn: 'Control Wiring Insulation Resistance',
    rule: 'min',
    threshold: 2.0,
    warningMarginPct: 25,
    unit: 'MΩ',
    standardCode: 'NETA ATS Sec 7.1.1.D.4',
    standardText: '≥ 2.0 MΩ (Bắt buộc nhị thứ)',
    criticalDesc: 'NHỎ HƠN TIÊU CHUẨN PMP: Nguy cơ chạm chập mạch nhị thứ, rơle mất nguồn hoặc tác động sai.',
    healthyDesc: 'ĐẠT TIÊU CHUẨN PMP: Mạch nhị thứ cách ly an toàn.'
  },

  // 5. Polarization Index (Chỉ số phân cực PI = R10m / R1m)
  polarizationIndex: {
    key: 'polarizationIndex',
    nameVi: 'Chỉ số phân cực PI (10 min / 1 min)',
    nameEn: 'Polarization Index (PI)',
    rule: 'min',
    threshold: 1.0,
    warningMarginPct: 25, // 1.0 to 1.25 is warning; > 1.5 is ideal
    unit: '',
    standardCode: 'NETA ATS Table 100.14 & IEEE 43',
    standardText: '≥ 1.0 (Bắt buộc; Khuyến nghị ≥ 1.5 - 2.0)',
    criticalDesc: 'NHỎ HƠN TIÊU CHUẨN PMP: Cách điện bị nhiễm ẩm nghiêm trọng hoặc bám bụi dẫn điện.',
    healthyDesc: 'ĐẠT TIÊU CHUẨN PMP: Chỉ số phân cực đạt yêu cầu.'
  },

  // 6. Dielectric Absorption Ratio (Chỉ số hấp thụ DAR = R60s / R30s)
  darValue: {
    key: 'darValue',
    nameVi: 'Chỉ số hấp thụ DAR (60s / 30s)',
    nameEn: 'Dielectric Absorption Ratio (DAR)',
    rule: 'min',
    threshold: 1.0,
    warningMarginPct: 25,
    unit: '',
    standardCode: 'NETA ATS Sec 7.2.1.1',
    standardText: '≥ 1.0 (Nguy hiểm nếu < 1.0, Tốt khi ≥ 1.25)',
    criticalDesc: 'NHỎ HƠN TIÊU CHUẨN PMP: Cảnh báo cách điện kém, bị nhiễm bẩn dẫn truyền.',
    healthyDesc: 'ĐẠT TIÊU CHUẨN PMP: Tỷ số hấp thụ đạt tiêu chuẩn.'
  },

  // 7. Battery Monoblock Voltage (Ắc quy 12V)
  batteryVoltage12V: {
    key: 'batteryVoltage12V',
    nameVi: 'Điện áp bình ắc quy Float 12V',
    nameEn: 'Battery Monoblock Float Voltage (12V)',
    rule: 'min',
    threshold: 13.2,
    warningMarginPct: 3,
    unit: 'V',
    standardCode: 'NETA ATS Sec 7.18 / IEEE 1188',
    standardText: '≥ 13.20 V (Dải chuẩn: 13.50V - 13.80V)',
    criticalDesc: 'NHỎ HƠN TIÊU CHUẨN PMP: Bình ắc quy bị sụt áp, hỏng cell hoặc sunfat hóa bản cực.',
    healthyDesc: 'ĐẠT TIÊU CHUẨN PMP: Điện áp float trong giới hạn tối ưu.'
  },

  // 8. SF6 Gas Pressure (Áp suất khí SF6)
  sf6Pressure: {
    key: 'sf6Pressure',
    nameVi: 'Áp suất khí SF6',
    nameEn: 'SF6 Gas Operating Pressure',
    rule: 'min',
    threshold: 5.0,
    warningMarginPct: 10,
    unit: 'bar',
    standardCode: 'NETA ATS Sec 7.5.4',
    standardText: '≥ 5.0 bar (Định mức 5.5 - 6.0 bar)',
    criticalDesc: 'NHỎ HƠN TIÊU CHUẨN PMP: Rò rỉ khí SF6, nguy cơ mất khả năng dập hồ quang máy cắt.',
    healthyDesc: 'ĐẠT TIÊU CHUẨN PMP: Áp suất khí SF6 đầy đủ.'
  },

  // 9. SF6 Gas Purity (Độ tinh khiết SF6)
  sf6Purity: {
    key: 'sf6Purity',
    nameVi: 'Độ tinh khiết khí SF6',
    nameEn: 'SF6 Gas Purity',
    rule: 'min',
    threshold: 97.0,
    warningMarginPct: 1.5,
    unit: '%',
    standardCode: 'NETA ATS Table 100.19',
    standardText: '≥ 97.0 %',
    criticalDesc: 'NHỎ HƠN TIÊU CHUẨN PMP: Khí SF6 bị nhiễm bẩn hoặc ẩm, giảm độ cách điện.',
    healthyDesc: 'ĐẠT TIÊU CHUẨN PMP: Độ tinh khiết khí SF6 đạt yêu cầu.'
  },

  // 10. Water Content in Oil (Độ ẩm trong dầu) - MAX rule
  waterContent: {
    key: 'waterContent',
    nameVi: 'Hàm lượng ẩm trong dầu ASTM D1533',
    nameEn: 'Moisture in Insulating Oil (ASTM D1533)',
    rule: 'max',
    threshold: 20.0,
    warningMarginPct: 20, // 16.0 to 20.0 is warning; > 20 is fail
    unit: 'ppm',
    standardCode: 'NETA ATS Table 100.4.1',
    standardText: '≤ 20.0 ppm',
    criticalDesc: 'VƯỢT TIÊU CHUẨN PMP: Dầu cách điện bị nhiễm ẩm nghiêm trọng.',
    healthyDesc: 'ĐẠT TIÊU CHUẨN PMP: Hàm lượng ẩm khô ráo, đạt chuẩn.'
  },

  // 11. Acetylene C2H2 DGA - MAX rule
  acetyleneC2h2: {
    key: 'acetyleneC2h2',
    nameVi: 'Khí Acetylene C2H2 trong dầu (DGA)',
    nameEn: 'Acetylene C2H2 Gas (IEEE C57.104)',
    rule: 'max',
    threshold: 1.0,
    warningMarginPct: 30,
    unit: 'ppm',
    standardCode: 'IEEE C57.104 & NETA ATS Sec 7.2.2',
    standardText: '≤ 1.0 ppm (Hồ quang điện nhiệt độ cao)',
    criticalDesc: 'VƯỢT TIÊU CHUẨN PMP: Xuất hiện hồ quang điện nguy hiểm bên trong máy biến áp.',
    healthyDesc: 'ĐẠT TIÊU CHUẨN PMP: Không có phóng điện hồ quang.'
  },

  // 12. Turns Ratio Error - MAX rule
  turnsRatioError: {
    key: 'turnsRatioError',
    nameVi: 'Sai số Tỷ số Biến áp Turns-Ratio',
    nameEn: 'Turns Ratio Deviation',
    rule: 'max',
    threshold: 0.5,
    warningMarginPct: 20,
    unit: '%',
    standardCode: 'IEEE C57.12.90 & NETA Sec 7.2.2.2',
    standardText: '≤ 0.5 %',
    criticalDesc: 'VƯỢT TIÊU CHUẨN PMP: Sai số tỷ số vòng dây vượt mức, nguy cơ chạm chập vòng dây hoặc nấc tap lỗi.',
    healthyDesc: 'ĐẠT TIÊU CHUẨN PMP: Tỷ số biến áp chuẩn xác.'
  },

  // 13. Winding Power Factor at 20°C - MAX rule
  windingPowerFactor: {
    key: 'windingPowerFactor',
    nameVi: 'Tổn hao điện môi Cuộn dây (Power Factor @ 20°C)',
    nameEn: 'Winding Insulation Power Factor at 20°C',
    rule: 'max',
    threshold: 0.5,
    warningMarginPct: 20,
    unit: '%',
    standardCode: 'NETA ATS Table 100.3',
    standardText: '≤ 0.50 % (Tối đa 1.0%)',
    criticalDesc: 'VƯỢT TIÊU CHUẨN PMP: Tổn hao điện môi cuộn dây tăng cao, già hóa cách điện.',
    healthyDesc: 'ĐẠT TIÊU CHUẨN PMP: Tổn hao điện môi cách điện thấp và ổn định.'
  },

  // 14. Bolted Resistance Deviation - MAX rule
  boltedDeviation: {
    key: 'boltedDeviation',
    nameVi: 'Độ lệch điện trở bu-lông 3 pha (DLRO)',
    nameEn: 'Bolted Joint Resistance Deviation',
    rule: 'max',
    threshold: 50.0,
    warningMarginPct: 20,
    unit: '%',
    standardCode: 'NETA ATS Table 100.12',
    standardText: '≤ 50.0 % (So với mối nối nhỏ nhất)',
    criticalDesc: 'VƯỢT TIÊU CHUẨN PMP: Mối nối tiếp xúc kém, lỏng bu-lông, nguy cơ phát nhiệt cháy mối nối.',
    healthyDesc: 'ĐẠT TIÊU CHUẨN PMP: Tiếp xúc bu-lông đồng đều 3 pha.'
  },

  // 15. Ground Resistance - MAX rule
  groundResistance: {
    key: 'groundResistance',
    nameVi: 'Điện trở tiếp địa trạm Fall-of-Potential',
    nameEn: 'Fall-of-Potential Ground Resistance (IEEE 81)',
    rule: 'max',
    threshold: 1.0,
    warningMarginPct: 20,
    unit: 'Ω',
    standardCode: 'NETA ATS Sec 7.13 & IEEE 81',
    standardText: '≤ 1.0 Ω (Trạm biến áp) / ≤ 5.0 Ω (Nhà máy)',
    criticalDesc: 'VƯỢT TIÊU CHUẨN PMP: Điện trở nối đất quá cao, nguy cơ điện áp bước và chạm nguy hiểm tính mạng.',
    healthyDesc: 'ĐẠT TIÊU CHUẨN PMP: Lưới nối đất trạm đạt chuẩn.'
  }
};

export interface PmpEvaluationResult {
  status: 'pass' | 'warning' | 'fail';
  isBelowStandard: boolean;
  value: number;
  threshold: number;
  unit: string;
  rule: 'min' | 'max';
  standardText: string;
  standardCode: string;
  badgeLabel: string;
  badgeBg: string;
  badgeText: string;
  inputBg: string;
  inputBorder: string;
  inputText: string;
  messageVi: string;
}

/**
 * Evaluates an input value against PMP standards.
 * Triggers vivid color highlight when value is SMALLER than standard (for 'min' rules)
 * or exceeds the safe threshold (for 'max' rules).
 */
export function evaluatePmpValue(
  ruleKeyOrName: string,
  rawVal: any,
  overrideMinThreshold?: number
): PmpEvaluationResult {
  const num = typeof rawVal === 'number'
    ? rawVal
    : parseFloat(String(rawVal).replace(/[^0-9.,-]/g, '').replace(',', '.'));

  const isNumeric = !isNaN(num) && rawVal !== '' && rawVal !== null && rawVal !== undefined;
  const currentVal = isNumeric ? num : 0;

  // Find standard rule by key or loose match
  const rule = PMP_STANDARD_RULES[ruleKeyOrName] || {
    key: ruleKeyOrName,
    nameVi: ruleKeyOrName,
    nameEn: ruleKeyOrName,
    rule: 'min',
    threshold: overrideMinThreshold ?? 1000,
    warningMarginPct: 15,
    unit: '',
    standardCode: 'Tiêu chuẩn PMP / NETA',
    standardText: `≥ ${overrideMinThreshold ?? 1000}`,
    criticalDesc: 'Giá trị nhỏ hơn ngưỡng tiêu chuẩn PMP.',
    healthyDesc: 'Giá trị đạt tiêu chuẩn PMP.'
  };

  const threshold = overrideMinThreshold ?? rule.threshold;
  const marginPct = rule.warningMarginPct ?? 15;

  let status: 'pass' | 'warning' | 'fail' = 'pass';
  let isBelowStandard = false;

  if (rule.rule === 'min') {
    if (currentVal < threshold) {
      status = 'fail';
      isBelowStandard = true;
    } else if (currentVal < threshold * (1 + marginPct / 100)) {
      status = 'warning';
      isBelowStandard = false;
    } else {
      status = 'pass';
      isBelowStandard = false;
    }
  } else {
    // 'max' rule: value must not exceed threshold
    if (currentVal > threshold) {
      status = 'fail';
      isBelowStandard = false;
    } else if (currentVal > threshold * (1 - marginPct / 100)) {
      status = 'warning';
      isBelowStandard = false;
    } else {
      status = 'pass';
      isBelowStandard = false;
    }
  }

  if (status === 'fail') {
    return {
      status: 'fail',
      isBelowStandard,
      value: currentVal,
      threshold,
      unit: rule.unit,
      rule: rule.rule,
      standardText: rule.standardText,
      standardCode: rule.standardCode,
      badgeLabel: isBelowStandard
        ? '⚠️ FAIL PMP: NHỎ HƠN TIÊU CHUẨN'
        : '⚠️ FAIL PMP: VƯỢT NGƯỠNG TIÊU CHUẨN',
      badgeBg: 'bg-rose-600',
      badgeText: 'text-white',
      inputBg: 'bg-rose-50/90 dark:bg-rose-950/40',
      inputBorder: 'border-2 border-rose-500 ring-2 ring-rose-200 dark:ring-rose-900',
      inputText: 'text-rose-950 dark:text-rose-200 font-black',
      messageVi: isBelowStandard
        ? `[TIÊU CHUẨN PMP: ${rule.standardText}] Giá trị hiện tại (${currentVal} ${rule.unit}) NHỎ HƠN tiêu chuẩn tối thiểu! ${rule.criticalDesc}`
        : `[TIÊU CHUẨN PMP: ${rule.standardText}] Giá trị hiện tại (${currentVal} ${rule.unit}) VƯỢT QUÁ ngưỡng an toàn! ${rule.criticalDesc}`
    };
  }

  if (status === 'warning') {
    return {
      status: 'warning',
      isBelowStandard: false,
      value: currentVal,
      threshold,
      unit: rule.unit,
      rule: rule.rule,
      standardText: rule.standardText,
      standardCode: rule.standardCode,
      badgeLabel: '⚡ CẢNH BÁO PMP: TIỆM CẬN NGƯỠNG',
      badgeBg: 'bg-amber-500',
      badgeText: 'text-slate-950',
      inputBg: 'bg-amber-50/90 dark:bg-amber-950/30',
      inputBorder: 'border-2 border-amber-500 ring-2 ring-amber-200 dark:ring-amber-900',
      inputText: 'text-amber-950 dark:text-amber-200 font-bold',
      messageVi: `[TIÊU CHUẨN PMP: ${rule.standardText}] Giá trị hiện tại (${currentVal} ${rule.unit}) tiệm cận ngưỡng giới hạn. Cần đưa vào kế hoạch giám sát theo dõi.`
    };
  }

  return {
    status: 'pass',
    isBelowStandard: false,
    value: currentVal,
    threshold,
    unit: rule.unit,
    rule: rule.rule,
    standardText: rule.standardText,
    standardCode: rule.standardCode,
    badgeLabel: '✓ ĐẠT TIÊU CHUẨN PMP',
    badgeBg: 'bg-emerald-600',
    badgeText: 'text-white',
    inputBg: 'bg-emerald-50/40 dark:bg-emerald-950/20',
    inputBorder: 'border border-emerald-400 focus:border-emerald-600 focus:ring-2 focus:ring-emerald-200',
    inputText: 'text-emerald-950 dark:text-emerald-200 font-bold',
    messageVi: `[TIÊU CHUẨN PMP: ${rule.standardText}] Giá trị hiện tại (${currentVal} ${rule.unit}) an toàn và hoàn toàn tuân thủ.`
  };
}
