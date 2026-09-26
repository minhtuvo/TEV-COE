/**
 * DGA Diagnostic Engine
 * According to IEEE C57.104, IEC 60599, CIGRE TB 771, and Duval Methodology.
 *
 * CRITICAL BUSINESS RULE:
 * IF (Condition == 1 AND Gassing_Rate == Normal):
 *   -> Set Status = "NORMAL / HEALTHY"
 *   -> SKIP fault diagnostic methods (Duval, Rogers, etc.) to prevent false alarms from low-concentration noise.
 *   -> Output recommendation: "Continue normal operation and standard annual sampling."
 */

export interface DgaInputGases {
  h2: number;
  ch4: number;
  c2h6: number;
  c2h4: number;
  c2h2: number;
  co: number;
  co2: number;
  o2?: number;
  n2?: number;
  // Additional optional physical/chemical properties
  moisture?: number;
  bdStrength?: number;
  acidity?: number;
  ffa?: number;
  estDp?: number;
  age?: number;
  loadFactor?: number;
}

export interface DgaPreviousSample {
  h2: number;
  ch4: number;
  c2h6: number;
  c2h4: number;
  c2h2: number;
  co: number;
  co2?: number;
  deltaMonths: number; // Time interval Δt in months
}

export interface GassingRateAnalysis {
  hasPreviousData: boolean;
  deltaMonths: number;
  rates: {
    gas: string;
    rAbs: number; // ppm/month
    rRel: number; // %/month
    isHigh: boolean;
  }[];
  activeGassingDetected: boolean;
  status: 'Normal' | 'Warning' | 'High';
  summary: string;
}

export interface DgaDiagnosticOutput {
  // Step 1: Basic Calculations
  tdcg: number;

  // Step 2: Condition & Gassing Rate
  conditionLevel: 1 | 2 | 3 | 4;
  conditionName: 'Condition 1' | 'Condition 2' | 'Condition 3' | 'Condition 4';
  conditionTitleVi: string;
  isIndividualNormal: boolean;
  exceededGases: string[];
  gassingRate: GassingRateAnalysis;
  status: 'NORMAL / HEALTHY' | 'FAULT DETECTED / ACTIVE GASSING';
  faultDiagnosisSkipped: boolean;
  skipReason: string;

  // Step 3: Diagnostic Methods Results
  matrix: Record<string, string>;
  detailedResults: {
    co2_co: {
      ratio: number;
      applicable: boolean;
      status: string;
      description: string;
    };
    keyGas: {
      principalGas: string;
      fault: string;
      description: string;
    };
    rogers: {
      r1: number; // CH4 / H2
      r2: number; // C2H2 / C2H4
      r5: number; // C2H4 / C2H6
      code: string;
      fault: string;
      description: string;
    };
    iec60599: {
      r1: number;
      r2: number;
      r5: number;
      code: string;
      fault: string;
      description: string;
    };
    dornenburg: {
      applicable: boolean;
      r1: number;
      r2: number;
      r3: number;
      r4: number;
      code: string;
      fault: string;
      description: string;
    };
    duvalT1: {
      pCH4: number;
      pC2H4: number;
      pC2H2: number;
      zone: string;
      faultName: string;
    };
    duvalT4: {
      applicable: boolean;
      pH2: number;
      pCH4: number;
      pC2H6: number;
      zone: string;
      faultName: string;
      refinedReason: string;
    };
    duvalT5: {
      applicable: boolean;
      pCH4: number;
      pC2H4: number;
      pC2H6: number;
      zone: string;
      faultName: string;
      refinedReason: string;
    };
    duvalP1: {
      x: number;
      y: number;
      zone: string;
      faultName: string;
    };
    duvalP2: {
      applicable: boolean;
      x: number;
      y: number;
      zone: string;
      faultName: string;
    };
    etra: {
      result: string;
      description: string;
    };
  };

  // Step 4: Synthesis & Recommendations
  synthesis: {
    primaryFault: string;
    primaryFaultVi: string;
    confidence: 'Cao' | 'Trung bình' | 'Thấp' | 'Không có sự cố (Condition 1)';
    paperInvolvement: string;
    summaryVi: string;
  };
  recommendations: string[];
  nextSamplingInterval: string;
  recommendedTests: string[];

  // Formatted sections for Vietnamese Report Generation
  reportSections: {
    part1_TDCG_GassingRate: {
      title: string;
      tdcgValue: string;
      conditionText: string;
      rateText: string;
      decisionText: string;
    };
    part2_MethodDiagnostics: {
      title: string;
      methods: {
        name: string;
        result: string;
        note: string;
      }[];
    };
    part3_FinalSynthesis: {
      title: string;
      faultType: string;
      confidence: string;
      paperInvolvement: string;
      conclusions: string[];
    };
    part4_ActionRecommendations: {
      title: string;
      samplingInterval: string;
      precautions: string[];
      diagnosticTests: string[];
    };
  };

  // CNAIM & Health Index
  currentHealthScore: string;
  futureHealthScore: string;
  currentHI: string;
  futureHI: string;
  currentPOF: string;
  futurePOF: string;
  eolYears: number;
  dpEstimation: number;
  condition: string; // Tốt / Trung bình / Kém
  tdcgColor: string;
  recommendationColor: string;
  healthScoreData: { name: string; 'Health Score': number; fill: string }[];
  timestamp: string;
}

// Point in polygon helper for Duval Pentagons
function pointInPolygon(point: [number, number], vs: number[][]): boolean {
  let inside = false;
  for (let i = 0, j = vs.length - 1; i < vs.length; j = i++) {
    const xi = vs[i][0], yi = vs[i][1];
    const xj = vs[j][0], yj = vs[j][1];
    const intersect = ((yi > point[1]) !== (yj > point[1])) &&
      (point[0] < (xj - xi) * (point[1] - yi) / (yj - yi) + xi);
    if (intersect) inside = !inside;
  }
  return inside;
}

export const PENTAGON_1_REGIONS = [
  { name: 'PD', polygon: [[0, 33], [-1, 33], [-1, 24.5], [0, 24.5]] },
  { name: 'D1', polygon: [[0, 40], [38, 12], [32, -6.1], [4, 16], [0, 1.5]] },
  { name: 'D2', polygon: [[4, 16], [32, -6.1], [24.3, -30], [0, -3], [0, 1.5]] },
  { name: 'T3', polygon: [[0, -3], [24.3, -30], [23.5, -32.4], [1, -32], [-6, -4]] },
  { name: 'T2', polygon: [[-6, -4], [1, -32.4], [-22.5, -32.4]] },
  { name: 'T1', polygon: [[-6, -4], [-22.5, -32.4], [-23.5, -32.4], [-35, 3], [0, 1.5], [0, -3]] },
  { name: 'S', polygon: [[0, 1.5], [-35, 3.1], [-38, 12.4], [0, 40], [0, 33], [-1, 33], [-1, 24.5], [0, 24.5]] }
];

export const PENTAGON_2_REGIONS = [
  { name: 'PD', polygon: [[0, 33], [-1, 33], [-1, 24.5], [0, 24.5]] },
  { name: 'D1', polygon: [[0, 40], [38, 12], [32, -6.1], [4, 16], [0, 1.5]] },
  { name: 'D2', polygon: [[4, 16], [32, -6.1], [24.3, -30], [0, -3], [0, 1.5]] },
  { name: 'S', polygon: [[0, 1.5], [-35, 3.1], [-38, 12.4], [0, 40], [0, 33], [-1, 33], [-1, 24.5], [0, 24.5]] },
  { name: 'T3', polygon: [[0, -3], [24.3, -30], [23.5, -32.4], [2.5, -32.4], [-3.5, -3]] },
  { name: 'C', polygon: [[-3.5, -3], [2.5, -32.4], [-21.5, -32.4], [-11, -8]] },
  { name: 'O', polygon: [[-3.5, -3], [-11, -8], [-21.5, -32.4], [-23.5, -32.4], [-35, 3.1], [0, 1.5], [0, -3]] }
];

/**
 * Main Diagnostic Algorithm Function
 */
export function executeDgaDiagnostics(
  gases: DgaInputGases,
  prevSample?: DgaPreviousSample | null
): DgaDiagnosticOutput {
  const h2 = Math.max(0, Number(gases.h2) || 0);
  const ch4 = Math.max(0, Number(gases.ch4) || 0);
  const c2h6 = Math.max(0, Number(gases.c2h6) || 0);
  const c2h4 = Math.max(0, Number(gases.c2h4) || 0);
  const c2h2 = Math.max(0, Number(gases.c2h2) || 0);
  const co = Math.max(0, Number(gases.co) || 0);
  const co2 = Math.max(0, Number(gases.co2) || 0);

  // ----------------------------------------------------
  // STEP 1: BASIC CALCULATIONS
  // TDCG = H2 + CH4 + C2H6 + C2H4 + C2H2 + CO (No CO2)
  // ----------------------------------------------------
  const tdcg = parseFloat((h2 + ch4 + c2h6 + c2h4 + c2h2 + co).toFixed(1));

  // ----------------------------------------------------
  // STEP 2: CONDITION & GASSING RATE EVALUATION
  // 1. IEEE C57.104 Condition Thresholds
  // Condition 1 Limits: H2<100, CH4<120, C2H6<65, C2H4<50, C2H2<35, CO<350
  // ----------------------------------------------------
  const exceededGases: string[] = [];
  if (h2 >= 100) exceededGases.push(`H₂ (${h2} ≥ 100 ppm)`);
  if (ch4 >= 120) exceededGases.push(`CH₄ (${ch4} ≥ 120 ppm)`);
  if (c2h6 >= 65) exceededGases.push(`C₂H₆ (${c2h6} ≥ 65 ppm)`);
  if (c2h4 >= 50) exceededGases.push(`C₂H₄ (${c2h4} ≥ 50 ppm)`);
  if (c2h2 >= 35) exceededGases.push(`C₂H₂ (${c2h2} ≥ 35 ppm)`);
  if (co >= 350) exceededGases.push(`CO (${co} ≥ 350 ppm)`);

  const isIndividualNormal = exceededGases.length === 0;

  let conditionLevel: 1 | 2 | 3 | 4 = 1;
  let conditionName: 'Condition 1' | 'Condition 2' | 'Condition 3' | 'Condition 4' = 'Condition 1';
  let conditionTitleVi = 'Mức 1 (Bình thường / Khỏe mạnh)';
  let tdcgColor = 'text-emerald-700 bg-emerald-100 border-emerald-300';

  if (tdcg <= 720 && isIndividualNormal) {
    conditionLevel = 1;
    conditionName = 'Condition 1';
    conditionTitleVi = 'Mức 1 (Bình thường / Khỏe mạnh)';
    tdcgColor = 'text-emerald-700 bg-emerald-100 border-emerald-300';
  } else if (tdcg <= 1920) {
    conditionLevel = 2;
    conditionName = 'Condition 2';
    conditionTitleVi = 'Mức 2 (Cảnh báo 1 / Tiền cảnh báo)';
    tdcgColor = 'text-amber-700 bg-amber-100 border-amber-300';
  } else if (tdcg <= 4630) {
    conditionLevel = 3;
    conditionName = 'Condition 3';
    conditionTitleVi = 'Mức 3 (Cảnh báo 2 / Sự cố đang phát triển)';
    tdcgColor = 'text-orange-700 bg-orange-100 border-orange-300';
  } else {
    conditionLevel = 4;
    conditionName = 'Condition 4';
    conditionTitleVi = 'Mức 4 (Nguy hiểm / Cần hành động ngay)';
    tdcgColor = 'text-rose-700 bg-rose-100 border-rose-300';
  }

  // 2. Gassing Rate Analysis
  let gassingRate: GassingRateAnalysis = {
    hasPreviousData: false,
    deltaMonths: 0,
    rates: [],
    activeGassingDetected: false,
    status: 'Normal',
    summary: 'Không có dữ liệu mẫu trước đó để tính tốc độ sinh khí.'
  };

  if (prevSample && prevSample.deltaMonths > 0) {
    const dt = prevSample.deltaMonths;
    const gasCheck = [
      { name: 'H2', cur: h2, prev: prevSample.h2, isActiveCombustible: true },
      { name: 'CH4', cur: ch4, prev: prevSample.ch4, isActiveCombustible: false },
      { name: 'C2H6', cur: c2h6, prev: prevSample.c2h6, isActiveCombustible: false },
      { name: 'C2H4', cur: c2h4, prev: prevSample.c2h4, isActiveCombustible: true },
      { name: 'C2H2', cur: c2h2, prev: prevSample.c2h2, isActiveCombustible: true },
      { name: 'CO', cur: co, prev: prevSample.co, isActiveCombustible: false }
    ];

    let activeGassing = false;
    const rates = gasCheck.map(item => {
      const delta = item.cur - item.prev;
      const rAbs = parseFloat((delta / dt).toFixed(2));
      const rRel = item.prev > 0 ? parseFloat(((delta / item.prev) * 100 / dt).toFixed(2)) : 0;
      const isHigh = item.isActiveCombustible && (rAbs >= 10 || rRel >= 10);
      if (isHigh) activeGassing = true;
      return { gas: item.name, rAbs, rRel, isHigh };
    });

    gassingRate = {
      hasPreviousData: true,
      deltaMonths: dt,
      rates,
      activeGassingDetected: activeGassing,
      status: activeGassing ? 'High' : 'Normal',
      summary: activeGassing
        ? 'Phát hiện tốc độ sinh khí hoạt tính cao (R_abs ≥ 10 ppm/tháng hoặc R_rel ≥ 10%/tháng với H₂, C₂H₄ hoặc C₂H₂).'
        : 'Tốc độ sinh khí nằm trong giới hạn bình thường (< 10 ppm/tháng).'
    };
  }

  // 3. Flowchart Decision Rule:
  // IF (Condition == 1 AND Gassing_Rate == Normal):
  //    -> Set Status = "NORMAL / HEALTHY"
  //    -> SKIP fault diagnostic methods (Duval, Rogers, etc.)
  // IF (Condition >= 2 OR Gassing_Rate == Warning/High):
  //    -> Set Status = "FAULT DETECTED / ACTIVE GASSING"
  //    -> PROCEED immediately to STEP 3
  const isHealthyCondition1 = (conditionLevel === 1) && (!gassingRate.activeGassingDetected);
  const status: 'NORMAL / HEALTHY' | 'FAULT DETECTED / ACTIVE GASSING' = isHealthyCondition1
    ? 'NORMAL / HEALTHY'
    : 'FAULT DETECTED / ACTIVE GASSING';
  const faultDiagnosisSkipped = isHealthyCondition1;
  const skipReason = isHealthyCondition1
    ? 'TDCG ≤ 720 ppm và toàn bộ khí riêng lẻ dưới ngưỡng giới hạn theo IEEE C57.104. Thuật toán tự động bỏ qua phân tích lỗi (Duval, Rogers, IEC) nhằm loại trừ báo động giả phát sinh từ tạp âm nồng độ thấp.'
    : '';

  // ----------------------------------------------------
  // STEP 3: FAULT DIAGNOSTIC METHODS (Run in Parallel)
  // ----------------------------------------------------

  // 1. CO2/CO Ratio
  let co2_co_applicable = false;
  let co2_co_ratio = 0;
  let co2_co_status = 'ND';
  let co2_co_desc = 'Chưa áp dụng';

  if (!faultDiagnosisSkipped) {
    if (co > 500 && co2 > 5000) {
      co2_co_applicable = true;
      co2_co_ratio = parseFloat((co2 / co).toFixed(2));
      if (co2_co_ratio < 3) {
        co2_co_status = 'C';
        co2_co_desc = 'Thoái hóa nhiệt xenlulo nghiêm trọng / Carbon hóa giấy cách điện (CO₂/CO < 3)';
      } else if (co2_co_ratio <= 10) {
        co2_co_status = 'OK';
        co2_co_desc = 'Lão hóa giấy xenlulo ở mức bình thường (3 ≤ CO₂/CO ≤ 10)';
      } else {
        co2_co_status = 'O';
        co2_co_desc = 'Lão hóa nhiệt độ thấp hoặc oxy hóa dầu cách điện (CO₂/CO > 10)';
      }
    } else {
      co2_co_status = 'ND';
      co2_co_desc = 'Chưa thỏa mãn điều kiện kích hoạt: CO > 500 ppm và CO₂ > 5000 ppm';
    }
  } else {
    co2_co_status = 'OK';
    co2_co_desc = 'Bình thường (Condition 1)';
  }

  // 2. Key Gas Method
  let keyGas_principal = '';
  let keyGas_fault = 'Normal';
  let keyGas_desc = 'Bình thường';

  if (!faultDiagnosisSkipped) {
    const combustibleMap: Record<string, number> = { h2, ch4, c2h6, c2h4, c2h2, co };
    let maxVal = -1;
    for (const [g, val] of Object.entries(combustibleMap)) {
      if (val > maxVal) {
        maxVal = val;
        keyGas_principal = g;
      }
    }

    if (keyGas_principal === 'h2') {
      if (c2h2 > 0.1 * h2 && c2h2 >= 5) {
        keyGas_fault = 'Arcing';
        keyGas_desc = 'Phóng điện hồ quang năng lượng cao (H₂ + C₂H₂ chủ đạo)';
      } else {
        keyGas_fault = 'PD / Corona';
        keyGas_desc = 'Phóng điện cục bộ (PD / Corona, H₂ chủ đạo)';
      }
    } else if (keyGas_principal === 'c2h2') {
      keyGas_fault = 'Arcing';
      keyGas_desc = 'Phóng điện hồ quang năng lượng cao (C₂H₂ chủ đạo)';
    } else if (keyGas_principal === 'c2h4') {
      keyGas_fault = 'Overheating (Oil)';
      keyGas_desc = 'Quá nhiệt dầu nhiệt độ cao > 700°C (C₂H₄ chủ đạo)';
    } else if (keyGas_principal === 'co') {
      keyGas_fault = 'Overheating (Paper)';
      keyGas_desc = 'Quá nhiệt giấy cách điện xenlulo (CO chủ đạo)';
    } else if (keyGas_principal === 'ch4' || keyGas_principal === 'c2h6') {
      keyGas_fault = 'Low Temp Thermal';
      keyGas_desc = 'Quá nhiệt độ thấp trong dầu hoặc lõi từ (CH₄/C₂H₆ chủ đạo)';
    }
  } else {
    keyGas_principal = 'None';
    keyGas_fault = 'Normal';
    keyGas_desc = 'Bình thường (Condition 1 - Không có khí chủ đạo gây lỗi)';
  }

  // 3. Rogers Ratio Method
  // R1 = CH4/H2, R2 = C2H2/C2H4, R5 = C2H4/C2H6
  const r1_rogers = h2 > 0 ? parseFloat((ch4 / h2).toFixed(3)) : 0;
  const r2_rogers = c2h4 > 0 ? parseFloat((c2h2 / c2h4).toFixed(3)) : 0;
  const r5_rogers = c2h6 > 0 ? parseFloat((c2h4 / c2h6).toFixed(3)) : 0;

  let rogers_code = 'ND';
  let rogers_fault = 'No Diagnosis / Unidentified Ratio';
  let rogers_desc = 'Chưa xác định';

  if (!faultDiagnosisSkipped) {
    if (r2_rogers < 0.1 && r1_rogers < 0.1 && r5_rogers < 1.0) {
      rogers_code = 'OK';
      rogers_fault = 'Normal';
      rogers_desc = 'Vận hành bình thường (Normal)';
    } else if (r2_rogers < 0.1 && r1_rogers >= 1.0 && r5_rogers < 1.0) {
      rogers_code = 'T1';
      rogers_fault = 'Thermal fault < 300°C';
      rogers_desc = 'Quá nhiệt nhiệt độ thấp < 300°C (T1)';
    } else if (r2_rogers < 0.1 && r1_rogers >= 1.0 && r5_rogers >= 1.0 && r5_rogers < 4.0) {
      rogers_code = 'T2';
      rogers_fault = 'Thermal fault 300°C-700°C';
      rogers_desc = 'Quá nhiệt nhiệt độ trung bình 300°C - 700°C (T2)';
    } else if (r2_rogers < 0.1 && r1_rogers >= 1.0 && r5_rogers >= 4.0) {
      rogers_code = 'T3';
      rogers_fault = 'Thermal fault > 700°C';
      rogers_desc = 'Quá nhiệt nhiệt độ cao > 700°C (T3)';
    } else if (r2_rogers >= 1.0 && r1_rogers >= 0.1 && r1_rogers < 0.5 && r5_rogers >= 1.0) {
      rogers_code = 'D1';
      rogers_fault = 'Low energy discharge';
      rogers_desc = 'Phóng điện năng lượng thấp (D1)';
    } else {
      rogers_code = 'ND';
      rogers_fault = 'No Diagnosis / Unidentified Ratio';
      rogers_desc = 'Không thuộc mã chuẩn Rogers';
    }
  } else {
    rogers_code = 'OK';
    rogers_fault = 'Normal';
    rogers_desc = 'Bình thường (Condition 1)';
  }

  // 4. IEC 60599 Ratio Method
  // R2 = C2H2/C2H4, R1 = CH4/H2, R5 = C2H4/C2H6
  let iec_code = 'ND';
  let iec_fault = 'No Diagnosis / Unidentified Ratio';
  let iec_desc = 'Chưa xác định';

  if (!faultDiagnosisSkipped) {
    if (r2_rogers < 0.1 && r1_rogers < 0.1 && r5_rogers < 0.2) {
      iec_code = 'PD';
      iec_fault = 'Partial Discharge';
      iec_desc = 'Phóng điện cục bộ (PD)';
    } else if (r2_rogers > 1.0 && r1_rogers >= 0.1 && r1_rogers <= 0.5 && r5_rogers > 1.0) {
      iec_code = 'D1';
      iec_fault = 'Discharge low energy';
      iec_desc = 'Phóng điện năng lượng thấp (D1)';
    } else if (r2_rogers >= 0.6 && r2_rogers <= 2.5 && r1_rogers >= 0.1 && r1_rogers <= 1.0 && r5_rogers > 2.0) {
      iec_code = 'D2';
      iec_fault = 'Discharge high energy';
      iec_desc = 'Phóng điện năng lượng cao (D2)';
    } else if (r2_rogers < 0.1 && r1_rogers > 1.0 && r5_rogers < 1.0) {
      iec_code = 'T1';
      iec_fault = 'Thermal fault < 300°C';
      iec_desc = 'Quá nhiệt nhiệt độ thấp < 300°C (T1)';
    } else if (r2_rogers < 0.1 && r1_rogers > 1.0 && r5_rogers >= 1.0 && r5_rogers <= 4.0) {
      iec_code = 'T2';
      iec_fault = 'Thermal fault 300°C-700°C';
      iec_desc = 'Quá nhiệt nhiệt độ trung bình 300°C - 700°C (T2)';
    } else if (r2_rogers < 0.2 && r1_rogers > 1.0 && r5_rogers > 4.0) {
      iec_code = 'T3';
      iec_fault = 'Thermal fault > 700°C';
      iec_desc = 'Quá nhiệt nhiệt độ cao > 700°C (T3)';
    } else {
      iec_code = 'ND';
      iec_fault = 'No Diagnosis / Unidentified Ratio';
      iec_desc = 'Không khớp tỷ số IEC 60599';
    }
  } else {
    iec_code = 'OK';
    iec_fault = 'Normal';
    iec_desc = 'Bình thường (Condition 1)';
  }

  // 5. Dornenburg Ratios
  // L1 limits: H2>100, CH4>120, CO>350, C2H2>35, C2H4>50, C2H6>65
  const dornenburg_applicable = (h2 > 100 || ch4 > 120 || co > 350 || c2h2 > 35 || c2h4 > 50 || c2h6 > 65);
  const d_r1 = h2 > 0 ? parseFloat((ch4 / h2).toFixed(3)) : 0;
  const d_r2 = c2h4 > 0 ? parseFloat((c2h2 / c2h4).toFixed(3)) : 0;
  const d_r3 = ch4 > 0 ? parseFloat((c2h2 / ch4).toFixed(3)) : 0;
  const d_r4 = c2h2 > 0 ? parseFloat((c2h6 / c2h2).toFixed(3)) : 0;

  let dornenburg_code = 'ND';
  let dornenburg_fault = 'No Diagnosis';
  let dornenburg_desc = 'Chưa xác định';

  if (!faultDiagnosisSkipped) {
    if (dornenburg_applicable) {
      if (d_r1 > 1.0 && d_r2 < 0.75 && d_r3 < 0.3 && d_r4 > 0.4) {
        dornenburg_code = 'T1';
        dornenburg_fault = 'Thermal Fault';
        dornenburg_desc = 'Quá nhiệt nhiệt phân dầu (Thermal Fault)';
      } else if (d_r1 < 0.1 && d_r3 < 0.3 && d_r4 > 0.4) {
        dornenburg_code = 'PD';
        dornenburg_fault = 'Corona / PD';
        dornenburg_desc = 'Phóng điện vầng quang / Corona (PD)';
      } else if (d_r1 > 0.1 && d_r1 < 1.0 && d_r2 > 0.75 && d_r3 > 0.3 && d_r4 < 0.4) {
        dornenburg_code = 'D2';
        dornenburg_fault = 'Arcing';
        dornenburg_desc = 'Phóng điện hồ quang (Arcing)';
      } else {
        dornenburg_code = 'ND';
        dornenburg_fault = 'Unidentified Ratio';
        dornenburg_desc = 'Tỷ số nằm ngoài bảng nhận diện Dornenburg';
      }
    } else {
      dornenburg_code = 'OK';
      dornenburg_fault = 'Normal';
      dornenburg_desc = 'Các khí chưa vượt ngưỡng L1 (Dornenburg)';
    }
  } else {
    dornenburg_code = 'OK';
    dornenburg_fault = 'Normal';
    dornenburg_desc = 'Bình thường (Condition 1)';
  }

  // 6. Duval Triangle 1 (T1) & Refinements (T4, T5)
  // %CH4, %C2H4, %C2H2
  const sumT1 = ch4 + c2h4 + c2h2;
  let pCH4_t1 = 0, pC2H4_t1 = 0, pC2H2_t1 = 0;
  let t1_zone = 'OK';
  let t1_name = 'Normal';

  if (!faultDiagnosisSkipped && sumT1 > 0) {
    pCH4_t1 = parseFloat(((ch4 / sumT1) * 100).toFixed(1));
    pC2H4_t1 = parseFloat(((c2h4 / sumT1) * 100).toFixed(1));
    pC2H2_t1 = parseFloat(((c2h2 / sumT1) * 100).toFixed(1));

    if (pCH4_t1 >= 98) {
      t1_zone = 'PD';
      t1_name = 'Phóng điện cục bộ (Partial Discharge)';
    } else if (pC2H2_t1 >= 29 || (pC2H2_t1 >= 13 && pC2H4_t1 >= 23 && pC2H4_t1 < 50) || (pC2H2_t1 >= 15 && pC2H4_t1 >= 50)) {
      t1_zone = 'D2';
      t1_name = 'Phóng điện năng lượng cao (Discharge D2)';
    } else if (pC2H2_t1 >= 13 && pC2H2_t1 < 29 && pC2H4_t1 < 23) {
      t1_zone = 'D1';
      t1_name = 'Phóng điện năng lượng thấp (Discharge D1)';
    } else if (pC2H2_t1 < 15 && pC2H4_t1 >= 50) {
      t1_zone = 'T3';
      t1_name = 'Quá nhiệt nhiệt độ cao > 700°C (T3)';
    } else if (pC2H2_t1 < 4 && pC2H4_t1 >= 20 && pC2H4_t1 < 50) {
      t1_zone = 'T2';
      t1_name = 'Quá nhiệt nhiệt độ trung bình 300-700°C (T2)';
    } else if (pC2H2_t1 < 4 && pC2H4_t1 < 20) {
      t1_zone = 'T1';
      t1_name = 'Quá nhiệt nhiệt độ thấp < 300°C (T1)';
    } else {
      t1_zone = 'DT';
      t1_name = 'Hỗn hợp quá nhiệt và phóng điện (DT)';
    }
  } else {
    t1_zone = 'OK';
    t1_name = 'Bình thường (Condition 1)';
  }

  // Duval Triangle 4 Refinement
  // Run if T1 result is PD, T1, or T2 (%H2, %CH4, %C2H6)
  const sumT4 = h2 + ch4 + c2h6;
  let pH2_t4 = 0, pCH4_t4 = 0, pC2H6_t4 = 0;
  let t4_applicable = false;
  let t4_zone = 'ND';
  let t4_name = 'Chưa áp dụng';
  let t4_refinedReason = '';

  if (!faultDiagnosisSkipped && ['PD', 'T1', 'T2'].includes(t1_zone) && sumT4 > 0) {
    t4_applicable = true;
    pH2_t4 = parseFloat(((h2 / sumT4) * 100).toFixed(1));
    pCH4_t4 = parseFloat(((ch4 / sumT4) * 100).toFixed(1));
    pC2H6_t4 = parseFloat(((c2h6 / sumT4) * 100).toFixed(1));
    t4_refinedReason = `Kích hoạt tinh chỉnh do Triangle 1 chuẩn đoán vùng ${t1_zone}`;

    if (pCH4_t4 >= 2 && pCH4_t4 < 15 && pC2H6_t4 < 1) {
      t4_zone = 'PD';
      t4_name = 'Phóng điện cục bộ (Partial Discharge)';
    } else if (pH2_t4 < 9 && pC2H6_t4 >= 30) {
      t4_zone = 'O';
      t4_name = 'Quá nhiệt nhiệt độ thấp < 250°C (Overheating O)';
    } else if ((pCH4_t4 >= 36 && pC2H6_t4 >= 24) || (pH2_t4 < 15 && pC2H6_t4 >= 24 && pC2H6_t4 < 30)) {
      t4_zone = 'C';
      t4_name = 'Carbon hóa giấy cách điện (Paper Carbonization C)';
    } else if (pH2_t4 >= 9 && pC2H6_t4 >= 46) {
      t4_zone = 'ND';
      t4_name = 'Chưa xác định rõ';
    } else {
      t4_zone = 'S';
      t4_name = 'Lỗi giải phóng khí dầu (Stray Gassing S)';
    }
  } else {
    t4_zone = faultDiagnosisSkipped ? 'OK' : 'ND';
    t4_name = faultDiagnosisSkipped ? 'Bình thường (Condition 1)' : 'Không cần thiết tinh chỉnh T4';
  }

  // Duval Triangle 5 Refinement
  // Run if T1 result is T2 or T3 (%CH4, %C2H4, %C2H6)
  const sumT5 = ch4 + c2h4 + c2h6;
  let pCH4_t5 = 0, pC2H4_t5 = 0, pC2H6_t5 = 0;
  let t5_applicable = false;
  let t5_zone = 'ND';
  let t5_name = 'Chưa áp dụng';
  let t5_refinedReason = '';

  if (!faultDiagnosisSkipped && ['T2', 'T3'].includes(t1_zone) && sumT5 > 0) {
    t5_applicable = true;
    pCH4_t5 = parseFloat(((ch4 / sumT5) * 100).toFixed(1));
    pC2H4_t5 = parseFloat(((c2h4 / sumT5) * 100).toFixed(1));
    pC2H6_t5 = parseFloat(((c2h6 / sumT5) * 100).toFixed(1));
    t5_refinedReason = `Kích hoạt tinh chỉnh do Triangle 1 chuẩn đoán vùng ${t1_zone}`;

    if (pC2H4_t5 < 1 && pC2H6_t5 >= 2 && pC2H6_t5 < 14) {
      t5_zone = 'PD';
      t5_name = 'Phóng điện cục bộ (PD)';
    } else if ((pC2H4_t5 >= 1 && pC2H4_t5 < 10 && pC2H6_t5 >= 2 && pC2H6_t5 < 14) || (pC2H4_t5 < 1 && pC2H6_t5 < 2) || (pC2H4_t5 < 10 && pC2H6_t5 >= 54)) {
      t5_zone = 'O';
      t5_name = 'Quá nhiệt nhiệt độ thấp < 250°C (O)';
    } else if (pC2H4_t5 < 10 && pC2H6_t5 >= 14 && pC2H6_t5 < 54) {
      t5_zone = 'S';
      t5_name = 'Khí sinh tự nhiên trong dầu (Stray Gassing S)';
    } else if (pC2H4_t5 >= 10 && pC2H4_t5 < 35 && pC2H6_t5 < 12) {
      t5_zone = 'T2';
      t5_name = 'Quá nhiệt 300°C-700°C (T2)';
    } else if (
      (pC2H4_t5 >= 35 && pC2H6_t5 < 12) ||
      (pC2H4_t5 >= 50 && pC2H6_t5 >= 12 && pC2H6_t5 < 14) ||
      (pC2H4_t5 >= 70 && pC2H6_t5 >= 14) ||
      (pC2H4_t5 >= 35 && pC2H6_t5 >= 30)
    ) {
      t5_zone = 'T3-H';
      t5_name = 'Quá nhiệt trong dầu nhiệt độ cao > 700°C (T3-H)';
    } else if (
      (pC2H4_t5 >= 10 && pC2H4_t5 < 50 && pC2H6_t5 >= 12 && pC2H6_t5 < 14) ||
      (pC2H4_t5 >= 10 && pC2H4_t5 < 70 && pC2H6_t5 >= 14 && pC2H6_t5 < 30)
    ) {
      t5_zone = 'C';
      t5_name = 'Lỗi nhiệt liên quan đến giấy cách điện xenlulo (C)';
    } else {
      t5_zone = 'ND';
      t5_name = 'Chưa xác định';
    }
  } else {
    t5_zone = faultDiagnosisSkipped ? 'OK' : 'ND';
    t5_name = faultDiagnosisSkipped ? 'Bình thường (Condition 1)' : 'Không cần thiết tinh chỉnh T5';
  }

  // 7. Duval Pentagon 1 & Pentagon 2
  // Coordinates (%H2, %C2H6, %CH4, %C2H4, %C2H2)
  const sumP = h2 + c2h6 + ch4 + c2h4 + c2h2;
  let p1_zone = 'OK';
  let p1_name = 'Bình thường';
  let p2_zone = 'OK';
  let p2_name = 'Bình thường';
  let p2_applicable = false;
  let pentagonCoords = { x: 50, y: 50 };

  if (!faultDiagnosisSkipped && sumP > 0) {
    const f1 = h2 / sumP;
    const f2 = c2h6 / sumP;
    const f3 = ch4 / sumP;
    const f4 = c2h4 / sumP;
    const f5 = c2h2 / sumP;

    const mathX = f1 * 0 + f2 * (-38) + f3 * (-23.5) + f4 * 23.5 + f5 * 38;
    const mathY = f1 * 40 + f2 * 12.4 + f3 * (-32.4) + f4 * (-32.4) + f5 * 12.4;
    pentagonCoords = { x: 50 - mathX, y: 50 - mathY };

    const pt: [number, number] = [mathX, mathY];

    for (const z of PENTAGON_1_REGIONS) {
      if (pointInPolygon(pt, z.polygon)) {
        p1_zone = z.name;
        break;
      }
    }

    const p1Names: Record<string, string> = {
      PD: 'Phóng điện cục bộ (Partial Discharge)',
      D1: 'Phóng điện năng lượng thấp (Low Energy Discharge)',
      D2: 'Phóng điện năng lượng cao (High Energy Discharge)',
      T1: 'Quá nhiệt nhiệt độ thấp < 300°C',
      T2: 'Quá nhiệt nhiệt độ 300°C - 700°C',
      T3: 'Quá nhiệt nhiệt độ cao > 700°C',
      S: 'Hiện tượng sinh khí trong dầu (Stray Gassing)'
    };
    p1_name = p1Names[p1_zone] || p1_zone;

    // Run Pentagon 2 if P1 indicates thermal fault (T1, T2, T3) or Overheating
    if (['T1', 'T2', 'T3', 'S'].includes(p1_zone)) {
      p2_applicable = true;
      for (const z of PENTAGON_2_REGIONS) {
        if (pointInPolygon(pt, z.polygon)) {
          p2_zone = z.name;
          break;
        }
      }
      const p2Names: Record<string, string> = {
        PD: 'Phóng điện cục bộ (PD)',
        D1: 'Phóng điện năng lượng thấp (D1)',
        D2: 'Phóng điện năng lượng cao (D2)',
        S: 'Khí sinh tự nhiên trong dầu (S)',
        T3: 'Quá nhiệt dầu nhiệt độ cao (T3-H)',
        C: 'Lỗi nhiệt liên quan đến giấy cách điện (C)',
        O: 'Quá nhiệt nhẹ < 250°C (O)'
      };
      p2_name = p2Names[p2_zone] || p2_zone;
    } else {
      p2_zone = 'ND';
      p2_name = 'Không yêu cầu tinh chỉnh P2 (Không có lỗi nhiệt)';
    }
  } else {
    p1_zone = 'OK';
    p1_name = 'Bình thường (Condition 1)';
    p2_zone = 'OK';
    p2_name = 'Bình thường (Condition 1)';
  }

  // 8. ETRA Method
  let etra_res = 'Normal';
  let etra_desc = 'Bình thường';
  if (!faultDiagnosisSkipped) {
    if (tdcg < 720) {
      etra_res = 'Normal';
      etra_desc = 'Điều kiện bình thường (TDCG < 720 ppm)';
    } else if (tdcg < 1920) {
      etra_res = 'Condition 2';
      etra_desc = 'Cần giám sát tăng cường (Condition 2)';
    } else if (tdcg < 4630) {
      etra_res = 'Condition 3';
      etra_desc = 'Cảnh báo mức cao (Condition 3)';
    } else {
      etra_res = 'Condition 4';
      etra_desc = 'Mức nghiêm trọng (Condition 4)';
    }
  } else {
    etra_res = 'Normal';
    etra_desc = 'Bình thường (Condition 1)';
  }

  // Construct Matrix
  const matrix: Record<string, string> = {
    'Duval T1': t1_zone,
    'Duval T4': t4_zone,
    'Duval T5': t5_zone,
    'Duval P1': p1_zone,
    'Duval P2': p2_zone,
    'KeyGas': keyGas_fault,
    'IEEE (Rogers)': rogers_code,
    'IEC Ratio': iec_code,
    'IEEE (Dornenburg)': dornenburg_code,
    'CO2/CO': co2_co_status,
    'ETRA': etra_res
  };

  // ----------------------------------------------------
  // STEP 4: SYNTHESIS & RECOMMENDATIONS
  // ----------------------------------------------------
  let primaryFault = 'Normal / Healthy';
  let primaryFaultVi = 'Không có sự cố (Bình thường)';
  let confidence: 'Cao' | 'Trung bình' | 'Thấp' | 'Không có sự cố (Condition 1)' = 'Không có sự cố (Condition 1)';
  let paperInvolvement = 'Không phát hiện suy giảm cách điện giấy';
  let summaryVi = '';
  const recommendations: string[] = [];
  let nextSamplingInterval = '12 tháng (Hàng năm)';
  const recommendedTests: string[] = [];

  if (isHealthyCondition1) {
    primaryFault = 'NORMAL / HEALTHY';
    primaryFaultVi = 'Bình thường / Khỏe mạnh';
    confidence = 'Không có sự cố (Condition 1)';
    paperInvolvement = 'Giấy cách điện xenlulo trong tình trạng tốt, không có thoái hóa bất thường.';
    summaryVi = 'Tổng khí cháy (TDCG) ở mức Condition 1 (≤ 720 ppm) và tất cả nồng độ khí riêng lẻ đều nằm dưới ngưỡng cảnh báo theo tiêu chuẩn IEEE C57.104. Thuật toán chẩn đoán bỏ qua các phương pháp phân tích tỷ số và tam giác Duval để ngăn ngừa báo động giả do nhiễu nồng độ thấp.';
    recommendations.push('Tiếp tục vận hành bình thường ở chế độ tải định mức.');
    recommendations.push('Duy trì chu kỳ lấy mẫu thử nghiệm DGA định kỳ 12 tháng/lần.');
    nextSamplingInterval = '12 tháng (Hàng năm)';
    recommendedTests.push('Lấy mẫu định kỳ DGA và thử nghiệm hóa lý dầu hàng năm theo quy trình vận hành.');
  } else {
    // Condition >= 2 or active gassing
    // Prioritize Duval Triangles / Pentagons and IEC 60599
    if (['D1', 'D2'].includes(t1_zone) || ['D1', 'D2'].includes(iec_code) || keyGas_fault === 'Arcing') {
      primaryFault = t1_zone === 'D2' || iec_code === 'D2' ? 'High Energy Arcing (D2)' : 'Low Energy Discharge (D1)';
      primaryFaultVi = t1_zone === 'D2' || iec_code === 'D2' ? 'Phóng điện hồ quang năng lượng cao (D2)' : 'Phóng điện năng lượng thấp (D1)';
      confidence = (t1_zone === iec_code) ? 'Cao' : 'Trung bình';
      recommendedTests.push('Đo phóng điện cục bộ online / acoustic / UHF PD');
      recommendedTests.push('Đo điện áp đánh thủng dầu và kiểm tra độ sạch hạt than (Carbon particles)');
      recommendedTests.push('Đo điện trở một chiều cuộn dây và điện trở cách điện');
    } else if (['T1', 'T2', 'T3', 'T3-H'].includes(t1_zone) || ['T1', 'T2', 'T3'].includes(iec_code)) {
      const tCode = t5_zone === 'T3-H' ? 'T3-H' : (t1_zone !== 'OK' && t1_zone !== 'ND' ? t1_zone : iec_code);
      primaryFault = `Thermal Fault (${tCode})`;
      primaryFaultVi = tCode === 'T3' || tCode === 'T3-H'
        ? 'Quá nhiệt nhiệt độ cao > 700°C (T3 / T3-H)'
        : tCode === 'T2'
          ? 'Quá nhiệt nhiệt độ trung bình 300°C - 700°C (T2)'
          : 'Quá nhiệt nhiệt độ thấp < 300°C (T1)';
      confidence = 'Cao';
      recommendedTests.push('Đo nhiệt độ điểm nóng cuộn dây & chụp ảnh nhiệt hồng ngoại (IR Thermography) toàn bộ vỏ thùng và đầu cốt');
      recommendedTests.push('Kiểm tra tiếp xúc bộ chuyển nấc OLTC / DETC và đấu nối thanh cái');
    } else if (t1_zone === 'PD' || iec_code === 'PD' || keyGas_fault.includes('PD')) {
      primaryFault = 'Partial Discharge (PD)';
      primaryFaultVi = 'Phóng điện cục bộ (PD / Corona)';
      confidence = 'Trung bình';
      recommendedTests.push('Đo TEV / Siêu âm định vị nguồn PD trên thành máy biến áp');
      recommendedTests.push('Kiểm tra hàm lượng ẩm trong dầu và suy giảm bọt khí');
    } else {
      primaryFault = 'Developing Fault / Active Gassing';
      primaryFaultVi = 'Sự cố đang phát triển / Xuất hiện khí bất thường';
      confidence = 'Trung bình';
      recommendedTests.push('Thực hiện lấy mẫu lặp lại sau 1-2 tuần để xác nhận xu hướng gia tăng');
    }

    // Cellulose paper involvement
    if (co2_co_status === 'C' || t4_zone === 'C' || t5_zone === 'C' || p2_zone === 'C') {
      paperInvolvement = 'CẢNH BÁO: Phát hiện dấu hiệu liên quan đến giấy cách điện xenlulo (Paper involvement / Cellulose degradation). Cần đo hàm lượng Furan (2-FAL) để đánh giá độ trùng hợp DP.';
    } else {
      paperInvolvement = 'Chưa ghi nhận thoái hóa giấy cách điện xenlulo cấp tính; sự cố chủ yếu xảy ra trong khối dầu cách điện.';
    }

    summaryVi = `Phát hiện dấu hiệu sự cố ở mức ${conditionName} (TDCG = ${tdcg} ppm). Kết quả tổng hợp đa phương pháp (Duval Triangle/Pentagon, IEC 60599, Rogers) chỉ ra dạng sự cố chính: ${primaryFaultVi}. ${paperInvolvement}`;

    // Action plan based on condition
    if (conditionLevel === 2) {
      recommendations.push('Tăng cường giám sát, rút ngắn chu kỳ lấy mẫu DGA xuống còn 1 - 3 tháng/lần.');
      recommendations.push('Theo dõi sát mức độ phụ tải và nhiệt độ cuộn dây / lớp dầu đỉnh.');
      nextSamplingInterval = '1 - 3 tháng';
    } else if (conditionLevel === 3) {
      recommendations.push('Cảnh báo mức độ nghiêm trọng: Rút ngắn chu kỳ lấy mẫu thử nghiệm DGA xuống còn 1 - 2 tuần/lần.');
      recommendations.push('Lên phương án chuẩn bị dừng máy để kiểm tra sửa chữa nếu tốc độ sinh khí không giảm.');
      recommendations.push('Giới hạn phụ tải vận hành máy biến áp dưới 80% định mức.');
      nextSamplingInterval = '1 - 2 tuần';
    } else if (conditionLevel === 4) {
      recommendations.push('NGUY CẤP: Lên kế hoạch tách máy biến áp ra khỏi vận hành để kiểm tra và xử lý ngay lập tức.');
      recommendations.push('Thực hiện lấy mẫu kiểm tra DGA hàng ngày nếu chưa thể tách máy ngay.');
      recommendations.push('Thực hiện trọn bộ thí nghiệm điện: đo PD, cách điện, tổn hao điện môi Tan-Delta, và đo điện trở một chiều.');
      nextSamplingInterval = 'Hàng ngày / Tách máy ngay';
    }
  }

  // ----------------------------------------------------
  // CNAIM & Health Index Calculation (Preserving existing logic)
  // ----------------------------------------------------
  const age = gases.age || 20;
  const loadFactor = gases.loadFactor || 50;

  let dutyFactor = 1;
  if (loadFactor <= 50) dutyFactor = 0.9;
  else if (loadFactor <= 70) dutyFactor = 0.95;
  else if (loadFactor <= 100) dutyFactor = 1;
  else dutyFactor = 1.4;

  const normalExpectedLife = 60;
  const expectedLife = normalExpectedLife / dutyFactor;
  const beta1 = Math.log(5.5 / 0.5) / expectedLife;
  let initialHealthScore = 0.5 * Math.exp(beta1 * age);
  initialHealthScore = Math.min(5.5, initialHealthScore);

  let dgaScore = 0;
  if (h2 > 100) dgaScore += 10 * 50; else if (h2 > 40) dgaScore += 4 * 50;
  if (ch4 > 150) dgaScore += 10 * 30; else if (ch4 > 50) dgaScore += 4 * 30;
  if (c2h4 > 150) dgaScore += 10 * 30; else if (c2h4 > 50) dgaScore += 4 * 30;
  if (c2h6 > 150) dgaScore += 10 * 30; else if (c2h6 > 50) dgaScore += 4 * 30;
  if (c2h2 > 20) dgaScore += 8 * 120; else if (c2h2 > 5) dgaScore += 4 * 120;

  const dgaTestCollar = dgaScore / 220;
  let dgaFactor = 1;
  if (dgaTestCollar > 7) dgaFactor = 1.5;
  else if (dgaTestCollar > 4) dgaFactor = 1.2;

  let oilScore = 0;
  const moisture = gases.moisture || 0;
  const acidity = gases.acidity || 0;
  const bdStrength = gases.bdStrength || 60;

  if (moisture > 35) oilScore += 8 * 80; else if (moisture > 25) oilScore += 4 * 80;
  if (acidity > 0.2) oilScore += 8 * 125; else if (acidity > 0.15) oilScore += 4 * 125;
  if (bdStrength < 30) oilScore += 10 * 80; else if (bdStrength < 40) oilScore += 4 * 80;

  let oilFactor = 1;
  if (oilScore > 1000) oilFactor = 1.2;
  else if (oilScore > 500) oilFactor = 1.1;
  else if (oilScore > 200) oilFactor = 1.05;

  const factors = [dgaFactor, oilFactor];
  const maxFactor = Math.max(...factors);
  let healthScoreFactor = maxFactor;
  if (maxFactor > 1) {
    const otherFactors = factors.filter(f => f !== maxFactor && f > 1);
    const sumOther = otherFactors.reduce((a, b) => a + (b - 1), 0);
    healthScoreFactor = maxFactor + (sumOther / 1.5);
  }

  let currentHealthScore = initialHealthScore * healthScoreFactor;
  currentHealthScore = Math.max(currentHealthScore, dgaTestCollar);
  currentHealthScore = Math.min(10, currentHealthScore);

  let beta2 = beta1;
  if (currentHealthScore > 0.5) {
    beta2 = Math.log(currentHealthScore / 0.5) / age;
    beta2 = Math.min(beta2, 2 * beta1);
  }

  let ageingReduction = 1;
  if (currentHealthScore > 5.5) ageingReduction = 1.5;
  else if (currentHealthScore > 2) ageingReduction = ((currentHealthScore - 2) / 7) + 1;

  const futureHealthScore = Math.min(15, currentHealthScore * Math.exp((beta2 / ageingReduction) * 10));

  const getHIBand = (score: number) => {
    if (score < 4) return 'HI1';
    if (score < 5.5) return 'HI2';
    if (score < 6.5) return 'HI3';
    if (score < 8) return 'HI4';
    return 'HI5';
  };

  const currentHI = getHIBand(currentHealthScore);
  const futureHI = getHIBand(futureHealthScore);

  const dpEstimation = gases.estDp || Math.max(200, Math.round(1000 - 150 * Math.log(co / 10 + 1)));
  const currentPOF = (0.00073 * Math.exp(0.5 * currentHealthScore)).toFixed(4);
  const futurePOF = (0.00073 * Math.exp(0.5 * futureHealthScore)).toFixed(4);
  const eolYears = currentHealthScore >= 7 ? 0 : Math.max(0, Math.round(Math.log(7.0 / currentHealthScore) / beta1));

  const condition = currentHealthScore < 4 ? 'Tốt' : currentHealthScore < 6.5 ? 'Trung bình' : 'Kém';
  const recommendationColor = currentHealthScore >= 6.5 ? 'red' : currentHealthScore >= 4 ? 'yellow' : 'emerald';

  const healthScoreData = [
    {
      name: 'Hiện tại',
      'Health Score': parseFloat(currentHealthScore.toFixed(2)),
      fill: currentHealthScore < 4 ? '#22c55e' : currentHealthScore < 5.5 ? '#84cc16' : currentHealthScore < 6.5 ? '#eab308' : currentHealthScore < 8 ? '#f97316' : '#ef4444'
    },
    {
      name: 'Sau 10 năm',
      'Health Score': parseFloat(futureHealthScore.toFixed(2)),
      fill: futureHealthScore < 4 ? '#22c55e' : futureHealthScore < 5.5 ? '#84cc16' : futureHealthScore < 6.5 ? '#eab308' : futureHealthScore < 8 ? '#f97316' : '#ef4444'
    }
  ];

  // ----------------------------------------------------
  // Output format structure (in Vietnamese as required)
  // ----------------------------------------------------
  const reportSections = {
    part1_TDCG_GassingRate: {
      title: '1. BẢNG TỔNG HỢP NỒNG ĐỘ VÀ TỐC ĐỘ TĂNG KHÍ (TDCG & GASSING RATE)',
      tdcgValue: `${tdcg} ppm`,
      conditionText: `${conditionName} (${conditionTitleVi})`,
      rateText: gassingRate.summary,
      decisionText: isHealthyCondition1
        ? 'KẾT LUẬN SƠ BỘ: NORMAL / HEALTHY. Không phát hiện sự cố hoạt tính, nồng độ khí ở mức tạp âm tự nhiên. BỎ QUA phân tích lỗi để loại trừ báo động giả.'
        : 'KẾT LUẬN SƠ BỘ: FAULT DETECTED / ACTIVE GASSING. Tổng khí cháy hoặc tốc độ sinh khí vượt ngưỡng giới hạn, tiến hành phân tích chẩn đoán lỗi chuyên sâu.'
    },
    part2_MethodDiagnostics: {
      title: '2. KẾT QUẢ CHẨN ĐOÁN CHI TIẾT TỪ CÁC PHƯƠNG PHÁP',
      methods: [
        {
          name: 'Tỷ số CO₂ / CO (Lão hóa giấy)',
          result: co2_co_status,
          note: co2_co_desc
        },
        {
          name: 'Phương pháp Khí chủ đạo (Key Gas)',
          result: keyGas_fault,
          note: keyGas_desc
        },
        {
          name: 'Tỷ số Rogers (IEEE PC57.104)',
          result: rogers_code,
          note: rogers_desc
        },
        {
          name: 'Tỷ số IEC 60599',
          result: iec_code,
          note: iec_desc
        },
        {
          name: 'Tỷ số Dornenburg',
          result: dornenburg_code,
          note: dornenburg_desc
        },
        {
          name: 'Tam giác Duval 1 (T1)',
          result: t1_zone,
          note: t1_name
        },
        {
          name: 'Tam giác Duval 4 (T4 tinh chỉnh)',
          result: t4_zone,
          note: t4_name
        },
        {
          name: 'Tam giác Duval 5 (T5 tinh chỉnh)',
          result: t5_zone,
          note: t5_name
        },
        {
          name: 'Ngũ giác Duval P1',
          result: p1_zone,
          note: p1_name
        },
        {
          name: 'Ngũ giác Duval P2 (tinh chỉnh)',
          result: p2_zone,
          note: p2_name
        },
        {
          name: 'Phương pháp Ma trận ETRA',
          result: etra_res,
          note: etra_desc
        }
      ]
    },
    part3_FinalSynthesis: {
      title: '3. KẾT LUẬN CUỐI CÙNG (DGA MATRIX SYNTHESIS)',
      faultType: primaryFaultVi,
      confidence: confidence,
      paperInvolvement: paperInvolvement,
      conclusions: [
        `Phân loại sự cố chính: ${primaryFaultVi} (Độ tin cậy: ${confidence})`,
        paperInvolvement,
        `Tình trạng sức khỏe máy biến áp: Health Index = ${currentHealthScore.toFixed(2)} (${currentHI}), xác suất sự cố hàng năm PoF = ${(parseFloat(currentPOF) * 100).toFixed(2)}%.`
      ]
    },
    part4_ActionRecommendations: {
      title: '4. KHUYẾN CÁO VẬN HÀNH BẢO TRÌ',
      samplingInterval: `Chu kỳ lấy mẫu thí nghiệm tiếp theo: ${nextSamplingInterval}`,
      precautions: recommendations,
      diagnosticTests: recommendedTests
    }
  };

  return {
    tdcg,
    conditionLevel,
    conditionName,
    conditionTitleVi,
    isIndividualNormal,
    exceededGases,
    gassingRate,
    status,
    faultDiagnosisSkipped,
    skipReason,
    matrix,
    detailedResults: {
      co2_co: {
        ratio: co2_co_ratio,
        applicable: co2_co_applicable,
        status: co2_co_status,
        description: co2_co_desc
      },
      keyGas: {
        principalGas: keyGas_principal,
        fault: keyGas_fault,
        description: keyGas_desc
      },
      rogers: {
        r1: r1_rogers,
        r2: r2_rogers,
        r5: r5_rogers,
        code: rogers_code,
        fault: rogers_fault,
        description: rogers_desc
      },
      iec60599: {
        r1: r1_rogers,
        r2: r2_rogers,
        r5: r5_rogers,
        code: iec_code,
        fault: iec_fault,
        description: iec_desc
      },
      dornenburg: {
        applicable: dornenburg_applicable,
        r1: d_r1,
        r2: d_r2,
        r3: d_r3,
        r4: d_r4,
        code: dornenburg_code,
        fault: dornenburg_fault,
        description: dornenburg_desc
      },
      duvalT1: {
        pCH4: pCH4_t1,
        pC2H4: pC2H4_t1,
        pC2H2: pC2H2_t1,
        zone: t1_zone,
        faultName: t1_name
      },
      duvalT4: {
        applicable: t4_applicable,
        pH2: pH2_t4,
        pCH4: pCH4_t4,
        pC2H6: pC2H6_t4,
        zone: t4_zone,
        faultName: t4_name,
        refinedReason: t4_refinedReason
      },
      duvalT5: {
        applicable: t5_applicable,
        pCH4: pCH4_t5,
        pC2H4: pC2H4_t5,
        pC2H6: pC2H6_t5,
        zone: t5_zone,
        faultName: t5_name,
        refinedReason: t5_refinedReason
      },
      duvalP1: {
        x: pentagonCoords.x,
        y: pentagonCoords.y,
        zone: p1_zone,
        faultName: p1_name
      },
      duvalP2: {
        applicable: p2_applicable,
        x: pentagonCoords.x,
        y: pentagonCoords.y,
        zone: p2_zone,
        faultName: p2_name
      },
      etra: {
        result: etra_res,
        description: etra_desc
      }
    },
    synthesis: {
      primaryFault,
      primaryFaultVi,
      confidence,
      paperInvolvement,
      summaryVi
    },
    recommendations,
    nextSamplingInterval,
    recommendedTests,
    reportSections,
    currentHealthScore: currentHealthScore.toFixed(2),
    futureHealthScore: futureHealthScore.toFixed(2),
    currentHI,
    futureHI,
    currentPOF,
    futurePOF,
    eolYears,
    dpEstimation,
    condition,
    tdcgColor,
    recommendationColor,
    healthScoreData,
    timestamp: new Date().toLocaleString('vi-VN')
  };
}
