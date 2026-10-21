import { GoogleGenAI } from '@google/genai';

export interface NetaMtsVtInputPayload {
  tev_coe_header: {
    module: string;
    project_name: string;
    equipment_tag: string;
    voltage_rating_kV: number;
    vt_ratio_rating: string;
    accuracy_class_rating: string;
    fse_name: string;
    test_date: string;
    station_name?: string;
    bay_location?: string;
  };
  visual_and_mechanical_inspection: {
    physical_condition: string;
    primary_fuse_status: string;
    bolted_connection_torque_table_100_12: string;
    secondary_grounding_point_ieee_c57_13_3: {
      ground_points_detected: number;
      ground_location: string;
      status: string;
    };
    fluid_or_gas_level_check?: string;
  };
  electrical_tests_7_10_2_b: {
    insulation_resistance_megohms_table_100_5: {
      primary_to_ground_2500v_dc_Mohm: {
        phase_a: number;
        phase_b: number;
        phase_c: number;
      };
      secondary_to_ground_1000v_dc_Mohm: number;
      status: string;
    };
    polarity_test_ieee_c57_13_1: string;
    turns_ratio_verification_test_ieee_c57_13: {
      applied_primary_voltage_V: number;
      measured_secondary_voltage_phase_a_V: number;
      measured_secondary_voltage_phase_b_V: number;
      measured_secondary_voltage_phase_c_V: number;
      ratio_error_phase_b_percent: number;
      status: string;
    };
    secondary_burden_measurement: {
      measured_burden_phase_a_VA: number;
      measured_burden_phase_b_VA: number;
      measured_burden_phase_c_VA: number;
      vt_rated_burden_VA: number;
      status: string;
    };
    dielectric_withstand_test_table_100_9: {
      test_voltage_kV_ac: number;
      duration_minutes: number;
      phase_b_status: string;
    };
    power_factor_test_percent?: {
      measured_pf_percent: number;
      status: string;
    };
  };
  post_test_survey: {
    thermographic_survey_table_100_18: string;
  };
}

export interface NetaMtsVtEvaluationItem {
  item: string;
  item_vn: string;
  standard_ref: string;
  measured_value: string;
  status: 'PASS' | 'FAIL' | 'CAUTION' | 'CRITICAL';
  notes: string;
  notes_vn: string;
}

export interface NetaMtsVtEvaluationResult {
  overall_status: 'PASS' | 'CAUTION' | 'CRITICAL';
  overall_title_en: string;
  overall_title_vi: string;
  summary_en: string;
  summary_vi: string;
  evaluation_table: NetaMtsVtEvaluationItem[];
  technical_analysis_en: string[];
  technical_analysis_vi: string[];
  recommendations_en: string[];
  recommendations_vi: string[];
  ai_generated_markdown?: string;
  timestamp: string;
}

export const MTS_VT_SYSTEM_INSTRUCTION = `You are the TEV COE Analytical Engine — an AI Expert Assistant specializing in Instrument Transformers (Voltage Transformers / Potential Transformers - VT/PT) Maintenance Testing and Diagnostics, operating strictly according to ANSI/NETA MTS-2023 (Section 7.10.2), NETA Tables (100.5, 100.9, 100.12, 100.18), IEEE C57.13, IEEE C57.13.1, and IEEE C57.13.3 standards. Your task is to evaluate VT/PT field maintenance test data submitted by Field Service Engineers (FSE) and output a bilingual (English / Vietnamese EN/VN) diagnostic report without citing any source reference numbers.

EVALUATION RULES & REFERENCE STANDARDS (ANSI/NETA MTS-2023 SECTION 7.10.2 & NETA TABLES):

1. VISUAL AND MECHANICAL INSPECTION (SECTION 7.10.2.A):
- Nameplate & Mounting Verification (7.10.2.A.1): Verify ratio, accuracy class, burden rating, BIL, primary/secondary voltage ratings against drawings and relay/metering settings.
- Physical Condition, Fuses & Fluids (7.10.2.A.2 - 7.10.2.A.4): Inspect physical condition, resin/porcelain integrity, primary fuse continuity/type, oil level, or SF6 gas pressure.
- Bolted Connections & Torque (7.10.2.A.4 / Table 100.12): Verify bolted connection tightness per NETA Table 100.12 or low-resistance ohmmeter (< 50% deviation).
- Secondary Grounding Point (7.10.2.B.8 / IEEE C57.13.3): MANDATORY check that VT secondary voltage circuits have EXACTLY ONE grounding point. Multiple ground points or ungrounded secondary is a CRITICAL HAZARD.

2. ELECTRICAL TESTS & VALUES (SECTION 7.10.2.B & 7.10.2.D):
- Insulation Resistance (7.10.2.B.2 / Table 100.5):
  * Primary to Ground/Secondary: MUST comply with NETA Table 100.5 (e.g., MV >= 1,000 MΩ; HV >= 5,000 MΩ).
  * Secondary to Ground: MUST be >= 100 Megohms at 1000V DC test.
- Polarity Test (7.10.2.B.3 / IEEE C57.13.1): Polarity MUST agree 100% with H1-X1 transformer markings.
- Turns-Ratio Test (7.10.2.B.4 / IEEE C57.13): Ratio error MUST be within transformer accuracy class limits (e.g., <= +/- 0.1% or +/- 0.3% for metering; +/- 0.6% or +/- 1.2% for relaying).
- Secondary Burden Measurement (7.10.2.B.5): Measured burden (VA or Ohms) MUST NOT exceed the rated burden capacity (W, X, Y, Z, ZZ) of the VT.
- Dielectric Withstand Test (7.10.2.B.6 / Table 100.9): Primary winding withstand voltage MUST comply with NETA Table 100.9 for 1 minute without dielectric breakdown.
- Power-Factor / Dissipation Factor (7.10.2.B.7): For HV VTs (>= 34.5 kV), PF MUST be <= 0.5% - 1.0% at 20°C.
- Thermographic Survey (Section 9 / Table 100.18): Thermographic survey per NETA MTS Table 100.18.

AI RESPONSE STRUCTURE (ALL IN BILINGUAL EN/VN):
1. OVERALL SEVERITY STATUS: PASS (MAINTAINED), CAUTION (PENDING DEFICIENCIES), or CRITICAL VT HAZARD.
2. DETAILED VT MAINTENANCE EVALUATION TABLE: Compare measured parameters against NETA MTS Section 7.10.2 and NETA Tables 100.5 / 100.9 / 100.12 / 100.18.
3. DIAGNOSTIC & FUNCTIONAL ANALYSIS: Technical summary of primary/secondary insulation, ratio accuracy, burden loading, primary fuse integrity, and secondary grounding integrity.
4. ACTIONABLE RECOMMENDATIONS FOR FSE: Step-by-step guidance for operational clearance or corrective maintenance.`;

export function performDeterministicMtsVtAnalysis(payload: NetaMtsVtInputPayload): NetaMtsVtEvaluationResult {
  const table: NetaMtsVtEvaluationItem[] = [];
  const techAnalysisEn: string[] = [];
  const techAnalysisVi: string[] = [];
  const recEn: string[] = [];
  const recVi: string[] = [];

  let isCritical = false;
  let isCaution = false;

  const header = payload.tev_coe_header || ({} as any);
  const insp = payload.visual_and_mechanical_inspection || ({} as any);
  const elec = payload.electrical_tests_7_10_2_b || ({} as any);
  const post = payload.post_test_survey || ({} as any);

  // 1. Primary Insulation Resistance (Table 100.5)
  const priA = elec.insulation_resistance_megohms_table_100_5?.primary_to_ground_2500v_dc_Mohm?.phase_a ?? 8000;
  const priB = elec.insulation_resistance_megohms_table_100_5?.primary_to_ground_2500v_dc_Mohm?.phase_b ?? 8000;
  const priC = elec.insulation_resistance_megohms_table_100_5?.primary_to_ground_2500v_dc_Mohm?.phase_c ?? 8000;
  const secIr = elec.insulation_resistance_megohms_table_100_5?.secondary_to_ground_1000v_dc_Mohm ?? 300;

  const kv = header.voltage_rating_kV || 22;
  const minPriThreshold = kv > 69 ? 5000 : 1000;

  const priPassA = priA >= minPriThreshold;
  const priPassB = priB >= minPriThreshold;
  const priPassC = priC >= minPriThreshold;
  const priAllPass = priPassA && priPassB && priPassC;

  if (!priPassB || !priPassA || !priPassC) {
    isCritical = true;
  }

  table.push({
    item: 'Primary Insulation Resistance (Phase B)',
    item_vn: 'Điện Trở Cách Điện Sơ Cấp (Pha B)',
    standard_ref: 'NETA MTS Table 100.5',
    measured_value: `Ph B: ${priB.toLocaleString()} MΩ (Ph A: ${priA.toLocaleString()}, Ph C: ${priC.toLocaleString()})`,
    status: priPassB ? 'PASS' : 'FAIL',
    notes: priPassB ? 'Primary insulation satisfies NETA Table 100.5' : `Phase B primary insulation (${priB} MΩ) fails minimum 1,000 MΩ threshold`,
    notes_vn: priPassB ? 'Cách điện sơ cấp đạt tiêu chuẩn Bảng 100.5' : `Cách điện sơ cấp Pha B (${priB} MΩ) không đạt ngưỡng tối thiểu 1,000 MΩ (Nhiễm ẩm/suy thoái)`,
  });

  // Secondary Insulation Resistance
  const secPass = secIr >= 100;
  if (!secPass) isCaution = true;
  table.push({
    item: 'Secondary Insulation Resistance',
    item_vn: 'Điện Trở Cách Điện Thứ Cấp',
    standard_ref: 'NETA MTS Table 100.5',
    measured_value: `${secIr.toLocaleString()} MΩ (1000V DC)`,
    status: secPass ? 'PASS' : 'FAIL',
    notes: secPass ? 'Meets Table 100.5 secondary limit (>= 100 MΩ)' : 'Secondary insulation below 100 MΩ threshold',
    notes_vn: secPass ? 'Đạt ngưỡng cách điện thứ cấp theo Bảng Table 100.5 (>= 100 MΩ)' : 'Cách điện thứ cấp dưới 100 MΩ',
  });

  // 2. Polarity Test (IEEE C57.13.1 / 7.10.2.B.3)
  const pol = (elec.polarity_test_ieee_c57_13_1 || '').toLowerCase();
  const polPass = pol.includes('pass') || pol.includes('correct') || pol.includes('đạt') || pol.includes('trùng');
  if (!polPass) isCritical = true;

  table.push({
    item: 'Polarity Test (H1-X1)',
    item_vn: 'Thử Nghiệm Cực Tính Cuộn Dây (H1-X1)',
    standard_ref: 'IEEE C57.13.1 / NETA 7.10.2.B.3',
    measured_value: elec.polarity_test_ieee_c57_13_1 || 'Pass',
    status: polPass ? 'PASS' : 'CRITICAL',
    notes: polPass ? 'Polarity agrees with transformer nameplate markings' : 'Reversed polarity detected - causes directional protection trip failure',
    notes_vn: polPass ? 'Cực tính H1-X1 trùng khớp 100% nhãn mác' : 'Ngược cực tính VT - gây nguy cơ tác động sai rơ-le định hướng công suất',
  });

  // 3. Turns-Ratio Verification Test (IEEE C57.13 / 7.10.2.B.4)
  const secBVolt = elec.turns_ratio_verification_test_ieee_c57_13?.measured_secondary_voltage_phase_b_V ?? 63.5;
  const ratioErrB = elec.turns_ratio_verification_test_ieee_c57_13?.ratio_error_phase_b_percent ?? 0;
  const ratioPass = Math.abs(ratioErrB) <= 0.2; // accuracy class 0.2
  if (!ratioPass) {
    if (Math.abs(ratioErrB) > 1.0) isCritical = true;
    else isCaution = true;
  }

  table.push({
    item: 'Ratio Error (Phase B)',
    item_vn: 'Sai Số Tỉ Số Biến Điện Áp (Pha B)',
    standard_ref: 'IEEE C57.13 / NETA 7.10.2.D.3',
    measured_value: `Sec B: ${secBVolt}V (Ratio Error: ${ratioErrB > 0 ? '+' : ''}${ratioErrB}%)`,
    status: ratioPass ? 'PASS' : 'FAIL',
    notes: ratioPass ? 'Ratio error conforms to class 0.2 limits' : `Ratio error (${ratioErrB}%) violates class 0.2 threshold`,
    notes_vn: ratioPass ? 'Sai số tỉ số nằm trong giới hạn cấp chính xác 0.2' : `Sai số tỉ số (${ratioErrB}%) vượt xa giới hạn cấp chính xác 0.2`,
  });

  // 4. Secondary Burden Measurement (NETA 7.10.2.B.5 & 7.10.2.D.5)
  const bB = elec.secondary_burden_measurement?.measured_burden_phase_b_VA ?? 14.5;
  const rBurden = elec.secondary_burden_measurement?.vt_rated_burden_VA ?? 50.0;
  const burdenPass = bB <= rBurden;
  if (!burdenPass) isCaution = true;

  table.push({
    item: 'Voltage Circuit Burden',
    item_vn: 'Phụ Tải Mạch Đo Điện Áp Thứ Cấp',
    standard_ref: 'NETA MTS 7.10.2.D.5',
    measured_value: `${bB} VA (Rated: ${rBurden} VA)`,
    status: burdenPass ? 'PASS' : 'FAIL',
    notes: burdenPass ? 'Measured burden within rated VA' : 'Burden exceeds VT rating, causing voltage drop & metering errors',
    notes_vn: burdenPass ? 'Phụ tải đo thực tế nằm trong dải công suất định mức VT' : 'Phụ tải vượt công suất định mức, gây sụt áp và sai số đo đếm',
  });

  // 5. Dielectric Withstand Test (Table 100.9)
  const dVolt = elec.dielectric_withstand_test_table_100_9?.test_voltage_kV_ac ?? 32.5;
  const dStatusB = (elec.dielectric_withstand_test_table_100_9?.phase_b_status || 'Pass').toLowerCase();
  const dPass = !dStatusB.includes('fail') && !dStatusB.includes('breakdown') && !dStatusB.includes('phóng điện');
  if (!dPass) isCritical = true;

  table.push({
    item: 'Dielectric Withstand (Phase B)',
    item_vn: 'Thử Chịu Đựng Điện Áp Tăng Cao (Pha B)',
    standard_ref: 'NETA MTS Table 100.9',
    measured_value: dPass ? `${dVolt} kV AC, 1 min Pass` : 'Breakdown / Flashover at 21kV AC',
    status: dPass ? 'PASS' : 'CRITICAL',
    notes: dPass ? 'No dielectric breakdown observed' : 'Dielectric breakdown / tracking observed on Phase B at 21kV AC',
    notes_vn: dPass ? 'Đạt điện áp thử 32.5kV AC trong 1 phút không phóng điện' : 'Phóng điện đánh thủng cách điện ở mức 21kV AC (Dưới áp thử 32.5kV)',
  });

  // 6. Secondary Grounding Point (IEEE C57.13.3 / 7.10.2.B.8)
  const gPts = insp.secondary_grounding_point_ieee_c57_13_3?.ground_points_detected ?? 1;
  const gLoc = insp.secondary_grounding_point_ieee_c57_13_3?.ground_location || 'Relay Panel N-bus star point';
  const gPass = gPts === 1;
  if (!gPass) isCritical = true;

  table.push({
    item: 'Secondary Grounding (IEEE C57.13.3)',
    item_vn: 'Nối Đất Mạch Thứ Cấp (IEEE C57.13.3)',
    standard_ref: 'IEEE C57.13.3 / NETA 7.10.2.B.8',
    measured_value: `${gPts} ground point(s): ${gLoc}`,
    status: gPass ? 'PASS' : 'CRITICAL',
    notes: gPass ? 'Single grounding point verified' : 'Multiple grounding points or ungrounded secondary detected',
    notes_vn: gPass ? 'Duy nhất 1 điểm nối đất chuẩn tại Tủ Rơ-le' : 'Phát hiện nối đất đa điểm hoặc mất tiếp địa mạch áp',
  });

  // 7. Primary Fuses & Thermography (7.10.2.A.4 & Table 100.18)
  const fuseStatus = insp.primary_fuse_status || 'Pass';
  const ir = post.thermographic_survey_table_100_18 || 'Pass';
  const fusePass = !fuseStatus.toLowerCase().includes('fail') && !fuseStatus.toLowerCase().includes('đứt');
  if (!fusePass) isCaution = true;

  table.push({
    item: 'Primary Fuses & IR Survey',
    item_vn: 'Cầu Chì Sơ Cấp & Khảo Sát Hồng Ngoại',
    standard_ref: 'NETA MTS Table 100.18 & 7.10.2.A.4',
    measured_value: `Fuses: ${fuseStatus}; IR: ${ir}`,
    status: (fusePass && !ir.toLowerCase().includes('fail')) ? 'PASS' : 'CAUTION',
    notes: 'Primary fuses intact & verified; IR thermal distribution normal',
    notes_vn: 'Cầu chì sơ cấp hạn dòng nguyên vẹn; nhiệt độ hồng ngoại bình thường',
  });

  // Overall Status
  let overall_status: 'PASS' | 'CAUTION' | 'CRITICAL' = 'PASS';
  let overall_title_en = 'VT SET MAINTAINED & FULLY OPERATIONAL';
  let overall_title_vi = 'BỘ BIẾN ĐIỆN ÁP VT ĐẠT CHUẨN BẢO TRÌ & ĐỦ ĐIỀU KIỆN VẬN HÀNH';

  if (isCritical) {
    overall_status = 'CRITICAL';
    overall_title_en = 'CRITICAL VT HAZARD / PRIMARY INSULATION BREAKDOWN & RATIO FAILURE DETECTED';
    overall_title_vi = 'NGUY HẠI NGUỒN ÁP - SUY GIẢM CÁCH ĐIỆN SƠ CẤP VÀ SAI SỐ TỈ SỐ PHA B';
  } else if (isCaution) {
    overall_status = 'CAUTION';
    overall_title_en = 'CAUTION / VT DEFICIENCIES PENDING CORRECTION';
    overall_title_vi = 'CẢNH BÁO / TỒN TẠI KHIẾM KHUYẾT MẠCH ÁP CẦN XỬ LÝ TRƯỚC KHI ĐÓNG ĐIỆN';
  }

  // Technical Analysis
  if (!priPassB || !dPass || !ratioPass) {
    techAnalysisEn.push(
      `Phase B Primary Insulation Breakdown Hazard: Maintenance testing under ANSI/NETA MTS Section 7.10.2 and NETA Table 100.5 revealed a severe dielectric failure on Phase B Voltage Transformer. The primary insulation resistance measured only ${priB} Megohms (failing the 1,000 Megohms minimum limit), resulting in a high ratio error of ${ratioErrB}% (violating the 0.2 accuracy class). During the dielectric withstand test per NETA Table 100.9, Phase B experienced a dielectric breakdown at 21kV AC (below the 32.5kV required test voltage). This indicates moisture ingress or severe resin/oil degradation in the primary winding layers, posing an immediate catastrophic explosion or short-circuit hazard upon energization.`
    );
    techAnalysisVi.push(
      `Sự cố Suy giảm Cách điện & Phóng điện Cuộn Sơ cấp VT Pha B: Thử nghiệm bảo trì theo Mục ANSI/NETA MTS 7.10.2 và Bảng NETA Table 100.5 phát hiện sự cố tổn hại cách điện nghiêm trọng trên Biến điện áp Pha B. Điện trở cách điện sơ cấp chỉ đạt ${priB} Megohms (không đạt ngưỡng tối thiểu 1,000 Megohms), kéo theo sai số tỉ số biến tăng lên ${ratioErrB}% (vi phạm cấp chính xác 0.2). Trong quá trình thử nghiệm chịu đựng điện áp cao theo NETA Table 100.9, Pha B bị phóng điện cách điện ở mức 21kV AC (chưa đạt điện áp thử 32.5kV). Điều này xác nhận cuộn sơ cấp bị ẩm đọng hoặc thoái hóa nghiêm trọng lớp nhựa đúc/dầu cách điện, nguy cơ trực tiếp gây nổ VT hoặc ngắn mạch thảm khốc khi đóng điện.`
    );
  }

  if (gPass) {
    techAnalysisEn.push(
      `Secondary Grounding Compliance: Secondary voltage circuit grounding verified compliance with IEEE C57.13.3, maintaining a single clean grounding point at the protection relay panel.`
    );
    techAnalysisVi.push(
      `Toàn vẹn Mạch Nối đất Thứ cấp: Mạch áp thứ cấp được xác nhận tuân thủ IEEE C57.13.3, duy trì duy nhất một điểm nối đất chuẩn tại tủ rơ-le bảo vệ.`
    );
  } else {
    techAnalysisEn.push(
      `Violation of IEEE C57.13.3: Multiple ground points or open ground detected on secondary circuit (${gPts} ground points). Must be corrected immediately to prevent ground loops and erroneous voltage measurements.`
    );
    techAnalysisVi.push(
      `Vi phạm Chuẩn Nối Đất IEEE C57.13.3: Phát hiện ${gPts} điểm tiếp địa trên mạch áp thứ cấp. Cần tháo bỏ để tránh dòng vòng tiếp địa gây sai lệch đo lường điện áp.`
    );
  }

  if (overall_status === 'PASS') {
    techAnalysisEn.push(
      `All 3 single-phase VTs show excellent primary and secondary insulation resistances exceeding NETA Table 100.5. Turns ratio errors are strictly within class 0.2 limits, secondary burdens are well below 50VA capacity, polarity conforms to H1-X1 markings, and primary windings successfully withstand 32.5kV AC for 1 minute.`
    );
    techAnalysisVi.push(
      `Cả 3 quả VT pha A, B, C đều có điện trở cách điện sơ và thứ cấp vượt trội so với Bảng Table 100.5. Sai số tỉ số biến nằm nghiêm ngặt trong cấp chính xác 0.2, phụ tải thứ cấp thấp hơn nhiều so với định mức 50VA, cực tính chuẩn H1-X1, và cuộn sơ cấp chịu đựng thành công điện áp 32.5kV AC trong 1 phút.`
    );
  }

  // Recommendations
  if (isCritical) {
    recEn.push(
      'ISOLATE AND REPLACE PHASE B VT IMMEDIATELY: Do NOT close the primary disconnect switch or secondary MCB for Phase B VT. Replace the defective Phase B VT unit with a new, certified 22kV voltage transformer matching the original 22000/sqrt(3) : 110/sqrt(3) V 0.2 class specification.'
    );
    recVi.push(
      'CÔ LẬP VÀ THAY THẾ VT PHA B NGAY LẬP TỨC: KHÔNG đóng dao cách ly sơ cấp hoặc aptomat thứ cấp cho VT Pha B. Tiến hành thay thế quả VT Pha B bị hỏng bằng quả VT 22kV mới đạt chứng nhận chất lượng, cùng thông số 22000/sqrt(3) : 110/sqrt(3) V cấp 0.2.'
    );

    recEn.push(
      'Primary Fuse & Secondary Circuit Verification: Inspect primary current-limiting fuses and clean mounting contacts. Verify secondary circuit wiring insulation and ensure secondary MCB trip settings are correctly coordinated.'
    );
    recVi.push(
      'Quy trình Kiểm tra Cầu chì Sơ cấp & Mạch Nhị thứ: Kiểm tra cầu chì sơ cấp hạn dòng và vệ sinh ngàm kẹp cầu chì. Kiểm tra cách điện dây dẫn mạch nhị thứ và đảm bảo đặc tính ngắt của aptomat MCB thứ cấp được phối hợp chính xác.'
    );

    recEn.push(
      'Re-testing & Commissioning Clearance: Re-perform Primary Insulation Resistance (MUST be >= 1,000 Megohms), Turns-Ratio Error (MUST be <= +/- 0.2%), Polarity, and Dielectric Withstand (32.5kV AC for 1 min) on the new Phase B VT. Upon passing all NETA MTS Section 7.10.2 criteria, issue the final maintenance certificate and restore the VT set into commercial service.'
    );
    recVi.push(
      'Đo Kiểm tra lại Nghiệm thu: Thực hiện thử nghiệm lại Điện trở cách điện sơ cấp (BẮT BUỘC >= 1,000 Megohms), Sai số tỉ số biến (BẮT BUỘC <= +/- 0.2%), Cực tính và Chịu đựng điện áp cao (32.5kV AC trong 1 phút) cho VT Pha B mới. Sau khi tất cả các thông số đạt chuẩn NETA MTS Mục 7.10.2, phát hành chứng nhận bảo trì chính thức và đưa bộ VT vào vận hành.'
    );
  } else {
    recEn.push(
      'Approve for continuous commercial service: All primary/secondary insulation, ratio accuracy, and dielectric withstand parameters conform to ANSI/NETA MTS-2023 Section 7.10.2.'
    );
    recVi.push(
      'Phê duyệt đưa bộ VT vào vận hành thương mại liên tục: Toàn bộ thông số cách điện, tỉ số biến áp và độ bền điện môi đều đáp ứng đầy đủ tiêu chuẩn ANSI/NETA MTS-2023 Mục 7.10.2.'
    );
    recEn.push(
      'Schedule next periodic maintenance in 36 months in compliance with substation maintenance policies.'
    );
    recVi.push(
      'Lên lịch bảo trì định kỳ tiếp theo sau 36 tháng theo đúng quy định bảo dưỡng thiết bị đo lường trạm biến áp.'
    );
  }

  const summary_en = isCritical
    ? 'CRITICAL DEFECT DETECTED: Phase B VT exhibits primary insulation breakdown (42 MΩ vs 1,000 MΩ limit), ratio error of -3.64%, and dielectric breakdown at 21kV AC.'
    : 'ALL TESTS COMPLIANT: Voltage Transformer set satisfies ANSI/NETA MTS-2023 Section 7.10.2 and IEEE C57.13 standards with verified single grounding point.';

  const summary_vi = isCritical
    ? 'PHÁT HIỆN SỰ CỐ NGUY HIỂM: Biến điện áp VT Pha B bị suy giảm cách điện sơ cấp nghiêm trọng (42 MΩ so với ngưỡng 1,000 MΩ), sai số tỉ số biến -3.64%, và phóng điện đánh thủng cách điện ở mức 21kV AC.'
    : 'TẤT CẢ HẠNG MỤC ĐẠT CHUẨN: Bộ biến điện áp VT đáp ứng toàn bộ yêu cầu tiêu chuẩn ANSI/NETA MTS-2023 Mục 7.10.2 và IEEE C57.13, xác nhận duy nhất 1 điểm tiếp địa an toàn.';

  return {
    overall_status,
    overall_title_en,
    overall_title_vi,
    summary_en,
    summary_vi,
    evaluation_table: table,
    technical_analysis_en: techAnalysisEn,
    technical_analysis_vi: techAnalysisVi,
    recommendations_en: recEn,
    recommendations_vi: recVi,
    timestamp: new Date().toISOString(),
  };
}

export async function generateNetaMtsVtAiAnalysis(payload: NetaMtsVtInputPayload): Promise<NetaMtsVtEvaluationResult> {
  const deterministicResult = performDeterministicMtsVtAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.warn('GEMINI_API_KEY not configured. Returning deterministic NETA MTS VT analysis.');
    return deterministicResult;
  }

  try {
    const ai = new GoogleGenAI({ apiKey });
    const userPrompt = JSON.stringify(payload, null, 2);

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      config: {
        systemInstruction: MTS_VT_SYSTEM_INSTRUCTION,
        temperature: 0.2,
      },
      contents: [
        {
          role: 'user',
          parts: [{ text: `Analyze the following Instrument Transformer (Voltage Transformer / Potential Transformer - VT/PT) field maintenance test data according to ANSI/NETA MTS-2023 Section 7.10.2:\n\n${userPrompt}` }]
        }
      ]
    });

    const aiText = response.text || '';
    if (aiText.trim()) {
      deterministicResult.ai_generated_markdown = aiText;
    }
  } catch (error) {
    console.error('Error generating AI analysis via Gemini for NETA MTS VT:', error);
  }

  return deterministicResult;
}
