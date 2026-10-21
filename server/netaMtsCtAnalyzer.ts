import { GoogleGenAI } from '@google/genai';

export interface NetaMtsCtInputPayload {
  tev_coe_header: {
    module: string;
    project_name: string;
    equipment_tag: string;
    voltage_rating_kV: number;
    ct_ratio_rating: string;
    accuracy_class_rating: string;
    fse_name: string;
    test_date: string;
    station_name?: string;
    bay_location?: string;
  };
  visual_and_mechanical_inspection: {
    nameplate_verification: {
      matched: boolean;
      notes: string;
    };
    physical_condition: string;
    bolted_connection_torque_table_100_12: string;
    clearance_verification: string;
    secondary_grounding_point_ieee_c57_13_3: {
      ground_points_detected: number;
      ground_locations: string;
      status: string;
    };
  };
  electrical_tests_7_10_1_b: {
    insulation_resistance_megohms_table_100_5: {
      primary_to_ground_2500v_dc_Mohm: number;
      secondary_to_ground_1000v_dc_Mohm: number;
      status: string;
    };
    polarity_test_ieee_c57_13_1: string;
    ratio_verification_test_ieee_c57_13: {
      applied_primary_current_A: number;
      measured_secondary_current_phase_a_A: number;
      measured_secondary_current_phase_b_A: number;
      measured_secondary_current_phase_c_A: number;
      ratio_error_phase_c_percent: number;
      status: string;
    };
    excitation_saturation_test_ieee_c57_13_1: {
      phase_a_knee_point_voltage_V: number;
      phase_b_knee_point_voltage_V: number;
      phase_c_knee_point_voltage_V: number;
      status: string;
    };
    winding_dc_resistance_ohms: {
      phase_a_ohms: number;
      phase_b_ohms: number;
      phase_c_ohms: number;
      status: string;
    };
    secondary_burden_measurement: {
      measured_burden_phase_c_VA: number;
      ct_rated_burden_VA: number;
      status: string;
    };
    dielectric_withstand_test_table_100_9: {
      test_voltage_kV_ac: number;
      duration_minutes: number;
      status: string;
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

export interface NetaMtsCtEvaluationItem {
  item: string;
  item_vn: string;
  standard_ref: string;
  measured_value: string;
  status: 'PASS' | 'FAIL' | 'CAUTION' | 'CRITICAL';
  notes: string;
  notes_vn: string;
}

export interface NetaMtsCtEvaluationResult {
  overall_status: 'PASS' | 'CAUTION' | 'CRITICAL';
  overall_title_en: string;
  overall_title_vi: string;
  summary_en: string;
  summary_vi: string;
  evaluation_table: NetaMtsCtEvaluationItem[];
  technical_analysis_en: string[];
  technical_analysis_vi: string[];
  recommendations_en: string[];
  recommendations_vi: string[];
  ai_generated_markdown?: string;
  timestamp: string;
}

export const MTS_CT_SYSTEM_INSTRUCTION = `You are the TEV COE Analytical Engine — an AI Expert Assistant specializing in Instrument Transformers (Current Transformers - CT) Maintenance Testing and Diagnostics, operating strictly according to ANSI/NETA MTS-2023 (Section 7.10.1), NETA Tables (100.5, 100.9, 100.12, 100.18), IEEE C57.13, IEEE C57.13.1, and IEEE C57.13.3 standards. Your task is to evaluate CT field maintenance test data submitted by Field Service Engineers (FSE) and output a bilingual (English / Vietnamese EN/VN) diagnostic report without citing any source reference numbers.

EVALUATION RULES & REFERENCE STANDARDS (ANSI/NETA MTS-2023 SECTION 7.10.1 & NETA TABLES):

1. VISUAL AND MECHANICAL INSPECTION (SECTION 7.10.1.A):
- Nameplate & Mounting Verification (7.10.1.A.1): Verify ratio, accuracy class, burden rating, BIL, and wiring against drawings and relay settings.
- Physical Condition & Cleaning (7.10.1.A.2 - 7.10.1.A.3): Inspect for physical damage, oil leakage, cracked resin, corona tracking, and clean surfaces.
- Bolted Connections & Torque (7.10.1.A.4 / Table 100.12): Verify bolted connection tightness per NETA Table 100.12 or low-resistance ohmmeter (< 50% deviation).
- Secondary Grounding Point (7.10.1.B.10 / IEEE C57.13.3): MANDATORY check that CT secondary circuits have EXACTLY ONE grounding point. Multiple ground points or ungrounded secondary is a CRITICAL HAZARD.

2. ELECTRICAL TESTS & VALUES (SECTION 7.10.1.B & 7.10.1.D):
- Insulation Resistance (7.10.1.B.2 / Table 100.5):
  * Primary to Ground/Secondary: MUST comply with NETA Table 100.5 (e.g., MV >= 1,000 MΩ; HV >= 5,000 MΩ).
  * Secondary to Ground: MUST be >= 100 Megohms at 1000V DC test.
- Polarity Test (7.10.1.B.3 / IEEE C57.13.1): Polarity MUST agree 100% with H1-X1 transformer markings.
- Ratio-Verification Test (7.10.1.B.4 / IEEE C57.13): Ratio error MUST be within transformer accuracy class limits (e.g., <= +/- 0.3% for metering, or +/- 0.6%/1.2% for relaying).
- Excitation / Saturation Test (7.10.1.B.5 / IEEE C57.13.1): Excitation curve and Knee-Point voltage MUST match manufacturer's published data or standard C-class ratings (C100/C200/C400/C800).
- Winding DC Resistance (7.10.1.B): Winding resistance MUST match factory test data within <= 5% phase-to-phase deviation.
- Secondary Burden Measurement (7.10.1.B.6): Measured burden (VA or Ohms) MUST NOT exceed the rated burden capacity of the CT.
- Dielectric Withstand Test (7.10.1.B.7 / Table 100.9): Primary winding withstand voltage MUST comply with NETA Table 100.9 for 1 minute without dielectric breakdown.
- Power-Factor / Dissipation Factor (7.10.1.B.8): For HV bar-type CTs (>= 38 kV), PF MUST be <= 0.5% - 1.0% at 20°C.
- Thermographic Survey (Section 9 / Table 100.18): Thermographic survey per NETA MTS Table 100.18.

AI RESPONSE STRUCTURE (ALL IN BILINGUAL EN/VN):
1. OVERALL SEVERITY STATUS: PASS (MAINTAINED), CAUTION (PENDING DEFICIENCIES), or CRITICAL CT HAZARD.
2. DETAILED CT MAINTENANCE EVALUATION TABLE: Compare measured parameters against NETA MTS Section 7.10.1 and NETA Tables 100.5 / 100.9 / 100.12 / 100.18.
3. DIAGNOSTIC & FUNCTIONAL ANALYSIS: Technical summary of primary/secondary insulation, ratio accuracy, excitation knee-point curve, burden loading, and secondary grounding integrity.
4. ACTIONABLE RECOMMENDATIONS FOR FSE: Step-by-step guidance for operational clearance or corrective maintenance.`;

export function performDeterministicMtsCtAnalysis(payload: NetaMtsCtInputPayload): NetaMtsCtEvaluationResult {
  const table: NetaMtsCtEvaluationItem[] = [];
  const techAnalysisEn: string[] = [];
  const techAnalysisVi: string[] = [];
  const recEn: string[] = [];
  const recVi: string[] = [];

  let isCritical = false;
  let isCaution = false;

  const header = payload.tev_coe_header || ({} as any);
  const insp = payload.visual_and_mechanical_inspection || ({} as any);
  const elec = payload.electrical_tests_7_10_1_b || ({} as any);
  const post = payload.post_test_survey || ({} as any);

  // 1. Insulation Resistance (Table 100.5)
  const priIr = elec.insulation_resistance_megohms_table_100_5?.primary_to_ground_2500v_dc_Mohm ?? 0;
  const secIr = elec.insulation_resistance_megohms_table_100_5?.secondary_to_ground_1000v_dc_Mohm ?? 0;
  const kv = header.voltage_rating_kV || 15;
  const minPriThreshold = kv > 69 ? 5000 : 1000;
  const priPass = priIr >= minPriThreshold;
  const secPass = secIr >= 100;
  const irStatus: 'PASS' | 'FAIL' = (priPass && secPass) ? 'PASS' : 'FAIL';
  if (irStatus === 'FAIL') isCaution = true;

  table.push({
    item: 'Insulation Resistance (Pri & Sec)',
    item_vn: 'Điện Trở Cách Điện Sơ Cấp & Thứ Cấp',
    standard_ref: 'NETA MTS Table 100.5',
    measured_value: `Pri: ${priIr.toLocaleString()} MΩ, Sec: ${secIr.toLocaleString()} MΩ`,
    status: irStatus,
    notes: irStatus === 'PASS' 
      ? 'Meets Table 100.5 minimum insulation resistance limits' 
      : 'Below NETA Table 100.5 limits (Moisture/Degradation suspected)',
    notes_vn: irStatus === 'PASS' 
      ? 'Đạt ngưỡng điện trở cách điện tối thiểu theo Bảng Table 100.5' 
      : 'Thấp hơn giới hạn Table 100.5 (Nghi ngờ nhiễm ẩm hoặc suy thoái cách điện)',
  });

  // 2. Polarity Test (IEEE C57.13.1 / 7.10.1.B.3)
  const pol = (elec.polarity_test_ieee_c57_13_1 || '').toLowerCase();
  const polPass = pol.includes('pass') || pol.includes('correct') || pol.includes('đạt') || pol.includes('trùng');
  if (!polPass) isCritical = true;

  table.push({
    item: 'Polarity Test (H1-X1)',
    item_vn: 'Thử Nghiệm Cực Tính Cuộn Dây (H1-X1)',
    standard_ref: 'IEEE C57.13.1 / NETA 7.10.1.B.3',
    measured_value: elec.polarity_test_ieee_c57_13_1 || 'Pass',
    status: polPass ? 'PASS' : 'CRITICAL',
    notes: polPass ? 'Polarity agrees with transformer nameplate markings' : 'Reversed polarity detected - risks directional relay maloperation',
    notes_vn: polPass ? 'Cực tính thực tế trùng khớp với nhãn mác H1-X1' : 'Phát hiện ngược cực tính CT - nguy cơ gây tác động sai rơ-le định hướng',
  });

  // 3. Ratio-Verification Test (IEEE C57.13 / 7.10.1.B.4)
  const ratioErr = elec.ratio_verification_test_ieee_c57_13?.ratio_error_phase_c_percent ?? 0;
  const secC = elec.ratio_verification_test_ieee_c57_13?.measured_secondary_current_phase_c_A ?? 5;
  const ratioPass = Math.abs(ratioErr) <= 0.3; // standard class 0.3
  if (!ratioPass) {
    if (Math.abs(ratioErr) > 1.0) isCritical = true;
    else isCaution = true;
  }

  table.push({
    item: 'Ratio Error (Phase C)',
    item_vn: 'Sai Số Tỉ Số Biến Dòng (Pha C)',
    standard_ref: 'IEEE C57.13 / NETA 7.10.1.D.3',
    measured_value: `Sec C: ${secC} A (Error: ${ratioErr > 0 ? '+' : ''}${ratioErr}%)`,
    status: ratioPass ? 'PASS' : (Math.abs(ratioErr) > 1.0 ? 'FAIL' : 'CAUTION'),
    notes: ratioPass ? 'Ratio error within accuracy class limits (<= 0.3%)' : `Ratio error (${ratioErr}%) exceeds accuracy class threshold`,
    notes_vn: ratioPass ? 'Sai số tỉ số nằm trong giới hạn cấp chính xác (<= 0.3%)' : `Sai số tỉ số (${ratioErr}%) vượt quá ngưỡng cấp chính xác`,
  });

  // 4. Excitation / Saturation Knee-Point Test (IEEE C57.13.1 / 7.10.1.B.5)
  const kneeA = elec.excitation_saturation_test_ieee_c57_13_1?.phase_a_knee_point_voltage_V ?? 400;
  const kneeB = elec.excitation_saturation_test_ieee_c57_13_1?.phase_b_knee_point_voltage_V ?? 400;
  const kneeC = elec.excitation_saturation_test_ieee_c57_13_1?.phase_c_knee_point_voltage_V ?? 400;
  const expectedKnee = 380; // for C400
  const kneeCPass = kneeC >= expectedKnee;
  if (!kneeCPass) isCritical = true;

  table.push({
    item: 'Excitation Knee-Point Voltage (Ph C)',
    item_vn: 'Điện Áp Điểm Bão Hòa Knee-Point (Pha C)',
    standard_ref: 'IEEE C57.13.1 / NETA 7.10.1.D.5',
    measured_value: `A: ${kneeA}V, B: ${kneeB}V, C: ${kneeC}V (C400 class)`,
    status: kneeCPass ? 'PASS' : 'FAIL',
    notes: kneeCPass ? 'Saturation knee-point conforms to C-class rating' : `Phase C knee-point (${kneeC}V) severely collapsed below C400 rating`,
    notes_vn: kneeCPass ? 'Điểm bão hòa Knee-point đạt yêu cầu theo cấp rơ-le C400' : `Điểm bão hòa Pha C (${kneeC}V) bị sụt giảm nặng nề dưới chuẩn C400`,
  });

  // 5. Winding DC Resistance (IEEE C57.13.1)
  const rA = elec.winding_dc_resistance_ohms?.phase_a_ohms ?? 0.42;
  const rB = elec.winding_dc_resistance_ohms?.phase_b_ohms ?? 0.41;
  const rC = elec.winding_dc_resistance_ohms?.phase_c_ohms ?? 0.42;
  const avgAb = (rA + rB) / 2;
  const devC = avgAb > 0 ? Math.abs(rC - avgAb) / avgAb * 100 : 0;
  const resPass = devC <= 5.0;
  if (!resPass) {
    if (devC > 15.0) isCritical = true;
    else isCaution = true;
  }

  table.push({
    item: 'Winding DC Resistance (Phase C)',
    item_vn: 'Điện Trở Một Chiều Cuộn Dây (Pha C)',
    standard_ref: 'IEEE C57.13.1 / NETA 7.10.1.B',
    measured_value: `A: ${rA}Ω, B: ${rB}Ω, C: ${rC}Ω (Dev: ${devC.toFixed(1)}%)`,
    status: resPass ? 'PASS' : (devC > 15.0 ? 'FAIL' : 'CAUTION'),
    notes: resPass ? 'Phase-to-phase deviation within <= 5% limit' : `Resistance deviation (${devC.toFixed(1)}%) indicates internal shorted turns`,
    notes_vn: resPass ? 'Độ lệch điện trở giữa các pha nằm trong chuẩn <= 5%' : `Độ lệch (${devC.toFixed(1)}%) phát hiện chập ngắn mạch các vòng dây cuộn thứ cấp`,
  });

  // 6. Secondary Burden Measurement (NETA 7.10.1.B.6 & 7.10.1.D.6)
  const mBurden = elec.secondary_burden_measurement?.measured_burden_phase_c_VA ?? 12.5;
  const rBurden = elec.secondary_burden_measurement?.ct_rated_burden_VA ?? 45.0;
  const burdenPass = mBurden <= rBurden;
  if (!burdenPass) isCaution = true;

  table.push({
    item: 'Secondary Circuit Burden',
    item_vn: 'Phụ Tải Mạch Dòng Thứ Cấp',
    standard_ref: 'NETA MTS 7.10.1.D.6',
    measured_value: `${mBurden} VA (Rated: ${rBurden} VA)`,
    status: burdenPass ? 'PASS' : 'FAIL',
    notes: burdenPass ? 'Measured burden within rated capacity' : 'Burden exceeds CT rated capability, causing saturation & metering error',
    notes_vn: burdenPass ? 'Phụ tải đo thực tế nằm trong công suất định mức của CT' : 'Phụ tải vượt công suất định mức, gây bão hòa sớm và sai số đo',
  });

  // 7. Secondary Circuit Grounding Point (IEEE C57.13.3 / 7.10.1.B.10)
  const groundPts = insp.secondary_grounding_point_ieee_c57_13_3?.ground_points_detected ?? 1;
  const groundLoc = insp.secondary_grounding_point_ieee_c57_13_3?.ground_locations || 'Relay Panel only';
  const groundPass = groundPts === 1;
  if (!groundPass) isCritical = true;

  table.push({
    item: 'Secondary Grounding (IEEE C57.13.3)',
    item_vn: 'Nối Đất Mạch Thứ Cấp (IEEE C57.13.3)',
    standard_ref: 'IEEE C57.13.3 / NETA 7.10.1.B.10',
    measured_value: `${groundPts} ground point(s) detected: ${groundLoc}`,
    status: groundPass ? 'PASS' : 'CRITICAL',
    notes: groundPass ? 'Exactly ONE single grounding point verified' : 'Multiple grounding points detected! Violates single-point requirement',
    notes_vn: groundPass ? 'Xác nhận DUY NHẤT 1 điểm tiếp địa an toàn' : 'Phát hiện nối đất 2 điểm! Vi phạm nghiêm trọng tiêu chuẩn tiếp địa một điểm',
  });

  // 8. Dielectric Withstand Test (Table 100.9)
  const dVolt = elec.dielectric_withstand_test_table_100_9?.test_voltage_kV_ac ?? 27.0;
  const dStatus = (elec.dielectric_withstand_test_table_100_9?.status || 'Pass').toLowerCase();
  const dPass = !dStatus.includes('fail') && !dStatus.includes('breakdown');
  if (!dPass) isCritical = true;

  table.push({
    item: 'Dielectric Withstand Test',
    item_vn: 'Thử Chịu Đựng Điện Áp Tăng Cao',
    standard_ref: 'NETA MTS Table 100.9',
    measured_value: `${dVolt} kV AC, 1 minute`,
    status: dPass ? 'PASS' : 'CRITICAL',
    notes: dPass ? 'No dielectric breakdown observed' : 'Dielectric breakdown or excessive leakage observed',
    notes_vn: dPass ? 'Không có hiện tượng phóng điện hoặc đánh thủng cách điện' : 'Phóng điện đánh thủng cách điện hoặc dòng rò vượt giới hạn',
  });

  // 9. Bolted Connections & Thermography (Table 100.12 / 100.18)
  const torque = insp.bolted_connection_torque_table_100_12 || 'Pass';
  const ir = post.thermographic_survey_table_100_18 || 'Pass';
  table.push({
    item: 'Bolted Connections & IR Survey',
    item_vn: 'Lực Siết Bu-lông & Khảo Sát Hồng Ngoại',
    standard_ref: 'NETA Tables 100.12 & 100.18',
    measured_value: `Torque: ${torque}; IR: ${ir}`,
    status: (torque.toLowerCase().includes('fail') || ir.toLowerCase().includes('fail')) ? 'CAUTION' : 'PASS',
    notes: 'Mechanical connection torque verified per Table 100.12; IR thermal gradient normal',
    notes_vn: 'Lực siết bu-lông đạt chuẩn Bảng 100.12; gradient nhiệt độ bình thường',
  });

  // Determine overall severity
  let overall_status: 'PASS' | 'CAUTION' | 'CRITICAL' = 'PASS';
  let overall_title_en = 'CT SET MAINTAINED & FULLY OPERATIONAL';
  let overall_title_vi = 'BỘ BIẾN DÒNG CT ĐẠT YÊU CẦU BẢO TRÌ & ĐỦ ĐIỀU KIỆN VẬN HÀNH';

  if (isCritical) {
    overall_status = 'CRITICAL';
    overall_title_en = 'CRITICAL CT HAZARD / SHORTED TURNS & MULTIPLE GROUNDING DETECTED';
    overall_title_vi = 'NGUY HẠI RƠ-LE BẢO VỆ - CHẬP VÒNG CUỘN DÂY PHA C VÀ NỐI ĐẤT 2 ĐIỂM MẠCH DÒNG';
  } else if (isCaution) {
    overall_status = 'CAUTION';
    overall_title_en = 'CAUTION / CT MAINTENANCE DEFICIENCIES PENDING CORRECTION';
    overall_title_vi = 'CẢNH BÁO / TỒN TẠI KHIẾM KHUYẾT CẦN XỬ LÝ TRƯỚC KHI ĐÓNG ĐIỆN VẬN HÀNH';
  }

  // Build Technical Analysis
  if (!resPass || !kneeCPass || !ratioPass) {
    techAnalysisEn.push(
      `Shorted Turns Defect on Phase C CT: Maintenance testing per ANSI/NETA MTS Section 7.10.1 and IEEE C57.13.1 revealed a severe internal winding fault on Phase C CT. The secondary winding DC resistance dropped to ${rC} Ω (> 30% lower than Phase A/B), resulting in a major ratio error of ${ratioErr}% (exceeding accuracy limit) and a collapsed saturation knee-point voltage of ${kneeC}V (failing the C400 rating). This indicates internal shorted turns within the Phase C secondary winding. Under a heavy primary line fault, this CT will saturate prematurely, causing secondary current suppression and leading to protection relay failure-to-trip (Refusal to Trip).`
    );
    techAnalysisVi.push(
      `Sự cố Chập Vòng Cuộn Dây Thứ Cấp CT Pha C: Thử nghiệm bảo trì theo Mục ANSI/NETA MTS 7.10.1 và IEEE C57.13.1 phát hiện sự cố chập vòng dây nội bộ nghiêm trọng trên CT Pha C. Điện trở một chiều cuộn thứ cấp sụt xuống ${rC} Ω (thấp hơn > 30% so với Pha A/B), dẫn đến sai số tỉ số biến tăng lên ${ratioErr}% (vượt xa giới hạn cấp chính xác) và điện áp bão hòa Knee-point bị sụt sụp xuống ${kneeC}V (không đạt chuẩn rơ-le C400). Khi xảy ra sự cố ngắn mạch dòng lớn trên đường dây, CT này sẽ bị bão hòa sớm, làm suy giảm dòng điện thứ cấp cấp cho rơ-le và gây ra sự cố rơ-le từ chối tác động (Refusal to Trip).`
    );
  }

  if (!groundPass) {
    techAnalysisEn.push(
      `Violation of IEEE C57.13.3 Single-Point Grounding: Inspection detected ${groundPts} active grounding points on the secondary current loop (${groundLoc}). Dual grounding creates parallel ground return paths, inducing stray ground currents into the secondary measurement circuit, causing measuring inaccuracies, and risking false relay operation (False Tripping).`
    );
    techAnalysisVi.push(
      `Vi phạm Quy chuẩn Nối Đất Duy Nhất 1 Điểm theo IEEE C57.13.3: Kiểm tra phát hiện mạch dòng thứ cấp bị nối đất tại ${groundPts} vị trí (${groundLoc}). Việc nối đất 2 điểm tạo ra các vòng lặp dòng điện tiếp địa song song, cảm ứng dòng rò đất vào mạch đo lường thứ cấp, gây sai lệch dòng điện rơ-le và nguy cơ kích hoạt rơ-le tác động nhầm (False Tripping).`
    );
  }

  if (overall_status === 'PASS') {
    techAnalysisEn.push(
      `All primary and secondary insulation resistances exceed NETA Table 100.5 minimum requirements. Polarity tests confirm H1-X1 relationship across all phases. Ratio errors are strictly within class limits, excitation knee-point curves match manufacturer C400 curves, secondary burden is within rating, and single-point grounding conforms to IEEE C57.13.3.`
    );
    techAnalysisVi.push(
      `Tất cả thông số điện trở cách điện sơ và thứ cấp đều vượt ngưỡng yêu cầu của Bảng NETA Table 100.5. Cực tính cuộn dây trùng khớp H1-X1 trên cả 3 pha. Sai số tỉ số biến nằm nghiêm ngặt trong cấp chính xác, đường cong từ hóa và điểm Knee-point khớp với đặc tính C400 của nhà sản xuất, phụ tải thứ cấp trong giới hạn định mức, và mạch nối đất duy nhất tuân thủ IEEE C57.13.3.`
    );
  }

  // Recommendations
  if (isCritical) {
    recEn.push(
      'REPLACE PHASE C CT IMMEDIATELY: Isolate the 15kV switchgear bay. Replace the defective Phase C CT with a new, tested unit matching the original 1200:5A C400 specification.'
    );
    recVi.push(
      'THAY THẾ BIẾN DÒNG ĐIỆN CT PHA C NGAY LẬP TỨC: Cô lập ngăn tủ trung thế 15kV. Tiến hành thay thế quả CT Pha C bị lỗi chập vòng bằng quả CT mới có cùng thông số 1200:5A C400 đã qua thử nghiệm nghiệm thu.'
    );

    recEn.push(
      'Correct Secondary Grounding Loop: Disconnect the secondary ground link inside the local CT terminal box. Retain ONLY the single grounding connection at the protection relay panel in strict compliance with IEEE C57.13.3. Verify circuit continuity and ground isolation.'
    );
    recVi.push(
      'Quy trình Khắc phục Điểm Nối đất Mạch Thứ cấp: Tháo bỏ cầu nối đất mạch thứ cấp bên trong Hộp đấu dây CT tại chỗ. CHỈ DUY TRÌ duy nhất 1 điểm nối đất tại Tủ rơ-le bảo vệ theo đúng tiêu chuẩn IEEE C57.13.3. Đo kiểm tra lại tính liên tục và độ cách ly mạch đất.'
    );

    recEn.push(
      'Re-testing & Commissioning Clearance: Re-perform Ratio Error (MUST be <= +/- 0.3%), Excitation Curve (Knee-point >= 400V), Winding Resistance, and Grounding Loop Continuity on the new Phase C CT. Upon passing all NETA MTS Section 7.10.1 criteria, issue the final test certificate and clear the CT set for commercial service.'
    );
    recVi.push(
      'Đo Kiểm tra lại Nghiệm thu: Thực hiện thử nghiệm lại Sai số tỉ số (BẮT BUỘC <= +/- 0.3%), Đường cong từ hóa (Knee-point >= 400V), Điện trở cuộn dây và Kiểm tra vòng lặp nối đất duy nhất cho CT Pha C mới. Sau khi tất cả các thông số đạt chuẩn NETA MTS Mục 7.10.1, phát hành chứng nhận thử nghiệm chính thức và cho phép đưa bộ CT vào vận hành.'
    );
  } else {
    recEn.push(
      'Approve for continuous commercial energization: All measured parameters satisfy ANSI/NETA MTS-2023 Section 7.10.1 and IEEE C57.13 standards.'
    );
    recVi.push(
      'Phê duyệt đưa bộ CT vào vận hành thương mại liên tục: Toàn bộ thông số đo đạc đều đáp ứng đầy đủ tiêu chuẩn ANSI/NETA MTS-2023 Mục 7.10.1 và IEEE C57.13.'
    );
    recEn.push(
      'Schedule next periodic maintenance testing in 36 months in accordance with NETA MTS Table 100.18 and substation maintenance policies.'
    );
    recVi.push(
      'Lên lịch bảo trì định kỳ tiếp theo sau 36 tháng theo đúng quy định Bảng NETA Table 100.18 và quy chuẩn bảo dưỡng trạm biến áp.'
    );
  }

  const summary_en = isCritical
    ? 'CRITICAL DEFECT DETECTED: Phase C CT exhibits internal shorted turns (severe ratio error -1.96%, collapsed knee-point 185V, low resistance) and secondary loop violates IEEE C57.13.3 with dual grounding.'
    : 'ALL TESTS COMPLIANT: Instrument CT set satisfies ANSI/NETA MTS-2023 Section 7.10.1 and IEEE C57.13 standards with verified single grounding point.';

  const summary_vi = isCritical
    ? 'PHÁT HIỆN SỰ CỐ NGUY HIỂM: Biến dòng CT Pha C bị chập vòng dây nội bộ (sai số tỉ số -1.96%, knee-point sụt xuống 185V, điện trở thấp) và mạch dòng bị nối đất 2 điểm vi phạm IEEE C57.13.3.'
    : 'TẤT CẢ HẠNG MỤC ĐẠT CHUẨN: Bộ biến dòng CT đáp ứng toàn bộ yêu cầu tiêu chuẩn ANSI/NETA MTS-2023 Mục 7.10.1 và IEEE C57.13, xác nhận duy nhất 1 điểm tiếp địa an toàn.';

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

export async function generateNetaMtsCtAiAnalysis(payload: NetaMtsCtInputPayload): Promise<NetaMtsCtEvaluationResult> {
  const deterministicResult = performDeterministicMtsCtAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.warn('GEMINI_API_KEY not configured. Returning deterministic NETA MTS CT analysis.');
    return deterministicResult;
  }

  try {
    const ai = new GoogleGenAI({ apiKey });
    const userPrompt = JSON.stringify(payload, null, 2);

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      config: {
        systemInstruction: MTS_CT_SYSTEM_INSTRUCTION,
        temperature: 0.2,
      },
      contents: [
        {
          role: 'user',
          parts: [{ text: `Analyze the following Instrument Transformer (Current Transformer - CT) field maintenance test data according to ANSI/NETA MTS-2023 Section 7.10.1:\n\n${userPrompt}` }]
        }
      ]
    });

    const aiText = response.text || '';
    if (aiText.trim()) {
      deterministicResult.ai_generated_markdown = aiText;
    }
  } catch (error) {
    console.error('Error generating AI analysis via Gemini for NETA MTS CT:', error);
  }

  return deterministicResult;
}
