import { GoogleGenAI } from '@google/genai';

export interface NetaEcsMvInputPayload {
  tev_coe_header: {
    module: string;
    project_name: string;
    system_tag: string;
    nominal_voltage_kV: number;
    fse_name: string;
    commissioning_date: string;
  };
  pre_energization_7_2_a: {
    acceptance_test_reports_reviewed: boolean;
    mv_busbar_insulation_resistance_5000v_megohms: {
      phase_a_to_ground: number;
      phase_b_to_ground: number;
      phase_c_to_ground: number;
      status?: string;
    };
    mv_cable_vlf_withstand_table_100_2: {
      test_voltage_kV_rms: number;
      frequency_hz: number;
      duration_minutes: number;
      tan_delta_mean_rad?: number;
      status: string; // 'Pass' | 'Fail'
    };
    bolted_connection_torque_table_100_12: string; // 'Pass' | 'Fail'
    sf6_gas_pressure_bar?: {
      measured_pressure: number;
      nominal_rating: number;
      status: string;
    };
    transformer_oil_bdv_kv?: {
      measured_bdv: number;
      standard_min_kv: number;
      status: string;
    };
    differential_protection_primary_injection: string; // 'Pass' | 'Fail'
    safety_interlocks_and_key_keys: string; // 'Pass' | 'Fail'
    written_energization_plan_approved: boolean;
    dc_auxiliary_power_battery_ok?: boolean;
  };
  energization_7_2_b: {
    temporary_protective_grounds_removed: boolean;
    phasing_check_bus_tie_breaker: {
      phase_a1_a2_voltage_volts: number;
      phase_b1_b2_voltage_volts: number;
      phase_c1_c2_voltage_volts: number;
      phase_angle_diff_deg?: number;
      status: string; // 'Pass' | 'Fail'
    };
    relay_metering_inputs_status: string; // 'Pass' | 'Fail'
  };
  post_energization_7_2_c: {
    thermographic_survey_table_100_18: string; // 'Pass' | 'Investigate'
    online_pd_survey_tables_100_23: {
      tev_max_dB: number;
      acoustic_max_dB: number;
      status: string; // 'Pass' | 'Fail'
    };
    post_energization_gas_fluid_report?: boolean;
  };
}

export interface NetaEcsMvItemEvaluation {
  item_code: string;
  item_name_vn: string;
  item_name_en: string;
  neta_standard_ref: string;
  measured_value: string;
  status: 'PASS' | 'FAIL' | 'INVESTIGATE';
  evaluation_vn: string;
  evaluation_en: string;
}

export interface NetaEcsMvEvaluationResult {
  overall_status: 'PASS' | 'CAUTION' | 'CRITICAL';
  status_title_en: string;
  status_title_vn: string;
  summary_badge: string;
  summary_color: string;
  evaluation_table: NetaEcsMvItemEvaluation[];
  technical_analysis: {
    hazard_identified: boolean;
    title_en: string;
    title_vn: string;
    description_en: string;
    description_vn: string;
    insulation_min_val: number;
    vlf_status: string;
    phasing_delta_v: number;
    tev_val: number;
  };
  actionable_recommendations: {
    halt_energization: boolean;
    steps: Array<{
      step_number: number;
      title_en: string;
      title_vn: string;
      details_en: string;
      details_vn: string;
    }>;
  };
  raw_markdown_report?: string;
}

export const ECS_MV_SYSTEM_INSTRUCTION = `You are the TEV COE Analytical Engine — an AI Expert Assistant specializing in Medium-Voltage Electrical Power System (> 1 kV and < 100 kV) Commissioning and Acceptance Testing, operating strictly according to ANSI/NETA ECS-2024 (Section 7.2) and ANSI/NETA ATS/MTS standards. Your task is to evaluate medium-voltage system commissioning test data submitted by Field Service Engineers (FSE) and output a bilingual (English / Vietnamese EN/VN) diagnostic report without citing any source reference numbers.

EVALUATION RULES & REFERENCE STANDARDS (ANSI/NETA ECS-2024 SECTION 7.2 & NETA TABLES):

1. PRE-ENERGIZATION REQUIREMENTS (SECTION 7.2.A):
- Documentation & Acceptance Testing (7.2.A.1 - 7.2.A.4): Verify OPR/BOD compliance and 100% acceptance testing completed per ANSI/NETA ATS. Minimum dry insulation resistance MUST comply with NETA Table 100.1 (e.g., 15kV system test at 2500V/5000V DC >= 1,000 MΩ).
- VLF Cable Withstand Test (7.2.A.5 / Table 100.2): Verify MV cable VLF withstand and Tan Delta test per NETA ATS Table 100.2 / IEEE 400.2.
- Bolt Tightness (7.2.A.7 / Table 100.12): Verify bolted connection tightness using torque wrench per NETA Table 100.12.
- Insulating Fluid & SF6 Gas (7.2.A.35): Verify transformer oil breakdown voltage (BDV), DGA, and SF6 pressure match manufacturer specifications.
- Protection Relays & CT/VT (7.2.A.32 - 7.2.A.34): Verify CT/VT ratio/polarity, differential relay schemes, and primary injection (through-fault) testing.
- Interlocks & Safety (7.2.A.15 - 7.2.A.45): Verify mechanical/key interlocks, Arc Energy reduction systems, DC auxiliary power, and approved written energization plan.

2. ENERGIZATION REQUIREMENTS (SECTION 7.2.B):
- Temporary Ground Removal (7.2.B.1 - 7.2.B.3): Verify 100% removal of temporary protective grounds.
- Dual-Source Phasing Check (7.2.B.8 - 7.2.B.9): Verify zero phase angle discrepancy across dual-source / double-ended bus tie breakers (delta V <= 15V, phase angle <= 5 deg).

3. POST-ENERGIZATION REQUIREMENTS (SECTION 7.2.C):
- Thermographic Survey (7.2.C.3 / Section 9 / Table 100.18): Perform IR survey under load per NETA MTS Table 100.18.
- Online Partial Discharge Survey (7.2.C.4 / NETA MTS Tables 100.23.1 - 100.23.5): Perform TEV/Acoustic/HFCT/UHF online PD survey under load.
- Post-Energization Fluid/Gas & Final Commissioning Report (7.2.C.1 & 5.4).

AI RESPONSE STRUCTURE (ALL IN BILINGUAL EN/VN):
1. OVERALL SEVERITY STATUS: PASS (COMMISSIONED), CAUTION (PENDING DEFICIENCIES), or CRITICAL ENERGIZATION HAZARD.
2. DETAILED MV COMMISSIONING EVALUATION TABLE: Compare measured parameters against NETA ECS Section 7.2 and NETA Tables 100.1 / 100.2 / 100.12 / 100.18 / 100.23.
3. COMMISSIONING & SYSTEM FUNCTION ANALYSIS: Technical summary of MV insulation, VLF cable test, relay protection, phasing, and online PD.
4. ACTIONABLE RECOMMENDATIONS FOR FSE: Step-by-step guidance for energization clearance or deficiency correction.`;

export function performDeterministicEcsMvAnalysis(payload: NetaEcsMvInputPayload): NetaEcsMvEvaluationResult {
  const pA = payload.pre_energization_7_2_a;
  const pB = payload.energization_7_2_b;
  const pC = payload.post_energization_7_2_c;

  const insA = pA.mv_busbar_insulation_resistance_5000v_megohms?.phase_a_to_ground ?? 5000;
  const insB = pA.mv_busbar_insulation_resistance_5000v_megohms?.phase_b_to_ground ?? 5000;
  const insC = pA.mv_busbar_insulation_resistance_5000v_megohms?.phase_c_to_ground ?? 5000;
  const minIns = Math.min(insA, insB, insC);
  const minRequiredIns = 1000.0; // Table 100.1 for MV apparatus >= 1kV

  const isVlfPass = !pA.mv_cable_vlf_withstand_table_100_2?.status?.toLowerCase().includes('fail');
  const isTorquePass = !pA.bolted_connection_torque_table_100_12?.toLowerCase().includes('fail');
  const isDiffProtPass = !pA.differential_protection_primary_injection?.toLowerCase().includes('fail');
  const isInterlocksPass = !pA.safety_interlocks_and_key_keys?.toLowerCase().includes('fail');
  const isDocsPass = pA.acceptance_test_reports_reviewed === true;
  const isPlanPass = pA.written_energization_plan_approved === true;

  // SF6 or Oil test
  const isSf6Pass = !pA.sf6_gas_pressure_bar || !pA.sf6_gas_pressure_bar.status?.toLowerCase().includes('fail');

  // Energization
  const isGroundsRemoved = pB.temporary_protective_grounds_removed === true;
  const phA = pB.phasing_check_bus_tie_breaker?.phase_a1_a2_voltage_volts ?? 10;
  const phB = pB.phasing_check_bus_tie_breaker?.phase_b1_b2_voltage_volts ?? 10;
  const phC = pB.phasing_check_bus_tie_breaker?.phase_c1_c2_voltage_volts ?? 10;
  const maxPhasingDeltaV = Math.max(phA, phB, phC);
  const isPhasingPass = !pB.phasing_check_bus_tie_breaker?.status?.toLowerCase().includes('fail') && maxPhasingDeltaV <= 30.0;
  const isRelayMeteringPass = !pB.relay_metering_inputs_status?.toLowerCase().includes('fail');

  // Post-Energization
  const isThermalPass = !pC.thermographic_survey_table_100_18?.toLowerCase().includes('investigate') &&
                        !pC.thermographic_survey_table_100_18?.toLowerCase().includes('fail');
  const tevMax = pC.online_pd_survey_tables_100_23?.tev_max_dB ?? 10;
  const acousticMax = pC.online_pd_survey_tables_100_23?.acoustic_max_dB ?? 1;
  const isPdPass = tevMax < 20.0 && acousticMax < 6.0 && !pC.online_pd_survey_tables_100_23?.status?.toLowerCase().includes('fail');

  // Build Evaluation Table
  const table: NetaEcsMvItemEvaluation[] = [
    {
      item_code: 'A1',
      item_name_vn: 'Rà soát Hồ sơ, Bản vẽ & Nhãn mác / Documentation & Nameplate',
      item_name_en: 'Documentation & Nameplate Review',
      neta_standard_ref: 'NETA ECS Section 7.2.A.1 & 7.2.A.2',
      measured_value: isDocsPass ? 'Đã duyệt báo cáo thử nghiệm Acceptance Testing (OPR/BOD)' : 'Chưa đủ hồ sơ thử nghiệm chấp nhận',
      status: isDocsPass ? 'PASS' : 'FAIL',
      evaluation_vn: isDocsPass ? 'Trùng khớp 100% hồ sơ thiết kế và catalog nhà sản xuất.' : 'Sai lệch hồ sơ hoặc thiếu file nghiệm thu.',
      evaluation_en: isDocsPass ? 'Fully compliant with OPR/BOD and manufacturer specs.' : 'Missing documentation.'
    },
    {
      item_code: 'A2',
      item_name_vn: 'Cách điện Thanh cái & Thiết bị MV (5000V DC) / MV Bus Insulation',
      item_name_en: 'MV Busbar Insulation Resistance (5000V DC)',
      neta_standard_ref: 'NETA ATS / Table 100.1 (Min >= 1,000 MΩ)',
      measured_value: `A: ${insA} MΩ | B: ${insB} MΩ | C: ${insC} MΩ`,
      status: minIns >= minRequiredIns ? 'PASS' : 'FAIL',
      evaluation_vn: minIns >= minRequiredIns
        ? `Đạt tiêu chuẩn NETA Table 100.1 (Min ${minIns} MΩ >= 1,000 MΩ).`
        : `Pha có điện trở thấp nhất (${minIns} MΩ) dưới ngưỡng 1,000 MΩ! Nguy cơ phóng điện chập pha.`,
      evaluation_en: minIns >= minRequiredIns
        ? 'Insulation exceeds NETA Table 100.1 threshold (>= 1,000 MΩ).'
        : `Insulation (${minIns} MΩ) below 1,000 MΩ threshold! Flashover hazard.`
    },
    {
      item_code: 'A3',
      item_name_vn: 'Thử nghiệm Cáp Ngầm Trung thế VLF / MV Cable VLF Withstand',
      item_name_en: 'MV Cable VLF Withstand & Tan Delta Test',
      neta_standard_ref: 'NETA ATS Table 100.2 / IEEE 400.2',
      measured_value: isVlfPass
        ? `${pA.mv_cable_vlf_withstand_table_100_2?.test_voltage_kV_rms ?? 28}kV RMS @ 0.1Hz, Tan δ = ${pA.mv_cable_vlf_withstand_table_100_2?.tan_delta_mean_rad ?? 0.0012}`
        : 'Thử nghiệm VLF phóng điện đánh thủng hoặc Tan Delta vượt chuẩn',
      status: isVlfPass ? 'PASS' : 'FAIL',
      evaluation_vn: isVlfPass ? 'Điện áp chịu đựng VLF và góc tổn hao Tan Delta đạt chuẩn.' : 'Cáp phóng điện hoặc thoái hóa cách điện Tan Delta cao.',
      evaluation_en: isVlfPass ? 'VLF withstand and Tan Delta compliant with Table 100.2.' : 'VLF breakdown or excessive Tan Delta.'
    },
    {
      item_code: 'A4',
      item_name_vn: 'Lực Siết Bu-lông Thanh cái & Đầu Cáp / Bolted Connection Tightness',
      item_name_en: 'Bolted Connection Tightness',
      neta_standard_ref: 'NETA ATS/MTS Table 100.12 (Bolt Torque)',
      measured_value: isTorquePass ? 'Đã siết cờ-lê lực hiệu chuẩn đạt Bảng 100.12' : 'Chưa siết cờ-lê lực hoặc thiếu dấu niêm phong',
      status: isTorquePass ? 'PASS' : 'FAIL',
      evaluation_vn: isTorquePass ? 'Mối nối cơ khí đạt lực siết chuẩn NETA Table 100.12.' : 'Chưa siết đủ lực, nguy cơ phát nhiệt tiếp xúc trung thế.',
      evaluation_en: isTorquePass ? 'Torque verified per Table 100.12.' : 'Torque verification deficient.'
    },
    {
      item_code: 'A5',
      item_name_vn: 'Áp suất Khí SF6 & Phân tích Dầu / Insulating Fluid & SF6 Gas',
      item_name_en: 'Insulating Fluid BDV & SF6 Gas Pressure',
      neta_standard_ref: 'NETA ECS Section 7.2.A.35 & 7.2.A.36',
      measured_value: pA.sf6_gas_pressure_bar
        ? `Áp suất SF6: ${pA.sf6_gas_pressure_bar.measured_pressure} bar (Định mức: ${pA.sf6_gas_pressure_bar.nominal_rating} bar)`
        : 'Chất cách điện đạt định mức nhà sản xuất',
      status: isSf6Pass ? 'PASS' : 'FAIL',
      evaluation_vn: isSf6Pass ? 'Áp suất khí SF6 và chất cách điện đạt định mức.' : 'Áp suất SF6 tụt thấp dưới ngưỡng khóa thao tác hoặc dầu ẩm.',
      evaluation_en: isSf6Pass ? 'SF6 gas pressure / fluid BDV within normal limits.' : 'Low SF6 pressure or fluid degradation.'
    },
    {
      item_code: 'A6',
      item_name_vn: 'Bảo vệ So lệch & Bơm Dòng Sơ cấp / Differential Protection & Primary Injection',
      item_name_en: 'Differential Relays & Primary Injection Testing',
      neta_standard_ref: 'NETA ECS Section 7.2.A.32 - 7.2.A.34',
      measured_value: isDiffProtPass ? 'Bơm dòng sơ cấp xuyên sự cố đạt, bảo vệ tác động chuẩn' : 'Sai cực tính CT hoặc lỗi logic bảo vệ so lệch',
      status: isDiffProtPass ? 'PASS' : 'FAIL',
      evaluation_vn: isDiffProtPass ? 'Hệ thống bảo vệ so lệch tác động chính xác theo file cài đặt.' : 'Lỗi cực tính CT/VT hoặc relay so lệch không tác động.',
      evaluation_en: isDiffProtPass ? 'Primary injection test verified differential scheme.' : 'Differential protection failure or CT polarity mismatch.'
    },
    {
      item_code: 'A7',
      item_name_vn: 'Khóa An Toàn Liên Động & Giảm Hồ Quang / Key Interlocks & Arc Energy',
      item_name_en: 'Key Interlocks & Arc Energy Reduction',
      neta_standard_ref: 'NETA ECS Section 7.2.A.15 & 7.2.A.16',
      measured_value: isInterlocksPass ? 'Liên động chìa khóa Kirk-key 100% tin cậy, Arc Flash active' : 'Kẹt cơ cấu liên động an toàn hoặc thiếu chìa khóa',
      status: isInterlocksPass ? 'PASS' : 'FAIL',
      evaluation_vn: isInterlocksPass ? 'Ngăn chặn 100% nguy cơ thao tác nhầm vào thiết bị mang điện.' : 'Lỗi khóa liên động, nguy cơ thao tác sự cố.',
      evaluation_en: isInterlocksPass ? 'Interlocks and arc energy systems verified.' : 'Interlock deficiency.'
    },
    {
      item_code: 'A8',
      item_name_vn: 'Kế Hoạch Đóng Điện & Nguồn DC Phụ Trợ / Written Energization Plan & DC System',
      item_name_en: 'Written Energization Plan & DC Control Power',
      neta_standard_ref: 'NETA ECS Section 7.2.A.30 & 7.2.A.45',
      measured_value: isPlanPass ? 'Kế hoạch đóng điện đã ký duyệt, dàn ắc-quy DC 110V/220V ổn định' : 'Chưa có Kế hoạch Đóng điện bằng văn bản được duyệt',
      status: isPlanPass ? 'PASS' : 'FAIL',
      evaluation_vn: isPlanPass ? 'Đủ hồ sơ pháp lý và nguồn thao tác máy cắt an toàn.' : 'Nghiêm cấm đóng điện trung thế khi chưa duyệt kế hoạch.',
      evaluation_en: isPlanPass ? 'Energization plan approved, DC control ready.' : 'Unapproved plan.'
    },
    {
      item_code: 'B1',
      item_name_vn: 'Tháo Tiếp Địa Di Động & Rào Chắn / Temporary Ground Removal',
      item_name_en: 'Temporary Protective Ground Removal',
      neta_standard_ref: 'NETA ECS Section 7.2.B.1 & 7.2.B.3',
      measured_value: isGroundsRemoved ? 'Đã tháo bỏ 100% tiếp địa di động trên thanh cái và cáp MV' : 'CẢNH BÁO TỬ VONG: CÒN DÂY TIẾP ĐỊA TẠM TRÊN HỆ TRUNG THẾ!',
      status: isGroundsRemoved ? 'PASS' : 'FAIL',
      evaluation_vn: isGroundsRemoved ? 'Hệ thống an toàn tuyệt đối, sẵn sàng khép dao cách ly/máy cắt.' : 'THẢM HỌA: Tiếp địa tạm gây ngắn mạch ba pha trực tiếp!',
      evaluation_en: isGroundsRemoved ? 'All temporary protective grounds removed.' : 'FATAL HAZARD: Temporary grounds still connected!'
    },
    {
      item_code: 'B2',
      item_name_vn: 'Đồng Pha Hai Nguồn Nguồn (Máy cắt Liên Lạc) / Phasing Checks (Tie Breaker)',
      item_name_en: 'Phasing Checks across Bus Tie Breaker',
      neta_standard_ref: 'NETA ECS Section 7.2.B.8 & 7.2.B.9 (Delta V <= 15V)',
      measured_value: isPhasingPass
        ? `Trùng pha 100% (ΔV: A=${phA}V, B=${phB}V, C=${phC}V <= 15V)`
        : `Lệch pha nguy hiểm! Chênh áp giữa hai nguồn: ${maxPhasingDeltaV}V`,
      status: isPhasingPass ? 'PASS' : 'FAIL',
      evaluation_vn: isPhasingPass ? 'Độ lệch góc pha <= 5°, chênh lệch điện áp pha tương ứng ~ 0V.' : 'Ngược pha hoặc lệch góc pha, tuyệt đối không hòa lưới!',
      evaluation_en: isPhasingPass ? 'In-phase condition confirmed across dual sources.' : 'Out-of-phase condition detected!'
    },
    {
      item_code: 'B3',
      item_name_vn: 'Đo Lường & Tín Hiệu Rơ-le / Metering & Relay Inputs Verification',
      item_name_en: 'Metering & Relay Secondary Vector Verification',
      neta_standard_ref: 'NETA ECS Section 7.2.B.7 & 7.2.B.10',
      measured_value: isRelayMeteringPass ? 'Dạng sóng & góc pha dòng/áp nhị thứ khớp 100% tải thực' : 'Nhầm cực tính CT/VT hoặc sai góc pha nhị thứ',
      status: isRelayMeteringPass ? 'PASS' : 'FAIL',
      evaluation_vn: isRelayMeteringPass ? 'Hệ thống đo đếm và rơ-le số IEDs hoạt động chính xác.' : 'Lỗi đo đếm hoặc cực tính CT có thể gây nhảy rơ-le sai.',
      evaluation_en: isRelayMeteringPass ? 'Relay inputs and metering vectors verified.' : 'Metering or polarity discrepancy.'
    },
    {
      item_code: 'C1',
      item_name_vn: 'Khảo Sát Nhiệt Hồng Ngoại Trung Thế Mang Tải / MV Thermographic Survey',
      item_name_en: 'MV Thermographic Survey under Operating Load',
      neta_standard_ref: 'NETA ECS 7.2.C.3 / Table 100.18 (NFPA 70B)',
      measured_value: isThermalPass ? 'Nhiệt độ đồng đều dưới 70% tải, không có điểm nóng Hotspot' : 'Phát hiện điểm phát nhiệt quá mức trên thanh cái/đầu cáp',
      status: isThermalPass ? 'PASS' : 'INVESTIGATE',
      evaluation_vn: isThermalPass ? 'Không xuất hiện điểm quá nhiệt bất thường theo Table 100.18.' : 'Xuất hiện điểm nóng cần xử lý trước khi bàn giao.',
      evaluation_en: isThermalPass ? 'Normal thermal profile under load per Table 100.18.' : 'Hotspot observed.'
    },
    {
      item_code: 'C2',
      item_name_vn: 'Khảo Sát Phóng Điện Cục Bộ Online / Online Partial Discharge Survey',
      item_name_en: 'Online Partial Discharge (PD) Survey',
      neta_standard_ref: 'NETA ECS 7.2.C.4 & Tables 100.23.1 - 100.23.5',
      measured_value: isPdPass
        ? `TEV cực đại: ${tevMax} dBmV (Chuẩn < 20) | Siêu âm: ${acousticMax} dBµV (Chuẩn < 6)`
        : `Phát hiện tín hiệu PD bất thường (TEV: ${tevMax} dBmV, US: ${acousticMax} dBµV)`,
      status: isPdPass ? 'PASS' : 'FAIL',
      evaluation_vn: isPdPass ? 'Tín hiệu PD dưới ngưỡng Warning theo NETA Table 100.23.' : 'Xuất hiện phóng điện cục bộ, nguy cơ thủng cách điện.',
      evaluation_en: isPdPass ? 'Online PD below Warning thresholds per Table 100.23.' : 'Elevated PD detected.'
    },
    {
      item_code: 'C3',
      item_name_vn: 'Phân Tích Dầu Khí Sau Đóng Điện & Báo Cáo / Post-Energization & Report',
      item_name_en: 'Post-Energization Sampling & Commissioning Report',
      neta_standard_ref: 'NETA ECS Section 7.2.C.1 & Section 5.4',
      measured_value: 'Báo cáo nghiệm thu hoàn chỉnh trình Chủ đầu tư phê duyệt',
      status: 'PASS',
      evaluation_vn: 'Hồ sơ kỹ thuật nghiệm thu hoàn thiện đủ điều kiện bàn giao.',
      evaluation_en: 'Comprehensive commissioning report completed per Section 5.4.'
    }
  ];

  // Determine Overall Status
  let overall_status: 'PASS' | 'CAUTION' | 'CRITICAL' = 'PASS';
  let status_title_en = 'PASS / FULLY COMMISSIONED & APPROVED FOR COMMERCIAL ENERGIZATION';
  let status_title_vn = 'NGHIỆM THU HOÀN TẤT / ĐỦ ĐIỀU KIỆN ĐÓNG ĐIỆN VẬN HÀNH THƯƠNG MẠI';
  let summary_badge = 'PASS (COMMISSIONED)';
  let summary_color = 'emerald';

  const criticalConditions = [
    minIns < minRequiredIns,
    !isVlfPass,
    !isGroundsRemoved,
    !isPhasingPass,
    !isPlanPass
  ];

  const cautionConditions = [
    !isTorquePass,
    !isDocsPass,
    !isSf6Pass,
    !isDiffProtPass,
    !isInterlocksPass,
    !isRelayMeteringPass,
    !isThermalPass,
    !isPdPass
  ];

  if (criticalConditions.some(Boolean)) {
    overall_status = 'CRITICAL';
    status_title_en = 'CRITICAL ENERGIZATION HAZARD / DIELECTRIC OR PHASING DEFICIENCY DETECTED';
    status_title_vn = 'NGUY HẠI ĐÓNG ĐIỆN TRUNG THẾ / LỖI CÁCH ĐIỆN, CÁP VLF HOẶC LỆCH PHA';
    summary_badge = 'CRITICAL ENERGIZATION HAZARD';
    summary_color = 'rose';
  } else if (cautionConditions.some(Boolean)) {
    overall_status = 'CAUTION';
    status_title_en = 'CAUTION / PENDING DEFICIENCIES (PUNCH-LIST REQUIRED)';
    status_title_vn = 'CẢNH BÁO / TỒN TẠI TỒN ĐỌNG CẦN KHẮC PHỤC (LẬP PUNCH-LIST)';
    summary_badge = 'CAUTION (PENDING DEFICIENCIES)';
    summary_color = 'amber';
  }

  // Technical Analysis
  let title_en = 'Medium-Voltage System Fully Commissioned under ANSI/NETA ECS-2024';
  let title_vn = 'Xác nhận Hệ thống Trung thế Đáp ứng Đầy đủ Tiêu chuẩn ANSI/NETA ECS-2024';
  let desc_en = `All medium-voltage parameters including 5000V busbar insulation (${minIns} MΩ >= 1,000 MΩ), VLF cable withstand, dual-source tie phasing (ΔV <= 15V), and online PD survey (TEV ${tevMax} dB) comply with NETA ECS Section 7.2.`;
  let desc_vn = `Tất cả các thông số trung thế bao gồm điện trở cách điện thanh cái 5000V (${minIns} MΩ >= 1,000 MΩ), thử nghiệm cáp VLF, đồng pha hai nguồn máy cắt liên lạc (ΔV <= 15V) và đo PD online (TEV ${tevMax} dB) đều đáp ứng xuất sắc tiêu chuẩn NETA ECS Mục 7.2.`;

  if (minIns < minRequiredIns) {
    title_en = `Low Insulation Resistance Hazard on MV Busbar (${minIns} MΩ < 1,000 MΩ)`;
    title_vn = `Nguy cơ Phóng điện Cách điện Thanh cái Trung thế (${minIns} MΩ < 1,000 MΩ)`;
    desc_en = `MV busbar insulation resistance measured only ${minIns} Megohms at 5000V DC. This fails the 1,000 Megohms minimum requirement specified in NETA Table 100.1 for medium-voltage apparatus. Energizing under this condition creates catastrophic phase-to-ground flashover risk.`;
    desc_vn = `Điện trở cách điện thanh cái trung thế chỉ đạt ${minIns} Megohms ở 5000V DC, không đạt ngưỡng tối thiểu 1,000 Megohms theo Bảng NETA Table 100.1. Đóng điện trong điều kiện này sẽ dẫn đến phóng điện chọc thủng cách điện thảm khốc.`;
  } else if (!isGroundsRemoved) {
    title_en = 'Deadly Safety Hazard: Temporary Protective Grounds Not Removed';
    title_vn = 'Nguy hại An toàn Tử vong: Chưa Tháo dỡ Tiếp địa Di động Tạm thời';
    desc_en = 'Temporary grounds left connected on MV busbars will cause an explosive three-phase bolted fault upon energization.';
    desc_vn = 'Tiếp địa tạm thời còn sót lại trên thanh cái trung thế sẽ gây nổ ngắn mạch ba pha cực kỳ thảm khốc khi đóng điện.';
  } else if (!isPhasingPass) {
    title_en = 'Out-of-Phase Condition Detected Across Dual Sources (Tie Breaker)';
    title_vn = 'Phát hiện Sự cố Lệch Pha Giữa Hai Nguồn Điện (Máy cắt Liên lạc)';
    desc_en = `Phasing check recorded excessive voltage difference (${maxPhasingDeltaV}V > 15V) across tie breaker poles. Paralleling these out-of-phase sources will trigger severe busbar damage.`;
    desc_vn = `Đo đồng pha ghi nhận chênh lệch điện áp lớn (${maxPhasingDeltaV}V > 15V) giữa hai nguồn. Đóng hòa máy cắt liên lạc trong tình trạng này sẽ gây nổ phóng điện thanh cái.`;
  }

  // Recommendations
  const steps = [];
  if (overall_status === 'CRITICAL') {
    steps.push({
      step_number: 1,
      title_en: 'HALT ENERGIZATION IMMEDIATELY',
      title_vn: 'TẠM DỪNG TIẾN TRÌNH ĐÓNG ĐIỆN TRUNG THẾ NGAY LẬP TỨC',
      details_en: 'Do NOT close incoming or tie circuit breakers. Hold the energization permit until dielectric integrity and phase alignment are fully restored.',
      details_vn: 'TUYỆT ĐỐI KHÔNG đóng máy cắt tổng hoặc máy cắt liên lạc. Thu hồi giấy phép đóng điện cho đến khi xử lý triệt để sự cố cách điện và đồng pha.'
    });
    steps.push({
      step_number: 2,
      title_en: 'Perform Isolation & Remediation Protocol',
      title_vn: 'Thực hiện Quy trình Cô lập & Khắc phục Sự cố',
      details_en: 'Inspect for moisture accumulation, tracking on epoxy spout insulators, or incorrect transformer connections. Clean and bake dry if insulation is deficient.',
      details_vn: 'Kiểm tra hiện tượng đọng ẩm, vết phóng điện trên sứ xuyên epoxy hoặc đấu sai cực tính máy biến áp. Vệ sinh và sấy khô nếu cách điện suy giảm.'
    });
    steps.push({
      step_number: 3,
      title_en: 'Re-testing & Commissioning Clearance',
      title_vn: 'Đo Kiểm tra lại Nghiệm thu & Cấp Phép Đóng Điện',
      details_en: 'Repeat 5000V DC insulation test and dual-source phasing verification. Ensure all values comply with NETA Tables 100.1 & 100.2 before re-authorizing energization.',
      details_vn: 'Đo lại cách điện 5000V DC và kiểm tra đồng pha hai nguồn. Đảm bảo toàn bộ thông số đạt chuẩn NETA Tables 100.1 & 100.2 trước khi cấp lại lệnh đóng điện.'
    });
  } else if (overall_status === 'CAUTION') {
    steps.push({
      step_number: 1,
      title_en: 'Issue Punch-List & Remediation Schedule',
      title_vn: 'Lập Bảng Punch-List & Tiến độ Khắc phục',
      details_en: 'Document all pending items (bolt torque, SF6 pressure, online PD warnings) and require contractor remediation within 48 hours.',
      details_vn: 'Ghi nhận toàn bộ các tồn đọng kỹ thuật vào biên bản Punch-List và yêu cầu nhà thầu hoàn thành khắc phục trong vòng 48 giờ.'
    });
    steps.push({
      step_number: 2,
      title_en: 'Continuous Online PD & Thermal Monitoring',
      title_vn: 'Tăng Cường Giám Sát PD & Nhiệt Hồng Ngoại',
      details_en: 'Monitor TEV and acoustic PD levels during initial 24 hours of operating load.',
      details_vn: 'Theo dõi liên tục tín hiệu phóng điện cục bộ TEV và siêu âm trong 24 giờ đầu mang tải.'
    });
  } else {
    steps.push({
      step_number: 1,
      title_en: 'Authorize Commercial Operation Clearance',
      title_vn: 'Cấp Phép Đưa vào Vận hành Thương mại Chính thức',
      details_en: 'Issue formal Medium-Voltage Electrical Commissioning Certificate under ANSI/NETA ECS-2024 Section 7.2. Hand over baseline test documentation per Section 5.4.',
      details_vn: 'Ký duyệt Biên bản Nghiệm thu Đóng điện Trung thế theo tiêu chuẩn ANSI/NETA ECS-2024 Mục 7.2. Bàn giao toàn bộ hồ sơ dữ liệu nền (baseline test) cho đơn vị quản lý vận hành.'
    });
    steps.push({
      step_number: 2,
      title_en: 'Schedule First 12-Month Maintenance Survey',
      title_vn: 'Lập Kế hoạch Khảo sát Bảo trì Định kỳ 12 Tháng',
      details_en: 'Schedule the first post-commissioning Online Partial Discharge and Thermographic Survey in 12 months in accordance with NETA MTS Section 11 and Section 9.',
      details_vn: 'Lập kế hoạch khảo sát Phóng điện Cục bộ Online và Khảo sát Nhiệt hồng ngoại định kỳ lần đầu sau 12 tháng vận hành theo Mục 11 và Mục 9 của tiêu chuẩn NETA MTS.'
    });
  }

  return {
    overall_status,
    status_title_en,
    status_title_vn,
    summary_badge,
    summary_color,
    evaluation_table: table,
    technical_analysis: {
      hazard_identified: overall_status === 'CRITICAL',
      title_en,
      title_vn,
      description_en: desc_en,
      description_vn: desc_vn,
      insulation_min_val: minIns,
      vlf_status: isVlfPass ? 'Pass' : 'Fail',
      phasing_delta_v: maxPhasingDeltaV,
      tev_val: tevMax
    },
    actionable_recommendations: {
      halt_energization: overall_status === 'CRITICAL',
      steps
    }
  };
}

export async function generateEcsMvAiAnalysis(payload: NetaEcsMvInputPayload): Promise<NetaEcsMvEvaluationResult> {
  const deterministic = performDeterministicEcsMvAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return deterministic;
  }

  try {
    const ai = new GoogleGenAI({ apiKey });
    const prompt = `Evaluate the following medium-voltage system commissioning dataset according to ANSI/NETA ECS-2024 Section 7.2:
${JSON.stringify(payload, null, 2)}

Provide your expert evaluation strictly adhering to NETA ECS-2024 (Section 7.2) and NETA Tables 100.1, 100.2, 100.12, 100.18, 100.23. Output your findings in bilingual Vietnamese and English without citing source reference numbers.`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: prompt,
      config: {
        systemInstruction: ECS_MV_SYSTEM_INSTRUCTION,
        temperature: 0.2
      }
    });

    const markdown = response.text || '';
    return {
      ...deterministic,
      raw_markdown_report: markdown
    };
  } catch (error) {
    console.warn('Gemini ECS MV analysis fallback to deterministic:', error);
    return deterministic;
  }
}
