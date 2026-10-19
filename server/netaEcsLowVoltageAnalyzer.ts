import { GoogleGenAI } from '@google/genai';

export interface NetaEcsLvInputPayload {
  tev_coe_header: {
    module: string;
    project_name: string;
    system_tag: string;
    voltage_rating_V: number;
    fse_name: string;
    commissioning_date: string;
  };
  pre_energization_7_1_a: {
    acceptance_test_reports_reviewed: boolean;
    busbar_insulation_resistance_1000v_megohms: {
      phase_a_to_ground: number;
      phase_b_to_ground: number;
      phase_c_to_ground: number;
      status?: string;
    };
    bolted_connection_torque_table_100_12: string; // 'Pass' | 'Fail'
    grounding_resistance_ohms: number;
    safety_interlocks_and_estop: string; // 'Pass' | 'Fail'
    written_energization_plan_approved: boolean;
    dc_control_power_battery_ready?: boolean;
  };
  energization_7_1_b: {
    temporary_protective_grounds_removed: boolean;
    safety_barricades_access_restricted?: boolean;
    phase_sequence_check: string; // 'Pass' | 'Fail'
    motor_rotation_check: string; // 'Pass' | 'Fail'
    alarms_and_metering_status: string; // 'Pass' | 'Fail'
  };
  post_energization_7_1_c: {
    thermographic_survey_table_100_18: string; // 'Pass' | 'Investigate'
    phase_load_balance_amps: {
      phase_a: number;
      phase_b: number;
      phase_c: number;
    };
    gfci_receptacles_tested: string; // 'Pass' | 'Fail'
  };
}

export interface NetaEcsLvItemEvaluation {
  item_code: string;
  item_name_vn: string;
  item_name_en: string;
  neta_standard_ref: string;
  measured_value: string;
  status: 'PASS' | 'FAIL' | 'INVESTIGATE';
  evaluation_vn: string;
  evaluation_en: string;
}

export interface NetaEcsLvEvaluationResult {
  overall_status: 'PASS' | 'CAUTION' | 'CRITICAL';
  status_title_en: string;
  status_title_vn: string;
  summary_badge: string;
  summary_color: string;
  evaluation_table: NetaEcsLvItemEvaluation[];
  technical_analysis: {
    hazard_identified: boolean;
    title_en: string;
    title_vn: string;
    description_en: string;
    description_vn: string;
    insulation_min_val: number;
    grounding_val: number;
    load_unbalance_pct: number;
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

export const ECS_LV_SYSTEM_INSTRUCTION = `You are the TEV COE Analytical Engine — an AI Expert Assistant specializing in Low-Voltage Electrical Power System (< 1 kV) Commissioning and Acceptance Testing, operating strictly according to ANSI/NETA ECS-2024 (Section 7.1) and ANSI/NETA ATS/MTS standards. Your task is to evaluate low-voltage system commissioning test data submitted by Field Service Engineers (FSE) and output a bilingual (English / Vietnamese EN/VN) diagnostic report without citing any source reference numbers.

EVALUATION RULES & REFERENCE STANDARDS (ANSI/NETA ECS-2024 SECTION 7.1 & NETA TABLES):

1. PRE-ENERGIZATION REQUIREMENTS (SECTION 7.1.A):
- Documentation & Acceptance Testing (7.1.A.1 - 7.1.A.4): Verify OPR/BOD compliance and 100% acceptance testing completed per ANSI/NETA ATS. Minimum dry insulation resistance MUST comply with NETA Table 100.1 (>= 1.0 MΩ for < 1kV apparatus).
- Bolt Tightness (7.1.A.7 / Table 100.12): Verify bolted connection tightness using torque wrench per NETA Table 100.12.
- Grounding & Bonding (Section 7.13 / IEEE 81): Grounding electrode system resistance MUST be <= 5.0 Ohms.
- Safety & Written Energization Plan (7.1.A.30 - 7.1.A.45): Verify interlocks, E-Stop, battery charge, and written energization plan approval.

2. ENERGIZATION REQUIREMENTS (SECTION 7.1.B):
- Safety & Temporary Ground Removal (7.1.B.1 - 7.1.B.3): Verify temporary grounds removed and access restricted.
- Phasing & Rotation (7.1.B.8 - 7.1.B.11): Verify A-B-C phase sequence and correct motor rotation.

3. POST-ENERGIZATION REQUIREMENTS (SECTION 7.1.C):
- Thermographic Survey (7.1.C.3 / Section 9 / Table 100.18): Perform IR survey under load.
- Cân bằng tải & Power Quality (7.1.C.5 - 7.1.C.7): Verify balanced phase loading and power quality survey.
- GFCI Receptacle Test (7.1.C.10): Verify 100% GFCI trip operation.

AI RESPONSE STRUCTURE (ALL IN BILINGUAL EN/VN):
1. OVERALL SEVERITY STATUS: PASS (COMMISSIONED), CAUTION (PENDING DEFICIENCIES), or CRITICAL ENERGIZATION HAZARD.
2. DETAILED LV COMMISSIONING EVALUATION TABLE: Compare measured parameters against NETA ECS Section 7.1 and NETA Tables 100.1 / 100.12 / 100.18.
3. COMMISSIONING & SYSTEM FUNCTION ANALYSIS: Technical summary of insulation, grounding, phase sequence, and thermographic results.
4. ACTIONABLE RECOMMENDATIONS FOR FSE: Step-by-step guidance for energization clearance or deficiency correction.`;

export function performDeterministicEcsLvAnalysis(payload: NetaEcsLvInputPayload): NetaEcsLvEvaluationResult {
  const pA = payload.pre_energization_7_1_a;
  const pB = payload.energization_7_1_b;
  const pC = payload.post_energization_7_1_c;

  const insA = pA.busbar_insulation_resistance_1000v_megohms?.phase_a_to_ground ?? 100;
  const insB = pA.busbar_insulation_resistance_1000v_megohms?.phase_b_to_ground ?? 100;
  const insC = pA.busbar_insulation_resistance_1000v_megohms?.phase_c_to_ground ?? 100;
  const minIns = Math.min(insA, insB, insC);

  const groundingR = pA.grounding_resistance_ohms ?? 1.0;
  const isTorquePass = !pA.bolted_connection_torque_table_100_12?.toLowerCase().includes('fail');
  const isInterlocksPass = !pA.safety_interlocks_and_estop?.toLowerCase().includes('fail');
  const isDocsPass = pA.acceptance_test_reports_reviewed === true;
  const isPlanPass = pA.written_energization_plan_approved === true;

  const isGroundsRemoved = pB.temporary_protective_grounds_removed === true;
  const isPhaseSeqPass = !pB.phase_sequence_check?.toLowerCase().includes('fail');
  const isMotorRotPass = !pB.motor_rotation_check?.toLowerCase().includes('fail');
  const isAlarmPass = !pB.alarms_and_metering_status?.toLowerCase().includes('fail');

  const isThermalPass = !pC.thermographic_survey_table_100_18?.toLowerCase().includes('investigate') &&
                        !pC.thermographic_survey_table_100_18?.toLowerCase().includes('fail');
  const isGfciPass = !pC.gfci_receptacles_tested?.toLowerCase().includes('fail');

  // Load Unbalance Calculation
  const lA = pC.phase_load_balance_amps?.phase_a ?? 100;
  const lB = pC.phase_load_balance_amps?.phase_b ?? 100;
  const lC = pC.phase_load_balance_amps?.phase_c ?? 100;
  const avgLoad = (lA + lB + lC) / 3;
  const maxDev = Math.max(Math.abs(lA - avgLoad), Math.abs(lB - avgLoad), Math.abs(lC - avgLoad));
  const unbalancePct = avgLoad > 0 ? (maxDev / avgLoad) * 100 : 0;
  const isLoadBalancePass = unbalancePct <= 15.0;

  // Build Evaluation Table
  const table: NetaEcsLvItemEvaluation[] = [
    {
      item_code: 'A1',
      item_name_vn: 'Rà soát Hồ sơ & Nhãn mác / Documentation & Nameplate Review',
      item_name_en: 'Documentation & Nameplate Review',
      neta_standard_ref: 'NETA ECS Section 7.1.A.1 - 7.1.A.6',
      measured_value: isDocsPass ? 'Đã duyệt báo cáo thử nghiệm Acceptance Testing (OPR/BOD)' : 'Chưa đủ hồ sơ thử nghiệm chấp nhận',
      status: isDocsPass ? 'PASS' : 'FAIL',
      evaluation_vn: isDocsPass ? 'Khớp 100% hồ sơ thiết kế và thông số định mức.' : 'Thiếu file báo cáo nghiệm thu chấp nhận.',
      evaluation_en: isDocsPass ? 'Fully compliant with project requirements (OPR/BOD).' : 'Missing acceptance test documentation.'
    },
    {
      item_code: 'A2',
      item_name_vn: 'Điện trở Cách điện Thanh cái / Busbar Insulation Resistance (1000V DC)',
      item_name_en: 'Busbar Insulation Resistance (1000V DC)',
      neta_standard_ref: 'NETA ATS / Table 100.1 (Min >= 1.0 MΩ)',
      measured_value: `Pha A: ${insA} MΩ | Pha B: ${insB} MΩ | Pha C: ${insC} MΩ`,
      status: minIns >= 1.0 ? 'PASS' : 'FAIL',
      evaluation_vn: minIns >= 1.0
        ? 'Tất cả các pha đạt chuẩn NETA Table 100.1 (Min >= 1.0 MΩ).'
        : `Pha có điện trở thấp nhất (${minIns} MΩ) dưới ngưỡng cho phép 1.0 MΩ! Nguy cơ ngắn mạch.`,
      evaluation_en: minIns >= 1.0
        ? 'All phases meet Table 100.1 minimum requirement (>= 1.0 MΩ).'
        : `Phase insulation (${minIns} MΩ) is below the 1.0 MΩ threshold! Flashover hazard.`
    },
    {
      item_code: 'A3',
      item_name_vn: 'Lực Siết Mối Nối Bu-lông / Bolted Connection Tightness',
      item_name_en: 'Bolted Connection Tightness',
      neta_standard_ref: 'NETA ATS/MTS Table 100.12 (Bolt Torque)',
      measured_value: isTorquePass ? 'Đã siết cờ-lê lực đạt tiêu chuẩn Table 100.12' : 'Chưa kiểm tra lực siết theo Bảng 100.12',
      status: isTorquePass ? 'PASS' : 'FAIL',
      evaluation_vn: isTorquePass ? 'Mối nối cơ khí đạt lực siết chuẩn NETA Table 100.12.' : 'Chưa siết lực đủ tiêu chuẩn, nguy cơ phát nhiệt tiếp xúc.',
      evaluation_en: isTorquePass ? 'Bolt torque verified per Table 100.12 standard.' : 'Torque verification missing or failed.'
    },
    {
      item_code: 'A4',
      item_name_vn: 'Điện trở Tiếp địa Cọc Đất / Grounding & Bonding Resistance',
      item_name_en: 'Grounding & Bonding Resistance',
      neta_standard_ref: 'NETA ATS Section 7.13 / IEEE 81 (Max <= 5.0 Ω)',
      measured_value: `${groundingR} Ohm`,
      status: groundingR <= 5.0 ? 'PASS' : 'FAIL',
      evaluation_vn: groundingR <= 5.0
        ? `Đạt tiêu chuẩn an toàn tiếp địa (<= 5.0 Ω).`
        : `Vượt ngưỡng 5.0 Ω (${groundingR} Ω). Nguy hiểm mất an toàn chạm vỏ!`,
      evaluation_en: groundingR <= 5.0
        ? 'Compliant with NETA Section 7.13 / IEEE 81 (<= 5.0 Ohms).'
        : `Exceeds 5.0 Ohms threshold (${groundingR} Ohms). Ground fault safety hazard.`
    },
    {
      item_code: 'A5',
      item_name_vn: 'Khóa An toàn, Liên động & E-Stop / Safety Interlocks & E-Stop',
      item_name_en: 'Safety Interlocks & E-Stop Chain',
      neta_standard_ref: 'NETA ECS Section 7.1.A.30 & 7.1.A.32',
      measured_value: isInterlocksPass ? 'Liên động cơ/điện tác động tốt, E-Stop ngắt máy cắt tổng' : 'Lỗi cơ cấu liên động hoặc E-Stop không tác động',
      status: isInterlocksPass ? 'PASS' : 'FAIL',
      evaluation_vn: isInterlocksPass ? 'Hệ thống liên động và ngắt khẩn cấp hoạt động chính xác.' : 'Lỗi khóa an toàn cơ/điện hoặc E-Stop.',
      evaluation_en: isInterlocksPass ? 'Interlocks and emergency trip operate correctly.' : 'Safety interlock or E-stop failure.'
    },
    {
      item_code: 'A6',
      item_name_vn: 'Kế hoạch Đóng điện & Nguồn Phụ trợ / Written Energization Plan',
      item_name_en: 'Written Energization Plan & Control Power',
      neta_standard_ref: 'NETA ECS Section 7.1.A.41 & 7.1.A.45',
      measured_value: isPlanPass ? 'Kế hoạch đóng điện đã được phê duyệt bằng văn bản' : 'Chưa có Kế hoạch Đóng điện bằng văn bản',
      status: isPlanPass ? 'PASS' : 'FAIL',
      evaluation_vn: isPlanPass ? 'Hồ sơ pháp lý tiền đóng điện đã được phê duyệt.' : 'Nghiêm cấm đóng điện khi chưa có kế hoạch được duyệt.',
      evaluation_en: isPlanPass ? 'Written energization plan signed and approved.' : 'Energization prohibited without approved plan.'
    },
    {
      item_code: 'B1',
      item_name_vn: 'An toàn & Tháo Tiếp địa Di động / Temporary Protective Grounds Removal',
      item_name_en: 'Temporary Protective Grounds Removal',
      neta_standard_ref: 'NETA ECS Section 7.1.B.1 & 7.1.B.2',
      measured_value: isGroundsRemoved ? 'Đã tháo bỏ 100% tiếp địa di động, rào chắn an toàn' : 'CẢNH BÁO: Còn tiếp địa di động trên thanh cái/tủ!',
      status: isGroundsRemoved ? 'PASS' : 'FAIL',
      evaluation_vn: isGroundsRemoved ? 'Khu vực an toàn, sẵn sàng đóng điện.' : 'CỰC KỲ NGUY HIỂM: Tiếp địa tạm chưa tháo gây ngắn mạch trực tiếp!',
      evaluation_en: isGroundsRemoved ? 'All temporary grounds removed and access controlled.' : 'DEADLY HAZARD: Temporary grounds still connected!'
    },
    {
      item_code: 'B2',
      item_name_vn: 'Thứ tự Pha & Góc Pha / Phase Sequence & Phasing',
      item_name_en: 'Phase Sequence & Phasing',
      neta_standard_ref: 'NETA ECS Section 7.1.B.8 & 7.1.B.9',
      measured_value: isPhaseSeqPass ? 'Thứ tự pha chuẩn A-B-C (U-V-W)' : 'Ngược thứ tự pha (A-C-B / U-W-V)',
      status: isPhaseSeqPass ? 'PASS' : 'FAIL',
      evaluation_vn: isPhaseSeqPass ? 'Khớp thứ tự pha chuẩn so với thiết kế.' : 'Lỗi ngược thứ tự pha, nguy cơ đảo chiều động cơ toàn nhà máy.',
      evaluation_en: isPhaseSeqPass ? 'Phase sequence A-B-C verified correct.' : 'Reverse phase sequence detected.'
    },
    {
      item_code: 'B3',
      item_name_vn: 'Chiều Quay Động cơ & Tải / Motor Rotation Check',
      item_name_en: 'Motor Rotation Check',
      neta_standard_ref: 'NETA ECS Section 7.1.B.10 & 7.1.B.11',
      measured_value: isMotorRotPass ? 'Động cơ quay đúng 100% chiều mũi tên chỉ dẫn' : 'Phát hiện động cơ quay ngược chiều',
      status: isMotorRotPass ? 'PASS' : 'FAIL',
      evaluation_vn: isMotorRotPass ? 'Chiều quay động cơ chuẩn xác.' : 'Động cơ quay ngược, cần hoán đổi cặp pha.',
      evaluation_en: isMotorRotPass ? 'Motor rotation matches direction arrows.' : 'Reverse motor rotation detected.'
    },
    {
      item_code: 'B4',
      item_name_vn: 'Cảnh báo & Đo lường / Alarms & Metering Check',
      item_name_en: 'Alarms & Metering Calibration',
      neta_standard_ref: 'NETA ECS Section 7.1.B.7 & 7.1.B.13',
      measured_value: isAlarmPass ? '0 cảnh báo lỗi active, đồng hồ đo áp/dòng chuẩn xác' : 'Còn cảnh báo sự cố Fault/Alarm trên tủ điện',
      status: isAlarmPass ? 'PASS' : 'FAIL',
      evaluation_vn: isAlarmPass ? 'Hệ thống đo đếm và bảo vệ ổn định.' : 'Còn lỗi chưa reset hoặc sai lệch đồng hồ đo.',
      evaluation_en: isAlarmPass ? 'Zero active alarms, metering values accurate.' : 'Active alarms or metering discrepancy.'
    },
    {
      item_code: 'C1',
      item_name_vn: 'Khảo sát Nhiệt Hồng ngoại mang Tải / Post-Energization Thermographic Survey',
      item_name_en: 'Post-Energization Thermographic Survey',
      neta_standard_ref: 'NETA ECS 7.1.C.3 / Table 100.18 (NFPA 70B)',
      measured_value: isThermalPass ? 'Nhiệt độ phân bố đồng đều dưới tải, không có điểm nóng' : 'Phát hiện điểm phát nhiệt Hotspot bất thường (Cần điều tra)',
      status: isThermalPass ? 'PASS' : 'INVESTIGATE',
      evaluation_vn: isThermalPass ? 'Không xuất hiện điểm quá nhiệt bất thường.' : 'Xuất hiện điểm nóng bất thường theo Bảng 100.18.',
      evaluation_en: isThermalPass ? 'Normal temperature profile under operating load.' : 'Abnormal thermal pattern detected per Table 100.18.'
    },
    {
      item_code: 'C2',
      item_name_vn: 'Cân bằng Tải & Chất lượng Điện năng / Load Balance & Power Quality',
      item_name_en: 'Phase Load Balance & Power Quality',
      neta_standard_ref: 'NETA ECS Section 7.1.C.6 & 7.1.C.7 (Unbalance <= 15%)',
      measured_value: `A: ${lA}A | B: ${lB}A | C: ${lC}A (Độ lệch: ${unbalancePct.toFixed(1)}%)`,
      status: isLoadBalancePass ? 'PASS' : 'FAIL',
      evaluation_vn: isLoadBalancePass
        ? `Tải cân bằng tốt giữa các pha (Độ lệch <= 15%).`
        : `Mất cân bằng pha (${unbalancePct.toFixed(1)}% > 15%). Cần phân bổ lại phụ tải.`,
      evaluation_en: isLoadBalancePass
        ? 'Phase loading balanced within acceptable limit (<= 15%).'
        : `Phase unbalance (${unbalancePct.toFixed(1)}% > 15%). Load redistribution required.`
    },
    {
      item_code: 'C3',
      item_name_vn: 'Thử nghiệm Ổ cắm Chống giật GFCI / GFCI Receptacles Test',
      item_name_en: 'GFCI Receptacles Functional Test',
      neta_standard_ref: 'NETA ECS Section 7.1.C.10',
      measured_value: isGfciPass ? '100% ổ cắm GFCI nhảy chính xác khi thử dòng rò' : 'Có ổ cắm GFCI không nhảy khi tạo dòng rò',
      status: isGfciPass ? 'PASS' : 'FAIL',
      evaluation_vn: isGfciPass ? 'Bảo vệ chống dòng rò hạ thế hoạt động hoàn hảo.' : 'Lỗi chức năng bảo vệ chống giật rò điện.',
      evaluation_en: isGfciPass ? '100% of GFCI receptacles tripped correctly.' : 'GFCI failure detected.'
    }
  ];

  // Determine Overall Status
  let overall_status: 'PASS' | 'CAUTION' | 'CRITICAL' = 'PASS';
  let status_title_en = 'COMMISSIONED & ACCEPTED / READY FOR FULL COMMERCIAL OPERATION';
  let status_title_vn = 'NGHIỆM THU ĐẠT CHUẨN / ĐỦ ĐIỀU KIỆN ĐÓNG ĐIỆN VẬN HÀNH THƯƠNG MẠI';
  let summary_badge = 'PASS (COMMISSIONED)';
  let summary_color = 'emerald';

  const criticalConditions = [
    minIns < 1.0,
    !isGroundsRemoved,
    !isPhaseSeqPass,
    groundingR > 5.0,
    !isPlanPass
  ];

  const cautionConditions = [
    !isTorquePass,
    !isDocsPass,
    !isInterlocksPass,
    !isMotorRotPass,
    !isAlarmPass,
    !isThermalPass,
    !isLoadBalancePass,
    !isGfciPass
  ];

  if (criticalConditions.some(Boolean)) {
    overall_status = 'CRITICAL';
    status_title_en = 'CRITICAL ENERGIZATION HAZARD / INSULATION OR SAFETY DEFICIENCY DETECTED';
    status_title_vn = 'NGUY HẠI ĐÓNG ĐIỆN NGHIÊM TRỌNG / LỖI SUY GIẢM CÁCH ĐIỆN HOẶC MẤT AN TOÀN';
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
  const hasLowInsulation = minIns < 1.0;
  let title_en = 'Full System Compliance Confirmed under ANSI/NETA ECS-2024';
  let title_vn = 'Xác nhận Hệ thống Đáp ứng Đầy đủ Tiêu chuẩn ANSI/NETA ECS-2024';
  let desc_en = `All low-voltage parameters including insulation resistance (${minIns} MΩ >= 1.0 MΩ), grounding electrode resistance (${groundingR} Ω <= 5.0 Ω), bolted connection torque, and load balance (${unbalancePct.toFixed(1)}%) meet NETA ECS requirements.`;
  let desc_vn = `Tất cả thông số hạ thế bao gồm điện trở cách điện (${minIns} MΩ >= 1.0 MΩ), điện trở tiếp địa (${groundingR} Ω <= 5.0 Ω), lực siết bu-lông và độ cân bằng tải (${unbalancePct.toFixed(1)}%) đều đạt chuẩn nghiệm thu NETA ECS.`;

  if (hasLowInsulation) {
    const lowPhase = insA < 1.0 ? 'Phase A' : insB < 1.0 ? 'Phase B' : 'Phase C';
    const lowPhaseVn = insA < 1.0 ? 'Pha A' : insB < 1.0 ? 'Pha B' : 'Pha C';
    title_en = `Low Insulation Resistance Hazard on ${lowPhase} (${minIns} MΩ)`;
    title_vn = `Nguy cơ Suy giảm Điện trở Cách điện Nghiêm trọng trên ${lowPhaseVn} (${minIns} Megohm)`;
    desc_en = `During pre-energization testing under NETA ECS Section 7.1.A.3, the insulation resistance for ${lowPhase} on ${payload.tev_coe_header.system_tag} measured only ${minIns} Megohm at 1000V DC. This fails the 1.0 Megohm minimum threshold specified in NETA Table 100.1 for low-voltage apparatus. Energizing the switchboard under this condition creates a severe short-circuit or phase-to-ground flashover hazard.`;
    desc_vn = `Trong quá trình kiểm tra tiền đóng điện theo Mục 7.1.A.3 của NETA ECS, điện trở cách điện ${lowPhaseVn} của ${payload.tev_coe_header.system_tag} chỉ đo được ${minIns} Megohm ở điện áp 1000V DC. Giá trị này không đạt ngưỡng tối thiểu 1.0 Megohm quy định tại Bảng NETA Table 100.1 dành cho thiết bị hạ thế. Việc cho phép đóng điện trong tình trạng này nguy cơ trực tiếp gây ngắn mạch hoặc phóng điện chạm đất thảm khốc.`;
  } else if (!isGroundsRemoved) {
    title_en = 'Deadly Safety Hazard: Temporary Protective Grounds Not Removed';
    title_vn = 'Nguy hại An toàn Tử vong: Chưa Tháo dỡ Tiếp địa Di động Tạm thời';
    desc_en = 'Temporary grounds left connected on the busbar will cause a catastrophic zero-impedance phase-to-ground bolted fault upon energization.';
    desc_vn = 'Tiếp địa tạm thời còn sót lại trên thanh cái sẽ gây sự cố ngắn mạch trực tiếp cực kỳ thảm khốc khi đóng điện.';
  } else if (!isPhaseSeqPass) {
    title_en = 'Reverse Phase Sequence Detected';
    title_vn = 'Phát hiện Sự cố Ngược Thứ tự Pha Hệ thống';
    desc_en = 'Reverse phase sequence will cause all three-phase AC induction motors to rotate in the opposite direction, risking severe mechanical equipment destruction.';
    desc_vn = 'Ngược thứ tự pha sẽ khiến toàn bộ động cơ không đồng bộ 3 pha quay ngược chiều, nguy cơ phá hỏng cơ cấu máy bơm và quạt công nghiệp.';
  }

  // Recommendations
  const steps = [];
  if (overall_status === 'CRITICAL') {
    steps.push({
      step_number: 1,
      title_en: 'HALT ENERGIZATION IMMEDIATELY',
      title_vn: 'TẠM DỪNG TIẾN TRÌNH ĐÓNG ĐIỆN NGAY LẬP TỨC',
      details_en: 'Do NOT close the main incoming circuit breaker. Issue an immediate hold on the written energization plan until all critical deficiencies are fully resolved.',
      details_vn: 'TUYỆT ĐỐI KHÔNG đóng máy cắt tổng hạ thế. Hoãn thực thi Kế hoạch Đóng điện cho đến khi khắc phục dứt điểm mọi nguy cơ.'
    });

    if (hasLowInsulation) {
      steps.push({
        step_number: 2,
        title_en: 'Insulation Troubleshooting & Isolation Protocol',
        title_vn: 'Quy trình Khoanh vùng & Khắc phục Suy giảm Cách điện',
        details_en: 'Disconnect secondary feeders and busway sections to isolate the low insulation point. Inspect for moisture accumulation, conductive dust, or damaged insulating support insulators along the busbar run. Clean insulators using electrical solvent and dry the cubicle.',
        details_vn: 'Tháo rời các nhánh cáp xuất tuyến và đoạn thanh cái dẫn để khoanh vùng điểm rò điện. Kiểm tra hiện tượng đọng sương/ẩm ướt, bụi kim loại hoặc nứt vỡ sứ đỡ thanh cái. Vệ sinh cách điện bằng dung dịch chuyên dụng và sấy khô khoang tủ.'
      });
      steps.push({
        step_number: 3,
        title_en: 'Re-testing & Commissioning Clearance',
        title_vn: 'Đo Kiểm tra lại Nghiệm thu & Cấp Phép Đóng Điện',
        details_en: 'Re-measure insulation resistance for all phases to ground. Ensure the value exceeds 1.0 Megohm (preferably > 50 Megohms for new switchboards) before re-issuing clearance for energization under Section 7.1.B.',
        details_vn: 'Đo lại điện trở cách điện các pha so với đất. Đảm bảo giá trị vượt ngưỡng 1.0 Megohm (khuyến nghị > 50 Megohms đối với tủ điện mới) trước khi cấp lại giấy phép đóng điện chính thức theo Mục 7.1.B.'
      });
    }
  } else if (overall_status === 'CAUTION') {
    steps.push({
      step_number: 1,
      title_en: 'Issue Punch-List & Remediation Schedule',
      title_vn: 'Lập Bảng Punch-List & Tiến độ Khắc phục',
      details_en: 'Document all pending non-critical items (bolted connection torque, load balance, GFCI trip) and require contractor remediation within 48 hours.',
      details_vn: 'Ghi nhận toàn bộ các tồn đọng kỹ thuật vào biên bản Punch-List và yêu cầu nhà thầu hoàn thành khắc phục trong vòng 48 giờ.'
    });
    steps.push({
      step_number: 2,
      title_en: 'Re-verify under Partial Operating Load',
      title_vn: 'Đối soát lại dưới Tải Vận hành Thực tế',
      details_en: 'Perform repeated thermographic scan and load balance measurement after load redistribution.',
      details_vn: 'Tiến hành chụp lại ảnh nhiệt và đo lại độ cân bằng pha sau khi tái phân bổ phụ tải.'
    });
  } else {
    steps.push({
      step_number: 1,
      title_en: 'Authorize Commercial Operation Clearance',
      title_vn: 'Cấp Phép Đưa vào Vận hành Thương mại Chính thức',
      details_en: 'Issue formal Commissioning Certificate under ANSI/NETA ECS-2024 Section 7.1. Archive all baseline acceptance test records into facility CMMS.',
      details_vn: 'Ký duyệt Biên bản Nghiệm thu Đóng điện theo tiêu chuẩn ANSI/NETA ECS-2024 Mục 7.1. Bàn giao toàn bộ dữ liệu cơ sở vào phần mềm quản lý bảo trì CMMS của nhà máy.'
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
      grounding_val: groundingR,
      load_unbalance_pct: Number(unbalancePct.toFixed(1))
    },
    actionable_recommendations: {
      halt_energization: overall_status === 'CRITICAL',
      steps
    }
  };
}

export async function generateEcsLvAiAnalysis(payload: NetaEcsLvInputPayload): Promise<NetaEcsLvEvaluationResult> {
  const deterministic = performDeterministicEcsLvAnalysis(payload);

  // If Gemini API key is available, run prompt enhancement
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return deterministic;
  }

  try {
    const ai = new GoogleGenAI({ apiKey });
    const prompt = `Evaluate the following low-voltage system commissioning dataset according to ANSI/NETA ECS-2024 Section 7.1:
${JSON.stringify(payload, null, 2)}

Provide your expert evaluation strictly adhering to NETA ECS-2024 (Section 7.1) and NETA Table 100.1 / Table 100.12. Output your findings in bilingual Vietnamese and English without citing source reference numbers.`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: prompt,
      config: {
        systemInstruction: ECS_LV_SYSTEM_INSTRUCTION,
        temperature: 0.2
      }
    });

    const markdown = response.text || '';
    return {
      ...deterministic,
      raw_markdown_report: markdown
    };
  } catch (error) {
    console.warn('Gemini ECS LV analysis fallback to deterministic:', error);
    return deterministic;
  }
}
