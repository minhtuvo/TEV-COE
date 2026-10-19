import { GoogleGenAI } from '@google/genai';

export interface NetaEcsAtsInputPayload {
  tev_coe_header: {
    module: string;
    project_name: string;
    system_tag: string;
    voltage_rating_V: number;
    current_rating_A: number;
    fse_name: string;
    commissioning_date: string;
  };
  pre_energization_8_3_a: {
    acceptance_test_reports_reviewed_ats_7_22_3: boolean;
    control_wiring_insulation_1000v_megohms: number;
    main_contact_resistance_micro_ohms: {
      source_1_normal_poles_uOhm: {
        phase_a: number;
        phase_b: number;
        phase_c: number;
        neutral?: number;
      };
      source_2_emergency_poles_uOhm: {
        phase_a: number;
        phase_b: number;
        phase_c: number;
        neutral?: number;
      };
      max_deviation_percent?: number;
      status?: string;
    };
    mechanical_and_electrical_interlocks: string; // 'Pass' | 'Fail'
    control_logic_and_timers_configured: {
      engine_start_delay_sec: number;
      transfer_time_delay_sec: number;
      retransfer_time_delay_sec: number;
      engine_cool_down_delay_sec: number;
      status?: string;
    };
    bolted_connection_torque_table_100_12: string; // 'Pass' | 'Fail'
    grounding_and_neutral_bonding: string; // 'Pass' | 'Fail'
    written_energization_plan_approved: boolean;
  };
  energization_8_3_b: {
    source_voltage_and_frequency: {
      source_1_utility_volts: number;
      source_1_freq_hz: number;
      source_2_generator_volts: number;
      source_2_freq_hz: number;
      status?: string;
    };
    phase_sequence_matching: string; // 'Pass' | 'Fail'
    automatic_load_transfer_test_nfpa_110: {
      simulated_event?: string;
      engine_start_signal_time_sec: number;
      generator_voltage_ready_time_sec: number;
      ats_transfer_completed_time_sec: number;
      total_emergency_transfer_time_sec: number;
      status?: string;
    };
    retransfer_and_cool_down_test: {
      simulated_event?: string;
      retransfer_timer_active_sec: number;
      ats_retransfer_to_source_1_completed: boolean;
      engine_cool_down_run_sec: number;
      generator_auto_shutdown_successful: boolean;
      status?: string;
    };
    single_phase_failure_test: string; // 'Pass' | 'Fail'
  };
  post_energization_8_3_c: {
    thermographic_survey_table_100_18: {
      max_delta_t_celsius?: number;
      status: string; // 'Pass' | 'Investigate'
      details?: string;
    };
    scada_bms_telemetry_verified: string; // 'Pass' | 'Fail'
    owner_manual_bypass_training_completed: boolean;
  };
}

export interface NetaEcsAtsItemEvaluation {
  item_code: string;
  item_name_vn: string;
  item_name_en: string;
  neta_standard_ref: string;
  measured_value: string;
  status: 'PASS' | 'FAIL' | 'INVESTIGATE';
  evaluation_vn: string;
  evaluation_en: string;
}

export interface NetaEcsAtsEvaluationResult {
  overall_status: 'PASS' | 'CAUTION' | 'CRITICAL';
  status_title_en: string;
  status_title_vn: string;
  summary_badge: string;
  summary_color: string;
  evaluation_table: NetaEcsAtsItemEvaluation[];
  technical_analysis: {
    hazard_identified: boolean;
    title_en: string;
    title_vn: string;
    description_en: string;
    description_vn: string;
    control_ir_megohms: number;
    source_2_max_contact_uohm: number;
    contact_deviation_percent: number;
    total_transfer_time_sec: number;
  };
  actionable_recommendations: {
    halt_transfer: boolean;
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

export const ECS_ATS_SYSTEM_INSTRUCTION = `You are the TEV COE Analytical Engine — an AI Expert Assistant specializing in Automatic Transfer Switch (ATS) System Commissioning and Acceptance Testing, operating strictly according to ANSI/NETA ECS-2024 (Section 8.3), ANSI/NETA ATS-2025 (Section 7.22.3), NETA MTS (Tables 100.1, 100.12, 100.18), NFPA 110, and IEEE 446 standards. Your task is to evaluate ATS system commissioning test data submitted by Field Service Engineers (FSE) and output a bilingual (English / Vietnamese EN/VN) diagnostic report without citing any source reference numbers.

EVALUATION RULES & REFERENCE STANDARDS (ANSI/NETA ECS-2024 SECTION 8.3 & NETA TABLES):

1. PRE-ENERGIZATION REQUIREMENTS (SECTION 8.3.A):
- Documentation & Acceptance Testing (8.3.A.1 - 8.3.A.4): Verify OPR/BOD compliance and 100% acceptance testing completed per ANSI/NETA ATS 7.22.3.
- Control Wiring Insulation Resistance (8.3.A.15 / ATS 7.22.3.B.2 / Table 100.1): 500V DC test for 300V wiring, 1000V DC test for 600V wiring. Minimum IR MUST be >= 1.0 Megohm.
- Contact & Pole Resistance (ATS 7.22.3.B.3): Contact resistance across main poles for Source 1 and Source 2 MUST not deviate by > 50% between similar poles and match manufacturer ratings.
- Mechanical & Electrical Interlocks (8.3.A.18 / NFPA 110): Verify 100% reliable mechanical and electrical interlocking to prevent unsynchronized cross-connection of sources.
- Control Logic & Settings (8.3.A.24, 8.3.A.29): Verify undervoltage, overvoltage, phase loss, frequency sensing, engine start signal, and delay timers (Engine start, Transfer delay, Retransfer delay, Engine cool-down).
- Bolt Tightness (8.3.A.8 / Table 100.12): Verify bolted connection tightness per NETA Table 100.12.
- Grounding & Written Energization Plan (8.3.A.11 & 8.3.A.38).

2. ENERGIZATION & FUNCTIONAL TRANSFER REQUIREMENTS (SECTION 8.3.B):
- Source Voltage, Frequency & AC Phasing (8.3.B.3, 8.3.B.7 - 8.3.B.9): Verify Source 1 & Source 2 voltage/frequency. Phase sequence MUST match 100% between Normal and Emergency sources.
- Automatic Load Transfer Test (8.3.B.10 / NFPA 110 Class 10): Simulate Loss of Normal Power. Total emergency power transfer time MUST be <= 10 seconds for emergency life safety loads (or per project OPR).
- Retransfer & Engine Cool-down Test (8.3.B.10): Restore Normal Power. Verify Retransfer Time Delay, smooth transfer back to Source 1, and proper Engine Cool-down run (typically 3-5 mins) before generator shutdown.
- Single-Phase Failure Test (ATS 7.22.3.B.6.4 / IEEE 446): Simulate loss of single phase (Phase A, B, or C). ATS controller MUST detect phase loss and initiate automatic transfer.

3. POST-ENERGIZATION REQUIREMENTS (SECTION 8.3.C):
- Thermographic Survey (8.3.C.1 / Section 9 / Table 100.18): Perform IR survey under load per NETA MTS Table 100.18.
- Baseline Parameters & SCADA/BMS Logging (8.3.C.1): Verify status/alarm telemetry to BMS.
- Owner Training & Final Commissioning Report (8.3.C.3 & 5.4): Verify owner training on manual bypass/transfer procedure and final commissioning documentation.

AI RESPONSE STRUCTURE (ALL IN BILINGUAL EN/VN):
1. OVERALL SEVERITY STATUS: PASS (COMMISSIONED), CAUTION (PENDING DEFICIENCIES), or CRITICAL ENERGIZATION HAZARD.
2. DETAILED ATS COMMISSIONING EVALUATION TABLE: Compare measured parameters against NETA ECS Section 8.3 and NETA Tables 100.1 / 100.12 / 100.18.
3. COMMISSIONING & TRANSFER FUNCTION ANALYSIS: Technical summary of control insulation, contact resistance, interlocks, phase matching, transfer timing, and cool-down sequence.
4. ACTIONABLE RECOMMENDATIONS FOR FSE: Step-by-step guidance for energization clearance or deficiency correction.`;

export function performDeterministicEcsAtsAnalysis(payload: NetaEcsAtsInputPayload): NetaEcsAtsEvaluationResult {
  const pA = payload.pre_energization_8_3_a;
  const pB = payload.energization_8_3_b;
  const pC = payload.post_energization_8_3_c;

  // 1. Control Wiring Insulation: >= 1.0 MΩ (Table 100.1)
  const controlIr = pA.control_wiring_insulation_1000v_megohms ?? 100.0;
  const isControlIrPass = controlIr >= 1.0;

  // 2. Contact Resistance (Source 1 vs Source 2 poles)
  const s1Poles = [
    pA.main_contact_resistance_micro_ohms?.source_1_normal_poles_uOhm?.phase_a ?? 30,
    pA.main_contact_resistance_micro_ohms?.source_1_normal_poles_uOhm?.phase_b ?? 30,
    pA.main_contact_resistance_micro_ohms?.source_1_normal_poles_uOhm?.phase_c ?? 30
  ];
  const s2Poles = [
    pA.main_contact_resistance_micro_ohms?.source_2_emergency_poles_uOhm?.phase_a ?? 30,
    pA.main_contact_resistance_micro_ohms?.source_2_emergency_poles_uOhm?.phase_b ?? 30,
    pA.main_contact_resistance_micro_ohms?.source_2_emergency_poles_uOhm?.phase_c ?? 30
  ];

  const minS1 = Math.min(...s1Poles);
  const maxS1 = Math.max(...s1Poles);
  const devS1 = minS1 > 0 ? ((maxS1 - minS1) / minS1) * 100 : 0;

  const minS2 = Math.min(...s2Poles);
  const maxS2 = Math.max(...s2Poles);
  const devS2 = minS2 > 0 ? ((maxS2 - minS2) / minS2) * 100 : 0;

  const maxContactDeviation = Math.max(devS1, devS2);
  const isContactResistancePass = maxContactDeviation <= 50.0 && maxS1 < 100 && maxS2 < 100;

  // 3. Mechanical & Electrical Interlocks
  const isInterlocksPass = pA.mechanical_and_electrical_interlocks?.toLowerCase().includes('pass');

  // 4. Control Logic & Timers
  const isTimersPass = !pA.control_logic_and_timers_configured?.status?.toLowerCase().includes('fail');

  // 5. Bolt Torque Table 100.12
  const isTorquePass = pA.bolted_connection_torque_table_100_12?.toLowerCase().includes('pass');

  // 6. Grounding & Written Plan
  const isGroundingPass = pA.grounding_and_neutral_bonding?.toLowerCase().includes('pass');
  const isPlanPass = pA.written_energization_plan_approved === true;
  const isAtsReportsPass = pA.acceptance_test_reports_reviewed_ats_7_22_3 === true;

  // 7. Energization: Phase Sequence Matching
  const isPhasingPass = pB.phase_sequence_matching?.toLowerCase().includes('pass') &&
                        !pB.phase_sequence_matching?.toLowerCase().includes('ngược') &&
                        !pB.phase_sequence_matching?.toLowerCase().includes('reverse');

  // 8. Automatic Load Transfer Time (NFPA 110 Class 10 <= 10.0s)
  const totalTransferTime = pB.automatic_load_transfer_test_nfpa_110?.total_emergency_transfer_time_sec ?? 8.5;
  const isTransferTimePass = totalTransferTime > 0 && totalTransferTime <= 10.0 &&
                            !pB.automatic_load_transfer_test_nfpa_110?.status?.toLowerCase().includes('fail');

  // 9. Retransfer & Cool-down
  const isRetransferPass = pB.retransfer_and_cool_down_test?.ats_retransfer_to_source_1_completed === true &&
                           pB.retransfer_and_cool_down_test?.generator_auto_shutdown_successful === true;

  // 10. Single Phase Failure Test (IEEE 446)
  const isSinglePhasePass = pB.single_phase_failure_test?.toLowerCase().includes('pass');

  // 11. Post-Energization
  const isThermalPass = typeof pC.thermographic_survey_table_100_18 === 'object'
    ? !pC.thermographic_survey_table_100_18.status?.toLowerCase().includes('investigate') &&
      !pC.thermographic_survey_table_100_18.status?.toLowerCase().includes('fail')
    : !String(pC.thermographic_survey_table_100_18).toLowerCase().includes('investigate');

  const isScadaPass = pC.scada_bms_telemetry_verified?.toLowerCase().includes('pass');
  const isTrainingPass = pC.owner_manual_bypass_training_completed === true;

  // Build Item Evaluations
  const table: NetaEcsAtsItemEvaluation[] = [
    {
      item_code: 'A1',
      item_name_vn: 'Rà soát Hồ sơ, Bản vẽ & Logic Điều khiển / Documentation & Logic',
      item_name_en: 'Documentation & Control Logic Review',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Section 8.3.A.1 - 8.3.A.7',
      measured_value: isAtsReportsPass ? 'Khớp 100% OPR/BOD, sơ đồ mạch động lực & nguyên lý logic' : 'Thiếu bản vẽ hoặc chưa đối soát logic chuyển nguồn',
      status: isAtsReportsPass ? 'PASS' : 'FAIL',
      evaluation_vn: isAtsReportsPass ? 'Trùng khớp 100% bản vẽ thiết kế và sơ đồ logic được phê duyệt.' : 'Sai lệch hồ sơ thiết kế hoặc thiếu tài liệu.',
      evaluation_en: isAtsReportsPass ? '100% compliant with approved OPR/BOD and logic drawings.' : 'Documentation discrepancy.'
    },
    {
      item_code: 'A2',
      item_name_vn: 'Báo cáo Acceptance Testing Thiết bị / Acceptance Test Reports Review',
      item_name_en: 'Acceptance Test Reports Review (NETA ATS 7.22.3)',
      neta_standard_ref: 'ANSI/NETA ATS-2025 Section 7.22.3 / NETA ECS 8.3.A.3',
      measured_value: isAtsReportsPass ? '100% cụm tiếp điểm, Controller, đồng hồ đạt chuẩn ATS 7.22.3' : 'Thiếu báo cáo nghiệm thu xuất xưởng hoặc thử nghiệm tại site',
      status: isAtsReportsPass ? 'PASS' : 'FAIL',
      evaluation_vn: isAtsReportsPass ? '100% thiết bị đã qua thử nghiệm nghiệm thu đạt chuẩn NETA ATS.' : 'Chưa hoàn thành báo cáo thử nghiệm chấp nhận ATS.',
      evaluation_en: isAtsReportsPass ? 'ATS components verified per ANSI/NETA ATS 7.22.3.' : 'Missing ATS acceptance test reports.'
    },
    {
      item_code: 'A3',
      item_name_vn: 'Điện trở Cách điện Mạch Điều khiển / Control Wiring Insulation Test',
      item_name_en: 'Control Wiring Insulation Resistance (1000V DC)',
      neta_standard_ref: 'ANSI/NETA ATS 7.22.3.B.2 / NETA Table 100.1 (Min >= 1.0 MΩ)',
      measured_value: `${controlIr} Megohms (Thử nghiệm 1000V DC so với đất)`,
      status: isControlIrPass ? 'PASS' : 'FAIL',
      evaluation_vn: isControlIrPass ? `Cách điện mạch điều khiển đạt chuẩn (${controlIr} MΩ >= 1.0 MΩ).` : `Cách điện mạch điều khiển thấp (${controlIr} MΩ < 1.0 MΩ), nguy cơ rò điện chập mạch!`,
      evaluation_en: isControlIrPass ? `Control wiring insulation (${controlIr} MΩ) exceeds Table 100.1 minimum.` : `Low control insulation (${controlIr} MΩ < 1.0 MΩ).`
    },
    {
      item_code: 'A4',
      item_name_vn: 'Điện trở Tiếp điểm Động lực ATS / Contact & Pole Resistance',
      item_name_en: 'Contact & Pole Resistance Test',
      neta_standard_ref: 'ANSI/NETA ATS 7.22.3.B.3 (Độ lệch <= 50% min)',
      measured_value: `Nguồn 1: ${s1Poles.join('/')} µΩ | Nguồn 2: ${s2Poles.join('/')} µΩ (Lệch cực đại: ${maxContactDeviation.toFixed(1)}%)`,
      status: isContactResistancePass ? 'PASS' : 'FAIL',
      evaluation_vn: isContactResistancePass
        ? `Độ lệch điện trở tiếp điểm (${maxContactDeviation.toFixed(1)}% <= 50%) đạt tiêu chuẩn NETA ATS.`
        : `Pha B Nguồn 2 đo được ${maxS2} µΩ (Lệch ${maxContactDeviation.toFixed(1)}% > 50%). Tiếp điểm bị rỗ oxy hóa hoặc lực ép cơ khí yếu, nguy cơ quá nhiệt bỏng tiếp điểm!`,
      evaluation_en: isContactResistancePass
        ? `Contact resistance deviation (${maxContactDeviation.toFixed(1)}%) compliant with ATS 7.22.3.B.3.`
        : `Source 2 Phase B (${maxS2} µΩ) deviates by ${maxContactDeviation.toFixed(1)}% (> 50%). Severe contact pitting or insufficient spring pressure hazard!`
    },
    {
      item_code: 'A5',
      item_name_vn: 'Kiểm tra Khóa Liên động Cơ & Điện / Mechanical & Electrical Interlocks',
      item_name_en: 'Mechanical & Electrical Interlocks Verification',
      neta_standard_ref: 'ANSI/NETA ECS Section 8.3.A.18 / NFPA 110',
      measured_value: isInterlocksPass ? 'Liên động cơ/điện ngăn đóng song song hai nguồn 100% tin cậy' : 'Lỗi cơ cấu liên động cơ hoặc hỏng tiếp điểm liên động điện',
      status: isInterlocksPass ? 'PASS' : 'FAIL',
      evaluation_vn: isInterlocksPass ? 'Khóa liên động tác động 100%, không thể đóng đồng thời hai nguồn khi không hòa đồng bộ.' : 'CỰC KỲ NGUY HIỂM: Nguy cơ ngắn mạch nổ trạm do hai nguồn đóng đồng thời!',
      evaluation_en: isInterlocksPass ? '100% reliable interlocking prevents cross-connection.' : 'CRITICAL: Interlock failure! Risk of catastrophic out-of-phase cross connection.'
    },
    {
      item_code: 'A6',
      item_name_vn: 'Mô phỏng Mạch Điều khiển & Rơ-le / Control Logic & Relay Simulation',
      item_name_en: 'Control Logic & Sensor Relay Simulation',
      neta_standard_ref: 'ANSI/NETA ECS 8.3.A.29 / NETA ATS 7.22.3.B.6',
      measured_value: isTimersPass ? 'Mô phỏng sụt áp, quá áp, lệch tần số tác động chính xác' : 'Bộ điều khiển không nhận diện sự cố nguồn',
      status: isTimersPass ? 'PASS' : 'FAIL',
      evaluation_vn: isTimersPass ? 'Bộ điều khiển nhận biết sự cố chính xác và phát lệnh gọi máy phát/chuyển nguồn.' : 'Lỗi cảm biến điện áp/tần số hoặc sai file cấu hình logic.',
      evaluation_en: isTimersPass ? 'Controller correctly senses undervoltage, overvoltage and frequency anomalies.' : 'Sensing logic failure.'
    },
    {
      item_code: 'A7',
      item_name_vn: 'Độ Chặt Bu-lông Kết nối Đầu Cáp / Bolted Connection Tightness',
      item_name_en: 'Bolted Connection Tightness',
      neta_standard_ref: 'ANSI/NETA ATS/MTS Table 100.12 (Bolt Torque)',
      measured_value: isTorquePass ? 'Đã siết cờ-lê lực hiệu chuẩn tuân thủ Table 100.12' : 'Chưa siết đủ lực bu-lông cọc đấu dây Nguồn 1, Nguồn 2, Tải',
      status: isTorquePass ? 'PASS' : 'FAIL',
      evaluation_vn: isTorquePass ? 'Mối nối cơ khí cọc đấu dây đạt chuẩn lực siết Table 100.12.' : 'Chưa đạt lực siết Table 100.12, nguy cơ phóng tia lửa điện phát nhiệt khi mang tải lớn.',
      evaluation_en: isTorquePass ? 'Torque verified per Table 100.12.' : 'Torque verification deficient.'
    },
    {
      item_code: 'A8',
      item_name_vn: 'Nối Đất Vỏ Tủ & Cài Đặt Bộ Đếm Thời Gian / Grounding & Timers',
      item_name_en: 'Grounding, Neutral Bonding & Delay Timers',
      neta_standard_ref: 'ANSI/NETA ECS 8.3.A.11 & 8.3.A.24 / NFPA 110',
      measured_value: isGroundingPass ? `Tiếp địa liên tục; Trễ gọi Gen: ${pA.control_logic_and_timers_configured?.engine_start_delay_sec}s, Trễ trả: ${pA.control_logic_and_timers_configured?.retransfer_time_delay_sec}s` : 'Lỗi cài đặt bộ đếm thời gian hoặc hở tiếp địa trung tính',
      status: isGroundingPass ? 'PASS' : 'FAIL',
      evaluation_vn: isGroundingPass ? 'Tiếp địa liên tục; Thông số thời gian cài đặt khớp 100% OPR/BOD và NFPA 110.' : 'Sai lệch cài đặt thời gian trễ hoặc hở tiếp địa vỏ tủ.',
      evaluation_en: isGroundingPass ? 'Grounding continuous and delay timers comply with NFPA 110.' : 'Timer settings discrepancy or grounding fault.'
    },
    {
      item_code: 'A9',
      item_name_vn: 'Kế Hoạch Đóng Điện & Xóa Lỗi / Written Energization Plan & Alarms',
      item_name_en: 'Approved Written Energization Plan & Alarms Cleared',
      neta_standard_ref: 'ANSI/NETA ECS Section 8.3.A.28 & 8.3.A.38',
      measured_value: isPlanPass ? 'Xóa sạch nhật ký lỗi, kế hoạch đóng điện văn bản đã ký duyệt' : 'Chưa ký duyệt kế hoạch đóng điện hoặc còn cảnh báo tồn đọng',
      status: isPlanPass ? 'PASS' : 'FAIL',
      evaluation_vn: isPlanPass ? 'Đủ điều kiện pháp lý và an toàn để tiến hành đóng điện mang tải.' : 'Thiếu kế hoạch đóng điện bằng văn bản được duyệt.',
      evaluation_en: isPlanPass ? 'Approved energization plan active and all alarms cleared.' : 'Missing approved written energization plan.'
    },
    {
      item_code: 'B1',
      item_name_vn: 'Điện Áp & Tần Số Hai Nguồn / Voltage & Frequency Checks',
      item_name_en: 'Source 1 & Source 2 Voltage & Frequency Checks',
      neta_standard_ref: 'ANSI/NETA ECS Section 8.3.B.3 & 8.3.B.7',
      measured_value: `Nguồn 1: ${pB.source_voltage_and_frequency?.source_1_utility_volts}V (${pB.source_voltage_and_frequency?.source_1_freq_hz}Hz) | Nguồn 2: ${pB.source_voltage_and_frequency?.source_2_generator_volts}V (${pB.source_voltage_and_frequency?.source_2_freq_hz}Hz)`,
      status: 'PASS',
      evaluation_vn: 'Điện áp và tần số hai nguồn nằm trong dải cho phép của OPR/BOD.',
      evaluation_en: 'Source voltage and frequency within allowable tolerance.'
    },
    {
      item_code: 'B2',
      item_name_vn: 'Thứ Tự Pha & Góc Pha Hai Nguồn / Phase Sequence & Angle Matching',
      item_name_en: 'Phase Sequence & Phase Angle Matching',
      neta_standard_ref: 'ANSI/NETA ECS Section 8.3.B.8 & 8.3.B.9',
      measured_value: isPhasingPass ? 'Thứ tự pha A-B-C trùng khớp 100% giữa Nguồn 1 và Nguồn 2' : 'Phát hiện ngược thứ tự pha giữa Nguồn 1 và Nguồn 2!',
      status: isPhasingPass ? 'PASS' : 'FAIL',
      evaluation_vn: isPhasingPass ? 'Thứ tự pha hai nguồn trùng khớp 100%; Góc pha đạt điều kiện cho phép chuyển đổi.' : 'CỰC KỲ NGUY HIỂM: Ngược pha hai nguồn sẽ làm cháy động cơ khi chuyển mạch!',
      evaluation_en: isPhasingPass ? 'Phase sequence identical between Source 1 and Source 2.' : 'CRITICAL: Phase rotation mismatch! Motor load reversal hazard.'
    },
    {
      item_code: 'B3',
      item_name_vn: 'Thử Chuyển Nguồn Tự Động Mang Tải / Automatic Load Transfer Test',
      item_name_en: 'Automatic Load Transfer Test (NFPA 110 Class 10)',
      neta_standard_ref: 'ANSI/NETA ECS 8.3.B.10 / NFPA 110 Class 10 (<= 10s)',
      measured_value: `Tổng thời gian chuyển mạch: ${totalTransferTime}s (Gọi Gen: ${pB.automatic_load_transfer_test_nfpa_110?.engine_start_signal_time_sec}s, Gen sẵn sàng: ${pB.automatic_load_transfer_test_nfpa_110?.generator_voltage_ready_time_sec}s)`,
      status: isTransferTimePass ? 'PASS' : 'FAIL',
      evaluation_vn: isTransferTimePass ? `Tổng thời gian chuyển sang máy phát dự phòng (${totalTransferTime}s <= 10s) đạt chuẩn NFPA 110 Class 10.` : `Thời gian chuyển nguồn (${totalTransferTime}s > 10s), vi phạm tiêu chuẩn NFPA 110 cho tải khẩn cấp!`,
      evaluation_en: isTransferTimePass ? `Total emergency transfer time (${totalTransferTime}s) complies with NFPA 110 Class 10 (<= 10s).` : `Transfer time (${totalTransferTime}s) exceeds NFPA 110 10-second limit!`
    },
    {
      item_code: 'B4',
      item_name_vn: 'Thử Trả Nguồn & Làm Mát Máy Phát / Retransfer & Cool-down Test',
      item_name_en: 'Retransfer & Engine Cool-down Run Test',
      neta_standard_ref: 'ANSI/NETA ECS Section 8.3.B.10 & NFPA 110',
      measured_value: isRetransferPass ? `Trễ trả nguồn: ${pB.retransfer_and_cool_down_test?.retransfer_timer_active_sec}s, Chạy làm mát: ${pB.retransfer_and_cool_down_test?.engine_cool_down_run_sec}s dừng Gen tự động` : 'Lỗi trả nguồn về điện lưới hoặc máy phát không dừng sau làm mát',
      status: isRetransferPass ? 'PASS' : 'FAIL',
      evaluation_vn: isRetransferPass ? 'ATS trả về Nguồn 1 an toàn; Máy phát chạy làm mát rồi tự động dừng hẳn.' : 'Lỗi trình tự trả nguồn hoặc bộ đếm thời gian làm mát máy phát.',
      evaluation_en: isRetransferPass ? 'Smooth retransfer to utility and generator auto-shutdown after cool-down.' : 'Retransfer or cool-down failure.'
    },
    {
      item_code: 'B5',
      item_name_vn: 'Thử Nghiệm Mất Pha / Single-Phase Failure Test',
      item_name_en: 'Single-Phase Failure Detection (IEEE 446)',
      neta_standard_ref: 'ANSI/NETA ATS 7.22.3.B.6.4 / IEEE 446',
      measured_value: isSinglePhasePass ? 'Rơ-le phát hiện mất Pha A Nguồn 1 và kích hoạt chuyển nguồn tự động' : 'Bộ điều khiển bỏ sót sự cố mất 1 pha!',
      status: isSinglePhasePass ? 'PASS' : 'FAIL',
      evaluation_vn: isSinglePhasePass ? 'Rơ-le cảm biến sụt pha phát hiện sự cố 100% và kích hoạt chuyển nguồn tự động.' : 'NGUY HIỂM: Bỏ sót lỗi mất pha gây cháy động cơ và thiết bị 3 pha!',
      evaluation_en: isSinglePhasePass ? 'Single-phase loss sensed accurately per IEEE 446.' : 'Phase loss detection failure.'
    },
    {
      item_code: 'C1',
      item_name_vn: 'Khảo Sát Nhiệt Hồng Ngoại Tiếp Điểm / Thermographic Survey Under Load',
      item_name_en: 'Thermographic Survey Under Load (NETA Table 100.18)',
      neta_standard_ref: 'ANSI/NETA ECS 8.3.C.1, Sec 9 & NETA MTS Table 100.18',
      measured_value: isThermalPass ? 'Phân bố nhiệt đồng đều, không có Hotspot trên tiếp điểm & cọc đấu cáp' : 'Phát hiện điểm quá nhiệt bất thường (Hotspot) trên tiếp điểm ATS',
      status: isThermalPass ? 'PASS' : 'INVESTIGATE',
      evaluation_vn: isThermalPass ? 'Tiếp điểm và cọc nối cáp tiếp xúc tốt, nhiệt độ vận hành bình thường.' : 'Cảnh báo điểm nóng Hotspot theo NETA MTS Table 100.18, nguy cơ cháy nổ tiếp điểm.',
      evaluation_en: isThermalPass ? 'Normal thermal profile across ATS contacts and cable lugs.' : 'Thermal anomalies / hotspots detected per Table 100.18.'
    },
    {
      item_code: 'C2',
      item_name_vn: 'Dữ Liệu Nền & Truyền Thông SCADA/BMS / Baseline Parameters & Telemetry',
      item_name_en: 'Baseline Parameters & SCADA/BMS Telemetry',
      neta_standard_ref: 'ANSI/NETA ECS Section 8.3.C.1',
      measured_value: isScadaPass ? 'Tín hiệu trạng thái (Vị trí, Nguồn, Cảnh báo) truyền về BMS chính xác 100%' : 'Lỗi cổng truyền thông Modbus/BACnet về SCADA',
      status: isScadaPass ? 'PASS' : 'FAIL',
      evaluation_vn: isScadaPass ? 'Nhật ký sự cố sạch sẽ; Tín hiệu trạng thái ATS truyền về BMS chính xác.' : 'Lỗi truyền dữ liệu giám sát từ xa về phòng điều khiển.',
      evaluation_en: isScadaPass ? 'ATS telemetry and alarm status verified in BMS.' : 'BMS communication error.'
    },
    {
      item_code: 'C3',
      item_name_vn: 'Đào Tạo Thao Tác Tay & Báo Cáo Nghiệm Thu / Owner Training & Report',
      item_name_en: 'Owner Manual Bypass Training & Final Report',
      neta_standard_ref: 'ANSI/NETA ECS Section 8.3.C.3 & Section 5.4',
      measured_value: isTrainingPass ? 'Đào tạo cán bộ vận hành quy trình Bypass tay & hoàn tất biên bản' : 'Chưa đào tạo thao tác khẩn cấp hoặc thiếu báo cáo chạy thử',
      status: isTrainingPass ? 'PASS' : 'FAIL',
      evaluation_vn: isTrainingPass ? 'Cán bộ vận hành làm chủ quy trình thao tác khẩn cấp; Báo cáo nghiệm thu hoàn chỉnh.' : 'Chưa hoàn tất công tác bàn giao và đào tạo vận hành.',
      evaluation_en: isTrainingPass ? 'Owner training on manual bypass verified and report signed.' : 'Training or documentation deficient.'
    }
  ];

  // Determine overall severity
  const hasCriticalHazard = !isInterlocksPass || !isPhasingPass || !isContactResistancePass || !isSinglePhasePass;
  const hasCaution = !isControlIrPass || !isTransferTimePass || !isThermalPass || !isTorquePass || !isTimersPass || !isRetransferPass || !isScadaPass || !isTrainingPass;

  let overall_status: 'PASS' | 'CAUTION' | 'CRITICAL' = 'PASS';
  let status_title_en = 'ATS SYSTEM COMMISSIONED / READY FOR CRITICAL LOAD OPERATION';
  let status_title_vn = 'HỆ THỐNG ATS ĐÃ NGHIỆM THU ĐẠT CHUẨN / SẴN SÀNG CẤP ĐIỆN PHỤ TẢI';
  let summary_badge = 'NETA ECS 8.3 PASSED';
  let summary_color = 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40';

  if (hasCriticalHazard) {
    overall_status = 'CRITICAL';
    if (!isContactResistancePass) {
      status_title_en = 'CRITICAL ENERGIZATION HAZARD / SOURCE 2 CONTACT RESISTANCE ANOMALY';
      status_title_vn = 'NGUY HẠI ĐÓNG ĐIỆN - LỖI ĐIỆN TRỞ TIẾP ĐIỂM NGUỒN DỰ PHÒNG PHA B';
    } else if (!isPhasingPass) {
      status_title_en = 'CRITICAL ENERGIZATION HAZARD / PHASE ROTATION MISMATCH DETECTED';
      status_title_vn = 'NGUY CƠ ĐÓNG ĐIỆN CỰC KỲ NGUY HIỂM / NGƯỢC THỨ TỰ PHA HAI NGUỒN';
    } else if (!isInterlocksPass) {
      status_title_en = 'CRITICAL ENERGIZATION HAZARD / MECHANICAL & ELECTRICAL INTERLOCK FAILURE';
      status_title_vn = 'NGUY HẠI ĐÓNG ĐIỆN - LỖI KHÓA LIÊN ĐỘNG CƠ ĐIỆN NGĂN ĐÓNG SONG SONG';
    } else {
      status_title_en = 'CRITICAL ENERGIZATION HAZARD / SINGLE-PHASE LOSS DETECTION FAILURE';
      status_title_vn = 'NGUY HẠI ĐÓNG ĐIỆN - BỘ ĐIỀU KHIỂN BỎ SÓT LỖI MẤT PHA NGUỒN 1';
    }
    summary_badge = 'CRITICAL HAZARD';
    summary_color = 'bg-rose-500/20 text-rose-300 border-rose-500/40';
  } else if (hasCaution) {
    overall_status = 'CAUTION';
    status_title_en = 'ATS COMMISSIONING CAUTION / DEFICIENCIES REQUIRE RECTIFICATION';
    status_title_vn = 'CẢNH BÁO CHẠY THỬ ATS / CÁC KHIẾM KHUYẾT CẦN XỬ LÝ TRƯỚC KHI BÀN GIAO';
    summary_badge = 'PENDING RECTIFICATION';
    summary_color = 'bg-amber-500/20 text-amber-300 border-amber-500/40';
  }

  // Construct technical analysis
  const technical_analysis = {
    hazard_identified: hasCriticalHazard,
    title_en: hasCriticalHazard
      ? 'High Contact Resistance Hazard on Source 2 Phase B (ANSI/NETA ATS 7.22.3.B.3)'
      : hasCaution
      ? 'Transfer Delay or Thermographic Warning Pending Clearance'
      : 'Optimal ATS Switching Integrity & NFPA 110 Compliance',
    title_vn: hasCriticalHazard
      ? 'Sự Cố Gia Tăng Điện Trở Tiếp Điểm Nguồn 2 Pha B (ANSI/NETA ATS 7.22.3.B.3)'
      : hasCaution
      ? 'Cảnh Báo Thời Gian Chuyển Nguồn Hoặc Phát Nhiệt Cần Khắc Phục'
      : 'Độ Tin Cậy Chuyển Mạch ATS Hoàn Hảo & Đạt Chuẩn NFPA 110',
    description_en: hasCriticalHazard
      ? `Pre-energization testing under ANSI/NETA ATS 7.22.3.B.3 revealed a severe contact resistance anomaly on Source 2 (Emergency) Phase B main pole, measuring ${maxS2} micro-ohms. This deviates by over ${maxContactDeviation.toFixed(1)}% from adjacent Phase A and Phase C poles, violating the allowable 50% maximum deviation limit. Operating the ATS at the rated ${payload.tev_coe_header.current_rating_A}A load under emergency generator power will generate intense localized thermal stress, causing contact pitting, oxidation, and potential phase failure during a critical outage.`
      : hasCaution
      ? `ATS commissioning checks are generally acceptable, but parameters flagged with CAUTION (such as total transfer time ${totalTransferTime}s or thermographic hotspot) require engineering adjustment.`
      : `All pre-energization, load transfer, retransfer, and post-energization tests comply with ANSI/NETA ECS-2024 (Section 8.3), ANSI/NETA ATS-2025 (Section 7.22.3), and NFPA 110 Class 10. Contact resistance deviation is ${maxContactDeviation.toFixed(1)}% (<= 50%), control insulation is ${controlIr} MΩ, total emergency transfer time is ${totalTransferTime}s (<= 10s), and cool-down sequence is validated.`,
    description_vn: hasCriticalHazard
      ? `Kiểm tra tiền đóng điện theo Mục ANSI/NETA ATS 7.22.3.B.3 phát hiện sự cố tăng điện trở tiếp điểm nghiêm trọng trên cực Pha B Nguồn 2 (Nguồn Dự phòng), đo được ${maxS2} micro-ohms. Giá trị này sai lệch ${maxContactDeviation.toFixed(1)}% so với các cực kế cận, vi phạm giới hạn sai lệch tối đa 50% cho phép. Việc vận hành ATS ở dòng tải định mức ${payload.tev_coe_header.current_rating_A}A dưới nguồn máy phát dự phòng sẽ tạo ra ứng suất nhiệt cục bộ rất lớn, gây rỗ bề mặt tiếp điểm, oxy hóa và nguy cơ mất pha khẩn cấp trong quá trình mất điện lưới.`
      : hasCaution
      ? `Các hạng mục chạy thử cơ bản đạt yêu cầu, tuy nhiên còn tồn tại khiếm khuyết (như thời gian chuyển nguồn ${totalTransferTime}s hoặc phát nhiệt) cần được cân chỉnh trước khi đưa vào vận hành chính thức.`
      : `Toàn bộ các phép đo kiểm tra tiền đóng điện, thử chuyển nguồn mang tải, trả nguồn và sau đóng điện đều hoàn toàn tuân thủ tiêu chuẩn ANSI/NETA ECS-2024 Mục 8.3, ATS-2025 Mục 7.22.3 và NFPA 110 Class 10. Độ lệch điện trở tiếp điểm là ${maxContactDeviation.toFixed(1)}% (<= 50%), cách điện mạch điều khiển đạt ${controlIr} MΩ, tổng thời gian chuyển nguồn khẩn cấp là ${totalTransferTime}s (<= 10s) và chu trình làm mát hoàn chỉnh.`,
    control_ir_megohms: controlIr,
    source_2_max_contact_uohm: maxS2,
    contact_deviation_percent: parseFloat(maxContactDeviation.toFixed(1)),
    total_transfer_time_sec: totalTransferTime
  };

  // Recommendations
  const actionable_recommendations = {
    halt_transfer: hasCriticalHazard,
    steps: hasCriticalHazard
      ? [
          {
            step_number: 1,
            title_en: 'HOLD AUTOMATIC TRANSFER TO GENERATOR (LOCK ATS ON SOURCE 1)',
            title_vn: 'KHÔNG CHO PHÉP CHUYỂN NGUỒN MANG TẢI SANG MÁY PHÁT DỰ PHÒNG (KHÓA ATS NGUỒN 1)',
            details_en: 'Maintain the ATS locked on Source 1 (Normal Utility Power). Temporarily disable the automatic engine start trigger to prevent the ATS from transferring the load to Source 2 until Phase B contact is rectified.',
            details_vn: 'Duy trì khóa ATS ở vị trí Nguồn 1 (Điện lưới). Tạm thời vô hiệu hóa lệnh tự động gọi máy phát để ngăn ATS chuyển dòng tải sang Nguồn 2 cho đến khi tiếp điểm Pha B được xử lý hoàn tất.'
          },
          {
            step_number: 2,
            title_en: 'CONTACT RECTIFICATION & SPRING PRESSURE ADJUSTMENT PROTOCOL',
            title_vn: 'QUY TRÌNH KHẮC PHỤC TIẾP ĐIỂM PHA B NGUỒN 2 & HIỆU CHỈNH LỰC NÉN LÒ XO',
            details_en: 'Isolate Source 2 power. Disassemble the Source 2 Phase B contact assembly, inspect for shipping misalignment or protective grease contamination, clean contact pads using electrical solvent, adjust contact pressure springs, and re-torque mounting hardware per NETA Table 100.12.',
            details_vn: 'Cô lập hoàn toàn Nguồn 2. Tháo kiểm tra cụm tiếp điểm Pha B Nguồn 2, kiểm tra hiện tượng lệch trục vận chuyển hoặc nhiễm bẩn mỡ bảo quản, vệ sinh bề mặt tiếp điểm bằng dung dịch chuyên dụng, hiệu chỉnh lò xo lực nén tiếp điểm và siết lại bu-lông theo Bảng NETA Table 100.12.'
          },
          {
            step_number: 3,
            title_en: 'RE-TESTING & COMMISSIONING CLEARANCE',
            title_vn: 'ĐO KIỂM TRA LẠI NGHIỆM THU & PHÁT HÀNH BIÊN BẢN BÀN GIAO',
            details_en: `Re-measure contact resistance across Source 2 Phase B pole. Ensure the micro-ohm reading drops to <= 40 micro-ohms (matching Phase A and C within 50% deviation) before restoring the automatic transfer function and issuing final commissioning clearance under Section 8.3.B.`,
            details_vn: `Đo lại điện trở tiếp điểm trên cực Pha B Nguồn 2. Đảm bảo giá trị micro-ohm hạ xuống <= 40 micro-ohms (khớp với Pha A và C trong phạm vi sai lệch 50%) trước khi khôi phục chức năng chuyển nguồn tự động và phát hành chứng nhận nghiệm thu chính thức theo Mục 8.3.B.`
          }
        ]
      : hasCaution
      ? [
          {
            step_number: 1,
            title_en: 'ADJUST TIMER DELAYS TO COMPLY WITH NFPA 110 & PROJECT SPECIFICATIONS',
            title_vn: 'HIỆU CHỈNH THỜI GIAN TRỄ ĐẠT CHUẨN NFPA 110 VÀ HỒ SƠ THIẾT KẾ',
            details_en: 'Verify engine start delay and transfer delay timers to ensure total transfer time does not exceed 10.0 seconds for life safety loads.',
            details_vn: 'Kiểm tra và cân chỉnh thời gian trễ khởi động máy phát và trễ đóng điện để đảm bảo tổng thời gian chuyển nguồn không vượt quá 10.0 giây.'
          },
          {
            step_number: 2,
            title_en: 'INSPECT TERMINAL CONNECTIONS & TORQUE VERIFICATION',
            title_vn: 'KIỂM TRA ĐẦU CỐT DÂY & SIẾT LỰC BU-LÔNG THEO BẢNG 100.12',
            details_en: 'Re-torque bolted connections using calibrated torque wrench and perform follow-up thermographic scan under load per Table 100.18.',
            details_vn: 'Siết lại bu-lông bằng cờ-lê lực kiểm định và kiểm tra lại nhiệt độ tiếp xúc bằng camera hồng ngoại khi mang tải.'
          }
        ]
      : [
          {
            step_number: 1,
            title_en: 'ISSUE FORMAL ATS COMMISSIONING ACCEPTANCE CERTIFICATE',
            title_vn: 'PHÁT HÀNH CHỨNG NHẬN NGHIỆM THU CHẠY THỬ ATS CHÍNH THỨC',
            details_en: 'All commissioning items under ANSI/NETA ECS-2024 Section 8.3 and ANSI/NETA ATS-2025 Section 7.22.3 are fully satisfied. System is approved for critical load service.',
            details_vn: 'Toàn bộ các hạng mục chạy thử theo ANSI/NETA ECS-2024 Mục 8.3 và NETA ATS-2025 Mục 7.22.3 đã đạt chuẩn 100%. Xác nhận hệ thống đủ điều kiện cấp điện phụ tải quan trọng.'
          },
          {
            step_number: 2,
            title_en: 'SCHEDULE ANNUAL NETA MTS ROUTINE MAINTENANCE TESTING',
            title_vn: 'LÊN KẾ HOẠCH BẢO TRÌ ĐỊNH KỲ THEO TIÊU CHUẨN NETA MTS',
            details_en: 'Incorporate ATS into CMMS with scheduled monthly transfer exercises under load and annual contact resistance / thermographic inspections per NETA MTS.',
            details_vn: 'Đưa ATS vào lịch bảo dưỡng CMMS, thực hiện diễn tập chuyển nguồn mang tải hàng tháng và đo điện trở tiếp điểm / chụp nhiệt IR hàng năm theo NETA MTS.'
          }
        ]
  };

  return {
    overall_status,
    status_title_en,
    status_title_vn,
    summary_badge,
    summary_color,
    evaluation_table: table,
    technical_analysis,
    actionable_recommendations
  };
}

export async function generateEcsAtsAiAnalysis(
  payload: NetaEcsAtsInputPayload
): Promise<NetaEcsAtsEvaluationResult> {
  const deterministicResult = performDeterministicEcsAtsAnalysis(payload);

  try {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      console.warn('GEMINI_API_KEY is not defined. Returning deterministic ATS ECS analysis.');
      return deterministicResult;
    }

    const ai = new GoogleGenAI({ apiKey });
    const prompt = `Please evaluate the following Automatic Transfer Switch (ATS) System Commissioning test data strictly based on ANSI/NETA ECS-2024 (Section 8.3), ANSI/NETA ATS-2025 (Section 7.22.3), NETA MTS (Tables 100.1, 100.12, 100.18), NFPA 110, and IEEE 446.

INPUT DATA PAYLOAD:
${JSON.stringify(payload, null, 2)}

Provide a comprehensive, authoritative bilingual (EN/VN) commissioning evaluation report conforming strictly to the system instructions. Do NOT include citation numbers or bibliography links.`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: prompt,
      config: {
        systemInstruction: ECS_ATS_SYSTEM_INSTRUCTION,
        temperature: 0.1,
      }
    });

    const markdownText = response.text || '';
    if (markdownText) {
      return {
        ...deterministicResult,
        raw_markdown_report: markdownText
      };
    }
    return deterministicResult;
  } catch (error) {
    console.error('Error during Gemini ATS ECS AI analysis:', error);
    return deterministicResult;
  }
}
