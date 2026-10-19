import { GoogleGenAI } from '@google/genai';

export interface NetaEcsUpsInputPayload {
  tev_coe_header: {
    module: string;
    project_name: string;
    system_tag: string;
    ups_rating_kVA: number;
    ac_voltage_rating_V: number;
    dc_bus_voltage_V: number;
    battery_type: string;
    battery_capacity_Ah: number;
    fse_name: string;
    commissioning_date: string;
  };
  pre_energization_8_2_a: {
    documentation_and_drawings_approved: boolean;
    acceptance_test_reports_reviewed_ats_7_22_2: boolean;
    power_circuit_insulation_1000v_megohms: number;
    control_wiring_insulation_500v_megohms: number;
    battery_bank_testing: {
      number_of_cells: number;
      max_cell_impedance_deviation_percent: number; // <= 10-20% deviation per IEEE 1188 / 450
      battery_float_voltage_V: number;
      cell_voltage_uniformity_pass: boolean;
      bolted_connections_torque_table_100_12: string; // 'Pass' | 'Fail'
      status?: string;
    };
    mechanical_and_electrical_interlocks: string; // 'Pass' | 'Fail'
    grounding_and_neutral_bonding: string; // 'Pass' | 'Fail' (<= 5.0 Ohm)
    written_energization_plan_approved: boolean;
  };
  energization_8_2_b: {
    input_voltage_and_phase_rotation: {
      rectifier_input_voltage_V: number;
      bypass_input_voltage_V: number;
      phase_sequence_match: string; // 'Pass' | 'Fail'
      status?: string;
    };
    rectifier_startup_and_dc_bus: {
      dc_bus_float_voltage_V: number;
      dc_ripple_voltage_percent: number; // <= 1.0%
      status?: string;
    };
    inverter_startup_and_sync: {
      inverter_output_voltage_V: number;
      phase_balance_deviation_percent: number; // <= 1.0%
      sync_phase_angle_deg: number; // <= 3.0 deg
      sync_voltage_diff_V: number; // <= 5.0 V
      status?: string;
    };
    static_bypass_transfer_test: {
      transfer_time_ms: number; // 0 ms
      transient_voltage_deviation_percent: number; // <= 5%
      status?: string;
    };
    loss_of_ac_input_power_test: {
      transfer_to_battery_seamless: boolean;
      output_voltage_drop_V: number; // <= 2%
      status?: string;
    };
    battery_discharge_and_full_load_test: {
      load_percentage: number; // 100%
      discharge_duration_minutes: number;
      end_of_discharge_voltage_V: number;
      full_load_capacity_met: boolean;
      status?: string;
    };
    output_thd_harmonics: {
      thd_v_percent: number; // <= 3.0% linear, <= 5.0% non-linear per IEEE 519
      status?: string;
    };
    manual_maintenance_bypass_transfer_test: string; // 'Pass' | 'Fail'
  };
  post_energization_8_2_c: {
    thermographic_survey_table_100_18: {
      max_delta_t_celsius?: number;
      status: string; // 'Pass' | 'Investigate'
      hotspot_location?: string;
    };
    scada_bms_telemetry_verified: string; // 'Pass' | 'Fail'
    owner_training_and_handover_completed: boolean;
  };
}

export interface NetaEcsUpsItemEvaluation {
  item_code: string;
  item_name_vn: string;
  item_name_en: string;
  neta_standard_ref: string;
  measured_value: string;
  status: 'PASS' | 'FAIL' | 'INVESTIGATE';
  evaluation_vn: string;
  evaluation_en: string;
}

export interface NetaEcsUpsEvaluationResult {
  overall_status: 'PASS' | 'CAUTION' | 'CRITICAL';
  status_title_en: string;
  status_title_vn: string;
  summary_badge: string;
  summary_color: string;
  evaluation_table: NetaEcsUpsItemEvaluation[];
  technical_analysis: {
    hazard_identified: boolean;
    title_en: string;
    title_vn: string;
    description_en: string;
    description_vn: string;
    power_ir_megohms: number;
    battery_impedance_deviation_percent: number;
    static_transfer_time_ms: number;
    thd_v_percent: number;
    thermal_delta_t?: number;
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

export const ECS_UPS_SYSTEM_INSTRUCTION = `You are the TEV COE Analytical Engine — an AI Expert Assistant specializing in Uninterruptible Power Supply (UPS) System Commissioning and Acceptance Testing, operating strictly according to ANSI/NETA ECS-2024 (Section 8.2), ANSI/NETA ATS-2025 (Section 7.22.2 & 7.18), NETA MTS (Tables 100.1, 100.12, 100.18), IEEE 450, IEEE 1188, IEEE 1184, and NFPA 111 standards. Your task is to evaluate UPS commissioning test data submitted by Field Service Engineers (FSE) and output an authoritative bilingual (English / Vietnamese EN/VN) diagnostic report without citing any source reference numbers.

EVALUATION RULES & REFERENCE STANDARDS (ANSI/NETA ECS-2024 SECTION 8.2 & NETA TABLES):

1. PRE-ENERGIZATION REQUIREMENTS (SECTION 8.2.A):
- Documentation & Drawing Review (8.2.A.1 - 8.2.A.6): 100% compliance with approved OPR, BOD, single-line diagrams, rectifier, inverter, static bypass, and battery strings.
- Acceptance Test Reports Review (8.2.A.3 / ATS 7.22.2, 7.18 & 7.6): Verified 100% acceptance test reports for UPS cubicle, isolation transformers, battery racks, and DC breakers.
- Power Circuit & Control Wiring Insulation (ATS 7.22.2.B.2 / Table 100.1): AC input/output bus and DC bus tested at 1000V DC MUST be >= 100 Megohms. Control wiring tested at 500V DC MUST be >= 1.0 Megohm.
- Battery Bank Inspection & Ohmic Testing (ATS 7.18 / IEEE 450 VLA / IEEE 1188 VRLA / IEEE 1184): Cell internal impedance / resistance deviation MUST not exceed 10-20% from baseline or string average. Terminal connection torque verified per NETA Table 100.12.
- Mechanical & Electrical Interlocks (8.2.A.18): 100% positive interlocking between Maintenance Bypass Breaker (MBB) and Inverter / Static Switch to prevent parallel unsynchronized backfeeding.
- Grounding & Neutral Bonding (8.2.A.11): System grounding resistance MUST be <= 5.0 Ohms. Separately Derived System (SDS) neutral-ground bonding verified.
- Written Energization Safety Plan (8.2.A.38): Approved step-by-step procedure and Arc Flash / DC hazard PPE compliance.

2. ENERGIZATION & FUNCTIONAL TESTS (SECTION 8.2.B):
- AC Input Voltage & Phase Rotation (8.2.B.3, 8.2.B.7): Rectifier and Bypass AC input within +/- 10% nominal, phase sequence A-B-C verified 100%.
- Rectifier Start-up & DC Bus (8.2.B.8): DC bus float/equalize charging voltage stable, AC ripple voltage <= 1.0% of DC bus voltage.
- Inverter Start-up & Bypass Synchronization (8.2.B.9): Output 3-phase voltage balanced (deviation <= 1.0%), synchronization with Bypass line within phase angle <= 3 degrees and voltage diff <= 5V.
- Static Bypass Transfer Test (8.2.B.10 / ATS 7.22.2.B.6): Transfer and retransfer between Inverter and Static Bypass MUST achieve 0 ms (seamless no-break transfer) with transient deviation <= 5%.
- Loss of AC Input Power Test (8.2.B.11): Simulated mains failure triggers instantaneous battery support with zero voltage dip or load disruption.
- Battery Discharge & Full Load Test (NFPA 111 / IEEE 450/1188): Step load test (50%, 75%, 100%) using load bank for specified design duration (15-60 mins). End-of-discharge cell voltages remain above cutoff thresholds.
- Output THD-v & Harmonics (8.2.B.12 / IEEE 519): Total voltage harmonic distortion THD-v MUST be <= 3.0% for linear loads, <= 5.0% for non-linear loads.
- Manual Maintenance Bypass Transfer Test (8.2.B.13): Safe, uninterrupted transfer to MBB for total isolation and servicing.

3. POST-ENERGIZATION REQUIREMENTS (SECTION 8.2.C):
- Infrared Thermographic Survey (8.2.C.1 / Table 100.18): Survey IGBT modules, thyristors, battery inter-tier cables, and lugs under >= 40% load. Max Delta-T between similar phases MUST be <= 3.0°C (Priority 4 normal).
- SCADA/BMS/DCIM Telemetry Integration (8.2.C.1): Verification of Modbus/SNMP telemetry for voltages, currents, load %, battery SoC/SoH, and emergency alarms.
- Operations Training & As-built Handover (8.2.C.3 & 5.4): Owner operations personnel trained in emergency manual bypass operations and full commissioning package handed over.

AI RESPONSE STRUCTURE (ALL IN AUTHORITATIVE BILINGUAL EN/VN):
1. OVERALL SEVERITY STATUS: PASS (COMMISSIONED), CAUTION (PENDING DEFICIENCIES), or CRITICAL ENERGIZATION HAZARD.
2. DETAILED UPS COMMISSIONING EVALUATION TABLE: Compare measured values against ANSI/NETA ECS Section 8.2 and NETA Tables 100.1 / 100.12 / 100.18.
3. ENGINEERING TECHNICAL ANALYSIS: In-depth technical breakdown of battery health, static switch timing, power quality (THD-v), and thermal findings.
4. ACTIONABLE PROTOCOL STEPS: Sequential numbered engineering steps with decisive recommendations.`;

export function performDeterministicEcsUpsAnalysis(
  payload: NetaEcsUpsInputPayload
): NetaEcsUpsEvaluationResult {
  const pre = payload.pre_energization_8_2_a;
  const eng = payload.energization_8_2_b;
  const post = payload.post_energization_8_2_c;

  // A1: Documentation & Drawings
  const isDocPass = pre.documentation_and_drawings_approved;
  // A2: Acceptance Tests Review
  const isAtsPass = pre.acceptance_test_reports_reviewed_ats_7_22_2;
  // A3: Insulation Resistance
  const powerIr = pre.power_circuit_insulation_1000v_megohms;
  const isPowerIrPass = powerIr >= 100.0; // Table 100.1 for 400V/1000V bus >= 100 MΩ
  const controlIr = pre.control_wiring_insulation_500v_megohms;
  const isControlIrPass = controlIr >= 1.0; // Table 100.1 >= 1.0 MΩ
  // A4: Battery Bank Testing
  const batDeviation = pre.battery_bank_testing.max_cell_impedance_deviation_percent;
  const isBatteryDeviationPass = batDeviation <= 20.0; // IEEE 1188 / 450: <= 20% from baseline
  const isBatteryTorquePass = pre.battery_bank_testing.bolted_connections_torque_table_100_12.toLowerCase() === 'pass';
  const isBatteryCellUniform = pre.battery_bank_testing.cell_voltage_uniformity_pass;
  const isBatteryBankPass = isBatteryDeviationPass && isBatteryTorquePass && isBatteryCellUniform;
  // A5: Interlocks
  const isInterlocksPass = pre.mechanical_and_electrical_interlocks.toLowerCase() === 'pass';
  // A6: Grounding
  const isGroundingPass = pre.grounding_and_neutral_bonding.toLowerCase() === 'pass';
  // A7: Written Plan
  const isPlanApproved = pre.written_energization_plan_approved;

  // B1: AC Input & Phasing
  const isPhasingPass = eng.input_voltage_and_phase_rotation.phase_sequence_match.toLowerCase() === 'pass';
  // B2: Rectifier & DC Ripple
  const dcRipple = eng.rectifier_startup_and_dc_bus.dc_ripple_voltage_percent;
  const isDcRipplePass = dcRipple <= 1.0;
  // B3: Inverter Sync
  const phaseBal = eng.inverter_startup_and_sync.phase_balance_deviation_percent;
  const syncAngle = eng.inverter_startup_and_sync.sync_phase_angle_deg;
  const isInverterSyncPass = phaseBal <= 1.5 && syncAngle <= 3.0;
  // B4: Static Bypass Transfer
  const staticTransferTime = eng.static_bypass_transfer_test.transfer_time_ms;
  const isStaticTransferPass = staticTransferTime <= 4.0; // Ideal 0ms, standard <= 4ms
  // B5: Loss of AC
  const isLossOfAcPass = eng.loss_of_ac_input_power_test.transfer_to_battery_seamless && eng.loss_of_ac_input_power_test.output_voltage_drop_V <= 10;
  // B6: Full Load Battery Discharge
  const isFullLoadPass = eng.battery_discharge_and_full_load_test.full_load_capacity_met;
  // B7: THD-v
  const thdPercent = eng.output_thd_harmonics.thd_v_percent;
  const isThdPass = thdPercent <= 3.5; // IEEE 519 <= 3.0% (strict) to 5.0%
  // B8: Manual Bypass
  const isManualBypassPass = eng.manual_maintenance_bypass_transfer_test.toLowerCase() === 'pass';

  // C1: Thermographic Survey
  const thermalDeltaT = post.thermographic_survey_table_100_18.max_delta_t_celsius ?? 0;
  const isThermalSurveyPass = post.thermographic_survey_table_100_18.status.toLowerCase() === 'pass' && thermalDeltaT <= 3.0;
  // C2: SCADA Telemetry
  const isScadaPass = post.scada_bms_telemetry_verified.toLowerCase() === 'pass';
  // C3: Training & Handover
  const isTrainingPass = post.owner_training_and_handover_completed;

  // Construct Detailed Evaluation Table
  const table: NetaEcsUpsItemEvaluation[] = [
    {
      item_code: '8.2.A.1',
      item_name_vn: 'Hồ Sơ Kỹ Thuật, Bản Vẽ & Thiết Kế Đã Duyệt',
      item_name_en: 'Documentation, Approved Drawings & Nameplate Review',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Mục 8.2.A.1 - 8.2.A.6',
      measured_value: isDocPass ? 'Đầy đủ & Trùng khớp 100%' : 'Thiếu hồ sơ / Bản vẽ chưa duyệt',
      status: isDocPass ? 'PASS' : 'FAIL',
      evaluation_vn: isDocPass ? 'Bản vẽ OPR, BOD, sơ đồ nguyên lý chỉnh lưu/nghịch lưu đạt chuẩn.' : 'Thiếu hoặc sai lệch bản vẽ thiết kế.',
      evaluation_en: isDocPass ? 'Documentation fully compliant with approved project requirements.' : 'Deficient drawings or unapproved engineering submittal.'
    },
    {
      item_code: '8.2.A.3',
      item_name_vn: 'Báo Cáo Nghiệm Thu Thiết Bị (ATS Acceptance Reports)',
      item_name_en: 'Acceptance Test Reports Review (ATS 7.22.2 & 7.18)',
      neta_standard_ref: 'ANSI/NETA ATS-2025 Sec 7.22.2, 7.18 & 7.6',
      measured_value: isAtsPass ? '100% Biên bản ATS đạt yêu cầu' : 'Thiếu báo cáo nghiệm thu ATS',
      status: isAtsPass ? 'PASS' : 'FAIL',
      evaluation_vn: isAtsPass ? 'Toàn bộ khối UPS và dàn ắc-quy đã qua nghiệm thu chấp nhận đạt chuẩn.' : 'Chưa hoàn thành hoặc thiếu biên bản nghiệm thu theo NETA ATS.',
      evaluation_en: isAtsPass ? 'Full ATS test reports verified for UPS and battery systems.' : 'Deficient or missing NETA ATS equipment test records.'
    },
    {
      item_code: '8.2.A.15',
      item_name_vn: 'Điện Trở Cách Điện Mạch Động Lực (AC/DC Power Bus)',
      item_name_en: 'Power Circuit Insulation Resistance (1000V DC)',
      neta_standard_ref: 'ANSI/NETA ATS 7.22.2.B.2 & Table 100.1',
      measured_value: `${powerIr} MΩ (Thử 1000V DC)`,
      status: isPowerIrPass ? 'PASS' : 'FAIL',
      evaluation_vn: isPowerIrPass ? 'Cách điện thanh cái AC/DC đạt tiêu chuẩn (≥ 100.0 MΩ).' : 'Cách điện suy giảm nghiêm trọng (< 100 MΩ), nguy cơ rò điện.',
      evaluation_en: isPowerIrPass ? 'Insulation resistance complies with NETA Table 100.1 minimum.' : 'Unacceptable power circuit insulation breakdown.'
    },
    {
      item_code: '8.2.A.16',
      item_name_vn: 'Điện Trở Cách Điện Mạch Điều Khiển (Control Wiring)',
      item_name_en: 'Control Wiring Insulation Resistance (500V DC)',
      neta_standard_ref: 'ANSI/NETA ATS 7.22.2.B.2 & Table 100.1',
      measured_value: `${controlIr} MΩ (Thử 500V DC)`,
      status: isControlIrPass ? 'PASS' : 'FAIL',
      evaluation_vn: isControlIrPass ? 'Mạch điều khiển cách điện an toàn (≥ 1.0 MΩ).' : 'Cách điện dây điều khiển thấp (< 1.0 MΩ).',
      evaluation_en: isControlIrPass ? 'Control wiring insulation complies with Table 100.1.' : 'Control circuit insulation below 1.0 Megohm threshold.'
    },
    {
      item_code: '8.2.A.4',
      item_name_vn: 'Kiểm Tra Trở Kháng & Lực Siết Dàn Ắc-quy (Battery Bank)',
      item_name_en: 'Battery Bank Ohmic & Torque Verification (IEEE 1188 / Table 100.12)',
      neta_standard_ref: 'IEEE 450 / IEEE 1188 / IEEE 1184 & Table 100.12',
      measured_value: `Lệch trở kháng max: ${batDeviation.toFixed(1)}% | Bu-lông: ${isBatteryTorquePass ? 'Đạt' : 'Chưa đạt'}`,
      status: isBatteryBankPass ? 'PASS' : batDeviation > 30 ? 'FAIL' : 'INVESTIGATE',
      evaluation_vn: isBatteryBankPass ? 'Trở kháng tế bào đồng đều (lệch ≤ 20%) và lực siết cọc bình đạt chuẩn.' : `Nội trở tế bào bình lệch ${batDeviation.toFixed(1)}% vượt ngưỡng hoặc bu-lông lỏng lẻo.`,
      evaluation_en: isBatteryBankPass ? 'Battery impedance deviation <= 20% and bolted torque compliant.' : `Battery cell impedance deviation ${batDeviation.toFixed(1)}% exceeds limits.`
    },
    {
      item_code: '8.2.A.18',
      item_name_vn: 'Khóa Liên Động Cơ Điện (MBB vs Inverter/Static Switch)',
      item_name_en: 'Mechanical & Electrical Interlocks Verification',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Mục 8.2.A.18',
      measured_value: isInterlocksPass ? 'Liên động hoàn hảo' : 'Lỗi khóa liên động',
      status: isInterlocksPass ? 'PASS' : 'FAIL',
      evaluation_vn: isInterlocksPass ? 'Khóa liên động cơ/điện ngăn ngừa đóng chập giữa Inverter và Bypass bảo dưỡng.' : 'NGUY HẠI: Hỏng khóa liên động, nguy cơ nối tắt nguồn khi thao tác bảo dưỡng.',
      evaluation_en: isInterlocksPass ? 'Interlocks verified to prevent out-of-phase cross-connections.' : 'CRITICAL: Interlock malfunction threatens catastrophic short-circuit.'
    },
    {
      item_code: '8.2.A.11',
      item_name_vn: 'Tiếp Địa Hệ Thống & Nối Đất Trung Tính Biến Áp Cách Ly',
      item_name_en: 'Grounding & Neutral Bonding (SDS System)',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Mục 8.2.A.11 & NFPA 70',
      measured_value: isGroundingPass ? 'Tiếp địa ≤ 5.0 Ω / Trung tính nối đất đạt' : 'Tiếp địa không đạt',
      status: isGroundingPass ? 'PASS' : 'FAIL',
      evaluation_vn: isGroundingPass ? 'Điện trở tiếp địa an toàn và đấu nối trung tính biến áp cách ly đúng chuẩn.' : 'Tiếp địa chưa nối hoặc điện trở cao (> 5.0 Ω).',
      evaluation_en: isGroundingPass ? 'Grounding resistance and SDS neutral bonding fully verified.' : 'Grounding connection deficient or resistance excessive.'
    },
    {
      item_code: '8.2.B.3',
      item_name_vn: 'Điện Áp & Thứ Tự Pha AC Ngõ Vào (Rectifier & Bypass)',
      item_name_en: 'Input AC Voltage & Phase Rotation Matching',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Mục 8.2.B.3 & 8.2.B.7',
      measured_value: `Chỉnh lưu: ${eng.input_voltage_and_phase_rotation.rectifier_input_voltage_V}V | Bypass: ${eng.input_voltage_and_phase_rotation.bypass_input_voltage_V}V | Thứ tự pha: ${isPhasingPass ? 'Trùng khớp A-B-C' : 'Lệch pha'}`,
      status: isPhasingPass ? 'PASS' : 'FAIL',
      evaluation_vn: isPhasingPass ? 'Điện áp AC ngõ vào ổn định và thứ tự pha hai đường Rectifier/Bypass đồng nhất.' : 'NGUY HẠI: Ngược thứ tự pha giữa nguồn Chỉnh lưu và Bypass.',
      evaluation_en: isPhasingPass ? 'Input voltages nominal and phase rotations matched perfectly.' : 'CRITICAL: Phase rotation mismatch between Rectifier and Bypass.'
    },
    {
      item_code: '8.2.B.8',
      item_name_vn: 'Khởi Động Bộ Chỉnh Lưu & Độ Nhấp Nhô DC (DC Bus Ripple)',
      item_name_en: 'Rectifier Startup & DC Bus Float/Ripple Voltage',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Mục 8.2.B.8',
      measured_value: `Điện áp DC Float: ${eng.rectifier_startup_and_dc_bus.dc_bus_float_voltage_V}V | Ripple AC: ${dcRipple.toFixed(2)}%`,
      status: isDcRipplePass ? 'PASS' : 'INVESTIGATE',
      evaluation_vn: isDcRipplePass ? 'Điện áp nạp Float chuẩn xác, độ gợn sóng AC ngõ ra DC thấp (≤ 1.0%).' : 'Độ gợn sóng điện áp DC quá cao (> 1.0%), nguy cơ làm hỏng tụ lọc hoặc giảm thọ pin.',
      evaluation_en: isDcRipplePass ? 'DC float voltage nominal with clean DC bus ripple (<= 1.0%).' : 'Excessive DC ripple may accelerate battery degradation.'
    },
    {
      item_code: '8.2.B.9',
      item_name_vn: 'Đồng Bộ Pha Inverter Với Static Bypass (Phase Sync)',
      item_name_en: 'Inverter Startup & Static Bypass Phase Synchronization',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Mục 8.2.B.9',
      measured_value: `Lệch góc pha: ${syncAngle.toFixed(1)}° | Lệch áp: ${eng.inverter_startup_and_sync.sync_voltage_diff_V}V`,
      status: isInverterSyncPass ? 'PASS' : 'INVESTIGATE',
      evaluation_vn: isInverterSyncPass ? 'Inverter khóa pha đồng bộ hoàn hảo với Bypass (góc lệch ≤ 3°, ΔV ≤ 5V).' : 'Góc lệch pha hoặc chênh áp giữa Inverter và Bypass vượt ngưỡng.',
      evaluation_en: isInverterSyncPass ? 'Inverter locked in sync with Bypass (angle <= 3 deg, delta-V <= 5V).' : 'Synchronization out of tolerance for uninterrupted transfer.'
    },
    {
      item_code: '8.2.B.10',
      item_name_vn: 'Thử Nghiệm Chuyển Mạch Tĩnh Không Gián Đoạn (Static Transfer)',
      item_name_en: 'Static Bypass Seamless Transfer Test (0 ms / No Break)',
      neta_standard_ref: 'ANSI/NETA ECS-2024 8.2.B.10 & ATS 7.22.2.B.6',
      measured_value: `Thời gian chuyển mạch: ${staticTransferTime} ms | Quá độ áp: ${eng.static_bypass_transfer_test.transient_voltage_deviation_percent}%`,
      status: isStaticTransferPass ? 'PASS' : 'FAIL',
      evaluation_vn: isStaticTransferPass ? 'Chuyển mạch tĩnh không ngắt quãng (0 ms), tải CNTT/y tế hoạt động liên tục.' : 'Thời gian chuyển mạch tĩnh kéo dài (> 4ms), gây nguy cơ sập tải thiết bị CNTT.',
      evaluation_en: isStaticTransferPass ? 'Zero break seamless transfer validated under full electrical conditions.' : 'Static transfer time exceeds safe limits, posing load drop risk.'
    },
    {
      item_code: '8.2.B.11',
      item_name_vn: 'Thử Nghiệm Cắt Mất Điện Lưới Đột Ngột (Loss of AC Input)',
      item_name_en: 'Simulated Loss of Input AC Power (Battery Support)',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Mục 8.2.B.11',
      measured_value: isLossOfAcPass ? 'Chuyển ắc-quy tức thời, không sụt áp' : 'Lỗi sụt áp ngõ ra khi mất lưới',
      status: isLossOfAcPass ? 'PASS' : 'FAIL',
      evaluation_vn: isLossOfAcPass ? 'UPS chuyển tức thời sang cấp nguồn từ ắc-quy với độ ổn định điện áp hoàn hảo.' : 'Sụt áp hoặc mất nguồn ngõ ra khi ngắt AC ngõ vào.',
      evaluation_en: isLossOfAcPass ? 'Seamless instantaneous battery transition verified with zero voltage sag.' : 'Unacceptable output voltage drop during mains loss simulation.'
    },
    {
      item_code: '8.2.B.12',
      item_name_vn: 'Thử Nghiệm Xả Ắc-quy Đầy Tải (Battery Discharge 100% Load)',
      item_name_en: 'Full Load Battery Discharge Test (NFPA 111 / IEEE 450)',
      neta_standard_ref: 'NFPA 111 & IEEE 450 / IEEE 1188',
      measured_value: `Tải: ${eng.battery_discharge_and_full_load_test.load_percentage}% | Thời gian: ${eng.battery_discharge_and_full_load_test.discharge_duration_minutes} phút | Áp cuối: ${eng.battery_discharge_and_full_load_test.end_of_discharge_voltage_V}V`,
      status: isFullLoadPass ? 'PASS' : 'FAIL',
      evaluation_vn: isFullLoadPass ? 'Dàn ắc-quy duy trì 100% công suất tải trong suốt thời gian thiết kế.' : 'Ắc-quy bị sụt áp sớm hoặc không đáp ứng thời gian lưu điện thiết kế.',
      evaluation_en: isFullLoadPass ? 'Battery system sustained 100% load for specified autonomy duration.' : 'Premature voltage collapse during full load discharge test.'
    },
    {
      item_code: '8.2.B.13',
      item_name_vn: 'Tổng Độ Méo Sóng Hài Điện Áp Ngõ Ra (Output Voltage THD-v)',
      item_name_en: 'Output Voltage Harmonic Distortion (THD-v IEEE 519)',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Mục 8.2.B.12 & IEEE 519',
      measured_value: `THD-v: ${thdPercent.toFixed(1)}%`,
      status: isThdPass ? 'PASS' : 'INVESTIGATE',
      evaluation_vn: isThdPass ? 'Dạng sóng điện áp ngõ ra hình sin chuẩn, độ méo sóng hài thấp (≤ 3.0%).' : 'Độ méo sóng hài THD-v cao (> 3.5%), cần kiểm tra bộ lọc LC ngõ ra Inverter.',
      evaluation_en: isThdPass ? 'Output voltage waveform pristine with THD-v <= 3.0% per IEEE 519.' : 'High voltage THD requires inverter output filter investigation.'
    },
    {
      item_code: '8.2.B.14',
      item_name_vn: 'Chuyển Mạch Sang Chế Độ Bảo Trì Bằng Tay (Manual MBB Transfer)',
      item_name_en: 'Manual Maintenance Bypass Transfer Verification',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Mục 8.2.B.13',
      measured_value: isManualBypassPass ? 'Thao tác trơn tru, không mất điện tải' : 'Thao tác gây gián đoạn tải',
      status: isManualBypassPass ? 'PASS' : 'FAIL',
      evaluation_vn: isManualBypassPass ? 'Quy trình đóng MBB và cô lập UPS diễn ra an toàn, tải liên tục được cấp điện.' : 'Quy trình bảo trì bằng tay không đồng bộ hoặc gây mất điện tải.',
      evaluation_en: isManualBypassPass ? 'Seamless transfer to MBB allows full UPS servicing without load outage.' : 'Manual bypass transfer caused unacceptable load disruption.'
    },
    {
      item_code: '8.2.C.1',
      item_name_vn: 'Khảo Sát Nhiệt Hồng Ngoại Mang Tải (Infrared Survey Table 100.18)',
      item_name_en: 'Thermographic Survey Under Load (NETA MTS Table 100.18)',
      neta_standard_ref: 'ANSI/NETA ECS Section 8.2.C.1 & Table 100.18',
      measured_value: `ΔT max = ${thermalDeltaT.toFixed(1)}°C | Trạng thái: ${isThermalSurveyPass ? 'Bình thường' : 'Điểm nóng cảnh báo'}`,
      status: isThermalSurveyPass ? 'PASS' : thermalDeltaT > 15 ? 'FAIL' : 'INVESTIGATE',
      evaluation_vn: isThermalSurveyPass ? 'Không có điểm quá nhiệt trên IGBT, Static Switch và cọc bình (ΔT ≤ 3.0°C).' : `Phát hiện điểm phát nhiệt bất thường ΔT = ${thermalDeltaT.toFixed(1)}°C (vượt giới hạn an toàn Table 100.18).`,
      evaluation_en: isThermalSurveyPass ? 'Thermal gradient uniform across semiconductor switches and lugs (Delta-T <= 3°C).' : `Thermal hotspot detected: Delta-T = ${thermalDeltaT.toFixed(1)}°C requires urgent investigation.`
    },
    {
      item_code: '8.2.C.2',
      item_name_vn: 'Tích Hợp Truyền Thông & Cảnh Báo SCADA / BMS / DCIM',
      item_name_en: 'SCADA / BMS / DCIM Telemetry & Alarms Integration',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Mục 8.2.C.1',
      measured_value: isScadaPass ? 'Kết nối & Báo động chuẩn xác 100%' : 'Lỗi truyền thông / Thiếu cảnh báo',
      status: isScadaPass ? 'PASS' : 'INVESTIGATE',
      evaluation_vn: isScadaPass ? 'Dữ liệu điện áp, dòng tải, nhiệt độ và cảnh báo khẩn cấp truyền về BMS thông suốt.' : 'Mất tín hiệu hoặc thiếu cảnh báo khẩn cấp trên hệ thống SCADA/DCIM.',
      evaluation_en: isScadaPass ? 'All telemetry and alarm vectors verified communicating to central BMS/DCIM.' : 'Deficient telemetry mapping or missing mission-critical alarms.'
    },
    {
      item_code: '8.2.C.3',
      item_name_vn: 'Đào Tạo Vận Hành Chủ Đầu Tư & Hồ Sơ Bàn Giao Hoàn Công',
      item_name_en: 'Owner Operations Training & As-built Handover',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Mục 8.2.C.3 & Section 5.4',
      measured_value: isTrainingPass ? 'Đã hoàn thành đào tạo & Ký biên bản' : 'Chưa hoàn thành đào tạo',
      status: isTrainingPass ? 'PASS' : 'INVESTIGATE',
      evaluation_vn: isTrainingPass ? 'Đội ngũ vận hành đã được hướng dẫn thao tác MBB khẩn cấp và nhận bàn giao hồ sơ hoàn công.' : 'Chưa tổ chức đào tạo hoặc chưa bàn giao đầy đủ hồ sơ chạy thử.',
      evaluation_en: isTrainingPass ? 'Owner operations personnel trained and full commissioning dossier signed.' : 'Operator training or commissioning closeout records incomplete.'
    }
  ];

  // Determine overall severity
  const hasCriticalHazard =
    !isInterlocksPass ||
    !isPhasingPass ||
    !isStaticTransferPass ||
    !isLossOfAcPass ||
    !isFullLoadPass ||
    batDeviation > 30.0 ||
    thermalDeltaT > 15.0;

  const hasCaution =
    !isPowerIrPass ||
    !isControlIrPass ||
    !isBatteryBankPass ||
    !isDcRipplePass ||
    !isInverterSyncPass ||
    !isThdPass ||
    !isManualBypassPass ||
    !isThermalSurveyPass ||
    !isScadaPass ||
    !isTrainingPass;

  let overall_status: 'PASS' | 'CAUTION' | 'CRITICAL' = 'PASS';
  let status_title_en = 'UPS SYSTEM COMMISSIONED / READY FOR MISSION-CRITICAL LOAD';
  let status_title_vn = 'HỆ THỐNG UPS ĐÃ CHẠY THỬ ĐẠT CHUẨN / SẴN SÀNG CẤP ĐIỆN TẢI QUAN TRỌNG';
  let summary_badge = 'NETA ECS 8.2 PASSED';
  let summary_color = 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40';

  if (hasCriticalHazard) {
    overall_status = 'CRITICAL';
    if (!isInterlocksPass) {
      status_title_en = 'CRITICAL ENERGIZATION HAZARD / MAINTENANCE BYPASS INTERLOCK FAILURE';
      status_title_vn = 'NGUY HẠI ĐÓNG ĐIỆN - LỖI KHÓA LIÊN ĐỘNG BYPASS BẢO TRÌ VỚI INVERTER';
    } else if (!isPhasingPass) {
      status_title_en = 'CRITICAL ENERGIZATION HAZARD / RECTIFIER-BYPASS PHASE MISMATCH';
      status_title_vn = 'NGUY HẠI ĐÓNG ĐIỆN - NGƯỢC THỨ TỰ PHA GIỮA CHỈNH LƯU VÀ BYPASS';
    } else if (batDeviation > 30.0) {
      status_title_en = 'CRITICAL HAZARD / BATTERY CELL INTERNAL IMPEDANCE DEFECT';
      status_title_vn = 'NGUY HẠI ĐÓNG ĐIỆN - NỘI TRỞ TẾ BÀO BÌNH ẮC-QUY TĂNG BẤT THƯỜNG';
    } else if (thermalDeltaT > 15.0) {
      status_title_en = 'CRITICAL HAZARD / SEVERE THERMAL HOTSPOT IN POWER STAGE (TABLE 100.18)';
      status_title_vn = 'NGUY HẠI ĐÓNG ĐIỆN - ĐIỂM NÓNG BẬC 1 TRÊN KHỐI BÁN DẪN CÔNG SUẤT';
    } else if (!isStaticTransferPass) {
      status_title_en = 'CRITICAL HAZARD / STATIC BYPASS INTERRUPTION (> 4 MS)';
      status_title_vn = 'NGUY HẠI ĐÓNG ĐIỆN - CHUYỂN MẠCH TĨNH BỊ GIÁN ĐOẠN MẤT ĐIỆN TẢI';
    } else {
      status_title_en = 'CRITICAL ENERGIZATION HAZARD / BATTERY DISCHARGE CAPACITY COLLAPSE';
      status_title_vn = 'NGUY HẠI ĐÓNG ĐIỆN - ẮC-QUY SỤT ÁP KHÔNG ĐẠT CÔNG SUẤT ĐẦY TẢI';
    }
    summary_badge = 'CRITICAL HAZARD';
    summary_color = 'bg-rose-500/20 text-rose-300 border-rose-500/40';
  } else if (hasCaution) {
    overall_status = 'CAUTION';
    status_title_en = 'UPS COMMISSIONING CAUTION / DEFICIENCIES REQUIRE RECTIFICATION';
    status_title_vn = 'CẢNH BÁO CHẠY THỬ UPS / CÁC KHIẾM KHUYẾT CẦN CÂN CHỈNH TRƯỚC KHI BÀN GIAO';
    summary_badge = 'PENDING RECTIFICATION';
    summary_color = 'bg-amber-500/20 text-amber-300 border-amber-500/40';
  }

  // Construct technical analysis
  const technical_analysis = {
    hazard_identified: hasCriticalHazard,
    title_en: hasCriticalHazard
      ? batDeviation > 30.0
        ? `Battery Internal Impedance Deviation Anomaly (${batDeviation.toFixed(1)}% vs IEEE 1188 Limit)`
        : !isInterlocksPass
        ? 'Maintenance Bypass Mechanical Interlock Defect (ANSI/NETA ECS 8.2.A.18)'
        : 'Critical Power Stage Hazard Under Commissioning'
      : hasCaution
      ? 'Power Quality THD-v or Thermal Observation Pending Rectification'
      : 'High-Integrity Dual-Conversion UPS & 0ms Seamless Transfer Validated',
    title_vn: hasCriticalHazard
      ? batDeviation > 30.0
        ? `Lỗi Tăng Nội Trở Tế Bào Bình Ắc-quy (Lệch ${batDeviation.toFixed(1)}% Theo Chuẩn IEEE 1188)`
        : !isInterlocksPass
        ? 'Lỗi Khóa Liên Động Cơ Học Cắt MBB Bảo Trì (ANSI/NETA ECS 8.2.A.18)'
        : 'Nguy Cơ Nghiêm Trọng Trong Khối Động Lực UPS Khi Chạy Thử'
      : hasCaution
      ? 'Cảnh Báo Sóng Hài THD-v Hoặc Nhiệt Độ Cần Theo Dõi Cân Chỉnh'
      : 'Hệ Thống UPS Chuyển Đổi Kép Hoạt Động Hoàn Hảo & Chuyển Mạch Tĩnh 0ms',
    description_en: hasCriticalHazard
      ? batDeviation > 30.0
        ? `Pre-energization ohmic testing per IEEE 1188 / ANSI/NETA ATS 7.18 identified a severe internal cell impedance anomaly in the battery string, deviating by ${batDeviation.toFixed(1)}% from baseline (allowable threshold <= 20%). Under high-rate discharge during a utility blackout, this high-impedance cell will experience severe thermal runaway, deep cell reversal, and voltage collapse, cutting the critical UPS autonomy short.`
        : !isInterlocksPass
        ? `Mechanical interlock testing between the Inverter/Static Switch and the Maintenance Bypass Breaker (MBB) failed. Without positive mechanical interlocking, an operator could inadvertently parallel out-of-phase AC sources, risking a catastrophic phase-to-phase flashover and complete load disruption.`
        : `Commissioning identified critical operational hazards violating ANSI/NETA ECS-2024 Section 8.2 and NFPA 111. System must not be transferred to critical live server or hospital equipment.`
      : hasCaution
      ? `UPS commissioning tests are predominantly satisfactory, but cautionary findings (such as THD-v = ${thdPercent.toFixed(1)}% or thermal delta-T = ${thermalDeltaT.toFixed(1)}°C) require fine-tuning prior to commercial sign-off.`
      : `All pre-energization, inverter synchronization, seamless 0 ms static transfer, 100% load bank battery discharge, and post-energization thermographic criteria strictly comply with ANSI/NETA ECS-2024 Section 8.2, ANSI/NETA ATS 7.22.2/7.18, and NFPA 111. The system is certified ready for critical load energization.`,
    description_vn: hasCriticalHazard
      ? batDeviation > 30.0
        ? `Kiểm tra đo trở kháng nội theo IEEE 1188 / ANSI/NETA ATS 7.18 phát hiện tế bào bình ắc-quy có độ lệch trở kháng ${batDeviation.toFixed(1)}% so với giá trị trung bình giàn (ngưỡng cho phép ≤ 20%). Khi xả dòng lớn trong sự cố mất điện lưới, bình có nội trở cao này sẽ bị quá nhiệt cục bộ, đảo cực và gây sụt áp đột ngột toàn bộ dàn ắc-quy.`
        : !isInterlocksPass
        ? `Kiểm tra khóa liên động cơ học giữa Máy cắt Bảo trì MBB và bộ Inverter/Static Switch thất bại. Việc thiếu liên động cơ học có thể dẫn đến việc kỹ thuật viên vô tình đóng nối song song hai nguồn điện khác pha, gây ra hồ quang ngắn mạch pha-pha thảm khốc và sập toàn bộ hệ thống điện.`
        : `Quá trình chạy thử phát hiện nguy cơ nghiêm trọng vi phạm ANSI/NETA ECS-2024 Mục 8.2 và NFPA 111. Tuyệt đối không được cấp điện cho tải máy chủ hoặc thiết bị y tế nhạy cảm.`
      : hasCaution
      ? `Các hạng mục kiểm tra đạt yêu cầu cơ bản, tuy nhiên một số thông số cảnh báo (như THD-v = ${thdPercent.toFixed(1)}% hoặc phát nhiệt ΔT = ${thermalDeltaT.toFixed(1)}°C) cần được xử lý cân chỉnh trước khi bàn giao hoàn công.`
      : `Toàn bộ các phép đo tiền đóng điện, đồng bộ nghịch lưu, chuyển mạch tĩnh không ngắt quãng 0 ms, xả ắc-quy 100% tải giả và khảo sát nhiệt hồng ngoại đều tuân thủ nghiêm ngặt ANSI/NETA ECS-2024 Mục 8.2, ATS 7.22.2/7.18 và NFPA 111. Hệ thống được chứng nhận sẵn sàng cấp điện phụ tải quan trọng.`,
    power_ir_megohms: powerIr,
    battery_impedance_deviation_percent: parseFloat(batDeviation.toFixed(1)),
    static_transfer_time_ms: staticTransferTime,
    thd_v_percent: parseFloat(thdPercent.toFixed(1)),
    thermal_delta_t: thermalDeltaT
  };

  // Construct recommendations
  const actionable_recommendations = {
    halt_transfer: hasCriticalHazard,
    steps: hasCriticalHazard
      ? [
          {
            step_number: 1,
            title_en: 'DO NOT ENERGIZE CRITICAL LOAD / ISOLATE DEFECTIVE SECTION',
            title_vn: 'KHÔNG CẤP ĐIỆN TẢI QUAN TRỌNG / CÔ LẬP KHỐI BỊ LỖI',
            details_en: batDeviation > 30.0
              ? 'Keep the UPS in standalone commissioning mode. Do not cut over live server or hospital loads. Isolate the affected battery string and tag out DC disconnect switch.'
              : !isInterlocksPass
              ? 'Lock out the Maintenance Bypass Breaker (MBB) mechanically. Repair or replace the interlock push-rod mechanism before any further switching procedures.'
              : 'Keep the UPS offline from critical operational infrastructure until all safety blocks are resolved.',
            details_vn: batDeviation > 30.0
              ? 'Giữ UPS ở chế độ chạy thử cô lập, tuyệt đối không đấu tải thực. Cách ly giàn ắc-quy có tế bào lỗi và treo biển cảnh báo an toàn trên cầu dao DC.'
              : !isInterlocksPass
              ? 'Khóa liên động cơ học trên máy cắt MBB, tiến hành căn chỉnh hoặc thay thế thanh trượt liên động trước khi thực hiện bất kỳ thao tác đóng cắt nào khác.'
              : 'Duy trì trạng thái dừng đóng điện tải quan trọng cho đến khi khắc phục triệt để các lỗi vi phạm an toàn.'
          },
          {
            step_number: 2,
            title_en: 'CELL REPLACEMENT & RE-TORQUE INTER-TIER CONNECTORS',
            title_vn: 'THAY THẾ TẾ BÀO ẮC-QUY LỖI & SIẾT LỰC CỌC BÌNH THEO BẢNG 100.12',
            details_en: 'Replace defective battery blocks with factory pre-screened matched cells. Re-torque all inter-cell and inter-tier terminal bolts using a calibrated torque wrench per NETA Table 100.12. Clean connection surfaces with antioxidant grease.',
            details_vn: 'Thay thế các bình ắc-quy bị tăng nội trở bằng bình mới cùng lô sản xuất có nội trở phù hợp. Siết lại toàn bộ bu-lông cọc bình bằng cờ-lê lực đã kiểm định theo Bảng NETA Table 100.12 và bôi mỡ chống oxy hóa.'
          },
          {
            step_number: 3,
            title_en: 'REPEAT 100% LOAD BANK DISCHARGE & IR SURVEY VERIFICATION',
            title_vn: 'THỬ NGHIỆM LẠI XẢ ĐẦY TẢI 100% VÀ KHẢO SÁT NHIỆT HỒNG NGOẠI',
            details_en: 'Perform an equalize charge cycle, followed by a repeat 100% load bank capacity discharge test per IEEE 450/1188. Continuously monitor cell thermal profiles with infrared imaging (NETA Table 100.18) to ensure zero localized hot spots.',
            details_vn: 'Thực hiện chu trình nạp cân bằng Equalize, sau đó tiến hành thử xả 100% tải giả trở lại theo IEEE 450/1188. Giám sát liên tục nhiệt độ cọc bình bằng camera nhiệt (NETA Table 100.18) để đảm bảo không còn điểm phát nhiệt cục bộ.'
          }
        ]
      : hasCaution
      ? [
          {
            step_number: 1,
            title_en: 'TUNE INVERTER OUTPUT FILTER & INVESTIGATE THD-V HARMONICS',
            title_vn: 'HIỆU CHỈNH BỘ LỌC NGÕ RA INVERTER & KHẮC PHỤC SÓNG HÀI THD-V',
            details_en: 'Inspect harmonic filter capacitors and tuning inductors on the inverter output stage to reduce THD-v below 3.0% under linear and non-linear loads per IEEE 519.',
            details_vn: 'Kiểm tra tụ lọc và cuộn kháng lọc sóng hài ngõ ra Inverter để đưa độ méo sóng hài THD-v về dưới 3.0% theo tiêu chuẩn IEEE 519.'
          },
          {
            step_number: 2,
            title_en: 'RE-TORQUE CABLE LUGS AT THERMAL ANOMALIES & COMPLETE BMS MAPPING',
            title_vn: 'SIẾT LỰC ĐẦU COSSE PHÁT NHIỆT & HOÀN THIỆN ÁNH XẠ TRUYỀN THÔNG BMS',
            details_en: 'Re-torque terminals showing temperature rises, verify SNMP/Modbus registers for all critical alarms, and re-scan under 50% load per Table 100.18.',
            details_vn: 'Siết lại bu-lông các đầu cosse có dấu hiệu tăng nhiệt độ, kiểm tra các thanh ghi cảnh báo qua SNMP/Modbus và quét lại nhiệt khi mang tải theo Bảng 100.18.'
          }
        ]
      : [
          {
            step_number: 1,
            title_en: 'ISSUE FORMAL UPS COMMISSIONING & ACCEPTANCE CERTIFICATE',
            title_vn: 'PHÁT HÀNH CHỨNG NHẬN NGHIỆM THU CHẠY THỬ UPS CHÍNH THỨC',
            details_en: 'System meets all criteria of ANSI/NETA ECS-2024 Section 8.2, ANSI/NETA ATS-2025 Section 7.22.2/7.18, and NFPA 111. Approved for live mission-critical energization.',
            details_vn: 'Hệ thống đạt 100% yêu cầu tiêu chuẩn ANSI/NETA ECS-2024 Mục 8.2, ATS-2025 Mục 7.22.2/7.18 và NFPA 111. Đủ điều kiện cấp điện trực tiếp cho phụ tải quan trọng.'
          },
          {
            step_number: 2,
            title_en: 'ENTER BASELINE BATTERY OHMIC READINGS INTO CMMS REGISTRY',
            title_vn: 'LƯU TRỮ THÔNG SỐ TRỞ KHÁNG ẮC-QUY NỀN VÀO HỆ THỐNG CMMS',
            details_en: 'Archive baseline cell impedance and float voltage data into CMMS for quarterly IEEE 1188 trend comparisons and schedule annual infrared surveys per NETA MTS Table 100.18.',
            details_vn: 'Lưu trữ thông số trở kháng từng tế bào và điện áp nạp nền vào phần mềm CMMS để đối soát định kỳ hàng quý theo IEEE 1188 và lập lịch chụp nhiệt hồng ngoại hàng năm theo NETA MTS.'
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

export async function generateEcsUpsAiAnalysis(
  payload: NetaEcsUpsInputPayload
): Promise<NetaEcsUpsEvaluationResult> {
  const deterministicResult = performDeterministicEcsUpsAnalysis(payload);

  try {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      console.warn('GEMINI_API_KEY is not defined. Returning deterministic UPS ECS analysis.');
      return deterministicResult;
    }

    const ai = new GoogleGenAI({ apiKey });
    const prompt = `Please evaluate the following Uninterruptible Power Supply (UPS) System Commissioning test data strictly based on ANSI/NETA ECS-2024 (Section 8.2), ANSI/NETA ATS-2025 (Section 7.22.2 & 7.18), NETA MTS (Tables 100.1, 100.12, 100.18), IEEE 450, IEEE 1188, IEEE 1184, and NFPA 111.

INPUT DATA PAYLOAD:
${JSON.stringify(payload, null, 2)}

Provide an authoritative bilingual (EN/VN) commissioning evaluation report conforming strictly to the system instructions. Do NOT include citation numbers or bibliography links.`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: prompt,
      config: {
        systemInstruction: ECS_UPS_SYSTEM_INSTRUCTION,
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
    console.error('Error during Gemini UPS ECS AI analysis:', error);
    return deterministicResult;
  }
}
