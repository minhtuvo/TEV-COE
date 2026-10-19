import { GoogleGenAI } from '@google/genai';

export interface NetaEcsSolarPvInputPayload {
  tev_coe_header: {
    module: string;
    project_name: string;
    system_tag: string;
    system_dc_voltage_rating_V: number; // 1000 or 1500
    fse_name: string;
    commissioning_date: string;
  };
  pre_energization_8_1_a: {
    acceptance_test_reports_reviewed_ats_7_29: boolean;
    dc_polarity_and_blocking_diodes: string; // 'Pass' | 'Fail' | description
    pv_string_dry_insulation_resistance_megohms: {
      voltage_rating_v?: number;
      string_01_to_03_avg: number;
      string_04: number;
      test_voltage_dc?: number;
      status?: string;
    };
    pv_string_wet_insulation_resistance_megohms?: {
      measured_val: number;
      status: string;
    };
    string_voc_measurements_volts: {
      string_01_to_03_avg_voc: number;
      string_04_voc: number;
      max_voc_deviation_percent: number;
      status?: string;
    };
    iv_curve_tracing_performance: {
      string_01_to_03_stc_pmax_percent: number;
      string_04_stc_pmax_percent: number;
      status?: string;
    };
    grounding_resistance_ohms: number;
    bolted_connection_torque_table_100_12: string; // 'Pass' | 'Fail'
    inverter_gateway_das_settings: string; // 'Pass' | 'Fail'
    written_energization_plan_approved: boolean;
  };
  energization_8_1_b: {
    safety_access_restricted: boolean;
    ac_phase_sequence_check: string; // 'Pass' | 'Fail'
    anti_islanding_protection_test_ieee_1547: {
      trip_time_seconds: number;
      status: string; // 'Pass' | 'Fail'
    };
    grid_code_and_power_factor_tuning: string; // 'Pass' | 'Fail'
  };
  post_energization_8_1_c: {
    das_weather_sensors_baseline: string; // 'Pass' | 'Fail'
    thermographic_survey_table_100_18: {
      max_delta_t_celsius?: number;
      status: string; // 'Pass' | 'Investigate'
      details?: string;
    };
    nerc_mod_025_2_reactive_power_verification: string; // 'Pass' | 'Fail'
  };
}

export interface NetaEcsSolarPvItemEvaluation {
  item_code: string;
  item_name_vn: string;
  item_name_en: string;
  neta_standard_ref: string;
  measured_value: string;
  status: 'PASS' | 'FAIL' | 'INVESTIGATE';
  evaluation_vn: string;
  evaluation_en: string;
}

export interface NetaEcsSolarPvEvaluationResult {
  overall_status: 'PASS' | 'CAUTION' | 'CRITICAL';
  status_title_en: string;
  status_title_vn: string;
  summary_badge: string;
  summary_color: string;
  evaluation_table: NetaEcsSolarPvItemEvaluation[];
  technical_analysis: {
    hazard_identified: boolean;
    title_en: string;
    title_vn: string;
    description_en: string;
    description_vn: string;
    dry_insulation_min_val: number;
    voc_deviation_percent: number;
    pmax_stc_percent: number;
    anti_islanding_time: number;
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

export const ECS_SOLAR_PV_SYSTEM_INSTRUCTION = `You are the TEV COE Analytical Engine — an AI Expert Assistant specializing in Solar Photovoltaic (PV) System Commissioning and Acceptance Testing, operating strictly according to ANSI/NETA ECS-2024 (Section 8.1), ANSI/NETA ATS-2025 (Section 7.29), IEC 62446-1, IEEE 1547, and NERC MOD-025-2 standards. Your task is to evaluate Solar PV system commissioning test data submitted by Field Service Engineers (FSE) and output a bilingual (English / Vietnamese EN/VN) diagnostic report without citing any source reference numbers.

EVALUATION RULES & REFERENCE STANDARDS (ANSI/NETA ECS-2024 SECTION 8.1 & NETA TABLES):

1. PRE-ENERGIZATION REQUIREMENTS (SECTION 8.1.A):
- Documentation & Acceptance Testing (8.1.A.1 - 8.1.A.6 & 8.1.A.31): Verify OPR/BOD compliance and 100% acceptance testing completed per ANSI/NETA ATS 7.29.
- Polarity & Blocking Diodes (8.1.A.28 - 8.1.A.29): Verify 100% correct DC polarity (+ to +, - to -) and blocking diode operation.
- Dry & Wet Insulation Resistance (8.1.A.30 / Table 100.1 / IEC 62446-1): Minimum IR for 1000V DC >= 1.0 MΩ; 1500V DC >= 5.0 MΩ.
- Voc, Isc & I-V Curve Performance (8.1.A.32 - 8.1.A.34 / ASTM E1036): String Voc and Isc deviation MUST be <= 5% across identical strings. Extrapolated STC Pmax MUST be >= 95% of module rating.
- Grounding & Bonding (8.1.A.27 / Section 7.13 / IEEE 81): Grounding electrode system resistance MUST be <= 5.0 Ohms.
- Bolt Tightness (8.1.A.8 / Table 100.12): Verify bolted connection tightness per NETA Table 100.12.
- Settings, Interlocks & Written Energization Plan (8.1.A.13, 8.1.A.25, 8.1.A.35 - 8.1.A.42): Verify inverter settings, gateways, DAS, E-Stop, and approved written energization plan.

2. ENERGIZATION REQUIREMENTS (SECTION 8.1.B):
- Safety & Access Control (8.1.B.1 - 8.1.B.6): Restrict access and follow approved energization plan steps.
- Voltage, Current & AC Phasing (8.1.B.7 - 8.1.B.9): Verify correct AC phase sequence and voltage/current vectors.
- Anti-Islanding Protection (8.1.B.10 / IEEE 1547): Verify mandatory anti-islanding trip function upon grid loss.
- Grid Code & Power Tuning (8.1.B.11): Verify grid code compliance, power factor Cos(phi), and stabilizer tuning.

3. POST-ENERGIZATION REQUIREMENTS (SECTION 8.1.C):
- Baseline Parameters & DAS (8.1.C.1): Establish baseline criteria for weather sensors and DAS logging.
- Thermographic Survey (8.1.C.2 / Section 9 / Table 100.18): Perform IR survey under load per NETA MTS Table 100.18.
- NERC MOD-025-2 Verification (8.1.C.5): Verify reactive power capability (P-Q curve) per NERC MOD-025-2.

AI RESPONSE STRUCTURE (ALL IN BILINGUAL EN/VN):
1. OVERALL SEVERITY STATUS: PASS (COMMISSIONED), CAUTION (PENDING DEFICIENCIES), or CRITICAL ENERGIZATION HAZARD.
2. DETAILED PV COMMISSIONING EVALUATION TABLE: Compare measured parameters against NETA ECS Section 8.1 and NETA Tables 100.1 / 100.12 / 100.18.
3. COMMISSIONING & SYSTEM FUNCTION ANALYSIS: Technical summary of PV array insulation, string Voc/Isc/I-V curves, anti-islanding, thermography, and reactive power capability.
4. ACTIONABLE RECOMMENDATIONS FOR FSE: Step-by-step guidance for energization clearance or deficiency correction.`;

export function performDeterministicEcsSolarPvAnalysis(payload: NetaEcsSolarPvInputPayload): NetaEcsSolarPvEvaluationResult {
  const pA = payload.pre_energization_8_1_a;
  const pB = payload.energization_8_1_b;
  const pC = payload.post_energization_8_1_c;
  const voltageRating = payload.tev_coe_header?.system_dc_voltage_rating_V || 1500;

  // Criteria thresholds
  // 1000V DC: >= 1.0 MΩ, 1500V DC: >= 5.0 MΩ
  const minRequiredDryIr = voltageRating >= 1500 ? 5.0 : 1.0;
  const string04DryIr = pA.pv_string_dry_insulation_resistance_megohms?.string_04 ?? 0;
  const string0103AvgDryIr = pA.pv_string_dry_insulation_resistance_megohms?.string_01_to_03_avg ?? 40.0;
  const minDryIr = Math.min(string04DryIr, string0103AvgDryIr);

  const isDryIrPass = string04DryIr >= minRequiredDryIr && string0103AvgDryIr >= minRequiredDryIr;

  // Polarity
  const isPolarityPass = pA.dc_polarity_and_blocking_diodes?.toLowerCase().includes('pass') &&
                         !pA.dc_polarity_and_blocking_diodes?.toLowerCase().includes('reverse') &&
                         !pA.dc_polarity_and_blocking_diodes?.toLowerCase().includes('ngược');

  // Voc deviation <= 5.0%
  const vocDeviation = pA.string_voc_measurements_volts?.max_voc_deviation_percent ?? 0;
  const isVocPass = vocDeviation <= 5.0;

  // I-V Curve Pmax STC >= 95%
  const pmaxStcString04 = pA.iv_curve_tracing_performance?.string_04_stc_pmax_percent ?? 0;
  const isIvCurvePass = pmaxStcString04 >= 95.0;

  // Grounding <= 5.0 Ohms
  const groundingR = pA.grounding_resistance_ohms ?? 1.0;
  const isGroundingPass = groundingR <= 5.0;

  // Torque
  const isTorquePass = pA.bolted_connection_torque_table_100_12?.toLowerCase().includes('pass');

  // Acceptance Reports
  const isAtsPass = pA.acceptance_test_reports_reviewed_ats_7_29 === true;

  // Settings & Interlocks
  const isSettingsPass = pA.inverter_gateway_das_settings?.toLowerCase().includes('pass');

  // Written Energization Plan
  const isPlanPass = pA.written_energization_plan_approved === true;

  // Energization Phase
  const isSafetyPass = pB.safety_access_restricted === true;
  const isPhasingPass = pB.ac_phase_sequence_check?.toLowerCase().includes('pass');

  // Anti-Islanding IEEE 1547: must trip within 2.0s (typical 0.1-0.5s)
  const tripTime = typeof pB.anti_islanding_protection_test_ieee_1547 === 'object'
    ? pB.anti_islanding_protection_test_ieee_1547.trip_time_seconds
    : 0.15;
  const isAntiIslandingPass = (typeof pB.anti_islanding_protection_test_ieee_1547 === 'object'
    ? pB.anti_islanding_protection_test_ieee_1547.status?.toLowerCase().includes('pass')
    : true) && tripTime > 0 && tripTime <= 2.0;

  // Grid Code
  const isGridCodePass = pB.grid_code_and_power_factor_tuning?.toLowerCase().includes('pass');

  // Post-Energization
  const isDasPass = pC.das_weather_sensors_baseline?.toLowerCase().includes('pass');

  // Thermographic Survey Table 100.18
  const isThermalPass = typeof pC.thermographic_survey_table_100_18 === 'object'
    ? !pC.thermographic_survey_table_100_18.status?.toLowerCase().includes('investigate') &&
      !pC.thermographic_survey_table_100_18.status?.toLowerCase().includes('fail')
    : !String(pC.thermographic_survey_table_100_18).toLowerCase().includes('investigate');

  // NERC MOD-025-2
  const isNercPass = pC.nerc_mod_025_2_reactive_power_verification?.toLowerCase().includes('pass');

  // Build Item Evaluations
  const table: NetaEcsSolarPvItemEvaluation[] = [
    {
      item_code: 'A1',
      item_name_vn: 'Rà soát Hồ sơ, Bản vẽ & Catalog / Documentation & Drawing Review',
      item_name_en: 'Documentation & Drawing Review',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Section 8.1.A.1 - 8.1.A.6',
      measured_value: isAtsPass ? 'Khớp 100% OPR/BOD, Layout String & Sơ đồ tủ gom' : 'Thiếu hồ sơ thiết kế hoặc chưa đối soát bản vẽ',
      status: isAtsPass ? 'PASS' : 'FAIL',
      evaluation_vn: isAtsPass ? 'Trùng khớp 100% bản vẽ thiết kế và thiết kế kỹ thuật được phê duyệt.' : 'Sai lệch hồ sơ thiết kế hoặc thiếu tài liệu.',
      evaluation_en: isAtsPass ? '100% compliant with approved OPR/BOD and string drawings.' : 'Documentation discrepancy or missing drawing approval.'
    },
    {
      item_code: 'A2',
      item_name_vn: 'Thử nghiệm Acceptance Testing Thiết bị / Acceptance Test Reports',
      item_name_en: 'Acceptance Test Reports Review',
      neta_standard_ref: 'ANSI/NETA ATS-2025 Sec 7.29 / NETA ECS 8.1.A.3 & 8.1.A.31',
      measured_value: isAtsPass ? '100% thiết bị (Inverter, Tấm pin, SPD, CB) đạt chuẩn ATS 7.29' : 'Thiếu báo cáo nghiệm thu ATS hoặc phát hiện lỗi thiết bị',
      status: isAtsPass ? 'PASS' : 'FAIL',
      evaluation_vn: isAtsPass ? '100% thiết bị đã qua thử nghiệm nghiệm thu đạt chuẩn NETA ATS.' : 'Chưa hoàn thành nghiệm thu xuất xưởng hoặc thử nghiệm tại site.',
      evaluation_en: isAtsPass ? 'All PV array components tested per ANSI/NETA ATS 7.29.' : 'Missing ATS acceptance test reports.'
    },
    {
      item_code: 'A3',
      item_name_vn: 'Thử Cực tính & Diode Chặn / Polarity & Blocking Diode Tests',
      item_name_en: 'Polarity & Blocking Diode Tests',
      neta_standard_ref: 'ANSI/NETA ECS 8.1.A.28 & 8.1.A.29 / NETA ATS 7.29.B.1',
      measured_value: isPolarityPass ? 'Cực tính DC (+/-) đúng 100%, Diode chặn bình thường' : 'Phát hiện ngược cực tính DC hoặc diode chặn ngắn mạch',
      status: isPolarityPass ? 'PASS' : 'FAIL',
      evaluation_vn: isPolarityPass ? 'Cực tính chính xác 100% (không ngược cực); Diode chặn hoạt động tốt.' : 'NGUY HIỂM: Ngược cực tính DC gây nổ Inverter hoặc cháy cáp!',
      evaluation_en: isPolarityPass ? '100% correct DC polarity verified; blocking diodes functional.' : 'CRITICAL: DC reverse polarity hazard!'
    },
    {
      item_code: 'A4',
      item_name_vn: `Điện trở Cách điện Khô & Ẩm (${voltageRating}V DC) / Dry & Wet Insulation`,
      item_name_en: `Dry & Wet Insulation Resistance (${voltageRating}V DC)`,
      neta_standard_ref: `NETA ECS 8.1.A.30 & Table 100.1 / IEC 62446-1 (Min >= ${minRequiredDryIr} MΩ)`,
      measured_value: `String 01-03: ${string0103AvgDryIr} MΩ | String 04: ${string04DryIr} MΩ`,
      status: isDryIrPass ? 'PASS' : 'FAIL',
      evaluation_vn: isDryIrPass
        ? `Đạt chuẩn NETA Table 100.1 (Min ${minDryIr} MΩ >= ${minRequiredDryIr} MΩ).`
        : `String 04 đo được ${string04DryIr} MΩ (< ${minRequiredDryIr} MΩ). Nguy cơ rò đất DC và chạm vỏ giá đỡ!`,
      evaluation_en: isDryIrPass
        ? `Insulation exceeds NETA Table 100.1 requirement (>= ${minRequiredDryIr} MΩ).`
        : `String 04 IR (${string04DryIr} MΩ) below ${minRequiredDryIr} MΩ requirement. Active DC ground fault hazard!`
    },
    {
      item_code: 'A5',
      item_name_vn: 'Điện áp Mở Mạch Voc & Quét Đường cong I-V / String Voc & I-V Curve',
      item_name_en: 'String Voc Deviation & I-V Curve Tracing',
      neta_standard_ref: 'ANSI/NETA ECS 8.1.A.32 - 8.1.A.34 / ASTM E1036',
      measured_value: `ΔVoc: ${vocDeviation}% (String 04 Voc = ${pA.string_voc_measurements_volts?.string_04_voc}V) | Pmax STC: ${pmaxStcString04}%`,
      status: isVocPass && isIvCurvePass ? 'PASS' : 'FAIL',
      evaluation_vn: isVocPass && isIvCurvePass
        ? `Voc độ lệch ${vocDeviation}% (<= 5%) và Pmax STC ${pmaxStcString04}% (>= 95%) đạt chuẩn.`
        : `Lỗi String 04: Lệch Voc ${vocDeviation}% (> 5%) và Pmax ${pmaxStcString04}% (< 95%). Bypass diode kích hoạt do hỏng cell!`,
      evaluation_en: isVocPass && isIvCurvePass
        ? `Voc deviation (${vocDeviation}%) and I-V Pmax (${pmaxStcString04}%) within allowable limits.`
        : `String 04 Voc deviation (${vocDeviation}%) exceeds 5% and Pmax dropped to ${pmaxStcString04}%. Bypass diodes engaged.`
    },
    {
      item_code: 'A6',
      item_name_vn: 'Tiếp Địa Khung Giá Đỡ & Cọc Đất Trạm / Grounding & Bonding',
      item_name_en: 'Grounding & Bonding System',
      neta_standard_ref: 'ANSI/NETA ECS 8.1.A.27 & Sec 7.13 / IEEE 81 (Max <= 5.0Ω)',
      measured_value: `${groundingR} Ohm (Liên kết khung tấm pin 100% liên tục)`,
      status: isGroundingPass ? 'PASS' : 'FAIL',
      evaluation_vn: isGroundingPass ? `Điện trở cọc đất trạm đạt chuẩn (${groundingR}Ω <= 5.0Ω).` : `Điện trở tiếp địa quá cao (${groundingR}Ω > 5.0Ω), nguy hiểm an toàn giật điện!`,
      evaluation_en: isGroundingPass ? `Grounding resistance (${groundingR}Ω) compliant with IEEE 81 (<= 5.0Ω).` : `High ground resistance (${groundingR}Ω > 5.0Ω).`
    },
    {
      item_code: 'A7',
      item_name_vn: 'Lực Siết Bu-lông Thanh Cái & Cọc Đấu / Bolted Connection Tightness',
      item_name_en: 'Bolted Connection Tightness',
      neta_standard_ref: 'ANSI/NETA ATS/MTS Table 100.12 (Bolt Torque)',
      measured_value: isTorquePass ? 'Đã siết cờ-lê lực hiệu chuẩn và niêm phong sơn Table 100.12' : 'Chưa siết đủ lực bu-lông thanh cái DC/AC',
      status: isTorquePass ? 'PASS' : 'FAIL',
      evaluation_vn: isTorquePass ? 'Mối nối cơ khí thanh cái tủ gom đạt chuẩn lực siết.' : 'Chưa đạt lực siết Table 100.12, nguy cơ phóng tia lửa điện phát nhiệt.',
      evaluation_en: isTorquePass ? 'Calibrated torque verified per Table 100.12.' : 'Torque verification deficient, loose terminal hazard.'
    },
    {
      item_code: 'A8',
      item_name_vn: 'Cài Đặt Biến Tần Inverter, DAS & Dừng Khẩn Cấp / Inverters & Interlocks',
      item_name_en: 'Inverter Settings, DAS & Emergency Stop',
      neta_standard_ref: 'ANSI/NETA ECS 8.1.A.13, 8.1.A.25, 8.1.A.35 & 8.1.A.38',
      measured_value: isSettingsPass ? 'Nạp file cấu hình phê duyệt, Gateway online, E-Stop ngắt tin cậy' : 'Sai thông số cài đặt Inverter hoặc E-Stop mất tác dụng',
      status: isSettingsPass ? 'PASS' : 'FAIL',
      evaluation_vn: isSettingsPass ? 'Inverter và cổng DAS cấu hình chuẩn xác, E-Stop an toàn.' : 'Lỗi cài đặt biến tần hoặc hệ thống dừng khẩn cấp E-Stop.',
      evaluation_en: isSettingsPass ? 'Inverter settings and E-stop logic verified.' : 'Inverter settings or E-stop logic deficient.'
    },
    {
      item_code: 'A9',
      item_name_vn: 'Kế Hoạch Đóng Điện Văn Bản & Xóa Cảnh Báo / Written Energization Plan',
      item_name_en: 'Approved Energization Plan & Alarms Cleared',
      neta_standard_ref: 'ANSI/NETA ECS Section 8.1.A.39 - 8.1.A.42',
      measured_value: isPlanPass ? 'Đã duyệt kế hoạch đóng điện, xóa sạch lịch sử lỗi Inverter' : 'Chưa ký duyệt kế hoạch đóng điện hoặc còn lỗi tồn đọng',
      status: isPlanPass ? 'PASS' : 'FAIL',
      evaluation_vn: isPlanPass ? 'Đủ điều kiện pháp lý và an toàn để triển khai đóng điện.' : 'Thiếu kế hoạch đóng điện bằng văn bản được duyệt.',
      evaluation_en: isPlanPass ? 'Approved energization plan active and all alarms cleared.' : 'Missing approved written energization plan.'
    },
    {
      item_code: 'B1',
      item_name_vn: 'Khoanh Vùng An Toàn & Trình Tự Đóng Điện / Access Control & Safety',
      item_name_en: 'Safety Access Control & Energization Protocol',
      neta_standard_ref: 'ANSI/NETA ECS Section 8.1.B.1 & 8.1.B.6',
      measured_value: isSafetyPass ? 'Khoanh vùng rào chắn, trang bị PPE hồ quang và tuân thủ quy trình' : 'Vi phạm vùng an toàn hoặc thao tác sai trình tự',
      status: isSafetyPass ? 'PASS' : 'FAIL',
      evaluation_vn: isSafetyPass ? 'Khu vực biến tần/tủ gom được cô lập an toàn tuyệt đối.' : 'Chưa khoanh vùng cách ly an toàn đóng điện.',
      evaluation_en: isSafetyPass ? 'Restricted perimeter and PPE enforced.' : 'Safety perimeter not secured.'
    },
    {
      item_code: 'B2',
      item_name_vn: 'Thứ Tự Pha AC & Điện Áp Đầu Ra / Voltage, Current & Phasing Checks',
      item_name_en: 'AC Voltage, Current & Phasing Checks',
      neta_standard_ref: 'ANSI/NETA ECS Section 8.1.B.7 - 8.1.B.9',
      measured_value: isPhasingPass ? 'Thứ tự pha AC trùng khớp lưới (A-B-C thuận), điện áp cân bằng' : 'Phát hiện ngược thứ tự pha AC hoặc mất cân bằng điện áp',
      status: isPhasingPass ? 'PASS' : 'FAIL',
      evaluation_vn: isPhasingPass ? 'Điện áp AC và góc pha đồng bộ 100% với lưới điện.' : 'Lỗi ngược thứ tự pha AC hoặc mất cân bằng nghiêm trọng.',
      evaluation_en: isPhasingPass ? 'AC phase rotation matches grid connection.' : 'AC phase rotation mismatch.'
    },
    {
      item_code: 'B3',
      item_name_vn: 'Thử Nghiệm Chống Đảo Lưới / Anti-Islanding Protection Test',
      item_name_en: 'Anti-Islanding Protection Test (IEEE 1547)',
      neta_standard_ref: 'ANSI/NETA ECS 8.1.B.10 & IEEE 1547',
      measured_value: isAntiIslandingPass ? `Inverter tự ngắt phát điện trong ${tripTime}s khi ngắt AC (<= 2.0s)` : `Lỗi ngắt đảo lưới: thời gian ngắt ${tripTime}s (> 2.0s) hoặc không ngắt`,
      status: isAntiIslandingPass ? 'PASS' : 'FAIL',
      evaluation_vn: isAntiIslandingPass ? `Chức năng chống đảo lưới bảo vệ công nhân vận hành đạt chuẩn IEEE 1547 (${tripTime}s).` : 'NGUY HIỂM: Inverter tiếp tục phát điện khi mất lưới, vi phạm nghiêm trọng IEEE 1547!',
      evaluation_en: isAntiIslandingPass ? `Inverter tripped in ${tripTime}s upon islanding condition per IEEE 1547.` : 'CRITICAL: Anti-islanding failure! Inverter failed to cease energizing grid.'
    },
    {
      item_code: 'B4',
      item_name_vn: 'Hòa Lưới, Grid Code & Điều Chỉnh Công Suất / Grid Code & Tuning',
      item_name_en: 'Grid Code Compliance & Power Factor Tuning',
      neta_standard_ref: 'ANSI/NETA ECS Section 8.1.B.11',
      measured_value: isGridCodePass ? 'Hòa lưới êm dịu, Cosφ = 0.99 lead/lag, tuân thủ Grid Code' : 'Dao động công suất hoặc vi phạm tiêu chuẩn hòa lưới',
      status: isGridCodePass ? 'PASS' : 'FAIL',
      evaluation_vn: isGridCodePass ? 'Đồng bộ hòa lưới ổn định, công suất phản kháng đáp ứng lệnh điều độ.' : 'Lỗi hệ số công suất hoặc dao động điện áp hòa lưới.',
      evaluation_en: isGridCodePass ? 'Grid code requirements and PF tuning verified.' : 'Grid code non-compliance.'
    },
    {
      item_code: 'C1',
      item_name_vn: 'Dữ Liệu Vận Hành Nền & Cảm Biến Thời Tiết / Baseline Parameters & DAS',
      item_name_en: 'Baseline Parameters & DAS Weather Station',
      neta_standard_ref: 'ANSI/NETA ECS Section 8.1.C.1',
      measured_value: isDasPass ? 'Pyranometer bức xạ và RTD nhiệt độ pin cập nhật chính xác' : 'Sai số cảm biến thời tiết hoặc dữ liệu DAS gián đoạn',
      status: isDasPass ? 'PASS' : 'FAIL',
      evaluation_vn: isDasPass ? 'Xác lập đường cơ sở hiệu suất PR (Performance Ratio) tin cậy.' : 'Cần hiệu chuẩn lại cảm biến bức xạ quang năng.',
      evaluation_en: isDasPass ? 'Baseline meteorological and DAS logging verified.' : 'Weather station calibration deficient.'
    },
    {
      item_code: 'C2',
      item_name_vn: 'Khảo Sát Nhiệt Hồng Ngoại Dàn Pin & Tủ Điện / Thermographic Survey',
      item_name_en: 'Thermographic Survey Under Load',
      neta_standard_ref: 'ANSI/NETA ECS 8.1.C.2, Sec 9 & NETA MTS Table 100.18',
      measured_value: isThermalPass ? 'Phân bố nhiệt đồng đều, không có Hotspot trên cell pin & MC4' : 'Phát hiện điểm nóng Hotspot cell pin hoặc phát nhiệt đầu nối MC4',
      status: isThermalPass ? 'PASS' : 'INVESTIGATE',
      evaluation_vn: isThermalPass ? 'Không có điểm quá nhiệt cục bộ, đầu cosse và jack nối MC4 tiếp xúc tốt.' : 'Cảnh báo điểm nóng Hotspot theo NETA MTS Table 100.18, nguy cơ cháy nổ tấm pin.',
      evaluation_en: isThermalPass ? 'Normal thermal pattern across modules and combiner terminals.' : 'Thermal anomalies / hotspots detected per Table 100.18.'
    },
    {
      item_code: 'C3',
      item_name_vn: 'Khả Năng Phát Công Suất Phản Kháng / NERC MOD-025-2 Verification',
      item_name_en: 'NERC MOD-025-2 Reactive Power Capability',
      neta_standard_ref: 'ANSI/NETA ECS Section 8.1.C.5 & NERC MOD-025-2',
      measured_value: isNercPass ? 'Xác nhận dải công suất phản kháng P-Q đáp ứng 100% biểu đồ' : 'Không đạt dải công suất phản kháng theo yêu cầu điều độ',
      status: isNercPass ? 'PASS' : 'FAIL',
      evaluation_vn: isNercPass ? 'Đáp ứng hoàn toàn tiêu chuẩn NERC MOD-025-2 về hỗ trợ điện áp lưới.' : 'Chưa đáp ứng dải huy động công suất phản kháng.',
      evaluation_en: isNercPass ? 'Reactive power capability curve complies with NERC MOD-025-2.' : 'Reactive power capability deficient.'
    }
  ];

  // Determine overall severity
  const hasCriticalHazard = !isPolarityPass || !isAntiIslandingPass || !isDryIrPass;
  const hasCaution = !isVocPass || !isIvCurvePass || !isThermalPass || !isTorquePass || !isAtsPass || !isGroundingPass || !isPlanPass;

  let overall_status: 'PASS' | 'CAUTION' | 'CRITICAL' = 'PASS';
  let status_title_en = 'SYSTEM FULLY COMMISSIONED / READY FOR COMMERCIAL OPERATION';
  let status_title_vn = 'HỆ THỐNG ĐÃ NGHIỆM THU ĐẠT CHUẨN / SẴN SÀNG VẬN HÀNH THƯƠNG MẠI';
  let summary_badge = 'NETA ECS 8.1 PASSED';
  let summary_color = 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40';

  if (hasCriticalHazard) {
    overall_status = 'CRITICAL';
    if (!isDryIrPass && !isVocPass) {
      status_title_en = 'CRITICAL ENERGIZATION HAZARD / STRING 04 INSULATION & VOC ANOMALY';
      status_title_vn = 'NGUY HẠI ĐÓNG ĐIỆN - LỖI CÁCH ĐIỆN & SỤT ÁP STRING 04 (CHẠM ĐẤT DC)';
    } else if (!isPolarityPass) {
      status_title_en = 'CRITICAL ENERGIZATION HAZARD / DC REVERSE POLARITY DETECTED';
      status_title_vn = 'NGUY CƠ ĐÓNG ĐIỆN CỰC KỲ NGUY HIỂM / NGƯỢC CỰC TÍNH DC';
    } else {
      status_title_en = 'CRITICAL ENERGIZATION HAZARD / ANTI-ISLANDING TRIP FAILURE';
      status_title_vn = 'NGUY CƠ ĐÓNG ĐIỆN / LỖI BẢO VỆ CHỐNG ĐẢO LƯỚI THEO IEEE 1547';
    }
    summary_badge = 'CRITICAL HAZARD';
    summary_color = 'bg-rose-500/20 text-rose-300 border-rose-500/40';
  } else if (hasCaution) {
    overall_status = 'CAUTION';
    status_title_en = 'COMMISSIONING CAUTION / DEFICIENCIES REQUIRE RECTIFICATION';
    status_title_vn = 'CẢNH BÁO CHẠY THỬ / CÁC KHIẾM KHUYẾT CẦN KHẮC PHỤC TRƯỚC KHI ĐÓNG ĐIỆN';
    summary_badge = 'PENDING RECTIFICATION';
    summary_color = 'bg-amber-500/20 text-amber-300 border-amber-500/40';
  }

  // Construct technical analysis
  const technical_analysis = {
    hazard_identified: hasCriticalHazard,
    title_en: hasCriticalHazard
      ? 'Active DC Ground Fault & String Degradation Hazard (NETA ECS 8.1.A.30 / IEC 62446-1)'
      : hasCaution
      ? 'Minor Deficiencies & Thermal Warnings Pending Clearance'
      : 'Optimal PV Array Performance & Commissioning Integrity',
    title_vn: hasCriticalHazard
      ? 'Sự Cố Suy Giảm Cách Điện Khô & Sụt Áp Voc String 04 (Chạm Đất DC)'
      : hasCaution
      ? 'Cảnh Báo Khiếm Khuyết Nhiệt Hoặc Lực Siết Bu-lông Cần Khắc Phục'
      : 'Hệ Thống Điện Mặt Trời Hoàn Hảo & Đạt Chuẩn Nghiệm Thu',
    description_en: hasCriticalHazard
      ? `Pre-energization testing under NETA ECS Section 8.1.A.30 revealed a severe insulation failure on String 04, measuring only ${string04DryIr} Megohm at ${voltageRating}V DC (violating the ${minRequiredDryIr} Megohm threshold in NETA Table 100.1). Concurrently, String 04 open-circuit voltage Voc dropped to ${pA.string_voc_measurements_volts?.string_04_voc}V (a ${vocDeviation}% deviation exceeding the 5% allowable limit), and its I-V curve showed a stepped notch resulting in a Pmax drop to ${pmaxStcString04}%. This indicates an active DC ground fault on String 04, causing bypass diodes to engage and shorting out damaged module cells.`
      : hasCaution
      ? `System commissioning parameters are largely acceptable, but attention is required for items flagged with CAUTION (such as Voc deviation ${vocDeviation}%, bolt torque or thermographic hotspots per Table 100.18).`
      : `All pre-energization, energization, and post-energization tests comply with ANSI/NETA ECS-2024 (Section 8.1), ATS-2025 Section 7.29, IEEE 1547, and NERC MOD-025-2. Array dry insulation is ${minDryIr} MΩ (>= ${minRequiredDryIr} MΩ), Voc deviation is ${vocDeviation}% (<= 5%), anti-islanding trip time is ${tripTime}s, and reactive power is validated.`,
    description_vn: hasCriticalHazard
      ? `Kiểm tra tiền đóng điện theo Mục 8.1.A.30 của NETA ECS phát hiện sự cố suy giảm cách điện nghiêm trọng trên String 04, chỉ đo được ${string04DryIr} Megohm ở điện áp ${voltageRating}V DC (vi phạm ngưỡng tối thiểu ${minRequiredDryIr} Megohm của Bảng NETA Table 100.1). Đồng thời, điện áp mở mạch Voc của String 04 sụt xuống ${pA.string_voc_measurements_volts?.string_04_voc}V (lệch ${vocDeviation}%, vượt giới hạn cho phép 5%), và đường cong I-V bị méo dạng nấc làm Pmax sụt giảm xuống ${pmaxStcString04}%. Điều này xác nhận có sự cố chạm đất DC trên String 04, ép diode bypass kích hoạt và làm ngắt chuỗi cell tấm pin bị tổn thương.`
      : hasCaution
      ? `Các thông số nghiệm thu cơ bản đáp ứng, tuy nhiên còn tồn tại khiếm khuyết cần khắc phục (như độ lệch Voc ${vocDeviation}%, lực siết bu-lông hoặc phát nhiệt hồng ngoại theo Bảng 100.18).`
      : `Toàn bộ các phép đo kiểm tra tiền đóng điện, trong đóng điện và sau đóng điện đều hoàn toàn tuân thủ tiêu chuẩn ANSI/NETA ECS-2024 Mục 8.1, ATS-2025 Mục 7.29, IEEE 1547 và NERC MOD-025-2. Điện trở cách điện khô đạt ${minDryIr} MΩ (>= ${minRequiredDryIr} MΩ), độ lệch Voc là ${vocDeviation}% (<= 5%), chống đảo lưới ngắt trong ${tripTime}s và công suất phản kháng đạt chuẩn.`,
    dry_insulation_min_val: minDryIr,
    voc_deviation_percent: vocDeviation,
    pmax_stc_percent: pmaxStcString04,
    anti_islanding_time: tripTime
  };

  // Recommendations
  const actionable_recommendations = {
    halt_energization: hasCriticalHazard,
    steps: hasCriticalHazard
      ? [
          {
            step_number: 1,
            title_en: 'ISOLATE STRING 04 IMMEDIATELY (OPEN DC SWITCH & REMOVE FUSES)',
            title_vn: 'CÔ LẬP CẦU CHÌ DC STRING 04 NGAY LẬP TỨC (NGẮT CÔNG TẮC & RÚT CẦU CHÌ)',
            details_en: 'Open the DC disconnect switch and remove DC fuses for String 04 at the combiner box immediately. Proceed with energization of Block 01 for all healthy strings (String 01 to 03) while keeping String 04 isolated.',
            details_vn: 'Ngắt công tắc DC và rút cầu chì DC cô lập String 04 tại tủ gom ngay lập tức. Cho phép tiến hành đóng điện Block 01 cho các string bình thường (String 01 đến 03) trong khi duy trì cô lập String 04.'
          },
          {
            step_number: 2,
            title_en: 'TROUBLESHOOTING PROTOCOL & CABLE FAULT LOCALIZATION',
            title_vn: 'QUY TRÌNH KHẮC PHỤC SỰ CỐ CÁP DC & KIỂM TRA TẤM PIN',
            details_en: 'Inspect the DC cable run of String 04 along the mounting structures to locate and repair the pinched insulation fault against the metal rails. Inspect PV modules on String 04 for thermal cell damage or shorted junction box diodes and replace defective modules.',
            details_vn: 'Kiểm tra dọc tuyến cáp DC của String 04 trên khung giá đỡ để tìm và xử lý điểm vỏ cáp bị dập kẹp vào xà thép. Kiểm tra các tấm pin trên String 04 xem có bị hư hỏng nhiệt cell hoặc hỏng diode junction box để tiến hành thay thế.'
          },
          {
            step_number: 3,
            title_en: 'RE-TESTING & COMMISSIONING CLEARANCE',
            title_vn: 'ĐO KIỂM TRA LẠI NGHIỆM THU & ĐÓNG ĐIỆN VẬN HÀNH',
            details_en: `Re-perform Dry Insulation Resistance (must be >= ${minRequiredDryIr} Megohms), String Voc (deviation <= 5%), and I-V Curve Tracing (STC Pmax >= 95%) on String 04. Upon passing, insert DC fuses and close the DC switch to integrate String 04 into full commercial operation.`,
            details_vn: `Thực hiện thử nghiệm lại Điện trở cách điện khô (BẮT BUỘC >= ${minRequiredDryIr} Megohms), Điện áp Voc (độ lệch <= 5%) và Quét đường cong I-V (Pmax >= 95%) cho String 04. Sau khi đạt, cắm lại cầu chì DC và đóng công tắc để đưa String 04 vào vận hành thương mại hoàn chỉnh.`
          }
        ]
      : hasCaution
      ? [
          {
            step_number: 1,
            title_en: 'RECTIFY PENDING MECHANICAL & ELECTRICAL DEFICIENCIES',
            title_vn: 'KHẮC PHỤC CÁC KHIẾM KHUYẾT CƠ KHÍ & ĐIỆN TỒN ĐỌNG',
            details_en: 'Re-torque bolted connections per Table 100.12 using calibrated torque wrenches. Investigate thermographic hotspots on combiner terminals or MC4 connectors.',
            details_vn: 'Siết lại lực bu-lông theo Bảng Table 100.12 bằng cờ-lê lực đã kiểm định. Khảo sát xử lý các điểm nóng hồng ngoại tại đầu cosse tủ gom hoặc jack MC4.'
          },
          {
            step_number: 2,
            title_en: 'STRING VOC & I-V CURVE NORMALIZATION',
            title_vn: 'ĐỒNG NHẤT ĐIỆN ÁP VOC & ĐẶC TÍNH I-V GIỮA CÁC CHUỖI',
            details_en: 'Check for partial module shading, soiled surfaces, or mismatched module counts in strings exceeding 5% Voc deviation.',
            details_vn: 'Kiểm tra bóng che cục bộ, bụi bẩn bám trên bề mặt hoặc số lượng tấm pin sai lệch giữa các string có độ lệch Voc vượt quá 5%.'
          }
        ]
      : [
          {
            step_number: 1,
            title_en: 'ISSUE FORMAL COMMERCIAL OPERATION CLEARANCE (COD)',
            title_vn: 'CẤP CHỨNG CHỈ ĐÓNG ĐIỆN VẬN HÀNH THƯƠNG MẠI (COD)',
            details_en: 'All commissioning items under ANSI/NETA ECS-2024 Section 8.1 are satisfied. System is certified safe for continuous grid-tied commercial operation.',
            details_vn: 'Tất cả các hạng mục chạy thử theo ANSI/NETA ECS-2024 Mục 8.1 đã đạt chuẩn 100%. Xác nhận hệ thống đủ điều kiện phát điện thương mại an toàn.'
          },
          {
            step_number: 2,
            title_en: 'ESTABLISH ROUTINE O&M AND NETA MTS INSPECTION CYCLES',
            title_vn: 'THIẾT LẬP CHU KỲ BẢO TRÌ ĐỊNH KỲ O&M VÀ NETA MTS',
            details_en: 'Transition to permanent CMMS monitoring, schedule 6-month thermal imaging survey per NETA MTS Table 100.18 and annual I-V curve trending.',
            details_vn: 'Chuyển giao sang hệ thống quản lý bảo trì CMMS, lên lịch chụp nhiệt định kỳ 6 tháng/lần theo NETA MTS Table 100.18 và đo suy hao I-V hàng năm.'
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

export async function generateEcsSolarPvAiAnalysis(
  payload: NetaEcsSolarPvInputPayload
): Promise<NetaEcsSolarPvEvaluationResult> {
  const deterministicResult = performDeterministicEcsSolarPvAnalysis(payload);

  try {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      console.warn('GEMINI_API_KEY is not defined. Returning deterministic Solar PV ECS analysis.');
      return deterministicResult;
    }

    const ai = new GoogleGenAI({ apiKey });
    const prompt = `Please evaluate the following Solar Photovoltaic (PV) System Commissioning test data strictly based on ANSI/NETA ECS-2024 (Section 8.1), ANSI/NETA ATS-2025 (Section 7.29), IEC 62446-1, IEEE 1547, and NERC MOD-025-2.

INPUT DATA PAYLOAD:
${JSON.stringify(payload, null, 2)}

Provide a comprehensive, authoritative bilingual (EN/VN) commissioning evaluation report conforming strictly to the system instructions. Do NOT include citation numbers or bibliography links.`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: prompt,
      config: {
        systemInstruction: ECS_SOLAR_PV_SYSTEM_INSTRUCTION,
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
    console.error('Error during Gemini Solar PV ECS AI analysis:', error);
    return deterministicResult;
  }
}
