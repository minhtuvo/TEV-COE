import { GoogleGenAI } from '@google/genai';

export interface NetaEcsHvInputPayload {
  tev_coe_header: {
    module: string;
    project_name: string;
    system_tag: string;
    nominal_voltage_kV: number;
    fse_name: string;
    commissioning_date: string;
  };
  pre_energization_7_3_a: {
    acceptance_test_reports_reviewed: boolean;
    ehv_busbar_insulation_resistance_10000v_megohms: {
      phase_a_to_ground: number;
      phase_b_to_ground: number;
      phase_c_to_ground: number;
      status?: string;
    };
    hv_cable_ac_withstand_and_offline_pd: {
      ac_withstand_pass: boolean;
      tan_delta_rad?: number;
      offline_pd_max_pc?: number;
      status: string; // 'Pass' | 'Fail'
    };
    sf6_gas_quality_analysis_table_100_19: {
      sf6_purity_percent: number;
      moisture_content_ppmv: number;
      dew_point_celsius: number;
      so2_content_ppmv: number;
      status?: string;
    };
    transformer_230kv_oil_dga_and_fluid_table_100_4: {
      breakdown_voltage_bdv_kV: number;
      moisture_content_ppm: number;
      power_factor_20c_percent: number;
      dga_combustible_gases_ppm: {
        h2: number;
        ch4: number;
        c2h2: number;
        c2h4: number;
        c2h6: number;
        co: number;
        co2: number;
      };
      status?: string;
    };
    protection_relays_and_nerc_prc_compliance: {
      line_differential_87l_test: string; // 'Pass' | 'Fail'
      distance_21_test: string; // 'Pass' | 'Fail'
      nerc_prc_002_4_dme_ddr_logging: string; // 'Pass' | 'Fail'
      nerc_prc_019_2_excitation_limiter_coordination: string; // 'Pass' | 'Fail'
      status?: string;
    };
    substation_grounding_grid_resistance_ieee_81_ohms: number;
    bolted_connection_torque_table_100_12: string; // 'Pass' | 'Fail'
    interlocks_and_online_monitoring: string; // 'Pass' | 'Fail'
    dual_dc_battery_banks_status: string; // 'Pass' | 'Fail'
    written_energization_plan_approved: boolean;
  };
  energization_7_3_b: {
    temporary_protective_grounds_removed: boolean;
    ehv_phasing_check_bus_tie_breaker: {
      phase_a1_a2_voltage_volts: number;
      phase_b1_b2_voltage_volts: number;
      phase_c1_c2_voltage_volts: number;
      phase_angle_diff_deg?: number;
      status?: string;
    };
    relay_vector_inputs_status: string; // 'Pass' | 'Fail'
  };
  post_energization_7_3_c: {
    thermographic_survey_table_100_18: string; // 'Pass' | 'Investigate'
    gis_online_pd_survey_tables_100_23: {
      uhf_max_dBmV: number;
      acoustic_max_dB: number;
      status: string; // 'Pass' | 'Fail'
    };
    post_energization_gas_fluid_report_completed: boolean;
  };
}

export interface NetaEcsHvItemEvaluation {
  item_code: string;
  item_name_vn: string;
  item_name_en: string;
  neta_standard_ref: string;
  measured_value: string;
  status: 'PASS' | 'FAIL' | 'INVESTIGATE';
  evaluation_vn: string;
  evaluation_en: string;
}

export interface NetaEcsHvEvaluationResult {
  overall_status: 'PASS' | 'CAUTION' | 'CRITICAL';
  status_title_en: string;
  status_title_vn: string;
  summary_badge: string;
  summary_color: string;
  evaluation_table: NetaEcsHvItemEvaluation[];
  technical_analysis: {
    hazard_identified: boolean;
    title_en: string;
    title_vn: string;
    description_en: string;
    description_vn: string;
    min_insulation_mohm: number;
    sf6_moisture_ppmv: number;
    sf6_dew_point_c: number;
    transformer_bdv_kv: number;
    ground_resistance_ohms: number;
    phasing_max_delta_v: number;
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

export const ECS_HV_SYSTEM_INSTRUCTION = `You are the TEV COE Analytical Engine — an AI Expert Assistant specializing in High-Voltage and Extra-High Voltage Electrical Power System (> 100 kV and < 1,000 kV) Commissioning and Acceptance Testing, operating strictly according to ANSI/NETA ECS-2024 (Section 7.3), ANSI/NETA ATS/MTS standards, NETA Tables (100.1, 100.2, 100.3, 100.4, 100.6, 100.12, 100.18, 100.19, 100.23), IEEE 81, IEEE C57.104, and NERC PRC standards (PRC-002-4, PRC-005-6, PRC-019-2, PRC-024-3, PRC-026-1). Your task is to evaluate HV/EHV system commissioning test data submitted by Field Service Engineers (FSE) and output a bilingual (English / Vietnamese EN/VN) diagnostic report without citing any source reference numbers.

EVALUATION RULES & REFERENCE STANDARDS (ANSI/NETA ECS-2024 SECTION 7.3 & NETA TABLES):

1. PRE-ENERGIZATION REQUIREMENTS (SECTION 7.3.A):
- Documentation & Acceptance Testing (7.3.A.1 - 7.3.A.4): Verify OPR/BOD compliance and 100% acceptance testing completed per ANSI/NETA ATS. Minimum dry insulation resistance MUST comply with NETA Table 100.1 (5000V/10000V DC test, >= 10,000 MΩ for HV/EHV apparatus).
- HV Cable AC Withstand & Offline PD Test (7.3.A.5 / Table 100.6.5 & 100.6.6 / IEEE 400.3): Verify power frequency AC withstand (20-300 Hz), Tan Delta, and offline PD test per NETA Table 100.6.6.
- SF6 Gas Quality & Purity (7.3.A.35 / Table 100.19): Verify SF6 gas purity >= 98%, moisture <= 99 ppmv (dew point <= -36°C), and SO2 <= 1.0 ppmv per NETA Table 100.19.
- Transformer Fluid & DGA Analysis (7.3.A.35 / Table 100.4.1 & IEEE C57.104): Verify oil breakdown voltage (BDV >= 50 kV for > 230kV), moisture <= 10 ppm, power factor <= 0.5% at 20°C, and DGA combustible gas limits (Condition 1).
- Protection Relays & NERC Compliance (7.3.A.32 - 7.3.A.34 & NERC PRC-002-4 / PRC-005-6 / PRC-019-2 / PRC-024-3): Verify primary injection, line differential (87L), distance (21), DME/DDR, and NERC excitation limiter coordination.
- Grounding Grid & Bolt Tightness (7.3.A.7 / Section 7.13 / Table 100.12 & IEEE 81): Substation grounding grid resistance MUST be <= 1.0 Ohm (<= 0.5 Ohm recommended for EHV substations). Bolt torque MUST comply with NETA Table 100.12.
- Interlocks, Arc Flash & DC Battery System (7.3.A.15, 7.3.A.30, 7.3.A.45): Verify key interlocks, online PD/DGA sensors, dual 110V/220V DC battery banks, and approved written energization plan.

2. ENERGIZATION REQUIREMENTS (SECTION 7.3.B):
- Temporary Ground Removal (7.3.B.1 - 7.3.B.3): Verify 100% removal of temporary protective grounds on HV/EHV buses.
- HV Dual-Source Phasing Check (7.3.B.8 - 7.3.B.9): Verify zero phase angle discrepancy across dual-source HV bus tie / bus coupler breakers (phase angle diff <= 5 deg).
- Relay Vector Inputs & SCADA (7.3.B.7 - 7.3.B.10): Verify secondary current/voltage vector angles on numerical IEDs and SCADA/EMS.

3. POST-ENERGIZATION REQUIREMENTS (SECTION 7.3.C):
- Thermographic Survey (7.3.C.3 / Section 9 / Table 100.18): Perform IR survey under load per NETA MTS Table 100.18.
- Online Partial Discharge Survey (7.3.C.4 / NETA MTS Tables 100.23.1 - 100.23.5): Perform UHF/Acoustic/HFCT/TEV online PD survey on GIS, cables, and transformers under load.
- Post-Energization Fluid/Gas & Final Commissioning Report (7.3.C.1 & 5.4).

AI RESPONSE STRUCTURE (ALL IN BILINGUAL EN/VN):
1. OVERALL SEVERITY STATUS: PASS (COMMISSIONED), CAUTION (PENDING DEFICIENCIES), or CRITICAL ENERGIZATION HAZARD.
2. DETAILED HV/EHV COMMISSIONING EVALUATION TABLE: Compare measured parameters against NETA ECS Section 7.3 and NETA Tables 100.1 / 100.4 / 100.6 / 100.12 / 100.18 / 100.19 / 100.23.
3. COMMISSIONING & SYSTEM FUNCTION ANALYSIS: Technical summary of HV insulation, SF6 gas purity, transformer DGA, relay protection/NERC compliance, dual-source phasing, and online PD.
4. ACTIONABLE RECOMMENDATIONS FOR FSE: Step-by-step guidance for energization clearance or deficiency correction.`;

export function performDeterministicEcsHvAnalysis(
  payload: NetaEcsHvInputPayload
): NetaEcsHvEvaluationResult {
  const pre = payload.pre_energization_7_3_a;
  const eng = payload.energization_7_3_b;
  const post = payload.post_energization_7_3_c;

  // A1/A2: Acceptance Test Reports
  const isAtsPass = pre.acceptance_test_reports_reviewed;

  // A2: Busbar Insulation Resistance (10kV DC, Table 100.1 >= 10,000 MΩ for HV/EHV)
  const busIr = pre.ehv_busbar_insulation_resistance_10000v_megohms;
  const minBusIr = Math.min(busIr.phase_a_to_ground, busIr.phase_b_to_ground, busIr.phase_c_to_ground);
  const isBusIrPass = minBusIr >= 10000.0;

  // A3: HV Cable AC Withstand & Offline PD (Table 100.6.5/100.6.6)
  const isCablePass = pre.hv_cable_ac_withstand_and_offline_pd.ac_withstand_pass;

  // A4: SF6 Gas Quality (Table 100.19: Purity >= 98%, Moisture <= 99 ppmv / dew point <= -36°C, SO2 <= 1.0 ppmv)
  const sf6 = pre.sf6_gas_quality_analysis_table_100_19;
  const isSf6PurityPass = sf6.sf6_purity_percent >= 98.0;
  const isSf6MoisturePass = sf6.moisture_content_ppmv <= 99.0 && sf6.dew_point_celsius <= -36.0;
  const isSf6So2Pass = sf6.so2_content_ppmv <= 1.0;
  const isSf6Pass = isSf6PurityPass && isSf6MoisturePass && isSf6So2Pass;

  // A5: Transformer Oil BDV & DGA (Table 100.4.1 BDV >= 50kV for > 230kV, moisture <= 10 ppm, PF <= 0.5%, C2H2 <= 1 ppm)
  const oil = pre.transformer_230kv_oil_dga_and_fluid_table_100_4;
  const isOilBdvPass = oil.breakdown_voltage_bdv_kV >= 50.0;
  const isOilMoisturePass = oil.moisture_content_ppm <= 10.0;
  const isOilPfPass = oil.power_factor_20c_percent <= 0.5;
  const isDgaC2h2Pass = oil.dga_combustible_gases_ppm.c2h2 <= 1.0; // Acetylene is critical for arcing
  const isOilDgaPass = isOilBdvPass && isOilMoisturePass && isOilPfPass && isDgaC2h2Pass;

  // A6: Protection Relays & NERC PRC Compliance
  const relays = pre.protection_relays_and_nerc_prc_compliance;
  const isRelayPass =
    relays.line_differential_87l_test.toLowerCase() === 'pass' &&
    relays.distance_21_test.toLowerCase() === 'pass' &&
    relays.nerc_prc_002_4_dme_ddr_logging.toLowerCase() === 'pass' &&
    relays.nerc_prc_019_2_excitation_limiter_coordination.toLowerCase() === 'pass';

  // A7: Ground Grid Resistance (IEEE 81: <= 1.0 Ohm, <= 0.5 Ohm for EHV)
  const groundRes = pre.substation_grounding_grid_resistance_ieee_81_ohms;
  const isGroundPass = groundRes <= 1.0;

  // A7: Bolted Torque (Table 100.12)
  const isTorquePass = pre.bolted_connection_torque_table_100_12.toLowerCase().includes('pass');

  // A8: Interlocks & Online Monitoring
  const isInterlocksPass = pre.interlocks_and_online_monitoring.toLowerCase().includes('pass');

  // A9: Dual DC Battery Banks & Written Plan
  const isDcBatteryPass = pre.dual_dc_battery_banks_status.toLowerCase().includes('pass');
  const isPlanApproved = pre.written_energization_plan_approved;

  // B1: Temporary Protective Ground Removal (CRITICAL)
  const isGroundsRemoved = eng.temporary_protective_grounds_removed;

  // B2: EHV Phasing Check (Delta V <= 60V, Angle <= 5°)
  const phasing = eng.ehv_phasing_check_bus_tie_breaker;
  const maxDeltaV = Math.max(
    phasing.phase_a1_a2_voltage_volts,
    phasing.phase_b1_b2_voltage_volts,
    phasing.phase_c1_c2_voltage_volts
  );
  const isPhasingPass = maxDeltaV <= 60; // Delta-V between same phases across dual sources should be near zero

  // B3: Relay Vector Inputs
  const isVectorPass = eng.relay_vector_inputs_status.toLowerCase().includes('pass');

  // C1: Thermographic Survey
  const isThermalPass = post.thermographic_survey_table_100_18.toLowerCase().includes('pass');

  // C2: Online PD Survey (NETA Table 100.23: UHF <= 6.0 dBmV, Acoustic <= 2.0 dB)
  const pd = post.gis_online_pd_survey_tables_100_23;
  const isOnlinePdPass = pd.status.toLowerCase().includes('pass') && pd.uhf_max_dBmV <= 6.0;

  // C3: Post-Energization Fluid/Gas Report
  const isPostReportPass = post.post_energization_gas_fluid_report_completed;

  // Construct Detailed Checklist Table
  const table: NetaEcsHvItemEvaluation[] = [
    {
      item_code: '7.3.A.1',
      item_name_vn: 'Rà Soát Hồ Sơ Kỹ Thuật & Báo Cáo Nghiệm Thu Thiết Bị',
      item_name_en: 'Documentation, Nameplate Review & Acceptance Reports',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Mục 7.3.A.1 - 7.3.A.4',
      measured_value: isAtsPass ? '100% Hồ sơ & Báo cáo ATS Đạt' : 'Thiếu hoặc sai lệch báo cáo',
      status: isAtsPass ? 'PASS' : 'FAIL',
      evaluation_vn: isAtsPass ? 'Trùng khớp 100% hồ sơ kỹ thuật OPR/BOD và thiết bị đã qua thử nghiệm nghiệm thu NETA ATS.' : 'Thiếu báo cáo nghiệm thu chấp nhận hoặc sai lệch hồ sơ.',
      evaluation_en: isAtsPass ? '100% compliance with approved submittals and NETA ATS reports verified.' : 'Incomplete test records or drawing discrepancy.'
    },
    {
      item_code: '7.3.A.2',
      item_name_vn: 'Điện Trở Cách Điện Thanh Cái Trạm 230kV / EHV Bus',
      item_name_en: 'EHV Busbar Insulation Resistance (10,000V DC)',
      neta_standard_ref: 'ANSI/NETA ATS Table 100.1 (Min ≥ 10,000 MΩ)',
      measured_value: `A: ${busIr.phase_a_to_ground} MΩ | B: ${busIr.phase_b_to_ground} MΩ | C: ${busIr.phase_c_to_ground} MΩ`,
      status: isBusIrPass ? 'PASS' : 'FAIL',
      evaluation_vn: isBusIrPass ? `Cách điện thanh cái 230kV đạt mức rất cao (${minBusIr.toLocaleString()} MΩ), vượt ngưỡng 10,000 MΩ Bảng Table 100.1.` : `Cách điện thanh cái thấp (${minBusIr} MΩ < 10,000 MΩ), nguy cơ rò điện.`,
      evaluation_en: isBusIrPass ? `Insulation resistance (${minBusIr.toLocaleString()} MΩ) exceeds Table 100.1 threshold.` : `Insulation resistance below Table 100.1 minimum limit.`
    },
    {
      item_code: '7.3.A.5',
      item_name_vn: 'Thử Nghiệm Cáp Ngầm Cao Thế AC & Phóng Điện Cục Bộ Offline',
      item_name_en: 'HV Cable AC Withstand & Offline PD (Table 100.6.6)',
      neta_standard_ref: 'ANSI/NETA ATS Table 100.6.5, 100.6.6 & IEEE 400.3',
      measured_value: isCablePass ? 'Chịu AC tần số thi công đạt, PD offline an toàn' : 'Không đạt điện áp chịu đựng hoặc PD cao',
      status: isCablePass ? 'PASS' : 'FAIL',
      evaluation_vn: isCablePass ? 'Cáp ngầm cao thế chịu đựng thành công điện áp AC và mức PD offline nằm trong giới hạn.' : 'Cáp phóng điện trong thử nghiệm hoặc góc tổn hao Tan Delta quá cao.',
      evaluation_en: isCablePass ? 'AC withstand test and offline PD within permissible NETA Table 100.6.6 limits.' : 'Dielectric breakdown or excessive offline PD detected.'
    },
    {
      item_code: '7.3.A.35',
      item_name_vn: 'Chất Lượng & Hàm Lượng Ẩm Khí SF6 Trạm GIS 230kV',
      item_name_en: 'SF6 Gas Quality & Moisture Analysis (Table 100.19)',
      neta_standard_ref: 'ANSI/NETA ATS Sec 7.5.4 & Table 100.19',
      measured_value: `Độ tinh khiết: ${sf6.sf6_purity_percent}% | Ẩm: ${sf6.moisture_content_ppmv} ppmv (Điểm sương ${sf6.dew_point_celsius}°C) | SO2: ${sf6.so2_content_ppmv} ppmv`,
      status: isSf6Pass ? 'PASS' : 'FAIL',
      evaluation_vn: isSf6Pass ? 'Khí SF6 đạt độ tinh khiết cao (≥ 98%), độ ẩm thấp (≤ 99 ppmv) và nồng độ SO2 an toàn.' : `NGUY HẠI: Độ ẩm khí SF6 (${sf6.moisture_content_ppmv} ppmv) vượt ngưỡng Bảng Table 100.19 (≤ 99 ppmv, điểm sương ≤ -36°C).`,
      evaluation_en: isSf6Pass ? 'SF6 gas parameters strictly comply with Table 100.19 limits.' : `CRITICAL: SF6 moisture (${sf6.moisture_content_ppmv} ppmv) violates Table 100.19 limit (<= 99 ppmv / dew point <= -36°C).`
    },
    {
      item_code: '7.3.A.36',
      item_name_vn: 'Phân Tích Dầu Cách Điện & Khí Hòa Tan DGA MBA 230kV',
      item_name_en: 'Transformer Oil Fluid & DGA Combustible Gas Analysis',
      neta_standard_ref: 'ANSI/NETA ATS Table 100.4.1 & IEEE C57.104',
      measured_value: `BDV: ${oil.breakdown_voltage_bdv_kV} kV | Ẩm dầu: ${oil.moisture_content_ppm} ppm | C2H2: ${oil.dga_combustible_gases_ppm.c2h2} ppm`,
      status: isOilDgaPass ? 'PASS' : 'FAIL',
      evaluation_vn: isOilDgaPass ? 'Điện áp đánh thủng BDV ≥ 50 kV, độ ẩm dầu thấp và khí cháy DGA nằm trong Condition 1 an toàn.' : 'Dầu biến áp cách điện kém hoặc xuất hiện khí cháy C2H2 (phóng điện hồ quang nội bộ).',
      evaluation_en: isOilDgaPass ? 'Dielectric breakdown and DGA combustible gases comply with IEEE C57.104 Condition 1.' : 'Dielectric failure or arcing detected in DGA.'
    },
    {
      item_code: '7.3.A.32',
      item_name_vn: 'Rơ-le Bảo Vệ 87L/21 & Phối Hợp Kích Từ NERC PRC',
      item_name_en: 'Protection Relays 87L/21 & NERC PRC Compliance',
      neta_standard_ref: 'NERC PRC-002-4 / PRC-005-6 / PRC-019-2 / PRC-024-3',
      measured_value: isRelayPass ? 'Thử bơm dòng sơ cấp 87L/21 & nhật ký NERC Đạt 100%' : 'Lỗi phối hợp bảo vệ hoặc thiếu dữ liệu NERC',
      status: isRelayPass ? 'PASS' : 'FAIL',
      evaluation_vn: isRelayPass ? 'Rơ-le so lệch đường dây 87L, bảo vệ khoảng cách 21 và thiết bị ghi sự cố DME/DDR đạt chuẩn NERC.' : 'Sai lệch góc pha nhị thứ hoặc cấu hình bảo vệ chưa tuân thủ tiêu chuẩn NERC.',
      evaluation_en: isRelayPass ? 'Line differential 87L, distance 21, and NERC compliance verified.' : 'Protection coordination defect or NERC record non-compliance.'
    },
    {
      item_code: '7.3.A.7',
      item_name_vn: 'Điện Trở Lưới Tiếp Địa Trạm & Lực Siết Bu-lông',
      item_name_en: 'Substation Grounding Grid (IEEE 81) & Bolted Torque (Table 100.12)',
      neta_standard_ref: 'IEEE 81 & NETA ATS/MTS Table 100.12',
      measured_value: `Điện trở lưới đất: ${groundRes.toFixed(2)} Ω | Lực siết bu-lông: ${isTorquePass ? 'Đạt Bảng 100.12' : 'Chưa đạt'}`,
      status: isGroundPass && isTorquePass ? 'PASS' : 'FAIL',
      evaluation_vn: isGroundPass && isTorquePass ? `Điện trở lưới tiếp địa trạm đạt ${groundRes} Ω (≤ 1.0 Ω) và lực siết bu-lông chuẩn xác.` : `Điện trở tiếp địa trạm quá cao (${groundRes} Ω > 1.0 Ω) hoặc bu-lông lỏng lẻo.`,
      evaluation_en: isGroundPass && isTorquePass ? `Grounding grid resistance ${groundRes} Ω meets IEEE 81 and torque is verified.` : `Ground grid resistance ${groundRes} Ω exceeds 1.0 Ohm limit.`
    },
    {
      item_code: '7.3.A.15',
      item_name_vn: 'Khóa Liên Động An Toàn & Hệ Thống Giảm Hồ Quang',
      item_name_en: 'Mechanical Key Interlocks & Arc Flash Mitigation',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Mục 7.3.A.15 & 7.3.A.16',
      measured_value: isInterlocksPass ? 'Liên động hoàn hảo, chìa khóa thu hồi đủ' : 'Lỗi khóa liên động',
      status: isInterlocksPass ? 'PASS' : 'FAIL',
      evaluation_vn: isInterlocksPass ? 'Khóa liên động giữa Dao cách ly và Máy cắt hoạt động chính xác 100%, ngăn ngừa thao tác nhầm.' : 'Lỗi liên động an toàn, nguy cơ đóng dao cách ly khi máy cắt đang mang tải.',
      evaluation_en: isInterlocksPass ? 'Mechanical and electrical key interlocks fully prevent incorrect switching operations.' : 'Interlock malfunction threatens hazardous switching.'
    },
    {
      item_code: '7.3.A.30',
      item_name_vn: 'Dàn Bình Ắc-quy DC 220V Kép & Phê Duyệt Kế Hoạch Đóng Điện',
      item_name_en: 'Dual DC Battery Systems & Approved Written Energization Plan',
      neta_standard_ref: 'ANSI/NETA ECS 7.3.A.30 & 7.3.A.45 / IEEE 450',
      measured_value: isDcBatteryPass && isPlanApproved ? 'Nguồn DC kép ổn định 100% & Kế hoạch đã ký duyệt' : 'Nguồn DC lỗi hoặc chưa duyệt kế hoạch',
      status: isDcBatteryPass && isPlanApproved ? 'PASS' : 'FAIL',
      evaluation_vn: isDcBatteryPass && isPlanApproved ? 'Nguồn điều khiển DC dự phòng tin cậy và kế hoạch đóng điện từng bước đã được ký duyệt.' : 'Nguồn ắc-quy DC suy thoái hoặc chưa có kế hoạch đóng điện bằng văn bản.',
      evaluation_en: isDcBatteryPass && isPlanApproved ? 'Dual battery systems stable and formal written plan approved.' : 'DC power defect or unapproved energization plan.'
    },
    {
      item_code: '7.3.B.1',
      item_name_vn: 'Xác Nhận Tháo Gỡ 100% Tiếp Địa Di Động Tạm Thời',
      item_name_en: 'Verification of 100% Temporary Protective Ground Removal',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Mục 7.3.B.1 & 7.3.B.3',
      measured_value: isGroundsRemoved ? 'Đã tháo bỏ & đối chiếu 100% dây tiếp địa' : 'NGUY HẠI: Còn sót tiếp địa di động trên thanh cái!',
      status: isGroundsRemoved ? 'PASS' : 'FAIL',
      evaluation_vn: isGroundsRemoved ? 'Toàn bộ các bộ tiếp địa di động tạm thời đã được tháo gỡ và ký biên bản đối soát an toàn.' : 'NGUY HẠI CHẾT NGƯỜI: Còn sót tiếp địa di động trên thanh cái, sẽ gây ngắn mạch pha-đất tức thời!',
      evaluation_en: isGroundsRemoved ? 'Zero temporary protective grounds remaining on HV busbars.' : 'FATAL HAZARD: Temporary grounds remaining will cause dead short-circuit on energization!'
    },
    {
      item_code: '7.3.B.8',
      item_name_vn: 'Kiểm Tra Đồng Pha Hai Nguồn 230kV Tại Máy Cắt Liên Lạc (Bus Tie)',
      item_name_en: 'Dual-Source 230kV Phasing Verification (Bus Tie Breaker)',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Mục 7.3.B.8 & 7.3.B.9',
      measured_value: `ΔV tương ứng: A-A = ${phasing.phase_a1_a2_voltage_volts}V | B-B = ${phasing.phase_b1_b2_voltage_volts}V | C-C = ${phasing.phase_c1_c2_voltage_volts}V`,
      status: isPhasingPass ? 'PASS' : 'FAIL',
      evaluation_vn: isPhasingPass ? `Độ lệch điện áp giữa các pha tương ứng rất nhỏ (ΔV ≤ ${maxDeltaV}V), hai nguồn 230kV hoàn toàn đồng pha.` : `NGUY HẠI: Lệch áp đồng pha lớn (ΔV = ${maxDeltaV}V), nguy cơ đóng nhầm hai nguồn ngược pha.`,
      evaluation_en: isPhasingPass ? `In-phase condition confirmed across 230kV dual sources (Delta-V <= ${maxDeltaV}V).` : `Out-of-phase condition detected across bus tie breaker.`
    },
    {
      item_code: '7.3.B.7',
      item_name_vn: 'Kiểm Tra Đồ Thị Véc-tơ Góc Pha Rơ-le & Tín Hiệu SCADA',
      item_name_en: 'Relay Vector Inputs & SCADA Phasing Validation',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Mục 7.3.B.7 & 7.3.B.10',
      measured_value: isVectorPass ? 'Véc-tơ dòng/áp chuẩn xác 100%, truyền thông SCADA thông suốt' : 'Lỗi góc pha véc-tơ',
      status: isVectorPass ? 'PASS' : 'FAIL',
      evaluation_vn: isVectorPass ? 'Đồ thị véc-tơ nhị thứ trên rơ-le trùng khớp hoàn toàn với chiều dòng tải thực tế.' : 'Đấu ngược cực tính biến dòng CT/VT hoặc sai lệch góc pha.',
      evaluation_en: isVectorPass ? 'Secondary current and voltage vector angles match actual load flow.' : 'CT/VT polarity error or vector angle mismatch.'
    },
    {
      item_code: '7.3.C.3',
      item_name_vn: 'Khảo Sát Nhiệt Hồng Ngoại Mang Tải (NETA MTS Table 100.18)',
      item_name_en: 'Infrared Thermographic Survey Under Load',
      neta_standard_ref: 'ANSI/NETA ECS Section 7.3.C.3 & Table 100.18',
      measured_value: isThermalPass ? 'Không có điểm quá nhiệt bất thường (ΔT ≤ 3.0°C)' : 'Phát hiện điểm nóng bất thường',
      status: isThermalPass ? 'PASS' : 'INVESTIGATE',
      evaluation_vn: isThermalPass ? 'Khảo sát nhiệt dưới tải mang dòng ổn định, các kẹp cực và đầu cáp tản nhiệt đồng đều.' : 'Phát hiện điểm nóng cục bộ trên tiếp xúc cao thế cần siết lại.',
      evaluation_en: isThermalPass ? 'Thermal profile balanced across all 230kV bushings and terminals.' : 'Thermal hotspot detected under load.'
    },
    {
      item_code: '7.3.C.4',
      item_name_vn: 'Đo Phóng Điện Cục Bộ Online Trạm GIS (UHF & Acoustic)',
      item_name_en: 'Online Partial Discharge Survey (NETA MTS Table 100.23)',
      neta_standard_ref: 'ANSI/NETA ECS Section 7.3.C.4 & Table 100.23.3',
      measured_value: `Cảm biến UHF: ${pd.uhf_max_dBmV} dBmV | Cảm biến Siêu âm: ${pd.acoustic_max_dB} dB`,
      status: isOnlinePdPass ? 'PASS' : 'FAIL',
      evaluation_vn: isOnlinePdPass ? `Tín hiệu PD trên trạm GIS ở mức an toàn (< 6.0 dBmV), không có hoạt động phóng điện bên trong khoang khí.` : `Phát hiện tín hiệu phóng điện cục bộ bất thường trong khoang GIS.`,
      evaluation_en: isOnlinePdPass ? `UHF and Acoustic online PD levels well below Warning threshold.` : `Abnormal partial discharge activity detected in GIS chamber.`
    },
    {
      item_code: '7.3.C.1',
      item_name_vn: 'Lấy Mẫu Khí/Dầu Sau Đóng Điện & Báo Cáo Hoàn Công',
      item_name_en: 'Post-Energization Gas/Fluid Re-test & Commissioning Dossier',
      neta_standard_ref: 'ANSI/NETA ECS-2024 Mục 7.3.C.1 & Section 5.4',
      measured_value: isPostReportPass ? 'Đã lấy mẫu sau 24h & Hoàn thiện hồ sơ bàn giao' : 'Chưa hoàn thành báo cáo',
      status: isPostReportPass ? 'PASS' : 'INVESTIGATE',
      evaluation_vn: isPostReportPass ? 'Lấy mẫu kiểm tra lại sau 24h mang tải và phát hành biên bản nghiệm thu chạy thử hoàn chỉnh.' : 'Chưa hoàn tất lấy mẫu kiểm tra sau đóng điện hoặc chưa nộp báo cáo hoàn công.',
      evaluation_en: isPostReportPass ? 'Post-energization re-sampling verified and commissioning dossier completed.' : 'Incomplete post-energization sampling or documentation.'
    }
  ];

  // Determine overall severity
  const hasCriticalHazard =
    !isGroundsRemoved ||
    !isPhasingPass ||
    !isSf6MoisturePass ||
    !isDgaC2h2Pass ||
    !isBusIrPass ||
    !isCablePass;

  const hasCaution =
    !isAtsPass ||
    !isSf6PurityPass ||
    !isSf6So2Pass ||
    !isOilBdvPass ||
    !isOilMoisturePass ||
    !isOilPfPass ||
    !isRelayPass ||
    !isGroundPass ||
    !isTorquePass ||
    !isInterlocksPass ||
    !isDcBatteryPass ||
    !isPlanApproved ||
    !isVectorPass ||
    !isThermalPass ||
    !isOnlinePdPass ||
    !isPostReportPass;

  let overall_status: 'PASS' | 'CAUTION' | 'CRITICAL' = 'PASS';
  let status_title_en = 'HV/EHV POWER SYSTEM COMMISSIONED / READY FOR COMMERCIAL GRID ENERGIZATION';
  let status_title_vn = 'HỆ THỐNG CAO THẾ & SIÊU CAO THẾ ĐÃ CHẠY THỬ ĐẠT CHUẨN / SẴN SÀNG HÒA LƯỚI QUỐC GIA';
  let summary_badge = 'NETA ECS 7.3 PASSED';
  let summary_color = 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40';

  if (hasCriticalHazard) {
    overall_status = 'CRITICAL';
    if (!isGroundsRemoved) {
      status_title_en = 'FATAL HAZARD / TEMPORARY PROTECTIVE GROUNDS REMAINING ON HV BUS';
      status_title_vn = 'NGUY HIỂM CHẾT NGƯỜI - CÒN SÓT TIẾP ĐỊA DI ĐỘNG TRÊN THANH CÁI CAO THẾ';
    } else if (!isPhasingPass) {
      status_title_en = 'CRITICAL ENERGIZATION HAZARD / DUAL-SOURCE 230KV PHASING MISMATCH';
      status_title_vn = 'NGUY HẠI ĐÓNG ĐIỆN - LỆCH PHA GIỮA HAI NGUỒN CAO THẾ 230KV TẠI BUS TIE';
    } else if (!isSf6MoisturePass) {
      status_title_en = 'CRITICAL ENERGIZATION HAZARD / SF6 GAS MOISTURE EXCEEDS SPECIFICATION';
      status_title_vn = 'NGUY HẠI ĐÓNG ĐIỆN - HÀM LƯỢNG ẨM KHÍ SF6 VƯỢT NGƯỠNG CHO PHÉP TẠI TRẠM GIS 230KV';
    } else if (!isDgaC2h2Pass) {
      status_title_en = 'CRITICAL ENERGIZATION HAZARD / TRANSFORMER OIL DGA ACETYLENE ARCING HAZARD';
      status_title_vn = 'NGUY HẠI ĐÓNG ĐIỆN - PHÁT HIỆN KHÍ ACETYLENE (C2H2) PHÓNG ĐIỆN TRONG MBA 230KV';
    } else {
      status_title_en = 'CRITICAL ENERGIZATION HAZARD / SEVERE DIELECTRIC BREAKDOWN DETECTED';
      status_title_vn = 'NGUY HẠI ĐÓNG ĐIỆN - CÁCH ĐIỆN CAO THẾ SUY GIẢM NGHIÊM TRỌNG';
    }
    summary_badge = 'CRITICAL HAZARD';
    summary_color = 'bg-rose-500/20 text-rose-300 border-rose-500/40';
  } else if (hasCaution) {
    overall_status = 'CAUTION';
    status_title_en = 'HV/EHV COMMISSIONING CAUTION / DEFICIENCIES REQUIRE RECTIFICATION';
    status_title_vn = 'CẢNH BÁO CHẠY THỬ CAO THẾ / CÁC KHIẾM KHUYẾT CẦN XỬ LÝ TRƯỚC KHI HÒA LƯỚI';
    summary_badge = 'PENDING RECTIFICATION';
    summary_color = 'bg-amber-500/20 text-amber-300 border-amber-500/40';
  }

  // Construct Technical Analysis
  const technical_analysis = {
    hazard_identified: hasCriticalHazard,
    title_en: hasCriticalHazard
      ? !isSf6MoisturePass
        ? 'High SF6 Gas Moisture Hazard on 230kV GIS Bay-01 (NETA Table 100.19)'
        : !isGroundsRemoved
        ? 'Fatal Grounding Hazard: Unremoved Temporary Grounds on EHV Bus'
        : 'Critical High-Voltage Electrical Hazard Under Commissioning'
      : hasCaution
      ? 'Grounding Resistance or Relay Settings Observation Pending Clearance'
      : 'EHV Dual-Source Phasing Integrity & NETA ECS 7.3 Compliance Validated',
    title_vn: hasCriticalHazard
      ? !isSf6MoisturePass
        ? 'Sự Cố Hàm Lượng Ẩm Khí SF6 Vượt Ngưỡng tại Bay GIS 230kV (NETA Table 100.19)'
        : !isGroundsRemoved
        ? 'Nguy Cơ Ngắn Mạch Chết Người: Chưa Tháo Tiếp Địa Di Động Trên Thanh Cái'
        : 'Nguy Cơ Đóng Điện Nghiêm Trọng Trong Hệ Thống Cao Thế & Siêu Cao Thế'
      : hasCaution
      ? 'Cảnh Báo Điện Trở Lưới Đất Hoặc Cài Đặt Rơ-le Cần Hoàn Thiện'
      : 'Hệ Thống Trạm 230kV Đạt Chuẩn Kỹ Thuật Tuyệt Đối & Sẵn Sàng Hòa Lưới',
    description_en: hasCriticalHazard
      ? !isSf6MoisturePass
        ? `Pre-energization testing under ANSI/NETA ATS Section 7.5.4 and NETA Table 100.19 revealed a critical gas quality violation on 230kV GIS Bay-01. The measured moisture content reached ${sf6.moisture_content_ppmv} ppmv with a dew point of ${sf6.dew_point_celsius}°C, which significantly fails the NETA Table 100.19 maximum limit of <= 99 ppmv (dew point <= -36°C). High moisture in 230kV SF6 insulation leads to hydrolysis under electrical arc switching, producing highly corrosive hydrofluoric acid (HF) and toxic SOF2/SO2F2 byproducts. This dramatically degrades internal spacer insulation strength, risking catastrophic internal arc flash breakdown upon energization.`
        : !isGroundsRemoved
        ? `Critical safety breach: Temporary protective grounds were left installed on the 230kV busbars. Energizing an EHV bus with attached protective grounds will cause a direct bolted three-phase or phase-to-ground short circuit, unleashing hundreds of kiloamperes of arc-flash energy and causing catastrophic switchgear destruction.`
        : `Severe dielectric failure detected in HV equipment, violating ANSI/NETA ECS-2024 Section 7.3 and IEEE standards. System must remain de-energized.`
      : hasCaution
      ? `HV/EHV commissioning parameters are largely compliant, but cautionary items (such as substation ground grid resistance ${groundRes} Ω or relay configuration) require engineering rectification.`
      : `All pre-energization, dual-source phasing verification, 100% protective ground removal, relay coordination (NERC PRC), and post-energization online PD criteria strictly comply with ANSI/NETA ECS-2024 Section 7.3 and NETA Tables 100.1, 100.4, 100.6, 100.12, 100.18, 100.19, and 100.23. The 230kV substation is fully certified for commercial grid operation.`,
    description_vn: hasCriticalHazard
      ? !isSf6MoisturePass
        ? `Kiểm tra tiền đóng điện theo Mục ANSI/NETA ATS 7.5.4 và Bảng NETA Table 100.19 phát hiện vi phạm tiêu chuẩn chất lượng khí SF6 nghiêm trọng tại Bay GIS 230kV Bay-01. Hàm lượng ẩm đo được lên tới ${sf6.moisture_content_ppmv} ppmv tương ứng với điểm sương ${sf6.dew_point_celsius}°C, vượt xa giới hạn tối đa <= 99 ppmv (điểm sương <= -36°C) của Bảng NETA Table 100.19. Độ ẩm cao trong môi trường khí SF6 230kV khi xảy ra thao tác dập hồ quang sẽ bị thủy phân tạo thành axit HF ăn mòn và các sản phẩm độc hại SOF2/SO2F2. Điều này làm suy giảm nghiêm trọng độ bền cách điện của sứ đỡ dạng đĩa (spacers) bên trong khoang GIS, nguy cơ trực tiếp gây sự cố phóng điện hồ quang nội bộ thảm khốc khi đóng điện.`
        : !isGroundsRemoved
        ? `Vi phạm an toàn đặc biệt nghiêm trọng: Bộ tiếp địa di động tạm thời vẫn còn trên thanh cái 230kV. Việc đóng điện vào thanh cái đang bị ngắn mạch tiếp địa sẽ giải phóng năng lượng hồ quang hàng trăm kA, phá hủy hoàn toàn trạm điện và gây nguy hiểm tính mạng!`
        : `Phát hiện sự cố cách điện nghiêm trọng trong hệ thống cao thế vi phạm tiêu chuẩn ANSI/NETA ECS-2024 Mục 7.3 và IEEE. Bắt buộc duy trì trạng thái cô lập toàn trạm.`
      : hasCaution
      ? `Các hạng mục chạy thử cơ bản đạt yêu cầu, tuy nhiên còn tồn tại điểm lưu ý (như điện trở lưới đất ${groundRes} Ω hoặc cài đặt rơ-le) cần được xử lý trước khi đóng điện chính thức.`
      : `Toàn bộ các phép đo tiền đóng điện, kiểm tra đồng pha hai nguồn 230kV, xác nhận tháo 100% tiếp địa di động, phối hợp rơ-le số NERC PRC và đo PD trực tuyến đều hoàn toàn tuân thủ tiêu chuẩn ANSI/NETA ECS-2024 Mục 7.3 và các Bảng NETA Table 100.1, 100.4, 100.6, 100.12, 100.18, 100.19, 100.23. Trạm biến áp 230kV đủ điều kiện bàn giao hòa lưới thương mại.`,
    min_insulation_mohm: minBusIr,
    sf6_moisture_ppmv: sf6.moisture_content_ppmv,
    sf6_dew_point_c: sf6.dew_point_celsius,
    transformer_bdv_kv: oil.breakdown_voltage_bdv_kV,
    ground_resistance_ohms: groundRes,
    phasing_max_delta_v: maxDeltaV
  };

  // Construct Actionable Recommendations
  const actionable_recommendations = {
    halt_energization: hasCriticalHazard,
    steps: hasCriticalHazard
      ? [
          {
            step_number: 1,
            title_en: 'HOLD ENERGIZATION OF 230KV GIS BAY-01 IMMEDIATELY',
            title_vn: 'HOÃN TIẾN TRÌNH ĐÓNG ĐIỆN BAY GIS 230KV-01 NGAY LẬP TỨC',
            details_en: 'Do NOT close the 230kV main circuit breaker or line disconnectors. Halt the energization sequence until SF6 gas evacuation, vacuum drying, and re-filling procedures are completed.',
            details_vn: 'KHÔNG đóng máy cắt tổng 230kV hoặc dao cách ly đường dây. Tạm dừng tiến trình đóng điện cho đến khi quy trình hút chân không sấy khô và nạp lại khí SF6 đạt tiêu chuẩn.'
          },
          {
            step_number: 2,
            title_en: 'SF6 GAS EVACUATION, VACUUM DRYING & RE-GASSING PROTOCOL',
            title_vn: 'QUY TRÌNH HÚT CHÂN KHÔNG SẤY KHÔ & NẠP LẠI KHÍ SF6 MỚI',
            details_en: 'Recover contaminated SF6 gas using an SF6 gas service cart equipped with molecular sieve filters. Evacuate 230kV GIS Bay-01 to a deep vacuum of <= 1.0 mbar and maintain vacuum for at least 4 hours to remove internal absorbed moisture. Refill with fresh, certified virgin SF6 gas to the rated pressure of 6.0 bar.',
            details_vn: 'Thu hồi lượng khí SF6 bị ẩm bằng xe sút khí SF6 chuyên dụng có trang bị bộ lọc rây phân tử. Thực hiện hút chân không khoang GIS 230kV Bay-01 xuống áp suất sâu <= 1.0 mbar và duy trì chân không tối thiểu 4 giờ để loại bỏ hoàn toàn độ ẩm đọng trên bề mặt sứ đĩa. Nạp lại bằng khí SF6 mới đạt chứng nhận chất lượng đến áp suất định mức 6.0 bar.'
          },
          {
            step_number: 3,
            title_en: 'RE-TESTING & FINAL COMMISSIONING CLEARANCE',
            title_vn: 'ĐO KIỂM TRA LẠI NGHIỆM THU & PHÁT HÀNH CHỨNG NHẬN',
            details_en: 'Re-measure SF6 gas purity (must be >= 98%), moisture content (must be <= 99 ppmv / dew point <= -36°C), and SO2 (must be <= 1.0 ppmv). Upon passing all NETA Table 100.19 criteria, issue the final Electrical Commissioning Certificate for 230KV-GIS-BAY-01 and approve commercial energization under Section 7.3.B.',
            details_vn: 'Đo kiểm tra lại độ tinh khiết khí SF6 (BẮT BUỘC >= 98%), hàm lượng ẩm (BẮT BUỘC <= 99 ppmv / điểm sương <= -36°C) và nồng độ SO2 (BẮT BUỘC <= 1.0 ppmv). Sau khi tất cả các thông số đạt chuẩn Bảng NETA Table 100.19, phát hành Giấy chứng nhận Nghiệm thu Chạy thử chính thức cho 230KV-GIS-BAY-01 và phê duyệt đóng điện vận hành thương mại theo Mục 7.3.B.'
          }
        ]
      : hasCaution
      ? [
          {
            step_number: 1,
            title_en: 'IMPROVE SUBSTATION GROUNDING GRID CONDUCTIVITY',
            title_vn: 'CẢI THIỆN ĐIỆN TRỞ LƯỚI TIẾP ĐỊA TRẠM CAO THẾ',
            details_en: 'Drive additional deep ground rods or treat soil around grounding perimeter to reduce grid resistance to <= 0.5 Ohm for EHV substations per IEEE 81.',
            details_vn: 'Đóng thêm cọc tiếp địa sâu hoặc xử lý hóa chất giảm điện trở đất xung quanh chu vi trạm để đưa điện trở lưới đất về <= 0.5 Ohm theo chuẩn trạm EHV IEEE 81.'
          },
          {
            step_number: 2,
            title_en: 'RE-TORQUE DETECTED THERMAL ANOMALIES & COMPLETE NERC LOGS',
            title_vn: 'SIẾT LẠI BU-LÔNG CÓ ĐIỂM NÓNG & HOÀN THIỆN HỒ SƠ NERC',
            details_en: 'Re-torque bolted terminals showing slight temperature rise per Table 100.12 and archive DME/DDR fault recorder settings in compliance with NERC PRC-002-4.',
            details_vn: 'Siết lại bu-lông các đầu cốt có dấu hiệu ấm lên bằng cờ-lê lực kiểm định và lưu trữ bản ghi sự cố DME/DDR theo quy định NERC PRC-002-4.'
          }
        ]
      : [
          {
            step_number: 1,
            title_en: 'ISSUE FORMAL 230KV EHV COMMISSIONING ACCEPTANCE CERTIFICATE',
            title_vn: 'PHÁT HÀNH CHỨNG NHẬN NGHIỆM THU ĐÓNG ĐIỆN CAO THẾ 230KV',
            details_en: 'All criteria under ANSI/NETA ECS-2024 Section 7.3, NETA Tables, and NERC PRC standards are satisfied. Approved for permanent commercial grid energization.',
            details_vn: 'Toàn bộ các hạng mục kiểm tra theo ANSI/NETA ECS-2024 Mục 7.3, các Bảng NETA và tiêu chuẩn NERC PRC đã đạt chuẩn 100%. Phê duyệt đóng điện hòa lưới vĩnh viễn.'
          },
          {
            step_number: 2,
            title_en: 'SCHEDULE ANNUAL NETA MTS ROUTINE PREVENTIVE MAINTENANCE',
            title_vn: 'LẬP KẾ HOẠCH BẢO TRÌ ĐỊNH KỲ THEO TIÊU CHUẨN NETA MTS',
            details_en: 'Incorporate 230kV GIS, cables, and 250MVA transformer into CMMS for annual infrared scans (Table 100.18), quarterly DGA sampling (IEEE C57.104), and continuous online PD monitoring (Table 100.23).',
            details_vn: 'Đưa trạm GIS 230kV, cáp lực và MBA 250MVA vào phần mềm CMMS để lập lịch chụp nhiệt hồng ngoại hàng năm (Bảng 100.18), lấy mẫu DGA định kỳ (IEEE C57.104) và giám sát PD online liên tục.'
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

export async function generateEcsHvAiAnalysis(
  payload: NetaEcsHvInputPayload
): Promise<NetaEcsHvEvaluationResult> {
  const deterministicResult = performDeterministicEcsHvAnalysis(payload);

  try {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      console.warn('GEMINI_API_KEY is not defined. Returning deterministic HV ECS analysis.');
      return deterministicResult;
    }

    const ai = new GoogleGenAI({ apiKey });
    const prompt = `Please evaluate the following High-Voltage and Extra-High Voltage System (> 100 kV and < 1,000 kV) Commissioning test data strictly based on ANSI/NETA ECS-2024 (Section 7.3), ANSI/NETA ATS/MTS standards, NETA Tables (100.1, 100.2, 100.3, 100.4, 100.6, 100.12, 100.18, 100.19, 100.23), IEEE 81, IEEE C57.104, and NERC PRC standards.

INPUT DATA PAYLOAD:
${JSON.stringify(payload, null, 2)}

Provide an authoritative bilingual (EN/VN) commissioning evaluation report conforming strictly to the system instructions. Do NOT include citation numbers or bibliography links.`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: prompt,
      config: {
        systemInstruction: ECS_HV_SYSTEM_INSTRUCTION,
        temperature: 0.1
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
    console.error('Error during Gemini HV ECS AI analysis:', error);
    return deterministicResult;
  }
}
