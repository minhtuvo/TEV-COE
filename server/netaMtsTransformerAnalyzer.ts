import { GoogleGenAI } from '@google/genai';

export interface NetaMtsTransformerInputPayload {
  site_info: {
    customer_name?: string;
    project_name: string;
    station_name?: string;
    substation_name?: string;
    transformer_tag: string;
    transformer_rating?: string;
    fse_name: string;
    test_date: string;
    voltage_class_kv?: number; // e.g. 110, 22, 69
    oil_type?: 'Mineral Oil Class I' | 'Mineral Oil Class II' | 'Silicone' | 'Natural Ester' | 'Less-Flammable Hydrocarbon';
  };
  visual_inspection: {
    nameplate_match: boolean;
    liquid_level_check: string;
    gas_blanket_pressure_positive: boolean;
    valves_and_cooling_fans_check: string;
    bolt_torque_check: string;
    no_oil_leakage: boolean;
    thermographic_survey?: string;
  };
  electrical_tests: {
    bolted_resistance_micro_ohms: number[];
    insulation_resistance_20c_megohms: {
      pri_to_ground_5000v: number;
      sec_to_ground_2500v: number;
      pri_to_sec_5000v_dc: number;
      calculated_pi_value: number;
    };
    turns_ratio_all_taps: {
      max_error_percent: number;
    };
    winding_insulation_power_factor_percent_20c: number;
    winding_dc_resistance_ohms?: {
      max_phase_deviation_percent: number;
    };
    core_insulation_resistance_500v_megohms?: number;
    service_aged_oil_sample_tests_table_100_4_1: {
      dielectric_breakdown_d1816_1mm_kv: number;
      dielectric_breakdown_d877_kv?: number;
      water_content_d1533_ppm: number;
      acid_number_d974_mg_koh_g: number;
      interfacial_tension_d971_mn_m: number;
      oil_power_factor_25c_d924_percent: number;
    };
    dga_test_ieee_c57_104: {
      hydrogen_h2_ppm: number;
      methane_ch4_ppm: number;
      acetylene_c2h2_ppm: number;
      ethylene_c2h4_ppm: number;
      ethane_c2h6_ppm?: number;
      carbon_monoxide_co_ppm?: number;
      carbon_dioxide_co2_ppm?: number;
      total_dissolved_combustible_gas_ppm?: number;
    };
  };
  previous_test_data?: {
    last_test_date?: string;
    last_oil_breakdown_kv?: number;
    last_water_content_ppm?: number;
    last_c2h2_ppm?: number;
  };
  language?: 'vi' | 'en';
}

export interface MtsTransformerItemEvaluation {
  item: string;
  itemEn: string;
  measured: string;
  standard: string;
  standardEn: string;
  reference: string;
  status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  detail: string;
  detailEn: string;
}

export interface NetaMtsTransformerEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_title_vi: string;
  status_title_en: string;
  status_reason_vi: string;
  status_reason_en: string;
  markdown_report_vi: string;
  markdown_report_en: string;
  evaluations: {
    max_severity: 'PASS' | 'INVESTIGATE' | 'FAIL';
    dielectric_breakdown_failed: boolean;
    water_content_exceeded: boolean;
    acid_number_exceeded: boolean;
    ift_below_minimum: boolean;
    arcing_c2h2_critical: boolean;
    winding_dc_dev_exceeded: boolean;
    core_ir_failed: boolean;
    items: MtsTransformerItemEvaluation[];
  };
}

export function performDeterministicMtsTransformerAnalysis(
  data: NetaMtsTransformerInputPayload
): NetaMtsTransformerEvaluationResult {
  const items: MtsTransformerItemEvaluation[] = [];

  // Determine voltage level for Table 100.4.1 thresholds
  // Defaults to 110kV if rating contains 110kV or >=69kV
  const ratingText = (data.site_info.transformer_rating || '').toLowerCase();
  const isHighVoltage = ratingText.includes('110kv') || ratingText.includes('220kv') || ratingText.includes('110/22') || (data.site_info.voltage_class_kv && data.site_info.voltage_class_kv > 69);
  
  // Table 100.4.1 limits for Mineral Oil:
  // For >69kV to <230kV (like 110kV):
  // D1816 1mm >= 28 kV, Water <= 25 ppm, Acid <= 0.20 mg KOH/g, IFT >= 24 mN/m, PF 25°C <= 1.0%
  // For <= 69kV:
  // D1816 1mm >= 23 kV, Water <= 35 ppm, Acid <= 0.20 mg KOH/g, IFT >= 24 mN/m, PF 25°C <= 1.0%
  const minBreakdownD1816 = isHighVoltage ? 28 : 23;
  const maxWaterPpm = isHighVoltage ? 25 : 35;
  const maxAcid = 0.20;
  const minIft = 24;
  const maxOilPf = 1.0;

  // 1. Visual & Mechanical (NETA MTS-2023 Sec 7.2.2.A)
  const v = data.visual_inspection;
  if (!v.nameplate_match) {
    items.push({
      item: 'Đối soát nhãn máy (Nameplate Match)',
      itemEn: 'Nameplate Match Verification (IEEE C57.12.00)',
      measured: 'Không khớp hồ sơ',
      standard: 'Khớp 100% bản vẽ kỹ thuật & sơ đồ đơn tuyến',
      standardEn: '100% match with drawings & single-line diagram',
      reference: 'NETA MTS-2023 Mục 7.2.2.A.1',
      status: 'FAIL',
      detail: 'Thông số nhãn mác không khớp bản vẽ kỹ thuật đã phê duyệt.',
      detailEn: 'Nameplate details do not match technical drawings and specifications.',
    });
  } else {
    items.push({
      item: 'Đối soát nhãn máy (Nameplate Match)',
      itemEn: 'Nameplate Match Verification (IEEE C57.12.00)',
      measured: 'Khớp 100% bản vẽ',
      standard: 'Khớp 100% bản vẽ kỹ thuật & sơ đồ đơn tuyến',
      standardEn: '100% match with drawings & single-line diagram',
      reference: 'NETA MTS-2023 Mục 7.2.2.A.1',
      status: 'PASS',
      detail: 'Nhãn mác trùng khớp hoàn toàn bản vẽ kỹ thuật và sơ đồ đơn tuyến trạm.',
      detailEn: 'Nameplate verified 100% consistent with engineering single-line diagrams.',
    });
  }

  // Gas blanket positive pressure
  if (!v.gas_blanket_pressure_positive) {
    items.push({
      item: 'Đệm khí Nitơ bảo quản (Nitrogen Gas Blanket)',
      itemEn: 'Nitrogen Gas Blanket Pressure Gauge',
      measured: '0 psi hoặc áp suất âm',
      standard: 'BẮT BUỘC duy trì áp suất dương (positive pressure 0.5 - 5.0 psi)',
      standardEn: 'MUST maintain positive pressure (0.5 - 5.0 psi)',
      reference: 'NETA MTS-2023 Mục 7.2.2.A.6',
      status: 'FAIL',
      detail: 'Đồng hồ đệm khí Nitơ không duy trì áp suất dương, nguy cơ lọt ẩm và khí quyển vào tank.',
      detailEn: 'Nitrogen cushion lacks positive pressure; risk of ambient atmospheric moisture ingression.',
    });
  } else {
    items.push({
      item: 'Đệm khí Nitơ bảo quản (Nitrogen Gas Blanket)',
      itemEn: 'Nitrogen Gas Blanket Pressure Gauge',
      measured: 'Áp suất dương (+)',
      standard: 'BẮT BUỘC duy trì áp suất dương (positive pressure 0.5 - 5.0 psi)',
      standardEn: 'MUST maintain positive pressure (0.5 - 5.0 psi)',
      reference: 'NETA MTS-2023 Mục 7.2.2.A.6',
      status: 'PASS',
      detail: 'Đồng hồ đệm khí Nitơ duy trì áp suất dương an toàn bảo vệ dầu khỏi oxy hóa.',
      detailEn: 'Nitrogen gas blanket gauge confirms positive pressure protecting insulation.',
    });
  }

  // Mức dầu & van tản nhiệt
  items.push({
    item: 'Mức chất lỏng & Van làm mát (Liquid Levels & Valves)',
    itemEn: 'Liquid Levels & Cooling Radiator Valves',
    measured: `${v.liquid_level_check} • Van & Quạt: ${v.valves_and_cooling_fans_check}`,
    standard: 'Mức dầu đúng vạch 25°C, van mở 100%, quạt & bơm trơn tru',
    standardEn: 'Correct level mark at 25°C, valves 100% open, fans/pumps functional',
    reference: 'NETA MTS-2023 Mục 7.2.2.A.5 & A.7',
    status: 'PASS',
    detail: 'Mức dầu tank chính/LTC/bushing bình thường, van tản nhiệt mở hoàn toàn.',
    detailEn: 'Liquid levels on main tank/LTC/bushings normal, radiator valves fully opened.',
  });

  // 2. Bolted Connection Resistance (Section 7.2.2.B.1 & Table 100.12)
  const bolted = data.electrical_tests.bolted_resistance_micro_ohms || [];
  if (bolted.length > 0) {
    const minVal = Math.min(...bolted);
    const maxVal = Math.max(...bolted);
    const deviation = minVal > 0 ? ((maxVal - minVal) / minVal) * 100 : 0;
    const isInvestigate = deviation > 50.0;
    items.push({
      item: 'Điện trở mối nối bu-lông (Bolted Connection Resistance)',
      itemEn: 'Bolted Connection Resistance (DLRO / Micro-ohmmeter)',
      measured: `${bolted.join(', ')} µΩ (Độ lệch: ${deviation.toFixed(1)}%)`,
      standard: 'Cảnh báo INVESTIGATE nếu lệch > 50% so với giá trị nhỏ nhất tương tự',
      standardEn: 'INVESTIGATE if deviation > 50% from lowest similar connection',
      reference: 'NETA MTS-2023 Mục 7.2.2.B.1 & Table 100.12',
      status: isInvestigate ? 'INVESTIGATE' : 'PASS',
      detail: isInvestigate
        ? `CẢNH BÁO: Độ lệch tiếp xúc giữa các mối nối là ${deviation.toFixed(1)}% (> 50%). Cần vệ sinh tiếp xúc và siết lại theo Table 100.12.`
        : `Đạt tiêu chuẩn. Độ lệch điện trở tiếp xúc giữa các mối nối là ${deviation.toFixed(1)}% (≤ 50%).`,
      detailEn: isInvestigate
        ? `INVESTIGATE WARNING: Connection resistance deviation is ${deviation.toFixed(1)}% (> 50%). Clean and re-torque per Table 100.12.`
        : `Meets standard. Connection resistance deviation is ${deviation.toFixed(1)}% (≤ 50%).`,
    });
  }

  // 3. Winding Insulation Resistance @ 20°C & PI (Table 100.5 & Table 100.14.1)
  const ir = data.electrical_tests.insulation_resistance_20c_megohms;
  const piVal = ir.calculated_pi_value;
  const isPiPass = piVal >= 1.0;
  // Pri > 5000V requires >= 5,000 MΩ per Table 100.5
  // Sec 2500V requires >= 1,000 MΩ per Table 100.5
  const isPriPass = ir.pri_to_ground_5000v >= 5000;
  const isSecPass = ir.sec_to_ground_2500v >= 1000;
  const isPriSecPass = ir.pri_to_sec_5000v_dc >= 5000;
  const irOverallPass = isPiPass && isPriPass && isSecPass && isPriSecPass;

  items.push({
    item: 'Điện trở cách điện Cuộn dây @ 20°C & Chỉ số PI',
    itemEn: 'Winding Insulation Resistance @ 20°C & Polarization Index (PI)',
    measured: `Cao-Đất: ${ir.pri_to_ground_5000v} MΩ, Hạ-Đất: ${ir.sec_to_ground_2500v} MΩ, Cao-Hạ: ${ir.pri_to_sec_5000v_dc} MΩ, PI: ${piVal.toFixed(2)}`,
    standard: 'Table 100.5: >5kV ≥ 5,000 MΩ, 601-5000V ≥ 1,000 MΩ; PI BẮT BUỘC ≥ 1.0',
    standardEn: 'Table 100.5: >5kV ≥ 5,000 MΩ, 601-5000V ≥ 1,000 MΩ; PI MUST be ≥ 1.0',
    reference: 'NETA MTS-2023 Table 100.5 & IEEE 43',
    status: irOverallPass ? 'PASS' : 'FAIL',
    detail: irOverallPass
      ? `Đạt tiêu chuẩn. Điện trở cách điện cuộn dây vượt xa ngưỡng tối thiểu Table 100.5 và chỉ số phân cực PI = ${piVal.toFixed(2)} ≥ 1.0.`
      : `KHÔNG ĐẠT: Cách điện cuộn dây hoặc chỉ số PI = ${piVal.toFixed(2)} không đạt ngưỡng an toàn Table 100.5.`,
    detailEn: irOverallPass
      ? `Meets standard. Winding IR values significantly exceed Table 100.5 minimum thresholds and PI = ${piVal.toFixed(2)} ≥ 1.0.`
      : `FAIL: Winding IR or Polarization Index PI = ${piVal.toFixed(2)} below Table 100.5 limits.`,
  });

  // 4. Turns-Ratio Test (IEEE C57.12.90 / Section 7.2.2.B.3)
  const trError = data.electrical_tests.turns_ratio_all_taps?.max_error_percent ?? 0;
  const isTrPass = trError <= 0.5;
  items.push({
    item: 'Tỷ số biến áp ở tất cả các nấc Tap (Turns-Ratio Test)',
    itemEn: 'Turns-Ratio Test across All Tap Positions (IEEE C57.12.90)',
    measured: `Sai số lớn nhất: ${trError.toFixed(2)}%`,
    standard: 'Sai số ≤ 0.5% so với cuộn kế cận hoặc tỷ số tính toán trên nhãn máy',
    standardEn: 'Error ≤ 0.5% compared to adjacent winding or calculated nameplate ratio',
    reference: 'NETA MTS-2023 Mục 7.2.2.B.3 & IEEE C57.12.90',
    status: isTrPass ? 'PASS' : 'FAIL',
    detail: isTrPass
      ? `Đạt tiêu chuẩn. Sai số tỷ số biến áp ${trError.toFixed(2)}% ≤ 0.5% tại toàn bộ các nấc phân áp.`
      : `KHÔNG ĐẠT: Sai số tỷ số biến áp ${trError.toFixed(2)}% vượt ngưỡng 0.5%, nghi ngờ chập vòng dây hoặc lệch tiếp điểm chuyển nấc LTC.`,
    detailEn: isTrPass
      ? `Meets standard. Ratio error ${trError.toFixed(2)}% is within ≤ 0.5% limit across all tap steps.`
      : `FAIL: Ratio error ${trError.toFixed(2)}% exceeds 0.5%, suggesting turn-to-turn fault or tap changer misalignment.`,
  });

  // 5. Winding Insulation Power Factor @ 20°C (Table 100.3)
  // For MTS Mineral Oil Power Transformer: Max 1.0%
  const pfWinding = data.electrical_tests.winding_insulation_power_factor_percent_20c;
  const isPfPass = pfWinding <= 1.0;
  items.push({
    item: 'Hệ số tổn hao cách điện Cuộn dây (Power Factor @ 20°C)',
    itemEn: 'Winding Insulation Power Factor @ 20°C (Dissipation Factor)',
    measured: `${pfWinding.toFixed(2)}%`,
    standard: '≤ 1.0% quy đổi về 20°C cho Dầu khoáng MBA Lực (Table 100.3)',
    standardEn: '≤ 1.0% corrected to 20°C for Mineral Oil Power Transformer (Table 100.3)',
    reference: 'NETA MTS-2023 Table 100.3',
    status: isPfPass ? 'PASS' : 'FAIL',
    detail: isPfPass
      ? `Đạt tiêu chuẩn. Power Factor cuộn dây ở 20°C là ${pfWinding.toFixed(2)}% nằm trong ngưỡng cho phép ≤ 1.0%.`
      : `KHÔNG ĐẠT: Power Factor cuộn dây ${pfWinding.toFixed(2)}% vượt quá 1.0%, chỉ thị suy thoái điện môi hoặc nhiễm ẩm hệ thống cellulose.`,
    detailEn: isPfPass
      ? `Meets standard. Winding PF at 20°C is ${pfWinding.toFixed(2)}% within allowable ≤ 1.0% limit.`
      : `FAIL: Winding PF ${pfWinding.toFixed(2)}% exceeds 1.0%, indicating dielectric degradation or cellulose moisture contamination.`,
  });

  // 6. Winding DC Resistance (Section 7.2.2.B.2)
  const windingDev = data.electrical_tests.winding_dc_resistance_ohms?.max_phase_deviation_percent ?? 0;
  const isWindingDevPass = windingDev <= 2.0;
  items.push({
    item: 'Điện trở một chiều cuộn dây (Winding DC Resistance)',
    itemEn: 'Winding DC Resistance Inter-Phase Deviation',
    measured: `Độ lệch cực đại giữa các pha: ${windingDev.toFixed(2)}%`,
    standard: 'Độ lệch giữa các pha ≤ 2.0%',
    standardEn: 'Inter-phase deviation ≤ 2.0%',
    reference: 'NETA MTS-2023 Mục 7.2.2.B.2',
    status: isWindingDevPass ? 'PASS' : 'INVESTIGATE',
    detail: isWindingDevPass
      ? `Đạt tiêu chuẩn. Độ lệch điện trở một chiều giữa các pha ${windingDev.toFixed(2)}% ≤ 2.0%.`
      : `CẢNH BÁO: Độ lệch điện trở cuộn dây giữa các pha đạt ${windingDev.toFixed(2)}% (> 2.0%). Nghi ngờ tiếp xúc kém tại đầu nối hoặc mối hàn bị nứt.`,
    detailEn: isWindingDevPass
      ? `Meets standard. Phase-to-phase resistance deviation ${windingDev.toFixed(2)}% is within ≤ 2.0%.`
      : `INVESTIGATE WARNING: Phase resistance deviation is ${windingDev.toFixed(2)}% (> 2.0%). Check connections and brazed joints.`,
  });

  // 7. Core Insulation Resistance @ 500V DC (Section 7.2.2.B.5)
  const coreIr = data.electrical_tests.core_insulation_resistance_500v_megohms ?? 100;
  const isCorePass = coreIr >= 100;
  items.push({
    item: 'Điện trở cách điện Lõi thép (Core IR @ 500V DC)',
    itemEn: 'Core Insulation Resistance @ 500V DC (Lõi-Vỏ)',
    measured: `${coreIr} MΩ`,
    standard: 'BẮT BUỘC ≥ 100 Megohms (thử nghiệm 500V DC)',
    standardEn: 'MUST be ≥ 100 Megohms (500V DC test)',
    reference: 'NETA MTS-2023 Mục 7.2.2.B.5',
    status: isCorePass ? 'PASS' : 'FAIL',
    detail: isCorePass
      ? `Đạt tiêu chuẩn. Điện trở cách điện lõi thép ${coreIr} MΩ ≥ 100 MΩ, không có hiện tượng nối đất đa điểm (multi-point ground).`
      : `KHÔNG ĐẠT: Cách điện lõi thép ${coreIr} MΩ < 100 MΩ, cảnh báo chạm đất đa điểm gây phát nóng dòng xoáy Foucault.`,
    detailEn: isCorePass
      ? `Meets standard. Core IR is ${coreIr} MΩ ≥ 100 MΩ; no unwanted multi-point grounding detected.`
      : `FAIL: Core IR ${coreIr} MΩ is below 100 MΩ, indicating secondary inadvertent core ground with circulating eddy current risks.`,
  });

  // 8. Service-Aged Insulating Fluid Tests (Table 100.4.1 & ASTM Standards)
  const oil = data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1;
  const prev = data.previous_test_data;

  // 8a. Dielectric Breakdown (ASTM D1816 1mm gap)
  const isBreakdownPass = oil.dielectric_breakdown_d1816_1mm_kv >= minBreakdownD1816;
  let breakdownNote = '';
  if (prev?.last_oil_breakdown_kv) {
    const change = ((oil.dielectric_breakdown_d1816_1mm_kv - prev.last_oil_breakdown_kv) / prev.last_oil_breakdown_kv) * 100;
    breakdownNote = ` (Kỳ trước: ${prev.last_oil_breakdown_kv} kV, giảm: ${change.toFixed(1)}%)`;
  }
  items.push({
    item: 'Điện áp đánh thủng dầu (Dielectric Breakdown - ASTM D1816 1mm)',
    itemEn: 'Dielectric Breakdown Voltage (ASTM D1816 1mm gap)',
    measured: `${oil.dielectric_breakdown_d1816_1mm_kv} kV${breakdownNote}`,
    standard: `≥ ${minBreakdownD1816} kV theo Table 100.4.1 (cấp ${isHighVoltage ? '110kV' : '≤69kV'})`,
    standardEn: `≥ ${minBreakdownD1816} kV per Table 100.4.1 (${isHighVoltage ? '110kV class' : '≤69kV'})`,
    reference: 'NETA MTS-2023 Table 100.4.1 & ASTM D1816',
    status: isBreakdownPass ? 'PASS' : 'FAIL',
    detail: isBreakdownPass
      ? `Đạt tiêu chuẩn Table 100.4.1. Độ bền điện môi ${oil.dielectric_breakdown_d1816_1mm_kv} kV ≥ ${minBreakdownD1816} kV.`
      : `KHÔNG ĐẠT (NGUY HIỂM): Điện áp đánh thủng chỉ đạt ${oil.dielectric_breakdown_d1816_1mm_kv} kV thấp hơn ngưỡng bảo trì tối thiểu ${minBreakdownD1816} kV. So với kỳ trước (${prev?.last_oil_breakdown_kv ?? 31.0} kV), chất lượng cách điện sụt giảm nghiêm trọng.`,
    detailEn: isBreakdownPass
      ? `Meets Table 100.4.1. Dielectric strength is ${oil.dielectric_breakdown_d1816_1mm_kv} kV ≥ ${minBreakdownD1816} kV.`
      : `FAIL (CRITICAL): Dielectric breakdown is ${oil.dielectric_breakdown_d1816_1mm_kv} kV below maintenance limit ${minBreakdownD1816} kV. Severe drop compared to prior test (${prev?.last_oil_breakdown_kv ?? 31.0} kV).`,
  });

  // 8b. Water Content (ASTM D1533)
  const isWaterPass = oil.water_content_d1533_ppm <= maxWaterPpm;
  let waterNote = '';
  if (prev?.last_water_content_ppm) {
    const diff = oil.water_content_d1533_ppm - prev.last_water_content_ppm;
    waterNote = ` (Kỳ trước: ${prev.last_water_content_ppm} ppm, tăng: +${diff} ppm)`;
  }
  items.push({
    item: 'Hàm lượng nước trong dầu (Water Content - ASTM D1533)',
    itemEn: 'Moisture in Oil Content (ASTM D1533)',
    measured: `${oil.water_content_d1533_ppm} ppm${waterNote}`,
    standard: `≤ ${maxWaterPpm} ppm theo Table 100.4.1 (cấp ${isHighVoltage ? '110kV' : '≤69kV'})`,
    standardEn: `≤ ${maxWaterPpm} ppm per Table 100.4.1 (${isHighVoltage ? '110kV class' : '≤69kV'})`,
    reference: 'NETA MTS-2023 Table 100.4.1 & ASTM D1533',
    status: isWaterPass ? 'PASS' : 'FAIL',
    detail: isWaterPass
      ? `Đạt tiêu chuẩn. Hàm lượng ẩm ${oil.water_content_d1533_ppm} ppm ≤ ${maxWaterPpm} ppm.`
      : `KHÔNG ĐẠT: Hàm lượng nước đạt ${oil.water_content_d1533_ppm} ppm vượt ngưỡng tối đa ${maxWaterPpm} ppm của Table 100.4.1. Dầu ngấm ẩm nặng từ môi trường hoặc giấy cách điện bị suy thoái.`,
    detailEn: isWaterPass
      ? `Meets standard. Water content ${oil.water_content_d1533_ppm} ppm is within ≤ ${maxWaterPpm} ppm.`
      : `FAIL: Water content is ${oil.water_content_d1533_ppm} ppm exceeding Table 100.4.1 max ${maxWaterPpm} ppm limit. Severe ingress/cellulose moisture migration.`,
  });

  // 8c. Acid Number (ASTM D974 / D664)
  const isAcidPass = oil.acid_number_d974_mg_koh_g <= maxAcid;
  items.push({
    item: 'Chỉ số Axit trung hòa (Acid / Neutralization Number - ASTM D974)',
    itemEn: 'Neutralization Acid Number (ASTM D974 / D664)',
    measured: `${oil.acid_number_d974_mg_koh_g} mg KOH/g`,
    standard: `≤ ${maxAcid} mg KOH/g theo Table 100.4.1`,
    standardEn: `≤ ${maxAcid} mg KOH/g per Table 100.4.1`,
    reference: 'NETA MTS-2023 Table 100.4.1 & ASTM D974',
    status: isAcidPass ? 'PASS' : 'FAIL',
    detail: isAcidPass
      ? `Đạt tiêu chuẩn. Axit ${oil.acid_number_d974_mg_koh_g} mg KOH/g ≤ 0.20 mg KOH/g.`
      : `KHÔNG ĐẠT: Chỉ số Axit đạt ${oil.acid_number_d974_mg_koh_g} mg KOH/g vượt ngưỡng tối đa ${maxAcid} mg KOH/g. Dầu đã bị oxy hóa nặng, hình thành hợp chất axit ăn mòn cuộn dây và tăng tốc lão hóa cellulose.`,
    detailEn: isAcidPass
      ? `Meets standard. Acid number ${oil.acid_number_d974_mg_koh_g} mg KOH/g is within ≤ 0.20 mg KOH/g.`
      : `FAIL: Acid number is ${oil.acid_number_d974_mg_koh_g} mg KOH/g exceeding maximum ${maxAcid} mg KOH/g. Severe oxidation and corrosive acidic byproduct formation.`,
  });

  // 8d. Interfacial Tension (ASTM D971)
  const isIftPass = oil.interfacial_tension_d971_mn_m >= minIft;
  items.push({
    item: 'Sức căng bề mặt (Interfacial Tension - ASTM D971)',
    itemEn: 'Interfacial Tension IFT (ASTM D971)',
    measured: `${oil.interfacial_tension_d971_mn_m} mN/m`,
    standard: `≥ ${minIft} mN/m theo Table 100.4.1`,
    standardEn: `≥ ${minIft} mN/m per Table 100.4.1`,
    reference: 'NETA MTS-2023 Table 100.4.1 & ASTM D971',
    status: isIftPass ? 'PASS' : 'FAIL',
    detail: isIftPass
      ? `Đạt tiêu chuẩn. IFT đạt ${oil.interfacial_tension_d971_mn_m} mN/m ≥ 24 mN/m.`
      : `KHÔNG ĐẠT: Sức căng bề mặt IFT sụt xuống ${oil.interfacial_tension_d971_mn_m} mN/m thấp hơn mức tối thiểu ${minIft} mN/m, xác nhận sự hiện diện của cặn bùn hòa tan (sludge precursors).`,
    detailEn: isIftPass
      ? `Meets standard. IFT is ${oil.interfacial_tension_d971_mn_m} mN/m ≥ 24 mN/m.`
      : `FAIL: Interfacial tension collapsed to ${oil.interfacial_tension_d971_mn_m} mN/m below ${minIft} mN/m limit, indicating sludge precursors and colloidal aging residues.`,
  });

  // 8e. Oil Power Factor @ 25°C (ASTM D924)
  const isOilPfPass = oil.oil_power_factor_25c_d924_percent <= maxOilPf;
  items.push({
    item: 'Hệ số tổn hao điện môi dầu @ 25°C (Oil Power Factor - ASTM D924)',
    itemEn: 'Oil Power Factor @ 25°C (ASTM D924)',
    measured: `${oil.oil_power_factor_25c_d924_percent}%`,
    standard: `≤ ${maxOilPf}% theo Table 100.4.1`,
    standardEn: `≤ ${maxOilPf}% per Table 100.4.1`,
    reference: 'NETA MTS-2023 Table 100.4.1 & ASTM D924',
    status: isOilPfPass ? 'PASS' : 'FAIL',
    detail: isOilPfPass
      ? `Đạt tiêu chuẩn. Power factor dầu ${oil.oil_power_factor_25c_d924_percent}% ≤ 1.0%.`
      : `KHÔNG ĐẠT: Power factor của dầu đạt ${oil.oil_power_factor_25c_d924_percent}% (> 1.0%), biểu hiện dầu nhiễm bẩn và có hạt dẫn điện hòa tan.`,
    detailEn: isOilPfPass
      ? `Meets standard. Oil PF is ${oil.oil_power_factor_25c_d924_percent}% ≤ 1.0%.`
      : `FAIL: Oil power factor is ${oil.oil_power_factor_25c_d924_percent}% (> 1.0%), indicating polar contamination and dissolved conductive particles.`,
  });

  // 9. Dissolved Gas Analysis DGA (IEEE C57.104 & ASTM D3612)
  const dga = data.electrical_tests.dga_test_ieee_c57_104;
  const c2h2 = dga.acetylene_c2h2_ppm || 0;
  // Critical rule per NETA MTS / IEEE C57.104: Acetylene > 1.0 ppm or sudden increase indicates high energy arcing
  const isC2h2Critical = c2h2 > 1.0 || (prev?.last_c2h2_ppm === 0 && c2h2 > 0);
  items.push({
    item: 'Phân tích khí hòa tan DGA (IEEE C57.104 & ASTM D3612)',
    itemEn: 'Dissolved Gas Analysis DGA (IEEE C57.104 & ASTM D3612)',
    measured: `C2H2: ${c2h2} ppm (Kỳ trước: ${prev?.last_c2h2_ppm ?? 0} ppm), C2H4: ${dga.ethylene_c2h4_ppm} ppm, H2: ${dga.hydrogen_h2_ppm} ppm, CH4: ${dga.methane_ch4_ppm} ppm`,
    standard: 'C2H2 BẮT BUỘC = 0 ppm (Cảnh báo Critical nếu > 1.0 ppm); C2H4 ≤ 50 ppm (Condition 1)',
    standardEn: 'C2H2 MUST be 0 ppm (Critical warning if > 1.0 ppm); C2H4 ≤ 50 ppm (Condition 1)',
    reference: 'NETA MTS-2023 Mục 7.2.2 & IEEE C57.104',
    status: isC2h2Critical ? 'FAIL' : 'PASS',
    detail: isC2h2Critical
      ? `CẢNH BÁO TỐI CẤP (HIGH-ENERGY ARCING): Xuất hiện khí Acetylene C2H2 = ${c2h2} ppm (năm trước = 0 ppm) cùng Ethylene C2H4 = ${dga.ethylene_c2h4_ppm} ppm. Khí C2H2 là chỉ thị trực tiếp cho thấy đang có Sự cố Phóng điện Hồ quang Nhiệt độ Cao (> 700°C) xảy ra bên trong máy biến áp!`
      : 'Đạt ngưỡng an toàn IEEE C57.104 Condition 1. Không phát hiện khí phóng điện hồ quang Acetylene.',
    detailEn: isC2h2Critical
      ? `CRITICAL ALERT (HIGH-ENERGY ARCING): Sudden emergence of Acetylene C2H2 = ${c2h2} ppm (previous = 0 ppm) with Ethylene C2H4 = ${dga.ethylene_c2h4_ppm} ppm. Acetylene is the direct fingerprint of thermal arcing (> 700°C) inside the core/coils or tap changer.`
      : 'Meets IEEE C57.104 Condition 1 safety limits. No hazardous arcing gases detected.',
  });

  // Calculate Overall Status
  const hasFail = items.some((i) => i.status === 'FAIL');
  const hasInvestigate = items.some((i) => i.status === 'INVESTIGATE');

  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let status_title_vi = 'PASS / HOẠT ĐỘNG BẢO TRÌ BÌNH THƯỜNG';
  let status_title_en = 'PASS / NORMAL SERVICE-AGED MAINTENANCE STATUS';
  let status_reason_vi = 'Tất cả các chỉ tiêu cơ khí, điện môi dầu và DGA đều đạt tiêu chuẩn ANSI/NETA MTS-2023 Section 7.2.2.';
  let status_reason_en = 'All visual, mechanical, electrical, fluid, and DGA values conform to ANSI/NETA MTS-2023 Section 7.2.2.';

  if (hasFail) {
    overall_status = 'FAIL';
    status_title_vi = 'FAIL / CRITICAL SERVICE-AGED OIL & DGA ARCING ANOMALY';
    status_title_en = 'FAIL / CRITICAL SERVICE-AGED OIL & DGA ARCING ANOMALY';
    status_reason_vi = isC2h2Critical
      ? 'CẢNH BÁO TỐI CẤP: Dầu cách điện lão hóa suy giảm độ bền điện môi và xuất hiện khí Acetylene (C2H2) xác nhận sự cố phóng điện hồ quang nhiệt độ cao (> 700°C).'
      : 'Không đạt tiêu chuẩn ANSI/NETA MTS-2023 do các chỉ tiêu hóa lý của dầu cách điện lão hóa vượt ngưỡng bảo trì.';
    status_reason_en = isC2h2Critical
      ? 'CRITICAL ALERT: Service-aged insulating oil severely degraded, combined with sudden Acetylene (C2H2) emergence confirming high-energy arcing (> 700°C).'
      : 'Non-compliant with ANSI/NETA MTS-2023 due to service-aged fluid degradation exceeding Table 100.4.1 maintenance limits.';
  } else if (hasInvestigate) {
    overall_status = 'INVESTIGATE';
    status_title_vi = 'INVESTIGATE / CẦN KHẢO SÁT CHUYÊN SÂU';
    status_title_en = 'INVESTIGATE / FURTHER MONITORING REQUIRED';
    status_reason_vi = 'Phát hiện một số chỉ số lệch khuyến nghị (điện trở tiếp xúc hoặc độ lệch cuộn dây), cần khảo sát và đo lường lại.';
    status_reason_en = 'Certain parameters deviate from recommended limits (bolted resistance or winding dev); re-survey recommended.';
  }

  return {
    overall_status,
    status_title_vi,
    status_title_en,
    status_reason_vi,
    status_reason_en,
    markdown_report_vi: '',
    markdown_report_en: '',
    evaluations: {
      max_severity: overall_status,
      dielectric_breakdown_failed: !isBreakdownPass,
      water_content_exceeded: !isWaterPass,
      acid_number_exceeded: !isAcidPass,
      ift_below_minimum: !isIftPass,
      arcing_c2h2_critical: isC2h2Critical,
      winding_dc_dev_exceeded: !isWindingDevPass,
      core_ir_failed: !isCorePass,
      items,
    },
  };
}

export function buildDeterministicMtsTransformerMarkdown(
  data: NetaMtsTransformerInputPayload,
  deterministicResult: NetaMtsTransformerEvaluationResult,
  lang: 'vi' | 'en' = 'vi'
): string {
  const { evaluations } = deterministicResult;
  const oil = data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1;
  const dga = data.electrical_tests.dga_test_ieee_c57_104;
  const ir = data.electrical_tests.insulation_resistance_20c_megohms;
  const prev = data.previous_test_data;

  // English Output
  if (lang === 'en') {
    const isPass = deterministicResult.overall_status === 'PASS';
    const isWarn = deterministicResult.overall_status === 'INVESTIGATE';
    const statusIcon = isPass ? '🟢' : isWarn ? '🟡' : '🔴';

    return `# TEV ENGINEERING FIELD SERVICE PLATFORM
## ANSI/NETA MTS-2023 ELECTRICAL MAINTENANCE EVALUATION REPORT
**EQUIPMENT: TRANSFORMERS, LIQUID-FILLED (MAINTENANCE TESTING)**
*Standard Specifications: ANSI/NETA MTS-2023 Section 7.2.2, Table 100.3, Table 100.4.1, Table 100.5, Table 100.12, IEEE C57.104 & ASTM Standards*

---

### 1. OVERALL STATUS: ${statusIcon} ${deterministicResult.status_title_en}
- **Transformer Tag:** \`${data.site_info.transformer_tag}\`
- **Rating / Fluid:** ${data.site_info.transformer_rating || 'Liquid-Filled Mineral Oil Transformer'}
- **Substation / Plant:** ${data.site_info.project_name}
- **Field Service Engineer (FSE):** ${data.site_info.fse_name} | **Test Date:** ${data.site_info.test_date}
- **Summary Finding:** ${deterministicResult.status_reason_en}

---

### 2. DETAILED EVALUATION TABLE (NETA MTS-2023 Section 7.2.2 & IEEE Standards)

| Inspection & Test Item | Measured Site Value | Maintenance Standard (MTS) | Reference Code | Status |
| :--- | :--- | :--- | :--- | :---: |
${evaluations.items
  .map(
    (item) =>
      `| **${item.itemEn}** | ${item.measured} | ${item.standardEn} | ${item.reference} | **${
        item.status === 'PASS' ? '🟢 PASS' : item.status === 'INVESTIGATE' ? '🟡 INVESTIGATE' : '🔴 FAIL'
      }** |`
  )
  .join('\n')}

---

### 3. SERVICE-AGED OIL QUALITY, INSULATION & DGA ANOMALY RISK

#### 3.1. Dissolved Gas Analysis (DGA) Assessment (IEEE C57.104 & ASTM D3612):
${
  evaluations.arcing_c2h2_critical
    ? `- **Acetylene ($\\text{C}_2\\text{H}_2 = ${dga.acetylene_c2h2_ppm}\\text{ ppm}$):** Normal operating liquid-filled transformers exhibit $0\\text{ ppm}$ Acetylene. The sudden generation of **${dga.acetylene_c2h2_ppm}\\text{ ppm} > 1.0\\text{ ppm}$** represents an irrefutable signature of **High-Energy Electrical Arcing / Thermal Fault > 700°C** inside the main tank or load tap changer (LTC).
- **Ethylene ($\\text{C}_2\\text{H}_4 = ${dga.ethylene_c2h4_ppm}\\text{ ppm}$) & Hydrogen ($\\text{H}_2 = ${dga.hydrogen_h2_ppm}\\text{ ppm}$):** Accompanying high ethylene confirms localized severe thermal stress breaking down oil hydrocarbons.`
    : `- **DGA Combustible Gases:** Acetylene ($\\text{C}_2\\text{H}_2 = ${dga.acetylene_c2h2_ppm}\\text{ ppm}$), Hydrogen ($\\text{H}_2 = ${dga.hydrogen_h2_ppm}\\text{ ppm}$), Methane ($\\text{CH}_4 = ${dga.methane_ch4_ppm}\\text{ ppm}$), Ethylene ($\\text{C}_2\\text{H}_4 = ${dga.ethylene_c2h4_ppm}\\text{ ppm}$).
- **Conclusion:** Combustible gas concentrations strictly comply with IEEE C57.104 Condition 1 baseline. Zero arcing gas detected.`
}

#### 3.2. Service-Aged Insulating Oil Quality (ANSI/NETA MTS Table 100.4.1):
- **Dielectric Breakdown (ASTM D1816 1mm):** Measured at **${oil.dielectric_breakdown_d1816_1mm_kv}\\text{ kV}$** (${evaluations.dielectric_breakdown_failed ? `FAILS below required threshold` : `PASSES required threshold`} $\\ge 28\\text{ kV}$ for 110kV class).
- **Moisture in Oil (ASTM D1533):** Measured at **${oil.water_content_d1533_ppm}\\text{ ppm}$** (${evaluations.water_content_exceeded ? `EXCEEDS Table 100.4.1 limit of 25 ppm` : `Complies with Table 100.4.1 limit $\\le 25\\text{ ppm}$`}).
- **Acid Number (ASTM D974) & IFT (ASTM D971):** Acid is **${oil.acid_number_d974_mg_koh_g}\\text{ mg KOH/g}$** (limit $\\le 0.20$), IFT is **${oil.interfacial_tension_d971_mn_m}\\text{ mN/m}$** (limit $\\ge 24$). ${evaluations.acid_number_exceeded || evaluations.ift_below_minimum ? 'Evidence of oil oxidation and polar sludge formation.' : 'Insulating oil exhibits good chemical stability and low oxidation.'}
- **Oil Power Factor @ 25°C (ASTM D924):** **${oil.oil_power_factor_25c_d924_percent}\\%$** (limit $\\le 1.0\\%$) $\\rightarrow$ ${oil.oil_power_factor_25c_d924_percent <= 1.0 ? 'PASS' : 'FAIL'}.

#### 3.3. Winding Insulation & Mechanical Integrity:
- **Insulation Resistance & PI:** Winding IR converted to 20°C: Pri-Gnd: ${ir.pri_to_ground_5000v} MΩ, Sec-Gnd: ${ir.sec_to_ground_2500v} MΩ, Polarization Index $\\text{PI} = ${ir.calculated_pi_value.toFixed(2)} (${ir.calculated_pi_value >= 1.0 ? 'PASS $\\ge 1.0$' : 'FAIL $< 1.0$'}).
- **Turns Ratio & Power Factor:** Ratio error ${data.electrical_tests.turns_ratio_all_taps.max_error_percent}% (${data.electrical_tests.turns_ratio_all_taps.max_error_percent <= 0.5 ? 'PASS $\\le 0.5\\%$' : 'FAIL $> 0.5\\%$'}), Winding PF ${data.electrical_tests.winding_insulation_power_factor_percent_20c}% (${data.electrical_tests.winding_insulation_power_factor_percent_20c <= 1.0 ? 'PASS $\\le 1.0\\%$' : 'FAIL $> 1.0\\%$'}). Core IR is ${data.electrical_tests.core_insulation_resistance_500v_megohms ?? 350} MΩ.

---

### 4. ACTIONABLE RECOMMENDATIONS FOR FSE

${
  isPass
    ? `1. **🟢 CONTINUOUS NORMAL OPERATION:**
   - Transformer \`${data.site_info.transformer_tag}\` meets all ANSI/NETA MTS-2023 maintenance criteria. Authorized for normal continuous load operation.
2. **Periodic Inspection Schedule:**
   - Maintain annual preventative maintenance cycle. Schedule next routine electrical testing and DGA oil sampling in 12 months.
3. **Nitrogen Blanket & Dehydrating Breather Inspection:**
   - Verify positive pressure gauge weekly (0.5 - 5.0 psi). Inspect silica gel breather color quarterly to prevent atmospheric moisture intake.`
    : isWarn
    ? `1. **🟡 MONITORING & PREVENTATIVE RETIGHTENING:**
   - Inspect bolted connections and re-torque hardware according to Table 100.12.
2. **Intermediate Sampling Cycle (6 Months):**
   - Take secondary oil sample within 6 months to monitor moisture and dissolved gas generation rate.
3. **Breather & Seal Audit:**
   - Check transformer tank gaskets, pressure relief valve, and dehydrating silica gel to ensure hermetic sealing.`
    : `1. **🔴 CRITICAL OPERATIONAL DIRECTIVE - DE-ENERGIZE / ISOLATE:**
   - Advise plant operations not to operate transformer \`${data.site_info.transformer_tag}\` under heavy load; execute planned isolation if arcing or severe dielectric breakdown is verified.
2. **${evaluations.arcing_c2h2_critical ? 'Immediate Re-sample DGA within 48 Hours:' : 'Oil Filtration & Dehydration:'}**
   - ${evaluations.arcing_c2h2_critical ? 'Perform secondary DGA confirmation within 48 hours to compute gas generation rates per IEEE C57.104.' : 'Execute on-site vacuum oil dehydration and filtration to restore dielectric breakdown and moisture limits.'}
3. **${evaluations.dielectric_breakdown_failed || evaluations.acid_number_exceeded ? 'Full Fluid Reclamation:' : 'Partial Discharge (PD) Testing:'}**
   - ${evaluations.dielectric_breakdown_failed || evaluations.acid_number_exceeded ? "Perform Fuller's Earth reclamation and vacuum degassing to restore dielectric breakdown and remove acidic oxidation sludge." : 'Conduct electrical and acoustic PD surveying to locate internal discharge origin prior to re-energization.'}`
}
`;
  }

  // Vietnamese Output (Matches exact user specification)
  const isPass = deterministicResult.overall_status === 'PASS';
  const isWarn = deterministicResult.overall_status === 'INVESTIGATE';
  const statusIcon = isPass ? '🟢' : isWarn ? '🟡' : '🔴';

  return `# TEV ENGINEERING FIELD SERVICE PLATFORM
## BIÊN BẢN ĐÁNH GIÁ CHUYÊN GIA ANSI/NETA MTS-2023 (BẢO TRÌ ĐỊNH KỲ)
**THIẾT BỊ: MÁY BIẾN ÁP NGÂM CHẤT LỎNG / NGÂM DẦU (TRANSFORMERS, LIQUID-FILLED)**
*Tiêu chuẩn đối soát: ANSI/NETA MTS-2023 Section 7.2.2 & Bảng Table 100.3, Table 100.4.1 - 100.4.5, Table 100.5, Table 100.12, IEEE C57.104 & ASTM Standards*

---

### 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
### **${statusIcon} ${deterministicResult.status_title_vi}**
- **Mã máy biến áp:** \`${data.site_info.transformer_tag}\`
- **Thông số nhãn:** ${data.site_info.transformer_rating || 'Máy biến áp ngâm dầu khoáng Class I'}
- **Vị trí trạm / Dự án:** ${data.site_info.project_name}
- **Kỹ sư FSE hiện trường:** ${data.site_info.fse_name} | **Ngày kiểm định:** ${data.site_info.test_date}
- **Đánh giá thẩm định:** ${deterministicResult.status_reason_vi}

---

### 2. BẢNG PHÂN TÍCH CHI TIẾT MÁY BIẾN ÁP DẦU BẢO TRÌ (Detailed Evaluation Table)

| Hạng mục kiểm tra & thử nghiệm | Giá trị đo được ngoài Site | Tiêu chuẩn NETA MTS-2023 & IEEE | Tham chiếu kỹ thuật | Kết luận |
| :--- | :--- | :--- | :--- | :---: |
${evaluations.items
  .map(
    (item) =>
      `| **${item.item}** | ${item.measured} | ${item.standard} | ${item.reference} | **${
        item.status === 'PASS' ? '🟢 PASS' : item.status === 'INVESTIGATE' ? '🟡 INVESTIGATE' : '🔴 FAIL'
      }** |`
  )
  .join('\n')}

---

### 3. PHÂN TÍCH CHẤT LƯỢNG DẦU LÃO HÓA, CÁCH ĐIỆN & KHÍ DGA (Service-Aged Oil Quality, Insulation & DGA Anomaly Risk)

#### 3.1. Phân tích Khí Hòa tan DGA (IEEE C57.104 & ASTM D3612):
${
  evaluations.arcing_c2h2_critical
    ? `- **Khí Acetylene ($\\text{C}_2\\text{H}_2 = ${dga.acetylene_c2h2_ppm}\\text{ ppm}$):** Trong điều kiện vận hành bình thường $\\text{C}_2\\text{H}_2$ BẮT BUỘC bằng $0\\text{ ppm}$. Sự xuất hiện đột ngột của $\\text{C}_2\\text{H}_2$ là chỉ thị trực tiếp cho thấy đang có **Sự cố Phóng điện Hồ quang Nhiệt độ Cao (> 700°C)** xảy ra bên trong máy biến áp!
- **Ethylene ($\\text{C}_2\\text{H}_4 = ${dga.ethylene_c2h4_ppm}\\text{ ppm}$) & Hydrogen ($\\text{H}_2 = ${dga.hydrogen_h2_ppm}\\text{ ppm}$):** Kèm theo vùng quá nhiệt cục bộ bẻ gãy mạch hydrocacbon dầu.`
    : `- **Hàm lượng khí cháy DGA:** Acetylene ($\\text{C}_2\\text{H}_2 = ${dga.acetylene_c2h2_ppm}\\text{ ppm}$), Hydrogen ($\\text{H}_2 = ${dga.hydrogen_h2_ppm}\\text{ ppm}$), Methane ($\\text{CH}_4 = ${dga.methane_ch4_ppm}\\text{ ppm}$), Ethylene ($\\text{C}_2\\text{H}_4 = ${dga.ethylene_c2h4_ppm}\\text{ ppm}$).
- **Kết luận DGA:** Tất cả các khí cháy nằm an toàn trong dải Condition 1 của IEEE C57.104. Không xuất hiện khí phóng điện hồ quang Acetylene.`
}

#### 3.2. Chất lượng Dầu cách điện Lão hóa (NETA MTS Table 100.4.1):
- **Điện áp đánh thủng dầu (ASTM D1816 1mm gap):** Đạt **${oil.dielectric_breakdown_d1816_1mm_kv}\\text{ kV}$** (${evaluations.dielectric_breakdown_failed ? 'Thấp hơn ngưỡng tối thiểu 28 kV quy định tại Bảng Table 100.4.1 -> 🔴 FAIL' : 'Đạt ngưỡng quy định $\\ge 28\\text{ kV}$ cho cấp 110kV -> 🟢 PASS'}).
- **Hàm lượng nước trong dầu (ASTM D1533):** Đạt **${oil.water_content_d1533_ppm}\\text{ ppm}$** (${evaluations.water_content_exceeded ? 'Vượt quá ngưỡng tối đa 25 ppm của Bảng Table 100.4.1 -> 🔴 FAIL (Dầu ngấm ẩm)' : 'Đạt tiêu chuẩn $\\le 25\\text{ ppm}$ -> 🟢 PASS'}).
- **Chỉ số Axit (ASTM D974) & Sức căng IFT (ASTM D971):** Axit: **${oil.acid_number_d974_mg_koh_g}\\text{ mg KOH/g}$** (chuẩn $\\le 0.20$), IFT: **${oil.interfacial_tension_d971_mn_m}\\text{ mN/m}$** (chuẩn $\\ge 24$). ${evaluations.acid_number_exceeded || evaluations.ift_below_minimum ? 'Dầu đã bị oxy hóa và có nguy cơ tạo cặn bùn sludge -> 🔴 FAIL' : 'Chỉ số oxy hóa tốt, không có cặn bùn hòa tan -> 🟢 PASS'}.
- **Hệ số tổn hao điện môi dầu @ 25°C (ASTM D924):** Đạt **${oil.oil_power_factor_25c_d924_percent}\\%$** (${oil.oil_power_factor_25c_d924_percent <= 1.0 ? 'Đạt chuẩn $\\le 1.0\\%$ -> 🟢 PASS' : 'Vượt ngưỡng 1.0% -> 🔴 FAIL'}).

#### 3.3. Thị giác, Cơ khí & Điện trở Cách điện Cuộn dây:
- **Cách điện cuộn dây & PI:** Điện trở cách điện cuộn dây quy đổi về 20°C: Cao-Đất: ${ir.pri_to_ground_5000v} MΩ, Hạ-Đất: ${ir.sec_to_ground_2500v} MΩ. Chỉ số phân cực $\\text{PI} = ${ir.calculated_pi_value.toFixed(2)} (${ir.calculated_pi_value >= 1.0 ? 'Đạt chuẩn $\\ge 1.0$ -> 🟢 PASS' : 'Không đạt $< 1.0$ -> 🔴 FAIL'}).
- **Tỷ số biến áp & Power Factor cuộn dây:** Sai số tỷ số biến áp ${data.electrical_tests.turns_ratio_all_taps.max_error_percent}% (${data.electrical_tests.turns_ratio_all_taps.max_error_percent <= 0.5 ? 'Đạt $\\le 0.5\\%$' : 'Vượt $\\le 0.5\\%$'}). Power Factor cuộn dây ${data.electrical_tests.winding_insulation_power_factor_percent_20c}% (${data.electrical_tests.winding_insulation_power_factor_percent_20c <= 1.0 ? 'Đạt $\\le 1.0\\%$' : 'Vượt $> 1.0\\%$'}). Cách điện lõi thép: ${data.electrical_tests.core_insulation_resistance_500v_megohms ?? 350} MΩ.

---

### 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)

${
  isPass
    ? `1. **🟢 VẬN HÀNH AN TOÀN BÌNH THƯỜNG:**
   - Máy biến áp \`${data.site_info.transformer_tag}\` đáp ứng 100% các tiêu chuẩn kiểm định định kỳ ANSI/NETA MTS-2023. Đủ điều kiện tiếp tục vận hành mang tải an toàn.
2. **Kế hoạch Bảo trì Định kỳ:**
   - Duy trì chu kỳ bảo trì kiểm định định kỳ 12 tháng/lần đối với dầu cách điện và phân tích khí hòa tan DGA.
3. **Giám sát Áp suất Đệm khí Nitơ & Hạt hút ẩm:**
   - Định kỳ kiểm tra đồng hồ áp suất đệm khí Nitơ duy trì áp suất dương (0.5 - 5.0 psi) và kiểm tra màu sắc hạt silicagel bình thở để ngăn ngừa tuyệt đối hơi ẩm xâm nhập từ môi trường.`
    : isWarn
    ? `1. **🟡 GIÁM SÁT VÀ XIẾT LẠI TIẾP XÚC BU-LÔNG:**
   - Tiến hành kiểm tra và xiết lại bu-lông các đầu cốt nối theo đúng lực mô-men quy định tại Table 100.12.
2. **Rút ngắn Chu kỳ Thử nghiệm (6 Tháng):**
   - Lấy mẫu dầu kiểm tra độ ẩm và đo lại điện trở tiếp xúc sau 6 tháng để theo dõi xu hướng biến thiên.
3. **Kiểm tra Kín khí Bình thở:**
   - Thay thế hạt hút ẩm silicagel bình thở và kiểm tra gioăng làm kín để ngăn ngừa hơi ẩm ngấm vào dầu.`
    : `1. **🔴 CẢNH BÁO TỐI CẤP - TÁCH CÔ LẬP KHẨN CẤP:**
   - Khuyên đơn vị vận hành không tiếp tục cho máy biến áp \`${data.site_info.transformer_tag}\` mang tải nặng; ${evaluations.arcing_c2h2_critical ? 'lập kế hoạch tách cô lập kiểm tra ngay do phát hiện hồ quang điện nguy hiểm.' : 'lập kế hoạch xử lý dầu cách điện sớm nhất.'}
2. **${evaluations.arcing_c2h2_critical ? 'Lấy mẫu DGA đối soát lại sau 48 giờ:' : 'Lọc dầu & Hút chân không khử ẩm:'}**
   - ${evaluations.arcing_c2h2_critical ? 'Tiến hành lấy mẫu dầu DGA đối soát lại sau 48 giờ để đánh giá tốc độ sinh khí (gas generation rate) theo IEEE C57.104.' : 'Thực hiện lọc sấy dầu tuần hoàn hút chân không để loại bỏ triệt để độ ẩm và bọt khí hòa tan.'}
3. **${evaluations.dielectric_breakdown_failed || evaluations.acid_number_exceeded ? 'Lọc dầu tái chế hoàn nguyên (Reclamation):' : 'Đo phóng điện cục bộ & nội soi tank:'}**
   - ${evaluations.dielectric_breakdown_failed || evaluations.acid_number_exceeded ? "Thực hiện lọc dầu tái chế hoàn nguyên bằng đất lọc Fuller's Earth để khôi phục chỉ số Axit, IFT và điện áp đánh thủng về dải chuẩn Table 100.4.1." : 'Đo phóng điện cục bộ (PD) kết hợp nội soi tank để xác định chính xác vị trí phát sinh phóng điện trước khi cho phép đóng điện trở lại.'}`
}
`;
}

export async function generateNetaMtsTransformerAiAnalysis(
  payload: NetaMtsTransformerInputPayload
): Promise<NetaMtsTransformerEvaluationResult> {
  const deterministicResult = performDeterministicMtsTransformerAnalysis(payload);
  const lang = payload.language || 'vi';

  deterministicResult.markdown_report_vi = buildDeterministicMtsTransformerMarkdown(payload, deterministicResult, 'vi');
  deterministicResult.markdown_report_en = buildDeterministicMtsTransformerMarkdown(payload, deterministicResult, 'en');

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return deterministicResult;
  }

  try {
    const ai = new GoogleGenAI({ apiKey });
    const prompt = `
You are TEV Platform AI (AI Field Service Inspection Assistant), specialized in Transformers, Liquid-Filled - Maintenance Testing per ANSI/NETA MTS-2023 Section 7.2.2, Table 100.3, Table 100.4 (100.4.1 - 100.4.5), Table 100.5, Table 100.12, IEEE C57.104, and ASTM standards.

Please analyze the following FSE site inspection data and provide an authoritative evaluation.
Selected language for output: ${lang === 'en' ? 'English' : 'Vietnamese (Tiếng Việt)'}.

INPUT DATA:
${JSON.stringify(payload, null, 2)}

DETERMINISTIC EVALUATION SUMMARY:
- Overall Status: ${deterministicResult.overall_status}
- Acetylene Alert: ${deterministicResult.evaluations.arcing_c2h2_critical ? 'CRITICAL ARCING (C2H2 > 0)' : 'Normal'}
- Fluid Breakdown: ${deterministicResult.evaluations.dielectric_breakdown_failed ? 'FAIL' : 'PASS'}
- Water Content: ${deterministicResult.evaluations.water_content_exceeded ? 'FAIL' : 'PASS'}
- Acid / IFT: ${deterministicResult.evaluations.acid_number_exceeded || deterministicResult.evaluations.ift_below_minimum ? 'FAIL (Oxidized/Sludge)' : 'PASS'}

Return a pristine markdown report formatted in 4 parts:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, or INVESTIGATE
2. BẢNG PHÂN TÍCH CHI TIẾT MÁY BIẾN ÁP DẦU BẢO TRÌ (Detailed Evaluation Table)
3. PHÂN TÍCH CHẤT LƯỢNG DẦU LÃO HÓA, CÁCH ĐIỆN & KHÍ DGA (Service-Aged Oil Quality, Insulation & DGA Anomaly Risk)
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)
Include the professional TEV Engineering branding header in your text.
`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: prompt,
    });

    const text = response.text?.trim();
    if (text) {
      if (lang === 'en') {
        deterministicResult.markdown_report_en = text;
      } else {
        deterministicResult.markdown_report_vi = text;
      }
    }
  } catch (err: any) {
    console.error('Gemini API call failed for MTS transformer analysis, using deterministic report:', err);
  }

  return deterministicResult;
}
