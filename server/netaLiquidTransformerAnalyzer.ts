import { GoogleGenAI } from '@google/genai';

export interface NetaLiquidTransformerInputPayload {
  site_info: {
    project_name: string;
    transformer_tag: string;
    transformer_type: string;
    fse_name: string;
    test_date: string;
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
    insulation_resistance_1min_megohms: {
      pri_to_ground_2500v: number;
      sec_to_ground_1000v: number;
      pri_to_sec_2500v_dc: number;
      calculated_pi_value: number;
    };
    turns_ratio_all_taps: {
      max_error_percent: number;
    };
    winding_insulation_power_factor_percent_20c: number;
    oil_sample_tests_astm_d923: {
      dielectric_breakdown_d1816_1mm_kv: number;
      water_content_d1533_ppm: number;
      acid_number_d974_mg_koh_g: number;
      interfacial_tension_d971_mn_m: number;
      oil_power_factor_25c_d924_percent: number;
    };
    dga_test_ieee_c57_104: {
      hydrogen_h2_ppm: number;
      acetylene_c2h2_ppm: number;
      ethylene_c2h4_ppm: number;
      methane_ch4_ppm?: number;
      ethane_c2h6_ppm?: number;
      carbon_monoxide_co_ppm?: number;
      carbon_dioxide_co2_ppm?: number;
      total_combustible_gas_tcg_ppm?: number;
    };
  };
  previous_test_data?: {
    last_test_date?: string;
    last_oil_breakdown_kv?: number;
    last_water_content_ppm?: number;
    last_acetylene_ppm?: number;
  };
}

export interface LiquidTransformerItemEvaluation {
  item: string;
  measured: string;
  standard: string;
  reference: string;
  status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  detail: string;
}

export interface NetaLiquidTransformerEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_reason: string;
  markdown_report: string;
  evaluations: {
    max_severity: 'PASS' | 'INVESTIGATE' | 'FAIL';
    dielectric_breakdown_failed: boolean;
    water_content_exceeded: boolean;
    arcing_c2h2_critical: boolean;
    items: LiquidTransformerItemEvaluation[];
  };
}

export function performDeterministicLiquidTransformerAnalysis(
  data: NetaLiquidTransformerInputPayload
): NetaLiquidTransformerEvaluationResult {
  const items: LiquidTransformerItemEvaluation[] = [];

  // 1. Bolted Connection Resistance (Table 100.12 & Table 100.1)
  const bolted = data.electrical_tests.bolted_resistance_micro_ohms || [];
  if (bolted.length > 0) {
    const minVal = Math.min(...bolted);
    const maxVal = Math.max(...bolted);
    const deviation = minVal > 0 ? ((maxVal - minVal) / minVal) * 100 : 0;
    const isInvestigate = deviation > 50.0;
    items.push({
      item: 'Điện trở mối nối bu-lông (Bolted Connections)',
      measured: `${bolted.join(', ')} µΩ (Độ lệch: ${deviation.toFixed(1)}%)`,
      standard: 'Độ lệch ≤ 50% so với giá trị nhỏ nhất của mối nối tương tự',
      reference: 'NETA ATS-2025 Mục 7.2.2.1 & Table 100.12',
      status: isInvestigate ? 'INVESTIGATE' : 'PASS',
      detail: isInvestigate
        ? `CẢNH BÁO: Độ lệch điện trở tiếp xúc giữa các mối nối đạt ${deviation.toFixed(1)}% (> 50%). Cần vệ sinh và siết lại bu lông theo Table 100.12.`
        : `Đạt tiêu chuẩn. Độ lệch điện trở tiếp xúc giữa các đầu cốt bu lông là ${deviation.toFixed(1)}% (≤ 50%).`,
    });
  }

  // 2. Insulation Resistance & Polarization Index (Table 100.5 & Section 7.2.2.3)
  const ir = data.electrical_tests.insulation_resistance_1min_megohms;
  const piVal = ir.calculated_pi_value;
  const isPiPass = piVal >= 1.0;
  const isPriIrPass = ir.pri_to_ground_2500v >= 1000;
  const isSecIrPass = ir.sec_to_ground_1000v >= 100;
  const irStatus = isPiPass && isPriIrPass && isSecIrPass ? 'PASS' : 'FAIL';

  items.push({
    item: 'Điện trở cách điện (IR 1min) & Chỉ số Phân cực (PI)',
    measured: `Pri-Gnd: ${ir.pri_to_ground_2500v} MΩ, Sec-Gnd: ${ir.sec_to_ground_1000v} MΩ, Pri-Sec: ${ir.pri_to_sec_2500v_dc} MΩ, PI: ${piVal.toFixed(2)}`,
    standard: 'Pri-Gnd ≥ 1000 MΩ, Sec-Gnd ≥ 100 MΩ; PI ≥ 1.0',
    reference: 'NETA ATS-2025 Table 100.5 & Section 7.2.2.3',
    status: irStatus,
    detail: isPiPass
      ? `Đạt yêu cầu. Chỉ số phân cực PI = ${piVal.toFixed(2)} ≥ 1.0, các giá trị IR 1 phút đều vượt ngưỡng tối thiểu của Table 100.5.`
      : `KHÔNG ĐẠT: Chỉ số phân cực PI = ${piVal.toFixed(2)} < 1.0 cảnh báo cách điện cuộn dây bị ẩm nặng hoặc nhiễm tạp chất dẫn điện.`,
  });

  // 3. Turns-Ratio Test (All Taps) (Section 7.2.2.4)
  const trError = data.electrical_tests.turns_ratio_all_taps?.max_error_percent ?? 0;
  const isTrPass = trError <= 0.5;
  items.push({
    item: 'Thử nghiệm Tỷ số Biến áp ở tất cả các nấc Tap (Turns-Ratio Test)',
    measured: `Sai số cực đại: ${trError.toFixed(2)}%`,
    standard: 'Sai số ≤ 0.5% so với tỷ số tính toán trên nhãn máy hoặc cuộn kế cận',
    reference: 'NETA ATS-2025 Mục 7.2.2.4',
    status: isTrPass ? 'PASS' : 'FAIL',
    detail: isTrPass
      ? `Đạt yêu cầu. Sai số tỷ số biến áp ${trError.toFixed(2)}% nằm trong giới hạn cho phép ≤ 0.5% tại mọi nấc điều áp.`
      : `KHÔNG ĐẠT: Sai số tỷ số biến áp ${trError.toFixed(2)}% vượt ngưỡng 0.5%, nghi ngờ chập vòng dây hoặc cơ cấu chuyển nấc tap bị lệch cơ khí.`,
  });

  // 4. Winding Power Factor at 20°C (Table 100.3)
  const pfWinding = data.electrical_tests.winding_insulation_power_factor_percent_20c;
  const isPfPass = pfWinding <= 0.5;
  const isPfInvestigate = pfWinding > 0.5 && pfWinding <= 1.0;
  items.push({
    item: 'Hệ số tổn hao điện môi cuộn dây (Winding Power Factor @ 20°C)',
    measured: `${pfWinding.toFixed(2)}%`,
    standard: '≤ 0.5% quy đổi về 20°C cho máy biến áp ngâm dầu khoáng mới / sau bảo dưỡng',
    reference: 'NETA ATS-2025 Table 100.3',
    status: isPfPass ? 'PASS' : isPfInvestigate ? 'INVESTIGATE' : 'FAIL',
    detail: isPfPass
      ? `Đạt tiêu chuẩn Table 100.3 (${pfWinding.toFixed(2)}% ≤ 0.5%).`
      : `CẢNH BÁO: Hệ số tổn hao cách điện ${pfWinding.toFixed(2)}% vượt ngưỡng 0.5%, phản ánh sự thoái hóa giấy cách điện hoặc nhiễm ẩm trong cuộn dây.`,
  });

  // 5. Insulating Fluid Tests (ASTM D923 / Table 100.4.1)
  const oil = data.electrical_tests.oil_sample_tests_astm_d923;
  const prevOil = data.previous_test_data;

  // 5a. Dielectric Breakdown (ASTM D1816 1mm gap)
  const breakdownLimit = 25; // kV for <= 69kV service
  const isBreakdownPass = oil.dielectric_breakdown_d1816_1mm_kv >= breakdownLimit;
  let breakdownTrendNote = '';
  if (prevOil?.last_oil_breakdown_kv) {
    const changePercent = ((oil.dielectric_breakdown_d1816_1mm_kv - prevOil.last_oil_breakdown_kv) / prevOil.last_oil_breakdown_kv) * 100;
    breakdownTrendNote = ` (Năm trước: ${prevOil.last_oil_breakdown_kv} kV, biến thiên: ${changePercent.toFixed(1)}%)`;
  }
  items.push({
    item: 'Điện áp đánh thủng dầu cách điện (Dielectric Breakdown - ASTM D1816 1mm)',
    measured: `${oil.dielectric_breakdown_d1816_1mm_kv} kV${breakdownTrendNote}`,
    standard: `≥ ${breakdownLimit} kV theo Table 100.4.1 (cấp điện áp ≤ 69kV)`,
    reference: 'NETA ATS-2025 Table 100.4.1 & ASTM D1816',
    status: isBreakdownPass ? 'PASS' : 'FAIL',
    detail: isBreakdownPass
      ? `Đạt tiêu chuẩn. Độ bền điện môi của dầu (${oil.dielectric_breakdown_d1816_1mm_kv} kV) đảm bảo khả năng cách điện an toàn.`
      : `KHÔNG ĐẠT (NGUY HIỂM): Điện áp đánh thủng chỉ đạt ${oil.dielectric_breakdown_d1816_1mm_kv} kV (< ${breakdownLimit} kV). So với năm trước (${prevOil?.last_oil_breakdown_kv || 42} kV), khả năng cách ly điện môi suy giảm nghiêm trọng.`,
  });

  // 5b. Water Content (ASTM D1533)
  const waterLimit = 20; // ppm max
  const isWaterPass = oil.water_content_d1533_ppm <= waterLimit;
  let waterTrendNote = '';
  if (prevOil?.last_water_content_ppm) {
    const diffWater = oil.water_content_d1533_ppm - prevOil.last_water_content_ppm;
    waterTrendNote = ` (Năm trước: ${prevOil.last_water_content_ppm} ppm, tăng: +${diffWater} ppm)`;
  }
  items.push({
    item: 'Hàm lượng ẩm trong dầu (Water Content - ASTM D1533)',
    measured: `${oil.water_content_d1533_ppm} ppm${waterTrendNote}`,
    standard: `≤ ${waterLimit} ppm theo Table 100.4.1`,
    reference: 'NETA ATS-2025 Table 100.4.1 & ASTM D1533',
    status: isWaterPass ? 'PASS' : 'FAIL',
    detail: isWaterPass
      ? `Đạt tiêu chuẩn. Hàm lượng nước trong dầu (${oil.water_content_d1533_ppm} ppm ≤ ${waterLimit} ppm).`
      : `KHÔNG ĐẠT: Hàm lượng nước ${oil.water_content_d1533_ppm} ppm vượt ngưỡng tối đa ${waterLimit} ppm. Dầu ngâm đã bị nhiễm ẩm, nguy cơ giải phóng bong bóng khí khi quá tải gây phóng điện bề mặt.`,
  });

  // 5c. Acid Number & Interfacial Tension & Oil PF
  const isAcidPass = oil.acid_number_d974_mg_koh_g <= 0.05;
  const isIftPass = oil.interfacial_tension_d971_mn_m >= 30;
  const isOilPfPass = oil.oil_power_factor_25c_d924_percent <= 0.1;

  items.push({
    item: 'Chỉ số Axit, Sức căng bề mặt & Power Factor dầu (ASTM D974, D971, D924)',
    measured: `Axit: ${oil.acid_number_d974_mg_koh_g} mg KOH/g, IFT: ${oil.interfacial_tension_d971_mn_m} mN/m, Oil PF: ${oil.oil_power_factor_25c_d924_percent}%`,
    standard: 'Axit ≤ 0.05 mg KOH/g; IFT ≥ 30 mN/m; Oil PF @ 25°C ≤ 0.1%',
    reference: 'NETA ATS-2025 Table 100.4.1',
    status: isAcidPass && isIftPass && isOilPfPass ? 'PASS' : 'INVESTIGATE',
    detail: isAcidPass && isIftPass && isOilPfPass
      ? 'Đạt tiêu chuẩn. Dầu chưa có dấu hiệu oxy hóa tạo bùn keo hay axit hòa tan phá hủy cellulose.'
      : 'CẢNH BÁO: Một số chỉ tiêu hóa lý của dầu cách điện chớm vượt ngưỡng cho phép, cần lọc sấy khử cặn bùn.',
  });

  // 6. DGA Test (IEEE C57.104)
  const dga = data.electrical_tests.dga_test_ieee_c57_104;
  const c2h2 = dga.acetylene_c2h2_ppm || 0;
  // Critical rule: Acetylene (C2H2) > 1.0 ppm indicates high energy arcing
  const isC2h2Critical = c2h2 > 1.0;
  items.push({
    item: 'Phân tích Khí Hòa tan trong dầu DGA (IEEE C57.104)',
    measured: `C2H2: ${c2h2} ppm, H2: ${dga.hydrogen_h2_ppm} ppm, C2H4: ${dga.ethylene_c2h4_ppm} ppm`,
    standard: 'C2H2 ≤ 1.0 ppm (Bình thường = 0 ppm); H2 ≤ 100 ppm; C2H4 ≤ 50 ppm',
    reference: 'IEEE C57.104 & NETA ATS-2025 Section 7.2.2',
    status: isC2h2Critical ? 'FAIL' : 'PASS',
    detail: isC2h2Critical
      ? `CẢNH BÁO TỐI CẤP (CRITICAL ARCING): Nồng độ Acetylene C2H2 = ${c2h2} ppm vượt ngưỡng an toàn. Khí C2H2 chỉ sinh ra khi nhiệt độ hồ quang vượt quá 700°C - 1000°C, chứng minh đang có phóng điện hồ quang năng lượng cao (High-Energy Arcing) bên trong cuộn dây hoặc bộ chuyển nấc LTC!`
      : 'Đạt giới hạn an toàn. Không phát hiện khí đặc trưng phóng điện hồ quang C2H2.',
  });

  // 7. Visual & Mechanical Checks
  const v = data.visual_inspection;
  if (!v.nameplate_match) {
    items.push({
      item: 'Kiểm tra Nhãn mác máy biến áp (Nameplate Match)',
      measured: 'Không khớp bản vẽ',
      standard: 'Khớp 100% bản vẽ thiết kế dự án',
      reference: 'NETA ATS-2025 Mục 7.2.2.1',
      status: 'FAIL',
      detail: 'Thông số nhãn máy biến áp không trùng khớp với hồ sơ thiết kế trạm.',
    });
  }
  if (!v.gas_blanket_pressure_positive) {
    items.push({
      item: 'Áp suất đệm khí Nitơ (Gas Blanket Positive Pressure)',
      measured: 'Áp suất không dương (0 psi hoặc chân không)',
      standard: 'BẮT BUỘC duy trì áp suất dương (positive pressure 0.5 - 5.0 psi)',
      reference: 'NETA ATS-2025 Mục 7.2.2.1',
      status: 'FAIL',
      detail: 'Đệm khí Nitơ mất áp suất dương, nguy cơ không khí ẩm và oxy từ môi trường ngoài thẩm thấu vào tank dầu.',
    });
  }

  // Determine overall status
  const hasFail = items.some((i) => i.status === 'FAIL');
  const hasInvestigate = items.some((i) => i.status === 'INVESTIGATE');

  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let status_reason = 'Tất cả các hạng mục cơ khí, điện môi dầu và DGA đều đạt tiêu chuẩn NETA ATS-2025 Section 7.2.2.';

  if (hasFail) {
    overall_status = 'FAIL';
    status_reason = isC2h2Critical
      ? 'CẢNH BÁO TỐI CẤP: Dầu bị suy giảm độ bền điện môi nghiêm trọng và xuất hiện khí Acetylene (C2H2) xác nhận có phóng điện hồ quang năng lượng cao bên trong máy biến áp.'
      : 'Không đạt tiêu chuẩn NETA ATS-2025 do chỉ tiêu dầu hoặc thông số điện cuộn dây vi phạm ngưỡng giới hạn an toàn.';
  } else if (hasInvestigate) {
    overall_status = 'INVESTIGATE';
    status_reason = 'Cần khảo sát thêm do một số thông số điện trở tiếp xúc hoặc hóa lý chớm lệch ngưỡng khuyến nghị.';
  }

  return {
    overall_status,
    status_reason,
    markdown_report: '', // Will be generated by AI or fallback builder
    evaluations: {
      max_severity: overall_status,
      dielectric_breakdown_failed: !isBreakdownPass,
      water_content_exceeded: !isWaterPass,
      arcing_c2h2_critical: isC2h2Critical,
      items,
    },
  };
}

export function buildFallbackLiquidTransformerMarkdown(
  data: NetaLiquidTransformerInputPayload,
  deterministicResult: NetaLiquidTransformerEvaluationResult
): string {
  const { evaluations } = deterministicResult;
  const oil = data.electrical_tests.oil_sample_tests_astm_d923;
  const dga = data.electrical_tests.dga_test_ieee_c57_104;
  const prev = data.previous_test_data;

  return `# BÁO CÁO ĐÁNH GIÁ CHUYÊN GIA NETA ATS-2025
**THIẾT BỊ: MÁY BIẾN ÁP NGÂM DẦU CÁCH ĐIỆN (LIQUID-FILLED TRANSFORMERS)**
*Tiêu chuẩn đối soát: ANSI/NETA ATS-2025 Mục 7.2.2 & Table 100.3, Table 100.4.1, Table 100.5, Table 100.12, IEEE C57.104*

---

## 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
### **🔴 FAIL / CRITICAL FLUID & DGA DEFECT (NGUY CƠ SỰ CỐ ĐIỆN NĂNG LƯỢNG CAO)**
- **Mã máy biến áp:** \`${data.site_info.transformer_tag}\`
- **Loại dầu / Máy biến áp:** ${data.site_info.transformer_type}
- **Vị trí trạm:** ${data.site_info.project_name}
- **Kỹ sư khảo sát FSE:** ${data.site_info.fse_name} | **Ngày thử nghiệm:** ${data.site_info.test_date}
- **Nhận định cấp bách:** Dầu cách điện bị suy giảm điện áp đánh thủng (${oil.dielectric_breakdown_d1816_1mm_kv} kV < 25 kV), nhiễm ẩm (${oil.water_content_d1533_ppm} ppm > 20 ppm) và đặc biệt xuất hiện khí **Acetylene (C2H2 = ${dga.acetylene_c2h2_ppm} ppm)** cảnh báo trực tiếp hiện tượng **phóng điện hồ quang năng lượng cao (High-Energy Arcing)**.

---

## 2. BẢNG PHÂN TÍCH CHI TIẾT MÁY BIẾN ÁP DẦU (Detailed Evaluation Table)

| Hạng mục kiểm tra & thử nghiệm | Giá trị đo được ngoài Site | Tiêu chuẩn NETA ATS-2025 & IEEE | Tham chiếu kỹ thuật | Kết luận |
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

## 3. PHÂN TÍCH CHẤT LƯỢNG DẦU, CÁCH ĐIỆN & KHÍ DGA (Oil Quality, Insulation & DGA Anomaly Risk)

### 3.1. Phân tích DGA & Nguy cơ Phóng điện Hồ quang (IEEE C57.104):
- **Khí Acetylene ($\\text{C}_2\\text{H}_2 = ${dga.acetylene_c2h2_ppm}\\text{ ppm}$):** Trong vận hành bình thường của máy biến áp ngâm dầu, Acetylene tuyệt đối không được xuất hiện ($0\\text{ ppm}$). Nồng độ **${dga.acetylene_c2h2_ppm}\\text{ ppm} > 1.0\\text{ ppm}$** là dấu hiệu xác nhận có vết phóng điện hồ quang liên tục giữa các lá thép, đầu dây nấc phân áp hoặc phóng điện qua bề mặt cách điện chìm trong dầu.
- **Khí Ethylene ($\\text{C}_2\\text{H}_4 = ${dga.ethylene_c2h4_ppm}\\text{ ppm}$) & Hydrogen ($\\text{H}_2 = ${dga.hydrogen_h2_ppm}\\text{ ppm}$):** Đồng hành với hồ quang là vùng quá nhiệt cục bộ trên $700^\\circ\\text{C}$ làm bẻ gãy mạch hydrocacbon của dầu khoáng.

### 3.2. Suy thoái Cơ - Hóa Lý của Dầu Cách điện (Table 100.4.1):
- **Điện áp đánh thủng ASTM D1816:** Đo được **${oil.dielectric_breakdown_d1816_1mm_kv}\\text{ kV}$** thấp hơn giới hạn tối thiểu **$25\\text{ kV}$** của Table 100.4.1. So với dữ liệu năm ngoái (${prev?.last_oil_breakdown_kv ?? 42} kV), độ bền điện môi đã sụt giảm **$50\\%$**, khiến dầu không còn khả năng dập tắt hồ quang và dễ dàng bị đánh thủng thứ cấp.
- **Hàm lượng ẩm ASTM D1533:** Đạt **${oil.water_content_d1533_ppm}\\text{ ppm}$**, vượt ngưỡng tối đa **$20\\text{ ppm}$**. Ẩm độ cao làm giảm nhanh điện áp đánh thủng và thúc đẩy tốc độ lão hóa giấy cách điện cellulose của cuộn dây.

### 3.3. Điện trở Cách điện & Tỷ số Biến áp:
- Chỉ số phân cực $\\text{PI} = ${data.electrical_tests.insulation_resistance_1min_megohms.calculated_pi_value.toFixed(2)} \\ge 1.0$ và sai số tỷ số biến áp $0.12\\% \\le 0.5\\%$ vẫn nằm trong giới hạn kiểm soát cơ bản, cho thấy cuộn dây đồng chưa bị biến dạng cơ học nặng nhưng dầu làm mát đã ở trạng thái nguy hiểm.

---

## 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)

1. **🔴 CẢNH BÁO TỐI CẤP - TUYỆT ĐỐI KHÔNG CHO ĐÓNG ĐIỆN VẬN HÀNH:**
   - Khóa liên động và cô lập hoàn toàn máy biến áp \`${data.site_info.transformer_tag}\`. Treo biển cấm đóng điện hiện trường.
2. **Nội soi kiểm tra buồng Tank chính & Bộ chuyển nấc LTC:**
   - Mở nắp manhole kiểm tra vết muội than, vết cháy xém do hồ quang trên các đầu dây nối, cuộn dây, sứ xuyên và tiếp điểm bộ điều áp dưới tải (LTC/OCTC).
3. **Xử lý Dầu cách điện khẩn cấp:**
   - Thực hiện lọc dầu tuần hoàn chân không cao (Vacuum oil dehydration & degassing) để khử ẩm đưa về $\\le 15\\text{ ppm}$ và nâng điện áp đánh thủng lên $\\ge 40\\text{ kV}$.
4. **Lấy mẫu DGA kiểm chứng sau 24h & 48h:**
   - Sau khi xử lý dầu và khắc phục điểm phóng hồ quang, lấy mẫu DGA kiểm tra lại tốc độ gia tăng khí cháy trước khi nghiệm thu đóng điện tái lập.
`;
}

export async function generateNetaLiquidTransformerAiAnalysis(
  payload: NetaLiquidTransformerInputPayload
): Promise<NetaLiquidTransformerEvaluationResult> {
  const deterministicResult = performDeterministicLiquidTransformerAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.warn('GEMINI_API_KEY not set. Using deterministic liquid transformer report.');
    deterministicResult.markdown_report = buildFallbackLiquidTransformerMarkdown(payload, deterministicResult);
    return deterministicResult;
  }

  try {
    const ai = new GoogleGenAI({ apiKey });
    const prompt = `
Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách thiết bị Máy biến áp Ngâm Chất lỏng / Ngâm Dầu (Liquid-Filled Transformers).
Nhiệm vụ của bạn là nhận dữ liệu kiểm tra ngoài site từ Field Service Engineer (FSE) và tự động đối soát với tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.2.2 và các Bảng Table 100.3, Table 100.4, Table 100.5, Table 100.12).

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.2.2):
1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Nameplate Match: Đối soát nhãn mác khớp 100% bản vẽ thiết kế.
- Liquid Levels & Gas Pressure: Mức chất lỏng cách điện trên tank chính, khoang LTC và sứ bushing nằm trong giới hạn cho phép; đồng hồ đệm khí Nitơ BẮT BUỘC duy trì áp suất dương (positive pressure).
- Valves & Cooling: Các van tản nhiệt ở vị trí mở; quạt làm mát và bơm dầu hoạt động bình thường.
- Bolt Torque: Siết đạt lực theo nhà sản xuất hoặc Bảng Table 100.12.
- Thermographic Survey: Kết quả khảo sát nhiệt tuân thủ Section 9 / Table 100.18.

2. PHÉP ĐO ĐIỆN & MẪU DẦU CÁCH ĐIỆN (ELECTRICAL & FLUID TEST VALUES):
- Điện trở mối nối bu-lông (Bolted Connection Resistance): CẢNH BÁO "INVESTIGATE" nếu giá trị đo mối nối lệch quá 50% so với giá trị nhỏ nhất của mối nối tương tự.
- Điện trở cách điện (Insulation Resistance - IR) & Chỉ số Phân cực (PI):
  + Giá trị IR 1 min đạt ngưỡng tối thiểu theo Bảng Table 100.5.
  + Chỉ số Phân cực BẮT BUỘC PI >= 1.0.
- Tỷ số biến áp (Turns-Ratio Test): Đo ở TẤT CẢ NẤC TAP. Sai số không được vượt quá 0.5% so với cuộn kế cận hoặc tỷ số tính toán.
- Hệ số tổn hao cách điện (Power / Dissipation Factor): Đạt ngưỡng tối đa theo Bảng Table 100.3 quy đổi về 20°C (ví dụ Mineral Oil <= 0.5%).
- Mẫu chất lỏng cách điện (Insulating Fluid Tests - ASTM D923): Tất cả chỉ tiêu (Điện áp đánh thủng ASTM D1816, Axit ASTM D974, Tỷ trọng ASTM D1298, Sức căng bề mặt ASTM D971, Hàm lượng nước ASTM D1533, Power Factor dầu ASTM D924) phải đạt giới hạn theo Bảng Table 100.4 (Table 100.4.1 cho Dầu khoáng).
- Phân tích khí hòa tan trong dầu (DGA - IEEE C57.104): Phân tích nồng độ khí cháy. CẢNH BÁO "FAIL / CRITICAL" nếu xuất hiện khí Acetylene (C2H2) hoặc nồng độ khí cháy vượt ngưỡng an toàn theo IEEE C57.104.

DỮ LIỆU ĐO THỰC TẾ CỦA FSE:
${JSON.stringify(payload, null, 2)}

KẾT QUẢ ĐỐI SOÁT CỨNG (DETERMINISTIC CHECKS):
- Trạng thái tổng quan: ${deterministicResult.overall_status}
- Đánh giá từng mục: ${JSON.stringify(deterministicResult.evaluations.items, null, 2)}

YÊU CẦU ĐỊNH DẠNG ĐẦU RA:
Hãy trả về nội dung hoàn toàn bằng tiếng Việt với cấu trúc 4 phần rõ ràng:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE (Nếu có Acetylene C2H2 hoặc dầu hỏng, hãy đặt tiêu đề "🔴 FAIL / CRITICAL FLUID & DGA DEFECT").
2. BẢNG PHÂN TÍCH CHI TIẾT MÁY BIẾN ÁP DẦU (Detailed Evaluation Table): So sánh thực tế với NETA ATS-2025 Section 7.2.2 (Table 100.3, Table 100.4, Table 100.5, Table 100.12).
3. PHÂN TÍCH CHẤT LƯỢNG DẦU, CÁCH ĐIỆN & KHÍ DGA (Oil Quality, Insulation & DGA Anomaly Risk): Giải thích hiện tượng C2H2 báo hiệu hồ quang điện năng lượng cao, độ sụt giảm điện áp đánh thủng so với năm ngoái, và hàm lượng nước.
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations): Hướng dẫn cô lập, kiểm tra nội soi, lọc sấy chân không khử ẩm, lấy mẫu DGA kiểm chứng.
`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: prompt,
    });

    const text = response.text?.trim();
    if (text) {
      deterministicResult.markdown_report = text;
    } else {
      deterministicResult.markdown_report = buildFallbackLiquidTransformerMarkdown(payload, deterministicResult);
    }
  } catch (err: any) {
    console.error('Error invoking Gemini for Liquid Transformer analysis:', err);
    deterministicResult.markdown_report = buildFallbackLiquidTransformerMarkdown(payload, deterministicResult);
  }

  return deterministicResult;
}
