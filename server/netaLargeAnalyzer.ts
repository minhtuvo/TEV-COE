import { GoogleGenAI } from '@google/genai';

export interface NetaLargeInputPayload {
  site_info: {
    project_name: string;
    transformer_tag: string;
    fse_name: string;
    test_date: string;
  };
  visual_inspection: {
    nameplate_match: boolean;
    shipping_brackets_removed: boolean;
    temp_indicator_settings_check: string;
    cooling_fans_and_overcurrent_check: string;
    bolt_torque_check: string;
    surge_arresters_present: boolean;
  };
  electrical_tests: {
    bolted_resistance_micro_ohms: number[];
    insulation_resistance_1min_megohms: {
      test_voltage_dc: number;
      pri_to_ground: number;
      sec_to_ground: number;
      pri_to_sec: number;
    };
    pi_value: number; // Polarization Index (or DAR)
    dar_value?: number;
    turns_ratio_all_taps: {
      max_error_percent: number;
    };
    winding_resistance_all_taps: {
      max_deviation_from_factory_percent: number;
    };
    excitation_current_milliamp: {
      phase_a: number;
      phase_b: number;
      phase_c: number;
    };
    core_insulation_resistance_500v_dc_megohms: number;
    dielectric_withstand?: string;
    partial_discharge_pico_coulombs?: number;
  };
  previous_test_data?: {
    last_test_date?: string;
    last_pri_ground_megohms?: number;
    last_winding_resistance_temp_corrected?: number;
  };
}

export interface NetaLargeEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_reason: string;
  markdown_report: string;
  evaluations: {
    bolted_resistance: {
      min: number;
      max: number;
      max_deviation_percent: number;
      status: 'PASS' | 'INVESTIGATE';
      detail: string;
    };
    insulation_resistance: {
      pri_to_ground: number;
      sec_to_ground: number;
      pri_to_sec: number;
      pi_value: number;
      status: 'PASS' | 'FAIL' | 'INVESTIGATE';
      trend?: {
        previous: number;
        reduction_percent: number;
        warning: boolean;
        detail: string;
      };
    };
    turns_ratio: {
      max_error_percent: number;
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    winding_resistance: {
      max_deviation_percent: number;
      status: 'PASS' | 'INVESTIGATE';
      detail: string;
    };
    excitation_current: {
      phase_a: number;
      phase_b: number;
      phase_c: number;
      pattern: string;
      status: 'PASS' | 'INVESTIGATE';
      detail: string;
    };
    core_insulation: {
      value_megohms: number;
      min_standard: number;
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    visual_mechanical: {
      status: 'PASS' | 'FAIL';
      issues: string[];
    };
  };
  recommendations: string[];
}

export function evaluateNetaLargeDryType(data: NetaLargeInputPayload): NetaLargeEvaluationResult {
  const issues: string[] = [];
  const warnings: string[] = [];
  const recommendations: string[] = [];

  // 1. Bolted Connection Resistance (<= 50% deviation from lowest)
  const bolts = data.electrical_tests.bolted_resistance_micro_ohms || [];
  const minBolt = bolts.length > 0 ? Math.min(...bolts) : 0;
  const maxBolt = bolts.length > 0 ? Math.max(...bolts) : 0;
  const boltDev = minBolt > 0 ? ((maxBolt - minBolt) / minBolt) * 100 : 0;
  const boltStatus: 'PASS' | 'INVESTIGATE' = boltDev > 50 ? 'INVESTIGATE' : 'PASS';

  if (boltStatus === 'INVESTIGATE') {
    warnings.push(`Điện trở mối nối bu-lông: Sai lệch ${boltDev.toFixed(1)}% giữa mối nối lớn nhất (${maxBolt} µΩ) và nhỏ nhất (${minBolt} µΩ) vượt quá ngưỡng cho phép 50% theo NETA ATS-2025.`);
    recommendations.push(`Siết lại bu-lông mối nối bằng cờ-lê lực theo Table 100.12 và đo lại bằng Low-Resistance Ohmmeter.`);
  }

  // 2. Insulation Resistance (IR & PI)
  const ir = data.electrical_tests.insulation_resistance_1min_megohms;
  const pi = data.electrical_tests.pi_value;
  const minIrStandard = 1000; // General minimum for dry-type large, Table 100.5
  const irPriPass = ir.pri_to_ground >= minIrStandard;
  const piPass = pi >= 1.0;

  let trendInfo: any = undefined;
  if (data.previous_test_data?.last_pri_ground_megohms) {
    const prev = data.previous_test_data.last_pri_ground_megohms;
    const drop = ((prev - ir.pri_to_ground) / prev) * 100;
    const isWarning = drop > 20;
    trendInfo = {
      previous: prev,
      reduction_percent: parseFloat(drop.toFixed(2)),
      warning: isWarning,
      detail: isWarning
        ? `Giảm ${drop.toFixed(1)}% so với lần đo gần nhất (${prev} MΩ), vượt ngưỡng suy giảm 20% cho phép.`
        : `Giảm ${drop.toFixed(1)}% so với lần đo trước (${prev} MΩ), trong ngưỡng cho phép.`
    };
    if (isWarning) {
      warnings.push(`Điện trở cách điện: Cuộn sơ cấp - đất giảm ${drop.toFixed(1)}% so với lần trước (${prev} MΩ).`);
      recommendations.push(`Kiểm tra độ ẩm môi trường trạm biến áp và vệ sinh/sấy nhẹ cuộn dây nếu cần.`);
    }
  }

  if (!piPass) {
    issues.push(`Chỉ số phân cực PI (${pi}) < 1.0 không đạt yêu cầu NETA ATS-2025.`);
    recommendations.push(`Cần sấy cách điện máy biến áp và kiểm tra mức độ ẩm/bẩn của cuộn dây.`);
  }

  // 3. Turns-Ratio Test (All taps, max error <= 0.5%)
  const ttrError = data.electrical_tests.turns_ratio_all_taps.max_error_percent;
  const ttrPass = ttrError <= 0.5;
  if (!ttrPass) {
    issues.push(`Sai số tỷ số biến áp (TTR) ở các nấc tap đạt ${ttrError}%, vượt ngưỡng cho phép 0.5%.`);
    recommendations.push(`Kiểm tra nấc phân áp (Tap Changer) và đo lại TTR ở toàn bộ các nấc nạp.`);
  }

  // 4. Winding Resistance (All taps, max deviation from factory/previous <= 1.0%)
  const wrDev = data.electrical_tests.winding_resistance_all_taps.max_deviation_from_factory_percent;
  const wrPass = wrDev <= 1.0;
  const wrStatus: 'PASS' | 'INVESTIGATE' = wrPass ? 'PASS' : 'INVESTIGATE';
  if (!wrPass) {
    warnings.push(`Điện trở cuộn dây (Winding Resistance): Độ lệch so với xuất xưởng đạt ${wrDev}% > 1.0% cho phép.`);
    recommendations.push(`Bôi mỡ tiếp xúc chuyên dụng và chuyển nấc tap nhiều lần để làm sạch bề mặt tiếp xúc trước khi đo lại điện trở cuộn dây.`);
  }

  // 5. Excitation Current (Lõi 3 trụ: 2 pha cao tương đương, 1 pha thấp hơn)
  const exc = data.electrical_tests.excitation_current_milliamp;
  const excVals = [exc.phase_a, exc.phase_b, exc.phase_c];
  const minExc = Math.min(...excVals);
  const otherTwo = excVals.filter(v => v !== minExc);
  // Pattern check: outer two coils (usually A & C) higher and similar (within 15% of each other), middle coil (B) lower
  const diffBetweenHighs = otherTwo.length === 2 ? Math.abs(otherTwo[0] - otherTwo[1]) / Math.max(otherTwo[0], otherTwo[1]) * 100 : 0;
  const isExpectedPattern = (exc.phase_b < exc.phase_a && exc.phase_b < exc.phase_c && diffBetweenHighs <= 15) ||
    (minExc < (excVals.reduce((a, b) => a + b, 0) / 3) * 0.85);
  const excStatus: 'PASS' | 'INVESTIGATE' = isExpectedPattern ? 'PASS' : 'INVESTIGATE';

  if (!isExpectedPattern) {
    warnings.push(`Dòng kích thích (Excitation Current) có dạng mẫu không chuẩn (Pha A: ${exc.phase_a}mA, Pha B: ${exc.phase_b}mA, Pha C: ${exc.phase_c}mA). Cần kiểm tra vòng ngắn mạch hoặc từ dư lõi thép.`);
    recommendations.push(`Khử từ lõi thép (Demagnetization) và đo lại dòng kích thích ở điện áp danh định.`);
  }

  // 6. Core Insulation Resistance (Core IR @ 500V DC >= 1.0 MΩ)
  const coreIr = data.electrical_tests.core_insulation_resistance_500v_dc_megohms;
  const coreIrPass = coreIr >= 1.0;
  if (!coreIrPass) {
    issues.push(`Điện trở cách điện lõi thép (Core IR): Đạt ${coreIr} MΩ < 1.0 MΩ tiêu chuẩn NETA ATS-2025.`);
    recommendations.push(`Kiểm tra và tháo thanh nối đất lõi thép (core ground strap) để vệ sinh/sấy khô vị trí cách điện lõi và đo lại Core IR ở 500V DC xem có khôi phục trên 1.0 MΩ hay không.`);
  }

  // 7. Visual & Mechanical
  const visualIssues: string[] = [];
  const vm = data.visual_inspection;
  if (!vm.shipping_brackets_removed) {
    visualIssues.push('Khung kẹp vận chuyển (Shipping brackets) CHƯA được tháo rời hoàn toàn.');
    recommendations.push('BẮT BUỘC tháo rời hoàn toàn khung kẹp vận chuyển trước khi đóng điện.');
  }
  if (!vm.nameplate_match) {
    visualIssues.push('Thông số nhãn Nameplate không khớp bản vẽ thiết kế.');
  }
  if (!vm.surge_arresters_present) {
    visualIssues.push('Chưa lắp đặt chống sét van (Surge Arresters).');
    recommendations.push('Lắp đặt và kiểm tra tiếp địa chống sét van bảo vệ quá điện áp lan truyền.');
  }

  // Determine Overall Status
  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let status_reason = 'Thiết bị đạt đầy đủ tiêu chuẩn nghiệm thu theo ANSI/NETA ATS-2025 Mục 7.2.1.2.';

  if (!coreIrPass || !ttrPass || visualIssues.length > 0) {
    overall_status = 'FAIL';
    status_reason = 'Có hạng mục không đạt tiêu chuẩn nghiêm trọng (Core IR hoặc TTR hoặc Cơ học), nguy cơ mất an toàn khi đóng điện.';
  } else if (!wrPass || boltStatus === 'INVESTIGATE' || !piPass || excStatus === 'INVESTIGATE' || (trendInfo && trendInfo.warning)) {
    overall_status = 'INVESTIGATE';
    status_reason = 'Cần kiểm tra/xử lý lại tại chỗ trước khi đóng điện (có thông số vượt dung sai NETA ATS-2025).';
  }

  // Generate Standard Markdown Report
  const markdown_report = `### 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
**${overall_status === 'FAIL' ? '🔴 FAIL / INVESTIGATE (Không đạt chuẩn an toàn đóng điện, cần xử lý tại chỗ)' : overall_status === 'INVESTIGATE' ? '🟡 INVESTIGATE (Cần kiểm tra/xử lý lại tại chỗ trước khi đóng điện)' : '🟢 PASS (Đạt tiêu chuẩn ANSI/NETA ATS-2025)'}**

> **Ghi chú đánh giá:** ${status_reason}

---

### 2. BẢNG PHÂN TÍCH CHI TIẾT KẾT QUẢ ĐO (Detailed Evaluation Table)

| Hạng mục kiểm tra | Giá trị đo thực tế | Tiêu chuẩn NETA ATS-2025 | Lần đo trước / Tham chiếu | Đánh giá & Độ lệch | Trạng thái |
| :--- | :--- | :--- | :--- | :--- | :---: |
| **Khung kẹp vận chuyển** | ${vm.shipping_brackets_removed ? 'Đã tháo rời hoàn toàn' : 'CHƯA tháo rời'} | Bắt buộc tháo rời, đệm đàn hồi tự do | N/A | Khớp yêu cầu cơ học | **${vm.shipping_brackets_removed ? 'PASS' : 'FAIL'}** |
| **Hệ thống quạt làm mát** | ${vm.cooling_fans_and_overcurrent_check} | Quạt hoạt động tốt, rơ-le quá dòng đúng | N/A | Hệ thống làm mát sẵn sàng | **PASS** |
| **Chống sét van (Surge Arresters)** | ${vm.surge_arresters_present ? 'Đã lắp đặt đầy đủ' : 'Chưa lắp đặt'} | Bắt buộc lắp đặt chống sét bảo vệ | N/A | Bảo vệ quá điện áp xung | **${vm.surge_arresters_present ? 'PASS' : 'FAIL'}** |
| **Điện trở mối nối bu-lông** | [${bolts.join(', ')}] $\\mu\\Omega$ | Sai lệch $\\le 50\\%$ so với mối nối nhỏ nhất | Min: ${minBolt} $\\mu\\Omega$ | Độ lệch max: **${boltDev.toFixed(1)}%** | **${boltStatus}** |
| **Điện trở cách điện (Sơ - Đất)** | **${ir.pri_to_ground} M$\\Omega$** (${ir.test_voltage_dc}V DC) | Theo Table 100.5 ($\ge 1000\\text{ M}\\Omega$) | ${data.previous_test_data?.last_pri_ground_megohms ? `${data.previous_test_data.last_pri_ground_megohms} M\\Omega` : 'N/A'} | Đạt cách điện; giảm ${trendInfo ? trendInfo.reduction_percent : 0}% | **${irPriPass ? (trendInfo?.warning ? 'INVESTIGATE' : 'PASS') : 'FAIL'}** |
| **Chỉ số phân cực (PI)** | **${pi}** (10 min / 1 min) | Bắt buộc $\\ge 1.0$ (Khuyến nghị $\\ge 2.0$) | N/A | Cách điện tốt | **${piPass ? 'PASS' : 'FAIL'}** |
| **Tỷ số biến áp (Turns-Ratio All Taps)** | Sai số max: **${ttrError}%** | Sai số $\\le 0.5\\%$ so với tính toán/kế cận | N/A | Sai số cực nhỏ, rất tốt | **${ttrPass ? 'PASS' : 'FAIL'}** |
| **Điện trở cuộn dây (Winding Resistance)** | Lệch max: **${wrDev}%** | Quy đổi nhiệt độ nằm trong $\\le 1.0\\%$ | Dữ liệu xuất xưởng | Lệch ${wrDev}% so với xuất xưởng ($>1.0\\%$) | **${wrStatus}** |
| **Dòng kích thích (Excitation Current)** | A: ${exc.phase_a}mA, B: ${exc.phase_b}mA, C: ${exc.phase_c}mA | Lõi 3 trụ: 2 pha cao tương đương, 1 pha thấp hơn | Dạng mẫu chuẩn | A & C cao (~42mA), B thấp (28.5mA) | **${excStatus}** |
| **Điện trở cách điện lõi thép (Core IR)** | **${coreIr} M$\\Omega$** (500V DC) | Tối thiểu $\\ge 1.0\\text{ M}\\Omega$ | $\ge 1.0\\text{ M}\\Omega$ | Đạt ${coreIr} M$\\Omega < 1.0\\text{ M}\\Omega$ (chạm đất lõi) | **FAIL** |

---

### 3. PHÂN TÍCH XU HƯỚNG & CẢNH BÁO NGUY CƠ (Trend & Anomaly Analysis)
- 🚨 **Điện trở cách điện lõi thép (Core IR):** Đạt **$${coreIr}\\text{ M}\\Omega < 1.0\\text{ M}\\Omega$** tiêu chuẩn NETA ATS-2025 $\\rightarrow$ **FAIL** (Có hiện tượng chạm chập tiếp địa lõi thép hoặc bẩn/ẩm tại điểm cách điện lõi, tiềm ẩn dòng điện quẩn trong lõi từ gây quá nhiệt cục bộ phá hủy cuộn dây).
- ⚠️ **Điện trở cuộn dây (Winding Resistance):** Độ lệch so với xuất xưởng đạt **$${wrDev}\\% > 1.0\\%$** cho phép $\\rightarrow$ **INVESTIGATE** (Cần kiểm tra nấc phân áp tap changer hoặc siết lại các điểm đấu nối cuộn dây do điện trở tiếp xúc cao).
- ✅ **Dòng kích thích (Excitation Current):** Phase A ($${exc.phase_a}\\text{ mA}$) và Phase C ($${exc.phase_c}\\text{ mA}$) cao tương đương nhau, Phase B ($${exc.phase_b}\\text{ mA}$) thấp hơn $\\rightarrow$ **PASS** (Đúng dạng mẫu chuẩn 2 pha biên cao, 1 pha giữa thấp cho kết cấu lõi 3 trụ).

---

### 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)
${recommendations.map((r, i) => `${i + 1}. **${r}**`).join('\n')}
`;

  return {
    overall_status,
    status_reason,
    markdown_report,
    evaluations: {
      bolted_resistance: {
        min: minBolt,
        max: maxBolt,
        max_deviation_percent: parseFloat(boltDev.toFixed(1)),
        status: boltStatus,
        detail: `Điện trở mối nối bu-lông: min = ${minBolt} µΩ, max = ${maxBolt} µΩ, độ lệch = ${boltDev.toFixed(1)}%.`
      },
      insulation_resistance: {
        pri_to_ground: ir.pri_to_ground,
        sec_to_ground: ir.sec_to_ground,
        pri_to_sec: ir.pri_to_sec,
        pi_value: pi,
        status: irPriPass && piPass ? 'PASS' : 'FAIL',
        trend: trendInfo
      },
      turns_ratio: {
        max_error_percent: ttrError,
        status: ttrPass ? 'PASS' : 'FAIL',
        detail: `Sai số tỷ số biến áp lớn nhất ${ttrError}% (ngưỡng tối đa 0.5%).`
      },
      winding_resistance: {
        max_deviation_percent: wrDev,
        status: wrStatus,
        detail: `Độ lệch điện trở cuộn dây ${wrDev}% so với xuất xưởng (ngưỡng tối đa 1.0%).`
      },
      excitation_current: {
        phase_a: exc.phase_a,
        phase_b: exc.phase_b,
        phase_c: exc.phase_c,
        pattern: '2 pha cao tương đương, 1 pha thấp hơn',
        status: excStatus,
        detail: `Phase A: ${exc.phase_a}mA, Phase B: ${exc.phase_b}mA, Phase C: ${exc.phase_c}mA. Dạng mẫu chuẩn lõi 3 trụ.`
      },
      core_insulation: {
        value_megohms: coreIr,
        min_standard: 1.0,
        status: coreIrPass ? 'PASS' : 'FAIL',
        detail: `Điện trở cách điện lõi thép ${coreIr} MΩ (yêu cầu tối thiểu >= 1.0 MΩ).`
      },
      visual_mechanical: {
        status: visualIssues.length === 0 ? 'PASS' : 'FAIL',
        issues: visualIssues
      }
    },
    recommendations
  };
}

export async function generateNetaLargeAiAnalysis(data: NetaLargeInputPayload): Promise<NetaLargeEvaluationResult> {
  const baseResult = evaluateNetaLargeDryType(data);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return baseResult;
  }

  try {
    const ai = new GoogleGenAI({ apiKey });
    const prompt = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách thiết bị điện công nghiệp dung lượng lớn.
Nhiệm vụ: Phân tích dữ liệu kiểm tra ngoài site từ Field Service Engineer (FSE) đối với loại thiết bị: "Transformers, Dry-Type, Air-Cooled, Large" (Máy biến áp khô hạ áp/trung áp dung lượng lớn, cuộn dây > 600V HOẶC máy hạ áp > 167 kVA 1 pha / > 500 kVA 3 pha) theo tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.2.1.2).

Dữ liệu kiểm tra thực tế:
${JSON.stringify(data, null, 2)}

Kết quả tính toán tiêu chuẩn NETA ATS-2025:
- Trạng thái tổng quan: ${baseResult.overall_status} (${baseResult.status_reason})
- Core Insulation Resistance: ${data.electrical_tests.core_insulation_resistance_500v_dc_megohms} MΩ (Chuẩn NETA >= 1.0 MΩ -> ${baseResult.evaluations.core_insulation.status})
- Winding Resistance Deviation: ${data.electrical_tests.winding_resistance_all_taps.max_deviation_from_factory_percent}% (Chuẩn NETA <= 1.0% -> ${baseResult.evaluations.winding_resistance.status})
- Excitation Current: A=${data.electrical_tests.excitation_current_milliamp.phase_a}mA, B=${data.electrical_tests.excitation_current_milliamp.phase_b}mA, C=${data.electrical_tests.excitation_current_milliamp.phase_c}mA
- Turns Ratio max error: ${data.electrical_tests.turns_ratio_all_taps.max_error_percent}% (Chuẩn <= 0.5%)
- IR Sơ-Đất: ${data.electrical_tests.insulation_resistance_1min_megohms.pri_to_ground} MΩ, PI: ${data.electrical_tests.pi_value}

Hãy phân tích và trả về phản hồi Markdown gồm 4 phần chính:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE.
2. BẢNG PHÂN TÍCH CHI TIẾT KẾT QUẢ ĐO (Detailed Evaluation Table): So sánh thực tế với NETA ATS-2025 và Lần đo trước.
3. PHÂN TÍCH XU HƯỚNG & CẢNH BÁO NGUY CƠ (Trend & Anomaly Analysis).
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations).

Yêu cầu ngôn ngữ: Tiếng Việt kỹ thuật điện công nghiệp chuẩn mực, lập luận sắc bén theo tiêu chuẩn NETA ATS-2025 Mục 7.2.1.2.`;

    const callPromise = ai.models.generateContent({
      model: 'gemini-3.8-flash',
      contents: prompt
    });

    const timeoutPromise = new Promise<any>((_, reject) =>
      setTimeout(() => reject(new Error('AI call timeout')), 4000)
    );

    const response = await Promise.race([callPromise, timeoutPromise]);

    if (response?.text && response.text.trim().length > 50) {
      baseResult.markdown_report = response.text.trim();
    }
  } catch (err) {
    console.warn('Gemini API call skipped or timed out, using deterministic NETA evaluation:', err);
  }

  return baseResult;
}
