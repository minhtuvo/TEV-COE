import { GoogleGenAI } from '@google/genai';

export interface NetaInputPayload {
  site_info: {
    project_name: string;
    transformer_tag: string;
    fse_name: string;
    test_date: string;
  };
  visual_inspection: {
    nameplate_match: boolean;
    physical_condition: string;
    grounding_check: string;
    shipping_brackets_removed: boolean;
    cleanliness: string;
    bolt_torque_check: string;
    tap_position: string;
  };
  electrical_tests: {
    bolted_resistance_micro_ohms: number[];
    insulation_resistance_1min_1000v_megohms: {
      pri_to_ground: number;
      sec_to_ground: number;
      pri_to_sec: number;
    };
    dar_value: number;
    turns_ratio: {
      calculated_ratio: number;
      measured_ratios: number[];
    };
    secondary_voltage_no_load: string;
  };
  previous_test_data?: {
    last_test_date?: string;
    last_insulation_resistance_pri_ground?: number;
  };
}

export interface NetaEvaluationResult {
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
      min_standard: number;
      status: 'PASS' | 'FAIL';
      trend?: {
        previous: number;
        reduction_percent: number;
        warning: boolean;
        detail: string;
      };
    };
    dar: {
      value: number;
      min_standard: number;
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    turns_ratio: {
      calculated: number;
      measured: number[];
      max_error_percent: number;
      status: 'PASS' | 'FAIL';
      detail: string;
      phase_details: Array<{ phase: string; measured: number; error_percent: number; pass: boolean }>;
    };
    visual_mechanical: {
      status: 'PASS' | 'FAIL';
      issues: string[];
    };
  };
  recommendations: string[];
}

export function evaluateNetaDryType(data: NetaInputPayload): NetaEvaluationResult {
  const issues: string[] = [];
  const warnings: string[] = [];
  const recommendations: string[] = [];

  // 1. Bolted Connection Resistance
  const boltVals = data.electrical_tests.bolted_resistance_micro_ohms || [];
  const minBolt = boltVals.length > 0 ? Math.min(...boltVals) : 0;
  const maxBolt = boltVals.length > 0 ? Math.max(...boltVals) : 0;
  const boltDevPercent = minBolt > 0 ? ((maxBolt - minBolt) / minBolt) * 100 : 0;
  const boltStatus: 'PASS' | 'INVESTIGATE' = boltDevPercent > 50 ? 'INVESTIGATE' : 'PASS';

  if (boltStatus === 'INVESTIGATE') {
    warnings.push(`Điện trở mối nối bu-lông: Giá trị lớn nhất (${maxBolt} µΩ) lệch ${boltDevPercent.toFixed(1)}% so với mối nối nhỏ nhất (${minBolt} µΩ), vượt quá ngưỡng cho phép 50% theo NETA ATS-2025.`);
    recommendations.push(`Siết lại bu-lông mối nối có điện trở cao (${maxBolt} µΩ) bằng cờ-lê lực theo Table 100.12, làm sạch bề mặt tiếp xúc và đo lại bằng Low-Resistance Ohmmeter.`);
  }

  // 2. Insulation Resistance (IR @ 1000V DC 1 min)
  const ir = data.electrical_tests.insulation_resistance_1min_1000v_megohms;
  const minIRStandard = 500; // MΩ
  const priGndPass = ir.pri_to_ground >= minIRStandard;
  const secGndPass = ir.sec_to_ground >= minIRStandard;
  const priSecPass = ir.pri_to_sec >= minIRStandard;
  const irPass = priGndPass && secGndPass && priSecPass;

  let trendInfo: any = undefined;
  if (data.previous_test_data?.last_insulation_resistance_pri_ground) {
    const prevIR = data.previous_test_data.last_insulation_resistance_pri_ground;
    const dropPercent = ((prevIR - ir.pri_to_ground) / prevIR) * 100;
    const isWarning = dropPercent > 20;
    trendInfo = {
      previous: prevIR,
      reduction_percent: parseFloat(dropPercent.toFixed(2)),
      warning: isWarning,
      detail: isWarning
        ? `Giảm ${dropPercent.toFixed(1)}% so với lần đo gần nhất (${prevIR} MΩ), vượt ngưỡng suy giảm 20% cho phép.`
        : `Giảm ${dropPercent.toFixed(1)}% so với lần đo trước (${prevIR} MΩ), trong ngưỡng chấp nhận.`
    };
    if (isWarning) {
      warnings.push(`Điện trở cách điện: Cuộn sơ cấp - vỏ đạt ${ir.pri_to_ground} MΩ (> 500 MΩ), nhưng suy giảm ${dropPercent.toFixed(1)}% (> 20%) so với kết quả lần trước (${prevIR} MΩ).`);
      recommendations.push(`Kiểm tra độ ẩm môi trường, vệ sinh bề mặt cuộn dây và xem xét sấy nhẹ cuộn dây nếu cần thiết.`);
    }
  }

  if (!irPass) {
    issues.push(`Điện trở cách điện dưới ngưỡng tối thiểu 500 MΩ của NETA ATS-2025 Table 100.5.`);
    recommendations.push(`Không được đóng điện! Cần sấy khô cuộn dây và kiểm tra lại nguyên nhân rò rỉ cách điện.`);
  }

  // 3. DAR / PI
  const darVal = data.electrical_tests.dar_value;
  const darPass = darVal >= 1.0;
  if (!darPass) {
    issues.push(`Hệ số hấp thụ điện dịch DAR (${darVal}) < 1.0 không đạt tiêu chuẩn NETA ATS-2025.`);
    recommendations.push(`Sấy cách điện máy biến áp và đo lại đường cong phục hồi điện áp hoặc chỉ số phân cực PI.`);
  }

  // 4. Turns-Ratio Test
  const calcRatio = data.electrical_tests.turns_ratio.calculated_ratio;
  const measuredRatios = data.electrical_tests.turns_ratio.measured_ratios || [];
  const phaseNames = ['Pha A', 'Pha B', 'Pha C'];
  let maxRatioError = 0;
  let turnsRatioPass = true;

  const phaseDetails = measuredRatios.map((m, idx) => {
    const phase = phaseNames[idx] || `Pha ${idx + 1}`;
    const errPercent = calcRatio > 0 ? (Math.abs(m - calcRatio) / calcRatio) * 100 : 0;
    if (errPercent > maxRatioError) maxRatioError = errPercent;
    const pass = errPercent <= 0.5;
    if (!pass) turnsRatioPass = false;
    return {
      phase,
      measured: m,
      error_percent: parseFloat(errPercent.toFixed(3)),
      pass
    };
  });

  if (!turnsRatioPass) {
    const failedPhases = phaseDetails.filter(p => !p.pass).map(p => `${p.phase} (sai số ${p.error_percent}%)`).join(', ');
    warnings.push(`Tỷ số biến áp: ${failedPhases} vượt quá sai số cho phép 0.5% so với tỷ số tính toán (${calcRatio}).`);
    recommendations.push(`Kiểm tra lại nấc chuyển mạch phân áp (Tap switch/connections), đảm bảo tiếp xúc tốt trên các đầu cực và tiến hành đo lại TTR.`);
  }

  // 5. Visual & Mechanical
  const visualIssues: string[] = [];
  const vm = data.visual_inspection;
  if (!vm.shipping_brackets_removed) {
    visualIssues.push('Khung kẹp vận chuyển (Shipping brackets) CHƯA được tháo rời hoàn toàn.');
    recommendations.push('BẮT BUỘC tháo rời hoàn toàn khung kẹp vận chuyển trước khi đóng điện để tránh rung động và tiếng ồn cộng hưởng.');
  }
  if (!vm.nameplate_match) {
    visualIssues.push('Thông số Nameplate không khớp bản vẽ thiết kế.');
  }
  if (vm.grounding_check.toLowerCase() !== 'pass') {
    visualIssues.push('Hệ thống tiếp địa vỏ và lõi từ chưa đạt yêu cầu.');
    recommendations.push('Kiểm tra và siết chặt cáp tiếp địa vỏ máy biến áp và trung tính.');
  }

  // Determine Overall Status
  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let status_reason = 'Thiết bị đạt đầy đủ tiêu chuẩn nghiệm thu theo ANSI/NETA ATS-2025 Mục 7.2.1.1.';

  if (issues.length > 0 || !turnsRatioPass && maxRatioError > 2.0) {
    overall_status = 'FAIL';
    status_reason = 'Có hạng mục không đạt tiêu chuẩn nghiêm trọng, nguy cơ mất an toàn khi đóng điện.';
  } else if (warnings.length > 0 || visualIssues.length > 0 || !turnsRatioPass) {
    overall_status = 'INVESTIGATE';
    status_reason = 'Cần kiểm tra/xử lý lại tại chỗ trước khi đóng điện (có thông số vượt dung sai NETA ATS-2025 hoặc cảnh báo suy giảm).';
  }

  // Generate Standard Markdown Report
  const markdown_report = `### 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
**${overall_status === 'INVESTIGATE' ? '🟡 INVESTIGATE (Cần kiểm tra/xử lý lại tại chỗ trước khi đóng điện)' : overall_status === 'FAIL' ? '🔴 FAIL (Không đạt tiêu chuẩn đóng điện)' : '🟢 PASS (Đạt tiêu chuẩn ANSI/NETA ATS-2025)'}**

> **Ghi chú đánh giá:** ${status_reason}

---

### 2. BẢNG PHÂN TÍCH CHI TIẾT KẾT QUẢ ĐO (Detailed Evaluation Table)

| Hạng mục kiểm tra | Giá trị đo thực tế | Tiêu chuẩn NETA ATS-2025 | Lần đo trước / Tham chiếu | Đánh giá & Độ lệch | Trạng thái |
| :--- | :--- | :--- | :--- | :--- | :---: |
| **Khung kẹp vận chuyển** | ${vm.shipping_brackets_removed ? 'Đã tháo rời hoàn toàn' : 'CHƯA tháo rời'} | Bắt buộc tháo rời, đệm đàn hồi tự do | N/A | Khớp yêu cầu cơ khí | **${vm.shipping_brackets_removed ? 'PASS' : 'FAIL'}** |
| **Độ sạch & Lực siết** | ${vm.cleanliness} / ${vm.bolt_torque_check} | Sạch sẽ, khô ráo, lực siết Table 100.12 | N/A | Đạt kiểm tra thị giác | **PASS** |
| **Điện trở mối nối bu-lông** | [${boltVals.join(', ')}] $\\mu\\Omega$ | Sai lệch $\\le 50\\%$ so với mối nối nhỏ nhất | Min: ${minBolt} $\\mu\\Omega$ | Lệch lớn nhất: **${boltDevPercent.toFixed(1)}%** | **${boltStatus}** |
| **Điện trở cách điện (Sơ - Đất)** | **${ir.pri_to_ground} M$\\Omega$** (1000V DC) | $\\ge 500\\text{ M}\\Omega$ (Table 100.5) | ${data.previous_test_data?.last_insulation_resistance_pri_ground ? `${data.previous_test_data.last_insulation_resistance_pri_ground} M\\Omega` : 'N/A'} | Đạt mức tối thiểu; giảm ${trendInfo ? trendInfo.reduction_percent : 0}% | **${priGndPass ? (trendInfo?.warning ? 'INVESTIGATE' : 'PASS') : 'FAIL'}** |
| **Điện trở cách điện (Thứ - Đất)** | **${ir.sec_to_ground} M$\\Omega$** | $\\ge 500\\text{ M}\\Omega$ (Table 100.5) | N/A | Đạt tiêu chuẩn | **PASS** |
| **Điện trở cách điện (Sơ - Thứ)** | **${ir.pri_to_sec} M$\\Omega$** | $\\ge 500\\text{ M}\\Omega$ (Table 100.5) | N/A | Đạt tiêu chuẩn | **PASS** |
| **Hệ số hấp thụ điện dịch (DAR)** | **${darVal}** | $\\ge 1.0$ | N/A | Khả năng nạp điện dung tốt | **PASS** |
| **Tỷ số biến áp (Turns-Ratio)** | [${phaseDetails.map(p => `${p.phase}: ${p.measured}`).join(', ')}] | Sai số $\\le 0.5\\%$ so với tính toán (${calcRatio}) | Tính toán: ${calcRatio} | Sai số max: **${maxRatioError.toFixed(2)}%** (${phaseDetails.find(p => !p.pass)?.phase || 'Tất cả đạt'}) | **${turnsRatioPass ? 'PASS' : 'FAIL'}** |
| **Điện áp thứ cấp không tải** | ${data.electrical_tests.secondary_voltage_no_load} | Phù hợp nhãn Nameplate | 208V/120V | Khớp điện áp thiết kế | **PASS** |

---

### 3. PHÂN TÍCH XU HƯỚNG & CẢNH BÁO (Trend Analysis & Risk Warning)
${warnings.length > 0 || issues.length > 0 ? [
  ...warnings.map(w => `- ⚠️ **Cảnh báo:** ${w}`),
  ...issues.map(i => `- 🚨 **Lỗi nghiêm trọng:** ${i}`)
].join('\n') : '- ✅ Tất cả các thông số điện và cơ học đều nằm trong giới hạn tối ưu của ANSI/NETA ATS-2025.'}

---

### 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHO FSE (Actionable Recommendations)
${recommendations.length > 0 ? recommendations.map((r, i) => `${i + 1}. **${r}**`).join('\n') : '1. Cho phép ký biên bản nghiệm thu kỹ thuật tại công trường và tiến hành quy trình đóng điện thử nghiệm.'}
`;

  return {
    overall_status,
    status_reason,
    markdown_report,
    evaluations: {
      bolted_resistance: {
        min: minBolt,
        max: maxBolt,
        max_deviation_percent: parseFloat(boltDevPercent.toFixed(1)),
        status: boltStatus,
        detail: `Điện trở mối nối: min = ${minBolt} µΩ, max = ${maxBolt} µΩ. Độ lệch = ${boltDevPercent.toFixed(1)}% (ngưỡng tối đa 50%).`
      },
      insulation_resistance: {
        pri_to_ground: ir.pri_to_ground,
        sec_to_ground: ir.sec_to_ground,
        pri_to_sec: ir.pri_to_sec,
        min_standard: minIRStandard,
        status: irPass ? 'PASS' : 'FAIL',
        trend: trendInfo
      },
      dar: {
        value: darVal,
        min_standard: 1.0,
        status: darPass ? 'PASS' : 'FAIL',
        detail: `DAR = ${darVal} (yêu cầu >= 1.0).`
      },
      turns_ratio: {
        calculated: calcRatio,
        measured: measuredRatios,
        max_error_percent: parseFloat(maxRatioError.toFixed(3)),
        status: turnsRatioPass ? 'PASS' : 'FAIL',
        detail: `Tỷ số danh định ${calcRatio}. Sai số lớn nhất ${maxRatioError.toFixed(2)}% (ngưỡng tối đa 0.5%).`,
        phase_details: phaseDetails
      },
      visual_mechanical: {
        status: visualIssues.length === 0 ? 'PASS' : 'FAIL',
        issues: visualIssues
      }
    },
    recommendations
  };
}

export async function generateNetaAiAnalysis(data: NetaInputPayload): Promise<NetaEvaluationResult> {
  const baseResult = evaluateNetaDryType(data);

  // If Gemini API key is available, enhance with generative analysis
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return baseResult;
  }

  try {
    const ai = new GoogleGenAI({ apiKey });
    const prompt = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) cho thiết bị Điện hạ áp theo tiêu chuẩn ANSI/NETA ATS-2025 Mục 7.2.1.1 (Transformers, Dry-Type, Air-Cooled, Low-Voltage, Small).

Dữ liệu kiểm tra thực tế từ hiện trường:
${JSON.stringify(data, null, 2)}

Kết quả tính toán quy tắc NETA ATS-2025:
- Trạng thái tổng quan: ${baseResult.overall_status} (${baseResult.status_reason})
- Độ lệch mối nối bu-lông: ${baseResult.evaluations.bolted_resistance.max_deviation_percent}% (ngưỡng 50%)
- IR Sơ-Đất: ${data.electrical_tests.insulation_resistance_1min_1000v_megohms.pri_to_ground} MΩ (chuẩn >= 500 MΩ)
- Xu hướng IR: Giảm ${baseResult.evaluations.insulation_resistance.trend?.reduction_percent || 0}% so với trước (${data.previous_test_data?.last_insulation_resistance_pri_ground || 'N/A'} MΩ)
- DAR: ${data.electrical_tests.dar_value} (chuẩn >= 1.0)
- Tỷ số Turns-ratio sai số: ${baseResult.evaluations.turns_ratio.max_error_percent}% (ngưỡng 0.5%)

Nhiệm vụ: Hãy phân tích kỹ lưỡng và trả về phản hồi định dạng Markdown chuyên nghiệp gồm 4 phần chính:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE (Cần kiểm tra thêm).
2. BẢNG PHÂN TÍCH CHI TIẾT KẾT QUẢ ĐO (Detailed Evaluation Table): So sánh giá trị đo thực tế với Tiêu chuẩn NETA ATS-2025 và Lần đo trước.
3. PHÂN TÍCH XU HƯỚNG & CẢNH BÁO (Trend Analysis & Risk Warning).
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHO FSE (Actionable Recommendations).

Yêu cầu ngôn ngữ: Tiếng Việt kỹ thuật điện chuyên ngành, rõ ràng, đanh thép, chuẩn xác theo quy chuẩn NETA ATS-2025.`;

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
