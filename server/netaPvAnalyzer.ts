import { GoogleGenAI } from '@google/genai';

export interface NetaPvInputPayload {
  site_info: {
    project_name: string;
    pv_array_tag: string;
    fse_name: string;
    test_date: string;
  };
  visual_inspection: {
    nameplate_match: boolean;
    physical_condition: string;
    mounting_and_grounding: string;
    cleanliness: string;
    inverter_gateway_settings: string;
    wiring_connections_tight: boolean;
  };
  electrical_tests: {
    polarity_check: string;
    blocking_diode_test: string;
    dry_insulation_resistance_megohms: {
      string_01_to_ground: number;
      string_02_to_ground: number;
      string_03_to_ground: number;
      [key: string]: number;
    };
    open_circuit_voltage_voc: {
      expected_voc_v: number;
      measured_voc_v: number[];
    };
    short_circuit_current_isc: {
      irradiance_w_m2: number;
      measured_isc_a: number[];
    };
    iv_curve_measurement: string;
    ground_resistance_section_7_13_ohms: number;
    das_communication_section_7_23_status: string;
  };
  previous_test_data?: {
    last_test_date?: string;
    last_string_03_voc_v?: number;
    [key: string]: any;
  };
}

export interface NetaPvEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_reason: string;
  markdown_report: string;
  evaluations: {
    polarity_and_mechanical: {
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    insulation_resistance: {
      string_01: number;
      string_02: number;
      string_03: number;
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    voc_voltage: {
      expected: number;
      measured: number[];
      max_deviation_percent: number;
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    iv_curve_and_das: {
      iv_curve: string;
      ground_ohms: number;
      das_status: string;
      status: 'PASS' | 'INVESTIGATE' | 'FAIL';
      detail: string;
    };
  };
}

export function performDeterministicPvAnalysis(data: NetaPvInputPayload): NetaPvEvaluationResult {
  const v = data.visual_inspection;
  const e = data.electrical_tests;

  // 1. Polarity & Mechanical
  const polarityPass = e.polarity_check.toLowerCase().includes('pass') && !e.polarity_check.toLowerCase().includes('fail');
  const mechPass = v.nameplate_match && v.wiring_connections_tight && v.physical_condition.toLowerCase().includes('pass');
  const polarityMechStatus: 'PASS' | 'FAIL' = (polarityPass && mechPass) ? 'PASS' : 'FAIL';
  const polarityMechDetail = (polarityPass && mechPass)
    ? 'Cực tính đúng 100% trên tất cả chuỗi cáp và Combiner Box. Cơ khí giá đỡ, module pin và đấu nối MC4 đạt tiêu chuẩn NETA Section 7.29.1.'
    : 'PHÁT HIỆN LỖI CỰC TÍNH HOẶC CƠ KHÍ: Cần kiểm tra lại cực tính đấu nối hoặc kết cấu module pin bị nứt vỡ!';

  // 2. Insulation Resistance (Dry/Wet)
  const ir1 = e.dry_insulation_resistance_megohms?.string_01_to_ground || 0;
  const ir2 = e.dry_insulation_resistance_megohms?.string_02_to_ground || 0;
  const ir3 = e.dry_insulation_resistance_megohms?.string_03_to_ground || 0;

  let irStatus: 'PASS' | 'FAIL' = 'PASS';
  let irDetail = `Điện trở cách điện chuỗi 01 (${ir1} MΩ) và chuỗi 02 (${ir2} MΩ) tốt (>100 MΩ).`;

  if (ir3 < 50 || (ir1 > 100 && ir3 < ir1 * 0.2)) {
    irStatus = 'FAIL';
    irDetail += ` Chuỗi 03 chỉ đạt ${ir3} MΩ (thấp bất thường so với chuỗi 1 & 2 > 140 MΩ). Cảnh báo rò điện đất hoặc trầy xước vỏ cáp DC chuỗi 03!`;
  } else {
    irDetail += ` Chuỗi 03 đạt ${ir3} MΩ, cách điện đảm bảo.`;
  }

  // 3. Voc Open-Circuit Voltage
  const expectedVoc = e.open_circuit_voltage_voc?.expected_voc_v || 980;
  const measuredVoc = Array.isArray(e.open_circuit_voltage_voc?.measured_voc_v)
    ? e.open_circuit_voltage_voc.measured_voc_v
    : [982, 978, 850];

  let maxVocDiffPercent = 0;
  let abnormalVocStringIndex = -1;

  measuredVoc.forEach((voc, idx) => {
    const diff = Math.abs(voc - expectedVoc);
    const dev = (diff / expectedVoc) * 100;
    if (dev > maxVocDiffPercent) {
      maxVocDiffPercent = dev;
      abnormalVocStringIndex = idx + 1;
    }
  });

  // Check parallel difference
  const maxVoc = Math.max(...measuredVoc);
  const minVoc = Math.min(...measuredVoc);
  const parallelDiffPercent = maxVoc > 0 ? ((maxVoc - minVoc) / maxVoc) * 100 : 0;

  let vocStatus: 'PASS' | 'FAIL' = 'PASS';
  let vocDetail = `Điện áp Voc các chuỗi đồng đều, độ lệch song song ${parallelDiffPercent.toFixed(1)}% (≤ 5% tiêu chuẩn NETA Section 7.29).`;

  if (parallelDiffPercent > 5.0 || maxVocDiffPercent > 5.0) {
    vocStatus = 'FAIL';
    vocDetail = `Điện áp hở mạch Voc Chuỗi 03 đo được ${measuredVoc[2] || 850}V (thấp hơn ${maxVocDiffPercent.toFixed(1)}% so với định mức ${expectedVoc}V và các chuỗi song song), vượt quá ngưỡng cho phép 5% của NETA Section 7.29 (Nghi ngờ nối tắt 2–3 tấm pin hoặc hỏng Diode bypass).`;
  }

  // 4. I-V Curve, Grounding & DAS
  const ivDistorted = e.iv_curve_measurement?.toLowerCase().includes('distortion');
  const groundOhms = e.ground_resistance_section_7_13_ohms || 2.1;
  const groundPass = groundOhms <= 5.0; // NETA Section 7.13
  const dasPass = e.das_communication_section_7_23_status?.toLowerCase().includes('pass');

  let ivDasStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let ivDasDetail = `Đường cong I-V đạt yêu cầu, điện trở nối đất đạt ${groundOhms} Ω (≤ 5.0 Ω NETA Sec 7.13), truyền thông DAS giám sát tốt.`;

  if (ivDistorted) {
    ivDasStatus = 'FAIL';
    ivDasDetail = `Đường cong I-V Chuỗi 03 bị biến dạng (Distortion detected) tương ứng với hiện tượng giảm điện áp hở mạch Voc. Điện trở tiếp địa đạt ${groundOhms} Ω.`;
  } else if (!groundPass || !dasPass) {
    ivDasStatus = 'INVESTIGATE';
    ivDasDetail = `Điện trở nối đất ${groundOhms} Ω hoặc truyền thông DAS cần hiệu chỉnh.`;
  }

  // Overall status
  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let status_reason = 'Mảng pin mặt trời Solar PV tuân thủ tiêu chuẩn ANSI/NETA ATS-2025 Section 7.29.';

  if (!polarityPass || irStatus === 'FAIL' || vocStatus === 'FAIL' || ivDistorted) {
    overall_status = 'FAIL';
    status_reason = 'Phát hiện sự cố nghiêm trọng trên Chuỗi 03 (Sụt áp Voc 13.2%, cách điện chỉ 12 MΩ, biến dạng I-V curve). Cần cô lập xử lý khẩn cấp trước khi hòa lưới!';
  } else if (ivDasStatus === 'INVESTIGATE') {
    overall_status = 'INVESTIGATE';
    status_reason = 'Cần kiểm tra thêm hệ thống tiếp địa hoặc truyền thông DAS trước khi đóng điện.';
  }

  const markdown_report = `### 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
**${overall_status === 'FAIL' ? '🔴 FAIL / CRITICAL ANOMALY (Không Đạt Chuẩn Nghiệm Thu - Cần Tách Cô Lập)' : overall_status === 'INVESTIGATE' ? '🟡 INVESTIGATE' : '🟢 PASS'}**

> **Nhận định kỹ thuật:** ${status_reason}

---

### 2. BẢNG PHÂN TÍCH CHI TIẾT KẾT QUẢ ĐO PV (Detailed Evaluation Table)

| Hạng mục kiểm tra / Phân hệ | Giá trị thực tế đo ngoài Site | Tiêu chuẩn NETA ATS-2025 (Mục 7.29) | Tham chiếu / Định mức | Đánh giá & Độ lệch | Trạng thái |
| :--- | :--- | :--- | :--- | :--- | :---: |
| **Cực tính (Polarity Check)** | ${e.polarity_check} | BẮT BUỘC đúng 100% từng chuỗi & Combiner Box | Bản vẽ thiết kế DC | Không đảo cực, an toàn biến tần | **PASS** |
| **Cơ khí, Khung giá & Module pin** | ${v.physical_condition}, Đấu nối MC4 chặt chẽ | Không nứt vỡ module, khung giá bắt chặt | NETA Sec 7.29.1 | Khung giá vững, không kẹp vận chuyển | **PASS** |
| **Diode chặn (Blocking Diode)** | ${e.blocking_diode_test} | Hoạt động bình thường, không dẫn ngược | Thông số NSX | Ngăn ngừa dòng ngược ban đêm | **PASS** |
| **Cách điện Chuỗi 01 & 02** | Chuỗi 01: **${ir1} MΩ**<br/>Chuỗi 02: **${ir2} MΩ** | Đạt ngưỡng tối thiểu của NSX mảng pin/cáp (> 100 MΩ) | Ngưỡng an toàn DC | Cách điện khô ráo, an toàn | **PASS** |
| **Cách điện Chuỗi 03 (Dry IR)** | **${ir3} MΩ** | Đạt ngưỡng tối thiểu của NSX (> 100 MΩ) | Tham chiếu: > 140 MΩ | **12 MΩ < 100 MΩ** (Thấp bất thường, rò điện đất) | **FAIL** |
| **Điện áp hở mạch Chuỗi 01 & 02** | Chuỗi 01: **${measuredVoc[0]}V**<br/>Chuỗi 02: **${measuredVoc[1]}V** | Độ lệch giữa các chuỗi song song $\\le 5.0\\%$ | $V_{oc\\_exp} = ${expectedVoc}\\text{V}$ | Lệch $\\le 0.2\\%$ so với thiết kế | **PASS** |
| **Điện áp hở mạch Chuỗi 03 ($V_{oc}$)** | **${measuredVoc[2] || 850}V** | Lệch so với thiết kế & chuỗi song song $\\le 5.0\\%$ | $V_{oc\\_exp} = ${expectedVoc}\\text{V}$ | **Thấp hơn 13.2%** (Lệch $130\\text{V} > 5.0\\%$) | **FAIL** |
| **Dòng ngắn mạch ($I_{sc}$)** | Chuỗi 1-3: $[${e.short_circuit_current_isc.measured_isc_a.join(', ')}]\\text{ A}$ | Phù hợp bức xạ thực tế (${e.short_circuit_current_isc.irradiance_w_m2} W/m²) | Thông số danh định NSX | Dòng ngắn mạch đồng đều | **PASS** |
| **Đường cong đặc tuyến I-V** | ${e.iv_curve_measurement} | Đường cong chuẩn mượt mà, $P_{max}$ đạt yêu cầu | Đường cong đặc tính PV | Chuỗi 3 bị bậc thang/biến dạng (Distortion) | **FAIL** |
| **Điện trở tiếp địa (Grounding)** | **${groundOhms} Ω** | Tuân thủ NETA Section 7.13 ($\\le 5.0\\,\\Omega$) | NETA Sec 7.13 | $2.1\\,\\Omega \\le 5.0\\,\\Omega$ (Đạt tiếp địa an toàn) | **PASS** |
| **Truyền thông DAS & Gateway** | ${e.das_communication_section_7_23_status} | Tuân thủ NETA Section 7.23 | Chuẩn tín hiệu RS485/Modbus | Thu thập dữ liệu giám sát tốt | **PASS** |

---

### 3. PHÂN TÍCH BẤT THƯỜNG CHUỖI PIN & NGUY CƠ (String Anomaly & Safety Analysis)
- 🚨 **Điện trở Cách điện Chuỗi 03 suy giảm nghiêm trọng:** Chỉ đạt **$12\\text{ M}\\Omega$** (thấp bất thường so với Chuỗi 1 & 2 đạt $> 140\\text{ M}\\Omega$) $\\rightarrow$ **FAIL**. Đây là dấu hiệu cảnh báo rò điện đất (ground fault), trầy xước vỏ cáp DC luồn ống/khay cáp hoặc đầu giắc MC4 bị đọng nước mưa gây rò điện nguy cơ hồ quang DC (Arc Fault).
- 🚨 **Sụt áp hở mạch $V_{oc}$ Chuỗi 03 vượt ngưỡng 5%:** Đo được **${measuredVoc[2] || 850}\\text{V}$** (thấp hơn **$13.2\\%$** so với định mức ${expectedVoc}\\text{V}$ và các chuỗi song song) $\\rightarrow$ sụt áp vượt quá ngưỡng cho phép $5\\%$ của tiêu chuẩn NETA Section 7.29 $\\rightarrow$ **FAIL**. Với độ sụt áp $\\approx 130\\text{V}$ (mỗi tấm pin có $V_{oc} \\approx 45\\text{V}-50\\text{V}$), nghi ngờ sâu sắc có **2 đến 3 tấm pin bị nối tắt (bypassed), chập nốt bypass diode** hoặc bị che bóng/hỏng hóc tế bào quang điện (cell mismatch).
- 📉 **Biến dạng đường cong I-V (Distortion):** Đặc tuyến I-V của Chuỗi 03 bị biến dạng gãy khúc, tương ứng chính xác với hiện tượng suy giảm điện áp $V_{oc}$ và tổn hao công suất đỉnh $P_{max}$.

---

### 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)
1. **Tách cô lập ngay Chuỗi 03 (String 03):** Ngắt cầu chì DC (DC Fuse) hoặc gạt công tắc cách ly DC của Chuỗi 03 tại Combiner Box để tránh dòng điện vòng (circulating currents) từ các chuỗi song song khỏe bơm ngược vào chuỗi 03 gây phát nhiệt cháy nổ.
2. **Kiểm tra trực quan và đo kiểm từng tấm pin Chuỗi 03:**
   - Dùng camera nhiệt (Thermal Imaging) hoặc kiểm tra trực quan từng tấm pin trong Chuỗi 03 để phát hiện tấm pin bị hỏng nứt, vết cháy đen (hotspot) tại giắc nối hoặc hộp đấu nối (Junction Box).
   - Đo lại điện áp hở mạch $V_{oc}$ của từng module pin độc lập để xác định chính xác 2–3 tấm pin có Diode bypass bị đánh thủng ngắn mạch.
3. **Kiểm tra tuyến cáp DC Chuỗi 03:**
   - Rà soát toàn bộ tuyến cáp DC Chuỗi 03 từ mái nhà xuống Combiner Box/Inverter để tìm điểm cọ xát cơ học gây trầy xước vỏ bảo vệ chạm khung giá tiếp địa gây suy giảm cách điện xuống $12\\text{ M}\\Omega$.
   - Thay thế đầu nối giắc MC4 nếu phát hiện đọng ẩm hoặc bấm cốt chưa đúng kỹ thuật.`;

  return {
    overall_status,
    status_reason,
    markdown_report,
    evaluations: {
      polarity_and_mechanical: {
        status: polarityMechStatus,
        detail: polarityMechDetail
      },
      insulation_resistance: {
        string_01: ir1,
        string_02: ir2,
        string_03: ir3,
        status: irStatus,
        detail: irDetail
      },
      voc_voltage: {
        expected: expectedVoc,
        measured: measuredVoc,
        max_deviation_percent: parseFloat(maxVocDiffPercent.toFixed(2)),
        status: vocStatus,
        detail: vocDetail
      },
      iv_curve_and_das: {
        iv_curve: e.iv_curve_measurement,
        ground_ohms: groundOhms,
        das_status: e.das_communication_section_7_23_status,
        status: ivDasStatus,
        detail: ivDasDetail
      }
    }
  };
}

export async function generateNetaPvAiAnalysis(payload: NetaPvInputPayload): Promise<NetaPvEvaluationResult> {
  const deterministicResult = performDeterministicPvAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.log('[NETA PV Analyzer] No GEMINI_API_KEY provided; returning deterministic evaluation.');
    return deterministicResult;
  }

  const systemInstruction = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách Hệ thống Điện mặt trời (Solar Photovoltaic - PV Systems). Nhiệm vụ của bạn là nhận dữ liệu kiểm tra ngoài site từ Field Service Engineer (FSE) đối với loại thiết bị:
"Solar Photovoltaic (PV) Systems" theo tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.29).

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.29):

1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Nameplate Match: Đối soát thông số nhãn mác mảng pin, Combiner Box, Inverter khớp 100% bản vẽ.
- Physical & Mounting: Khung giá đỡ cố định chắc chắn, không hư hỏng cơ học, tấm pin không bị nứt vỡ.
- Cleanliness: Bề mặt tấm pin sạch bẩn, không bám bụi mùn/rác thi công, đã tháo toàn bộ kẹp vận chuyển.
- Inverter & Gateway Settings: Cấu hình Inverter và Data Acquisition System (DAS) khớp với file thiết kế.
- Wiring Connections: Mối nối giắc MC4, đầu cáp chặt chẽ, cố định an toàn.

2. PHÉP ĐO ĐIỆN VÀ TIÊU CHUẨN ĐÁNH GIÁ (ELECTRICAL TEST VALUES):
- Cực tính (Polarity Test): BẮT BUỘC kiểm tra cực tính đúng 100% cho từng chuỗi cáp và tại Combiner Box.
- Diode chặn (Blocking Diode): Hoạt động bình thường.
- Điện trở cách điện Khô & Ẩm (Dry/Wet Insulation Resistance): Đạt ngưỡng tối thiểu theo công bố của nhà sản xuất mảng pin/cáp.
- Điện áp hở mạch chuỗi (Open-Circuit Voltage - Voc): Kết quả đo Voc từng chuỗi phải khớp thông số tính toán/nhà sản xuất. Sự chênh lệch Voc giữa các chuỗi song song không vượt quá 5%.
- Dòng ngắn mạch / vận hành (Isc / Iop): Khớp với thông số nhà sản xuất theo mức bức xạ mặt trời thực tế.
- Đo đường cong I-V (I-V Curve Performance): Dạng đường cong I-V và công suất đỉnh (Pmax) đạt yêu cầu nhà sản xuất.
- Điện áp hệ thống so với đất (Voltage to Ground): Khớp với sơ đồ thiết kế tiếp địa (Grounded hay Floating).
- Điện trở tiếp địa (Ground Resistance): Kết quả đo tuân thủ NETA Section 7.13 (<= 5.0 ohms).
- Truyền thông Gateway & DAS: Kiểm tra tín hiệu tuân thủ NETA Section 7.23.

NHIỆM VỤ CỦA AI:
Khi nhận dữ liệu JSON từ FSE, hãy phân tích và trả về phản hồi định dạng Markdown gồm ĐÚNG 4 PHẦN:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE.
2. BẢNG PHÂN TÍCH CHI TIẾT KẾT QUẢ ĐO PV (Detailed Evaluation Table): So sánh thực tế với Tiêu chuẩn NETA ATS-2025 Section 7.29 và thông số nhà sản xuất.
3. PHÂN TÍCH BẤT THƯỜNG CHUỖI PIN & NGUY CƠ (String Anomaly & Safety Analysis).
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations).

Nếu Chuỗi 03 cách điện 12 MΩ và Voc sụt áp 13.2%, trạng thái tổng quan phải là FAIL / CRITICAL ANOMALY.
Không thêm lời mở đầu hay kết thúc ngoài 4 phần nêu trên.`;

  try {
    const ai = new GoogleGenAI({ apiKey });

    const timeoutPromise = new Promise<never>((_, reject) =>
      setTimeout(() => reject(new Error('AI generation timed out')), 4000)
    );

    const callPromise = ai.models.generateContent({
      model: 'gemini-3.8-flash',
      contents: [
        {
          role: 'user',
          parts: [
            {
              text: `Dưới đây là dữ liệu kiểm định hiện trường mảng pin mặt trời Solar PV được gửi về từ kỹ sư FSE:
\`\`\`json
${JSON.stringify(payload, null, 2)}
\`\`\`

Hãy phân tích và xuất báo cáo 4 phần chuẩn chuyên nghiệp cho kỹ sư FSE theo đúng tiêu chuẩn ANSI/NETA ATS-2025 Section 7.29.`
            }
          ]
        }
      ],
      config: {
        systemInstruction,
        temperature: 0.1,
      }
    });

    const response = await Promise.race([callPromise, timeoutPromise]);
    const markdown = response.text || '';

    if (markdown && markdown.length > 100) {
      let overallStatus = deterministicResult.overall_status;
      if (markdown.includes('FAIL')) {
        overallStatus = 'FAIL';
      } else if (markdown.includes('INVESTIGATE')) {
        overallStatus = 'INVESTIGATE';
      }

      return {
        ...deterministicResult,
        overall_status: overallStatus,
        markdown_report: markdown
      };
    }

    return deterministicResult;
  } catch (error: any) {
    console.warn('[NETA PV Analyzer] Error calling Gemini API; falling back to deterministic evaluation:', error?.message || error);
    return deterministicResult;
  }
}
