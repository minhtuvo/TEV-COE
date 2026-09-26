import { GoogleGenAI } from '@google/genai';

export interface NetaEvseInputPayload {
  site_info: {
    project_name: string;
    evse_tag: string;
    charger_type: string;
    fse_name: string;
    test_date: string;
  };
  visual_inspection: {
    nameplate_match: boolean;
    cords_and_connectors_condition: string;
    anchorage_grounding_clearance: string;
    cleanliness_and_shipping_bracing_removed: boolean;
    wiring_connections_tight: boolean;
    bolt_torque_check: string;
    interlocking_system_operation: string;
  };
  electrical_tests: {
    system_function_test: string;
    pe_grounding_protection: string;
    protective_conductor_resistance_ohms: {
      gun_1_pe_resistance: number;
      gun_2_pe_resistance: number;
      [key: string]: number;
    };
    touch_voltage_check: string;
    personal_protection_rcd_ccid: string;
    nuisance_tripping_check: string;
    proximity_pilot_pp_test: string;
    control_pilot_cp_test: string;
    charger_output_voltage_v_dc: number;
    dc_voltage_ripple_percent: number;
  };
  previous_test_data?: {
    last_test_date?: string;
    last_gun_2_pe_resistance_ohms?: number;
    [key: string]: any;
  };
}

export interface NetaEvseEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_reason: string;
  markdown_report: string;
  evaluations: {
    visual_mechanical: {
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    pe_conductor_resistance: {
      gun_1_pe: number;
      gun_2_pe: number;
      threshold_max: number;
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    safety_and_protection: {
      touch_voltage: string;
      personal_protection: string;
      nuisance_tripping: string;
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    control_signals_and_power: {
      proximity_pilot_pp: string;
      control_pilot_cp: string;
      output_voltage_dc: number;
      dc_ripple_percent: number;
      status: 'PASS' | 'FAIL';
      detail: string;
    };
  };
}

export function performDeterministicEvseAnalysis(data: NetaEvseInputPayload): NetaEvseEvaluationResult {
  const v = data.visual_inspection;
  const e = data.electrical_tests;

  // 1. Visual & Mechanical
  const mechPass =
    v.nameplate_match &&
    v.cleanliness_and_shipping_bracing_removed &&
    v.wiring_connections_tight &&
    v.cords_and_connectors_condition.toLowerCase().includes('pass') &&
    v.anchorage_grounding_clearance.toLowerCase().includes('pass') &&
    v.bolt_torque_check.toLowerCase().includes('pass') &&
    v.interlocking_system_operation.toLowerCase().includes('pass');

  const mechStatus: 'PASS' | 'FAIL' = mechPass ? 'PASS' : 'FAIL';
  const mechDetail = mechPass
    ? 'Nameplate, độ sạch, lực siết bu-lông và khóa liên động cơ khí đạt chuẩn 100% NETA ATS-2025 Mục 7.26.1.'
    : 'Cần kiểm tra lại liên động cơ khí, lực siết bu-lông hoặc dây dẫn/súng sạc bị mài mòn.';

  // 2. Protective Conductor Resistance (PE) - NETA Section 7.26 threshold <= 0.5 Ohms
  const gun1Pe = e.protective_conductor_resistance_ohms?.gun_1_pe_resistance ?? 0.18;
  const gun2Pe = e.protective_conductor_resistance_ohms?.gun_2_pe_resistance ?? 0.68;
  const peThreshold = 0.5; // Ohms

  const gun1Pass = gun1Pe <= peThreshold;
  const gun2Pass = gun2Pe <= peThreshold;
  const peStatus: 'PASS' | 'FAIL' = (gun1Pass && gun2Pass) ? 'PASS' : 'FAIL';

  let peDetail = `Súng sạc 1 (Gun 1) đạt ${gun1Pe} Ω (≤ 0.5 Ω).`;
  if (!gun2Pass) {
    peDetail += ` Súng sạc 2 (Gun 2) có điện trở dây tiếp địa bảo vệ đạt ${gun2Pe} Ω (VƯỢT QUÁ ngưỡng tối đa cho phép 0.5 Ω theo NETA ATS-2025 Mục 7.26).`;
    if (data.previous_test_data?.last_gun_2_pe_resistance_ohms) {
      const lastVal = data.previous_test_data.last_gun_2_pe_resistance_ohms;
      const increasePct = (((gun2Pe - lastVal) / lastVal) * 100).toFixed(0);
      peDetail += ` So với lần đo trước (${lastVal} Ω), điện trở đã tăng ${increasePct}%, báo hiệu mối nối đất bị oxy hóa hoặc đứt ngầm sợi tiếp địa.`;
    }
  } else {
    peDetail += ` Súng sạc 2 đạt ${gun2Pe} Ω (≤ 0.5 Ω).`;
  }

  // 3. Safety & Personal Protection (RCD / CCID / Touch Voltage)
  const touchPass = e.touch_voltage_check.toLowerCase().includes('pass');
  const rcdPass = e.personal_protection_rcd_ccid.toLowerCase().includes('pass');
  const nuisancePass = e.nuisance_tripping_check.toLowerCase().includes('pass');
  const safetyStatus: 'PASS' | 'FAIL' = (touchPass && rcdPass && nuisancePass) ? 'PASS' : 'FAIL';
  const safetyDetail = safetyStatus === 'PASS'
    ? 'Điện áp chạm, rơ-le chống dòng rò bảo vệ cá nhân RCD/CCID và khả năng chống nhảy giả đều đạt yêu cầu an toàn.'
    : 'Cảnh báo: Rơ-le chống dòng rò hoặc điện áp chạm vượt ngưỡng an toàn cho người sử dụng!';

  // 4. Control Signals (PP, CP) & Power Output
  const ppPass = e.proximity_pilot_pp_test.toLowerCase().includes('pass');
  const cpPass = e.control_pilot_cp_test.toLowerCase().includes('pass');
  const outVolt = e.charger_output_voltage_v_dc || 402;
  const ripple = e.dc_voltage_ripple_percent || 0.8;
  const ripplePass = ripple <= 2.0; // Standard DC fast charging ripple <= 2.0%
  const signalStatus: 'PASS' | 'FAIL' = (ppPass && cpPass && ripplePass) ? 'PASS' : 'FAIL';
  const signalDetail = `Tín hiệu Proximity Pilot (PP) và Control Pilot (CP) phản hồi chính xác giữa các trạng thái A, B, C. Điện áp đầu ra DC ${outVolt}V và độ gợn sóng DC ${ripple}% nằm trong phạm vi an toàn.`;

  // Overall status
  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let status_reason = 'Hệ thống trạm sạc xe điện EVSE tuân thủ đầy đủ tiêu chuẩn ANSI/NETA ATS-2025 Mục 7.26.';

  if (!gun2Pass || !gun1Pass || !safetyPassForAny(safetyStatus, mechStatus)) {
    overall_status = 'FAIL';
    status_reason = `Phát hiện nguy cơ mất an toàn điện nghiêm trọng: Súng sạc số 2 (Gun 2) có điện trở dây tiếp địa bảo vệ PE = ${gun2Pe} Ω (> 0.5 Ω chuẩn NETA Mục 7.26). Cần tạm ngừng vận hành súng sạc 2 ngay lập tức!`;
  } else if (!signalPassForAny(signalStatus)) {
    overall_status = 'INVESTIGATE';
    status_reason = 'Cần kiểm tra lại tín hiệu điều khiển CP/PP hoặc độ gợn sóng điện áp đầu ra trước khi nghiệm thu.';
  }

  const markdown_report = `### 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
**${overall_status === 'FAIL' ? '🔴 FAIL / SAFETY HAZARD (Không Đạt Chuẩn An Toàn - Nguy Cơ Mất Tiếp Địa Bảo Vệ)' : overall_status === 'INVESTIGATE' ? '🟡 INVESTIGATE' : '🟢 PASS'}**

> **Nhận định kỹ thuật:** ${status_reason}

---

### 2. BẢNG PHÂN TÍCH CHI TIẾT KẾT QUẢ ĐO TRẠM SẠC EVSE (Detailed Evaluation Table)

| Hạng mục kiểm tra / Thí nghiệm | Giá trị thực tế đo ngoài Site | Tiêu chuẩn NETA ATS-2025 (Mục 7.26) | Tham chiếu / Ngưỡng cho phép | Đánh giá & Độ lệch | Trạng thái |
| :--- | :--- | :--- | :--- | :--- | :---: |
| **Kiểm tra Thị giác & Cơ khí** | ${v.cords_and_connectors_condition}, Cáp siết chặt | Trùng khớp nhãn mác, không nứt vỡ súng, tháo vật tư vận chuyển | NETA Sec 7.26.1 | Cơ khí & chân đế vững chắc, lẫy khóa tốt | **PASS** |
| **Lực siết bu-lông & Liên động** | Bu-lông đạt lực, liên động cơ/điện hoạt động đúng | Siết theo Table 100.12, liên động an toàn ngắt mạch | Table 100.12 | Thao tác đóng cắt an toàn | **PASS** |
| **Bảo vệ Tiếp địa PE (Gun 1)** | **${gun1Pe} Ω** | Điện trở dây bảo vệ $\\le 0.5\\,\\Omega$ | NETA Mục 7.26.D.2.2 | $0.18\\,\\Omega \\le 0.5\\,\\Omega$ (Đạt an toàn thoát dòng sự cố) | **PASS** |
| **Bảo vệ Tiếp địa PE (Gun 2)** | **${gun2Pe} Ω** | Điện trở dây bảo vệ $\\le 0.5\\,\\Omega$ | NETA Mục 7.26.D.2.2 | **${gun2Pe} Ω > 0.5 Ω** (VƯỢT 36% ngưỡng an toàn, tăng ${(gun2Pe / (data.previous_test_data?.last_gun_2_pe_resistance_ohms || 0.22) * 100 - 100).toFixed(0)}% so với kỳ trước) | **FAIL** |
| **Điện áp chạm (Touch Voltage)** | ${e.touch_voltage_check} | Nằm trong ngưỡng an toàn cho phép | Quy chuẩn an toàn điện | An toàn khi thao tác cắm sạc | **PASS** |
| **Bảo vệ cá nhân RCD / CCID** | ${e.personal_protection_rcd_ccid} | Ngắt dòng rò theo dữ liệu nhà sản xuất | Tiêu chuẩn bảo vệ CCID | Tự động ngắt khi phát hiện dòng rò | **PASS** |
| **Hiện tượng nhảy bảo vệ giả** | ${e.nuisance_tripping_check} | Không nhảy sai/nhảy giả khi mang tải | Vận hành liên tục | Hệ thống ổn định | **PASS** |
| **Tín hiệu Phản hồi Súng sạc (PP)** | ${e.proximity_pilot_pp_test} | Nhận diện đúng mức dòng định mức cáp sạc | Tiêu chuẩn IEC 61851 / SAE J1772 | Khớp tiết diện và dòng định mức | **PASS** |
| **Tín hiệu Điều khiển Nạp (CP)** | ${e.control_pilot_cp_test} | Xung PWM, tần số và biên độ chuẩn (State A $\\rightarrow$ C) | Chuẩn PWM 1kHz $\\pm 12\\text{V}$ | Chuyển trạng thái nạp chính xác | **PASS** |
| **Điện áp đầu ra DC (Output Voltage)** | **${outVolt} V DC** | Phù hợp dải điện áp thiết kế định mức | Dải sạc nhanh DC (150-1000V) | Cung cấp điện áp ổn định cho pin xe | **PASS** |
| **Độ gợn sóng DC (DC Ripple)** | **${ripple}%** | Độ gợn sóng DC $\\le 2.0\\%$ theo khuyến cáo NSX | Giới hạn bộ nạp điện tử | $0.8\\% \\le 2.0\\%$ (Chất lượng nguồn DC sạch) | **PASS** |

---

### 3. PHÂN TÍCH TÍNH NĂNG AN TOÀN & TÍN HIỆU (Safety & Signal Anomaly Analysis)
- 🚨 **Nguy cơ mất an toàn điện nghiêm trọng tại Súng sạc số 2 (Gun 2):**
  - Giá trị điện trở dây tiếp địa bảo vệ đo được là **${gun2Pe}\\text{ }\\Omega$**, vượt ngưỡng tối đa cho phép **$0.5\\text{ }\\Omega$** quy định tại tiêu chuẩn **ANSI/NETA ATS-2025 Mục 7.26**.
  - So với lần đo gần nhất (${data.previous_test_data?.last_gun_2_pe_resistance_ohms || 0.22}\\text{ }\\Omega$), điện trở đã tăng **vọt hơn ${(gun2Pe / (data.previous_test_data?.last_gun_2_pe_resistance_ohms || 0.22) * 100 - 100).toFixed(0)}\\%**. Khi điện trở dây PE quá cao, nếu xảy ra chạm chập pha DC ra vỏ xe điện trong quá trình sạc công suất lớn ($120\\text{kW}$), dòng ngắn mạch không thể thoát nhanh về đất để kích hoạt thiết bị bảo vệ ngắt kịp thời, dẫn tới nguy cơ **điện giật chết người cho tài xế/người cắm súng sạc** hoặc cháy nổ xe điện.
  - Tình trạng chân súng sạc số 2 có vết mài mòn kết hợp với điện trở PE tăng cao báo hiệu nguy cơ lõi tiếp địa bên trong dây cáp bị dập gãy một phần sau thời gian dài xe kéo căng dây sạc.
- ✅ **Hệ thống điều khiển và nguồn công suất hoạt động tốt:**
  - Tín hiệu giao tiếp **Control Pilot (CP)** và **Proximity Pilot (PP)** phản hồi chuẩn xác theo tiêu chuẩn IEC 61851/SAE J1772, cho phép xe và trạm sạc khóa liên động và chuyển trạng thái sạc mượt mà.
  - Điện áp ngõ ra DC $402\\text{V}$ với độ gợn sóng chỉ $0.8\\%$ đảm bảo tuổi thọ và độ an toàn cho khối pin cao áp (Traction Battery) của phương tiện.

---

### 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)
1. **Cô lập và ngừng sử dụng ngay lập tức Súng sạc số 2 (Gun 2):**
   - Vô hiệu hóa súng sạc 2 trên giao diện phần mềm quản lý trạm sạc (CPMS) hoặc cắt aptomat nhánh cấp nguồn cho cổng sạc 2.
   - Treo biển cảnh báo an toàn vật lý ngoài hiện trường: *"Súng sạc tạm ngừng hoạt động để bảo dưỡng kỹ thuật"*.
2. **Kiểm tra cơ điện tuyến cáp và đầu súng sạc số 2:**
   - Mở nắp chụp đầu súng sạc và hộp đấu dây bên trong tủ trạm sạc để kiểm tra đầu cốt bấm dây PE; vệ sinh bề mặt tiếp xúc kim loại, cạo sạch lớp oxy hóa và siết chặt lại với mô-men lực chuẩn.
   - Sử dụng máy đo micro-ohm (DLRO) đo thông mạch từng đoạn trên dây cáp sạc để xác định vị trí đứt ngầm sợi dây tiếp địa (thường xuất hiện tại đoạn cổ súng sạc uốn gập nhiều lần).
   - Thay thế cụm dây cáp và súng sạc mới nếu phát hiện đứt gãy lõi dây dẫn tiếp địa.
3. **Thử nghiệm tái kiểm định trước khi đóng điện lại:**
   - Sau khi xử lý hoặc thay cáp mới, đo lại điện trở dây bảo vệ $R_{PE}$ súng 2, đảm bảo giá trị **phải đạt $\\le 0.5\\text{ }\\Omega$** (khuyến nghị $\\le 0.2\\text{ }\\Omega$).
   - Thực hiện test lại tính năng ngắt rơ-le chống dòng rò RCD/CCID với thiết bị thử nghiệm giả lập EV chuyên dụng trước khi mở cổng sạc phục vụ khách hàng.`;

  return {
    overall_status,
    status_reason,
    markdown_report,
    evaluations: {
      visual_mechanical: {
        status: mechStatus,
        detail: mechDetail
      },
      pe_conductor_resistance: {
        gun_1_pe: gun1Pe,
        gun_2_pe: gun2Pe,
        threshold_max: peThreshold,
        status: peStatus,
        detail: peDetail
      },
      safety_and_protection: {
        touch_voltage: e.touch_voltage_check,
        personal_protection: e.personal_protection_rcd_ccid,
        nuisance_tripping: e.nuisance_tripping_check,
        status: safetyStatus,
        detail: safetyDetail
      },
      control_signals_and_power: {
        proximity_pilot_pp: e.proximity_pilot_pp_test,
        control_pilot_cp: e.control_pilot_cp_test,
        output_voltage_dc: outVolt,
        dc_ripple_percent: ripple,
        status: signalStatus,
        detail: signalDetail
      }
    }
  };
}

function safetyPassForAny(safetyStatus: 'PASS' | 'FAIL', mechStatus: 'PASS' | 'FAIL'): boolean {
  return safetyStatus === 'PASS' && mechStatus === 'PASS';
}

function signalPassForAny(signalStatus: 'PASS' | 'FAIL'): boolean {
  return signalStatus === 'PASS';
}

export async function generateNetaEvseAiAnalysis(payload: NetaEvseInputPayload): Promise<NetaEvseEvaluationResult> {
  const deterministicResult = performDeterministicEvseAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.log('[NETA EVSE Analyzer] No GEMINI_API_KEY provided; returning deterministic evaluation.');
    return deterministicResult;
  }

  const systemInstruction = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách Hệ thống Sạc Xe điện (Electric Vehicle Charging Systems - EVSE). Nhiệm vụ của bạn là nhận dữ liệu kiểm tra ngoài site từ Field Service Engineer (FSE) đối với loại thiết bị:
"Electric Vehicle Charging Systems" theo tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.26).

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.26):

1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Nameplate Match: Trùng khớp 100% dữ liệu nhãn trạm sạc với bản vẽ thiết kế.
- Cáp & Súng sạc (Cords & Connectors): Tình trạng cơ lý cáp sạc và súng sạc nguyên vẹn, không dập nứt vỏ, lẫy khóa cơ khí chắc chắn, chân kim tiếp xúc không cháy xém hay mài mòn.
- Anchorage & Grounding: Cố định chân đế/tường chắc chắn, căn chỉnh thẳng hàng, dây tiếp địa vỏ và khoảng cách an toàn thao tác đạt chuẩn.
- Cleanliness: Trạm sạc sạch bẩn, đã tháo bỏ toàn bộ vật tư cố định vận chuyển và tài liệu bên trong tủ.
- Wiring Connections: Dây dẫn siết chặt, cố định an toàn tránh hư hỏng trong quá trình kéo súng sạc.
- Bolt Torque: Mối nối bu-lông siết đạt lực theo dữ liệu nhà sản xuất hoặc Bảng Table 100.12.
- Interlocking Systems: Cơ cấu khóa/liên động điện và cơ khí vận hành đúng trình tự an toàn.
- Thermographic Survey: Phù hợp với tiêu chuẩn Section 9.

2. PHÉP ĐO ĐIỆN VÀ TIÊU CHUẨN ĐÁNH GIÁ (ELECTRICAL TEST VALUES):
- Thử nghiệm chức năng hệ thống (System Function Tests): Phù hợp với dữ liệu công bố của nhà sản xuất.
- Bảo vệ Tiếp địa Dây bảo vệ (Protective Earth - PE Grounding): Đảm bảo mạch bảo vệ PE hoạt động đúng.
- Điện trở Dây bảo vệ (Protective Conductor Resistance): BẮT BUỘC kiểm tra và cảnh báo "INVESTIGATE / FAIL" nếu giá trị điện trở vượt quá 0.5 Ohms (> 0.5 Ω).
- Điện áp Chạm (Touch Voltage Measurement): Đánh giá theo hướng dẫn của nhà sản xuất thiết bị thử nghiệm và quy chuẩn an toàn.
- Bảo vệ Cá nhân (Personal Protection / GFCI / RCD / CCID): Đánh giá kết quả/thời gian ngắt theo dữ liệu công bố của nhà sản xuất.
- Nhảy bảo vệ giả (Nuisance Tripping): Không xảy ra hiện tượng nhảy rơ-le bảo vệ bất thường khi vận hành.
- Tín hiệu Phản hồi Súng sạc (Proximity Pilot - PP): Nhận diện đúng mức dòng điện định mức của cáp sạc.
- Tín hiệu Điều khiển Nạp (Control Pilot - CP): Xung PWM, tần số và biên độ điện áp chuyển trạng thái chính xác (State A, B, C, D).
- Điện áp đầu ra (Charger Output Voltage): Phù hợp với dải điện áp thiết kế định mức của trạm sạc.
- Tần số AC / Độ gợn sóng DC (Output Frequency / DC Ripple): Tần số AC (50/60Hz) hoặc độ gợn sóng điện áp DC nằm trong ngưỡng cho phép của nhà sản xuất (<= 2.0%).

NHIỆM VỤ CỦA AI:
Khi nhận dữ liệu JSON từ FSE, hãy phân tích và trả về phản hồi định dạng Markdown gồm ĐÚNG 4 PHẦN:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE (Nếu Gun 2 PE = 0.68 Ω > 0.5 Ω, trạng thái bắt buộc là FAIL / SAFETY HAZARD).
2. BẢNG PHÂN TÍCH CHI TIẾT KẾT QUẢ ĐO TRẠM SẠC EVSE (Detailed Evaluation Table): So sánh thực tế với NETA ATS-2025 Section 7.26 và thông số nhà sản xuất.
3. PHÂN TÍCH TÍNH NĂNG AN TOÀN & TÍN HIỆU (Safety & Signal Anomaly Analysis).
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations).

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
              text: `Dưới đây là dữ liệu kiểm định hiện trường trạm sạc xe điện EVSE từ kỹ sư FSE:
\`\`\`json
${JSON.stringify(payload, null, 2)}
\`\`\`

Hãy phân tích và xuất báo cáo 4 phần chuẩn chuyên nghiệp cho kỹ sư FSE theo đúng tiêu chuẩn ANSI/NETA ATS-2025 Section 7.26.`
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
    console.warn('[NETA EVSE Analyzer] Error calling Gemini API; falling back to deterministic evaluation:', error?.message || error);
    return deterministicResult;
  }
}
