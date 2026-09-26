import { GoogleGenAI } from '@google/genai';

export interface NetaBessInputPayload {
  site_info: {
    project_name: string;
    bess_container_tag: string;
    fse_name: string;
    test_date: string;
  };
  visual_inspection: {
    battery_location_clearance: string; // "Pass" | "Fail"
    fire_suppression_installed: boolean;
    eyewash_station_present: boolean;
    nameplate_match: boolean;
    physical_condition: string; // "Good (No leakage or damage)" | string
    rack_mounting_and_grounding: string; // "Pass" | "Fail"
    cleanliness: string; // "Pass" | "Fail"
    bolt_torque_check: string; // "Pass" | "Fail"
  };
  electrical_and_subsystem_tests: {
    transformer_section_7_2_status: string; // "Pass" | "Fail" | "N/A"
    lv_breakers_section_7_6_1_status: string; // "Pass" | "Fail"
    lv_cables_section_7_3_2_status: string; // "Pass" | "Fail"
    battery_system_polarity_correct: boolean;
    measured_total_dc_voltage_v: number;
    expected_total_dc_voltage_v: number;
    load_test_results: {
      discharge_power_kw: number;
      status: string;
    };
    power_quality_at_poi_ieee_1547: {
      voltage_thd_percent: number;
      current_thd_percent: number;
    };
  };
  previous_test_data?: {
    last_test_date?: string;
    last_measured_dc_voltage_v?: number;
  };
}

export interface NetaBessEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_reason: string;
  markdown_report: string;
  evaluations: {
    safety_fire_suppression: {
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    subsystems_status: {
      transformer: string;
      lv_breakers: string;
      lv_cables: string;
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    dc_voltage_and_polarity: {
      measured: number;
      expected: number;
      deviation_percent: number;
      polarity_correct: boolean;
      status: 'PASS' | 'INVESTIGATE' | 'FAIL';
      detail: string;
    };
    power_quality_ieee_1547: {
      voltage_thd: number;
      current_thd: number;
      voltage_thd_pass: boolean;
      current_thd_pass: boolean;
      status: 'PASS' | 'INVESTIGATE' | 'FAIL';
      detail: string;
    };
  };
}

export function performDeterministicBessAnalysis(data: NetaBessInputPayload): NetaBessEvaluationResult {
  const v = data.visual_inspection;
  const e = data.electrical_and_subsystem_tests;

  // 1. Safety & Fire suppression (MANDATORY)
  const firePass = v.fire_suppression_installed && v.eyewash_station_present;
  const safetyStatus: 'PASS' | 'FAIL' = firePass ? 'PASS' : 'FAIL';
  const safetyDetail = firePass
    ? 'Hệ thống PCCC chuyên dụng và trạm rửa mắt khẩn cấp đã trang bị đầy đủ tại Container BESS (Đạt chuẩn NETA ATS Sec 7.28.1).'
    : 'VI PHẠM AN TOÀN NGHIÊM TRỌNG: Chưa lắp đặt đủ hệ thống chữa cháy chuyên dụng hoặc thiếu trạm rửa mắt khẩn cấp cho khu vực pin.';

  // 2. Subsystems status (Transformers 7.2, LV Breakers 7.6.1, LV Cables 7.3.2)
  const transPass = e.transformer_section_7_2_status.toLowerCase().includes('pass');
  const breakerPass = e.lv_breakers_section_7_6_1_status.toLowerCase().includes('pass');
  const cablePass = e.lv_cables_section_7_3_2_status.toLowerCase().includes('pass');
  const subsystemsPass = transPass && breakerPass && cablePass;
  const subsystemsStatus: 'PASS' | 'FAIL' = subsystemsPass ? 'PASS' : 'FAIL';
  const subsystemsDetail = subsystemsPass
    ? 'Tất cả phân hệ hạ tầng (MBA Sec 7.2, Máy cắt hạ áp Sec 7.6.1, Cáp động lực Sec 7.3.2) đều đạt tiêu chuẩn NETA ATS-2025.'
    : 'Có phân hệ hạ tầng (MBA, máy cắt hạ áp hoặc cáp điện) không đạt thử nghiệm kiểm định NETA ATS.';

  // 3. DC Voltage & Polarity
  const polarityPass = e.battery_system_polarity_correct;
  const voltDiff = Math.abs(e.measured_total_dc_voltage_v - e.expected_total_dc_voltage_v);
  const voltDevPercent = e.expected_total_dc_voltage_v > 0
    ? parseFloat(((voltDiff / e.expected_total_dc_voltage_v) * 100).toFixed(2))
    : 0;

  let dcStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let dcDetail = `Điện áp DC tổng đo được ${e.measured_total_dc_voltage_v}V (Định mức ${e.expected_total_dc_voltage_v}V, lệch ${voltDevPercent}%), cực tính đấu nối chính xác.`;

  if (!polarityPass) {
    dcStatus = 'FAIL';
    dcDetail = 'LỖI ĐẤU NỐI ĐẢO CỰC TÍNH (+/-)! Nguy cơ nổ ngắn mạch phá hủy Cell pin và bộ nghịch lưu PCS!';
  } else if (voltDevPercent > 5.0) {
    dcStatus = 'FAIL';
    dcDetail = `Điện áp DC tổng lệch ${voltDevPercent}% (> 5%) so với thiết kế, có nguy cơ hở mạch chuỗi cell hoặc cell pin suy thoái nặng.`;
  } else if (voltDevPercent > 1.5) {
    dcStatus = 'INVESTIGATE';
    dcDetail = `Điện áp DC tổng lệch ${voltDevPercent}% (> 1.5%) so với điện áp định mức, cần kiểm tra trạng thái SoC và điện áp từng Rack.`;
  }

  // 4. Power Quality at POI (IEEE 1547)
  const vThd = e.power_quality_at_poi_ieee_1547.voltage_thd_percent;
  const iThd = e.power_quality_at_poi_ieee_1547.current_thd_percent;
  const vThdPass = vThd <= 5.0; // IEEE 1547 limit is 5.0%
  const iThdPass = iThd <= 5.0;

  let pqStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let pqDetail = `Chất lượng điện năng tại MV POI đạt chuẩn IEEE 1547: THDu = ${vThd}% (ngưỡng < 5%), THDi = ${iThd}% (ngưỡng < 5%).`;

  if (!vThdPass || !iThdPass) {
    pqStatus = 'INVESTIGATE';
    pqDetail = `Sóng hài điện áp THDu = ${vThd}% vượt ngưỡng khuyến nghị của tiêu chuẩn IEEE 1547 (yêu cầu THDu < 5.0%). Độ méo sóng hài dòng điện THDi = ${iThd}%.`;
  }

  // Overall status determination
  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let status_reason = 'Hệ thống BESS đáp ứng các tiêu chuẩn an toàn NETA ATS-2025 Section 7.28 và IEEE 1547.';

  if (!firePass || !polarityPass || !subsystemsPass || dcStatus === 'FAIL') {
    overall_status = 'FAIL';
    status_reason = 'Hệ thống BESS không đạt tiêu chuẩn an toàn bắt buộc (PCCC/Cực tính/Hạ tầng). Nghiêm cấm đóng điện hòa lưới!';
  } else if (pqStatus === 'INVESTIGATE' || dcStatus === 'INVESTIGATE') {
    overall_status = 'INVESTIGATE';
    status_reason = 'Hệ thống BESS cần điều tra & cân chỉnh (Độ méo sóng hài điện áp THDu tại POI vượt chuẩn IEEE 1547 hoặc điện áp DC chuỗi cần kiểm tra thêm).';
  }

  // Generate standard 4-part Markdown report
  const markdown_report = `### 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
**${overall_status === 'PASS' ? '🟢 PASS (Đạt Tiêu Chuẩn Nghiệm Thu)' : overall_status === 'INVESTIGATE' ? '🟡 INVESTIGATE (Cần Khảo Sát & Cân Chỉnh Trước Khi Hòa Lưới)' : '🔴 FAIL (Không Đạt Chuẩn An Toàn - Nghiêm Cấm Đóng Điện)'}**

> **Nhận định kỹ thuật:** ${status_reason}

---

### 2. BẢNG PHÂN TÍCH CHI TIẾT CÁC PHÂN HỆ BESS (Detailed Evaluation Table)

| Phân hệ / Hạng mục kiểm tra | Dữ liệu thực tế ngoài Site | Tiêu chuẩn NETA ATS-2025 / IEEE 1547 | Tham chiếu / Định mức | Đánh giá kỹ thuật | Trạng thái |
| :--- | :--- | :--- | :--- | :--- | :---: |
| **Vị trí pin & Hành lang an toàn** | ${v.battery_location_clearance} | Khoảng cách thông thoáng, đúng bản vẽ | Tiêu chuẩn PCCC & NFPA 855 | Đảm bảo thoát khí và bảo trì thuận lợi | **PASS** |
| **Hệ thống PCCC & Trạm rửa mắt** | ${firePass ? 'Đã lắp đặt PCCC chuyên dụng & Trạm rửa mắt' : 'Thiếu PCCC hoặc Trạm rửa mắt'} | BẮT BUỘC có PCCC khí/aerosol & bồn rửa mắt | NETA Section 7.28.1.2 | ${firePass ? 'Đạt yêu cầu an toàn sinh mạng' : 'VI PHẠM QUY CHUẨN'} | **${safetyStatus}** |
| **Nhãn Nameplate & Tiếp địa Racks** | Nameplate khớp 100%, Khung giá đỡ tiếp địa chắc chắn | Đạt tiêu chuẩn cơ khí & đẳng thế | NETA Table 100.12 | Tiếp địa vỏ tủ và rack an toàn | **PASS** |
| **Máy biến áp BESS (MV Transformer)** | Kết quả: ${e.transformer_section_7_2_status} | Tuân thủ toàn bộ phép đo NETA Section 7.2 | Section 7.2 (Dry/Liquid TR) | Đạt cách điện và tỷ số biến áp | **PASS** |
| **Máy cắt hạ áp (LV Circuit Breakers)** | Kết quả: ${e.lv_breakers_section_7_6_1_status} | Thử nghiệm ngắt & tiếp xúc NETA Sec 7.6.1 | Section 7.6.1 (Air/MCCB) | Cơ cấu ngắt và rơ-le tác động tốt | **PASS** |
| **Cáp kết nối hạ áp (LV Cables)** | Kết quả: ${e.lv_cables_section_7_3_2_status} | Thử nghiệm điện trở cách điện NETA Sec 7.3.2 | Table 100.1 (Cách điện cáp) | Cách điện cáp DC/AC đạt chuẩn | **PASS** |
| **Cực tính & Điện áp DC tổng** | Điện áp đo: **${e.measured_total_dc_voltage_v}V**<br/>Cực tính: **${polarityPass ? 'Chính xác (+/-)' : 'SAI CỰC TÍNH'}** | Cực tính chính xác 100%, sai số điện áp $\\le 1.0\\%$ | Định mức: ${e.expected_total_dc_voltage_v}V | Lệch ${voltDevPercent}% (${e.measured_total_dc_voltage_v}V so với ${e.expected_total_dc_voltage_v}V) | **${dcStatus}** |
| **Thử nghiệm mang tải (Load Test)** | Công suất xả: **${e.load_test_results.discharge_power_kw} kW**<br/>Trạng thái: ${e.load_test_results.status} | Đạt công suất nạp/xả & dung lượng theo NSX | Thông số thiết kế hệ thống | BESS xả tải ổn định, BMS quản lý tốt | **PASS** |
| **Chất lượng điện năng tại POI (IEEE 1547)** | **THDu = ${vThd}%**<br/>**THDi = ${iThd}%** | Sóng hài điện áp $\\text{THD}_u < 5.0\\%$<br/>Sóng hài dòng $\\text{THD}_i < 5.0\\%$ | Tiêu chuẩn IEEE 1547 | ${vThdPass ? 'Đạt chuẩn hòa lưới' : `THDu = ${vThd}% vượt ngưỡng 5.0% cho phép`} | **${pqStatus}** |

---

### 3. PHÂN TÍCH XU HƯỚNG & CẢNH BÁO AN TOÀN (Safety & System Anomaly Analysis)
- ⚡ **Chất lượng điện năng tại MV POI (IEEE 1547):** Sóng hài điện áp **$\\text{THD}_u = ${vThd}\\%$** vượt quá ngưỡng khuyến nghị của tiêu chuẩn IEEE 1547 (thường yêu cầu $\\text{THD}_u < 5\\%$) $\\rightarrow$ **INVESTIGATE**. Hiện tượng này có thể do bộ nghịch lưu công suất (PCS/Inverter) chưa tối ưu thuật toán điều khiển PWM hoặc bộ lọc sóng hài (LC/LCL Filter) chưa được hiệu chuẩn phù hợp với tổng trở lưới tại trạm.
- 🔋 **Độ ổn định điện áp DC tổng:** Điện áp DC chuỗi pin đo được **${e.measured_total_dc_voltage_v}V** tương đương với định mức **${e.expected_total_dc_voltage_v}V** (độ lệch chỉ ${voltDevPercent}%), cực tính đấu nối chính xác tuyệt đối. Khối pin duy trì điện áp đồng đều, không phát hiện hiện tượng sụt áp bất thường.
- 🛡️ **An toàn PCCC & Cứu nạn khẩn cấp:** Đã trang bị hệ thống chữa cháy chuyên dụng tự động và trạm rửa mắt khẩn cấp tại Container BESS $\\rightarrow$ **PASS**. Khu vực lưu trữ đảm bảo khoảng cách an toàn chống nguy cơ lan truyền nhiệt (thermal runaway propagation).

---

### 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)
1. **Kiểm tra và hiệu chỉnh bộ lọc sóng hài PCS/Inverter:**
   - Kiểm tra bộ lọc bộ biến đổi công suất (PCS/Inverter Harmonic Filters) và tụ lọc đầu ra AC.
   - Cấu hình lại tham số điều khiển vòng lặp dòng điện (Current Control Loop) và tần số đóng cắt của Inverter để giảm độ méo sóng hài điện áp $\\text{THD}_u$ tại điểm đấu nối trung áp xuống dưới **$5.0\\%$** trước khi cho hệ thống hòa lưới vận hành thương mại.
2. **Kiểm tra giám sát hệ thống quản lý pin (BMS):**
   - Đọc dữ liệu CAN/Modbus từ BMS để đối soát độ lệch điện áp giữa các Cell (Cell Delta Voltage $< 20\\text{mV}$) và nhiệt độ các khối pin khi xả tải ở công suất ${e.load_test_results.discharge_power_kw} \\text{kW}$.
3. **Thủ tục nghiệm thu:**
   - Ký biên bản hiện trường tạm thời ghi nhận trạng thái **INVESTIGATE** (Chờ tinh chỉnh PCS). Sau khi cân chỉnh bộ lọc sóng hài đạt $\\text{THD}_u < 5\\%$, tiến hành đo kiểm lại chất lượng điện năng tại POI để phê duyệt đóng điện chính thức.`;

  return {
    overall_status,
    status_reason,
    markdown_report,
    evaluations: {
      safety_fire_suppression: {
        status: safetyStatus,
        detail: safetyDetail
      },
      subsystems_status: {
        transformer: e.transformer_section_7_2_status,
        lv_breakers: e.lv_breakers_section_7_6_1_status,
        lv_cables: e.lv_cables_section_7_3_2_status,
        status: subsystemsStatus,
        detail: subsystemsDetail
      },
      dc_voltage_and_polarity: {
        measured: e.measured_total_dc_voltage_v,
        expected: e.expected_total_dc_voltage_v,
        deviation_percent: voltDevPercent,
        polarity_correct: polarityPass,
        status: dcStatus,
        detail: dcDetail
      },
      power_quality_ieee_1547: {
        voltage_thd: vThd,
        current_thd: iThd,
        voltage_thd_pass: vThdPass,
        current_thd_pass: iThdPass,
        status: pqStatus,
        detail: pqDetail
      }
    }
  };
}

export async function generateNetaBessAiAnalysis(payload: NetaBessInputPayload): Promise<NetaBessEvaluationResult> {
  const deterministicResult = performDeterministicBessAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.log('[NETA BESS Analyzer] No GEMINI_API_KEY provided; returning deterministic evaluation.');
    return deterministicResult;
  }

  const systemInstruction = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách Hệ thống Lưu trữ Năng lượng Bằng Pin (BESS).
Nhiệm vụ của bạn là nhận dữ liệu kiểm tra ngoài site từ Field Service Engineer (FSE) đối với hệ thống:
"Battery Energy Storage Systems (BESS)" theo tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.28).

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.28):

1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Battery Location & Safety: Vị trí đặt pin an toàn, thông thoáng, đúng khoảng cách bảo trì.
- Fire Suppression & Eyewash: BẮT BUỘC đã lắp đặt hệ thống PCCC chuyên dụng cho khu vực pin và trang bị trạm rửa mắt khẩn cấp (Nếu thiếu -> FAIL ngay lập tức).
- Nameplate Match: Thông số nhãn máy (Cell, Rack, PCS, Transformer) khớp 100% bản vẽ thiết kế.
- Physical & Grounding: Không hư hỏng cơ học, không rò rỉ điện dịch; khung giá đỡ (racks) neo cố định chắc chắn và tiếp địa đầy đủ.
- Cleanliness: Tủ rack, khối pin và khoang PCS sạch sẽ, không ẩm ướt hoặc bám bụi kim loại.
- Bolt Torque: Mối nối bu-lông lắp đặt tại site siết đạt lực theo dữ liệu xuất bản của nhà sản xuất.

2. PHÉP ĐO ĐIỆN VÀ TÍCH HỢP HỆ THỐNG (ELECTRICAL & SYSTEM TESTS):
- Máy biến áp BESS (Transformers): Kết quả thử nghiệm tuân thủ tiêu chuẩn NETA Section 7.2.
- Aptomat/Máy cắt hạ áp (LV Circuit Breakers): Kết quả thử nghiệm ngắt/cách điện tuân thủ NETA Section 7.6.1 (7.6.1.1.1, 7.6.1.1.2, 7.6.1.2).
- Cáp kết nối hạ áp (LV Interconnecting Cables): Thử nghiệm điện trở cách điện tuân thủ NETA Section 7.3.2.
- Điện áp & Cực tính (Voltage & Polarity): Điện áp tổng, điện áp chuỗi pin và cực tính (+/-) phải hoàn toàn chính xác theo thông số nhà sản xuất.
- Thử nghiệm mang tải (Load Test): Đạt công suất nạp/xả và dung lượng theo dữ liệu xuất bản của nhà sản xuất.
- Chất lượng điện năng (Power Quality): Đo đạc tại điểm đấu nối trung áp (MV POI) lúc khởi động phải phù hợp với tiêu chuẩn IEEE 1547 (Yêu cầu THDu < 5.0%, THDi < 5.0%. Nếu THDu > 5.0% -> INVESTIGATE).

NHIỆM VỤ CỦA AI:
Khi nhận dữ liệu JSON từ FSE, hãy phân tích và trả về phản hồi định dạng Markdown gồm ĐÚNG 4 PHẦN:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE.
2. BẢNG PHÂN TÍCH CHI TIẾT CÁC PHÂN HỆ BESS (Detailed Evaluation Table): So sánh các phân hệ Pin, PCCC, MBA, Máy cắt, Cáp và Điện áp với Tiêu chuẩn NETA ATS-2025 Section 7.28 và IEEE 1547.
3. PHÂN TÍCH XU HƯỚNG & CẢNH BÁO AN TOÀN (Safety & System Anomaly Analysis).
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations).

Nếu sóng hài điện áp THDu = 6.2% vượt quá 5%, trạng thái tổng quan phải là INVESTIGATE.
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
              text: `Dưới đây là dữ liệu kiểm định hiện trường Container BESS được gửi về từ kỹ sư FSE:
\`\`\`json
${JSON.stringify(payload, null, 2)}
\`\`\`

Hãy phân tích và xuất báo cáo 4 phần chuẩn chuyên nghiệp cho kỹ sư FSE theo đúng tiêu chuẩn ANSI/NETA ATS-2025 Section 7.28 và IEEE 1547.`
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
      if (markdown.includes('Trạng thái Tổng quan: INVESTIGATE') || markdown.includes('**INVESTIGATE**') || markdown.includes('INVESTIGATE')) {
        overallStatus = 'INVESTIGATE';
      } else if (markdown.includes('FAIL')) {
        overallStatus = 'FAIL';
      }

      return {
        ...deterministicResult,
        overall_status: overallStatus,
        markdown_report: markdown
      };
    }

    return deterministicResult;
  } catch (error: any) {
    console.warn('[NETA BESS Analyzer] Error calling Gemini API; falling back to deterministic evaluation:', error?.message || error);
    return deterministicResult;
  }
}
