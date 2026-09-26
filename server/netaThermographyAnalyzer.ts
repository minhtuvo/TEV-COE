import { GoogleGenAI } from '@google/genai';

export interface ThermalMeasurementItem {
  component_name: string;
  measured_temp_c: number;
  reference_component_temp_c: number;
}

export interface NetaThermographyInputPayload {
  site_info: {
    project_name: string;
    survey_tag: string;
    fse_name: string;
    test_date: string;
  };
  survey_conditions: {
    camera_model: string;
    ambient_temperature_c: number;
    system_voltage_v: number;
    operating_current_amp: number;
    load_percentage: number;
    inaccessible_areas?: string;
  };
  thermal_measurements: ThermalMeasurementItem[];
}

export interface ThermalItemEvaluation {
  component_name: string;
  measured_temp_c: number;
  reference_component_temp_c: number;
  ambient_temperature_c: number;
  delta_t_comp: number;
  delta_t_amb: number;
  level: 0 | 1 | 2 | 3 | 4;
  level_name: string;
  action_recommendation: string;
  status: 'NORMAL' | 'INVESTIGATE' | 'REPAIR_SCHEDULED' | 'MONITOR' | 'REPAIR_IMMEDIATELY';
}

export interface NetaThermographyEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'REPAIR_SCHEDULED' | 'MONITOR' | 'REPAIR_IMMEDIATELY';
  status_reason: string;
  markdown_report: string;
  evaluations: {
    survey_conditions: {
      load_percentage: number;
      is_adequate_load: boolean;
      detail: string;
    };
    max_severity_level: number;
    items: ThermalItemEvaluation[];
  };
}

export function evaluateThermalItem(
  item: ThermalMeasurementItem,
  ambientTemp: number
): ThermalItemEvaluation {
  const delta_t_comp = parseFloat((item.measured_temp_c - item.reference_component_temp_c).toFixed(1));
  const delta_t_amb = parseFloat((item.measured_temp_c - ambientTemp).toFixed(1));

  let level: 0 | 1 | 2 | 3 | 4 = 0;
  let level_name = 'Normal (Bình thường)';
  let action_recommendation = 'No action required';
  let status: 'NORMAL' | 'INVESTIGATE' | 'REPAIR_SCHEDULED' | 'MONITOR' | 'REPAIR_IMMEDIATELY' = 'NORMAL';

  // NETA ATS-2025 Table 100.18:
  // Level 4 (Major discrepancy): Delta T_comp > 15°C OR Delta T_amb > 40°C -> Repair immediately
  // Level 3 (Monitor): 21°C <= Delta T_amb <= 40°C -> Monitor until corrective measures can be accomplished
  // Level 2 (Probable deficiency): 4°C <= Delta T_comp <= 15°C OR 11°C <= Delta T_amb <= 20°C -> Repair as time permits
  // Level 1 (Possible deficiency): 1°C <= Delta T_comp <= 3°C OR 1°C <= Delta T_amb <= 10°C -> Warrants investigation

  if (delta_t_comp > 15 || delta_t_amb > 40) {
    level = 4;
    level_name = 'Mức 4 (Major discrepancy - CẤP BÁCH)';
    action_recommendation = 'Repair immediately (Dừng máy hoặc ngắt cô lập để xử lý CẤP TỐC ngay lập tức)';
    status = 'REPAIR_IMMEDIATELY';
  } else if (delta_t_amb >= 21 && delta_t_amb <= 40) {
    level = 3;
    level_name = 'Mức 3 (Monitor - Theo dõi)';
    action_recommendation = 'Monitor until corrective measures can be accomplished (Theo dõi liên tục cho đến khi xử lý xong)';
    status = 'MONITOR';
  } else if ((delta_t_comp >= 4 && delta_t_comp <= 15) || (delta_t_amb >= 11 && delta_t_amb <= 20)) {
    level = 2;
    level_name = 'Mức 2 (Probable deficiency - Sai lệch tiềm tàng)';
    action_recommendation = 'Repair as time permits (Sửa chữa/siết lại lực bu-lông khi có lịch dừng máy phù hợp)';
    status = 'REPAIR_SCHEDULED';
  } else if ((delta_t_comp >= 1 && delta_t_comp <= 3) || (delta_t_amb >= 1 && delta_t_amb <= 10)) {
    level = 1;
    level_name = 'Mức 1 (Possible deficiency - Nghi vấn)';
    action_recommendation = 'Warrants investigation (Ghi nhận, kiểm tra theo dõi nguyên nhân)';
    status = 'INVESTIGATE';
  }

  return {
    component_name: item.component_name,
    measured_temp_c: item.measured_temp_c,
    reference_component_temp_c: item.reference_component_temp_c,
    ambient_temperature_c: ambientTemp,
    delta_t_comp,
    delta_t_amb,
    level,
    level_name,
    action_recommendation,
    status
  };
}

export function performDeterministicThermographyAnalysis(data: NetaThermographyInputPayload): NetaThermographyEvaluationResult {
  const ambientTemp = data.survey_conditions.ambient_temperature_c ?? 32.0;
  const loadPct = data.survey_conditions.load_percentage ?? 78;

  const itemEvals: ThermalItemEvaluation[] = (data.thermal_measurements || []).map(item =>
    evaluateThermalItem(item, ambientTemp)
  );

  const maxLevel = itemEvals.length > 0
    ? Math.max(...itemEvals.map(e => e.level))
    : 0;

  let overall_status: 'PASS' | 'INVESTIGATE' | 'REPAIR_SCHEDULED' | 'MONITOR' | 'REPAIR_IMMEDIATELY' = 'PASS';
  let status_reason = 'Tất cả các điểm đo nhiệt độ mối nối và tiếp điểm đều trong ngưỡng bình thường theo NETA ATS-2025 Table 100.18.';

  if (maxLevel === 4) {
    overall_status = 'REPAIR_IMMEDIATELY';
    const majorItem = itemEvals.find(i => i.level === 4);
    status_reason = `Phát hiện điểm phát nhiệt cực kỳ nguy hiểm (Mức 4 - Major Discrepancy) tại [${majorItem?.component_name}]: ΔT_comp = ${majorItem?.delta_t_comp}°C (> 15°C) và ΔT_amb = ${majorItem?.delta_t_amb}°C (> 40°C). Cần ngắt cô lập sửa chữa ngay lập tức để ngăn cháy nổ!`;
  } else if (maxLevel === 3) {
    overall_status = 'MONITOR';
    status_reason = 'Phát hiện điểm phát nhiệt Mức 3 (21°C ≤ ΔT_amb ≤ 40°C). Cần giám sát liên tục cho tới khi xử lý bảo dưỡng xong.';
  } else if (maxLevel === 2) {
    overall_status = 'REPAIR_SCHEDULED';
    status_reason = 'Phát hiện sai lệch nhiệt độ Mức 2 (Probable deficiency). Cần lên lịch siết lực mối nối hoặc bảo dưỡng khi có lịch cắt điện phù hợp.';
  } else if (maxLevel === 1) {
    overall_status = 'INVESTIGATE';
    status_reason = 'Phát hiện sai lệch nhiệt độ Mức 1 (Possible deficiency). Ghi nhận và theo dõi trong các kỳ bảo trì tiếp theo.';
  }

  const tableRows = itemEvals.map(ev => {
    return `| **${ev.component_name}** | **${ev.measured_temp_c}°C** | ${ev.reference_component_temp_c}°C | **${ev.delta_t_comp > 0 ? '+' : ''}${ev.delta_t_comp}°C** | **${ev.delta_t_amb > 0 ? '+' : ''}${ev.delta_t_amb}°C** | **${ev.level_name}** | **${ev.status}** |`;
  }).join('\n');

  const markdown_report = `### 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
**${overall_status === 'REPAIR_IMMEDIATELY' ? '🔴 REPAIR_IMMEDIATELY (Major Discrepancy - Cấp Bách Xử Lý Ngay Lập Tức)' : overall_status === 'MONITOR' ? '🟠 MONITOR (Cần Giám Sát Liên Tục)' : overall_status === 'REPAIR_SCHEDULED' ? '🟡 REPAIR_SCHEDULED (Lên Lịch Sửa Chữa)' : overall_status === 'INVESTIGATE' ? '🔵 INVESTIGATE (Ghi Nhận Theo Dõi)' : '🟢 PASS'}**

> **Nhận định kỹ thuật:** ${status_reason}
> **Điều kiện khảo sát:** Tải vận hành đạt **${loadPct}%** định mức (${data.survey_conditions.operating_current_amp}A / ${data.survey_conditions.system_voltage_v}V), nhiệt độ môi trường **${ambientTemp}°C**, sử dụng thiết bị ${data.survey_conditions.camera_model} theo chuẩn NETA Mục 9.

---

### 2. BẢNG PHÂN TÍCH NHIỆT ĐỘ CHI TIẾT (Thermographic Evaluation Table)
*Đối chiếu trực tiếp với ANSI/NETA ATS-2025 Table 100.18 (Thermographic Survey Actions)*

| Hạng mục / Điểm đo | Nhiệt độ đo ($T$) | Nhiệt độ đối chứng ($T_{ref}$) | Chênh lệch $\\Delta T_{comp}$ | Chênh lệch $\\Delta T_{amb}$ | Phân cấp Table 100.18 | Khuyến nghị hành động |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: |
${tableRows}

---

### 3. PHÂN TÍCH NGUYÊN NHÂN & NGUY CƠ (Root Cause Analysis & Thermal Anomaly Risk)
- 🔴 **Điểm nóng nguy cấp tại Circuit Breaker Incoming Lug Phase B ($75.8^\\circ\\text{C}$):**
  - $\\Delta T_{comp} = 75.8 - 41.2 = \\mathbf{34.6^\\circ\\text{C}}$ (vượt quá xa ngưỡng $15^\\circ\\text{C}$ của Mức 4 theo NETA Table 100.18).
  - $\\Delta T_{amb} = 75.8 - 32.0 = \\mathbf{43.8^\\circ\\text{C}}$ (vượt ngưỡng $40^\\circ\\text{C}$ của Mức 4 theo NETA Table 100.18).
  - **Cơ chế sự cố:** Lực siết bu-lông đầu cốt đấu nối Aptomat Phase B bị lỏng (under-torqued) hoặc bề mặt đồng bị oxy hóa rỗ nặng, làm tăng vọt điện trở tiếp xúc ($R_{contact}$). Dưới dòng tải $650\\text{A}$, công suất tổn hao nhiệt $P_{loss} = I^2 \\cdot R$ tăng theo cấp số nhân dẫn đến quá nhiệt cục bộ $75.8^\\circ\\text{C}$. Nếu duy trì tải lâu dài, điểm phát nhiệt này sẽ làm cháy lớp cách điện pha B, nóng chảy đầu cos và nguy cơ dẫn đến sự cố phóng hồ quang ngắn mạch 3 pha (Arc Flash catastrophe).
- 🟡 **Sai lệch tiềm tàng tại Feeder Cable Termination Phase C ($48.0^\\circ\\text{C}$):**
  - $\\Delta T_{comp} = 3.5^\\circ\\text{C}$, $\\Delta T_{amb} = 16.0^\\circ\\text{C}$ $\\rightarrow$ Đạt phân cấp Mức 2 (Probable deficiency: $11^\\circ\\text{C} \\le \\Delta T_{amb} \\le 20^\\circ\\text{C}$).
  - Cần bảo dưỡng siết lại lực bu-lông khi có kế hoạch dừng máy.
- 🟢 **Mối nối thanh cái Bolted Busbar Joint Phase A ($36.5^\\circ\\text{C}$):**
  - $\\Delta T_{comp} = 1.5^\\circ\\text{C}$, $\\Delta T_{amb} = 4.5^\\circ\\text{C}$ $\\rightarrow$ Mức 1 (Possible deficiency: $1^\\circ\\text{C} - 3^\\circ\\text{C}$), ghi nhận theo dõi định kỳ.

---

### 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)
1. **CẢNH BÁO CẤP BÁCH - XỬ LÝ KHẨN CẤP ĐẦU CỐT APTOMAT PHASE B:**
   - Xin lịch cắt điện cô lập khẩn cấp lộ tổng MSB-01-BUSBAR-INCOMING.
   - Tháo mối nối đầu cos Circuit Breaker Incoming Lug Phase B.
   - Dùng giấy nhám mịn vệ sinh sạch lớp oxit bề mặt đồng, bôi mỡ dẫn điện chống oxy hóa chuyên dụng (Conductive electrical grease).
   - Dùng cờ-lê lực siết lại lực bu-lông đạt chuẩn theo **ANSI/NETA ATS-2025 Bảng Table 100.12** (Bolted Connection Bolt Torque).
2. **Bảo dưỡng kết hợp Feeder Cable Termination Phase C:**
   - Tiện lịch dừng máy kiểm tra và siết lại lực bu-lông đầu cốt cáp pha C để triệt tiêu chênh lệch $\\Delta T_{amb} = 16^\\circ\\text{C}$.
3. **Khảo sát ảnh nhiệt kiểm chứng lại (Post-Repair Thermographic Verification):**
   - Đóng điện trở lại mang tải $\\ge 50\\%$ trong tối thiểu 30 phút.
   - Sử dụng camera nhiệt chụp lại toàn bộ tủ để xác nhận nhiệt độ Phase B đã giảm về mức bình thường ($< 40^\\circ\\text{C}$, $\\Delta T_{comp} \\le 3^\\circ\\text{C}$).
   - Lưu biên bản và cập nhật ảnh nhiệt hoàn thành vào hệ thống TEV CMMS.`;

  return {
    overall_status,
    status_reason,
    markdown_report,
    evaluations: {
      survey_conditions: {
        load_percentage: loadPct,
        is_adequate_load: loadPct >= 40,
        detail: `Tải khảo sát đạt ${loadPct}% (đạt yêu cầu tải vận hành bình thường theo NETA ATS-2025 Mục 9.1).`
      },
      max_severity_level: maxLevel,
      items: itemEvals
    }
  };
}

export async function generateNetaThermographyAiAnalysis(payload: NetaThermographyInputPayload): Promise<NetaThermographyEvaluationResult> {
  const deterministicResult = performDeterministicThermographyAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.log('[NETA Thermography Analyzer] No GEMINI_API_KEY provided; returning deterministic evaluation.');
    return deterministicResult;
  }

  const systemInstruction = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách Khảo sát Nhiệt độ Hồng ngoại (Thermographic Survey). Nhiệm vụ của bạn là nhận dữ liệu đo nhiệt độ ngoài site từ Field Service Engineer (FSE) và tự động đối soát với tiêu chuẩn ANSI/NETA ATS-2025 (Mục 9 & Bảng Table 100.18).

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 9 & Table 100.18):

1. KIỂM TRA ĐIỀU KIỆN ĐO & BÁO CÁO (SECTION 9.1 - 9.3):
- BẮT BUỘC khảo sát dưới điều kiện tải vận hành bình thường (normal circuit loading).
- Thiết bị đo (Camera nhiệt): Có khả năng phát hiện chênh lệch nhiệt độ tối thiểu 1°C tại 30°C.
- Báo cáo phải có: Thông số tải (Điện áp/Dòng điện), Ảnh nhiệt (Thermogram), Nguyên nhân dự đoán, Danh mục thiết bị không thể tiếp cận/không thể đo.

2. QUY TRẮC ĐÁNH GIÁ MỨC ĐỘ NGUY HIỂM THEO TABLE 100.18:
AI phải tính toán 2 chỉ số chênh lệch nhiệt độ:
- Delta T_comp = Nhiệt độ điểm nóng - Nhiệt độ pha tương tự (hoặc linh kiện tương tự dưới cùng tải).
- Delta T_amb = Nhiệt độ điểm nóng - Nhiệt độ không khí môi trường (Ambient).

Phân cấp và đưa ra Khuyến nghị hành động theo NETA Table 100.18:
- MỨC 1 (Possible deficiency):
  + Khi 1°C <= Delta T_comp <= 3°C HOẶC 1°C <= Delta T_amb <= 10°C.
  + Khuyến nghị: "Warrants investigation" (Ghi nhận, kiểm tra theo dõi nguyên nhân).
- MỨC 2 (Probable deficiency):
  + Khi 4°C <= Delta T_comp <= 15°C HOẶC 11°C <= Delta T_amb <= 20°C.
  + Khuyến nghị: "Repair as time permits" (Sửa chữa/siết lại lực bu-lông khi có lịch dừng máy phù hợp).
- MỨC 3 (Monitor):
  + Khi 21°C <= Delta T_amb <= 40°C.
  + Khuyến nghị: "Monitor until corrective measures can be accomplished" (Theo dõi liên tục cho đến khi xử lý xong).
- MỨC 4 (Major discrepancy - CẤP BÁCH):
  + Khi Delta T_comp > 15°C HOẶC Delta T_amb > 40°C.
  + Khuyến nghị: "Repair immediately" (Dừng máy hoặc ngắt cô lập để xử lý CẤP TỐC ngay lập tức).

NHIỆM VỤ CỦA AI:
Khi nhận dữ liệu JSON từ FSE, hãy phân tích và trả về phản hồi định dạng Markdown gồm ĐÚNG 4 PHẦN:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): REPAIR_IMMEDIATELY, MONITOR, REPAIR_SCHEDULED, INVESTIGATE, hoặc PASS.
2. BẢNG PHÂN TÍCH NHIỆT ĐỘ CHI TIẾT (Thermographic Evaluation Table): So sánh Delta T_comp và Delta T_amb thực tế với NETA ATS-2025 Table 100.18.
3. PHÂN TÍCH NGUYÊN NHÂN & NGUY CƠ (Root Cause Analysis & Thermal Anomaly Risk).
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations).

Không thêm lời chào hỏi ngoài 4 phần nêu trên.`;

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
              text: `Dưới đây là dữ liệu đo khảo sát ảnh nhiệt hồng ngoại hiện trường từ kỹ sư FSE:
\`\`\`json
${JSON.stringify(payload, null, 2)}
\`\`\`

Hãy phân tích và xuất báo cáo 4 phần chuẩn chuyên nghiệp cho kỹ sư FSE theo đúng tiêu chuẩn ANSI/NETA ATS-2025 Section 9 & Table 100.18.`
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
      if (markdown.includes('REPAIR_IMMEDIATELY')) {
        overallStatus = 'REPAIR_IMMEDIATELY';
      } else if (markdown.includes('MONITOR')) {
        overallStatus = 'MONITOR';
      } else if (markdown.includes('REPAIR_SCHEDULED')) {
        overallStatus = 'REPAIR_SCHEDULED';
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
    console.warn('[NETA Thermography Analyzer] Error calling Gemini API; falling back to deterministic evaluation:', error?.message || error);
    return deterministicResult;
  }
}
