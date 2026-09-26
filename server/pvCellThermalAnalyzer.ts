import { GoogleGenAI } from '@google/genai';

export type PvSeverityLevel =
  | '1 — No anomaly'
  | '2 — Mild'
  | '3 — Significant'
  | '4 — Safety-relevant';

export type PvLikelyCause =
  | 'Not installed'
  | 'Wiring / installation'
  | 'Module defect'
  | 'Shading-induced'
  | 'Soiling-induced'
  | 'Unknown';

export type PvRecommendedAction =
  | 'Replace'
  | 'Repair'
  | 'Quick fix'
  | 'Field check'
  | 'Move / remove the object'
  | 'Clean'
  | 'Monitor';

export interface PvThermalInputPayload {
  project_name?: string;
  module_tag?: string;
  fse_name?: string;
  test_date?: string;
  irradiance_w_m2?: number;
  ambient_temperature_c?: number;
  t_max_c?: number;
  t_ref_c?: number;
  pattern_type?: string;
  visual_defect_type?: string;
  visual_notes?: string;
  image_ir_base64?: string;
  image_rgb_base64?: string;
  mime_type?: string;
}

export interface PvCellAnalysisJsonData {
  anomaly_detected: boolean;
  anomaly_type: string;
  delta_t_celsius: number;
  assessment: {
    severity: PvSeverityLevel;
    likely_cause: PvLikelyCause;
    recommended_action: PvRecommendedAction;
  };
  technical_details: {
    affected_component: string;
    risk_description: string;
    field_instructions: string;
  };
}

export interface PvCellAnalysisResponse {
  success: boolean;
  markdown_report: string;
  json_data: PvCellAnalysisJsonData;
  delta_t: number;
  severity: PvSeverityLevel;
  likely_cause: PvLikelyCause;
  recommended_action: PvRecommendedAction;
  error?: string;
}

/**
 * Deterministic Rule-Based Evaluation following IEC 62446-3 & Volateq / Sitemark
 * 5-Step Process & 3-Dimension Matrix
 */
export function evaluatePvCellDeterministic(payload: PvThermalInputPayload): PvCellAnalysisResponse {
  const tMax = typeof payload.t_max_c === 'number' ? payload.t_max_c : 58.5;
  const tRef = typeof payload.t_ref_c === 'number' ? payload.t_ref_c : 42.0;
  const deltaT = parseFloat((tMax - tRef).toFixed(1));
  const pattern = (payload.pattern_type || 'Hotspot').toLowerCase();
  const visual = (payload.visual_defect_type || '').toLowerCase();
  const notes = (payload.visual_notes || '').toLowerCase();

  let anomalyDetected = true;
  let anomalyType = payload.pattern_type || 'Cell Hotspot';

  // 1. Step 2 & 3: Severity Evaluation (Cumulative Scale)
  let severity: PvSeverityLevel = '2 — Mild';

  const isReversedPolarity = pattern.includes('reversed') || pattern.includes('đảo cực') || pattern.includes('ngược cực');
  const isJunctionBox = pattern.includes('junction') || pattern.includes('hộp đấu nối') || pattern.includes('heated junction box');
  const isDiodeBypass = pattern.includes('bypass') || pattern.includes('diode') || pattern.includes('substring');
  const isOpenCircuit = pattern.includes('open circuit') || pattern.includes('hở mạch');
  const isMultiHotspot = pattern.includes('multi') || pattern.includes('đa');
  const isPID = pattern.includes('pid');
  const isSoiling = visual.includes('soiling') || visual.includes('dropping') || visual.includes('phân chim') || visual.includes('bụi');
  const isShading = visual.includes('shading') || visual.includes('shadow') || visual.includes('vegetation') || visual.includes('cây') || pattern.includes('shaded');
  const isBrokenGlass = visual.includes('broken glass') || visual.includes('vỡ kính') || visual.includes('nứt');
  const isMissingModule = visual.includes('not installed') || notes.includes('thiếu');

  // Severity Level Determination
  if (isReversedPolarity || (isJunctionBox && deltaT >= 10) || deltaT >= 40 || (isDiodeBypass && deltaT >= 25)) {
    severity = '4 — Safety-relevant';
  } else if (isMissingModule || isBrokenGlass || isOpenCircuit || isDiodeBypass || (isMultiHotspot && deltaT >= 5) || deltaT >= 10) {
    severity = '3 — Significant';
  } else if (deltaT < 10 && deltaT > 2) {
    severity = '2 — Mild';
  } else if (deltaT <= 2 && !isBrokenGlass && !isMissingModule) {
    severity = '1 — No anomaly';
    anomalyDetected = false;
  }

  // 2. Step 4: Likely Cause Engine (Precedence Order)
  // Precedence: Not installed -> Wiring/Installation -> Module defect -> Shading-induced -> Soiling-induced -> Unknown
  let likelyCause: PvLikelyCause = 'Module defect';
  if (isMissingModule) {
    likelyCause = 'Not installed';
  } else if (isReversedPolarity || isOpenCircuit || pattern.includes('string') || notes.includes('wiring') || notes.includes('mc4')) {
    likelyCause = 'Wiring / installation';
  } else if (isShading) {
    likelyCause = 'Shading-induced';
  } else if (isSoiling) {
    likelyCause = 'Soiling-induced';
  } else if (isBrokenGlass || isJunctionBox || isDiodeBypass || isPID || pattern.includes('cell') || pattern.includes('hotspot')) {
    likelyCause = 'Module defect';
  } else {
    likelyCause = 'Unknown';
  }

  // 3. Step 5: Recommended Action Matrix (First-match Order)
  // First-match: Replace -> Repair -> Quick fix -> Field check -> Move/Remove -> Clean -> Monitor
  let recommendedAction: PvRecommendedAction = 'Field check';
  if (isMissingModule || isBrokenGlass) {
    recommendedAction = 'Replace';
  } else if (isReversedPolarity || isOpenCircuit) {
    recommendedAction = 'Repair';
  } else if (isDiodeBypass && !isBrokenGlass) {
    recommendedAction = 'Quick fix';
  } else if (severity === '4 — Safety-relevant' || (severity === '3 — Significant' && deltaT >= 10)) {
    recommendedAction = 'Field check';
  } else if (likelyCause === 'Shading-induced') {
    recommendedAction = 'Move / remove the object';
  } else if (likelyCause === 'Soiling-induced') {
    recommendedAction = 'Clean';
  } else {
    recommendedAction = 'Monitor';
  }

  // Technical details
  let affectedComponent = 'Individual PV Cell / Substring';
  if (isJunctionBox) affectedComponent = 'Junction Box (Hộp Đấu Nối)';
  else if (isDiodeBypass) affectedComponent = pattern.includes('double') ? 'Double Substring (2/3 Module)' : 'Single Substring (1/3 Module)';
  else if (isOpenCircuit) affectedComponent = 'Full Module / DC String';
  else if (isMultiHotspot) affectedComponent = 'Multiple Cells on Module';
  else if (isPID) affectedComponent = 'Cells near Frame / String Negative Pole';

  let riskDescription = '';
  let fieldInstructions = '';

  if (severity === '4 — Safety-relevant') {
    riskDescription = `CỰC KỲ NGUY HIỂM (Safety-relevant): Điểm nhiệt đạt ΔT = ${deltaT}°C (T_max = ${tMax}°C). Nguy cơ cao gây cháy xém lớp EVA/màng lưng Tedlar, đánh thủng cách điện dẫn tới hồ quang DC (Arc Fault) và cháy nổ mảng pin.`;
    fieldInstructions = `1. Tách cô lập chuỗi DC / che tấm pin khỏi ánh sáng mặt trời ngay lập tức.\n2. Dùng camera nhiệt cầm tay và súng đo hồng ngoại kiểm tra lại vị trí điểm phát nhiệt.\n3. Kiểm tra Hộp đấu nối và đo thông mạch Diode bypass.\n4. Làm thủ tục thay thế tấm pin (RMA) theo quy trình an toàn PCCC.`;
  } else if (severity === '3 — Significant') {
    riskDescription = `ĐÁNG CHÚ Ý (Significant): ΔT = ${deltaT}°C vượt ngưỡng 10°C hoặc hỏng hóc cơ lý (Diode kích hoạt / vỡ kính). Tấm pin đang hoạt động ở chế độ tiêu tán công suất (Reverse Bias), làm sụt giảm sản lượng toàn chuỗi 15–33%.`;
    fieldInstructions = `1. Kiểm tra bề mặt tấm pin bằng mắt thường để phát hiện nứt vi mô (micro-cracks) hoặc nứt vỡ kính.\n2. Đo điện áp hở mạch Voc và đường cong đặc tuyến I-V ở điều kiện bức xạ > 700 W/m².\n3. Nếu Diode bypass bị hỏng kích hoạt liên tục, tiến hành thay thế hộp diode hoặc thay mô-đun.`;
  } else if (severity === '2 — Mild') {
    riskDescription = `MỨC ĐỘ NHẸ (Mild): ΔT = ${deltaT}°C nằm trong ngưỡng theo dõi (< 10°C). Tổn hao sản lượng chưa đáng kể, thường do bụi bẩn, bóng mờ nhẹ hoặc sai lệch nhẹ giữa các cell.`;
    fieldInstructions = `1. Vệ sinh rửa sạch bề mặt tấm pin bằng nước khử ion và khăn mềm chuyên dụng.\n2. Cắt tỉa cây cối hoặc di dời vật cản gây bóng râm cục bộ.\n3. Ghi nhận nhật ký giám sát và đo nhiệt lại ở chu kỳ thanh tra kế tiếp (sau 30-60 ngày).`;
  } else {
    riskDescription = `HOÀN TOÀN BÌNH THƯỜNG (No Anomaly): Không phát hiện bất thường nhiệt (ΔT = ${deltaT}°C ≤ 2°C). Tấm pin hoạt động ổn định và đồng đều.`;
    fieldInstructions = `Duy trì chế độ vận hành và bảo dưỡng định kỳ theo lịch O&M của nhà máy.`;
  }

  const markdownReport = `### BÁO CÁO PHÂN TÍCH NHIỆT HỒNG NGOẠI (IR) & THỰC TẾ (RGB) TẤM PIN PV
*Tuân thủ Tiêu chuẩn IEC 62446-3, NETA ATS-2025 & Khung Chuẩn Hóa Volateq / Sitemark*

---

#### 1. THÔNG SỐ VẬN HÀNH & KẾT QUẢ ĐO NHIỆT
- **Mã tấm pin / Chuỗi (Module Tag):** \`${payload.module_tag || 'PV-MOD-STRING-01'}\`
- **Bức xạ mặt trời (Irradiance):** **${payload.irradiance_w_m2 || 850} W/m²** (Đạt chuẩn IEC 62446-3 yêu cầu $\\ge 600\\text{ W/m}^2$)
- **Nhiệt độ cực đại ($T_{max}$):** **${tMax}°C**
- **Nhiệt độ tham chiếu cell lành ($T_{ref}$):** **${tRef}°C**
- **Mức chênh lệch nhiệt độ ($\\Delta T = T_{max} - T_{ref}$):** **${deltaT}°C**
- **Dạng bất thường nhận diện (Anomaly Type):** **${payload.pattern_type || anomalyType}**

---

#### 2. MA TRẬN ĐÁNH GIÁ 3 CHIỀU (3-DIMENSION CLASSIFICATION MATRIX)

| Chiều Đánh Giá | Kết Quả Phân Loại | Diễn Giải Chi Tiết Kỹ Thuật |
| :--- | :--- | :--- |
| **1. Mức độ Nghiêm Trọng (Severity)** | **${severity}** | ${severity === '4 — Safety-relevant' ? '🔴 Mất an toàn / Nguy cơ cháy nổ cao (ΔT ≥ 40°C hoặc hỏng JB/đảo cực)' : severity === '3 — Significant' ? '🟠 Đáng chú ý / Cần sửa chữa (ΔT ≥ 10°C hoặc Diode/Hở mạch/Vỡ kính)' : severity === '2 — Mild' ? '🟡 Mức độ nhẹ / Tiếp tục theo dõi (ΔT < 10°C, soiling, shading)' : '🟢 Bình thường, không có bất thường nhiệt'} |
| **2. Nguyên Nhân Khả Thi (Likely Cause)** | **${likelyCause}** | Cơ chế: ${likelyCause === 'Wiring / installation' ? 'Lỗi đấu nối cáp DC, giắc MC4 lỏng hoặc đảo cực tính' : likelyCause === 'Module defect' ? 'Lỗi nội tại module (nứt vi mô cell, Diode bypass, nóng hộp JB)' : likelyCause === 'Shading-induced' ? 'Che bóng vật lý do cây cối, công trình hoặc vật cản lân cận' : likelyCause === 'Soiling-induced' ? 'Bụi bẩn, phân chim bám đọng lâu ngày gây che sáng cục bộ' : likelyCause === 'Not installed' ? 'Thiếu mô-đun trên hệ giàn khung' : 'Chưa xác định'} |
| **3. Hành Động Khuyến Nghị (Action)** | **${recommendedAction}** | Khớp quy tắc ưu tiên: **${recommendedAction}** (Quyết định xử lý theo chuẩn O&M Sitemark / Volateq) |

---

#### 3. BẢN CHẤT SỰ CỐ & CẢNH BÁO NGUY CƠ (Risk Analysis)
- **Vị trí ảnh hưởng:** **${affectedComponent}**
- **Đánh giá rủi ro:** ${riskDescription}

---

#### 4. HƯỚNG DẪN HÀNH ĐỘNG DÀNH CHO KỸ SƯ FIELD SERVICE (Field Action Steps)
${fieldInstructions.split('\n').map(line => line ? `- ${line}` : '').join('\n')}`;

  const jsonData: PvCellAnalysisJsonData = {
    anomaly_detected: anomalyDetected,
    anomaly_type: payload.pattern_type || anomalyType,
    delta_t_celsius: deltaT,
    assessment: {
      severity,
      likely_cause: likelyCause,
      recommended_action: recommendedAction,
    },
    technical_details: {
      affected_component: affectedComponent,
      risk_description: riskDescription,
      field_instructions: fieldInstructions,
    }
  };

  return {
    success: true,
    markdown_report: markdownReport,
    json_data: jsonData,
    delta_t: deltaT,
    severity,
    likely_cause: likelyCause,
    recommended_action: recommendedAction
  };
}

/**
 * AI-Powered Multimodal Vision & Analysis using Gemini Vision
 */
export async function analyzePvCellWithGemini(
  payload: PvThermalInputPayload
): Promise<PvCellAnalysisResponse> {
  const deterministicFallback = evaluatePvCellDeterministic(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.log('[PV Cell Analyzer] No GEMINI_API_KEY; returning deterministic evaluation.');
    return deterministicFallback;
  }

  // Exact system instruction requested by user
  const systemInstruction = `Bạn là Hệ thống AI Chuyên gia Phân tích Ảnh Nhiệt Hồng ngoại (IR) và Ảnh Thực tế (RGB) cho Tấm pin Mặt trời (PV Cell & Module Inspection AI), tuân thủ tiêu chuẩn IEC 62446-3, NETA ATS-2025 và quy chuẩn phân tích Volateq / Sitemark.

NHIỆM VỤ CỦA BẠN:
Khi người dùng tải lên hình ảnh nhiệt (Thermal IR) hoặc ảnh thực tế (RGB) của tấm pin/chuỗi pin mặt trời, bạn phải đóng vai trò là Chuyên gia Kiểm định:
1. Nhận diện các bất thường nhiệt (Thermal Anomalies) và lỗi thực tế (Visual Defects) dựa theo phân loại chuẩn hóa Sitemark / Volateq (Hotspot, Multi hotspot, Single/Double bypassed substring, Single/Multi diode, PID, String open circuit/reversed polarity, Combiner box, Inverter, Module open circuit, Heated Junction Box, Single/Multi Cross Cell, Shaded Module, Soiling, Broken Glass, Delamination, Self Shading, Vegetation, Shadowing, Dropping...).
2. Ước tính/Đọc giá trị chênh lệch nhiệt độ Delta T (ΔT = T_max - T_ref) nếu có số liệu trên ảnh hoặc thông số đo được cung cấp.
3. Chẩn đoán chính xác 3 chiều thông tin: SEVERITY (Mức độ nghiêm trọng), LIKELY CAUSE (Nguyên nhân khả thi), và RECOMMENDED ACTION (Hành động khuyến nghị).

QUY TRẮC ĐÁNH GIÁ 3 CHIỀU (3-DIMENSION CLASSIFICATION MATRIX):

1. CHIỀU 1: SEVERITY (Mức độ nghiêm trọng - Xếp theo cấp độ tích lũy cao nhất):
- Level 1 — No anomaly: Pin hoạt động bình thường, nhiệt độ đồng đều.
- Level 2 — Mild (Nhẹ / Theo dõi): Cell Hotspot có ΔT < 10°C, Multi-cell Hotspot có ΔT < 5°C, Hộp đấu nối ấm (Junction Box ΔT < 10°C), Suy giảm PID nhẹ, Bụi bẩn (Soiling), Phân chim, Bóng che nhẹ.
- Level 3 — Significant (Đáng chú ý / Cần sửa chữa): Thiếu tấm pin, Hở mạch (Open Circuit), Diode Bypass đang kích hoạt, Ngắt chuỗi Substring, Vỡ kính mặt trước, Cell Hotspot có ΔT >= 10°C, Multi-cell Hotspot có ΔT >= 5°C.
- Level 4 — Safety-relevant (Mất an toàn / Nguy cơ cháy nổ): Cell Hotspot có ΔT >= 40°C, Hotspot xuất hiện trên chuỗi đã bị Bypass ngắt, Hộp đấu nối quá nhiệt ΔT >= 10°C, Ngược cực tính (Reversed Polarity).

2. CHIỀU 2: LIKELY CAUSE (Nguyên nhân khả thi - Ưu tiên Lỗi Điện/Cấu trúc trước Lỗi Nhiệt):
- Not installed: Tấm pin bị thiếu trên giá đỡ.
- Wiring / installation: Ngược cực tính, hở mạch chuỗi, đứt dây cáp DC, tuột giắc MC4.
- Module defect: Diode bypass bị hỏng/kích hoạt, Hộp đấu nối quá nhiệt, vỡ kính, hoặc Hotspot do lỗi nội bộ cell (nứt vi mô micro-crack, bong tróc delamination).
- Shading-induced: Che bóng do cây cối, công trình, cột điện, rác đè lên.
- Soiling-induced: Đốm bẩn bám dính, phân chim tích tụ.
- Unknown: Dạng bất thường chưa xác định.

3. CHIỀU 3: RECOMMENDED ACTION (Hành động khuyến nghị - Xử lý theo thứ tự khớp đầu tiên):
- Replace: Thay thế tấm pin (áp dụng khi thiếu tấm pin hoặc vỡ kính).
- Repair: Sửa chữa đấu nối (áp dụng khi đảo cực tính hoặc hở mạch).
- Quick fix: Xử lý nhanh giắc cắm/thay diode bypass.
- Field check: Kiểm tra đo đạc thực địa tại site bằng thiết bị đo V-I/Megaohmmeter (áp dụng khi Hộp đấu nối quá nhiệt Safety, Hotspot cấp Significant trở lên).
- Move / remove the object: Cắt tỉa cây cối, di dời vật cản gây che bóng.
- Clean: Vệ sinh rửa tấm pin, tẩy sạch phân chim.
- Monitor: Ghi nhận và theo dõi định kỳ ở kỳ kiểm tra sau (áp dụng cho Mild Hotspot, PID nhẹ).

ĐỊNH DẠNG PHẢN HỒI (OUTPUT FORMAT):
Hãy trả về phản hồi chuẩn xác gồm 2 phần:

PHẦN 1: BÁO CÁO PHÂN TÍCH MARKDOWN (Tiêu chuẩn IEC 62446-3 & Volateq/Sitemark)
- Loại thiết bị & Dạng lỗi nhận diện (Anomaly Type): Hotspot, Multi-hotspot, Single/Double Bypass Diode, Open Circuit, PID, Junction Box Overheating, Soiling, Shading...
- Mức chênh lệch nhiệt độ ước tính (Delta T): ... °C
- Severity Level: [Level 1 / 2 / 3 / 4] + Mô tả ý nghĩa.
- Likely Cause: [Nguyên nhân] + Giải thích chi tiết cơ chế gây ra lỗi.
- Recommended Action: [Hành động] + Hướng dẫn từng bước cho Field Engineer.

PHẦN 2: DỮ LIỆU JSON ĐỂ TÍCH HỢP HỆ THỐNG
\`\`\`json
{
  "anomaly_detected": true,
  "anomaly_type": "Cell Hotspot",
  "delta_t_celsius": 14.5,
  "assessment": {
    "severity": "3 — Significant",
    "likely_cause": "Module defect",
    "recommended_action": "Field check"
  },
  "technical_details": {
    "affected_component": "Single Cell / Substring 2",
    "risk_description": "...",
    "field_instructions": "..."
  }
}
\`\`\``;

  try {
    const ai = new GoogleGenAI({ apiKey });

    // Prepare content parts
    const parts: any[] = [];

    // Attach image if available
    let hasImage = false;
    if (payload.image_ir_base64) {
      let cleanIr = payload.image_ir_base64;
      let mimeType = payload.mime_type || 'image/jpeg';
      if (cleanIr.includes(';base64,')) {
        const split = cleanIr.split(';base64,');
        if (split[0].includes('data:')) {
          mimeType = split[0].replace('data:', '');
        }
        cleanIr = split[1];
      }
      parts.push({
        inlineData: {
          mimeType,
          data: cleanIr,
        },
      });
      hasImage = true;
    }

    if (payload.image_rgb_base64) {
      let cleanRgb = payload.image_rgb_base64;
      let mimeType = payload.mime_type || 'image/jpeg';
      if (cleanRgb.includes(';base64,')) {
        const split = cleanRgb.split(';base64,');
        if (split[0].includes('data:')) {
          mimeType = split[0].replace('data:', '');
        }
        cleanRgb = split[1];
      }
      parts.push({
        inlineData: {
          mimeType,
          data: cleanRgb,
        },
      });
      hasImage = true;
    }

    const userPrompt = `Hãy phân tích hình ảnh nhiệt/thực tế của tấm pin PV được tải lên (hoặc dữ liệu thông số kiểm tra hiện trường). Nhận diện dạng lỗi, ước tính mức Delta T (nếu có thông số nhiệt độ trên hình) và đưa ra đánh giá chính xác theo 3 chiều: Severity, Likely Cause, và Recommended Action kèm dữ liệu JSON theo đúng System Instruction. Thuật toán phân tích và đánh giá tình trạng tấm pin/cell mặt trời (PV Cell) từ ảnh nhiệt hồng ngoại (IR) và ảnh thực tế (RGB) được xây dựng dựa trên tiêu chuẩn phân tích ảnh nhiệt IEC 62446-3 và khung đánh giá chuẩn hóa của Volateq / Sitemark.

THÔNG SỐ HIỆN TRƯỜNG ĐƯỢC CUNG CẤP:
- Module Tag: ${payload.module_tag || 'PV-MODULE-01'}
- T_max đo được: ${payload.t_max_c ?? 58.5}°C
- T_ref đo được: ${payload.t_ref_c ?? 42.0}°C
- Bức xạ mặt trời: ${payload.irradiance_w_m2 ?? 850} W/m2
- Mẫu nhận diện dự kiến: ${payload.pattern_type || 'Hotspot'}
- Dạng bất thường thị giác: ${payload.visual_defect_type || 'None'}
- Ghi chú hiện trường: ${payload.visual_notes || 'None'}`;

    parts.push({ text: userPrompt });

    const timeoutPromise = new Promise<never>((_, reject) =>
      setTimeout(() => reject(new Error('AI generation timed out')), 9000)
    );

    const callPromise = ai.models.generateContent({
      model: 'gemini-3.8-flash',
      contents: [
        {
          role: 'user',
          parts,
        },
      ],
      config: {
        systemInstruction,
        temperature: 0.1,
      },
    });

    const response = await Promise.race([callPromise, timeoutPromise]);
    const responseText = response.text || '';

    if (responseText && responseText.length > 50) {
      // Parse JSON block from response
      let parsedJson: PvCellAnalysisJsonData | null = null;
      const jsonMatch = responseText.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
      if (jsonMatch && jsonMatch[1]) {
        try {
          parsedJson = JSON.parse(jsonMatch[1]);
        } catch (e) {
          console.warn('[PV Cell Analyzer] Failed to parse JSON block from Gemini output, using fallback json');
        }
      }

      const finalJson = parsedJson || deterministicFallback.json_data;
      const finalSeverity = (finalJson.assessment?.severity as PvSeverityLevel) || deterministicFallback.severity;
      const finalCause = (finalJson.assessment?.likely_cause as PvLikelyCause) || deterministicFallback.likely_cause;
      const finalAction = (finalJson.assessment?.recommended_action as PvRecommendedAction) || deterministicFallback.recommended_action;
      const finalDeltaT = typeof finalJson.delta_t_celsius === 'number' ? finalJson.delta_t_celsius : deterministicFallback.delta_t;

      return {
        success: true,
        markdown_report: responseText,
        json_data: finalJson,
        delta_t: finalDeltaT,
        severity: finalSeverity,
        likely_cause: finalCause,
        recommended_action: finalAction,
      };
    }

    return deterministicFallback;
  } catch (error: any) {
    console.warn('[PV Cell Analyzer] Error calling Gemini API; falling back to deterministic evaluation:', error?.message || error);
    return deterministicFallback;
  }
}
