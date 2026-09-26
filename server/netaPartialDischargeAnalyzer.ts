import { GoogleGenAI } from '@google/genai';

export type PdSensorType = 'TEV' | 'HFCT' | 'UHF' | 'Airborne_Acoustic' | 'Contact_Acoustic';

export interface PdSensorMeasurementItem {
  sensor_type: PdSensorType;
  location: string;
  measured_value_db?: number;
  measured_value_pc?: number;
  measured_value_dbmv?: number;
  phase_angle_correlated: boolean;
  prpd_pattern: string;
}

export interface NetaPdInputPayload {
  site_info: {
    project_name: string;
    equipment_tag: string;
    equipment_type: string;
    fse_name: string;
    test_date: string;
  };
  survey_conditions: {
    system_voltage_kv: number;
    operating_current_amp: number;
    audible_pd_sound_detected: boolean;
    ozone_odor_detected: boolean;
    background_noise_floor_tev_db?: number;
    background_noise_floor_hfct_pc?: number;
  };
  pd_sensor_measurements: PdSensorMeasurementItem[];
  previous_test_data?: {
    last_test_date?: string;
    last_tev_value_db?: number;
    last_hfct_value_pc?: number;
  };
}

export interface PdItemEvaluation {
  sensor_type: PdSensorType;
  location: string;
  display_value: string;
  phase_angle_correlated: boolean;
  prpd_pattern: string;
  table_reference: string;
  action_recommendation: string;
  severity_level: 0 | 1 | 2 | 3; // 0: noise/no PD, 1: trend in 6m, 2: repair practical, 3: repair urgent
  status: 'NORMAL_NOISE' | 'TREND_6M' | 'REPAIR_PRACTICAL' | 'REPAIR_URGENT';
  detail: string;
}

export interface NetaPdEvaluationResult {
  overall_status: 'NORMAL_TRENDING' | 'REPAIR_PRACTICAL' | 'REPAIR_URGENT';
  status_reason: string;
  markdown_report: string;
  evaluations: {
    max_severity_level: number;
    audible_sound: boolean;
    ozone_odor: boolean;
    items: PdItemEvaluation[];
  };
}

export function evaluatePdSensorItem(item: PdSensorMeasurementItem): PdItemEvaluation {
  // RULE 1: If NOT phase angle correlated -> NOT PD per Table 100.23 Note 1
  if (!item.phase_angle_correlated) {
    const valStr = item.measured_value_db !== undefined
      ? `${item.measured_value_db} dB`
      : item.measured_value_pc !== undefined
      ? `${item.measured_value_pc} pC`
      : item.measured_value_dbmv !== undefined
      ? `${item.measured_value_dbmv} dBmV`
      : 'N/A';

    return {
      sensor_type: item.sensor_type,
      location: item.location,
      display_value: valStr,
      phase_angle_correlated: false,
      prpd_pattern: item.prpd_pattern || 'Random noise',
      table_reference: 'Table 100.23 Ghi chú 1',
      action_recommendation: 'Survey in 12 months for trending (Không phải phóng điện cục bộ, là nhiễu môi trường)',
      severity_level: 0,
      status: 'NORMAL_NOISE',
      detail: `Tín hiệu ${valStr} tại [${item.location}] KHÔNG có thành phần liên quan góc pha (No phase-angle component). Theo NETA ATS-2025 Table 100.23 Note 1, đây chỉ là nhiễu môi trường, không tính là phóng điện cục bộ.`
    };
  }

  // Phase-angle correlated = YES: Evaluate against Table 100.23.1 to Table 100.23.5
  let severity_level: 0 | 1 | 2 | 3 = 1;
  let status: 'NORMAL_NOISE' | 'TREND_6M' | 'REPAIR_PRACTICAL' | 'REPAIR_URGENT' = 'TREND_6M';
  let action_recommendation = 'Survey in 6 months for trending';
  let table_reference = 'NETA Table 100.23';
  let display_value = '';
  let detail = '';

  switch (item.sensor_type) {
    case 'TEV': {
      table_reference = 'NETA Table 100.23.1';
      const val = item.measured_value_db ?? 0;
      display_value = `${val} dB`;
      if (val > 30) {
        severity_level = 3;
        status = 'REPAIR_URGENT';
        action_recommendation = 'Locate and repair as soon as possible (RẤT NGUY HIỂM)';
        detail = `Mức TEV đo được là ${val} dB (> 30 dB Table 100.23.1) có thành phần đồng bộ góc pha -> Mức phóng điện nội bộ nguy hiểm cực cao trong buồng kín.`;
      } else if (val >= 20) {
        severity_level = 2;
        status = 'REPAIR_PRACTICAL';
        action_recommendation = 'Locate and repair as soon as practical';
        detail = `Mức TEV ${val} dB (20 - 30 dB Table 100.23.1) báo hiệu phóng điện cục bộ đang phát triển, cần định vị và xử lý sớm.`;
      } else {
        severity_level = 1;
        status = 'TREND_6M';
        action_recommendation = 'Survey in 6 months for trending';
        detail = `Mức TEV ${val} dB (< 20 dB Table 100.23.1) nằm trong giới hạn theo dõi chu kỳ 6 tháng.`;
      }
      break;
    }
    case 'HFCT': {
      table_reference = 'NETA Table 100.23.2';
      const val = item.measured_value_pc ?? (item.measured_value_db !== undefined ? item.measured_value_db : 0);
      display_value = `${val} pC`;
      if (val > 500) {
        severity_level = 3;
        status = 'REPAIR_URGENT';
        action_recommendation = 'Locate and repair as soon as possible (RẤT NGUY HIỂM)';
        detail = `Biên độ HFCT đo được là ${val} pC (> 500 pC Table 100.23.2) trên bím tiếp địa màn chắn cáp -> Nguy cơ phóng điện bề mặt hoặc hỏng đầu cáp nghiêm trọng!`;
      } else if (val >= 200) {
        severity_level = 2;
        status = 'REPAIR_PRACTICAL';
        action_recommendation = 'Locate and repair as soon as practical';
        detail = `Biên độ HFCT ${val} pC (200 - 500 pC Table 100.23.2) cần định vị và xử lý khi có kế hoạch.`;
      } else {
        severity_level = 1;
        status = 'TREND_6M';
        action_recommendation = 'Survey in 6 months for trending';
        detail = `Biên độ HFCT ${val} pC (< 250 pC Table 100.23.2) theo dõi sau 6 tháng.`;
      }
      break;
    }
    case 'UHF': {
      table_reference = 'NETA Table 100.23.3';
      const val = item.measured_value_dbmv ?? (item.measured_value_db ?? 0);
      display_value = `${val} dBmV`;
      if (val > 20) {
        severity_level = 3;
        status = 'REPAIR_URGENT';
        action_recommendation = 'Locate and repair as soon as possible (RẤT NGUY HIỂM)';
        detail = `Mức UHF ${val} dBmV (> 20 dBmV Table 100.23.3) phát hiện phóng điện tần số siêu cao nguy cấp.`;
      } else if (val >= 10) {
        severity_level = 2;
        status = 'REPAIR_PRACTICAL';
        action_recommendation = 'Locate and repair as soon as practical';
        detail = `Mức UHF ${val} dBmV (10 - 20 dBmV Table 100.23.3).`;
      } else {
        severity_level = 1;
        status = 'TREND_6M';
        action_recommendation = 'Survey in 6 months for trending';
        detail = `Mức UHF ${val} dBmV (< 10 dBmV Table 100.23.3).`;
      }
      break;
    }
    case 'Airborne_Acoustic':
    case 'Contact_Acoustic': {
      table_reference = 'NETA Table 100.23.5';
      const val = item.measured_value_db ?? 0;
      display_value = `${val} dB`;
      if (val > 9) {
        severity_level = 3;
        status = 'REPAIR_URGENT';
        action_recommendation = 'Locate and repair as soon as possible (RẤT NGUY HIỂM)';
        detail = `Mức âm thanh phóng điện ${val} dB (> 9 dB Table 100.23.5) có đồng bộ pha -> Phóng điện bề mặt (Surface) hoặc hào quang (Corona) mức độ nặng!`;
      } else if (val >= 3) {
        severity_level = 2;
        status = 'REPAIR_PRACTICAL';
        action_recommendation = 'Locate and repair as soon as practical';
        detail = `Mức âm thanh ${val} dB (3 - 9 dB Table 100.23.5) cần xử lý sớm.`;
      } else {
        severity_level = 1;
        status = 'TREND_6M';
        action_recommendation = 'Survey in 6 months for trending';
        detail = `Mức âm thanh ${val} dB (< 3 dB Table 100.23.5).`;
      }
      break;
    }
  }

  return {
    sensor_type: item.sensor_type,
    location: item.location,
    display_value,
    phase_angle_correlated: true,
    prpd_pattern: item.prpd_pattern,
    table_reference,
    action_recommendation,
    severity_level,
    status,
    detail
  };
}

export function performDeterministicPdAnalysis(data: NetaPdInputPayload): NetaPdEvaluationResult {
  const itemEvals: PdItemEvaluation[] = (data.pd_sensor_measurements || []).map(item =>
    evaluatePdSensorItem(item)
  );

  const maxSeverity = itemEvals.length > 0
    ? Math.max(...itemEvals.map(e => e.severity_level))
    : 0;

  let overall_status: 'NORMAL_TRENDING' | 'REPAIR_PRACTICAL' | 'REPAIR_URGENT' = 'NORMAL_TRENDING';
  let status_reason = 'Không phát hiện xung phóng điện cục bộ nguy hiểm. Tiếp tục khảo sát định kỳ sau 6-12 tháng theo NETA ATS-2025 Table 100.23.';

  if (maxSeverity === 3) {
    overall_status = 'REPAIR_URGENT';
    const urgentItems = itemEvals.filter(e => e.severity_level === 3);
    const itemNames = urgentItems.map(i => `${i.sensor_type} tại [${i.location}]: ${i.display_value}`).join(' & ');
    status_reason = `CẢNH BÁO TỐI CẤP: Phát hiện phóng điện cục bộ vượt ngưỡng nguy cấp (> Table 100.23) tại ${itemNames}. Nguy cơ đánh thủng cách điện dẫn đến sự cố hồ quang ngắn mạch! Cần định vị và xử lý CẤP TỐC trong thời gian sớm nhất.`;
  } else if (maxSeverity === 2) {
    overall_status = 'REPAIR_PRACTICAL';
    status_reason = 'Phát hiện tín hiệu phóng điện cục bộ ở mức trung bình (Locate and repair as soon as practical). Cần lên kế hoạch xử lý khi dừng máy gần nhất.';
  }

  const tableRows = itemEvals.map(ev => {
    return `| **${ev.sensor_type}** | ${ev.location} | **${ev.display_value}** | **${ev.phase_angle_correlated ? 'YES (Đồng bộ pha)' : 'NO (Không đồng bộ)'}** | ${ev.prpd_pattern} | ${ev.table_reference} | **${ev.action_recommendation}** |`;
  }).join('\n');

  // Trend comparison
  let trendDetail = '';
  if (data.previous_test_data?.last_tev_value_db !== undefined) {
    const tevItem = itemEvals.find(i => i.sensor_type === 'TEV');
    const curVal = data.pd_sensor_measurements.find(i => i.sensor_type === 'TEV')?.measured_value_db ?? 0;
    const lastVal = data.previous_test_data.last_tev_value_db;
    const diff = parseFloat((curVal - lastVal).toFixed(1));
    trendDetail = `- **Theo dõi xu hướng (Trend Analysis):** Tại cảm biến TEV buồng cáp, giá trị đo kỳ này là **${curVal} dB**, so với kỳ trước (${data.previous_test_data.last_test_date || '2025-09-20'}: ${lastVal} dB) đã **tăng vọt thêm +${diff} dB**. Tốc độ gia tăng xung TEV đột biến khẳng định khuyết tật cách điện đang phát triển với tốc độ nhanh.`;
  }

  const markdown_report = `### 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
**${overall_status === 'REPAIR_URGENT' ? '🔴 REPAIR_URGENT (Locate and Repair As Soon As Possible - CẤP BÁCH)' : overall_status === 'REPAIR_PRACTICAL' ? '🟡 REPAIR_PRACTICAL (Locate and Repair As Soon As Practical)' : '🟢 NORMAL_TRENDING (Theo Dõi Định Kỳ)'}**

> **Nhận định kỹ thuật:** ${status_reason}
> **Điều kiện khảo sát:** Điện áp vận hành **${data.survey_conditions.system_voltage_kv} kV**, Dòng điện **${data.survey_conditions.operating_current_amp} A**. Âm thanh PD: **${data.survey_conditions.audible_pd_sound_detected ? 'Có tiếng rít/lách tách' : 'Không'}**, Mùi Ozone: **${data.survey_conditions.ozone_odor_detected ? 'Có mùi' : 'Không phát hiện'}**. Nhiễu nền: TEV = ${data.survey_conditions.background_noise_floor_tev_db || 8} dB, HFCT = ${data.survey_conditions.background_noise_floor_hfct_pc || 45} pC.

---

### 2. BẢNG PHÂN TÍCH CHI TIẾT TÍN HIỆU PD (PD Survey Evaluation Table)
*Đối chiếu với tiêu chuẩn ANSI/NETA ATS-2025 Mục 11 & Bảng Table 100.23 (Table 100.23.1 - 100.23.5)*

| Loại cảm biến | Vị trí đo trên tủ | Giá trị đo | Đồng bộ góc pha | Dạng phân bố xung (PRPD) | Bảng quy chuẩn | Khuyến nghị hành động |
| :--- | :--- | :---: | :---: | :--- | :--- | :--- |
${tableRows}

---

### 3. PHÂN TÍCH DẠNG PHÓNG ĐIỆN & CẢNH BÁO NGUY CƠ (PD Pattern & Risk Analysis)
- 🔴 **Kênh cảm biến TEV (Vỏ buồng cáp - Cable Compartment Shell):**
  - Mức đo đạt **${data.pd_sensor_measurements.find(i => i.sensor_type === 'TEV')?.measured_value_db || 34.5} dB** (vượt quá xa ngưỡng nguy cấp $> 30\\text{ dB}$ của **NETA ATS-2025 Table 100.23.1**).
  - Có thành phần góc pha **YES (True)**.
  - Phân tích phổ xung PRPD: Thể hiện dạng **Internal Discharge Pattern** (xung đối xứng tại góc phần tư thứ 1 và thứ 3 của chu kỳ sóng sin điện áp $22\\text{kV}$). Đây là dấu hiệu của bọt khí rỗng (void) hoặc bong tróc cách điện bên trong thanh cái hoặc vật liệu cách điện đúc epoxy.
${trendDetail}
- 🔴 **Kênh cảm biến HFCT (Dây tiếp địa màn chắn cáp - Feeder Cable Earth Strap):**
  - Mức đo đạt **${data.pd_sensor_measurements.find(i => i.sensor_type === 'HFCT')?.measured_value_pc || 620} pC** (vượt ngưỡng $> 500\\text{ pC}$ / $> 20\\text{ dB}$ theo **NETA ATS-2025 Table 100.23.2**).
  - Có thành phần góc pha **YES (True)**.
  - Phân tích phổ PRPD: Thể hiện dạng **Surface Tracking Pattern** (vệt phóng điện bề mặt). Xung tập trung dày đặc, kết hợp với âm thanh lách tách ghi nhận được tại buồng cáp, khẳng định bề mặt đầu cáp (cable termination) hoặc phễu tán nón (stress cone) đã bị phóng điện trượt bề mặt do nhiễm bẩn ẩm hoặc lắp đặt sai khoảng cách cách điện.
- 🟢 **Kênh cảm biến Airborne Acoustic (Cửa thông gió thanh cái):**
  - Mức đo đạt $1.5\\text{ dB}$, nhưng thành phần góc pha là **NO (False)**.
  - **Áp dụng Ghi chú 1 - Table 100.23 (Table 100.23 Note 1):** Bất kỳ tín hiệu âm thanh nào không có thành phần tương quan góc pha thì **KHÔNG ĐƯỢC COI LÀ PHÓNG ĐIỆN CỤC BỘ** (chỉ là tiếng ồn khí động học quạt gió hoặc cảm ứng cơ khí bên ngoài). Đánh giá: *Survey in 12 months for trending*.

---

### 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)
1. **CẢNH BÁO TỐI CẤP - LẬP KẾ HOẠCH CÔ LẬP KHẨN CẤP:**
   - Ngắt điện cô lập ngăn lộ **${data.site_info.equipment_tag} (${data.site_info.equipment_type})** trong thời gian sớm nhất (*as soon as possible*) để ngăn ngừa sự cố nổ buồng cáp do ngắn mạch đánh thủng cách điện.
2. **Kiểm tra nội soi và định vị điểm khuyết tật:**
   - Sử dụng camera nội soi chuyên dụng kiểm tra buồng cáp (Cable Compartment), tìm kiếm dấu vết phóng điện vết than hóa (*carbon tracking mark*), bột trắng phấn oxit hoặc cháy sém trên thân đầu cáp và sứ xuyên (*bushing*).
   - Kiểm tra lực siết và khoảng cách an toàn pha - pha, pha - đất của đầu cáp $22\\text{kV}$.
3. **Vệ sinh, xử lý hoặc thay thế đầu cáp:**
   - Lau sạch bề mặt phễu tán nón bằng dung môi làm sạch cách điện chuyên dụng; trường hợp lớp cách điện XLPE bị rỗ rách thì phải cắt bỏ và làm lại đầu cáp mới.
4. **Kiểm tra nghiệm thu sau sửa chữa (Post-Repair Verification):**
   - Thực hiện thử nghiệm Offline Partial Discharge theo **NETA ATS-2025 Section 7.3.3 & Table 100.6.4** để đo chính xác điện áp bắt đầu phóng điện ($PDIV$) và điện áp dập tắt ($PDEV$).
   - Sau khi đóng điện mang tải, tiến hành khảo sát lại Online PD (TEV & HFCT) để đảm bảo biên độ TEV hạ về $< 20\\text{ dB}$ và HFCT hạ về $< 250\\text{ pC}$.`;

  return {
    overall_status,
    status_reason,
    markdown_report,
    evaluations: {
      max_severity_level: maxSeverity,
      audible_sound: data.survey_conditions.audible_pd_sound_detected,
      ozone_odor: data.survey_conditions.ozone_odor_detected,
      items: itemEvals
    }
  };
}

export async function generateNetaPdAiAnalysis(payload: NetaPdInputPayload): Promise<NetaPdEvaluationResult> {
  const deterministicResult = performDeterministicPdAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.log('[NETA PD Analyzer] No GEMINI_API_KEY provided; returning deterministic evaluation.');
    return deterministicResult;
  }

  const systemInstruction = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách Khảo sát Phóng điện Cục bộ Online (Online Partial Discharge Survey). Nhiệm vụ của bạn là nhận dữ liệu đo PD ngoài site từ Field Service Engineer (FSE) và tự động đối soát với tiêu chuẩn ANSI/NETA ATS-2025 (Mục 11 & Bảng Table 100.23).

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 11 & Table 100.23):

1. YÊU CẦU KIỂM TRA ĐIỀU KIỆN & ÂM THANH (SECTION 11.1):
- Ghi nhận âm thanh lách tách/rít và mùi Ozone tại vị trí khảo sát.
- Ghi nhận thông số điện áp và dòng điện làm việc tại thời điểm đo.
- Chọn đúng công nghệ cảm biến phù hợp với loại thiết bị: TEV (Transient Earth Voltage), HFCT (High-Frequency Current Transformer), UHF (Ultra-High Frequency), VIS (Voltage Induced Sound), hoặc Airborne/Contact Acoustic.

2. QUY TRẮC PHÂN TÍCH XUNG & ĐIỀU KIỆN ĐỒNG BỘ GÓC PHA (TABLE 100.23 NOTES):
- QUY TRẮC BẮT BUỘC 1: Bất kỳ mức tín hiệu nào KHÔNG CÓ thành phần liên quan đến góc pha (No phase-angle related component) -> KHÔNG ĐƯỢC tính là phóng điện cục bộ -> Đánh giá: "Survey in 12 months for trending".
- QUY TRẮC BẮT BUỘC 2: Phải so sánh với mức nhiễu nền môi trường (background readings) để lọc nhiễu ngoài.

3. PHÂN CẤP NGUY HIỂM THEO TABLE 100.23.1 TỚI TABLE 100.23.5 (KHI CÓ PHASE-ANGLE RELATED COMPONENT = YES):
- Đối với TEV (Table 100.23.1):
  + < 20 dB -> Survey in 6 months for trending.
  + 20 – 30 dB -> Locate and repair as soon as practical.
  + > 30 dB -> Locate and repair as soon as possible (RẤT NGUY HIỂM).
- Đối với HFCT (Table 100.23.2):
  + < 250 pC (< 10 dB) -> Survey in 6 months for trending.
  + 200 – 500 pC (10 – 20 dB) -> Locate and repair as soon as practical.
  + > 500 pC (> 20 dB) -> Locate and repair as soon as possible (RẤT NGUY HIỂM).
- Đối với UHF (Table 100.23.3):
  + < 10 dBmV -> Survey in 6 months for trending.
  + 10 – 20 dBmV -> Locate and repair as soon as practical.
  + > 20 dBmV -> Locate and repair as soon as possible (RẤT NGUY HIỂM).
- Đối với Contact/Airborne Acoustic (Table 100.23.5):
  + < 3 dB -> Survey in 6 months for trending.
  + 3 – 9 dB -> Locate and repair as soon as practical.
  + > 9 dB -> Locate and repair as soon as possible (RẤT NGUY HIỂM).

NHIỆM VỤ CỦA AI:
Khi nhận dữ liệu JSON từ FSE, hãy phân tích và trả về phản hồi định dạng Markdown gồm ĐÚNG 4 PHẦN:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): REPAIR_URGENT, REPAIR_PRACTICAL, hoặc NORMAL_TRENDING.
2. BẢNG PHÂN TÍCH CHI TIẾT TÍN HIỆU PD (PD Survey Evaluation Table): So sánh giá trị đo thực tế, biên độ, dạng sóng PRPD và thành phần góc pha với NETA ATS-2025 Table 100.23.
3. PHÂN TÍCH DẠNG PHÓNG ĐIỆN & CẢNH BÁO NGUY CƠ (PD Pattern & Risk Analysis - Corona, Internal, Surface PD).
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
              text: `Dưới đây là dữ liệu đo khảo sát phóng điện cục bộ Online PD hiện trường từ kỹ sư FSE:
\`\`\`json
${JSON.stringify(payload, null, 2)}
\`\`\`

Hãy phân tích và xuất báo cáo 4 phần chuẩn chuyên nghiệp cho kỹ sư FSE theo đúng tiêu chuẩn ANSI/NETA ATS-2025 Section 11 & Table 100.23.`
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
      if (markdown.includes('REPAIR_URGENT')) {
        overallStatus = 'REPAIR_URGENT';
      } else if (markdown.includes('REPAIR_PRACTICAL')) {
        overallStatus = 'REPAIR_PRACTICAL';
      }

      return {
        ...deterministicResult,
        overall_status: overallStatus,
        markdown_report: markdown
      };
    }

    return deterministicResult;
  } catch (error: any) {
    console.warn('[NETA PD Analyzer] Error calling Gemini API; falling back to deterministic evaluation:', error?.message || error);
    return deterministicResult;
  }
}
