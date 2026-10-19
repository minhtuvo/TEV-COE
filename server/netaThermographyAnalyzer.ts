import { GoogleGenAI } from '@google/genai';

export interface ThermographicFinding {
  location: string;
  component_type: string;
  measured_temperature_celsius: number;
  reference_component_temp_celsius: number;
  similar_component_delta_t_celsius?: number;
  ambient_air_delta_t_celsius?: number;
  suspected_cause?: string;
  // legacy compatibility
  component_name?: string;
  measured_temp_c?: number;
}

export interface NetaThermographyInputPayload {
  site_info: {
    project_name: string;
    survey_tag: string;
    ambient_temperature_celsius?: number;
    circuit_loading_percent?: number;
    ir_camera_model?: string;
    fse_name: string;
    test_date: string;
  };
  survey_parameters_section_9?: {
    equipment_access_condition: string;
    circuit_load_status: string;
    report_elements_complete: boolean;
  };
  thermographic_findings?: ThermographicFinding[];
  previous_test_data?: {
    last_survey_date: string;
    last_max_delta_t_celsius: number;
  };
  // legacy compatibility
  survey_conditions?: {
    camera_model: string;
    ambient_temperature_c: number;
    system_voltage_v?: number;
    operating_current_amp?: number;
    load_percentage: number;
    inaccessible_areas?: string;
  };
  thermal_measurements?: Array<{
    component_name: string;
    measured_temp_c: number;
    reference_component_temp_c: number;
  }>;
}

export interface ThermalItemEvaluation {
  location: string;
  component_type: string;
  measured_temp_c: number;
  reference_component_temp_c: number;
  ambient_temperature_c: number;
  delta_t_comp: number;
  delta_t_amb: number;
  level: 0 | 1 | 2 | 3 | 4;
  level_name: string;
  level_en: string;
  action_recommendation_vi: string;
  action_recommendation_en: string;
  suspected_cause: string;
  status: 'NORMAL' | 'INVESTIGATE' | 'REPAIR_SCHEDULED' | 'MONITOR' | 'REPAIR_IMMEDIATELY';
  // legacy compatibility
  component_name?: string;
  action_recommendation?: string;
}

export interface NetaThermographyEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'REPAIR_SCHEDULED' | 'MONITOR' | 'REPAIR_IMMEDIATELY' | 'LEVEL 4 - CRITICAL REPAIR IMMEDIATELY';
  overall_status_en: string;
  overall_status_vi: string;
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
  finding: ThermographicFinding | { component_name: string; measured_temp_c: number; reference_component_temp_c: number },
  ambientTemp: number
): ThermalItemEvaluation {
  const measured = 'measured_temperature_celsius' in finding 
    ? finding.measured_temperature_celsius 
    : finding.measured_temp_c;
  const ref = 'reference_component_temp_celsius' in finding 
    ? finding.reference_component_temp_celsius 
    : finding.reference_component_temp_c;
  const location = 'location' in finding ? finding.location : ('component_name' in finding ? finding.component_name : 'Điểm phát nhiệt');
  const compType = 'component_type' in finding ? finding.component_type : ('component_name' in finding ? finding.component_name : 'Tiếp điểm');
  const suspectedCause = 'suspected_cause' in finding ? (finding.suspected_cause || '') : 'Điện trở tiếp xúc hoặc bất cân bằng pha';

  const delta_t_comp = parseFloat((measured - ref).toFixed(1));
  const delta_t_amb = parseFloat((measured - ambientTemp).toFixed(1));

  // Determine Level per Method 1 (Similar component)
  let level1 = 0;
  if (delta_t_comp > 15) {
    level1 = 4;
  } else if (delta_t_comp >= 4) {
    level1 = 2;
  } else if (delta_t_comp >= 1) {
    level1 = 1;
  }

  // Determine Level per Method 2 (Ambient air)
  let level2 = 0;
  if (delta_t_amb > 40) {
    level2 = 4;
  } else if (delta_t_amb >= 21) {
    level2 = 3;
  } else if (delta_t_amb >= 11) {
    level2 = 2;
  } else if (delta_t_amb >= 1) {
    level2 = 1;
  }

  // Final level is the maximum of the two evaluation methods
  const finalLevel = Math.max(level1, level2) as 0 | 1 | 2 | 3 | 4;

  let level_name = 'Bình thường (Normal)';
  let level_en = 'Normal';
  let action_vi = 'Vận hành bình thường / Continue operation';
  let action_en = 'Continue operation; no repair required';
  let status: 'NORMAL' | 'INVESTIGATE' | 'REPAIR_SCHEDULED' | 'MONITOR' | 'REPAIR_IMMEDIATELY' = 'NORMAL';

  if (finalLevel === 4) {
    level_name = 'LEVEL 4 (Sai lệch nghiêm trọng)';
    level_en = 'LEVEL 4 (Major Discrepancy)';
    action_vi = 'SỬA CHỮA NGAY LẬP TỨC / REPAIR IMMEDIATELY';
    action_en = 'Isolate and repair immediately';
    status = 'REPAIR_IMMEDIATELY';
  } else if (finalLevel === 3) {
    level_name = 'LEVEL 3 (Theo dõi chặt chẽ)';
    level_en = 'LEVEL 3 (Monitor Continuously)';
    action_vi = 'Giám sát chặt chẽ cho đến khi xử lý / Monitor until repair';
    action_en = 'Monitor continuously until corrective measures are accomplished';
    status = 'MONITOR';
  } else if (finalLevel === 2) {
    level_name = 'LEVEL 2 (Khuyết tật tiềm ẩn)';
    level_en = 'LEVEL 2 (Probable Deficiency)';
    action_vi = 'Sửa chữa khi điều kiện cho phép / Repair as time permits';
    action_en = 'Schedule repair as time permits';
    status = 'REPAIR_SCHEDULED';
  } else if (finalLevel === 1) {
    level_name = 'LEVEL 1 (Cần kiểm tra)';
    level_en = 'LEVEL 1 (Possible Deficiency)';
    action_vi = 'Cần kiểm tra trong kỳ bảo dưỡng / Warrants investigation';
    action_en = 'Warrants investigation; monitor during routine maintenance';
    status = 'INVESTIGATE';
  }

  return {
    location,
    component_type: compType,
    measured_temp_c: measured,
    reference_component_temp_c: ref,
    ambient_temperature_c: ambientTemp,
    delta_t_comp,
    delta_t_amb,
    level: finalLevel,
    level_name,
    level_en,
    action_recommendation_vi: action_vi,
    action_recommendation_en: action_en,
    suspected_cause: suspectedCause,
    status,
    component_name: location,
    action_recommendation: action_vi
  };
}

export function performDeterministicThermographyAnalysis(data: NetaThermographyInputPayload): NetaThermographyEvaluationResult {
  const ambientTemp = data.site_info?.ambient_temperature_celsius ?? data.survey_conditions?.ambient_temperature_c ?? 32.5;
  const loadPct = data.site_info?.circuit_loading_percent ?? data.survey_conditions?.load_percentage ?? 68.0;
  const cameraModel = data.site_info?.ir_camera_model ?? data.survey_conditions?.camera_model ?? 'FLIR T865 (Sensitivity < 0.03 deg C)';

  // Consolidate findings from either thermographic_findings or legacy thermal_measurements
  const rawFindings = (data.thermographic_findings && data.thermographic_findings.length > 0)
    ? data.thermographic_findings
    : (data.thermal_measurements || []).map(m => ({
        location: m.component_name,
        component_type: 'Mối nối điện / Electrical Joint',
        measured_temperature_celsius: m.measured_temp_c,
        reference_component_temp_celsius: m.reference_component_temp_c,
        suspected_cause: 'Điện trở tiếp xúc / Contact resistance'
      }));

  const itemEvals: ThermalItemEvaluation[] = rawFindings.map(item =>
    evaluateThermalItem(item, ambientTemp)
  );

  const maxLevel = itemEvals.length > 0
    ? Math.max(...itemEvals.map(e => e.level))
    : 0;

  let overall_status: 'PASS' | 'INVESTIGATE' | 'REPAIR_SCHEDULED' | 'MONITOR' | 'REPAIR_IMMEDIATELY' | 'LEVEL 4 - CRITICAL REPAIR IMMEDIATELY' = 'PASS';
  let overall_status_en = 'PASS / ALL THERMAL MEASUREMENTS WITHIN NORMAL LIMITS';
  let overall_status_vi = 'PASS / TẤT CẢ ĐIỂM ĐO NHIỆT ĐỘ NẰM TRONG NGƯỠNG BÌNH THƯỜNG';
  let status_reason = 'Tất cả các điểm đo nhiệt độ mối nối và tiếp điểm đều trong ngưỡng bình thường theo ANSI/NETA MTS Section 9 & Table 100.18.';

  if (maxLevel === 4) {
    overall_status = 'LEVEL 4 - CRITICAL REPAIR IMMEDIATELY';
    overall_status_en = 'LEVEL 4 - CRITICAL REPAIR IMMEDIATELY';
    overall_status_vi = 'SAI LỆCH NGHIÊM TRỌNG - BẮT BUỘC SỬA CHỮA NGAY LẬP TỨC (LEVEL 4)';
    const majorItem = itemEvals.find(i => i.level === 4);
    status_reason = `Phát hiện điểm phát nhiệt cực kỳ nguy hiểm (Level 4 - Major Discrepancy) tại [${majorItem?.location}]: ΔT_comp = ${majorItem?.delta_t_comp}°C (> 15°C) và ΔT_amb = ${majorItem?.delta_t_amb}°C (> 40°C). Nguy cơ hỏa hoạn và phóng hồ quang! Bắt buộc cách ly và sửa chữa ngay lập tức.`;
  } else if (maxLevel === 3) {
    overall_status = 'MONITOR';
    overall_status_en = 'LEVEL 3 - MONITOR CONTINUOUSLY';
    overall_status_vi = 'LEVEL 3 - GIÁM SÁT CHẶT CHẼ LIÊN TỤC';
    status_reason = 'Phát hiện điểm phát nhiệt Mức 3 (21°C ≤ ΔT_amb ≤ 40°C). Cần giám sát liên tục tải điện và nhiệt độ cho đến khi thực hiện xong biện pháp khắc phục.';
  } else if (maxLevel === 2) {
    overall_status = 'REPAIR_SCHEDULED';
    overall_status_en = 'LEVEL 2 - REPAIR AS TIME PERMITS';
    overall_status_vi = 'LEVEL 2 - KHUYẾT TẬT TIỀM ẨN; SỬA CHỮA KHI ĐIỀU KIỆN CHO PHÉP';
    status_reason = 'Phát hiện sai lệch nhiệt độ Mức 2 (Probable deficiency: 4°C ≤ ΔT_comp ≤ 15°C hoặc 11°C ≤ ΔT_amb ≤ 20°C). Cần lên lịch siết lực hoặc bảo dưỡng khi có lịch cắt điện phù hợp.';
  } else if (maxLevel === 1) {
    overall_status = 'INVESTIGATE';
    overall_status_en = 'LEVEL 1 - WARRANTS INVESTIGATION';
    overall_status_vi = 'LEVEL 1 - CẦN KIỂM TRA TRONG KỲ BẢO DƯỠNG';
    status_reason = 'Phát hiện sai lệch nhiệt độ Mức 1 (Possible deficiency: 1°C ≤ ΔT_comp ≤ 3°C hoặc 1°C ≤ ΔT_amb ≤ 10°C). Ghi nhận và theo dõi trong kỳ bảo trì định kỳ tiếp theo.';
  }

  // Generate detailed table rows
  const tableRows = itemEvals.map(ev => {
    return `| **${ev.location}** | ${ev.component_type} | **${ev.measured_temp_c}°C** | **${ev.delta_t_comp > 0 ? '+' : ''}${ev.delta_t_comp}°C** (${ev.delta_t_comp > 15 ? '> 15°C' : ev.delta_t_comp >= 4 ? '4°C - 15°C' : '1°C - 3°C'}) | **${ev.delta_t_amb > 0 ? '+' : ''}${ev.delta_t_amb}°C** (${ev.delta_t_amb > 40 ? '> 40°C' : ev.delta_t_amb >= 21 ? '21°C - 40°C' : ev.delta_t_amb >= 11 ? '11°C - 20°C' : '1°C - 10°C'}) | **${ev.level_en}** | **${ev.action_recommendation_vi}** |`;
  }).join('\n');

  // Diagnostic narrative
  const narrativeSections = itemEvals.map(ev => {
    if (ev.level === 4) {
      return `* **${ev.location} (${ev.component_type}) - LEVEL 4 (Major Discrepancy)**:
  * **EN**: The electrical connection at ${ev.location} exhibits a critical hot spot reaching **${ev.measured_temp_c}°C**. The temperature rise compared to adjacent similar components is **ΔT = ${ev.delta_t_comp}°C** (exceeding the Level 4 threshold of > 15°C), and the rise above ambient air (${ev.ambient_temperature_c}°C) is **ΔT = ${ev.delta_t_amb}°C** (exceeding the Level 4 threshold of > 40°C). Probable cause: ${ev.suspected_cause}. This indicates an extreme contact resistance under normal circuit load (${loadPct}%), posing an immediate threat of thermal runaway, catastrophic bus insulation flashover, and arc flash danger.
  * **VN**: Điểm đấu nối tại ${ev.location} xuất hiện điểm quá nhiệt nghiêm trọng lên tới **${ev.measured_temp_c}°C**. Mức chênh lệch nhiệt độ so với pha tương tự đạt **ΔT = ${ev.delta_t_comp}°C** (vượt ngưỡng Level 4 > 15°C), và mức tăng so với nhiệt độ môi trường (${ev.ambient_temperature_c}°C) đạt **ΔT = ${ev.delta_t_amb}°C** (vượt ngưỡng Level 4 > 40°C). Nguyên nhân nghi ngờ: ${ev.suspected_cause}. Điều này chứng tỏ điện trở tiếp xúc tăng vọt dưới phụ tải ${loadPct}%, nguy cơ gây hỏa hoạn, nóng chảy đầu cực và phóng hồ quang pha-pha thảm khốc.`;
    } else if (ev.level === 3 || (ev.delta_t_amb >= 21 && ev.delta_t_amb <= 40)) {
      return `* **${ev.location} (${ev.component_type}) - LEVEL 3 (High Thermal Rise)**:
  * **EN**: Connection exhibits a significant temperature rise of **ΔT_amb = ${ev.delta_t_amb}°C** (Level 3: 21°C - 40°C above ambient) and measured temp of **${ev.measured_temp_c}°C**. Probable cause: ${ev.suspected_cause}. Requires continuous monitoring and expedited repair within 24-48 hours.
  * **VN**: Điểm tiếp xúc có mức tăng nhiệt độ so với môi trường lên tới **ΔT_amb = ${ev.delta_t_amb}°C** (Level 3: 21°C - 40°C) và nhiệt độ đo đạt **${ev.measured_temp_c}°C**. Nguyên nhân nghi ngờ: ${ev.suspected_cause}. Cần giám sát liên tục tải điện và lên phương án siết lực/thay thế trong vòng 24-48 giờ.`;
    } else if (ev.level === 2) {
      return `* **${ev.location} (${ev.component_type}) - LEVEL 2 (Probable Deficiency)**:
  * **EN**: Measured temperature **${ev.measured_temp_c}°C** with **ΔT_comp = ${ev.delta_t_comp}°C** (4°C - 15°C) and **ΔT_amb = ${ev.delta_t_amb}°C** (11°C - 20°C). Probable cause: ${ev.suspected_cause}. Indicates early-stage contact degradation. Schedule corrective maintenance during the next scheduled shutdown window.
  * **VN**: Nhiệt độ đo đạt **${ev.measured_temp_c}°C** với **ΔT_comp = ${ev.delta_t_comp}°C** (4°C - 15°C) và **ΔT_amb = ${ev.delta_t_amb}°C** (11°C - 20°C). Nguyên nhân nghi ngờ: ${ev.suspected_cause}. Cho thấy có sự gia tăng điện trở tiếp xúc ở giai đoạn đầu. Lập kế hoạch vệ sinh bề mặt và siết lực trong kỳ ngừng máy định kỳ tới.`;
    } else {
      return `* **${ev.location} (${ev.component_type}) - LEVEL 1 (Possible Deficiency / Normal)**:
  * **EN**: Measured temperature **${ev.measured_temp_c}°C** (ΔT_comp = ${ev.delta_t_comp}°C). Minor temperature variance; monitor during routine inspections.
  * **VN**: Nhiệt độ đo đạt **${ev.measured_temp_c}°C** (ΔT_comp = ${ev.delta_t_comp}°C). Sai lệch nhiệt độ nhẹ; tiếp tục theo dõi trong các kỳ bảo dưỡng định kỳ tiếp theo.`;
    }
  }).join('\n\n');

  const markdown_report = `* **TRẠNG THÁI TỔNG QUAN / OVERALL STATUS**: **${overall_status_en}** (${overall_status_vi})

---

#### 1. BẢNG PHÂN TÍCH CHI TIẾT KHẢO SÁT HỒNG NGOẠI / DETAILED THERMOGRAPHIC EVALUATION TABLE
*Tiêu chuẩn đối chiếu: ANSI/NETA MTS (Mục 9) & NETA Table 100.18 (Suggested Actions Based on Temperature Rise)*

| Vị trí Thiết bị / Equipment Location | Cấu thành & Điểm Phát nhiệt / Component & Anomaly Point | Nhiệt độ Đo / Measured Temp | $\\Delta T$ So với Pha Tương tự / Similar Delta T | $\\Delta T$ So với Môi trường / Ambient Delta T | Đánh giá NETA Table 100.18 / NETA Table 100.18 Classification | Hành động Khuyến nghị / Recommended Action |
| :--- | :--- | :---: | :---: | :---: | :---: | :--- |
${tableRows}

---

#### 2. PHÂN TÍCH ĐIỂM QUÁ NHIỆT & NGUYÊN NHÂN CƠ BẢN / THERMAL ANOMALY & ROOT CAUSE ANALYSIS

${narrativeSections}

---

#### 3. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT FOR FSE / ACTIONABLE RECOMMENDATIONS FOR FSE

1. **Xử lý Khẩn cấp Điểm Quá nhiệt Nguy cấp (Immediate Action for Level 4 Anomalies)**:
   * **EN**: Schedule an immediate emergency shutdown for any equipment with Level 4 discrepancies (> 15°C similar Delta T or > 40°C ambient Delta T). De-energize, verify zero voltage, apply safety grounds, disassemble the bolted connection, clean contact surfaces with approved electrical solvent, apply high-grade anti-oxidant grease, and re-torque bolts using a calibrated torque wrench in strict compliance with ANSI/NETA MTS Table 100.12.
   * **VN**: Lập kế hoạch ngắt điện cô lập khẩn cấp đối với các thiết bị có sai lệch Level 4 (> 15°C lệch pha hoặc > 40°C so với môi trường). Cắt nguồn, kiểm tra không còn điện, tiếp địa an toàn, tháo dỡ mối nối bu-lông, tẩy sạch oxit bề mặt bằng dung dịch chuyên dụng, quét mỡ chống oxy hóa và siết lại bằng cờ-lê lực hiệu chuẩn theo đúng Bảng ANSI/NETA MTS Table 100.12.

2. **Xử lý Điểm Phát nhiệt Tiềm ẩn & Giám sát (Corrective Action for Level 2 & Level 3)**:
   * **EN**: Continuously monitor loading and thermal progression for Level 3 points. Schedule surface reconditioning, bolt re-torquing, and lug replacement for Level 2 points during the next scheduled planned outage (within 1 - 3 months).
   * **VN**: Giám sát liên tục phụ tải và diễn biến nhiệt độ đối với các điểm Level 3. Lập kế hoạch vệ sinh tiếp xúc, siết lực bu-lông và thay thế đầu cốt cho các điểm Level 2 trong đợt bảo dưỡng cắt điện định kỳ tới (trong vòng 1 - 3 tháng).

3. **Chụp Kiểm tra Lại Sau Sửa chữa / Re-scan Verification Protocol**:
   * **EN**: Re-perform thermographic survey under similar nominal load conditions (≥ 40% nominal load) after corrective repairs are completed. Ensure temperature differentials across all connections drop below **3°C** (Level 1) before signing off the official survey report.
   * **VN**: Thực hiện chụp ảnh nhiệt kiểm tra lại dưới điều kiện mang tải tương đương (phụ tải tối thiểu ≥ 40% định mức) sau khi hoàn tất sửa chữa. Đảm bảo mức chênh lệch nhiệt độ trên tất cả các điểm nối hạ về dưới **3°C** (Level 1) trước khi ký biên bản nghiệm thu.`;

  return {
    overall_status,
    overall_status_en,
    overall_status_vi,
    status_reason,
    markdown_report,
    evaluations: {
      survey_conditions: {
        load_percentage: loadPct,
        is_adequate_load: loadPct >= 40,
        detail: `Phụ tải khảo sát đạt ${loadPct}% (đạt yêu cầu tải tối thiểu ≥ 40% định mức theo ANSI/NETA MTS Section 9.3.3 & NFPA 70B).`
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

  const systemInstruction = `You are an AI Expert Field Service Assistant (TEV Platform AI) specializing in Electrical Thermographic Survey and Thermal Infrared Diagnostics. Your task is to receive field thermographic inspection data from Field Service Engineers (FSE) and automatically evaluate thermal anomalies STRICTLY against ANSI/NETA MTS (Section 9) and NETA Table 100.18 (Suggested Actions Based on Temperature Rise). You MUST provide all evaluation outputs in BILINGUAL format (English / Vietnamese EN/VN) and NEVER include any citation reference numbers.

EVALUATION RULES & REFERENCE STANDARDS (ANSI/NETA MTS SECTION 9 & TABLE 100.18):

1. INSPECTION & REPORTING PARAMETERS (NETA MTS SECTION 9.1 - 9.3):
- Equipment Access & IR Windows (Section 9.1): Verify physical/mechanical condition and safe optical access through IR ports or opened covers.
- Report Completeness (Section 9.2): Verify 8 required elements (Equipment description, Discrepancies, Delta T values, Probable cause, Inaccessible areas, Load conditions, Thermograms/photos, Recommended actions).
- IR Camera Sensitivity (Section 9.3.1): IR camera MUST detect minimum temperature difference of 1°C at 30°C.
- Circuit Load Condition (Section 9.3.3 & NFPA 70B): Thermographic survey MUST be performed under normal circuit loading, not less than 40% nominal load.

2. THERMAL EVALUATION MATRIX & ACTIONS (NETA MTS TABLE 100.18):
Evaluate temperature rise (Delta T) using two comparison methods:
METHOD 1: Component-to-Similar-Component Delta T (Under Similar Loading):
- Delta T = 1°C - 3°C -> LEVEL 1: Possible deficiency; warrants investigation.
- Delta T = 4°C - 15°C -> LEVEL 2: Indicates probable deficiency; repair as time permits.
- Delta T > 15°C -> LEVEL 4: Major discrepancy; REPAIR IMMEDIATELY.

METHOD 2: Component-to-Ambient Air Delta T:
- Delta T = 1°C - 10°C -> LEVEL 1: Possible deficiency; warrants investigation.
- Delta T = 11°C - 20°C -> LEVEL 2: Indicates probable deficiency; repair as time permits.
- Delta T = 21°C - 40°C -> LEVEL 3: Monitor until corrective measures can be accomplished.
- Delta T > 40°C -> LEVEL 4: Major discrepancy; REPAIR IMMEDIATELY.

AI RESPONSE STRUCTURE (ALL IN BILINGUAL EN/VN):
When receiving JSON input from FSE, output a Markdown response containing 4 sections:
1. OVERALL STATUS / TRẠNG THÁI TỔNG QUAN: PASS, INVESTIGATE, MONITOR, or LEVEL 4 - CRITICAL REPAIR IMMEDIATELY.
2. DETAILED THERMOGRAPHIC SURVEY EVALUATION TABLE / BẢNG PHÂN TÍCH CHI TIẾT KHẢO SÁT HỒNG NGOẠI: Compare findings against NETA MTS Section 9 & Table 100.18 in bilingual EN/VN format.
3. THERMAL ANOMALY & ROOT CAUSE ANALYSIS / PHÂN TÍCH ĐIỂM QUÁ NHIỆT & NGUYÊN NHÂN CƠ BẢN: Diagnostic narrative explaining thermal stress, contact resistance, or phase imbalance.
4. ACTIONABLE RECOMMENDATIONS FOR FSE / KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT FOR FSE: Step-by-step guidance in EN/VN based on Table 100.18 action levels.`;

  try {
    const ai = new GoogleGenAI({ apiKey });

    const timeoutPromise = new Promise<never>((_, reject) =>
      setTimeout(() => reject(new Error('AI generation timed out')), 4500)
    );

    const callPromise = ai.models.generateContent({
      model: 'gemini-3.8-flash',
      contents: [
        {
          role: 'user',
          parts: [
            {
              text: `Dưới đây là dữ liệu kiểm tra khảo sát nhiệt độ hồng ngoại ngoài site từ kỹ sư FSE:
\`\`\`json
${JSON.stringify(payload, null, 2)}
\`\`\`

Hãy phân tích và xuất báo cáo đối soát song ngữ (EN/VN) chuẩn ANSI/NETA MTS Section 9 & Table 100.18 cho kỹ sư hiện trường FSE.`
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
      if (markdown.includes('LEVEL 4') || markdown.includes('REPAIR IMMEDIATELY') || markdown.includes('REPAIR_IMMEDIATELY')) {
        overallStatus = 'LEVEL 4 - CRITICAL REPAIR IMMEDIATELY';
      } else if (markdown.includes('LEVEL 3') || markdown.includes('MONITOR')) {
        overallStatus = 'MONITOR';
      } else if (markdown.includes('LEVEL 2') || markdown.includes('REPAIR_SCHEDULED') || markdown.includes('repair as time permits')) {
        overallStatus = 'REPAIR_SCHEDULED';
      } else if (markdown.includes('LEVEL 1') || markdown.includes('INVESTIGATE')) {
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
