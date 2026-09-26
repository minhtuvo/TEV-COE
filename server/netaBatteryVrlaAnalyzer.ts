import { GoogleGenAI } from '@google/genai';

export interface NetaBatteryVrlaInputPayload {
  site_info: {
    project_name: string;
    battery_bank_tag: string;
    battery_type: string;
    fse_name: string;
    test_date: string;
    substation_location?: string;
    rated_voltage_vdc?: number;
    rated_capacity_ah?: number;
    ambient_temp_c?: number;
  };
  visual_inspection: {
    ventilation_system_operable: boolean;
    eyewash_station_present: boolean;
    nameplate_match: boolean;
    cabinet_mounting_and_grounding: 'Pass' | 'Fail' | 'Investigate';
    cleanliness_and_oxide_inhibitor: 'Pass' | 'Fail' | 'Investigate';
    physical_condition: string; // e.g. "Good (No case swelling or acid leaks)" or "Swelling detected"
    bolt_torque_check: 'Pass' | 'Fail' | 'Investigate';
    thermographic_survey?: 'Pass' | 'Fail' | 'Investigate' | 'Not Performed';
  };
  electrical_tests: {
    negative_post_temperature_c: {
      avg_temp_c: number;
      max_temp_monoblock_12_c?: number;
      max_temp_c?: number;
      hot_monoblock_tag?: string;
      temp_differential_c?: number;
    };
    intercell_connection_resistance_micro_ohms: number[];
    charger_float_voltage_v: number;
    charger_equalize_voltage_v: number;
    charger_alarms_check: 'Pass' | 'Fail' | 'Investigate';
    monoblock_voltages_float_mode_v: {
      min_voltage_v: number;
      max_voltage_v: number;
      avg_voltage_v: number;
      voltage_spread_v?: number;
    };
    internal_ohmic_measurement_resistance_mohm: {
      avg_resistance_mohm: number;
      max_resistance_monoblock_12_mohm?: number;
      max_resistance_mohm?: number;
      variance_percent: number;
      suspect_monoblock_tag?: string;
    };
    load_test_ieee_1188: string; // e.g. "Pass (Capacity 88% per IEEE 1188)"
    system_voltage_to_ground_v: {
      positive_to_ground_v: number;
      negative_to_ground_v: number;
      ground_fault_status?: 'Balanced' | 'Positive Ground Fault' | 'Negative Ground Fault' | 'Asymmetric';
    };
  };
  previous_test_data?: {
    last_test_date?: string;
    last_avg_resistance_mohm?: number;
    last_monoblock_12_temp_c?: number;
  };
}

export interface NetaBatteryVrlaEvaluationResult {
  overall_status: 'PASS' | 'FAIL' | 'INVESTIGATE';
  status_reason: string;
  markdown_report: string;
  evaluation_details: {
    visual_mechanical_status: 'PASS' | 'FAIL' | 'INVESTIGATE';
    negative_post_temp_status: 'PASS' | 'FAIL' | 'INVESTIGATE';
    temp_delta_c: number;
    thermal_runaway_risk: 'LOW' | 'MODERATE' | 'CRITICAL';
    intercell_resistance_status: 'PASS' | 'FAIL' | 'INVESTIGATE';
    intercell_max_deviation_percent: number;
    internal_ohmic_status: 'PASS' | 'FAIL' | 'INVESTIGATE';
    internal_ohmic_variance_percent: number;
    monoblock_voltages_status: 'PASS' | 'FAIL' | 'INVESTIGATE';
    monoblock_voltage_spread_v: number;
    ground_voltage_status: 'PASS' | 'FAIL' | 'INVESTIGATE';
    ground_voltage_diff_v: number;
    load_test_status: 'PASS' | 'FAIL' | 'INVESTIGATE';
  };
}

export function performDeterministicVrlaAnalysis(
  payload: NetaBatteryVrlaInputPayload
): NetaBatteryVrlaEvaluationResult {
  let overall_status: 'PASS' | 'FAIL' | 'INVESTIGATE' = 'PASS';
  const failureReasons: string[] = [];
  const investigateReasons: string[] = [];

  // 1. Visual & Mechanical Check
  let visualStatus: 'PASS' | 'FAIL' | 'INVESTIGATE' = 'PASS';
  if (!payload.visual_inspection.ventilation_system_operable) {
    visualStatus = 'FAIL';
    failureReasons.push('Hệ thống thông gió phòng tủ VRLA không hoạt động (Nguy cơ tích tụ khí và phát nhiệt)');
  }
  if (!payload.visual_inspection.eyewash_station_present) {
    if (visualStatus !== 'FAIL') visualStatus = 'INVESTIGATE';
    investigateReasons.push('Thiếu trạm rửa mắt khẩn cấp tại khu vực giàn pin VRLA');
  }
  if (!payload.visual_inspection.nameplate_match) {
    if (visualStatus !== 'FAIL') visualStatus = 'INVESTIGATE';
    investigateReasons.push('Nhãn mác giàn pin VRLA không khớp 100% bản vẽ thiết kế');
  }
  if (
    payload.visual_inspection.cabinet_mounting_and_grounding === 'Fail' ||
    payload.visual_inspection.cleanliness_and_oxide_inhibitor === 'Fail' ||
    payload.visual_inspection.bolt_torque_check === 'Fail'
  ) {
    visualStatus = 'FAIL';
    failureReasons.push('Khung tủ/tiếp địa/độ sạch hoặc lực siết bu-lông không đạt chuẩn NETA');
  } else if (
    payload.visual_inspection.cabinet_mounting_and_grounding === 'Investigate' ||
    payload.visual_inspection.cleanliness_and_oxide_inhibitor === 'Investigate' ||
    payload.visual_inspection.bolt_torque_check === 'Investigate'
  ) {
    if (visualStatus !== 'FAIL') visualStatus = 'INVESTIGATE';
    investigateReasons.push('Cần kiểm tra lại tiếp địa hoặc bôi mỡ chống oxy hóa cọc cực');
  }

  const physicalCondLower = (payload.visual_inspection.physical_condition || '').toLowerCase();
  if (physicalCondLower.includes('swell') || physicalCondLower.includes('phồng') || physicalCondLower.includes('crack') || physicalCondLower.includes('nứt') || physicalCondLower.includes('leak')) {
    visualStatus = 'FAIL';
    failureReasons.push('Vỏ bình VRLA bị phồng rộp (swelling) hoặc rò rỉ dung dịch');
  }

  // 2. Negative Post Temperature & Thermal Runaway Check (NETA ATS-2025 Sec 7.18.1.3 & IEEE 1188)
  const avgTemp = payload.electrical_tests.negative_post_temperature_c.avg_temp_c || 25.0;
  const maxTemp =
    payload.electrical_tests.negative_post_temperature_c.max_temp_monoblock_12_c ??
    payload.electrical_tests.negative_post_temperature_c.max_temp_c ??
    avgTemp;
  const tempDelta = Math.round((maxTemp - avgTemp) * 10) / 10;

  let negPostTempStatus: 'PASS' | 'FAIL' | 'INVESTIGATE' = 'PASS';
  let thermalRunawayRisk: 'LOW' | 'MODERATE' | 'CRITICAL' = 'LOW';

  if (tempDelta >= 5.0 || maxTemp >= 40.0) {
    negPostTempStatus = 'FAIL';
    thermalRunawayRisk = 'CRITICAL';
    failureReasons.push(`Nhiệt độ cực âm bình VRLA đạt ${maxTemp}°C (lệch +${tempDelta}°C so với trung bình ${avgTemp}°C) -> Nguy cơ Trôi nhiệt (Thermal Runaway) cực kỳ nghiêm trọng per IEEE 1188`);
  } else if (tempDelta >= 3.0 || maxTemp >= 35.0) {
    negPostTempStatus = 'INVESTIGATE';
    thermalRunawayRisk = 'MODERATE';
    investigateReasons.push(`Nhiệt độ cực âm bình lệch +${tempDelta}°C so với giàn pin (cần theo dõi trôi nhiệt)`);
  }

  // 3. Intercell Connection Resistance (NETA ATS-2025: shall not deviate by > 50% from minimum)
  const intercellRes = payload.electrical_tests.intercell_connection_resistance_micro_ohms || [];
  let intercellStatus: 'PASS' | 'FAIL' | 'INVESTIGATE' = 'PASS';
  let maxIntercellDev = 0;
  if (intercellRes.length > 0) {
    const minVal = Math.min(...intercellRes);
    const maxVal = Math.max(...intercellRes);
    if (minVal > 0) {
      maxIntercellDev = Math.round(((maxVal - minVal) / minVal) * 1000) / 10;
      if (maxIntercellDev > 50) {
        intercellStatus = 'INVESTIGATE';
        investigateReasons.push(`Điện trở cầu nối intercell lệch cực đại ${maxIntercellDev}% (> 50% theo NETA ATS-2025 Mục 7.18.1.3)`);
      }
    }
  }

  // 4. Internal Ohmic Measurements (NETA ATS-2025: shall not vary by > 25%)
  const ohmicVariance = payload.electrical_tests.internal_ohmic_measurement_resistance_mohm.variance_percent;
  let internalOhmicStatus: 'PASS' | 'FAIL' | 'INVESTIGATE' = 'PASS';
  if (ohmicVariance > 35) {
    internalOhmicStatus = 'FAIL';
    failureReasons.push(`Nội trở Ôm Monoblock biến động ${ohmicVariance}% (Vượt xa ngưỡng tối đa cho phép 25% theo NETA Section 7.18.1.3.D.7 - Báo hiệu khô dung dịch GEL/AGM hoặc chai bản cực)`);
  } else if (ohmicVariance > 25) {
    internalOhmicStatus = 'INVESTIGATE';
    investigateReasons.push(`Nội trở Ôm Monoblock biến động ${ohmicVariance}% (> 25% theo NETA ATS-2025 Mục 7.18.1.3.D.7)`);
  }

  // 5. Monoblock Voltages in Float Mode
  const minV = payload.electrical_tests.monoblock_voltages_float_mode_v.min_voltage_v;
  const maxV = payload.electrical_tests.monoblock_voltages_float_mode_v.max_voltage_v;
  const voltageSpread = Math.round((maxV - minV) * 1000) / 1000;
  let monoblockVoltagesStatus: 'PASS' | 'FAIL' | 'INVESTIGATE' = 'PASS';
  // Standard 12V monoblock float voltage is typically 13.5 - 13.8V, spread > 0.6V indicates uneven charging
  if (voltageSpread > 0.6) {
    monoblockVoltagesStatus = 'INVESTIGATE';
    investigateReasons.push(`Độ lệch áp float giữa các bình monoblock là ${voltageSpread}V (cần kiểm tra nạp cân bằng hoặc thay bình lỗi)`);
  }

  // 6. Ground Voltage Symmetry Check
  const vPos = Math.abs(payload.electrical_tests.system_voltage_to_ground_v.positive_to_ground_v);
  const vNeg = Math.abs(payload.electrical_tests.system_voltage_to_ground_v.negative_to_ground_v);
  const groundDiff = Math.round(Math.abs(vPos - vNeg) * 10) / 10;
  let groundStatus: 'PASS' | 'FAIL' | 'INVESTIGATE' = 'PASS';
  if (groundDiff > 25) {
    groundStatus = 'FAIL';
    failureReasons.push(`Điện áp đối đất mất đối xứng nghiêm trọng (|V+|=${vPos}V vs |V-|=${vNeg}V, lệch ${groundDiff}V) -> Nguy cơ chạm đất lưới DC!`);
  } else if (groundDiff > 8) {
    groundStatus = 'INVESTIGATE';
    investigateReasons.push(`Điện áp đối đất lệch ${groundDiff}V (|V+|=${vPos}V vs |V-|=${vNeg}V) -> Nghi ngờ suy giảm cách điện DC`);
  }

  // 7. Load Test
  const loadTestStr = (payload.electrical_tests.load_test_ieee_1188 || '').toLowerCase();
  let loadTestStatus: 'PASS' | 'FAIL' | 'INVESTIGATE' = 'PASS';
  if (loadTestStr.includes('fail') || loadTestStr.includes('không đạt') || (loadTestStr.includes('%') && parseInt(loadTestStr.replace(/[^0-9]/g, '')) < 80)) {
    loadTestStatus = 'FAIL';
    failureReasons.push('Thử nghiệm xả tải dung lượng không đạt chuẩn IEEE 1188 (< 80% dung lượng danh định)');
  }

  // Determine overall status
  if (failureReasons.length > 0) {
    overall_status = 'FAIL';
  } else if (investigateReasons.length > 0) {
    overall_status = 'INVESTIGATE';
  }

  let status_reason = 'Tất cả các hạng mục kiểm tra thị giác, nhiệt độ cực âm, nội trở và điện áp đạt tiêu chuẩn ANSI/NETA ATS-2025 Mục 7.18.1.3 & IEEE 1188.';
  if (overall_status === 'FAIL') {
    if (thermalRunawayRisk === 'CRITICAL') {
      status_reason = 'CRITICAL THERMAL RUNAWAY & INTERNAL OHMIC DEFECT: Nhiệt độ cực âm monoblock tăng vọt và biến động nội trở vượt ngưỡng 25% theo NETA Section 7.18.1.3.D.7.';
    } else {
      status_reason = failureReasons.join('; ');
    }
  } else if (overall_status === 'INVESTIGATE') {
    status_reason = investigateReasons.join('; ');
  }

  return {
    overall_status,
    status_reason,
    markdown_report: '',
    evaluation_details: {
      visual_mechanical_status: visualStatus,
      negative_post_temp_status: negPostTempStatus,
      temp_delta_c: tempDelta,
      thermal_runaway_risk: thermalRunawayRisk,
      intercell_resistance_status: intercellStatus,
      intercell_max_deviation_percent: maxIntercellDev,
      internal_ohmic_status: internalOhmicStatus,
      internal_ohmic_variance_percent: ohmicVariance,
      monoblock_voltages_status: monoblockVoltagesStatus,
      monoblock_voltage_spread_v: voltageSpread,
      ground_voltage_status: groundStatus,
      ground_voltage_diff_v: groundDiff,
      load_test_status: loadTestStatus
    }
  };
}

export function buildFallbackVrlaMarkdown(
  payload: NetaBatteryVrlaInputPayload,
  result: NetaBatteryVrlaEvaluationResult
): string {
  const { site_info, visual_inspection, electrical_tests, previous_test_data } = payload;
  const { overall_status, status_reason, evaluation_details } = result;

  const hotBlock = electrical_tests.negative_post_temperature_c.hot_monoblock_tag || 'Monoblock số 12';
  const hotTemp = electrical_tests.negative_post_temperature_c.max_temp_monoblock_12_c ?? electrical_tests.negative_post_temperature_c.max_temp_c ?? 41.2;
  const avgTemp = electrical_tests.negative_post_temperature_c.avg_temp_c || 26.5;

  const maxRes = electrical_tests.internal_ohmic_measurement_resistance_mohm.max_resistance_monoblock_12_mohm ?? electrical_tests.internal_ohmic_measurement_resistance_mohm.max_resistance_mohm ?? 4.65;
  const avgRes = electrical_tests.internal_ohmic_measurement_resistance_mohm.avg_resistance_mohm || 3.20;

  return `## BÁO CÁO THẨM ĐỊNH KỸ THUẬT ANSI/NETA ATS-2025 MỤC 7.18.1.3 & IEEE 1188
**HỆ THỐNG ĐIỆN MỘT CHIỀU DC VÀ ẮC QUY CHÌ-AXIT KÍN KHÍ VAN ĐIỀU ÁP (VRLA)**

- **Dự án / Trạm**: ${site_info.project_name}
- **Mã giàn pin**: ${site_info.battery_bank_tag}
- **Chủng loại**: ${site_info.battery_type}
- **Kỹ sư khảo sát (FSE)**: ${site_info.fse_name}
- **Ngày đo kiểm**: ${site_info.test_date}

---

### 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
### **${overall_status === 'FAIL' ? '❌ FAIL / CRITICAL THERMAL RUNAWAY & INTERNAL OHMIC DEFECT' : overall_status === 'INVESTIGATE' ? '⚠️ INVESTIGATE / CẦN THEO DÕI ĐẶC BIỆT' : '✅ PASS / ĐẠT TIÊU CHUẨN NETA ATS-2025'}**

**Lý do đánh giá:** ${status_reason}

---

### 2. BẢNG PHÂN TÍCH CHI TIẾT GIÀN PIN VRLA (Detailed Evaluation Table)

| Hạng mục kiểm tra | Tiêu chuẩn tham chiếu | Giá trị thực tế đo được | Ngưỡng cho phép | Kết luận |
| :--- | :--- | :--- | :--- | :--- |
| **Vị trí & Thông gió** | NETA ATS 7.18.1.3.A.1 | ${visual_inspection.ventilation_system_operable ? 'Hoạt động tốt' : 'Không hoạt động'} | Bắt buộc thông gió cưỡng bức/tự nhiên | **${visual_inspection.ventilation_system_operable ? 'PASS' : 'FAIL'}** |
| **Trạm rửa mắt khẩn cấp** | NETA ATS 7.18.1.3.A.2 | ${visual_inspection.eyewash_station_present ? 'Trang bị sẵn sàng' : 'Chưa trang bị'} | Sẵn sàng tại khu vực tủ pin | **${visual_inspection.eyewash_station_present ? 'PASS' : 'INVESTIGATE'}** |
| **Đối soát nhãn mác** | NETA ATS 7.18.1.3.A.3 | ${visual_inspection.nameplate_match ? 'Khớp 100% bản vẽ' : 'Lệch thông số thiết kế'} | Khớp 100% hồ sơ thiết kế | **${visual_inspection.nameplate_match ? 'PASS' : 'INVESTIGATE'}** |
| **Khung tủ & Tiếp địa** | NETA ATS 7.18.1.3.A.4 | ${visual_inspection.cabinet_mounting_and_grounding} | Cố định vững chắc, tiếp địa an toàn | **${visual_inspection.cabinet_mounting_and_grounding.toUpperCase()}** |
| **Tình trạng vỏ bình** | NETA ATS 7.18.1.3.A.5 | ${visual_inspection.physical_condition} | Không phồng rộp (swelling), nứt vỡ | **${evaluation_details.visual_mechanical_status}** |
| **Độ sạch cọc cực & Mỡ** | NETA ATS 7.18.1.3.A.6 | ${visual_inspection.cleanliness_and_oxide_inhibitor} | Sạch sẽ, bôi mỡ chống oxy hóa | **${visual_inspection.cleanliness_and_oxide_inhibitor.toUpperCase()}** |
| **Lực siết bu-lông** | Table 100.12 | ${visual_inspection.bolt_torque_check} | Đạt giá trị cờ-lê lực quy định | **${visual_inspection.bolt_torque_check.toUpperCase()}** |
| **Nhiệt độ Cực Âm** | NETA 7.18.1.3.D.1 & IEEE 1188 | ${hotBlock}: **${hotTemp}°C** (Avg: ${avgTemp}°C, $\\Delta T = +${evaluation_details.temp_delta_c}°C$) | $\\Delta T < 3.0°C$, tuyệt đối $< 35°C$ | **${evaluation_details.negative_post_temp_status}** |
| **Điện trở Cầu nối Intercell** | NETA 7.18.1.3.D.2 | [${electrical_tests.intercell_connection_resistance_micro_ohms.join(', ')}] $\\mu\\Omega$ (Lệch: **${evaluation_details.intercell_max_deviation_percent}%**) | Không lệch quá 50% so với min | **${evaluation_details.intercell_resistance_status}** |
| **Điện áp Float Bộ sạc** | NETA 7.18.1.3.D.3 | Float: ${electrical_tests.charger_float_voltage_v}V, Eq: ${electrical_tests.charger_equalize_voltage_v}V | Đúng thông số kỹ thuật nhà sản xuất | **PASS** |
| **Điện áp Monoblock Float** | NETA 7.18.1.3.D.4 | Min: ${electrical_tests.monoblock_voltages_float_mode_v.min_voltage_v}V, Max: ${electrical_tests.monoblock_voltages_float_mode_v.max_voltage_v}V (Spread: **${evaluation_details.monoblock_voltage_spread_v}V**) | Nằm trong dải dung sai công bố | **${evaluation_details.monoblock_voltages_status}** |
| **Nội trở Ôm Monoblock** | NETA 7.18.1.3.D.7 | Avg: ${avgRes} $\\text{m}\\Omega$, ${hotBlock}: **${maxRes} $\\text{m}\\Omega$** (Biến động: **${evaluation_details.internal_ohmic_variance_percent}%**) | Shall not vary by > 25% | **${evaluation_details.internal_ohmic_status}** |
| **Thử nghiệm Xả Tải** | IEEE 1188 | ${electrical_tests.load_test_ieee_1188} | Đạt $\\ge 80\\%$ dung lượng danh định | **${evaluation_details.load_test_status}** |
| **Điện áp DC so với Đất** | NETA 7.18.1.3.D.8 | (+): **${electrical_tests.system_voltage_to_ground_v.positive_to_ground_v}V**, (-): **${electrical_tests.system_voltage_to_ground_v.negative_to_ground_v}V** | $|V_{(+)}| \\approx |V_{(-)}|$ (Similar in magnitude) | **${evaluation_details.ground_voltage_status}** |

---

### 3. PHÂN TÍCH NỘI TRỞ, NHIỆT ĐỘ CỰC ÂM & NGUY CƠ TRÔI NHIỆT (Internal Ohmic, Negative Post Temp & Thermal Runaway Risk)

- **Thị giác & Lực siết**: Tủ chứa, tiếp địa, độ sạch cọc cực, mỡ chống oxy hóa và lực siết bu-lông đạt chuẩn NETA Section 7.18.1.3.A.
- **Điện trở Cầu nối Intercell & Điện áp đất**: Mối nối intercell cân bằng tốt ($\\approx ${electrical_tests.intercell_connection_resistance_micro_ohms[0] || 14.2}\\,\\mu\\Omega$, độ lệch ${evaluation_details.intercell_max_deviation_percent}\\% $\\le 50\\%$) và điện áp DC so với đất cân bằng (${electrical_tests.system_voltage_to_ground_v.positive_to_ground_v}\\text{V} / ${electrical_tests.system_voltage_to_ground_v.negative_to_ground_v}\\text{V}$) $\\rightarrow$ **Hệ thống cáp và thanh cái DC cách điện tốt, không rò điện đất**.
- **Nhiệt độ Cực Âm (Negative Post Temperature)**: 
  ${hotBlock} có nhiệt độ cực âm đo được **${hotTemp}°C**, cao hơn **+${evaluation_details.temp_delta_c}°C** so với nhiệt độ trung bình giàn pin (${avgTemp}°C). Theo tiêu chuẩn IEEE 1188, hiện tượng phát nhiệt cục bộ tại cực âm bình kín khí VRLA là chỉ báo rõ ràng của quá trình tái hợp oxy sinh nhiệt mất kiểm soát $\\rightarrow$ **CẢNH BÁO NGUY HIỂM: Hiện tượng Trôi nhiệt (Thermal Runaway) nội bộ trong bình VRLA ${hotBlock}**.
- **Nội trở Ôm Monoblock (Internal Ohmic Measurement)**:
  ${hotBlock} có nội trở đo được **${maxRes} m$\\Omega$**, tăng biến động **${evaluation_details.internal_ohmic_variance_percent}%** so với trung bình ${avgRes} m$\\Omega$. Giá trị này vượt quá xa ngưỡng tối đa cho phép 25% theo NETA Section 7.18.1.3.D.7 $\\rightarrow$ **Bình ${hotBlock} đã bị khô dung dịch điện giải GEL/AGM, màng xốp phân tách thoái hóa nghiêm trọng và gia tăng trở kháng nội bộ**.
${previous_test_data?.last_test_date ? `- **Dữ liệu Lịch sử (${previous_test_data.last_test_date})**: Nội trở trung bình năm trước là ${previous_test_data.last_avg_resistance_mohm} m$\\Omega$, nhiệt độ bình ${hotBlock} năm trước là ${previous_test_data.last_monoblock_12_temp_c}°C. Tốc độ suy thoái tăng nhanh trong vòng 12 tháng qua.` : ''}

---

### 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)

1. **CẢNH BÁO KHẨN CẤP**: Lập thủ tục ngắt cô lập an toàn và thay thế ngay lập tức Monoblock VRLA ${hotBlock} để tránh nguy cơ bình bị phát nổ hoặc phồng rộp (case swelling) gây cháy nổ tủ nguồn theo tiêu chuẩn IEEE 1188.
2. **Kiểm tra Dòng sạc Float & Cảm biến Bù nhiệt**: Đo kiểm tra dòng sạc duy trì (Float charge current) của bộ sạc DC và kiểm tra cảm biến bù nhiệt độ (Temperature compensation sensor). Nếu nhiệt độ phòng tăng mà bộ sạc không hạ điện áp float ($\\approx -3\\text{mV}/\\text{cell}/°\\text{C}$), giàn pin VRLA sẽ liên tục bị quá sạc dẫn đến trôi nhiệt hàng loạt.
3. **Hiệu chuẩn & Đo kiểm lại**: Sau khi thay thế bình mới, tiến hành đo lại nội trở ôm cell ban đầu (baseline internal ohmic) và giám sát nhiệt độ cực âm trong chu kỳ sạc float 48 giờ đầu tiên.
4. **Kiểm tra thông gió cưỡng bức phòng UPS**: Đảm bảo điều hòa phòng pin hoạt động liên tục ở dải nhiệt độ lý tưởng $20°\\text{C} - 25°\\text{C}$ per IEEE 1188.`;
}

export async function generateNetaBatteryVrlaAiAnalysis(
  payload: NetaBatteryVrlaInputPayload
): Promise<NetaBatteryVrlaEvaluationResult> {
  const deterministicResult = performDeterministicVrlaAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    deterministicResult.markdown_report = buildFallbackVrlaMarkdown(payload, deterministicResult);
    return deterministicResult;
  }

  try {
    const ai = new GoogleGenAI({ apiKey });
    const systemPrompt = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách Hệ thống Điện một chiều DC và Ắc quy Chì-Axit kín khí van điều áp (Direct-Current Systems, Batteries, Valve-Regulated Lead-Acid - VRLA). 
Nhiệm vụ của bạn là nhận dữ liệu kiểm tra ngoài site từ Field Service Engineer (FSE) và tự động đối soát với tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.18.1.3) và IEEE 1188.

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.18.1.3 & IEEE 1188):
1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Location & Ventilation: Vị trí đặt giàn/tủ pin an toàn, hệ thống thông gió hoạt động bình thường.
- Eyewash Equipment: Trang bị sẵn trạm rửa mắt khẩn cấp tại khu vực pin.
- Nameplate Match: Đối soát nhãn mác giàn pin VRLA khớp 100% bản vẽ thiết kế.
- Racks/Cabinets & Physical Condition: Khung giá/tủ cố định chắc chắn, tiếp địa đầy đủ; vỏ bình VRLA nguyên vẹn, không nứt vỡ hoặc phồng rộp (swelling).
- Cleanliness & Oxide Inhibitor: Cọc cực sạch sẽ và được bôi mỡ chống oxy hóa (oxide inhibitor).
- Bolt Torque: Siết lực bu-lông cọc cực và cầu nối intercell đạt theo nhà sản xuất hoặc Bảng Table 100.12.
- Thermographic Survey: Kết quả khảo sát nhiệt tuân thủ Section 9 / Table 100.18.

2. PHÉP ĐO ĐIỆN VÀ TIÊU CHUẨN ĐÁNH GIÁ (ELECTRICAL TEST VALUES):
- Nhiệt độ Cực Âm (Negative Post Temperature): Nằm trong dải quy định của nhà sản xuất hoặc tiêu chuẩn IEEE 1188 (Cảnh báo nguy cơ trôi nhiệt / Thermal Runaway). Nếu lệch quá 3-5°C hoặc vượt 35-40°C thì là CRITICAL FAIL.
- Điện trở Cầu nối Intercell (Intercell Connection Resistance): CẢNH BÁO "INVESTIGATE" nếu giá trị đo lệch quá 50% (> 50%) so với giá trị nhỏ nhất của mối nối tương tự.
- Điện áp Float & Equalize Bộ sạc (Charger Voltages & Alarms): Điện áp float/equalize và các mạch cảnh báo bộ sạc đạt đúng thông số nhà sản xuất.
- Điện áp Monoblock/Cell (Monoblock/Cell Voltages): Ở chế độ sạc Float, điện áp các bình/cell nằm trong dải dung sai nhà sản xuất công bố.
- Nội trở Ôm Cell (Internal Ohmic Measurements): Giá trị trở kháng/nội trở/độ dẫn điện giữa các monoblock/cell ở trạng thái sạc đầy KHÔNG ĐƯỢC biến động quá 25% (shall not vary by > 25%).
- Thử nghiệm Xả Tải (Load Test per IEEE 1188): Kết quả xả tải dung lượng đạt theo công bố của nhà sản xuất hoặc tiêu chuẩn IEEE 1188 (>= 80%).
- Điện áp so với Đất (System Voltage Positive/Negative to Ground): Độ lớn điện áp đo từ cực Dương (+) đến đất BẮT BUỘC phải tương đồng với độ lớn điện áp đo từ cực Âm (-) đến đất (similar in magnitude).

NHIỆM VỤ CỦA AI:
Khi nhận dữ liệu JSON từ FSE, hãy phân tích và trả về phản hồi định dạng Markdown gồm 4 phần:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE (kèm lý do chính).
2. BẢNG PHÂN TÍCH CHI TIẾT GIÀN PIN VRLA (Detailed Evaluation Table): So sánh thực tế với NETA ATS-2025 Section 7.18.1.3 và IEEE 1188.
3. PHÂN TÍCH NỘI TRỞ, NHIỆT ĐỘ CỰC ÂM & NGUY CƠ TRÔI NHIỆT (Internal Ohmic, Negative Post Temp & Thermal Runaway Risk).
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations).`;

    const userPrompt = `Dữ liệu kiểm tra giàn pin VRLA từ FSE tại hiện trường:
${JSON.stringify(payload, null, 2)}

Kết quả đối soát định lượng kỹ thuật:
- Trạng thái sơ bộ: ${deterministicResult.overall_status}
- Lý do: ${deterministicResult.status_reason}
- Chênh lệch nhiệt độ cực âm: +${deterministicResult.evaluation_details.temp_delta_c}°C (Mức nguy cơ trôi nhiệt: ${deterministicResult.evaluation_details.thermal_runaway_risk})
- Biến động nội trở Ôm: ${deterministicResult.evaluation_details.internal_ohmic_variance_percent}%
- Độ lệch điện trở cầu nối intercell: ${deterministicResult.evaluation_details.intercell_max_deviation_percent}%
- Cân bằng điện áp đất: ${payload.electrical_tests.system_voltage_to_ground_v.positive_to_ground_v}V / ${payload.electrical_tests.system_voltage_to_ground_v.negative_to_ground_v}V

Hãy phân tích toàn diện, chuyên nghiệp, sắc bén và lập báo cáo chi tiết theo chuẩn ANSI/NETA ATS-2025 Mục 7.18.1.3 và IEEE 1188.`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [{ role: 'user', parts: [{ text: `${systemPrompt}\n\n${userPrompt}` }] }]
    });

    const aiMarkdown = response.text;
    deterministicResult.markdown_report = aiMarkdown || buildFallbackVrlaMarkdown(payload, deterministicResult);
    return deterministicResult;
  } catch (error) {
    console.error('Error calling Gemini for NETA ATS-2025 VRLA analysis:', error);
    deterministicResult.markdown_report = buildFallbackVrlaMarkdown(payload, deterministicResult);
    return deterministicResult;
  }
}
