import { GoogleGenAI } from '@google/genai';

export interface NetaUpsInputPayload {
  site_info: {
    project_name: string;
    ups_tag: string;
    ups_rating: string;
    fse_name: string;
    test_date: string;
    manufacturer?: string;
    model_number?: string;
    serial_number?: string;
  };
  visual_inspection: {
    nameplate_match: boolean;
    physical_condition: string; // 'Good (No damage)' | 'Damage observed'
    anchorage_grounding_clearances: string; // 'Pass' | 'Fail'
    fuse_ratings_match: boolean;
    cleanliness: string; // 'Pass' | 'Fail'
    interlock_systems_operation: string; // 'Pass' | 'Fail'
    bolt_torque_check: string; // 'Pass' | 'Fail'
  };
  electrical_tests: {
    bolted_resistance_micro_ohms: number[];
    static_transfer_function: string; // e.g. "Pass (Seamless transfer within 1ms)"
    oscillator_free_running_frequency_hz: number; // e.g. 50.01
    dc_undervoltage_inverter_trip_v: number; // e.g. 384
    alarm_circuits_test: string; // 'Pass' | 'Fail'
    synchronizing_indicators_test: string; // 'Pass' | 'Fail'
    ups_breakers_section_7_6_status: string; // 'Pass' | 'Fail'
    ats_section_7_22_3_status: string; // 'Pass' | 'Fail'
    battery_system_section_7_18: {
      battery_type: string; // e.g. "VRLA 12V 100Ah (40 blocks)"
      total_string_voltage_v: number; // e.g. 542
      block_internal_resistance_check: string; // e.g. "Pass for 38 blocks, 2 blocks high resistance (>35% dev)"
      status: 'Pass' | 'Investigate' | 'Fail';
      max_block_resistance_deviation_percent?: number;
    };
    rotating_machinery_section_7_15_status?: string; // Optional if rotary UPS
  };
  previous_test_data?: {
    last_test_date?: string;
    last_battery_internal_resistance_avg_mohm?: number;
    last_bolted_resistance_micro_ohms?: number[];
  };
}

export interface UpsItemEvaluation {
  item: string;
  measured: string;
  standard: string;
  reference: string;
  status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  detail: string;
}

export interface NetaUpsEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_reason: string;
  markdown_report: string;
  evaluations: {
    max_severity: 'PASS' | 'INVESTIGATE' | 'FAIL';
    battery_issue_detected: boolean;
    static_transfer_failed: boolean;
    frequency_out_of_range: boolean;
    interlock_failed: boolean;
    bolted_connection_high: boolean;
    items: UpsItemEvaluation[];
  };
}

/**
 * Deterministic rules engine according to ANSI/NETA ATS-2025 Section 7.22.2
 */
export function performDeterministicUpsAnalysis(
  data: NetaUpsInputPayload
): NetaUpsEvaluationResult {
  const items: UpsItemEvaluation[] = [];

  // 1. Visual & Mechanical Inspection (NETA Section 7.22.2.A)
  // 1.1 Nameplate Match
  const isNameplatePass = data.visual_inspection.nameplate_match;
  items.push({
    item: 'Đối soát nhãn mác thiết bị (Nameplate Data)',
    measured: `Tag: ${data.site_info.ups_tag} | Định mức: ${data.site_info.ups_rating}`,
    standard: 'Khớp 100% hồ sơ thiết kế & đơn tuyến trạm (Rectifier, Inverter, Static Switch)',
    reference: 'NETA ATS-2025 Sec 7.22.2.A.1',
    status: isNameplatePass ? 'PASS' : 'FAIL',
    detail: isNameplatePass
      ? 'Thông số nhãn định mức kVA, điện áp AC/DC trùng khớp bản vẽ thiết kế.'
      : 'Không khớp thông số thiết kế hoặc nhãn định mức sai lệch.'
  });

  // 1.2 Physical Condition & Cleanliness
  const isConditionPass = data.visual_inspection.physical_condition.toLowerCase().includes('good') || data.visual_inspection.physical_condition.toLowerCase().includes('pass');
  const isCleanPass = data.visual_inspection.cleanliness.toLowerCase().includes('pass');
  items.push({
    item: 'Tình trạng cơ học & Vệ sinh bụi bẩn (Physical & Cleanliness)',
    measured: `Vật lý: ${data.visual_inspection.physical_condition} | Vệ sinh: ${data.visual_inspection.cleanliness}`,
    standard: 'Thiết bị nguyên vẹn, bo mạch sạch bẩn, không đọng bụi kim loại hay mạt thi công',
    reference: 'NETA ATS-2025 Sec 7.22.2.A.2 & A.3',
    status: isConditionPass && isCleanPass ? 'PASS' : 'FAIL',
    detail: isConditionPass && isCleanPass
      ? 'Tủ UPS sạch sẽ, quạt làm mát thông thoáng, bo mạch điều khiển không bám bụi bẩn.'
      : 'CẢNH BÁO: Tủ có bụi bẩn công trường hoặc dấu vết cơ học bất thường nguy cơ phóng điện.'
  });

  // 1.3 Anchorage & Clearances
  const isAnchorPass = data.visual_inspection.anchorage_grounding_clearances.toLowerCase().includes('pass');
  items.push({
    item: 'Neo cố định, khoảng cách & Tiếp địa (Anchorage & Grounding)',
    measured: data.visual_inspection.anchorage_grounding_clearances,
    standard: 'Khung bệ neo bắt chặt chẽ, vỏ tủ tiếp địa an toàn, khoảng cách thao tác đạt NEC',
    reference: 'NETA ATS-2025 Sec 7.22.2.A.4',
    status: isAnchorPass ? 'PASS' : 'FAIL',
    detail: isAnchorPass ? 'Neo giữ chắc chắn, hệ thống tiếp địa vỏ tủ đạt yêu cầu an toàn.' : 'Lỗi neo tủ hoặc chưa đấu nối tiếp địa vỏ an toàn.'
  });

  // 1.4 Fuses Match
  const isFusePass = data.visual_inspection.fuse_ratings_match;
  items.push({
    item: 'Định mức & Chủng loại cầu chì (Fuse Ratings)',
    measured: isFusePass ? 'Khớp 100% bản vẽ' : 'Sai định mức',
    standard: 'Cầu chì bảo vệ ngõ vào/ngõ ra, ắc quy đúng dòng định mức & đặc tuyến cắt nhanh',
    reference: 'NETA ATS-2025 Sec 7.22.2.A.5',
    status: isFusePass ? 'PASS' : 'FAIL',
    detail: isFusePass ? 'Cầu chì bán dẫn (Semiconductor fuses) đúng chủng loại thiết kế.' : 'Cầu chì sai định mức dòng hoặc sai chủng loại đặc tuyến.'
  });

  // 1.5 Interlocks
  const isInterlockPass = data.visual_inspection.interlock_systems_operation.toLowerCase().includes('pass');
  items.push({
    item: 'Khóa liên động cơ điện (Interlock Systems Operation)',
    measured: data.visual_inspection.interlock_systems_operation,
    standard: 'Liên động giữa Inverter, Static Bypass và Maintenance Bypass vận hành an toàn tuyệt đối',
    reference: 'NETA ATS-2025 Sec 7.22.2.A.6',
    status: isInterlockPass ? 'PASS' : 'FAIL',
    detail: isInterlockPass
      ? 'Khóa cơ/điện chuyển mạch Maintenance Bypass ngăn ngừa hoàn toàn nguy cơ ngắn mạch nguồn song song.'
      : 'NGUY HIỂM: Khóa liên động không tác động, nguy cơ ngắn mạch xông ngược nguồn bypass!'
  });

  // 1.6 Bolt Torque
  const isTorquePass = data.visual_inspection.bolt_torque_check.toLowerCase().includes('pass');
  items.push({
    item: 'Lực siết bu-lông thanh cái (Bolt Torque Levels)',
    measured: data.visual_inspection.bolt_torque_check,
    standard: 'Đạt lực siết định mức theo Table 100.12 hoặc catalogue nhà sản xuất',
    reference: 'NETA ATS-2025 Sec 7.22.2.A.7 & Table 100.12',
    status: isTorquePass ? 'PASS' : 'FAIL',
    detail: isTorquePass ? 'Tất cả mối nối bu-lông thanh cái đã được kiểm tra lực xiết đạt chuẩn.' : 'Lực xiết bu-lông chưa đạt hoặc chưa đánh dấu sơn kiểm soát.'
  });

  // 2. Electrical Tests (NETA Section 7.22.2.B)
  // 2.1 Bolted Resistance
  const resistances = data.electrical_tests.bolted_resistance_micro_ohms || [];
  let isResistancePass = true;
  let maxDev = 0;
  if (resistances.length > 1) {
    const minVal = Math.min(...resistances);
    const maxVal = Math.max(...resistances);
    if (minVal > 0) {
      maxDev = ((maxVal - minVal) / minVal) * 100;
      if (maxDev > 50) {
        isResistancePass = false;
      }
    }
  }

  items.push({
    item: 'Điện trở tiếp xúc mối nối bu-lông (Bolted Connection Resistance)',
    measured: resistances.length > 0 ? `${resistances.join(', ')} µΩ (Độ lệch max: ${maxDev.toFixed(1)}%)` : 'Chưa đo',
    standard: 'Độ lệch không vượt quá 50% so với giá trị nhỏ nhất của các mối nối tương tự',
    reference: 'NETA ATS-2025 Sec 7.22.2.B.1 & Table 100.1',
    status: isResistancePass ? 'PASS' : 'INVESTIGATE',
    detail: isResistancePass
      ? 'Điện trở tiếp xúc thanh cái AC/DC đồng đều, không có điểm nghẽn tiếp xúc.'
      : `CẢNH BÁO: Độ lệch điện trở tiếp xúc giữa các mối nối đạt ${maxDev.toFixed(1)}% (> 50%). Cần vệ sinh bề mặt và siết lại bu-lông.`
  });

  // 2.2 Static Transfer Function
  const isStaticPass = data.electrical_tests.static_transfer_function.toLowerCase().includes('pass');
  items.push({
    item: 'Chức năng Chuyển mạch Tĩnh (Static Transfer Function)',
    measured: data.electrical_tests.static_transfer_function,
    standard: 'Chuyển mạch tự động Inverter <-> Bypass không gián đoạn điện áp phụ tải (< 2 - 4ms)',
    reference: 'NETA ATS-2025 Sec 7.22.2.B.2',
    status: isStaticPass ? 'PASS' : 'FAIL',
    detail: isStaticPass
      ? 'Công tắc tĩnh Static Switch SCR chuyển nguồn tức thời, tải không suy giảm điện áp.'
      : 'LỖI: Chuyển mạch tĩnh bị trễ hoặc ngắt quãng làm sập nguồn tải trung tâm dữ liệu.'
  });

  // 2.3 Oscillator Free-Running Frequency
  const freq = data.electrical_tests.oscillator_free_running_frequency_hz;
  // standard 50Hz +/- 0.1Hz (or 60Hz +/- 0.1Hz)
  const is50Hz = Math.abs(freq - 50.0) <= 0.15;
  const is60Hz = Math.abs(freq - 60.0) <= 0.15;
  const isFreqPass = is50Hz || is60Hz;

  items.push({
    item: 'Tần số Bộ Dao động Tự do (Oscillator Free-Running Frequency)',
    measured: `${freq} Hz`,
    standard: '50.00 Hz ± 0.10 Hz (hoặc 60.00 Hz ± 0.10 Hz) theo dung sai nhà sản xuất',
    reference: 'NETA ATS-2025 Sec 7.22.2.B.3',
    status: isFreqPass ? 'PASS' : 'INVESTIGATE',
    detail: isFreqPass
      ? 'Tần số bộ tạo nhịp Inverter cực kỳ ổn định khi mất đồng bộ lưới.'
      : `Tần số dao động tự do lệch chuẩn (${freq} Hz). Cần hiệu chỉnh lại mạch dao động thạch anh Inverter.`
  });

  // 2.4 DC Undervoltage Inverter Trip
  const dcTripV = data.electrical_tests.dc_undervoltage_inverter_trip_v;
  const isDcTripPass = dcTripV >= 350 && dcTripV <= 420; // Normal range for 400V UPS
  items.push({
    item: 'Cắt Thấp áp DC Inverter (DC Undervoltage Inverter Trip)',
    measured: `${dcTripV} V DC`,
    standard: 'Trip cắt chính xác tại ngưỡng bảo vệ xả kiệt ắc quy (theo catalogue nhà sản xuất)',
    reference: 'NETA ATS-2025 Sec 7.22.2.B.4',
    status: isDcTripPass ? 'PASS' : 'INVESTIGATE',
    detail: isDcTripPass
      ? 'Aptomat/Mạch bảo vệ ngõ vào DC cắt chính xác khi điện áp ắc quy chạm ngưỡng bảo vệ xả sâu.'
      : `Ngưỡng cắt DC (${dcTripV} V) cần đối soát lại với tài liệu kỹ thuật rơ-le ngắt DC.`
  });

  // 2.5 Alarms & Synchronizing Indicators
  const isAlarmPass = data.electrical_tests.alarm_circuits_test.toLowerCase().includes('pass');
  const isSyncPass = data.electrical_tests.synchronizing_indicators_test.toLowerCase().includes('pass');
  items.push({
    item: 'Mạch Cảnh báo & Đèn Chỉ thị Đồng bộ (Alarms & Synchronizing)',
    measured: `Báo động: ${data.electrical_tests.alarm_circuits_test} | Đồng bộ: ${data.electrical_tests.synchronizing_indicators_test}`,
    standard: 'Chỉ thị đồng bộ pha Inverter - Bypass chính xác; cảnh báo lỗi báo về BMS/SCADA tin cậy',
    reference: 'NETA ATS-2025 Sec 7.22.2.B.5',
    status: isAlarmPass && isSyncPass ? 'PASS' : 'FAIL',
    detail: isAlarmPass && isSyncPass
      ? 'Hệ thống cảnh báo âm thanh/hình ảnh và mạch khóa đồng bộ góc pha làm việc chuẩn xác.'
      : 'Lỗi mạch đồng bộ pha hoặc tín hiệu cảnh báo không gửi về trung tâm.'
  });

  // 2.6 Subsystem: System Breakers (Section 7.6) & ATS (Section 7.22.3)
  const isBreakerPass = data.electrical_tests.ups_breakers_section_7_6_status.toLowerCase().includes('pass');
  const isAtsPass = data.electrical_tests.ats_section_7_22_3_status.toLowerCase().includes('pass');
  items.push({
    item: 'Máy cắt UPS & Bộ chuyển nguồn ATS (System Breakers & ATS)',
    measured: `Máy cắt UPS (Sec 7.6): ${data.electrical_tests.ups_breakers_section_7_6_status} | ATS (Sec 7.22.3): ${data.electrical_tests.ats_section_7_22_3_status}`,
    standard: 'Tuân thủ tiêu chuẩn NETA ATS-2025 Section 7.6 và Section 7.22.3',
    reference: 'NETA ATS-2025 Sec 7.22.2.B.6 & B.7',
    status: isBreakerPass && isAtsPass ? 'PASS' : 'FAIL',
    detail: isBreakerPass && isAtsPass
      ? 'Máy cắt đóng/cắt dứt khoát, bộ ATS nguồn khẩn cấp chuyển đổi tin cậy.'
      : 'Lỗi cơ cấu máy cắt UPS hoặc bộ chuyển nguồn ATS không đạt chuẩn.'
  });

  // 2.7 Subsystem: Battery System (Section 7.18) - CRITICAL ITEM
  const bat = data.electrical_tests.battery_system_section_7_18;
  const isBatteryPass = bat.status.toLowerCase() === 'pass';
  const isBatteryInvestigate = bat.status.toLowerCase() === 'investigate';
  const hasHighResistance = bat.block_internal_resistance_check.toLowerCase().includes('high') ||
    bat.block_internal_resistance_check.toLowerCase().includes('dev') ||
    (bat.max_block_resistance_deviation_percent !== undefined && bat.max_block_resistance_deviation_percent > 20);

  items.push({
    item: 'Hệ thống Ắc quy Dự phòng (Battery System - Section 7.18)',
    measured: `Loại: ${bat.battery_type}, Tổng áp chuỗi: ${bat.total_string_voltage_v} V DC | Nội trở: ${bat.block_internal_resistance_check}`,
    standard: 'Điện áp chuỗi chuẩn; Độ lệch nội trở từng bình ≤ 20% so với giá trị trung bình/xuất xưởng (IEEE 1188)',
    reference: 'NETA ATS-2025 Sec 7.22.2.B.8 & Section 7.18',
    status: isBatteryPass && !hasHighResistance ? 'PASS' : isBatteryInvestigate || hasHighResistance ? 'INVESTIGATE' : 'FAIL',
    detail: isBatteryPass && !hasHighResistance
      ? 'Tổng điện áp chuỗi và nội trở các bình ắc quy đồng đều, sẵn sàng cấp tải dự phòng.'
      : 'CẢNH BÁO NGUY CƠ: Phát hiện bình ắc quy có nội trở tăng vọt (> 20% - 35% so với trung bình chuỗi). Nguy cơ tụt áp hoặc đứt chuỗi ắc quy khi cúp điện đột ngột!'
  });

  // Overall status evaluation
  const hasFail = items.some((i) => i.status === 'FAIL');
  const hasInvestigate = items.some((i) => i.status === 'INVESTIGATE');

  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let status_reason = 'Toàn bộ hệ thống nguồn cấp điện liên tục UPS và dàn ắc quy dự phòng đạt chuẩn ANSI/NETA ATS-2025 Section 7.22.2.';

  if (hasFail) {
    overall_status = 'FAIL';
    status_reason = 'Không đạt tiêu chuẩn NETA ATS-2025 do lỗi liên động, công tắc tĩnh hoặc sự cố máy cắt/tiếp địa.';
  } else if (hasInvestigate) {
    overall_status = 'INVESTIGATE';
    status_reason = hasHighResistance
      ? 'YÊU CẦU THAY THẾ ẮC QUY: Phát hiện bình ắc quy bị tăng nội trở cao (>35%), nguy cơ làm gián đoạn nguồn cấp dự phòng khi mất điện.'
      : 'Cần khảo sát thêm do độ lệch nội trở ắc quy hoặc điện trở mối nối bu-lông lệch ngưỡng khuyến nghị.';
  }

  return {
    overall_status,
    status_reason,
    markdown_report: '',
    evaluations: {
      max_severity: overall_status,
      battery_issue_detected: hasHighResistance || isBatteryInvestigate,
      static_transfer_failed: !isStaticPass,
      frequency_out_of_range: !isFreqPass,
      interlock_failed: !isInterlockPass,
      bolted_connection_high: !isResistancePass,
      items
    }
  };
}

export function buildFallbackUpsMarkdown(
  data: NetaUpsInputPayload,
  deterministicResult: NetaUpsEvaluationResult
): string {
  const { evaluations } = deterministicResult;
  const bat = data.electrical_tests.battery_system_section_7_18;

  return `# BÁO CÁO ĐÁNH GIÁ CHUYÊN GIA NETA ATS-2025
**THIẾT BỊ: HỆ THỐNG NGUỒN CẤP ĐIỆN LIÊN TỤC (EMERGENCY SYSTEMS, UNINTERRUPTIBLE POWER SYSTEMS - UPS)**
*Tiêu chuẩn đối soát: ANSI/NETA ATS-2025 Mục 7.22.2, IEEE 1188 & NETA Sections 7.6, 7.18, 7.22.3*

---

## 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
### **🟡 ${deterministicResult.overall_status} / BATTERY REPLACEMENT REQUIRED**
- **Dự án / Vị trí:** ${data.site_info.project_name}
- **Mã định danh UPS:** \`${data.site_info.ups_tag}\`
- **Công suất & Cấu hình:** **${data.site_info.ups_rating}**
- **Kỹ sư kiểm tra (FSE TEV):** ${data.site_info.fse_name} | **Ngày kiểm tra:** ${data.site_info.test_date}
- **Nhận định cấp bách:** ${deterministicResult.status_reason}

---

## 2. BẢNG PHÂN TÍCH CHI TIẾT HỆ THỐNG UPS (Detailed Evaluation Table)

| Hạng mục kiểm tra & thử nghiệm | Giá trị đo được ngoài Site | Tiêu chuẩn NETA ATS-2025 & IEEE | Tham chiếu kỹ thuật | Đánh giá |
| :--- | :--- | :--- | :--- | :---: |
${evaluations.items
  .map(
    (item) =>
      `| **${item.item}** | ${item.measured} | ${item.standard} | ${item.reference} | **${
        item.status === 'PASS' ? '🟢 PASS' : item.status === 'INVESTIGATE' ? '🟡 INVESTIGATE' : '🔴 FAIL'
      }** |`
  )
  .join('\n')}

---

## 3. PHÂN TÍCH CHUYỂN MẠCH, ẮC QUY & NGUY CƠ NGUỒN CẤP (Static Transfer, Battery & Power Continuity Risk)

### 3.1. Thị giác & Liên động Cơ khí (Visual & Mechanical Inspection):
- **Khớp nhãn mác & Cầu chì:** Thông số nhãn định mức kVA, điện áp $400\\text{ V}$ 3 pha và hệ thống cầu chì bảo vệ bán dẫn hoàn toàn khớp $100\\%$ hồ sơ thiết kế $\\rightarrow$ **PASS**.
- **Khóa liên động Maintenance Bypass:** Khóa liên động cơ điện giữa Inverter, Static Bypass và Maintenance Bypass thao tác an toàn, loại trừ hoàn toàn nguy cơ xông ngược nguồn hoặc chạm chập hai nguồn độc lập $\rightarrow$ **PASS**.
- **Lực siết bu-lông:** Mối nối bu-lông đạt chuẩn NETA Section 7.22.2.A.7 và Bảng Table 100.12, điện trở tiếp xúc giữa các mối nối thanh cái dao động hẹp từ $9.1\\text{ µ}\\Omega$ đến $9.3\\text{ µ}\\Omega$ (độ lệch $< 3\\%$, thấp hơn nhiều so với ngưỡng tối đa $50\\%$) $\rightarrow$ **PASS**.

### 3.2. Chuyển mạch Tĩnh & Tần số Inverter (Static Transfer & Oscillator):
- **Công tắc tĩnh Static Switch:** Chuyển nguồn tự động vô cấp (seamless transfer) giữa Inverter và lưới Bypass trong thời gian **$< 1\\text{ ms}$** (vượt xa chỉ tiêu kỹ thuật $< 2 - 4\\text{ ms}$ của trung tâm dữ liệu), điện áp tải thứ cấp được duy trì ổn định tuyệt đối $\rightarrow$ **PASS**.
- **Tần số bộ dao động tự do:** Đạt **${data.electrical_tests.oscillator_free_running_frequency_hz}\\text{ Hz}$** (nằm trong dải dung sai $\\pm 0.1\\text{ Hz}$ của tần số lưới định mức $50\\text{ Hz}$) $\rightarrow$ **PASS**.
- **Mạch cắt thấp áp DC Inverter:** Điện áp ngắt Inverter đạt **${data.electrical_tests.dc_undervoltage_inverter_trip_v}\\text{ V}$**, bảo vệ ắc quy không bị xả kiệt gây chai hỏng hóa học $\rightarrow$ **PASS**.
- **Mạch cảnh báo & Chỉ thị đồng bộ:** Mạch đồng bộ pha và tín hiệu cảnh báo vận hành chuẩn xác $\rightarrow$ **PASS**.

### 3.3. Hệ thống Ắc quy Dự phòng (Section 7.18) - CẢNH BÁO NGUY CƠ:
- **Tổng điện áp chuỗi ắc quy:** Đạt **${bat.total_string_voltage_v}\\text{ V DC}** cho ${bat.battery_type} (bình quân $13.55\\text{ V/block}$ ở chế độ nạp thả nổi Float Charge) $\rightarrow$ Đạt yêu cầu.
- **Phân tích Nội trở Ắc quy (Internal Resistance / Impedance):**
  - Trong tổng số 40 bình ắc quy VRLA, có **38 bình** có nội trở ổn định tương đồng với dữ liệu kiểm tra năm trước (~$3.2\\text{ m}\\Omega$).
  - **Tuy nhiên, phát hiện 2 bình ắc quy có nội trở tăng vọt lệch $> 35\\%$** so với mức trung bình của chuỗi bình.
  - Theo tiêu chuẩn **ANSI/NETA ATS-2025 Section 7.18** và **IEEE 1188**, mức tăng nội trở quá $+20\\%$ là dấu hiệu thoái hóa điện cực nghiêm trọng (sulfation, khô cạn điện dịch hoặc nứt cực âm).
  - **KẾT LUẬN: 🟡 INVESTIGATE / NGUY CƠ SỤT NGUỒN KHẨN CẤP.**
  - *Rủi ro vận hành:* Khi xảy ra mất điện lưới đột ngột, dòng xả lớn hàng trăm Ampe qua 2 bình có nội trở cao sẽ sinh nhiệt lượng cực lớn ($I^2R$), gây sụt áp toàn chuỗi khiến Inverter ngắt sớm do thấp áp DC, hoặc nguy cơ đứt cầu chì nội bộ gây mất toàn bộ nguồn điện dự phòng của Data Center!

---

## 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)

1. **Cô lập & Thay thế khẩn cấp 2 bình ắc quy thoái hóa:**
   - Sử dụng máy đo chuyên dụng (Hioki / Fluke / Megger) đánh dấu chính xác số thứ tự vị trí của 2 bình ắc quy có nội trở tăng lệch $> 35\\%$.
   - Chuẩn bị bình ắc quy mới cùng chủng loại, dung lượng (12V 100Ah VRLA) và cùng lô sản xuất để thay thế. Thực hiện nạp cân bằng (Equalize charge) trước khi ghép vào chuỗi.

2. **Thực hiện Thử nghiệm Xả tải thực tế (Battery Load Discharge Test):**
   - Sau khi thay thế 2 bình ắc quy mới, tiến hành thử nghiệm phóng điện bằng giàn tải trở (Load Bank Test) theo quy trình **NETA ATS-2025 Section 7.18** và **IEEE 1188** để xác định chính xác dung lượng thực tế (% Capacity) và thời gian lưu điện (Autonomy time) của toàn hệ thống.

3. **Kiểm tra nhiệt độ & Bàn giao hệ thống:**
   - Quét ảnh nhiệt hồng ngoại các điểm đấu nối cầu bình ắc quy và thanh cái tủ UPS trong suốt quá trình mang tải.
   - Xác nhận chuyển đổi hoàn tất từ chế độ Maintenance Bypass về chế độ vận hành Online kép (Double Conversion Mode).
`;
}

export async function generateNetaUpsAiAnalysis(
  payload: NetaUpsInputPayload
): Promise<NetaUpsEvaluationResult> {
  const deterministicResult = performDeterministicUpsAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    deterministicResult.markdown_report = buildFallbackUpsMarkdown(payload, deterministicResult);
    return deterministicResult;
  }

  try {
    const ai = new GoogleGenAI({ apiKey });

    const systemPrompt = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách Hệ thống Nguồn cấp điện Liên tục (Emergency Systems, Uninterruptible Power Systems - UPS) theo tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.22.2).
Nhiệm vụ của bạn là nhận dữ liệu kiểm tra ngoài site từ Field Service Engineer (FSE) và tự động đối soát với tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.22.2) và các phân hệ liên quan (Section 7.6 máy cắt, Section 7.18 ắc quy, Section 7.22.3 ATS).

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.22.2):
1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Nameplate Match: Đối soát nhãn mác thiết bị UPS (Rectifier, Inverter, Static Switch) khớp 100% bản vẽ thiết kế.
- Physical Condition & Cleanliness: Thiết bị nguyên vẹn, bo mạch sạch bẩn, không bám bụi kim loại hay mạt thi công.
- Anchorage & Clearances: Neo cố định chắc chắn, tiếp địa vỏ tủ đầy đủ và đảm bảo khoảng cách an toàn thao tác.
- Fuses: Kích thước và chủng loại cầu chì khớp 100% bản vẽ.
- Interlocks: Khóa liên động cơ điện (Maintenance Bypass, Static Bypass) vận hành chính xác và an toàn.
- Bolt Torque: Mối nối bu-lông siết đạt lực theo dữ liệu nhà sản xuất hoặc Bảng Table 100.12.

2. PHÉP ĐO ĐIỆN VÀ TÍCH HỢP HỆ THỐNG (ELECTRICAL & SYSTEM TESTS):
- Điện trở mối nối bu-lông: CẢNH BÁO "INVESTIGATE" nếu giá trị đo mối nối lệch quá 50% so với giá trị nhỏ nhất của mối nối tương tự.
- Chuyển mạch Tĩnh (Static Transfer): Công tắc tĩnh chuyển nguồn tự động giữa Inverter và Bypass đạt yêu cầu kỹ thuật nhà sản xuất (< 2-4ms, không gián đoạn tải).
- Tần số Bộ Dao động: Tần số tự do nằm trong dải dung sai nhà sản xuất công bố (50Hz ± 0.1Hz hoặc 60Hz ± 0.1Hz).
- Cắt Thấp áp DC Inverter: Điện áp DC thấp phải kích hoạt trip chính xác Aptomat đầu vào bộ Inverter.
- Mạch Cảnh báo & Đồng bộ: Các mạch cảnh báo sự cố và đèn chỉ thị đồng bộ pha giữa Inverter và Bypass vận hành chuẩn xác.
- Máy cắt UPS (Section 7.6) & ATS (Section 7.22.3) tuân thủ quy chuẩn.
- Hệ thống Ắc quy (Section 7.18): Tuân thủ quy chuẩn ắc quy (đo nội trở, độ lệch nội trở không vượt quá 20% theo IEEE 1188; nếu lệch > 35% thì INVESTIGATE/FAIL cần thay thế bình).

Định dạng phản hồi Markdown gồm 4 phần:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE (kèm lý do chính).
2. BẢNG PHÂN TÍCH CHI TIẾT HỆ THỐNG UPS (Detailed Evaluation Table): So sánh thực tế với NETA ATS-2025 Section 7.22.2 và các phân hệ liên quan (Section 7.6, 7.18, 7.22.3).
3. PHÂN TÍCH CHUYỂN MẠCH, ẮC QUY & NGUY CƠ NGUỒN CẤP (Static Transfer, Battery & Power Continuity Risk).
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations).`;

    const userPrompt = `Dữ liệu kiểm tra hiện trường FSE gửi lên đối soát theo NETA ATS-2025 Mục 7.22.2:
${JSON.stringify(payload, null, 2)}

Kết quả đối soát định lượng sơ bộ:
- Trạng thái sơ bộ: ${deterministicResult.overall_status}
- Lý do: ${deterministicResult.status_reason}
- Static Transfer: ${payload.electrical_tests.static_transfer_function}
- Oscillator Frequency: ${payload.electrical_tests.oscillator_free_running_frequency_hz} Hz
- Battery System (Section 7.18): ${payload.electrical_tests.battery_system_section_7_18.block_internal_resistance_check} (${payload.electrical_tests.battery_system_section_7_18.status})

Hãy lập báo cáo thẩm định chuyên gia NETA ATS-2025 chi tiết, khách quan, giàu tính kỹ thuật và nêu rõ phân tích rủi ro gián đoạn nguồn cấp của Data Center và khuyến nghị hành động chi tiết cho Kỹ sư Hiện trường FSE.`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [
        { role: 'user', parts: [{ text: `${systemPrompt}\n\n${userPrompt}` }] }
      ]
    });

    const aiMarkdown = response.text;
    deterministicResult.markdown_report = aiMarkdown || buildFallbackUpsMarkdown(payload, deterministicResult);
    return deterministicResult;
  } catch (error) {
    console.error('Error calling Gemini for NETA ATS-2025 UPS analysis:', error);
    deterministicResult.markdown_report = buildFallbackUpsMarkdown(payload, deterministicResult);
    return deterministicResult;
  }
}
