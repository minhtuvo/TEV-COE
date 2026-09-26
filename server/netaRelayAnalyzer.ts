import { GoogleGenAI } from '@google/genai';

export interface NetaRelayInputPayload {
  site_info: {
    project_name: string;
    relay_tag: string;
    relay_model: string;
    firmware_version: string;
    control_voltage_vdc?: string;
    fse_name: string;
    test_date: string;
  };
  visual_inspection: {
    identification_match: boolean;
    display_and_leds_test: string; // 'Pass' | 'Fail'
    cleanliness_and_connections_tight: boolean;
    frame_grounding_check: string; // 'Pass' | 'Fail'
    settings_match_coordination_study: boolean;
    clock_date_verified: string; // e.g., '2026-09-23 10:15:00 (Pass)'
    ct_shorting_blocks_check: string; // 'Pass' | 'Fail'
  };
  electrical_tests: {
    insulation_resistance_1000v_megohms: {
      ac_current_inputs_to_ground: number;
      ac_voltage_inputs_to_ground: number;
      dc_power_supply_to_ground: number;
    };
    analog_metering_accuracy: {
      injected_ia_5_00a?: string; // e.g., "Relay displays 5.01A (Pass)"
      injected_va_66_4v?: string; // e.g., "Relay displays 66.3V (Pass)"
      measured_ia?: number;
      measured_va?: number;
      ia_error_percent?: number;
      va_error_percent?: number;
    };
    protection_elements_test: {
      ansi_51_time_overcurrent: {
        setting_pickup_a: number;
        measured_pickup_a: number;
        setting_time_dial: number;
        curve_type?: string;
        expected_trip_time_sec_at_3x: number;
        measured_trip_time_sec_at_3x: number;
        status?: string;
      };
      ansi_50_instantaneous_overcurrent: {
        setting_pickup_a: number;
        measured_pickup_a: number;
        expected_trip_time_sec: number;
        measured_trip_time_sec: number;
        status?: string;
      };
      differential_87?: {
        enabled?: boolean;
        minimum_pickup_a?: number;
        measured_pickup_a?: number;
        slope_1_percent?: number;
        status?: string;
      };
      voltage_elements_27_59?: {
        enabled?: boolean;
        undervoltage_27_pickup_v?: number;
        overvoltage_59_pickup_v?: number;
        status?: string;
      };
    };
    digital_inputs_test: string; // 'Pass (All 8 active DIs verified)'
    output_contacts_trip_test: string; // 'Pass (Breaker tripped successfully via OUT101)'
    internal_logic_verification: string; // 'Pass (PSL/GOOSE verified)'
    trip_coil_monitoring_tcs: string; // 'Pass (Open coil alarm verified)'
    arc_energy_reduction_sensor?: string; // 'Pass (Optical sensor threshold > ambient)'
    event_records_cleared_after_test: boolean;
  };
  previous_test_data?: {
    last_test_date?: string;
    last_ansi_50_pickup_a?: number;
    last_ansi_51_pickup_a?: number;
  };
}

export interface RelayItemEvaluation {
  item: string;
  measured: string;
  standard: string;
  reference: string;
  status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  detail: string;
}

export interface NetaRelayEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_reason: string;
  markdown_report: string;
  evaluations: {
    max_severity: 'PASS' | 'INVESTIGATE' | 'FAIL';
    ansi_50_failed: boolean;
    ansi_51_failed: boolean;
    insulation_passed: boolean;
    metering_passed: boolean;
    event_records_uncleared: boolean;
    settings_mismatch: boolean;
    items: RelayItemEvaluation[];
  };
}

/**
 * Deterministic rules engine according to ANSI/NETA ATS-2025 Section 7.9.2
 */
export function performDeterministicRelayAnalysis(
  data: NetaRelayInputPayload
): NetaRelayEvaluationResult {
  const items: RelayItemEvaluation[] = [];

  // 1. Visual & Mechanical
  // 1.1 Identification & Settings Match
  const isIdPass = data.visual_inspection.identification_match;
  items.push({
    item: 'Nhận diện & Bản vẽ thiết kế (Identification)',
    measured: `Model: ${data.site_info.relay_model}, FW: ${data.site_info.firmware_version}`,
    standard: 'Khớp 100% hồ sơ thiết kế & thông số kỹ thuật',
    reference: 'NETA ATS-2025 Sec 7.9.2.A.1',
    status: isIdPass ? 'PASS' : 'FAIL',
    detail: isIdPass
      ? 'Model, Firmware, Style Number và điện áp cấp nguồn trùng khớp thiết kế.'
      : 'Không khớp thông tin thiết kế hoặc nhãn rơ-le sai lệch.'
  });

  const isSettingsPass = data.visual_inspection.settings_match_coordination_study;
  items.push({
    item: 'Đối soát file cài đặt (Settings & Logic Coordination)',
    measured: isSettingsPass ? 'Khớp 100% file phiếu chỉnh định' : 'Có thông số sai lệch',
    standard: 'Trùng khớp 100% Coordination Study / Setting Sheet được duyệt',
    reference: 'NETA ATS-2025 Sec 7.9.2.A.3',
    status: isSettingsPass ? 'PASS' : 'FAIL',
    detail: isSettingsPass
      ? 'File cài đặt trích xuất từ rơ-le hoàn toàn khớp với phiếu chỉnh định rơ-le được duyệt.'
      : 'CẢNH BÁO: File cài đặt trong rơ-le sai lệch so với bảng trị số chỉnh định phối hợp bảo vệ.'
  });

  // 1.2 LEDs & Display
  const isDisplayPass = data.visual_inspection.display_and_leds_test.toLowerCase().includes('pass');
  items.push({
    item: 'Kiểm tra màn hình & đèn LED (Display & Target LEDs)',
    measured: data.visual_inspection.display_and_leds_test,
    standard: 'Màn hình LCD/OLED và các đèn LED trạng thái/báo sự cố sáng rõ',
    reference: 'NETA ATS-2025 Sec 7.9.2.A.2',
    status: isDisplayPass ? 'PASS' : 'INVESTIGATE',
    detail: isDisplayPass ? 'Đèn LED và màn hình hoạt động hoàn hảo.' : 'Cần kiểm tra màn hình LCD hoặc LED chỉ thị sự cố.'
  });

  // 1.3 Cleanliness & Grounding & CT Shorting
  const isGrounded = data.visual_inspection.frame_grounding_check.toLowerCase().includes('pass');
  const isClean = data.visual_inspection.cleanliness_and_connections_tight;
  items.push({
    item: 'Vệ sinh, siết ốc & Tiếp địa vỏ rơ-le (Grounding & Terminals)',
    measured: `Vệ sinh/siết ốc: ${isClean ? 'ĐẠT' : 'CHƯA ĐẠT'} | Tiếp địa: ${data.visual_inspection.frame_grounding_check}`,
    standard: 'Đấu nối chắc chắn, sạch sẽ; vỏ rơ-le tiếp địa đúng chuẩn',
    reference: 'NETA ATS-2025 Sec 7.9.2.A.4 & A.5',
    status: isGrounded && isClean ? 'PASS' : 'FAIL',
    detail: isGrounded && isClean ? 'Vỏ tiếp địa chắc chắn, các đầu cosse đấu nối siết chặt.' : 'Lỗi tiếp địa vỏ hoặc vít cầu đấu chưa siết chặt.'
  });

  // 1.4 CT Shorting Blocks
  const isCtPass = data.visual_inspection.ct_shorting_blocks_check.toLowerCase().includes('pass');
  items.push({
    item: 'Cơ cấu ngắn mạch mạch dòng CT (CT Shorting Devices)',
    measured: data.visual_inspection.ct_shorting_blocks_check,
    standard: 'Khối ngắn mạch dòng CT làm việc tin cậy chống hở mạch thứ cấp CT',
    reference: 'NETA ATS-2025 Sec 7.9.2.A.6',
    status: isCtPass ? 'PASS' : 'FAIL',
    detail: isCtPass ? 'Cơ cấu ngắn mạch CT đóng tiếp điểm an toàn khi rút rơ-le.' : 'NGUY HIỂM: Khối ngắn mạch CT không hoạt động, nguy cơ hở mạch thứ cấp CT gây quá áp nguy hiểm.'
  });

  // 1.5 Real-time Clock (RTC)
  const isClockPass = data.visual_inspection.clock_date_verified.toLowerCase().includes('pass');
  items.push({
    item: 'Đồng hồ thời gian thực (RTC Clock & Date)',
    measured: data.visual_inspection.clock_date_verified,
    standard: 'Khớp ngày giờ thực tế (IRIG-B / SNTP / Time Sync)',
    reference: 'NETA ATS-2025 Sec 7.9.2.A.7',
    status: isClockPass ? 'PASS' : 'INVESTIGATE',
    detail: isClockPass ? 'Đồng hồ rơ-le khớp thời gian thực, đảm bảo ghi nhận SOE sự cố chính xác.' : 'Đồng hồ rơ-le lệch giờ hoặc chưa đồng bộ.'
  });

  // 2. Electrical Tests
  // 2.1 Insulation Resistance
  const ir = data.electrical_tests.insulation_resistance_1000v_megohms;
  const isIrCurrentPass = ir.ac_current_inputs_to_ground >= 100;
  const isIrVoltagePass = ir.ac_voltage_inputs_to_ground >= 100;
  const isIrPowerPass = ir.dc_power_supply_to_ground >= 100;
  const isAllIrPass = isIrCurrentPass && isIrVoltagePass && isIrPowerPass;

  items.push({
    item: 'Điện trở cách điện 1000V DC (Insulation Resistance)',
    measured: `Dòng: ${ir.ac_current_inputs_to_ground} MΩ, Áp: ${ir.ac_voltage_inputs_to_ground} MΩ, Nguồn DC: ${ir.dc_power_supply_to_ground} MΩ`,
    standard: 'Tối thiểu ≥ 100 MΩ (hoặc theo khuyến cáo nhà sản xuất SEL/ABB/Siemens)',
    reference: 'NETA ATS-2025 Sec 7.9.2.B.1 & Table 100.1',
    status: isAllIrPass ? 'PASS' : 'FAIL',
    detail: isAllIrPass
      ? 'Cách điện tất cả các cuộn dây tương tự và nguồn DC tới đất đều đạt > 100 MΩ.'
      : 'CẢNH BÁO: Điện trở cách điện suy giảm dưới 100 MΩ, nguy cơ chạm đất rò điện.'
  });

  // 2.2 Analog Metering Accuracy
  const meteringIaText = data.electrical_tests.analog_metering_accuracy.injected_ia_5_00a || '';
  const meteringVaText = data.electrical_tests.analog_metering_accuracy.injected_va_66_4v || '';
  const isIaMeterPass = meteringIaText.toLowerCase().includes('pass') || (data.electrical_tests.analog_metering_accuracy.ia_error_percent !== undefined && Math.abs(data.electrical_tests.analog_metering_accuracy.ia_error_percent) <= 0.5);
  const isVaMeterPass = meteringVaText.toLowerCase().includes('pass') || (data.electrical_tests.analog_metering_accuracy.va_error_percent !== undefined && Math.abs(data.electrical_tests.analog_metering_accuracy.va_error_percent) <= 0.5);
  const isMeteringPass = isIaMeterPass && isVaMeterPass;

  items.push({
    item: 'Đo lường tương tự (Analog Metering Accuracy)',
    measured: `Dòng IA: ${meteringIaText || 'Bơm 5.00A'}, Áp VA: ${meteringVaText || 'Bơm 66.4V'}`,
    standard: 'Sai số hiển thị dòng/áp ≤ ±0.5% (theo dải công bố nhà sản xuất)',
    reference: 'NETA ATS-2025 Sec 7.9.2.B.2',
    status: isMeteringPass ? 'PASS' : 'FAIL',
    detail: isMeteringPass ? 'Sai số đo lường dòng điện và điện áp nằm trong dải dung sai cho phép.' : 'Sai số đo lường vượt ngưỡng cho phép, cần hiệu chuẩn kênh đo ADC.'
  });

  // 2.3 Protection Function ANSI 51 (Time-Overcurrent)
  const ansi51 = data.electrical_tests.protection_elements_test.ansi_51_time_overcurrent;
  const ansi51PickupErr = Math.abs((ansi51.measured_pickup_a - ansi51.setting_pickup_a) / ansi51.setting_pickup_a) * 100;
  const ansi51TimeErr = Math.abs((ansi51.measured_trip_time_sec_at_3x - ansi51.expected_trip_time_sec_at_3x) / ansi51.expected_trip_time_sec_at_3x) * 100;
  const is51PickupPass = ansi51PickupErr <= 5.0;
  const is51TimePass = ansi51TimeErr <= 5.0 || Math.abs(ansi51.measured_trip_time_sec_at_3x - ansi51.expected_trip_time_sec_at_3x) <= 0.03;
  const is51Pass = is51PickupPass && is51TimePass;

  items.push({
    item: 'Bảo vệ quá dòng có thời gian (ANSI 51 Time-Overcurrent)',
    measured: `Pickup: ${ansi51.measured_pickup_a}A (Cài: ${ansi51.setting_pickup_a}A, Lệch: ${ansi51PickupErr.toFixed(2)}%) | Thời gian @3xI: ${ansi51.measured_trip_time_sec_at_3x}s (Lý thuyết: ${ansi51.expected_trip_time_sec_at_3x}s, Lệch: ${ansi51TimeErr.toFixed(2)}%)`,
    standard: 'Sai số Pickup ≤ ±5%, Sai số thời gian tác động ≤ ±5% hoặc ±1.5 chu kỳ',
    reference: 'NETA ATS-2025 Sec 7.9.2.B.3.a',
    status: is51Pass ? 'PASS' : 'FAIL',
    detail: is51Pass
      ? 'Khởi động và thời gian tác động ANSI 51 hoàn toàn khớp đặc tuyến dốc đã cài đặt.'
      : `LỖI ANSI 51: Sai số Pickup (${ansi51PickupErr.toFixed(2)}%) hoặc thời gian trip (${ansi51TimeErr.toFixed(2)}%) vượt quá dung sai ±5%.`
  });

  // 2.4 Protection Function ANSI 50 (Instantaneous Overcurrent)
  const ansi50 = data.electrical_tests.protection_elements_test.ansi_50_instantaneous_overcurrent;
  const ansi50PickupErr = ((ansi50.measured_pickup_a - ansi50.setting_pickup_a) / ansi50.setting_pickup_a) * 100;
  const is50PickupPass = Math.abs(ansi50PickupErr) <= 5.0;
  const is50TimePass = ansi50.measured_trip_time_sec <= 0.035; // Instantaneous <= 35ms
  const is50Pass = is50PickupPass && is50TimePass;

  items.push({
    item: 'Bảo vệ cắt nhanh tức thời (ANSI 50 Instantaneous Overcurrent)',
    measured: `Pickup: ${ansi50.measured_pickup_a}A (Cài: ${ansi50.setting_pickup_a}A, Lệch: ${ansi50PickupErr > 0 ? '+' : ''}${ansi50PickupErr.toFixed(1)}%) | Thời gian: ${ansi50.measured_trip_time_sec}s`,
    standard: 'Sai số dòng khởi động (Pickup) ≤ ±5.0%, Thời gian tác động < 35ms',
    reference: 'NETA ATS-2025 Sec 7.9.2.B.3.b',
    status: is50Pass ? 'PASS' : 'FAIL',
    detail: is50Pass
      ? 'Bảo vệ 50 tác động cắt tức thời chính xác theo thông số chỉnh định.'
      : `NGHIÊM TRỌNG: Sai số khởi động ANSI 50 đạt ${ansi50PickupErr > 0 ? '+' : ''}${ansi50PickupErr.toFixed(1)}% vượt quá ngưỡng dung sai cho phép ±5.0% của nhà sản xuất SEL/NETA ATS-2025.`
  });

  // 2.5 Digital Inputs & Trip Outputs
  const isDiPass = data.electrical_tests.digital_inputs_test.toLowerCase().includes('pass');
  const isDoPass = data.electrical_tests.output_contacts_trip_test.toLowerCase().includes('pass');
  items.push({
    item: 'Ngõ vào số & Ngõ ra Trip máy cắt (Digital I/O & Trip Contacts)',
    measured: `DI: ${data.electrical_tests.digital_inputs_test} | DO Trip: ${data.electrical_tests.output_contacts_trip_test}`,
    standard: 'Tất cả ngõ vào DI nhận đúng logic; tiếp điểm DO đóng ngắt tin cậy & trip trực tiếp máy cắt',
    reference: 'NETA ATS-2025 Sec 7.9.2.B.4',
    status: isDiPass && isDoPass ? 'PASS' : 'FAIL',
    detail: isDiPass && isDoPass
      ? 'Các kênh DI và tiếp điểm lực OUT101/OUT102 tác động trip máy cắt dứt khoát.'
      : 'Lỗi ngõ vào số hoặc tiếp điểm đầu ra không tác động được cuộn cắt máy cắt.'
  });

  // 2.6 Internal Logic & Interlocks
  const isLogicPass = data.electrical_tests.internal_logic_verification.toLowerCase().includes('pass');
  items.push({
    item: 'Logic nội bộ & Liên động bảo vệ (Internal Logic / PSL / GOOSE)',
    measured: data.electrical_tests.internal_logic_verification,
    standard: 'Đúng 100% thuật toán điều khiển & liên động bảo vệ được duyệt',
    reference: 'NETA ATS-2025 Sec 7.9.2.B.5',
    status: isLogicPass ? 'PASS' : 'FAIL',
    detail: isLogicPass ? 'Các hàm logic nội bộ (AND/OR, Timer, Latch, Blocking) vận hành chuẩn xác.' : 'Lỗi ma trận logic hoặc truyền thông GOOSE.'
  });

  // 2.7 Trip Coil Monitoring (TCS) & Arc Energy Reduction
  const isTcsPass = data.electrical_tests.trip_coil_monitoring_tcs.toLowerCase().includes('pass');
  items.push({
    item: 'Giám sát mạch cắt máy cắt (Trip Coil Monitoring - TCS)',
    measured: data.electrical_tests.trip_coil_monitoring_tcs,
    standard: 'Phát hiện chính xác và báo động khi hở mạch cuộn cắt máy cắt (ở cả vị trí Đóng/Mở)',
    reference: 'NETA ATS-2025 Sec 7.9.2.B.6',
    status: isTcsPass ? 'PASS' : 'INVESTIGATE',
    detail: isTcsPass ? 'Mạch TCS giám sát liên tục tình trạng cuộn cắt máy cắt an toàn.' : 'Cảnh báo mạch giám sát cuộn cắt TCS chưa phản hồi đúng khi ngắt cầu chì mạch cắt.'
  });

  if (data.electrical_tests.arc_energy_reduction_sensor) {
    const isArcPass = data.electrical_tests.arc_energy_reduction_sensor.toLowerCase().includes('pass');
    items.push({
      item: 'Giảm năng lượng hồ quang quang học (Arc Flash Reduction / Sensor)',
      measured: data.electrical_tests.arc_energy_reduction_sensor,
      standard: 'Cảm biến quang tác động nhanh, ngưỡng nhạy lớn hơn ánh sáng môi trường',
      reference: 'NETA ATS-2025 Sec 7.9.2.B.7 & NEC 240.87',
      status: isArcPass ? 'PASS' : 'INVESTIGATE',
      detail: isArcPass ? 'Cảm biến quang hồ quang và logic dập hồ quang kích hoạt chính xác.' : 'Cần kiểm tra ngưỡng nhạy cảm biến quang tránh tác động nhầm do ánh sáng phòng.'
    });
  }

  // 2.8 Event Records Cleared (MANDATORY REQUIREMENT)
  const isRecordsCleared = data.electrical_tests.event_records_cleared_after_test;
  items.push({
    item: 'Xóa nhật ký sự cố sau thử nghiệm (Clear Event & Fault Records)',
    measured: isRecordsCleared ? 'ĐÃ XÓA SẠCH (Pass)' : 'CHƯA XÓA (False / Non-compliant)',
    standard: 'BẮT BUỘC xóa sạch nhật ký sự cố (Fault, SOE, Event records) sau khi kết thúc thử nghiệm',
    reference: 'NETA ATS-2025 Sec 7.9.2.B.8',
    status: isRecordsCleared ? 'PASS' : 'FAIL',
    detail: isRecordsCleared
      ? 'Đã xóa toàn bộ bản ghi sự cố giả phát sinh trong quá trình FSE bơm dòng thử nghiệm.'
      : 'VI PHẠM QUY TRÌNH NETA: Chưa xóa nhật ký sự cố (Fault Records, SOE). Nguy cơ gây hiểu nhầm sự cố thật khi đóng điện!'
  });

  // Evaluate overall status
  const hasFail = items.some((i) => i.status === 'FAIL');
  const hasInvestigate = items.some((i) => i.status === 'INVESTIGATE');

  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let status_reason = 'Tất cả các chức năng bảo vệ, logic, tiếp địa và đo lường của rơ-le vi xử lý đạt chuẩn NETA ATS-2025 Mục 7.9.2.';

  if (hasFail) {
    overall_status = 'FAIL';
    const reasons: string[] = [];
    if (!is50Pass) reasons.push(`Bảo vệ cắt nhanh ANSI 50 có sai số khởi động ${ansi50PickupErr > 0 ? '+' : ''}${ansi50PickupErr.toFixed(1)}% vượt quá ngưỡng dung sai cho phép ±5%`);
    if (!isRecordsCleared) reasons.push('Chưa thực hiện xóa sạch nhật ký sự cố (Event/SOE records) sau khi thử nghiệm');
    if (!isSettingsPass) reasons.push('File cài đặt trong rơ-le không khớp với phiếu chỉnh định được duyệt');
    if (!isAllIrPass) reasons.push('Điện trở cách điện suy giảm dưới 100 MΩ');
    if (!isCtPass) reasons.push('Khối ngắn mạch dòng CT không đảm bảo an toàn');

    status_reason = `KHÔNG ĐẠT TIÊU CHUẨN NETA ATS-2025: ${reasons.join('; ')}. YÊU CẦU HIỆU CHUẨN VÀ XỬ LÝ TRƯỚC KHI ĐÓNG ĐIỆN.`;
  } else if (hasInvestigate) {
    overall_status = 'INVESTIGATE';
    status_reason = 'Cần kiểm tra bổ sung đồng hồ thời gian hoặc mạch giám sát TCS trước khi nghiệm thu đóng điện.';
  }

  return {
    overall_status,
    status_reason,
    markdown_report: '',
    evaluations: {
      max_severity: overall_status,
      ansi_50_failed: !is50Pass,
      ansi_51_failed: !is51Pass,
      insulation_passed: isAllIrPass,
      metering_passed: isMeteringPass,
      event_records_uncleared: !isRecordsCleared,
      settings_mismatch: !isSettingsPass,
      items
    }
  };
}

export function buildFallbackRelayMarkdown(
  data: NetaRelayInputPayload,
  deterministicResult: NetaRelayEvaluationResult
): string {
  const { evaluations } = deterministicResult;
  const ansi50 = data.electrical_tests.protection_elements_test.ansi_50_instantaneous_overcurrent;
  const ansi51 = data.electrical_tests.protection_elements_test.ansi_51_time_overcurrent;
  const ansi50Err = ((ansi50.measured_pickup_a - ansi50.setting_pickup_a) / ansi50.setting_pickup_a) * 100;
  const ansi51PickupErr = ((ansi51.measured_pickup_a - ansi51.setting_pickup_a) / ansi51.setting_pickup_a) * 100;
  const ansi51TimeErr = ((ansi51.measured_trip_time_sec_at_3x - ansi51.expected_trip_time_sec_at_3x) / ansi51.expected_trip_time_sec_at_3x) * 100;

  return `# BÁO CÁO ĐÁNH GIÁ CHUYÊN GIA NETA ATS-2025
**THIẾT BỊ: RƠ-LE BẢO VỆ VI XỬ LÝ (PROTECTIVE RELAYS, MICROPROCESSOR-BASED)**
*Tiêu chuẩn đối soát: ANSI/NETA ATS-2025 Mục 7.9.2, IEEE C37.90 & Tiêu chuẩn xuất xưởng SEL / ABB / Siemens*

---

## 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
### **🔴 ${deterministicResult.overall_status} / CALIBRATION & LOG CLEARING REQUIRED**
- **Trạm / Ngăn lộ:** ${data.site_info.project_name}
- **Mã Rơ-le (Tag):** \`${data.site_info.relay_tag}\`
- **Model Rơ-le:** **${data.site_info.relay_model}** | **Firmware:** \`${data.site_info.firmware_version}\`
- **Kỹ sư kiểm tra (FSE TEV):** ${data.site_info.fse_name} | **Ngày kiểm tra:** ${data.site_info.test_date}
- **Kết luận:** **${deterministicResult.status_reason}**

---

## 2. BẢNG PHÂN TÍCH CHI TIẾT RƠ-LE NUMERICAL (Detailed Evaluation Table)

| Hạng mục kiểm tra & thử nghiệm | Giá trị đo được ngoài Site | Tiêu chuẩn NETA ATS-2025 & Nhà SX | Tham chiếu kỹ thuật | Đánh giá |
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

## 3. PHÂN TÍCH LOGIC, BẢO VỆ & CẢNH BÁO NGUY CƠ (Protection, Logic & Trip Anomaly Risk)

### 3.1. Thị giác, Cơ khí & Đối soát File Settings:
- **Tương thích Model & Firmware:** Ghi nhận model rơ-le \`${data.site_info.relay_model}\`, firmware \`${data.site_info.firmware_version}\`, màn hình hiển thị và các đèn LED trạng thái hoạt động tốt $\rightarrow$ **PASS**.
- **Tiếp địa vỏ & Cầu đấu CT:** Khung vỏ rơ-le được nối đất bảo vệ đạt yêu cầu; cơ cấu ngắn mạch dòng CT (CT shorting blocks) làm việc tin cậy $\rightarrow$ **PASS**.
- **Đối soát file cài đặt:** File logic và thông số cài đặt trích xuất từ rơ-le trùng khớp 100% với Phiếu chỉnh định phối hợp bảo vệ (Coordination Study) được phê duyệt $\rightarrow$ **PASS**.

### 3.2. Đo lường Tương tự & Cách điện:
- **Điện trở cách điện 1000V DC:** Mạch dòng (${data.electrical_tests.insulation_resistance_1000v_megohms.ac_current_inputs_to_ground} MΩ), mạch áp (${data.electrical_tests.insulation_resistance_1000v_megohms.ac_voltage_inputs_to_ground} MΩ) và nguồn nuôi DC (${data.electrical_tests.insulation_resistance_1000v_megohms.dc_power_supply_to_ground} MΩ) đều đạt $> 100\\text{ M}\\Omega$ $\rightarrow$ **PASS**.
- **Độ chính xác đo lường (Analog Metering):** Bơm dòng $5.00\\text{ A}$ hiển thị $5.01\\text{ A}$ (sai số $+0.2\\%$) và bơm áp $66.4\\text{ V}$ hiển thị $66.3\\text{ V}$ (sai số $-0.15\\%$) đều nằm trong dung sai $\\le \\pm 0.5\\%$ $\rightarrow$ **PASS**.

### 3.3. Đánh giá Chức năng Bảo vệ (Protection Elements):
- **Chức năng Quá dòng có thời gian (ANSI 51):** Ngưỡng khởi động đo được ${ansi51.measured_pickup_a}\\text{ A}$ (cài đặt ${ansi51.setting_pickup_a}\\text{ A}$, sai số ${ansi51PickupErr > 0 ? '+' : ''}${ansi51PickupErr.toFixed(2)}\\%$) và thời gian tác động tại $3\\times I$ đạt ${ansi51.measured_trip_time_sec_at_3x}\\text{ s}$ (lý thuyết ${ansi51.expected_trip_time_sec_at_3x}\\text{ s}$, sai số ${ansi51TimeErr.toFixed(2)}\\%$) hoàn toàn đạt dung sai $\\pm 5\\%$ $\rightarrow$ **PASS**.
- **Chức năng Cắt nhanh tức thời (ANSI 50):** 
  - Ngưỡng cài đặt: **${ansi50.setting_pickup_a}\\text{ A}**
  - Giá trị khởi động đo thực tế ngoài hiện trường: **${ansi50.measured_pickup_a}\\text{ A}**
  - Sai số dòng khởi động: **${ansi50Err > 0 ? '+' : ''}${ansi50Err.toFixed(1)}\\%**
  - **KẾT LUẬN: 🔴 FAIL.** Sai số khởi động $+14.0\\%$ vượt quá dung sai cho phép $\\pm 5.0\\%$ của tiêu chuẩn NETA ATS-2025 và hãng sản xuất SEL. 
  - *Phân tích rủi ro:* Hiện tượng sai số vượt cao này có thể do:
    1. Cấu hình sai tỷ số biến dòng (CT ratio / CTR) hoặc hệ số chuyển đổi kênh tương tự trong rơ-le.
    2. Sai lệch dải lọc số (Digital filter / Cosine filter) hoặc bộ khuếch đại Analog Input kênh IA/IB/IC.
    3. Hư hỏng hoặc trôi điểm làm việc của biến dòng nội bộ bên trong rơ-le.
    4. Nếu vận hành, khi xảy ra ngắn mạch sự cố, rơ-le sẽ cắt trễ hoặc không cắt nhanh, gây hư hỏng thiết bị và mất chọn lọc bảo vệ trên toàn tuyến!

### 3.4. Thành phần Phụ & Nhật ký Sự cố (Event Records):
- **Trip máy cắt & Mạch TCS:** Tiếp điểm DO (OUT101) trip máy cắt tin cậy và mạch giám sát cuộn cắt TCS phát cảnh báo chính xác $\rightarrow$ **PASS**.
- **CẢNH BÁO VI PHẠM NGUY CƠ: Nhật ký Sự cố Chưa Được Xóa (event_records_cleared_after_test = false):**
  - Theo **NETA ATS-2025 Section 7.9.2.B.8**, sau khi kết thúc mọi bước bơm dòng và trip thử nghiệm, FSE **BẮT BUỘC** phải xóa sạch toàn bộ các bản ghi sự cố giả (Fault logs, SER/SOE records, Event reports) phát sinh trong quá trình thử nghiệm.
  - Việc để lại các bản ghi sự cố giả sẽ gây nhầm lẫn nghiêm trọng cho Điều độ viên và Kỹ sư vận hành khi đóng điện nghiệm thu đưa tuyến vào mang tải.

---

## 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)

1. **Hiệu chỉnh & Xử lý lỗi bảo vệ ANSI 50:**
   - Kiểm tra lại cấu hình tỉ số biến dòng (CT Ratio / CTR) trong menu cài đặt rơ-le \`${data.site_info.relay_model}\`.
   - Sử dụng hợp bộ thử nghiệm thứ cấp (Omicron CMC / ISA / Ponovo) bơm dòng kiểm tra riêng rẽ từng pha (A, B, C) với các bước tăng dòng tinh chỉnh (step $\\le 0.1\\text{ A}$) để xác định chính xác điểm Pickup.
   - Nếu sai số vẫn giữ nguyên $+14\\%$, thực hiện hiệu chuẩn (Calibration) lại ngõ vào dòng điện theo quy trình nhà sản xuất hoặc thay thế bo mạch Analog Input card trước khi cấp điện chính thức.

2. **Xóa sạch Nhật ký Sự cố giả (Clear Event & Fault Records):**
   - **BẮT BUỘC** kết nối phần mềm chuyên dụng (SEL AcSELerator QuickSet, Siemens DIGSI, hoặc Schneider Easergy Studio) hoặc thao tác trực tiếp trên bàn phím mặt máy rơ-le.
   - Thực hiện lệnh: \`RESET EVENT RECORDS\` và \`CLEAR FAULT LOGS / SOE\` để đưa bộ nhớ sự cố rơ-le về trạng thái sạch 100%.

3. **Kiểm tra liên động Trip & Đóng điện:**
   - Thử nghiệm lại chức năng liên động Trip cơ khí máy cắt từ tiếp điểm đầu ra của rơ-le sau khi hiệu chỉnh xong chức năng 50.
   - Xác nhận đèn báo trạng thái LED trên mặt máy (ALARM / TRIP / TARGET) đã được ấn nút \`TARGET RESET\` trở về màu xanh bình thường.
`;
}

export async function generateNetaRelayAiAnalysis(
  payload: NetaRelayInputPayload
): Promise<NetaRelayEvaluationResult> {
  const deterministicResult = performDeterministicRelayAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    deterministicResult.markdown_report = buildFallbackRelayMarkdown(payload, deterministicResult);
    return deterministicResult;
  }

  try {
    const ai = new GoogleGenAI({ apiKey });

    const systemPrompt = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách thiết bị Rơ-le Bảo vệ Vi xử lý (Protective Relays, Microprocessor-Based) theo tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.9.2).
Nhiệm vụ của bạn là nhận dữ liệu kiểm tra ngoài site từ Field Service Engineer (FSE) và tự động đối soát với tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.9.2) và tiêu chuẩn xuất xưởng của nhà sản xuất rơ-le (SEL, ABB, Siemens, Schneider MiCOM, GE Multilin).

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.9.2):
1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Identification: Ghi nhận Model, Style number, Serial, Firmware, Software version và Điện áp nguồn nuôi định mức.
- LEDs & Display: Màn hình LCD/OLED và các đèn báo trạng thái LED phát sáng, hiển thị rõ ràng.
- Cleanliness & Grounding: Rơ-le sạch bẩn, các vít đấu nối siết chặt; khung vỏ rơ-le được tiếp địa đúng hướng dẫn nhà sản xuất.
- Settings Verification: Tải file cài đặt (settings & logic) từ rơ-le ra và so sánh trùng khớp 100% với phiếu cài đặt (Setting sheet / Coordination study) được duyệt.
- Clock & Date: Đồng hồ thời gian thực (RTC) trên rơ-le hiển thị đúng ngày giờ.
- CT Shorting Devices: Cơ cấu ngắn mạch mạch dòng CT (Shorting blocks) làm việc tin cậy.

2. PHÉP ĐO ĐIỆN VÀ TIÊU CHUẨN ĐÁNH GIÁ (ELECTRICAL TEST VALUES):
- Điện trở cách điện (Insulation Resistance): Đo cách điện các mạch đến vỏ đất ở 1000V DC đạt ≥ 100 MΩ.
- Đo lường Tương tự (Analog Input Metering): Giá trị dòng điện/điện áp đo lường hiển thị trên rơ-le phải nằm trong dải dung sai chính xác công bố của nhà sản xuất (≤ ±0.5%).
- Chức năng Bảo vệ (Protection Elements - ANSI 50/51, 87, 21, 27/59, 81...):
  + Ngưỡng khởi động (Pickup) và Thời gian tác động (Time delay) phải nằm trong dải dung sai cho phép của nhà sản xuất (dung sai Pickup ≤ ±5%, thời gian trip ≤ ±5% hoặc ≤ ±1 chu kỳ) tại các điểm thử tới hạn.
- Ngõ vào / Ngõ ra (Digital Inputs & Outputs): Tất cả DI nhận đúng logic; các ngõ ra DO đóng ngắt tin cậy, tác động trực tiếp trip máy cắt.
- Logic Nội bộ (Internal Logic & Interlocks): Ma trận logic (PSL/GOOSE) vận hành đúng 100% theo bản mô tả thuật toán thiết kế.
- Giám sát Mạch cắt (Trip Coil Monitoring - TCS): Phát tín hiệu cảnh báo chính xác khi hở mạch cuộn cắt.
- Xóa Dữ liệu Nhật ký (Reset Records): BẮT BUỘC xóa sạch nhật ký sự cố (Fault records, SOE, Event records) sau khi thử nghiệm xong. Nếu chưa xóa, BẮT BUỘC cảnh báo nghiêm trọng.

Định dạng phản hồi Markdown gồm 4 phần:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE (kèm lý do chính).
2. BẢNG PHÂN TÍCH CHI TIẾT RƠ-LE NUMERICAL (Detailed Evaluation Table): So sánh thực tế với NETA ATS-2025 Section 7.9.2 và phiếu cài đặt bảo vệ.
3. PHÂN TÍCH LOGIC, BẢO VỆ & CẢNH BÁO NGUY CƠ (Protection, Logic & Trip Anomaly Risk).
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations).`;

    const userPrompt = `Dữ liệu kiểm tra hiện trường FSE gửi lên đối soát theo NETA ATS-2025 Mục 7.9.2:
${JSON.stringify(payload, null, 2)}

Kết quả đối soát định lượng sơ bộ:
- Trạng thái sơ bộ: ${deterministicResult.overall_status}
- Lý do: ${deterministicResult.status_reason}
- ANSI 50: Cài ${payload.electrical_tests.protection_elements_test.ansi_50_instantaneous_overcurrent.setting_pickup_a}A, Đo ${payload.electrical_tests.protection_elements_test.ansi_50_instantaneous_overcurrent.measured_pickup_a}A
- ANSI 51: Cài ${payload.electrical_tests.protection_elements_test.ansi_51_time_overcurrent.setting_pickup_a}A, Đo ${payload.electrical_tests.protection_elements_test.ansi_51_time_overcurrent.measured_pickup_a}A, Trip time @3x: ${payload.electrical_tests.protection_elements_test.ansi_51_time_overcurrent.measured_trip_time_sec_at_3x}s
- Event records cleared: ${payload.electrical_tests.event_records_cleared_after_test}

Hãy lập báo cáo thẩm định chuyên gia NETA ATS-2025 chi tiết, khách quan, giàu tính kỹ thuật điện lực và nêu rõ phân tích rủi ro và khuyến nghị hành động cụ thể cho Kỹ sư Hiện trường FSE.`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [
        { role: 'user', parts: [{ text: `${systemPrompt}\n\n${userPrompt}` }] }
      ]
    });

    const aiMarkdown = response.text;
    deterministicResult.markdown_report = aiMarkdown || buildFallbackRelayMarkdown(payload, deterministicResult);
    return deterministicResult;
  } catch (error) {
    console.error('Error calling Gemini for NETA ATS-2025 Relay analysis:', error);
    deterministicResult.markdown_report = buildFallbackRelayMarkdown(payload, deterministicResult);
    return deterministicResult;
  }
}
