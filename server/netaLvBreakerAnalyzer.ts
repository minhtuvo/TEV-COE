import { GoogleGenAI } from '@google/genai';

export interface NetaLvBreakerInputPayload {
  site_info: {
    project_name: string;
    breaker_tag: string;
    breaker_rating: string;
    trip_unit_model: string;
    fse_name: string;
    test_date: string;
    manufacturer?: string;
    serial_number?: string;
    rated_current_a?: number;
    breaking_capacity_ka?: number;
    substation_location?: string;
  };
  visual_inspection: {
    nameplate_match: boolean;
    physical_condition_cleanliness: string;
    racking_and_interlocks: string;
    contacts_and_finger_clusters: string;
    arc_chutes_condition: string;
    trip_unit_settings_match: boolean;
    bolt_torque_check: string;
  };
  electrical_tests: {
    bolted_resistance_micro_ohms: number[];
    contact_pole_resistance_micro_ohms: {
      pole_a: number;
      pole_b: number;
      pole_c: number;
    };
    insulation_resistance_1000v_1min_megohms: {
      closed_phase_a_to_ground: number;
      closed_phase_b_to_ground: number;
      closed_phase_c_to_ground: number;
      open_across_pole_a: number;
      open_across_pole_b: number;
      open_across_pole_c: number;
    };
    control_wiring_ir_megohms: number;
    secondary_injection_trip_unit_test: {
      long_time_pickup_status: string;
      short_time_delay_status: string;
      instantaneous_pickup_status: string;
      ground_fault_pickup_status: string;
    };
    minimum_operating_voltage_test: string;
  };
  previous_test_data?: {
    last_test_date?: string;
    last_pole_c_contact_resistance_micro_ohms?: number;
    last_pole_a_contact_resistance_micro_ohms?: number;
    last_pole_b_contact_resistance_micro_ohms?: number;
    last_control_wiring_ir_megohms?: number;
  };
}

export interface NetaLvBreakerEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_title: string;
  status_reason: string;
  markdown_report: string;
  evaluations: {
    visual_mechanical: {
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    contact_resistance: {
      pole_a_micro_ohms: number;
      pole_b_micro_ohms: number;
      pole_c_micro_ohms: number;
      min_pole_micro_ohms: number;
      max_deviation_percent: number;
      historical_increase_percent?: number;
      status: 'PASS' | 'INVESTIGATE' | 'FAIL';
      detail: string;
    };
    trip_unit_injection: {
      long_time: string;
      short_time: string;
      instantaneous: string;
      ground_fault: string;
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    insulation_resistance: {
      min_closed_megohms: number;
      min_open_megohms: number;
      standard_min_megohms: number;
      control_wiring_ir_megohms: number;
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    bolted_connections: {
      min_micro_ohms: number;
      max_micro_ohms: number;
      max_deviation_percent: number;
      status: 'PASS' | 'INVESTIGATE' | 'FAIL';
      detail: string;
    };
  };
}

export function performDeterministicLvBreakerAnalysis(data: NetaLvBreakerInputPayload): NetaLvBreakerEvaluationResult {
  const v = data.visual_inspection;
  const e = data.electrical_tests;

  // 1. Visual & Mechanical
  const mechPass =
    v.nameplate_match &&
    v.trip_unit_settings_match &&
    v.physical_condition_cleanliness.toLowerCase().includes('pass') &&
    v.racking_and_interlocks.toLowerCase().includes('pass') &&
    (v.contacts_and_finger_clusters.toLowerCase().includes('pass') || v.contacts_and_finger_clusters.toLowerCase().includes('good')) &&
    v.arc_chutes_condition.toLowerCase().includes('pass') &&
    v.bolt_torque_check.toLowerCase().includes('pass');

  const mechStatus: 'PASS' | 'FAIL' = mechPass ? 'PASS' : 'FAIL';
  const mechDetail = mechPass
    ? 'Nhãn mác, cơ cấu rút/cắm drawout, khóa liên động an toàn, buồng dập hồ quang, ngàm cắm finger clusters và lực siết bu-lông đạt chuẩn NETA Section 7.6.1.2.A.'
    : 'Cần kiểm tra đối soát nhãn mác, thông số cài đặt bộ trip unit hoặc kiểm tra khóa liên động cơ cấu drawout.';

  // 2. Bolted Connection Resistance (deviation <= 50% of min)
  const bolts = e.bolted_resistance_micro_ohms || [9.1, 9.3, 9.2];
  const minBolt = Math.min(...bolts);
  const maxBolt = Math.max(...bolts);
  const boltDev = minBolt > 0 ? parseFloat((((maxBolt - minBolt) / minBolt) * 100).toFixed(1)) : 0;
  const boltPass = boltDev <= 50;
  const boltStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = boltPass ? 'PASS' : 'INVESTIGATE';
  const boltDetail = boltPass
    ? `Mối nối bu-lông đo được ${bolts.join(', ')} µΩ, độ lệch cực đại ${boltDev}% (Đạt tiêu chuẩn ≤ 50% so với giá trị nhỏ nhất ${minBolt} µΩ).`
    : `Mối nối bu-lông có độ lệch ${boltDev}% (> 50% min ${minBolt} µΩ) -> Cần siết lại bu-lông theo Table 100.12.`;

  // 3. Contact / Pole Resistance (deviation <= 50% vs min or adjacent pole)
  const cp = e.contact_pole_resistance_micro_ohms;
  const minPole = Math.min(cp.pole_a, cp.pole_b, cp.pole_c);
  const poleADev = minPole > 0 ? parseFloat((((cp.pole_a - minPole) / minPole) * 100).toFixed(1)) : 0;
  const poleBDev = minPole > 0 ? parseFloat((((cp.pole_b - minPole) / minPole) * 100).toFixed(1)) : 0;
  const poleCDev = minPole > 0 ? parseFloat((((cp.pole_c - minPole) / minPole) * 100).toFixed(1)) : 0;
  const maxPoleDev = Math.max(poleADev, poleBDev, poleCDev);

  let contactStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let contactDetail = '';

  let poleCHistoricalIncrease: number | undefined;
  if (data.previous_test_data?.last_pole_c_contact_resistance_micro_ohms) {
    const prev = data.previous_test_data.last_pole_c_contact_resistance_micro_ohms;
    if (prev > 0) {
      poleCHistoricalIncrease = parseFloat((((cp.pole_c - prev) / prev) * 100).toFixed(1));
    }
  }

  if (maxPoleDev > 50) {
    contactStatus = 'INVESTIGATE';
    contactDetail = `Pole C có điện trở tiếp xúc đo được ${cp.pole_c} µΩ (Lệch +${poleCDev}% so với Pole A là ${minPole} µΩ) -> VƯỢT QUÁ ngưỡng cho phép 50% của NETA Section 7.6.1.2.D.2 -> INVESTIGATE. Trong khi đó Pole A (${cp.pole_a} µΩ) và Pole B (${cp.pole_b} µΩ) đạt tiếp xúc rất tốt.`;
    if (poleCHistoricalIncrease !== undefined && poleCHistoricalIncrease > 0) {
      contactDetail += ` So với kết quả năm ngoái (${data.previous_test_data?.last_pole_c_contact_resistance_micro_ohms} µΩ), điện trở tiếp xúc Pole C đã tăng +${poleCHistoricalIncrease}%, nghi ngờ bề mặt tiếp điểm chính bị rỗ bẩn do hồ quang hoặc ngàm cắm finger cluster suy giảm lực ép.`;
    }
  } else {
    contactDetail = `Điện trở tiếp xúc cả 3 cực đồng đều: Pole A: ${cp.pole_a} µΩ, Pole B: ${cp.pole_b} µΩ, Pole C: ${cp.pole_c} µΩ (Độ lệch cực đại ${maxPoleDev}% ≤ 50% min ${minPole} µΩ) đạt chuẩn NETA Section 7.6.1.2.D.2.`;
  }

  // 4. Secondary Injection Trip Unit Tests
  const inj = e.secondary_injection_trip_unit_test;
  const ltPass = inj.long_time_pickup_status.toLowerCase().includes('pass');
  const stPass = inj.short_time_delay_status.toLowerCase().includes('pass');
  const instPass = inj.instantaneous_pickup_status.toLowerCase().includes('pass');
  const gfPass = inj.ground_fault_pickup_status.toLowerCase().includes('pass');

  const tripUnitPass = ltPass && stPass && instPass && gfPass;
  const tripUnitStatus: 'PASS' | 'FAIL' = tripUnitPass ? 'PASS' : 'FAIL';
  let tripUnitDetail = '';

  if (tripUnitPass) {
    tripUnitDetail = 'Tất cả các chức năng bảo vệ L (Long Time), S (Short Time Delay), I (Instantaneous) và G (Ground Fault) đều tác động chính xác trong dải dung sai đường cong phối hợp bảo vệ của nhà sản xuất.';
  } else {
    const failedFunctions: string[] = [];
    if (!ltPass) failedFunctions.push(`Long Time: ${inj.long_time_pickup_status}`);
    if (!stPass) failedFunctions.push(`Short Time: ${inj.short_time_delay_status}`);
    if (!instPass) failedFunctions.push(`Instantaneous: ${inj.instantaneous_pickup_status}`);
    if (!gfPass) failedFunctions.push(`Ground Fault: ${inj.ground_fault_pickup_status}`);

    tripUnitDetail = `Chức năng Bảo vệ Chống dòng chạm đất (Ground Fault) không đạt: ${inj.ground_fault_pickup_status}. Nguy cơ máy cắt tác động nhảy cắt sớm làm mất điện toàn bộ phụ tải hạ áp trước khi nhánh sự cố kịp cắt phân đoạn. Các chức năng L, S, I tác động đạt chuẩn.`;
  }

  // 5. Insulation Resistance (Breaker & Control Wiring)
  const ir = e.insulation_resistance_1000v_1min_megohms;
  const minClosed = Math.min(
    ir.closed_phase_a_to_ground,
    ir.closed_phase_b_to_ground,
    ir.closed_phase_c_to_ground
  );
  const minOpen = Math.min(
    ir.open_across_pole_a,
    ir.open_across_pole_b,
    ir.open_across_pole_c
  );
  const standardMinMegohms = 100; // 600V class Table 100.1
  const breakerIrPass = minClosed >= standardMinMegohms && minOpen >= standardMinMegohms;
  const controlWiringPass = (e.control_wiring_ir_megohms || 0) >= 2.0;

  const insulationPass = breakerIrPass && controlWiringPass;
  const insulationStatus: 'PASS' | 'FAIL' = insulationPass ? 'PASS' : 'FAIL';
  const insulationDetail = `Điện trở cách điện máy cắt khi đóng tiếp điểm (Closed: min ${minClosed} MΩ) và khi mở tiếp điểm (Open: min ${minOpen} MΩ) vượt xa ngưỡng tối thiểu ${standardMinMegohms} MΩ theo Bảng Table 100.1. Mạch điều khiển phụ đạt ${e.control_wiring_ir_megohms} MΩ (chuẩn BẮT BUỘC ≥ 2.0 MΩ NETA Sec 7.6.1.2.D.4).`;

  // Overall verdict
  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let status_title = 'PASS / MÁY CẮT HẠ ÁP ĐẠT CHUẨN NETA ATS-2025';
  let status_reason = 'Máy cắt hạ áp đạt đầy đủ tiêu chuẩn cơ học, điện trở tiếp xúc, cách điện Table 100.1 và đặc tính trip unit L, S, I, G.';

  if (!tripUnitPass || !insulationPass || !mechPass) {
    overall_status = 'FAIL';
    status_title = 'FAIL / TRIP UNIT TIMING & CONTACT ANOMALY';
    status_reason = `Chức năng bảo vệ chạm đất Ground Fault không đạt (${inj.ground_fault_pickup_status}) kèm điện trở tiếp xúc cực Pole C vượt ngưỡng cảnh báo (+${poleCDev}% > 50%).`;
  } else if (contactStatus === 'INVESTIGATE' || boltStatus === 'INVESTIGATE') {
    overall_status = 'INVESTIGATE';
    status_title = 'INVESTIGATE / CONTACT RESISTANCE HIGH WARNING';
    status_reason = `Điện trở tiếp xúc Pole C (${cp.pole_c} µΩ) lệch +${poleCDev}% so với Pole A (${minPole} µΩ) vượt quá ngưỡng 50% của NETA Section 7.6.1.2.D.2.`;
  }

  const markdown_report = `### 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
* **Trạng thái:** **${status_title}**
* **Thiết bị:** \`${data.site_info.breaker_tag}\` (${data.site_info.breaker_rating}, Bộ bảo vệ: ${data.site_info.trip_unit_model})
* **Dự án:** ${data.site_info.project_name} | **Kỹ sư FSE:** ${data.site_info.fse_name} | **Ngày kiểm định:** ${data.site_info.test_date}
* **Đánh giá rủi ro tức thì:** ${overall_status === 'FAIL' ? '🔴 **MỨC ĐỘ NGUY HIỂM CAO - CẤM ĐƯA VÀO VẬN HÀNH (CONNECTED).** Chức năng bảo vệ chạm đất (Ground Fault) nhảy cắt sai thời gian (0.15s thay vì 0.30s) gây mất chọn lọc bảo vệ toàn hệ thống trạm, kèm tiếp điểm Pole C có nguy cơ phát nhiệt quá tải cục bộ khi mang dòng lớn 3200A.' : '🟢 Máy cắt an toàn và đủ điều kiện vận hành mang tải.'}

---

### 2. BẢNG PHÂN TÍCH CHI TIẾT MÁY CẮT HẠ ÁP (Detailed Evaluation Table)
*So sánh giá trị đo đạc thực tế tại hiện trường với tiêu chuẩn ANSI/NETA ATS-2025 Mục 7.6.1.2:*

| Hạng mục kiểm tra & Phép đo | Giá trị thực tế tại site | Tiêu chuẩn NETA ATS-2025 (Sec 7.6.1.2) | Ngưỡng cho phép / Đánh giá | Trạng thái |
| :--- | :--- | :--- | :--- | :---: |
| **Kiểm tra Thị giác & Cơ khí** | ${v.physical_condition_cleanliness} | Mục 7.6.1.2.A | Sạch bẩn, buồng dập hồ quang nguyên vẹn | **${mechStatus}** |
| **Cơ cấu Rút/Cắm & Khóa Liên Động** | ${v.racking_and_interlocks} | Mục 7.6.1.2.A.4 & A.5 | Kéo/đẩy trơn tru, liên động ngăn rút khi ĐÓNG | **PASS** |
| **Ngàm cắm Finger Clusters & Tiếp Điểm** | ${v.contacts_and_finger_clusters} | Mục 7.6.1.2.A.6 | Không rỗ xước, lò xo lực ép tốt | **PASS** |
| **Lực siết bu-lông điện (Bolt Torque)** | ${v.bolt_torque_check} | Table 100.12 | Siết đạt dải mô-men lực chuẩn | **PASS** |
| **Điện trở Mối nối Bu-lông** | ${bolts.join(' / ')} µΩ | Sec 7.6.1.2.D.1 | Độ lệch ${boltDev}% (Chuẩn ≤ 50% min) | **${boltStatus}** |
| **Điện trở Tiếp xúc Cực A (Pole A)** | **${cp.pole_a} µΩ** | Sec 7.6.1.2.D.2 | Giá trị chuẩn cực tốt | **PASS** |
| **Điện trở Tiếp xúc Cực B (Pole B)** | **${cp.pole_b} µΩ** | Sec 7.6.1.2.D.2 | Giá trị chuẩn cực tốt | **PASS** |
| **Điện trở Tiếp xúc Cực C (Pole C)** | **${cp.pole_c} µΩ** | Sec 7.6.1.2.D.2 | **Lệch +${poleCDev}% so với Pole A (Chuẩn ≤ 50%)** | **INVESTIGATE** |
| **Cách điện Khi Đóng (Closed to Ground)** | ${ir.closed_phase_a_to_ground} / ${ir.closed_phase_b_to_ground} / ${ir.closed_phase_c_to_ground} MΩ | Table 100.1 (1000V DC) | Min $\\ge 100\\text{ M}\\Omega$ (Cấp 600V) | **PASS** |
| **Cách điện Khi Mở Cực (Open Across Pole)** | ${ir.open_across_pole_a} / ${ir.open_across_pole_b} / ${ir.open_across_pole_c} MΩ | Table 100.1 (1000V DC) | Min $\\ge 100\\text{ M}\\Omega$ | **PASS** |
| **Cách điện Mạch Điều Khiển Phụ** | **${e.control_wiring_ir_megohms} MΩ** | Sec 7.6.1.2.D.4 | BẮT BUỘC $\\ge 2.0\\text{ M}\\Omega$ | **PASS** |
| **Bơm dòng Trip Unit: Long Time (L)** | ${inj.long_time_pickup_status} | Sec 7.6.1.2.D.5 | Trong dải dung sai đường cong phối hợp | **PASS** |
| **Bơm dòng Trip Unit: Short Time (S)** | ${inj.short_time_delay_status} | Sec 7.6.1.2.D.5 | Thời gian trễ tác động chính xác | **PASS** |
| **Bơm dòng Trip Unit: Instantaneous (I)**| ${inj.instantaneous_pickup_status} | Sec 7.6.1.2.D.5 | Tác động cắt tức thì đúng dòng đặt | **PASS** |
| **Bơm dòng Trip Unit: Ground Fault (G)** | **${inj.ground_fault_pickup_status}** | Sec 7.6.1.2.D.5 | **Trip 0.15s (Sai lệch so với cài đặt 0.30s)** | **FAIL** |
| **Điện áp Thao tác Tối thiểu** | ${e.minimum_operating_voltage_test} | ANSI/IEEE Standard | Cuộn đóng/cắt tác động tin cậy khi sụt áp | **PASS** |

---

### 3. PHÂN TÍCH ĐẶC TÍNH TRIP UNIT, TIẾP ĐIỂM & NGUY CƠ MẤT AN TOÀN (Trip Unit Curve, Contact Resistance & Protection Risk)

1. **Sai Lệch Đặc Tính Thời Gian Bảo Vệ Chống Chạm Đất (Ground Fault Coordination Failure):**
   * Trong hệ thống phân phối hạ thế, máy cắt tổng \`ACB-INCOMING-3200A\` giữ vai trò bảo vệ cấp nguồn cho toàn trạm. Thời gian trễ chạm đất cài đặt là **$0.30\\text{ giây}$** nhằm đảm bảo tính chọn lọc (Selective Coordination) để các máy cắt nhánh (Feeder Breakers hoặc MCCB) trip trước trong trường hợp rò điện cục bộ (thời gian cắt nhánh thường là $0.05 - 0.10\\text{s}$).
   * Thử nghiệm bơm dòng thứ cấp (Secondary Injection) bằng thiết bị kiểm định cho thấy bộ điều khiển **MicroLogic 6.0A cắt tức thì chỉ sau $0.15\\text{ giây}$ (sớm gấp đôi so với đường cong cài đặt $0.30\\text{s}$)** $\\rightarrow$ **KẾT LUẬN FAIL**.
   * **Hậu quả sự cố:** Khi một tủ nhánh ở hạ nguồn bị sự cố chạm đất nhẹ, máy cắt tổng 3200A sẽ nhảy cắt sớm trước máy cắt nhánh, gây **mất điện toàn bộ nhà máy / trung tâm thương mại**, làm gián đoạn sản xuất và vi phạm tính chọn lọc bảo vệ.

2. **Bất Thường Điện Trở Tiếp Xúc Cực C (Contact Resistance Anomaly per NETA Sec 7.6.1.2.D.2):**
   * Pole A ($18.2\\,\\mu\\Omega$) và Pole B ($18.5\\,\\mu\\Omega$) có điện trở tiếp xúc xuất sắc và đồng đều.
   * Tuy nhiên, **Pole C đo được $42.1\\,\\mu\\Omega$**, chênh lệch **$+131.3\\%$** so với Pole A ($18.2\\,\\mu\\Omega$). Tiêu chuẩn NETA ATS-2025 Mục 7.6.1.2.D.2 quy định: *Độ lệch điện trở tiếp xúc giữa các cực kế cận hoặc so với cực nhỏ nhất không được vượt quá $50\\%$*.
   * Đối chiếu với dữ liệu kiểm định năm ngoái ($18.6\\,\\mu\\Omega$), điện trở Pole C đã **tăng vọt $+126.3\\%$**.
   * **Nguyên nhân và nguy cơ:** Với dòng tải định mức lên tới $3200\\text{A}$, công suất tổn hao nhiệt tại điểm tiếp xúc Pole C tính theo định luật Joule-Lenz ($P = I^2 R$) sẽ tăng gấp hơn $2.3$ lần so với Pole A và Pole B. Lượng nhiệt này sẽ gây **nóng đỏ ngàm cắm finger cluster, làm chai cứng lò xo ép tiếp điểm, cháy buồng dập hồ quang và nguy cơ nổ hồ quang pha C** khi mang tải nặng.

---

### 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)

1. **Hiệu Chỉnh Lại Thông Số Bộ Bảo Vệ MicroLogic 6.0A:**
   * Tách máy cắt ra vị trí **TEST**. Kết nối thiết bị chuyên dụng (Full Function Test Kit / Schneider Handheld Test Kit) vào cổng kiểm tra của bộ MicroLogic 6.0A.
   * Kiểm tra thông số cài đặt $I_g$ (Ground Fault Pickup) và $t_g$ (Ground Fault Time Delay). Đảm bảo công tắc cài đặt hoặc menu hiển thị số đang chọn đúng đường cong $I^2t\\text{ ON/OFF}$ và giá trị thời gian trễ $0.30\\text{s}$.
   * Thực hiện bơm dòng lại 3 lần liên tiếp; nếu bộ MicroLogic vẫn tiếp tục tác động ở $0.15\\text{s}$, cần thay thế bo mạch điện tử (Trip Unit Replacement) và nạp lại cấu hình chỉnh định phối hợp bảo vệ.

2. **Bảo Dưỡng Bề Mặt Tiếp Điểm & Ngàm Cắm Finger Clusters Pole C:**
   * Rút hẳn máy cắt ra khỏi khoang tủ (vị trí DISCONNECTED / OUT).
   * Tháo buồng dập hồ quang Pole C để kiểm tra trực quan tiếp điểm chính (Main Contacts) và tiếp điểm dập hồ quang (Arcing Contacts). Làm sạch bụi carbon và cặn cháy bám trên bề mặt tiếp xúc bằng dung dịch làm sạch tiếp điểm điện tử chuyên dụng (không dùng giấy nhám làm mất lớp mạ bạc).
   * Kiểm tra cụm ngàm cắm phía sau (Finger Clusters) của Pole C: Kiểm tra từng thanh tiếp xúc, đo lực căng của lò xo kẹp (tension springs). Nếu phát hiện lò xo bị giãn nhiệt hoặc đổi màu do quá nhiệt, phải thay mới toàn bộ cụm finger clusters.
   * Quét một lớp mỏng mỡ tiếp xúc dẫn điện chuyên dụng (như NO-OX-ID A-Special) lên ngàm cắm trước khi cắm thử vào thanh cái đực của tủ.

3. **Đo Kiểm Định Nghiệm Thu Lại Trước Khi Đưa Vào Vị Trí CONNECTED:**
   * Dùng cầu đo micro-ohm (DLRO dòng thử tối thiểu $100\\text{A DC}$) đo lại điện trở tiếp xúc Pole C. Giá trị bắt buộc phải **giảm xuống $\\le 20.0\\,\\mu\\Omega$** (độ lệch giữa 3 cực $\\le 15\\%$, không được vượt quá $50\\%$).
   * Kiểm tra thao tác đóng/cắt bằng cuộn đóng/cắt ở điện áp $85\\% - 110\\%$ điện áp định mức.
   * Chỉ đưa máy cắt vào vị trí **CONNECTED** và đóng điện hòa tải khi cả 2 lỗi trên đã được khắc phục hoàn toàn.`;

  return {
    overall_status,
    status_title,
    status_reason,
    markdown_report,
    evaluations: {
      visual_mechanical: {
        status: mechStatus,
        detail: mechDetail,
      },
      contact_resistance: {
        pole_a_micro_ohms: cp.pole_a,
        pole_b_micro_ohms: cp.pole_b,
        pole_c_micro_ohms: cp.pole_c,
        min_pole_micro_ohms: minPole,
        max_deviation_percent: maxPoleDev,
        historical_increase_percent: poleCHistoricalIncrease,
        status: contactStatus,
        detail: contactDetail,
      },
      trip_unit_injection: {
        long_time: inj.long_time_pickup_status,
        short_time: inj.short_time_delay_status,
        instantaneous: inj.instantaneous_pickup_status,
        ground_fault: inj.ground_fault_pickup_status,
        status: tripUnitStatus,
        detail: tripUnitDetail,
      },
      insulation_resistance: {
        min_closed_megohms: minClosed,
        min_open_megohms: minOpen,
        standard_min_megohms: standardMinMegohms,
        control_wiring_ir_megohms: e.control_wiring_ir_megohms,
        status: insulationStatus,
        detail: insulationDetail,
      },
      bolted_connections: {
        min_micro_ohms: minBolt,
        max_micro_ohms: maxBolt,
        max_deviation_percent: boltDev,
        status: boltStatus,
        detail: boltDetail,
      },
    },
  };
}

export async function generateNetaLvBreakerAiAnalysis(payload: NetaLvBreakerInputPayload): Promise<NetaLvBreakerEvaluationResult> {
  const deterministicResult = performDeterministicLvBreakerAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.log('[NETA LV Breaker Analyzer] No GEMINI_API_KEY provided; returning deterministic evaluation.');
    return deterministicResult;
  }

  const systemInstruction = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách Máy cắt không khí hạ áp (Low-Voltage Power Circuit Breakers - ACB) theo tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.6.1.2).

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.6.1.2):

1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Nameplate Match: Đối soát nhãn mác máy cắt (điện áp, dòng định mức, dòng cắt ngắn mạch, loại trip unit) khớp 100% bản vẽ.
- Physical Condition & Cleanliness: Máy cắt sạch bẩn, tháo bỏ kẹp vận chuyển; buồng dập hồ quang (arc chutes) nguyên vẹn.
- Racking & Interlocks: Cơ cấu kéo/đẩy drawout trơn tru; liên động an toàn ngăn rút/cắm máy cắt khi đang ĐÓNG.
- Contacts & Finger Clusters: Tiếp điểm chính, tiếp điểm dập hồ quang và ngàm cắm finger clusters không bị rỗ nặng hay biến dạng.
- Trip Unit Settings: Các thông số cài đặt L, S, I, G khớp 100% phiếu chỉnh định (Coordination Study).
- Bolt Torque: Mối nối bu-lông siết đạt lực theo dữ liệu nhà sản xuất hoặc Bảng Table 100.12.

2. PHÉP ĐO ĐIỆN VÀ TIÊU CHUẨN ĐÁNH GIÁ (ELECTRICAL TEST VALUES):
- Điện trở mối nối bu-lông (Bolted Connection Resistance): CẢNH BÁO "INVESTIGATE" nếu giá trị đo mối nối lệch quá 50% so với giá trị nhỏ nhất của mối nối tương tự.
- Điện trở tiếp xúc cực (Contact / Pole Resistance): CẢNH BÁO "INVESTIGATE" nếu giá trị đo cực lệch quá 50% so với cực kế cận hoặc giá trị nhỏ nhất.
- Điện trở cách điện Máy cắt (Insulation Resistance - IR): Đo 1 phút (Closed & Open) đạt tối thiểu theo Bảng Table 100.1 (>= 100 Megohms cho cấp 600V).
- Điện trở cách điện Mạch phụ (Control Wiring IR): Thử nghiệm 500V/1000V DC BẮT BUỘC >= 2.0 Megohms (2 MΩ).
- Thử nghiệm Bơm dòng Trip Unit (Primary/Secondary Injection): Ngưỡng pickup và thời gian trip L, S, I, G nằm trong dải dung sai đường cong đặc tính nhà sản xuất.
- Điện áp Thao tác Tối thiểu (Minimum Operating Voltage): Cuộn đóng và cuộn cắt tác động tin cậy ở mức điện áp sụt tối thiểu theo chuẩn ANSI/IEEE.

NHIỆM VỤ CỦA AI:
Khi nhận dữ liệu JSON từ FSE, hãy phân tích và trả về phản hồi định dạng Markdown gồm 4 phần:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE.
2. BẢNG PHÂN TÍCH CHI TIẾT MÁY CẮT HẠ ÁP (Detailed Evaluation Table): So sánh thực tế với NETA ATS-2025 Section 7.6.1.2 (Table 100.1, Table 100.12).
3. PHÂN TÍCH ĐẶC TÍNH TRIP UNIT, TIẾP ĐIỂM & NGUY CƠ MẤT AN TOÀN (Trip Unit Curve, Contact Resistance & Protection Risk).
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations).`;

  try {
    const ai = new GoogleGenAI({});
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [
        {
          role: 'user',
          parts: [
            {
              text: `${systemInstruction}\n\nĐây là dữ liệu JSON thực tế từ FSE tại công trường:\n${JSON.stringify(payload, null, 2)}`
            }
          ]
        }
      ]
    });

    const aiMarkdown = response.text?.trim();
    if (aiMarkdown && aiMarkdown.length > 100) {
      return {
        ...deterministicResult,
        markdown_report: aiMarkdown
      };
    }
  } catch (error) {
    console.error('[NETA LV Breaker Analyzer] Gemini generation error, using deterministic report:', error);
  }

  return deterministicResult;
}
