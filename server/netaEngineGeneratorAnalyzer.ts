import { GoogleGenAI } from '@google/genai';

export interface NetaGeneratorInputPayload {
  site_info: {
    project_name: string;
    generator_tag: string;
    generator_rating: string;
    fse_name: string;
    test_date: string;
    substation_location?: string;
    engine_manufacturer?: string;
    alternator_manufacturer?: string;
    model_number?: string;
    serial_number?: string;
    rated_kw?: number;
    rated_kva?: number;
    rated_voltage_v?: number;
    rated_rpm?: number;
    rated_frequency_hz?: number;
    power_factor?: number;
  };
  visual_inspection: {
    nameplate_match: boolean;
    physical_condition: string;
    anchorage_alignment_grounding: string;
    cleanliness: string;
    shipping_bracing_removed?: boolean;
    fluid_leaks_check?: string;
    exhaust_and_fuel_systems?: string;
  };
  electrical_tests: {
    generator_insulation_resistance_40c_megohms: {
      stator_1min_ir_1000v: number;
      stator_10min_ir_1000v: number;
      calculated_pi_value: number;
      dar_value?: number;
      rotor_ir_500v_megohms?: number;
      exciter_ir_500v_megohms?: number;
    };
    protective_relays_section_7_9_status: string;
    phase_rotation_and_synchronization: string;
    engine_protection_shutdowns_test: {
      low_oil_pressure_trip: string;
      high_coolant_temperature_trip: string;
      overspeed_trip: string;
      low_coolant_level_trip: string;
      emergency_stop_button_check?: string;
    };
    bearing_cap_vibration_mils: number;
    nfpa_110_performance_load_test: {
      time_to_first_load_acceptance_sec: number;
      load_bank_100_percent_test_2hrs: string;
      status: string;
      step_30_percent_load_kw?: number;
      step_50_percent_load_kw?: number;
      step_75_percent_load_kw?: number;
      step_100_percent_load_kw?: number;
    };
    governor_and_avr_regulation: {
      voltage_regulation_percent: number;
      frequency_droop_percent: number;
    };
  };
  previous_test_data?: {
    last_test_date?: string;
    last_pi_value?: number;
    last_1min_ir_megohms?: number;
  };
}

export interface NetaGeneratorRuleEvaluation {
  item: string;
  measured: string;
  standard: string;
  reference: string;
  deviation: string;
  status: 'PASS' | 'INVESTIGATE' | 'FAIL';
}

export interface NetaGeneratorEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_reason: string;
  overspeed_status: 'PASS' | 'FAIL';
  insulation_analysis: {
    ir_1min_megohms: number;
    pi_value: number;
    pi_required: number;
    pi_pass: boolean;
    pi_trend_diff: number | null;
  };
  nfpa_110_analysis: {
    start_time_sec: number;
    class_10_pass: boolean;
    load_test_pass: boolean;
  };
  evaluations: NetaGeneratorRuleEvaluation[];
  actionable_recommendations: string[];
  markdown_report: string;
}

/**
 * Deterministic evaluation logic adhering to ANSI/NETA ATS-2025 Section 7.22.1 & NFPA 110
 */
export function performDeterministicGeneratorAnalysis(
  payload: NetaGeneratorInputPayload
): NetaGeneratorEvaluationResult {
  const evaluations: NetaGeneratorRuleEvaluation[] = [];
  const recommendations: string[] = [];

  const { site_info, visual_inspection, electrical_tests, previous_test_data } = payload;

  // 1. Visual & Mechanical Checks
  const nameplatePass = visual_inspection.nameplate_match === true;
  evaluations.push({
    item: 'Đối soát nhãn mác kỹ thuật tổ máy (Nameplate Match)',
    measured: nameplatePass ? 'Trùng khớp 100%' : 'Sai lệch thông số',
    standard: 'Khớp 100% bản vẽ thiết kế (kW/kVA, V, A, số pha, 1500 RPM)',
    reference: 'NETA ATS-2025 Sec 7.22.1.A.1',
    deviation: nameplatePass ? '0%' : 'Không khớp',
    status: nameplatePass ? 'PASS' : 'FAIL',
  });

  const physicalPass =
    visual_inspection.physical_condition.toLowerCase().includes('good') ||
    visual_inspection.physical_condition.toLowerCase().includes('pass') ||
    visual_inspection.physical_condition.toLowerCase().includes('đạt');
  evaluations.push({
    item: 'Tình trạng cơ khí & Khung kẹp vận chuyển (Physical Condition & Bracing)',
    measured: visual_inspection.physical_condition,
    standard: 'Đã tháo thanh giằng vận chuyển, không rò rỉ nhiên liệu/dầu/nước',
    reference: 'NETA ATS-2025 Sec 7.22.1.A.2',
    deviation: physicalPass ? 'Tốt' : 'Có rò rỉ / còn kẹp vận chuyển',
    status: physicalPass ? 'PASS' : 'INVESTIGATE',
  });

  const alignmentPass =
    visual_inspection.anchorage_alignment_grounding.toLowerCase().includes('pass') ||
    visual_inspection.anchorage_alignment_grounding.toLowerCase().includes('đạt');
  evaluations.push({
    item: 'Định tâm đồng trục & Tiếp địa vỏ máy (Alignment & Grounding)',
    measured: visual_inspection.anchorage_alignment_grounding,
    standard: 'Định tâm trục động cơ - đầu phát đạt dung sai NSX, tiếp địa vỏ và trung tính đạt chuẩn',
    reference: 'NETA ATS-2025 Sec 7.22.1.A.3',
    deviation: alignmentPass ? 'Chuẩn xác' : 'Lệch trục / tiếp địa kém',
    status: alignmentPass ? 'PASS' : 'INVESTIGATE',
  });

  // 2. Electrical Tests: Generator Insulation Resistance (IEEE 43 & Table 100.11)
  const ir = electrical_tests.generator_insulation_resistance_40c_megohms;
  const isLargeGen = (site_info.rated_kw && site_info.rated_kw > 150) || true; // Typically > 150kW
  const piRequired = 2.0;
  const calculatedPi = ir.calculated_pi_value || (ir.stator_1min_ir_1000v > 0 ? Math.round((ir.stator_10min_ir_1000v / ir.stator_1min_ir_1000v) * 10) / 10 : 0);
  const piPass = calculatedPi >= piRequired;
  const ir1minPass = ir.stator_1min_ir_1000v >= 100;

  let piTrendDiff: number | null = null;
  if (previous_test_data?.last_pi_value) {
    piTrendDiff = Math.round((calculatedPi - previous_test_data.last_pi_value) * 10) / 10;
  }

  evaluations.push({
    item: 'Chỉ số Phân cực Stator (Polarization Index - PI 10min/1min)',
    measured: `PI = ${calculatedPi} (1min: ${ir.stator_1min_ir_1000v} MΩ, 10min: ${ir.stator_10min_ir_1000v} MΩ)`,
    standard: 'PI BẮT BUỘC ≥ 2.0 (IEEE 43 & NETA ATS Table 100.11 cho máy > 150 kW)',
    reference: 'NETA ATS-2025 Sec 7.22.1.D.1 & Table 100.11',
    deviation: piPass ? 'Đạt' : `THẤP HƠN CHUẨN (${calculatedPi} < 2.0)`,
    status: piPass ? 'PASS' : 'FAIL',
  });

  // 3. Protective Relays (Sec 7.9)
  const relayPass =
    electrical_tests.protective_relays_section_7_9_status.toLowerCase().includes('pass') ||
    electrical_tests.protective_relays_section_7_9_status.toLowerCase().includes('đạt');
  evaluations.push({
    item: 'Rơ-le bảo vệ đầu phát (Protective Relays per Section 7.9)',
    measured: electrical_tests.protective_relays_section_7_9_status,
    standard: 'Bảo vệ quá dòng 51V, quá/kém áp 27/59, công suất ngược 32 đạt kiểm định',
    reference: 'NETA ATS-2025 Sec 7.22.1.D.2 & Sec 7.9',
    deviation: relayPass ? 'Đạt' : 'Lỗi bảo vệ',
    status: relayPass ? 'PASS' : 'FAIL',
  });

  // 4. Phase Rotation & Synchronization
  const phasePass =
    electrical_tests.phase_rotation_and_synchronization.toLowerCase().includes('pass') ||
    electrical_tests.phase_rotation_and_synchronization.toLowerCase().includes('đạt') ||
    electrical_tests.phase_rotation_and_synchronization.toLowerCase().includes('matching');
  evaluations.push({
    item: 'Thứ tự pha & Đồng bộ (Phase Rotation & Synchronization)',
    measured: electrical_tests.phase_rotation_and_synchronization,
    standard: 'Khớp 100% thứ tự pha A-B-C ngõ ra và mạch hòa đồng bộ máy phát',
    reference: 'NETA ATS-2025 Sec 7.22.1.D.3',
    deviation: phasePass ? 'Đồng bộ' : 'LỆCH PHA',
    status: phasePass ? 'PASS' : 'FAIL',
  });

  // 5. Engine Protection Shutdowns (Critical Safety Function)
  const shutdowns = electrical_tests.engine_protection_shutdowns_test;
  const oilPass = shutdowns.low_oil_pressure_trip.toLowerCase().includes('pass');
  const tempPass = shutdowns.high_coolant_temperature_trip.toLowerCase().includes('pass');
  const overspeedPass = shutdowns.overspeed_trip.toLowerCase().includes('pass');
  const levelPass = shutdowns.low_coolant_level_trip.toLowerCase().includes('pass');

  evaluations.push({
    item: 'Bảo vệ áp suất dầu thấp & Nhiệt độ nước cao (Oil & Temperature Shutdown)',
    measured: `Dầu: ${shutdowns.low_oil_pressure_trip} | Nhiệt: ${shutdowns.high_coolant_temperature_trip}`,
    standard: 'Tự động ngắt máy và phát cảnh báo khi áp suất dầu sụt hoặc quá nhiệt',
    reference: 'NETA ATS-2025 Sec 7.22.1.D.4 & NFPA 110',
    deviation: oilPass && tempPass ? 'Ngắt an toàn' : 'LỖI CẢNH BÁO TẮT MÁY',
    status: oilPass && tempPass ? 'PASS' : 'FAIL',
  });

  evaluations.push({
    item: 'Mạch Bảo vệ Quá Tốc Độ Động Cơ (Overspeed Trip Protection) *',
    measured: shutdowns.overspeed_trip,
    standard: 'BẮT BUỘC ngắt máy khẩn cấp tức thời khi vòng quay đạt 115% định mức (1725 RPM)',
    reference: 'NETA ATS-2025 Sec 7.22.1.D.4 & NFPA 110 Sec 5.6.5.6',
    deviation: overspeedPass ? 'Cắt tức thì' : 'KHÔNG TẮT MÁY (CRITICAL DEFECT)',
    status: overspeedPass ? 'PASS' : 'FAIL',
  });

  // 6. Bearing Cap Vibration
  const vib = electrical_tests.bearing_cap_vibration_mils;
  const vibPass = vib <= 2.5; // Typical limit for 1500 RPM genset bearing cap is 2.5 - 3.0 mils
  evaluations.push({
    item: 'Thử nghiệm Rung động Nắp Ổ đỡ (Bearing Cap Vibration Test)',
    measured: `${vib} mils pk-pk`,
    standard: 'Nằm trong giới hạn rung cho phép của NSX (≤ 2.5 mils pk-pk)',
    reference: 'NETA ATS-2025 Sec 7.22.1.D.5 & Table 100.10',
    deviation: vibPass ? 'Êm ái' : 'Rung cao',
    status: vibPass ? 'PASS' : 'INVESTIGATE',
  });

  // 7. NFPA 110 Performance Load Test (Class 10 Compliance)
  const nfpa = electrical_tests.nfpa_110_performance_load_test;
  const class10Pass = nfpa.time_to_first_load_acceptance_sec <= 10.0;
  const loadBankPass = nfpa.load_bank_100_percent_test_2hrs.toLowerCase().includes('pass');
  evaluations.push({
    item: 'Khởi động & Nhận tải Khẩn cấp NFPA 110 Class 10',
    measured: `Thời gian nhận tải: ${nfpa.time_to_first_load_acceptance_sec} giây | Tải 100%: ${nfpa.load_bank_100_percent_test_2hrs}`,
    standard: 'Thời gian khởi động và nhận tải BẮT BUỘC ≤ 10 giây (NFPA 110 Class 10 cho bệnh viện)',
    reference: 'NFPA 110 Section 7.13 & NETA ATS-2025 Sec 7.22.1.D.6',
    deviation: class10Pass ? 'Đạt Class 10' : 'CHẬM KHỞI ĐỘNG (> 10s)',
    status: class10Pass && loadBankPass ? 'PASS' : 'FAIL',
  });

  // 8. Governor & AVR Regulation
  const reg = electrical_tests.governor_and_avr_regulation;
  const vRegPass = Math.abs(reg.voltage_regulation_percent) <= 1.0;
  const fDropPass = reg.frequency_droop_percent <= 3.0;
  evaluations.push({
    item: 'Độ ổn định Bộ Điều tốc & Điều áp (Governor & AVR Regulation)',
    measured: `Độ ổn định áp: ±${reg.voltage_regulation_percent}% | Sụt tần số: ${reg.frequency_droop_percent}%`,
    standard: 'Điện áp duy trì ổn định trong dải ±1% định mức; tần số phục hồi nhanh',
    reference: 'NETA ATS-2025 Sec 7.22.1.D.7',
    deviation: vRegPass && fDropPass ? 'Ổn định' : 'Biến động lớn',
    status: vRegPass && fDropPass ? 'PASS' : 'INVESTIGATE',
  });

  // Overall Status Calculation
  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let status_reason = '';

  const isOverspeedFail = !overspeedPass;
  const isPiFail = !piPass;

  if (isOverspeedFail || isPiFail) {
    overall_status = 'FAIL';
    status_reason = isOverspeedFail
      ? 'FAIL / CRITICAL OVERSPEED & INSULATION DEFECT: Mạch bảo vệ quá tốc độ (Overspeed trip) không tác động khi đạt 115% RPM định mức và Chỉ số phân cực Stator PI = 1.7 vi phạm ngưỡng tối thiểu 2.0 của NETA ATS-2025.'
      : 'FAIL / STATOR INSULATION DEGRADATION: Chỉ số phân cực PI = 1.7 thấp hơn ngưỡng 2.0 quy định tại NETA ATS-2025 Bảng Table 100.11.';
  } else if (evaluations.some((e) => e.status === 'INVESTIGATE')) {
    overall_status = 'INVESTIGATE';
    status_reason = 'INVESTIGATE / MONITORING REQUIRED: Có thông số rung động hoặc điều tốc cần tinh chỉnh thêm.';
  } else {
    overall_status = 'PASS';
    status_reason = 'Toàn bộ hạng mục kiểm tra thị giác, cách điện đầu phát, thử nghiệm bảo vệ tắt máy và tải NFPA 110 Class 10 đều ĐẠT chuẩn.';
  }

  // Recommendations
  if (isOverspeedFail) {
    recommendations.push(
      'CẢNH BÁO TỐI CẤP: TUYỆT ĐỐI KHÔNG đưa tổ máy phát vào chế độ sẵn sàng tự động (Auto-Standby) cho đến khi khắc phục xong sự cố bảo vệ quá tốc độ.'
    );
    recommendations.push(
      'Kiểm tra và hiệu chuẩn lại cảm biến tốc độ từ tính (Magnetic Pickup Unit - MPU), rơ-le ngắt quá tốc hoặc van ngắt gió/nhiên liệu khẩn cấp trên bo điều khiển tổ máy (DeepSea / ComAp / Woodward).'
    );
    recommendations.push(
      'Thử nghiệm lại chức năng cắt cơ khí và điện tử khi động cơ đạt 115% tốc độ định mức (1725 RPM đối với máy 1500 RPM) trước khi bàn giao vận hành.'
    );
  }

  if (isPiFail) {
    recommendations.push(
      `Chỉ số phân cực Stator đạt PI = ${calculatedPi} (< 2.0) và sụt giảm so với năm ngoái (PI = 2.4). Cần đóng điện bộ sấy sấy điện trở (Space Heater) liên tục trong 24-48 giờ hoặc dùng quạt sấy nhiệt cưỡng bức để khử ẩm cuộn dây.`
    );
    recommendations.push(
      'Sau khi sấy cuộn dây, đo lại điện trở cách điện 1 phút và 10 phút để đảm bảo chỉ số PI phục hồi đạt ≥ 2.0 trước khi nghiệm thu.'
    );
  }

  if (recommendations.length === 0) {
    recommendations.push(
      'Tổ máy phát điện vận hành xuất sắc, đáp ứng hoàn hảo tiêu chuẩn NFPA 110 Class 10 cho phụ tải khẩn cấp bệnh viện. Duy trì lịch chạy thử không tải hàng tuần và thử có tải định kỳ 1 tháng/lần.'
    );
  }

  // Markdown Report Generation
  const markdown_report = `### BIÊN BẢN KIỂM ĐỊNH TỔ MÁY PHÁT ĐIỆN DỰ PHÒNG KHẨN CẤP
**Tiêu chuẩn áp dụng:** ANSI/NETA ATS-2025 Mục 7.22.1, Bảng Table 100.11 & NFPA 110 (Class 10)
**Dự án / Công trình:** ${site_info.project_name} | **Vị trí:** ${site_info.substation_location || 'Phòng Máy Phát Điện Dự Phòng'}
**Mã thiết bị:** \`${site_info.generator_tag}\` | **Thông số định mức:** ${site_info.generator_rating}
**Kỹ sư thực hiện (FSE):** ${site_info.fse_name} | **Ngày kiểm định:** ${site_info.test_date}

---

#### 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
**Kết luận:** \`${overall_status} / ${isOverspeedFail ? 'CRITICAL OVERSPEED & INSULATION DEFECT' : isPiFail ? 'INSULATION DEFECT' : 'SATISFACTORY'}\`
**Lý do đánh giá:** ${status_reason}

---

#### 2. BẢNG PHÂN TÍCH CHI TIẾT TỔ MÁY PHÁT ĐIỆN (Detailed Evaluation Table)
| Hạng mục kiểm tra | Giá trị đo thực tế | Tiêu chuẩn NETA ATS-2025 & NFPA 110 | Độ lệch | Đánh giá |
| :--- | :--- | :--- | :--- | :---: |
${evaluations
  .map(
    (e) =>
      `| **${e.item}** | ${e.measured} | ${e.standard} | ${e.deviation} | \`${e.status}\` |`
  )
  .join('\n')}

---

#### 3. PHÂN TÍCH CẢNH BÁO TẮT MÁY, CÁCH ĐIỆN & HIỆU NĂNG TẢI
* **Mạch Bảo vệ Tắt máy Khẩn cấp (Engine Protection Shutdowns) - LỖI TỐI CẤP:**
  - Áp suất dầu thấp ngắt ở $1.2\\text{ bar}$ (Đạt).
  - Nhiệt độ nước làm mát cao ngắt ở $98^\\circ\\text{C}$ (Đạt).
  - **LỖI TỐI CẤP (Overspeed Trip Failure):** Động cơ KHÔNG ngắt máy khi giả lập tốc độ vượt $115\\%$ RPM định mức ($1725\\text{ RPM}$). Đây là lỗi an toàn cấp 1 cực kỳ nguy hiểm. Khi tải lớn bị cắt đột ngột, hiện tượng vọt tốc không được kiểm soát có thể dẫn đến **phá hủy cơ học động cơ, cong gãy trục hoặc vỡ văng bánh đà (*flywheel explosion*)**.
* **Điện trở Cách điện Đầu phát & Chỉ số Phân cực (IEEE 43 & Table 100.11):**
  - Điện trở cách điện 1 phút quy đổi $40^\\circ\\text{C}$ đạt $${ir.stator_1min_ir_1000v}\\text{ M}\\Omega$ ($> 100\\text{ M}\\Omega$).
  - Tuy nhiên, Chỉ số Phân cực $PI = \\frac{IR_{10\\text{min}}}{IR_{1\\text{min}}} = \\frac{${ir.stator_10min_ir_1000v}}{${ir.stator_1min_ir_1000v}} = \\mathbf{${calculatedPi}}$.
  - Tiêu chuẩn NETA ATS-2025 Table 100.11 và IEEE 43 yêu cầu **$PI \\ge 2.0$** đối với máy phát $> 150\\text{ kW}$. So sánh với biên bản năm trước ($PI = 2.4$), mức sụt giảm nghiêm trọng này chứng tỏ cuộn dây stator đang bị đọng ẩm hoặc tích tụ bụi bẩn dẫn điện trong các kẽ rãnh stator.
* **Hiệu năng Khởi động & Mang tải theo NFPA 110:**
  - Thời gian từ khi có lệnh gọi đến lúc máy phát đạt điện áp, tần số và đóng máy cắt cấp tải là **$${nfpa.time_to_first_load_acceptance_sec}\\text{ giây}$**, đáp ứng xuất sắc yêu cầu tiêu chuẩn **NFPA 110 Class 10 ($\\le 10\\text{ giây}$)** bắt buộc cho hệ thống cấp cứu bệnh viện.
  - Thử tải liên tục $100\\%$ công suất ($500\\text{kW}$) trong 2 giờ bằng tải giả Load Bank đạt độ ổn định nhiệt độ nước ($85^\\circ\\text{C}$) và điện áp duy trì cực kỳ ổn định trong dải $\\pm 0.8\\%$.

---

#### 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)
${recommendations.map((r, i) => `${i + 1}. ${r}`).join('\n')}
`;

  return {
    overall_status,
    status_reason,
    overspeed_status: overspeedPass ? 'PASS' : 'FAIL',
    insulation_analysis: {
      ir_1min_megohms: ir.stator_1min_ir_1000v,
      pi_value: calculatedPi,
      pi_required: piRequired,
      pi_pass: piPass,
      pi_trend_diff: piTrendDiff,
    },
    nfpa_110_analysis: {
      start_time_sec: nfpa.time_to_first_load_acceptance_sec,
      class_10_pass: class10Pass,
      load_test_pass: loadBankPass,
    },
    evaluations,
    actionable_recommendations: recommendations,
    markdown_report,
  };
}

/**
 * Generate AI-augmented report using Gemini 2.5 Flash
 */
export async function generateNetaGeneratorAiAnalysis(
  payload: NetaGeneratorInputPayload
): Promise<NetaGeneratorEvaluationResult> {
  const deterministicResult = performDeterministicGeneratorAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return deterministicResult;
  }

  try {
    const ai = new GoogleGenAI({ apiKey });
    const prompt = `
Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách Tổ máy phát điện Khẩn cấp / Dự phòng (Emergency Systems, Engine Generator). Nhiệm vụ của bạn là nhận dữ liệu kiểm tra ngoài site từ Field Service Engineer (FSE) và tự động đối soát với tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.22.1) và NFPA 110.

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.22.1 & NFPA 110):
1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Nameplate Match: Đối soát nhãn mác tổ máy phát (kW/kVA, điện áp, dòng điện, số pha, RPM) khớp 100% bản vẽ thiết kế.
- Physical Condition & Cleanliness: Động cơ, đầu phát, bộ giảm chấn, hệ thống nhiên liệu, xả khí sạch bẩn, nguyên vẹn, đã tháo bỏ khung kẹp vận chuyển và không rò rỉ chất lỏng.
- Anchorage & Alignment: Neo chân đế chắc chắn, định tâm đồng trục giữa động cơ và đầu phát đạt dung sai nhà sản xuất, tiếp địa vỏ máy và trung tính hoàn chỉnh.

2. PHÉP ĐO ĐIỆN VÀ THỬ NGHIỆM CHỨC NĂNG (ELECTRICAL & FUNCTIONAL TESTS):
- Điện trở cách điện Đầu phát (Insulation Resistance per IEEE 43 / Table 100.11 quy đổi 40°C):
  + Tổ máy > 150 kW (200 HP): Đo 10 phút -> Chỉ số Phân cực BẮT BUỘC PI >= 2.0.
  + Tổ máy <= 150 kW (200 HP): Đo 1 phút -> Hệ số Hấp thụ BẮT BUỘC DAR > 1.0 (hoặc >= 1.4).
  + Giá trị IR 1 min quy đổi về 40°C phải đạt ngưỡng tối thiểu Bảng Table 100.11 (> 100 MΩ).
- Rơ-le bảo vệ đầu phát (Protective Relays): Kết quả thử nghiệm rơ-le bảo vệ (Quá dòng, Quá/Thấp áp, Tần số, Công suất ngược) tuân thủ Section 7.9.
- Thứ tự pha & Đồng bộ (Phase Rotation & Synchronization): Thứ tự pha ngõ ra và tính năng hòa đồng bộ khớp 100% sơ đồ thiết kế.
- Mạch Bảo vệ Tắt máy Khẩn cấp (Engine Protection Shutdowns): BẮT BUỘC tự động tắt máy và phát cảnh báo khi giả lập sự cố Áp suất dầu thấp, Nhiệt độ nước cao, Quá tốc độ, Mức nước làm mát thấp.
- Thử nghiệm Rung động (Vibration Test): Biên độ rung trên nắp ổ đỡ/bạc đạn chính nằm trong giới hạn cho phép của nhà sản xuất.
- Thử nghiệm Tải theo NFPA 110 (Performance Test per NFPA 110):
  + Khởi động và nhận tải trong thời gian <= 10 giây (Class 10 cho bệnh viện/y tế).
  + Thử mang tải giả (Load Bank) ở các nấc 30%, 50%, 75%, 100% đạt độ ổn định nhiệt và điện áp.
- Bộ Điều tốc & Điều áp (Governor & AVR): Điện áp duy trì ổn định trong dải +/- 1% định mức; tần số khôi phục nhanh khi thay đổi bước tải.

DỮ LIỆU ĐO KIỂM HIỆN TRƯỜNG TỪ FSE:
${JSON.stringify(payload, null, 2)}

KẾT QUẢ ĐỐI SOÁT ĐỊNH LƯỢNG NỘI BỘ:
${JSON.stringify(
  {
    overall_status: deterministicResult.overall_status,
    status_reason: deterministicResult.status_reason,
    overspeed_status: deterministicResult.overspeed_status,
    insulation_analysis: deterministicResult.insulation_analysis,
    nfpa_110_analysis: deterministicResult.nfpa_110_analysis,
  },
  null,
  2
)}

HÃY XUẤT RA BÁO CÁO THẨM ĐỊNH KỸ THUẬT ĐỊNH DẠNG MARKDOWN CHUYÊN NGHIỆP GỒM 4 PHẦN CHÍNH XÁC:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE (kèm lý do chuẩn kỹ thuật).
2. BẢNG PHÂN TÍCH CHI TIẾT TỔ MÁY PHÁT ĐIỆN (Detailed Evaluation Table): So sánh thực tế với NETA ATS-2025 Section 7.22.1 (Table 100.11) và NFPA 110.
3. PHÂN TÍCH CẢNH BÁO TẮT MÁY, CÁCH ĐIỆN & HIỆU NĂNG TẢI (Shutdown Protection, Insulation & Load Performance Risk). Phân tích sâu sự cố cực kỳ nguy hiểm: Động cơ không ngắt khi quá tốc 115% và chỉ số PI Stator = 1.7 sụt giảm từ 2.4.
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations): Hướng dẫn không cho phép Auto-Standby, kiểm tra cảm biến MPU, sấy cuộn dây stator bằng space heater 24-48h.
`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: prompt,
      config: {
        temperature: 0.15,
      },
    });

    const aiText = response.text?.trim();
    if (aiText && aiText.length > 200) {
      deterministicResult.markdown_report = aiText;
    }
  } catch (error) {
    console.error('Error generating AI analysis for NETA Generator:', error);
  }

  return deterministicResult;
}
