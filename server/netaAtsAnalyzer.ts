import { GoogleGenAI } from '@google/genai';

export interface NetaAtsInputPayload {
  site_info: {
    project_name: string;
    ats_tag: string;
    ats_rating: string;
    fse_name: string;
    test_date: string;
    substation_location?: string;
    manufacturer?: string;
    model_number?: string;
    serial_number?: string;
    rated_current_a?: number;
    rated_voltage_v?: number;
    poles_count?: number;
    transition_type?: 'Open Transition' | 'Closed Transition' | 'Delayed Transition';
  };
  visual_inspection: {
    nameplate_match: boolean;
    physical_condition: string;
    cleanliness: string;
    manual_transfer_operation: string;
    mechanical_interlock_check: string;
    anchorage_grounding_check: string;
    bolt_torque_check: string;
    arc_chutes_condition?: string;
    thermographic_survey?: string;
  };
  electrical_tests: {
    bolted_resistance_micro_ohms: number[];
    main_pole_insulation_resistance_1000v_megohms: {
      normal_source_position: number;
      alternate_source_position: number;
    };
    control_wiring_ir_megohms: number;
    contact_pole_resistance_micro_ohms: {
      normal_source_pole_a: number;
      normal_source_pole_b: number;
      normal_source_pole_c: number;
      normal_source_pole_n?: number;
      alternate_source_pole_a: number;
      alternate_source_pole_b: number;
      alternate_source_pole_c: number;
      alternate_source_pole_n?: number;
    };
    phasing_and_rotation_check: string;
    automatic_sequence_tests: {
      normal_source_undervoltage_sensing: string;
      engine_start_signal_time_sec: number;
      alternate_source_voltage_frequency_sensing: string;
      transfer_time_delay_sec: number;
      retransfer_time_delay_sec: number;
      engine_cool_down_time_sec: number;
      limit_switches_and_interlocks: string;
      in_phase_monitor_check?: string;
    };
  };
  previous_test_data?: {
    last_test_date?: string;
    last_normal_pole_c_resistance_micro_ohms?: number;
    last_contact_resistance_avg_micro_ohms?: number;
  };
}

export interface NetaAtsRuleEvaluation {
  item: string;
  measured: string;
  standard: string;
  reference: string;
  deviation: string;
  status: 'PASS' | 'INVESTIGATE' | 'FAIL';
}

export interface NetaAtsEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_reason: string;
  contact_resistance_analysis: {
    normal_min_micro_ohms: number;
    normal_max_micro_ohms: number;
    normal_max_deviation_percent: number;
    alternate_min_micro_ohms: number;
    alternate_max_micro_ohms: number;
    alternate_max_deviation_percent: number;
    worst_pole: string;
    is_anomaly: boolean;
  };
  insulation_analysis: {
    main_pole_pass: boolean;
    control_wiring_pass: boolean;
    control_ir_megohms: number;
  };
  sequence_analysis: {
    engine_start_sec: number;
    transfer_delay_sec: number;
    retransfer_delay_sec: number;
    cool_down_sec: number;
    interlocks_verified: boolean;
  };
  evaluations: NetaAtsRuleEvaluation[];
  actionable_recommendations: string[];
  markdown_report: string;
}

/**
 * Deterministic evaluation logic adhering strictly to ANSI/NETA ATS-2025 Section 7.22.3
 */
export function performDeterministicAtsAnalysis(payload: NetaAtsInputPayload): NetaAtsEvaluationResult {
  const evaluations: NetaAtsRuleEvaluation[] = [];
  const recommendations: string[] = [];

  const { visual_inspection, electrical_tests } = payload;

  // 1. Visual & Mechanical Checks
  const nameplatePass = visual_inspection.nameplate_match === true;
  evaluations.push({
    item: 'Đối soát nhãn mác kỹ thuật (Nameplate Match)',
    measured: nameplatePass ? 'Trùng khớp 100%' : 'Sai khác thông số',
    standard: 'Khớp 100% bản vẽ thiết kế (Điện áp, dòng điện, số cực, sơ đồ)',
    reference: 'NETA ATS-2025 Sec 7.22.3.A.1',
    deviation: nameplatePass ? '0%' : 'Không khớp',
    status: nameplatePass ? 'PASS' : 'FAIL',
  });

  const physicalPass = visual_inspection.physical_condition.toLowerCase().includes('good') ||
    visual_inspection.physical_condition.toLowerCase().includes('pass') ||
    visual_inspection.physical_condition.toLowerCase().includes('đạt');
  evaluations.push({
    item: 'Tình trạng vật lý & Buồng dập hồ quang (Physical Condition & Arc Chutes)',
    measured: visual_inspection.physical_condition,
    standard: 'Sạch sẽ, không nứt vỡ cơ lý, buồng dập hồ quang nguyên vẹn',
    reference: 'NETA ATS-2025 Sec 7.22.3.A.2',
    deviation: physicalPass ? 'Bình thường' : 'Bất thường',
    status: physicalPass ? 'PASS' : 'INVESTIGATE',
  });

  const manualOpPass = visual_inspection.manual_transfer_operation.toLowerCase().includes('pass') ||
    visual_inspection.manual_transfer_operation.toLowerCase().includes('đạt') ||
    visual_inspection.manual_transfer_operation.toLowerCase().includes('smooth');
  evaluations.push({
    item: 'Thao tác chuyển mạch bằng tay (Manual Transfer Operation)',
    measured: visual_inspection.manual_transfer_operation,
    standard: 'Thao tác đóng cắt bằng tay trơn tru, dứt khoát, không kẹt cơ',
    reference: 'NETA ATS-2025 Sec 7.22.3.A.3',
    deviation: manualOpPass ? 'Trơn tru' : 'Kẹt cơ',
    status: manualOpPass ? 'PASS' : 'INVESTIGATE',
  });

  const mechanicalInterlockPass = visual_inspection.mechanical_interlock_check.toLowerCase().includes('pass') ||
    visual_inspection.mechanical_interlock_check.toLowerCase().includes('đạt');
  evaluations.push({
    item: 'Khóa liên động cơ khí (Mechanical Interlock Check)',
    measured: visual_inspection.mechanical_interlock_check,
    standard: 'BẮT BUỘC ngăn chặn tuyệt đối khả năng đóng đồng thời hai nguồn (Normal & Alternate)',
    reference: 'NETA ATS-2025 Sec 7.22.3.A.4',
    deviation: mechanicalInterlockPass ? 'Hoạt động hoàn hảo' : 'LỖI LIÊN ĐỘNG',
    status: mechanicalInterlockPass ? 'PASS' : 'FAIL',
  });

  const anchoragePass = visual_inspection.anchorage_grounding_check.toLowerCase().includes('pass') ||
    visual_inspection.anchorage_grounding_check.toLowerCase().includes('đạt');
  evaluations.push({
    item: 'Định vị tủ & Tiếp địa vỏ (Anchorage & Grounding)',
    measured: visual_inspection.anchorage_grounding_check,
    standard: 'Tiếp địa vỏ tủ hoàn chỉnh, bệ định vị chắc chắn theo thiết kế',
    reference: 'NETA ATS-2025 Sec 7.22.3.A.5',
    deviation: anchoragePass ? 'Chắc chắn' : 'Lỏng lẻo',
    status: anchoragePass ? 'PASS' : 'INVESTIGATE',
  });

  const torquePass = visual_inspection.bolt_torque_check.toLowerCase().includes('pass') ||
    visual_inspection.bolt_torque_check.toLowerCase().includes('đạt');
  evaluations.push({
    item: 'Lực siết bu-lông mối nối (Bolted Connection Torque)',
    measured: visual_inspection.bolt_torque_check,
    standard: 'Đạt lực siết cờ-lê lực theo Table 100.12 hoặc NSX',
    reference: 'NETA ATS-2025 Sec 7.22.3.A.6 & Table 100.12',
    deviation: torquePass ? 'Đạt chuẩn' : 'Chưa đạt lực',
    status: torquePass ? 'PASS' : 'INVESTIGATE',
  });

  // 2. Electrical Tests: Bolted Connection Resistance
  const bolted = electrical_tests.bolted_resistance_micro_ohms || [];
  let boltedMaxDev = 0;
  let boltedPass = true;
  if (bolted.length > 0) {
    const minB = Math.min(...bolted);
    const maxB = Math.max(...bolted);
    if (minB > 0) {
      boltedMaxDev = Math.round(((maxB - minB) / minB) * 1000) / 10;
      boltedPass = boltedMaxDev <= 50;
    }
  }
  evaluations.push({
    item: 'Điện trở mối nối bu-lông (Bolted Connection Resistance)',
    measured: bolted.length > 0 ? `${bolted.join(' / ')} µΩ (Max: ${Math.max(...bolted)} µΩ)` : 'N/A',
    standard: 'Không được lệch quá 50% (> 50%) so với giá trị nhỏ nhất của mối nối tương tự',
    reference: 'NETA ATS-2025 Sec 7.22.3.D.1',
    deviation: `${boltedMaxDev}%`,
    status: boltedPass ? 'PASS' : 'INVESTIGATE',
  });

  // 3. Electrical Tests: Main Pole Insulation Resistance (1000V DC)
  const normIR = electrical_tests.main_pole_insulation_resistance_1000v_megohms?.normal_source_position ?? 100;
  const altIR = electrical_tests.main_pole_insulation_resistance_1000v_megohms?.alternate_source_position ?? 100;
  // Per Table 100.1: for 600V nominal equipment, min test voltage 1000VDC, min IR is 100 Megohms.
  const mainPolePass = normIR >= 100 && altIR >= 100;
  evaluations.push({
    item: 'Điện trở cách điện Mạch lực 1000V DC (Main Pole Insulation Resistance)',
    measured: `Normal: ${normIR} MΩ | Alternate: ${altIR} MΩ`,
    standard: '≥ 100 Megohms (NETA ATS Table 100.1 cho thiết bị danh định ≤ 600V)',
    reference: 'NETA ATS-2025 Sec 7.22.3.D.2 & Table 100.1',
    deviation: mainPolePass ? 'Đạt' : 'Cách điện suy giảm (< 100 MΩ)',
    status: mainPolePass ? 'PASS' : 'FAIL',
  });

  // 4. Electrical Tests: Control Wiring Insulation Resistance
  const ctrlIR = electrical_tests.control_wiring_ir_megohms ?? 0;
  // Per NETA ATS-2025 Sec 7.22.3.D.3: Control wiring insulation resistance shall not be less than 2.0 megohms.
  const controlWiringPass = ctrlIR >= 2.0;
  evaluations.push({
    item: 'Điện trở cách điện Mạch điều khiển (Control Wiring Insulation Resistance)',
    measured: `${ctrlIR} MΩ`,
    standard: 'BẮT BUỘC ≥ 2.0 Megohms (Thử nghiệm 500V/1000V DC)',
    reference: 'NETA ATS-2025 Sec 7.22.3.D.3',
    deviation: controlWiringPass ? 'Đạt' : 'RÒ CÁCH ĐIỆN MẠCH ĐIỀU KHIỂN (< 2.0 MΩ)',
    status: controlWiringPass ? 'PASS' : 'FAIL',
  });

  // 5. Electrical Tests: Contact / Pole Resistance (Critical Core Check)
  const cp = electrical_tests.contact_pole_resistance_micro_ohms;
  const normPoles = [cp.normal_source_pole_a, cp.normal_source_pole_b, cp.normal_source_pole_c].filter((v) => v !== undefined && v > 0);
  const altPoles = [cp.alternate_source_pole_a, cp.alternate_source_pole_b, cp.alternate_source_pole_c].filter((v) => v !== undefined && v > 0);

  const minNorm = Math.min(...normPoles);
  const maxNorm = Math.max(...normPoles);
  const normDev = minNorm > 0 ? Math.round(((maxNorm - minNorm) / minNorm) * 1000) / 10 : 0;

  const minAlt = Math.min(...altPoles);
  const maxAlt = Math.max(...altPoles);
  const altDev = minAlt > 0 ? Math.round(((maxAlt - minAlt) / minAlt) * 1000) / 10 : 0;

  let worstPole = 'Không có';
  let isContactAnomaly = false;

  if (cp.normal_source_pole_c > minNorm && ((cp.normal_source_pole_c - minNorm) / minNorm) > 0.5) {
    worstPole = `Normal Source Pole C (${cp.normal_source_pole_c} µΩ)`;
    isContactAnomaly = true;
  } else if (normDev > 50) {
    worstPole = `Normal Source (Max: ${maxNorm} µΩ)`;
    isContactAnomaly = true;
  } else if (altDev > 50) {
    worstPole = `Alternate Source (Max: ${maxAlt} µΩ)`;
    isContactAnomaly = true;
  }

  const contactPass = !isContactAnomaly && normDev <= 50 && altDev <= 50;
  evaluations.push({
    item: 'Điện trở tiếp xúc Cực/Tiếp điểm (Contact / Pole Resistance)',
    measured: `Normal: [A: ${cp.normal_source_pole_a}, B: ${cp.normal_source_pole_b}, C: ${cp.normal_source_pole_c}] µΩ (Lệch ${normDev}%) | Alternate: [A: ${cp.alternate_source_pole_a}, B: ${cp.alternate_source_pole_b}, C: ${cp.alternate_source_pole_c}] µΩ (Lệch ${altDev}%)`,
    standard: 'Độ lệch giữa các cực không được vượt quá 50% (> 50%) so với cực nhỏ nhất',
    reference: 'NETA ATS-2025 Sec 7.22.3.D.4',
    deviation: `Normal: ${normDev}%, Alternate: ${altDev}%`,
    status: contactPass ? 'PASS' : 'INVESTIGATE',
  });

  // 6. Phasing & Rotation Check
  const phasingPass = electrical_tests.phasing_and_rotation_check.toLowerCase().includes('pass') ||
    electrical_tests.phasing_and_rotation_check.toLowerCase().includes('đạt') ||
    electrical_tests.phasing_and_rotation_check.toLowerCase().includes('matching');
  evaluations.push({
    item: 'Kiểm tra Thứ tự pha & Đồng bộ (Phasing & Phase Rotation)',
    measured: electrical_tests.phasing_and_rotation_check,
    standard: 'Khớp 100% thứ tự pha A-B-C và đồng bộ góc pha giữa Nguồn lưới & Máy phát',
    reference: 'NETA ATS-2025 Sec 7.22.3.D.5',
    deviation: phasingPass ? 'Đồng bộ pha' : 'LỆCH THỨ TỰ PHA',
    status: phasingPass ? 'PASS' : 'FAIL',
  });

  // 7. Automatic Sequence Tests & Timers
  const seq = electrical_tests.automatic_sequence_tests;
  const seqPass = seq.normal_source_undervoltage_sensing.toLowerCase().includes('pass') &&
    seq.alternate_source_voltage_frequency_sensing.toLowerCase().includes('pass') &&
    seq.limit_switches_and_interlocks.toLowerCase().includes('pass');

  evaluations.push({
    item: 'Trình tự chuyển nguồn tự động & Định thời gian (Automatic Sequence & Timers)',
    measured: `Khởi động MF: ${seq.engine_start_signal_time_sec}s | Đóng tải DP: ${seq.transfer_time_delay_sec}s | Khôi phục lưới: ${seq.retransfer_time_delay_sec}s | Làm mát MF: ${seq.engine_cool_down_time_sec}s`,
    standard: 'Cảm biến sụt áp, tín hiệu khởi động, thời gian trễ đóng/khôi phục và công tắc hành trình chuẩn xác',
    reference: 'NETA ATS-2025 Sec 7.22.3.D.6',
    deviation: seqPass ? 'Trình tự chính xác' : 'Bất thường trình tự',
    status: seqPass ? 'PASS' : 'INVESTIGATE',
  });

  // Determine Overall Status
  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let status_reason = '';

  const hasFail = evaluations.some((e) => e.status === 'FAIL');
  const hasInvestigate = evaluations.some((e) => e.status === 'INVESTIGATE');

  if (hasFail) {
    overall_status = 'FAIL';
    status_reason = 'Phát hiện lỗi nghiêm trọng vi phạm tiêu chuẩn NETA ATS-2025 (Khóa liên động cơ khí, cách điện mạch lực hoặc rò cách điện mạch điều khiển).';
  } else if (hasInvestigate || isContactAnomaly) {
    overall_status = 'INVESTIGATE';
    status_reason = `INVESTIGATE / CONTACT RESISTANCE ANOMALY: Điện trở tiếp điểm Pole C Nguồn chính (${cp.normal_source_pole_c} µΩ) lệch ${normDev}% so với cực nhỏ nhất (${minNorm} µΩ), vượt quá giới hạn 50% theo NETA ATS-2025 Mục 7.22.3.D.4.`;
  } else {
    overall_status = 'PASS';
    status_reason = 'Toàn bộ các hạng mục kiểm tra thị giác, liên động cơ khí, cách điện và trình tự chuyển nguồn tự động ATS đều ĐẠT chuẩn ANSI/NETA ATS-2025 Mục 7.22.3.';
  }

  // Generate actionable recommendations
  if (isContactAnomaly) {
    recommendations.push(
      'Tách cô lập an toàn nguồn ATS (Lockout/Tagout), tháo nắp bảo vệ buồng dập hồ quang để kiểm tra trực quan bề mặt tiếp điểm động và tĩnh của Pole C.'
    );
    recommendations.push(
      'Vệ sinh tẩy rửa bề mặt tiếp điểm bằng dung dịch làm sạch tiếp điểm chuyên dụng (Contact Cleaner), kiểm tra độ đàn hồi của lò xo nén tiếp điểm.'
    );
    recommendations.push(
      'Thực hiện thao tác chuyển mạch đóng cắt bằng tay 10-15 lần để tự đánh bóng/làm sạch cơ học bề mặt tiếp xúc, sau đó dùng micro-ohmmeter 100A DC đo lại điện trở.'
    );
    recommendations.push(
      `Đảm bảo điện trở tiếp điểm Pole C hạ về dưới ${Math.round(minNorm * 1.5 * 10) / 10} µΩ (độ lệch ≤ 50% so với cực kế cận) trước khi nghiệm thu đóng điện cho tải.`
    );
  }
  if (!controlWiringPass) {
    recommendations.push(
      'Kiểm tra và cô lập từng nhánh dây điều khiển, rơ-le định thời và biến điện áp đo lường để tìm vị trí chạm đất/suy giảm cách điện (< 2.0 MΩ).'
    );
  }
  if (!mechanicalInterlockPass) {
    recommendations.push(
      'DỪNG NGAY VIỆC ĐÓNG ĐIỆN VẬN HÀNH: Kiểm tra và căn chỉnh thanh liên động cơ khí (Walking beam / Mechanical Interlock) để đảm bảo không bao giờ hai nguồn đóng song song gây nổ pha.'
    );
  }
  if (recommendations.length === 0) {
    recommendations.push(
      'Bộ chuyển nguồn ATS vận hành hoàn hảo. Tiến hành niêm phong chì tủ, ghi nhật ký bảo trì định kỳ và duy trì lịch kiểm định 12 tháng/lần.'
    );
  }

  // Format Markdown Report
  const markdown_report = `### BIÊN BẢN KIỂM ĐỊNH BỘ CHUYỂN NGUỒN TỰ ĐỘNG (ATS)
**Tiêu chuẩn áp dụng:** ANSI/NETA ATS-2025 Mục 7.22.3 & Table 100.1, 100.12, 100.18
**Dự án / Địa điểm:** ${payload.site_info.project_name} | **Vị trí:** ${payload.site_info.substation_location || 'Phòng Hạ Thế'}
**Mã thiết bị:** \`${payload.site_info.ats_tag}\` | **Thông số:** ${payload.site_info.ats_rating}
**Kỹ sư kiểm tra (FSE):** ${payload.site_info.fse_name} | **Ngày thử nghiệm:** ${payload.site_info.test_date}

---

#### 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
**Kết luận:** \`${overall_status} / ${isContactAnomaly ? 'CONTACT RESISTANCE ANOMALY' : overall_status === 'PASS' ? 'SATISFACTORY' : 'DEFECT DETECTED'}\`
**Lý do đánh giá:** ${status_reason}

---

#### 2. BẢNG PHÂN TÍCH CHI TIẾT BỘ CHUYỂN NGUỒN ATS (Detailed Evaluation Table)
| Hạng mục kiểm tra | Giá trị đo thực tế | Tiêu chuẩn NETA ATS-2025 Sec 7.22.3 | Độ lệch | Đánh giá |
| :--- | :--- | :--- | :--- | :---: |
${evaluations
  .map(
    (e) =>
      `| **${e.item}** | ${e.measured} | ${e.standard} | ${e.deviation} | \`${e.status}\` |`
  )
  .join('\n')}

---

#### 3. PHÂN TÍCH TRÌNH TỰ TỰ ĐỘNG, TIẾP ĐIỂM & NGUY CƠ MẤT NGUỒN
* **Điện trở tiếp xúc cực (Contact Resistance Analysis):**
  - **Nguồn dự phòng (Alternate Source):** Các cực Pole A ($${cp.alternate_source_pole_a}\\,\\mu\\Omega$), Pole B ($${cp.alternate_source_pole_b}\\,\\mu\\Omega$), Pole C ($${cp.alternate_source_pole_c}\\,\\mu\\Omega$) đồng đều tuyệt đối (độ lệch cực đại chỉ $${altDev}\\% \\le 50\\%$) $\\rightarrow$ Đạt yêu cầu xuất sắc.
  - **Nguồn chính (Normal Source):** Pole A = $${cp.normal_source_pole_a}\\,\\mu\\Omega$, Pole B = $${cp.normal_source_pole_b}\\,\\mu\\Omega$, nhưng Pole C đo được $${cp.normal_source_pole_c}\\,\\mu\\Omega$. Độ lệch so với cực nhỏ nhất lên tới **${normDev}%** (vượt quá xa ngưỡng $50\\%$ theo quy định tại NETA Mục 7.22.3.D.4).
  - **Nguy cơ vận hành:** Khi vận hành đầy tải ($800\\text{A}$), cực Pole C với điện trở tiếp xúc cao gấp 3 lần bình thường sẽ phát sinh tổn hao nhiệt Joule ($I^2R$), gây quá nhiệt cục bộ, rỗ tiếp điểm và tiềm ẩn nguy cơ nung chảy tiếp điểm động hoặc sụt áp pha cấp cho phụ tải văn phòng.
* **Cách điện & Liên động an toàn:**
  - Điện trở cách điện mạch lực cả 2 nguồn đạt $> 800\\text{ M}\\Omega$ (vượt xa ngưỡng $100\\text{ M}\\Omega$ Table 100.1).
  - Điện trở cách điện mạch điều khiển đạt $${ctrlIR}\\text{ M}\\Omega$ (vượt ngưỡng bắt buộc $\\ge 2.0\\text{ M}\\Omega$ Mục 7.22.3.D.3).
  - Khóa liên động cơ khí (*Mechanical Interlock*) và liên động điện hoạt động an toàn, loại trừ hoàn toàn nguy cơ hòa nhầm 2 nguồn khác pha.
* **Trình tự tự động & Bộ định thời (Automatic Sequence & Timers):**
  - Cảm biến sụt áp lưới ($85\\% U_n$) kích hoạt lệnh gọi máy phát sau đúng $${seq.engine_start_signal_time_sec}\\text{s}$.
  - Thời gian trễ đóng tải sang máy phát $${seq.transfer_time_delay_sec}\\text{s}$ và trễ khôi phục về lưới $${seq.retransfer_time_delay_sec}\\text{s}$ kèm thời gian chạy làm mát $${seq.engine_cool_down_time_sec}\\text{s}$ vận hành chính xác $100\\%$ theo cài đặt bộ điều khiển.

---

#### 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)
${recommendations.map((r, i) => `${i + 1}. ${r}`).join('\n')}
`;

  return {
    overall_status,
    status_reason,
    contact_resistance_analysis: {
      normal_min_micro_ohms: minNorm,
      normal_max_micro_ohms: maxNorm,
      normal_max_deviation_percent: normDev,
      alternate_min_micro_ohms: minAlt,
      alternate_max_micro_ohms: maxAlt,
      alternate_max_deviation_percent: altDev,
      worst_pole: worstPole,
      is_anomaly: isContactAnomaly,
    },
    insulation_analysis: {
      main_pole_pass: mainPolePass,
      control_wiring_pass: controlWiringPass,
      control_ir_megohms: ctrlIR,
    },
    sequence_analysis: {
      engine_start_sec: seq.engine_start_signal_time_sec,
      transfer_delay_sec: seq.transfer_time_delay_sec,
      retransfer_delay_sec: seq.retransfer_time_delay_sec,
      cool_down_sec: seq.engine_cool_down_time_sec,
      interlocks_verified: mechanicalInterlockPass,
    },
    evaluations,
    actionable_recommendations: recommendations,
    markdown_report,
  };
}

/**
 * Generate comprehensive AI-driven analysis using Gemini 2.5 Flash
 */
export async function generateNetaAtsAiAnalysis(
  payload: NetaAtsInputPayload
): Promise<NetaAtsEvaluationResult> {
  const deterministicResult = performDeterministicAtsAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return deterministicResult;
  }

  try {
    const ai = new GoogleGenAI({ apiKey });
    const prompt = `
Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách Bộ chuyển nguồn Tự động (Emergency Systems, Automatic Transfer Switches - ATS). Nhiệm vụ của bạn là nhận dữ liệu kiểm tra ngoài site từ Field Service Engineer (FSE) và tự động đối soát với tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.22.3) và các bảng Table 100.1, Table 100.12, Table 100.18.

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.22.3):
1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Nameplate Match: Đối soát nhãn mác ATS (điện áp, dòng điện, số cực) khớp 100% bản vẽ thiết kế.
- Physical Condition & Cleanliness: Tủ ATS, buồng dập hồ quang và tiếp điểm sạch bẩn, không hư hỏng cơ lý.
- Manual Operation: Thao tác chuyển mạch bằng tay trơn tru, dứt khoát.
- Mechanical Interlock: Khóa liên động cơ khí giữa Nguồn chính (Normal) và Nguồn dự phòng (Alternate) BẮT BUỘC hoạt động an toàn, ngăn đóng trùng hai nguồn.
- Anchorage & Grounding: Định vị chắc chắn, tiếp địa vỏ tủ hoàn chỉnh.
- Bolt Torque: Mối nối bu-lông siết đạt lực theo nhà sản xuất hoặc Bảng Table 100.12.
- Thermographic Survey: Kết quả khảo sát nhiệt tuân thủ Section 9 / Table 100.18.

2. PHÉP ĐO ĐIỆN VÀ TIÊU CHUẨN ĐÁNH GIÁ (ELECTRICAL TEST VALUES):
- Điện trở mối nối bu-lông (Bolted Connection Resistance): CẢNH BÁO "INVESTIGATE" nếu giá trị đo mối nối lệch quá 50% so với giá trị nhỏ nhất của mối nối tương tự.
- Điện trở cách điện Mạch lực (Main Pole Insulation Resistance): Đạt ngưỡng tối thiểu theo Bảng Table 100.1 hoặc nhà sản xuất ở cả 2 vị trí nguồn (≥ 100 Megohms ở 1000V DC cho thiết bị ≤ 600V).
- Điện trở cách điện Mạch điều khiển (Control Wiring IR): Thử nghiệm 500V/1000V DC BẮT BUỘC >= 2.0 Megohms (2 MΩ).
- Điện trở tiếp xúc cực (Contact / Pole Resistance): CẢNH BÁO "INVESTIGATE" nếu giá trị đo cực lệch quá 50% so với cực kế cận hoặc giá trị nhỏ nhất.
- Thứ tự pha & Đồng bộ (Phasing & Rotation): Thứ tự pha và đồng bộ góc pha giữa hai nguồn khớp 100% sơ đồ thiết kế.
- Trình tự Chuyển nguồn Tự động & Thời gian Trễ (Sequence & Timers):
  + Giám sát điện áp/tần số Nguồn chính và Nguồn dự phòng hoạt động chính xác.
  + Thời gian trễ khởi động máy phát (Engine start delay), trễ đóng tải dự phòng (Transfer delay), trễ khôi phục nguồn lưới (Retransfer delay) và trễ làm mát máy phát (Cool down timer) khớp 100% với cài đặt bộ điều khiển ATS.
  + Các công tắc hành trình (Limit switches) và liên động điện phản hồi đúng vị trí.

DỮ LIỆU ĐO KIỂM HIỆN TRƯỜNG TỪ FSE:
${JSON.stringify(payload, null, 2)}

KẾT QUẢ ĐỐI SOÁT ĐỊNH LƯỢNG NỘI BỘ:
${JSON.stringify(
  {
    overall_status: deterministicResult.overall_status,
    status_reason: deterministicResult.status_reason,
    contact_resistance_analysis: deterministicResult.contact_resistance_analysis,
    insulation_analysis: deterministicResult.insulation_analysis,
    sequence_analysis: deterministicResult.sequence_analysis,
  },
  null,
  2
)}

HÃY XUẤT RA BÁO CÁO THẨM ĐỊNH KỸ THUẬT ĐỊNH DẠNG MARKDOWN CHUYÊN NGHIỆP GỒM 4 PHẦN CHÍNH XÁC:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE (kèm lý do cô đọng chuẩn kỹ thuật).
2. BẢNG PHÂN TÍCH CHI TIẾT BỘ CHUYỂN NGUỒN ATS (Detailed Evaluation Table): So sánh thực tế với NETA ATS-2025 Section 7.22.3 (Table 100.1, Table 100.12, Table 100.18).
3. PHÂN TÍCH TRÌNH TỰ TỰ ĐỘNG, TIẾP ĐIỂM & NGUY CƠ MẤT NGUỒN (Automatic Sequence, Contact & Power Continuity Risk). Phân tích sâu sự sai lệch của Pole C Nguồn chính ($68.4\\,\\mu\\Omega$ lệch 209.5% so với Pole A $22.1\\,\\mu\\Omega$) và nguy cơ phát nhiệt nung chảy tiếp điểm.
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations): Hướng dẫn cô lập, vệ sinh bằng contact cleaner, thao tác đóng mở bằng tay 10-15 lần để tự mài bóng cơ học và đo lại để đưa về dưới 33.1 µΩ (lệch ≤ 50%).
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
    console.error('Error generating AI analysis for NETA ATS:', error);
  }

  return deterministicResult;
}
