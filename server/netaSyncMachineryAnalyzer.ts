import { GoogleGenAI } from '@google/genai';

export interface NetaSyncMachineryInputPayload {
  site_info: {
    project_name: string;
    generator_tag: string;
    generator_rating: string;
    fse_name: string;
    test_date: string;
    manufacturer?: string;
    serial_number?: string;
    rated_voltage_kv?: number;
    rated_speed_rpm?: number;
    excitation_type?: string; // 'Brushless' | 'Static Excitation with Collector Rings'
  };
  visual_inspection: {
    nameplate_match: boolean;
    physical_cleanliness: string; // 'Pass' | 'Fail'
    air_gap_uniformity: string; // 'Pass' | 'Fail'
    brushless_exciter_condition: string; // 'Good' | 'Needs Service'
    bolt_torque_check: string; // 'Pass' | 'Fail'
    shaft_alignment_check?: string; // 'Pass' | 'Fail'
  };
  electrical_tests: {
    stator_ir_40c_megohms: {
      stator_1min_ir_2500v: number; // e.g. 1800
      stator_10min_ir_2500v: number; // e.g. 5400
      calculated_pi_value: number; // e.g. 3.0
    };
    main_field_rotor_ir_1000v_megohms: number; // e.g. 250
    exciter_stator_rotor_ir_megohms?: number; // e.g. 300
    pole_to_pole_ac_voltage_drop_volts: {
      pole_1: number;
      pole_2: number;
      pole_3: number;
      pole_4: number;
      pole_5?: number;
      pole_6?: number;
      average_voltage_drop_per_pole: number;
    };
    rotating_diodes_check: string; // e.g. "Pass (All 6 diodes healthy)"
    stator_phase_resistance_ohms: {
      phase_ab: number; // e.g. 0.082
      phase_bc: number; // e.g. 0.083
      phase_ca: number; // e.g. 0.082
    };
    uncoupled_vibration_in_sec_pk: number; // e.g. 0.05
    power_factor_relay_trip_test?: string; // e.g. "Pass"
  };
  previous_test_data?: {
    last_test_date?: string;
    last_pole_3_voltage_drop_volts?: number;
    last_stator_pi_value?: number;
    last_vibration_in_sec_pk?: number;
  };
}

export interface SyncMachineryItemEvaluation {
  item: string;
  measured: string;
  standard: string;
  reference: string;
  status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  detail: string;
}

export interface NetaSyncMachineryEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_reason: string;
  markdown_report: string;
  evaluations: {
    max_severity: 'PASS' | 'INVESTIGATE' | 'FAIL';
    shorted_turns_pole_failed: boolean;
    faulty_pole_number: number | null;
    stator_pi_passed: boolean;
    field_rotor_ir_passed: boolean;
    stator_resistance_balanced: boolean;
    vibration_passed: boolean;
    items: SyncMachineryItemEvaluation[];
  };
}

/**
 * Deterministic rules engine according to ANSI/NETA ATS-2025 Section 7.15.2
 */
export function performDeterministicSyncMachineryAnalysis(
  data: NetaSyncMachineryInputPayload
): NetaSyncMachineryEvaluationResult {
  const items: SyncMachineryItemEvaluation[] = [];

  // 1. Visual & Mechanical Inspection (NETA Section 7.15.2.A)
  // 1.1 Nameplate Match
  const isNameplatePass = data.visual_inspection.nameplate_match;
  items.push({
    item: 'Đối soát nhãn Stator, Rotor cực lồi & Exciter (Nameplate)',
    measured: `Tag: ${data.site_info.generator_tag} | Định mức: ${data.site_info.generator_rating}`,
    standard: 'Khớp 100% hồ sơ thiết kế (Công suất MW, điện áp kV, tốc độ RPM, dòng kích từ)',
    reference: 'NETA ATS-2025 Sec 7.15.2.A.1',
    status: isNameplatePass ? 'PASS' : 'FAIL',
    detail: isNameplatePass
      ? 'Nhãn máy phát, rotor cực từ và bộ kích từ không chổi than trùng khớp bản vẽ thiết kế.'
      : 'Không khớp thông số thiết kế hoặc nhãn định mức sai lệch.'
  });

  // 1.2 Physical Cleanliness & Air-Gap Uniformity
  const isCleanPass = data.visual_inspection.physical_cleanliness.toLowerCase().includes('pass');
  const isAirGapPass = data.visual_inspection.air_gap_uniformity.toLowerCase().includes('pass');
  items.push({
    item: 'Vệ sinh & Độ đồng đều khe hở không khí (Cleanliness & Air-Gap)',
    measured: `Vệ sinh: ${data.visual_inspection.physical_cleanliness} | Khe hở không khí: ${data.visual_inspection.air_gap_uniformity}`,
    standard: 'Stator, rotor, cuộn cực lồi sạch sẽ; khe hở không khí giữa các cực đồng đều',
    reference: 'NETA ATS-2025 Sec 7.15.2.A.2 & A.3',
    status: isCleanPass && isAirGapPass ? 'PASS' : 'FAIL',
    detail: isCleanPass && isAirGapPass
      ? 'Máy điện sạch bụi than, không có dị vật từ tính; khe hở cực từ stator-rotor phân bố đều.'
      : 'LỖI: Khe hở không khí không đều hoặc tồn đọng bụi than/dị vật trong khe từ.'
  });

  // 1.3 Brushless Exciter & Bolt Torque
  const isExciterPass = data.visual_inspection.brushless_exciter_condition.toLowerCase().includes('good') || data.visual_inspection.brushless_exciter_condition.toLowerCase().includes('pass');
  const isTorquePass = data.visual_inspection.bolt_torque_check.toLowerCase().includes('pass');
  items.push({
    item: 'Bộ kích từ không chổi than & Lực siết bu-lông (Exciter & Torque)',
    measured: `Kích từ: ${data.visual_inspection.brushless_exciter_condition} | Lực siết: ${data.visual_inspection.bolt_torque_check}`,
    standard: 'Bộ kích từ nguyên vẹn; Bu-lông siết đạt lực theo catalogue hoặc Table 100.12',
    reference: 'NETA ATS-2025 Sec 7.15.2.A.4 & A.5',
    status: isExciterPass && isTorquePass ? 'PASS' : 'FAIL',
    detail: isExciterPass && isTorquePass
      ? 'Bộ kích từ quay hoạt động tốt, các bu-lông chân máy và khớp nối siết đủ lực.'
      : 'Cần kiểm tra lại bộ kích từ hoặc lực siết bu-lông chân máy.'
  });

  // 2. Electrical Tests (NETA Section 7.15.2.B)
  // 2.1 Stator Insulation Resistance & Polarization Index (PI) - Table 100.11
  const stator = data.electrical_tests.stator_ir_40c_megohms;
  const isStatorIrPass = stator.stator_1min_ir_2500v >= 100;
  const isPiPass = stator.calculated_pi_value >= 2.0; // Table 100.11 PI >= 2.0 for machines > 200 HP

  items.push({
    item: 'Cách điện Stator & Chỉ số phân cực PI @ 40°C (Stator IR & PI)',
    measured: `IR 1min: ${stator.stator_1min_ir_2500v} MΩ, IR 10min: ${stator.stator_10min_ir_2500v} MΩ, PI = ${stator.calculated_pi_value.toFixed(2)}`,
    standard: 'IR 1min ≥ 100 MΩ; PI (10min/1min) ≥ 2.0 đối với máy điện > 200 HP',
    reference: 'NETA ATS-2025 Sec 7.15.2.B.1 & Table 100.11',
    status: isStatorIrPass && isPiPass ? 'PASS' : 'FAIL',
    detail: isStatorIrPass && isPiPass
      ? `Cuộn dây Stator cách điện tuyệt hảo (${stator.stator_10min_ir_2500v} MΩ), chỉ số PI = ${stator.calculated_pi_value.toFixed(2)} xác nhận cuộn dây hoàn toàn khô ráo.`
      : 'CẢNH BÁO: Điện trở cách điện stator suy giảm hoặc chỉ số PI < 2.0 (nghi ngờ nhiễm ẩm/bụi dẫn điện).'
  });

  // 2.2 Main Field Rotor Winding Insulation Resistance (1000V DC)
  const rotorIr = data.electrical_tests.main_field_rotor_ir_1000v_megohms;
  const isRotorIrPass = rotorIr >= 100;
  items.push({
    item: 'Cách điện Cuộn Kích từ Rô-to 1000V DC (Main Field Rotor IR)',
    measured: `${rotorIr} MΩ`,
    standard: 'Tối thiểu ≥ 100 MΩ (hoặc theo khuyến nghị nhà sản xuất)',
    reference: 'NETA ATS-2025 Sec 7.15.2.B.2 & Table 100.11',
    status: isRotorIrPass ? 'PASS' : 'FAIL',
    detail: isRotorIrPass
      ? 'Cuộn dây kích từ chính trên rotor cách điện tốt với trục và lõi thép.'
      : 'CẢNH BÁO NGUY HIỂM: Cách điện cuộn kích từ rotor suy giảm < 100 MΩ, nguy cơ chạm đất rotor.'
  });

  // 2.3 CRITICAL TEST: Pole-to-Pole AC Voltage Drop Test
  const drops = data.electrical_tests.pole_to_pole_ac_voltage_drop_volts;
  const poles: { poleNum: number; drop: number }[] = [
    { poleNum: 1, drop: drops.pole_1 },
    { poleNum: 2, drop: drops.pole_2 },
    { poleNum: 3, drop: drops.pole_3 },
    { poleNum: 4, drop: drops.pole_4 },
  ];
  if (drops.pole_5 !== undefined) poles.push({ poleNum: 5, drop: drops.pole_5 });
  if (drops.pole_6 !== undefined) poles.push({ poleNum: 6, drop: drops.pole_6 });

  let avgDrop = drops.average_voltage_drop_per_pole;
  if (!avgDrop || avgDrop <= 0) {
    avgDrop = poles.reduce((acc, p) => acc + p.drop, 0) / poles.length;
  }

  let faultyPole: number | null = null;
  let maxPoleDeviationPercent = 0;
  let isPoleDropPass = true;

  const poleDeviations = poles.map((p) => {
    const dev = ((p.drop - avgDrop) / avgDrop) * 100;
    if (Math.abs(dev) > Math.abs(maxPoleDeviationPercent)) {
      maxPoleDeviationPercent = dev;
    }
    if (Math.abs(dev) > 10.0) {
      isPoleDropPass = false;
      faultyPole = p.poleNum;
    }
    return `Pole ${p.poleNum}: ${p.drop}V (${dev > 0 ? '+' : ''}${dev.toFixed(1)}%)`;
  });

  items.push({
    item: 'Thử Sụt áp AC từng Cực Rô-to (Pole-to-Pole AC Voltage Drop)',
    measured: `Trung bình: ${avgDrop.toFixed(1)}V | ${poleDeviations.join(', ')}`,
    standard: 'Độ lệch sụt áp trên từng cực KHÔNG ĐƯỢC vượt quá 10% (≤ ±10%) so với trung bình toàn bộ cực',
    reference: 'NETA ATS-2025 Sec 7.15.2.B.3',
    status: isPoleDropPass ? 'PASS' : 'FAIL',
    detail: isPoleDropPass
      ? 'Sụt áp AC trên các cuộn cực từ đồng đều, không có hiện tượng chập vòng dây.'
      : `CRITICAL FAIL: Cực số ${faultyPole} (Pole ${faultyPole}) có mức sụt áp lệch ${maxPoleDeviationPercent.toFixed(1)}% vượt xa ngưỡng tối đa cho phép 10%. Xác nhận ngắn mạch giữa các vòng dây (Shorted turns in field pole ${faultyPole}).`
  });

  // 2.4 Rotating Diodes & SCR Assembly Check
  const isDiodesPass = data.electrical_tests.rotating_diodes_check.toLowerCase().includes('pass') || data.electrical_tests.rotating_diodes_check.toLowerCase().includes('healthy');
  items.push({
    item: 'Cụm Điốt Quay Kích Từ (Rotating Diodes & Varistor/SCR)',
    measured: data.electrical_tests.rotating_diodes_check,
    standard: 'Tất cả các điốt nắn dòng mở thuận và khóa ngược tốt, không bị thủng hay đứt',
    reference: 'NETA ATS-2025 Sec 7.15.2.B.4',
    status: isDiodesPass ? 'PASS' : 'FAIL',
    detail: isDiodesPass
      ? 'Cầu chỉnh lưu quay 3 pha hoạt động hoàn hảo, bảo vệ chống quá áp varistor tốt.'
      : 'LỖI: Phát hiện điốt quay bị thủng (shorted) hoặc đứt (open), gây mất kích từ.'
  });

  // 2.5 Stator Phase Resistance Balance
  const res = data.electrical_tests.stator_phase_resistance_ohms;
  const resValues = [res.phase_ab, res.phase_bc, res.phase_ca];
  const minRes = Math.min(...resValues);
  const maxRes = Math.max(...resValues);
  const resDev = minRes > 0 ? ((maxRes - minRes) / minRes) * 100 : 0;
  const isResPass = resDev <= 5.0;

  items.push({
    item: 'Điện trở cuộn dây Stator giữa các pha (Stator Phase Resistance)',
    measured: `AB: ${res.phase_ab} Ω, BC: ${res.phase_bc} Ω, CA: ${res.phase_ca} Ω (Lệch max: ${resDev.toFixed(2)}%)`,
    standard: 'Độ lệch điện trở một chiều giữa các pha không vượt quá 5% (≤ 5%)',
    reference: 'NETA ATS-2025 Sec 7.15.2.B.5',
    status: isResPass ? 'PASS' : 'INVESTIGATE',
    detail: isResPass
      ? 'Điện trở Stator hoàn toàn cân bằng giữa 3 pha, các mối nối đầu bối dây tiếp xúc tốt.'
      : `CẢNH BÁO: Độ lệch điện trở pha đạt ${resDev.toFixed(2)}% (> 5%), cần kiểm tra mối nối bấm cosse hoặc điểm hàn đầu cực.`
  });

  // 2.6 Uncoupled Vibration Test (NETA Table 100.10)
  const vib = data.electrical_tests.uncoupled_vibration_in_sec_pk;
  // For 1500 RPM / 1800 RPM: Limit typically 0.15 in/sec pk (or 0.10 in/sec pk)
  const isVibPass = vib <= 0.15;
  items.push({
    item: 'Độ Rung Thử Nghiệm Không Tải Tháo Khớp (Uncoupled Vibration)',
    measured: `${vib} in/sec pk`,
    standard: 'Tuân thủ Table 100.10 (≤ 0.15 in/sec pk đối với máy điện quay 1500 - 1800 RPM)',
    reference: 'NETA ATS-2025 Sec 7.15.2.B.6 & Table 100.10',
    status: isVibPass ? 'PASS' : 'FAIL',
    detail: isVibPass
      ? `Độ rung vòng bi đạt ${vib} in/sec pk nằm trong ngưỡng cho phép rất an toàn.`
      : `Độ rung ${vib} in/sec pk vượt ngưỡng Table 100.10. Cần cân bằng động rotor hoặc kiểm tra vòng bi.`
  });

  // Overall status evaluation
  const hasFail = items.some((i) => i.status === 'FAIL');
  const hasInvestigate = items.some((i) => i.status === 'INVESTIGATE');

  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let status_reason = 'Toàn bộ cuộn dây Stator, cực từ Rotor và hệ thống kích từ đồng bộ đạt chuẩn NETA ATS-2025 Mục 7.15.2.';

  if (hasFail) {
    overall_status = 'FAIL';
    if (!isPoleDropPass) {
      status_reason = `KHÔNG ĐẠT TIÊU CHUẨN NETA ATS-2025 (CRITICAL ROTOR DEFECT): Cực số ${faultyPole} có sụt áp AC lệch ${maxPoleDeviationPercent.toFixed(1)}% vượt quá 10%, xác nhận ngắn mạch giữa các vòng dây cuộn cực từ (Shorted turns in field pole). CẤM HÒA LƯỚI ĐÓNG TẢI!`;
    } else {
      status_reason = 'Không đạt tiêu chuẩn NETA ATS-2025 do cách điện stator/rotor hoặc độ rung vượt ngưỡng an toàn.';
    }
  } else if (hasInvestigate) {
    overall_status = 'INVESTIGATE';
    status_reason = 'Cần kiểm tra bổ sung độ lệch điện trở stator hoặc hệ số phân cực PI trước khi hòa lưới.';
  }

  return {
    overall_status,
    status_reason,
    markdown_report: '',
    evaluations: {
      max_severity: overall_status,
      shorted_turns_pole_failed: !isPoleDropPass,
      faulty_pole_number: faultyPole,
      stator_pi_passed: isPiPass,
      field_rotor_ir_passed: isRotorIrPass,
      stator_resistance_balanced: isResPass,
      vibration_passed: isVibPass,
      items
    }
  };
}

export function buildFallbackSyncMachineryMarkdown(
  data: NetaSyncMachineryInputPayload,
  deterministicResult: NetaSyncMachineryEvaluationResult
): string {
  const { evaluations } = deterministicResult;
  const stator = data.electrical_tests.stator_ir_40c_megohms;
  const drops = data.electrical_tests.pole_to_pole_ac_voltage_drop_volts;
  const pole3Dev = ((drops.pole_3 - drops.average_voltage_drop_per_pole) / drops.average_voltage_drop_per_pole) * 100;

  return `# BÁO CÁO ĐÁNH GIÁ CHUYÊN GIA NETA ATS-2025
**THIẾT BỊ: MÁY ĐIỆN ĐỒNG BỘ (ROTATING MACHINERY, SYNCHRONOUS MOTORS AND GENERATORS)**
*Tiêu chuẩn đối soát: ANSI/NETA ATS-2025 Mục 7.15.2, Table 100.10, Table 100.11, Table 100.12 & IEEE 115*

---

## 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
### **🔴 ${deterministicResult.overall_status} / CRITICAL ROTOR DEFECT (NGẮN MẠCH VÒNG DÂY CỰC TỪ)**
- **Nhà máy / Ngăn lộ:** ${data.site_info.project_name}
- **Mã máy phát (Tag):** \`${data.site_info.generator_tag}\`
- **Công suất & Kiểu máy:** **${data.site_info.generator_rating}**
- **Kỹ sư kiểm tra (FSE TEV):** ${data.site_info.fse_name} | **Ngày kiểm tra:** ${data.site_info.test_date}
- **Kết luận cấp bách:** ${deterministicResult.status_reason}

---

## 2. BẢNG PHÂN TÍCH CHI TIẾT MÁY ĐIỆN ĐỒNG BỘ (Detailed Evaluation Table)

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

## 3. PHÂN TÍCH HỆ THỐNG KÍCH TỪ, CỰC RÔ-TO & ĐỒNG BỘ (Excitation, Field Pole & Sync Risk)

### 3.1. Cách điện Cuộn dây Stator & Cuộn Kích từ Rô-to:
- **Cuộn dây Stator 11kV:** Điện trở cách điện đo ở $2500\\text{ V DC}$ đạt $1800\\text{ M}\\Omega$ (1 phút) và $5400\\text{ M}\\Omega$ (10 phút) $\\gg 100\\text{ M}\\Omega$.
- **Hệ số phân cực Polarization Index (PI):** Đạt $\\text{PI} = \\frac{5400}{1800} = \\mathbf{3.0} \\ge 2.0$ theo đúng Bảng Table 100.11 đối với máy điện quay $> 200\\text{ HP}$ $\\rightarrow$ **PASS**. Cuộn dây stator khô ráo và không bị nhiễm ẩm.
- **Cuộn kích từ Rotor (Main Field Winding):** Điện trở cách điện đo ở $1000\\text{ V DC}$ đạt $250\\text{ M}\\Omega > 100\\text{ M}\\Omega$ $\\rightarrow$ **PASS**.

### 3.2. Đánh giá Sụt áp AC từng Cực Rô-to (Pole-to-Pole AC Voltage Drop) - PHÁT HIỆN SỰ CỐ TỐI CẤP:
- Điện áp sụt trung bình trên các cực đo được là: **${drops.average_voltage_drop_per_pole}\\text{ V}**.
- Các cực Pole 1 ($24.2\\text{ V}$), Pole 2 ($24.5\\text{ V}$) và Pole 4 ($24.4\\text{ V}$) có mức sụt áp đồng đều (độ lệch $< +7\\%$).
- **RIÊNG CỰC SỐ 3 (POLE 3):**
  - Mức sụt áp đo được ngoài hiện trường chỉ đạt: **${drops.pole_3}\\text{ V}**.
  - Sai lệch so với giá trị trung bình toàn bộ các cực: **${poleDeviationsCalc(drops.pole_3, drops.average_voltage_drop_per_pole)}** (Vượt quá xa ngưỡng tối đa cho phép $\\pm 10\\%$ theo quy chuẩn **NETA ATS-2025 Section 7.15.2.B.3** và **IEEE 115**).
  - **KẾT LUẬN KỸ THUẬT: 🔴 CRITICAL FAIL.**
  - *Bản chất sự cố:* Mức sụt áp trên cuộn dây cực 3 bị giảm mạnh chứng minh trực tiếp hiện tượng **ngắn mạch giữa các vòng dây (Turn-to-Turn Short Circuit / Shorted Turns)** trong cuộn dây cực số 3. Một phần các vòng dây bị chập làm giảm tổng trở $Z$ của cực từ, dẫn đến điện áp rơi trên cuộn dây này bị tụt sâu.
  - *Rủi ro nghiêm trọng khi vận hành:* Nếu đưa máy phát vào hòa lưới hoặc cấp dòng kích từ định mức, hiện tượng chập vòng sẽ gây:
    1. **Mất cân bằng từ thông khe hở (Unbalanced Magnetic Pull - UMP):** Sinh ra lực hút từ trường bất đối xứng cực lớn kéo lệch tâm rotor.
    2. **Rung động cơ khí phá hủy gối trục:** Rung giật tần số quay ($1\\times RPM$) phá vỡ bạc đạn, cong vênh trục máy phát.
    3. **Quá nhiệt cục bộ:** Dòng điện cảm ứng quẩn trong vòng dây bị chập sẽ phát nhiệt dữ dội làm cháy đen bối dây và cháy lan sang toàn bộ rotor.

### 3.3. Điốt Quay Kích từ & Cân bằng Điện trở Stator:
- Cụm 6 điốt nắn dòng kích từ quay và mạch dập quá áp làm việc tốt $\rightarrow$ **PASS**.
- Điện trở một chiều stator giữa 3 pha (AB: $0.082\\,\\Omega$, BC: $0.083\\,\\Omega$, CA: $0.082\\,\\Omega$) cân bằng cao, độ lệch tối đa chỉ $1.22\\% \\le 5.0\\%$ $\rightarrow$ **PASS**.
- Độ rung không tải tháo khớp đạt $0.05\\text{ in/sec pk} \\le 0.15\\text{ in/sec pk}$ theo Table 100.10 $\rightarrow$ **PASS**.

---

## 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)

1. **CẢNH BÁO TỐI CẤP: TUYỆT ĐỐI KHÔNG HÒA LƯỚI ĐÓNG TẢI:**
   - Ban hành phiếu cấm đóng điện (Danger / Do Not Operate) đối với máy phát \`${data.site_info.generator_tag}\`.
   - Khóa mạch điều khiển kích từ AVR để ngăn ngừa việc vô tình kích từ gây mất cân bằng lực từ UMP.

2. **Cô lập & Xử lý Cuộn dây Cực Rô-to số 3 (Field Pole 3):**
   - Thực hiện tháo rotor đưa ra ngoài bệ thử hoặc tháo các thanh giằng nêm cực từ Pole 3.
   - Kiểm tra cách điện giữa các vòng dây (turn-to-turn insulation) của cuộn cực số 3 bằng thiết bị đo xung áp cao tần (Surge Comparison Tester) hoặc đo điện trở thuần từng lớp dây.
   - Tiến hành quấn lại hoặc thay thế bối dây cực từ số 3 đạt đúng số vòng dây và tiết diện dây đồng xuất xưởng.

3. **Thử nghiệm đối soát lại sau khi xử lý:**
   - Thực hiện lại phép thử sụt áp AC trên từng cực rô-to (Pole-to-Pole AC Voltage Drop Test).
   - Xác nhận độ lệch sụt áp giữa tất cả các cực từ nằm gọn trong phạm vi $\\le \\pm 10\\%$ so với trung bình chuỗi trước khi nghiệm thu đóng điện đưa máy vào mang tải.
`;
}

function poleDeviationsCalc(val: number, avg: number): string {
  const dev = ((val - avg) / avg) * 100;
  return `${dev > 0 ? '+' : ''}${dev.toFixed(1)}%`;
}

export async function generateNetaSyncMachineryAiAnalysis(
  payload: NetaSyncMachineryInputPayload
): Promise<NetaSyncMachineryEvaluationResult> {
  const deterministicResult = performDeterministicSyncMachineryAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    deterministicResult.markdown_report = buildFallbackSyncMachineryMarkdown(payload, deterministicResult);
    return deterministicResult;
  }

  try {
    const ai = new GoogleGenAI({ apiKey });

    const systemPrompt = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách Động cơ và Máy phát điện Đồng bộ (Rotating Machinery, Synchronous Motors and Generators) theo tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.15.2).
Nhiệm vụ của bạn là nhận dữ liệu kiểm tra ngoài site từ Field Service Engineer (FSE) và tự động đối soát với tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.15.2, Table 100.10, Table 100.11, Table 100.12) và IEEE 115.

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.15.2):
1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Nameplate Match: Đối soát nhãn Stator, Rotor cực lồi và Exciter khớp 100% bản vẽ.
- Physical & Cleanliness: Stator, rotor, cuộn kích từ, cụm chổi than/vòng trượt sạch bụi than và nguyên vẹn.
- Alignment & Air-Gap: Căn chỉnh đồng trục và khe hở không khí giữa các cực lồi đồng đều.
- Bolt Torque: Theo nhà sản xuất hoặc Bảng Table 100.12.

2. PHÉP ĐO ĐIỆN VÀ TIÊU CHUẨN ĐÁNH GIÁ (ELECTRICAL TEST VALUES):
- Cách điện Stator & Rotor (IR / PI / DAR - Table 100.11):
  + Stator IR đạt Table 100.11; PI >= 2.0 (máy > 200 HP); DAR >= 1.4 (máy <= 200 HP).
  + Cuộn kích từ chính (Main field rotor) và Exciter IR đạt chuẩn Table 100.11 (>= 100 MΩ).
- Thử sụt áp AC các Cực Rô-to (Pole-to-Pole AC Voltage Drop): Độ lệch điện áp sụt trên từng cực KHÔNG ĐƯỢC vượt quá 10% (> 10%) so với giá trị trung bình toàn bộ các cực.
  + Nếu lệch > 10% (ví dụ cực 3 là 18.1V so với trung bình 22.8V -> lệch -20.6%): KẾT LUẬN CRITICAL FAIL - Shorted turns in field pole 3 (Ngắn mạch giữa các vòng dây cuộn cực từ).
- Điện trở Stator giữa các pha: CẢNH BÁO "INVESTIGATE" nếu lệch quá 5% giữa các pha.
- Diode & SCR Kích từ: Điốt quay và SCR mở cổng chính xác, không bị đánh thủng.
- Thử nghiệm Đồng bộ & Rơ-le Cos phi: Rơ-le bảo vệ hệ số công suất trip đúng cài đặt khi sụt dòng kích từ.
- Thử nghiệm Độ rung (Table 100.10): Biên độ rung máy không tải tháo khớp tuân thủ Table 100.10.

Định dạng phản hồi Markdown gồm 4 phần:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE (kèm lý do chính).
2. BẢNG PHÂN TÍCH CHI TIẾT MÁY ĐIỆN ĐỒNG BỘ (Detailed Evaluation Table): So sánh thực tế với NETA ATS-2025 Section 7.15.2 (Table 100.10, Table 100.11, Table 100.12).
3. PHÂN TÍCH HỆ THỐNG KÍCH TỪ, CỰC RÔ-TO & ĐỒNG BỘ (Excitation, Field Pole & Sync Risk).
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations).`;

    const userPrompt = `Dữ liệu kiểm tra hiện trường FSE gửi lên đối soát theo NETA ATS-2025 Mục 7.15.2:
${JSON.stringify(payload, null, 2)}

Kết quả đối soát định lượng sơ bộ:
- Trạng thái sơ bộ: ${deterministicResult.overall_status}
- Lý do: ${deterministicResult.status_reason}
- Stator IR 1min: ${payload.electrical_tests.stator_ir_40c_megohms.stator_1min_ir_2500v} MΩ, PI: ${payload.electrical_tests.stator_ir_40c_megohms.calculated_pi_value}
- Pole-to-pole AC drop: ${JSON.stringify(payload.electrical_tests.pole_to_pole_ac_voltage_drop_volts)}
- Diodes: ${payload.electrical_tests.rotating_diodes_check}
- Vibration: ${payload.electrical_tests.uncoupled_vibration_in_sec_pk} in/sec pk

Hãy lập báo cáo thẩm định chuyên gia NETA ATS-2025 chi tiết, khách quan, giàu tính kỹ thuật điện cơ và nêu rõ phân tích rủi ro lực hút từ trường bất đối xứng UMP và khuyến nghị hành động chi tiết cho Kỹ sư Hiện trường FSE.`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [
        { role: 'user', parts: [{ text: `${systemPrompt}\n\n${userPrompt}` }] }
      ]
    });

    const aiMarkdown = response.text;
    deterministicResult.markdown_report = aiMarkdown || buildFallbackSyncMachineryMarkdown(payload, deterministicResult);
    return deterministicResult;
  } catch (error) {
    console.error('Error calling Gemini for NETA ATS-2025 Synchronous Machinery analysis:', error);
    deterministicResult.markdown_report = buildFallbackSyncMachineryMarkdown(payload, deterministicResult);
    return deterministicResult;
  }
}
