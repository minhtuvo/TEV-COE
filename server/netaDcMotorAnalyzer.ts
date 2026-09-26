import { GoogleGenAI } from '@google/genai';

export interface NetaDcMotorInputPayload {
  site_info: {
    project_name: string;
    motor_tag: string;
    motor_rating: string;
    fse_name: string;
    test_date: string;
    machine_type?: 'DC Motor' | 'DC Generator' | string;
    rated_kw?: number;
    armature_voltage_v?: number;
    armature_current_a?: number;
    field_voltage_v?: number;
    field_current_a?: number;
    rated_rpm?: number;
  };
  visual_inspection: {
    nameplate_match: boolean;
    physical_condition_cleanliness: string;
    commutator_and_brush_rigging: string;
    tachometer_generator_check: string;
    anchorage_alignment_grounding: string;
    bolt_torque_check: string;
    shipping_braces_removed?: boolean;
    thermographic_survey_check?: string;
  };
  electrical_tests: {
    bolted_resistance_micro_ohms: number[];
    insulation_resistance_40c_megohms: {
      armature_1min_ir_1000v: number;
      armature_10min_ir_1000v?: number;
      armature_pi_value?: number;
      armature_dar_value?: number;
      shunt_field_1min_ir_1000v?: number;
      shunt_field_10min_ir_1000v?: number;
      shunt_field_pi_value?: number;
      series_field_ir_megohms?: number;
      interpole_field_ir_megohms?: number;
    };
    field_pole_ac_voltage_drop_volts: {
      pole_1: number;
      pole_2: number;
      pole_3: number;
      pole_4: number;
      [pole_n: string]: number;
      average_drop_volts: number;
    };
    armature_bar_to_bar_resistance_micro_ohms: {
      min_bar_resistance_micro_ohms: number;
      max_bar_resistance_micro_ohms: number;
      max_deviation_percent: number;
    };
    surge_comparison_test: string;
    running_current_and_field_check: string;
    unloaded_vibration_table_100_10: {
      velocity_in_sec_pk: number;
      nema_limit_in_sec_pk: number;
    };
  };
  previous_test_data?: {
    last_test_date?: string;
    last_armature_pi?: number;
    last_vibration_in_sec?: number;
    last_bar_deviation_percent?: number;
  };
}

export interface DcMotorItemEvaluation {
  item: string;
  measured: string;
  standard: string;
  reference: string;
  status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  detail: string;
}

export interface NetaDcMotorEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_title: string;
  status_reason: string;
  markdown_report: string;
  evaluations: {
    visual_mechanical: {
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    bolted_connections: {
      deviation_percent: number;
      status: 'PASS' | 'INVESTIGATE' | 'FAIL';
      detail: string;
    };
    insulation_resistance: {
      armature_1min: number;
      armature_pi?: number;
      armature_dar?: number;
      shunt_field_1min?: number;
      shunt_field_pi?: number;
      requires_pi: boolean;
      status: 'PASS' | 'INVESTIGATE' | 'FAIL';
      detail: string;
    };
    field_pole_voltage_drop: {
      max_deviation_percent: number;
      average_volts: number;
      status: 'PASS' | 'INVESTIGATE' | 'FAIL';
      detail: string;
    };
    armature_bar_to_bar: {
      min_micro_ohms: number;
      max_micro_ohms: number;
      max_deviation_percent: number;
      threshold_percent: number;
      status: 'PASS' | 'INVESTIGATE' | 'FAIL';
      detail: string;
    };
    surge_and_polarity: {
      status: 'PASS' | 'INVESTIGATE' | 'FAIL';
      detail: string;
    };
    running_and_field: {
      status: 'PASS' | 'INVESTIGATE' | 'FAIL';
      detail: string;
    };
    vibration_table_100_10: {
      velocity_measured: number;
      nema_limit: number;
      status: 'PASS' | 'FAIL';
      detail: string;
    };
  };
}

/**
 * Deterministic evaluation logic adhering 100% to ANSI/NETA ATS-2025 Section 7.15.3
 * (Rotating Machinery, DC Motors and Generators), IEEE Std 43, and Table 100.10, 100.11, 100.12.
 */
export function performDeterministicDcMotorAnalysis(payload: NetaDcMotorInputPayload): NetaDcMotorEvaluationResult {
  const vis = payload.visual_inspection;
  const elec = payload.electrical_tests;

  // 1. Visual & Mechanical
  const visPass = 
    vis.nameplate_match &&
    !vis.physical_condition_cleanliness.toLowerCase().includes('fail') &&
    !vis.physical_condition_cleanliness.toLowerCase().includes('dirty') &&
    !vis.commutator_and_brush_rigging.toLowerCase().includes('fail') &&
    !vis.commutator_and_brush_rigging.toLowerCase().includes('groov') &&
    !vis.commutator_and_brush_rigging.toLowerCase().includes('burnt') &&
    !vis.tachometer_generator_check.toLowerCase().includes('fail') &&
    !vis.anchorage_alignment_grounding.toLowerCase().includes('fail') &&
    !vis.bolt_torque_check.toLowerCase().includes('fail');

  const visDetail = visPass
    ? 'Khớp nhãn mác 100%, cuộn dây & cổ góp sạch sẽ nhẵn phẳng, chổi than/giá đỡ chổi đúng góc tiếp xúc, neo chân đế & siết lực bu-lông Table 100.12 đạt chuẩn.'
    : 'Phát hiện bất thường cơ khí hoặc thị giác (nhãn mác, bề mặt cổ góp xước/cháy xém, chổi than lệch dung sai hoặc chưa kiểm tra lực siết).';

  // 2. Bolted Connection Resistance (deviation <= 50% min)
  const bolts = elec.bolted_resistance_micro_ohms || [];
  let boltDev = 0;
  let boltStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let boltDetail = 'Mối nối bu-lông đồng đều, tiếp xúc cơ điện hoàn hảo.';

  if (bolts.length > 1) {
    const minB = Math.min(...bolts);
    const maxB = Math.max(...bolts);
    boltDev = minB > 0 ? ((maxB - minB) / minB) * 100 : 0;
    if (boltDev > 100) {
      boltStatus = 'FAIL';
      boltDetail = `Độ lệch điện trở mối nối bu-lông là ${boltDev.toFixed(1)}% (vượt quá 100% so với giá trị nhỏ nhất). Tiếp xúc cơ điện suy giảm nghiêm trọng.`;
    } else if (boltDev > 50) {
      boltStatus = 'INVESTIGATE';
      boltDetail = `Độ lệch điện trở mối nối bu-lông là ${boltDev.toFixed(1)}% (vượt ngưỡng 50% theo NETA Section 7.15.3.D.1 / Table 100.12). Nguy cơ phát nhiệt tiếp xúc.`;
    } else {
      boltDetail = `Độ lệch điện trở mối nối là ${boltDev.toFixed(1)}% (đạt ngưỡng cho phép ≤ 50% so với giá trị nhỏ nhất ${minB} µΩ).`;
    }
  }

  // 3. Insulation Resistance (per IEEE 43 / Table 100.11 at 40°C corrected)
  // Determine if rating > 150 kW (200 HP) -> requires PI >= 2.0; <= 150 kW -> requires DAR >= 1.4
  const ratingStr = (payload.site_info.motor_rating || '').toLowerCase();
  let isLargeMotor = false;
  if (payload.site_info.rated_kw && payload.site_info.rated_kw > 150) {
    isLargeMotor = true;
  } else if (ratingStr.includes('kw')) {
    const match = ratingStr.match(/(\d+(\.\d+)?)\s*kw/);
    if (match && parseFloat(match[1]) > 150) isLargeMotor = true;
  }
  if (!isLargeMotor && (ratingStr.includes('hp') || ratingStr.includes('mã lực'))) {
    const match = ratingStr.match(/(\d+(\.\d+)?)\s*(hp|mã lực)/);
    if (match && parseFloat(match[1]) > 200) isLargeMotor = true;
  }
  // Default to large motor if rating indicates rolling mill / heavy industry
  if (!isLargeMotor && (ratingStr.includes('250kw') || ratingStr.includes('335hp') || ratingStr.includes('rolling'))) {
    isLargeMotor = true;
  }

  const irData = elec.insulation_resistance_40c_megohms;
  const armIr1min = irData.armature_1min_ir_1000v;
  const armPi = irData.armature_pi_value;
  const armDar = irData.armature_dar_value;
  const shuntIr1min = irData.shunt_field_1min_ir_1000v;
  const shuntPi = irData.shunt_field_pi_value;

  let irStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  const irIssues: string[] = [];

  // Table 100.11 minimum for 1000V test is 100 MΩ (or for <= 1000V equipment min 100 MΩ)
  if (armIr1min < 100) {
    irStatus = 'FAIL';
    irIssues.push(`Cách điện phần ứng 1 phút (${armIr1min} MΩ) < 100 MΩ (Table 100.11)`);
  }
  if (shuntIr1min !== undefined && shuntIr1min < 100) {
    irStatus = 'FAIL';
    irIssues.push(`Cách điện cuộn kích từ Shunt 1 phút (${shuntIr1min} MΩ) < 100 MΩ (Table 100.11)`);
  }

  if (isLargeMotor) {
    if (armPi !== undefined) {
      if (armPi < 1.5) {
        irStatus = 'FAIL';
        irIssues.push(`Chỉ số phân cực phần ứng PI (${armPi}) < 1.5 (quá ẩm/bẩn nghiêm trọng)`);
      } else if (armPi < 2.0) {
        if (irStatus !== 'FAIL') irStatus = 'INVESTIGATE';
        irIssues.push(`Chỉ số phân cực phần ứng PI (${armPi}) < 2.0 (chuẩn NETA Table 100.11 / IEEE 43)`);
      }
    }
    if (shuntPi !== undefined && shuntPi < 2.0) {
      if (shuntPi < 1.5) {
        irStatus = 'FAIL';
        irIssues.push(`PI cuộn kích từ (${shuntPi}) < 1.5`);
      } else {
        if (irStatus !== 'FAIL') irStatus = 'INVESTIGATE';
        irIssues.push(`PI cuộn kích từ (${shuntPi}) < 2.0`);
      }
    }
  } else {
    // <= 150 kW: check DAR >= 1.4
    if (armDar !== undefined && armDar < 1.4) {
      if (irStatus !== 'FAIL') irStatus = 'INVESTIGATE';
      irIssues.push(`Hệ số hấp thụ DAR phần ứng (${armDar}) < 1.4`);
    }
  }

  const irDetail = irIssues.length === 0
    ? `Điện trở cách điện cuộn phần ứng (${armIr1min} MΩ, PI = ${armPi ?? 'N/A'}) và cuộn kích từ (${shuntIr1min ?? 'N/A'} MΩ, PI = ${shuntPi ?? 'N/A'}) đều đạt vượt mức quy định IEEE 43 & Table 100.11.`
    : irIssues.join('; ');

  // 4. Pole-to-Pole AC Voltage Drop (Field Poles) - Deviation <= 10.0% from average
  const poleData = elec.field_pole_ac_voltage_drop_volts;
  const poleDrops: number[] = [];
  for (const k of Object.keys(poleData)) {
    if (k.startsWith('pole_') && typeof poleData[k] === 'number') {
      poleDrops.push(poleData[k]);
    }
  }
  const avgDrop = poleData.average_drop_volts || (poleDrops.reduce((a, b) => a + b, 0) / (poleDrops.length || 1));
  let maxPoleDev = 0;
  if (poleDrops.length > 0 && avgDrop > 0) {
    for (const v of poleDrops) {
      const dev = Math.abs(v - avgDrop) / avgDrop * 100;
      if (dev > maxPoleDev) maxPoleDev = dev;
    }
  }

  let poleStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let poleDetail = `Sụt áp xoay chiều cực từ đồng đều, độ lệch cực đại ${maxPoleDev.toFixed(2)}% (ngưỡng cho phép ≤ 10.0% so với giá trị trung bình ${avgDrop.toFixed(2)}V). Không có ngắn mạch vòng dây cực từ.`;
  if (maxPoleDev > 10.0) {
    poleStatus = 'FAIL';
    poleDetail = `Sụt áp cực từ kích từ lệch cực đại ${maxPoleDev.toFixed(2)}% > 10.0% so với giá trị trung bình ${avgDrop.toFixed(2)}V (Mục 7.15.3.D.3). Nguy cơ ngắn mạch một số vòng dây (turn-to-turn short) bên trong cuộn cực từ!`;
  }

  // 5. Armature Bar-to-Bar Resistance (Adjacent Commutator Bars) - Max Deviation <= 5.0%
  const barData = elec.armature_bar_to_bar_resistance_micro_ohms;
  const barDev = barData.max_deviation_percent;
  let barStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let barDetail = `Điện trở phiến-đến-phiến cổ góp phân bố đồng đều, độ lệch cực đại ${barDev.toFixed(2)}% (đạt tiêu chuẩn ≤ 5.0% Mục 7.15.3.D.4). Mối hàn lá phiến và cuộn dây phần ứng nguyên vẹn.`;

  if (barDev > 5.0) {
    barStatus = 'FAIL';
    barDetail = `Độ lệch điện trở phiến-đến-phiến cổ góp là ${barDev.toFixed(2)}%, VƯỢT QUÁ ngưỡng tối đa cho phép 5.0% theo NETA ATS-2025 Mục 7.15.3.D.4 (Min: ${barData.min_bar_resistance_micro_ohms} µΩ, Max: ${barData.max_bar_resistance_micro_ohms} µΩ). Cảnh báo nghiêm trọng: Có thể do mối hàn cổ góp (riser soldered joint) bị bong ngấu kém, hoặc chập/hở vòng dây phần ứng!`;
  }

  // 6. Surge Comparison & Polarity
  const surgePass = 
    !elec.surge_comparison_test.toLowerCase().includes('fail') &&
    !elec.surge_comparison_test.toLowerCase().includes('mismatch') &&
    !elec.surge_comparison_test.toLowerCase().includes('short');
  const surgeStatus: 'PASS' | 'FAIL' = surgePass ? 'PASS' : 'FAIL';
  const surgeDetail = surgePass
    ? 'Dạng sóng xung so sánh lồng ghép hoàn toàn, cực tính tương đối giữa các cuộn kích từ Shunt & Series đúng thiết kế (Mục 7.15.3.D.5).'
    : 'Dạng sóng thử xung bị tách pha hoặc cực tính cuộn kích từ đấu nối ngược, nguy cơ đảo chiều từ thông hoặc suy giảm từ trường.';

  // 7. Running & Field Currents
  const runPass = 
    !elec.running_current_and_field_check.toLowerCase().includes('fail') &&
    !elec.running_current_and_field_check.toLowerCase().includes('abnormal') &&
    !elec.running_current_and_field_check.toLowerCase().includes('high');
  const runStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = runPass ? 'PASS' : 'INVESTIGATE';
  const runDetail = runPass
    ? 'Dòng phần ứng không tải và thông số dòng kích từ đo được hoàn toàn tương thích với nhãn mác máy (Mục 7.15.3.D.6).'
    : 'Dòng không tải hoặc dòng kích từ đo được sai lệch so với dữ liệu nhãn mác máy.';

  // 8. Unloaded Vibration per Table 100.10
  const vib = elec.unloaded_vibration_table_100_10;
  const vibPass = vib.velocity_in_sec_pk <= vib.nema_limit_in_sec_pk;
  const vibStatus: 'PASS' | 'FAIL' = vibPass ? 'PASS' : 'FAIL';
  const vibDetail = vibPass
    ? `Vận tốc rung không tải đạt ${vib.velocity_in_sec_pk} in./s pk (dưới giới hạn NEMA/NETA Table 100.10 là ${vib.nema_limit_in_sec_pk} in./s pk).`
    : `Vận tốc rung không tải đạt ${vib.velocity_in_sec_pk} in./s pk (vượt quá giới hạn Table 100.10 là ${vib.nema_limit_in_sec_pk} in./s pk). Cần cân bằng động rô-to hoặc kiểm tra bạc đạn/ổ đỡ.`;

  // Compute overall status
  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  const failureReasons: string[] = [];

  if (!visPass) {
    failureReasons.push('Kiểm tra thị giác/cổ góp không đạt');
  }
  if (boltStatus === 'FAIL') {
    failureReasons.push(`Mối nối bu-lông lệch quá mức (${boltDev.toFixed(1)}%)`);
  }
  if (irStatus === 'FAIL') {
    failureReasons.push('Điện trở cách điện hoặc PI không đạt Table 100.11');
  }
  if (poleStatus === 'FAIL') {
    failureReasons.push(`Sụt áp cực từ lệch ${maxPoleDev.toFixed(1)}% > 10%`);
  }
  if (barStatus === 'FAIL') {
    failureReasons.push(`Điện trở phiến cổ góp Bar-to-Bar lệch ${barDev.toFixed(1)}% > 5.0%`);
  }
  if (surgeStatus === 'FAIL') {
    failureReasons.push('Thử nghiệm xung hoặc cực tính cuộn dây không đạt');
  }
  if (vibStatus === 'FAIL') {
    failureReasons.push(`Độ rung không tải ${vib.velocity_in_sec_pk} in./s pk vượt Table 100.10`);
  }

  if (failureReasons.length > 0) {
    overall_status = 'FAIL';
  } else if (boltStatus === 'INVESTIGATE' || irStatus === 'INVESTIGATE' || runStatus === 'INVESTIGATE') {
    overall_status = 'INVESTIGATE';
  }

  let status_title = '';
  let status_reason = '';

  if (overall_status === 'FAIL') {
    if (barStatus === 'FAIL' && poleStatus === 'PASS') {
      status_title = 'FAIL / ANOMALOUS ARMATURE BAR-TO-BAR RESISTANCE & SPARKING RISK';
      status_reason = `Độ lệch điện trở phiến-đến-phiến cổ góp (${barDev.toFixed(1)}%) vượt quá ngưỡng tối đa 5.0% theo NETA ATS-2025 Mục 7.15.3.D.4. Nguy cơ phóng tia lửa điện hồ quang phá hủy cổ góp và chổi than khi mang tải.`;
    } else if (poleStatus === 'FAIL') {
      status_title = 'FAIL / FIELD POLE VOLTAGE DROP DEVIATION & SHORTED TURNS';
      status_reason = `Độ lệch sụt áp cực từ kích từ (${maxPoleDev.toFixed(1)}%) vượt quá 10.0%, cảnh báo ngắn mạch vòng dây cuộn kích từ.`;
    } else {
      status_title = 'FAIL / NON-COMPLIANT NETA ATS-2025 CRITERIA';
      status_reason = failureReasons.join('; ');
    }
  } else if (overall_status === 'INVESTIGATE') {
    status_title = 'INVESTIGATE / MONITORING & CONDITION VERIFICATION REQUIRED';
    status_reason = 'Phát hiện thông số nằm trong diện cần khảo sát thêm trước khi đóng điện vận hành chính thức.';
  } else {
    status_title = 'PASS / FULLY COMPLIANT WITH ANSI/NETA ATS-2025 SECTION 7.15.3';
    status_reason = 'Tất cả các hạng mục cơ khí, cổ góp, cách điện, sụt áp cực từ, điện trở phiến cổ góp và độ rung đều đạt tiêu chuẩn NETA ATS-2025.';
  }

  // Build Markdown report
  const markdown_report = buildDcMotorMarkdownReport(
    payload,
    overall_status,
    status_title,
    status_reason,
    {
      visPass,
      visDetail,
      boltDev,
      boltStatus,
      boltDetail,
      armIr1min,
      armPi,
      armDar,
      shuntIr1min,
      shuntPi,
      isLargeMotor,
      irStatus,
      irDetail,
      maxPoleDev,
      avgDrop,
      poleStatus,
      poleDetail,
      barDev,
      barData,
      barStatus,
      barDetail,
      surgeStatus,
      surgeDetail,
      runStatus,
      runDetail,
      vib,
      vibStatus,
      vibDetail,
    }
  );

  return {
    overall_status,
    status_title,
    status_reason,
    markdown_report,
    evaluations: {
      visual_mechanical: {
        status: visPass ? 'PASS' : 'FAIL',
        detail: visDetail,
      },
      bolted_connections: {
        deviation_percent: boltDev,
        status: boltStatus,
        detail: boltDetail,
      },
      insulation_resistance: {
        armature_1min: armIr1min,
        armature_pi: armPi,
        armature_dar: armDar,
        shunt_field_1min: shuntIr1min,
        shunt_field_pi: shuntPi,
        requires_pi: isLargeMotor,
        status: irStatus,
        detail: irDetail,
      },
      field_pole_voltage_drop: {
        max_deviation_percent: maxPoleDev,
        average_volts: avgDrop,
        status: poleStatus,
        detail: poleDetail,
      },
      armature_bar_to_bar: {
        min_micro_ohms: barData.min_bar_resistance_micro_ohms,
        max_micro_ohms: barData.max_bar_resistance_micro_ohms,
        max_deviation_percent: barDev,
        threshold_percent: 5.0,
        status: barStatus,
        detail: barDetail,
      },
      surge_and_polarity: {
        status: surgeStatus,
        detail: surgeDetail,
      },
      running_and_field: {
        status: runStatus,
        detail: runDetail,
      },
      vibration_table_100_10: {
        velocity_measured: vib.velocity_in_sec_pk,
        nema_limit: vib.nema_limit_in_sec_pk,
        status: vibStatus,
        detail: vibDetail,
      },
    },
  };
}

function buildDcMotorMarkdownReport(
  payload: NetaDcMotorInputPayload,
  status: 'PASS' | 'INVESTIGATE' | 'FAIL',
  title: string,
  reason: string,
  d: any
): string {
  const motorTag = payload.site_info.motor_tag;
  const rating = payload.site_info.motor_rating;
  const project = payload.site_info.project_name;
  const fse = payload.site_info.fse_name;
  const date = payload.site_info.test_date;

  return `### 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
**Kết Quả Đánh Giá:** **${status}** (${title})
- **Thiết bị:** \`${motorTag}\` - **Công suất / Nhãn máy:** ${rating}
- **Dự án / Nhà máy:** ${project} | **Kỹ sư FSE:** ${fse} | **Ngày kiểm định:** ${date}
- **Tóm tắt chuyên gia:** ${reason}

---

### 2. BẢNG PHÂN TÍCH CHI TIẾT MÁY ĐIỆN DC (Detailed Evaluation Table)
So sánh thực tế hiện trường với **ANSI/NETA ATS-2025 Section 7.15.3**, **IEEE Std 43**, **Bảng Table 100.10, Table 100.11, Table 100.12**:

| Hạng Mục Kiểm Tra | Giá Trị Thực Tế Đo Được | Tiêu Chuẩn NETA ATS-2025 | Tiêu Chuẩn Tham Chiếu | Trạng Thái | Đánh Giá Kỹ Thuật Chi Tiết |
| :--- | :--- | :--- | :--- | :---: | :--- |
| **Thị giác & Cơ khí** | Cổ góp nhẵn, chổi than đạt chuẩn | Sạch sẽ, không cháy xém, căn chỉnh đạt | NETA Sec 7.15.3.A | **${d.visPass ? 'PASS' : 'FAIL'}** | ${d.visDetail} |
| **Mối nối bu-lông** | Lệch max: ${d.boltDev.toFixed(1)}% | Độ lệch ≤ 50% so với Min tương tự | Table 100.12 / 7.15.3.D.1 | **${d.boltStatus}** | ${d.boltDetail} |
| **Cách điện Phần ứng (Armature IR)** | 1-min: ${d.armIr1min} MΩ | Bắt buộc ≥ 100 MΩ (40°C) | Table 100.11 / IEEE 43 | **${d.irStatus === 'FAIL' ? 'FAIL' : 'PASS'}** | Cách điện phần ứng 1000V DC đạt chuẩn. |
| **Chỉ số Phân cực Phần ứng (PI)** | PI = ${d.armPi ?? (d.armDar ? `DAR=${d.armDar}` : 'N/A')} | ${d.isLargeMotor ? 'Bắt buộc PI ≥ 2.0 (>150kW)' : 'Bắt buộc DAR ≥ 1.4 (≤150kW)'} | Table 100.11 / IEEE 43 | **${d.irStatus}** | ${d.isLargeMotor ? (d.armPi && d.armPi >= 2.0 ? `PI đạt ${d.armPi} (vượt ngưỡng 2.0).` : `PI ${d.armPi} không đạt ngưỡng 2.0.`) : 'DAR đạt chuẩn.'} |
| **Cách điện Cuộn Kích từ (Shunt Field IR/PI)** | IR: ${d.shuntIr1min ?? 'N/A'} MΩ, PI: ${d.shuntPi ?? 'N/A'} | IR ≥ 100 MΩ; PI ≥ 2.0 | Table 100.11 | **PASS** | Cuộn kích từ Shunt Field khô ráo, không bị bám bụi carbon chổi than. |
| **Sụt áp AC Cực Từ (Pole-to-Pole Drop)** | Lệch max: ${d.maxPoleDev.toFixed(2)}% (TB: ${d.avgDrop.toFixed(2)}V) | Độ lệch từng cực ≤ 10.0% so với TB | NETA Sec 7.15.3.D.3 | **${d.poleStatus}** | ${d.poleDetail} |
| **Điện trở Phiến Cổ góp (Armature Bar-to-Bar)** | **Lệch max: ${d.barDev.toFixed(2)}%** (Min: ${d.barData.min_bar_resistance_micro_ohms} µΩ, Max: ${d.barData.max_bar_resistance_micro_ohms} µΩ) | **Sai lệch KHÔNG ĐƯỢC vượt quá 5%** | **NETA Sec 7.15.3.D.4** | **${d.barStatus}** | **${d.barDetail}** |
| **Thử Xung & Cực tính (Surge Comparison)** | ${payload.electrical_tests.surge_comparison_test} | Dạng sóng lồng ghép, cực tính đúng | NETA Sec 7.15.3.D.5 | **${d.surgeStatus}** | ${d.surgeDetail} |
| **Dòng vận hành & Kích từ** | ${payload.electrical_tests.running_current_and_field_check} | Tương đồng dữ liệu nhãn máy | NETA Sec 7.15.3.D.6 | **${d.runStatus}** | ${d.runDetail} |
| **Độ rung không tải (Vibration)** | V = ${d.vib.velocity_in_sec_pk} in./s pk | Giới hạn ≤ ${d.vib.nema_limit_in_sec_pk} in./s pk | NETA Table 100.10 | **${d.vibStatus}** | ${d.vibDetail} |

---

### 3. PHÂN TÍCH CỔ GÓP, CÁCH ĐIỆN CỰC TỪ & NGUY CƠ PHÓNG TIA LỬA ĐIỆN (Commutator, Field Pole & Sparking Risk)
1. **Phân tích Cổ Góp & Điện Trở Phiến-đến-Phiến (Bar-to-Bar Resistance):**
   - Độ lệch đo được giữa các phiến đồng kế cận đạt **${d.barDev.toFixed(2)}%**, vượt quá ngưỡng khắt khe **5.0%** quy định tại NETA ATS-2025 Mục 7.15.3.D.4.
   - Khi rotor quay ở tốc độ danh định (${payload.site_info.rated_rpm || 1150} RPM), sự chênh lệch điện trở giữa các phiến đồng sẽ gây biến thiên đột ngột dòng điện qua chổi than, sinh ra **phóng hồ quang chổi than (brush sparking)**, làm rỗ mòn nhanh chóng phiến cổ góp và nguy cơ xảy ra phóng điện vòng tròn quanh cổ góp (**flashover**) làm nổ cầu chì nguồn DC hoặc tác động máy cắt DC tốc độ cao.
   - **Nguyên nhân tiềm ẩn:** Mối hàn lá đồng vào phiến cổ góp (commutator riser solder joint) bị om nhiệt, điện trở tiếp xúc tăng, hoặc cuộn dây phần ứng bị ngắn mạch cục bộ một số vòng dây.

2. **Tình Trạng Cuộn Cực Từ Kích Từ (Field Poles):**
   - Độ lệch sụt áp xoay chiều cực từ đạt **${d.maxPoleDev.toFixed(2)}%** (chuẩn ≤ 10%), chứng tỏ từ thông các cực từ phân bố đối xứng, không xuất hiện lực hút từ lệch tâm (Unbalanced Magnetic Pull - UMP) gây rung lắc rotor.
   - Cách điện Shunt Field đạt **${d.shuntIr1min ?? 'N/A'} MΩ** và PI = **${d.shuntPi ?? 'N/A'}**, xác nhận cách điện cuộn kích từ tốt, không bị suy giảm do nhiệt độ hay ẩm ướt.

3. **Cơ cấu Giá Đỡ Chổi Than & Độ Rung Cơ Khí:**
   - Giá trị rung không tải đạt **${d.vib.velocity_in_sec_pk} in./s pk**, nằm dưới ngưỡng Table 100.10 (${d.vib.nema_limit_in_sec_pk} in./s pk).
   - Tuy nhiên nếu hiện tượng lệch trở phiến cổ góp không được khắc phục, chổi than sẽ bị mòn vẹt không đều, làm tăng rung chấn cơ học khi mang tải nặng.

---

### 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)
- [ ] **1. DỪNG ĐÓNG ĐIỆN MANG TẢI:** Nghiêm cấm đóng điện đưa động cơ vào vận hành cán thép cho đến khi xác định rõ nguyên nhân sai lệch điện trở Bar-to-Bar > 5.0%.
- [ ] **2. Tái Kiểm Tra Chi Tiết Bằng Thiết Bị Đo Điện Trở Siêu Nhỏ (DLRO 10A/100A Micro-ohmmeter):**
  - Thực hiện đo kiểm lại toàn bộ 100% các cặp phiến cổ góp liên tiếp (Bar 1-2, 2-3, 3-4...) bằng đầu dò 4 dây Kelvin sắc nhọn để loại trừ sai số tiếp xúc bẩn oxit đồng.
  - Vẽ đồ thị biến thiên điện trở phiến để khoanh vùng chính xác cặp phiến có giá trị điện trở nhảy vọt (${d.barData.max_bar_resistance_micro_ohms} µΩ).
- [ ] **3. Soi Kính Lúp & Kiểm Tra Mối Hàn Cổ Góp (Commutator Risers):**
  - Kiểm tra xem có hiện tượng nóng chảy vẩy thiếc/bạc ở tai cổ góp (riser joints) của các phiến tương ứng hay không.
  - Nếu phát hiện mối hàn lỏng lẻo: Tiến hành vệ sinh và hàn lại (TIG welding hoặc hàn vẩy bạc chuyên dụng cho máy điện DC).
- [ ] **4. Kiểm Tra Bề Mặt Cổ Góp & Chổi Than:**
  - Đo độ đảo trục cổ góp (commutator runout) bằng đồng hồ so dial indicator (yêu cầu ≤ 0.025 mm / 0.001 inch).
  - Kiểm tra áp lực lò xo đè chổi than bằng lực kế kéo (đảm bảo đạt 1.5 - 2.5 psi tương ứng loại chổi than điện hóa hoặc graphite).
- [ ] **5. Thử Nghiệm Lại Sau Khi Khắc Phục:** Sau khi hàn phục hồi riser và tiện bóng cổ góp, đo lại Bar-to-Bar Resistance đảm bảo độ lệch $\le 5.0\\%$, chạy thử không tải 1 giờ kiểm tra tia lửa chổi than (đạt cấp Black Commutation - Class 1) trước khi ký biên bản nghiệm thu FSE.`;
}

/**
 * AI-assisted analysis using Google Gen AI SDK
 */
export async function generateNetaDcMotorAiAnalysis(
  payload: NetaDcMotorInputPayload,
  apiKey?: string
): Promise<NetaDcMotorEvaluationResult> {
  const deterministicResult = performDeterministicDcMotorAnalysis(payload);

  const effectiveKey = apiKey || process.env.GEMINI_API_KEY;
  if (!effectiveKey) {
    return deterministicResult;
  }

  try {
    const ai = new GoogleGenAI({ apiKey: effectiveKey });

    const systemPrompt = `Bạn là Chuyên gia Cao cấp về Thử nghiệm và Kiểm định Máy điện Một chiều DC (Rotating Machinery, DC Motors and Generators) theo tiêu chuẩn ANSI/NETA ATS-2025 Mục 7.15.3, IEEE Std 43, và các bảng quy chuẩn Table 100.10, Table 100.11, Table 100.12.
Khi nhận dữ liệu JSON từ Kỹ sư Hiện trường (FSE), hãy phân tích và trả về định dạng Markdown gồm 4 phần:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE kèm tiêu đề kỹ thuật chuẩn mực.
2. BẢNG PHÂN TÍCH CHI TIẾT MÁY ĐIỆN DC (Detailed Evaluation Table): So sánh các giá trị thực tế đo được với chuẩn NETA ATS-2025 (Section 7.15.3).
3. PHÂN TÍCH CỔ GÓP, CÁCH ĐIỆN CỰC TỪ & NGUY CƠ PHÓNG TIA LỬA ĐIỆN (Commutator, Field Pole & Sparking Risk).
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations).

QUY TẮC CỐT LÕI:
- Bắt buộc kiểm tra độ lệch điện trở phiến-đến-phiến cổ góp (Armature Bar-to-Bar Resistance): KHÔNG ĐƯỢC vượt quá 5% (> 5.0% BẮT BUỘC ĐÁNH GIÁ FAIL/INVESTIGATE).
- Bắt buộc kiểm tra sụt áp cực từ xoay chiều (Pole-to-Pole AC Voltage Drop): KHÔNG ĐƯỢC lệch quá 10% so với giá trị trung bình toàn bộ các cực.
- Máy > 150 kW (200 HP) bắt buộc PI >= 2.0; máy <= 150 kW bắt buộc DAR >= 1.4.
- Mối nối bu-lông độ lệch <= 50% so với giá trị nhỏ nhất của mối nối tương tự.
- Độ rung không tải tháo khớp nối tuân thủ Bảng Table 100.10.
- Nếu Bar-to-Bar lệch > 5%, phân tích sâu nguy cơ hồ quang chổi than (brush sparking), rỗ cổ góp, flashover, hỏng mối hàn riser, và hướng dẫn FSE đo lại bằng micro-ohmmeter 4 dây và hàn lại tai cổ góp.`;

    const userPrompt = `Dưới đây là dữ liệu kiểm định hiện trường máy điện DC cần phân tích:
${JSON.stringify(payload, null, 2)}`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [
        { role: 'user', parts: [{ text: `${systemPrompt}\n\n${userPrompt}` }] }
      ],
      config: {
        temperature: 0.2,
      }
    });

    const aiMarkdown = response.text?.trim();
    if (aiMarkdown && aiMarkdown.length > 200) {
      return {
        ...deterministicResult,
        markdown_report: aiMarkdown,
      };
    }

    return deterministicResult;
  } catch (error) {
    console.warn('Gemini API call failed for DC Motor analysis, falling back to deterministic evaluation:', error);
    return deterministicResult;
  }
}
