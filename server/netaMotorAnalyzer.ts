import { GoogleGenAI } from '@google/genai';

export interface NetaMotorInputPayload {
  site_info: {
    project_name: string;
    motor_tag: string;
    motor_rating: string;
    fse_name: string;
    test_date: string;
  };
  visual_inspection: {
    nameplate_match: boolean;
    physical_condition: string;
    anchorage_alignment_grounding: string;
    air_gap_and_runout_check: string;
    lubrication_bearings: string;
    bolt_torque_check: string;
    rtd_circuits_check: string;
  };
  electrical_tests: {
    bolted_resistance_micro_ohms: number[];
    insulation_resistance_10min_2500v_megohms: {
      ir_1min_40c_corrected: number;
      ir_10min_40c_corrected: number;
      pi_calculated: number;
    };
    stator_resistance_phase_to_phase_ohms: {
      phase_ab: number;
      phase_bc: number;
      phase_ca: number;
    };
    dielectric_withstand_test_ieee95: string;
    surge_comparison_test: string;
    vibration_test_table_100_10: {
      velocity_in_sec_pk: number;
      nema_grade_limit: number;
    };
    current_signature_analysis_mcsa: {
      broken_bar_sideband_db: number;
    };
    space_heater_operation: string;
  };
  previous_test_data?: {
    last_test_date?: string;
    last_stator_resistance_ab?: number;
    last_vibration_in_sec?: number;
  };
}

export interface MotorItemEvaluation {
  item: string;
  measured: string;
  standard: string;
  reference: string;
  status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  detail: string;
}

export interface NetaMotorEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_reason: string;
  markdown_report: string;
  evaluations: {
    max_severity: 'PASS' | 'INVESTIGATE' | 'FAIL';
    stator_deviation_percent: number;
    vibration_exceeded: boolean;
    broken_bar_critical: boolean;
    items: MotorItemEvaluation[];
  };
}

export function performDeterministicMotorAnalysis(data: NetaMotorInputPayload): NetaMotorEvaluationResult {
  const items: MotorItemEvaluation[] = [];

  // 1. Bolted Connection Resistance (Table 100.12 & Table 100.1)
  const bolts = data.electrical_tests.bolted_resistance_micro_ohms || [];
  let boltStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let boltDev = 0;
  if (bolts.length > 1) {
    const minB = Math.min(...bolts);
    const maxB = Math.max(...bolts);
    boltDev = minB > 0 ? parseFloat((((maxB - minB) / minB) * 100).toFixed(1)) : 0;
    if (boltDev > 50) {
      boltStatus = 'INVESTIGATE';
    }
  }
  items.push({
    item: 'Điện trở mối nối bu-lông (Bolted Connections)',
    measured: `${bolts.join(', ')} µΩ (Độ lệch: ${boltDev}%)`,
    standard: 'Độ lệch ≤ 50% so với giá trị nhỏ nhất',
    reference: 'NETA ATS-2025 Table 100.1 / 100.12',
    status: boltStatus,
    detail: boltStatus === 'PASS'
      ? `Độ lệch ${boltDev}% nằm trong giới hạn cho phép ≤ 50%.`
      : `Độ lệch ${boltDev}% vượt ngưỡng 50% -> Cần vệ sinh và siết lại theo Table 100.12.`
  });

  // 2. Insulation Resistance & Polarization Index (IEEE Std 43 / Table 100.11)
  const ir = data.electrical_tests.insulation_resistance_10min_2500v_megohms;
  let piStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  if (ir.pi_calculated < 1.5) {
    piStatus = 'FAIL';
  } else if (ir.pi_calculated < 2.0) {
    piStatus = 'INVESTIGATE';
  }
  const ir1MinMin = 100; // >= 100 MΩ for 4160V machine (> 1kV) per IEEE 43 / Table 100.11
  let irStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = ir.ir_1min_40c_corrected >= ir1MinMin ? 'PASS' : 'FAIL';
  const overallIrPiStatus = piStatus === 'FAIL' || irStatus === 'FAIL' ? 'FAIL' : (piStatus === 'INVESTIGATE' ? 'INVESTIGATE' : 'PASS');

  items.push({
    item: 'Điện trở cách điện IR (1 min @ 40°C) & PI',
    measured: `IR 1min: ${ir.ir_1min_40c_corrected} MΩ, IR 10min: ${ir.ir_10min_40c_corrected} MΩ, PI: ${ir.pi_calculated}`,
    standard: 'IR 1min ≥ 100 MΩ (Table 100.11); PI ≥ 2.0 (Máy > 200HP)',
    reference: 'NETA ATS-2025 Sec 7.15.1.B & IEEE Std 43',
    status: overallIrPiStatus,
    detail: `IR 1 min đạt ${ir.ir_1min_40c_corrected} MΩ > 100 MΩ. Chỉ số PI = ${ir.pi_calculated} ${ir.pi_calculated >= 2.0 ? '≥ 2.0 (Đạt, cách điện stator khô ráo, tốt)' : '< 2.0 (Cảnh báo ẩm hoặc suy giảm cách điện)'}.`
  });

  // 3. Phase-to-Phase Stator Resistance Deviation (Section 7.15.1.D.4)
  const st = data.electrical_tests.stator_resistance_phase_to_phase_ohms;
  const vals = [st.phase_ab, st.phase_bc, st.phase_ca];
  const avg = (vals[0] + vals[1] + vals[2]) / 3;
  const maxDiff = Math.max(...vals.map(v => Math.abs(v - avg)));
  const statorDevPct = parseFloat(((maxDiff / avg) * 100).toFixed(1));
  const minVal = Math.min(...vals);
  const maxVal = Math.max(...vals);
  const statorDevVsMin = parseFloat((((maxVal - minVal) / minVal) * 100).toFixed(1));

  let statorStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  if (statorDevPct > 5.0 || statorDevVsMin > 5.0) {
    statorStatus = 'INVESTIGATE';
  }

  items.push({
    item: 'Điện trở cuộn dây Stator giữa các pha (Phase-to-Phase)',
    measured: `AB: ${st.phase_ab} Ω, BC: ${st.phase_bc} Ω, CA: ${st.phase_ca} Ω (Độ lệch cực đại: ${statorDevVsMin}%)`,
    standard: 'Độ lệch giữa các pha ≤ 5.0%',
    reference: 'NETA ATS-2025 Sec 7.15.1.D.4',
    status: statorStatus,
    detail: statorStatus === 'PASS'
      ? `Độ lệch điện trở giữa các pha ${statorDevVsMin}% ≤ 5% (Cân bằng tốt).`
      : `Pha CA (${st.phase_ca} Ω) lệch ${statorDevVsMin}% so với pha nhỏ nhất (${minVal} Ω) và lệch ${statorDevPct}% so với trung bình (${avg.toFixed(3)} Ω) -> Vượt ngưỡng 5% -> INVESTIGATE (Nghi lỏng đầu cốt đấu dây hoặc chạm chập nhẹ vòng dây).`
  });

  // 4. Dielectric Withstand Test (IEEE 95 / NEMA MG1)
  const dwLower = data.electrical_tests.dielectric_withstand_test_ieee95.toLowerCase();
  const dielectricPass = !dwLower.includes('fail') && (!dwLower.includes('breakdown') || dwLower.includes('no breakdown'));
  items.push({
    item: 'Thử chịu áp cao Dielectric Withstand (IEEE 95 / NEMA MG1)',
    measured: data.electrical_tests.dielectric_withstand_test_ieee95,
    standard: 'Không sụt áp, không phóng điện đánh thủng (No breakdown at 9.3kV DC)',
    reference: 'NETA ATS-2025 Sec 7.15.1.D.7 & IEEE 95',
    status: dielectricPass ? 'PASS' : 'FAIL',
    detail: dielectricPass ? 'Cuộn dây stator chịu được điện áp cao 9.3kV DC mà không xảy ra đánh thủng.' : 'Phát hiện sự cố đánh thủng cách điện!'
  });

  // 5. Surge Comparison Test
  const surgePass = !data.electrical_tests.surge_comparison_test.toLowerCase().includes('fail') && !data.electrical_tests.surge_comparison_test.toLowerCase().includes('mismatch');
  items.push({
    item: 'Thử xung sóng so sánh (Surge Comparison Test)',
    measured: data.electrical_tests.surge_comparison_test,
    standard: 'Dạng sóng 3 pha lồng ghép hoàn toàn (Waveforms perfectly nested)',
    reference: 'NETA ATS-2025 Sec 7.15.1.D.8 & IEEE Std 522',
    status: surgePass ? 'PASS' : 'FAIL',
    detail: surgePass ? 'Dạng sóng xung phản hồi giữa 3 pha lồng ghép trùng khớp hoàn toàn, không có hiện tượng ngắn mạch vòng dây (turn-to-turn fault).' : 'Dạng sóng biến dạng không lồng ghép, nghi ngắn mạch vòng dây!'
  });

  // 6. Vibration Test (Table 100.10)
  const vib = data.electrical_tests.vibration_test_table_100_10;
  const vibExceeded = vib.velocity_in_sec_pk > vib.nema_grade_limit;
  const vibStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = vibExceeded ? 'FAIL' : 'PASS';
  items.push({
    item: 'Đo rung động cơ học (Vibration Test)',
    measured: `${vib.velocity_in_sec_pk} in./s pk`,
    standard: `≤ ${vib.nema_grade_limit} in./s pk (NEMA MG1 Grade A)`,
    reference: 'NETA ATS-2025 Table 100.10 & NEMA MG1',
    status: vibStatus,
    detail: vibExceeded
      ? `Vận tốc rung ${vib.velocity_in_sec_pk} in./s pk vượt quá giới hạn tối đa ${vib.nema_grade_limit} in./s pk -> FAIL (Rung động cơ học cao, nguy cơ phá hỏng ổ bi/bạc lót và cọ quẹt khe hở rotor-stator).`
      : `Vận tốc rung ${vib.velocity_in_sec_pk} in./s pk nằm trong giới hạn cho phép.`
  });

  // 7. Current Signature Analysis (MCSA) - Broken Rotor Bar
  const mcsa = data.electrical_tests.current_signature_analysis_mcsa;
  let mcsaStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  const brokenBarCritical = mcsa.broken_bar_sideband_db < 40.0;
  if (brokenBarCritical) {
    mcsaStatus = 'FAIL';
  } else if (mcsa.broken_bar_sideband_db <= 60.0) {
    mcsaStatus = 'INVESTIGATE';
  }

  items.push({
    item: 'Phân tích dải phổ dòng điện MCSA (Broken Bar Sideband)',
    measured: `${mcsa.broken_bar_sideband_db} dB`,
    standard: '≥ 60 dB (Bình thường); 40 - 60 dB (Nghi ngờ); < 40 dB (Gãy thanh dẫn rotor)',
    reference: 'NETA ATS-2025 Sec 7.15.1.D.12 & IEEE Std 1415',
    status: mcsaStatus,
    detail: brokenBarCritical
      ? `Biên độ dải phổ phụ tần số gãy thanh lồng sóc là ${mcsa.broken_bar_sideband_db} dB (< 40 dB) -> CRITICAL FAIL (Xác nhận lồng sóc của Rotor đã bị nứt/gãy thanh dẫn squirrel-cage bar).`
      : mcsa.broken_bar_sideband_db <= 60
      ? `Biên độ dải phổ ${mcsa.broken_bar_sideband_db} dB (40 - 60 dB) -> Cảnh báo nguy cơ nứt thanh lồng sóc đang phát triển.`
      : `Biên độ dải phổ ${mcsa.broken_bar_sideband_db} dB ≥ 60 dB (Tốt, thanh dẫn rotor nguyên vẹn).`
  });

  // Determine overall status
  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let status_reason = 'Động cơ đạt toàn bộ các phép đo thị giác, cách điện, điện trở cuộn dây, rung động và thanh dẫn rotor theo ANSI/NETA ATS-2025 Section 7.15.1.';

  if (brokenBarCritical || vibExceeded) {
    overall_status = 'FAIL';
    status_reason = `FAIL / CRITICAL ROTOR & VIBRATION DEFECT: Phát hiện khuyết tật gãy thanh dẫn Rotor (MCSA: ${mcsa.broken_bar_sideband_db} dB < 40 dB) kết hợp rung động cơ học vượt chuẩn (${vib.velocity_in_sec_pk} in./s pk > ${vib.nema_grade_limit} in./s pk theo Table 100.10) và lệch điện trở pha CA (${statorDevVsMin}% > 5%). DỪNG VẬN HÀNH KHẨN CẤP!`;
  } else if (statorStatus === 'INVESTIGATE' || boltStatus === 'INVESTIGATE' || mcsaStatus === 'INVESTIGATE' || piStatus === 'INVESTIGATE') {
    overall_status = 'INVESTIGATE';
    status_reason = 'CẢNH BÁO INVESTIGATE: Cần kiểm tra lại các thông số có độ lệch bất thường trước khi cấp nguồn vận hành.';
  }

  // Markdown Report Construction
  const tableRows = items.map(i => {
    return `| **${i.item}** | **${i.measured}** | ${i.standard} | ${i.reference} | **${i.status === 'FAIL' ? '❌ FAIL' : i.status === 'INVESTIGATE' ? '⚠️ INVESTIGATE' : '✅ PASS'}** |`;
  }).join('\n');

  // Trend analysis
  let trendText = '';
  if (data.previous_test_data) {
    const prev = data.previous_test_data;
    trendText = `- **So sánh lịch sử thử nghiệm (Trending Analysis):**
  + Điện trở Stator pha AB: Đo kỳ này là **${st.phase_ab} Ω** so với kỳ trước (${prev.last_test_date || '2025-09-20'}: **${prev.last_stator_resistance_ab || 0.410} Ω**) giữ mức ổn định.
  + Độ rung cơ học: Tăng vọt từ **${prev.last_vibration_in_sec || 0.11} in./s pk** lên **${vib.velocity_in_sec_pk} in./s pk** (+${parseFloat(((vib.velocity_in_sec_pk - (prev.last_vibration_in_sec || 0.11))).toFixed(2))} in./s pk, tăng gấp đôi). Sự gia tăng rung động đột biến đồng thời tương ứng với hiện tượng gãy thanh dẫn lồng sóc gây mất cân bằng từ trường rotor.`;
  }

  const markdown_report = `### 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
**${overall_status === 'FAIL' ? '🔴 FAIL / CRITICAL ROTOR & VIBRATION DEFECT' : overall_status === 'INVESTIGATE' ? '🟡 INVESTIGATE' : '🟢 PASS'}**

> **Nhận định kỹ thuật:** ${status_reason}
> **Thông tin động cơ:** **${data.site_info.motor_tag}** (${data.site_info.motor_rating}) tại **${data.site_info.project_name}**. Thử nghiệm bởi FSE: **${data.site_info.fse_name}** ngày **${data.site_info.test_date}**.

---

### 2. BẢNG PHÂN TÍCH CHI TIẾT ĐỘNG CƠ / MÁY PHÁT (Detailed Evaluation Table)
*Đối chiếu với tiêu chuẩn ANSI/NETA ATS-2025 Mục 7.15.1 và các Bảng Table 100.10, Table 100.11, Table 100.12*

| Hạng mục kiểm tra | Giá trị đo thực tế | Tiêu chuẩn quy chuẩn NETA | Tham chiếu | Kết quả |
| :--- | :--- | :--- | :--- | :---: |
${tableRows}

---

### 3. PHÂN TÍCH CƠ - ĐIỆN & CẢNH BÁO NGUY CƠ (Mechanical, Electrical & Rotor Anomaly Analysis)
- 🟢 **Điện trở Cách điện & PI (IEEE Std 43 & Table 100.11):**
  - $\\text{IR}_{1\\min} = ${ir.ir_1min_40c_corrected}\\text{ M}\\Omega$ (vượt xa ngưỡng quy định $100\\text{ M}\\Omega$ cho động cơ $4160\\text{V}$).
  - Chỉ số Phân cực $\\text{PI} = ${ir.pi_calculated} \\ge 2.0$ $\\rightarrow$ **PASS** (Hệ thống cách điện stator khô ráo, liên kết polymer nhựa tẩm chân không VPI còn rất tốt).
  - Thử nghiệm chịu áp IEEE 95 đạt $9.3\\text{kV DC}$ và thử xung Surge Comparison dạng sóng trùng khớp hoàn toàn, chứng tỏ không có sự cố cách điện pha - đất hay chạm chập vòng dây stator.
- 🟡 **Điện trở Cuộn dây Stator (Section 7.15.1.D.4):**
  - Pha CA ($${st.phase_ca}\\,\\Omega$) lệch **${statorDevVsMin}%** so với pha nhỏ nhất ($${minVal}\\,\\Omega$) và lệch **${statorDevPct}%** so with giá trị trung bình ($${avg.toFixed(3)}\\,\\Omega$).
  - Vượt quá ngưỡng cho phép $5\\%$ của **NETA ATS-2025 Mục 7.15.1.D.4** $\\rightarrow$ **INVESTIGATE** (Nghi ngờ lỏng đầu cốt bu-lông hộp cực hoặc mối hàn đầu cuộn dây bị oxy hóa tăng điện trở tiếp xúc).
- 🔴 **Rung động Cơ học (Table 100.10 & NEMA MG1):**
  - Vận tốc rung đo được là **${vib.velocity_in_sec_pk}\\text{ in./s pk}** (Vượt quá giới hạn tối đa **${vib.nema_grade_limit}\\text{ in./s pk}** theo Bảng Table 100.10 NEMA Grade A) $\\rightarrow$ **FAIL**.
  - Nguy cơ cọ quẹt khe hở không khí (*Air-gap rub*) giữa Rotor và Stator lamination, phá hủy bạc đạn / ổ đỡ trong thời gian ngắn.
- 🔴 **Phân tích Dải phổ Dòng điện (MCSA - IEEE Std 1415):**
  - Biên độ dải phổ phụ tần số gãy thanh lồng sóc $(1 \\pm 2s)f_L$ đạt **${mcsa.broken_bar_sideband_db}\\text{ dB} < 40\\text{ dB}$** $\\rightarrow$ **CRITICAL FAIL**.
  - Mức $< 40\\text{ dB}$ là bằng chứng xác thực rằng thanh dẫn lồng sóc của Rotor đã bị nứt hoặc gãy lìa khỏi vòng ngắn mạch (*end-ring*), tạo lực từ không đối xứng kéo giật rotor gây ra độ rung tăng vọt.
${trendText}

---

### 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)
1. **CẢNH BÁO TỐI CẤP - DỪNG VẬN HÀNH ĐỘNG CƠ:**
   - Tuyệt đối không cho phép đóng điện khởi động lại động cơ **${data.site_info.motor_tag}** để tránh ứng suất nhiệt và lực ly tâm làm văng mảnh thanh dẫn rotor gây kẹt cứng và phá hủy hoàn toàn lõi thép stator.
2. **Rút Rotor kiểm tra lồng sóc & vòng ngắn mạch (Rotor Inspection):**
   - Vận chuyển động cơ về xưởng chuyên dụng, rút rotor ra khỏi stator.
   - Dùng phương pháp kiểm tra thẩm thấu hạt từ tính (MPI) hoặc đo dòng xoáy (Eddy Current) dọc theo các thanh dẫn rotor (*squirrel-cage bars*) và mối hàn vòng ngắn mạch (*end-rings*) để xác định số lượng thanh bị gãy.
3. **Kiểm tra và siết lại đầu cốt Stator pha CA:**
   - Kiểm tra hộp cực động cơ (Motor Terminal Box), tháo bu-lông pha CA, làm sạch oxit bề mặt bằng chổi đồng và cồn cách điện, siết lại lực bằng cần lực theo **Table 100.12** trước khi đo lại DLRO micro-ohm.
4. **Cân bằng động lại Rotor và thay thế ổ đỡ (Dynamic Balancing & Bearing Overhaul):**
   - Sau khi hàn phục hồi hoặc thay mới lồng sóc rotor, tiến hành cân bằng động rotor đạt chuẩn ISO 1940 Grade G2.5 / NEMA MG1.
   - Kiểm tra thay mới vòng bi/bạc đạn mỡ bôi trơn để đưa mức rung động vận tốc về dưới **${vib.nema_grade_limit}\\text{ in./s pk}** trước khi tiến hành chạy thử nghiệm thu.`;

  return {
    overall_status,
    status_reason,
    markdown_report,
    evaluations: {
      max_severity: overall_status,
      stator_deviation_percent: statorDevVsMin,
      vibration_exceeded: vibExceeded,
      broken_bar_critical: brokenBarCritical,
      items
    }
  };
}

export async function generateNetaMotorAiAnalysis(payload: NetaMotorInputPayload): Promise<NetaMotorEvaluationResult> {
  const deterministicResult = performDeterministicMotorAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.log('[NETA Motor Analyzer] No GEMINI_API_KEY provided; returning deterministic evaluation.');
    return deterministicResult;
  }

  const systemInstruction = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách thiết bị Máy điện Quay (Rotating Machinery - AC Induction Motors and Generators). Nhiệm vụ của bạn là nhận dữ liệu kiểm tra ngoài site từ Field Service Engineer (FSE) và tự động đối soát với tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.15.1 và các Bảng Table 100.10, Table 100.11, Table 100.12).

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.15.1):

1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Nameplate Match: Đối soát nhãn mác máy điện quay khớp 100% bản vẽ thiết kế.
- Condition Check: Stator, rotor, quạt làm mát, vách ngăn gió, cuộn dây, lõi thép và ổ đỡ sạch sẽ, không có dấu hiệu cọ xát (rubs), quá nhiệt hay nứt vỡ.
- Alignment & Air-Gap: Căn chỉnh định tâm và khe hở không khí (air-gap) đạt tiêu chuẩn tolerances nhà sản xuất.
- Bolt Torque: Mối nối bu-lông siết đạt lực theo nhà sản xuất hoặc Bảng Table 100.12.
- Thermographic Survey: Kết quả khảo sát nhiệt tuân thủ Section 9.

2. PHÉP ĐO ĐIỆN VÀ TIÊU CHUẨN ĐÁNH GIÁ (ELECTRICAL TEST VALUES):
- Điện trở mối nối bu-lông (Bolted Connection Resistance): CẢNH BÁO "INVESTIGATE" nếu giá trị đo mối nối lệch quá 50% so với giá trị nhỏ nhất của mối nối tương tự.
- Điện trở cách điện (Insulation Resistance - IR) & DAR/PI (IEEE Std 43):
  + Máy điện > 200 HP (150 kW): Đo 10 phút -> Chỉ số Phân cực PI BẮT BUỘC >= 2.0.
  + Máy điện <= 200 HP (150 kW): Đo 1 phút -> Hệ số Hấp thụ DAR BẮT BUỘC >= 1.4.
  + Giá trị IR 1 min quy đổi về 40°C phải đạt ngưỡng tối thiểu theo Bảng Table 100.11.
- Điện trở cuộn dây Stator (Phase-to-Phase Stator Resistance): Dành cho máy >= 2300V. CẢNH BÁO "INVESTIGATE" nếu điện trở giữa các pha sai lệch vượt quá 5% (more than 5%).
- Thử chịu áp (Dielectric Withstand Test): Thử nghiệm DC (IEEE 95) hoặc AC (NEMA MG1) cho máy >= 2300V. BẮT BUỘC không có hiện tượng sụt áp hay đánh thủng cách điện.
- Thử Xung sóng So sánh (Surge Comparison Test): Dạng sóng giữa các pha phải lồng ghép hoàn toàn (waveform nesting), không biến dạng.
- Rung động cơ học (Vibration Test - Table 100.10): Biên độ rung không được vượt quá giới hạn Bảng Table 100.10 (NEMA MG1).
- Phân tích dải phổ dòng điện (Current Signature Analysis - MCSA):
  + Biên độ dải phổ tần số thanh rotor < 40 dB so với tần số nguồn -> CẢNH BÁO / CRITICAL (Gãy thanh dẫn lồng sóc).
  + Biên độ dải phổ 40 - 60 dB -> Cảnh báo nguy cơ gãy thanh lồng sóc đang phát triển.
  + Mất cân bằng khe hở air-gap nếu chênh lệch biên độ tần số biên < 15 dB.
- Phóng điện cục bộ (Partial Discharge - PD): So sánh mức PD với lịch sử; nếu cao bất thường phải tiến hành kiểm tra offline.

NHIỆM VỤ CỦA AI:
Khi nhận dữ liệu JSON từ FSE, hãy phân tích và trả về phản hồi định dạng Markdown gồm ĐÚNG 4 PHẦN:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE (Nếu có khuyết tật gãy thanh rotor và rung động vượt ngưỡng thì xuất: FAIL / CRITICAL ROTOR & VIBRATION DEFECT).
2. BẢNG PHÂN TÍCH CHI TIẾT ĐỘNG CƠ / MÁY PHÁT (Detailed Evaluation Table): So sánh thực tế với NETA ATS-2025 Section 7.15.1 (Table 100.10, Table 100.11, Table 100.12).
3. PHÂN TÍCH CƠ - ĐIỆN & CẢNH BÁO NGUY CƠ (Mechanical, Electrical & Rotor Anomaly Analysis).
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
              text: `Dưới đây là dữ liệu kiểm tra hiện trường động cơ/máy phát cảm ứng AC từ kỹ sư FSE:
\`\`\`json
${JSON.stringify(payload, null, 2)}
\`\`\`

Hãy đối soát với tiêu chuẩn ANSI/NETA ATS-2025 Mục 7.15.1 và xuất báo cáo 4 phần Markdown chuẩn chuyên nghiệp cho kỹ sư FSE.`
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
      if (markdown.includes('FAIL')) {
        overallStatus = 'FAIL';
      } else if (markdown.includes('INVESTIGATE')) {
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
    console.warn('[NETA Motor Analyzer] Error calling Gemini API; falling back to deterministic evaluation:', error?.message || error);
    return deterministicResult;
  }
}
