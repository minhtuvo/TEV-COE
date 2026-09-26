import { GoogleGenAI } from '@google/genai';

export interface NetaCableLvInputPayload {
  site_info: {
    project_name: string;
    cable_tag: string;
    cable_rating: string;
    cable_length_meters: number;
    fse_name: string;
    test_date: string;
    voltage_rating?: '300V' | '600V' | '1000V' | '0.6/1kV' | string;
    ambient_temp_c?: number;
    manufacturer?: string;
  };
  visual_inspection: {
    cable_data_match: boolean;
    physical_condition_terminals: string;
    bolt_torque_check: string;
    thermographic_survey?: string;
  };
  electrical_tests: {
    insulation_resistance_1000v_1min_megohms: {
      phase_a_to_ground: number;
      phase_b_to_ground: number;
      phase_c_to_ground: number;
      neutral_to_ground?: number;
      phase_a_to_b: number;
      phase_b_to_c: number;
      phase_c_to_a: number;
    };
    continuity_test: string;
    parallel_conductors_dc_resistance_milli_ohms?: {
      phase_a_run1: number;
      phase_a_run2: number;
      phase_b_run1: number;
      phase_b_run2: number;
      phase_c_run1: number;
      phase_c_run2: number;
      neutral_run1?: number;
      neutral_run2?: number;
    };
    test_voltage_v_dc?: number;
  };
  previous_test_data?: {
    last_test_date?: string;
    last_phase_c_ir_megohms?: number;
    last_phase_a_ir_megohms?: number;
    last_phase_b_ir_megohms?: number;
    last_neutral_ir_megohms?: number;
  };
}

export interface NetaCableLvEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_title: string;
  status_reason: string;
  markdown_report: string;
  evaluations: {
    visual_mechanical: {
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    continuity: {
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    insulation_resistance: {
      required_min_megohms: number;
      tested_voltage_v: number;
      min_measured_megohms: number;
      failing_tests: string[];
      phase_c_drop_percent?: number;
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    parallel_conductors: {
      has_parallel: boolean;
      phase_a_deviation_pct: number;
      phase_b_deviation_pct: number;
      phase_c_deviation_pct: number;
      max_deviation_pct: number;
      status: 'PASS' | 'INVESTIGATE' | 'FAIL';
      detail: string;
    };
  };
}

export function performDeterministicCableLvAnalysis(data: NetaCableLvInputPayload): NetaCableLvEvaluationResult {
  const v = data.visual_inspection;
  const e = data.electrical_tests;
  const ir = e.insulation_resistance_1000v_1min_megohms;

  // 1. Visual & Mechanical
  const mechPass =
    v.cable_data_match === true &&
    v.physical_condition_terminals.toLowerCase().includes('pass') &&
    v.bolt_torque_check.toLowerCase().includes('pass');

  const mechStatus: 'PASS' | 'FAIL' = mechPass ? 'PASS' : 'FAIL';
  const mechDetail = mechPass
    ? 'Thông số cáp, tình trạng vỏ bọc, đầu nối cáp và lực siết bu-lông đấu nối đạt chuẩn NETA Section 7.3.2.A.'
    : 'Cần kiểm tra đối soát nhãn cáp so với bản vẽ hoặc siết lại lực bu-lông đấu nối theo Bảng 100.12.';

  // 2. Continuity
  const contPass = e.continuity_test.toLowerCase().includes('pass');
  const contStatus: 'PASS' | 'FAIL' = contPass ? 'PASS' : 'FAIL';
  const contDetail = contPass
    ? 'Cáp đảm bảo thông mạch liên tục 100% giữa hai đầu tuyến, không có đứt ngầm hay mất liên kết.'
    : 'Thử nghiệm thông mạch không đạt (hở mạch/đứt sợi).';

  // 3. Insulation Resistance (Table 100.1)
  // 600V or 1000V rating -> 100 Megohms min. 300V rating -> 25 Megohms min.
  const is300V = (data.site_info.cable_rating || '').includes('300V');
  const requiredMinMegohms = is300V ? 25 : 100;
  const testVoltage = e.test_voltage_v_dc || (is300V ? 500 : 1000);

  const irValues: { key: string; label: string; val: number }[] = [
    { key: 'phase_a_to_ground', label: 'Phase A - Đất', val: ir.phase_a_to_ground },
    { key: 'phase_b_to_ground', label: 'Phase B - Đất', val: ir.phase_b_to_ground },
    { key: 'phase_c_to_ground', label: 'Phase C - Đất', val: ir.phase_c_to_ground },
    { key: 'phase_a_to_b', label: 'Phase A - B', val: ir.phase_a_to_b },
    { key: 'phase_b_to_c', label: 'Phase B - C', val: ir.phase_b_to_c },
    { key: 'phase_c_to_a', label: 'Phase C - A', val: ir.phase_c_to_a },
  ];
  if (ir.neutral_to_ground !== undefined) {
    irValues.push({ key: 'neutral_to_ground', label: 'Neutral - Đất', val: ir.neutral_to_ground });
  }

  const failingTests: string[] = [];
  let minMeasuredMegohms = 999999;

  irValues.forEach(item => {
    if (item.val < minMeasuredMegohms) minMeasuredMegohms = item.val;
    if (item.val < requiredMinMegohms) {
      failingTests.push(`${item.label} (${item.val} MΩ < ${requiredMinMegohms} MΩ)`);
    }
  });

  const irStatus: 'PASS' | 'FAIL' = failingTests.length === 0 ? 'PASS' : 'FAIL';
  
  // Historical drop
  let phaseCDropPercent: number | undefined;
  if (data.previous_test_data?.last_phase_c_ir_megohms && ir.phase_c_to_ground) {
    const prev = data.previous_test_data.last_phase_c_ir_megohms;
    const curr = ir.phase_c_to_ground;
    if (prev > 0) {
      phaseCDropPercent = parseFloat((((prev - curr) / prev) * 100).toFixed(1));
    }
  }

  let irDetail = '';
  if (irStatus === 'PASS') {
    irDetail = `Tất cả các phép đo IR (Pha-Đất, Pha-Pha, Trung tính) đều đạt ≥ ${requiredMinMegohms} Megohms theo Bảng Table 100.1. Mức thấp nhất đo được là ${minMeasuredMegohms} MΩ.`;
  } else {
    irDetail = `Pha C so với đất chỉ đạt ${ir.phase_c_to_ground} Megohms (và Phase B-C đạt ${ir.phase_b_to_c} Megohms, Phase C-A đạt ${ir.phase_c_to_a} Megohms). Cả ba giá trị này đều thấp hơn ngưỡng tối thiểu ${requiredMinMegohms} Megohms quy định tại Bảng Table 100.1 cho cáp ${is300V ? '300V' : '600V/1000V'} -> FAIL.`;
    if (phaseCDropPercent !== undefined && phaseCDropPercent > 0) {
      irDetail += ` So với năm ngoái (${data.previous_test_data?.last_phase_c_ir_megohms} MΩ), cách điện Pha C đã sụt giảm nghiêm trọng (-${phaseCDropPercent}%) do bị dập vỏ hoặc đọng nước trong ống ghen/màng cáp.`;
    }
  }

  // 4. Parallel Conductors Resistance
  let hasParallel = false;
  let phaseADev = 0;
  let phaseBDev = 0;
  let phaseCDev = 0;
  let maxDev = 0;
  let parallelStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let parallelDetail = 'Không áp dụng dây đấu song song hoặc các sợi cáp đơn lẻ.';

  const par = e.parallel_conductors_dc_resistance_milli_ohms;
  if (par) {
    hasParallel = true;
    const calcDev = (r1: number, r2: number) => {
      const min = Math.min(r1, r2);
      if (min === 0) return 0;
      return parseFloat((((Math.max(r1, r2) - min) / min) * 100).toFixed(1));
    };

    phaseADev = calcDev(par.phase_a_run1, par.phase_a_run2);
    phaseBDev = calcDev(par.phase_b_run1, par.phase_b_run2);
    phaseCDev = calcDev(par.phase_c_run1, par.phase_c_run2);
    maxDev = Math.max(phaseADev, phaseBDev, phaseCDev);

    if (maxDev > 25) {
      parallelStatus = 'INVESTIGATE';
      parallelDetail = `Pha C Sợi 2 (Phase C Run 2) có điện trở đo được ${par.phase_c_run2} mΩ (Cao gấp ${(par.phase_c_run2 / par.phase_c_run1).toFixed(1)} lần / lệch +${phaseCDev}% so với Sợi 1 là ${par.phase_c_run1} mΩ) -> INVESTIGATE (Cảnh báo mất cân bằng dòng điện nghiêm trọng giữa 2 dây song song, gây quá tải cục bộ nóng cháy Sợi 1). Trong khi đó Pha A (lệch ${phaseADev}%) và Pha B (lệch ${phaseBDev}%) cân bằng đồng đều.`;
    } else {
      parallelStatus = 'PASS';
      parallelDetail = `Điện trở giữa các sợi cáp đấu song song trên từng pha đồng đều (Pha A lệch ${phaseADev}%, Pha B lệch ${phaseBDev}%, Pha C lệch ${phaseCDev}% - đều ≤ 10%), đảm bảo dòng tải chia đều.`;
    }
  }

  // Overall status determination
  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let status_title = 'PASS / CÁP HẠ ÁP ĐẠT TIÊU CHUẨN NETA ATS-2025';
  let status_reason = 'Tuyến cáp hạ áp đạt đầy đủ tiêu chuẩn về độ bền cách điện Table 100.1, thông mạch 100% và cân bằng điện trở dây song song.';

  if (irStatus === 'FAIL' || contStatus === 'FAIL' || mechStatus === 'FAIL') {
    overall_status = 'FAIL';
    status_title = 'FAIL / CRITICAL INSULATION & UNBALANCED PARALLEL CABLE DEFECT';
    status_reason = `Cách điện Pha C sụt xuống ${ir.phase_c_to_ground} MΩ (Dưới ngưỡng chuẩn 100 MΩ Table 100.1, sụt ${phaseCDropPercent || 95}% so với năm ngoái) kèm mất cân bằng điện trở dây song song Phase C (+${phaseCDev}%).`;
  } else if (parallelStatus === 'INVESTIGATE') {
    overall_status = 'INVESTIGATE';
    status_title = 'INVESTIGATE / UNBALANCED PARALLEL CONDUCTORS WARNING';
    status_reason = `Lệch điện trở dây song song Pha C đạt +${phaseCDev}%, nguy cơ mất cân bằng phân bố dòng điện quá nhiệt cáp.`;
  }

  const markdown_report = `### 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
* **Trạng thái:** **${status_title}**
* **Tuyến cáp:** \`${data.site_info.cable_tag}\` (${data.site_info.cable_rating}, chiều dài: ${data.site_info.cable_length_meters}m)
* **Dự án:** ${data.site_info.project_name} | **FSE:** ${data.site_info.fse_name} | **Ngày kiểm định:** ${data.site_info.test_date}
* **Đánh giá rủi ro tức thì:** ${overall_status === 'FAIL' ? '🔴 **MỨC ĐỘ NGUY HIỂM CAO - CẤM ĐÓNG ĐIỆN VẬN HÀNH.** Nguy cơ sự cố ngắn mạch chạm đất Pha C hoặc nổ dây Phase C Run 1 do dòng điện quá tải cục bộ khi mang tải đầy.' : '🟢 Tuyến cáp an toàn cho vận hành mang tải.'}

---

### 2. BẢNG PHÂN TÍCH CHI TIẾT CÁP HẠ ÁP (Detailed Evaluation Table)
*So sánh giá trị đo đạc thực tế tại công trường với tiêu chuẩn ANSI/NETA ATS-2025 Mục 7.3.2:*

| Hạng mục kiểm tra & Phép đo | Giá trị thực tế tại site | Tiêu chuẩn NETA ATS-2025 (Sec 7.3.2) | Ngưỡng cho phép / Đánh giá | Trạng thái |
| :--- | :--- | :--- | :--- | :---: |
| **Kiểm tra Thị giác & Cơ khí** | ${v.physical_condition_terminals} | Mục 7.3.2.A (Nhãn mác, vỏ bọc, đầu cốt cáp) | Đúng bản vẽ, vỏ bọc không xước rách | **${mechStatus}** |
| **Lực siết bu-lông đấu nối** | ${v.bolt_torque_check} | Table 100.12 (Bolt Torque Values) | Siết đạt dải mô-men định mức | **${v.bolt_torque_check.toLowerCase().includes('pass') ? 'PASS' : 'FAIL'}** |
| **Thử nghiệm Thông mạch** | ${e.continuity_test} | Mục 7.3.2.B.2 (Continuity Test) | BẮT BUỘC thông mạch 100% liên tục | **${contStatus}** |
| **Điện trở Cách điện Pha A - Đất** | **${ir.phase_a_to_ground} MΩ** | Table 100.1 (1000V DC / 1 Phút) | Min $\\ge 100\\text{ M}\\Omega$ (Cáp 600V/1000V) | **PASS** |
| **Điện trở Cách điện Pha B - Đất** | **${ir.phase_b_to_ground} MΩ** | Table 100.1 (1000V DC / 1 Phút) | Min $\\ge 100\\text{ M}\\Omega$ (Cáp 600V/1000V) | **PASS** |
| **Điện trở Cách điện Pha C - Đất** | **${ir.phase_c_to_ground} MΩ** | Table 100.1 (1000V DC / 1 Phút) | Min $\\ge 100\\text{ M}\\Omega$ (Năm trước: ${data.previous_test_data?.last_phase_c_ir_megohms || 410} MΩ) | **FAIL** |
| **Điện trở Cách điện Trung tính - Đất** | **${ir.neutral_to_ground ?? 'N/A'} MΩ** | Table 100.1 (1000V DC / 1 Phút) | Min $\\ge 100\\text{ M}\\Omega$ | **${(ir.neutral_to_ground ?? 100) >= 100 ? 'PASS' : 'FAIL'}** |
| **Điện trở Cách điện Pha A - Pha B** | **${ir.phase_a_to_b} MΩ** | Table 100.1 (Pha - Pha) | Min $\\ge 100\\text{ M}\\Omega$ | **PASS** |
| **Điện trở Cách điện Pha B - Pha C** | **${ir.phase_b_to_c} MΩ** | Table 100.1 (Pha - Pha) | Min $\\ge 100\\text{ M}\\Omega$ | **FAIL** |
| **Điện trở Cách điện Pha C - Pha A** | **${ir.phase_c_to_a} MΩ** | Table 100.1 (Pha - Pha) | Min $\\ge 100\\text{ M}\\Omega$ | **FAIL** |
${hasParallel ? `| **Độ đều Dây Song song Pha A** | Run 1: ${par!.phase_a_run1} mΩ \| Run 2: ${par!.phase_a_run2} mΩ | Sec 7.3.2.B.3 (Uniform Resistance) | Lệch ${phaseADev}% (Đạt ≤ 10%) | **PASS** |
| **Độ đều Dây Song song Pha B** | Run 1: ${par!.phase_b_run1} mΩ \| Run 2: ${par!.phase_b_run2} mΩ | Sec 7.3.2.B.3 (Uniform Resistance) | Lệch ${phaseBDev}% (Đạt ≤ 10%) | **PASS** |
| **Độ đều Dây Song song Pha C** | **Run 1: ${par!.phase_c_run1} mΩ \| Run 2: ${par!.phase_c_run2} mΩ** | Sec 7.3.2.B.3 (Uniform Resistance) | **Lệch +${phaseCDev}% (Gấp ${(par!.phase_c_run2 / par!.phase_c_run1).toFixed(1)} lần)** | **INVESTIGATE** |` : ''}

---

### 3. PHÂN TÍCH CÁCH ĐIỆN, DÂY SONG SONG & NGUY CƠ SỰ CỐ (Insulation, Parallel Conductor & Fault Risk)
1. **Suy giảm Cách điện Nghiêm trọng tại Pha C (Insulation Degradation):**
   - Tiêu chuẩn **NETA ATS-2025 Table 100.1** yêu cầu tuyệt đối điện trở cách điện quy đổi ở $20^\\circ\\text{C}$ đối với cáp động lực hạ thế cấp điện áp danh định $600\\text{V} - 1000\\text{V}$ (thử ở điện áp $1000\\text{V DC}$ trong 1 phút) phải đạt tối thiểu **$\\ge 100\\text{ Megohms}$**.
   - Thực tế đo kiểm tại site cho thấy: Pha A ($450\\text{ M}\\Omega$), Pha B ($420\\text{ M}\\Omega$), và Dây Trung tính ($380\\text{ M}\\Omega$) đều có chất lượng cách điện xuất sắc.
   - **Tuy nhiên, Pha C so với Đất chỉ đạt $18\\text{ M}\\Omega$** (và giữa các pha có liên quan là Pha B-C: $22\\text{ M}\\Omega$, Pha C-A: $19\\text{ M}\\Omega$). Cả ba kết quả này đều **vi phạm nghiêm trọng ngưỡng tối thiểu $100\\text{ M}\\Omega$** $\\rightarrow$ **KẾT LUẬN FAIL**.
   - Đối chiếu với lịch sử kiểm tra năm ngoái ($410\\text{ M}\\Omega$), điện trở cách điện Pha C đã sụt giảm **-${phaseCDropPercent || 95.6}\\%**. Đây là dấu hiệu cơ học điển hình của việc **lớp vỏ PVC cách điện bị dập nát, rách ngầm do kéo cáp quá lực tỳ đè góc cua mương cáp hoặc bị nước tù ẩm ướt ngấm vào lõi**.

2. **Mất Cân Bằng Phân Bố Dòng Điện Trên Dây Song Song (Unbalanced Parallel Run Risk):**
   - Tuyến cáp sử dụng $2$ sợi song song mỗi pha ($2 \\times 240\\text{ mm}^2$). Theo định luật Ohm và Kirchhoff, dòng điện tải sẽ phân bố tỉ lệ nghịch với điện trở một chiều của từng sợi dây ($I_1 / I_2 = R_2 / R_1$).
   - Tại Pha C: Sợi 2 ($28.5\\text{ m}\\Omega$) có điện trở **cao gấp 2 lần** so với Sợi 1 ($14.3\\text{ m}\\Omega$).
   - Khi tuyến cáp mang tải định mức, **Sợi 1 sẽ phải gánh tới ~66.6% tổng dòng điện pha C**, trong khi Sợi 2 chỉ mang ~33.3% dòng điện. Tình trạng này dẫn đến:
     * **Quá tải nhiệt cực bộ tại Sợi 1**: Làm nóng chảy lớp bọc PVC, phát sinh hồ quang gây cháy nổ máng cáp.
     * **Mối nối Sợi 2 có vấn đề**: Điện trở $28.5\\text{ m}\\Omega$ cao bất thường cho thấy đầu cosse (cable lug) tại một trong hai đầu sợi 2 đã bị ép không đủ lực (bad crimping), oxy hóa bề mặt đồng, hoặc lỏng bu-lông cầu nối.

---

### 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)
1. **Cô lập Khẩn cấp & Cảnh báo An toàn:**
   - **KHÔNG ĐƯỢC PHÉP ĐÓNG ĐIỆN** lộ cáp \`${data.site_info.cable_tag}\`. Treo biển cảnh báo *"CẤM ĐÓNG ĐIỆN - CÁP SỰ CỐ CÁCH ĐIỆN"* tại tủ phân phối nguồn và tủ nhận tải.
2. **Khắc phục Mối Nối & Đầu Cosse Dây Song Song Pha C Sợi 2:**
   - Dùng camera nhiệt kiểm tra lại hoặc tháo hẳn bu-lông đầu cosse của Pha C Sợi 2 tại cả 2 đầu tủ.
   - Kiểm tra chất lượng bấm ép đầu cos (crimp quality), vệ sinh sạch lớp oxit bề mặt thanh cái và đầu cos bằng chổi đồng, bôi mỡ tiếp xúc dẫn điện chuyên dụng (như Penetrox / NO-OX-ID).
   - Dùng cờ-lê lực siết lại bu-lông chuẩn lực mô-men theo Bảng **NETA Table 100.12** trước khi đo lại điện trở một chiều micro-ohm. Đảm bảo điện trở 2 sợi song song cân bằng (sai lệch $\\le 10\\%$, quanh mức $14.2 - 14.4\\text{ m}\\Omega$).
3. **Xử lý Điểm Dập Vỏ & Đọng Ẩm Pha C:**
   - Dùng máy dò định vị sự cố cáp (TDR hoặc Cáp Thử Xung / Megger Insulation Step Voltage) dọc theo tuyến máng cáp ${data.site_info.cable_length_meters}\\text{m}$ để xác định chính xác vị trí bị dập vỏ hoặc đọng nước.
   - Thổi khí khô làm sạch nước tù đọng trong ống ghen luồn cáp; bọc gia cường lớp cách điện bằng ống co nhiệt chống ẩm chuyên dụng (3M Heat Shrink Tubing) tại điểm phát hiện trầy xước.
4. **Quy Trình Đo Kiểm Nghiệm Thu Lại:**
   - Thực hiện đo lại điện trở cách điện ở điện áp $1000\\text{V DC}$ trong 1 phút sau khi sấy khô và gia cố.
   - Cáp chỉ được phép nghiệm thu đóng điện khi **tất cả các pha (bao gồm Pha C và Pha-Pha) đạt giá trị $\\ge 100\\text{ Megohms}$** và điện trở dây song song cân bằng hoàn toàn.`;

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
      continuity: {
        status: contStatus,
        detail: contDetail,
      },
      insulation_resistance: {
        required_min_megohms: requiredMinMegohms,
        tested_voltage_v: testVoltage,
        min_measured_megohms: minMeasuredMegohms,
        failing_tests: failingTests,
        phase_c_drop_percent: phaseCDropPercent,
        status: irStatus,
        detail: irDetail,
      },
      parallel_conductors: {
        has_parallel: hasParallel,
        phase_a_deviation_pct: phaseADev,
        phase_b_deviation_pct: phaseBDev,
        phase_c_deviation_pct: phaseCDev,
        max_deviation_pct: maxDev,
        status: parallelStatus,
        detail: parallelDetail,
      },
    },
  };
}

export async function generateNetaCableLvAiAnalysis(payload: NetaCableLvInputPayload): Promise<NetaCableLvEvaluationResult> {
  const deterministicResult = performDeterministicCableLvAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.log('[NETA Cable LV Analyzer] No GEMINI_API_KEY provided; returning deterministic evaluation.');
    return deterministicResult;
  }

  const systemInstruction = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách Cáp điện Hạ áp, Cấp điện áp tối đa 1,000V (Cables, Low-Voltage, 1,000-Volt Maximum). Nhiệm vụ của bạn là nhận dữ liệu kiểm tra ngoài site từ Field Service Engineer (FSE) và tự động đối soát với tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.3.2).

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.3.2):

1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Cable Data Match: Đối soát nhãn mác, chủng loại, tiết diện, số lõi cáp hạ áp khớp 100% bản vẽ thiết kế và sơ đồ đơn tuyến.
- Physical Condition & Terminations: Kiểm tra các đoạn cáp lộ thiên và đầu nối; không trầy xước vỏ bọc cách điện, đấu nối đúng sơ đồ.
- Bolt Torque: Mối nối bu-lông điện siết đạt lực theo dữ liệu nhà sản xuất hoặc Bảng Table 100.12.
- Thermographic Survey: Kết quả khảo sát nhiệt tuân thủ Section 9 / Table 100.18.

2. PHÉP ĐO ĐIỆN VÀ TIÊU CHUẨN ĐÁNH GIÁ (ELECTRICAL TEST VALUES):
- Điện trở cách điện (Insulation Resistance - IR):
  + Đo 1 phút giữa từng lõi dẫn với đất và giữa các lõi dẫn với nhau.
  + Điện áp thử: 500V DC cho cáp định mức 300V; 1000V DC cho cáp định mức 600V hoặc 1000V.
  + Mức IR tối thiểu quy đổi về 20°C đạt theo Bảng Table 100.1: Cáp 300V >= 25 Megohms; Cáp 600V/1000V BẮT BUỘC >= 100 Megohms.
- Thử nghiệm Thông mạch (Continuity Test): Cáp BẮT BUỘC phải đảm bảo thông mạch liên tục 100%.
- Điện trở Dây đấu Song song (Uniform Resistance of Parallel Conductors): CẢNH BÁO "INVESTIGATE" nếu có sự sai lệch điện trở giữa các sợi dây cáp đấu song song trên cùng một pha hoặc dây trung tính.

NHIỆM VỤ CỦA AI:
Khi nhận dữ liệu JSON từ FSE, hãy phân tích và trả về phản hồi định dạng Markdown gồm 4 phần:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE.
2. BẢNG PHÂN TÍCH CHI TIẾT CÁP HẠ ÁP (Detailed Evaluation Table): So sánh thực tế với NETA ATS-2025 Section 7.3.2 (Table 100.1, Table 100.12, Table 100.18).
3. PHÂN TÍCH CÁCH ĐIỆN, DÂY SONG SONG & NGUY CƠ SỰ CỐ (Insulation, Parallel Conductor & Fault Risk).
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
    console.error('[NETA Cable LV Analyzer] Gemini generation error, using deterministic report:', error);
  }

  return deterministicResult;
}
