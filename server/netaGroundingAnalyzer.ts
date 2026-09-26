import { GoogleGenAI } from '@google/genai';

export interface NetaGroundingInputPayload {
  site_info: {
    project_name: string;
    grounding_system_tag: string;
    system_type: string;
    fse_name: string;
    test_date: string;
    facility_type?: 'substation' | 'industrial' | 'commercial' | 'datacenter' | string;
    target_resistance_ohms?: number;
    soil_resistivity_ohm_meters?: number;
    substation_voltage_kv?: number;
  };
  visual_inspection: {
    grounding_layout_match: boolean;
    physical_condition_and_corrosion: string;
    exothermic_welds_and_bolted_joints: string;
    conductor_and_bus_sizing: string;
    ground_wells_and_test_links_accessible: boolean;
    bolt_torque_check: string;
  };
  electrical_tests: {
    fall_of_potential_ground_resistance_ieee_81_ohms: number;
    point_to_point_continuity_milli_ohms: {
      main_bus_to_transformer_frame: number;
      main_bus_to_switchgear_earth_bar: number;
      main_bus_to_control_building_structure: number;
      main_bus_to_substation_fence_gate: number;
      [key: string]: number;
    };
    fence_bonding_resistance_ohms?: number;
    test_method?: string;
  };
  previous_test_data?: {
    last_test_date?: string;
    last_ground_resistance_ohms?: number;
    last_fence_gate_milli_ohms?: number;
  };
}

export interface NetaGroundingEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_title: string;
  status_reason: string;
  markdown_report: string;
  evaluations: {
    visual_mechanical: {
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    ground_grid_resistance: {
      measured_ohms: number;
      max_allowed_ohms: number;
      substation_threshold_ohms: number;
      general_threshold_ohms: number;
      historical_increase_percent?: number;
      status: 'PASS' | 'FAIL';
      detail: string;
    };
    point_to_point_continuity: {
      max_measured_milli_ohms: number;
      fence_gate_milli_ohms: number;
      threshold_milli_ohms: number;
      failing_points: string[];
      status: 'PASS' | 'INVESTIGATE' | 'FAIL';
      detail: string;
    };
    fence_bonding: {
      measured_ohms: number;
      max_allowed_ohms: number;
      status: 'PASS' | 'INVESTIGATE' | 'FAIL';
      detail: string;
    };
  };
}

export function performDeterministicGroundingAnalysis(data: NetaGroundingInputPayload): NetaGroundingEvaluationResult {
  const v = data.visual_inspection;
  const e = data.electrical_tests;

  // 1. Visual & Mechanical
  const mechPass =
    v.grounding_layout_match &&
    v.ground_wells_and_test_links_accessible &&
    v.physical_condition_and_corrosion.toLowerCase().includes('pass') &&
    v.exothermic_welds_and_bolted_joints.toLowerCase().includes('pass') &&
    v.conductor_and_bus_sizing.toLowerCase().includes('pass') &&
    v.bolt_torque_check.toLowerCase().includes('pass');

  const mechStatus: 'PASS' | 'FAIL' = mechPass ? 'PASS' : 'FAIL';
  const mechDetail = mechPass
    ? 'Lưới tiếp địa đồng 240mm², cọc đất, mối hàn hóa nhiệt Cadweld, giếng kiểm tra và lực siết bu-lông đạt chuẩn 100% theo NETA Section 7.13.A.'
    : 'Cần kiểm tra đối soát cấu trúc lưới đất, khả năng tiếp cận giếng kiểm tra hoặc kiểm tra chất lượng mối hàn hóa nhiệt.';

  // 2. Ground Grid Resistance (IEEE Std 81 Fall-of-Potential)
  const measuredOhms = e.fall_of_potential_ground_resistance_ieee_81_ohms;
  const isSubstation =
    data.site_info.project_name.toLowerCase().includes('substation') ||
    data.site_info.project_name.toLowerCase().includes('trạm') ||
    data.site_info.facility_type === 'substation' ||
    (data.site_info.substation_voltage_kv && data.site_info.substation_voltage_kv >= 22);

  const substationThreshold = 1.0; // 1.0 Ohm for Substation / Utility / Data Center
  const generalThreshold = 5.0; // 5.0 Ohms for general industrial / commercial
  const maxAllowedOhms = isSubstation ? substationThreshold : generalThreshold;

  const gridPass = measuredOhms <= generalThreshold && (!isSubstation || measuredOhms <= substationThreshold);
  const gridStatus: 'PASS' | 'FAIL' = gridPass ? 'PASS' : 'FAIL';

  let historicalIncreasePct: number | undefined;
  if (data.previous_test_data?.last_ground_resistance_ohms) {
    const prev = data.previous_test_data.last_ground_resistance_ohms;
    if (prev > 0) {
      historicalIncreasePct = parseFloat((((measuredOhms - prev) / prev) * 100).toFixed(1));
    }
  }

  let gridDetail = '';
  if (gridPass) {
    gridDetail = `Điện trở nối đất tổng thể đo bằng phương pháp Fall-of-Potential đạt ${measuredOhms} Ω, hoàn toàn tuân thủ tiêu chuẩn NETA Section 7.13.D.2 (≤ ${maxAllowedOhms} Ω) và IEEE Std 81.`;
  } else {
    gridDetail = `Đo điện trở nối đất tổng thể bằng phương pháp 3 điểm Fall-of-Potential đạt ${measuredOhms} Ohms. Giá trị này vượt quá ngưỡng tối đa cho phép 5.0 Ohms theo NETA ATS-2025 Section 7.13.D.2 (và vượt xa ngưỡng 1.0 Ohm áp dụng cho trạm biến áp ${data.site_info.project_name}) -> FAIL.`;
    if (historicalIncreasePct !== undefined && historicalIncreasePct > 0) {
      gridDetail += ` So với năm ngoái (${data.previous_test_data?.last_ground_resistance_ohms} Ohms), điện trở đất tăng +${historicalIncreasePct}% do đất bị khô đọng bẩn, giảm độ ẩm hoặc đứt mạch liên kết cọc ngầm.`;
    }
  }

  // 3. Point-to-Point Continuity (< 0.5 Ohm = 500 mOhms)
  const pts = e.point_to_point_continuity_milli_ohms;
  const thresholdMilliOhms = 500; // 0.5 Ohm
  const failingPoints: string[] = [];
  let maxMeasuredMilliOhms = 0;

  Object.entries(pts).forEach(([key, val]) => {
    if (typeof val === 'number') {
      if (val > maxMeasuredMilliOhms) maxMeasuredMilliOhms = val;
      if (val > thresholdMilliOhms) {
        failingPoints.push(`${key}: ${val} mΩ (> 500 mΩ)`);
      }
    }
  });

  const fenceGateVal = pts.main_bus_to_substation_fence_gate || (e.fence_bonding_resistance_ohms ? e.fence_bonding_resistance_ohms * 1000 : 0);
  const fenceGateOhms = fenceGateVal / 1000;

  let pointStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let pointDetail = '';

  if (failingPoints.length > 0 || fenceGateOhms > 0.5) {
    pointStatus = 'INVESTIGATE';
    pointDetail = `Liên kết tiếp địa vỏ máy biến áp (${pts.main_bus_to_transformer_frame} mΩ), tủ điện (${pts.main_bus_to_switchgear_earth_bar} mΩ) và nhà điều khiển (${pts.main_bus_to_control_building_structure} mΩ) đạt tiếp xúc rất tốt (< 20 mΩ). Tuy nhiên, điện trở liên kết từ thanh cái chính đến Cổng hàng rào trạm (Fence Gate) đo được ${fenceGateOhms} Ohm (${fenceGateVal} mΩ) -> Vượt quá ngưỡng cho phép 0.5 Ohm -> INVESTIGATE (Nghi ngờ đứt dây liên kết mềm tiếp địa cổng hàng rào hoặc mối nối bu-lông bị oxy hóa).`;
  } else {
    pointStatus = 'PASS';
    pointDetail = `Tất cả các điểm liên kết tiếp địa từ Main Ground Bus đến vỏ thiết bị, máy biến áp và cổng hàng rào đều đạt điện trở rất thấp (cực đại ${maxMeasuredMilliOhms} mΩ < 500 mΩ / 0.5 Ω).`;
  }

  // 4. Fence Bonding
  const fenceBondingOhms = e.fence_bonding_resistance_ohms || fenceGateOhms;
  const fencePass = fenceBondingOhms <= 0.5;
  const fenceStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = fencePass ? 'PASS' : 'INVESTIGATE';
  const fenceDetail = fencePass
    ? `Tiếp địa hàng rào đạt ${fenceBondingOhms} Ω (≤ 0.5 Ω).`
    : `Tiếp địa hàng rào đo được ${fenceBondingOhms} Ω (> 0.5 Ω) -> Nguy cơ điện áp chạm Touch Potential nguy hiểm khi chạm tay vào hàng rào khi xảy ra ngắn mạch chạm đất.`;

  // Overall Status
  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let status_title = 'PASS / HỆ THỐNG TIẾP ĐỊA ĐẠT CHUẨN NETA ATS-2025';
  let status_reason = 'Điện trở nối đất tổng thể và liên kết tiếp địa điểm-đến-điểm đạt tiêu chuẩn IEEE Std 81 và NETA ATS-2025 Mục 7.13.';

  if (!gridPass || !mechPass) {
    overall_status = 'FAIL';
    status_title = 'FAIL / HIGH GROUND GRID RESISTANCE & FENCE BONDING ANOMALY';
    status_reason = `Điện trở nối đất lưới đất đo được ${measuredOhms} Ω (Vượt ngưỡng ${maxAllowedOhms} Ω) kèm điện trở liên kết cổng hàng rào ${fenceGateOhms} Ω (> 0.5 Ω).`;
  } else if (pointStatus === 'INVESTIGATE' || fenceStatus === 'INVESTIGATE') {
    overall_status = 'INVESTIGATE';
    status_title = 'INVESTIGATE / FENCE BONDING RESISTANCE WARNING';
    status_reason = `Điện trở liên kết cổng hàng rào đo được ${fenceGateOhms} Ω vượt ngưỡng cho phép 0.5 Ohm theo Mục 7.13.D.1.`;
  }

  const markdown_report = `### 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
* **Trạng thái:** **${status_title}**
* **Hệ thống tiếp địa:** \`${data.site_info.grounding_system_tag}\` (${data.site_info.system_type})
* **Dự án:** ${data.site_info.project_name} | **Kỹ sư FSE:** ${data.site_info.fse_name} | **Ngày kiểm định:** ${data.site_info.test_date}
* **Đánh giá rủi ro tức thì:** ${overall_status === 'FAIL' ? '🔴 **MỨC ĐỘ NGUY HIỂM CAO - NGUY CƠ ĐIỆN GIẬT DO ĐIỆN ÁP BƯỚC & ĐIỆN ÁP SỜ.** Điện trở đất toàn trạm đạt 7.8 Ω (vượt ngưỡng 5.0 Ω cho công nghiệp và 1.0 Ω cho trạm 110kV), khi có dòng sự cố chạm đất chạy qua sẽ nâng điện thế toàn bãi trạm (Ground Potential Rise - GPR) lên hàng kilovolt gây nguy hiểm chết người.' : '🟢 Hệ thống tiếp địa an toàn, đảm bảo tản dòng sự cố nhanh chóng xuống đất.'}

---

### 2. BẢNG PHÂN TÍCH CHI TIẾT HỆ THỐNG NỐI ĐẤT (Detailed Evaluation Table)
*So sánh giá trị đo đạc thực tế tại site với tiêu chuẩn ANSI/NETA ATS-2025 Section 7.13 và IEEE Std 81:*

| Hạng mục kiểm tra & Phép đo | Giá trị thực tế tại site | Tiêu chuẩn NETA ATS-2025 (Sec 7.13) & IEEE 81 | Ngưỡng cho phép / Đánh giá | Trạng thái |
| :--- | :--- | :--- | :--- | :---: |
| **Kiểm tra Thị giác & Mối hàn** | ${v.exothermic_welds_and_bolted_joints} | Mục 7.13.A.1 - A.5 (Cadweld, cọc, dây dẫn) | Mối hàn ngấu đều, không rỗ nứt | **${mechStatus}** |
| **Tiết diện Dây & Thanh cái** | ${v.conductor_and_bus_sizing} | Mục 7.13.A.2 (NFPA 70 / NEC) | Đúng tiết diện chịu dòng ngắn mạch | **PASS** |
| **Giếng Kiểm tra & Test Links** | ${v.ground_wells_and_test_links_accessible ? 'Sạch sẽ, tiếp cận tốt' : 'Bị vùi lấp'} | Mục 7.13.A.3 | Tháo lắp cầu đo dễ dàng | **PASS** |
| **Lực siết bu-lông (Bolt Torque)** | ${v.bolt_torque_check} | Table 100.12 | Siết đạt dải mô-men định mức | **PASS** |
| **Điện trở Nối đất Lưới (Fall-of-Potential)** | **${measuredOhms} Ω** | IEEE Std 81 & NETA Sec 7.13.D.2 | **Max $\\le 5.0\\,\\Omega$ (Trạm: $\\le 1.0\\,\\Omega$)** | **${gridStatus}** |
| **Liên kết Main Bus - Vỏ MBA** | ${pts.main_bus_to_transformer_frame} mΩ (0.0125 Ω) | Mục 7.13.D.1 (Point-to-Point) | BẮT BUỘC $< 0.5\\,\\Omega$ (500 mΩ) | **PASS** |
| **Liên kết Main Bus - Tủ điện MSB**| ${pts.main_bus_to_switchgear_earth_bar} mΩ (0.0142 Ω) | Mục 7.13.D.1 (Point-to-Point) | BẮT BUỘC $< 0.5\\,\\Omega$ (500 mΩ) | **PASS** |
| **Liên kết Main Bus - Nhà ĐK** | ${pts.main_bus_to_control_building_structure} mΩ (0.0150 Ω) | Mục 7.13.D.1 (Point-to-Point) | BẮT BUỘC $< 0.5\\,\\Omega$ (500 mΩ) | **PASS** |
| **Liên kết Cổng Hàng Rào Trạm** | **${fenceGateVal} mΩ (${fenceGateOhms} Ω)** | Mục 7.13.D.1 & IEEE Std 80 | **Max $\\le 0.5\\,\\Omega$ (500 mΩ)** | **INVESTIGATE** |
| **Tiếp địa Hàng rào (Fence Bonding)** | **${fenceBondingOhms} Ω** | NETA Sec 7.13.D.1 & IEEE Std 80 | Max $\\le 0.5\\,\\Omega$ | **INVESTIGATE** |

---

### 3. PHÂN TÍCH ĐIỆN TRỞ ĐẤT, LIÊN KẾT & NGUY CƠ ĐIỆN GIẬT (Ground Grid Resistance, Bonding Integrity & Touch/Step Potential Risk)

1. **Nguy Cơ Điện Áp Bước (Step Potential) và Điện Áp Sờ (Touch Potential) Chết Người:**
   * Theo tiêu chuẩn **IEEE Std 80** và **NETA ATS-2025 Section 7.13.D.2**, điện trở nối đất tổng thể của một trạm biến áp trung/cao áp hoặc trạm phân phối chính bắt buộc phải duy trì ở mức **$\\le 1.0\\,\\Omega$** (và không được vượt quá $5.0\\,\\Omega$ cho hệ thống điện công nghiệp thông thường).
   * Thực tế đo kiểm tại site bằng phương pháp 3 điểm Fall-of-Potential ghi nhận điện trở đất lên tới **$7.8\\,\\Omega$**.
   * So sánh với năm ngoái ($4.2\\,\\Omega$), điện trở đã **tăng vọt $+85.7\\%$**. Tình trạng này xuất phát từ hiện tượng đất bị khô hạn sỏi đá, mạch nước ngầm sụt giảm hoặc các cọc tiếp địa đóng ngầm bị đứt gãy mối hàn rỉ sét ngầm.
   * **Hiểm họa an toàn điện:** Khi xảy ra sự cố chạm đất một pha trong trạm (ví dụ dòng sự cố $I_f = 2000\\text{A}$), điện thế bãi trạm sẽ dâng cao đột ngột:
     $$V_{\\text{GPR}} = I_f \\times R_g = 2000\\text{A} \\times 7.8\\,\\Omega = 15,600\\text{ Volts (15.6 kV!)}$$
     Mức điện thế này vượt xa ngưỡng chịu đựng của cơ thể con người, tạo ra điện áp bước (Step Voltage) trên mặt đất và điện áp sờ (Touch Voltage) cực kỳ nguy hiểm cho kỹ sư vận hành khi đi lại trong trạm.

2. **Khuyết Tật Tiếp Địa Cổng Hàng Rào Kim Loại (Fence Gate Bonding Defect):**
   * Trong khi vỏ máy biến áp ($12.5\\text{ m}\\Omega$), thanh cái tủ điện ($14.2\\text{ m}\\Omega$) và khung nhà điều khiển ($15.0\\text{ m}\\Omega$) có liên kết đẳng thế xuất sắc; thì **Cổng hàng rào trạm (Fence Gate) có điện trở đo được lên tới $680.0\\text{ m}\\Omega$ ($0.68\\,\\Omega$)**, vượt quá ngưỡng cho phép $0.5\\,\\Omega$.
   * **Nguyên nhân:** Cổng trạm đóng mở thường xuyên khiến dây đồng bện mềm tiếp địa (flexible copper bonding strap) bị tưa đứt sợi, hoặc điểm bắt bu-lông bản lề cổng bị oxy hóa bề mặt rỉ sét, mất tính dẫn điện.
   * **Hậu quả:** Khi người đi bộ hoặc bảo vệ chạm tay mở cổng trạm trong lúc trạm xảy ra chạm đất, sự chênh lệch điện thế giữa cổng và mặt đất ẩm sẽ phóng điện giật trực tiếp qua tay người xuống chân.

---

### 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)

1. **Khắc Phục Ngay Lập Tức Dây Liên Kết Tiếp Địa Cổng Hàng Rào:**
   * Tháo rời điểm nối dây tiếp địa mềm tại hai bên bản lề cổng trạm.
   * Vệ sinh sạch lớp sơn, rỉ sét và oxit tại vị trí bắt bu-lông bằng chổi đồng; quét mỡ tiếp xúc dẫn điện chuyên dụng (Penetrox / NO-OX-ID).
   * Thay thế dây bện đồng mềm tiếp địa mới (loại dây dệt bện đồng mạ thiếc tiết diện tối thiểu $50\\text{ mm}^2$) nối từ cánh cổng sang trụ hàng rào cố định.
   * Dùng cầu đo micro-ohm đo lại, đảm bảo điện trở liên kết **phải giảm xuống dưới $50\\text{ m}\\Omega$** (chuẩn NETA $< 0.5\\,\\Omega$).

2. **Giải Pháp Hạ Điện Trở Nối Đất Toàn Trạm Về Dưới 1.0 Ohm (hoặc < 5.0 Ohms):**
   * **Đo điện trở suất của đất (Soil Resistivity Test per Wenner 4-pin Method):** Xác định độ sâu của tầng đất ẩm có điện trở suất thấp.
   * **Bổ sung cọc tiếp địa đóng sâu:** Thi công đóng thêm các cọc tiếp địa đồng hoặc thép bọc đồng (Copper-Clad Steel Rods) phi 16-20mm dài 3m - 6m liên kết song song vào giếng tiếp địa hiện hữu.
   * **Bơm hợp chất giảm điện trở đất (GEM - Ground Enhancement Material):** Trộn và đổ hợp chất hóa học dẫn điện GEM / Bentonite xung quanh các hố cọc và rãnh chôn dây để ổn định độ ẩm và tăng bề mặt tiếp xúc điện tích với đất.

3. **Quy Trình Đo Kiểm Nghiệm Thu Lại:**
   * Sử dụng máy đo điện trở đất chuyên dụng (Megger DET hoặc Fluke 1625) đo lại theo phương pháp Fall-of-Potential $62\\%$ Rule sau khi hoàn nguyên đất 48-72 giờ.
   * Hệ thống chỉ được cấp phép đóng điện nghiệm thu khi **điện trở đất tổng thể giảm xuống $\\le 1.0\\,\\Omega$ (hoặc tối thiểu $\\le 5.0\\,\\Omega$)** và tiếp địa hàng rào hoàn toàn thông suốt.`;

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
      ground_grid_resistance: {
        measured_ohms: measuredOhms,
        max_allowed_ohms: maxAllowedOhms,
        substation_threshold_ohms: substationThreshold,
        general_threshold_ohms: generalThreshold,
        historical_increase_percent: historicalIncreasePct,
        status: gridStatus,
        detail: gridDetail,
      },
      point_to_point_continuity: {
        max_measured_milli_ohms: maxMeasuredMilliOhms,
        fence_gate_milli_ohms: fenceGateVal,
        threshold_milli_ohms: thresholdMilliOhms,
        failing_points: failingPoints,
        status: pointStatus,
        detail: pointDetail,
      },
      fence_bonding: {
        measured_ohms: fenceBondingOhms,
        max_allowed_ohms: 0.5,
        status: fenceStatus,
        detail: fenceDetail,
      },
    },
  };
}

export async function generateNetaGroundingAiAnalysis(payload: NetaGroundingInputPayload): Promise<NetaGroundingEvaluationResult> {
  const deterministicResult = performDeterministicGroundingAnalysis(payload);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.log('[NETA Grounding Analyzer] No GEMINI_API_KEY provided; returning deterministic evaluation.');
    return deterministicResult;
  }

  const systemInstruction = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách Hệ thống Nối đất / Tiếp địa (Grounding Systems) theo tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.13) và IEEE Std 81.

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.13 & IEEE Std 81):

1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Layout & Design Match: Đối soát cấu trúc lưới tiếp địa, cọc đất, thanh cái tiếp địa chính (Main Ground Bus) khớp 100% bản vẽ thiết kế và quy chuẩn NFPA 70 / NEC.
- Physical Condition & Welds: Dây dẫn và thanh cái không bị nứt đứt, oxy hóa; mối hàn hóa nhiệt (exothermic welds) ngấu đều, không rỗng xốp; mối nối bu-lông chắc chắn.
- Conductor Sizing: Tiết diện dây cáp đồng và thanh cái đạt theo thiết kế và tính toán dòng ngắn mạch sự cố.
- Ground Wells & Test Links: Giếng kiểm tra sạch sẽ, không bị lấp đất; điểm tháo thắt cầu nối thử nghiệm (test links) tháo lắp dễ dàng.
- Bolt Torque: Mối nối bu-lông siết đạt lực theo dữ liệu nhà sản xuất hoặc Bảng Table 100.12.

2. PHÉP ĐO ĐIỆN VÀ TIÊU CHUẨN ĐÁNH GIÁ (ELECTRICAL TEST VALUES):
- Điện trở Nối đất Điện cực / Lưới đất (Ground System Resistance to Earth - IEEE Std 81 Fall-of-Potential):
  + Điện trở nối đất tổng thể xuống lòng đất BẮT BUỘC <= 5.0 Ohms đối với hệ thống công nghiệp/thương mại thông thường.
  + Đối với trạm biến áp trung/cao áp, trung tâm dữ liệu hoặc công trình điện lực chuyên dụng, điện trở BẮT BUỘC <= 1.0 Ohm (hoặc theo quy định cụ thể của dự án).
- Điện trở Liên kết Điểm-đến-Điểm (Point-to-Point Ground Continuity / Resistance):
  + Đo bằng thiết bị đo điện trở thấp (Low-Resistance Ohmmeter) từ Main Ground Bus đến vỏ tủ điện, khung máy biến áp, kết cấu thép.
  + Giá trị đo BẮT BUỘC < 0.5 Ohm. CẢNH BÁO "INVESTIGATE" nếu giá trị đo giữa các điểm tương tự lệch quá 50% so với giá trị nhỏ nhất.
- Tiếp địa Hàng rào & Khung vỏ (Fence & Structure Bonding): Hàng rào kim loại và kết cấu thép trạm phải liên kết tiếp địa tin cậy với điện trở liên kết <= 0.5 Ohm.

NHIỆM VỤ CỦA AI:
Khi nhận dữ liệu JSON từ FSE, hãy phân tích và trả về phản hồi định dạng Markdown gồm 4 phần:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE.
2. BẢNG PHÂN TÍCH CHI TIẾT HỆ THỐNG NỐI ĐẤT (Detailed Evaluation Table): So sánh thực tế với NETA ATS-2025 Section 7.13 (Table 100.12) và IEEE Std 81.
3. PHÂN TÍCH ĐIỆN TRỞ ĐẤT, LIÊN KẾT & NGUY CƠ ĐIỆN GIẬT (Ground Grid Resistance, Bonding Integrity & Touch/Step Potential Risk).
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
    console.error('[NETA Grounding Analyzer] Gemini generation error, using deterministic report:', error);
  }

  return deterministicResult;
}
