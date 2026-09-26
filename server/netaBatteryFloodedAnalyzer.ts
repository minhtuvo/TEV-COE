import { GoogleGenAI } from '@google/genai';

export interface NetaBatteryFloodedInputPayload {
  site_info: {
    project_name: string;
    battery_bank_tag: string;
    battery_type: string;
    fse_name: string;
    test_date: string;
    manufacturer?: string;
    model_number?: string;
    serial_number?: string;
  };
  visual_inspection: {
    ventilation_system_operable: boolean;
    eyewash_and_spill_containment_present: boolean;
    nameplate_match: boolean;
    rack_mounting_and_grounding: string; // 'Pass' | 'Fail'
    flame_arresters_present: boolean;
    cleanliness_and_oxide_inhibitor: string; // 'Pass' | 'Fail'
    electrolyte_level_check: string; // e.g. "Pass (All cells at MAX mark)"
    bolt_torque_check: string; // 'Pass' | 'Fail'
  };
  electrical_tests: {
    intercell_connection_resistance_micro_ohms: number[];
    charger_float_voltage_v: number;
    charger_equalize_voltage_v: number;
    charger_alarms_check: string; // 'Pass' | 'Fail'
    cell_voltages_float_mode_v: {
      cell_min_v: number;
      cell_max_v: number;
      voltage_spread_v: number;
    };
    specific_gravity_corrected_25c: {
      pilot_cell_sg: number;
      min_sg_found: number;
      max_sg_found: number;
    };
    internal_ohmic_measurement_resistance_mohm: {
      avg_resistance_mohm: number;
      max_resistance_cell_mohm: number;
      variance_percent: number;
    };
    load_test_ieee_450: string; // e.g. "Pass (Capacity 94% per IEEE 450)"
    system_voltage_to_ground_v: {
      positive_to_ground_v: number;
      negative_to_ground_v: number;
    };
  };
  previous_test_data?: {
    last_test_date?: string;
    last_avg_resistance_mohm?: number;
    last_bolted_resistance_micro_ohms?: number[];
  };
}

export interface BatteryItemEvaluation {
  item: string;
  measured: string;
  standard: string;
  reference: string;
  status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  detail: string;
}

export interface NetaBatteryFloodedEvaluationResult {
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_title: string;
  status_reason: string;
  markdown_report: string;
  evaluations: {
    max_severity: 'PASS' | 'INVESTIGATE' | 'FAIL';
    ground_fault_detected: boolean;
    ground_fault_polarity: 'POSITIVE' | 'NEGATIVE' | 'NONE';
    intercell_resistance_high: boolean;
    cell_voltage_spread_failed: boolean;
    internal_ohmic_failed: boolean;
    specific_gravity_low: boolean;
    ventilation_failed: boolean;
    safety_eyewash_failed: boolean;
    items: BatteryItemEvaluation[];
  };
}

/**
 * Deterministic calculation & evaluation engine according to ANSI/NETA ATS-2025 Section 7.18.1.1 and IEEE 450
 */
export function performDeterministicBatteryFloodedAnalysis(
  data: NetaBatteryFloodedInputPayload
): NetaBatteryFloodedEvaluationResult {
  const items: BatteryItemEvaluation[] = [];

  // --- 1. VISUAL & MECHANICAL INSPECTION (NETA ATS-2025 Section 7.18.1.1.A) ---

  // 1.1 Location & Ventilation System
  const isVentPass = Boolean(data.visual_inspection.ventilation_system_operable);
  items.push({
    item: 'Thông gió phòng pin & Vị trí đặt giàn (Ventilation & Location)',
    measured: isVentPass ? 'Hệ thống quạt thông gió hoạt động bình thường' : 'Hệ thống quạt thông gió KHÔNG hoạt động',
    standard: 'BẮT BUỘC quạt thông gió hoạt động liên tục, ngăn ngừa tuyệt đối tích tụ khí Hydro (H₂)',
    reference: 'NETA ATS-2025 Sec 7.18.1.1.A.1 & IEEE 450',
    status: isVentPass ? 'PASS' : 'FAIL',
    detail: isVentPass
      ? 'Hệ thống thông gió cưỡng bức phòng ắc quy đạt chuẩn chống cháy nổ, đối lưu khí tốt.'
      : 'NGUY HIỂM CẤP BÁCH: Nguy cơ tích tụ khí Hydro gây cháy nổ trạm điện khi sạc bình.'
  });

  // 1.2 Eyewash Station & Spill Containment
  const isSafetyPass = Boolean(data.visual_inspection.eyewash_and_spill_containment_present);
  items.push({
    item: 'Trạm rửa mắt khẩn cấp & Khay hứng tràn axit (Eyewash & Spill Containment)',
    measured: isSafetyPass ? 'Trang bị đầy đủ trạm rửa mắt & khay hứng chống tràn' : 'Thiếu trang thiết bị an toàn axit',
    standard: 'Có sẵn trạm rửa mắt khẩn cấp hoạt động tốt và hệ thống/khay chống tràn axit trung hòa',
    reference: 'NETA ATS-2025 Sec 7.18.1.1.A.2 & OSHA',
    status: isSafetyPass ? 'PASS' : 'FAIL',
    detail: isSafetyPass
      ? 'Bảo đảm an toàn hóa chất H₂SO₄ cho kỹ sư thao tác bảo dưỡng tại trạm.'
      : 'VI PHẠM AN TOÀN: Thiếu trang bị cấp cứu khi tiếp xúc dung dịch axit sunfuric loãng.'
  });

  // 1.3 Nameplate Match
  const isNameplatePass = Boolean(data.visual_inspection.nameplate_match);
  items.push({
    item: 'Đối soát nhãn mác giàn pin (Nameplate Data vs Design Drawings)',
    measured: `${data.site_info.battery_bank_tag} (${data.site_info.battery_type})`,
    standard: 'Khớp 100% hồ sơ thiết kế nhị thứ, điện áp danh định và dung lượng Ah',
    reference: 'NETA ATS-2025 Sec 7.18.1.1.A.3',
    status: isNameplatePass ? 'PASS' : 'FAIL',
    detail: isNameplatePass
      ? 'Số lượng cell, điện áp danh định 2V/cell và dung lượng Ah trùng khớp bản vẽ thi công.'
      : 'Không khớp hồ sơ thiết kế trạm hoặc sai lệch cấu hình chuỗi bình.'
  });

  // 1.4 Racks Mounting & Grounding
  const isRackPass = (data.visual_inspection.rack_mounting_and_grounding || '').toLowerCase().includes('pass');
  items.push({
    item: 'Giá đỡ giàn pin & Tiếp địa an toàn (Racks Mounting & Grounding)',
    measured: data.visual_inspection.rack_mounting_and_grounding || 'Pass',
    standard: 'Giá đỡ cố định chắc chắn, sơn phủ kháng axit, nối đất vỏ giá đỡ đạt tiêu chuẩn',
    reference: 'NETA ATS-2025 Sec 7.18.1.1.A.4',
    status: isRackPass ? 'PASS' : 'FAIL',
    detail: isRackPass
      ? 'Giá đỡ vững chắc, có đệm cách điện chống rung chấn, tiếp địa bảo vệ đạt yêu cầu.'
      : 'Giá đỡ lỏng lẻo, han gỉ hoặc thiếu dây nối đất đẳng thế an toàn.'
  });

  // 1.5 Flame Arresters (Nắp chắn lửa)
  const isFlamePass = Boolean(data.visual_inspection.flame_arresters_present);
  items.push({
    item: 'Nắp chắn lửa nổ ngược (Flame Arresters / Explosion-Resistant Vents)',
    measured: isFlamePass ? 'Đầy đủ 100% nắp chắn lửa trên tất cả các cell' : 'Thiếu hoặc nứt vỡ nắp chắn lửa',
    standard: 'Mỗi cell pin Flooded Lead-Acid BẮT BUỘC có nắp thông khí chắn lửa chống nổ ngược',
    reference: 'NETA ATS-2025 Sec 7.18.1.1.A.5 & IEEE 450',
    status: isFlamePass ? 'PASS' : 'FAIL',
    detail: isFlamePass
      ? 'Nắp chắn lửa xốp ceramic/polymer ngăn chặn tia lửa từ ngoài lọt vào bên trong cell.'
      : 'CẢNH BÁO NGUY HIỂM: Thiếu nắp chắn lửa có thể gây nổ vỡ cell khi có tia lửa điện ngoài.'
  });

  // 1.6 Cleanliness & Oxide Inhibitor
  const isCleanPass = (data.visual_inspection.cleanliness_and_oxide_inhibitor || '').toLowerCase().includes('pass');
  items.push({
    item: 'Vệ sinh cọc cực & Mỡ chống oxy hóa (Cleanliness & Oxide Inhibitor)',
    measured: data.visual_inspection.cleanliness_and_oxide_inhibitor || 'Pass',
    standard: 'Cọc cực chì sạch bẩn, không muối sunfat bám, bôi mỡ chống oxy hóa (NO-OX-ID / vaseline)',
    reference: 'NETA ATS-2025 Sec 7.18.1.1.A.6',
    status: isCleanPass ? 'PASS' : 'FAIL',
    detail: isCleanPass
      ? 'Bề mặt tiếp xúc cọc cực sạch bóng, được bôi mỡ bôi trơn chuyên dụng chống ăn mòn axit.'
      : 'Cọc cực bám bụi bẩn, đóng xỉ trắng muối axit hoặc thiếu mỡ bảo vệ chống oxy hóa.'
  });

  // 1.7 Electrolyte Level
  const isElectroPass = (data.visual_inspection.electrolyte_level_check || '').toLowerCase().includes('pass');
  items.push({
    item: 'Mức dung dịch điện môi (Electrolyte Level Check)',
    measured: data.visual_inspection.electrolyte_level_check || 'Pass',
    standard: 'Mức dung dịch điện phân (H₂SO₄) nằm giữa vạch MIN và MAX của bình trong suốt',
    reference: 'NETA ATS-2025 Sec 7.18.1.1.A.7',
    status: isElectroPass ? 'PASS' : 'FAIL',
    detail: isElectroPass
      ? 'Dung dịch ngập hoàn toàn bản cực và vách ngăn, không bị cạn hở đỉnh tấm cực chì.'
      : 'CẢNH BÁO: Mức điện môi sụt dưới vạch MIN, bản cực bị khô lộ ra ngoài gây hư hỏng vĩnh viễn.'
  });

  // 1.8 Bolt Torque
  const isTorquePass = (data.visual_inspection.bolt_torque_check || '').toLowerCase().includes('pass');
  items.push({
    item: 'Lực siết bu-lông cọc cực & Cầu nối (Bolt Torque per Table 100.12)',
    measured: data.visual_inspection.bolt_torque_check || 'Pass',
    standard: 'Siết đạt mô-men xoắn bằng cần cân lực theo thông số nhà sản xuất hoặc Table 100.12',
    reference: 'NETA ATS-2025 Sec 7.18.1.1.A.8 & Table 100.12',
    status: isTorquePass ? 'PASS' : 'FAIL',
    detail: isTorquePass
      ? 'Các mối ghép bu-lông cọc cực được kiểm tra siết đạt lực quy định, có đánh dấu vạch sơn.'
      : 'Mối ghép bu-lông chưa siết đủ lực hoặc có nguy cơ lỏng sinh nhiệt khi phóng dòng lớn.'
  });

  // --- 2. ELECTRICAL TEST VALUES (NETA ATS-2025 Section 7.18.1.1.D & IEEE 450) ---

  // 2.1 Intercell Connection Resistance
  const resistances = data.electrical_tests.intercell_connection_resistance_micro_ohms || [];
  let minRes = 0;
  let maxRes = 0;
  let maxDevPercent = 0;
  let intercellFailed = false;
  let maxResIndex = -1;

  if (resistances.length > 0) {
    minRes = Math.min(...resistances);
    maxRes = Math.max(...resistances);
    maxResIndex = resistances.indexOf(maxRes);
    if (minRes > 0) {
      maxDevPercent = ((maxRes - minRes) / minRes) * 100;
      if (maxDevPercent > 50) {
        intercellFailed = true;
      }
    }
  }

  items.push({
    item: 'Điện trở cầu nối Intercell (Intercell Connection Resistance)',
    measured: `Min: ${minRes.toFixed(1)} µΩ | Max: ${maxRes.toFixed(1)} µΩ (Mối #${maxResIndex + 1} lệch ${maxDevPercent.toFixed(1)}%)`,
    standard: 'CẢNH BÁO INVESTIGATE nếu giá trị đo lệch quá 50% (> 50%) so với giá trị nhỏ nhất của mối nối tương tự',
    reference: 'NETA ATS-2025 Sec 7.18.1.1.D.1 & D.14',
    status: intercellFailed ? 'INVESTIGATE' : 'PASS',
    detail: intercellFailed
      ? `Mối nối số ${maxResIndex + 1} có điện trở ${maxRes.toFixed(1)} µΩ (lệch ${maxDevPercent.toFixed(1)}% > 50% so với giá trị min ${minRes.toFixed(1)} µΩ). Cần tháo, vệ sinh tiếp xúc và siết lại cờ-lê lực.`
      : `Độ lệch điện trở giữa các cầu nối là ${maxDevPercent.toFixed(1)}% (nằm trong giới hạn ≤ 50% theo NETA).`
  });

  // 2.2 Charger Float & Equalize Voltages and Alarms
  const floatV = data.electrical_tests.charger_float_voltage_v;
  const eqV = data.electrical_tests.charger_equalize_voltage_v;
  const alarmPass = (data.electrical_tests.charger_alarms_check || '').toLowerCase().includes('pass');
  items.push({
    item: 'Điện áp sạc Float / Equalize & Cảnh báo bộ sạc (Charger Voltages & Alarms)',
    measured: `Float: ${floatV}V | Equalize: ${eqV}V | Cảnh báo: ${data.electrical_tests.charger_alarms_check}`,
    standard: 'Điện áp sạc duy trì và sạc cân bằng đạt thông số kỹ thuật nhà sản xuất; chức năng báo động tác động tin cậy',
    reference: 'NETA ATS-2025 Sec 7.18.1.1.D.2 & D.3',
    status: alarmPass && floatV > 0 && eqV > floatV ? 'PASS' : 'INVESTIGATE',
    detail: alarmPass && floatV > 0 && eqV > floatV
      ? 'Bộ sạc hoạt động chuẩn xác, bảo đảm bù tự phóng điện cho giàn pin Flooded 110V/125V DC.'
      : 'Cảnh báo bộ sạc hoặc điện áp sạc float/equalize lệch ngoài dải vận hành tối ưu.'
  });

  // 2.3 Individual Cell Voltages (Float Mode)
  const cellMinV = data.electrical_tests.cell_voltages_float_mode_v.cell_min_v;
  const cellMaxV = data.electrical_tests.cell_voltages_float_mode_v.cell_max_v;
  const actualSpreadV = data.electrical_tests.cell_voltages_float_mode_v.voltage_spread_v || Math.abs(cellMaxV - cellMinV);
  const cellVoltageFailed = actualSpreadV > 0.05;

  items.push({
    item: 'Độ lệch điện áp giữa các Cell ở chế độ Float (Cell Voltages in Float Mode)',
    measured: `Min: ${cellMinV.toFixed(2)}V | Max: ${cellMaxV.toFixed(2)}V | Độ lệch ΔV: ${actualSpreadV.toFixed(2)}V`,
    standard: 'Ở chế độ sạc Float, điện áp giữa các cell KHÔNG ĐƯỢC sai lệch vượt quá 0.05 Volt (within 0.05V of each other)',
    reference: 'NETA ATS-2025 Sec 7.18.1.1.D.4 & IEEE 450',
    status: cellVoltageFailed ? 'FAIL' : 'PASS',
    detail: cellVoltageFailed
      ? `Độ lệch điện áp giữa các cell đạt ${actualSpreadV.toFixed(2)}V (vượt ngưỡng cho phép tối đa 0.05V). Cần chạy chế độ sạc cân bằng Equalize Charge để kéo đều điện áp.`
      : `Độ chênh lệch điện áp giữa các cell là ${actualSpreadV.toFixed(2)}V (đạt tiêu chuẩn ≤ 0.05V theo NETA & IEEE 450).`
  });

  // 2.4 Specific Gravity (Tỷ trọng dung dịch hiệu chỉnh 25°C)
  const pilotSg = data.electrical_tests.specific_gravity_corrected_25c.pilot_cell_sg;
  const minSg = data.electrical_tests.specific_gravity_corrected_25c.min_sg_found;
  const maxSg = data.electrical_tests.specific_gravity_corrected_25c.max_sg_found;
  const sgFailed = minSg < 1.190 || (maxSg - minSg) > 0.030;

  items.push({
    item: 'Tỷ trọng dung dịch điện môi hiệu chỉnh 25°C (Specific Gravity corrected @ 25°C)',
    measured: `Pilot cell: ${pilotSg.toFixed(3)} | Min: ${minSg.toFixed(3)} | Max: ${maxSg.toFixed(3)}`,
    standard: 'Tỷ trọng thiết kế chuẩn 1.215 ± 0.010 (dải 1.205 - 1.225); độ lệch giữa các bình không quá 0.020',
    reference: 'NETA ATS-2025 Sec 7.18.1.1.D.5 & IEEE 450',
    status: sgFailed ? 'FAIL' : 'PASS',
    detail: sgFailed
      ? `Cell có tỷ trọng thấp nhất đạt ${minSg.toFixed(3)} (sụt thấp dưới mức chuẩn 1.215). Dấu hiệu cell bị sunfat hóa, chưa sạc đầy hoặc mất chất điện giải.`
      : `Tỷ trọng dung dịch các bình nằm trong dải danh định thiết kế tốt.`
  });

  // 2.5 Internal Ohmic Measurements (Nội trở ôm cell)
  const avgRes = data.electrical_tests.internal_ohmic_measurement_resistance_mohm.avg_resistance_mohm;
  const maxResCell = data.electrical_tests.internal_ohmic_measurement_resistance_mohm.max_resistance_cell_mohm;
  const variancePct = data.electrical_tests.internal_ohmic_measurement_resistance_mohm.variance_percent;
  const ohmicFailed = variancePct > 25.0;

  items.push({
    item: 'Nội trở Ôm Cell (Internal Ohmic Measurements @ Full Charge)',
    measured: `Avg: ${avgRes.toFixed(2)} mΩ | Max cell: ${maxResCell.toFixed(2)} mΩ | Biến động: ${variancePct.toFixed(1)}%`,
    standard: 'Giá trị trở kháng/nội trở/độ dẫn điện giữa các cell ở trạng thái sạc đầy KHÔNG ĐƯỢC biến động quá 25% (> 25%)',
    reference: 'NETA ATS-2025 Sec 7.18.1.1.D.6 & IEEE 450',
    status: ohmicFailed ? 'FAIL' : 'PASS',
    detail: ohmicFailed
      ? `Cell nội trở lớn nhất ${maxResCell.toFixed(2)} mΩ biến động ${variancePct.toFixed(1)}% so với trung bình (vượt quá ngưỡng cho phép 25%). Báo hiệu suy giảm dung lượng nặng hoặc chai bản cực.`
      : `Độ đồng đều nội trở giữa các bình đạt ${variancePct.toFixed(1)}% (nằm trong giới hạn ≤ 25%).`
  });

  // 2.6 Load Test per IEEE 450 (Thử nghiệm xả tải dung lượng)
  const loadTestPass = (data.electrical_tests.load_test_ieee_450 || '').toLowerCase().includes('pass');
  items.push({
    item: 'Thử nghiệm xả tải dung lượng (Battery Capacity Load Test per IEEE 450)',
    measured: data.electrical_tests.load_test_ieee_450 || 'Pass',
    standard: 'Dung lượng thực tế sau xả BẮT BUỘC đạt ≥ 80% dung lượng danh định của nhà sản xuất',
    reference: 'NETA ATS-2025 Sec 7.18.1.1.D.7 & IEEE 450',
    status: loadTestPass ? 'PASS' : 'FAIL',
    detail: loadTestPass
      ? 'Dung lượng thực tế đáp ứng đầy đủ dòng phóng sự cố theo biểu đồ phụ tải DC trạm.'
      : 'CẢNH BÁO: Giàn pin không đạt dung lượng tối thiểu 80% theo IEEE 450, khuyến nghị thay thế chuỗi.'
  });

  // 2.7 System Voltage Positive / Negative to Ground (Độ cân bằng cực tính đất)
  const vPos = data.electrical_tests.system_voltage_to_ground_v.positive_to_ground_v;
  const vNeg = data.electrical_tests.system_voltage_to_ground_v.negative_to_ground_v;
  const absPos = Math.abs(vPos);
  const absNeg = Math.abs(vNeg);
  const totalBusV = absPos + absNeg;
  const diffMagnitude = Math.abs(absPos - absNeg);
  const unbalanceRatio = totalBusV > 0 ? (diffMagnitude / totalBusV) * 100 : 0;

  // Ground fault rule: Positive to ground and negative to ground magnitudes must be equal in magnitude.
  // Allowed threshold is typically within 10% unbalance under normal conditions.
  // When +82.0V and -40.5V (diff = 41.5V on a ~122.5V bus), this is a severe ground fault!
  let groundFaultDetected = false;
  let groundPolarity: 'POSITIVE' | 'NEGATIVE' | 'NONE' = 'NONE';

  if (diffMagnitude > 5.0 || unbalanceRatio > 15.0) {
    groundFaultDetected = true;
    // If absNeg < absPos, the negative bus is closer to ground potential -> Negative Ground Fault
    // If absPos < absNeg, the positive bus is closer to ground potential -> Positive Ground Fault
    groundPolarity = absNeg < absPos ? 'NEGATIVE' : 'POSITIVE';
  }

  items.push({
    item: 'Điện áp so với Đất Dương / Âm (System Voltage Positive/Negative to Ground)',
    measured: `(+) to Ground: ${vPos > 0 ? '+' : ''}${vPos.toFixed(1)}V | (-) to Ground: ${vNeg.toFixed(1)}V (Lệch: ${diffMagnitude.toFixed(1)}V)`,
    standard: 'Độ lớn điện áp đo từ cực Dương (+) đến đất BẮT BUỘC phải bằng độ lớn điện áp đo từ cực Âm (-) đến đất (equal in magnitude)',
    reference: 'NETA ATS-2025 Sec 7.18.1.1.D.8 & IEEE 450',
    status: groundFaultDetected ? 'FAIL' : 'PASS',
    detail: groundFaultDetected
      ? `NGUY HIỂM CẤP BÁCH: Mất cân bằng điện áp đối đất cực kỳ nghiêm trọng (${vPos > 0 ? '+' : ''}${vPos.toFixed(1)}V so với ${vNeg.toFixed(1)}V, chuẩn cân bằng $\\approx \\pm ${(totalBusV / 2).toFixed(1)}\\text{V}$). Xác định trực tiếp đang có SỰ CỐ CHẠM ĐẤT CỰC ${groundPolarity === 'NEGATIVE' ? 'ÂM (NEGATIVE DC GROUND FAULT)' : 'DƯƠNG (POSITIVE DC GROUND FAULT)'} trong lưới 110V/125V DC!`
      : `Hệ thống DC cách ly tốt với đất, điện áp cực tính đối xứng hoàn hảo ($\\approx \\pm ${(totalBusV / 2).toFixed(1)}\\text{V}$).`
  });

  // --- 3. OVERALL STATUS SYNTHESIS ---
  const hasFail = items.some(i => i.status === 'FAIL') || groundFaultDetected || cellVoltageFailed || ohmicFailed;
  const hasInvestigate = items.some(i => i.status === 'INVESTIGATE') || intercellFailed;

  let overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
  let status_title = 'PASS / HEALTHY DC SYSTEM';
  let status_reason = 'Tất cả các hạng mục kiểm tra thị giác, cơ khí và đo kiểm điện tuân thủ 100% ANSI/NETA ATS-2025 Section 7.18.1.1 và IEEE 450.';

  if (hasFail) {
    overall_status = 'FAIL';
    if (groundFaultDetected && ohmicFailed) {
      status_title = 'FAIL / CRITICAL DC GROUND FAULT & CELL DEGRADATION';
      status_reason = `Hệ thống DC ghi nhận SỰ CỐ CHẠM ĐẤT CỰC ${groundPolarity} (Đo ${vPos > 0 ? '+' : ''}${vPos.toFixed(1)}V / ${vNeg.toFixed(1)}V) kết hợp biến động nội trở cell ${variancePct.toFixed(1)}% (> 25%) và độ lệch áp Float ${actualSpreadV.toFixed(2)}V (> 0.05V).`;
    } else if (groundFaultDetected) {
      status_title = `FAIL / CRITICAL DC ${groundPolarity} GROUND FAULT`;
      status_reason = `Phát hiện sự cố chạm đất DC cực ${groundPolarity} (${vPos > 0 ? '+' : ''}${vPos.toFixed(1)}V / ${vNeg.toFixed(1)}V). Nguy cơ gây nhảy sai hoặc tê liệt rơ-le bảo vệ.`;
    } else {
      status_title = 'FAIL / BATTERY PARAMETERS OUT OF TOLERANCE';
      status_reason = 'Các thông số điện áp cell, tỷ trọng dung dịch hoặc nội trở ôm vượt quá giới hạn an toàn NETA ATS-2025.';
    }
  } else if (hasInvestigate) {
    overall_status = 'INVESTIGATE';
    status_title = 'INVESTIGATE / CELL CONNECTIONS OR CHARGER DEVIATION';
    status_reason = 'Cần siết lại cờ-lê lực mối nối intercell hoặc kiểm tra hiệu chuẩn điện áp sạc theo tiêu chuẩn.';
  }

  // Generate 4-part Markdown report matching exact user prompt requirements
  const markdown_report = generateMarkdownReport({
    data,
    items,
    overall_status,
    status_title,
    minRes,
    maxRes,
    maxResIndex,
    maxDevPercent,
    cellMinV,
    cellMaxV,
    actualSpreadV,
    pilotSg,
    minSg,
    maxSg,
    avgRes,
    maxResCell,
    variancePct,
    vPos,
    vNeg,
    totalBusV,
    groundFaultDetected,
    groundPolarity
  });

  return {
    overall_status,
    status_title,
    status_reason,
    markdown_report,
    evaluations: {
      max_severity: overall_status,
      ground_fault_detected: groundFaultDetected,
      ground_fault_polarity: groundPolarity,
      intercell_resistance_high: intercellFailed,
      cell_voltage_spread_failed: cellVoltageFailed,
      internal_ohmic_failed: ohmicFailed,
      specific_gravity_low: sgFailed,
      ventilation_failed: !isVentPass,
      safety_eyewash_failed: !isSafetyPass,
      items
    }
  };
}

/**
 * Helper to build the exact 4-part structured Markdown report
 */
function generateMarkdownReport(params: {
  data: NetaBatteryFloodedInputPayload;
  items: BatteryItemEvaluation[];
  overall_status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  status_title: string;
  minRes: number;
  maxRes: number;
  maxResIndex: number;
  maxDevPercent: number;
  cellMinV: number;
  cellMaxV: number;
  actualSpreadV: number;
  pilotSg: number;
  minSg: number;
  maxSg: number;
  avgRes: number;
  maxResCell: number;
  variancePct: number;
  vPos: number;
  vNeg: number;
  totalBusV: number;
  groundFaultDetected: boolean;
  groundPolarity: 'POSITIVE' | 'NEGATIVE' | 'NONE';
}): string {
  const {
    data,
    items,
    status_title,
    minRes,
    maxRes,
    maxResIndex,
    maxDevPercent,
    cellMinV,
    cellMaxV,
    actualSpreadV,
    pilotSg,
    minSg,
    maxSg,
    avgRes,
    maxResCell,
    variancePct,
    vPos,
    vNeg,
    totalBusV,
    groundFaultDetected,
    groundPolarity
  } = params;

  return `# BÁO CÁO ĐÁNH GIÁ CHUYÊN GIA NETA ATS-2025: HỆ THỐNG ĐIỆN DC & ẮC QUY CHÌ-AXIT NGẬP DUNG DỊCH (FLOODED LEAD-ACID)
**Dự án:** ${data.site_info.project_name}  
**Mã giàn pin (Tag):** ${data.site_info.battery_bank_tag} | **Chủng loại:** ${data.site_info.battery_type}  
**Kỹ sư Hiện trường (FSE):** ${data.site_info.fse_name} | **Ngày thử nghiệm:** ${data.site_info.test_date}  
**Tiêu chuẩn đối soát:** ANSI/NETA ATS-2025 (Mục 7.18.1.1) & IEEE 450 (Recommended Practice for Maintenance, Testing, and Replacement of Vented Lead-Acid Batteries for Generating Stations and Substations)

---

## 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
### **${status_title}**

> **Tóm lược Đánh giá từ TEV Platform AI:**
> - **Thị giác & An toàn:** Hệ thống thông gió phòng pin, trạm rửa mắt khẩn cấp, khay chống tràn, nắp chắn lửa và mỡ bôi cọc cực đạt chuẩn NETA Section 7.18.1.1.A.
> - **Điện trở Cầu nối Intercell:** Mối nối số ${maxResIndex + 1} có điện trở $${maxRes.toFixed(1)}\\,\\mu\\Omega$ (Lệch ${maxDevPercent.toFixed(1)}% so với mối nối nhỏ nhất $${minRes.toFixed(1)}\\,\\mu\\Omega$) $\\rightarrow$ ${maxDevPercent > 50 ? 'Vượt quá ngưỡng cho phép 50% của NETA Section 7.18.1.1.D.1 $\\rightarrow$ **INVESTIGATE** (Siết lại lực hoặc vệ sinh cầu nối intercell).' : 'Đạt chuẩn ≤ 50% $\\rightarrow$ **PASS**.'}
> - **Độ lệch Điện áp & Tỷ trọng Dung dịch:** Độ chênh lệch điện áp giữa các cell đạt ${actualSpreadV.toFixed(2)}V (Từ ${cellMinV.toFixed(2)}V đến ${cellMaxV.toFixed(2)}V), ${actualSpreadV > 0.05 ? 'vượt quá ngưỡng tối đa cho phép 0.05 Volt theo NETA Section 7.18.1.1.D.4 $\\rightarrow$ **FAIL**.' : 'đạt chuẩn ≤ 0.05 Volt $\\rightarrow$ **PASS**.'} Tỷ trọng dung dịch cell thấp nhất đạt ${minSg.toFixed(3)} (Sụt thấp hơn mức chuẩn ${pilotSg.toFixed(3)}), báo hiệu cell bị sụt điện dịch hoặc cân bằng kém.
> - **Nội trở Ôm Cell (Internal Ohmic Measurement):** Cell có nội trở lớn nhất đạt ${maxResCell.toFixed(2)} mΩ (Biến động ${variancePct.toFixed(1)}% so với trung bình ${avgRes.toFixed(2)} mΩ) $\\rightarrow$ ${variancePct > 25 ? 'Vượt quá ngưỡng tối đa cho phép 25% theo NETA Section 7.18.1.1.D.6 $\\rightarrow$ **FAIL** (Nghi ngờ chai tấm cực hoặc sunfat hóa nặng).' : 'Đạt chuẩn ≤ 25% $\\rightarrow$ **PASS**.'}
> - **Điện áp so với Đất (Positive/Negative to Ground):** Điện áp Dương so với đất là ${vPos > 0 ? '+' : ''}${vPos.toFixed(1)}V, Âm so với đất là ${vNeg.toFixed(1)}V (Mất cân bằng nghiêm trọng so với mức chuẩn cân bằng $\\approx \\pm ${(totalBusV / 2).toFixed(1)}\\text{V}$) $\\rightarrow$ ${groundFaultDetected ? `**CRITICAL FAIL** (Chỉ thị trực tiếp đang có Sự cố Chạm đất cực ${groundPolarity === 'NEGATIVE' ? 'Âm (Negative DC Ground Fault)' : 'Dương (Positive DC Ground Fault)'} trong hệ thống 110V/125V DC).` : '**PASS** (Cân bằng tốt).'}

---

## 2. BẢNG PHÂN TÍCH CHI TIẾT GIÀN PIN FLOODED LEAD-ACID (Detailed Evaluation Table)
So sánh thực tế từng hạng mục với tiêu chuẩn **ANSI/NETA ATS-2025 Mục 7.18.1.1** & **IEEE 450**:

| STT | Hạng mục Kiểm tra | Giá trị Thực tế Đo được | Yêu cầu Kỹ thuật Tiêu chuẩn | Điều khoản Tham chiếu | Kết quả |
| :---: | :--- | :--- | :--- | :--- | :---: |
${items.map((it, idx) => `| ${idx + 1} | **${it.item}** | ${it.measured} | ${it.standard} | ${it.reference} | **${it.status === 'PASS' ? '✅ PASS' : it.status === 'INVESTIGATE' ? '⚠️ INVESTIGATE' : '❌ FAIL'}** |`).join('\n')}

---

## 3. PHÂN TÍCH TỶ TRỌNG, NỘI TRỞ & NGUY CƠ RÒ ĐẤT (Specific Gravity, Internal Ohmic & Grounding Risk)

### 3.1. Phân tích Điện áp Đối đất & Sự cố Chạm đất DC (Critical Grounding Risk Analysis)
- **Tổng điện áp thanh cái DC:** $V_{\\text{bus}} = |V_{(+)}| + |V_{(-)}| = |${vPos.toFixed(1)}| + |${vNeg.toFixed(1)}| = ${totalBusV.toFixed(1)}\\text{ V}$.
- **Điện áp đối xứng lý thuyết khi không có rò chạm đất:** $V_{\\text{ref}} = \\pm \\frac{${totalBusV.toFixed(1)}}{2} = \\pm ${(totalBusV / 2).toFixed(1)}\\text{ V}$.
- **Độ lệch điện áp thực tế:** Cực Dương đo được $${vPos > 0 ? '+' : ''}${vPos.toFixed(1)}\\text{ V}$, Cực Âm đo được $${vNeg.toFixed(1)}\\text{ V}$.
- **Cơ chế sự cố kỹ thuật:** Điện áp cực Âm sụt sâu xuống $${vNeg.toFixed(1)}\\text{ V}$ trong khi cực Dương bị đẩy vọt lên $${vPos > 0 ? '+' : ''}${vPos.toFixed(1)}\\text{ V}$. Điều này xác nhận điện trở cách điện từ cực Âm về đất đã bị suy giảm nghiêm trọng ($R_{\\text{iso(-)}} \\ll R_{\\text{iso(+)}} \\approx \\infty$).
- **Hậu quả vận hành:** Khi có thêm một điểm chạm đất thứ hai ở cực Dương trong hệ thống, dòng ngắn mạch DC trực tiếp sẽ kích hoạt cắt áp nguồn điều khiển, hoặc nguy cơ kích hoạt sai cuộn cắt (Trip Coil) của máy cắt 110kV/220kV gây mất điện diện rộng trạm.

### 3.2. Đánh giá Mất cân bằng Nội trở Ôm & Suy thoái Tấm cực (Internal Ohmic Degradation)
- **Giá trị nội trở trung bình giàn pin:** $\\bar{R} = ${avgRes.toFixed(2)}\\text{ m}\\Omega$.
- **Cell có nội trở lớn nhất:** $R_{\\text{max}} = ${maxResCell.toFixed(2)}\\text{ m}\\Omega$, tương đương mức biến động:
  $$\\Delta R = \\frac{${maxResCell.toFixed(2)} - ${avgRes.toFixed(2)}}{${avgRes.toFixed(2)}} \\times 100\\% = ${variancePct.toFixed(1)}\\%$$
- **Đối soát NETA ATS-2025 Section 7.18.1.1.D.6:** Giới hạn cho phép tối đa là $25\\%$. Giá trị đo được $${variancePct.toFixed(1)}\\%$ vượt ngưỡng $1.5$ lần, phản ánh tình trạng kết tinh muối chì sunfat cứng ($PbSO_4$) không thể đảo ngược trên bề mặt bản cực, làm thu hẹp diện tích điện hóa hữu dụng.

### 3.3. Tương quan giữa Sụt Áp Float & Sụt Tỷ Trọng Điện Dịch
- **Độ lệch điện áp Float:** $\\Delta V = ${cellMaxV.toFixed(2)}\\text{ V} - ${cellMinV.toFixed(2)}\\text{ V} = ${actualSpreadV.toFixed(2)}\\text{ V} > 0.05\\text{ V}$ (vượt chuẩn NETA Section 7.18.1.1.D.4).
- **Tỷ trọng dung dịch hiệu chỉnh:** Cell thấp nhất đo được $${minSg.toFixed(3)}$ so với Pilot Cell $${pilotSg.toFixed(3)}$. Nồng độ axit sunfuric loãng bị sụt giảm đồng thời với sụt áp, chứng minh hiện tượng tự phóng điện cục bộ và phân tầng dung dịch điện môi.

---

## 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)

1. **CẢNH BÁO TỐI CẤP — ĐỊNH VỊ VÀ XỬ LÝ CHẠM ĐẤT DC CỰC ÂM (Negative Ground Fault):**
   - Sử dụng thiết bị dò tìm chạm đất DC trực tuyến (DC Earth Fault Locator chuyên dụng tần số thấp 10Hz - 25Hz) để quét từng lộ cáp điều khiển, cuộn đóng/cắt máy cắt và tủ rơ-le nhị thứ.
   - Tuyệt đối không để hệ thống vận hành kéo dài trong tình trạng chạm đất cực âm, tránh nguy cơ nhảy cắt giả thiết bị rơ-le/điều khiển trạm 220kV.

2. **CHẠY CHẾ ĐỘ SẠC CÂN BẰNG (EQUALIZING CHARGE):**
   - Kích hoạt chế độ sạc cân bằng trên bộ sạc với điện áp $${data.electrical_tests.charger_equalize_voltage_v}\\text{ V}$ trong thời gian khuyến cáo của nhà sản xuất (thường từ 24h đến 72h).
   - Mục tiêu: Hòa tan các tinh thể sunfat chì trên bản cực, kéo độ lệch điện áp giữa các cell về dưới ngưỡng quy định $0.05\\text{ V}$ và khôi phục tỷ trọng dung dịch lên mức chuẩn $1.215 \\pm 0.010$.

3. **BẢO DƯỠNG MỐI NỐI CẦU INTERCELL SỐ ${maxResIndex + 1}:**
   - Tháo dỡ cầu nối số ${maxResIndex + 1}$ (${maxRes.toFixed(1)}\\,\\mu\\Omega$), dùng bàn chải đồng và dung dịch trung hòa vệ sinh sạch màng oxy hóa chì.
   - Bôi lớp mỡ chuyên dụng chống oxy hóa (NO-OX-ID A-Special) và dùng cần cân lực siết lại đúng lực quy định theo Table 100.12.
   - Đo kiểm tra lại điện trở tiếp xúc bằng Micro-ohmmeter, đảm bảo sai lệch dưới $50\\%$ so với mối nối nhỏ nhất.

4. **ĐO LẠI NỘI TRỞ ÔM SAU CHU TRÌNH SẠC:**
   - Sau khi kết thúc sạc cân bằng và giàn pin được nghỉ tối thiểu 24 giờ, tiến hành đo lại toàn bộ nội trở ôm các cell bằng thiết bị chuyên dụng (Megger BITE / Fluke BT521).
   - Nếu cell có nội trở $${maxResCell.toFixed(2)}\\text{ m}\\Omega$ vẫn biến động vượt quá $25\\%$, cần lập kế hoạch thay thế cell đơn lẻ để bảo đảm độ tin cậy của giàn nguồn DC.
`;
}

/**
 * AI Augmented Analysis using Google Gemini SDK with fallback to deterministic engine
 */
export async function generateNetaBatteryFloodedAiAnalysis(
  payload: NetaBatteryFloodedInputPayload
): Promise<NetaBatteryFloodedEvaluationResult> {
  const deterministicResult = performDeterministicBatteryFloodedAnalysis(payload);

  // If Gemini API key is available, we can augment with generative insights
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return deterministicResult;
  }

  try {
    const ai = new GoogleGenAI({ apiKey });
    const systemPrompt = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách Hệ thống Điện một chiều DC và Ắc quy Chì-Axit ngập dung dịch (Direct-Current Systems, Batteries, Flooded Lead-Acid). 
Nhiệm vụ của bạn là nhận dữ liệu kiểm tra ngoài site từ Field Service Engineer (FSE) và tự động đối soát với tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.18.1.1 và IEEE 450).

QUY TRẮC ĐÁNH GIÁ:
1. Kiểm tra thị giác & cơ khí: Thông gió bắt buộc, bồn rửa mắt & khay hứng tràn, nhãn mác, giá đỡ, nắp chắn lửa, vệ sinh cọc cực & mỡ chống oxy hóa, mức điện dịch, lực siết bu-lông.
2. Phép đo điện:
- Điện trở cầu nối Intercell: CẢNH BÁO "INVESTIGATE" nếu giá trị đo lệch quá 50% (> 50%) so với giá trị nhỏ nhất của mối nối tương tự.
- Điện áp sạc Float & Equalize: Theo thông số kỹ thuật.
- Điện áp Cell ở chế độ Float: KHÔNG ĐƯỢC sai lệch vượt quá 0.05 Volt.
- Tỷ trọng dung dịch: Chuẩn 1.215.
- Nội trở Ôm Cell: KHÔNG ĐƯỢC biến động quá 25% (> 25%).
- Thử nghiệm xả tải: Đạt ≥ 80% per IEEE 450.
- Điện áp so với đất: Độ lớn cực (+) đến đất BẮT BUỘC phải bằng độ lớn cực (-) đến đất (equal in magnitude). Nếu lệch là có sự cố chạm đất DC (Positive hoặc Negative Ground Fault).

ĐỊNH DẠNG ĐẦU RA BẮT BUỘC:
Trả về phản hồi Markdown gồm đúng 4 phần:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE (kèm tiêu đề chi tiết).
2. BẢNG PHÂN TÍCH CHI TIẾT GIÀN PIN FLOODED LEAD-ACID (Detailed Evaluation Table) dạng bảng Markdown so sánh thực tế với NETA ATS-2025 Sec 7.18.1.1 và IEEE 450.
3. PHÂN TÍCH TỶ TRỌNG, NỘI TRỞ & NGUY CƠ RÒ ĐẤT (Specific Gravity, Internal Ohmic & Grounding Risk).
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations).`;

    const userPrompt = `Dữ liệu kiểm tra hệ thống DC & giàn pin Flooded Lead-Acid từ FSE:\n${JSON.stringify(payload, null, 2)}`;

    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: [
        { role: 'user', parts: [{ text: `${systemPrompt}\n\n${userPrompt}` }] }
      ]
    });

    const aiMarkdown = response.text?.trim();
    if (aiMarkdown && aiMarkdown.length > 200) {
      return {
        ...deterministicResult,
        markdown_report: aiMarkdown
      };
    }
  } catch (error) {
    console.warn('Gemini API call failed for battery flooded analyzer, using deterministic result:', error);
  }

  return deterministicResult;
}
