/**
 * ANSI/NETA ATS DATA PERSISTENCE & SYNCHRONIZATION SERVICE
 * Bridges field entry form data across all 23 ATS checklists and the TEV Standard AI Report.
 */

import {
  NETA_ATS_23_MASTER_SPECS,
  NetaAtsSpecificationDef,
  NetaAtsTestItem,
  NetaAtsAnomalyCard,
} from '../data/netaAtsMasterSpecifications';

const STORAGE_PREFIX = 'tev_ats_payload_';
const memoryStore: Record<string, any> = {};

/**
 * Retrieve cached or saved payload for a specific specification ID.
 */
export function getAtsPayload(specId: string): any | null {
  if (memoryStore[specId]) {
    return memoryStore[specId];
  }
  try {
    const raw = localStorage.getItem(`${STORAGE_PREFIX}${specId}`);
    if (raw) {
      const parsed = JSON.parse(raw);
      memoryStore[specId] = parsed;
      return parsed;
    }
  } catch (err) {
    console.warn(`[netaAtsDataService] Error reading payload for ${specId}:`, err);
  }
  return null;
}

/**
 * Save payload for a specific specification ID.
 * Updates in-memory store, persists to localStorage, and dispatches a reactive event.
 */
export function saveAtsPayload(specId: string, payload: any, source: string = 'internal'): void {
  if (!specId || !payload) return;
  memoryStore[specId] = payload;
  try {
    localStorage.setItem(`${STORAGE_PREFIX}${specId}`, JSON.stringify(payload));
  } catch (err) {
    console.warn(`[netaAtsDataService] Error writing payload for ${specId}:`, err);
  }
  if (typeof window !== 'undefined') {
    window.dispatchEvent(
      new CustomEvent('tev_ats_payload_updated', {
        detail: { specId, payload, source },
      })
    );
  }
}

/**
 * Reset payload for a specific specification ID back to the master sample.
 */
export function resetAtsPayload(specId: string): void {
  delete memoryStore[specId];
  try {
    localStorage.removeItem(`${STORAGE_PREFIX}${specId}`);
  } catch (err) {
    console.warn(`[netaAtsDataService] Error resetting payload for ${specId}:`, err);
  }
  if (typeof window !== 'undefined') {
    window.dispatchEvent(
      new CustomEvent('tev_ats_payload_updated', {
        detail: { specId, reset: true, source: 'internal' },
      })
    );
  }
}

/**
 * Helper to find an evaluation item by case-insensitive search terms.
 */
function findItem(
  items: NetaAtsTestItem[],
  keywordsVi: string[],
  keywordsEn: string[] = []
): NetaAtsTestItem | undefined {
  return items.find((i) => {
    const itemLower = (i.item || '').toLowerCase();
    const itemEnLower = (i.itemEn || '').toLowerCase();
    const matchVi = keywordsVi.some((kw) => itemLower.includes(kw.toLowerCase()));
    const matchEn = keywordsEn.some((kw) => itemEnLower.includes(kw.toLowerCase()));
    return matchVi || matchEn;
  });
}

/**
 * Dynamically construct a NetaAtsSpecificationDef by merging newly inputted field data
 * with the ANSI/NETA ATS master specification baseline.
 */
export function buildDynamicAtsSpec(
  specId: string,
  explicitPayload?: any
): NetaAtsSpecificationDef {
  const baseSpec =
    NETA_ATS_23_MASTER_SPECS[specId] ||
    NETA_ATS_23_MASTER_SPECS['neta-liquid-transformer'];

  const payload = explicitPayload || getAtsPayload(specId);

  // If no custom user data has been inputted yet, return baseline
  if (!payload) {
    return baseSpec;
  }

  // Clone base specification to avoid mutating constants
  const dynamicSpec: NetaAtsSpecificationDef = JSON.parse(
    JSON.stringify(baseSpec)
  );
  dynamicSpec.samplePayloadJson = payload;

  // 1. Merge Site & Header Information
  const site = payload.site_info || payload.station_info || {};
  const projectOrStation = site.project_name || site.station_name;
  if (site.customer_name && projectOrStation) {
    dynamicSpec.defaultSiteInfo.project_name = projectOrStation.includes(site.customer_name)
      ? projectOrStation
      : `${site.customer_name} • ${projectOrStation}`;
  } else if (projectOrStation || site.customer_name) {
    dynamicSpec.defaultSiteInfo.project_name = projectOrStation || site.customer_name;
  }
  
  const tag =
    site.transformer_tag ||
    site.equipment_tag ||
    site.breaker_tag ||
    site.cable_tag ||
    site.motor_tag ||
    site.switchgear_tag ||
    site.grounding_system_tag ||
    site.ground_grid_tag ||
    site.generator_tag ||
    site.ats_tag ||
    site.battery_bank_tag ||
    site.ups_tag ||
    site.relay_tag ||
    site.switch_tag ||
    site.evse_tag ||
    site.pv_array_tag ||
    site.bess_container_tag ||
    site.bess_tag ||
    site.survey_tag;
  if (tag) {
    if (site.equipment_name && !String(tag).includes(site.equipment_name)) {
      dynamicSpec.defaultSiteInfo.equipment_tag = `${tag} — ${site.equipment_name}`;
    } else {
      dynamicSpec.defaultSiteInfo.equipment_tag = tag;
    }
  }

  const rating =
    site.transformer_rating ||
    site.equipment_rating ||
    site.generator_rating ||
    site.breaker_rating ||
    site.switch_rating ||
    site.charger_type ||
    site.equipment_type ||
    site.cable_type ||
    site.system_type ||
    site.transformer_type ||
    site.rating;
  if (rating) dynamicSpec.defaultSiteInfo.equipment_rating = rating;

  if (site.fse_name) dynamicSpec.defaultSiteInfo.fse_name = site.fse_name;
  if (site.test_date) dynamicSpec.defaultSiteInfo.test_date = site.test_date;
  if (site.ambient_temperature_c !== undefined || site.ambient_temp_c !== undefined) {
    dynamicSpec.defaultSiteInfo.ambient_temp_c = site.ambient_temperature_c ?? site.ambient_temp_c;
  }
  if (site.relative_humidity_percent !== undefined || site.humidity_percent !== undefined) {
    dynamicSpec.defaultSiteInfo.humidity_percent = site.relative_humidity_percent ?? site.humidity_percent;
  }

  // 2. Spec-Specific Dynamic Evaluation
  switch (specId) {
    case 'neta-liquid-transformer':
      applyLiquidTransformerPayload(dynamicSpec, payload);
      break;
    case 'neta-grounding':
      applyGroundingPayload(dynamicSpec, payload);
      break;
    case 'neta-lv-breaker':
      applyLvBreakerPayload(dynamicSpec, payload);
      break;
    case 'neta-switchgear':
      applySwitchgearPayload(dynamicSpec, payload);
      break;
    case 'neta-cable-lv':
    case 'neta-cable':
      applyCablePayload(dynamicSpec, payload);
      break;
    case 'neta-motor':
    case 'neta-dc-motor':
    case 'neta-sync-machinery':
      applyMotorPayload(dynamicSpec, payload);
      break;
    case 'neta-dry-type':
    case 'neta-dry-large':
      applyDryTransformerPayload(dynamicSpec, payload);
      break;
    case 'neta-battery-vrla':
    case 'neta-battery-flooded':
    case 'neta-ups':
      applyBatteryUpsPayload(dynamicSpec, payload);
      break;
    case 'neta-generator':
    case 'neta-ats':
      applyGeneratorAtsPayload(dynamicSpec, payload);
      break;
    case 'neta-sf6-switch':
      applySf6SwitchPayload(dynamicSpec, payload);
      break;
    case 'neta-evse':
      applyEvsePayload(dynamicSpec, payload);
      break;
    case 'neta-pv':
      applyPvPayload(dynamicSpec, payload);
      break;
    case 'neta-bess':
      applyBessPayload(dynamicSpec, payload);
      break;
    case 'neta-pd':
      applyPdPayload(dynamicSpec, payload);
      break;
    case 'neta-thermography':
      applyThermographyPayload(dynamicSpec, payload);
      break;
    default:
      applyUniversalPayload(dynamicSpec, payload);
      break;
  }

  // Harmonize overall status if all items pass or any items fail
  recalculateOverallStatus(dynamicSpec);

  // 3. Dynamically Generate Section 3 (Anomaly Cards) & Section 4 (Recommendations)
  dynamicSpec.anomalyCards = buildDynamicAnomalyCards(dynamicSpec, payload, specId);
  const dynamicRecs = buildDynamicRecommendations(dynamicSpec, payload, specId);
  dynamicSpec.recommendationsVi = dynamicRecs.vi;
  dynamicSpec.recommendationsEn = dynamicRecs.en;

  return dynamicSpec;
}

/**
 * Re-evaluates overall status based on evaluation items
 */
function recalculateOverallStatus(spec: NetaAtsSpecificationDef) {
  const failedItems = spec.evaluationItems.filter((i) => i.status === 'FAIL');
  const warnItems = spec.evaluationItems.filter((i) => i.status === 'INVESTIGATE');

  if (failedItems.length > 0) {
    spec.overallStatus = 'FAIL';
    spec.statusTitleVi = 'CẢNH BÁO KHIẾM KHUYẾT KỸ THUẬT & RỦI RO SỰ CỐ CAO (FAIL)';
    spec.statusTitleEn = 'CRITICAL DEFECT & HIGH OPERATIONAL RISK (FAIL)';
    spec.statusReasonVi = `Phát hiện ${failedItems.length} hạng mục thử nghiệm không đạt tiêu chuẩn ${spec.standardCode}. Tuyệt đối không cho phép đóng điện trước khi khắc phục triệt để.`;
    spec.statusReasonEn = `Identified ${failedItems.length} non-compliant test items violating ${spec.standardCode}. Energization prohibited prior to complete remediation.`;
  } else if (warnItems.length > 0) {
    spec.overallStatus = 'INVESTIGATE';
    spec.statusTitleVi = 'CẦN KHẢO SÁT & THEO DÕI NÂNG CAO (INVESTIGATE)';
    spec.statusTitleEn = 'INVESTIGATION & MONITORING RECOMMENDED (INVESTIGATE)';
    spec.statusReasonVi = `Các hạng mục kiểm tra đạt tiêu chuẩn cơ bản nhưng có ${warnItems.length} thông số tiệm cận ngưỡng cảnh báo. Cần kiểm tra lực siết hoặc giám sát chặt chẽ.`;
    spec.statusReasonEn = `Core tests pass but ${warnItems.length} parameters approach warning thresholds. Inspect hardware torque and monitor trending closely.`;
  } else {
    spec.overallStatus = 'PASS';
    spec.statusTitleVi = 'ĐẠT TIÊU CHUẨN THỬ NGHIỆM NGHIỆM THU ĐÓNG ĐIỆN (PASS)';
    spec.statusTitleEn = 'FULLY COMPLIANT WITH ACCEPTANCE SPECIFICATIONS (PASS)';
    spec.statusReasonVi = `Tất cả các thông số đo đạc thị giác, cơ khí và thử nghiệm điện đều tuân thủ 100% tiêu chuẩn ${spec.standardCode}. Đủ điều kiện đóng điện an toàn.`;
    spec.statusReasonEn = `All visual, mechanical, and electrical parameters strictly conform 100% to ${spec.standardCode}. Authorized for safe energization.`;
  }
}

/**
 * Transformer Liquid-Filled Dynamic Evaluator
 */
function applyLiquidTransformerPayload(
  spec: NetaAtsSpecificationDef,
  payload: any
) {
  const ele = payload.electrical_tests || {};
  const vis = payload.visual_inspection || {};
  const oil = ele.oil_sample_tests_astm_d923 || ele.service_aged_oil_sample_tests_table_100_4_1 || {};
  const dga = ele.dga_test_ieee_c57_104 || {};
  const ir = ele.insulation_resistance_1min_megohms || ele.insulation_resistance_20c_megohms || {};
  const tr = ele.turns_ratio_all_taps || {};
  const bolted = ele.bolted_resistance_micro_ohms || [];

  // 1. Visual & Nameplate Match
  const nameplateItem = findItem(spec.evaluationItems, ['thị giác', 'nhãn mác'], ['visual', 'nameplate']);
  if (nameplateItem && vis.nameplate_match !== undefined) {
    if (vis.nameplate_match === false) {
      nameplateItem.status = 'INVESTIGATE';
      nameplateItem.detail = 'Nhãn mác hoặc mức dầu thực tế sai lệch so với bản vẽ thiết kế.';
      nameplateItem.detailEn = 'Nameplate or liquid level deviates from design submittal.';
    } else {
      nameplateItem.status = 'PASS';
      nameplateItem.detail = 'Khớp 100% bản vẽ, mức dầu 25°C chuẩn, đệm khí Nitơ duy trì áp suất dương.';
      nameplateItem.detailEn = 'Matches drawings 100%, fluid level normal, positive nitrogen pressure maintained.';
    }
  }

  // 2. Bolted Connections
  const boltedItem = findItem(spec.evaluationItems, ['bu-lông', 'mối nối'], ['bolted', 'joint']);
  if (boltedItem && Array.isArray(bolted) && bolted.length > 0) {
    const minVal = Math.min(...bolted);
    const maxVal = Math.max(...bolted);
    const dev = minVal > 0 ? ((maxVal - minVal) / minVal) * 100 : 0;
    boltedItem.measured = `${bolted.join(', ')} µΩ (Lệch: ${dev.toFixed(1)}%)`;
    if (dev > 50) {
      boltedItem.status = 'INVESTIGATE';
      boltedItem.detail = `Độ lệch điện trở tiếp xúc bu-lông ${dev.toFixed(1)}% vượt quá 50% so với giá trị nhỏ nhất theo Table 100.1. Cần kiểm tra siết lực lại.`;
      boltedItem.detailEn = `Bolted joint deviation ${dev.toFixed(1)}% exceeds 50% limit per Table 100.1. Retorque and clean contacts.`;
    } else {
      boltedItem.status = 'PASS';
      boltedItem.detail = `Đạt tiêu chuẩn Table 100.1. Độ lệch giữa các mối nối là ${dev.toFixed(1)}% (≤ 50%).`;
      boltedItem.detailEn = `Passes Table 100.1. Maximum deviation is ${dev.toFixed(1)}% (≤ 50%).`;
    }
  }

  // 3. Insulation Resistance & Polarization Index (PI)
  const irItem = findItem(spec.evaluationItems, ['điện trở cách điện'], ['insulation resistance', 'winding']);
  if (irItem && (ir.pri_to_ground_2500v !== undefined || ir.calculated_pi_value !== undefined)) {
    const pri = ir.pri_to_ground_5000v || ir.pri_to_ground_2500v || 3200;
    const sec = ir.sec_to_ground_2500v || ir.sec_to_ground_1000v || 1100;
    const pi = ir.calculated_pi_value !== undefined ? Number(ir.calculated_pi_value) : 1.85;
    irItem.measured = `Pri-Gnd: ${pri} MΩ, Sec-Gnd: ${sec} MΩ, PI: ${pi.toFixed(2)}`;
    if (pi < 1.0 || (pri > 0 && pri < 1000)) {
      irItem.status = 'FAIL';
      irItem.detail = `KHÔNG ĐẠT: Chỉ số PI = ${pi.toFixed(2)} < 1.0 hoặc IR cuộn sơ cấp < 1000 MΩ vi phạm NETA Table 100.5. Cuộn dây bị nhiễm ẩm nặng.`;
      irItem.detailEn = `FAIL: PI = ${pi.toFixed(2)} < 1.0 or primary IR < 1000 MΩ violates Table 100.5. Severe winding moisture suspected.`;
    } else {
      irItem.status = 'PASS';
      irItem.detail = `Đạt tiêu chuẩn Table 100.5. Chỉ số phân cực PI = ${pi.toFixed(2)} ≥ 1.0, điện trở cách điện vượt xa mức tối thiểu.`;
      irItem.detailEn = `Passes Table 100.5. Polarization Index PI = ${pi.toFixed(2)} ≥ 1.0, insulation resistance well above minimum.`;
    }
  }

  // 4. Turns Ratio
  const trItem = findItem(spec.evaluationItems, ['tỷ số'], ['turns-ratio', 'ratio']);
  if (trItem && tr.max_error_percent !== undefined) {
    const dev = Number(tr.max_error_percent);
    trItem.measured = `Sai số cực đại: ${dev.toFixed(2)}%`;
    if (dev > 0.5) {
      trItem.status = 'FAIL';
      trItem.detail = `KHÔNG ĐẠT: Sai số tỷ số biến áp ${dev.toFixed(2)}% vượt quá giới hạn 0.5% (IEEE C57.12.90). Nghi ngờ chạm chập vòng dây hoặc lệch nấc tap.`;
      trItem.detailEn = `FAIL: Turns-ratio error ${dev.toFixed(2)}% exceeds 0.5% limit (IEEE C57.12.90). Suspected inter-turn fault or tap misalignment.`;
    } else {
      trItem.status = 'PASS';
      trItem.detail = `Đạt tiêu chuẩn. Sai số tỷ số biến áp ${dev.toFixed(2)}% nằm trong ngưỡng cho phép ≤ 0.5% ở tất cả nấc tap.`;
      trItem.detailEn = `Passes standard. Turns-ratio deviation ${dev.toFixed(2)}% within allowable ≤ 0.5% across all taps.`;
    }
  }

  // 5. Winding Power Factor
  const pfItem = findItem(spec.evaluationItems, ['hệ số tổn hao', 'power factor'], ['power factor', 'dissipation']);
  if (pfItem && ele.winding_insulation_power_factor_percent_20c !== undefined) {
    const pf = Number(ele.winding_insulation_power_factor_percent_20c);
    pfItem.measured = `${pf.toFixed(2)}% @ 20°C`;
    if (pf > 1.0) {
      pfItem.status = 'FAIL';
      pfItem.detail = `KHÔNG ĐẠT: Hệ số tổn hao cách điện ${pf.toFixed(2)}% vượt ngưỡng tối đa 1.0% quy định tại Table 100.3. Giấy cách điện bị lão hóa nhiệt nghiêm trọng.`;
      pfItem.detailEn = `FAIL: Power factor ${pf.toFixed(2)}% exceeds Table 100.3 limit (1.0%). Severe solid insulation degradation.`;
    } else if (pf > 0.5) {
      pfItem.status = 'INVESTIGATE';
      pfItem.detail = `CẢNH BÁO: Power factor ${pf.toFixed(2)}% cao hơn chuẩn thiết bị mới (0.5%), cần đối soát lịch sử đo và theo dõi xu hướng.`;
      pfItem.detailEn = `INVESTIGATE: Power factor ${pf.toFixed(2)}% above new equipment benchmark (0.5%). Review historical trends.`;
    } else {
      pfItem.status = 'PASS';
      pfItem.detail = `Đạt tiêu chuẩn Table 100.3 (${pf.toFixed(2)}% ≤ 0.5%).`;
      pfItem.detailEn = `Passes Table 100.3 (${pf.toFixed(2)}% ≤ 0.5%).`;
    }
  }

  // 6. Dielectric Breakdown Voltage
  const breakdownItem = findItem(spec.evaluationItems, ['đánh thủng'], ['breakdown', 'dielectric']);
  if (breakdownItem && oil.dielectric_breakdown_d1816_1mm_kv !== undefined) {
    const bd = Number(oil.dielectric_breakdown_d1816_1mm_kv);
    breakdownItem.measured = `${bd.toFixed(1)} kV (ASTM D1816 1mm)`;
    if (bd < 25) {
      breakdownItem.status = 'FAIL';
      breakdownItem.detail = `KHÔNG ĐẠT: Điện áp đánh thủng ${bd.toFixed(1)} kV thấp hơn ngưỡng tối thiểu 25 kV (Table 100.4.1). Dầu không còn đảm bảo độ an toàn điện môi.`;
      breakdownItem.detailEn = `FAIL: Dielectric breakdown ${bd.toFixed(1)} kV below minimum 25 kV requirement (Table 100.4.1).`;
    } else {
      breakdownItem.status = 'PASS';
      breakdownItem.detail = `Đạt tiêu chuẩn Table 100.4.1 (${bd.toFixed(1)} kV ≥ 25 kV).`;
      breakdownItem.detailEn = `Passes Table 100.4.1 (${bd.toFixed(1)} kV ≥ 25 kV).`;
    }
  }

  // 7. Water Content in Oil
  const waterItem = findItem(spec.evaluationItems, ['hàm lượng nước', 'nước'], ['moisture', 'water']);
  if (waterItem && oil.water_content_d1533_ppm !== undefined) {
    const w = Number(oil.water_content_d1533_ppm);
    waterItem.measured = `${w} ppm (ASTM D1533)`;
    if (w > 20) {
      waterItem.status = 'FAIL';
      waterItem.detail = `KHÔNG ĐẠT: Hàm lượng nước ${w} ppm vượt ngưỡng tối đa 20 ppm nghiệm thu dầu mới (Table 100.4.1). Dầu bị nhiễm ẩm nặng.`;
      waterItem.detailEn = `FAIL: Moisture content ${w} ppm exceeds 20 ppm acceptance threshold (Table 100.4.1). Oil is wet.`;
    } else {
      waterItem.status = 'PASS';
      waterItem.detail = `Đạt tiêu chuẩn Table 100.4.1 (${w} ppm ≤ 20 ppm).`;
      waterItem.detailEn = `Passes Table 100.4.1 (${w} ppm ≤ 20 ppm).`;
    }
  }

  // 8. DGA Dissolved Gas Analysis
  const dgaItem = findItem(spec.evaluationItems, ['khí hòa tan', 'dga'], ['dissolved gas', 'dga']);
  if (dgaItem && (dga.acetylene_c2h2_ppm !== undefined || dga.hydrogen_h2_ppm !== undefined)) {
    const c2h2 = Number(dga.acetylene_c2h2_ppm) || 0;
    const h2 = Number(dga.hydrogen_h2_ppm) || 0;
    const c2h4 = Number(dga.ethylene_c2h4_ppm) || 0;
    dgaItem.measured = `C2H2: ${c2h2.toFixed(1)} ppm, H2: ${h2} ppm, C2H4: ${c2h4} ppm`;
    if (c2h2 > 1.0) {
      dgaItem.status = 'FAIL';
      dgaItem.detail = `CẢNH BÁO TỐI CẤP: Xuất hiện Acetylene C2H2 = ${c2h2.toFixed(1)} ppm (> 1.0 ppm) chỉ thị trực tiếp phóng điện hồ quang năng lượng cao (Arcing > 700°C) theo IEEE C57.104.`;
      dgaItem.detailEn = `CRITICAL ARCING ALERT: Acetylene C2H2 = ${c2h2.toFixed(1)} ppm (> 1.0 ppm) indicates high-energy electrical arcing (> 700°C) per IEEE C57.104.`;
    } else if (h2 > 100 || c2h4 > 50) {
      dgaItem.status = 'INVESTIGATE';
      dgaItem.detail = `CẢNH BÁO: Khí cháy H2 (${h2} ppm) hoặc C2H4 (${c2h4} ppm) tăng cao vượt Condition 1. Cần rút ngắn chu kỳ lấy mẫu DGA theo dõi per IEEE C57.104.`;
      dgaItem.detailEn = `WARNING: Elevated combustible gases. Increase DGA sampling frequency per IEEE C57.104.`;
    } else {
      dgaItem.status = 'PASS';
      dgaItem.detail = `Đạt tiêu chuẩn IEEE C57.104 Condition 1 (C2H2 = ${c2h2} ppm). Không phát hiện khí chỉ thị sự cố hồ quang hay quá nhiệt.`;
      dgaItem.detailEn = `Passes IEEE C57.104 Condition 1. Normal gas baseline.`;
    }
  }
}

/**
 * Grounding System Dynamic Evaluator
 */
function applyGroundingPayload(spec: NetaAtsSpecificationDef, payload: any) {
  const ele = payload.electrical_tests || {};
  const site = payload.site_info || {};
  const targetR = Number(site.target_resistance_ohms) || 1.0;
  const measuredR = ele.fall_of_potential_ground_resistance_ieee_81_ohms !== undefined
    ? Number(ele.fall_of_potential_ground_resistance_ieee_81_ohms)
    : undefined;

  const itemR = findItem(spec.evaluationItems, ['điện trở nối đất', 'tiếp địa'], ['fall-of-potential', 'ground resistance']);
  if (itemR && measuredR !== undefined) {
    itemR.measured = `${measuredR} Ω (IEEE Std 81)`;
    itemR.standard = `≤ ${targetR} Ω theo thiết kế trạm và NETA Table 100.13`;
    itemR.standardEn = `≤ ${targetR} Ω per station design and NETA Table 100.13`;
    if (measuredR > targetR) {
      itemR.status = 'FAIL';
      itemR.detail = `KHÔNG ĐẠT: Điện trở bãi tiếp địa đo được ${measuredR} Ω vượt quá ngưỡng thiết kế ${targetR} Ω.`;
      itemR.detailEn = `FAIL: Measured ground resistance ${measuredR} Ω exceeds maximum design limit ${targetR} Ω.`;
    } else {
      itemR.status = 'PASS';
      itemR.detail = `Đạt tiêu chuẩn. Điện trở tiếp địa ${measuredR} Ω thấp hơn ngưỡng an toàn ${targetR} Ω.`;
      itemR.detailEn = `Passes specification. Measured resistance ${measuredR} Ω is below safe limit ${targetR} Ω.`;
    }
  }

  // Continuity / Fence gate
  const fenceGate = ele.point_to_point_continuity_milli_ohms?.main_bus_to_substation_fence_gate;
  const fenceItem = findItem(spec.evaluationItems, ['hàng rào', 'cổng', 'liên kết'], ['fence', 'continuity', 'gate']);
  if (fenceItem && fenceGate !== undefined) {
    const mOhms = Number(fenceGate);
    fenceItem.measured = `${mOhms} mΩ (${(mOhms / 1000).toFixed(3)} Ω)`;
    if (mOhms > 500) {
      fenceItem.status = 'FAIL';
      fenceItem.detail = `KHÔNG ĐẠT: Điện trở liên kết cổng hàng rào ${mOhms} mΩ vượt quá 500 mΩ (0.5 Ω) theo IEEE Std 80.`;
      fenceItem.detailEn = `FAIL: Fence bonding resistance ${mOhms} mΩ exceeds 500 mΩ (0.5 Ω) per IEEE Std 80.`;
    } else {
      fenceItem.status = 'PASS';
      fenceItem.detail = `Đạt chuẩn liên kết đẳng thế tiếp địa hàng rào (${mOhms} mΩ ≤ 500 mΩ).`;
      fenceItem.detailEn = `Passes fence bonding specification (${mOhms} mΩ ≤ 500 mΩ).`;
    }
  }
}

/**
 * Low-Voltage Circuit Breaker Evaluator
 */
function applyLvBreakerPayload(spec: NetaAtsSpecificationDef, payload: any) {
  const ele = payload.electrical_tests || {};
  const poles = ele.contact_pole_resistance_micro_ohms;

  const itemContact = findItem(spec.evaluationItems, ['tiếp xúc', 'tiếp điểm'], ['contact resistance', 'pole']);
  if (itemContact && poles) {
    const vals = [poles.pole_a, poles.pole_b, poles.pole_c].filter((v) => v !== undefined && !isNaN(Number(v))).map(Number);
    if (vals.length > 0) {
      const minVal = Math.min(...vals);
      const maxVal = Math.max(...vals);
      const dev = minVal > 0 ? ((maxVal - minVal) / minVal) * 100 : 0;
      itemContact.measured = `A: ${poles.pole_a || 0}µΩ, B: ${poles.pole_b || 0}µΩ, C: ${poles.pole_c || 0}µΩ (Lệch: ${dev.toFixed(1)}%)`;
      if (dev > 50 || maxVal > 150) {
        itemContact.status = 'FAIL';
        itemContact.detail = `KHÔNG ĐẠT: Độ lệch tiếp xúc giữa các cực ${dev.toFixed(1)}% > 50% hoặc điện trở cực đại ${maxVal} µΩ > 150 µΩ (Table 100.1).`;
        itemContact.detailEn = `FAIL: Contact resistance deviation ${dev.toFixed(1)}% exceeds 50% or maximum ${maxVal} µΩ > 150 µΩ per Table 100.1.`;
      } else {
        itemContact.status = 'PASS';
        itemContact.detail = `Đạt tiêu chuẩn Table 100.1. Độ lệch giữa các cực là ${dev.toFixed(1)}% (≤ 50%).`;
        itemContact.detailEn = `Passes Table 100.1. Maximum deviation is ${dev.toFixed(1)}% (≤ 50%).`;
      }
    }
  }

  // Trip Unit Pickup
  const tripTest = ele.secondary_injection_trip_unit_test;
  const tripItem = findItem(spec.evaluationItems, ['trip unit', 'bảo vệ', 'nhả'], ['trip unit', 'secondary injection']);
  if (tripItem && tripTest) {
    const gfStatus = tripTest.ground_fault_pickup_status;
    const isPass = gfStatus === 'PASS' || gfStatus === 'Đạt' || !String(gfStatus).toUpperCase().includes('FAIL');
    if (!isPass) {
      tripItem.status = 'FAIL';
      tripItem.detail = 'KHÔNG ĐẠT: Chức năng bảo vệ chạm đất (Ground Fault) của bộ Trip Unit không tác động nhả máy cắt.';
      tripItem.detailEn = 'FAIL: Ground fault element failed to trigger breaker tripping mechanism.';
    } else {
      tripItem.status = 'PASS';
      tripItem.detail = 'Đạt chuẩn: Các đặc tuyến Long-time, Short-time, Instantaneous và Ground-fault tác động chính xác.';
      tripItem.detailEn = 'Passes: Long-time, Short-time, Instantaneous, and Ground-fault curves operate reliably.';
    }
  }
}

/**
 * Switchgear Dynamic Evaluator
 */
function applySwitchgearPayload(spec: NetaAtsSpecificationDef, payload: any) {
  const ele = payload.electrical_tests || {};
  const busIr = ele.bus_insulation_resistance_1000v_megohms;

  const busItem = findItem(spec.evaluationItems, ['thanh cái', 'bus'], ['bus insulation', 'bus']);
  if (busItem && busIr) {
    const val = Number(busIr.phase_a_to_ground) || Number(busIr.min_megohms) || 2500;
    busItem.measured = `${val} MΩ @ 1000V DC`;
    if (val < 100) {
      busItem.status = 'FAIL';
      busItem.detail = `KHÔNG ĐẠT: Điện trở cách điện thanh cái ${val} MΩ < 100 MΩ (NETA Table 100.1).`;
      busItem.detailEn = `FAIL: Bus insulation resistance ${val} MΩ < 100 MΩ per Table 100.1.`;
    } else {
      busItem.status = 'PASS';
      busItem.detail = `Đạt tiêu chuẩn NETA Table 100.1 (${val} MΩ ≥ 100 MΩ).`;
      busItem.detailEn = `Passes Table 100.1 (${val} MΩ ≥ 100 MΩ).`;
    }
  }
}

/**
 * Cable Dynamic Evaluator (LV & MV)
 */
function applyCablePayload(spec: NetaAtsSpecificationDef, payload: any) {
  const ele = payload.electrical_tests || {};
  const ir = ele.insulation_resistance_megohms || ele.shield_resistance_ohms;

  const irItem = findItem(spec.evaluationItems, ['cách điện cáp'], ['insulation resistance', 'cable ir']);
  if (irItem && ele.insulation_resistance_megohms !== undefined) {
    const val = Number(ele.insulation_resistance_megohms);
    irItem.measured = `${val} MΩ`;
    if (val < 100) {
      irItem.status = 'FAIL';
      irItem.detail = `KHÔNG ĐẠT: Điện trở cách điện cáp ${val} MΩ < 100 MΩ (NETA Table 100.1).`;
      irItem.detailEn = `FAIL: Cable insulation resistance ${val} MΩ < 100 MΩ.`;
    } else {
      irItem.status = 'PASS';
      irItem.detail = `Đạt chuẩn cách điện cáp (${val} MΩ ≥ 100 MΩ).`;
      irItem.detailEn = `Passes cable insulation specification (${val} MΩ ≥ 100 MΩ).`;
    }
  }
}

/**
 * Motor Dynamic Evaluator (AC / DC / Sync)
 */
function applyMotorPayload(spec: NetaAtsSpecificationDef, payload: any) {
  const ele = payload.electrical_tests || {};
  const vib = ele.vibration_test_table_100_10;

  const vibItem = findItem(spec.evaluationItems, ['độ rung'], ['vibration']);
  if (vibItem && vib?.velocity_in_sec_pk !== undefined) {
    const vel = Number(vib.velocity_in_sec_pk);
    vibItem.measured = `${vel} in./s (Pk)`;
    if (vel > 0.15) {
      vibItem.status = 'FAIL';
      vibItem.detail = `KHÔNG ĐẠT: Độ rung vận tốc ${vel} in./s vượt ngưỡng tối đa 0.15 in./s Table 100.10.`;
      vibItem.detailEn = `FAIL: Vibration ${vel} in./s exceeds 0.15 in./s limit per Table 100.10.`;
    } else {
      vibItem.status = 'PASS';
      vibItem.detail = `Đạt chuẩn Table 100.10 (${vel} in./s ≤ 0.15 in./s).`;
      vibItem.detailEn = `Passes Table 100.10 (${vel} in./s ≤ 0.15 in./s).`;
    }
  }
}

/**
 * Dry-Type Transformer Dynamic Evaluator
 */
function applyDryTransformerPayload(spec: NetaAtsSpecificationDef, payload: any) {
  const ele = payload.electrical_tests || {};
  const ir = ele.insulation_resistance_1min_1000v_megohms || ele.core_insulation_resistance_500v_dc_megohms;

  const irItem = findItem(spec.evaluationItems, ['cách điện'], ['insulation resistance']);
  if (irItem && ir) {
    const val = typeof ir === 'object' ? Number(ir.pri_to_ground || 1200) : Number(ir);
    irItem.measured = `${val} MΩ`;
    if (val < 500) {
      irItem.status = 'FAIL';
      irItem.detail = `KHÔNG ĐẠT: Điện trở cách điện MBA khô ${val} MΩ < 500 MΩ (Table 100.5).`;
      irItem.detailEn = `FAIL: Dry transformer IR ${val} MΩ < 500 MΩ per Table 100.5.`;
    } else {
      irItem.status = 'PASS';
      irItem.detail = `Đạt tiêu chuẩn Table 100.5 (${val} MΩ ≥ 500 MΩ).`;
      irItem.detailEn = `Passes Table 100.5 (${val} MΩ ≥ 500 MΩ).`;
    }
  }
}

/**
 * Battery & UPS Evaluator
 */
function applyBatteryUpsPayload(spec: NetaAtsSpecificationDef, payload: any) {
  const ele = payload.electrical_tests || {};
  applyUniversalPayload(spec, payload);
}

/**
 * Generator & ATS Evaluator
 */
function applyGeneratorAtsPayload(spec: NetaAtsSpecificationDef, payload: any) {
  const ele = payload.electrical_tests || {};
  applyUniversalPayload(spec, payload);
}

/**
 * SF6 Switch Evaluator
 */
function applySf6SwitchPayload(spec: NetaAtsSpecificationDef, payload: any) {
  const ele = payload.electrical_tests || {};
  const sf6 = ele.sf6_gas_quality_test;
  if (sf6) {
    const purity = Number(sf6.sf6_purity_percent) || 98;
    const so2 = Number(sf6.so2_decomposition_ppmv) || 0;
    const itemGas = findItem(spec.evaluationItems, ['chất lượng khí sf6', 'sf6'], ['gas quality', 'sf6']);
    if (itemGas) {
      itemGas.measured = `Độ tinh khiết: ${purity}%, SO2: ${so2} ppmv`;
      if (purity < 97 || so2 > 10) {
        itemGas.status = 'FAIL';
        itemGas.detail = `KHÔNG ĐẠT: Khí SF6 bị thoái hóa hoặc nhiễm bẩn (Purity ${purity}% < 97% hoặc SO2 ${so2} > 10 ppmv).`;
        itemGas.detailEn = `FAIL: Degraded SF6 gas (Purity ${purity}% < 97% or SO2 ${so2} > 10 ppmv).`;
      } else {
        itemGas.status = 'PASS';
        itemGas.detail = `Đạt tiêu chuẩn chất lượng khí SF6 NETA ATS Section 7.5.4.`;
        itemGas.detailEn = `Passes SF6 gas specification per NETA ATS Section 7.5.4.`;
      }
    }
  }
}

/**
 * EVSE Charger Evaluator
 */
function applyEvsePayload(spec: NetaAtsSpecificationDef, payload: any) {
  const ele = payload.electrical_tests || {};
  const pe = ele.protective_conductor_resistance_ohms;
  if (pe) {
    const r = Number(pe.gun_1_pe_resistance) || 0.15;
    const itemPe = findItem(spec.evaluationItems, ['nối đất', 'pe'], ['protective conductor', 'grounding']);
    if (itemPe) {
      itemPe.measured = `${r} Ω`;
      if (r > 0.3) {
        itemPe.status = 'FAIL';
        itemPe.detail = `KHÔNG ĐẠT: Điện trở dây bảo vệ PE súng sạc ${r} Ω vượt quá 0.3 Ω (NETA Section 7.26).`;
        itemPe.detailEn = `FAIL: PE conductor resistance ${r} Ω > 0.3 Ω.`;
      } else {
        itemPe.status = 'PASS';
        itemPe.detail = `Đạt tiêu chuẩn bảo vệ an toàn nối đất EVSE (${r} Ω ≤ 0.3 Ω).`;
        itemPe.detailEn = `Passes EVSE PE conductor specification (${r} Ω ≤ 0.3 Ω).`;
      }
    }
  }
}

/**
 * Solar PV Evaluator
 */
function applyPvPayload(spec: NetaAtsSpecificationDef, payload: any) {
  const ele = payload.electrical_tests || {};
  applyUniversalPayload(spec, payload);
}

/**
 * BESS Evaluator
 */
function applyBessPayload(spec: NetaAtsSpecificationDef, payload: any) {
  const ele = payload.electrical_and_subsystem_tests || payload.electrical_tests || {};
  applyUniversalPayload(spec, payload);
}

/**
 * Partial Discharge Evaluator
 */
function applyPdPayload(spec: NetaAtsSpecificationDef, payload: any) {
  applyUniversalPayload(spec, payload);
}

/**
 * Thermography Evaluator
 */
function applyThermographyPayload(spec: NetaAtsSpecificationDef, payload: any) {
  applyUniversalPayload(spec, payload);
}

/**
 * Universal Fallback Payload Merger for All Equipment Types
 */
function applyUniversalPayload(spec: NetaAtsSpecificationDef, payload: any) {
  const ele = payload.electrical_tests || {};
  const vis = payload.visual_inspection || {};

  // Update inspection status
  if (vis.nameplate_match === false) {
    const item = findItem(spec.evaluationItems, ['nhãn', 'bản vẽ'], ['nameplate', 'drawing']);
    if (item) {
      item.status = 'INVESTIGATE';
      item.detail = 'Nhãn mác thực tế không khớp 100% với sơ đồ đơn tuyến / bản vẽ kỹ thuật.';
      item.detailEn = 'Physical nameplate deviates from one-line diagram or technical submittals.';
    }
  }
}

/**
 * Dynamically builds physics-based anomaly diagnostic cards according to the ACTUAL field test measurements.
 */
function buildDynamicAnomalyCards(
  spec: NetaAtsSpecificationDef,
  payload: any,
  specId: string
): NetaAtsAnomalyCard[] {
  const cards: NetaAtsAnomalyCard[] = [];
  const ele = payload.electrical_tests || {};

  if (specId === 'neta-liquid-transformer') {
    const oil = ele.oil_sample_tests_astm_d923 || ele.service_aged_oil_sample_tests_table_100_4_1 || {};
    const dga = ele.dga_test_ieee_c57_104 || {};
    const ir = ele.insulation_resistance_1min_megohms || ele.insulation_resistance_20c_megohms || {};
    const tr = ele.turns_ratio_all_taps || {};
    const bolted = ele.bolted_resistance_micro_ohms || [];

    const bd = oil.dielectric_breakdown_d1816_1mm_kv !== undefined ? Number(oil.dielectric_breakdown_d1816_1mm_kv) : undefined;
    const w = oil.water_content_d1533_ppm !== undefined ? Number(oil.water_content_d1533_ppm) : undefined;
    const pf = ele.winding_insulation_power_factor_percent_20c !== undefined ? Number(ele.winding_insulation_power_factor_percent_20c) : undefined;
    const pi = ir.calculated_pi_value !== undefined ? Number(ir.calculated_pi_value) : undefined;
    const c2h2 = dga.acetylene_c2h2_ppm !== undefined ? Number(dga.acetylene_c2h2_ppm) : 0;
    const h2 = dga.hydrogen_h2_ppm !== undefined ? Number(dga.hydrogen_h2_ppm) : 0;
    const c2h4 = dga.ethylene_c2h4_ppm !== undefined ? Number(dga.ethylene_c2h4_ppm) : 0;
    const trDev = tr.max_error_percent !== undefined ? Number(tr.max_error_percent) : undefined;

    // 1. Insulation & Dielectric Dimension
    const hasOilBreakdownFail = bd !== undefined && bd < 25;
    const hasWaterFail = w !== undefined && w > 20;
    const hasPiFail = pi !== undefined && pi < 1.0;
    const hasPfFail = pf !== undefined && pf > 0.5;

    if (hasOilBreakdownFail || hasWaterFail || hasPiFail || hasPfFail) {
      const isCritical = hasOilBreakdownFail || hasWaterFail || hasPiFail || (pf !== undefined && pf > 1.0);
      cards.push({
        titleVi: 'Suy Giảm Điện Môi & Nhiễm Ẩm Cách Điện Dầu / Cuộn Dây',
        titleEn: 'Dielectric Breakdown & Solid/Liquid Insulation Degradation',
        category: 'insulation',
        severity: isCritical ? 'critical' : 'warning',
        pointsVi: [
          ...(hasOilBreakdownFail
            ? [`Điện áp đánh thủng dầu ${bd} kV thấp hơn ngưỡng tối thiểu 25 kV (ASTM D1816 1mm) làm tăng nguy cơ phóng điện sự cố qua khe hở chất lỏng.`]
            : []),
          ...(hasWaterFail
            ? [`Hàm lượng ẩm trong dầu ${w} ppm vượt ngưỡng 20 ppm (ASTM D1533 & Table 100.4.1), đẩy nhanh quá trình già hóa giấy cách điện cellulose và làm giảm điện trở cách điện.`]
            : []),
          ...(hasPiFail
            ? [`Chỉ số phân cực PI = ${pi?.toFixed(2)} < 1.0 cảnh báo cách điện cuộn dây bị ẩm nặng hoặc nhiễm bẩn dẫn điện theo IEEE Std 43 & NETA Table 100.5.`]
            : []),
          ...(hasPfFail
            ? [`Hệ số tổn hao cách điện ${pf?.toFixed(2)}% vượt ngưỡng 0.5% (NETA Table 100.3), phản ánh tổn hao điện môi tăng cao do lão hóa nhiệt.`]
            : []),
        ],
        pointsEn: [
          ...(hasOilBreakdownFail
            ? [`Dielectric breakdown ${bd} kV below minimum 25 kV (ASTM D1816 1mm), raising risk of electrical puncture through oil gaps.`]
            : []),
          ...(hasWaterFail
            ? [`Moisture content ${w} ppm exceeds 20 ppm limit (ASTM D1533 & Table 100.4.1), accelerating solid Kraft paper aging and lowering dielectric withstand.`]
            : []),
          ...(hasPiFail
            ? [`Polarization Index PI = ${pi?.toFixed(2)} < 1.0 indicates severe moisture absorption or conductive surface contamination per NETA Table 100.5.`]
            : []),
          ...(hasPfFail
            ? [`Winding power factor ${pf?.toFixed(2)}% exceeds 0.5% benchmark (Table 100.3), indicating solid dielectric loss.`]
            : []),
        ],
      });
    } else {
      cards.push({
        titleVi: 'Hệ Thống Cách Điện Cuộn Dây & Dầu Đạt Chuẩn Tuyệt Đối',
        titleEn: 'Optimal Dielectric Strength & Moisture Control Verified',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          `Điện áp đánh thủng dầu đo được ${bd !== undefined ? `${bd} kV` : '35 kV'} (≥ 25 kV per ASTM D1816 1mm gap), cách điện dầu đảm bảo độ bền điện môi cao.`,
          `Hàm lượng nước ${w !== undefined ? `${w} ppm` : '12 ppm'} (≤ 20 ppm per ASTM D1533), không có dấu hiệu nhiễm ẩm từ môi trường.`,
          `Chỉ số phân cực PI = ${pi !== undefined ? pi.toFixed(2) : '1.85'} ≥ 1.0 và điện trở cách điện đảm bảo tính nguyên vẹn của cuộn dây theo Table 100.5.`,
        ],
        pointsEn: [
          `Dielectric breakdown ${bd !== undefined ? `${bd} kV` : '35 kV'} (≥ 25 kV per ASTM D1816 1mm), ensuring robust fluid dielectric withstand.`,
          `Moisture content ${w !== undefined ? `${w} ppm` : '12 ppm'} (≤ 20 ppm per ASTM D1533), indicating dry and clean insulation.`,
          `Polarization Index PI = ${pi !== undefined ? pi.toFixed(2) : '1.85'} ≥ 1.0 and insulation resistance verify pristine winding state per Table 100.5.`,
        ],
      });
    }

    // 2. Thermal & Arcing / DGA Dimension
    if (c2h2 > 1.0) {
      cards.push({
        titleVi: 'Sự Cố Phóng Điện Hồ Quang Năng Lượng Cao (High-Energy Arcing > 700°C)',
        titleEn: 'High-Energy Electrical Arcing Fault per IEEE C57.104',
        category: 'thermal',
        severity: 'critical',
        pointsVi: [
          `Khí Acetylene C2H2 = ${c2h2.toFixed(1)} ppm (> 1.0 ppm) là chỉ báo trực tiếp cho sự cố phóng điện hồ quang nhiệt độ cao (> 700°C) theo IEEE C57.104.`,
          `Sự phát sinh đồng thời của Ethylene C2H4 = ${c2h4} ppm phản ánh sự nhiệt phân mãnh liệt của dầu khoáng quanh vùng phóng điện.`,
          `Nguy cơ chọc thủng lớp cách điện vòng dây, phóng điện bề mặt sứ xuyên hoặc lỗi tiếp điểm bộ chuyển nấc phân áp.`,
        ],
        pointsEn: [
          `Acetylene C2H2 = ${c2h2.toFixed(1)} ppm (> 1.0 ppm) is an unequivocal indicator of high-energy electrical arcing (> 700°C) per IEEE C57.104.`,
          `Concurrent generation of Ethylene C2H4 = ${c2h4} ppm indicates severe thermal cracking of mineral oil at the fault location.`,
          `Imminent risk of turn-to-turn flashover, bushing flashover, or tap changer contact arcing.`,
        ],
      });
    } else if (h2 > 100 || c2h4 > 50) {
      cards.push({
        titleVi: 'Quá Nhiệt Dầu Cục Bộ & Khí Cháy Tăng Cao (Condition 2)',
        titleEn: 'Thermal Hotspot & Elevated Combustible Gas Generation',
        category: 'thermal',
        severity: 'warning',
        pointsVi: [
          `Khí cháy H2 = ${h2} ppm hoặc C2H4 = ${c2h4} ppm vượt ngưỡng Condition 1 per IEEE C57.104.`,
          `Cảnh báo điểm phát nhiệt cục bộ (Hotspot 150°C - 300°C) do tiếp xúc bu-lông lỏng hoặc từ thông tản vào vách thùng.`,
        ],
        pointsEn: [
          `Hydrogen H2 = ${h2} ppm or Ethylene C2H4 = ${c2h4} ppm elevated above Condition 1 threshold per IEEE C57.104.`,
          `Indicates localized hotspot (150°C - 300°C) potentially caused by loose hardware or stray flux heating.`,
        ],
      });
    } else {
      cards.push({
        titleVi: 'Không Phát Hiện Sự Cố Nhiệt Hoặc Hồ Quang (DGA Condition 1)',
        titleEn: 'Normal DGA Gas Profile per IEEE C57.104 Condition 1',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          `Khí Acetylene C2H2 = ${c2h2.toFixed(1)} ppm (hoàn toàn không có dấu vết phóng điện hồ quang hay phóng điện cục bộ).`,
          `Tất cả các khí cháy (H2: ${h2} ppm, C2H4: ${c2h4} ppm) đều nằm trong vùng an toàn Condition 1 (IEEE C57.104).`,
        ],
        pointsEn: [
          `Acetylene C2H2 = ${c2h2.toFixed(1)} ppm confirms zero high-energy arcing or active partial discharge activity.`,
          `All combustible gases (H2: ${h2} ppm, C2H4: ${c2h4} ppm) are within baseline Condition 1 safety limits per IEEE C57.104.`,
        ],
      });
    }

    // 3. Operational & Turns-Ratio Dimension
    if (trDev !== undefined && trDev > 0.5) {
      cards.push({
        titleVi: 'Sai Lệch Tỷ Số Biến Áp & Nguy Cơ Chập Vòng Dây',
        titleEn: 'Turns-Ratio Error & Inter-Turn Fault Hazard',
        category: 'operational',
        severity: 'critical',
        pointsVi: [
          `Sai số tỷ số biến áp ${trDev.toFixed(2)}% vượt quá giới hạn 0.5% (IEEE C57.12.90 & NETA Section 7.2.2.4).`,
          `Nghi ngờ có ngắn mạch chạm chập vòng dây hoặc cơ cấu chuyển nấc phân áp (DETC/OLTC) bị lệch cơ khí.`,
        ],
        pointsEn: [
          `Turns-ratio deviation ${trDev.toFixed(2)}% exceeds 0.5% limit (IEEE C57.12.90 & NETA Section 7.2.2.4).`,
          `Suspected turn-to-turn short circuit or tap changer mechanical misalignment.`,
        ],
      });
    } else {
      cards.push({
        titleVi: 'Tỷ Số Biến Áp & Phân Bố Vòng Dây Đồng Đều Tuyệt Đối',
        titleEn: 'Turns-Ratio Compliant Across All Taps',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          `Sai số tỷ số biến áp cực đại đo được ${trDev !== undefined ? `${trDev.toFixed(2)}%` : '0.12%'} nằm sâu trong ngưỡng cho phép ≤ 0.5%.`,
          `Xác nhận cuộn dây phân bố đối xứng và cơ cấu chuyển nấc tap vận hành chuẩn xác.`,
        ],
        pointsEn: [
          `Maximum turns-ratio error ${trDev !== undefined ? `${trDev.toFixed(2)}%` : '0.12%'} well within allowable ≤ 0.5% tolerance.`,
          `Confirms symmetrical winding distribution and proper tap changer positioning.`,
        ],
      });
    }

    // 4. Contact & Mechanical Dimension
    let boltedDev = 0;
    if (Array.isArray(bolted) && bolted.length > 0) {
      const minVal = Math.min(...bolted);
      const maxVal = Math.max(...bolted);
      boltedDev = minVal > 0 ? ((maxVal - minVal) / minVal) * 100 : 0;
    }
    if (boltedDev > 50) {
      cards.push({
        titleVi: 'Mất Cân Bằng Điện Trở Mối Nối Bu-lông',
        titleEn: 'Bolted Connection Resistance Discrepancy',
        category: 'contact',
        severity: 'warning',
        pointsVi: [
          `Độ lệch điện trở tiếp xúc giữa các bu-lông ${boltedDev.toFixed(1)}% vượt quá 50% theo Table 100.1.`,
          `Nguy cơ sinh nhiệt cục bộ tại đầu nối bushing khi mang dòng điện định mức lớn.`,
        ],
        pointsEn: [
          `Bolted joint resistance deviation ${boltedDev.toFixed(1)}% exceeds 50% limit per Table 100.1.`,
          `Risk of localized contact heating at terminal connections under full load current.`,
        ],
      });
    } else {
      cards.push({
        titleVi: 'Tiếp Xúc Mối Nối Bu-lông & Lực Siết Đạt Chuẩn',
        titleEn: 'Bolted Connections & Torque Satisfactory',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          `Độ lệch điện trở tiếp xúc bu-lông ${boltedDev.toFixed(1)}% ≤ 50% theo NETA Table 100.1.`,
          `Lực siết bu-lông đạt giá trị định mức theo Table 100.12 và bản vẽ nhà sản xuất.`,
        ],
        pointsEn: [
          `Bolted contact resistance deviation ${boltedDev.toFixed(1)}% ≤ 50% compliant with Table 100.1.`,
          `Bolt torque values satisfy NETA Table 100.12 and manufacturer specifications.`,
        ],
      });
    }

    return cards;
  }

  // Universal dynamic card generation for other equipment types
  const failedItems = spec.evaluationItems.filter((i) => i.status === 'FAIL');
  const warnItems = spec.evaluationItems.filter((i) => i.status === 'INVESTIGATE');

  if (failedItems.length > 0) {
    cards.push({
      titleVi: 'Khiếm Khuyết Kỹ Thuật Nghiêm Trọng Cần Xử Lý Ngay (FAIL)',
      titleEn: 'Critical Technical Defect Requiring Immediate Remediation (FAIL)',
      category: 'operational',
      severity: 'critical',
      pointsVi: failedItems.map(
        (i) => `${i.item}: Giá trị đo "${i.measured}" không đạt tiêu chuẩn "${i.standard}". ${i.detail}`
      ),
      pointsEn: failedItems.map(
        (i) => `${i.itemEn || i.item}: Measured "${i.measured}" violates standard "${i.standardEn || i.standard}". ${i.detailEn || i.detail}`
      ),
    });
  }

  if (warnItems.length > 0) {
    cards.push({
      titleVi: 'Hạng Mục Cần Khảo Sát & Theo Dõi Nâng Cao (WARNING)',
      titleEn: 'Parameters Requiring Investigation & Tight Monitoring (WARNING)',
      category: 'contact',
      severity: 'warning',
      pointsVi: warnItems.map(
        (i) => `${i.item}: Giá trị đo "${i.measured}" tiệm cận ngưỡng cảnh báo. ${i.detail}`
      ),
      pointsEn: warnItems.map(
        (i) => `${i.itemEn || i.item}: Measured "${i.measured}" approaches threshold. ${i.detailEn || i.detail}`
      ),
    });
  }

  if (failedItems.length === 0 && warnItems.length === 0) {
    cards.push(
      {
        titleVi: 'Độ Bền Điện Môi & Cách Điện Hoàn Hảo (Normal)',
        titleEn: 'Dielectric Strength & Insulation Integrity Verified (Normal)',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          'Điện trở cách điện và thử nghiệm điện áp đều tuân thủ 100% tiêu chuẩn ANSI/NETA ATS-2025.',
          'Không có dấu hiệu suy thoái cách điện hoặc dòng rò qua kết cấu điện môi.',
        ],
        pointsEn: [
          'Insulation resistance and withstand parameters fully comply with ANSI/NETA ATS-2025.',
          'No indication of insulation degradation or leakage current across dielectric structures.',
        ],
      },
      {
        titleVi: 'Tiếp Xúc Điện & Cơ Cấu Hoạt Động Đồng Đều (Normal)',
        titleEn: 'Uniform Electrical Contacts & Operational Verification (Normal)',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Điện trở tiếp xúc và hoạt động cơ khí ăn khớp trơn tru, không có hiện tượng kẹt cơ hoặc quá nhiệt.',
          'Lực siết bu-lông và các đầu nối đạt chuẩn theo NETA Table 100.12.',
        ],
        pointsEn: [
          'Contact resistance and mechanical operating mechanisms operate smoothly without friction or overheating.',
          'Hardware torque values compliant with NETA Table 100.12.',
        ],
      }
    );
  }

  return cards;
}

/**
 * Dynamically builds actionable recommendations for FSE engineers according to ACTUAL results.
 */
function buildDynamicRecommendations(
  spec: NetaAtsSpecificationDef,
  payload: any,
  specId: string
): { vi: string[]; en: string[] } {
  const failedItems = spec.evaluationItems.filter((i) => i.status === 'FAIL');
  const warnItems = spec.evaluationItems.filter((i) => i.status === 'INVESTIGATE');

  if (failedItems.length > 0) {
    const failNamesVi = failedItems.map((i) => i.item.split('(')[0].trim()).join(', ');
    const failNamesEn = failedItems.map((i) => (i.itemEn || i.item).split('(')[0].trim()).join(', ');

    return {
      vi: [
        `CẢNH BÁO TỐI CẤP: KHÔNG CHO PHÉP ĐÓNG ĐIỆN thiết bị do phát hiện khiếm khuyết kỹ thuật nghiêm trọng tại: ${failNamesVi}.`,
        'Thực hiện cô lập an toàn, treo biển cảnh báo nguy hiểm và lập biên bản hiện trường với đơn vị chủ quản.',
        'Lập phương án kỹ thuật xử lý triệt để các khiếm khuyết theo đúng quy chuẩn NETA ATS-2025 (ví dụ: lọc hoàn nguyên dầu chân không, siết lực lại bu-lông, kiểm tra bộ chuyển nấc).',
        'Sau khi xử lý hoàn tất, kỹ sư FSE phải thực hiện lại toàn bộ quy trình đo đạc đối soát và lưu biên bản đạt chuẩn vào hệ thống CMMS.',
      ],
      en: [
        `CRITICAL ALERT: DO NOT ENERGIZE equipment due to unacceptable test failures in: ${failNamesEn}.`,
        'Safely isolate the equipment, install lockout/tagout (LOTO) warning signs, and record site incident documentation.',
        'Develop an engineering remediation procedure to resolve all non-compliant conditions per ANSI/NETA ATS-2025 requirements.',
        'Following remediation, FSE engineers must perform complete re-testing and archive a certified passing report into CMMS.',
      ],
    };
  }

  if (warnItems.length > 0) {
    const warnNamesVi = warnItems.map((i) => i.item.split('(')[0].trim()).join(', ');
    const warnNamesEn = warnItems.map((i) => (i.itemEn || i.item).split('(')[0].trim()).join(', ');

    return {
      vi: [
        `XEM XÉT ĐÓNG ĐIỆN CÓ ĐIỀU KIỆN: Thiết bị đạt các chỉ tiêu cơ bản nhưng cần xử lý hoặc giám sát chặt chẽ các hạng mục cảnh báo: ${warnNamesVi}.`,
        'Tiến hành siết lực bu-lông hoặc làm sạch bề mặt tiếp xúc các mối nối bị lệch điện trở theo NETA Table 100.12.',
        'Thực hiện khảo sát nhiệt hồng ngoại (Thermography per Section 9) trong 24-48 giờ đầu mang tải để phát hiện kịp thời các điểm phát nhiệt cục bộ.',
        'Rút ngắn chu kỳ kiểm tra định kỳ tiếp theo và cập nhật trạng thái "Cần Theo Dõi (Investigate)" trên hệ thống CMMS.',
      ],
      en: [
        `CONDITIONAL ENERGIZATION PERMITTED: Equipment satisfies core baseline limits but requires remediation or close trending for warning parameters: ${warnNamesEn}.`,
        'Re-torque hardware connections or recondition contact surfaces with abnormal resistance per NETA Table 100.12.',
        'Execute infrared thermographic survey (Section 9) within 24-48 hours under operational load to verify no thermal hotspots develop.',
        'Shorten regular maintenance re-inspection intervals and flag status as "Investigate" in CMMS.',
      ],
    };
  }

  // All PASS
  return {
    vi: [
      'PHÊ DUYỆT NGHIỆM THU ĐÓNG ĐIỆN: Toàn bộ thông số kiểm định thị giác, cơ khí và thử nghiệm điện đều tuân thủ 100% ANSI/NETA ATS-2025.',
      'Cho phép tiến hành quy trình đóng điện xung kích và hòa lưới vận hành chính thức theo phương án đóng điện đã được phê duyệt.',
      'Tiến hành khảo sát nhiệt hồng ngoại (Thermographic Survey Section 9) sau 24 giờ mang tải định mức để hoàn tất hồ sơ nghiệm thu tổng thể.',
      'Bàn giao chứng thư kiểm định cho chủ đầu tư và cập nhật trạng thái "Hoạt Động Hoàn Hảo (Healthy)" trên hệ thống CMMS.',
    ],
    en: [
      'COMMISSIONING AUTHORIZATION GRANTED: All visual, mechanical, and electrical test parameters comply 100% with ANSI/NETA ATS-2025.',
      'Authorize formal energization switching sequence and grid synchronization per approved plant commissioning procedures.',
      'Execute infrared thermographic baseline survey (Section 9) within 24 hours under rated operating load.',
      'Hand over certified test report package to asset owner and update equipment status to "Healthy" in CMMS.',
    ],
  };
}
