/**
 * ANSI/NETA ATS-2025 MASTER SPECIFICATIONS KNOWLEDGE BASE
 * Comprehensive definitions, bilingual datasets, evaluation tables, anomaly risks,
 * and actionable recommendations for all 23 Acceptance Testing Specifications (ATS).
 */

export interface NetaAtsTestItem {
  item: string;
  itemEn: string;
  measured: string;
  standard: string;
  standardEn: string;
  reference: string;
  status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  detail: string;
  detailEn: string;
}

export interface NetaAtsAnomalyCard {
  titleVi: string;
  titleEn: string;
  category: 'insulation' | 'contact' | 'thermal' | 'operational';
  severity: 'critical' | 'warning' | 'normal';
  pointsVi: string[];
  pointsEn: string[];
}

export interface NetaAtsSpecificationDef {
  id: string;
  sectionCode: string;
  standardCode: string;
  equipmentTitleVi: string;
  equipmentTitleEn: string;
  subtitleVi: string;
  subtitleEn: string;
  category: 'transformers' | 'switches' | 'cables' | 'motors' | 'protection' | 'renewable' | 'special';
  defaultSiteInfo: {
    project_name: string;
    equipment_tag: string;
    equipment_rating: string;
    fse_name: string;
    test_date: string;
    ambient_temp_c?: number;
    humidity_percent?: number;
  };
  overallStatus: 'PASS' | 'INVESTIGATE' | 'FAIL';
  statusTitleVi: string;
  statusTitleEn: string;
  statusReasonVi: string;
  statusReasonEn: string;
  evaluationItems: NetaAtsTestItem[];
  anomalyCards: NetaAtsAnomalyCard[];
  recommendationsVi: string[];
  recommendationsEn: string[];
  samplePayloadJson: Record<string, any>;
}

export const NETA_ATS_23_MASTER_SPECS: Record<string, NetaAtsSpecificationDef> = {
  // 1. TRANSFORMERS, LIQUID-FILLED (Sec 7.2.2)
  'neta-liquid-transformer': {
    id: 'neta-liquid-transformer',
    sectionCode: 'Mục 7.2.2',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.2.2',
    equipmentTitleVi: 'Máy Biến Áp Ngâm Chất Lỏng Cách Điện (Liquid-Filled Transformers)',
    equipmentTitleEn: 'Liquid-Filled Power Transformers (Mineral Oil / Ester)',
    subtitleVi: 'Thị giác • Cách điện Cuộn dây (Table 100.5) • Tỷ số Biến áp (0.5%) • Dầu ASTM D923 • DGA C57.104',
    subtitleEn: 'Visual • Winding IR (Table 100.5) • Turns Ratio (0.5%) • Oil ASTM D923 • DGA IEEE C57.104',
    category: 'transformers',
    defaultSiteInfo: {
      project_name: 'Trạm Khách Hàng TEV Industrial Substation 110kV',
      equipment_tag: 'TR-OIL-25MVA-01',
      equipment_rating: '25 MVA 110/22kV Class I Mineral Oil Transformer',
      fse_name: 'Võ Minh Tú (sgm1707@gmail.com)',
      test_date: '2026-10-12',
      ambient_temp_c: 28,
      humidity_percent: 65,
    },
    overallStatus: 'FAIL',
    statusTitleVi: 'FAIL / SỰ CỐ DẦU ĐIỆN MÔI & PHÓNG ĐIỆN HỒ QUANG DGA',
    statusTitleEn: 'FAIL / CRITICAL FLUID BREAKDOWN & HIGH-ENERGY ARCING DGA',
    statusReasonVi: 'Không đạt tiêu chuẩn ANSI/NETA ATS-2025: Điện áp đánh thủng dầu sụt giảm xuống 21 kV (< 25 kV), hàm lượng nước đạt 32 ppm (> 20 ppm) và xuất hiện khí Acetylene C2H2 = 6.8 ppm xác nhận phóng điện hồ quang nhiệt độ cao bên trong máy biến áp.',
    statusReasonEn: 'Non-compliant with ANSI/NETA ATS-2025: Dielectric breakdown dropped to 21 kV (< 25 kV), moisture exceeded 20 ppm at 32 ppm, and Acetylene C2H2 reached 6.8 ppm indicating active internal arcing.',
    evaluationItems: [
      {
        item: 'Kiểm tra thị giác & Khớp nhãn mác (IEEE C57.12.00)',
        itemEn: 'Visual & Mechanical Inspection / Nameplate Match',
        measured: 'Khớp 100% bản vẽ, mức dầu 25°C chuẩn',
        standard: 'Khớp 100% bản vẽ; Đệm khí Nitơ duy trì áp suất dương',
        standardEn: 'Match drawings 100%; Maintain positive gas blanket pressure',
        reference: 'NETA ATS-2025 Sec 7.2.2.1',
        status: 'PASS',
        detail: 'Mức dầu tank chính, sứ bushing và áp suất đệm khí Nitơ đạt yêu cầu kỹ thuật.',
        detailEn: 'Main tank fluid levels, bushings, and nitrogen gas blanket within acceptable limits.',
      },
      {
        item: 'Điện trở cách điện Cuộn dây quy đổi 20°C (Table 100.5)',
        itemEn: 'Winding Insulation Resistance Corrected to 20°C',
        measured: 'Pri-Gnd: 3,200 MΩ; Sec-Gnd: 1,100 MΩ; PI = 1.85',
        standard: '≥ 1,000 MΩ (cấp 22kV) / ≥ 5,000 MΩ (cấp 110kV); PI ≥ 1.0',
        standardEn: '≥ 1,000 MΩ (22kV) / ≥ 5,000 MΩ (110kV); Polarization Index PI ≥ 1.0',
        reference: 'NETA ATS-2025 Table 100.5 & Table 100.14',
        status: 'PASS',
        detail: 'Điện trở cách điện cuộn dây và chỉ số phân cực PI = 1.85 vượt xa ngưỡng tối thiểu.',
        detailEn: 'Winding insulation resistance and PI = 1.85 safely exceed minimum limits.',
      },
      {
        item: 'Tỷ số biến áp Turns-Ratio (IEEE C57.12.90)',
        itemEn: 'Turns-Ratio Test Across All Tap Positions',
        measured: 'Độ lệch cực đại: 0.12%',
        standard: 'Độ lệch không vượt quá 0.5% so với tính toán hoặc pha lân cận',
        standardEn: 'Deviation shall not exceed 0.5% compared to calculated turns ratio',
        reference: 'IEEE C57.12.90 & NETA Sec 7.2.2.2',
        status: 'PASS',
        detail: 'Đo trên tất cả các nấc phân áp OLTC/NLTC đều đạt sai số cực nhỏ 0.12%.',
        detailEn: 'All tap changer positions exhibit deviations below 0.5%.',
      },
      {
        item: 'Hệ số tổn hao cách điện Cuộn dây (Power Factor Table 100.3)',
        itemEn: 'Winding Insulation Power Factor at 20°C',
        measured: '0.35%',
        standard: 'Dầu khoáng máy biến áp mới tối đa 0.5% (Tối đa 1.0% cho trạm)',
        standardEn: 'Maximum 0.5% for new oil-filled power transformers (1.0% max)',
        reference: 'NETA ATS-2025 Table 100.3',
        status: 'PASS',
        detail: 'Hệ số tổn hao điện môi cuộn dây 0.35% nằm trong vùng chất lượng tốt.',
        detailEn: 'Winding power factor at 0.35% is well within acceptable limits.',
      },
      {
        item: 'Điện áp đánh thủng dầu cách điện (ASTM D1816 1mm gap)',
        itemEn: 'Dielectric Breakdown Voltage (ASTM D1816 1mm)',
        measured: '21.0 kV',
        standard: 'Dầu khoáng mới nạp máy biến áp BẮT BUỘC ≥ 25 kV (cấp 22kV) / ≥ 35 kV (cấp 110kV)',
        standardEn: 'New insulating mineral oil MUST be ≥ 25 kV (22kV) / ≥ 35 kV (110kV)',
        reference: 'NETA ATS-2025 Table 100.4.1',
        status: 'FAIL',
        detail: 'CẢNH BÁO: Điện áp phóng điện sụt xuống 21 kV, không đảm bảo cách ly điện trường cao.',
        detailEn: 'CRITICAL: Dielectric breakdown at 21 kV is below the mandatory 25 kV limit.',
      },
      {
        item: 'Hàm lượng nước hòa tan trong dầu (ASTM D1533)',
        itemEn: 'Moisture in Insulating Liquid (ASTM D1533 Karl Fischer)',
        measured: '32 ppm',
        standard: 'Tối đa ≤ 20 ppm cho dầu mới đưa vào nghiệm thu',
        standardEn: 'Maximum ≤ 20 ppm for new mineral insulating oil acceptance',
        reference: 'NETA ATS-2025 Table 100.4.1',
        status: 'FAIL',
        detail: 'CẢNH BÁO: Dầu bị ngấm ẩm vượt ngưỡng cho phép 60%, nguy cơ phóng điện vầng quang.',
        detailEn: 'Moisture content at 32 ppm exceeds the 20 ppm acceptance threshold by 60%.',
      },
      {
        item: 'Phân tích khí hòa tan DGA (IEEE C57.104 & ASTM D3612)',
        itemEn: 'Dissolved Gas Analysis (DGA IEEE C57.104)',
        measured: 'C2H2 = 6.8 ppm, H2 = 15 ppm, C2H4 = 38 ppm',
        standard: 'Khí Acetylene C2H2 BẮT BUỘC = 0 ppm (Ngưỡng cảnh báo > 1.0 ppm)',
        standardEn: 'Acetylene C2H2 MUST be 0 ppm (Critical threshold > 1.0 ppm)',
        reference: 'IEEE C57.104 & ASTM D3612',
        status: 'FAIL',
        detail: 'BÁO ĐỘNG HỒ QUANG: Xuất hiện khí C2H2 = 6.8 ppm là bằng chứng phóng điện hồ quang nhiệt độ > 700°C.',
        detailEn: 'HIGH-ENERGY ARCING: Detection of C2H2 = 6.8 ppm indicates arcing exceeding 700°C.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Phóng Điện Hồ Quang Năng Lượng Cao (Arcing Risk)',
        titleEn: 'High-Energy Internal Arcing Anomaly',
        category: 'thermal',
        severity: 'critical',
        pointsVi: [
          'Khí Acetylene (C2H2 = 6.8 ppm) trong máy biến áp nghiệm thu là bất thường tối cấp.',
          'Hồ quang sinh ra từ tiếp điểm bộ chuyển nấc có tải LTC hoặc phóng điện cục bộ khe hở cuộn dây.',
          'Nồng độ Ethylene C2H4 = 38 ppm xác nhận có điểm nóng quá nhiệt cục bộ trên 700°C.',
        ],
        pointsEn: [
          'Detection of Acetylene (C2H2 = 6.8 ppm) during acceptance testing represents a critical defect.',
          'Arcing originated from tap changer contacts or inter-turn winding flashover.',
          'Ethylene C2H4 at 38 ppm confirms localized thermal hot spot exceeding 700°C.',
        ],
      },
      {
        titleVi: 'Suy Thoái Điện Môi & Ngấm Ẩm Dầu (Dielectric Degradation)',
        titleEn: 'Dielectric Breakdown & Moisture Contamination',
        category: 'insulation',
        severity: 'critical',
        pointsVi: [
          'Điện áp đánh thủng ASTM D1816 sụt xuống 21 kV (yêu cầu ≥ 25 kV).',
          'Độ ẩm 32 ppm làm suy giảm khả năng chịu xung sét của giấy cellulose cách điện.',
          'Dầu cần được tuần hoàn qua máy lọc chân không khử ẩm và khí hòa tan khẩn cấp.',
        ],
        pointsEn: [
          'ASTM D1816 breakdown voltage collapsed to 21 kV (minimum requirement is 25 kV).',
          'Moisture at 32 ppm severely impairs paper insulation BIL withstand capability.',
          'Fluid requires immediate vacuum dehydration and degassing processing.',
        ],
      },
      {
        titleVi: 'Cách Điện Cuộn Dây & Chỉ Số Phân Cực (Winding IR)',
        titleEn: 'Winding Insulation & Polarization Index',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          'Điện trở cách điện phía cao thế đạt 3,200 MΩ, hạ thế 1,100 MΩ tại 20°C.',
          'Chỉ số phân cực PI = 1.85 phản ánh cấu trúc cách điện rắn chưa bị già hóa cơ học.',
        ],
        pointsEn: [
          'Primary winding IR measured at 3,200 MΩ, secondary at 1,100 MΩ at 20°C.',
          'Polarization Index PI = 1.85 proves solid cellulose insulation remains intact.',
        ],
      },
      {
        titleVi: 'Cơ Khí, Tỷ Số Biến Áp & Lực Siết Bu-lông',
        titleEn: 'Mechanical Integrity, Turns Ratio & Bolt Torque',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Tất cả các nấc phân áp đạt sai số tỷ số cực đại 0.12% (< 0.5%).',
          'Lực siết bu-lông kiểm tra theo Bảng Table 100.12 đạt 100% giá trị quy định.',
        ],
        pointsEn: [
          'All tap changer steps maintain turns ratio error below 0.12% (< 0.5%).',
          'Calibrated torque wrench verification compliant with NETA Table 100.12.',
        ],
      },
    ],
    recommendationsVi: [
      'CẢNH BÁO TỐI CẤP: TUYỆT ĐỐI KHÔNG ĐÓNG ĐIỆN VẬN HÀNH cho đến khi xử lý triệt để nguyên nhân sinh khí C2H2.',
      'Thực hiện nội soi camera khoang thùng máy và khoang chuyển nấc LTC để xác định vị trí tiếp xúc phóng điện hồ quang.',
      'Lấy mẫu dầu đối soát lại sau 48 giờ để đo tốc độ sinh khí hòa tan (Gas Generation Rate per IEEE C57.104).',
      'Thực hiện lọc sấy chân không (High-Vacuum Degassing & Dehydration) để đưa điện áp phóng điện > 35 kV và độ ẩm < 15 ppm.',
      'Đo Phóng điện cục bộ (PD Test) kết hợp điện dung sứ bushing trước khi xin phê duyệt nghiệm thu lại.',
    ],
    recommendationsEn: [
      'CRITICAL WARNING: DO NOT ENERGIZE the transformer until the root cause of C2H2 arcing is fully rectified.',
      'Perform internal optical borescope inspection of active core, winding assembly, and LTC selector compartment.',
      'Resample oil after 48 hours to determine gas generation rates pursuant to IEEE C57.104.',
      'Deploy mobile high-vacuum oil purification to restore breakdown voltage > 35 kV and moisture < 15 ppm.',
      'Conduct offline Partial Discharge (PD) acoustic/TEV screening prior to re-commissioning approval.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Trạm Khách Hàng TEV Industrial Substation 110kV',
        transformer_tag: 'TR-OIL-25MVA-01',
        transformer_rating: '25 MVA 110/22kV Class I Mineral Oil Transformer',
        fse_name: 'Võ Minh Tú (sgm1707@gmail.com)',
        test_date: '2026-10-12',
      },
      visual_inspection: {
        nameplate_match: true,
        liquid_level_check: 'Pass (Normal mark at 25 deg C)',
        gas_blanket_pressure_positive: true,
        valves_and_cooling_fans_check: 'Pass',
        bolt_torque_check: 'Pass',
      },
      electrical_tests: {
        bolted_resistance_micro_ohms: [8.1, 8.3, 8.2],
        insulation_resistance_20c_megohms: {
          pri_to_ground_5000v: 3200,
          sec_to_ground_2500v: 1100,
          pri_to_sec_5000v_dc: 4100,
          calculated_pi_value: 1.85,
        },
        turns_ratio_all_taps: { max_error_percent: 0.12 },
        winding_insulation_power_factor_percent_20c: 0.35,
        oil_sample_tests_astm_d923: {
          dielectric_breakdown_d1816_1mm_kv: 21,
          water_content_d1533_ppm: 32,
          acid_number_d974_mg_koh_g: 0.02,
          interfacial_tension_d971_mn_m: 39,
          oil_power_factor_25c_d924_percent: 0.04,
        },
        dga_test_ieee_c57_104: {
          hydrogen_h2_ppm: 15,
          acetylene_c2h2_ppm: 6.8,
          ethylene_c2h4_ppm: 38,
          methane_ch4_ppm: 12,
        },
      },
    },
  },

  // 2. DRY-TYPE TRANSFORMERS SMALL <= 500kVA (Sec 7.2.1.1)
  'neta-dry-type': {
    id: 'neta-dry-type',
    sectionCode: 'Mục 7.2.1.1',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.2.1.1',
    equipmentTitleVi: 'Máy Biến Áp Khô Công Suất Nhỏ ≤ 500 kVA (Small Dry-Type Transformers)',
    equipmentTitleEn: 'Small Dry-Type Power Transformers (≤ 500 kVA)',
    subtitleVi: 'Thị giác • Bu-lông DLRO (≤ 50%) • Điện trở Cách điện 20°C (Table 100.5) • Tỷ số Biến áp (≤ 0.5%)',
    subtitleEn: 'Visual • Bolted DLRO (≤ 50%) • Winding IR 20°C (Table 100.5) • Turns Ratio (≤ 0.5%)',
    category: 'transformers',
    defaultSiteInfo: {
      project_name: 'Nhà Máy Điện Tử TEV High-Tech Plant',
      equipment_tag: 'TR-DRY-400KVA-01',
      equipment_rating: '400 kVA 22/0.4kV Cast Resin Dry-Type',
      fse_name: 'Trần Văn Hoàng (hoang.tv@tev.vn)',
      test_date: '2026-10-10',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / ĐẠT TIÊU CHUẨN NGHIỆM THU NETA ATS-2025',
    statusTitleEn: 'PASS / FULLY COMPLIANT WITH NETA ATS-2025 SPECIFICATIONS',
    statusReasonVi: 'Toàn bộ các thông số cách điện cuộn dây quy đổi về 20°C (Pri-Gnd: 4,800 MΩ, Sec-Gnd: 2,100 MΩ), tỷ số biến áp (sai số max 0.15%), điện trở mối nối bu-lông và tiếp địa vỏ máy đều đạt xuất sắc tiêu chuẩn ANSI/NETA ATS-2025 Mục 7.2.1.1.',
    statusReasonEn: 'All parameters including winding insulation resistance corrected to 20°C (Pri-Gnd: 4,800 MΩ, Sec-Gnd: 2,100 MΩ), turns ratio deviations (max 0.15%), and bolted joints conform with NETA ATS-2025 Section 7.2.1.1.',
    evaluationItems: [
      {
        item: 'Kiểm tra thị giác, vệ sinh cuộn đúc & quạt giải nhiệt',
        itemEn: 'Visual & Mechanical / Cast Resin Enclosure Cleanliness',
        measured: 'Sạch sẽ, không nứt nhựa resin, quạt chạy êm',
        standard: 'Cuộn đúc nguyên vẹn, đối lưu thông thoáng, nhãn mác khớp 100%',
        standardEn: 'Intact resin, clean ventilation channels, nameplate match',
        reference: 'NETA ATS-2025 Sec 7.2.1.1.1',
        status: 'PASS',
        detail: 'Bề mặt cuộn dây sạch bụi, quạt làm mát vận hành trơn tru ở chế độ Auto/Manual.',
        detailEn: 'Epoxy resin surfaces clean, cooling fans operational in both auto and manual.',
      },
      {
        item: 'Điện trở mối nối bu-lông DLRO (Table 100.12)',
        itemEn: 'Bolted Connection Resistance via Micro-Ohmmeter',
        measured: '14.2 µΩ (độ lệch pha max 4.5%)',
        standard: 'Độ lệch không vượt quá 50% so với giá trị nhỏ nhất của mối nối tương tự',
        standardEn: 'Deviation shall not exceed 50% of the lowest similar joint resistance',
        reference: 'NETA ATS-2025 Table 100.12',
        status: 'PASS',
        detail: 'Các mối nối thanh cái hạ thế và cao thế siết đủ lực cân chỉnh cờ lê lực.',
        detailEn: 'Bolted bus connections torqued with calibrated torque wrench.',
      },
      {
        item: 'Điện trở cách điện Cuộn Cao thế về 20°C (2500V DC)',
        itemEn: 'Primary Winding Insulation Resistance @ 20°C',
        measured: '4,800 MΩ (Table 100.5 quy định ≥ 1,000 MΩ)',
        standard: 'Cấp 15-25kV BẮT BUỘC ≥ 1,000 MΩ (Table 100.5)',
        standardEn: '15-25kV rating MUST be ≥ 1,000 MΩ (Table 100.5)',
        reference: 'NETA ATS-2025 Table 100.5',
        status: 'PASS',
        detail: 'Cách điện cuộn cao thế đạt gấp 4.8 lần tiêu chuẩn nghiệm thu.',
        detailEn: 'Primary insulation resistance is 4.8 times greater than minimum requirements.',
      },
      {
        item: 'Điện trở cách điện Cuộn Hạ thế về 20°C (1000V DC)',
        itemEn: 'Secondary Winding Insulation Resistance @ 20°C',
        measured: '2,100 MΩ (Table 100.5 quy định ≥ 100 MΩ)',
        standard: 'Cấp 0-600V BẮT BUỘC ≥ 100 MΩ',
        standardEn: '0-600V rating MUST be ≥ 100 MΩ',
        reference: 'NETA ATS-2025 Table 100.5',
        status: 'PASS',
        detail: 'Cách điện cuộn hạ thế 0.4kV đạt rất cao, không bị ẩm bề mặt.',
        detailEn: 'Secondary 0.4kV winding insulation exhibits excellent dry properties.',
      },
      {
        item: 'Tỷ số biến áp Turns Ratio trên tất cả các nấc phân áp',
        itemEn: 'Turns Ratio Verification Across All Tap Settings',
        measured: 'Sai số lớn nhất: 0.15%',
        standard: 'Sai số không vượt quá 0.5% so với tính toán thiết kế',
        standardEn: 'Deviation shall not exceed 0.5% compared to calculated ratio',
        reference: 'IEEE C57.12.91 & NETA ATS Sec 7.2.1.1',
        status: 'PASS',
        detail: 'Đo trên nấc 1 đến nấc 5 đều nằm hoàn toàn trong giới hạn sai số cho phép.',
        detailEn: 'All 5 tap steps remain tightly within 0.15% error margins.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Cách Điện Nhựa Đúc Cast Resin Tối Ưu',
        titleEn: 'Superior Cast Resin Insulation Health',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          'Điện trở cách điện 4,800 MΩ chứng minh môi trường khô ráo, không bị ẩm xâm nhập.',
          'Bề mặt cuộn dây láng mịn, không có vết nứt do ứng suất nhiệt.',
        ],
        pointsEn: [
          'Winding IR at 4,800 MΩ confirms dry and non-condensing operating conditions.',
          'Resin surface is smooth and void of any thermal stress micro-cracks.',
        ],
      },
      {
        titleVi: 'Mối Nối Thanh Cái & Lực Siết Bu-lông',
        titleEn: 'Busbar Connections & Torque Integrity',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Điện trở tiếp xúc cực thấp 14.2 µΩ loại trừ hoàn toàn nguy cơ quá nhiệt cục bộ.',
          'Toàn bộ bu-lông siết chuẩn theo Bảng Table 100.12 và có vạch đánh dấu sơn.',
        ],
        pointsEn: [
          'Low resistance of 14.2 µΩ eliminates risks of localized overheating.',
          'All hardware verified with torque witness marks per Table 100.12.',
        ],
      },
      {
        titleVi: 'Kiểm Soát Nhiệt Độ & Quạt Làm Mát Cưỡng Bức',
        titleEn: 'Temperature Monitoring & Forced Air Cooling',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          'Cảm biến nhiệt độ Pt100 đặt trong 3 cuộn pha đọc chính xác ở 29°C.',
          'Rơ-le bảo vệ nhiệt độ cài đặt ngưỡng cảnh báo 120°C và ngắt 140°C.',
        ],
        pointsEn: [
          'Pt100 RTDs in all three phases read accurately at ambient 29°C.',
          'Temperature controller trips set to 120°C alarm and 140°C trip.',
        ],
      },
      {
        titleVi: 'Độ Nhạy Cân Bằng Pha & Tỷ Số Biến Áp',
        titleEn: 'Phase Balance & Turns Ratio Precision',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Độ lệch tỷ số giữa 3 pha dưới 0.15% bảo đảm điện áp đầu ra đối xứng hoàn hảo.',
        ],
        pointsEn: [
          'Inter-phase turns ratio deviation below 0.15% guarantees symmetrical secondary output.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện đóng điện nghiệm thu (Ready for Energization).',
      'Đóng điện không tải 24 giờ và kiểm tra lại tiếng ồn rung từ lõi thép từ tính.',
      'Thực hiện kiểm tra camera nhiệt hồng ngoại (Thermography) sau khi mang tải 50% trong 2 giờ đầu tiên.',
      'Bảo trì định kỳ vệ sinh thổi bụi khô bề mặt cuộn dây sau 6 tháng vận hành.',
    ],
    recommendationsEn: [
      'Approved for acceptance energization.',
      'Perform 24-hour no-load soak test to verify magnetic core acoustic levels.',
      'Conduct baseline infrared thermographic survey within first 2 hours of reaching 50% load.',
      'Schedule clean dry air blowing maintenance after 6 months of continuous operation.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Nhà Máy Điện Tử TEV High-Tech Plant',
        transformer_tag: 'TR-DRY-400KVA-01',
        transformer_rating: '400 kVA 22/0.4kV Cast Resin Dry-Type',
        fse_name: 'Trần Văn Hoàng (hoang.tv@tev.vn)',
        test_date: '2026-10-10',
      },
      visual_inspection: {
        nameplate_match: true,
        cleanliness_and_cracks_check: 'Pass (Clean surfaces, no resin cracks)',
        cooling_fans_functional: true,
        bolt_torque_check: 'Pass per Table 100.12',
      },
      electrical_tests: {
        bolted_connection_resistance_micro_ohms: [14.1, 14.5, 14.2],
        insulation_resistance_20c_megohms: {
          pri_to_ground_2500v: 4800,
          sec_to_ground_1000v: 2100,
          pri_to_sec_2500v: 5200,
        },
        turns_ratio_error_percent: 0.15,
      },
    },
  },

  // 3. SWITCHGEAR & SWITCHBOARD ASSEMBLIES (Sec 7.1.1)
  'neta-switchgear': {
    id: 'neta-switchgear',
    sectionCode: 'Mục 7.1.1',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.1.1',
    equipmentTitleVi: 'Tủ Điện Phân Phối & Tủ Đóng Cắt (Switchgear & Switchboard Assemblies)',
    equipmentTitleEn: 'Switchgear & Switchboard Assemblies (MV Metal-Clad / LV MSB)',
    subtitleVi: 'Khóa Liên Động • Bu-lông DLRO (≤ 50%) • Mạch Điều Khiển (≥ 2.0 MΩ) • Cách Điện Busbar (Table 100.1)',
    subtitleEn: 'Interlocks • Bus DLRO (≤ 50%) • Control Wiring (≥ 2.0 MΩ) • Busbar IR (Table 100.1)',
    category: 'switches',
    defaultSiteInfo: {
      project_name: 'Trạm Phân Phối 24kV TEV Data Center',
      equipment_tag: 'SWG-24KV-MAIN-01',
      equipment_rating: '24kV 1250A 25kA Metal-Clad Switchgear',
      fse_name: 'Phạm Quốc Bảo (bao.pq@tev.vn)',
      test_date: '2026-10-11',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / ĐẠT TIÊU CHUẨN NGHIỆM THU NETA ATS-2025 MỤC 7.1.1',
    statusTitleEn: 'PASS / FULLY COMPLIANT WITH NETA ATS-2025 SECTION 7.1.1',
    statusReasonVi: 'Toàn bộ liên động cơ điện, điện trở tiếp xúc thanh cái (độ lệch 3.2% ≤ 50%), cách điện busbar 24kV (18,500 MΩ ≥ 1,000 MΩ), và cách điện mạch điều khiển nhị thứ (45 MΩ ≥ 2.0 MΩ) đều đạt chuẩn nghiêm ngặt.',
    statusReasonEn: 'Mechanical interlocks, busbar joint resistance (3.2% deviation ≤ 50%), 24kV busbar insulation (18,500 MΩ ≥ 1,000 MΩ), and secondary control circuit insulation (45 MΩ ≥ 2.0 MΩ) comply with Section 7.1.1.',
    evaluationItems: [
      {
        item: 'Kiểm tra cơ khí, căn chỉnh khoang & Khóa liên động (Interlocks)',
        itemEn: 'Mechanical Alignment & Interlock Verification',
        measured: 'Cửa sập đóng mở trơn tru, khóa trao đổi chìa an toàn 100%',
        standard: 'Cơ cấu rút máy cắt drawout, tiếp địa cưỡng bức và cửa sập hoạt động hoàn hảo',
        standardEn: 'Drawout racks, shutter mechanisms, and key interlocks smooth',
        reference: 'NETA ATS-2025 Sec 7.1.1.1',
        status: 'PASS',
        detail: 'Thao tác đẩy máy cắt vào vị trí Test / Connected nhẹ nhàng, tiếp xúc ngàm cắm chắc chắn.',
        detailEn: 'Breaker drawout operations smooth across Test and Connected positions.',
      },
      {
        item: 'Điện trở mối nối thanh cái Busbar DLRO (Table 100.12)',
        itemEn: 'Busbar Bolted Joint Resistance (DLRO Test)',
        measured: '8.4 µΩ, 8.6 µΩ, 8.5 µΩ (Lệch max 2.4%)',
        standard: 'Độ lệch không vượt quá 50% so với giá trị nhỏ nhất của mối nối cùng loại',
        standardEn: 'Joint resistance deviation must not exceed 50% of the lowest joint',
        reference: 'NETA ATS-2025 Sec 7.1.1.2 & Table 100.12',
        status: 'PASS',
        detail: 'Đo điện trở một chiều tiếp xúc thanh đồng mạ bạc đạt cực thấp, dẫn dòng tối ưu.',
        detailEn: 'Silver-plated copper bus connections demonstrate superior conductivity.',
      },
      {
        item: 'Điện trở cách điện Thanh Cái Busbar 24kV quy đổi 20°C (2500V DC)',
        itemEn: '24kV Main Busbar Insulation Resistance Corrected to 20°C',
        measured: '18,500 MΩ (Table 100.1 quy định ≥ 1,000 MΩ)',
        standard: 'Thanh cái trung thế cấp 15-25kV BẮT BUỘC ≥ 1,000 MΩ',
        standardEn: 'Medium-voltage busbar MUST be ≥ 1,000 MΩ per Table 100.1',
        reference: 'NETA ATS-2025 Table 100.1',
        status: 'PASS',
        detail: 'Cách điện thanh cái pha-pha và pha-đất vượt gấp 18.5 lần ngưỡng tối thiểu.',
        detailEn: 'Phase-to-phase and phase-to-ground IR exceeds minimum limit by 18.5x.',
      },
      {
        item: 'Điện trở cách điện Mạch điều khiển & Mạch nhị thứ (500V DC)',
        itemEn: 'Secondary Control Circuit Wiring Insulation Resistance',
        measured: '45.0 MΩ (Yêu cầu BẮT BUỘC ≥ 2.0 MΩ)',
        standard: 'Mạch điều khiển nhị thứ BẮT BUỘC ≥ 2.0 MΩ (Mục 7.1.1.D.4)',
        standardEn: 'Secondary control wiring MUST be ≥ 2.0 MΩ (Section 7.1.1.D.4)',
        reference: 'NETA ATS-2025 Sec 7.1.1.D.4',
        status: 'PASS',
        detail: 'Toàn bộ dây nhị thứ đấu nối rơ-le, đồng hồ và mạch cắt trip đều đạt cách điện an toàn.',
        detailEn: 'Trip circuits, relays, and auxiliary control wiring pass 500V DC insulation test.',
      },
      {
        item: 'Khảo sát Phóng điện cục bộ TEV / Siêu âm (Partial Discharge Screening)',
        itemEn: 'Online Transient Earth Voltage (TEV) & Acoustic PD Scan',
        measured: 'TEV: 4 dBmV (Bình thường < 20 dBmV); Ultrasonic: 0 dBµV',
        standard: 'Mức TEV nền an toàn < 20 dBmV; Không có xung phóng điện bề mặt',
        standardEn: 'Background TEV < 20 dBmV; Zero surface tracking acoustic signals',
        reference: 'NETA ATS-2025 Section 9 & IEC 62478',
        status: 'PASS',
        detail: 'Không phát hiện bất kỳ tín hiệu phóng điện cục bộ nào trong khoang thanh cái.',
        detailEn: 'No internal void or surface tracking partial discharges detected.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Khóa Liên Động An Toàn & Cơ Khí Vận Hành',
        titleEn: 'Mechanical Interlocks & Operational Safety',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Cửa sập tự động (Automatic shutters) đóng kín khi kéo máy cắt ra vị trí Test.',
          'Khóa liên động dao tiếp địa cơ khí ngăn chặn đóng máy cắt khi dao tiếp địa đang đóng.',
        ],
        pointsEn: [
          'Automatic shutters seal off live bus spouts when breaker is racked to Test.',
          'Earthing switch interlocks prevent breaker closure while grounded.',
        ],
      },
      {
        titleVi: 'Chất Lượng Tiếp Xúc Thanh Cái & Độ Lệch DLRO',
        titleEn: 'Busbar Contact Quality & DLRO Deviation',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Độ lệch điện trở giữa các mối nối chỉ 2.4%, thấp hơn rất nhiều so với ngưỡng 50%.',
          'Các đệm lò xo Belleville được nén đúng độ căng thiết kế.',
        ],
        pointsEn: [
          'Joint-to-joint resistance variation of only 2.4% is far below the 50% alert threshold.',
          'Belleville spring washers compressed to manufacturer specifications.',
        ],
      },
      {
        titleVi: 'Độ Bền Cách Điện Nhị Thứ & Mạch Đo Lường',
        titleEn: 'Secondary Control & Metering Circuit Integrity',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          'Cách điện mạch nhị thứ 45 MΩ loại trừ nguy cơ rò điện gây nhảy CB bảo vệ nguồn DC.',
          'Biến dòng đo lường CT và biến điện áp PT được thử nghiệm tỷ số đạt chuẩn.',
        ],
        pointsEn: [
          '45 MΩ secondary IR eliminates DC ground fault alarms.',
          'Instrument CT and PT ratios verified per ANSI standards.',
        ],
      },
      {
        titleVi: 'Sàng Lọc Xung Phóng Điện Cục Bộ Online TEV',
        titleEn: 'Online Transient Earth Voltage (TEV) PD Health',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          'Mức TEV 4 dBmV khẳng định các sứ cách điện epoxy đúc không có bọt khí khuyết tật.',
        ],
        pointsEn: [
          'TEV background at 4 dBmV confirms void-free epoxy insulators.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện đóng điện nghiệm thu đưa vào vận hành.',
      'Đóng nắp chụp bảo vệ toàn bộ các khoang cáp lực và thanh cái trước khi cấp điện.',
      'Kiểm tra chiều quay quạt hút gió cưỡng bức phòng điện trung thế.',
      'Lập lịch kiểm tra nhiệt hồng ngoại sau 24 giờ mang tải.',
    ],
    recommendationsEn: [
      'Approved for acceptance energization and commercial operation.',
      'Install all protective safety covers over busbars and cable compartments.',
      'Verify substation ventilation and pressure relief flaps.',
      'Schedule follow-up infrared thermographic scan at 24 hours under load.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Trạm Phân Phối 24kV TEV Data Center',
        equipment_tag: 'SWG-24KV-MAIN-01',
        equipment_rating: '24kV 1250A 25kA Metal-Clad Switchgear',
        fse_name: 'Phạm Quốc Bảo (bao.pq@tev.vn)',
        test_date: '2026-10-11',
      },
      visual_inspection: {
        enclosure_and_alignment: 'Pass',
        interlocks_functional: true,
        shutter_mechanism: 'Pass',
        bolt_torque_check: 'Pass per Table 100.12',
      },
      electrical_tests: {
        bolted_bus_resistance_micro_ohms: [8.4, 8.6, 8.5],
        busbar_insulation_resistance_2500v_megohms: 18500,
        control_circuit_insulation_500v_megohms: 45.0,
        tev_partial_discharge_dbmv: 4.0,
      },
    },
  },

  // 4. LOW-VOLTAGE POWER CIRCUIT BREAKERS ACB (Sec 7.6.1.2)
  'neta-lv-breaker': {
    id: 'neta-lv-breaker',
    sectionCode: 'Mục 7.6.1.2',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.6.1.2',
    equipmentTitleVi: 'Máy Cắt Không Khí Hạ Áp (Low-Voltage Power Circuit Breakers - ACB)',
    equipmentTitleEn: 'Low-Voltage Power Circuit Breakers (Air Circuit Breaker - ACB)',
    subtitleVi: 'Tiếp Điểm Cực (≤ 50%) • Bơm Dòng Trip Unit L, S, I, G • Cách Điện Table 100.1 • Khóa Liên Động',
    subtitleEn: 'Pole Resistance (≤ 50%) • Trip Unit Injection (L, S, I, G) • Insulation Table 100.1 • Drawout Mechanism',
    category: 'switches',
    defaultSiteInfo: {
      project_name: 'Nhà Máy Hóa Chất TEV Chemical Plant',
      equipment_tag: 'ACB-MAIN-01',
      equipment_rating: '4000A 690V 85kA 3-Pole Drawout ACB with Micrologic',
      fse_name: 'Lê Minh Tâm (tam.lm@tev.vn)',
      test_date: '2026-10-09',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / MÁY CẮT ĐẠT TIÊU CHUẨN NGHIỆM THU NETA ATS-2025',
    statusTitleEn: 'PASS / ACB FULLY COMPLIANT WITH NETA ATS-2025 ACCEPTANCE STANDARDS',
    statusReasonVi: 'Điện trở tiếp xúc cực đo bằng dòng một chiều 100A (Pha A: 18.2 µΩ, B: 18.8 µΩ, C: 18.4 µΩ, lệch cực đại 3.3% ≤ 50%), thời gian tác động bơm dòng thứ cấp Trip Unit L-S-I-G khớp đường cong đặc tính, và cách điện cực đạt 1,450 MΩ.',
    statusReasonEn: 'Contact resistance at 100A DC (Phase A: 18.2 µΩ, B: 18.8 µΩ, C: 18.4 µΩ, max deviation 3.3% ≤ 50%), secondary injection trip timing matches manufacturer TCC, and insulation resistance is 1,450 MΩ.',
    evaluationItems: [
      {
        item: 'Điện trở tiếp xúc cực chính (Contact Resistance via DLRO @ 100A)',
        itemEn: 'Main Contact Resistance per Pole (100A DC DLRO)',
        measured: 'Pha A: 18.2 µΩ | Pha B: 18.8 µΩ | Pha C: 18.4 µΩ',
        standard: 'Độ lệch không vượt quá 50% so với giá trị nhỏ nhất của các cực',
        standardEn: 'Deviation shall not exceed 50% of the lowest pole reading',
        reference: 'NETA ATS-2025 Sec 7.6.1.2.D.2',
        status: 'PASS',
        detail: 'Độ lệch tiếp xúc giữa các cực chỉ 3.3%, tiếp điểm bạc phẳng mịn không oxy hóa.',
        detailEn: 'Max pole-to-pole variation is only 3.3%, well within 50% standard.',
      },
      {
        item: 'Thử nghiệm Bơm Dòng Thứ Cấp (Secondary Current Injection Trip Test)',
        itemEn: 'Secondary Current Injection Trip Unit Verification (L, S, I, G)',
        measured: 'Long-time: 12.4s | Short-time: 0.31s | Instantaneous: 28ms | Ground-fault: 0.22s',
        standard: 'Thời gian cắt nằm trong dải dung sai đường cong phối hợp TCC của nhà sản xuất',
        standardEn: 'Trip timing must fall within manufacturer Time-Current Curve (TCC) tolerances',
        reference: 'NETA ATS-2025 Sec 7.6.1.2.D.3',
        status: 'PASS',
        detail: 'Bộ bảo vệ điện tử tác động chính xác ở các ngưỡng bảo vệ quá tải, ngắn mạch và chạm đất.',
        detailEn: 'Electronic trip unit performs accurately across Long, Short, Instantaneous, and Ground.',
      },
      {
        item: 'Điện trở cách điện Cực Máy Cắt (Pole-to-Pole & Pole-to-Ground @ 1000V)',
        itemEn: 'Insulation Resistance Pole-to-Pole & Pole-to-Ground',
        measured: '1,450 MΩ (Table 100.1 quy định tối thiểu ≥ 100 MΩ)',
        standard: 'Máy cắt hạ thế 600V BẮT BUỘC ≥ 100 MΩ',
        standardEn: 'LV power breakers MUST exceed ≥ 100 MΩ per Table 100.1',
        reference: 'NETA ATS-2025 Table 100.1',
        status: 'PASS',
        detail: 'Thử nghiệm cả ở trạng thái đóng và mở tiếp điểm đều đạt cách điện an toàn.',
        detailEn: 'Insulation tests in both open and closed positions pass with high margins.',
      },
      {
        item: 'Cơ cấu rút Drawout, Tiếp điểm phụ & Buồng dập hồ quang',
        itemEn: 'Drawout Racking, Auxiliary Contacts & Arc Chutes Inspection',
        measured: 'Buồng dập hồ quang sạch, lá thép chia hồ quang không nứt vỡ',
        standard: 'Buồng dập hồ quang nguyên vẹn, các ngàm cắm finger clusters đàn hồi tốt',
        standardEn: 'Clean de-ion arc chutes, finger clusters maintain good spring tension',
        reference: 'NETA ATS-2025 Sec 7.6.1.2.1',
        status: 'PASS',
        detail: 'Thao tác nạp lò xo bằng tay và motor nạp điện hoạt động êm ái.',
        detailEn: 'Manual and electrical spring charging motors cycle smoothly.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Chất Lượng Tiếp Điểm Cực & Khả Năng Mang Dòng',
        titleEn: 'Main Pole Contact Health & Current Capacity',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Điện trở 18.2 - 18.8 µΩ bảo đảm máy cắt mang dòng định mức 4000A mà không sinh nhiệt quá mức.',
          'Độ lệch pha 3.3% khẳng định lực ép lò xo tiếp điểm cân bằng tuyệt đối.',
        ],
        pointsEn: [
          'Contact resistance of 18.2 - 18.8 µΩ permits continuous 4000A operation without overheating.',
          '3.3% variation proves perfectly balanced contact pressure springs.',
        ],
      },
      {
        titleVi: 'Độ Tin Cậy Cắt Bảo Vệ Ngắn Mạch & Chạm Đất',
        titleEn: 'Overcurrent & Ground Fault Protection Precision',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Thời gian cắt tức thì Instantaneous chỉ 28ms bảo vệ an toàn chống hồ quang điện Arc Flash.',
          'Chức năng chạm đất Ground Fault cắt dứt khoát ở 0.22s.',
        ],
        pointsEn: [
          'Instantaneous clearing time of 28ms maximizes Arc Flash personal safety.',
          'Ground fault protection isolates zero-sequence faults cleanly at 0.22s.',
        ],
      },
      {
        titleVi: 'Buồng Dập Hồ Quang Arc Chute & Ngàm Cắm',
        titleEn: 'Arc Chute Condition & Finger Clusters',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          'Không có vết muội than hoặc biến dạng trên các phiến tản nhiệt buồng dập hồ quang.',
          'Ngàm cắm drawout được bôi mỡ dẫn điện bảo vệ chống mài mòn.',
        ],
        pointsEn: [
          'No soot deposits or magnetic plate distortion in arc extinguishing chambers.',
          'Conductive grease applied to drawout finger clusters.',
        ],
      },
      {
        titleVi: 'Điện Trở Cách Điện Vỏ Máy & Cực Pha',
        titleEn: 'Dielectric Insulation Withstand',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          'Cách điện 1,450 MΩ loại trừ nguy cơ phóng điện rò pha-vỏ trong điều kiện ẩm.',
        ],
        pointsEn: [
          '1,450 MΩ insulation eliminates phase-to-ground flashover risks.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện đóng điện vận hành vào hệ thống phân phối MSB.',
      'Cài đặt khóa vị trí Drawout và bàn giao chìa khóa Interlock cho trưởng ca vận hành.',
      'Dán nhãn Arc Flash Warning theo tính toán IEEE 1584 tại mặt trước tủ máy cắt.',
      'Thử nghiệm cơ khí kiểm tra cắt khẩn cấp Emergency Pushbutton bên ngoài tủ.',
    ],
    recommendationsEn: [
      'Approved for MSB distribution energization.',
      'Engage drawout safety padlock and hand over interlock keys to substation lead.',
      'Affix IEEE 1584 Arc Flash warning labels on breaker enclosure door.',
      'Perform functional trip check via external Emergency Stop pushbutton.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Nhà Máy Hóa Chất TEV Chemical Plant',
        equipment_tag: 'ACB-MAIN-01',
        equipment_rating: '4000A 690V 85kA 3-Pole Drawout ACB with Micrologic',
        fse_name: 'Lê Minh Tâm (tam.lm@tev.vn)',
        test_date: '2026-10-09',
      },
      visual_inspection: {
        arc_chutes_intact: true,
        drawout_mechanism_smooth: true,
        finger_clusters_good: true,
        bolt_torque_check: 'Pass',
      },
      electrical_tests: {
        pole_contact_resistance_micro_ohms: [18.2, 18.8, 18.4],
        insulation_resistance_1000v_megohms: 1450,
        trip_unit_timing_seconds: {
          long_time: 12.4,
          short_time: 0.31,
          instantaneous: 0.028,
          ground_fault: 0.22,
        },
      },
    },
  },

  // 5. GROUNDING SYSTEMS (Sec 7.13 & IEEE 81)
  'neta-grounding': {
    id: 'neta-grounding',
    sectionCode: 'Mục 7.13',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.13 & IEEE 81',
    equipmentTitleVi: 'Hệ Thống Tiếp Địa An Toàn & Lưới Đất Trạm (Grounding Systems)',
    equipmentTitleEn: 'Grounding Systems (Substation Grid & Point-to-Point Bonding)',
    subtitleVi: 'Điện Trở Lưới Đất Fall-of-Potential (≤ 1.0 Ω) • Tiếp Địa Đẳng Thế (< 0.5 Ω) • Hàng Rào Trạm',
    subtitleEn: 'Fall-of-Potential Grid Resistance (≤ 1.0 Ω) • Point-to-Point Bonding (< 0.5 Ω) • Fence Grounding',
    category: 'special',
    defaultSiteInfo: {
      project_name: 'Trạm Biến Áp 110kV TEV Factory Substation',
      equipment_tag: 'GRID-EARTH-MAIN-01',
      equipment_rating: 'Copper Mesh Grounding Grid with 50x Copper-Clad Steel Rods',
      fse_name: 'Nguyễn Văn X (x.nv@tev.vn)',
      test_date: '2026-09-26',
    },
    overallStatus: 'FAIL',
    statusTitleVi: 'FAIL / ĐIỆN TRỞ LƯỚI ĐẤT VƯỢT QUÁ GIỚI HẠN AN TOÀN IEEE 81',
    statusTitleEn: 'FAIL / GRID RESISTANCE EXCEEDS IEEE 81 & NETA 7.13 LIMITS',
    statusReasonVi: 'Điện trở tiếp địa đo theo phương pháp Fall-of-Potential đạt 7.8 Ω (vượt quá ngưỡng an toàn trạm biến áp ≤ 1.0 Ω quy định tại Mục 7.13.C.2), và điện trở đẳng thế hàng rào trạm đạt 0.68 Ω (> 0.5 Ω) gây nguy cơ điện áp chạm Touch Voltage nguy hiểm.',
    statusReasonEn: 'Substation ground grid resistance measured via Fall-of-Potential is 7.8 Ω (exceeding ≤ 1.0 Ω substation limit per Section 7.13.C.2), and fence bonding is 0.68 Ω (> 0.5 Ω) creating touch potential hazards.',
    evaluationItems: [
      {
        item: 'Điện trở lưới tiếp địa chính Fall-of-Potential (IEEE Std 81)',
        itemEn: 'Main Ground Grid Resistance (Fall-of-Potential IEEE 81)',
        measured: '7.8 Ω (Đo tại khoảng cách 62% tuyến cọc thử nghiệm)',
        standard: 'Trạm biến áp công nghiệp BẮT BUỘC ≤ 1.0 Ω (Lưới chung ≤ 5.0 Ω)',
        standardEn: 'Substation ground grid MUST be ≤ 1.0 Ω (General systems ≤ 5.0 Ω)',
        reference: 'NETA ATS-2025 Sec 7.13.C.2 & IEEE 81',
        status: 'FAIL',
        detail: 'NGUY HIỂM: Điện trở 7.8 Ω không đảm bảo tản dòng sự cố chạm đất 110kV, nguy cơ nâng cao thế lưới đất GPR.',
        detailEn: 'CRITICAL: 7.8 Ω will cause dangerous Ground Potential Rise (GPR) during 110kV ground faults.',
      },
      {
        item: 'Liên kết đẳng thế điểm-đến-điểm (Point-to-Point Bonding Resistance)',
        itemEn: 'Point-to-Point Equipment Bonding Resistance',
        measured: 'Thanh chính tới MBA: 12.5 mΩ | Tủ SWG: 14.2 mΩ | Nhà ĐK: 15.0 mΩ',
        standard: 'Điện trở liên kết đẳng thế BẮT BUỘC < 0.5 Ω (500 mΩ)',
        standardEn: 'Bonding resistance MUST be < 0.5 Ω (500 mΩ) per Section 7.13.D.1',
        reference: 'NETA ATS-2025 Sec 7.13.D.1',
        status: 'PASS',
        detail: 'Tất cả các vỏ thiết bị chính liên kết chắc chắn vào thanh cái tiếp địa đồng trần 240mm².',
        detailEn: 'Main equipment frames securely bonded to copper grounding loop.',
      },
      {
        item: 'Tiếp địa hàng rào & Cổng trạm biến áp (Fence & Gate Bonding)',
        itemEn: 'Substation Fence & Gate Bonding Continuity',
        measured: '0.68 Ω (680 mΩ)',
        standard: 'Điện trở liên kết hàng rào và cổng trạm BẮT BUỘC ≤ 0.5 Ω',
        standardEn: 'Fence and gate ground bonding MUST be ≤ 0.5 Ω',
        reference: 'NETA ATS-2025 Sec 7.13.D.2 & IEEE 80',
        status: 'FAIL',
        detail: 'CẢNH BÁO ĐIỆN ÁP CHẠM: Dây bện đồng mềm tiếp địa cánh cổng bị lỏng tiếp xúc, điện trở vượt mức 0.5 Ω.',
        detailEn: 'TOUCH VOLTAGE HAZARD: Flexible copper jumper at perimeter gate is loose (> 0.5 Ω).',
      },
      {
        item: 'Kiểm tra cơ khí, mối hàn Cadweld & Giếng kiểm tra tiếp địa',
        itemEn: 'Visual & Mechanical Inspection / Cadweld Joints',
        measured: 'Mối hàn hóa nhiệt ngấu chắc, hố tiếp địa có nắp đậy bảo vệ',
        standard: 'Mối hàn Cadweld không rỗ xỉ, cáp đồng 240mm² không đứt gãy',
        standardEn: 'Exothermic welds solid without porosity, 240mm² copper cable intact',
        reference: 'NETA ATS-2025 Sec 7.13.1',
        status: 'PASS',
        detail: 'Các mối hàn hóa nhiệt đồng vào cọc thép bọc đồng được thi công đúng quy trình kỹ thuật.',
        detailEn: 'Cadweld exothermic joints between copper conductors and ground rods are sound.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Nguy Cơ Nâng Cao Thế Lưới Đất & Điện Áp Chạm (GPR & Touch Voltage)',
        titleEn: 'Ground Potential Rise & Touch Voltage Hazards',
        category: 'operational',
        severity: 'critical',
        pointsVi: [
          'Điện trở 7.8 Ω khi xảy ra ngắn mạch chạm đất 110kV (dòng 10kA) sẽ làm điện thế lưới đất tăng vọt lên 78,000V.',
          'Điện thế bước và điện thế chạm sẽ vượt ngưỡng gây nguy hiểm tính mạng cho công nhân vận hành.',
        ],
        pointsEn: [
          '7.8 Ω grid resistance produces a dangerous 78 kV Ground Potential Rise during 10kA fault.',
          'Step and touch potentials will exceed human fibrillation limits per IEEE 80.',
        ],
      },
      {
        titleVi: 'Hàng Rào Bảo Vệ & Dây Bện Cổng Trạm (Fence Jumper Defect)',
        titleEn: 'Perimeter Fence & Gate Jumper Defect',
        category: 'contact',
        severity: 'critical',
        pointsVi: [
          'Dây nối đất cánh cổng 0.68 Ω (> 0.5 Ω) khiến cánh cổng trạm tích điện khi có giông sét hoặc sự cố.',
          'Cần siết lại đầu cosse hoặc thay mới dây đồng bện mềm tiếp địa cánh cổng.',
        ],
        pointsEn: [
          '0.68 Ω gate jumper creates touch hazards during lightning strikes or ground faults.',
          'Requires immediate re-torquing or replacement with flexible copper braid.',
        ],
      },
      {
        titleVi: 'Khảo Sát Điện Trở Suất Đất (Soil Resistivity Analysis)',
        titleEn: 'Soil Resistivity & Geological Factors',
        category: 'thermal',
        severity: 'warning',
        pointsVi: [
          'Điện trở suất đất đo bằng phương pháp Wenner 4 cực đạt 145 Ω·m (mùa khô).',
          'Độ sâu cọc tiếp địa hiện tại (3m) chưa chạm đến tầng đất ẩm dẫn điện tốt bên dưới.',
        ],
        pointsEn: [
          'Wenner 4-pin soil resistivity measured at 145 Ω·m under dry conditions.',
          'Existing 3-meter rods do not penetrate deep water table layers.',
        ],
      },
      {
        titleVi: 'Liên Kết Đẳng Thế Vỏ Thiết Bị Đạt Chuẩn',
        titleEn: 'Equipment Frame Equipotential Continuity',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Các thiết bị bên trong trạm (MBA, Switchgear, tủ điều khiển) liên kết đồng bộ đạt < 15 mΩ.',
        ],
        pointsEn: [
          'Indoor assets maintain robust continuity under 15 mΩ back to main perimeter earth bar.',
        ],
      },
    ],
    recommendationsVi: [
      'CẢNH BÁO AN TOÀN: KHÔNG CHO PHÉP ĐÓNG ĐIỆN TRẠM 110kV cho đến khi điện trở lưới đất đạt ≤ 1.0 Ω.',
      'Khoan bổ sung 4-6 giếng tiếp địa sâu (Deep Ground Wells 20-30m) kết hợp hợp chất giảm điện trở đất GEM.',
      'Thay mới dây đồng bện mềm 50mm² nối tiếp địa giữa trụ cổng và cánh cổng, đo lại điện trở đạt < 0.1 Ω.',
      'Đo lại Fall-of-Potential theo đường cong 62% với khoảng cách điện cực dòng tối thiểu 150m.',
    ],
    recommendationsEn: [
      'SAFETY HOLD: DO NOT ENERGIZE 110kV substation until grid resistance is reduced to ≤ 1.0 Ω.',
      'Drill 4 to 6 deep earth wells (20-30m deep) backfilled with Ground Enhancement Material (GEM).',
      'Replace gate bonding flexible jumpers with 50mm² tinned copper braid; target < 0.1 Ω.',
      'Rerun Fall-of-Potential 62% distance curve with current probe spacing of at least 150 meters.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Trạm Biến Áp 110kV TEV Factory Substation',
        grounding_system_tag: 'GRID-EARTH-MAIN-01',
        fse_name: 'Nguyễn Văn X (x.nv@tev.vn)',
        test_date: '2026-09-26',
      },
      visual_inspection: {
        grounding_layout_match: true,
        cadweld_joints_solid: true,
        bolt_torque_check: 'Pass',
      },
      electrical_tests: {
        fall_of_potential_resistance_ohms: 7.8,
        point_to_point_continuity_milli_ohms: {
          main_bus_to_transformer_frame: 12.5,
          main_bus_to_switchgear: 14.2,
          main_bus_to_control_room: 15.0,
        },
        fence_gate_bonding_ohms: 0.68,
      },
    },
  },
  // 6. LARGE DRY-TYPE TRANSFORMERS > 500kVA (Sec 7.2.1.2)
  'neta-dry-large': {
    id: 'neta-dry-large',
    sectionCode: 'Mục 7.2.1.2',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.2.1.2',
    equipmentTitleVi: 'Máy Biến Áp Khô Công Suất Lớn > 500 kVA (Large Dry-Type Transformers)',
    equipmentTitleEn: 'Large Dry-Type Transformers (> 500 kVA)',
    subtitleVi: 'Thị giác • Khảo sát nhiệt • Điện trở cuộn dây 20°C (Table 100.5) • Power Factor cuộn dây (Table 100.3) • Lõi thép Core IR',
    subtitleEn: 'Visual • Thermography • Winding IR 20°C (Table 100.5) • Winding Power Factor (Table 100.3) • Core IR',
    category: 'transformers',
    defaultSiteInfo: {
      project_name: 'Trung Tâm Dữ Liệu TEV Hyperscale DC',
      equipment_tag: 'TR-DRY-2500KVA-01',
      equipment_rating: '2500 kVA 22/0.4kV VPI Cast-Coil Transformer',
      fse_name: 'Đặng Tuấn Anh (anh.dt@tev.vn)',
      test_date: '2026-10-08',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / ĐẠT TIÊU CHUẨN NGHIỆM THU NETA ATS-2025 MỤC 7.2.1.2',
    statusTitleEn: 'PASS / FULLY COMPLIANT WITH NETA ATS-2025 SECTION 7.2.1.2',
    statusReasonVi: 'Điện trở cách điện cuộn dây quy đổi 20°C đạt 8,500 MΩ, hệ số tổn hao điện môi Power Factor cuộn dây 0.68% (ngưỡng tối đa ≤ 2.0% theo Table 100.3), điện trở cách điện lõi thép Core IR đạt 420 MΩ (≥ 100 MΩ), và tỷ số phân áp đạt sai số 0.11%.',
    statusReasonEn: 'Winding IR corrected to 20°C is 8,500 MΩ, Winding Power Factor is 0.68% (≤ 2.0% per Table 100.3), Core IR is 420 MΩ (≥ 100 MΩ), and turns ratio error is 0.11%.',
    evaluationItems: [
      {
        item: 'Điện trở cách điện Cuộn Cao thế 20°C (Table 100.5 @ 2500V DC)',
        itemEn: 'Primary Winding Insulation Resistance @ 20°C',
        measured: '8,500 MΩ (Yêu cầu ≥ 1,000 MΩ)',
        standard: 'Cấp điện áp 15-25kV BẮT BUỘC ≥ 1,000 MΩ',
        standardEn: '15-25kV rated windings MUST be ≥ 1,000 MΩ',
        reference: 'NETA ATS-2025 Table 100.5',
        status: 'PASS',
        detail: 'Cuộn cao thế VPI đạt cách điện xuất sắc, không phát hiện rò rỉ điện trường.',
        detailEn: 'Primary VPI windings exhibit superior insulation withstand.',
      },
      {
        item: 'Hệ số tổn hao cách điện Cuộn dây (Winding Power Factor Table 100.3)',
        itemEn: 'Winding Insulation Power Factor at 20°C',
        measured: '0.68%',
        standard: 'Máy biến áp khô công suất lớn BẮT BUỘC ≤ 2.0% tại 20°C',
        standardEn: 'Large dry-type transformers MUST be ≤ 2.0% at 20°C',
        reference: 'NETA ATS-2025 Table 100.3',
        status: 'PASS',
        detail: 'Hệ số tổn hao điện môi 0.68% khẳng định cuộn dây khô ráo và liên kết keo nhựa hoàn hảo.',
        detailEn: '0.68% dissipation factor confirms resin bonding void-free.',
      },
      {
        item: 'Điện trở cách điện Lõi Thép đến Vỏ (Core-to-Ground IR @ 500V DC)',
        itemEn: 'Core Insulation Resistance (Core-to-Ground Test)',
        measured: '420 MΩ',
        standard: 'Điện trở cách điện lõi thép BẮT BUỘC ≥ 100 MΩ',
        standardEn: 'Core insulation resistance MUST be ≥ 100 MΩ',
        reference: 'NETA ATS-2025 Sec 7.2.1.2.D.6',
        status: 'PASS',
        detail: 'Không có hiện tượng tiếp địa đa điểm (multiple ground loops) trong lõi từ.',
        detailEn: 'Core grounded at a single dedicated point; no stray ground loops.',
      },
      {
        item: 'Tỷ số biến áp Turns-Ratio trên tất cả các nấc Tap (IEEE C57.12.91)',
        itemEn: 'Turns Ratio Verification Across All Steps',
        measured: 'Sai số cực đại: 0.11%',
        standard: 'Độ lệch không vượt quá 0.5% so với tính toán thiết kế',
        standardEn: 'Deviation shall not exceed 0.5% compared to calculated ratio',
        reference: 'NETA ATS-2025 Sec 7.2.1.2',
        status: 'PASS',
        detail: 'Tỷ số phân áp giữa các pha A-B-C đồng nhất cao độ.',
        detailEn: 'All taps demonstrate uniform voltage transformation.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Hệ Số Tổn Hao Điện Môi & Chất Lượng Keo VPI',
        titleEn: 'Dielectric Dissipation Factor & Resin Integrity',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          'Hệ số tổn hao 0.68% thấp hơn rất nhiều so với trần 2.0% của Bảng Table 100.3.',
          'Không có bọt khí tồn đọng trong quá trình tẩm chân không áp lực VPI.',
        ],
        pointsEn: [
          'Dissipation factor of 0.68% well below Table 100.3 limit of 2.0%.',
          'Void-free resin impregnation confirmed.',
        ],
      },
      {
        titleVi: 'Cách Điện Lõi Thép & Tiếp Địa Một Điểm',
        titleEn: 'Core Grounding & Multi-Point Loop Prevention',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Core IR 420 MΩ ngăn chặn dòng xoáy tuần hoàn gây nóng cục bộ gông từ.',
        ],
        pointsEn: [
          '420 MΩ core IR prevents circulating eddy currents in frame clamps.',
        ],
      },
      {
        titleVi: 'Đối Lưu Quạt Cưỡng Bức & Cảm Biến RTD',
        titleEn: 'Forced Air Fans & Pt100 RTD Sensors',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          'Quạt làm mát AF (Air Forced) kích hoạt tự động ở 100°C và tắt ở 80°C.',
        ],
        pointsEn: [
          'Forced air cooling fans engage automatically at 100°C and reset at 80°C.',
        ],
      },
      {
        titleVi: 'Điện Trở Mối Nối Thanh Cái Đồng Hạ Thế',
        titleEn: 'LV Copper Busbar Bolted Joint Health',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Mối nối thanh cái 0.4kV chịu dòng 3600A đạt điện trở 6.8 µΩ (lệch 2.1%).',
        ],
        pointsEn: [
          '0.4kV 3600A busbar joints measure 6.8 µΩ (2.1% max deviation).',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện đóng điện nghiệm thu đưa vào vận hành.',
      'Đóng điện không tải 12 giờ để sấy khô tự nhiên bề mặt cuộn dây.',
      'Khảo sát nhiệt hồng ngoại tại các đầu cosse cáp sau khi mang tải 50%.',
    ],
    recommendationsEn: [
      'Approved for acceptance energization.',
      'Perform 12-hour no-load soak test to acclimatize windings.',
      'Perform infrared thermography survey on cable terminations at 50% load.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Trung Tâm Dữ Liệu TEV Hyperscale DC',
        transformer_tag: 'TR-DRY-2500KVA-01',
        fse_name: 'Đặng Tuấn Anh (anh.dt@tev.vn)',
        test_date: '2026-10-08',
      },
      electrical_tests: {
        insulation_resistance_pri_megohms: 8500,
        insulation_resistance_sec_megohms: 3400,
        winding_power_factor_percent: 0.68,
        core_insulation_resistance_megohms: 420,
        turns_ratio_max_error_percent: 0.11,
      },
    },
  },

  // 7. LOW-VOLTAGE CABLES <= 1000V (Sec 7.3.2)
  'neta-cable-lv': {
    id: 'neta-cable-lv',
    sectionCode: 'Mục 7.3.2',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.3.2',
    equipmentTitleVi: 'Cáp Điện Lực Hạ Áp ≤ 1000V (Low-Voltage Power Cables)',
    equipmentTitleEn: 'Low-Voltage Power Cables (≤ 1000V Rated)',
    subtitleVi: 'Thị giác • Cách điện 1000V DC (Table 100.1 ≥ 100 MΩ) • Mối nối DLRO (≤ 50%) • Đúng thứ tự pha',
    subtitleEn: 'Visual • Insulation 1000V DC (Table 100.1 ≥ 100 MΩ) • Joint DLRO (≤ 50%) • Phase Phasing',
    category: 'cables',
    defaultSiteInfo: {
      project_name: 'Tòa Nhà Văn Phòng TEV Tower',
      equipment_tag: 'CBL-LV-FEEDER-01',
      equipment_rating: '0.6/1kV 4x1C 240mm2 Cu/XLPE/PVC Feeder Cable',
      fse_name: 'Ngô Quốc Cường (cuong.nq@tev.vn)',
      test_date: '2026-10-07',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / ĐẠT TIÊU CHUẨN NGHIỆM THU NETA ATS-2025 MỤC 7.3.2',
    statusTitleEn: 'PASS / COMPLIANT WITH NETA ATS-2025 SECTION 7.3.2',
    statusReasonVi: 'Toàn bộ các pha cáp lực hạ áp đo tại 1000V DC trong 1 phút đạt điện trở cách điện từ 2,400 MΩ đến 2,800 MΩ (vượt xa ngưỡng 100 MΩ theo Table 100.1), điện trở đầu cosse bấm đạt 12.5 µΩ, và thứ tự pha đối soát chuẩn xác.',
    statusReasonEn: 'All phases tested at 1000V DC for 1 minute exhibit insulation resistance between 2,400 MΩ and 2,800 MΩ (far exceeding 100 MΩ Table 100.1 limit), lug resistance is 12.5 µΩ, and phase rotation verified.',
    evaluationItems: [
      {
        item: 'Điện trở cách điện Cáp Pha-Pha & Pha-Đất (1000V DC trong 1 phút)',
        itemEn: 'Insulation Resistance Phase-to-Phase & Phase-to-Ground',
        measured: 'A-G: 2,600 MΩ | B-G: 2,800 MΩ | C-G: 2,400 MΩ | N-G: 2,500 MΩ',
        standard: 'Cáp hạ thế cách điện XLPE BẮT BUỘC ≥ 100 MΩ (Table 100.1)',
        standardEn: 'LV XLPE cables MUST exceed ≥ 100 MΩ per Table 100.1',
        reference: 'NETA ATS-2025 Table 100.1 & Sec 7.3.2',
        status: 'PASS',
        detail: 'Lớp bọc cách điện XLPE khô ráo, không bị trầy xước trong quá trình kéo rải cáp.',
        detailEn: 'XLPE insulation is intact with no pulling mechanical abrasions.',
      },
      {
        item: 'Điện trở tiếp xúc đầu cosse bấm ép (Lug Bolted Connection DLRO)',
        itemEn: 'Lug Bolted Joint Resistance (DLRO Test)',
        measured: '12.5 µΩ (độ lệch giữa các pha max 4.0%)',
        standard: 'Độ lệch không vượt quá 50% so với giá trị nhỏ nhất của mối nối tương tự',
        standardEn: 'Deviation must not exceed 50% of lowest similar connection',
        reference: 'NETA ATS-2025 Table 100.12',
        status: 'PASS',
        detail: 'Đầu cosse đồng mạ thiếc bấm ép đủ số vết ép thủy lực, siết bu-lông chuẩn lực.',
        detailEn: 'Hydraulically crimped copper lugs torqued to Table 100.12 standards.',
      },
      {
        item: 'Kiểm tra thứ tự pha & Thông mạch liên tục (Phasing & Continuity)',
        itemEn: 'Phasing & Electrical Continuity Verification',
        measured: 'Pha A-B-C đúng thứ tự quay thuận (Clockwise Rotation)',
        standard: 'Đúng sơ đồ đơn tuyến, không chéo pha, thông mạch 100%',
        standardEn: 'Match one-line diagram, correct phase sequence, 100% continuity',
        reference: 'NETA ATS-2025 Sec 7.3.2.1',
        status: 'PASS',
        detail: 'Màu băng pha (Đỏ-Vàng-Xanh-Đen) đánh dấu chuẩn xác tại hai đầu tuyến cáp.',
        detailEn: 'Phase identification tags match on both source and destination terminations.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Chất Lượng Vỏ Bọc & Cách Điện XLPE',
        titleEn: 'XLPE Insulation & Jacket Integrity',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          'Điện trở cách điện 2,400 - 2,800 MΩ khẳng định vỏ cáp không bị ẩm xâm nhập qua hố ga.',
        ],
        pointsEn: [
          '2,400 to 2,800 MΩ IR proves zero moisture ingress along cable trench routing.',
        ],
      },
      {
        titleVi: 'Chất Lượng Ép Đầu Cosse & Lực Siết',
        titleEn: 'Crimped Lug & Bolt Torque Quality',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Điện trở 12.5 µΩ bảo đảm mang dòng 400A không bị phát nhiệt cục bộ.',
        ],
        pointsEn: [
          '12.5 µΩ resistance guarantees safe 400A ampacity without hot spots.',
        ],
      },
      {
        titleVi: 'Bán Kính Uốn Cong & Cố Định Thang Cáp',
        titleEn: 'Bending Radius & Cable Tray Clamping',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Bán kính uốn cong đạt > 8 lần đường kính ngoài của cáp per NEC/NFPA 70.',
        ],
        pointsEn: [
          'Bending radius complies with > 8x outer cable diameter per NEC.',
        ],
      },
      {
        titleVi: 'Chống Cháy Lan & Chặn Khe Hở (Firestop)',
        titleEn: 'Penetration Firestopping Seal',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          'Vị trí xuyên tường phòng điện được trám vữa chống cháy lan đạt tiêu chuẩn 2 giờ.',
        ],
        pointsEn: [
          'Wall penetration sealed with 2-hour rated firestop sealant.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện đóng điện mang tải.',
      'Dán nhãn số lộ và tên tải ở cả hai đầu tủ phân phối và tủ tiêu thụ.',
      'Chụp ảnh nhiệt kiểm tra mối nối sau 2 giờ mang tải toàn phần.',
    ],
    recommendationsEn: [
      'Approved for energization and full load delivery.',
      'Affix feeder circuit labels at both source and termination panels.',
      'Perform thermographic scan of cable lugs after 2 hours under load.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Tòa Nhà Văn Phòng TEV Tower',
        cable_tag: 'CBL-LV-FEEDER-01',
        fse_name: 'Ngô Quốc Cường (cuong.nq@tev.vn)',
        test_date: '2026-10-07',
      },
      electrical_tests: {
        insulation_resistance_phase_a_gnd_megohms: 2600,
        insulation_resistance_phase_b_gnd_megohms: 2800,
        insulation_resistance_phase_c_gnd_megohms: 2400,
        bolted_connection_micro_ohms: 12.5,
        phase_rotation_correct: true,
      },
    },
  },

  // 8. MEDIUM & HIGH-VOLTAGE CABLES (Sec 7.3.3)
  'neta-cable': {
    id: 'neta-cable',
    sectionCode: 'Mục 7.3.3',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.3.3 & IEEE 400.2',
    equipmentTitleVi: 'Cáp Điện Lực Trung & Cao Thế (Medium & High-Voltage Shielded Cables)',
    equipmentTitleEn: 'Medium & High-Voltage Shielded Power Cables',
    subtitleVi: 'Thị giác • Màn chắn đồng • Điện trở cách điện 5000V DC • Thử VLF AC Withstand (IEEE 400.2) • Tan Delta / VLF TD',
    subtitleEn: 'Visual • Copper Shield • IR 5000V DC • VLF AC Withstand (IEEE 400.2) • VLF Tan Delta / TD Stability',
    category: 'cables',
    defaultSiteInfo: {
      project_name: 'Nhà Máy Điện Gió TEV Wind Power Plant 100MW',
      equipment_tag: 'CBL-24KV-COLLECTOR-01',
      equipment_rating: '24kV 3x1C 300mm2 Cu/XLPE/CWS/HDPE Underground Cable (2.5 km)',
      fse_name: 'Nguyễn Tiến Dũng (dung.nt@tev.vn)',
      test_date: '2026-10-06',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / CÁP TRUNG THẾ ĐẠT CHUẨN NETA ATS-2025 & IEEE 400.2',
    statusTitleEn: 'PASS / MV CABLE COMPLIANT WITH NETA ATS-2025 & IEEE 400.2',
    statusReasonVi: 'Cáp 24kV vượt qua thử nghiệm chịu đựng điện áp cao xoay chiều tần số cực thấp VLF AC Withstand tại 28kV RMS trong 30 phút mà không phóng điện, điện trở cách điện đạt 45,000 MΩ (≥ 5,000 MΩ), và độ lệch Tan Delta TD Stability < 0.1 x 10^-3.',
    statusReasonEn: '24kV cable passed VLF AC Withstand at 28kV RMS for 30 minutes without breakdown, insulation resistance is 45,000 MΩ (≥ 5,000 MΩ), and Tan Delta stability < 0.1 x 10^-3.',
    evaluationItems: [
      {
        item: 'Thử nghiệm Chịu đựng Điện áp VLF AC Withstand (0.1 Hz per IEEE 400.2)',
        itemEn: 'VLF AC High-Potential Withstand Test (0.1 Hz Sine Wave)',
        measured: '28.0 kV RMS trong 30 phút (Không xảy ra phóng điện đánh thủng)',
        standard: 'Duy trì điện áp thử nghiệm liên tục trong 30-60 phút mà không có sự cố cách điện',
        standardEn: 'Withstand test voltage for 30-60 minutes without insulation flashover',
        reference: 'IEEE Std 400.2 & NETA ATS-2025 Table 100.6',
        status: 'PASS',
        detail: 'Cả 3 pha cáp đều chịu đựng thành công điện áp thử nghiệm nghiệm thu 28kV RMS.',
        detailEn: 'All three phases successfully withstood 28kV RMS acceptance stress.',
      },
      {
        item: 'Điện trở cách điện Cáp sau thử VLF (5000V DC trong 1 phút)',
        itemEn: 'Insulation Resistance Post-VLF Withstand (5000V DC)',
        measured: '45,000 MΩ (Table 100.1 quy định ≥ 5,000 MΩ)',
        standard: 'Cáp trung thế cấp 15-25kV BẮT BUỘC ≥ 5,000 MΩ',
        standardEn: '15-25kV shielded cable MUST be ≥ 5,000 MΩ per Table 100.1',
        reference: 'NETA ATS-2025 Table 100.1',
        status: 'PASS',
        detail: 'Điện trở cách điện cực cao chứng minh lớp XLPE không bị tổn thương suy thoái.',
        detailEn: 'High post-test IR proves XLPE tree-free condition.',
      },
      {
        item: 'Điện trở liên tục màn chắn kim loại đồng (Metallic Shield Continuity)',
        itemEn: 'Copper Wire Shield (CWS) Continuity & Resistance',
        measured: '0.42 Ω trên chiều dài 2.5 km tuyến cáp',
        standard: 'Điện trở màn chắn đồng đều giữa 3 pha, không bị đứt gãy màn chắn',
        standardEn: 'Shield resistance balanced across phases; no open shields',
        reference: 'NETA ATS-2025 Sec 7.3.3.1',
        status: 'PASS',
        detail: 'Màn chắn đồng thoát dòng dung và bảo vệ an toàn khi có ngắn mạch chạm đất.',
        detailEn: 'Copper tape/wire shield continuous to drain capacitive currents.',
      },
      {
        item: 'Đo Hệ số tổn hao điện môi VLF Tan Delta & Ổn định TD (IEEE 400.2)',
        itemEn: 'VLF Tan Delta (TD) Mean Value & TD Stability',
        measured: 'Mean TD: 1.2 x 10^-3 | TD Stability: 0.04 x 10^-3',
        standard: 'Cáp XLPE mới: Mean TD < 4.0 x 10^-3 và TD Stability < 0.1 x 10^-3',
        standardEn: 'New XLPE: Mean TD < 4.0 x 10^-3 and TD Stability < 0.1 x 10^-3',
        reference: 'IEEE 400.2 Table 5 & Table 6',
        status: 'PASS',
        detail: 'Phân loại tình trạng cách điện: Không có suy thoái bẫy nước (No Action Required).',
        detailEn: 'Condition assessment: Water tree free insulation (No Action Required).',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Khả Năng Chịu Đựng Điện Áp Cao VLF Withstand',
        titleEn: 'VLF AC Withstand Resilience',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          'Thử nghiệm 28kV RMS mô phỏng quá điện áp đóng cắt mà không gây tổn thương cáp.',
          'Dòng rò trong suốt 30 phút duy trì ổn định, không có hiện tượng tăng vọt.',
        ],
        pointsEn: [
          '28kV RMS simulated switching surges with zero insulation deterioration.',
          'Leakage current remained stable throughout the 30-minute test period.',
        ],
      },
      {
        titleVi: 'Đặc Tuyến Tan Delta & Suy Thoái Cây Nước (Water Trees)',
        titleEn: 'Tan Delta Diagnostics & Water Tree Assessment',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          'Độ ổn định TD 0.04 x 10^-3 khẳng định cách điện đồng nhất trên toàn bộ 2.5km.',
        ],
        pointsEn: [
          'TD stability of 0.04 x 10^-3 proves uniform insulation along 2.5km run.',
        ],
      },
      {
        titleVi: 'Đầu Cáp Co Nguội & Hộp Nối (Terminations & Splices)',
        titleEn: 'Cold-Shrink Terminations & In-Line Splices',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Ống điều hòa điện trường (Stress control tube) lắp đặt đúng kích thước gọt vỏ.',
          'Dây tiếp địa màn chắn bọc cao su chịu thời tiết lắp đặt chuẩn kỹ thuật.',
        ],
        pointsEn: [
          'Geometric stress control tubes positioned precisely per manufacturer template.',
          'Shield grounding braids sealed with weatherproof mastic.',
        ],
      },
      {
        titleVi: 'Tiếp Địa Màn Chắn Một Đầu / Hai Đầu (Shield Bonding)',
        titleEn: 'Shield Grounding Configuration',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Màn chắn được tiếp địa chắc chắn tại cả hai đầu xuất tuyến trạm biến áp.',
        ],
        pointsEn: [
          'Metallic shields solidly grounded at substation switchgear switch spouts.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện đóng điện hòa lưới tuyến cáp gom 24kV.',
      'Đậy kín nắp hố ga kỹ thuật và lắp đặt bơm chìm tự động chống ngập nước kéo dài.',
      'Lập hồ sơ lưu trữ chỉ số Tan Delta ban đầu để làm căn cứ so sánh kiểm tra định kỳ.',
    ],
    recommendationsEn: [
      'Approved for 24kV collector feeder grid connection.',
      'Seal all underground cable manholes and verify automated sump pump operation.',
      'Archive baseline VLF Tan Delta curves in CMMS for condition-based trending.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Nhà Máy Điện Gió TEV Wind Power Plant 100MW',
        cable_tag: 'CBL-24KV-COLLECTOR-01',
        fse_name: 'Nguyễn Tiến Dũng (dung.nt@tev.vn)',
        test_date: '2026-10-06',
      },
      electrical_tests: {
        vlf_withstand_test_kv_rms: 28.0,
        vlf_test_duration_minutes: 30,
        insulation_resistance_post_vlf_megohms: 45000,
        shield_continuity_ohms: 0.42,
        tan_delta_mean: 0.0012,
        tan_delta_stability: 0.00004,
      },
    },
  },

  // 9. AC INDUCTION MOTORS (Sec 7.15.1)
  'neta-motor': {
    id: 'neta-motor',
    sectionCode: 'Mục 7.15.1',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.15.1 & IEEE 43',
    equipmentTitleVi: 'Động Cơ Không Đồng Bộ AC (AC Induction Motors)',
    equipmentTitleEn: 'AC Induction Motors (Squirrel-Cage Rotor)',
    subtitleVi: 'Thị giác • Căn chỉnh trục • Điện trở Stator 20°C (Table 100.11) • Chỉ số PI (IEEE 43) • Điện trở một chiều (≤ 2%)',
    subtitleEn: 'Visual • Alignment • Stator IR 20°C (Table 100.11) • PI Index (IEEE 43) • Winding Resistance (≤ 2%)',
    category: 'motors',
    defaultSiteInfo: {
      project_name: 'Trạm Bơm Cấp Nước Công Nghiệp TEV Water Plant',
      equipment_tag: 'MTR-PUMP-6.6KV-01',
      equipment_rating: '1200 kW 6.6kV 1485 RPM TEFC Induction Motor',
      fse_name: 'Hoàng Văn Thắng (thang.hv@tev.vn)',
      test_date: '2026-10-05',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / ĐỘNG CƠ ĐẠT TIÊU CHUẨN NGHIỆM THU NETA ATS-2025',
    statusTitleEn: 'PASS / MOTOR FULLY COMPLIANT WITH NETA ATS-2025 & IEEE 43',
    statusReasonVi: 'Điện trở cách điện Stator tại 2500V DC quy đổi về 20°C đạt 4,200 MΩ (ngưỡng tối thiểu 100 MΩ theo Table 100.11), chỉ số phân cực PI = 2.45 (vượt ngưỡng yêu cầu PI ≥ 2.0 cho cách điện Class F per IEEE 43), và độ lệch điện trở một chiều giữa 3 pha chỉ 0.85% (≤ 2.0%).',
    statusReasonEn: 'Stator insulation resistance at 2500V DC corrected to 20°C is 4,200 MΩ (≥ 100 MΩ Table 100.11), Polarization Index PI = 2.45 (≥ 2.0 Class F per IEEE 43), and winding resistance phase deviation is 0.85% (≤ 2.0%).',
    evaluationItems: [
      {
        item: 'Điện trở cách điện Stator quy đổi 20°C (2500V DC per Table 100.11)',
        itemEn: 'Stator Winding Insulation Resistance @ 20°C',
        measured: '4,200 MΩ',
        standard: 'Động cơ điện áp 2300V - 7200V BẮT BUỘC ≥ 100 MΩ (Table 100.11)',
        standardEn: 'Motors rated 2300V - 7200V MUST be ≥ 100 MΩ per Table 100.11',
        reference: 'NETA ATS-2025 Table 100.11 & IEEE 43',
        status: 'PASS',
        detail: 'Cách điện cuộn dây Stator đạt gấp 42 lần tiêu chuẩn nghiệm thu.',
        detailEn: 'Stator winding insulation is 42 times above minimum requirement.',
      },
      {
        item: 'Chỉ số phân cực Polarization Index PI (R10min / R1min per IEEE 43)',
        itemEn: 'Polarization Index PI (10-min / 1-min Ratio per IEEE 43)',
        measured: 'PI = 2.45',
        standard: 'Hệ thống cách điện cấp Class F/H BẮT BUỘC PI ≥ 2.0',
        standardEn: 'Class F / Class H insulation systems MUST exhibit PI ≥ 2.0',
        reference: 'IEEE Std 43-2013 Table 3',
        status: 'PASS',
        detail: 'Đường cong dòng nạp điện môi dốc đứng, khẳng định cuộn dây khô sạch hoàn toàn.',
        detailEn: 'Dielectric absorption profile confirms pristine dry insulation.',
      },
      {
        item: 'Điện trở một chiều cuộn dây Stator (Winding DC Resistance Balance)',
        itemEn: 'Stator Winding DC Resistance Balance Across Phases',
        measured: 'U-V: 0.282 Ω | V-W: 0.284 Ω | W-U: 0.283 Ω (Lệch cực đại: 0.85%)',
        standard: 'Độ lệch điện trở giữa các pha BẮT BUỘC ≤ 2.0% so với giá trị trung bình',
        standardEn: 'Inter-phase resistance deviation MUST be ≤ 2.0% of the average',
        reference: 'NETA ATS-2025 Sec 7.15.1.D.2',
        status: 'PASS',
        detail: 'Cuộn dây 3 pha đối xứng hoàn hảo, các mối nối đầu bối dây hàn chắc chắn.',
        detailEn: 'Stator windings symmetrically balanced, ensuring even magnetic pull.',
      },
      {
        item: 'Kiểm tra căn chỉnh đồng tâm trục (Shaft Alignment & Air Gap)',
        itemEn: 'Shaft Laser Alignment & Air Gap Uniformity',
        measured: 'Độ lệch góc: 0.03 mm/100mm | Độ lệch tâm: 0.02 mm',
        standard: 'Độ lệch trục nằm trong dung sai cho phép của nhà sản xuất (≤ 0.05 mm)',
        standardEn: 'Shaft alignment within precision tolerances (≤ 0.05 mm)',
        reference: 'NETA ATS-2025 Sec 7.15.1.1',
        status: 'PASS',
        detail: 'Căn chỉnh laser khớp nối giữa động cơ và bơm đạt độ chính xác cao.',
        detailEn: 'Precision laser alignment to pump shaft verified.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Độ Bền Cách Điện Stator & Chỉ Số Phân Cực',
        titleEn: 'Stator Insulation Health & Polarization Index',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          'Chỉ số PI = 2.45 khẳng định không có bụi bẩn hay hơi ẩm thẩm thấu vào rãnh stato.',
        ],
        pointsEn: [
          'PI = 2.45 proves clean, dry, and void-free stator slot insulation.',
        ],
      },
      {
        titleVi: 'Cân Bằng Điện Trở Pha & Nguy Cơ Quá Nhiệt',
        titleEn: 'Phase Resistance Symmetry & Thermal Balance',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Độ lệch 0.85% loại trừ nguy cơ mất cân bằng dòng điện gây rung giật cơ khí.',
        ],
        pointsEn: [
          '0.85% deviation prevents current unbalance and motor vibration harmonics.',
        ],
      },
      {
        titleVi: 'Điện Trở Nhiệt Sấy Cuộn Dây (Space Heater)',
        titleEn: 'Anti-Condensation Space Heater Operation',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          'Điện trở sấy Space Heater 220V/500W hoạt động tốt khi động cơ dừng.',
        ],
        pointsEn: [
          'Anti-condensation space heaters engage automatically upon motor shutdown.',
        ],
      },
      {
        titleVi: 'Hệ Thống Bôi Trơn Vòng Bi (Bearing Lubrication)',
        titleEn: 'Bearing Lubrication & RTD Monitoring',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Vòng bi được tra mỡ tổng hợp chịu nhiệt Polyurea đúng dung tích nhà sản xuất.',
        ],
        pointsEn: [
          'Bearings lubricated with specified high-temperature polyurea grease.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện đóng điện chạy thử không tải (Solo Run 2-4 hours).',
      'Đo độ rung rung chấn cơ khí (Vibration Analysis per ISO 10816-3) trong quá trình chạy không tải.',
      'Kiểm tra chiều quay trục động cơ khớp với chiều mũi tên cánh bơm trước khi nối khớp nối.',
    ],
    recommendationsEn: [
      'Approved for uncoupled solo run (2 to 4 hours).',
      'Perform ISO 10816-3 mechanical vibration baseline survey during uncoupled run.',
      'Verify phase rotation matches pump impeller directional arrow before coupling.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Trạm Bơm Cấp Nước Công Nghiệp TEV Water Plant',
        motor_tag: 'MTR-PUMP-6.6KV-01',
        fse_name: 'Hoàng Văn Thắng (thang.hv@tev.vn)',
        test_date: '2026-10-05',
      },
      electrical_tests: {
        stator_insulation_20c_megohms: 4200,
        polarization_index_pi: 2.45,
        winding_resistance_phase_deviation_percent: 0.85,
        shaft_alignment_angular_mm: 0.03,
        space_heater_functional: true,
      },
    },
  },

  // 10. SYNCHRONOUS MOTORS & GENERATORS (Sec 7.15.2)
  'neta-sync-machinery': {
    id: 'neta-sync-machinery',
    sectionCode: 'Mục 7.15.2',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.15.2 & IEEE 43',
    equipmentTitleVi: 'Máy Điện Đồng Bộ (Synchronous Motors & Generators)',
    equipmentTitleEn: 'Synchronous Motors & Generators (Stator, Rotor & Exciter)',
    subtitleVi: 'Thị giác • Cách điện Stator (Table 100.11) • Cách điện Rotor (≥ 10 MΩ) • Mạch kích từ Exciter • Diode quay',
    subtitleEn: 'Visual • Stator IR (Table 100.11) • Rotor IR (≥ 10 MΩ) • Exciter System • Rotating Diode Bridge',
    category: 'motors',
    defaultSiteInfo: {
      project_name: 'Nhà Máy Nhiệt Điện Sinh Khối TEV Biomass Plant',
      equipment_tag: 'GEN-SYNC-15MW-01',
      equipment_rating: '15 MW 11kV 3000 RPM Salient-Pole Synchronous Generator',
      fse_name: 'Vũ Đình Cảnh (canh.vd@tev.vn)',
      test_date: '2026-10-04',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / MÁY ĐIỆN ĐỒNG BỘ ĐẠT CHUẨN NETA ATS-2025 MỤC 7.15.2',
    statusTitleEn: 'PASS / SYNCHRONOUS MACHINE COMPLIANT WITH NETA ATS-2025 SEC 7.15.2',
    statusReasonVi: 'Cách điện Stator đạt 12,000 MΩ (PI = 2.85), cách điện cuộn Rotor đạt 85 MΩ (vượt xa ngưỡng ≥ 10 MΩ), hệ thống kích từ không chổi than Brushless Exciter và cầu diode quay đạt kết quả thử nghiệm thuận-nghịch hoàn hảo.',
    statusReasonEn: 'Stator IR is 12,000 MΩ (PI = 2.85), Field Rotor IR is 85 MΩ (exceeding ≥ 10 MΩ threshold), brushless exciter and rotating diode bridge pass forward-reverse tests.',
    evaluationItems: [
      {
        item: 'Điện trở cách điện Stator quy đổi 20°C (5000V DC)',
        itemEn: 'Stator Winding Insulation Resistance @ 20°C',
        measured: '12,000 MΩ | PI = 2.85',
        standard: 'Máy điện 11kV BẮT BUỘC ≥ 100 MΩ; PI ≥ 2.0 (Class F)',
        standardEn: '11kV stator MUST be ≥ 100 MΩ; PI ≥ 2.0 per IEEE 43',
        reference: 'NETA ATS-2025 Table 100.11 & IEEE 43',
        status: 'PASS',
        detail: 'Cách điện cuộn dây Stator đạt chất lượng vượt trội.',
        detailEn: 'Stator insulation exhibits top-tier dielectric integrity.',
      },
      {
        item: 'Điện trở cách điện Cuộn kích từ Rotor (Field Winding IR @ 500V DC)',
        itemEn: 'Rotor Field Winding Insulation Resistance',
        measured: '85.0 MΩ',
        standard: 'Cuộn dây Rotor máy điện đồng bộ BẮT BUỘC ≥ 10 MΩ',
        standardEn: 'Synchronous rotor field winding MUST be ≥ 10 MΩ',
        reference: 'NETA ATS-2025 Sec 7.15.2.D.4',
        status: 'PASS',
        detail: 'Cuộn cực từ rotor cách điện an toàn với trục máy, không chạm đất kích từ.',
        detailEn: 'Rotor pole windings well isolated from rotor shaft forged body.',
      },
      {
        item: 'Kiểm tra Cầu Diode quay & Bảo vệ quá áp Varistor (Rotating Rectifier)',
        itemEn: 'Rotating Diode Rectifier Bridge & Varistor Check',
        measured: '6 Diode đo điện áp thuận 0.55V, điện trở nghịch vô cùng; Varistor bình thường',
        standard: 'Tỷ số thuận/nghịch diode rõ ràng, varistor không bị chập nổ',
        standardEn: 'Proper forward/reverse diode conduction ratio; intact surge suppressors',
        reference: 'NETA ATS-2025 Sec 7.15.2.1',
        status: 'PASS',
        detail: 'Cầu chỉnh lưu quay hoạt động tốt, đảm bảo cung cấp dòng kích từ DC ổn định.',
        detailEn: 'Rotating rectifier wheel functions properly with healthy surge protection.',
      },
      {
        item: 'Điện trở cách điện Gối trục (Insulated Bearing Pedestal IR)',
        itemEn: 'Shaft Bearing Pedestal Insulation Resistance',
        measured: '2.8 MΩ (Yêu cầu BẮT BUỘC ≥ 1.0 MΩ)',
        standard: 'Gối trục cách điện BẮT BUỘC ≥ 1.0 MΩ để ngăn dòng điện trục (Shaft Current)',
        standardEn: 'Insulated bearing pedestal MUST be ≥ 1.0 MΩ to mitigate shaft currents',
        reference: 'NETA ATS-2025 Sec 7.15.2.D.8',
        status: 'PASS',
        detail: 'Cách điện gối đỡ phía không tải NDE loại trừ nguy cơ rỗ mòn điện vòng bi.',
        detailEn: 'NDE bearing pedestal insulation prevents electrical fluting damage.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Cách Điện Stator & Độ Bền Điện Môi Lõi Từ',
        titleEn: 'Stator Insulation & Core Dielectric Integrity',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          'Cách điện 12,000 MΩ và PI = 2.85 bảo vệ an toàn cho máy phát điện 11kV.',
        ],
        pointsEn: [
          '12,000 MΩ and PI = 2.85 guarantee robust 11kV generator dielectric safety.',
        ],
      },
      {
        titleVi: 'Cuộn Kích Từ Rotor & Cầu Diode Quay',
        titleEn: 'Rotor Field Windings & Rotating Rectifiers',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          '85 MΩ Rotor IR loại trừ nguy cơ ngắn mạch vòng dây kích từ gây rung cơ khí.',
        ],
        pointsEn: [
          '85 MΩ field IR prevents rotor turn-to-turn shorts and thermal unbalance.',
        ],
      },
      {
        titleVi: 'Cách Điện Gối Trục Ngăn Dòng Điện Trục EDM',
        titleEn: 'Bearing Pedestal Insulation & EDM Mitigation',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Cách điện gối đỡ 2.8 MΩ triệt tiêu dòng rò cảm ứng gây rỗ bề mặt con lăn.',
        ],
        pointsEn: [
          '2.8 MΩ pedestal insulation blocks capacitive shaft circulating currents.',
        ],
      },
      {
        titleVi: 'Hệ Thống Điều Khiển Kích Từ Tự Động (AVR)',
        titleEn: 'Automatic Voltage Regulator (AVR) Response',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          'Thử nghiệm bước nhảy điện áp AVR phản hồi nhanh trong vòng 0.2s.',
        ],
        pointsEn: [
          'Digital AVR transient step response stabilizes output within 0.2s.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện quay trục máy phát đạt tốc độ định mức 3000 RPM.',
      'Thử nghiệm đặc tính không tải và ngắn mạch máy phát trước khi hòa lưới.',
      'Kiểm tra tiếp xúc chổi than nối đất trục (Shaft Grounding Brush) thoát điện tích tĩnh.',
    ],
    recommendationsEn: [
      'Approved for mechanical roll-up to rated 3000 RPM.',
      'Perform open-circuit saturation and sustained short-circuit testing prior to synchronization.',
      'Inspect shaft grounding brush contact to ensure static charge dissipation.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Nhà Máy Nhiệt Điện Sinh Khối TEV Biomass Plant',
        generator_tag: 'GEN-SYNC-15MW-01',
        fse_name: 'Vũ Đình Cảnh (canh.vd@tev.vn)',
        test_date: '2026-10-04',
      },
      electrical_tests: {
        stator_insulation_20c_megohms: 12000,
        stator_pi_ratio: 2.85,
        rotor_field_insulation_megohms: 85.0,
        bearing_insulation_megohms: 2.8,
        rotating_diodes_healthy: true,
      },
    },
  },

  // 11. DC MOTORS & GENERATORS (Sec 7.15.3)
  'neta-dc-motor': {
    id: 'neta-dc-motor',
    sectionCode: 'Mục 7.15.3',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.15.3 & Table 100.11',
    equipmentTitleVi: 'Động Cơ & Máy Phát Điện Một Chiều DC (DC Motors & Generators)',
    equipmentTitleEn: 'DC Motors & Generators (Armature, Commutator & Field)',
    subtitleVi: 'Thị giác • Cổ góp Commutator • Chổi than • Cách điện Armature (Table 100.11) • Cuộn Shunt & Series Field',
    subtitleEn: 'Visual • Commutator Film • Carbon Brushes • Armature IR (Table 100.11) • Shunt & Series Field IR',
    category: 'motors',
    defaultSiteInfo: {
      project_name: 'Nhà Máy Cán Thép TEV Steel Mill',
      equipment_tag: 'MTR-DC-ROLLING-01',
      equipment_rating: '800 kW 750V DC 600 RPM Shunt-Wound DC Motor',
      fse_name: 'Trịnh Công Sơn (son.tc@tev.vn)',
      test_date: '2026-10-03',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / ĐỘNG CƠ DC ĐẠT TIÊU CHUẨN NGHIỆM THU NETA ATS-2025',
    statusTitleEn: 'PASS / DC MOTOR FULLY COMPLIANT WITH NETA ATS-2025 SEC 7.15.3',
    statusReasonVi: 'Điện trở cách điện cuộn phần ứng Armature đạt 650 MΩ (≥ 100 MΩ), cuộn kích từ Shunt Field đạt 820 MΩ, cổ góp đồng bóng đều màu hạt dẻ, và chổi than tiếp xúc đạt > 85% diện tích bề mặt.',
    statusReasonEn: 'Armature insulation resistance is 650 MΩ (≥ 100 MΩ), Shunt Field IR is 820 MΩ, commutator exhibits chocolate-brown patina, and brush contact exceeds 85% area.',
    evaluationItems: [
      {
        item: 'Điện trở cách điện Cuộn phần ứng Armature (1000V DC per Table 100.11)',
        itemEn: 'Armature Winding Insulation Resistance @ 1000V DC',
        measured: '650 MΩ (Table 100.11 quy định tối thiểu ≥ 100 MΩ)',
        standard: 'Phần ứng DC cấp điện áp 600V - 1000V BẮT BUỘC ≥ 100 MΩ',
        standardEn: 'DC armatures rated 600V - 1000V MUST be ≥ 100 MΩ',
        reference: 'NETA ATS-2025 Table 100.11',
        status: 'PASS',
        detail: 'Rãnh cổ góp sạch bụi than, không có cầu nối carbon giữa các phiến góp.',
        detailEn: 'Mica undercuts clean of carbon dust; high insulation barrier verified.',
      },
      {
        item: 'Điện trở cách điện Cuộn kích từ song song Shunt Field & Cuộn phụ',
        itemEn: 'Shunt Field & Interpole Winding Insulation Resistance',
        measured: '820 MΩ (1000V DC)',
        standard: 'Cuộn kích từ máy điện DC BẮT BUỘC ≥ 100 MΩ',
        standardEn: 'DC field windings MUST exceed ≥ 100 MΩ',
        reference: 'NETA ATS-2025 Sec 7.15.3.D.2',
        status: 'PASS',
        detail: 'Cuộn kích từ cách điện tốt, không bị lão hóa do nhiệt.',
        detailEn: 'Shunt field and interpoles exhibit robust insulation withstand.',
      },
      {
        item: 'Kiểm tra Cổ góp Commutator & Rãnh Mica (Commutator & Undercut)',
        itemEn: 'Commutator Surface & Mica Undercut Depth Inspection',
        measured: 'Rãnh mica xẻ sâu 1.2 mm, vát mép phiến góp 45° sạch sẽ',
        standard: 'Mica xẻ rãnh sâu 0.8 - 1.5 mm, không có ba via đồng gờ cạnh',
        standardEn: 'Mica undercut 0.8 - 1.5 mm; copper bars chamfered without burrs',
        reference: 'NETA ATS-2025 Sec 7.15.3.1',
        status: 'PASS',
        detail: 'Bề mặt cổ góp láng mịn, sẵn sàng cho quá trình tạo màng oxit đồng bảo vệ.',
        detailEn: 'Commutator surface smooth; ready for copper oxide patina formation.',
      },
      {
        item: 'Kiểm tra Chổi than, Giá đỡ chổi & Áp lực lò xo (Carbon Brushes)',
        itemEn: 'Carbon Brush Seating & Spring Pressure Verification',
        measured: 'Áp lực lò xo: 18 kPa (Chuẩn 16-20 kPa) | Tiếp xúc chổi: > 85%',
        standard: 'Áp lực phân bổ đều, độ tiếp xúc bề mặt chổi than đạt tối thiểu 75%',
        standardEn: 'Even spring tension; minimum 75% brush face bedding contact',
        reference: 'NETA ATS-2025 Sec 7.15.3.1',
        status: 'PASS',
        detail: 'Chổi than mài rà khớp khít với độ cong cổ góp, dây bện pigtail dẫn điện tốt.',
        detailEn: 'Brushes properly seated to curvature; copper pigtails secured.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Bề Mặt Cổ Góp & Màng Oxit Bảo Vệ',
        titleEn: 'Commutator Surface & Protective Patina',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Mica rãnh sâu 1.2 mm ngăn ngừa hiện tượng đánh lửa vòng tròn (Flashover).',
        ],
        pointsEn: [
          '1.2 mm mica undercut eliminates circumferential flashover arcing.',
        ],
      },
      {
        titleVi: 'Chổi Than Carbon & Cân Bằng Áp Lực Lò Xo',
        titleEn: 'Brush Bedding & Spring Tension Balance',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Áp lực 18 kPa giữ chổi than không bị rung nảy khi động cơ quay tốc độ cao.',
        ],
        pointsEn: [
          '18 kPa spring force prevents brush bounce and micro-sparking under load.',
        ],
      },
      {
        titleVi: 'Cách Điện Phần Ứng & Cuộn Phụ Interpole',
        titleEn: 'Armature & Interpole Insulation Health',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          'Cách điện 650 MΩ bảo đảm vận hành an toàn với bộ biến đổi nguồn Thyristor DC Drive.',
        ],
        pointsEn: [
          '650 MΩ armature IR safely handles Thyristor DC Drive dV/dt pulses.',
        ],
      },
      {
        titleVi: 'Hệ Thống Thổi Gió Cưỡng Bức & Lọc Bụi',
        titleEn: 'Forced Blower & Air Filter Maintenance',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          'Quạt thổi cưỡng bức tạo áp suất dương ngăn bụi than thâm nhập cuộn kích từ.',
        ],
        pointsEn: [
          'Positive pressure blower prevents carbon dust settling on field coils.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện cấp nguồn chạy thử nghiệm không tải ở tốc độ thấp 100 RPM.',
      'Quan sát tia lửa chổi than (Sparking Grade per IEC 60034-3: yêu cầu đạt Grade 1 - không có tia lửa).',
      'Thổi sạch bụi than cổ góp bằng khí nén khô áp lực thấp sau 24 giờ vận hành đầu tiên.',
    ],
    recommendationsEn: [
      'Approved for low-speed 100 RPM no-load commissioning run.',
      'Monitor brush sparking grade per IEC 60034-3 (must meet Grade 1 sparkless).',
      'Blow down carbon dust with low-pressure dry air after first 24 operating hours.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Nhà Máy Cán Thép TEV Steel Mill',
        motor_tag: 'MTR-DC-ROLLING-01',
        fse_name: 'Trịnh Công Sơn (son.tc@tev.vn)',
        test_date: '2026-10-03',
      },
      electrical_tests: {
        armature_insulation_megohms: 650,
        shunt_field_insulation_megohms: 820,
        brush_spring_pressure_kpa: 18.0,
        brush_bedding_contact_percent: 88,
      },
    },
  },

  // 12. EMERGENCY ENGINE GENERATORS (Sec 7.22.1 & NFPA 110)
  'neta-generator': {
    id: 'neta-generator',
    sectionCode: 'Mục 7.22.1',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.22.1 & NFPA 110',
    equipmentTitleVi: 'Tổ Máy Phát Điện Khẩn Cấp / Dự Phòng (Emergency Engine Generators)',
    equipmentTitleEn: 'Emergency Engine Generator Sets (Diesel Genset & NFPA 110)',
    subtitleVi: 'Thị giác • Bảo vệ Tắt máy Quá tốc (115%) • Cách điện Stator IEEE 43 (PI ≥ 2.0) • Cấp tải NFPA 110 Class 10 (≤ 10s)',
    subtitleEn: 'Visual • Overspeed Trip (115%) • Stator IR IEEE 43 (PI ≥ 2.0) • NFPA 110 Class 10 Load Step (≤ 10s)',
    category: 'motors',
    defaultSiteInfo: {
      project_name: 'Bệnh Viện Đa Khoa Quốc Tế TEV Hospital',
      equipment_tag: 'GEN-DIESEL-2000KVA-01',
      equipment_rating: '2000 kVA 400/230V 1500 RPM Cummins Diesel Generator',
      fse_name: 'Nguyễn Văn Minh (minh.nv@tev.vn)',
      test_date: '2026-10-02',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / TỔ MÁY PHÁT ĐẠT CHUẨN NETA ATS-2025 & NFPA 110',
    statusTitleEn: 'PASS / GENSET COMPLIANT WITH NETA ATS-2025 & NFPA 110 CLASS 10',
    statusReasonVi: 'Thời gian khởi động và nhận 100% tải khẩn cấp chỉ 8.2 giây (đáp ứng tiêu chuẩn NFPA 110 Class 10 ≤ 10s), bảo vệ ngắt quá tốc độ tác động chính xác ở 114.5% (ngưỡng 115%), cách điện đầu phát đạt 2,800 MΩ (PI = 2.4), và tần số ổn định trong dung sai ±0.5%.',
    statusReasonEn: 'Emergency startup and load acceptance completed in 8.2 seconds (NFPA 110 Class 10 ≤ 10s), overspeed trip activated at 114.5% (115% limit), alternator IR is 2,800 MΩ (PI = 2.4), and frequency deviation is within ±0.5%.',
    evaluationItems: [
      {
        item: 'Thời gian khởi động & Nhận tải khẩn cấp (NFPA 110 Class 10 Test)',
        itemEn: 'Emergency Cold Start & Full Load Acceptance (NFPA 110)',
        measured: '8.2 giây (Đạt điện áp định mức 400V và đóng máy cắt cấp tải)',
        standard: 'Hệ thống nguồn khẩn cấp Loại 1 BẮT BUỘC nhận tải trong vòng ≤ 10 giây',
        standardEn: 'Emergency Level 1 EPS MUST accept load within ≤ 10 seconds',
        reference: 'NFPA 110 Sec 7.13 & NETA ATS Sec 7.22.1',
        status: 'PASS',
        detail: 'Động cơ khởi động nhanh chóng, máy cắt cấp tải đóng hoàn tất ở giây thứ 8.2.',
        detailEn: 'Fast diesel cranking, generator breaker closes at 8.2s mark.',
      },
      {
        item: 'Bảo vệ Tắt máy An toàn: Cắt Quá Tốc Độ (Overspeed Trip Test)',
        itemEn: 'Safety Shutdown Functions: Engine Overspeed Trip',
        measured: 'Tác động ngắt ở 1718 RPM (114.5% tốc độ định mức 1500 RPM)',
        standard: 'Bảo vệ quá tốc độ BẮT BUỘC cắt nhiên liệu ngắt tức thì tại ≤ 115%',
        standardEn: 'Overspeed shutdown MUST trip instantly at ≤ 115% rated RPM',
        reference: 'NETA ATS-2025 Sec 7.22.1.D.3',
        status: 'PASS',
        detail: 'Van ngắt gió khẩn cấp và rơ-le nhiên liệu cắt dứt khoát bảo vệ an toàn động cơ.',
        detailEn: 'Electronic ECM and air shutoff valve trip cleanly at 114.5%.',
      },
      {
        item: 'Điện trở cách điện Đầu phát Stator (1000V DC per IEEE 43)',
        itemEn: 'Alternator Stator Insulation Resistance @ 1000V DC',
        measured: '2,800 MΩ | PI = 2.40',
        standard: 'Đầu phát hạ thế BẮT BUỘC ≥ 100 MΩ; PI ≥ 2.0 (Class H)',
        standardEn: 'LV alternator MUST be ≥ 100 MΩ; PI ≥ 2.0 per IEEE 43',
        reference: 'NETA ATS-2025 Table 100.11 & IEEE 43',
        status: 'PASS',
        detail: 'Đầu phát điện khô ráo, không bị ẩm ngưng tụ từ khí thải.',
        detailEn: 'Alternator insulation exceeds minimum requirements.',
      },
      {
        item: 'Thử nghiệm Tải giả 100% Load Bank trong 2 giờ (Full-Load Test)',
        itemEn: 'Full-Load Load Bank Endurance Run (100% for 2 Hours)',
        measured: 'Tải 1600 kW (100%): Nhiệt độ nước 88°C, Áp suất dầu 4.8 bar, Điện áp 400V ±0.8%',
        standard: 'Vận hành ổn định đầy tải mà không báo lỗi quá nhiệt hoặc sụt áp',
        standardEn: 'Continuous full-load stability with steady coolant temp and oil pressure',
        reference: 'NETA ATS-2025 Sec 7.22.1.D.4',
        status: 'PASS',
        detail: 'Hệ thống làm mát két nước tản nhiệt và quạt giải nhiệt vận hành tối ưu.',
        detailEn: 'Radiator cooling and oil lubrication perform within normal operating bands.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Khả Năng Cấp Nguồn Khẩn Cấp NFPA 110 Class 10',
        titleEn: 'NFPA 110 Class 10 Emergency Response',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Thời gian 8.2s bảo đảm các phòng mổ và thiết bị y tế không bị gián đoạn điện.',
        ],
        pointsEn: [
          '8.2s response guarantees uninterrupted emergency power to hospital operating theaters.',
        ],
      },
      {
        titleVi: 'Bảo Vệ Động Cơ Diesel & Tắt Máy Khẩn Cấp',
        titleEn: 'Diesel Engine Safety Shutdown Functions',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          'Tất cả 4 cảm biến an toàn (Quá tốc, Áp suất dầu thấp, Nhiệt độ nước cao, Mức nước thấp) hoạt động 100%.',
        ],
        pointsEn: [
          'All 4 safety interlocks (Overspeed, Low Oil Pressure, High Coolant Temp, Low Water) functional.',
        ],
      },
      {
        titleVi: 'Chất Lượng Điện Áp & Độ Ổn Định Tần Số AVR/Governor',
        titleEn: 'Voltage & Frequency Regulation Dynamics',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Độ sụt tần số khi nhận bước tải 50% chỉ 1.8 Hz và hồi phục trong vòng 1.2 giây.',
        ],
        pointsEn: [
          'Frequency transient dip during 50% block load is only 1.8 Hz, recovering within 1.2s.',
        ],
      },
      {
        titleVi: 'Hệ Thống Ắc Quy Khởi Động & Bộ Sạc Tự Động',
        titleEn: 'Dual Starting Battery Banks & Float Charger',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Ắc quy đề 24V đạt dòng phóng khởi động CCA 1200A, bộ sạc duy trì 27.6V ổn định.',
        ],
        pointsEn: [
          '24V starter battery delivers 1200A CCA; float charger maintains steady 27.6V.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện nghiệm thu bàn giao hệ thống nguồn khẩn cấp tòa nhà.',
      'Duy trì chế độ tự động chạy không tải định kỳ hàng tuần 15 phút (Weekly Exercise).',
      'Đổ đầy bình dầu diesel dự phòng duy trì tối thiểu 24 giờ vận hành liên tục (NFPA 110).',
    ],
    recommendationsEn: [
      'Approved for hospital Level 1 emergency power commission handover.',
      'Configure automated weekly 15-minute uncoupled exercise test.',
      'Maintain day tank and bulk fuel storage at full capacity for 24-hour run time.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Bệnh Viện Đa Khoa Quốc Tế TEV Hospital',
        generator_tag: 'GEN-DIESEL-2000KVA-01',
        fse_name: 'Nguyễn Văn Minh (minh.nv@tev.vn)',
        test_date: '2026-10-02',
      },
      electrical_tests: {
        startup_to_load_seconds: 8.2,
        overspeed_trip_rpm: 1718,
        alternator_insulation_megohms: 2800,
        load_bank_test_kw: 1600,
        voltage_regulation_percent: 0.8,
      },
    },
  },

  // 13. AUTOMATIC TRANSFER SWITCHES ATS (Sec 7.22.3)
  'neta-ats': {
    id: 'neta-ats',
    sectionCode: 'Mục 7.22.3',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.22.3',
    equipmentTitleVi: 'Bộ Chuyển Nguồn Tự Động (Automatic Transfer Switches - ATS)',
    equipmentTitleEn: 'Automatic Transfer Switches (ATS Transition & Interlocks)',
    subtitleVi: 'Khóa Liên Động Cơ Khí • Điện Trở Tiếp Điểm (≤ 50%) • Cách Điện Điều Khiển (≥ 2 MΩ) • Trình Tự Tự Động',
    subtitleEn: 'Mechanical Interlocks • Contact DLRO (≤ 50%) • Control Insulation (≥ 2 MΩ) • Auto Sequence Simulation',
    category: 'switches',
    defaultSiteInfo: {
      project_name: 'Trung Tâm Dữ Liệu Ngân Hàng TEV Bank Data Center',
      equipment_tag: 'ATS-400V-1600A-01',
      equipment_rating: '1600A 400V 4-Pole Closed Transition ATS with Asco Controller',
      fse_name: 'Đỗ Hữu Trí (tri.dh@tev.vn)',
      test_date: '2026-10-01',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / BỘ CHUYỂN NGUỒN ATS ĐẠT CHUẨN NETA ATS-2025 MỤC 7.22.3',
    statusTitleEn: 'PASS / ATS FULLY COMPLIANT WITH NETA ATS-2025 SECTION 7.22.3',
    statusReasonVi: 'Khóa liên động cơ điện ngăn chặn đóng đồng thời hai nguồn không đồng bộ hoạt động tuyệt đối an toàn, điện trở tiếp điểm cực chính đo tại 100A DC đạt 16.2 µΩ (lệch cực đại 2.8% ≤ 50%), cách điện mạch điều khiển đạt 38 MΩ (≥ 2.0 MΩ), và chu trình mô phỏng mất điện lưới chuyển nguồn máy phát trơn tru trong 3.2 giây.',
    statusReasonEn: 'Mechanical interlocks prevent asynchronous cross-connection, contact resistance is 16.2 µΩ (2.8% deviation ≤ 50%), control circuit IR is 38 MΩ (≥ 2.0 MΩ), and simulated outage transfer operates seamlessly in 3.2s.',
    evaluationItems: [
      {
        item: 'Kiểm tra Khóa liên động Cơ khí & Điện khí (Mechanical & Electrical Interlocks)',
        itemEn: 'Mechanical & Electrical Interlock Safety Verification',
        measured: 'Khóa liên động cơ khí chặn 100% không cho phép 2 nguồn đóng đồng thời',
        standard: 'Tuyệt đối ngăn chặn đóng nhầm gây đoản mạch giữa nguồn Lưới và Máy phát',
        standardEn: 'Strictly prohibit paralleling unsynchronized sources',
        reference: 'NETA ATS-2025 Sec 7.22.3.1',
        status: 'PASS',
        detail: 'Cơ cấu chốt cơ khí và công tắc phụ liên động điện hoạt động tin cậy.',
        detailEn: 'Heavy-duty mechanical walking beam interlock operates without jamming.',
      },
      {
        item: 'Điện trở tiếp xúc cực đóng mạch (Contact Resistance DLRO @ 100A DC)',
        itemEn: 'Main Power Contact Resistance (DLRO Test)',
        measured: 'Nguồn Lưới: 16.2 µΩ | Nguồn Máy phát: 16.5 µΩ (Lệch cực đại 2.8%)',
        standard: 'Độ lệch không vượt quá 50% so với giá trị nhỏ nhất của các cực',
        standardEn: 'Deviation shall not exceed 50% of lowest pole reading',
        reference: 'NETA ATS-2025 Table 100.12',
        status: 'PASS',
        detail: 'Bề mặt tiếp điểm mạ bạc phẳng khít, lực ép lò xo tiếp điểm cân bằng.',
        detailEn: 'Silver-alloy contacts maintain tight contact under high spring tension.',
      },
      {
        item: 'Điện trở cách điện Cực Động lực & Mạch điều khiển (Table 100.1)',
        itemEn: 'Power Pole & Control Circuit Insulation Resistance',
        measured: 'Cực động lực: 2,100 MΩ (1000V DC) | Mạch điều khiển: 38 MΩ (500V DC)',
        standard: 'Cực động lực BẮT BUỘC ≥ 100 MΩ; Mạch điều khiển BẮT BUỘC ≥ 2.0 MΩ',
        standardEn: 'Power poles MUST be ≥ 100 MΩ; Control circuit MUST be ≥ 2.0 MΩ',
        reference: 'NETA ATS-2025 Table 100.1 & Sec 7.22.3.D.3',
        status: 'PASS',
        detail: 'Cách điện an toàn giữa các pha, pha-đất và mạch cấp tín hiệu khởi động máy phát.',
        detailEn: 'Insulation margins safely isolate control board logic from 400V power poles.',
      },
      {
        item: 'Mô phỏng Mất điện Lưới & Trình tự Tự động Chuyển nguồn (Auto Sequence)',
        itemEn: 'Loss of Normal Power Simulation & Auto Transfer Timing',
        measured: 'Mất điện lưới -> Phát lệnh đề máy phát: 1.0s -> Chuyển nguồn ATS: 3.2s',
        standard: 'Đúng chu trình thời gian cài đặt bộ điều khiển (Timers TDES, TDNE, TDEN)',
        standardEn: 'Matches programmed timing sequences per controller schedule',
        reference: 'NETA ATS-2025 Sec 7.22.3.D.4',
        status: 'PASS',
        detail: 'Trình tự chuyển đổi từ Lưới sang Máy phát và tự động hoàn trả khi Lưới phục hồi hoàn hảo.',
        detailEn: 'Seamless transfer sequence and retransfer upon utility restoration.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Khóa Liên Động Cơ Khí Chống Chập Nguồn',
        titleEn: 'Mechanical Walking-Beam Interlock Integrity',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Thanh liên động cơ khí ngăn ngừa nguy cơ cấp ngược điện từ máy phát về lưới trung thế.',
        ],
        pointsEn: [
          'Interlock mechanism prevents hazardous backfeeding into de-energized utility grid.',
        ],
      },
      {
        titleVi: 'Điện Trở Tiếp Điểm & Khả Năng Chịu Dòng Ngắn Mạch',
        titleEn: 'Contact Resistance & Withstand Rating (WCR)',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Tiếp điểm 16.2 µΩ bảo đảm độ bền cơ nhiệt khi có dòng ngắn mạch 65kA chạy qua.',
        ],
        pointsEn: [
          '16.2 µΩ contact resistance ensures 65kA Withstand and Closing Rating (WCR).',
        ],
      },
      {
        titleVi: 'Độ Nhạy Cảm Biến Điện Áp & Tần Số Bộ Điều Khiển',
        titleEn: 'Microprocessor Controller Voltage & Frequency Sensing',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Bộ điều khiển nhận biết sụt áp pha < 85% hoặc mất pha ngay lập tức trong 0.1s.',
        ],
        pointsEn: [
          'Controller senses phase under-voltage < 85% or phase loss within 0.1s.',
        ],
      },
      {
        titleVi: 'Thời Gian Làm Nguội Động Cơ Khi Hoàn Trả (Engine Cooldown)',
        titleEn: 'Engine Cooldown Timing Sequence',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          'Bộ định thời duy trì chạy không tải làm mát máy phát 5 phút trước khi dừng hẳn.',
        ],
        pointsEn: [
          'Engine cooldown timer programmed for 5 minutes to prevent turbocharger coking.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện đóng điện vận hành chuyển nguồn tự động.',
      'Kích hoạt tính năng tự động tập dượt chuyển nguồn có tải định kỳ hàng tháng (Monthly Load Test).',
      'Kiểm tra định kỳ siết bu-lông cáp động lực vào thanh đồng ATS sau 6 tháng.',
    ],
    recommendationsEn: [
      'Approved for live automatic standby service.',
      'Enable programmed monthly loaded transfer exercise cycle.',
      'Schedule 6-month infrared and torque verification of incoming/outgoing power lugs.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Trung Tâm Dữ Liệu Ngân Hàng TEV Bank Data Center',
        ats_tag: 'ATS-400V-1600A-01',
        fse_name: 'Đỗ Hữu Trí (tri.dh@tev.vn)',
        test_date: '2026-10-01',
      },
      electrical_tests: {
        normal_contact_resistance_micro_ohms: 16.2,
        emergency_contact_resistance_micro_ohms: 16.5,
        power_pole_insulation_megohms: 2100,
        control_circuit_insulation_megohms: 38.0,
        transfer_time_seconds: 3.2,
      },
    },
  },

  // 14. VRLA BATTERY SYSTEMS (Sec 7.18.2 & IEEE 1188)
  'neta-battery-vrla': {
    id: 'neta-battery-vrla',
    sectionCode: 'Mục 7.18.2',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.18.2 & IEEE 1188',
    equipmentTitleVi: 'Hệ Thống Ắc Quy Kín Khí VRLA (Valve-Regulated Lead-Acid Batteries)',
    equipmentTitleEn: 'VRLA Battery Systems (Stationary Valve-Regulated Lead-Acid)',
    subtitleVi: 'Thị giác • Nội trở Cell (IEEE 1188) • Điện áp Phao Float (±0.05V) • Điện trở cầu nối Cút (≤ 50%) • Nhiệt độ',
    subtitleEn: 'Visual • Internal Ohmic Resistance (IEEE 1188) • Float Voltage (±0.05V) • Intercell Connector DLRO • Thermal Scan',
    category: 'special',
    defaultSiteInfo: {
      project_name: 'Phòng Nguồn UPS Trung Tâm Dữ Liệu TEV Data Center',
      equipment_tag: 'BAT-VRLA-240V-STRING-01',
      equipment_rating: '240V DC String (40x 12V 100Ah AGM VRLA Batteries)',
      fse_name: 'Lương Thế Vinh (vinh.lt@tev.vn)',
      test_date: '2026-09-30',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / HỆ THỐNG ẮC QUY VRLA ĐẠT CHUẨN NETA ATS-2025 & IEEE 1188',
    statusTitleEn: 'PASS / VRLA BATTERY SYSTEM COMPLIANT WITH NETA ATS-2025 & IEEE 1188',
    statusReasonVi: 'Nội trở các bình ắc quy đo bằng thiết bị đo chuyên dụng đạt giá trị trung bình 4.25 mΩ (độ lệch giữa các cell cực đại 6.2% nằm an toàn dưới ngưỡng cảnh báo 20% theo IEEE 1188), điện áp phao đồng đều 13.62V ±0.03V/bình, và điện trở cầu nối intercell đạt 18.5 µΩ (lệch 4.1% ≤ 50%).',
    statusReasonEn: 'Internal ohmic resistance averages 4.25 mΩ (max cell variation 6.2% is well below 20% IEEE 1188 threshold), float voltage is balanced at 13.62V ±0.03V/unit, and intercell resistance is 18.5 µΩ (4.1% variation ≤ 50%).',
    evaluationItems: [
      {
        item: 'Nội trở bình ắc quy Internal Ohmic Value (IEEE Std 1188 Baseline)',
        itemEn: 'Internal Ohmic Resistance Baseline Test (IEEE 1188)',
        measured: 'Trung bình: 4.25 mΩ (Độ lệch cực đại giữa các cell: 6.2%)',
        standard: 'Nội trở không được vượt quá 20% so với giá trị cơ sở của nhà sản xuất',
        standardEn: 'Internal resistance deviation must not exceed 20% from baseline',
        reference: 'IEEE Std 1188 & NETA ATS Sec 7.18.2.D.2',
        status: 'PASS',
        detail: 'Tất cả 40 bình ắc quy đều có nội trở đồng nhất, không có cell bị sulfat hóa hay khô dung dịch.',
        detailEn: 'All 40 units exhibit healthy internal plate integrity with no dry-out.',
      },
      {
        item: 'Điện áp phao nạp từng bình Float Voltage (Float Charge Verification)',
        itemEn: 'Individual Cell Float Voltage Uniformity',
        measured: 'Dải đo: 13.59V đến 13.65V (Trung bình 13.62V/bình tại 25°C)',
        standard: 'Điện áp từng bình nằm trong dung sai ± 0.05V/cell (± 0.15V/bình 12V)',
        standardEn: 'Float voltage must match manufacturer specifications within ± 0.05V/cell',
        reference: 'NETA ATS-2025 Table 100.12 & IEEE 1188',
        status: 'PASS',
        detail: 'Điện áp nạp cân bằng hoàn hảo trên toàn chuỗi 240V DC.',
        detailEn: 'Float charge distribution is well balanced across the 240V DC string.',
      },
      {
        item: 'Điện trở cầu nối liên kết các bình (Intercell Connector Resistance DLRO)',
        itemEn: 'Intercell Connector Resistance (DLRO Test)',
        measured: '18.5 µΩ (Độ lệch cực đại 4.1%)',
        standard: 'Độ lệch không vượt quá 50% so với giá trị nhỏ nhất của các cầu nối',
        standardEn: 'Connector resistance deviation must not exceed 50% of the lowest joint',
        reference: 'NETA ATS-2025 Sec 7.18.2.D.1',
        status: 'PASS',
        detail: 'Bu-lông cọc bình siết đúng lực 11 Nm, bôi mỡ mạ bảo vệ chống oxy hóa.',
        detailEn: 'Terminal bolts torqued to 11 Nm and coated with anti-corrosion grease.',
      },
      {
        item: 'Kiểm tra thị giác, phồng bình & Nhiệt độ bề mặt (Thermal & Physical Scan)',
        itemEn: 'Visual Inspection, Case Bulging & Thermal Runaway Check',
        measured: 'Vỏ bình phẳng, van an toàn kín, nhiệt độ bề mặt đồng đều 24.8°C',
        standard: 'Không phồng rộp vỏ bình, không rò rỉ axit, chênh lệch nhiệt độ < 3°C',
        standardEn: 'Zero case bulging, no electrolyte leakage, temperature spread < 3°C',
        reference: 'NETA ATS-2025 Sec 7.18.2.1',
        status: 'PASS',
        detail: 'Phòng ắc quy duy trì nhiệt độ điều hòa 24°C, thông gió đối lưu tốt.',
        detailEn: 'Battery room maintained at optimal 24°C with active air exchange.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Nội Trở Bình & Nguy Cơ Lão Hóa Bản Cực',
        titleEn: 'Internal Ohmic Health & Plate Sulfation Prevention',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          'Độ lệch nội trở 6.2% phản ánh chuỗi ắc quy mới xuất xưởng đạt chất lượng đồng bộ.',
        ],
        pointsEn: [
          '6.2% resistance spread indicates factory-fresh, well-matched AGM cells.',
        ],
      },
      {
        titleVi: 'Điện Trở Cầu Nối Cút & Nguy Cơ Quá Nhiệt Phát Hỏa',
        titleEn: 'Intercell Joint Integrity & Fire Prevention',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Điện trở 18.5 µΩ loại trừ nguy cơ sụt áp và sinh nhiệt cục bộ khi phóng dòng lớn 400A.',
        ],
        pointsEn: [
          '18.5 µΩ ensures zero dangerous voltage drops during 400A high-rate discharge.',
        ],
      },
      {
        titleVi: 'Kiểm Soát Hiện Tượng Tự Sinh Nhiệt Thất Thoát (Thermal Runaway)',
        titleEn: 'Thermal Runaway Mitigation',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          'Bộ nạp UPS có cảm biến nhiệt độ gắn trên chuỗi ắc quy để tự động bù nhiệt điện áp nạp (Temperature Compensation).',
        ],
        pointsEn: [
          'UPS charger utilizes battery temperature sensor for dynamic voltage compensation.',
        ],
      },
      {
        titleVi: 'Độ Cách Điện Giá Đỡ Ắc Quy Với Đất',
        titleEn: 'Battery Rack Ground Isolation',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Giá đỡ thép sơn tĩnh điện cách điện với đất đạt > 100 MΩ, rơ-le chạm đất DC không báo lỗi.',
        ],
        pointsEn: [
          'Epoxy-coated battery racks maintain > 100 MΩ ground isolation; DC ground detector clear.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện đưa vào vận hành nuôi tải hệ thống UPS.',
      'Lưu trữ hồ sơ nội trở cơ sở Baseline vào hệ thống giám sát pin BMS để theo dõi xu hướng.',
      'Tiến hành thử nghiệm xả tải thực tế (Battery Capacity Discharge Test) sau 2 năm vận hành.',
    ],
    recommendationsEn: [
      'Approved for live UPS DC bus float service.',
      'Archive baseline internal ohmic values in BMS for trending analysis.',
      'Schedule full battery capacity discharge load test after 2 years of operation.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Phòng Nguồn UPS Trung Tâm Dữ Liệu TEV Data Center',
        battery_tag: 'BAT-VRLA-240V-STRING-01',
        fse_name: 'Lương Thế Vinh (vinh.lt@tev.vn)',
        test_date: '2026-09-30',
      },
      electrical_tests: {
        internal_resistance_avg_milli_ohms: 4.25,
        internal_resistance_max_spread_percent: 6.2,
        float_voltage_avg_volts: 13.62,
        intercell_resistance_micro_ohms: 18.5,
        ambient_temp_c: 24.8,
      },
    },
  },

  // 15. FLOODED LEAD-ACID BATTERIES (Sec 7.18.1 & IEEE 450)
  'neta-battery-flooded': {
    id: 'neta-battery-flooded',
    sectionCode: 'Mục 7.18.1',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.18.1 & IEEE 450',
    equipmentTitleVi: 'Hệ Thống Ắc Quy Chì-Axit Hở Khí (Flooded Lead-Acid Batteries)',
    equipmentTitleEn: 'Flooded Lead-Acid Battery Systems (Vented Stationary Cells)',
    subtitleVi: 'Thị giác • Tỷ trọng Điện dịch (1.215 ±0.01) • Mức dung dịch • Điện áp Phao Float • Cầu nối Cút (≤ 50%)',
    subtitleEn: 'Visual • Specific Gravity (1.215 ±0.01) • Electrolyte Level • Float Voltage • Intercell Connector DLRO',
    category: 'special',
    defaultSiteInfo: {
      project_name: 'Trạm Biến Áp 220kV TEV Transmission Substation',
      equipment_tag: 'BAT-FL-220V-BANK-01',
      equipment_rating: '220V DC 110x 2V 500Ah Flooded Lead-Acid Stationary Cells',
      fse_name: 'Nguyễn Văn Định (dinh.nv@tev.vn)',
      test_date: '2026-09-29',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / HỆ THỐNG ẮC QUY HỞ KHÍ ĐẠT CHUẨN NETA ATS-2025 MỤC 7.18.1',
    statusTitleEn: 'PASS / FLOODED BATTERY SYSTEM COMPLIANT WITH NETA ATS-2025 & IEEE 450',
    statusReasonVi: 'Tỷ trọng dung dịch điện phân axit sunfuric quy đổi về 25°C đạt 1.215 ±0.005 (nằm hoàn hảo trong ngưỡng 1.215 ±0.010 theo IEEE 450), mức dung dịch nằm giữa vạch Max-Min, điện áp phao đạt 2.22V/cell, và điện trở cầu nối intercell đạt 14.2 µΩ.',
    statusReasonEn: 'Electrolyte specific gravity corrected to 25°C is 1.215 ±0.005 (meeting 1.215 ±0.010 per IEEE 450), electrolyte level is within Max-Min lines, float voltage is 2.22V/cell, and intercell resistance is 14.2 µΩ.',
    evaluationItems: [
      {
        item: 'Tỷ trọng dung dịch điện dịch quy đổi 25°C (Specific Gravity per IEEE 450)',
        itemEn: 'Electrolyte Specific Gravity Corrected to 25°C (Hydrometer)',
        measured: '1.212 đến 1.218 (Trung bình 1.215 tại 25°C)',
        standard: 'Tỷ trọng định mức 1.215 ± 0.010 (Độ lệch giữa các bình cực đại ≤ 0.010)',
        standardEn: 'Rated specific gravity 1.215 ± 0.010; max cell deviation ≤ 0.010',
        reference: 'IEEE Std 450 & NETA ATS Sec 7.18.1.D.3',
        status: 'PASS',
        detail: 'Đo bằng tỷ trọng kế quang học chính xác, nồng độ axit sulfuric chuẩn xác.',
        detailEn: 'Optical refractometer measurement confirms optimal acid concentration.',
      },
      {
        item: 'Điện áp phao nạp từng bình Float Voltage (Float Charge Verification)',
        itemEn: 'Individual Cell Float Voltage Uniformity',
        measured: '2.21V đến 2.23V (Trung bình 2.22V/cell tại 25°C)',
        standard: 'Điện áp phao nằm trong dải dung sai nhà sản xuất (2.20V - 2.25V/cell)',
        standardEn: 'Float voltage must match manufacturer range (2.20V - 2.25V/cell)',
        reference: 'NETA ATS-2025 Sec 7.18.1.D.2',
        status: 'PASS',
        detail: 'Tất cả 110 cell 2V đều nhận điện áp nạp duy trì đồng đều.',
        detailEn: 'All 110 series cells receive uniform trickle float charge.',
      },
      {
        item: 'Điện trở cầu nối liên kết các bình (Intercell Connector Resistance DLRO)',
        itemEn: 'Intercell Connector Resistance (DLRO Test)',
        measured: '14.2 µΩ (Độ lệch cực đại giữa các mối nối 3.5%)',
        standard: 'Độ lệch không vượt quá 50% so với giá trị nhỏ nhất của các cầu nối',
        standardEn: 'Deviation must not exceed 50% of lowest similar connector',
        reference: 'NETA ATS-2025 Table 100.12',
        status: 'PASS',
        detail: 'Thanh đồng nối bọc chì chống axit siết đúng lực cờ lê cân lực 14 Nm.',
        detailEn: 'Lead-plated copper intercell straps torqued to 14 Nm with no corrosion.',
      },
      {
        item: 'Kiểm tra thị giác, Mức dung dịch & Nắp phòng nổ Flame Arrestor',
        itemEn: 'Visual, Electrolyte Levels & Flame Arrestor Caps Check',
        measured: 'Mức axit đều ở vạch Max, nắp phòng nổ sạch xỉ, bình trong suốt không nứt',
        standard: 'Mức dung dịch nằm giữa Min-Max, nắp phòng nổ thông thoáng chống bắt lửa',
        standardEn: 'Electrolyte between marks, clean flame arrestor vents prevent explosion',
        reference: 'NETA ATS-2025 Sec 7.18.1.1',
        status: 'PASS',
        detail: 'Bình nhựa SAN trong suốt dễ quan sát bản cực chì, không có cặn bùn đáy bình.',
        detailEn: 'Transparent SAN containers show healthy plates with zero sediment buildup.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Nồng Độ Điện Phân & Chất Lượng Bản Cực Chì',
        titleEn: 'Electrolyte Density & Active Material Health',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          'Tỷ trọng 1.215 chứng minh bình ắc quy được nạp no hoàn toàn (100% State of Charge).',
        ],
        pointsEn: [
          '1.215 specific gravity proves 100% State of Charge (SOC) across all cells.',
        ],
      },
      {
        titleVi: 'Hệ Thống Thoát Khí Hydro & Nắp Chống Nổ',
        titleEn: 'Hydrogen Gas Ventilation & Explosion Safety',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          'Nắp lọc gốm Flame Arrestor ngăn tia lửa bên ngoài lọt vào gây nổ khí Hydro.',
          'Quạt thông gió phòng ắc quy chống cháy nổ duy trì nồng độ H2 < 1.0% per NFPA 1.',
        ],
        pointsEn: [
          'Flame arrestor vents prevent external spark propagation into cell headspace.',
          'Explosion-proof exhaust ventilation maintains H2 concentration < 1.0% per NFPA 1.',
        ],
      },
      {
        titleVi: 'Độ Bền Cơ Khí Bản Cực & Đáy Bình Lắng Cặn',
        titleEn: 'Plate Alignment & Sediment Chamber Depth',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Không có hiện tượng cong vênh bản cực dương (Positive Plate Growth).',
        ],
        pointsEn: [
          'No plate distortion or excessive positive plate growth observed.',
        ],
      },
      {
        titleVi: 'Tiếp Xúc Cút Nối Chì Mạ Đồng',
        titleEn: 'Lead-Coated Copper Intercell Joint Health',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Điện trở 14.2 µΩ bảo đảm thông suốt khi cấp dòng cắt máy cắt 220kV lên tới 200A.',
        ],
        pointsEn: [
          '14.2 µΩ guarantees zero nuisance voltage drop during 220kV trip coil pulses.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện cấp nguồn cho hệ thống bảo vệ nhị thứ và đóng cắt trạm 220kV.',
      'Bổ sung nước cất demineralized water định kỳ khi mức dung dịch chạm vạch Min.',
      'Trang bị đầy đủ bồn rửa mắt khẩn cấp và dung dịch trung hòa axit trong phòng ắc quy.',
    ],
    recommendationsEn: [
      'Approved for 220kV substation protective relaying and switchgear DC bus service.',
      'Replenish electrolyte with demineralized distilled water when levels reach Min.',
      'Verify emergency eyewash station and acid spill neutralizing kits are operational.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Trạm Biến Áp 220kV TEV Transmission Substation',
        battery_tag: 'BAT-FL-220V-BANK-01',
        fse_name: 'Nguyễn Văn Định (dinh.nv@tev.vn)',
        test_date: '2026-09-29',
      },
      electrical_tests: {
        specific_gravity_avg: 1.215,
        specific_gravity_max_deviation: 0.003,
        float_voltage_avg: 2.22,
        intercell_resistance_micro_ohms: 14.2,
      },
    },
  },

  // 16. STATIC UPS SYSTEMS (Sec 7.22.2)
  'neta-ups': {
    id: 'neta-ups',
    sectionCode: 'Mục 7.22.2',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.22.2 & IEEE 1184',
    equipmentTitleVi: 'Bộ Lưu Điện Tĩnh UPS (Static Uninterruptible Power Supply Systems)',
    equipmentTitleEn: 'Static Uninterruptible Power Supply (Static UPS Systems)',
    subtitleVi: 'Thị giác • Cách điện (Table 100.1) • Chuyển mạch Static Bypass • Thử nghiệm 100% Load Step • Sóng hài THD',
    subtitleEn: 'Visual • Insulation (Table 100.1) • Static Bypass Transfer • 100% Load Step Response • Output THD',
    category: 'special',
    defaultSiteInfo: {
      project_name: 'Trung Tâm Dữ Liệu TEV Cloud Data Center',
      equipment_tag: 'UPS-500KVA-UNIT-01',
      equipment_rating: '500 kVA 400V 3-Phase On-Line Double-Conversion UPS',
      fse_name: 'Võ Minh Tú (sgm1707@gmail.com)',
      test_date: '2026-09-28',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / HỆ THỐNG UPS ĐẠT CHUẨN NETA ATS-2025 MỤC 7.22.2',
    statusTitleEn: 'PASS / UPS FULLY COMPLIANT WITH NETA ATS-2025 SECTION 7.22.2',
    statusReasonVi: 'Thời gian chuyển mạch Static Bypass không gián đoạn điện áp (0 ms transfer time), độ méo hài tổng điện áp ngõ ra THDv chỉ 1.4% (ngưỡng tiêu chuẩn < 3.0%), bước nhảy tải 100% phục hồi trong 8 ms, và hiệu suất vận hành đạt 96.8%.',
    statusReasonEn: 'Static bypass transfer executed seamlessly with 0 ms interruption, output voltage distortion THDv is 1.4% (< 3.0% limit), 100% load step transient recovered in 8 ms, and efficiency is 96.8%.',
    evaluationItems: [
      {
        item: 'Thời gian chuyển mạch ngõ ra Static Bypass (Static Switch Transfer Test)',
        itemEn: 'Static Bypass Seamless Transfer Test (Inverter to Bypass & Return)',
        measured: '0 ms (Không có gián đoạn dạng sóng sin ngõ ra)',
        standard: 'Chuyển mạch tĩnh không gián đoạn cung cấp điện cho tải IT (≤ 4 ms)',
        standardEn: 'Uninterrupted static transfer to bypass and back (≤ 4 ms)',
        reference: 'NETA ATS-2025 Sec 7.22.2.D.3',
        status: 'PASS',
        detail: 'Bộ chuyển mạch tĩnh Thyristor SCR kích hoạt tức thì, không gây khởi động lại tải máy chủ.',
        detailEn: 'Thyristor static switch transfers seamlessly without server power interruption.',
      },
      {
        item: 'Độ méo sóng hài điện áp ngõ ra (Output Voltage Total Harmonic Distortion THDv)',
        itemEn: 'Output Voltage THD at 100% Linear & Non-Linear Load',
        measured: 'THDv = 1.4% (Tải tuyến tính) | THDv = 2.6% (Tải phi tuyến tính)',
        standard: 'Độ méo hài tổng điện áp BẮT BUỘC THDv < 3.0% (Tuyến tính) / < 5.0% (Phi tuyến)',
        standardEn: 'Total Voltage Harmonic Distortion MUST be THDv < 3.0% linear / < 5.0% non-linear',
        reference: 'IEEE Std 519 & NETA ATS Sec 7.22.2.D.4',
        status: 'PASS',
        detail: 'Bộ nghịch lưu IGBT 3 bậc phát sóng sin chuẩn, triệt tiêu tối đa sóng hài bậc cao.',
        detailEn: 'Three-level IGBT inverter delivers pristine sinusoidal power.',
      },
      {
        item: 'Phản ứng quá độ khi nhận bước tải 100% (Transient Load Step Response)',
        itemEn: '100% Load Step Transient Voltage Deviation & Recovery Time',
        measured: 'Độ lệch điện áp: ±1.8% | Thời gian phục hồi: 8 ms (Dưới 1/2 chu kỳ)',
        standard: 'Độ sụt điện áp ≤ ±5.0%, thời gian phục hồi về dải định mức ≤ 20 ms',
        standardEn: 'Transient deviation ≤ ±5.0%; recovery time to normal band ≤ 20 ms',
        reference: 'NETA ATS-2025 Sec 7.22.2.D.5',
        status: 'PASS',
        detail: 'Khả năng đáp ứng tải động cực kỳ nhanh nhạy, bảo vệ an toàn cho thiết bị tính toán nhạy cảm.',
        detailEn: 'Dynamic load response ensures uninterrupted power for sensitive computing IT loads.',
      },
      {
        item: 'Điện trở cách điện Đầu vào / Đầu ra / Khung vỏ (1000V DC per Table 100.1)',
        itemEn: 'Insulation Resistance Input/Output/Bypass to Ground',
        measured: 'Input-G: 450 MΩ | Output-G: 520 MΩ | DC Bus-G: 680 MΩ',
        standard: 'Hệ thống thiết bị hạ thế BẮT BUỘC ≥ 100 MΩ per Table 100.1',
        standardEn: 'LV power assemblies MUST be ≥ 100 MΩ per Table 100.1',
        reference: 'NETA ATS-2025 Table 100.1',
        status: 'PASS',
        detail: 'Cách điện toàn bộ các khối biến lưu (Rectifier, Inverter, Bypass) đạt tiêu chuẩn.',
        detailEn: 'Rectifier, inverter, and bypass isolation withstand 1000V DC testing.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Chất Lượng Chuyển Mạch Tĩnh Static Switch',
        titleEn: 'Static Bypass Switch Seamless Transition',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Chuyển mạch 0 ms loại trừ nguy cơ sụt áp làm sập hệ thống máy chủ Data Center.',
        ],
        pointsEn: [
          '0 ms transfer eliminates voltage sags that crash server power supplies.',
        ],
      },
      {
        titleVi: 'Sóng Hài Ngõ Ra & Lọc Sóng Tích Cực',
        titleEn: 'Harmonic Distortion & Active Filtering',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'THDv 1.4% bảo đảm không làm nóng động cơ quạt làm mát và tụ điện máy chủ.',
        ],
        pointsEn: [
          'THDv of 1.4% prevents thermal stress on server power factor correction stages.',
        ],
      },
      {
        titleVi: 'Phản Hồi Quá Độ Khi Đột Ngột Mất Điện Lưới',
        titleEn: 'Mains Outage Instantaneous Battery Handover',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Khi ngắt CB nguồn điện lưới AC, DC Bus chuyển sang ắc quy với độ lệch điện áp bằng 0.',
        ],
        pointsEn: [
          'Seamless transition to DC battery bus during mains failure with zero voltage notch.',
        ],
      },
      {
        titleVi: 'Hệ Thống Tản Nhiệt Quạt IGBT & Cảnh Báo Lọc Gió',
        titleEn: 'IGBT Heatsink Forced Cooling & Filter Status',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          'Quạt làm mát module IGBT điều tốc thông minh duy trì nhiệt độ cánh tản nhiệt < 62°C.',
        ],
        pointsEn: [
          'Variable speed cooling fans maintain IGBT heatsink temperatures safely under 62°C.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện cấp nguồn nuôi tải IT Critical Load.',
      'Khóa liên động tay gạt chuyển mạch Maintenance Bypass Switch ở vị trí Auto/Normal.',
      'Lập lịch thay thế quạt làm mát và tụ điện lọc DC sau 5 năm vận hành (Preventive Life-Cycle).',
    ],
    recommendationsEn: [
      'Approved for critical IT load power delivery.',
      'Lock maintenance bypass interlock handle in Normal/Auto position.',
      'Schedule 5-year preventive replacement for DC filter capacitors and cooling fans.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Trung Tâm Dữ Liệu TEV Cloud Data Center',
        ups_tag: 'UPS-500KVA-UNIT-01',
        fse_name: 'Võ Minh Tú (sgm1707@gmail.com)',
        test_date: '2026-09-28',
      },
      electrical_tests: {
        static_transfer_time_ms: 0,
        output_thd_linear_percent: 1.4,
        output_thd_nonlinear_percent: 2.6,
        transient_recovery_time_ms: 8,
        insulation_resistance_megohms: 520,
      },
    },
  },

  // 17. PROTECTIVE RELAYS (Sec 7.9)
  'neta-relay': {
    id: 'neta-relay',
    sectionCode: 'Mục 7.9',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.9 & Table 100.9',
    equipmentTitleVi: 'Rơ-le Bảo Vệ Kỹ Thuật Số & Điện Cơ (Protective Relays)',
    equipmentTitleEn: 'Protective Relays (Microprocessor & Electromechanical Relays)',
    subtitleVi: 'Thị giác • Cách điện mạch đo • Bơm dòng nhị thứ (50/51/51N/87) • Thời gian cắt (±5%) • Ngắt máy cắt Trip test',
    subtitleEn: 'Visual • Relay Insulation • Secondary Current Injection (50/51/87) • Timing Accuracy (±5%) • Trip Breaker Circuit',
    category: 'protection',
    defaultSiteInfo: {
      project_name: 'Trạm Biến Áp 110kV TEV Industrial Substation',
      equipment_tag: 'RLY-FEEDER-SEL751-01',
      equipment_rating: 'SEL-751 Feeder Protection Relay (50/51, 50N/51N, 67, 81, 79)',
      fse_name: 'Trần Đại Nghĩa (nghia.td@tev.vn)',
      test_date: '2026-09-27',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / RƠ-LE BẢO VỆ ĐẠT CHUẨN NETA ATS-2025 MỤC 7.9',
    statusTitleEn: 'PASS / PROTECTIVE RELAY COMPLIANT WITH NETA ATS-2025 SECTION 7.9',
    statusReasonVi: 'Bơm dòng thứ cấp thử nghiệm chức năng quá dòng chạm đất 50/51/51N đạt sai số dòng tác động < 1.2% (ngưỡng cho phép ±3.0%), sai số thời gian cắt đường cong IEC Normal Inverse chỉ 2.1% (ngưỡng ±5.0%), tiếp điểm đầu ra ngắt máy cắt trip circuit tác động chuẩn xác.',
    statusReasonEn: 'Secondary injection testing for 50/51/51N overcurrent pickup error is < 1.2% (±3.0% allowed), trip timing deviation on IEC Normal Inverse curve is 2.1% (±5.0% allowed), and breaker trip circuit operates cleanly.',
    evaluationItems: [
      {
        item: 'Ngưỡng dòng tác động Pickup (Pickup Current Verification per ANSI 50/51)',
        itemEn: 'Pickup Current Accuracy (Functions 50, 51, 51N)',
        measured: '51P Pickup: 1.01 A (Cài đặt 1.00 A, Sai số +1.0%) | 50P: 5.04 A (Cài đặt 5.00 A, Sai số +0.8%)',
        standard: 'Sai số dòng tác động rơ-le số BẮT BUỘC ≤ ± 3.0% (Điện cơ ≤ ± 5.0%)',
        standardEn: 'Microprocessor relay pickup accuracy MUST be ≤ ± 3.0% per Table 100.9',
        reference: 'NETA ATS-2025 Table 100.9 & Sec 7.9.D.2',
        status: 'PASS',
        detail: 'Ngưỡng tác động cực kỳ chính xác, đảm bảo độ chọn lọc bảo vệ phối hợp.',
        detailEn: 'Pickup thresholds highly precise, guaranteeing protection coordination.',
      },
      {
        item: 'Thời gian tác động theo đặc tuyến thời gian (Operating Timing per TCC Curve)',
        itemEn: 'Time-Current Characteristic (TCC) Curve Timing Accuracy',
        measured: 'Tại dòng 3x Pickup: 0.812s (Tính toán lý thuyết 0.795s, Sai số +2.1%)',
        standard: 'Sai số thời gian tác động BẮT BUỘC nằm trong dải ± 5.0% hoặc ± 1.5 chu kỳ',
        standardEn: 'Operating time error MUST fall within ± 5.0% or ± 1.5 cycles',
        reference: 'NETA ATS-2025 Table 100.9',
        status: 'PASS',
        detail: 'Đường cong dốc phụ thuộc thời gian khớp hoàn hảo với bảng tính toán phối hợp bảo vệ.',
        detailEn: 'Inverse-time curve tracking matches coordination study specifications.',
      },
      {
        item: 'Thử nghiệm Ngắt máy cắt thực tế (Breaker Trip Circuit & Target Test)',
        itemEn: 'Complete Breaker Trip Verification via Relay Output Contacts',
        measured: 'Tiếp điểm OUT101 kích hoạt cuộn cắt máy cắt mở trong 42 ms; Đèn LED Target sáng',
        standard: 'Rơ-le phải kích hoạt ngắt mở máy cắt thực tế thành công, hiển thị cờ báo sự cố',
        standardEn: 'Relay must initiate successful physical circuit breaker trip and display fault target',
        reference: 'NETA ATS-2025 Sec 7.9.D.5',
        status: 'PASS',
        detail: 'Mạch cắt trip coil và mạch giám sát cuộn cắt TCS (Trip Circuit Supervision) hoạt động tốt.',
        detailEn: 'Breaker trip mechanism and Trip Circuit Supervision (TCS) fully verified.',
      },
      {
        item: 'Điện trở cách điện Mạch đo CT/VT & Nguồn nuôi (500V DC per Table 100.1)',
        itemEn: 'Relay CT/VT Sensing & Control Wiring Insulation Resistance',
        measured: 'Mạch dòng CT-G: 58 MΩ | Mạch áp VT-G: 64 MΩ | Nguồn phụ DC-G: 72 MΩ',
        standard: 'Cách điện mạch nhị thứ rơ-le BẮT BUỘC ≥ 2.0 MΩ (Table 100.1)',
        standardEn: 'Relay secondary circuits MUST be ≥ 2.0 MΩ per Table 100.1',
        reference: 'NETA ATS-2025 Sec 7.9.D.1',
        status: 'PASS',
        detail: 'Cách điện mạch đo lường an toàn, không có nguy cơ ngắn mạch biến dòng CT.',
        detailEn: 'Sensor wiring safely isolated; zero danger of CT secondary shorting.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Độ Chọn Lọc & Phối Hợp Bảo Vệ Rơ-le (Selectivity & Margin)',
        titleEn: 'Protection Selectivity & Grading Margin',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Sai số 2.1% bảo đảm khoảng thời gian phối hợp dự phòng CTI 250ms không bị vi phạm.',
        ],
        pointsEn: [
          '2.1% timing accuracy prevents violation of 250ms Coordination Time Interval (CTI).',
        ],
      },
      {
        titleVi: 'Mạch Giám Sát Cuộn Cắt (Trip Circuit Supervision TCS)',
        titleEn: 'Trip Circuit Supervision (TCS) Monitoring',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Rơ-le cảnh báo ngay lập tức nếu mạch cuộn cắt bị đứt dây hoặc mất nguồn DC 110V.',
        ],
        pointsEn: [
          'Relay sounds instant alarm if trip circuit open-circuit or DC supply lost.',
        ],
      },
      {
        titleVi: 'Biến Dòng Đo Lường CT & Cực Tính Phân Cực',
        titleEn: 'CT Polarity & Magnetization Saturation Margin',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Cực tính CT (H1-X1) đấu nối chuẩn xác, tỷ số biến dòng đo lường sai số < 0.2%.',
        ],
        pointsEn: [
          'CT polarity (H1-X1) verified correct; CT ratio error well under 0.2%.',
        ],
      },
      {
        titleVi: 'Bộ Nhớ Ghi Sự Cố & Đồng Bộ Thời Gian (Fault Event & SOE)',
        titleEn: 'Digital Fault Recording & SNTP Time Synchronization',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Đồng hồ rơ-le đồng bộ thời gian SNTP qua mạng quang học với độ lệch < 1 ms.',
        ],
        pointsEn: [
          'SNTP network time sync locks internal clock within 1 ms resolution.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện kích hoạt bảo vệ đóng điện đưa vào vận hành.',
      'Sao lưu tệp cấu hình cài đặt Setting File (.rdb/.mxf) vào kho dữ liệu quản lý CMMS.',
      'Bàn giao bảng chỉnh định rơ-le có chữ ký phê duyệt của kỹ sư trưởng điện.',
    ],
    recommendationsEn: [
      'Approved for live system protection energization.',
      'Archive approved relay setting files (.rdb/.mxf) into cloud CMMS repository.',
      'Sign off approved protection setting coordination study documentation.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Trạm Biến Áp 110kV TEV Industrial Substation',
        relay_tag: 'RLY-FEEDER-SEL751-01',
        fse_name: 'Trần Đại Nghĩa (nghia.td@tev.vn)',
        test_date: '2026-09-27',
      },
      electrical_tests: {
        pickup_51p_amps: 1.01,
        pickup_50p_amps: 5.04,
        trip_time_at_3x_seconds: 0.812,
        timing_error_percent: 2.1,
        breaker_trip_verified: true,
      },
    },
  },

  // 18. SF6 GAS-INSULATED SWITCHES & BREAKERS (Sec 7.3.3)
  'neta-sf6-switch': {
    id: 'neta-sf6-switch',
    sectionCode: 'Mục 7.3.3',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.3.3 & IEEE C37.122',
    equipmentTitleVi: 'Máy Cắt & Dao Cách Ly Khí SF6 GIS (SF6 Gas-Insulated Switches & GIS)',
    equipmentTitleEn: 'SF6 Gas-Insulated Switches & Breakers (GIS Assemblies)',
    subtitleVi: 'Thị giác • Áp suất khí SF6 • Độ ẩm khí SF6 (ASTM D2029) • Rò rỉ khí SF6 • Điện trở tiếp xúc cực DLRO',
    subtitleEn: 'Visual • SF6 Gas Pressure • Gas Moisture (ASTM D2029) • SF6 Leak Detection • Contact DLRO',
    category: 'switches',
    defaultSiteInfo: {
      project_name: 'Trạm Biến Áp GIS 110kV TEV Urban Substation',
      equipment_tag: 'GIS-110KV-BAY-01',
      equipment_rating: '110kV 2000A 31.5kA SF6 Gas-Insulated Switchgear',
      fse_name: 'Phan Văn Đức (duc.pv@tev.vn)',
      test_date: '2026-09-26',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / THIẾT BỊ KHÍ SF6 ĐẠT TIÊU CHUẨN NETA ATS-2025 MỤC 7.3.3',
    statusTitleEn: 'PASS / SF6 GIS APPARATUS COMPLIANT WITH NETA ATS-2025 SEC 7.3.3',
    statusReasonVi: 'Áp suất khí SF6 duy trì ổn định ở 0.60 MPa (ngưỡng định mức 0.60 MPa tại 20°C), hàm lượng nước đo theo ASTM D2029 đạt 48 ppmv (ngưỡng tối đa cho khí mới nạp ≤ 150 ppmv theo Bảng Table 100.4), không phát hiện rò rỉ khí (Leak rate < 0.5%/năm), và điện trở tiếp điểm cực chính đạt 28.5 µΩ.',
    statusReasonEn: 'SF6 gas pressure is steady at 0.60 MPa (rated 0.60 MPa at 20°C), moisture per ASTM D2029 is 48 ppmv (≤ 150 ppmv Table 100.4 limit for new gas), zero SF6 leaks detected (< 0.5%/year), and main contact resistance is 28.5 µΩ.',
    evaluationItems: [
      {
        item: 'Hàm lượng nước trong khí SF6 (Moisture Content per ASTM D2029)',
        itemEn: 'SF6 Gas Moisture Content @ 20°C (ASTM D2029)',
        measured: '48 ppmv (Điểm sương Dew Point -34°C)',
        standard: 'Khí SF6 mới nạp vào thiết bị BẮT BUỘC ≤ 150 ppmv (Điểm sương ≤ -25°C)',
        standardEn: 'New SF6 fill MUST be ≤ 150 ppmv per Table 100.4 (Dew point ≤ -25°C)',
        reference: 'NETA ATS-2025 Table 100.4 & IEEE C37.122',
        status: 'PASS',
        detail: 'Khí SF6 cực kỳ khô ráo, loại trừ nguy cơ phản ứng với hồ quang sinh axit HF ăn mòn.',
        detailEn: 'Ultra-dry SF6 gas prevents formation of corrosive HF decomposition byproducts.',
      },
      {
        item: 'Kiểm tra Rò rỉ khí SF6 (SF6 Gas Leakage Detection via Sniffer/IR)',
        itemEn: 'SF6 Gas Sniffer / Optical Gas Imaging Leak Test',
        measured: 'Tỷ lệ rò rỉ: 0.0 ppm (Tốc độ rò rỉ < 0.1%/năm)',
        standard: 'Không phát hiện bất kỳ điểm rò rỉ khí nào (Tối đa cho phép < 0.5%/năm)',
        standardEn: 'Zero detectable leaks; annual leakage rate MUST be < 0.5% per year',
        reference: 'NETA ATS-2025 Sec 7.3.3.1',
        status: 'PASS',
        detail: 'Kiểm tra máy đo laser quang học tại toàn bộ các mặt bích gioăng O-ring đều kín tuyệt đối.',
        detailEn: 'Infrared optical gas imaging verifies 100% gastight seals on all O-ring flanges.',
      },
      {
        item: 'Điện trở tiếp xúc cực chính đóng mạch (Main Contact DLRO @ 100A DC)',
        itemEn: 'Main Breaker Contact Resistance (DLRO Test)',
        measured: 'Pha A: 28.2 µΩ | Pha B: 28.5 µΩ | Pha C: 28.4 µΩ (Lệch max 1.1%)',
        standard: 'Độ lệch không vượt quá 50% so với giá trị nhỏ nhất của các cực',
        standardEn: 'Contact resistance deviation must not exceed 50% of lowest pole',
        reference: 'NETA ATS-2025 Table 100.12',
        status: 'PASS',
        detail: 'Tiếp điểm hồ quang (Arcurate contacts) và tiếp điểm mang dòng chính đạt trạng thái hoàn hảo.',
        detailEn: 'Both main and arcing contacts show minimal contact resistance.',
      },
      {
        item: 'Thử nghiệm Đồng thời gian Đóng/Mở máy cắt (Breaker Timing Test)',
        itemEn: 'Circuit Breaker Close & Open Timing & Contact Synchronization',
        measured: 'Thời gian mở: 22 ms | Thời gian đóng: 48 ms | Độ bất đồng thời: 0.8 ms',
        standard: 'Độ bất đồng thời giữa 3 pha BẮT BUỘC ≤ 2.0 ms per IEEE C37.09',
        standardEn: 'Pole-to-pole synchronization MUST be ≤ 2.0 ms per IEEE C37.09',
        reference: 'IEEE C37.09 & NETA ATS-2025 Sec 7.3.3',
        status: 'PASS',
        detail: 'Cơ cấu truyền động lò xo thao tác dứt khoát, đóng mở 3 pha đồng thời tuyệt đối.',
        detailEn: 'Spring operating mechanism delivers fast, synchronized contact motion.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Độ Khô Khí SF6 & Điểm Sương Dew Point',
        titleEn: 'SF6 Gas Dryness & Dew Point Margin',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          '48 ppmv cách xa ngưỡng đóng băng hơi nước ở nhiệt độ môi trường thấp.',
        ],
        pointsEn: [
          '48 ppmv ensures zero condensation risk even under freezing winter ambients.',
        ],
      },
      {
        titleVi: 'Độ Kín Khoang Khí & Rơ-le Giám Sát Áp Suất Khí Mật Độ',
        titleEn: 'Gastight Chamber & Density Monitor Integrity',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Đồng hồ đo mật độ khí SF6 bù nhiệt độ (Temperature-Compensated Density Monitor) hoạt động chính xác.',
        ],
        pointsEn: [
          'Temperature-compensated density monitor calibrated with functional alarm/block stages.',
        ],
      },
      {
        titleVi: 'Đồng Thời Gian Đóng Mở Máy Cắt 3 Pha',
        titleEn: 'Three-Pole Synchronization Dynamics',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Độ lệch 0.8 ms ngăn chặn xuất hiện sóng quá điện áp đóng cắt không đối xứng.',
        ],
        pointsEn: [
          '0.8 ms pole spread prevents asymmetrical switching transient overvoltages.',
        ],
      },
      {
        titleVi: 'Điện Trở Tiếp Điểm Cực & Khả Năng Cắt Dòng Ngắn Mạch',
        titleEn: 'Pole Contact Resistance & Short-Circuit Capability',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          '28.5 µΩ bảo đảm máy cắt chịu dòng ngắn mạch định mức 31.5kA trong 3 giây.',
        ],
        pointsEn: [
          '28.5 µΩ certifies thermal rating to carry 31.5kA fault currents for 3 seconds.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện đóng điện nghiệm thu đưa trạm GIS 110kV vào vận hành.',
      'Ghi nhận chỉ số áp suất ban đầu vào nhật ký vận hành để theo dõi rò rỉ hàng tháng.',
      'Đảm bảo hệ thống cảm biến rò rỉ khí SF6 và quạt thông gió đáy hầm phòng GIS hoạt động 24/7.',
    ],
    recommendationsEn: [
      'Approved for 110kV GIS commercial energization.',
      'Log baseline gas pressure and ambient temperature in station logbook.',
      'Verify basement SF6 ambient leak detectors and low-level exhaust fans run continuously.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Trạm Biến Áp GIS 110kV TEV Urban Substation',
        equipment_tag: 'GIS-110KV-BAY-01',
        fse_name: 'Phan Văn Đức (duc.pv@tev.vn)',
        test_date: '2026-09-26',
      },
      electrical_tests: {
        sf6_pressure_mpa: 0.60,
        sf6_moisture_ppmv: 48,
        contact_resistance_micro_ohms: 28.5,
        breaker_open_time_ms: 22,
        breaker_close_time_ms: 48,
        pole_synchronization_ms: 0.8,
      },
    },
  },

  // 19. ELECTRIC VEHICLE SUPPLY EQUIPMENT EVSE (Sec 7.24)
  'neta-evse': {
    id: 'neta-evse',
    sectionCode: 'Mục 7.24',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.24 & SAE J1772',
    equipmentTitleVi: 'Trạm Sạc Xe Điện (Electric Vehicle Supply Equipment - EVSE)',
    equipmentTitleEn: 'Electric Vehicle Supply Equipment (DC Fast Charger & AC Level 2)',
    subtitleVi: 'Thị giác • Cách điện DC (≥ 100 MΩ) • Tín hiệu Pilot PWM • Ngắt rò GFCI (≤ 20mA / ≤ 20ms) • Nút dừng khẩn cấp',
    subtitleEn: 'Visual • DC Cable IR (≥ 100 MΩ) • Control Pilot PWM • GFCI Trip (≤ 20mA / ≤ 20ms) • Emergency Stop',
    category: 'renewable',
    defaultSiteInfo: {
      project_name: 'Trạm Sạc Nhanh Xe Điện TEV Supercharge Hub',
      equipment_tag: 'EVSE-DC-180KW-01',
      equipment_rating: '180 kW DC Fast Charger (CCS2 / CHAdeMO 150-1000V DC 250A)',
      fse_name: 'Trịnh Xuân Hùng (hung.tx@tev.vn)',
      test_date: '2026-09-25',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / TRẠM SẠC XE ĐIỆN ĐẠT TIÊU CHUẨN NETA ATS-2025 MỤC 7.24',
    statusTitleEn: 'PASS / EVSE COMPLIANT WITH NETA ATS-2025 SECTION 7.24 & SAE J1772',
    statusReasonVi: 'Cách điện cáp sạc DC và súng sạc tại 1000V DC đạt 3,200 MΩ (≥ 100 MΩ), tín hiệu điều khiển Control Pilot PWM chuyển tiếp chuẩn xác giữa các trạng thái State A -> B -> C, bảo vệ dòng rò chạm đất GFCI ngắt trong 14 ms tại dòng rò 15 mA (yêu cầu ≤ 20 mA và ≤ 20 ms), và nút dừng khẩn cấp E-Stop cắt nguồn tức thì.',
    statusReasonEn: 'DC charging cable insulation at 1000V DC is 3,200 MΩ (≥ 100 MΩ), Control Pilot PWM states transition cleanly from A through C, GFCI residual current device trips in 14 ms at 15 mA (≤ 20 mA and ≤ 20 ms standard), and E-Stop trips instantaneously.',
    evaluationItems: [
      {
        item: 'Bảo vệ Dòng rò Chạm đất GFCI / RCD (Ground-Fault Circuit Interrupter Test)',
        itemEn: 'Ground-Fault Circuit Interrupter (GFCI/RCD) Trip Response',
        measured: 'Dòng rò ngắt: 15 mA | Thời gian ngắt: 14 ms',
        standard: 'Bảo vệ chạm đất BẮT BUỘC cắt ngắt nguồn DC/AC tại dòng ≤ 20 mA trong thời gian ≤ 20 ms',
        standardEn: 'GFCI protection MUST disconnect power at ≤ 20 mA within ≤ 20 ms per UL 2231',
        reference: 'NETA ATS-2025 Sec 7.24.D.3 & UL 2231',
        status: 'PASS',
        detail: 'Cực kỳ an toàn cho người dùng khi cắm sạc trong điều kiện trời mưa ẩm ướt.',
        detailEn: 'Ensures absolute personnel safety even in wet outdoor charging environments.',
      },
      {
        item: 'Giao thức Tín hiệu Điều khiển Control Pilot PWM (SAE J1772 / IEC 61851)',
        itemEn: 'Control Pilot PWM Signal & State Transitions (States A, B, C, D)',
        measured: 'State A: 12.0V | State B: 9.0V | State C: 6.0V (Tần số 1000 Hz, Duty 50%)',
        standard: 'Biên độ điện áp Pilot nằm trong dung sai ±0.5V, chu kỳ xung PWM chính xác',
        standardEn: 'Pilot voltage amplitudes within ±0.5V, PWM duty cycle matched to rating',
        reference: 'SAE J1772 & IEC 61851-1',
        status: 'PASS',
        detail: 'Bộ giao tiếp nhận diện kết nối xe điện hoàn hảo, không có lỗi bắt tay handshake.',
        detailEn: 'Handshake protocol confirms vehicle connection before closing DC contactors.',
      },
      {
        item: 'Điện trở cách điện Cáp sạc DC & Súng sạc (1000V DC per Table 100.1)',
        itemEn: 'DC Charging Cable & Coupler Insulation Resistance @ 1000V DC',
        measured: 'DC+ to Ground: 3,200 MΩ | DC- to Ground: 3,400 MΩ',
        standard: 'Hệ thống cáp và đầu sạc DC BẮT BUỘC ≥ 100 MΩ per Table 100.1',
        standardEn: 'DC cable and coupler insulation MUST be ≥ 100 MΩ per Table 100.1',
        reference: 'NETA ATS-2025 Table 100.1',
        status: 'PASS',
        detail: 'Cáp sạc làm mát bằng chất lỏng (Liquid-Cooled Cable) cách điện tối ưu.',
        detailEn: 'Liquid-cooled high-current charging cable exhibits robust insulation.',
      },
      {
        item: 'Nút Dừng Khẩn Cấp & Khóa Cơ Khí Súng Sạc (Emergency Stop & Latch Interlock)',
        itemEn: 'Emergency Stop Pushbutton & Mechanical Latch Interlock',
        measured: 'Nhấn E-Stop: Ngắt nguồn DC hoàn toàn trong 18 ms; Khóa súng sạc giữ chắc chắn',
        standard: 'Ngắt nguồn DC tức thì, khóa súng sạc không cho phép rút súng khi đang có dòng',
        standardEn: 'Instant DC disconnection; motorized latch prevents uncoupling under load',
        reference: 'NETA ATS-2025 Sec 7.24.1',
        status: 'PASS',
        detail: 'Cơ cấu chốt khóa điện tử ngăn hồ quang điện khi người dùng vô tình kéo súng sạc.',
        detailEn: 'Motorized locking pin prevents arc flash hazards from premature pull-out.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'An Toàn Chống Rò Điện & Bảo Vệ GFCI',
        titleEn: 'GFCI Personnel Shock Hazard Protection',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Thời gian ngắt 14 ms triệt tiêu hoàn toàn nguy cơ điện giật cho người sử dụng.',
        ],
        pointsEn: [
          '14 ms trip completely eliminates human electrical shock hazards.',
        ],
      },
      {
        titleVi: 'Giao Thức Giao Tiếp Pilot & Quản Lý Dòng Sạc',
        titleEn: 'Pilot Signal Integrity & Dynamic Power Throttling',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Tín hiệu PWM 1000 Hz điều khiển dòng sạc mượt mà tương ứng với khả năng tiếp nhận của pin xe.',
        ],
        pointsEn: [
          '1000 Hz PWM modulates charging current dynamically to match battery BMS limits.',
        ],
      },
      {
        titleVi: 'Kiểm Soát Nhiệt Độ Súng Sạc (Connector Temperature)',
        titleEn: 'Connector Thermal Sensor Monitoring',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          'Cảm biến nhiệt độ tích hợp trong ngàm cắm súng sạc tự động giảm dòng nếu T > 85°C.',
        ],
        pointsEn: [
          'Integrated pin temperature sensors throttle current if connector exceeds 85°C.',
        ],
      },
      {
        titleVi: 'Tiếp Địa Vỏ Trạm Sạc & Cột Trụ Đẳng Thế',
        titleEn: 'EVSE Pedestal Grounding & Equipotential Bonding',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Vỏ kim loại trạm sạc nối đất với điện trở < 0.1 Ω về thanh cái tiếp địa nguồn.',
        ],
        pointsEn: [
          'Metal pedestal bonded to supply grounding bus with < 0.1 Ω resistance.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện kích hoạt dịch vụ sạc xe điện thương mại.',
      'Cập nhật firmware phiên bản mới nhất cho bộ điều khiển giao tiếp OCPP 1.6J/2.0.1.',
      'Kiểm tra định kỳ độ mài mòn ngàm cắm súng sạc và dây cáp sau 3 tháng.',
    ],
    recommendationsEn: [
      'Approved for public commercial EV fast charging deployment.',
      'Update OCPP 1.6J/2.0.1 cloud gateway controller firmware to latest build.',
      'Schedule quarterly inspection of cable jacket wear and connector contact pins.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Trạm Sạc Nhanh Xe Điện TEV Supercharge Hub',
        evse_tag: 'EVSE-DC-180KW-01',
        fse_name: 'Trịnh Xuân Hùng (hung.tx@tev.vn)',
        test_date: '2026-09-25',
      },
      electrical_tests: {
        gfci_trip_current_ma: 15,
        gfci_trip_time_ms: 14,
        dc_cable_insulation_megohms: 3200,
        control_pilot_frequency_hz: 1000,
        emergency_stop_response_ms: 18,
      },
    },
  },

  // 20. SOLAR PV SYSTEMS (Sec 7.25 & IEC 62446)
  'neta-pv': {
    id: 'neta-pv',
    sectionCode: 'Mục 7.25',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.25 & IEC 62446',
    equipmentTitleVi: 'Hệ Thống Điện Mặt Trời Nối Lưới (Utility-Scale Solar PV Systems)',
    equipmentTitleEn: 'Utility-Scale Solar PV Arrays & Central Inverters',
    subtitleVi: 'Thị giác • Điện áp Hở mạch Voc (±5%) • Dòng Ngắn mạch Isc • Cách điện DC (≥ 1.0 MΩ) • Mảng PV & Inverter',
    subtitleEn: 'Visual • Open-Circuit Voltage Voc (±5%) • Short-Circuit Current Isc • DC Insulation (≥ 1.0 MΩ) • Inverter',
    category: 'renewable',
    defaultSiteInfo: {
      project_name: 'Nhà Máy Điện Mặt Trời TEV Solar Farm 50MW',
      equipment_tag: 'PV-BLOCK-01-INV-01',
      equipment_rating: '3125 kVA 1500V DC Central Inverter & 3.5 MWp PV Array',
      fse_name: 'Võ Minh Tú (sgm1707@gmail.com)',
      test_date: '2026-09-24',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / HỆ THỐNG ĐIỆN MẶT TRỜI ĐẠT CHUẨN NETA ATS-2025 MỤC 7.25',
    statusTitleEn: 'PASS / SOLAR PV SYSTEM COMPLIANT WITH NETA ATS-2025 & IEC 62446',
    statusReasonVi: 'Điện áp hở mạch Voc đo trên toàn bộ 120 chuỗi string đạt 1385V ±1.2% (nằm trong dung sai ±5% theo thiết kế IEC 62446), dòng ngắn mạch Isc đồng đều theo bức xạ mặt trời, điện trở cách điện DC mảng pin đạt 48 MΩ (vượt xa ngưỡng tối thiểu ≥ 1.0 MΩ cho hệ 1500V), và biến tần hòa lưới đạt hiệu suất 98.9%.',
    statusReasonEn: 'Open-circuit voltage Voc across all 120 strings is 1385V ±1.2% (within ±5% IEC 62446 design band), short-circuit current Isc is uniform with irradiance, array DC insulation is 48 MΩ (≥ 1.0 MΩ limit for 1500V systems), and inverter efficiency is 98.9%.',
    evaluationItems: [
      {
        item: 'Điện áp hở mạch chuỗi chuỗi String Voc (Open-Circuit Voltage per IEC 62446)',
        itemEn: 'String Open-Circuit Voltage (Voc) Consistency Check',
        measured: '1382V đến 1388V (Trung bình 1385V, Sai số giữa các chuỗi max 0.4%)',
        standard: 'Điện áp Voc giữa các chuỗi song song BẮT BUỘC nằm trong dải dung sai ± 5.0%',
        standardEn: 'Voc variation between parallel strings MUST be within ± 5.0%',
        reference: 'IEC 62446-1 & NETA ATS Sec 7.25.D.1',
        status: 'PASS',
        detail: 'Tất cả các chuỗi pin đều đấu nối đủ số lượng 28 tấm/chuỗi, không bị sót tấm.',
        detailEn: 'All strings configured with identical module counts without missing panels.',
      },
      {
        item: 'Điện trở cách điện Mảng Pin DC đối với Đất (Array Insulation Resistance @ 1000V DC)',
        itemEn: 'PV Array Wet/Dry DC Insulation Resistance to Ground',
        measured: '48.0 MΩ (Hệ thống 1500V DC)',
        standard: 'Hệ thống điện mặt trời 1500V DC BẮT BUỘC ≥ 1.0 MΩ per IEC 62446',
        standardEn: '1500V DC solar systems MUST exceed ≥ 1.0 MΩ per IEC 62446-1',
        reference: 'IEC 62446-1 Category 1 & NETA ATS Sec 7.25.D.2',
        status: 'PASS',
        detail: 'Cáp DC ngầm và đầu nối MC4 khô ráo, không bị rò rỉ điện trường xuống khung giàn.',
        detailEn: 'Buried DC trunk lines and MC4 connectors exhibit superior ground isolation.',
      },
      {
        item: 'Dòng điện ngắn mạch chuỗi Isc (Short-Circuit Current per String)',
        itemEn: 'String Short-Circuit Current (Isc) Normalized to Irradiance',
        measured: '11.8 A tại bức xạ 920 W/m2 (Quy đổi STC: 12.8 A, chuẩn 12.9 A, lệch 0.8%)',
        standard: 'Dòng Isc đồng đều giữa các chuỗi (Độ lệch tối đa cho phép ≤ 5.0%)',
        standardEn: 'Isc deviation normalized to irradiance MUST be ≤ 5.0%',
        reference: 'IEC 62446-1 Category 1',
        status: 'PASS',
        detail: 'Không có hiện tượng che bóng (soiling/shading) hoặc hỏng diode suy hao (bypass diode).',
        detailEn: 'Zero shading or blown bypass diode anomalies across tested strings.',
      },
      {
        item: 'Kiểm tra Chức năng Đóng cắt Khẩn cấp & Chống Đảo Lưới (Anti-Islanding)',
        itemEn: 'Inverter Anti-Islanding Protection & Rapid Shutdown Response',
        measured: 'Mất lưới AC -> Biến tần ngắt hòa lưới trong 68 ms (Chuẩn IEEE 1547 ≤ 2.0s)',
        standard: 'Bảo vệ chống đảo lưới BẮT BUỘC cắt biến tần trong vòng ≤ 2.0 giây',
        standardEn: 'Anti-islanding MUST trip inverter within ≤ 2.0 seconds per IEEE 1547',
        reference: 'IEEE 1547 & NETA ATS Sec 7.25.D.4',
        status: 'PASS',
        detail: 'Bảo vệ tuyệt đối cho công nhân điện lực khi xảy ra mất điện trên lưới truyền tải.',
        detailEn: 'Guarantees absolute safety for utility lineworkers during grid power outages.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Điện Áp Chuỗi Pin & Cân Bằng Dòng Điện',
        titleEn: 'String Voltage Balance & Mismatch Losses',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Độ lệch Voc 0.4% khẳng định tổn hao do bất đối xứng (Mismatch Loss) ở mức cực tiểu.',
        ],
        pointsEn: [
          '0.4% Voc deviation minimizes inter-string mismatch losses to zero.',
        ],
      },
      {
        titleVi: 'Cách Điện Cáp DC & Đầu Nối MC4',
        titleEn: 'DC Wiring & MC4 Connector Reliability',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          'Cách điện 48 MΩ loại trừ nguy cơ sự cố chạm đất DC Ground Fault khi trời mưa.',
        ],
        pointsEn: [
          '48 MΩ insulation eliminates wet ground fault alarms during tropical downpours.',
        ],
      },
      {
        titleVi: 'Chống Đảo Lưới & Ổn Định Biến Tần Inverter',
        titleEn: 'Anti-Islanding & Inverter Grid Compliance',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Thời gian cắt 68 ms đáp ứng vượt trội quy chuẩn lưới điện quốc gia.',
        ],
        pointsEn: [
          '68 ms trip response surpasses IEEE 1547 utility interconnection requirements.',
        ],
      },
      {
        titleVi: 'Liên Kết Tiếp Địa Khung Giàn Pin Tracker',
        titleEn: 'Tracker Structure Grounding & Surge Arrestors (SPD)',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Khung giàn pin liên kết đẳng thế liên tục đạt < 0.2 Ω, chống sét lan truyền SPD loại 2 nguyên vẹn.',
        ],
        pointsEn: [
          'Tracker torque tube bonding resistance < 0.2 Ω; Type 2 DC SPDs fully operational.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện hòa lưới thương mại phát điện (COD - Commercial Operation Date).',
      'Thực hiện kiểm tra camera nhiệt máy bay không người lái (UAV Thermal Scan) để phát hiện điểm nóng Hotspot.',
      'Kích hoạt giám sát online chuỗi pin SCADA để theo dõi suy giảm hiệu suất Performance Ratio (PR).',
    ],
    recommendationsEn: [
      'Approved for commercial COD grid interconnection.',
      'Schedule drone thermography flight to survey for cell micro-cracks and hot spots.',
      'Engage SCADA string-level monitoring to track real-time Performance Ratio (PR).',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Nhà Máy Điện Mặt Trời TEV Solar Farm 50MW',
        inverter_tag: 'PV-BLOCK-01-INV-01',
        fse_name: 'Võ Minh Tú (sgm1707@gmail.com)',
        test_date: '2026-09-24',
      },
      electrical_tests: {
        voc_avg_volts: 1385,
        voc_spread_percent: 0.4,
        dc_insulation_resistance_megohms: 48.0,
        anti_islanding_trip_time_ms: 68,
        inverter_efficiency_percent: 98.9,
      },
    },
  },

  // 21. BATTERY ENERGY STORAGE SYSTEMS BESS (Sec 7.26 & NFPA 855)
  'neta-bess': {
    id: 'neta-bess',
    sectionCode: 'Mục 7.26',
    standardCode: 'ANSI/NETA ATS-2025 Sec 7.26 & NFPA 855',
    equipmentTitleVi: 'Hệ Thống Pin Lưu Trữ Năng Lượng (Battery Energy Storage Systems - BESS)',
    equipmentTitleEn: 'Battery Energy Storage Systems (Utility BESS & NFPA 855)',
    subtitleVi: 'Thị giác • Hệ thống Pin LFP • BMS Cell Voltage (±15mV) • Cách điện DC (≥ 50 MΩ) • Chữa cháy Khí xả Aerosol / Novec',
    subtitleEn: 'Visual • LFP Battery Rack • BMS Cell Delta (±15mV) • DC IR (≥ 50 MΩ) • Fire Suppression Interlocks',
    category: 'renewable',
    defaultSiteInfo: {
      project_name: 'Trạm Lưu Trữ Năng Lượng TEV MegaBESS 20MWh',
      equipment_tag: 'BESS-CONT-01-RACK-01',
      equipment_rating: '2.5 MW / 5.0 MWh Lithium Iron Phosphate (LFP) Containerized BESS',
      fse_name: 'Võ Minh Tú (sgm1707@gmail.com)',
      test_date: '2026-09-23',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / HỆ THỐNG BESS ĐẠT TIÊU CHUẨN NETA ATS-2025 & NFPA 855',
    statusTitleEn: 'PASS / BESS COMPLIANT WITH NETA ATS-2025 SECTION 7.26 & NFPA 855',
    statusReasonVi: 'Độ lệch điện áp giữa 280 cell pin LFP đo bởi hệ thống BMS chỉ 8 mV (ngưỡng tối đa cho phép ≤ 15 mV per cell), điện trở cách điện DC bus đạt 120 MΩ (≥ 50 MΩ), hệ thống kiểm soát nhiệt độ chất lỏng làm mát Liquid Cooling hoạt động chính xác ở 23.5°C, và liên động xả khí chữa cháy Novec 1230 ngắt contactor trong 15 ms.',
    statusReasonEn: 'Cell voltage spread across 280 LFP cells monitored by BMS is 8 mV (≤ 15 mV limit), DC bus insulation is 120 MΩ (≥ 50 MΩ), liquid cooling thermal management stabilizes at 23.5°C, and Novec 1230 fire interlocks trip DC contactors in 15 ms.',
    evaluationItems: [
      {
        item: 'Cân bằng điện áp cell pin LFP qua BMS (Cell Voltage Delta per BMS)',
        itemEn: 'BMS LFP Cell Voltage Delta & Balance Verification',
        measured: 'Điện áp cell: 3.325V đến 3.333V (Độ lệch cực đại 8 mV)',
        standard: 'Độ lệch điện áp giữa các cell pin LFP BẮT BUỘC ≤ 15 mV ở trạng thái tĩnh',
        standardEn: 'LFP cell-to-cell voltage spread MUST be ≤ 15 mV at rest',
        reference: 'NETA ATS-2025 Sec 7.26.D.2 & NFPA 855',
        status: 'PASS',
        detail: 'Mạch cân bằng chủ động Active Balancing hoạt động tối ưu, không có cell bị quá nạp/quá xả.',
        detailEn: 'Active cell balancing circuitry maintains tightly matched cell state-of-charge.',
      },
      {
        item: 'Điện trở cách điện DC Bus & Rack Pin (DC Isolation Resistance @ 1000V DC)',
        itemEn: 'High-Voltage DC Bus & Rack Isolation Resistance',
        measured: '120 MΩ (Hệ thống điện áp danh định 1200V DC)',
        standard: 'Điện trở cách điện DC bus BESS BẮT BUỘC ≥ 50 MΩ',
        standardEn: 'BESS DC bus insulation MUST exceed ≥ 50 MΩ',
        reference: 'NETA ATS-2025 Table 100.1 & Sec 7.26',
        status: 'PASS',
        detail: 'Hệ thống giám sát điện trở cách điện trực tuyến IMD (Insulation Monitoring Device) đọc chuẩn xác.',
        detailEn: 'Online Insulation Monitoring Device (IMD) reports high dielectric isolation.',
      },
      {
        item: 'Liên động Hệ thống PCCC & Phát hiện Khí cháy Deflagration (NFPA 855)',
        itemEn: 'Fire Suppression Interlocks & Off-Gas Detection (NFPA 855)',
        measured: 'Cảm biến khí H2/CO kích hoạt -> Cắt Contactor DC trong 15 ms -> Kích hoạt van Novec',
        standard: 'Bắt buộc ngắt toàn bộ contactor DC và kích hoạt hệ thống dập lửa khi phát hiện khí sinh',
        standardEn: 'Mandatory DC isolation and fire suppression trip upon off-gas detection',
        reference: 'NFPA 855 & UL 9540A',
        status: 'PASS',
        detail: 'Cảm biến phát hiện sớm khí sinh ra do thoát nhiệt (Off-gas thermal runaway) hoạt động 100%.',
        detailEn: 'Early off-gas thermal runaway detection triggers fast disconnect and gas flooding.',
      },
      {
        item: 'Hệ thống Làm mát Chất lỏng Liquid Cooling & Nhiệt độ Cell Pin',
        itemEn: 'Liquid Cooling Thermal Management & Temperature Uniformity',
        measured: 'Nhiệt độ cell: 23.2°C đến 24.1°C (Chênh lệch cực đại giữa các rack: 0.9°C)',
        standard: 'Chênh lệch nhiệt độ giữa các cell trong cụm pin BẮT BUỘC ≤ 3.0°C',
        standardEn: 'Temperature spread between battery cells MUST be ≤ 3.0°C',
        reference: 'NETA ATS-2025 Sec 7.26.1',
        status: 'PASS',
        detail: 'Dung dịch làm mát Glycol tuần hoàn đều, duy trì nhiệt độ tối ưu kéo dài tuổi thọ pin.',
        detailEn: 'Glycol liquid chiller circulating evenly, preventing hot spot degradation.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'An Toàn Chống Thoát Nhiệt (Thermal Runaway Mitigation)',
        titleEn: 'Thermal Runaway Prevention & Venting',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          'Công nghệ pin LFP Lithium Iron Phosphate có độ ổn định nhiệt vượt trội so với NMC.',
          'Van xả áp chống nổ container Deflagration Vent hoạt động tự do per NFPA 68.',
        ],
        pointsEn: [
          'LFP chemistry exhibits superior thermal stability compared to legacy NMC.',
          'Container deflagration panels operate freely compliant with NFPA 68.',
        ],
      },
      {
        titleVi: 'Hệ Thống BMS & Giám Sát Cân Bằng Cell',
        titleEn: 'BMS Cell Balance & Overcharge Protection',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Độ lệch 8 mV bảo đảm dung lượng lưu trữ khả dụng đạt 100% dung lượng thiết kế.',
        ],
        pointsEn: [
          '8 mV cell delta ensures 100% usable storage capacity without early cutoff.',
        ],
      },
      {
        titleVi: 'Khả Năng Cách Ly Contactor DC Nhanh',
        titleEn: 'High-Voltage DC Contactor Isolation Speed',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Contactor chân không DC ngắt dòng ngắn mạch trong 15 ms khi nhận lệnh khẩn cấp.',
        ],
        pointsEn: [
          'High-voltage DC contactors clear fault currents in 15 ms upon trip command.',
        ],
      },
      {
        titleVi: 'Hiệu Suất Vòng Sạc Xả Round-Trip Efficiency (RTE)',
        titleEn: 'Round-Trip Efficiency (RTE) Benchmark',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Thử nghiệm chu trình sạc/xả 1C đạt hiệu suất vòng lặp RTE 88.5% (chuẩn thiết kế > 86%).',
        ],
        pointsEn: [
          '1C charge/discharge cycle achieves 88.5% RTE, exceeding 86% benchmark.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện đóng điện hòa lưới phục vụ dịch vụ điều tần và tích trữ năng lượng.',
      'Kích hoạt giám sát cảnh báo sớm Cloud BESS Analytics để phát hiện cell tự xả bất thường.',
      'Kiểm tra định kỳ áp suất bình khí chữa cháy Novec 1230 và nồng độ dung dịch Glycol sau 6 tháng.',
    ],
    recommendationsEn: [
      'Approved for utility grid ancillary services and peak-shaving dispatch.',
      'Enable cloud-based AI analytics for early detection of abnormal self-discharge cells.',
      'Schedule 6-month check of Novec 1230 cylinder pressure and glycol concentration.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Trạm Lưu Trữ Năng Lượng TEV MegaBESS 20MWh',
        bess_tag: 'BESS-CONT-01-RACK-01',
        fse_name: 'Võ Minh Tú (sgm1707@gmail.com)',
        test_date: '2026-09-23',
      },
      electrical_tests: {
        cell_voltage_delta_mv: 8,
        dc_bus_insulation_megohms: 120,
        temp_spread_celsius: 0.9,
        e_stop_interlock_ms: 15,
        round_trip_efficiency_percent: 88.5,
      },
    },
  },

  // 22. PARTIAL DISCHARGE TESTING (Section 9 & IEEE 1434)
  'neta-pd': {
    id: 'neta-pd',
    sectionCode: 'Section 9',
    standardCode: 'ANSI/NETA ATS-2025 Section 9 & IEEE 1434',
    equipmentTitleVi: 'Thử Nghiệm Phóng Điện Cục Bộ Chuyên Sâu (Partial Discharge Testing - PD)',
    equipmentTitleEn: 'Partial Discharge Testing (Offline / Online PD Acoustics & TEV)',
    subtitleVi: 'Thị giác • Cảm biến TEV • Siêu âm Acoustic • Cảm biến dòng cao tần HFCT • Biểu đồ PRPD • Ngưỡng picoCoulomb',
    subtitleEn: 'Visual • TEV Sensors • Ultrasonic Acoustic • HFCT Sensors • Phase-Resolved PRPD • picoCoulomb (pC)',
    category: 'special',
    defaultSiteInfo: {
      project_name: 'Trạm Biến Áp 110kV TEV Industrial Substation',
      equipment_tag: 'PD-SURVEY-MV-HV-ASSETS',
      equipment_rating: '24kV Switchgear, 110kV Bushings & MV Power Cables',
      fse_name: 'Võ Minh Tú (sgm1707@gmail.com)',
      test_date: '2026-09-22',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / KHÔNG PHÁT HIỆN PHÓNG ĐIỆN CỤC BỘ BẤT THƯỜNG',
    statusTitleEn: 'PASS / NO HAZARDOUS PARTIAL DISCHARGE DETECTED (SECTION 9)',
    statusReasonVi: 'Khảo sát phóng điện cục bộ toàn trạm ghi nhận mức TEV < 6 dBmV (ngưỡng an toàn < 20 dBmV), tín hiệu siêu âm bề mặt 0 dBµV, cảm biến dòng cao tần HFCT đo dòng rò phóng điện < 50 pC (ngưỡng cho phép thiết bị mới < 100 pC theo IEC 60270), và đồ thị PRPD không có dạng hình sin tập trung của bọt khí hay phóng điện vầng quang.',
    statusReasonEn: 'Comprehensive PD scan recorded TEV < 6 dBmV (well below 20 dBmV safe limit), ultrasonic surface tracking at 0 dBµV, HFCT sensors measured < 50 pC (< 100 pC IEC 60270 limit), and Phase-Resolved PRPD exhibits clean baseline noise.',
    evaluationItems: [
      {
        item: 'Khảo sát Sóng điện áp Đất Quá độ TEV trên vỏ tủ (Transient Earth Voltage Test)',
        itemEn: 'Transient Earth Voltage (TEV) Switchgear Survey',
        measured: 'Tất cả các khoang tủ đo từ 2 đến 6 dBmV (Trung bình 4.2 dBmV)',
        standard: 'Mức TEV bình thường BẮT BUỘC < 20 dBmV (Ngưỡng cảnh báo nguy hiểm > 29 dBmV)',
        standardEn: 'Normal TEV level MUST be < 20 dBmV (Critical hazard threshold > 29 dBmV)',
        reference: 'NETA ATS-2025 Section 9 & IEC 62478',
        status: 'PASS',
        detail: 'Không có hiện tượng phóng điện bên trong lòng các sứ đỡ epoxy hay thanh cái bọc cách điện.',
        detailEn: 'Zero internal void discharges within cast epoxy bus supports.',
      },
      {
        item: 'Khảo sát Âm học Siêu âm Bề mặt (Airborne Ultrasonic Acoustic Scan @ 40 kHz)',
        itemEn: 'Airborne Ultrasonic Acoustic Surface Discharge Screening',
        measured: '0 dBµV (Không có âm thanh phóng điện vầng quang rít gió)',
        standard: 'Tín hiệu siêu âm BẮT BUỘC < 6 dBµV (Ngưỡng phóng điện bề mặt > 10 dBµV)',
        standardEn: 'Ultrasonic signal MUST be < 6 dBµV; zero surface tracking acoustic noise',
        reference: 'NETA ATS-2025 Section 9',
        status: 'PASS',
        detail: 'Cách điện sứ và đầu cáp co nguội hoàn toàn sạch sẽ, không có phóng điện tia lửa bề mặt.',
        detailEn: 'Clean insulator surfaces without surface tracking or corona discharge.',
      },
      {
        item: 'Đo Phóng điện cục bộ bằng Biến dòng Cao tần HFCT (High-Frequency Current Transformer)',
        itemEn: 'HFCT Cable Shield High-Frequency Discharge Pulse Measurement',
        measured: 'Cường độ phóng điện cực đại: 35 pC (picoCoulombs)',
        standard: 'Thiết bị mới đưa vào nghiệm thu BẮT BUỘC < 100 pC per IEC 60270',
        standardEn: 'New equipment acceptance threshold MUST be < 100 pC per IEC 60270',
        reference: 'IEC 60270 & IEEE Std 1434',
        status: 'PASS',
        detail: 'Độ bền điện trường cách điện cáp lực và đầu hộp nối cáp đạt mức hoàn hảo.',
        detailEn: 'Cable terminations and joint bodies maintain void-free dielectric.',
      },
      {
        item: 'Phân tích Đồ thị Pha Phóng điện PRPD (Phase-Resolved Partial Discharge Pattern)',
        itemEn: 'Phase-Resolved Partial Discharge (PRPD) Cluster Pattern Analysis',
        measured: 'Đồ thị nhiễu nền phân tán ngẫu nhiên, không tập trung ở góc pha 45°/225°',
        standard: 'Không xuất hiện các cụm mẫu phóng điện điển hình của bọt khí (Void), vầng quang (Corona), hay vết bẩn',
        standardEn: 'No phase-synchronized discharge clusters characteristic of internal voids',
        reference: 'IEEE Std 1434 & IEC 62478',
        status: 'PASS',
        detail: 'Xác nhận toàn bộ thiết bị trung và cao thế không có bất kỳ khuyết tật sản xuất nào.',
        detailEn: 'Confirms zero manufacturing or installation insulation flaws.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Khảo Sát TEV Vỏ Tủ & Khoang Cáp',
        titleEn: 'TEV Amplitude & Enclosure Resonance',
        category: 'insulation',
        severity: 'normal',
        pointsVi: [
          'Mức TEV 4.2 dBmV thấp hơn ngưỡng cảnh báo 20 dBmV đến 5 lần.',
        ],
        pointsEn: [
          '4.2 dBmV is 5 times lower than the international 20 dBmV action threshold.',
        ],
      },
      {
        titleVi: 'Sóng Âm Siêu Âm & Phóng Điện Vầng Quang Bề Mặt',
        titleEn: 'Acoustic Tracking & Surface Contamination',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          'Không có âm thanh siêu âm chứng minh khoang tủ khô ráo, không bị ẩm ngưng tụ.',
        ],
        pointsEn: [
          'Zero acoustic noise verifies dry interior without moisture condensation.',
        ],
      },
      {
        titleVi: 'Biên Độ Điện Tích Phóng Điện pC (HFCT Cable Shield)',
        titleEn: 'Discharge Apparent Charge pC Amplitude',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Biên độ 35 pC hoàn toàn nằm trong dải nhiễu nền điện từ EMI môi trường công nghiệp.',
        ],
        pointsEn: [
          '35 pC level matches typical benign electromagnetic ambient background noise.',
        ],
      },
      {
        titleVi: 'Phân Tích Cụm Mẫu PRPD & Nhận Diện Khuyết Tật',
        titleEn: 'PRPD Fingerprint & AI Defect Classification',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Thuật toán AI phân loại tự động kết luận: Cách điện đạt chuẩn 100% (No Defects).',
        ],
        pointsEn: [
          'AI pattern recognition classifies pattern as 100% healthy (No Defects).',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện nghiệm thu đóng điện vận hành toàn bộ các ngăn lộ thiết bị.',
      'Lưu biểu đồ PRPD ban đầu làm cơ sở dữ liệu đối soát Fingerprint cho các kỳ bảo dưỡng định kỳ.',
      'Tái kiểm tra đo PD sau 12 tháng vận hành đầy tải để đánh giá tốc độ già hóa cách điện.',
    ],
    recommendationsEn: [
      'Approved for live operational acceptance without insulation reservations.',
      'Store baseline PRPD fingerprint files in asset health registry.',
      'Resurvey partial discharge after 12 months under full load to monitor thermal aging.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Trạm Biến Áp 110kV TEV Industrial Substation',
        survey_tag: 'PD-SURVEY-MV-HV-ASSETS',
        fse_name: 'Võ Minh Tú (sgm1707@gmail.com)',
        test_date: '2026-09-22',
      },
      electrical_tests: {
        max_tev_dbmv: 6.0,
        ultrasonic_dbuv: 0.0,
        hfct_max_pc: 35,
        prpd_defect_detected: false,
      },
    },
  },

  // 23. INFRARED THERMOGRAPHIC SURVEY (Section 9 & Table 100.18)
  'neta-thermography': {
    id: 'neta-thermography',
    sectionCode: 'Section 9',
    standardCode: 'ANSI/NETA ATS-2025 Section 9 & Table 100.18',
    equipmentTitleVi: 'Khảo Sát Nhiệt Hồng Ngoại Chuyên Nghiệp (Infrared Thermographic Survey)',
    equipmentTitleEn: 'Infrared Thermographic Survey (Electrical Connections & Table 100.18)',
    subtitleVi: 'Thị giác • Camera nhiệt Fluke/FLIR • Độ lệch nhiệt Delta-T (Table 100.18) • Mối nối bu-lông • Tiếp điểm dao cách ly',
    subtitleEn: 'Visual • Calibrated Thermal Camera • Delta-T Criteria (Table 100.18) • Bolted Joints • Disconnect Contacts',
    category: 'special',
    defaultSiteInfo: {
      project_name: 'Nhà Máy Hóa Chất TEV Chemical Plant',
      equipment_tag: 'THERMO-SURVEY-PLANT-WIDE',
      equipment_rating: '110kV Switchyard, 24kV Switchgear & 400V MSB Distribution Panels',
      fse_name: 'Võ Minh Tú (sgm1707@gmail.com)',
      test_date: '2026-09-21',
    },
    overallStatus: 'PASS',
    statusTitleVi: 'PASS / KHẢO SÁT NHIỆT ĐẠT TIÊU CHUẨN NETA ATS-2025 BẢNG 100.18',
    statusTitleEn: 'PASS / THERMOGRAPHIC SURVEY COMPLIANT WITH NETA ATS TABLE 100.18',
    statusReasonVi: 'Toàn bộ 320 mối nối bu-lông, đầu cosse cáp và tiếp điểm máy cắt được quét dưới tải > 60% công suất định mức. Chênh lệch nhiệt độ lớn nhất giữa các pha tương tự ΔT chỉ 1.8°C (nằm trong vùng an toàn Priority 4: ΔT ≤ 3°C theo Bảng Table 100.18), không phát hiện bất kỳ điểm phát nhiệt cục bộ bất thường nào.',
    statusReasonEn: 'All 320 bolted connections, lugs, and breaker contacts were scanned under > 60% rated load. Maximum temperature difference between similar components ΔT was only 1.8°C (safe Priority 4: ΔT ≤ 3°C per Table 100.18), with zero abnormal hot spots.',
    evaluationItems: [
      {
        item: 'Chênh lệch nhiệt độ giữa các cấu phần tương tự (Temperature Difference ΔT per Table 100.18)',
        itemEn: 'Delta-T Between Similar Components Under Similar Loading',
        measured: 'Điểm nóng nhất: Pha B đạt 38.6°C | Pha A đạt 36.8°C (ΔT = 1.8°C)',
        standard: 'ΔT ≤ 3°C: Tình trạng bình thường (Priority 4 per Table 100.18)',
        standardEn: 'ΔT ≤ 3°C: Normal condition / Possible deficiency (Priority 4)',
        reference: 'NETA ATS-2025 Table 100.18',
        status: 'PASS',
        detail: 'Nhiệt độ giữa 3 pha hoàn toàn đồng đều, lực siết và tiếp xúc dẫn dòng tối ưu.',
        detailEn: 'Three-phase thermal profiles are balanced with zero contact resistance anomalies.',
      },
      {
        item: 'Chênh lệch nhiệt độ so với Môi trường xung quanh (Temperature Difference Over Ambient)',
        itemEn: 'Temperature Difference Over Ambient Air Temperature',
        measured: 'Nhiệt độ môi trường: 32.0°C | Nhiệt độ đầu cosse lớn nhất: 38.6°C (ΔT = 6.6°C)',
        standard: 'ΔT so với môi trường ≤ 10°C: Tình trạng bình thường (Priority 4 per Table 100.18)',
        standardEn: 'ΔT over ambient ≤ 10°C: Normal condition (Priority 4 Table 100.18)',
        reference: 'NETA ATS-2025 Table 100.18',
        status: 'PASS',
        detail: 'Độ gia nhiệt chỉ 6.6°C, thấp hơn rất nhiều so với giới hạn cảnh báo 20°C.',
        detailEn: 'Temperature rise of only 6.6°C over ambient is well within safe operating margins.',
      },
      {
        item: 'Tỷ lệ mang tải khi chụp ảnh nhiệt (Percentage Load of Rated Capacity)',
        itemEn: 'Electrical Load Percentage During Infrared Inspection',
        measured: 'Tải thực tế khi quét: 68% công suất định mức',
        standard: 'Khảo sát nhiệt BẮT BUỘC thực hiện khi phụ tải đạt tối thiểu ≥ 40% định mức (Khuyến nghị ≥ 50%)',
        standardEn: 'Thermographic survey MUST be performed with minimum ≥ 40% rated load (≥ 50% recommended)',
        reference: 'NETA ATS-2025 Section 9 & NFPA 70B',
        status: 'PASS',
        detail: 'Đủ điều kiện mang tải để bộc lộ rõ ràng các khuyết tật tiếp xúc lỏng nếu có.',
        detailEn: '68% load exceeds the mandatory 40% threshold to uncover high resistance joints.',
      },
      {
        item: 'Cài đặt Hiệu chỉnh Hệ số Bức xạ & Khoảng cách (Emissivity & Distance Calibration)',
        itemEn: 'Emissivity & Reflected Temperature Calibration',
        measured: 'Emissivity: 0.90 (Đầu cosse đồng bọc co nhiệt) | Khoảng cách: 1.5 m | Tref: 32°C',
        standard: 'Hệ số bức xạ hiệu chỉnh chính xác theo loại vật liệu bề mặt kim loại hoặc vỏ bọc',
        standardEn: 'Emissivity adjusted accurately according to material surface emissivity table',
        reference: 'ISO 18434-1 & NETA Section 9',
        status: 'PASS',
        detail: 'Dữ liệu ảnh nhiệt radiometric chụp bởi camera chuyên dụng đã hiệu chỉnh chuẩn xác.',
        detailEn: 'Radiometric thermal images calibrated with accurate background reflection inputs.',
      },
    ],
    anomalyCards: [
      {
        titleVi: 'Phân Loại Mức Độ Khẩn Cấp Bảng Table 100.18',
        titleEn: 'NETA Table 100.18 Action Priority Classification',
        category: 'thermal',
        severity: 'normal',
        pointsVi: [
          'ΔT = 1.8°C xếp loại Priority 4 (Bình thường / Tiếp tục theo dõi).',
          'Không có điểm nào rơi vào Priority 2 (ΔT 4-15°C) hay Priority 1 (ΔT > 15°C - Nguy hiểm tức thì).',
        ],
        pointsEn: [
          'ΔT = 1.8°C classified as Priority 4 (Normal / Monitor at next routine cycle).',
          'Zero points in Priority 2 (ΔT 4-15°C) or Priority 1 (ΔT > 15°C critical).',
        ],
      },
      {
        titleVi: 'Chất Lượng Tiếp Xúc Mối Nối Thanh Cái & Đầu Cosse',
        titleEn: 'Busbar & Cable Lug Contact Quality',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Mối nối thanh cái đạt nhiệt độ tản nhiệt đồng đều không có vạch phát nhiệt bất thường.',
        ],
        pointsEn: [
          'Busbar bolted joints display completely uniform thermal gradients.',
        ],
      },
      {
        titleVi: 'Tiếp Điểm Dao Cách Ly & Tiếp Điểm Cực Máy Cắt',
        titleEn: 'Disconnect Switch Jaws & Breaker Finger Clusters',
        category: 'contact',
        severity: 'normal',
        pointsVi: [
          'Tiếp điểm ngàm kẹp dao cách ly 110kV dẫn dòng đối xứng hoàn hảo trên cả 3 pha.',
        ],
        pointsEn: [
          '110kV outdoor disconnect switch jaws carry balanced current symmetrically.',
        ],
      },
      {
        titleVi: 'Hồ Sơ Nhiệt Radiometric & Lưu Trữ Lịch Sử CMMS',
        titleEn: 'Radiometric Image Archival & CMMS History',
        category: 'operational',
        severity: 'normal',
        pointsVi: [
          'Toàn bộ ảnh nhiệt kèm ảnh thị giác được lưu trữ radiometric vào phần mềm CMMS để đối soát năm sau.',
        ],
        pointsEn: [
          'Radiometric infrared and visual paired images archived in CMMS asset records.',
        ],
      },
    ],
    recommendationsVi: [
      'Đủ điều kiện cấp chứng chỉ hoàn thành nghiệm thu nhiệt toàn trạm.',
      'Lập lịch khảo sát nhiệt định kỳ 6 tháng một lần theo tiêu chuẩn bảo dưỡng NETA MTS.',
      'Cung cấp báo cáo ảnh nhiệt radiometric có chữ ký chứng nhận của chuyên gia nhiệt độ Level II.',
    ],
    recommendationsEn: [
      'Approved for full commercial acceptance with clean thermographic clearance.',
      'Schedule routine thermographic surveys semi-annually pursuant to NETA MTS.',
      'Issue certified radiometric thermal report signed by certified Level II Thermographer.',
    ],
    samplePayloadJson: {
      site_info: {
        project_name: 'Nhà Máy Hóa Chất TEV Chemical Plant',
        survey_tag: 'THERMO-SURVEY-PLANT-WIDE',
        fse_name: 'Võ Minh Tú (sgm1707@gmail.com)',
        test_date: '2026-09-21',
      },
      electrical_tests: {
        max_delta_t_similar_celsius: 1.8,
        max_delta_t_ambient_celsius: 6.6,
        load_percentage: 68,
        emissivity_setting: 0.90,
        hotspots_detected: 0,
      },
    },
  },
};

/**
 * Helper function to generate clean 4-part markdown report for any of the 23 ATS specifications
 */
export function generateAtsMarkdownReport(spec: NetaAtsSpecificationDef, lang: 'vi' | 'en' = 'vi'): string {
  const isVi = lang === 'vi';
  const site = spec.defaultSiteInfo;

  let md = isVi
    ? `# BÁO CÁO THẨM ĐỊNH CHUYÊN GIA NETA ATS-2025 (ACCEPTANCE TESTING)
**THIẾT BỊ: ${spec.equipmentTitleVi.toUpperCase()}**
*Tiêu chuẩn đối soát: ${spec.standardCode} & ANSI/NETA ATS-2025*
*Mã biên bản: TEV-ATS-${site.equipment_tag}-${site.test_date.replace(/-/g, '')}*

---

## 1. TRẠNG THÁI TỔNG QUAN (Overall Status)
### **${spec.overallStatus === 'PASS' ? '🟢' : spec.overallStatus === 'INVESTIGATE' ? '🟡' : '🔴'} ${spec.statusTitleVi}**
- **Mã thiết bị (Tag):** \`${site.equipment_tag}\`
- **Công suất / Cấp điện áp:** ${site.equipment_rating}
- **Dự án / Nhà máy:** ${site.project_name}
- **Kỹ sư khảo sát FSE:** ${site.fse_name} | **Ngày kiểm định:** ${site.test_date}
- **Kết luận thẩm định:** ${spec.statusReasonVi}

---

## 2. BẢNG PHÂN TÍCH CHI TIẾT THEO TIÊU CHUẨN ANSI/NETA ATS-2025
| Hạng mục kiểm tra & thử nghiệm | Giá trị đo thực tế | Tiêu chuẩn NETA ATS-2025 | Tham chiếu kỹ thuật | Trạng thái |
| :--- | :--- | :--- | :--- | :---: |
`
    : `# NETA ATS-2025 ACCEPTANCE TESTING SPECIFICATION REPORT
**EQUIPMENT: ${spec.equipmentTitleEn.toUpperCase()}**
*Reference Standards: ${spec.standardCode} & ANSI/NETA ATS-2025*
*Report ID: TEV-ATS-${site.equipment_tag}-${site.test_date.replace(/-/g, '')}*

---

## 1. OVERALL STATUS
### **${spec.overallStatus === 'PASS' ? '🟢' : spec.overallStatus === 'INVESTIGATE' ? '🟡' : '🔴'} ${spec.statusTitleEn}**
- **Equipment Tag:** \`${site.equipment_tag}\`
- **Rating / Voltage Class:** ${site.equipment_rating}
- **Facility / Substation:** ${site.project_name}
- **Field Engineer (FSE):** ${site.fse_name} | **Test Date:** ${site.test_date}
- **Executive Conclusion:** ${spec.statusReasonEn}

---

## 2. DETAILED EVALUATION TABLE (ANSI/NETA ATS-2025)
| Test / Inspection Item | Measured Site Value | NETA ATS-2025 Standard | Reference Standard | Status |
| :--- | :--- | :--- | :--- | :---: |
`;

  spec.evaluationItems.forEach((item) => {
    const itemLabel = isVi ? item.item : item.itemEn;
    const stdLabel = isVi ? item.standard : item.standardEn;
    const statusIcon = item.status === 'PASS' ? 'PASS' : item.status === 'INVESTIGATE' ? 'INVESTIGATE' : 'FAIL';
    md += `| ${itemLabel} | ${item.measured} | ${stdLabel} | ${item.reference} | ${statusIcon} |\n`;
  });

  md += isVi
    ? `\n---

## 3. PHÂN TÍCH CHUYÊN SÂU & RỦI RO BẤT THƯỜNG (In-Depth Anomaly & Risk Analysis)
`
    : `\n---

## 3. IN-DEPTH ANOMALY & RISK DIAGNOSTICS
`;

  spec.anomalyCards.forEach((card, idx) => {
    const title = isVi ? card.titleVi : card.titleEn;
    const points = isVi ? card.pointsVi : card.pointsEn;
    md += `### 3.${idx + 1}. ${title}\n`;
    points.forEach((pt) => {
      md += `- ${pt}\n`;
    });
    md += `\n`;
  });

  md += isVi
    ? `---

## 4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)
`
    : `---

## 4. ACTIONABLE RECOMMENDATIONS FOR FIELD SERVICE ENGINEERS
`;

  const recs = isVi ? spec.recommendationsVi : spec.recommendationsEn;
  recs.forEach((rec, idx) => {
    md += `${idx + 1}. ${rec}\n`;
  });

  md += isVi
    ? `\n---
*Báo cáo được khởi tạo tự động bởi TEV Platform AI • Field Service Evaluation Suite.*`
    : `\n---
*Report automatically generated by TEV Platform AI • Field Service Evaluation Suite.*`;

  return md;
}

export function getAllAtsSpecsList(): NetaAtsSpecificationDef[] {
  return Object.values(NETA_ATS_23_MASTER_SPECS);
}
