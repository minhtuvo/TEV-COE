import { 
  DEFAULT_CUSTOMERS, 
  DEFAULT_EQUIPMENT, 
  DEFAULT_WORK_ORDERS, 
  DEFAULT_INVENTORY,
  CustomerRecord,
  EquipmentMasterRecord,
  FieldInspectionReport
} from './fieldSyncService';

export {
  DEFAULT_CUSTOMERS,
  DEFAULT_EQUIPMENT,
  DEFAULT_WORK_ORDERS,
  DEFAULT_INVENTORY
};

/**
 * Historical inspection reports powering simulation charts, trending graphs,
 * DGA triangles, and 2-way Google Sheets synchronization.
 */
export const DEFAULT_SIMULATION_REPORTS: any[] = [
  // --- MÁY BIẾN ÁP TRF-110-01 (Multiple Inspection Cycles for Trending) ---
  {
    id: 'REP-2026-09-001',
    date: '15/09/2026',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    factory: 'Trạm biến áp TEV 110kV',
    location: 'Sân phân phối 110kV - Vịnh T1',
    equipmentId: 'TRF-110-01',
    equipmentName: 'Máy biến áp chính T1 110/22kV 63MVA',
    type: 'Máy biến áp',
    inspector: 'KS Nguyễn Văn An (FSE Lead)',
    status: 'healthy',
    health: 94,
    notes: 'Kỳ kiểm tra Q3/2026: DGA khí hòa tan ở mức bình thường theo IEEE C57.104. Dầu trong, không phát hiện bọt khí.',
    measurements: {
      oilTemp: '52.0',
      windingTemp: '61.5',
      irHighLow: '3200',
      irHighEarth: '3100',
      h2: '15',
      ch4: '12',
      c2h6: '8',
      c2h4: '10',
      c2h2: '0',
      co: '185',
      co2: '1680',
      o2: '480',
      n2: '42000',
      dielectricStrength: '68',
      furan: '0.04',
      oilMoisture: '12'
    },
    rawData: ['15/09/2026', 'Công ty Cổ phần Năng lượng TEV', 'Trạm biến áp TEV 110kV', 'Sân phân phối 110kV', 'TRF-110-01', 'Máy biến áp chính T1', 'Máy biến áp', 94, 'healthy', '52.0', '61.5', '3200', '3100', 'normal', '15', '68', '0.04', '12', '4', '0.75', 'https://drive.google.com/reports/TRF-110-01_202609.pdf']
  },
  {
    id: 'REP-2026-08-001',
    date: '15/08/2026',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    factory: 'Trạm biến áp TEV 110kV',
    location: 'Sân phân phối 110kV - Vịnh T1',
    equipmentId: 'TRF-110-01',
    equipmentName: 'Máy biến áp chính T1 110/22kV 63MVA',
    type: 'Máy biến áp',
    inspector: 'KS Trần Minh Hùng',
    status: 'healthy',
    health: 93,
    notes: 'Kỳ kiểm tra tháng 8: Nhiệt độ dầu 50.5°C, các thông số cách điện đạt tiêu chuẩn.',
    measurements: {
      oilTemp: '50.5',
      windingTemp: '59.0',
      irHighLow: '3150',
      irHighEarth: '3050',
      h2: '14',
      ch4: '11',
      c2h6: '7',
      c2h4: '9',
      c2h2: '0',
      co: '175',
      co2: '1620',
      o2: '490',
      n2: '41500',
      dielectricStrength: '66',
      furan: '0.04',
      oilMoisture: '13'
    }
  },
  {
    id: 'REP-2026-07-001',
    date: '15/07/2026',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    factory: 'Trạm biến áp TEV 110kV',
    location: 'Sân phân phối 110kV - Vịnh T1',
    equipmentId: 'TRF-110-01',
    equipmentName: 'Máy biến áp chính T1 110/22kV 63MVA',
    type: 'Máy biến áp',
    inspector: 'KS Nguyễn Văn An',
    status: 'healthy',
    health: 92,
    notes: 'Kỳ kiểm tra tháng 7: Vận hành bình thường, tải 72%.',
    measurements: {
      oilTemp: '48.0',
      windingTemp: '57.0',
      irHighLow: '3100',
      irHighEarth: '2980',
      h2: '12',
      ch4: '10',
      c2h6: '6',
      c2h4: '8',
      c2h2: '0',
      co: '160',
      co2: '1550',
      o2: '510',
      n2: '43000',
      dielectricStrength: '65',
      furan: '0.03',
      oilMoisture: '14'
    }
  },
  {
    id: 'REP-2026-06-001',
    date: '15/06/2026',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    factory: 'Trạm biến áp TEV 110kV',
    location: 'Sân phân phối 110kV - Vịnh T1',
    equipmentId: 'TRF-110-01',
    equipmentName: 'Máy biến áp chính T1 110/22kV 63MVA',
    type: 'Máy biến áp',
    inspector: 'KS Lê Hoàng Long',
    status: 'healthy',
    health: 91,
    notes: 'Kỳ kiểm tra định kỳ 6 tháng đầu năm: Baseline kiểm định hoàn tất.',
    measurements: {
      oilTemp: '47.5',
      windingTemp: '56.0',
      irHighLow: '3050',
      irHighEarth: '2900',
      h2: '10',
      ch4: '9',
      c2h6: '5',
      c2h4: '7',
      c2h2: '0',
      co: '150',
      co2: '1490',
      o2: '520',
      n2: '44000',
      dielectricStrength: '64',
      furan: '0.03',
      oilMoisture: '14'
    }
  },

  // --- MÁY BIẾN ÁP TRF-01 (Mô phỏng Dashboard mặc định) ---
  {
    id: 'REP-2026-09-TRF01',
    date: '16/09/2026',
    customer: 'Công ty Điện lực A',
    factory: 'Nhà máy Bắc Ninh',
    location: 'Trạm biến áp 110kV',
    equipmentId: 'TRF-01',
    equipmentName: 'Máy biến áp T1',
    type: 'Máy biến áp',
    inspector: 'FSE Engineer',
    status: 'healthy',
    health: 95,
    notes: 'Báo cáo kiểm tra định kỳ MBA T1. Cách điện tốt, DGA đạt mức chuẩn.',
    measurements: {
      oilTemp: '55.0',
      windingTemp: '64.0',
      irHighLow: '3500',
      irHighEarth: '3400',
      h2: '18',
      ch4: '14',
      c2h6: '9',
      c2h4: '11',
      c2h2: '0',
      co: '190',
      co2: '1720',
      o2: '470',
      n2: '41000',
      dielectricStrength: '70',
      furan: '0.03',
      oilMoisture: '11'
    }
  },
  {
    id: 'REP-2026-08-TRF01',
    date: '16/08/2026',
    customer: 'Công ty Điện lực A',
    factory: 'Nhà máy Bắc Ninh',
    location: 'Trạm biến áp 110kV',
    equipmentId: 'TRF-01',
    equipmentName: 'Máy biến áp T1',
    type: 'Máy biến áp',
    inspector: 'FSE Engineer',
    status: 'healthy',
    health: 94,
    notes: 'Theo dõi thông số nhiệt độ và DGA định kỳ tháng 8.',
    measurements: {
      oilTemp: '53.0',
      windingTemp: '62.0',
      irHighLow: '3400',
      irHighEarth: '3300',
      h2: '16',
      ch4: '13',
      c2h6: '8',
      c2h4: '10',
      c2h2: '0',
      co: '180',
      co2: '1650',
      o2: '480',
      n2: '42000',
      dielectricStrength: '68',
      furan: '0.03',
      oilMoisture: '12'
    }
  },
  {
    id: 'REP-2026-07-TRF01',
    date: '16/07/2026',
    customer: 'Công ty Điện lực A',
    factory: 'Nhà máy Bắc Ninh',
    location: 'Trạm biến áp 110kV',
    equipmentId: 'TRF-01',
    equipmentName: 'Máy biến áp T1',
    type: 'Máy biến áp',
    inspector: 'FSE Engineer',
    status: 'healthy',
    health: 93,
    notes: 'Theo dõi thông số nhiệt độ và DGA định kỳ tháng 7.',
    measurements: {
      oilTemp: '51.5',
      windingTemp: '60.0',
      irHighLow: '3300',
      irHighEarth: '3200',
      h2: '15',
      ch4: '12',
      c2h6: '7',
      c2h4: '9',
      c2h2: '0',
      co: '170',
      co2: '1580',
      o2: '490',
      n2: '43000',
      dielectricStrength: '67',
      furan: '0.02',
      oilMoisture: '13'
    }
  },

  // --- TỦ ĐIỆN SWG-MV-24KV-01 & SWG-01 (TEV & Ultrasonic Partial Discharge) ---
  {
    id: 'REP-2026-09-002',
    date: '14/09/2026',
    customer: 'Tập đoàn Thép Hòa Phát Dung Quất',
    factory: 'Khu liên hợp Gang Thép Dung Quất',
    location: 'Trạm điện phân phối Xưởng Cán Thép',
    equipmentId: 'SWG-MV-24KV-01',
    equipmentName: 'Tủ hợp bộ đóng cắt trung thế 24kV Metal-Clad 8 ngăn',
    type: 'Tủ điện',
    inspector: 'KS Trần Đình Tuấn',
    status: 'healthy',
    health: 93,
    notes: 'Đo phóng điện cục bộ TEV và Siêu âm (Ultrasonic): TEV max 12 dBmV (ngưỡng an toàn < 20 dBmV), Siêu âm 8 dBµV.',
    measurements: {
      tev: '12.0',
      ultrasonic: '8.0',
      maxBusbarTemp: '48.5',
      contactRes: '16.5'
    }
  },
  {
    id: 'REP-2026-08-002',
    date: '14/08/2026',
    customer: 'Tập đoàn Thép Hòa Phát Dung Quất',
    factory: 'Khu liên hợp Gang Thép Dung Quất',
    location: 'Trạm điện phân phối Xưởng Cán Thép',
    equipmentId: 'SWG-MV-24KV-01',
    equipmentName: 'Tủ hợp bộ đóng cắt trung thế 24kV Metal-Clad 8 ngăn',
    type: 'Tủ điện',
    inspector: 'KS Trần Đình Tuấn',
    status: 'healthy',
    health: 92,
    notes: 'Kỳ tháng 8: TEV 10.5 dBmV, Siêu âm 7.5 dBµV, thanh cái 47.0°C.',
    measurements: {
      tev: '10.5',
      ultrasonic: '7.5',
      maxBusbarTemp: '47.0',
      contactRes: '16.8'
    }
  },
  {
    id: 'REP-2026-07-002',
    date: '14/07/2026',
    customer: 'Tập đoàn Thép Hòa Phát Dung Quất',
    factory: 'Khu liên hợp Gang Thép Dung Quất',
    location: 'Trạm điện phân phối Xưởng Cán Thép',
    equipmentId: 'SWG-MV-24KV-01',
    equipmentName: 'Tủ hợp bộ đóng cắt trung thế 24kV Metal-Clad 8 ngăn',
    type: 'Tủ điện',
    inspector: 'KS Trần Đình Tuấn',
    status: 'healthy',
    health: 92,
    notes: 'Kỳ tháng 7: TEV 9.0 dBmV, Siêu âm 6.8 dBµV, thanh cái 46.2°C.',
    measurements: {
      tev: '9.0',
      ultrasonic: '6.8',
      maxBusbarTemp: '46.2',
      contactRes: '17.0'
    }
  },

  // --- ĐỘNG CƠ MTR-PUMP-450KW (3-Phase Vibration, Temp & Tan-Delta) ---
  {
    id: 'REP-2026-09-003',
    date: '12/09/2026',
    customer: 'Tổ hợp Sản xuất Ô tô VinFast',
    factory: 'Nhà máy Cát Hải - Hải Phòng',
    location: 'Trạm Bơm Làm Mát Trung Tâm - Bơm P1',
    equipmentId: 'MTR-PUMP-450KW',
    equipmentName: 'Động cơ bơm tuần hoàn nước giải nhiệt 450kW',
    type: 'Động cơ',
    inspector: 'KS Lê Hoàng Long',
    status: 'healthy',
    health: 91,
    notes: 'Độ rung RMS DE 1.85 mm/s (tiêu chuẩn ISO 10816-3 Zone A). Nhiệt độ Stator 72.5°C, Vòng bi 63.0°C.',
    measurements: {
      vibration: '1.85',
      statorTemp: '72.5',
      bearingTemp: '63.0',
      voltageImbalance: '0.45',
      irR: '850',
      piR: '3.4',
      ddR: '1.8',
      pdR: '3200',
      elcidR: '85',
      contactRes: '14.2'
    }
  },
  {
    id: 'REP-2026-08-003',
    date: '12/08/2026',
    customer: 'Tổ hợp Sản xuất Ô tô VinFast',
    factory: 'Nhà máy Cát Hải - Hải Phòng',
    location: 'Trạm Bơm Làm Mát Trung Tâm - Bơm P1',
    equipmentId: 'MTR-PUMP-450KW',
    equipmentName: 'Động cơ bơm tuần hoàn nước giải nhiệt 450kW',
    type: 'Động cơ',
    inspector: 'KS Lê Hoàng Long',
    status: 'healthy',
    health: 90,
    notes: 'Độ rung RMS DE 1.80 mm/s. Nhiệt độ Stator 71.0°C.',
    measurements: {
      vibration: '1.80',
      statorTemp: '71.0',
      bearingTemp: '62.0',
      voltageImbalance: '0.40',
      irR: '840',
      piR: '3.3',
      ddR: '1.9',
      pdR: '3300',
      elcidR: '88',
      contactRes: '14.5'
    }
  },
  {
    id: 'REP-2026-07-003',
    date: '12/07/2026',
    customer: 'Tổ hợp Sản xuất Ô tô VinFast',
    factory: 'Nhà máy Cát Hải - Hải Phòng',
    location: 'Trạm Bơm Làm Mát Trung Tâm - Bơm P1',
    equipmentId: 'MTR-PUMP-450KW',
    equipmentName: 'Động cơ bơm tuần hoàn nước giải nhiệt 450kW',
    type: 'Động cơ',
    inspector: 'KS Lê Hoàng Long',
    status: 'healthy',
    health: 90,
    notes: 'Độ rung RMS DE 1.75 mm/s. Nhiệt độ Stator 69.5°C.',
    measurements: {
      vibration: '1.75',
      statorTemp: '69.5',
      bearingTemp: '60.5',
      voltageImbalance: '0.42',
      irR: '830',
      piR: '3.2',
      ddR: '2.0',
      pdR: '3400',
      elcidR: '90',
      contactRes: '14.8'
    }
  },

  // --- MÁY CẮT MC-ACB-3200A ---
  {
    id: 'REP-2026-09-004',
    date: '20/09/2026',
    customer: 'Samsung Electronics Vietnam (SEVT)',
    factory: 'Nhà máy SEVT Thái Nguyên',
    location: 'Phòng MSB 01 - Ngăn lộ tổng 01',
    equipmentId: 'MC-ACB-3200A',
    equipmentName: 'Máy cắt không khí tổng ACB Incomer 3200A',
    type: 'Máy cắt',
    inspector: 'KS Nguyễn Văn An',
    status: 'healthy',
    health: 98,
    notes: 'Điện trở tiếp xúc 3 pha: A=18.2 µΩ, B=18.5 µΩ, C=18.4 µΩ. Thời gian đóng 48ms, cắt 32ms.',
    measurements: {
      contactResPhaseA: '18.2',
      contactResPhaseB: '18.5',
      contactResPhaseC: '18.4',
      insulationOpenPoles: '2400',
      insulationToGround: '2100',
      closingTimeMs: '48',
      openingTimeMs: '32'
    }
  },

  // --- MÁY PHÁT ĐIỆN GEN-STANDBY-2000KVA ---
  {
    id: 'REP-2026-09-005',
    date: '18/09/2026',
    customer: 'Samsung Electronics Vietnam (SEVT)',
    factory: 'Nhà máy SEVT Thái Nguyên',
    location: 'Nhà máy phát điện dự phòng - Máy số 1',
    equipmentId: 'GEN-STANDBY-2000KVA',
    equipmentName: 'Máy phát điện dự phòng khẩn cấp Cummins 2000kVA',
    type: 'Máy phát',
    inspector: 'KS Trần Minh Hùng',
    status: 'healthy',
    health: 97,
    notes: 'Thử tải chuyển nguồn ATS tự động: Thời gian nhận tải 8.5 giây, tần số 50.05 Hz ổn định, áp 401V.',
    measurements: {
      insulationStator: '1250',
      insulationRotor: '680',
      voltageL1L2: '401',
      frequencyHz: '50.05',
      oilPressureBar: '4.8',
      coolantTempC: '82',
      batteryVoltageV: '26.8',
      atsTransferTimeSec: '8.5'
    }
  }
];

/**
 * Ensures initial default data contains all master equipment if empty.
 */
export function getInitialEquipmentList(storedList?: any[]): any[] {
  if (Array.isArray(storedList) && storedList.length > 0) {
    // Merge any missing default equipment so the user always has rich items
    const existingIds = new Set(storedList.map(e => e.id?.toLowerCase()));
    const missingDefaults = DEFAULT_EQUIPMENT.filter(e => !existingIds.has(e.id.toLowerCase()));
    return [...storedList, ...missingDefaults];
  }
  return DEFAULT_EQUIPMENT;
}

/**
 * Ensures initial reports contain simulation trend data if empty.
 */
export function getInitialReportsList(storedReports?: any[]): any[] {
  if (Array.isArray(storedReports) && storedReports.length > 0) {
    const existingIds = new Set(storedReports.map(r => r.id));
    const missingDefaults = DEFAULT_SIMULATION_REPORTS.filter(r => !existingIds.has(r.id));
    return [...storedReports, ...missingDefaults];
  }
  return DEFAULT_SIMULATION_REPORTS;
}
