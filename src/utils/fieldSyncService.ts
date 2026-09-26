/**
 * fieldSyncService.ts
 * Enterprise Data Architecture & 2-Way Synchronization Engine for TEV Service Platform
 * 
 * Supports:
 * - Master Customer Directory (Sheet 'khachhang')
 * - Master Equipment Directory (Sheet 'thietbi')
 * - Dynamic Technical Schemas for:
 *   + Máy biến áp (Transformer)
 *   + Máy cắt (Circuit Breaker - ACB / VCB / SF6)
 *   + Động cơ (Electric Motor)
 *   + Máy phát (Engine Generator)
 *   + Tủ điện (Switchgear / RMU / MSB)
 * - Field Inspection Records (Sheet 'TEV Service Flatform')
 * - Real-time Bidirectional (2-way) Sync between App & Google Sheets
 */

export type EquipmentCategoryType = 
  | 'Máy biến áp'
  | 'Máy cắt'
  | 'Động cơ'
  | 'Máy phát'
  | 'Tủ điện'
  | 'Inverter';

export interface CustomerRecord {
  id: string;              // Mã khách hàng (VD: KH-SAMSUNG, KH-EVN-01)
  name: string;            // Tên khách hàng (VD: Công ty Điện lực A, Samsung SEVT)
  factories: string[];     // Danh sách nhà máy / phân xưởng
  email?: string;
  phone?: string;
  address?: string;
  contactPerson?: string;
  notes?: string;
  createdAt?: string;
  updatedAt?: string;
}

// Technical Nameplate Specifications per Equipment Type
export interface TransformerSpecs {
  ratedPowerKva?: number;       // Công suất định mức (kVA/MVA)
  primaryVoltageKv?: number;     // Điện áp sơ cấp (kV, VD: 22kV, 110kV)
  secondaryVoltageV?: number;    // Điện áp thứ cấp (V, VD: 400V)
  ratedCurrentA?: number;        // Dòng định mức (A)
  vectorGroup?: string;          // Tổ đấu dây (Dyn11, YNd11...)
  coolingType?: string;          // Làm mát (ONAN, ONAF, AN, AF)
  insulationType?: string;       // Dầu khoáng (Mineral Oil), Silicon, Cast Resin (Khô)
  impedanceZ?: number;           // Uk% (5.5%, 6.0%)
  oilWeightKg?: number;          // Khối lượng dầu
}

export interface BreakerSpecs {
  ratedVoltageKv?: number;       // Cấp điện áp (kV: 0.4kV, 24kV, 36kV, 110kV)
  ratedCurrentA?: number;        // Dòng định mức (A: 630A, 1600A, 3200A, 4000A)
  breakingCapacityKa?: number;   // Dòng cắt ngắn mạch Icu (kA: 25kA, 50kA, 65kA)
  breakerType?: string;          // ACB (Không khí), VCB (Chân không), SF6
  operatingMechanism?: string;   // Động cơ lò xo (Motor Spring), Điện từ, Thủy lực
  tripUnitModel?: string;        // MicroLogic, Ekip, PR122, Sepam...
  polesCount?: number;           // 3 cực, 4 cực
}

export interface MotorSpecs {
  ratedPowerKw?: number;         // Công suất (kW/HP)
  ratedVoltageV?: number;        // Điện áp định mức (380V, 660V, 6600V)
  ratedCurrentA?: number;        // Dòng định mức (A)
  ratedSpeedRpm?: number;        // Tốc độ quay (vòng/phút, VD: 1450, 2950 rpm)
  poleCount?: number;            // Số cực (2, 4, 6, 8)
  insulationClass?: string;      // Cấp cách điện (Class F, H, B)
  powerFactorCosPhi?: number;    // Cos phi (0.85, 0.88)
  ipRating?: string;             // Chuẩn bảo vệ (IP55, IP65)
  bearingTypeDe?: string;        // Loại vòng bi DE / NDE (6314 C3...)
}

export interface GeneratorSpecs {
  ratedPowerKva?: number;        // Công suất định mức (kVA/kW)
  ratedVoltageV?: number;        // Điện áp phát (400V, 6300V, 10500V)
  ratedCurrentA?: number;        // Dòng điện định mức (A)
  frequencyHz?: number;          // Tần số (50Hz / 60Hz)
  powerFactorCosPhi?: number;    // Hệ số công suất (0.8)
  engineManufacturer?: string;   // Hãng động cơ (Cummins, Perkins, Mitsubishi, MTU)
  alternatorBrand?: string;      // Hãng đầu phát (Stamford, Leroy Somer, Mecc Alte)
  excitationSystem?: string;     // Kích từ (AVR Brushless, PMG, AREP)
  coolingSystem?: string;        // Làm mát két nước, trao đổi nhiệt
}

export interface SwitchgearSpecs {
  ratedVoltageKv?: number;       // Cấp điện áp danh định (0.4kV, 24kV, 36kV)
  busbarRatingA?: number;        // Dòng thanh cái chính Busbar (1250A, 2500A, 4000A)
  shortCircuitKa?: number;       // Khả năng chịu ngắn mạch (25kA/1s, 50kA/1s)
  cubicleCount?: number;         // Số ngăn tủ (Ngăn lộ vào, phân đoạn, ngăn lộ ra)
  formType?: string;             // Form tủ (Form 2b, Form 3b, Form 4b)
  ipRating?: string;             // Chuẩn bảo vệ vỏ (IP31, IP42, IP54)
  enclosureType?: string;        // Metal-Clad, Metal-Enclosed, RMU
}

export interface EquipmentMasterRecord {
  id: string;                    // Mã thiết bị (Tag ID: TR-01, ACB-MSB-01, MTR-PUMP-01...)
  name: string;                  // Tên thiết bị
  customer: string;              // Tên khách hàng (liên kết sheet khachhang)
  factory: string;               // Nhà máy / Site
  location: string;              // Vị trí / Trạm / Ngăn lộ
  type: EquipmentCategoryType;   // Phân loại thiết bị
  manufacturer?: string;         // Hãng sản xuất (ABB, Schneider, Siemens, Dong Anh, Thibidi...)
  model?: string;                // Model / Kiểu
  serialNumber?: string;         // Số chế tạo / Serial No
  yearOfManufacture?: string;    // Năm sản xuất / Đưa vào vận hành
  criticality?: 'A' | 'B' | 'C'; // Mức độ quan trọng rủi ro
  status?: 'healthy' | 'warning' | 'critical'; // Trạng thái hiện tại
  health?: number;               // Điểm sức khỏe HI (0 - 100%)
  lastCheck?: string;            // Ngày kiểm tra gần nhất (DD/MM/YYYY)
  pmCycleMonths?: number;        // Chu kỳ bảo trì định kỳ (tháng: 3, 6, 12)
  notes?: string;
  source?: 'sheet' | 'app' | 'ocr'; // Nguồn nhập
  
  // Specific Technical Specs based on equipment type
  specs?: {
    transformer?: TransformerSpecs;
    breaker?: BreakerSpecs;
    motor?: MotorSpecs;
    generator?: GeneratorSpecs;
    switchgear?: SwitchgearSpecs;
    rawText?: string;
  };

  // Specific Test Measurements (lưu lại kết quả kiểm tra gần nhất)
  measurements?: Record<string, any>;
  rawData?: any[];
}

export interface FieldInspectionReport {
  id: string;                    // Mã báo cáo kiểm tra (REP-YYYY-MM-XXXX)
  date: string;                  // Ngày kiểm tra
  customer: string;              // Khách hàng
  factory: string;               // Nhà máy / Site
  location: string;              // Vị trí
  equipmentId: string;           // Mã thiết bị
  equipmentName: string;         // Tên thiết bị
  type: EquipmentCategoryType;   // Loại thiết bị
  inspector: string;             // Kỹ sư FSE
  status: 'healthy' | 'warning' | 'critical';
  health: number;                // Chỉ số HI %
  notes: string;                 // Nhận xét / Khuyến nghị
  measurements: Record<string, any>;
  fileUrl?: string;              // Link biên bản PDF / Ảnh chụp nameplate
  createdAt?: string;
}

// ============================================================================
// DEFAULT MASTER DATA FOR SEEDING & PRE-MIGRATION BACKUP
// ============================================================================

export const DEFAULT_CUSTOMERS: CustomerRecord[] = [
  {
    id: 'KH-TEV-01',
    name: 'Công ty Cổ phần Năng lượng TEV',
    factories: ['Nhà máy Bắc Ninh', 'Nhà máy Hải Phòng', 'Trạm biến áp TEV 110kV'],
    email: 'contact@tev-energy.vn',
    phone: '024.3988.7766',
    address: 'KCN Quế Võ, Tỉnh Bắc Ninh',
    contactPerson: 'Kỹ sư Trưởng Nguyễn Văn An',
    notes: 'Khách hàng trọng điểm - Hệ thống 110kV & 22kV'
  },
  {
    id: 'KH-SEVT-02',
    name: 'Samsung Electronics Vietnam (SEVT)',
    factories: ['Nhà máy SEVT Thái Nguyên', 'Phân xưởng Phụ trợ Utility 1', 'Substation 110kV Phổ Yên'],
    email: 'facilities.sevt@samsung.com',
    phone: '0208.356.9999',
    address: 'KCN Yên Bình, Phổ Yên, Thái Nguyên',
    contactPerson: 'Trưởng phòng Điện Mr. Park / Mr. Hùng',
    notes: 'Yêu cầu kiểm tra NETA ATS định kỳ 6 tháng/lần'
  },
  {
    id: 'KH-VINFAST-03',
    name: 'Tổ hợp Sản xuất Ô tô VinFast',
    factories: ['Nhà máy Cát Hải - Hải Phòng', 'Xưởng Hàn Thân Vỏ', 'Trạm Biến áp 110kV Đình Vũ'],
    email: 'maintenance@vinfast.vn',
    phone: '0225.388.9988',
    address: 'Đảo Cát Hải, Huyện Cát Hải, Hải Phòng',
    contactPerson: 'KS Lê Hoàng Long',
    notes: 'Hệ thống tủ điện trung thế 24kV RMU & Robot Motors'
  },
  {
    id: 'KH-HOAPHAT-04',
    name: 'Tập đoàn Thép Hòa Phát Dung Quất',
    factories: ['Khu liên hợp Gang Thép Dung Quất', 'Xưởng Cán Thép', 'Trạm 220kV Hòa Phát'],
    email: 'dienluc.dungquat@hoaphat.com.vn',
    phone: '0255.362.8888',
    address: 'KKT Dung Quất, Huyện Bình Sơn, Quảng Ngãi',
    contactPerson: 'Mr. Trần Đình Tuấn',
    notes: 'Môi trường nhiệt cao, bụi dẫn điện - Kiểm tra IR & TEV thường xuyên'
  }
];

export const DEFAULT_EQUIPMENT: EquipmentMasterRecord[] = [
  // --- MÁY BIẾN ÁP ---
  {
    id: 'TRF-110-01',
    name: 'Máy biến áp chính T1 110/22kV 63MVA',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    factory: 'Trạm biến áp TEV 110kV',
    location: 'Sân phân phối 110kV - Vịnh T1',
    type: 'Máy biến áp',
    manufacturer: 'Đông Anh EEMC',
    model: 'TBA-63MVA-110',
    serialNumber: 'EEMC-2022-9981',
    yearOfManufacture: '2022',
    criticality: 'A',
    status: 'healthy',
    health: 94,
    lastCheck: '15/09/2026',
    pmCycleMonths: 6,
    specs: {
      transformer: {
        ratedPowerKva: 63000,
        primaryVoltageKv: 110,
        secondaryVoltageV: 22000,
        ratedCurrentA: 330,
        vectorGroup: 'YNd11',
        coolingType: 'ONAF',
        insulationType: 'Mineral Oil (Dầu khoáng TCVN)',
        impedanceZ: 10.5,
        oilWeightKg: 18500
      }
    },
    measurements: {
      oilTemp: '52', windingTemp: '61', irHighLow: '3200', irHighEarth: '3100',
      dielectricStrength: '68', furan: '0.04', oilMoisture: '12',
      h2: '15', ch4: '12', c2h4: '8', c2h2: '0', co: '180', co2: '1650'
    }
  },
  {
    id: 'TRF-SEVT-02',
    name: 'Máy biến áp phân phối khô Cast Resin 2500kVA',
    customer: 'Samsung Electronics Vietnam (SEVT)',
    factory: 'Phân xưởng Phụ trợ Utility 1',
    location: 'Phòng MBA Khô Tầng 1 - Cell TR02',
    type: 'Máy biến áp',
    manufacturer: 'ABB Transformer',
    model: 'Resibloc 2500kVA 22/0.4kV',
    serialNumber: 'ABB-VN-2023-4589',
    yearOfManufacture: '2023',
    criticality: 'A',
    status: 'healthy',
    health: 96,
    lastCheck: '10/08/2026',
    pmCycleMonths: 12,
    specs: {
      transformer: {
        ratedPowerKva: 2500,
        primaryVoltageKv: 22,
        secondaryVoltageV: 400,
        ratedCurrentA: 3608,
        vectorGroup: 'Dyn11',
        coolingType: 'AN/AF',
        insulationType: 'Cast Resin (Cuộn đúc chân không)',
        impedanceZ: 6.0
      }
    },
    measurements: {
      windingTemp: '75', irHighLow: '4500', irHighEarth: '4200'
    }
  },

  // --- MÁY CẮT (CIRCUIT BREAKER) ---
  {
    id: 'MC-ACB-3200A',
    name: 'Máy cắt không khí tổng ACB Incomer 3200A',
    customer: 'Samsung Electronics Vietnam (SEVT)',
    factory: 'Nhà máy SEVT Thái Nguyên',
    location: 'Phòng MSB 01 - Ngăn lộ tổng 01',
    type: 'Máy cắt',
    manufacturer: 'Schneider Electric',
    model: 'MasterPact NW32 H1 3P Drawout',
    serialNumber: 'SCH-2024-8841',
    yearOfManufacture: '2024',
    criticality: 'A',
    status: 'healthy',
    health: 98,
    lastCheck: '20/09/2026',
    pmCycleMonths: 6,
    specs: {
      breaker: {
        ratedVoltageKv: 0.4,
        ratedCurrentA: 3200,
        breakingCapacityKa: 65,
        breakerType: 'ACB (Air Circuit Breaker)',
        operatingMechanism: 'Motor Lò xo trữ năng 220VAC',
        tripUnitModel: 'MicroLogic 6.0A Ground Fault',
        polesCount: 3
      }
    },
    measurements: {
      contactResPhaseA: '18.2', contactResPhaseB: '18.5', contactResPhaseC: '18.4',
      insulationOpenPoles: '2400', insulationToGround: '2100',
      closingTimeMs: '48', openingTimeMs: '32'
    }
  },
  {
    id: 'MC-VCB-24KV',
    name: 'Máy cắt chân không VCB 24kV 1250A 25kA',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    factory: 'Nhà máy Bắc Ninh',
    location: 'Tủ trung thế RMU Ngăn Feeder F02',
    type: 'Máy cắt',
    manufacturer: 'Siemens',
    model: 'SION 3AE 24kV 1250A',
    serialNumber: 'SIE-2023-1120',
    yearOfManufacture: '2023',
    criticality: 'B',
    status: 'healthy',
    health: 92,
    lastCheck: '05/09/2026',
    pmCycleMonths: 12,
    specs: {
      breaker: {
        ratedVoltageKv: 24,
        ratedCurrentA: 1250,
        breakingCapacityKa: 25,
        breakerType: 'VCB (Vacuum Circuit Breaker)',
        operatingMechanism: 'Motor Lò xo 110VDC',
        tripUnitModel: 'Rơ le bảo vệ SIPROTEC 5 7SJ82',
        polesCount: 3
      }
    },
    measurements: {
      contactResPhaseA: '28.5', contactResPhaseB: '29.1', contactResPhaseC: '28.8',
      vacuumIntegrity: 'Pass (25kV AC/1min)', insulationClosedToGround: '4800'
    }
  },

  // --- ĐỘNG CƠ ĐIỆN (ELECTRIC MOTOR) ---
  {
    id: 'MTR-PUMP-450KW',
    name: 'Động cơ bơm tuần hoàn nước giải nhiệt 450kW',
    customer: 'Tổ hợp Sản xuất Ô tô VinFast',
    factory: 'Nhà máy Cát Hải - Hải Phòng',
    location: 'Trạm Bơm Làm Mát Trung Tâm - Bơm P1',
    type: 'Động cơ',
    manufacturer: 'Siemens / SIMOTICS GP',
    model: '1LA8 357-4AB60 450kW 380V',
    serialNumber: 'MTR-SIE-2023-7721',
    yearOfManufacture: '2023',
    criticality: 'A',
    status: 'healthy',
    health: 91,
    lastCheck: '12/09/2026',
    pmCycleMonths: 3,
    specs: {
      motor: {
        ratedPowerKw: 450,
        ratedVoltageV: 380,
        ratedCurrentA: 820,
        ratedSpeedRpm: 1485,
        poleCount: 4,
        insulationClass: 'Class F (155°C)',
        powerFactorCosPhi: 0.88,
        ipRating: 'IP55',
        bearingTypeDe: 'SKF 6320 C3'
      }
    },
    measurements: {
      vibrationDeRms: '1.85', vibrationNdeRms: '1.62',
      statorTemp: '72.5', bearingTempDe: '63.0', bearingTempNde: '58.5',
      irStatorMegohms: '850', polarizationIndexPi: '3.4', voltageImbalance: '0.45'
    }
  },

  // --- MÁY PHÁT ĐIỆN (ENGINE GENERATOR) ---
  {
    id: 'GEN-STANDBY-2000KVA',
    name: 'Máy phát điện dự phòng khẩn cấp Cummins 2000kVA',
    customer: 'Samsung Electronics Vietnam (SEVT)',
    factory: 'Nhà máy SEVT Thái Nguyên',
    location: 'Nhà máy phát điện dự phòng - Máy số 1',
    type: 'Máy phát',
    manufacturer: 'Cummins Power Generation',
    model: 'C2200 D5 2000kVA 400V 50Hz',
    serialNumber: 'CUM-2023-55241',
    yearOfManufacture: '2023',
    criticality: 'A',
    status: 'healthy',
    health: 97,
    lastCheck: '18/09/2026',
    pmCycleMonths: 3,
    specs: {
      generator: {
        ratedPowerKva: 2000,
        ratedVoltageV: 400,
        ratedCurrentA: 2886,
        frequencyHz: 50,
        powerFactorCosPhi: 0.8,
        engineManufacturer: 'Cummins QSK60-G4 V16 Turbocharged',
        alternatorBrand: 'Stamford PI734F1 Brushless',
        excitationSystem: 'PMG + MX321 AVR',
        coolingSystem: 'Két nước giải nhiệt gắn liền 50°C'
      }
    },
    measurements: {
      insulationStator: '1250', insulationRotor: '680',
      voltageL1L2: '401', frequencyHz: '50.05', oilPressureBar: '4.8',
      coolantTempC: '82', batteryVoltageV: '26.8', atsTransferTimeSec: '8.5'
    }
  },

  // --- TỦ ĐIỆN PHÂN PHỐI & ĐÓNG CẮT (SWITCHGEAR) ---
  {
    id: 'SWG-MV-24KV-01',
    name: 'Tủ hợp bộ đóng cắt trung thế 24kV Metal-Clad 8 ngăn',
    customer: 'Tập đoàn Thép Hòa Phát Dung Quất',
    factory: 'Khu liên hợp Gang Thép Dung Quất',
    location: 'Trạm điện phân phối Xưởng Cán Thép',
    type: 'Tủ điện',
    manufacturer: 'Schneider Electric',
    model: 'PIX 24kV 1250A 25kA/1s Form 4b',
    serialNumber: 'SCH-PIX-2024-1011',
    yearOfManufacture: '2024',
    criticality: 'A',
    status: 'healthy',
    health: 93,
    lastCheck: '14/09/2026',
    pmCycleMonths: 6,
    specs: {
      switchgear: {
        ratedVoltageKv: 24,
        busbarRatingA: 1250,
        shortCircuitKa: 25,
        cubicleCount: 8,
        formType: 'Form 4b',
        ipRating: 'IP4X',
        enclosureType: 'Metal-Clad LSC2B-PM'
      }
    },
    measurements: {
      tevDbMax: '12', ultrasonicDbuv: '8', maxBusbarTemp: '48.5',
      contactResBusbar: '16.5', controlWiringIr: '38.0'
    }
  }
];

// ============================================================================
// CONVERTERS & FORMATTERS FOR GOOGLE SHEETS
// ============================================================================

/**
 * Creates sheet row for sheet 'khachhang'
 */
export function customerToSheetRow(c: CustomerRecord): any[] {
  return [
    c.id,
    c.name,
    Array.isArray(c.factories) ? c.factories.join(', ') : (c.factories || ''),
    c.email || '',
    c.phone || '',
    c.address || '',
    c.contactPerson || '',
    c.notes || '',
    c.createdAt || new Date().toISOString(),
    c.updatedAt || new Date().toISOString()
  ];
}

/**
 * Creates sheet row for sheet 'thietbi' (Master Asset Table)
 */
export function equipmentToSheetRow(e: EquipmentMasterRecord): any[] {
  // Format summary of technical specs into readable string
  let specsSummary = '';
  if (e.type === 'Máy biến áp' && e.specs?.transformer) {
    const t = e.specs.transformer;
    specsSummary = `${t.ratedPowerKva || ''} kVA | ${t.primaryVoltageKv || ''}/${(t.secondaryVoltageV || 400) / 1000} kV | ${t.vectorGroup || 'Dyn11'} | ${t.insulationType || ''}`;
  } else if (e.type === 'Máy cắt' && e.specs?.breaker) {
    const b = e.specs.breaker;
    specsSummary = `${b.ratedCurrentA || ''}A | ${b.breakingCapacityKa || ''}kA | ${b.ratedVoltageKv || ''}kV | ${b.breakerType || ''}`;
  } else if (e.type === 'Động cơ' && e.specs?.motor) {
    const m = e.specs.motor;
    specsSummary = `${m.ratedPowerKw || ''} kW | ${m.ratedVoltageV || ''}V | ${m.ratedSpeedRpm || ''} rpm | ${m.insulationClass || ''}`;
  } else if (e.type === 'Máy phát' && e.specs?.generator) {
    const g = e.specs.generator;
    specsSummary = `${g.ratedPowerKva || ''} kVA | ${g.ratedVoltageV || ''}V | ${g.engineManufacturer || ''}`;
  } else if (e.type === 'Tủ điện' && e.specs?.switchgear) {
    const s = e.specs.switchgear;
    specsSummary = `${s.ratedVoltageKv || ''}kV | Busbar ${s.busbarRatingA || ''}A | ${s.cubicleCount || ''} Ngăn | ${s.formType || ''}`;
  } else {
    specsSummary = e.notes || '';
  }

  return [
    e.id,
    e.name,
    e.customer,
    e.factory,
    e.location,
    e.type,
    e.manufacturer || '',
    e.model || '',
    e.serialNumber || '',
    e.yearOfManufacture || '',
    specsSummary,
    e.status || 'healthy',
    e.health !== undefined ? e.health : 90,
    e.criticality || 'B',
    e.lastCheck || new Date().toLocaleDateString('vi-VN'),
    e.pmCycleMonths || 6
  ];
}

/**
 * Creates sheet row for 'TEV Service Flatform' (Inspection Log)
 */
export function inspectionReportToSheetRow(r: FieldInspectionReport): any[] {
  return [
    r.id,
    r.date,
    r.customer,
    r.factory,
    r.location,
    r.equipmentId,
    r.equipmentName,
    r.type,
    r.inspector || 'FSE Engineer',
    r.status === 'healthy' ? 'Đạt (Healthy)' : r.status === 'warning' ? 'Cảnh báo (Warning)' : 'Nguy hiểm (Critical)',
    r.health,
    r.notes || '',
    r.fileUrl || ''
  ];
}

/**
 * Maps raw sheet row from 'khachhang' back to CustomerRecord
 */
export function sheetRowToCustomer(row: any[]): CustomerRecord | null {
  if (!row || row.length < 2) return null;
  const id = row[0]?.toString().trim();
  const name = row[1]?.toString().trim();
  if (!id || !name || id.toLowerCase().includes('mã khách')) return null;

  const factoriesRaw = row[2] ? row[2].toString() : '';
  const factories = factoriesRaw.split(',').map((f: string) => f.trim()).filter((f: string) => f.length > 0);

  return {
    id,
    name,
    factories: factories.length > 0 ? factories : ['Nhà máy chính'],
    email: row[3]?.toString().trim() || '',
    phone: row[4]?.toString().trim() || '',
    address: row[5]?.toString().trim() || '',
    contactPerson: row[6]?.toString().trim() || '',
    notes: row[7]?.toString().trim() || '',
    createdAt: row[8]?.toString().trim() || new Date().toISOString(),
    updatedAt: row[9]?.toString().trim() || new Date().toISOString()
  };
}

/**
 * Maps raw sheet row from 'thietbi' back to EquipmentMasterRecord
 */
export function sheetRowToEquipment(row: any[]): EquipmentMasterRecord | null {
  if (!row || row.length < 3) return null;
  const id = row[0]?.toString().trim();
  const name = row[1]?.toString().trim();
  if (!id || !name || id.toLowerCase().includes('mã thiết bị') || id.toLowerCase().includes('tag')) return null;

  const customer = row[2]?.toString().trim() || '';
  const factory = row[3]?.toString().trim() || '';
  const location = row[4]?.toString().trim() || '';
  const typeStr = row[5]?.toString().trim() || 'Máy biến áp';

  let type: EquipmentCategoryType = 'Máy biến áp';
  const lower = typeStr.toLowerCase();
  if (lower.includes('cắt') || lower.includes('breaker') || lower.includes('acb') || lower.includes('vcb')) {
    type = 'Máy cắt';
  } else if (lower.includes('động cơ') || lower.includes('motor') || lower.includes('bơm')) {
    type = 'Động cơ';
  } else if (lower.includes('phát') || lower.includes('generator') || lower.includes('gen')) {
    type = 'Máy phát';
  } else if (lower.includes('tủ') || lower.includes('switchgear') || lower.includes('rmu') || lower.includes('msb')) {
    type = 'Tủ điện';
  } else if (lower.includes('inverter') || lower.includes('biến tần')) {
    type = 'Inverter';
  }

  const manufacturer = row[6]?.toString().trim() || '';
  const model = row[7]?.toString().trim() || '';
  const serialNumber = row[8]?.toString().trim() || '';
  const yearOfManufacture = row[9]?.toString().trim() || '';
  const specsRaw = row[10]?.toString().trim() || '';
  
  let status: 'healthy' | 'warning' | 'critical' = 'healthy';
  const statusRaw = row[11]?.toString().toLowerCase() || '';
  if (statusRaw.includes('critical') || statusRaw.includes('nguy hiểm') || statusRaw.includes('hỏng')) {
    status = 'critical';
  } else if (statusRaw.includes('warn') || statusRaw.includes('cảnh báo')) {
    status = 'warning';
  }

  const health = parseFloat(row[12]?.toString().replace(',', '.') || '90') || 90;
  const criticality = (row[13]?.toString().trim().toUpperCase() as 'A' | 'B' | 'C') || 'B';
  const lastCheck = row[14]?.toString().trim() || new Date().toLocaleDateString('vi-VN');
  const pmCycleMonths = parseInt(row[15]?.toString() || '6', 10) || 6;

  return {
    id,
    name,
    customer,
    factory,
    location,
    type,
    manufacturer,
    model,
    serialNumber,
    yearOfManufacture,
    criticality,
    status,
    health,
    lastCheck,
    pmCycleMonths,
    notes: specsRaw,
    source: 'sheet',
    rawData: [...row]
  };
}
