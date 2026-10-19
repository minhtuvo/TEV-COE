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
  | 'Inverter'
  | 'Cáp điện'
  | 'Pin & UPS'
  | 'Rơ le bảo vệ';

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
    id: 'KH001',
    name: 'Công ty Cổ phần Năng lượng TEV',
    factories: ['Nhà máy Bắc Ninh', 'Nhà máy Hải Phòng', 'Trạm biến áp TEV 110kV'],
    email: 'contact@tev-energy.vn',
    phone: '024.3988.7766',
    address: 'KCN Quế Võ, Tỉnh Bắc Ninh',
    contactPerson: 'Kỹ sư Trưởng Nguyễn Văn An',
    notes: 'Khách hàng trọng điểm - Hệ thống 110kV & 22kV'
  },
  {
    id: 'KH002',
    name: 'Samsung Electronics Vietnam (SEVT)',
    factories: ['Nhà máy SEVT Thái Nguyên', 'Phân xưởng Phụ trợ Utility 1', 'Substation 110kV Phổ Yên'],
    email: 'facilities.sevt@samsung.com',
    phone: '0208.356.9999',
    address: 'KCN Yên Bình, Phổ Yên, Thái Nguyên',
    contactPerson: 'Trưởng phòng Điện Mr. Park / Mr. Hùng',
    notes: 'Yêu cầu kiểm tra NETA ATS định kỳ 6 tháng/lần'
  },
  {
    id: 'KH003',
    name: 'Tổ hợp Sản xuất Ô tô VinFast',
    factories: ['Nhà máy Cát Hải - Hải Phòng', 'Xưởng Hàn Thân Vỏ', 'Trạm Biến áp 110kV Đình Vũ'],
    email: 'maintenance@vinfast.vn',
    phone: '0225.388.9988',
    address: 'Đảo Cát Hải, Huyện Cát Hải, Hải Phòng',
    contactPerson: 'KS Lê Hoàng Long',
    notes: 'Hệ thống tủ điện trung thế 24kV RMU & Robot Motors'
  },
  {
    id: 'KH004',
    name: 'Tập đoàn Thép Hòa Phát Dung Quất',
    factories: ['Khu liên hợp Gang Thép Dung Quất', 'Xưởng Cán Thép', 'Trạm 220kV Hòa Phát'],
    email: 'dienluc.dungquat@hoaphat.com.vn',
    phone: '0255.362.8888',
    address: 'KKT Dung Quất, Huyện Bình Sơn, Quảng Ngãi',
    contactPerson: 'Mr. Trần Đình Tuấn',
    notes: 'Môi trường nhiệt cao, bụi dẫn điện - Kiểm tra IR & TEV thường xuyên'
  },
  {
    id: 'KH005',
    name: 'Super Energy',
    factories: ['Nhà máy Điện Mặt Trời Super Energy', 'Trạm Biến Áp 22/110kV Super Energy'],
    email: 'operations@superenergy.vn',
    phone: '028.7300.9988',
    address: 'KCN Lộc An - Bình Sơn, Tỉnh Đồng Nai',
    contactPerson: 'Mr. Patrick Tan (Operations Director)',
    notes: 'Hệ thống Điện mặt trời Solar Farm & Tủ trung thế 24kV'
  }
];

import { MASTER_EQUIPMENT_104 } from './masterEquipment104';

export const DEFAULT_EQUIPMENT: EquipmentMasterRecord[] = MASTER_EQUIPMENT_104;

export const DEFAULT_WORK_ORDERS: any[] = [
  // --- 1. PHIẾU YÊU CẦU CÔNG VIỆC WR MỚI MỞ NGOÀI SITE (Chờ Daily Meeting xét duyệt - Priority 1 Khẩn Cấp) ---
  {
    wrId: 'WR-2026-001',
    id: 'WO-2026-PENDING-01',
    workPermitId: '',
    title: 'Phát hiện rò rỉ dầu & tiếng rít phóng điện cực cao MBA T2 110kV',
    description: 'Field Service / Operator đi tuần tra kiểm tra site phát hiện rò rỉ van xả đáy dầu MBA T2 và có tiếng phóng điện rít nhẹ tại sứ xuyên 110kV pha B. Cần họp Daily Meeting phê duyệt cô lập khẩn cấp.',
    equipmentId: 'TRF-110-01',
    equipmentName: 'Máy biến áp chính T1 110/22kV 63MVA',
    failureCode: 'ELEC-02',
    customer: 'Super Energy',
    customerId: 'KH-SUPER-05',
    factory: 'Nhà máy Điện Gió Mũi Né',
    workStream: 'CM',
    isUnplanned: true,
    priority: 'urgent',
    priorityCode: 'P1',
    workflowStage: '1_wr_pending',
    status: 'initiated',
    reportedBy: 'KS. Hoàng Đình Trọng (Site Operator FSE)',
    reportedAt: '2026-10-16T07:15:00.000Z',
    assignedTo: 'Chờ phân công Engineer sau Daily Meeting',
    responsibleApprove: 'Plant Manager / Maintenance Manager',
    responsibleDo: 'Đội Phản ứng nhanh TEV',
    blockingRequired: true,
    deEnergyRequired: true,
    dueDate: '2026-10-16',
    usedMaterials: 'Bộ gioăng Viton chịu nhiệt, Bộ thử áp lực dầu chuyên dụng',
    createdAt: '2026-10-16T07:15:00.000Z',
    updatedAt: '2026-10-16T07:30:00.000Z',
    downtimeStart: '',
    repairStart: '',
    repairEnd: '',
    restartTime: ''
  },

  // --- 2. PHIẾU WR ĐÃ ĐƯỢC DUYỆT TRONG DAILY MEETING -> THÀNH WO GIAO ENGINEER (Priority 3 - 7 ngày) ---
  {
    wrId: 'WR-2026-002',
    id: 'WO-2026-013',
    workPermitId: '',
    title: 'Thay thế bộ tiếp điểm phụ & motor nạp lò xo Máy cắt ACB Ngăn 02',
    description: 'Báo cáo từ ca trực: Motor nạp lò xo ACB Incomer có hiện tượng nóng bất thường (68°C) và nạp chậm > 15s. Đã thông qua Daily Meeting 08:00 sáng, MM phê duyệt chuyển thành WO giao Engineer lập phương án thi công.',
    equipmentId: 'MC-ACB-3200A',
    equipmentName: 'Máy cắt không khí tổng ACB Incomer 3200A',
    failureCode: 'MECH-01',
    customer: 'Samsung Electronics Vietnam (SEVT)',
    customerId: 'KH-SEVT-02',
    factory: 'Nhà máy SEVT Thái Nguyên',
    workStream: 'CM',
    isUnplanned: true,
    priority: 'high',
    priorityCode: 'P3',
    workflowStage: '2_daily_approved',
    status: 'approved',
    reportedBy: 'KTV. Vũ Văn Hải (Ca trực Vận hành)',
    reportedAt: '2026-10-15T16:30:00.000Z',
    dailyMeetingApprovedAt: '2026-10-16T08:30:00.000Z',
    dailyMeetingApprovedByRole: 'MM',
    dailyMeetingApprovedByName: 'Nguyễn Văn Minh (Maintenance Manager)',
    dailyMeetingNotes: 'Phê duyệt ưu tiên P3 xử lý trong 7 ngày. Giao KS Tuấn lập kế hoạch LOTO và chuẩn bị phụ tùng dự phòng.',
    assignedTo: 'KS. Lê Minh Tuấn',
    responsibleApprove: 'Maintenance Manager',
    responsibleDo: 'Tổ Cơ điện TEV',
    blockingRequired: true,
    deEnergyRequired: true,
    dueDate: '2026-10-23',
    usedMaterials: 'Motor nạp lò xo MCH Schneider 220VAC (1 bộ), Tiếp điểm phụ OF/SDE (2 cặp)',
    createdAt: '2026-10-15T16:30:00.000Z',
    updatedAt: '2026-10-16T08:35:00.000Z',
    downtimeStart: '',
    repairStart: '',
    repairEnd: '',
    restartTime: ''
  },

  // --- 3. WO ĐANG TRONG GIAI ĐOẠN ENGINEER PLANNING (LOTO / Cô Lập Năng Lượng / Phụ Tùng - Priority 7 Shutdown) ---
  {
    wrId: 'WR-2026-003',
    id: 'WO-2026-014',
    workPermitId: '',
    title: 'Bảo dưỡng thanh cái chính Busbar 24kV & Thay sứ xuyên (Shutdown toàn trạm)',
    description: 'Engineer đang lập biện pháp an toàn (Planning Phase): Lập danh mục điểm cô lập Lockout/Tagout (LOTO), quy trình de-energy cắt điện từ ngăn lộ trạm 110kV, chuẩn bị máy nâng và giàn giáo an toàn.',
    equipmentId: 'SWG-MV-24KV-01',
    equipmentName: 'Tủ hợp bộ đóng cắt trung thế 24kV Metal-Clad 8 ngăn',
    failureCode: 'PD-SURFACE',
    customer: 'Tập đoàn Thép Hòa Phát Dung Quất',
    customerId: 'KH-HOAPHAT-04',
    factory: 'Khu liên hợp Gang Thép Dung Quất',
    workStream: 'CM',
    isUnplanned: true,
    priority: 'high',
    priorityCode: 'P7',
    workflowStage: '3_planning',
    status: 'planned',
    reportedBy: 'KS. Trần Đình Tuấn',
    reportedAt: '2026-10-12T09:00:00.000Z',
    dailyMeetingApprovedAt: '2026-10-13T08:15:00.000Z',
    dailyMeetingApprovedByRole: 'PM',
    dailyMeetingApprovedByName: 'Trần Hải Đăng (Plant Manager)',
    dailyMeetingNotes: 'Duyệt chuyển sang lịch Shutdown toàn trạm vào cuối tháng. Yêu cầu Engineer lập biện pháp an toàn LOTO nghiêm ngặt.',
    assignedTo: 'KS. Phạm Quốc Hưng',
    responsibleApprove: 'Operation Manager & Plant Manager',
    responsibleDo: 'Nhóm Thí nghiệm & Bảo trì Cao thế TEV',
    blockingRequired: true,
    deEnergyRequired: true,
    lotoTagNumbers: 'LOTO-HP-24KV-01 đến 08',
    isolationPoints: 'Dao cách ly DS-271, DS-272, Máy cắt liên lạc Bus-Tie Q00, Tiếp địa di động 3 pha',
    dueDate: '2026-10-30',
    usedMaterials: 'Sứ xuyên thanh cái 24kV 1250A (3 quả), Bộ tiếp địa di động 24kV tiêu chuẩn IEC',
    createdAt: '2026-10-12T09:00:00.000Z',
    updatedAt: '2026-10-16T09:00:00.000Z',
    downtimeStart: '',
    repairStart: '',
    repairEnd: '',
    restartTime: ''
  },

  // --- 4. WO CẢI TIẾN THIẾT KẾ DMR - ĐÃ CẤP WORK PERMIT ĐANG THI CÔNG HIỆN TRƯỜNG (Priority 9 DMR) ---
  {
    wrId: 'WR-2026-004',
    id: 'WO-2026-015',
    workPermitId: 'WP-TEV-DMR-09',
    title: 'DMR: Lắp đặt hệ thống cảm biến giám sát online nhiệt độ không dây thanh cái SAW',
    description: 'Design Modification Request (DMR): Cải tiến kỹ thuật bổ sung 24 cảm biến nhiệt độ không dây công nghệ sóng âm bề mặt (SAW) không dùng pin vào các điểm nối thanh cái tủ trung thế để giám sát liên tục về SCADA.',
    equipmentId: 'SWG-MV-24KV-01',
    equipmentName: 'Tủ hợp bộ đóng cắt trung thế 24kV Metal-Clad 8 ngăn',
    failureCode: 'HOTSPOT-BUSBAR',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    customerId: 'KH-TEV-01',
    factory: 'Trạm 110kV Hòa Khánh',
    workStream: 'CM',
    isUnplanned: true,
    priority: 'medium',
    priorityCode: 'P9',
    workflowStage: '4_wp_issued',
    status: 'in-progress',
    reportedBy: 'KS. Đặng Văn Nam (FSE Project Lead)',
    reportedAt: '2026-10-05T10:00:00.000Z',
    dailyMeetingApprovedAt: '2026-10-06T08:30:00.000Z',
    dailyMeetingApprovedByRole: 'DM',
    dailyMeetingApprovedByName: 'Lê Văn Bảy (Deputy Plant Manager)',
    dailyMeetingNotes: 'Phê duyệt dự án cải tiến thiết kế DMR. Cho phép cấp Work Permit thi công có kiểm soát LOTO.',
    assignedTo: 'KS. Đặng Văn Nam',
    responsibleApprove: 'Deputy Manager & Safety Manager',
    responsibleDo: 'Đội Thi công Nâng cấp Hệ thống TEV',
    blockingRequired: true,
    deEnergyRequired: true,
    lotoTagNumbers: 'LOTO-WP-DMR-2026',
    isolationPoints: 'Ngăn tủ Incomer 24kV, Khóa liên động cơ Interlock chìa khóa Castell',
    wpIssuedAt: '2026-10-16T08:00:00.000Z',
    wpIssuedBy: 'Cán bộ An toàn Nguyễn Thanh Tùng',
    dueDate: '2026-10-18',
    usedMaterials: 'Bộ cảm biến SAW Transceiver (24 bộ), Bộ thu tín hiệu Reader Anten 433MHz (2 bộ)',
    createdAt: '2026-10-05T10:00:00.000Z',
    updatedAt: '2026-10-16T08:00:00.000Z',
    downtimeStart: '2026-10-16 08:30',
    repairStart: '2026-10-16 09:00',
    repairEnd: '',
    restartTime: ''
  },

  {
    id: 'WO-2026-001',
    workPermitId: 'WP-2026-001',
    title: 'Tủ bị phóng điện',
    description: 'Thực hiện kiểm tra phóng điện cục bộ và an toàn tủ trung thế',
    equipmentId: 'MVCB-001',
    equipmentName: 'Tủ máy cắt trung thế MVCB-001 24kV',
    failureCode: 'ELEC-01',
    customer: 'Super Energy',
    customerId: 'KH-SUPER-05',
    type: 'corrective',
    workStream: 'CM',
    isUnplanned: true,
    priority: 'High',
    priorityCode: 'P3',
    workflowStage: '3_planning',
    status: 'planned',
    assignedTo: 'sgm1707@gmail.com',
    responsibleApprove: 'FSE Lead / Trưởng nhóm',
    responsibleDo: 'Kỹ sư hiện trường',
    blockingRequired: true,
    dueDate: '2026-10-02',
    usedMaterials: 'Hóa chất vệ sinh cách điện, Bộ đo TEV & Cảm biến siêu âm',
    createdAt: '2026-09-28T08:00:00.000Z',
    updatedAt: '2026-09-29T02:00:00.000Z'
  },
  {
    id: 'WO-2026-002',
    workPermitId: 'WP-SEVT-104',
    title: 'Bảo trì ngăn lộ & Tiếp điểm Máy cắt ACB Incomer 3200A',
    description: 'Vệ sinh buồng dập hồ quang, đo điện trở tiếp xúc 3 pha R-Y-B, kiểm tra cơ cấu nạp lò xo và thử nghiệm trip unit điện tử.',
    equipmentId: 'MC-ACB-3200A',
    equipmentName: 'Máy cắt không khí tổng ACB Incomer 3200A',
    failureCode: 'MECH-WEAR',
    customer: 'Samsung Electronics Vietnam (SEVT)',
    customerId: 'KH-SEVT-02',
    type: 'corrective',
    isUnplanned: false,
    priority: 'Critical',
    status: 'completed',
    assignedTo: 'KS. Lê Minh Tuấn',
    responsibleApprove: 'GĐ Kỹ thuật Samsung SEVT',
    responsibleDo: 'Nhóm Thí nghiệm Cao thế TEV',
    blockingRequired: true,
    dueDate: '2026-09-25',
    usedMaterials: 'Bộ tiếp điểm hồ quang Arc Chute (3 bộ), Mỡ bôi trơn Kluber Isoflex (1 tuýp)',
    createdAt: '2026-09-18T09:00:00.000Z',
    updatedAt: '2026-09-25T16:00:00.000Z',
    downtimeStart: '2026-09-25 08:00',
    repairStart: '2026-09-25 08:30',
    repairEnd: '2026-09-25 15:30',
    restartTime: '2026-09-25 16:00'
  },
  {
    id: 'WO-2026-003',
    workPermitId: 'WP-VF-091',
    title: 'Đo rung động & Thay thế vòng bi Động cơ bơm 450kW',
    description: 'Phát hiện rung động DE tăng lên 4.2mm/s, tiến hành thay thế vòng bi SKF 6320 C3, cân bằng động laser trục động cơ.',
    equipmentId: 'MTR-PUMP-450KW',
    equipmentName: 'Động cơ bơm tuần hoàn nước giải nhiệt 450kW',
    failureCode: 'BEARING-VIBRATION',
    customer: 'Tổ hợp Sản xuất Ô tô VinFast',
    customerId: 'KH-VINFAST-03',
    type: 'corrective',
    isUnplanned: true,
    priority: 'High',
    status: 'completed',
    assignedTo: 'KS. Phạm Quốc Hưng',
    responsibleApprove: 'Quản đốc VinFast Cát Hải',
    responsibleDo: 'Tổ Cơ điện TEV Service',
    blockingRequired: false,
    dueDate: '2026-09-22',
    usedMaterials: 'Vòng bi SKF 6320 C3 (2 cái), Mỡ chịu nhiệt Mobil Polyrex EM (2 hộp)',
    createdAt: '2026-09-19T10:15:00.000Z',
    updatedAt: '2026-09-22T17:00:00.000Z',
    downtimeStart: '2026-09-21 22:00',
    repairStart: '2026-09-22 01:00',
    repairEnd: '2026-09-22 15:00',
    restartTime: '2026-09-22 16:30'
  },
  {
    id: 'WO-2026-004',
    workPermitId: 'WP-HP-552',
    title: 'Khảo sát phóng điện cục bộ TEV & Siêu âm Tủ trung thế 24kV',
    description: 'Thực hiện đo online mức TEV và cảm biến siêu âm Airborne tại 8 khoang tủ phân phối, phát hiện phóng điện nhẹ tại thanh cái Busbar khoang F03.',
    equipmentId: 'SWG-MV-24KV-01',
    equipmentName: 'Tủ hợp bộ đóng cắt trung thế 24kV Metal-Clad 8 ngăn',
    failureCode: 'PD-SURFACE',
    customer: 'Tập đoàn Thép Hòa Phát Dung Quất',
    customerId: 'KH-HOAPHAT-04',
    type: 'inspection',
    isUnplanned: false,
    priority: 'Medium',
    status: 'in_progress',
    assignedTo: 'KS. Trần Hoàng Long',
    responsibleApprove: 'Trưởng ban Quản lý Năng lượng',
    responsibleDo: 'Kỹ sư Thí nghiệm Hiện trường',
    blockingRequired: false,
    dueDate: '2026-10-10',
    usedMaterials: 'Thiết bị đo TEV UltraTEV Plus2, Cảm biến Parabolic Siêu âm',
    createdAt: '2026-09-24T08:30:00.000Z',
    updatedAt: '2026-09-28T11:00:00.000Z',
    downtimeStart: '',
    repairStart: '',
    repairEnd: '',
    restartTime: ''
  },
  {
    id: 'WO-2026-005',
    workPermitId: 'WP-SEVT-205',
    title: 'Thử tải Load Bank 100% Máy phát điện dự phòng Cummins 2000kVA',
    description: 'Thực hiện chạy thử tải bậc 25%, 50%, 75%, 100% công suất định mức theo tiêu chuẩn NFPA 110, kiểm tra đáp ứng điện áp và thời gian hòa điện lưới tự động.',
    equipmentId: 'GEN-STANDBY-2000KVA',
    equipmentName: 'Máy phát điện dự phòng khẩn cấp Cummins 2000kVA',
    failureCode: 'ROUTINE-PM',
    customer: 'Samsung Electronics Vietnam (SEVT)',
    customerId: 'KH-SEVT-02',
    type: 'preventive',
    isUnplanned: false,
    priority: 'Medium',
    status: 'completed',
    assignedTo: 'KS. Đặng Văn Nam',
    responsibleApprove: 'Trưởng ca Vận hành SEVT',
    responsibleDo: 'Đội Dịch vụ Kỹ thuật Máy phát TEV',
    blockingRequired: false,
    dueDate: '2026-09-26',
    usedMaterials: 'Lọc nhớt Fleetguard LF9009 (4 cái), Lọc nhiên liệu FS1000 (4 cái), Nước làm mát Donaldson (40L)',
    createdAt: '2026-09-21T07:45:00.000Z',
    updatedAt: '2026-09-26T18:00:00.000Z',
    downtimeStart: '',
    repairStart: '2026-09-26 08:00',
    repairEnd: '2026-09-26 16:30',
    restartTime: '2026-09-26 17:00'
  },
  {
    id: 'WO-2026-006',
    workPermitId: 'WP-TEV-2603',
    title: 'Thí nghiệm điện môi Furan & Khí hòa tan DGA MBA khô 2500kVA',
    description: 'Lấy mẫu kiểm tra cách điện cuộn đúc nhựa epoxy, đo điện trở cách điện High-Low, High-Earth và đối soát nhiệt độ hồng ngoại các điểm nối cọc cực.',
    equipmentId: 'TRF-SEVT-02',
    equipmentName: 'Máy biến áp phân phối khô Cast Resin 2500kVA',
    failureCode: 'INSULATION-TEST',
    customer: 'Samsung Electronics Vietnam (SEVT)',
    customerId: 'KH-SEVT-02',
    type: 'preventive',
    isUnplanned: false,
    priority: 'Low',
    status: 'open',
    assignedTo: 'KS. Nguyễn Văn An',
    responsibleApprove: 'Trưởng phòng Bảo trì SEVT',
    responsibleDo: 'Nhóm Thí nghiệm TEV',
    blockingRequired: true,
    dueDate: '2026-10-15',
    usedMaterials: 'Chổi lau vệ sinh cách điện chuyên dụng, Cồn công nghiệp 99% (5 Lít)',
    createdAt: '2026-09-25T14:20:00.000Z',
    updatedAt: '2026-09-27T09:10:00.000Z',
    downtimeStart: '',
    repairStart: '',
    repairEnd: '',
    restartTime: ''
  },
  {
    id: 'WO-2026-007',
    workPermitId: 'WP-TEV-2604',
    title: 'Bảo dưỡng Máy cắt chân không VCB 24kV 1250A',
    description: 'Đo độ mòn tiếp điểm chân không, đo thời gian đóng cắt 3 pha, kiểm tra hành trình nén tiếp điểm và tra mỡ cơ cấu lò xo.',
    equipmentId: 'MC-VCB-24KV',
    equipmentName: 'Máy cắt chân không VCB 24kV 1250A 25kA',
    failureCode: 'VACUUM-CHECK',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    customerId: 'KH-TEV-01',
    type: 'preventive',
    isUnplanned: false,
    priority: 'Medium',
    status: 'in_progress',
    assignedTo: 'KS. Lê Minh Tuấn',
    responsibleApprove: 'TP. Kỹ thuật Điện Lực',
    responsibleDo: 'Đội Bảo trì Trạm 110kV',
    blockingRequired: true,
    dueDate: '2026-10-08',
    usedMaterials: 'Mỡ tra cơ cấu chuyên dụng Kluber Isoflex Topas L32 (1 tuýp)',
    createdAt: '2026-09-23T11:00:00.000Z',
    updatedAt: '2026-09-28T15:00:00.000Z',
    downtimeStart: '2026-09-28 08:00',
    repairStart: '2026-09-28 08:30',
    repairEnd: '',
    restartTime: ''
  },
  {
    id: 'WO-2026-008',
    workPermitId: 'WP-VF-095',
    title: 'Thử nghiệm cao áp VLF 0.1Hz tuyến cáp ngầm 22kV XLPE 3x240mm²',
    description: 'Thử nghiệm Hipot VLF 0.1Hz tần số cực thấp theo tiêu chuẩn IEEE 400.2, đo dòng rò và đo Tan-Delta đánh giá lão hóa lớp cách điện cáp.',
    equipmentId: 'CBL-22KV-01',
    equipmentName: 'Tuyến cáp ngầm trung thế 22kV XLPE 3x240mm²',
    failureCode: 'TAN-DELTA-TEST',
    customer: 'Tổ hợp Sản xuất Ô tô VinFast',
    customerId: 'KH-VINFAST-03',
    type: 'preventive',
    isUnplanned: false,
    priority: 'High',
    status: 'open',
    assignedTo: 'KS. Nguyễn Thành Đạt',
    responsibleApprove: 'Chỉ huy trưởng NETA',
    responsibleDo: 'Nhóm Thí nghiệm Cáp TEV',
    blockingRequired: true,
    dueDate: '2026-10-12',
    usedMaterials: 'Băng keo cách điện 3M Scotch 23 (6 cuộn), Hộp đầu cáp co nhiệt Raychem (1 bộ)',
    createdAt: '2026-09-26T09:30:00.000Z',
    updatedAt: '2026-09-28T16:20:00.000Z',
    downtimeStart: '',
    repairStart: '',
    repairEnd: '',
    restartTime: ''
  },
  {
    id: 'WO-2026-009',
    workPermitId: 'WP-TEV-2605',
    title: 'Đo nội trở & xả tải dung lượng giàn pin UPS 220VDC 500Ah',
    description: 'Đo nội trở từng bình ắc quy theo chuẩn IEEE 1188, xả tải giả định kỳ 3 giờ để đo dung lượng thực tế và đối soát nhiệt độ bằng camera nhiệt.',
    equipmentId: 'BESS-220V-01',
    equipmentName: 'Giàn pin nguồn điều khiển 220VDC VRLA 500Ah',
    failureCode: 'BATTERY-IMPEDANCE',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    customerId: 'KH-TEV-01',
    type: 'preventive',
    isUnplanned: false,
    priority: 'Medium',
    status: 'completed',
    assignedTo: 'KS. Bùi Văn Hùng',
    responsibleApprove: 'Kỹ sư Trưởng trạm 110kV',
    responsibleDo: 'Tổ Nguồn DC & Tự động hóa',
    blockingRequired: false,
    dueDate: '2026-09-27',
    usedMaterials: 'Cầu chì DC 100A (4 cái), Khăn lau chống tĩnh điện (1 gói)',
    createdAt: '2026-09-22T08:00:00.000Z',
    updatedAt: '2026-09-27T17:30:00.000Z',
    downtimeStart: '',
    repairStart: '2026-09-27 08:30',
    repairEnd: '2026-09-27 16:30',
    restartTime: '2026-09-27 17:00'
  },
  {
    id: 'WO-2026-010',
    workPermitId: 'WP-HP-555',
    title: 'Hiệu chuẩn đặc tuyến Rơ le bảo vệ so lệch SEL-787',
    description: 'Bơm dòng nhị thứ thử nghiệm đặc tuyến so lệch 87T, quá dòng 50/51 và kiểm tra tiếp điểm ngắt cuộn cắt máy cắt 110kV.',
    equipmentId: 'RLY-DIFF-787',
    equipmentName: 'Rơ le bảo vệ so lệch máy biến áp SEL-787',
    failureCode: 'RELAY-CALIBRATION',
    customer: 'Tập đoàn Thép Hòa Phát Dung Quất',
    customerId: 'KH-HOAPHAT-04',
    type: 'preventive',
    isUnplanned: false,
    priority: 'High',
    status: 'in_progress',
    assignedTo: 'KS. Vũ Quang Huy',
    responsibleApprove: 'Trưởng phòng Thí nghiệm Điện',
    responsibleDo: 'Nhóm Rơ le & Điều khiển TEV',
    blockingRequired: true,
    dueDate: '2026-10-06',
    usedMaterials: 'Hợp bộ thử nghiệm Omicron CMC 356, Hàng kẹp thử nghiệm Phoenix Contact (12 cái)',
    createdAt: '2026-09-24T13:40:00.000Z',
    updatedAt: '2026-09-28T14:15:00.000Z',
    downtimeStart: '2026-09-28 13:00',
    repairStart: '2026-09-28 13:30',
    repairEnd: '',
    restartTime: ''
  },
  {
    id: 'WO-2026-011',
    workPermitId: 'WP-SEVT-209',
    title: 'Xử lý quá nhiệt thanh cái Busbar tủ điện MSB 4000A',
    description: 'Chụp ảnh nhiệt phát hiện điểm tiếp xúc thanh cái pha B đạt 88°C, vệ sinh bề mặt tiếp xúc đồng, siết lực lại bulong bằng cờ lê lực theo chuẩn DIN 43673.',
    equipmentId: 'SWG-MSB-4000A',
    equipmentName: 'Tủ phân phối hạ thế tổng MSB 4000A',
    failureCode: 'HOTSPOT-BUSBAR',
    customer: 'Samsung Electronics Vietnam (SEVT)',
    customerId: 'KH-SEVT-02',
    type: 'corrective',
    isUnplanned: true,
    priority: 'Critical',
    status: 'completed',
    assignedTo: 'KS. Trần Hoàng Long',
    responsibleApprove: 'Trưởng ban Quản lý Năng lượng SEVT',
    responsibleDo: 'Đội Cơ điện Phản ứng nhanh TEV',
    blockingRequired: true,
    dueDate: '2026-09-26',
    usedMaterials: 'Bulong siết lực Inox M12 kèm long đền Belleville (24 bộ), Mỡ dẫn điện tản nhiệt (1 hộp)',
    createdAt: '2026-09-25T15:00:00.000Z',
    updatedAt: '2026-09-26T21:00:00.000Z',
    downtimeStart: '2026-09-26 18:00',
    repairStart: '2026-09-26 18:30',
    repairEnd: '2026-09-26 20:30',
    restartTime: '2026-09-26 21:00'
  },
  {
    id: 'WO-2026-012',
    workPermitId: 'WP-TEV-2606',
    title: 'Kiểm tra đo điện trở tiếp địa hệ thống chống sét & nối đất trạm',
    description: 'Đo điện trở đất cột thu lôi và lưới nối đất trạm 110kV bằng phương pháp 3 cực theo tiêu chuẩn TCVN 9385:2012, điện trở đo được đạt 0.38 Ohm.',
    equipmentId: 'GND-GRID-110KV',
    equipmentName: 'Lưới tiếp địa an toàn & chống sét trạm 110kV',
    failureCode: 'GROUND-RESISTANCE',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    customerId: 'KH-TEV-01',
    type: 'inspection',
    isUnplanned: false,
    priority: 'Low',
    status: 'completed',
    assignedTo: 'KS. Phạm Quốc Hưng',
    responsibleApprove: 'Phòng An toàn Lao động',
    responsibleDo: 'Nhóm Đo lường Hiện trường TEV',
    blockingRequired: false,
    dueDate: '2026-09-28',
    usedMaterials: 'Đồng hồ đo tiếp địa Chauvin Arnoux CA 6460, Hóa chất giảm điện trở GEM (2 bao)',
    createdAt: '2026-09-27T08:00:00.000Z',
    updatedAt: '2026-09-28T16:00:00.000Z',
    downtimeStart: '',
    repairStart: '2026-09-28 09:00',
    repairEnd: '2026-09-28 15:30',
    restartTime: ''
  },
  // --- HỆ THỐNG PHIẾU BẢO TRÌ ĐỊNH KỲ PM (CHU KỲ 07, 14, 30, 60, 90, 180, 360 NGÀY - PRIORITY 6) ---
  {
    id: 'WO-PM-7D-01',
    workPermitId: 'WP-PM-7D-01',
    title: 'PM 07 Ngày: Kiểm tra mức dầu rò rỉ, màu hạt hút ẩm Silicagel MBA',
    description: 'Bảo trì định kỳ tuần (07 days): Kiểm tra mức dầu chỉ thị qua kính quan sát, kiểm tra độ đổi màu hạt hút ẩm bình thở Silicagel, ghi chép áp suất khí đệm Nitrogen.',
    equipmentId: 'TRF-110-01',
    equipmentName: 'Máy biến áp chính T1 110/22kV 63MVA',
    failureCode: 'ROUTINE-PM',
    customer: 'Super Energy',
    customerId: 'KH-SUPER-05',
    factory: 'Nhà máy Điện Gió Mũi Né',
    workStream: 'PM',
    pmCycleDays: 7,
    type: 'preventive',
    isUnplanned: false,
    priority: 'low',
    priorityCode: 'P6',
    workflowStage: '4_wp_issued',
    status: 'in-progress',
    assignedTo: 'KS. Hoàng Đình Trọng',
    responsibleApprove: 'Trưởng ban Vận hành',
    responsibleDo: 'Kỹ sư Giám sát Tuần',
    blockingRequired: false,
    dueDate: '2026-10-20',
    pmFrequency: 'weekly',
    estimatedTime: 2,
    actualTimeSpent: 1.5,
    usedMaterials: 'Hạt hút ẩm Silicagel chống ẩm (5kg), Dầu lau công nghiệp',
    createdAt: '2026-10-13T08:00:00.000Z',
    updatedAt: '2026-10-16T08:00:00.000Z'
  },
  {
    id: 'WO-PM-14D-01',
    workPermitId: 'WP-PM-14D-01',
    title: 'PM 14 Ngày: Quét TEV & Siêu âm phòng ngừa phóng điện tủ trung thế',
    description: 'Bảo trì định kỳ 14 ngày (Bi-weekly): Sử dụng máy đo UltraTEV Plus2 quét 14 khoang tủ trung thế 24kV, đo siêu âm tại các khe thông gió và đối chiếu lịch sử phóng điện.',
    equipmentId: 'SWG-MV-24KV-01',
    equipmentName: 'Tủ hợp bộ đóng cắt trung thế 24kV Metal-Clad 8 ngăn',
    failureCode: 'PD-SURFACE',
    customer: 'Tập đoàn Thép Hòa Phát Dung Quất',
    customerId: 'KH-HOAPHAT-04',
    factory: 'Khu liên hợp Gang Thép Dung Quất',
    workStream: 'PM',
    pmCycleDays: 14,
    type: 'preventive',
    isUnplanned: false,
    priority: 'medium',
    priorityCode: 'P6',
    workflowStage: '3_planning',
    status: 'planned',
    assignedTo: 'KS. Trần Đình Tuấn',
    responsibleApprove: 'Maintenance Manager',
    responsibleDo: 'Nhóm Thí nghiệm Hiện trường',
    blockingRequired: false,
    dueDate: '2026-10-24',
    pmFrequency: '2-weeks',
    estimatedTime: 3,
    actualTimeSpent: 0,
    usedMaterials: 'Máy đo phóng điện cục bộ UltraTEV Plus2, Tai nghe chuyên dụng',
    createdAt: '2026-10-10T08:00:00.000Z',
    updatedAt: '2026-10-16T08:30:00.000Z'
  },
  {
    id: 'WO-PM-2026-001',
    workPermitId: 'WP-PM-2026-001',
    title: 'PM 90 Ngày: Bảo trì định kỳ Quý 3 - Máy biến áp dầu T1 2000kVA',
    description: 'Lấy mẫu dầu DGA, đo điện trở cách điện 10kV, kiểm tra rơ-le Buchholz & van an toàn áp suất theo chu kỳ 90 ngày.',
    equipmentId: 'TRF-110-01',
    equipmentName: 'Máy biến áp chính T1 110/22kV 63MVA (Điện Gió Mũi Né)',
    failureCode: 'ROUTINE-PM',
    customer: 'Super Energy',
    customerId: 'KH-SUPER-05',
    factory: 'Nhà máy Điện Gió Mũi Né',
    workStream: 'PM',
    pmCycleDays: 90,
    type: 'preventive',
    isUnplanned: false,
    priority: 'high',
    priorityCode: 'P6',
    workflowStage: '6_wo_closed',
    status: 'completed',
    assignedTo: 'KS. Nguyễn Văn S',
    responsibleApprove: 'GĐ Kỹ thuật Super Energy',
    responsibleDo: 'Kỹ sư FSE TEV',
    blockingRequired: true,
    dueDate: '2026-09-08',
    completedDate: '2026-09-07',
    pmFrequency: '3-months',
    estimatedTime: 6,
    actualTimeSpent: 5.5,
    usedMaterials: 'Bộ chai lấy mẫu dầu thủy tinh Syringe DGA, Gioăng làm kín van',
    createdAt: '2026-09-01T08:00:00.000Z',
    updatedAt: '2026-09-07T16:00:00.000Z'
  },
  {
    id: 'WO-PM-2026-002',
    workPermitId: 'WP-PM-2026-002',
    title: 'Bảo trì định kỳ Tháng 9 - Bộ chuyển nguồn tự động ATS 800A',
    description: 'Vệ sinh tiếp điểm cực Pole C, kiểm tra khóa liên động cơ khí, đo điện trở mạch điều khiển.',
    equipmentId: 'SWG-LV-01',
    equipmentName: 'Tủ chuyển nguồn ATS Khẩn cấp (800A 4P)',
    failureCode: 'ATS-INSPECTION',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    customerId: 'KH-TEV-01',
    factory: 'Tòa nhà TEV Office Tower',
    type: 'preventive',
    isUnplanned: false,
    priority: 'urgent',
    status: 'completed',
    assignedTo: 'KS. Trần Văn H',
    responsibleApprove: 'Quản lý Vận hành Tòa nhà',
    responsibleDo: 'Kỹ sư Cơ điện TEV',
    blockingRequired: true,
    dueDate: '2026-09-12',
    completedDate: '2026-09-11',
    pmFrequency: 'monthly',
    estimatedTime: 4,
    actualTimeSpent: 4.5,
    usedMaterials: 'Bình xịt vệ sinh tiếp điểm CRC Contact Cleaner, Mỡ dẫn điện tiếp xúc',
    createdAt: '2026-09-02T08:00:00.000Z',
    updatedAt: '2026-09-11T17:00:00.000Z'
  },
  {
    id: 'WO-PM-2026-003',
    workPermitId: 'WP-PM-2026-003',
    title: 'Bảo trì định kỳ Tháng 9 - Tổ máy phát Diesel khẩn cấp 500kW',
    description: 'Kiểm tra bảo vệ ngắt quá tốc độ (115%), sấy cuộn stator nâng PI, thử tải nấc 100% Load Bank.',
    equipmentId: 'GEN-STANDBY-2000KVA',
    equipmentName: 'Tổ máy phát điện Diesel 500kW (Cummins)',
    failureCode: 'GEN-TEST',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    customerId: 'KH-TEV-01',
    factory: 'Bệnh Viện TEV General Hospital',
    type: 'preventive',
    isUnplanned: false,
    priority: 'urgent',
    status: 'in-progress',
    assignedTo: 'KS. Nguyễn Văn S',
    responsibleApprove: 'Trưởng ban An toàn Y tế',
    responsibleDo: 'Đội Máy phát TEV Service',
    blockingRequired: false,
    dueDate: '2026-09-21',
    pmFrequency: 'monthly',
    estimatedTime: 8,
    actualTimeSpent: 4,
    usedMaterials: 'Lọc gió Donaldson, Nước làm mát chống đông, Axit châm bình ắc quy',
    createdAt: '2026-09-05T08:00:00.000Z',
    updatedAt: '2026-09-21T10:00:00.000Z'
  },
  {
    id: 'WO-PM-2026-004',
    workPermitId: 'WP-PM-2026-004',
    title: 'Bảo dưỡng kiểm tra Giàn ắc quy VRLA AGM 120kVA Data Center',
    description: 'Đo nội trở nội suy mΩ, kiểm tra cân bằng điện áp cực đất và nhiệt độ cực âm.',
    equipmentId: 'BAT-BESS-01',
    equipmentName: 'Giàn ắc quy VRLA 120kVA (12V x 40 bình)',
    failureCode: 'BATT-IMPEDANCE',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    customerId: 'KH-TEV-01',
    factory: 'Trung Tâm Dữ Liệu TEV Tier-III',
    type: 'preventive',
    isUnplanned: false,
    priority: 'high',
    status: 'initiated',
    assignedTo: 'KS. Lê Hoàng D',
    responsibleApprove: 'GĐ Trung tâm Dữ liệu',
    responsibleDo: 'Kỹ sư Nguồn DC TEV',
    blockingRequired: false,
    dueDate: '2026-09-18',
    pmFrequency: 'monthly',
    estimatedTime: 5,
    actualTimeSpent: 0,
    usedMaterials: 'Đồng hồ đo nội trở Hioki BT3554, Giẻ lau chống tĩnh điện',
    createdAt: '2026-09-08T08:00:00.000Z',
    updatedAt: '2026-09-18T09:00:00.000Z'
  },
  {
    id: 'WO-PM-2026-005',
    workPermitId: 'WP-PM-2026-005',
    title: 'Bảo trì & Kiểm tra rơ-le số bảo vệ xuất tuyến 22kV SEL-751',
    description: 'Bơm dòng thứ cấp thử đặc tính thời gian quá dòng 50/51, kiểm tra cổng SCADA Modbus.',
    equipmentId: 'REL-SEL-751-01',
    equipmentName: 'Rơ-le bảo vệ kỹ thuật số SEL-751 Ngăn Lộ 471',
    failureCode: 'RELAY-TEST',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    customerId: 'KH-TEV-01',
    factory: 'Khu Công Nghiệp VSIP 1',
    type: 'preventive',
    isUnplanned: false,
    priority: 'medium',
    status: 'completed',
    assignedTo: 'KS. Phạm Minh T',
    responsibleApprove: 'Trưởng nhóm Bảo vệ Rơ-le',
    responsibleDo: 'Nhóm Thí nghiệm TEV',
    blockingRequired: true,
    dueDate: '2026-09-15',
    completedDate: '2026-09-14',
    pmFrequency: '6-months',
    estimatedTime: 4,
    actualTimeSpent: 3.5,
    usedMaterials: 'Dây cắm thí nghiệm an toàn 4mm, Hàng kẹp Phoenix Contact',
    createdAt: '2026-09-05T08:00:00.000Z',
    updatedAt: '2026-09-14T16:00:00.000Z'
  },
  {
    id: 'WO-PM-2026-006',
    workPermitId: 'WP-PM-2026-006',
    title: 'Đo phóng điện cục bộ Partial Discharge (PD) Tủ trung thế RMU',
    description: 'Đo cảm biến TEV & Siêu âm màng ngăn tủ cáp đầu vào trạm 22kV.',
    equipmentId: 'SWG-MV-24KV-01',
    equipmentName: 'Tủ máy cắt trung thế RMU Schneider 24kV',
    failureCode: 'PD-ULTRASOUND',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    customerId: 'KH-TEV-01',
    factory: 'Nhà máy Sản xuất TEV Tân Uyên',
    type: 'preventive',
    isUnplanned: false,
    priority: 'high',
    status: 'initiated',
    assignedTo: 'KS. Trần Văn H',
    responsibleApprove: 'Quản đốc Nhà máy Tân Uyên',
    responsibleDo: 'Kỹ sư Thí nghiệm Hiện trường',
    blockingRequired: false,
    dueDate: '2026-09-26',
    pmFrequency: 'monthly',
    estimatedTime: 3,
    actualTimeSpent: 0,
    usedMaterials: 'Máy đo UltraTEV Plus2, Cảm biến tiếp xúc vỏ kim loại',
    createdAt: '2026-09-12T08:00:00.000Z',
    updatedAt: '2026-09-20T10:00:00.000Z'
  },
  {
    id: 'WO-PM-2026-007',
    workPermitId: 'WP-PM-2026-007',
    title: 'Khảo sát ảnh nhiệt hồng ngoại (IR) toàn bộ thanh cái hạ thế',
    description: 'Quét camera Flir E76 các đầu cosse cáp tổng tủ MSB khi mang tải trên 75%.',
    equipmentId: 'MC-ACB-3200A',
    equipmentName: 'Tủ phân phối chính MSB 2500A Khối Sản Xuất',
    failureCode: 'THERMO-SCAN',
    customer: 'Samsung Electronics Vietnam (SEVT)',
    customerId: 'KH-SEVT-02',
    factory: 'Nhà máy Sản xuất TEV Tân Uyên',
    type: 'preventive',
    isUnplanned: false,
    priority: 'medium',
    status: 'initiated',
    assignedTo: 'KS. Lê Hoàng D',
    responsibleApprove: 'Trưởng ban An toàn Cơ điện',
    responsibleDo: 'Kỹ sư Chẩn đoán IR',
    blockingRequired: false,
    dueDate: '2026-09-27',
    pmFrequency: '3-months',
    estimatedTime: 4,
    actualTimeSpent: 0,
    usedMaterials: 'Camera nhiệt Flir E76, Thước đo khoảng cách laser',
    createdAt: '2026-09-15T08:00:00.000Z',
    updatedAt: '2026-09-21T09:00:00.000Z'
  },
  {
    id: 'WO-PM-2026-008',
    workPermitId: 'WP-PM-2026-008',
    title: 'Bảo trì thí nghiệm VLF cáp ngầm trung thế 24kV Cu/XLPE 3x240mm2',
    description: 'Thử nghiệm điện áp cao tần số thấp 0.1Hz và đo góc tổn hao điện môi Tan Delta.',
    equipmentId: 'CAB-24KV-01',
    equipmentName: 'Tuyến cáp ngầm trung thế 24kV cấp trạm biến áp',
    failureCode: 'VLF-HIPOT',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    customerId: 'KH-TEV-01',
    factory: 'Khu Công Nghiệp VSIP 1',
    type: 'preventive',
    isUnplanned: false,
    priority: 'high',
    status: 'initiated',
    assignedTo: 'KS. Phạm Minh T',
    responsibleApprove: 'Trưởng ban Quản lý Hạ tầng KCN',
    responsibleDo: 'Đội Thí nghiệm Cáp lực TEV',
    blockingRequired: true,
    dueDate: '2026-09-20',
    pmFrequency: '1-year',
    estimatedTime: 6,
    actualTimeSpent: 0,
    usedMaterials: 'Thiết bị HVA28 VLF Test Set, Tiếp địa di động cáp',
    createdAt: '2026-09-10T08:00:00.000Z',
    updatedAt: '2026-09-19T14:00:00.000Z'
  },
  {
    id: 'WO-PM-2026-009',
    workPermitId: 'WP-PM-2026-009',
    title: 'Bảo trì trạm sạc xe điện nhanh DC Fast EVSE 120kW',
    description: 'Kiểm tra điện trở tiếp địa súng sạc, kiểm tra quạt làm mát công suất và cách điện DC.',
    equipmentId: 'EVSE-DC-120KW',
    equipmentName: 'Trạm sạc xe điện cao tốc 120kW Dual Gun',
    failureCode: 'EVSE-INSPECTION',
    customer: 'Tổ hợp Sản xuất Ô tô VinFast',
    customerId: 'KH-VINFAST-03',
    factory: 'Trạm Sạc Cao Tốc Long Thành',
    type: 'preventive',
    isUnplanned: false,
    priority: 'low',
    status: 'completed',
    assignedTo: 'KS. Nguyễn Văn S',
    responsibleApprove: 'Quản lý Trạm Sạc EVSE',
    responsibleDo: 'Kỹ sư Dịch vụ TEV',
    blockingRequired: false,
    dueDate: '2026-09-05',
    completedDate: '2026-09-04',
    pmFrequency: 'monthly',
    estimatedTime: 3,
    actualTimeSpent: 3,
    usedMaterials: 'Hóa chất vệ sinh cổng sạc CCS2, Lọc bụi công nghiệp',
    createdAt: '2026-09-01T08:00:00.000Z',
    updatedAt: '2026-09-04T17:00:00.000Z'
  },
  {
    id: 'WO-PM-2026-010',
    workPermitId: 'WP-PM-2026-010',
    title: 'Kiểm tra định kỳ Quý - Dàn Inverter Năng lượng mặt trời áp mái 500kW',
    description: 'Kiểm tra đường đặc tính I-V string, đo điện trở cách điện DC mảng pin và vệ sinh tản nhiệt.',
    equipmentId: 'INV-SOLAR-01',
    equipmentName: 'Chuỗi biến tần Solar Inverter Sungrow 110kW',
    failureCode: 'INVERTER-IV-CURVE',
    customer: 'Super Energy',
    customerId: 'KH-SUPER-05',
    factory: 'Nhà máy Điện Gió Mũi Né',
    type: 'preventive',
    isUnplanned: false,
    priority: 'medium',
    status: 'in-progress',
    assignedTo: 'KS. Lê Hoàng D',
    responsibleApprove: 'Trưởng ban Năng lượng Tái tạo',
    responsibleDo: 'Đội O&M Solar & Wind TEV',
    blockingRequired: false,
    dueDate: '2026-09-29',
    pmFrequency: '3-months',
    estimatedTime: 5,
    actualTimeSpent: 2,
    usedMaterials: 'Đồng hồ đo Solar I-V Seaward PV200, Quạt hút tản nhiệt dự phòng',
    createdAt: '2026-09-18T08:00:00.000Z',
    updatedAt: '2026-09-24T11:00:00.000Z'
  }
];

export const DEFAULT_INVENTORY: any[] = [
  {
    id: 'VT-NYNAS-OIL',
    name: 'Dầu khoáng cách điện máy biến áp Nynas Nytro 4000X',
    sku: 'OIL-NYNAS-4000X',
    category: 'Hóa chất & Dầu cách điện',
    quantity: 1200,
    unit: 'Lít',
    minStock: 400,
    location: 'Kho Hóa chất - Kệ A1',
    price: 65000,
    createdAt: '2026-09-01T08:00:00.000Z',
    updatedAt: '2026-09-28T10:00:00.000Z'
  },
  {
    id: 'VT-SKF-6320C3',
    name: 'Vòng bi động cơ cao tốc SKF 6320 C3 Explorer',
    sku: 'BRG-SKF-6320C3',
    category: 'Vòng bi & Cơ khí',
    quantity: 16,
    unit: 'Cái',
    minStock: 4,
    location: 'Kho Cơ khí - Ngăn B03',
    price: 4850000,
    createdAt: '2026-09-01T08:00:00.000Z',
    updatedAt: '2026-09-25T14:00:00.000Z'
  },
  {
    id: 'VT-SF6-GAS-50KG',
    name: 'Bình khí tinh khiết SF6 cách điện cao thế 50kg',
    sku: 'GAS-SF6-50KG',
    category: 'Khí cách điện',
    quantity: 8,
    unit: 'Bình',
    minStock: 2,
    location: 'Kho Khí đặc chủng - Khu C',
    price: 28000000,
    createdAt: '2026-09-05T08:00:00.000Z',
    updatedAt: '2026-09-26T09:00:00.000Z'
  },
  {
    id: 'VT-ARC-CHUTE-NW',
    name: 'Bộ buồng dập hồ quang máy cắt MasterPact NW',
    sku: 'PART-SCH-NW-ARC',
    category: 'Phụ tùng máy cắt',
    quantity: 12,
    unit: 'Bộ',
    minStock: 3,
    location: 'Kho Điện cao thế - Kệ D1',
    price: 9200000,
    createdAt: '2026-09-10T08:00:00.000Z',
    updatedAt: '2026-09-25T11:00:00.000Z'
  },
  {
    id: 'VT-CABLE-3M-24KV',
    name: 'Bộ đầu cáp co nguội 3M ngoài trời 24kV 3x240mm²',
    sku: 'TERM-3M-QT3-24KV',
    category: 'Vật tư cáp điện',
    quantity: 10,
    unit: 'Bộ',
    minStock: 3,
    location: 'Kho Cáp & Phụ kiện - Kệ E2',
    price: 6800000,
    createdAt: '2026-09-08T08:00:00.000Z',
    updatedAt: '2026-09-26T14:00:00.000Z'
  },
  {
    id: 'VT-LUBE-MOBIL-EM',
    name: 'Mỡ bò chịu nhiệt & chống phóng điện Mobil Polyrex EM',
    sku: 'GREASE-MOBIL-EM',
    category: 'Dầu mỡ bôi trơn',
    quantity: 24,
    unit: 'Hộp',
    minStock: 6,
    location: 'Kho Hóa chất - Kệ A3',
    price: 380000,
    createdAt: '2026-09-02T08:00:00.000Z',
    updatedAt: '2026-09-22T10:00:00.000Z'
  },
  {
    id: 'VT-FILTER-FLEETGUARD',
    name: 'Bộ lọc nhiên liệu & lọc nhớt Cummins Fleetguard FS1000/LF9009',
    sku: 'FLT-CUMMINS-SET',
    category: 'Vật tư máy phát',
    quantity: 18,
    unit: 'Bộ',
    minStock: 4,
    location: 'Kho Động cơ - Kệ F1',
    price: 1850000,
    createdAt: '2026-09-05T08:00:00.000Z',
    updatedAt: '2026-09-26T16:00:00.000Z'
  },
  {
    id: 'VT-SILICAGEL-25KG',
    name: 'Hạt hút ẩm biến tính Silicagel bình thở MBA 25kg',
    sku: 'DES-SILICA-25KG',
    category: 'Vật tư máy biến áp',
    quantity: 20,
    unit: 'Bao',
    minStock: 5,
    location: 'Kho Hóa chất - Kệ A2',
    price: 1250000,
    createdAt: '2026-09-03T08:00:00.000Z',
    updatedAt: '2026-09-27T08:00:00.000Z'
  },
  {
    id: 'VT-FUSE-DC-100A',
    name: 'Cầu chì gốm bảo vệ nguồn DC hệ thống UPS 250V 100A',
    sku: 'FUSE-BUSSMANN-100A',
    category: 'Thiết bị bảo vệ',
    quantity: 35,
    unit: 'Cái',
    minStock: 10,
    location: 'Kho Linh kiện - Ngăn C05',
    price: 450000,
    createdAt: '2026-09-10T08:00:00.000Z',
    updatedAt: '2026-09-27T10:00:00.000Z'
  },
  {
    id: 'VT-TAPE-3M-SCOTCH23',
    name: 'Băng keo cách điện tự dính cao su 3M Scotch 23 19mmx9.1m',
    sku: 'TAPE-3M-SCOTCH23',
    category: 'Vật tư tiêu hao',
    quantity: 60,
    unit: 'Cuộn',
    minStock: 15,
    location: 'Kho Tiêu hao - Kệ G1',
    price: 220000,
    createdAt: '2026-09-05T08:00:00.000Z',
    updatedAt: '2026-09-28T09:00:00.000Z'
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
