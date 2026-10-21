import { CustomerRecord, EquipmentMasterRecord, EquipmentCategoryType, DEFAULT_CUSTOMERS } from '../utils/fieldSyncService';
import { DEFAULT_TEV_FSE_LIST, FseEngineer } from '../components/FseEngineerDropdown';

export const CUSTOMERS_STORAGE_KEY = 'tev_customers_master_v1';
export const EQUIPMENT_STORAGE_KEY = 'tev_equipment_cache';
export const FSE_STORAGE_KEY = 'tev_registered_fse_engineers_v2';

// Custom event names for cross-component synchrony
export const EVT_CUSTOMERS_UPDATED = 'tev_customers_updated';
export const EVT_EQUIPMENT_UPDATED = 'tev_equipment_updated';
export const EVT_FSE_UPDATED = 'tev_fse_updated';

/**
 * Default sample equipment list if storage is empty
 */
export const DEFAULT_EQUIPMENT_LIST: EquipmentMasterRecord[] = [
  {
    id: 'TR-01',
    name: 'Máy biến áp T1 2500kVA 22/0.4kV',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    factory: 'Nhà máy Bắc Ninh',
    location: 'Trạm biến áp chính TBA-01',
    type: 'Máy biến áp',
    manufacturer: 'Thibidi / ABB',
    model: 'ONAN-2500',
    serialNumber: 'TB-2022-8891',
    yearOfManufacture: '2022',
    criticality: 'A',
    status: 'healthy',
    health: 96,
    lastCheck: '12/03/2025',
    pmCycleMonths: 6,
    notes: 'Dầu khoáng tiêu chuẩn Mineral Oil, cấp bảo vệ IP54'
  },
  {
    id: 'TR-02',
    name: 'Máy biến áp T2 3200kVA 22/0.4kV',
    customer: 'Tập đoàn Điện tử Samsung SEVT',
    factory: 'Khu công nghệ cao Phổ Yên',
    location: 'Trạm điện trung tâm Subst-02',
    type: 'Máy biến áp',
    manufacturer: 'Dong Anh EEMC',
    model: 'ONAF-3200',
    serialNumber: 'DA-2021-4412',
    yearOfManufacture: '2021',
    criticality: 'A',
    status: 'healthy',
    health: 92,
    lastCheck: '18/02/2025',
    pmCycleMonths: 6,
    notes: 'Kiểm định NETA ATS định kỳ, DGA bình thường'
  },
  {
    id: 'MVCB-01',
    name: 'Máy cắt trung thế VCB 24kV 1250A 25kA',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    factory: 'Nhà máy Hải Phòng',
    location: 'Tủ xuất tuyến MV-Cubicle 03',
    type: 'Máy cắt',
    manufacturer: 'Schneider Electric',
    model: 'EVOLIS 24kV',
    serialNumber: 'SE-VCB-9982',
    yearOfManufacture: '2020',
    criticality: 'A',
    status: 'healthy',
    health: 95,
    lastCheck: '10/01/2025',
    pmCycleMonths: 12,
    notes: 'R_tiếp xúc 28µΩ, thời gian cắt 38ms đạt tiêu chuẩn'
  },
  {
    id: 'MTR-PUMP-01',
    name: 'Động cơ bơm nước làm mát chính 315kW 6.6kV',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    factory: 'Trạm biến áp TEV 110kV',
    location: 'Nhà bơm trung tâm Pump-Station',
    type: 'Động cơ',
    manufacturer: 'Siemens',
    model: '1LA8 315kW',
    serialNumber: 'SM-2019-1120',
    yearOfManufacture: '2019',
    criticality: 'B',
    status: 'healthy',
    health: 89,
    lastCheck: '05/03/2025',
    pmCycleMonths: 6,
    notes: 'Độ rung DE 1.8mm/s, nhiệt độ gối đỡ 62°C'
  },
  {
    id: 'GEN-STBY-01',
    name: 'Máy phát điện dự phòng 1500kVA 400V',
    customer: 'Công ty Cổ phần Sữa Vinamilk',
    factory: 'Nhà máy Tiên Sơn',
    location: 'Nhà máy phát khẩn cấp Genset-Room',
    type: 'Máy phát',
    manufacturer: 'Cummins Power Generation',
    model: 'QSK50-G4',
    serialNumber: 'CU-2023-3310',
    yearOfManufacture: '2023',
    criticality: 'A',
    status: 'healthy',
    health: 98,
    lastCheck: '15/02/2025',
    pmCycleMonths: 3,
    notes: 'ATS chuyển nguồn tự động 4.2s, hòa đồng bộ chuẩn'
  },
  {
    id: 'GND-GRID-01',
    name: 'Hệ thống tiếp địa trạm biến áp 110kV',
    customer: 'Công ty Cổ phần Năng lượng TEV',
    factory: 'Trạm biến áp TEV 110kV',
    location: 'Sân phân phối ngoài trời 110kV',
    type: 'Tủ điện',
    manufacturer: 'TEV Earthing Engineering',
    model: 'Deep-Well Mesh Grid',
    serialNumber: 'GW-2022-01',
    yearOfManufacture: '2022',
    criticality: 'A',
    status: 'healthy',
    health: 98,
    lastCheck: '01/03/2025',
    pmCycleMonths: 12,
    notes: 'Điện trở đất IEEE 81 đo được 0.42Ω (đạt ≤ 1.0Ω)'
  }
];

// ============================================================================
// 1. CUSTOMERS DIRECTORY (Sheet 'khachhang')
// ============================================================================

export function normalizeCustomerId(id: string, indexFallback = 1): string {
  if (!id) return `KH${String(indexFallback).padStart(3, '0')}`;
  const trimmed = id.trim().toUpperCase();
  if (/^KH\d{3,}$/.test(trimmed)) return trimmed;
  const digitsMatch = trimmed.match(/(\d+)$/);
  if (digitsMatch) {
    const num = parseInt(digitsMatch[1], 10);
    if (!isNaN(num)) return `KH${String(num).padStart(3, '0')}`;
  }
  return trimmed;
}

export function getNextCustomerId(existingCustomers?: CustomerRecord[]): string {
  const list = existingCustomers || getRegisteredCustomers();
  let maxSeq = 0;
  list.forEach(c => {
    if (!c || !c.id) return;
    const match = c.id.match(/^KH-?(\d+)$/i);
    if (match) {
      const seq = parseInt(match[1], 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
      return;
    }
    const matchSuffix = c.id.match(/^KH-.*-(\d+)$/i);
    if (matchSuffix) {
      const seq = parseInt(matchSuffix[1], 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
      return;
    }
    const trailing = c.id.match(/(\d+)$/);
    if (trailing) {
      const seq = parseInt(trailing[1], 10);
      if (!isNaN(seq) && seq < 10000 && seq > maxSeq) maxSeq = seq;
    }
  });
  return `KH${String(maxSeq + 1).padStart(3, '0')}`;
}

export function getRegisteredCustomers(): CustomerRecord[] {
  try {
    const saved = localStorage.getItem(CUSTOMERS_STORAGE_KEY);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) {
        // Ensure all customer IDs conform to standard KHXXX
        let hasChanges = false;
        const normalized = parsed.map((c: CustomerRecord, idx: number) => {
          const stdId = normalizeCustomerId(c.id, idx + 1);
          if (stdId !== c.id) {
            hasChanges = true;
            return { ...c, id: stdId };
          }
          return c;
        });
        if (hasChanges) {
          try {
            localStorage.setItem(CUSTOMERS_STORAGE_KEY, JSON.stringify(normalized));
          } catch {}
        }
        return normalized;
      }
    }
  } catch (e) {
    console.warn('Error reading customers from localStorage:', e);
  }
  return DEFAULT_CUSTOMERS;
}

export function saveCustomerRecord(customer: CustomerRecord): CustomerRecord[] {
  const current = getRegisteredCustomers();
  const cleanId = (customer.id || '').trim();
  const cleanName = (customer.name || '').trim();
  if (!cleanId || !cleanName) return current;

  const nowIso = new Date().toISOString();
  const updatedCustomer: CustomerRecord = {
    ...customer,
    id: cleanId,
    name: cleanName,
    factories: Array.isArray(customer.factories) && customer.factories.length > 0 
      ? customer.factories 
      : ['Nhà máy chính'],
    createdAt: customer.createdAt || nowIso,
    updatedAt: nowIso
  };

  const idx = current.findIndex(c => c.id.toLowerCase() === cleanId.toLowerCase());
  let next: CustomerRecord[];
  if (idx >= 0) {
    next = [...current];
    next[idx] = { ...next[idx], ...updatedCustomer };
  } else {
    next = [updatedCustomer, ...current];
  }

  try {
    localStorage.setItem(CUSTOMERS_STORAGE_KEY, JSON.stringify(next));
  } catch (e) {}

  window.dispatchEvent(new CustomEvent(EVT_CUSTOMERS_UPDATED, { detail: next }));
  return next;
}

export function deleteCustomerRecord(customerId: string): CustomerRecord[] {
  const current = getRegisteredCustomers();
  const cleanId = (customerId || '').trim().toLowerCase();
  const next = current.filter(c => c.id.toLowerCase() !== cleanId);

  try {
    localStorage.setItem(CUSTOMERS_STORAGE_KEY, JSON.stringify(next));
  } catch (e) {}

  window.dispatchEvent(new CustomEvent(EVT_CUSTOMERS_UPDATED, { detail: next }));
  return next;
}

// ============================================================================
// 2. EQUIPMENT DIRECTORY (Sheet 'thietbi')
// ============================================================================

export function getRegisteredEquipment(): EquipmentMasterRecord[] {
  try {
    const cached = localStorage.getItem(EQUIPMENT_STORAGE_KEY);
    if (cached) {
      const parsed = JSON.parse(cached);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed;
      }
    }
  } catch (e) {
    console.warn('Error reading equipment from localStorage:', e);
  }
  return DEFAULT_EQUIPMENT_LIST;
}

export function saveEquipmentRecord(equipment: Partial<EquipmentMasterRecord> & { id: string; name: string }): EquipmentMasterRecord[] {
  const current = getRegisteredEquipment();
  const cleanId = (equipment.id || '').trim();
  const cleanName = (equipment.name || '').trim();
  if (!cleanId || !cleanName) return current;

  const nowIso = new Date().toISOString();
  const updatedItem: EquipmentMasterRecord = {
    id: cleanId,
    name: cleanName,
    customer: equipment.customer || 'Công ty Cổ phần Năng lượng TEV',
    factory: equipment.factory || 'Nhà máy chính',
    location: equipment.location || 'Trạm điện phân phối',
    type: equipment.type || 'Máy biến áp',
    manufacturer: equipment.manufacturer || 'ABB / Schneider',
    model: equipment.model || '',
    serialNumber: equipment.serialNumber || '',
    yearOfManufacture: equipment.yearOfManufacture || new Date().getFullYear().toString(),
    criticality: equipment.criticality || 'B',
    status: equipment.status || 'healthy',
    health: typeof equipment.health === 'number' ? equipment.health : 95,
    lastCheck: equipment.lastCheck || new Date().toLocaleDateString('vi-VN'),
    pmCycleMonths: equipment.pmCycleMonths || 6,
    notes: equipment.notes || '',
    source: 'app',
    specs: equipment.specs,
    measurements: equipment.measurements
  };

  const idx = current.findIndex(e => e.id.toLowerCase() === cleanId.toLowerCase());
  let next: EquipmentMasterRecord[];
  if (idx >= 0) {
    next = [...current];
    next[idx] = { ...next[idx], ...updatedItem };
  } else {
    next = [updatedItem, ...current];
  }

  try {
    localStorage.setItem(EQUIPMENT_STORAGE_KEY, JSON.stringify(next));
  } catch (e) {}

  window.dispatchEvent(new CustomEvent(EVT_EQUIPMENT_UPDATED, { detail: next }));
  return next;
}

export function deleteEquipmentRecord(equipmentId: string): EquipmentMasterRecord[] {
  const current = getRegisteredEquipment();
  const cleanId = (equipmentId || '').trim().toLowerCase();
  const next = current.filter(e => e.id.toLowerCase() !== cleanId);

  try {
    localStorage.setItem(EQUIPMENT_STORAGE_KEY, JSON.stringify(next));
  } catch (e) {}

  window.dispatchEvent(new CustomEvent(EVT_EQUIPMENT_UPDATED, { detail: next }));
  return next;
}

// ============================================================================
// 3. FSE ENGINEERS DIRECTORY (Sheet 'FSE')
// ============================================================================

export function getRegisteredFseList(): FseEngineer[] {
  try {
    const saved = localStorage.getItem(FSE_STORAGE_KEY);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) {
        // Ensure default Võ Minh Tú is always first
        const tu = parsed.find((e: FseEngineer) => e.email === 'sgm1707@gmail.com') || DEFAULT_TEV_FSE_LIST[0];
        const others = parsed.filter((e: FseEngineer) => e.email !== 'sgm1707@gmail.com');
        return [tu, ...others];
      }
    }
  } catch (e) {
    console.warn('Error reading FSE from localStorage:', e);
  }
  return DEFAULT_TEV_FSE_LIST;
}

export function saveFseEngineerRecord(engineer: Partial<FseEngineer> & { name: string; email: string }): FseEngineer[] {
  const current = getRegisteredFseList();
  const cleanName = (engineer.name || '').trim();
  const cleanEmail = (engineer.email || '').trim().toLowerCase();
  if (!cleanName || !cleanEmail) return current;

  const id = engineer.id || `fse-${Date.now()}`;
  const updatedEngineer: FseEngineer = {
    id,
    name: cleanName,
    email: cleanEmail,
    role: engineer.role || 'Kỹ sư Thí nghiệm Hiện trường NETA',
    phone: engineer.phone || '',
    company: engineer.company || 'TEV Field Service'
  };

  const idx = current.findIndex(e => e.email.toLowerCase() === cleanEmail);
  let next: FseEngineer[];
  if (idx >= 0) {
    next = [...current];
    next[idx] = { ...next[idx], ...updatedEngineer };
  } else {
    next = [...current, updatedEngineer];
  }

  try {
    localStorage.setItem(FSE_STORAGE_KEY, JSON.stringify(next));
  } catch (e) {}

  window.dispatchEvent(new CustomEvent(EVT_FSE_UPDATED, { detail: next }));
  return next;
}

export function deleteFseEngineerRecord(idOrEmail: string): FseEngineer[] {
  const current = getRegisteredFseList();
  const cleanTarget = (idOrEmail || '').trim().toLowerCase();
  
  // Do not allow deleting the primary lead account
  if (cleanTarget === 'sgm1707@gmail.com') {
    return current;
  }

  const next = current.filter(e => 
    e.email.toLowerCase() !== cleanTarget && 
    (e.id || '').toLowerCase() !== cleanTarget
  );

  try {
    localStorage.setItem(FSE_STORAGE_KEY, JSON.stringify(next));
  } catch (e) {}

  window.dispatchEvent(new CustomEvent(EVT_FSE_UPDATED, { detail: next }));
  return next;
}

// ============================================================================
// 4. SITE PREPARATION & PROJECT DIRECTORY (Sheet 'Dự án' / 'duan')
// ============================================================================

export const PROJECTS_STORAGE_KEY = 'tev_projects_master_v2';
export const EVT_PROJECTS_UPDATED = 'tev_projects_updated';

export interface ProjectEquipmentItem {
  id?: string;
  kksCode: string; // e.g. "=10EAA01GS001", "+110kV-TR01"
  name: string; // e.g. "Máy biến áp chính T1 110/22kV 63MVA"
  type: string; // e.g. "Máy biến áp", "Máy cắt", "Động cơ", "Tủ điện", "Cáp điện", "Tiếp địa", "Pin & UPS"
  specs?: string; // e.g. "63MVA 110/22kV Thibidi"
  location?: string; // e.g. "Sân phân phối 110kV"
  notes?: string;
  testStandard?: string; // e.g. "ANSI/NETA ATS-2025"
  standard?: 'ats' | 'mts' | 'ecs';
  status?: 'chờ_test' | 'đã_test' | 'cảnh_báo' | 'đạt';
}

export interface ProjectMasterRecord {
  id: string; // e.g. "DA-2026-001"
  name: string; // e.g. "Dự án Thí nghiệm Định kỳ TBA 110kV Super Energy"
  customerId: string; // e.g. "KH001"
  customerName: string; // e.g. "Công ty Cổ phần Năng lượng Super Energy"
  factory: string; // e.g. "Nhà máy Điện Gió Mũi Né"
  status: 'chuẩn_bị' | 'đang_thực_hiện' | 'hoàn_thành';
  standard?: 'ats' | 'mts' | 'ecs';
  leadFse?: string; // e.g. "Võ Minh Tú"
  startDate?: string;
  endDate?: string;
  planningStartDate?: string; // e.g. "2026-03-01"
  planningEndDate?: string; // e.g. "2026-03-25"
  notes?: string;
  equipmentList: ProjectEquipmentItem[];
  createdAt: string;
  updatedAt: string;
}

export const DEFAULT_PROJECTS_LIST: ProjectMasterRecord[] = [
  {
    id: 'DA-2026-001',
    name: 'Dự án Thí nghiệm Định kỳ TBA 110kV Super Energy Mũi Né',
    customerId: 'KH001',
    customerName: 'Công ty Cổ phần Năng lượng Super Energy',
    factory: 'Nhà máy Điện Gió Mũi Né',
    status: 'đang_thực_hiện',
    standard: 'ats',
    leadFse: 'Võ Minh Tú (sgm1707@gmail.com)',
    startDate: '2026-03-01',
    endDate: '2026-03-25',
    planningStartDate: '2026-03-01',
    planningEndDate: '2026-03-25',
    notes: 'Thí nghiệm toàn diện MBA 110kV, Máy cắt SF6, Tủ trung thế 22kV và Hệ thống Tiếp địa theo NETA ATS-2025',
    equipmentList: [
      {
        kksCode: '=10EAA01GS001',
        name: 'Máy biến áp chính T1 110/22kV 63MVA',
        type: 'Máy biến áp',
        specs: '63MVA 110/22kV ONAN/ONAF Thibidi',
        location: 'Sân phân phối ngoài trời 110kV',
        testStandard: 'ANSI/NETA ATS-2025 Mục 7.2.2',
        standard: 'ats',
        status: 'đã_test'
      },
      {
        kksCode: '+110kV-131-Q0',
        name: 'Máy cắt 110kV ngăn lộ 131 (VCB/SF6)',
        type: 'Máy cắt',
        specs: '123kV 2500A 40kA SF6 Alstom',
        location: 'Ngăn lộ xuất tuyến 131',
        testStandard: 'ANSI/NETA ATS-2025 Mục 7.6.3',
        standard: 'ats',
        status: 'chờ_test'
      },
      {
        kksCode: '=10BBA01',
        name: 'Hệ thống tủ trung thế hợp bộ 22kV Switchgear C01-C08',
        type: 'Tủ điện',
        specs: '24kV 2000A 25kA Schneider SM6',
        location: 'Phòng phân phối trung thế 22kV',
        testStandard: 'ANSI/NETA ATS-2025 Mục 7.1.1',
        standard: 'ats',
        status: 'chờ_test'
      },
      {
        kksCode: '10LBA10AA001',
        name: 'Tuyến cáp ngầm 24kV cấp nguồn tự dùng 3x240mm2',
        type: 'Cáp điện',
        specs: '24kV 3x240 Cu/XLPE/PVC/DSTA/PVC',
        location: 'Hào cáp chính từ T1 sang tủ tự dùng',
        testStandard: 'ANSI/NETA ATS-2025 Mục 7.3.3',
        standard: 'ats',
        status: 'chờ_test'
      },
      {
        kksCode: '=00BAT01',
        name: 'Hệ thống giàn ắc quy trạm 110V DC 200Ah',
        type: 'Pin & UPS',
        specs: '110V DC 200Ah AGM VRLA Fiamm',
        location: 'Phòng DC & Battery',
        testStandard: 'ANSI/NETA ATS-2025 Mục 7.18.1.3',
        standard: 'ats',
        status: 'đã_test'
      },
      {
        kksCode: '=10EAA01-EARTH',
        name: 'Lưới tiếp địa trạm biến áp 110kV',
        type: 'Tiếp địa',
        specs: 'Đồng trần C120 liên kết cọc thép bọc đồng sâu 12m',
        location: 'Toàn trạm 110kV',
        testStandard: 'ANSI/NETA ATS-2025 Mục 7.13 & IEEE Std 81',
        standard: 'ats',
        status: 'đã_test'
      }
    ],
    createdAt: '2026-03-01T08:00:00.000Z',
    updatedAt: '2026-03-01T08:00:00.000Z'
  },
  {
    id: 'DA-2026-002',
    name: 'Dự án Nghiệm thu NETA ATS Hệ thống Điện VinFast Hải Phòng',
    customerId: 'KH002',
    customerName: 'Công ty TNHH Sản xuất và Kinh doanh VinFast',
    factory: 'Tổ hợp Nhà máy Cát Hải - Hải Phòng',
    status: 'chuẩn_bị',
    standard: 'ats',
    leadFse: 'Nguyễn Văn Bình',
    startDate: '2026-03-15',
    endDate: '2026-04-10',
    planningStartDate: '2026-03-15',
    planningEndDate: '2026-04-10',
    notes: 'Nghiệm thu đóng điện trạm biến áp khô và tủ phân phối xưởng Thân vỏ & Xưởng Sơn',
    equipmentList: [
      {
        kksCode: '+VF-TR01',
        name: 'Máy biến áp khô 2500kVA 22/0.4kV Xưởng Thân vỏ',
        type: 'Máy biến áp',
        specs: '2500kVA 22/0.4kV Cast Resin GE',
        location: 'Trạm điện Subst-01 Body Shop',
        testStandard: 'ANSI/NETA ATS-2025 Mục 7.2.1.2',
        standard: 'ats',
        status: 'chờ_test'
      },
      {
        kksCode: '+VF-ACB01',
        name: 'Máy cắt không khí ACB Masterpact MTZ2 4000A',
        type: 'Máy cắt',
        specs: 'ACB 4000A 65kA Micrologic 6.0X',
        location: 'Tủ MSB-01 Body Shop',
        testStandard: 'ANSI/NETA ATS-2025 Mục 7.6.1.2',
        standard: 'ats',
        status: 'chờ_test'
      },
      {
        kksCode: '+VF-MTR-PAINT',
        name: 'Động cơ quạt thông gió phòng sơn 250kW 380V',
        type: 'Động cơ',
        specs: '250kW 380V 1480rpm ABB IP55',
        location: 'Phòng sơn Paint Shop',
        testStandard: 'ANSI/NETA ATS-2025 Mục 7.15.1',
        standard: 'ats',
        status: 'chờ_test'
      },
      {
        kksCode: '+VF-UPS-DC',
        name: 'Hệ thống bộ lưu điện UPS 120kVA Line-Interactive',
        type: 'Pin & UPS',
        specs: '120kVA 3/3 Riello Master HP',
        location: 'Phòng Server & SCADA trung tâm',
        testStandard: 'ANSI/NETA ATS-2025 Mục 7.22',
        standard: 'ats',
        status: 'chờ_test'
      }
    ],
    createdAt: '2026-03-05T08:00:00.000Z',
    updatedAt: '2026-03-05T08:00:00.000Z'
  },
  {
    id: 'DA-2026-003',
    name: 'Dự án Commissioning NETA ECS Hệ Thống Solar & Pin BESS 50MWh',
    customerId: 'KH001',
    customerName: 'Công ty Cổ phần Năng lượng Super Energy',
    factory: 'Nhà máy Điện Mặt Trời Phong Điền',
    status: 'đang_thực_hiện',
    standard: 'ecs',
    leadFse: 'Trần Đình Trọng (fse.trong@tev.vn)',
    startDate: '2026-03-10',
    endDate: '2026-04-20',
    planningStartDate: '2026-03-10',
    planningEndDate: '2026-04-20',
    notes: 'Commissioning đóng điện chạy thử toàn diện Solar Farm 50MWp & Hệ thống Lưu trữ Pin BESS theo tiêu chuẩn ANSI/NETA ECS-2020',
    equipmentList: [
      {
        kksCode: '=50PV-INV01',
        name: 'Inverter trung tâm Solar Farm 3125kVA Sungrow',
        type: 'Pin & UPS',
        specs: '3125kVA 1500V DC / 600V AC Sungrow SG3125HV',
        location: 'Trạm Inverter Block 01',
        testStandard: 'ANSI/NETA ECS Section 8.1 Renewable PV',
        standard: 'ecs',
        status: 'đã_test'
      },
      {
        kksCode: '=50BESS-CONT01',
        name: 'Container Pin Lưu trữ Năng lượng BESS 5MWh LFP',
        type: 'Pin & UPS',
        specs: '1500V DC 5MWh CATL EnerC LFP Liquid Cooling',
        location: 'Khu vực BESS B01',
        testStandard: 'ANSI/NETA ECS Section 8.2 Energy Storage',
        standard: 'ecs',
        status: 'đang_test' as any
      },
      {
        kksCode: '=50RMU-MV01',
        name: 'Tủ RMU 24kV kết nối xuất tuyến Solar-BESS',
        type: 'Tủ điện',
        specs: '24kV 630A 20kA ABB SafeRing CCV',
        location: 'Trạm nâng áp Solar Substation',
        testStandard: 'ANSI/NETA ECS Section 7.2 MV Switchgear',
        standard: 'ecs',
        status: 'chờ_test'
      }
    ],
    createdAt: '2026-03-08T08:00:00.000Z',
    updatedAt: '2026-03-08T08:00:00.000Z'
  },
  {
    id: 'DA-2026-004',
    name: 'Dự án Bảo Trì Định Kỳ NETA MTS Trạm Biến Áp & Busway Formosa',
    customerId: 'KH003',
    customerName: 'Tập đoàn Thép Formosa Hà Tĩnh',
    factory: 'Trạm điện trung tâm 110/22kV Formosa',
    status: 'chuẩn_bị',
    standard: 'mts',
    leadFse: 'Lê Hoàng Nam (nam.lh@tev.vn)',
    startDate: '2026-04-01',
    endDate: '2026-04-30',
    planningStartDate: '2026-04-01',
    planningEndDate: '2026-04-30',
    notes: 'Bảo trì định kỳ NETA MTS Mục 7.4 Thanh cái vỏ kim loại Busway 24kV, Biến áp lực 40MVA và Phân tích DGA dầu',
    equipmentList: [
      {
        kksCode: '=FHS-BUSWAY-MV1',
        name: 'Thanh cái dẫn điện vỏ kim loại Busway 24kV 2500A Tuyến A',
        type: 'Cáp điện',
        specs: '24kV 2500A IP68 Cast Resin Busway Henikwon',
        location: 'Trục nối từ MBA 40MVA sang tủ Switchgear 22kV',
        testStandard: 'ANSI/NETA MTS-2023 Mục 7.4 Busways',
        standard: 'mts',
        status: 'chờ_test'
      },
      {
        kksCode: '=FHS-TR02-40MVA',
        name: 'Máy biến áp lực T2 110/22kV 40MVA ngâm dầu',
        type: 'Máy biến áp',
        specs: '40MVA 115/23kV ONAN/ONAF Dongfang',
        location: 'Trạm biến áp trung tâm',
        testStandard: 'ANSI/NETA MTS-2023 Mục 7.2.2 Liquid Transformers',
        standard: 'mts',
        status: 'chờ_test'
      }
    ],
    createdAt: '2026-03-12T08:00:00.000Z',
    updatedAt: '2026-03-12T08:00:00.000Z'
  }
];

export function getRegisteredProjects(): ProjectMasterRecord[] {
  try {
    const saved = localStorage.getItem(PROJECTS_STORAGE_KEY);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) {
        // Ensure planning dates and standard backward compatibility
        return parsed.map((p: any) => ({
          ...p,
          standard: p.standard || (p.name?.toLowerCase().includes('ecs') ? 'ecs' : p.name?.toLowerCase().includes('mts') ? 'mts' : 'ats'),
          planningStartDate: p.planningStartDate || p.startDate || '2026-03-01',
          planningEndDate: p.planningEndDate || p.endDate || '2026-03-25'
        }));
      }
    }
  } catch (e) {
    console.warn('Error reading projects from localStorage:', e);
  }
  return DEFAULT_PROJECTS_LIST;
}

export function saveProjectRecord(project: ProjectMasterRecord): ProjectMasterRecord[] {
  const current = getRegisteredProjects();
  const cleanId = (project.id || '').trim();
  const cleanName = (project.name || '').trim();
  if (!cleanId || !cleanName) return current;

  const nowIso = new Date().toISOString();
  const updatedProject: ProjectMasterRecord = {
    ...project,
    id: cleanId,
    name: cleanName,
    equipmentList: Array.isArray(project.equipmentList) ? project.equipmentList : [],
    createdAt: project.createdAt || nowIso,
    updatedAt: nowIso
  };

  const idx = current.findIndex(p => p.id.toLowerCase() === cleanId.toLowerCase());
  let next: ProjectMasterRecord[];
  if (idx >= 0) {
    next = [...current];
    next[idx] = { ...next[idx], ...updatedProject };
  } else {
    next = [updatedProject, ...current];
  }

  try {
    localStorage.setItem(PROJECTS_STORAGE_KEY, JSON.stringify(next));
  } catch (e) {}

  window.dispatchEvent(new CustomEvent(EVT_PROJECTS_UPDATED, { detail: next }));
  return next;
}

export function deleteProjectRecord(projectId: string): ProjectMasterRecord[] {
  const current = getRegisteredProjects();
  const cleanId = (projectId || '').trim().toLowerCase();
  const next = current.filter(p => p.id.toLowerCase() !== cleanId);

  try {
    localStorage.setItem(PROJECTS_STORAGE_KEY, JSON.stringify(next));
  } catch (e) {}

  window.dispatchEvent(new CustomEvent(EVT_PROJECTS_UPDATED, { detail: next }));
  return next;
}

export function getNextProjectId(existingProjects?: ProjectMasterRecord[]): string {
  const currentYear = new Date().getFullYear();
  const list = existingProjects || getRegisteredProjects();
  let maxSeq = 0;

  list.forEach(p => {
    if (!p || !p.id) return;
    const match = p.id.match(/^DA-(\d{4})-(\d+)$/i);
    if (match) {
      const yr = parseInt(match[1], 10);
      const seq = parseInt(match[2], 10);
      if (yr === currentYear && !isNaN(seq) && seq > maxSeq) {
        maxSeq = seq;
      }
      return;
    }
    const matchSimple = p.id.match(/^DA-(\d+)$/i);
    if (matchSimple) {
      const seq = parseInt(matchSimple[1], 10);
      if (!isNaN(seq) && seq > maxSeq) {
        maxSeq = seq;
      }
    }
  });

  return `DA-${currentYear}-${String(maxSeq + 1).padStart(3, '0')}`;
}

// ============================================================================
// 5. GOOGLE SHEETS TAB EXPORT HEADERS & ROW CONVERTERS
// ============================================================================

export const SHEET_TAB_NAMES = {
  CUSTOMER: 'khachhang',
  EQUIPMENT: 'thietbi',
  FSE: 'FSE',
  PROJECT: 'Dự án'
};

export const SHEET_FSE_HEADERS = [
  'Mã FSE',
  'Họ và tên Kỹ sư',
  'Email',
  'Vai trò / Chức danh',
  'Số điện thoại',
  'Đơn vị / Công ty',
  'Ngày tạo',
  'Ngày cập nhật'
];

export const SHEET_PROJECT_HEADERS = [
  'Mã Dự Án',
  'Tên Dự Án',
  'Mã KH',
  'Tên Khách Hàng',
  'Nhà Máy / Trạm',
  'Tiêu Chuẩn',
  'Số Lượng Thiết Bị',
  'Danh Sách KKS & Thiết Bị',
  'Kỹ Sư Trưởng',
  'Planning Bắt Đầu',
  'Planning Kết Thúc',
  'Trạng Thái',
  'Ghi Chú',
  'Ngày Cập Nhật'
];

export function fseToSheetRow(fse: FseEngineer): any[] {
  const nowStr = new Date().toISOString();
  return [
    fse.id || `fse-${fse.email}`,
    fse.name,
    fse.email,
    fse.role,
    fse.phone || '',
    fse.company || 'TEV Field Service',
    nowStr,
    nowStr
  ];
}

export function sheetRowToFse(row: any[]): FseEngineer | null {
  if (!row || row.length < 3) return null;
  const col0 = String(row[0] || '').trim();
  const col1 = String(row[1] || '').trim();
  const col2 = String(row[2] || '').trim();

  // Skip header row
  if (col0.toLowerCase().includes('mã fse') || col1.toLowerCase().includes('họ và tên') || col2.toLowerCase().includes('email')) {
    return null;
  }

  const name = col1 || col0;
  const email = col2.includes('@') ? col2 : (col1.includes('@') ? col1 : `${col0.toLowerCase().replace(/[^a-z0-9]/g, '')}@tev.vn`);
  const id = col0 || `fse-${Date.now()}`;

  return {
    id,
    name,
    email,
    role: String(row[3] || 'Kỹ sư Thí nghiệm Hiện trường NETA'),
    phone: String(row[4] || ''),
    company: String(row[5] || 'TEV Field Service')
  };
}

export function projectToSheetRow(project: ProjectMasterRecord): any[] {
  const nowStr = new Date().toISOString();
  const equipmentSummary = (project.equipmentList || [])
    .map(eq => `${eq.kksCode || eq.id}: ${eq.name} (${eq.type})`)
    .join('; ');

  return [
    project.id,
    project.name,
    project.customerId || '',
    project.customerName || '',
    project.factory || '',
    (project.standard || 'ats').toUpperCase(),
    (project.equipmentList || []).length,
    equipmentSummary,
    project.leadFse || '',
    project.planningStartDate || project.startDate || '',
    project.planningEndDate || project.endDate || '',
    project.status || 'chuẩn_bị',
    project.notes || '',
    nowStr
  ];
}

export function sheetRowToProject(row: any[]): ProjectMasterRecord | null {
  if (!row || row.length < 2) return null;
  const col0 = String(row[0] || '').trim();
  const col1 = String(row[1] || '').trim();

  // Skip header row
  if (col0.toLowerCase().includes('mã dự án') || col1.toLowerCase().includes('tên dự án')) {
    return null;
  }

  // Equipment list summary can be at row[7] (standard) or row[6] (fallback)
  const candidate7 = String(row[7] || '').trim();
  const candidate6 = String(row[6] || '').trim();
  const rawEquipmentSummary = candidate7.includes(':') ? candidate7 : (candidate6.includes(':') ? candidate6 : candidate7);
  const parsedEquipmentList: ProjectEquipmentItem[] = [];

  if (rawEquipmentSummary) {
    const items = rawEquipmentSummary.split(';');
    items.forEach(item => {
      const trimmed = item.trim();
      if (!trimmed) return;
      const colonIdx = trimmed.indexOf(':');
      if (colonIdx > 0) {
        const kks = trimmed.substring(0, colonIdx).trim();
        const rest = trimmed.substring(colonIdx + 1).trim();
        const parenStart = rest.lastIndexOf('(');
        const parenEnd = rest.lastIndexOf(')');
        let name = rest;
        let type = 'Máy biến áp';
        if (parenStart > 0 && parenEnd > parenStart) {
          name = rest.substring(0, parenStart).trim();
          type = rest.substring(parenStart + 1, parenEnd).trim();
        }
        parsedEquipmentList.push({
          kksCode: kks,
          name: name || kks,
          type: type || 'Thiết bị điện'
        });
      } else {
        parsedEquipmentList.push({
          kksCode: trimmed,
          name: trimmed,
          type: 'Thiết bị điện'
        });
      }
    });
  }

  const rawStd = String(row[5] || '').toLowerCase();
  const standard: 'ats' | 'mts' | 'ecs' = rawStd.includes('mts') ? 'mts' : rawStd.includes('ecs') ? 'ecs' : 'ats';
  const nowStr = new Date().toISOString();
  return {
    id: col0 || `DA-${Date.now()}`,
    name: col1 || 'Dự án mới',
    customerId: String(row[2] || ''),
    customerName: String(row[3] || ''),
    factory: String(row[4] || ''),
    standard,
    leadFse: String(row[8] || row[7] || ''),
    planningStartDate: String(row[9] || ''),
    planningEndDate: String(row[10] || ''),
    startDate: String(row[9] || ''),
    endDate: String(row[10] || ''),
    status: (row[11] === 'đang_thực_hiện' || row[11] === 'hoàn_thành') ? row[11] : 'chuẩn_bị',
    notes: String(row[12] || ''),
    equipmentList: parsedEquipmentList,
    createdAt: nowStr,
    updatedAt: String(row[13] || nowStr)
  };
}

export const EVT_NAVIGATE_TO_NETA_FORM = 'tev_navigate_to_neta_form';

/**
 * Maps an equipment type name (or Vietnamese description) to its corresponding NETA Spec ID
 */
export function resolveEquipmentTypeToSpecId(type?: string, name?: string): string {
  const t = `${type || ''} ${name || ''}`.toLowerCase();
  if (t.includes('cáp hạ') || t.includes('cable-lv') || t.includes('0.4kv') || t.includes('hạ thế') || t.includes('hạ áp')) return 'neta-cable-lv';
  if (t.includes('cáp') || t.includes('cable') || t.includes('feeder')) return 'neta-cable';
  if (t.includes('tiếp địa') || t.includes('ground') || t.includes('earth') || t.includes('nối đất')) return 'neta-grounding';
  if (t.includes('flooded') || t.includes('ngập') || t.includes('axit')) return 'neta-battery-flooded';
  if (t.includes('pin') || t.includes('ups') || t.includes('battery') || t.includes('ắc quy')) return 'neta-battery-vrla';
  if (t.includes('rơ le') || t.includes('relay')) return 'neta-relay';
  if (t.includes('inverter') || t.includes('pv') || t.includes('solar') || t.includes('biến tần')) return 'neta-pv';
  if (t.includes('cắt') || t.includes('breaker') || t.includes('vcb') || t.includes('acb') || t.includes('mcb')) return 'neta-lv-breaker';
  if (t.includes('biến áp khô') || t.includes('dry')) return 'neta-dry-type';
  if (t.includes('biến áp') || t.includes('transformer') || t.includes('mba')) return 'neta-liquid-transformer';
  if (t.includes('tủ') || t.includes('switchgear') || t.includes('rmu')) return 'neta-switchgear';
  if (t.includes('động cơ') || t.includes('motor')) return 'neta-motor';
  if (t.includes('phát') || t.includes('generator') || t.includes('genset')) return 'neta-generator';
  if (t.includes('ats') || t.includes('chuyển nguồn')) return 'neta-ats';
  if (t.includes('bess') || t.includes('lưu trữ')) return 'neta-bess';
  return 'neta-liquid-transformer';
}
