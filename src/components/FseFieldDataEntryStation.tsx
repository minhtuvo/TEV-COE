import React, { useState, useEffect, useMemo } from 'react';
import {
  FileSpreadsheet,
  CheckCircle2,
  AlertTriangle,
  RefreshCw,
  Search,
  Plus,
  Camera,
  Layers,
  Wrench,
  Clock,
  Send,
  Zap,
  Power,
  RotateCcw,
  ShieldCheck,
  Droplets,
  Calendar,
  Building,
  MapPin,
  ClipboardList,
  Flame,
  Check,
  UploadCloud,
  DownloadCloud,
  Database,
  ArrowDownUp,
  ExternalLink,
  ChevronDown,
  ChevronRight,
  Info,
  SlidersHorizontal,
  X,
  UserCheck,
  Lock,
  ArrowRight,
  FileCheck,
  Briefcase,
  ListChecks,
  FolderPlus,
  Cable,
  FileCheck2,
  Sparkles,
  Scale,
  Gauge,
  ShieldAlert
} from 'lucide-react';
import { NameplateScannerModal } from './NameplateScannerModal';
import { ExtractedNameplateData } from '../../server/nameplateOcrAnalyzer';
import { parseVersionTimestamp } from '../utils/syncReconciliationService';
import {
  getRegisteredProjects,
  ProjectMasterRecord,
  ProjectEquipmentItem,
  EVT_PROJECTS_UPDATED,
  EVT_NAVIGATE_TO_NETA_FORM,
  getNextCustomerId
} from '../services/masterDirectoryService';
import { MasterDirectorySheetManagerModal } from './MasterDirectorySheetManagerModal';
import { NetaAtsUniversalReportView } from './NetaAtsUniversalReportView';
import {
  NETA_ATS_23_MASTER_SPECS,
  NetaAtsSpecificationDef,
  NetaAtsTestItem
} from '../data/netaAtsMasterSpecifications';
import {
  buildDynamicAtsSpec,
  saveAtsPayload,
  getAtsPayload
} from '../services/netaAtsDataService';

export interface CustomerItem {
  id: string;
  name: string;
  factories?: string[];
  email?: string;
  phone?: string;
  address?: string;
  createdAt?: string;
  updatedAt?: string;
}

export interface EquipmentItem {
  id: string;
  name: string;
  type: string;
  customer?: string;
  customerId?: string;
  factory?: string;
  location?: string;
  status?: 'healthy' | 'warning' | 'critical';
  health?: number;
  lastCheck?: string;
  nameplate?: any;
  technicalSpecs?: Record<string, any>;
  createdAt?: string;
  updatedAt?: string;
}

export interface WorkOrderItem {
  id: string;
  workPermitId: string;
  title: string;
  description: string;
  equipmentId: string | string[];
  equipmentName?: string;
  failureCode: string;
  customer: string;
  customerId?: string;
  factory?: string;
  type: string;
  isUnplanned: boolean;
  priority: string;
  status: string;
  assignedTo: string;
  responsibleApprove: string;
  responsibleDo: string;
  blockingRequired: boolean;
  dueDate: string;
  usedMaterials: any[] | string;
  createdAt: string;
  updatedAt: string;
  downtimeStart: string;
  repairStart: string;
  repairEnd: string;
  restartTime: string;
}

interface Props {
  customers: CustomerItem[];
  allEquipment: EquipmentItem[];
  workOrders: WorkOrderItem[];
  isGoogleConnected: boolean;
  onConnectGoogle: () => void;
  onSaveCustomer: (customer: CustomerItem) => Promise<void>;
  onSaveEquipment: (equipment: EquipmentItem) => Promise<void>;
  onSaveWorkOrder: (wo: WorkOrderItem, syncToSheetsImmediately?: boolean) => Promise<void>;
  onSaveReport?: (reportData: any) => Promise<void> | void;
  onSyncAllToSheets: () => Promise<void>;
  onFetchFromSheets: () => Promise<void>;
  isSyncing: boolean;
  userEmail?: string;
  onSyncAppToSheetsAndFirestore?: () => Promise<void>;
  onSyncSheetsToFirestoreAndApp?: () => Promise<void>;
}

export const FAILURE_CODES = [
  { code: 'ELEC-01', desc: 'Suy giảm điện trở cách điện (Low Insulation Resistance)' },
  { code: 'ELEC-02', desc: 'Phóng điện cục bộ / Quá dòng nhị thứ (Partial Discharge / Overcurrent)' },
  { code: 'ELEC-03', desc: 'Tổn hao điện môi Tan-Delta / Tip-up vượt ngưỡng' },
  { code: 'MECH-01', desc: 'Kẹt cơ cấu đóng cắt / lò xo / cần thao tác (Mechanical Jam)' },
  { code: 'MECH-02', desc: 'Độ rung vượt chuẩn ISO 10816 (High Bearing/Shaft Vibration)' },
  { code: 'THERM-01', desc: 'Phát nóng cục bộ đầu nối / busbar (Hotspot IR Inspection)' },
  { code: 'THERM-02', desc: 'Quá nhiệt cuộn dây Stator / MBA (Winding Overheating)' },
  { code: 'LEAK-01', desc: 'Rò rỉ dầu cách điện / áp suất khí SF6 tụt (Oil / SF6 Leakage)' },
  { code: 'TRIP-01', desc: 'Cắt bảo vệ rơle / Trip Unit (Relay / Electronic Trip)' },
  { code: 'DGA-01', desc: 'Khí cháy hòa tan DGA gia tăng (Acetylene C2H2 / Arc Flash)' },
  { code: 'GEN-01', desc: 'Mất kích từ / Sai lệch điện áp tần số máy phát (AVR / Excitation)' },
  { code: 'PM-ROUTINE', desc: 'Bảo trì định kỳ / Thí nghiệm kiểm định theo kế hoạch' },
  { code: 'OTHER-99', desc: 'Nguyên nhân khác (Ghi rõ trong phần mô tả)' }
];

export const WORK_TYPES = [
  'Bảo trì định kỳ (PM)',
  'Sửa chữa đột xuất (CM)',
  'Kiểm định thử nghiệm (Testing)',
  'Cải tạo nâng cấp',
  'Xử lý sự cố khẩn cấp'
];

export const WORK_STATUSES = [
  'Mới tạo',
  'Đang thực hiện',
  'Chờ phê duyệt',
  'Hoàn thành',
  'Tạm dừng',
  'Hủy'
];

export const PRIORITIES = [
  { value: 'urgent', label: 'Khẩn cấp (P1)' },
  { value: 'high', label: 'Cao (P2)' },
  { value: 'medium', label: 'Trung bình (P3)' },
  { value: 'low', label: 'Thấp (P4)' }
];

/**
 * Tự động tính toán mã WO tiếp theo (WO-YYYY-STT) và WP tiếp theo (WP-YYYY-STT)
 * theo đúng định dạng chuẩn, ví dụ: WO-2026-001, WO-2026-002, WO-2026-003...
 */
export const generateNextWoAndWpIds = (orders: { id?: string; workPermitId?: string }[] = []) => {
  const currentYear = new Date().getFullYear();
  let maxWoSeq = 0;
  let maxWpSeq = 0;

  (orders || []).forEach(item => {
    // 1. Phân tích mã WO: WO-2026-001, WO-001, WO-2026-1, etc.
    if (item.id) {
      const str = String(item.id).trim();
      const match = str.match(/^WO-(?:(\d{4})-)?(\d+)$/i);
      if (match) {
        const yr = match[1] ? parseInt(match[1], 10) : currentYear;
        const seq = parseInt(match[2], 10);
        if (!isNaN(seq) && (yr === currentYear || !match[1])) {
          if (seq < 10000 && seq > maxWoSeq) {
            maxWoSeq = seq;
          }
        }
      } else {
        const fallback = str.match(/(\d+)$/);
        if (fallback) {
          const seq = parseInt(fallback[1], 10);
          if (!isNaN(seq) && seq < 10000 && seq > maxWoSeq) {
            maxWoSeq = seq;
          }
        }
      }
    }

    // 2. Phân tích mã WP: WP-2026-001, WP-001, etc.
    const wp = item.workPermitId;
    if (wp) {
      const str = String(wp).trim();
      const match = str.match(/^WP-(?:(\d{4})-)?(\d+)$/i);
      if (match) {
        const yr = match[1] ? parseInt(match[1], 10) : currentYear;
        const seq = parseInt(match[2], 10);
        if (!isNaN(seq) && (yr === currentYear || !match[1])) {
          if (seq < 10000 && seq > maxWpSeq) {
            maxWpSeq = seq;
          }
        }
      } else {
        const fallback = str.match(/(\d+)$/);
        if (fallback) {
          const seq = parseInt(fallback[1], 10);
          if (!isNaN(seq) && seq < 10000 && seq > maxWpSeq) {
            maxWpSeq = seq;
          }
        }
      }
    }
  });

  const nextSeq = Math.max(maxWoSeq, maxWpSeq) + 1;
  const sttStr = String(nextSeq).padStart(3, '0');

  return {
    nextWoId: `WO-${currentYear}-${sttStr}`,
    nextWpId: `WP-${currentYear}-${sttStr}`,
    nextSeq
  };
};

export const FseFieldDataEntryStation: React.FC<Props> = ({
  customers,
  allEquipment,
  workOrders,
  isGoogleConnected,
  onConnectGoogle,
  onSaveCustomer,
  onSaveEquipment,
  onSaveWorkOrder,
  onSaveReport,
  onSyncAllToSheets,
  onFetchFromSheets,
  isSyncing,
  userEmail = 'sgm1707@gmail.com',
  onSyncAppToSheetsAndFirestore,
  onSyncSheetsToFirestoreAndApp
}) => {
  // Sync status & config
  const [spreadsheetInput, setSpreadsheetInput] = useState(() => {
    return localStorage.getItem('tev_spreadsheet_id') || '1k8m5x9P_TEV_Service_Flatform';
  });
  const [showConfig, setShowConfig] = useState(false);
  const [syncStatusMsg, setSyncStatusMsg] = useState<string | null>(null);

  // Local reactive list of work orders so UI updates IMMEDIATELY when saved
  const [localWorkOrders, setLocalWorkOrders] = useState<WorkOrderItem[]>(() => workOrders || []);
  const [highlightedWoId, setHighlightedWoId] = useState<string | null>(null);
  const [tableSearchQuery, setTableSearchQuery] = useState<string>('');
  const [tableStatusFilter, setTableStatusFilter] = useState<string>('all');
  const [tablePageSize, setTablePageSize] = useState<number>(25);

  useEffect(() => {
    if (workOrders && workOrders.length > 0) {
      setLocalWorkOrders(workOrders);
    }
  }, [workOrders]);

  // Unified displayOrders: Combines localWorkOrders & workOrders, filters and sorts newest first
  const displayOrders = useMemo(() => {
    const map = new Map<string, WorkOrderItem>();
    (workOrders || []).forEach(w => { if (w.id) map.set(w.id.toLowerCase(), w); });
    (localWorkOrders || []).forEach(w => { if (w.id) map.set(w.id.toLowerCase(), w); });
    let list = Array.from(map.values());

    if (tableStatusFilter !== 'all') {
      list = list.filter(w => (w.status || '').toLowerCase().includes(tableStatusFilter.toLowerCase()));
    }

    if (tableSearchQuery.trim()) {
      const q = tableSearchQuery.toLowerCase().trim();
      list = list.filter(w =>
        (w.id && w.id.toLowerCase().includes(q)) ||
        (w.workPermitId && w.workPermitId.toLowerCase().includes(q)) ||
        (w.title && w.title.toLowerCase().includes(q)) ||
        (w.customer && w.customer.toLowerCase().includes(q)) ||
        (w.failureCode && w.failureCode.toLowerCase().includes(q)) ||
        (w.assignedTo && w.assignedTo.toLowerCase().includes(q)) ||
        (w.equipmentId && String(w.equipmentId).toLowerCase().includes(q))
      );
    }

    return list.sort((a, b) => {
      const timeA = parseVersionTimestamp(a.updatedAt || a.createdAt);
      const timeB = parseVersionTimestamp(b.updatedAt || b.createdAt);
      if (timeB !== timeA) return timeB - timeA;
      return (b.id || '').localeCompare(a.id || '');
    });
  }, [localWorkOrders, workOrders, tableSearchQuery, tableStatusFilter]);

  // Step 1: Customer Selection State
  const [selectedCustomerId, setSelectedCustomerId] = useState<string>('');
  const [customerSearch, setCustomerSearch] = useState('');
  const [showCustomerDropdown, setShowCustomerDropdown] = useState(false);
  const [showNewCustomerModal, setShowNewCustomerModal] = useState(false);
  const [newCustomer, setNewCustomer] = useState({
    id: '',
    name: '',
    factory: '',
    phone: '',
    email: '',
    address: ''
  });

  // Project Pre-allocation State (Sheet 'duan')
  const [projectsList, setProjectsList] = useState<ProjectMasterRecord[]>(() => getRegisteredProjects());
  const [selectedProjectId, setSelectedProjectId] = useState<string>('');
  const [projectSearch, setProjectSearch] = useState<string>('');
  const [showProjectDropdown, setShowProjectDropdown] = useState<boolean>(false);
  const [showMasterDirectoryModal, setShowMasterDirectoryModal] = useState<boolean>(false);

  // Step 2: Equipment Selection State (Now supports 'project' with KKS codes)
  const [equipmentSelectMode, setEquipmentSelectMode] = useState<'project' | 'existing' | 'scan' | 'manual'>('project');
  const [selectedEquipmentId, setSelectedEquipmentId] = useState<string>('');
  const [equipmentSearch, setEquipmentSearch] = useState('');
  const [showEquipmentDropdown, setShowEquipmentDropdown] = useState(false);
  const [showScannerModal, setShowScannerModal] = useState(false);
  const [manualEquipment, setManualEquipment] = useState({
    id: '',
    name: '',
    type: 'Máy cắt',
    factory: '',
    location: '',
    manufacturer: '',
    model: '',
    serial: '',
    year: new Date().getFullYear().toString()
  });

  // Selected Equipment Type (for dynamic specs)
  const [currentEqType, setCurrentEqType] = useState<string>('Máy cắt');

  // Modal to open full NETA ATS 23-spec Universal Report / Checklist
  const [showNetaUniversalModal, setShowNetaUniversalModal] = useState<boolean>(false);

  // Helper to normalize equipment type to the standard categories
  const resolveEquipmentType = (type?: string): string => {
    if (!type) return 'Máy cắt';
    const t = type.toLowerCase();
    if (t.includes('cáp') || t.includes('cable') || t.includes('feeder') || t.includes('đoạn cáp')) return 'Cáp điện';
    if (t.includes('tiếp địa') || t.includes('ground') || t.includes('earth') || t.includes('nối đất')) return 'Tiếp địa';
    if (t.includes('rơ le') || t.includes('relay') || t.includes('bảo vệ')) return 'Rơ le bảo vệ';
    if (t.includes('pin') || t.includes('ups') || t.includes('battery') || t.includes('ắc quy') || t.includes('bess')) return 'Pin & UPS';
    if (t.includes('biến áp') || t.includes('transformer') || t.includes('mba')) return 'Máy biến áp';
    if (t.includes('cắt') || t.includes('breaker') || t.includes('vcb') || t.includes('acb') || t.includes('mcb')) return 'Máy cắt';
    if (t.includes('động cơ') || t.includes('motor')) return 'Động cơ';
    if (t.includes('phát') || t.includes('generator') || t.includes('genset')) return 'Máy phát';
    if (t.includes('tủ') || t.includes('switchgear') || t.includes('rmu') || t.includes('ngăn lộ')) return 'Tủ điện';
    return type;
  };

  // Helper to map equipment category to official NETA ATS specification ID
  const getSpecIdForType = (type: string): string => {
    const t = (type || '').toLowerCase();
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
    return 'neta-cable';
  };

  // Step 3: Dynamic Technical Specs State (Full NETA ATS categories)
  const [techSpecs, setTechSpecs] = useState<Record<string, any>>({
    // Breaker specs
    ratedVoltageKv: '24',
    ratedCurrentA: '630',
    shortCircuitKa: '25',
    contactResistanceR: '38.5',
    contactResistanceY: '39.0',
    contactResistanceB: '38.8',
    closeTimeMs: '42.0',
    openTimeMs: '28.5',
    insulationIrMo: '2500',
    controlCircuitIrMo: '150',
    sf6GasPressureBar: '6.0',
    tripUnitLsig: 'Đạt (Pass)',
    springChargeMechanism: 'Bình thường',

    // Motor specs
    ratedPowerKw: '160',
    motorVoltageV: '400',
    motorCurrentA: '290',
    motorRpm: '1485',
    vibrationDeMmS: '1.8',
    vibrationNdeMmS: '1.6',
    statorTempC: '72',
    bearingTempC: '65',
    phaseImbalancePct: '0.8',
    motorIrMo: '1200',
    motorPi: '2.8',
    motorDd: '1.4',
    motorTanDeltaPct: '1.5',

    // Generator specs
    genPowerKva: '1500',
    genVoltageV: '400',
    genCurrentA: '2165',
    genCosPhi: '0.8',
    genHz: '50.0',
    genRpm: '1500',
    genIrStatorMo: '3500',
    genIrRotorMo: '1800',
    genPi: '3.1',
    genStartSec: '7.2',
    genLoadBankPct: '100',
    genOilPressureBar: '4.5',
    genWaterTempC: '82',
    genOverspeedTrip: 'Cắt 115% đạt',

    // Switchgear specs
    swgVoltageKv: '24',
    swgBusCurrentA: '2500',
    swgBusResistanceR: '22.0',
    swgBusResistanceY: '21.5',
    swgBusResistanceB: '22.3',
    swgTevDbmv: '12',
    swgUltrasonicDbuv: '8',
    swgPulsePps: '15',
    swgControlIrMo: '250',
    swgInterlockStatus: 'Tốt / Đầy đủ',

    // Transformer specs
    trfPowerKva: '2000',
    trfPriVoltageKv: '22',
    trfSecVoltageV: '400',
    trfVectorGroup: 'Dyn11',
    trfOilTempC: '58',
    trfWindingTempC: '68',
    trfIrHighLowMo: '4500',
    trfIrHighEarthMo: '3800',
    trfDgaH2: '15',
    trfDgaCh4: '8',
    trfDgaC2h2: '0',
    trfDgaC2h4: '4',
    trfDgaCo: '210',
    trfDgaCo2: '1800',
    trfBreakdownKv: '68',
    trfMoisturePpm: '12',

    // Cable (Cáp điện) specs - ANSI/NETA ATS 7.3 (LV / MV & DC Solar Inverter)
    cableVoltageRating: '1000V DC / 0.6-1kV',
    cableSizeMm2: '1x6mm2 DC Solar Feeder',
    cableLengthM: '150',
    cableTestVoltageKv: '1.0kV DC',
    cableIrPhasePhaseMo: '5000',
    cableIrPhaseEarthMo: '4500',
    cableHipotVoltageKv: '2.5kV DC',
    cableLeakageCurrentUa: '0.8',
    cableShieldContinuity: 'Đạt (Pass)',
    cableConductorResistanceMo: '3.08',
    cableAtsStatus: 'Đạt tiêu chuẩn ANSI/NETA ATS Mục 7.3.2',

    // Grounding (Tiếp địa) specs - ANSI/NETA ATS 7.13 & IEEE 81
    gndSystemType: 'Lưới tiếp địa trạm TBA & Inverter',
    gndTestMethod: 'Fall-of-Potential 3 cực',
    gndResistanceOhm: '0.65',
    gndRequiredMaxOhm: '1.0',
    gndEquipotentialBondingMo: '15',
    gndVisualCondition: 'Tốt (Mối hàn hóa nhiệt chuẩn, không ăn mòn)',
    gndAtsStatus: 'Đạt tiêu chuẩn ANSI/NETA ATS Mục 7.13',

    // Battery & UPS (Pin & UPS) specs - ANSI/NETA ATS 7.18 & 7.22
    batSystemType: '110V DC Giàn Ắc quy VRLA AGM',
    batCapacityAh: '200Ah',
    batTotalVoltageV: '124.5',
    batFloatCurrentA: '0.45',
    batCellInternalResMo: '0.85',
    batMinCellVoltageV: '2.24',
    batMaxCellVoltageV: '2.26',
    batLoadTestDurationMin: '60',
    batAtsStatus: 'Đạt tiêu chuẩn ANSI/NETA ATS Mục 7.18',

    // Protective Relay (Rơ le bảo vệ) specs - ANSI/NETA ATS 7.9
    relayModel: 'SEL-751 / Schneider Micom P122',
    relayFunctionCodes: '50/51, 50N/51N, 87, 27/59',
    relayPickupCurrentA: '5.02',
    relayTripTimeMs: '120',
    relayErrorPct: '0.4%',
    relayTripContactTest: 'Đóng cắt tốt / Trip Test OK',
    relayAtsStatus: 'Đạt tiêu chuẩn ANSI/NETA ATS Mục 7.9',

    // Inverter specs (Solar & BESS)
    invRatedPowerKw: '250kW',
    invDcInputVoltageV: '850V DC',
    invAcOutputVoltageV: '400V AC',
    invThdCurrentPct: '1.8%',
    invAntiIslanding: 'Tác động ngắt < 100ms đạt'
  });

  // Active NETA specification ID based on selected equipment type
  const activeNetaSpecId = useMemo(() => {
    return getSpecIdForType(currentEqType);
  }, [currentEqType]);

  // Dynamic NETA ATS specification definition
  const activeNetaSpec = useMemo(() => {
    return buildDynamicAtsSpec(activeNetaSpecId);
  }, [activeNetaSpecId, techSpecs]);

  // Real-time editable evaluation items for the active NETA form
  const [netaFormItems, setNetaFormItems] = useState<NetaAtsTestItem[]>([]);

  // Update NETA form items whenever the spec or equipment changes
  useEffect(() => {
    if (activeNetaSpec && activeNetaSpec.evaluationItems) {
      setNetaFormItems(activeNetaSpec.evaluationItems);
    }
  }, [activeNetaSpecId]);

  // Step 4: Work Order (WO) Form State - Exactly 23 Fields
  const [woData, setWoData] = useState<Partial<WorkOrderItem>>(() => {
    const { nextWoId, nextWpId } = generateNextWoAndWpIds(workOrders || []);
    return {
      id: nextWoId,
      workPermitId: nextWpId,
      title: 'Kiểm định thử nghiệm & bảo trì thiết bị tại hiện trường',
      description: 'Thực hiện kiểm tra điện trở tiếp xúc, đo cách điện, kiểm định thông số kỹ thuật theo quy trình NETA ATS.',
      equipmentId: '',
      equipmentName: '',
      failureCode: 'PM-ROUTINE',
      customer: '',
      customerId: '',
      factory: '',
      type: 'Kiểm định thử nghiệm (Testing)',
      isUnplanned: false,
      priority: 'medium',
      status: 'Mới tạo',
      assignedTo: userEmail || 'FSE Onsite',
      responsibleApprove: 'FSE Lead / Trưởng ca',
      responsibleDo: 'Kỹ sư hiện trường (FSE Onsite)',
      blockingRequired: true,
      dueDate: new Date(Date.now() + 86400000 * 3).toISOString().split('T')[0],
      usedMaterials: 'Hóa chất vệ sinh tiếp điểm CRC, Giấy đo cách điện, Băng keo hạ thế 3M',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      downtimeStart: '',
      repairStart: '',
      repairEnd: '',
      restartTime: ''
    };
  });

  // Automatically ensure WO-YYYY-STT and WP-YYYY-STT match current next STT
  useEffect(() => {
    if (localWorkOrders.length > 0) {
      setWoData(prev => {
        // If empty or matches generic prefix
        if (!prev.id || prev.id.startsWith('WO-')) {
          const { nextWoId, nextWpId } = generateNextWoAndWpIds(localWorkOrders);
          return {
            ...prev,
            id: nextWoId,
            workPermitId: nextWpId
          };
        }
        return prev;
      });
    }
  }, [localWorkOrders.length]);

  const handleAutoGenerateNextIds = () => {
    const { nextWoId, nextWpId } = generateNextWoAndWpIds(localWorkOrders);
    setWoData(prev => ({
      ...prev,
      id: nextWoId,
      workPermitId: nextWpId
    }));
    setSyncStatusMsg(`Đã cập nhật mã tiếp theo: ${nextWoId} & ${nextWpId}`);
    setTimeout(() => setSyncStatusMsg(null), 3000);
  };

  const [savingLoading, setSavingLoading] = useState(false);
  const [saveSuccessNotice, setSaveSuccessNotice] = useState<string | null>(null);

  // Sync selected customer with WO and equipment filter
  const currentCustomer = useMemo(() => {
    return customers.find(c => c.id === selectedCustomerId) || null;
  }, [customers, selectedCustomerId]);

  // Filter equipment based on selected customer
  const filteredEquipment = useMemo(() => {
    if (!selectedCustomerId) return allEquipment;
    const cust = customers.find(c => c.id === selectedCustomerId);
    return allEquipment.filter(e => 
      e.customerId === selectedCustomerId || 
      (cust && e.customer && e.customer.toLowerCase().includes(cust.name.toLowerCase()))
    );
  }, [allEquipment, selectedCustomerId, customers]);

  // Selected Equipment Object
  const currentEquipment = useMemo(() => {
    return allEquipment.find(e => e.id === selectedEquipmentId) || null;
  }, [allEquipment, selectedEquipmentId]);

  // Update WO customer when customer changes
  useEffect(() => {
    if (currentCustomer) {
      setWoData(prev => ({
        ...prev,
        customer: currentCustomer.name,
        customerId: currentCustomer.id,
        factory: (currentCustomer.factories && currentCustomer.factories[0]) || prev.factory || ''
      }));
    }
  }, [currentCustomer]);

  // Update WO equipment and type when equipment changes
  useEffect(() => {
    if (currentEquipment) {
      setSelectedEquipmentId(currentEquipment.id);
      const resolved = resolveEquipmentType(currentEquipment.type);
      setCurrentEqType(resolved);
      setWoData(prev => ({
        ...prev,
        equipmentId: currentEquipment.id,
        equipmentName: currentEquipment.name,
        factory: currentEquipment.factory || prev.factory || ''
      }));
      if (currentEquipment.technicalSpecs) {
        setTechSpecs(prev => ({ ...prev, ...currentEquipment.technicalSpecs }));
      }
    }
  }, [currentEquipment]);

  // Sync projects list across app
  useEffect(() => {
    const handlePrjUpdate = () => setProjectsList(getRegisteredProjects());
    window.addEventListener(EVT_PROJECTS_UPDATED, handlePrjUpdate);
    return () => window.removeEventListener(EVT_PROJECTS_UPDATED, handlePrjUpdate);
  }, []);

  // Handle direct navigation to NETA Field Form from Project Sheet / Master Directory
  const handleNavigateToNetaFromProject = (
    project: ProjectMasterRecord,
    equipment: ProjectEquipmentItem,
    standard: 'ats' | 'mts' = 'ats'
  ) => {
    // 1. Select the project
    setSelectedProjectId(project.id);
    setProjectSearch('');
    setShowProjectDropdown(false);

    // Auto-select Customer from project
    if (project.customerId) {
      setSelectedCustomerId(project.customerId);
    }

    // 2. Resolve equipment and category
    const resolved = resolveEquipmentType(equipment.type);
    setCurrentEqType(resolved);
    setSelectedEquipmentId(equipment.kksCode || equipment.id);
    setEquipmentSelectMode('project');

    // 3. Populate Work Order
    const stdLabel = standard === 'mts' ? 'ANSI/NETA MTS-2023 (MST)' : 'ANSI/NETA ATS-2025 (ATS)';
    setWoData(prev => ({
      ...prev,
      projectId: project.id,
      customer: project.customerName,
      customerId: project.customerId,
      factory: project.factory || prev.factory || project.customerName,
      equipmentId: equipment.kksCode || equipment.id,
      equipmentName: equipment.kksCode ? `${equipment.kksCode} - ${equipment.name}` : equipment.name,
      equipmentType: resolved,
      substation: equipment.location || prev.substation,
      assignedFse: project.leadFse || prev.assignedFse,
      title: `Kiểm định & thử nghiệm ${standard === 'mts' ? 'NETA MST' : 'NETA ATS'} - ${equipment.name} (${equipment.kksCode})`,
      description: `Dự án: [${project.id}] ${project.name}. Thiết bị: [${equipment.kksCode}] ${equipment.name} (${equipment.type}). Tiêu chuẩn kiểm định: ${stdLabel} tại ${project.factory || project.customerName}.`
    }));

    // 4. Update technical specification defaults
    if (equipment.specs) {
      setTechSpecs(prev => ({
        ...prev,
        rawSpecs: equipment.specs,
        cableVoltageRating: resolved === 'Cáp điện' ? (equipment.specs || prev.cableVoltageRating || '22kV 3x240mm2 Cu/XLPE/PVC') : prev.cableVoltageRating
      }));
    }

    // 5. Close master directory modal
    setShowMasterDirectoryModal(false);
    setProjectsList(getRegisteredProjects());

    // 6. Navigate directly to top-level NETA ATS/MTS Form via App.tsx
    window.dispatchEvent(
      new CustomEvent(EVT_NAVIGATE_TO_NETA_FORM, {
        detail: { project, equipment, standard }
      })
    );
  };

  // Listen to cross-modal navigation events
  useEffect(() => {
    const handleSelectWo = (e: any) => {
      if (e.detail) {
        handleSelectProject(e.detail);
        setShowMasterDirectoryModal(false);
      }
    };
    window.addEventListener('tev_select_project_for_wo', handleSelectWo);
    return () => {
      window.removeEventListener('tev_select_project_for_wo', handleSelectWo);
    };
  }, [projectsList]);

  const currentProject = projectsList.find(p => p.id === selectedProjectId);

  // Handle project selection from site
  const handleSelectProject = (prj: ProjectMasterRecord) => {
    setSelectedProjectId(prj.id);
    setProjectSearch('');
    setShowProjectDropdown(false);

    // Auto-select Customer from project
    if (prj.customerId) {
      setSelectedCustomerId(prj.customerId);
    }

    // Auto-populate Work Order fields
    setWoData(prev => ({
      ...prev,
      customer: prj.customerName,
      customerId: prj.customerId,
      factory: prj.factory || prev.factory || '',
      title: `Kiểm định & thử nghiệm NETA ATS - ${prj.name}`,
      description: `Dự án: [${prj.id}] ${prj.name}. Thực hiện thử nghiệm thiết bị điện theo tiêu chuẩn ANSI/NETA ATS-2025 tại ${prj.factory}.`
    }));

    // Switch equipment tab to 'project'
    setEquipmentSelectMode('project');
  };

  // Handle single equipment pick from project (with KKS code)
  const handleSelectProjectEquipment = (item: ProjectEquipmentItem) => {
    setSelectedEquipmentId(item.kksCode);
    const resolved = resolveEquipmentType(item.type);
    setCurrentEqType(resolved);
    setWoData(prev => ({
      ...prev,
      equipmentId: item.kksCode,
      equipmentName: item.name,
      factory: currentProject?.factory || prev.factory || ''
    }));

    if (item.specs) {
      setTechSpecs(prev => ({ ...prev, rawSpecs: item.specs }));
    }
  };

  // Generate next customer IDs strictly adhering to KHXXX
  const handleGenerateNewCustomer = () => {
    const genId = getNextCustomerId(customers as any);
    setNewCustomer({
      id: genId,
      name: '',
      factory: '',
      phone: '',
      email: '',
      address: ''
    });
    setShowNewCustomerModal(true);
  };

  const handleSaveNewCustomer = async () => {
    if (!newCustomer.name) {
      alert('Vui lòng nhập tên khách hàng');
      return;
    }
    let cleanId = newCustomer.id.trim();
    if (!cleanId) {
      cleanId = getNextCustomerId(customers as any);
    } else {
      cleanId = cleanId.toUpperCase();
      if (!cleanId.startsWith('KH')) {
        const digits = cleanId.replace(/\D/g, '');
        cleanId = `KH${(digits || '1').padStart(3, '0')}`;
      }
    }

    const customerItem: CustomerItem = {
      id: cleanId,
      name: newCustomer.name,
      factories: newCustomer.factory ? [newCustomer.factory] : [],
      phone: newCustomer.phone,
      email: newCustomer.email,
      address: newCustomer.address,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    try {
      await onSaveCustomer(customerItem);
      setSelectedCustomerId(customerItem.id);
      setShowNewCustomerModal(false);
      setSyncStatusMsg(`Đã tạo khách hàng "${customerItem.name}" (${customerItem.id}) chuẩn KHXXX và đồng bộ 2 chiều!`);
      setTimeout(() => setSyncStatusMsg(null), 4000);
    } catch (e: any) {
      alert('Lỗi tạo khách hàng: ' + e.message);
    }
  };

  const handleSaveManualEquipment = async () => {
    if (!manualEquipment.name || !manualEquipment.id) {
      alert('Vui lòng nhập Mã và Tên thiết bị');
      return;
    }
    const eqItem: EquipmentItem = {
      id: manualEquipment.id,
      name: manualEquipment.name,
      type: manualEquipment.type,
      customerId: selectedCustomerId || undefined,
      customer: currentCustomer ? currentCustomer.name : undefined,
      factory: manualEquipment.factory || (currentCustomer?.factories?.[0] || ''),
      location: manualEquipment.location,
      status: 'healthy',
      health: 100,
      nameplate: {
        manufacturer: manualEquipment.manufacturer,
        model: manualEquipment.model,
        serial: manualEquipment.serial,
        year: manualEquipment.year
      },
      technicalSpecs: { ...techSpecs },
      lastCheck: new Date().toLocaleDateString('vi-VN'),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    try {
      await onSaveEquipment(eqItem);
      setSelectedEquipmentId(eqItem.id);
      setCurrentEqType(eqItem.type);
      setEquipmentSelectMode('existing');
      setSyncStatusMsg(`Đã lưu thiết bị "${eqItem.name}" (${eqItem.id}) và đồng bộ lên sheet thietbi!`);
      setTimeout(() => setSyncStatusMsg(null), 4000);
    } catch (e: any) {
      alert('Lỗi lưu thiết bị: ' + e.message);
    }
  };

  // Handle OCR scan completion
  const handleApplyScannedNameplate = async (extracted: ExtractedNameplateData) => {
    setShowScannerModal(false);
    let matchedType = 'Máy biến áp';
    const rawType = (extracted.equipment_type || '').toLowerCase();
    if (rawType.includes('breaker') || rawType.includes('cắt') || rawType.includes('vcb') || rawType.includes('acb')) {
      matchedType = 'Máy cắt';
    } else if (rawType.includes('motor') || rawType.includes('động cơ') || rawType.includes('bơm')) {
      matchedType = 'Động cơ';
    } else if (rawType.includes('generator') || rawType.includes('phát') || rawType.includes('tổ máy')) {
      matchedType = 'Máy phát';
    } else if (rawType.includes('switchgear') || rawType.includes('tủ điện') || rawType.includes('rmu')) {
      matchedType = 'Tủ điện';
    }

    const newId = extracted.equipment_tag || `EQ-SCAN-${Date.now().toString().slice(-4)}`;
    const newName = extracted.equipment_name || `${extracted.manufacturer || 'OEM'} ${extracted.equipment_type || 'Thiết bị'}`;

    // Update specs with extracted data
    setTechSpecs(prev => ({
      ...prev,
      ratedVoltageKv: extracted.primary_voltage_kv ? String(extracted.primary_voltage_kv) : prev.ratedVoltageKv,
      ratedPowerKw: extracted.rated_power_kva ? String(extracted.rated_power_kva) : prev.ratedPowerKw,
      trfPowerKva: extracted.rated_power_kva ? String(extracted.rated_power_kva) : prev.trfPowerKva,
      trfPriVoltageKv: extracted.primary_voltage_kv ? String(extracted.primary_voltage_kv) : prev.trfPriVoltageKv,
      trfSecVoltageV: extracted.secondary_voltage_v ? String(extracted.secondary_voltage_v) : prev.trfSecVoltageV,
      trfVectorGroup: extracted.vector_group || prev.trfVectorGroup
    }));

    const scannedEquipment: EquipmentItem = {
      id: newId,
      name: newName,
      type: matchedType,
      customerId: selectedCustomerId || undefined,
      customer: currentCustomer ? currentCustomer.name : undefined,
      factory: currentCustomer?.factories?.[0] || 'Hiện trường Site',
      location: 'Hiện trường (Site nameplate scan)',
      status: 'healthy',
      health: 100,
      nameplate: {
        manufacturer: extracted.manufacturer,
        serial: extracted.serial_number,
        year: extracted.manufacturing_year,
        standards: extracted.standards,
        fullText: extracted.full_nameplate_text
      },
      technicalSpecs: { ...techSpecs },
      lastCheck: new Date().toLocaleDateString('vi-VN'),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    try {
      await onSaveEquipment(scannedEquipment);
      setSelectedEquipmentId(scannedEquipment.id);
      setCurrentEqType(matchedType);
      setSyncStatusMsg(`Quét AI thành công: Bóc tách nameplate "${newName}" -> Đồng bộ lên sheet thietbi!`);
      setTimeout(() => setSyncStatusMsg(null), 5000);
    } catch (e: any) {
      console.error('Save scanned eq error:', e);
    }
  };

  // 1-Click Timestamp Buttons for Work Order
  const recordTimestamp = (field: 'downtimeStart' | 'repairStart' | 'repairEnd' | 'restartTime') => {
    const nowStr = new Date().toLocaleString('vi-VN');
    setWoData(prev => ({
      ...prev,
      [field]: nowStr,
      updatedAt: new Date().toISOString()
    }));
  };

  // Submit Work Order (Syncs to Firestore & 23-column Sheet TEV Service Flatform)
  const handleSaveAndSyncWorkOrder = async () => {
    if (!woData.title) {
      alert('Vui lòng nhập Tiêu đề Work Order');
      return;
    }
    if (!woData.customer) {
      alert('Vui lòng chọn hoặc tạo Khách hàng');
      return;
    }
    if (!woData.equipmentId) {
      alert('Vui lòng chọn hoặc thêm Thiết bị kiểm tra');
      return;
    }

    setSavingLoading(true);
    try {
      const now = new Date().toISOString();
      const { nextWoId, nextWpId } = generateNextWoAndWpIds(localWorkOrders);

      // Resolve customerId and format strictly KHXXX
      let resolvedCustomerId = woData.customerId || '';
      if (!resolvedCustomerId && woData.customer) {
        const found = customers.find(c => 
          c.name.toLowerCase() === woData.customer.toLowerCase() || 
          c.id.toLowerCase() === woData.customer.toLowerCase()
        );
        if (found) resolvedCustomerId = found.id;
      }

      const finalWO: WorkOrderItem = {
        id: (woData.id && /^WO-\d{4}-\d+$/i.test(woData.id)) ? woData.id : nextWoId,
        workPermitId: (woData.workPermitId && /^WP-\d{4}-\d+$/i.test(woData.workPermitId)) ? woData.workPermitId : nextWpId,
        title: woData.title,
        description: woData.description || '',
        equipmentId: woData.equipmentId,
        equipmentName: woData.equipmentName || (currentEquipment?.name || ''),
        failureCode: woData.failureCode || 'PM-ROUTINE',
        customer: woData.customer,
        customerId: resolvedCustomerId,
        factory: woData.factory || '',
        type: woData.type || 'Kiểm định thử nghiệm (Testing)',
        isUnplanned: !!woData.isUnplanned,
        priority: woData.priority || 'medium',
        status: woData.status || 'Mới tạo',
        assignedTo: woData.assignedTo || userEmail,
        responsibleApprove: woData.responsibleApprove || 'FSE Lead',
        responsibleDo: woData.responsibleDo || 'FSE Onsite',
        blockingRequired: !!woData.blockingRequired,
        dueDate: woData.dueDate || new Date().toISOString().split('T')[0],
        usedMaterials: woData.usedMaterials || '',
        createdAt: woData.createdAt || now,
        updatedAt: now,
        downtimeStart: woData.downtimeStart || '',
        repairStart: woData.repairStart || '',
        repairEnd: woData.repairEnd || '',
        restartTime: woData.restartTime || ''
      };

      // Call parent to save to Firestore & sync to Google Sheets (23 columns)
      await onSaveWorkOrder(finalWO, true);

      // Sync to corresponding Equipment Sheet on Google Sheets (Máy cắt, Động cơ, Máy phát, Tủ điện, Máy biến áp, etc.)
      try {
        const tokens = localStorage.getItem('google_tokens');
        const spreadsheetId = localStorage.getItem('tev_spreadsheet_id') || '1k8m5x9P_TEV_Service_Flatform';
        const headers: Record<string, string> = { 'Content-Type': 'application/json' };
        if (tokens) headers['Authorization'] = `Bearer ${tokens}`;
        if (spreadsheetId) headers['x-spreadsheet-id'] = spreadsheetId;

        await fetch('/api/sheets/sync-record', {
          method: 'POST',
          headers,
          body: JSON.stringify({
            recordType: 'equipmentTest',
            spreadsheetId,
            data: {
              ...techSpecs,
              equipmentType: currentEqType,
              woId: finalWO.id,
              equipmentId: finalWO.equipmentId,
              equipmentName: finalWO.equipmentName,
              customer: finalWO.customer,
              location: finalWO.factory,
              testDate: new Date().toLocaleDateString('vi-VN'),
              inspector: finalWO.assignedTo
            }
          })
        });
      } catch (eqSyncErr) {
        console.warn('Equipment sheet sync error:', eqSyncErr);
      }

      // Immediately update localWorkOrders reactive state with newest entry on top
      const updatedList = [finalWO, ...localWorkOrders.filter(w => w.id !== finalWO.id)];
      setLocalWorkOrders(updatedList);
      setHighlightedWoId(finalWO.id);

      setSaveSuccessNotice(`Đã lưu và đồng bộ 2 chiều thành công WO [${finalWO.id}] vào file Google Sheet "TEV Service Flatform" & Sheet thiết bị "${currentEqType}"!`);
      setTimeout(() => setSaveSuccessNotice(null), 6000);

      // Auto generate next WO (WO-YYYY-STT) and WP (WP-YYYY-STT) code for subsequent work
      const subsequent = generateNextWoAndWpIds(updatedList);
      setWoData(prev => ({
        ...prev,
        id: subsequent.nextWoId,
        workPermitId: subsequent.nextWpId,
        downtimeStart: '',
        repairStart: '',
        repairEnd: '',
        restartTime: ''
      }));
    } catch (e: any) {
      alert('Lỗi lưu và đồng bộ: ' + (e.message || String(e)));
    } finally {
      setSavingLoading(false);
    }
  };

  // Initialize and format all equipment sheets at once
  const handleInitAllEquipmentSheets = async () => {
    setSyncStatusMsg('Đang tự động khởi tạo & chuẩn hóa các sheet thiết bị trên file TEV Service Flatform...');
    try {
      const tokens = localStorage.getItem('google_tokens');
      const spreadsheetId = localStorage.getItem('tev_spreadsheet_id') || '1k8m5x9P_TEV_Service_Flatform';
      const headers: Record<string, string> = { 'Content-Type': 'application/json' };
      if (tokens) headers['Authorization'] = `Bearer ${tokens}`;
      if (spreadsheetId) headers['x-spreadsheet-id'] = spreadsheetId;

      const res = await fetch('/api/sheets/init-tev-sheets', {
        method: 'POST',
        headers,
        body: JSON.stringify({ spreadsheetId })
      });
      const json = await res.json();
      if (json.success) {
        setSyncStatusMsg(`Khởi tạo thành công: ${json.sheets?.join(', ')}`);
        setTimeout(() => setSyncStatusMsg(null), 6000);
      } else {
        alert('Lỗi khởi tạo sheet: ' + (json.error || 'Vui lòng kiểm tra quyền Google Drive'));
      }
    } catch (err: any) {
      alert('Lỗi kết nối: ' + err.message);
    }
  };

  const handleSaveSpreadsheetConfig = () => {
    localStorage.setItem('tev_spreadsheet_id', spreadsheetInput);
    setShowConfig(false);
    setSyncStatusMsg('Đã lưu cấu hình Google Sheet ID thành công!');
    setTimeout(() => setSyncStatusMsg(null), 3000);
  };

  return (
    <div className="space-y-6 max-w-6xl mx-auto">
      {/* Nameplate Scanner Modal */}
      <NameplateScannerModal
        isOpen={showScannerModal}
        onClose={() => setShowScannerModal(false)}
        onApplyData={handleApplyScannedNameplate}
      />

      {/* TOP HEADER: 2-WAY SYNC CONTROLLER & STATUS */}
      <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 rounded-2xl p-5 text-white shadow-lg border border-indigo-900/50">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="w-12 h-12 rounded-xl bg-blue-600/30 border border-blue-400/40 flex items-center justify-center text-blue-400 shrink-0">
              <FileSpreadsheet size={26} />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-xs font-black uppercase tracking-wider px-2 py-0.5 rounded bg-blue-500/20 text-blue-300 border border-blue-400/30">
                  TEV SERVICE PLATFORM 2-WAY SYNC
                </span>
                <span className="text-xs text-slate-300 font-medium hidden sm:inline">
                  Chuẩn hóa dữ liệu hiện trường FSE
                </span>
              </div>
              <h1 className="text-lg sm:text-xl font-bold text-white mt-1">
                Trạm Nhập Liệu Hiện Trường & Đồng Bộ Google Sheet 2 Chiều
              </h1>
              <p className="text-xs text-slate-400 mt-0.5">
                Đồng nhất danh mục Khách hàng (<code className="text-amber-300">khachhang</code>), Thiết bị (<code className="text-cyan-300">thietbi</code>) và Phiếu Công Việc 23 cột (<code className="text-emerald-300">TEV Service Flatform</code>).
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2 self-start md:self-center">
            {!isGoogleConnected ? (
              <button
                type="button"
                onClick={onConnectGoogle}
                className="px-3.5 py-2 bg-amber-500 hover:bg-amber-400 text-slate-950 font-bold text-xs rounded-xl flex items-center gap-1.5 shadow-md transition-all active:scale-95"
              >
                <UploadCloud size={16} />
                <span>Kết Nối Google Drive</span>
              </button>
            ) : (
              <div className="flex items-center gap-1 px-3 py-1.5 rounded-xl bg-emerald-500/20 text-emerald-300 border border-emerald-400/30 text-xs font-bold">
                <CheckCircle2 size={15} />
                <span>Drive Đã Kết Nối</span>
              </div>
            )}

            <button
              type="button"
              onClick={handleInitAllEquipmentSheets}
              className="px-3.5 py-2 bg-indigo-600 hover:bg-indigo-500 text-white font-bold text-xs rounded-xl flex items-center gap-1.5 shadow-md transition-all active:scale-95"
              title="Tự động khởi tạo & chuẩn hóa toàn bộ các Sheet Thiết Bị (Máy biến áp, Máy cắt, Động cơ, Máy phát, Tủ điện, Cáp điện, Pin/UPS, Rơle) trên file Google Sheets"
            >
              <Layers size={14} />
              <span>Tạo Các Sheet Thiết Bị</span>
            </button>

            {/* FLOW 1: APP -> GOOGLE SHEET -> FIRESTORE */}
            <button
              type="button"
              onClick={onSyncAppToSheetsAndFirestore || onSyncAllToSheets}
              disabled={isSyncing}
              className="px-3.5 py-2 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-bold text-xs rounded-xl flex items-center gap-1.5 shadow-md transition-all active:scale-95 disabled:opacity-50"
              title="Flow 1: Lưu & đồng bộ từ App lên Google Sheet và lưu vào Firestore"
            >
              <Database size={14} className="text-emerald-200" />
              <span>App ➔ Sheet ➔ Firestore</span>
            </button>

            {/* FLOW 2: GOOGLE SHEET -> FIRESTORE -> APP */}
            <button
              type="button"
              onClick={onSyncSheetsToFirestoreAndApp || onFetchFromSheets}
              disabled={isSyncing}
              className="px-3.5 py-2 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white font-bold text-xs rounded-xl flex items-center gap-1.5 shadow-md transition-all active:scale-95 disabled:opacity-50"
              title="Flow 2: Nạp dữ liệu mới nhất từ Google Sheet, lưu vào Firestore và cập nhật App"
            >
              <DownloadCloud size={14} className="text-blue-200" />
              <span>Sheet ➔ Firestore ➔ App</span>
            </button>

            <button
              type="button"
              onClick={() => setShowConfig(!showConfig)}
              className="p-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl border border-slate-700 transition-colors"
              title="Cấu hình Google Sheet ID"
            >
              <SlidersHorizontal size={15} />
            </button>
          </div>
        </div>

        {/* Sheet Tabs Status Bar */}
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5 mt-4 pt-3 border-t border-slate-800/80 text-xs">
          <div className="flex items-center justify-between px-3 py-2 rounded-xl bg-slate-950/60 border border-slate-800">
            <span className="text-slate-400 flex items-center gap-1.5">
              <span className="w-2 h-2 rounded-full bg-emerald-400"></span>
              Sheet WO (23 Cột):
            </span>
            <span className="font-mono font-bold text-emerald-300">TEV Service Flatform ({workOrders.length})</span>
          </div>
          <div className="flex items-center justify-between px-3 py-2 rounded-xl bg-slate-950/60 border border-slate-800">
            <span className="text-slate-400 flex items-center gap-1.5">
              <span className="w-2 h-2 rounded-full bg-amber-400"></span>
              Sheet Khách Hàng:
            </span>
            <span className="font-mono font-bold text-amber-300">khachhang ({customers.length})</span>
          </div>
          <div className="flex items-center justify-between px-3 py-2 rounded-xl bg-slate-950/60 border border-slate-800">
            <span className="text-slate-400 flex items-center gap-1.5">
              <span className="w-2 h-2 rounded-full bg-cyan-400"></span>
              Sheet Thiết Bị:
            </span>
            <span className="font-mono font-bold text-cyan-300">thietbi ({allEquipment.length})</span>
          </div>
        </div>

        {/* Dynamic Equipment Sheets Badges */}
        <div className="flex flex-wrap items-center gap-2 mt-2 pt-2 border-t border-slate-800/50 text-[11px]">
          <span className="text-slate-400 font-bold">Các Sheet Thiết Bị:</span>
          {['Máy biến áp', 'Máy cắt', 'Động cơ', 'Máy phát', 'Tủ điện', 'Cáp điện', 'Pin & UPS', 'Rơ le bảo vệ'].map(sheetName => (
            <span
              key={sheetName}
              onClick={() => setCurrentEqType(sheetName === 'Pin & UPS' ? 'Pin / BESS' : (sheetName === 'Rơ le bảo vệ' ? 'Rơle' : sheetName))}
              className={`px-2.5 py-1 rounded-lg cursor-pointer transition-all border ${
                currentEqType.includes(sheetName.slice(0, 4))
                  ? 'bg-indigo-600/40 text-indigo-200 border-indigo-400/60 font-bold scale-105'
                  : 'bg-slate-900/60 text-slate-300 border-slate-700/60 hover:bg-slate-800'
              }`}
            >
              📄 {sheetName}
            </span>
          ))}
        </div>

        {/* Sheet Config Popdown */}
        {showConfig && (
          <div className="mt-4 p-4 rounded-xl bg-slate-950 border border-slate-800 animate-in fade-in duration-150">
            <div className="flex items-center justify-between mb-2">
              <span className="text-xs font-bold text-slate-200 flex items-center gap-1.5">
                <FileSpreadsheet size={15} className="text-blue-400" />
                Cấu hình Đường dẫn Google Sheet (Spreadsheet ID / URL)
              </span>
              <button onClick={() => setShowConfig(false)} className="text-slate-500 hover:text-slate-300">
                <X size={15} />
              </button>
            </div>
            <div className="flex gap-2">
              <input
                type="text"
                value={spreadsheetInput}
                onChange={(e) => setSpreadsheetInput(e.target.value)}
                placeholder="Dán link Google Sheet hoặc ID (VD: 1k8m5x9P...)"
                className="flex-1 bg-slate-900 border border-slate-700 text-white rounded-lg px-3 py-2 text-xs font-mono focus:border-blue-500 outline-none"
              />
              <button
                type="button"
                onClick={handleSaveSpreadsheetConfig}
                className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs rounded-lg transition-colors shrink-0"
              >
                Lưu ID
              </button>
            </div>
            <p className="text-[11px] text-slate-400 mt-2">
              Mẹo: Hệ thống tự động bóc tách ID từ đường link <code className="text-blue-300">https://docs.google.com/spreadsheets/d/ID/edit</code>.
            </p>
          </div>
        )}

        {/* Notification Toast */}
        {syncStatusMsg && (
          <div className="mt-3 p-2.5 rounded-xl bg-emerald-500/20 border border-emerald-400/40 text-emerald-200 text-xs font-bold flex items-center gap-2 animate-in slide-in-from-top-1">
            <CheckCircle2 size={16} className="text-emerald-400 shrink-0" />
            <span>{syncStatusMsg}</span>
          </div>
        )}
      </div>

      {/* SUCCESS BANNER WHEN SAVED */}
      {saveSuccessNotice && (
        <div className="p-4 rounded-2xl bg-emerald-600 text-white font-bold text-sm flex items-center justify-between shadow-lg animate-in slide-in-from-top-2">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 rounded-lg bg-white/20 flex items-center justify-center shrink-0">
              <FileCheck size={20} />
            </div>
            <span>{saveSuccessNotice}</span>
          </div>
          <button onClick={() => setSaveSuccessNotice(null)} className="text-white/80 hover:text-white p-1">
            <X size={18} />
          </button>
        </div>
      )}

      {/* 4-STEP FIELD ENTRY FORM */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        
        {/* LEFT COLUMN: STEPS 0, 1 & 2 (PROJECT, CUSTOMER & EQUIPMENT UNIFIED SELECTION) */}
        <div className="lg:col-span-5 space-y-6">

          {/* STEP 0: DỰ ÁN CÔNG TÁC HIỆN TRƯỜNG (AUTO-LOAD KHÁCH HÀNG & THIẾT BỊ KKS) */}
          <div className="bg-gradient-to-r from-indigo-900 via-indigo-950 to-slate-900 rounded-2xl p-4 sm:p-5 border border-indigo-700/60 shadow-md text-white space-y-3">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="w-8 h-8 rounded-xl bg-amber-400 text-slate-950 flex items-center justify-center font-black shadow-xs">
                  <Briefcase size={16} />
                </div>
                <div>
                  <div className="flex items-center gap-1.5">
                    <span className="text-[10px] font-black uppercase tracking-wider px-1.5 py-0.2 rounded bg-amber-400/20 text-amber-300 border border-amber-400/30">
                      Field Project Prep
                    </span>
                    <span className="text-xs text-slate-300 font-medium">Sheet Dự án</span>
                  </div>
                  <h3 className="text-sm font-bold text-white mt-0.5">
                    Chọn Mã Dự Án Hiện Trường
                  </h3>
                </div>
              </div>

              <button
                type="button"
                onClick={() => setShowMasterDirectoryModal(true)}
                className="px-3 py-1.5 bg-indigo-700/80 hover:bg-indigo-600 text-white border border-indigo-500/50 rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all shadow-xs"
                title="Mở bảng quản lý & soạn thiết bị cho dự án mới"
              >
                <FolderPlus size={14} className="text-amber-300" />
                <span>+ Soạn Dự Án</span>
              </button>
            </div>

            {/* Dropdown Project Selector */}
            <div className="relative">
              <div className="relative">
                <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  type="text"
                  placeholder="Nhập hoặc chọn Mã Dự Án (VD: DA-2026-001, DA-2026-002...)..."
                  value={projectSearch || (currentProject ? `${currentProject.id} - ${currentProject.name}` : '')}
                  onChange={(e) => {
                    setProjectSearch(e.target.value);
                    setShowProjectDropdown(true);
                  }}
                  onFocus={() => setShowProjectDropdown(true)}
                  className="w-full pl-9 pr-8 py-2.5 bg-slate-800/90 border border-indigo-500/40 focus:bg-slate-900 focus:border-amber-400 focus:ring-2 focus:ring-amber-400/20 rounded-xl text-xs font-bold text-white outline-none transition-all placeholder:text-slate-400"
                />
                <button
                  type="button"
                  onClick={() => setShowProjectDropdown(!showProjectDropdown)}
                  className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-300 hover:text-white"
                >
                  <ChevronDown size={16} />
                </button>
              </div>

              {/* Project Dropdown Results */}
              {showProjectDropdown && (
                <div className="absolute z-40 w-full mt-1.5 bg-slate-900 border border-indigo-700/60 rounded-xl shadow-2xl max-h-64 overflow-y-auto divide-y divide-slate-800 animate-in fade-in-50 duration-150">
                  {projectsList
                    .filter(p =>
                      !projectSearch ||
                      p.id.toLowerCase().includes(projectSearch.toLowerCase()) ||
                      p.name.toLowerCase().includes(projectSearch.toLowerCase()) ||
                      p.customerName.toLowerCase().includes(projectSearch.toLowerCase())
                    )
                    .map(prj => (
                      <div
                        key={prj.id}
                        onClick={() => handleSelectProject(prj)}
                        className={`p-3 hover:bg-indigo-900/60 cursor-pointer transition-colors flex items-center justify-between text-xs ${
                          selectedProjectId === prj.id ? 'bg-indigo-900/90 font-bold text-amber-300' : 'text-slate-200'
                        }`}
                      >
                        <div>
                          <div className="font-bold flex items-center gap-2">
                            <span className="font-mono text-amber-400">{prj.id}</span>
                            <span className="text-white">{prj.name}</span>
                          </div>
                          <div className="text-[11px] text-slate-400 mt-0.5">
                            Khách hàng: <span className="text-blue-300 font-semibold">{prj.customerId} - {prj.customerName}</span> • Trạm: {prj.factory} • <span className="text-amber-300 font-mono font-semibold">{prj.equipmentList?.length || 0} thiết bị KKS</span>
                          </div>
                        </div>
                        {selectedProjectId === prj.id && (
                          <Check size={16} className="text-amber-400 shrink-0" />
                        )}
                      </div>
                    ))}

                  <div
                    onClick={() => {
                      setShowProjectDropdown(false);
                      setShowMasterDirectoryModal(true);
                    }}
                    className="p-2.5 bg-indigo-950/80 hover:bg-indigo-900 text-amber-300 text-xs font-bold cursor-pointer flex items-center gap-1.5 border-t border-indigo-800"
                  >
                    <Plus size={14} />
                    <span>Tạo dự án mới hoặc soạn danh sách thiết bị KKS...</span>
                  </div>
                </div>
              )}
            </div>

            {/* Active Project Card Preview */}
            {currentProject ? (
              <div className="p-3 rounded-xl bg-indigo-950/70 border border-indigo-600/40 text-xs flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                <div className="space-y-0.5">
                  <div className="text-emerald-400 font-bold flex items-center gap-1.5">
                    <CheckCircle2 size={13} />
                    <span>Đã nạp dự án: <strong>{currentProject.id}</strong> - {currentProject.name}</span>
                  </div>
                  <div className="text-[11px] text-slate-300">
                    Khách hàng: <strong>{currentProject.customerName}</strong> ({currentProject.factory}) • KKS sẵn sàng: <strong className="text-amber-300 font-mono">{currentProject.equipmentList?.length || 0} thiết bị</strong>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={() => {
                    setSelectedProjectId('');
                    setEquipmentSelectMode('existing');
                  }}
                  className="text-[10px] text-slate-400 hover:text-white underline self-end sm:self-center"
                >
                  Bỏ chọn dự án
                </button>
              </div>
            ) : (
              <div className="p-2.5 rounded-xl bg-indigo-950/40 border border-dashed border-indigo-700/40 text-slate-400 text-[11px] flex items-center justify-between">
                <span>Chưa chọn mã dự án. Bạn có thể chọn để tự động load toàn bộ thiết bị KKS.</span>
              </div>
            )}
          </div>

          {/* STEP 1: KHÁCH HÀNG (ĐỒNG NHẤT TỪ FILE SHEET KHACHHANG) */}
          <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <span className="w-7 h-7 rounded-lg bg-amber-500 text-slate-950 font-black text-xs flex items-center justify-center shadow-xs">
                  1
                </span>
                <div>
                  <h3 className="font-bold text-slate-900 text-sm">
                    Khách Hàng (Sheet <code className="text-amber-700 bg-amber-50 px-1 py-0.5 rounded">khachhang</code>)
                  </h3>
                  <p className="text-xs text-slate-500">Chọn danh sách thả có sẵn hoặc tạo mới</p>
                </div>
              </div>

              <button
                type="button"
                onClick={handleGenerateNewCustomer}
                className="px-3 py-1.5 bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-300 rounded-xl text-xs font-bold flex items-center gap-1 transition-all active:scale-95 shadow-2xs"
              >
                <Plus size={14} />
                <span>+ Tạo Khách Hàng</span>
              </button>
            </div>

            {/* Dropdown Customer Selector */}
            <div className="relative">
              <label className="block text-xs font-bold text-slate-700 mb-1">
                Lựa chọn khách hàng hiện trường
              </label>
              <div className="relative">
                <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  type="text"
                  placeholder="Gõ tìm kiếm tên hoặc mã khách hàng..."
                  value={customerSearch || (currentCustomer ? `${currentCustomer.id} - ${currentCustomer.name}` : '')}
                  onChange={(e) => {
                    setCustomerSearch(e.target.value);
                    setShowCustomerDropdown(true);
                  }}
                  onFocus={() => setShowCustomerDropdown(true)}
                  className="w-full pl-9 pr-8 py-2.5 bg-slate-50 border border-slate-300 focus:bg-white focus:border-amber-500 focus:ring-2 focus:ring-amber-200 rounded-xl text-xs font-semibold text-slate-900 outline-none transition-all"
                />
                <button
                  type="button"
                  onClick={() => setShowCustomerDropdown(!showCustomerDropdown)}
                  className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                >
                  <ChevronDown size={16} />
                </button>
              </div>

              {/* Customer Dropdown Results */}
              {showCustomerDropdown && (
                <div className="absolute z-30 w-full mt-1.5 bg-white border border-slate-200 rounded-xl shadow-xl max-h-56 overflow-y-auto divide-y divide-slate-100 animate-in fade-in-50 duration-150">
                  {customers
                    .filter(c => 
                      !customerSearch ||
                      c.name.toLowerCase().includes(customerSearch.toLowerCase()) ||
                      c.id.toLowerCase().includes(customerSearch.toLowerCase())
                    )
                    .map(cust => (
                      <div
                        key={cust.id}
                        onClick={() => {
                          setSelectedCustomerId(cust.id);
                          setCustomerSearch('');
                          setShowCustomerDropdown(false);
                        }}
                        className={`p-3 hover:bg-amber-50/60 cursor-pointer transition-colors flex items-center justify-between ${
                          selectedCustomerId === cust.id ? 'bg-amber-50 font-bold text-amber-900' : 'text-slate-800'
                        }`}
                      >
                        <div>
                          <div className="text-xs font-bold text-slate-900">{cust.name}</div>
                          <div className="text-[11px] text-slate-500 mt-0.5">
                            Mã: <span className="font-mono text-amber-700">{cust.id}</span>
                            {cust.factories && cust.factories.length > 0 && ` • Trạm/Nhà máy: ${cust.factories.join(', ')}`}
                          </div>
                        </div>
                        {selectedCustomerId === cust.id && (
                          <Check size={16} className="text-amber-600 shrink-0" />
                        )}
                      </div>
                    ))}

                  <div
                    onClick={handleGenerateNewCustomer}
                    className="p-3 bg-amber-50/40 hover:bg-amber-100/60 text-amber-800 text-xs font-bold cursor-pointer flex items-center gap-2 border-t border-amber-200"
                  >
                    <Plus size={15} />
                    <span>Thêm khách hàng mới tại hiện trường...</span>
                  </div>
                </div>
              )}
            </div>

            {/* Selected Customer Card Preview */}
            {currentCustomer ? (
              <div className="p-3.5 rounded-xl bg-amber-50/50 border border-amber-200 text-xs space-y-1.5">
                <div className="flex items-center justify-between">
                  <span className="font-black text-slate-900">{currentCustomer.name}</span>
                  <span className="px-2 py-0.5 rounded bg-amber-200/60 text-amber-900 font-mono font-bold text-[10px]">
                    {currentCustomer.id}
                  </span>
                </div>
                <div className="text-slate-600 flex items-center gap-1.5">
                  <Building size={13} className="text-amber-700 shrink-0" />
                  <span>{currentCustomer.factories?.join(', ') || 'Chưa phân nhà máy'}</span>
                </div>
                {currentCustomer.phone && (
                  <div className="text-slate-500 text-[11px]">Hotline: {currentCustomer.phone}</div>
                )}
              </div>
            ) : (
              <div className="p-3 rounded-xl bg-slate-50 border border-dashed border-slate-300 text-slate-400 text-xs text-center">
                Chưa chọn khách hàng. Vui lòng chọn từ danh sách hoặc tạo mới.
              </div>
            )}
          </div>

          {/* STEP 2: THIẾT BỊ KIỂM TRA (SẴN CÓ TỪ SHEET THIETBI / SCAN NAMEPLATE / NHẬP THỦ CÔNG) */}
          <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <span className="w-7 h-7 rounded-lg bg-blue-600 text-white font-black text-xs flex items-center justify-center shadow-xs">
                  2
                </span>
                <div>
                  <h3 className="font-bold text-slate-900 text-sm">
                    Thiết Bị Kiểm Tra (Sheet <code className="text-blue-700 bg-blue-50 px-1 py-0.5 rounded">thietbi</code>)
                  </h3>
                  <p className="text-xs text-slate-500">Sẵn có FSE chuẩn bị trước, Scan nameplate hoặc Nhập tay</p>
                </div>
              </div>
            </div>

            {/* 4 Entry Modes Tabs */}
            <div className="grid grid-cols-4 gap-1 p-1 bg-slate-100 rounded-xl text-xs font-bold text-slate-600">
              <button
                type="button"
                onClick={() => setEquipmentSelectMode('project')}
                className={`py-1.5 px-1.5 rounded-lg transition-all text-center flex items-center justify-center gap-1 ${
                  equipmentSelectMode === 'project'
                    ? 'bg-indigo-600 text-white shadow-xs'
                    : 'hover:text-slate-900 text-indigo-700'
                }`}
                title="Chọn thiết bị & mã KKS từ dự án đã chuẩn bị"
              >
                <Briefcase size={12} />
                <span className="truncate">Dự Án ({currentProject?.equipmentList?.length || 0})</span>
              </button>
              <button
                type="button"
                onClick={() => setEquipmentSelectMode('existing')}
                className={`py-1.5 px-1.5 rounded-lg transition-all text-center ${
                  equipmentSelectMode === 'existing'
                    ? 'bg-white text-blue-700 shadow-xs'
                    : 'hover:text-slate-900'
                }`}
              >
                Master ({filteredEquipment.length})
              </button>
              <button
                type="button"
                onClick={() => setShowScannerModal(true)}
                className="py-1.5 px-1.5 rounded-lg transition-all text-center flex items-center justify-center gap-1 text-amber-700 bg-amber-50 hover:bg-amber-100 border border-amber-200"
              >
                <Camera size={12} />
                <span>Scan</span>
              </button>
              <button
                type="button"
                onClick={() => setEquipmentSelectMode('manual')}
                className={`py-1.5 px-1.5 rounded-lg transition-all text-center ${
                  equipmentSelectMode === 'manual'
                    ? 'bg-white text-blue-700 shadow-xs'
                    : 'hover:text-slate-900'
                }`}
              >
                + Gõ tay
              </button>
            </div>

            {/* MODE 1: SELECT EQUIPMENT FROM ACTIVE PROJECT (KKS CODE PRE-ALLOCATED) */}
            {equipmentSelectMode === 'project' && (
              <div className="space-y-3">
                {currentProject ? (
                  <div className="space-y-2.5">
                    <div className="flex items-center justify-between text-xs">
                      <span className="font-bold text-slate-800 flex items-center gap-1.5">
                        <ListChecks size={14} className="text-indigo-600" />
                        Thiết bị trong dự án {currentProject.id}:
                      </span>
                      <span className="text-[11px] font-mono font-bold text-indigo-700 bg-indigo-50 px-2 py-0.5 rounded">
                        {currentProject.equipmentList?.length || 0} thiết bị KKS
                      </span>
                    </div>

                    {currentProject.equipmentList && currentProject.equipmentList.length > 0 ? (
                      <div className="space-y-2 max-h-72 overflow-y-auto pr-1">
                        {currentProject.equipmentList.map((item, idx) => {
                          const isSelected = (selectedEquipmentId || '').toLowerCase() === item.kksCode.toLowerCase();
                          return (
                            <div
                              key={item.kksCode || idx}
                              onClick={() => handleSelectProjectEquipment(item)}
                              className={`p-3 rounded-xl border cursor-pointer transition-all flex items-start justify-between gap-2 ${
                                isSelected
                                  ? 'bg-indigo-50/90 border-indigo-500 ring-2 ring-indigo-400 text-indigo-950 shadow-xs'
                                  : 'bg-white border-slate-200 hover:border-indigo-300 hover:bg-slate-50'
                              }`}
                            >
                              <div className="min-w-0">
                                <div className="flex items-center gap-1.5">
                                  <span className="font-mono font-black text-xs text-indigo-700">
                                    {item.kksCode}
                                  </span>
                                  <span className="text-[10px] bg-indigo-100 text-indigo-900 font-semibold px-1.5 py-0.2 rounded">
                                    {item.type}
                                  </span>
                                </div>
                                <div className="font-bold text-slate-900 text-xs mt-1">
                                  {item.name}
                                </div>
                                {item.location && (
                                  <div className="text-[11px] text-slate-500 mt-0.5 flex items-center gap-1">
                                    <MapPin size={11} className="text-slate-400 shrink-0" />
                                    <span className="truncate">{item.location}</span>
                                    {item.specs && <span>• {item.specs}</span>}
                                  </div>
                                )}
                              </div>

                              <div className="flex flex-col items-end gap-1 shrink-0">
                                <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${
                                  item.status === 'đã_test'
                                    ? 'bg-emerald-100 text-emerald-800'
                                    : 'bg-amber-100 text-amber-900'
                                }`}>
                                  {item.status === 'đã_test' ? 'Đã test' : 'Chờ test'}
                                </span>
                                {isSelected && (
                                  <div className="flex flex-col items-end gap-1 mt-1">
                                    <span className="text-[10px] font-bold text-indigo-700 flex items-center gap-0.5">
                                      <Check size={12} /> Đang chọn
                                    </span>
                                    <button
                                      type="button"
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        setShowNetaUniversalModal(true);
                                      }}
                                      className="text-[10px] font-bold text-emerald-700 hover:text-emerald-900 bg-emerald-50 px-1.5 py-0.5 rounded border border-emerald-200 flex items-center gap-1 shadow-2xs"
                                      title="Mở Biểu Mẫu Thử Nghiệm NETA ATS"
                                    >
                                      <ShieldCheck size={10} />
                                      <span>Form NETA</span>
                                    </button>
                                  </div>
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    ) : (
                      <div className="p-4 rounded-xl bg-amber-50/60 border border-amber-200 text-xs text-amber-900 space-y-2">
                        <p className="font-medium">Dự án này chưa có thiết bị nào trong danh sách chuẩn bị.</p>
                        <button
                          type="button"
                          onClick={() => setShowMasterDirectoryModal(true)}
                          className="px-3 py-1.5 bg-indigo-700 hover:bg-indigo-800 text-white rounded-lg font-bold flex items-center gap-1.5 shadow-2xs"
                        >
                          <Plus size={13} />
                          <span>Soạn thiết bị & mã KKS cho dự án này</span>
                        </button>
                      </div>
                    )}
                  </div>
                ) : (
                  <div className="p-4 rounded-xl bg-slate-50 border border-dashed border-slate-300 text-xs text-center text-slate-500 space-y-1.5">
                    <p className="font-bold text-slate-700">Chưa chọn Mã Dự Án Hiện Trường</p>
                    <p className="text-[11px] text-slate-500">
                      Vui lòng chọn Mã Dự Án ở khung bước 0 phía trên, hoặc chuyển qua tab <strong>Master</strong> để chọn thiết bị.
                    </p>
                  </div>
                )}
              </div>
            )}

            {/* MODE 2: SELECT FROM EXISTING PRE-PREPARED SHEET THIETBI */}
            {equipmentSelectMode === 'existing' && (
              <div className="space-y-3">
                <div className="relative">
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    Chọn thiết bị FSE đã chuẩn bị trước
                  </label>
                  <div className="relative">
                    <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                    <input
                      type="text"
                      placeholder="Tìm theo mã hoặc tên thiết bị..."
                      value={equipmentSearch || (currentEquipment ? `${currentEquipment.id} - ${currentEquipment.name}` : '')}
                      onChange={(e) => {
                        setEquipmentSearch(e.target.value);
                        setShowEquipmentDropdown(true);
                      }}
                      onFocus={() => setShowEquipmentDropdown(true)}
                      className="w-full pl-9 pr-8 py-2.5 bg-slate-50 border border-slate-300 focus:bg-white focus:border-blue-500 focus:ring-2 focus:ring-blue-200 rounded-xl text-xs font-semibold text-slate-900 outline-none transition-all"
                    />
                    <button
                      type="button"
                      onClick={() => setShowEquipmentDropdown(!showEquipmentDropdown)}
                      className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                    >
                      <ChevronDown size={16} />
                    </button>
                  </div>

                  {showEquipmentDropdown && (
                    <div className="absolute z-30 w-full mt-1.5 bg-white border border-slate-200 rounded-xl shadow-xl max-h-56 overflow-y-auto divide-y divide-slate-100 animate-in fade-in-50 duration-150">
                      {filteredEquipment
                        .filter(e => 
                          !equipmentSearch ||
                          e.name.toLowerCase().includes(equipmentSearch.toLowerCase()) ||
                          e.id.toLowerCase().includes(equipmentSearch.toLowerCase()) ||
                          e.type.toLowerCase().includes(equipmentSearch.toLowerCase())
                        )
                        .map(eq => (
                          <div
                            key={eq.id}
                            onClick={() => {
                              setSelectedEquipmentId(eq.id);
                              const resolved = resolveEquipmentType(eq.type);
                              setCurrentEqType(resolved);
                              setWoData(prev => ({
                                ...prev,
                                equipmentId: eq.id,
                                equipmentName: eq.name,
                                factory: eq.factory || prev.factory || ''
                              }));
                              setEquipmentSearch('');
                              setShowEquipmentDropdown(false);
                            }}
                            className={`p-3 hover:bg-blue-50/60 cursor-pointer transition-colors flex items-center justify-between ${
                              selectedEquipmentId === eq.id ? 'bg-blue-50 font-bold text-blue-900' : 'text-slate-800'
                            }`}
                          >
                            <div>
                              <div className="text-xs font-bold text-slate-900 flex items-center gap-1.5">
                                <span>{eq.name}</span>
                                <span className="text-[10px] px-1.5 py-0.2 rounded bg-slate-100 text-slate-600 border border-slate-200">
                                  {eq.type}
                                </span>
                              </div>
                              <div className="text-[11px] text-slate-500 mt-0.5">
                                Tag: <span className="font-mono text-blue-700">{eq.id}</span>
                                {eq.location && ` • Vị trí: ${eq.location}`}
                              </div>
                            </div>
                            {selectedEquipmentId === eq.id && (
                              <Check size={16} className="text-blue-600 shrink-0" />
                            )}
                          </div>
                        ))}

                      <div
                        onClick={() => setEquipmentSelectMode('manual')}
                        className="p-3 bg-blue-50/40 hover:bg-blue-100/60 text-blue-800 text-xs font-bold cursor-pointer flex items-center gap-2 border-t border-blue-200"
                      >
                        <Plus size={15} />
                        <span>Không tìm thấy? Nhập thủ công thiết bị mới tại site...</span>
                      </div>
                    </div>
                  )}
                </div>

                {/* Selected Equipment Card Preview */}
                {currentEquipment && (
                  <div className="p-3.5 rounded-xl bg-blue-50/50 border border-blue-200 text-xs space-y-1.5">
                    <div className="flex items-center justify-between">
                      <span className="font-black text-slate-900">{currentEquipment.name}</span>
                      <span className="px-2 py-0.5 rounded bg-blue-200/60 text-blue-900 font-mono font-bold text-[10px]">
                        {currentEquipment.id}
                      </span>
                    </div>
                    <div className="flex items-center gap-3 text-slate-600">
                      <span className="font-semibold text-blue-700">{currentEquipment.type}</span>
                      <span>•</span>
                      <span className="flex items-center gap-1"><MapPin size={12} /> {currentEquipment.location || currentEquipment.factory || 'Trạm chính'}</span>
                    </div>
                    {currentEquipment.nameplate?.manufacturer && (
                      <div className="text-slate-500 text-[11px]">
                        Hãng: {currentEquipment.nameplate.manufacturer} {currentEquipment.nameplate.model ? `(${currentEquipment.nameplate.model})` : ''}
                      </div>
                    )}
                    {/* NETA Form Badge & Shortcut */}
                    <div className="mt-2 pt-2 border-t border-blue-200/60 flex items-center justify-between">
                      <div className="flex items-center gap-1.5 text-xs text-blue-900 font-bold">
                        <FileCheck2 size={13} className="text-emerald-600" />
                        <span>Form NETA:</span>
                        <span className="px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 font-mono text-[10px]">
                          {activeNetaSpec.sectionCode}
                        </span>
                      </div>
                      <button
                        type="button"
                        onClick={() => setShowNetaUniversalModal(true)}
                        className="text-[11px] font-bold text-indigo-700 hover:text-indigo-900 underline flex items-center gap-1"
                      >
                        <ExternalLink size={11} />
                        <span>Xem Biểu Mẫu NETA</span>
                      </button>
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* MODE B: MANUAL ENTRY FROM SITE */}
            {equipmentSelectMode === 'manual' && (
              <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-3 text-xs">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-slate-800">Thêm thiết bị mới tại công trường</span>
                  <button
                    type="button"
                    onClick={() => setEquipmentSelectMode('existing')}
                    className="text-slate-400 hover:text-slate-600"
                  >
                    <X size={15} />
                  </button>
                </div>

                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="block text-[11px] font-bold text-slate-700 mb-1">Mã TB (Tag) *</label>
                    <input
                      type="text"
                      placeholder="VD: VCB-24-05"
                      value={manualEquipment.id}
                      onChange={(e) => setManualEquipment({ ...manualEquipment, id: e.target.value })}
                      className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg font-mono font-semibold"
                    />
                  </div>
                  <div>
                    <label className="block text-[11px] font-bold text-slate-700 mb-1">Loại thiết bị</label>
                    <select
                      value={manualEquipment.type}
                      onChange={(e) => {
                        setManualEquipment({ ...manualEquipment, type: e.target.value });
                        setCurrentEqType(e.target.value);
                      }}
                      className="w-full px-2 py-1.5 bg-white border border-slate-300 rounded-lg font-bold text-blue-800"
                    >
                      <option value="Máy cắt">Máy cắt (ACB / VCB / SF6)</option>
                      <option value="Động cơ">Động cơ (Motor)</option>
                      <option value="Máy phát">Máy phát điện (Generator)</option>
                      <option value="Tủ điện">Tủ điện (Switchgear / RMU)</option>
                      <option value="Máy biến áp">Máy biến áp (Transformer)</option>
                      <option value="Inverter">Inverter Solar / Wind</option>
                      <option value="Cáp điện">Cáp điện lực (Cable)</option>
                      <option value="Tiếp địa">Hệ thống tiếp địa</option>
                    </select>
                  </div>
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Tên thiết bị *</label>
                  <input
                    type="text"
                    placeholder="VD: Máy cắt chân không lộ 471"
                    value={manualEquipment.name}
                    onChange={(e) => setManualEquipment({ ...manualEquipment, name: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg font-semibold"
                  />
                </div>

                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="block text-[11px] font-bold text-slate-700 mb-1">Hãng sản xuất</label>
                    <input
                      type="text"
                      placeholder="VD: Schneider / ABB"
                      value={manualEquipment.manufacturer}
                      onChange={(e) => setManualEquipment({ ...manualEquipment, manufacturer: e.target.value })}
                      className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg"
                    />
                  </div>
                  <div>
                    <label className="block text-[11px] font-bold text-slate-700 mb-1">Vị trí ngăn lộ</label>
                    <input
                      type="text"
                      placeholder="VD: Gian phân phối 24kV"
                      value={manualEquipment.location}
                      onChange={(e) => setManualEquipment({ ...manualEquipment, location: e.target.value })}
                      className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg"
                    />
                  </div>
                </div>

                <button
                  type="button"
                  onClick={handleSaveManualEquipment}
                  className="w-full py-2 bg-blue-600 hover:bg-blue-500 text-white font-bold rounded-lg transition-colors flex items-center justify-center gap-1.5 shadow-xs"
                >
                  <Check size={14} />
                  <span>Lưu Thiết Bị & Đồng Bộ Lên Sheet thietbi</span>
                </button>
              </div>
            )}
          </div>
        </div>

        {/* RIGHT COLUMN: STEP 3 (DYNAMIC SPECS) & STEP 4 (WORK ORDER 23 COLUMNS) */}
        <div className="lg:col-span-7 space-y-6">

          {/* STEP 3: DYNAMIC TECHNICAL SPECIFICATION FIELDS */}
          <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2.5">
                <span className="w-7 h-7 rounded-lg bg-indigo-600 text-white font-black text-xs flex items-center justify-center shadow-xs">
                  3
                </span>
                <div>
                  <h3 className="font-bold text-slate-900 text-sm flex items-center gap-2">
                    <span>Thông Số Kỹ Thuật Động:</span>
                    <span className="px-2.5 py-0.5 rounded-full bg-indigo-50 text-indigo-700 border border-indigo-200 text-xs font-extrabold">
                      {currentEqType}
                    </span>
                  </h3>
                  <p className="text-xs text-slate-500">Mỗi loại thiết bị có các trường thí nghiệm & kiểm định đặc thù riêng</p>
                </div>
              </div>

              {/* Selector to switch specs on the fly */}
              <div className="flex items-center gap-1 text-xs">
                <select
                  value={currentEqType}
                  onChange={(e) => setCurrentEqType(e.target.value)}
                  className="bg-slate-50 border border-slate-300 rounded-lg px-2.5 py-1 font-bold text-slate-700 outline-none"
                >
                  <option value="Cáp điện">Cáp điện (Power Cables - NETA 7.3)</option>
                  <option value="Máy cắt">Máy cắt (Circuit Breaker - NETA 7.6)</option>
                  <option value="Máy biến áp">Máy biến áp (Transformer - NETA 7.2)</option>
                  <option value="Động cơ">Động cơ (Motor - NETA 7.15)</option>
                  <option value="Máy phát">Máy phát (Generator - NETA 7.15/7.22)</option>
                  <option value="Tủ điện">Tủ điện (Switchgear - NETA 7.1)</option>
                  <option value="Tiếp địa">Hệ thống tiếp địa (Grounding - NETA 7.13)</option>
                  <option value="Rơ le bảo vệ">Rơ le bảo vệ (Protective Relays - NETA 7.9)</option>
                  <option value="Pin & UPS">Pin & UPS / Hệ thống DC (NETA 7.18)</option>
                </select>
              </div>
            </div>

            {/* DYNAMIC FORM PER EQUIPMENT TYPE */}
            {/* TYPE 1: MÁY CẮT (CIRCUIT BREAKER) */}
            {currentEqType === 'Máy cắt' && (
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 text-xs">
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Điện áp định mức (kV)</label>
                  <input
                    type="text"
                    value={techSpecs.ratedVoltageKv || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, ratedVoltageKv: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-semibold"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Dòng định mức (A)</label>
                  <input
                    type="text"
                    value={techSpecs.ratedCurrentA || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, ratedCurrentA: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-semibold"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Dòng cắt ngắn mạch Icu (kA)</label>
                  <input
                    type="text"
                    value={techSpecs.shortCircuitKa || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, shortCircuitKa: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-semibold"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Điện trở tiếp xúc Pha R (μΩ)</label>
                  <input
                    type="text"
                    value={techSpecs.contactResistanceR || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, contactResistanceR: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Điện trở tiếp xúc Pha Y (μΩ)</label>
                  <input
                    type="text"
                    value={techSpecs.contactResistanceY || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, contactResistanceY: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Điện trở tiếp xúc Pha B (μΩ)</label>
                  <input
                    type="text"
                    value={techSpecs.contactResistanceB || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, contactResistanceB: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Thời gian đóng t_close (ms)</label>
                  <input
                    type="text"
                    value={techSpecs.closeTimeMs || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, closeTimeMs: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Thời gian cắt t_open (ms)</label>
                  <input
                    type="text"
                    value={techSpecs.openTimeMs || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, openTimeMs: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Cách điện cực-vỏ IR (MΩ)</label>
                  <input
                    type="text"
                    value={techSpecs.insulationIrMo || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, insulationIrMo: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Mạch điều khiển (≥ 2.0 MΩ)</label>
                  <input
                    type="text"
                    value={techSpecs.controlCircuitIrMo || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, controlCircuitIrMo: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Áp suất khí SF6 / Chân không</label>
                  <input
                    type="text"
                    value={techSpecs.sf6GasPressureBar || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, sf6GasPressureBar: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Thử nghiệm Trip Unit L-S-I-G</label>
                  <input
                    type="text"
                    value={techSpecs.tripUnitLsig || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, tripUnitLsig: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg"
                  />
                </div>
              </div>
            )}

            {/* TYPE 2: ĐỘNG CƠ (MOTOR) */}
            {currentEqType === 'Động cơ' && (
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 text-xs">
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Công suất P (kW / HP)</label>
                  <input
                    type="text"
                    value={techSpecs.ratedPowerKw || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, ratedPowerKw: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-semibold"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Điện áp định mức (V)</label>
                  <input
                    type="text"
                    value={techSpecs.motorVoltageV || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, motorVoltageV: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Tốc độ quay (RPM)</label>
                  <input
                    type="text"
                    value={techSpecs.motorRpm || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, motorRpm: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Rung ổ bi DE (mm/s RMS)</label>
                  <input
                    type="text"
                    value={techSpecs.vibrationDeMmS || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, vibrationDeMmS: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Rung ổ bi NDE (mm/s RMS)</label>
                  <input
                    type="text"
                    value={techSpecs.vibrationNdeMmS || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, vibrationNdeMmS: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Nhiệt độ Stator (°C)</label>
                  <input
                    type="text"
                    value={techSpecs.statorTempC || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, statorTempC: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Cách điện Stator IR (MΩ)</label>
                  <input
                    type="text"
                    value={techSpecs.motorIrMo || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, motorIrMo: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Chỉ số phân cực PI (≥ 2.0)</label>
                  <input
                    type="text"
                    value={techSpecs.motorPi || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, motorPi: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-bold text-indigo-700"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Tổn hao Tan-Delta (%)</label>
                  <input
                    type="text"
                    value={techSpecs.motorTanDeltaPct || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, motorTanDeltaPct: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
              </div>
            )}

            {/* TYPE 3: MÁY PHÁT ĐIỆN (GENERATOR) */}
            {currentEqType === 'Máy phát' && (
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 text-xs">
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Công suất định mức (kVA/kW)</label>
                  <input
                    type="text"
                    value={techSpecs.genPowerKva || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, genPowerKva: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-semibold"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Hệ số công suất Cosφ</label>
                  <input
                    type="text"
                    value={techSpecs.genCosPhi || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, genCosPhi: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Tần số (Hz) & RPM</label>
                  <input
                    type="text"
                    value={`${techSpecs.genHz || '50'}Hz • ${techSpecs.genRpm || '1500'}rpm`}
                    onChange={(e) => setTechSpecs({ ...techSpecs, genHz: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Khởi động khẩn cấp (NFPA ≤ 10s)</label>
                  <input
                    type="text"
                    value={techSpecs.genStartSec || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, genStartSec: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-bold text-emerald-700"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Thử tải Load Bank (%)</label>
                  <input
                    type="text"
                    value={techSpecs.genLoadBankPct || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, genLoadBankPct: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Cách điện Stator & PI</label>
                  <input
                    type="text"
                    value={`${techSpecs.genIrStatorMo || '3500'}MΩ (PI: ${techSpecs.genPi || '3.1'})`}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-semibold"
                  />
                </div>
              </div>
            )}

            {/* TYPE 4: TỦ ĐIỆN (SWITCHGEAR) */}
            {currentEqType === 'Tủ điện' && (
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 text-xs">
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Cấp điện áp tủ (kV)</label>
                  <input
                    type="text"
                    value={techSpecs.swgVoltageKv || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, swgVoltageKv: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Dòng thanh cái Busbar (A)</label>
                  <input
                    type="text"
                    value={techSpecs.swgBusCurrentA || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, swgBusCurrentA: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Mức TEV PD (dBmV) [≤ 20]</label>
                  <input
                    type="text"
                    value={techSpecs.swgTevDbmv || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, swgTevDbmv: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-bold text-indigo-700"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Siêu âm Ultrasonic (dBμV)</label>
                  <input
                    type="text"
                    value={techSpecs.swgUltrasonicDbuv || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, swgUltrasonicDbuv: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Mật độ xung TEV (pps)</label>
                  <input
                    type="text"
                    value={techSpecs.swgPulsePps || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, swgPulsePps: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Khóa liên động cơ điện</label>
                  <input
                    type="text"
                    value={techSpecs.swgInterlockStatus || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, swgInterlockStatus: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg"
                  />
                </div>
              </div>
            )}

            {/* TYPE 5: MÁY BIẾN ÁP (TRANSFORMER) */}
            {currentEqType === 'Máy biến áp' && (
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 text-xs">
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Công suất định mức (kVA)</label>
                  <input
                    type="text"
                    value={techSpecs.trfPowerKva || ''}
                    onChange={(e) => setTechSpecs({ ...techSpecs, trfPowerKva: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-semibold"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Cấp điện áp Cao/Hạ (kV)</label>
                  <input
                    type="text"
                    value={`${techSpecs.trfPriVoltageKv || '22'} / ${techSpecs.trfSecVoltageV || '400'}V`}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Nhiệt độ dầu / Cuộn dây (°C)</label>
                  <input
                    type="text"
                    value={`${techSpecs.trfOilTempC || '58'}°C / ${techSpecs.trfWindingTempC || '68'}°C`}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Khí C2H2 Acetylene (ppm)</label>
                  <input
                    type="text"
                    value={techSpecs.trfDgaC2h2 || '0'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, trfDgaC2h2: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-bold text-rose-600"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Độ bền điện môi dầu (kV)</label>
                  <input
                    type="text"
                    value={techSpecs.trfBreakdownKv || '68'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, trfBreakdownKv: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-bold text-emerald-600"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Độ ẩm trong dầu (ppm)</label>
                  <input
                    type="text"
                    value={techSpecs.trfMoisturePpm || '12'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, trfMoisturePpm: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
              </div>
            )}

            {/* TYPE 6: CÁP ĐIỆN LỰC (POWER CABLES - ANSI/NETA ATS 7.3) */}
            {currentEqType === 'Cáp điện' && (
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 text-xs">
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Cấp điện áp & Cấu tạo cáp</label>
                  <input
                    type="text"
                    value={techSpecs.cableVoltageRating || '22kV 3x240mm2 Cu/XLPE/PVC'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, cableVoltageRating: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-semibold"
                    placeholder="VD: 22kV 3x240mm2 hoặc 0.6/1kV"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Tiết diện & Cấu tạo ruột dẫn</label>
                  <input
                    type="text"
                    value={techSpecs.cableSizeMm2 || '3x240mm2'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, cableSizeMm2: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                    placeholder="VD: 3x240 + 1x120 mm2"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Chiều dài tuyến cáp (m)</label>
                  <input
                    type="text"
                    value={techSpecs.cableLengthM || '350m'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, cableLengthM: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Cách điện Pha A - Đất/Shield (MΩ)</label>
                  <input
                    type="text"
                    value={techSpecs.cableIrPhaseEarthMo || '4500'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, cableIrPhaseEarthMo: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-bold text-emerald-700"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Cách điện Pha B - Đất/Shield (MΩ)</label>
                  <input
                    type="text"
                    value={techSpecs.cableIrPhaseBMo || '4200'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, cableIrPhaseBMo: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-bold text-emerald-700"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Cách điện Pha C - Đất/Shield (MΩ)</label>
                  <input
                    type="text"
                    value={techSpecs.cableIrPhaseCMo || '4600'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, cableIrPhaseCMo: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-bold text-emerald-700"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Cách điện Pha - Pha (MΩ)</label>
                  <input
                    type="text"
                    value={techSpecs.cableIrPhasePhaseMo || '5000'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, cableIrPhasePhaseMo: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Thử cao áp VLF Hipot (kV RMS / 30p)</label>
                  <input
                    type="text"
                    value={techSpecs.cableHipotVoltageKv || '28kV RMS (0.1Hz, 30 min)'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, cableHipotVoltageKv: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-semibold"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Tổn hao Tan-Delta (Mean TD x10⁻³)</label>
                  <input
                    type="text"
                    value={techSpecs.cableTanDelta || '1.2 (< 4.0)'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, cableTanDelta: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Phóng điện cục bộ PD cáp (pC)</label>
                  <input
                    type="text"
                    value={techSpecs.cablePdPc || '8 pC (≤ 10 pC)'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, cablePdPc: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Điện trở màn chắn Shield (Ω)</label>
                  <input
                    type="text"
                    value={techSpecs.cableShieldContinuity || '0.45 Ω (Pass)'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, cableShieldContinuity: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Thứ tự pha & Mối nối DLRO (μΩ)</label>
                  <input
                    type="text"
                    value={techSpecs.cableConductorResistanceMo || '12.5 μΩ • Đúng thứ tự pha'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, cableConductorResistanceMo: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
              </div>
            )}

            {/* TYPE 7: TIẾP ĐỊA (GROUNDING SYSTEMS - ANSI/NETA ATS 7.13) */}
            {currentEqType === 'Tiếp địa' && (
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 text-xs">
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Hệ thống tiếp địa</label>
                  <input
                    type="text"
                    value={techSpecs.gndSystemType || 'Lưới tiếp địa trạm TBA & Thiết bị'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, gndSystemType: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Phương pháp đo</label>
                  <input
                    type="text"
                    value={techSpecs.gndTestMethod || 'Fall-of-Potential 3 cực 62%'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, gndTestMethod: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Điện trở tiếp địa đo được (Ω)</label>
                  <input
                    type="text"
                    value={techSpecs.gndResistanceOhm || '0.65'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, gndResistanceOhm: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-bold text-emerald-700"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Yêu cầu tối đa cho phép (Ω)</label>
                  <input
                    type="text"
                    value={techSpecs.gndRequiredMaxOhm || '1.0'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, gndRequiredMaxOhm: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Liên kết vỏ thiết bị (mΩ)</label>
                  <input
                    type="text"
                    value={techSpecs.gndEquipotentialBondingMo || '15'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, gndEquipotentialBondingMo: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Mối hàn hóa nhiệt Cadweld</label>
                  <input
                    type="text"
                    value={techSpecs.gndVisualCondition || 'Đạt chuẩn, không ăn mòn'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, gndVisualCondition: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg"
                  />
                </div>
              </div>
            )}

            {/* TYPE 8: RƠ LE BẢO VỆ (PROTECTIVE RELAYS - ANSI/NETA ATS 7.9) */}
            {currentEqType === 'Rơ le bảo vệ' && (
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 text-xs">
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Hãng & Model Rơ le</label>
                  <input
                    type="text"
                    value={techSpecs.relayModel || 'SEL-751 / MiCOM P122'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, relayModel: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-semibold"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Mã chức năng bảo vệ ANSI</label>
                  <input
                    type="text"
                    value={techSpecs.relayFunctionCodes || '50/51, 50N/51N, 87, 27/59'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, relayFunctionCodes: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Dòng tác động đo được Pickup (A)</label>
                  <input
                    type="text"
                    value={techSpecs.relayPickupCurrentA || '5.02'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, relayPickupCurrentA: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-bold text-indigo-700"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Thời gian tác động Trip (ms)</label>
                  <input
                    type="text"
                    value={techSpecs.relayTripTimeMs || '120'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, relayTripTimeMs: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Sai số đo kiểm (%)</label>
                  <input
                    type="text"
                    value={techSpecs.relayErrorPct || '0.4% (≤ 2.0%)'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, relayErrorPct: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Thử nghiệm trip cắt máy cắt</label>
                  <input
                    type="text"
                    value={techSpecs.relayTripContactTest || 'Đóng cắt tốt / Trip OK'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, relayTripContactTest: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg"
                  />
                </div>
              </div>
            )}

            {/* TYPE 9: PIN & UPS / HỆ THỐNG DC (BATTERIES & UPS - ANSI/NETA ATS 7.18) */}
            {currentEqType === 'Pin & UPS' && (
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 text-xs">
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Loại ắc quy / Hệ thống DC</label>
                  <input
                    type="text"
                    value={techSpecs.batSystemType || '110V DC Giàn Ắc quy VRLA AGM'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, batSystemType: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-semibold"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Dung lượng định mức (Ah)</label>
                  <input
                    type="text"
                    value={techSpecs.batCapacityAh || '200Ah'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, batCapacityAh: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Tổng điện áp chuỗi (VDC)</label>
                  <input
                    type="text"
                    value={techSpecs.batTotalVoltageV || '124.5'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, batTotalVoltageV: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-bold text-indigo-700"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Nội trở cell lớn nhất (mΩ)</label>
                  <input
                    type="text"
                    value={techSpecs.batCellInternalResMo || '0.85'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, batCellInternalResMo: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Dải điện áp bình (V/cell)</label>
                  <input
                    type="text"
                    value={`${techSpecs.batMinCellVoltageV || '2.24'} - ${techSpecs.batMaxCellVoltageV || '2.26'}V`}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Thử tải xả Load Test (phút)</label>
                  <input
                    type="text"
                    value={techSpecs.batLoadTestDurationMin || '60 phút @ 0.1C'}
                    onChange={(e) => setTechSpecs({ ...techSpecs, batLoadTestDurationMin: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
              </div>
            )}
          </div>

          {/* DEDICATED NETA ATS TEST FORM & CHECKLIST (BIỂU MẪU KIỂM ĐỊNH & THỬ NGHIỆM TIÊU CHUẨN NETA ATS-2025) */}
          <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2.5">
                <span className="w-7 h-7 rounded-lg bg-amber-500 text-slate-950 font-black text-xs flex items-center justify-center shadow-xs">
                  <ShieldCheck size={16} />
                </span>
                <div>
                  <h3 className="font-bold text-slate-900 text-sm flex items-center gap-2 flex-wrap">
                    <span>Biểu Mẫu Thử Nghiệm Chuẩn ANSI/NETA ATS:</span>
                    <span className="px-2 py-0.5 rounded-full bg-amber-50 text-amber-900 border border-amber-200 font-mono text-xs font-black">
                      {activeNetaSpec.sectionCode} - {currentEqType}
                    </span>
                  </h3>
                  <p className="text-xs text-slate-500 mt-0.5">
                    {activeNetaSpec.standardCode} • {activeNetaSpec.subtitleVi || activeNetaSpec.equipmentTitleVi}
                  </p>
                </div>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => {
                    const allPassed = netaFormItems.map(item => ({ ...item, status: 'PASS' as const }));
                    setNetaFormItems(allPassed);
                    saveAtsPayload(activeNetaSpecId, {
                      ...getAtsPayload(activeNetaSpecId),
                      evaluationItems: allPassed
                    });
                  }}
                  className="px-3 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-300 rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all active:scale-95 shadow-2xs"
                  title="Đánh dấu tất cả hạng mục đạt chuẩn NETA ATS"
                >
                  <CheckCircle2 size={13} className="text-emerald-600" />
                  <span>Đạt Toàn Bộ (Pass All)</span>
                </button>

                <button
                  type="button"
                  onClick={() => setShowNetaUniversalModal(true)}
                  className="px-3.5 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all active:scale-95 shadow-xs"
                >
                  <Sparkles size={13} className="text-amber-300" />
                  <span>Mở Báo Cáo Nghiệm Thu NETA</span>
                </button>
              </div>
            </div>

            {/* Overall Compliance Status Banner */}
            <div className={`p-3.5 rounded-xl border flex items-start gap-3 ${
              activeNetaSpec.overallStatus === 'PASS'
                ? 'bg-emerald-50/70 border-emerald-200 text-emerald-950'
                : 'bg-amber-50/70 border-amber-200 text-amber-950'
            }`}>
              <div className={`p-1.5 rounded-lg shrink-0 mt-0.5 ${
                activeNetaSpec.overallStatus === 'PASS' ? 'bg-emerald-500 text-white' : 'bg-amber-500 text-white'
              }`}>
                {activeNetaSpec.overallStatus === 'PASS' ? <ShieldCheck size={18} /> : <AlertTriangle size={18} />}
              </div>
              <div className="space-y-1 text-xs">
                <div className="font-black text-sm flex items-center gap-2">
                  <span>{activeNetaSpec.statusTitleVi}</span>
                </div>
                <p className="text-slate-600 leading-relaxed">
                  {activeNetaSpec.statusReasonVi}
                </p>
              </div>
            </div>

            {/* Test Items Evaluation Table */}
            <div className="space-y-2">
              <div className="flex items-center justify-between text-xs font-bold text-slate-700">
                <span>Hạng mục kiểm tra & thí nghiệm điện ({netaFormItems.length} hạng mục):</span>
                <span className="text-[11px] text-slate-400">Nhấp để thay đổi trạng thái Đạt / Không đạt / Theo dõi</span>
              </div>

              <div className="border border-slate-200 rounded-xl overflow-hidden divide-y divide-slate-100 max-h-96 overflow-y-auto">
                {netaFormItems.map((item, idx) => (
                  <div key={idx} className="p-3 bg-white hover:bg-slate-50/80 transition-colors flex flex-col md:flex-row md:items-center justify-between gap-3 text-xs">
                    <div className="space-y-1 min-w-0 flex-1">
                      <div className="font-bold text-slate-900 flex items-center gap-2">
                        <span className="font-mono text-slate-400 text-[10px]">#{idx + 1}</span>
                        <span>{item.item}</span>
                      </div>
                      <div className="text-[11px] text-slate-500">
                        Tiêu chuẩn: <span className="font-semibold text-slate-700">{item.standard}</span>
                        {item.reference && <span className="text-slate-400 ml-1">({item.reference})</span>}
                      </div>
                      {item.detail && (
                        <div className="text-[11px] text-slate-600 italic">
                          • {item.detail}
                        </div>
                      )}
                    </div>

                    <div className="flex items-center gap-2 shrink-0">
                      <div className="text-right">
                        <div className="text-[10px] text-slate-400 font-semibold">Giá trị đo kiểm:</div>
                        <input
                          type="text"
                          value={item.measured || ''}
                          onChange={(e) => {
                            const updated = [...netaFormItems];
                            updated[idx] = { ...updated[idx], measured: e.target.value };
                            setNetaFormItems(updated);
                          }}
                          className="px-2 py-1 bg-slate-50 border border-slate-300 rounded font-mono font-bold text-slate-800 text-[11px] w-48 text-right focus:bg-white focus:border-indigo-500 outline-none"
                        />
                      </div>

                      <div className="flex items-center gap-1">
                        <button
                          type="button"
                          onClick={() => {
                            const updated = [...netaFormItems];
                            updated[idx] = { ...updated[idx], status: 'PASS' };
                            setNetaFormItems(updated);
                          }}
                          className={`px-2 py-1 rounded font-bold text-[10px] transition-all ${
                            item.status === 'PASS'
                              ? 'bg-emerald-600 text-white shadow-xs'
                              : 'bg-slate-100 text-slate-500 hover:bg-slate-200'
                          }`}
                        >
                          PASS
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            const updated = [...netaFormItems];
                            updated[idx] = { ...updated[idx], status: 'REVIEW' };
                            setNetaFormItems(updated);
                          }}
                          className={`px-2 py-1 rounded font-bold text-[10px] transition-all ${
                            item.status === 'REVIEW'
                              ? 'bg-amber-500 text-white shadow-xs'
                              : 'bg-slate-100 text-slate-500 hover:bg-slate-200'
                          }`}
                        >
                          CẢNH BÁO
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            const updated = [...netaFormItems];
                            updated[idx] = { ...updated[idx], status: 'FAIL' };
                            setNetaFormItems(updated);
                          }}
                          className={`px-2 py-1 rounded font-bold text-[10px] transition-all ${
                            item.status === 'FAIL'
                              ? 'bg-rose-600 text-white shadow-xs'
                              : 'bg-slate-100 text-slate-500 hover:bg-slate-200'
                          }`}
                        >
                          FAIL
                        </button>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            {/* Sync to Work Order Description button */}
            <div className="flex items-center justify-between pt-1">
              <span className="text-[11px] text-slate-500">
                Toàn bộ dữ liệu đo kiểm NETA được lưu tự động và tích hợp vào biên bản thí nghiệm.
              </span>
              <button
                type="button"
                onClick={() => {
                  const summaryText = `[ĐẠT CHUẨN NETA ATS ${activeNetaSpec.sectionCode}] ${activeNetaSpec.statusReasonVi}`;
                  setWoData(prev => ({
                    ...prev,
                    description: prev.description ? `${prev.description}\n${summaryText}` : summaryText
                  }));
                  alert('Đã cập nhật kết luận NETA ATS vào nội dung phiếu Work Order!');
                }}
                className="text-xs font-bold text-indigo-700 hover:text-indigo-900 underline flex items-center gap-1"
              >
                <ClipboardList size={13} />
                <span>Nạp kết luận NETA vào Work Order</span>
              </button>
            </div>
          </div>

          {/* STEP 4: LẬP PHIẾU CÔNG VIỆC WORK ORDER & WORK PERMIT (CHUẨN 23 TRƯỜNG TEV SERVICE FLATFORM) */}
          <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2.5">
                <span className="w-7 h-7 rounded-lg bg-emerald-600 text-white font-black text-xs flex items-center justify-center shadow-xs">
                  4
                </span>
                <div>
                  <h3 className="font-bold text-slate-900 text-sm">
                    Phiếu Công Việc Work Order & Work Permit (23 Cột Chuẩn TEV)
                  </h3>
                  <p className="text-xs text-slate-500">Đồng bộ tự động 2 chiều lên sheet <code className="text-emerald-700 font-bold">TEV Service Flatform</code></p>
                </div>
              </div>

              <span className="px-2.5 py-1 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200 text-xs font-mono font-bold">
                {woData.id}
              </span>
            </div>

            {/* Form Fields: All 23 Headers matched verbatim */}
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 text-xs">
              {/* 1. Mã WO */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">1. Mã WO *</label>
                <input
                  type="text"
                  value={woData.id || ''}
                  onChange={(e) => setWoData({ ...woData, id: e.target.value })}
                  className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-bold text-emerald-800"
                />
              </div>

              {/* 2. Mã Work Permit */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">2. Mã Work Permit *</label>
                <input
                  type="text"
                  value={woData.workPermitId || ''}
                  onChange={(e) => setWoData({ ...woData, workPermitId: e.target.value })}
                  className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-bold text-blue-800"
                />
              </div>

              {/* 6. Failure Code */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">6. Failure Code *</label>
                <select
                  value={woData.failureCode || 'PM-ROUTINE'}
                  onChange={(e) => setWoData({ ...woData, failureCode: e.target.value })}
                  className="w-full px-2 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-semibold text-rose-700"
                >
                  {FAILURE_CODES.map(fc => (
                    <option key={fc.code} value={fc.code}>{fc.code} - {fc.desc}</option>
                  ))}
                </select>
              </div>

              {/* 3. Tiêu đề */}
              <div className="sm:col-span-2">
                <label className="block text-[11px] font-bold text-slate-700 mb-1">3. Tiêu đề *</label>
                <input
                  type="text"
                  value={woData.title || ''}
                  onChange={(e) => setWoData({ ...woData, title: e.target.value })}
                  placeholder="Nhập tiêu đề công việc thực hiện..."
                  className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg font-semibold text-slate-900"
                />
              </div>

              {/* 8. Loại công việc */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">8. Loại công việc</label>
                <select
                  value={woData.type || 'Kiểm định thử nghiệm (Testing)'}
                  onChange={(e) => setWoData({ ...woData, type: e.target.value })}
                  className="w-full px-2 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-semibold"
                >
                  {WORK_TYPES.map(t => (
                    <option key={t} value={t}>{t}</option>
                  ))}
                </select>
              </div>

              {/* 4. Mô tả */}
              <div className="sm:col-span-3">
                <label className="block text-[11px] font-bold text-slate-700 mb-1">4. Mô tả công việc</label>
                <textarea
                  rows={2}
                  value={woData.description || ''}
                  onChange={(e) => setWoData({ ...woData, description: e.target.value })}
                  placeholder="Mô tả chi tiết nội dung bảo dưỡng, thí nghiệm hoặc khắc phục sự cố..."
                  className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg text-xs"
                />
              </div>

              {/* 5. Thiết bị (auto from step 2) */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">5. Thiết bị</label>
                <input
                  type="text"
                  disabled
                  value={`${woData.equipmentId || ''} ${woData.equipmentName ? `- ${woData.equipmentName}` : ''}`}
                  className="w-full px-2.5 py-1.5 bg-slate-100 border border-slate-300 rounded-lg font-semibold text-slate-700 cursor-not-allowed"
                />
              </div>

              {/* 7. Khách hàng (auto from step 1) */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">7. Khách hàng</label>
                <input
                  type="text"
                  disabled
                  value={woData.customer || ''}
                  className="w-full px-2.5 py-1.5 bg-slate-100 border border-slate-300 rounded-lg font-semibold text-slate-700 cursor-not-allowed"
                />
              </div>

              {/* 9. Ngoài kế hoạch */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">9. Ngoài kế hoạch?</label>
                <select
                  value={woData.isUnplanned ? 'Có' : 'Không'}
                  onChange={(e) => setWoData({ ...woData, isUnplanned: e.target.value === 'Có' })}
                  className="w-full px-2 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-bold"
                >
                  <option value="Không">Không (Kế hoạch định kỳ)</option>
                  <option value="Có">Có (Sự cố đột xuất Unplanned)</option>
                </select>
              </div>

              {/* 10. Mức độ ưu tiên */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">10. Mức độ ưu tiên</label>
                <select
                  value={woData.priority || 'medium'}
                  onChange={(e) => setWoData({ ...woData, priority: e.target.value })}
                  className="w-full px-2 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-bold"
                >
                  {PRIORITIES.map(p => (
                    <option key={p.value} value={p.value}>{p.label}</option>
                  ))}
                </select>
              </div>

              {/* 11. Trạng thái */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">11. Trạng thái</label>
                <select
                  value={woData.status || 'Mới tạo'}
                  onChange={(e) => setWoData({ ...woData, status: e.target.value })}
                  className="w-full px-2 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-bold text-blue-700"
                >
                  {WORK_STATUSES.map(s => (
                    <option key={s} value={s}>{s}</option>
                  ))}
                </select>
              </div>

              {/* 12. Người thực hiện (PIC) */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">12. Người thực hiện (PIC)</label>
                <input
                  type="text"
                  value={woData.assignedTo || ''}
                  onChange={(e) => setWoData({ ...woData, assignedTo: e.target.value })}
                  className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg font-medium"
                />
              </div>

              {/* 13. Vai trò Phê duyệt */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">13. Vai trò Phê duyệt</label>
                <input
                  type="text"
                  value={woData.responsibleApprove || ''}
                  onChange={(e) => setWoData({ ...woData, responsibleApprove: e.target.value })}
                  className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg"
                />
              </div>

              {/* 14. Vai trò Thực hiện */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">14. Vai trò Thực hiện</label>
                <input
                  type="text"
                  value={woData.responsibleDo || ''}
                  onChange={(e) => setWoData({ ...woData, responsibleDo: e.target.value })}
                  className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg"
                />
              </div>

              {/* 15. Yêu cầu Cô lập LOTO */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">15. Yêu cầu Cô lập nguồn?</label>
                <select
                  value={woData.blockingRequired ? 'Có' : 'Không'}
                  onChange={(e) => setWoData({ ...woData, blockingRequired: e.target.value === 'Có' })}
                  className="w-full px-2 py-1.5 bg-slate-50 border border-slate-300 rounded-lg font-bold text-amber-800"
                >
                  <option value="Có">Có (Lockout/Tagout Bắt buộc)</option>
                  <option value="Không">Không (Kiểm tra Online)</option>
                </select>
              </div>

              {/* 16. Hạn hoàn thành */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">16. Hạn hoàn thành</label>
                <input
                  type="date"
                  value={woData.dueDate ? woData.dueDate.split('T')[0] : ''}
                  onChange={(e) => setWoData({ ...woData, dueDate: e.target.value })}
                  className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg"
                />
              </div>

              {/* 17. Vật tư sử dụng */}
              <div className="sm:col-span-2">
                <label className="block text-[11px] font-bold text-slate-700 mb-1">17. Vật tư sử dụng</label>
                <input
                  type="text"
                  value={typeof woData.usedMaterials === 'string' ? woData.usedMaterials : JSON.stringify(woData.usedMaterials || '')}
                  onChange={(e) => setWoData({ ...woData, usedMaterials: e.target.value })}
                  placeholder="Ghi nhận các vật tư phụ tùng đã dùng..."
                  className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg"
                />
              </div>
            </div>

            {/* 1-TOUCH TIMESTAMP LOGGING (FIELDS 20, 21, 22, 23) */}
            <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                  <Clock size={15} className="text-indigo-600" />
                  Ghi nhận mốc thời gian dừng máy & sửa chữa (Cột 20 - 23)
                </span>
                <span className="text-[11px] text-slate-500">Nhấn nút để ghi nhận thời gian thực tế tức thì</span>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                {/* 20. Bắt đầu dừng máy */}
                <div className="bg-white p-2.5 rounded-lg border border-slate-200 space-y-1">
                  <div className="flex items-center justify-between">
                    <span className="text-[11px] font-bold text-rose-700">20. Bắt đầu dừng máy</span>
                    <button
                      type="button"
                      onClick={() => recordTimestamp('downtimeStart')}
                      className="px-2 py-0.5 rounded bg-rose-50 hover:bg-rose-100 text-rose-700 font-bold text-[10px]"
                    >
                      Bấm giờ
                    </button>
                  </div>
                  <input
                    type="text"
                    value={woData.downtimeStart || ''}
                    onChange={(e) => setWoData({ ...woData, downtimeStart: e.target.value })}
                    placeholder="Chưa ghi nhận"
                    className="w-full text-xs font-mono bg-slate-50 border border-slate-200 rounded px-2 py-1 text-slate-700"
                  />
                </div>

                {/* 21. Bắt đầu sửa chữa */}
                <div className="bg-white p-2.5 rounded-lg border border-slate-200 space-y-1">
                  <div className="flex items-center justify-between">
                    <span className="text-[11px] font-bold text-amber-700">21. Bắt đầu sửa chữa</span>
                    <button
                      type="button"
                      onClick={() => recordTimestamp('repairStart')}
                      className="px-2 py-0.5 rounded bg-amber-50 hover:bg-amber-100 text-amber-700 font-bold text-[10px]"
                    >
                      Bấm giờ
                    </button>
                  </div>
                  <input
                    type="text"
                    value={woData.repairStart || ''}
                    onChange={(e) => setWoData({ ...woData, repairStart: e.target.value })}
                    placeholder="Chưa ghi nhận"
                    className="w-full text-xs font-mono bg-slate-50 border border-slate-200 rounded px-2 py-1 text-slate-700"
                  />
                </div>

                {/* 22. Kết thúc sửa chữa */}
                <div className="bg-white p-2.5 rounded-lg border border-slate-200 space-y-1">
                  <div className="flex items-center justify-between">
                    <span className="text-[11px] font-bold text-blue-700">22. Kết thúc sửa</span>
                    <button
                      type="button"
                      onClick={() => recordTimestamp('repairEnd')}
                      className="px-2 py-0.5 rounded bg-blue-50 hover:bg-blue-100 text-blue-700 font-bold text-[10px]"
                    >
                      Bấm giờ
                    </button>
                  </div>
                  <input
                    type="text"
                    value={woData.repairEnd || ''}
                    onChange={(e) => setWoData({ ...woData, repairEnd: e.target.value })}
                    placeholder="Chưa ghi nhận"
                    className="w-full text-xs font-mono bg-slate-50 border border-slate-200 rounded px-2 py-1 text-slate-700"
                  />
                </div>

                {/* 23. Chạy lại máy */}
                <div className="bg-white p-2.5 rounded-lg border border-slate-200 space-y-1">
                  <div className="flex items-center justify-between">
                    <span className="text-[11px] font-bold text-emerald-700">23. Chạy lại máy</span>
                    <button
                      type="button"
                      onClick={() => recordTimestamp('restartTime')}
                      className="px-2 py-0.5 rounded bg-emerald-50 hover:bg-emerald-100 text-emerald-700 font-bold text-[10px]"
                    >
                      Bấm giờ
                    </button>
                  </div>
                  <input
                    type="text"
                    value={woData.restartTime || ''}
                    onChange={(e) => setWoData({ ...woData, restartTime: e.target.value })}
                    placeholder="Chưa ghi nhận"
                    className="w-full text-xs font-mono bg-slate-50 border border-slate-200 rounded px-2 py-1 text-slate-700"
                  />
                </div>
              </div>
            </div>

            {/* ACTION BUTTON: SAVE & 2-WAY SYNC TO SHEET TEV SERVICE FLATFORM */}
            <div className="pt-2 flex flex-col sm:flex-row items-center gap-3">
              <button
                type="button"
                onClick={handleSaveAndSyncWorkOrder}
                disabled={savingLoading || isSyncing}
                className="w-full sm:flex-1 py-3 px-6 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-extrabold rounded-xl shadow-lg transition-all active:scale-98 flex items-center justify-center gap-2 text-sm disabled:opacity-50"
              >
                {savingLoading ? (
                  <>
                    <RefreshCw size={18} className="animate-spin" />
                    <span>Đang Lưu & Đồng Bộ 2 Chiều Lên Google Sheet...</span>
                  </>
                ) : (
                  <>
                    <UploadCloud size={18} />
                    <span>Lưu & Đồng Bộ 2 Chiều Lên Sheet "TEV Service Flatform"</span>
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* RECENT WORK ORDERS ON SHEET "TEV SERVICE FLATFORM" TABLE */}
      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="p-4 bg-slate-900 text-white flex flex-col md:flex-row md:items-center justify-between gap-3">
          <div className="flex items-center gap-2.5">
            <ClipboardList size={18} className="text-emerald-400" />
            <div>
              <h3 className="font-bold text-sm flex items-center gap-2">
                <span>Nhật Ký Phiếu Công Việc Đã Đồng Bộ</span>
                <code className="text-emerald-300 text-xs px-2 py-0.5 rounded bg-emerald-950/80 border border-emerald-800">TEV Service Flatform</code>
              </h3>
              <p className="text-[11px] text-slate-400">Tự động hiển thị các phiếu mới lưu & đồng bộ 2 chiều (Mới nhất lên đầu)</p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <div className="relative">
              <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                type="text"
                value={tableSearchQuery}
                onChange={(e) => setTableSearchQuery(e.target.value)}
                placeholder="Tìm mã WO, thiết bị, KH..."
                className="pl-8 pr-3 py-1.5 bg-slate-800 border border-slate-700 rounded-lg text-xs text-white placeholder-slate-400 outline-none focus:border-emerald-500 w-44 sm:w-56"
              />
              {tableSearchQuery && (
                <button onClick={() => setTableSearchQuery('')} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-white">
                  <X size={12} />
                </button>
              )}
            </div>

            <select
              value={tableStatusFilter}
              onChange={(e) => setTableStatusFilter(e.target.value)}
              className="bg-slate-800 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-slate-200 font-medium outline-none"
            >
              <option value="all">Tất cả trạng thái</option>
              <option value="Mới tạo">Mới tạo</option>
              <option value="Đang thực hiện">Đang thực hiện</option>
              <option value="Chờ phê duyệt">Chờ phê duyệt</option>
              <option value="Hoàn thành">Hoàn thành</option>
            </select>

            <select
              value={tablePageSize}
              onChange={(e) => setTablePageSize(Number(e.target.value))}
              className="bg-slate-800 border border-slate-700 rounded-lg px-2 py-1.5 text-xs text-slate-200 outline-none"
            >
              <option value={10}>Hiện 10</option>
              <option value={25}>Hiện 25</option>
              <option value={50}>Hiện 50</option>
              <option value={-1}>Tất cả ({displayOrders.length})</option>
            </select>

            <button
              onClick={onFetchFromSheets}
              disabled={isSyncing}
              className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-lg border border-slate-700 transition-colors"
              title="Làm mới từ Google Sheet"
            >
              <RefreshCw size={14} className={isSyncing ? 'animate-spin text-emerald-400' : ''} />
            </button>
          </div>
        </div>

        <div className="overflow-x-auto max-h-[500px] overflow-y-auto">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-50 text-slate-600 font-bold border-b border-slate-200 sticky top-0 z-10 shadow-xs">
              <tr>
                <th className="px-3.5 py-2.5">Mã WO</th>
                <th className="px-3.5 py-2.5">Mã Work Permit</th>
                <th className="px-3.5 py-2.5">Khách hàng</th>
                <th className="px-3.5 py-2.5">Thiết bị</th>
                <th className="px-3.5 py-2.5">Tiêu đề</th>
                <th className="px-3.5 py-2.5">Failure Code</th>
                <th className="px-3.5 py-2.5">Trạng thái</th>
                <th className="px-3.5 py-2.5">Người thực hiện</th>
                <th className="px-3.5 py-2.5">Dừng máy</th>
                <th className="px-3.5 py-2.5">Chạy lại máy</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 font-medium">
              {displayOrders.slice(0, tablePageSize === -1 ? displayOrders.length : tablePageSize).map((wo, i) => {
                const isJustSaved = wo.id === highlightedWoId;
                return (
                  <tr
                    key={wo.id || i}
                    className={`transition-colors ${
                      isJustSaved 
                        ? 'bg-emerald-50/90 border-l-4 border-l-emerald-600 font-semibold' 
                        : 'hover:bg-slate-50/80'
                    }`}
                  >
                    <td className="px-3.5 py-2.5 font-mono font-bold text-emerald-700 whitespace-nowrap">
                      <div className="flex items-center gap-1.5">
                        <span>{wo.id}</span>
                        {isJustSaved && (
                          <span className="px-1.5 py-0.5 rounded bg-emerald-600 text-white text-[9px] font-bold uppercase animate-pulse">
                            Mới đồng bộ
                          </span>
                        )}
                      </div>
                    </td>
                    <td className="px-3.5 py-2.5 font-mono text-blue-700 whitespace-nowrap">{wo.workPermitId || '-'}</td>
                    <td className="px-3.5 py-2.5 font-bold text-slate-800">{wo.customer || '-'}</td>
                    <td className="px-3.5 py-2.5 font-semibold text-slate-700">
                      {Array.isArray(wo.equipmentId) ? wo.equipmentId.join(', ') : (wo.equipmentId || wo.equipmentName || '-')}
                    </td>
                    <td className="px-3.5 py-2.5 text-slate-800 max-w-xs truncate" title={wo.title}>{wo.title}</td>
                    <td className="px-3.5 py-2.5 font-mono text-rose-700 whitespace-nowrap">{wo.failureCode || '-'}</td>
                    <td className="px-3.5 py-2.5 whitespace-nowrap">
                      <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${
                        wo.status === 'Hoàn thành' ? 'bg-emerald-100 text-emerald-800' :
                        wo.status === 'Đang thực hiện' ? 'bg-blue-100 text-blue-800' :
                        wo.status === 'Chờ phê duyệt' ? 'bg-amber-100 text-amber-800' :
                        'bg-slate-100 text-slate-700'
                      }`}>
                        {wo.status}
                      </span>
                    </td>
                    <td className="px-3.5 py-2.5 text-slate-600 whitespace-nowrap">{wo.assignedTo}</td>
                    <td className="px-3.5 py-2.5 text-slate-500 font-mono text-[11px] whitespace-nowrap">{wo.downtimeStart || '-'}</td>
                    <td className="px-3.5 py-2.5 text-slate-500 font-mono text-[11px] whitespace-nowrap">{wo.restartTime || '-'}</td>
                  </tr>
                );
              })}
              {displayOrders.length === 0 && (
                <tr>
                  <td colSpan={10} className="px-4 py-10 text-center text-slate-400">
                    Chưa có phiếu công việc nào phù hợp với bộ lọc. Hãy điền thông tin bên trên và bấm "Lưu & Đồng bộ" để tạo bản ghi đầu tiên!
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
        <div className="p-3 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-xs text-slate-500">
          <span>
            Đang hiển thị <strong>{Math.min(tablePageSize === -1 ? displayOrders.length : tablePageSize, displayOrders.length)}</strong> trên tổng số <strong>{displayOrders.length}</strong> phiếu công việc
          </span>
          <span className="text-[11px] text-slate-400">
            Dữ liệu đồng bộ tự động 2 chiều giữa App ↔ Firestore ↔ Google Sheet
          </span>
        </div>
      </div>

      {/* MODAL: TẠO MỚI KHÁCH HÀNG TẠI HIỆN TRƯỜNG */}
      {showNewCustomerModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full overflow-hidden animate-in zoom-in-95 duration-150">
            <div className="p-4 bg-amber-500 text-slate-950 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Building size={18} />
                <h3 className="font-bold text-sm">Thêm Mới Khách Hàng Hiện Trường</h3>
              </div>
              <button onClick={() => setShowNewCustomerModal(false)} className="text-slate-950/70 hover:text-slate-950">
                <X size={18} />
              </button>
            </div>

            <div className="p-5 space-y-3 text-xs">
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">Mã khách hàng</label>
                <input
                  type="text"
                  value={newCustomer.id}
                  onChange={(e) => setNewCustomer({ ...newCustomer, id: e.target.value })}
                  placeholder="VD: KH-008"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg font-mono font-bold text-amber-900"
                />
              </div>
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">Tên khách hàng / Đơn vị *</label>
                <input
                  type="text"
                  value={newCustomer.name}
                  onChange={(e) => setNewCustomer({ ...newCustomer, name: e.target.value })}
                  placeholder="VD: Công ty Nhiệt Điện Phả Lại"
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg font-semibold"
                />
              </div>
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">Nhà máy / Trạm / Site</label>
                <input
                  type="text"
                  value={newCustomer.factory}
                  onChange={(e) => setNewCustomer({ ...newCustomer, factory: e.target.value })}
                  placeholder="VD: Trạm 110kV Phả Lại 2"
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg"
                />
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Số điện thoại</label>
                  <input
                    type="text"
                    value={newCustomer.phone}
                    onChange={(e) => setNewCustomer({ ...newCustomer, phone: e.target.value })}
                    placeholder="VD: 0988xxx..."
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Email</label>
                  <input
                    type="email"
                    value={newCustomer.email}
                    onChange={(e) => setNewCustomer({ ...newCustomer, email: e.target.value })}
                    placeholder="contact@customer.com"
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg"
                  />
                </div>
              </div>
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">Địa chỉ</label>
                <input
                  type="text"
                  value={newCustomer.address}
                  onChange={(e) => setNewCustomer({ ...newCustomer, address: e.target.value })}
                  placeholder="VD: KCN Quế Võ, Bắc Ninh"
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg"
                />
              </div>
            </div>

            <div className="p-4 bg-slate-50 border-t border-slate-100 flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setShowNewCustomerModal(false)}
                className="px-4 py-2 rounded-xl text-slate-600 font-bold text-xs hover:bg-slate-200 transition-colors"
              >
                Hủy
              </button>
              <button
                type="button"
                onClick={handleSaveNewCustomer}
                className="px-5 py-2 rounded-xl bg-amber-500 hover:bg-amber-400 text-slate-950 font-bold text-xs shadow-md transition-all active:scale-95"
              >
                Lưu & Thêm Vào Dropdown
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MASTER DIRECTORY & PROJECT PREPARATION SHEET MODAL */}
      <MasterDirectorySheetManagerModal
        isOpen={showMasterDirectoryModal}
        onClose={() => {
          setShowMasterDirectoryModal(false);
          setProjectsList(getRegisteredProjects());
        }}
        isGoogleConnected={isGoogleConnected}
        googleSheetUrl={localStorage.getItem('tev_spreadsheet_id') ? `https://docs.google.com/spreadsheets/d/${localStorage.getItem('tev_spreadsheet_id')}` : undefined}
        onSyncAll={onSyncAllToSheets}
        initialTab="duan"
        onNavigateToNetaForm={handleNavigateToNetaFromProject}
        onSelectProjectForWorkOrder={(prj) => {
          handleSelectProject(prj);
          setShowMasterDirectoryModal(false);
        }}
      />

      {/* NETA ATS UNIVERSAL REPORT MODAL */}
      {showNetaUniversalModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/80 backdrop-blur-sm flex items-center justify-center p-2 sm:p-4 animate-in fade-in duration-200">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-6xl max-h-[94vh] flex flex-col overflow-hidden border border-slate-300">
            <div className="p-4 bg-slate-900 text-white flex items-center justify-between border-b border-slate-800">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-amber-500 rounded-lg text-slate-950 font-black">
                  <ShieldCheck size={20} />
                </div>
                <div>
                  <h3 className="font-bold text-base text-white flex items-center gap-2 flex-wrap">
                    <span>Biên Bản Thí Nghiệm Toàn Diện ANSI/NETA ATS-2025</span>
                    <span className="px-2 py-0.5 rounded text-xs font-mono font-semibold bg-amber-400/20 text-amber-300 border border-amber-400/30">
                      {activeNetaSpec.sectionCode} - {currentEqType}
                    </span>
                  </h3>
                  <p className="text-xs text-slate-400 mt-0.5">
                    {activeNetaSpec.standardCode} • Đánh giá tự động AI • In PDF & Xuất báo cáo kỹ thuật TEV
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setShowNetaUniversalModal(false)}
                className="w-9 h-9 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white flex items-center justify-center transition-colors"
                title="Đóng"
              >
                <X size={18} />
              </button>
            </div>

            <div className="flex-1 overflow-y-auto p-4 bg-slate-50">
              <NetaAtsUniversalReportView
                initialSpecId={activeNetaSpecId}
                initialLang="vi"
                onSaveReport={() => {
                  const summaryText = `[NETA ATS ${activeNetaSpec.sectionCode}] ${activeNetaSpec.statusReasonVi || 'Đạt tiêu chuẩn thí nghiệm nghiệm thu NETA ATS'}`;
                  setWoData(prev => ({
                    ...prev,
                    description: prev.description ? `${prev.description}\n${summaryText}` : summaryText
                  }));
                  setShowNetaUniversalModal(false);
                }}
              />
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
