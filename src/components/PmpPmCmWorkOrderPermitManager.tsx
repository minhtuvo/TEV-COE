import React, { useState, useMemo } from 'react';
import {
  Calendar,
  ClipboardList,
  CheckCircle,
  AlertCircle,
  Clock,
  Search,
  Plus,
  Printer,
  Download,
  RefreshCw,
  Zap,
  Shield,
  ShieldAlert,
  ArrowRight,
  TrendingUp,
  BarChart3,
  FileSpreadsheet,
  X,
  FileText,
  Save,
  Trash2,
  Lock,
  Building,
  Mail,
  SlidersHorizontal,
  ChevronRight,
  UserCheck,
  Check,
  AlertTriangle,
  Flame,
  PowerOff,
  Wrench,
  Layers,
  ArrowUpRight,
  Users,
  Maximize2,
  Minimize2,
  CalendarDays,
  RotateCcw
} from 'lucide-react';
import { CustomerSelectDropdown } from './CustomerSelectDropdown';
import { EquipmentSelectDropdown } from './EquipmentSelectDropdown';
import { FseEngineerDropdown } from './FseEngineerDropdown';
import { PmScheduleDashboard } from './PmScheduleDashboard';
import { getNextWorkOrderId, getNextWorkPermitId, getNextWorkRequestId } from '../services/netaCmmsBridgeService';

// ============================================================================
// 1. DATA TYPES & CMMS SCHEMA
// ============================================================================

export type PriorityCode = 'P1' | 'P3' | 'P6' | 'P7' | 'P9';

export type WorkStream = 'PM' | 'CM';

export type PmCycleDays = 7 | 14 | 30 | 60 | 90 | 180 | 360 | 720;

export type WorkflowStage =
  | '1_wr_pending'         // WR mới tạo từ site, chờ họp Daily Meeting
  | '2_daily_approved'     // Đã họp duyệt -> Thành WO, giao Engineer planning
  | '3_planning'           // Engineer đang planning LOTO, de-energy, spare parts
  | '4_wp_issued'          // Đã cấp Work Permit, đang thi công tại hiện trường
  | '5_wp_closed'          // Đã nghiệm thu hiện trường, rút LOTO, đóng WP
  | '6_wo_closed';         // Đã đóng WO trên CMMS, hoàn tất toàn bộ chu trình

export interface WorkOrderItem {
  // Identification
  id: string;                    // Mã WO (WO-YYYY-XXX)
  wrId?: string;                 // Mã Work Request ban đầu (WR-YYYY-XXX)
  workPermitId: string;          // Mã Work Permit (WP-YYYY-XXX)
  title: string;                 // Tiêu đề công việc
  description: string;           // Mô tả chi tiết khiếm khuyết / nội dung bảo trì
  equipmentId: string | string[];// Mã thiết bị KKS
  equipmentName?: string;        // Tên thiết bị
  failureCode: string;           // Mã lỗi sự cố
  customer: string;              // Khách hàng / Nhà máy
  customerId?: string;           // Mã khách hàng
  factory?: string;              // Phân xưởng / Trạm

  // Two Core Work Streams: PM vs CM
  workStream?: WorkStream;       // 'PM' (định kỳ) hoặc 'CM' (sự cố từ WR)
  pmCycleDays?: PmCycleDays;     // 7, 14, 30, 60, 90, 180, 360 days
  isUnplanned: boolean;          // true: CM (ngoài KH) | false: PM (định kỳ)

  // Standard Industrial Plant Priorities
  priority: string;              // Chuẩn text (low, medium, high, urgent)
  priorityCode?: PriorityCode;   // P1, P3, P6, P7, P9

  // Workflow Stages
  workflowStage?: WorkflowStage; // 1_wr_pending -> 2_daily_approved -> 3_planning -> 4_wp_issued -> 5_wp_closed -> 6_wo_closed
  status: string;                // initiated, approved, planned, scheduled, in-progress, completed, overdue, cancelled

  // Daily Review Meeting Approval (MM / OM / DM / PM)
  reportedBy?: string;           // Operator / Field Service Engineer phát hiện tại site
  reportedAt?: string;           // Ngày giờ phát hiện lỗi tại site
  dailyMeetingApprovedAt?: string;
  dailyMeetingApprovedByRole?: 'MM' | 'OM' | 'DM' | 'PM';
  dailyMeetingApprovedByName?: string;
  dailyMeetingNotes?: string;

  // Engineer Planning Phase
  assignedTo: string;            // Lead Engineer phụ trách (PIC)
  responsibleApprove: string;    // TM, MM, OM, PM
  responsibleDo: string;         // Everybody, ME, PE, Subcontractor
  blockingRequired: boolean;     // Yêu cầu cô lập nguồn LOTO
  deEnergyRequired?: boolean;    // Cắt điện / Cô lập năng lượng
  lotoTagNumbers?: string;       // Mã thẻ khóa LOTO
  isolationPoints?: string;      // Điểm cô lập: CB, Dao cách ly, Van
  usedMaterials: any[] | string; // Phụ tùng / Vật tư dự phòng đã dùng & cần dùng
  dueDate: string;               // Hạn hoàn thành SLA

  // Work Permit Execution & Closure
  wpIssuedAt?: string;
  wpIssuedBy?: string;
  wpClosedAt?: string;
  wpClosedBy?: string;
  wpCloseNotes?: string;

  // CMMS WO Closure
  woClosedAt?: string;
  woClosedBy?: string;
  rootCause?: string;
  actualTimeSpent?: number;
  estimatedTime?: number;
  laborCount?: number;
  laborCost?: number;
  partCost?: number;
  attachments?: string[];

  // 1-Touch Timestamps (Cols 20-23)
  createdAt: string;             // Cột 18: Ngày tạo
  updatedAt: string;             // Cột 19: Cập nhật cuối
  downtimeStart: string;         // Cột 20: Bắt đầu dừng máy
  repairStart: string;           // Cột 21: Bắt đầu sửa
  repairEnd: string;             // Cột 22: Kết thúc sửa
  restartTime: string;           // Cột 23: Chạy lại máy

  // Legacy compat
  type: string;
  pmFrequency?: string;
}

export interface CustomerItem {
  id: string;
  name: string;
  factories?: string[];
  [key: string]: any;
}

export interface EquipmentItem {
  id: string;
  name: string;
  type?: string;
  customer?: string;
  factory?: string;
  location?: string;
  health?: number;
  status?: string;
  lastCheck?: string;
  pmCycleMonths?: number;
  [key: string]: any;
}

interface Props {
  workOrders: WorkOrderItem[];
  allEquipment: EquipmentItem[];
  customers: CustomerItem[];
  isGoogleConnected: boolean;
  onConnectGoogle?: () => void;
  onSaveWorkOrder: (wo: WorkOrderItem, syncToSheetsImmediately?: boolean) => Promise<void>;
  onDeleteWorkOrder?: (id: string) => Promise<void> | void;
  onSyncAllToSheets?: () => Promise<void>;
  onFetchFromSheets?: () => Promise<void>;
  isSyncing?: boolean;
  onNavigateToNetaForm?: (equipment: any, standard: 'ats' | 'mts') => void;
  onSendPMAlerts?: (forceAll?: boolean) => Promise<void> | void;
  onOpenEmailConfig?: () => void;
  pmUpcomingTasks?: any[];
  userEmail?: string;
}

// ============================================================================
// 2. CONFIGURATIONS & STANDARDS
// ============================================================================

export const PLANT_PRIORITIES: Record<PriorityCode, {
  label: string;
  title: string;
  desc: string;
  badgeClass: string;
  bgLight: string;
  textClass: string;
  slaText: string;
  icon: any;
}> = {
  P1: {
    label: 'Priority 1',
    title: 'P1: Làm Ngay (Emergency / Dừng Máy)',
    desc: 'Sự cố nguy hiểm, cháy nổ, dừng dây chuyền, bắt buộc xử lý khẩn cấp trong 2-4h.',
    badgeClass: 'bg-rose-600 text-white font-black',
    bgLight: 'bg-rose-50 border-rose-200',
    textClass: 'text-rose-700',
    slaText: 'Khẩn cấp (ngay lập tức / trong 2-4h)',
    icon: Flame
  },
  P3: {
    label: 'Priority 3',
    title: 'P3: Cấp Bách (Thời hạn trong 1 tuần)',
    desc: 'Hư hỏng nhưng chưa dừng máy ngay, phải xử lý dứt điểm trong vòng 7 ngày.',
    badgeClass: 'bg-amber-500 text-white font-bold',
    bgLight: 'bg-amber-50 border-amber-200',
    textClass: 'text-amber-800',
    slaText: 'Thời hạn hoàn thành ≤ 7 ngày (1 tuần)',
    icon: Clock
  },
  P6: {
    label: 'Priority 6',
    title: 'P6: Bảo Trì Định Kỳ PM (Routine PM)',
    desc: 'Bảo dưỡng định kỳ theo chu kỳ 07d, 14d, 30d, 60d, 90d, 180d, 360d hoặc khiếm khuyết nhỏ.',
    badgeClass: 'bg-blue-600 text-white font-bold',
    bgLight: 'bg-blue-50 border-blue-200',
    textClass: 'text-blue-700',
    slaText: 'Theo chu kỳ lịch PM định kỳ',
    icon: Calendar
  },
  P7: {
    label: 'Priority 7',
    title: 'P7: Shutdown (Chỉ làm khi dừng máy toàn trạm)',
    desc: 'Công việc phức tạp bắt buộc phải chờ toàn trạm hoặc nhà máy dừng máy (Turnaround / Outage).',
    badgeClass: 'bg-purple-600 text-white font-bold',
    bgLight: 'bg-purple-50 border-purple-200',
    textClass: 'text-purple-700',
    slaText: 'Đợt Shutdown / Cắt điện toàn trạm kế tiếp',
    icon: PowerOff
  },
  P9: {
    label: 'Priority 9',
    title: 'P9: DMR (Design Modification Request)',
    desc: 'Yêu cầu thay đổi, cải tiến thiết kế kỹ thuật, bổ sung thiết bị hoặc thay đổi công nghệ.',
    badgeClass: 'bg-indigo-700 text-white font-bold',
    bgLight: 'bg-indigo-50 border-indigo-200',
    textClass: 'text-indigo-800',
    slaText: 'Theo phê duyệt dự án MOC / Cải tiến',
    icon: Wrench
  }
};

export const PM_CYCLE_OPTIONS: Array<{ days: PmCycleDays; label: string; desc: string }> = [
  { days: 7, label: '07 Ngày (Hàng tuần)', desc: 'Kiểm tra mức dầu, rò rỉ, vệ sinh bộ thở hạt hút ẩm' },
  { days: 14, label: '14 Ngày (Nửa tháng)', desc: 'Kiểm tra phóng điện TEV, nhiệt độ đầu cốt, ắc quy trạm' },
  { days: 30, label: '30 Ngày (Hàng tháng)', desc: 'Bảo trì định kỳ tủ điện, camera nhiệt tiếp xúc, hệ thống làm mát' },
  { days: 60, label: '60 Ngày (2 tháng)', desc: 'Kiểm tra cơ cấu đóng cắt máy cắt, đo dòng rò chống sét van' },
  { days: 90, label: '90 Ngày (3 tháng / Quý)', desc: 'Thử nghiệm rơ-le bảo vệ, phân tích rung động động cơ, đo điện trở tiếp đất' },
  { days: 180, label: '180 Ngày (6 tháng / Nửa năm)', desc: 'Lấy mẫu dầu MBA phân tích DGA sắc ký khí, thử nghiệm cao áp VLF' },
  { days: 360, label: '360 Ngày (1 năm / Hằng năm)', desc: 'Đại tu tổng thể, thử nghiệm NETA MTS toàn diện, thử tải máy phát' },
  { days: 720, label: '720 Ngày (2 năm / Trung tu)', desc: 'Trung tu, thí nghiệm nghiệm thu định kỳ chuyên sâu' }
];

export const WORKFLOW_STAGES: Record<WorkflowStage, {
  step: number;
  label: string;
  badge: string;
  shortTitle: string;
  desc: string;
}> = {
  '1_wr_pending': {
    step: 1,
    label: '1. Yêu Cầu WR (Chờ Daily Meeting)',
    shortTitle: '1. Chờ Duyệt WR',
    badge: 'bg-amber-100 text-amber-900 border-amber-300',
    desc: 'Field Service / Operator đi site phát hiện bất thường và mở WR.'
  },
  '2_daily_approved': {
    step: 2,
    label: '2. Đã Duyệt WR ➔ WO (Giao Engineer)',
    shortTitle: '2. Đã Thành WO',
    badge: 'bg-blue-100 text-blue-900 border-blue-300',
    desc: 'Họp Daily Meeting duyệt (MM/OM/DM/PM), cấp mã WO và giao Engineer.'
  },
  '3_planning': {
    step: 3,
    label: '3. Engineer Planning (LOTO / Vật Tư)',
    shortTitle: '3. Đang Planning',
    badge: 'bg-indigo-100 text-indigo-900 border-indigo-300',
    desc: 'Kỹ sư chuẩn bị: De-energy, thẻ khóa LOTO, phụ tùng, tạo Work Permit.'
  },
  '4_wp_issued': {
    step: 4,
    label: '4. Đã Cấp WP (Đang Thi Công Hiện Trường)',
    shortTitle: '4. WP Đang Mở',
    badge: 'bg-purple-100 text-purple-900 border-purple-300',
    desc: 'Cán bộ an toàn cấp phép WP, đội công tác thực hiện tại trạm.'
  },
  '5_wp_closed': {
    step: 5,
    label: '5. Đã Đóng WP (Nghiệm Thu An Toàn / Rút LOTO)',
    shortTitle: '5. Đã Đóng WP',
    badge: 'bg-teal-100 text-teal-900 border-teal-300',
    desc: 'Hoàn tất hiện trường, rút LOTO, hoàn trả mặt bằng, đóng giấy phép.'
  },
  '6_wo_closed': {
    step: 6,
    label: '6. Đã Đóng WO Trên CMMS (Hoàn Tất)',
    shortTitle: '6. Đã Đóng WO',
    badge: 'bg-emerald-100 text-emerald-900 border-emerald-300',
    desc: 'Nghiệm thu kỹ thuật, cập nhật lịch sử thiết bị và đóng WO trên hệ thống.'
  }
};

export const APPROVER_ROLES = [
  { role: 'MM', title: 'Maintenance Manager (Trưởng phòng Bảo trì)' },
  { role: 'OM', title: 'Operation Manager (Trưởng phòng Vận hành)' },
  { role: 'DM', title: 'Deputy Plant Manager (Phó Giám đốc Nhà máy)' },
  { role: 'PM', title: 'Plant Manager (Giám đốc Nhà máy)' }
];

export const FAILURE_CODES = [
  { code: 'ROUTINE-PM', desc: 'Bảo trì định kỳ chuẩn theo chu kỳ ngày (Routine PM)' },
  { code: 'ELEC-01', desc: 'Suy giảm điện trở cách điện cuộn dây (Insulation Degradation)' },
  { code: 'ELEC-02', desc: 'Phóng điện cục bộ bên trong tủ/máy (Partial Discharge PD)' },
  { code: 'ELEC-03', desc: 'Lệch tỷ số biến áp / Lệch điện trở DC (Turns/DC Imbalance)' },
  { code: 'THERM-01', desc: 'Quá nhiệt tiếp xúc / Busbar / Đầu cáp (Thermal Hotspot)' },
  { code: 'DGA-01', desc: 'Sinh khí cháy MBA vượt ngưỡng IEEE C57.104 (Combustible Gas)' },
  { code: 'MECH-01', desc: 'Rung động vòng bi / Kẹt cơ cấu cơ khí (Vibration / Mechanical)' },
  { code: 'VLF-HIPOT', desc: 'Thử nghiệm cao áp VLF cáp ngầm 22kV (VLF Cable Hipot)' },
  { code: 'BATT-IMP', desc: 'Nội trở ắc quy tăng / Điện áp sụt giảm (Battery Impedance)' },
  { code: 'RELAY-TEST', desc: 'Sai lệch ngưỡng tác động bảo vệ rơ-le (Relay Trip Deviation)' },
  { code: 'ATS-INSP', desc: 'Lỗi cơ cấu chuyển nguồn tự động ATS (Transfer Switch Failure)' },
  { code: 'SOLAR-IV', desc: 'Suy giảm đặc tuyến I-V chuỗi PV Inverter (Solar IV Degradation)' },
  { code: 'WIND-VIB', desc: 'Rung động hộp số Tua-bin Gió (Wind Gearbox Vibration)' }
];

// Helper to deduce workflow stage from existing workOrder fields
function deduceWorkflowStage(wo: WorkOrderItem): WorkflowStage {
  if (wo.workflowStage) return wo.workflowStage;
  if (wo.status === 'completed' || wo.woClosedAt) return '6_wo_closed';
  if (wo.wpClosedAt) return '5_wp_closed';
  if (wo.workPermitId && (wo.status === 'in-progress' || wo.status === 'scheduled')) return '4_wp_issued';
  if (wo.status === 'planned') return '3_planning';
  if (wo.status === 'approved') return '2_daily_approved';
  if (wo.wrId && !wo.dailyMeetingApprovedAt && wo.status === 'initiated') return '1_wr_pending';
  return '2_daily_approved';
}

// Helper to deduce PriorityCode
function deducePriorityCode(wo: WorkOrderItem): PriorityCode {
  if (wo.priorityCode) return wo.priorityCode;
  if (!wo.isUnplanned && (wo.type === 'preventive' || wo.workStream === 'PM')) return 'P6';
  const p = (wo.priority || '').toLowerCase();
  if (p === 'urgent' || p === 'critical' || p === 'p1') return 'P1';
  if (p === 'high' || p === 'p3') return 'P3';
  if (p === 'p7' || wo.title?.toLowerCase().includes('shutdown')) return 'P7';
  if (p === 'p9' || wo.title?.toLowerCase().includes('dmr')) return 'P9';
  return 'P3';
}

// ============================================================================
// HELPERS: DATETIME-LOCAL CONVERSION FOR MILESTONES (COLUMNS 20-23)
// ============================================================================
function parseToDatetimeLocal(val?: string): string {
  if (!val || typeof val !== 'string') return '';
  const s = val.trim();
  if (!s || s.toLowerCase() === 'chưa ghi nhận') return '';
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(s)) return s.slice(0, 16);
  if (/^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}/.test(s)) return s.replace(/\s+/, 'T').slice(0, 16);
  const dmyMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s*(\d{2}:\d{2})?/);
  if (dmyMatch) {
    const [, d, m, y, t] = dmyMatch;
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}T${t || '08:00'}`;
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    const hours = String(d.getHours()).padStart(2, '0');
    const minutes = String(d.getMinutes()).padStart(2, '0');
    return `${year}-${month}-${day}T${hours}:${minutes}`;
  }
  return '';
}

function formatFromDatetimeLocal(val: string): string {
  if (!val) return '';
  return val.replace('T', ' ').slice(0, 16);
}

function getNowFormattedDatetime(): string {
  const d = new Date();
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  const hours = String(d.getHours()).padStart(2, '0');
  const minutes = String(d.getMinutes()).padStart(2, '0');
  return `${year}-${month}-${day} ${hours}:${minutes}`;
}

function calculateMilestoneDiffHours(t0Str?: string, t1Str?: string): number | null {
  if (!t0Str || !t1Str) return null;
  const d0 = new Date(t0Str.replace(' ', 'T')).getTime();
  const d1 = new Date(t1Str.replace(' ', 'T')).getTime();
  if (isNaN(d0) || isNaN(d1) || d1 < d0) return null;
  return Number(((d1 - d0) / (1000 * 60 * 60)).toFixed(1));
}

// ============================================================================
// 3. MAIN COMPONENT: PmpPmCmWorkOrderPermitManager
// ============================================================================

export const PmpPmCmWorkOrderPermitManager: React.FC<Props> = ({
  workOrders,
  allEquipment,
  customers,
  isGoogleConnected,
  onConnectGoogle,
  onSaveWorkOrder,
  onDeleteWorkOrder,
  onSyncAllToSheets,
  onFetchFromSheets,
  isSyncing = false,
  onNavigateToNetaForm,
  onSendPMAlerts,
  onOpenEmailConfig,
  pmUpcomingTasks = [],
  userEmail = 'sgm1707@gmail.com'
}) => {
  // Navigation Tabs: Default to 'pm-dashboard' as requested
  const [activeSubTab, setActiveSubTab] = useState<
    'all-work' | 'daily-meeting' | 'pm-matrix' | 'wp-loto' | 'pm-dashboard' | 'alerts' | 'sheets'
  >('pm-dashboard');

  // View modes for Work Orders Table: 'full-23-cols' | 'workflow-kanban' | 'compact'
  const [viewMode, setViewMode] = useState<'full-23-cols' | 'workflow-kanban' | 'compact'>('full-23-cols');

  // Filter state
  const [searchQuery, setSearchQuery] = useState('');
  const [filterStream, setFilterStream] = useState<'all' | 'PM' | 'CM'>('all');
  const [filterPriority, setFilterPriority] = useState<string>('all');
  const [filterStage, setFilterStage] = useState<string>('all');
  const [filterLoto, setFilterLoto] = useState<string>('all');
  const [filterAssignee, setFilterAssignee] = useState<string>('all');

  // Modals
  const [isWrModalOpen, setIsWrModalOpen] = useState(false);               // Modal 1: Mở WR mới (Operator/Site)
  const [isPmModalOpen, setIsPmModalOpen] = useState(false);               // Modal 2: Tạo PM định kỳ (07d - 360d)
  const [isDailyApproveModalOpen, setIsDailyApproveModalOpen] = useState(false); // Modal 3: Họp Daily Meeting duyệt WR -> WO
  const [isPlanningModalOpen, setIsPlanningModalOpen] = useState(false);   // Modal 4: Engineer Planning LOTO & Mở WP
  const [isCloseWpModalOpen, setIsCloseWpModalOpen] = useState(false);     // Modal 5: Đóng WP nghiệm thu hiện trường
  const [isCloseWoModalOpen, setIsCloseWoModalOpen] = useState(false);     // Modal 6: Đóng WO trên CMMS
  const [isEdit23ColsModalOpen, setIsEdit23ColsModalOpen] = useState(false); // Modal 7: Sửa chi tiết 23 Cột
  const [isFullScreenEdit23Modal, setIsFullScreenEdit23Modal] = useState(true); // Hiển thị toàn màn hình theo yêu cầu
  const [isPrintPermitOpen, setIsPrintPermitOpen] = useState(false);       // Modal 8: In giấy phép Work Permit A4

  // Selected item for active modal action
  const [activeItem, setActiveItem] = useState<Partial<WorkOrderItem>>({});
  const [permitToPrint, setPermitToPrint] = useState<WorkOrderItem | null>(null);

  // Notice alert
  const [actionNotice, setActionNotice] = useState<string | null>(null);

  // Calculate high-level plant KPIs
  const plantKpis = useMemo(() => {
    const total = workOrders.length;
    let pmCount = 0;
    let cmCount = 0;
    let wrPendingCount = 0;
    let woPlanningCount = 0;
    let wpExecutingCount = 0;
    let wpClosedCount = 0;
    let woClosedCount = 0;
    let lotoRequiredCount = 0;
    let p1EmergencyCount = 0;
    let p3UrgentCount = 0;

    let totalDowntimeHours = 0;

    workOrders.forEach(w => {
      const stage = deduceWorkflowStage(w);
      const pri = deducePriorityCode(w);
      const isPm = !w.isUnplanned && (w.type === 'preventive' || w.workStream === 'PM');

      if (isPm) pmCount++;
      else cmCount++;

      if (stage === '1_wr_pending') wrPendingCount++;
      else if (stage === '2_daily_approved' || stage === '3_planning') woPlanningCount++;
      else if (stage === '4_wp_issued') wpExecutingCount++;
      else if (stage === '5_wp_closed') wpClosedCount++;
      else if (stage === '6_wo_closed') woClosedCount++;

      if (w.blockingRequired || w.deEnergyRequired) lotoRequiredCount++;
      if (pri === 'P1') p1EmergencyCount++;
      if (pri === 'P3') p3UrgentCount++;

      if (w.downtimeStart && w.restartTime) {
        const start = new Date(w.downtimeStart).getTime();
        const end = new Date(w.restartTime).getTime();
        if (!isNaN(start) && !isNaN(end) && end > start) {
          totalDowntimeHours += (end - start) / (1000 * 3600);
        }
      }
    });

    return {
      total,
      pmCount,
      cmCount,
      pmPercent: total > 0 ? Math.round((pmCount / total) * 100) : 0,
      wrPendingCount,
      woPlanningCount,
      wpExecutingCount,
      wpClosedCount,
      woClosedCount,
      lotoRequiredCount,
      p1EmergencyCount,
      p3UrgentCount,
      totalDowntimeHours: totalDowntimeHours.toFixed(1)
    };
  }, [workOrders]);

  // Filtered orders list
  const filteredOrders = useMemo(() => {
    return workOrders.filter(w => {
      const stage = deduceWorkflowStage(w);
      const pri = deducePriorityCode(w);
      const isPm = !w.isUnplanned && (w.type === 'preventive' || w.workStream === 'PM');

      // Search
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        const eqStr = Array.isArray(w.equipmentId) ? w.equipmentId.join(' ') : (w.equipmentId || '');
        const match =
          (w.id || '').toLowerCase().includes(q) ||
          (w.wrId || '').toLowerCase().includes(q) ||
          (w.workPermitId || '').toLowerCase().includes(q) ||
          (w.title || '').toLowerCase().includes(q) ||
          (w.description || '').toLowerCase().includes(q) ||
          (w.customer || '').toLowerCase().includes(q) ||
          (w.assignedTo || '').toLowerCase().includes(q) ||
          (w.failureCode || '').toLowerCase().includes(q) ||
          eqStr.toLowerCase().includes(q);
        if (!match) return false;
      }

      // Filter Stream (PM vs CM)
      if (filterStream === 'PM' && !isPm) return false;
      if (filterStream === 'CM' && isPm) return false;

      // Filter Priority
      if (filterPriority !== 'all' && pri !== filterPriority) return false;

      // Filter Stage
      if (filterStage !== 'all' && stage !== filterStage) return false;

      // Filter LOTO
      if (filterLoto === 'yes' && !w.blockingRequired && !w.deEnergyRequired) return false;
      if (filterLoto === 'no' && (w.blockingRequired || w.deEnergyRequired)) return false;

      // Filter Assignee
      if (filterAssignee !== 'all' && w.assignedTo !== filterAssignee) return false;

      return true;
    });
  }, [workOrders, searchQuery, filterStream, filterPriority, filterStage, filterLoto, filterAssignee]);

  // Unique assignees for filter
  const uniqueAssignees = useMemo(() => {
    return Array.from(new Set(workOrders.map(w => w.assignedTo).filter(Boolean))).sort();
  }, [workOrders]);

  // Handler: Open modal to create a new Work Request (WR) from site
  const handleOpenCreateWr = () => {
    const nextWrId = getNextWorkRequestId(workOrders);
    const today = new Date().toISOString().split('T')[0];

    setActiveItem({
      wrId: nextWrId,
      id: '',                           // Chưa có WO ID (sẽ được cấp khi Daily Meeting duyệt)
      workPermitId: '',                 // Chưa có WP ID
      workStream: 'CM',
      isUnplanned: true,
      priorityCode: 'P1',               // Default: P1 để cảnh báo xử lý nhanh
      priority: 'urgent',
      workflowStage: '1_wr_pending',
      status: 'initiated',
      title: '',
      description: '',
      equipmentId: allEquipment[0]?.id || '',
      equipmentName: allEquipment[0]?.name || '',
      customer: customers[0]?.name || 'Tổ hợp Sản xuất Ô tô VinFast',
      customerId: customers[0]?.id || 'KH-VINFAST-01',
      factory: customers[0]?.factories?.[0] || 'Nhà máy Cát Hải - Hải Phòng',
      reportedBy: 'Kỹ sư Hiện trường / Operator',
      reportedAt: new Date().toISOString(),
      dueDate: today,
      blockingRequired: true,
      deEnergyRequired: true,
      failureCode: 'ELEC-01',
      type: 'corrective',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      usedMaterials: ''
    });
    setIsWrModalOpen(true);
  };

  // Handler: Open modal to create a recurring Preventive Maintenance (PM) work plan
  const handleOpenCreatePm = () => {
    const nextWoId = getNextWorkOrderId(workOrders);
    const nextWpId = getNextWorkPermitId(workOrders);
    const today = new Date().toISOString().split('T')[0];

    setActiveItem({
      id: nextWoId,
      workPermitId: nextWpId,
      workStream: 'PM',
      pmCycleDays: 30,                  // Default: 30 ngày
      isUnplanned: false,
      priorityCode: 'P6',               // Priority 6 cho PM
      priority: 'medium',
      workflowStage: '2_daily_approved',
      status: 'approved',
      title: `Bảo dưỡng định kỳ 30 ngày - ${allEquipment[0]?.name || 'Thiết bị trạm'}`,
      description: 'Thực hiện kiểm tra định kỳ 30 ngày: đo điện trở cách điện, kiểm tra mức dầu, chụp ảnh nhiệt tiếp xúc, vệ sinh công nghiệp.',
      equipmentId: allEquipment[0]?.id || '',
      equipmentName: allEquipment[0]?.name || '',
      customer: customers[0]?.name || '',
      customerId: customers[0]?.id || '',
      factory: customers[0]?.factories?.[0] || '',
      assignedTo: 'Võ Minh Tú',
      responsibleApprove: 'TM',
      responsibleDo: 'ME',
      blockingRequired: true,
      deEnergyRequired: true,
      dueDate: new Date(Date.now() + 30 * 86400000).toISOString().split('T')[0],
      failureCode: 'ROUTINE-PM',
      type: 'preventive',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      usedMaterials: 'Mỡ bôi trơn chuyên dụng, giẻ lau công nghiệp, hóa chất làm sạch'
    });
    setIsPmModalOpen(true);
  };

  // Handler: Save Work Request (WR) from site
  const handleSaveWr = async () => {
    if (!activeItem.title || !activeItem.equipmentId) {
      alert('Vui lòng nhập Tiêu đề và chọn Thiết bị cần sửa chữa!');
      return;
    }

    const nextWrId = activeItem.wrId || getNextWorkRequestId(workOrders);
    // Provide a provisional WO placeholder if needed
    const nextWoId = activeItem.id || getNextWorkOrderId(workOrders);
    const nextWpId = activeItem.workPermitId || getNextWorkPermitId(workOrders);

    const payload: WorkOrderItem = {
      id: nextWoId,
      wrId: nextWrId,
      workPermitId: nextWpId,
      title: activeItem.title || '',
      description: activeItem.description || '',
      equipmentId: activeItem.equipmentId || '',
      equipmentName: activeItem.equipmentName || '',
      customer: activeItem.customer || '',
      customerId: activeItem.customerId || '',
      factory: activeItem.factory || '',
      workStream: 'CM',
      isUnplanned: true,
      priorityCode: activeItem.priorityCode || 'P1',
      priority: activeItem.priorityCode === 'P1' ? 'urgent' : activeItem.priorityCode === 'P3' ? 'high' : 'medium',
      workflowStage: '1_wr_pending',
      status: 'initiated',
      failureCode: activeItem.failureCode || 'ELEC-01',
      reportedBy: activeItem.reportedBy || 'Field Service Engineer',
      reportedAt: activeItem.reportedAt || new Date().toISOString(),
      assignedTo: activeItem.assignedTo || 'Chờ phân công',
      responsibleApprove: 'MM',
      responsibleDo: 'Everybody',
      blockingRequired: Boolean(activeItem.blockingRequired),
      deEnergyRequired: Boolean(activeItem.deEnergyRequired),
      dueDate: activeItem.dueDate || '',
      usedMaterials: activeItem.usedMaterials || '',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      downtimeStart: '',
      repairStart: '',
      repairEnd: '',
      restartTime: '',
      type: 'corrective'
    };

    try {
      await onSaveWorkOrder(payload, true);
      setIsWrModalOpen(false);
      setActionNotice(`Đã tạo thành công Phiếu Yêu Cầu WR [${payload.wrId}]! Sẵn sàng đưa vào Daily Meeting duyệt.`);
      setTimeout(() => setActionNotice(null), 5000);
    } catch (e: any) {
      alert('Lỗi lưu WR: ' + e.message);
    }
  };

  // Handler: Save Preventive Maintenance (PM) Plan
  const handleSavePm = async () => {
    if (!activeItem.title || !activeItem.equipmentId) {
      alert('Vui lòng nhập Tiêu đề và chọn Thiết bị bảo dưỡng!');
      return;
    }

    const payload: WorkOrderItem = {
      id: activeItem.id || getNextWorkOrderId(workOrders),
      wrId: activeItem.wrId || '',
      workPermitId: activeItem.workPermitId || getNextWorkPermitId(workOrders),
      title: activeItem.title || '',
      description: activeItem.description || '',
      equipmentId: activeItem.equipmentId || '',
      equipmentName: activeItem.equipmentName || '',
      customer: activeItem.customer || '',
      customerId: activeItem.customerId || '',
      factory: activeItem.factory || '',
      workStream: 'PM',
      pmCycleDays: activeItem.pmCycleDays || 30,
      isUnplanned: false,
      priorityCode: 'P6',
      priority: 'medium',
      workflowStage: '2_daily_approved',
      status: 'approved',
      failureCode: 'ROUTINE-PM',
      assignedTo: activeItem.assignedTo || 'Kỹ sư FSE',
      responsibleApprove: activeItem.responsibleApprove || 'TM',
      responsibleDo: activeItem.responsibleDo || 'ME',
      blockingRequired: Boolean(activeItem.blockingRequired),
      deEnergyRequired: Boolean(activeItem.deEnergyRequired),
      dueDate: activeItem.dueDate || '',
      usedMaterials: activeItem.usedMaterials || '',
      createdAt: activeItem.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      downtimeStart: '',
      repairStart: '',
      repairEnd: '',
      restartTime: '',
      type: 'preventive'
    };

    try {
      await onSaveWorkOrder(payload, true);
      setIsPmModalOpen(false);
      setActionNotice(`Đã thiết lập Lịch PM Định Kỳ [${payload.id}] (Chu kỳ ${payload.pmCycleDays} ngày) thành công!`);
      setTimeout(() => setActionNotice(null), 5000);
    } catch (e: any) {
      alert('Lỗi lưu PM: ' + e.message);
    }
  };

  // Handler: Daily Meeting Approval (WR -> WO)
  const handleApproveWrInDailyMeeting = async () => {
    if (!activeItem.id) {
      alert('Không tìm thấy phiếu!');
      return;
    }

    const updated: WorkOrderItem = {
      ...(activeItem as WorkOrderItem),
      workflowStage: '2_daily_approved',
      status: 'approved',
      dailyMeetingApprovedAt: new Date().toISOString(),
      dailyMeetingApprovedByRole: activeItem.dailyMeetingApprovedByRole || 'MM',
      dailyMeetingApprovedByName: activeItem.dailyMeetingApprovedByName || 'Nguyễn Văn Quản Lý',
      dailyMeetingNotes: activeItem.dailyMeetingNotes || 'Đã thông qua trong cuộc họp Daily Meeting. Yêu cầu lập kế hoạch LOTO và thực hiện đúng thời hạn.',
      assignedTo: activeItem.assignedTo || 'Võ Minh Tú',
      updatedAt: new Date().toISOString()
    };

    try {
      await onSaveWorkOrder(updated, true);
      setIsDailyApproveModalOpen(false);
      setActionNotice(`Họp Daily Meeting: Đã phê duyệt WR [${updated.wrId || updated.id}] ➔ Chuyển thành WO [${updated.id}] và giao cho ${updated.assignedTo}!`);
      setTimeout(() => setActionNotice(null), 5000);
    } catch (e: any) {
      alert('Lỗi phê duyệt: ' + e.message);
    }
  };

  // Handler: Save Engineer Planning & Issue Work Permit (WP)
  const handleSavePlanningAndIssueWp = async () => {
    if (!activeItem.id) return;

    const updated: WorkOrderItem = {
      ...(activeItem as WorkOrderItem),
      workflowStage: '4_wp_issued',
      status: 'in-progress',
      workPermitId: activeItem.workPermitId || getNextWorkPermitId(workOrders),
      wpIssuedAt: new Date().toISOString(),
      wpIssuedBy: activeItem.wpIssuedBy || 'Cán bộ An toàn (SO)',
      lotoTagNumbers: activeItem.lotoTagNumbers || 'LOTO-TR-01',
      isolationPoints: activeItem.isolationPoints || 'Máy cắt Q01, Dao cách ly DS01, Khóa van dầu',
      deEnergyRequired: Boolean(activeItem.deEnergyRequired),
      blockingRequired: Boolean(activeItem.blockingRequired),
      usedMaterials: activeItem.usedMaterials || '',
      updatedAt: new Date().toISOString(),
      repairStart: activeItem.repairStart || new Date().toISOString()
    };

    try {
      await onSaveWorkOrder(updated, true);
      setIsPlanningModalOpen(false);
      setActionNotice(`Đã hoàn tất Planning & Cấp Giấy Phép Work Permit [${updated.workPermitId}]! Hiện trường bắt đầu thi công.`);
      setTimeout(() => setActionNotice(null), 5000);
    } catch (e: any) {
      alert('Lỗi cấp WP: ' + e.message);
    }
  };

  // Handler: Close Work Permit (WP) at Site
  const handleCloseWorkPermit = async () => {
    if (!activeItem.id) return;

    const updated: WorkOrderItem = {
      ...(activeItem as WorkOrderItem),
      workflowStage: '5_wp_closed',
      status: 'in-progress',
      wpClosedAt: new Date().toISOString(),
      wpClosedBy: activeItem.wpClosedBy || 'Kỹ sư FSE / Cán bộ An toàn',
      wpCloseNotes: activeItem.wpCloseNotes || 'Hiện trường đã dọn dẹp vệ sinh sạch sẽ, tháo toàn bộ tiếp địa và gỡ bỏ thẻ khóa LOTO an toàn.',
      repairEnd: new Date().toISOString(),
      restartTime: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    try {
      await onSaveWorkOrder(updated, true);
      setIsCloseWpModalOpen(false);
      setActionNotice(`Đã đóng Giấy Phép Work Permit [${updated.workPermitId}] an toàn! Có thể đóng WO trên CMMS.`);
      setTimeout(() => setActionNotice(null), 5000);
    } catch (e: any) {
      alert('Lỗi đóng WP: ' + e.message);
    }
  };

  // Handler: Final CMMS WO Closure
  const handleCloseWorkOrderOnCmms = async () => {
    if (!activeItem.id) return;

    const updated: WorkOrderItem = {
      ...(activeItem as WorkOrderItem),
      workflowStage: '6_wo_closed',
      status: 'completed',
      woClosedAt: new Date().toISOString(),
      woClosedBy: activeItem.woClosedBy || 'Maintenance Manager / Plant CMMS Admin',
      rootCause: activeItem.rootCause || 'Đã khắc phục hoàn toàn theo quy trình kỹ thuật',
      actualTimeSpent: activeItem.actualTimeSpent || 3.0,
      updatedAt: new Date().toISOString()
    };

    try {
      await onSaveWorkOrder(updated, true);
      setIsCloseWoModalOpen(false);
      setActionNotice(`🎉 Đã đóng thành công WO [${updated.id}] trên CMMS! Toàn bộ chu trình bảo dưỡng khép kín.`);
      setTimeout(() => setActionNotice(null), 5000);
    } catch (e: any) {
      alert('Lỗi đóng WO: ' + e.message);
    }
  };

  // Handler: 1-Touch Timestamps (Cols 20-23)
  const handleRecordTimestamp = async (
    order: WorkOrderItem,
    field: 'downtimeStart' | 'repairStart' | 'repairEnd' | 'restartTime'
  ) => {
    const nowIso = new Date().toISOString();
    const updated: WorkOrderItem = {
      ...order,
      [field]: nowIso,
      updatedAt: nowIso
    };

    try {
      await onSaveWorkOrder(updated, true);
      setActionNotice(`Đã ghi nhận mốc ${field} cho [${order.id}] lúc ${new Date().toLocaleTimeString('vi-VN')}`);
      setTimeout(() => setActionNotice(null), 3500);
    } catch (e) {
      console.error(e);
    }
  };

  // Handler: Export 23-column CSV
  const handleExport23ColumnsCsv = () => {
    const headers = [
      '1. Mã WO',
      '2. Mã Work Permit',
      '3. Tiêu đề công việc',
      '4. Mô tả chi tiết',
      '5. Mã / Tên thiết bị',
      '6. Failure Code',
      '7. Khách hàng / Nhà máy',
      '8. Loại công việc (PM/CM)',
      '9. Ngoài kế hoạch?',
      '10. Mức độ ưu tiên (P1/P3/P6/P7/P9)',
      '11. Trạng thái vòng đời CMMS',
      '12. Người thực hiện (PIC)',
      '13. Vai trò Phê duyệt (Approve)',
      '14. Vai trò Thực hiện (Do)',
      '15. Yêu cầu cô lập LOTO',
      '16. Hạn hoàn thành (SLA)',
      '17. Vật tư sử dụng',
      '18. Ngày tạo (Created)',
      '19. Cập nhật cuối (Updated)',
      '20. Bắt đầu dừng máy',
      '21. Bắt đầu sửa chữa',
      '22. Kết thúc sửa chữa',
      '23. Chạy lại máy'
    ];

    const rows = filteredOrders.map(w => {
      const pri = deducePriorityCode(w);
      const stage = deduceWorkflowStage(w);
      const equipStr = Array.isArray(w.equipmentId) ? w.equipmentId.join('; ') : (w.equipmentId || w.equipmentName || '');
      const materialsStr = Array.isArray(w.usedMaterials)
        ? w.usedMaterials.map((m: any) => `${m.name || m}`).join('; ')
        : (w.usedMaterials || '');

      return [
        `"${w.id || ''}"`,
        `"${w.workPermitId || ''}"`,
        `"${(w.title || '').replace(/"/g, '""')}"`,
        `"${(w.description || '').replace(/"/g, '""')}"`,
        `"${equipStr.replace(/"/g, '""')}"`,
        `"${w.failureCode || ''}"`,
        `"${(w.customer || '').replace(/"/g, '""')}"`,
        `"${w.workStream || (w.isUnplanned ? 'CM' : 'PM')}"`,
        `"${w.isUnplanned ? 'Có (CM)' : 'Không (PM)'}"`,
        `"${pri}"`,
        `"${WORKFLOW_STAGES[stage].label}"`,
        `"${w.assignedTo || ''}"`,
        `"${w.responsibleApprove || w.dailyMeetingApprovedByRole || 'MM'}"`,
        `"${w.responsibleDo || 'Everybody'}"`,
        `"${w.blockingRequired || w.deEnergyRequired ? 'Có (LOTO)' : 'Không'}"`,
        `"${w.dueDate || ''}"`,
        `"${materialsStr.replace(/"/g, '""')}"`,
        `"${w.createdAt || ''}"`,
        `"${w.updatedAt || ''}"`,
        `"${w.downtimeStart || ''}"`,
        `"${w.repairStart || ''}"`,
        `"${w.repairEnd || ''}"`,
        `"${w.restartTime || ''}"`
      ].join(',');
    });

    const csvContent = '\uFEFF' + [headers.join(','), ...rows].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `TEV_Plant_CMMS_WR_WO_WP_23Col_${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="space-y-5">
      {/* ==================================================================== */}
      {/* 1. PLANT ASSET MANAGEMENT & WORKFLOW LIFECYCLE HEADER                */}
      {/* ==================================================================== */}
      <div className="bg-gradient-to-r from-slate-900 via-slate-800 to-indigo-950 rounded-2xl p-6 text-white border border-slate-700 shadow-md">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6">
          <div className="space-y-2.5 max-w-3xl">
            <div className="flex flex-wrap items-center gap-2">
              <span className="px-3 py-1 rounded-full text-xs font-bold uppercase tracking-wider bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 flex items-center gap-1.5">
                <CheckCircle size={14} className="text-emerald-400" />
                Chuẩn CMMS Nhà Máy TEV
              </span>
              <span className="px-3 py-1 rounded-full text-xs font-medium bg-blue-500/20 text-blue-300 border border-blue-500/30 flex items-center gap-1.5">
                <Shield size={14} className="text-blue-400" />
                WR ➔ Họp Daily Meeting ➔ WO ➔ WP (LOTO) ➔ Đóng CMMS
              </span>
              {isGoogleConnected && (
                <span className="px-3 py-1 rounded-full text-xs font-medium bg-cyan-500/20 text-cyan-300 border border-cyan-500/30 flex items-center gap-1.5">
                  <span className="w-2 h-2 rounded-full bg-cyan-400 animate-pulse"></span>
                  Đồng bộ Google Sheet: TEV Service Flatform
                </span>
              )}
            </div>

            <h1 className="text-2xl sm:text-3xl font-extrabold tracking-tight text-white flex items-center gap-3">
              <ClipboardList className="text-blue-400" size={30} />
              Quản Trị Phiếu Công Việc & Giấy Phép Work Permit Chuẩn Nhà Máy
            </h1>

            <p className="text-xs sm:text-sm text-slate-300 leading-relaxed">
              Quy trình khép kín: <strong>Bảo trì định kỳ PM</strong> (07d, 14d, 30d, 60d, 90d, 180d, 360d - Priority 6) và 
              <strong> Sửa chữa sự cố CM</strong> (Field Service / Operator mở phiếu WR ➔ Họp Daily Meeting phê duyệt: MM / OM / DM / PM ➔ Chuyển thành WO ➔ Kỹ sư Planning LOTO / De-energize / Vật tư ➔ Cấp Work Permit thi công ➔ Đóng WP & Đóng WO).
            </p>
          </div>

          {/* Action Buttons: 2 Core Entrypoints (WR for CM & Create PM Plan) */}
          <div className="flex flex-col sm:flex-row lg:flex-col gap-2.5 shrink-0">
            <div className="flex items-center gap-2">
              {/* Button 1: Operator / Field Service opens WR */}
              <button
                onClick={handleOpenCreateWr}
                className="flex-1 flex items-center justify-center gap-2 px-4 py-2.5 bg-gradient-to-r from-rose-600 to-amber-600 hover:from-rose-700 hover:to-amber-700 text-white font-bold text-xs rounded-xl shadow-md transition-all hover:scale-[1.02] active:scale-[0.98]"
                title="Field Service / Operator đi site phát hiện khiếm khuyết mở phiếu WR/CM mới"
              >
                <AlertTriangle size={16} />
                <span>+ Mở Yêu Cầu WR/CM</span>
              </button>

              {/* Button 2: Create recurring PM Plan */}
              <button
                onClick={handleOpenCreatePm}
                className="flex-1 flex items-center justify-center gap-2 px-4 py-2.5 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white font-bold text-xs rounded-xl shadow-md transition-all hover:scale-[1.02] active:scale-[0.98]"
                title="Lập kế hoạch bảo trì định kỳ PM (7d, 14d, 30d, 60d, 90d, 180d, 360d - Priority 6)"
              >
                <Calendar size={16} />
                <span>+ Lập Kế Hoạch PM (07d - 360d)</span>
              </button>
            </div>

            {/* Sub utilities */}
            <div className="flex items-center gap-2">
              {onFetchFromSheets && (
                <button
                  onClick={() => onFetchFromSheets()}
                  disabled={isSyncing}
                  className="flex-1 flex items-center justify-center gap-1.5 px-3 py-2 bg-slate-800 hover:bg-slate-700 border border-slate-600 text-slate-200 text-xs font-semibold rounded-lg transition-colors disabled:opacity-50"
                  title="Nạp dữ liệu từ Google Sheets TEV Service Flatform"
                >
                  <RefreshCw size={13} className={isSyncing ? 'animate-spin text-cyan-400' : 'text-cyan-400'} />
                  <span>Nạp Sheets</span>
                </button>
              )}
              {onSyncAllToSheets && (
                <button
                  onClick={() => onSyncAllToSheets()}
                  disabled={isSyncing}
                  className="flex-1 flex items-center justify-center gap-1.5 px-3 py-2 bg-slate-800 hover:bg-slate-700 border border-slate-600 text-slate-200 text-xs font-semibold rounded-lg transition-colors disabled:opacity-50"
                  title="Đẩy toàn bộ phiếu công việc lên Google Sheets TEV Service Flatform"
                >
                  <FileSpreadsheet size={13} className="text-emerald-400" />
                  <span>Đẩy Sheets</span>
                </button>
              )}
              <button
                onClick={handleExport23ColumnsCsv}
                className="flex items-center justify-center gap-1.5 px-3 py-2 bg-slate-800 hover:bg-slate-700 border border-slate-600 text-slate-200 text-xs font-semibold rounded-lg transition-colors"
                title="Xuất bảng Excel/CSV Chuẩn TEV"
              >
                <Download size={13} className="text-amber-400" />
                <span>Xuất CSV</span>
              </button>
            </div>
          </div>
        </div>

        {/* Action Notice */}
        {actionNotice && (
          <div className="mt-4 p-3 bg-emerald-950/80 border border-emerald-500/60 rounded-xl text-emerald-200 text-xs font-medium flex items-center justify-between animate-fadeIn">
            <span className="flex items-center gap-2">
              <CheckCircle size={16} className="text-emerald-400 shrink-0" />
              {actionNotice}
            </span>
            <button onClick={() => setActionNotice(null)} className="text-emerald-400 hover:text-white">
              <X size={14} />
            </button>
          </div>
        )}
      </div>

      {/* ==================================================================== */}
      {/* 2. INTERACTIVE PIPELINE BREADCRUMB: 6 STAGES OF PLANT CMMS           */}
      {/* ==================================================================== */}
      <div className="bg-white rounded-2xl p-3 border border-slate-200 shadow-xs">
        <div className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-2 flex items-center justify-between">
          <span>Quy Trình Quản Lý Vòng Đời Phiếu Công Việc (Nhấn vào từng bước để lọc)</span>
          <span className="font-mono text-slate-700">Tổng: {plantKpis.total} phiếu</span>
        </div>

        <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-2">
          {Object.entries(WORKFLOW_STAGES).map(([key, info]) => {
            const count = workOrders.filter(w => deduceWorkflowStage(w) === key).length;
            const isSelected = filterStage === key;

            return (
              <button
                key={key}
                onClick={() => setFilterStage(isSelected ? 'all' : key)}
                className={`p-2.5 rounded-xl border text-left transition-all relative overflow-hidden ${
                  isSelected
                    ? 'border-blue-600 bg-blue-50/70 shadow-xs ring-2 ring-blue-500/20'
                    : 'border-slate-200 hover:border-slate-300 bg-slate-50/50 hover:bg-slate-100/50'
                }`}
              >
                <div className="flex items-center justify-between mb-1">
                  <span className="text-[10px] font-bold text-slate-400">BƯỚC {info.step}</span>
                  <span className={`px-1.5 py-0.2 rounded-full text-[10px] font-black ${
                    count > 0 ? 'bg-slate-900 text-white' : 'bg-slate-200 text-slate-600'
                  }`}>
                    {count}
                  </span>
                </div>
                <div className="font-bold text-xs text-slate-900 truncate" title={info.label}>
                  {info.shortTitle}
                </div>
                <div className="text-[10px] text-slate-500 truncate mt-0.5" title={info.desc}>
                  {info.desc}
                </div>
              </button>
            );
          })}
        </div>
      </div>

      {/* ==================================================================== */}
      {/* 3. PLANT ASSET KPI SUMMARY CARDS                                     */}
      {/* ==================================================================== */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3 text-xs">
        {/* Stream Ratio: PM vs CM */}
        <div className="bg-white p-3.5 rounded-xl border border-slate-200 shadow-2xs">
          <div className="flex items-center justify-between text-slate-400 mb-1">
            <span className="font-bold text-[10px] uppercase">Cơ Cấu PM / CM</span>
            <Layers size={14} className="text-blue-500" />
          </div>
          <div className="text-xl font-black text-slate-900">
            {plantKpis.pmPercent}% <span className="text-xs font-normal text-slate-500">PM</span>
          </div>
          <div className="mt-1 text-[11px] text-slate-500 flex items-center justify-between">
            <span>PM: {plantKpis.pmCount}</span>
            <span>CM: {plantKpis.cmCount}</span>
          </div>
        </div>

        {/* Chờ Họp Daily Meeting Duyệt (WR Pending) */}
        <div className={`p-3.5 rounded-xl border shadow-2xs transition-all ${
          plantKpis.wrPendingCount > 0 ? 'bg-amber-50/60 border-amber-200' : 'bg-white border-slate-200'
        }`}>
          <div className="flex items-center justify-between text-slate-400 mb-1">
            <span className="font-bold text-[10px] uppercase text-amber-800">Chờ Họp Daily Duyệt</span>
            <Users size={14} className="text-amber-600" />
          </div>
          <div className={`text-xl font-black ${plantKpis.wrPendingCount > 0 ? 'text-amber-700' : 'text-slate-800'}`}>
            {plantKpis.wrPendingCount} <span className="text-xs font-normal text-slate-500">WR</span>
          </div>
          <div className="mt-1 text-[11px] text-amber-800 font-medium">
            {plantKpis.wrPendingCount > 0 ? 'Cần MM/OM duyệt' : 'Đã duyệt hết'}
          </div>
        </div>

        {/* Priority 1 (Làm ngay) & Priority 3 */}
        <div className={`p-3.5 rounded-xl border shadow-2xs transition-all ${
          plantKpis.p1EmergencyCount > 0 ? 'bg-rose-50/60 border-rose-200' : 'bg-white border-slate-200'
        }`}>
          <div className="flex items-center justify-between text-slate-400 mb-1">
            <span className="font-bold text-[10px] uppercase text-rose-800">Sự Cố P1 & P3</span>
            <Flame size={14} className="text-rose-600" />
          </div>
          <div className="text-xl font-black text-rose-700">
            {plantKpis.p1EmergencyCount} <span className="text-xs font-semibold text-slate-600">P1</span> / {plantKpis.p3UrgentCount} <span className="text-xs font-semibold text-slate-600">P3</span>
          </div>
          <div className="mt-1 text-[11px] text-slate-500">
            {plantKpis.p1EmergencyCount > 0 ? 'Cần can thiệp khẩn cấp' : 'Không có sự cố P1'}
          </div>
        </div>

        {/* Đang Thi Công (WP Mở tại site) */}
        <div className="bg-white p-3.5 rounded-xl border border-slate-200 shadow-2xs">
          <div className="flex items-center justify-between text-slate-400 mb-1">
            <span className="font-bold text-[10px] uppercase">WP Mở Tại Site</span>
            <Wrench size={14} className="text-purple-600" />
          </div>
          <div className="text-xl font-black text-purple-700">
            {plantKpis.wpExecutingCount} <span className="text-xs font-normal text-slate-500">phiếu</span>
          </div>
          <div className="mt-1 text-[11px] text-slate-500">
            Đang làm việc tại hiện trường
          </div>
        </div>

        {/* Kiểm Soát LOTO / De-energize */}
        <div className="bg-white p-3.5 rounded-xl border border-slate-200 shadow-2xs">
          <div className="flex items-center justify-between text-slate-400 mb-1">
            <span className="font-bold text-[10px] uppercase">Cô Lập Nguồn (LOTO)</span>
            <Lock size={14} className="text-indigo-600" />
          </div>
          <div className="text-xl font-black text-indigo-900">
            {plantKpis.lotoRequiredCount} <span className="text-xs font-normal text-slate-500">phiếu</span>
          </div>
          <div className="mt-1 text-[11px] text-indigo-700 font-medium">
            100% tuân thủ an toàn
          </div>
        </div>

        {/* Downtime Hours Tracker (Cols 20-23) */}
        <div className="bg-white p-3.5 rounded-xl border border-slate-200 shadow-2xs">
          <div className="flex items-center justify-between text-slate-400 mb-1">
            <span className="font-bold text-[10px] uppercase">Giờ Dừng Máy</span>
            <Clock size={14} className="text-amber-600" />
          </div>
          <div className="text-xl font-black text-amber-700">
            {plantKpis.totalDowntimeHours} <span className="text-xs font-normal text-slate-500">h</span>
          </div>
          <div className="mt-1 text-[11px] text-slate-500">
            Tính từ mốc Dừng máy ➔ Chạy lại
          </div>
        </div>
      </div>

      {/* ==================================================================== */}
      {/* 4. MAIN SUB-TABS: WORKFLOW-DRIVEN SECTIONS                           */}
      {/* ==================================================================== */}
      <div className="bg-white rounded-2xl p-1.5 border border-slate-200 shadow-xs flex flex-wrap gap-1.5 text-xs">
        {/* TAB 1: Dashboard & Tiến Độ PM displayed first as requested */}
        <button
          onClick={() => setActiveSubTab('pm-dashboard')}
          className={`flex items-center gap-2 py-2 px-3.5 rounded-xl font-bold transition-all ${
            activeSubTab === 'pm-dashboard'
              ? 'bg-blue-600 text-white shadow-xs'
              : 'text-slate-600 hover:bg-slate-100'
          }`}
        >
          <BarChart3 size={15} />
          <span>Dashboard & Tiến Độ PM</span>
        </button>

        {/* TAB 2: Danh Sách & Quản Lý Phiếu */}
        <button
          onClick={() => setActiveSubTab('all-work')}
          className={`flex items-center gap-2 py-2 px-3.5 rounded-xl font-bold transition-all ${
            activeSubTab === 'all-work'
              ? 'bg-blue-600 text-white shadow-xs'
              : 'text-slate-600 hover:bg-slate-100'
          }`}
        >
          <ClipboardList size={15} />
          <span>Danh Sách & Quản Lý Phiếu</span>
          <span className={`text-[10px] px-1.5 py-0.2 rounded-full font-black ${
            activeSubTab === 'all-work' ? 'bg-blue-500 text-white' : 'bg-slate-200 text-slate-700'
          }`}>
            {filteredOrders.length}
          </span>
        </button>

        <button
          onClick={() => setActiveSubTab('daily-meeting')}
          className={`flex items-center gap-2 py-2 px-3.5 rounded-xl font-bold transition-all ${
            activeSubTab === 'daily-meeting'
              ? 'bg-blue-600 text-white shadow-xs'
              : 'text-slate-600 hover:bg-slate-100'
          }`}
        >
          <Users size={15} />
          <span>Họp Giao Ban Daily Meeting (Duyệt WR ➔ WO)</span>
          {plantKpis.wrPendingCount > 0 && (
            <span className="text-[10px] px-1.5 py-0.2 rounded-full font-black bg-amber-400 text-slate-950 animate-pulse">
              {plantKpis.wrPendingCount} cần duyệt
            </span>
          )}
        </button>

        <button
          onClick={() => setActiveSubTab('pm-matrix')}
          className={`flex items-center gap-2 py-2 px-3.5 rounded-xl font-bold transition-all ${
            activeSubTab === 'pm-matrix'
              ? 'bg-blue-600 text-white shadow-xs'
              : 'text-slate-600 hover:bg-slate-100'
          }`}
        >
          <Calendar size={15} />
          <span>Ma Trận Chu Kỳ PM (07d - 360d)</span>
        </button>

        <button
          onClick={() => setActiveSubTab('wp-loto')}
          className={`flex items-center gap-2 py-2 px-3.5 rounded-xl font-bold transition-all ${
            activeSubTab === 'wp-loto'
              ? 'bg-blue-600 text-white shadow-xs'
              : 'text-slate-600 hover:bg-slate-100'
          }`}
        >
          <Lock size={15} />
          <span>Giấy Phép Work Permit & An Toàn LOTO ({plantKpis.lotoRequiredCount})</span>
        </button>

        <button
          onClick={() => setActiveSubTab('alerts')}
          className={`flex items-center gap-2 py-2 px-3.5 rounded-xl font-bold transition-all ${
            activeSubTab === 'alerts'
              ? 'bg-blue-600 text-white shadow-xs'
              : 'text-slate-600 hover:bg-slate-100'
          }`}
        >
          <Mail size={15} />
          <span>Cảnh Báo Tự Động (Cloud Functions)</span>
          {pmUpcomingTasks.length > 0 && (
            <span className="text-[10px] px-1.5 py-0.2 rounded-full font-black bg-amber-400 text-slate-950">
              {pmUpcomingTasks.length}
            </span>
          )}
        </button>
      </div>

      {/* ==================================================================== */}
      {/* 5. TAB 1: ALL WORK ORDERS & WORK PERMITS (23 COLS MATRIX & VIEWS)    */}
      {/* ==================================================================== */}
      {activeSubTab === 'all-work' && (
        <div className="space-y-4">
          {/* Advanced Plant Filters */}
          <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-xs space-y-3">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
              {/* Search */}
              <div className="relative flex-1">
                <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                <input
                  type="text"
                  placeholder="Tìm theo Mã WO, WR, WP, Thiết bị, Khách hàng, Failure Code, Kỹ sư..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs sm:text-sm focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 font-medium"
                />
              </div>

              {/* View Switcher */}
              <div className="flex items-center gap-1 bg-slate-100 p-1 rounded-xl shrink-0 border border-slate-200">
                <button
                  onClick={() => setViewMode('full-23-cols')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                    viewMode === 'full-23-cols' ? 'bg-white text-blue-700 shadow-xs' : 'text-slate-600 hover:text-slate-900'
                  }`}
                  title="Chế độ Bảng Chi Tiết Chuẩn TEV"
                >
                  📊 Bảng Chi Tiết
                </button>
                <button
                  onClick={() => setViewMode('workflow-kanban')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                    viewMode === 'workflow-kanban' ? 'bg-white text-blue-700 shadow-xs' : 'text-slate-600 hover:text-slate-900'
                  }`}
                  title="Chế độ Bảng Quy Trình 6 Bước CMMS"
                >
                  📌 Quy Trình 6 Bước
                </button>
                <button
                  onClick={() => setViewMode('compact')}
                  className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                    viewMode === 'compact' ? 'bg-white text-blue-700 shadow-xs' : 'text-slate-600 hover:text-slate-900'
                  }`}
                  title="Chế độ Thẻ Vận Hành Gọn Gàng"
                >
                  📱 Thẻ Vận Hành
                </button>
              </div>
            </div>

            {/* Filter Dropdowns */}
            <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 gap-2 text-xs">
              {/* Dòng Công Việc PM vs CM */}
              <div>
                <label className="block text-[10px] font-bold text-slate-500 mb-0.5">Dòng công việc</label>
                <select
                  value={filterStream}
                  onChange={(e) => setFilterStream(e.target.value as any)}
                  className="w-full px-2 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-semibold text-slate-700"
                >
                  <option value="all">Tất cả (PM & CM)</option>
                  <option value="PM">Bảo trì định kỳ PM (7d - 360d)</option>
                  <option value="CM">Sửa chữa sự cố CM (WR ➔ WO)</option>
                </select>
              </div>

              {/* Priority Standard: P1, P3, P6, P7, P9 */}
              <div>
                <label className="block text-[10px] font-bold text-slate-500 mb-0.5">Mức ưu tiên (Priority)</label>
                <select
                  value={filterPriority}
                  onChange={(e) => setFilterPriority(e.target.value)}
                  className="w-full px-2 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-semibold text-slate-700"
                >
                  <option value="all">Tất cả Priority</option>
                  <option value="P1">P1: Làm ngay (Emergency)</option>
                  <option value="P3">P3: Trong 1 tuần (Urgent)</option>
                  <option value="P6">P6: Bảo trì định kỳ (Routine PM)</option>
                  <option value="P7">P7: Shutdown (Dừng trạm)</option>
                  <option value="P9">P9: DMR (Cải tiến thiết kế)</option>
                </select>
              </div>

              {/* Vòng đời Stage */}
              <div>
                <label className="block text-[10px] font-bold text-slate-500 mb-0.5">Giai đoạn vòng đời</label>
                <select
                  value={filterStage}
                  onChange={(e) => setFilterStage(e.target.value)}
                  className="w-full px-2 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-semibold text-slate-700"
                >
                  <option value="all">Tất cả giai đoạn</option>
                  {Object.entries(WORKFLOW_STAGES).map(([k, v]) => (
                    <option key={k} value={k}>{v.label}</option>
                  ))}
                </select>
              </div>

              {/* LOTO */}
              <div>
                <label className="block text-[10px] font-bold text-slate-500 mb-0.5">Cô lập nguồn (LOTO)</label>
                <select
                  value={filterLoto}
                  onChange={(e) => setFilterLoto(e.target.value)}
                  className="w-full px-2 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-semibold text-slate-700"
                >
                  <option value="all">Tất cả</option>
                  <option value="yes">Bắt buộc LOTO / De-energize</option>
                  <option value="no">Không cô lập</option>
                </select>
              </div>

              {/* Kỹ sư PIC */}
              <div>
                <label className="block text-[10px] font-bold text-slate-500 mb-0.5">Kỹ sư phụ trách</label>
                <select
                  value={filterAssignee}
                  onChange={(e) => setFilterAssignee(e.target.value)}
                  className="w-full px-2 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-semibold text-slate-700"
                >
                  <option value="all">Tất cả Kỹ sư</option>
                  {uniqueAssignees.map(a => (
                    <option key={a} value={a}>{a}</option>
                  ))}
                </select>
              </div>
            </div>
          </div>

          {/* VIEW MODE 1: FULL 23-COLUMN MATRIX (The main requirement) */}
          {viewMode === 'full-23-cols' && (
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="p-3 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-2 text-xs">
                <div className="flex items-center gap-2">
                  <span className="font-bold text-slate-800">
                    Bảng Quản Trị Chuẩn TEV Service Flatform
                  </span>
                  <span className="px-2 py-0.5 bg-blue-100 text-blue-800 rounded font-bold">
                    {filteredOrders.length} phiếu hiển thị
                  </span>
                </div>
                <div className="flex items-center gap-3 text-[11px] text-slate-500">
                  <span className="flex items-center gap-1">
                    <span className="w-2.5 h-2.5 rounded bg-blue-500"></span> Định danh &amp; WP
                  </span>
                  <span className="flex items-center gap-1">
                    <span className="w-2.5 h-2.5 rounded bg-amber-500"></span> Phân loại &amp; KH
                  </span>
                  <span className="flex items-center gap-1">
                    <span className="w-2.5 h-2.5 rounded bg-indigo-500"></span> Phân công RACI
                  </span>
                  <span className="flex items-center gap-1">
                    <span className="w-2.5 h-2.5 rounded bg-rose-500"></span> LOTO &amp; Vật tư
                  </span>
                  <span className="flex items-center gap-1">
                    <span className="w-2.5 h-2.5 rounded bg-emerald-500"></span> Mốc giờ dừng máy
                  </span>
                </div>
              </div>

              <div className="overflow-x-auto max-h-[650px] relative">
                <table className="w-full text-left border-collapse text-xs">
                  <thead className="sticky top-0 z-20 bg-slate-100/95 backdrop-blur-xs shadow-xs text-slate-700 font-bold uppercase tracking-wider text-[11px] border-b border-slate-300">
                    <tr>
                      <th className="p-3 bg-slate-100 sticky left-0 z-30 shadow-xs border-r border-slate-200 min-w-[130px]">
                        1. Mã WO
                      </th>
                      <th className="p-3 border-r border-slate-200 min-w-[130px]">2. Mã Work Permit</th>
                      <th className="p-3 border-r border-slate-200 min-w-[220px]">3. Tiêu đề công việc</th>
                      <th className="p-3 border-r border-slate-200 min-w-[180px]">4. Mô tả chi tiết</th>
                      <th className="p-3 border-r border-slate-200 min-w-[160px]">5. Thiết bị</th>
                      <th className="p-3 border-r border-slate-200 min-w-[140px]">6. Failure Code</th>
                      <th className="p-3 border-r border-slate-200 min-w-[160px]">7. Khách hàng</th>
                      <th className="p-3 border-r border-slate-200 min-w-[130px]">8. Loại CV (PM/CM)</th>
                      <th className="p-3 border-r border-slate-200 min-w-[100px]">9. Ngoài KH?</th>
                      <th className="p-3 border-r border-slate-200 min-w-[130px]">10. Ưu tiên (Priority)</th>
                      <th className="p-3 border-r border-slate-200 min-w-[170px]">11. Vòng đời CMMS</th>
                      <th className="p-3 border-r border-slate-200 min-w-[140px]">12. PIC thực hiện</th>
                      <th className="p-3 border-r border-slate-200 min-w-[110px]">13. Phê duyệt</th>
                      <th className="p-3 border-r border-slate-200 min-w-[110px]">14. Thực hiện</th>
                      <th className="p-3 border-r border-slate-200 min-w-[120px]">15. Cô lập LOTO</th>
                      <th className="p-3 border-r border-slate-200 min-w-[110px]">16. Hạn hoàn thành</th>
                      <th className="p-3 border-r border-slate-200 min-w-[180px]">17. Vật tư sử dụng</th>
                      <th className="p-3 border-r border-slate-200 min-w-[130px]">18. Ngày tạo</th>
                      <th className="p-3 border-r border-slate-200 min-w-[130px]">19. Cập nhật</th>
                      <th className="p-3 border-r border-slate-200 min-w-[150px] bg-rose-50/50 text-rose-800">
                        20. Bắt đầu dừng máy
                      </th>
                      <th className="p-3 border-r border-slate-200 min-w-[150px] bg-amber-50/50 text-amber-800">
                        21. Bắt đầu sửa
                      </th>
                      <th className="p-3 border-r border-slate-200 min-w-[150px] bg-blue-50/50 text-blue-800">
                        22. Kết thúc sửa
                      </th>
                      <th className="p-3 border-r border-slate-200 min-w-[150px] bg-emerald-50/50 text-emerald-800">
                        23. Chạy lại máy
                      </th>
                      <th className="p-3 sticky right-0 z-30 bg-slate-100 shadow-xs text-right min-w-[200px]">
                        Thao tác theo Quy trình
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {filteredOrders.length === 0 ? (
                      <tr>
                        <td colSpan={24} className="p-10 text-center text-slate-400">
                          Không tìm thấy phiếu công việc nào phù hợp với bộ lọc hiện tại.
                        </td>
                      </tr>
                    ) : (
                      filteredOrders.map((wo, index) => {
                        const pri = deducePriorityCode(wo);
                        const stage = deduceWorkflowStage(wo);
                        const priCfg = PLANT_PRIORITIES[pri];
                        const stageCfg = WORKFLOW_STAGES[stage];
                        const isPm = !wo.isUnplanned && (wo.type === 'preventive' || wo.workStream === 'PM');
                        const equipStr = Array.isArray(wo.equipmentId) ? wo.equipmentId.join(', ') : (wo.equipmentId || wo.equipmentName || '---');
                        const materialsStr = Array.isArray(wo.usedMaterials)
                          ? wo.usedMaterials.map((m: any) => `${m.name || m}`).join(', ')
                          : (wo.usedMaterials || 'Chưa sử dụng');

                        return (
                          <tr
                            key={wo.id || index}
                            className="hover:bg-blue-50/40 transition-colors group"
                          >
                            {/* 1. Mã WO (Sticky Left) */}
                            <td className="p-3 bg-white group-hover:bg-blue-50/40 sticky left-0 z-10 shadow-xs border-r border-slate-200 font-mono font-bold text-emerald-700">
                              <button
                                onClick={() => {
                                  setActiveItem(wo);
                                  setIsEdit23ColsModalOpen(true);
                                }}
                                className="hover:underline flex items-center gap-1"
                                title="Nhấn để xem / chỉnh sửa phiếu"
                              >
                                {wo.id || wo.wrId || 'WO-CHƯA-DUYỆT'}
                              </button>
                              {wo.wrId && (
                                <div className="text-[10px] text-slate-400 font-normal">
                                  WR: {wo.wrId}
                                </div>
                              )}
                            </td>

                            {/* 2. Mã WP */}
                            <td className="p-3 border-r border-slate-200 font-mono font-bold text-blue-800">
                              {wo.workPermitId ? (
                                <span className="px-1.5 py-0.5 rounded bg-blue-50 border border-blue-200">
                                  {wo.workPermitId}
                                </span>
                              ) : (
                                <span className="text-slate-400 italic">Chưa cấp</span>
                              )}
                            </td>

                            {/* 3. Tiêu đề */}
                            <td className="p-3 border-r border-slate-200 font-semibold text-slate-900 max-w-[220px] truncate" title={wo.title}>
                              {wo.title}
                            </td>

                            {/* 4. Mô tả */}
                            <td className="p-3 border-r border-slate-200 text-slate-600 max-w-[180px] truncate" title={wo.description}>
                              {wo.description || '---'}
                            </td>

                            {/* 5. Thiết bị */}
                            <td className="p-3 border-r border-slate-200 font-medium text-slate-800">
                              <div className="truncate max-w-[160px]" title={equipStr}>
                                {equipStr}
                              </div>
                            </td>

                            {/* 6. Failure Code */}
                            <td className="p-3 border-r border-slate-200">
                              <span className="px-2 py-0.5 rounded font-mono font-bold text-[10px] bg-rose-50 text-rose-700 border border-rose-200">
                                {wo.failureCode || (isPm ? 'ROUTINE-PM' : 'ELEC-01')}
                              </span>
                            </td>

                            {/* 7. Khách hàng */}
                            <td className="p-3 border-r border-slate-200 text-slate-700 truncate max-w-[160px]" title={wo.customer}>
                              {wo.customer || '---'}
                            </td>

                            {/* 8. Loại CV (PM vs CM) */}
                            <td className="p-3 border-r border-slate-200">
                              {isPm ? (
                                <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-blue-100 text-blue-800 border border-blue-200">
                                  PM {wo.pmCycleDays ? `(${wo.pmCycleDays}d)` : '(Định kỳ)'}
                                </span>
                              ) : (
                                <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-100 text-amber-800 border border-amber-200">
                                  CM (Sự cố WR)
                                </span>
                              )}
                            </td>

                            {/* 9. Ngoài KH */}
                            <td className="p-3 border-r border-slate-200 text-center font-bold">
                              {wo.isUnplanned ? (
                                <span className="px-2 py-0.5 rounded bg-rose-100 text-rose-800 text-[10px]">Có (CM)</span>
                              ) : (
                                <span className="px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 text-[10px]">Không (PM)</span>
                              )}
                            </td>

                            {/* 10. Mức độ ưu tiên (P1, P3, P6, P7, P9) */}
                            <td className="p-3 border-r border-slate-200">
                              <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${priCfg.badgeClass}`} title={priCfg.desc}>
                                {pri} - {priCfg.label}
                              </span>
                            </td>

                            {/* 11. Vòng đời CMMS Stage */}
                            <td className="p-3 border-r border-slate-200">
                              <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold border ${stageCfg.badge}`}>
                                {stageCfg.shortTitle}
                              </span>
                            </td>

                            {/* 12. PIC Người thực hiện */}
                            <td className="p-3 border-r border-slate-200 font-medium text-slate-800">
                              <div className="flex items-center gap-1.5 truncate max-w-[140px]">
                                <UserCheck size={12} className="text-blue-500 shrink-0" />
                                <span>{wo.assignedTo || 'Chờ gán'}</span>
                              </div>
                            </td>

                            {/* 13. Phê duyệt (MM, OM, DM, PM) */}
                            <td className="p-3 border-r border-slate-200 font-mono font-semibold text-slate-700">
                              {wo.dailyMeetingApprovedByRole || wo.responsibleApprove || 'MM'}
                            </td>

                            {/* 14. Thực hiện */}
                            <td className="p-3 border-r border-slate-200 font-mono font-semibold text-slate-700">
                              {wo.responsibleDo || 'Everybody'}
                            </td>

                            {/* 15. Cô lập LOTO */}
                            <td className="p-3 border-r border-slate-200 text-center">
                              {wo.blockingRequired || wo.deEnergyRequired ? (
                                <span className="px-2 py-0.5 rounded bg-amber-100 text-amber-800 font-bold text-[10px] flex items-center justify-center gap-1">
                                  <Lock size={10} /> LOTO
                                </span>
                              ) : (
                                <span className="text-slate-400">Không</span>
                              )}
                            </td>

                            {/* 16. Hạn hoàn thành SLA */}
                            <td className="p-3 border-r border-slate-200 font-medium text-slate-700">
                              {wo.dueDate ? (wo.dueDate.includes('T') ? wo.dueDate.split('T')[0] : wo.dueDate) : '---'}
                            </td>

                            {/* 17. Vật tư */}
                            <td className="p-3 border-r border-slate-200 text-slate-600 truncate max-w-[180px]" title={materialsStr}>
                              {materialsStr}
                            </td>

                            {/* 18. Ngày tạo */}
                            <td className="p-3 border-r border-slate-200 text-slate-500 text-[10px]">
                              {wo.createdAt ? new Date(wo.createdAt).toLocaleDateString('vi-VN') : '---'}
                            </td>

                            {/* 19. Cập nhật */}
                            <td className="p-3 border-r border-slate-200 text-slate-500 text-[10px]">
                              {wo.updatedAt ? new Date(wo.updatedAt).toLocaleDateString('vi-VN') : '---'}
                            </td>

                            {/* 20. Bắt đầu dừng máy */}
                            <td className="p-2 border-r border-slate-200 bg-rose-50/20">
                              <div className="flex items-center justify-between gap-1">
                                <span className="text-[10px] font-mono text-slate-700 truncate max-w-[95px]">
                                  {wo.downtimeStart ? new Date(wo.downtimeStart).toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' }) : '---'}
                                </span>
                                <button
                                  type="button"
                                  onClick={() => handleRecordTimestamp(wo, 'downtimeStart')}
                                  className="px-1.5 py-0.5 bg-rose-100 hover:bg-rose-200 text-rose-800 rounded font-bold text-[10px]"
                                  title="Ghi nhận giờ dừng máy tức thì"
                                >
                                  ⏱️
                                </button>
                              </div>
                            </td>

                            {/* 21. Bắt đầu sửa */}
                            <td className="p-2 border-r border-slate-200 bg-amber-50/20">
                              <div className="flex items-center justify-between gap-1">
                                <span className="text-[10px] font-mono text-slate-700 truncate max-w-[95px]">
                                  {wo.repairStart ? new Date(wo.repairStart).toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' }) : '---'}
                                </span>
                                <button
                                  type="button"
                                  onClick={() => handleRecordTimestamp(wo, 'repairStart')}
                                  className="px-1.5 py-0.5 bg-amber-100 hover:bg-amber-200 text-amber-800 rounded font-bold text-[10px]"
                                  title="Ghi nhận giờ bắt đầu sửa chữa"
                                >
                                  ⏱️
                                </button>
                              </div>
                            </td>

                            {/* 22. Kết thúc sửa */}
                            <td className="p-2 border-r border-slate-200 bg-blue-50/20">
                              <div className="flex items-center justify-between gap-1">
                                <span className="text-[10px] font-mono text-slate-700 truncate max-w-[95px]">
                                  {wo.repairEnd ? new Date(wo.repairEnd).toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' }) : '---'}
                                </span>
                                <button
                                  type="button"
                                  onClick={() => handleRecordTimestamp(wo, 'repairEnd')}
                                  className="px-1.5 py-0.5 bg-blue-100 hover:bg-blue-200 text-blue-800 rounded font-bold text-[10px]"
                                  title="Ghi nhận giờ kết thúc sửa chữa"
                                >
                                  ⏱️
                                </button>
                              </div>
                            </td>

                            {/* 23. Chạy lại máy */}
                            <td className="p-2 border-r border-slate-200 bg-emerald-50/20">
                              <div className="flex items-center justify-between gap-1">
                                <span className="text-[10px] font-mono text-slate-700 truncate max-w-[95px]">
                                  {wo.restartTime ? new Date(wo.restartTime).toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' }) : '---'}
                                </span>
                                <button
                                  type="button"
                                  onClick={() => handleRecordTimestamp(wo, 'restartTime')}
                                  className="px-1.5 py-0.5 bg-emerald-100 hover:bg-emerald-200 text-emerald-800 rounded font-bold text-[10px]"
                                  title="Ghi nhận giờ chạy lại máy trạm"
                                >
                                  ⏱️
                                </button>
                              </div>
                            </td>

                            {/* Actions Column (Contextual by Stage) */}
                            <td className="p-2.5 sticky right-0 z-10 bg-white group-hover:bg-blue-50/40 shadow-xs text-right whitespace-nowrap">
                              <div className="flex items-center justify-end gap-1.5">
                                {/* Action based on stage */}
                                {stage === '1_wr_pending' && (
                                  <button
                                    onClick={() => {
                                      setActiveItem(wo);
                                      setIsDailyApproveModalOpen(true);
                                    }}
                                    className="px-2 py-1 bg-amber-500 hover:bg-amber-600 text-white font-bold rounded text-[11px] shadow-xs"
                                    title="Họp Daily Meeting duyệt WR ➔ WO"
                                  >
                                    Duyệt WR
                                  </button>
                                )}

                                {(stage === '2_daily_approved' || stage === '3_planning') && (
                                  <button
                                    onClick={() => {
                                      setActiveItem(wo);
                                      setIsPlanningModalOpen(true);
                                    }}
                                    className="px-2 py-1 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded text-[11px] shadow-xs"
                                    title="Planning LOTO, vật tư & Cấp Work Permit"
                                  >
                                    Cấp WP
                                  </button>
                                )}

                                {stage === '4_wp_issued' && (
                                  <button
                                    onClick={() => {
                                      setActiveItem(wo);
                                      setIsCloseWpModalOpen(true);
                                    }}
                                    className="px-2 py-1 bg-teal-600 hover:bg-teal-700 text-white font-bold rounded text-[11px] shadow-xs"
                                    title="Nghiệm thu an toàn & Đóng Work Permit tại site"
                                  >
                                    Đóng WP
                                  </button>
                                )}

                                {stage === '5_wp_closed' && (
                                  <button
                                    onClick={() => {
                                      setActiveItem(wo);
                                      setIsCloseWoModalOpen(true);
                                    }}
                                    className="px-2 py-1 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded text-[11px] shadow-xs"
                                    title="Nghiệm thu toàn bộ & Đóng WO trên CMMS"
                                  >
                                    Đóng WO
                                  </button>
                                )}

                                {/* Print Work Permit */}
                                <button
                                  onClick={() => {
                                    setPermitToPrint(wo);
                                    setIsPrintPermitOpen(true);
                                  }}
                                  className="p-1 rounded bg-slate-100 hover:bg-slate-200 text-slate-700"
                                  title="In Giấy phép Work Permit chuẩn A4"
                                >
                                  <Printer size={13} />
                                </button>

                                {/* Trigger NETA Test Form */}
                                {onNavigateToNetaForm && (
                                  <button
                                    onClick={() => {
                                      const eqId = Array.isArray(wo.equipmentId) ? wo.equipmentId[0] : wo.equipmentId;
                                      const targetEq = allEquipment.find(e => e.id.toLowerCase() === (eqId || '').toLowerCase()) || {
                                        id: eqId || 'EQ-01',
                                        name: wo.equipmentName || 'Thiết bị trạm',
                                        type: 'Máy biến áp',
                                        customer: wo.customer,
                                        factory: wo.factory
                                      };
                                      onNavigateToNetaForm(targetEq, isPm ? 'mts' : 'ats');
                                    }}
                                    className="p-1 rounded bg-amber-50 hover:bg-amber-100 text-amber-700"
                                    title="Mở Form Thí Nghiệm NETA (ATS/MTS)"
                                  >
                                    <Zap size={13} />
                                  </button>
                                )}

                                {/* Edit 23 Cols */}
                                <button
                                  onClick={() => {
                                    setActiveItem(wo);
                                    setIsEdit23ColsModalOpen(true);
                                  }}
                                  className="p-1 rounded bg-blue-50 hover:bg-blue-100 text-blue-700"
                                  title="Chỉnh sửa chi tiết"
                                >
                                  <Wrench size={13} />
                                </button>

                                {/* Delete */}
                                {onDeleteWorkOrder && (
                                  <button
                                    onClick={() => {
                                      if (confirm(`Xóa phiếu [${wo.id || wo.wrId}]?`)) {
                                        onDeleteWorkOrder(wo.id);
                                      }
                                    }}
                                    className="p-1 rounded hover:bg-rose-100 text-slate-400 hover:text-rose-600"
                                    title="Xóa phiếu"
                                  >
                                    <Trash2 size={13} />
                                  </button>
                                )}
                              </div>
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* VIEW MODE 2: WORKFLOW 6-STEP PIPELINE KANBAN */}
          {viewMode === 'workflow-kanban' && (
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6 gap-3">
              {Object.entries(WORKFLOW_STAGES).map(([stageKey, cfg]) => {
                const stageOrders = filteredOrders.filter(w => deduceWorkflowStage(w) === stageKey);
                return (
                  <div key={stageKey} className="bg-slate-50 rounded-2xl border border-slate-200 p-3 flex flex-col min-h-[460px]">
                    <div className="flex items-center justify-between pb-2 border-b border-slate-200 mb-2">
                      <span className="font-bold text-xs text-slate-800">{cfg.shortTitle}</span>
                      <span className="px-2 py-0.5 rounded-full bg-white text-slate-700 text-[10px] font-bold shadow-2xs">
                        {stageOrders.length}
                      </span>
                    </div>

                    <div className="space-y-2 overflow-y-auto flex-1">
                      {stageOrders.map(wo => {
                        const pri = deducePriorityCode(wo);
                        const priCfg = PLANT_PRIORITIES[pri];
                        const isPm = !wo.isUnplanned && (wo.type === 'preventive' || wo.workStream === 'PM');

                        return (
                          <div
                            key={wo.id}
                            className="bg-white p-3 rounded-xl border border-slate-200 shadow-2xs hover:shadow-xs transition-all space-y-1.5"
                          >
                            <div className="flex items-center justify-between">
                              <span className="font-mono font-bold text-xs text-emerald-800">
                                {wo.id || wo.wrId}
                              </span>
                              <span className={`px-1.5 py-0.2 rounded text-[10px] font-bold ${priCfg.badgeClass}`}>
                                {pri}
                              </span>
                            </div>

                            <div className="font-bold text-xs text-slate-900 line-clamp-2">
                              {wo.title}
                            </div>

                            <div className="text-[11px] text-slate-500 truncate">
                              {Array.isArray(wo.equipmentId) ? wo.equipmentId.join(', ') : wo.equipmentId}
                            </div>

                            <div className="flex items-center justify-between pt-1 border-t border-slate-100 text-[10px]">
                              <span className="text-slate-600 font-medium">
                                {isPm ? `PM (${wo.pmCycleDays || 30}d)` : 'CM (WR)'}
                              </span>
                              <button
                                onClick={() => {
                                  setActiveItem(wo);
                                  setIsEdit23ColsModalOpen(true);
                                }}
                                className="text-blue-600 font-bold hover:underline"
                              >
                                Chi tiết
                              </button>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* VIEW MODE 3: COMPACT CARDS */}
          {viewMode === 'compact' && (
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
              {filteredOrders.map(wo => {
                const pri = deducePriorityCode(wo);
                const stage = deduceWorkflowStage(wo);
                const priCfg = PLANT_PRIORITIES[pri];
                const stageCfg = WORKFLOW_STAGES[stage];

                return (
                  <div key={wo.id} className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs hover:border-blue-400 transition-all space-y-2.5">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2 font-mono font-bold">
                        <span className="text-sm text-emerald-800">{wo.id || wo.wrId}</span>
                        {wo.workPermitId && (
                          <span className="text-xs px-1.5 py-0.5 rounded bg-blue-50 text-blue-800 border border-blue-200">
                            {wo.workPermitId}
                          </span>
                        )}
                      </div>
                      <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${priCfg.badgeClass}`}>
                        {pri}
                      </span>
                    </div>

                    <h4 className="font-bold text-slate-900 text-sm line-clamp-1">{wo.title}</h4>
                    <p className="text-xs text-slate-500 line-clamp-2">{wo.description}</p>

                    <div className="p-2.5 bg-slate-50 rounded-lg text-xs space-y-1">
                      <div className="flex justify-between">
                        <span className="text-slate-500">Thiết bị:</span>
                        <span className="font-semibold text-slate-800">{Array.isArray(wo.equipmentId) ? wo.equipmentId.join(', ') : wo.equipmentId}</span>
                      </div>
                      <div className="flex justify-between">
                        <span className="text-slate-500">Giai đoạn:</span>
                        <span className="font-bold text-blue-700">{stageCfg.shortTitle}</span>
                      </div>
                      <div className="flex justify-between">
                        <span className="text-slate-500">PIC:</span>
                        <span className="font-medium text-slate-800">{wo.assignedTo || 'Chờ phân công'}</span>
                      </div>
                    </div>

                    <div className="flex items-center justify-between pt-2 border-t border-slate-100 text-xs">
                      <button
                        onClick={() => {
                          setPermitToPrint(wo);
                          setIsPrintPermitOpen(true);
                        }}
                        className="px-2.5 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold flex items-center gap-1"
                      >
                        <Printer size={13} /> In Permit
                      </button>

                      <button
                        onClick={() => {
                          setActiveItem(wo);
                          setIsEdit23ColsModalOpen(true);
                        }}
                        className="px-3 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-700 text-white font-bold"
                      >
                        Sửa Chi Tiết
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}

      {/* ==================================================================== */}
      {/* 6. TAB 2: DAILY MEETING REVIEW BOARD (MM / OM / DM / PM)             */}
      {/* ==================================================================== */}
      {activeSubTab === 'daily-meeting' && (
        <div className="space-y-4">
          <div className="bg-gradient-to-r from-amber-950 via-slate-900 to-slate-900 rounded-2xl p-6 text-white border border-amber-900/50 shadow-md">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div className="space-y-1 max-w-2xl">
                <span className="px-2.5 py-1 rounded-full text-xs font-bold uppercase tracking-wider bg-amber-500/20 text-amber-300 border border-amber-500/30">
                  🤝 Cuộc Họp Giao Ban Đầu Ngày (Daily Operations & Maintenance Review)
                </span>
                <h3 className="text-xl font-extrabold text-white">
                  Xét Duyệt Phiếu Yêu Cầu Công Việc (WR ➔ Work Order)
                </h3>
                <p className="text-xs text-slate-300 leading-relaxed">
                  Các phiếu yêu cầu công việc WR do Field Service / Operator phát hiện tại site được thẩm tra trong Daily Meeting.
                  Khi <strong>Maintenance Manager (MM)</strong>, <strong>Operation Manager (OM)</strong>, <strong>Deputy Manager (DM)</strong> hoặc <strong>Plant Manager (PM)</strong> phê duyệt, WR sẽ chính thức chuyển thành WO và giao Engineer planning an toàn LOTO & vật tư.
                </p>
              </div>

              <div className="p-3 bg-white/10 rounded-xl border border-white/10 text-center shrink-0">
                <div className="text-2xl font-black text-amber-400">{plantKpis.wrPendingCount}</div>
                <div className="text-xs text-slate-300">Phiếu WR chờ duyệt hôm nay</div>
              </div>
            </div>
          </div>

          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
            <div className="p-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
              <h4 className="font-bold text-slate-900 text-sm">
                Danh Sách Phiếu Yêu Cầu WR Cần Họp Phê Duyệt ({workOrders.filter(w => deduceWorkflowStage(w) === '1_wr_pending').length})
              </h4>
              <button
                onClick={handleOpenCreateWr}
                className="px-3 py-1.5 bg-rose-600 hover:bg-rose-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5"
              >
                <Plus size={14} /> Mở Thêm WR
              </button>
            </div>

            <div className="divide-y divide-slate-100">
              {workOrders.filter(w => deduceWorkflowStage(w) === '1_wr_pending').length === 0 ? (
                <div className="p-10 text-center text-slate-400">
                  Hiện tại không có phiếu yêu cầu WR nào đang chờ họp Daily Meeting duyệt. Mọi việc đều được thông qua!
                </div>
              ) : (
                workOrders.filter(w => deduceWorkflowStage(w) === '1_wr_pending').map(wo => {
                  const pri = deducePriorityCode(wo);
                  const priCfg = PLANT_PRIORITIES[pri];

                  return (
                    <div key={wo.id} className="p-4 hover:bg-amber-50/30 transition-colors flex flex-col md:flex-row md:items-center justify-between gap-4">
                      <div className="space-y-1.5 max-w-2xl">
                        <div className="flex items-center gap-2.5">
                          <span className="font-mono font-bold text-xs text-amber-800 bg-amber-100 px-2 py-0.5 rounded">
                            {wo.wrId || wo.id}
                          </span>
                          <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${priCfg.badgeClass}`}>
                            {pri} - {priCfg.label}
                          </span>
                          <span className="text-xs text-slate-500">
                            Người báo cáo: <strong>{wo.reportedBy || 'Operator'}</strong> ({wo.reportedAt ? new Date(wo.reportedAt).toLocaleDateString('vi-VN') : 'Hôm nay'})
                          </span>
                        </div>

                        <div className="font-bold text-slate-900 text-sm">{wo.title}</div>
                        <div className="text-xs text-slate-600">{wo.description}</div>

                        <div className="flex flex-wrap items-center gap-3 text-[11px] text-slate-500 pt-1">
                          <span>Thiết bị: <strong className="text-slate-800">{Array.isArray(wo.equipmentId) ? wo.equipmentId.join(', ') : wo.equipmentId}</strong></span>
                          <span>•</span>
                          <span>Trạm / Nhà máy: {wo.factory || wo.customer}</span>
                          <span>•</span>
                          <span>Failure Code: <span className="font-mono text-rose-700 font-bold">{wo.failureCode}</span></span>
                        </div>
                      </div>

                      <div className="flex items-center gap-2 shrink-0">
                        <button
                          onClick={() => {
                            setActiveItem(wo);
                            setIsDailyApproveModalOpen(true);
                          }}
                          className="px-4 py-2 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 text-white font-bold text-xs rounded-xl shadow-xs flex items-center gap-1.5"
                        >
                          <Check size={14} /> Phê Duyệt WR ➔ Chuyển WO
                        </button>
                      </div>
                    </div>
                  );
                })
              )}
            </div>
          </div>
        </div>
      )}

      {/* ==================================================================== */}
      {/* 7. TAB 3: PM RECURRING CYCLE MATRIX (07d - 360d)                     */}
      {/* ==================================================================== */}
      {activeSubTab === 'pm-matrix' && (
        <div className="space-y-4">
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-xs">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-4 border-b border-slate-100">
              <div>
                <h3 className="font-bold text-slate-900 text-base">
                  Tiêu Chuẩn Chu Kỳ Bảo Trì Định Kỳ PM (07d, 14d, 30d, 60d, 90d, 180d, 360d)
                </h3>
                <p className="text-xs text-slate-500">
                  Tất cả các công việc bảo trì định kỳ chuẩn đều mang mã <strong>Priority 6 (Routine PM)</strong>
                </p>
              </div>

              <button
                onClick={handleOpenCreatePm}
                className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs rounded-xl flex items-center gap-1.5 shadow-xs"
              >
                <Plus size={15} /> Thiết Lập Lịch PM Cho Thiết Bị
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 mt-4">
              {PM_CYCLE_OPTIONS.map(c => {
                const count = workOrders.filter(w => !w.isUnplanned && (w.pmCycleDays === c.days)).length;
                return (
                  <div key={c.days} className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/60 space-y-2">
                    <div className="flex items-center justify-between">
                      <span className="font-bold text-xs text-blue-700">{c.label}</span>
                      <span className="px-2 py-0.5 rounded-full bg-blue-100 text-blue-800 text-[10px] font-bold">
                        {count} thiết bị
                      </span>
                    </div>
                    <p className="text-[11px] text-slate-600 leading-relaxed">
                      {c.desc}
                    </p>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Equipment Schedule List */}
          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
            <div className="p-4 border-b border-slate-200 bg-slate-50 flex items-center justify-between">
              <span className="font-bold text-slate-900 text-sm">Danh Mục Thiết Bị & Chu Kỳ Bảo Dưỡng PM</span>
              <span className="text-xs text-slate-500">{allEquipment.length} thiết bị</span>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead className="bg-slate-100 text-slate-700 uppercase font-bold text-[11px] border-b border-slate-200">
                  <tr>
                    <th className="p-3">Mã Thiết Bị KKS</th>
                    <th className="p-3">Tên Thiết Bị & Loại</th>
                    <th className="p-3">Khách Hàng / Trạm</th>
                    <th className="p-3">Chu Kỳ PM Tiêu Chuẩn</th>
                    <th className="p-3">Hạn Kế Tiếp</th>
                    <th className="p-3">Sức Khỏe</th>
                    <th className="p-3 text-right">Thao Tác</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {allEquipment.map(eq => (
                    <tr key={eq.id} className="hover:bg-slate-50 transition-colors">
                      <td className="p-3 font-mono font-bold text-blue-700">{eq.id}</td>
                      <td className="p-3">
                        <div className="font-bold text-slate-900">{eq.name}</div>
                        <div className="text-[11px] text-slate-500">{eq.type || 'Thiết bị điện'}</div>
                      </td>
                      <td className="p-3 text-slate-700">
                        <div>{eq.customer || '---'}</div>
                        <div className="text-[11px] text-slate-500">{eq.factory || eq.location}</div>
                      </td>
                      <td className="p-3">
                        <span className="px-2 py-0.5 rounded bg-blue-50 text-blue-700 font-bold border border-blue-200">
                          {eq.pmCycleMonths ? `${eq.pmCycleMonths * 30} ngày` : '30 ngày'}
                        </span>
                      </td>
                      <td className="p-3 font-medium text-slate-700">
                        {eq.lastCheck ? `${eq.lastCheck}` : 'Đến hạn'}
                      </td>
                      <td className="p-3">
                        <span className="px-2 py-0.5 rounded-full font-bold text-[10px] bg-emerald-100 text-emerald-800">
                          {eq.health ?? 92}%
                        </span>
                      </td>
                      <td className="p-3 text-right">
                        <button
                          onClick={() => {
                            const nextWo = getNextWorkOrderId(workOrders);
                            const nextWp = getNextWorkPermitId(workOrders);
                            setActiveItem({
                              id: nextWo,
                              workPermitId: nextWp,
                              title: `Bảo dưỡng định kỳ 30d - ${eq.name}`,
                              description: `Thực hiện bảo dưỡng định kỳ và thử nghiệm NETA cho ${eq.name}`,
                              equipmentId: eq.id,
                              equipmentName: eq.name,
                              customer: eq.customer || '',
                              factory: eq.factory || eq.location || '',
                              workStream: 'PM',
                              pmCycleDays: 30,
                              isUnplanned: false,
                              priorityCode: 'P6',
                              priority: 'medium',
                              workflowStage: '2_daily_approved',
                              status: 'approved',
                              assignedTo: 'Võ Minh Tú',
                              blockingRequired: true,
                              deEnergyRequired: true,
                              dueDate: new Date(Date.now() + 30 * 86400000).toISOString().split('T')[0]
                            });
                            setIsPmModalOpen(true);
                          }}
                          className="px-2.5 py-1.5 bg-blue-50 hover:bg-blue-100 text-blue-700 font-bold rounded-lg transition-colors"
                        >
                          Lập Phiếu PM
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================================== */}
      {/* 8. TAB 4: WORK PERMITS & LOTO DE-ENERGIZATION SAFETY MATRIX          */}
      {/* ==================================================================== */}
      {activeSubTab === 'wp-loto' && (
        <div className="space-y-4">
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-xs space-y-3">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div>
                <h3 className="font-bold text-slate-900 text-base flex items-center gap-2">
                  <Lock size={18} className="text-indigo-600" />
                  Quản Lý Giấy Phép Work Permit & An Toàn Cô Lập Nguồn (LOTO)
                </h3>
                <p className="text-xs text-slate-500">
                  Kiểm soát an toàn trước khi vào trạm: Cắt điện (De-energization), khóa LOTO, thử không còn áp và nối đất lưu động.
                </p>
              </div>

              <div className="flex items-center gap-2">
                <span className="px-3 py-1 bg-amber-50 text-amber-800 border border-amber-200 rounded-lg text-xs font-bold">
                  {workOrders.filter(w => w.blockingRequired || w.deEnergyRequired).length} Phiếu Yêu Cầu LOTO
                </span>
              </div>
            </div>
          </div>

          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead className="bg-slate-100 text-slate-700 uppercase font-bold text-[11px] border-b border-slate-200">
                  <tr>
                    <th className="p-3">Mã WP</th>
                    <th className="p-3">Mã WO Liên Quan</th>
                    <th className="p-3">Tiêu Đề Công Việc</th>
                    <th className="p-3">Thiết Bị</th>
                    <th className="p-3">Trạng Thái Giấy Phép</th>
                    <th className="p-3">Điểm Cô Lập / Thẻ LOTO</th>
                    <th className="p-3">Người Cấp / Đóng</th>
                    <th className="p-3 text-right">Thao Tác Giấy Phép</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {workOrders.filter(w => w.blockingRequired || w.workPermitId).map(wo => {
                    const stage = deduceWorkflowStage(wo);
                    return (
                      <tr key={wo.id} className="hover:bg-slate-50 transition-colors">
                        <td className="p-3 font-mono font-bold text-blue-700">
                          {wo.workPermitId || 'Chưa cấp'}
                        </td>
                        <td className="p-3 font-mono font-bold text-emerald-700">
                          {wo.id}
                        </td>
                        <td className="p-3 font-semibold text-slate-900 max-w-[220px] truncate" title={wo.title}>
                          {wo.title}
                        </td>
                        <td className="p-3 text-slate-700">
                          {Array.isArray(wo.equipmentId) ? wo.equipmentId.join(', ') : wo.equipmentId}
                        </td>
                        <td className="p-3">
                          {stage === '4_wp_issued' ? (
                            <span className="px-2 py-0.5 rounded-full font-bold text-[10px] bg-purple-100 text-purple-800 border border-purple-200">
                              Đang Mở (Hiệu Lực Tại Site)
                            </span>
                          ) : stage === '5_wp_closed' || stage === '6_wo_closed' ? (
                            <span className="px-2 py-0.5 rounded-full font-bold text-[10px] bg-emerald-100 text-emerald-800 border border-emerald-200">
                              Đã Đóng An Toàn
                            </span>
                          ) : (
                            <span className="px-2 py-0.5 rounded-full font-bold text-[10px] bg-amber-100 text-amber-800 border border-amber-200">
                              Đang Chuẩn Bị Planning
                            </span>
                          )}
                        </td>
                        <td className="p-3 text-slate-600">
                          {wo.isolationPoints || wo.lotoTagNumbers || 'Cắt MC lộ tổng & khóa LOTO'}
                        </td>
                        <td className="p-3 text-slate-600">
                          {wo.wpIssuedBy || wo.assignedTo || '---'}
                        </td>
                        <td className="p-3 text-right">
                          <div className="flex items-center justify-end gap-1.5">
                            <button
                              onClick={() => {
                                setPermitToPrint(wo);
                                setIsPrintPermitOpen(true);
                              }}
                              className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-semibold flex items-center gap-1"
                            >
                              <Printer size={13} /> In WP
                            </button>
                            {stage === '4_wp_issued' && (
                              <button
                                onClick={() => {
                                  setActiveItem(wo);
                                  setIsCloseWpModalOpen(true);
                                }}
                                className="px-2.5 py-1 bg-teal-600 hover:bg-teal-700 text-white rounded-lg text-xs font-bold"
                              >
                                Đóng WP
                              </button>
                            )}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================================== */}
      {/* 9. TAB 5: PM SCHEDULE DASHBOARD INTEGRATION                          */}
      {/* ==================================================================== */}
      {activeSubTab === 'pm-dashboard' && (
        <PmScheduleDashboard
          workOrders={workOrders}
          allEquipment={allEquipment}
          onOpenWorkOrderModal={(wo) => {
            setActiveItem(wo);
            setIsEdit23ColsModalOpen(true);
          }}
          onSendAlert={() => onSendPMAlerts?.(true)}
          onOpenEmailConfig={onOpenEmailConfig}
          onUpdateWorkOrderStatus={async (id, status) => {
            const wo = workOrders.find(w => w.id === id);
            if (wo) {
              await onSaveWorkOrder({ ...wo, status, updatedAt: new Date().toISOString() }, true);
            }
          }}
        />
      )}

      {/* ==================================================================== */}
      {/* 10. TAB 6: AUTOMATED EMAIL ALERTS (3 DAYS NOTICE)                    */}
      {/* ==================================================================== */}
      {activeSubTab === 'alerts' && (
        <div className="space-y-4">
          <div className="bg-gradient-to-br from-slate-900 via-slate-800 to-indigo-950 rounded-2xl p-6 text-white border border-slate-700 shadow-md">
            <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6">
              <div className="space-y-2 max-w-2xl">
                <span className="px-3 py-1 rounded-full text-xs font-bold uppercase tracking-wider bg-blue-500/20 text-blue-300 border border-blue-500/30">
                  ⚡ Firebase Cloud Functions & Nodemailer
                </span>
                <h3 className="text-xl font-bold text-white">
                  Cảnh Báo Tự Động Trước 3 Ngày Đến Hạn Phiếu PM
                </h3>
                <p className="text-sm text-slate-300 leading-relaxed">
                  Hệ thống tự động quét mỗi ngày lúc 08:00 AM, kiểm tra các phiếu bảo trì định kỳ sắp đến hạn trong vòng 3 ngày
                  hoặc đã quá hạn để gửi email thông báo chi tiết đến kỹ sư hiện trường và quản trị viên ({userEmail}).
                </p>
              </div>

              <div className="flex items-center gap-3 shrink-0">
                {onSendPMAlerts && (
                  <button
                    onClick={() => onSendPMAlerts(true)}
                    className="flex items-center gap-2 px-4 py-2.5 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white font-bold text-sm rounded-xl shadow-md transition-all"
                  >
                    <Mail size={16} />
                    <span>Gửi Thử Nghiệm Ngay</span>
                  </button>
                )}
                {onOpenEmailConfig && (
                  <button
                    onClick={onOpenEmailConfig}
                    className="flex items-center gap-2 px-4 py-2.5 bg-slate-800 hover:bg-slate-700 border border-slate-600 text-slate-200 font-bold text-sm rounded-xl transition-colors"
                  >
                    <SlidersHorizontal size={16} />
                    <span>Cấu Hình Email</span>
                  </button>
                )}
              </div>
            </div>
          </div>

          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 space-y-4">
            <h4 className="font-bold text-slate-900 text-base">
              Danh Sách Nhiệm Vụ PM Cần Cảnh Báo ({pmUpcomingTasks.length})
            </h4>

            {pmUpcomingTasks.length === 0 ? (
              <div className="p-8 text-center text-slate-400 bg-slate-50 rounded-xl border border-slate-100">
                Hiện tại không có phiếu PM nào sắp đến hạn trong 3 ngày tới hoặc bị quá hạn. Tiến độ hoàn hảo!
              </div>
            ) : (
              <div className="divide-y divide-slate-100">
                {pmUpcomingTasks.map((task, idx) => (
                  <div key={idx} className="py-3 flex items-center justify-between gap-4">
                    <div>
                      <div className="font-bold text-slate-900 text-sm">{task.title || task.equipmentName}</div>
                      <div className="text-xs text-slate-500">Mã WO: <span className="font-mono font-bold text-blue-700">{task.id}</span> • Thiết bị: {task.equipmentId}</div>
                    </div>
                    <div className="flex items-center gap-3">
                      <span className="px-2.5 py-1 rounded bg-amber-50 text-amber-700 border border-amber-200 font-bold text-xs">
                        Hạn: {task.dueDate}
                      </span>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}

      {/* ==================================================================== */}
      {/* 11. MODAL 1: MỞ PHIẾU YÊU CẦU WR MỚI (OPERATOR / FIELD SERVICE SITE) */}
      {/* ==================================================================== */}
      {isWrModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex items-center justify-center p-3 sm:p-4 overflow-y-auto">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden my-6 border border-slate-200">
            <div className="px-6 py-4 bg-gradient-to-r from-rose-900 via-rose-800 to-slate-900 text-white flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-xl bg-rose-600 flex items-center justify-center font-bold">
                  <AlertTriangle size={18} />
                </div>
                <div>
                  <h3 className="font-bold text-base">Mở Phiếu Yêu Cầu Công Việc (WR - Work Request)</h3>
                  <p className="text-xs text-rose-200">Dành cho Field Service / Operator ghi nhận khiếm khuyết thiết bị ngoài site</p>
                </div>
              </div>
              <button onClick={() => setIsWrModalOpen(false)} className="text-slate-300 hover:text-white">
                <X size={20} />
              </button>
            </div>

            <div className="p-6 space-y-4 text-xs">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Mã Yêu Cầu (WR ID)</label>
                  <input
                    type="text"
                    disabled
                    value={activeItem.wrId || ''}
                    className="w-full px-3 py-2 bg-slate-100 border border-slate-300 rounded-lg font-mono font-bold text-rose-700"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Người phát hiện (Reported By)</label>
                  <input
                    type="text"
                    value={activeItem.reportedBy || ''}
                    onChange={(e) => setActiveItem({ ...activeItem, reportedBy: e.target.value })}
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg font-medium"
                    placeholder="VD: Nguyễn Văn Tuấn (Kỹ sư FSE)"
                  />
                </div>
              </div>

              {/* Priority Standard: P1, P3, P7, P9 */}
              <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
                <label className="block text-[11px] font-bold text-slate-800">
                  Mức Độ Ưu Tiên Chuẩn Nhà Máy (Priority Selection) *
                </label>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  {(['P1', 'P3', 'P7', 'P9'] as PriorityCode[]).map(code => {
                    const cfg = PLANT_PRIORITIES[code];
                    const isSelected = activeItem.priorityCode === code;
                    return (
                      <button
                        type="button"
                        key={code}
                        onClick={() => setActiveItem({ ...activeItem, priorityCode: code })}
                        className={`p-2.5 rounded-xl border text-left transition-all ${
                          isSelected
                            ? `${cfg.bgLight} border-2 ring-1 shadow-2xs`
                            : 'bg-white border-slate-200 hover:border-slate-300'
                        }`}
                      >
                        <div className="flex items-center justify-between">
                          <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${cfg.badgeClass}`}>
                            {cfg.label}
                          </span>
                          {isSelected && <Check size={14} className={cfg.textClass} />}
                        </div>
                        <div className="font-bold text-xs text-slate-900 mt-1">{cfg.title}</div>
                        <div className="text-[10px] text-slate-500 mt-0.5 line-clamp-2">{cfg.desc}</div>
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* Equipment & Customer */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Thiết bị xảy ra sự cố *</label>
                  <select
                    value={typeof activeItem.equipmentId === 'string' ? activeItem.equipmentId : (activeItem.equipmentId?.[0] || '')}
                    onChange={(e) => {
                      const sel = allEquipment.find(eq => eq.id === e.target.value);
                      setActiveItem({
                        ...activeItem,
                        equipmentId: e.target.value,
                        equipmentName: sel?.name || ''
                      });
                    }}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg font-semibold"
                  >
                    <option value="">-- Chọn thiết bị --</option>
                    {allEquipment.map(eq => (
                      <option key={eq.id} value={eq.id}>{eq.id} - {eq.name}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Mã lỗi sự cố (Failure Code)</label>
                  <select
                    value={activeItem.failureCode || 'ELEC-01'}
                    onChange={(e) => setActiveItem({ ...activeItem, failureCode: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg font-semibold text-rose-700"
                  >
                    {FAILURE_CODES.map(fc => (
                      <option key={fc.code} value={fc.code}>{fc.code} - {fc.desc}</option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Title & Description */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">Tiêu đề khiếm khuyết phát hiện *</label>
                <input
                  type="text"
                  value={activeItem.title || ''}
                  onChange={(e) => setActiveItem({ ...activeItem, title: e.target.value })}
                  placeholder="VD: Quá nhiệt đầu cốt 24kV pha B máy biến áp T1 / Phóng điện rò rỉ"
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg font-semibold text-sm focus:ring-2 focus:ring-rose-500"
                />
              </div>

              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">Mô tả hiện tượng chi tiết tại site</label>
                <textarea
                  rows={3}
                  value={activeItem.description || ''}
                  onChange={(e) => setActiveItem({ ...activeItem, description: e.target.value })}
                  placeholder="Mô tả các dấu hiệu, mùi khét, tiếng kêu bất thường, số đo nhiệt độ đo bằng súng hồng ngoại..."
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs"
                />
              </div>

              {/* Safety: LOTO & De-energize */}
              <div className="p-3 rounded-xl bg-amber-50 border border-amber-200 flex items-center justify-between">
                <div>
                  <span className="font-bold text-slate-900 text-xs block">Yêu cầu cô lập điện & khóa LOTO</span>
                  <span className="text-[11px] text-slate-600">Sự cố có bắt buộc phải cắt nguồn và treo thẻ khóa an toàn không?</span>
                </div>
                <input
                  type="checkbox"
                  checked={Boolean(activeItem.blockingRequired)}
                  onChange={(e) => setActiveItem({ ...activeItem, blockingRequired: e.target.checked, deEnergyRequired: e.target.checked })}
                  className="w-4 h-4 text-amber-600 rounded"
                />
              </div>
            </div>

            <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between">
              <button
                type="button"
                onClick={() => setIsWrModalOpen(false)}
                className="px-4 py-2 border border-slate-300 rounded-xl text-xs font-semibold text-slate-700 hover:bg-slate-100"
              >
                Hủy bỏ
              </button>

              <button
                type="button"
                onClick={handleSaveWr}
                className="px-5 py-2.5 bg-gradient-to-r from-rose-600 to-amber-600 hover:from-rose-700 hover:to-amber-700 text-white font-bold text-xs rounded-xl shadow-md"
              >
                Gửi Phiếu Yêu Cầu WR Lên Daily Meeting
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================================== */}
      {/* 12. MODAL 2: TẠO KẾ HOẠCH PM ĐỊNH KỲ (07d - 360d)                     */}
      {/* ==================================================================== */}
      {isPmModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex items-center justify-center p-3 sm:p-4 overflow-y-auto">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden my-6 border border-slate-200">
            <div className="px-6 py-4 bg-gradient-to-r from-blue-900 via-indigo-950 to-slate-900 text-white flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-xl bg-blue-600 flex items-center justify-center font-bold">
                  <Calendar size={18} />
                </div>
                <div>
                  <h3 className="font-bold text-base">Thiết Lập Kế Hoạch Bảo Trì Định Kỳ (PM)</h3>
                  <p className="text-xs text-blue-200">Chu kỳ tiêu chuẩn: 07d, 14d, 30d, 60d, 90d, 180d, 360d (Priority 6)</p>
                </div>
              </div>
              <button onClick={() => setIsPmModalOpen(false)} className="text-slate-300 hover:text-white">
                <X size={20} />
              </button>
            </div>

            <div className="p-6 space-y-4 text-xs">
              {/* Select PM Cycle Days */}
              <div>
                <label className="block text-[11px] font-bold text-slate-800 mb-1.5">
                  Chọn Chu Kỳ Ngày (PM Frequency Cycle) *
                </label>
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                  {PM_CYCLE_OPTIONS.map(opt => (
                    <button
                      type="button"
                      key={opt.days}
                      onClick={() => setActiveItem({
                        ...activeItem,
                        pmCycleDays: opt.days,
                        dueDate: new Date(Date.now() + opt.days * 86400000).toISOString().split('T')[0]
                      })}
                      className={`p-2.5 rounded-xl border text-center transition-all ${
                        activeItem.pmCycleDays === opt.days
                          ? 'border-blue-600 bg-blue-50 text-blue-900 font-bold shadow-xs'
                          : 'border-slate-200 bg-white hover:bg-slate-50 text-slate-700'
                      }`}
                    >
                      <div className="text-xs font-bold">{opt.days} Ngày</div>
                      <div className="text-[10px] text-slate-500 mt-0.5 truncate">{opt.label.split('(')[1]?.replace(')', '') || ''}</div>
                    </button>
                  ))}
                </div>
              </div>

              {/* Equipment */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">Thiết bị bảo trì *</label>
                <select
                  value={typeof activeItem.equipmentId === 'string' ? activeItem.equipmentId : (activeItem.equipmentId?.[0] || '')}
                  onChange={(e) => {
                    const sel = allEquipment.find(eq => eq.id === e.target.value);
                    setActiveItem({
                      ...activeItem,
                      equipmentId: e.target.value,
                      equipmentName: sel?.name || '',
                      title: `Bảo dưỡng định kỳ ${activeItem.pmCycleDays || 30}d - ${sel?.name || 'Thiết bị trạm'}`
                    });
                  }}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg font-semibold"
                >
                  <option value="">-- Chọn thiết bị --</option>
                  {allEquipment.map(eq => (
                    <option key={eq.id} value={eq.id}>{eq.id} - {eq.name}</option>
                  ))}
                </select>
              </div>

              {/* Title & Scope */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">Tiêu đề công việc PM *</label>
                <input
                  type="text"
                  value={activeItem.title || ''}
                  onChange={(e) => setActiveItem({ ...activeItem, title: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg font-semibold text-sm"
                />
              </div>

              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">Mô tả các hạng mục kiểm tra định kỳ</label>
                <textarea
                  rows={3}
                  value={activeItem.description || ''}
                  onChange={(e) => setActiveItem({ ...activeItem, description: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs"
                />
              </div>

              {/* Engineer & Due Date */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <FseEngineerDropdown
                    label="Kỹ sư phụ trách (PIC)"
                    value={activeItem.assignedTo || ''}
                    onChange={(val) => setActiveItem({ ...activeItem, assignedTo: val })}
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Hạn hoàn thành (Due Date)</label>
                  <input
                    type="date"
                    value={activeItem.dueDate || ''}
                    onChange={(e) => setActiveItem({ ...activeItem, dueDate: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg"
                  />
                </div>
              </div>
            </div>

            <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between">
              <button
                type="button"
                onClick={() => setIsPmModalOpen(false)}
                className="px-4 py-2 border border-slate-300 rounded-xl text-xs font-semibold text-slate-700 hover:bg-slate-100"
              >
                Hủy bỏ
              </button>

              <button
                type="button"
                onClick={handleSavePm}
                className="px-5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs rounded-xl shadow-md"
              >
                Lưu & Thiết Lập Kế Hoạch PM
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================================== */}
      {/* 13. MODAL 3: DAILY MEETING APPROVAL (WR ➔ WORK ORDER)                */}
      {/* ==================================================================== */}
      {isDailyApproveModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex items-center justify-center p-3 sm:p-4 overflow-y-auto">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden my-6 border border-slate-200">
            <div className="px-6 py-4 bg-gradient-to-r from-emerald-900 via-teal-900 to-slate-900 text-white flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-xl bg-emerald-600 flex items-center justify-center font-bold">
                  <Check size={18} />
                </div>
                <div>
                  <h3 className="font-bold text-base">Họp Daily Meeting: Phê Duyệt WR ➔ Chuyển Thành WO</h3>
                  <p className="text-xs text-emerald-200">Ủy quyền phê duyệt: Maintenance Manager / Operation Manager / Plant Manager</p>
                </div>
              </div>
              <button onClick={() => setIsDailyApproveModalOpen(false)} className="text-slate-300 hover:text-white">
                <X size={20} />
              </button>
            </div>

            <div className="p-6 space-y-4 text-xs">
              {/* WR Info Recap */}
              <div className="p-3 bg-amber-50 rounded-xl border border-amber-200 space-y-1">
                <div className="flex items-center justify-between">
                  <span className="font-mono font-bold text-amber-900">Mã WR: {activeItem.wrId || activeItem.id}</span>
                  <span className="font-bold text-slate-700">Mã WO sẽ cấp: <strong className="font-mono text-emerald-700">{activeItem.id || 'WO-TỰ-ĐỘNG'}</strong></span>
                </div>
                <div className="font-bold text-sm text-slate-900">{activeItem.title}</div>
                <div className="text-slate-600 text-[11px]">{activeItem.description}</div>
              </div>

              {/* Approver Role Selection */}
              <div>
                <label className="block text-[11px] font-bold text-slate-800 mb-1.5">
                  Vai Trò Cán Bộ Phê Duyệt Trong Buổi Họp *
                </label>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  {APPROVER_ROLES.map(r => (
                    <button
                      type="button"
                      key={r.role}
                      onClick={() => setActiveItem({ ...activeItem, dailyMeetingApprovedByRole: r.role as any })}
                      className={`p-2.5 rounded-xl border text-left transition-all ${
                        activeItem.dailyMeetingApprovedByRole === r.role
                          ? 'border-emerald-600 bg-emerald-50 text-emerald-950 font-bold shadow-2xs'
                          : 'border-slate-200 bg-white hover:bg-slate-50 text-slate-700'
                      }`}
                    >
                      <div className="font-bold">{r.role}</div>
                      <div className="text-[10px] text-slate-500">{r.title}</div>
                    </button>
                  ))}
                </div>
              </div>

              {/* Assign Lead Engineer */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <FseEngineerDropdown
                    label="Giao cho Kỹ Sư Phụ Trách Planning (Lead Engineer) *"
                    value={activeItem.assignedTo || ''}
                    onChange={(val) => setActiveItem({ ...activeItem, assignedTo: val })}
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Thời hạn cam kết hoàn thành (SLA Due Date)</label>
                  <input
                    type="date"
                    value={activeItem.dueDate ? activeItem.dueDate.split('T')[0] : ''}
                    onChange={(e) => setActiveItem({ ...activeItem, dueDate: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg font-medium"
                  />
                </div>
              </div>

              {/* Daily Meeting Directives / Notes */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">
                  Ý kiến chỉ đạo từ cuộc họp Daily Meeting
                </label>
                <textarea
                  rows={2}
                  value={activeItem.dailyMeetingNotes || ''}
                  onChange={(e) => setActiveItem({ ...activeItem, dailyMeetingNotes: e.target.value })}
                  placeholder="VD: Cần phối hợp với ca vận hành cắt điện lúc 13h00; kiểm tra phụ tùng roong phớt sẵn có trong kho..."
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs"
                />
              </div>
            </div>

            <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between">
              <button
                type="button"
                onClick={() => setIsDailyApproveModalOpen(false)}
                className="px-4 py-2 border border-slate-300 rounded-xl text-xs font-semibold text-slate-700 hover:bg-slate-100"
              >
                Hủy bỏ
              </button>

              <button
                type="button"
                onClick={handleApproveWrInDailyMeeting}
                className="px-5 py-2.5 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 text-white font-bold text-xs rounded-xl shadow-md flex items-center gap-1.5"
              >
                <Check size={15} /> Xác Nhận Phê Duyệt ➔ Chuyển Thành WO
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================================== */}
      {/* 14. MODAL 4: ENGINEER PLANNING (LOTO / DE-ENERGY / SPARE PARTS / WP)  */}
      {/* ==================================================================== */}
      {isPlanningModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex items-center justify-center p-3 sm:p-4 overflow-y-auto">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden my-6 border border-slate-200">
            <div className="px-6 py-4 bg-gradient-to-r from-indigo-900 via-slate-900 to-slate-900 text-white flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-xl bg-indigo-600 flex items-center justify-center font-bold">
                  <Wrench size={18} />
                </div>
                <div>
                  <h3 className="font-bold text-base">Engineer Planning & Cấp Giấy Phép Work Permit (WP)</h3>
                  <p className="text-xs text-indigo-200">Lập kế hoạch an toàn LOTO, cô lập nguồn, vật tư phụ tùng và phát hành WP</p>
                </div>
              </div>
              <button onClick={() => setIsPlanningModalOpen(false)} className="text-slate-300 hover:text-white">
                <X size={20} />
              </button>
            </div>

            <div className="p-6 space-y-4 text-xs">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Mã Lệnh Công Việc (WO)</label>
                  <input
                    type="text"
                    disabled
                    value={activeItem.id || ''}
                    className="w-full px-3 py-2 bg-slate-100 border border-slate-300 rounded-lg font-mono font-bold text-emerald-800"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Mã Giấy Phép Work Permit (WP)</label>
                  <input
                    type="text"
                    value={activeItem.workPermitId || getNextWorkPermitId(workOrders)}
                    onChange={(e) => setActiveItem({ ...activeItem, workPermitId: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg font-mono font-bold text-blue-800"
                  />
                </div>
              </div>

              {/* LOTO & De-energize Details */}
              <div className="p-3.5 rounded-xl bg-amber-50/70 border border-amber-200 space-y-3">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-amber-900 text-xs flex items-center gap-1.5">
                    <Lock size={15} /> Kế Hoạch An Toàn Cô Lập Điện & Khóa LOTO Bắt Buộc
                  </span>
                  <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-amber-200 text-amber-900">
                    Bắt Buộc
                  </span>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                  <div>
                    <label className="block text-[10px] font-bold text-slate-700 mb-1">Mã thẻ khóa LOTO (Lockout Tag #)</label>
                    <input
                      type="text"
                      value={activeItem.lotoTagNumbers || ''}
                      onChange={(e) => setActiveItem({ ...activeItem, lotoTagNumbers: e.target.value })}
                      placeholder="VD: LOTO-MBA-01, PADLOCK-92"
                      className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg font-mono"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold text-slate-700 mb-1">Các điểm cô lập (Isolation Points)</label>
                    <input
                      type="text"
                      value={activeItem.isolationPoints || ''}
                      onChange={(e) => setActiveItem({ ...activeItem, isolationPoints: e.target.value })}
                      placeholder="VD: Cắt MC tổng, kéo dao tiếp địa DS, khóa van"
                      className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg font-mono"
                    />
                  </div>
                </div>
              </div>

              {/* Spare parts & Materials */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">Vật tư phụ tùng cần chuẩn bị (Spare Parts Planning)</label>
                <textarea
                  rows={2}
                  value={typeof activeItem.usedMaterials === 'string' ? activeItem.usedMaterials : JSON.stringify(activeItem.usedMaterials || '')}
                  onChange={(e) => setActiveItem({ ...activeItem, usedMaterials: e.target.value })}
                  placeholder="Ghi rõ danh mục: Roong phớt, mỡ bôi trơn, băng dính cách điện 3M, dầu biến áp..."
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg"
                />
              </div>

              {/* Permit Issuer */}
              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">Cán bộ phát hành giấy phép (Permit Issuer / Safety Officer)</label>
                <input
                  type="text"
                  value={activeItem.wpIssuedBy || 'KS. Nguyễn Văn An (Safety Officer)'}
                  onChange={(e) => setActiveItem({ ...activeItem, wpIssuedBy: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg font-medium"
                />
              </div>
            </div>

            <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between">
              <button
                type="button"
                onClick={() => setIsPlanningModalOpen(false)}
                className="px-4 py-2 border border-slate-300 rounded-xl text-xs font-semibold text-slate-700 hover:bg-slate-100"
              >
                Hủy bỏ
              </button>

              <button
                type="button"
                onClick={handleSavePlanningAndIssueWp}
                className="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs rounded-xl shadow-md flex items-center gap-1.5"
              >
                <Shield size={14} /> Hoàn Tất Planning & Phát Hành Work Permit
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================================== */}
      {/* 15. MODAL 5: CLOSE WORK PERMIT (NGHIỆM THU HIỆN TRƯỜNG & RÚT LOTO)   */}
      {/* ==================================================================== */}
      {isCloseWpModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex items-center justify-center p-3 sm:p-4 overflow-y-auto">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-xl overflow-hidden my-6 border border-slate-200">
            <div className="px-6 py-4 bg-gradient-to-r from-teal-900 via-slate-900 to-slate-900 text-white flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-xl bg-teal-600 flex items-center justify-center font-bold">
                  <CheckCircle size={18} />
                </div>
                <div>
                  <h3 className="font-bold text-base">Đóng Giấy Phép Work Permit (Hiện Trường)</h3>
                  <p className="text-xs text-teal-200">Nghiệm thu an toàn, tháo tiếp địa, rút thẻ LOTO và hoàn trả mặt bằng</p>
                </div>
              </div>
              <button onClick={() => setIsCloseWpModalOpen(false)} className="text-slate-300 hover:text-white">
                <X size={20} />
              </button>
            </div>

            <div className="p-6 space-y-4 text-xs">
              <div className="p-3 bg-teal-50 border border-teal-200 rounded-xl space-y-1">
                <div className="font-bold text-teal-950">Mã Work Permit: {activeItem.workPermitId}</div>
                <div className="text-slate-700">Công việc: {activeItem.title}</div>
              </div>

              {/* Safety signoff checklist */}
              <div className="space-y-2 p-3 bg-slate-50 rounded-xl border border-slate-200">
                <div className="font-bold text-slate-800">Xác nhận an toàn hiện trường bắt buộc:</div>
                <div className="space-y-1.5 text-slate-700">
                  <label className="flex items-center gap-2 cursor-pointer">
                    <input type="checkbox" defaultChecked className="rounded text-teal-600" />
                    <span>1. Toàn bộ nhân sự và dụng cụ đã rút khỏi vị trí nguy hiểm</span>
                  </label>
                  <label className="flex items-center gap-2 cursor-pointer">
                    <input type="checkbox" defaultChecked className="rounded text-teal-600" />
                    <span>2. Đã tháo dỡ toàn bộ các bộ dây nối đất lưu động</span>
                  </label>
                  <label className="flex items-center gap-2 cursor-pointer">
                    <input type="checkbox" defaultChecked className="rounded text-teal-600" />
                    <span>3. Đã tháo khóa và gỡ bỏ thẻ LOTO tại tất cả các điểm cô lập</span>
                  </label>
                  <label className="flex items-center gap-2 cursor-pointer">
                    <input type="checkbox" defaultChecked className="rounded text-teal-600" />
                    <span>4. Mặt bằng làm việc sạch sẽ, đủ điều kiện đóng điện chạy lại</span>
                  </label>
                </div>
              </div>

              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">Ghi chú bàn giao & nghiệm thu</label>
                <textarea
                  rows={2}
                  value={activeItem.wpCloseNotes || ''}
                  onChange={(e) => setActiveItem({ ...activeItem, wpCloseNotes: e.target.value })}
                  placeholder="Xác nhận thiết bị vận hành tốt, không rò rỉ, sẵn sàng đóng điện..."
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg"
                />
              </div>
            </div>

            <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between">
              <button
                type="button"
                onClick={() => setIsCloseWpModalOpen(false)}
                className="px-4 py-2 border border-slate-300 rounded-xl text-xs font-semibold text-slate-700 hover:bg-slate-100"
              >
                Hủy bỏ
              </button>

              <button
                type="button"
                onClick={handleCloseWorkPermit}
                className="px-5 py-2.5 bg-teal-600 hover:bg-teal-700 text-white font-bold text-xs rounded-xl shadow-md"
              >
                Ký Xác Nhận Đóng Work Permit Hiện Trường
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================================== */}
      {/* 16. MODAL 6: FINAL CLOSE WO ON CMMS (KẾT THÚC TOÀN BỘ CHU TRÌNH)     */}
      {/* ==================================================================== */}
      {isCloseWoModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex items-center justify-center p-3 sm:p-4 overflow-y-auto">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-xl overflow-hidden my-6 border border-slate-200">
            <div className="px-6 py-4 bg-gradient-to-r from-emerald-900 via-slate-900 to-slate-900 text-white flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-xl bg-emerald-600 flex items-center justify-center font-bold">
                  <CheckCircle size={18} />
                </div>
                <div>
                  <h3 className="font-bold text-base">Đóng Phiếu Công Việc (Close WO Trên CMMS)</h3>
                  <p className="text-xs text-emerald-200">Kết thúc toàn bộ chu trình, lưu hồ sơ lý lịch thiết bị và cập nhật lên Google Sheets</p>
                </div>
              </div>
              <button onClick={() => setIsCloseWoModalOpen(false)} className="text-slate-300 hover:text-white">
                <X size={20} />
              </button>
            </div>

            <div className="p-6 space-y-4 text-xs">
              <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-xl space-y-1">
                <div className="font-bold text-emerald-950">Mã Lệnh WO: {activeItem.id}</div>
                <div className="text-slate-700">Tiêu đề: {activeItem.title}</div>
              </div>

              <div>
                <label className="block text-[11px] font-bold text-slate-700 mb-1">Kết luận kỹ thuật & Nguyên nhân gốc rễ (Root Cause)</label>
                <textarea
                  rows={2}
                  value={activeItem.rootCause || ''}
                  onChange={(e) => setActiveItem({ ...activeItem, rootCause: e.target.value })}
                  placeholder="Ghi nhận nguyên nhân sự cố và biện pháp phòng ngừa tái diễn..."
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Thời gian can thiệp thực tế (Giờ)</label>
                  <input
                    type="number"
                    value={activeItem.actualTimeSpent || ''}
                    onChange={(e) => setActiveItem({ ...activeItem, actualTimeSpent: parseFloat(e.target.value) || 0 })}
                    placeholder="VD: 3.5"
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">Người đóng WO trên hệ thống</label>
                  <input
                    type="text"
                    value={activeItem.woClosedBy || 'Maintenance Manager (MM)'}
                    onChange={(e) => setActiveItem({ ...activeItem, woClosedBy: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg font-medium"
                  />
                </div>
              </div>
            </div>

            <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between">
              <button
                type="button"
                onClick={() => setIsCloseWoModalOpen(false)}
                className="px-4 py-2 border border-slate-300 rounded-xl text-xs font-semibold text-slate-700 hover:bg-slate-100"
              >
                Hủy bỏ
              </button>

              <button
                type="button"
                onClick={handleCloseWorkOrderOnCmms}
                className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-xl shadow-md"
              >
                Hoàn Tất & Đóng WO Trên CMMS
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================================== */}
      {/* 17. MODAL 7: CHỈNH SỬA CHI TIẾT ĐẦY ĐỦ 23 CỘT CHUẨN TEV              */}
      {/* ==================================================================== */}
      {isEdit23ColsModalOpen && (
        <div className={`fixed inset-0 z-50 flex ${
          isFullScreenEdit23Modal 
            ? 'p-0 w-screen h-screen' 
            : 'bg-slate-900/80 backdrop-blur-xs items-center justify-center p-2 sm:p-4 overflow-y-auto'
        }`}>
          <div className={`bg-white shadow-2xl flex flex-col overflow-hidden ${
            isFullScreenEdit23Modal
              ? 'w-full h-full rounded-none border-0'
              : 'w-full max-w-6xl rounded-2xl my-auto max-h-[94vh] border border-slate-200'
          }`}>
            <div className="px-6 py-4 bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white flex items-center justify-between shrink-0 shadow-md">
              <div>
                <div className="flex items-center gap-2">
                  <h3 className="font-bold text-base sm:text-lg">Chỉnh Sửa Chi Tiết Phiếu Công Việc TEV Service Flatform</h3>
                  {isFullScreenEdit23Modal && (
                    <span className="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-blue-500/30 text-blue-300 border border-blue-400/30">
                      Toàn Màn Hình
                    </span>
                  )}
                </div>
                <p className="text-xs text-slate-300">
                  WO: <span className="font-mono font-bold text-emerald-400">{activeItem.id || '---'}</span>
                  {' • '}
                  WP: <span className="font-mono font-bold text-cyan-300">{activeItem.workPermitId || '---'}</span>
                  {' • '}
                  Thiết bị: <span className="font-semibold text-amber-300">{Array.isArray(activeItem.equipmentId) ? activeItem.equipmentId.join(', ') : (activeItem.equipmentId || '---')}</span>
                </p>
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setIsFullScreenEdit23Modal(!isFullScreenEdit23Modal)}
                  className="px-2.5 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-slate-300 hover:text-white transition-colors flex items-center gap-1.5 text-xs font-semibold"
                  title={isFullScreenEdit23Modal ? 'Thu nhỏ cửa sổ' : 'Phóng to toàn màn hình'}
                >
                  {isFullScreenEdit23Modal ? (
                    <>
                      <Minimize2 size={16} />
                      <span className="hidden sm:inline">Thu nhỏ</span>
                    </>
                  ) : (
                    <>
                      <Maximize2 size={16} />
                      <span className="hidden sm:inline">Toàn màn hình</span>
                    </>
                  )}
                </button>
                <button
                  type="button"
                  onClick={() => setIsEdit23ColsModalOpen(false)}
                  className="p-1.5 rounded-lg bg-white/10 hover:bg-rose-600 text-slate-300 hover:text-white transition-colors"
                  title="Đóng cửa sổ"
                >
                  <X size={20} />
                </button>
              </div>
            </div>

            <div className={`p-6 md:p-8 overflow-y-auto space-y-6 text-xs ${
              isFullScreenEdit23Modal ? 'flex-1 max-h-none' : 'max-h-[75vh]'
            }`}>
              {/* KHỐI 1: THÔNG TIN NHẬN DIỆN VÀ PHÂN LOẠI CÔNG VIỆC */}
              <div className="bg-slate-50/70 p-4 rounded-xl border border-slate-200/80 space-y-4">
                <div className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center gap-1.5 border-b border-slate-200 pb-2">
                  <span className="w-2 h-2 rounded-full bg-blue-600"></span>
                  Thông Tin Chung Phiếu Công Việc
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                  <div>
                    <label className="block text-[11px] font-bold text-slate-700 mb-1">1. Mã WO</label>
                    <input
                      type="text"
                      value={activeItem.id || ''}
                      onChange={(e) => setActiveItem({ ...activeItem, id: e.target.value })}
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg font-mono font-bold text-emerald-800 bg-emerald-50/20"
                    />
                  </div>
                  <div>
                    <label className="block text-[11px] font-bold text-slate-700 mb-1">2. Mã Work Permit</label>
                    <input
                      type="text"
                      value={activeItem.workPermitId || ''}
                      onChange={(e) => setActiveItem({ ...activeItem, workPermitId: e.target.value })}
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg font-mono font-bold text-blue-800 bg-blue-50/20"
                    />
                  </div>
                  <div>
                    <label className="block text-[11px] font-bold text-slate-700 mb-1">6. Failure Code</label>
                    <select
                      value={activeItem.failureCode || 'ROUTINE-PM'}
                      onChange={(e) => setActiveItem({ ...activeItem, failureCode: e.target.value })}
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg font-semibold text-rose-700 bg-white"
                    >
                      {FAILURE_CODES.map(fc => (
                        <option key={fc.code} value={fc.code}>{fc.code} - {fc.desc}</option>
                      ))}
                    </select>
                  </div>
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">3. Tiêu đề công việc</label>
                  <input
                    type="text"
                    value={activeItem.title || ''}
                    onChange={(e) => setActiveItem({ ...activeItem, title: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg font-semibold"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-slate-700 mb-1">4. Mô tả chi tiết</label>
                  <textarea
                    rows={2}
                    value={activeItem.description || ''}
                    onChange={(e) => setActiveItem({ ...activeItem, description: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs"
                  />
                </div>
              </div>

              {/* KHỐI 2: ĐIỀU ĐỘ, TIẾN ĐỘ & NGƯỜI THỰC HIỆN */}
              <div className="bg-slate-50/70 p-4 rounded-xl border border-slate-200/80 space-y-4">
                <div className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center gap-1.5 border-b border-slate-200 pb-2">
                  <span className="w-2 h-2 rounded-full bg-indigo-600"></span>
                  Tiến Độ, Mức Ưu Tiên & Phân Công Nhân Sự
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                  <div>
                    <label className="block text-[11px] font-bold text-slate-700 mb-1">10. Mức ưu tiên (Priority)</label>
                    <select
                      value={activeItem.priorityCode || 'P3'}
                      onChange={(e) => setActiveItem({ ...activeItem, priorityCode: e.target.value as any })}
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg font-bold bg-white"
                    >
                      <option value="P1">P1: Làm ngay (Emergency/Breakdown)</option>
                      <option value="P3">P3: Trong 1 tuần (High Priority)</option>
                      <option value="P6">P6: Bảo trì định kỳ (Scheduled PM)</option>
                      <option value="P7">P7: Shutdown (Dừng nhà máy)</option>
                      <option value="P9">P9: DMR (Defect & Reliability)</option>
                    </select>
                  </div>

                  <div>
                    <label className="block text-[11px] font-bold text-slate-700 mb-1">11. Vòng đời CMMS Stage</label>
                    <select
                      value={activeItem.workflowStage || '3_planning'}
                      onChange={(e) => setActiveItem({ ...activeItem, workflowStage: e.target.value as any })}
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg font-bold text-blue-700 bg-white"
                    >
                      {Object.entries(WORKFLOW_STAGES).map(([k, v]) => (
                        <option key={k} value={k}>{v.label}</option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <FseEngineerDropdown
                      label="12. Người thực hiện (PIC)"
                      value={activeItem.assignedTo || ''}
                      onChange={(val) => setActiveItem({ ...activeItem, assignedTo: val })}
                    />
                  </div>
                </div>
              </div>

              {/* KHỐI 3: MỐC THỜI GIAN DỪNG MÁY (CỘT 20-23) - CALENDAR PICKER CHUẨN */}
              <div className="p-5 bg-gradient-to-br from-slate-50 to-indigo-50/30 border border-slate-200 rounded-2xl space-y-4 shadow-xs">
                <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-200/80 pb-3">
                  <div>
                    <span className="font-bold text-slate-900 text-sm flex items-center gap-2">
                      <CalendarDays size={18} className="text-indigo-600" />
                      Mốc Thời Gian Dừng Máy &amp; Sửa Chữa (Chuẩn TEV Service Flatform)
                    </span>
                    <p className="text-[11px] text-slate-500 mt-0.5">
                      Lựa chọn trực tiếp trên Calendar · Hỗ trợ bấm "Bây giờ" hoặc mở lịch chọn ngày giờ chính xác · Chuẩn tính toán ARP-E
                    </p>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="text-[11px] font-mono text-indigo-700 bg-indigo-50 border border-indigo-200 px-2.5 py-1 rounded-lg flex items-center gap-1.5 font-semibold">
                      <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse inline-block"></span>
                      Định dạng chuẩn: YYYY-MM-DD HH:mm
                    </span>
                  </div>
                </div>

                {/* Milestone Step Progress Indicator */}
                <div className="grid grid-cols-2 md:grid-cols-4 gap-2 text-center text-[10px] font-bold text-slate-600 pb-1">
                  <div className="p-1.5 bg-rose-50 border border-rose-200 rounded-lg text-rose-800">
                    Bước 1: 🛑 Dừng máy
                  </div>
                  <div className="p-1.5 bg-amber-50 border border-amber-200 rounded-lg text-amber-800">
                    Bước 2: 🔧 Bắt đầu sửa
                  </div>
                  <div className="p-1.5 bg-blue-50 border border-blue-200 rounded-lg text-blue-800">
                    Bước 3: 🏁 Kết thúc sửa
                  </div>
                  <div className="p-1.5 bg-emerald-50 border border-emerald-200 rounded-lg text-emerald-800">
                    Bước 4: ⚡ Chạy lại máy
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                  {/* 20. Bắt đầu dừng máy */}
                  <div className="bg-white p-3.5 rounded-xl border-2 border-rose-200 shadow-xs hover:border-rose-400 transition-all space-y-2">
                    <div className="flex items-center justify-between">
                      <label className="text-[11px] text-rose-700 font-bold flex items-center gap-1.5">
                        <span className="w-2.5 h-2.5 rounded-full bg-rose-500"></span>
                        20. Bắt đầu dừng máy
                      </label>
                      <button
                        type="button"
                        onClick={() => {
                          const now = getNowFormattedDatetime();
                          setActiveItem({ ...activeItem, downtimeStart: now });
                        }}
                        className="px-2 py-0.5 rounded bg-rose-50 hover:bg-rose-100 text-[10px] text-rose-700 font-bold transition-colors"
                        title="Điền thời gian hiện tại"
                      >
                        🕒 Bây giờ
                      </button>
                    </div>

                    <div className="relative">
                      <input
                        id="downtimeStartInput"
                        type="datetime-local"
                        value={parseToDatetimeLocal(activeItem.downtimeStart)}
                        onChange={(e) => {
                          const formatted = formatFromDatetimeLocal(e.target.value);
                          setActiveItem({ ...activeItem, downtimeStart: formatted });
                        }}
                        onClick={(e) => {
                          try { (e.target as any).showPicker?.(); } catch {}
                        }}
                        className="w-full text-xs font-mono font-bold border border-slate-300 rounded-lg pl-8 pr-2 py-2 bg-slate-50/40 hover:bg-white focus:bg-white focus:border-rose-500 focus:ring-2 focus:ring-rose-200 outline-none transition-all cursor-pointer text-slate-800"
                      />
                      <CalendarDays size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-rose-600 pointer-events-none" />
                    </div>

                    <div className="flex items-center justify-between text-[11px] pt-1 border-t border-slate-100">
                      <span className="font-mono text-slate-600 truncate">
                        {activeItem.downtimeStart || <span className="text-slate-400 italic">Chưa ghi nhận</span>}
                      </span>
                      {activeItem.downtimeStart && (
                        <button
                          type="button"
                          onClick={() => setActiveItem({ ...activeItem, downtimeStart: '' })}
                          className="text-slate-400 hover:text-rose-600 ml-1 font-semibold text-[10px]"
                        >
                          Xóa
                        </button>
                      )}
                    </div>
                  </div>

                  {/* 21. Bắt đầu sửa */}
                  <div className="bg-white p-3.5 rounded-xl border-2 border-amber-200 shadow-xs hover:border-amber-400 transition-all space-y-2">
                    <div className="flex items-center justify-between">
                      <label className="text-[11px] text-amber-700 font-bold flex items-center gap-1.5">
                        <span className="w-2.5 h-2.5 rounded-full bg-amber-500"></span>
                        21. Bắt đầu sửa
                      </label>
                      <button
                        type="button"
                        onClick={() => {
                          const now = getNowFormattedDatetime();
                          setActiveItem({ ...activeItem, repairStart: now });
                        }}
                        className="px-2 py-0.5 rounded bg-amber-50 hover:bg-amber-100 text-[10px] text-amber-700 font-bold transition-colors"
                        title="Điền thời gian hiện tại"
                      >
                        🕒 Bây giờ
                      </button>
                    </div>

                    <div className="relative">
                      <input
                        id="repairStartInput"
                        type="datetime-local"
                        value={parseToDatetimeLocal(activeItem.repairStart)}
                        onChange={(e) => {
                          const formatted = formatFromDatetimeLocal(e.target.value);
                          setActiveItem({ ...activeItem, repairStart: formatted });
                        }}
                        onClick={(e) => {
                          try { (e.target as any).showPicker?.(); } catch {}
                        }}
                        className="w-full text-xs font-mono font-bold border border-slate-300 rounded-lg pl-8 pr-2 py-2 bg-slate-50/40 hover:bg-white focus:bg-white focus:border-amber-500 focus:ring-2 focus:ring-amber-200 outline-none transition-all cursor-pointer text-slate-800"
                      />
                      <CalendarDays size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-amber-600 pointer-events-none" />
                    </div>

                    <div className="flex items-center justify-between text-[11px] pt-1 border-t border-slate-100">
                      <span className="font-mono text-slate-600 truncate">
                        {activeItem.repairStart || <span className="text-slate-400 italic">Chưa ghi nhận</span>}
                      </span>
                      {activeItem.repairStart && (
                        <button
                          type="button"
                          onClick={() => setActiveItem({ ...activeItem, repairStart: '' })}
                          className="text-slate-400 hover:text-amber-600 ml-1 font-semibold text-[10px]"
                        >
                          Xóa
                        </button>
                      )}
                    </div>
                  </div>

                  {/* 22. Kết thúc sửa */}
                  <div className="bg-white p-3.5 rounded-xl border-2 border-blue-200 shadow-xs hover:border-blue-400 transition-all space-y-2">
                    <div className="flex items-center justify-between">
                      <label className="text-[11px] text-blue-700 font-bold flex items-center gap-1.5">
                        <span className="w-2.5 h-2.5 rounded-full bg-blue-500"></span>
                        22. Kết thúc sửa
                      </label>
                      <button
                        type="button"
                        onClick={() => {
                          const now = getNowFormattedDatetime();
                          setActiveItem({ ...activeItem, repairEnd: now });
                        }}
                        className="px-2 py-0.5 rounded bg-blue-50 hover:bg-blue-100 text-[10px] text-blue-700 font-bold transition-colors"
                        title="Điền thời gian hiện tại"
                      >
                        🕒 Bây giờ
                      </button>
                    </div>

                    <div className="relative">
                      <input
                        id="repairEndInput"
                        type="datetime-local"
                        value={parseToDatetimeLocal(activeItem.repairEnd)}
                        onChange={(e) => {
                          const formatted = formatFromDatetimeLocal(e.target.value);
                          setActiveItem({ ...activeItem, repairEnd: formatted });
                        }}
                        onClick={(e) => {
                          try { (e.target as any).showPicker?.(); } catch {}
                        }}
                        className="w-full text-xs font-mono font-bold border border-slate-300 rounded-lg pl-8 pr-2 py-2 bg-slate-50/40 hover:bg-white focus:bg-white focus:border-blue-500 focus:ring-2 focus:ring-blue-200 outline-none transition-all cursor-pointer text-slate-800"
                      />
                      <CalendarDays size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-blue-600 pointer-events-none" />
                    </div>

                    <div className="flex items-center justify-between text-[11px] pt-1 border-t border-slate-100">
                      <span className="font-mono text-slate-600 truncate">
                        {activeItem.repairEnd || <span className="text-slate-400 italic">Chưa ghi nhận</span>}
                      </span>
                      {activeItem.repairEnd && (
                        <button
                          type="button"
                          onClick={() => setActiveItem({ ...activeItem, repairEnd: '' })}
                          className="text-slate-400 hover:text-blue-600 ml-1 font-semibold text-[10px]"
                        >
                          Xóa
                        </button>
                      )}
                    </div>
                  </div>

                  {/* 23. Chạy lại máy */}
                  <div className="bg-white p-3.5 rounded-xl border-2 border-emerald-200 shadow-xs hover:border-emerald-400 transition-all space-y-2">
                    <div className="flex items-center justify-between">
                      <label className="text-[11px] text-emerald-700 font-bold flex items-center gap-1.5">
                        <span className="w-2.5 h-2.5 rounded-full bg-emerald-500"></span>
                        23. Chạy lại máy
                      </label>
                      <button
                        type="button"
                        onClick={() => {
                          const now = getNowFormattedDatetime();
                          setActiveItem({ ...activeItem, restartTime: now });
                        }}
                        className="px-2 py-0.5 rounded bg-emerald-50 hover:bg-emerald-100 text-[10px] text-emerald-700 font-bold transition-colors"
                        title="Điền thời gian hiện tại"
                      >
                        🕒 Bây giờ
                      </button>
                    </div>

                    <div className="relative">
                      <input
                        id="restartTimeInput"
                        type="datetime-local"
                        value={parseToDatetimeLocal(activeItem.restartTime)}
                        onChange={(e) => {
                          const formatted = formatFromDatetimeLocal(e.target.value);
                          setActiveItem({ ...activeItem, restartTime: formatted });
                        }}
                        onClick={(e) => {
                          try { (e.target as any).showPicker?.(); } catch {}
                        }}
                        className="w-full text-xs font-mono font-bold border border-slate-300 rounded-lg pl-8 pr-2 py-2 bg-slate-50/40 hover:bg-white focus:bg-white focus:border-emerald-500 focus:ring-2 focus:ring-emerald-200 outline-none transition-all cursor-pointer text-slate-800"
                      />
                      <CalendarDays size={15} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-emerald-600 pointer-events-none" />
                    </div>

                    <div className="flex items-center justify-between text-[11px] pt-1 border-t border-slate-100">
                      <span className="font-mono text-slate-600 truncate">
                        {activeItem.restartTime || <span className="text-slate-400 italic">Chưa ghi nhận</span>}
                      </span>
                      {activeItem.restartTime && (
                        <button
                          type="button"
                          onClick={() => setActiveItem({ ...activeItem, restartTime: '' })}
                          className="text-slate-400 hover:text-emerald-600 ml-1 font-semibold text-[10px]"
                        >
                          Xóa
                        </button>
                      )}
                    </div>
                  </div>
                </div>

                {/* Live Calculated Durations (ARP-E Standard) */}
                {(() => {
                  const waitHrs = calculateMilestoneDiffHours(activeItem.downtimeStart, activeItem.repairStart);
                  const repairHrs = calculateMilestoneDiffHours(activeItem.repairStart, activeItem.repairEnd);
                  const restartHrs = calculateMilestoneDiffHours(activeItem.repairEnd, activeItem.restartTime);
                  const totalDtHrs = calculateMilestoneDiffHours(activeItem.downtimeStart, activeItem.restartTime);

                  return (
                    <div className="bg-slate-900 text-white rounded-xl p-4 flex flex-wrap items-center justify-between gap-4 text-xs shadow-inner">
                      <div className="flex flex-wrap items-center gap-5">
                        <div className="flex items-center gap-1.5">
                          <span className="text-slate-400 text-[11px]">Thời gian chờ (T_wait):</span>
                          <span className="font-mono font-bold text-amber-400 text-sm">
                            {waitHrs !== null ? `${waitHrs} giờ` : '---'}
                          </span>
                        </div>
                        <div className="w-px h-5 bg-slate-700 hidden sm:block"></div>
                        <div className="flex items-center gap-1.5">
                          <span className="text-slate-400 text-[11px]">Sửa chữa thực tế (T_rep):</span>
                          <span className="font-mono font-bold text-emerald-400 text-sm">
                            {repairHrs !== null ? `${repairHrs} giờ` : '---'}
                          </span>
                        </div>
                        <div className="w-px h-5 bg-slate-700 hidden sm:block"></div>
                        <div className="flex items-center gap-1.5">
                          <span className="text-slate-400 text-[11px]">Kiểm tra chạy lại:</span>
                          <span className="font-mono font-bold text-blue-400 text-sm">
                            {restartHrs !== null ? `${restartHrs} giờ` : '---'}
                          </span>
                        </div>
                        <div className="w-px h-5 bg-slate-700 hidden sm:block"></div>
                        <div className="flex items-center gap-1.5">
                          <span className="text-slate-400 text-[11px]">Tổng dừng máy (Total DT):</span>
                          <span className="font-mono font-black text-rose-400 text-sm">
                            {totalDtHrs !== null ? `${totalDtHrs} giờ` : '---'}
                          </span>
                        </div>
                      </div>

                      {waitHrs !== null && repairHrs !== null && waitHrs > repairHrs && (
                        <span className="px-2.5 py-1 rounded-lg text-[11px] font-bold bg-amber-500/20 text-amber-300 border border-amber-500/40">
                          ⚠️ MWT ({waitHrs}h) &gt; MTTR ({repairHrs}h) - Cần tối ưu kho vật tư & khâu lập kế hoạch
                        </span>
                      )}
                    </div>
                  );
                })()}
              </div>
            </div>

            <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between">
              <button
                type="button"
                onClick={() => setIsEdit23ColsModalOpen(false)}
                className="px-4 py-2 border border-slate-300 rounded-xl text-xs font-semibold text-slate-700 hover:bg-slate-100"
              >
                Hủy bỏ
              </button>

              <button
                type="button"
                onClick={async () => {
                  if (activeItem.id) {
                    await onSaveWorkOrder(activeItem as WorkOrderItem, true);
                    setIsEdit23ColsModalOpen(false);
                    setActionNotice(`Đã cập nhật chi tiết cho [${activeItem.id}] thành công!`);
                    setTimeout(() => setActionNotice(null), 4000);
                  }
                }}
                className="px-5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs rounded-xl shadow-md"
              >
                Lưu Thay Đổi & Đồng Bộ Sheets
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================================== */}
      {/* 18. MODAL 8: IN GIẤY PHÉP WORK PERMIT KHỔ A4 CHUẨN CÔNG NGHIỆP        */}
      {/* ==================================================================== */}
      {isPrintPermitOpen && permitToPrint && (
        <div className="fixed inset-0 bg-slate-900/70 backdrop-blur-xs z-50 flex items-center justify-center p-3 sm:p-4 overflow-y-auto">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-3xl overflow-hidden my-6 border border-slate-300">
            {/* Action Bar */}
            <div className="px-6 py-3 bg-slate-900 text-white flex items-center justify-between print:hidden">
              <div className="flex items-center gap-2">
                <Printer size={18} className="text-amber-400" />
                <span className="font-bold text-sm">Xem Trước Bản In Giấy Phép Work Permit (A4 Chuẩn Nhà Máy)</span>
              </div>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => window.print()}
                  className="px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-lg flex items-center gap-1.5 shadow-xs transition-colors"
                >
                  <Printer size={14} />
                  <span>In Ra Giấy / Lưu PDF</span>
                </button>
                <button
                  onClick={() => setIsPrintPermitOpen(false)}
                  className="p-1 rounded-lg hover:bg-white/10 text-slate-400 hover:text-white"
                >
                  <X size={20} />
                </button>
              </div>
            </div>

            {/* A4 Printable Sheet */}
            <div className="p-8 text-slate-900 space-y-5 font-serif text-xs">
              <div className="border-b-2 border-slate-900 pb-3 text-center space-y-1">
                <div className="text-[11px] uppercase font-sans tracking-widest text-slate-600 font-bold">
                  CÔNG TY TNHH KỸ THUẬT NĂNG LƯỢNG TEV - BAN QUẢN LÝ AN TOÀN & BẢO TRÌ NHÀ MÁY
                </div>
                <h2 className="text-lg font-bold uppercase tracking-wide text-slate-900">
                  GIẤY PHÉP LÀM VIỆC VÀ CÔ LẬP NĂNG LƯỢNG (PERMIT TO WORK & LOTO)
                </h2>
                <div className="text-[11px] italic text-slate-600">
                  (Căn cứ quy chuẩn kỹ thuật an toàn điện & quy trình phê duyệt Daily Meeting CMMS)
                </div>
                <div className="flex justify-between items-center text-[11px] font-mono font-bold pt-2 px-2 border-t border-slate-200 mt-2">
                  <span>MÃ GIẤY PHÉP (WP): {permitToPrint.workPermitId || 'WP-PENDING'}</span>
                  <span>MÃ LỆNH CÔNG VIỆC (WO): {permitToPrint.id}</span>
                  {permitToPrint.wrId && <span>MÃ YÊU CẦU (WR): {permitToPrint.wrId}</span>}
                </div>
              </div>

              {/* General Info */}
              <div className="border border-slate-300 p-3 rounded space-y-1.5">
                <div className="grid grid-cols-2 gap-2">
                  <div><strong>Khách hàng / Đơn vị:</strong> {permitToPrint.customer}</div>
                  <div><strong>Phân xưởng / Trạm:</strong> {permitToPrint.factory || 'Trạm chính'}</div>
                  <div><strong>Thiết bị thực hiện:</strong> {Array.isArray(permitToPrint.equipmentId) ? permitToPrint.equipmentId.join(', ') : permitToPrint.equipmentId}</div>
                  <div><strong>Mức ưu tiên:</strong> {permitToPrint.priorityCode || 'P3'}</div>
                  <div><strong>Chỉ huy trực tiếp (PIC):</strong> {permitToPrint.assignedTo}</div>
                  <div><strong>Thời hạn hiệu lực:</strong> {permitToPrint.dueDate || 'Trong ngày'}</div>
                  <div className="col-span-2"><strong>Nội dung công việc:</strong> {permitToPrint.title}</div>
                  <div className="col-span-2"><strong>Mô tả chi tiết:</strong> {permitToPrint.description || 'Thực hiện bảo dưỡng an toàn'}</div>
                </div>
              </div>

              {/* Safety LOTO Checklist */}
              <div className="border border-slate-300 p-3 rounded space-y-1.5">
                <div className="font-bold uppercase text-slate-800">
                  Biện Pháp Cô Lập Nguồn & An Toàn Bắt Buộc (LOTO Safety Conditions):
                </div>
                <div className="grid grid-cols-2 gap-1.5">
                  <div>✓ Đã cắt điện và cô lập tại: <strong>{permitToPrint.isolationPoints || 'MC tổng'}</strong></div>
                  <div>✓ Đã treo thẻ khóa LOTO số: <strong>{permitToPrint.lotoTagNumbers || 'LOTO-TR-01'}</strong></div>
                  <div>✓ Đã dùng bút thử điện kiểm tra không còn áp</div>
                  <div>✓ Đã nối đất di động lưu động 3 pha an toàn</div>
                  <div>✓ Đặt rào chắn, biển báo "CẤM ĐÓNG ĐIỆN! CÓ NGƯỜI LÀM VIỆC"</div>
                  <div>✓ Đầy đủ trang bị BHLĐ cá nhân (Găng tay cách điện, kính, mũ)</div>
                </div>
              </div>

              {/* Timestamps */}
              <div className="grid grid-cols-4 gap-2 border border-slate-300 p-2 rounded text-center">
                <div>
                  <div className="font-bold text-[10px] uppercase">Bắt Đầu Dừng</div>
                  <div className="font-mono text-[10px]">{permitToPrint.downtimeStart ? new Date(permitToPrint.downtimeStart).toLocaleTimeString('vi-VN') : '---'}</div>
                </div>
                <div>
                  <div className="font-bold text-[10px] uppercase">Bắt Đầu Sửa</div>
                  <div className="font-mono text-[10px]">{permitToPrint.repairStart ? new Date(permitToPrint.repairStart).toLocaleTimeString('vi-VN') : '---'}</div>
                </div>
                <div>
                  <div className="font-bold text-[10px] uppercase">Kết Thúc Sửa</div>
                  <div className="font-mono text-[10px]">{permitToPrint.repairEnd ? new Date(permitToPrint.repairEnd).toLocaleTimeString('vi-VN') : '---'}</div>
                </div>
                <div>
                  <div className="font-bold text-[10px] uppercase">Chạy Lại Máy</div>
                  <div className="font-mono text-[10px]">{permitToPrint.restartTime ? new Date(permitToPrint.restartTime).toLocaleTimeString('vi-VN') : '---'}</div>
                </div>
              </div>

              {/* Signatures */}
              <div className="grid grid-cols-3 gap-4 text-center text-xs pt-3">
                <div className="space-y-10">
                  <div>
                    <div className="font-bold">CHỈ HUY TRỰC TIẾP (PIC)</div>
                    <div className="text-[10px] italic">(Tiếp nhận hiện trường & thi công)</div>
                  </div>
                  <div className="font-bold">{permitToPrint.assignedTo || 'Kỹ sư FSE'}</div>
                </div>

                <div className="space-y-10">
                  <div>
                    <div className="font-bold">CÁN BỘ AN TOÀN (SO)</div>
                    <div className="text-[10px] italic">(Kiểm tra điều kiện LOTO & cấp WP)</div>
                  </div>
                  <div className="font-bold">{permitToPrint.wpIssuedBy || 'KS. Nguyễn Văn An'}</div>
                </div>

                <div className="space-y-10">
                  <div>
                    <div className="font-bold">NGƯỜI PHÊ DUYỆT ({permitToPrint.dailyMeetingApprovedByRole || 'MM'})</div>
                    <div className="text-[10px] italic">(Họp Daily Meeting phê duyệt)</div>
                  </div>
                  <div className="font-bold">{permitToPrint.dailyMeetingApprovedByName || 'Trưởng phòng Bảo trì'}</div>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
