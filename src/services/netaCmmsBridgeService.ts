/**
 * NETA-CMMS DATA BRIDGE & UNIFIED ID GENERATION SERVICE
 * 
 * 1. Formats and auto-increments IDs:
 *    - Work Order: WO-YYYY-XXX (e.g. WO-2026-001, WO-2026-013)
 *    - Work Permit: WP-YYYY-XXX (e.g. WP-2026-001, WP-2026-013)
 *    - Customer: KHXXX (e.g. KH001, KH002, KH006)
 * 
 * 2. Resolves field discrepancies between FSE (NETA ATS/MTS) and CMMS:
 *    - Harmonizes NETA inspection checklists into canonical CMMS Work Orders
 *    - Ensures all 31 CMMS columns and WorkOrder interfaces are cleanly populated
 *    - Synchronizes equipment health and Google Sheets rows
 */

import { CustomerRecord, EquipmentMasterRecord, DEFAULT_CUSTOMERS, DEFAULT_WORK_ORDERS } from '../utils/fieldSyncService';
import { getRegisteredCustomers } from './masterDirectoryService';

export const EVT_CMMS_WO_UPDATED = 'tev_cmms_wo_updated';
export const EVT_CMMS_REPORT_SAVED = 'tev_cmms_report_saved';

export interface CmmsWorkOrder {
  id: string; // WO-YYYY-XXX
  workPermitId: string; // WP-YYYY-XXX
  title: string;
  description: string;
  equipmentId: string | string[];
  equipmentName?: string;
  customerId?: string; // KHXXX
  customer: string;
  factory: string;
  type: string;
  priority: string;
  status: string;
  isUnplanned: boolean;
  assignedTo: string;
  responsibleApprove: string;
  responsibleDo: string;
  blockingRequired: boolean;
  dueDate: string;
  createdAt: string;
  updatedAt: string;
  completedAt?: string | null;
  pmFrequency?: string;
  failureCode?: string;
  rootCause?: string;
  actualTimeSpent?: number;
  estimatedTime?: number;
  laborCount?: number;
  laborCost?: number;
  partCost?: number;
  attachments?: string[];
  usedMaterials?: any;
  downtimeStart?: string;
  repairStart?: string;
  repairEnd?: string;
  restartTime?: string;
  // Inspection & Linkage fields
  healthIndex?: number;
  netaReportId?: string;
  inspectionSummary?: string;
  testStandard?: string;
  testDate?: string;
  reportData?: any;
}

// ============================================================================
// 1. STANDARDIZED ID GENERATION (WO-YYYY-XXX, WP-YYYY-XXX, KHXXX)
// ============================================================================

/**
 * Calculates the next Work Order ID in format: WO-YYYY-XXX
 * Automatically extracts the highest existing sequence number for the current year and adds 1.
 */
export function getNextWorkOrderId(existingOrders?: any[], targetYear?: number): string {
  const currentYear = targetYear || new Date().getFullYear();
  const orders = Array.isArray(existingOrders) && existingOrders.length > 0 
    ? existingOrders 
    : getStoredWorkOrdersFallback();

  let maxSeq = 0;

  orders.forEach(wo => {
    if (!wo) return;
    const rawId = String(wo.id || '').trim();
    if (!rawId) return;

    // Pattern 1: WO-2026-001, WO-2026-1, WO-2025-012
    const matchFull = rawId.match(/^WO-(\d{4})-(\d+)$/i);
    if (matchFull) {
      const yr = parseInt(matchFull[1], 10);
      const seq = parseInt(matchFull[2], 10);
      if (yr === currentYear && !isNaN(seq) && seq > maxSeq) {
        maxSeq = seq;
      }
      return;
    }

    // Pattern 2: WO-PM-2026-001
    const matchPm = rawId.match(/^WO-PM-(\d{4})-(\d+)$/i);
    if (matchPm) {
      const yr = parseInt(matchPm[1], 10);
      const seq = parseInt(matchPm[2], 10);
      if (yr === currentYear && !isNaN(seq) && seq > maxSeq) {
        maxSeq = seq;
      }
      return;
    }

    // Pattern 3: WO-001, WO-12 (assumed current year)
    const matchSimple = rawId.match(/^WO-(\d+)$/i);
    if (matchSimple) {
      const seq = parseInt(matchSimple[1], 10);
      if (!isNaN(seq) && seq < 10000 && seq > maxSeq) {
        maxSeq = seq;
      }
      return;
    }

    // Pattern 4: Trailing digits
    const trailing = rawId.match(/(\d+)$/);
    if (trailing) {
      const seq = parseInt(trailing[1], 10);
      if (!isNaN(seq) && seq < 10000 && seq > maxSeq) {
        maxSeq = seq;
      }
    }
  });

  const nextSeq = maxSeq + 1;
  const seqStr = String(nextSeq).padStart(3, '0');
  return `WO-${currentYear}-${seqStr}`;
}

/**
 * Calculates the next Work Permit ID in format: WP-YYYY-XXX
 * Automatically extracts the highest existing sequence number for the current year and adds 1.
 */
export function getNextWorkPermitId(existingOrders?: any[], targetYear?: number): string {
  const currentYear = targetYear || new Date().getFullYear();
  const orders = Array.isArray(existingOrders) && existingOrders.length > 0 
    ? existingOrders 
    : getStoredWorkOrdersFallback();

  let maxSeq = 0;

  orders.forEach(wo => {
    if (!wo) return;
    const rawWp = String(wo.workPermitId || '').trim();
    if (!rawWp) return;

    // Pattern 1: WP-2026-001
    const matchFull = rawWp.match(/^WP-(\d{4})-(\d+)$/i);
    if (matchFull) {
      const yr = parseInt(matchFull[1], 10);
      const seq = parseInt(matchFull[2], 10);
      if (yr === currentYear && !isNaN(seq) && seq > maxSeq) {
        maxSeq = seq;
      }
      return;
    }

    // Pattern 2: WP-PM-2026-001
    const matchPm = rawWp.match(/^WP-PM-(\d{4})-(\d+)$/i);
    if (matchPm) {
      const yr = parseInt(matchPm[1], 10);
      const seq = parseInt(matchPm[2], 10);
      if (yr === currentYear && !isNaN(seq) && seq > maxSeq) {
        maxSeq = seq;
      }
      return;
    }

    // Pattern 3: WP-001
    const matchSimple = rawWp.match(/^WP-(\d+)$/i);
    if (matchSimple) {
      const seq = parseInt(matchSimple[1], 10);
      if (!isNaN(seq) && seq < 10000 && seq > maxSeq) {
        maxSeq = seq;
      }
      return;
    }

    // Pattern 4: Trailing digits
    const trailing = rawWp.match(/(\d+)$/);
    if (trailing) {
      const seq = parseInt(trailing[1], 10);
      if (!isNaN(seq) && seq < 10000 && seq > maxSeq) {
        maxSeq = seq;
      }
    }
  });

  // If no WP exists or max is 0, align with the WO sequence if available
  if (maxSeq === 0) {
    const woNext = getNextWorkOrderId(orders, currentYear);
    const woMatch = woNext.match(/(\d+)$/);
    if (woMatch) {
      maxSeq = parseInt(woMatch[1], 10) - 1;
    }
  }

  const nextSeq = maxSeq + 1;
  const seqStr = String(nextSeq).padStart(3, '0');
  return `WP-${currentYear}-${seqStr}`;
}

/**
 * Calculates the next Work Request (WR) ID in format: WR-YYYY-XXX
 */
export function getNextWorkRequestId(existingOrders?: any[], targetYear?: number): string {
  const currentYear = targetYear || new Date().getFullYear();
  const orders = Array.isArray(existingOrders) && existingOrders.length > 0 
    ? existingOrders 
    : getStoredWorkOrdersFallback();

  let maxSeq = 0;

  orders.forEach(wo => {
    if (!wo) return;
    const rawWr = String(wo.wrId || wo.workRequestId || '').trim();
    if (!rawWr) return;

    const matchFull = rawWr.match(/^WR-(\d{4})-(\d+)$/i);
    if (matchFull) {
      const yr = parseInt(matchFull[1], 10);
      const seq = parseInt(matchFull[2], 10);
      if (yr === currentYear && !isNaN(seq) && seq > maxSeq) {
        maxSeq = seq;
      }
      return;
    }

    const trailing = rawWr.match(/(\d+)$/);
    if (trailing) {
      const seq = parseInt(trailing[1], 10);
      if (!isNaN(seq) && seq < 10000 && seq > maxSeq) {
        maxSeq = seq;
      }
    }
  });

  const nextSeq = maxSeq + 1;
  const seqStr = String(nextSeq).padStart(3, '0');
  return `WR-${currentYear}-${seqStr}`;
}

/**
 * Calculates the next Customer Code strictly in format: KHXXX (e.g. KH001, KH002, KH003...)
 * Scans all existing customers (KH001, KH-001, KH-TEV-01, etc.) and extracts highest integer.
 */
export function getNextCustomerId(existingCustomers?: CustomerRecord[]): string {
  const customers = Array.isArray(existingCustomers) && existingCustomers.length > 0 
    ? existingCustomers 
    : getRegisteredCustomers();

  let maxSeq = 0;

  customers.forEach(cust => {
    if (!cust) return;
    const rawId = String(cust.id || '').trim();
    if (!rawId) return;

    // Pattern 1: KH001, KH012, KH-001, KH-1
    const matchDirect = rawId.match(/^KH-?(\d+)$/i);
    if (matchDirect) {
      const seq = parseInt(matchDirect[1], 10);
      if (!isNaN(seq) && seq > maxSeq) {
        maxSeq = seq;
      }
      return;
    }

    // Pattern 2: KH-TEV-01, KH-SEVT-02, KH-VINFAST-03, KH-HOAPHAT-04, KH-SUPER-05
    const matchSuffix = rawId.match(/^KH-.*-(\d+)$/i);
    if (matchSuffix) {
      const seq = parseInt(matchSuffix[1], 10);
      if (!isNaN(seq) && seq > maxSeq) {
        maxSeq = seq;
      }
      return;
    }

    // Pattern 3: Any trailing digits after prefix
    const matchTrailing = rawId.match(/(\d+)$/);
    if (matchTrailing) {
      const seq = parseInt(matchTrailing[1], 10);
      if (!isNaN(seq) && seq < 10000 && seq > maxSeq) {
        maxSeq = seq;
      }
    }
  });

  const nextSeq = maxSeq + 1;
  const seqStr = String(nextSeq).padStart(3, '0');
  return `KH${seqStr}`;
}

function getStoredWorkOrdersFallback(): any[] {
  try {
    const saved = localStorage.getItem('tev_cached_work_orders');
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) return parsed;
    }
  } catch {}
  return DEFAULT_WORK_ORDERS;
}

// ============================================================================
// 2. FSE (NETA) TO CMMS DATA HARMONIZATION ADAPTER
// ============================================================================

export interface HarmonizeFseParams {
  rawPayload?: any;
  data?: any;
  siteInfo?: any;
  analysisResult?: any;
  evaluation?: any;
  equipmentId?: string;
  equipmentName?: string;
  equipmentType?: string;
  customer?: string;
  customerId?: string;
  factory?: string;
  fseName?: string;
  testDate?: string;
  type?: string;
  existingWorkOrders?: any[];
  existingCustomers?: CustomerRecord[];
  existingEquipment?: EquipmentMasterRecord[];
}

export interface HarmonizedResult {
  workOrder: CmmsWorkOrder;
  report: any;
  equipmentUpdate: {
    id: string;
    health: number;
    status: 'healthy' | 'warning' | 'critical';
    lastCheck: string;
  };
}

/**
 * Harmonizes and maps disparate FSE (NETA) fields into the canonical CMMS schema.
 * Reconciles:
 *  - site_info (project_name, transformer_tag, fse_name, test_date) -> customer, equipmentId, assignedTo, dueDate
 *  - overall_status (PASS/INVESTIGATE/FAIL) -> status, priority, failureCode
 *  - health_index -> health (0-100)
 *  - test measurements -> description & report notes
 *  - auto-allocates WO-YYYY-XXX and WP-YYYY-XXX
 */
export function harmonizeFseToCmms(params: HarmonizeFseParams): HarmonizedResult {
  const netaData = params.data || params.rawPayload?.data || params.rawPayload || {};
  const siteInfo = params.siteInfo || netaData?.site_info || {};
  const analysis = params.analysisResult || params.evaluation || params.rawPayload?.analysisResult || {};

  // 1. Resolve Equipment Tag & Name
  const equipTag = String(
    params.equipmentId ||
    siteInfo.transformer_tag ||
    siteInfo.grounding_system_tag ||
    siteInfo.breaker_tag ||
    siteInfo.motor_tag ||
    siteInfo.generator_tag ||
    siteInfo.switchgear_tag ||
    siteInfo.cable_tag ||
    siteInfo.survey_tag ||
    siteInfo.switch_tag ||
    siteInfo.pv_array_tag ||
    siteInfo.evse_tag ||
    siteInfo.battery_bank_tag ||
    siteInfo.equipment_tag ||
    'EQ-01'
  ).trim();

  // Find matching equipment from master list
  const matchedEq = (params.existingEquipment || []).find(
    e => e.id.toLowerCase() === equipTag.toLowerCase()
  );

  const equipName = params.equipmentName || matchedEq?.name || siteInfo.transformer_name || `Thiết bị ${equipTag}`;
  const equipType = params.equipmentType || matchedEq?.type || siteInfo.transformer_type || 'Máy biến áp';

  // 2. Resolve Customer Code & Customer Name
  let customerName = String(
    params.customer ||
    matchedEq?.customer ||
    siteInfo.project_name ||
    siteInfo.customer_name ||
    'Công ty Cổ phần Năng lượng TEV'
  ).trim();

  let customerId = params.customerId || '';
  const customersList = params.existingCustomers || getRegisteredCustomers();

  if (!customerId) {
    // Try matching customer name
    const foundCust = customersList.find(
      c => c.name.toLowerCase() === customerName.toLowerCase() ||
           c.factories.some(f => f.toLowerCase() === customerName.toLowerCase())
    );
    if (foundCust) {
      customerId = foundCust.id;
      customerName = foundCust.name;
    } else {
      // Allocate next KHXXX code
      customerId = getNextCustomerId(customersList);
    }
  }

  // Ensure CustomerId conforms to KHXXX if needed
  if (!customerId.startsWith('KH')) {
    customerId = `KH${customerId.replace(/\D/g, '').padStart(3, '0')}`;
  }

  const factory = params.factory || matchedEq?.factory || siteInfo.location || siteInfo.substation || 'Trạm chính';

  // 3. Resolve FSE Engineer
  const fseName = String(
    params.fseName ||
    siteInfo.fse_name ||
    siteInfo.inspector ||
    'Võ Minh Tú (sgm1707@gmail.com)'
  ).trim();

  // 4. Resolve Test Date
  const rawDate = params.testDate || siteInfo.test_date || new Date().toISOString().split('T')[0];
  let formattedDisplayDate = rawDate;
  if (rawDate.includes('-')) {
    const parts = rawDate.split('-');
    if (parts.length === 3) formattedDisplayDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
  }

  // 5. Evaluate Status, Health Index & Failure Codes
  const overallStatus = String(analysis.overall_status || analysis.verdict || 'PASS').toUpperCase();
  let cmmsStatus = 'completed';
  let cmmsPriority = 'medium';
  let failureCode = 'PM-ROUTINE';
  let healthScore = 95;

  if (overallStatus.includes('FAIL') || overallStatus.includes('CRITICAL') || overallStatus.includes('IMMEDIATELY')) {
    cmmsStatus = 'in-progress';
    cmmsPriority = 'urgent';
    failureCode = 'ELEC-01'; // Electrical breakdown
    healthScore = analysis.health_index || analysis.health_score || 35;
  } else if (overallStatus.includes('INVESTIGATE') || overallStatus.includes('WARNING') || overallStatus.includes('SCHEDULED')) {
    cmmsStatus = 'in-progress';
    cmmsPriority = 'high';
    failureCode = 'THERM-01'; // Degraded insulation / overheating
    healthScore = analysis.health_index || analysis.health_score || 72;
  } else {
    cmmsStatus = 'completed';
    cmmsPriority = 'medium';
    failureCode = 'PM-ROUTINE';
    healthScore = analysis.health_index || analysis.health_score || 96;
  }

  // 6. Generate IDs
  const woId = getNextWorkOrderId(params.existingWorkOrders);
  const wpId = getNextWorkPermitId(params.existingWorkOrders);

  const dateDigits = rawDate.replace(/\D/g, '').slice(0, 8);
  const cleanTag = equipTag.replace(/[^a-zA-Z0-9]/g, '').slice(0, 8);
  const netaReportId = `REP-ATS-${dateDigits}-${cleanTag}`;

  // 7. Title & Technical Description
  const title = `[NETA ATS] Kiểm định & thử nghiệm ${equipName} (${equipTag})`;
  const statusReason = analysis.status_reason || analysis.status_reason_vi || 'Biên bản kiểm tra kỹ thuật hiện trường NETA ATS-2025';
  const description = `${statusReason}. Đánh giá tổng quan: ${overallStatus} (Chỉ số HI: ${healthScore}%). Kỹ sư FSE thực hiện: ${fseName}. Khách hàng: ${customerName} (${customerId}).`;

  const nowIso = new Date().toISOString();

  // 8. Canonical Work Order Object
  const workOrder: CmmsWorkOrder = {
    id: woId,
    workPermitId: wpId,
    title,
    description,
    equipmentId: equipTag,
    equipmentName: equipName,
    customerId,
    customer: customerName,
    factory,
    type: 'preventive',
    priority: cmmsPriority,
    status: cmmsStatus,
    isUnplanned: false,
    assignedTo: fseName,
    responsibleApprove: 'FSE Lead / Technical Manager',
    responsibleDo: 'Kỹ sư Thí nghiệm NETA Hiện trường',
    blockingRequired: cmmsPriority === 'urgent',
    dueDate: rawDate,
    createdAt: nowIso,
    updatedAt: nowIso,
    completedAt: cmmsStatus === 'completed' ? nowIso : null,
    pmFrequency: '6-months',
    failureCode,
    rootCause: analysis.recommendations ? (Array.isArray(analysis.recommendations) ? analysis.recommendations.join('; ') : analysis.recommendations) : '',
    actualTimeSpent: 3.5,
    estimatedTime: 4.0,
    laborCount: 2,
    laborCost: 1500000,
    partCost: 0,
    attachments: [netaReportId],
    usedMaterials: 'Hóa chất làm sạch cách điện, Thiết bị đo NETA ATS',
    downtimeStart: '',
    repairStart: rawDate,
    repairEnd: rawDate,
    restartTime: '',
    healthIndex: healthScore,
    netaReportId,
    inspectionSummary: statusReason,
    testStandard: params.type || 'ANSI/NETA ATS-2025',
    testDate: formattedDisplayDate,
    reportData: {
      netaData,
      analysisResult: analysis
    }
  };

  // 9. Canonical Report Object
  const report = {
    id: netaReportId,
    woId,
    equipmentId: equipTag,
    equipmentName: equipName,
    customer: customerName,
    factory,
    date: formattedDisplayDate,
    testDate: rawDate,
    inspector: fseName,
    status: healthScore >= 85 ? 'healthy' : healthScore >= 60 ? 'warning' : 'critical',
    notes: description,
    type: equipType,
    testStandard: params.type || 'ANSI/NETA ATS-2025',
    overallStatus,
    health: healthScore,
    details: {
      netaData,
      analysisResult: analysis
    }
  };

  // 10. Equipment Health Update
  const equipmentUpdate = {
    id: equipTag,
    health: healthScore,
    status: (healthScore >= 85 ? 'healthy' : healthScore >= 60 ? 'warning' : 'critical') as 'healthy' | 'warning' | 'critical',
    lastCheck: formattedDisplayDate
  };

  return {
    workOrder,
    report,
    equipmentUpdate
  };
}
