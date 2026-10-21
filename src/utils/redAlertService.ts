/**
 * redAlertService.ts
 * Enterprise Diagnostic Engine for High-Voltage & Critical Fleet Asset Management
 * 
 * Condition triggers:
 * 1. Health Index (HI) < 50% (Ngưỡng tới hạn theo tiêu chuẩn NETA/IEEE)
 * 2. Overdue Maintenance > 30 Days (Quá hạn bảo trì định kỳ hoặc phiếu công việc quá 30 ngày)
 */

export interface RedAlertEquipment {
  equipment: any;
  id: string;
  name: string;
  customer: string;
  factory: string;
  location: string;
  type: string;
  criticality?: string;
  
  // Health Metrics
  health: number;
  isLowHealth: boolean; // health < 50%
  
  // Maintenance Metrics
  lastCheckDateStr: string;
  nextDueDateStr: string;
  cycleMonths: number;
  overdueDays: number;
  isOverdueMaintenance: boolean; // overdueDays > 30
  
  // Classification
  alertType: 'both' | 'low_health' | 'overdue_pm';
  reasons: string[];
  severity: 'critical' | 'high';
  
  // Linked Diagnostics & PM Details
  linkedWorkOrder?: any;
  diagnosticNotes?: string;
}

export interface RedAlertSummary {
  totalAlerts: number;
  lowHealthCount: number;
  overdueCount: number;
  bothCount: number;
  criticalUnitsCount: number;
  alerts: RedAlertEquipment[];
}

/**
 * Safely parse date from various formats:
 * - DD/MM/YYYY
 * - YYYY-MM-DD
 * - ISO string
 * - Timestamp
 */
export function parseDate(dateStr?: any): Date | null {
  if (!dateStr) return null;
  if (dateStr instanceof Date) return isNaN(dateStr.getTime()) ? null : dateStr;
  
  if (typeof dateStr === 'number') {
    const d = new Date(dateStr);
    return isNaN(d.getTime()) ? null : d;
  }
  
  const str = String(dateStr).trim();
  if (!str) return null;

  // Format: DD/MM/YYYY or DD-MM-YYYY
  const dmyMatch = str.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
  if (dmyMatch) {
    const day = parseInt(dmyMatch[1], 10);
    const month = parseInt(dmyMatch[2], 10) - 1;
    const year = parseInt(dmyMatch[3], 10);
    const d = new Date(year, month, day);
    return isNaN(d.getTime()) ? null : d;
  }

  // Format: YYYY-MM-DD or ISO
  const isoDate = new Date(str);
  if (!isNaN(isoDate.getTime())) {
    return isoDate;
  }

  return null;
}

/**
 * Format Date to standard Vietnamese display DD/MM/YYYY
 */
export function formatDate(date: Date | null): string {
  if (!date || isNaN(date.getTime())) return 'Chưa xác định';
  const d = String(date.getDate()).padStart(2, '0');
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const y = date.getFullYear();
  return `${d}/${m}/${y}`;
}

/**
 * Evaluates whether an equipment triggers the Critical Red Alert condition:
 * - Health Index < 50%
 * - OR Overdue Maintenance > 30 days
 */
export function evaluateEquipmentRedAlert(
  eq: any,
  workOrders: any[] = [],
  referenceDate: Date = new Date()
): RedAlertEquipment | null {
  if (!eq || !eq.id) return null;

  // 1. Health Index check
  const rawHealth = typeof eq.health === 'number' 
    ? eq.health 
    : typeof eq.healthIndex === 'number' 
      ? eq.healthIndex 
      : typeof eq.healthScore === 'number'
        ? eq.healthScore
        : 100;
  
  const health = Math.round(rawHealth);
  const isLowHealth = health < 50;

  // 2. Overdue Maintenance calculation
  const now = referenceDate;
  const cycleMonths = Number(eq.pmCycleMonths) || 6;
  
  // A. Check linked pending/active work orders (preventive or corrective)
  let maxWoOverdueDays = 0;
  let linkedWo: any = null;

  if (Array.isArray(workOrders) && workOrders.length > 0) {
    const eqIdClean = String(eq.id).toLowerCase().trim();
    const relatedOrders = workOrders.filter(wo => {
      if (!wo) return false;
      const rawId = wo.equipmentId;
      if (Array.isArray(rawId)) {
        return rawId.some(id => String(id).toLowerCase().trim() === eqIdClean);
      }
      return String(rawId || '').toLowerCase().trim() === eqIdClean;
    });

    for (const wo of relatedOrders) {
      if (wo.status === 'completed' || wo.status === 'cancelled') continue;
      if (wo.dueDate) {
        const woDue = parseDate(wo.dueDate);
        if (woDue && now > woDue) {
          const days = Math.floor((now.getTime() - woDue.getTime()) / (1000 * 60 * 60 * 24));
          if (days > maxWoOverdueDays) {
            maxWoOverdueDays = days;
            linkedWo = wo;
          }
        }
      }
    }
  }

  // B. Check equipment's own scheduled maintenance dueDate / nextMaintenance
  let maxDirectOverdueDays = 0;
  let nextDueDate: Date | null = null;
  const directDueStr = eq.nextMaintenance || eq.maintenanceDueDate || eq.dueDate;
  if (directDueStr) {
    const d = parseDate(directDueStr);
    if (d) {
      nextDueDate = d;
      if (now > d) {
        maxDirectOverdueDays = Math.floor((now.getTime() - d.getTime()) / (1000 * 60 * 60 * 24));
      }
    }
  }

  // C. Check equipment's lastCheck and maintenance interval (pmCycleMonths)
  let lastCheckDate: Date | null = null;
  let cycleOverdueDays = 0;
  const rawLastCheck = eq.lastCheck || eq.lastMaintenance || eq.lastInspectionDate;

  if (rawLastCheck) {
    lastCheckDate = parseDate(rawLastCheck);
    if (lastCheckDate) {
      // Calculate next scheduled PM date based on cycle
      const calculatedDue = new Date(lastCheckDate);
      calculatedDue.setMonth(calculatedDue.getMonth() + cycleMonths);
      
      if (!nextDueDate) {
        nextDueDate = calculatedDue;
      }
      
      if (now > calculatedDue) {
        cycleOverdueDays = Math.floor((now.getTime() - calculatedDue.getTime()) / (1000 * 60 * 60 * 24));
      }
    }
  }

  // D. Check manual/mock overdueMaintenanceDays if set
  const manualOverdueDays = Number(eq.overdueMaintenanceDays) || Number(eq.overdueDays) || 0;

  // Determine overall overdue days
  const overdueDays = Math.max(0, maxWoOverdueDays, maxDirectOverdueDays, cycleOverdueDays, manualOverdueDays);
  const isOverdueMaintenance = overdueDays > 30;

  // If neither condition is met, return null (equipment is not in Red Alert state)
  if (!isLowHealth && !isOverdueMaintenance) {
    return null;
  }

  // Build diagnostic reasons
  const reasons: string[] = [];
  if (isLowHealth) {
    reasons.push(`Health Index giảm xuống ${health}% (Dưới ngưỡng tới hạn 50%)`);
  }
  if (isOverdueMaintenance) {
    const dueStr = nextDueDate ? formatDate(nextDueDate) : 'N/A';
    reasons.push(`Quá hạn bảo trì ${overdueDays} ngày (Hạn chót: ${dueStr} - Chu kỳ ${cycleMonths} tháng)`);
    if (linkedWo) {
      reasons.push(`Phiếu CMMS chưa hoàn thành: ${linkedWo.id || linkedWo.title}`);
    }
  }

  const alertType: 'both' | 'low_health' | 'overdue_pm' = 
    isLowHealth && isOverdueMaintenance 
      ? 'both' 
      : isLowHealth 
        ? 'low_health' 
        : 'overdue_pm';

  return {
    equipment: eq,
    id: eq.id,
    name: eq.name || `Thiết bị ${eq.id}`,
    customer: eq.customer || 'Chưa gán',
    factory: eq.factory || 'Chưa gán',
    location: eq.location || 'Trạm biến áp / Phòng điện',
    type: eq.type || 'Thiết bị điện',
    criticality: eq.criticality || 'A',
    health,
    isLowHealth,
    lastCheckDateStr: lastCheckDate ? formatDate(lastCheckDate) : (eq.lastCheck || 'Chưa kiểm tra'),
    nextDueDateStr: nextDueDate ? formatDate(nextDueDate) : 'Đã quá hạn',
    cycleMonths,
    overdueDays,
    isOverdueMaintenance,
    alertType,
    reasons,
    severity: (alertType === 'both' || health < 35 || overdueDays > 60) ? 'critical' : 'high',
    linkedWorkOrder: linkedWo,
    diagnosticNotes: eq.notes || eq.diagnosticNotes || (isLowHealth ? 'Khuyến nghị: Đo cách điện IR & phân tích khí DGA ngay' : 'Khuyến nghị: Lập kế hoạch bảo dưỡng định kỳ')
  };
}

/**
 * Scans a fleet of equipment and produces a sorted Red Alert summary report
 */
export function getFleetRedAlerts(
  equipmentList: any[] = [],
  workOrders: any[] = [],
  filterCustomer: string = 'all',
  filterFactory: string = 'all',
  referenceDate: Date = new Date()
): RedAlertSummary {
  // Filter by customer & factory if selected
  const targetEquipment = equipmentList.filter(eq => {
    if (!eq) return false;
    const matchCust = filterCustomer === 'all' || eq.customer === filterCustomer;
    const matchFac = filterFactory === 'all' || eq.factory === filterFactory;
    return matchCust && matchFac;
  });

  const alerts: RedAlertEquipment[] = [];
  let lowHealthCount = 0;
  let overdueCount = 0;
  let bothCount = 0;
  let criticalUnitsCount = 0;

  for (const eq of targetEquipment) {
    const alert = evaluateEquipmentRedAlert(eq, workOrders, referenceDate);
    if (alert) {
      alerts.push(alert);
      if (alert.isLowHealth) lowHealthCount++;
      if (alert.isOverdueMaintenance) overdueCount++;
      if (alert.alertType === 'both') bothCount++;
      if (alert.severity === 'critical') criticalUnitsCount++;
    }
  }

  // Sort: 'both' first, then lowest health, then highest overdue days
  alerts.sort((a, b) => {
    if (a.alertType === 'both' && b.alertType !== 'both') return -1;
    if (b.alertType === 'both' && a.alertType !== 'both') return 1;
    if (a.health !== b.health) return a.health - b.health;
    return b.overdueDays - a.overdueDays;
  });

  return {
    totalAlerts: alerts.length,
    lowHealthCount,
    overdueCount,
    bothCount,
    criticalUnitsCount,
    alerts
  };
}
