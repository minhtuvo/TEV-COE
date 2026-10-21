import { ProjectMasterRecord } from '../services/masterDirectoryService';

export type PmpHealthStatus = 'completed' | 'on_track' | 'at_risk' | 'critical_overdue' | 'planned';

export interface PmpProjectHealth {
  status: PmpHealthStatus;
  label: string;
  shortLabel: string;
  spi: number; // Schedule Performance Index (e.g. 1.05)
  plannedProgressPercent: number; // Expected % by reference date
  actualProgressPercent: number; // Actual KKS completion %
  varianceDays: number; // Planned End vs Reference Date
  isOverdue: boolean;
  isEndingSoon: boolean;
  durationDays: number;
  startDateFormatted: string;
  endDateFormatted: string;
  colorScheme: {
    badgeBg: string;
    badgeText: string;
    badgeBorder: string;
    dotColor: string;
    barGradient: string;
    solidBg: string;
    textColor: string;
    borderColor: string;
    cardBorderHighlight: string;
  };
}

/**
 * Standard reference date for PM calculations.
 * In production/2026 timeline: default to 2026-03-15 or current date.
 */
export const DEFAULT_PMP_REFERENCE_DATE = new Date('2026-03-15T00:00:00');

/**
 * Parse an ISO or YYYY-MM-DD date safely
 */
export function parseDateSafe(dateStr?: string, fallback = '2026-03-01'): Date {
  if (!dateStr) return new Date(fallback);
  try {
    const d = new Date(dateStr);
    return isNaN(d.getTime()) ? new Date(fallback) : d;
  } catch {
    return new Date(fallback);
  }
}

/**
 * Calculate PMP Schedule Performance & RAG Status according to PMI/PMP standard:
 * - Green (On Track): SPI >= 0.95 or actual progress meets planned timeline
 * - Amber (At Risk): 0.80 <= SPI < 0.95 or approaching deadline with high pending workload
 * - Red (Critical / Overdue): SPI < 0.80 or deadline exceeded with incomplete testing
 * - Blue (Planned / Baseline): Project not yet started (Reference date < Start date)
 * - Emerald / Slate (Completed): 100% equipment tested or marked complete
 */
export function calculatePmpProjectHealth(
  prj: ProjectMasterRecord,
  referenceDate: Date = DEFAULT_PMP_REFERENCE_DATE
): PmpProjectHealth {
  const rawStart = prj.planningStartDate || prj.startDate || '2026-03-01';
  const rawEnd = prj.planningEndDate || prj.endDate || '2026-03-25';

  const pStart = parseDateSafe(rawStart, '2026-03-01');
  const pEnd = parseDateSafe(rawEnd, '2026-03-25');

  // Duration in days
  const durationDays = Math.max(1, Math.round((pEnd.getTime() - pStart.getTime()) / (1000 * 60 * 60 * 24)) + 1);

  // KKS completion %
  const totalEq = (prj.equipmentList || []).length;
  const testedEq = (prj.equipmentList || []).filter((e) => e.status === 'đã_test').length;
  const actualProgressPercent =
    totalEq > 0
      ? Math.round((testedEq / totalEq) * 100)
      : prj.status === 'hoàn_thành'
      ? 100
      : prj.status === 'đang_thực_hiện'
      ? 50
      : 0;

  const nowMs = referenceDate.getTime();
  const startMs = pStart.getTime();
  const endMs = pEnd.getTime();

  // Days remaining until deadline
  const daysUntilDeadline = Math.round((endMs - nowMs) / (1000 * 60 * 60 * 24));
  const isOverdue = nowMs > endMs && actualProgressPercent < 100 && prj.status !== 'hoàn_thành';
  const isEndingSoon = daysUntilDeadline >= 0 && daysUntilDeadline <= 5 && actualProgressPercent < 80;

  // Expected progress calculation (Planned Value timeline)
  let plannedProgressPercent = 0;
  if (nowMs < startMs) {
    plannedProgressPercent = 0;
  } else if (nowMs >= endMs) {
    plannedProgressPercent = 100;
  } else {
    const elapsedDays = Math.max(0, Math.round((nowMs - startMs) / (1000 * 60 * 60 * 24)));
    plannedProgressPercent = Math.min(100, Math.max(1, Math.round((elapsedDays / durationDays) * 100)));
  }

  // Schedule Performance Index (SPI = Actual Progress / Planned Progress)
  let spi = 1.0;
  if (plannedProgressPercent > 0) {
    spi = Number((actualProgressPercent / plannedProgressPercent).toFixed(2));
  } else {
    spi = 1.0;
  }

  // Determine PMP Health Status
  let status: PmpHealthStatus = 'on_track';
  let label = 'Đúng hạn (On Track)';
  let shortLabel = 'On Track';

  if (prj.status === 'hoàn_thành' || (totalEq > 0 && testedEq === totalEq)) {
    status = 'completed';
    label = 'Hoàn tất (Completed)';
    shortLabel = 'Hoàn thành';
  } else if (isOverdue || (plannedProgressPercent === 100 && actualProgressPercent < 100)) {
    status = 'critical_overdue';
    label = 'Trễ hạn (Critical / Overdue)';
    shortLabel = 'Trễ hạn';
  } else if (nowMs < startMs && prj.status === 'chuẩn_bị') {
    status = 'planned';
    label = 'Kế hoạch cơ sở (Baseline)';
    shortLabel = 'Kế hoạch';
  } else if (isEndingSoon || spi < 0.85 || actualProgressPercent < plannedProgressPercent - 15) {
    status = 'at_risk';
    label = 'Cảnh báo trễ (At Risk)';
    shortLabel = 'Cảnh báo';
  } else {
    status = 'on_track';
    label = 'Đúng tiến độ (On Track)';
    shortLabel = 'Đúng hạn';
  }

  // PMP Standard Color Scheme
  const colorMap: Record<PmpHealthStatus, PmpProjectHealth['colorScheme']> = {
    completed: {
      badgeBg: 'bg-emerald-500/15',
      badgeText: 'text-emerald-700',
      badgeBorder: 'border-emerald-300',
      dotColor: 'bg-emerald-500',
      barGradient: 'from-emerald-600 via-teal-600 to-emerald-700',
      solidBg: 'bg-emerald-600',
      textColor: 'text-emerald-700',
      borderColor: 'border-emerald-300',
      cardBorderHighlight: 'hover:border-emerald-400'
    },
    on_track: {
      badgeBg: 'bg-emerald-500/15',
      badgeText: 'text-emerald-800',
      badgeBorder: 'border-emerald-400/60',
      dotColor: 'bg-emerald-500',
      barGradient: 'from-emerald-600 via-emerald-500 to-teal-600',
      solidBg: 'bg-emerald-600',
      textColor: 'text-emerald-800',
      borderColor: 'border-emerald-300',
      cardBorderHighlight: 'hover:border-emerald-400'
    },
    at_risk: {
      badgeBg: 'bg-amber-500/15',
      badgeText: 'text-amber-900',
      badgeBorder: 'border-amber-400',
      dotColor: 'bg-amber-500',
      barGradient: 'from-amber-500 via-amber-600 to-orange-600',
      solidBg: 'bg-amber-500',
      textColor: 'text-amber-900',
      borderColor: 'border-amber-300',
      cardBorderHighlight: 'hover:border-amber-400'
    },
    critical_overdue: {
      badgeBg: 'bg-rose-500/15',
      badgeText: 'text-rose-800',
      badgeBorder: 'border-rose-400',
      dotColor: 'bg-rose-600',
      barGradient: 'from-rose-600 via-red-600 to-rose-700',
      solidBg: 'bg-rose-600',
      textColor: 'text-rose-800',
      borderColor: 'border-rose-300',
      cardBorderHighlight: 'hover:border-rose-400'
    },
    planned: {
      badgeBg: 'bg-blue-500/15',
      badgeText: 'text-blue-800',
      badgeBorder: 'border-blue-300',
      dotColor: 'bg-blue-500',
      barGradient: 'from-blue-600 via-indigo-600 to-blue-700',
      solidBg: 'bg-blue-600',
      textColor: 'text-blue-800',
      borderColor: 'border-blue-300',
      cardBorderHighlight: 'hover:border-blue-400'
    }
  };

  return {
    status,
    label,
    shortLabel,
    spi,
    plannedProgressPercent,
    actualProgressPercent,
    varianceDays: daysUntilDeadline,
    isOverdue,
    isEndingSoon,
    durationDays,
    startDateFormatted: rawStart,
    endDateFormatted: rawEnd,
    colorScheme: colorMap[status]
  };
}
