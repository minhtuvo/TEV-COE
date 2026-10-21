import React, { useState, useMemo, useRef } from 'react';
import {
  Calendar,
  Clock,
  CheckCircle2,
  AlertCircle,
  AlertTriangle,
  Filter,
  Layers,
  ChevronDown,
  ChevronUp,
  Sparkles,
  Zap,
  TrendingUp,
  Activity,
  User,
  Building,
  Tag,
  ArrowRight,
  ShieldCheck,
  Flag,
  HelpCircle,
  BarChart3
} from 'lucide-react';
import { ProjectMasterRecord, ProjectEquipmentItem } from '../services/masterDirectoryService';
import {
  calculatePmpProjectHealth,
  PmpHealthStatus,
  PmpProjectHealth,
  DEFAULT_PMP_REFERENCE_DATE,
  parseDateSafe
} from '../utils/pmpStandard';

interface ProjectPortfolioGanttTimelineProps {
  projects: ProjectMasterRecord[];
  onSelectProject?: (project: ProjectMasterRecord) => void;
  onOpenManageEquipment?: (project: ProjectMasterRecord) => void;
  onOpenAddProject?: () => void;
  selectedProjectId?: string | null;
}

export type StandardFilter = 'all' | 'ats' | 'mts' | 'ecs';
export type PmpStatusFilter = 'all' | 'on_track' | 'at_risk' | 'critical_overdue' | 'planned' | 'completed';
export type TimeScaleMode = '3m' | '6m' | '12m';
export type ColorCodingMode = 'pmp' | 'neta';

export const ProjectPortfolioGanttTimeline: React.FC<ProjectPortfolioGanttTimelineProps> = ({
  projects,
  onSelectProject,
  onOpenManageEquipment,
  onOpenAddProject,
  selectedProjectId
}) => {
  const [standardFilter, setStandardFilter] = useState<StandardFilter>('all');
  const [pmpFilter, setPmpFilter] = useState<PmpStatusFilter>('all');
  const [colorMode, setColorMode] = useState<ColorCodingMode>('pmp');
  const [timeScale, setTimeScale] = useState<TimeScaleMode>('6m');
  const [isExpanded, setIsExpanded] = useState<boolean>(true);
  const [hoveredProject, setHoveredProject] = useState<ProjectMasterRecord | null>(null);
  const timelineScrollRef = useRef<HTMLDivElement>(null);

  // Reference Base Date for Timeline calculation (2026 local time)
  const currentYear = 2026;
  const today = useMemo(() => DEFAULT_PMP_REFERENCE_DATE, []);

  // Compute Timeline boundaries
  const timelineBounds = useMemo(() => {
    let startMonth = 0; // Jan
    let endMonth = 5; // Jun (6 months default)

    if (timeScale === '3m') {
      startMonth = 1; // Feb
      endMonth = 4; // May
    } else if (timeScale === '6m') {
      startMonth = 0; // Jan
      endMonth = 5; // Jun
    } else {
      startMonth = 0; // Jan
      endMonth = 11; // Dec
    }

    const startDate = new Date(currentYear, startMonth, 1);
    const endDate = new Date(currentYear, endMonth + 1, 0, 23, 59, 59);
    const totalDays = Math.max(1, Math.round((endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24)));

    // Generate month segments
    const months = [];
    for (let m = startMonth; m <= endMonth; m++) {
      const monthStart = new Date(currentYear, m, 1);
      const monthEnd = new Date(currentYear, m + 1, 0);
      const daysInMonth = monthEnd.getDate();
      months.push({
        monthIndex: m,
        name: `Tháng ${m + 1}`,
        year: currentYear,
        days: daysInMonth,
        startDate: monthStart,
        endDate: monthEnd
      });
    }

    return {
      startDate,
      endDate,
      totalDays,
      months
    };
  }, [timeScale, currentYear]);

  // Pre-calculate PMP Health for each project
  const projectHealthMap = useMemo(() => {
    const map = new Map<string, PmpProjectHealth>();
    projects.forEach(p => {
      map.set(p.id, calculatePmpProjectHealth(p, today));
    });
    return map;
  }, [projects, today]);

  // Filter projects by Standard and PMP Health status
  const filteredProjects = useMemo(() => {
    return projects.filter(p => {
      // Standard filter
      if (standardFilter !== 'all') {
        const std = p.standard || (p.name?.toLowerCase().includes('ecs') ? 'ecs' : p.name?.toLowerCase().includes('mts') ? 'mts' : 'ats');
        if (std !== standardFilter) return false;
      }
      // PMP Status filter
      if (pmpFilter !== 'all') {
        const health = projectHealthMap.get(p.id);
        if (health && health.status !== pmpFilter) return false;
      }
      return true;
    });
  }, [projects, standardFilter, pmpFilter, projectHealthMap]);

  // Portfolio KPIs according to PMP Standards
  const portfolioStats = useMemo(() => {
    const total = projects.length;
    let onTrackCount = 0;
    let atRiskCount = 0;
    let criticalCount = 0;
    let plannedCount = 0;
    let completedCount = 0;

    let totalEquipment = 0;
    let testedEquipment = 0;
    let countAts = 0;
    let countMts = 0;
    let countEcs = 0;

    projects.forEach(p => {
      const std = p.standard || (p.name?.toLowerCase().includes('ecs') ? 'ecs' : p.name?.toLowerCase().includes('mts') ? 'mts' : 'ats');
      if (std === 'ats') countAts++;
      else if (std === 'mts') countMts++;
      else if (std === 'ecs') countEcs++;

      const health = projectHealthMap.get(p.id);
      if (health) {
        if (health.status === 'completed') completedCount++;
        else if (health.status === 'on_track') onTrackCount++;
        else if (health.status === 'at_risk') atRiskCount++;
        else if (health.status === 'critical_overdue') criticalCount++;
        else if (health.status === 'planned') plannedCount++;
      }

      (p.equipmentList || []).forEach(eq => {
        totalEquipment++;
        if (eq.status === 'đã_test') testedEquipment++;
      });
    });

    const completionRate = totalEquipment > 0 ? Math.round((testedEquipment / totalEquipment) * 100) : 0;

    return {
      total,
      onTrackCount,
      atRiskCount,
      criticalCount,
      plannedCount,
      completedCount,
      countAts,
      countMts,
      countEcs,
      totalEquipment,
      testedEquipment,
      completionRate
    };
  }, [projects, projectHealthMap]);

  // Helper: compute bar left and width in %
  const computeBarPosition = (p: ProjectMasterRecord) => {
    const health = projectHealthMap.get(p.id) || calculatePmpProjectHealth(p, today);
    const pStart = parseDateSafe(health.startDateFormatted, '2026-03-01');
    const pEnd = parseDateSafe(health.endDateFormatted, '2026-03-25');

    const timelineStartMs = timelineBounds.startDate.getTime();
    const timelineEndMs = timelineBounds.endDate.getTime();
    const timelineSpanMs = Math.max(1, timelineEndMs - timelineStartMs);

    const startMs = Math.max(timelineStartMs, Math.min(timelineEndMs, pStart.getTime()));
    const endMs = Math.max(timelineStartMs, Math.min(timelineEndMs, pEnd.getTime()));

    const leftPercent = Math.max(0, Math.min(100, ((startMs - timelineStartMs) / timelineSpanMs) * 100));
    const widthPercent = Math.max(3, Math.min(100 - leftPercent, ((endMs - startMs) / timelineSpanMs) * 100));

    const totalEq = (p.equipmentList || []).length;
    const testedEq = (p.equipmentList || []).filter(e => e.status === 'đã_test').length;
    const std = p.standard || (p.name?.toLowerCase().includes('ecs') ? 'ecs' : p.name?.toLowerCase().includes('mts') ? 'mts' : 'ats');

    return {
      leftPercent,
      widthPercent,
      health,
      totalEq,
      testedEq,
      std
    };
  };

  // Today marker calculation
  const todayPosition = useMemo(() => {
    const timelineStartMs = timelineBounds.startDate.getTime();
    const timelineEndMs = timelineBounds.endDate.getTime();
    const timelineSpanMs = Math.max(1, timelineEndMs - timelineStartMs);
    const todayMs = today.getTime();

    if (todayMs < timelineStartMs || todayMs > timelineEndMs) return null;
    return ((todayMs - timelineStartMs) / timelineSpanMs) * 100;
  }, [today, timelineBounds]);

  return (
    <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden transition-all mb-4">
      {/* Top PM Executive Overview Bar - PMP Design Standard */}
      <div className="bg-gradient-to-r from-slate-900 via-slate-800 to-indigo-950 text-white p-4">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1">
            <div className="flex items-center gap-2.5 flex-wrap">
              <span className="px-2.5 py-1 rounded-lg bg-emerald-500/20 text-emerald-300 border border-emerald-500/40 font-mono text-[11px] font-bold uppercase tracking-wider flex items-center gap-1.5 shadow-2xs">
                <ShieldCheck size={13} className="text-emerald-400" />
                <span>PMP® Portfolio Cockpit (PMI Standard)</span>
              </span>
              <span className="px-2 py-0.5 rounded-md bg-white/10 text-slate-300 font-mono text-[10px] uppercase">
                EVM Schedule Variance
              </span>
              <h3 className="text-base sm:text-lg font-black tracking-tight flex items-center gap-2">
                <span>Timeline Tiến Độ Dự Án (Portfolio Gantt View)</span>
              </h3>
            </div>
            <p className="text-xs text-slate-300 max-w-3xl leading-relaxed">
              Hệ thống quản lý tiến độ danh mục dự án theo <strong>chuẩn PMP / PMI (RAG Status: Đỏ - Vàng - Xanh)</strong>. Giám sát chỉ số tiến độ kế hoạch cơ sở (Planned Baseline Start - End) đồng thời phân loại tiêu chuẩn thử nghiệm <strong>ANSI/NETA ATS, MTS và ECS</strong>.
            </p>
          </div>

          {/* Action & Toggle Controls */}
          <div className="flex items-center gap-2 shrink-0 self-end lg:self-center">
            {/* Color Mode Switcher */}
            <div className="flex items-center bg-black/40 p-1 rounded-xl border border-white/10 text-xs">
              <button
                type="button"
                onClick={() => setColorMode('pmp')}
                className={`px-2.5 py-1 rounded-lg font-bold flex items-center gap-1.5 transition-all ${
                  colorMode === 'pmp'
                    ? 'bg-emerald-600 text-white shadow-xs'
                    : 'text-slate-300 hover:text-white'
                }`}
                title="Màu sắc theo chuẩn PMP (RAG: On Track, At Risk, Critical Overdue)"
              >
                <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse" />
                <span>Chuẩn Màu PMP</span>
              </button>
              <button
                type="button"
                onClick={() => setColorMode('neta')}
                className={`px-2.5 py-1 rounded-lg font-bold flex items-center gap-1.5 transition-all ${
                  colorMode === 'neta'
                    ? 'bg-indigo-600 text-white shadow-xs'
                    : 'text-slate-300 hover:text-white'
                }`}
                title="Màu sắc theo chuẩn NETA (ATS: Xanh, MTS: Vàng, ECS: Ngọc)"
              >
                <span>Chuẩn NETA</span>
              </button>
            </div>

            {onOpenAddProject && (
              <button
                type="button"
                onClick={onOpenAddProject}
                className="px-3 py-1.5 bg-gradient-to-r from-amber-500 to-amber-600 hover:from-amber-600 hover:to-amber-700 text-slate-950 font-black text-xs rounded-xl shadow-xs flex items-center gap-1.5 active:scale-95 transition-all"
              >
                <span>+ Soạn Dự Án</span>
              </button>
            )}

            <button
              type="button"
              onClick={() => setIsExpanded(!isExpanded)}
              className="p-1.5 text-slate-300 hover:text-white hover:bg-white/10 rounded-xl transition-colors"
              title={isExpanded ? 'Thu gọn biểu đồ Gantt' : 'Mở rộng biểu đồ Gantt'}
            >
              {isExpanded ? <ChevronUp size={18} /> : <ChevronDown size={18} />}
            </button>
          </div>
        </div>

        {/* PMP 5-Pillar Health KPI Cards */}
        <div className="grid grid-cols-2 sm:grid-cols-5 gap-2.5 mt-4 pt-3 border-t border-slate-700/60">
          {/* Card 1: Total & Standards */}
          <div className="bg-white/5 rounded-xl p-2.5 border border-white/10 flex items-center justify-between">
            <div>
              <div className="text-[11px] text-slate-400 font-medium">Tổng Danh Mục</div>
              <div className="text-xl font-black text-white mt-0.5">{portfolioStats.total} <span className="text-xs font-normal text-slate-400">dự án</span></div>
            </div>
            <div className="flex flex-col items-end gap-0.5">
              <span className="text-[9px] font-mono px-1 py-0.2 rounded bg-blue-500/20 text-blue-300 border border-blue-400/30 font-bold">
                ATS: {portfolioStats.countAts}
              </span>
              <span className="text-[9px] font-mono px-1 py-0.2 rounded bg-amber-500/20 text-amber-300 border border-amber-400/30 font-bold">
                MTS: {portfolioStats.countMts}
              </span>
              <span className="text-[9px] font-mono px-1 py-0.2 rounded bg-teal-500/20 text-teal-300 border border-teal-400/30 font-bold">
                ECS: {portfolioStats.countEcs}
              </span>
            </div>
          </div>

          {/* Card 2: PMP Green - On Track */}
          <div className="bg-emerald-950/40 rounded-xl p-2.5 border border-emerald-500/30 flex items-center justify-between">
            <div>
              <div className="text-[11px] text-emerald-300 font-medium flex items-center gap-1">
                <span className="w-2 h-2 rounded-full bg-emerald-400" />
                <span>PMP: Đúng Hạn</span>
              </div>
              <div className="text-xl font-black text-emerald-400 mt-0.5">{portfolioStats.onTrackCount}</div>
            </div>
            <div className="w-8 h-8 rounded-lg bg-emerald-500/20 flex items-center justify-center text-emerald-300">
              <CheckCircle2 size={16} />
            </div>
          </div>

          {/* Card 3: PMP Amber - At Risk */}
          <div className="bg-amber-950/40 rounded-xl p-2.5 border border-amber-500/30 flex items-center justify-between">
            <div>
              <div className="text-[11px] text-amber-300 font-medium flex items-center gap-1">
                <span className="w-2 h-2 rounded-full bg-amber-400" />
                <span>PMP: Cảnh Báo Trễ</span>
              </div>
              <div className="text-xl font-black text-amber-400 mt-0.5">{portfolioStats.atRiskCount}</div>
            </div>
            <div className="w-8 h-8 rounded-lg bg-amber-500/20 flex items-center justify-center text-amber-300">
              <AlertTriangle size={16} />
            </div>
          </div>

          {/* Card 4: PMP Red - Critical / Overdue */}
          <div className="bg-rose-950/40 rounded-xl p-2.5 border border-rose-500/30 flex items-center justify-between">
            <div>
              <div className="text-[11px] text-rose-300 font-medium flex items-center gap-1">
                <span className="w-2 h-2 rounded-full bg-rose-500 animate-pulse" />
                <span>PMP: Trễ Hạn</span>
              </div>
              <div className="text-xl font-black text-rose-400 mt-0.5">{portfolioStats.criticalCount}</div>
            </div>
            <div className="w-8 h-8 rounded-lg bg-rose-500/20 flex items-center justify-center text-rose-300">
              <AlertCircle size={16} />
            </div>
          </div>

          {/* Card 5: Equipment KKS Earned Value Progress */}
          <div className="bg-white/5 rounded-xl p-2.5 border border-white/10 flex items-center justify-between">
            <div>
              <div className="text-[11px] text-indigo-300 font-medium">Tiến Độ KKS Thực Tế</div>
              <div className="text-lg font-black text-white mt-0.5">
                {portfolioStats.testedEquipment}/{portfolioStats.totalEquipment} <span className="text-xs text-indigo-300 font-semibold">({portfolioStats.completionRate}%)</span>
              </div>
            </div>
            <div className="w-8 h-8 rounded-lg bg-indigo-500/20 flex items-center justify-center text-indigo-300">
              <BarChart3 size={16} />
            </div>
          </div>
        </div>
      </div>

      {/* Collapsible Gantt Body */}
      {isExpanded && (
        <div className="p-3 sm:p-4 space-y-3 bg-slate-50/60">
          {/* Interactive Filters Bar */}
          <div className="flex flex-wrap items-center justify-between gap-2.5 pb-2 border-b border-slate-200">
            {/* Filter by PMP Health Status */}
            <div className="flex items-center gap-1 text-xs flex-wrap">
              <span className="text-slate-600 font-bold flex items-center gap-1 mr-1">
                <ShieldCheck size={13} className="text-emerald-600" /> Sức khỏe PMP:
              </span>
              <button
                type="button"
                onClick={() => setPmpFilter('all')}
                className={`px-2.5 py-1 rounded-lg font-bold transition-all text-xs ${
                  pmpFilter === 'all'
                    ? 'bg-slate-800 text-white shadow-xs'
                    : 'bg-white text-slate-600 border border-slate-200 hover:bg-slate-100'
                }`}
              >
                Tất cả ({projects.length})
              </button>

              <button
                type="button"
                onClick={() => setPmpFilter('on_track')}
                className={`px-2.5 py-1 rounded-lg font-bold transition-all text-xs flex items-center gap-1 ${
                  pmpFilter === 'on_track'
                    ? 'bg-emerald-600 text-white shadow-xs'
                    : 'bg-white text-emerald-800 border border-emerald-200 hover:bg-emerald-50'
                }`}
              >
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-500" />
                <span>Đúng hạn ({portfolioStats.onTrackCount})</span>
              </button>

              <button
                type="button"
                onClick={() => setPmpFilter('at_risk')}
                className={`px-2.5 py-1 rounded-lg font-bold transition-all text-xs flex items-center gap-1 ${
                  pmpFilter === 'at_risk'
                    ? 'bg-amber-500 text-slate-950 shadow-xs'
                    : 'bg-white text-amber-900 border border-amber-200 hover:bg-amber-50'
                }`}
              >
                <span className="w-1.5 h-1.5 rounded-full bg-amber-500" />
                <span>Cảnh báo ({portfolioStats.atRiskCount})</span>
              </button>

              <button
                type="button"
                onClick={() => setPmpFilter('critical_overdue')}
                className={`px-2.5 py-1 rounded-lg font-bold transition-all text-xs flex items-center gap-1 ${
                  pmpFilter === 'critical_overdue'
                    ? 'bg-rose-600 text-white shadow-xs'
                    : 'bg-white text-rose-800 border border-rose-200 hover:bg-rose-50'
                }`}
              >
                <span className="w-1.5 h-1.5 rounded-full bg-rose-500" />
                <span>Trễ hạn ({portfolioStats.criticalCount})</span>
              </button>

              <button
                type="button"
                onClick={() => setPmpFilter('completed')}
                className={`px-2.5 py-1 rounded-lg font-bold transition-all text-xs flex items-center gap-1 ${
                  pmpFilter === 'completed'
                    ? 'bg-teal-700 text-white shadow-xs'
                    : 'bg-white text-teal-800 border border-teal-200 hover:bg-teal-50'
                }`}
              >
                <span>Hoàn tất ({portfolioStats.completedCount})</span>
              </button>
            </div>

            {/* Filter by Standard (ATS, MTS, ECS) & Time Scale */}
            <div className="flex items-center gap-2 flex-wrap">
              {/* Standard Filter */}
              <div className="flex items-center bg-white p-0.5 rounded-lg border border-slate-300 text-xs font-bold">
                {(['all', 'ats', 'mts', 'ecs'] as StandardFilter[]).map(std => {
                  const isActive = standardFilter === std;
                  const labels: Record<StandardFilter, string> = {
                    all: 'Tất cả tiêu chuẩn',
                    ats: 'ATS',
                    mts: 'MTS',
                    ecs: 'ECS'
                  };
                  return (
                    <button
                      key={std}
                      type="button"
                      onClick={() => setStandardFilter(std)}
                      className={`px-2 py-0.5 rounded transition-all ${
                        isActive
                          ? std === 'ecs'
                            ? 'bg-teal-600 text-white'
                            : std === 'mts'
                            ? 'bg-amber-500 text-slate-950'
                            : std === 'ats'
                            ? 'bg-blue-600 text-white'
                            : 'bg-slate-700 text-white'
                          : 'text-slate-600 hover:bg-slate-100'
                      }`}
                    >
                      {labels[std]}
                    </button>
                  );
                })}
              </div>

              {/* TimeScale Toggle */}
              <div className="flex items-center bg-white p-0.5 rounded-lg border border-slate-300 text-xs font-bold">
                <button
                  type="button"
                  onClick={() => setTimeScale('3m')}
                  className={`px-2 py-0.5 rounded ${timeScale === '3m' ? 'bg-indigo-600 text-white' : 'text-slate-600 hover:bg-slate-100'}`}
                >
                  3 Tháng
                </button>
                <button
                  type="button"
                  onClick={() => setTimeScale('6m')}
                  className={`px-2 py-0.5 rounded ${timeScale === '6m' ? 'bg-indigo-600 text-white' : 'text-slate-600 hover:bg-slate-100'}`}
                >
                  6 Tháng
                </button>
                <button
                  type="button"
                  onClick={() => setTimeScale('12m')}
                  className={`px-2 py-0.5 rounded ${timeScale === '12m' ? 'bg-indigo-600 text-white' : 'text-slate-600 hover:bg-slate-100'}`}
                >
                  1 Năm
                </button>
              </div>
            </div>
          </div>

          {/* GANTT CHART VISUAL CANVAS */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-2xs overflow-hidden">
            <div className="overflow-x-auto" ref={timelineScrollRef}>
              <div className="min-w-[900px]">
                {/* Timeline Header Row */}
                <div className="grid grid-cols-12 border-b border-slate-200 bg-slate-100/90 text-xs text-slate-700 font-bold">
                  {/* Left Column: Project Info Header */}
                  <div className="col-span-4 p-2.5 border-r border-slate-200 flex items-center justify-between">
                    <span>Dự Án / PMP Health / FSE</span>
                    <span className="text-[10px] text-slate-500 font-mono">Planning Range</span>
                  </div>

                  {/* Right Column: Dynamic Month Intervals */}
                  <div className="col-span-8 relative flex">
                    {timelineBounds.months.map((m, idx) => (
                      <div
                        key={idx}
                        className="flex-1 p-2 text-center border-r border-slate-200 last:border-r-0"
                      >
                        <div className="font-bold text-slate-800">{m.name}</div>
                        <div className="text-[10px] text-slate-400 font-normal">{m.days} ngày</div>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Projects Gantt Rows */}
                <div className="divide-y divide-slate-100">
                  {filteredProjects.length > 0 ? (
                    filteredProjects.map((prj) => {
                      const pos = computeBarPosition(prj);
                      const isSelected = selectedProjectId === prj.id;
                      const health = pos.health;

                      // Standard Badge Styling
                      const standardBadge =
                        pos.std === 'ecs' ? (
                          <span className="px-1.5 py-0.2 rounded font-mono font-bold text-[9px] bg-teal-100 text-teal-800 border border-teal-300">
                            ECS
                          </span>
                        ) : pos.std === 'mts' ? (
                          <span className="px-1.5 py-0.2 rounded font-mono font-bold text-[9px] bg-amber-100 text-amber-900 border border-amber-300">
                            MTS
                          </span>
                        ) : (
                          <span className="px-1.5 py-0.2 rounded font-mono font-bold text-[9px] bg-blue-100 text-blue-800 border border-blue-300">
                            ATS
                          </span>
                        );

                      // Bar Color Decision: PMP Health vs NETA standard
                      let barGradient = health.colorScheme.barGradient;
                      if (colorMode === 'neta') {
                        barGradient =
                          pos.std === 'ecs'
                            ? 'from-teal-600 via-emerald-600 to-teal-700'
                            : pos.std === 'mts'
                            ? 'from-amber-500 via-amber-600 to-amber-700'
                            : 'from-blue-600 via-indigo-600 to-indigo-700';
                      }

                      return (
                        <div
                          key={prj.id}
                          className={`grid grid-cols-12 hover:bg-slate-50 transition-colors group ${
                            isSelected ? 'bg-indigo-50/50 ring-1 ring-inset ring-indigo-400' : ''
                          }`}
                        >
                          {/* Project Summary Cell */}
                          <div className="col-span-4 p-2.5 border-r border-slate-200 space-y-1">
                            <div className="flex items-center justify-between gap-1.5">
                              <div className="flex items-center gap-1.5">
                                <span className="font-mono font-black text-slate-800 text-xs">
                                  {prj.id}
                                </span>
                                {standardBadge}
                              </div>

                              {/* PMP Standard RAG Health Badge */}
                              <span
                                className={`text-[10px] px-2 py-0.5 rounded-full font-bold border flex items-center gap-1 ${health.colorScheme.badgeBg} ${health.colorScheme.badgeText} ${health.colorScheme.badgeBorder}`}
                                title={`Chỉ số tiến độ SPI: ${health.spi}. Kế hoạch: ${health.plannedProgressPercent}%, Thực tế: ${health.actualProgressPercent}%`}
                              >
                                <span className={`w-1.5 h-1.5 rounded-full ${health.colorScheme.dotColor}`} />
                                <span>{health.shortLabel}</span>
                              </span>
                            </div>

                            <div
                              onClick={() => onSelectProject && onSelectProject(prj)}
                              className="font-bold text-slate-900 text-xs hover:text-indigo-600 cursor-pointer line-clamp-1"
                              title={prj.name}
                            >
                              {prj.name}
                            </div>

                            <div className="text-[11px] text-slate-500 flex items-center justify-between">
                              <span className="truncate max-w-[170px]" title={prj.customerName}>
                                {prj.customerName}
                              </span>
                              <span className="font-mono font-semibold text-slate-600 text-[10px]">
                                {health.startDateFormatted} → {health.endDateFormatted}
                              </span>
                            </div>
                          </div>

                          {/* Gantt Timeline Bar Cell */}
                          <div className="col-span-8 relative py-3 px-2 flex items-center">
                            {/* Background Month Grid Guides */}
                            <div className="absolute inset-0 flex pointer-events-none">
                              {timelineBounds.months.map((_, mIdx) => (
                                <div
                                  key={mIdx}
                                  className="flex-1 border-r border-slate-100 last:border-r-0"
                                />
                              ))}
                            </div>

                            {/* Today Line Indicator */}
                            {todayPosition !== null && (
                              <div
                                style={{ left: `${todayPosition}%` }}
                                className="absolute top-0 bottom-0 w-0.5 bg-rose-500 z-10 pointer-events-none"
                              >
                                <span className="absolute -top-1 -translate-x-1/2 bg-rose-500 text-white text-[8px] font-bold px-1 rounded-full uppercase tracking-tighter shadow-2xs whitespace-nowrap">
                                  Hôm nay
                                </span>
                              </div>
                            )}

                            {/* PMP Planned Baseline Track (Outer Frame) */}
                            <div
                              style={{
                                left: `${pos.leftPercent}%`,
                                width: `${pos.widthPercent}%`
                              }}
                              className="absolute h-8 rounded-lg border border-dashed border-slate-300 bg-slate-100/60 pointer-events-none"
                              title="PMP Baseline Schedule Window"
                            />

                            {/* Active Gantt Bar (PMP Color Formatted) */}
                            <div
                              style={{
                                left: `${pos.leftPercent}%`,
                                width: `${pos.widthPercent}%`
                              }}
                              onClick={() => onSelectProject && onSelectProject(prj)}
                              onMouseEnter={() => setHoveredProject(prj)}
                              onMouseLeave={() => setHoveredProject(null)}
                              className={`relative h-7 rounded-lg bg-gradient-to-r ${barGradient} text-white shadow-xs cursor-pointer flex items-center px-2 justify-between overflow-hidden group/bar hover:ring-2 hover:ring-indigo-400 hover:scale-[1.01] transition-all z-0`}
                            >
                              {/* Earned Value / Tested Equipment Fill */}
                              <div
                                style={{ width: `${health.actualProgressPercent}%` }}
                                className="absolute top-0 left-0 bottom-0 bg-white/20 pointer-events-none"
                              />

                              {/* Inside Bar Label */}
                              <div className="relative z-10 flex items-center gap-1.5 min-w-0">
                                <span className="font-bold text-[11px] truncate drop-shadow-xs">
                                  {prj.name}
                                </span>
                              </div>

                              <div className="relative z-10 flex items-center gap-2 shrink-0 text-[10px] font-mono">
                                <span className="bg-black/30 px-1.5 py-0.2 rounded font-bold">
                                  {health.durationDays} ngày
                                </span>
                                {pos.totalEq > 0 && (
                                  <span className="bg-white/25 px-1.5 py-0.2 rounded font-black hidden sm:inline-block">
                                    {health.actualProgressPercent}% ({pos.testedEq}/{pos.totalEq})
                                  </span>
                                )}
                              </div>
                            </div>
                          </div>
                        </div>
                      );
                    })
                  ) : (
                    <div className="p-8 text-center text-slate-500 space-y-2">
                      <Calendar size={28} className="mx-auto text-slate-300" />
                      <p className="text-xs font-semibold">
                        Không có dự án nào khớp với bộ lọc PMP / tiêu chuẩn hiện tại.
                      </p>
                    </div>
                  )}
                </div>
              </div>
            </div>

            {/* Gantt Timeline Footer: PMP & NETA Official Legend */}
            <div className="p-3 bg-slate-50 border-t border-slate-200 space-y-2 px-4">
              {/* PMP Standard Legend Strip */}
              <div className="flex flex-wrap items-center justify-between gap-3 text-xs text-slate-600">
                <div className="flex items-center gap-3.5 flex-wrap text-[11px]">
                  <span className="font-extrabold text-slate-800 flex items-center gap-1">
                    <ShieldCheck size={13} className="text-emerald-600" />
                    <span>Màu chuẩn PMP (RAG Status):</span>
                  </span>

                  <span className="inline-flex items-center gap-1.5">
                    <span className="w-2.5 h-2.5 rounded-full bg-emerald-500" />
                    <span><strong>Xanh lá (On Track)</strong>: SPI ≥ 0.95 (Đúng tiến độ)</span>
                  </span>

                  <span className="inline-flex items-center gap-1.5">
                    <span className="w-2.5 h-2.5 rounded-full bg-amber-500" />
                    <span><strong>Vàng cam (At Risk)</strong>: Cảnh báo trễ / Gần hạn chót</span>
                  </span>

                  <span className="inline-flex items-center gap-1.5">
                    <span className="w-2.5 h-2.5 rounded-full bg-rose-500 animate-pulse" />
                    <span><strong>Đỏ cờ (Critical)</strong>: Quá deadline / Trễ nghiêm trọng</span>
                  </span>

                  <span className="inline-flex items-center gap-1.5">
                    <span className="w-2.5 h-2.5 rounded-full bg-blue-500" />
                    <span><strong>Xanh dương (Planned)</strong>: Kế hoạch cơ sở</span>
                  </span>

                  <span className="inline-flex items-center gap-1.5">
                    <span className="w-2.5 h-2.5 rounded-full bg-teal-700" />
                    <span><strong>Xanh ngọc sẫm</strong>: Đã đóng / Hoàn thành</span>
                  </span>
                </div>

                <div className="text-[11px] text-slate-500 font-medium">
                  Hiển thị <strong>{filteredProjects.length}</strong> / <strong>{projects.length}</strong> dự án trên trục thời gian năm {currentYear}
                </div>
              </div>

              {/* NETA Standards Indicator Strip */}
              <div className="flex flex-wrap items-center gap-3 pt-1 border-t border-slate-200 text-[10px] text-slate-500">
                <span className="font-bold text-slate-700">Mã tiêu chuẩn NETA:</span>
                <span className="inline-flex items-center gap-1 font-mono">
                  <span className="w-2 h-2 rounded-sm bg-blue-600" />
                  <strong>ATS</strong>: Acceptance Testing (Nghiệm thu đóng điện)
                </span>
                <span className="inline-flex items-center gap-1 font-mono">
                  <span className="w-2 h-2 rounded-sm bg-amber-500" />
                  <strong>MTS</strong>: Maintenance Testing (Bảo dưỡng định kỳ)
                </span>
                <span className="inline-flex items-center gap-1 font-mono">
                  <span className="w-2 h-2 rounded-sm bg-teal-600" />
                  <strong>ECS</strong>: Electrical Commissioning (Vận hành chạy thử)
                </span>
                <span className="text-slate-300">|</span>
                <span className="inline-flex items-center gap-1">
                  <span className="w-2 h-2 rounded-full bg-rose-500" />
                  <span>Đường mốc "Hôm nay" (Today Marker: {DEFAULT_PMP_REFERENCE_DATE.toISOString().split('T')[0]})</span>
                </span>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
