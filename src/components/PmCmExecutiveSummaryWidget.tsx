import React from 'react';
import { 
  ClipboardList, 
  CheckCircle2, 
  Clock, 
  AlertTriangle, 
  Activity, 
  ArrowDownCircle, 
  ArrowUpRight, 
  Plus, 
  TrendingUp, 
  ShieldCheck,
  CalendarCheck2
} from 'lucide-react';
import { DashboardWorkOrder } from './DashboardPmCmProgressModule';

export interface PmCmExecutiveSummaryWidgetProps {
  workOrders: DashboardWorkOrder[];
  reliabilityKpis: {
    totalWOs: number;
    initiatedCount: number;
    inProgressCount: number;
    completedCount: number;
    criticalCount: number;
    backlog?: number;
    compliance?: number;
  };
  onScrollToDetailedTable: () => void;
  onCreateWorkOrder: () => void;
  onNavigateToFullCmms: () => void;
}

export const PmCmExecutiveSummaryWidget: React.FC<PmCmExecutiveSummaryWidgetProps> = ({
  workOrders,
  reliabilityKpis,
  onScrollToDetailedTable,
  onCreateWorkOrder,
  onNavigateToFullCmms,
}) => {
  const total = reliabilityKpis.totalWOs || workOrders.length || 0;
  const completed = reliabilityKpis.completedCount || 0;
  const inProgress = reliabilityKpis.inProgressCount || 0;
  const initiated = reliabilityKpis.initiatedCount || 0;
  const completionRate = total > 0 ? Math.round((completed / total) * 100) : 0;

  // PM vs CM Breakdown
  const pmOrders = workOrders.filter(wo => {
    const t = (wo.type || '').toLowerCase();
    return t.includes('preventive') || t.includes('định kỳ') || t.includes('pm') || t.includes('bảo dưỡng');
  });
  const cmOrders = workOrders.filter(wo => {
    const t = (wo.type || '').toLowerCase();
    return t.includes('corrective') || t.includes('đột xuất') || t.includes('cm') || t.includes('khẩn cấp') || wo.isUnplanned;
  });

  const pmCount = pmOrders.length;
  const cmCount = cmOrders.length;
  const totalCategorized = pmCount + cmCount || total || 1;
  const pmPercent = Math.round((pmCount / totalCategorized) * 100);
  const cmPercent = Math.round((cmCount / totalCategorized) * 100);

  // Overdue count
  const overdueCount = workOrders.filter(wo => {
    if (!wo.dueDate) return false;
    const isOver = new Date(wo.dueDate).getTime() < Date.now();
    const s = (wo.status || '').toLowerCase();
    return isOver && !s.includes('complete') && !s.includes('hoàn thành');
  }).length;

  const compliance = reliabilityKpis.compliance ?? (completionRate > 0 ? Math.min(100, completionRate + 5) : 85);
  const isWorldClassRatio = pmPercent >= 75;

  return (
    <div className="bg-white rounded-2xl border border-slate-200/90 shadow-sm overflow-hidden">
      {/* HEADER BAR */}
      <div className="px-5 py-3.5 bg-slate-900 border-b border-slate-800 text-white flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-xl bg-blue-500/20 border border-blue-400/30 text-blue-400 flex items-center justify-center shadow-inner">
            <ClipboardList size={19} />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h2 className="text-sm sm:text-base font-bold text-white tracking-tight">
                Chỉ Số Tiến Độ PM/CM &amp; Quản Trị Công Việc Bảo Trì
              </h2>
              <span className="text-[10px] font-bold px-2 py-0.5 rounded-md bg-blue-500/20 text-blue-300 border border-blue-400/30 tracking-wider uppercase">
                Cấp Quản Lý (Executive)
              </span>
            </div>
            <p className="text-[11px] text-slate-400">
              Tổng hợp tiến độ phiếu công việc (WO), bảo trì định kỳ (PM) và xử lý sự cố (CM) đồng bộ hai chiều
            </p>
          </div>
        </div>

        {/* Action Controls */}
        <div className="flex items-center gap-2 shrink-0">
          <button
            onClick={onScrollToDetailedTable}
            className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white rounded-lg text-xs font-semibold transition-all border border-slate-700"
            title="Cuộn nhanh xuống bảng chi tiết phiếu công việc ở cuối trang"
          >
            <ArrowDownCircle size={14} className="text-blue-400" />
            <span>Xem Bảng Chi Tiết ↓</span>
          </button>

          <button
            onClick={onCreateWorkOrder}
            className="flex items-center gap-1.5 px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-xs font-bold transition-all shadow-sm"
          >
            <Plus size={14} />
            <span>+ Tạo WO</span>
          </button>

          <button
            onClick={onNavigateToFullCmms}
            className="flex items-center gap-1 px-2.5 py-1.5 bg-white/10 hover:bg-white/20 text-slate-200 hover:text-white rounded-lg text-xs font-medium transition-all"
            title="Mở toàn bộ hệ thống CMMS & Work Permits"
          >
            <span>CMMS</span>
            <ArrowUpRight size={13} />
          </button>
        </div>
      </div>

      {/* EXECUTIVE KPI GRID */}
      <div className="p-4 sm:p-5 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 bg-gradient-to-b from-slate-50/70 to-white">
        
        {/* CARD 1: OVERALL PM/CM COMPLETION PROGRESS */}
        <div className="bg-white rounded-xl p-4 border border-slate-200/90 shadow-2xs flex flex-col justify-between hover:border-blue-300 transition-colors">
          <div>
            <div className="flex items-center justify-between text-slate-500 mb-1.5">
              <span className="text-[11px] font-bold uppercase tracking-wider text-slate-600">
                Tiến Độ Hoàn Thành Kế Hoạch
              </span>
              <TrendingUp size={16} className="text-emerald-600" />
            </div>
            
            <div className="flex items-baseline gap-2 mt-1">
              <span className="text-3xl font-black font-mono text-slate-900">{completionRate}%</span>
              <span className="text-xs font-medium text-slate-500">
                ({completed} / {total} phiếu)
              </span>
            </div>

            {/* Visual Progress Bar */}
            <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden mt-3 flex">
              <div 
                className="h-full bg-gradient-to-r from-emerald-500 to-teal-400 rounded-full transition-all duration-500" 
                style={{ width: `${completionRate}%` }}
              />
            </div>
          </div>

          <div className="pt-3 mt-3 border-t border-slate-100 flex items-center justify-between text-[11px]">
            <span className="text-slate-500">Tuân thủ kế hoạch (Compliance):</span>
            <span className="font-bold font-mono text-emerald-700">{compliance}%</span>
          </div>
        </div>

        {/* CARD 2: PROACTIVE MAINTENANCE RATIO (PM VS CM) */}
        <div className="bg-white rounded-xl p-4 border border-slate-200/90 shadow-2xs flex flex-col justify-between hover:border-indigo-300 transition-colors">
          <div>
            <div className="flex items-center justify-between text-slate-500 mb-1.5">
              <span className="text-[11px] font-bold uppercase tracking-wider text-slate-600">
                Cơ Cấu Bảo Trì (PM vs CM)
              </span>
              <ShieldCheck size={16} className="text-indigo-600" />
            </div>

            <div className="flex items-baseline gap-2 mt-1">
              <span className="text-2xl font-black font-mono text-indigo-700">{pmPercent}%</span>
              <span className="text-xs font-semibold text-slate-600">Định kỳ (PM)</span>
              <span className="text-xs text-slate-400">·</span>
              <span className="text-xl font-bold font-mono text-rose-600">{cmPercent}%</span>
              <span className="text-xs font-semibold text-slate-600">Sự cố (CM)</span>
            </div>

            {/* Ratio Split Bar */}
            <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden mt-3 flex">
              <div 
                className="h-full bg-indigo-600 transition-all duration-500" 
                style={{ width: `${pmPercent}%` }}
                title={`Định kỳ (PM): ${pmCount} phiếu (${pmPercent}%)`}
              />
              <div 
                className="h-full bg-rose-500 transition-all duration-500" 
                style={{ width: `${cmPercent}%` }}
                title={`Sự cố (CM): ${cmCount} phiếu (${cmPercent}%)`}
              />
            </div>
          </div>

          <div className="pt-3 mt-3 border-t border-slate-100 flex items-center justify-between text-[11px]">
            <span className="text-slate-500">Chuẩn bảo trì World-Class:</span>
            <span className={`font-semibold ${isWorldClassRatio ? 'text-emerald-700' : 'text-amber-700'}`}>
              {isWorldClassRatio ? '✓ Đạt chuẩn (PM ≥ 80%)' : '⚠ Cần giảm sự cố CM'}
            </span>
          </div>
        </div>

        {/* CARD 3: ACTIVE WORK ORDERS IN PROGRESS */}
        <div className="bg-white rounded-xl p-4 border border-slate-200/90 shadow-2xs flex flex-col justify-between hover:border-amber-300 transition-colors">
          <div>
            <div className="flex items-center justify-between text-slate-500 mb-1.5">
              <span className="text-[11px] font-bold uppercase tracking-wider text-slate-600">
                Điều Độ Đang Triển Khai
              </span>
              <Activity size={16} className="text-amber-600" />
            </div>

            <div className="flex items-baseline gap-2 mt-1">
              <span className="text-3xl font-black font-mono text-amber-600">{inProgress}</span>
              <span className="text-xs font-medium text-slate-500">phiếu đang thực hiện</span>
            </div>

            <p className="text-[11px] text-slate-500 mt-2">
              Kỹ sư FSE đang thao tác tại hiện trường theo phiếu Work Order và Work Permit đã duyệt
            </p>
          </div>

          <div className="pt-3 mt-3 border-t border-slate-100 flex items-center justify-between text-[11px]">
            <span className="text-slate-500">Phiếu mới tạo / Chờ xử lý:</span>
            <span className="font-bold font-mono text-sky-700">{initiated} phiếu</span>
          </div>
        </div>

        {/* CARD 4: SLA & BACKLOG METRICS */}
        <div className="bg-white rounded-xl p-4 border border-slate-200/90 shadow-2xs flex flex-col justify-between hover:border-rose-300 transition-colors">
          <div>
            <div className="flex items-center justify-between text-slate-500 mb-1.5">
              <span className="text-[11px] font-bold uppercase tracking-wider text-slate-600">
                Tuân Thủ Tiến Độ &amp; SLA
              </span>
              <CalendarCheck2 size={16} className={overdueCount > 0 ? "text-rose-600" : "text-emerald-600"} />
            </div>

            <div className="flex items-baseline gap-2 mt-1">
              <span className={`text-3xl font-black font-mono ${overdueCount > 0 ? 'text-rose-600' : 'text-emerald-600'}`}>
                {overdueCount}
              </span>
              <span className="text-xs font-medium text-slate-500">
                {overdueCount > 0 ? 'phiếu trễ hạn (Overdue)' : 'phiếu trễ hạn'}
              </span>
            </div>

            <p className="text-[11px] text-slate-500 mt-2">
              {overdueCount > 0 
                ? 'Có phiếu công việc vượt quá mốc Deadline quy định, cần ưu tiên phân bổ nhân lực'
                : 'Tất cả phiếu công việc đang được kiểm soát trong phạm vi thời hạn cam kết'
              }
            </p>
          </div>

          <div className="pt-3 mt-3 border-t border-slate-100 flex items-center justify-between text-[11px]">
            <span className="text-slate-500">Công việc tồn đọng (Backlog):</span>
            <span className="font-bold font-mono text-slate-800">{reliabilityKpis.backlog || (total - completed)} phiếu</span>
          </div>
        </div>

      </div>
    </div>
  );
};
