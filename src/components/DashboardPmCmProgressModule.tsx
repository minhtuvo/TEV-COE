import React from 'react';
import { 
  ClipboardList, 
  CheckCircle, 
  Clock, 
  AlertTriangle, 
  Activity, 
  Plus, 
  RefreshCw, 
  ArrowUpRight, 
  Search, 
  Eye, 
  Check, 
  Play,
  Factory,
  TrendingUp,
  ShieldCheck
} from 'lucide-react';

export interface DashboardWorkOrder {
  id: string;
  title: string;
  description?: string;
  equipmentId?: string | string[];
  equipmentName?: string;
  customer?: string;
  customerId?: string;
  type?: string;
  priority?: string;
  status?: string;
  assignedTo?: string;
  responsibleDo?: string;
  dueDate?: string;
  workPermitId?: string;
  isUnplanned?: boolean;
}

export interface DashboardPmCmProgressModuleProps {
  workOrders: DashboardWorkOrder[];
  dashboardFilteredWorkOrders: DashboardWorkOrder[];
  dashboardWoFilter: 'all' | 'corrective' | 'preventive' | 'in-progress' | 'initiated' | 'overdue';
  setDashboardWoFilter: (filter: 'all' | 'corrective' | 'preventive' | 'in-progress' | 'initiated' | 'overdue') => void;
  dashboardWoSearch: string;
  setDashboardWoSearch: (query: string) => void;
  reliabilityKpis: {
    totalWOs: number;
    initiatedCount: number;
    inProgressCount: number;
    completedCount: number;
    criticalCount: number;
    backlog?: number;
    compliance?: number;
  };
  selectedCustomer: string;
  isGoogleConnected: boolean;
  isSyncing: boolean;
  onFetchFromSheets: () => void;
  onOpenWorkOrderDetail: (wo: DashboardWorkOrder) => void;
  onQuickUpdateStatus: (woId: string, newStatus: string) => void;
  onCreateWorkOrder: () => void;
  onNavigateToFullCmms: () => void;
  onNavigateToFieldStation: () => void;
}

export const DashboardPmCmProgressModule: React.FC<DashboardPmCmProgressModuleProps> = ({
  workOrders,
  dashboardFilteredWorkOrders,
  dashboardWoFilter,
  setDashboardWoFilter,
  dashboardWoSearch,
  setDashboardWoSearch,
  reliabilityKpis,
  selectedCustomer,
  isGoogleConnected,
  isSyncing,
  onFetchFromSheets,
  onOpenWorkOrderDetail,
  onQuickUpdateStatus,
  onCreateWorkOrder,
  onNavigateToFullCmms,
  onNavigateToFieldStation
}) => {
  const totalWOs = reliabilityKpis.totalWOs || workOrders.length || 0;
  const completedCount = reliabilityKpis.completedCount || 0;
  const inProgressCount = reliabilityKpis.inProgressCount || 0;
  const initiatedCount = reliabilityKpis.initiatedCount || 0;
  const completionPercent = totalWOs > 0 ? Math.round((completedCount / totalWOs) * 100) : 0;

  const pmCount = workOrders.filter(wo => {
    const t = (wo.type || '').toLowerCase();
    return t.includes('preventive') || t.includes('định kỳ') || t.includes('pm') || t.includes('bảo dưỡng');
  }).length;

  const cmCount = workOrders.filter(wo => {
    const t = (wo.type || '').toLowerCase();
    return t.includes('corrective') || t.includes('đột xuất') || t.includes('cm') || t.includes('khẩn cấp') || wo.isUnplanned;
  }).length;

  return (
    <div id="cmms-work-orders-table" className="bg-white rounded-2xl border border-slate-200/90 shadow-md overflow-hidden transition-all">
      {/* 1. EXECUTIVE HEADER & PROGRESS BANNER */}
      <div className="p-5 bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1.5">
            <div className="flex items-center gap-2.5">
              <div className="w-10 h-10 rounded-xl bg-blue-500/20 border border-blue-400/30 text-blue-400 flex items-center justify-center shadow-inner">
                <ClipboardList size={22} />
              </div>
              <div>
                <div className="flex items-center gap-2.5">
                  <h2 className="text-lg sm:text-xl font-bold tracking-tight text-white">
                    Tiến Độ PM/CM &amp; Quản Lý Bảo Trì Hiện Trường
                  </h2>
                  <span className="px-2.5 py-0.5 rounded-full text-xs font-bold bg-blue-500/20 text-blue-300 border border-blue-400/30 uppercase tracking-wider">
                    CMMS Trực Quan
                  </span>
                </div>
                <p className="text-xs text-slate-300">
                  Theo dõi tiến độ xử lý sự cố (CM), bảo dưỡng phòng ngừa (PM) và đồng bộ Google Sheets hai chiều
                </p>
              </div>
            </div>
          </div>

          {/* Action Toolbar */}
          <div className="flex flex-wrap items-center gap-2 shrink-0">
            <button
              onClick={onNavigateToFullCmms}
              className="px-3.5 py-2 bg-white/10 hover:bg-white/20 text-white rounded-xl text-xs font-semibold transition-all border border-white/10 flex items-center gap-1.5 shadow-xs"
              title="Chuyển đến màn hình Quản lý CMMS chuyên sâu"
            >
              <span>Phân hệ CMMS đầy đủ</span>
              <ArrowUpRight size={14} className="text-blue-300" />
            </button>

            <button
              onClick={onCreateWorkOrder}
              className="px-3.5 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-xs font-bold transition-all shadow-sm hover:shadow-blue-500/20 flex items-center gap-1.5"
            >
              <Plus size={15} />
              <span>+ Tạo phiếu WO/WR</span>
            </button>

            <button
              onClick={onFetchFromSheets}
              disabled={!isGoogleConnected || isSyncing}
              className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 rounded-xl text-xs font-medium transition-colors flex items-center gap-1.5 disabled:opacity-50"
              title="Tải lại phiếu từ Google Sheet TEV Service Flatform"
            >
              <RefreshCw size={13} className={isSyncing ? 'animate-spin text-blue-400' : 'text-slate-400'} />
              <span className="hidden sm:inline">Đồng bộ Sheets</span>
            </button>
          </div>
        </div>

        {/* 2. PROGRESS BAR & QUICK STAT CARDS */}
        <div className="mt-5 pt-4 border-t border-white/10 space-y-3">
          {/* Progress Bar Row */}
          <div className="bg-white/5 border border-white/10 rounded-xl p-3.5 space-y-2">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-1 text-xs">
              <div className="flex items-center gap-2 font-semibold">
                <TrendingUp size={15} className="text-emerald-400" />
                <span>Tiến độ hoàn thành kế hoạch bảo trì:</span>
                <span className="text-emerald-400 font-bold font-mono text-sm">{completionPercent}%</span>
                <span className="text-slate-400 font-normal">({completedCount} / {totalWOs} phiếu hoàn tất)</span>
              </div>
              <div className="flex items-center gap-3 text-[11px] text-slate-300">
                <span className="flex items-center gap-1">
                  <span className="w-2 h-2 rounded-full bg-emerald-400"></span> Hoàn tất ({completedCount})
                </span>
                <span className="flex items-center gap-1">
                  <span className="w-2 h-2 rounded-full bg-amber-400"></span> Đang làm ({inProgressCount})
                </span>
                <span className="flex items-center gap-1">
                  <span className="w-2 h-2 rounded-full bg-sky-400"></span> Chờ xử lý ({initiatedCount})
                </span>
              </div>
            </div>

            {/* Visual Bar */}
            <div className="w-full h-2.5 bg-slate-800 rounded-full overflow-hidden flex shadow-inner">
              <div 
                className="h-full bg-gradient-to-r from-emerald-500 to-teal-400 transition-all duration-500 rounded-full" 
                style={{ width: `${completionPercent}%` }}
                title={`Hoàn thành: ${completionPercent}%`}
              />
              <div 
                className="h-full bg-amber-400 transition-all duration-500" 
                style={{ width: `${totalWOs > 0 ? (inProgressCount / totalWOs) * 100 : 0}%` }}
                title={`Đang làm: ${inProgressCount}`}
              />
              <div 
                className="h-full bg-sky-400 transition-all duration-500" 
                style={{ width: `${totalWOs > 0 ? (initiatedCount / totalWOs) * 100 : 0}%` }}
                title={`Mới tạo/Chờ: ${initiatedCount}`}
              />
            </div>
          </div>

          {/* Quick Metrics KPI Pills */}
          <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2">
            <button
              onClick={() => setDashboardWoFilter('all')}
              className={`p-2.5 rounded-xl border text-left transition-all ${
                dashboardWoFilter === 'all'
                  ? 'bg-blue-600/30 border-blue-400/80 ring-1 ring-blue-400'
                  : 'bg-white/5 border-white/10 hover:bg-white/10'
              }`}
            >
              <p className="text-[10px] uppercase font-bold text-slate-400 tracking-wider">Tất cả phiếu</p>
              <p className="text-lg font-bold text-white mt-0.5">{totalWOs}</p>
            </button>

            <button
              onClick={() => setDashboardWoFilter('initiated')}
              className={`p-2.5 rounded-xl border text-left transition-all ${
                dashboardWoFilter === 'initiated'
                  ? 'bg-sky-500/30 border-sky-400/80 ring-1 ring-sky-400'
                  : 'bg-white/5 border-white/10 hover:bg-white/10'
              }`}
            >
              <p className="text-[10px] uppercase font-bold text-sky-300 tracking-wider flex items-center gap-1">
                <Clock size={11} /> Mới tạo / Chờ
              </p>
              <p className="text-lg font-bold text-sky-200 mt-0.5">{initiatedCount}</p>
            </button>

            <button
              onClick={() => setDashboardWoFilter('in-progress')}
              className={`p-2.5 rounded-xl border text-left transition-all ${
                dashboardWoFilter === 'in-progress'
                  ? 'bg-amber-500/30 border-amber-400/80 ring-1 ring-amber-400'
                  : 'bg-white/5 border-white/10 hover:bg-white/10'
              }`}
            >
              <p className="text-[10px] uppercase font-bold text-amber-300 tracking-wider flex items-center gap-1">
                <Activity size={11} /> Đang xử lý
              </p>
              <p className="text-lg font-bold text-amber-200 mt-0.5">{inProgressCount}</p>
            </button>

            <button
              onClick={() => setDashboardWoFilter('preventive')}
              className={`p-2.5 rounded-xl border text-left transition-all ${
                dashboardWoFilter === 'preventive'
                  ? 'bg-emerald-500/30 border-emerald-400/80 ring-1 ring-emerald-400'
                  : 'bg-white/5 border-white/10 hover:bg-white/10'
              }`}
            >
              <p className="text-[10px] uppercase font-bold text-emerald-300 tracking-wider flex items-center gap-1">
                <CheckCircle size={11} /> Định kỳ (PM)
              </p>
              <p className="text-lg font-bold text-emerald-200 mt-0.5">{pmCount}</p>
            </button>

            <button
              onClick={() => setDashboardWoFilter('corrective')}
              className={`p-2.5 rounded-xl border text-left transition-all ${
                dashboardWoFilter === 'corrective'
                  ? 'bg-rose-500/30 border-rose-400/80 ring-1 ring-rose-400'
                  : 'bg-white/5 border-white/10 hover:bg-white/10'
              }`}
            >
              <p className="text-[10px] uppercase font-bold text-rose-300 tracking-wider flex items-center gap-1">
                <AlertTriangle size={11} /> Sự cố (CM)
              </p>
              <p className="text-lg font-bold text-rose-200 mt-0.5">{cmCount}</p>
            </button>

            <button
              onClick={() => setDashboardWoFilter('overdue')}
              className={`p-2.5 rounded-xl border text-left transition-all ${
                dashboardWoFilter === 'overdue'
                  ? 'bg-purple-500/30 border-purple-400/80 ring-1 ring-purple-400'
                  : 'bg-white/5 border-white/10 hover:bg-white/10'
              }`}
            >
              <p className="text-[10px] uppercase font-bold text-purple-300 tracking-wider flex items-center gap-1">
                <ShieldCheck size={11} /> Quá hạn
              </p>
              <p className="text-lg font-bold text-purple-200 mt-0.5">
                {workOrders.filter(wo => {
                  if (!wo.dueDate) return false;
                  const isOver = new Date(wo.dueDate).getTime() < Date.now();
                  const s = (wo.status || '').toLowerCase();
                  return isOver && !s.includes('complete') && !s.includes('hoàn thành');
                }).length}
              </p>
            </button>
          </div>
        </div>
      </div>

      {/* 3. FILTER, SEARCH & CONTROLS */}
      <div className="p-4 bg-slate-50 border-b border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-3">
        <div className="flex flex-wrap items-center gap-1.5 w-full sm:w-auto">
          <span className="text-xs text-slate-500 font-semibold mr-1">Lọc nhanh:</span>
          {[
            { id: 'all', label: 'Tất cả' },
            { id: 'initiated', label: 'Mới tạo' },
            { id: 'in-progress', label: 'Đang làm' },
            { id: 'corrective', label: 'Sửa đột xuất (CM)' },
            { id: 'preventive', label: 'Bảo trì định kỳ (PM)' },
            { id: 'overdue', label: 'Quá hạn' },
          ].map(f => (
            <button
              key={f.id}
              onClick={() => setDashboardWoFilter(f.id as any)}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${
                dashboardWoFilter === f.id
                  ? 'bg-slate-900 text-white shadow-xs'
                  : 'bg-white text-slate-700 hover:bg-slate-200/80 border border-slate-200'
              }`}
            >
              {f.label}
            </button>
          ))}
        </div>

        <div className="relative w-full sm:w-80">
          <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input
            type="text"
            placeholder="Tìm theo mã WO, thiết bị, khách hàng..."
            value={dashboardWoSearch}
            onChange={(e) => setDashboardWoSearch(e.target.value)}
            className="w-full text-xs pl-9 pr-7 py-2 border border-slate-200 rounded-xl bg-white focus:outline-none focus:ring-2 focus:ring-blue-500 font-medium placeholder:text-slate-400 shadow-2xs"
          />
          {dashboardWoSearch && (
            <button
              onClick={() => setDashboardWoSearch('')}
              className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 text-xs font-bold"
            >
              ✕
            </button>
          )}
        </div>
      </div>

      {/* 4. WORK ORDERS TABLE */}
      <div className="overflow-x-auto max-h-[460px]">
        <table className="w-full text-left border-collapse">
          <thead className="sticky top-0 bg-slate-100/95 backdrop-blur-xs border-b border-slate-200 text-[11px] uppercase tracking-wider text-slate-600 font-bold z-10">
            <tr>
              <th className="p-3">Mã WO &amp; Tiêu đề</th>
              <th className="p-3">Thiết bị &amp; Khách hàng</th>
              <th className="p-3">Loại hình</th>
              <th className="p-3">Ưu tiên</th>
              <th className="p-3">Trạng thái</th>
              <th className="p-3">Phụ trách</th>
              <th className="p-3">Hạn chót</th>
              <th className="p-3 text-right">Thao tác</th>
            </tr>
          </thead>
          <tbody className="text-xs divide-y divide-slate-100">
            {dashboardFilteredWorkOrders.length > 0 ? (
              dashboardFilteredWorkOrders.map((wo) => {
                const eqStr = Array.isArray(wo.equipmentId) ? wo.equipmentId.join(', ') : (wo.equipmentId || 'Chưa gán');
                const isHigh = wo.priority === 'Critical' || wo.priority === 'High' || wo.priority === 'Khẩn cấp';
                const statusLower = (wo.status || '').toLowerCase();
                const isDone = statusLower.includes('complete') || statusLower.includes('hoàn thành');
                const isInProg = statusLower.includes('progress') || statusLower.includes('đang');
                
                return (
                  <tr key={wo.id} className="hover:bg-slate-50/80 transition-colors">
                    <td className="p-3">
                      <div className="font-mono font-bold text-blue-700 flex items-center gap-1.5">
                        <span>{wo.id}</span>
                        {wo.workPermitId && (
                          <span className="text-[10px] font-normal text-slate-400 font-sans">
                            ({wo.workPermitId})
                          </span>
                        )}
                      </div>
                      <div className="font-semibold text-slate-900 mt-0.5 line-clamp-1" title={wo.title}>
                        {wo.title}
                      </div>
                      {wo.description && (
                        <div className="text-[11px] text-slate-500 line-clamp-1 mt-0.5" title={wo.description}>
                          {wo.description}
                        </div>
                      )}
                    </td>
                    <td className="p-3">
                      <div className="font-semibold text-slate-800 flex items-center gap-1">
                        <span className="px-1.5 py-0.5 rounded bg-slate-100 border border-slate-200 text-slate-700 font-mono text-[10px]">
                          {eqStr}
                        </span>
                        {wo.equipmentName && (
                          <span className="text-[11px] text-slate-600 line-clamp-1">
                            {wo.equipmentName}
                          </span>
                        )}
                      </div>
                      <div className="text-[11px] text-slate-500 mt-0.5 flex items-center gap-1">
                        <Factory size={11} className="text-slate-400 shrink-0" />
                        <span className="font-medium text-slate-700 truncate">{wo.customer || 'Chưa phân loại'}</span>
                      </div>
                    </td>
                    <td className="p-3">
                      <span className={`px-2 py-0.5 rounded text-[10px] font-semibold ${
                        (wo.type || '').toLowerCase().includes('corrective') || (wo.type || '').toLowerCase().includes('đột xuất') || wo.isUnplanned
                          ? 'bg-rose-50 text-rose-700 border border-rose-200'
                          : (wo.type || '').toLowerCase().includes('preventive') || (wo.type || '').toLowerCase().includes('định kỳ')
                          ? 'bg-blue-50 text-blue-700 border border-blue-200'
                          : 'bg-slate-100 text-slate-700 border border-slate-200'
                      }`}>
                        {wo.type || (wo.isUnplanned ? 'Sửa đột xuất (CM)' : 'Định kỳ (PM)')}
                      </span>
                    </td>
                    <td className="p-3">
                      <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                        isHigh
                          ? 'bg-rose-100 text-rose-800 border border-rose-300'
                          : wo.priority === 'Medium' || wo.priority === 'Trung bình'
                          ? 'bg-amber-100 text-amber-800 border border-amber-300'
                          : 'bg-slate-100 text-slate-700 border border-slate-200'
                      }`}>
                        {wo.priority || 'Medium'}
                      </span>
                    </td>
                    <td className="p-3">
                      <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold ${
                        isDone
                          ? 'bg-emerald-100 text-emerald-800 border border-emerald-200'
                          : isInProg
                          ? 'bg-amber-100 text-amber-800 border border-amber-200 animate-pulse'
                          : 'bg-sky-100 text-sky-800 border border-sky-200'
                      }`}>
                        <span className={`w-1.5 h-1.5 rounded-full ${
                          isDone ? 'bg-emerald-600' : isInProg ? 'bg-amber-600' : 'bg-sky-600'
                        }`}></span>
                        {isDone ? 'Hoàn thành' : isInProg ? 'Đang xử lý' : (wo.status || 'Mới tạo')}
                      </span>
                    </td>
                    <td className="p-3 text-slate-600">
                      <div className="text-[11px] font-medium text-slate-800">{wo.assignedTo || 'Chưa phân công'}</div>
                      {wo.responsibleDo && <div className="text-[10px] text-slate-400">{wo.responsibleDo}</div>}
                    </td>
                    <td className="p-3 font-mono text-slate-600 text-[11px]">
                      {wo.dueDate ? (
                        <span className={new Date(wo.dueDate).getTime() < Date.now() && !isDone ? 'text-rose-600 font-bold' : ''}>
                          {wo.dueDate}
                        </span>
                      ) : (
                        <span className="text-slate-400">-</span>
                      )}
                    </td>
                    <td className="p-3 text-right">
                      <div className="flex items-center justify-end gap-1.5">
                        {!isDone && (
                          <button
                            onClick={() => onQuickUpdateStatus(wo.id, isInProg ? 'completed' : 'in-progress')}
                            className="px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 text-[11px] font-semibold rounded-lg transition-colors flex items-center gap-1"
                            title={isInProg ? 'Đánh dấu hoàn thành' : 'Chuyển sang Đang xử lý'}
                          >
                            {isInProg ? (
                              <>
                                <Check size={12} className="text-emerald-600" />
                                <span>Hoàn thành</span>
                              </>
                            ) : (
                              <>
                                <Play size={11} className="text-blue-600" />
                                <span>Bắt đầu</span>
                              </>
                            )}
                          </button>
                        )}
                        <button
                          onClick={() => onOpenWorkOrderDetail(wo)}
                          className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                          title="Xem chi tiết phiếu trong CMMS"
                        >
                          <Eye size={16} />
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })
            ) : (
              <tr>
                <td colSpan={8} className="p-8 text-center text-slate-500">
                  <div className="flex flex-col items-center justify-center gap-2">
                    <ClipboardList size={32} className="text-slate-300" />
                    <p className="text-sm font-semibold text-slate-700">Không tìm thấy phiếu CMMS nào phù hợp với bộ lọc</p>
                    <p className="text-xs text-slate-400">
                      Thử chọn lại khách hàng &quot;Tất cả&quot; hoặc xóa từ khóa tìm kiếm
                    </p>
                  </div>
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* 5. FOOTER */}
      <div className="p-3.5 border-t border-slate-100 bg-slate-50 flex flex-wrap items-center justify-between text-xs text-slate-500">
        <div className="flex items-center gap-2">
          <span>Hiển thị <strong>{dashboardFilteredWorkOrders.length}</strong> / <strong>{workOrders.length}</strong> phiếu công việc.</span>
          {selectedCustomer !== 'all' && (
            <span className="text-blue-600 font-semibold">(Đang lọc theo: {selectedCustomer})</span>
          )}
        </div>
        <div className="flex items-center gap-3">
          <button
            onClick={onNavigateToFieldStation}
            className="text-blue-600 hover:text-blue-800 font-semibold transition-colors flex items-center gap-1"
          >
            <span>Trạm nhập liệu &amp; Thí nghiệm hiện trường</span>
            <ArrowUpRight size={13} />
          </button>
        </div>
      </div>
    </div>
  );
};
