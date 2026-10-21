import React, { useState, useMemo } from 'react';
import {
  AlertOctagon,
  AlertTriangle,
  Flame,
  Activity,
  Clock,
  ChevronDown,
  ChevronUp,
  Wrench,
  FileText,
  Search,
  ExternalLink,
  ShieldAlert,
  Zap,
  Factory,
  MapPin,
  CheckCircle2,
  Copy,
  Check,
  Filter,
  Download,
  Eye,
  History,
  X
} from 'lucide-react';
import { RedAlertSummary, RedAlertEquipment } from '../utils/redAlertService';

interface Props {
  summary: RedAlertSummary;
  onFilterToEquipment?: (equipmentId: string) => void;
  onFilterToAllRedAlerts?: () => void;
  onClearRedAlertFilter?: () => void;
  isFilterActive?: boolean;
  onOpenWorkOrderModal?: (equipment: any) => void;
  onNavigateToInspection?: (equipment: any) => void;
  onSelectRiskDetail?: (equipmentId: string) => void;
  onViewEquipmentProfile?: (equipment: any) => void;
}

export const EquipmentRedAlertSection: React.FC<Props> = ({
  summary,
  onFilterToEquipment,
  onFilterToAllRedAlerts,
  onClearRedAlertFilter,
  isFilterActive = false,
  onOpenWorkOrderModal,
  onNavigateToInspection,
  onSelectRiskDetail,
  onViewEquipmentProfile
}) => {
  const [isExpanded, setIsExpanded] = useState<boolean>(true);
  const [activeTab, setActiveTab] = useState<'all' | 'low_health' | 'overdue_pm' | 'both'>('all');
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [copied, setCopied] = useState<boolean>(false);

  // Filter alerts based on active tab and search query
  const filteredAlerts = useMemo(() => {
    return summary.alerts.filter((item) => {
      // Tab filter
      if (activeTab === 'low_health' && !item.isLowHealth) return false;
      if (activeTab === 'overdue_pm' && !item.isOverdueMaintenance) return false;
      if (activeTab === 'both' && item.alertType !== 'both') return false;

      // Search filter
      if (searchTerm.trim()) {
        const q = searchTerm.toLowerCase().trim();
        const matchId = item.id.toLowerCase().includes(q);
        const matchName = item.name.toLowerCase().includes(q);
        const matchFactory = item.factory.toLowerCase().includes(q);
        const matchLocation = item.location.toLowerCase().includes(q);
        return matchId || matchName || matchFactory || matchLocation;
      }
      return true;
    });
  }, [summary.alerts, activeTab, searchTerm]);

  // Copy urgent alert summary to clipboard
  const handleCopyAlerts = () => {
    const lines = [
      `=== BÁO CÁO CẢNH BÁO ĐỎ - MODULE QUẢN LÝ THIẾT BỊ (TEV FLEET) ===`,
      `Thời gian xuất: ${new Date().toLocaleString('vi-VN')}`,
      `Tiêu chuẩn: NETA MTS & IEEE 493 (Health Index < 50% | Quá hạn bảo trì > 30 ngày)`,
      `Tổng số thiết bị phát hiện: ${summary.totalAlerts}`,
      `- Thiết bị có Health Index < 50%: ${summary.lowHealthCount}`,
      `- Thiết bị quá hạn bảo trì > 30 ngày: ${summary.overdueCount}`,
      `- Thiết bị nguy cơ kép (vừa HI < 50% vừa quá hạn > 30d): ${summary.bothCount}`,
      ``,
      ...summary.alerts.map((a, i) => 
        `${i + 1}. [${a.id}] ${a.name} (${a.factory} - ${a.location})\n   - HI: ${a.health}% | Quá hạn: ${a.overdueDays} ngày (Chu kỳ: ${a.pmCycleMonths} tháng | Lần cuối: ${a.lastCheck})\n   - Lý do: ${a.reasons.join(' | ')}`
      )
    ];
    navigator.clipboard.writeText(lines.join('\n'));
    setCopied(true);
    setTimeout(() => setCopied(false), 2500);
  };

  // Export CSV
  const handleExportCSV = () => {
    const headers = ['Mã Thiết Bị', 'Tên Thiết Bị', 'Khách Hàng', 'Nhà Máy', 'Vị Trí', 'Loại', 'Health Index (%)', 'Số Ngày Quá Hạn', 'Kiểm Tra Lần Cuối', 'Hạn Kế Tiếp', 'Lý Do Cảnh Báo'];
    const rows = summary.alerts.map(a => [
      `"${a.id}"`,
      `"${a.name.replace(/"/g, '""')}"`,
      `"${a.customer}"`,
      `"${a.factory}"`,
      `"${a.location.replace(/"/g, '""')}"`,
      `"${a.type}"`,
      a.health,
      a.overdueDays,
      `"${a.lastCheck}"`,
      `"${a.dueDate}"`,
      `"${a.reasons.join('; ').replace(/"/g, '""')}"`
    ]);
    const csvContent = 'data:text/csv;charset=utf-8,\uFEFF' + [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', `Canh_Bao_Do_Thiet_Bi_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // If there are 0 alerts across the fleet
  if (summary.totalAlerts === 0) {
    return (
      <div className="bg-emerald-50/70 border border-emerald-200/90 rounded-xl p-4 flex items-center justify-between shadow-xs">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-lg bg-emerald-100 flex items-center justify-center text-emerald-600">
            <CheckCircle2 size={22} />
          </div>
          <div>
            <h3 className="font-bold text-emerald-950 text-sm">Hệ Thống Thiết Bị Đạt Chuẩn Vận Hành An Toàn</h3>
            <p className="text-xs text-emerald-700 mt-0.5">
              Tất cả thiết bị đều có Health Index &ge; 50% và không có thiết bị nào quá hạn bảo trì quá 30 ngày (Tiêu chuẩn NETA MTS / IEEE 493).
            </p>
          </div>
        </div>
        <span className="text-xs font-semibold text-emerald-800 bg-emerald-100/80 px-3 py-1 rounded-md">
          Chỉ số an toàn 100%
        </span>
      </div>
    );
  }

  return (
    <div className="bg-rose-50/30 border-2 border-rose-500 rounded-xl shadow-md overflow-hidden transition-all duration-200">
      {/* SECTION HEADER BAR */}
      <div className="bg-gradient-to-r from-rose-900 via-rose-800 to-red-950 text-white p-4 sm:p-5">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="flex items-start sm:items-center gap-3">
            <div className="relative shrink-0">
              <div className="w-11 h-11 rounded-lg bg-rose-600/90 border border-rose-400/50 flex items-center justify-center shadow-inner">
                <ShieldAlert className="text-white" size={24} />
              </div>
              <span className="absolute -top-1 -right-1 flex h-3.5 w-3.5">
                <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-rose-400 opacity-75"></span>
                <span className="relative inline-flex rounded-full h-3.5 w-3.5 bg-rose-500 border-2 border-rose-900"></span>
              </span>
            </div>

            <div>
              <div className="flex flex-wrap items-center gap-2">
                <span className="px-2 py-0.5 rounded text-[11px] font-black uppercase tracking-wider bg-rose-500 text-white shadow-xs">
                  🚨 CẢNH BÁO ĐỎ CẤP 1
                </span>
                <span className="text-xs font-medium text-rose-200/90 hidden sm:inline">
                  NETA MTS / IEEE 493 Standard
                </span>
                {isFilterActive && (
                  <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-bold bg-amber-400 text-slate-900 shadow-xs animate-pulse">
                    <Filter size={11} /> Đang lọc bảng dưới theo Cảnh báo đỏ
                  </span>
                )}
              </div>
              <h3 className="text-base sm:text-lg font-bold text-white tracking-tight mt-0.5 flex items-center gap-2">
                Phát Hiện {summary.totalAlerts} Thiết Bị Có Health Index &lt; 50% Hoặc Quá Hạn Bảo Trì &gt; 30 Ngày
              </h3>
              <p className="text-xs text-rose-200/80 mt-0.5">
                Cần kích hoạt phiếu công tác CMMS, kiểm tra thử nghiệm thực địa NETA hoặc đánh giá ma trận rủi ro ngay lập tức.
              </p>
            </div>
          </div>

          {/* Quick Metrics Badges & Action Buttons */}
          <div className="flex flex-wrap items-center gap-2 self-end lg:self-center">
            {/* Direct filter toggle on equipment table */}
            {onFilterToAllRedAlerts && (
              <button
                onClick={isFilterActive ? onClearRedAlertFilter : onFilterToAllRedAlerts}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition-all shadow-sm ${
                  isFilterActive
                    ? 'bg-amber-400 text-slate-950 hover:bg-amber-300 ring-2 ring-amber-200'
                    : 'bg-rose-600 hover:bg-rose-500 text-white border border-rose-400/40'
                }`}
                title={isFilterActive ? "Hủy lọc bảng thiết bị" : "Lọc danh sách bảng bên dưới chỉ hiển thị thiết bị cảnh báo đỏ"}
              >
                {isFilterActive ? (
                  <>
                    <X size={13} className="text-slate-950" />
                    <span>Hủy lọc bảng ({summary.totalAlerts})</span>
                  </>
                ) : (
                  <>
                    <Filter size={13} />
                    <span>Lọc danh sách bên dưới ({summary.totalAlerts})</span>
                  </>
                )}
              </button>
            )}

            <button
              onClick={handleExportCSV}
              className="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-rose-950/60 hover:bg-rose-900/80 text-rose-100 border border-rose-700/60 flex items-center gap-1 transition-colors"
              title="Xuất file CSV danh sách cảnh báo đỏ"
            >
              <Download size={13} />
              <span className="hidden sm:inline">Xuất CSV</span>
            </button>

            <button
              onClick={handleCopyAlerts}
              className="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-rose-950/60 hover:bg-rose-900/80 text-rose-100 border border-rose-700/60 flex items-center gap-1 transition-colors"
              title="Sao chép báo cáo sự cố gửi cho FSE / Quản lý"
            >
              {copied ? (
                <>
                  <Check size={13} className="text-emerald-400" />
                  <span className="text-emerald-300 font-bold">Đã sao chép!</span>
                </>
              ) : (
                <>
                  <Copy size={13} />
                  <span className="hidden sm:inline">Sao chép báo cáo</span>
                </>
              )}
            </button>

            <button
              onClick={() => setIsExpanded(!isExpanded)}
              className="px-2.5 py-1.5 rounded-lg text-xs font-bold bg-white text-rose-950 hover:bg-rose-100 flex items-center gap-1 transition-colors shadow-sm ml-1"
            >
              {isExpanded ? (
                <>
                  <span>Thu gọn</span>
                  <ChevronUp size={15} />
                </>
              ) : (
                <>
                  <span>Xem chi tiết ({summary.totalAlerts})</span>
                  <ChevronDown size={15} />
                </>
              )}
            </button>
          </div>
        </div>

        {/* METRIC COUNTER CHIPS */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 mt-3.5 pt-3 border-t border-rose-700/50 text-xs">
          <div className="bg-rose-950/40 rounded-lg p-2 border border-rose-700/40">
            <span className="text-rose-300 text-[11px] block">Tổng thiết bị cảnh báo đỏ</span>
            <div className="flex items-baseline gap-1.5 mt-0.5">
              <span className="text-lg font-black text-white">{summary.totalAlerts}</span>
              <span className="text-[10px] text-rose-200">thiết bị tới hạn</span>
            </div>
          </div>

          <div className="bg-rose-950/40 rounded-lg p-2 border border-rose-700/40">
            <span className="text-rose-300 text-[11px] block">Health Index &lt; 50%</span>
            <div className="flex items-baseline gap-1.5 mt-0.5">
              <span className="text-lg font-black text-rose-300">{summary.lowHealthCount}</span>
              <span className="text-[10px] text-rose-200">suy giảm cách điện/sức khỏe</span>
            </div>
          </div>

          <div className="bg-rose-950/40 rounded-lg p-2 border border-rose-700/40">
            <span className="text-rose-300 text-[11px] block">Quá hạn bảo trì &gt; 30 ngày</span>
            <div className="flex items-baseline gap-1.5 mt-0.5">
              <span className="text-lg font-black text-rose-300">{summary.overdueCount}</span>
              <span className="text-[10px] text-rose-200">chậm tiến độ PM</span>
            </div>
          </div>

          <div className="bg-rose-950/40 rounded-lg p-2 border border-rose-700/40">
            <span className="text-rose-300 text-[11px] block">Nguy cơ kép (Cả hai)</span>
            <div className="flex items-baseline gap-1.5 mt-0.5">
              <span className="text-lg font-black text-amber-300">{summary.bothCount}</span>
              <span className="text-[10px] text-rose-200">vừa HI&lt;50% vừa quá hạn</span>
            </div>
          </div>
        </div>
      </div>

      {/* EXPANDED CONTENT: FILTER TABS, SEARCH, AND ALERT CARDS */}
      {isExpanded && (
        <div className="p-4 sm:p-5 bg-white space-y-4">
          {/* CONTROL BAR: TABS & SEARCH */}
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-rose-100">
            {/* Filter Tabs */}
            <div className="flex flex-wrap items-center gap-1.5">
              <button
                onClick={() => setActiveTab('all')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                  activeTab === 'all'
                    ? 'bg-rose-600 text-white shadow-xs'
                    : 'bg-rose-50 text-rose-800 hover:bg-rose-100'
                }`}
              >
                Tất cả cảnh báo ({summary.totalAlerts})
              </button>
              <button
                onClick={() => setActiveTab('low_health')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                  activeTab === 'low_health'
                    ? 'bg-rose-600 text-white shadow-xs'
                    : 'bg-rose-50 text-rose-800 hover:bg-rose-100'
                }`}
              >
                Health Index &lt; 50% ({summary.lowHealthCount})
              </button>
              <button
                onClick={() => setActiveTab('overdue_pm')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                  activeTab === 'overdue_pm'
                    ? 'bg-rose-600 text-white shadow-xs'
                    : 'bg-rose-50 text-rose-800 hover:bg-rose-100'
                }`}
              >
                Quá hạn &gt; 30 ngày ({summary.overdueCount})
              </button>
              <button
                onClick={() => setActiveTab('both')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                  activeTab === 'both'
                    ? 'bg-rose-600 text-white shadow-xs'
                    : 'bg-rose-50 text-rose-800 hover:bg-rose-100'
                }`}
              >
                Nguy cơ kép ({summary.bothCount})
              </button>
            </div>

            {/* In-Alert Search */}
            <div className="relative w-full sm:w-64">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={14} />
              <input
                type="text"
                placeholder="Tìm mã hoặc tên thiết bị cảnh báo..."
                className="w-full pl-8 pr-3 py-1.5 bg-slate-50 border border-slate-200 focus:bg-white focus:border-rose-500 focus:ring-1 focus:ring-rose-200 rounded-lg text-xs outline-none transition-all"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>
          </div>

          {/* EQUIPMENT CARDS GRID */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
            {filteredAlerts.length > 0 ? (
              filteredAlerts.map((item) => (
                <div
                  key={item.id}
                  className="bg-white rounded-xl border border-rose-200 hover:border-rose-400 p-4 shadow-xs transition-all hover:shadow-sm flex flex-col justify-between relative overflow-hidden group"
                >
                  {/* Left severity strip */}
                  <div className={`absolute top-0 left-0 bottom-0 w-1.5 ${
                    item.alertType === 'both' ? 'bg-rose-600' : 'bg-red-500'
                  }`} />

                  <div>
                    {/* Top Row: ID, Badges, Category */}
                    <div className="flex items-start justify-between gap-2 pl-2">
                      <div>
                        <div className="flex flex-wrap items-center gap-1.5">
                          <span className="font-mono text-xs font-bold text-rose-700 bg-rose-50 border border-rose-200 px-2 py-0.5 rounded">
                            {item.id}
                          </span>
                          <span className="px-2 py-0.5 rounded text-[10px] font-semibold bg-slate-100 text-slate-700">
                            {item.type}
                          </span>
                          {item.alertType === 'both' ? (
                            <span className="px-1.5 py-0.5 rounded text-[10px] font-black bg-rose-600 text-white flex items-center gap-1">
                              <Flame size={10} /> Nguy cơ kép (HI + Quá hạn)
                            </span>
                          ) : item.isLowHealth ? (
                            <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-rose-100 text-rose-800 border border-rose-200">
                              HI &lt; 50%
                            </span>
                          ) : (
                            <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-amber-100 text-amber-800 border border-amber-200">
                              Quá hạn &gt; 30d
                            </span>
                          )}
                        </div>
                        <h4 className="font-bold text-slate-900 text-sm mt-1 group-hover:text-rose-600 transition-colors">
                          {item.name}
                        </h4>
                        <div className="flex items-center gap-2 text-xs text-slate-500 mt-0.5">
                          <span className="flex items-center gap-1">
                            <Factory size={11} className="text-slate-400" />
                            {item.factory}
                          </span>
                          <span>•</span>
                          <span className="flex items-center gap-1">
                            <MapPin size={11} className="text-slate-400" />
                            {item.location}
                          </span>
                        </div>
                      </div>

                      {/* Quick Focus Button */}
                      {onFilterToEquipment && (
                        <button
                          onClick={() => onFilterToEquipment(item.id)}
                          className="shrink-0 p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-colors border border-transparent hover:border-rose-200"
                          title="Lọc bảng thiết bị bên dưới theo thiết bị này"
                        >
                          <Eye size={16} />
                        </button>
                      )}
                    </div>

                    {/* Middle Row: Two Pillars (Health Gauge & Overdue Timeline) */}
                    <div className="grid grid-cols-2 gap-2 mt-3 pl-2">
                      {/* Pillar 1: Health Index Gauge */}
                      <div className={`p-2.5 rounded-lg border text-xs flex flex-col justify-between ${
                        item.health < 50
                          ? 'bg-rose-50/70 border-rose-200 text-rose-950'
                          : 'bg-slate-50 border-slate-200 text-slate-800'
                      }`}>
                        <div className="flex items-center justify-between">
                          <span className="text-[11px] font-medium text-slate-600 flex items-center gap-1">
                            <Activity size={12} className={item.health < 50 ? "text-rose-600" : "text-slate-400"} />
                            Health Index
                          </span>
                          <span className={`font-black text-sm ${item.health < 50 ? 'text-rose-600' : 'text-slate-700'}`}>
                            {item.health}%
                          </span>
                        </div>
                        <div className="w-full bg-slate-200 h-2 rounded-full overflow-hidden mt-1.5">
                          <div
                            className={`h-full rounded-full ${
                              item.health < 50 ? 'bg-rose-600' : 'bg-emerald-500'
                            }`}
                            style={{ width: `${Math.min(item.health, 100)}%` }}
                          />
                        </div>
                        <div className="flex items-center justify-between text-[10px] text-slate-500 mt-1">
                          <span>Ngưỡng NETA: 50%</span>
                          {item.health < 50 ? (
                            <span className="text-rose-600 font-bold">Thấp hơn ngưỡng {(50 - item.health)}%</span>
                          ) : (
                            <span className="text-emerald-600 font-medium">Bình thường</span>
                          )}
                        </div>
                      </div>

                      {/* Pillar 2: Overdue Status */}
                      <div className={`p-2.5 rounded-lg border text-xs flex flex-col justify-between ${
                        item.overdueDays > 30
                          ? 'bg-rose-50/70 border-rose-200 text-rose-950'
                          : 'bg-slate-50 border-slate-200 text-slate-800'
                      }`}>
                        <div className="flex items-center justify-between">
                          <span className="text-[11px] font-medium text-slate-600 flex items-center gap-1">
                            <Clock size={12} className={item.overdueDays > 30 ? "text-rose-600" : "text-slate-400"} />
                            Tiến độ PM
                          </span>
                          <span className={`font-black text-xs ${item.overdueDays > 30 ? 'text-rose-600' : 'text-slate-700'}`}>
                            {item.overdueDays > 0 ? `Quá hạn ${item.overdueDays} ngày` : 'Đúng hạn'}
                          </span>
                        </div>
                        <div className="text-[11px] text-slate-600 mt-1 flex flex-col gap-0.5">
                          <div className="flex justify-between">
                            <span>Kỳ kiểm tra:</span>
                            <span className="font-semibold text-slate-800">{item.pmCycleMonths || 6} tháng</span>
                          </div>
                          <div className="flex justify-between">
                            <span>Lần cuối:</span>
                            <span className="font-mono text-slate-700">{item.lastCheck || 'Chưa rõ'}</span>
                          </div>
                        </div>
                        <div className="text-[10px] text-rose-600 font-semibold mt-0.5">
                          {item.overdueDays > 30 ? `Vượt quá giới hạn cho phép ${item.overdueDays - 30} ngày` : 'Trong hạn'}
                        </div>
                      </div>
                    </div>

                    {/* Diagnostic Alarm Reasons */}
                    <div className="mt-2.5 pl-2">
                      <div className="bg-rose-50/40 rounded-lg p-2 border border-rose-100 text-xs text-rose-900 space-y-1">
                        {item.reasons.map((reason, idx) => (
                          <div key={idx} className="flex items-start gap-1.5 text-[11px]">
                            <AlertTriangle size={12} className="text-rose-600 shrink-0 mt-0.5" />
                            <span>{reason}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>

                  {/* BOTTOM ACTION BUTTONS */}
                  <div className="mt-3 pt-2.5 border-t border-slate-100 pl-2 flex flex-wrap items-center justify-between gap-2">
                    <div className="flex items-center gap-1.5">
                      {onViewEquipmentProfile && (
                        <button
                          onClick={() => onViewEquipmentProfile(item.raw)}
                          className="px-2.5 py-1 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition-colors flex items-center gap-1"
                          title="Xem hồ sơ lý lịch kỹ thuật chi tiết của thiết bị"
                        >
                          <History size={12} className="text-slate-500" />
                          <span>Hồ sơ</span>
                        </button>
                      )}

                      {onSelectRiskDetail && (
                        <button
                          onClick={() => onSelectRiskDetail(item.id)}
                          className="px-2.5 py-1 text-xs font-semibold text-purple-700 bg-purple-50 hover:bg-purple-100 border border-purple-200 rounded-lg transition-colors flex items-center gap-1"
                          title="Xem phân tích ma trận rủi ro và FMEA"
                        >
                          <Zap size={12} className="text-purple-600" />
                          <span>Ma trận rủi ro</span>
                        </button>
                      )}
                    </div>

                    <div className="flex items-center gap-1.5">
                      {onNavigateToInspection && (
                        <button
                          onClick={() => onNavigateToInspection(item.raw)}
                          className="px-2.5 py-1 text-xs font-semibold text-blue-700 bg-blue-50 hover:bg-blue-100 border border-blue-200 rounded-lg transition-colors flex items-center gap-1"
                          title="Nhập biên bản thử nghiệm NETA / IEEE"
                        >
                          <FileText size={12} />
                          <span>Nhập NETA</span>
                        </button>
                      )}

                      {onOpenWorkOrderModal && (
                        <button
                          onClick={() => onOpenWorkOrderModal(item.raw)}
                          className="px-3 py-1 text-xs font-bold text-white bg-rose-600 hover:bg-rose-700 rounded-lg transition-colors flex items-center gap-1 shadow-xs"
                          title="Tạo phiếu công tác bảo trì / sửa chữa CMMS khẩn cấp"
                        >
                          <Wrench size={12} />
                          <span>Tạo CMMS</span>
                        </button>
                      )}
                    </div>
                  </div>
                </div>
              ))
            ) : (
              <div className="col-span-full py-8 text-center text-slate-500 bg-slate-50 rounded-xl border border-dashed border-slate-200">
                <p className="text-xs">Không tìm thấy thiết bị nào khớp với bộ lọc cảnh báo đỏ.</p>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};
