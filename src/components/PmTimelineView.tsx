import React, { useState, useMemo, useRef } from 'react';
import {
  Calendar,
  ChevronLeft,
  ChevronRight,
  Clock,
  AlertTriangle,
  CheckCircle2,
  AlertOctagon,
  Wrench,
  Search,
  Filter,
  ArrowRight,
  Sparkles,
  Info,
  CalendarDays,
  Layers,
  Zap,
  TrendingUp,
  Maximize2
} from 'lucide-react';

export interface PmTimelineViewProps {
  workOrders: any[];
  allEquipment: any[];
  onOpenWorkOrderModal?: (workOrder?: any) => void;
  onUpdateWorkOrderStatus?: (woId: string, newStatus: string) => void;
  initialMonth?: string; // YYYY-MM
}

// Built-in realistic PM seed items to ensure calendar has robust PM tasks across the month
const SEED_TIMELINE_ORDERS = [
  {
    id: 'WO-PM-2026-001',
    title: 'Bảo trì định kỳ Quý 3 - Máy biến áp dầu T1 2000kVA',
    equipmentId: 'MBA-T1-2000KVA',
    equipmentName: 'Máy biến áp chính T1 (2000kVA 22/0.4kV)',
    equipmentType: 'Máy biến áp',
    factory: 'Nhà máy Điện Gió Mũi Né',
    priority: 'high',
    type: 'preventive',
    status: 'completed',
    assignedTo: 'Nguyễn Văn S',
    dueDate: '2026-09-07',
    estimatedHours: 6
  },
  {
    id: 'WO-PM-2026-002',
    title: 'Bảo dưỡng định kỳ Tháng 9 - Bộ chuyển nguồn ATS 800A',
    equipmentId: 'ATS-LV-800A-01',
    equipmentName: 'Tủ chuyển nguồn ATS Khẩn cấp',
    equipmentType: 'Tủ điện',
    factory: 'Trạm Phân Phối KCN VSIP 1',
    priority: 'medium',
    type: 'preventive',
    status: 'completed',
    assignedTo: 'Trần Đình K',
    dueDate: '2026-09-10',
    estimatedHours: 4
  },
  {
    id: 'WO-PM-2026-003',
    title: 'Thí nghiệm Phóng điện cục bộ (PD) & TEV Tủ RMU 24kV',
    equipmentId: 'RMU-24KV-S01',
    equipmentName: 'Tủ RMU Schneider 24kV Ngăn lộ E01',
    equipmentType: 'Tủ điện',
    factory: 'Nhà máy Điện Mặt Trời Super Energy',
    priority: 'urgent',
    type: 'preventive',
    status: 'initiated',
    dueDate: '2026-09-18', // Overdue
    assignedTo: 'Lê Hoàng D',
    estimatedHours: 5
  },
  {
    id: 'WO-PM-2026-004',
    title: 'Kiểm tra bảo dưỡng định kỳ Tủ máy cắt MVCB-001',
    equipmentId: 'MVCB-001',
    equipmentName: 'Tủ máy cắt trung thế MVCB-001 24kV',
    equipmentType: 'Tủ điện',
    factory: 'Nhà máy Điện Mặt Trời Super Energy',
    priority: 'high',
    type: 'preventive',
    status: 'in_progress',
    dueDate: '2026-09-24', // Today
    assignedTo: 'Nguyễn Văn S',
    estimatedHours: 6
  },
  {
    id: 'WO-PM-2026-005',
    title: 'Bảo trì hệ thống ắc quy dự phòng UPS 40kVA',
    equipmentId: 'UPS-BATT-40KVA',
    equipmentName: 'Hệ ắc quy VRLA 12V-100Ah Trạm điều khiển',
    equipmentType: 'Pin & UPS',
    factory: 'Trạm Biến Áp 110kV Bắc Ninh',
    priority: 'medium',
    type: 'preventive',
    status: 'initiated',
    dueDate: '2026-09-26', // Due in 2 days
    assignedTo: 'Vũ Quốc A',
    estimatedHours: 4
  },
  {
    id: 'WO-PM-2026-006',
    title: 'Kiểm tra đo đặc tính I-V Chuỗi Inverter Năng lượng mặt trời 110kW',
    equipmentId: 'PV-INV-100K-03',
    equipmentName: 'Chuỗi biến tần Solar Inverter Sungrow 110kW',
    equipmentType: 'Inverter',
    factory: 'Nhà máy Điện Gió Mũi Né',
    priority: 'high',
    type: 'preventive',
    status: 'in_progress',
    dueDate: '2026-09-29', // Due in 5 days
    assignedTo: 'Lê Hoàng D',
    estimatedHours: 5
  },
  {
    id: 'WO-PM-2026-007',
    title: 'Thử nghiệm Rơ-le bảo vệ so lệch đường dây SEL-311L',
    equipmentId: 'RELAY-SEL-311L',
    equipmentName: 'Rơ-le số bảo vệ kỹ thuật số SEL-311L',
    equipmentType: 'Rơ le',
    factory: 'Trạm Phân Phối KCN VSIP 1',
    priority: 'medium',
    type: 'preventive',
    status: 'initiated',
    dueDate: '2026-09-30',
    assignedTo: 'Phạm Minh T',
    estimatedHours: 4
  }
];

export const PmTimelineView: React.FC<PmTimelineViewProps> = ({
  workOrders,
  allEquipment,
  onOpenWorkOrderModal,
  onUpdateWorkOrderStatus,
  initialMonth = '2026-09'
}) => {
  const [selectedMonth, setSelectedMonth] = useState<string>(initialMonth);
  const [filterType, setFilterType] = useState<string>('all');
  const [filterPriority, setFilterPriority] = useState<string>('all');
  const [filterStatus, setFilterStatus] = useState<string>('all');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [highlightGapDaysOnly, setHighlightGapDaysOnly] = useState<boolean>(false);

  const scrollContainerRef = useRef<HTMLDivElement>(null);

  // App current reference date
  const todayStr = '2026-09-24';
  const todayDate = useMemo(() => new Date(todayStr), []);

  // Merge live PM work orders with realistic seeds
  const mergedPmOrders = useMemo(() => {
    const livePmList = (workOrders || []).filter(
      (wo) => wo.type === 'preventive' || wo.pmFrequency || (wo.id && wo.id.startsWith('WO-PM'))
    );
    const existingIds = new Set(livePmList.map((w) => w.id?.toLowerCase()));
    const result = [...livePmList];
    for (const seed of SEED_TIMELINE_ORDERS) {
      if (!existingIds.has(seed.id.toLowerCase())) {
        result.push(seed);
      }
    }
    return result;
  }, [workOrders]);

  // Compute Days in Selected Month
  const { daysInMonth, year, monthIndex } = useMemo(() => {
    const [yStr, mStr] = selectedMonth.split('-');
    const y = parseInt(yStr, 10) || 2026;
    const m = parseInt(mStr, 10) || 9; // 1-12
    const totalDays = new Date(y, m, 0).getDate();

    const days = [];
    for (let d = 1; d <= totalDays; d++) {
      const dateObj = new Date(y, m - 1, d);
      const dayNum = String(d).padStart(2, '0');
      const dateStr = `${y}-${String(m).padStart(2, '0')}-${dayNum}`;
      const dayOfWeek = dateObj.getDay(); // 0: CN, 1: T2...
      const dayName = ['CN', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7'][dayOfWeek];
      const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
      const isToday = dateStr === todayStr;

      days.push({
        dayNumber: d,
        dateStr,
        dayName,
        isWeekend,
        isToday,
        dateObj
      });
    }

    return { daysInMonth: days, year: y, monthIndex: m };
  }, [selectedMonth, todayStr]);

  // Map Orders to Date String
  const filteredOrders = useMemo(() => {
    return mergedPmOrders.filter((wo) => {
      if (filterType !== 'all' && (wo.equipmentType || '').toLowerCase() !== filterType.toLowerCase()) {
        return false;
      }
      if (filterPriority !== 'all' && (wo.priority || '').toLowerCase() !== filterPriority.toLowerCase()) {
        return false;
      }
      if (filterStatus !== 'all' && (wo.status || '').toLowerCase() !== filterStatus.toLowerCase()) {
        return false;
      }
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const titleMatch = (wo.title || '').toLowerCase().includes(q);
        const idMatch = (wo.id || '').toLowerCase().includes(q);
        const eqMatch = (wo.equipmentName || wo.equipmentId || '').toLowerCase().includes(q);
        const picMatch = (wo.assignedTo || '').toLowerCase().includes(q);
        if (!titleMatch && !idMatch && !eqMatch && !picMatch) return false;
      }
      return true;
    });
  }, [mergedPmOrders, filterType, filterPriority, filterStatus, searchQuery]);

  // Group orders by due date
  const ordersByDate = useMemo(() => {
    const map = new Map<string, any[]>();
    filteredOrders.forEach((wo) => {
      const due = (wo.dueDate || '').split('T')[0];
      if (due) {
        const list = map.get(due) || [];
        list.push(wo);
        map.set(due, list);
      }
    });
    return map;
  }, [filteredOrders]);

  // Calculate Maintenance Gaps (Consecutive days with ZERO scheduled work orders)
  const maintenanceGapsAnalysis = useMemo(() => {
    const gaps: Array<{
      startDate: string;
      endDate: string;
      durationDays: number;
      startDay: number;
      endDay: number;
      isFuture: boolean;
      isPast: boolean;
    }> = [];

    let currentGapStart: number | null = null;

    daysInMonth.forEach((day, index) => {
      const orders = ordersByDate.get(day.dateStr) || [];
      const hasWork = orders.length > 0;

      if (!hasWork) {
        if (currentGapStart === null) {
          currentGapStart = day.dayNumber;
        }
      } else {
        if (currentGapStart !== null) {
          const prevDay = daysInMonth[index - 1];
          const duration = prevDay.dayNumber - currentGapStart + 1;
          if (duration >= 2) { // Consider gap if 2 or more consecutive days without PM
            const startDateStr = `${year}-${String(monthIndex).padStart(2, '0')}-${String(currentGapStart).padStart(2, '0')}`;
            const isFuture = startDateStr >= todayStr;
            const isPast = prevDay.dateStr < todayStr;
            gaps.push({
              startDate: startDateStr,
              endDate: prevDay.dateStr,
              durationDays: duration,
              startDay: currentGapStart,
              endDay: prevDay.dayNumber,
              isFuture,
              isPast
            });
          }
          currentGapStart = null;
        }
      }
    });

    // Check if ends with a gap
    if (currentGapStart !== null) {
      const lastDay = daysInMonth[daysInMonth.length - 1];
      const duration = lastDay.dayNumber - currentGapStart + 1;
      if (duration >= 2) {
        const startDateStr = `${year}-${String(monthIndex).padStart(2, '0')}-${String(currentGapStart).padStart(2, '0')}`;
        gaps.push({
          startDate: startDateStr,
          endDate: lastDay.dateStr,
          durationDays: duration,
          startDay: currentGapStart,
          endDay: lastDay.dayNumber,
          isFuture: startDateStr >= todayStr,
          isPast: lastDay.dateStr < todayStr
        });
      }
    }

    const totalGapDays = gaps.reduce((acc, g) => acc + g.durationDays, 0);
    const longestGap = gaps.length > 0 ? Math.max(...gaps.map(g => g.durationDays)) : 0;
    const upcomingGaps = gaps.filter(g => g.isFuture);
    const pastGaps = gaps.filter(g => g.isPast);

    // Overloaded days (> 2 work orders)
    const overloadDays = daysInMonth.filter(d => (ordersByDate.get(d.dateStr) || []).length >= 3);

    return {
      gaps,
      totalGapDays,
      longestGap,
      upcomingGaps,
      pastGaps,
      overloadDays
    };
  }, [daysInMonth, ordersByDate, year, monthIndex, todayStr]);

  // Is a day inside a maintenance gap?
  const getGapInfoForDay = (dayNumber: number) => {
    return maintenanceGapsAnalysis.gaps.find(
      g => dayNumber >= g.startDay && dayNumber <= g.endDay
    ) || null;
  };

  // Scroll handlers
  const handleScroll = (direction: 'left' | 'right') => {
    if (scrollContainerRef.current) {
      const offset = direction === 'left' ? -380 : 380;
      scrollContainerRef.current.scrollBy({ left: offset, behavior: 'smooth' });
    }
  };

  const handleScrollToToday = () => {
    const todayIndex = daysInMonth.findIndex(d => d.isToday);
    if (todayIndex >= 0 && scrollContainerRef.current) {
      const targetScroll = Math.max(0, todayIndex * 140 - 200);
      scrollContainerRef.current.scrollTo({ left: targetScroll, behavior: 'smooth' });
    }
  };

  return (
    <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden flex flex-col">
      {/* TIMELINE TOP HEADER */}
      <div className="p-5 border-b border-slate-200 bg-gradient-to-r from-slate-900 via-slate-800 to-indigo-950 text-white">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2.5">
              <div className="w-9 h-9 rounded-xl bg-blue-600/30 border border-blue-400/40 flex items-center justify-center text-blue-400 shadow-xs">
                <CalendarDays className="w-5 h-5" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <h2 className="text-base sm:text-lg font-bold text-white tracking-tight">
                    Lịch Dòng Thời Gian Bảo Trì (PM Timeline Calendar)
                  </h2>
                  <span className="px-2 py-0.5 rounded-full text-[10px] font-black bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                    Live Gantt
                  </span>
                </div>
                <p className="text-xs text-slate-300 mt-0.5">
                  Theo dõi trực quan hạn hoàn thành phiếu PM theo trục ngang &amp; Phát hiện sớm các khoảng trống bảo trì (Maintenance Gaps)
                </p>
              </div>
            </div>
          </div>

          {/* Quick Date Navigator & Controls */}
          <div className="flex flex-wrap items-center gap-2">
            <div className="flex items-center bg-slate-800/80 rounded-xl p-1 border border-slate-700 text-xs">
              <input
                type="month"
                value={selectedMonth}
                onChange={(e) => setSelectedMonth(e.target.value)}
                className="bg-transparent text-white font-bold outline-none px-2 py-1 text-xs cursor-pointer"
              />
            </div>

            <button
              onClick={handleScrollToToday}
              className="px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-xs font-bold transition-all shadow-sm flex items-center gap-1.5 active:scale-95"
              title="Cuộn tới ngày hôm nay"
            >
              <Clock className="w-3.5 h-3.5" />
              <span>Hôm Nay ({todayStr.slice(8, 10)}/{todayStr.slice(5, 7)})</span>
            </button>

            <div className="flex items-center gap-1 bg-slate-800/80 rounded-xl p-0.5 border border-slate-700">
              <button
                onClick={() => handleScroll('left')}
                className="p-1.5 hover:bg-slate-700 text-slate-300 hover:text-white rounded-lg transition-colors"
                title="Cuộn sang trái"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>
              <button
                onClick={() => handleScroll('right')}
                className="p-1.5 hover:bg-slate-700 text-slate-300 hover:text-white rounded-lg transition-colors"
                title="Cuộn sang phải"
              >
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        </div>

        {/* MAINTENANCE GAPS SUMMARY KPI BANNER */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mt-4 pt-4 border-t border-slate-700/60">
          <div className="p-2.5 rounded-xl bg-slate-800/60 border border-slate-700 flex items-center justify-between">
            <div>
              <p className="text-[10px] uppercase font-bold text-slate-400">Khoảng Trống Bảo Trì (Gaps)</p>
              <p className="text-lg font-black text-amber-300 font-mono">
                {maintenanceGapsAnalysis.gaps.length} <span className="text-xs font-normal text-slate-400">khoảng</span>
              </p>
            </div>
            <div className="w-8 h-8 rounded-lg bg-amber-500/10 text-amber-400 flex items-center justify-center font-bold text-xs">
              ⚡
            </div>
          </div>

          <div className="p-2.5 rounded-xl bg-slate-800/60 border border-slate-700 flex items-center justify-between">
            <div>
              <p className="text-[10px] uppercase font-bold text-slate-400">Tổng Ngày Trống Lịch PM</p>
              <p className="text-lg font-black text-emerald-400 font-mono">
                {maintenanceGapsAnalysis.totalGapDays} <span className="text-xs font-normal text-slate-400">ngày</span>
              </p>
            </div>
            <div className="w-8 h-8 rounded-lg bg-emerald-500/10 text-emerald-400 flex items-center justify-center font-bold text-xs">
              📅
            </div>
          </div>

          <div className="p-2.5 rounded-xl bg-slate-800/60 border border-slate-700 flex items-center justify-between">
            <div>
              <p className="text-[10px] uppercase font-bold text-slate-400">Khoảng Trống Dài Nhất</p>
              <p className="text-lg font-black text-sky-300 font-mono">
                {maintenanceGapsAnalysis.longestGap} <span className="text-xs font-normal text-slate-400">ngày liên tiếp</span>
              </p>
            </div>
            <div className="w-8 h-8 rounded-lg bg-sky-500/10 text-sky-400 flex items-center justify-center font-bold text-xs">
              ⏳
            </div>
          </div>

          <div className="p-2.5 rounded-xl bg-slate-800/60 border border-slate-700 flex items-center justify-between">
            <div>
              <p className="text-[10px] uppercase font-bold text-slate-400">Khoảng Trống Sắp Tới</p>
              <p className="text-lg font-black text-rose-300 font-mono">
                {maintenanceGapsAnalysis.upcomingGaps.length} <span className="text-xs font-normal text-slate-400">cần tối ưu</span>
              </p>
            </div>
            <div className="w-8 h-8 rounded-lg bg-rose-500/10 text-rose-400 flex items-center justify-center font-bold text-xs">
              🎯
            </div>
          </div>
        </div>
      </div>

      {/* FILTER CONTROLS BAR */}
      <div className="p-3 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 text-xs">
        <div className="flex flex-wrap items-center gap-2">
          {/* Search Input */}
          <div className="relative">
            <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              type="text"
              placeholder="Tìm WO, thiết bị, kỹ sư..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="pl-8 pr-3 py-1.5 bg-white border border-slate-200 rounded-lg text-slate-700 outline-none focus:border-blue-500 text-xs w-48 transition-all"
            />
          </div>

          {/* Status Filter */}
          <select
            value={filterStatus}
            onChange={(e) => setFilterStatus(e.target.value)}
            className="px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-slate-700 font-medium outline-none focus:border-blue-500 cursor-pointer"
          >
            <option value="all">Tất cả trạng thái</option>
            <option value="initiated">Mới tạo / Chờ duyệt</option>
            <option value="in_progress">Đang thực hiện</option>
            <option value="completed">Đã hoàn thành</option>
          </select>

          {/* Priority Filter */}
          <select
            value={filterPriority}
            onChange={(e) => setFilterPriority(e.target.value)}
            className="px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-slate-700 font-medium outline-none focus:border-blue-500 cursor-pointer"
          >
            <option value="all">Tất cả mức ưu tiên</option>
            <option value="urgent">Khẩn cấp (Urgent)</option>
            <option value="high">Cao (High)</option>
            <option value="medium">Trung bình (Medium)</option>
            <option value="low">Thấp (Low)</option>
          </select>

          {/* Equipment Type Filter */}
          <select
            value={filterType}
            onChange={(e) => setFilterType(e.target.value)}
            className="px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-slate-700 font-medium outline-none focus:border-blue-500 cursor-pointer"
          >
            <option value="all">Tất cả loại thiết bị</option>
            <option value="Máy biến áp">Máy biến áp</option>
            <option value="Tủ điện">Tủ điện / RMU</option>
            <option value="Động cơ">Động cơ</option>
            <option value="Inverter">Inverter Solar</option>
            <option value="Pin & UPS">Pin &amp; UPS</option>
            <option value="Rơ le">Rơ le bảo vệ</option>
          </select>
        </div>

        {/* Gap Highlighting Toggle */}
        <div className="flex items-center gap-2">
          <label className="flex items-center gap-1.5 cursor-pointer text-slate-700 select-none">
            <input
              type="checkbox"
              checked={highlightGapDaysOnly}
              onChange={(e) => setHighlightGapDaysOnly(e.target.checked)}
              className="rounded text-amber-600 focus:ring-amber-500"
            />
            <span className="font-semibold text-amber-800">⚡ Chỉ xem ngày có Khoảng Trống Bảo Trì (Gap Focus)</span>
          </label>
        </div>
      </div>

      {/* HORIZONTAL TIMELINE CALENDAR TRACK */}
      <div className="relative overflow-hidden bg-slate-100 flex-1 min-h-[460px]">
        <div
          ref={scrollContainerRef}
          className="overflow-x-auto overflow-y-auto max-h-[600px] p-4 select-none scrollbar-thin scrollbar-thumb-slate-300"
        >
          <div className="inline-flex gap-3 min-w-full pb-4">
            {daysInMonth.map((day) => {
              const orders = ordersByDate.get(day.dateStr) || [];
              const gap = getGapInfoForDay(day.dayNumber);
              const isGapDay = gap !== null;
              const isGapStart = gap && gap.startDay === day.dayNumber;

              if (highlightGapDaysOnly && !isGapDay) {
                return null;
              }

              return (
                <div
                  key={day.dateStr}
                  className={`shrink-0 w-44 rounded-2xl flex flex-col transition-all duration-200 border ${
                    day.isToday
                      ? 'bg-blue-50/80 border-blue-400 ring-2 ring-blue-500/20 shadow-md'
                      : isGapDay
                      ? 'bg-amber-50/40 border-amber-200/90 shadow-2xs'
                      : day.isWeekend
                      ? 'bg-slate-50/80 border-slate-200'
                      : 'bg-white border-slate-200 shadow-2xs'
                  }`}
                >
                  {/* Day Header Box */}
                  <div className={`p-2.5 rounded-t-2xl border-b flex items-center justify-between ${
                    day.isToday
                      ? 'bg-blue-600 text-white border-blue-700'
                      : isGapDay
                      ? 'bg-amber-100/70 text-amber-900 border-amber-200'
                      : 'bg-slate-50 text-slate-800 border-slate-100'
                  }`}>
                    <div className="flex items-center gap-1.5">
                      <span className={`text-base font-black font-mono leading-none ${
                        day.isToday ? 'text-white' : 'text-slate-900'
                      }`}>
                        {String(day.dayNumber).padStart(2, '0')}
                      </span>
                      <div className="flex flex-col">
                        <span className="text-[10px] font-bold uppercase leading-none opacity-80">
                          {day.dayName}
                        </span>
                        <span className="text-[9px] opacity-70 leading-none mt-0.5">
                          T{monthIndex}
                        </span>
                      </div>
                    </div>

                    <div className="flex items-center gap-1">
                      {day.isToday && (
                        <span className="px-1.5 py-0.5 rounded-md text-[9px] font-black bg-white text-blue-700 shadow-2xs uppercase tracking-wide">
                          Hôm Nay
                        </span>
                      )}

                      {orders.length > 0 ? (
                        <span className={`px-1.5 py-0.5 rounded-full text-[10px] font-bold ${
                          day.isToday
                            ? 'bg-blue-700 text-white'
                            : 'bg-blue-100 text-blue-800'
                        }`}>
                          {orders.length} WO
                        </span>
                      ) : (
                        <span className="text-[10px] text-slate-400 font-semibold italic">
                          Trống
                        </span>
                      )}
                    </div>
                  </div>

                  {/* Day Content Area */}
                  <div className="p-2.5 flex-1 flex flex-col gap-2 min-h-[300px]">
                    {/* GAP NOTIFICATION CARD (If gap starts on this day) */}
                    {isGapStart && (
                      <div className="p-2 rounded-xl bg-amber-500/10 border border-amber-300 text-amber-900 text-[11px] leading-tight flex items-start gap-1.5 mb-1 animate-pulse">
                        <span className="text-amber-600 font-bold shrink-0">⚡</span>
                        <div>
                          <strong className="block font-bold">Khoảng trống {gap.durationDays} ngày</strong>
                          <span className="text-[10px] text-amber-700">
                            ({gap.startDate.slice(8, 10)}/T{monthIndex} - {gap.endDate.slice(8, 10)}/T{monthIndex})
                          </span>
                          <p className="text-[9px] text-amber-800 mt-0.5">
                            Cơ hội bố trí thêm PM hoặc bảo dưỡng định kỳ
                          </p>
                        </div>
                      </div>
                    )}

                    {/* WORK ORDER CARDS FOR THIS DATE */}
                    {orders.length > 0 ? (
                      orders.map((wo: any) => {
                        const isCompleted = wo.status === 'completed';
                        const isOverdue = !isCompleted && wo.dueDate < todayStr;
                        const isDueSoon = !isCompleted && !isOverdue && wo.dueDate <= '2026-09-30';

                        let badgeColor = 'bg-slate-100 text-slate-700 border-slate-200';
                        if (isCompleted) {
                          badgeColor = 'bg-emerald-50 text-emerald-800 border-emerald-300';
                        } else if (isOverdue) {
                          badgeColor = 'bg-rose-50 text-rose-800 border-rose-300';
                        } else if (isDueSoon) {
                          badgeColor = 'bg-amber-50 text-amber-800 border-amber-300';
                        } else {
                          badgeColor = 'bg-blue-50 text-blue-800 border-blue-300';
                        }

                        let priorityColor = 'bg-blue-100 text-blue-700';
                        if (wo.priority === 'urgent') priorityColor = 'bg-rose-100 text-rose-700 font-bold';
                        else if (wo.priority === 'high') priorityColor = 'bg-orange-100 text-orange-700 font-bold';
                        else if (wo.priority === 'medium') priorityColor = 'bg-amber-100 text-amber-700';

                        return (
                          <div
                            key={wo.id}
                            onClick={() => onOpenWorkOrderModal?.(wo)}
                            className={`p-2.5 rounded-xl border transition-all cursor-pointer hover:shadow-md hover:scale-[1.02] active:scale-98 flex flex-col justify-between gap-1.5 ${badgeColor}`}
                          >
                            <div className="flex items-center justify-between gap-1">
                              <span className="font-mono text-[10px] font-black tracking-tight text-slate-800 truncate">
                                {wo.id}
                              </span>
                              <span className={`px-1.5 py-0.2 rounded text-[9px] uppercase font-bold shrink-0 ${priorityColor}`}>
                                {wo.priority || 'PM'}
                              </span>
                            </div>

                            <p className="text-[11px] font-bold text-slate-900 line-clamp-2 leading-tight">
                              {wo.title}
                            </p>

                            <div className="flex items-center gap-1 text-[10px] text-slate-600 truncate">
                              <Wrench className="w-3 h-3 text-slate-400 shrink-0" />
                              <span className="truncate">{wo.equipmentName || wo.equipmentId}</span>
                            </div>

                            <div className="pt-1.5 mt-0.5 border-t border-slate-200/60 flex items-center justify-between text-[10px]">
                              <span className="font-medium text-slate-700 truncate max-w-[90px]">
                                👤 {wo.assignedTo || 'Chưa gán'}
                              </span>

                              {isCompleted ? (
                                <span className="flex items-center gap-0.5 text-emerald-700 font-bold">
                                  <CheckCircle2 className="w-3 h-3 text-emerald-600" /> Xong
                                </span>
                              ) : isOverdue ? (
                                <span className="flex items-center gap-0.5 text-rose-700 font-bold animate-pulse">
                                  <AlertOctagon className="w-3 h-3 text-rose-600" /> Trễ
                                </span>
                              ) : (
                                <span className="text-slate-500 font-medium">
                                  {wo.estimatedHours ? `${wo.estimatedHours}h` : 'PM'}
                                </span>
                              )}
                            </div>
                          </div>
                        );
                      })
                    ) : (
                      <div className="flex-1 flex flex-col items-center justify-center text-center p-3 text-slate-400">
                        <span className="text-lg opacity-40 mb-1">💤</span>
                        <span className="text-[10px] font-semibold text-slate-400">
                          {isGapDay ? 'Khoảng trống bảo trì' : 'Không có phiếu'}
                        </span>
                        <span className="text-[9px] text-slate-400 mt-0.5">
                          Sẵn sàng xếp lịch
                        </span>
                      </div>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* TIMELINE FOOTER LEGEND & OPTIMIZATION INSIGHTS */}
      <div className="p-4 bg-white border-t border-slate-200 flex flex-col md:flex-row items-start md:items-center justify-between gap-4 text-xs">
        <div className="flex flex-wrap items-center gap-4">
          <span className="text-slate-500 font-bold uppercase text-[10px]">Chú Thích Màu Sắc:</span>
          <div className="flex items-center gap-1.5">
            <span className="w-2.5 h-2.5 rounded-full bg-emerald-500" />
            <span className="text-slate-700">Đã hoàn thành</span>
          </div>
          <div className="flex items-center gap-1.5">
            <span className="w-2.5 h-2.5 rounded-full bg-rose-500 animate-pulse" />
            <span className="text-slate-700 font-medium">Quá hạn (Overdue)</span>
          </div>
          <div className="flex items-center gap-1.5">
            <span className="w-2.5 h-2.5 rounded-full bg-amber-500" />
            <span className="text-slate-700">Sắp đến hạn (≤7d)</span>
          </div>
          <div className="flex items-center gap-1.5">
            <span className="w-2.5 h-2.5 rounded-full bg-blue-500" />
            <span className="text-slate-700">Đang thực hiện</span>
          </div>
          <div className="flex items-center gap-1.5">
            <span className="px-1.5 py-0.5 rounded bg-amber-100 text-amber-800 font-bold text-[10px]">
              ⚡ Khoảng trống
            </span>
            <span className="text-slate-700">Maintenance Gap</span>
          </div>
        </div>

        <div className="text-[11px] text-slate-500 flex items-center gap-1.5">
          <Info className="w-3.5 h-3.5 text-blue-500 shrink-0" />
          <span>
            Bấm vào từng thẻ phiếu để xem chi tiết, phân công kỹ sư hoặc cập nhật trạng thái nghiệm thu.
          </span>
        </div>
      </div>
    </div>
  );
};
