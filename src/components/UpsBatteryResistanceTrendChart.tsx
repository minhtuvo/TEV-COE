import React, { useState, useMemo } from 'react';
import {
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
  ReferenceLine
} from 'recharts';
import {
  TrendingUp,
  BatteryCharging,
  AlertTriangle,
  ShieldAlert,
  ShieldCheck,
  Plus,
  Trash2,
  RotateCcw,
  Download,
  Info,
  Calendar,
  Layers,
  ArrowUpRight,
  Maximize2,
  Activity,
  Sliders,
  CheckCircle2,
  FileSpreadsheet
} from 'lucide-react';

export interface UpsBatteryResistanceRecord {
  id: string;
  date: string; // YYYY-MM-DD
  cycleLabel: string; // e.g. "Q1/2024 (Baseline)"
  stringAvgResistance: number; // mΩ
  worstCellResistance: number; // mΩ
  minCellResistance: number; // mΩ
  worstCellId: string; // e.g. "Cell #18"
  baselineResistance: number; // mΩ
  maxDeviationPercent: number; // %
  floatVoltageV: number; // V
  temperatureCelsius: number; // °C
  fseInspector: string;
  status: 'NORMAL' | 'CAUTION' | 'CRITICAL';
  notes: string;
  cell1?: number;
  cell12?: number;
  cell18?: number;
  cell36?: number;
}

export interface UpsBatteryResistanceTrendChartProps {
  systemTag?: string;
  batteryType?: string;
  batteryCapacityAh?: number;
  currentDeviationPercent?: number;
  currentWorstCell?: string;
  lang?: 'vi' | 'en';
  onUpdateCurrentValue?: (deviation: number, worstCell: string) => void;
  className?: string;
}

// Preset datasets for various field systems
const PRESET_SYSTEM_DATASETS: Record<string, UpsBatteryResistanceRecord[]> = {
  'UPS-CRITICAL-DATACENTER-01': [
    {
      id: 'rec-1',
      date: '2024-03-15',
      cycleLabel: 'Q1/2024 (Nghiệm Thu Gốc)',
      stringAvgResistance: 1.25,
      worstCellResistance: 1.26,
      minCellResistance: 1.23,
      worstCellId: 'Cell #18',
      baselineResistance: 1.25,
      maxDeviationPercent: 0.8,
      floatVoltageV: 13.65,
      temperatureCelsius: 23.5,
      fseInspector: 'Võ Minh Tú',
      status: 'NORMAL',
      notes: 'Nghiệm thu đóng điện ban đầu (Baseline Acceptance). Tất cả 40 cell đồng đều hoàn hảo.',
      cell1: 1.24,
      cell12: 1.25,
      cell18: 1.26,
      cell36: 1.23
    },
    {
      id: 'rec-2',
      date: '2024-09-18',
      cycleLabel: 'Q3/2024 (Định Kỳ 6T)',
      stringAvgResistance: 1.27,
      worstCellResistance: 1.31,
      minCellResistance: 1.24,
      worstCellId: 'Cell #18',
      baselineResistance: 1.25,
      maxDeviationPercent: 4.8,
      floatVoltageV: 13.62,
      temperatureCelsius: 24.0,
      fseInspector: 'Trần Văn Long',
      status: 'NORMAL',
      notes: 'Bảo dưỡng định kỳ 6 tháng NETA MTS. Lệch trở kháng tăng nhẹ nhưng vẫn < 10%.',
      cell1: 1.25,
      cell12: 1.26,
      cell18: 1.31,
      cell36: 1.24
    },
    {
      id: 'rec-3',
      date: '2025-03-22',
      cycleLabel: 'Q1/2025 (Định Kỳ 12T)',
      stringAvgResistance: 1.29,
      worstCellResistance: 1.41,
      minCellResistance: 1.25,
      worstCellId: 'Cell #18',
      baselineResistance: 1.25,
      maxDeviationPercent: 12.8,
      floatVoltageV: 13.58,
      temperatureCelsius: 24.8,
      fseInspector: 'Võ Minh Tú',
      status: 'CAUTION',
      notes: 'Cảnh báo IEEE 1188: Cell #18 tăng trở kháng +12.8% so với baseline. Tăng tần suất đo kiểm.',
      cell1: 1.26,
      cell12: 1.28,
      cell18: 1.41,
      cell36: 1.25
    },
    {
      id: 'rec-4',
      date: '2025-09-15',
      cycleLabel: 'Q3/2025 (Định Kỳ 18T)',
      stringAvgResistance: 1.31,
      worstCellResistance: 1.55,
      minCellResistance: 1.26,
      worstCellId: 'Cell #18',
      baselineResistance: 1.25,
      maxDeviationPercent: 24.0,
      floatVoltageV: 13.50,
      temperatureCelsius: 25.5,
      fseInspector: 'Nguyễn Hoàng Sơn',
      status: 'CAUTION',
      notes: 'Cell #18 vượt ngưỡng 20% (đạt +24.0%). Khuyến cáo kiểm tra lực siết cọc bình và nội trở điện cực.',
      cell1: 1.28,
      cell12: 1.29,
      cell18: 1.55,
      cell36: 1.26
    },
    {
      id: 'rec-5',
      date: '2026-03-20',
      cycleLabel: 'Q1/2026 (Định Kỳ 24T)',
      stringAvgResistance: 1.33,
      worstCellResistance: 1.68,
      minCellResistance: 1.27,
      worstCellId: 'Cell #18',
      baselineResistance: 1.25,
      maxDeviationPercent: 34.4,
      floatVoltageV: 13.42,
      temperatureCelsius: 26.2,
      fseInspector: 'Võ Minh Tú',
      status: 'CRITICAL',
      notes: 'Suy thoái nghiêm trọng trên Cell #18 (+34.4%). Nguy cơ đảo cực khi xả tải cao.',
      cell1: 1.29,
      cell12: 1.31,
      cell18: 1.68,
      cell36: 1.27
    },
    {
      id: 'rec-6',
      date: '2026-10-18',
      cycleLabel: 'Q3/2026 (Chạy Thử Hiện Tại)',
      stringAvgResistance: 1.35,
      worstCellResistance: 1.82,
      minCellResistance: 1.28,
      worstCellId: 'Cell #18',
      baselineResistance: 1.25,
      maxDeviationPercent: 45.6,
      floatVoltageV: 13.31,
      temperatureCelsius: 27.8,
      fseInspector: 'Võ Minh Tú',
      status: 'CRITICAL',
      notes: 'NGUY HẠI KHẨN CẤP: Cell #18 tăng vọt +45.6%, bu-lông lỏng phát nhiệt Delta-T = 18.2°C. Bắt buộc thay thế block trước khi cấp điện tải thực.',
      cell1: 1.30,
      cell12: 1.32,
      cell18: 1.82,
      cell36: 1.28
    }
  ],
  'UPS-DATACENTER-TIER3-A': [
    {
      id: 'rec-t1',
      date: '2024-06-10',
      cycleLabel: 'Q2/2024 (Nghiệm Thu)',
      stringAvgResistance: 1.22,
      worstCellResistance: 1.23,
      minCellResistance: 1.20,
      worstCellId: 'Cell #08',
      baselineResistance: 1.22,
      maxDeviationPercent: 0.8,
      floatVoltageV: 13.65,
      temperatureCelsius: 22.8,
      fseInspector: 'Võ Minh Tú',
      status: 'NORMAL',
      notes: 'Nghiệm thu đạt chuẩn 100% NETA ATS 7.18.',
      cell1: 1.21,
      cell12: 1.22,
      cell18: 1.22,
      cell36: 1.23
    },
    {
      id: 'rec-t2',
      date: '2025-06-12',
      cycleLabel: 'Q2/2025 (Định Kỳ 12T)',
      stringAvgResistance: 1.24,
      worstCellResistance: 1.26,
      minCellResistance: 1.21,
      worstCellId: 'Cell #08',
      baselineResistance: 1.22,
      maxDeviationPercent: 3.2,
      floatVoltageV: 13.62,
      temperatureCelsius: 23.1,
      fseInspector: 'Trần Văn Long',
      status: 'NORMAL',
      notes: 'Tình trạng bình hoàn hảo, độ lệch trở kháng thấp.',
      cell1: 1.23,
      cell12: 1.24,
      cell18: 1.24,
      cell36: 1.25
    },
    {
      id: 'rec-t3',
      date: '2026-10-18',
      cycleLabel: 'Q3/2026 (Hiện Tại)',
      stringAvgResistance: 1.25,
      worstCellResistance: 1.28,
      minCellResistance: 1.22,
      worstCellId: 'Cell #08',
      baselineResistance: 1.22,
      maxDeviationPercent: 4.9,
      floatVoltageV: 13.60,
      temperatureCelsius: 23.5,
      fseInspector: 'Võ Minh Tú',
      status: 'NORMAL',
      notes: 'Đạt chuẩn 100% NETA ECS Section 8.2 & IEEE 1188.',
      cell1: 1.24,
      cell12: 1.25,
      cell18: 1.25,
      cell36: 1.27
    }
  ]
};

export const UpsBatteryResistanceTrendChart: React.FC<UpsBatteryResistanceTrendChartProps> = ({
  systemTag = 'UPS-CRITICAL-DATACENTER-01',
  batteryType = 'VRLA (12V 150Ah x 40 bình)',
  batteryCapacityAh = 150,
  currentDeviationPercent,
  currentWorstCell = 'Cell #18',
  lang = 'vi',
  onUpdateCurrentValue,
  className = ''
}) => {
  // Normalize system tag key for dataset lookup
  const normalizedKey = useMemo(() => {
    if (systemTag.includes('UPS-DATACENTER-TIER3') || systemTag.includes('TIER3')) {
      return 'UPS-DATACENTER-TIER3-A';
    }
    return 'UPS-CRITICAL-DATACENTER-01';
  }, [systemTag]);

  // Dataset State
  const [records, setRecords] = useState<UpsBatteryResistanceRecord[]>(() => {
    return PRESET_SYSTEM_DATASETS[normalizedKey] || PRESET_SYSTEM_DATASETS['UPS-CRITICAL-DATACENTER-01'];
  });

  // Controls State
  const [viewMode, setViewMode] = useState<'resistance' | 'deviation'>('resistance');
  const [visibleSeries, setVisibleSeries] = useState<{
    worstCell: boolean;
    stringAvg: boolean;
    minCell: boolean;
    baseline: boolean;
    cell18: boolean;
    cell1: boolean;
    cell12: boolean;
  }>({
    worstCell: true,
    stringAvg: true,
    minCell: true,
    baseline: true,
    cell18: true,
    cell1: false,
    cell12: false
  });

  // Modal / Input form for new cycle
  const [showAddModal, setShowAddModal] = useState<boolean>(false);
  const [newDate, setNewDate] = useState<string>('2026-12-15');
  const [newCycleLabel, setNewCycleLabel] = useState<string>('Q4/2026 (Tái Kiểm Định)');
  const [newAvg, setNewAvg] = useState<number>(1.36);
  const [newWorst, setNewWorst] = useState<number>(1.85);
  const [newMin, setNewMin] = useState<number>(1.28);
  const [newWorstCellId, setNewWorstCellId] = useState<string>('Cell #18');
  const [newNotes, setNewNotes] = useState<string>('Đo kiểm tra lại sau khi xử lý vệ sinh và siết lực cọc bình.');

  // Baseline resistance for this system
  const baseline = records.length > 0 ? records[0].baselineResistance : 1.25;

  // Warning thresholds according to IEEE 1188
  const warningLimitResistance = baseline * 1.20; // +20%
  const criticalLimitResistance = baseline * 1.50; // +50%

  // Latest Record
  const latestRecord = records[records.length - 1];

  // Prepared Chart Data
  const chartData = useMemo(() => {
    return records.map((r) => {
      const devWorst = Number((((r.worstCellResistance - r.baselineResistance) / r.baselineResistance) * 100).toFixed(1));
      const devAvg = Number((((r.stringAvgResistance - r.baselineResistance) / r.baselineResistance) * 100).toFixed(1));
      const devMin = Number((((r.minCellResistance - r.baselineResistance) / r.baselineResistance) * 100).toFixed(1));
      const dev18 = r.cell18 ? Number((((r.cell18 - r.baselineResistance) / r.baselineResistance) * 100).toFixed(1)) : 0;
      const dev1 = r.cell1 ? Number((((r.cell1 - r.baselineResistance) / r.baselineResistance) * 100).toFixed(1)) : 0;
      const dev12 = r.cell12 ? Number((((r.cell12 - r.baselineResistance) / r.baselineResistance) * 100).toFixed(1)) : 0;

      return {
        ...r,
        displayLabel: r.cycleLabel.split(' ')[0], // e.g. "Q1/2024"
        devWorst,
        devAvg,
        devMin,
        dev18,
        dev1,
        dev12,
        warnThresholdRes: warningLimitResistance,
        critThresholdRes: criticalLimitResistance,
        warnThresholdDev: 20,
        critThresholdDev: 50
      };
    });
  }, [records, warningLimitResistance, criticalLimitResistance]);

  // Handle Add New Record
  const handleAddRecord = () => {
    const deviation = Number((((newWorst - baseline) / baseline) * 100).toFixed(1));
    let status: 'NORMAL' | 'CAUTION' | 'CRITICAL' = 'NORMAL';
    if (deviation > 30) status = 'CRITICAL';
    else if (deviation > 15) status = 'CAUTION';

    const newRec: UpsBatteryResistanceRecord = {
      id: `rec-${Date.now()}`,
      date: newDate,
      cycleLabel: newCycleLabel,
      stringAvgResistance: newAvg,
      worstCellResistance: newWorst,
      minCellResistance: newMin,
      worstCellId: newWorstCellId,
      baselineResistance: baseline,
      maxDeviationPercent: deviation,
      floatVoltageV: 13.4,
      temperatureCelsius: 25.0,
      fseInspector: 'Võ Minh Tú',
      status,
      notes: newNotes,
      cell18: newWorst,
      cell1: newMin,
      cell12: newAvg
    };

    const updated = [...records, newRec];
    setRecords(updated);
    setShowAddModal(false);

    if (onUpdateCurrentValue) {
      onUpdateCurrentValue(deviation, newWorstCellId);
    }
  };

  // Reset to Defaults
  const handleReset = () => {
    setRecords(PRESET_SYSTEM_DATASETS[normalizedKey] || PRESET_SYSTEM_DATASETS['UPS-CRITICAL-DATACENTER-01']);
  };

  // Export CSV
  const handleExportCsv = () => {
    const headers = [
      'Cycle Label',
      'Date',
      'Baseline (mOhm)',
      'String Avg (mOhm)',
      'Worst Cell (mOhm)',
      'Min Cell (mOhm)',
      'Worst Cell ID',
      'Deviation (%)',
      'Status',
      'Notes'
    ];
    const rows = records.map((r) => [
      `"${r.cycleLabel}"`,
      r.date,
      r.baselineResistance,
      r.stringAvgResistance,
      r.worstCellResistance,
      r.minCellResistance,
      `"${r.worstCellId}"`,
      r.maxDeviationPercent,
      r.status,
      `"${r.notes.replace(/"/g, '""')}"`
    ]);

    const csvContent = 'data:text/csv;charset=utf-8,' + [headers.join(','), ...rows.map((e) => e.join(','))].join('\n');
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', `UPS_Battery_Resistance_Trend_${systemTag.replace(/[^a-zA-Z0-9]/g, '_')}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Custom Tooltip
  const CustomTooltip = ({ active, payload, label }: any) => {
    if (!active || !payload || !payload.length) return null;
    const data = payload[0].payload as typeof chartData[0];

    return (
      <div className="bg-slate-900 border border-slate-700 p-3 rounded-xl shadow-2xl text-xs max-w-xs space-y-1.5 z-50">
        <div className="font-bold text-white border-b border-slate-800 pb-1 flex justify-between items-center">
          <span>{data.cycleLabel}</span>
          <span className="text-[10px] text-slate-400 font-mono">{data.date}</span>
        </div>

        <div className="space-y-1 pt-0.5">
          <div className="flex justify-between items-center text-rose-300">
            <span>Tế bào cao nhất ({data.worstCellId}):</span>
            <span className="font-bold font-mono">
              {viewMode === 'resistance' ? `${data.worstCellResistance} mΩ` : `+${data.devWorst}%`}
            </span>
          </div>

          <div className="flex justify-between items-center text-indigo-300">
            <span>Trung bình toàn chuỗi:</span>
            <span className="font-bold font-mono">
              {viewMode === 'resistance' ? `${data.stringAvgResistance} mΩ` : `+${data.devAvg}%`}
            </span>
          </div>

          <div className="flex justify-between items-center text-emerald-300">
            <span>Tế bào thấp nhất:</span>
            <span className="font-bold font-mono">
              {viewMode === 'resistance' ? `${data.minCellResistance} mΩ` : `+${data.devMin}%`}
            </span>
          </div>

          <div className="flex justify-between items-center text-slate-400 border-t border-slate-800 pt-1">
            <span>Baseline nghiệm thu gốc:</span>
            <span className="font-mono">{data.baselineResistance} mΩ</span>
          </div>

          <div className="flex justify-between items-center">
            <span className="text-slate-400">Độ lệch so với Baseline:</span>
            <span
              className={`font-black font-mono ${
                data.maxDeviationPercent > 30
                  ? 'text-rose-400'
                  : data.maxDeviationPercent > 15
                  ? 'text-amber-400'
                  : 'text-emerald-400'
              }`}
            >
              +{data.maxDeviationPercent}%
            </span>
          </div>

          <div className="text-[11px] text-slate-400 italic pt-1 border-t border-slate-800">
            {data.notes}
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className={`bg-slate-900/90 rounded-2xl border border-slate-800 p-5 shadow-xl space-y-4 ${className}`}>
      {/* Chart Header */}
      <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-3 border-b border-slate-800 pb-4">
        <div>
          <div className="flex items-center gap-2 flex-wrap mb-1">
            <span className="px-2 py-0.5 text-[10px] font-bold rounded bg-indigo-500/20 text-indigo-300 border border-indigo-400/30">
              IEEE 1188 / IEEE 450
            </span>
            <span className="px-2 py-0.5 text-[10px] font-bold rounded bg-slate-800 text-slate-300 border border-slate-700">
              ANSI/NETA MTS 7.18
            </span>
            <span
              className={`px-2 py-0.5 text-[10px] font-black rounded border ${
                latestRecord.status === 'CRITICAL'
                  ? 'bg-rose-500/20 text-rose-300 border-rose-500/40'
                  : latestRecord.status === 'CAUTION'
                  ? 'bg-amber-500/20 text-amber-300 border-amber-500/40'
                  : 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40'
              }`}
            >
              {latestRecord.status === 'CRITICAL'
                ? '🔴 NGUY HẠI NỘI TRỞ TĂNG CAO'
                : latestRecord.status === 'CAUTION'
                ? '🟡 CẢNH BÁO SUY THOÁI'
                : '🟢 ĐẠT CHUẨN IEEE 1188'}
            </span>
          </div>

          <h3 className="text-base font-bold text-white flex items-center gap-2">
            <BatteryCharging className="text-indigo-400 shrink-0" size={20} />
            <span>
              {lang === 'vi'
                ? 'Biểu Đồ Xu Hướng Lịch Sử Nội Trở Bình Ắc-quy (IEEE 1188 Battery Impedance Trend)'
                : 'Battery Internal Resistance Historical Trend (IEEE 1188)'}
            </span>
          </h3>

          <p className="text-xs text-slate-400 mt-0.5">
            {lang === 'vi'
              ? `Hệ thống: ${systemTag} | Loại bình: ${batteryType} (${batteryCapacityAh}Ah) | Baseline gốc: ${baseline.toFixed(2)} mΩ`
              : `System: ${systemTag} | Battery: ${batteryType} (${batteryCapacityAh}Ah) | Baseline: ${baseline.toFixed(2)} mΩ`}
          </p>
        </div>

        {/* Action Buttons */}
        <div className="flex items-center gap-2 flex-wrap">
          {/* View mode toggle */}
          <div className="bg-slate-950 p-1 rounded-xl border border-slate-800 flex items-center gap-1 text-xs">
            <button
              onClick={() => setViewMode('resistance')}
              className={`px-2.5 py-1 rounded-lg font-medium transition ${
                viewMode === 'resistance'
                  ? 'bg-indigo-600 text-white shadow'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              Giá Trị (mΩ)
            </button>
            <button
              onClick={() => setViewMode('deviation')}
              className={`px-2.5 py-1 rounded-lg font-medium transition ${
                viewMode === 'deviation'
                  ? 'bg-indigo-600 text-white shadow'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              Độ Lệch (%)
            </button>
          </div>

          <button
            onClick={() => setShowAddModal(true)}
            className="px-3 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold flex items-center gap-1.5 transition shadow"
          >
            <Plus size={14} />
            <span>Thêm Kỳ Đo</span>
          </button>

          <button
            onClick={handleExportCsv}
            className="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-medium border border-slate-700 flex items-center gap-1.5 transition"
            title="Xuất CSV"
          >
            <Download size={14} />
            <span>CSV</span>
          </button>

          <button
            onClick={handleReset}
            className="p-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white border border-slate-700 transition"
            title="Khôi phục mặc định"
          >
            <RotateCcw size={14} />
          </button>
        </div>
      </div>

      {/* KPI Stats Bar */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <div className="bg-slate-950/70 p-3 rounded-xl border border-slate-800">
          <div className="text-[11px] text-slate-400 flex items-center gap-1">
            <Activity size={12} className="text-rose-400" />
            <span>Tế Bào Cao Nhất ({latestRecord.worstCellId})</span>
          </div>
          <div className="text-base font-bold text-rose-300 mt-1">
            {latestRecord.worstCellResistance} mΩ
          </div>
          <div className="text-[10px] text-slate-400 mt-0.5">
            Lệch +{latestRecord.maxDeviationPercent}% so với baseline
          </div>
        </div>

        <div className="bg-slate-950/70 p-3 rounded-xl border border-slate-800">
          <div className="text-[11px] text-slate-400 flex items-center gap-1">
            <Layers size={12} className="text-indigo-400" />
            <span>Trung Bình Toàn Chuỗi</span>
          </div>
          <div className="text-base font-bold text-indigo-300 mt-1">
            {latestRecord.stringAvgResistance} mΩ
          </div>
          <div className="text-[10px] text-slate-400 mt-0.5">
            Baseline: {latestRecord.baselineResistance} mΩ
          </div>
        </div>

        <div className="bg-slate-950/70 p-3 rounded-xl border border-slate-800">
          <div className="text-[11px] text-slate-400 flex items-center gap-1">
            <AlertTriangle size={12} className="text-amber-400" />
            <span>Ngưỡng Cảnh Báo (+20%)</span>
          </div>
          <div className="text-base font-bold text-amber-300 mt-1">
            {warningLimitResistance.toFixed(2)} mΩ
          </div>
          <div className="text-[10px] text-slate-400 mt-0.5">IEEE 1188 Cảnh báo suy thoái</div>
        </div>

        <div className="bg-slate-950/70 p-3 rounded-xl border border-slate-800">
          <div className="text-[11px] text-slate-400 flex items-center gap-1">
            <ShieldAlert size={12} className="text-rose-500" />
            <span>Ngưỡng Thay Thế (+50%)</span>
          </div>
          <div className="text-base font-bold text-rose-400 mt-1">
            {criticalLimitResistance.toFixed(2)} mΩ
          </div>
          <div className="text-[10px] text-slate-400 mt-0.5">Nguy cơ nổ bình / đứt mạch</div>
        </div>
      </div>

      {/* Series Visibility Toggles */}
      <div className="flex items-center gap-2 flex-wrap text-xs text-slate-400 bg-slate-950/40 p-2 rounded-xl border border-slate-800/60">
        <span className="text-[11px] font-semibold text-slate-400 flex items-center gap-1">
          <Sliders size={12} />
          <span>Hiển thị đường xu hướng:</span>
        </span>

        <label className="flex items-center gap-1.5 cursor-pointer hover:text-white transition">
          <input
            type="checkbox"
            checked={visibleSeries.worstCell}
            onChange={(e) => setVisibleSeries({ ...visibleSeries, worstCell: e.target.checked })}
            className="rounded bg-slate-800 border-slate-700 text-rose-500 focus:ring-0"
          />
          <span className="text-rose-400 font-medium">Tế Bào Lỗi Nhất ({latestRecord.worstCellId})</span>
        </label>

        <label className="flex items-center gap-1.5 cursor-pointer hover:text-white transition">
          <input
            type="checkbox"
            checked={visibleSeries.stringAvg}
            onChange={(e) => setVisibleSeries({ ...visibleSeries, stringAvg: e.target.checked })}
            className="rounded bg-slate-800 border-slate-700 text-indigo-500 focus:ring-0"
          />
          <span className="text-indigo-400 font-medium">Trung Bình Toàn Chuỗi</span>
        </label>

        <label className="flex items-center gap-1.5 cursor-pointer hover:text-white transition">
          <input
            type="checkbox"
            checked={visibleSeries.minCell}
            onChange={(e) => setVisibleSeries({ ...visibleSeries, minCell: e.target.checked })}
            className="rounded bg-slate-800 border-slate-700 text-emerald-500 focus:ring-0"
          />
          <span className="text-emerald-400 font-medium">Tế Bào Tốt Nhất</span>
        </label>

        <label className="flex items-center gap-1.5 cursor-pointer hover:text-white transition">
          <input
            type="checkbox"
            checked={visibleSeries.baseline}
            onChange={(e) => setVisibleSeries({ ...visibleSeries, baseline: e.target.checked })}
            className="rounded bg-slate-800 border-slate-700 text-slate-400 focus:ring-0"
          />
          <span className="text-slate-400 font-medium">Baseline Nghiệm Thu Gốc</span>
        </label>
      </div>

      {/* Main Recharts Line Chart */}
      <div className="h-80 w-full pt-2">
        <ResponsiveContainer width="100%" height="100%">
          <LineChart data={chartData} margin={{ top: 15, right: 30, left: 10, bottom: 25 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#334155" opacity={0.5} />
            <XAxis
              dataKey="displayLabel"
              stroke="#94a3b8"
              fontSize={11}
              tickLine={false}
              dy={8}
            />
            <YAxis
              stroke="#94a3b8"
              fontSize={11}
              tickLine={false}
              domain={viewMode === 'resistance' ? ['auto', 'auto'] : [0, 'auto']}
              unit={viewMode === 'resistance' ? ' mΩ' : '%'}
            />
            <Tooltip content={<CustomTooltip />} />
            <Legend
              verticalAlign="top"
              height={36}
              wrapperStyle={{ fontSize: '11px', paddingTop: '0px' }}
            />

            {/* Threshold Reference Lines */}
            {viewMode === 'resistance' ? (
              <>
                <ReferenceLine
                  y={warningLimitResistance}
                  stroke="#f59e0b"
                  strokeDasharray="4 4"
                  strokeWidth={1.5}
                  label={{
                    value: `Cảnh Báo IEEE 1188 (+20%): ${warningLimitResistance.toFixed(2)} mΩ`,
                    fill: '#f59e0b',
                    fontSize: 10,
                    position: 'insideTopRight'
                  }}
                />
                <ReferenceLine
                  y={criticalLimitResistance}
                  stroke="#f43f5e"
                  strokeDasharray="4 4"
                  strokeWidth={1.5}
                  label={{
                    value: `Nguy Hiểm Thay Thế (+50%): ${criticalLimitResistance.toFixed(2)} mΩ`,
                    fill: '#f43f5e',
                    fontSize: 10,
                    position: 'insideTopRight'
                  }}
                />
              </>
            ) : (
              <>
                <ReferenceLine
                  y={20}
                  stroke="#f59e0b"
                  strokeDasharray="4 4"
                  strokeWidth={1.5}
                  label={{
                    value: 'Ngưỡng Cảnh Báo IEEE 1188 (+20%)',
                    fill: '#f59e0b',
                    fontSize: 10,
                    position: 'insideTopRight'
                  }}
                />
                <ReferenceLine
                  y={50}
                  stroke="#f43f5e"
                  strokeDasharray="4 4"
                  strokeWidth={1.5}
                  label={{
                    value: 'Ngưỡng Thay Thế Khẩn Cấp (+50%)',
                    fill: '#f43f5e',
                    fontSize: 10,
                    position: 'insideTopRight'
                  }}
                />
              </>
            )}

            {/* Lines */}
            {visibleSeries.baseline && (
              <Line
                type="monotone"
                dataKey={viewMode === 'resistance' ? 'baselineResistance' : () => 0}
                name={viewMode === 'resistance' ? 'Baseline Nghiệm Thu Gốc (mΩ)' : 'Baseline Gốc (0%)'}
                stroke="#64748b"
                strokeWidth={1.5}
                strokeDasharray="3 3"
                dot={false}
              />
            )}

            {visibleSeries.minCell && (
              <Line
                type="monotone"
                dataKey={viewMode === 'resistance' ? 'minCellResistance' : 'devMin'}
                name={viewMode === 'resistance' ? 'Tế Bào Thấp Nhất (mΩ)' : 'Tế Bào Thấp Nhất (%)'}
                stroke="#10b981"
                strokeWidth={2}
                dot={{ r: 3, fill: '#10b981' }}
                activeDot={{ r: 5 }}
              />
            )}

            {visibleSeries.stringAvg && (
              <Line
                type="monotone"
                dataKey={viewMode === 'resistance' ? 'stringAvgResistance' : 'devAvg'}
                name={viewMode === 'resistance' ? 'Trung Bình Dàn Bình (mΩ)' : 'Trung Bình Dàn Bình (%)'}
                stroke="#6366f1"
                strokeWidth={2.5}
                dot={{ r: 4, fill: '#6366f1' }}
                activeDot={{ r: 6 }}
              />
            )}

            {visibleSeries.worstCell && (
              <Line
                type="monotone"
                dataKey={viewMode === 'resistance' ? 'worstCellResistance' : 'devWorst'}
                name={
                  viewMode === 'resistance'
                    ? `Tế Bào Cao Nhất - ${latestRecord.worstCellId} (mΩ)`
                    : `Tế Bào Cao Nhất - ${latestRecord.worstCellId} (%)`
                }
                stroke="#f43f5e"
                strokeWidth={3}
                dot={{ r: 5, fill: '#f43f5e' }}
                activeDot={{ r: 7 }}
              />
            )}
          </LineChart>
        </ResponsiveContainer>
      </div>

      {/* Historical Inspection Log Table */}
      <div className="bg-slate-950/80 rounded-xl border border-slate-800 overflow-hidden">
        <div className="p-3 border-b border-slate-800 flex justify-between items-center">
          <span className="text-xs font-bold text-slate-300 flex items-center gap-1.5">
            <Calendar size={14} className="text-indigo-400" />
            <span>Nhật Ký Đo Đạc Nội Trở Các Kỳ (IEEE 1188 / NETA MTS Inspection Log)</span>
          </span>
          <span className="text-[11px] text-slate-500">{records.length} Kỳ Kiểm Định</span>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-900 text-slate-400 uppercase text-[10px] tracking-wider">
              <tr>
                <th className="py-2.5 px-3">Kỳ Kiểm Định</th>
                <th className="py-2.5 px-3">Ngày Đo</th>
                <th className="py-2.5 px-3">Trung Bình</th>
                <th className="py-2.5 px-3">Cao Nhất (Tế Bào)</th>
                <th className="py-2.5 px-3">Độ Lệch vs Baseline</th>
                <th className="py-2.5 px-3 text-center">Trạng Thái</th>
                <th className="py-2.5 px-3">Ghi Chú & Khuyến Nghị</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-800">
              {records.map((r) => (
                <tr key={r.id} className="hover:bg-slate-900/50 transition">
                  <td className="py-2 px-3 font-medium text-white">{r.cycleLabel}</td>
                  <td className="py-2 px-3 text-slate-400 font-mono text-[11px]">{r.date}</td>
                  <td className="py-2 px-3 text-indigo-300 font-mono font-semibold">
                    {r.stringAvgResistance} mΩ
                  </td>
                  <td className="py-2 px-3 text-rose-300 font-mono font-bold">
                    {r.worstCellResistance} mΩ ({r.worstCellId})
                  </td>
                  <td className="py-2 px-3 font-mono font-black">
                    <span
                      className={
                        r.maxDeviationPercent > 30
                          ? 'text-rose-400'
                          : r.maxDeviationPercent > 15
                          ? 'text-amber-400'
                          : 'text-emerald-400'
                      }
                    >
                      +{r.maxDeviationPercent}%
                    </span>
                  </td>
                  <td className="py-2 px-3 text-center">
                    <span
                      className={`px-2 py-0.5 rounded text-[10px] font-black ${
                        r.status === 'CRITICAL'
                          ? 'bg-rose-500/20 text-rose-300 border border-rose-500/40'
                          : r.status === 'CAUTION'
                          ? 'bg-amber-500/20 text-amber-300 border border-amber-500/40'
                          : 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'
                      }`}
                    >
                      {r.status}
                    </span>
                  </td>
                  <td className="py-2 px-3 text-slate-300 text-[11px] max-w-xs truncate" title={r.notes}>
                    {r.notes}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Add New Cycle Modal */}
      {showAddModal && (
        <div className="fixed inset-0 bg-black/70 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-700 rounded-2xl p-6 max-w-md w-full shadow-2xl space-y-4">
            <div className="flex justify-between items-center border-b border-slate-800 pb-3">
              <h4 className="text-sm font-bold text-white flex items-center gap-2">
                <Plus size={16} className="text-indigo-400" />
                <span>Thêm Dữ Liệu Đo Nội Trở Kỳ Mới</span>
              </h4>
              <button
                onClick={() => setShowAddModal(false)}
                className="text-slate-400 hover:text-white text-xs"
              >
                ✕
              </button>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="text-slate-400 block mb-1">Tên Kỳ Đo</label>
                <input
                  type="text"
                  value={newCycleLabel}
                  onChange={(e) => setNewCycleLabel(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-700 rounded px-2.5 py-1.5 text-white"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-slate-400 block mb-1">Ngày Đo</label>
                  <input
                    type="date"
                    value={newDate}
                    onChange={(e) => setNewDate(e.target.value)}
                    className="w-full bg-slate-950 border border-slate-700 rounded px-2.5 py-1.5 text-white"
                  />
                </div>
                <div>
                  <label className="text-slate-400 block mb-1">Tế Bào Lỗi Nhất</label>
                  <input
                    type="text"
                    value={newWorstCellId}
                    onChange={(e) => setNewWorstCellId(e.target.value)}
                    className="w-full bg-slate-950 border border-slate-700 rounded px-2.5 py-1.5 text-white"
                  />
                </div>
              </div>

              <div className="grid grid-cols-3 gap-2">
                <div>
                  <label className="text-slate-400 block mb-1">Trung Bình (mΩ)</label>
                  <input
                    type="number"
                    step="0.01"
                    value={newAvg}
                    onChange={(e) => setNewAvg(Number(e.target.value))}
                    className="w-full bg-slate-950 border border-slate-700 rounded px-2.5 py-1.5 text-white"
                  />
                </div>
                <div>
                  <label className="text-slate-400 block mb-1">Lớn Nhất (mΩ)</label>
                  <input
                    type="number"
                    step="0.01"
                    value={newWorst}
                    onChange={(e) => setNewWorst(Number(e.target.value))}
                    className="w-full bg-slate-950 border border-slate-700 rounded px-2.5 py-1.5 text-white"
                  />
                </div>
                <div>
                  <label className="text-slate-400 block mb-1">Nhỏ Nhất (mΩ)</label>
                  <input
                    type="number"
                    step="0.01"
                    value={newMin}
                    onChange={(e) => setNewMin(Number(e.target.value))}
                    className="w-full bg-slate-950 border border-slate-700 rounded px-2.5 py-1.5 text-white"
                  />
                </div>
              </div>

              <div>
                <label className="text-slate-400 block mb-1">Ghi Chú Đánh Giá Hiện Trường</label>
                <textarea
                  value={newNotes}
                  onChange={(e) => setNewNotes(e.target.value)}
                  rows={3}
                  className="w-full bg-slate-950 border border-slate-700 rounded px-2.5 py-1.5 text-white"
                />
              </div>
            </div>

            <div className="flex justify-end gap-2 border-t border-slate-800 pt-3">
              <button
                onClick={() => setShowAddModal(false)}
                className="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs"
              >
                Hủy
              </button>
              <button
                onClick={handleAddRecord}
                className="px-4 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white font-semibold text-xs shadow"
              >
                Lưu Vào Biểu Đồ
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
