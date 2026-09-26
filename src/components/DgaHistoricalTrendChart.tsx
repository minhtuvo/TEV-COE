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
  Plus,
  Trash2,
  RotateCcw,
  Download,
  AlertTriangle,
  Info,
  Calendar,
  CheckCircle2,
  Layers,
  ArrowUpRight,
  Flame,
  Zap,
  Activity,
  Maximize2,
  Eye,
  EyeOff
} from 'lucide-react';

export interface DgaHistoryPoint {
  id: string;
  date: string; // YYYY-MM-DD or DD/MM/YYYY
  label?: string; // e.g. "Kỳ 1/2025"
  h2: number;
  ch4: number;
  c2h6: number;
  c2h4: number;
  c2h2: number;
  co: number;
  co2: number;
  samplingPoint?: string;
  notes?: string;
}

interface DgaHistoricalTrendChartProps {
  currentFormValues?: {
    h2?: string | number;
    ch4?: string | number;
    c2h6?: string | number;
    c2h4?: string | number;
    c2h2?: string | number;
    co?: string | number;
    co2?: string | number;
    o2?: string | number;
    n2?: string | number;
  };
  equipmentCode?: string;
  equipmentName?: string;
  onLoadSampleToForm?: (sample: Partial<DgaHistoryPoint>) => void;
  onAnalyzeWithSample?: (sample: Partial<DgaHistoryPoint>) => void;
}

// Baseline historical samples for standard periodic transformer testing
const DEFAULT_HISTORICAL_SAMPLES: DgaHistoryPoint[] = [
  {
    id: 'dga-hist-1',
    date: '2025-03-15',
    label: 'Kỳ 1 (15/03/2025)',
    h2: 18,
    ch4: 22,
    c2h6: 14,
    c2h4: 10,
    c2h2: 0.1,
    co: 220,
    co2: 1950,
    samplingPoint: 'Van đáy MBA',
    notes: 'Kiểm tra định kỳ sau 6 tháng vận hành. Dầu trong giới hạn bình thường.'
  },
  {
    id: 'dga-hist-2',
    date: '2025-09-20',
    label: 'Kỳ 2 (20/09/2025)',
    h2: 32,
    ch4: 38,
    c2h6: 25,
    c2h4: 18,
    c2h2: 0.2,
    co: 295,
    co2: 2400,
    samplingPoint: 'Van đáy MBA',
    notes: 'Phụ tải trung bình 70%, xuất hiện xu hướng sinh khí nhiệt nhẹ.'
  },
  {
    id: 'dga-hist-3',
    date: '2026-03-10',
    label: 'Kỳ 3 (10/03/2026)',
    h2: 68,
    ch4: 82,
    c2h6: 48,
    c2h4: 35,
    c2h2: 0.7,
    co: 420,
    co2: 3250,
    samplingPoint: 'Van đáy MBA',
    notes: 'Tốc độ sinh CH4 và C2H4 tăng đều. Cần theo dõi chặt chẽ.'
  },
  {
    id: 'dga-hist-4',
    date: '2026-09-24',
    label: 'Kỳ 4 (Hiện tại - 24/09/2026)',
    h2: 145,
    ch4: 172,
    c2h6: 98,
    c2h4: 92,
    c2h2: 2.8,
    co: 590,
    co2: 4400,
    samplingPoint: 'Van đáy MBA',
    notes: 'Cảnh báo: C2H2 tăng vọt lên 2.8 ppm (vượt ngưỡng 1.0 ppm), xuất hiện phóng điện hồ quang.'
  }
];

// Gas configuration: colors, symbols, standard IEEE limits (90th percentile IEEE C57.104-2008 / 2019)
export const GAS_CONFIGS: Record<string, {
  name: string;
  formula: string;
  fullName: string;
  color: string;
  unit: string;
  limit90th: number; // IEEE Condition 1 limit in ppm
  hazardType: string;
  yAxisId: 'left' | 'right';
}> = {
  h2: {
    name: 'H2',
    formula: 'H₂',
    fullName: 'Hydrogen (Hydro)',
    color: '#2563eb', // Blue
    unit: 'ppm',
    limit90th: 100,
    hazardType: 'Phóng điện vầng quang / PD',
    yAxisId: 'left'
  },
  ch4: {
    name: 'CH4',
    formula: 'CH₄',
    fullName: 'Methane (Metan)',
    color: '#059669', // Emerald
    unit: 'ppm',
    limit90th: 120,
    hazardType: 'Quá nhiệt nhiệt độ thấp (< 300°C)',
    yAxisId: 'left'
  },
  c2h6: {
    name: 'C2H6',
    formula: 'C₂H₆',
    fullName: 'Ethane (Etan)',
    color: '#d97706', // Amber
    unit: 'ppm',
    limit90th: 65,
    hazardType: 'Quá nhiệt nhiệt độ trung bình (300-500°C)',
    yAxisId: 'left'
  },
  c2h4: {
    name: 'C2H4',
    formula: 'C₂H₄',
    fullName: 'Ethylene (Etilen)',
    color: '#ea580c', // Orange
    unit: 'ppm',
    limit90th: 50,
    hazardType: 'Quá nhiệt nhiệt độ cao (> 700°C)',
    yAxisId: 'left'
  },
  c2h2: {
    name: 'C2H2',
    formula: 'C₂H₂',
    fullName: 'Acetylene (Axetilen)',
    color: '#dc2626', // Bright Red
    unit: 'ppm',
    limit90th: 1, // Extremely strict: > 1 ppm is alarming!
    hazardType: 'Hồ quang điện năng lượng cao (Arcing)',
    yAxisId: 'left'
  },
  co: {
    name: 'CO',
    formula: 'CO',
    fullName: 'Carbon Monoxide',
    color: '#7c3aed', // Purple
    unit: 'ppm',
    limit90th: 350,
    hazardType: 'Lão hóa cách điện Cellulose',
    yAxisId: 'left'
  },
  co2: {
    name: 'CO2',
    formula: 'CO₂',
    fullName: 'Carbon Dioxide',
    color: '#64748b', // Slate
    unit: 'ppm',
    limit90th: 2500,
    hazardType: 'Suy thoái cách điện rắn Cellulose',
    yAxisId: 'right' // Put on right axis due to thousands ppm scale
  },
  tdcg: {
    name: 'TDCG',
    formula: 'TDCG',
    fullName: 'Total Combustible Gas (Tổng khí cháy)',
    color: '#4338ca', // Indigo
    unit: 'ppm',
    limit90th: 720,
    hazardType: 'Tổng hợp khí cháy IEEE C57.104',
    yAxisId: 'left'
  }
};

export const DgaHistoricalTrendChart: React.FC<DgaHistoricalTrendChartProps> = ({
  currentFormValues,
  equipmentCode = 'TRF-01',
  equipmentName = 'Máy biến áp T1',
  onLoadSampleToForm,
  onAnalyzeWithSample
}) => {
  // Historical data state initialized from localStorage or defaults
  const [history, setHistory] = useState<DgaHistoryPoint[]>(() => {
    const saved = localStorage.getItem(`dga_history_${equipmentCode}`);
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        console.warn('Failed to parse saved dga history', e);
      }
    }
    return DEFAULT_HISTORICAL_SAMPLES;
  });

  // Active gas visibility filters
  const [activeGases, setActiveGases] = useState<Record<string, boolean>>({
    h2: true,
    ch4: true,
    c2h6: true,
    c2h4: true,
    c2h2: true,
    co: true,
    co2: false, // Default hidden to keep combustible gases clearly legible
    tdcg: true
  });

  // Reference lines toggle
  const [showThresholdLines, setShowThresholdLines] = useState<boolean>(true);

  // Separate CO2 / Right Y-axis
  const [useSeparateCo2Axis, setUseSeparateCo2Axis] = useState<boolean>(true);

  // New Sample modal / drawer state
  const [showAddModal, setShowAddModal] = useState<boolean>(false);
  const [newSample, setNewSample] = useState<Partial<DgaHistoryPoint>>({
    date: new Date().toISOString().split('T')[0],
    h2: 0,
    ch4: 0,
    c2h6: 0,
    c2h4: 0,
    c2h2: 0,
    co: 0,
    co2: 0,
    samplingPoint: 'Van đáy MBA',
    notes: ''
  });

  // Save to localStorage when history changes
  const saveHistory = (updated: DgaHistoryPoint[]) => {
    setHistory(updated);
    try {
      localStorage.setItem(`dga_history_${equipmentCode}`, JSON.stringify(updated));
    } catch (e) {
      console.warn('Could not save DGA history to localStorage', e);
    }
  };

  // Toggle gas visibility
  const toggleGas = (gasKey: string) => {
    setActiveGases(prev => ({
      ...prev,
      [gasKey]: !prev[gasKey]
    }));
  };

  // Select all or combustible only
  const setAllGases = (state: boolean) => {
    setActiveGases({
      h2: state,
      ch4: state,
      c2h6: state,
      c2h4: state,
      c2h2: state,
      co: state,
      co2: state,
      tdcg: state
    });
  };

  const setOnlyCombustibleGases = () => {
    setActiveGases({
      h2: true,
      ch4: true,
      c2h6: true,
      c2h4: true,
      c2h2: true,
      co: true,
      co2: false,
      tdcg: true
    });
  };

  // Add current form values as a new historical point
  const handleAddCurrentFormAsHistory = () => {
    if (!currentFormValues) return;
    const h2 = parseFloat(String(currentFormValues.h2 || 0)) || 0;
    const ch4 = parseFloat(String(currentFormValues.ch4 || 0)) || 0;
    const c2h6 = parseFloat(String(currentFormValues.c2h6 || 0)) || 0;
    const c2h4 = parseFloat(String(currentFormValues.c2h4 || 0)) || 0;
    const c2h2 = parseFloat(String(currentFormValues.c2h2 || 0)) || 0;
    const co = parseFloat(String(currentFormValues.co || 0)) || 0;
    const co2 = parseFloat(String(currentFormValues.co2 || 0)) || 0;

    const todayStr = new Date().toISOString().split('T')[0];
    const newPoint: DgaHistoryPoint = {
      id: `dga-hist-${Date.now()}`,
      date: todayStr,
      label: `Lần đo ${new Date().toLocaleDateString('vi-VN')}`,
      h2,
      ch4,
      c2h6,
      c2h4,
      c2h2,
      co,
      co2,
      samplingPoint: 'Van đáy MBA',
      notes: `Dữ liệu đo hiện trường nạp lúc ${new Date().toLocaleTimeString('vi-VN')}`
    };

    const updated = [...history, newPoint].sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
    saveHistory(updated);
  };

  // Handle add custom sample
  const handleSaveCustomSample = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newSample.date) return;
    const item: DgaHistoryPoint = {
      id: `dga-hist-${Date.now()}`,
      date: newSample.date,
      label: newSample.label || `Lần đo ${newSample.date}`,
      h2: Number(newSample.h2 || 0),
      ch4: Number(newSample.ch4 || 0),
      c2h6: Number(newSample.c2h6 || 0),
      c2h4: Number(newSample.c2h4 || 0),
      c2h2: Number(newSample.c2h2 || 0),
      co: Number(newSample.co || 0),
      co2: Number(newSample.co2 || 0),
      samplingPoint: newSample.samplingPoint || 'Van đáy MBA',
      notes: newSample.notes || ''
    };

    const updated = [...history, item].sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
    saveHistory(updated);
    setShowAddModal(false);
  };

  // Reset to default sample history
  const handleResetDefaults = () => {
    if (window.confirm('Khôi phục danh sách dữ liệu lịch sử đo mẫu chuẩn theo IEEE C57.104?')) {
      saveHistory(DEFAULT_HISTORICAL_SAMPLES);
    }
  };

  // Delete a sample
  const handleDeleteSample = (id: string) => {
    const updated = history.filter(item => item.id !== id);
    saveHistory(updated);
  };

  // Prepare chart dataset with computed TDCG and formatted labels
  const chartData = useMemo(() => {
    return history.map(item => {
      const tdcg = item.h2 + item.ch4 + item.c2h6 + item.c2h4 + item.c2h2 + item.co;
      return {
        ...item,
        tdcg,
        displayDate: item.date.split('-').reverse().join('/') // DD/MM/YYYY
      };
    });
  }, [history]);

  // Compute generation rates (tốc độ sinh khí ppm/tháng) between the two most recent samples
  const generationRateAnalysis = useMemo(() => {
    if (chartData.length < 2) return null;
    const prev = chartData[chartData.length - 2];
    const latest = chartData[chartData.length - 1];

    const dPrev = new Date(prev.date).getTime();
    const dLatest = new Date(latest.date).getTime();
    const diffDays = Math.max(1, Math.round((dLatest - dPrev) / (1000 * 60 * 60 * 24)));
    const diffMonths = Math.max(0.1, diffDays / 30.4375);

    const deltaC2H2 = (latest.c2h2 - prev.c2h2);
    const rateC2H2PerMonth = (deltaC2H2 / diffMonths);

    const deltaTdcg = (latest.tdcg - prev.tdcg);
    const rateTdcgPerMonth = (deltaTdcg / diffMonths);

    const deltaH2 = (latest.h2 - prev.h2);
    const rateH2PerMonth = (deltaH2 / diffMonths);

    const deltaC2H4 = (latest.c2h4 - prev.c2h4);
    const rateC2H4PerMonth = (deltaC2H4 / diffMonths);

    // IEEE C57.104-2008 Table 3 Rate limits (Condition 1 limit approx 10 ppm/day or 300 ppm/month)
    const isRapidGeneration = rateTdcgPerMonth > 90 || rateC2H2PerMonth > 0.5;

    return {
      prevDate: prev.displayDate,
      latestDate: latest.displayDate,
      diffDays,
      diffMonths: diffMonths.toFixed(1),
      deltaTdcg: deltaTdcg.toFixed(1),
      rateTdcgPerMonth: rateTdcgPerMonth.toFixed(1),
      deltaC2H2: deltaC2H2.toFixed(2),
      rateC2H2PerMonth: rateC2H2PerMonth.toFixed(2),
      deltaH2: deltaH2.toFixed(1),
      rateH2PerMonth: rateH2PerMonth.toFixed(1),
      deltaC2H4: deltaC2H4.toFixed(1),
      rateC2H4PerMonth: rateC2H4PerMonth.toFixed(1),
      isRapidGeneration,
      latestTdcg: latest.tdcg,
      latestC2H2: latest.c2h2
    };
  }, [chartData]);

  // Check if current form values have any gas entered
  const hasCurrentValues = useMemo(() => {
    if (!currentFormValues) return false;
    const total = ['h2', 'ch4', 'c2h6', 'c2h4', 'c2h2', 'co', 'co2']
      .reduce((sum, key) => sum + (parseFloat(String(currentFormValues[key as keyof typeof currentFormValues] || 0)) || 0), 0);
    return total > 0;
  }, [currentFormValues]);

  return (
    <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden mb-8">
      {/* Header Bar */}
      <div className="p-6 border-b border-slate-100 bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2.5">
              <div className="p-2 rounded-lg bg-blue-500/20 text-blue-400 border border-blue-500/30">
                <TrendingUp size={22} />
              </div>
              <div>
                <h3 className="text-xl font-bold tracking-tight text-white flex items-center gap-2">
                  Biểu Đồ Xu Hướng Nồng Độ Khí Hòa Tan (DGA LineChart)
                  <span className="text-xs font-semibold px-2 py-0.5 rounded-full bg-blue-500/20 text-blue-300 border border-blue-400/30">
                    IEEE C57.104
                  </span>
                </h3>
                <p className="text-xs text-slate-300 mt-0.5">
                  Giám sát tiến triển nồng độ các loại khí cháy &amp; tốc độ sinh khí theo thời gian cho thiết bị: <span className="font-semibold text-amber-300">{equipmentCode} - {equipmentName}</span>
                </p>
              </div>
            </div>
          </div>

          <div className="flex items-center flex-wrap gap-2">
            {hasCurrentValues && (
              <button
                type="button"
                onClick={handleAddCurrentFormAsHistory}
                className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-emerald-500 text-white hover:bg-emerald-600 transition-colors shadow-sm"
                title="Ghi nhận giá trị đang nhập trên form thành điểm đo lịch sử mới nhất"
              >
                <Plus size={14} />
                Lưu Dữ Liệu Form Vào Biểu Đồ
              </button>
            )}

            <button
              type="button"
              onClick={() => setShowAddModal(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-lg bg-white/10 hover:bg-white/20 text-white transition-colors border border-white/20"
            >
              <Plus size={14} />
              Thêm Lần Đo Mới
            </button>

            <button
              type="button"
              onClick={handleResetDefaults}
              className="flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium rounded-lg bg-white/10 hover:bg-white/20 text-slate-200 transition-colors"
              title="Đặt lại 4 mẫu đo tiêu chuẩn theo chu kỳ"
            >
              <RotateCcw size={13} />
              Mẫu Chuẩn
            </button>
          </div>
        </div>

        {/* Gas Filter Pills */}
        <div className="mt-5 pt-4 border-t border-slate-700/60 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex items-center flex-wrap gap-2">
            <span className="text-slate-400 font-medium mr-1 flex items-center gap-1">
              <Layers size={13} /> Khí hiển thị:
            </span>
            {Object.keys(GAS_CONFIGS).map(gasKey => {
              const cfg = GAS_CONFIGS[gasKey];
              const isActive = activeGases[gasKey];
              return (
                <button
                  key={gasKey}
                  type="button"
                  onClick={() => toggleGas(gasKey)}
                  className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-bold transition-all border ${
                    isActive
                      ? 'bg-slate-800 text-white shadow-sm border-slate-600'
                      : 'bg-slate-900/60 text-slate-400 border-slate-800 opacity-60 hover:opacity-90'
                  }`}
                  style={{
                    borderLeftColor: cfg.color,
                    borderLeftWidth: '3px'
                  }}
                  title={`${cfg.fullName} - Ngưỡng IEEE 90th: ${cfg.limit90th} ppm`}
                >
                  <span
                    className="w-2 h-2 rounded-full inline-block"
                    style={{ backgroundColor: isActive ? cfg.color : '#64748b' }}
                  />
                  <span>{cfg.formula}</span>
                  {isActive ? <Eye size={12} className="opacity-70" /> : <EyeOff size={12} className="opacity-40" />}
                </button>
              );
            })}
          </div>

          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={setOnlyCombustibleGases}
              className="text-slate-300 hover:text-white underline text-[11px]"
            >
              Chỉ khí cháy (Ẩn CO₂)
            </button>
            <button
              type="button"
              onClick={() => setAllGases(true)}
              className="text-slate-300 hover:text-white underline text-[11px]"
            >
              Hiện tất cả
            </button>
            <label className="flex items-center gap-1.5 cursor-pointer text-slate-300 text-[11px]">
              <input
                type="checkbox"
                checked={showThresholdLines}
                onChange={e => setShowThresholdLines(e.target.checked)}
                className="rounded border-slate-700 text-blue-500 focus:ring-0"
              />
              <span>Đường giới hạn IEEE C57.104</span>
            </label>
          </div>
        </div>
      </div>

      {/* Main Chart Area */}
      <div className="p-6">
        {chartData.length === 0 ? (
          <div className="h-72 flex flex-col items-center justify-center text-slate-400 bg-slate-50 rounded-xl border border-dashed border-slate-300">
            <AlertTriangle size={36} className="text-amber-500 mb-2" />
            <p className="font-semibold text-slate-700">Chưa có dữ liệu lịch sử đo DGA</p>
            <p className="text-xs text-slate-500 mt-1 max-w-sm text-center">
              Nhấn "Mẫu Chuẩn" để tải dữ liệu kiểm tra 4 kỳ hoặc nhập các mẫu phân tích phòng thí nghiệm ngoài site.
            </p>
            <button
              type="button"
              onClick={handleResetDefaults}
              className="mt-3 px-3 py-1.5 bg-slate-900 text-white text-xs font-semibold rounded-lg hover:bg-slate-800"
            >
              Nạp Mẫu Lịch Sử Chuẩn
            </button>
          </div>
        ) : (
          <div className="w-full">
            {/* High-Alert Warning Banner if C2H2 > 1 or TDCG > 1920 */}
            {generationRateAnalysis && generationRateAnalysis.latestC2H2 > 1.0 && (
              <div className="mb-5 p-3.5 bg-rose-50 border border-rose-200 rounded-lg flex items-start gap-3 text-rose-800 text-xs animate-pulse">
                <Flame size={18} className="text-rose-600 mt-0.5 flex-shrink-0" />
                <div>
                  <div className="font-bold uppercase tracking-wider text-rose-900">
                    Cảnh Báo Nguy Hiểm: Phát Hiện Khí Acetylene (C₂H₂ = {generationRateAnalysis.latestC2H2} ppm &gt; 1.0 ppm)
                  </div>
                  <div className="mt-0.5 text-rose-700">
                    Theo IEEE C57.104, Acetylene là dấu hiệu đặc trưng của hiện tượng <strong>hồ quang điện (Electric Arcing)</strong> phá hủy cách điện bên trong máy biến áp. Yêu cầu dừng máy để đo điện áp đánh thủng dầu, kiểm tra bộ chuyển nấc OLTC và soi nội soi thùng máy.
                  </div>
                </div>
              </div>
            )}

            {/* Responsive Recharts LineChart */}
            <div className="h-80 w-full">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart
                  data={chartData}
                  margin={{ top: 15, right: 30, left: 10, bottom: 20 }}
                >
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                  
                  <XAxis
                    dataKey="displayDate"
                    tick={{ fill: '#475569', fontSize: 12, fontWeight: 500 }}
                    axisLine={{ stroke: '#cbd5e1' }}
                    tickLine={{ stroke: '#cbd5e1' }}
                    dy={8}
                  />

                  {/* Left Y-Axis for Combustible gases (ppm) */}
                  <YAxis
                    yAxisId="left"
                    tick={{ fill: '#475569', fontSize: 11 }}
                    axisLine={{ stroke: '#cbd5e1' }}
                    tickLine={{ stroke: '#cbd5e1' }}
                    label={{
                      value: 'Nồng độ Khí cháy (ppm)',
                      angle: -90,
                      position: 'insideLeft',
                      fill: '#64748b',
                      fontSize: 12,
                      offset: -2
                    }}
                  />

                  {/* Right Y-Axis for CO2 (ppm) if active */}
                  {activeGases.co2 && (
                    <YAxis
                      yAxisId="right"
                      orientation="right"
                      tick={{ fill: '#64748b', fontSize: 11 }}
                      axisLine={{ stroke: '#94a3b8' }}
                      tickLine={{ stroke: '#94a3b8' }}
                      label={{
                        value: 'CO₂ (ppm)',
                        angle: 90,
                        position: 'insideRight',
                        fill: '#64748b',
                        fontSize: 12,
                        offset: 4
                      }}
                    />
                  )}

                  {/* Custom Tooltip */}
                  <Tooltip
                    content={({ active, payload, label }) => {
                      if (!active || !payload || !payload.length) return null;
                      const point = payload[0]?.payload as (DgaHistoryPoint & { tdcg: number; displayDate: string });
                      return (
                        <div className="bg-slate-900 text-white p-3.5 rounded-xl shadow-xl border border-slate-700 text-xs min-w-[240px]">
                          <div className="flex items-center justify-between border-b border-slate-700 pb-2 mb-2">
                            <span className="font-bold text-amber-400 flex items-center gap-1.5">
                              <Calendar size={13} /> {point.displayDate}
                            </span>
                            <span className="text-[11px] text-slate-400">{point.samplingPoint || 'Van đáy MBA'}</span>
                          </div>

                          <div className="space-y-1.5">
                            {Object.keys(GAS_CONFIGS).map(gasKey => {
                              if (!activeGases[gasKey]) return null;
                              const cfg = GAS_CONFIGS[gasKey];
                              const val = gasKey === 'tdcg' ? point.tdcg : (point as any)[gasKey];
                              const isOver = val > cfg.limit90th;
                              return (
                                <div key={gasKey} className="flex items-center justify-between">
                                  <span className="flex items-center gap-1.5">
                                    <span className="w-2 h-2 rounded-full" style={{ backgroundColor: cfg.color }} />
                                    <span className="text-slate-300">{cfg.formula}:</span>
                                  </span>
                                  <span className={`font-mono font-bold ${isOver ? 'text-rose-400' : 'text-white'}`}>
                                    {val} {cfg.unit}
                                    {isOver && <span className="ml-1 text-[10px] text-rose-400 font-semibold">(Vượt IEEE)</span>}
                                  </span>
                                </div>
                              );
                            })}
                          </div>

                          {point.notes && (
                            <div className="mt-2 pt-2 border-t border-slate-800 text-[11px] text-slate-300 italic">
                              "{point.notes}"
                            </div>
                          )}
                        </div>
                      );
                    }}
                  />

                  {/* Standard IEEE C57.104 Reference Lines */}
                  {showThresholdLines && (
                    <>
                      {activeGases.c2h2 && (
                        <ReferenceLine
                          yAxisId="left"
                          y={1}
                          stroke="#ef4444"
                          strokeDasharray="4 4"
                          strokeWidth={1.5}
                          label={{
                            value: 'IEEE C2H2 Giới hạn: 1 ppm',
                            fill: '#dc2626',
                            fontSize: 10,
                            position: 'insideTopRight'
                          }}
                        />
                      )}
                      {activeGases.tdcg && (
                        <ReferenceLine
                          yAxisId="left"
                          y={720}
                          stroke="#4338ca"
                          strokeDasharray="3 3"
                          strokeWidth={1}
                          label={{
                            value: 'TDCG Cond 1 Limit: 720 ppm',
                            fill: '#4338ca',
                            fontSize: 10,
                            position: 'insideBottomRight'
                          }}
                        />
                      )}
                      {activeGases.h2 && (
                        <ReferenceLine
                          yAxisId="left"
                          y={100}
                          stroke="#2563eb"
                          strokeDasharray="2 4"
                          strokeWidth={1}
                          label={{
                            value: 'H2 Limit: 100 ppm',
                            fill: '#2563eb',
                            fontSize: 10,
                            position: 'insideTopLeft'
                          }}
                        />
                      )}
                    </>
                  )}

                  {/* Render Lines for Each Active Gas */}
                  {activeGases.h2 && (
                    <Line
                      yAxisId="left"
                      type="monotone"
                      dataKey="h2"
                      name="H2 (Hydro)"
                      stroke={GAS_CONFIGS.h2.color}
                      strokeWidth={2.5}
                      dot={{ r: 4, fill: GAS_CONFIGS.h2.color }}
                      activeDot={{ r: 6 }}
                    />
                  )}
                  {activeGases.ch4 && (
                    <Line
                      yAxisId="left"
                      type="monotone"
                      dataKey="ch4"
                      name="CH4 (Metan)"
                      stroke={GAS_CONFIGS.ch4.color}
                      strokeWidth={2.5}
                      dot={{ r: 4, fill: GAS_CONFIGS.ch4.color }}
                      activeDot={{ r: 6 }}
                    />
                  )}
                  {activeGases.c2h6 && (
                    <Line
                      yAxisId="left"
                      type="monotone"
                      dataKey="c2h6"
                      name="C2H6 (Etan)"
                      stroke={GAS_CONFIGS.c2h6.color}
                      strokeWidth={2}
                      dot={{ r: 4, fill: GAS_CONFIGS.c2h6.color }}
                      activeDot={{ r: 6 }}
                    />
                  )}
                  {activeGases.c2h4 && (
                    <Line
                      yAxisId="left"
                      type="monotone"
                      dataKey="c2h4"
                      name="C2H4 (Etilen)"
                      stroke={GAS_CONFIGS.c2h4.color}
                      strokeWidth={2.5}
                      dot={{ r: 4, fill: GAS_CONFIGS.c2h4.color }}
                      activeDot={{ r: 6 }}
                    />
                  )}
                  {activeGases.c2h2 && (
                    <Line
                      yAxisId="left"
                      type="monotone"
                      dataKey="c2h2"
                      name="C2H2 (Axetilen)"
                      stroke={GAS_CONFIGS.c2h2.color}
                      strokeWidth={3}
                      dot={{ r: 5, fill: GAS_CONFIGS.c2h2.color }}
                      activeDot={{ r: 7 }}
                    />
                  )}
                  {activeGases.co && (
                    <Line
                      yAxisId="left"
                      type="monotone"
                      dataKey="co"
                      name="CO (Carbon Monoxide)"
                      stroke={GAS_CONFIGS.co.color}
                      strokeWidth={2}
                      dot={{ r: 4, fill: GAS_CONFIGS.co.color }}
                      activeDot={{ r: 6 }}
                    />
                  )}
                  {activeGases.co2 && (
                    <Line
                      yAxisId="right"
                      type="monotone"
                      dataKey="co2"
                      name="CO2 (Carbon Dioxide)"
                      stroke={GAS_CONFIGS.co2.color}
                      strokeWidth={2}
                      strokeDasharray="4 2"
                      dot={{ r: 4, fill: GAS_CONFIGS.co2.color }}
                      activeDot={{ r: 6 }}
                    />
                  )}
                  {activeGases.tdcg && (
                    <Line
                      yAxisId="left"
                      type="monotone"
                      dataKey="tdcg"
                      name="TDCG (Tổng khí cháy)"
                      stroke={GAS_CONFIGS.tdcg.color}
                      strokeWidth={3}
                      dot={{ r: 4, fill: GAS_CONFIGS.tdcg.color }}
                      activeDot={{ r: 6 }}
                    />
                  )}
                </LineChart>
              </ResponsiveContainer>
            </div>

            {/* Rate of Gas Generation KPI Cards */}
            {generationRateAnalysis && (
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mt-6 pt-4 border-t border-slate-100">
                <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
                  <div className="text-[11px] font-bold text-slate-500 uppercase tracking-wider">Khoảng thời gian đo</div>
                  <div className="text-sm font-bold text-slate-900 mt-0.5">
                    {generationRateAnalysis.prevDate} → {generationRateAnalysis.latestDate}
                  </div>
                  <div className="text-[11px] text-slate-500 mt-0.5">
                    {generationRateAnalysis.diffDays} ngày ({generationRateAnalysis.diffMonths} tháng)
                  </div>
                </div>

                <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
                  <div className="text-[11px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1">
                    <TrendingUp size={12} className="text-blue-500" /> Tốc độ sinh TDCG
                  </div>
                  <div className={`text-base font-extrabold mt-0.5 ${Number(generationRateAnalysis.rateTdcgPerMonth) > 60 ? 'text-amber-600' : 'text-slate-900'}`}>
                    +{generationRateAnalysis.rateTdcgPerMonth} <span className="text-xs font-normal text-slate-500">ppm/tháng</span>
                  </div>
                  <div className="text-[11px] text-slate-500">
                    Δ TDCG: +{generationRateAnalysis.deltaTdcg} ppm
                  </div>
                </div>

                <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
                  <div className="text-[11px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1">
                    <Flame size={12} className="text-rose-500" /> Tốc độ sinh C₂H₂ (Hồ quang)
                  </div>
                  <div className={`text-base font-extrabold mt-0.5 ${Number(generationRateAnalysis.rateC2H2PerMonth) > 0.2 ? 'text-rose-600' : 'text-slate-900'}`}>
                    +{generationRateAnalysis.rateC2H2PerMonth} <span className="text-xs font-normal text-slate-500">ppm/tháng</span>
                  </div>
                  <div className="text-[11px] text-slate-500">
                    Δ C₂H₂: +{generationRateAnalysis.deltaC2H2} ppm
                  </div>
                </div>

                <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
                  <div className="text-[11px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1">
                    <Zap size={12} className="text-indigo-500" /> Tốc độ sinh C₂H₄ (Nhiệt cao)
                  </div>
                  <div className="text-base font-extrabold text-slate-900 mt-0.5">
                    +{generationRateAnalysis.rateC2H4PerMonth} <span className="text-xs font-normal text-slate-500">ppm/tháng</span>
                  </div>
                  <div className="text-[11px] text-slate-500">
                    Δ C₂H₄: +{generationRateAnalysis.deltaC2H4} ppm
                  </div>
                </div>
              </div>
            )}
          </div>
        )}
      </div>

      {/* Historical Data Table with 1-Click Load into Analysis Form */}
      <div className="border-t border-slate-200 bg-slate-50/70 p-6">
        <div className="flex items-center justify-between mb-3">
          <div>
            <h4 className="text-sm font-bold text-slate-900 flex items-center gap-2">
              <Calendar size={16} className="text-blue-600" />
              Bảng Số Liệu Các Lần Lấy Mẫu ({history.length} mốc kiểm tra)
            </h4>
            <p className="text-xs text-slate-500">
              Nhấn <strong>"Nạp vào Form"</strong> để chuyển số liệu của bất kỳ lần đo nào vào bộ chẩn đoán Duval &amp; Health Index.
            </p>
          </div>
        </div>

        <div className="overflow-x-auto rounded-lg border border-slate-200 bg-white shadow-sm">
          <table className="w-full text-xs text-left">
            <thead className="bg-slate-100/90 text-slate-700 font-semibold uppercase text-[10px] tracking-wider border-b border-slate-200">
              <tr>
                <th className="py-2.5 px-3">Ngày Lấy Mẫu</th>
                <th className="py-2.5 px-2 text-center text-blue-700">H₂</th>
                <th className="py-2.5 px-2 text-center text-emerald-700">CH₄</th>
                <th className="py-2.5 px-2 text-center text-amber-700">C₂H₆</th>
                <th className="py-2.5 px-2 text-center text-orange-700">C₂H₄</th>
                <th className="py-2.5 px-2 text-center text-rose-700 font-bold">C₂H₂</th>
                <th className="py-2.5 px-2 text-center text-purple-700">CO</th>
                <th className="py-2.5 px-2 text-center text-slate-600">CO₂</th>
                <th className="py-2.5 px-2 text-center text-indigo-700 font-bold">TDCG</th>
                <th className="py-2.5 px-3">Ghi Chú Hiện Trường</th>
                <th className="py-2.5 px-3 text-right">Thao Tác</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {chartData.map((item, idx) => {
                const isC2H2Alarm = item.c2h2 >= 1.0;
                return (
                  <tr key={item.id} className="hover:bg-slate-50/80 transition-colors">
                    <td className="py-2.5 px-3 font-semibold text-slate-800 whitespace-nowrap">
                      {item.displayDate}
                      {item.label && <div className="text-[10px] text-slate-400 font-normal">{item.label}</div>}
                    </td>
                    <td className="py-2.5 px-2 text-center font-mono font-medium text-slate-700">{item.h2}</td>
                    <td className="py-2.5 px-2 text-center font-mono font-medium text-slate-700">{item.ch4}</td>
                    <td className="py-2.5 px-2 text-center font-mono font-medium text-slate-700">{item.c2h6}</td>
                    <td className="py-2.5 px-2 text-center font-mono font-medium text-slate-700">{item.c2h4}</td>
                    <td className={`py-2.5 px-2 text-center font-mono font-bold ${isC2H2Alarm ? 'text-rose-600 bg-rose-50' : 'text-slate-800'}`}>
                      {item.c2h2}
                    </td>
                    <td className="py-2.5 px-2 text-center font-mono font-medium text-slate-700">{item.co}</td>
                    <td className="py-2.5 px-2 text-center font-mono font-medium text-slate-500">{item.co2}</td>
                    <td className="py-2.5 px-2 text-center font-mono font-bold text-indigo-700">{item.tdcg}</td>
                    <td className="py-2.5 px-3 text-slate-500 text-[11px] max-w-xs truncate" title={item.notes}>
                      {item.notes || '—'}
                    </td>
                    <td className="py-2.5 px-3 text-right whitespace-nowrap space-x-1.5">
                      {onLoadSampleToForm && (
                        <button
                          type="button"
                          onClick={() => onLoadSampleToForm(item)}
                          className="px-2 py-1 text-[11px] font-semibold bg-blue-50 text-blue-600 hover:bg-blue-100 rounded border border-blue-200 transition-colors"
                          title="Điền các thông số khí này vào biểu mẫu bên trái"
                        >
                          Nạp vào Form
                        </button>
                      )}
                      <button
                        type="button"
                        onClick={() => handleDeleteSample(item.id)}
                        className="p-1 text-slate-400 hover:text-rose-600 rounded transition-colors"
                        title="Xóa điểm đo này"
                      >
                        <Trash2 size={13} />
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {/* Add Custom Sample Modal */}
      {showAddModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-4">
          <div className="bg-white rounded-xl shadow-2xl max-w-lg w-full overflow-hidden border border-slate-200">
            <div className="px-6 py-4 bg-slate-900 text-white flex items-center justify-between">
              <h3 className="font-bold text-sm flex items-center gap-2">
                <Plus size={16} className="text-blue-400" />
                Thêm Điểm Đo Lịch Sử DGA Mới
              </h3>
              <button
                type="button"
                onClick={() => setShowAddModal(false)}
                className="text-slate-400 hover:text-white text-lg font-bold"
              >
                ✕
              </button>
            </div>

            <form onSubmit={handleSaveCustomSample} className="p-6 space-y-4 text-xs">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Ngày lấy mẫu *</label>
                  <input
                    type="date"
                    required
                    value={newSample.date}
                    onChange={e => setNewSample({ ...newSample, date: e.target.value })}
                    className="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Nhãn / Tiêu đề kỳ đo</label>
                  <input
                    type="text"
                    placeholder="VD: Kỳ 5 (2027)"
                    value={newSample.label || ''}
                    onChange={e => setNewSample({ ...newSample, label: e.target.value })}
                    className="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                </div>
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">Nồng độ các loại khí (ppm) *</label>
                <div className="grid grid-cols-4 gap-2 bg-slate-50 p-3 rounded-lg border border-slate-200">
                  <div>
                    <span className="block text-[10px] font-bold text-blue-700 uppercase">H₂</span>
                    <input
                      type="number"
                      step="any"
                      value={newSample.h2 || 0}
                      onChange={e => setNewSample({ ...newSample, h2: parseFloat(e.target.value) || 0 })}
                      className="w-full px-2 py-1 border border-slate-300 rounded text-center font-mono"
                    />
                  </div>
                  <div>
                    <span className="block text-[10px] font-bold text-emerald-700 uppercase">CH₄</span>
                    <input
                      type="number"
                      step="any"
                      value={newSample.ch4 || 0}
                      onChange={e => setNewSample({ ...newSample, ch4: parseFloat(e.target.value) || 0 })}
                      className="w-full px-2 py-1 border border-slate-300 rounded text-center font-mono"
                    />
                  </div>
                  <div>
                    <span className="block text-[10px] font-bold text-amber-700 uppercase">C₂H₆</span>
                    <input
                      type="number"
                      step="any"
                      value={newSample.c2h6 || 0}
                      onChange={e => setNewSample({ ...newSample, c2h6: parseFloat(e.target.value) || 0 })}
                      className="w-full px-2 py-1 border border-slate-300 rounded text-center font-mono"
                    />
                  </div>
                  <div>
                    <span className="block text-[10px] font-bold text-orange-700 uppercase">C₂H₄</span>
                    <input
                      type="number"
                      step="any"
                      value={newSample.c2h4 || 0}
                      onChange={e => setNewSample({ ...newSample, c2h4: parseFloat(e.target.value) || 0 })}
                      className="w-full px-2 py-1 border border-slate-300 rounded text-center font-mono"
                    />
                  </div>

                  <div>
                    <span className="block text-[10px] font-bold text-rose-700 uppercase">C₂H₂</span>
                    <input
                      type="number"
                      step="any"
                      value={newSample.c2h2 || 0}
                      onChange={e => setNewSample({ ...newSample, c2h2: parseFloat(e.target.value) || 0 })}
                      className="w-full px-2 py-1 border border-rose-300 bg-rose-50/50 rounded text-center font-mono font-bold"
                    />
                  </div>
                  <div>
                    <span className="block text-[10px] font-bold text-purple-700 uppercase">CO</span>
                    <input
                      type="number"
                      step="any"
                      value={newSample.co || 0}
                      onChange={e => setNewSample({ ...newSample, co: parseFloat(e.target.value) || 0 })}
                      className="w-full px-2 py-1 border border-slate-300 rounded text-center font-mono"
                    />
                  </div>
                  <div>
                    <span className="block text-[10px] font-bold text-slate-700 uppercase">CO₂</span>
                    <input
                      type="number"
                      step="any"
                      value={newSample.co2 || 0}
                      onChange={e => setNewSample({ ...newSample, co2: parseFloat(e.target.value) || 0 })}
                      className="w-full px-2 py-1 border border-slate-300 rounded text-center font-mono"
                    />
                  </div>
                  <div className="flex flex-col justify-end">
                    <span className="block text-[10px] font-bold text-indigo-700 uppercase">TDCG</span>
                    <div className="px-2 py-1 bg-indigo-50 border border-indigo-200 rounded text-center font-mono font-bold text-indigo-900">
                      {((newSample.h2 || 0) + (newSample.ch4 || 0) + (newSample.c2h6 || 0) + (newSample.c2h4 || 0) + (newSample.c2h2 || 0) + (newSample.co || 0)).toFixed(1)}
                    </div>
                  </div>
                </div>
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">Ghi chú hiện trường</label>
                <textarea
                  rows={2}
                  placeholder="Ghi chú về phụ tải, thời tiết, vị trí lấy mẫu van đáy..."
                  value={newSample.notes || ''}
                  onChange={e => setNewSample({ ...newSample, notes: e.target.value })}
                  className="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
                />
              </div>

              <div className="flex justify-end gap-2 pt-3 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setShowAddModal(false)}
                  className="px-4 py-2 border border-slate-300 rounded-lg text-slate-600 hover:bg-slate-50 font-medium"
                >
                  Hủy
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 font-semibold shadow-sm"
                >
                  Lưu Vào Lịch Sử Đo
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};
