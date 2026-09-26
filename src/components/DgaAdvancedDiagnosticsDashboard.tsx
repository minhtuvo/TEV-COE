import React, { useState, useMemo } from 'react';
import {
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  ReferenceLine
} from 'recharts';
import {
  Activity,
  AlertTriangle,
  CheckCircle2,
  Clock,
  Download,
  Upload,
  FileText,
  Sliders,
  TrendingUp,
  Flame,
  Zap,
  Info,
  Layers,
  Save,
  Calendar,
  ChevronRight,
  Sparkles,
  RefreshCw,
  X,
  User,
  Edit3,
  Check,
  ArrowUp,
  ArrowDown,
  ArrowRight,
  History,
  Thermometer,
  ChevronDown,
  ChevronUp
} from 'lucide-react';
import { regionsT1, regionsP1 } from '../App';
import { executeDgaDiagnostics } from '../utils/dgaDiagnosticEngine';
import { DgaParamTrendInput, computeDgaParamTrend } from './dga/DgaParamTrendInput';
import { DEFAULT_BASELINE_DGA_RECORD } from '../utils/dgaHistoryService';
import {
  DgaDetailsModal,
  RatioSettingsModal,
  RiskAssessmentModal,
  RatioAnalysisModal,
  SampleHistoryModal,
  TransformerOverviewReportModal,
  EventLogModal
} from './dga/DgaModals';

export interface DgaDataInput {
  h2: string | number;
  ch4: string | number;
  c2h6: string | number;
  c2h4: string | number;
  c2h2: string | number;
  co: string | number;
  co2: string | number;
  o2?: string | number;
  n2?: string | number;
  moisture?: string | number;
  acidity?: string | number;
  bdStrength?: string | number;
  age?: string | number;
  loadFactor?: string | number;
  estDp?: string | number;
}

interface DgaAdvancedDiagnosticsDashboardProps {
  currentFormValues: DgaDataInput;
  equipmentCode?: string;
  equipmentName?: string;
  customerName?: string;
  sampleDate?: string;
  engineerName?: string;
  analysisResult?: any;
  onAnalyze?: () => void;
  onUpdateFormValues?: (values: Partial<DgaDataInput>) => void;
  onExportPdf?: () => void;
  onUpdateAssetInfo?: (info: { equipmentCode?: string; equipmentName?: string; sampleDate?: string; engineerName?: string }) => void;
  lastRecordedValues?: any;
}

// 7-day or historical timeline samples matching the user's dashboard screenshot
const DEFAULT_TIMELINE = [
  { date: 'May 08', fullDate: '2025-05-08 08:15', h2: 2850, ch4: 340, c2h6: 42, c2h4: 28, c2h2: 5.2, co: 58, co2: 2400 },
  { date: 'May 09', fullDate: '2025-05-09 08:15', h2: 2900, ch4: 355, c2h6: 44, c2h4: 29, c2h2: 5.4, co: 59, co2: 2450 },
  { date: 'May 10', fullDate: '2025-05-10 08:15', h2: 2800, ch4: 330, c2h6: 40, c2h4: 27, c2h2: 5.0, co: 56, co2: 2380 },
  { date: 'May 11', fullDate: '2025-05-11 08:15', h2: 3050, ch4: 380, c2h6: 46, c2h4: 32, c2h2: 6.1, co: 60, co2: 2520 },
  { date: 'May 12', fullDate: '2025-05-12 08:15', h2: 3120, ch4: 395, c2h6: 47, c2h4: 34, c2h2: 6.8, co: 61, co2: 2590 },
  { date: 'May 13', fullDate: '2025-05-13 08:15', h2: 3200, ch4: 405, c2h6: 48, c2h4: 35, c2h2: 7.2, co: 62, co2: 2650 },
  { date: 'May 14', fullDate: '2025-05-14 08:15', h2: 3280, ch4: 412, c2h6: 48, c2h4: 36, c2h2: 7.6, co: 62, co2: 2700 }
];

export const DgaAdvancedDiagnosticsDashboard: React.FC<DgaAdvancedDiagnosticsDashboardProps> = ({
  currentFormValues,
  equipmentCode = 'TR-01',
  equipmentName = 'Máy biến áp T1 110kV',
  customerName = 'EVN / Substation',
  sampleDate = '2025-05-14T08:15',
  engineerName = 'Nguyễn Văn Tuấn (Kỹ sư DGA)',
  analysisResult,
  onAnalyze,
  onUpdateFormValues,
  onExportPdf,
  onUpdateAssetInfo,
  lastRecordedValues
}) => {
  // Historical record fallback
  const histRecord = lastRecordedValues || DEFAULT_BASELINE_DGA_RECORD;

  // Asset info editable state
  const [assetCode, setAssetCode] = useState<string>(equipmentCode);
  const [assetName, setAssetName] = useState<string>(equipmentName);
  const [entryDate, setEntryDate] = useState<string>(sampleDate);
  const [entryEngineer, setEntryEngineer] = useState<string>(engineerName);
  const [isEditingAsset, setIsEditingAsset] = useState<boolean>(false);
  const [savedSuccessMsg, setSavedSuccessMsg] = useState<string>('');
  const [showTrendInputs, setShowTrendInputs] = useState<boolean>(false);

  // Filters state
  const [selectedAsset, setSelectedAsset] = useState<string>(`${equipmentCode} — ${equipmentName}`);
  const [timeRange, setTimeRange] = useState<string>('Last 7 days');
  const [dgaSystem, setDgaSystem] = useState<string>('DGA-1 — Online DGA System');
  const [activeMethodTab, setActiveMethodTab] = useState<string>('triangle');
  const [isGasTrendLogScale, setIsGasTrendLogScale] = useState<boolean>(true);
  const [isRatioTrendLogScale, setIsRatioTrendLogScale] = useState<boolean>(true);

  // Modals state
  const [showDgaDetailsModal, setShowDgaDetailsModal] = useState<boolean>(false);
  const [showRatioSettingsModal, setShowRatioSettingsModal] = useState<boolean>(false);
  const [showRiskAssessmentModal, setShowRiskAssessmentModal] = useState<boolean>(false);
  const [showRatioAnalysisModal, setShowRatioAnalysisModal] = useState<boolean>(false);
  const [showSampleHistoryModal, setShowSampleHistoryModal] = useState<boolean>(false);
  const [showReportModal, setShowReportModal] = useState<boolean>(false);
  const [showEventLogModal, setShowEventLogModal] = useState<boolean>(false);
  const [showImportExportModal, setShowImportExportModal] = useState<boolean>(false);
  const [importExportTab, setImportExportTab] = useState<'import' | 'export'>('import');
  const [showAnnotationsModal, setShowAnnotationsModal] = useState<boolean>(false);
  const [annotations, setAnnotations] = useState([
    { id: 1, date: 'May 13, 2025 18:05', title: 'CO level elevated — below alarm threshold' },
    { id: 2, date: 'May 11, 2025 10:20', title: 'Load transfer event' },
    { id: 3, date: 'May 09, 2025 14:15', title: 'DGA system maintenance' }
  ]);
  const [newAnnotationText, setNewAnnotationText] = useState('');

  const handleSaveAssetInfo = () => {
    onUpdateAssetInfo?.({
      equipmentCode: assetCode,
      equipmentName: assetName,
      sampleDate: entryDate,
      engineerName: entryEngineer
    });
    setSelectedAsset(`${assetCode} — ${assetName}`);
    setIsEditingAsset(false);
    setSavedSuccessMsg('Đã lưu thông tin thiết bị & kỹ sư!');
    setTimeout(() => setSavedSuccessMsg(''), 3000);
  };

  const handleOpenDgaDetails = () => {
    if (onAnalyze) onAnalyze();
    setShowDgaDetailsModal(true);
  };

  const handleLoadSampleFromHistory = (sample: any) => {
    if (onUpdateFormValues) {
      onUpdateFormValues({
        h2: sample.h2,
        ch4: sample.ch4,
        c2h6: sample.c2h6,
        c2h4: sample.c2h4,
        c2h2: sample.c2h2,
        co: sample.co,
        co2: sample.co2
      });
    }
    if (sample.date) setEntryDate(sample.date.replace(' ', 'T'));
    if (sample.engineer) setEntryEngineer(sample.engineer);
    if (onAnalyze) onAnalyze();
    setSavedSuccessMsg(`Đã nạp mẫu ${sample.sampleId} vào phân tích!`);
    setTimeout(() => setSavedSuccessMsg(''), 3000);
  };

  // Gas values parse
  const h2 = parseFloat(String(currentFormValues.h2 || 0)) || 0;
  const ch4 = parseFloat(String(currentFormValues.ch4 || 0)) || 0;
  const c2h6 = parseFloat(String(currentFormValues.c2h6 || 0)) || 0;
  const c2h4 = parseFloat(String(currentFormValues.c2h4 || 0)) || 0;
  const c2h2 = parseFloat(String(currentFormValues.c2h2 || 0)) || 0;
  const co = parseFloat(String(currentFormValues.co || 0)) || 0;
  const co2 = parseFloat(String(currentFormValues.co2 || 0)) || 0;
  const tdcg = h2 + ch4 + c2h6 + c2h4 + c2h2 + co;

  // Real-time engine analysis
  const engineAnalysis = useMemo(() => {
    return executeDgaDiagnostics({
      h2,
      ch4,
      c2h6,
      c2h4,
      c2h2,
      co,
      co2,
      moisture: parseFloat(String(currentFormValues.moisture || 12)),
      bdStrength: parseFloat(String(currentFormValues.bdStrength || 58)),
      acidity: parseFloat(String(currentFormValues.acidity || 0.04)),
      ffa: parseFloat(String(currentFormValues.ffa || 0)),
      estDp: parseFloat(String(currentFormValues.estDp || 650)),
      age: parseFloat(String(currentFormValues.age || 20)),
      loadFactor: parseFloat(String(currentFormValues.loadFactor || 65))
    });
  }, [h2, ch4, c2h6, c2h4, c2h2, co, co2, currentFormValues]);

  const activeAnalysis = analysisResult || engineAnalysis;
  const isHealthyCondition1 = activeAnalysis.faultDiagnosisSkipped ?? (
    tdcg <= 720 && h2 < 100 && ch4 < 120 && c2h6 < 65 && c2h4 < 50 && c2h2 < 35 && co < 350
  );

  // Gas Ratios
  const rCH4_H2 = h2 > 0 ? (ch4 / h2) : 0;
  const rC2H2_C2H4 = c2h4 > 0 ? (c2h2 / c2h4) : 0;
  const rC2H4_C2H6 = c2h6 > 0 ? (c2h4 / c2h6) : 0;
  const rC2H2_CH4 = ch4 > 0 ? (c2h2 / ch4) : 0;
  const rC2H6_C2H2 = c2h2 > 0 ? (c2h6 / c2h2) : 0;
  const rCO2_CO = co > 0 ? (co2 / co) : 0;

  // Active status checks according to IEEE C57.104 Condition 1 rule
  const isAlarmC2H2 = !isHealthyCondition1 && c2h2 >= 1.0;
  const isAlertC2H2 = !isHealthyCondition1 && c2h2 >= 0.5 && c2h2 < 1.0;
  const isTdcgHigh = !isHealthyCondition1 && tdcg > 720;
  const hasAlarm = !isHealthyCondition1 && (isAlarmC2H2 || tdcg > 1920);
  const hasAlert = !isHealthyCondition1 && (isAlertC2H2 || (tdcg > 720 && tdcg <= 1920));

  // Overall Risk Level calculation (0-100)
  const riskScore = useMemo(() => {
    if (isHealthyCondition1) return 18;
    let score = 25;
    if (tdcg > 720) score += 20;
    if (tdcg > 1920) score += 25;
    if (c2h2 > 1.0) score += 30;
    else if (c2h2 > 0.1) score += 15;
    if (activeAnalysis?.currentHealthScore) {
      score = Math.min(100, Math.round(score * 0.4 + (parseFloat(activeAnalysis.currentHealthScore) * 10) * 0.6));
    }
    return Math.min(95, Math.max(15, score));
  }, [isHealthyCondition1, tdcg, c2h2, activeAnalysis]);

  // Combine timeline with current active sample for charts
  const chartTimeline = useMemo(() => {
    const list = [...DEFAULT_TIMELINE];
    // Update last entry with current form values if provided
    if (tdcg > 0) {
      list[list.length - 1] = {
        date: 'May 14 (Now)',
        fullDate: '2025-05-14 08:15',
        h2: Math.max(0.1, h2),
        ch4: Math.max(0.1, ch4),
        c2h6: Math.max(0.1, c2h6),
        c2h4: Math.max(0.1, c2h4),
        c2h2: Math.max(0.01, c2h2),
        co: Math.max(0.1, co),
        co2: Math.max(0.1, co2)
      };
    }
    return list.map(item => ({
      ...item,
      r1: item.h2 > 0 ? parseFloat((item.ch4 / item.h2).toFixed(3)) : 0.01,
      r2: item.c2h4 > 0 ? parseFloat((item.c2h2 / item.c2h4).toFixed(3)) : 0.01,
      r3: item.c2h6 > 0 ? parseFloat((item.c2h4 / item.c2h6).toFixed(3)) : 0.01
    }));
  }, [h2, ch4, c2h6, c2h4, c2h2, co, co2, tdcg]);

  // Active diagnostic methods data (Condition 1 skips faults)
  const rogersInterpretation = isHealthyCondition1 ? 'Normal (Cond 1)' : (activeAnalysis?.matrix?.['IEEE (Rogers)'] || (rCH4_H2 < 0.1 && rC2H2_C2H4 < 0.1 ? 'PD' : 'Normal'));
  const duvalT1Interpretation = isHealthyCondition1 ? 'Normal (Cond 1)' : (activeAnalysis?.matrix?.['Duval T1'] || 'Normal');
  const duvalP1Interpretation = isHealthyCondition1 ? 'Normal (Cond 1)' : (activeAnalysis?.matrix?.['Duval P1'] || 'Normal');
  const dornenburgInterpretation = isHealthyCondition1 ? 'Normal (Cond 1)' : (activeAnalysis?.matrix?.['IEEE (Dornenburg)'] || 'Normal');

  // Ternary Duval Triangle coordinates calculation
  const triangleCoords = useMemo(() => {
    const sumT1 = ch4 + c2h4 + c2h2;
    if (sumT1 <= 0) return { x: 50, y: 50, pCH4: 0, pC2H4: 0, pC2H2: 0 };
    const pCH4 = (ch4 / sumT1) * 100;
    const pC2H4 = (c2h4 / sumT1) * 100;
    const pC2H2 = (c2h2 / sumT1) * 100;

    // Standard ternary to 2D Cartesian mapping inside viewBox 0 0 100 100
    const x = 6.7 * (pC2H2 / 100) + 93.3 * (pC2H4 / 100) + 50 * (pCH4 / 100);
    const y = 80 * (pC2H2 / 100) + 80 * (pC2H4 / 100) + 5 * (pCH4 / 100);
    return { x, y, pCH4: pCH4.toFixed(1), pC2H4: pC2H4.toFixed(1), pC2H2: pC2H2.toFixed(1) };
  }, [ch4, c2h4, c2h2]);

  // Pentagon coordinates calculation
  const pentagonCoords = useMemo(() => {
    const sumP = h2 + c2h6 + ch4 + c2h4 + c2h2;
    if (sumP <= 0) return { x: 50, y: 50 };
    const p1 = h2 / sumP;
    const p2 = c2h6 / sumP;
    const p3 = ch4 / sumP;
    const p4 = c2h4 / sumP;
    const p5 = c2h2 / sumP;

    const mathX = p1 * 0 + p2 * (-38) + p3 * (-23.5) + p4 * 23.5 + p5 * 38;
    const mathY = p1 * 40 + p2 * 12.4 + p3 * (-32.4) + p4 * (-32.4) + p5 * 12.4;
    return { x: 50 - mathX, y: 50 - mathY };
  }, [h2, c2h6, ch4, c2h4, c2h2]);

  const handleAddAnnotation = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newAnnotationText.trim()) return;
    const now = new Date().toLocaleDateString('en-US', { month: 'short', day: '2-digit', year: 'numeric' }) + ' ' +
                new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
    setAnnotations(prev => [{ id: Date.now(), date: now, title: newAnnotationText.trim() }, ...prev]);
    setNewAnnotationText('');
  };

  return (
    <div className="space-y-6 font-sans text-slate-800">
      {/* 1. TOP HEADER & CONTEXT BAR */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-200/80 p-5 transition-all">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="w-11 h-11 rounded-xl bg-blue-50 text-blue-600 flex items-center justify-center border border-blue-100 shadow-xs">
              <Activity size={24} className="stroke-[2.2]" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-xl font-bold text-slate-900 tracking-tight">
                  DGA Monitoring &amp; Diagnostics
                </h1>
                <div
                  className="group relative cursor-pointer text-slate-400 hover:text-slate-600"
                  title="Chẩn đoán nâng cao khí hòa tan dầu biến áp theo IEC 60599 & IEEE C57.104"
                >
                  <Info size={16} />
                </div>
              </div>
              <p className="text-xs text-slate-500 font-medium">
                Trending, gas ratios, alarms and diagnostic interpretation for dissolved gas analysis.
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-3">
            {/* Transformer Asset Code Input */}
            <div className="flex flex-col">
              <label className="text-[11px] font-semibold text-slate-500 mb-0.5">Mã thiết bị (Asset Code)</label>
              <input
                type="text"
                value={assetCode}
                onChange={e => {
                  const val = e.target.value;
                  setAssetCode(val);
                  if (onUpdateAssetInfo) {
                    onUpdateAssetInfo({ assetCode: val, assetName, entryDate, entryEngineer });
                  }
                }}
                placeholder="TR-01"
                className="w-24 text-xs font-bold bg-white border border-slate-300 focus:border-blue-500 rounded-lg px-2.5 py-1.5 outline-none focus:ring-2 focus:ring-blue-100 text-slate-800 shadow-2xs transition-all"
              />
            </div>

            {/* Transformer Asset Name Input */}
            <div className="flex flex-col">
              <label className="text-[11px] font-semibold text-slate-500 mb-0.5">Tên máy biến áp (Asset Name)</label>
              <input
                type="text"
                value={assetName}
                onChange={e => {
                  const val = e.target.value;
                  setAssetName(val);
                  if (onUpdateAssetInfo) {
                    onUpdateAssetInfo({ assetCode, assetName: val, entryDate, entryEngineer });
                  }
                }}
                placeholder="MBA T1 110kV Bắc Ninh"
                className="w-56 text-xs font-semibold bg-white border border-slate-300 focus:border-blue-500 rounded-lg px-2.5 py-1.5 outline-none focus:ring-2 focus:ring-blue-100 text-slate-800 shadow-2xs transition-all"
              />
            </div>

            {/* Quick Preset Selector */}
            <div className="flex flex-col">
              <label className="text-[11px] font-semibold text-slate-500 mb-0.5">Chọn mẫu nhanh</label>
              <select
                value={selectedAsset}
                onChange={e => {
                  const val = e.target.value;
                  setSelectedAsset(val);
                  const parts = val.split(' — ');
                  const newCode = parts[0] || assetCode;
                  const newName = parts[1] || assetName;
                  setAssetCode(newCode);
                  setAssetName(newName);
                  if (onUpdateAssetInfo) {
                    onUpdateAssetInfo({ assetCode: newCode, assetName: newName, entryDate, entryEngineer });
                  }
                }}
                className="text-xs font-medium bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-lg px-2.5 py-1.5 outline-none focus:ring-2 focus:ring-blue-500 cursor-pointer text-slate-700"
              >
                <option value={`${assetCode} — ${assetName}`}>{assetCode} — {assetName}</option>
                <option value="TR-01 — MBA T1 110kV Bắc Ninh">TR-01 — MBA T1 110kV Bắc Ninh (Chuẩn)</option>
                <option value="TR-101 — MBA T2 220kV Đông Anh">TR-101 — MBA T2 220kV Đông Anh (Cảnh báo nhiệt)</option>
                <option value="TR-102 — MBA T3 110/22kV 40MVA">TR-102 — MBA T3 110/22kV 40MVA (Hồ quang nhẹ)</option>
              </select>
            </div>

            {/* Ngày nhập liệu / Thời gian lấy mẫu */}
            <div className="flex flex-col">
              <label className="text-[11px] font-semibold text-slate-500 mb-0.5 flex items-center gap-1">
                <Calendar size={11} className="text-slate-400" />
                Ngày nhập liệu / lấy mẫu
              </label>
              <input
                type="text"
                value={entryDate}
                onChange={e => {
                  const val = e.target.value;
                  setEntryDate(val);
                  if (onUpdateAssetInfo) {
                    onUpdateAssetInfo({ assetCode, assetName, entryDate: val, entryEngineer });
                  }
                }}
                placeholder="14/05/2025 08:15"
                className="text-xs font-semibold bg-white border border-slate-300 focus:border-blue-500 rounded-lg px-2.5 py-1.5 outline-none focus:ring-2 focus:ring-blue-100 text-slate-800 w-36 shadow-2xs transition-all"
              />
            </div>

            {/* Tên kỹ sư nhập liệu */}
            <div className="flex flex-col">
              <label className="text-[11px] font-semibold text-slate-500 mb-0.5 flex items-center gap-1">
                <User size={11} className="text-slate-400" />
                Kỹ sư nhập liệu
              </label>
              <input
                type="text"
                value={entryEngineer}
                onChange={e => {
                  const val = e.target.value;
                  setEntryEngineer(val);
                  if (onUpdateAssetInfo) {
                    onUpdateAssetInfo({ assetCode, assetName, entryDate, entryEngineer: val });
                  }
                }}
                placeholder="Nguyễn Văn Tuấn"
                className="text-xs font-semibold bg-white border border-slate-300 focus:border-blue-500 rounded-lg px-2.5 py-1.5 outline-none focus:ring-2 focus:ring-blue-100 text-slate-800 w-36 shadow-2xs transition-all"
              />
            </div>

            {/* Time Range Select */}
            <div className="flex flex-col">
              <label className="text-[11px] font-semibold text-slate-400 mb-0.5">Time range</label>
              <div className="relative">
                <select
                  value={timeRange}
                  onChange={e => setTimeRange(e.target.value)}
                  className="text-xs font-semibold bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-lg pl-3 pr-8 py-1.5 outline-none focus:ring-2 focus:ring-blue-500 cursor-pointer text-slate-700"
                >
                  <option value="Last 7 days">Last 7 days</option>
                  <option value="Last 30 days">Last 30 days</option>
                  <option value="Last 6 months">Last 6 months</option>
                  <option value="All historical data">All historical data</option>
                </select>
                <Calendar size={13} className="absolute right-2.5 top-2.5 text-slate-400 pointer-events-none" />
              </div>
            </div>

            {/* Save / Apply Button */}
            <div className="flex items-end pt-3 sm:pt-0">
              <button
                type="button"
                onClick={handleSaveAssetInfo}
                className="inline-flex items-center gap-1.5 bg-blue-600 hover:bg-blue-700 active:bg-blue-800 text-white text-xs font-bold px-3.5 py-2 rounded-lg shadow-sm transition-all cursor-pointer"
                title="Lưu thông tin thiết bị và người nhập liệu"
              >
                <Save size={14} />
                <span>Lưu thông tin</span>
              </button>
            </div>
          </div>
        </div>

        {savedSuccessMsg && (
          <div className="mt-3 p-2 bg-emerald-50 border border-emerald-200 rounded-lg text-emerald-800 text-xs font-semibold flex items-center gap-2 animate-in fade-in">
            <Check size={14} className="text-emerald-600" />
            <span>{savedSuccessMsg}</span>
          </div>
        )}
      </div>

      {/* 1.5 VISUAL TREND PARAMETERS & HISTORICAL COMPARISON BAR */}
      <div className="bg-white rounded-2xl border border-slate-200 shadow-2xs overflow-hidden">
        <div className="p-4 bg-gradient-to-r from-slate-50 via-blue-50/40 to-slate-50 border-b border-slate-200/80 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-xl bg-blue-600 text-white flex items-center justify-center shadow-xs flex-shrink-0">
              <Thermometer size={18} />
            </div>
            <div>
              <div className="flex items-center gap-2 flex-wrap">
                <span className="text-sm font-extrabold text-slate-900">
                  Tham số đo &amp; Chỉ báo xu hướng so với lịch sử (Visual Trend Indicator)
                </span>
                <span className="text-[10px] font-mono font-bold text-blue-700 bg-blue-100/70 px-2 py-0.5 rounded-full border border-blue-200">
                  Cơ sở dữ liệu: {histRecord.date}
                </span>
              </div>
              <p className="text-[11px] text-slate-500 mt-0.5">
                Mũi tên trực quan (▲ Đỏ: Tăng nguy hiểm / ▼ Xanh: Giảm cải thiện) tự động đối soát với mẫu đo kỳ trước ({histRecord.label || histRecord.source})
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              type="button"
              onClick={() => setShowTrendInputs(!showTrendInputs)}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all shadow-2xs cursor-pointer ${
                showTrendInputs
                  ? 'bg-blue-600 text-white shadow-xs'
                  : 'bg-white hover:bg-slate-100 text-slate-700 border border-slate-300'
              }`}
            >
              <Sliders size={13} />
              <span>{showTrendInputs ? 'Thu gọn tham số' : 'Mở bảng tham số (▲/▼)'}</span>
              {showTrendInputs ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
            </button>
          </div>
        </div>

        {/* Collapsed Compact Trend Summary Strip */}
        {!showTrendInputs && (
          <div className="px-4 py-2.5 bg-white text-xs flex items-center justify-between gap-3 overflow-x-auto">
            <div className="flex items-center gap-3 text-[11px] text-slate-600 flex-shrink-0">
              <span className="font-bold text-slate-800 flex items-center gap-1">
                <History size={12} className="text-blue-500" />
                Đối soát kỳ trước:
              </span>
              {[
                { key: 'h2', label: 'H₂', val: h2, hist: histRecord.gases.h2 },
                { key: 'ch4', label: 'CH₄', val: ch4, hist: histRecord.gases.ch4 },
                { key: 'c2h4', label: 'C₂H₄', val: c2h4, hist: histRecord.gases.c2h4 },
                { key: 'c2h2', label: 'C₂H₂', val: c2h2, hist: histRecord.gases.c2h2 },
                { key: 'co', label: 'CO', val: co, hist: histRecord.gases.co }
              ].map(item => {
                const tr = computeDgaParamTrend(item.val, item.hist, false);
                return (
                  <span key={item.key} className="inline-flex items-center gap-1 bg-slate-50 px-2 py-0.5 rounded border border-slate-200 font-mono">
                    <span className="font-bold text-slate-700">{item.label}:</span>
                    <span>{item.val}</span>
                    {tr.hasHistorical && (
                      <span className={`inline-flex items-center gap-0.5 text-[10px] font-bold ${
                        tr.severity === 'worse' ? 'text-rose-600' : tr.severity === 'better' ? 'text-emerald-600' : 'text-slate-400'
                      }`}>
                        {tr.direction === 'up' && <ArrowUp size={10} />}
                        {tr.direction === 'down' && <ArrowDown size={10} />}
                        {tr.direction === 'flat' && <ArrowRight size={10} />}
                        {tr.badgeText}
                      </span>
                    )}
                  </span>
                );
              })}
            </div>

            <button
              type="button"
              onClick={() => setShowTrendInputs(true)}
              className="text-[11px] font-bold text-blue-600 hover:text-blue-800 flex-shrink-0 flex items-center gap-0.5"
            >
              <span>Xem chi tiết 9 khí</span>
              <ChevronRight size={12} />
            </button>
          </div>
        )}

        {/* Expanded Full Parameter Grid with Visual Trend Indicators */}
        {showTrendInputs && (
          <div className="p-5 bg-white border-t border-slate-100 space-y-4 animate-in fade-in duration-150">
            <div>
              <div className="text-xs font-bold text-slate-800 uppercase tracking-wider mb-2.5 flex items-center justify-between">
                <span>Nồng độ 9 loại khí hòa tan (ppm) &amp; Mũi tên xu hướng</span>
                <span className="text-[10px] text-slate-500 font-normal">Chuẩn IEEE C57.104 Condition 1</span>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-5 gap-3">
                <DgaParamTrendInput
                  paramKey="h2"
                  label="H₂"
                  subLabel="Hydrogen"
                  unit="ppm"
                  standard="≤ 100"
                  value={currentFormValues.h2}
                  onChange={(val) => onUpdateFormValues?.({ h2: val })}
                  historicalValue={histRecord.gases.h2}
                  historicalDate={histRecord.date}
                  isHigherBetter={false}
                  placeholder="Nhập H₂"
                />
                <DgaParamTrendInput
                  paramKey="ch4"
                  label="CH₄"
                  subLabel="Methane"
                  unit="ppm"
                  standard="≤ 120"
                  value={currentFormValues.ch4}
                  onChange={(val) => onUpdateFormValues?.({ ch4: val })}
                  historicalValue={histRecord.gases.ch4}
                  historicalDate={histRecord.date}
                  isHigherBetter={false}
                  placeholder="Nhập CH₄"
                />
                <DgaParamTrendInput
                  paramKey="c2h6"
                  label="C₂H₆"
                  subLabel="Ethane"
                  unit="ppm"
                  standard="≤ 65"
                  value={currentFormValues.c2h6}
                  onChange={(val) => onUpdateFormValues?.({ c2h6: val })}
                  historicalValue={histRecord.gases.c2h6}
                  historicalDate={histRecord.date}
                  isHigherBetter={false}
                  placeholder="Nhập C₂H₆"
                />
                <DgaParamTrendInput
                  paramKey="c2h4"
                  label="C₂H₄"
                  subLabel="Ethylene"
                  unit="ppm"
                  standard="≤ 50"
                  value={currentFormValues.c2h4}
                  onChange={(val) => onUpdateFormValues?.({ c2h4: val })}
                  historicalValue={histRecord.gases.c2h4}
                  historicalDate={histRecord.date}
                  isHigherBetter={false}
                  placeholder="Nhập C₂H₄"
                />
                <DgaParamTrendInput
                  paramKey="c2h2"
                  label="C₂H₂"
                  subLabel="Acetylene"
                  unit="ppm"
                  standard="≤ 1.0"
                  value={currentFormValues.c2h2}
                  onChange={(val) => onUpdateFormValues?.({ c2h2: val })}
                  historicalValue={histRecord.gases.c2h2}
                  historicalDate={histRecord.date}
                  isHigherBetter={false}
                  placeholder="Nhập C₂H₂"
                />
                <DgaParamTrendInput
                  paramKey="co"
                  label="CO"
                  subLabel="Carbon Monoxide"
                  unit="ppm"
                  standard="≤ 350"
                  value={currentFormValues.co}
                  onChange={(val) => onUpdateFormValues?.({ co: val })}
                  historicalValue={histRecord.gases.co}
                  historicalDate={histRecord.date}
                  isHigherBetter={false}
                  placeholder="Nhập CO"
                />
                <DgaParamTrendInput
                  paramKey="co2"
                  label="CO₂"
                  subLabel="Carbon Dioxide"
                  unit="ppm"
                  standard="≤ 2500"
                  value={currentFormValues.co2}
                  onChange={(val) => onUpdateFormValues?.({ co2: val })}
                  historicalValue={histRecord.gases.co2}
                  historicalDate={histRecord.date}
                  isHigherBetter={false}
                  placeholder="Nhập CO₂"
                />
                <DgaParamTrendInput
                  paramKey="o2"
                  label="O₂"
                  subLabel="Oxygen"
                  unit="ppm"
                  standard="≤ 3500"
                  value={currentFormValues.o2 || ''}
                  onChange={(val) => onUpdateFormValues?.({ o2: val })}
                  historicalValue={histRecord.gases.o2}
                  historicalDate={histRecord.date}
                  isHigherBetter={false}
                  placeholder="Nhập O₂"
                />
                <DgaParamTrendInput
                  paramKey="n2"
                  label="N₂"
                  subLabel="Nitrogen"
                  unit="ppm"
                  standard="≤ 50000"
                  value={currentFormValues.n2 || ''}
                  onChange={(val) => onUpdateFormValues?.({ n2: val })}
                  historicalValue={histRecord.gases.n2}
                  historicalDate={histRecord.date}
                  isHigherBetter={false}
                  placeholder="Nhập N₂"
                />
                <div className="flex flex-col justify-center gap-2 p-2.5 bg-slate-50 rounded-xl border border-slate-200">
                  <span className="text-[11px] font-bold text-slate-700">Tổng khí cháy (TDCG)</span>
                  <div className="text-xl font-black text-slate-900">{tdcg.toFixed(0)} <span className="text-xs font-normal">ppm</span></div>
                  <button
                    type="button"
                    onClick={() => {
                      if (onAnalyze) onAnalyze();
                      setSavedSuccessMsg('Đã cập nhật tính toán DGA với số liệu mới!');
                      setTimeout(() => setSavedSuccessMsg(''), 3000);
                    }}
                    className="w-full py-1.5 bg-blue-600 hover:bg-blue-700 active:bg-blue-800 text-white rounded-lg text-xs font-bold shadow-2xs transition-colors flex items-center justify-center gap-1 cursor-pointer"
                  >
                    <Activity size={13} />
                    <span>Chẩn đoán ngay</span>
                  </button>
                </div>
              </div>
            </div>

            {/* Other Parameters with Trend Indicators */}
            <div className="border-t border-slate-100 pt-3">
              <div className="text-xs font-bold text-slate-800 uppercase tracking-wider mb-2.5 flex items-center justify-between">
                <span>Thông số lý hóa dầu &amp; Tình trạng MBA (ASTM / IEC)</span>
                <span className="text-[10px] text-slate-500 font-normal">ASTM D1533 / D1816 / D974</span>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                <DgaParamTrendInput
                  paramKey="moisture"
                  label="Moisture"
                  subLabel="Ẩm trong dầu"
                  unit="ppm"
                  standard="≤ 20"
                  value={currentFormValues.moisture ?? 12}
                  onChange={(val) => onUpdateFormValues?.({ moisture: val })}
                  historicalValue={histRecord.otherParams.moisture}
                  historicalDate={histRecord.date}
                  isHigherBetter={false}
                  placeholder="Nhập Moisture"
                />
                <DgaParamTrendInput
                  paramKey="bdStrength"
                  label="BD Strength"
                  subLabel="Đánh thủng"
                  unit="kV"
                  standard="≥ 50"
                  value={currentFormValues.bdStrength ?? 58}
                  onChange={(val) => onUpdateFormValues?.({ bdStrength: val })}
                  historicalValue={histRecord.otherParams.bdStrength}
                  historicalDate={histRecord.date}
                  isHigherBetter={true}
                  placeholder="Nhập BD Strength"
                />
                <DgaParamTrendInput
                  paramKey="acidity"
                  label="Acidity"
                  subLabel="Chỉ số axit"
                  unit="mgKOH/g"
                  standard="≤ 0.10"
                  value={currentFormValues.acidity ?? 0.04}
                  onChange={(val) => onUpdateFormValues?.({ acidity: val })}
                  historicalValue={histRecord.otherParams.acidity}
                  historicalDate={histRecord.date}
                  isHigherBetter={false}
                  placeholder="Nhập Acidity"
                />
                <DgaParamTrendInput
                  paramKey="estDp"
                  label="Est DP"
                  subLabel="Độ trùng hợp"
                  unit="DP"
                  standard="≥ 500"
                  value={currentFormValues.estDp ?? 650}
                  onChange={(val) => onUpdateFormValues?.({ estDp: val })}
                  historicalValue={histRecord.otherParams.estDp}
                  historicalDate={histRecord.date}
                  isHigherBetter={true}
                  placeholder="Nhập DP"
                />
              </div>
            </div>
          </div>
        )}
      </div>

      {/* 2. TOP 4 STATUS & KPI CARDS */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Card 1: DGA Status */}
        <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-2xs hover:shadow-sm transition-all flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">DGA Status</span>
              <span
                className="text-slate-400 hover:text-slate-600 cursor-pointer"
                title="Xem chi tiết phân tích DGA"
                onClick={handleOpenDgaDetails}
              >
                <Info size={14} />
              </span>
            </div>
            <div className="mt-3 flex items-center gap-3">
              <div className={`w-10 h-10 rounded-full flex items-center justify-center flex-shrink-0 ${
                hasAlarm ? 'bg-rose-100 text-rose-600' : hasAlert ? 'bg-amber-100 text-amber-600' : 'bg-emerald-100 text-emerald-600'
              }`}>
                {hasAlarm ? <Flame size={22} /> : hasAlert ? <AlertTriangle size={20} /> : <CheckCircle2 size={22} />}
              </div>
              <div>
                <div className="text-lg font-extrabold text-slate-900 leading-tight">
                  {hasAlarm ? 'Alarm' : hasAlert ? 'Alert / Warning' : 'Normal'}
                </div>
                <div className="text-[11px] text-slate-500 mt-0.5">
                  {hasAlarm
                    ? `C₂H₂ = ${c2h2} ppm (Vượt ngưỡng hồ quang)`
                    : hasAlert
                    ? 'Xuất hiện khí cháy cần giám sát'
                    : 'No abnormal gas levels detected'}
                </div>
              </div>
            </div>
          </div>
          <div className="mt-4 pt-3 border-t border-slate-100">
            <button
              type="button"
              onClick={handleOpenDgaDetails}
              className="text-xs font-semibold text-blue-600 hover:text-blue-800 flex items-center gap-1 group"
            >
              <span>View DGA details</span>
              <ChevronRight size={13} className="group-hover:translate-x-0.5 transition-transform" />
            </button>
          </div>
        </div>

        {/* Card 2: Latest Sample Time */}
        <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-2xs hover:shadow-sm transition-all flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">Latest Sample Time</span>
              <span className="text-slate-400 hover:text-slate-600 cursor-pointer" onClick={() => setShowSampleHistoryModal(true)}>
                <Info size={14} />
              </span>
            </div>
            <div className="mt-3 flex items-center gap-3">
              <div className="w-10 h-10 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center flex-shrink-0">
                <Clock size={22} />
              </div>
              <div>
                <div className="text-sm font-extrabold text-slate-900 leading-tight">
                  {entryDate}
                </div>
                <div className="text-[11px] text-slate-500 mt-0.5 line-clamp-1">
                  Kỹ sư: <span className="font-semibold text-slate-700">{entryEngineer}</span>
                </div>
              </div>
            </div>
          </div>
          <div className="mt-4 pt-3 border-t border-slate-100">
            <button
              type="button"
              onClick={() => setShowSampleHistoryModal(true)}
              className="text-xs font-semibold text-blue-600 hover:text-blue-800 flex items-center gap-1 cursor-pointer"
            >
              <span>View sample history</span>
              <ChevronRight size={13} />
            </button>
          </div>
        </div>

        {/* Card 3: Gas Ratio Alarm Status */}
        <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-2xs hover:shadow-sm transition-all flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">Gas Ratio Alarm Status</span>
              <span className="text-slate-400 hover:text-slate-600 cursor-pointer" onClick={() => setShowRatioSettingsModal(true)}>
                <Info size={14} />
              </span>
            </div>
            <div className="mt-3 grid grid-cols-4 gap-1 text-center">
              <div className="flex flex-col items-center">
                <span className="text-[10px] text-slate-500 font-bold mb-1">CH₄/H₂</span>
                <span className="w-6 h-6 rounded-full bg-emerald-100 text-emerald-600 flex items-center justify-center text-xs">
                  <CheckCircle2 size={14} />
                </span>
              </div>
              <div className="flex flex-col items-center">
                <span className="text-[10px] text-slate-500 font-bold mb-1">C₂H₂/C₂H₄</span>
                <span className={`w-6 h-6 rounded-full flex items-center justify-center text-xs ${
                  !isHealthyCondition1 && rC2H2_C2H4 > 0.1 ? 'bg-amber-100 text-amber-600' : 'bg-emerald-100 text-emerald-600'
                }`}>
                  {!isHealthyCondition1 && rC2H2_C2H4 > 0.1 ? <AlertTriangle size={13} /> : <CheckCircle2 size={14} />}
                </span>
              </div>
              <div className="flex flex-col items-center">
                <span className="text-[10px] text-slate-500 font-bold mb-1">C₂H₄/C₂H₆</span>
                <span className="w-6 h-6 rounded-full bg-emerald-100 text-emerald-600 flex items-center justify-center text-xs">
                  <CheckCircle2 size={14} />
                </span>
              </div>
              <div className="flex flex-col items-center">
                <span className="text-[10px] text-slate-500 font-bold mb-1">Other</span>
                <span className="w-6 h-6 rounded-full bg-emerald-100 text-emerald-600 flex items-center justify-center text-xs">
                  <CheckCircle2 size={14} />
                </span>
              </div>
            </div>
            <div className="text-[11px] text-slate-500 mt-2 text-center">
              {isHealthyCondition1
                ? 'All ratio alarms normal (Condition 1 - Healthy)'
                : rC2H2_C2H4 > 0.1
                ? 'C₂H₂/C₂H₄ tỷ lệ bất thường'
                : 'All ratio alarms normal'}
            </div>
          </div>
          <div className="mt-4 pt-3 border-t border-slate-100">
            <button
              type="button"
              onClick={() => setShowRatioSettingsModal(true)}
              className="text-xs font-semibold text-blue-600 hover:text-blue-800 flex items-center gap-1 cursor-pointer"
            >
              <span>View ratio settings</span>
              <ChevronRight size={13} />
            </button>
          </div>
        </div>

        {/* Card 4: Overall Risk Level */}
        <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-2xs hover:shadow-sm transition-all flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">Overall Risk Level</span>
              <span className="text-slate-400 hover:text-slate-600 cursor-pointer" onClick={() => setShowRiskAssessmentModal(true)}>
                <Info size={14} />
              </span>
            </div>
            <div className="mt-3 flex items-center gap-3.5">
              {/* Radial Score Gauge */}
              <div className="relative w-12 h-12 flex-shrink-0 flex items-center justify-center">
                <svg className="w-12 h-12 transform -rotate-90">
                  <circle
                    cx="24"
                    cy="24"
                    r="19"
                    stroke="#e2e8f0"
                    strokeWidth="4"
                    fill="transparent"
                  />
                  <circle
                    cx="24"
                    cy="24"
                    r="19"
                    stroke={riskScore > 65 ? '#ef4444' : riskScore > 40 ? '#f59e0b' : '#10b981'}
                    strokeWidth="4"
                    fill="transparent"
                    strokeDasharray="119.38"
                    strokeDashoffset={119.38 - (119.38 * riskScore) / 100}
                    strokeLinecap="round"
                    className="transition-all duration-700"
                  />
                </svg>
                <div className="absolute inset-0 flex flex-col items-center justify-center text-center">
                  <span className="text-xs font-extrabold text-slate-900">{riskScore}</span>
                  <span className="text-[8px] text-slate-400 -mt-1">/100</span>
                </div>
              </div>

              <div>
                <div className={`text-sm font-extrabold ${riskScore > 65 ? 'text-rose-600' : riskScore > 40 ? 'text-amber-600' : 'text-emerald-700'}`}>
                  {riskScore > 65 ? 'High Risk' : riskScore > 40 ? 'Moderate Risk' : 'Low Risk'}
                </div>
                <div className="text-[11px] text-slate-500 mt-0.5">
                  {riskScore > 65 ? 'Immediate inspection' : riskScore > 40 ? 'Review trend closely' : 'Condition stable'}
                </div>
              </div>
            </div>
          </div>
          <div className="mt-4 pt-3 border-t border-slate-100">
            <button
              type="button"
              onClick={() => setShowRiskAssessmentModal(true)}
              className="text-xs font-semibold text-blue-600 hover:text-blue-800 flex items-center gap-1 cursor-pointer"
            >
              <span>Risk assessment</span>
              <ChevronRight size={13} />
            </button>
          </div>
        </div>
      </div>

      {/* 3. MIDDLE SECTION: 3-PANEL VISUALIZATION GRID */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
        {/* Panel 1: Key Gas Trends (ppm) */}
        <div className="lg:col-span-5 bg-white rounded-xl border border-slate-200 p-5 shadow-2xs flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between mb-2">
              <div className="flex items-center gap-2">
                <h3 className="font-bold text-slate-900 text-sm flex items-center gap-1.5">
                  Key Gas Trends (ppm)
                </h3>
                <span className="text-slate-400 hover:text-slate-600 cursor-pointer">
                  <Info size={13} />
                </span>
              </div>
              <div className="flex items-center gap-1 bg-slate-100 p-0.5 rounded-lg border border-slate-200 text-[11px]">
                <button
                  type="button"
                  onClick={() => setIsGasTrendLogScale(true)}
                  className={`px-2 py-0.5 rounded font-bold transition-all ${
                    isGasTrendLogScale ? 'bg-white text-blue-600 shadow-2xs' : 'text-slate-500 hover:text-slate-800'
                  }`}
                >
                  Log scale
                </button>
                <button
                  type="button"
                  onClick={() => setIsGasTrendLogScale(false)}
                  className={`px-2 py-0.5 rounded font-bold transition-all ${
                    !isGasTrendLogScale ? 'bg-white text-blue-600 shadow-2xs' : 'text-slate-500 hover:text-slate-800'
                  }`}
                >
                  Linear
                </button>
              </div>
            </div>

            {/* Gas Pills Legend */}
            <div className="flex flex-wrap items-center gap-2 mb-3 text-[11px] font-semibold text-slate-600">
              <span className="inline-flex items-center gap-1"><span className="w-2.5 h-0.5 bg-[#2563eb] rounded-full inline-block"></span> H₂</span>
              <span className="inline-flex items-center gap-1"><span className="w-2.5 h-0.5 bg-[#16a34a] rounded-full inline-block"></span> CH₄</span>
              <span className="inline-flex items-center gap-1"><span className="w-2.5 h-0.5 bg-[#dc2626] rounded-full inline-block"></span> C₂H₆</span>
              <span className="inline-flex items-center gap-1"><span className="w-2.5 h-0.5 bg-[#9333ea] rounded-full inline-block"></span> C₂H₄</span>
              <span className="inline-flex items-center gap-1"><span className="w-2.5 h-0.5 bg-[#ea580c] rounded-full inline-block"></span> C₂H₂</span>
              <span className="inline-flex items-center gap-1"><span className="w-2.5 h-0.5 bg-[#475569] rounded-full inline-block"></span> CO</span>
            </div>

            {/* LineChart */}
            <div className="h-56 w-full">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={chartTimeline} margin={{ top: 5, right: 10, left: -20, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                  <XAxis dataKey="date" tick={{ fill: '#64748b', fontSize: 10 }} />
                  <YAxis
                    scale={isGasTrendLogScale ? 'log' : 'auto'}
                    domain={isGasTrendLogScale ? [0.1, 10000] : [0, 'auto']}
                    tick={{ fill: '#64748b', fontSize: 10 }}
                    allowDataOverflow
                  />
                  <Tooltip
                    content={({ active, payload, label }) => {
                      if (!active || !payload || !payload.length) return null;
                      return (
                        <div className="bg-slate-900 text-white p-2.5 rounded-lg shadow-xl text-[11px] min-w-[150px] border border-slate-700">
                          <div className="font-bold text-amber-300 border-b border-slate-700 pb-1 mb-1">{label}</div>
                          {payload.map((item: any) => (
                            <div key={item.name} className="flex justify-between py-0.5">
                              <span style={{ color: item.color }}>{item.name}:</span>
                              <span className="font-mono font-bold text-white">{item.value} ppm</span>
                            </div>
                          ))}
                        </div>
                      );
                    }}
                  />
                  <Line type="monotone" dataKey="h2" name="H₂" stroke="#2563eb" strokeWidth={2} dot={{ r: 3 }} />
                  <Line type="monotone" dataKey="ch4" name="CH₄" stroke="#16a34a" strokeWidth={2} dot={{ r: 3 }} />
                  <Line type="monotone" dataKey="c2h6" name="C₂H₆" stroke="#dc2626" strokeWidth={1.8} dot={{ r: 3 }} />
                  <Line type="monotone" dataKey="c2h4" name="C₂H₄" stroke="#9333ea" strokeWidth={1.8} dot={{ r: 3 }} />
                  <Line type="monotone" dataKey="c2h2" name="C₂H₂" stroke="#ea580c" strokeWidth={2.5} dot={{ r: 4 }} />
                  <Line type="monotone" dataKey="co" name="CO" stroke="#475569" strokeWidth={1.5} dot={{ r: 3 }} />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </div>
          <div className="mt-3 pt-2 border-t border-slate-100 flex justify-between items-center text-xs">
            <button
              type="button"
              onClick={() => setShowDgaDetailsModal(true)}
              className="text-blue-600 hover:text-blue-800 font-semibold flex items-center gap-1 cursor-pointer"
            >
              <span>View full gas analysis</span>
              <ChevronRight size={13} />
            </button>
            <span className="text-[10px] text-slate-400">Total: {tdcg.toFixed(0)} ppm</span>
          </div>
        </div>

        {/* Panel 2: Gas Ratio Trending */}
        <div className="lg:col-span-4 bg-white rounded-xl border border-slate-200 p-5 shadow-2xs flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between mb-2">
              <div className="flex items-center gap-2">
                <h3 className="font-bold text-slate-900 text-sm flex items-center gap-1.5">
                  Gas Ratio Trending
                </h3>
                <span className="text-slate-400 hover:text-slate-600 cursor-pointer" onClick={() => setShowRatioSettingsModal(true)}>
                  <Info size={13} />
                </span>
              </div>
              <div className="flex items-center gap-1 bg-slate-100 p-0.5 rounded-lg border border-slate-200 text-[11px]">
                <button
                  type="button"
                  onClick={() => setIsRatioTrendLogScale(true)}
                  className={`px-2 py-0.5 rounded font-bold transition-all ${
                    isRatioTrendLogScale ? 'bg-white text-blue-600 shadow-2xs' : 'text-slate-500 hover:text-slate-800'
                  }`}
                >
                  Log scale
                </button>
                <button
                  type="button"
                  onClick={() => setIsRatioTrendLogScale(false)}
                  className={`px-2 py-0.5 rounded font-bold transition-all ${
                    !isRatioTrendLogScale ? 'bg-white text-blue-600 shadow-2xs' : 'text-slate-500 hover:text-slate-800'
                  }`}
                >
                  Linear
                </button>
              </div>
            </div>

            {/* Ratio Legend & Thresholds */}
            <div className="flex flex-wrap items-center justify-between gap-2 mb-3 text-[11px] font-semibold text-slate-600">
              <div className="flex items-center gap-2">
                <span className="inline-flex items-center gap-1"><span className="w-2.5 h-0.5 bg-[#2563eb] rounded-full inline-block"></span> CH₄/H₂</span>
                <span className="inline-flex items-center gap-1"><span className="w-2.5 h-0.5 bg-[#16a34a] rounded-full inline-block"></span> C₂H₂/C₂H₄</span>
                <span className="inline-flex items-center gap-1"><span className="w-2.5 h-0.5 bg-[#dc2626] rounded-full inline-block"></span> C₂H₄/C₂H₆</span>
              </div>
            </div>

            {/* Ratio LineChart with Alarm & Alert Thresholds */}
            <div className="h-56 w-full relative">
              <ResponsiveContainer width="100%" height="100%">
                <LineChart data={chartTimeline} margin={{ top: 12, right: 10, left: -20, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                  <XAxis dataKey="date" tick={{ fill: '#64748b', fontSize: 10 }} />
                  <YAxis
                    scale={isRatioTrendLogScale ? 'log' : 'auto'}
                    domain={isRatioTrendLogScale ? [0.01, 100] : [0, 'auto']}
                    tick={{ fill: '#64748b', fontSize: 10 }}
                    allowDataOverflow
                  />
                  <ReferenceLine
                    y={10}
                    stroke="#ef4444"
                    strokeDasharray="4 3"
                    strokeWidth={1.5}
                    label={{ value: 'Alarm threshold', fill: '#dc2626', fontSize: 9, position: 'insideTopRight' }}
                  />
                  <ReferenceLine
                    y={1.0}
                    stroke="#f59e0b"
                    strokeDasharray="3 3"
                    strokeWidth={1.2}
                    label={{ value: 'Alert threshold', fill: '#d97706', fontSize: 9, position: 'insideTopRight' }}
                  />
                  <Tooltip
                    content={({ active, payload, label }) => {
                      if (!active || !payload || !payload.length) return null;
                      return (
                        <div className="bg-slate-900 text-white p-2.5 rounded-lg shadow-xl text-[11px] min-w-[140px] border border-slate-700">
                          <div className="font-bold text-amber-300 border-b border-slate-700 pb-1 mb-1">{label}</div>
                          <div className="flex justify-between py-0.5 text-blue-300">
                            <span>CH₄/H₂:</span>
                            <span className="font-mono font-bold">{payload[0]?.payload?.r1}</span>
                          </div>
                          <div className="flex justify-between py-0.5 text-emerald-300">
                            <span>C₂H₂/C₂H₄:</span>
                            <span className="font-mono font-bold">{payload[0]?.payload?.r2}</span>
                          </div>
                          <div className="flex justify-between py-0.5 text-rose-300">
                            <span>C₂H₄/C₂H₆:</span>
                            <span className="font-mono font-bold">{payload[0]?.payload?.r3}</span>
                          </div>
                        </div>
                      );
                    }}
                  />
                  <Line type="monotone" dataKey="r1" name="CH₄/H₂" stroke="#2563eb" strokeWidth={2} dot={{ r: 3 }} />
                  <Line type="monotone" dataKey="r2" name="C₂H₂/C₂H₄" stroke="#16a34a" strokeWidth={2} dot={{ r: 3 }} />
                  <Line type="monotone" dataKey="r3" name="C₂H₄/C₂H₆" stroke="#dc2626" strokeWidth={2} dot={{ r: 3 }} />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </div>
          <div className="mt-3 pt-2 border-t border-slate-100 flex justify-between items-center text-xs">
            <button
              type="button"
              onClick={() => setShowRatioAnalysisModal(true)}
              className="text-blue-600 hover:text-blue-800 font-semibold flex items-center gap-1 cursor-pointer"
            >
              <span>View ratio analysis</span>
              <ChevronRight size={13} />
            </button>
            <span className="text-[10px] text-slate-400">IEC 60599 / Rogers</span>
          </div>
        </div>

        {/* Panel 3: Active Events / DGA Alarm Summary */}
        <div className="lg:col-span-3 bg-white rounded-xl border border-slate-200 p-5 shadow-2xs flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between mb-3">
              <div className="flex items-center gap-1.5">
                <h3 className="font-bold text-slate-900 text-sm">Active Events</h3>
                <span className="text-slate-400" onClick={() => setShowEventLogModal(true)}>
                  <Info size={13} />
                </span>
              </div>
              <button
                type="button"
                onClick={() => setShowEventLogModal(true)}
                className="text-xs font-bold text-blue-600 hover:text-blue-800 cursor-pointer"
              >
                View all
              </button>
            </div>

            <div className="text-[11px] font-bold text-slate-400 uppercase tracking-wider mb-2">
              DGA Alarm Summary
            </div>

            {/* Chronological Event Items */}
            <div className="space-y-3">
              <div className="flex items-start gap-2.5">
                <span className="w-5 h-5 rounded-full bg-emerald-100 text-emerald-600 flex items-center justify-center flex-shrink-0 mt-0.5">
                  <CheckCircle2 size={13} />
                </span>
                <div className="flex-1 min-w-0">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-800 truncate">All gas levels normal</span>
                    <span className="text-[10px] text-slate-400 flex-shrink-0">May 14, 08:15</span>
                  </div>
                  <p className="text-[11px] text-slate-500 line-clamp-1">No gas concentrations above normal limits</p>
                </div>
              </div>

              <div className="flex items-start gap-2.5">
                <span className="w-5 h-5 rounded-full bg-emerald-100 text-emerald-600 flex items-center justify-center flex-shrink-0 mt-0.5">
                  <CheckCircle2 size={13} />
                </span>
                <div className="flex-1 min-w-0">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-800 truncate">Gas ratios normal</span>
                    <span className="text-[10px] text-slate-400 flex-shrink-0">May 14, 08:15</span>
                  </div>
                  <p className="text-[11px] text-slate-500 line-clamp-1">All monitored ratios within limits</p>
                </div>
              </div>

              <div className="flex items-start gap-2.5">
                <span className="w-5 h-5 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center flex-shrink-0 mt-0.5">
                  <Info size={13} />
                </span>
                <div className="flex-1 min-w-0">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-800 truncate">Moisture level normal</span>
                    <span className="text-[10px] text-slate-400 flex-shrink-0">May 13, 22:41</span>
                  </div>
                  <p className="text-[11px] text-slate-500 line-clamp-1">Moisture content within normal range</p>
                </div>
              </div>

              <div className="flex items-start gap-2.5">
                <span className="w-5 h-5 rounded-full bg-amber-100 text-amber-600 flex items-center justify-center flex-shrink-0 mt-0.5">
                  <AlertTriangle size={13} />
                </span>
                <div className="flex-1 min-w-0">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-amber-800 truncate">CO level elevated</span>
                    <span className="text-[10px] text-slate-400 flex-shrink-0">May 13, 18:05</span>
                  </div>
                  <p className="text-[11px] text-slate-500 line-clamp-1">CO trending upward but below alarm</p>
                </div>
              </div>

              <div className="flex items-start gap-2.5">
                <span className="w-5 h-5 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center flex-shrink-0 mt-0.5">
                  <Info size={13} />
                </span>
                <div className="flex-1 min-w-0">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-800 truncate">New sample collected</span>
                    <span className="text-[10px] text-slate-400 flex-shrink-0">May 13, 08:15</span>
                  </div>
                  <p className="text-[11px] text-slate-500 line-clamp-1">Online DGA sample successfully collected</p>
                </div>
              </div>
            </div>
          </div>
          <div className="mt-3 pt-2 border-t border-slate-100">
            <span className="text-xs font-semibold text-blue-600 hover:text-blue-800 flex items-center gap-1 cursor-pointer">
              <span>View event log</span>
              <ChevronRight size={13} />
            </span>
          </div>
        </div>
      </div>

      {/* 4. ADVANCED DIAGNOSTICS METHODS TITLE & 5 GRAPHICAL CARDS */}
      <div>
        <div className="flex items-center justify-between mb-4">
          <h2 className="text-lg font-bold text-slate-900 tracking-tight">
            Advanced diagnostics methods
          </h2>
          <span className="text-xs text-slate-500 font-medium">Click method to focus or drill down</span>
        </div>

        {/* 5 Cards Grid */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
          {/* Method 1: Duval Triangle */}
          <div
            onClick={() => setActiveMethodTab('triangle')}
            className={`bg-white rounded-xl border p-4 shadow-2xs cursor-pointer transition-all flex flex-col justify-between ${
              activeMethodTab === 'triangle' ? 'border-blue-500 ring-2 ring-blue-100' : 'border-slate-200 hover:border-slate-300'
            }`}
          >
            <div>
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs font-bold text-slate-800">Duval Triangle</span>
                <span className="text-[10px] text-blue-600 font-bold bg-blue-50 px-1.5 py-0.5 rounded">T1</span>
              </div>
              {/* Mini SVG Triangle Preview with Active Dot */}
              <div className="relative w-full aspect-[4/3] flex items-center justify-center bg-slate-50 rounded-lg p-2 border border-slate-100">
                <svg viewBox="0 0 100 100" className="w-full h-full max-h-24">
                  <polygon points="50,10 90,85 10,85" fill="#f8fafc" stroke="#cbd5e1" strokeWidth="1" />
                  {/* PD zone top */}
                  <polygon points="50,10 52,18 48,18" fill="#38bdf8" />
                  {/* T1, T2, T3 zones */}
                  <polygon points="48,18 35,45 65,45 52,18" fill="#4ade80" />
                  <polygon points="35,45 25,65 75,65 65,45" fill="#facc15" />
                  <polygon points="25,65 40,85 60,85 75,65" fill="#fb923c" />
                  {/* D1, D2 zones */}
                  <polygon points="10,85 25,65 40,85" fill="#f43f5e" />
                  <polygon points="90,85 75,65 60,85" fill="#ef4444" />
                  {/* Plotted Dot */}
                  <circle cx={triangleCoords.x} cy={triangleCoords.y} r="3" fill="#1e293b" stroke="#ffffff" strokeWidth="1" />
                </svg>
              </div>
            </div>
            <div className="mt-3 pt-2 border-t border-slate-100">
              <div className="text-[10px] text-slate-400 font-medium">Interpretation</div>
              <div className="text-xs font-bold text-slate-800 flex items-center gap-1.5 mt-0.5">
                <span className="w-2 h-2 rounded-full bg-blue-500 inline-block"></span>
                <span>{duvalT1Interpretation}</span>
              </div>
            </div>
          </div>

          {/* Method 2: Duval Pentagon */}
          <div
            onClick={() => setActiveMethodTab('pentagon')}
            className={`bg-white rounded-xl border p-4 shadow-2xs cursor-pointer transition-all flex flex-col justify-between ${
              activeMethodTab === 'pentagon' ? 'border-blue-500 ring-2 ring-blue-100' : 'border-slate-200 hover:border-slate-300'
            }`}
          >
            <div>
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs font-bold text-slate-800">Duval Pentagon</span>
                <span className="text-[10px] text-blue-600 font-bold bg-blue-50 px-1.5 py-0.5 rounded">P1</span>
              </div>
              {/* Mini SVG Pentagon Preview */}
              <div className="relative w-full aspect-[4/3] flex items-center justify-center bg-slate-50 rounded-lg p-2 border border-slate-100">
                <svg viewBox="0 0 100 100" className="w-full h-full max-h-24">
                  {/* Pentagon shape */}
                  <polygon points="50,15 85,40 70,80 30,80 15,40" fill="#f1f5f9" stroke="#94a3b8" strokeWidth="0.8" />
                  <polygon points="50,25 75,45 65,72 35,72 25,45" fill="#c7d2fe" opacity="0.6" />
                  <polygon points="50,35 65,48 60,65 40,65 35,48" fill="#a7f3d0" opacity="0.7" />
                  {/* Plotted Dot */}
                  <circle cx={pentagonCoords.x} cy={pentagonCoords.y} r="3" fill="#0f172a" stroke="#ffffff" strokeWidth="1" />
                </svg>
              </div>
            </div>
            <div className="mt-3 pt-2 border-t border-slate-100">
              <div className="text-[10px] text-slate-400 font-medium">Interpretation</div>
              <div className="text-xs font-bold text-slate-800 flex items-center gap-1.5 mt-0.5">
                <span className="w-2 h-2 rounded-full bg-emerald-500 inline-block"></span>
                <span>{duvalP1Interpretation}</span>
              </div>
            </div>
          </div>

          {/* Method 3: Rogers Ratio */}
          <div
            onClick={() => setActiveMethodTab('rogers')}
            className={`bg-white rounded-xl border p-4 shadow-2xs cursor-pointer transition-all flex flex-col justify-between ${
              activeMethodTab === 'rogers' ? 'border-blue-500 ring-2 ring-blue-100' : 'border-slate-200 hover:border-slate-300'
            }`}
          >
            <div>
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs font-bold text-slate-800">Rogers Ratio</span>
                <span className="text-[10px] text-blue-600 font-bold bg-blue-50 px-1.5 py-0.5 rounded">IEC</span>
              </div>
              <div className="bg-slate-50 rounded-lg p-2.5 border border-slate-100 space-y-1 text-xs">
                <div className="flex justify-between">
                  <span className="text-slate-500">R1 (CH₄/H₂):</span>
                  <span className="font-mono font-bold text-slate-800">{rCH4_H2.toFixed(2)}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500">R2 (C₂H₂/C₂H₄):</span>
                  <span className="font-mono font-bold text-slate-800">{rC2H2_C2H4.toFixed(2)}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500">R3 (C₂H₄/C₂H₆):</span>
                  <span className="font-mono font-bold text-slate-800">{rC2H4_C2H6.toFixed(2)}</span>
                </div>
              </div>
            </div>
            <div className="mt-3 pt-2 border-t border-slate-100">
              <div className="text-[10px] text-slate-400 font-medium">Interpretation</div>
              <div className="text-xs font-bold text-slate-800 flex items-center gap-1.5 mt-0.5">
                <span className="w-2 h-2 rounded-full bg-emerald-500 inline-block"></span>
                <span>{rogersInterpretation}</span>
              </div>
            </div>
          </div>

          {/* Method 4: Key Gases (ppm) */}
          <div
            onClick={() => setActiveMethodTab('keygas')}
            className={`bg-white rounded-xl border p-4 shadow-2xs cursor-pointer transition-all flex flex-col justify-between ${
              activeMethodTab === 'keygas' ? 'border-blue-500 ring-2 ring-blue-100' : 'border-slate-200 hover:border-slate-300'
            }`}
          >
            <div>
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs font-bold text-slate-800">Key Gas (ppm)</span>
                <span className="text-[10px] text-blue-600 font-bold bg-blue-50 px-1.5 py-0.5 rounded">TDCG</span>
              </div>
              <div className="bg-slate-50 rounded-lg p-2 border border-slate-100 grid grid-cols-1 sm:grid-cols-2 gap-x-2 gap-y-1 text-[11px]">
                {/* H2 */}
                {(() => {
                  const tr = computeDgaParamTrend(h2, histRecord.gases.h2, false);
                  return (
                    <div className="flex items-center justify-between">
                      <span className="text-slate-500 flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-blue-600"></span>H₂:</span>
                      <div className="flex items-center gap-1">
                        <span className="font-mono font-bold">{h2}</span>
                        {tr.hasHistorical && (
                          <span className={`inline-flex items-center text-[9px] font-bold px-1 py-0.2 rounded font-mono ${
                            tr.severity === 'worse' ? 'text-rose-700 bg-rose-50 border border-rose-200' :
                            tr.severity === 'better' ? 'text-emerald-700 bg-emerald-50 border border-emerald-200' :
                            'text-slate-500 bg-slate-100'
                          }`} title={`So với lần trước (${histRecord.gases.h2}): ${tr.diff > 0 ? '+' : ''}${tr.diff.toFixed(1)} (${tr.percent > 0 ? '+' : ''}${tr.percent.toFixed(0)}%)`}>
                            {tr.direction === 'up' && <ArrowUp size={9} />}
                            {tr.direction === 'down' && <ArrowDown size={9} />}
                            {tr.direction === 'flat' && <ArrowRight size={9} />}
                            {tr.direction === 'up' ? `+${Math.abs(tr.diff).toFixed(0)}` : tr.direction === 'down' ? `-${Math.abs(tr.diff).toFixed(0)}` : '0'}
                          </span>
                        )}
                      </div>
                    </div>
                  );
                })()}

                {/* CH4 */}
                {(() => {
                  const tr = computeDgaParamTrend(ch4, histRecord.gases.ch4, false);
                  return (
                    <div className="flex items-center justify-between">
                      <span className="text-slate-500 flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-emerald-600"></span>CH₄:</span>
                      <div className="flex items-center gap-1">
                        <span className="font-mono font-bold">{ch4}</span>
                        {tr.hasHistorical && (
                          <span className={`inline-flex items-center text-[9px] font-bold px-1 py-0.2 rounded font-mono ${
                            tr.severity === 'worse' ? 'text-rose-700 bg-rose-50 border border-rose-200' :
                            tr.severity === 'better' ? 'text-emerald-700 bg-emerald-50 border border-emerald-200' :
                            'text-slate-500 bg-slate-100'
                          }`} title={`So với lần trước (${histRecord.gases.ch4}): ${tr.diff > 0 ? '+' : ''}${tr.diff.toFixed(1)} (${tr.percent > 0 ? '+' : ''}${tr.percent.toFixed(0)}%)`}>
                            {tr.direction === 'up' && <ArrowUp size={9} />}
                            {tr.direction === 'down' && <ArrowDown size={9} />}
                            {tr.direction === 'flat' && <ArrowRight size={9} />}
                            {tr.direction === 'up' ? `+${Math.abs(tr.diff).toFixed(0)}` : tr.direction === 'down' ? `-${Math.abs(tr.diff).toFixed(0)}` : '0'}
                          </span>
                        )}
                      </div>
                    </div>
                  );
                })()}

                {/* C2H6 */}
                {(() => {
                  const tr = computeDgaParamTrend(c2h6, histRecord.gases.c2h6, false);
                  return (
                    <div className="flex items-center justify-between">
                      <span className="text-slate-500 flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-red-600"></span>C₂H₆:</span>
                      <div className="flex items-center gap-1">
                        <span className="font-mono font-bold">{c2h6}</span>
                        {tr.hasHistorical && (
                          <span className={`inline-flex items-center text-[9px] font-bold px-1 py-0.2 rounded font-mono ${
                            tr.severity === 'worse' ? 'text-rose-700 bg-rose-50 border border-rose-200' :
                            tr.severity === 'better' ? 'text-emerald-700 bg-emerald-50 border border-emerald-200' :
                            'text-slate-500 bg-slate-100'
                          }`} title={`So với lần trước (${histRecord.gases.c2h6}): ${tr.diff > 0 ? '+' : ''}${tr.diff.toFixed(1)} (${tr.percent > 0 ? '+' : ''}${tr.percent.toFixed(0)}%)`}>
                            {tr.direction === 'up' && <ArrowUp size={9} />}
                            {tr.direction === 'down' && <ArrowDown size={9} />}
                            {tr.direction === 'flat' && <ArrowRight size={9} />}
                            {tr.direction === 'up' ? `+${Math.abs(tr.diff).toFixed(0)}` : tr.direction === 'down' ? `-${Math.abs(tr.diff).toFixed(0)}` : '0'}
                          </span>
                        )}
                      </div>
                    </div>
                  );
                })()}

                {/* C2H4 */}
                {(() => {
                  const tr = computeDgaParamTrend(c2h4, histRecord.gases.c2h4, false);
                  return (
                    <div className="flex items-center justify-between">
                      <span className="text-slate-500 flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-purple-600"></span>C₂H₄:</span>
                      <div className="flex items-center gap-1">
                        <span className="font-mono font-bold">{c2h4}</span>
                        {tr.hasHistorical && (
                          <span className={`inline-flex items-center text-[9px] font-bold px-1 py-0.2 rounded font-mono ${
                            tr.severity === 'worse' ? 'text-rose-700 bg-rose-50 border border-rose-200' :
                            tr.severity === 'better' ? 'text-emerald-700 bg-emerald-50 border border-emerald-200' :
                            'text-slate-500 bg-slate-100'
                          }`} title={`So với lần trước (${histRecord.gases.c2h4}): ${tr.diff > 0 ? '+' : ''}${tr.diff.toFixed(1)} (${tr.percent > 0 ? '+' : ''}${tr.percent.toFixed(0)}%)`}>
                            {tr.direction === 'up' && <ArrowUp size={9} />}
                            {tr.direction === 'down' && <ArrowDown size={9} />}
                            {tr.direction === 'flat' && <ArrowRight size={9} />}
                            {tr.direction === 'up' ? `+${Math.abs(tr.diff).toFixed(0)}` : tr.direction === 'down' ? `-${Math.abs(tr.diff).toFixed(0)}` : '0'}
                          </span>
                        )}
                      </div>
                    </div>
                  );
                })()}

                {/* C2H2 */}
                {(() => {
                  const tr = computeDgaParamTrend(c2h2, histRecord.gases.c2h2, false);
                  return (
                    <div className="flex items-center justify-between col-span-1 sm:col-span-2 pt-0.5 border-t border-slate-200">
                      <span className="text-rose-600 font-bold flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-rose-600"></span>C₂H₂ (Hồ quang):</span>
                      <div className="flex items-center gap-1.5">
                        <span className="font-mono font-bold text-rose-600">{c2h2} ppm</span>
                        {tr.hasHistorical && (
                          <span className={`inline-flex items-center gap-0.5 text-[9px] font-bold px-1.5 py-0.2 rounded font-mono ${
                            tr.severity === 'worse' ? 'text-rose-700 bg-rose-50 border border-rose-300' :
                            tr.severity === 'better' ? 'text-emerald-700 bg-emerald-50 border border-emerald-300' :
                            'text-slate-600 bg-slate-100'
                          }`} title={`So với lần trước (${histRecord.gases.c2h2}): ${tr.diff > 0 ? '+' : ''}${tr.diff.toFixed(2)} (${tr.percent > 0 ? '+' : ''}${tr.percent.toFixed(1)}%)`}>
                            {tr.direction === 'up' && <ArrowUp size={9} className="stroke-[2.5]" />}
                            {tr.direction === 'down' && <ArrowDown size={9} className="stroke-[2.5]" />}
                            {tr.direction === 'flat' && <ArrowRight size={9} />}
                            <span>{tr.diff > 0 ? `+${tr.diff.toFixed(1)}` : tr.diff < 0 ? `-${Math.abs(tr.diff).toFixed(1)}` : '0.0'}</span>
                          </span>
                        )}
                      </div>
                    </div>
                  );
                })()}
              </div>
            </div>
            <div className="mt-3 pt-2 border-t border-slate-100">
              <div className="text-[10px] text-slate-400 font-medium">Total combustible gas</div>
              <div className="text-xs font-bold text-slate-900 mt-0.5">
                {tdcg.toFixed(0)} ppm
              </div>
            </div>
          </div>

          {/* Method 5: Doernenburg */}
          <div
            onClick={() => setActiveMethodTab('dornenburg')}
            className={`bg-white rounded-xl border p-4 shadow-2xs cursor-pointer transition-all flex flex-col justify-between ${
              activeMethodTab === 'dornenburg' ? 'border-blue-500 ring-2 ring-blue-100' : 'border-slate-200 hover:border-slate-300'
            }`}
          >
            <div>
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs font-bold text-slate-800">Doernenburg</span>
                <span className="text-[10px] text-blue-600 font-bold bg-blue-50 px-1.5 py-0.5 rounded">IEEE</span>
              </div>
              <div className="bg-slate-50 rounded-lg p-2.5 border border-slate-100 space-y-1 text-xs">
                <div className="flex justify-between">
                  <span className="text-slate-500">CH₄/H₂:</span>
                  <span className="font-mono font-bold text-slate-800">{rCH4_H2.toFixed(2)}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500">C₂H₂/C₂H₄:</span>
                  <span className="font-mono font-bold text-slate-800">{rC2H2_C2H4.toFixed(2)}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500">C₂H₂/CH₄:</span>
                  <span className="font-mono font-bold text-slate-800">{rC2H2_CH4.toFixed(2)}</span>
                </div>
              </div>
            </div>
            <div className="mt-3 pt-2 border-t border-slate-100">
              <div className="text-[10px] text-slate-400 font-medium">Interpretation</div>
              <div className="text-xs font-bold text-slate-800 flex items-center gap-1.5 mt-0.5">
                <span className="w-2 h-2 rounded-full bg-emerald-500 inline-block"></span>
                <span>{dornenburgInterpretation}</span>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* 5. DEEP-DIVE: TRACK THE RELATIONSHIP BETWEEN GASES (Image 3) */}
      <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
        <div className="max-w-3xl mb-6">
          <h2 className="text-xl font-bold text-slate-900 tracking-tight">
            Track the relationship between gases
          </h2>
          <p className="text-xs text-slate-600 mt-1.5 leading-relaxed">
            Gas ratios remove the effect of absolute concentration and highlight the underlying chemical and thermal processes occurring inside the transformer. Advanced DGA combines automatic ratio calculation, continuous trending and clear alarm visualization to support accurate diagnostics and confident engineering interpretation.
          </p>
        </div>

        {isHealthyCondition1 && (
          <div className="mb-6 p-4 rounded-xl bg-emerald-50/90 border border-emerald-200 flex items-start gap-3 shadow-xs">
            <CheckCircle2 size={22} className="text-emerald-600 flex-shrink-0 mt-0.5" />
            <div className="text-xs text-emerald-950">
              <span className="font-bold text-sm block text-emerald-900">
                Tình trạng thiết bị: NORMAL / HEALTHY (Condition 1 — IEEE C57.104 &amp; IEC 60599)
              </span>
              <p className="mt-1 text-emerald-800 leading-relaxed">
                Tổng khí cháy TDCG = <strong>{tdcg.toFixed(1)} ppm</strong> (≤ 720 ppm) và toàn bộ các khí riêng lẻ đều nằm dưới ngưỡng cảnh báo giới hạn (H₂&lt;100, CH₄&lt;120, C₂H₆&lt;65, C₂H₄&lt;50, C₂H₂&lt;35, CO&lt;350 ppm).
                Theo quy chuẩn chẩn đoán quốc tế, hệ thống tự động <strong>bỏ qua phân loại lỗi</strong> (Duval Triangles, Duval Pentagons, Rogers, IEC, Dornenburg) để loại trừ triệt để báo động giả do tạp âm đo lường ở nồng độ cực thấp.
              </p>
              <div className="mt-2 inline-flex items-center gap-1.5 font-bold text-emerald-700 bg-emerald-100/80 px-2.5 py-1 rounded-md text-[11px]">
                <span>Khuyến cáo: Duy trì chế độ vận hành bình thường và lấy mẫu định kỳ hàng năm (Annual sampling).</span>
              </div>
            </div>
          </div>
        )}

        {/* Ratio Trend Example with Alert & Alarm Milestones */}
        <div className="bg-slate-50/80 rounded-xl p-4 border border-slate-200 mb-6">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-3">
            <div className="text-xs font-bold text-slate-800 flex items-center gap-2">
              <span>Ratio trend example — C₂H₂/C₂H₄</span>
              <span className="text-[10px] font-semibold text-blue-600 bg-blue-50 px-2 py-0.5 rounded border border-blue-200">
                Current: {rC2H2_C2H4.toFixed(2)}
              </span>
            </div>
            <div className="flex items-center gap-4 text-[11px] font-semibold text-slate-500">
              <span className="flex items-center gap-1.5">
                <span className="w-3 h-0.5 bg-blue-600 inline-block"></span> C₂H₂/C₂H₄ ratio
              </span>
              <span className="flex items-center gap-1.5">
                <span className="w-3 h-0.5 bg-amber-500 border-dashed inline-block"></span> Alarm level 2 (Alert)
              </span>
              <span className="flex items-center gap-1.5">
                <span className="w-3 h-0.5 bg-rose-600 border-dashed inline-block"></span> Alarm level 1 (Alarm)
              </span>
            </div>
          </div>

          <div className="h-44 w-full">
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={chartTimeline} margin={{ top: 15, right: 20, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="date" tick={{ fill: '#64748b', fontSize: 10 }} />
                <YAxis scale="log" domain={[0.001, 10]} tick={{ fill: '#64748b', fontSize: 10 }} />
                <ReferenceLine y={1.0} stroke="#ef4444" strokeDasharray="4 3" strokeWidth={1.5} label={{ value: 'Level 1 alarm', fill: '#dc2626', fontSize: 10 }} />
                <ReferenceLine y={0.1} stroke="#f59e0b" strokeDasharray="3 3" strokeWidth={1.2} label={{ value: 'Level 2 alert', fill: '#d97706', fontSize: 10 }} />
                <Tooltip />
                <Line type="monotone" dataKey="r2" name="C₂H₂/C₂H₄" stroke="#2563eb" strokeWidth={2.5} dot={{ r: 4 }} />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* 2-Column Ratio Explanations & Trending States */}
        <div className="grid grid-cols-1 md:grid-cols-12 gap-6">
          {/* Col 1: Key gas ratios cards */}
          <div className="md:col-span-7 space-y-3">
            <h4 className="text-sm font-bold text-slate-900 border-b border-slate-100 pb-2 flex items-center gap-2">
              <Sliders size={16} className="text-blue-600" />
              Key gas ratios &amp; Diagnostic Indications
            </h4>

            <div className="divide-y divide-slate-100 text-xs">
              <div className="py-2.5 flex items-start gap-3">
                <span className="w-20 px-2 py-1 rounded bg-blue-50 text-blue-700 font-mono font-bold text-center flex-shrink-0">
                  CH₄/H₂
                </span>
                <div className="flex-1">
                  <div className="font-bold text-slate-800">Thermal severity indicator</div>
                  <p className="text-slate-500 text-[11px] mt-0.5">Higher values suggest elevated temperature stress inside the tank.</p>
                </div>
              </div>

              <div className="py-2.5 flex items-start gap-3">
                <span className="w-20 px-2 py-1 rounded bg-rose-50 text-rose-700 font-mono font-bold text-center flex-shrink-0">
                  C₂H₂/C₂H₄
                </span>
                <div className="flex-1">
                  <div className="font-bold text-slate-800">Arcing activity indicator</div>
                  <p className="text-slate-500 text-[11px] mt-0.5">Effective for identifying high-energy electrical discharge conditions.</p>
                </div>
              </div>

              <div className="py-2.5 flex items-start gap-3">
                <span className="w-20 px-2 py-1 rounded bg-amber-50 text-amber-700 font-mono font-bold text-center flex-shrink-0">
                  C₂H₂/CH₄
                </span>
                <div className="flex-1">
                  <div className="font-bold text-slate-800">Discharge vs. thermal indicator</div>
                  <p className="text-slate-500 text-[11px] mt-0.5">Helps distinguish electrical arcing from severe overheating.</p>
                </div>
              </div>

              <div className="py-2.5 flex items-start gap-3">
                <span className="w-20 px-2 py-1 rounded bg-indigo-50 text-indigo-700 font-mono font-bold text-center flex-shrink-0">
                  C₂H₆/C₂H₂
                </span>
                <div className="flex-1">
                  <div className="font-bold text-slate-800">Arcing intensity indicator</div>
                  <p className="text-slate-500 text-[11px] mt-0.5">Higher values indicate lower energy discharge or incipient faults.</p>
                </div>
              </div>

              <div className="py-2.5 flex items-start gap-3">
                <span className="w-20 px-2 py-1 rounded bg-purple-50 text-purple-700 font-mono font-bold text-center flex-shrink-0">
                  C₂H₄/C₂H₆
                </span>
                <div className="flex-1">
                  <div className="font-bold text-slate-800">Thermal cracking indicator</div>
                  <p className="text-slate-500 text-[11px] mt-0.5">Higher values point to higher temperature thermal faults (&gt; 700°C).</p>
                </div>
              </div>

              <div className="py-2.5 flex items-start gap-3">
                <span className="w-20 px-2 py-1 rounded bg-slate-100 text-slate-700 font-mono font-bold text-center flex-shrink-0">
                  CO₂/CO
                </span>
                <div className="flex-1">
                  <div className="font-bold text-slate-800">Cellulose aging indicator</div>
                  <p className="text-slate-500 text-[11px] mt-0.5">Reflects paper oxidation and degradation processes (&lt; 3 indicates severe paper decomposition).</p>
                </div>
              </div>
            </div>
          </div>

          {/* Col 2: Trending & Alarm Philosophy */}
          <div className="md:col-span-5 space-y-4">
            <div className="bg-slate-50 rounded-xl p-4 border border-slate-200">
              <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wider mb-3">Gas ratio trending</h4>
              <div className="space-y-3 text-xs">
                <div className="flex items-start gap-2.5">
                  <span className="w-6 h-6 rounded-full bg-emerald-100 text-emerald-600 flex items-center justify-center flex-shrink-0 mt-0.5">
                    <CheckCircle2 size={14} />
                  </span>
                  <div>
                    <div className="font-bold text-slate-800">Stable</div>
                    <p className="text-[11px] text-slate-500">Ratio remains within normal range with minimal fluctuation. Transformer is operating normally.</p>
                  </div>
                </div>

                <div className="flex items-start gap-2.5">
                  <span className="w-6 h-6 rounded-full bg-amber-100 text-amber-600 flex items-center justify-center flex-shrink-0 mt-0.5">
                    <TrendingUp size={14} />
                  </span>
                  <div>
                    <div className="font-bold text-slate-800">Developing</div>
                    <p className="text-[11px] text-slate-500">Ratios show a gradual upward trend. Continue monitoring and review related gas trends.</p>
                  </div>
                </div>

                <div className="flex items-start gap-2.5">
                  <span className="w-6 h-6 rounded-full bg-rose-100 text-rose-600 flex items-center justify-center flex-shrink-0 mt-0.5">
                    <Flame size={14} />
                  </span>
                  <div>
                    <div className="font-bold text-slate-800">Rapidly changing</div>
                    <p className="text-[11px] text-slate-500">Ratios change quickly or exceed alarm thresholds. Investigate promptly and correlate with gas data.</p>
                  </div>
                </div>
              </div>
            </div>

            <div className="bg-slate-50 rounded-xl p-4 border border-slate-200">
              <h4 className="text-xs font-bold text-slate-800 uppercase tracking-wider mb-2">Gas ratio alarms</h4>
              <ul className="text-[11px] text-slate-600 space-y-2">
                <li><strong>Threshold-based awareness:</strong> Configurable alarm levels (Level 1 &amp; 2) highlight ratios that exceed defined limits.</li>
                <li><strong>Early indication:</strong> Ratio alarms can provide earlier warning than individual gas concentration alarms.</li>
                <li><strong>Correlated insight:</strong> Alarm events are correlated with gas concentrations, trends and other ratios for accurate interpretation.</li>
              </ul>
            </div>
          </div>
        </div>

        {/* Pro Tip Callout Box */}
        <div className="mt-6 p-4 rounded-xl bg-gradient-to-r from-blue-50 to-indigo-50 border border-blue-200/80 flex items-center gap-3">
          <div className="w-9 h-9 rounded-lg bg-blue-600 text-white flex items-center justify-center flex-shrink-0 shadow-xs">
            <Sparkles size={18} />
          </div>
          <div className="text-xs">
            <span className="font-extrabold text-blue-900">Gas concentration + gas ratio + trend context = better diagnostic visibility.</span>
            <p className="text-blue-700 mt-0.5">Understand the why behind the numbers and take informed engineering action.</p>
          </div>
        </div>
      </div>

      {/* 6. REFERENCE FRAMEWORK (Image 1 & 4) */}
      <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
        <h3 className="text-sm font-bold text-slate-900 uppercase tracking-wider mb-4">
          Reference framework
        </h3>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <div className="p-4 rounded-xl bg-slate-50 border border-slate-200">
            <div className="text-base font-extrabold text-blue-700">IEC 60599</div>
            <p className="text-xs text-slate-600 mt-1 leading-relaxed">
              Mineral oil-filled electrical equipment — Guidance on the interpretation of gases in oil-immersed equipment.
            </p>
          </div>

          <div className="p-4 rounded-xl bg-slate-50 border border-slate-200">
            <div className="text-base font-extrabold text-blue-700">IEC 60567</div>
            <p className="text-xs text-slate-600 mt-1 leading-relaxed">
              Mineral oil-filled electrical equipment — Sampling of free and dissolved gases and analysis by gas chromatography.
            </p>
          </div>

          <div className="p-4 rounded-xl bg-slate-50 border border-slate-200">
            <div className="text-base font-extrabold text-blue-700">IEEE C57.104</div>
            <p className="text-xs text-slate-600 mt-1 leading-relaxed">
              Guide for the interpretation of Gases Generated in Oil-Immersed Transformers.
            </p>
          </div>
        </div>
      </div>

      {/* 7. BOTTOM ACTION CARDS: ANNOTATIONS, REPORT, CSV IMPORT/EXPORT (Image 2 & 5) */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
        {/* Card 1: Trend Chart Annotations */}
        <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-2xs flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between mb-3">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-2">
                <TrendingUp size={15} className="text-purple-600" />
                Trend Chart Annotations
              </span>
              <span className="text-slate-400">
                <Info size={13} />
              </span>
            </div>
            <div className="space-y-2 text-xs">
              {annotations.slice(0, 3).map(item => (
                <div key={item.id} className="border-b border-slate-100 pb-1.5">
                  <div className="text-[10px] text-slate-400">{item.date}</div>
                  <div className="text-slate-700 font-medium line-clamp-1">{item.title}</div>
                </div>
              ))}
            </div>
          </div>
          <div className="mt-4 pt-2 border-t border-slate-100">
            <button
              type="button"
              onClick={() => setShowAnnotationsModal(true)}
              className="text-xs font-semibold text-blue-600 hover:text-blue-800 flex items-center gap-1"
            >
              <span>Manage annotations</span>
              <ChevronRight size={13} />
            </button>
          </div>
        </div>

        {/* Card 2: Transformer Overview Report */}
        <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-2xs flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between mb-3">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-2">
                <FileText size={15} className="text-blue-600" />
                Transformer Overview Report
              </span>
              <span className="text-slate-400">
                <Info size={13} />
              </span>
            </div>
            <div className="space-y-1 text-xs">
              <div className="flex justify-between py-1 border-b border-slate-100">
                <span className="text-slate-500">Generated on</span>
                <span className="font-semibold text-slate-800">May 14, 2025 08:20</span>
              </div>
              <div className="flex justify-between py-1 border-b border-slate-100">
                <span className="text-slate-500">Report type</span>
                <span className="font-semibold text-slate-800">DGA Summary Report</span>
              </div>
              <div className="flex justify-between py-1">
                <span className="text-slate-500">Time range</span>
                <span className="font-semibold text-slate-800">May 8 – May 14, 2025</span>
              </div>
            </div>
          </div>
          <div className="mt-4 pt-2 border-t border-slate-100 flex items-center justify-between text-xs">
            <button
              type="button"
              onClick={() => setShowReportModal(true)}
              className="font-semibold text-blue-600 hover:text-blue-800 flex items-center gap-1 cursor-pointer"
            >
              <span>Open report</span>
              <span>→</span>
            </button>
            <button
              type="button"
              onClick={() => {
                if (onExportPdf) onExportPdf();
                else setShowReportModal(true);
              }}
              className="inline-flex items-center gap-1 font-semibold text-slate-700 hover:text-slate-900 bg-slate-100 px-2.5 py-1 rounded border border-slate-200 cursor-pointer"
            >
              <Download size={13} />
              Download PDF
            </button>
          </div>
        </div>

        {/* Card 3: CSV Import / Export */}
        <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-2xs flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between mb-3">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-2">
                <Upload size={15} className="text-emerald-600" />
                CSV / Excel Import &amp; Export
              </span>
              <span className="text-slate-400">
                <Info size={13} />
              </span>
            </div>
            <p className="text-xs text-slate-500 mb-3 leading-relaxed">
              Export latest DGA dataset to CSV/Excel, or upload historical chromatography records for trending.
            </p>
            <div className="flex gap-2">
              <button
                type="button"
                onClick={() => { setImportExportTab('export'); setShowImportExportModal(true); }}
                className="flex-1 py-1.5 px-2.5 text-xs font-semibold rounded-lg bg-blue-50 text-blue-700 hover:bg-blue-100 border border-blue-200 transition-colors text-center"
              >
                Export CSV
              </button>
              <button
                type="button"
                onClick={() => { setImportExportTab('import'); setShowImportExportModal(true); }}
                className="flex-1 py-1.5 px-2.5 text-xs font-semibold rounded-lg bg-slate-100 text-slate-700 hover:bg-slate-200 border border-slate-200 transition-colors text-center"
              >
                Import CSV
              </button>
            </div>
          </div>
          <div className="mt-4 pt-2 border-t border-slate-100">
            <span
              onClick={() => { setImportExportTab('import'); setShowImportExportModal(true); }}
              className="text-xs font-semibold text-blue-600 hover:text-blue-800 flex items-center gap-1 cursor-pointer"
            >
              <span>View import history</span>
              <ChevronRight size={13} />
            </span>
          </div>
        </div>
      </div>

      {/* IMPORT / EXPORT MODAL (Image 5) */}
      {showImportExportModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-2xl w-full overflow-hidden border border-slate-200 animate-in fade-in zoom-in duration-150">
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <button
                  type="button"
                  onClick={() => setImportExportTab('import')}
                  className={`text-sm font-bold pb-1 border-b-2 transition-all ${
                    importExportTab === 'import' ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-400 hover:text-slate-600'
                  }`}
                >
                  Import Data
                </button>
                <button
                  type="button"
                  onClick={() => setImportExportTab('export')}
                  className={`text-sm font-bold pb-1 border-b-2 transition-all ${
                    importExportTab === 'export' ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-400 hover:text-slate-600'
                  }`}
                >
                  Export Data
                </button>
              </div>
              <button
                type="button"
                onClick={() => setShowImportExportModal(false)}
                className="text-slate-400 hover:text-slate-600 p-1"
              >
                <X size={18} />
              </button>
            </div>

            <div className="p-6">
              {importExportTab === 'import' ? (
                <div className="space-y-5">
                  <div className="border-2 border-dashed border-slate-300 rounded-xl p-8 flex flex-col items-center justify-center text-center hover:border-blue-500 transition-colors bg-slate-50/50 cursor-pointer">
                    <div className="w-12 h-12 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center mb-3">
                      <Upload size={22} />
                    </div>
                    <h4 className="text-sm font-bold text-slate-800">Drag and drop file here</h4>
                    <p className="text-xs text-slate-500 mt-1">or</p>
                    <label className="mt-2.5 px-4 py-1.5 bg-blue-600 hover:bg-blue-700 text-white font-semibold text-xs rounded-lg cursor-pointer shadow-xs">
                      Choose File
                      <input type="file" accept=".csv,.xlsx" className="hidden" />
                    </label>
                    <p className="text-[11px] text-slate-400 mt-3">
                      Supported formats: .csv, .xlsx | Max file size: 500 MB
                    </p>
                  </div>

                  <div>
                    <h5 className="text-xs font-bold text-slate-700 uppercase tracking-wider mb-2">Recent imports</h5>
                    <div className="border border-slate-200 rounded-lg overflow-hidden text-xs">
                      <div className="bg-slate-100 px-3 py-2 font-semibold text-slate-600 flex justify-between">
                        <span>File name</span>
                        <span>Status</span>
                      </div>
                      <div className="divide-y divide-slate-100">
                        <div className="px-3 py-2 flex items-center justify-between">
                          <div>
                            <span className="font-semibold text-slate-800">DGA_2025_Q1.xlsx</span>
                            <div className="text-[10px] text-slate-400">15 May 2025 10:32 • 12,845 records</div>
                          </div>
                          <span className="text-[11px] font-bold text-emerald-600 bg-emerald-50 px-2 py-0.5 rounded flex items-center gap-1">
                            <CheckCircle2 size={12} /> Completed
                          </span>
                        </div>
                        <div className="px-3 py-2 flex items-center justify-between">
                          <div>
                            <span className="font-semibold text-slate-800">PD_Monitoring_Mar.xlsx</span>
                            <div className="text-[10px] text-slate-400">02 May 2025 14:18 • 8,230 records</div>
                          </div>
                          <span className="text-[11px] font-bold text-emerald-600 bg-emerald-50 px-2 py-0.5 rounded flex items-center gap-1">
                            <CheckCircle2 size={12} /> Completed
                          </span>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="space-y-4 text-xs">
                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">Asset / Transformer</label>
                    <select className="w-full p-2 border border-slate-300 rounded-lg outline-none font-medium">
                      <option>{equipmentCode} — {equipmentName}</option>
                      <option>TR-101 — 230 / 132 / 33 kV</option>
                    </select>
                  </div>

                  <div>
                    <label className="block font-semibold text-slate-700 mb-1">Data categories</label>
                    <div className="flex flex-wrap gap-2">
                      <span className="px-2.5 py-1 bg-blue-50 text-blue-700 font-semibold rounded-md border border-blue-200">DGA ✓</span>
                      <span className="px-2.5 py-1 bg-blue-50 text-blue-700 font-semibold rounded-md border border-blue-200">Partial Discharge ✓</span>
                      <span className="px-2.5 py-1 bg-blue-50 text-blue-700 font-semibold rounded-md border border-blue-200">Oil Quality ✓</span>
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">Time range</label>
                      <input type="text" defaultValue="01 Apr 2024 — 30 Apr 2025" className="w-full p-2 border border-slate-300 rounded-lg" />
                    </div>
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">File format</label>
                      <select className="w-full p-2 border border-slate-300 rounded-lg font-medium">
                        <option>Microsoft Excel (.xlsx)</option>
                        <option>Comma-Separated Values (.csv)</option>
                      </select>
                    </div>
                  </div>

                  <div className="pt-4 flex justify-end gap-2 border-t border-slate-100">
                    <button
                      type="button"
                      onClick={() => setShowImportExportModal(false)}
                      className="px-4 py-2 border border-slate-300 rounded-lg text-slate-700 font-semibold"
                    >
                      Close
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        // Generate simple CSV download
                        const rows = [
                          ['Date', 'H2', 'CH4', 'C2H6', 'C2H4', 'C2H2', 'CO', 'CO2'],
                          ...chartTimeline.map(t => [t.fullDate, t.h2, t.ch4, t.c2h6, t.c2h4, t.c2h2, t.co, t.co2])
                        ];
                        const csvContent = 'data:text/csv;charset=utf-8,' + rows.map(e => e.join(',')).join('\n');
                        const encodedUri = encodeURI(csvContent);
                        const link = document.createElement('a');
                        link.setAttribute('href', encodedUri);
                        link.setAttribute('download', `DGA_Export_${equipmentCode}.csv`);
                        document.body.appendChild(link);
                        link.click();
                        document.body.removeChild(link);
                        setShowImportExportModal(false);
                      }}
                      className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-lg shadow-sm"
                    >
                      Export Data
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ANNOTATIONS MODAL */}
      {showAnnotationsModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-lg w-full overflow-hidden border border-slate-200">
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-900 text-white">
              <h3 className="font-bold text-sm flex items-center gap-2">
                <TrendingUp size={16} className="text-blue-400" />
                Manage Trend Chart Annotations
              </h3>
              <button
                type="button"
                onClick={() => setShowAnnotationsModal(false)}
                className="text-slate-400 hover:text-white"
              >
                <X size={18} />
              </button>
            </div>

            <div className="p-6 space-y-4 text-xs">
              <form onSubmit={handleAddAnnotation} className="space-y-2">
                <label className="block font-semibold text-slate-700">Add new engineering event / note</label>
                <div className="flex gap-2">
                  <input
                    type="text"
                    placeholder="e.g., Oil degasification, Tap changer inspection..."
                    value={newAnnotationText}
                    onChange={e => setNewAnnotationText(e.target.value)}
                    className="flex-1 px-3 py-2 border border-slate-300 rounded-lg outline-none focus:ring-2 focus:ring-blue-500"
                  />
                  <button
                    type="submit"
                    className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-lg whitespace-nowrap shadow-xs"
                  >
                    Add
                  </button>
                </div>
              </form>

              <div className="pt-2">
                <span className="block font-bold text-slate-500 uppercase tracking-wider text-[10px] mb-2">
                  Recorded Annotations ({annotations.length})
                </span>
                <div className="space-y-2 max-h-60 overflow-y-auto pr-1">
                  {annotations.map(item => (
                    <div key={item.id} className="p-2.5 rounded-lg bg-slate-50 border border-slate-200 flex items-center justify-between">
                      <div>
                        <div className="font-semibold text-slate-800">{item.title}</div>
                        <div className="text-[10px] text-slate-400 mt-0.5">{item.date}</div>
                      </div>
                      <button
                        type="button"
                        onClick={() => setAnnotations(prev => prev.filter(a => a.id !== item.id))}
                        className="text-slate-400 hover:text-rose-600 p-1"
                      >
                        <X size={14} />
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            <div className="px-6 py-3 bg-slate-50 border-t border-slate-200 flex justify-end">
              <button
                type="button"
                onClick={() => setShowAnnotationsModal(false)}
                className="px-4 py-1.5 bg-slate-900 text-white text-xs font-bold rounded-lg"
              >
                Done
              </button>
            </div>
          </div>
        </div>
      )}

      {/* DGA Details Modal */}
      <DgaDetailsModal
        isOpen={showDgaDetailsModal}
        onClose={() => setShowDgaDetailsModal(false)}
        currentGases={currentFormValues}
        currentFormValues={currentFormValues}
        analysisResult={activeAnalysis}
        equipmentCode={assetCode}
        equipmentName={assetName}
        sampleDate={entryDate}
        engineerName={entryEngineer}
        onUpdateFormValues={onUpdateFormValues}
      />

      {/* Ratio Settings Modal */}
      <RatioSettingsModal
        isOpen={showRatioSettingsModal}
        onClose={() => setShowRatioSettingsModal(false)}
      />

      {/* Risk Assessment Modal */}
      <RiskAssessmentModal
        isOpen={showRiskAssessmentModal}
        onClose={() => setShowRiskAssessmentModal(false)}
        riskScore={riskScore}
        currentGases={currentFormValues}
        currentFormValues={currentFormValues}
        analysisResult={activeAnalysis}
        equipmentCode={assetCode}
        equipmentName={assetName}
        sampleDate={entryDate}
        engineerName={entryEngineer}
      />

      {/* Ratio Analysis Modal */}
      <RatioAnalysisModal
        isOpen={showRatioAnalysisModal}
        onClose={() => setShowRatioAnalysisModal(false)}
        currentGases={currentFormValues}
        currentFormValues={currentFormValues}
        analysisResult={activeAnalysis}
        equipmentCode={assetCode}
        equipmentName={assetName}
      />

      {/* Sample History Modal */}
      <SampleHistoryModal
        isOpen={showSampleHistoryModal}
        onClose={() => setShowSampleHistoryModal(false)}
        equipmentCode={assetCode}
        equipmentName={assetName}
        onLoadSample={handleLoadSampleFromHistory}
      />

      {/* Transformer Overview Report Modal */}
      <TransformerOverviewReportModal
        isOpen={showReportModal}
        onClose={() => setShowReportModal(false)}
        equipmentCode={assetCode}
        equipmentName={assetName}
        customerName={customerName}
        sampleDate={entryDate}
        engineerName={entryEngineer}
        currentGases={currentFormValues}
        currentFormValues={currentFormValues}
        analysisResult={activeAnalysis}
        riskScore={riskScore}
        onDownloadPdf={onExportPdf}
      />

      {/* Event Log Modal */}
      <EventLogModal
        isOpen={showEventLogModal}
        onClose={() => setShowEventLogModal(false)}
        equipmentCode={assetCode}
        equipmentName={assetName}
      />
    </div>
  );
};
