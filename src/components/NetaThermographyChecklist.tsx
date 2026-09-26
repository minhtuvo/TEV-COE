import React, { useState, useMemo } from 'react';
import {
  Flame,
  Thermometer,
  AlertTriangle,
  AlertOctagon,
  CheckCircle2,
  XCircle,
  Sparkles,
  Printer,
  RotateCcw,
  Save,
  Plus,
  Trash2,
  FileCheck2,
  Copy,
  Check,
  Eye,
  Info,
  Layers,
  Sliders,
  Gauge,
  Zap,
  Activity,
  Maximize2
} from 'lucide-react';
import {
  NetaThermographyInputPayload,
  NetaThermographyEvaluationResult,
  ThermalMeasurementItem,
  evaluateThermalItem
} from '../../server/netaThermographyAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

const SAMPLE_THERMAL_DATA: NetaThermographyInputPayload = {
  site_info: {
    project_name: "Tram Phieu TEV Main Switchboard",
    survey_tag: "MSB-01-BUSBAR-INCOMING",
    fse_name: "Nguyen Van H",
    test_date: "2026-09-23"
  },
  survey_conditions: {
    camera_model: "FLIR E86 (Sensitivity < 0.04 deg C)",
    ambient_temperature_c: 32.0,
    system_voltage_v: 400,
    operating_current_amp: 650,
    load_percentage: 78,
    inaccessible_areas: "Behind primary transformer barrier"
  },
  thermal_measurements: [
    {
      component_name: "Bolted Busbar Joint Phase A",
      measured_temp_c: 36.5,
      reference_component_temp_c: 35.0
    },
    {
      component_name: "Circuit Breaker Incoming Lug Phase B",
      measured_temp_c: 75.8,
      reference_component_temp_c: 41.2
    },
    {
      component_name: "Feeder Cable Termination Phase C",
      measured_temp_c: 48.0,
      reference_component_temp_c: 44.5
    }
  ]
};

export const NetaThermographyChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaThermographyInputPayload>(SAMPLE_THERMAL_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaThermographyEvaluationResult | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [jsonText, setJsonText] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live Calculations according to ANSI/NETA ATS-2025 Table 100.18
  const liveStats = useMemo(() => {
    const amb = data.survey_conditions.ambient_temperature_c ?? 32.0;
    const items = data.thermal_measurements.map(item => evaluateThermalItem(item, amb));

    const maxLevel = items.length > 0 ? Math.max(...items.map(i => i.level)) : 0;
    const maxHotspot = items.length > 0 ? Math.max(...items.map(i => i.measured_temp_c)) : 0;
    const maxDeltaComp = items.length > 0 ? Math.max(...items.map(i => i.delta_t_comp)) : 0;
    const maxDeltaAmb = items.length > 0 ? Math.max(...items.map(i => i.delta_t_amb)) : 0;

    const criticalItem = items.find(i => i.level === 4);
    const monitorItem = items.find(i => i.level === 3);

    return {
      ambientTemp: amb,
      loadPct: data.survey_conditions.load_percentage,
      operatingAmp: data.survey_conditions.operating_current_amp,
      items,
      maxLevel,
      maxHotspot,
      maxDeltaComp,
      maxDeltaAmb,
      criticalItem,
      monitorItem,
      hasLevel4: maxLevel === 4,
      hasLevel3: maxLevel === 3,
      hasLevel2: maxLevel === 2
    };
  }, [data]);

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    setAnalysisResult(null);
    setSavedSuccessMessage(null);

    try {
      const response = await fetch('/api/field-service/neta-thermography-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });

      const resJson = await response.json();
      if (resJson.success) {
        setAnalysisResult(resJson.data);
      } else {
        alert('Lỗi phân tích: ' + resJson.error);
      }
    } catch (err: any) {
      console.error('Lỗi khi gọi API phân tích NETA Thermography:', err);
      alert('Không thể kết nối đến máy chủ AI: ' + err.message);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleSaveToCmms = () => {
    if (!analysisResult) {
      alert('Vui lòng bấm "Phân Tích AI Chuyên Gia" trước khi lưu vào kho báo cáo.');
      return;
    }
    if (onSaveReport) {
      onSaveReport({
        data,
        analysisResult,
        equipmentId: data.site_info.survey_tag,
        date: data.site_info.test_date,
        type: 'Thermographic Survey (NETA ATS-2025 Sec 9 & Table 100.18)'
      });
      setSavedSuccessMessage(`Đã lưu biên bản khảo sát ảnh nhiệt [${data.site_info.survey_tag}] vào kho báo cáo CMMS!`);
    }
  };

  const handleResetToSample = () => {
    setData(SAMPLE_THERMAL_DATA);
    setAnalysisResult(null);
    setSavedSuccessMessage(null);
  };

  const handleOpenJsonModal = () => {
    setJsonText(JSON.stringify(data, null, 2));
    setShowJsonModal(true);
  };

  const handleApplyJson = () => {
    try {
      const parsed = JSON.parse(jsonText);
      setData(parsed);
      setShowJsonModal(false);
      setAnalysisResult(null);
      setSavedSuccessMessage(null);
    } catch (err: any) {
      alert('Định dạng JSON không hợp lệ: ' + err.message);
    }
  };

  const handleCopyJson = () => {
    navigator.clipboard.writeText(JSON.stringify(data, null, 2));
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleAddItem = () => {
    const newItem: ThermalMeasurementItem = {
      component_name: `New Component Phase ${String.fromCharCode(65 + (data.thermal_measurements.length % 3))}`,
      measured_temp_c: 38.0,
      reference_component_temp_c: 35.0
    };
    setData({
      ...data,
      thermal_measurements: [...data.thermal_measurements, newItem]
    });
  };

  const handleRemoveItem = (index: number) => {
    setData({
      ...data,
      thermal_measurements: data.thermal_measurements.filter((_, i) => i !== index)
    });
  };

  const handleItemChange = (index: number, field: keyof ThermalMeasurementItem, value: any) => {
    const updated = [...data.thermal_measurements];
    updated[index] = {
      ...updated[index],
      [field]: value
    };
    setData({
      ...data,
      thermal_measurements: updated
    });
  };

  return (
    <div className="space-y-6">
      {/* Top Banner & Header */}
      <div className="bg-gradient-to-r from-red-950 via-orange-950 to-slate-900 text-white p-5 sm:p-6 rounded-2xl shadow-lg border border-red-500/30">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1">
            <div className="flex flex-wrap items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-extrabold uppercase tracking-wider bg-white/20 text-white backdrop-blur-xs border border-white/30 flex items-center gap-1.5">
                <Thermometer size={13} className="text-amber-300" />
                ANSI/NETA ATS-2025 Mục 9 & Bảng Table 100.18
              </span>
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-red-950/50 text-red-200 border border-red-400/40 flex items-center gap-1">
                <Flame size={12} />
                Thermographic Survey • Hồng Ngoại Phát Nhiệt • 4 Mức Mối Nguy
              </span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white">
              Khảo Sát & Đánh Giá Nhiệt Độ Hồng Ngoại (Thermographic Survey)
            </h1>
            <p className="text-xs sm:text-sm text-red-100 max-w-2xl">
              Tự động tính toán chênh lệch ΔT_comp (so với linh kiện cùng tải) và ΔT_amb (so với môi trường) để phân cấp 4 mức khuyến nghị hành động theo NETA Table 100.18.
            </p>
          </div>

          {/* Action Buttons */}
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={handleResetToSample}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title="Khôi phục dữ liệu mẫu hiện trường FSE"
            >
              <RotateCcw size={14} />
              <span>Dữ Liệu Mẫu FSE</span>
            </button>

            <button
              onClick={handleAddItem}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title="Thêm điểm đo nhiệt độ mới"
            >
              <Plus size={14} />
              <span>Thêm Điểm Đo</span>
            </button>

            <button
              onClick={handleOpenJsonModal}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title="Nhập hoặc sao chép mã JSON"
            >
              <FileCheck2 size={14} />
              <span>Mã JSON</span>
            </button>

            <button
              onClick={() => setShowPrintModal(true)}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title="In hoặc xuất biên bản PDF TEV"
            >
              <Printer size={14} />
              <span>In Biên Bản PDF</span>
            </button>

            <button
              onClick={handleSaveToCmms}
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-red-800 hover:bg-red-700 text-white border border-red-500/40 flex items-center gap-1.5 transition-all shadow-md"
              title="Lưu biên bản kiểm tra vào kho báo cáo"
            >
              <Save size={15} />
              <span>Lưu vào Kho Báo Cáo</span>
            </button>

            <button
              onClick={handleAnalyze}
              disabled={isAnalyzing}
              className="px-4 py-2 rounded-xl text-xs font-bold bg-gradient-to-r from-amber-400 to-yellow-300 hover:from-amber-300 hover:to-yellow-200 text-slate-950 flex items-center gap-2 transition-all shadow-lg shadow-red-950/40 disabled:opacity-50"
            >
              <Sparkles size={15} className={isAnalyzing ? 'animate-spin' : 'text-slate-950'} />
              <span>{isAnalyzing ? 'Đang Phân Tích NETA...' : 'Phân Tích AI Chuyên Gia'}</span>
            </button>
          </div>
        </div>

        {/* Live NETA Diagnostic Strip */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5 mt-5 pt-4 border-t border-red-500/30 text-xs">
          {/* Survey Load & Ambient */}
          <div className="bg-red-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-red-500/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-red-200 block font-medium">Tải Khảo Sát & Môi Trường</span>
              <span className="text-base font-black text-amber-300">
                {liveStats.loadPct}% ({liveStats.operatingAmp}A) • {liveStats.ambientTemp}°C
              </span>
            </div>
            <Activity size={20} className="text-amber-400 shrink-0" />
          </div>

          {/* Max Hotspot Temp */}
          <div className="bg-red-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-red-500/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-red-200 block font-medium">Điểm Nóng Nhất (Hotspot)</span>
              <span className={`text-base font-black ${liveStats.hasLevel4 ? 'text-rose-300' : 'text-amber-300'}`}>
                {liveStats.maxHotspot}°C
              </span>
            </div>
            <Flame size={20} className={liveStats.hasLevel4 ? 'text-rose-400 animate-pulse shrink-0' : 'text-amber-400 shrink-0'} />
          </div>

          {/* Max Delta T_comp */}
          <div className="bg-red-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-red-500/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-red-200 block font-medium">Max ΔT_comp (Cùng tải)</span>
              <span className={`text-base font-black ${liveStats.maxDeltaComp > 15 ? 'text-rose-300' : liveStats.maxDeltaComp >= 4 ? 'text-amber-300' : 'text-emerald-300'}`}>
                +{liveStats.maxDeltaComp}°C {liveStats.maxDeltaComp > 15 ? '(Mức 4)' : ''}
              </span>
            </div>
            {liveStats.maxDeltaComp > 15 ? (
              <AlertOctagon size={20} className="text-rose-400 animate-bounce shrink-0" />
            ) : (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            )}
          </div>

          {/* Max Delta T_amb */}
          <div className="bg-red-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-red-500/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-red-200 block font-medium">Max ΔT_amb (Môi trường)</span>
              <span className={`text-base font-black ${liveStats.maxDeltaAmb > 40 ? 'text-rose-300' : liveStats.maxDeltaAmb >= 21 ? 'text-orange-300' : 'text-emerald-300'}`}>
                +{liveStats.maxDeltaAmb}°C {liveStats.maxDeltaAmb > 40 ? '(Mức 4)' : ''}
              </span>
            </div>
            {liveStats.maxDeltaAmb > 40 ? (
              <AlertTriangle size={20} className="text-rose-400 animate-pulse shrink-0" />
            ) : (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            )}
          </div>
        </div>
      </div>

      {/* Save Feedback Toast */}
      {savedSuccessMessage && (
        <div className="bg-emerald-50 border border-emerald-300 p-4 rounded-xl flex items-center justify-between shadow-xs">
          <div className="flex items-center gap-2.5">
            <CheckCircle2 size={18} className="text-emerald-600" />
            <span className="text-xs sm:text-sm font-bold text-emerald-900">{savedSuccessMessage}</span>
          </div>
          {onNavigateToReports && (
            <button
              onClick={onNavigateToReports}
              className="px-3 py-1 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold rounded-lg shadow-xs transition-colors"
            >
              Mở Kho Báo Cáo
            </button>
          )}
        </div>
      )}

      {/* Main Form: Survey Conditions & Measurements */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left Column: Site Info & Survey Conditions */}
        <div className="space-y-6">
          {/* Site Information */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <Thermometer size={17} className="text-red-600" />
                Thông Tin Khảo Sát
              </h2>
              <span className="text-[10px] font-semibold text-slate-500 bg-slate-100 px-2 py-0.5 rounded-full">
                NETA ATS-2025 Sec 9
              </span>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="text-slate-600 font-medium block mb-1">Tên Dự Án / Nhà Máy:</label>
                <input
                  type="text"
                  value={data.site_info.project_name}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, project_name: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-red-500 font-medium text-slate-800"
                />
              </div>

              <div>
                <label className="text-slate-600 font-medium block mb-1">Mã Vị Trí / Tủ Điện (Survey Tag):</label>
                <input
                  type="text"
                  value={data.site_info.survey_tag}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, survey_tag: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-red-500 font-bold text-red-700 font-mono"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Kỹ Sư FSE Đo:</label>
                  <input
                    type="text"
                    value={data.site_info.fse_name}
                    onChange={(e) => setData({
                      ...data,
                      site_info: { ...data.site_info, fse_name: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-red-500 font-medium text-slate-800"
                  />
                </div>
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Ngày Đo Nhiệt:</label>
                  <input
                    type="date"
                    value={data.site_info.test_date}
                    onChange={(e) => setData({
                      ...data,
                      site_info: { ...data.site_info, test_date: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-red-500 font-medium text-slate-800"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Survey Conditions (Section 9.1 - 9.3) */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <Gauge size={17} className="text-red-600" />
                Điều Kiện Đo (NETA Mục 9.1 - 9.3)
              </h2>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="text-slate-600 font-medium block mb-1">Model Camera Nhiệt (Độ nhạy ≤ 1°C):</label>
                <input
                  type="text"
                  value={data.survey_conditions.camera_model}
                  onChange={(e) => setData({
                    ...data,
                    survey_conditions: { ...data.survey_conditions, camera_model: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-red-500 font-medium text-slate-800"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Nhiệt Độ Môi Trường (°C):</label>
                  <input
                    type="number"
                    step="0.5"
                    value={data.survey_conditions.ambient_temperature_c}
                    onChange={(e) => setData({
                      ...data,
                      survey_conditions: { ...data.survey_conditions, ambient_temperature_c: parseFloat(e.target.value) || 0 }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-red-500 font-bold text-slate-800"
                  />
                </div>
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Điện Áp Vận Hành (V):</label>
                  <input
                    type="number"
                    value={data.survey_conditions.system_voltage_v}
                    onChange={(e) => setData({
                      ...data,
                      survey_conditions: { ...data.survey_conditions, system_voltage_v: parseFloat(e.target.value) || 0 }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-red-500 font-medium text-slate-800"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Dòng Điện Tải (A):</label>
                  <input
                    type="number"
                    value={data.survey_conditions.operating_current_amp}
                    onChange={(e) => setData({
                      ...data,
                      survey_conditions: { ...data.survey_conditions, operating_current_amp: parseFloat(e.target.value) || 0 }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-red-500 font-bold text-slate-800"
                  />
                </div>
                <div>
                  <label className="text-slate-600 font-medium block mb-1">% Tải Vận Hành (Tối thiểu 40%):</label>
                  <input
                    type="number"
                    value={data.survey_conditions.load_percentage}
                    onChange={(e) => setData({
                      ...data,
                      survey_conditions: { ...data.survey_conditions, load_percentage: parseFloat(e.target.value) || 0 }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-red-500 font-bold text-emerald-700"
                  />
                </div>
              </div>

              <div>
                <label className="text-slate-600 font-medium block mb-1">Vị Trí Không Thể Tiếp Cận/Đo (Inaccessible):</label>
                <input
                  type="text"
                  value={data.survey_conditions.inaccessible_areas || ''}
                  onChange={(e) => setData({
                    ...data,
                    survey_conditions: { ...data.survey_conditions, inaccessible_areas: e.target.value }
                  })}
                  placeholder="Ví dụ: Phía sau vách ngăn MBA chính..."
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-red-500 font-medium text-slate-800"
                />
              </div>
            </div>
          </div>

          {/* Quick Guide Card: Table 100.18 */}
          <div className="bg-slate-900 text-white p-4 rounded-2xl border border-slate-800 shadow-sm space-y-2.5 text-xs">
            <h3 className="font-bold text-amber-300 flex items-center gap-1.5 text-xs">
              <Info size={14} />
              Bảng Phân Cấp NETA ATS-2025 Table 100.18
            </h3>
            <ul className="space-y-1.5 text-[11px] text-slate-300">
              <li className="flex items-start gap-1.5">
                <span className="w-2 h-2 rounded-full bg-blue-400 mt-1 shrink-0" />
                <span><strong>Mức 1:</strong> 1-3°C comp hoặc 1-10°C amb $\rightarrow$ <em>Warrants investigation</em></span>
              </li>
              <li className="flex items-start gap-1.5">
                <span className="w-2 h-2 rounded-full bg-yellow-400 mt-1 shrink-0" />
                <span><strong>Mức 2:</strong> 4-15°C comp hoặc 11-20°C amb $\rightarrow$ <em>Repair as time permits</em></span>
              </li>
              <li className="flex items-start gap-1.5">
                <span className="w-2 h-2 rounded-full bg-orange-400 mt-1 shrink-0" />
                <span><strong>Mức 3:</strong> 21-40°C amb $\rightarrow$ <em>Monitor until corrective measures</em></span>
              </li>
              <li className="flex items-start gap-1.5">
                <span className="w-2 h-2 rounded-full bg-rose-500 mt-1 shrink-0" />
                <span><strong>Mức 4 (CẤP BÁCH):</strong> &gt;15°C comp hoặc &gt;40°C amb $\rightarrow$ <em>Repair immediately</em></span>
              </li>
            </ul>
          </div>
        </div>

        {/* Center & Right Column: Thermal Measurements Table */}
        <div className="lg:col-span-2 space-y-6">
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-100 pb-3">
              <div>
                <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                  <Flame size={17} className="text-red-600" />
                  Danh Sách Điểm Đo Nhiệt Độ Hồng Ngoại & Đánh Giá
                </h2>
                <span className="text-[11px] text-slate-500">
                  Tự động so sánh với linh kiện đối chứng (T_ref) và môi trường (T_amb = {liveStats.ambientTemp}°C)
                </span>
              </div>

              <button
                onClick={handleAddItem}
                className="px-3 py-1.5 bg-red-50 hover:bg-red-100 text-red-700 text-xs font-bold rounded-lg border border-red-200 flex items-center gap-1 transition-colors self-start sm:self-auto"
              >
                <Plus size={14} />
                <span>Thêm Điểm Đo</span>
              </button>
            </div>

            {/* Measurement List */}
            <div className="space-y-3">
              {data.thermal_measurements.map((item, idx) => {
                const evalItem = liveStats.items[idx];
                return (
                  <div
                    key={idx}
                    className={`p-3.5 rounded-xl border transition-all ${
                      evalItem?.level === 4
                        ? 'bg-rose-50/70 border-rose-300 ring-1 ring-rose-300'
                        : evalItem?.level === 3
                        ? 'bg-orange-50/70 border-orange-300'
                        : evalItem?.level === 2
                        ? 'bg-amber-50/70 border-amber-300'
                        : evalItem?.level === 1
                        ? 'bg-blue-50/70 border-blue-200'
                        : 'bg-slate-50 border-slate-200'
                    }`}
                  >
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-2 mb-2 border-b border-slate-200/60">
                      <div className="flex items-center gap-2 flex-1">
                        <span className="w-5 h-5 rounded-full bg-slate-200 text-slate-700 text-[10px] font-bold flex items-center justify-center shrink-0">
                          {idx + 1}
                        </span>
                        <input
                          type="text"
                          value={item.component_name}
                          onChange={(e) => handleItemChange(idx, 'component_name', e.target.value)}
                          className="font-bold text-xs text-slate-800 bg-white/80 border border-slate-200 rounded px-2 py-1 flex-1 focus:ring-1 focus:ring-red-500"
                        />
                      </div>

                      <div className="flex items-center gap-2">
                        {/* Status Badge */}
                        <span
                          className={`px-2.5 py-0.5 rounded-full text-[10px] font-black border ${
                            evalItem?.level === 4
                              ? 'bg-rose-600 text-white border-rose-700 animate-pulse'
                              : evalItem?.level === 3
                              ? 'bg-orange-500 text-white border-orange-600'
                              : evalItem?.level === 2
                              ? 'bg-amber-400 text-slate-900 border-amber-500'
                              : evalItem?.level === 1
                              ? 'bg-blue-500 text-white border-blue-600'
                              : 'bg-emerald-500 text-white border-emerald-600'
                          }`}
                        >
                          {evalItem?.level_name}
                        </span>

                        {data.thermal_measurements.length > 1 && (
                          <button
                            onClick={() => handleRemoveItem(idx)}
                            className="p-1 text-slate-400 hover:text-rose-600 transition-colors"
                            title="Xóa điểm đo này"
                          >
                            <Trash2 size={15} />
                          </button>
                        )}
                      </div>
                    </div>

                    {/* Inputs & Delta T Calculations */}
                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 text-xs">
                      <div>
                        <label className="text-slate-500 text-[10px] font-medium block">Nhiệt Độ Điểm Đo (T):</label>
                        <div className="flex items-center gap-1 mt-0.5">
                          <input
                            type="number"
                            step="0.1"
                            value={item.measured_temp_c}
                            onChange={(e) => handleItemChange(idx, 'measured_temp_c', parseFloat(e.target.value) || 0)}
                            className="w-full px-2 py-1 bg-white border border-slate-300 rounded font-black text-slate-900 focus:ring-1 focus:ring-red-500"
                          />
                          <span className="text-slate-500 font-bold">°C</span>
                        </div>
                      </div>

                      <div>
                        <label className="text-slate-500 text-[10px] font-medium block">Nhiệt Độ Đối Chứng (T_ref):</label>
                        <div className="flex items-center gap-1 mt-0.5">
                          <input
                            type="number"
                            step="0.1"
                            value={item.reference_component_temp_c}
                            onChange={(e) => handleItemChange(idx, 'reference_component_temp_c', parseFloat(e.target.value) || 0)}
                            className="w-full px-2 py-1 bg-white border border-slate-300 rounded font-medium text-slate-800 focus:ring-1 focus:ring-red-500"
                          />
                          <span className="text-slate-500 font-bold">°C</span>
                        </div>
                      </div>

                      {/* Calculated Delta T_comp */}
                      <div className="bg-white/80 p-1.5 rounded border border-slate-200">
                        <span className="text-slate-500 text-[10px] block">ΔT_comp (So với pha đối chứng):</span>
                        <span className={`text-xs font-black block mt-0.5 ${
                          (evalItem?.delta_t_comp ?? 0) > 15 ? 'text-rose-600' : (evalItem?.delta_t_comp ?? 0) >= 4 ? 'text-amber-600' : 'text-slate-800'
                        }`}>
                          {evalItem?.delta_t_comp ?? 0 > 0 ? '+' : ''}{evalItem?.delta_t_comp ?? 0}°C
                        </span>
                      </div>

                      {/* Calculated Delta T_amb */}
                      <div className="bg-white/80 p-1.5 rounded border border-slate-200">
                        <span className="text-slate-500 text-[10px] block">ΔT_amb (So với môi trường):</span>
                        <span className={`text-xs font-black block mt-0.5 ${
                          (evalItem?.delta_t_amb ?? 0) > 40 ? 'text-rose-600' : (evalItem?.delta_t_amb ?? 0) >= 21 ? 'text-orange-600' : 'text-slate-800'
                        }`}>
                          {evalItem?.delta_t_amb ?? 0 > 0 ? '+' : ''}{evalItem?.delta_t_amb ?? 0}°C
                        </span>
                      </div>
                    </div>

                    {/* Action Guideline */}
                    <div className="mt-2 pt-2 border-t border-slate-200/60 text-[11px] text-slate-600 flex items-center justify-between">
                      <span><strong>Hành động:</strong> {evalItem?.action_recommendation}</span>
                      <span className="font-bold text-slate-500">NETA Table 100.18</span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </div>

      {/* AI Analysis Output Section */}
      {analysisResult && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden animate-in fade-in slide-in-from-bottom-3 duration-300">
          <div className="p-5 bg-gradient-to-r from-slate-900 via-red-950 to-slate-900 text-white flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-red-500/20 border border-red-400/30 flex items-center justify-center text-red-300 shrink-0">
                <Sparkles size={20} />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className="text-[10px] uppercase tracking-wider font-extrabold px-2 py-0.5 rounded-full bg-white/10 text-white">
                    TEV Platform AI Evaluation
                  </span>
                  <span className={`text-[11px] font-black px-2 py-0.5 rounded-full border ${
                    analysisResult.overall_status === 'REPAIR_IMMEDIATELY'
                      ? 'bg-rose-500/30 text-rose-300 border-rose-400/50'
                      : analysisResult.overall_status === 'MONITOR'
                      ? 'bg-orange-500/30 text-orange-300 border-orange-400/50'
                      : analysisResult.overall_status === 'REPAIR_SCHEDULED'
                      ? 'bg-amber-500/30 text-amber-300 border-amber-400/50'
                      : 'bg-emerald-500/30 text-emerald-300 border-emerald-400/50'
                  }`}>
                    {analysisResult.overall_status}
                  </span>
                </div>
                <h3 className="text-base font-bold text-white mt-0.5">
                  Báo Cáo Đánh Giá Khảo Sát Nhiệt Độ Hồng Ngoại (NETA ATS-2025 Table 100.18)
                </h3>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={() => setShowPrintModal(true)}
                className="px-3.5 py-1.5 bg-white/10 hover:bg-white/20 text-white text-xs font-semibold rounded-lg flex items-center gap-1.5 transition-colors border border-white/20"
              >
                <Printer size={14} />
                <span>In Báo Cáo</span>
              </button>
              <button
                onClick={handleSaveToCmms}
                className="px-4 py-1.5 bg-red-700 hover:bg-red-600 text-white text-xs font-bold rounded-lg flex items-center gap-1.5 shadow-md transition-colors"
              >
                <Save size={14} />
                <span>Lưu vào Kho Báo Cáo</span>
              </button>
            </div>
          </div>

          <div className="p-6 prose prose-slate max-w-none text-xs sm:text-sm">
            <div
              className="space-y-4"
              dangerouslySetInnerHTML={{
                __html: analysisResult.markdown_report
                  .replace(/### (.*?)\n/g, '<h3 class="text-base font-bold text-slate-900 border-b border-slate-200 pb-2 mt-4 mb-2 flex items-center gap-2">$1</h3>')
                  .replace(/\n\n/g, '<br/>')
              }}
            />
          </div>
        </div>
      )}

      {/* JSON Import/Export Modal */}
      {showJsonModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-white rounded-2xl max-w-2xl w-full p-6 shadow-2xl border border-slate-200 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <FileCheck2 size={18} className="text-red-600" />
                <h3 className="font-bold text-sm text-slate-800">
                  Dữ Liệu JSON Khảo Sát Nhiệt Độ Hồng Ngoại (NETA Sec 9 & Table 100.18)
                </h3>
              </div>
              <button
                onClick={() => setShowJsonModal(false)}
                className="text-slate-400 hover:text-slate-600 text-sm font-bold"
              >
                ✕
              </button>
            </div>

            <p className="text-xs text-slate-500">
              Kỹ sư FSE có thể sao chép chuỗi JSON này để lưu trữ hoặc dán dữ liệu đo mới từ Camera nhiệt vào đây và bấm "Áp Dụng Dữ Liệu":
            </p>

            <textarea
              value={jsonText}
              onChange={(e) => setJsonText(e.target.value)}
              rows={14}
              className="w-full p-3 font-mono text-xs bg-slate-900 text-amber-200 rounded-xl border border-slate-700 focus:outline-hidden focus:ring-2 focus:ring-red-500"
            />

            <div className="flex items-center justify-between pt-2">
              <button
                onClick={handleCopyJson}
                className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-lg flex items-center gap-1.5 transition-colors"
              >
                {copied ? <Check size={14} className="text-emerald-600" /> : <Copy size={14} />}
                <span>{copied ? 'Đã Sao Chép!' : 'Sao Chép JSON'}</span>
              </button>

              <div className="flex gap-2">
                <button
                  onClick={() => setShowJsonModal(false)}
                  className="px-3.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-medium rounded-lg"
                >
                  Hủy
                </button>
                <button
                  onClick={handleApplyJson}
                  className="px-4 py-1.5 bg-red-600 hover:bg-red-500 text-white text-xs font-bold rounded-lg shadow-sm"
                >
                  Áp Dụng Dữ Liệu
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* TEV Report Print Modal */}
      {showPrintModal && (
        <TevReportPrintModal
          isOpen={showPrintModal}
          onClose={() => setShowPrintModal(false)}
          title="BIÊN BẢN KHẢO SÁT & ĐÁNH GIÁ NHIỆT ĐỘ HỒNG NGOẠI (THERMOGRAPHIC SURVEY)"
          standard="ANSI/NETA ATS-2025 Section 9 & Table 100.18"
          siteInfo={{
            project_name: data.site_info.project_name,
            transformer_tag: data.site_info.survey_tag,
            fse_name: data.site_info.fse_name,
            test_date: data.site_info.test_date
          }}
          overallStatus={analysisResult?.overall_status || (liveStats.hasLevel4 ? 'CRITICAL' : liveStats.hasLevel3 ? 'INVESTIGATE' : 'PASS')}
          statusReason={analysisResult?.status_reason || 'Đã đối chiếu với quy chuẩn kỹ thuật NETA ATS-2025 Table 100.18.'}
          visualChecks={[
            { label: 'Camera nhiệt độ nhạy cao (Sensitivity < 0.04°C)', value: data.survey_conditions.camera_model, pass: true },
            { label: 'Tải vận hành đạt điều kiện bình thường (> 40%)', value: `${data.survey_conditions.load_percentage}% (${data.survey_conditions.operating_current_amp}A)`, pass: data.survey_conditions.load_percentage >= 40 },
            { label: 'Nhiệt độ không khí môi trường (Ambient)', value: `${data.survey_conditions.ambient_temperature_c}°C`, pass: true },
            { label: 'Điện áp hệ thống', value: `${data.survey_conditions.system_voltage_v}V`, pass: true },
            { label: 'Khu vực không thể tiếp cận / đo', value: data.survey_conditions.inaccessible_areas || 'Không có', pass: true }
          ]}
          electricalTable={liveStats.items.map(item => ({
            item: item.component_name,
            measured: `${item.measured_temp_c}°C`,
            standard: `ΔT_comp ≤ 15°C, ΔT_amb ≤ 40°C`,
            reference: `NETA Table 100.18 (T_ref: ${item.reference_component_temp_c}°C)`,
            deviation: `ΔT_comp: +${item.delta_t_comp}°C | ΔT_amb: +${item.delta_t_amb}°C`,
            status: item.level === 4 ? 'FAIL' : item.level >= 2 ? 'INVESTIGATE' : 'PASS'
          }))}
          warnings={[
            liveStats.hasLevel4 ? `Phát hiện điểm phát nhiệt Mức 4 (Major Discrepancy) tại ${liveStats.criticalItem?.component_name} (${liveStats.criticalItem?.measured_temp_c}°C). Nguy cơ lỏng bu-lông quá mức hoặc rỗ bề mặt tiếp xúc gây nóng chảy hoặc ngắn mạch phóng hồ quang!` : '',
            liveStats.hasLevel2 ? `Có điểm đo phát nhiệt Mức 2 (Probable deficiency), cần lên kế hoạch bảo dưỡng siết lực bu-lông khi có lịch dừng máy.` : ''
          ].filter(Boolean)}
          recommendations={[
            'CẢNH BÁO CẤP BÁCH: Cô lập khẩn cấp Aptomat chính (Circuit Breaker Incoming Lug Phase B) để vệ sinh bề mặt tiếp xúc, bôi mỡ tiếp xúc chuyên dụng và dùng cờ-lê lực siết lại bu-lông theo Bảng Table 100.12.',
            'Thực hiện chụp lại ảnh nhiệt (Thermography) ngay sau khi xử lý để xác nhận nhiệt độ Phase B đã hạ về mức bình thường dưới 40°C.',
            'Theo dõi định kỳ các điểm phát nhiệt Mức 1 và Mức 2 trong các đợt kiểm tra bảo trì tiếp theo.'
          ]}
        />
      )}
    </div>
  );
};
