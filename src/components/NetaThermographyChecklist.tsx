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
  Maximize2,
  ChevronDown,
  ChevronUp,
  Globe,
  ShieldCheck,
  Clock,
  ArrowRight
} from 'lucide-react';
import {
  NetaThermographyInputPayload,
  NetaThermographyEvaluationResult,
  ThermographicFinding,
  ThermalItemEvaluation,
  evaluateThermalItem,
  performDeterministicThermographyAnalysis
} from '../../server/netaThermographyAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';
import { CustomerSelectDropdown } from './CustomerSelectDropdown';
import { EquipmentSelectDropdown } from './EquipmentSelectDropdown';
import { FseEngineerDropdown } from './FseEngineerDropdown';
import { useAtsFormSync } from '../hooks/useAtsFormSync';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

// ----------------------------------------------------------------------------
// NETA MTS SECTION 9 & TABLE 100.18 BILINGUAL SAMPLE DATA
// ----------------------------------------------------------------------------
const SAMPLE_MTS_THERMAL_DATA: NetaThermographyInputPayload = {
  site_info: {
    project_name: "Tram Bien Ap TEV High-Tech Substation / TEV Substation",
    survey_tag: "IR-SURVEY-SWITCHGEAR-BUSDUCT-01",
    ambient_temperature_celsius: 32.5,
    circuit_loading_percent: 68.0,
    ir_camera_model: "FLIR T865 (Sensitivity < 0.03 deg C)",
    fse_name: "Nguyen Van F",
    test_date: "2026-10-16"
  },
  survey_parameters_section_9: {
    equipment_access_condition: "Pass (Scanned through certified IR viewports and opened back panels)",
    circuit_load_status: "Pass (68% load >= 40% NETA minimum loading requirement)",
    report_elements_complete: true
  },
  thermographic_findings: [
    {
      location: "Tủ Hạ thế MSB-01 / Main Switchboard - ACB Incomer Phase B Connection",
      component_type: "Bolted Busbar Joint / Mối nối bu-lông thanh cái",
      measured_temperature_celsius: 88.5,
      reference_component_temp_celsius: 42.0,
      similar_component_delta_t_celsius: 46.5,
      ambient_air_delta_t_celsius: 56.0,
      suspected_cause: "Loose bolted connection or improper torque / Lỏng bu-lông siết thanh cái"
    },
    {
      location: "Máy biến áp TR-02 / Transformer TR-02 LV Bushing Phase C Lug",
      component_type: "Cable Termination Lug / Đầu cốt cáp hạ thế",
      measured_temperature_celsius: 48.0,
      reference_component_temp_celsius: 41.5,
      similar_component_delta_t_celsius: 6.5,
      ambient_air_delta_t_celsius: 15.5,
      suspected_cause: "Slight contact resistance or uneven cable current / Điện trở tiếp xúc nhẹ"
    },
    {
      location: "Tủ Phân phối DB-FL01 / Feeder Breaker #04 Terminal",
      component_type: "MCCB Lug Terminal / Cọc đấu dây MCCB",
      measured_temperature_celsius: 61.0,
      reference_component_temp_celsius: 39.0,
      similar_component_delta_t_celsius: 22.0,
      ambient_air_delta_t_celsius: 28.5,
      suspected_cause: "Oxidized contact surface / Bề mặt tiếp điểm bị oxy hóa"
    }
  ],
  previous_test_data: {
    last_survey_date: "2025-10-15",
    last_max_delta_t_celsius: 3.5
  }
};

export const NetaThermographyChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useAtsFormSync<NetaThermographyInputPayload>('neta-thermography', SAMPLE_MTS_THERMAL_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaThermographyEvaluationResult | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showTable10018Modal, setShowTable10018Modal] = useState(false);
  const [showAiPromptModal, setShowAiPromptModal] = useState(false);
  const [copiedPrompt, setCopiedPrompt] = useState(false);
  const [jsonText, setJsonText] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);
  const [lang, setLang] = useState<'vi' | 'en'>('vi');

  const isVi = lang === 'vi';

  // Live Calculations according to ANSI/NETA MTS Section 9 & Table 100.18
  const liveStats = useMemo(() => {
    const amb = data.site_info.ambient_temperature_celsius ?? 32.5;
    const loadPct = data.site_info.circuit_loading_percent ?? 68.0;
    const findings = data.thermographic_findings || [];

    const items: ThermalItemEvaluation[] = findings.map(item => evaluateThermalItem(item, amb));

    const maxLevel = items.length > 0 ? Math.max(...items.map(i => i.level)) : 0;
    const maxHotspot = items.length > 0 ? Math.max(...items.map(i => i.measured_temp_c)) : 0;
    const maxDeltaComp = items.length > 0 ? Math.max(...items.map(i => i.delta_t_comp)) : 0;
    const maxDeltaAmb = items.length > 0 ? Math.max(...items.map(i => i.delta_t_amb)) : 0;

    const criticalItem = items.find(i => i.level === 4);
    const monitorItem = items.find(i => i.level === 3);

    const isLoadPass = loadPct >= 40.0;

    return {
      ambientTemp: amb,
      loadPct,
      isLoadPass,
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
        // Fallback to client-side deterministic evaluation
        const fallback = performDeterministicThermographyAnalysis(data);
        setAnalysisResult(fallback);
      }
    } catch (err: any) {
      console.warn('Fallback to client-side deterministic evaluation:', err);
      const fallback = performDeterministicThermographyAnalysis(data);
      setAnalysisResult(fallback);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleSaveToCmms = () => {
    if (!analysisResult) {
      alert(isVi ? 'Vui lòng bấm "Phân Tích AI NETA Table 100.18" trước khi lưu vào kho báo cáo.' : 'Please run AI Analysis before saving to report repository.');
      return;
    }
    if (onSaveReport) {
      onSaveReport({
        data,
        analysisResult,
        equipmentId: data.site_info.survey_tag,
        date: data.site_info.test_date,
        type: 'Thermographic Survey (ANSI/NETA MTS Section 9 & Table 100.18)'
      });
      setSavedSuccessMessage(
        isVi
          ? `Đã lưu biên bản khảo sát hồng ngoại [${data.site_info.survey_tag}] vào kho báo cáo CMMS!`
          : `Saved thermographic survey [${data.site_info.survey_tag}] to CMMS repository!`
      );
    }
  };

  const handleResetToSample = () => {
    setData(SAMPLE_MTS_THERMAL_DATA);
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
      alert((isVi ? 'Định dạng JSON không hợp lệ: ' : 'Invalid JSON format: ') + err.message);
    }
  };

  const handleCopyJson = () => {
    navigator.clipboard.writeText(JSON.stringify(data, null, 2));
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleAddFinding = () => {
    const count = (data.thermographic_findings || []).length;
    const newFinding: ThermographicFinding = {
      location: `Tủ Phân phối Lộ #${count + 1} / Feeder #${count + 1} Terminal`,
      component_type: "Bolted Busbar Joint / Mối nối bu-lông",
      measured_temperature_celsius: 42.0,
      reference_component_temp_celsius: 38.0,
      suspected_cause: "Kiểm tra định kỳ / Routine scan"
    };
    setData({
      ...data,
      thermographic_findings: [...(data.thermographic_findings || []), newFinding]
    });
  };

  const handleRemoveFinding = (index: number) => {
    setData({
      ...data,
      thermographic_findings: (data.thermographic_findings || []).filter((_, i) => i !== index)
    });
  };

  const handleFindingChange = (index: number, field: keyof ThermographicFinding, value: any) => {
    const updated = [...(data.thermographic_findings || [])];
    updated[index] = {
      ...updated[index],
      [field]: value
    };
    setData({
      ...data,
      thermographic_findings: updated
    });
  };

  return (
    <div className="space-y-6">
      {/* ---------------------------------------------------------------------- */}
      {/* 3-ROW OFFICIAL TEV HEADER BANNER */}
      {/* ---------------------------------------------------------------------- */}
      <div className="bg-gradient-to-r from-slate-950 via-slate-900 to-rose-950 text-white p-6 sm:p-7 rounded-2xl shadow-xl border border-rose-500/30 space-y-4">
        {/* Row 1: Brand & Standard Badges */}
        <div className="flex items-center justify-between gap-3 flex-wrap">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-rose-500 to-amber-500 text-slate-950 flex items-center justify-center font-black text-xl shadow-md shrink-0">
              TEV
            </div>
            <div className="flex items-center gap-2.5 flex-wrap">
              <span className="text-xs sm:text-sm font-black uppercase tracking-wider text-rose-300">
                TEV FLATFORM AI • FIELD SERVICE EVALUATION REPORT
              </span>
              <span className="text-xs font-bold uppercase tracking-wider bg-rose-500/20 text-rose-300 px-2.5 py-0.5 rounded border border-rose-400/40">
                ANSI/NETA MTS Section 9 & Table 100.18
              </span>
              <span className="text-xs bg-slate-800 text-slate-300 px-2 py-0.5 rounded border border-slate-700">
                NFPA 70B
              </span>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={() => setShowTable10018Modal(true)}
              className="px-3 py-1 bg-amber-500/20 hover:bg-amber-500/30 border border-amber-400/40 rounded-full text-xs text-amber-300 font-bold flex items-center gap-1.5 transition-colors"
            >
              <Info size={13} />
              <span>{isVi ? 'Tra cứu Bảng NETA Table 100.18' : 'View NETA Table 100.18'}</span>
            </button>

            <button
              onClick={() => setShowAiPromptModal(true)}
              className="px-3 py-1 bg-rose-500/20 hover:bg-rose-500/30 border border-rose-400/40 rounded-full text-xs text-rose-300 font-bold flex items-center gap-1.5 transition-colors"
              title={isVi ? "System Instructions Prompt chuẩn NETA MTS Section 9 & Table 100.18 cho Google AI Studio" : "System Instructions Prompt for Google AI Studio"}
            >
              <Sparkles size={13} className="text-rose-400" />
              <span>{isVi ? '📋 Prompt AI Studio (Song Ngữ)' : '📋 AI Studio Prompt'}</span>
            </button>
          </div>
        </div>

        {/* Row 2: Full-Width Horizontal Title */}
        <div className="pt-2 border-t border-slate-800">
          <h1 className="text-xl sm:text-2xl font-black text-white tracking-tight leading-snug">
            {isVi
              ? 'Khảo Sát & Kiểm Tra Hồng Ngoại Thiết Bị Điện (Electrical Thermographic Survey)'
              : 'Electrical Equipment Thermographic Survey & Infrared Inspection (NETA MTS)'}
          </h1>
          <p className="text-xs sm:text-sm text-slate-300 mt-0.5 leading-relaxed font-medium">
            {isVi
              ? 'Thử nghiệm bảo trì thiết bị điện phân phối, máy biến áp & thanh cái mang tải; đối soát chênh lệch nhiệt độ ΔT so với pha tương tự và môi trường theo chuẩn ANSI/NETA MTS (Mục 9 & Bảng Table 100.18).'
              : 'Maintenance testing for electrical switchgear, transformers & busbars under load; thermal anomaly evaluation against ANSI/NETA MTS Section 9 & Table 100.18 (Level 1-4).'}
          </p>
        </div>

        {/* Row 3: Action Toolbar */}
        <div className="flex flex-wrap items-center justify-between gap-3 pt-2 border-t border-slate-800/80">
          {/* Language Selector */}
          <div className="flex items-center gap-1 bg-slate-900/90 p-1 rounded-xl border border-slate-700/80">
            <button
              onClick={() => setLang('vi')}
              className={`px-3 py-1.5 rounded-lg text-xs font-black transition-all flex items-center gap-1.5 ${
                isVi
                  ? 'bg-rose-600 text-white shadow-md'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              <span>🇻🇳 Tiếng Việt</span>
            </button>
            <button
              onClick={() => setLang('en')}
              className={`px-3 py-1.5 rounded-lg text-xs font-black transition-all flex items-center gap-1.5 ${
                !isVi
                  ? 'bg-rose-600 text-white shadow-md'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              <span>🇺🇸 English (EN)</span>
            </button>
          </div>

          {/* Action Buttons */}
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={handleResetToSample}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title={isVi ? "Khôi phục dữ liệu mẫu hiện trường FSE" : "Reset sample FSE data"}
            >
              <RotateCcw size={14} />
              <span>{isVi ? 'Dữ Liệu Mẫu FSE' : 'Sample Data'}</span>
            </button>

            <button
              onClick={handleAddFinding}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title={isVi ? "Thêm điểm đo phát nhiệt mới" : "Add anomaly point"}
            >
              <Plus size={14} />
              <span>{isVi ? 'Thêm Điểm Đo' : 'Add Anomaly'}</span>
            </button>

            <button
              onClick={handleOpenJsonModal}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title={isVi ? "Nhập hoặc sao chép mã JSON" : "Import / Export JSON"}
            >
              <FileCheck2 size={14} />
              <span>{isVi ? 'Mã JSON' : 'JSON'}</span>
            </button>

            <button
              onClick={() => setShowPrintModal(true)}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title={isVi ? "In hoặc xuất biên bản PDF TEV" : "Print PDF Report"}
            >
              <Printer size={14} />
              <span>{isVi ? 'In Biên Bản PDF' : 'Print PDF'}</span>
            </button>

            <button
              onClick={handleSaveToCmms}
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-rose-800 hover:bg-rose-700 text-white border border-rose-500/40 flex items-center gap-1.5 transition-all shadow-md"
              title={isVi ? "Lưu biên bản kiểm tra vào kho báo cáo" : "Save report to CMMS"}
            >
              <Save size={15} />
              <span>{isVi ? 'Lưu vào Kho Báo Cáo' : 'Save to CMMS'}</span>
            </button>

            <button
              onClick={handleAnalyze}
              disabled={isAnalyzing}
              className="px-4 py-2 rounded-xl text-xs font-black bg-gradient-to-r from-amber-400 to-yellow-300 hover:from-amber-300 hover:to-yellow-200 text-slate-950 flex items-center gap-2 transition-all shadow-lg shadow-rose-950/40 disabled:opacity-50"
            >
              <Sparkles size={15} className={isAnalyzing ? 'animate-spin' : 'text-slate-950'} />
              <span>{isAnalyzing ? (isVi ? 'Đang Phân Tích NETA...' : 'Analyzing NETA...') : (isVi ? 'Phân Tích AI NETA Table 100.18' : 'AI NETA Table 100.18')}</span>
            </button>
          </div>
        </div>

        {/* Live NETA Diagnostic Strip */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5 pt-3 border-t border-rose-500/30 text-xs">
          {/* Circuit Loading Condition Section 9.3.3 */}
          <div className="bg-rose-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-rose-500/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-rose-200 block font-medium">
                {isVi ? 'Tải Mạch (NETA 9.3.3 ≥ 40%)' : 'Circuit Load (≥ 40%)'}
              </span>
              <span className={`text-base font-black ${liveStats.isLoadPass ? 'text-emerald-300' : 'text-rose-300'}`}>
                {liveStats.loadPct}% {liveStats.isLoadPass ? '(Pass)' : '(Low Load)'}
              </span>
            </div>
            {liveStats.isLoadPass ? (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            ) : (
              <AlertTriangle size={20} className="text-rose-400 shrink-0" />
            )}
          </div>

          {/* Max Hotspot Temp */}
          <div className="bg-rose-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-rose-500/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-rose-200 block font-medium">
                {isVi ? 'Điểm Nóng Nhất (Hotspot)' : 'Max Hotspot Temp'}
              </span>
              <span className={`text-base font-black ${liveStats.hasLevel4 ? 'text-rose-300' : 'text-amber-300'}`}>
                {liveStats.maxHotspot}°C
              </span>
            </div>
            <Flame size={20} className={liveStats.hasLevel4 ? 'text-rose-400 animate-pulse shrink-0' : 'text-amber-400 shrink-0'} />
          </div>

          {/* Max Delta T_comp */}
          <div className="bg-rose-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-rose-500/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-rose-200 block font-medium">
                {isVi ? 'Max ΔT Pha Tương Tự' : 'Max Similar ΔT'}
              </span>
              <span className={`text-base font-black ${liveStats.maxDeltaComp > 15 ? 'text-rose-300' : liveStats.maxDeltaComp >= 4 ? 'text-amber-300' : 'text-emerald-300'}`}>
                +{liveStats.maxDeltaComp}°C {liveStats.maxDeltaComp > 15 ? '(Level 4)' : liveStats.maxDeltaComp >= 4 ? '(Level 2)' : ''}
              </span>
            </div>
            {liveStats.maxDeltaComp > 15 ? (
              <AlertOctagon size={20} className="text-rose-400 animate-bounce shrink-0" />
            ) : (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            )}
          </div>

          {/* Max Delta T_amb */}
          <div className="bg-rose-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-rose-500/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-rose-200 block font-medium">
                {isVi ? 'Max ΔT So Môi Trường' : 'Max Ambient ΔT'}
              </span>
              <span className={`text-base font-black ${liveStats.maxDeltaAmb > 40 ? 'text-rose-300' : liveStats.maxDeltaAmb >= 21 ? 'text-orange-300' : 'text-emerald-300'}`}>
                +{liveStats.maxDeltaAmb}°C {liveStats.maxDeltaAmb > 40 ? '(Level 4)' : liveStats.maxDeltaAmb >= 21 ? '(Level 3)' : ''}
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
              {isVi ? 'Mở Kho Báo Cáo' : 'Open Reports'}
            </button>
          )}
        </div>
      )}

      {/* ---------------------------------------------------------------------- */}
      {/* SECTION A & FORM INPUTS */}
      {/* ---------------------------------------------------------------------- */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left Column: Site Info & Section 9 Survey Parameters */}
        <div className="space-y-6">
          {/* Site Information */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <Thermometer size={17} className="text-rose-600" />
                {isVi ? 'Thông Tin Trạm & Thiết Bị Khảo Sát' : 'Site & Survey Information'}
              </h2>
              <span className="text-[10px] font-semibold text-slate-500 bg-slate-100 px-2 py-0.5 rounded-full">
                NETA MTS Sec 9
              </span>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <CustomerSelectDropdown
                  label={isVi ? "Mã & Tên Khách Hàng (Dự Án / Nhà Máy)" : "Project / Substation Name"}
                  value={data.site_info.project_name}
                  onChange={(val, cust) => setData({
                    ...data,
                    site_info: { ...data.site_info, project_name: cust?.name || val }
                  })}
                  placeholder="Chọn hoặc bổ sung mã khách hàng..."
                />
              </div>

              <div>
                <EquipmentSelectDropdown
                  label={isVi ? "Tên & Mã Thiết Bị Khảo Sát (Survey Tag)" : "Survey Tag / Equipment ID"}
                  value={data.site_info.survey_tag}
                  selectedCustomer={data.site_info.project_name}
                  onChange={(val, eq) => setData({
                    ...data,
                    site_info: { ...data.site_info, survey_tag: eq?.id || val }
                  })}
                  placeholder="Chọn hoặc bổ sung thiết bị khảo sát..."
                />
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                <div>
                  <FseEngineerDropdown
                    label={isVi ? "Kỹ Sư FSE Đo Nhiệt" : "FSE Engineer"}
                    value={data.site_info.fse_name}
                    onChange={(val) => setData({
                      ...data,
                      site_info: { ...data.site_info, fse_name: val }
                    })}
                  />
                </div>
                <div>
                  <label className="text-slate-600 font-semibold block mb-1">
                    {isVi ? "Ngày Đo Nhiệt:" : "Test Date:"}
                  </label>
                  <input
                    type="date"
                    value={data.site_info.test_date}
                    onChange={(e) => setData({
                      ...data,
                      site_info: { ...data.site_info, test_date: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-rose-500 font-medium text-slate-800"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Survey Conditions & Parameters (NETA MTS Section 9.1 - 9.3) */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <Gauge size={17} className="text-rose-600" />
                {isVi ? 'Tham Số Khảo Sát (NETA MTS Mục 9.1 - 9.3)' : 'Survey Parameters (NETA MTS 9.1 - 9.3)'}
              </h2>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="text-slate-600 font-medium block mb-1">
                  {isVi ? "Model Camera Nhiệt (Độ nhạy ≤ 1°C per 9.3.1):" : "IR Camera Model (Sensitivity ≤ 1°C):"}
                </label>
                <input
                  type="text"
                  value={data.site_info.ir_camera_model || ''}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, ir_camera_model: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-rose-500 font-medium text-slate-800"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">
                    {isVi ? "Nhiệt Độ Môi Trường (°C):" : "Ambient Temp (°C):"}
                  </label>
                  <input
                    type="number"
                    step="0.5"
                    value={data.site_info.ambient_temperature_celsius ?? 32.5}
                    onChange={(e) => setData({
                      ...data,
                      site_info: { ...data.site_info, ambient_temperature_celsius: parseFloat(e.target.value) || 0 }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-rose-500 font-bold text-slate-800"
                  />
                </div>
                <div>
                  <label className="text-slate-600 font-medium block mb-1">
                    {isVi ? "% Tải Mạch (≥ 40% per 9.3.3):" : "Circuit Load % (≥ 40%):"}
                  </label>
                  <input
                    type="number"
                    step="1"
                    value={data.site_info.circuit_loading_percent ?? 68}
                    onChange={(e) => setData({
                      ...data,
                      site_info: { ...data.site_info, circuit_loading_percent: parseFloat(e.target.value) || 0 }
                    })}
                    className={`w-full px-3 py-2 bg-slate-50 border rounded-lg focus:outline-hidden focus:ring-2 focus:ring-rose-500 font-black ${
                      (data.site_info.circuit_loading_percent ?? 68) >= 40 ? 'border-emerald-300 text-emerald-700' : 'border-rose-300 text-rose-700'
                    }`}
                  />
                </div>
              </div>

              {/* Checklist Section 9.1 - 9.3 Controls */}
              <div className="pt-2 border-t border-slate-100 space-y-2">
                <span className="text-[11px] font-bold text-slate-700 block">
                  {isVi ? 'Quy trình kiểm tra tại site (NETA MTS 9.1 - 9.3):' : 'Field Inspection Check (NETA MTS 9.1 - 9.3):'}
                </span>

                <div className="p-2.5 rounded-lg bg-slate-50 border border-slate-200 space-y-1">
                  <div className="flex items-center justify-between">
                    <span className="font-semibold text-slate-700">A1. {isVi ? 'Tiếp cận & Cửa sổ hồng ngoại (IR Viewports):' : 'Access & IR Viewports:'}</span>
                    <span className="text-emerald-700 font-bold text-[10px] bg-emerald-100 px-1.5 py-0.5 rounded">Pass</span>
                  </div>
                  <input
                    type="text"
                    value={data.survey_parameters_section_9?.equipment_access_condition || ''}
                    onChange={(e) => setData({
                      ...data,
                      survey_parameters_section_9: {
                        ...data.survey_parameters_section_9!,
                        equipment_access_condition: e.target.value
                      }
                    })}
                    className="w-full text-[11px] bg-white border border-slate-200 rounded p-1 text-slate-700"
                  />
                </div>

                <div className="p-2.5 rounded-lg bg-slate-50 border border-slate-200 space-y-1">
                  <div className="flex items-center justify-between">
                    <span className="font-semibold text-slate-700">A4. {isVi ? 'Báo cáo đầy đủ 8 thành phần Mục 9.2:' : 'Report Complete (8 Elements):'}</span>
                    <input
                      type="checkbox"
                      checked={data.survey_parameters_section_9?.report_elements_complete ?? true}
                      onChange={(e) => setData({
                        ...data,
                        survey_parameters_section_9: {
                          ...data.survey_parameters_section_9!,
                          report_elements_complete: e.target.checked
                        }
                      })}
                      className="w-4 h-4 text-rose-600 rounded"
                    />
                  </div>
                  <span className="text-[10px] text-slate-500 block">
                    {isVi ? 'Bao gồm: Mô tả, Điểm bất thường, ΔT, Nguyên nhân, Inaccessible, Tải, Ảnh nhiệt & Hành động.' : 'Includes: Description, Anomalies, ΔT, Probable cause, Inaccessible, Load, Thermogram & Action.'}
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Center & Right Column: Thermographic Findings Table */}
        <div className="lg:col-span-2 space-y-6">
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-100 pb-3">
              <div>
                <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                  <Flame size={17} className="text-rose-600" />
                  {isVi ? 'Danh Sách Điểm Đo Phát Nhiệt (Thermographic Findings)' : 'Thermographic Findings & Anomaly List'}
                </h2>
                <span className="text-[11px] text-slate-500">
                  {isVi
                    ? `So sánh với pha tương tự (ΔT_comp) và nhiệt độ môi trường (T_amb = ${liveStats.ambientTemp}°C) per NETA Table 100.18`
                    : `Evaluated against similar component (ΔT_comp) and ambient (T_amb = ${liveStats.ambientTemp}°C) per NETA Table 100.18`}
                </span>
              </div>

              <button
                onClick={handleAddFinding}
                className="px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-700 text-xs font-bold rounded-lg border border-rose-200 flex items-center gap-1 transition-colors self-start sm:self-auto"
              >
                <Plus size={14} />
                <span>{isVi ? 'Thêm Điểm Đo' : 'Add Finding'}</span>
              </button>
            </div>

            {/* Measurement List */}
            <div className="space-y-3.5">
              {(data.thermographic_findings || []).map((item, idx) => {
                const evalItem = liveStats.items[idx];
                return (
                  <div
                    key={idx}
                    className={`p-4 rounded-xl border transition-all ${
                      evalItem?.level === 4
                        ? 'bg-rose-50/80 border-rose-300 ring-1 ring-rose-400'
                        : evalItem?.level === 3
                        ? 'bg-orange-50/80 border-orange-300'
                        : evalItem?.level === 2
                        ? 'bg-amber-50/80 border-amber-300'
                        : evalItem?.level === 1
                        ? 'bg-blue-50/80 border-blue-200'
                        : 'bg-slate-50 border-slate-200'
                    }`}
                  >
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-2.5 mb-2.5 border-b border-slate-200/70">
                      <div className="flex items-center gap-2 flex-1">
                        <span className="w-5 h-5 rounded-full bg-slate-200 text-slate-800 text-[10px] font-bold flex items-center justify-center shrink-0">
                          {idx + 1}
                        </span>
                        <input
                          type="text"
                          value={item.location}
                          onChange={(e) => handleFindingChange(idx, 'location', e.target.value)}
                          placeholder="Vị trí thiết bị (Location)..."
                          className="font-bold text-xs text-slate-800 bg-white/90 border border-slate-200 rounded px-2.5 py-1 flex-1 focus:ring-1 focus:ring-rose-500"
                        />
                      </div>

                      <div className="flex items-center gap-2">
                        {/* Status Badge */}
                        <span
                          className={`px-3 py-1 rounded-full text-[10px] font-black border uppercase tracking-wider ${
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
                          {evalItem?.level_en || `Level ${evalItem?.level}`}
                        </span>

                        {(data.thermographic_findings || []).length > 1 && (
                          <button
                            onClick={() => handleRemoveFinding(idx)}
                            className="p-1 text-slate-400 hover:text-rose-600 transition-colors"
                            title="Xóa điểm đo này"
                          >
                            <Trash2 size={15} />
                          </button>
                        )}
                      </div>
                    </div>

                    {/* Inputs & Delta T Calculations */}
                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5 text-xs">
                      <div>
                        <label className="text-slate-600 text-[10px] font-semibold block">
                          {isVi ? 'Nhiệt Độ Đo (T):' : 'Measured Temp (T):'}
                        </label>
                        <div className="flex items-center gap-1 mt-0.5">
                          <input
                            type="number"
                            step="0.1"
                            value={item.measured_temperature_celsius}
                            onChange={(e) => handleFindingChange(idx, 'measured_temperature_celsius', parseFloat(e.target.value) || 0)}
                            className="w-full px-2 py-1 bg-white border border-slate-300 rounded font-black text-slate-900 focus:ring-1 focus:ring-rose-500"
                          />
                          <span className="text-slate-600 font-bold">°C</span>
                        </div>
                      </div>

                      <div>
                        <label className="text-slate-600 text-[10px] font-semibold block">
                          {isVi ? 'Pha Đối Chứng (T_ref):' : 'Similar Ref Temp:'}
                        </label>
                        <div className="flex items-center gap-1 mt-0.5">
                          <input
                            type="number"
                            step="0.1"
                            value={item.reference_component_temp_celsius}
                            onChange={(e) => handleFindingChange(idx, 'reference_component_temp_celsius', parseFloat(e.target.value) || 0)}
                            className="w-full px-2 py-1 bg-white border border-slate-300 rounded font-medium text-slate-800 focus:ring-1 focus:ring-rose-500"
                          />
                          <span className="text-slate-600 font-bold">°C</span>
                        </div>
                      </div>

                      {/* Calculated Delta T_comp */}
                      <div className="bg-white/90 p-2 rounded-lg border border-slate-200">
                        <span className="text-slate-500 text-[10px] block font-medium">
                          {isVi ? 'ΔT Pha Tương Tự:' : 'Similar ΔT (comp):'}
                        </span>
                        <span className={`text-xs font-black block mt-0.5 ${
                          (evalItem?.delta_t_comp ?? 0) > 15 ? 'text-rose-600' : (evalItem?.delta_t_comp ?? 0) >= 4 ? 'text-amber-600' : 'text-slate-800'
                        }`}>
                          {evalItem?.delta_t_comp ?? 0 > 0 ? '+' : ''}{evalItem?.delta_t_comp ?? 0}°C
                        </span>
                      </div>

                      {/* Calculated Delta T_amb */}
                      <div className="bg-white/90 p-2 rounded-lg border border-slate-200">
                        <span className="text-slate-500 text-[10px] block font-medium">
                          {isVi ? 'ΔT So Môi Trường:' : 'Ambient ΔT (amb):'}
                        </span>
                        <span className={`text-xs font-black block mt-0.5 ${
                          (evalItem?.delta_t_amb ?? 0) > 40 ? 'text-rose-600' : (evalItem?.delta_t_amb ?? 0) >= 21 ? 'text-orange-600' : 'text-slate-800'
                        }`}>
                          {evalItem?.delta_t_amb ?? 0 > 0 ? '+' : ''}{evalItem?.delta_t_amb ?? 0}°C
                        </span>
                      </div>
                    </div>

                    {/* Component Type & Suspected Cause */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 mt-2.5 pt-2 border-t border-slate-200/60 text-xs">
                      <div>
                        <label className="text-[10px] text-slate-500 block">
                          {isVi ? 'Loại cấu thành (Component Type):' : 'Component Type:'}
                        </label>
                        <input
                          type="text"
                          value={item.component_type}
                          onChange={(e) => handleFindingChange(idx, 'component_type', e.target.value)}
                          className="w-full text-[11px] bg-white border border-slate-200 rounded px-2 py-1 mt-0.5"
                        />
                      </div>
                      <div>
                        <label className="text-[10px] text-slate-500 block">
                          {isVi ? 'Nguyên nhân nghi ngờ (Suspected Cause):' : 'Suspected Cause:'}
                        </label>
                        <input
                          type="text"
                          value={item.suspected_cause || ''}
                          onChange={(e) => handleFindingChange(idx, 'suspected_cause', e.target.value)}
                          className="w-full text-[11px] bg-white border border-slate-200 rounded px-2 py-1 mt-0.5"
                        />
                      </div>
                    </div>

                    {/* Action Guideline per Table 100.18 */}
                    <div className="mt-2 pt-2 border-t border-slate-200/60 text-[11px] flex flex-col sm:flex-row sm:items-center justify-between gap-1">
                      <span className="text-slate-700">
                        <strong>{isVi ? 'Hành động khuyến nghị:' : 'Recommended Action:'}</strong>{' '}
                        <span className="font-semibold text-rose-700">
                          {isVi ? evalItem?.action_recommendation_vi : evalItem?.action_recommendation_en}
                        </span>
                      </span>
                      <span className="text-[10px] font-bold text-slate-400 bg-slate-100 px-2 py-0.5 rounded shrink-0">
                        NETA Table 100.18
                      </span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </div>

      {/* ---------------------------------------------------------------------- */}
      {/* AI ANALYSIS OUTPUT SECTION (MATCHING SECTION IV SPECIFICATIONS) */}
      {/* ---------------------------------------------------------------------- */}
      {analysisResult && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-xl overflow-hidden animate-in fade-in slide-in-from-bottom-3 duration-300">
          <div className="p-5 sm:p-6 bg-gradient-to-r from-slate-950 via-slate-900 to-rose-950 text-white flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-4 border-b border-rose-900/60">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-rose-500/20 border border-rose-400/30 flex items-center justify-center text-rose-300 shrink-0">
                <Sparkles size={20} />
              </div>
              <div>
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="text-[10px] uppercase tracking-wider font-extrabold px-2 py-0.5 rounded-full bg-white/10 text-white">
                    TEV Platform AI Evaluation
                  </span>
                  <span className={`text-[11px] font-black px-2.5 py-0.5 rounded-full border ${
                    analysisResult.overall_status.includes('LEVEL 4') || analysisResult.overall_status === 'REPAIR_IMMEDIATELY'
                      ? 'bg-rose-600 text-white border-rose-400 animate-pulse'
                      : analysisResult.overall_status === 'MONITOR'
                      ? 'bg-orange-500 text-white border-orange-400'
                      : analysisResult.overall_status === 'REPAIR_SCHEDULED'
                      ? 'bg-amber-400 text-slate-900 border-amber-500'
                      : 'bg-emerald-500 text-white border-emerald-400'
                  }`}>
                    {analysisResult.overall_status_en || analysisResult.overall_status}
                  </span>
                </div>
                <h3 className="text-base sm:text-lg font-black text-white mt-1">
                  {isVi
                    ? 'Báo Cáo Đánh Giá Khảo Sát Nhiệt Độ Hồng Ngoại (ANSI/NETA MTS Section 9)'
                    : 'Thermographic Survey Field Evaluation Report (ANSI/NETA MTS Section 9)'}
                </h3>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={() => setShowPrintModal(true)}
                className="px-3.5 py-2 bg-white/10 hover:bg-white/20 text-white text-xs font-bold rounded-xl flex items-center gap-1.5 transition-colors border border-white/20 shadow-xs"
              >
                <Printer size={14} />
                <span>{isVi ? 'In Biên Bản PDF' : 'Print PDF'}</span>
              </button>
              <button
                onClick={handleSaveToCmms}
                className="px-4 py-2 bg-rose-700 hover:bg-rose-600 text-white text-xs font-bold rounded-xl flex items-center gap-1.5 shadow-md transition-colors"
              >
                <Save size={14} />
                <span>{isVi ? 'Lưu vào Kho Báo Cáo' : 'Save Report'}</span>
              </button>
            </div>
          </div>

          {/* Section 1: Detailed Table Formatted with Wide Columns */}
          <div className="p-6 space-y-6">
            <div className="space-y-3">
              <div className="flex items-center justify-between">
                <h4 className="text-sm font-black text-slate-800 uppercase tracking-wide flex items-center gap-2">
                  <span className="w-2 h-2 rounded-full bg-rose-600" />
                  1. {isVi ? 'BẢNG PHÂN TÍCH CHI TIẾT KHẢO SÁT HỒNG NGOẠI' : 'DETAILED THERMOGRAPHIC EVALUATION TABLE'}
                </h4>
                <span className="text-[11px] text-slate-500 font-medium">
                  ANSI/NETA MTS Section 9 & Table 100.18
                </span>
              </div>

              <div className="overflow-x-auto rounded-xl border border-slate-200">
                <table className="w-full text-left text-xs border-collapse min-w-[1050px]">
                  <thead>
                    <tr className="bg-slate-100 text-slate-800 font-bold border-b border-slate-200">
                      <th className="py-3 px-4 w-[28%] min-w-[260px]">{isVi ? 'Vị trí Thiết bị / Location' : 'Equipment Location'}</th>
                      <th className="py-3 px-3 w-[20%]">{isVi ? 'Cấu thành & Điểm Phát nhiệt' : 'Component & Anomaly Point'}</th>
                      <th className="py-3 px-3 text-center w-[12%]">{isVi ? 'Nhiệt độ Đo' : 'Measured Temp'}</th>
                      <th className="py-3 px-3 text-center w-[14%]">{isVi ? 'ΔT Pha Tương Tự' : 'Similar ΔT'}</th>
                      <th className="py-3 px-3 text-center w-[12%]">{isVi ? 'ΔT Môi Trường' : 'Ambient ΔT'}</th>
                      <th className="py-3 px-3 text-center w-[14%]">{isVi ? 'Phân loại Table 100.18' : 'Table 100.18 Class'}</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 font-medium text-slate-700">
                    {liveStats.items.map((ev, i) => (
                      <tr key={i} className="hover:bg-slate-50 transition-colors">
                        <td className="py-3 px-4 font-bold text-slate-900 break-words">
                          {ev.location}
                        </td>
                        <td className="py-3 px-3 text-slate-600 break-words">
                          {ev.component_type}
                        </td>
                        <td className="py-3 px-3 text-center font-black text-slate-900">
                          {ev.measured_temp_c}°C
                        </td>
                        <td className="py-3 px-3 text-center font-bold">
                          <span className={`${ev.delta_t_comp > 15 ? 'text-rose-600 font-black' : ev.delta_t_comp >= 4 ? 'text-amber-600' : 'text-slate-700'}`}>
                            +{ev.delta_t_comp}°C
                          </span>
                          <span className="block text-[10px] text-slate-400">
                            {ev.delta_t_comp > 15 ? '(> 15°C)' : ev.delta_t_comp >= 4 ? '(4°C - 15°C)' : '(1°C - 3°C)'}
                          </span>
                        </td>
                        <td className="py-3 px-3 text-center font-bold">
                          <span className={`${ev.delta_t_amb > 40 ? 'text-rose-600 font-black' : ev.delta_t_amb >= 21 ? 'text-orange-600' : 'text-slate-700'}`}>
                            +{ev.delta_t_amb}°C
                          </span>
                          <span className="block text-[10px] text-slate-400">
                            {ev.delta_t_amb > 40 ? '(> 40°C)' : ev.delta_t_amb >= 21 ? '(21°C - 40°C)' : '(11°C - 20°C)'}
                          </span>
                        </td>
                        <td className="py-3 px-3 text-center">
                          <span className={`px-2.5 py-1 rounded-full text-[10px] font-black uppercase tracking-wider inline-block ${
                            ev.level === 4
                              ? 'bg-rose-600 text-white'
                              : ev.level === 3
                              ? 'bg-orange-500 text-white'
                              : ev.level === 2
                              ? 'bg-amber-400 text-slate-900'
                              : 'bg-blue-500 text-white'
                          }`}>
                            {ev.level_en}
                          </span>
                          <span className="block text-[10px] text-rose-700 mt-0.5 font-bold">
                            {isVi ? ev.action_recommendation_vi : ev.action_recommendation_en}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Markdown Report Body */}
            <div className="prose prose-slate max-w-none text-xs sm:text-sm pt-4 border-t border-slate-200">
              <div
                className="space-y-4"
                dangerouslySetInnerHTML={{
                  __html: analysisResult.markdown_report
                    .replace(/### (.*?)\n/g, '<h3 class="text-base font-bold text-slate-900 border-b border-slate-200 pb-2 mt-4 mb-2 flex items-center gap-2">$1</h3>')
                    .replace(/#### (.*?)\n/g, '<h4 class="text-sm font-bold text-slate-800 mt-3 mb-1.5">$1</h4>')
                    .replace(/\n\n/g, '<br/>')
                }}
              />
            </div>
          </div>
        </div>
      )}

      {/* ---------------------------------------------------------------------- */}
      {/* NETA TABLE 100.18 REFERENCE MODAL */}
      {/* ---------------------------------------------------------------------- */}
      {showTable10018Modal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/70 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-white rounded-2xl max-w-3xl w-full p-6 shadow-2xl border border-slate-200 space-y-4 max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <Flame size={20} className="text-rose-600" />
                <h3 className="font-bold text-base text-slate-900">
                  ANSI/NETA MTS Table 100.18 - Thermographic Survey Actions
                </h3>
              </div>
              <button
                onClick={() => setShowTable10018Modal(false)}
                className="text-slate-400 hover:text-slate-600 text-lg font-bold"
              >
                ✕
              </button>
            </div>

            <p className="text-xs text-slate-600">
              {isVi
                ? 'Bảng giá trị tham chiếu độ chênh nhiệt độ (ΔT) và hành động xử lý kỹ thuật đề xuất theo chuẩn bảo dưỡng thiết bị điện ANSI/NETA MTS Table 100.18:'
                : 'Reference table for temperature difference (ΔT) and suggested actions based on temperature rise per ANSI/NETA MTS Table 100.18:'}
            </p>

            <div className="overflow-x-auto rounded-xl border border-slate-200">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="bg-slate-100 text-slate-800 font-bold border-b border-slate-200">
                    <th className="py-2.5 px-3">Mức / Level</th>
                    <th className="py-2.5 px-3">ΔT Pha Tương Tự (Similar)</th>
                    <th className="py-2.5 px-3">ΔT So Môi Trường (Ambient)</th>
                    <th className="py-2.5 px-3">Hành động Khuyến nghị / Recommended Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-slate-700">
                  <tr className="bg-blue-50/50">
                    <td className="py-2.5 px-3 font-bold text-blue-700">Level 1</td>
                    <td className="py-2.5 px-3 font-semibold">1°C - 3°C</td>
                    <td className="py-2.5 px-3 font-semibold">1°C - 10°C</td>
                    <td className="py-2.5 px-3">
                      <strong>Cần kiểm tra / Possible deficiency; warrants investigation</strong>
                      <span className="block text-[11px] text-slate-500">Ghi nhận; theo dõi trong kỳ bảo dưỡng định kỳ tiếp theo.</span>
                    </td>
                  </tr>
                  <tr className="bg-amber-50/50">
                    <td className="py-2.5 px-3 font-bold text-amber-700">Level 2</td>
                    <td className="py-2.5 px-3 font-semibold">4°C - 15°C</td>
                    <td className="py-2.5 px-3 font-semibold">11°C - 20°C</td>
                    <td className="py-2.5 px-3">
                      <strong>Khuyết tật tiềm ẩn; Sửa chữa khi điều kiện cho phép / Indicates probable deficiency; repair as time permits</strong>
                      <span className="block text-[11px] text-slate-500">Lập kế hoạch siết lực hoặc vệ sinh tiếp điểm khi có lịch dừng máy phù hợp.</span>
                    </td>
                  </tr>
                  <tr className="bg-orange-50/50">
                    <td className="py-2.5 px-3 font-bold text-orange-700">Level 3</td>
                    <td className="py-2.5 px-3 font-semibold text-slate-400"><em>(N/A)</em></td>
                    <td className="py-2.5 px-3 font-semibold">21°C - 40°C</td>
                    <td className="py-2.5 px-3">
                      <strong>Giám sát chặt chẽ cho đến khi xử lý / Monitor until corrective measures can be accomplished</strong>
                      <span className="block text-[11px] text-slate-500">Giám sát liên tục tải điện và nhiệt độ cho đến khi thực hiện xong biện pháp khắc phục.</span>
                    </td>
                  </tr>
                  <tr className="bg-rose-50/60 font-bold">
                    <td className="py-2.5 px-3 text-rose-700">Level 4</td>
                    <td className="py-2.5 px-3 text-rose-700">&gt; 15°C</td>
                    <td className="py-2.5 px-3 text-rose-700">&gt; 40°C</td>
                    <td className="py-2.5 px-3 text-rose-900">
                      <strong>Sai lệch nghiêm trọng; SỬA CHỮA NGAY LẬP TỨC / Major discrepancy; repair immediately</strong>
                      <span className="block text-[11px] text-rose-600 font-normal">Rất nguy hiểm (nguy cơ cháy nổ / phóng điện hồ quang); bắt buộc cô lập và khắc phục ngay lập tức.</span>
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>

            <div className="flex justify-end pt-2">
              <button
                onClick={() => setShowTable10018Modal(false)}
                className="px-4 py-1.5 bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold rounded-lg"
              >
                {isVi ? 'Đóng' : 'Close'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ---------------------------------------------------------------------- */}
      {/* JSON IMPORT / EXPORT MODAL */}
      {/* ---------------------------------------------------------------------- */}
      {showJsonModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-white rounded-2xl max-w-2xl w-full p-6 shadow-2xl border border-slate-200 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <FileCheck2 size={18} className="text-rose-600" />
                <h3 className="font-bold text-sm text-slate-800">
                  {isVi ? 'Dữ Liệu JSON Khảo Sát Hồng Ngoại (NETA MTS Sec 9 & Table 100.18)' : 'JSON Input Payload (NETA MTS Sec 9 & Table 100.18)'}
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
              {isVi
                ? 'Kỹ sư FSE có thể sao chép chuỗi JSON này để lưu trữ hoặc dán dữ liệu đo mới từ Camera nhiệt vào đây và bấm "Áp Dụng Dữ Liệu":'
                : 'Copy or paste raw thermographic inspection JSON to populate the field checklist:'}
            </p>

            <textarea
              value={jsonText}
              onChange={(e) => setJsonText(e.target.value)}
              rows={14}
              className="w-full p-3 font-mono text-xs bg-slate-900 text-rose-200 rounded-xl border border-slate-700 focus:outline-hidden focus:ring-2 focus:ring-rose-500"
            />

            <div className="flex items-center justify-between pt-2">
              <button
                onClick={handleCopyJson}
                className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-lg flex items-center gap-1.5 transition-colors"
              >
                {copied ? <Check size={14} className="text-emerald-600" /> : <Copy size={14} />}
                <span>{copied ? (isVi ? 'Đã Sao Chép!' : 'Copied!') : (isVi ? 'Sao Chép JSON' : 'Copy JSON')}</span>
              </button>

              <div className="flex gap-2">
                <button
                  onClick={() => setShowJsonModal(false)}
                  className="px-3.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-medium rounded-lg"
                >
                  {isVi ? 'Hủy' : 'Cancel'}
                </button>
                <button
                  onClick={handleApplyJson}
                  className="px-4 py-1.5 bg-rose-600 hover:bg-rose-500 text-white text-xs font-bold rounded-lg shadow-sm"
                >
                  {isVi ? 'Áp Dụng Dữ Liệu' : 'Apply JSON'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ---------------------------------------------------------------------- */}
      {/* GOOGLE AI STUDIO SYSTEM INSTRUCTION PROMPT MODAL */}
      {/* ---------------------------------------------------------------------- */}
      {showAiPromptModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/70 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-white rounded-2xl max-w-3xl w-full p-6 shadow-2xl border border-slate-200 space-y-4 max-h-[90vh] flex flex-col">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3 shrink-0">
              <div className="flex items-center gap-2">
                <Sparkles size={18} className="text-rose-600" />
                <h3 className="font-bold text-sm text-slate-900">
                  {isVi ? 'System Instruction Prompt Dành Cho Google AI Studio (Chuẩn NETA MTS Sec 9 & Table 100.18)' : 'Google AI Studio System Instruction Prompt (NETA MTS Sec 9 & Table 100.18)'}
                </h3>
              </div>
              <button
                onClick={() => setShowAiPromptModal(false)}
                className="text-slate-400 hover:text-slate-600 text-sm font-bold"
              >
                ✕
              </button>
            </div>

            <div className="text-xs text-slate-600 space-y-1.5 shrink-0 bg-rose-50 p-3 rounded-xl border border-rose-200">
              <p className="font-bold text-rose-900">
                {isVi ? '📋 Hướng dẫn sử dụng trên Google AI Studio:' : '📋 How to use on Google AI Studio:'}
              </p>
              <p>
                {isVi
                  ? 'Bấm nút "Sao Chép Prompt", sau đó mở Google AI Studio và dán trực tiếp vào mục "System Instructions". Mô hình AI sẽ tự động thẩm định số liệu đo nhiệt hồng ngoại theo đúng quy tắc NETA Table 100.18 và phản hồi song ngữ EN/VN hoàn chỉnh.'
                  : 'Click "Copy Prompt", open Google AI Studio and paste directly into "System Instructions". The AI model will evaluate field thermography data strictly per NETA Table 100.18.'}
              </p>
            </div>

            <div className="flex-1 overflow-y-auto">
              <textarea
                readOnly
                value={`You are an AI Expert Field Service Assistant (TEV Platform AI) specializing in Electrical Thermographic Survey and Thermal Infrared Diagnostics. Your task is to receive field thermographic inspection data from Field Service Engineers (FSE) and automatically evaluate thermal anomalies STRICTLY against ANSI/NETA MTS (Section 9) and NETA Table 100.18 (Suggested Actions Based on Temperature Rise). You MUST provide all evaluation outputs in BILINGUAL format (English / Vietnamese EN/VN) and NEVER include any citation reference numbers.

EVALUATION RULES & REFERENCE STANDARDS (ANSI/NETA MTS SECTION 9 & TABLE 100.18):

1. INSPECTION & REPORTING PARAMETERS (NETA MTS SECTION 9.1 - 9.3):
- Equipment Access & IR Windows (Section 9.1): Verify physical/mechanical condition and safe optical access through IR ports or opened covers.
- Report Completeness (Section 9.2): Verify 8 required elements (Equipment description, Discrepancies, Delta T values, Probable cause, Inaccessible areas, Load conditions, Thermograms/photos, Recommended actions).
- IR Camera Sensitivity (Section 9.3.1): IR camera MUST detect minimum temperature difference of 1°C at 30°C.
- Circuit Load Condition (Section 9.3.3 & NFPA 70B): Thermographic survey MUST be performed under normal circuit loading, not less than 40% nominal load.

2. THERMAL EVALUATION MATRIX & ACTIONS (NETA MTS TABLE 100.18):
Evaluate temperature rise (Delta T) using two comparison methods:
METHOD 1: Component-to-Similar-Component Delta T (Under Similar Loading):
- Delta T = 1°C - 3°C -> LEVEL 1: Possible deficiency; warrants investigation.
- Delta T = 4°C - 15°C -> LEVEL 2: Indicates probable deficiency; repair as time permits.
- Delta T > 15°C -> LEVEL 4: Major discrepancy; REPAIR IMMEDIATELY.

METHOD 2: Component-to-Ambient Air Delta T:
- Delta T = 1°C - 10°C -> LEVEL 1: Possible deficiency; warrants investigation.
- Delta T = 11°C - 20°C -> LEVEL 2: Indicates probable deficiency; repair as time permits.
- Delta T = 21°C - 40°C -> LEVEL 3: Monitor until corrective measures can be accomplished.
- Delta T > 40°C -> LEVEL 4: Major discrepancy; REPAIR IMMEDIATELY.

AI RESPONSE STRUCTURE (ALL IN BILINGUAL EN/VN):
When receiving JSON input from FSE, output a Markdown response containing 4 sections:
1. OVERALL STATUS / TRẠNG THÁI TỔNG QUAN: PASS, INVESTIGATE, MONITOR, or CRITICAL REPAIR IMMEDIATELY.
2. DETAILED THERMOGRAPHIC SURVEY EVALUATION TABLE / BẢNG PHÂN TÍCH CHI TIẾT KHẢO SÁT HỒNG NGOẠI: Compare findings against NETA MTS Section 9 & Table 100.18.
3. ROOT CAUSE & ANOMALY RISK ASSESSMENT / NGUYÊN NHÂN SƠ BỘ & RỦI RO PHÁT NHIỆT: Electrical contact resistance, loose bolts, phase imbalance, overloaded conductor, oxidation.
4. FSE ACTIONABLE RECOMMENDATIONS / KHUYẾN NGHỊ KỸ THUẬT FSE: Specific corrective measures per NETA Table 100.18 and NETA MTS Table 100.12 (Torque specifications).`}
                rows={16}
                className="w-full p-3 font-mono text-xs bg-slate-900 text-rose-200 rounded-xl border border-slate-700 focus:outline-hidden"
              />
            </div>

            <div className="flex items-center justify-between pt-2 border-t border-slate-100 shrink-0">
              <button
                onClick={() => {
                  const promptStr = `You are an AI Expert Field Service Assistant (TEV Platform AI) specializing in Electrical Thermographic Survey and Thermal Infrared Diagnostics. Your task is to receive field thermographic inspection data from Field Service Engineers (FSE) and automatically evaluate thermal anomalies STRICTLY against ANSI/NETA MTS (Section 9) and NETA Table 100.18 (Suggested Actions Based on Temperature Rise). You MUST provide all evaluation outputs in BILINGUAL format (English / Vietnamese EN/VN) and NEVER include any citation reference numbers.

EVALUATION RULES & REFERENCE STANDARDS (ANSI/NETA MTS SECTION 9 & TABLE 100.18):

1. INSPECTION & REPORTING PARAMETERS (NETA MTS SECTION 9.1 - 9.3):
- Equipment Access & IR Windows (Section 9.1): Verify physical/mechanical condition and safe optical access through IR ports or opened covers.
- Report Completeness (Section 9.2): Verify 8 required elements (Equipment description, Discrepancies, Delta T values, Probable cause, Inaccessible areas, Load conditions, Thermograms/photos, Recommended actions).
- IR Camera Sensitivity (Section 9.3.1): IR camera MUST detect minimum temperature difference of 1°C at 30°C.
- Circuit Load Condition (Section 9.3.3 & NFPA 70B): Thermographic survey MUST be performed under normal circuit loading, not less than 40% nominal load.

2. THERMAL EVALUATION MATRIX & ACTIONS (NETA MTS TABLE 100.18):
Evaluate temperature rise (Delta T) using two comparison methods:
METHOD 1: Component-to-Similar-Component Delta T (Under Similar Loading):
- Delta T = 1°C - 3°C -> LEVEL 1: Possible deficiency; warrants investigation.
- Delta T = 4°C - 15°C -> LEVEL 2: Indicates probable deficiency; repair as time permits.
- Delta T > 15°C -> LEVEL 4: Major discrepancy; REPAIR IMMEDIATELY.

METHOD 2: Component-to-Ambient Air Delta T:
- Delta T = 1°C - 10°C -> LEVEL 1: Possible deficiency; warrants investigation.
- Delta T = 11°C - 20°C -> LEVEL 2: Indicates probable deficiency; repair as time permits.
- Delta T = 21°C - 40°C -> LEVEL 3: Monitor until corrective measures can be accomplished.
- Delta T > 40°C -> LEVEL 4: Major discrepancy; REPAIR IMMEDIATELY.

AI RESPONSE STRUCTURE (ALL IN BILINGUAL EN/VN):
When receiving JSON input from FSE, output a Markdown response containing 4 sections:
1. OVERALL STATUS / TRẠNG THÁI TỔNG QUAN: PASS, INVESTIGATE, MONITOR, or CRITICAL REPAIR IMMEDIATELY.
2. DETAILED THERMOGRAPHIC SURVEY EVALUATION TABLE / BẢNG PHÂN TÍCH CHI TIẾT KHẢO SÁT HỒNG NGOẠI: Compare findings against NETA MTS Section 9 & Table 100.18.
3. ROOT CAUSE & ANOMALY RISK ASSESSMENT / NGUYÊN NHÂN SƠ BỘ & RỦI RO PHÁT NHIỆT: Electrical contact resistance, loose bolts, phase imbalance, overloaded conductor, oxidation.
4. FSE ACTIONABLE RECOMMENDATIONS / KHUYẾN NGHỊ KỸ THUẬT FSE: Specific corrective measures per NETA Table 100.18 and NETA MTS Table 100.12 (Torque specifications).`;
                  navigator.clipboard.writeText(promptStr);
                  setCopiedPrompt(true);
                  setTimeout(() => setCopiedPrompt(false), 2500);
                }}
                className="px-4 py-2 bg-rose-600 hover:bg-rose-500 text-white text-xs font-bold rounded-xl flex items-center gap-1.5 transition-all shadow-md shadow-rose-600/30"
              >
                {copiedPrompt ? <Check size={15} /> : <Copy size={15} />}
                <span>{copiedPrompt ? (isVi ? 'Đã Sao Chép Prompt Thành Công!' : 'Prompt Copied to Clipboard!') : (isVi ? 'Sao Chép Prompt System Instructions' : 'Copy System Instructions')}</span>
              </button>

              <button
                onClick={() => setShowAiPromptModal(false)}
                className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-xl"
              >
                {isVi ? 'Đóng' : 'Close'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ---------------------------------------------------------------------- */}
      {/* TEV OFFICIAL REPORT PRINT MODAL */}
      {/* ---------------------------------------------------------------------- */}
      {showPrintModal && (
        <TevReportPrintModal
          isOpen={showPrintModal}
          onClose={() => setShowPrintModal(false)}
          title="BIÊN BẢN KHẢO SÁT & ĐÁNH GIÁ NHIỆT ĐỘ HỒNG NGOẠI (THERMOGRAPHIC SURVEY)"
          titleEn="ELECTRICAL EQUIPMENT THERMOGRAPHIC SURVEY EVALUATION REPORT"
          standard="ANSI/NETA MTS Section 9 & Table 100.18 (NFPA 70B)"
          fseName={data.site_info.fse_name}
          siteInfo={{
            project_name: data.site_info.project_name,
            transformer_tag: data.site_info.survey_tag,
            fse_name: data.site_info.fse_name,
            test_date: data.site_info.test_date
          }}
          overallStatus={liveStats.hasLevel4 ? 'FAIL' : liveStats.hasLevel3 || liveStats.hasLevel2 ? 'INVESTIGATE' : 'PASS'}
          statusReason={analysisResult?.status_reason || 'Đã đối chiếu trực tiếp với quy chuẩn kỹ thuật ANSI/NETA MTS Section 9 & Table 100.18.'}
          visualChecks={[
            {
              label: 'Điều kiện tiếp cận & Cửa sổ hồng ngoại (IR Viewports 9.1)',
              value: data.survey_parameters_section_9?.equipment_access_condition || 'Tiếp cận an toàn',
              pass: true
            },
            {
              label: 'Độ nhạy & Hiệu chuẩn máy ảnh nhiệt (Camera Sensitivity 9.3.1)',
              value: data.site_info.ir_camera_model || 'FLIR T865 (Sensitivity < 0.03°C)',
              pass: true
            },
            {
              label: 'Điều kiện mang tải mạch điện (Circuit Load ≥ 40% per 9.3.3)',
              value: `${data.site_info.circuit_loading_percent}% tải định mức`,
              pass: (data.site_info.circuit_loading_percent ?? 68) >= 40
            },
            {
              label: 'Báo cáo đầy đủ 8 thành phần quy định NETA MTS Mục 9.2',
              value: (data.survey_parameters_section_9?.report_elements_complete ?? true) ? 'Đầy đủ 8 thành phần' : 'Thiếu dữ liệu',
              pass: data.survey_parameters_section_9?.report_elements_complete ?? true
            }
          ]}
          electricalTable={liveStats.items.map(item => ({
            item: item.location,
            itemEn: `${item.component_type} @ ${item.location}`,
            measured: `${item.measured_temp_c}°C`,
            standard: 'ΔT_comp ≤ 15°C & ΔT_amb ≤ 40°C',
            standardEn: 'Similar ΔT ≤ 15°C, Ambient ΔT ≤ 40°C (NETA 100.18)',
            reference: `T_ref: ${item.reference_component_temp_c}°C, T_amb: ${item.ambient_temperature_c}°C`,
            deviation: `ΔT_comp: +${item.delta_t_comp}°C | ΔT_amb: +${item.delta_t_amb}°C (${item.level_en})`,
            status: item.level === 4 ? 'FAIL' : item.level >= 2 ? 'INVESTIGATE' : 'PASS'
          }))}
          warnings={[
            liveStats.hasLevel4 ? `CẢNH BÁO LEVEL 4: Phát hiện điểm phát nhiệt nguy cấp tại [${liveStats.criticalItem?.location}] đạt ${liveStats.criticalItem?.measured_temp_c}°C (ΔT_comp = +${liveStats.criticalItem?.delta_t_comp}°C > 15°C). Nguy cơ hỏa hoạn và phóng hồ quang! BẮT BUỘC NGẮT CÔ LẬP SỬA CHỮA NGAY LẬP TỨC!` : '',
            liveStats.hasLevel3 ? `LƯU Ý LEVEL 3: Có điểm phát nhiệt tăng cao trên 21°C so với môi trường. Yêu cầu giám sát chặt chẽ liên tục.` : '',
            liveStats.hasLevel2 ? `KHUYẾT TẬT TIỀM ẨN LEVEL 2: Cần lên lịch siết lực lại bu-lông và vệ sinh bề mặt trong kỳ bảo dưỡng tới.` : ''
          ].filter(Boolean)}
          recommendations={[
            'Xử lý khẩn cấp điểm phát nhiệt Level 4: Ngắt điện cô lập, tháo rã mối nối bu-lông, vệ sinh bề mặt tiếp xúc, bôi mỡ dẫn điện chống oxy hóa và siết lực đạt chuẩn NETA MTS Table 100.12.',
            'Kiểm tra và siết lại toàn bộ cọc đấu dây các thiết bị lân cận thuộc cùng nhánh phân phối.',
            'Thực hiện chụp lại ảnh nhiệt kiểm chứng dưới tải tương đương sau khi hoàn thành sửa chữa; đảm bảo ΔT hạ về dưới 3°C trước khi ký biên bản nghiệm thu.'
          ]}
        />
      )}
    </div>
  );
};
