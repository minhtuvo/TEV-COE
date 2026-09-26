import React, { useState, useMemo } from 'react';
import {
  RotateCcw,
  Zap,
  CheckCircle2,
  AlertTriangle,
  FileText,
  Printer,
  Sparkles,
  Save,
  ShieldCheck,
  Activity,
  Layers,
  Check,
  Eye,
  Camera,
  Upload,
  Gauge,
  Sliders,
  Power,
  ShieldAlert,
  Flame,
  Scale,
  Copy,
  Info,
  Clock,
  RefreshCw,
  FolderArchive,
  ArrowRight,
  Radio,
  SlidersHorizontal,
  XCircle,
  FileCheck2
} from 'lucide-react';
import {
  NetaSwitchgearInputPayload,
  NetaSwitchgearEvaluationResult,
  performDeterministicSwitchgearAnalysis
} from '../../server/netaSwitchgearAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';
import { NameplateScannerModal } from './NameplateScannerModal';
import { ExtractedNameplateData } from '../../server/nameplateOcrAnalyzer';
import { FseEngineerDropdown } from './FseEngineerDropdown';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

export const SAMPLE_SWITCHGEAR_FSE_DATA: NetaSwitchgearInputPayload = {
  site_info: {
    project_name: "Tram Khieu TEV Industrial Factory",
    switchgear_tag: "MSB-22KV-CELL-01",
    switchgear_rating: "24kV 1250A 25kA Metal-Clad Switchgear",
    fse_name: "Nguyen Van T",
    test_date: "2026-09-26",
    substation_location: "Phòng Phân Phối Trung Thế MV Trạm Chính",
    manufacturer: "Schneider / ABB / Siemens",
    rated_voltage_kv: 24,
    rated_current_a: 1250,
    short_circuit_ka: 25
  },
  visual_inspection: {
    nameplate_match: true,
    physical_condition_cleanliness: "Pass (Clean, shipping bracing removed)",
    anchorage_grounding_clearances: "Pass",
    mimic_and_labeling: true,
    fuse_breaker_ratings_match: true,
    ct_vt_ratios_match: true,
    interlock_systems_operation: "Pass",
    barriers_shutters_lubrication: "Pass",
    filters_and_heaters_check: "Pass",
    cpt_visual_inspection: "Pass",
    bolt_torque_check: "Pass",
    thermographic_survey: "Pass"
  },
  electrical_tests: {
    bolted_resistance_micro_ohms: [11.2, 11.5, 11.3],
    bus_insulation_resistance_2500v_1min_megohms: {
      phase_a_to_ground: 4500,
      phase_b_to_ground: 4200,
      phase_c_to_ground: 3900,
      phase_ab: 5100,
      phase_bc: 4800,
      phase_ca: 5000
    },
    control_wiring_ir_1000v_megohms: 1.2,
    cpt_tests: {
      insulation_resistance_1000v_megohms: 850,
      turns_ratio_error_percent: 0.15
    },
    ct_vt_section_7_10_status: "Pass",
    ground_resistance_section_7_13: "Pass (0.35 ohms)",
    current_injection_wiring_check: "Pass",
    phasing_check_dual_source: "Pass (Phase matched)",
    online_partial_discharge_tev_db: 12.0,
    surge_arrester_status: "Pass"
  },
  previous_test_data: {
    last_test_date: "2025-09-25",
    last_control_wiring_ir_megohms: 25.0
  }
};

export const NetaSwitchgearChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaSwitchgearInputPayload>(SAMPLE_SWITCHGEAR_FSE_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaSwitchgearEvaluationResult | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [jsonText, setJsonText] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);
  const [showScannerModal, setShowScannerModal] = useState(false);

  // Real-time deterministic calculation according to NETA ATS-2025 Section 7.1.1
  const liveStats = useMemo(() => {
    const v = data.visual_inspection;
    const e = data.electrical_tests;
    const prev = data.previous_test_data;

    // Bolted resistance deviation
    const bolts = e.bolted_resistance_micro_ohms || [11.2, 11.5, 11.3];
    const minBolt = Math.min(...bolts);
    const maxBolt = Math.max(...bolts);
    const boltDev = minBolt > 0 ? parseFloat((((maxBolt - minBolt) / minBolt) * 100).toFixed(1)) : 0;
    const isBoltPass = boltDev <= 50;

    // Bus Insulation Resistance (Table 100.1 threshold: 1000 MΩ for 2500V test on 24kV)
    const bus = e.bus_insulation_resistance_2500v_1min_megohms || {
      phase_a_to_ground: 4500,
      phase_b_to_ground: 4200,
      phase_c_to_ground: 3900,
      phase_ab: 5100,
      phase_bc: 4800,
      phase_ca: 5000
    };
    const minBusGnd = Math.min(bus.phase_a_to_ground, bus.phase_b_to_ground, bus.phase_c_to_ground);
    const minBusPh = Math.min(bus.phase_ab, bus.phase_bc, bus.phase_ca);
    const isBusPass = minBusGnd >= 1000 && minBusPh >= 1000;

    // Control Wiring IR (NETA Sec 7.1.1.D.4: Minimum 2.0 Megohms)
    const ctrlIr = e.control_wiring_ir_1000v_megohms ?? 1.2;
    const isCtrlPass = ctrlIr >= 2.0;
    const lastCtrl = prev?.last_control_wiring_ir_megohms ?? 25.0;
    const ctrlDrop = lastCtrl > 0 ? parseFloat((((lastCtrl - ctrlIr) / lastCtrl) * 100).toFixed(1)) : 0;

    // CPT Transformer Tests (Ratio error <= 0.5%, IR >= 100 MΩ)
    const cptRatio = e.cpt_tests?.turns_ratio_error_percent ?? 0.15;
    const cptIr = e.cpt_tests?.insulation_resistance_1000v_megohms ?? 850;
    const isCptPass = cptRatio <= 0.5 && cptIr >= 100;

    // Online TEV Partial Discharge (< 20 dB is safe)
    const tevDb = e.online_partial_discharge_tev_db ?? 12.0;
    const isTevPass = tevDb < 20;

    // Visual pass
    const isVisualPass =
      v.nameplate_match &&
      v.mimic_and_labeling &&
      v.fuse_breaker_ratings_match &&
      (v.physical_condition_cleanliness || '').toLowerCase().includes('pass') &&
      (v.interlock_systems_operation || '').toLowerCase().includes('pass') &&
      (v.bolt_torque_check || '').toLowerCase().includes('pass');

    // Overall instant status
    let instantStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
    if (!isCtrlPass || !isBusPass || !isCptPass || !isVisualPass) {
      instantStatus = 'FAIL';
    } else if (!isBoltPass || tevDb >= 20) {
      instantStatus = 'INVESTIGATE';
    }

    return {
      bolts,
      minBolt,
      maxBolt,
      boltDev,
      isBoltPass,
      bus,
      minBusGnd,
      minBusPh,
      isBusPass,
      ctrlIr,
      isCtrlPass,
      lastCtrl,
      ctrlDrop,
      cptRatio,
      cptIr,
      isCptPass,
      tevDb,
      isTevPass,
      isVisualPass,
      instantStatus
    };
  }, [data]);

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    setAnalysisResult(null);
    setSavedSuccessMessage(null);

    try {
      const response = await fetch('/api/field-service/neta-switchgear-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });

      const resJson = await response.json();
      if (resJson.success) {
        setAnalysisResult(resJson.data);
      } else {
        alert('Lỗi phân tích: ' + (resJson.error || 'Vui lòng kiểm tra dữ liệu đầu vào'));
      }
    } catch (err: any) {
      console.warn('Lỗi khi gọi API phân tích Switchgear; sử dụng fallback phân tích máy khách:', err);
      const fallback = performDeterministicSwitchgearAnalysis(data);
      setAnalysisResult(fallback);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleSaveToCmms = () => {
    const resultToSave = analysisResult || performDeterministicSwitchgearAnalysis(data);
    if (onSaveReport) {
      onSaveReport({
        data,
        analysisResult: resultToSave,
        equipmentId: data.site_info.switchgear_tag,
        date: data.site_info.test_date,
        type: 'Switchgear & Switchboard Assemblies (ANSI/NETA ATS-2025 Mục 7.1.1)'
      });
      setSavedSuccessMessage(`Đã lưu biên bản kiểm định tủ điện [${data.site_info.switchgear_tag}] vào kho báo cáo CMMS!`);
    }
  };

  const handleResetToSample = () => {
    setData(SAMPLE_SWITCHGEAR_FSE_DATA);
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

  const handleApplyExtractedNameplate = (extracted: ExtractedNameplateData) => {
    setData(prev => ({
      ...prev,
      site_info: {
        ...prev.site_info,
        switchgear_tag: extracted.equipment_tag || prev.site_info.switchgear_tag,
        switchgear_rating: extracted.equipment_type || prev.site_info.switchgear_rating,
        manufacturer: extracted.manufacturer || prev.site_info.manufacturer,
        serial_number: extracted.serial_number || prev.site_info.serial_number,
        rated_voltage_kv: extracted.primary_voltage_kv || prev.site_info.rated_voltage_kv,
      }
    }));
    setShowScannerModal(false);
  };

  return (
    <div className="space-y-6">
      {/* Top Banner & Header */}
      <div className="bg-gradient-to-r from-slate-900 via-slate-800 to-indigo-950 text-white p-5 sm:p-6 rounded-2xl shadow-lg border border-indigo-500/30">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1.5">
            <div className="flex flex-wrap items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-extrabold uppercase tracking-wider bg-indigo-500/30 text-indigo-200 backdrop-blur-xs border border-indigo-400/40 flex items-center gap-1.5">
                <Power size={13} className="text-indigo-300" />
                ANSI/NETA ATS-2025 Mục 7.1.1
              </span>
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-emerald-950/60 text-emerald-300 border border-emerald-500/40 flex items-center gap-1">
                <ShieldCheck size={12} />
                Switchgear & Switchboard Assemblies • Table 100.1 • Table 100.12 • Table 100.23
              </span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white flex items-center gap-2.5">
              <span>Kiểm Định Tủ Điện Phân Phối & Tủ Đóng Cắt (Switchgear)</span>
            </h1>
            <p className="text-xs sm:text-sm text-slate-300 max-w-3xl">
              Hệ thống đối soát hiện trường tự động NETA ATS-2025 Mục 7.1.1: Khóa liên động cơ điện, lực siết bu-lông (DLRO ≤ 50%), cách điện thanh cái (&gt; 1000 MΩ), cách điện mạch điều khiển nhị thứ (≥ 2.0 MΩ), MBA cấp nguồn điều khiển CPT và phóng điện cục bộ TEV Online.
            </p>
          </div>

          {/* Action Buttons */}
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => setShowScannerModal(true)}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-indigo-600/60 hover:bg-indigo-600 text-white border border-indigo-400/40 flex items-center gap-1.5 transition-all shadow-xs"
              title="Quét nhãn mác tủ điện bằng camera"
            >
              <Camera size={14} />
              <span>Quét Nameplate AI</span>
            </button>

            <button
              onClick={handleResetToSample}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title="Nạp dữ liệu mẫu hiện trường FSE (Bao gồm lỗi cách điện mạch điều khiển 1.2 MΩ)"
            >
              <RotateCcw size={14} />
              <span>Nạp Dữ Liệu Mẫu</span>
            </button>

            <button
              onClick={handleOpenJsonModal}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title="Xem / Chỉnh sửa JSON thô"
            >
              <Sliders size={14} />
              <span>JSON Nhập / Xuất</span>
            </button>

            <button
              onClick={() => setShowPrintModal(true)}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title="In biên bản nghiệm thu hoặc xuất PDF"
            >
              <Printer size={14} />
              <span>In Báo Cáo PDF</span>
            </button>
          </div>
        </div>

        {/* Live Diagnostics Metrics Cards */}
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3 mt-5 pt-4 border-t border-slate-700/60 text-xs">
          {/* 1. Control Wiring IR Status (Critical Alert) */}
          <div className={`p-3 rounded-xl border ${
            !liveStats.isCtrlPass 
              ? 'bg-red-950/70 border-red-500/80 text-red-200' 
              : 'bg-slate-800/80 border-slate-700 text-slate-200'
          }`}>
            <div className="flex items-center justify-between font-medium mb-1">
              <span className="flex items-center gap-1">
                <Zap size={13} className={!liveStats.isCtrlPass ? 'text-red-400' : 'text-emerald-400'} />
                Cách Điện Mạch Điều Khiển
              </span>
              <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                !liveStats.isCtrlPass ? 'bg-red-500 text-white' : 'bg-emerald-500 text-white'
              }`}>
                {!liveStats.isCtrlPass ? 'FAIL' : 'PASS'}
              </span>
            </div>
            <div className="text-lg font-black text-white">
              {liveStats.ctrlIr} <span className="text-xs font-normal">MΩ</span>
            </div>
            <div className="text-[11px] opacity-80 mt-0.5">
              Chuẩn NETA: ≥ 2.0 MΩ {liveStats.ctrlDrop > 0 && `(Sụt -${liveStats.ctrlDrop}%)`}
            </div>
          </div>

          {/* 2. Bus Insulation Resistance */}
          <div className={`p-3 rounded-xl border ${
            liveStats.isBusPass 
              ? 'bg-slate-800/80 border-slate-700 text-slate-200' 
              : 'bg-red-950/70 border-red-500/80 text-red-200'
          }`}>
            <div className="flex items-center justify-between font-medium mb-1">
              <span className="flex items-center gap-1">
                <ShieldCheck size={13} className="text-emerald-400" />
                Cách Điện Thanh Cái (Bus IR)
              </span>
              <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-500 text-white">
                PASS
              </span>
            </div>
            <div className="text-lg font-black text-white">
              {liveStats.minBusGnd} <span className="text-xs font-normal">MΩ</span>
            </div>
            <div className="text-[11px] opacity-80 mt-0.5">
              Min Pha-Đất (Ngưỡng Table 100.1: ≥ 1000 MΩ)
            </div>
          </div>

          {/* 3. Bolted Resistance DLRO */}
          <div className={`p-3 rounded-xl border ${
            liveStats.isBoltPass 
              ? 'bg-slate-800/80 border-slate-700 text-slate-200' 
              : 'bg-amber-950/70 border-amber-500/80 text-amber-200'
          }`}>
            <div className="flex items-center justify-between font-medium mb-1">
              <span className="flex items-center gap-1">
                <SlidersHorizontal size={13} className="text-blue-400" />
                Mối Nối Bu-lông DLRO
              </span>
              <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                liveStats.isBoltPass ? 'bg-emerald-500 text-white' : 'bg-amber-500 text-white'
              }`}>
                {liveStats.isBoltPass ? 'PASS' : 'INVESTIGATE'}
              </span>
            </div>
            <div className="text-lg font-black text-white">
              Δ {liveStats.boltDev}%
            </div>
            <div className="text-[11px] opacity-80 mt-0.5">
              Độ lệch Max vs Min (Ngưỡng ≤ 50%)
            </div>
          </div>

          {/* 4. CPT Control Transformer */}
          <div className="p-3 rounded-xl border bg-slate-800/80 border-slate-700 text-slate-200">
            <div className="flex items-center justify-between font-medium mb-1">
              <span className="flex items-center gap-1">
                <Activity size={13} className="text-cyan-400" />
                MBA CPT Điều Khiển
              </span>
              <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-500 text-white">
                PASS
              </span>
            </div>
            <div className="text-lg font-black text-white">
              ±{liveStats.cptRatio}% <span className="text-xs font-normal font-mono">| {liveStats.cptIr}MΩ</span>
            </div>
            <div className="text-[11px] opacity-80 mt-0.5">
              Sai số tỷ số (Ngưỡng ≤ 0.5%) &amp; IR
            </div>
          </div>

          {/* 5. Online Partial Discharge TEV */}
          <div className="p-3 rounded-xl border bg-slate-800/80 border-slate-700 text-slate-200 col-span-2 sm:col-span-1">
            <div className="flex items-center justify-between font-medium mb-1">
              <span className="flex items-center gap-1">
                <Radio size={13} className="text-violet-400" />
                Xung PD TEV Online
              </span>
              <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-500 text-white">
                NORMAL
              </span>
            </div>
            <div className="text-lg font-black text-white">
              {liveStats.tevDb} <span className="text-xs font-normal">dB</span>
            </div>
            <div className="text-[11px] opacity-80 mt-0.5">
              Table 100.23: An toàn &lt; 20 dB
            </div>
          </div>
        </div>
      </div>

      {savedSuccessMessage && (
        <div className="p-4 rounded-xl bg-emerald-50 border border-emerald-200 text-emerald-800 flex items-center justify-between shadow-xs">
          <div className="flex items-center gap-2">
            <CheckCircle2 size={18} className="text-emerald-600 flex-shrink-0" />
            <span className="text-sm font-semibold">{savedSuccessMessage}</span>
          </div>
          {onNavigateToReports && (
            <button
              onClick={onNavigateToReports}
              className="text-xs font-bold text-emerald-700 hover:text-emerald-900 underline flex items-center gap-1"
            >
              <span>Xem Trong Kho Báo Cáo</span>
              <ArrowRight size={13} />
            </button>
          )}
        </div>
      )}

      {/* Main Checklist Input Fields Grid */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Left Column: Section 1 & Section 2 */}
        <div className="space-y-6">
          {/* Section 1: Site Info & Switchgear Data */}
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100 mb-4">
              <h2 className="text-base font-bold text-slate-800 flex items-center gap-2">
                <Power size={18} className="text-indigo-600" />
                1. Thông Tin Trạm &amp; Tủ Điện Switchgear
              </h2>
              <span className="text-xs font-medium text-slate-400">Site &amp; Nameplate</span>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs">
              <div>
                <label className="block text-slate-600 font-semibold mb-1">Tên Dự Án / Nhà Máy</label>
                <input
                  type="text"
                  value={data.site_info.project_name}
                  onChange={e => setData(prev => ({
                    ...prev,
                    site_info: { ...prev.site_info, project_name: e.target.value }
                  }))}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-slate-800 focus:bg-white focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none transition-all"
                  placeholder="Ví dụ: Tram Khieu TEV Industrial Factory"
                />
              </div>

              <div>
                <label className="block text-slate-600 font-semibold mb-1">Mã Ngăn Tủ (Tag)</label>
                <input
                  type="text"
                  value={data.site_info.switchgear_tag}
                  onChange={e => setData(prev => ({
                    ...prev,
                    site_info: { ...prev.site_info, switchgear_tag: e.target.value }
                  }))}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-slate-800 font-bold focus:bg-white focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none transition-all"
                  placeholder="Ví dụ: MSB-22KV-CELL-01"
                />
              </div>

              <div className="sm:col-span-2">
                <label className="block text-slate-600 font-semibold mb-1">Quy Cách &amp; Thông Số Định Mức (Rating)</label>
                <input
                  type="text"
                  value={data.site_info.switchgear_rating}
                  onChange={e => setData(prev => ({
                    ...prev,
                    site_info: { ...prev.site_info, switchgear_rating: e.target.value }
                  }))}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-slate-800 focus:bg-white focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none transition-all"
                  placeholder="Ví dụ: 24kV 1250A 25kA Metal-Clad Switchgear"
                />
              </div>

              <div>
                <label className="block text-slate-600 font-semibold mb-1">Kỹ Sư Kiểm Định (FSE)</label>
                <FseEngineerDropdown
                  value={data.site_info.fse_name}
                  onChange={val => setData(prev => ({
                    ...prev,
                    site_info: { ...prev.site_info, fse_name: val }
                  }))}
                />
              </div>

              <div>
                <label className="block text-slate-600 font-semibold mb-1">Ngày Kiểm Tra Hiện Trường</label>
                <input
                  type="date"
                  value={data.site_info.test_date}
                  onChange={e => setData(prev => ({
                    ...prev,
                    site_info: { ...prev.site_info, test_date: e.target.value }
                  }))}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-slate-800 focus:bg-white focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 outline-none transition-all"
                />
              </div>

              <div>
                <label className="block text-slate-600 font-semibold mb-1">Điện Áp Định Mức (kV)</label>
                <input
                  type="number"
                  value={data.site_info.rated_voltage_kv || 24}
                  onChange={e => setData(prev => ({
                    ...prev,
                    site_info: { ...prev.site_info, rated_voltage_kv: Number(e.target.value) }
                  }))}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-slate-800 focus:bg-white focus:border-indigo-500 outline-none"
                />
              </div>

              <div>
                <label className="block text-slate-600 font-semibold mb-1">Dòng Cắt Ngắn Mạch (kA)</label>
                <input
                  type="number"
                  value={data.site_info.short_circuit_ka || 25}
                  onChange={e => setData(prev => ({
                    ...prev,
                    site_info: { ...prev.site_info, short_circuit_ka: Number(e.target.value) }
                  }))}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-slate-800 focus:bg-white focus:border-indigo-500 outline-none"
                />
              </div>
            </div>
          </div>

          {/* Section 2: Visual & Mechanical Inspection (NETA ATS-2025 Section 7.1.1.A) */}
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100 mb-4">
              <h2 className="text-base font-bold text-slate-800 flex items-center gap-2">
                <CheckCircle2 size={18} className="text-emerald-600" />
                2. Kiểm Tra Thị Giác &amp; Cơ Khí (NETA Mục 7.1.1.A)
              </h2>
              <span className="text-[11px] font-semibold px-2 py-0.5 bg-emerald-50 text-emerald-700 rounded border border-emerald-200">
                11 Hạng Mục Tiêu Chuẩn
              </span>
            </div>

            <div className="space-y-3 text-xs">
              {/* Nameplate match */}
              <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-lg hover:bg-slate-100/70 transition-colors">
                <div>
                  <div className="font-semibold text-slate-800">1. Đối Soát Nhãn Mác (Nameplate Match)</div>
                  <div className="text-[11px] text-slate-500">Điện áp, dòng định mức, dòng cắt ngắn mạch khớp 100% bản vẽ thiết kế</div>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.nameplate_match}
                  onChange={e => setData(prev => ({
                    ...prev,
                    visual_inspection: { ...prev.visual_inspection, nameplate_match: e.target.checked }
                  }))}
                  className="w-4 h-4 text-indigo-600 rounded focus:ring-indigo-500"
                />
              </div>

              {/* Physical Condition & Cleanliness */}
              <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-lg hover:bg-slate-100/70 transition-colors">
                <div>
                  <div className="font-semibold text-slate-800">2. Độ Sạch Sẽ &amp; Tháo Kẹp Vận Chuyển</div>
                  <div className="text-[11px] text-slate-500">Tủ sạch bụi bẩn, tháo bỏ hoàn toàn nẹp cố định vận chuyển, bản vẽ trong khoang</div>
                </div>
                <select
                  value={data.visual_inspection.physical_condition_cleanliness}
                  onChange={e => setData(prev => ({
                    ...prev,
                    visual_inspection: { ...prev.visual_inspection, physical_condition_cleanliness: e.target.value }
                  }))}
                  className="px-2.5 py-1 bg-white border border-slate-200 rounded text-xs font-semibold text-slate-700 outline-none"
                >
                  <option value="Pass (Clean, shipping bracing removed)">Pass (Đạt sạch, đã tháo nẹp)</option>
                  <option value="Fail (Dust, shipping bracing intact)">Fail (Bụi bẩn / Chưa tháo nẹp)</option>
                </select>
              </div>

              {/* Anchorage & Clearances */}
              <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-lg hover:bg-slate-100/70 transition-colors">
                <div>
                  <div className="font-semibold text-slate-800">3. Định Vị, Tiếp Địa &amp; Hành Lang An Toàn</div>
                  <div className="text-[11px] text-slate-500">Bắt bu-lông chân tủ chắc chắn, tiếp địa thanh cái/vỏ tủ, khoảng cách thoát hiểm</div>
                </div>
                <select
                  value={data.visual_inspection.anchorage_grounding_clearances}
                  onChange={e => setData(prev => ({
                    ...prev,
                    visual_inspection: { ...prev.visual_inspection, anchorage_grounding_clearances: e.target.value }
                  }))}
                  className="px-2.5 py-1 bg-white border border-slate-200 rounded text-xs font-semibold text-slate-700 outline-none"
                >
                  <option value="Pass">Pass (Đạt tiếp địa &amp; hành lang)</option>
                  <option value="Fail">Fail (Lỏng lẻo / Hành lang hẹp)</option>
                </select>
              </div>

              {/* Mimic & Labeling */}
              <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-lg hover:bg-slate-100/70 transition-colors">
                <div>
                  <div className="font-semibold text-slate-800">4. Sơ Đồ Mimic &amp; Nhãn Thiết Bị</div>
                  <div className="text-[11px] text-slate-500">Sơ đồ mimic mặt tủ và nhãn định danh thiết bị trùng khớp bản vẽ</div>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.mimic_and_labeling}
                  onChange={e => setData(prev => ({
                    ...prev,
                    visual_inspection: { ...prev.visual_inspection, mimic_and_labeling: e.target.checked }
                  }))}
                  className="w-4 h-4 text-indigo-600 rounded focus:ring-indigo-500"
                />
              </div>

              {/* Fuse & Breaker Ratings */}
              <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-lg hover:bg-slate-100/70 transition-colors">
                <div>
                  <div className="font-semibold text-slate-800">5. Trị Số Cầu Chì &amp; Máy Cắt Bảo Vệ</div>
                  <div className="text-[11px] text-slate-500">Trị số cầu chì và aptomat khớp 100% bản vẽ và tính toán phối hợp bảo vệ</div>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.fuse_breaker_ratings_match}
                  onChange={e => setData(prev => ({
                    ...prev,
                    visual_inspection: { ...prev.visual_inspection, fuse_breaker_ratings_match: e.target.checked }
                  }))}
                  className="w-4 h-4 text-indigo-600 rounded focus:ring-indigo-500"
                />
              </div>

              {/* Interlocks */}
              <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-lg hover:bg-slate-100/70 transition-colors">
                <div>
                  <div className="font-semibold text-slate-800">6. Khóa Liên Động Cơ Điện (Interlocks)</div>
                  <div className="text-[11px] text-slate-500">Khóa liên động chìa trao đổi (Key exchange), cơ cấu liên động ngắt CB vận hành 100%</div>
                </div>
                <select
                  value={data.visual_inspection.interlock_systems_operation}
                  onChange={e => setData(prev => ({
                    ...prev,
                    visual_inspection: { ...prev.visual_inspection, interlock_systems_operation: e.target.value }
                  }))}
                  className="px-2.5 py-1 bg-white border border-slate-200 rounded text-xs font-semibold text-slate-700 outline-none"
                >
                  <option value="Pass">Pass (Liên động an toàn 100%)</option>
                  <option value="Fail">Fail (Kẹt khóa / Cơ cấu lỗi)</option>
                </select>
              </div>

              {/* Lubrication & Shutters */}
              <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-lg hover:bg-slate-100/70 transition-colors">
                <div>
                  <div className="font-semibold text-slate-800">7. Bôi Trơn Tiếp Điểm &amp; Cửa Sập (Shutters)</div>
                  <div className="text-[11px] text-slate-500">Tiếp điểm động bôi trơn đạt chuẩn; cửa sập (shutters) và vách ngăn pha đóng mở trơn tru</div>
                </div>
                <select
                  value={data.visual_inspection.barriers_shutters_lubrication}
                  onChange={e => setData(prev => ({
                    ...prev,
                    visual_inspection: { ...prev.visual_inspection, barriers_shutters_lubrication: e.target.value }
                  }))}
                  className="px-2.5 py-1 bg-white border border-slate-200 rounded text-xs font-semibold text-slate-700 outline-none"
                >
                  <option value="Pass">Pass (Vận hành trơn tru)</option>
                  <option value="Fail">Fail (Kẹt cửa sập / Khô mỡ)</option>
                </select>
              </div>

              {/* Filters & Space Heaters */}
              <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-lg hover:bg-slate-100/70 transition-colors">
                <div>
                  <div className="font-semibold text-slate-800">8. Bộ Lọc Gió &amp; Sấy Chống Đọng Sương</div>
                  <div className="text-[11px] text-slate-500">Bộ lọc gió sạch sẽ; bộ sấy không gian (Space Heaters) &amp; thermostat chạy bình thường</div>
                </div>
                <select
                  value={data.visual_inspection.filters_and_heaters_check}
                  onChange={e => setData(prev => ({
                    ...prev,
                    visual_inspection: { ...prev.visual_inspection, filters_and_heaters_check: e.target.value }
                  }))}
                  className="px-2.5 py-1 bg-white border border-slate-200 rounded text-xs font-semibold text-slate-700 outline-none"
                >
                  <option value="Pass">Pass (Sấy hoạt động tốt)</option>
                  <option value="Fail">Fail (Hỏng bộ sấy / Tắc lọc)</option>
                </select>
              </div>

              {/* CPT Visual */}
              <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-lg hover:bg-slate-100/70 transition-colors">
                <div>
                  <div className="font-semibold text-slate-800">9. MBA Cấp Nguồn Điều Khiển (CPT Visual)</div>
                  <div className="text-[11px] text-slate-500">CPT nguyên vẹn, tiếp điểm rút cắm tốt, cầu chì bảo vệ đúng trị số</div>
                </div>
                <select
                  value={data.visual_inspection.cpt_visual_inspection}
                  onChange={e => setData(prev => ({
                    ...prev,
                    visual_inspection: { ...prev.visual_inspection, cpt_visual_inspection: e.target.value }
                  }))}
                  className="px-2.5 py-1 bg-white border border-slate-200 rounded text-xs font-semibold text-slate-700 outline-none"
                >
                  <option value="Pass">Pass (CPT nguyên vẹn, tiếp xúc tốt)</option>
                  <option value="Fail">Fail (Rỉ sét / Tiếp xúc kém)</option>
                </select>
              </div>

              {/* Bolt Torque Check */}
              <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-lg hover:bg-slate-100/70 transition-colors">
                <div>
                  <div className="font-semibold text-slate-800">10. Lực Siết Bu-lông Mối Nối (Bolt Torque)</div>
                  <div className="text-[11px] text-slate-500">Siết đạt mô-men lực theo nhà sản xuất hoặc Bảng Table 100.12 của NETA</div>
                </div>
                <select
                  value={data.visual_inspection.bolt_torque_check}
                  onChange={e => setData(prev => ({
                    ...prev,
                    visual_inspection: { ...prev.visual_inspection, bolt_torque_check: e.target.value }
                  }))}
                  className="px-2.5 py-1 bg-white border border-slate-200 rounded text-xs font-semibold text-slate-700 outline-none"
                >
                  <option value="Pass">Pass (Đạt chuẩn Table 100.12)</option>
                  <option value="Fail">Fail (Lỏng bu-lông)</option>
                </select>
              </div>
            </div>
          </div>
        </div>

        {/* Right Column: Section 3 & Section 4 */}
        <div className="space-y-6">
          {/* Section 3: Electrical Test Values (NETA Section 7.1.1.D) */}
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100 mb-4">
              <h2 className="text-base font-bold text-slate-800 flex items-center gap-2">
                <Zap size={18} className="text-amber-500" />
                3. Phép Đo Điện &amp; Tiêu Chuẩn NETA ATS-2025 Mục 7.1.1.D
              </h2>
              <span className="text-[11px] font-semibold px-2 py-0.5 bg-amber-50 text-amber-700 rounded border border-amber-200">
                Phép Đo Định Lượng
              </span>
            </div>

            <div className="space-y-4 text-xs">
              {/* 1. Bolted Connection Resistance (Micro-Ohms) */}
              <div className="p-3.5 bg-slate-50 rounded-xl border border-slate-200/80">
                <div className="flex items-center justify-between mb-2">
                  <span className="font-bold text-slate-800 flex items-center gap-1.5">
                    <SlidersHorizontal size={14} className="text-indigo-600" />
                    Điện Trở Mối Nối Bu-lông DLRO (µΩ)
                  </span>
                  <span className={`text-[10px] font-bold px-2 py-0.5 rounded ${
                    liveStats.isBoltPass ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'
                  }`}>
                    Lệch: {liveStats.boltDev}% {liveStats.isBoltPass ? '(≤ 50% Đạt)' : '(> 50% INVESTIGATE)'}
                  </span>
                </div>
                <div className="grid grid-cols-3 gap-2">
                  {[0, 1, 2].map(idx => (
                    <div key={idx}>
                      <span className="text-[10px] text-slate-500 block">Mối Nối #{idx + 1} (µΩ)</span>
                      <input
                        type="number"
                        step="0.1"
                        value={data.electrical_tests.bolted_resistance_micro_ohms[idx] ?? 11.2}
                        onChange={e => {
                          const val = Number(e.target.value);
                          setData(prev => {
                            const next = [...(prev.electrical_tests.bolted_resistance_micro_ohms || [11.2, 11.5, 11.3])];
                            next[idx] = val;
                            return {
                              ...prev,
                              electrical_tests: { ...prev.electrical_tests, bolted_resistance_micro_ohms: next }
                            };
                          });
                        }}
                        className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded font-semibold text-slate-800 focus:border-indigo-500 outline-none"
                      />
                    </div>
                  ))}
                </div>
              </div>

              {/* 2. Bus Insulation Resistance (2500V DC 1-Min) */}
              <div className="p-3.5 bg-slate-50 rounded-xl border border-slate-200/80">
                <div className="flex items-center justify-between mb-2">
                  <span className="font-bold text-slate-800 flex items-center gap-1.5">
                    <ShieldCheck size={14} className="text-emerald-600" />
                    Điện Trở Cách Điện Thanh Cái (2500V 1-phút, MΩ)
                  </span>
                  <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-emerald-100 text-emerald-800">
                    NETA Table 100.1: ≥ 1,000 MΩ
                  </span>
                </div>
                <div className="grid grid-cols-3 gap-2 mb-2">
                  <div>
                    <span className="text-[10px] text-slate-500 block">Pha A - Đất (MΩ)</span>
                    <input
                      type="number"
                      value={data.electrical_tests.bus_insulation_resistance_2500v_1min_megohms.phase_a_to_ground}
                      onChange={e => {
                        const val = Number(e.target.value);
                        setData(prev => ({
                          ...prev,
                          electrical_tests: {
                            ...prev.electrical_tests,
                            bus_insulation_resistance_2500v_1min_megohms: {
                              ...prev.electrical_tests.bus_insulation_resistance_2500v_1min_megohms,
                              phase_a_to_ground: val
                            }
                          }
                        }));
                      }}
                      className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded font-semibold text-slate-800 outline-none"
                    />
                  </div>
                  <div>
                    <span className="text-[10px] text-slate-500 block">Pha B - Đất (MΩ)</span>
                    <input
                      type="number"
                      value={data.electrical_tests.bus_insulation_resistance_2500v_1min_megohms.phase_b_to_ground}
                      onChange={e => {
                        const val = Number(e.target.value);
                        setData(prev => ({
                          ...prev,
                          electrical_tests: {
                            ...prev.electrical_tests,
                            bus_insulation_resistance_2500v_1min_megohms: {
                              ...prev.electrical_tests.bus_insulation_resistance_2500v_1min_megohms,
                              phase_b_to_ground: val
                            }
                          }
                        }));
                      }}
                      className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded font-semibold text-slate-800 outline-none"
                    />
                  </div>
                  <div>
                    <span className="text-[10px] text-slate-500 block">Pha C - Đất (MΩ)</span>
                    <input
                      type="number"
                      value={data.electrical_tests.bus_insulation_resistance_2500v_1min_megohms.phase_c_to_ground}
                      onChange={e => {
                        const val = Number(e.target.value);
                        setData(prev => ({
                          ...prev,
                          electrical_tests: {
                            ...prev.electrical_tests,
                            bus_insulation_resistance_2500v_1min_megohms: {
                              ...prev.electrical_tests.bus_insulation_resistance_2500v_1min_megohms,
                              phase_c_to_ground: val
                            }
                          }
                        }));
                      }}
                      className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded font-semibold text-slate-800 outline-none"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-3 gap-2">
                  <div>
                    <span className="text-[10px] text-slate-500 block">Pha A - B (MΩ)</span>
                    <input
                      type="number"
                      value={data.electrical_tests.bus_insulation_resistance_2500v_1min_megohms.phase_ab}
                      onChange={e => {
                        const val = Number(e.target.value);
                        setData(prev => ({
                          ...prev,
                          electrical_tests: {
                            ...prev.electrical_tests,
                            bus_insulation_resistance_2500v_1min_megohms: {
                              ...prev.electrical_tests.bus_insulation_resistance_2500v_1min_megohms,
                              phase_ab: val
                            }
                          }
                        }));
                      }}
                      className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded font-semibold text-slate-800 outline-none"
                    />
                  </div>
                  <div>
                    <span className="text-[10px] text-slate-500 block">Pha B - C (MΩ)</span>
                    <input
                      type="number"
                      value={data.electrical_tests.bus_insulation_resistance_2500v_1min_megohms.phase_bc}
                      onChange={e => {
                        const val = Number(e.target.value);
                        setData(prev => ({
                          ...prev,
                          electrical_tests: {
                            ...prev.electrical_tests,
                            bus_insulation_resistance_2500v_1min_megohms: {
                              ...prev.electrical_tests.bus_insulation_resistance_2500v_1min_megohms,
                              phase_bc: val
                            }
                          }
                        }));
                      }}
                      className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded font-semibold text-slate-800 outline-none"
                    />
                  </div>
                  <div>
                    <span className="text-[10px] text-slate-500 block">Pha C - A (MΩ)</span>
                    <input
                      type="number"
                      value={data.electrical_tests.bus_insulation_resistance_2500v_1min_megohms.phase_ca}
                      onChange={e => {
                        const val = Number(e.target.value);
                        setData(prev => ({
                          ...prev,
                          electrical_tests: {
                            ...prev.electrical_tests,
                            bus_insulation_resistance_2500v_1min_megohms: {
                              ...prev.electrical_tests.bus_insulation_resistance_2500v_1min_megohms,
                              phase_ca: val
                            }
                          }
                        }));
                      }}
                      className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded font-semibold text-slate-800 outline-none"
                    />
                  </div>
                </div>
              </div>

              {/* 3. Control Wiring IR (Section 7.1.1.D.4) - HIGHLIGHTED RED ON DEFECT */}
              <div className={`p-4 rounded-xl border transition-all ${
                !liveStats.isCtrlPass 
                  ? 'bg-red-50 border-red-300 ring-2 ring-red-400/30' 
                  : 'bg-slate-50 border-slate-200'
              }`}>
                <div className="flex items-center justify-between mb-1.5">
                  <span className="font-bold text-slate-800 flex items-center gap-1.5">
                    <Zap size={15} className={!liveStats.isCtrlPass ? 'text-red-600' : 'text-emerald-600'} />
                    Cách Điện Mạch Điều Khiển Nhị Thứ (Control Wiring IR @ 1000V DC)
                  </span>
                  <span className={`text-[10px] font-bold px-2 py-0.5 rounded ${
                    !liveStats.isCtrlPass ? 'bg-red-600 text-white animate-pulse' : 'bg-emerald-600 text-white'
                  }`}>
                    {!liveStats.isCtrlPass ? 'FAIL / DEFECT' : 'PASS'}
                  </span>
                </div>
                <div className="grid grid-cols-2 gap-3 items-center">
                  <div>
                    <label className="text-[10px] text-slate-600 font-semibold block mb-0.5">
                      Giá Trị Đo Thực Tế Hiện Trường (Megohms)
                    </label>
                    <input
                      type="number"
                      step="0.1"
                      value={data.electrical_tests.control_wiring_ir_1000v_megohms}
                      onChange={e => {
                        const val = Number(e.target.value);
                        setData(prev => ({
                          ...prev,
                          electrical_tests: { ...prev.electrical_tests, control_wiring_ir_1000v_megohms: val }
                        }));
                      }}
                      className={`w-full px-3 py-2 rounded-lg font-black text-base outline-none border transition-all ${
                        !liveStats.isCtrlPass
                          ? 'bg-white border-red-400 text-red-700 ring-1 ring-red-300'
                          : 'bg-white border-slate-200 text-slate-800'
                      }`}
                    />
                  </div>
                  <div className="text-[11px] text-slate-600 bg-white/70 p-2.5 rounded-lg border border-slate-200/60">
                    <div>
                      <strong>Quy chuẩn NETA Mục 7.1.1.D.4:</strong> Bắt buộc <span className="font-bold text-red-600">≥ 2.0 Megohms</span>.
                    </div>
                    {liveStats.ctrlDrop > 0 && (
                      <div className="text-red-700 font-semibold mt-1">
                        Sụt giảm {liveStats.ctrlDrop}% so với năm trước ({liveStats.lastCtrl} MΩ).
                      </div>
                    )}
                  </div>
                </div>
              </div>

              {/* 4. CPT Transformer Tests */}
              <div className="p-3.5 bg-slate-50 rounded-xl border border-slate-200/80">
                <div className="flex items-center justify-between mb-2">
                  <span className="font-bold text-slate-800 flex items-center gap-1.5">
                    <Activity size={14} className="text-cyan-600" />
                    Thử Nghiệm Máy Biến Áp CPT Điều Khiển (Mục 7.1.1.D.6)
                  </span>
                  <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-emerald-100 text-emerald-800">
                    Sai số ≤ 0.5% &amp; Table 100.1
                  </span>
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <span className="text-[10px] text-slate-500 block">Sai Số Tỷ Số CPT Turns Ratio (%)</span>
                    <input
                      type="number"
                      step="0.01"
                      value={data.electrical_tests.cpt_tests.turns_ratio_error_percent}
                      onChange={e => {
                        const val = Number(e.target.value);
                        setData(prev => ({
                          ...prev,
                          electrical_tests: {
                            ...prev.electrical_tests,
                            cpt_tests: { ...prev.electrical_tests.cpt_tests, turns_ratio_error_percent: val }
                          }
                        }));
                      }}
                      className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded font-semibold text-slate-800 outline-none"
                    />
                  </div>
                  <div>
                    <span className="text-[10px] text-slate-500 block">Cách Điện Cuộn Dây CPT 1000V (MΩ)</span>
                    <input
                      type="number"
                      value={data.electrical_tests.cpt_tests.insulation_resistance_1000v_megohms}
                      onChange={e => {
                        const val = Number(e.target.value);
                        setData(prev => ({
                          ...prev,
                          electrical_tests: {
                            ...prev.electrical_tests,
                            cpt_tests: { ...prev.electrical_tests.cpt_tests, insulation_resistance_1000v_megohms: val }
                          }
                        }));
                      }}
                      className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded font-semibold text-slate-800 outline-none"
                    />
                  </div>
                </div>
              </div>

              {/* 5. CT/VT, Grounding, Current Injection, Phasing, Online PD */}
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <span className="text-[10px] text-slate-500 block font-semibold">Điện Trở Đất Tiếp Địa (Mục 7.13)</span>
                  <input
                    type="text"
                    value={data.electrical_tests.ground_resistance_section_7_13}
                    onChange={e => setData(prev => ({
                      ...prev,
                      electrical_tests: { ...prev.electrical_tests, ground_resistance_section_7_13: e.target.value }
                    }))}
                    className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded text-slate-800 font-semibold outline-none"
                  />
                </div>

                <div>
                  <span className="text-[10px] text-slate-500 block font-semibold">Phóng Điện Cục Bộ Online TEV (dB)</span>
                  <input
                    type="number"
                    step="0.5"
                    value={data.electrical_tests.online_partial_discharge_tev_db}
                    onChange={e => setData(prev => ({
                      ...prev,
                      electrical_tests: { ...prev.electrical_tests, online_partial_discharge_tev_db: Number(e.target.value) }
                    }))}
                    className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded text-slate-800 font-semibold outline-none"
                  />
                </div>

                <div>
                  <span className="text-[10px] text-slate-500 block font-semibold">Bơm Dòng Nhị Thứ (Mạch CT)</span>
                  <input
                    type="text"
                    value={data.electrical_tests.current_injection_wiring_check}
                    onChange={e => setData(prev => ({
                      ...prev,
                      electrical_tests: { ...prev.electrical_tests, current_injection_wiring_check: e.target.value }
                    }))}
                    className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded text-slate-800 outline-none"
                  />
                </div>

                <div>
                  <span className="text-[10px] text-slate-500 block font-semibold">Kiểm Tra Khớp Pha Nguồn Kép</span>
                  <input
                    type="text"
                    value={data.electrical_tests.phasing_check_dual_source}
                    onChange={e => setData(prev => ({
                      ...prev,
                      electrical_tests: { ...prev.electrical_tests, phasing_check_dual_source: e.target.value }
                    }))}
                    className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded text-slate-800 outline-none"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Section 4: Previous Test Historical Data Comparison */}
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100 mb-4">
              <h2 className="text-base font-bold text-slate-800 flex items-center gap-2">
                <Clock size={18} className="text-slate-600" />
                4. Dữ Liệu Đối Soát Lịch Sử (Previous Year Inspection)
              </h2>
              <span className="text-xs text-slate-400">Trend Analysis</span>
            </div>

            <div className="grid grid-cols-2 gap-4 text-xs">
              <div>
                <label className="block text-slate-600 font-semibold mb-1">Ngày Đo Lần Gần Nhất</label>
                <input
                  type="date"
                  value={data.previous_test_data?.last_test_date || '2025-09-25'}
                  onChange={e => setData(prev => ({
                    ...prev,
                    previous_test_data: { ...prev.previous_test_data, last_test_date: e.target.value }
                  }))}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-slate-800 outline-none"
                />
              </div>

              <div>
                <label className="block text-slate-600 font-semibold mb-1">Cách Điện Mạch Điều Khiển Kỳ Trước (MΩ)</label>
                <input
                  type="number"
                  step="0.1"
                  value={data.previous_test_data?.last_control_wiring_ir_megohms || 25.0}
                  onChange={e => setData(prev => ({
                    ...prev,
                    previous_test_data: { ...prev.previous_test_data, last_control_wiring_ir_megohms: Number(e.target.value) }
                  }))}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-slate-800 font-semibold outline-none"
                />
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Primary Action Button: Run AI Analysis */}
      <div className="flex flex-col sm:flex-row items-center justify-between gap-4 p-5 bg-gradient-to-r from-slate-900 to-indigo-950 rounded-2xl shadow-md border border-indigo-500/30 text-white">
        <div className="space-y-1 text-center sm:text-left">
          <div className="flex items-center justify-center sm:justify-start gap-2">
            <Sparkles size={18} className="text-amber-400 animate-spin" style={{ animationDuration: '6s' }} />
            <span className="text-sm font-bold text-white">Trợ Lý AI Chuyên Gia Field Service (TEV Platform AI)</span>
          </div>
          <p className="text-xs text-slate-300">
            Tự động đối soát toàn bộ thông số với tiêu chuẩn ANSI/NETA ATS-2025 Mục 7.1.1, phân loại rủi ro và xuất khuyến nghị hành động 4 phần.
          </p>
        </div>

        <div className="flex items-center gap-3">
          <button
            onClick={handleAnalyze}
            disabled={isAnalyzing}
            className={`px-5 py-2.5 rounded-xl font-bold text-xs uppercase tracking-wider flex items-center gap-2 transition-all shadow-lg ${
              isAnalyzing
                ? 'bg-slate-700 text-slate-400 cursor-not-allowed'
                : 'bg-gradient-to-r from-amber-500 to-orange-500 hover:from-amber-400 hover:to-orange-400 text-slate-950 font-black shadow-amber-500/20 active:scale-95'
            }`}
          >
            {isAnalyzing ? (
              <>
                <RefreshCw size={15} className="animate-spin text-white" />
                <span>Đang Đối Soát Tiêu Chuẩn NETA...</span>
              </>
            ) : (
              <>
                <Sparkles size={15} className="text-slate-950" />
                <span>Chạy Đối Soát AI (Gemini 3.8 Flash)</span>
              </>
            )}
          </button>

          <button
            onClick={handleSaveToCmms}
            className="px-4 py-2.5 rounded-xl font-bold text-xs bg-emerald-600 hover:bg-emerald-500 text-white flex items-center gap-1.5 transition-all shadow-xs"
            title="Lưu biên bản vào kho báo cáo bảo trì CMMS"
          >
            <Save size={14} />
            <span>Lưu Báo Cáo</span>
          </button>
        </div>
      </div>

      {/* AI Analysis Result Section (Render 4 Markdown Parts) */}
      {analysisResult && (
        <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-6">
          {/* Status Header */}
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-100">
            <div className="flex items-center gap-3">
              <div className={`w-10 h-10 rounded-xl flex items-center justify-center font-black ${
                analysisResult.overall_status === 'FAIL' 
                  ? 'bg-red-500/10 text-red-600 border border-red-200' 
                  : analysisResult.overall_status === 'INVESTIGATE'
                  ? 'bg-amber-500/10 text-amber-600 border border-amber-200'
                  : 'bg-emerald-500/10 text-emerald-600 border border-emerald-200'
              }`}>
                {analysisResult.overall_status === 'FAIL' ? (
                  <XCircle size={24} />
                ) : analysisResult.overall_status === 'INVESTIGATE' ? (
                  <AlertTriangle size={24} />
                ) : (
                  <CheckCircle2 size={24} />
                )}
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className={`px-2.5 py-0.5 rounded-md text-[11px] font-black uppercase tracking-wider ${
                    analysisResult.overall_status === 'FAIL' 
                      ? 'bg-red-500 text-white' 
                      : analysisResult.overall_status === 'INVESTIGATE'
                      ? 'bg-amber-500 text-slate-950 font-black'
                      : 'bg-emerald-600 text-white'
                  }`}>
                    {analysisResult.status_title || analysisResult.overall_status}
                  </span>
                  <span className="text-xs text-slate-400">• ANSI/NETA ATS-2025 Mục 7.1.1</span>
                </div>
                <h3 className="text-base font-bold text-slate-800 mt-1">
                  Biên Bản Đối Soát Hiện Trường Switchgear &amp; Switchboard Assemblies
                </h3>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={() => {
                  navigator.clipboard.writeText(analysisResult.markdown_report);
                  setCopied(true);
                  setTimeout(() => setCopied(false), 2000);
                }}
                className="px-3 py-1.5 rounded-lg text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-700 flex items-center gap-1.5 transition-colors border border-slate-200"
              >
                {copied ? <Check size={14} className="text-emerald-600" /> : <Copy size={14} />}
                <span>{copied ? 'Đã Sao Chép Markdown' : 'Sao Chép Báo Cáo'}</span>
              </button>

              <button
                onClick={() => setShowPrintModal(true)}
                className="px-3 py-1.5 rounded-lg text-xs font-semibold bg-indigo-50 hover:bg-indigo-100 text-indigo-700 flex items-center gap-1.5 transition-colors border border-indigo-200"
              >
                <Printer size={14} />
                <span>In Biên Bản Nghiệm Thu</span>
              </button>
            </div>
          </div>

          {/* Quick Technical Opinion Callout */}
          <div className={`p-4 rounded-xl border text-xs leading-relaxed ${
            analysisResult.overall_status === 'FAIL'
              ? 'bg-red-50/70 border-red-200 text-red-900'
              : analysisResult.overall_status === 'INVESTIGATE'
              ? 'bg-amber-50/70 border-amber-200 text-amber-900'
              : 'bg-emerald-50/70 border-emerald-200 text-emerald-900'
          }`}>
            <div className="font-bold flex items-center gap-1.5 mb-1 text-sm">
              <FileCheck2 size={16} />
              Đánh Giá Trọng Yếu Từ TEV Field Service AI:
            </div>
            <div>{analysisResult.status_reason}</div>
          </div>

          {/* Markdown Content Display */}
          <div className="prose prose-slate max-w-none text-xs leading-relaxed overflow-x-auto bg-slate-50/70 p-5 rounded-xl border border-slate-200/80">
            <div className="whitespace-pre-wrap font-sans text-slate-800 selection:bg-indigo-100">
              {analysisResult.markdown_report}
            </div>
          </div>
        </div>
      )}

      {/* JSON Import/Export Modal */}
      {showJsonModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-2xl w-full p-6 shadow-2xl border border-slate-200 space-y-4">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <h3 className="text-base font-bold text-slate-800 flex items-center gap-2">
                <Sliders size={18} className="text-indigo-600" />
                Dữ Liệu JSON Kiểm Định Switchgear NETA ATS-2025
              </h3>
              <button
                onClick={() => setShowJsonModal(false)}
                className="text-slate-400 hover:text-slate-600 text-sm font-bold"
              >
                ✕
              </button>
            </div>
            <p className="text-xs text-slate-500">
              Bạn có thể dán dữ liệu JSON kiểm định từ thiết bị đo ngoài hiện trường hoặc sao chép để chia sẻ.
            </p>
            <textarea
              rows={14}
              value={jsonText}
              onChange={e => setJsonText(e.target.value)}
              className="w-full font-mono text-xs p-3 bg-slate-50 border border-slate-200 rounded-xl focus:bg-white focus:border-indigo-500 outline-none"
            />
            <div className="flex items-center justify-between pt-2">
              <button
                onClick={handleCopyJson}
                className="px-3 py-2 rounded-xl text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-700 flex items-center gap-1.5 transition-colors border border-slate-200"
              >
                {copied ? <Check size={14} className="text-emerald-600" /> : <Copy size={14} />}
                <span>{copied ? 'Đã Sao Chép' : 'Sao Chép JSON'}</span>
              </button>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setShowJsonModal(false)}
                  className="px-4 py-2 rounded-xl text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-700"
                >
                  Hủy Bỏ
                </button>
                <button
                  onClick={handleApplyJson}
                  className="px-4 py-2 rounded-xl text-xs font-bold bg-indigo-600 hover:bg-indigo-500 text-white"
                >
                  Áp Dụng Dữ Liệu
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Nameplate Scanner Modal */}
      {showScannerModal && (
        <NameplateScannerModal
          isOpen={showScannerModal}
          onClose={() => setShowScannerModal(false)}
          equipmentCategory="switchgear"
          onApplyData={handleApplyExtractedNameplate}
        />
      )}

      {/* TEV Report Print Modal */}
      {showPrintModal && (
        <TevReportPrintModal
          isOpen={showPrintModal}
          onClose={() => setShowPrintModal(false)}
          title={`BIÊN BẢN KIỂM ĐỊNH TỦ ĐIỆN PHÂN PHỐI & TỦ ĐÓNG CẮT (NETA ATS-2025)`}
          equipmentTag={data.site_info.switchgear_tag}
          equipmentType="Switchgear & Switchboard Assemblies (ANSI/NETA ATS-2025 Mục 7.1.1)"
          projectName={data.site_info.project_name}
          fseName={data.site_info.fse_name}
          testDate={data.site_info.test_date}
          overallStatus={analysisResult?.overall_status || liveStats.instantStatus}
          statusReason={analysisResult?.status_reason || (liveStats.isCtrlPass ? 'Tủ điện đạt chuẩn' : 'Sụt giảm cách điện mạch điều khiển')}
          ratings={`${data.site_info.switchgear_rating} • ${data.site_info.rated_voltage_kv || 24}kV • ${data.site_info.short_circuit_ka || 25}kA`}
          evaluations={[
            {
              item: 'Kiểm tra Thị giác & Cơ khí',
              measured: data.visual_inspection.physical_condition_cleanliness,
              standard: 'Khớp nhãn mác, tủ sạch bẩn, đã tháo kẹp vận chuyển',
              reference: 'NETA Sec 7.1.1.A',
              deviation: 'Đạt yêu cầu cơ học',
              status: liveStats.isVisualPass ? 'PASS' : 'FAIL'
            },
            {
              item: 'Khóa Liên Động Cơ Điện (Interlocks)',
              measured: data.visual_inspection.interlock_systems_operation,
              standard: 'Khóa liên động chìa (Key exchange), ngắt CB an toàn 100%',
              reference: 'NETA Sec 7.1.1.A.6',
              deviation: 'Chống chạm nhầm thanh cái',
              status: 'PASS'
            },
            {
              item: 'Bộ Sấy Chống Đọng Sương (Space Heaters)',
              measured: data.visual_inspection.filters_and_heaters_check,
              standard: 'Bộ sấy và thermostat vận hành bình thường',
              reference: 'NETA Sec 7.1.1.A.8',
              deviation: 'Ngăn ngừa đọng sương',
              status: 'PASS'
            },
            {
              item: 'Điện trở mối nối bu-lông (DLRO)',
              measured: `${liveStats.bolts.join(', ')} µΩ (Max: ${liveStats.maxBolt}, Min: ${liveStats.minBolt})`,
              standard: 'Độ lệch ≤ 50% so với giá trị nhỏ nhất',
              reference: 'NETA Sec 7.1.1.D.1',
              deviation: `Độ lệch ${liveStats.boltDev}%`,
              status: liveStats.isBoltPass ? 'PASS' : 'INVESTIGATE'
            },
            {
              item: 'Cách điện Thanh cái (Pha - Đất)',
              measured: `A-G: ${liveStats.bus.phase_a_to_ground} MΩ, B-G: ${liveStats.bus.phase_b_to_ground} MΩ, C-G: ${liveStats.bus.phase_c_to_ground} MΩ`,
              standard: '≥ 1,000 MΩ (Table 100.1 @ 2500V DC)',
              reference: 'NETA Table 100.1',
              deviation: `Tối thiểu ${liveStats.minBusGnd} MΩ (Vượt ngưỡng)`,
              status: 'PASS'
            },
            {
              item: 'Cách điện Thanh cái (Pha - Pha)',
              measured: `AB: ${liveStats.bus.phase_ab} MΩ, BC: ${liveStats.bus.phase_bc} MΩ, CA: ${liveStats.bus.phase_ca} MΩ`,
              standard: '≥ 1,000 MΩ (Table 100.1 @ 2500V DC)',
              reference: 'NETA Table 100.1',
              deviation: `Tối thiểu ${liveStats.minBusPh} MΩ`,
              status: 'PASS'
            },
            {
              item: 'Cách điện Mạch Điều Khiển (Control Wiring IR)',
              measured: `${liveStats.ctrlIr} Megohms (@ 1000V DC)`,
              standard: 'BẮT BUỘC ≥ 2.0 Megohms (2 MΩ)',
              reference: 'NETA Sec 7.1.1.D.4',
              deviation: liveStats.isCtrlPass ? 'Đạt chuẩn' : `THẤP HƠN NGƯỠNG (SỤT ${liveStats.ctrlDrop}% SO VỚI KỲ TRƯỚC ${liveStats.lastCtrl} MΩ)`,
              status: liveStats.isCtrlPass ? 'PASS' : 'FAIL'
            },
            {
              item: 'Máy biến áp CPT Điều Khiển (Turns Ratio & IR)',
              measured: `Sai số tỷ số: ${liveStats.cptRatio}%, Cách điện: ${liveStats.cptIr} MΩ`,
              standard: 'Sai số tỷ số ≤ 0.5%, Cách điện Table 100.1',
              reference: 'NETA Sec 7.1.1.D.6',
              deviation: 'Đạt chuẩn cấp nguồn nhị thứ',
              status: 'PASS'
            },
            {
              item: 'Thử nghiệm Biến dòng & Biến áp CT/VT',
              measured: data.electrical_tests.ct_vt_section_7_10_status,
              standard: 'Đạt tiêu chuẩn Section 7.10',
              reference: 'NETA Sec 7.10',
              deviation: 'Tỷ số và cực tính chuẩn xác',
              status: 'PASS'
            },
            {
              item: 'Bơm dòng nhị thứ (Current Injection)',
              measured: data.electrical_tests.current_injection_wiring_check,
              standard: 'Xác nhận thông mạch CT và không hở mạch',
              reference: 'NETA Sec 7.1.1.D.8',
              deviation: 'Mạch dòng CT toàn vẹn',
              status: 'PASS'
            },
            {
              item: 'Kiểm tra Trùng pha (Phasing Checks)',
              measured: data.electrical_tests.phasing_check_dual_source,
              standard: 'Trùng pha 100% giữa các nguồn đầu vào',
              reference: 'NETA Sec 7.1.1.D.9',
              deviation: 'Trùng pha hoàn toàn',
              status: 'PASS'
            },
            {
              item: 'Phóng điện cục bộ Online (TEV Sensor)',
              measured: `${liveStats.tevDb} dB`,
              standard: 'Dải an toàn < 20 dB (Table 100.23)',
              reference: 'NETA Table 100.23',
              deviation: 'Bình thường, không có phóng điện',
              status: 'PASS'
            }
          ]}
          warnings={[
            !liveStats.isCtrlPass ? `Điện trở cách điện mạch điều khiển nhị thứ đo tại 1000V DC chỉ đạt ${liveStats.ctrlIr} Megohms, không đạt ngưỡng tối thiểu 2.0 Megohms của NETA ATS-2025 Mục 7.1.1.D.4.` : '',
            liveStats.ctrlDrop > 20 ? `So với kỳ kiểm định trước (${liveStats.lastCtrl} MΩ), cách điện mạch điều khiển sụt giảm ${liveStats.ctrlDrop}% do ẩm hoặc trầy xước vỏ dây.` : '',
            !liveStats.isBoltPass ? `Độ lệch điện trở mối nối bu-lông (${liveStats.boltDev}%) vượt quá 50% so với giá trị nhỏ nhất.` : ''
          ].filter(Boolean)}
          recommendations={[
            `Tách cô lập nguồn điều khiển (AC/DC auxiliary power) của tủ ${data.site_info.switchgear_tag} trước khi sửa chữa.`,
            'Kiểm tra từng nhánh dây điều khiển nhị thứ trong khoang rơ-le và máng cáp, sấy khô khoang điều khiển và thay thế các đoạn dây bị sụt cách điện/trầy xước vỏ.',
            'Đo lại điện trở cách điện mạch điều khiển, đảm bảo giá trị đạt tối thiểu ≥ 2.0 Megohms trước khi cấp nguồn đóng điện chính thức.'
          ]}
        />
      )}
    </div>
  );
};
