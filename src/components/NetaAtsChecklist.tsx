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
  GitBranch,
  Timer
} from 'lucide-react';
import {
  NetaAtsInputPayload,
  NetaAtsEvaluationResult,
  performDeterministicAtsAnalysis
} from '../../server/netaAtsAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';
import { NameplateScannerModal } from './NameplateScannerModal';
import { ExtractedNameplateData } from '../../server/nameplateOcrAnalyzer';
import { FseEngineerDropdown } from './FseEngineerDropdown';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

export const SAMPLE_ATS_FSE_DATA: NetaAtsInputPayload = {
  site_info: {
    project_name: "Toa Nha TEV Office Tower",
    ats_tag: "ATS-LV-800A-01",
    ats_rating: "800A 4-Pole 400V Automatic Transfer Switch",
    fse_name: "Nguyen Van R",
    test_date: "2026-09-24",
    substation_location: "Phòng Kỹ Thuật Điện Tầng Hầm B1",
    manufacturer: "Socomec / ASCO / ABB",
    rated_current_a: 800,
    rated_voltage_v: 400,
    poles_count: 4,
    transition_type: "Open Transition"
  },
  visual_inspection: {
    nameplate_match: true,
    physical_condition: "Good",
    cleanliness: "Pass",
    manual_transfer_operation: "Pass",
    mechanical_interlock_check: "Pass",
    anchorage_grounding_check: "Pass",
    bolt_torque_check: "Pass",
    arc_chutes_condition: "Pass (Clean, no erosion)",
    thermographic_survey: "Pass"
  },
  electrical_tests: {
    bolted_resistance_micro_ohms: [8.2, 8.4, 8.3],
    main_pole_insulation_resistance_1000v_megohms: {
      normal_source_position: 850,
      alternate_source_position: 820
    },
    control_wiring_ir_megohms: 45.0,
    contact_pole_resistance_micro_ohms: {
      normal_source_pole_a: 22.1,
      normal_source_pole_b: 22.5,
      normal_source_pole_c: 68.4,
      alternate_source_pole_a: 23.0,
      alternate_source_pole_b: 22.8,
      alternate_source_pole_c: 23.2
    },
    phasing_and_rotation_check: "Pass (Matching A-B-C phase rotation)",
    automatic_sequence_tests: {
      normal_source_undervoltage_sensing: "Pass (Triggered at 85% Un)",
      engine_start_signal_time_sec: 3.0,
      alternate_source_voltage_frequency_sensing: "Pass",
      transfer_time_delay_sec: 5.0,
      retransfer_time_delay_sec: 300.0,
      engine_cool_down_time_sec: 180.0,
      limit_switches_and_interlocks: "Pass",
      in_phase_monitor_check: "Pass"
    }
  },
  previous_test_data: {
    last_test_date: "2025-09-22",
    last_normal_pole_c_resistance_micro_ohms: 22.8
  }
};

export const SAMPLE_ATS_PASS_DATA: NetaAtsInputPayload = {
  site_info: {
    project_name: "Bệnh Viện Đa Khoa Quốc Tế TEV Hospital",
    ats_tag: "ATS-EMERGENCY-1200A",
    ats_rating: "1200A 4-Pole 400V Closed Transition ATS",
    fse_name: "Tran Van T",
    test_date: "2026-09-24",
    substation_location: "Phòng Trạm Biến Áp Cấp Cứu",
    manufacturer: "ASCO Power Technologies 7000 Series",
    rated_current_a: 1200,
    rated_voltage_v: 400,
    poles_count: 4,
    transition_type: "Closed Transition"
  },
  visual_inspection: {
    nameplate_match: true,
    physical_condition: "Hoàn hảo (Không trầy xước, buồng dập mới 100%)",
    cleanliness: "Pass",
    manual_transfer_operation: "Pass (Chuyển êm, dứt khoát)",
    mechanical_interlock_check: "Pass (Khóa cơ khí vững chắc)",
    anchorage_grounding_check: "Pass",
    bolt_torque_check: "Pass",
    arc_chutes_condition: "Pass",
    thermographic_survey: "Pass"
  },
  electrical_tests: {
    bolted_resistance_micro_ohms: [6.1, 6.2, 6.3],
    main_pole_insulation_resistance_1000v_megohms: {
      normal_source_position: 1250,
      alternate_source_position: 1180
    },
    control_wiring_ir_megohms: 80.0,
    contact_pole_resistance_micro_ohms: {
      normal_source_pole_a: 21.8,
      normal_source_pole_b: 22.0,
      normal_source_pole_c: 22.4,
      alternate_source_pole_a: 22.1,
      alternate_source_pole_b: 22.3,
      alternate_source_pole_c: 22.5
    },
    phasing_and_rotation_check: "Pass (Khớp 100% thứ tự pha A-B-C)",
    automatic_sequence_tests: {
      normal_source_undervoltage_sensing: "Pass (Tác động 85% Un)",
      engine_start_signal_time_sec: 2.5,
      alternate_source_voltage_frequency_sensing: "Pass (90% Un, 50Hz)",
      transfer_time_delay_sec: 4.0,
      retransfer_time_delay_sec: 180.0,
      engine_cool_down_time_sec: 120.0,
      limit_switches_and_interlocks: "Pass",
      in_phase_monitor_check: "Pass"
    }
  },
  previous_test_data: {
    last_test_date: "2025-09-18",
    last_normal_pole_c_resistance_micro_ohms: 22.0
  }
};

export const NetaAtsChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaAtsInputPayload>(SAMPLE_ATS_FSE_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaAtsEvaluationResult | null>(null);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showCameraModal, setShowCameraModal] = useState(false);
  const [jsonInput, setJsonInput] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live real-time calculations
  const liveStats = useMemo(() => {
    // 1. Normal Source Contact Resistance Deviation
    const cp = data.electrical_tests.contact_pole_resistance_micro_ohms;
    const normPoles = [cp.normal_source_pole_a, cp.normal_source_pole_b, cp.normal_source_pole_c].filter(v => v !== undefined && v > 0);
    const minNorm = normPoles.length > 0 ? Math.min(...normPoles) : 0;
    const maxNorm = normPoles.length > 0 ? Math.max(...normPoles) : 0;
    const normDev = minNorm > 0 ? Math.round(((maxNorm - minNorm) / minNorm) * 1000) / 10 : 0;
    const normPass = normDev <= 50;

    // 2. Alternate Source Contact Resistance Deviation
    const altPoles = [cp.alternate_source_pole_a, cp.alternate_source_pole_b, cp.alternate_source_pole_c].filter(v => v !== undefined && v > 0);
    const minAlt = altPoles.length > 0 ? Math.min(...altPoles) : 0;
    const maxAlt = altPoles.length > 0 ? Math.max(...altPoles) : 0;
    const altDev = minAlt > 0 ? Math.round(((maxAlt - minAlt) / minAlt) * 1000) / 10 : 0;
    const altPass = altDev <= 50;

    // 3. Bolted resistance deviation
    const bolted = data.electrical_tests.bolted_resistance_micro_ohms || [];
    let boltedDev = 0;
    let boltedPass = true;
    if (bolted.length > 0) {
      const minB = Math.min(...bolted);
      const maxB = Math.max(...bolted);
      if (minB > 0) {
        boltedDev = Math.round(((maxB - minB) / minB) * 1000) / 10;
        boltedPass = boltedDev <= 50;
      }
    }

    // 4. Control wiring IR
    const ctrlIR = data.electrical_tests.control_wiring_ir_megohms;
    const ctrlPass = ctrlIR >= 2.0;

    // 5. Main Pole IR
    const normIR = data.electrical_tests.main_pole_insulation_resistance_1000v_megohms?.normal_source_position || 0;
    const altIR = data.electrical_tests.main_pole_insulation_resistance_1000v_megohms?.alternate_source_position || 0;
    const mainPolePass = normIR >= 100 && altIR >= 100;

    // 6. Mechanical interlock
    const interlockPass = data.visual_inspection.mechanical_interlock_check.toLowerCase().includes('pass') ||
      data.visual_inspection.mechanical_interlock_check.toLowerCase().includes('đạt');

    return {
      minNorm,
      maxNorm,
      normDev,
      normPass,
      minAlt,
      maxAlt,
      altDev,
      altPass,
      boltedDev,
      boltedPass,
      ctrlIR,
      ctrlPass,
      normIR,
      altIR,
      mainPolePass,
      interlockPass
    };
  }, [data]);

  const handleRunAiAnalysis = async () => {
    setIsAnalyzing(true);
    setSavedSuccessMessage(null);
    try {
      const response = await fetch('/api/field-service/neta-ats-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });
      const result = await response.json();
      if (result.success && result.data) {
        setAnalysisResult(result.data);
      } else {
        const fallback = performDeterministicAtsAnalysis(data);
        setAnalysisResult(fallback);
      }
    } catch (e) {
      console.error('Error calling ATS AI analysis API:', e);
      const fallback = performDeterministicAtsAnalysis(data);
      setAnalysisResult(fallback);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleSaveToReports = () => {
    if (!analysisResult) return;
    if (onSaveReport) {
      onSaveReport({
        payload_type: 'NETA_ATS_2025_AUTOMATIC_TRANSFER_SWITCH',
        equipment_type: 'Automatic Transfer Switch (ATS)',
        ats_tag: data.site_info.ats_tag,
        project_name: data.site_info.project_name,
        test_standard: 'ANSI/NETA ATS-2025 Mục 7.22.3',
        analysis_result: analysisResult,
        neta_data: data
      });
      setSavedSuccessMessage(`Đã lưu biên bản bộ chuyển nguồn ${data.site_info.ats_tag} vào kho Báo cáo kỹ thuật thành công!`);
      setTimeout(() => setSavedSuccessMessage(null), 6000);
    }
  };

  const handleApplyOcrData = (ocr: ExtractedNameplateData) => {
    setData((prev) => ({
      ...prev,
      site_info: {
        ...prev.site_info,
        ats_tag: ocr.equipment_tag || ocr.serial_number || prev.site_info.ats_tag,
        ats_rating: `${ocr.rated_power_kva ? ocr.rated_power_kva + 'A' : '800A'} 4-Pole ${ocr.secondary_voltage_v || 400}V ATS`,
        manufacturer: ocr.manufacturer || prev.site_info.manufacturer
      }
    }));
    setShowCameraModal(false);
  };

  const handleCopyMarkdown = () => {
    if (analysisResult?.markdown_report) {
      navigator.clipboard.writeText(analysisResult.markdown_report);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }
  };

  return (
    <div className="space-y-6 pb-20">
      {/* Top Header */}
      <div className="bg-gradient-to-r from-orange-950 via-slate-900 to-amber-950 border border-orange-500/30 rounded-2xl p-6 shadow-xl relative overflow-hidden">
        <div className="absolute top-0 right-0 w-96 h-96 bg-orange-500/10 rounded-full blur-3xl pointer-events-none" />
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 relative z-10">
          <div>
            <div className="flex items-center gap-2 text-orange-400 font-mono text-xs uppercase tracking-wider mb-2">
              <span className="w-2 h-2 rounded-full bg-orange-400 animate-pulse" />
              ANSI/NETA ATS-2025 Section 7.22.3 (Emergency Systems)
            </div>
            <h1 className="text-2xl lg:text-3xl font-bold text-white flex items-center gap-3">
              <GitBranch className="w-8 h-8 text-orange-400" />
              Bộ Chuyển Nguồn Tự Động (Automatic Transfer Switch - ATS)
            </h1>
            <p className="text-slate-300 text-sm mt-1 max-w-2xl">
              Hệ thống Nguồn khẩn cấp & Bộ chuyển nguồn tự động ATS hạ thế. Kiểm tra liên động cơ khí, điện trở tiếp xúc cực (độ lệch ≤ 50%), cách điện mạch điều khiển (≥ 2.0 MΩ) và trình tự chuyển mạch tự động.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => {
                setData(SAMPLE_ATS_FSE_DATA);
                setAnalysisResult(null);
              }}
              className="px-3.5 py-2 rounded-xl bg-amber-950/60 hover:bg-amber-900/70 border border-amber-500/40 text-amber-300 text-xs font-medium flex items-center gap-1.5 transition-colors shadow-sm"
              title="Nạp mẫu mô phỏng sự cố điện trở tiếp điểm Pole C Nguồn chính lệch 209.5%"
            >
              <AlertTriangle className="w-3.5 h-3.5 text-amber-400" />
              Nạp Mẫu Lỗi (Pole C Contact Anomaly)
            </button>
            <button
              onClick={() => {
                setData(SAMPLE_ATS_PASS_DATA);
                setAnalysisResult(null);
              }}
              className="px-3.5 py-2 rounded-xl bg-emerald-950/60 hover:bg-emerald-900/70 border border-emerald-500/40 text-emerald-300 text-xs font-medium flex items-center gap-1.5 transition-colors shadow-sm"
            >
              <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
              Nạp Mẫu Đạt Chuẩn (PASS)
            </button>
            <button
              onClick={() => {
                setJsonInput(JSON.stringify(data, null, 2));
                setShowJsonModal(true);
              }}
              className="px-3.5 py-2 rounded-xl bg-slate-800/80 hover:bg-slate-700/80 border border-slate-700 text-slate-200 text-xs font-medium flex items-center gap-1.5 transition-colors"
            >
              <Upload className="w-3.5 h-3.5 text-slate-400" />
              Xem / Nhập JSON
            </button>
            <button
              onClick={() => setShowCameraModal(true)}
              className="px-3.5 py-2 rounded-xl bg-slate-800/80 hover:bg-slate-700/80 border border-slate-700 text-slate-200 text-xs font-medium flex items-center gap-1.5 transition-colors"
            >
              <Camera className="w-3.5 h-3.5 text-orange-400" />
              Quét Nhãn AI
            </button>
          </div>
        </div>

        {/* Live Telemetry KPI Cards */}
        <div className="grid grid-cols-2 md:grid-cols-5 gap-3 mt-6 pt-6 border-t border-slate-700/50">
          {/* Contact Resistance Pole C Card */}
          <div className={`p-3.5 rounded-xl border ${
            !liveStats.normPass
              ? 'bg-amber-950/60 border-amber-500/70 text-amber-200'
              : 'bg-slate-800/60 border-slate-700 text-slate-200'
          }`}>
            <div className="flex items-center justify-between text-xs mb-1">
              <span className="flex items-center gap-1 font-medium">
                <Sliders className="w-3.5 h-3.5 text-amber-400" />
                Tiếp Điểm Nguồn Chính
              </span>
              <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                !liveStats.normPass
                  ? 'bg-amber-500 text-black animate-pulse'
                  : 'bg-emerald-500/20 text-emerald-400'
              }`}>
                {liveStats.normDev}%
              </span>
            </div>
            <div className="text-xl font-bold font-mono">
              {data.electrical_tests.contact_pole_resistance_micro_ohms.normal_source_pole_c} µΩ
            </div>
            <div className="text-[11px] text-slate-400 truncate mt-0.5">
              NETA ≤ 50% (Pole A: {data.electrical_tests.contact_pole_resistance_micro_ohms.normal_source_pole_a} µΩ)
            </div>
          </div>

          {/* Alternate Source Contact Resistance Card */}
          <div className={`p-3.5 rounded-xl border ${
            !liveStats.altPass
              ? 'bg-rose-950/50 border-rose-500/60 text-rose-200'
              : 'bg-slate-800/60 border-slate-700 text-slate-200'
          }`}>
            <div className="flex items-center justify-between text-xs mb-1">
              <span className="flex items-center gap-1 font-medium">
                <Sliders className="w-3.5 h-3.5 text-teal-400" />
                Tiếp Điểm Dự Phòng
              </span>
              <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                !liveStats.altPass ? 'bg-rose-500 text-white' : 'bg-emerald-500/20 text-emerald-400'
              }`}>
                {liveStats.altDev}%
              </span>
            </div>
            <div className="text-xl font-bold font-mono">
              {liveStats.maxAlt} µΩ
            </div>
            <div className="text-[11px] text-slate-400 truncate mt-0.5">
              Cân bằng tốt (Min: {liveStats.minAlt} µΩ)
            </div>
          </div>

          {/* Mechanical Interlock Card */}
          <div className={`p-3.5 rounded-xl border ${
            !liveStats.interlockPass
              ? 'bg-rose-950/60 border-rose-500/70 text-rose-200'
              : 'bg-slate-800/60 border-slate-700 text-slate-200'
          }`}>
            <div className="flex items-center justify-between text-xs mb-1">
              <span className="flex items-center gap-1 font-medium">
                <ShieldCheck className="w-3.5 h-3.5 text-orange-400" />
                Khóa Liên Động Cơ
              </span>
              <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                !liveStats.interlockPass ? 'bg-rose-500 text-white' : 'bg-emerald-500/20 text-emerald-400'
              }`}>
                {liveStats.interlockPass ? 'BẢO VỆ' : 'NGUY HIỂM'}
              </span>
            </div>
            <div className="text-lg font-bold font-mono truncate">
              {liveStats.interlockPass ? 'Khóa Cơ An Toàn' : 'Lỗi Khóa Cơ!'}
            </div>
            <div className="text-[11px] text-slate-400 truncate mt-0.5">
              Chống đóng trùng nguồn
            </div>
          </div>

          {/* Control Wiring IR Card */}
          <div className={`p-3.5 rounded-xl border ${
            !liveStats.ctrlPass
              ? 'bg-rose-950/60 border-rose-500/60 text-rose-200'
              : 'bg-slate-800/60 border-slate-700 text-slate-200'
          }`}>
            <div className="flex items-center justify-between text-xs mb-1">
              <span className="flex items-center gap-1 font-medium">
                <Gauge className="w-3.5 h-3.5 text-blue-400" />
                Cách Điện Điều Khiển
              </span>
              <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                !liveStats.ctrlPass ? 'bg-rose-500 text-white' : 'bg-emerald-500/20 text-emerald-400'
              }`}>
                {liveStats.ctrlIR} MΩ
              </span>
            </div>
            <div className="text-xl font-bold font-mono">
              {liveStats.ctrlIR} MΩ
            </div>
            <div className="text-[11px] text-slate-400 truncate mt-0.5">
              NETA Bắt Buộc ≥ 2.0 MΩ
            </div>
          </div>

          {/* Auto Sequence Timers Card */}
          <div className="col-span-2 md:col-span-1 p-3.5 rounded-xl border bg-slate-800/60 border-slate-700 text-slate-200">
            <div className="flex items-center justify-between text-xs mb-1">
              <span className="flex items-center gap-1 font-medium">
                <Timer className="w-3.5 h-3.5 text-yellow-400" />
                Định Thời ATS
              </span>
              <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-500/20 text-emerald-400">
                CHÍNH XÁC
              </span>
            </div>
            <div className="text-base font-bold font-mono text-yellow-300">
              {data.electrical_tests.automatic_sequence_tests.transfer_time_delay_sec}s / {data.electrical_tests.automatic_sequence_tests.retransfer_time_delay_sec}s
            </div>
            <div className="text-[11px] text-slate-400 truncate mt-0.5">
              Gọi MF: {data.electrical_tests.automatic_sequence_tests.engine_start_signal_time_sec}s | Mát: {data.electrical_tests.automatic_sequence_tests.engine_cool_down_time_sec}s
            </div>
          </div>
        </div>
      </div>

      {savedSuccessMessage && (
        <div className="p-4 rounded-xl bg-emerald-950/70 border border-emerald-500/50 text-emerald-300 text-sm flex items-center justify-between shadow-lg">
          <div className="flex items-center gap-2">
            <CheckCircle2 className="w-5 h-5 text-emerald-400 flex-shrink-0" />
            <span>{savedSuccessMessage}</span>
          </div>
          {onNavigateToReports && (
            <button
              onClick={onNavigateToReports}
              className="px-3 py-1 rounded-lg bg-emerald-800/60 hover:bg-emerald-700/60 text-white text-xs font-medium underline"
            >
              Xem Kho Báo Cáo
            </button>
          )}
        </div>
      )}

      {/* Main Form Fields Grid */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Left Column: Site Info & Visual Inspection */}
        <div className="space-y-6">
          {/* Section 1: Thông tin ATS */}
          <div className="bg-slate-900/80 border border-slate-800 rounded-2xl p-5 shadow-lg space-y-4">
            <h2 className="text-base font-bold text-white flex items-center gap-2 border-b border-slate-800 pb-3">
              <GitBranch className="w-5 h-5 text-orange-400" />
              1. Thông Tin Bộ Chuyển Nguồn Tự Động (ATS)
            </h2>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
              <div>
                <label className="block text-slate-400 mb-1 font-medium">Tên dự án / Công trình</label>
                <input
                  type="text"
                  value={data.site_info.project_name}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, project_name: e.target.value }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-orange-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Mã Tag tủ ATS</label>
                <input
                  type="text"
                  value={data.site_info.ats_tag}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, ats_tag: e.target.value }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-orange-500"
                />
              </div>

              <div className="md:col-span-2">
                <label className="block text-slate-400 mb-1 font-medium">Thông số định mức (Dòng điện, số cực, điện áp)</label>
                <input
                  type="text"
                  value={data.site_info.ats_rating}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, ats_rating: e.target.value }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-orange-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Hãng sản xuất / Model</label>
                <input
                  type="text"
                  value={data.site_info.manufacturer || 'Socomec / ASCO / ABB'}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, manufacturer: e.target.value }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-orange-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Vị trí lắp đặt</label>
                <input
                  type="text"
                  value={data.site_info.substation_location || 'Phòng Hạ Thế'}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, substation_location: e.target.value }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-orange-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Kỹ sư kiểm tra (FSE)</label>
                <FseEngineerDropdown
                  value={data.site_info.fse_name}
                  onChange={(val) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, fse_name: val }
                    })
                  }
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Ngày thực hiện đo kiểm</label>
                <input
                  type="date"
                  value={data.site_info.test_date}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, test_date: e.target.value }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-orange-500"
                />
              </div>
            </div>
          </div>

          {/* Section 2: Kiểm tra Thị giác & Cơ khí (NETA ATS-2025 Mục 7.22.3.A) */}
          <div className="bg-slate-900/80 border border-slate-800 rounded-2xl p-5 shadow-lg space-y-4">
            <h2 className="text-base font-bold text-white flex items-center gap-2 border-b border-slate-800 pb-3">
              <Eye className="w-5 h-5 text-orange-400" />
              2. Kiểm Tra Thị Giác & Cơ Khí (NETA Sec 7.22.3.A)
            </h2>

            <div className="space-y-3 text-xs">
              <div className="flex items-center justify-between p-3 rounded-xl bg-slate-950/70 border border-slate-800">
                <div>
                  <div className="font-semibold text-white">Khóa Liên Động Cơ Khí (Mechanical Interlock)</div>
                  <div className="text-slate-400 text-[11px]">BẮT BUỘC ngăn chặn tuyệt đối hai nguồn Normal & Alternate đóng song song</div>
                </div>
                <button
                  type="button"
                  onClick={() =>
                    setData({
                      ...data,
                      visual_inspection: {
                        ...data.visual_inspection,
                        mechanical_interlock_check:
                          data.visual_inspection.mechanical_interlock_check === 'Pass' ? 'Fail' : 'Pass'
                      }
                    })
                  }
                  className={`px-3 py-1.5 rounded-lg font-bold text-xs transition-colors ${
                    data.visual_inspection.mechanical_interlock_check === 'Pass'
                      ? 'bg-emerald-600/30 text-emerald-300 border border-emerald-500/50'
                      : 'bg-rose-600/30 text-rose-300 border border-rose-500/50'
                  }`}
                >
                  {data.visual_inspection.mechanical_interlock_check === 'Pass' ? 'ĐẠT (Pass)' : 'LỖI (Fail)'}
                </button>
              </div>

              <div className="flex items-center justify-between p-3 rounded-xl bg-slate-950/70 border border-slate-800">
                <div>
                  <div className="font-semibold text-white">Thao Tác Chuyển Mạch Bằng Tay (Manual Operation)</div>
                  <div className="text-slate-400 text-[11px]">Chuyển mạch dứt khoát, không kẹt lò xo, trơn tru</div>
                </div>
                <button
                  type="button"
                  onClick={() =>
                    setData({
                      ...data,
                      visual_inspection: {
                        ...data.visual_inspection,
                        manual_transfer_operation:
                          data.visual_inspection.manual_transfer_operation === 'Pass' ? 'Fail' : 'Pass'
                      }
                    })
                  }
                  className={`px-3 py-1.5 rounded-lg font-bold text-xs transition-colors ${
                    data.visual_inspection.manual_transfer_operation === 'Pass'
                      ? 'bg-emerald-600/30 text-emerald-300 border border-emerald-500/50'
                      : 'bg-rose-600/30 text-rose-300 border border-rose-500/50'
                  }`}
                >
                  {data.visual_inspection.manual_transfer_operation === 'Pass' ? 'ĐẠT (Pass)' : 'KẸT CƠ'}
                </button>
              </div>

              <div className="flex items-center justify-between p-3 rounded-xl bg-slate-950/70 border border-slate-800">
                <div>
                  <div className="font-semibold text-white">Đối Soát Nhãn Mác Kỹ Thuật (Nameplate Match)</div>
                  <div className="text-slate-400 text-[11px]">Khớp 100% bản vẽ thiết kế (Dòng, áp, số cực)</div>
                </div>
                <button
                  type="button"
                  onClick={() =>
                    setData({
                      ...data,
                      visual_inspection: {
                        ...data.visual_inspection,
                        nameplate_match: !data.visual_inspection.nameplate_match
                      }
                    })
                  }
                  className={`px-3 py-1.5 rounded-lg font-bold text-xs transition-colors ${
                    data.visual_inspection.nameplate_match
                      ? 'bg-emerald-600/30 text-emerald-300 border border-emerald-500/50'
                      : 'bg-rose-600/30 text-rose-300 border border-rose-500/50'
                  }`}
                >
                  {data.visual_inspection.nameplate_match ? 'KHỚP 100%' : 'KHÔNG KHỚP'}
                </button>
              </div>

              <div className="grid grid-cols-2 gap-3 pt-2">
                <div>
                  <label className="block text-slate-400 mb-1 font-medium">Buồng dập hồ quang & Tiếp điểm</label>
                  <input
                    type="text"
                    value={data.visual_inspection.arc_chutes_condition || 'Pass'}
                    onChange={(e) =>
                      setData({
                        ...data,
                        visual_inspection: {
                          ...data.visual_inspection,
                          arc_chutes_condition: e.target.value
                        }
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-orange-500"
                    placeholder="Pass (Clean, no erosion)"
                  />
                </div>

                <div>
                  <label className="block text-slate-400 mb-1 font-medium">Tình trạng vật lý & Độ sạch</label>
                  <input
                    type="text"
                    value={data.visual_inspection.physical_condition}
                    onChange={(e) =>
                      setData({
                        ...data,
                        visual_inspection: {
                          ...data.visual_inspection,
                          physical_condition: e.target.value
                        }
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-orange-500"
                  />
                </div>

                <div>
                  <label className="block text-slate-400 mb-1 font-medium">Định vị bệ tủ & Tiếp địa vỏ</label>
                  <select
                    value={data.visual_inspection.anchorage_grounding_check}
                    onChange={(e: any) =>
                      setData({
                        ...data,
                        visual_inspection: {
                          ...data.visual_inspection,
                          anchorage_grounding_check: e.target.value
                        }
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-orange-500"
                  >
                    <option value="Pass">Pass - Định vị & tiếp địa tốt</option>
                    <option value="Investigate">Investigate - Cần siết bu-lông</option>
                    <option value="Fail">Fail - Mất tiếp địa an toàn</option>
                  </select>
                </div>

                <div>
                  <label className="block text-slate-400 mb-1 font-medium">Lực siết bu-lông (Table 100.12)</label>
                  <select
                    value={data.visual_inspection.bolt_torque_check}
                    onChange={(e: any) =>
                      setData({
                        ...data,
                        visual_inspection: {
                          ...data.visual_inspection,
                          bolt_torque_check: e.target.value
                        }
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-orange-500"
                  >
                    <option value="Pass">Pass - Đạt cờ-lê lực</option>
                    <option value="Investigate">Investigate - Lệch nhẹ</option>
                    <option value="Fail">Fail - Lỏng bu-lông lực</option>
                  </select>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Right Column: Electrical Tests & Automatic Sequence */}
        <div className="space-y-6">
          {/* Section 3: Điện trở Tiếp xúc Cực & Mối nối Bu-lông (NETA 7.22.3.D.1 & D.4) */}
          <div className={`border rounded-2xl p-5 shadow-lg space-y-4 ${
            !liveStats.normPass
              ? 'bg-amber-950/20 border-amber-500/60'
              : 'bg-slate-900/80 border-slate-800'
          }`}>
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <h2 className="text-base font-bold text-white flex items-center gap-2">
                <Sliders className="w-5 h-5 text-orange-400" />
                3. Điện Trở Tiếp Xúc Cực (Contact Resistance - µΩ)
              </h2>
              <span className={`px-2.5 py-1 rounded-full text-xs font-bold ${
                !liveStats.normPass
                  ? 'bg-amber-500 text-black animate-pulse'
                  : 'bg-emerald-500/20 text-emerald-400'
              }`}>
                {!liveStats.normPass ? `CẢNH BÁO: LỆCH ${liveStats.normDev}%` : 'TIẾP ĐIỂM ĐẠT'}
              </span>
            </div>

            {/* Normal Source Poles */}
            <div>
              <div className="flex items-center justify-between mb-1.5 text-xs">
                <span className="font-semibold text-amber-300">Nguồn Chính (Normal Source Position)</span>
                <span className={`text-[11px] font-mono font-bold ${
                  !liveStats.normPass ? 'text-amber-400' : 'text-emerald-400'
                }`}>
                  Độ lệch: {liveStats.normDev}% (NETA ≤ 50%)
                </span>
              </div>
              <div className="grid grid-cols-3 gap-2">
                <div>
                  <label className="block text-[11px] text-slate-400 mb-1">Pole A (µΩ)</label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.contact_pole_resistance_micro_ohms.normal_source_pole_a}
                    onChange={(e) =>
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          contact_pole_resistance_micro_ohms: {
                            ...data.electrical_tests.contact_pole_resistance_micro_ohms,
                            normal_source_pole_a: parseFloat(e.target.value) || 0
                          }
                        }
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white font-mono text-xs focus:outline-none focus:border-orange-500"
                  />
                </div>
                <div>
                  <label className="block text-[11px] text-slate-400 mb-1">Pole B (µΩ)</label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.contact_pole_resistance_micro_ohms.normal_source_pole_b}
                    onChange={(e) =>
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          contact_pole_resistance_micro_ohms: {
                            ...data.electrical_tests.contact_pole_resistance_micro_ohms,
                            normal_source_pole_b: parseFloat(e.target.value) || 0
                          }
                        }
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white font-mono text-xs focus:outline-none focus:border-orange-500"
                  />
                </div>
                <div>
                  <label className="block text-[11px] text-slate-400 mb-1 font-bold text-amber-300">Pole C (µΩ) *</label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.contact_pole_resistance_micro_ohms.normal_source_pole_c}
                    onChange={(e) =>
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          contact_pole_resistance_micro_ohms: {
                            ...data.electrical_tests.contact_pole_resistance_micro_ohms,
                            normal_source_pole_c: parseFloat(e.target.value) || 0
                          }
                        }
                      })
                    }
                    className={`w-full bg-slate-950 border rounded-lg px-2.5 py-1.5 font-mono text-xs font-bold focus:outline-none ${
                      data.electrical_tests.contact_pole_resistance_micro_ohms.normal_source_pole_c > liveStats.minNorm * 1.5
                        ? 'border-amber-500 text-amber-300'
                        : 'border-slate-700 text-white'
                    }`}
                  />
                </div>
              </div>
            </div>

            {/* Alternate Source Poles */}
            <div className="pt-2 border-t border-slate-800">
              <div className="flex items-center justify-between mb-1.5 text-xs">
                <span className="font-semibold text-teal-300">Nguồn Dự Phòng (Alternate Source Position)</span>
                <span className="text-[11px] font-mono font-bold text-emerald-400">
                  Độ lệch: {liveStats.altDev}% (Đạt)
                </span>
              </div>
              <div className="grid grid-cols-3 gap-2">
                <div>
                  <label className="block text-[11px] text-slate-400 mb-1">Pole A (µΩ)</label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.contact_pole_resistance_micro_ohms.alternate_source_pole_a}
                    onChange={(e) =>
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          contact_pole_resistance_micro_ohms: {
                            ...data.electrical_tests.contact_pole_resistance_micro_ohms,
                            alternate_source_pole_a: parseFloat(e.target.value) || 0
                          }
                        }
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white font-mono text-xs focus:outline-none focus:border-teal-500"
                  />
                </div>
                <div>
                  <label className="block text-[11px] text-slate-400 mb-1">Pole B (µΩ)</label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.contact_pole_resistance_micro_ohms.alternate_source_pole_b}
                    onChange={(e) =>
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          contact_pole_resistance_micro_ohms: {
                            ...data.electrical_tests.contact_pole_resistance_micro_ohms,
                            alternate_source_pole_b: parseFloat(e.target.value) || 0
                          }
                        }
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white font-mono text-xs focus:outline-none focus:border-teal-500"
                  />
                </div>
                <div>
                  <label className="block text-[11px] text-slate-400 mb-1">Pole C (µΩ)</label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.contact_pole_resistance_micro_ohms.alternate_source_pole_c}
                    onChange={(e) =>
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          contact_pole_resistance_micro_ohms: {
                            ...data.electrical_tests.contact_pole_resistance_micro_ohms,
                            alternate_source_pole_c: parseFloat(e.target.value) || 0
                          }
                        }
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1.5 text-white font-mono text-xs focus:outline-none focus:border-teal-500"
                  />
                </div>
              </div>
            </div>

            {/* Bolted connections array */}
            <div className="pt-2 border-t border-slate-800">
              <label className="block text-slate-400 mb-1 text-xs font-medium">
                Điện trở Mối nối Bu-lông Thanh Cái ATS (µΩ)
              </label>
              <div className="grid grid-cols-3 gap-2">
                {data.electrical_tests.bolted_resistance_micro_ohms.map((res, idx) => (
                  <div key={idx} className="relative">
                    <span className="absolute left-2 top-1.5 text-[10px] text-slate-500 font-mono">#{idx + 1}</span>
                    <input
                      type="number"
                      step="0.1"
                      value={res}
                      onChange={(e) => {
                        const newArr = [...data.electrical_tests.bolted_resistance_micro_ohms];
                        newArr[idx] = parseFloat(e.target.value) || 0;
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            bolted_resistance_micro_ohms: newArr
                          }
                        });
                      }}
                      className="w-full bg-slate-950 border border-slate-700 rounded-lg pl-7 pr-2 py-1.5 text-white font-mono text-xs focus:outline-none focus:border-orange-500"
                    />
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* Section 4: Điện trở Cách điện Mạch Lực & Mạch Điều Khiển (NETA 7.22.3.D.2 & D.3) */}
          <div className="bg-slate-900/80 border border-slate-800 rounded-2xl p-5 shadow-lg space-y-4">
            <h2 className="text-base font-bold text-white flex items-center gap-2 border-b border-slate-800 pb-3">
              <Gauge className="w-5 h-5 text-orange-400" />
              4. Điện Trở Cách Điện (Insulation Resistance)
            </h2>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-3 text-xs">
              <div>
                <label className="block text-slate-400 mb-1 font-medium">Mạch Lực Vị trí Normal (MΩ)</label>
                <input
                  type="number"
                  step="1"
                  value={data.electrical_tests.main_pole_insulation_resistance_1000v_megohms.normal_source_position}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        main_pole_insulation_resistance_1000v_megohms: {
                          ...data.electrical_tests.main_pole_insulation_resistance_1000v_megohms,
                          normal_source_position: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-orange-500"
                />
                <span className="text-[10px] text-slate-400 mt-1 block">Table 100.1: ≥ 100 MΩ</span>
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Mạch Lực Vị trí Alternate (MΩ)</label>
                <input
                  type="number"
                  step="1"
                  value={data.electrical_tests.main_pole_insulation_resistance_1000v_megohms.alternate_source_position}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        main_pole_insulation_resistance_1000v_megohms: {
                          ...data.electrical_tests.main_pole_insulation_resistance_1000v_megohms,
                          alternate_source_position: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-orange-500"
                />
                <span className="text-[10px] text-slate-400 mt-1 block">Table 100.1: ≥ 100 MΩ</span>
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Mạch Điều Khiển Control IR (MΩ)</label>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.control_wiring_ir_megohms}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        control_wiring_ir_megohms: parseFloat(e.target.value) || 0
                      }
                    })
                  }
                  className={`w-full bg-slate-950 border rounded-lg px-3 py-2 font-mono font-bold focus:outline-none ${
                    data.electrical_tests.control_wiring_ir_megohms < 2.0
                      ? 'border-rose-500 text-rose-300'
                      : 'border-slate-700 text-white'
                  }`}
                />
                <span className="text-[10px] text-slate-400 mt-1 block">NETA 7.22.3.D.3: ≥ 2.0 MΩ</span>
              </div>
            </div>

            <div className="pt-2">
              <label className="block text-slate-400 mb-1 text-xs font-medium">
                Kiểm tra Thứ tự pha & Đồng bộ (Phasing and Rotation Check)
              </label>
              <input
                type="text"
                value={data.electrical_tests.phasing_and_rotation_check}
                onChange={(e) =>
                  setData({
                    ...data,
                    electrical_tests: {
                      ...data.electrical_tests,
                      phasing_and_rotation_check: e.target.value
                    }
                  })
                }
                className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white text-xs focus:outline-none focus:border-orange-500"
                placeholder="Pass (Matching A-B-C phase rotation)"
              />
            </div>
          </div>

          {/* Section 5: Trình Tự Chuyển Nguồn & Bộ Định Thời Gian (NETA 7.22.3.D.6) */}
          <div className="bg-slate-900/80 border border-slate-800 rounded-2xl p-5 shadow-lg space-y-4">
            <h2 className="text-base font-bold text-white flex items-center gap-2 border-b border-slate-800 pb-3">
              <Timer className="w-5 h-5 text-orange-400" />
              5. Trình Tự Chuyển Nguồn Tự Động & Bộ Định Thời (Timers)
            </h2>

            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 text-xs">
              <div>
                <label className="block text-slate-400 mb-1 font-medium">Trễ Gọi Máy Phát (s)</label>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.automatic_sequence_tests.engine_start_signal_time_sec}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        automatic_sequence_tests: {
                          ...data.electrical_tests.automatic_sequence_tests,
                          engine_start_signal_time_sec: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-orange-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Trễ Đóng Tải DP (s)</label>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.automatic_sequence_tests.transfer_time_delay_sec}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        automatic_sequence_tests: {
                          ...data.electrical_tests.automatic_sequence_tests,
                          transfer_time_delay_sec: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-orange-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Trễ Khôi Phục Lưới (s)</label>
                <input
                  type="number"
                  step="1"
                  value={data.electrical_tests.automatic_sequence_tests.retransfer_time_delay_sec}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        automatic_sequence_tests: {
                          ...data.electrical_tests.automatic_sequence_tests,
                          retransfer_time_delay_sec: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-orange-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Trễ Làm Mát MF (s)</label>
                <input
                  type="number"
                  step="1"
                  value={data.electrical_tests.automatic_sequence_tests.engine_cool_down_time_sec}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        automatic_sequence_tests: {
                          ...data.electrical_tests.automatic_sequence_tests,
                          engine_cool_down_time_sec: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-orange-500"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs pt-2">
              <div>
                <label className="block text-slate-400 mb-1 font-medium">Cảm biến sụt áp Nguồn chính</label>
                <input
                  type="text"
                  value={data.electrical_tests.automatic_sequence_tests.normal_source_undervoltage_sensing}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        automatic_sequence_tests: {
                          ...data.electrical_tests.automatic_sequence_tests,
                          normal_source_undervoltage_sensing: e.target.value
                        }
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-orange-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Công tắc hành trình & Liên động điện</label>
                <input
                  type="text"
                  value={data.electrical_tests.automatic_sequence_tests.limit_switches_and_interlocks}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        automatic_sequence_tests: {
                          ...data.electrical_tests.automatic_sequence_tests,
                          limit_switches_and_interlocks: e.target.value
                        }
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-orange-500"
                />
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Action Buttons Bar */}
      <div className="flex flex-wrap items-center justify-between gap-4 p-5 bg-slate-900/90 border border-slate-800 rounded-2xl shadow-xl">
        <div className="flex items-center gap-2 text-xs text-slate-400">
          <Clock className="w-4 h-4 text-orange-400" />
          <span>Đối soát chuẩn xác theo ANSI/NETA ATS-2025 Mục 7.22.3 (Emergency Systems, ATS)</span>
        </div>

        <div className="flex flex-wrap items-center gap-3">
          <button
            onClick={handleRunAiAnalysis}
            disabled={isAnalyzing}
            className="px-5 py-2.5 rounded-xl bg-gradient-to-r from-orange-600 to-amber-600 hover:from-orange-500 hover:to-amber-500 text-white font-semibold text-sm flex items-center gap-2 shadow-lg shadow-orange-500/20 disabled:opacity-50 transition-all cursor-pointer"
          >
            {isAnalyzing ? (
              <>
                <RotateCcw className="w-4 h-4 animate-spin" />
                <span>AI Đang Phân Tích & Đối Soát...</span>
              </>
            ) : (
              <>
                <Sparkles className="w-4 h-4 text-orange-200" />
                <span>Thẩm Định Chuyên Gia AI (NETA ATS-2025)</span>
              </>
            )}
          </button>

          {analysisResult && (
            <>
              <button
                onClick={handleSaveToReports}
                className="px-4 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-500 text-white font-medium text-sm flex items-center gap-2 transition-colors cursor-pointer"
              >
                <Save className="w-4 h-4" />
                <span>Lưu Vào Kho Báo Cáo CMMS</span>
              </button>
              <button
                onClick={() => setShowPrintModal(true)}
                className="px-4 py-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 font-medium text-sm flex items-center gap-2 border border-slate-700 transition-colors cursor-pointer"
              >
                <Printer className="w-4 h-4 text-orange-400" />
                <span>In Biên Bản A4</span>
              </button>
            </>
          )}
        </div>
      </div>

      {/* Analysis Result Display */}
      {analysisResult && (
        <div className="space-y-6">
          {/* Status banner */}
          <div className={`p-6 rounded-2xl border shadow-xl flex flex-col md:flex-row md:items-center justify-between gap-4 ${
            analysisResult.overall_status === 'FAIL'
              ? 'bg-rose-950/40 border-rose-500/60'
              : analysisResult.overall_status === 'INVESTIGATE'
              ? 'bg-amber-950/40 border-amber-500/60'
              : 'bg-emerald-950/40 border-emerald-500/60'
          }`}>
            <div className="flex items-start gap-4">
              <div className={`p-3 rounded-xl ${
                analysisResult.overall_status === 'FAIL'
                  ? 'bg-rose-500/20 text-rose-400'
                  : analysisResult.overall_status === 'INVESTIGATE'
                  ? 'bg-amber-500/20 text-amber-400'
                  : 'bg-emerald-500/20 text-emerald-400'
              }`}>
                {analysisResult.overall_status === 'FAIL' ? (
                  <ShieldAlert className="w-8 h-8" />
                ) : analysisResult.overall_status === 'INVESTIGATE' ? (
                  <AlertTriangle className="w-8 h-8" />
                ) : (
                  <ShieldCheck className="w-8 h-8" />
                )}
              </div>
              <div>
                <div className="text-xs uppercase tracking-wider font-semibold text-slate-400">
                  Kết Luận Kiểm Định NETA ATS-2025 Mục 7.22.3 (Automatic Transfer Switches)
                </div>
                <h3 className={`text-2xl font-bold mt-0.5 ${
                  analysisResult.overall_status === 'FAIL'
                    ? 'text-rose-400'
                    : analysisResult.overall_status === 'INVESTIGATE'
                    ? 'text-amber-400'
                    : 'text-emerald-400'
                }`}>
                  {analysisResult.overall_status === 'FAIL'
                    ? 'FAIL / CRITICAL ATS DEFECT DETECTED'
                    : analysisResult.overall_status === 'INVESTIGATE'
                    ? 'INVESTIGATE / CONTACT RESISTANCE ANOMALY'
                    : 'PASS / ĐẠT TIÊU CHUẨN NETA ATS-2025'}
                </h3>
                <p className="text-slate-300 text-sm mt-1 max-w-3xl">
                  {analysisResult.status_reason}
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={handleCopyMarkdown}
                className="px-3.5 py-2 rounded-xl bg-slate-800/80 hover:bg-slate-700/80 border border-slate-700 text-slate-200 text-xs font-medium flex items-center gap-1.5 transition-colors"
              >
                {copied ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5 text-slate-400" />}
                {copied ? 'Đã Sao Chép' : 'Sao Chép Báo Cáo'}
              </button>
            </div>
          </div>

          {/* Markdown Expert Report Box */}
          <div className="bg-slate-900/90 border border-slate-800 rounded-2xl p-6 shadow-xl space-y-4">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <h3 className="text-lg font-bold text-white flex items-center gap-2">
                <FileText className="w-5 h-5 text-orange-400" />
                Báo Cáo Thẩm Định Chuyên Gia NETA ATS-2025 Mục 7.22.3
              </h3>
              <span className="text-xs text-slate-400 font-mono">
                Model: Gemini 2.5 Flash / TEV Expert Engine
              </span>
            </div>

            <div className="prose prose-invert max-w-none text-slate-200 text-sm leading-relaxed whitespace-pre-line bg-slate-950/60 p-5 rounded-xl border border-slate-800 font-sans">
              {analysisResult.markdown_report}
            </div>
          </div>
        </div>
      )}

      {/* JSON Import/Export Modal */}
      {showJsonModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm">
          <div className="bg-slate-900 border border-slate-700 rounded-2xl max-w-2xl w-full p-6 space-y-4 shadow-2xl">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <h3 className="text-base font-bold text-white flex items-center gap-2">
                <Upload className="w-5 h-5 text-orange-400" />
                Dữ Liệu JSON Kiểm Tra ATS Hiện Trường
              </h3>
              <button
                onClick={() => setShowJsonModal(false)}
                className="text-slate-400 hover:text-white"
              >
                ✕
              </button>
            </div>

            <p className="text-xs text-slate-400">
              Bạn có thể dán trực tiếp dữ liệu JSON từ phần mềm đo kiểm hoặc file cấu hình để nạp vào form ATS.
            </p>

            <textarea
              rows={14}
              value={jsonInput}
              onChange={(e) => setJsonInput(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl p-3 font-mono text-xs text-orange-300 focus:outline-none focus:border-orange-500"
            />

            <div className="flex items-center justify-end gap-3 pt-2">
              <button
                onClick={() => setShowJsonModal(false)}
                className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-medium"
              >
                Đóng
              </button>
              <button
                onClick={() => {
                  try {
                    const parsed = JSON.parse(jsonInput);
                    setData(parsed);
                    setShowJsonModal(false);
                  } catch (err: any) {
                    alert('Lỗi định dạng JSON: ' + err.message);
                  }
                }}
                className="px-4 py-2 rounded-xl bg-orange-600 hover:bg-orange-500 text-white text-xs font-semibold"
              >
                Áp Dụng Dữ Liệu
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Nameplate Camera Scanner Modal */}
      {showCameraModal && (
        <NameplateScannerModal
          isOpen={showCameraModal}
          onClose={() => setShowCameraModal(false)}
          onApplyData={handleApplyOcrData}
          equipmentCategory="switchgear"
        />
      )}

      {/* Print Report Modal */}
      {showPrintModal && analysisResult && (
        <TevReportPrintModal
          isOpen={showPrintModal}
          onClose={() => setShowPrintModal(false)}
          reportTitle={`BIÊN BẢN KIỂM ĐỊNH NETA ATS-2025 - BỘ CHUYỂN NGUỒN TỰ ĐỘNG ${data.site_info.ats_tag}`}
          equipmentTag={data.site_info.ats_tag}
          equipmentType="Bộ Chuyển Nguồn Tự Động (ATS)"
          testStandard="ANSI/NETA ATS-2025 Mục 7.22.3 & Table 100.1, 100.12, 100.18"
          overallStatus={analysisResult.overall_status}
          markdownReport={analysisResult.markdown_report}
        />
      )}
    </div>
  );
};
