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
  Timer,
  Cpu,
  TrendingDown,
  AlertOctagon
} from 'lucide-react';
import {
  NetaGeneratorInputPayload,
  NetaGeneratorEvaluationResult,
  performDeterministicGeneratorAnalysis
} from '../../server/netaEngineGeneratorAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';
import { NameplateScannerModal } from './NameplateScannerModal';
import { ExtractedNameplateData } from '../../server/nameplateOcrAnalyzer';
import { FseEngineerDropdown } from './FseEngineerDropdown';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

export const SAMPLE_GENERATOR_FSE_DATA: NetaGeneratorInputPayload = {
  site_info: {
    project_name: "Benh Vien TEV General Hospital",
    generator_tag: "GEN-EMERGENCY-500KW-01",
    generator_rating: "500kW / 625kVA 400V 1500RPM Diesel Generator Set",
    fse_name: "Nguyen Van S",
    test_date: "2026-09-24",
    substation_location: "Phòng Máy Phát Điện Dự Phòng Khối Nhà B",
    engine_manufacturer: "Cummins / Perkins",
    alternator_manufacturer: "Stamford / Leroy Somer",
    rated_kw: 500,
    rated_kva: 625,
    rated_voltage_v: 400,
    rated_rpm: 1500,
    rated_frequency_hz: 50,
    power_factor: 0.8
  },
  visual_inspection: {
    nameplate_match: true,
    physical_condition: "Good (Shipping bracing removed, no fluid leaks)",
    anchorage_alignment_grounding: "Pass",
    cleanliness: "Pass",
    shipping_bracing_removed: true,
    fluid_leaks_check: "Pass",
    exhaust_and_fuel_systems: "Pass"
  },
  electrical_tests: {
    generator_insulation_resistance_40c_megohms: {
      stator_1min_ir_1000v: 650,
      stator_10min_ir_1000v: 1105,
      calculated_pi_value: 1.7,
      rotor_ir_500v_megohms: 250,
      exciter_ir_500v_megohms: 180
    },
    protective_relays_section_7_9_status: "Pass",
    phase_rotation_and_synchronization: "Pass (Matching A-B-C rotation)",
    engine_protection_shutdowns_test: {
      low_oil_pressure_trip: "Pass (Engine tripped at 1.2 bar)",
      high_coolant_temperature_trip: "Pass (Engine tripped at 98 deg C)",
      overspeed_trip: "Fail (Engine failed to trip at 115% rated RPM)",
      low_coolant_level_trip: "Pass",
      emergency_stop_button_check: "Pass"
    },
    bearing_cap_vibration_mils: 1.8,
    nfpa_110_performance_load_test: {
      time_to_first_load_acceptance_sec: 8.5,
      load_bank_100_percent_test_2hrs: "Pass (Maintained 500kW for 2 hours, coolant temp stable at 85 deg C)",
      status: "Pass",
      step_30_percent_load_kw: 150,
      step_50_percent_load_kw: 250,
      step_75_percent_load_kw: 375,
      step_100_percent_load_kw: 500
    },
    governor_and_avr_regulation: {
      voltage_regulation_percent: 0.8,
      frequency_droop_percent: 1.2
    }
  },
  previous_test_data: {
    last_test_date: "2025-09-22",
    last_pi_value: 2.4,
    last_1min_ir_megohms: 720
  }
};

export const SAMPLE_GENERATOR_PASS_DATA: NetaGeneratorInputPayload = {
  site_info: {
    project_name: "Trung Tâm Dữ Liệu TEV Tier-III Data Center",
    generator_tag: "GEN-DATACENTER-1000KW",
    generator_rating: "1000kW / 1250kVA 400V 1500RPM Prime Genset",
    fse_name: "Le Hoang D",
    test_date: "2026-09-24",
    substation_location: "Khu Vực Máy Phát Điện Dự Phòng Khẩn Cấp",
    engine_manufacturer: "Mitsubishi Heavy Industries",
    alternator_manufacturer: "Stamford HCI634J",
    rated_kw: 1000,
    rated_kva: 1250,
    rated_voltage_v: 400,
    rated_rpm: 1500,
    rated_frequency_hz: 50,
    power_factor: 0.8
  },
  visual_inspection: {
    nameplate_match: true,
    physical_condition: "Tốt (Thanh giằng vận chuyển đã tháo, hoàn toàn khô ráo)",
    anchorage_alignment_grounding: "Pass (Định tâm laser chuẩn, tiếp địa đạt 1.2Ω)",
    cleanliness: "Pass",
    shipping_bracing_removed: true,
    fluid_leaks_check: "Pass",
    exhaust_and_fuel_systems: "Pass"
  },
  electrical_tests: {
    generator_insulation_resistance_40c_megohms: {
      stator_1min_ir_1000v: 950,
      stator_10min_ir_1000v: 2660,
      calculated_pi_value: 2.8,
      rotor_ir_500v_megohms: 420,
      exciter_ir_500v_megohms: 310
    },
    protective_relays_section_7_9_status: "Pass",
    phase_rotation_and_synchronization: "Pass (Đồng bộ pha A-B-C)",
    engine_protection_shutdowns_test: {
      low_oil_pressure_trip: "Pass (Cắt máy tại 1.3 bar)",
      high_coolant_temperature_trip: "Pass (Cắt máy tại 97°C)",
      overspeed_trip: "Pass (Cắt máy tức thì tại 1725 RPM - 115% n_đm)",
      low_coolant_level_trip: "Pass",
      emergency_stop_button_check: "Pass"
    },
    bearing_cap_vibration_mils: 1.2,
    nfpa_110_performance_load_test: {
      time_to_first_load_acceptance_sec: 7.2,
      load_bank_100_percent_test_2hrs: "Pass (Giữ vững 1000kW trong 2h, nước làm mát 82°C)",
      status: "Pass",
      step_30_percent_load_kw: 300,
      step_50_percent_load_kw: 500,
      step_75_percent_load_kw: 750,
      step_100_percent_load_kw: 1000
    },
    governor_and_avr_regulation: {
      voltage_regulation_percent: 0.5,
      frequency_droop_percent: 0.8
    }
  },
  previous_test_data: {
    last_test_date: "2025-09-15",
    last_pi_value: 2.7,
    last_1min_ir_megohms: 910
  }
};

export const NetaEngineGeneratorChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaGeneratorInputPayload>(SAMPLE_GENERATOR_FSE_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaGeneratorEvaluationResult | null>(null);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showCameraModal, setShowCameraModal] = useState(false);
  const [jsonInput, setJsonInput] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live real-time stats
  const liveStats = useMemo(() => {
    const ir = data.electrical_tests.generator_insulation_resistance_40c_megohms;
    const pi = ir.calculated_pi_value || (ir.stator_1min_ir_1000v > 0 ? Math.round((ir.stator_10min_ir_1000v / ir.stator_1min_ir_1000v) * 10) / 10 : 0);
    const piPass = pi >= 2.0;

    let piTrend = 0;
    if (data.previous_test_data?.last_pi_value) {
      piTrend = Math.round((pi - data.previous_test_data.last_pi_value) * 10) / 10;
    }

    const shutdowns = data.electrical_tests.engine_protection_shutdowns_test;
    const overspeedPass = shutdowns.overspeed_trip.toLowerCase().includes('pass');
    const oilPass = shutdowns.low_oil_pressure_trip.toLowerCase().includes('pass');
    const tempPass = shutdowns.high_coolant_temperature_trip.toLowerCase().includes('pass');

    const nfpa = data.electrical_tests.nfpa_110_performance_load_test;
    const class10Pass = nfpa.time_to_first_load_acceptance_sec <= 10.0;

    const reg = data.electrical_tests.governor_and_avr_regulation;
    const regPass = Math.abs(reg.voltage_regulation_percent) <= 1.0;

    const vib = data.electrical_tests.bearing_cap_vibration_mils;
    const vibPass = vib <= 2.5;

    return {
      pi,
      piPass,
      piTrend,
      overspeedPass,
      oilPass,
      tempPass,
      class10Pass,
      timeToLoad: nfpa.time_to_first_load_acceptance_sec,
      regPass,
      vReg: reg.voltage_regulation_percent,
      fDroop: reg.frequency_droop_percent,
      vib,
      vibPass
    };
  }, [data]);

  const handleRunAiAnalysis = async () => {
    setIsAnalyzing(true);
    setSavedSuccessMessage(null);
    try {
      const response = await fetch('/api/field-service/neta-generator-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });
      const result = await response.json();
      if (result.success && result.data) {
        setAnalysisResult(result.data);
      } else {
        const fallback = performDeterministicGeneratorAnalysis(data);
        setAnalysisResult(fallback);
      }
    } catch (e) {
      console.error('Error calling Generator AI analysis API:', e);
      const fallback = performDeterministicGeneratorAnalysis(data);
      setAnalysisResult(fallback);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleSaveToReports = () => {
    if (!analysisResult) return;
    if (onSaveReport) {
      onSaveReport({
        payload_type: 'NETA_ATS_2025_ENGINE_GENERATOR',
        equipment_type: 'Emergency Engine Generator',
        generator_tag: data.site_info.generator_tag,
        project_name: data.site_info.project_name,
        test_standard: 'ANSI/NETA ATS-2025 Mục 7.22.1 & NFPA 110',
        analysis_result: analysisResult,
        neta_data: data
      });
      setSavedSuccessMessage(`Đã lưu biên bản tổ máy phát ${data.site_info.generator_tag} vào kho Báo cáo kỹ thuật thành công!`);
      setTimeout(() => setSavedSuccessMessage(null), 6000);
    }
  };

  const handleApplyOcrData = (ocr: ExtractedNameplateData) => {
    setData((prev) => ({
      ...prev,
      site_info: {
        ...prev.site_info,
        generator_tag: ocr.equipment_tag || ocr.serial_number || prev.site_info.generator_tag,
        generator_rating: `${ocr.rated_power_kva || 500}kW 400V 1500RPM Genset`,
        engine_manufacturer: ocr.manufacturer || prev.site_info.engine_manufacturer
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
      {/* Top Banner Header */}
      <div className="bg-gradient-to-r from-red-950 via-slate-900 to-amber-950 border border-red-500/30 rounded-2xl p-6 shadow-xl relative overflow-hidden">
        <div className="absolute top-0 right-0 w-96 h-96 bg-red-500/10 rounded-full blur-3xl pointer-events-none" />
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 relative z-10">
          <div>
            <div className="flex items-center gap-2 text-red-400 font-mono text-xs uppercase tracking-wider mb-2">
              <span className="w-2 h-2 rounded-full bg-red-400 animate-pulse" />
              ANSI/NETA ATS-2025 Section 7.22.1 & NFPA 110 (Emergency Systems)
            </div>
            <h1 className="text-2xl lg:text-3xl font-bold text-white flex items-center gap-3">
              <Power className="w-8 h-8 text-amber-400" />
              Tổ Máy Phát Điện Khẩn Cấp / Dự Phòng (Engine Generator)
            </h1>
            <p className="text-slate-300 text-sm mt-1 max-w-2xl">
              Hệ thống Nguồn điện Khẩn cấp & Máy phát Diesel. Kiểm tra bảo vệ tắt máy (Áp suất dầu, quá nhiệt, quá tốc độ), cách điện stator IEEE 43 (PI ≥ 2.0) và khả năng nhận tải khẩn cấp NFPA 110 Class 10 (≤ 10 giây).
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => {
                setData(SAMPLE_GENERATOR_FSE_DATA);
                setAnalysisResult(null);
              }}
              className="px-3.5 py-2 rounded-xl bg-red-950/60 hover:bg-red-900/70 border border-red-500/40 text-red-300 text-xs font-medium flex items-center gap-1.5 transition-colors shadow-sm"
              title="Nạp mẫu mô phỏng sự cố: Quá tốc độ không ngắt (Overspeed Fail) và PI Stator sụt giảm còn 1.7"
            >
              <AlertOctagon className="w-3.5 h-3.5 text-red-400" />
              Nạp Mẫu Lỗi (Overspeed Fail & Low PI)
            </button>
            <button
              onClick={() => {
                setData(SAMPLE_GENERATOR_PASS_DATA);
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
              <Camera className="w-3.5 h-3.5 text-amber-400" />
              Quét Nhãn AI
            </button>
          </div>
        </div>

        {/* Live Telemetry KPI Cards */}
        <div className="grid grid-cols-2 md:grid-cols-5 gap-3 mt-6 pt-6 border-t border-slate-700/50">
          {/* Overspeed Trip Card (Critical) */}
          <div className={`p-3.5 rounded-xl border ${
            !liveStats.overspeedPass
              ? 'bg-rose-950/70 border-rose-500 text-rose-200 ring-2 ring-rose-500/50'
              : 'bg-slate-800/60 border-slate-700 text-slate-200'
          }`}>
            <div className="flex items-center justify-between text-xs mb-1">
              <span className="flex items-center gap-1 font-medium">
                <AlertOctagon className="w-3.5 h-3.5 text-rose-400" />
                Cắt Quá Tốc Độ (115%)
              </span>
              <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                !liveStats.overspeedPass ? 'bg-rose-500 text-white animate-pulse' : 'bg-emerald-500/20 text-emerald-400'
              }`}>
                {!liveStats.overspeedPass ? 'CRITICAL FAIL' : 'AN TOÀN'}
              </span>
            </div>
            <div className="text-lg font-bold font-mono truncate">
              {liveStats.overspeedPass ? 'Cắt Tức Thì 1725 RPM' : 'KHÔNG TẮT MÁY!'}
            </div>
            <div className="text-[11px] text-slate-400 truncate mt-0.5">
              NFPA 110 Sec 5.6.5.6
            </div>
          </div>

          {/* Stator Polarization Index PI Card */}
          <div className={`p-3.5 rounded-xl border ${
            !liveStats.piPass
              ? 'bg-amber-950/60 border-amber-500/70 text-amber-200'
              : 'bg-slate-800/60 border-slate-700 text-slate-200'
          }`}>
            <div className="flex items-center justify-between text-xs mb-1">
              <span className="flex items-center gap-1 font-medium">
                <Gauge className="w-3.5 h-3.5 text-amber-400" />
                Chỉ Số Phân Cực PI
              </span>
              <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                !liveStats.piPass ? 'bg-amber-500 text-black' : 'bg-emerald-500/20 text-emerald-400'
              }`}>
                {liveStats.pi < 2.0 ? 'ẨM / BẨN' : 'ĐẠT'}
              </span>
            </div>
            <div className="text-xl font-bold font-mono">
              PI = {liveStats.pi}
            </div>
            <div className="text-[11px] text-slate-400 truncate mt-0.5 flex items-center gap-1">
              {liveStats.piTrend < 0 ? (
                <span className="text-rose-400 flex items-center">
                  <TrendingDown className="w-3 h-3" /> {liveStats.piTrend} (vs năm trước)
                </span>
              ) : (
                <span>Table 100.11: PI ≥ 2.0</span>
              )}
            </div>
          </div>

          {/* NFPA 110 Class 10 Card */}
          <div className={`p-3.5 rounded-xl border ${
            !liveStats.class10Pass
              ? 'bg-rose-950/60 border-rose-500/60 text-rose-200'
              : 'bg-slate-800/60 border-slate-700 text-slate-200'
          }`}>
            <div className="flex items-center justify-between text-xs mb-1">
              <span className="flex items-center gap-1 font-medium">
                <Timer className="w-3.5 h-3.5 text-blue-400" />
                Thời Gian Nhận Tải
              </span>
              <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                liveStats.class10Pass ? 'bg-emerald-500/20 text-emerald-400' : 'bg-rose-500 text-white'
              }`}>
                {liveStats.class10Pass ? 'CLASS 10' : 'CHẬM'}
              </span>
            </div>
            <div className="text-xl font-bold font-mono">
              {liveStats.timeToLoad} giây
            </div>
            <div className="text-[11px] text-slate-400 truncate mt-0.5">
              NFPA 110 Bắt Buộc ≤ 10s
            </div>
          </div>

          {/* Governor & AVR Card */}
          <div className={`p-3.5 rounded-xl border ${
            !liveStats.regPass
              ? 'bg-amber-950/60 border-amber-500/60 text-amber-200'
              : 'bg-slate-800/60 border-slate-700 text-slate-200'
          }`}>
            <div className="flex items-center justify-between text-xs mb-1">
              <span className="flex items-center gap-1 font-medium">
                <Sliders className="w-3.5 h-3.5 text-teal-400" />
                Ổn Định Điện Áp AVR
              </span>
              <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-500/20 text-emerald-400">
                ±{Math.abs(liveStats.vReg)}%
              </span>
            </div>
            <div className="text-xl font-bold font-mono">
              ±{Math.abs(liveStats.vReg)}%
            </div>
            <div className="text-[11px] text-slate-400 truncate mt-0.5">
              NETA ≤ ±1.0% | Droop {liveStats.fDroop}%
            </div>
          </div>

          {/* Vibration Card */}
          <div className={`col-span-2 md:col-span-1 p-3.5 rounded-xl border ${
            !liveStats.vibPass
              ? 'bg-amber-950/60 border-amber-500/60 text-amber-200'
              : 'bg-slate-800/60 border-slate-700 text-slate-200'
          }`}>
            <div className="flex items-center justify-between text-xs mb-1">
              <span className="flex items-center gap-1 font-medium">
                <Activity className="w-3.5 h-3.5 text-purple-400" />
                Rung Nắp Ổ Đỡ
              </span>
              <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-500/20 text-emerald-400">
                ÊM ÁI
              </span>
            </div>
            <div className="text-xl font-bold font-mono">
              {liveStats.vib} mils
            </div>
            <div className="text-[11px] text-slate-400 truncate mt-0.5">
              Giới hạn NSX ≤ 2.5 mils
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

      {/* Main Grid: Left & Right Columns */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Left Column: Generator Info & Visual Inspection */}
        <div className="space-y-6">
          {/* Section 1: Thông tin Máy phát */}
          <div className="bg-slate-900/80 border border-slate-800 rounded-2xl p-5 shadow-lg space-y-4">
            <h2 className="text-base font-bold text-white flex items-center gap-2 border-b border-slate-800 pb-3">
              <Power className="w-5 h-5 text-amber-400" />
              1. Thông Tin Tổ Máy Phát Điện (Engine Generator)
            </h2>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
              <div>
                <label className="block text-slate-400 mb-1 font-medium">Tên dự án / Bệnh viện</label>
                <input
                  type="text"
                  value={data.site_info.project_name}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, project_name: e.target.value }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-amber-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Mã Tag tổ máy</label>
                <input
                  type="text"
                  value={data.site_info.generator_tag}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, generator_tag: e.target.value }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-amber-500"
                />
              </div>

              <div className="md:col-span-2">
                <label className="block text-slate-400 mb-1 font-medium">Thông số định mức tổ máy (kW / kVA / V / RPM)</label>
                <input
                  type="text"
                  value={data.site_info.generator_rating}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, generator_rating: e.target.value }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-amber-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Hãng động cơ Diesel</label>
                <input
                  type="text"
                  value={data.site_info.engine_manufacturer || 'Cummins / Perkins'}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, engine_manufacturer: e.target.value }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-amber-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Hãng đầu phát (Alternator)</label>
                <input
                  type="text"
                  value={data.site_info.alternator_manufacturer || 'Stamford / Leroy Somer'}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, alternator_manufacturer: e.target.value }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-amber-500"
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
                <label className="block text-slate-400 mb-1 font-medium">Ngày thử nghiệm</label>
                <input
                  type="date"
                  value={data.site_info.test_date}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, test_date: e.target.value }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-amber-500"
                />
              </div>
            </div>
          </div>

          {/* Section 2: Kiểm tra Thị giác & Cơ khí (NETA Sec 7.22.1.A) */}
          <div className="bg-slate-900/80 border border-slate-800 rounded-2xl p-5 shadow-lg space-y-4">
            <h2 className="text-base font-bold text-white flex items-center gap-2 border-b border-slate-800 pb-3">
              <Eye className="w-5 h-5 text-amber-400" />
              2. Kiểm Tra Thị Giác & Cơ Khí (NETA Sec 7.22.1.A)
            </h2>

            <div className="space-y-3 text-xs">
              <div className="flex items-center justify-between p-3 rounded-xl bg-slate-950/70 border border-slate-800">
                <div>
                  <div className="font-semibold text-white">Khung Kẹp Vận Chuyển (Shipping Bracing Removed)</div>
                  <div className="text-slate-400 text-[11px]">BẮT BUỘC tháo bỏ thanh giằng khóa lò xo giảm chấn trước khi chạy</div>
                </div>
                <button
                  type="button"
                  onClick={() =>
                    setData({
                      ...data,
                      visual_inspection: {
                        ...data.visual_inspection,
                        shipping_bracing_removed: !data.visual_inspection.shipping_bracing_removed
                      }
                    })
                  }
                  className={`px-3 py-1.5 rounded-lg font-bold text-xs transition-colors ${
                    data.visual_inspection.shipping_bracing_removed
                      ? 'bg-emerald-600/30 text-emerald-300 border border-emerald-500/50'
                      : 'bg-rose-600/30 text-rose-300 border border-rose-500/50'
                  }`}
                >
                  {data.visual_inspection.shipping_bracing_removed ? 'ĐÃ THÁO BỎ' : 'CHƯA THÁO!'}
                </button>
              </div>

              <div className="flex items-center justify-between p-3 rounded-xl bg-slate-950/70 border border-slate-800">
                <div>
                  <div className="font-semibold text-white">Đối Soát Nhãn Mác Kỹ Thuật (Nameplate Match)</div>
                  <div className="text-slate-400 text-[11px]">Khớp 100% bản vẽ thiết kế (kW/kVA, V, A, số pha, 1500 RPM)</div>
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
                  <label className="block text-slate-400 mb-1 font-medium">Tình trạng vật lý & Rò rỉ chất lỏng</label>
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
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-amber-500"
                  />
                </div>

                <div>
                  <label className="block text-slate-400 mb-1 font-medium">Định tâm trục & Tiếp địa vỏ máy</label>
                  <select
                    value={data.visual_inspection.anchorage_alignment_grounding}
                    onChange={(e: any) =>
                      setData({
                        ...data,
                        visual_inspection: {
                          ...data.visual_inspection,
                          anchorage_alignment_grounding: e.target.value
                        }
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-amber-500"
                  >
                    <option value="Pass">Pass - Đồng trục & tiếp địa tốt</option>
                    <option value="Investigate">Investigate - Cần cân chỉnh</option>
                    <option value="Fail">Fail - Lệch tâm / mất tiếp địa</option>
                  </select>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Right Column: Electrical Tests & Shutdown Protections */}
        <div className="space-y-6">
          {/* Section 3: Mạch Bảo Vệ Tắt Máy Khẩn Cấp (Engine Shutdowns) */}
          <div className={`border rounded-2xl p-5 shadow-lg space-y-4 ${
            !liveStats.overspeedPass
              ? 'bg-rose-950/30 border-rose-500/70'
              : 'bg-slate-900/80 border-slate-800'
          }`}>
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <h2 className="text-base font-bold text-white flex items-center gap-2">
                <AlertOctagon className="w-5 h-5 text-rose-400" />
                3. Mạch Bảo Vệ Tắt Máy Khẩn Cấp (Engine Shutdowns)
              </h2>
              <span className={`px-2.5 py-1 rounded-full text-xs font-bold ${
                !liveStats.overspeedPass
                  ? 'bg-rose-500 text-white animate-pulse'
                  : 'bg-emerald-500/20 text-emerald-400'
              }`}>
                {!liveStats.overspeedPass ? 'LỖI CẮT QUÁ TỐC ĐỘ' : 'BẢO VỆ ĐẠT'}
              </span>
            </div>

            <div className="space-y-3 text-xs">
              {/* Overspeed Trip (Highlighted) */}
              <div className={`p-3 rounded-xl border flex items-center justify-between ${
                !liveStats.overspeedPass
                  ? 'bg-rose-900/40 border-rose-500/70'
                  : 'bg-slate-950/70 border-slate-800'
              }`}>
                <div>
                  <div className="font-bold text-white flex items-center gap-1.5">
                    <span className="text-rose-400">●</span>
                    Quá Tốc Độ Động Cơ (Overspeed Trip - 115% Định Mức) *
                  </div>
                  <div className="text-slate-400 text-[11px]">
                    BẮT BUỘC tự ngắt khi vòng quay đạt 1725 RPM để chống nổ bánh đà
                  </div>
                </div>

                <select
                  value={data.electrical_tests.engine_protection_shutdowns_test.overspeed_trip.includes('Fail') ? 'Fail' : 'Pass'}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        engine_protection_shutdowns_test: {
                          ...data.electrical_tests.engine_protection_shutdowns_test,
                          overspeed_trip:
                            e.target.value === 'Fail'
                              ? 'Fail (Engine failed to trip at 115% rated RPM)'
                              : 'Pass (Engine tripped instantly at 115% rated RPM)'
                        }
                      }
                    })
                  }
                  className={`px-3 py-1.5 rounded-lg font-bold text-xs focus:outline-none ${
                    data.electrical_tests.engine_protection_shutdowns_test.overspeed_trip.includes('Fail')
                      ? 'bg-rose-600 text-white border border-rose-400'
                      : 'bg-emerald-600/30 text-emerald-300 border border-emerald-500'
                  }`}
                >
                  <option value="Pass">Pass - Tự ngắt máy tức thì</option>
                  <option value="Fail">Fail - Không ngắt máy!</option>
                </select>
              </div>

              {/* Low oil pressure */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                <div>
                  <label className="block text-slate-400 mb-1 font-medium">Bảo vệ Áp suất dầu bôi trơn thấp</label>
                  <input
                    type="text"
                    value={data.electrical_tests.engine_protection_shutdowns_test.low_oil_pressure_trip}
                    onChange={(e) =>
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          engine_protection_shutdowns_test: {
                            ...data.electrical_tests.engine_protection_shutdowns_test,
                            low_oil_pressure_trip: e.target.value
                          }
                        }
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-amber-500"
                  />
                </div>

                <div>
                  <label className="block text-slate-400 mb-1 font-medium">Bảo vệ Nhiệt độ nước làm mát cao</label>
                  <input
                    type="text"
                    value={data.electrical_tests.engine_protection_shutdowns_test.high_coolant_temperature_trip}
                    onChange={(e) =>
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          engine_protection_shutdowns_test: {
                            ...data.electrical_tests.engine_protection_shutdowns_test,
                            high_coolant_temperature_trip: e.target.value
                          }
                        }
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-amber-500"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Section 4: Điện trở Cách điện Đầu phát & Chỉ số Phân cực (IEEE 43) */}
          <div className={`border rounded-2xl p-5 shadow-lg space-y-4 ${
            !liveStats.piPass
              ? 'bg-amber-950/20 border-amber-500/60'
              : 'bg-slate-900/80 border-slate-800'
          }`}>
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <h2 className="text-base font-bold text-white flex items-center gap-2">
                <Gauge className="w-5 h-5 text-amber-400" />
                4. Điện Trở Cách Điện & Chỉ Số Phân Cực PI (IEEE 43)
              </h2>
              <span className={`px-2.5 py-1 rounded-full text-xs font-bold ${
                !liveStats.piPass
                  ? 'bg-amber-500 text-black animate-pulse'
                  : 'bg-emerald-500/20 text-emerald-400'
              }`}>
                {!liveStats.piPass ? `PI = ${liveStats.pi} < 2.0 (ẨM/BẨN)` : 'PI ĐẠT CHUẨN'}
              </span>
            </div>

            <div className="grid grid-cols-3 gap-3 text-xs">
              <div>
                <label className="block text-slate-400 mb-1 font-medium">IR 1 phút Stator 1000V (MΩ)</label>
                <input
                  type="number"
                  step="1"
                  value={data.electrical_tests.generator_insulation_resistance_40c_megohms.stator_1min_ir_1000v}
                  onChange={(e) => {
                    const v1 = parseFloat(e.target.value) || 0;
                    const v10 = data.electrical_tests.generator_insulation_resistance_40c_megohms.stator_10min_ir_1000v;
                    const newPi = v1 > 0 ? Math.round((v10 / v1) * 10) / 10 : 0;
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        generator_insulation_resistance_40c_megohms: {
                          ...data.electrical_tests.generator_insulation_resistance_40c_megohms,
                          stator_1min_ir_1000v: v1,
                          calculated_pi_value: newPi
                        }
                      }
                    });
                  }}
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-2 text-white font-mono focus:outline-none focus:border-amber-500"
                />
                <span className="text-[10px] text-slate-400 mt-1 block">Table 100.11: ≥ 100 MΩ</span>
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">IR 10 phút Stator 1000V (MΩ)</label>
                <input
                  type="number"
                  step="1"
                  value={data.electrical_tests.generator_insulation_resistance_40c_megohms.stator_10min_ir_1000v}
                  onChange={(e) => {
                    const v10 = parseFloat(e.target.value) || 0;
                    const v1 = data.electrical_tests.generator_insulation_resistance_40c_megohms.stator_1min_ir_1000v;
                    const newPi = v1 > 0 ? Math.round((v10 / v1) * 10) / 10 : 0;
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        generator_insulation_resistance_40c_megohms: {
                          ...data.electrical_tests.generator_insulation_resistance_40c_megohms,
                          stator_10min_ir_1000v: v10,
                          calculated_pi_value: newPi
                        }
                      }
                    });
                  }}
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-2 text-white font-mono focus:outline-none focus:border-amber-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-bold text-amber-300">Chỉ số Phân cực (PI) *</label>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.generator_insulation_resistance_40c_megohms.calculated_pi_value}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        generator_insulation_resistance_40c_megohms: {
                          ...data.electrical_tests.generator_insulation_resistance_40c_megohms,
                          calculated_pi_value: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className={`w-full bg-slate-950 border rounded-lg px-2.5 py-2 font-mono font-bold focus:outline-none ${
                    data.electrical_tests.generator_insulation_resistance_40c_megohms.calculated_pi_value < 2.0
                      ? 'border-amber-500 text-amber-300'
                      : 'border-slate-700 text-white'
                  }`}
                />
                <span className="text-[10px] text-slate-400 mt-1 block">BẮT BUỘC PI ≥ 2.0</span>
              </div>
            </div>
          </div>

          {/* Section 5: Thử Nghiệm Tải NFPA 110 & Độ Ổn Định AVR */}
          <div className="bg-slate-900/80 border border-slate-800 rounded-2xl p-5 shadow-lg space-y-4">
            <h2 className="text-base font-bold text-white flex items-center gap-2 border-b border-slate-800 pb-3">
              <Timer className="w-5 h-5 text-amber-400" />
              5. Thử Nghiệm Tải NFPA 110 & Điều Tốc / Điều Áp AVR
            </h2>

            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 text-xs">
              <div>
                <label className="block text-slate-400 mb-1 font-medium">Thời gian nhận tải (s)</label>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.nfpa_110_performance_load_test.time_to_first_load_acceptance_sec}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        nfpa_110_performance_load_test: {
                          ...data.electrical_tests.nfpa_110_performance_load_test,
                          time_to_first_load_acceptance_sec: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-amber-500"
                />
                <span className="text-[10px] text-slate-400 mt-1 block">NFPA 110: ≤ 10.0s</span>
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Độ ổn định áp AVR (±%)</label>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.governor_and_avr_regulation.voltage_regulation_percent}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        governor_and_avr_regulation: {
                          ...data.electrical_tests.governor_and_avr_regulation,
                          voltage_regulation_percent: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-amber-500"
                />
                <span className="text-[10px] text-slate-400 mt-1 block">NETA: ≤ ±1.0%</span>
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Sụt tần số Droop (%)</label>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.governor_and_avr_regulation.frequency_droop_percent}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        governor_and_avr_regulation: {
                          ...data.electrical_tests.governor_and_avr_regulation,
                          frequency_droop_percent: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-amber-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Rung nắp ổ đỡ (mils)</label>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.bearing_cap_vibration_mils}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        bearing_cap_vibration_mils: parseFloat(e.target.value) || 0
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-amber-500"
                />
                <span className="text-[10px] text-slate-400 mt-1 block">NSX: ≤ 2.5 mils</span>
              </div>
            </div>

            <div className="pt-2">
              <label className="block text-slate-400 mb-1 text-xs font-medium">
                Kết quả thử tải giả 100% liên tục 2 giờ (Load Bank Test)
              </label>
              <input
                type="text"
                value={data.electrical_tests.nfpa_110_performance_load_test.load_bank_100_percent_test_2hrs}
                onChange={(e) =>
                  setData({
                    ...data,
                    electrical_tests: {
                      ...data.electrical_tests,
                      nfpa_110_performance_load_test: {
                        ...data.electrical_tests.nfpa_110_performance_load_test,
                        load_bank_100_percent_test_2hrs: e.target.value
                      }
                    }
                  })
                }
                className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white text-xs focus:outline-none focus:border-amber-500"
              />
            </div>
          </div>
        </div>
      </div>

      {/* Action Buttons Bar */}
      <div className="flex flex-wrap items-center justify-between gap-4 p-5 bg-slate-900/90 border border-slate-800 rounded-2xl shadow-xl">
        <div className="flex items-center gap-2 text-xs text-slate-400">
          <Clock className="w-4 h-4 text-amber-400" />
          <span>Đối soát chuẩn xác theo ANSI/NETA ATS-2025 Mục 7.22.1 & NFPA 110 (Emergency Generators)</span>
        </div>

        <div className="flex flex-wrap items-center gap-3">
          <button
            onClick={handleRunAiAnalysis}
            disabled={isAnalyzing}
            className="px-5 py-2.5 rounded-xl bg-gradient-to-r from-red-600 via-amber-600 to-yellow-600 hover:from-red-500 hover:to-yellow-500 text-white font-semibold text-sm flex items-center gap-2 shadow-lg shadow-amber-500/20 disabled:opacity-50 transition-all cursor-pointer"
          >
            {isAnalyzing ? (
              <>
                <RotateCcw className="w-4 h-4 animate-spin" />
                <span>AI Đang Phân Tích & Đối Soát...</span>
              </>
            ) : (
              <>
                <Sparkles className="w-4 h-4 text-amber-200" />
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
                <Printer className="w-4 h-4 text-amber-400" />
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
                  Kết Luận Thẩm Định NETA ATS-2025 Mục 7.22.1 & NFPA 110 (Emergency Engine Generator)
                </div>
                <h3 className={`text-2xl font-bold mt-0.5 ${
                  analysisResult.overall_status === 'FAIL'
                    ? 'text-rose-400'
                    : analysisResult.overall_status === 'INVESTIGATE'
                    ? 'text-amber-400'
                    : 'text-emerald-400'
                }`}>
                  {analysisResult.overall_status === 'FAIL'
                    ? 'FAIL / CRITICAL OVERSPEED & INSULATION DEFECT'
                    : analysisResult.overall_status === 'INVESTIGATE'
                    ? 'INVESTIGATE / MONITORING REQUIRED'
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
                <FileText className="w-5 h-5 text-amber-400" />
                Báo Cáo Thẩm Định Chuyên Gia NETA ATS-2025 Mục 7.22.1 & NFPA 110
              </h3>
              <span className="text-xs text-slate-400 font-mono">
                Model: Gemini 2.5 Flash / TEV Engine Generator AI
              </span>
            </div>

            <div className="prose prose-invert max-w-none text-slate-200 text-sm leading-relaxed whitespace-pre-line bg-slate-950/60 p-5 rounded-xl border border-slate-800 font-sans">
              {analysisResult.markdown_report}
            </div>
          </div>
        </div>
      )}

      {/* JSON Modal */}
      {showJsonModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm">
          <div className="bg-slate-900 border border-slate-700 rounded-2xl max-w-2xl w-full p-6 space-y-4 shadow-2xl">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <h3 className="text-base font-bold text-white flex items-center gap-2">
                <Upload className="w-5 h-5 text-amber-400" />
                Dữ Liệu JSON Kiểm Định Tổ Máy Phát Điện
              </h3>
              <button
                onClick={() => setShowJsonModal(false)}
                className="text-slate-400 hover:text-white"
              >
                ✕
              </button>
            </div>

            <p className="text-xs text-slate-400">
              Dán trực tiếp JSON kiểm định ngoài hiện trường để nạp nhanh vào biểu mẫu máy phát điện.
            </p>

            <textarea
              rows={14}
              value={jsonInput}
              onChange={(e) => setJsonInput(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl p-3 font-mono text-xs text-amber-300 focus:outline-none focus:border-amber-500"
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
                className="px-4 py-2 rounded-xl bg-amber-600 hover:bg-amber-500 text-white text-xs font-semibold"
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
          equipmentCategory="motor"
        />
      )}

      {/* Print Report Modal */}
      {showPrintModal && analysisResult && (
        <TevReportPrintModal
          isOpen={showPrintModal}
          onClose={() => setShowPrintModal(false)}
          reportTitle={`BIÊN BẢN KIỂM ĐỊNH NETA ATS-2025 - TỔ MÁY PHÁT ĐIỆN ${data.site_info.generator_tag}`}
          equipmentTag={data.site_info.generator_tag}
          equipmentType="Tổ Máy Phát Điện Khẩn Cấp (Engine Generator)"
          testStandard="ANSI/NETA ATS-2025 Mục 7.22.1 & NFPA 110"
          overallStatus={analysisResult.overall_status}
          markdownReport={analysisResult.markdown_report}
        />
      )}
    </div>
  );
};
