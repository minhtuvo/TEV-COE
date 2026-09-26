import React, { useState, useMemo } from 'react';
import {
  BatteryCharging,
  Zap,
  CheckCircle2,
  AlertTriangle,
  FileText,
  Printer,
  Sparkles,
  Save,
  RotateCcw,
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
  Thermometer,
  Scale,
  Copy,
  Info,
  Clock
} from 'lucide-react';
import {
  NetaBatteryVrlaInputPayload,
  NetaBatteryVrlaEvaluationResult,
  performDeterministicVrlaAnalysis
} from '../../server/netaBatteryVrlaAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';
import { NameplateScannerModal } from './NameplateScannerModal';
import { ExtractedNameplateData } from '../../server/nameplateOcrAnalyzer';
import { FseEngineerDropdown } from './FseEngineerDropdown';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

export const SAMPLE_BATTERY_VRLA_FSE_DATA: NetaBatteryVrlaInputPayload = {
  site_info: {
    project_name: "Tram Bien Ap TEV 110kV UPS Room",
    battery_bank_tag: "BAT-VRLA-220V-01",
    battery_type: "Valve-Regulated Lead-Acid (18 Monoblocks 12V 100Ah VRLA AGM)",
    fse_name: "Nguyen Van Q",
    test_date: "2026-09-24",
    substation_location: "Phòng UPS Trạm 110kV",
    rated_voltage_vdc: 220,
    rated_capacity_ah: 100,
    ambient_temp_c: 25.0
  },
  visual_inspection: {
    ventilation_system_operable: true,
    eyewash_station_present: true,
    nameplate_match: true,
    cabinet_mounting_and_grounding: "Pass",
    cleanliness_and_oxide_inhibitor: "Pass",
    physical_condition: "Good (No case swelling or acid leaks)",
    bolt_torque_check: "Pass",
    thermographic_survey: "Pass"
  },
  electrical_tests: {
    negative_post_temperature_c: {
      avg_temp_c: 26.5,
      max_temp_monoblock_12_c: 41.2,
      hot_monoblock_tag: "Monoblock số 12"
    },
    intercell_connection_resistance_micro_ohms: [14.1, 14.3, 14.2, 14.5],
    charger_float_voltage_v: 243.0,
    charger_equalize_voltage_v: 252.0,
    charger_alarms_check: "Pass",
    monoblock_voltages_float_mode_v: {
      min_voltage_v: 13.10,
      max_voltage_v: 13.62,
      avg_voltage_v: 13.50
    },
    internal_ohmic_measurement_resistance_mohm: {
      avg_resistance_mohm: 3.20,
      max_resistance_monoblock_12_mohm: 4.65,
      variance_percent: 45.3,
      suspect_monoblock_tag: "Monoblock số 12"
    },
    load_test_ieee_1188: "Pass (Capacity 88% per IEEE 1188)",
    system_voltage_to_ground_v: {
      positive_to_ground_v: 121.2,
      negative_to_ground_v: -121.8
    }
  },
  previous_test_data: {
    last_test_date: "2025-09-22",
    last_avg_resistance_mohm: 3.15,
    last_monoblock_12_temp_c: 26.0
  }
};

export const SAMPLE_BATTERY_VRLA_PASS_DATA: NetaBatteryVrlaInputPayload = {
  site_info: {
    project_name: "Trung Tâm Dữ Liệu TEV Data Center",
    battery_bank_tag: "BAT-VRLA-220V-02",
    battery_type: "Valve-Regulated Lead-Acid (18 Monoblocks 12V 150Ah VRLA GEL)",
    fse_name: "Tran Van T",
    test_date: "2026-09-24",
    substation_location: "Phòng Nguồn DC Tầng 2",
    rated_voltage_vdc: 220,
    rated_capacity_ah: 150,
    ambient_temp_c: 24.5
  },
  visual_inspection: {
    ventilation_system_operable: true,
    eyewash_station_present: true,
    nameplate_match: true,
    cabinet_mounting_and_grounding: "Pass",
    cleanliness_and_oxide_inhibitor: "Pass",
    physical_condition: "Hoàn hảo (Không phồng rộp, nắp van kín)",
    bolt_torque_check: "Pass",
    thermographic_survey: "Pass"
  },
  electrical_tests: {
    negative_post_temperature_c: {
      avg_temp_c: 25.2,
      max_temp_monoblock_12_c: 25.9,
      hot_monoblock_tag: "Bình số 7"
    },
    intercell_connection_resistance_micro_ohms: [13.8, 14.0, 13.9, 14.1],
    charger_float_voltage_v: 243.0,
    charger_equalize_voltage_v: 252.0,
    charger_alarms_check: "Pass",
    monoblock_voltages_float_mode_v: {
      min_voltage_v: 13.48,
      max_voltage_v: 13.56,
      avg_voltage_v: 13.52
    },
    internal_ohmic_measurement_resistance_mohm: {
      avg_resistance_mohm: 2.85,
      max_resistance_monoblock_12_mohm: 3.12,
      variance_percent: 9.5,
      suspect_monoblock_tag: "Không có"
    },
    load_test_ieee_1188: "Pass (Capacity 96% per IEEE 1188)",
    system_voltage_to_ground_v: {
      positive_to_ground_v: 121.5,
      negative_to_ground_v: -121.5
    }
  },
  previous_test_data: {
    last_test_date: "2025-09-18",
    last_avg_resistance_mohm: 2.80,
    last_monoblock_12_temp_c: 25.0
  }
};

export const NetaBatteryVrlaChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaBatteryVrlaInputPayload>(SAMPLE_BATTERY_VRLA_FSE_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaBatteryVrlaEvaluationResult | null>(null);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showCameraModal, setShowCameraModal] = useState(false);
  const [jsonInput, setJsonInput] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live real-time calculations
  const liveStats = useMemo(() => {
    // 1. Negative Post Temperature & Thermal Runaway check
    const avgTemp = data.electrical_tests.negative_post_temperature_c.avg_temp_c || 25.0;
    const maxTemp =
      data.electrical_tests.negative_post_temperature_c.max_temp_monoblock_12_c ??
      data.electrical_tests.negative_post_temperature_c.max_temp_c ??
      avgTemp;
    const deltaT = Math.round((maxTemp - avgTemp) * 10) / 10;
    let thermalRisk: 'LOW' | 'MODERATE' | 'CRITICAL' = 'LOW';
    if (deltaT >= 5.0 || maxTemp >= 40.0) {
      thermalRisk = 'CRITICAL';
    } else if (deltaT >= 3.0 || maxTemp >= 35.0) {
      thermalRisk = 'MODERATE';
    }

    // 2. Intercell connection resistance
    const resistances = data.electrical_tests.intercell_connection_resistance_micro_ohms || [];
    let minRes = 0;
    let maxRes = 0;
    let intercellDev = 0;
    let intercellPass = true;
    if (resistances.length > 0) {
      minRes = Math.min(...resistances);
      maxRes = Math.max(...resistances);
      if (minRes > 0) {
        intercellDev = Math.round(((maxRes - minRes) / minRes) * 1000) / 10;
        intercellPass = intercellDev <= 50;
      }
    }

    // 3. Internal Ohmic variance
    const ohmicVar = data.electrical_tests.internal_ohmic_measurement_resistance_mohm.variance_percent;
    const ohmicPass = ohmicVar <= 25;

    // 4. Monoblock voltage spread
    const minV = data.electrical_tests.monoblock_voltages_float_mode_v.min_voltage_v;
    const maxV = data.electrical_tests.monoblock_voltages_float_mode_v.max_voltage_v;
    const spreadV = Math.round((maxV - minV) * 1000) / 1000;
    const voltagePass = spreadV <= 0.6;

    // 5. System voltage to ground symmetry
    const vPos = Math.abs(data.electrical_tests.system_voltage_to_ground_v.positive_to_ground_v);
    const vNeg = Math.abs(data.electrical_tests.system_voltage_to_ground_v.negative_to_ground_v);
    const groundDiff = Math.round(Math.abs(vPos - vNeg) * 10) / 10;
    let groundSymmetryStatus = 'Cân bằng (|V+| ≈ |V-|)';
    let isGroundFault = false;
    if (groundDiff > 25) {
      isGroundFault = true;
      groundSymmetryStatus = vPos < vNeg ? 'CẢNH BÁO: Rò/Chạm đất Cực Dương (+)' : 'CẢNH BÁO: Rò/Chạm đất Cực Âm (-)';
    } else if (groundDiff > 8) {
      groundSymmetryStatus = 'Cảnh báo: Lệch điện áp đất';
    }

    return {
      avgTemp,
      maxTemp,
      deltaT,
      thermalRisk,
      intercellDev,
      intercellPass,
      minRes,
      maxRes,
      ohmicVar,
      ohmicPass,
      spreadV,
      voltagePass,
      vPos,
      vNeg,
      groundDiff,
      groundSymmetryStatus,
      isGroundFault
    };
  }, [data]);

  const handleRunAiAnalysis = async () => {
    setIsAnalyzing(true);
    setSavedSuccessMessage(null);
    try {
      const response = await fetch('/api/field-service/neta-battery-vrla-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });
      const result = await response.json();
      if (result.success && result.data) {
        setAnalysisResult(result.data);
      } else {
        const fallback = performDeterministicVrlaAnalysis(data);
        setAnalysisResult(fallback);
      }
    } catch (e) {
      console.error('Error calling VRLA AI analysis API:', e);
      const fallback = performDeterministicVrlaAnalysis(data);
      setAnalysisResult(fallback);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleSaveToReports = () => {
    if (!analysisResult) return;
    if (onSaveReport) {
      onSaveReport({
        payload_type: 'NETA_ATS_2025_BATTERY_VRLA',
        equipment_type: 'Valve-Regulated Lead-Acid (VRLA) Battery Bank',
        battery_bank_tag: data.site_info.battery_bank_tag,
        project_name: data.site_info.project_name,
        test_standard: 'ANSI/NETA ATS-2025 Mục 7.18.1.3 & IEEE 1188',
        analysis_result: analysisResult,
        neta_data: data
      });
      setSavedSuccessMessage(`Đã lưu biên bản giàn pin VRLA ${data.site_info.battery_bank_tag} vào kho Báo cáo kỹ thuật thành công!`);
      setTimeout(() => setSavedSuccessMessage(null), 6000);
    }
  };

  const handleApplyOcrData = (ocr: ExtractedNameplateData) => {
    setData((prev) => ({
      ...prev,
      site_info: {
        ...prev.site_info,
        battery_bank_tag: ocr.equipment_tag || ocr.serial_number || prev.site_info.battery_bank_tag,
        battery_type: ocr.equipment_type || ocr.standards || prev.site_info.battery_type,
        rated_voltage_vdc: ocr.secondary_voltage_v ?? prev.site_info.rated_voltage_vdc
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
      <div className="bg-gradient-to-r from-teal-900 via-emerald-950 to-slate-900 border border-teal-500/30 rounded-2xl p-6 shadow-xl relative overflow-hidden">
        <div className="absolute top-0 right-0 w-96 h-96 bg-teal-500/10 rounded-full blur-3xl pointer-events-none" />
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 relative z-10">
          <div>
            <div className="flex items-center gap-2 text-teal-400 font-mono text-xs uppercase tracking-wider mb-2">
              <span className="w-2 h-2 rounded-full bg-teal-400 animate-pulse" />
              ANSI/NETA ATS-2025 Section 7.18.1.3 & IEEE 1188
            </div>
            <h1 className="text-2xl lg:text-3xl font-bold text-white flex items-center gap-3">
              <BatteryCharging className="w-8 h-8 text-teal-400" />
              Ắc Quy Chì-Axit Kín Khí (VRLA - Section 7.18.1.3)
            </h1>
            <p className="text-slate-300 text-sm mt-1 max-w-2xl">
              Hệ thống Điện một chiều DC và Ắc quy Chì-Axit van điều áp kín khí (VRLA AGM/GEL). Giám sát trôi nhiệt (Thermal Runaway), nội trở Ôm per NETA 7.18.1.3.D.7, độ lệch cầu nối & điện áp đối đất.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => {
                setData(SAMPLE_BATTERY_VRLA_FSE_DATA);
                setAnalysisResult(null);
              }}
              className="px-3.5 py-2 rounded-xl bg-rose-950/60 hover:bg-rose-900/70 border border-rose-500/40 text-rose-300 text-xs font-medium flex items-center gap-1.5 transition-colors shadow-sm"
              title="Nạp mẫu mô phỏng sự cố trôi nhiệt và chai nội trở bình số 12"
            >
              <Flame className="w-3.5 h-3.5 text-rose-400" />
              Nạp Mẫu Lỗi (Thermal Runaway & Ohmic)
            </button>
            <button
              onClick={() => {
                setData(SAMPLE_BATTERY_VRLA_PASS_DATA);
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
              <Camera className="w-3.5 h-3.5 text-teal-400" />
              Quét Nhãn AI
            </button>
          </div>
        </div>

        {/* Live Telemetry KPI Cards */}
        <div className="grid grid-cols-2 md:grid-cols-5 gap-3 mt-6 pt-6 border-t border-slate-700/50">
          {/* Thermal Runaway Card */}
          <div className={`p-3.5 rounded-xl border ${
            liveStats.thermalRisk === 'CRITICAL'
              ? 'bg-rose-950/50 border-rose-500/60 text-rose-200'
              : liveStats.thermalRisk === 'MODERATE'
              ? 'bg-amber-950/50 border-amber-500/60 text-amber-200'
              : 'bg-slate-800/60 border-slate-700 text-slate-200'
          }`}>
            <div className="flex items-center justify-between text-xs mb-1">
              <span className="flex items-center gap-1 font-medium">
                <Thermometer className="w-3.5 h-3.5 text-rose-400" />
                Cực Âm / Trôi Nhiệt
              </span>
              <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                liveStats.thermalRisk === 'CRITICAL'
                  ? 'bg-rose-500 text-white animate-pulse'
                  : liveStats.thermalRisk === 'MODERATE'
                  ? 'bg-amber-500 text-black'
                  : 'bg-emerald-500/20 text-emerald-400'
              }`}>
                {liveStats.thermalRisk}
              </span>
            </div>
            <div className="text-xl font-bold font-mono">
              {liveStats.maxTemp}°C
              <span className="text-xs font-normal text-slate-400 ml-1">
                (ΔT +{liveStats.deltaT}°C)
              </span>
            </div>
            <div className="text-[11px] text-slate-400 truncate mt-0.5">
              Avg: {liveStats.avgTemp}°C per IEEE 1188
            </div>
          </div>

          {/* Internal Ohmic Variance Card */}
          <div className={`p-3.5 rounded-xl border ${
            !liveStats.ohmicPass
              ? 'bg-rose-950/50 border-rose-500/60 text-rose-200'
              : 'bg-slate-800/60 border-slate-700 text-slate-200'
          }`}>
            <div className="flex items-center justify-between text-xs mb-1">
              <span className="flex items-center gap-1 font-medium">
                <Gauge className="w-3.5 h-3.5 text-teal-400" />
                Biến Động Nội Trở
              </span>
              <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                !liveStats.ohmicPass ? 'bg-rose-500 text-white' : 'bg-emerald-500/20 text-emerald-400'
              }`}>
                {liveStats.ohmicVar}%
              </span>
            </div>
            <div className="text-xl font-bold font-mono">
              {data.electrical_tests.internal_ohmic_measurement_resistance_mohm.max_resistance_monoblock_12_mohm || 4.65} mΩ
            </div>
            <div className="text-[11px] text-slate-400 truncate mt-0.5">
              NETA ≤ 25% (Thực tế: {liveStats.ohmicVar}%)
            </div>
          </div>

          {/* Intercell Connection Resistance Deviation Card */}
          <div className={`p-3.5 rounded-xl border ${
            !liveStats.intercellPass
              ? 'bg-amber-950/50 border-amber-500/60 text-amber-200'
              : 'bg-slate-800/60 border-slate-700 text-slate-200'
          }`}>
            <div className="flex items-center justify-between text-xs mb-1">
              <span className="flex items-center gap-1 font-medium">
                <Sliders className="w-3.5 h-3.5 text-amber-400" />
                Cầu Nối Intercell
              </span>
              <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                !liveStats.intercellPass ? 'bg-amber-500 text-black' : 'bg-emerald-500/20 text-emerald-400'
              }`}>
                {liveStats.intercellDev}%
              </span>
            </div>
            <div className="text-xl font-bold font-mono">
              {liveStats.maxRes} µΩ
            </div>
            <div className="text-[11px] text-slate-400 truncate mt-0.5">
              NETA ≤ 50% deviation
            </div>
          </div>

          {/* Monoblock Voltage Spread Card */}
          <div className={`p-3.5 rounded-xl border ${
            !liveStats.voltagePass
              ? 'bg-amber-950/50 border-amber-500/60 text-amber-200'
              : 'bg-slate-800/60 border-slate-700 text-slate-200'
          }`}>
            <div className="flex items-center justify-between text-xs mb-1">
              <span className="flex items-center gap-1 font-medium">
                <Power className="w-3.5 h-3.5 text-teal-400" />
                Lệch Áp Monoblock
              </span>
              <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                !liveStats.voltagePass ? 'bg-amber-500 text-black' : 'bg-emerald-500/20 text-emerald-400'
              }`}>
                Δ {liveStats.spreadV} V
              </span>
            </div>
            <div className="text-xl font-bold font-mono">
              {data.electrical_tests.monoblock_voltages_float_mode_v.avg_voltage_v} V
            </div>
            <div className="text-[11px] text-slate-400 truncate mt-0.5">
              Min: {data.electrical_tests.monoblock_voltages_float_mode_v.min_voltage_v}V - Max: {data.electrical_tests.monoblock_voltages_float_mode_v.max_voltage_v}V
            </div>
          </div>

          {/* DC Ground Voltage Symmetry Card */}
          <div className={`col-span-2 md:col-span-1 p-3.5 rounded-xl border ${
            liveStats.isGroundFault
              ? 'bg-rose-950/60 border-rose-500/60 text-rose-200'
              : 'bg-slate-800/60 border-slate-700 text-slate-200'
          }`}>
            <div className="flex items-center justify-between text-xs mb-1">
              <span className="flex items-center gap-1 font-medium">
                <Scale className="w-3.5 h-3.5 text-cyan-400" />
                Điện Áp Đối Đất
              </span>
              <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                liveStats.isGroundFault ? 'bg-rose-500 text-white' : 'bg-emerald-500/20 text-emerald-400'
              }`}>
                Δ {liveStats.groundDiff}V
              </span>
            </div>
            <div className="text-base font-bold font-mono text-cyan-300">
              +{liveStats.vPos}V / -{liveStats.vNeg}V
            </div>
            <div className="text-[11px] text-slate-400 truncate mt-0.5" title={liveStats.groundSymmetryStatus}>
              {liveStats.groundSymmetryStatus}
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
        {/* Left Column: Site Info, Visual & Mechanical Inspection */}
        <div className="space-y-6">
          {/* Section: Thông tin giàn pin VRLA */}
          <div className="bg-slate-900/80 border border-slate-800 rounded-2xl p-5 shadow-lg space-y-4">
            <h2 className="text-base font-bold text-white flex items-center gap-2 border-b border-slate-800 pb-3">
              <BatteryCharging className="w-5 h-5 text-teal-400" />
              1. Thông Tin Giàn Pin VRLA & Trạm Nguồn
            </h2>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
              <div>
                <label className="block text-slate-400 mb-1 font-medium">Tên dự án / Trạm</label>
                <input
                  type="text"
                  value={data.site_info.project_name}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, project_name: e.target.value }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-teal-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Mã Tag giàn pin VRLA</label>
                <input
                  type="text"
                  value={data.site_info.battery_bank_tag}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, battery_bank_tag: e.target.value }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-teal-500"
                />
              </div>

              <div className="md:col-span-2">
                <label className="block text-slate-400 mb-1 font-medium">Chủng loại ắc quy (Monoblocks & Công nghệ)</label>
                <input
                  type="text"
                  value={data.site_info.battery_type}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, battery_type: e.target.value }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-teal-500"
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
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-teal-500"
                />
              </div>
            </div>
          </div>

          {/* Section: Kiểm tra Thị giác & Cơ khí (NETA ATS-2025 Section 7.18.1.3.A) */}
          <div className="bg-slate-900/80 border border-slate-800 rounded-2xl p-5 shadow-lg space-y-4">
            <h2 className="text-base font-bold text-white flex items-center gap-2 border-b border-slate-800 pb-3">
              <Eye className="w-5 h-5 text-teal-400" />
              2. Kiểm Tra Thị Giác & Cơ Khí (NETA ATS 7.18.1.3.A)
            </h2>

            <div className="space-y-3 text-xs">
              <div className="flex items-center justify-between p-3 rounded-xl bg-slate-950/70 border border-slate-800">
                <div>
                  <div className="font-semibold text-white">Hệ thống Thông gió Phòng Tủ VRLA</div>
                  <div className="text-slate-400 text-[11px]">Bắt buộc thông gió bình thường, tản nhiệt và thoát khí</div>
                </div>
                <button
                  type="button"
                  onClick={() =>
                    setData({
                      ...data,
                      visual_inspection: {
                        ...data.visual_inspection,
                        ventilation_system_operable: !data.visual_inspection.ventilation_system_operable
                      }
                    })
                  }
                  className={`px-3 py-1.5 rounded-lg font-bold text-xs transition-colors ${
                    data.visual_inspection.ventilation_system_operable
                      ? 'bg-emerald-600/30 text-emerald-300 border border-emerald-500/50'
                      : 'bg-rose-600/30 text-rose-300 border border-rose-500/50'
                  }`}
                >
                  {data.visual_inspection.ventilation_system_operable ? 'ĐẠT (Operable)' : 'LỖI (Inoperable)'}
                </button>
              </div>

              <div className="flex items-center justify-between p-3 rounded-xl bg-slate-950/70 border border-slate-800">
                <div>
                  <div className="font-semibold text-white">Trang bị Trạm Rửa Mắt Khẩn Cấp (Eyewash)</div>
                  <div className="text-slate-400 text-[11px]">Có sẵn tại khu vực giàn/tủ pin VRLA</div>
                </div>
                <button
                  type="button"
                  onClick={() =>
                    setData({
                      ...data,
                      visual_inspection: {
                        ...data.visual_inspection,
                        eyewash_station_present: !data.visual_inspection.eyewash_station_present
                      }
                    })
                  }
                  className={`px-3 py-1.5 rounded-lg font-bold text-xs transition-colors ${
                    data.visual_inspection.eyewash_station_present
                      ? 'bg-emerald-600/30 text-emerald-300 border border-emerald-500/50'
                      : 'bg-amber-600/30 text-amber-300 border border-amber-500/50'
                  }`}
                >
                  {data.visual_inspection.eyewash_station_present ? 'SẴN SÀNG' : 'CHƯA TRANG BỊ'}
                </button>
              </div>

              <div className="flex items-center justify-between p-3 rounded-xl bg-slate-950/70 border border-slate-800">
                <div>
                  <div className="font-semibold text-white">Đối Soát Nhãn Mác Thiết Bị (Nameplate Match)</div>
                  <div className="text-slate-400 text-[11px]">Khớp 100% hồ sơ thiết kế và thông số nhà sản xuất</div>
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

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Tình trạng Vỏ bình VRLA (Kiểm tra Phồng rộp / Swelling)</label>
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
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-teal-500"
                  placeholder="e.g. Good (No case swelling or acid leaks)"
                />
              </div>

              <div className="grid grid-cols-2 gap-3 pt-2">
                <div>
                  <label className="block text-slate-400 mb-1 font-medium">Khung tủ / Giá & Tiếp địa</label>
                  <select
                    value={data.visual_inspection.cabinet_mounting_and_grounding}
                    onChange={(e: any) =>
                      setData({
                        ...data,
                        visual_inspection: {
                          ...data.visual_inspection,
                          cabinet_mounting_and_grounding: e.target.value
                        }
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-teal-500"
                  >
                    <option value="Pass">Pass - Đạt tiếp địa</option>
                    <option value="Investigate">Investigate - Cần siết lại</option>
                    <option value="Fail">Fail - Mất tiếp địa</option>
                  </select>
                </div>

                <div>
                  <label className="block text-slate-400 mb-1 font-medium">Cọc cực & Mỡ chống oxy hóa</label>
                  <select
                    value={data.visual_inspection.cleanliness_and_oxide_inhibitor}
                    onChange={(e: any) =>
                      setData({
                        ...data,
                        visual_inspection: {
                          ...data.visual_inspection,
                          cleanliness_and_oxide_inhibitor: e.target.value
                        }
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-teal-500"
                  >
                    <option value="Pass">Pass - Sạch & Đã bôi mỡ</option>
                    <option value="Investigate">Investigate - Bụi bẩn cọc</option>
                    <option value="Fail">Fail - Oxy hóa nặng</option>
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
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-teal-500"
                  >
                    <option value="Pass">Pass - Đạt cờ-lê lực</option>
                    <option value="Investigate">Investigate - Lệch nhẹ</option>
                    <option value="Fail">Fail - Lỏng bu-lông</option>
                  </select>
                </div>

                <div>
                  <label className="block text-slate-400 mb-1 font-medium">Khảo sát nhiệt hồng ngoại</label>
                  <select
                    value={data.visual_inspection.thermographic_survey || 'Pass'}
                    onChange={(e: any) =>
                      setData({
                        ...data,
                        visual_inspection: {
                          ...data.visual_inspection,
                          thermographic_survey: e.target.value
                        }
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-teal-500"
                  >
                    <option value="Pass">Pass - Bình thường</option>
                    <option value="Investigate">Investigate - Cảnh báo</option>
                    <option value="Fail">Fail - Điểm nóng bất thường</option>
                  </select>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Right Column: Electrical Tests (NETA ATS-2025 Section 7.18.1.3.D) */}
        <div className="space-y-6">
          {/* Section: Nhiệt độ Cực Âm & Nguy cơ Trôi Nhiệt (IEEE 1188) */}
          <div className={`border rounded-2xl p-5 shadow-lg space-y-4 ${
            liveStats.thermalRisk === 'CRITICAL'
              ? 'bg-rose-950/30 border-rose-500/60'
              : 'bg-slate-900/80 border-slate-800'
          }`}>
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <h2 className="text-base font-bold text-white flex items-center gap-2">
                <Thermometer className="w-5 h-5 text-rose-400" />
                3. Nhiệt Độ Cực Âm & Trôi Nhiệt (IEEE 1188)
              </h2>
              <span className={`px-2.5 py-1 rounded-full text-xs font-bold ${
                liveStats.thermalRisk === 'CRITICAL'
                  ? 'bg-rose-500 text-white animate-pulse'
                  : liveStats.thermalRisk === 'MODERATE'
                  ? 'bg-amber-500 text-black'
                  : 'bg-emerald-500/20 text-emerald-400'
              }`}>
                {liveStats.thermalRisk === 'CRITICAL' ? 'CẢNH BÁO TRÔI NHIỆT (CRITICAL)' : `RỦI RO: ${liveStats.thermalRisk}`}
              </span>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-3 text-xs">
              <div>
                <label className="block text-slate-400 mb-1 font-medium">Nhiệt độ TB Giàn Pin (°C)</label>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.negative_post_temperature_c.avg_temp_c}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        negative_post_temperature_c: {
                          ...data.electrical_tests.negative_post_temperature_c,
                          avg_temp_c: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-teal-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Nhiệt độ Cực Âm Cao Nhất (°C)</label>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.negative_post_temperature_c.max_temp_monoblock_12_c ?? 41.2}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        negative_post_temperature_c: {
                          ...data.electrical_tests.negative_post_temperature_c,
                          max_temp_monoblock_12_c: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className={`w-full bg-slate-950 border rounded-lg px-3 py-2 text-white font-mono font-bold focus:outline-none ${
                    liveStats.maxTemp >= 40.0 || liveStats.deltaT >= 5.0
                      ? 'border-rose-500 text-rose-300'
                      : 'border-slate-700'
                  }`}
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Vị trí Bình Nóng Nhất</label>
                <input
                  type="text"
                  value={data.electrical_tests.negative_post_temperature_c.hot_monoblock_tag || 'Monoblock số 12'}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        negative_post_temperature_c: {
                          ...data.electrical_tests.negative_post_temperature_c,
                          hot_monoblock_tag: e.target.value
                        }
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-teal-500"
                />
              </div>
            </div>

            <div className="p-3 rounded-xl bg-slate-950/60 border border-slate-800 text-[11px] text-slate-300 flex items-start gap-2">
              <Info className="w-4 h-4 text-teal-400 flex-shrink-0 mt-0.5" />
              <div>
                <span className="font-semibold text-white">Chuẩn IEEE 1188:</span> Sự gia tăng nhiệt độ bất thường ở cực âm (Negative Post Temperature) là dấu hiệu điển hình của phản ứng tái hợp oxy tỏa nhiệt không kiểm soát (*Thermal Runaway*). Khi nhiệt độ vượt quá 35°C hoặc lệch trên 3.0°C so với trung bình giàn pin, FSE cần kiểm tra dòng float và xem xét thay bình.
              </div>
            </div>
          </div>

          {/* Section: Nội trở Ôm & Điện trở Cầu Nối Intercell */}
          <div className="bg-slate-900/80 border border-slate-800 rounded-2xl p-5 shadow-lg space-y-4">
            <h2 className="text-base font-bold text-white flex items-center gap-2 border-b border-slate-800 pb-3">
              <Gauge className="w-5 h-5 text-teal-400" />
              4. Nội Trở Ôm Cell & Cầu Nối Intercell
            </h2>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-3 text-xs">
              <div>
                <label className="block text-slate-400 mb-1 font-medium">Nội trở TB Giàn (mΩ)</label>
                <input
                  type="number"
                  step="0.01"
                  value={data.electrical_tests.internal_ohmic_measurement_resistance_mohm.avg_resistance_mohm}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        internal_ohmic_measurement_resistance_mohm: {
                          ...data.electrical_tests.internal_ohmic_measurement_resistance_mohm,
                          avg_resistance_mohm: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-teal-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Nội trở Bình Cao Nhất (mΩ)</label>
                <input
                  type="number"
                  step="0.01"
                  value={data.electrical_tests.internal_ohmic_measurement_resistance_mohm.max_resistance_monoblock_12_mohm ?? 4.65}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        internal_ohmic_measurement_resistance_mohm: {
                          ...data.electrical_tests.internal_ohmic_measurement_resistance_mohm,
                          max_resistance_monoblock_12_mohm: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-teal-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Độ biến động Nội trở (%)</label>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.internal_ohmic_measurement_resistance_mohm.variance_percent}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        internal_ohmic_measurement_resistance_mohm: {
                          ...data.electrical_tests.internal_ohmic_measurement_resistance_mohm,
                          variance_percent: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className={`w-full bg-slate-950 border rounded-lg px-3 py-2 font-mono font-bold focus:outline-none ${
                    data.electrical_tests.internal_ohmic_measurement_resistance_mohm.variance_percent > 25
                      ? 'border-rose-500 text-rose-300'
                      : 'border-slate-700 text-white'
                  }`}
                />
                <span className="text-[10px] text-slate-400 mt-1 block">NETA 7.18.1.3.D.7: ≤ 25%</span>
              </div>
            </div>

            {/* Intercell resistances array */}
            <div className="pt-2">
              <div className="flex items-center justify-between mb-1.5 text-xs">
                <span className="text-slate-400 font-medium">Điện trở Cầu nối Intercell (µΩ)</span>
                <span className={`text-[11px] font-bold ${
                  liveStats.intercellPass ? 'text-emerald-400' : 'text-amber-400'
                }`}>
                  Lệch: {liveStats.intercellDev}% (Ngưỡng NETA ≤ 50%)
                </span>
              </div>
              <div className="grid grid-cols-4 gap-2">
                {data.electrical_tests.intercell_connection_resistance_micro_ohms.map((res, idx) => (
                  <div key={idx} className="relative">
                    <span className="absolute left-2 top-2 text-[10px] text-slate-500 font-mono">#{idx + 1}</span>
                    <input
                      type="number"
                      step="0.1"
                      value={res}
                      onChange={(e) => {
                        const newArr = [...data.electrical_tests.intercell_connection_resistance_micro_ohms];
                        newArr[idx] = parseFloat(e.target.value) || 0;
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            intercell_connection_resistance_micro_ohms: newArr
                          }
                        });
                      }}
                      className="w-full bg-slate-950 border border-slate-700 rounded-lg pl-8 pr-2 py-1.5 text-white font-mono text-xs focus:outline-none focus:border-teal-500"
                    />
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* Section: Điện áp Bộ Sạc, Monoblock & Điện áp Đối Đất */}
          <div className="bg-slate-900/80 border border-slate-800 rounded-2xl p-5 shadow-lg space-y-4">
            <h2 className="text-base font-bold text-white flex items-center gap-2 border-b border-slate-800 pb-3">
              <Power className="w-5 h-5 text-teal-400" />
              5. Điện Áp Sạc, Monoblock Float & Điện Áp Đất
            </h2>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-3 text-xs">
              <div>
                <label className="block text-slate-400 mb-1 font-medium">Điện áp Float Bộ sạc (V)</label>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.charger_float_voltage_v}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        charger_float_voltage_v: parseFloat(e.target.value) || 0
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-teal-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Điện áp Equalize (V)</label>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.charger_equalize_voltage_v}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        charger_equalize_voltage_v: parseFloat(e.target.value) || 0
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-teal-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Cảnh báo Bộ sạc (Alarms)</label>
                <select
                  value={data.electrical_tests.charger_alarms_check}
                  onChange={(e: any) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        charger_alarms_check: e.target.value
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-teal-500"
                >
                  <option value="Pass">Pass - Bình thường</option>
                  <option value="Investigate">Investigate - Có cảnh báo</option>
                  <option value="Fail">Fail - Lỗi bộ sạc</option>
                </select>
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Áp Monoblock Min (V)</label>
                <input
                  type="number"
                  step="0.01"
                  value={data.electrical_tests.monoblock_voltages_float_mode_v.min_voltage_v}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        monoblock_voltages_float_mode_v: {
                          ...data.electrical_tests.monoblock_voltages_float_mode_v,
                          min_voltage_v: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-teal-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Áp Monoblock Max (V)</label>
                <input
                  type="number"
                  step="0.01"
                  value={data.electrical_tests.monoblock_voltages_float_mode_v.max_voltage_v}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        monoblock_voltages_float_mode_v: {
                          ...data.electrical_tests.monoblock_voltages_float_mode_v,
                          max_voltage_v: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white font-mono focus:outline-none focus:border-teal-500"
                />
              </div>

              <div>
                <label className="block text-slate-400 mb-1 font-medium">Thử nghiệm Xả tải</label>
                <input
                  type="text"
                  value={data.electrical_tests.load_test_ieee_1188}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        load_test_ieee_1188: e.target.value
                      }
                    })
                  }
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-teal-500"
                />
              </div>
            </div>

            {/* System Voltage to Ground */}
            <div className="pt-2 border-t border-slate-800">
              <label className="block text-slate-400 mb-1 text-xs font-medium">
                Điện Áp So Với Đất (System Voltage Positive/Negative to Ground per NETA 7.18.1.3.D.8)
              </label>
              <div className="grid grid-cols-2 gap-3 text-xs">
                <div>
                  <span className="text-[11px] text-slate-500">Cực Dương (+) đối đất (VDC)</span>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.system_voltage_to_ground_v.positive_to_ground_v}
                    onChange={(e) =>
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          system_voltage_to_ground_v: {
                            ...data.electrical_tests.system_voltage_to_ground_v,
                            positive_to_ground_v: parseFloat(e.target.value) || 0
                          }
                        }
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-cyan-300 font-mono font-bold focus:outline-none focus:border-teal-500"
                  />
                </div>
                <div>
                  <span className="text-[11px] text-slate-500">Cực Âm (-) đối đất (VDC)</span>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.system_voltage_to_ground_v.negative_to_ground_v}
                    onChange={(e) =>
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          system_voltage_to_ground_v: {
                            ...data.electrical_tests.system_voltage_to_ground_v,
                            negative_to_ground_v: parseFloat(e.target.value) || 0
                          }
                        }
                      })
                    }
                    className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-cyan-300 font-mono font-bold focus:outline-none focus:border-teal-500"
                  />
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Action Buttons Bar */}
      <div className="flex flex-wrap items-center justify-between gap-4 p-5 bg-slate-900/90 border border-slate-800 rounded-2xl shadow-xl">
        <div className="flex items-center gap-2 text-xs text-slate-400">
          <Clock className="w-4 h-4 text-teal-400" />
          <span>Đối soát chuẩn xác theo ANSI/NETA ATS-2025 Mục 7.18.1.3 & IEEE 1188</span>
        </div>

        <div className="flex flex-wrap items-center gap-3">
          <button
            onClick={handleRunAiAnalysis}
            disabled={isAnalyzing}
            className="px-5 py-2.5 rounded-xl bg-gradient-to-r from-teal-600 to-emerald-600 hover:from-teal-500 hover:to-emerald-500 text-white font-semibold text-sm flex items-center gap-2 shadow-lg shadow-teal-500/20 disabled:opacity-50 transition-all cursor-pointer"
          >
            {isAnalyzing ? (
              <>
                <RotateCcw className="w-4 h-4 animate-spin" />
                <span>AI Đang Phân Tích & Đối Soát...</span>
              </>
            ) : (
              <>
                <Sparkles className="w-4 h-4 text-teal-200" />
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
                <Printer className="w-4 h-4 text-teal-400" />
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
                  Kết Luận Kiểm Định NETA ATS-2025 Mục 7.18.1.3 & IEEE 1188
                </div>
                <h3 className={`text-2xl font-bold mt-0.5 ${
                  analysisResult.overall_status === 'FAIL'
                    ? 'text-rose-400'
                    : analysisResult.overall_status === 'INVESTIGATE'
                    ? 'text-amber-400'
                    : 'text-emerald-400'
                }`}>
                  {analysisResult.overall_status === 'FAIL'
                    ? 'FAIL / CRITICAL THERMAL RUNAWAY & INTERNAL OHMIC DEFECT'
                    : analysisResult.overall_status === 'INVESTIGATE'
                    ? 'INVESTIGATE / CẦN THEO DÕI ĐẶC BIỆT'
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
                <FileText className="w-5 h-5 text-teal-400" />
                Báo Cáo Thẩm Định Chuyên Gia NETA ATS-2025 & IEEE 1188
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
                <Upload className="w-5 h-5 text-teal-400" />
                Dữ Liệu JSON Kiểm Tra Hiện Trường VRLA
              </h3>
              <button
                onClick={() => setShowJsonModal(false)}
                className="text-slate-400 hover:text-white"
              >
                ✕
              </button>
            </div>

            <p className="text-xs text-slate-400">
              Bạn có thể dán trực tiếp dữ liệu JSON từ app đo kiểm hoặc file cấu hình vào đây để nạp nhanh vào form.
            </p>

            <textarea
              rows={14}
              value={jsonInput}
              onChange={(e) => setJsonInput(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl p-3 font-mono text-xs text-teal-300 focus:outline-none focus:border-teal-500"
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
                className="px-4 py-2 rounded-xl bg-teal-600 hover:bg-teal-500 text-white text-xs font-semibold"
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
          equipmentCategory="battery"
        />
      )}

      {/* Print Report Modal */}
      {showPrintModal && analysisResult && (
        <TevReportPrintModal
          isOpen={showPrintModal}
          onClose={() => setShowPrintModal(false)}
          reportTitle={`BIÊN BẢN KIỂM ĐỊNH NETA ATS-2025 - GIÀN PIN VRLA ${data.site_info.battery_bank_tag}`}
          equipmentTag={data.site_info.battery_bank_tag}
          equipmentType="Ắc quy Chì-Axit Kín Khí (VRLA)"
          testStandard="ANSI/NETA ATS-2025 Mục 7.18.1.3 & IEEE 1188"
          overallStatus={analysisResult.overall_status}
          markdownReport={analysisResult.markdown_report}
        />
      )}
    </div>
  );
};
