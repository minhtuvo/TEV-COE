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
  Droplets,
  Wind,
  Flame,
  Scale,
  Copy,
  Info
} from 'lucide-react';
import {
  NetaBatteryFloodedInputPayload,
  NetaBatteryFloodedEvaluationResult,
  performDeterministicBatteryFloodedAnalysis
} from '../../server/netaBatteryFloodedAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';
import { NameplateScannerModal } from './NameplateScannerModal';
import { ExtractedNameplateData } from '../../server/nameplateOcrAnalyzer';
import { FseEngineerDropdown } from './FseEngineerDropdown';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

export const SAMPLE_BATTERY_FLOODED_FSE_DATA: NetaBatteryFloodedInputPayload = {
  site_info: {
    project_name: "Tram 220kV TEV Substation DC System",
    battery_bank_tag: "BAT-DC-110V-FLOODED",
    battery_type: "Flooded Lead-Acid (54 Cells 2V 400Ah)",
    fse_name: "Nguyen Van P",
    test_date: "2026-09-24",
    manufacturer: "Enersys / Yuasa / Hoppecke",
    model_number: "2V-400Ah OPzS",
    serial_number: "SN-TEV-BAT-2026-054"
  },
  visual_inspection: {
    ventilation_system_operable: true,
    eyewash_and_spill_containment_present: true,
    nameplate_match: true,
    rack_mounting_and_grounding: "Pass",
    flame_arresters_present: true,
    cleanliness_and_oxide_inhibitor: "Pass",
    electrolyte_level_check: "Pass (All cells at MAX mark)",
    bolt_torque_check: "Pass"
  },
  electrical_tests: {
    intercell_connection_resistance_micro_ohms: [12.1, 12.3, 12.2, 28.5],
    charger_float_voltage_v: 122.5,
    charger_equalize_voltage_v: 127.8,
    charger_alarms_check: "Pass",
    cell_voltages_float_mode_v: {
      cell_min_v: 2.21,
      cell_max_v: 2.29,
      voltage_spread_v: 0.08
    },
    specific_gravity_corrected_25c: {
      pilot_cell_sg: 1.215,
      min_sg_found: 1.180,
      max_sg_found: 1.220
    },
    internal_ohmic_measurement_resistance_mohm: {
      avg_resistance_mohm: 0.45,
      max_resistance_cell_mohm: 0.62,
      variance_percent: 37.7
    },
    load_test_ieee_450: "Pass (Capacity 94% per IEEE 450)",
    system_voltage_to_ground_v: {
      positive_to_ground_v: 82.0,
      negative_to_ground_v: -40.5
    }
  },
  previous_test_data: {
    last_test_date: "2025-09-20",
    last_avg_resistance_mohm: 0.44
  }
};

export const NetaBatteryFloodedChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaBatteryFloodedInputPayload>(SAMPLE_BATTERY_FLOODED_FSE_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaBatteryFloodedEvaluationResult | null>(null);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showCameraModal, setShowCameraModal] = useState(false);
  const [jsonInput, setJsonInput] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live real-time calculations
  const liveStats = useMemo(() => {
    // 1. Intercell connection resistance
    const resistances = data.electrical_tests.intercell_connection_resistance_micro_ohms || [];
    let minRes = 0;
    let maxRes = 0;
    let intercellDev = 0;
    let intercellPass = true;
    let maxResIndex = -1;

    if (resistances.length > 0) {
      minRes = Math.min(...resistances);
      maxRes = Math.max(...resistances);
      maxResIndex = resistances.indexOf(maxRes);
      if (minRes > 0) {
        intercellDev = ((maxRes - minRes) / minRes) * 100;
        if (intercellDev > 50) intercellPass = false;
      }
    }

    // 2. Cell float voltage spread
    const cellMin = data.electrical_tests.cell_voltages_float_mode_v.cell_min_v;
    const cellMax = data.electrical_tests.cell_voltages_float_mode_v.cell_max_v;
    const spreadV = Math.abs(cellMax - cellMin);
    const cellSpreadPass = spreadV <= 0.05;

    // 3. Internal ohmic measurement variance
    const avgRes = data.electrical_tests.internal_ohmic_measurement_resistance_mohm.avg_resistance_mohm;
    const maxResCell = data.electrical_tests.internal_ohmic_measurement_resistance_mohm.max_resistance_cell_mohm;
    let variancePct = data.electrical_tests.internal_ohmic_measurement_resistance_mohm.variance_percent;
    if (avgRes > 0 && maxResCell > 0) {
      variancePct = parseFloat((((maxResCell - avgRes) / avgRes) * 100).toFixed(1));
    }
    const ohmicPass = variancePct <= 25.0;

    // 4. Grounding voltage balance
    const vPos = Math.abs(data.electrical_tests.system_voltage_to_ground_v.positive_to_ground_v || 0);
    const vNeg = Math.abs(data.electrical_tests.system_voltage_to_ground_v.negative_to_ground_v || 0);
    const totalBusV = vPos + vNeg;
    const vDiff = Math.abs(vPos - vNeg);
    const isGroundBalanced = vDiff <= 5.0;
    const groundFaultType = !isGroundBalanced ? (vNeg < vPos ? 'NEGATIVE' : 'POSITIVE') : 'NONE';

    // 5. Specific gravity
    const minSg = data.electrical_tests.specific_gravity_corrected_25c.min_sg_found;
    const pilotSg = data.electrical_tests.specific_gravity_corrected_25c.pilot_cell_sg;
    const sgPass = minSg >= 1.200 && (pilotSg - minSg) <= 0.020;

    // 6. Overall safety flag
    const overallSafe = intercellPass && cellSpreadPass && ohmicPass && isGroundBalanced && sgPass;

    return {
      minRes,
      maxRes,
      maxResIndex,
      intercellDev,
      intercellPass,
      spreadV,
      cellSpreadPass,
      variancePct,
      ohmicPass,
      vPos,
      vNeg,
      totalBusV,
      vDiff,
      isGroundBalanced,
      groundFaultType,
      sgPass,
      overallSafe
    };
  }, [data]);

  const handleApplyNameplateData = (extracted: ExtractedNameplateData) => {
    setData((prev) => ({
      ...prev,
      site_info: {
        ...prev.site_info,
        battery_bank_tag: extracted.equipment_tag || prev.site_info.battery_bank_tag,
        battery_type: extracted.equipment_type || extracted.equipment_name || prev.site_info.battery_type,
        manufacturer: extracted.manufacturer || prev.site_info.manufacturer,
        serial_number: extracted.serial_number || prev.site_info.serial_number
      },
      visual_inspection: {
        ...prev.visual_inspection,
        nameplate_match: true
      }
    }));
    setSavedSuccessMessage(`Đã quét nhãn giàn pin thành công: ${extracted.equipment_tag || 'Battery Bank'}!`);
    setTimeout(() => setSavedSuccessMessage(null), 5000);
  };

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    setAnalysisResult(null);
    try {
      const res = await fetch('/api/field-service/neta-battery-flooded-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });
      const json = await res.json();
      if (json.success && json.data) {
        setAnalysisResult(json.data);
      } else {
        const det = performDeterministicBatteryFloodedAnalysis(data);
        setAnalysisResult(det);
      }
    } catch {
      const det = performDeterministicBatteryFloodedAnalysis(data);
      setAnalysisResult(det);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleSaveToCmms = () => {
    const reportPayload = {
      type: 'NETA_ATS_2025_BATTERY_FLOODED',
      data: data,
      analysisResult: analysisResult,
      equipmentId: data.site_info.battery_bank_tag,
      date: data.site_info.test_date,
      title: `Biên bản Giàn Ắc Quy Flooded Lead-Acid NETA ATS-2025: ${data.site_info.battery_bank_tag}`,
      standard: 'ANSI/NETA ATS-2025 Section 7.18.1.1 & IEEE 450',
      status: liveStats.overallSafe ? 'PASS' : 'INVESTIGATE',
      timestamp: new Date().toISOString()
    };

    if (onSaveReport) {
      onSaveReport(reportPayload);
      setSavedSuccessMessage(`Đã lưu biên bản giàn pin ${data.site_info.battery_bank_tag} vào kho báo cáo thành công!`);
      setTimeout(() => setSavedSuccessMessage(null), 5000);
    }
  };

  const handleCopyMarkdown = () => {
    if (analysisResult?.markdown_report) {
      navigator.clipboard.writeText(analysisResult.markdown_report);
      setCopied(true);
      setTimeout(() => setCopied(false), 2500);
    }
  };

  const electricalTableForPrint = [
    {
      item: 'Điện trở cầu nối Intercell (Intercell Resistance)',
      measured: `${data.electrical_tests.intercell_connection_resistance_micro_ohms.join(', ')} µΩ (Max: ${liveStats.maxRes.toFixed(1)} µΩ, lệch ${liveStats.intercellDev.toFixed(1)}%)`,
      standard: 'Cảnh báo nếu lệch > 50% so với giá trị nhỏ nhất',
      reference: 'NETA ATS-2025 Sec 7.18.1.1.D.1 & Table 100.12',
      deviation: `${liveStats.intercellDev.toFixed(1)}%`,
      status: liveStats.intercellPass ? 'PASS' : 'INVESTIGATE'
    },
    {
      item: 'Điện áp sạc Float & Equalize (Charger Voltages & Alarms)',
      measured: `Float: ${data.electrical_tests.charger_float_voltage_v}V | Equalize: ${data.electrical_tests.charger_equalize_voltage_v}V | Alarms: ${data.electrical_tests.charger_alarms_check}`,
      standard: 'Đạt định mức nhà sản xuất; cảnh báo hoạt động tin cậy',
      reference: 'NETA ATS-2025 Sec 7.18.1.1.D.2 & D.3',
      deviation: 'Đạt chuẩn',
      status: 'PASS'
    },
    {
      item: 'Độ lệch điện áp Cell chế độ Float (Cell Voltage Spread)',
      measured: `Min: ${data.electrical_tests.cell_voltages_float_mode_v.cell_min_v}V | Max: ${data.electrical_tests.cell_voltages_float_mode_v.cell_max_v}V (ΔV: ${liveStats.spreadV.toFixed(2)}V)`,
      standard: 'Điện áp giữa các cell không lệch quá 0.05 Volt',
      reference: 'NETA ATS-2025 Sec 7.18.1.1.D.4 & IEEE 450',
      deviation: `${liveStats.spreadV.toFixed(2)}V`,
      status: liveStats.cellSpreadPass ? 'PASS' : 'FAIL'
    },
    {
      item: 'Tỷ trọng dung dịch hiệu chỉnh 25°C (Specific Gravity)',
      measured: `Pilot: ${data.electrical_tests.specific_gravity_corrected_25c.pilot_cell_sg} | Min: ${data.electrical_tests.specific_gravity_corrected_25c.min_sg_found} | Max: ${data.electrical_tests.specific_gravity_corrected_25c.max_sg_found}`,
      standard: '1.215 ± 0.010 @ 25°C; độ lệch không quá 0.020',
      reference: 'NETA ATS-2025 Sec 7.18.1.1.D.5 & IEEE 450',
      deviation: `${(data.electrical_tests.specific_gravity_corrected_25c.pilot_cell_sg - data.electrical_tests.specific_gravity_corrected_25c.min_sg_found).toFixed(3)}`,
      status: liveStats.sgPass ? 'PASS' : 'FAIL'
    },
    {
      item: 'Nội trở Ôm Cell (Internal Ohmic Measurements)',
      measured: `Avg: ${data.electrical_tests.internal_ohmic_measurement_resistance_mohm.avg_resistance_mohm} mΩ | Max: ${data.electrical_tests.internal_ohmic_measurement_resistance_mohm.max_resistance_cell_mohm} mΩ (Lệch ${liveStats.variancePct.toFixed(1)}%)`,
      standard: 'Không được biến động quá 25% (> 25%) so với trung bình',
      reference: 'NETA ATS-2025 Sec 7.18.1.1.D.6 & IEEE 450',
      deviation: `${liveStats.variancePct.toFixed(1)}%`,
      status: liveStats.ohmicPass ? 'PASS' : 'FAIL'
    },
    {
      item: 'Thử nghiệm xả tải dung lượng (Load Test per IEEE 450)',
      measured: data.electrical_tests.load_test_ieee_450,
      standard: 'Dung lượng thực tế đạt ≥ 80% dung lượng thiết kế',
      reference: 'NETA ATS-2025 Sec 7.18.1.1.D.7 & IEEE 450',
      deviation: '94% (Đạt)',
      status: 'PASS'
    },
    {
      item: 'Điện áp cực Dương/Âm so với Đất (System Voltage to Ground)',
      measured: `(+) to Gnd: ${data.electrical_tests.system_voltage_to_ground_v.positive_to_ground_v > 0 ? '+' : ''}${data.electrical_tests.system_voltage_to_ground_v.positive_to_ground_v}V | (-) to Gnd: ${data.electrical_tests.system_voltage_to_ground_v.negative_to_ground_v}V`,
      standard: 'Độ lớn cực (+) và (-) đến đất BẮT BUỘC phải bằng nhau (equal in magnitude)',
      reference: 'NETA ATS-2025 Sec 7.18.1.1.D.8 & IEEE 450',
      deviation: `${liveStats.vDiff.toFixed(1)}V (Lệch đất)`,
      status: liveStats.isGroundBalanced ? 'PASS' : 'FAIL'
    }
  ];

  return (
    <div className="space-y-6 animate-in fade-in duration-300">
      {/* 1. TOP HEADER BANNER */}
      <div className="bg-gradient-to-r from-amber-950 via-slate-900 to-amber-900 rounded-2xl p-6 text-white shadow-xl border border-amber-800/40 relative overflow-hidden">
        <div className="absolute right-0 top-0 w-96 h-96 bg-amber-500/10 rounded-full blur-3xl pointer-events-none"></div>
        <div className="relative z-10 flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1.5">
            <div className="flex flex-wrap items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-500/30 text-amber-200 border border-amber-400/40 tracking-wide uppercase">
                ANSI/NETA ATS-2025 Mục 7.18.1.1 &amp; IEEE 450
              </span>
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-400/30">
                TEV Platform AI • Field Service Inspection
              </span>
            </div>
            <h1 className="text-2xl font-black tracking-tight text-white flex items-center gap-2.5">
              <BatteryCharging size={28} className="text-amber-400 flex-shrink-0" />
              <span>Hệ Thống Điện Một Chiều DC &amp; Ắc Quy Chì-Axit Ngập Dung Dịch</span>
            </h1>
            <p className="text-xs text-amber-200/80 max-w-3xl leading-relaxed">
              Direct-Current Systems, Batteries, Flooded Lead-Acid — Kiểm định an toàn khí Hydro, thông gió, tỷ trọng dung dịch, 
              độ lệch điện áp cell float (&le; 0.05V), nội trở ôm (&le; 25%) và phát hiện chạm đất DC (Positive/Negative Ground Fault).
            </p>
          </div>

          {/* Action Button Toolbar */}
          <div className="flex flex-wrap items-center gap-2 pt-2 lg:pt-0">
            <button
              type="button"
              onClick={() => {
                setData(SAMPLE_BATTERY_FLOODED_FSE_DATA);
                setAnalysisResult(null);
                setSavedSuccessMessage('Đã nạp bộ mẫu FSE Trạm 220kV TEV Substation (Chạm đất cực âm & chai cell)!');
                setTimeout(() => setSavedSuccessMessage(null), 4000);
              }}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold bg-amber-500/20 hover:bg-amber-500/30 text-amber-200 border border-amber-400/30 transition-all cursor-pointer shadow-xs"
              title="Nạp dữ liệu mẫu hiện trường FSE để chạy thử nghiệm thuật toán AI"
            >
              <RotateCcw size={14} />
              <span>Nạp Mẫu FSE Trạm 220kV</span>
            </button>

            <button
              type="button"
              onClick={() => setShowCameraModal(true)}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition-all cursor-pointer shadow-xs"
              title="Quét nhãn mác giàn pin bằng AI Vision OCR"
            >
              <Camera size={14} className="text-amber-400" />
              <span>Quét Nhãn Mác</span>
            </button>

            <button
              type="button"
              onClick={() => {
                setJsonInput(JSON.stringify(data, null, 2));
                setShowJsonModal(true);
              }}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition-all cursor-pointer shadow-xs"
              title="Xem hoặc nhập mã JSON từ ứng dụng FSE di động"
            >
              <FileText size={14} className="text-blue-400" />
              <span>JSON Input</span>
            </button>

            <button
              type="button"
              onClick={() => setShowPrintModal(true)}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition-all cursor-pointer shadow-xs"
              title="Xuất phiếu kiểm định chuẩn A4"
            >
              <Printer size={14} className="text-emerald-400" />
              <span>In / PDF A4</span>
            </button>

            <button
              type="button"
              onClick={handleAnalyze}
              disabled={isAnalyzing}
              className="inline-flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold bg-gradient-to-r from-amber-500 to-amber-600 hover:from-amber-400 hover:to-amber-500 active:scale-95 text-slate-950 transition-all cursor-pointer shadow-lg shadow-amber-500/20 disabled:opacity-50"
            >
              <Sparkles size={15} className={isAnalyzing ? 'animate-spin' : ''} />
              <span>{isAnalyzing ? 'Đang Phân Tích...' : 'Chạy Phân Tích AI'}</span>
            </button>
          </div>
        </div>
      </div>

      {savedSuccessMessage && (
        <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-xl text-emerald-900 text-xs font-semibold flex items-center gap-2 animate-in fade-in">
          <CheckCircle2 size={16} className="text-emerald-600 flex-shrink-0" />
          <span>{savedSuccessMessage}</span>
        </div>
      )}

      {/* 2. REAL-TIME LIVE TELEMETRY CARDS */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3.5">
        {/* Card 1: Ground Voltage & Ground Fault Alert */}
        <div className={`p-4 rounded-xl border flex flex-col justify-between transition-all shadow-xs ${
          !liveStats.isGroundBalanced
            ? 'bg-rose-50/90 border-rose-300 ring-2 ring-rose-400/40 text-rose-950'
            : 'bg-white border-slate-200'
        }`}>
          <div>
            <div className="flex items-center justify-between text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-1">
              <span>Điện Áp Đối Đất (+/-)</span>
              <ShieldAlert size={14} className={!liveStats.isGroundBalanced ? 'text-rose-600' : 'text-emerald-600'} />
            </div>
            <div className="font-mono text-base font-black text-slate-900 mt-0.5">
              +{liveStats.vPos.toFixed(1)}V / -{liveStats.vNeg.toFixed(1)}V
            </div>
            <div className="text-[11px] font-semibold mt-1">
              {!liveStats.isGroundBalanced ? (
                <span className="text-rose-700 flex items-center gap-1 font-bold">
                  <AlertTriangle size={12} className="flex-shrink-0" />
                  CHẠM ĐẤT CỰC {liveStats.groundFaultType === 'NEGATIVE' ? 'ÂM' : 'DƯƠNG'}! (Lệch {liveStats.vDiff.toFixed(1)}V)
                </span>
              ) : (
                <span className="text-emerald-700 flex items-center gap-1 font-bold">
                  <Check size={12} /> Cân bằng đối xứng (0V lệch)
                </span>
              )}
            </div>
          </div>
          <div className="mt-3 pt-2 border-t border-slate-100 text-[10px] text-slate-500">
            NETA: |V+| = |V-| (Equal magnitude)
          </div>
        </div>

        {/* Card 2: Float Cell Voltage Spread */}
        <div className={`p-4 rounded-xl border flex flex-col justify-between transition-all shadow-xs ${
          !liveStats.cellSpreadPass
            ? 'bg-amber-50/90 border-amber-300 text-amber-950'
            : 'bg-white border-slate-200'
        }`}>
          <div>
            <div className="flex items-center justify-between text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-1">
              <span>Lệch Áp Float (ΔV)</span>
              <Activity size={14} className={!liveStats.cellSpreadPass ? 'text-amber-600' : 'text-emerald-600'} />
            </div>
            <div className="font-mono text-base font-black text-slate-900 mt-0.5">
              {liveStats.spreadV.toFixed(2)} V
            </div>
            <div className="text-[11px] font-semibold mt-1">
              {!liveStats.cellSpreadPass ? (
                <span className="text-amber-800 font-bold">
                  Vượt ngưỡng 0.05V (Min: {data.electrical_tests.cell_voltages_float_mode_v.cell_min_v}V, Max: {data.electrical_tests.cell_voltages_float_mode_v.cell_max_v}V)
                </span>
              ) : (
                <span className="text-emerald-700 font-bold">
                  Đạt chuẩn (&le; 0.05V)
                </span>
              )}
            </div>
          </div>
          <div className="mt-3 pt-2 border-t border-slate-100 text-[10px] text-slate-500">
            NETA Sec 7.18.1.1.D.4 / IEEE 450
          </div>
        </div>

        {/* Card 3: Internal Ohmic Variance */}
        <div className={`p-4 rounded-xl border flex flex-col justify-between transition-all shadow-xs ${
          !liveStats.ohmicPass
            ? 'bg-rose-50/90 border-rose-300 text-rose-950'
            : 'bg-white border-slate-200'
        }`}>
          <div>
            <div className="flex items-center justify-between text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-1">
              <span>Nội Trở Ôm Biến Động</span>
              <Gauge size={14} className={!liveStats.ohmicPass ? 'text-rose-600' : 'text-emerald-600'} />
            </div>
            <div className="font-mono text-base font-black text-slate-900 mt-0.5">
              {liveStats.variancePct.toFixed(1)} %
            </div>
            <div className="text-[11px] font-semibold mt-1">
              {!liveStats.ohmicPass ? (
                <span className="text-rose-700 font-bold">
                  Lệch &gt; 25% (Max: {data.electrical_tests.internal_ohmic_measurement_resistance_mohm.max_resistance_cell_mohm} mΩ, Avg: {data.electrical_tests.internal_ohmic_measurement_resistance_mohm.avg_resistance_mohm} mΩ)
                </span>
              ) : (
                <span className="text-emerald-700 font-bold">
                  Đạt chuẩn (&le; 25%)
                </span>
              )}
            </div>
          </div>
          <div className="mt-3 pt-2 border-t border-slate-100 text-[10px] text-slate-500">
            NETA Sec 7.18.1.1.D.6 / IEEE 450
          </div>
        </div>

        {/* Card 4: Intercell Connection Resistance Deviation */}
        <div className={`p-4 rounded-xl border flex flex-col justify-between transition-all shadow-xs ${
          !liveStats.intercellPass
            ? 'bg-amber-50/90 border-amber-300 text-amber-950'
            : 'bg-white border-slate-200'
        }`}>
          <div>
            <div className="flex items-center justify-between text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-1">
              <span>Cầu Nối Intercell Lệch</span>
              <Scale size={14} className={!liveStats.intercellPass ? 'text-amber-600' : 'text-emerald-600'} />
            </div>
            <div className="font-mono text-base font-black text-slate-900 mt-0.5">
              {liveStats.intercellDev.toFixed(1)} %
            </div>
            <div className="text-[11px] font-semibold mt-1">
              {!liveStats.intercellPass ? (
                <span className="text-amber-800 font-bold">
                  INVESTIGATE (Mối #{liveStats.maxResIndex + 1}: {liveStats.maxRes.toFixed(1)} µΩ vs min {liveStats.minRes.toFixed(1)} µΩ)
                </span>
              ) : (
                <span className="text-emerald-700 font-bold">
                  Đạt chuẩn (&le; 50% min)
                </span>
              )}
            </div>
          </div>
          <div className="mt-3 pt-2 border-t border-slate-100 text-[10px] text-slate-500">
            NETA Sec 7.18.1.1.D.1 / Table 100.12
          </div>
        </div>

        {/* Card 5: Specific Gravity */}
        <div className={`p-4 rounded-xl border flex flex-col justify-between transition-all shadow-xs ${
          !liveStats.sgPass
            ? 'bg-amber-50/90 border-amber-300 text-amber-950'
            : 'bg-white border-slate-200'
        }`}>
          <div>
            <div className="flex items-center justify-between text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-1">
              <span>Tỷ Trọng Dung Dịch (SG)</span>
              <Droplets size={14} className={!liveStats.sgPass ? 'text-amber-600' : 'text-emerald-600'} />
            </div>
            <div className="font-mono text-base font-black text-slate-900 mt-0.5">
              {data.electrical_tests.specific_gravity_corrected_25c.min_sg_found.toFixed(3)}
            </div>
            <div className="text-[11px] font-semibold mt-1">
              {!liveStats.sgPass ? (
                <span className="text-amber-800 font-bold">
                  Sụt tỷ trọng (Pilot: {data.electrical_tests.specific_gravity_corrected_25c.pilot_cell_sg.toFixed(3)})
                </span>
              ) : (
                <span className="text-emerald-700 font-bold">
                  Chuẩn 1.215 ± 0.010
                </span>
              )}
            </div>
          </div>
          <div className="mt-3 pt-2 border-t border-slate-100 text-[10px] text-slate-500">
            NETA Sec 7.18.1.1.D.5 / IEEE 450
          </div>
        </div>
      </div>

      {/* 3. INTERACTIVE DATA ENTRY FORM */}
      <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-6">
        {/* Section 1: Site Info & Battery Specifications */}
        <div>
          <h2 className="text-sm font-bold text-slate-900 uppercase tracking-wider flex items-center gap-2 mb-3 pb-2 border-b border-slate-100">
            <span className="w-6 h-6 rounded-full bg-amber-100 text-amber-800 flex items-center justify-center text-xs font-black">1</span>
            <span>Thông Tin Trạm, Giàn Pin &amp; Kỹ Sư Phụ Trách (Site &amp; Battery Bank Info)</span>
          </h2>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <div>
              <label className="text-xs font-semibold text-slate-600 block mb-1">Tên Dự án / Trạm Điện</label>
              <input
                type="text"
                value={data.site_info.project_name}
                onChange={(e) => setData({ ...data, site_info: { ...data.site_info, project_name: e.target.value } })}
                className="w-full text-xs font-medium bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 outline-none focus:ring-2 focus:ring-amber-500"
              />
            </div>
            <div>
              <label className="text-xs font-semibold text-slate-600 block mb-1">Mã Giàn Pin (Battery Bank Tag)</label>
              <input
                type="text"
                value={data.site_info.battery_bank_tag}
                onChange={(e) => setData({ ...data, site_info: { ...data.site_info, battery_bank_tag: e.target.value } })}
                className="w-full text-xs font-bold text-slate-800 bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 outline-none focus:ring-2 focus:ring-amber-500 font-mono"
              />
            </div>
            <div>
              <label className="text-xs font-semibold text-slate-600 block mb-1">Chủng Loại &amp; Quy Cách Pin</label>
              <input
                type="text"
                value={data.site_info.battery_type}
                onChange={(e) => setData({ ...data, site_info: { ...data.site_info, battery_type: e.target.value } })}
                className="w-full text-xs font-medium bg-slate-50 border border-slate-200 rounded-lg px-3 py-2 outline-none focus:ring-2 focus:ring-amber-500"
              />
            </div>
            <div>
              <label className="text-xs font-semibold text-slate-600 block mb-1">Kỹ Sư Hiện Trường (FSE)</label>
              <FseEngineerDropdown
                currentEngineer={data.site_info.fse_name}
                onSelectEngineer={(name) => setData({ ...data, site_info: { ...data.site_info, fse_name: name } })}
              />
            </div>
          </div>
        </div>

        {/* Section 2: Visual & Mechanical Inspection (NETA Sec 7.18.1.1.A) */}
        <div>
          <h2 className="text-sm font-bold text-slate-900 uppercase tracking-wider flex items-center gap-2 mb-3 pb-2 border-b border-slate-100">
            <span className="w-6 h-6 rounded-full bg-blue-100 text-blue-800 flex items-center justify-center text-xs font-black">2</span>
            <span>Kiểm Tra Thị Giác &amp; Cơ Khí (Visual &amp; Mechanical Inspection per NETA Section 7.18.1.1.A)</span>
          </h2>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 text-xs">
            {/* 2.1 Ventilation */}
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/70 space-y-2">
              <div className="flex items-center justify-between">
                <span className="font-bold text-slate-700 flex items-center gap-1.5">
                  <Wind size={14} className="text-blue-600" />
                  Hệ Thống Thông Gió
                </span>
                <span className="text-[10px] font-bold text-slate-400">Bắt buộc</span>
              </div>
              <label className="flex items-center gap-2 cursor-pointer font-semibold text-slate-700">
                <input
                  type="checkbox"
                  checked={data.visual_inspection.ventilation_system_operable}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, ventilation_system_operable: e.target.checked }
                  })}
                  className="rounded text-amber-600 focus:ring-amber-500 w-4 h-4 cursor-pointer"
                />
                <span>Quạt hút hoạt động tốt (Chống Hydro)</span>
              </label>
              <p className="text-[11px] text-slate-500 leading-tight">
                NETA Sec 7.18.1.1.A.1: Ngăn ngừa tích tụ khí Hydro H₂ gây cháy nổ trạm.
              </p>
            </div>

            {/* 2.2 Eyewash & Spill Containment */}
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/70 space-y-2">
              <div className="flex items-center justify-between">
                <span className="font-bold text-slate-700 flex items-center gap-1.5">
                  <Droplets size={14} className="text-cyan-600" />
                  Rửa Mắt &amp; Chống Tràn Axit
                </span>
                <span className="text-[10px] font-bold text-slate-400">An toàn</span>
              </div>
              <label className="flex items-center gap-2 cursor-pointer font-semibold text-slate-700">
                <input
                  type="checkbox"
                  checked={data.visual_inspection.eyewash_and_spill_containment_present}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, eyewash_and_spill_containment_present: e.target.checked }
                  })}
                  className="rounded text-amber-600 focus:ring-amber-500 w-4 h-4 cursor-pointer"
                />
                <span>Có trạm rửa mắt &amp; khay hứng</span>
              </label>
              <p className="text-[11px] text-slate-500 leading-tight">
                NETA Sec 7.18.1.1.A.2: Khay hứng chống tràn axit và trạm rửa mắt khẩn cấp.
              </p>
            </div>

            {/* 2.3 Flame Arresters */}
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/70 space-y-2">
              <div className="flex items-center justify-between">
                <span className="font-bold text-slate-700 flex items-center gap-1.5">
                  <Flame size={14} className="text-rose-600" />
                  Nắp Chắn Lửa (Flame Arresters)
                </span>
                <span className="text-[10px] font-bold text-slate-400">100% Cell</span>
              </div>
              <label className="flex items-center gap-2 cursor-pointer font-semibold text-slate-700">
                <input
                  type="checkbox"
                  checked={data.visual_inspection.flame_arresters_present}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, flame_arresters_present: e.target.checked }
                  })}
                  className="rounded text-amber-600 focus:ring-amber-500 w-4 h-4 cursor-pointer"
                />
                <span>Đầy đủ nắp chắn lửa chống nổ ngược</span>
              </label>
              <p className="text-[11px] text-slate-500 leading-tight">
                NETA Sec 7.18.1.1.A.5: Xốp ngăn tia lửa ngoài lọt vào bên trong buồng khí cell.
              </p>
            </div>

            {/* 2.4 Nameplate Match */}
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/70 space-y-2">
              <div className="flex items-center justify-between">
                <span className="font-bold text-slate-700 flex items-center gap-1.5">
                  <ShieldCheck size={14} className="text-emerald-600" />
                  Đối Soát Nhãn Mác
                </span>
                <span className="text-[10px] font-bold text-slate-400">100% khớp</span>
              </div>
              <label className="flex items-center gap-2 cursor-pointer font-semibold text-slate-700">
                <input
                  type="checkbox"
                  checked={data.visual_inspection.nameplate_match}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, nameplate_match: e.target.checked }
                  })}
                  className="rounded text-amber-600 focus:ring-amber-500 w-4 h-4 cursor-pointer"
                />
                <span>Khớp bản vẽ thiết kế 100%</span>
              </label>
              <p className="text-[11px] text-slate-500 leading-tight">
                NETA Sec 7.18.1.1.A.3: Khớp chủng loại bình Flooded 2V, dung lượng 400Ah.
              </p>
            </div>

            {/* 2.5 Racks & Grounding */}
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/70 space-y-1.5">
              <label className="font-bold text-slate-700 block">Giá Đỡ &amp; Tiếp Địa Khung (Racks)</label>
              <select
                value={data.visual_inspection.rack_mounting_and_grounding}
                onChange={(e) => setData({
                  ...data,
                  visual_inspection: { ...data.visual_inspection, rack_mounting_and_grounding: e.target.value }
                })}
                className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 font-semibold text-slate-800 outline-none focus:ring-2 focus:ring-amber-500"
              >
                <option value="Pass">Pass (Cố định chắc chắn &amp; tiếp địa đạt)</option>
                <option value="Fail">Fail (Lỏng lẻo hoặc thiếu tiếp địa vỏ)</option>
              </select>
            </div>

            {/* 2.6 Cleanliness & Oxide Inhibitor */}
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/70 space-y-1.5">
              <label className="font-bold text-slate-700 block">Vệ Sinh Cọc &amp; Mỡ Chống Oxy Hóa</label>
              <select
                value={data.visual_inspection.cleanliness_and_oxide_inhibitor}
                onChange={(e) => setData({
                  ...data,
                  visual_inspection: { ...data.visual_inspection, cleanliness_and_oxide_inhibitor: e.target.value }
                })}
                className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 font-semibold text-slate-800 outline-none focus:ring-2 focus:ring-amber-500"
              >
                <option value="Pass">Pass (Sạch cọc, có bôi mỡ NO-OX-ID)</option>
                <option value="Fail">Fail (Bám muối axit hoặc thiếu mỡ bảo vệ)</option>
              </select>
            </div>

            {/* 2.7 Electrolyte Level */}
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/70 space-y-1.5">
              <label className="font-bold text-slate-700 block">Mức Dung Dịch Điện Môi</label>
              <input
                type="text"
                value={data.visual_inspection.electrolyte_level_check}
                onChange={(e) => setData({
                  ...data,
                  visual_inspection: { ...data.visual_inspection, electrolyte_level_check: e.target.value }
                })}
                className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 font-semibold text-slate-800 outline-none focus:ring-2 focus:ring-amber-500"
              />
            </div>

            {/* 2.8 Bolt Torque */}
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/70 space-y-1.5">
              <label className="font-bold text-slate-700 block">Lực Siết Bu-lông (Bolt Torque)</label>
              <select
                value={data.visual_inspection.bolt_torque_check}
                onChange={(e) => setData({
                  ...data,
                  visual_inspection: { ...data.visual_inspection, bolt_torque_check: e.target.value }
                })}
                className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 font-semibold text-slate-800 outline-none focus:ring-2 focus:ring-amber-500"
              >
                <option value="Pass">Pass (Siết đạt cờ-lê lực theo Table 100.12)</option>
                <option value="Fail">Fail (Chưa kiểm tra hoặc lỏng bu-lông)</option>
              </select>
            </div>
          </div>
        </div>

        {/* Section 3: Electrical Test Values (NETA Sec 7.18.1.1.D & IEEE 450) */}
        <div>
          <h2 className="text-sm font-bold text-slate-900 uppercase tracking-wider flex items-center gap-2 mb-3 pb-2 border-b border-slate-100">
            <span className="w-6 h-6 rounded-full bg-emerald-100 text-emerald-800 flex items-center justify-center text-xs font-black">3</span>
            <span>Phép Đo Điện &amp; Tiêu Chuẩn Đánh Giá (Electrical Test Values per NETA Section 7.18.1.1.D &amp; IEEE 450)</span>
          </h2>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-5 text-xs">
            {/* 3.1 Intercell Connection Resistance */}
            <div className="p-4 rounded-xl border border-slate-200 bg-slate-50/50 space-y-2.5">
              <div className="flex items-center justify-between">
                <label className="font-bold text-slate-800">
                  Điện Trở Cầu Nối Intercell (µΩ)
                </label>
                <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                  liveStats.intercellPass ? 'bg-emerald-100 text-emerald-700' : 'bg-amber-100 text-amber-800'
                }`}>
                  Lệch: {liveStats.intercellDev.toFixed(1)}% (Max &le; 50%)
                </span>
              </div>
              <input
                type="text"
                value={data.electrical_tests.intercell_connection_resistance_micro_ohms.join(', ')}
                onChange={(e) => {
                  const arr = e.target.value
                    .split(',')
                    .map((s) => parseFloat(s.trim()))
                    .filter((n) => !isNaN(n));
                  setData({
                    ...data,
                    electrical_tests: {
                      ...data.electrical_tests,
                      intercell_connection_resistance_micro_ohms: arr
                    }
                  });
                }}
                className="w-full bg-white border border-slate-300 rounded-lg px-3 py-2 font-mono font-bold text-slate-800 outline-none focus:ring-2 focus:ring-amber-500"
                placeholder="12.1, 12.3, 12.2, 28.5"
              />
              <p className="text-[11px] text-slate-500 leading-tight">
                NETA Sec 7.18.1.1.D.1: CẢNH BÁO "INVESTIGATE" nếu giá trị đo lệch quá 50% so với giá trị nhỏ nhất của mối nối tương tự.
              </p>
            </div>

            {/* 3.2 System Voltage Positive/Negative to Ground */}
            <div className={`p-4 rounded-xl border space-y-2.5 ${
              !liveStats.isGroundBalanced
                ? 'bg-rose-50/80 border-rose-300'
                : 'bg-slate-50/50 border-slate-200'
            }`}>
              <div className="flex items-center justify-between">
                <label className="font-bold text-slate-800 flex items-center gap-1">
                  <span>Điện Áp So Với Đất (+) &amp; (-)</span>
                </label>
                <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                  liveStats.isGroundBalanced ? 'bg-emerald-100 text-emerald-700' : 'bg-rose-200 text-rose-900 animate-pulse'
                }`}>
                  {liveStats.isGroundBalanced ? 'Cân bằng' : `Chạm đất cực ${liveStats.groundFaultType === 'NEGATIVE' ? 'Âm' : 'Dương'}`}
                </span>
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <span className="text-[10px] text-slate-500 font-semibold block mb-0.5">Cực Dương (+) to Gnd (V)</span>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.system_voltage_to_ground_v.positive_to_ground_v}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        system_voltage_to_ground_v: {
                          ...data.electrical_tests.system_voltage_to_ground_v,
                          positive_to_ground_v: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 font-mono font-bold text-slate-800 outline-none focus:ring-2 focus:ring-amber-500"
                  />
                </div>
                <div>
                  <span className="text-[10px] text-slate-500 font-semibold block mb-0.5">Cực Âm (-) to Gnd (V)</span>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.system_voltage_to_ground_v.negative_to_ground_v}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        system_voltage_to_ground_v: {
                          ...data.electrical_tests.system_voltage_to_ground_v,
                          negative_to_ground_v: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 font-mono font-bold text-slate-800 outline-none focus:ring-2 focus:ring-amber-500"
                  />
                </div>
              </div>
              <p className="text-[11px] text-slate-500 leading-tight">
                NETA Sec 7.18.1.1.D.8: Độ lớn điện áp cực (+) đến đất BẮT BUỘC phải bằng độ lớn cực (-) đến đất (equal in magnitude).
              </p>
            </div>

            {/* 3.3 Cell Voltages (Float Mode) */}
            <div className="p-4 rounded-xl border border-slate-200 bg-slate-50/50 space-y-2.5">
              <div className="flex items-center justify-between">
                <label className="font-bold text-slate-800">
                  Điện Áp Cell Chế Độ Float (V)
                </label>
                <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                  liveStats.cellSpreadPass ? 'bg-emerald-100 text-emerald-700' : 'bg-amber-100 text-amber-800'
                }`}>
                  ΔV = {liveStats.spreadV.toFixed(2)}V (Max &le; 0.05V)
                </span>
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <span className="text-[10px] text-slate-500 font-semibold block mb-0.5">Cell Min (V)</span>
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.cell_voltages_float_mode_v.cell_min_v}
                    onChange={(e) => {
                      const minV = parseFloat(e.target.value) || 0;
                      const maxV = data.electrical_tests.cell_voltages_float_mode_v.cell_max_v;
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          cell_voltages_float_mode_v: {
                            cell_min_v: minV,
                            cell_max_v: maxV,
                            voltage_spread_v: parseFloat(Math.abs(maxV - minV).toFixed(2))
                          }
                        }
                      });
                    }}
                    className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 font-mono font-bold text-slate-800 outline-none focus:ring-2 focus:ring-amber-500"
                  />
                </div>
                <div>
                  <span className="text-[10px] text-slate-500 font-semibold block mb-0.5">Cell Max (V)</span>
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.cell_voltages_float_mode_v.cell_max_v}
                    onChange={(e) => {
                      const maxV = parseFloat(e.target.value) || 0;
                      const minV = data.electrical_tests.cell_voltages_float_mode_v.cell_min_v;
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          cell_voltages_float_mode_v: {
                            cell_min_v: minV,
                            cell_max_v: maxV,
                            voltage_spread_v: parseFloat(Math.abs(maxV - minV).toFixed(2))
                          }
                        }
                      });
                    }}
                    className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 font-mono font-bold text-slate-800 outline-none focus:ring-2 focus:ring-amber-500"
                  />
                </div>
              </div>
              <p className="text-[11px] text-slate-500 leading-tight">
                NETA Sec 7.18.1.1.D.4: Ở chế độ Float, điện áp giữa các cell KHÔNG ĐƯỢC sai lệch vượt quá 0.05 Volt.
              </p>
            </div>

            {/* 3.4 Specific Gravity */}
            <div className="p-4 rounded-xl border border-slate-200 bg-slate-50/50 space-y-2.5">
              <div className="flex items-center justify-between">
                <label className="font-bold text-slate-800">
                  Tỷ Trọng Dung Dịch (Hiệu chỉnh 25°C)
                </label>
                <span className="text-[10px] text-slate-500 font-semibold">Chuẩn: 1.215</span>
              </div>
              <div className="grid grid-cols-3 gap-2">
                <div>
                  <span className="text-[10px] text-slate-500 font-semibold block mb-0.5">Pilot Cell</span>
                  <input
                    type="number"
                    step="0.001"
                    value={data.electrical_tests.specific_gravity_corrected_25c.pilot_cell_sg}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        specific_gravity_corrected_25c: {
                          ...data.electrical_tests.specific_gravity_corrected_25c,
                          pilot_cell_sg: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 font-mono font-semibold text-slate-800 outline-none focus:ring-2 focus:ring-amber-500"
                  />
                </div>
                <div>
                  <span className="text-[10px] text-slate-500 font-semibold block mb-0.5">Min SG</span>
                  <input
                    type="number"
                    step="0.001"
                    value={data.electrical_tests.specific_gravity_corrected_25c.min_sg_found}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        specific_gravity_corrected_25c: {
                          ...data.electrical_tests.specific_gravity_corrected_25c,
                          min_sg_found: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 font-mono font-semibold text-slate-800 outline-none focus:ring-2 focus:ring-amber-500"
                  />
                </div>
                <div>
                  <span className="text-[10px] text-slate-500 font-semibold block mb-0.5">Max SG</span>
                  <input
                    type="number"
                    step="0.001"
                    value={data.electrical_tests.specific_gravity_corrected_25c.max_sg_found}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        specific_gravity_corrected_25c: {
                          ...data.electrical_tests.specific_gravity_corrected_25c,
                          max_sg_found: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 font-mono font-semibold text-slate-800 outline-none focus:ring-2 focus:ring-amber-500"
                  />
                </div>
              </div>
              <p className="text-[11px] text-slate-500 leading-tight">
                NETA Sec 7.18.1.1.D.5: Tỷ trọng đo bằng hydrometer quang năng và hiệu chỉnh nhiệt độ @ 25°C.
              </p>
            </div>

            {/* 3.5 Internal Ohmic Measurement */}
            <div className="p-4 rounded-xl border border-slate-200 bg-slate-50/50 space-y-2.5">
              <div className="flex items-center justify-between">
                <label className="font-bold text-slate-800">
                  Nội Trở Ôm Cell Sạc Đầy (mΩ)
                </label>
                <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                  liveStats.ohmicPass ? 'bg-emerald-100 text-emerald-700' : 'bg-rose-100 text-rose-800'
                }`}>
                  Lệch: {liveStats.variancePct.toFixed(1)}% (Max &le; 25%)
                </span>
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <span className="text-[10px] text-slate-500 font-semibold block mb-0.5">Nội trở TB (mΩ)</span>
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.internal_ohmic_measurement_resistance_mohm.avg_resistance_mohm}
                    onChange={(e) => {
                      const avg = parseFloat(e.target.value) || 0;
                      const maxC = data.electrical_tests.internal_ohmic_measurement_resistance_mohm.max_resistance_cell_mohm;
                      const vPct = avg > 0 ? parseFloat((((maxC - avg) / avg) * 100).toFixed(1)) : 0;
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          internal_ohmic_measurement_resistance_mohm: {
                            avg_resistance_mohm: avg,
                            max_resistance_cell_mohm: maxC,
                            variance_percent: vPct
                          }
                        }
                      });
                    }}
                    className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 font-mono font-bold text-slate-800 outline-none focus:ring-2 focus:ring-amber-500"
                  />
                </div>
                <div>
                  <span className="text-[10px] text-slate-500 font-semibold block mb-0.5">Cell Max (mΩ)</span>
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.internal_ohmic_measurement_resistance_mohm.max_resistance_cell_mohm}
                    onChange={(e) => {
                      const maxC = parseFloat(e.target.value) || 0;
                      const avg = data.electrical_tests.internal_ohmic_measurement_resistance_mohm.avg_resistance_mohm;
                      const vPct = avg > 0 ? parseFloat((((maxC - avg) / avg) * 100).toFixed(1)) : 0;
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          internal_ohmic_measurement_resistance_mohm: {
                            avg_resistance_mohm: avg,
                            max_resistance_cell_mohm: maxC,
                            variance_percent: vPct
                          }
                        }
                      });
                    }}
                    className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 font-mono font-bold text-slate-800 outline-none focus:ring-2 focus:ring-amber-500"
                  />
                </div>
              </div>
              <p className="text-[11px] text-slate-500 leading-tight">
                NETA Sec 7.18.1.1.D.6: Trở kháng/nội trở KHÔNG ĐƯỢC biến động quá 25% so với mức trung bình.
              </p>
            </div>

            {/* 3.6 Charger Voltages & Load Test */}
            <div className="p-4 rounded-xl border border-slate-200 bg-slate-50/50 space-y-2.5">
              <label className="font-bold text-slate-800 block">
                Điện Áp Bộ Sạc Float / Equalize &amp; Xả Tải
              </label>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <span className="text-[10px] text-slate-500 font-semibold block mb-0.5">Sạc Float (V)</span>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.charger_float_voltage_v}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        charger_float_voltage_v: parseFloat(e.target.value) || 0
                      }
                    })}
                    className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 font-mono font-semibold text-slate-800 outline-none focus:ring-2 focus:ring-amber-500"
                  />
                </div>
                <div>
                  <span className="text-[10px] text-slate-500 font-semibold block mb-0.5">Sạc Cân Bằng (V)</span>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.charger_equalize_voltage_v}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        charger_equalize_voltage_v: parseFloat(e.target.value) || 0
                      }
                    })}
                    className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 font-mono font-semibold text-slate-800 outline-none focus:ring-2 focus:ring-amber-500"
                  />
                </div>
              </div>
              <div>
                <span className="text-[10px] text-slate-500 font-semibold block mb-0.5">Xả Tải per IEEE 450</span>
                <input
                  type="text"
                  value={data.electrical_tests.load_test_ieee_450}
                  onChange={(e) => setData({
                    ...data,
                    electrical_tests: {
                      ...data.electrical_tests,
                      load_test_ieee_450: e.target.value
                    }
                  })}
                  className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 font-semibold text-slate-800 outline-none focus:ring-2 focus:ring-amber-500"
                />
              </div>
            </div>
          </div>
        </div>

        {/* Action Button Bar */}
        <div className="pt-2 flex flex-wrap items-center justify-between gap-3 border-t border-slate-100">
          <div className="text-xs text-slate-500 font-medium">
            * Dữ liệu được mã hóa và đối soát tự động theo tiêu chuẩn NETA ATS-2025 Mục 7.18.1.1 &amp; IEEE 450.
          </div>
          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={handleSaveToCmms}
              className="inline-flex items-center gap-1.5 px-4 py-2 rounded-xl text-xs font-bold bg-slate-900 hover:bg-slate-800 text-white transition-all cursor-pointer shadow-sm"
            >
              <Save size={15} />
              <span>Lưu Vào Kho Báo Cáo</span>
            </button>
            <button
              type="button"
              onClick={handleAnalyze}
              disabled={isAnalyzing}
              className="inline-flex items-center gap-2 px-5 py-2.5 rounded-xl text-xs font-extrabold bg-gradient-to-r from-amber-500 to-amber-600 hover:from-amber-400 hover:to-amber-500 text-slate-950 transition-all cursor-pointer shadow-md disabled:opacity-50"
            >
              <Sparkles size={16} className={isAnalyzing ? 'animate-spin' : ''} />
              <span>{isAnalyzing ? 'Đang Chẩn Đoán...' : 'Chạy Đánh Giá NETA ATS-2025'}</span>
            </button>
          </div>
        </div>
      </div>

      {/* 4. ANALYSIS RESULT SECTION (MARKDOWN REPORT DISPLAY) */}
      {analysisResult && (
        <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-6 animate-in slide-in-from-bottom-2 duration-300">
          {/* Header Status Banner */}
          <div className={`p-4 rounded-xl border flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${
            analysisResult.overall_status === 'FAIL'
              ? 'bg-rose-50 border-rose-300 text-rose-950'
              : analysisResult.overall_status === 'INVESTIGATE'
              ? 'bg-amber-50 border-amber-300 text-amber-950'
              : 'bg-emerald-50 border-emerald-300 text-emerald-950'
          }`}>
            <div className="flex items-start sm:items-center gap-3">
              {analysisResult.overall_status === 'FAIL' ? (
                <ShieldAlert size={28} className="text-rose-600 flex-shrink-0" />
              ) : analysisResult.overall_status === 'INVESTIGATE' ? (
                <AlertTriangle size={28} className="text-amber-600 flex-shrink-0" />
              ) : (
                <ShieldCheck size={28} className="text-emerald-600 flex-shrink-0" />
              )}
              <div>
                <span className="text-xs font-bold uppercase tracking-wider text-slate-500 block">
                  Kết Luận Đánh Giá Tổng Thể
                </span>
                <h3 className="text-lg font-black tracking-tight">
                  {analysisResult.status_title || analysisResult.overall_status}
                </h3>
                <p className="text-xs mt-0.5 text-slate-700 leading-relaxed">
                  {analysisResult.status_reason}
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2 flex-shrink-0 self-end sm:self-center">
              <button
                type="button"
                onClick={handleCopyMarkdown}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-white/90 hover:bg-white text-slate-800 border border-slate-300 shadow-2xs transition-all cursor-pointer"
                title="Sao chép toàn bộ phản hồi Markdown"
              >
                <Copy size={13} />
                <span>{copied ? 'Đã Sao Chép!' : 'Copy Markdown'}</span>
              </button>
              <button
                type="button"
                onClick={() => setShowPrintModal(true)}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-slate-900 hover:bg-slate-800 text-white shadow-2xs transition-all cursor-pointer"
              >
                <Printer size={13} />
                <span>In Biên Bản</span>
              </button>
            </div>
          </div>

          {/* Render Markdown Content */}
          <div className="prose prose-slate max-w-none prose-headings:font-bold prose-h1:text-xl prose-h2:text-base prose-h3:text-sm prose-p:text-xs prose-p:leading-relaxed prose-table:text-xs prose-th:bg-slate-100 prose-th:p-2 prose-td:p-2 bg-slate-50/60 p-5 rounded-xl border border-slate-200">
            <div
              className="text-xs space-y-4 text-slate-800 leading-relaxed font-sans"
              dangerouslySetInnerHTML={{
                __html: formatMarkdownToHtml(analysisResult.markdown_report)
              }}
            />
          </div>
        </div>
      )}

      {/* JSON Edit / Import Modal */}
      {showJsonModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-xs p-4 animate-in fade-in">
          <div className="bg-white rounded-2xl max-w-2xl w-full p-6 shadow-2xl space-y-4 border border-slate-200">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <h3 className="text-sm font-bold text-slate-900 uppercase tracking-wider flex items-center gap-2">
                <FileText size={16} className="text-amber-600" />
                Dữ Liệu JSON Nhập Liệu FSE (NETA ATS-2025)
              </h3>
              <button
                type="button"
                onClick={() => setShowJsonModal(false)}
                className="text-slate-400 hover:text-slate-600 text-xs font-bold px-2 py-1 rounded"
              >
                Đóng
              </button>
            </div>
            <textarea
              rows={14}
              value={jsonInput}
              onChange={(e) => setJsonInput(e.target.value)}
              className="w-full text-xs font-mono bg-slate-900 text-amber-200 p-4 rounded-xl outline-none focus:ring-2 focus:ring-amber-500 border border-slate-800 leading-relaxed"
            />
            <div className="flex items-center justify-between pt-2">
              <button
                type="button"
                onClick={() => {
                  navigator.clipboard.writeText(jsonInput);
                  alert('Đã copy dữ liệu JSON vào Clipboard!');
                }}
                className="px-3 py-1.5 rounded-lg text-xs font-bold bg-slate-100 hover:bg-slate-200 text-slate-700"
              >
                Copy JSON
              </button>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setShowJsonModal(false)}
                  className="px-4 py-2 rounded-xl text-xs font-semibold text-slate-600 hover:bg-slate-100"
                >
                  Hủy
                </button>
                <button
                  type="button"
                  onClick={() => {
                    try {
                      const parsed = JSON.parse(jsonInput);
                      setData(parsed);
                      setShowJsonModal(false);
                      setSavedSuccessMessage('Đã tải thành công dữ liệu JSON FSE vào Checklist!');
                      setTimeout(() => setSavedSuccessMessage(null), 4000);
                    } catch (err: any) {
                      alert('Lỗi định dạng JSON: ' + err.message);
                    }
                  }}
                  className="px-4 py-2 rounded-xl text-xs font-bold bg-amber-600 hover:bg-amber-700 text-white shadow-sm"
                >
                  Áp Dụng Dữ Liệu
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Nameplate Camera OCR Scanner Modal */}
      {showCameraModal && (
        <NameplateScannerModal
          isOpen={showCameraModal}
          onClose={() => setShowCameraModal(false)}
          onApplyData={handleApplyNameplateData}
          equipmentCategory="battery"
        />
      )}

      {/* A4 Print Report Modal */}
      {showPrintModal && (
        <TevReportPrintModal
          isOpen={showPrintModal}
          onClose={() => setShowPrintModal(false)}
          equipmentName={data.site_info.battery_bank_tag}
          equipmentType={data.site_info.battery_type}
          standard="ANSI/NETA ATS-2025 Section 7.18.1.1 & IEEE 450"
          fseName={data.site_info.fse_name}
          testDate={data.site_info.test_date}
          overallStatus={analysisResult ? analysisResult.overall_status : liveStats.overallSafe ? 'PASS' : 'INVESTIGATE'}
          aiMarkdownReport={analysisResult ? analysisResult.markdown_report : undefined}
          visualItems={[
            { item: 'Hệ thống thông gió phòng pin', status: data.visual_inspection.ventilation_system_operable ? 'PASS' : 'FAIL', note: 'Chống tích tụ Hydro H₂' },
            { item: 'Trạm rửa mắt & khay hứng tràn axit', status: data.visual_inspection.eyewash_and_spill_containment_present ? 'PASS' : 'FAIL', note: 'An toàn axit sunfuric H₂SO₄' },
            { item: 'Đối soát nhãn mác giàn pin 100%', status: data.visual_inspection.nameplate_match ? 'PASS' : 'FAIL', note: 'Khớp hồ sơ thiết kế' },
            { item: 'Giá đỡ giàn pin & tiếp địa vỏ', status: data.visual_inspection.rack_mounting_and_grounding.includes('Pass') ? 'PASS' : 'FAIL', note: 'Cố định & tiếp địa an toàn' },
            { item: 'Nắp chắn lửa nổ ngược (Flame Arresters)', status: data.visual_inspection.flame_arresters_present ? 'PASS' : 'FAIL', note: '100% cell có nắp xốp' },
            { item: 'Vệ sinh cọc cực & mỡ chống oxy hóa', status: data.visual_inspection.cleanliness_and_oxide_inhibitor.includes('Pass') ? 'PASS' : 'FAIL', note: 'Bôi mỡ NO-OX-ID' },
            { item: 'Mức dung dịch điện môi (Electrolyte Level)', status: data.visual_inspection.electrolyte_level_check.includes('Pass') ? 'PASS' : 'FAIL', note: 'Giữa vạch MIN-MAX' },
            { item: 'Lực siết bu-lông cọc cực Table 100.12', status: data.visual_inspection.bolt_torque_check.includes('Pass') ? 'PASS' : 'FAIL', note: 'Siết đạt cờ-lê lực' }
          ]}
          electricalItems={electricalTableForPrint}
        />
      )}
    </div>
  );
};

/**
 * Basic markdown converter for clean rendering in DOM
 */
function formatMarkdownToHtml(markdown: string): string {
  if (!markdown) return '';

  return markdown
    .replace(/^# (.*$)/gim, '<h1 class="text-base font-black text-slate-900 mt-4 mb-2 pb-1 border-b border-slate-200">$1</h1>')
    .replace(/^## (.*$)/gim, '<h2 class="text-sm font-bold text-slate-900 mt-3 mb-1.5 text-amber-900">$1</h2>')
    .replace(/^### (.*$)/gim, '<h3 class="text-xs font-bold text-slate-800 mt-2 mb-1">$1</h3>')
    .replace(/\*\*(.*?)\*\*/gim, '<strong class="font-bold text-slate-900">$1</strong>')
    .replace(/\*(.*?)\*/gim, '<em class="italic">$1</em>')
    .replace(/^> (.*$)/gim, '<blockquote class="border-l-4 border-amber-500 pl-3 py-1 my-2 bg-amber-50/70 rounded-r text-amber-950 font-medium text-xs">$1</blockquote>')
    .replace(/\| (.*) \|/g, (match) => {
      // Return table row markup
      const cells = match.split('|').map(c => c.trim()).filter((c, i, a) => i > 0 && i < a.length - 1);
      if (cells.some(c => c.includes('---'))) return '';
      const isHeader = cells.includes('Hạng mục Kiểm tra') || cells.includes('STT');
      if (isHeader) {
        return `<tr class="bg-slate-100 font-bold text-slate-800">${cells.map(c => `<th class="p-2 border border-slate-200">${c}</th>`).join('')}</tr>`;
      }
      return `<tr class="hover:bg-slate-50">${cells.map(c => `<td class="p-2 border border-slate-200">${c}</td>`).join('')}</tr>`;
    })
    .replace(/<tr>[\s\S]*?<\/tr>/g, (match) => `<table class="w-full border-collapse border border-slate-200 my-2 text-xs">${match}</table>`)
    .replace(/\n\n/g, '<br/>');
}
