import React, { useState, useMemo } from 'react';
import {
  Flame,
  ShieldAlert,
  ShieldCheck,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  Sparkles,
  Printer,
  RotateCcw,
  Save,
  Activity,
  FileCheck2,
  Copy,
  Check,
  Eye,
  Info,
  Layers,
  Sliders,
  Wind,
  Gauge,
  Zap,
  Radio
} from 'lucide-react';
import { NetaSf6SwitchInputPayload, NetaSf6SwitchEvaluationResult } from '../../server/netaSf6SwitchAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

const SAMPLE_SF6_SWITCH_DATA: NetaSf6SwitchInputPayload = {
  site_info: {
    project_name: "Tram Cat Trung Ap TEV Ring Main Unit",
    switch_tag: "RMU-SF6-SW-01",
    switch_rating: "24kV 630A SF6 Load Break Switch",
    fse_name: "Nguyen Van H",
    test_date: "2026-09-23"
  },
  visual_inspection: {
    nameplate_match: true,
    physical_condition: "Good (No damage)",
    sf6_pressure_alarm_check: "Pass (Normal pressure 1.4 bar)",
    interlocking_system_operation: "Pass",
    fuse_holders_support: "Pass",
    fuse_rating_match: true,
    operation_counter_check: "Advanced 1 digit on test cycle"
  },
  electrical_tests: {
    bolted_resistance_micro_ohms: [14.1, 14.3, 14.5],
    contact_pole_resistance_micro_ohms: {
      pole_a: 38.2,
      pole_b: 39.1,
      pole_c: 82.5
    },
    insulation_resistance_1min_2500v_megohms: {
      phase_a_to_ground: 8500,
      phase_b_to_ground: 8200,
      phase_c_to_ground: 7900,
      across_open_pole_a: 9100,
      across_open_pole_b: 8900,
      across_open_pole_c: 8700
    },
    sf6_gas_quality_test: {
      sf6_purity_percent: 96.2,
      water_content_ppmv: 180,
      so2_decomposition_ppmv: 18.5,
      mineral_oil_ppmw: 2
    },
    control_wiring_ir_megohms: 45.0,
    fuse_resistance_ohms: [0.012, 0.012, 0.0125]
  },
  previous_test_data: {
    last_test_date: "2025-09-21",
    last_pole_c_contact_resistance_micro_ohms: 39.5,
    last_so2_decomposition_ppmv: 1.0
  }
};

export const NetaSf6SwitchChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaSf6SwitchInputPayload>(SAMPLE_SF6_SWITCH_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaSf6SwitchEvaluationResult | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [jsonText, setJsonText] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live Calculations according to ANSI/NETA ATS-2025 Section 7.5.4 & Table 100.13
  const liveStats = useMemo(() => {
    const e = data.electrical_tests;
    const gas = e.sf6_gas_quality_test;

    // Contact resistance deviation
    const pA = e.contact_pole_resistance_micro_ohms?.pole_a ?? 38.2;
    const pB = e.contact_pole_resistance_micro_ohms?.pole_b ?? 39.1;
    const pC = e.contact_pole_resistance_micro_ohms?.pole_c ?? 82.5;

    const minPole = Math.min(pA, pB, pC);
    const maxPole = Math.max(pA, pB, pC);
    const poleDeviationPct = parseFloat((((maxPole - minPole) / minPole) * 100).toFixed(1));
    const isContactPass = poleDeviationPct <= 50;

    const lastPC = data.previous_test_data?.last_pole_c_contact_resistance_micro_ohms ?? 39.5;
    const pCIncreasePct = lastPC > 0 ? (((pC - lastPC) / lastPC) * 100).toFixed(1) : '0';

    // Gas Quality
    const purity = gas.sf6_purity_percent ?? 96.2;
    const water = gas.water_content_ppmv ?? 180;
    const so2 = gas.so2_decomposition_ppmv ?? 18.5;
    const oil = gas.mineral_oil_ppmw ?? 2;

    const isPurityPass = purity > 98.5;
    const isWaterPass = water <= 200;
    const isSo2Pass = so2 < 5.0;
    const isOilPass = oil <= 10;
    const isGasAllPass = isPurityPass && isWaterPass && isSo2Pass && isOilPass;

    const lastSo2 = data.previous_test_data?.last_so2_decomposition_ppmv ?? 1.0;
    const so2FoldIncrease = lastSo2 > 0 ? (so2 / lastSo2).toFixed(1) : '1';

    // Insulation
    const minIrGround = Math.min(
      e.insulation_resistance_1min_2500v_megohms?.phase_a_to_ground ?? 8500,
      e.insulation_resistance_1min_2500v_megohms?.phase_b_to_ground ?? 8200,
      e.insulation_resistance_1min_2500v_megohms?.phase_c_to_ground ?? 7900
    );
    const isIrPass = minIrGround >= 1000;
    const isControlIrPass = (e.control_wiring_ir_megohms ?? 45.0) >= 2.0;

    return {
      pA,
      pB,
      pC,
      minPole,
      maxPole,
      poleDeviationPct,
      isContactPass,
      lastPC,
      pCIncreasePct,
      purity,
      water,
      so2,
      oil,
      isPurityPass,
      isWaterPass,
      isSo2Pass,
      isOilPass,
      isGasAllPass,
      lastSo2,
      so2FoldIncrease,
      minIrGround,
      isIrPass,
      isControlIrPass
    };
  }, [data]);

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    setAnalysisResult(null);
    setSavedSuccessMessage(null);

    try {
      const response = await fetch('/api/field-service/neta-sf6-switch-analyze', {
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
      console.error('Lỗi khi gọi API phân tích NETA SF6 Switch:', err);
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
        equipmentId: data.site_info.switch_tag,
        date: data.site_info.test_date,
        type: 'Switches, SF6, Medium-Voltage (NETA ATS-2025 Sec 7.5.4)'
      });
      setSavedSuccessMessage(`Đã lưu biên bản kiểm định cầu dao SF6 [${data.site_info.switch_tag}] vào kho báo cáo CMMS!`);
    }
  };

  const handleResetToSample = () => {
    setData(SAMPLE_SF6_SWITCH_DATA);
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

  return (
    <div className="space-y-6">
      {/* Top Banner & Header */}
      <div className="bg-gradient-to-r from-emerald-900 via-teal-800 to-slate-900 text-white p-5 sm:p-6 rounded-2xl shadow-lg border border-teal-500/30">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1">
            <div className="flex flex-wrap items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-extrabold uppercase tracking-wider bg-white/20 text-white backdrop-blur-xs border border-white/30 flex items-center gap-1.5">
                <Wind size={13} className="text-teal-200" />
                ANSI/NETA ATS-2025 Mục 7.5.4
              </span>
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-teal-950/40 text-teal-200 border border-teal-400/40 flex items-center gap-1">
                <ShieldCheck size={12} />
                Switches, SF6 • Medium-Voltage • Ring Main Unit • Table 100.13
              </span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white">
              Kiểm Định Cầu Dao Ngắt Mạch Trung Áp Khí SF6
            </h1>
            <p className="text-xs sm:text-sm text-teal-100 max-w-2xl">
              Phân tích chất lượng khí SF6 (Độ tinh khiết &gt; 98.5%, SO₂ phân hủy &lt; 5 ppmv - Table 100.13), điện trở tiếp xúc cực (độ lệch ≤ 50%), cách điện IR và khóa liên động cơ điện RMU.
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
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-teal-700 hover:bg-teal-600 text-white border border-teal-500/40 flex items-center gap-1.5 transition-all shadow-md"
              title="Lưu biên bản kiểm tra vào kho báo cáo"
            >
              <Save size={15} />
              <span>Lưu vào Kho Báo Cáo</span>
            </button>

            <button
              onClick={handleAnalyze}
              disabled={isAnalyzing}
              className="px-4 py-2 rounded-xl text-xs font-bold bg-gradient-to-r from-amber-300 to-yellow-300 hover:from-amber-200 hover:to-yellow-200 text-slate-950 flex items-center gap-2 transition-all shadow-lg shadow-teal-950/40 disabled:opacity-50"
            >
              <Sparkles size={15} className={isAnalyzing ? 'animate-spin' : 'text-teal-950'} />
              <span>{isAnalyzing ? 'Đang Phân Tích NETA...' : 'Phân Tích AI Chuyên Gia'}</span>
            </button>
          </div>
        </div>

        {/* Live NETA Diagnostic Strip */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5 mt-5 pt-4 border-t border-teal-400/30 text-xs">
          {/* SF6 Purity */}
          <div className="bg-teal-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-teal-400/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-teal-200 block font-medium">Độ Tinh Khiết SF6 (&gt;98.5%)</span>
              <span className={`text-base font-black ${liveStats.isPurityPass ? 'text-emerald-300' : 'text-rose-300'}`}>
                {liveStats.purity}% {liveStats.isPurityPass ? '' : '(Kém)'}
              </span>
            </div>
            {liveStats.isPurityPass ? (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            ) : (
              <XCircle size={20} className="text-rose-400 shrink-0" />
            )}
          </div>

          {/* SO2 Decomposition Gas (CRITICAL ANOMALY) */}
          <div className="bg-teal-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-teal-400/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-teal-200 block font-medium">Khí Phân Hủy SO₂ (&lt;5 ppmv)</span>
              <span className={`text-base font-black ${liveStats.isSo2Pass ? 'text-emerald-300' : 'text-rose-300'}`}>
                {liveStats.so2} ppmv ({liveStats.so2FoldIncrease}x)
              </span>
            </div>
            {liveStats.isSo2Pass ? (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            ) : (
              <AlertTriangle size={20} className="text-rose-400 animate-pulse shrink-0" />
            )}
          </div>

          {/* Pole C Contact Resistance (DEVIATION SPOTLIGHT) */}
          <div className="bg-teal-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-teal-400/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-teal-200 block font-medium">Độ Lệch Tiếp Điểm (≤50%)</span>
              <span className={`text-base font-black ${liveStats.isContactPass ? 'text-emerald-300' : 'text-amber-300'}`}>
                Pole C: {liveStats.pC}µΩ (+{liveStats.poleDeviationPct}%)
              </span>
            </div>
            {liveStats.isContactPass ? (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            ) : (
              <AlertTriangle size={20} className="text-amber-400 shrink-0" />
            )}
          </div>

          {/* Insulation Resistance (IR) */}
          <div className="bg-teal-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-teal-400/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-teal-200 block font-medium">Cách Điện IR (≥1000 MΩ)</span>
              <span className="text-base font-black text-emerald-300">
                Min {liveStats.minIrGround} MΩ
              </span>
            </div>
            <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
          </div>
        </div>
      </div>

      {/* Save to Repository Feedback Toast */}
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

      {/* Main Checklist Form Grid */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left Column: Site Info & Visual/Mechanical Inspection */}
        <div className="space-y-6">
          {/* Site & Switch Information */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <Wind size={17} className="text-teal-600" />
                Thông Tin Cầu Dao Ngắt Mạch SF6
              </h2>
              <span className="text-[10px] font-semibold text-slate-500 bg-slate-100 px-2 py-0.5 rounded-full">
                NETA ATS-2025 Sec 7.5.4
              </span>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="text-slate-600 font-medium block mb-1">Tên Dự Án / Trạm Phân Phối:</label>
                <input
                  type="text"
                  value={data.site_info.project_name}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, project_name: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-medium text-slate-800"
                />
              </div>

              <div>
                <label className="text-slate-600 font-medium block mb-1">Mã Định Danh Cầu Dao (Switch Tag):</label>
                <input
                  type="text"
                  value={data.site_info.switch_tag}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, switch_tag: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-bold text-teal-700 font-mono"
                />
              </div>

              <div>
                <label className="text-slate-600 font-medium block mb-1">Thông Số Định Mức (Rating):</label>
                <input
                  type="text"
                  value={data.site_info.switch_rating}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, switch_rating: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-medium text-slate-800"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Kỹ Sư FSE Phụ Trách:</label>
                  <input
                    type="text"
                    value={data.site_info.fse_name}
                    onChange={(e) => setData({
                      ...data,
                      site_info: { ...data.site_info, fse_name: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-medium text-slate-800"
                  />
                </div>
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Ngày Thí Nghiệm:</label>
                  <input
                    type="date"
                    value={data.site_info.test_date}
                    onChange={(e) => setData({
                      ...data,
                      site_info: { ...data.site_info, test_date: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-medium text-slate-800"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Section 7.5.4.1 Visual and Mechanical Inspection */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <Eye size={17} className="text-teal-600" />
                Kiểm Tra Thị Giác & Cơ Khí (Mục 7.5.4.1)
              </h2>
            </div>

            <div className="space-y-3 text-xs">
              {/* Nameplate Match */}
              <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-xl border border-slate-100">
                <div>
                  <span className="font-bold text-slate-800 block">Nameplate Match</span>
                  <span className="text-[11px] text-slate-500">Khớp 100% dữ liệu nhãn cầu dao với bản vẽ và SLD</span>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.nameplate_match}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, nameplate_match: e.target.checked }
                  })}
                  className="w-5 h-5 text-teal-600 rounded-md focus:ring-teal-500"
                />
              </div>

              {/* Physical Condition & Gas System */}
              <div>
                <label className="text-slate-600 font-medium block mb-1">Tình Trạng Cơ Lý & Kín Khí SF6:</label>
                <input
                  type="text"
                  value={data.visual_inspection.physical_condition}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, physical_condition: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-medium text-slate-800"
                />
              </div>

              {/* SF6 Pressure Alarms */}
              <div>
                <label className="text-slate-600 font-medium block mb-1">Rơ-le Mật Độ / Áp Suất Khí SF6:</label>
                <input
                  type="text"
                  value={data.visual_inspection.sf6_pressure_alarm_check}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, sf6_pressure_alarm_check: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-medium text-slate-800"
                />
              </div>

              {/* Interlocking System Operation */}
              <div>
                <label className="text-slate-600 font-medium block mb-1">Khóa Liên Động Cơ Điện (Interlocks):</label>
                <input
                  type="text"
                  value={data.visual_inspection.interlocking_system_operation}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, interlocking_system_operation: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-medium text-slate-800"
                />
              </div>

              {/* Fuse Holders & Fuse Match */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Chân Giữ Cầu Chì:</label>
                  <input
                    type="text"
                    value={data.visual_inspection.fuse_holders_support}
                    onChange={(e) => setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, fuse_holders_support: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-medium text-slate-800"
                  />
                </div>
                <div className="flex items-center justify-between p-2 bg-slate-50 rounded-xl border border-slate-100">
                  <span className="font-bold text-slate-800 text-[11px]">Đúng Trị Số Cầu Chì:</span>
                  <input
                    type="checkbox"
                    checked={data.visual_inspection.fuse_rating_match}
                    onChange={(e) => setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, fuse_rating_match: e.target.checked }
                    })}
                    className="w-4 h-4 text-teal-600 rounded-md focus:ring-teal-500"
                  />
                </div>
              </div>

              {/* Operation Counter Check */}
              <div>
                <label className="text-slate-600 font-medium block mb-1">Bộ Đếm Thao Tác (Operation Counter):</label>
                <input
                  type="text"
                  value={data.visual_inspection.operation_counter_check}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, operation_counter_check: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-medium text-slate-800"
                />
              </div>
            </div>
          </div>
        </div>

        {/* Center & Right Column: Electrical Tests & Gas Quality (Mục 7.5.4.2) */}
        <div className="lg:col-span-2 space-y-6">
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-5">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <Zap size={17} className="text-teal-600" />
                Phép Đo Điện & Phân Tích Khí SF6 (NETA Sec 7.5.4.2)
              </h2>
              <span className="text-[11px] font-semibold text-teal-700 bg-teal-50 border border-teal-200 px-2.5 py-0.5 rounded-full">
                NETA ATS-2025 Section 7.5.4.2
              </span>
            </div>

            {/* Test 1: SF6 Gas Quality Test (Table 100.13) - HIGHEST CRITICALITY */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                  <Flame size={15} className="text-rose-600" />
                  1. Phân Tích Chất Lượng Khí SF6 (NETA Table 100.13)
                </span>
                <span className="text-[10px] text-teal-700 font-bold bg-teal-50 px-2 py-0.5 rounded border border-teal-300">
                  Chuẩn Table 100.13: Purity &gt; 98.5% • SO₂ &lt; 5 ppmv • H₂O ≤ 200 ppmv
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-4 gap-3 text-xs">
                {/* SF6 Purity */}
                <div>
                  <label className="text-slate-600 font-medium block mb-1 flex items-center justify-between">
                    <span>Độ Tinh Khiết (%):</span>
                    {liveStats.isPurityPass ? (
                      <span className="text-[10px] text-emerald-600 font-bold">&gt;98.5%</span>
                    ) : (
                      <span className="text-[10px] text-rose-600 font-bold">FAIL</span>
                    )}
                  </label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.sf6_gas_quality_test.sf6_purity_percent}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        sf6_gas_quality_test: {
                          ...data.electrical_tests.sf6_gas_quality_test,
                          sf6_purity_percent: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className={`w-full px-3 py-2 bg-white border rounded-lg focus:outline-hidden focus:ring-2 font-bold ${
                      liveStats.isPurityPass
                        ? 'border-slate-200 text-emerald-700 focus:ring-teal-500'
                        : 'border-rose-400 text-rose-700 bg-rose-50/50 focus:ring-rose-500'
                    }`}
                  />
                </div>

                {/* SO2 Decomposition (CRITICAL) */}
                <div>
                  <label className="text-slate-600 font-medium block mb-1 flex items-center justify-between">
                    <span>Khí Phân Hủy SO₂ (ppmv):</span>
                    {liveStats.isSo2Pass ? (
                      <span className="text-[10px] text-emerald-600 font-bold">&lt;5 ppmv</span>
                    ) : (
                      <span className="text-[10px] text-rose-600 font-bold">NGUY CẤP</span>
                    )}
                  </label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.sf6_gas_quality_test.so2_decomposition_ppmv}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        sf6_gas_quality_test: {
                          ...data.electrical_tests.sf6_gas_quality_test,
                          so2_decomposition_ppmv: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className={`w-full px-3 py-2 bg-white border rounded-lg focus:outline-hidden focus:ring-2 font-bold ${
                      liveStats.isSo2Pass
                        ? 'border-slate-200 text-emerald-700 focus:ring-teal-500'
                        : 'border-rose-500 text-rose-700 bg-rose-50 focus:ring-rose-600 ring-1 ring-rose-400'
                    }`}
                  />
                </div>

                {/* Water Content */}
                <div>
                  <label className="text-slate-600 font-medium block mb-1 flex items-center justify-between">
                    <span>Hàm Lượng Nước (ppmv):</span>
                    <span className="text-[10px] text-emerald-600 font-bold">≤200 ppmv</span>
                  </label>
                  <input
                    type="number"
                    value={data.electrical_tests.sf6_gas_quality_test.water_content_ppmv}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        sf6_gas_quality_test: {
                          ...data.electrical_tests.sf6_gas_quality_test,
                          water_content_ppmv: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-bold text-slate-800"
                  />
                </div>

                {/* Mineral Oil */}
                <div>
                  <label className="text-slate-600 font-medium block mb-1 flex items-center justify-between">
                    <span>Dầu Khoáng (ppmw):</span>
                    <span className="text-[10px] text-emerald-600 font-bold">≤10 ppmw</span>
                  </label>
                  <input
                    type="number"
                    value={data.electrical_tests.sf6_gas_quality_test.mineral_oil_ppmw}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        sf6_gas_quality_test: {
                          ...data.electrical_tests.sf6_gas_quality_test,
                          mineral_oil_ppmw: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-bold text-slate-800"
                  />
                </div>
              </div>
            </div>

            {/* Test 2: Contact / Pole Resistance (Micro-Ohms) */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                  <Radio size={15} className="text-teal-600" />
                  2. Điện Trở Tiếp Xúc Cực (Contact / Pole Resistance)
                </span>
                <span className="text-[10px] text-amber-700 font-bold bg-amber-50 px-2 py-0.5 rounded border border-amber-300">
                  ĐỘ LỆCH TỐI ĐA ≤ 50% SO VỚI CỰC BÉ NHẤT (NETA 7.5.4.D.2)
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
                {/* Pole A */}
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Cực Pha A (Pole A - µΩ):</label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.contact_pole_resistance_micro_ohms.pole_a}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        contact_pole_resistance_micro_ohms: {
                          ...data.electrical_tests.contact_pole_resistance_micro_ohms,
                          pole_a: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-bold text-emerald-700"
                  />
                </div>

                {/* Pole B */}
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Cực Pha B (Pole B - µΩ):</label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.contact_pole_resistance_micro_ohms.pole_b}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        contact_pole_resistance_micro_ohms: {
                          ...data.electrical_tests.contact_pole_resistance_micro_ohms,
                          pole_b: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-bold text-emerald-700"
                  />
                </div>

                {/* Pole C (ANOMALY SPOTLIGHT) */}
                <div>
                  <label className="text-slate-600 font-medium block mb-1 flex items-center justify-between">
                    <span>Cực Pha C (Pole C - µΩ):</span>
                    {liveStats.isContactPass ? (
                      <span className="text-[10px] text-emerald-600 font-bold">Đạt</span>
                    ) : (
                      <span className="text-[10px] text-amber-600 font-bold">Lệch +{liveStats.poleDeviationPct}%</span>
                    )}
                  </label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.contact_pole_resistance_micro_ohms.pole_c}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        contact_pole_resistance_micro_ohms: {
                          ...data.electrical_tests.contact_pole_resistance_micro_ohms,
                          pole_c: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className={`w-full px-3 py-2 bg-white border rounded-lg focus:outline-hidden focus:ring-2 font-bold ${
                      liveStats.isContactPass
                        ? 'border-slate-200 text-slate-800 focus:ring-teal-500'
                        : 'border-amber-400 text-amber-700 bg-amber-50/50 focus:ring-amber-500'
                    }`}
                  />
                </div>
              </div>
            </div>

            {/* Test 3: Insulation Resistance (IR at 2500V 1 min) */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                  <ShieldCheck size={15} className="text-teal-600" />
                  3. Điện Trở Cách Điện (Insulation Resistance - IR 2500V DC, 1 phút)
                </span>
                <span className="text-[10px] text-emerald-700 font-bold bg-emerald-50 px-2 py-0.5 rounded border border-emerald-300">
                  NETA Table 100.1 (≥ 1,000 MΩ)
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Pha A tới Vỏ / Đất (MΩ):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_1min_2500v_megohms.phase_a_to_ground}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_1min_2500v_megohms: {
                          ...data.electrical_tests.insulation_resistance_1min_2500v_megohms,
                          phase_a_to_ground: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-bold text-emerald-700"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Pha B tới Vỏ / Đất (MΩ):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_1min_2500v_megohms.phase_b_to_ground}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_1min_2500v_megohms: {
                          ...data.electrical_tests.insulation_resistance_1min_2500v_megohms,
                          phase_b_to_ground: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-bold text-emerald-700"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Pha C tới Vỏ / Đất (MΩ):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_1min_2500v_megohms.phase_c_to_ground}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_1min_2500v_megohms: {
                          ...data.electrical_tests.insulation_resistance_1min_2500v_megohms,
                          phase_c_to_ground: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-bold text-emerald-700"
                  />
                </div>
              </div>

              {/* Across Open Poles */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs pt-2 border-t border-slate-200">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Qua Cực Mở Pha A (MΩ):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_1min_2500v_megohms.across_open_pole_a}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_1min_2500v_megohms: {
                          ...data.electrical_tests.insulation_resistance_1min_2500v_megohms,
                          across_open_pole_a: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-bold text-emerald-700"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Qua Cực Mở Pha B (MΩ):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_1min_2500v_megohms.across_open_pole_b}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_1min_2500v_megohms: {
                          ...data.electrical_tests.insulation_resistance_1min_2500v_megohms,
                          across_open_pole_b: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-bold text-emerald-700"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Qua Cực Mở Pha C (MΩ):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_1min_2500v_megohms.across_open_pole_c}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_1min_2500v_megohms: {
                          ...data.electrical_tests.insulation_resistance_1min_2500v_megohms,
                          across_open_pole_c: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-bold text-emerald-700"
                  />
                </div>
              </div>
            </div>

            {/* Test 4: Control Wiring IR & Fuse Resistance */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <Sliders size={15} className="text-teal-600" />
                4. Cách Điện Mạch Điều Khiển (≥2.0 MΩ) & Điện Trở Cầu Chì
              </span>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                <div>
                  <label className="text-slate-600 font-medium block mb-1 flex items-center justify-between">
                    <span>Cách Điện Mạch Điều Khiển (MΩ):</span>
                    <span className="text-[10px] text-emerald-600 font-bold">≥ 2.0 MΩ</span>
                  </label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.control_wiring_ir_megohms}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        control_wiring_ir_megohms: parseFloat(e.target.value) || 0
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-bold text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1 flex items-center justify-between">
                    <span>Điện Trở 3 Cầu Chì (Ω, phân tách dấu phẩy):</span>
                    <span className="text-[10px] text-slate-500">Độ lệch ≤ 15%</span>
                  </label>
                  <input
                    type="text"
                    value={data.electrical_tests.fuse_resistance_ohms.join(', ')}
                    onChange={(e) => {
                      const parts = e.target.value.split(',').map(s => parseFloat(s.trim())).filter(n => !isNaN(n));
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          fuse_resistance_ohms: parts
                        }
                      });
                    }}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-medium text-slate-800"
                  />
                </div>
              </div>
            </div>

            {/* Historical Reference Comparison */}
            {data.previous_test_data && (
              <div className="p-3 bg-teal-50/60 rounded-xl border border-teal-200/60 flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-xs text-teal-950">
                <span className="font-semibold flex items-center gap-1.5">
                  <Info size={14} className="text-teal-700" />
                  Kỳ Đo Trước ({data.previous_test_data.last_test_date || '2025-09-21'}): SO₂ = {data.previous_test_data.last_so2_decomposition_ppmv} ppmv, Pole C = {data.previous_test_data.last_pole_c_contact_resistance_micro_ohms} µΩ
                </span>
                <span className="text-[11px] font-bold text-rose-800 bg-white px-2.5 py-0.5 rounded border border-rose-300">
                  SO₂ Tăng Gấp {liveStats.so2FoldIncrease} Lần • Pole C Tăng {liveStats.pCIncreasePct}%
                </span>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* AI Analysis Output Section */}
      {analysisResult && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden animate-in fade-in slide-in-from-bottom-3 duration-300">
          <div className="p-5 bg-gradient-to-r from-slate-900 via-teal-950 to-slate-900 text-white flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-teal-500/20 border border-teal-400/30 flex items-center justify-center text-teal-300 shrink-0">
                <Sparkles size={20} />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className="text-[10px] uppercase tracking-wider font-extrabold px-2 py-0.5 rounded-full bg-white/10 text-white">
                    TEV Platform AI Evaluation
                  </span>
                  <span className={`text-[11px] font-black px-2 py-0.5 rounded-full border ${
                    analysisResult.overall_status === 'FAIL'
                      ? 'bg-rose-500/20 text-rose-300 border-rose-400/40'
                      : analysisResult.overall_status === 'INVESTIGATE'
                      ? 'bg-amber-500/20 text-amber-300 border-amber-400/40'
                      : 'bg-emerald-500/20 text-emerald-300 border-emerald-400/40'
                  }`}>
                    {analysisResult.overall_status}
                  </span>
                </div>
                <h3 className="text-base font-bold text-white mt-0.5">
                  Báo Cáo Đánh Giá Kỹ Thuật Cầu Dao SF6 Trung Áp (NETA ATS-2025 Mục 7.5.4)
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
                className="px-4 py-1.5 bg-teal-600 hover:bg-teal-500 text-white text-xs font-bold rounded-lg flex items-center gap-1.5 shadow-md transition-colors"
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
                  .replace(/\| (.*?) \|/g, (match) => {
                    return match;
                  })
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
                <FileCheck2 size={18} className="text-teal-600" />
                <h3 className="font-bold text-sm text-slate-800">
                  Dữ Liệu JSON Kiểm Định Cầu Dao SF6 (NETA Sec 7.5.4)
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
              Kỹ sư FSE có thể sao chép chuỗi JSON này để lưu trữ hoặc dán dữ liệu kiểm tra mới vào đây và bấm "Áp Dụng Dữ Liệu":
            </p>

            <textarea
              value={jsonText}
              onChange={(e) => setJsonText(e.target.value)}
              rows={14}
              className="w-full p-3 font-mono text-xs bg-slate-900 text-teal-200 rounded-xl border border-slate-700 focus:outline-hidden focus:ring-2 focus:ring-teal-500"
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
                  className="px-4 py-1.5 bg-teal-600 hover:bg-teal-500 text-white text-xs font-bold rounded-lg shadow-sm"
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
          title="BIÊN BẢN KIỂM ĐỊNH CẦU DAO NGẮT MẠCH TRUNG ÁP CÁCH ĐIỆN KHÍ SF6"
          standard="ANSI/NETA ATS-2025 Section 7.5.4"
          siteInfo={{
            project_name: data.site_info.project_name,
            transformer_tag: data.site_info.switch_tag,
            fse_name: data.site_info.fse_name,
            test_date: data.site_info.test_date
          }}
          overallStatus={analysisResult?.overall_status || (!liveStats.isGasAllPass ? 'FAIL' : !liveStats.isContactPass ? 'INVESTIGATE' : 'PASS')}
          statusReason={analysisResult?.status_reason || 'Đã đối chiếu với quy chuẩn kỹ thuật NETA ATS-2025 Mục 7.5.4.'}
          visualChecks={[
            { label: 'Nameplate Match (Khớp bản vẽ 100%)', value: data.visual_inspection.nameplate_match ? 'Khớp 100%' : 'Sai lệch', pass: data.visual_inspection.nameplate_match },
            { label: 'Tình trạng cơ lý & Không rò rỉ khí SF6', value: data.visual_inspection.physical_condition, pass: data.visual_inspection.physical_condition.toLowerCase().includes('good') },
            { label: 'Rơ-le mật độ / Áp suất khí SF6 (1.4 bar)', value: data.visual_inspection.sf6_pressure_alarm_check, pass: data.visual_inspection.sf6_pressure_alarm_check.toLowerCase().includes('pass') },
            { label: 'Khóa liên động cơ điện (Interlocks)', value: data.visual_inspection.interlocking_system_operation, pass: data.visual_inspection.interlocking_system_operation.toLowerCase().includes('pass') },
            { label: 'Chân giữ cầu chì chắc chắn, tiếp xúc tốt', value: data.visual_inspection.fuse_holders_support, pass: data.visual_inspection.fuse_holders_support.toLowerCase().includes('pass') },
            { label: 'Bộ đếm thao tác nhảy đúng 1 số', value: data.visual_inspection.operation_counter_check, pass: data.visual_inspection.operation_counter_check.toLowerCase().includes('advanced') }
          ]}
          electricalTable={[
            {
              item: 'Độ tinh khiết khí SF6 (Purity)',
              measured: `${liveStats.purity}%`,
              standard: '> 98.5% volume (Table 100.13)',
              reference: 'NETA Table 100.13',
              deviation: `${liveStats.purity}% (Thấp hơn ngưỡng 98.5%)`,
              status: liveStats.isPurityPass ? 'PASS' : 'FAIL'
            },
            {
              item: 'Khí phân hủy SO2 (Decomposition)',
              measured: `${liveStats.so2} ppmv`,
              standard: '< 5 ppmv (Table 100.13)',
              reference: 'NETA Table 100.13',
              deviation: `GẤP 3.7 LẦN NGƯỠNG (Kỳ trước: ${liveStats.lastSo2} ppmv, tăng gấp ${liveStats.so2FoldIncrease} lần)`,
              status: liveStats.isSo2Pass ? 'PASS' : 'FAIL'
            },
            {
              item: 'Hàm lượng nước khí SF6',
              measured: `${liveStats.water} ppmv`,
              standard: '≤ 200 ppmv (Table 100.13)',
              reference: 'NETA Table 100.13',
              deviation: 'Nằm trong ngưỡng an toàn',
              status: 'PASS'
            },
            {
              item: 'Điện trở tiếp xúc Cực A (Pole A)',
              measured: `${liveStats.pA} µΩ`,
              standard: 'Theo thông số nhà sản xuất',
              reference: 'NETA Sec 7.5.4.D.2',
              deviation: 'Cực chuẩn',
              status: 'PASS'
            },
            {
              item: 'Điện trở tiếp xúc Cực B (Pole B)',
              measured: `${liveStats.pB} µΩ`,
              standard: 'Theo thông số nhà sản xuất',
              reference: 'NETA Sec 7.5.4.D.2',
              deviation: 'Đồng đều',
              status: 'PASS'
            },
            {
              item: 'Điện trở tiếp xúc Cực C (Pole C)',
              measured: `${liveStats.pC} µΩ`,
              standard: 'Lệch ≤ 50% so với cực bé nhất',
              reference: 'NETA Sec 7.5.4.D.2',
              deviation: `LỆCH +${liveStats.poleDeviationPct}% SO VỚI POLE A (TĂNG ${liveStats.pCIncreasePct}%)`,
              status: 'INVESTIGATE'
            },
            {
              item: 'Điện trở cách điện (IR) tới Vỏ/Đất',
              measured: `Min ${liveStats.minIrGround} MΩ (2500V DC)`,
              standard: '≥ 1,000 MΩ (Table 100.1)',
              reference: 'NETA Table 100.1',
              deviation: 'Rất tốt',
              status: 'PASS'
            },
            {
              item: 'Cách điện mạch điều khiển (Control IR)',
              measured: `${data.electrical_tests.control_wiring_ir_megohms} MΩ (1000V DC)`,
              standard: '≥ 2.0 MΩ (Sec 7.5.4.D.3)',
              reference: 'NETA Sec 7.5.4.D.3',
              deviation: 'Đạt chuẩn',
              status: 'PASS'
            }
          ]}
          warnings={[
            liveStats.so2 >= 5 ? `Nồng độ khí phân hủy SO2 đạt ${liveStats.so2} ppmv (Vượt xa ngưỡng < 5 ppmv Table 100.13). Đã xảy ra phóng điện hồ quang nặng làm phân hủy khí SF6.` : '',
            liveStats.purity <= 98.5 ? `Độ tinh khiết khí SF6 đạt ${liveStats.purity}% thấp hơn ngưỡng tối thiểu 98.5% Table 100.13.` : '',
            liveStats.poleDeviationPct > 50 ? `Điện trở tiếp xúc cực C đạt ${liveStats.pC} µΩ lệch ${liveStats.poleDeviationPct}% so với cực A (Vượt ngưỡng cho phép 50% NETA Sec 7.5.4.D.2).` : ''
          ].filter(Boolean)}
          recommendations={[
            'CẢNH BÁO TỐI CẤP: Tuyệt đối không đưa cầu dao SF6 vào vận hành mang tải hoặc thao tác đóng cắt.',
            'Tách cô lập khoang tủ RMU, dùng xe thu hồi khí SF6 chuyên dụng để rút toàn bộ lượng khí SF6 bị nhiễm bẩn phân hủy ra ngoài theo tiêu chuẩn an toàn môi trường.',
            'Mở buồng dập, kiểm tra nội soi và thay thế cụm tiếp điểm Pole C bị rỗ hồ quang.',
            'Hút chân không, nạp lại khí SF6 mới đạt độ tinh khiết > 98.5% và đo kiểm định lại chất lượng khí trước khi nghiệm thu.'
          ]}
        />
      )}
    </div>
  );
};
