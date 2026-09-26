import React, { useState, useMemo } from 'react';
import {
  RotateCcw,
  Activity,
  AlertTriangle,
  AlertOctagon,
  CheckCircle2,
  XCircle,
  Sparkles,
  Printer,
  Save,
  FileCheck2,
  Copy,
  Check,
  Info,
  Layers,
  Cpu,
  Gauge,
  TrendingUp,
  ShieldCheck,
  Zap,
  Wind
} from 'lucide-react';
import {
  NetaMotorInputPayload,
  NetaMotorEvaluationResult,
  performDeterministicMotorAnalysis
} from '../../server/netaMotorAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';
import { FseEngineerDropdown } from './FseEngineerDropdown';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

const SAMPLE_MOTOR_DATA: NetaMotorInputPayload = {
  site_info: {
    project_name: "Tram Bom Nuoc TEV Industrial Plant",
    motor_tag: "MOT-PUMP-4160V-01",
    motor_rating: "4160V 500HP (375kW) 3-Phase Induction Motor",
    fse_name: "Võ Minh Tú (sgm1707@gmail.com)",
    test_date: "2026-09-23"
  },
  visual_inspection: {
    nameplate_match: true,
    physical_condition: "Good (Clean baffles and windings)",
    anchorage_alignment_grounding: "Pass",
    air_gap_and_runout_check: "Pass (Within limits)",
    lubrication_bearings: "Pass",
    bolt_torque_check: "Pass",
    rtd_circuits_check: "Pass"
  },
  electrical_tests: {
    bolted_resistance_micro_ohms: [10.2, 10.5, 10.8],
    insulation_resistance_10min_2500v_megohms: {
      ir_1min_40c_corrected: 850,
      ir_10min_40c_corrected: 2125,
      pi_calculated: 2.5
    },
    stator_resistance_phase_to_phase_ohms: {
      phase_ab: 0.412,
      phase_bc: 0.415,
      phase_ca: 0.448
    },
    dielectric_withstand_test_ieee95: "Pass (No breakdown at 9.3kV DC)",
    surge_comparison_test: "Pass (Waveforms perfectly nested)",
    vibration_test_table_100_10: {
      velocity_in_sec_pk: 0.22,
      nema_grade_limit: 0.15
    },
    current_signature_analysis_mcsa: {
      broken_bar_sideband_db: 35.2
    },
    space_heater_operation: "Pass"
  },
  previous_test_data: {
    last_test_date: "2025-09-20",
    last_stator_resistance_ab: 0.410,
    last_vibration_in_sec: 0.11
  }
};

export const NetaMotorChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaMotorInputPayload>(SAMPLE_MOTOR_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaMotorEvaluationResult | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [jsonText, setJsonText] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live Calculations according to ANSI/NETA ATS-2025 Section 7.15.1
  const liveStats = useMemo(() => {
    const analysis = performDeterministicMotorAnalysis(data);
    const st = data.electrical_tests.stator_resistance_phase_to_phase_ohms;
    const vals = [st.phase_ab, st.phase_bc, st.phase_ca];
    const minVal = Math.min(...vals);
    const maxVal = Math.max(...vals);
    const devPct = minVal > 0 ? parseFloat((((maxVal - minVal) / minVal) * 100).toFixed(1)) : 0;

    const vib = data.electrical_tests.vibration_test_table_100_10;
    const vibExceeded = vib.velocity_in_sec_pk > vib.nema_grade_limit;

    const mcsa = data.electrical_tests.current_signature_analysis_mcsa;
    const brokenBarCritical = mcsa.broken_bar_sideband_db < 40.0;

    const lastVib = data.previous_test_data?.last_vibration_in_sec || 0.11;
    const vibDelta = parseFloat((vib.velocity_in_sec_pk - lastVib).toFixed(2));

    return {
      analysis,
      pi: data.electrical_tests.insulation_resistance_10min_2500v_megohms.pi_calculated,
      devPct,
      statorInvestigate: devPct > 5.0,
      vibVal: vib.velocity_in_sec_pk,
      vibLimit: vib.nema_grade_limit,
      vibExceeded,
      vibDelta,
      mcsaDb: mcsa.broken_bar_sideband_db,
      brokenBarCritical,
      isFail: analysis.overall_status === 'FAIL',
      isInvestigate: analysis.overall_status === 'INVESTIGATE'
    };
  }, [data]);

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    setAnalysisResult(null);
    setSavedSuccessMessage(null);

    try {
      const response = await fetch('/api/field-service/neta-motor-analyze', {
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
      console.error('Lỗi khi gọi API phân tích NETA Motor:', err);
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
        equipmentId: data.site_info.motor_tag,
        date: data.site_info.test_date,
        type: 'Rotating Machinery AC Motor/Generator (NETA ATS-2025 Sec 7.15.1)'
      });
      setSavedSuccessMessage(`Đã lưu biên bản kiểm định động cơ [${data.site_info.motor_tag}] vào kho báo cáo CMMS!`);
    }
  };

  const handleResetToSample = () => {
    setData(SAMPLE_MOTOR_DATA);
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
      <div className="bg-gradient-to-r from-blue-950 via-slate-900 to-indigo-950 text-white p-5 sm:p-6 rounded-2xl shadow-lg border border-blue-500/30">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1">
            <div className="flex flex-wrap items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-extrabold uppercase tracking-wider bg-white/20 text-white backdrop-blur-xs border border-white/30 flex items-center gap-1.5">
                <RotateCcw size={13} className="text-cyan-300" />
                ANSI/NETA ATS-2025 Mục 7.15.1 & Table 100.10, 100.11, 100.12
              </span>
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-blue-900/60 text-blue-200 border border-blue-400/40 flex items-center gap-1">
                <Activity size={12} />
                Rotating Machinery • AC Induction Motors & Generators • MCSA • Rung Động • PI • Độ Lệch Stator
              </span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white">
              Kiểm Tra & Thử Nghiệm Máy Điện Quay (AC Motors & Generators)
            </h1>
            <p className="text-xs sm:text-sm text-blue-100 max-w-2xl">
              Động cơ và máy phát cảm ứng xoay chiều. Tự động kiểm tra độ lệch điện trở stator (&le; 5%), chỉ số phân cực PI (&ge; 2.0), giới hạn rung động Table 100.10 và phổ dòng điện MCSA phát hiện gãy thanh dẫn rotor (&lt; 40 dB).
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
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-blue-700 hover:bg-blue-600 text-white border border-blue-400/40 flex items-center gap-1.5 transition-all shadow-md"
              title="Lưu biên bản kiểm tra vào kho báo cáo"
            >
              <Save size={15} />
              <span>Lưu vào Kho Báo Cáo</span>
            </button>

            <button
              onClick={handleAnalyze}
              disabled={isAnalyzing}
              className="px-4 py-2 rounded-xl text-xs font-bold bg-gradient-to-r from-amber-400 to-yellow-300 hover:from-amber-300 hover:to-yellow-200 text-slate-950 flex items-center gap-2 transition-all shadow-lg shadow-blue-950/40 disabled:opacity-50"
            >
              <Sparkles size={15} className={isAnalyzing ? 'animate-spin' : 'text-slate-950'} />
              <span>{isAnalyzing ? 'Đang Phân Tích NETA...' : 'Phân Tích AI Chuyên Gia'}</span>
            </button>
          </div>
        </div>

        {/* Live NETA Diagnostic Strip */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5 mt-5 pt-4 border-t border-blue-500/30 text-xs">
          {/* Insulation Resistance & PI */}
          <div className="bg-blue-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-blue-500/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-blue-200 block font-medium">Chỉ Số Phân Cực (PI)</span>
              <span className={`text-base font-black ${liveStats.pi >= 2.0 ? 'text-emerald-300' : 'text-rose-300'}`}>
                PI = {liveStats.pi} {liveStats.pi >= 2.0 ? '(≥ 2.0)' : '(< 2.0)'}
              </span>
            </div>
            {liveStats.pi >= 2.0 ? (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            ) : (
              <AlertTriangle size={20} className="text-rose-400 shrink-0" />
            )}
          </div>

          {/* Stator Resistance Deviation */}
          <div className="bg-blue-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-blue-500/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-blue-200 block font-medium">Lệch Điện Trở Pha Stator</span>
              <span className={`text-base font-black ${liveStats.statorInvestigate ? 'text-amber-300' : 'text-emerald-300'}`}>
                {liveStats.devPct}% {liveStats.statorInvestigate ? '(> 5% Lệch)' : '(≤ 5%)'}
              </span>
            </div>
            {liveStats.statorInvestigate ? (
              <AlertTriangle size={20} className="text-amber-400 shrink-0" />
            ) : (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            )}
          </div>

          {/* Mechanical Vibration */}
          <div className="bg-blue-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-blue-500/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-blue-200 block font-medium">Rung Động (Table 100.10)</span>
              <span className={`text-base font-black ${liveStats.vibExceeded ? 'text-rose-300' : 'text-emerald-300'}`}>
                {liveStats.vibVal} in./s {liveStats.vibExceeded ? '(> 0.15)' : ''}
              </span>
            </div>
            {liveStats.vibExceeded ? (
              <AlertOctagon size={20} className="text-rose-400 animate-bounce shrink-0" />
            ) : (
              <Activity size={20} className="text-emerald-400 shrink-0" />
            )}
          </div>

          {/* MCSA Broken Rotor Bar */}
          <div className="bg-blue-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-blue-500/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-blue-200 block font-medium">MCSA Gãy Thanh Lồng Sóc</span>
              <span className={`text-base font-black ${liveStats.brokenBarCritical ? 'text-rose-300 animate-pulse' : 'text-emerald-300'}`}>
                {liveStats.mcsaDb} dB {liveStats.brokenBarCritical ? '(< 40 dB Cấp bách)' : ''}
              </span>
            </div>
            {liveStats.brokenBarCritical ? (
              <AlertOctagon size={20} className="text-rose-400 shrink-0" />
            ) : (
              <Cpu size={20} className="text-blue-400 shrink-0" />
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

      {/* Main Grid: Form Inputs */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left Column: Site Info & Visual Inspection */}
        <div className="space-y-6">
          {/* Site & Motor Information */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <RotateCcw size={17} className="text-blue-600" />
                Thông Tin Động Cơ / Máy Phát
              </h2>
              <span className="text-[10px] font-semibold text-slate-500 bg-slate-100 px-2 py-0.5 rounded-full">
                NETA Mục 7.15.1
              </span>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="text-slate-600 font-medium block mb-1">Tên Trạm / Dự Án:</label>
                <input
                  type="text"
                  value={data.site_info.project_name}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, project_name: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-blue-500 font-medium text-slate-800"
                />
              </div>

              <div>
                <label className="text-slate-600 font-medium block mb-1">Mã Động Cơ (Motor Tag):</label>
                <input
                  type="text"
                  value={data.site_info.motor_tag}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, motor_tag: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-blue-500 font-bold text-blue-700 font-mono"
                />
              </div>

              <div>
                <label className="text-slate-600 font-medium block mb-1">Thông Số Định Mức (Rating):</label>
                <input
                  type="text"
                  value={data.site_info.motor_rating}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, motor_rating: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-blue-500 font-medium text-slate-800"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Kỹ Sư FSE TEV:</label>
                  <FseEngineerDropdown
                    value={data.site_info.fse_name}
                    onChange={(val) => setData({
                      ...data,
                      site_info: { ...data.site_info, fse_name: val }
                    })}
                  />
                </div>
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Ngày Thử Nghiệm:</label>
                  <input
                    type="date"
                    value={data.site_info.test_date}
                    onChange={(e) => setData({
                      ...data,
                      site_info: { ...data.site_info, test_date: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-blue-500 font-medium text-slate-800"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Visual & Mechanical Inspection (Section 7.15.1.A) */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <ShieldCheck size={17} className="text-blue-600" />
                Kiểm Tra Thị Giác & Cơ Khí (Mục 7.15.1.A)
              </h2>
            </div>

            <div className="space-y-2.5 text-xs">
              <label className="flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-200 cursor-pointer">
                <span className="font-medium text-slate-700">1. Đối soát nhãn mác (Nameplate Match 100%)</span>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.nameplate_match}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, nameplate_match: e.target.checked }
                  })}
                  className="w-4 h-4 text-blue-600 rounded focus:ring-blue-500"
                />
              </label>

              <div>
                <label className="text-slate-500 text-[11px] font-medium block mb-1">2. Tình trạng vật lý Stator/Rotor/Vách ngăn:</label>
                <input
                  type="text"
                  value={data.visual_inspection.physical_condition}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, physical_condition: e.target.value }
                  })}
                  className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded text-slate-800 text-xs"
                />
              </div>

              <div>
                <label className="text-slate-500 text-[11px] font-medium block mb-1">3. Định tâm & Khe hở không khí (Air-Gap):</label>
                <input
                  type="text"
                  value={data.visual_inspection.air_gap_and_runout_check}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, air_gap_and_runout_check: e.target.value }
                  })}
                  className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded text-slate-800 text-xs"
                />
              </div>

              <div>
                <label className="text-slate-500 text-[11px] font-medium block mb-1">4. Ổ đỡ bôi trơn & Lực siết bu-lông:</label>
                <input
                  type="text"
                  value={data.visual_inspection.lubrication_bearings}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, lubrication_bearings: e.target.value }
                  })}
                  className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded text-slate-800 text-xs"
                />
              </div>
            </div>
          </div>

          {/* Historical Trend Card */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-3 text-xs">
            <h3 className="font-bold text-slate-800 flex items-center gap-1.5 border-b border-slate-100 pb-2">
              <TrendingUp size={15} className="text-blue-600" />
              Dữ Liệu Thử Nghiệm Năm Ngoái (Trending)
            </h3>
            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="text-slate-500 text-[10px] block mb-1">Ngày Thử Trước:</label>
                <input
                  type="date"
                  value={data.previous_test_data?.last_test_date || '2025-09-20'}
                  onChange={(e) => setData({
                    ...data,
                    previous_test_data: {
                      ...data.previous_test_data,
                      last_test_date: e.target.value
                    }
                  })}
                  className="w-full px-2 py-1 bg-slate-50 border border-slate-200 rounded text-slate-700"
                />
              </div>
              <div>
                <label className="text-slate-500 text-[10px] block mb-1">Rung Động Trước (in./s):</label>
                <input
                  type="number"
                  step="0.01"
                  value={data.previous_test_data?.last_vibration_in_sec || 0.11}
                  onChange={(e) => setData({
                    ...data,
                    previous_test_data: {
                      ...data.previous_test_data,
                      last_vibration_in_sec: parseFloat(e.target.value) || 0
                    }
                  })}
                  className="w-full px-2 py-1 bg-slate-50 border border-slate-200 rounded font-bold text-slate-800"
                />
              </div>
            </div>
            <div className="p-2 rounded bg-rose-50 border border-rose-200 text-[11px] text-rose-900 font-medium">
              Biến thiên độ rung: <strong className="text-rose-700">+{liveStats.vibDelta} in./s pk</strong> (tăng gấp đôi so với kỳ trước, cảnh báo mất cân bằng rotor nghiêm trọng).
            </div>
          </div>
        </div>

        {/* Center & Right Column: Electrical Tests */}
        <div className="lg:col-span-2 space-y-6">
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-5">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div>
                <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                  <Zap size={17} className="text-blue-600" />
                  Các Phép Đo Điện Động Cơ & Đánh Giá NETA (Section 7.15.1.B - D)
                </h2>
                <span className="text-[11px] text-slate-500">
                  Tự động đối soát IEEE 43, IEEE 95, IEEE 522, IEEE 1415 và NEMA MG1
                </span>
              </div>
            </div>

            {/* Test 1: Bolted Connection Resistance */}
            <div className="p-3.5 bg-slate-50 rounded-xl border border-slate-200 space-y-2 text-xs">
              <div className="flex items-center justify-between">
                <span className="font-bold text-slate-800">1. Điện trở mối nối bu-lông (Bolted Connections):</span>
                <span className="text-[10px] text-slate-500 font-mono">NETA Table 100.12 (≤ 50% lệch)</span>
              </div>
              <div className="grid grid-cols-3 gap-2">
                {[0, 1, 2].map((idx) => (
                  <div key={idx}>
                    <label className="text-slate-500 text-[10px] block">Cực {idx + 1} (µΩ):</label>
                    <input
                      type="number"
                      step="0.1"
                      value={data.electrical_tests.bolted_resistance_micro_ohms[idx] ?? 10.5}
                      onChange={(e) => {
                        const arr = [...data.electrical_tests.bolted_resistance_micro_ohms];
                        arr[idx] = parseFloat(e.target.value) || 0;
                        setData({
                          ...data,
                          electrical_tests: { ...data.electrical_tests, bolted_resistance_micro_ohms: arr }
                        });
                      }}
                      className="w-full px-2 py-1 bg-white border border-slate-300 rounded font-semibold text-slate-800"
                    />
                  </div>
                ))}
              </div>
            </div>

            {/* Test 2: Insulation Resistance & PI */}
            <div className="p-3.5 bg-blue-50/60 rounded-xl border border-blue-200 space-y-2 text-xs">
              <div className="flex items-center justify-between">
                <span className="font-bold text-blue-900">2. Điện trở cách điện IR & Chỉ số phân cực PI (IEEE Std 43):</span>
                <span className={`text-[10px] font-black px-2 py-0.5 rounded-full ${liveStats.pi >= 2.0 ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'}`}>
                  {liveStats.pi >= 2.0 ? 'PI ĐẠT (≥ 2.0)' : 'PI KHÔNG ĐẠT'}
                </span>
              </div>
              <div className="grid grid-cols-3 gap-2">
                <div>
                  <label className="text-slate-600 text-[10px] block">IR 1 min @ 40°C (MΩ):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_10min_2500v_megohms.ir_1min_40c_corrected}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_10min_2500v_megohms: {
                          ...data.electrical_tests.insulation_resistance_10min_2500v_megohms,
                          ir_1min_40c_corrected: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-2 py-1 bg-white border border-slate-300 rounded font-bold text-slate-800"
                  />
                </div>
                <div>
                  <label className="text-slate-600 text-[10px] block">IR 10 min @ 40°C (MΩ):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_10min_2500v_megohms.ir_10min_40c_corrected}
                    onChange={(e) => {
                      const val10 = parseFloat(e.target.value) || 0;
                      const val1 = data.electrical_tests.insulation_resistance_10min_2500v_megohms.ir_1min_40c_corrected;
                      const newPi = val1 > 0 ? parseFloat((val10 / val1).toFixed(2)) : 1.0;
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_10min_2500v_megohms: {
                            ...data.electrical_tests.insulation_resistance_10min_2500v_megohms,
                            ir_10min_40c_corrected: val10,
                            pi_calculated: newPi
                          }
                        }
                      });
                    }}
                    className="w-full px-2 py-1 bg-white border border-slate-300 rounded font-bold text-slate-800"
                  />
                </div>
                <div>
                  <label className="text-slate-600 text-[10px] block">Chỉ số Phân cực (PI):</label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.insulation_resistance_10min_2500v_megohms.pi_calculated}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_10min_2500v_megohms: {
                          ...data.electrical_tests.insulation_resistance_10min_2500v_megohms,
                          pi_calculated: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-2 py-1 bg-white border border-slate-300 rounded font-black text-blue-900"
                  />
                </div>
              </div>
            </div>

            {/* Test 3: Stator Winding Resistance (Phase-to-Phase) */}
            <div className={`p-3.5 rounded-xl border space-y-2 text-xs transition-colors ${
              liveStats.statorInvestigate ? 'bg-amber-50/80 border-amber-300' : 'bg-slate-50 border-slate-200'
            }`}>
              <div className="flex items-center justify-between">
                <span className="font-bold text-slate-800">3. Điện trở cuộn dây Stator giữa các pha (Phase-to-Phase):</span>
                <span className={`text-[10px] font-black px-2 py-0.5 rounded-full ${
                  liveStats.statorInvestigate ? 'bg-amber-200 text-amber-900 animate-pulse' : 'bg-emerald-100 text-emerald-800'
                }`}>
                  Độ lệch: {liveStats.devPct}% {liveStats.statorInvestigate ? '(> 5% VƯỢT CHUẨN)' : '(≤ 5% ĐẠT)'}
                </span>
              </div>
              <div className="grid grid-cols-3 gap-2">
                <div>
                  <label className="text-slate-600 text-[10px] block">Pha A - B (Ω):</label>
                  <input
                    type="number"
                    step="0.001"
                    value={data.electrical_tests.stator_resistance_phase_to_phase_ohms.phase_ab}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        stator_resistance_phase_to_phase_ohms: {
                          ...data.electrical_tests.stator_resistance_phase_to_phase_ohms,
                          phase_ab: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-2 py-1 bg-white border border-slate-300 rounded font-mono font-bold text-slate-800"
                  />
                </div>
                <div>
                  <label className="text-slate-600 text-[10px] block">Pha B - C (Ω):</label>
                  <input
                    type="number"
                    step="0.001"
                    value={data.electrical_tests.stator_resistance_phase_to_phase_ohms.phase_bc}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        stator_resistance_phase_to_phase_ohms: {
                          ...data.electrical_tests.stator_resistance_phase_to_phase_ohms,
                          phase_bc: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-2 py-1 bg-white border border-slate-300 rounded font-mono font-bold text-slate-800"
                  />
                </div>
                <div>
                  <label className="text-slate-600 text-[10px] block">Pha C - A (Ω):</label>
                  <input
                    type="number"
                    step="0.001"
                    value={data.electrical_tests.stator_resistance_phase_to_phase_ohms.phase_ca}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        stator_resistance_phase_to_phase_ohms: {
                          ...data.electrical_tests.stator_resistance_phase_to_phase_ohms,
                          phase_ca: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-2 py-1 bg-white border border-slate-300 rounded font-mono font-bold text-slate-800"
                  />
                </div>
              </div>
            </div>

            {/* Test 4 & 5: Dielectric & Surge Comparison */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 space-y-1.5">
                <label className="font-bold text-slate-800 block">4. Thử chịu áp Dielectric (IEEE 95 DC):</label>
                <input
                  type="text"
                  value={data.electrical_tests.dielectric_withstand_test_ieee95}
                  onChange={(e) => setData({
                    ...data,
                    electrical_tests: { ...data.electrical_tests, dielectric_withstand_test_ieee95: e.target.value }
                  })}
                  className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded text-slate-800"
                />
              </div>

              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 space-y-1.5">
                <label className="font-bold text-slate-800 block">5. Thử xung Surge Comparison (IEEE 522):</label>
                <input
                  type="text"
                  value={data.electrical_tests.surge_comparison_test}
                  onChange={(e) => setData({
                    ...data,
                    electrical_tests: { ...data.electrical_tests, surge_comparison_test: e.target.value }
                  })}
                  className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded text-slate-800"
                />
              </div>
            </div>

            {/* Test 6 & 7: Mechanical Vibration & MCSA */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
              {/* Vibration Table 100.10 */}
              <div className={`p-3.5 rounded-xl border space-y-2 ${
                liveStats.vibExceeded ? 'bg-rose-50 border-rose-300' : 'bg-slate-50 border-slate-200'
              }`}>
                <div className="flex items-center justify-between">
                  <span className="font-bold text-slate-800">6. Rung động (Table 100.10):</span>
                  <span className={`text-[10px] font-black px-2 py-0.5 rounded-full ${
                    liveStats.vibExceeded ? 'bg-rose-200 text-rose-900 animate-pulse' : 'bg-emerald-100 text-emerald-800'
                  }`}>
                    {liveStats.vibExceeded ? 'FAIL (> Giới hạn)' : 'PASS'}
                  </span>
                </div>
                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="text-slate-500 text-[10px] block">Vận tốc đo (in./s pk):</label>
                    <input
                      type="number"
                      step="0.01"
                      value={data.electrical_tests.vibration_test_table_100_10.velocity_in_sec_pk}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          vibration_test_table_100_10: {
                            ...data.electrical_tests.vibration_test_table_100_10,
                            velocity_in_sec_pk: parseFloat(e.target.value) || 0
                          }
                        }
                      })}
                      className="w-full px-2 py-1 bg-white border border-slate-300 rounded font-bold text-slate-800"
                    />
                  </div>
                  <div>
                    <label className="text-slate-500 text-[10px] block">Giới hạn NEMA Grade A:</label>
                    <input
                      type="number"
                      step="0.01"
                      value={data.electrical_tests.vibration_test_table_100_10.nema_grade_limit}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          vibration_test_table_100_10: {
                            ...data.electrical_tests.vibration_test_table_100_10,
                            nema_grade_limit: parseFloat(e.target.value) || 0
                          }
                        }
                      })}
                      className="w-full px-2 py-1 bg-white border border-slate-300 rounded font-bold text-slate-800"
                    />
                  </div>
                </div>
              </div>

              {/* MCSA Broken Rotor Bar */}
              <div className={`p-3.5 rounded-xl border space-y-2 ${
                liveStats.brokenBarCritical ? 'bg-rose-50 border-rose-300 ring-1 ring-rose-300' : 'bg-slate-50 border-slate-200'
              }`}>
                <div className="flex items-center justify-between">
                  <span className="font-bold text-slate-800">7. MCSA Gãy Thanh Rotor:</span>
                  <span className={`text-[10px] font-black px-2 py-0.5 rounded-full ${
                    liveStats.brokenBarCritical ? 'bg-rose-600 text-white animate-pulse' : 'bg-emerald-100 text-emerald-800'
                  }`}>
                    {liveStats.brokenBarCritical ? 'CRITICAL DEFECT' : 'PASS'}
                  </span>
                </div>
                <div>
                  <label className="text-slate-500 text-[10px] block">Dải phổ phụ tần số gãy thanh (dB):</label>
                  <div className="flex items-center gap-1.5 mt-0.5">
                    <input
                      type="number"
                      step="0.1"
                      value={data.electrical_tests.current_signature_analysis_mcsa.broken_bar_sideband_db}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          current_signature_analysis_mcsa: {
                            broken_bar_sideband_db: parseFloat(e.target.value) || 0
                          }
                        }
                      })}
                      className="w-full px-2 py-1 bg-white border border-slate-300 rounded font-black text-rose-700"
                    />
                    <span className="text-[11px] text-slate-500 font-bold shrink-0">dB</span>
                  </div>
                  <span className="text-[10px] text-slate-500 block mt-1">
                    Ngưỡng NETA: &lt; 40 dB (Gãy thanh lồng sóc) • 40-60 dB (Cảnh báo) • &ge; 60 dB (Tốt)
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* AI Analysis Output Section */}
      {analysisResult && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden animate-in fade-in slide-in-from-bottom-3 duration-300">
          <div className="p-5 bg-gradient-to-r from-slate-900 via-blue-950 to-slate-900 text-white flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-blue-500/20 border border-blue-400/30 flex items-center justify-center text-blue-300 shrink-0">
                <Sparkles size={20} />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className="text-[10px] uppercase tracking-wider font-extrabold px-2 py-0.5 rounded-full bg-white/10 text-white">
                    TEV Platform AI Evaluation
                  </span>
                  <span className={`text-[11px] font-black px-2 py-0.5 rounded-full border ${
                    analysisResult.overall_status === 'FAIL'
                      ? 'bg-rose-500/30 text-rose-300 border-rose-400/50'
                      : analysisResult.overall_status === 'INVESTIGATE'
                      ? 'bg-amber-500/30 text-amber-300 border-amber-400/50'
                      : 'bg-emerald-500/30 text-emerald-300 border-emerald-400/50'
                  }`}>
                    {analysisResult.overall_status}
                  </span>
                </div>
                <h3 className="text-base font-bold text-white mt-0.5">
                  Báo Cáo Kiểm Định Động Cơ / Máy Phát (NETA ATS-2025 Section 7.15.1)
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
                className="px-4 py-1.5 bg-blue-700 hover:bg-blue-600 text-white text-xs font-bold rounded-lg flex items-center gap-1.5 shadow-md transition-colors"
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
                <FileCheck2 size={18} className="text-blue-600" />
                <h3 className="font-bold text-sm text-slate-800">
                  Dữ Liệu JSON Kiểm Tra Động Cơ (NETA Sec 7.15.1)
                </h3>
              </div>
              <button
                onClick={() => setShowJsonModal(false)}
                className="text-slate-400 hover:text-slate-600 text-sm font-bold"
              >
                ✕
              </button>
            </div>

            <textarea
              value={jsonText}
              onChange={(e) => setJsonText(e.target.value)}
              rows={14}
              className="w-full p-3 font-mono text-xs bg-slate-900 text-blue-200 rounded-xl border border-slate-700 focus:outline-hidden focus:ring-2 focus:ring-blue-500"
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
                  className="px-4 py-1.5 bg-blue-600 hover:bg-blue-500 text-white text-xs font-bold rounded-lg shadow-sm"
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
          title="BIÊN BẢN KIỂM ĐỊNH MÁY ĐIỆN QUAY (AC INDUCTION MOTORS & GENERATORS)"
          standard="ANSI/NETA ATS-2025 Section 7.15.1 & Tables 100.10, 100.11, 100.12"
          siteInfo={{
            project_name: data.site_info.project_name,
            transformer_tag: data.site_info.motor_tag,
            fse_name: data.site_info.fse_name,
            test_date: data.site_info.test_date
          }}
          overallStatus={analysisResult?.overall_status === 'FAIL' ? 'CRITICAL' : analysisResult?.overall_status === 'INVESTIGATE' ? 'INVESTIGATE' : 'PASS'}
          statusReason={analysisResult?.status_reason || 'Đã đối soát với tiêu chuẩn ANSI/NETA ATS-2025 Section 7.15.1.'}
          visualChecks={[
            { label: 'Đối soát nhãn mác kỹ thuật', value: data.visual_inspection.nameplate_match ? 'Khớp 100% bản vẽ' : 'Lệch', pass: data.visual_inspection.nameplate_match },
            { label: 'Tình trạng vật lý Stator/Rotor/Vách ngăn', value: data.visual_inspection.physical_condition, pass: true },
            { label: 'Định tâm & Khe hở không khí (Air-Gap)', value: data.visual_inspection.air_gap_and_runout_check, pass: true },
            { label: 'Bôi trơn ổ đỡ & Bu-lông siết lực', value: data.visual_inspection.lubrication_bearings, pass: true },
            { label: 'Mạch đo nhiệt RTD & Sấy chống ẩm', value: data.visual_inspection.rtd_circuits_check, pass: true }
          ]}
          electricalTable={liveStats.analysis.evaluations.items.map(item => ({
            item: item.item,
            measured: item.measured,
            standard: item.standard,
            reference: item.reference,
            deviation: item.detail,
            status: item.status
          }))}
          warnings={[
            liveStats.brokenBarCritical ? 'CẢNH BÁO TỐI CẤP: Dải phổ dòng điện MCSA đạt 35.2 dB (< 40 dB). Xác nhận lồng sóc của Rotor đã bị nứt/gãy thanh dẫn squirrel-cage bar! DỪNG VẬN HÀNH NGAY.' : '',
            liveStats.vibExceeded ? `Rung động cơ học ${liveStats.vibVal} in./s pk vượt quá ngưỡng tối đa ${liveStats.vibLimit} in./s pk (Table 100.10 NEMA Grade A).` : '',
            liveStats.statorInvestigate ? `Độ lệch điện trở cuộn dây stator pha CA đạt ${liveStats.devPct}% vượt ngưỡng 5% của NETA ATS-2025 Section 7.15.1.D.4.` : ''
          ].filter(Boolean)}
          recommendations={[
            'DỪNG VẬN HÀNH ĐỘNG CƠ CẤP BÁCH để ngăn ngừa ứng suất lực ly tâm làm bung thanh dẫn phá hủy Stator.',
            'Rút rotor kiểm tra thẩm thấu hạt từ tính (MPI) tìm vết nứt thanh dẫn lồng sóc và vòng ngắn mạch end-rings.',
            'Vệ sinh và siết lại đầu cốt bu-lông pha CA tại hộp cực động cơ theo Bảng Table 100.12.',
            'Cân bằng động lại rotor và đại tu thay thế vòng bi ổ đỡ để hạ độ rung về dưới 0.15 in./s pk.'
          ]}
        />
      )}
    </div>
  );
};
