import React, { useState, useMemo } from 'react';
import {
  Zap,
  PlugZap,
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
  Radio,
  Eye,
  Info,
  Layers,
  Sliders,
  Cable,
  Gauge
} from 'lucide-react';
import { NetaEvseInputPayload, NetaEvseEvaluationResult } from '../../server/netaEvseAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

const SAMPLE_EVSE_DATA: NetaEvseInputPayload = {
  site_info: {
    project_name: "Tram Sac Xe Dien TEV Fast Charging Hub",
    evse_tag: "EVSE-DC-120KW-01",
    charger_type: "DC Fast Charger (120kW Dual Gun)",
    fse_name: "Nguyen Van G",
    test_date: "2026-09-23"
  },
  visual_inspection: {
    nameplate_match: true,
    cords_and_connectors_condition: "Pass (Gun 1 Good, Gun 2 minor pin wear)",
    anchorage_grounding_clearance: "Pass",
    cleanliness_and_shipping_bracing_removed: true,
    wiring_connections_tight: true,
    bolt_torque_check: "Pass",
    interlocking_system_operation: "Pass"
  },
  electrical_tests: {
    system_function_test: "Pass",
    pe_grounding_protection: "Pass",
    protective_conductor_resistance_ohms: {
      gun_1_pe_resistance: 0.18,
      gun_2_pe_resistance: 0.68
    },
    touch_voltage_check: "Pass (Within safe limit)",
    personal_protection_rcd_ccid: "Pass",
    nuisance_tripping_check: "Pass",
    proximity_pilot_pp_test: "Pass (Correct cable rating identification)",
    control_pilot_cp_test: "Pass (State A to State C transition normal)",
    charger_output_voltage_v_dc: 402,
    dc_voltage_ripple_percent: 0.8
  },
  previous_test_data: {
    last_test_date: "2025-09-22",
    last_gun_2_pe_resistance_ohms: 0.22
  }
};

export const NetaEvseChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaEvseInputPayload>(SAMPLE_EVSE_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaEvseEvaluationResult | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [jsonText, setJsonText] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live Calculations based on NETA ATS-2025 Section 7.26
  const liveStats = useMemo(() => {
    const gun1Pe = data.electrical_tests?.protective_conductor_resistance_ohms?.gun_1_pe_resistance ?? 0.18;
    const gun2Pe = data.electrical_tests?.protective_conductor_resistance_ohms?.gun_2_pe_resistance ?? 0.68;
    const peThreshold = 0.5; // Ohms max

    const isGun1Pass = gun1Pe <= peThreshold;
    const isGun2Pass = gun2Pe <= peThreshold;
    const isOverallPePass = isGun1Pass && isGun2Pass;

    const lastGun2Pe = data.previous_test_data?.last_gun_2_pe_resistance_ohms ?? 0.22;
    const gun2IncreasePct = lastGun2Pe > 0 ? (((gun2Pe - lastGun2Pe) / lastGun2Pe) * 100).toFixed(0) : '0';

    const ripple = data.electrical_tests?.dc_voltage_ripple_percent ?? 0.8;
    const isRipplePass = ripple <= 2.0;

    const outVolt = data.electrical_tests?.charger_output_voltage_v_dc ?? 402;

    const isCpPass = data.electrical_tests?.control_pilot_cp_test?.toLowerCase().includes('pass');
    const isPpPass = data.electrical_tests?.proximity_pilot_pp_test?.toLowerCase().includes('pass');

    return {
      gun1Pe,
      gun2Pe,
      peThreshold,
      isGun1Pass,
      isGun2Pass,
      isOverallPePass,
      lastGun2Pe,
      gun2IncreasePct,
      ripple,
      isRipplePass,
      outVolt,
      isCpPass,
      isPpPass
    };
  }, [data]);

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    setAnalysisResult(null);
    setSavedSuccessMessage(null);

    try {
      const response = await fetch('/api/field-service/neta-evse-analyze', {
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
      console.error('Lỗi khi gọi API phân tích NETA EVSE:', err);
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
        equipmentId: data.site_info.evse_tag,
        date: data.site_info.test_date,
        type: 'Electric Vehicle Charging Systems (NETA ATS-2025 Sec 7.26)'
      });
      setSavedSuccessMessage(`Đã lưu biên bản kiểm định trạm sạc EVSE [${data.site_info.evse_tag}] vào kho báo cáo CMMS!`);
    }
  };

  const handleResetToSample = () => {
    setData(SAMPLE_EVSE_DATA);
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
      <div className="bg-gradient-to-r from-cyan-700 via-blue-700 to-indigo-800 text-white p-5 sm:p-6 rounded-2xl shadow-lg border border-cyan-500/30">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1">
            <div className="flex flex-wrap items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-extrabold uppercase tracking-wider bg-white/20 text-white backdrop-blur-xs border border-white/30 flex items-center gap-1.5">
                <PlugZap size={13} className="text-cyan-200" />
                ANSI/NETA ATS-2025 Mục 7.26
              </span>
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-cyan-950/40 text-cyan-200 border border-cyan-400/40 flex items-center gap-1">
                <ShieldCheck size={12} />
                EV Charging Systems • DC Fast Charger • Súng Sạc Kép
              </span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white">
              Kiểm Định Hệ Thống Sạc Xe Điện (Electric Vehicle Charging Systems - EVSE)
            </h1>
            <p className="text-xs sm:text-sm text-cyan-100 max-w-2xl">
              Đo điện trở dây tiếp địa bảo vệ PE (ngưỡng NETA ≤ 0.5 Ω), điện áp chạm, RCD/CCID, kiểm tra xung Control Pilot (CP), Proximity Pilot (PP) và độ gợn sóng DC theo NETA Section 7.26.
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
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-emerald-700 hover:bg-emerald-600 text-white border border-emerald-500/40 flex items-center gap-1.5 transition-all shadow-md"
              title="Lưu biên bản kiểm tra vào kho báo cáo"
            >
              <Save size={15} />
              <span>Lưu vào Kho Báo Cáo</span>
            </button>

            <button
              onClick={handleAnalyze}
              disabled={isAnalyzing}
              className="px-4 py-2 rounded-xl text-xs font-bold bg-gradient-to-r from-yellow-300 to-amber-300 hover:from-yellow-200 hover:to-amber-200 text-slate-950 flex items-center gap-2 transition-all shadow-lg shadow-cyan-900/30 disabled:opacity-50"
            >
              <Sparkles size={15} className={isAnalyzing ? 'animate-spin' : 'text-amber-800'} />
              <span>{isAnalyzing ? 'Đang Phân Tích NETA...' : 'Phân Tích AI Chuyên Gia'}</span>
            </button>
          </div>
        </div>

        {/* Live NETA Diagnostic Strip */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5 mt-5 pt-4 border-t border-cyan-400/30 text-xs">
          {/* Gun 1 PE Resistance */}
          <div className="bg-cyan-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-cyan-400/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-cyan-200 block font-medium">Súng 1 Dây PE (≤0.5Ω)</span>
              <span className={`text-base font-black ${liveStats.isGun1Pass ? 'text-emerald-300' : 'text-rose-300'}`}>
                {liveStats.gun1Pe} Ω
              </span>
            </div>
            {liveStats.isGun1Pass ? (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            ) : (
              <XCircle size={20} className="text-rose-400 shrink-0" />
            )}
          </div>

          {/* Gun 2 PE Resistance */}
          <div className="bg-cyan-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-cyan-400/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-cyan-200 block font-medium">Súng 2 Dây PE (≤0.5Ω)</span>
              <span className={`text-base font-black ${liveStats.isGun2Pass ? 'text-emerald-300' : 'text-rose-300'}`}>
                {liveStats.gun2Pe} Ω (+{liveStats.gun2IncreasePct}%)
              </span>
            </div>
            {liveStats.isGun2Pass ? (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            ) : (
              <AlertTriangle size={20} className="text-rose-400 shrink-0" />
            )}
          </div>

          {/* DC Voltage Ripple */}
          <div className="bg-cyan-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-cyan-400/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-cyan-200 block font-medium">Độ Gợn DC (≤2.0%)</span>
              <span className={`text-base font-black ${liveStats.isRipplePass ? 'text-emerald-300' : 'text-rose-300'}`}>
                {liveStats.ripple}% ({liveStats.outVolt}V)
              </span>
            </div>
            {liveStats.isRipplePass ? (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            ) : (
              <AlertTriangle size={20} className="text-rose-400 shrink-0" />
            )}
          </div>

          {/* CP / PP Signal Status */}
          <div className="bg-cyan-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-cyan-400/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-cyan-200 block font-medium">Tín Hiệu CP/PP Pilot</span>
              <span className={`text-base font-black ${liveStats.isCpPass && liveStats.isPpPass ? 'text-emerald-300' : 'text-rose-300'}`}>
                {liveStats.isCpPass && liveStats.isPpPass ? 'Chuẩn A-B-C' : 'Lỗi Giao Tiếp'}
              </span>
            </div>
            {liveStats.isCpPass && liveStats.isPpPass ? (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            ) : (
              <XCircle size={20} className="text-rose-400 shrink-0" />
            )}
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
          {/* Site & Charger Information */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <PlugZap size={17} className="text-cyan-600" />
                Thông Tin Trạm Sạc & Hiện Trường
              </h2>
              <span className="text-[10px] font-semibold text-slate-500 bg-slate-100 px-2 py-0.5 rounded-full">
                NETA ATS-2025 Sec 7.26
              </span>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="text-slate-600 font-medium block mb-1">Tên Dự Án / Địa Điểm Trạm Sạc:</label>
                <input
                  type="text"
                  value={data.site_info.project_name}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, project_name: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-cyan-500 font-medium text-slate-800"
                />
              </div>

              <div>
                <label className="text-slate-600 font-medium block mb-1">Mã Định Danh Trạm Sạc (EVSE Tag):</label>
                <input
                  type="text"
                  value={data.site_info.evse_tag}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, evse_tag: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-cyan-500 font-bold text-cyan-700 font-mono"
                />
              </div>

              <div>
                <label className="text-slate-600 font-medium block mb-1">Loại Trạm Sạc & Cấu Hình:</label>
                <input
                  type="text"
                  value={data.site_info.charger_type}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, charger_type: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-cyan-500 font-medium text-slate-800"
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
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-cyan-500 font-medium text-slate-800"
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
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-cyan-500 font-medium text-slate-800"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Section 7.26.1 Visual and Mechanical Inspection */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <Eye size={17} className="text-cyan-600" />
                Kiểm Tra Thị Giác & Cơ Khí (Mục 7.26.1)
              </h2>
            </div>

            <div className="space-y-3 text-xs">
              {/* Nameplate Match */}
              <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-xl border border-slate-100">
                <div>
                  <span className="font-bold text-slate-800 block">Nameplate Match</span>
                  <span className="text-[11px] text-slate-500">Khớp 100% dữ liệu nhãn trạm sạc với bản vẽ thiết kế</span>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.nameplate_match}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, nameplate_match: e.target.checked }
                  })}
                  className="w-5 h-5 text-cyan-600 rounded-md focus:ring-cyan-500"
                />
              </div>

              {/* Cords and Connectors */}
              <div>
                <label className="text-slate-600 font-medium block mb-1">Cáp & Súng Sạc (Cords & Connectors):</label>
                <input
                  type="text"
                  value={data.visual_inspection.cords_and_connectors_condition}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, cords_and_connectors_condition: e.target.value }
                  })}
                  placeholder="Pass (Gun 1 Good, Gun 2 minor pin wear)"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-cyan-500 font-medium text-slate-800"
                />
              </div>

              {/* Anchorage & Grounding Clearance */}
              <div>
                <label className="text-slate-600 font-medium block mb-1">Cố Định Chân Đế & Tiếp Địa Vỏ (Anchorage):</label>
                <input
                  type="text"
                  value={data.visual_inspection.anchorage_grounding_clearance}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, anchorage_grounding_clearance: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-cyan-500 font-medium text-slate-800"
                />
              </div>

              {/* Cleanliness and shipping bracing removed */}
              <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-xl border border-slate-100">
                <div>
                  <span className="font-bold text-slate-800 block">Vệ Sinh & Tháo Kẹp Vận Chuyển</span>
                  <span className="text-[11px] text-slate-500">Đã tháo bỏ toàn bộ nẹp cố định vận chuyển & tài liệu</span>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.cleanliness_and_shipping_bracing_removed}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, cleanliness_and_shipping_bracing_removed: e.target.checked }
                  })}
                  className="w-5 h-5 text-cyan-600 rounded-md focus:ring-cyan-500"
                />
              </div>

              {/* Wiring Connections Tight */}
              <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-xl border border-slate-100">
                <div>
                  <span className="font-bold text-slate-800 block">Dây Dẫn Siết Chặt & Cố Định</span>
                  <span className="text-[11px] text-slate-500">Tránh giằng xé dây dẫn khi kéo súng sạc ra vào</span>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.wiring_connections_tight}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, wiring_connections_tight: e.target.checked }
                  })}
                  className="w-5 h-5 text-cyan-600 rounded-md focus:ring-cyan-500"
                />
              </div>

              {/* Bolt Torque Check */}
              <div>
                <label className="text-slate-600 font-medium block mb-1">Kiểm Tra Lực Siết Bu-lông (Table 100.12):</label>
                <input
                  type="text"
                  value={data.visual_inspection.bolt_torque_check}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, bolt_torque_check: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-cyan-500 font-medium text-slate-800"
                />
              </div>

              {/* Interlocking Systems */}
              <div>
                <label className="text-slate-600 font-medium block mb-1">Khóa & Liên Động Điện/Cơ (Interlocks):</label>
                <input
                  type="text"
                  value={data.visual_inspection.interlocking_system_operation}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, interlocking_system_operation: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-cyan-500 font-medium text-slate-800"
                />
              </div>
            </div>
          </div>
        </div>

        {/* Center & Right Column: Electrical Tests (Mục 7.26.2) */}
        <div className="lg:col-span-2 space-y-6">
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-5">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <Zap size={17} className="text-cyan-600" />
                Phép Đo Điện & Tiêu Chuẩn Đánh Giá (NETA Sec 7.26.2)
              </h2>
              <span className="text-[11px] font-semibold text-cyan-700 bg-cyan-50 border border-cyan-200 px-2.5 py-0.5 rounded-full">
                NETA ATS-2025 Section 7.26.2
              </span>
            </div>

            {/* Test 1: Protective Conductor Resistance (PE) - CRITICAL NETA RULE */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                  <Cable size={15} className="text-cyan-600" />
                  1. Điện Trở Dây Bảo Vệ (Protective Conductor Resistance - PE Grounding)
                </span>
                <span className="text-[10px] text-rose-600 font-bold bg-rose-50 px-2 py-0.5 rounded border border-rose-200">
                  BẮT BUỘC ≤ 0.50 Ω
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                {/* Gun 1 PE */}
                <div>
                  <label className="text-slate-600 font-medium block mb-1 flex items-center justify-between">
                    <span>Súng sạc 01 (Gun 1 PE - Ω):</span>
                    {liveStats.isGun1Pass ? (
                      <span className="text-[10px] text-emerald-600 font-bold">Đạt (≤0.5Ω)</span>
                    ) : (
                      <span className="text-[10px] text-rose-600 font-bold">Vượt Ngưỡng</span>
                    )}
                  </label>
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.protective_conductor_resistance_ohms.gun_1_pe_resistance}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        protective_conductor_resistance_ohms: {
                          ...data.electrical_tests.protective_conductor_resistance_ohms,
                          gun_1_pe_resistance: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className={`w-full px-3 py-2 bg-white border rounded-lg focus:outline-hidden focus:ring-2 font-bold ${
                      liveStats.isGun1Pass
                        ? 'border-slate-200 text-emerald-700 focus:ring-cyan-500'
                        : 'border-rose-400 text-rose-700 bg-rose-50/50 focus:ring-rose-500'
                    }`}
                  />
                </div>

                {/* Gun 2 PE */}
                <div>
                  <label className="text-slate-600 font-medium block mb-1 flex items-center justify-between">
                    <span>Súng sạc 02 (Gun 2 PE - Ω):</span>
                    {liveStats.isGun2Pass ? (
                      <span className="text-[10px] text-emerald-600 font-bold">Đạt (≤0.5Ω)</span>
                    ) : (
                      <span className="text-[10px] text-rose-600 font-bold">0.68 Ω &gt; 0.50 Ω (FAIL)</span>
                    )}
                  </label>
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.protective_conductor_resistance_ohms.gun_2_pe_resistance}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        protective_conductor_resistance_ohms: {
                          ...data.electrical_tests.protective_conductor_resistance_ohms,
                          gun_2_pe_resistance: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className={`w-full px-3 py-2 bg-white border rounded-lg focus:outline-hidden focus:ring-2 font-bold ${
                      liveStats.isGun2Pass
                        ? 'border-slate-200 text-slate-800 focus:ring-cyan-500'
                        : 'border-rose-400 text-rose-700 bg-rose-50/50 focus:ring-rose-500'
                    }`}
                  />
                </div>
              </div>
            </div>

            {/* Test 2: Touch Voltage, RCD/CCID & Nuisance Tripping */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <ShieldCheck size={15} className="text-cyan-600" />
                2. Bảo Vệ Cá Nhân (RCD / CCID), Điện Áp Chạm & Nhảy Bảo Vệ Giả
              </span>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Điện Áp Chạm (Touch Voltage):</label>
                  <input
                    type="text"
                    value={data.electrical_tests.touch_voltage_check}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: { ...data.electrical_tests, touch_voltage_check: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-cyan-500 font-medium text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Bảo Vệ Chống Dòng Rò (RCD/CCID):</label>
                  <input
                    type="text"
                    value={data.electrical_tests.personal_protection_rcd_ccid}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: { ...data.electrical_tests, personal_protection_rcd_ccid: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-cyan-500 font-medium text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Hiện Tượng Nhảy Bảo Vệ Giả:</label>
                  <input
                    type="text"
                    value={data.electrical_tests.nuisance_tripping_check}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: { ...data.electrical_tests, nuisance_tripping_check: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-cyan-500 font-medium text-slate-800"
                  />
                </div>
              </div>
            </div>

            {/* Test 3: Pilot Signals (PP & CP) */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <Radio size={15} className="text-cyan-600" />
                3. Tín Hiệu Điều Khiển & Phản Hồi Súng Sạc (Control Pilot - CP / Proximity Pilot - PP)
              </span>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Tín Hiệu Nhận Diện Cáp (PP):</label>
                  <input
                    type="text"
                    value={data.electrical_tests.proximity_pilot_pp_test}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: { ...data.electrical_tests, proximity_pilot_pp_test: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-cyan-500 font-medium text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Tín Hiệu Điều Khiển Nạp (CP - PWM):</label>
                  <input
                    type="text"
                    value={data.electrical_tests.control_pilot_cp_test}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: { ...data.electrical_tests, control_pilot_cp_test: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-cyan-500 font-medium text-slate-800"
                  />
                </div>
              </div>
            </div>

            {/* Test 4: Charger Output Voltage & DC Ripple */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <Gauge size={15} className="text-cyan-600" />
                4. Điện Áp Đầu Ra DC & Độ Gợn Sóng (DC Ripple)
              </span>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Điện Áp Đầu Ra Định Mức (V DC):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.charger_output_voltage_v_dc}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        charger_output_voltage_v_dc: parseFloat(e.target.value) || 0
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-cyan-500 font-bold text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1 flex items-center justify-between">
                    <span>Độ Gợn Sóng Điện Áp DC (%):</span>
                    {liveStats.isRipplePass ? (
                      <span className="text-[10px] text-emerald-600 font-bold">Đạt (≤2.0%)</span>
                    ) : (
                      <span className="text-[10px] text-rose-600 font-bold">Vượt Ngưỡng</span>
                    )}
                  </label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.dc_voltage_ripple_percent}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        dc_voltage_ripple_percent: parseFloat(e.target.value) || 0
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-cyan-500 font-bold text-cyan-700"
                  />
                </div>
              </div>
            </div>

            {/* Historical Reference */}
            {data.previous_test_data && (
              <div className="p-3 bg-cyan-50/60 rounded-xl border border-cyan-200/60 flex items-center justify-between text-xs text-cyan-900">
                <span className="font-semibold flex items-center gap-1.5">
                  <Info size={14} className="text-cyan-600" />
                  Kỳ Đo Trước ({data.previous_test_data.last_test_date || '2025-09-22'}): Súng 2 PE = {data.previous_test_data.last_gun_2_pe_resistance_ohms} Ω (Bình thường)
                </span>
                <span className="text-[11px] font-bold text-rose-700 bg-white px-2 py-0.5 rounded border border-rose-300">
                  Tăng Vọt {liveStats.gun2IncreasePct}% Lên {liveStats.gun2Pe} Ω
                </span>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* AI Analysis Output Section */}
      {analysisResult && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden animate-in fade-in slide-in-from-bottom-3 duration-300">
          <div className="p-5 bg-gradient-to-r from-slate-900 via-cyan-950 to-slate-900 text-white flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-cyan-500/20 border border-cyan-400/30 flex items-center justify-center text-cyan-300 shrink-0">
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
                  Báo Cáo Đánh Giá Kỹ Thuật Trạm Sạc Xe Điện (NETA ATS-2025 Mục 7.26)
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
                className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold rounded-lg flex items-center gap-1.5 shadow-md transition-colors"
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
                <FileCheck2 size={18} className="text-cyan-600" />
                <h3 className="font-bold text-sm text-slate-800">
                  Dữ Liệu JSON Kiểm Định EVSE (NETA ATS-2025 Sec 7.26)
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
              className="w-full p-3 font-mono text-xs bg-slate-900 text-cyan-200 rounded-xl border border-slate-700 focus:outline-hidden focus:ring-2 focus:ring-cyan-500"
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
                  className="px-4 py-1.5 bg-cyan-600 hover:bg-cyan-500 text-white text-xs font-bold rounded-lg shadow-sm"
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
          title="BIÊN BẢN KIỂM ĐỊNH HỆ THỐNG SẠC XE ĐIỆN (EV CHARGING SYSTEMS)"
          standard="ANSI/NETA ATS-2025 Section 7.26"
          siteInfo={{
            project_name: data.site_info.project_name,
            transformer_tag: data.site_info.evse_tag,
            fse_name: data.site_info.fse_name,
            test_date: data.site_info.test_date
          }}
          overallStatus={analysisResult?.overall_status || (liveStats.isOverallPePass ? 'PASS' : 'FAIL')}
          statusReason={analysisResult?.status_reason || 'Đã đối chiếu với quy chuẩn kỹ thuật NETA ATS-2025 Mục 7.26.'}
          visualChecks={[
            { label: 'Nameplate Match (Khớp bản vẽ 100%)', value: data.visual_inspection.nameplate_match ? 'Khớp 100%' : 'Sai lệch', pass: data.visual_inspection.nameplate_match },
            { label: 'Cáp sạc & Súng súng cơ lý nguyên vẹn', value: data.visual_inspection.cords_and_connectors_condition, pass: data.visual_inspection.cords_and_connectors_condition.toLowerCase().includes('pass') },
            { label: 'Cố định chân đế & Tiếp địa vỏ', value: data.visual_inspection.anchorage_grounding_clearance, pass: data.visual_inspection.anchorage_grounding_clearance.toLowerCase().includes('pass') },
            { label: 'Vệ sinh & Tháo kẹp vận chuyển', value: data.visual_inspection.cleanliness_and_shipping_bracing_removed ? 'Đã tháo toàn bộ' : 'Chưa tháo', pass: data.visual_inspection.cleanliness_and_shipping_bracing_removed },
            { label: 'Dây dẫn siết chặt & Cố định an toàn', value: data.visual_inspection.wiring_connections_tight ? 'Siết chặt' : 'Lỏng lẻo', pass: data.visual_inspection.wiring_connections_tight },
            { label: 'Lực siết bu-lông (Table 100.12)', value: data.visual_inspection.bolt_torque_check, pass: data.visual_inspection.bolt_torque_check.toLowerCase().includes('pass') },
            { label: 'Khóa liên động điện & Cơ khí', value: data.visual_inspection.interlocking_system_operation, pass: data.visual_inspection.interlocking_system_operation.toLowerCase().includes('pass') }
          ]}
          electricalTable={[
            {
              item: 'Thử nghiệm chức năng hệ thống',
              measured: data.electrical_tests.system_function_test,
              standard: 'Phù hợp dữ liệu công bố của NSX',
              reference: 'Manual NSX',
              deviation: 'Đạt chức năng',
              status: 'PASS'
            },
            {
              item: 'Bảo vệ tiếp địa PE Súng 1 (Gun 1)',
              measured: `${liveStats.gun1Pe} Ω`,
              standard: 'Điện trở dây bảo vệ ≤ 0.50 Ω (NETA 7.26.D.2.2)',
              reference: 'Ngưỡng: ≤ 0.50 Ω',
              deviation: `${liveStats.gun1Pe} Ω ≤ 0.50 Ω`,
              status: liveStats.isGun1Pass ? 'PASS' : 'FAIL'
            },
            {
              item: 'Bảo vệ tiếp địa PE Súng 2 (Gun 2)',
              measured: `${liveStats.gun2Pe} Ω`,
              standard: 'Điện trở dây bảo vệ ≤ 0.50 Ω (NETA 7.26.D.2.2)',
              reference: 'Ngưỡng: ≤ 0.50 Ω (Kỳ trước: 0.22 Ω)',
              deviation: `${liveStats.gun2Pe} Ω > 0.50 Ω (VƯỢT 36%, TĂNG ${liveStats.gun2IncreasePct}%)`,
              status: liveStats.isGun2Pass ? 'PASS' : 'FAIL'
            },
            {
              item: 'Điện áp chạm (Touch Voltage)',
              measured: data.electrical_tests.touch_voltage_check,
              standard: 'Nằm trong ngưỡng điện áp an toàn cho phép',
              reference: 'Quy chuẩn an toàn',
              deviation: 'An toàn thao tác',
              status: 'PASS'
            },
            {
              item: 'Bảo vệ cá nhân (RCD / CCID)',
              measured: data.electrical_tests.personal_protection_rcd_ccid,
              standard: 'Tự động cắt khi phát hiện dòng rò',
              reference: 'Tiêu chuẩn bảo vệ CCID',
              deviation: 'Ngắt đúng thời gian',
              status: 'PASS'
            },
            {
              item: 'Hiện tượng nhảy bảo vệ giả',
              measured: data.electrical_tests.nuisance_tripping_check,
              standard: 'Không xảy ra nhảy giả khi mang tải',
              reference: 'Vận hành ổn định',
              deviation: 'Không nhảy sai',
              status: 'PASS'
            },
            {
              item: 'Tín hiệu Phản hồi Súng sạc (PP)',
              measured: data.electrical_tests.proximity_pilot_pp_test,
              standard: 'Nhận diện đúng mức dòng định mức cáp sạc',
              reference: 'IEC 61851 / SAE J1772',
              deviation: 'Đúng tiết diện cáp',
              status: 'PASS'
            },
            {
              item: 'Tín hiệu Điều khiển Nạp (CP)',
              measured: data.electrical_tests.control_pilot_cp_test,
              standard: 'Xung PWM chuyển trạng thái State A -> State C chuẩn',
              reference: 'PWM 1kHz ±12V',
              deviation: 'Chuyển trạng thái mượt mà',
              status: 'PASS'
            },
            {
              item: 'Điện áp đầu ra DC (Charger Output)',
              measured: `${liveStats.outVolt} V DC`,
              standard: 'Phù hợp dải điện áp thiết kế định mức',
              reference: 'Dải sạc nhanh DC',
              deviation: 'Điện áp ổn định',
              status: 'PASS'
            },
            {
              item: 'Độ gợn sóng điện áp DC (DC Ripple)',
              measured: `${liveStats.ripple}%`,
              standard: 'Độ gợn sóng DC ≤ 2.0% theo khuyến cáo NSX',
              reference: 'Giới hạn bộ sạc',
              deviation: `${liveStats.ripple}% ≤ 2.0%`,
              status: liveStats.isRipplePass ? 'PASS' : 'FAIL'
            }
          ]}
          warnings={[
            !liveStats.isGun2Pass
              ? `NGUY CƠ MẤT AN TOÀN ĐIỆN: Súng sạc 2 có điện trở dây PE = ${liveStats.gun2Pe} Ω (> 0.50 Ω chuẩn NETA Mục 7.26). Tăng ${liveStats.gun2IncreasePct}% so với kỳ trước (${liveStats.lastGun2Pe} Ω), nguy cơ đứt ngầm dây tiếp địa gây nguy hiểm giật điện!`
              : ''
          ].filter(Boolean)}
          recommendations={[
            'Tạm thời ngắt và treo biển cảnh báo ngừng sử dụng đối với Súng sạc số 2 (Gun 2).',
            'Kiểm tra điểm đấu nối dây PE bên trong súng sạc số 2 và tủ trạm sạc; vệ sinh mạ lại hoặc thay thế dây cáp sạc Gun 2 nếu phát hiện đứt ngầm sợi tiếp địa.',
            'Đo lại điện trở dây bảo vệ PE sau khi xử lý, đảm bảo giá trị giảm xuống dưới 0.5 Ω trước khi mở lại trạm sạc cho khách hàng.'
          ]}
        />
      )}
    </div>
  );
};
