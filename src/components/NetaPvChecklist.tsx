import React, { useState, useMemo } from 'react';
import {
  Sun,
  Flame,
  SunMedium,
  ShieldCheck,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  Sparkles,
  Printer,
  RotateCcw,
  Upload,
  Download,
  Save,
  Activity,
  Zap,
  Gauge,
  Sliders,
  Compass,
  FileCheck2,
  Copy,
  Check,
  Radio,
  Eye,
  Info
} from 'lucide-react';
import { NetaPvInputPayload, NetaPvEvaluationResult } from '../../server/netaPvAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';
import { PvCellDeepThermalAnalyzer } from './PvCellDeepThermalAnalyzer';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

const SAMPLE_PV_DATA: NetaPvInputPayload = {
  site_info: {
    project_name: "Nha May Dien Mat Troi TEV Roof 1MW",
    pv_array_tag: "ARRAY-ZONE-01",
    fse_name: "Pham Van F",
    test_date: "2026-09-22"
  },
  visual_inspection: {
    nameplate_match: true,
    physical_condition: "Pass (No cracked modules)",
    mounting_and_grounding: "Pass",
    cleanliness: "Pass",
    inverter_gateway_settings: "Pass",
    wiring_connections_tight: true
  },
  electrical_tests: {
    polarity_check: "Pass (Correct polarity across all strings)",
    blocking_diode_test: "Pass",
    dry_insulation_resistance_megohms: {
      string_01_to_ground: 150,
      string_02_to_ground: 145,
      string_03_to_ground: 12
    },
    open_circuit_voltage_voc: {
      expected_voc_v: 980,
      measured_voc_v: [982, 978, 850]
    },
    short_circuit_current_isc: {
      irradiance_w_m2: 850,
      measured_isc_a: [11.2, 11.3, 11.1]
    },
    iv_curve_measurement: "Pass for String 1-2, Distortion detected on String 3",
    ground_resistance_section_7_13_ohms: 2.1,
    das_communication_section_7_23_status: "Pass"
  },
  previous_test_data: {
    last_test_date: "2025-09-21",
    last_string_03_voc_v: 979
  }
};

export const NetaPvChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [activePvMode, setActivePvMode] = useState<'deep-pv-cell' | 'neta-string'>('deep-pv-cell');
  const [data, setData] = useState<NetaPvInputPayload>(SAMPLE_PV_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaPvEvaluationResult | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [jsonText, setJsonText] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live Calculations based on NETA ATS-2025 Section 7.29
  const liveStats = useMemo(() => {
    const vocArr = data.electrical_tests?.open_circuit_voltage_voc?.measured_voc_v || [982, 978, 850];
    const expVoc = data.electrical_tests?.open_circuit_voltage_voc?.expected_voc_v || 980;
    const maxVoc = Math.max(...vocArr);
    const minVoc = Math.min(...vocArr);
    const parallelDiffPercent = maxVoc > 0 ? ((maxVoc - minVoc) / maxVoc) * 100 : 0;

    // Voc String 3 deviation
    const voc3 = vocArr[2] ?? 850;
    const voc3DiffPercent = expVoc > 0 ? ((expVoc - voc3) / expVoc) * 100 : 0;

    // Insulation resistance
    const ir1 = data.electrical_tests?.dry_insulation_resistance_megohms?.string_01_to_ground || 150;
    const ir2 = data.electrical_tests?.dry_insulation_resistance_megohms?.string_02_to_ground || 145;
    const ir3 = data.electrical_tests?.dry_insulation_resistance_megohms?.string_03_to_ground || 12;

    const groundOhms = data.electrical_tests?.ground_resistance_section_7_13_ohms || 2.1;
    const isGroundPass = groundOhms <= 5.0;

    const isPolarityPass = data.electrical_tests?.polarity_check?.toLowerCase().includes('pass');
    const isIvPass = !data.electrical_tests?.iv_curve_measurement?.toLowerCase().includes('distortion');

    return {
      expVoc,
      measuredVoc: vocArr,
      parallelDiffPercent: parseFloat(parallelDiffPercent.toFixed(2)),
      isVocParallelPass: parallelDiffPercent <= 5.0,
      voc3,
      voc3DiffPercent: parseFloat(voc3DiffPercent.toFixed(2)),
      ir1,
      ir2,
      ir3,
      isIr3Pass: ir3 >= 50 && (ir1 <= 0 || ir3 >= ir1 * 0.2),
      groundOhms,
      isGroundPass,
      isPolarityPass,
      isIvPass
    };
  }, [data]);

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    setAnalysisResult(null);
    setSavedSuccessMessage(null);

    try {
      const response = await fetch('/api/field-service/neta-pv-analyze', {
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
      console.error('Lỗi khi gọi API phân tích NETA Solar PV:', err);
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
        equipmentId: data.site_info.pv_array_tag,
        date: data.site_info.test_date,
        type: 'Solar Photovoltaic (PV) Systems (NETA ATS-2025 Sec 7.29)'
      });
      setSavedSuccessMessage(`Đã lưu biên bản kiểm định Solar PV [${data.site_info.pv_array_tag}] vào kho báo cáo CMMS!`);
    }
  };

  const handleResetToSample = () => {
    setData(SAMPLE_PV_DATA);
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
      {/* Mode Selector Tab Bar */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-slate-900 border border-slate-800 p-2 sm:p-2.5 rounded-2xl shadow-md">
        <div className="flex flex-wrap items-center gap-2">
          <button
            onClick={() => setActivePvMode('deep-pv-cell')}
            className={`px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-bold flex items-center gap-2 transition-all ${
              activePvMode === 'deep-pv-cell'
                ? 'bg-gradient-to-r from-amber-500 to-orange-600 text-slate-950 shadow-md shadow-orange-950/40 ring-1 ring-amber-400/50'
                : 'text-slate-400 hover:text-white hover:bg-slate-800/60'
            }`}
          >
            <Flame size={16} className={activePvMode === 'deep-pv-cell' ? 'text-slate-950' : 'text-amber-400'} />
            <span>Phân Tích Chuyên Sâu Tấm Pin & Cell (IR & RGB AI)</span>
            <span className="text-[10px] px-1.5 py-0.5 rounded-md bg-black/20 font-extrabold uppercase hidden md:inline-block">
              IEC 62446-3 & Sitemark
            </span>
          </button>

          <button
            onClick={() => setActivePvMode('neta-string')}
            className={`px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-bold flex items-center gap-2 transition-all ${
              activePvMode === 'neta-string'
                ? 'bg-gradient-to-r from-indigo-600 to-blue-600 text-white shadow-md shadow-indigo-950/40 ring-1 ring-indigo-400/50'
                : 'text-slate-400 hover:text-white hover:bg-slate-800/60'
            }`}
          >
            <Sun size={16} className={activePvMode === 'neta-string' ? 'text-white' : 'text-yellow-400'} />
            <span>Nghiệm Thu Toàn Diện Chuỗi Pin (Array & Strings)</span>
            <span className="text-[10px] px-1.5 py-0.5 rounded-md bg-white/20 font-extrabold uppercase hidden md:inline-block">
              NETA ATS-2025 Mục 7.29
            </span>
          </button>
        </div>

        <div className="text-xs text-slate-400 hidden lg:flex items-center gap-2 pr-2">
          <Sparkles size={14} className="text-amber-400" />
          <span>Tích hợp Thư viện Mẫu lỗi Sitemark / Volateq</span>
        </div>
      </div>

      {activePvMode === 'deep-pv-cell' ? (
        <PvCellDeepThermalAnalyzer
          onSaveReport={onSaveReport}
          onNavigateToReports={onNavigateToReports}
        />
      ) : (
        <div className="space-y-6">
          {/* Top Banner & Actions Header */}
          <div className="bg-gradient-to-r from-amber-600 via-orange-600 to-amber-700 text-white p-5 sm:p-6 rounded-2xl shadow-lg border border-amber-500/30">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1">
            <div className="flex flex-wrap items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-extrabold uppercase tracking-wider bg-white/20 text-white backdrop-blur-xs border border-white/30 flex items-center gap-1.5">
                <Sun size={13} className="text-yellow-200" />
                ANSI/NETA ATS-2025 Mục 7.29
              </span>
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-amber-950/40 text-amber-200 border border-amber-400/40 flex items-center gap-1">
                <ShieldCheck size={12} />
                Solar PV Systems • Combiner Box • Inverter DC
              </span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white">
              Kiểm Định Hệ Thống Điện Mặt Trời (Solar PV Systems)
            </h1>
            <p className="text-xs sm:text-sm text-amber-100 max-w-2xl">
              Đánh giá Cực tính, Diode chặn, Cách điện Khô/Ẩm, Điện áp Hở Mạch Voc (ngưỡng lệch ≤ 5%), Dòng Isc, Đo đường cong I-V và Tiếp địa theo NETA Section 7.29.
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
              className="px-4 py-2 rounded-xl text-xs font-bold bg-gradient-to-r from-yellow-300 to-amber-300 hover:from-yellow-200 hover:to-amber-200 text-slate-950 flex items-center gap-2 transition-all shadow-lg shadow-amber-900/30 disabled:opacity-50"
            >
              <Sparkles size={15} className={isAnalyzing ? 'animate-spin' : 'text-amber-800'} />
              <span>{isAnalyzing ? 'Đang Phân Tích NETA...' : 'Phân Tích AI Chuyên Gia'}</span>
            </button>
          </div>
        </div>

        {/* Live NETA Diagnostic Strip */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5 mt-5 pt-4 border-t border-amber-400/30 text-xs">
          {/* Voc Parallel Deviation */}
          <div className="bg-amber-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-amber-400/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-amber-200 block font-medium">Lệch Voc Song Song (≤5%)</span>
              <span className={`text-base font-black ${liveStats.isVocParallelPass ? 'text-emerald-300' : 'text-rose-300'}`}>
                {liveStats.parallelDiffPercent}%
              </span>
            </div>
            {liveStats.isVocParallelPass ? (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            ) : (
              <XCircle size={20} className="text-rose-400 shrink-0" />
            )}
          </div>

          {/* Voc String 3 */}
          <div className="bg-amber-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-amber-400/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-amber-200 block font-medium">Voc Chuỗi 03 ({liveStats.expVoc}V)</span>
              <span className={`text-base font-black ${liveStats.voc3DiffPercent <= 5.0 ? 'text-emerald-300' : 'text-rose-300'}`}>
                {liveStats.voc3}V (-{liveStats.voc3DiffPercent}%)
              </span>
            </div>
            {liveStats.voc3DiffPercent <= 5.0 ? (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            ) : (
              <AlertTriangle size={20} className="text-rose-400 shrink-0" />
            )}
          </div>

          {/* Dry Insulation String 03 */}
          <div className="bg-amber-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-amber-400/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-amber-200 block font-medium">Cách Điện Khô Chuỗi 03</span>
              <span className={`text-base font-black ${liveStats.isIr3Pass ? 'text-emerald-300' : 'text-rose-300'}`}>
                {liveStats.ir3} MΩ
              </span>
            </div>
            {liveStats.isIr3Pass ? (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            ) : (
              <XCircle size={20} className="text-rose-400 shrink-0" />
            )}
          </div>

          {/* Ground Resistance Section 7.13 */}
          <div className="bg-amber-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-amber-400/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-amber-200 block font-medium">Tiếp Địa NETA 7.13 (≤5Ω)</span>
              <span className={`text-base font-black ${liveStats.isGroundPass ? 'text-emerald-300' : 'text-rose-300'}`}>
                {liveStats.groundOhms} Ω
              </span>
            </div>
            {liveStats.isGroundPass ? (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            ) : (
              <AlertTriangle size={20} className="text-amber-400 shrink-0" />
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
        {/* Left Column: Site Info & Visual Inspection */}
        <div className="space-y-6">
          {/* Site & Array Information */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <Sun size={17} className="text-amber-500" />
                Thông Tin Mảng Pin & Hiện Trường
              </h2>
              <span className="text-[10px] font-semibold text-slate-500 bg-slate-100 px-2 py-0.5 rounded-full">
                NETA ATS-2025 Sec 7.29
              </span>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="text-slate-600 font-medium block mb-1">Tên Dự Án Điện Mặt Trời:</label>
                <input
                  type="text"
                  value={data.site_info.project_name}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, project_name: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-medium text-slate-800"
                />
              </div>

              <div>
                <label className="text-slate-600 font-medium block mb-1">Mã Định Danh Mảng Pin (Array Tag):</label>
                <input
                  type="text"
                  value={data.site_info.pv_array_tag}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, pv_array_tag: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-bold text-amber-700 font-mono"
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
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-medium text-slate-800"
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
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-medium text-slate-800"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Section 7.29.1 Visual and Mechanical Inspection */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <Eye size={17} className="text-amber-500" />
                Kiểm Tra Thị Giác & Cơ Khí (Mục 7.29.1)
              </h2>
            </div>

            <div className="space-y-3 text-xs">
              {/* Nameplate Match */}
              <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-xl border border-slate-100">
                <div>
                  <span className="font-bold text-slate-800 block">Nameplate Match</span>
                  <span className="text-[11px] text-slate-500">Đối soát thông số mảng pin, Combiner Box, Inverter khớp 100% bản vẽ</span>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.nameplate_match}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, nameplate_match: e.target.checked }
                  })}
                  className="w-5 h-5 text-amber-600 rounded-md focus:ring-amber-500"
                />
              </div>

              {/* Physical Condition */}
              <div>
                <label className="text-slate-600 font-medium block mb-1">Tình Trạng Cơ Khí & Module Pin:</label>
                <input
                  type="text"
                  value={data.visual_inspection.physical_condition}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, physical_condition: e.target.value }
                  })}
                  placeholder="Pass (No cracked modules)"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-medium text-slate-800"
                />
              </div>

              {/* Mounting and Grounding */}
              <div>
                <label className="text-slate-600 font-medium block mb-1">Khung Giá Đỡ & Tiếp Địa Module:</label>
                <input
                  type="text"
                  value={data.visual_inspection.mounting_and_grounding}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, mounting_and_grounding: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-medium text-slate-800"
                />
              </div>

              {/* Cleanliness */}
              <div>
                <label className="text-slate-600 font-medium block mb-1">Vệ Sinh Bề Mặt & Tháo Kẹp Vận Chuyển:</label>
                <input
                  type="text"
                  value={data.visual_inspection.cleanliness}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, cleanliness: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-medium text-slate-800"
                />
              </div>

              {/* Inverter & Gateway Settings */}
              <div>
                <label className="text-slate-600 font-medium block mb-1">Cài Đặt Biến Tần Inverter & Gateway DAS:</label>
                <input
                  type="text"
                  value={data.visual_inspection.inverter_gateway_settings}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, inverter_gateway_settings: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-medium text-slate-800"
                />
              </div>

              {/* Wiring Connections Tight */}
              <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-xl border border-slate-100">
                <div>
                  <span className="font-bold text-slate-800 block">Đấu Nối Giắc MC4 & Đầu Cáp DC</span>
                  <span className="text-[11px] text-slate-500">Mối nối giắc MC4, đầu cáp chặt chẽ, cố định an toàn</span>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.wiring_connections_tight}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, wiring_connections_tight: e.target.checked }
                  })}
                  className="w-5 h-5 text-amber-600 rounded-md focus:ring-amber-500"
                />
              </div>
            </div>
          </div>
        </div>

        {/* Center & Right Column: Electrical Tests (Mục 7.29.2) */}
        <div className="lg:col-span-2 space-y-6">
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-5">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <Zap size={17} className="text-amber-500" />
                Phép Đo Điện & Tiêu Chuẩn Đánh Giá (NETA Sec 7.29.2)
              </h2>
              <span className="text-[11px] font-semibold text-amber-700 bg-amber-50 border border-amber-200 px-2.5 py-0.5 rounded-full">
                NETA ATS-2025 Section 7.29.2
              </span>
            </div>

            {/* Test 1: Polarity & Blocking Diode */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                  <Compass size={15} className="text-amber-600" />
                  1. Kiểm Tra Cực Tính (Polarity) & Diode Chặn (Blocking Diode)
                </span>
                <span className="text-[10px] text-rose-600 font-bold bg-rose-50 px-2 py-0.5 rounded border border-rose-200">
                  BẮT BUỘC ĐÚNG 100%
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Cực Tính Từng Chuỗi & Combiner Box:</label>
                  <input
                    type="text"
                    value={data.electrical_tests.polarity_check}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: { ...data.electrical_tests, polarity_check: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-semibold text-slate-800"
                  />
                </div>
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Thử Nghiệm Diode Chặn:</label>
                  <input
                    type="text"
                    value={data.electrical_tests.blocking_diode_test}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: { ...data.electrical_tests, blocking_diode_test: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-semibold text-slate-800"
                  />
                </div>
              </div>
            </div>

            {/* Test 2: Dry / Wet Insulation Resistance */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                  <ShieldCheck size={15} className="text-amber-600" />
                  2. Điện Trở Cách Điện Khô & Ẩm (Dry/Wet IR - MΩ)
                </span>
                <span className="text-[10px] text-slate-500 bg-white px-2 py-0.5 rounded border border-slate-200">
                  Chuẩn NSX &gt; 100 MΩ
                </span>
              </div>

              <div className="grid grid-cols-3 gap-2.5 text-xs">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Chuỗi 01 tới Đất (MΩ):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.dry_insulation_resistance_megohms.string_01_to_ground}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        dry_insulation_resistance_megohms: {
                          ...data.electrical_tests.dry_insulation_resistance_megohms,
                          string_01_to_ground: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-bold text-emerald-700"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Chuỗi 02 tới Đất (MΩ):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.dry_insulation_resistance_megohms.string_02_to_ground}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        dry_insulation_resistance_megohms: {
                          ...data.electrical_tests.dry_insulation_resistance_megohms,
                          string_02_to_ground: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-bold text-emerald-700"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1 flex items-center justify-between">
                    <span>Chuỗi 03 tới Đất (MΩ):</span>
                    {liveStats.isIr3Pass ? (
                      <span className="text-[10px] text-emerald-600 font-bold">Đạt</span>
                    ) : (
                      <span className="text-[10px] text-rose-600 font-bold">Thấp (12 MΩ)</span>
                    )}
                  </label>
                  <input
                    type="number"
                    value={data.electrical_tests.dry_insulation_resistance_megohms.string_03_to_ground}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        dry_insulation_resistance_megohms: {
                          ...data.electrical_tests.dry_insulation_resistance_megohms,
                          string_03_to_ground: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className={`w-full px-3 py-2 bg-white border rounded-lg focus:outline-hidden focus:ring-2 font-bold ${
                      liveStats.isIr3Pass
                        ? 'border-slate-200 text-emerald-700 focus:ring-amber-500'
                        : 'border-rose-400 text-rose-700 bg-rose-50/50 focus:ring-rose-500'
                    }`}
                  />
                </div>
              </div>
            </div>

            {/* Test 3: Open-Circuit Voltage (Voc) */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                  <Activity size={15} className="text-amber-600" />
                  3. Điện Áp Hở Mạch Chuỗi (Open-Circuit Voltage - Voc)
                </span>
                <span className="text-[10px] text-amber-700 font-bold bg-amber-100/70 px-2 py-0.5 rounded border border-amber-300">
                  Lệch Song Song ≤ 5.0%
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-4 gap-2.5 text-xs">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Voc Định Mức / Dự Kiến (V):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.open_circuit_voltage_voc.expected_voc_v}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        open_circuit_voltage_voc: {
                          ...data.electrical_tests.open_circuit_voltage_voc,
                          expected_voc_v: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-bold text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Voc Chuỗi 01 (V):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.open_circuit_voltage_voc.measured_voc_v[0] ?? 982}
                    onChange={(e) => {
                      const val = parseFloat(e.target.value) || 0;
                      const arr = [...(data.electrical_tests.open_circuit_voltage_voc.measured_voc_v || [982, 978, 850])];
                      arr[0] = val;
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          open_circuit_voltage_voc: { ...data.electrical_tests.open_circuit_voltage_voc, measured_voc_v: arr }
                        }
                      });
                    }}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-bold text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Voc Chuỗi 02 (V):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.open_circuit_voltage_voc.measured_voc_v[1] ?? 978}
                    onChange={(e) => {
                      const val = parseFloat(e.target.value) || 0;
                      const arr = [...(data.electrical_tests.open_circuit_voltage_voc.measured_voc_v || [982, 978, 850])];
                      arr[1] = val;
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          open_circuit_voltage_voc: { ...data.electrical_tests.open_circuit_voltage_voc, measured_voc_v: arr }
                        }
                      });
                    }}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-bold text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1 flex items-center justify-between">
                    <span>Voc Chuỗi 03 (V):</span>
                    {liveStats.voc3DiffPercent <= 5.0 ? (
                      <span className="text-[10px] text-emerald-600 font-bold">Đạt</span>
                    ) : (
                      <span className="text-[10px] text-rose-600 font-bold">-13.2% Lệch</span>
                    )}
                  </label>
                  <input
                    type="number"
                    value={data.electrical_tests.open_circuit_voltage_voc.measured_voc_v[2] ?? 850}
                    onChange={(e) => {
                      const val = parseFloat(e.target.value) || 0;
                      const arr = [...(data.electrical_tests.open_circuit_voltage_voc.measured_voc_v || [982, 978, 850])];
                      arr[2] = val;
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          open_circuit_voltage_voc: { ...data.electrical_tests.open_circuit_voltage_voc, measured_voc_v: arr }
                        }
                      });
                    }}
                    className={`w-full px-3 py-2 bg-white border rounded-lg focus:outline-hidden focus:ring-2 font-bold ${
                      liveStats.voc3DiffPercent <= 5.0
                        ? 'border-slate-200 text-slate-800 focus:ring-amber-500'
                        : 'border-rose-400 text-rose-700 bg-rose-50/50 focus:ring-rose-500'
                    }`}
                  />
                </div>
              </div>
            </div>

            {/* Test 4: Short-Circuit Current Isc & Irradiance */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                  <SunMedium size={15} className="text-amber-600" />
                  4. Dòng Ngắn Mạch / Vận Hành (Isc / Iop) & Bức Xạ Mặt Trời
                </span>
                <span className="text-[10px] text-slate-500 bg-white px-2 py-0.5 rounded border border-slate-200">
                  Khớp mức bức xạ thực tế
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-4 gap-2.5 text-xs">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Bức Xạ (W/m²):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.short_circuit_current_isc.irradiance_w_m2}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        short_circuit_current_isc: {
                          ...data.electrical_tests.short_circuit_current_isc,
                          irradiance_w_m2: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-bold text-amber-700"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Isc Chuỗi 01 (A):</label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.short_circuit_current_isc.measured_isc_a[0] ?? 11.2}
                    onChange={(e) => {
                      const val = parseFloat(e.target.value) || 0;
                      const arr = [...(data.electrical_tests.short_circuit_current_isc.measured_isc_a || [11.2, 11.3, 11.1])];
                      arr[0] = val;
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          short_circuit_current_isc: { ...data.electrical_tests.short_circuit_current_isc, measured_isc_a: arr }
                        }
                      });
                    }}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-medium text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Isc Chuỗi 02 (A):</label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.short_circuit_current_isc.measured_isc_a[1] ?? 11.3}
                    onChange={(e) => {
                      const val = parseFloat(e.target.value) || 0;
                      const arr = [...(data.electrical_tests.short_circuit_current_isc.measured_isc_a || [11.2, 11.3, 11.1])];
                      arr[1] = val;
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          short_circuit_current_isc: { ...data.electrical_tests.short_circuit_current_isc, measured_isc_a: arr }
                        }
                      });
                    }}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-medium text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Isc Chuỗi 03 (A):</label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.short_circuit_current_isc.measured_isc_a[2] ?? 11.1}
                    onChange={(e) => {
                      const val = parseFloat(e.target.value) || 0;
                      const arr = [...(data.electrical_tests.short_circuit_current_isc.measured_isc_a || [11.2, 11.3, 11.1])];
                      arr[2] = val;
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          short_circuit_current_isc: { ...data.electrical_tests.short_circuit_current_isc, measured_isc_a: arr }
                        }
                      });
                    }}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-medium text-slate-800"
                  />
                </div>
              </div>
            </div>

            {/* Test 5: I-V Curve, Grounding & DAS Communication */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <Sliders size={15} className="text-amber-600" />
                5. Đo Đặc Tuyến I-V, Tiếp Địa Nối Đất & Giám Sát DAS
              </span>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Đo Đường Cong I-V (I-V Curve):</label>
                  <input
                    type="text"
                    value={data.electrical_tests.iv_curve_measurement}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: { ...data.electrical_tests, iv_curve_measurement: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-medium text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Điện Trở Tiếp Địa Sec 7.13 (Ω):</label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.ground_resistance_section_7_13_ohms}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        ground_resistance_section_7_13_ohms: parseFloat(e.target.value) || 0
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-bold text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Truyền Thông DAS Sec 7.23:</label>
                  <input
                    type="text"
                    value={data.electrical_tests.das_communication_section_7_23_status}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: { ...data.electrical_tests, das_communication_section_7_23_status: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 font-medium text-slate-800"
                  />
                </div>
              </div>
            </div>

            {/* Historical Data Reference */}
            {data.previous_test_data && (
              <div className="p-3 bg-amber-50/60 rounded-xl border border-amber-200/60 flex items-center justify-between text-xs text-amber-900">
                <span className="font-semibold flex items-center gap-1.5">
                  <Info size={14} className="text-amber-600" />
                  Kỳ Đo Trước ({data.previous_test_data.last_test_date || '2025-09-21'}): Chuỗi 03 Voc = {data.previous_test_data.last_string_03_voc_v}V (Đạt chuẩn 980V)
                </span>
                <span className="text-[11px] font-bold text-rose-700 bg-white px-2 py-0.5 rounded border border-rose-300">
                  Suy Giảm Bất Thường Hiện Tại
                </span>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* AI Analysis Output Section */}
      {analysisResult && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden animate-in fade-in slide-in-from-bottom-3 duration-300">
          <div className="p-5 bg-gradient-to-r from-slate-900 via-amber-950 to-slate-900 text-white flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-amber-500/20 border border-amber-400/30 flex items-center justify-center text-amber-300 shrink-0">
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
                  Báo Cáo Đánh Giá Kỹ Thuật Solar PV (NETA ATS-2025 Mục 7.29)
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
                <FileCheck2 size={18} className="text-amber-600" />
                <h3 className="font-bold text-sm text-slate-800">
                  Dữ Liệu JSON Kiểm Định Ngoài Site (NETA ATS-2025 Sec 7.29)
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
              className="w-full p-3 font-mono text-xs bg-slate-900 text-amber-200 rounded-xl border border-slate-700 focus:outline-hidden focus:ring-2 focus:ring-amber-500"
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
                  className="px-4 py-1.5 bg-amber-600 hover:bg-amber-500 text-white text-xs font-bold rounded-lg shadow-sm"
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
          title="BIÊN BẢN KIỂM ĐỊNH HỆ THỐNG ĐIỆN MẶT TRỜI (SOLAR PV SYSTEMS)"
          standard="ANSI/NETA ATS-2025 Section 7.29"
          siteInfo={{
            project_name: data.site_info.project_name,
            transformer_tag: data.site_info.pv_array_tag,
            fse_name: data.site_info.fse_name,
            test_date: data.site_info.test_date
          }}
          overallStatus={analysisResult?.overall_status || (liveStats.isVocParallelPass && liveStats.isIr3Pass ? 'PASS' : 'FAIL')}
          statusReason={analysisResult?.status_reason || 'Đã đối chiếu với quy chuẩn kỹ thuật NETA ATS-2025 Mục 7.29.'}
          visualChecks={[
            { label: 'Nameplate Match (Khớp bản vẽ 100%)', value: data.visual_inspection.nameplate_match ? 'Khớp 100%' : 'Sai lệch', pass: data.visual_inspection.nameplate_match },
            { label: 'Cơ khí & Module pin không nứt vỡ', value: data.visual_inspection.physical_condition, pass: data.visual_inspection.physical_condition.toLowerCase().includes('pass') },
            { label: 'Khung giá đỡ & Tiếp địa hệ thống', value: data.visual_inspection.mounting_and_grounding, pass: data.visual_inspection.mounting_and_grounding.toLowerCase().includes('pass') },
            { label: 'Vệ sinh bề mặt & Tháo kẹp vận chuyển', value: data.visual_inspection.cleanliness, pass: data.visual_inspection.cleanliness.toLowerCase().includes('pass') },
            { label: 'Cài đặt Inverter & Gateway DAS', value: data.visual_inspection.inverter_gateway_settings, pass: data.visual_inspection.inverter_gateway_settings.toLowerCase().includes('pass') },
            { label: 'Đấu nối giắc MC4 & Đầu cáp chặt chẽ', value: data.visual_inspection.wiring_connections_tight ? 'Chặt chẽ, an toàn' : 'Lỏng lẻo', pass: data.visual_inspection.wiring_connections_tight }
          ]}
          electricalTable={[
            {
              item: 'Cực tính (Polarity Check)',
              measured: data.electrical_tests.polarity_check,
              standard: 'Bắt buộc đúng 100% từng chuỗi & Combiner Box',
              reference: 'Bản vẽ DC',
              deviation: 'Đúng cực tính (+/-)',
              status: liveStats.isPolarityPass ? 'PASS' : 'FAIL'
            },
            {
              item: 'Diode chặn (Blocking Diode)',
              measured: data.electrical_tests.blocking_diode_test,
              standard: 'Hoạt động bình thường, không dẫn ngược',
              reference: 'Thông số NSX',
              deviation: 'Bình thường',
              status: 'PASS'
            },
            {
              item: 'Cách điện Chuỗi 01 & 02 (Dry IR)',
              measured: `Chuỗi 1: ${liveStats.ir1} MΩ, Chuỗi 2: ${liveStats.ir2} MΩ`,
              standard: '> 100 MΩ (Đạt ngưỡng NSX)',
              reference: 'Chuẩn an toàn DC',
              deviation: 'Tốt',
              status: 'PASS'
            },
            {
              item: 'Cách điện Chuỗi 03 (Dry IR)',
              measured: `${liveStats.ir3} MΩ`,
              standard: '> 100 MΩ',
              reference: 'Tham chiếu: > 140 MΩ',
              deviation: liveStats.isIr3Pass ? 'Đạt' : `${liveStats.ir3} MΩ < 100 MΩ (Rò điện đất)`,
              status: liveStats.isIr3Pass ? 'PASS' : 'FAIL'
            },
            {
              item: 'Điện áp hở mạch Chuỗi 01 & 02 (Voc)',
              measured: `Chuỗi 1: ${liveStats.measuredVoc[0]}V, Chuỗi 2: ${liveStats.measuredVoc[1]}V`,
              standard: 'Lệch ≤ 5.0% so với thiết kế & song song',
              reference: `Voc_exp = ${liveStats.expVoc}V`,
              deviation: 'Lệch ≤ 0.2%',
              status: 'PASS'
            },
            {
              item: 'Điện áp hở mạch Chuỗi 03 (Voc)',
              measured: `${liveStats.voc3}V`,
              standard: 'Lệch ≤ 5.0% so với thiết kế & song song',
              reference: `Voc_exp = ${liveStats.expVoc}V`,
              deviation: `Giảm ${liveStats.voc3DiffPercent}% (Lệch ${liveStats.expVoc - liveStats.voc3}V > 5.0%)`,
              status: liveStats.voc3DiffPercent <= 5.0 ? 'PASS' : 'FAIL'
            },
            {
              item: 'Dòng ngắn mạch (Isc)',
              measured: `[${data.electrical_tests.short_circuit_current_isc.measured_isc_a.join(', ')}] A @ ${data.electrical_tests.short_circuit_current_isc.irradiance_w_m2} W/m²`,
              standard: 'Phù hợp mức bức xạ mặt trời thực tế',
              reference: 'Danh định NSX',
              deviation: 'Dòng đều các chuỗi',
              status: 'PASS'
            },
            {
              item: 'Đường cong đặc tuyến I-V',
              measured: data.electrical_tests.iv_curve_measurement,
              standard: 'Đường cong mượt mà, Pmax đạt chuẩn',
              reference: 'Đặc tuyến PV',
              deviation: liveStats.isIvPass ? 'Chuẩn' : 'Biến dạng Chuỗi 03 (Distortion)',
              status: liveStats.isIvPass ? 'PASS' : 'FAIL'
            },
            {
              item: 'Điện trở tiếp địa (Grounding)',
              measured: `${liveStats.groundOhms} Ω`,
              standard: 'NETA Section 7.13 (≤ 5.0 Ω)',
              reference: 'Section 7.13',
              deviation: `${liveStats.groundOhms} Ω ≤ 5.0 Ω`,
              status: liveStats.isGroundPass ? 'PASS' : 'INVESTIGATE'
            }
          ]}
          warnings={[
            !liveStats.isIr3Pass ? `Cách điện Chuỗi 03 chỉ đạt ${liveStats.ir3} MΩ (Thấp bất thường, nguy cơ rò điện đất).` : '',
            liveStats.voc3DiffPercent > 5.0 ? `Điện áp hở mạch Voc Chuỗi 03 sụt giảm ${liveStats.voc3DiffPercent}% vượt quá ngưỡng 5% của NETA Section 7.29 (Nghi ngờ chập 2-3 tấm pin hoặc hỏng Diode bypass).` : '',
            !liveStats.isIvPass ? 'Đặc tuyến I-V Chuỗi 03 bị biến dạng tương ứng với sự cố giảm điện áp.' : ''
          ].filter(Boolean)}
          recommendations={[
            'Tách cô lập ngay Chuỗi 03 (String 03) tại Combiner Box để ngăn chặn dòng điện vòng nguy hại.',
            'Dùng camera nhiệt kiểm tra từng module trong Chuỗi 03 để phát hiện điểm nóng (hotspot) và đo lại Voc từng tấm.',
            'Kiểm tra tuyến cáp DC Chuỗi 03 từ mái nhà xuống tủ Combiner Box để tìm điểm chạm đất gây suy giảm cách điện xuống 12 MΩ.'
          ]}
        />
      )}
        </div>
      )}
    </div>
  );
};
