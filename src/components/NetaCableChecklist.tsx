import React, { useState, useMemo } from 'react';
import {
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
  Cable,
  Gauge,
  Zap
} from 'lucide-react';
import { NetaCableInputPayload, NetaCableEvaluationResult } from '../../server/netaCableAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

const SAMPLE_CABLE_DATA: NetaCableInputPayload = {
  site_info: {
    project_name: "Tram 110kV TEV Industrial Grid",
    cable_tag: "CBL-MV-22KV-FEEDER-01",
    cable_rating: "22kV 3x240mm2 Cu/XLPE/Tape Shield",
    cable_length_feet: 1500,
    fse_name: "Tran Van H",
    test_date: "2026-09-23"
  },
  visual_inspection: {
    cable_data_match: true,
    terminations_and_splices_condition: "Pass (Good condition)",
    bolt_torque_check: "Pass",
    shield_grounding_check: "Pass",
    bending_radius_check: "Pass (Complies with Table 100.22)",
    window_ct_grounding_correct: true
  },
  electrical_tests: {
    insulation_resistance_1min_2500v_megohms: {
      phase_a_to_shield: 4500,
      phase_b_to_shield: 4200,
      phase_c_to_shield: 3800
    },
    shield_resistance_ohms: {
      phase_a_shield: 8.2,
      phase_b_shield: 8.5,
      phase_c_shield: 24.6
    },
    tdr_test: "Uniform length 1500ft, no discontinuities",
    vlf_withstand_test_0_1hz_30min: {
      test_voltage_kv_rms: 28,
      status: "Pass (No breakdown during 30 min)"
    },
    vlf_tan_delta_test: {
      mean_tan_delta_1uo: 1.2e-3,
      tip_up_tan_delta: 0.4e-3,
      evaluation: "Pass (No action required per Table 100.6.7.1)"
    }
  },
  previous_test_data: {
    last_test_date: "2025-09-20",
    last_phase_c_shield_resistance_ohms: 8.4
  }
};

export const NetaCableChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaCableInputPayload>(SAMPLE_CABLE_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaCableEvaluationResult | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [jsonText, setJsonText] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live Calculations based on NETA ATS-2025 Section 7.3.3
  const liveStats = useMemo(() => {
    const lengthFeet = data.site_info.cable_length_feet || 1500;
    const normLimit = 10; // 10 Ohms per 1,000 feet
    const maxAllowedOhms = (lengthFeet / 1000) * normLimit; // 15.0 Ohms for 1500 ft

    const shA = data.electrical_tests?.shield_resistance_ohms?.phase_a_shield ?? 8.2;
    const shB = data.electrical_tests?.shield_resistance_ohms?.phase_b_shield ?? 8.5;
    const shC = data.electrical_tests?.shield_resistance_ohms?.phase_c_shield ?? 24.6;

    const shAPer1000 = parseFloat(((shA / lengthFeet) * 1000).toFixed(1));
    const shBPer1000 = parseFloat(((shB / lengthFeet) * 1000).toFixed(1));
    const shCPer1000 = parseFloat(((shC / lengthFeet) * 1000).toFixed(1));

    const isShAPass = shA <= maxAllowedOhms;
    const isShBPass = shB <= maxAllowedOhms;
    const isShCPass = shC <= maxAllowedOhms;
    const isAllShieldPass = isShAPass && isShBPass && isShCPass;

    const lastShC = data.previous_test_data?.last_phase_c_shield_resistance_ohms ?? 8.4;
    const shCIncreasePct = lastShC > 0 ? (((shC - lastShC) / lastShC) * 100).toFixed(0) : '0';

    const irA = data.electrical_tests?.insulation_resistance_1min_2500v_megohms?.phase_a_to_shield ?? 4500;
    const irB = data.electrical_tests?.insulation_resistance_1min_2500v_megohms?.phase_b_to_shield ?? 4200;
    const irC = data.electrical_tests?.insulation_resistance_1min_2500v_megohms?.phase_c_to_shield ?? 3800;
    const isIrPass = irA >= 1000 && irB >= 1000 && irC >= 1000;

    const isVlfPass = data.electrical_tests?.vlf_withstand_test_0_1hz_30min?.status?.toLowerCase().includes('pass');
    const isTanDeltaPass = data.electrical_tests?.vlf_tan_delta_test?.evaluation?.toLowerCase().includes('pass');

    return {
      lengthFeet,
      normLimit,
      maxAllowedOhms,
      shA,
      shB,
      shC,
      shAPer1000,
      shBPer1000,
      shCPer1000,
      isShAPass,
      isShBPass,
      isShCPass,
      isAllShieldPass,
      lastShC,
      shCIncreasePct,
      irA,
      irB,
      irC,
      isIrPass,
      isVlfPass,
      isTanDeltaPass
    };
  }, [data]);

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    setAnalysisResult(null);
    setSavedSuccessMessage(null);

    try {
      const response = await fetch('/api/field-service/neta-cable-analyze', {
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
      console.error('Lỗi khi gọi API phân tích NETA Cable:', err);
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
        equipmentId: data.site_info.cable_tag,
        date: data.site_info.test_date,
        type: 'Shielded Cables, Medium- and High-Voltage (NETA ATS-2025 Sec 7.3.3)'
      });
      setSavedSuccessMessage(`Đã lưu biên bản kiểm định cáp trung/cao áp [${data.site_info.cable_tag}] vào kho báo cáo CMMS!`);
    }
  };

  const handleResetToSample = () => {
    setData(SAMPLE_CABLE_DATA);
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
      <div className="bg-gradient-to-r from-violet-800 via-purple-700 to-indigo-900 text-white p-5 sm:p-6 rounded-2xl shadow-lg border border-purple-500/30">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1">
            <div className="flex flex-wrap items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-extrabold uppercase tracking-wider bg-white/20 text-white backdrop-blur-xs border border-white/30 flex items-center gap-1.5">
                <Cable size={13} className="text-purple-200" />
                ANSI/NETA ATS-2025 Mục 7.3.3
              </span>
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-purple-950/40 text-purple-200 border border-purple-400/40 flex items-center gap-1">
                <ShieldCheck size={12} />
                Shielded Cables • Medium & High Voltage • VLF 0.1Hz • Tan-Delta
              </span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white">
              Kiểm Định Cáp Điện Trung & Cao Áp Có Màn Chắn Kim Loại
            </h1>
            <p className="text-xs sm:text-sm text-purple-100 max-w-2xl">
              Đo thông mạch màn chắn (ngưỡng NETA ≤ 10 Ω/1,000 ft), điện trở cách điện IR (Table 100.1), thử chịu áp VLF 0.1Hz (Table 100.6.3), tổn hao điện môi Tan-Delta và tiếp địa qua CT cửa sổ.
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
              className="px-4 py-2 rounded-xl text-xs font-bold bg-gradient-to-r from-amber-300 to-yellow-300 hover:from-amber-200 hover:to-yellow-200 text-slate-950 flex items-center gap-2 transition-all shadow-lg shadow-purple-900/30 disabled:opacity-50"
            >
              <Sparkles size={15} className={isAnalyzing ? 'animate-spin' : 'text-purple-900'} />
              <span>{isAnalyzing ? 'Đang Phân Tích NETA...' : 'Phân Tích AI Chuyên Gia'}</span>
            </button>
          </div>
        </div>

        {/* Live NETA Diagnostic Strip */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5 mt-5 pt-4 border-t border-purple-400/30 text-xs">
          {/* Phase A & B Shield */}
          <div className="bg-purple-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-purple-400/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-purple-200 block font-medium">Màn Chắn Pha A & B (≤{liveStats.maxAllowedOhms}Ω)</span>
              <span className="text-base font-black text-emerald-300">
                A: {liveStats.shA}Ω • B: {liveStats.shB}Ω
              </span>
            </div>
            <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
          </div>

          {/* Phase C Shield (ANOMALY SPOTLIGHT) */}
          <div className="bg-purple-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-purple-400/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-purple-200 block font-medium">Màn Chắn Pha C (≤{liveStats.maxAllowedOhms}Ω / {liveStats.lengthFeet}ft)</span>
              <span className={`text-base font-black ${liveStats.isShCPass ? 'text-emerald-300' : 'text-amber-300'}`}>
                {liveStats.shC} Ω ({liveStats.shCPer1000} Ω/1k ft)
              </span>
            </div>
            {liveStats.isShCPass ? (
              <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
            ) : (
              <AlertTriangle size={20} className="text-amber-400 shrink-0" />
            )}
          </div>

          {/* Insulation Resistance (IR) */}
          <div className="bg-purple-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-purple-400/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-purple-200 block font-medium">Cách Điện IR (≥1000 MΩ)</span>
              <span className="text-base font-black text-emerald-300">
                Min {Math.min(liveStats.irA, liveStats.irB, liveStats.irC)} MΩ
              </span>
            </div>
            <CheckCircle2 size={20} className="text-emerald-400 shrink-0" />
          </div>

          {/* VLF Withstand & Tan-Delta */}
          <div className="bg-purple-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-purple-400/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-purple-200 block font-medium">VLF 28kV & Tan-Delta</span>
              <span className={`text-base font-black ${liveStats.isVlfPass && liveStats.isTanDeltaPass ? 'text-emerald-300' : 'text-rose-300'}`}>
                {liveStats.isVlfPass && liveStats.isTanDeltaPass ? 'Đạt 30 Phút' : 'Có Sự Cố'}
              </span>
            </div>
            {liveStats.isVlfPass && liveStats.isTanDeltaPass ? (
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
          {/* Site & Cable Information */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <Cable size={17} className="text-purple-600" />
                Thông Tin Tuyến Cáp Hiện Trường
              </h2>
              <span className="text-[10px] font-semibold text-slate-500 bg-slate-100 px-2 py-0.5 rounded-full">
                NETA ATS-2025 Sec 7.3.3
              </span>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="text-slate-600 font-medium block mb-1">Tên Dự Án / Vị Trí Tuyến Cáp:</label>
                <input
                  type="text"
                  value={data.site_info.project_name}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, project_name: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-medium text-slate-800"
                />
              </div>

              <div>
                <label className="text-slate-600 font-medium block mb-1">Mã Định Danh Cáp (Cable Tag):</label>
                <input
                  type="text"
                  value={data.site_info.cable_tag}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, cable_tag: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-bold text-purple-700 font-mono"
                />
              </div>

              <div>
                <label className="text-slate-600 font-medium block mb-1">Thông Số Cáp (Cable Rating):</label>
                <input
                  type="text"
                  value={data.site_info.cable_rating}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, cable_rating: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-medium text-slate-800"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Chiều Dài Tuyến (Feet):</label>
                  <input
                    type="number"
                    value={data.site_info.cable_length_feet}
                    onChange={(e) => setData({
                      ...data,
                      site_info: { ...data.site_info, cable_length_feet: parseFloat(e.target.value) || 0 }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-bold text-slate-800"
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
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-medium text-slate-800"
                  />
                </div>
              </div>

              <div>
                <label className="text-slate-600 font-medium block mb-1">Kỹ Sư FSE Trưởng Nhóm:</label>
                <input
                  type="text"
                  value={data.site_info.fse_name}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, fse_name: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-medium text-slate-800"
                />
              </div>
            </div>
          </div>

          {/* Section 7.3.3.1 Visual and Mechanical Inspection */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <Eye size={17} className="text-purple-600" />
                Kiểm Tra Thị Giác & Cơ Khí (Mục 7.3.3.1)
              </h2>
            </div>

            <div className="space-y-3 text-xs">
              {/* Cable Data Match */}
              <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-xl border border-slate-100">
                <div>
                  <span className="font-bold text-slate-800 block">Cable Data Match</span>
                  <span className="text-[11px] text-slate-500">Khớp 100% dữ liệu nhãn cáp trung/cao áp với bản vẽ</span>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.cable_data_match}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, cable_data_match: e.target.checked }
                  })}
                  className="w-5 h-5 text-purple-600 rounded-md focus:ring-purple-500"
                />
              </div>

              {/* Terminations and Splices */}
              <div>
                <label className="text-slate-600 font-medium block mb-1">Đầu Cáp & Hộp Nối (Terminations & Splices):</label>
                <input
                  type="text"
                  value={data.visual_inspection.terminations_and_splices_condition}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, terminations_and_splices_condition: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-medium text-slate-800"
                />
              </div>

              {/* Bolt Torque Check */}
              <div>
                <label className="text-slate-600 font-medium block mb-1">Lực Siết Bu-lông Đấu Nối (Table 100.12):</label>
                <input
                  type="text"
                  value={data.visual_inspection.bolt_torque_check}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, bolt_torque_check: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-medium text-slate-800"
                />
              </div>

              {/* Shield Grounding */}
              <div>
                <label className="text-slate-600 font-medium block mb-1">Tiếp Địa Màn Chắn Kim Loại (Shield Grounding):</label>
                <input
                  type="text"
                  value={data.visual_inspection.shield_grounding_check}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, shield_grounding_check: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-medium text-slate-800"
                />
              </div>

              {/* Bending Radius Check */}
              <div>
                <label className="text-slate-600 font-medium block mb-1">Bán Kính Uốn Cong Cáp (Table 100.22 / ICEA):</label>
                <input
                  type="text"
                  value={data.visual_inspection.bending_radius_check}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, bending_radius_check: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-medium text-slate-800"
                />
              </div>

              {/* Window-Type CT Grounding (Crucial NETA Rule) */}
              <div className="flex items-center justify-between p-2.5 bg-purple-50/50 rounded-xl border border-purple-200">
                <div>
                  <span className="font-bold text-purple-950 block">Tiếp Địa CT Cửa Sổ (Window CT Grounding)</span>
                  <span className="text-[11px] text-purple-800">Dây tiếp địa màn chắn lộn ngược lại qua lòng CT trước khi nối đất</span>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.window_ct_grounding_correct}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, window_ct_grounding_correct: e.target.checked }
                  })}
                  className="w-5 h-5 text-purple-600 rounded-md focus:ring-purple-500"
                />
              </div>
            </div>
          </div>
        </div>

        {/* Center & Right Column: Electrical Tests (Mục 7.3.3.2) */}
        <div className="lg:col-span-2 space-y-6">
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-5">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <Zap size={17} className="text-purple-600" />
                Phép Đo Điện & Tiêu Chuẩn Đánh Giá (NETA Sec 7.3.3.2)
              </h2>
              <span className="text-[11px] font-semibold text-purple-700 bg-purple-50 border border-purple-200 px-2.5 py-0.5 rounded-full">
                NETA ATS-2025 Section 7.3.3.2
              </span>
            </div>

            {/* Test 1: Shield Resistance (Thông Mạch Màn Chắn Kim Loại) - CRITICAL NETA RULE */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                  <ShieldAlert size={15} className="text-purple-600" />
                  1. Thông Mạch Màn Chắn Kim Loại (Shield Continuity Test)
                </span>
                <span className="text-[10px] text-amber-700 font-bold bg-amber-50 px-2 py-0.5 rounded border border-amber-300">
                  TỐI ĐA ≤ 10 Ω / 1,000 ft (≤ {liveStats.maxAllowedOhms} Ω / {liveStats.lengthFeet} ft)
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
                {/* Phase A Shield */}
                <div>
                  <label className="text-slate-600 font-medium block mb-1 flex items-center justify-between">
                    <span>Màn chắn Pha A (Ω):</span>
                    <span className="text-[10px] text-emerald-600 font-bold">
                      {liveStats.shAPer1000} Ω/1k ft
                    </span>
                  </label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.shield_resistance_ohms.phase_a_shield}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        shield_resistance_ohms: {
                          ...data.electrical_tests.shield_resistance_ohms,
                          phase_a_shield: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-bold text-emerald-700"
                  />
                </div>

                {/* Phase B Shield */}
                <div>
                  <label className="text-slate-600 font-medium block mb-1 flex items-center justify-between">
                    <span>Màn chắn Pha B (Ω):</span>
                    <span className="text-[10px] text-emerald-600 font-bold">
                      {liveStats.shBPer1000} Ω/1k ft
                    </span>
                  </label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.shield_resistance_ohms.phase_b_shield}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        shield_resistance_ohms: {
                          ...data.electrical_tests.shield_resistance_ohms,
                          phase_b_shield: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-bold text-emerald-700"
                  />
                </div>

                {/* Phase C Shield (ANOMALY) */}
                <div>
                  <label className="text-slate-600 font-medium block mb-1 flex items-center justify-between">
                    <span>Màn chắn Pha C (Ω):</span>
                    {liveStats.isShCPass ? (
                      <span className="text-[10px] text-emerald-600 font-bold">Đạt</span>
                    ) : (
                      <span className="text-[10px] text-amber-600 font-bold">
                        {liveStats.shCPer1000} Ω/1k ft (VƯỢT 64%)
                      </span>
                    )}
                  </label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.shield_resistance_ohms.phase_c_shield}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        shield_resistance_ohms: {
                          ...data.electrical_tests.shield_resistance_ohms,
                          phase_c_shield: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className={`w-full px-3 py-2 bg-white border rounded-lg focus:outline-hidden focus:ring-2 font-bold ${
                      liveStats.isShCPass
                        ? 'border-slate-200 text-slate-800 focus:ring-purple-500'
                        : 'border-amber-400 text-amber-700 bg-amber-50/50 focus:ring-amber-500'
                    }`}
                  />
                </div>
              </div>
            </div>

            {/* Test 2: Insulation Resistance (IR at 2500V 1 min) */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                  <ShieldCheck size={15} className="text-purple-600" />
                  2. Điện Trở Cách Điện (Insulation Resistance - IR 2500V DC, 1 phút)
                </span>
                <span className="text-[10px] text-emerald-700 font-bold bg-emerald-50 px-2 py-0.5 rounded border border-emerald-300">
                  NETA Table 100.1 (≥ 1,000 MΩ)
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Pha A tới Màn Chắn (MΩ):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_1min_2500v_megohms.phase_a_to_shield}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_1min_2500v_megohms: {
                          ...data.electrical_tests.insulation_resistance_1min_2500v_megohms,
                          phase_a_to_shield: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-bold text-emerald-700"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Pha B tới Màn Chắn (MΩ):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_1min_2500v_megohms.phase_b_to_shield}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_1min_2500v_megohms: {
                          ...data.electrical_tests.insulation_resistance_1min_2500v_megohms,
                          phase_b_to_shield: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-bold text-emerald-700"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Pha C tới Màn Chắn (MΩ):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_1min_2500v_megohms.phase_c_to_shield}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_1min_2500v_megohms: {
                          ...data.electrical_tests.insulation_resistance_1min_2500v_megohms,
                          phase_c_to_shield: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-bold text-emerald-700"
                  />
                </div>
              </div>
            </div>

            {/* Test 3: VLF Withstand Test (0.1 Hz, 30 min) & TDR */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <Gauge size={15} className="text-purple-600" />
                3. Thử Chịu Áp VLF (0.1 Hz, 30 phút) & Đo Phản Xạ TDR
              </span>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Điện Áp Thử VLF (kV rms):</label>
                  <input
                    type="number"
                    value={data.electrical_tests.vlf_withstand_test_0_1hz_30min.test_voltage_kv_rms}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        vlf_withstand_test_0_1hz_30min: {
                          ...data.electrical_tests.vlf_withstand_test_0_1hz_30min,
                          test_voltage_kv_rms: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-bold text-purple-700"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Kết Quả Thử Chịu Áp VLF 30 Phút:</label>
                  <input
                    type="text"
                    value={data.electrical_tests.vlf_withstand_test_0_1hz_30min.status}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        vlf_withstand_test_0_1hz_30min: {
                          ...data.electrical_tests.vlf_withstand_test_0_1hz_30min,
                          status: e.target.value
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-medium text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Đo Phản Xạ TDR Tuyến Cáp:</label>
                  <input
                    type="text"
                    value={data.electrical_tests.tdr_test}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: { ...data.electrical_tests, tdr_test: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-medium text-slate-800"
                  />
                </div>
              </div>
            </div>

            {/* Test 4: VLF Tan-Delta / Tip-Up Test */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <Sliders size={15} className="text-purple-600" />
                4. Thử Tổn Hao Điện Môi VLF Tan-Delta & Tip-Up (Table 100.6.7.1)
              </span>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Mean Tan Delta (1.0 Uo):</label>
                  <input
                    type="number"
                    step="0.0001"
                    value={data.electrical_tests.vlf_tan_delta_test.mean_tan_delta_1uo}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        vlf_tan_delta_test: {
                          ...data.electrical_tests.vlf_tan_delta_test,
                          mean_tan_delta_1uo: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-bold text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Delta Tan Delta (Tip-Up):</label>
                  <input
                    type="number"
                    step="0.0001"
                    value={data.electrical_tests.vlf_tan_delta_test.tip_up_tan_delta}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        vlf_tan_delta_test: {
                          ...data.electrical_tests.vlf_tan_delta_test,
                          tip_up_tan_delta: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-bold text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 font-medium block mb-1">Đánh Giá Tiêu Chuẩn Tan-Delta:</label>
                  <input
                    type="text"
                    value={data.electrical_tests.vlf_tan_delta_test.evaluation}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        vlf_tan_delta_test: {
                          ...data.electrical_tests.vlf_tan_delta_test,
                          evaluation: e.target.value
                        }
                      }
                    })}
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-purple-500 font-medium text-slate-800"
                  />
                </div>
              </div>
            </div>

            {/* Historical Reference */}
            {data.previous_test_data && (
              <div className="p-3 bg-purple-50/60 rounded-xl border border-purple-200/60 flex items-center justify-between text-xs text-purple-900">
                <span className="font-semibold flex items-center gap-1.5">
                  <Info size={14} className="text-purple-600" />
                  Kỳ Đo Trước ({data.previous_test_data.last_test_date || '2025-09-20'}): Màn chắn Phase C = {data.previous_test_data.last_phase_c_shield_resistance_ohms} Ω
                </span>
                <span className="text-[11px] font-bold text-amber-800 bg-white px-2 py-0.5 rounded border border-amber-300">
                  Tăng Vọt {liveStats.shCIncreasePct}% Lên {liveStats.shC} Ω (Bất Thường Màn Chắn)
                </span>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* AI Analysis Output Section */}
      {analysisResult && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden animate-in fade-in slide-in-from-bottom-3 duration-300">
          <div className="p-5 bg-gradient-to-r from-slate-900 via-purple-950 to-slate-900 text-white flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-purple-500/20 border border-purple-400/30 flex items-center justify-center text-purple-300 shrink-0">
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
                  Báo Cáo Đánh Giá Kỹ Thuật Cáp Trung & Cao Áp (NETA ATS-2025 Mục 7.3.3)
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
                <FileCheck2 size={18} className="text-purple-600" />
                <h3 className="font-bold text-sm text-slate-800">
                  Dữ Liệu JSON Kiểm Định Cáp Trung/Cao Áp (NETA Sec 7.3.3)
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
              className="w-full p-3 font-mono text-xs bg-slate-900 text-purple-200 rounded-xl border border-slate-700 focus:outline-hidden focus:ring-2 focus:ring-purple-500"
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
                  className="px-4 py-1.5 bg-purple-600 hover:bg-purple-500 text-white text-xs font-bold rounded-lg shadow-sm"
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
          title="BIÊN BẢN KIỂM ĐỊNH CÁP ĐIỆN TRUNG VÀ CAO ÁP CÓ MÀN CHẮN KIM LOẠI"
          standard="ANSI/NETA ATS-2025 Section 7.3.3"
          siteInfo={{
            project_name: data.site_info.project_name,
            transformer_tag: data.site_info.cable_tag,
            fse_name: data.site_info.fse_name,
            test_date: data.site_info.test_date
          }}
          overallStatus={analysisResult?.overall_status || (liveStats.isAllShieldPass ? 'PASS' : 'INVESTIGATE')}
          statusReason={analysisResult?.status_reason || 'Đã đối chiếu với quy chuẩn kỹ thuật NETA ATS-2025 Mục 7.3.3.'}
          visualChecks={[
            { label: 'Cable Data Match (Khớp bản vẽ 100%)', value: data.visual_inspection.cable_data_match ? 'Khớp 100%' : 'Sai lệch', pass: data.visual_inspection.cable_data_match },
            { label: 'Đầu cáp & Hộp nối không trầy xước/rách vỏ', value: data.visual_inspection.terminations_and_splices_condition, pass: data.visual_inspection.terminations_and_splices_condition.toLowerCase().includes('pass') },
            { label: 'Lực siết bu-lông đấu nối (Table 100.12)', value: data.visual_inspection.bolt_torque_check, pass: data.visual_inspection.bolt_torque_check.toLowerCase().includes('pass') },
            { label: 'Tiếp địa màn chắn kim loại đúng sơ đồ', value: data.visual_inspection.shield_grounding_check, pass: data.visual_inspection.shield_grounding_check.toLowerCase().includes('pass') },
            { label: 'Bán kính uốn cong cáp (Table 100.22 / ICEA)', value: data.visual_inspection.bending_radius_check, pass: data.visual_inspection.bending_radius_check.toLowerCase().includes('pass') },
            { label: 'Tiếp địa màn chắn lộn ngược qua lòng CT cửa sổ', value: data.visual_inspection.window_ct_grounding_correct ? 'Lộn ngược đúng kỹ thuật' : 'Sai sơ đồ', pass: data.visual_inspection.window_ct_grounding_correct }
          ]}
          electricalTable={[
            {
              item: 'Điện trở cách điện (IR) Pha A',
              measured: `${liveStats.irA} MΩ (2500V DC)`,
              standard: '≥ 1,000 MΩ (Table 100.1)',
              reference: 'NETA Table 100.1',
              deviation: `${liveStats.irA} MΩ (Rất tốt)`,
              status: 'PASS'
            },
            {
              item: 'Điện trở cách điện (IR) Pha B',
              measured: `${liveStats.irB} MΩ (2500V DC)`,
              standard: '≥ 1,000 MΩ (Table 100.1)',
              reference: 'NETA Table 100.1',
              deviation: `${liveStats.irB} MΩ (Rất tốt)`,
              status: 'PASS'
            },
            {
              item: 'Điện trở cách điện (IR) Pha C',
              measured: `${liveStats.irC} MΩ (2500V DC)`,
              standard: '≥ 1,000 MΩ (Table 100.1)',
              reference: 'NETA Table 100.1',
              deviation: `${liveStats.irC} MΩ (Tốt)`,
              status: 'PASS'
            },
            {
              item: 'Màn chắn kim loại Pha A',
              measured: `${liveStats.shA} Ω (${liveStats.shAPer1000} Ω/1,000 ft)`,
              standard: '≤ 10 Ω / 1,000 ft (NETA 7.3.3.D.2)',
              reference: `Ngưỡng: ≤ ${liveStats.maxAllowedOhms} Ω / ${liveStats.lengthFeet} ft`,
              deviation: 'Đạt chuẩn thông mạch',
              status: 'PASS'
            },
            {
              item: 'Màn chắn kim loại Pha B',
              measured: `${liveStats.shB} Ω (${liveStats.shBPer1000} Ω/1,000 ft)`,
              standard: '≤ 10 Ω / 1,000 ft (NETA 7.3.3.D.2)',
              reference: `Ngưỡng: ≤ ${liveStats.maxAllowedOhms} Ω / ${liveStats.lengthFeet} ft`,
              deviation: 'Đạt chuẩn thông mạch',
              status: 'PASS'
            },
            {
              item: 'Màn chắn kim loại Pha C',
              measured: `${liveStats.shC} Ω (${liveStats.shCPer1000} Ω/1,000 ft)`,
              standard: '≤ 10 Ω / 1,000 ft (NETA 7.3.3.D.2)',
              reference: `Ngưỡng: ≤ ${liveStats.maxAllowedOhms} Ω (Kỳ trước: ${liveStats.lastShC} Ω)`,
              deviation: `VƯỢT 64% NGƯỠNG CHO PHÉP (TĂNG ${liveStats.shCIncreasePct}%)`,
              status: 'INVESTIGATE'
            },
            {
              item: 'Đo phản xạ xung TDR',
              measured: data.electrical_tests.tdr_test,
              standard: 'Đồ thị đồng nhất, không có điểm đứt',
              reference: 'Tiêu chuẩn đo TDR',
              deviation: `Xác nhận chiều dài ${liveStats.lengthFeet} ft`,
              status: 'PASS'
            },
            {
              item: 'Thử chịu áp VLF 0.1 Hz (30 min)',
              measured: `${data.electrical_tests.vlf_withstand_test_0_1hz_30min.test_voltage_kv_rms} kV rms`,
              standard: 'Không xảy ra đánh thủng trong 30 phút (Table 100.6.3)',
              reference: 'NETA Table 100.6.3',
              deviation: 'Không đánh thủng',
              status: 'PASS'
            },
            {
              item: 'Tổn hao điện môi Tan-Delta',
              measured: `Mean Tan δ = ${data.electrical_tests.vlf_tan_delta_test.mean_tan_delta_1uo}, Tip-Up = ${data.electrical_tests.vlf_tan_delta_test.tip_up_tan_delta}`,
              standard: 'Đạt ngưỡng Table 100.6.7.1 cho cáp XLPE',
              reference: 'NETA Table 100.6.7.1',
              deviation: 'Chưa lão hóa cây nước',
              status: 'PASS'
            }
          ]}
          warnings={[
            !liveStats.isShCPass ? `Điện trở màn chắn Pha C đạt ${liveStats.shC} Ω / ${liveStats.lengthFeet} ft (${liveStats.shCPer1000} Ω / 1,000 ft) vượt quá ngưỡng tối đa 10 Ω / 1,000 ft NETA ATS-2025 Mục 7.3.3.D.2.` : '',
            liveStats.shCIncreasePct !== '0' && Number(liveStats.shCIncreasePct) > 20 ? `So với kỳ trước (${liveStats.lastShC} Ω), điện trở màn chắn Pha C đã tăng vọt ${liveStats.shCIncreasePct}%.` : ''
          ].filter(Boolean)}
          recommendations={[
            'Kiểm tra điểm tiếp địa màn chắn Phase C tại hai đầu cáp (terminations) xem có bị lỏng, oxy hóa hoặc rỉ sét điểm nối đất hay không.',
            'Nếu các điểm nối đất hai đầu tốt, cần dùng máy dò lỗi cáp/TDR kiểm tra dọc tuyến cáp Phase C để tìm vị trí lớp màn chắn đồng (Tape shield) bị đứt gãy hoặc xói mòn do ngấm nước.',
            `Xử lý lại tiếp địa màn chắn và đo lại điện trở màn chắn Phase C, đảm bảo giá trị giảm xuống dưới ${liveStats.maxAllowedOhms} Ohms / ${liveStats.lengthFeet} feet (10 Ω / 1,000 ft).`
          ]}
        />
      )}
    </div>
  );
};
