import React, { useState, useMemo } from 'react';
import {
  Droplets,
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
  Gauge,
  Flame,
  Zap,
  ShieldCheck,
  TrendingDown,
  ArrowRight,
  Camera
} from 'lucide-react';
import {
  NetaLiquidTransformerInputPayload,
  NetaLiquidTransformerEvaluationResult,
  performDeterministicLiquidTransformerAnalysis
} from '../../server/netaLiquidTransformerAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';
import { NameplateScannerModal } from './NameplateScannerModal';
import { ExtractedNameplateData } from '../../server/nameplateOcrAnalyzer';
import { FseEngineerDropdown } from './FseEngineerDropdown';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

const SAMPLE_LIQUID_TRANSFORMER_DATA: NetaLiquidTransformerInputPayload = {
  site_info: {
    project_name: "Tram Khach Hang TEV Industrial Substation",
    transformer_tag: "TR-OIL-2000KVA-01",
    transformer_type: "Mineral Oil Liquid-Filled Transformer",
    fse_name: "Võ Minh Tú (sgm1707@gmail.com)",
    test_date: "2026-09-23"
  },
  visual_inspection: {
    nameplate_match: true,
    liquid_level_check: "Pass (Normal mark at 25 deg C)",
    gas_blanket_pressure_positive: true,
    valves_and_cooling_fans_check: "Pass",
    bolt_torque_check: "Pass",
    no_oil_leakage: true,
    thermographic_survey: "Pass (Within limits Table 100.18)"
  },
  electrical_tests: {
    bolted_resistance_micro_ohms: [8.1, 8.3, 8.2],
    insulation_resistance_1min_megohms: {
      pri_to_ground_2500v: 3200,
      sec_to_ground_1000v: 1100,
      pri_to_sec_2500v_dc: 4100,
      calculated_pi_value: 1.85
    },
    turns_ratio_all_taps: {
      max_error_percent: 0.12
    },
    winding_insulation_power_factor_percent_20c: 0.35,
    oil_sample_tests_astm_d923: {
      dielectric_breakdown_d1816_1mm_kv: 21,
      water_content_d1533_ppm: 32,
      acid_number_d974_mg_koh_g: 0.02,
      interfacial_tension_d971_mn_m: 39,
      oil_power_factor_25c_d924_percent: 0.04
    },
    dga_test_ieee_c57_104: {
      hydrogen_h2_ppm: 15,
      acetylene_c2h2_ppm: 6.8,
      ethylene_c2h4_ppm: 38,
      methane_ch4_ppm: 12,
      carbon_monoxide_co_ppm: 180,
      carbon_dioxide_co2_ppm: 2400
    }
  },
  previous_test_data: {
    last_test_date: "2025-09-22",
    last_oil_breakdown_kv: 42,
    last_water_content_ppm: 15,
    last_acetylene_ppm: 0.0
  }
};

export const NetaLiquidTransformerChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaLiquidTransformerInputPayload>(SAMPLE_LIQUID_TRANSFORMER_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaLiquidTransformerEvaluationResult | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showCameraModal, setShowCameraModal] = useState(false);
  const [jsonText, setJsonText] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  const handleApplyNameplateData = (extracted: ExtractedNameplateData) => {
    setData((prev) => ({
      ...prev,
      site_info: {
        ...prev.site_info,
        transformer_tag: extracted.equipment_tag || prev.site_info.transformer_tag,
        transformer_type: (extracted.equipment_type as any) || prev.site_info.transformer_type,
      },
      visual_inspection: {
        ...prev.visual_inspection,
        nameplate_match: true,
      }
    }));
    setSavedSuccessMessage(`Đã quét nhãn mác thành công: ${extracted.equipment_tag} (${extracted.manufacturer || 'Máy biến áp'}) - Công suất: ${extracted.rated_power_kva || 2000}kVA, ${extracted.primary_voltage_kv || 22}kV. Đã tự động cập nhật vào Checklist!`);
    setTimeout(() => setSavedSuccessMessage(null), 6000);
  };

  // Live Calculations according to ANSI/NETA ATS-2025 Section 7.2.2
  const liveStats = useMemo(() => {
    const analysis = performDeterministicLiquidTransformerAnalysis(data);
    const oil = data.electrical_tests.oil_sample_tests_astm_d923;
    const dga = data.electrical_tests.dga_test_ieee_c57_104;
    const prev = data.previous_test_data;

    const breakdownFail = oil.dielectric_breakdown_d1816_1mm_kv < 25;
    const waterFail = oil.water_content_d1533_ppm > 20;
    const arcingFail = (dga.acetylene_c2h2_ppm || 0) > 1.0;
    const piFail = data.electrical_tests.insulation_resistance_1min_megohms.calculated_pi_value < 1.0;
    const ratioFail = (data.electrical_tests.turns_ratio_all_taps?.max_error_percent || 0) > 0.5;

    let breakdownDropPct = 0;
    if (prev?.last_oil_breakdown_kv) {
      breakdownDropPct = ((oil.dielectric_breakdown_d1816_1mm_kv - prev.last_oil_breakdown_kv) / prev.last_oil_breakdown_kv) * 100;
    }

    return {
      analysis,
      breakdownFail,
      waterFail,
      arcingFail,
      piFail,
      ratioFail,
      breakdownDropPct: parseFloat(breakdownDropPct.toFixed(1))
    };
  }, [data]);

  const handleRunAiAnalysis = async () => {
    setIsAnalyzing(true);
    try {
      const res = await fetch('/api/field-service/neta-liquid-transformer-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      const json = await res.json();
      if (json.success && json.data) {
        setAnalysisResult(json.data);
      } else {
        alert('Lỗi phân tích: ' + (json.error || 'Vui lòng kiểm tra lại dữ liệu'));
      }
    } catch (e: any) {
      console.error(e);
      // Fallback locally
      const fallback = performDeterministicLiquidTransformerAnalysis(data);
      setAnalysisResult(fallback);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleSaveToCmms = () => {
    const finalReport = analysisResult || performDeterministicLiquidTransformerAnalysis(data);
    const reportPayload = {
      id: `NETA-OIL-TR-${Date.now()}`,
      equipmentId: data.site_info.transformer_tag,
      equipmentName: `Máy biến áp ngâm dầu ${data.site_info.transformer_tag}`,
      type: 'NETA ATS-2025 Liquid-Filled Transformer (Sec 7.2.2)',
      site: data.site_info.project_name,
      inspector: data.site_info.fse_name,
      date: data.site_info.test_date,
      overallStatus: finalReport.overall_status,
      summary: finalReport.status_reason,
      markdownReport: finalReport.markdown_report,
      rawPayload: data,
      createdAt: new Date().toISOString()
    };

    if (onSaveReport) {
      onSaveReport(reportPayload);
      setSavedSuccessMessage(`Đã lưu biên bản ${reportPayload.id} vào kho dữ liệu CMMS thành công!`);
      setTimeout(() => setSavedSuccessMessage(null), 4000);
    }
  };

  const openJsonModal = () => {
    setJsonText(JSON.stringify(data, null, 2));
    setShowJsonModal(true);
  };

  const applyJsonText = () => {
    try {
      const parsed = JSON.parse(jsonText);
      setData(parsed);
      setShowJsonModal(false);
    } catch {
      alert('Chuỗi JSON không đúng định dạng. Vui lòng kiểm tra lại!');
    }
  };

  return (
    <div className="space-y-6">
      {/* Header Banner */}
      <div className="bg-gradient-to-r from-amber-700 via-yellow-800 to-slate-900 text-white p-6 rounded-2xl shadow-lg border border-amber-600/30">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-3">
              <span className="p-2.5 bg-amber-500/20 rounded-xl border border-amber-400/30 text-amber-300">
                <Droplets size={26} />
              </span>
              <div>
                <div className="flex items-center gap-2">
                  <span className="text-xs font-bold uppercase tracking-wider bg-amber-500/30 text-amber-200 px-2 py-0.5 rounded border border-amber-400/40">
                    ANSI/NETA ATS-2025 Mục 7.2.2
                  </span>
                  <span className="text-xs bg-slate-800/80 text-slate-300 px-2 py-0.5 rounded">
                    ASTM D923 & IEEE C57.104
                  </span>
                </div>
                <h1 className="text-xl md:text-2xl font-black mt-1 tracking-tight">
                  Máy biến áp Ngâm Chất lỏng / Ngâm Dầu (Liquid-Filled Transformers)
                </h1>
                <p className="text-xs md:text-sm text-amber-100/80 mt-0.5">
                  Thử nghiệm điện trở cuộn dây, cách điện, điện áp đánh thủng dầu ASTM D1816, hàm lượng ẩm D1533 và phân tích khí hòa tan DGA
                </p>
              </div>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={() => setShowCameraModal(true)}
              className="px-3.5 py-2 bg-amber-400 hover:bg-amber-300 text-slate-950 font-bold rounded-lg text-xs flex items-center gap-1.5 transition-all shadow-sm active:scale-95"
              title="Quét nhãn mác máy biến áp bằng camera và tự động điền thông số"
            >
              <Camera size={14} /> Quét Nhãn Nameplate (AI OCR)
            </button>
            <button
              onClick={() => setData(SAMPLE_LIQUID_TRANSFORMER_DATA)}
              className="px-3 py-2 bg-white/10 hover:bg-white/20 text-white rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-all border border-white/20"
              title="Tải kịch bản mẫu trạm khách hàng TEV Substation"
            >
              <Info size={14} /> Dữ liệu mẫu FSE
            </button>
            <button
              onClick={openJsonModal}
              className="px-3 py-2 bg-white/10 hover:bg-white/20 text-white rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-all border border-white/20"
            >
              <Copy size={14} /> Mã JSON
            </button>
            <button
              onClick={handleRunAiAnalysis}
              disabled={isAnalyzing}
              className="px-4 py-2 bg-gradient-to-r from-amber-500 to-yellow-500 hover:from-amber-400 hover:to-yellow-400 text-slate-950 font-bold rounded-lg text-xs flex items-center gap-1.5 shadow-md transition-all active:scale-95 disabled:opacity-50"
            >
              <Sparkles size={15} />
              {isAnalyzing ? 'Đang phân tích...' : 'Phân tích AI Chuyên gia NETA'}
            </button>
          </div>
        </div>
      </div>

      {savedSuccessMessage && (
        <div className="p-4 bg-emerald-50 border border-emerald-200 text-emerald-800 rounded-xl flex items-center justify-between text-sm animate-fade-in shadow-xs">
          <div className="flex items-center gap-2 font-medium">
            <CheckCircle2 size={18} className="text-emerald-600" />
            {savedSuccessMessage}
          </div>
          {onNavigateToReports && (
            <button
              onClick={onNavigateToReports}
              className="text-emerald-700 underline font-semibold text-xs hover:text-emerald-900"
            >
              Xem danh sách báo cáo
            </button>
          )}
        </div>
      )}

      {/* Live NETA Diagnostic Strip */}
      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3">
        {/* Breakdown Voltage */}
        <div className={`p-3.5 rounded-xl border ${liveStats.breakdownFail ? 'bg-rose-50 border-rose-300 text-rose-900' : 'bg-emerald-50 border-emerald-300 text-emerald-900'} shadow-xs`}>
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold uppercase tracking-wider">Đánh thủng D1816</span>
            <Droplets size={16} className={liveStats.breakdownFail ? 'text-rose-600' : 'text-emerald-600'} />
          </div>
          <div className="text-xl font-black mt-1">
            {data.electrical_tests.oil_sample_tests_astm_d923.dielectric_breakdown_d1816_1mm_kv} kV
          </div>
          <div className="text-[11px] mt-0.5 flex items-center gap-1">
            {liveStats.breakdownFail ? (
              <span className="text-rose-700 font-bold flex items-center gap-0.5">
                <AlertOctagon size={12} /> &lt; 25 kV (Table 100.4.1)
              </span>
            ) : (
              <span className="text-emerald-700 font-semibold">≥ 25 kV Đạt</span>
            )}
          </div>
          {data.previous_test_data?.last_oil_breakdown_kv && (
            <div className="text-[10px] text-slate-500 mt-1 flex items-center gap-0.5">
              <TrendingDown size={11} className="text-rose-500" />
              Năm ngoái: {data.previous_test_data.last_oil_breakdown_kv} kV ({liveStats.breakdownDropPct}%)
            </div>
          )}
        </div>

        {/* Water Content */}
        <div className={`p-3.5 rounded-xl border ${liveStats.waterFail ? 'bg-rose-50 border-rose-300 text-rose-900' : 'bg-emerald-50 border-emerald-300 text-emerald-900'} shadow-xs`}>
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold uppercase tracking-wider">Ẩm trong dầu D1533</span>
            <Activity size={16} className={liveStats.waterFail ? 'text-rose-600' : 'text-emerald-600'} />
          </div>
          <div className="text-xl font-black mt-1">
            {data.electrical_tests.oil_sample_tests_astm_d923.water_content_d1533_ppm} ppm
          </div>
          <div className="text-[11px] mt-0.5">
            {liveStats.waterFail ? (
              <span className="text-rose-700 font-bold flex items-center gap-0.5">
                <AlertTriangle size={12} /> &gt; 20 ppm Nhiễm ẩm
              </span>
            ) : (
              <span className="text-emerald-700 font-semibold">≤ 20 ppm Đạt</span>
            )}
          </div>
        </div>

        {/* DGA C2H2 Arcing */}
        <div className={`p-3.5 rounded-xl border ${liveStats.arcingFail ? 'bg-rose-100 border-rose-400 text-rose-950 animate-pulse-slow' : 'bg-emerald-50 border-emerald-300 text-emerald-900'} shadow-xs`}>
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold uppercase tracking-wider">DGA Acetylene C2H2</span>
            <Flame size={16} className={liveStats.arcingFail ? 'text-rose-600' : 'text-emerald-600'} />
          </div>
          <div className="text-xl font-black mt-1">
            {data.electrical_tests.dga_test_ieee_c57_104.acetylene_c2h2_ppm} ppm
          </div>
          <div className="text-[11px] mt-0.5">
            {liveStats.arcingFail ? (
              <span className="text-rose-800 font-black">
                🔴 HỒ QUANG ĐIỆN
              </span>
            ) : (
              <span className="text-emerald-700 font-semibold">≤ 1.0 ppm Bình thường</span>
            )}
          </div>
        </div>

        {/* PI Calculated */}
        <div className={`p-3.5 rounded-xl border ${liveStats.piFail ? 'bg-rose-50 border-rose-300 text-rose-900' : 'bg-emerald-50 border-emerald-300 text-emerald-900'} shadow-xs`}>
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold uppercase tracking-wider">Chỉ số PI</span>
            <Gauge size={16} className={liveStats.piFail ? 'text-rose-600' : 'text-emerald-600'} />
          </div>
          <div className="text-xl font-black mt-1">
            {data.electrical_tests.insulation_resistance_1min_megohms.calculated_pi_value.toFixed(2)}
          </div>
          <div className="text-[11px] mt-0.5">
            {liveStats.piFail ? (
              <span className="text-rose-700 font-bold">&lt; 1.0 Không đạt</span>
            ) : (
              <span className="text-emerald-700 font-semibold">≥ 1.0 (NETA Sec 7.2.2)</span>
            )}
          </div>
        </div>

        {/* Turns Ratio Error */}
        <div className={`p-3.5 rounded-xl border ${liveStats.ratioFail ? 'bg-rose-50 border-rose-300 text-rose-900' : 'bg-emerald-50 border-emerald-300 text-emerald-900'} shadow-xs`}>
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold uppercase tracking-wider">Sai số Tỷ số Biến áp</span>
            <Layers size={16} className={liveStats.ratioFail ? 'text-rose-600' : 'text-emerald-600'} />
          </div>
          <div className="text-xl font-black mt-1">
            {data.electrical_tests.turns_ratio_all_taps.max_error_percent}%
          </div>
          <div className="text-[11px] mt-0.5">
            {liveStats.ratioFail ? (
              <span className="text-rose-700 font-bold">&gt; 0.5% Không đạt</span>
            ) : (
              <span className="text-emerald-700 font-semibold">≤ 0.5% Đạt nấc tap</span>
            )}
          </div>
        </div>

        {/* Winding PF at 20°C */}
        <div className="p-3.5 rounded-xl border bg-slate-50 border-slate-200 text-slate-900 shadow-xs">
          <div className="flex items-center justify-between">
            <span className="text-xs font-bold uppercase tracking-wider">Winding PF @ 20°C</span>
            <Zap size={16} className="text-amber-500" />
          </div>
          <div className="text-xl font-black mt-1 text-slate-800">
            {data.electrical_tests.winding_insulation_power_factor_percent_20c}%
          </div>
          <div className="text-[11px] mt-0.5 text-slate-600">
            Tiêu chuẩn: ≤ 0.5% (Table 100.3)
          </div>
        </div>
      </div>

      {/* Main Checklist Input Form */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Left Column: Site Info, Visual & Mechanical, Electrical Tests */}
        <div className="space-y-6">
          {/* Site Info */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-xs p-5">
            <div className="flex items-center justify-between mb-4 border-b border-slate-100 pb-2">
              <h3 className="text-sm font-bold text-slate-900 uppercase tracking-wider flex items-center gap-2">
                <span className="w-2 h-2 rounded-full bg-amber-500"></span>
                1. Thông Tin Trạm & Nhãn Máy Biến Áp Dầu
              </h3>
              <button
                type="button"
                onClick={() => setShowCameraModal(true)}
                className="px-2.5 py-1 bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-300 rounded font-bold text-[11px] transition-all flex items-center gap-1 shadow-2xs"
              >
                <Camera size={13} className="text-amber-600" /> Quét Nameplate
              </button>
            </div>

            <div className="flex items-center justify-between mb-3 bg-amber-50/70 border border-amber-200/80 rounded-lg p-2.5 px-3">
              <div className="flex items-center gap-2 text-xs text-amber-950 font-medium">
                <Camera size={15} className="text-amber-600 shrink-0" />
                <span>Dùng Camera quét bảng tên kim loại để AI tự bóc tách và điền form</span>
              </div>
              <button
                type="button"
                onClick={() => setShowCameraModal(true)}
                className="px-3 py-1 bg-amber-600 hover:bg-amber-700 text-white rounded font-bold text-xs shadow-2xs transition-all shrink-0 ml-2"
              >
                Mở Camera
              </button>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Dự án / Trạm biến áp</label>
                <input
                  type="text"
                  value={data.site_info.project_name}
                  onChange={(e) => setData({ ...data, site_info: { ...data.site_info, project_name: e.target.value } })}
                  className="w-full text-xs font-medium border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-amber-500 focus:outline-none"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Mã định danh (Tag Name)</label>
                <input
                  type="text"
                  value={data.site_info.transformer_tag}
                  onChange={(e) => setData({ ...data, site_info: { ...data.site_info, transformer_tag: e.target.value } })}
                  className="w-full text-xs font-medium border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-amber-500 focus:outline-none font-mono"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Loại dầu / Kết cấu MBA</label>
                <select
                  value={data.site_info.transformer_type}
                  onChange={(e) => setData({ ...data, site_info: { ...data.site_info, transformer_type: e.target.value } })}
                  className="w-full text-xs font-medium border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-amber-500 focus:outline-none"
                >
                  <option value="Mineral Oil Liquid-Filled Transformer">Mineral Oil (Dầu khoáng tiêu chuẩn - Table 100.4.1)</option>
                  <option value="Silicone Liquid-Filled Transformer">Silicone Liquid (Table 100.4.2)</option>
                  <option value="Natural Ester Liquid-Filled Transformer">Natural/Synthetic Ester (Table 100.4.3)</option>
                </select>
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">
                  Kỹ sư kiểm tra (FSE TEV)
                </label>
                <FseEngineerDropdown
                  value={data.site_info.fse_name}
                  onChange={(fseName) => setData({ ...data, site_info: { ...data.site_info, fse_name: fseName } })}
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Ngày thử nghiệm</label>
                <input
                  type="date"
                  value={data.site_info.test_date}
                  onChange={(e) => setData({ ...data, site_info: { ...data.site_info, test_date: e.target.value } })}
                  className="w-full text-xs font-medium border border-slate-300 rounded-lg px-2.5 py-2 focus:ring-2 focus:ring-amber-500 focus:outline-none"
                />
              </div>
            </div>
          </div>

          {/* Visual & Mechanical Inspection */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-xs p-5">
            <h3 className="text-sm font-bold text-slate-900 uppercase tracking-wider mb-4 flex items-center gap-2 border-b border-slate-100 pb-2">
              <span className="w-2 h-2 rounded-full bg-amber-500"></span>
              2. Kiểm Tra Thị Giác & Cơ Khí (Section 7.2.2.1)
            </h3>
            <div className="space-y-3">
              <div className="flex items-center justify-between p-3 bg-slate-50 rounded-lg border border-slate-200">
                <span className="text-xs font-semibold text-slate-700">Khớp nhãn mác 100% bản vẽ (Nameplate Match)</span>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.nameplate_match}
                  onChange={(e) => setData({ ...data, visual_inspection: { ...data.visual_inspection, nameplate_match: e.target.checked } })}
                  className="w-4 h-4 text-amber-600 rounded focus:ring-amber-500"
                />
              </div>

              <div className="flex items-center justify-between p-3 bg-slate-50 rounded-lg border border-slate-200">
                <div>
                  <div className="text-xs font-semibold text-slate-700">Áp suất đệm khí Nitơ dương (Positive Gas Pressure)</div>
                  <div className="text-[11px] text-slate-500">BẮT BUỘC duy trì áp suất dương (0.5 - 5.0 psi) chống lọt ẩm</div>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.gas_blanket_pressure_positive}
                  onChange={(e) => setData({ ...data, visual_inspection: { ...data.visual_inspection, gas_blanket_pressure_positive: e.target.checked } })}
                  className="w-4 h-4 text-amber-600 rounded focus:ring-amber-500"
                />
              </div>

              <div className="flex items-center justify-between p-3 bg-slate-50 rounded-lg border border-slate-200">
                <span className="text-xs font-semibold text-slate-700">Không rò rỉ dầu ở cánh tản nhiệt & van xả (No Oil Leakage)</span>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.no_oil_leakage}
                  onChange={(e) => setData({ ...data, visual_inspection: { ...data.visual_inspection, no_oil_leakage: e.target.checked } })}
                  className="w-4 h-4 text-amber-600 rounded focus:ring-amber-500"
                />
              </div>

              <div className="grid grid-cols-2 gap-3 pt-1">
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1">Mức dầu tank chính & LTC</label>
                  <input
                    type="text"
                    value={data.visual_inspection.liquid_level_check}
                    onChange={(e) => setData({ ...data, visual_inspection: { ...data.visual_inspection, liquid_level_check: e.target.value } })}
                    className="w-full text-xs font-medium border border-slate-300 rounded-lg px-3 py-2"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1">Lực siết bu-lông (Table 100.12)</label>
                  <input
                    type="text"
                    value={data.visual_inspection.bolt_torque_check}
                    onChange={(e) => setData({ ...data, visual_inspection: { ...data.visual_inspection, bolt_torque_check: e.target.value } })}
                    className="w-full text-xs font-medium border border-slate-300 rounded-lg px-3 py-2"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Electrical Tests: Bolted & IR / PI */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-xs p-5">
            <h3 className="text-sm font-bold text-slate-900 uppercase tracking-wider mb-4 flex items-center gap-2 border-b border-slate-100 pb-2">
              <span className="w-2 h-2 rounded-full bg-amber-500"></span>
              3. Thử Nghiệm Điện Trở Cuộn Dây & Cách Điện (Section 7.2.2.2 - 7.2.2.4)
            </h3>
            
            <div className="space-y-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Điện trở mối nối bu-lông 3 pha (µΩ) - Table 100.12 (Độ lệch ≤ 50%)
                </label>
                <div className="grid grid-cols-3 gap-2">
                  {data.electrical_tests.bolted_resistance_micro_ohms.map((val, idx) => (
                    <input
                      key={idx}
                      type="number"
                      step="0.1"
                      value={val}
                      onChange={(e) => {
                        const newBolted = [...data.electrical_tests.bolted_resistance_micro_ohms];
                        newBolted[idx] = parseFloat(e.target.value) || 0;
                        setData({
                          ...data,
                          electrical_tests: { ...data.electrical_tests, bolted_resistance_micro_ohms: newBolted }
                        });
                      }}
                      className="text-xs font-semibold border border-slate-300 rounded-lg px-2.5 py-1.5 text-center font-mono"
                    />
                  ))}
                </div>
              </div>

              <div className="p-3 bg-slate-50 rounded-lg border border-slate-200 space-y-2.5">
                <div className="text-xs font-bold text-slate-800 flex items-center justify-between">
                  <span>Điện trở cách điện IR (1 min) & Chỉ số PI (Table 100.5)</span>
                  <span className="text-[11px] font-semibold text-amber-700">PI BẮT BUỘC ≥ 1.0</span>
                </div>
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                  <div>
                    <label className="block text-[10px] font-medium text-slate-500 mb-0.5">Sơ cấp-Đất (2.5kV)</label>
                    <input
                      type="number"
                      value={data.electrical_tests.insulation_resistance_1min_megohms.pri_to_ground_2500v}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1min_megohms: {
                            ...data.electrical_tests.insulation_resistance_1min_megohms,
                            pri_to_ground_2500v: parseFloat(e.target.value) || 0
                          }
                        }
                      })}
                      className="w-full text-xs font-bold border border-slate-300 rounded px-2 py-1 font-mono"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-medium text-slate-500 mb-0.5">Thứ cấp-Đất (1kV)</label>
                    <input
                      type="number"
                      value={data.electrical_tests.insulation_resistance_1min_megohms.sec_to_ground_1000v}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1min_megohms: {
                            ...data.electrical_tests.insulation_resistance_1min_megohms,
                            sec_to_ground_1000v: parseFloat(e.target.value) || 0
                          }
                        }
                      })}
                      className="w-full text-xs font-bold border border-slate-300 rounded px-2 py-1 font-mono"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-medium text-slate-500 mb-0.5">Sơ-Thứ (2.5kV DC)</label>
                    <input
                      type="number"
                      value={data.electrical_tests.insulation_resistance_1min_megohms.pri_to_sec_2500v_dc}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1min_megohms: {
                            ...data.electrical_tests.insulation_resistance_1min_megohms,
                            pri_to_sec_2500v_dc: parseFloat(e.target.value) || 0
                          }
                        }
                      })}
                      className="w-full text-xs font-bold border border-slate-300 rounded px-2 py-1 font-mono"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-medium text-slate-500 mb-0.5">Chỉ số PI (10m/1m)</label>
                    <input
                      type="number"
                      step="0.05"
                      value={data.electrical_tests.insulation_resistance_1min_megohms.calculated_pi_value}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1min_megohms: {
                            ...data.electrical_tests.insulation_resistance_1min_megohms,
                            calculated_pi_value: parseFloat(e.target.value) || 0
                          }
                        }
                      })}
                      className="w-full text-xs font-black border border-slate-300 rounded px-2 py-1 font-mono text-amber-700 bg-amber-50"
                    />
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Sai số Tỷ số Biến áp (%) - max ≤ 0.5%
                  </label>
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.turns_ratio_all_taps.max_error_percent}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        turns_ratio_all_taps: { max_error_percent: parseFloat(e.target.value) || 0 }
                      }
                    })}
                    className="w-full text-xs font-bold border border-slate-300 rounded-lg px-3 py-2 font-mono"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Winding Power Factor @ 20°C (%) ≤ 0.5%
                  </label>
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.winding_insulation_power_factor_percent_20c}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        winding_insulation_power_factor_percent_20c: parseFloat(e.target.value) || 0
                      }
                    })}
                    className="w-full text-xs font-bold border border-slate-300 rounded-lg px-3 py-2 font-mono"
                  />
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Right Column: Oil Sample Tests ASTM D923 & DGA IEEE C57.104 */}
        <div className="space-y-6">
          {/* Oil Sample Tests */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-xs p-5">
            <h3 className="text-sm font-bold text-slate-900 uppercase tracking-wider mb-4 flex items-center justify-between border-b border-slate-100 pb-2">
              <span className="flex items-center gap-2">
                <span className="w-2 h-2 rounded-full bg-amber-500"></span>
                4. Chỉ Tiêu Dầu Cách Điện (ASTM D923 & Table 100.4.1)
              </span>
              <span className="text-[11px] font-bold text-amber-700 bg-amber-50 px-2 py-0.5 rounded border border-amber-200">
                Dầu khoáng Mineral Oil
              </span>
            </h3>

            <div className="space-y-3.5">
              <div className="grid grid-cols-2 gap-3">
                <div className={`p-3 rounded-lg border ${liveStats.breakdownFail ? 'bg-rose-50 border-rose-300' : 'bg-slate-50 border-slate-200'}`}>
                  <label className="block text-xs font-bold text-slate-800 mb-1">
                    Điện áp đánh thủng D1816 1mm (kV)
                  </label>
                  <input
                    type="number"
                    value={data.electrical_tests.oil_sample_tests_astm_d923.dielectric_breakdown_d1816_1mm_kv}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        oil_sample_tests_astm_d923: {
                          ...data.electrical_tests.oil_sample_tests_astm_d923,
                          dielectric_breakdown_d1816_1mm_kv: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full text-base font-black border border-slate-300 rounded-md px-3 py-1.5 font-mono text-slate-900 bg-white"
                  />
                  <div className="text-[11px] text-slate-500 mt-1">Ngưỡng Table 100.4.1: ≥ 25 kV</div>
                </div>

                <div className={`p-3 rounded-lg border ${liveStats.waterFail ? 'bg-rose-50 border-rose-300' : 'bg-slate-50 border-slate-200'}`}>
                  <label className="block text-xs font-bold text-slate-800 mb-1">
                    Hàm lượng ẩm D1533 (ppm)
                  </label>
                  <input
                    type="number"
                    value={data.electrical_tests.oil_sample_tests_astm_d923.water_content_d1533_ppm}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        oil_sample_tests_astm_d923: {
                          ...data.electrical_tests.oil_sample_tests_astm_d923,
                          water_content_d1533_ppm: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full text-base font-black border border-slate-300 rounded-md px-3 py-1.5 font-mono text-slate-900 bg-white"
                  />
                  <div className="text-[11px] text-slate-500 mt-1">Ngưỡng Table 100.4.1: ≤ 20 ppm</div>
                </div>
              </div>

              <div className="grid grid-cols-3 gap-3">
                <div>
                  <label className="block text-[11px] font-semibold text-slate-600 mb-1">Trị số Axit (mg KOH/g)</label>
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.oil_sample_tests_astm_d923.acid_number_d974_mg_koh_g}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        oil_sample_tests_astm_d923: {
                          ...data.electrical_tests.oil_sample_tests_astm_d923,
                          acid_number_d974_mg_koh_g: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full text-xs font-bold border border-slate-300 rounded px-2.5 py-1.5 font-mono"
                  />
                  <span className="text-[10px] text-slate-400">≤ 0.05 mg KOH/g</span>
                </div>
                <div>
                  <label className="block text-[11px] font-semibold text-slate-600 mb-1">Sức căng IFT (mN/m)</label>
                  <input
                    type="number"
                    step="1"
                    value={data.electrical_tests.oil_sample_tests_astm_d923.interfacial_tension_d971_mn_m}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        oil_sample_tests_astm_d923: {
                          ...data.electrical_tests.oil_sample_tests_astm_d923,
                          interfacial_tension_d971_mn_m: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full text-xs font-bold border border-slate-300 rounded px-2.5 py-1.5 font-mono"
                  />
                  <span className="text-[10px] text-slate-400">≥ 30 mN/m</span>
                </div>
                <div>
                  <label className="block text-[11px] font-semibold text-slate-600 mb-1">Power Factor dầu (%)</label>
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.oil_sample_tests_astm_d923.oil_power_factor_25c_d924_percent}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        oil_sample_tests_astm_d923: {
                          ...data.electrical_tests.oil_sample_tests_astm_d923,
                          oil_power_factor_25c_d924_percent: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full text-xs font-bold border border-slate-300 rounded px-2.5 py-1.5 font-mono"
                  />
                  <span className="text-[10px] text-slate-400">≤ 0.1% @ 25°C</span>
                </div>
              </div>
            </div>
          </div>

          {/* DGA Dissolved Gas Analysis IEEE C57.104 */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-xs p-5">
            <h3 className="text-sm font-bold text-slate-900 uppercase tracking-wider mb-4 flex items-center justify-between border-b border-slate-100 pb-2">
              <span className="flex items-center gap-2">
                <span className="w-2 h-2 rounded-full bg-amber-500"></span>
                5. Phân Tích Khí Hòa Tan DGA (IEEE C57.104)
              </span>
              <span className="text-[11px] font-bold text-rose-700 bg-rose-50 px-2 py-0.5 rounded border border-rose-200">
                Phát hiện Hồ quang & Quá nhiệt
              </span>
            </h3>

            <div className="space-y-3">
              <div className={`p-3 rounded-lg border ${liveStats.arcingFail ? 'bg-rose-50 border-rose-300' : 'bg-slate-50 border-slate-200'}`}>
                <div className="flex items-center justify-between mb-1">
                  <label className="text-xs font-black text-rose-950 flex items-center gap-1.5">
                    <Flame size={14} className="text-rose-600" />
                    Acetylene C2H2 (ppm) - CẢNH BÁO HỒ QUANG NĂNG LƯỢNG CAO
                  </label>
                  <span className="text-[10px] font-bold uppercase px-1.5 py-0.5 bg-rose-200 text-rose-800 rounded">
                    Limit: ≤ 1.0 ppm
                  </span>
                </div>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.dga_test_ieee_c57_104.acetylene_c2h2_ppm}
                  onChange={(e) => setData({
                    ...data,
                    electrical_tests: {
                      ...data.electrical_tests,
                      dga_test_ieee_c57_104: {
                        ...data.electrical_tests.dga_test_ieee_c57_104,
                        acetylene_c2h2_ppm: parseFloat(e.target.value) || 0
                      }
                    }
                  })}
                  className="w-full text-base font-black border border-rose-300 rounded-md px-3 py-1.5 font-mono text-rose-900 bg-white"
                />
                <div className="text-[11px] text-rose-700 mt-1">
                  {data.electrical_tests.dga_test_ieee_c57_104.acetylene_c2h2_ppm > 1.0
                    ? '⚠️ CẢNH BÁO TỐI CẤP: Xuất hiện Acetylene xác nhận có phóng hồ quang > 700°C bên trong máy biến áp!'
                    : 'Bình thường không có khí Acetylene.'}
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Hydrogen H2 (ppm) - limit ≤ 100
                  </label>
                  <input
                    type="number"
                    value={data.electrical_tests.dga_test_ieee_c57_104.hydrogen_h2_ppm}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        dga_test_ieee_c57_104: {
                          ...data.electrical_tests.dga_test_ieee_c57_104,
                          hydrogen_h2_ppm: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full text-xs font-bold border border-slate-300 rounded-lg px-3 py-1.5 font-mono"
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Ethylene C2H4 (ppm) - limit ≤ 50
                  </label>
                  <input
                    type="number"
                    value={data.electrical_tests.dga_test_ieee_c57_104.ethylene_c2h4_ppm}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        dga_test_ieee_c57_104: {
                          ...data.electrical_tests.dga_test_ieee_c57_104,
                          ethylene_c2h4_ppm: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full text-xs font-bold border border-slate-300 rounded-lg px-3 py-1.5 font-mono"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Previous Year Trend Comparison */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-xs p-5">
            <h3 className="text-sm font-bold text-slate-900 uppercase tracking-wider mb-4 flex items-center gap-2 border-b border-slate-100 pb-2">
              <span className="w-2 h-2 rounded-full bg-amber-500"></span>
              6. So Sánh Xu Hướng Lịch Sử (Trending Data)
            </h3>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Đánh thủng năm trước (kV)</label>
                <input
                  type="number"
                  value={data.previous_test_data?.last_oil_breakdown_kv || 42}
                  onChange={(e) => setData({
                    ...data,
                    previous_test_data: {
                      ...data.previous_test_data,
                      last_oil_breakdown_kv: parseFloat(e.target.value) || 0
                    }
                  })}
                  className="w-full text-xs font-bold border border-slate-300 rounded-lg px-3 py-1.5 font-mono"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Hàm lượng nước năm trước (ppm)</label>
                <input
                  type="number"
                  value={data.previous_test_data?.last_water_content_ppm || 15}
                  onChange={(e) => setData({
                    ...data,
                    previous_test_data: {
                      ...data.previous_test_data,
                      last_water_content_ppm: parseFloat(e.target.value) || 0
                    }
                  })}
                  className="w-full text-xs font-bold border border-slate-300 rounded-lg px-3 py-1.5 font-mono"
                />
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Action Toolbar */}
      <div className="flex flex-wrap items-center justify-between gap-3 p-4 bg-white rounded-xl border border-slate-200 shadow-xs">
        <div className="flex items-center gap-2">
          <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">Trạng thái ước lượng:</span>
          {liveStats.arcingFail || liveStats.breakdownFail || liveStats.waterFail ? (
            <span className="px-2.5 py-1 bg-rose-100 text-rose-800 font-black rounded-lg text-xs flex items-center gap-1 border border-rose-300">
              <AlertOctagon size={13} /> 🔴 FAIL / CRITICAL FLUID & DGA DEFECT
            </span>
          ) : (
            <span className="px-2.5 py-1 bg-emerald-100 text-emerald-800 font-bold rounded-lg text-xs flex items-center gap-1 border border-emerald-300">
              <CheckCircle2 size={13} /> 🟢 PASS NETA ATS-2025
            </span>
          )}
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={() => setShowPrintModal(true)}
            className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 rounded-lg text-xs font-bold flex items-center gap-1.5 transition-all"
          >
            <Printer size={15} /> In Biên Bản PDF
          </button>
          <button
            onClick={handleSaveToCmms}
            className="px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 transition-all shadow-xs"
          >
            <Save size={15} /> Lưu vào Kho Báo Cáo CMMS
          </button>
          <button
            onClick={handleRunAiAnalysis}
            disabled={isAnalyzing}
            className="px-5 py-2 bg-gradient-to-r from-amber-600 to-yellow-600 hover:from-amber-500 hover:to-yellow-500 text-slate-950 rounded-lg text-xs font-black flex items-center gap-1.5 shadow-md transition-all active:scale-95 disabled:opacity-50"
          >
            <Sparkles size={15} />
            {isAnalyzing ? 'Đang thẩm định AI...' : 'Xuất Báo Cáo Chuyên Gia NETA'}
          </button>
        </div>
      </div>

      {/* Analysis Result Output Display */}
      {analysisResult && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-md p-6 space-y-6 animate-fade-in">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-4">
            <div>
              <div className="flex items-center gap-2">
                <span className="text-xs font-black uppercase tracking-wider bg-slate-900 text-amber-400 px-2.5 py-0.5 rounded">
                  BÁO CÁO THẨM ĐỊNH AI NETA ATS-2025
                </span>
                <span className="text-xs text-slate-500">Mục 7.2.2 & Table 100.4.1</span>
              </div>
              <h2 className="text-xl font-black text-slate-900 mt-1">
                Kết Quả Thử Nghiệm Máy Biến Áp Ngâm Dầu ({data.site_info.transformer_tag})
              </h2>
            </div>
            <div>
              {analysisResult.overall_status === 'FAIL' ? (
                <div className="px-4 py-2 bg-rose-100 text-rose-800 font-black rounded-xl border border-rose-300 text-center">
                  <div className="text-xs">KẾT LUẬN TỔNG QUAN</div>
                  <div className="text-base flex items-center justify-center gap-1">
                    <AlertOctagon size={18} /> FAIL / CRITICAL FLUID & DGA DEFECT
                  </div>
                </div>
              ) : analysisResult.overall_status === 'INVESTIGATE' ? (
                <div className="px-4 py-2 bg-amber-100 text-amber-800 font-black rounded-xl border border-amber-300 text-center">
                  <div className="text-xs">KẾT LUẬN TỔNG QUAN</div>
                  <div className="text-base flex items-center justify-center gap-1">
                    <AlertTriangle size={18} /> INVESTIGATE / CẦN KHẢO SÁT
                  </div>
                </div>
              ) : (
                <div className="px-4 py-2 bg-emerald-100 text-emerald-800 font-black rounded-xl border border-emerald-300 text-center">
                  <div className="text-xs">KẾT LUẬN TỔNG QUAN</div>
                  <div className="text-base flex items-center justify-center gap-1">
                    <CheckCircle2 size={18} /> PASS / ĐẠT TIÊU CHUẨN
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* Markdown Content */}
          <div className="prose prose-slate max-w-none text-slate-800 text-sm leading-relaxed bg-slate-50/50 p-6 rounded-xl border border-slate-200">
            <div
              className="space-y-4"
              dangerouslySetInnerHTML={{
                __html: analysisResult.markdown_report
                  .replace(/# (.*?)\n/g, '<h1 class="text-xl font-bold text-slate-900 border-b pb-2 mt-4">$1</h1>')
                  .replace(/## (.*?)\n/g, '<h2 class="text-lg font-bold text-slate-800 mt-4 mb-2 flex items-center gap-2"><span class="w-1.5 h-4 bg-amber-600 rounded"></span>$1</h2>')
                  .replace(/### (.*?)\n/g, '<h3 class="text-base font-bold text-slate-800 mt-3 mb-1">$1</h3>')
                  .replace(/\*\*(.*?)\*\*/g, '<strong class="font-bold text-slate-900">$1</strong>')
                  .replace(/`([^`]+)`/g, '<code class="bg-slate-200 px-1 py-0.5 rounded text-amber-800 font-mono text-xs">$1</code>')
                  .replace(/\n\n/g, '<p class="my-2"></p>')
                  .replace(/\| (.*?) \|/g, (match) => {
                    return `<div class="overflow-x-auto my-2 text-xs font-mono">${match}</div>`;
                  })
              }}
            />
          </div>
        </div>
      )}

      {/* JSON Import/Export Modal */}
      {showJsonModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 p-4 backdrop-blur-xs">
          <div className="bg-white rounded-2xl max-w-2xl w-full p-6 shadow-2xl space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h3 className="text-base font-bold text-slate-800 flex items-center gap-2">
                <Copy size={18} className="text-amber-600" />
                Mã Nguồn JSON Dữ Liệu Khảo Sát FSE
              </h3>
              <button
                onClick={() => setShowJsonModal(false)}
                className="text-slate-400 hover:text-slate-600 text-lg font-bold"
              >
                ✕
              </button>
            </div>
            <p className="text-xs text-slate-500">
              Bạn có thể dán dữ liệu kiểm tra máy biến áp dầu từ máy đo hoặc sao chép mã JSON chuẩn để tích hợp API.
            </p>
            <textarea
              rows={14}
              value={jsonText}
              onChange={(e) => setJsonText(e.target.value)}
              className="w-full text-xs font-mono p-3 bg-slate-50 border border-slate-300 rounded-xl focus:outline-none focus:ring-2 focus:ring-amber-500"
            />
            <div className="flex items-center justify-between">
              <button
                onClick={() => {
                  navigator.clipboard.writeText(jsonText);
                  setCopied(true);
                  setTimeout(() => setCopied(false), 2000);
                }}
                className="px-3.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-all"
              >
                {copied ? <Check size={14} className="text-emerald-600" /> : <Copy size={14} />}
                {copied ? 'Đã sao chép!' : 'Sao chép JSON'}
              </button>
              <div className="flex gap-2">
                <button
                  onClick={() => setShowJsonModal(false)}
                  className="px-4 py-1.5 border border-slate-300 text-slate-700 rounded-lg text-xs font-semibold hover:bg-slate-50"
                >
                  Đóng
                </button>
                <button
                  onClick={applyJsonText}
                  className="px-4 py-1.5 bg-amber-600 text-white rounded-lg text-xs font-bold hover:bg-amber-700 shadow-xs"
                >
                  Nạp Dữ Liệu Này
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* PDF Print Modal */}
      {showPrintModal && (
        <TevReportPrintModal
          report={{
            id: `NETA-OIL-TR-${data.site_info.transformer_tag}`,
            equipmentId: data.site_info.transformer_tag,
            equipmentName: `Máy biến áp ngâm dầu ${data.site_info.transformer_tag}`,
            substation: data.site_info.project_name,
            location: 'Trạm biến áp ngoài trời / Trong nhà',
            voltageLevel: '22kV / 0.4kV (2000kVA)',
            testingEngineer: data.site_info.fse_name,
            overallStatus: analysisResult?.overall_status || (liveStats.arcingFail || liveStats.breakdownFail ? 'FAIL' : 'PASS'),
            testDate: data.site_info.test_date,
            summary: analysisResult?.status_reason || 'Đánh giá đối soát NETA ATS-2025 Mục 7.2.2 cho máy biến áp ngâm dầu.',
            detailedFindings: analysisResult?.markdown_report || performDeterministicLiquidTransformerAnalysis(data).markdown_report,
            recommendations: 'Không cho đóng điện. Cô lập máy biến áp, lọc dầu hút chân không khử ẩm và kiểm tra nội soi vết hồ quang Acetylene.',
            sensorData: [
              { name: 'Điện áp đánh thủng D1816', value: `${data.electrical_tests.oil_sample_tests_astm_d923.dielectric_breakdown_d1816_1mm_kv} kV`, threshold: '≥ 25 kV (Table 100.4.1)', status: liveStats.breakdownFail ? 'FAIL' : 'PASS' },
              { name: 'Hàm lượng ẩm D1533', value: `${data.electrical_tests.oil_sample_tests_astm_d923.water_content_d1533_ppm} ppm`, threshold: '≤ 20 ppm', status: liveStats.waterFail ? 'FAIL' : 'PASS' },
              { name: 'DGA Acetylene C2H2', value: `${data.electrical_tests.dga_test_ieee_c57_104.acetylene_c2h2_ppm} ppm`, threshold: '≤ 1.0 ppm (IEEE C57.104)', status: liveStats.arcingFail ? 'CRITICAL' : 'PASS' },
              { name: 'Chỉ số phân cực PI', value: `${data.electrical_tests.insulation_resistance_1min_megohms.calculated_pi_value.toFixed(2)}`, threshold: '≥ 1.0 (NETA Sec 7.2.2)', status: liveStats.piFail ? 'FAIL' : 'PASS' },
              { name: 'Sai số Tỷ số Biến áp', value: `${data.electrical_tests.turns_ratio_all_taps.max_error_percent}%`, threshold: '≤ 0.5%', status: liveStats.ratioFail ? 'FAIL' : 'PASS' },
              { name: 'Winding PF @ 20°C', value: `${data.electrical_tests.winding_insulation_power_factor_percent_20c}%`, threshold: '≤ 0.5% (Table 100.3)', status: 'PASS' }
            ]
          }}
          onClose={() => setShowPrintModal(false)}
        />
      )}

      {/* Camera Nameplate OCR Modal */}
      {showCameraModal && (
        <NameplateScannerModal
          isOpen={showCameraModal}
          onClose={() => setShowCameraModal(false)}
          onApplyData={handleApplyNameplateData}
          currentTag={data.site_info.transformer_tag}
        />
      )}
    </div>
  );
};
