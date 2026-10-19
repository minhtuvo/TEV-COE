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
  Camera,
  Globe,
  RefreshCw,
  FileText,
  Sliders,
  Share2,
  Database,
  Calendar,
  Wrench
} from 'lucide-react';
import {
  NetaMtsTransformerInputPayload,
  NetaMtsTransformerEvaluationResult,
  performDeterministicMtsTransformerAnalysis,
  buildDeterministicMtsTransformerMarkdown
} from '../../server/netaMtsTransformerAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';
import { NameplateScannerModal } from './NameplateScannerModal';
import { ExtractedNameplateData } from '../../server/nameplateOcrAnalyzer';
import { FseEngineerDropdown } from './FseEngineerDropdown';
import { useAtsFormSync } from '../hooks/useAtsFormSync';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
  initialLanguage?: 'vi' | 'en';
  initialDisplayMode?: 'form' | 'tev-report';
}

export const SAMPLE_MTS_PROMPT_DATA: NetaMtsTransformerInputPayload = {
  site_info: {
    project_name: "Nha May Hoa Chat TEV Chemical Plant",
    transformer_tag: "TR-MAIN-110KV-01",
    transformer_rating: "25MVA 110/22kV Class I Mineral Oil Transformer",
    fse_name: "Nguyen Van Z",
    test_date: "2026-10-12",
    voltage_class_kv: 110,
    oil_type: "Mineral Oil Class I"
  },
  visual_inspection: {
    nameplate_match: true,
    liquid_level_check: "Pass (Normal mark at 25 deg C)",
    gas_blanket_pressure_positive: true,
    valves_and_cooling_fans_check: "Pass",
    bolt_torque_check: "Pass",
    no_oil_leakage: true,
    thermographic_survey: "Pass (Compliant with Section 9 / Table 100.18)"
  },
  electrical_tests: {
    bolted_resistance_micro_ohms: [7.8, 8.0, 7.9],
    insulation_resistance_20c_megohms: {
      pri_to_ground_5000v: 8200,
      sec_to_ground_2500v: 3500,
      pri_to_sec_5000v_dc: 9100,
      calculated_pi_value: 1.45
    },
    turns_ratio_all_taps: {
      max_error_percent: 0.18
    },
    winding_insulation_power_factor_percent_20c: 0.85,
    winding_dc_resistance_ohms: {
      max_phase_deviation_percent: 1.2
    },
    core_insulation_resistance_500v_megohms: 350,
    service_aged_oil_sample_tests_table_100_4_1: {
      dielectric_breakdown_d1816_1mm_kv: 19.5,
      dielectric_breakdown_d877_kv: 24.0,
      water_content_d1533_ppm: 42,
      acid_number_d974_mg_koh_g: 0.28,
      interfacial_tension_d971_mn_m: 21,
      oil_power_factor_25c_d924_percent: 1.4
    },
    dga_test_ieee_c57_104: {
      hydrogen_h2_ppm: 45,
      methane_ch4_ppm: 28,
      acetylene_c2h2_ppm: 8.5,
      ethylene_c2h4_ppm: 115,
      ethane_c2h6_ppm: 18,
      carbon_monoxide_co_ppm: 320,
      carbon_dioxide_co2_ppm: 2800,
      total_dissolved_combustible_gas_ppm: 206.5
    }
  },
  previous_test_data: {
    last_test_date: "2024-10-10",
    last_oil_breakdown_kv: 31.0,
    last_water_content_ppm: 22,
    last_c2h2_ppm: 0.0
  },
  language: 'vi'
};

export const NetaMtsTransformerChecklist: React.FC<Props> = ({
  onSaveReport,
  onNavigateToReports,
  initialLanguage = 'vi',
  initialDisplayMode = 'form'
}) => {
  const [lang, setLang] = useState<'vi' | 'en'>(initialLanguage);
  const [mtsViewMode, setMtsViewMode] = useState<'form' | 'tev-report'>(initialDisplayMode);
  const [data, setData] = useAtsFormSync<NetaMtsTransformerInputPayload>('neta-mts-transformer', SAMPLE_MTS_PROMPT_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaMtsTransformerEvaluationResult | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showCameraModal, setShowCameraModal] = useState(false);
  const [jsonText, setJsonText] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);
  const [activeFormTab, setActiveFormTab] = useState<'overview' | 'visual' | 'electrical' | 'oil' | 'dga'>('overview');

  React.useEffect(() => {
    setMtsViewMode(initialDisplayMode);
  }, [initialDisplayMode]);

  // Whenever user modifies data in the form, automatically reset analysisResult so currentEvaluation re-evaluates fresh
  React.useEffect(() => {
    setAnalysisResult(null);
  }, [data]);

  React.useEffect(() => {
    const handleResetMode = () => setMtsViewMode('form');
    window.addEventListener('tev_mts_open_form', handleResetMode);
    return () => window.removeEventListener('tev_mts_open_form', handleResetMode);
  }, []);

  // Compute live deterministic baseline
  const liveStats = useMemo(() => {
    const analysis = performDeterministicMtsTransformerAnalysis({ ...data, language: lang });
    const oil = data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1;
    const dga = data.electrical_tests.dga_test_ieee_c57_104;
    const prev = data.previous_test_data;

    const is110kV = (data.site_info.transformer_rating || '').includes('110') || (data.site_info.voltage_class_kv && data.site_info.voltage_class_kv >= 69);
    const breakdownThreshold = is110kV ? 28 : 23;
    const waterThreshold = is110kV ? 25 : 35;

    const breakdownFail = oil.dielectric_breakdown_d1816_1mm_kv < breakdownThreshold;
    const waterFail = oil.water_content_d1533_ppm > waterThreshold;
    const acidFail = oil.acid_number_d974_mg_koh_g > 0.20;
    const iftFail = oil.interfacial_tension_d971_mn_m < 24;
    const oilPfFail = oil.oil_power_factor_25c_d924_percent > 1.0;
    const arcingFail = (dga.acetylene_c2h2_ppm || 0) > 1.0 || ((prev?.last_c2h2_ppm === 0 || !prev?.last_c2h2_ppm) && (dga.acetylene_c2h2_ppm || 0) > 0);
    const piFail = data.electrical_tests.insulation_resistance_20c_megohms.calculated_pi_value < 1.0;
    const ratioFail = (data.electrical_tests.turns_ratio_all_taps?.max_error_percent || 0) > 0.5;
    const windingPfFail = data.electrical_tests.winding_insulation_power_factor_percent_20c > 1.0;

    let breakdownDropPct = 0;
    if (prev?.last_oil_breakdown_kv) {
      breakdownDropPct = ((oil.dielectric_breakdown_d1816_1mm_kv - prev.last_oil_breakdown_kv) / prev.last_oil_breakdown_kv) * 100;
    }

    return {
      analysis,
      breakdownFail,
      breakdownThreshold,
      waterFail,
      waterThreshold,
      acidFail,
      iftFail,
      oilPfFail,
      arcingFail,
      piFail,
      ratioFail,
      windingPfFail,
      breakdownDropPct,
      overallSafe: !breakdownFail && !waterFail && !acidFail && !iftFail && !arcingFail && !piFail && !ratioFail && !windingPfFail
    };
  }, [data, lang]);

  const handleApplyNameplateData = (extracted: ExtractedNameplateData) => {
    setData((prev) => ({
      ...prev,
      site_info: {
        ...prev.site_info,
        transformer_tag: extracted.equipment_tag || prev.site_info.transformer_tag,
        transformer_rating: extracted.rated_power_kva
          ? `${extracted.rated_power_kva}kVA ${extracted.primary_voltage_kv || 110}/${(extracted as any).secondary_voltage_kv || (extracted.secondary_voltage_v ? extracted.secondary_voltage_v / 1000 : 22)}kV Mineral Oil Transformer`
          : prev.site_info.transformer_rating,
        voltage_class_kv: extracted.primary_voltage_kv || prev.site_info.voltage_class_kv,
      },
      visual_inspection: {
        ...prev.visual_inspection,
        nameplate_match: true,
      }
    }));
    setSavedSuccessMessage(
      lang === 'vi'
        ? `Đã quét nhãn thành công: ${extracted.equipment_tag || 'MBA'} - Công suất: ${extracted.rated_power_kva || 25000}kVA. Đã tự động cập nhật vào biên bản MTS!`
        : `Nameplate successfully captured: ${extracted.equipment_tag || 'Transformer'} - Rating: ${extracted.rated_power_kva || 25000}kVA. Updated into MTS form!`
    );
    setTimeout(() => setSavedSuccessMessage(null), 6000);
  };

  const handleRunAiAnalysis = async () => {
    setIsAnalyzing(true);
    try {
      const response = await fetch('/api/field-service/neta-mts-transformer-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ...data, language: lang }),
      });

      if (response.ok) {
        const json = await response.json();
        if (json.success && json.data) {
          setAnalysisResult(json.data);
          setSavedSuccessMessage(
            lang === 'vi'
              ? 'Trợ lý AI TEV Platform đã phân tích thành công theo ANSI/NETA MTS-2023 Mục 7.2.2 & IEEE C57.104!'
              : 'TEV Platform AI has completed evaluation per ANSI/NETA MTS-2023 Section 7.2.2 & IEEE C57.104!'
          );
          setTimeout(() => setSavedSuccessMessage(null), 5000);
          return;
        }
      }
      // Fallback if API unavailable
      const fallback = performDeterministicMtsTransformerAnalysis({ ...data, language: lang });
      setAnalysisResult(fallback);
      setSavedSuccessMessage(
        lang === 'vi'
          ? 'Đã đối soát hoàn tất với bộ quy chuẩn ANSI/NETA MTS-2023 & IEEE C57.104.'
          : 'Deterministic evaluation completed per ANSI/NETA MTS-2023 & IEEE C57.104.'
      );
      setTimeout(() => setSavedSuccessMessage(null), 5000);
    } catch (err) {
      console.error('Error invoking MTS Transformer AI analysis:', err);
      const fallback = performDeterministicMtsTransformerAnalysis({ ...data, language: lang });
      setAnalysisResult(fallback);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleResetToSamplePrompt = () => {
    setData(SAMPLE_MTS_PROMPT_DATA);
    setAnalysisResult(null);
    setSavedSuccessMessage(
      lang === 'vi'
        ? 'Đã tải thành công Mẫu Dữ Liệu Kiểm Định MTS (Nha May Hoa Chat TEV Chemical Plant / TR-MAIN-110KV-01)!'
        : 'Loaded MTS benchmark sample data (TEV Chemical Plant / TR-MAIN-110KV-01)!'
    );
    setTimeout(() => setSavedSuccessMessage(null), 5000);
  };

  const handleCopyMarkdown = () => {
    const activeReport = analysisResult
      ? (lang === 'en' ? analysisResult.markdown_report_en : analysisResult.markdown_report_vi)
      : (lang === 'en' ? liveStats.analysis.markdown_report_en : liveStats.analysis.markdown_report_vi);
    navigator.clipboard.writeText(activeReport);
    setCopied(true);
    setTimeout(() => setCopied(false), 2500);
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
      setSavedSuccessMessage(
        lang === 'vi' ? 'Đã cập nhật dữ liệu JSON kiểm định hiện trường thành công!' : 'Updated JSON field test data successfully!'
      );
      setTimeout(() => setSavedSuccessMessage(null), 5000);
    } catch (err: any) {
      alert(`JSON Error: ${err.message}`);
    }
  };

  const handleSaveToCmms = () => {
    const currentEvaluation = analysisResult || liveStats.analysis;
    const reportPayload = {
      type: 'NETA_MTS_2023_LIQUID_TRANSFORMER',
      equipmentId: data.site_info.transformer_tag,
      date: data.site_info.test_date,
      title: `${lang === 'vi' ? 'Biên bản MTS-2023 MBA Dầu' : 'MTS-2023 Liquid Transformer Report'}: ${data.site_info.transformer_tag}`,
      standard: 'ANSI/NETA MTS-2023 Section 7.2.2 & IEEE C57.104',
      analysisResult: currentEvaluation,
      data: data,
      site_info: data.site_info,
      measurements: {
        dielectric_breakdown_kv: data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.dielectric_breakdown_d1816_1mm_kv,
        water_content_ppm: data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.water_content_d1533_ppm,
        acid_number: data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.acid_number_d974_mg_koh_g,
        ift_mn_m: data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.interfacial_tension_d971_mn_m,
        oil_power_factor_percent: data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.oil_power_factor_25c_d924_percent,
        c2h2_ppm: data.electrical_tests.dga_test_ieee_c57_104.acetylene_c2h2_ppm,
        c2h4_ppm: data.electrical_tests.dga_test_ieee_c57_104.ethylene_c2h4_ppm,
        h2_ppm: data.electrical_tests.dga_test_ieee_c57_104.hydrogen_h2_ppm,
        pi_value: data.electrical_tests.insulation_resistance_20c_megohms.calculated_pi_value,
        winding_pf_20c: data.electrical_tests.winding_insulation_power_factor_percent_20c,
        turns_ratio_error: data.electrical_tests.turns_ratio_all_taps.max_error_percent,
        core_ir_megohms: data.electrical_tests.core_insulation_resistance_500v_megohms,
      }
    };

    if (onSaveReport) {
      onSaveReport(reportPayload);
    }
    setSavedSuccessMessage(
      lang === 'vi'
        ? 'Đã lưu biên bản kiểm định bảo trì MTS vào hệ thống & đồng bộ tự động!'
        : 'Saved MTS maintenance test report to system & synced successfully!'
    );
    setTimeout(() => setSavedSuccessMessage(null), 5000);
  };

  const currentEvaluation = useMemo(() => {
    // ALWAYS dynamically evaluate from the current live field data!
    const res = performDeterministicMtsTransformerAnalysis({ ...data, language: lang });
    res.markdown_report_vi = buildDeterministicMtsTransformerMarkdown(data, res, 'vi');
    res.markdown_report_en = buildDeterministicMtsTransformerMarkdown(data, res, 'en');
    if (analysisResult) {
      // If AI analysis was explicitly run, preserve AI expert narratives while keeping evaluation status and numbers synchronized with latest data
      return {
        ...analysisResult,
        overall_status: res.overall_status,
        status_title_vi: res.status_title_vi,
        status_title_en: res.status_title_en,
        status_reason_vi: res.status_reason_vi,
        status_reason_en: res.status_reason_en,
        evaluations: res.evaluations,
        markdown_report_vi: res.markdown_report_vi,
        markdown_report_en: res.markdown_report_en
      };
    }
    return res;
  }, [data, lang, analysisResult]);

  // Real-time PMP helper
  const evalPmpField = (status: 'PASS' | 'INVESTIGATE' | 'FAIL', standardText: string) => {
    if (status === 'FAIL') {
      return {
        inputClass: 'border-rose-500 bg-rose-50/80 text-rose-950 font-black shadow-inner ring-1 ring-rose-400 focus:ring-rose-500',
        cardBorder: 'border-rose-400 bg-rose-50/30',
        badgeClass: 'bg-rose-100 text-rose-800 border-rose-300 font-black',
        badgeText: lang === 'vi' ? '✕ KHÔNG ĐẠT (FAIL)' : '✕ FAIL',
        icon: <XCircle size={10} className="text-rose-600 inline mr-0.5" />,
        standardText
      };
    }
    if (status === 'INVESTIGATE') {
      return {
        inputClass: 'border-amber-400 bg-amber-50/70 text-amber-950 font-black shadow-inner ring-1 ring-amber-300 focus:ring-amber-500',
        cardBorder: 'border-amber-400 bg-amber-50/30',
        badgeClass: 'bg-amber-100 text-amber-900 border-amber-300 font-black',
        badgeText: lang === 'vi' ? '⚠️ CẢNH BÁO (INVESTIGATE)' : '⚠️ INVESTIGATE',
        icon: <AlertTriangle size={10} className="text-amber-600 inline mr-0.5" />,
        standardText
      };
    }
    return {
      inputClass: 'border-emerald-400 bg-emerald-50/40 text-emerald-950 font-bold focus:ring-emerald-500',
      cardBorder: 'border-slate-200 bg-slate-50',
      badgeClass: 'bg-emerald-100 text-emerald-800 border-emerald-300 font-bold',
      badgeText: lang === 'vi' ? '✓ ĐẠT (PASS)' : '✓ PASS',
      icon: <CheckCircle2 size={10} className="text-emerald-600 inline mr-0.5" />,
      standardText
    };
  };

  // Real-time bolted connections deviation
  const boltedDev = useMemo(() => {
    const arr = data.electrical_tests.bolted_resistance_micro_ohms || [];
    if (!arr.length) return 0;
    const min = Math.min(...arr);
    const max = Math.max(...arr);
    return min > 0 ? ((max - min) / min) * 100 : 0;
  }, [data.electrical_tests.bolted_resistance_micro_ohms]);

  return (
    <div className="space-y-6">
      {/* SUCCESS BANNER TOAST */}
      {savedSuccessMessage && (
        <div className="p-4 rounded-xl bg-emerald-600 text-white font-medium flex items-center justify-between shadow-lg animate-in fade-in duration-300">
          <div className="flex items-center gap-2">
            <CheckCircle2 size={20} className="shrink-0" />
            <span>{savedSuccessMessage}</span>
          </div>
          <button
            onClick={() => setSavedSuccessMessage(null)}
            className="p-1 hover:bg-emerald-700 rounded-lg text-emerald-100 hover:text-white"
          >
            ✕
          </button>
        </div>
      )}

      {/* TOP HEADER: TEV BRANDING & DUAL LANGUAGE SWITCHER */}
      <div className="bg-gradient-to-r from-slate-950 via-slate-900 to-amber-950 text-white rounded-2xl p-6 shadow-xl border border-amber-900/40">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6">
          <div className="space-y-2">
            <div className="flex items-center gap-2.5 flex-wrap">
              {/* TEV Platform Logo Badge */}
              <div className="flex items-center gap-1.5 px-3 py-1 bg-amber-500/20 border border-amber-400/40 rounded-lg shadow-inner">
                <span className="font-black text-amber-400 text-base tracking-wider">TEV</span>
                <span className="text-[10px] font-bold text-amber-200 uppercase tracking-widest pl-1 border-l border-amber-400/40">
                  Engineering AI
                </span>
              </div>
              <span className="px-2.5 py-0.5 rounded-full text-xs font-bold uppercase tracking-wider bg-rose-500/20 text-rose-300 border border-rose-400/30">
                MODULE MTS-2023
              </span>
              <span className="px-2.5 py-0.5 rounded-full text-xs font-bold uppercase tracking-wider bg-blue-500/20 text-blue-300 border border-blue-400/30">
                ANSI/NETA MTS-2023 SEC 7.2.2
              </span>
            </div>

            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white flex items-center gap-3">
              <Droplets className="text-amber-400 shrink-0" size={26} />
              <span>
                {lang === 'vi'
                  ? 'Máy Biến Áp Ngâm Chất Lỏng / Dầu (Maintenance Testing)'
                  : 'Liquid-Filled Transformers (Maintenance Testing Specifications)'}
              </span>
            </h1>

            <p className="text-xs sm:text-sm text-slate-300 max-w-3xl leading-relaxed">
              {lang === 'vi'
                ? 'Biểu mẫu chuẩn PMP dành cho Kỹ sư Hiện trường (FSE) theo ANSI/NETA MTS-2023 (Mục 7.2.2), Table 100.4.1 & IEEE C57.104. Hiện tại hệ thống MTS tích hợp 01 biểu mẫu chuẩn (Máy biến áp ngâm dầu) — các biểu mẫu MTS thiết bị khác sẽ được cập nhật tiếp theo.'
                : 'PMP-standard Field Service Checklist per ANSI/NETA MTS-2023 (Section 7.2.2), Table 100.4.1 & IEEE C57.104. Currently featuring 1 standardized MTS form (Liquid-Filled Transformer) — additional MTS forms will be updated soon.'}
            </p>
          </div>

          {/* Action Hub & Dual-Language Control */}
          <div className="flex flex-wrap lg:flex-col items-stretch sm:items-end gap-2.5 shrink-0">
            {/* Language Switcher */}
            <div className="flex items-center gap-1 bg-slate-900/90 p-1 rounded-xl border border-slate-700 shadow-sm">
              <Globe size={15} className="text-slate-400 ml-1.5 mr-0.5" />
              <button
                type="button"
                onClick={() => setLang('vi')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                  lang === 'vi'
                    ? 'bg-amber-500 text-slate-950 shadow-md shadow-amber-500/30 font-black'
                    : 'text-slate-300 hover:text-white'
                }`}
              >
                🇻🇳 Tiếng Việt
              </button>
              <button
                type="button"
                onClick={() => setLang('en')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                  lang === 'en'
                    ? 'bg-amber-500 text-slate-950 shadow-md shadow-amber-500/30 font-black'
                    : 'text-slate-300 hover:text-white'
                }`}
              >
                🇺🇸 English
              </button>
            </div>

            <div className="flex items-center gap-2 flex-wrap">
              <button
                type="button"
                onClick={handleResetToSamplePrompt}
                className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-amber-300 border border-amber-500/30 rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all shadow-sm"
                title={lang === 'vi' ? 'Nạp lại bộ dữ liệu JSON mô phỏng trong yêu cầu' : 'Load default MTS simulation JSON'}
              >
                <RefreshCw size={14} />
                <span>{lang === 'vi' ? 'Nạp Dữ Liệu Mẫu Prompt' : 'Load Prompt Sample'}</span>
              </button>

              <button
                type="button"
                onClick={() => setShowCameraModal(true)}
                className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-blue-300 border border-blue-500/30 rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all shadow-sm"
              >
                <Camera size={14} />
                <span>{lang === 'vi' ? 'Quét Bảng Tên' : 'Scan Nameplate'}</span>
              </button>

              <button
                type="button"
                onClick={handleOpenJsonModal}
                className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-emerald-300 border border-emerald-500/30 rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all shadow-sm"
              >
                <FileText size={14} />
                <span>JSON FSE</span>
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* CRITICAL LIVE STATUS STRIP */}
      <div className={`p-4 rounded-2xl border text-sm flex flex-col md:flex-row md:items-center justify-between gap-4 transition-all shadow-sm ${
        currentEvaluation.overall_status === 'FAIL'
          ? 'bg-rose-50 border-rose-300 text-rose-950'
          : currentEvaluation.overall_status === 'INVESTIGATE'
          ? 'bg-amber-50 border-amber-300 text-amber-950'
          : 'bg-emerald-50 border-emerald-300 text-emerald-950'
      }`}>
        <div className="space-y-1">
          <div className="flex items-center gap-2 flex-wrap font-bold">
            {currentEvaluation.overall_status === 'FAIL' ? (
              <span className="flex items-center gap-1 text-rose-700 font-black">
                <AlertOctagon size={18} />
                <span>{lang === 'vi' ? currentEvaluation.status_title_vi : currentEvaluation.status_title_en}</span>
              </span>
            ) : currentEvaluation.overall_status === 'INVESTIGATE' ? (
              <span className="flex items-center gap-1 text-amber-700 font-black">
                <AlertTriangle size={18} />
                <span>{lang === 'vi' ? currentEvaluation.status_title_vi : currentEvaluation.status_title_en}</span>
              </span>
            ) : (
              <span className="flex items-center gap-1 text-emerald-700 font-black">
                <CheckCircle2 size={18} />
                <span>{lang === 'vi' ? currentEvaluation.status_title_vi : currentEvaluation.status_title_en}</span>
              </span>
            )}
            <span className="text-xs px-2 py-0.5 rounded-full font-mono font-semibold bg-white/70 border border-current">
              {data.site_info.transformer_tag}
            </span>
          </div>

          <p className="text-xs text-slate-700">
            {lang === 'vi' ? currentEvaluation.status_reason_vi : currentEvaluation.status_reason_en}
          </p>

          <div className="flex items-center gap-3 pt-1 text-xs text-slate-600 flex-wrap">
            <span className={`font-semibold ${liveStats.arcingFail ? 'text-rose-600 font-bold' : ''}`}>
              ⚡ C2H2: {data.electrical_tests.dga_test_ieee_c57_104.acetylene_c2h2_ppm} ppm {liveStats.arcingFail ? '(Arcing Fault!)' : ''}
            </span>
            <span>•</span>
            <span className={`font-semibold ${liveStats.breakdownFail ? 'text-rose-600 font-bold' : ''}`}>
              🧪 {lang === 'vi' ? 'Đánh thủng D1816:' : 'Breakdown D1816:'} {data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.dielectric_breakdown_d1816_1mm_kv} kV
            </span>
            <span>•</span>
            <span className={`font-semibold ${liveStats.waterFail ? 'text-rose-600 font-bold' : ''}`}>
              💧 {lang === 'vi' ? 'Ẩm D1533:' : 'Water D1533:'} {data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.water_content_d1533_ppm} ppm
            </span>
            <span>•</span>
            <span className={`font-semibold ${liveStats.acidFail ? 'text-rose-600 font-bold' : ''}`}>
              ⚗️ {lang === 'vi' ? 'Axit:' : 'Acid:'} {data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.acid_number_d974_mg_koh_g} mg KOH/g
            </span>
            <span>•</span>
            <span className="font-semibold text-emerald-700">
              🛡️ PI: {data.electrical_tests.insulation_resistance_20c_megohms.calculated_pi_value}
            </span>
          </div>
        </div>

        <div className="flex items-center gap-2 shrink-0 flex-wrap">
          <button
            type="button"
            onClick={() => setMtsViewMode(mtsViewMode === 'form' ? 'tev-report' : 'form')}
            className={`px-4 py-2.5 font-black rounded-xl text-xs flex items-center gap-2 shadow-md transition-all active:scale-95 ${
              mtsViewMode === 'tev-report'
                ? 'bg-slate-900 hover:bg-slate-800 text-amber-300 border border-amber-400/40'
                : 'bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white'
            }`}
          >
            <Sparkles size={15} className="text-amber-300" />
            <span>
              {mtsViewMode === 'tev-report'
                ? (lang === 'vi' ? '← Quay Lại Biểu Mẫu Nhập Liệu' : '← Back to Entry Form')
                : (lang === 'vi' ? 'Xem Báo Cáo Chuẩn TEV AI ⚡' : 'View TEV AI Report ⚡')}
            </span>
          </button>

          <button
            type="button"
            onClick={async () => {
              await handleRunAiAnalysis();
              setMtsViewMode('tev-report');
            }}
            disabled={isAnalyzing}
            className="px-4 py-2.5 bg-gradient-to-r from-amber-600 to-amber-500 hover:from-amber-500 hover:to-amber-400 text-slate-950 font-black rounded-xl text-xs flex items-center gap-2 shadow-md shadow-amber-500/20 active:scale-95 transition-all disabled:opacity-50"
          >
            {isAnalyzing ? (
              <RefreshCw size={15} className="animate-spin" />
            ) : (
              <Sparkles size={15} />
            )}
            <span>
              {isAnalyzing
                ? (lang === 'vi' ? 'Đang Phân Tích AI...' : 'Analyzing AI...')
                : (lang === 'vi' ? 'Phân Tích & Xuất Báo Cáo AI' : 'Run TEV AI Analysis')}
            </span>
          </button>
        </div>
      </div>

      {/* PMP 2-STEP MODE SWITCHER STRIP */}
      <div className="bg-slate-900 rounded-xl p-1.5 border border-slate-800 flex items-center justify-between gap-2 flex-wrap">
        <div className="flex items-center gap-1.5 bg-slate-950 p-1 rounded-lg border border-slate-800 flex-wrap sm:flex-nowrap">
          <button
            type="button"
            onClick={() => setMtsViewMode('form')}
            className={`flex items-center gap-1.5 px-4 py-2 rounded-lg text-xs font-black transition-all ${
              mtsViewMode === 'form'
                ? 'bg-amber-500 text-slate-950 shadow-md shadow-amber-500/20'
                : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
          >
            <Sliders size={14} />
            <span>
              {lang === 'vi'
                ? 'BƯỚC 1: BIỂU MẪU NHẬP LIỆU HIỆN TRƯỜNG (MTS-2023 FORM)'
                : 'STEP 1: FIELD MEASUREMENT ENTRY FORM (MTS-2023)'}
            </span>
          </button>

          <button
            type="button"
            onClick={() => setMtsViewMode('tev-report')}
            className={`flex items-center gap-1.5 px-4 py-2 rounded-lg text-xs font-black transition-all ${
              mtsViewMode === 'tev-report'
                ? 'bg-blue-600 text-white shadow-md shadow-blue-600/30'
                : 'text-slate-400 hover:text-white hover:bg-slate-800'
            }`}
          >
            <Sparkles size={14} className="text-amber-400" />
            <span>
              {lang === 'vi'
                ? 'BƯỚC 2: XEM BÁO CÁO CHUẨN TEV AI 4 PHẦN (MTS-2023)'
                : 'STEP 2: VIEW TEV STANDARD 4-PART AI REPORT'}
            </span>
          </button>
        </div>

        <span className="text-[11px] text-amber-300 font-mono px-2">
          {mtsViewMode === 'form' ? 'CHẾ ĐỘ: NHẬP BIỂU MẪU FSE' : 'CHẾ ĐỘ: BÁO CÁO CHUẨN TEV AI'}
        </span>
      </div>

      {mtsViewMode === 'form' && (
        <div className="space-y-6">
          {/* FORM SECTIONS NAVIGATION TABS */}
          <div className="flex items-center gap-2 overflow-x-auto pb-1 border-b border-slate-200">
            {[
              { id: 'overview', vi: '1. Khách Hàng, Dự Án & Thiết Bị', en: '1. Customer, Project & Asset', icon: Sliders },
              { id: 'visual', vi: '2. Thị Giác & Cơ Khí (Sec 7.2.2.A)', en: '2. Visual & Mechanical (7.2.2.A)', icon: CheckCircle2 },
              { id: 'electrical', vi: '3. Phép Đo Điện & Cuộn Dây', en: '3. Electrical & Winding Tests', icon: Zap },
              { id: 'oil', vi: '4. Dầu Cách Điện Lão Hóa (Table 100.4.1)', en: '4. Service-Aged Oil (Table 100.4.1)', icon: Droplets },
              { id: 'dga', vi: '5. Khí Hòa Tan DGA (IEEE C57.104)', en: '5. DGA Gas Analysis (C57.104)', icon: Flame },
            ].map((tab) => {
              const TabIcon = tab.icon;
              return (
                <button
                  key={tab.id}
                  type="button"
                  onClick={() => setActiveFormTab(tab.id as any)}
                  className={`flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-xs font-bold whitespace-nowrap transition-all ${
                    activeFormTab === tab.id
                      ? 'bg-amber-500 text-slate-950 shadow-sm'
                      : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
                  }`}
                >
                  <TabIcon size={14} />
                  <span>{lang === 'vi' ? tab.vi : tab.en}</span>
                </button>
              );
            })}
          </div>

          {/* TAB CONTENT 1: OVERVIEW & NAMEPLATE */}
          {activeFormTab === 'overview' && (
            <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm space-y-4">
              <div className="flex items-center justify-between flex-wrap gap-2">
                <h3 className="font-bold text-slate-900 text-sm flex items-center gap-2">
                  <Sliders size={16} className="text-amber-600" />
                  <span>{lang === 'vi' ? 'Thông Tin Khách Hàng, Dự Án & Thiết Bị Thí Nghiệm (Chuẩn PMP)' : 'Customer, Project & Tested Equipment Information'}</span>
                </h3>
                <span className="text-[11px] bg-amber-50 text-amber-900 border border-amber-200 px-2.5 py-0.5 rounded-full font-semibold">
                  {lang === 'vi' ? 'Tự động đồng bộ từ Quản lý dự án (Mục 4. Soạn dự án)' : 'Auto-synced from Project Management'}
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 text-xs">
                <div>
                  <label className="block text-slate-700 font-bold mb-1">
                    {lang === 'vi' ? 'Tên Khách Hàng (Master)' : 'Customer Name (Master)'}
                  </label>
                  <input
                    type="text"
                    placeholder="VD: Công ty Cổ phần Năng lượng TEV"
                    value={data.site_info.customer_name || ''}
                    onChange={(e) =>
                      setData({
                        ...data,
                        site_info: { ...data.site_info, customer_name: e.target.value }
                      })
                    }
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-amber-500 bg-slate-50 font-bold text-blue-900"
                  />
                </div>

                <div>
                  <label className="block text-slate-700 font-bold mb-1">
                    {lang === 'vi' ? 'Tên Dự Án / Trạm Thí Nghiệm' : 'Project / Substation Name'}
                  </label>
                  <input
                    type="text"
                    value={data.site_info.project_name}
                    onChange={(e) =>
                      setData({
                        ...data,
                        site_info: { ...data.site_info, project_name: e.target.value }
                      })
                    }
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-amber-500 bg-slate-50 font-semibold text-slate-900"
                  />
                </div>

                <div>
                  <label className="block text-slate-700 font-bold mb-1">
                    {lang === 'vi' ? 'Mã Tag Thiết Bị / KKS Code' : 'Transformer Tag / KKS Code'}
                  </label>
                  <input
                    type="text"
                    value={data.site_info.transformer_tag}
                    onChange={(e) =>
                      setData({
                        ...data,
                        site_info: { ...data.site_info, transformer_tag: e.target.value }
                      })
                    }
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-amber-500 font-mono font-bold text-indigo-800 bg-slate-50"
                  />
                </div>

            <div>
              <label className="block text-slate-600 font-semibold mb-1">
                {lang === 'vi' ? 'Thông số kỹ thuật / Nhãn máy' : 'Rating & Transformer Type'}
              </label>
              <input
                type="text"
                value={data.site_info.transformer_rating || ''}
                onChange={(e) =>
                  setData({
                    ...data,
                    site_info: { ...data.site_info, transformer_rating: e.target.value }
                  })
                }
                className="w-full px-3 py-2 border rounded-lg focus:ring-2 focus:ring-amber-500 bg-slate-50"
              />
            </div>

            <div>
              <label className="block text-slate-600 font-semibold mb-1">
                {lang === 'vi' ? 'Kỹ sư FSE thí nghiệm' : 'Field Service Engineer (FSE)'}
              </label>
              <FseEngineerDropdown
                currentEngineer={data.site_info.fse_name}
                onSelectEngineer={(eng) =>
                  setData({
                    ...data,
                    site_info: { ...data.site_info, fse_name: eng }
                  })
                }
              />
            </div>

            <div>
              <label className="block text-slate-600 font-semibold mb-1">
                {lang === 'vi' ? 'Ngày thí nghiệm' : 'Test Date'}
              </label>
              <input
                type="date"
                value={data.site_info.test_date}
                onChange={(e) =>
                  setData({
                    ...data,
                    site_info: { ...data.site_info, test_date: e.target.value }
                  })
                }
                className="w-full px-3 py-2 border rounded-lg focus:ring-2 focus:ring-amber-500 bg-slate-50 font-mono"
              />
            </div>

            <div>
              <label className="block text-slate-600 font-semibold mb-1">
                {lang === 'vi' ? 'Cấp điện áp thử nghiệm (kV)' : 'Voltage Class (kV)'}
              </label>
              <select
                value={data.site_info.voltage_class_kv || 110}
                onChange={(e) =>
                  setData({
                    ...data,
                    site_info: { ...data.site_info, voltage_class_kv: Number(e.target.value) }
                  })
                }
                className="w-full px-3 py-2 border rounded-lg focus:ring-2 focus:ring-amber-500 bg-slate-50 font-bold"
              >
                <option value={22}>22 kV (≤ 69kV - Table 100.4.1 Min 23kV)</option>
                <option value={35}>35 kV (≤ 69kV - Table 100.4.1 Min 23kV)</option>
                <option value={110}>110 kV (&gt;69kV - &lt;230kV - Table 100.4.1 Min 28kV)</option>
                <option value={220}>220 kV (&gt;69kV - &lt;230kV - Table 100.4.1 Min 28kV)</option>
                <option value={500}>500 kV (≥ 230kV - Table 100.4.1 Min 30kV)</option>
              </select>
            </div>
          </div>
        </div>
      )}

      {/* TAB CONTENT 2: VISUAL & MECHANICAL */}
      {activeFormTab === 'visual' && (
        <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="font-bold text-slate-900 text-sm flex items-center gap-2">
              <CheckCircle2 size={16} className="text-emerald-600" />
              <span>{lang === 'vi' ? 'Kiểm Tra Thị Giác & Cơ Khí (ANSI/NETA MTS-2023 Mục 7.2.2.A)' : 'Visual & Mechanical Inspection (NETA MTS-2023 Sec 7.2.2.A)'}</span>
            </h3>
            <span className="text-xs text-slate-500 font-semibold">IEEE C57.12.00 / NETA Table 100.12</span>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
            <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 flex items-center justify-between">
              <div>
                <div className="font-bold text-slate-800">
                  {lang === 'vi' ? '1. Đối soát nhãn mác (IEEE C57.12.00)' : '1. Nameplate Match (IEEE C57.12.00)'}
                </div>
                <div className="text-slate-500 text-[11px]">
                  {lang === 'vi' ? 'Khớp 100% bản vẽ kỹ thuật & sơ đồ đơn tuyến' : '100% matches drawings & single line diagrams'}
                </div>
              </div>
              <label className="flex items-center gap-2 cursor-pointer font-bold">
                <input
                  type="checkbox"
                  checked={data.visual_inspection.nameplate_match}
                  onChange={(e) =>
                    setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, nameplate_match: e.target.checked }
                    })
                  }
                  className="rounded text-amber-600 focus:ring-amber-500 w-4 h-4"
                />
                <span className={data.visual_inspection.nameplate_match ? 'text-emerald-700' : 'text-rose-600'}>
                  {data.visual_inspection.nameplate_match ? 'PASS' : 'FAIL'}
                </span>
              </label>
            </div>

            <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 flex items-center justify-between">
              <div>
                <div className="font-bold text-slate-800">
                  {lang === 'vi' ? '2. Đệm khí Nitơ duy trì áp suất dương (+)' : '2. Nitrogen Gas Blanket Positive Pressure (+)'}
                </div>
                <div className="text-slate-500 text-[11px]">
                  {lang === 'vi' ? 'BẮT BUỘC áp suất dương (positive pressure 0.5 - 5 psi)' : 'MUST maintain positive pressure gauge reading'}
                </div>
              </div>
              <label className="flex items-center gap-2 cursor-pointer font-bold">
                <input
                  type="checkbox"
                  checked={data.visual_inspection.gas_blanket_pressure_positive}
                  onChange={(e) =>
                    setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, gas_blanket_pressure_positive: e.target.checked }
                    })
                  }
                  className="rounded text-amber-600 focus:ring-amber-500 w-4 h-4"
                />
                <span className={data.visual_inspection.gas_blanket_pressure_positive ? 'text-emerald-700' : 'text-rose-600'}>
                  {data.visual_inspection.gas_blanket_pressure_positive ? 'PASS (Positive)' : 'FAIL (Zero/Negative)'}
                </span>
              </label>
            </div>

            <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
              <label className="font-bold text-slate-800 block mb-1">
                {lang === 'vi' ? '3. Mức chất lỏng trên tank chính, LTC, bushing' : '3. Liquid Level (Main Tank, LTC, Bushings)'}
              </label>
              <input
                type="text"
                value={data.visual_inspection.liquid_level_check}
                onChange={(e) =>
                  setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, liquid_level_check: e.target.value }
                  })
                }
                className="w-full px-2.5 py-1.5 border rounded-lg bg-white"
              />
            </div>

            <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
              <label className="font-bold text-slate-800 block mb-1">
                {lang === 'vi' ? '4. Van tản nhiệt mở 100%, quạt & bơm dầu' : '4. Valves Open, Radiator Fans & Pumps'}
              </label>
              <input
                type="text"
                value={data.visual_inspection.valves_and_cooling_fans_check}
                onChange={(e) =>
                  setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, valves_and_cooling_fans_check: e.target.value }
                  })
                }
                className="w-full px-2.5 py-1.5 border rounded-lg bg-white"
              />
            </div>

            <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 flex items-center justify-between">
              <div>
                <div className="font-bold text-slate-800">
                  {lang === 'vi' ? '5. Lực siết bu-lông (NETA Table 100.12)' : '5. Bolt Torque (NETA Table 100.12)'}
                </div>
                <div className="text-slate-500 text-[11px]">
                  {lang === 'vi' ? 'Siết lực calib cờ lê lực theo nhà sản xuất / Table 100.12' : 'Calibrated torque wrench per Table 100.12'}
                </div>
              </div>
              <input
                type="text"
                value={data.visual_inspection.bolt_torque_check}
                onChange={(e) =>
                  setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, bolt_torque_check: e.target.value }
                  })
                }
                className="w-28 px-2 py-1 border rounded-lg bg-white text-center font-bold"
              />
            </div>

            <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 flex items-center justify-between">
              <div>
                <div className="font-bold text-slate-800">
                  {lang === 'vi' ? '6. Không rò rỉ dầu (No Oil Leakage)' : '6. No Oil Leakage on Tank & Valves'}
                </div>
                <div className="text-slate-500 text-[11px]">
                  {lang === 'vi' ? 'Kiểm tra mặt bích, roăng cao su, kính quan sát' : 'Flanges, gaskets, sight glass dry and intact'}
                </div>
              </div>
              <label className="flex items-center gap-2 cursor-pointer font-bold">
                <input
                  type="checkbox"
                  checked={data.visual_inspection.no_oil_leakage}
                  onChange={(e) =>
                    setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, no_oil_leakage: e.target.checked }
                    })
                  }
                  className="rounded text-amber-600 focus:ring-amber-500 w-4 h-4"
                />
                <span className={data.visual_inspection.no_oil_leakage ? 'text-emerald-700' : 'text-rose-600'}>
                  {data.visual_inspection.no_oil_leakage ? 'PASS (No leak)' : 'FAIL (Leaking)'}
                </span>
              </label>
            </div>
          </div>
        </div>
      )}

      {/* TAB CONTENT 3: ELECTRICAL TESTS */}
      {activeFormTab === 'electrical' && (
        <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="font-bold text-slate-900 text-sm flex items-center gap-2">
              <Zap size={16} className="text-amber-600" />
              <span>{lang === 'vi' ? 'Phép Đo Điện & Cuộn Dây (NETA MTS-2023 Sec 7.2.2.B)' : 'Electrical Maintenance Tests (NETA MTS-2023 Sec 7.2.2.B)'}</span>
            </h3>
            <span className="text-xs text-slate-500 font-semibold">Table 100.3, Table 100.5, Table 100.12</span>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 text-xs">
            {/* Bolted resistance */}
            {(() => {
              const pmp = evalPmpField(boltedDev > 50 ? 'INVESTIGATE' : 'PASS', 'Table 100.12 ≤ 50%');
              return (
                <div className={`p-3.5 rounded-xl border transition-all ${pmp.cardBorder} space-y-2`}>
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <span className="font-bold text-slate-800">
                      {lang === 'vi' ? 'Điện trở mối nối bu-lông (µΩ)' : 'Bolted Resistance (µΩ)'}
                    </span>
                    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] border ${pmp.badgeClass}`}>
                      {pmp.icon}
                      <span>{pmp.badgeText}</span>
                    </span>
                  </div>
                  <div className="flex items-center justify-between text-[11px] text-slate-500">
                    <span>Chuẩn Table 100.12: Lệch ≤ 50% min</span>
                    <span className="font-mono font-bold text-slate-700">Lệch: {boltedDev.toFixed(1)}%</span>
                  </div>
                  <div className="flex gap-2">
                    {data.electrical_tests.bolted_resistance_micro_ohms.map((val, idx) => (
                      <input
                        key={idx}
                        type="number"
                        step="0.1"
                        value={val}
                        onChange={(e) => {
                          const next = [...data.electrical_tests.bolted_resistance_micro_ohms];
                          next[idx] = Number(e.target.value);
                          setAnalysisResult(null);
                          setData({
                            ...data,
                            electrical_tests: {
                              ...data.electrical_tests,
                              bolted_resistance_micro_ohms: next,
                            }
                          });
                        }}
                        className={`w-1/3 px-2 py-1.5 border rounded-lg font-mono text-center transition-all ${pmp.inputClass}`}
                      />
                    ))}
                  </div>
                </div>
              );
            })()}

            {/* Winding IR @ 20°C: Pri, Sec, Pri-Sec, PI */}
            {(() => {
              const pri = data.electrical_tests.insulation_resistance_20c_megohms.pri_to_ground_5000v;
              const pi = data.electrical_tests.insulation_resistance_20c_megohms.calculated_pi_value;
              const irStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = (pi < 1.0 || (pri > 0 && pri < 1000)) ? 'FAIL' : pi < 1.5 ? 'INVESTIGATE' : 'PASS';
              const pmp = evalPmpField(irStatus, 'Table 100.5 (PI ≥ 1.0, IR ≥ 1000 MΩ)');

              return (
                <div className={`p-3.5 rounded-xl border transition-all ${pmp.cardBorder} space-y-2`}>
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <span className="font-bold text-slate-800">
                      {lang === 'vi' ? 'Cách điện Cuộn dây & PI @ 20°C' : 'Winding IR & PI @ 20°C'}
                    </span>
                    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] border ${pmp.badgeClass}`}>
                      {pmp.icon}
                      <span>{pmp.badgeText}</span>
                    </span>
                  </div>
                  <div className="text-[11px] text-slate-500">
                    <span>Chuẩn Table 100.5: PI ≥ 1.0 • Pri-Gnd ≥ 1000 MΩ</span>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <span className="text-[10px] text-slate-500 block">Pri-Gnd (5000V)</span>
                      <input
                        type="number"
                        value={data.electrical_tests.insulation_resistance_20c_megohms.pri_to_ground_5000v}
                        onChange={(e) => {
                          setAnalysisResult(null);
                          setData({
                            ...data,
                            electrical_tests: {
                              ...data.electrical_tests,
                              insulation_resistance_20c_megohms: {
                                ...data.electrical_tests.insulation_resistance_20c_megohms,
                                pri_to_ground_5000v: Number(e.target.value)
                              }
                            }
                          });
                        }}
                        className={`w-full px-2 py-1.5 border rounded-lg font-mono text-xs font-bold transition-all ${
                          pri < 1000 ? 'border-rose-400 bg-rose-50 text-rose-950 ring-1 ring-rose-400' : 'border-slate-300 bg-white'
                        }`}
                      />
                    </div>
                    <div>
                      <span className="text-[10px] text-slate-500 block">Sec-Gnd (2500V)</span>
                      <input
                        type="number"
                        value={data.electrical_tests.insulation_resistance_20c_megohms.sec_to_ground_2500v}
                        onChange={(e) => {
                          setAnalysisResult(null);
                          setData({
                            ...data,
                            electrical_tests: {
                              ...data.electrical_tests,
                              insulation_resistance_20c_megohms: {
                                ...data.electrical_tests.insulation_resistance_20c_megohms,
                                sec_to_ground_2500v: Number(e.target.value)
                              }
                            }
                          });
                        }}
                        className="w-full px-2 py-1.5 border border-slate-300 rounded-lg bg-white font-mono text-xs font-bold"
                      />
                    </div>
                    <div>
                      <span className="text-[10px] text-slate-500 block">Pri-Sec (5000V DC)</span>
                      <input
                        type="number"
                        value={data.electrical_tests.insulation_resistance_20c_megohms.pri_to_sec_5000v_dc}
                        onChange={(e) => {
                          setAnalysisResult(null);
                          setData({
                            ...data,
                            electrical_tests: {
                              ...data.electrical_tests,
                              insulation_resistance_20c_megohms: {
                                ...data.electrical_tests.insulation_resistance_20c_megohms,
                                pri_to_sec_5000v_dc: Number(e.target.value)
                              }
                            }
                          });
                        }}
                        className="w-full px-2 py-1.5 border border-slate-300 rounded-lg bg-white font-mono text-xs font-bold"
                      />
                    </div>
                    <div>
                      <span className="text-[10px] font-bold block text-slate-700">Chỉ số PI (≥ 1.0)</span>
                      <input
                        type="number"
                        step="0.01"
                        value={data.electrical_tests.insulation_resistance_20c_megohms.calculated_pi_value}
                        onChange={(e) => {
                          setAnalysisResult(null);
                          setData({
                            ...data,
                            electrical_tests: {
                              ...data.electrical_tests,
                              insulation_resistance_20c_megohms: {
                                ...data.electrical_tests.insulation_resistance_20c_megohms,
                                calculated_pi_value: Number(e.target.value)
                              }
                            }
                          });
                        }}
                        className={`w-full px-2 py-1.5 border rounded-lg font-mono text-xs font-black transition-all ${
                          pi < 1.0
                            ? 'border-rose-500 bg-rose-100 text-rose-950 ring-1 ring-rose-400'
                            : pi < 1.5
                            ? 'border-amber-400 bg-amber-50 text-amber-950 ring-1 ring-amber-300'
                            : 'border-emerald-400 bg-emerald-50 text-emerald-950'
                        }`}
                      />
                    </div>
                  </div>
                </div>
              );
            })()}

            {/* Turns Ratio & Winding Power Factor */}
            {(() => {
              const ratioErr = data.electrical_tests.turns_ratio_all_taps.max_error_percent;
              const pf = data.electrical_tests.winding_insulation_power_factor_percent_20c;
              const ratioStatus: 'PASS' | 'FAIL' = ratioErr > 0.5 ? 'FAIL' : 'PASS';
              const pfStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = pf > 1.0 ? 'FAIL' : pf > 0.5 ? 'INVESTIGATE' : 'PASS';
              const combinedStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = (ratioStatus === 'FAIL' || pfStatus === 'FAIL') ? 'FAIL' : pfStatus === 'INVESTIGATE' ? 'INVESTIGATE' : 'PASS';
              const pmp = evalPmpField(combinedStatus, 'Table 100.3 (≤ 0.5% ratio, ≤ 1.0% PF)');

              return (
                <div className={`p-3.5 rounded-xl border transition-all ${pmp.cardBorder} space-y-2`}>
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <span className="font-bold text-slate-800">
                      {lang === 'vi' ? 'Tỷ số biến áp & Power Factor @ 20°C' : 'Turns-Ratio & Power Factor @ 20°C'}
                    </span>
                    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] border ${pmp.badgeClass}`}>
                      {pmp.icon}
                      <span>{pmp.badgeText}</span>
                    </span>
                  </div>
                  <div>
                    <div className="flex items-center justify-between text-[11px] mb-1">
                      <span className="text-slate-500">Sai số tỷ số biến áp Max (%) - Chuẩn ≤ 0.5%</span>
                      <span className={`font-mono font-bold text-[10px] px-1.5 py-0.2 rounded ${ratioErr > 0.5 ? 'bg-rose-100 text-rose-800' : 'bg-emerald-100 text-emerald-800'}`}>
                        {ratioErr > 0.5 ? 'VƯỢT CHUẨN' : 'ĐẠT CHUẨN'}
                      </span>
                    </div>
                    <input
                      type="number"
                      step="0.01"
                      value={data.electrical_tests.turns_ratio_all_taps.max_error_percent}
                      onChange={(e) => {
                        setAnalysisResult(null);
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            turns_ratio_all_taps: { max_error_percent: Number(e.target.value) }
                          }
                        });
                      }}
                      className={`w-full px-2.5 py-1.5 border rounded-lg font-mono text-xs font-bold transition-all ${
                        ratioErr > 0.5
                          ? 'border-rose-500 bg-rose-100 text-rose-950 ring-1 ring-rose-400'
                          : 'border-emerald-400 bg-emerald-50/50 text-emerald-950'
                      }`}
                    />
                  </div>
                  <div>
                    <div className="flex items-center justify-between text-[11px] mb-1">
                      <span className="text-slate-500">Winding Power Factor (%) - Table 100.3 ≤ 1.0%</span>
                      <span className={`font-mono font-bold text-[10px] px-1.5 py-0.2 rounded ${pf > 1.0 ? 'bg-rose-100 text-rose-800' : pf > 0.5 ? 'bg-amber-100 text-amber-800' : 'bg-emerald-100 text-emerald-800'}`}>
                        {pf > 1.0 ? 'FAIL' : pf > 0.5 ? 'INVESTIGATE' : 'PASS'}
                      </span>
                    </div>
                    <input
                      type="number"
                      step="0.01"
                      value={data.electrical_tests.winding_insulation_power_factor_percent_20c}
                      onChange={(e) => {
                        setAnalysisResult(null);
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            winding_insulation_power_factor_percent_20c: Number(e.target.value)
                          }
                        });
                      }}
                      className={`w-full px-2.5 py-1.5 border rounded-lg font-mono text-xs font-bold transition-all ${
                        pf > 1.0
                          ? 'border-rose-500 bg-rose-100 text-rose-950 ring-1 ring-rose-400'
                          : pf > 0.5
                          ? 'border-amber-400 bg-amber-50 text-amber-950 ring-1 ring-amber-300'
                          : 'border-emerald-400 bg-emerald-50/50 text-emerald-950'
                      }`}
                    />
                  </div>
                </div>
              );
            })()}

            {/* Winding DC Resistance deviation & Core IR */}
            {(() => {
              const dev = data.electrical_tests.winding_dc_resistance_ohms?.max_phase_deviation_percent ?? 1.2;
              const core = data.electrical_tests.core_insulation_resistance_500v_megohms ?? 350;
              const dcStatus: 'PASS' | 'INVESTIGATE' | 'FAIL' = dev > 5.0 ? 'FAIL' : dev > 2.0 ? 'INVESTIGATE' : 'PASS';
              const coreStatus: 'PASS' | 'FAIL' = core < 100 ? 'FAIL' : 'PASS';
              const combined: 'PASS' | 'INVESTIGATE' | 'FAIL' = (dcStatus === 'FAIL' || coreStatus === 'FAIL') ? 'FAIL' : dcStatus === 'INVESTIGATE' ? 'INVESTIGATE' : 'PASS';
              const pmp = evalPmpField(combined, 'NETA Section 7.2.2.B (≤ 2.0% Dev, ≥ 100 MΩ Core)');

              return (
                <div className={`p-3.5 rounded-xl border transition-all ${pmp.cardBorder} space-y-2`}>
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <span className="font-bold text-slate-800">
                      {lang === 'vi' ? 'Điện trở DC & Cách điện lõi thép' : 'Winding DC & Core IR'}
                    </span>
                    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] border ${pmp.badgeClass}`}>
                      {pmp.icon}
                      <span>{pmp.badgeText}</span>
                    </span>
                  </div>
                  <div>
                    <div className="flex items-center justify-between text-[11px] mb-1">
                      <span className="text-slate-500">Lệch điện trở DC giữa các pha (%) ≤ 2.0%</span>
                      <span className={`font-mono font-bold text-[10px] px-1.5 py-0.2 rounded ${dev > 5.0 ? 'bg-rose-100 text-rose-800' : dev > 2.0 ? 'bg-amber-100 text-amber-800' : 'bg-emerald-100 text-emerald-800'}`}>
                        {dev > 2.0 ? 'LỆCH > 2%' : 'ĐẠT'}
                      </span>
                    </div>
                    <input
                      type="number"
                      step="0.1"
                      value={dev}
                      onChange={(e) => {
                        setAnalysisResult(null);
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            winding_dc_resistance_ohms: { max_phase_deviation_percent: Number(e.target.value) }
                          }
                        });
                      }}
                      className={`w-full px-2.5 py-1.5 border rounded-lg font-mono text-xs font-bold transition-all ${
                        dev > 5.0
                          ? 'border-rose-500 bg-rose-100 text-rose-950 ring-1 ring-rose-400'
                          : dev > 2.0
                          ? 'border-amber-400 bg-amber-50 text-amber-950 ring-1 ring-amber-300'
                          : 'border-emerald-400 bg-emerald-50/50 text-emerald-950'
                      }`}
                    />
                  </div>
                  <div>
                    <div className="flex items-center justify-between text-[11px] mb-1">
                      <span className="text-slate-500">Cách điện lõi thép 500V DC (MΩ) ≥ 100 MΩ</span>
                      <span className={`font-mono font-bold text-[10px] px-1.5 py-0.2 rounded ${core < 100 ? 'bg-rose-100 text-rose-800' : 'bg-emerald-100 text-emerald-800'}`}>
                        {core < 100 ? 'KHÔNG ĐẠT' : 'ĐẠT'}
                      </span>
                    </div>
                    <input
                      type="number"
                      value={core}
                      onChange={(e) => {
                        setAnalysisResult(null);
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            core_insulation_resistance_500v_megohms: Number(e.target.value)
                          }
                        });
                      }}
                      className={`w-full px-2.5 py-1.5 border rounded-lg font-mono text-xs font-bold transition-all ${
                        core < 100
                          ? 'border-rose-500 bg-rose-100 text-rose-950 ring-1 ring-rose-400'
                          : 'border-emerald-400 bg-emerald-50/50 text-emerald-950'
                      }`}
                    />
                  </div>
                </div>
              );
            })()}
          </div>
        </div>
      )}

      {/* TAB CONTENT 4: SERVICE-AGED OIL SAMPLES */}
      {activeFormTab === 'oil' && (
        <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="font-bold text-slate-900 text-sm flex items-center gap-2">
              <Droplets size={16} className="text-amber-600" />
              <span>{lang === 'vi' ? 'Mẫu Dầu Cách Điện Lão Hóa (ASTM D923 & NETA MTS Table 100.4.1)' : 'Service-Aged Insulating Fluid (ASTM D923 & NETA Table 100.4.1)'}</span>
            </h3>
            <span className="text-xs px-2.5 py-0.5 rounded-full font-bold bg-amber-100 text-amber-900 border border-amber-300">
              {lang === 'vi' ? 'Dầu Khoáng Mineral Oil Class I' : 'Class I Mineral Oil'}
            </span>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 text-xs">
            {/* Breakdown ASTM D1816 1mm */}
            {(() => {
              const bd = data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.dielectric_breakdown_d1816_1mm_kv;
              const status: 'PASS' | 'FAIL' = bd < liveStats.breakdownThreshold ? 'FAIL' : 'PASS';
              const pmp = evalPmpField(status, `Table 100.4.1 (≥ ${liveStats.breakdownThreshold} kV)`);

              return (
                <div className={`p-3.5 rounded-xl border transition-all ${pmp.cardBorder} space-y-1.5`}>
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <span className="font-bold text-slate-800">
                      {lang === 'vi' ? 'Điện áp đánh thủng D1816 (1mm)' : 'Breakdown D1816 (1mm gap)'}
                    </span>
                    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] border ${pmp.badgeClass}`}>
                      {pmp.icon}
                      <span>{pmp.badgeText}</span>
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-500">
                    {lang === 'vi' ? `Chuẩn Table 100.4.1: Tối thiểu ≥ ${liveStats.breakdownThreshold} kV` : `Table 100.4.1 limit: ≥ ${liveStats.breakdownThreshold} kV`}
                  </p>
                  <div className="flex items-center gap-2 pt-1">
                    <input
                      type="number"
                      step="0.1"
                      value={bd}
                      onChange={(e) => {
                        setAnalysisResult(null);
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            service_aged_oil_sample_tests_table_100_4_1: {
                              ...data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1,
                              dielectric_breakdown_d1816_1mm_kv: Number(e.target.value)
                            }
                          }
                        });
                      }}
                      className={`w-full px-2.5 py-1.5 border rounded-lg font-mono text-sm font-black transition-all ${pmp.inputClass}`}
                    />
                    <span className="font-bold text-slate-600">kV</span>
                  </div>
                  {status === 'FAIL' && (
                    <p className="text-[11px] text-rose-700 font-bold pt-0.5">
                      ✕ {lang === 'vi' ? `Dưới ngưỡng tối thiểu ${liveStats.breakdownThreshold} kV, dầu bị nhiễm bẩn!` : `Below limit ${liveStats.breakdownThreshold} kV!`}
                    </p>
                  )}
                </div>
              );
            })()}

            {/* Water Content ASTM D1533 */}
            {(() => {
              const w = data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.water_content_d1533_ppm;
              const status: 'PASS' | 'FAIL' = w > liveStats.waterThreshold ? 'FAIL' : 'PASS';
              const pmp = evalPmpField(status, `Table 100.4.1 (≤ ${liveStats.waterThreshold} ppm)`);

              return (
                <div className={`p-3.5 rounded-xl border transition-all ${pmp.cardBorder} space-y-1.5`}>
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <span className="font-bold text-slate-800">
                      {lang === 'vi' ? 'Hàm lượng ẩm D1533' : 'Moisture in Oil D1533'}
                    </span>
                    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] border ${pmp.badgeClass}`}>
                      {pmp.icon}
                      <span>{pmp.badgeText}</span>
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-500">
                    {lang === 'vi' ? `Chuẩn Table 100.4.1: Tối đa ≤ ${liveStats.waterThreshold} ppm` : `Table 100.4.1 limit: ≤ ${liveStats.waterThreshold} ppm`}
                  </p>
                  <div className="flex items-center gap-2 pt-1">
                    <input
                      type="number"
                      value={w}
                      onChange={(e) => {
                        setAnalysisResult(null);
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            service_aged_oil_sample_tests_table_100_4_1: {
                              ...data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1,
                              water_content_d1533_ppm: Number(e.target.value)
                            }
                          }
                        });
                      }}
                      className={`w-full px-2.5 py-1.5 border rounded-lg font-mono text-sm font-black transition-all ${pmp.inputClass}`}
                    />
                    <span className="font-bold text-slate-600">ppm</span>
                  </div>
                  {status === 'FAIL' && (
                    <p className="text-[11px] text-rose-700 font-bold pt-0.5">
                      ✕ {lang === 'vi' ? `Vượt quá giới hạn tối đa ${liveStats.waterThreshold} ppm, dầu bị ẩm!` : `Exceeds max limit ${liveStats.waterThreshold} ppm!`}
                    </p>
                  )}
                </div>
              );
            })()}

            {/* Neutralization Acid Number ASTM D974 */}
            {(() => {
              const acid = data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.acid_number_d974_mg_koh_g;
              const status: 'PASS' | 'INVESTIGATE' | 'FAIL' = acid > 0.20 ? 'FAIL' : acid > 0.15 ? 'INVESTIGATE' : 'PASS';
              const pmp = evalPmpField(status, 'Table 100.4.1 (≤ 0.20 mg KOH/g)');

              return (
                <div className={`p-3.5 rounded-xl border transition-all ${pmp.cardBorder} space-y-1.5`}>
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <span className="font-bold text-slate-800">
                      {lang === 'vi' ? 'Chỉ số Axit trung hòa D974' : 'Acid Number D974'}
                    </span>
                    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] border ${pmp.badgeClass}`}>
                      {pmp.icon}
                      <span>{pmp.badgeText}</span>
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-500">
                    {lang === 'vi' ? 'Chuẩn Table 100.4.1: Giới hạn ≤ 0.20 mg KOH/g' : 'Table 100.4.1 limit: ≤ 0.20 mg KOH/g'}
                  </p>
                  <div className="flex items-center gap-2 pt-1">
                    <input
                      type="number"
                      step="0.01"
                      value={acid}
                      onChange={(e) => {
                        setAnalysisResult(null);
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            service_aged_oil_sample_tests_table_100_4_1: {
                              ...data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1,
                              acid_number_d974_mg_koh_g: Number(e.target.value)
                            }
                          }
                        });
                      }}
                      className={`w-full px-2.5 py-1.5 border rounded-lg font-mono text-sm font-black transition-all ${pmp.inputClass}`}
                    />
                    <span className="font-bold text-slate-600">mg KOH/g</span>
                  </div>
                  {status === 'FAIL' && (
                    <p className="text-[11px] text-rose-700 font-bold pt-0.5">
                      ✕ {lang === 'vi' ? 'Vượt quá 0.20 mg KOH/g, dầu oxy hóa tạo cặn bùn sludge!' : 'Exceeds 0.20 mg KOH/g!'}
                    </p>
                  )}
                </div>
              );
            })()}

            {/* Interfacial Tension ASTM D971 */}
            {(() => {
              const ift = data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.interfacial_tension_d971_mn_m;
              const status: 'PASS' | 'INVESTIGATE' | 'FAIL' = ift < 24 ? 'FAIL' : ift < 28 ? 'INVESTIGATE' : 'PASS';
              const pmp = evalPmpField(status, 'Table 100.4.1 (≥ 24 mN/m)');

              return (
                <div className={`p-3.5 rounded-xl border transition-all ${pmp.cardBorder} space-y-1.5`}>
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <span className="font-bold text-slate-800">
                      {lang === 'vi' ? 'Sức căng bề mặt IFT D971' : 'Interfacial Tension IFT D971'}
                    </span>
                    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] border ${pmp.badgeClass}`}>
                      {pmp.icon}
                      <span>{pmp.badgeText}</span>
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-500">
                    {lang === 'vi' ? 'Chuẩn Table 100.4.1: Tối thiểu ≥ 24 mN/m' : 'Table 100.4.1 limit: ≥ 24 mN/m'}
                  </p>
                  <div className="flex items-center gap-2 pt-1">
                    <input
                      type="number"
                      step="0.1"
                      value={ift}
                      onChange={(e) => {
                        setAnalysisResult(null);
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            service_aged_oil_sample_tests_table_100_4_1: {
                              ...data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1,
                              interfacial_tension_d971_mn_m: Number(e.target.value)
                            }
                          }
                        });
                      }}
                      className={`w-full px-2.5 py-1.5 border rounded-lg font-mono text-sm font-black transition-all ${pmp.inputClass}`}
                    />
                    <span className="font-bold text-slate-600">mN/m</span>
                  </div>
                  {status === 'FAIL' && (
                    <p className="text-[11px] text-rose-700 font-bold pt-0.5">
                      ✕ {lang === 'vi' ? 'Dưới 24 mN/m, xác nhận có cặn phân cực hòa tan!' : 'Below 24 mN/m, polar sludge confirmed!'}
                    </p>
                  )}
                </div>
              );
            })()}

            {/* Oil Power Factor @ 25°C ASTM D924 */}
            {(() => {
              const pf = data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.oil_power_factor_25c_d924_percent;
              const status: 'PASS' | 'INVESTIGATE' | 'FAIL' = pf > 1.0 ? 'FAIL' : pf > 0.5 ? 'INVESTIGATE' : 'PASS';
              const pmp = evalPmpField(status, 'Table 100.4.1 (≤ 1.0%)');

              return (
                <div className={`p-3.5 rounded-xl border transition-all ${pmp.cardBorder} space-y-1.5`}>
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <span className="font-bold text-slate-800">
                      {lang === 'vi' ? 'Power Factor dầu @ 25°C D924' : 'Oil Power Factor @ 25°C D924'}
                    </span>
                    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] border ${pmp.badgeClass}`}>
                      {pmp.icon}
                      <span>{pmp.badgeText}</span>
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-500">
                    {lang === 'vi' ? 'Chuẩn Table 100.4.1: Giới hạn ≤ 1.0%' : 'Table 100.4.1 limit: ≤ 1.0%'}
                  </p>
                  <div className="flex items-center gap-2 pt-1">
                    <input
                      type="number"
                      step="0.01"
                      value={pf}
                      onChange={(e) => {
                        setAnalysisResult(null);
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            service_aged_oil_sample_tests_table_100_4_1: {
                              ...data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1,
                              oil_power_factor_25c_d924_percent: Number(e.target.value)
                            }
                          }
                        });
                      }}
                      className={`w-full px-2.5 py-1.5 border rounded-lg font-mono text-sm font-black transition-all ${pmp.inputClass}`}
                    />
                    <span className="font-bold text-slate-600">%</span>
                  </div>
                  {status === 'FAIL' && (
                    <p className="text-[11px] text-rose-700 font-bold pt-0.5">
                      ✕ {lang === 'vi' ? 'Vượt quá 1.0%, tổn hao điện môi dầu tăng cao!' : 'Exceeds 1.0% limit!'}
                    </p>
                  )}
                </div>
              );
            })()}

            {/* Previous inspection comparative info */}
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50 space-y-2">
              <div className="font-bold text-slate-800">
                {lang === 'vi' ? 'Dữ liệu kỳ kiểm tra trước (Trend)' : 'Historical Baseline Trend'}
              </div>
              <div className="grid grid-cols-2 gap-2 text-[11px]">
                <div>
                  <span className="text-slate-500 block">Ngày trước:</span>
                  <input
                    type="date"
                    value={data.previous_test_data?.last_test_date || '2024-10-10'}
                    onChange={(e) => {
                      setAnalysisResult(null);
                      setData({
                        ...data,
                        previous_test_data: {
                          ...data.previous_test_data,
                          last_test_date: e.target.value
                        }
                      });
                    }}
                    className="w-full px-2 py-1 border rounded bg-white font-mono"
                  />
                </div>
                <div>
                  <span className="text-slate-500 block">Đánh thủng cũ:</span>
                  <input
                    type="number"
                    step="0.1"
                    value={data.previous_test_data?.last_oil_breakdown_kv || 31.0}
                    onChange={(e) => {
                      setAnalysisResult(null);
                      setData({
                        ...data,
                        previous_test_data: {
                          ...data.previous_test_data,
                          last_oil_breakdown_kv: Number(e.target.value)
                        }
                      });
                    }}
                    className="w-full px-2 py-1 border rounded bg-white font-mono font-bold"
                  />
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* TAB CONTENT 5: DISSOLVED GAS ANALYSIS DGA */}
      {activeFormTab === 'dga' && (
        <div className="bg-white rounded-2xl p-5 border border-slate-200 shadow-sm space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="font-bold text-slate-900 text-sm flex items-center gap-2">
              <Flame size={16} className="text-rose-600" />
              <span>{lang === 'vi' ? 'Phân Tích Khí Hòa Tan DGA (IEEE C57.104 & ASTM D3612)' : 'Dissolved Gas Analysis DGA (IEEE C57.104 & ASTM D3612)'}</span>
            </h3>
            <span className="text-xs px-2.5 py-0.5 rounded-full font-bold bg-rose-100 text-rose-900 border border-rose-300">
              {lang === 'vi' ? 'Đánh Giá Khí Cháy & Hồ Quang' : 'Combustible & Arcing Gases'}
            </span>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 text-xs">
            {/* Acetylene C2H2 - Critical */}
            {(() => {
              const c2h2 = data.electrical_tests.dga_test_ieee_c57_104.acetylene_c2h2_ppm;
              const status: 'PASS' | 'INVESTIGATE' | 'FAIL' = c2h2 > 1.0 ? 'FAIL' : c2h2 > 0 ? 'INVESTIGATE' : 'PASS';
              const pmp = evalPmpField(status, 'IEEE C57.104 (C2H2 = 0 ppm)');

              return (
                <div className={`p-4 rounded-xl border transition-all ${pmp.cardBorder} space-y-1.5`}>
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <span className="font-black text-rose-700 flex items-center gap-1">
                      <AlertOctagon size={16} />
                      <span>Acetylene (C2H2)</span>
                    </span>
                    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] border ${pmp.badgeClass}`}>
                      {pmp.icon}
                      <span>{pmp.badgeText}</span>
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-600">
                    {lang === 'vi' ? 'Chỉ thị Phóng điện Hồ quang > 700°C (Limit: 0 ppm)' : 'Fingerprint of High-Energy Arcing > 700°C'}
                  </p>
                  <div className="flex items-center gap-2 pt-1">
                    <input
                      type="number"
                      step="0.1"
                      value={c2h2}
                      onChange={(e) => {
                        setAnalysisResult(null);
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            dga_test_ieee_c57_104: {
                              ...data.electrical_tests.dga_test_ieee_c57_104,
                              acetylene_c2h2_ppm: Number(e.target.value)
                            }
                          }
                        });
                      }}
                      className={`w-full px-3 py-1.5 border rounded-lg font-mono text-base font-black transition-all ${pmp.inputClass}`}
                    />
                    <span className="font-bold text-rose-800">ppm</span>
                  </div>
                  {status === 'FAIL' ? (
                    <p className="text-[11px] text-rose-700 font-bold pt-1">
                      🚨 {lang === 'vi' ? 'CẢNH BÁO TỐI CẤP: Hồ quang năng lượng cao > 700°C!' : 'CRITICAL: High-energy arcing fault!'}
                    </p>
                  ) : status === 'INVESTIGATE' ? (
                    <p className="text-[11px] text-amber-800 font-bold pt-1">
                      ⚠️ {lang === 'vi' ? 'Có vết C2H2, cần lấy mẫu kiểm tra sau 48h!' : 'Trace C2H2 detected, re-sample in 48h!'}
                    </p>
                  ) : (
                    <p className="text-[11px] text-emerald-700 font-medium pt-1">
                      ✓ {lang === 'vi' ? 'An toàn 100%, không có hồ quang.' : '100% safe baseline.'}
                    </p>
                  )}
                </div>
              );
            })()}

            {/* Ethylene C2H4 */}
            {(() => {
              const c2h4 = data.electrical_tests.dga_test_ieee_c57_104.ethylene_c2h4_ppm;
              const status: 'PASS' | 'INVESTIGATE' | 'FAIL' = c2h4 > 100 ? 'FAIL' : c2h4 > 50 ? 'INVESTIGATE' : 'PASS';
              const pmp = evalPmpField(status, 'IEEE C57.104 (≤ 50 ppm)');

              return (
                <div className={`p-3.5 rounded-xl border transition-all ${pmp.cardBorder} space-y-1`}>
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <span className="font-bold text-slate-800">Ethylene (C2H4)</span>
                    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] border ${pmp.badgeClass}`}>
                      {pmp.icon}
                      <span>{pmp.badgeText}</span>
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-500">
                    {lang === 'vi' ? 'Quá nhiệt cục bộ dầu (≤ 50 ppm Cond 1)' : 'High temp oil overheating'}
                  </p>
                  <div className="flex items-center gap-2 pt-1">
                    <input
                      type="number"
                      value={c2h4}
                      onChange={(e) => {
                        setAnalysisResult(null);
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            dga_test_ieee_c57_104: {
                              ...data.electrical_tests.dga_test_ieee_c57_104,
                              ethylene_c2h4_ppm: Number(e.target.value)
                            }
                          }
                        });
                      }}
                      className={`w-full px-2.5 py-1.5 border rounded-lg font-mono font-bold transition-all ${pmp.inputClass}`}
                    />
                    <span className="font-bold text-slate-600">ppm</span>
                  </div>
                </div>
              );
            })()}

            {/* Hydrogen H2 */}
            {(() => {
              const h2 = data.electrical_tests.dga_test_ieee_c57_104.hydrogen_h2_ppm;
              const status: 'PASS' | 'INVESTIGATE' | 'FAIL' = h2 > 700 ? 'FAIL' : h2 > 100 ? 'INVESTIGATE' : 'PASS';
              const pmp = evalPmpField(status, 'IEEE C57.104 (≤ 100 ppm)');

              return (
                <div className={`p-3.5 rounded-xl border transition-all ${pmp.cardBorder} space-y-1`}>
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <span className="font-bold text-slate-800">Hydrogen (H2)</span>
                    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] border ${pmp.badgeClass}`}>
                      {pmp.icon}
                      <span>{pmp.badgeText}</span>
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-500">
                    {lang === 'vi' ? 'Phóng điện PD / quá nhiệt (≤ 100 ppm)' : 'PD or thermal breakdown'}
                  </p>
                  <div className="flex items-center gap-2 pt-1">
                    <input
                      type="number"
                      value={h2}
                      onChange={(e) => {
                        setAnalysisResult(null);
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            dga_test_ieee_c57_104: {
                              ...data.electrical_tests.dga_test_ieee_c57_104,
                              hydrogen_h2_ppm: Number(e.target.value)
                            }
                          }
                        });
                      }}
                      className={`w-full px-2.5 py-1.5 border rounded-lg font-mono font-bold transition-all ${pmp.inputClass}`}
                    />
                    <span className="font-bold text-slate-600">ppm</span>
                  </div>
                </div>
              );
            })()}

            {/* Methane CH4 */}
            {(() => {
              const ch4 = data.electrical_tests.dga_test_ieee_c57_104.methane_ch4_ppm;
              const status: 'PASS' | 'INVESTIGATE' | 'FAIL' = ch4 > 400 ? 'FAIL' : ch4 > 120 ? 'INVESTIGATE' : 'PASS';
              const pmp = evalPmpField(status, 'IEEE C57.104 (≤ 120 ppm)');

              return (
                <div className={`p-3.5 rounded-xl border transition-all ${pmp.cardBorder} space-y-1`}>
                  <div className="flex items-center justify-between gap-1 flex-wrap">
                    <span className="font-bold text-slate-800">Methane (CH4)</span>
                    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] border ${pmp.badgeClass}`}>
                      {pmp.icon}
                      <span>{pmp.badgeText}</span>
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-500">
                    {lang === 'vi' ? 'Quá nhiệt dầu nhiệt độ thấp (≤ 120 ppm)' : 'Low temp thermal fault'}
                  </p>
                  <div className="flex items-center gap-2 pt-1">
                    <input
                      type="number"
                      value={ch4}
                      onChange={(e) => {
                        setAnalysisResult(null);
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            dga_test_ieee_c57_104: {
                              ...data.electrical_tests.dga_test_ieee_c57_104,
                              methane_ch4_ppm: Number(e.target.value)
                            }
                          }
                        });
                      }}
                      className={`w-full px-2.5 py-1.5 border rounded-lg font-mono font-bold transition-all ${pmp.inputClass}`}
                    />
                    <span className="font-bold text-slate-600">ppm</span>
                  </div>
                </div>
              );
            })()}
          </div>
        </div>
      )}

          {/* BOTTOM PMP ACTION BAR BELOW FORM: SWITCH TO TEV AI STANDARD REPORT */}
          <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 rounded-2xl p-5 text-white border border-amber-500/40 shadow-xl flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
            <div className="space-y-1">
              <div className="flex items-center gap-2">
                <span className="px-2 py-0.5 rounded bg-amber-500 text-slate-950 font-black text-[10px] uppercase tracking-wider">
                  BƯỚC TIẾP THEO (CHUẨN PMP)
                </span>
                <span className="font-bold text-sm text-amber-300">
                  {lang === 'vi' ? 'Hoàn tất nhập số liệu biểu mẫu hiện trường?' : 'Finished entering field measurement parameters?'}
                </span>
              </div>
              <p className="text-xs text-slate-300">
                {lang === 'vi'
                  ? `Nhấn nút bên phải để chuyển sang xem Báo Cáo Chuẩn TEV AI 4 Phần cho thiết bị "${data.site_info.transformer_tag}" (${data.site_info.project_name}).`
                  : `Click the button on the right to view the 4-Part TEV AI Standard Report for "${data.site_info.transformer_tag}" (${data.site_info.project_name}).`}
              </p>
            </div>

            <div className="flex items-center gap-2.5 shrink-0 w-full sm:w-auto justify-end">
              <button
                type="button"
                onClick={handleSaveToCmms}
                className="px-4 py-2.5 bg-slate-800 hover:bg-slate-700 text-emerald-300 border border-emerald-500/40 rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all"
              >
                <Save size={15} />
                <span>{lang === 'vi' ? 'Lưu Dữ Liệu' : 'Save Data'}</span>
              </button>

              <button
                type="button"
                onClick={() => {
                  setAnalysisResult(null);
                  setMtsViewMode('tev-report');
                  setSavedSuccessMessage(
                    lang === 'vi'
                      ? 'Đã cập nhật và đồng bộ toàn bộ 4 phần Báo Cáo Chuẩn TEV AI theo các số liệu hiện trường mới nhất!'
                      : 'Updated all 4 sections of TEV AI Report with latest field measurements!'
                  );
                  setTimeout(() => setSavedSuccessMessage(null), 4000);
                  window.scrollTo({ top: 0, behavior: 'smooth' });
                }}
                className="px-5 py-3 bg-gradient-to-r from-amber-500 via-amber-400 to-yellow-400 hover:from-amber-400 hover:to-yellow-300 text-slate-950 font-black rounded-xl text-xs sm:text-sm flex items-center gap-2 shadow-lg shadow-amber-500/30 active:scale-95 transition-all"
              >
                <Sparkles size={16} />
                <span>{lang === 'vi' ? 'Chuyển Qua Xem Báo Cáo Chuẩn TEV AI →' : 'Switch to TEV AI Standard Report →'}</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* PROFESSIONAL AI EVALUATION RESULT & TEV BRANDED CERTIFICATE SECTION       */}
      {/* ========================================================================= */}
      {mtsViewMode === 'tev-report' && (
      <div className="bg-white rounded-2xl border-2 border-slate-800 shadow-xl overflow-hidden animate-in fade-in duration-200">
        {/* TEV Official Certificate Top Header */}
        <div className="bg-gradient-to-r from-slate-950 via-slate-900 to-blue-950 p-6 sm:p-7 text-white border-b-2 border-amber-500 space-y-4">
          {/* Row 1: TEV Brand, Category & Standard Badge */}
          <div className="flex items-center justify-between gap-4 flex-wrap">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-amber-500 text-slate-950 flex items-center justify-center font-black text-xl shadow-md shrink-0">
                TEV
              </div>
              <div className="flex items-center gap-2.5 flex-wrap">
                <span className="text-xs sm:text-sm font-black uppercase tracking-wider text-amber-400">
                  TEV FLATFORM AI • FIELD SERVICE EVALUATION REPORT
                </span>
                <span className="px-2.5 py-0.5 rounded text-[11px] font-black uppercase tracking-wider bg-white/10 text-slate-200 border border-white/20">
                  ANSI/NETA MTS-2023 Mục 7.2.2
                </span>
              </div>
            </div>

            <div className="hidden sm:flex items-center gap-1.5 px-3 py-1 bg-amber-500/20 border border-amber-400/30 rounded-full text-xs text-amber-300 font-bold">
              <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
              <span>Báo Cáo Bảo Trì NETA MTS</span>
            </div>
          </div>

          {/* Row 2: Full-Width Horizontal Equipment Title & Subtitle */}
          <div className="pt-2.5 border-t border-slate-800/80">
            <h2 className="text-xl sm:text-2xl lg:text-3xl font-black text-white tracking-tight leading-snug">
              {lang === 'vi'
                ? 'Máy Biến Áp Ngâm Chất Lỏng / Dầu (Liquid-Filled Transformers)'
                : 'Liquid-Filled Transformers Maintenance Evaluation'}
            </h2>
            <p className="text-xs sm:text-sm text-slate-300 mt-1 max-w-5xl leading-relaxed font-medium">
              {lang === 'vi'
                ? `Biên bản đối soát số: TEV-MTS-${data.site_info.transformer_tag}-${data.site_info.test_date.replace(/-/g, '')} • Dự án: ${data.site_info.project_name} • Kỹ sư FSE: ${data.site_info.fse_name}`
                : `Official Report No: TEV-MTS-${data.site_info.transformer_tag}-${data.site_info.test_date.replace(/-/g, '')} • Substation: ${data.site_info.project_name} • FSE: ${data.site_info.fse_name}`}
            </p>
          </div>

          {/* Row 3: Dedicated Action Toolbar Beneath Equipment Title */}
          <div className="pt-3 border-t border-slate-800 flex items-center gap-2.5 flex-wrap">
            {/* Language Switcher */}
            <div className="flex items-center bg-slate-950 p-1 rounded-xl border border-slate-800">
              <Globe size={13} className="text-slate-400 ml-1.5 mr-1" />
              <button
                type="button"
                onClick={() => setLang('vi')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                  lang === 'vi' ? 'bg-amber-500 text-slate-950 font-black shadow-xs' : 'text-slate-400 hover:text-white'
                }`}
              >
                🇻🇳 Tiếng Việt
              </button>
              <button
                type="button"
                onClick={() => setLang('en')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                  lang === 'en' ? 'bg-amber-500 text-slate-950 font-black shadow-xs' : 'text-slate-400 hover:text-white'
                }`}
              >
                🇺🇸 English
              </button>
            </div>

            <button
              type="button"
              onClick={() => {
                setMtsViewMode('form');
                window.scrollTo({ top: 0, behavior: 'smooth' });
              }}
              className="px-3.5 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-xs font-black flex items-center gap-1.5 shadow-md shadow-blue-600/30 transition-all border border-blue-400"
              title={lang === 'vi' ? 'Quay lại biểu mẫu nhập thông số FSE' : 'Switch to Field Entry Form'}
            >
              <Layers size={14} className="text-amber-300" />
              <span>{lang === 'vi' ? 'Biểu Mẫu Nhập Liệu 📝' : 'Field Entry Form 📝'}</span>
            </button>

            <button
              type="button"
              onClick={() => setShowPrintModal(true)}
              className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md shadow-blue-600/30 transition-all"
            >
              <Printer size={15} />
              <span>{lang === 'vi' ? 'In / Xuất PDF' : 'Print / Export PDF'}</span>
            </button>

            <button
              type="button"
              onClick={handleCopyMarkdown}
              className="px-3.5 py-2 bg-slate-800 hover:bg-slate-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 border border-slate-700 transition-all shadow-sm"
              title="Sao chép toàn bộ báo cáo Markdown"
            >
              {copied ? <Check size={14} className="text-emerald-400" /> : <Copy size={14} />}
              <span>{copied ? (lang === 'vi' ? 'Đã chép!' : 'Copied!') : (lang === 'vi' ? 'Sao chép MD' : 'Copy MD')}</span>
            </button>

            <button
              type="button"
              onClick={handleSaveToCmms}
              className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md shadow-emerald-600/30 transition-all"
            >
              <Save size={15} />
              <span>{lang === 'vi' ? 'Lưu Báo Cáo CMMS' : 'Save to CMMS'}</span>
            </button>

            <button
              type="button"
              onClick={handleRunAiAnalysis}
              disabled={isAnalyzing}
              className="px-4 py-2 bg-gradient-to-r from-amber-500 to-yellow-500 hover:from-amber-400 hover:to-yellow-400 text-slate-950 rounded-xl text-xs font-black flex items-center gap-1.5 shadow-md shadow-amber-500/20 transition-all"
            >
              <Sparkles size={14} />
              <span>{isAnalyzing ? (lang === 'vi' ? 'Đang phân tích...' : 'Analyzing...') : (lang === 'vi' ? 'Chạy AI NETA Đối Soát' : 'Run AI Evaluation')}</span>
            </button>
          </div>
        </div>

        {/* SECTION 1: OVERALL STATUS */}
        <div className={`p-6 border-b text-sm ${
          currentEvaluation.overall_status === 'FAIL'
            ? 'bg-rose-50/70 border-rose-200'
            : currentEvaluation.overall_status === 'INVESTIGATE'
            ? 'bg-amber-50/70 border-amber-200'
            : 'bg-emerald-50/70 border-emerald-200'
        }`}>
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div>
              <span className="text-xs uppercase font-extrabold tracking-wider text-slate-500 block mb-1">
                {lang === 'vi' ? '1. TRẠNG THÁI TỔNG QUAN (Overall Status)' : '1. OVERALL STATUS'}
              </span>
              <div className="flex items-center gap-3">
                <span className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl font-black text-sm tracking-wide uppercase shadow-sm ${
                  currentEvaluation.overall_status === 'FAIL'
                    ? 'bg-rose-600 text-white'
                    : currentEvaluation.overall_status === 'INVESTIGATE'
                    ? 'bg-amber-600 text-white'
                    : 'bg-emerald-600 text-white'
                }`}>
                  {currentEvaluation.overall_status === 'FAIL' && <XCircle size={16} />}
                  {currentEvaluation.overall_status === 'INVESTIGATE' && <AlertTriangle size={16} />}
                  {currentEvaluation.overall_status === 'PASS' && <CheckCircle2 size={16} />}
                  <span>{lang === 'vi' ? currentEvaluation.status_title_vi : currentEvaluation.status_title_en}</span>
                </span>
              </div>
            </div>

            <div className="text-xs space-y-1 bg-white/80 p-3 rounded-xl border border-slate-200/80">
              <div><strong className="text-slate-600">{lang === 'vi' ? 'Mã máy biến áp:' : 'Transformer Tag:'}</strong> <span className="font-mono font-bold text-slate-900">{data.site_info.transformer_tag}</span></div>
              <div><strong className="text-slate-600">{lang === 'vi' ? 'Dự án / Nhà máy:' : 'Substation / Plant:'}</strong> <span className="font-semibold text-slate-900">{data.site_info.project_name}</span></div>
              <div><strong className="text-slate-600">{lang === 'vi' ? 'Kỹ sư FSE:' : 'Field Engineer:'}</strong> <span className="font-semibold text-slate-900">{data.site_info.fse_name}</span></div>
            </div>
          </div>
        </div>

        {/* SECTION 2: DETAILED EVALUATION TABLE */}
        <div className="p-6 border-b border-slate-200">
          <div className="flex items-center justify-between mb-3">
            <h3 className="font-black text-slate-900 text-sm uppercase tracking-wide flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-amber-500"></span>
              <span>
                {lang === 'vi'
                  ? '2. BẢNG PHÂN TÍCH CHI TIẾT MÁY BIẾN ÁP DẦU BẢO TRÌ (Detailed Evaluation Table)'
                  : '2. DETAILED EVALUATION TABLE (NETA MTS-2023 Section 7.2.2)'}
              </span>
            </h3>
            <span className="text-xs text-slate-500 font-semibold hidden sm:inline">
              ANSI/NETA MTS-2023 Table 100.3, 100.4.1, 100.5, 100.12 & IEEE C57.104
            </span>
          </div>

          <div className="overflow-x-auto rounded-xl border border-slate-200">
            <table className="w-full text-left text-xs border-collapse min-w-[1050px]">
              <thead>
                <tr className="bg-slate-100 text-slate-800 font-bold border-b border-slate-200 text-xs">
                  <th className="p-4 w-[42%] min-w-[380px]">{lang === 'vi' ? 'Hạng mục kiểm tra & thử nghiệm' : 'Inspection / Test Item'}</th>
                  <th className="p-4 w-[16%] min-w-[150px]">{lang === 'vi' ? 'Giá trị đo thực tế' : 'Measured Site Value'}</th>
                  <th className="p-4 w-[22%] min-w-[210px]">{lang === 'vi' ? 'Tiêu chuẩn NETA MTS-2023' : 'NETA MTS-2023 Standard'}</th>
                  <th className="p-4 w-[11%] min-w-[110px]">{lang === 'vi' ? 'Tham chiếu kỹ thuật' : 'Reference Standard'}</th>
                  <th className="p-4 w-[9%] min-w-[95px] text-center">{lang === 'vi' ? 'Kết luận' : 'Status'}</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {currentEvaluation.evaluations.items.map((item, idx) => (
                  <tr key={idx} className={item.status === 'FAIL' ? 'bg-rose-50/40 hover:bg-rose-50/60' : item.status === 'INVESTIGATE' ? 'bg-amber-50/40 hover:bg-amber-50/60' : 'hover:bg-slate-50/80 transition-colors'}>
                    <td className="p-4 font-bold text-slate-900 align-top">
                      <div className="text-xs sm:text-sm text-slate-900 font-bold leading-snug break-words">{lang === 'vi' ? item.item : item.itemEn}</div>
                      <div className="text-[11px] text-slate-600 font-normal pt-1.5 leading-relaxed break-words">
                        {lang === 'vi' ? item.detail : item.detailEn}
                      </div>
                    </td>
                    <td className="p-4 font-mono font-bold text-slate-900 whitespace-nowrap align-top">
                      <span className="bg-slate-100 px-2 py-1 rounded border border-slate-200/80 inline-block">
                        {item.measured}
                      </span>
                    </td>
                    <td className="p-4 text-slate-800 align-top leading-relaxed break-words text-xs">
                      {lang === 'vi' ? item.standard : item.standardEn}
                    </td>
                    <td className="p-4 text-slate-600 font-mono text-[11px] align-top whitespace-nowrap">
                      <span className="bg-blue-50/60 text-blue-900 px-2 py-0.5 rounded border border-blue-200/50">
                        {item.reference}
                      </span>
                    </td>
                    <td className="p-4 text-center whitespace-nowrap align-top">
                      <span className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full font-black text-[11px] uppercase shadow-2xs ${
                        item.status === 'PASS'
                          ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                          : item.status === 'INVESTIGATE'
                          ? 'bg-amber-100 text-amber-800 border border-amber-300'
                          : 'bg-rose-100 text-rose-800 border border-rose-300'
                      }`}>
                        {item.status === 'PASS' && <CheckCircle2 size={13} />}
                        {item.status === 'INVESTIGATE' && <AlertTriangle size={13} />}
                        {item.status === 'FAIL' && <XCircle size={13} />}
                        <span>{item.status}</span>
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* SECTION 3: SERVICE-AGED OIL QUALITY, INSULATION & DGA ANOMALY RISK */}
        <div className="p-6 border-b border-slate-200 bg-slate-50/50 space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="font-black text-slate-900 text-sm uppercase tracking-wide flex items-center gap-2">
              <span className={`w-2.5 h-2.5 rounded-full ${
                currentEvaluation.overall_status === 'PASS'
                  ? 'bg-emerald-500'
                  : currentEvaluation.overall_status === 'INVESTIGATE'
                  ? 'bg-amber-500'
                  : 'bg-rose-500'
              }`}></span>
              <span>
                {lang === 'vi'
                  ? '3. PHÂN TÍCH CHẤT LƯỢNG DẦU LÃO HÓA, CÁCH ĐIỆN & KHÍ DGA (Oil Quality, Insulation & DGA Risk)'
                  : '3. SERVICE-AGED OIL QUALITY, INSULATION & DGA ANOMALY RISK'}
              </span>
            </h3>
            <span className={`px-2.5 py-0.5 rounded-full text-xs font-black uppercase ${
              currentEvaluation.overall_status === 'PASS'
                ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                : currentEvaluation.overall_status === 'INVESTIGATE'
                ? 'bg-amber-100 text-amber-900 border border-amber-300'
                : 'bg-rose-100 text-rose-800 border border-rose-300'
            }`}>
              {currentEvaluation.overall_status}
            </span>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
            {/* Box 3.1: DGA and Arcing Fault */}
            <div className={`p-4 rounded-xl border space-y-2 shadow-xs ${
              currentEvaluation.evaluations.arcing_c2h2_critical
                ? 'bg-rose-50/80 border-rose-300'
                : 'bg-white border-slate-200'
            }`}>
              <div className={`font-bold flex items-center gap-1.5 text-xs ${
                currentEvaluation.evaluations.arcing_c2h2_critical ? 'text-rose-700' : 'text-slate-800'
              }`}>
                <Flame size={16} />
                <span>
                  {lang === 'vi'
                    ? 'Phân tích Khí Hòa tan DGA & Hồ quang (IEEE C57.104 & ASTM D3612)'
                    : 'DGA Dissolved Gas Analysis & Arcing Risk (IEEE C57.104)'}
                </span>
              </div>
              <ul className="space-y-1.5 text-slate-700 leading-relaxed list-disc pl-4">
                {currentEvaluation.evaluations.arcing_c2h2_critical ? (
                  <>
                    <li className="text-rose-900 font-medium">
                      <strong>Khí Acetylene (C2H2 = {data.electrical_tests.dga_test_ieee_c57_104.acetylene_c2h2_ppm} ppm):</strong>{' '}
                      {lang === 'vi'
                        ? 'CẢNH BÁO TỐI CẤP: Sự phát sinh khí C2H2 (> 1.0 ppm) là chỉ thị trực tiếp cho thấy đang có Sự cố Phóng điện Hồ quang Nhiệt độ Cao (> 700°C) xảy ra bên trong máy biến áp!'
                        : 'CRITICAL ALERT: Emergence of C2H2 (> 1.0 ppm) is an undeniable fingerprint of High-Energy Arcing (> 700°C) inside the transformer!'}
                    </li>
                    <li>
                      <strong>Ethylene (C2H4 = {data.electrical_tests.dga_test_ieee_c57_104.ethylene_c2h4_ppm} ppm) & Hydrogen (H2 = {data.electrical_tests.dga_test_ieee_c57_104.hydrogen_h2_ppm} ppm):</strong>{' '}
                      {lang === 'vi'
                        ? 'Xác nhận kèm theo vùng quá nhiệt cục bộ bẻ gãy mạch hydrocacbon dầu.'
                        : 'Confirms accompanying severe localized thermal decomposition of insulating fluid.'}
                    </li>
                  </>
                ) : (
                  <>
                    <li className="text-emerald-900 font-medium">
                      <strong>Khí Acetylene (C2H2 = {data.electrical_tests.dga_test_ieee_c57_104.acetylene_c2h2_ppm || 0} ppm):</strong>{' '}
                      {lang === 'vi'
                        ? 'An toàn 100%. Không phát hiện khí phóng điện hồ quang nhiệt độ cao.'
                        : '100% Safe. Zero high-energy thermal arcing gases detected.'}
                    </li>
                    <li>
                      <strong>Tổng khí cháy DGA (TDCG):</strong> H2 = {data.electrical_tests.dga_test_ieee_c57_104.hydrogen_h2_ppm} ppm, CH4 = {data.electrical_tests.dga_test_ieee_c57_104.methane_ch4_ppm} ppm, C2H4 = {data.electrical_tests.dga_test_ieee_c57_104.ethylene_c2h4_ppm} ppm. Nằm an toàn trong dải IEEE C57.104 Condition 1.
                    </li>
                  </>
                )}
              </ul>
            </div>

            {/* Box 3.2: Oil aging ASTM */}
            <div className={`p-4 rounded-xl border space-y-2 shadow-xs ${
              currentEvaluation.evaluations.dielectric_breakdown_failed || currentEvaluation.evaluations.water_content_exceeded || currentEvaluation.evaluations.acid_number_exceeded
                ? 'bg-rose-50/80 border-rose-300'
                : 'bg-white border-slate-200'
            }`}>
              <div className={`font-bold flex items-center gap-1.5 text-xs ${
                currentEvaluation.evaluations.dielectric_breakdown_failed || currentEvaluation.evaluations.water_content_exceeded
                  ? 'text-rose-700'
                  : 'text-amber-800'
              }`}>
                <Droplets size={16} />
                <span>
                  {lang === 'vi'
                    ? 'Chất lượng Dầu cách điện Lão hóa (NETA MTS Table 100.4.1)'
                    : 'Service-Aged Fluid Quality (NETA Table 100.4.1)'}
                </span>
              </div>
              <ul className="space-y-1.5 text-slate-700 leading-relaxed list-disc pl-4">
                <li>
                  <strong>{lang === 'vi' ? 'Điện áp đánh thủng D1816:' : 'Breakdown D1816:'}</strong>{' '}
                  Đạt {data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.dielectric_breakdown_d1816_1mm_kv} kV{' '}
                  {currentEvaluation.evaluations.dielectric_breakdown_failed
                    ? (lang === 'vi' ? `(Thấp hơn ngưỡng tối thiểu ${liveStats.breakdownThreshold} kV) -> 🔴 FAIL` : `(Below ${liveStats.breakdownThreshold} kV limit) -> 🔴 FAIL`)
                    : (lang === 'vi' ? `(Đạt ngưỡng tối thiểu ≥ ${liveStats.breakdownThreshold} kV) -> 🟢 PASS` : `(Meets limit ≥ ${liveStats.breakdownThreshold} kV) -> 🟢 PASS`)}
                </li>
                <li>
                  <strong>{lang === 'vi' ? 'Hàm lượng nước D1533:' : 'Moisture D1533:'}</strong>{' '}
                  Đạt {data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.water_content_d1533_ppm} ppm{' '}
                  {currentEvaluation.evaluations.water_content_exceeded
                    ? (lang === 'vi' ? `(Vượt quá ngưỡng tối đa ${liveStats.waterThreshold} ppm) -> 🔴 FAIL (Dầu ngấm ẩm)` : `(Exceeds ${liveStats.waterThreshold} ppm limit) -> 🔴 FAIL`)
                    : (lang === 'vi' ? `(Đạt giới hạn cho phép ≤ ${liveStats.waterThreshold} ppm) -> 🟢 PASS` : `(Meets limit ≤ ${liveStats.waterThreshold} ppm) -> 🟢 PASS`)}
                </li>
                <li>
                  <strong>{lang === 'vi' ? 'Axit D974 & Sức căng IFT D971:' : 'Acid D974 & IFT D971:'}</strong>{' '}
                  Axit = {data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.acid_number_d974_mg_koh_g} mg KOH/g (chuẩn ≤ 0.20), IFT = {data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.interfacial_tension_d971_mn_m} mN/m (chuẩn ≥ 24){' '}
                  {currentEvaluation.evaluations.acid_number_exceeded || currentEvaluation.evaluations.ift_below_minimum
                    ? '-> 🔴 FAIL (Dầu oxy hóa và sinh cặn bùn)'
                    : '-> 🟢 PASS (Dầu ổn định hóa học tốt)'}
                </li>
              </ul>
            </div>

            {/* Box 3.3: Insulation & Winding Integrity */}
            <div className="bg-white p-4 rounded-xl border border-slate-200 space-y-2 shadow-xs">
              <div className="font-bold text-slate-900 flex items-center gap-1.5 text-xs text-emerald-700">
                <ShieldCheck size={16} />
                <span>
                  {lang === 'vi'
                    ? 'Thị giác, Cơ khí & Điện trở Cách điện Cuộn dây'
                    : 'Visual, Mechanical & Winding Insulation'}
                </span>
              </div>
              <ul className="space-y-1.5 text-slate-700 leading-relaxed list-disc pl-4">
                <li>
                  {lang === 'vi'
                    ? 'Mức chất lỏng, áp suất đệm khí Nitơ dương, lực siết bu-lông đạt chuẩn NETA MTS Section 7.2.2.A.'
                    : 'Liquid level, positive nitrogen gas blanket, bolt torque conform to NETA MTS Section 7.2.2.A.'}
                </li>
                <li>
                  {lang === 'vi'
                    ? `Cách điện cuộn dây quy đổi 20°C: Cao-Đất: ${data.electrical_tests.insulation_resistance_20c_megohms.pri_to_ground_5000v} MΩ, Hạ-Đất: ${data.electrical_tests.insulation_resistance_20c_megohms.sec_to_ground_2500v} MΩ. PI = ${data.electrical_tests.insulation_resistance_20c_megohms.calculated_pi_value} (${data.electrical_tests.insulation_resistance_20c_megohms.calculated_pi_value >= 1.0 ? '≥ 1.0 -> 🟢 PASS' : '< 1.0 -> 🔴 FAIL'}).`
                    : `Winding IR: Pri-Gnd: ${data.electrical_tests.insulation_resistance_20c_megohms.pri_to_ground_5000v} MΩ. Polarization Index PI = ${data.electrical_tests.insulation_resistance_20c_megohms.calculated_pi_value} (${data.electrical_tests.insulation_resistance_20c_megohms.calculated_pi_value >= 1.0 ? '≥ 1.0 -> 🟢 PASS' : '< 1.0 -> 🔴 FAIL'}).`}
                </li>
              </ul>
            </div>

            {/* Box 3.4: Turns Ratio & Core IR */}
            <div className="bg-white p-4 rounded-xl border border-slate-200 space-y-2 shadow-xs">
              <div className="font-bold text-slate-900 flex items-center gap-1.5 text-xs text-blue-700">
                <Gauge size={16} />
                <span>
                  {lang === 'vi'
                    ? 'Tỷ số Biến áp, Power Factor Cuộn Dây & Cách điện Lõi'
                    : 'Turns Ratio, Winding PF & Core Insulation'}
                </span>
              </div>
              <ul className="space-y-1.5 text-slate-700 leading-relaxed list-disc pl-4">
                <li>
                  {lang === 'vi'
                    ? `Sai số tỷ số biến áp ${data.electrical_tests.turns_ratio_all_taps.max_error_percent}% (${data.electrical_tests.turns_ratio_all_taps.max_error_percent <= 0.5 ? '≤ 0.5% -> 🟢 PASS' : '> 0.5% -> 🔴 FAIL'}).`
                    : `Turns-ratio error is ${data.electrical_tests.turns_ratio_all_taps.max_error_percent}% (${data.electrical_tests.turns_ratio_all_taps.max_error_percent <= 0.5 ? '≤ 0.5% -> 🟢 PASS' : '> 0.5% -> 🔴 FAIL'}).`}
                </li>
                <li>
                  {lang === 'vi'
                    ? `Power Factor cuộn dây ở 20°C: ${data.electrical_tests.winding_insulation_power_factor_percent_20c}% (${data.electrical_tests.winding_insulation_power_factor_percent_20c <= 1.0 ? '≤ 1.0% Table 100.3 -> 🟢 PASS' : '> 1.0% -> 🔴 FAIL'}).`
                    : `Winding Power Factor at 20°C: ${data.electrical_tests.winding_insulation_power_factor_percent_20c}% (${data.electrical_tests.winding_insulation_power_factor_percent_20c <= 1.0 ? '≤ 1.0% Table 100.3 -> 🟢 PASS' : '> 1.0% -> 🔴 FAIL'}).`}
                </li>
                <li>
                  {lang === 'vi'
                    ? `Điện trở cách điện lõi thép: ${data.electrical_tests.core_insulation_resistance_500v_megohms ?? 350} MΩ (≥ 100 MΩ -> 🟢 PASS).`
                    : `Core insulation resistance: ${data.electrical_tests.core_insulation_resistance_500v_megohms ?? 350} MΩ (≥ 100 MΩ -> 🟢 PASS).`}
                </li>
              </ul>
            </div>
          </div>
        </div>

        {/* SECTION 4: ACTIONABLE RECOMMENDATIONS */}
        <div className="p-6 bg-slate-900 text-white space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="font-black text-amber-400 text-sm uppercase tracking-wide flex items-center gap-2">
              <span className="w-2 h-2 rounded-full bg-amber-400"></span>
              <span>
                {lang === 'vi'
                  ? '4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO KỸ SƯ FSE (Actionable Recommendations)'
                  : '4. ACTIONABLE RECOMMENDATIONS FOR FSE'}
              </span>
            </h3>
            <span className={`text-xs px-2.5 py-0.5 rounded-full font-bold ${
              currentEvaluation.overall_status === 'PASS'
                ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-400/30'
                : currentEvaluation.overall_status === 'INVESTIGATE'
                ? 'bg-amber-500/20 text-amber-300 border border-amber-400/30'
                : 'bg-rose-500/20 text-rose-300 border border-rose-400/30'
            }`}>
              {currentEvaluation.overall_status === 'PASS'
                ? (lang === 'vi' ? 'VẬN HÀNH AN TOÀN' : 'SAFE OPERATION')
                : currentEvaluation.overall_status === 'INVESTIGATE'
                ? (lang === 'vi' ? 'THEO DÕI NÂNG CAO' : 'MONITORING REQUIRED')
                : (lang === 'vi' ? 'HÀNH ĐỘNG CẤP BÁCH' : 'IMMEDIATE DIRECTIVES')}
            </span>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
            {currentEvaluation.overall_status === 'PASS' ? (
              <>
                <div className="p-4 rounded-xl bg-slate-800/80 border border-emerald-500/40 space-y-2">
                  <div className="font-black text-emerald-400 flex items-center gap-2 text-xs">
                    <CheckCircle2 size={16} />
                    <span>1. {lang === 'vi' ? 'VẬN HÀNH BÌNH THƯỜNG LIÊN TỤC' : 'NORMAL CONTINUOUS OPERATION'}</span>
                  </div>
                  <p className="text-slate-300 leading-relaxed">
                    {lang === 'vi'
                      ? `Máy biến áp ${data.site_info.transformer_tag} đạt 100% các tiêu chí kiểm định NETA MTS-2023. Đủ điều kiện tiếp tục vận hành mang tải an toàn.`
                      : `Transformer ${data.site_info.transformer_tag} meets 100% NETA MTS-2023 maintenance criteria. Authorized for normal operation.`}
                  </p>
                </div>
                <div className="p-4 rounded-xl bg-slate-800/80 border border-emerald-500/40 space-y-2">
                  <div className="font-black text-emerald-400 flex items-center gap-2 text-xs">
                    <Calendar size={16} />
                    <span>2. {lang === 'vi' ? 'KẾ HOẠCH BẢO TRÌ ĐỊNH KỲ (12 THÁNG)' : 'ANNUAL PM SCHEDULE (12 MONTHS)'}</span>
                  </div>
                  <p className="text-slate-300 leading-relaxed">
                    {lang === 'vi'
                      ? 'Duy trì kế hoạch bảo dưỡng định kỳ 12 tháng/lần cho thử nghiệm dầu và phân tích DGA tiếp theo.'
                      : 'Maintain annual preventative maintenance testing schedule for oil quality and DGA screening.'}
                  </p>
                </div>
                <div className="p-4 rounded-xl bg-slate-800/80 border border-blue-500/40 space-y-2">
                  <div className="font-black text-blue-400 flex items-center gap-2 text-xs">
                    <ShieldCheck size={16} />
                    <span>3. {lang === 'vi' ? 'KIỂM TRA BÌNH THỞ SILICAGEL' : 'INSPECT DEHYDRATING BREATHER'}</span>
                  </div>
                  <p className="text-slate-300 leading-relaxed">
                    {lang === 'vi'
                      ? 'Kiểm tra màu sắc hạt hút ẩm bình thở định kỳ hàng quý và đảm bảo đệm khí Nitơ duy trì áp suất dương (0.5 - 5.0 psi).'
                      : 'Inspect dehydrating breather silica gel quarterly and verify positive nitrogen pressure cushion (0.5 - 5.0 psi).'}
                  </p>
                </div>
                <div className="p-4 rounded-xl bg-slate-800/80 border border-blue-500/40 space-y-2">
                  <div className="font-black text-blue-400 flex items-center gap-2 text-xs">
                    <Sliders size={16} />
                    <span>4. {lang === 'vi' ? 'THEO DÕI NHIỆT ĐỘ CUỘN DÂY & DẦU' : 'WINDING & OIL TEMPERATURE TRENDING'}</span>
                  </div>
                  <p className="text-slate-300 leading-relaxed">
                    {lang === 'vi'
                      ? 'Giám sát đồng hồ đo nhiệt độ cuộn dây và dầu trên rơ le bảo vệ, ghi nhận dữ liệu định kỳ vào nhật ký vận hành trạm.'
                      : 'Monitor top oil and winding temperature indicators, recording operational logs regularly into SCADA/CMMS.'}
                  </p>
                </div>
              </>
            ) : currentEvaluation.overall_status === 'INVESTIGATE' ? (
              <>
                <div className="p-4 rounded-xl bg-slate-800/80 border border-amber-500/40 space-y-2">
                  <div className="font-black text-amber-400 flex items-center gap-2 text-xs">
                    <Wrench size={16} />
                    <span>1. {lang === 'vi' ? 'SIẾT LẠI TIẾP XÚC BU-LÔNG' : 'RE-TORQUE BOLTED CONNECTIONS'}</span>
                  </div>
                  <p className="text-slate-300 leading-relaxed">
                    {lang === 'vi'
                      ? 'Kiểm tra vệ sinh bề mặt tiếp xúc và siết lại bu-lông đầu cốt bushing theo lực mô-men quy định tại Table 100.12.'
                      : 'Clean contact surfaces and re-torque bolted connection hardware per Table 100.12 standards.'}
                  </p>
                </div>
                <div className="p-4 rounded-xl bg-slate-800/80 border border-amber-500/40 space-y-2">
                  <div className="font-black text-amber-400 flex items-center gap-2 text-xs">
                    <RefreshCw size={16} />
                    <span>2. {lang === 'vi' ? 'RÚT NGẮN CHU KỲ LẤY MẪU (6 THÁNG)' : 'INTERMEDIATE SAMPLING (6 MONTHS)'}</span>
                  </div>
                  <p className="text-slate-300 leading-relaxed">
                    {lang === 'vi'
                      ? 'Lấy mẫu dầu kiểm tra lại độ ẩm và DGA sau 6 tháng để theo dõi tốc độ biến thiên trước khi đưa ra quyết định xử lý.'
                      : 'Take secondary oil sample within 6 months to track moisture and gas generation trends.'}
                  </p>
                </div>
                <div className="p-4 rounded-xl bg-slate-800/80 border border-blue-500/40 space-y-2">
                  <div className="font-black text-blue-400 flex items-center gap-2 text-xs">
                    <Droplets size={16} />
                    <span>3. {lang === 'vi' ? 'THAY HẠT HÚT ẨM SILICAGEL' : 'REPLACE SILICA GEL BREATHER'}</span>
                  </div>
                  <p className="text-slate-300 leading-relaxed">
                    {lang === 'vi'
                      ? 'Thay mới hạt hút ẩm silicagel nếu có dấu hiệu đổi màu quá 2/3 để ngăn ngừa độ ẩm thâm nhập vào dầu.'
                      : 'Replace silica gel desiccant charge if over 2/3 saturated to prevent ambient moisture contamination.'}
                  </p>
                </div>
                <div className="p-4 rounded-xl bg-slate-800/80 border border-blue-500/40 space-y-2">
                  <div className="font-black text-blue-400 flex items-center gap-2 text-xs">
                    <Gauge size={16} />
                    <span>4. {lang === 'vi' ? 'CHỤP ẢNH NHIỆT HỒNG NGOẠI (IR)' : 'THERMOGRAPHIC SURVEY'}</span>
                  </div>
                  <p className="text-slate-300 leading-relaxed">
                    {lang === 'vi'
                      ? 'Thực hiện quét ảnh nhiệt hồng ngoại các điểm đấu nối đầu cực để phát hiện sớm các điểm quá nhiệt cục bộ.'
                      : 'Perform infrared thermographic scan on bushing terminals to detect localized hot spots under load.'}
                  </p>
                </div>
              </>
            ) : (
              <>
                <div className="p-4 rounded-xl bg-slate-800/80 border border-rose-500/40 space-y-2">
                  <div className="font-black text-rose-400 flex items-center gap-2 text-xs">
                    <AlertOctagon size={16} />
                    <span>1. {lang === 'vi' ? 'CẢNH BÁO TỐI CẤP: KHÔNG MANG TẢI NẶNG / TÁCH CÔ LẬP' : 'CRITICAL: DO NOT OPERATE HEAVY / ISOLATE NOW'}</span>
                  </div>
                  <p className="text-slate-300 leading-relaxed">
                    {lang === 'vi'
                      ? `Khuyên đơn vị vận hành không tiếp tục cho máy biến áp ${data.site_info.transformer_tag} mang tải nặng; ${currentEvaluation.evaluations.arcing_c2h2_critical ? 'lập kế hoạch tách cô lập kiểm tra ngay do phát hiện phóng điện hồ quang.' : 'lập kế hoạch lọc xử lý dầu khẩn cấp.'}`
                      : `Advise plant operations not to operate transformer ${data.site_info.transformer_tag} under heavy load; execute planned isolation.`}
                  </p>
                </div>

                <div className="p-4 rounded-xl bg-slate-800/80 border border-amber-500/40 space-y-2">
                  <div className="font-black text-amber-400 flex items-center gap-2 text-xs">
                    <RefreshCw size={16} />
                    <span>2. {currentEvaluation.evaluations.arcing_c2h2_critical ? (lang === 'vi' ? 'LẤY MẪU DGA ĐỐI SOÁT LẠI SAU 48 GIỜ' : 'RE-SAMPLE DGA WITHIN 48 HOURS') : (lang === 'vi' ? 'LỌC DẦU HÚT CHÂN KHÔNG KHỬ NƯỚC' : 'VACUUM OIL DEHYDRATION')}</span>
                  </div>
                  <p className="text-slate-300 leading-relaxed">
                    {currentEvaluation.evaluations.arcing_c2h2_critical
                      ? (lang === 'vi'
                        ? 'Tiến hành lấy mẫu dầu DGA đối soát lại sau 48 giờ để đánh giá tốc độ sinh khí (gas generation rate per day) theo IEEE C57.104.'
                        : 'Take secondary oil sample for DGA cross-check within 48 hours to determine gas generation rates (ppm/day) per IEEE C57.104.')
                      : (lang === 'vi'
                        ? 'Thực hiện lọc sấy dầu tuần hoàn hút chân không khử nước và khí hòa tan để phục hồi độ bền điện môi.'
                        : 'Execute on-site vacuum dehydration and degassing to restore dielectric breakdown voltage.')}
                  </p>
                </div>

                <div className="p-4 rounded-xl bg-slate-800/80 border border-blue-500/40 space-y-2">
                  <div className="font-black text-blue-400 flex items-center gap-2 text-xs">
                    <Droplets size={16} />
                    <span>3. {lang === 'vi' ? 'LỌC DẦU TÁI CHẾ HOÀN NGUYÊN (RECLAMATION)' : 'FULL OIL RECLAMATION & DEHYDRATION'}</span>
                  </div>
                  <p className="text-slate-300 leading-relaxed">
                    {lang === 'vi'
                      ? "Thực hiện lọc dầu tái chế hoàn nguyên bằng đất lọc Fuller's Earth kết hợp hút chân không để khôi phục chỉ số Axit, IFT, Water content và Dielectric breakdown về dải Bảng Table 100.4.1."
                      : "Execute full fluid reclamation using Fuller's Earth chemical filtering combined with high-vacuum dehydration/degassing."}
                  </p>
                </div>

                <div className="p-4 rounded-xl bg-slate-800/80 border border-emerald-500/40 space-y-2">
                  <div className="font-black text-emerald-400 flex items-center gap-2 text-xs">
                    <Zap size={16} />
                    <span>4. {lang === 'vi' ? 'ĐO PHÓNG ĐIỆN CỤC BỘ (PD) & NỘI SOI TANK' : 'PARTIAL DISCHARGE & TANK BORESCOPE'}</span>
                  </div>
                  <p className="text-slate-300 leading-relaxed">
                    {lang === 'vi'
                      ? 'Tiến hành đo Phóng điện cục bộ (PD) và kiểm tra nội soi khoang tank chính/bộ chuyển nấc để định vị điểm phóng điện trước khi đóng điện vận hành trở lại.'
                      : 'Conduct electrical/acoustic PD testing and borescope inspection of core/coils and tap changer contacts before re-energizing.'}
                  </p>
                </div>
              </>
            )}
          </div>

          {/* FSE Signoff strip */}
          <div className="pt-4 border-t border-slate-800 flex flex-col sm:flex-row sm:items-center justify-between text-xs text-slate-400 gap-2">
            <div>
              <span className="text-slate-500">{lang === 'vi' ? 'Kỹ sư FSE thẩm định:' : 'Certified FSE Engineer:'}</span>{' '}
              <strong className="text-amber-300">{data.site_info.fse_name}</strong>
            </div>
            <div>
              <span className="text-slate-500">{lang === 'vi' ? 'Đơn vị thực hiện:' : 'Inspection Org:'}</span>{' '}
              <strong className="text-white">TEV Technical Services Co., Ltd</strong>
            </div>
          </div>

          {/* Return to Form Button at Bottom of Report */}
          <div className="pt-3 border-t border-slate-800 flex items-center justify-between flex-wrap gap-3">
            <button
              type="button"
              onClick={() => {
                setMtsViewMode('form');
                window.scrollTo({ top: 0, behavior: 'smooth' });
              }}
              className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-amber-300 border border-amber-500/40 rounded-xl text-xs font-bold flex items-center gap-2 transition-all"
            >
              <Sliders size={14} />
              <span>{lang === 'vi' ? '← Quay Lại Biểu Mẫu Nhập Liệu (Chỉnh Sửa Thông Số)' : '← Back to Field Entry Form'}</span>
            </button>

            <button
              type="button"
              onClick={() => setShowPrintModal(true)}
              className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md"
            >
              <Printer size={14} />
              <span>{lang === 'vi' ? 'In / Xuất Báo Cáo PDF Chuẩn TEV' : 'Print / Export Official PDF'}</span>
            </button>
          </div>
        </div>
      </div>
      )}

      {/* JSON DATA MODAL */}
      {showJsonModal && (
        <div className="fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-4 backdrop-blur-xs">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden flex flex-col max-h-[85vh]">
            <div className="p-4 bg-slate-900 text-white flex items-center justify-between">
              <div className="flex items-center gap-2">
                <FileText size={18} className="text-amber-400" />
                <h3 className="font-bold text-sm">
                  {lang === 'vi' ? 'Dữ Liệu JSON Kiểm Định Hiện Trường (NETA MTS-2023)' : 'FSE Field Data JSON (NETA MTS-2023)'}
                </h3>
              </div>
              <button onClick={() => setShowJsonModal(false)} className="text-slate-400 hover:text-white">
                ✕
              </button>
            </div>
            <div className="p-4 flex-1 overflow-y-auto">
              <p className="text-xs text-slate-500 mb-2">
                {lang === 'vi'
                  ? 'Bạn có thể chỉnh sửa trực tiếp dữ liệu JSON đầu vào dưới đây để mô phỏng các kịch bản kiểm tra khác nhau.'
                  : 'You can directly edit the JSON input below to simulate different field testing scenarios.'}
              </p>
              <textarea
                value={jsonText}
                onChange={(e) => setJsonText(e.target.value)}
                className="w-full h-80 p-3 font-mono text-xs border rounded-xl bg-slate-900 text-emerald-400 focus:ring-2 focus:ring-amber-500 resize-none"
              />
            </div>
            <div className="p-4 bg-slate-50 border-t border-slate-200 flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setShowJsonModal(false)}
                className="px-4 py-2 border rounded-xl text-xs font-semibold text-slate-600 hover:bg-slate-100"
              >
                {lang === 'vi' ? 'Đóng' : 'Close'}
              </button>
              <button
                type="button"
                onClick={handleApplyJson}
                className="px-4 py-2 bg-amber-500 hover:bg-amber-400 text-slate-950 font-bold rounded-xl text-xs"
              >
                {lang === 'vi' ? 'Áp Dụng Dữ Liệu' : 'Apply JSON Data'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* NAMEPLATE CAMERA SCANNER MODAL */}
      <NameplateScannerModal
        isOpen={showCameraModal}
        onClose={() => setShowCameraModal(false)}
        onApplyData={handleApplyNameplateData}
      />

      {/* PROFESSIONAL TEV REPORT PRINT MODAL */}
      <TevReportPrintModal
        isOpen={showPrintModal}
        onClose={() => setShowPrintModal(false)}
        title={lang === 'vi' ? 'BIÊN BẢN BẢO TRÌ MÁY BIẾN ÁP NGÂM DẦU (NETA MTS-2023)' : 'LIQUID TRANSFORMER MAINTENANCE REPORT (NETA MTS-2023)'}
        standard="ANSI/NETA MTS-2023 Section 7.2.2 & IEEE C57.104"
        fseName={data.site_info.fse_name}
        siteInfo={{
          project_name: data.site_info.project_name,
          transformer_tag: data.site_info.transformer_tag,
          fse_name: data.site_info.fse_name,
          test_date: data.site_info.test_date,
        }}
        overallStatus={currentEvaluation.overall_status}
        statusReason={lang === 'vi' ? currentEvaluation.status_reason_vi : currentEvaluation.status_reason_en}
        visualChecks={[
          { label: 'Đối soát nhãn mác (IEEE C57.12.00)', value: data.visual_inspection.nameplate_match ? 'Khớp 100%' : 'Không khớp', pass: data.visual_inspection.nameplate_match },
          { label: 'Đệm khí Nitơ duy trì áp suất dương', value: data.visual_inspection.gas_blanket_pressure_positive ? 'Áp suất dương (+)' : 'Áp suất 0 / Âm', pass: data.visual_inspection.gas_blanket_pressure_positive },
          { label: 'Mức chất lỏng (Tank/LTC/Bushing)', value: data.visual_inspection.liquid_level_check, pass: true },
          { label: 'Van tản nhiệt, quạt & bơm làm mát', value: data.visual_inspection.valves_and_cooling_fans_check, pass: true },
          { label: 'Lực siết bu-lông (NETA Table 100.12)', value: data.visual_inspection.bolt_torque_check, pass: true },
          { label: 'Không rò rỉ dầu', value: data.visual_inspection.no_oil_leakage ? 'Khô ráo' : 'Rò rỉ', pass: data.visual_inspection.no_oil_leakage },
        ]}
        electricalTable={currentEvaluation.evaluations.items.map((i) => ({
          item: lang === 'vi' ? i.item : i.itemEn,
          measured: i.measured,
          standard: lang === 'vi' ? i.standard : i.standardEn,
          reference: i.reference,
          deviation: i.status === 'FAIL' ? 'VƯỢT NGƯỠNG' : 'ĐẠT',
          status: i.status,
        }))}
        warnings={[
          liveStats.arcingFail ? `CẢNH BÁO TỐI CẤP DGA: Xuất hiện khí Acetylene C2H2 = ${data.electrical_tests.dga_test_ieee_c57_104.acetylene_c2h2_ppm} ppm xác nhận có hồ quang năng lượng cao > 700°C bên trong máy biến áp!` : '',
          liveStats.breakdownFail ? `Dầu cách điện bị đánh thủng ở ${data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.dielectric_breakdown_d1816_1mm_kv} kV (thấp hơn giới hạn tối thiểu ${liveStats.breakdownThreshold} kV Table 100.4.1).` : '',
          liveStats.waterFail ? `Hàm lượng ẩm ${data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.water_content_d1533_ppm} ppm vượt quá ngưỡng cho phép ${liveStats.waterThreshold} ppm.` : '',
          liveStats.acidFail ? `Chỉ số Axit ${data.electrical_tests.service_aged_oil_sample_tests_table_100_4_1.acid_number_d974_mg_koh_g} mg KOH/g vượt quá 0.20 mg KOH/g báo hiệu dầu oxy hóa và tạo cặn bùn sludge.` : '',
        ].filter(Boolean)}
        recommendations={
          currentEvaluation.overall_status === 'PASS'
            ? [
                `1. Máy biến áp ${data.site_info.transformer_tag} đạt 100% các tiêu chí kiểm định NETA MTS-2023. Đủ điều kiện tiếp tục vận hành mang tải an toàn.`,
                '2. Duy trì kế hoạch bảo dưỡng định kỳ 12 tháng/lần cho thử nghiệm dầu và phân tích DGA tiếp theo.',
                '3. Kiểm tra màu sắc hạt hút ẩm bình thở định kỳ hàng quý và đảm bảo đệm khí Nitơ duy trì áp suất dương (0.5 - 5.0 psi).',
                '4. Giám sát đồng hồ đo nhiệt độ cuộn dây và dầu trên rơ le bảo vệ, ghi nhận dữ liệu định kỳ vào nhật ký vận hành trạm.'
              ]
            : currentEvaluation.overall_status === 'INVESTIGATE'
            ? [
                '1. Kiểm tra vệ sinh bề mặt tiếp xúc và siết lại bu-lông đầu cốt bushing theo lực mô-men quy định tại Table 100.12.',
                '2. Lấy mẫu dầu kiểm tra lại độ ẩm và DGA sau 6 tháng để theo dõi tốc độ biến thiên.',
                '3. Thay mới hạt hút ẩm silicagel nếu có dấu hiệu đổi màu quá 2/3 để ngăn ngừa độ ẩm thâm nhập vào dầu.',
                '4. Thực hiện quét ảnh nhiệt hồng ngoại các điểm đấu nối đầu cực để phát hiện sớm các điểm quá nhiệt cục bộ.'
              ]
            : [
                `1. CẢNH BÁO TỐI CẤP: Khuyên đơn vị vận hành không tiếp tục cho máy biến áp ${data.site_info.transformer_tag} mang tải nặng; lập kế hoạch tách cô lập kiểm tra ngay.`,
                '2. Tiến hành lấy mẫu dầu DGA đối soát lại sau 48 giờ để đánh giá tốc độ sinh khí (gas generation rate) per IEEE C57.104.',
                "3. Thực hiện lọc dầu tái chế hoàn nguyên (Reclamation/Regeneration) bằng đất lọc Fuller's Earth kết hợp hút chân không khử nước và bọt khí.",
                '4. Tiến hành đo Phóng điện cục bộ (PD) và kiểm tra nội soi khoang tank chính/bộ chuyển nấc để định vị điểm phóng điện trước khi đóng điện vận hành trở lại.'
              ]
        }
      />
    </div>
  );
};
