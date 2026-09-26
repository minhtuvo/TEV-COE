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
  Zap,
  Code2,
  Download,
  Upload,
  ArrowRight,
  TrendingDown,
  Scale
} from 'lucide-react';
import { NetaCableLvInputPayload, NetaCableLvEvaluationResult } from '../../server/netaCableLvAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

const SAMPLE_LV_CABLE_DEFECT_DATA: NetaCableLvInputPayload = {
  site_info: {
    project_name: "Toa Nha TEV Commercial Center",
    cable_tag: "CBL-LV-MAIN-FEEDER-01",
    cable_rating: "0.6/1kV Cu/PVC/PVC 4x(2x240mm2) Parallel Feeder",
    cable_length_meters: 120,
    fse_name: "Nguyen Van U",
    test_date: "2026-09-26",
    voltage_rating: "0.6/1kV",
    ambient_temp_c: 30
  },
  visual_inspection: {
    cable_data_match: true,
    physical_condition_terminals: "Pass (Good condition, properly terminated)",
    bolt_torque_check: "Pass",
    thermographic_survey: "Pass (Table 100.18 - No thermal anomaly prior to outage)"
  },
  electrical_tests: {
    insulation_resistance_1000v_1min_megohms: {
      phase_a_to_ground: 450,
      phase_b_to_ground: 420,
      phase_c_to_ground: 18,
      neutral_to_ground: 380,
      phase_a_to_b: 510,
      phase_b_to_c: 22,
      phase_c_to_a: 19
    },
    continuity_test: "Pass (All conductors continuous)",
    parallel_conductors_dc_resistance_milli_ohms: {
      phase_a_run1: 14.2,
      phase_a_run2: 14.3,
      phase_b_run1: 14.1,
      phase_b_run2: 14.2,
      phase_c_run1: 14.3,
      phase_c_run2: 28.5
    },
    test_voltage_v_dc: 1000
  },
  previous_test_data: {
    last_test_date: "2025-09-25",
    last_phase_c_ir_megohms: 410,
    last_phase_a_ir_megohms: 480,
    last_phase_b_ir_megohms: 460,
    last_neutral_ir_megohms: 400
  }
};

const SAMPLE_LV_CABLE_PASS_DATA: NetaCableLvInputPayload = {
  site_info: {
    project_name: "Nha May TEV High-Tech Park",
    cable_tag: "CBL-LV-SUBSTATION-F02",
    cable_rating: "0.6/1kV Cu/XLPE/PVC 4x(2x185mm2) Feeder",
    cable_length_meters: 85,
    fse_name: "Nguyen Van U",
    test_date: "2026-09-26",
    voltage_rating: "0.6/1kV",
    ambient_temp_c: 28
  },
  visual_inspection: {
    cable_data_match: true,
    physical_condition_terminals: "Pass (Excellent condition, clean & dry)",
    bolt_torque_check: "Pass",
    thermographic_survey: "Pass (Section 9 compliant)"
  },
  electrical_tests: {
    insulation_resistance_1000v_1min_megohms: {
      phase_a_to_ground: 520,
      phase_b_to_ground: 490,
      phase_c_to_ground: 480,
      neutral_to_ground: 450,
      phase_a_to_b: 580,
      phase_b_to_c: 560,
      phase_c_to_a: 570
    },
    continuity_test: "Pass (All conductors continuous 100%)",
    parallel_conductors_dc_resistance_milli_ohms: {
      phase_a_run1: 11.2,
      phase_a_run2: 11.3,
      phase_b_run1: 11.1,
      phase_b_run2: 11.2,
      phase_c_run1: 11.2,
      phase_c_run2: 11.4
    },
    test_voltage_v_dc: 1000
  },
  previous_test_data: {
    last_test_date: "2025-09-20",
    last_phase_c_ir_megohms: 500,
    last_phase_a_ir_megohms: 540,
    last_phase_b_ir_megohms: 510,
    last_neutral_ir_megohms: 470
  }
};

export const NetaCableLvChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaCableLvInputPayload>(SAMPLE_LV_CABLE_DEFECT_DATA);
  const [activeTab, setActiveTab] = useState<'visual' | 'electrical' | 'ai-report' | 'ai-prompt'>('electrical');
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaCableLvEvaluationResult | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showPromptModal, setShowPromptModal] = useState(false);
  const [jsonText, setJsonText] = useState('');
  const [copiedPrompt, setCopiedPrompt] = useState(false);
  const [copiedJson, setCopiedJson] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live Calculations according to ANSI/NETA ATS-2025 Mục 7.3.2 & Table 100.1
  const liveStats = useMemo(() => {
    const is300V = (data.site_info.cable_rating || '').includes('300V');
    const minStandardMegohms = is300V ? 25 : 100;
    const ir = data.electrical_tests.insulation_resistance_1000v_1min_megohms;

    const irA = ir.phase_a_to_ground ?? 0;
    const irB = ir.phase_b_to_ground ?? 0;
    const irC = ir.phase_c_to_ground ?? 0;
    const irN = ir.neutral_to_ground ?? minStandardMegohms;
    const irAB = ir.phase_a_to_b ?? 0;
    const irBC = ir.phase_b_to_c ?? 0;
    const irCA = ir.phase_c_to_a ?? 0;

    const isPhaseAPass = irA >= minStandardMegohms;
    const isPhaseBPass = irB >= minStandardMegohms;
    const isPhaseCPass = irC >= minStandardMegohms;
    const isPhaseNPass = irN >= minStandardMegohms;
    const isPhaseABPass = irAB >= minStandardMegohms;
    const isPhaseBCPass = irBC >= minStandardMegohms;
    const isPhaseCAPass = irCA >= minStandardMegohms;

    const isAllIrPass = isPhaseAPass && isPhaseBPass && isPhaseCPass && isPhaseNPass && isPhaseABPass && isPhaseBCPass && isPhaseCAPass;

    // Last year Phase C comparison
    const lastIrC = data.previous_test_data?.last_phase_c_ir_megohms || 410;
    const irCDropPct = lastIrC > 0 ? (((lastIrC - irC) / lastIrC) * 100).toFixed(1) : '0';

    // Parallel conductors delta calculation
    const par = data.electrical_tests.parallel_conductors_dc_resistance_milli_ohms;
    let maxParallelDelta = 0;
    let phaseADelta = 0;
    let phaseBDelta = 0;
    let phaseCDelta = 0;

    if (par) {
      const calcPct = (r1: number, r2: number) => {
        const min = Math.min(r1, r2);
        if (min === 0) return 0;
        return parseFloat((((Math.max(r1, r2) - min) / min) * 100).toFixed(1));
      };
      phaseADelta = calcPct(par.phase_a_run1, par.phase_a_run2);
      phaseBDelta = calcPct(par.phase_b_run1, par.phase_b_run2);
      phaseCDelta = calcPct(par.phase_c_run1, par.phase_c_run2);
      maxParallelDelta = Math.max(phaseADelta, phaseBDelta, phaseCDelta);
    }

    const isParallelPass = maxParallelDelta <= 15;
    const isParallelInvestigate = maxParallelDelta > 15;

    // Continuity
    const isContinuityPass = data.electrical_tests.continuity_test.toLowerCase().includes('pass');

    // Visual Mechanical
    const isVisualPass =
      data.visual_inspection.cable_data_match &&
      data.visual_inspection.physical_condition_terminals.toLowerCase().includes('pass') &&
      data.visual_inspection.bolt_torque_check.toLowerCase().includes('pass');

    // Overall Live Verdict
    let verdict: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
    if (!isAllIrPass || !isContinuityPass || !isVisualPass) {
      verdict = 'FAIL';
    } else if (isParallelInvestigate) {
      verdict = 'INVESTIGATE';
    }

    return {
      minStandardMegohms,
      irA,
      irB,
      irC,
      irN,
      irAB,
      irBC,
      irCA,
      isPhaseCPass,
      isAllIrPass,
      lastIrC,
      irCDropPct,
      phaseADelta,
      phaseBDelta,
      phaseCDelta,
      maxParallelDelta,
      isParallelPass,
      isParallelInvestigate,
      isContinuityPass,
      isVisualPass,
      verdict
    };
  }, [data]);

  const SYSTEM_INSTRUCTION_PROMPT = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách Cáp điện Hạ áp, Cấp điện áp tối đa 1,000V (Cables, Low-Voltage, 1,000-Volt Maximum). Nhiệm vụ của bạn là nhận dữ liệu kiểm tra ngoài site từ Field Service Engineer (FSE) và tự động đối soát với tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.3.2).

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.3.2):

1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Cable Data Match: Đối soát nhãn mác, chủng loại, tiết diện, số lõi cáp hạ áp khớp 100% bản vẽ thiết kế và sơ đồ đơn tuyến.
- Physical Condition & Terminations: Kiểm tra các đoạn cáp lộ thiên và đầu nối; không trầy xước vỏ bọc cách điện, đấu nối đúng sơ đồ.
- Bolt Torque: Mối nối bu-lông điện siết đạt lực theo dữ liệu nhà sản xuất hoặc Bảng Table 100.12.
- Thermographic Survey: Kết quả khảo sát nhiệt tuân thủ Section 9 / Table 100.18.

2. PHÉP ĐO ĐIỆN VÀ TIÊU CHUẨN ĐÁNH GIÁ (ELECTRICAL TEST VALUES):
- Điện trở cách điện (Insulation Resistance - IR):
  + Đo 1 phút giữa từng lõi dẫn với đất và giữa các lõi dẫn với nhau.
  + Điện áp thử: 500V DC cho cáp định mức 300V; 1000V DC cho cáp định mức 600V hoặc 1000V.
  + Mức IR tối thiểu quy đổi về 20°C đạt theo Bảng Table 100.1: Cáp 300V >= 25 Megohms; Cáp 600V/1000V BẮT BUỘC >= 100 Megohms.
- Thử nghiệm Thông mạch (Continuity Test): Cáp BẮT BUỘC phải đảm bảo thông mạch liên tục 100%.
- Điện trở Dây đấu Song song (Uniform Resistance of Parallel Conductors): CẢNH BÁO "INVESTIGATE" nếu có sự sai lệch điện trở giữa các sợi dây cáp đấu song song trên cùng một pha hoặc dây trung tính.

NHIỆM VỤ CỦA AI:
Khi nhận dữ liệu JSON từ FSE, hãy phân tích và trả về phản hồi định dạng Markdown gồm 4 phần:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE.
2. BẢNG PHÂN TÍCH CHI TIẾT CÁP HẠ ÁP (Detailed Evaluation Table): So sánh thực tế với NETA ATS-2025 Section 7.3.2 (Table 100.1, Table 100.12, Table 100.18).
3. PHÂN TÍCH CÁCH ĐIỆN, DÂY SONG SONG & NGUY CƠ SỰ CỐ (Insulation, Parallel Conductor & Fault Risk).
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations).`;

  const handleRunAiAnalysis = async () => {
    try {
      setIsAnalyzing(true);
      const res = await fetch('/api/field-service/neta-cable-lv-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });
      const result = await res.json();
      if (result.success && result.data) {
        setAnalysisResult(result.data);
        setActiveTab('ai-report');
      } else {
        alert('Lỗi phân tích: ' + (result.error || 'Không nhận được kết quả.'));
      }
    } catch (err: any) {
      console.error('Error invoking AI analysis:', err);
      alert('Không thể kết nối đến máy chủ AI: ' + err.message);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleSaveToCmms = () => {
    if (!analysisResult) {
      alert('Vui lòng thực hiện "Phân Tích AI NETA ATS-2025" trước khi lưu báo cáo vào CMMS.');
      return;
    }
    if (onSaveReport) {
      onSaveReport({
        data,
        analysisResult,
        equipmentId: data.site_info.cable_tag,
        date: data.site_info.test_date,
        type: 'Low-Voltage Cables (NETA ATS-2025 Mục 7.3.2)'
      });
      setSavedSuccessMessage(`Đã lưu thành công biên bản kiểm định cáp ${data.site_info.cable_tag} vào kho báo cáo CMMS!`);
      setTimeout(() => setSavedSuccessMessage(null), 5000);
    }
  };

  const handleCopyPrompt = () => {
    navigator.clipboard.writeText(SYSTEM_INSTRUCTION_PROMPT);
    setCopiedPrompt(true);
    setTimeout(() => setCopiedPrompt(false), 2000);
  };

  const handleCopyJson = () => {
    navigator.clipboard.writeText(JSON.stringify(data, null, 2));
    setCopiedJson(true);
    setTimeout(() => setCopiedJson(false), 2000);
  };

  const handleExportJson = () => {
    const dataStr = 'data:text/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(data, null, 2));
    const downloadAnchor = document.createElement('a');
    downloadAnchor.setAttribute('href', dataStr);
    downloadAnchor.setAttribute('download', `${data.site_info.cable_tag}_NETA_ATS_2025.json`);
    document.body.appendChild(downloadAnchor);
    downloadAnchor.click();
    downloadAnchor.remove();
  };

  const handleImportJson = () => {
    try {
      const parsed = JSON.parse(jsonText);
      if (parsed.site_info && parsed.electrical_tests && parsed.visual_inspection) {
        setData(parsed);
        setShowJsonModal(false);
        setJsonText('');
        alert('Đã nhập dữ liệu kiểm tra thành công!');
      } else {
        alert('Định dạng JSON không hợp lệ theo chuẩn NETA ATS-2025 Cáp Hạ Áp.');
      }
    } catch (e: any) {
      alert('Lỗi cú pháp JSON: ' + e.message);
    }
  };

  return (
    <div className="space-y-6">
      {/* HEADER BANNER */}
      <div className="bg-gradient-to-r from-emerald-950 via-teal-900 to-slate-900 rounded-2xl p-6 text-white shadow-xl border border-emerald-700/40 relative overflow-hidden">
        <div className="absolute right-0 top-0 w-96 h-96 bg-emerald-500/10 rounded-full blur-3xl pointer-events-none" />
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6 relative z-10">
          <div>
            <div className="flex flex-wrap items-center gap-2 mb-2">
              <span className="px-3 py-1 bg-emerald-500/20 text-emerald-300 border border-emerald-400/30 rounded-full text-xs font-bold uppercase tracking-wider flex items-center gap-1.5">
                <Cable size={14} className="text-emerald-400" />
                ANSI/NETA ATS-2025 Mục 7.3.2
              </span>
              <span className="px-3 py-1 bg-teal-500/20 text-teal-300 border border-teal-400/30 rounded-full text-xs font-semibold">
                Table 100.1 (IR ≥ 100 MΩ)
              </span>
              <span className="px-3 py-1 bg-blue-500/20 text-blue-300 border border-blue-400/30 rounded-full text-xs font-semibold">
                Uniform Resistance (Dây song song)
              </span>
            </div>
            <h1 className="text-2xl lg:text-3xl font-extrabold text-white flex items-center gap-3">
              Cáp Điện Hạ Áp (Cables, Low-Voltage, 1,000V Max)
            </h1>
            <p className="text-emerald-200/90 text-sm mt-1 max-w-3xl">
              Hệ thống kiểm định hiện trường tự động: Đối soát nhãn cáp, độ bền cách điện 1000V DC (Table 100.1), thông mạch liên tục 100% và độ đều điện trở một chiều giữa các sợi dây đấu song song.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => {
                setData(SAMPLE_LV_CABLE_DEFECT_DATA);
                setAnalysisResult(null);
              }}
              className="px-3.5 py-2 bg-rose-600/30 hover:bg-rose-600/40 text-rose-200 border border-rose-500/40 rounded-xl text-xs font-semibold transition-all flex items-center gap-1.5"
              title="Tải dữ liệu mẫu có lỗi suy giảm cách điện pha C và mất cân bằng sợi song song"
            >
              <AlertTriangle size={14} className="text-rose-400" />
              Mẫu Lỗi Sự Cố (Fail)
            </button>

            <button
              onClick={() => {
                setData(SAMPLE_LV_CABLE_PASS_DATA);
                setAnalysisResult(null);
              }}
              className="px-3.5 py-2 bg-emerald-600/30 hover:bg-emerald-600/40 text-emerald-200 border border-emerald-500/40 rounded-xl text-xs font-semibold transition-all flex items-center gap-1.5"
              title="Tải tuyến cáp đạt chuẩn nghiệm thu"
            >
              <CheckCircle2 size={14} className="text-emerald-400" />
              Mẫu Chuẩn (Pass)
            </button>

            <button
              onClick={() => setShowPromptModal(true)}
              className="px-3.5 py-2 bg-purple-600/40 hover:bg-purple-600/50 text-purple-200 border border-purple-400/40 rounded-xl text-xs font-semibold transition-all flex items-center gap-1.5"
            >
              <Code2 size={14} className="text-purple-300" />
              Prompt AI Studio
            </button>

            <button
              onClick={handleRunAiAnalysis}
              disabled={isAnalyzing}
              className="px-5 py-2 bg-gradient-to-r from-emerald-500 to-teal-500 hover:from-emerald-400 hover:to-teal-400 text-slate-950 font-bold rounded-xl text-sm transition-all shadow-lg shadow-emerald-500/30 flex items-center gap-2 disabled:opacity-50"
            >
              {isAnalyzing ? (
                <>
                  <div className="w-4 h-4 border-2 border-slate-950 border-t-transparent rounded-full animate-spin" />
                  Đang phân tích...
                </>
              ) : (
                <>
                  <Sparkles size={16} />
                  Phân Tích AI NETA ATS-2025
                </>
              )}
            </button>
          </div>
        </div>

        {/* LIVE METRICS TILES */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mt-6 pt-6 border-t border-emerald-800/40 text-xs">
          <div className="bg-slate-900/60 p-3 rounded-xl border border-emerald-500/20">
            <span className="text-slate-400 block mb-1">Cách điện Pha C - Đất:</span>
            <div className="flex items-center gap-2">
              <span className={`text-base font-bold ${liveStats.isPhaseCPass ? 'text-emerald-400' : 'text-rose-400'}`}>
                {liveStats.irC} MΩ
              </span>
              <span className={`text-[11px] px-1.5 py-0.5 rounded font-bold ${liveStats.isPhaseCPass ? 'bg-emerald-500/20 text-emerald-300' : 'bg-rose-500/20 text-rose-300'}`}>
                {liveStats.isPhaseCPass ? 'ĐẠT' : `FAIL (< ${liveStats.minStandardMegohms} MΩ)`}
              </span>
            </div>
            <span className="text-[10px] text-slate-400 mt-1 block">
              So với kỳ trước ({liveStats.lastIrC} MΩ): <span className="text-rose-400 font-semibold">-{liveStats.irCDropPct}%</span>
            </span>
          </div>

          <div className="bg-slate-900/60 p-3 rounded-xl border border-emerald-500/20">
            <span className="text-slate-400 block mb-1">Lệch dây song song Pha C:</span>
            <div className="flex items-center gap-2">
              <span className={`text-base font-bold ${liveStats.isParallelPass ? 'text-emerald-400' : 'text-amber-400'}`}>
                +{liveStats.phaseCDelta}%
              </span>
              <span className={`text-[11px] px-1.5 py-0.5 rounded font-bold ${liveStats.isParallelPass ? 'bg-emerald-500/20 text-emerald-300' : 'bg-amber-500/20 text-amber-300'}`}>
                {liveStats.isParallelPass ? 'ĐẠT' : 'INVESTIGATE'}
              </span>
            </div>
            <span className="text-[10px] text-slate-400 mt-1 block">
              Run 1 ({data.electrical_tests.parallel_conductors_dc_resistance_milli_ohms?.phase_c_run1 || 0} mΩ) vs Run 2 ({data.electrical_tests.parallel_conductors_dc_resistance_milli_ohms?.phase_c_run2 || 0} mΩ)
            </span>
          </div>

          <div className="bg-slate-900/60 p-3 rounded-xl border border-emerald-500/20">
            <span className="text-slate-400 block mb-1">Thử nghiệm Thông Mạch:</span>
            <div className="flex items-center gap-2">
              <span className="text-base font-bold text-emerald-400">100% Liên tục</span>
              <span className="text-[11px] px-1.5 py-0.5 rounded font-bold bg-emerald-500/20 text-emerald-300">
                PASS
              </span>
            </div>
            <span className="text-[10px] text-slate-400 mt-1 block">Mục 7.3.2.B.2</span>
          </div>

          <div className="bg-slate-900/60 p-3 rounded-xl border border-emerald-500/20">
            <span className="text-slate-400 block mb-1">Trạng thái tổng quan NETA:</span>
            <div className="flex items-center gap-2">
              <span className={`text-base font-extrabold ${liveStats.verdict === 'PASS' ? 'text-emerald-400' : liveStats.verdict === 'INVESTIGATE' ? 'text-amber-400' : 'text-rose-400'}`}>
                {liveStats.verdict}
              </span>
              <span className={`text-[11px] px-1.5 py-0.5 rounded font-bold ${liveStats.verdict === 'PASS' ? 'bg-emerald-500/20 text-emerald-300' : liveStats.verdict === 'INVESTIGATE' ? 'bg-amber-500/20 text-amber-300' : 'bg-rose-500/20 text-rose-300'}`}>
                {liveStats.verdict === 'FAIL' ? 'CẤM ĐÓNG ĐIỆN' : liveStats.verdict === 'INVESTIGATE' ? 'CẦN KIỂM TRA' : 'AN TOÀN'}
              </span>
            </div>
            <span className="text-[10px] text-slate-400 mt-1 block">Table 100.1 & Sec 7.3.2</span>
          </div>
        </div>
      </div>

      {savedSuccessMessage && (
        <div className="bg-emerald-50 border border-emerald-200 text-emerald-800 p-4 rounded-xl flex items-center justify-between shadow-sm animate-fade-in">
          <div className="flex items-center gap-2 font-medium">
            <CheckCircle2 size={18} className="text-emerald-600" />
            {savedSuccessMessage}
          </div>
          {onNavigateToReports && (
            <button
              onClick={onNavigateToReports}
              className="text-xs font-bold text-emerald-700 hover:text-emerald-800 underline flex items-center gap-1"
            >
              Mở Kho Báo Cáo <ArrowRight size={14} />
            </button>
          )}
        </div>
      )}

      {/* QUICK ACTIONS BAR */}
      <div className="flex flex-wrap items-center justify-between gap-3 bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
        <div className="flex items-center gap-1 bg-slate-100 p-1 rounded-lg">
          <button
            onClick={() => setActiveTab('electrical')}
            className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-all ${
              activeTab === 'electrical' ? 'bg-white text-emerald-700 shadow-sm' : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Phép Đo Điện NETA (Table 100.1)
          </button>
          <button
            onClick={() => setActiveTab('visual')}
            className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-all ${
              activeTab === 'visual' ? 'bg-white text-emerald-700 shadow-sm' : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Thị Giác & Cơ Khí (Sec 7.3.2.A)
          </button>
          <button
            onClick={() => setActiveTab('ai-report')}
            className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-all flex items-center gap-1 ${
              activeTab === 'ai-report' ? 'bg-white text-emerald-700 shadow-sm' : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <Sparkles size={13} className="text-emerald-600" />
            Kết Quả AI Chuyên Gia
            {analysisResult && (
              <span className={`w-2 h-2 rounded-full ${analysisResult.overall_status === 'PASS' ? 'bg-emerald-500' : 'bg-rose-500'}`} />
            )}
          </button>
          <button
            onClick={() => setActiveTab('ai-prompt')}
            className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-all flex items-center gap-1 ${
              activeTab === 'ai-prompt' ? 'bg-white text-purple-700 shadow-sm' : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <Code2 size={13} className="text-purple-600" />
            Prompt Google AI Studio
          </button>
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={handleCopyJson}
            className="px-3 py-1.5 border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-lg text-xs font-medium flex items-center gap-1.5 transition-colors"
            title="Sao chép toàn bộ JSON dữ liệu kiểm tra"
          >
            {copiedJson ? <Check size={14} className="text-emerald-600" /> : <Copy size={14} />}
            {copiedJson ? 'Đã sao chép JSON' : 'Copy JSON'}
          </button>

          <button
            onClick={handleExportJson}
            className="px-3 py-1.5 border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-lg text-xs font-medium flex items-center gap-1.5 transition-colors"
            title="Tải tệp JSON về máy tính"
          >
            <Download size={14} />
            Xuất JSON
          </button>

          <button
            onClick={() => setShowJsonModal(true)}
            className="px-3 py-1.5 border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-lg text-xs font-medium flex items-center gap-1.5 transition-colors"
            title="Dán hoặc nhập dữ liệu JSON từ site"
          >
            <Upload size={14} />
            Nhập JSON
          </button>

          <button
            onClick={() => setShowPrintModal(true)}
            className="px-3 py-1.5 border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-lg text-xs font-medium flex items-center gap-1.5 transition-colors"
          >
            <Printer size={14} />
            In / PDF
          </button>

          <button
            onClick={handleSaveToCmms}
            className="px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-sm transition-colors"
          >
            <Save size={14} />
            Lưu CMMS
          </button>
        </div>
      </div>

      {/* TAB 1: PHÉP ĐO ĐIỆN (ELECTRICAL TESTS) */}
      {activeTab === 'electrical' && (
        <div className="space-y-6">
          {/* SECTION 1: INSULATION RESISTANCE (TABLE 100.1) */}
          <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-100">
              <div>
                <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                  <Zap size={18} className="text-amber-500" />
                  1. Điện Trở Cách Điện (Insulation Resistance - IR) per Table 100.1
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  Điện áp thử: 1000V DC trong 1 phút (cho cáp 600V hoặc 1000V). Ngưỡng tối thiểu bắt buộc: <strong className="text-slate-800">≥ 100 Megohms (MΩ)</strong>.
                </p>
              </div>
              <div className="flex items-center gap-3 text-xs bg-slate-50 px-3 py-2 rounded-lg border border-slate-200">
                <span className="text-slate-500">Điện áp thử nghiệm:</span>
                <span className="font-bold text-slate-800">1000V DC (1 Phút)</span>
              </div>
            </div>

            {/* IR Phase-to-Ground & Neutral */}
            <div className="mt-5">
              <h4 className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-3">
                Đo Cách Điện Lõi Cáp với Đất (Conductor-to-Ground)
              </h4>
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
                  <label className="text-xs font-semibold text-slate-600 block mb-1">Phase A - Ground (MΩ)</label>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_1000v_1min_megohms.phase_a_to_ground}
                    onChange={(e) =>
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1000v_1min_megohms: {
                            ...data.electrical_tests.insulation_resistance_1000v_1min_megohms,
                            phase_a_to_ground: parseFloat(e.target.value) || 0
                          }
                        }
                      })
                    }
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md font-bold text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                  />
                  <div className="flex items-center justify-between text-[11px] mt-1.5">
                    <span className="text-slate-500">Chuẩn: ≥ 100 MΩ</span>
                    <span className="text-emerald-600 font-bold">PASS</span>
                  </div>
                </div>

                <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
                  <label className="text-xs font-semibold text-slate-600 block mb-1">Phase B - Ground (MΩ)</label>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_1000v_1min_megohms.phase_b_to_ground}
                    onChange={(e) =>
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1000v_1min_megohms: {
                            ...data.electrical_tests.insulation_resistance_1000v_1min_megohms,
                            phase_b_to_ground: parseFloat(e.target.value) || 0
                          }
                        }
                      })
                    }
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md font-bold text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                  />
                  <div className="flex items-center justify-between text-[11px] mt-1.5">
                    <span className="text-slate-500">Chuẩn: ≥ 100 MΩ</span>
                    <span className="text-emerald-600 font-bold">PASS</span>
                  </div>
                </div>

                <div className={`p-3 rounded-lg border ${liveStats.isPhaseCPass ? 'bg-slate-50 border-slate-200' : 'bg-rose-50/80 border-rose-300 ring-2 ring-rose-400/20'}`}>
                  <div className="flex items-center justify-between mb-1">
                    <label className="text-xs font-bold text-rose-700">Phase C - Ground (MΩ)</label>
                    <span className="text-[10px] px-1.5 py-0.2 bg-rose-200 text-rose-800 rounded font-bold">DEFECT</span>
                  </div>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_1000v_1min_megohms.phase_c_to_ground}
                    onChange={(e) =>
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1000v_1min_megohms: {
                            ...data.electrical_tests.insulation_resistance_1000v_1min_megohms,
                            phase_c_to_ground: parseFloat(e.target.value) || 0
                          }
                        }
                      })
                    }
                    className="w-full px-3 py-2 bg-white border border-rose-300 rounded-md font-extrabold text-rose-600 focus:outline-none focus:ring-2 focus:ring-rose-500"
                  />
                  <div className="flex items-center justify-between text-[11px] mt-1.5">
                    <span className="text-slate-500">Chuẩn: ≥ 100 MΩ</span>
                    <span className="text-rose-600 font-extrabold">FAIL ({liveStats.irC} &lt; 100 MΩ)</span>
                  </div>
                  <div className="mt-1 text-[10px] text-rose-600 font-medium">
                    📉 Năm ngoái: <strong>{liveStats.lastIrC} MΩ</strong> (Sụt -{liveStats.irCDropPct}%)
                  </div>
                </div>

                <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
                  <label className="text-xs font-semibold text-slate-600 block mb-1">Neutral - Ground (MΩ)</label>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_1000v_1min_megohms.neutral_to_ground ?? 380}
                    onChange={(e) =>
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1000v_1min_megohms: {
                            ...data.electrical_tests.insulation_resistance_1000v_1min_megohms,
                            neutral_to_ground: parseFloat(e.target.value) || 0
                          }
                        }
                      })
                    }
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md font-bold text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                  />
                  <div className="flex items-center justify-between text-[11px] mt-1.5">
                    <span className="text-slate-500">Chuẩn: ≥ 100 MΩ</span>
                    <span className="text-emerald-600 font-bold">PASS</span>
                  </div>
                </div>
              </div>
            </div>

            {/* IR Phase-to-Phase */}
            <div className="mt-6 pt-5 border-t border-slate-100">
              <h4 className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-3">
                Đo Cách Điện Giữa Các Pha với Nhau (Conductor-to-Conductor)
              </h4>
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
                  <label className="text-xs font-semibold text-slate-600 block mb-1">Phase A - Phase B (MΩ)</label>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_1000v_1min_megohms.phase_a_to_b}
                    onChange={(e) =>
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1000v_1min_megohms: {
                            ...data.electrical_tests.insulation_resistance_1000v_1min_megohms,
                            phase_a_to_b: parseFloat(e.target.value) || 0
                          }
                        }
                      })
                    }
                    className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md font-bold text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                  />
                  <div className="flex items-center justify-between text-[11px] mt-1.5">
                    <span className="text-slate-500">Chuẩn: ≥ 100 MΩ</span>
                    <span className="text-emerald-600 font-bold">PASS</span>
                  </div>
                </div>

                <div className="p-3 bg-rose-50/80 rounded-lg border border-rose-300">
                  <label className="text-xs font-bold text-rose-700 block mb-1">Phase B - Phase C (MΩ)</label>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_1000v_1min_megohms.phase_b_to_c}
                    onChange={(e) =>
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1000v_1min_megohms: {
                            ...data.electrical_tests.insulation_resistance_1000v_1min_megohms,
                            phase_b_to_c: parseFloat(e.target.value) || 0
                          }
                        }
                      })
                    }
                    className="w-full px-3 py-2 bg-white border border-rose-300 rounded-md font-extrabold text-rose-600 focus:outline-none focus:ring-2 focus:ring-rose-500"
                  />
                  <div className="flex items-center justify-between text-[11px] mt-1.5">
                    <span className="text-slate-500">Chuẩn: ≥ 100 MΩ</span>
                    <span className="text-rose-600 font-extrabold">FAIL ({liveStats.irBC} MΩ)</span>
                  </div>
                </div>

                <div className="p-3 bg-rose-50/80 rounded-lg border border-rose-300">
                  <label className="text-xs font-bold text-rose-700 block mb-1">Phase C - Phase A (MΩ)</label>
                  <input
                    type="number"
                    value={data.electrical_tests.insulation_resistance_1000v_1min_megohms.phase_c_to_a}
                    onChange={(e) =>
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1000v_1min_megohms: {
                            ...data.electrical_tests.insulation_resistance_1000v_1min_megohms,
                            phase_c_to_a: parseFloat(e.target.value) || 0
                          }
                        }
                      })
                    }
                    className="w-full px-3 py-2 bg-white border border-rose-300 rounded-md font-extrabold text-rose-600 focus:outline-none focus:ring-2 focus:ring-rose-500"
                  />
                  <div className="flex items-center justify-between text-[11px] mt-1.5">
                    <span className="text-slate-500">Chuẩn: ≥ 100 MΩ</span>
                    <span className="text-rose-600 font-extrabold">FAIL ({liveStats.irCA} MΩ)</span>
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* SECTION 2: PARALLEL CONDUCTORS DC RESISTANCE */}
          <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-100">
              <div>
                <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                  <Scale size={18} className="text-blue-600" />
                  2. Điện Trở Một Chiều Dây Đấu Song Song (Uniform Resistance of Parallel Conductors)
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  ANSI/NETA ATS-2025 Mục 7.3.2.B.3: Đo điện trở một chiều của từng sợi cáp song song. Báo động <strong className="text-amber-600 font-semibold">"INVESTIGATE"</strong> nếu có sự sai lệch điện trở giữa các sợi dây cáp đấu song song trên cùng một pha.
                </p>
              </div>
              <span className={`px-3 py-1 rounded-full text-xs font-bold border ${liveStats.isParallelPass ? 'bg-emerald-50 text-emerald-700 border-emerald-300' : 'bg-amber-50 text-amber-800 border-amber-300'}`}>
                {liveStats.isParallelPass ? 'Cân bằng tải đạt' : 'Cảnh báo lệch dòng tải'}
              </span>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mt-5">
              {/* Phase A Runs */}
              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
                <div className="flex items-center justify-between mb-3">
                  <span className="font-bold text-slate-800 text-sm">Pha A (2 Sợi Song Song)</span>
                  <span className="text-xs font-semibold px-2 py-0.5 rounded bg-emerald-100 text-emerald-800">
                    Lệch {liveStats.phaseADelta}%
                  </span>
                </div>
                <div className="space-y-3">
                  <div>
                    <label className="text-xs text-slate-500 block mb-1">Sợi 1 - Phase A Run 1 (mΩ)</label>
                    <input
                      type="number"
                      step="0.1"
                      value={data.electrical_tests.parallel_conductors_dc_resistance_milli_ohms?.phase_a_run1 || 14.2}
                      onChange={(e) =>
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            parallel_conductors_dc_resistance_milli_ohms: {
                              ...data.electrical_tests.parallel_conductors_dc_resistance_milli_ohms!,
                              phase_a_run1: parseFloat(e.target.value) || 0
                            }
                          }
                        })
                      }
                      className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md font-medium text-slate-800"
                    />
                  </div>
                  <div>
                    <label className="text-xs text-slate-500 block mb-1">Sợi 2 - Phase A Run 2 (mΩ)</label>
                    <input
                      type="number"
                      step="0.1"
                      value={data.electrical_tests.parallel_conductors_dc_resistance_milli_ohms?.phase_a_run2 || 14.3}
                      onChange={(e) =>
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            parallel_conductors_dc_resistance_milli_ohms: {
                              ...data.electrical_tests.parallel_conductors_dc_resistance_milli_ohms!,
                              phase_a_run2: parseFloat(e.target.value) || 0
                            }
                          }
                        })
                      }
                      className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md font-medium text-slate-800"
                    />
                  </div>
                </div>
              </div>

              {/* Phase B Runs */}
              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
                <div className="flex items-center justify-between mb-3">
                  <span className="font-bold text-slate-800 text-sm">Pha B (2 Sợi Song Song)</span>
                  <span className="text-xs font-semibold px-2 py-0.5 rounded bg-emerald-100 text-emerald-800">
                    Lệch {liveStats.phaseBDelta}%
                  </span>
                </div>
                <div className="space-y-3">
                  <div>
                    <label className="text-xs text-slate-500 block mb-1">Sợi 1 - Phase B Run 1 (mΩ)</label>
                    <input
                      type="number"
                      step="0.1"
                      value={data.electrical_tests.parallel_conductors_dc_resistance_milli_ohms?.phase_b_run1 || 14.1}
                      onChange={(e) =>
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            parallel_conductors_dc_resistance_milli_ohms: {
                              ...data.electrical_tests.parallel_conductors_dc_resistance_milli_ohms!,
                              phase_b_run1: parseFloat(e.target.value) || 0
                            }
                          }
                        })
                      }
                      className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md font-medium text-slate-800"
                    />
                  </div>
                  <div>
                    <label className="text-xs text-slate-500 block mb-1">Sợi 2 - Phase B Run 2 (mΩ)</label>
                    <input
                      type="number"
                      step="0.1"
                      value={data.electrical_tests.parallel_conductors_dc_resistance_milli_ohms?.phase_b_run2 || 14.2}
                      onChange={(e) =>
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            parallel_conductors_dc_resistance_milli_ohms: {
                              ...data.electrical_tests.parallel_conductors_dc_resistance_milli_ohms!,
                              phase_b_run2: parseFloat(e.target.value) || 0
                            }
                          }
                        })
                      }
                      className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md font-medium text-slate-800"
                    />
                  </div>
                </div>
              </div>

              {/* Phase C Runs - Defect Highlight */}
              <div className="p-4 bg-amber-50/70 rounded-xl border border-amber-300 ring-2 ring-amber-400/20">
                <div className="flex items-center justify-between mb-3">
                  <span className="font-bold text-amber-900 text-sm">Pha C (2 Sợi Song Song)</span>
                  <span className="text-xs font-extrabold px-2 py-0.5 rounded bg-amber-200 text-amber-900">
                    Lệch +{liveStats.phaseCDelta}% ⚠️
                  </span>
                </div>
                <div className="space-y-3">
                  <div>
                    <label className="text-xs text-slate-600 block mb-1 font-medium">Sợi 1 - Phase C Run 1 (mΩ)</label>
                    <input
                      type="number"
                      step="0.1"
                      value={data.electrical_tests.parallel_conductors_dc_resistance_milli_ohms?.phase_c_run1 || 14.3}
                      onChange={(e) =>
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            parallel_conductors_dc_resistance_milli_ohms: {
                              ...data.electrical_tests.parallel_conductors_dc_resistance_milli_ohms!,
                              phase_c_run1: parseFloat(e.target.value) || 0
                            }
                          }
                        })
                      }
                      className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md font-bold text-slate-900"
                    />
                  </div>
                  <div>
                    <label className="text-xs text-amber-800 block mb-1 font-bold">Sợi 2 - Phase C Run 2 (mΩ - Gấp 2 lần!)</label>
                    <input
                      type="number"
                      step="0.1"
                      value={data.electrical_tests.parallel_conductors_dc_resistance_milli_ohms?.phase_c_run2 || 28.5}
                      onChange={(e) =>
                        setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            parallel_conductors_dc_resistance_milli_ohms: {
                              ...data.electrical_tests.parallel_conductors_dc_resistance_milli_ohms!,
                              phase_c_run2: parseFloat(e.target.value) || 0
                            }
                          }
                        })
                      }
                      className="w-full px-3 py-2 bg-white border border-amber-400 rounded-md font-extrabold text-amber-700"
                    />
                  </div>
                </div>
                <p className="text-[11px] text-amber-800 mt-2 font-medium">
                  Cảnh báo: Sợi 2 tăng trở gấp đôi sẽ ép dòng điện dồn về Sợi 1 (~66.6% tải), gây nguy cơ cháy vỏ bọc do quá tải nhiệt cục bộ!
                </p>
              </div>
            </div>
          </div>

          {/* SECTION 3: CONTINUITY TEST */}
          <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div>
                <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                  <CheckCircle2 size={18} className="text-emerald-500" />
                  3. Thử Nghiệm Thông Mạch (Continuity Test)
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  ANSI/NETA ATS-2025 Mục 7.3.2.B.2: Cáp BẮT BUỘC phải đảm bảo thông mạch liên tục 100% giữa hai đầu tuyến.
                </p>
              </div>
              <div className="w-full sm:w-72">
                <input
                  type="text"
                  value={data.electrical_tests.continuity_test}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        continuity_test: e.target.value
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-md text-sm font-semibold text-slate-800 focus:bg-white focus:outline-none"
                />
              </div>
            </div>
          </div>
        </div>
      )}

      {/* TAB 2: THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL) */}
      {activeTab === 'visual' && (
        <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm space-y-6">
          <div className="pb-4 border-b border-slate-100">
            <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
              <Eye size={18} className="text-teal-600" />
              Kiểm Tra Thị Giác & Cơ Khí (Visual & Mechanical Inspection per Section 7.3.2.A)
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Đối soát thông số cáp, tình trạng vỏ bọc, đầu cáp và lực siết bu-lông đấu nối.
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <div className="flex items-center justify-between">
                <label className="text-sm font-semibold text-slate-800">
                  Cable Data Match (Khớp bản vẽ thiết kế)
                </label>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.cable_data_match}
                  onChange={(e) =>
                    setData({
                      ...data,
                      visual_inspection: {
                        ...data.visual_inspection,
                        cable_data_match: e.target.checked
                      }
                    })
                  }
                  className="w-5 h-5 text-emerald-600 rounded border-slate-300 focus:ring-emerald-500"
                />
              </div>
              <p className="text-xs text-slate-500">
                Đối soát nhãn mác, tiết diện ({data.site_info.cable_rating}), số lõi và chiều dài khớp 100% bản vẽ thiết kế và sơ đồ đơn tuyến.
              </p>
            </div>

            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <label className="text-sm font-semibold text-slate-800 block">
                Physical Condition & Terminations (Tình trạng vỏ & đầu nối)
              </label>
              <input
                type="text"
                value={data.visual_inspection.physical_condition_terminals}
                onChange={(e) =>
                  setData({
                    ...data,
                    visual_inspection: {
                      ...data.visual_inspection,
                      physical_condition_terminals: e.target.value
                    }
                  })
                }
                className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md text-sm font-medium text-slate-800"
              />
              <p className="text-xs text-slate-500">
                Kiểm tra các đoạn cáp lộ thiên và đầu nối; không trầy xước vỏ bọc cách điện, đấu nối đúng sơ đồ pha.
              </p>
            </div>

            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <label className="text-sm font-semibold text-slate-800 block">
                Bolt Torque Check (Lực siết bu-lông Table 100.12)
              </label>
              <input
                type="text"
                value={data.visual_inspection.bolt_torque_check}
                onChange={(e) =>
                  setData({
                    ...data,
                    visual_inspection: {
                      ...data.visual_inspection,
                      bolt_torque_check: e.target.value
                    }
                  })
                }
                className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md text-sm font-medium text-slate-800"
              />
              <p className="text-xs text-slate-500">
                Mối nối bu-lông điện siết đạt lực theo dữ liệu nhà sản xuất hoặc Bảng Table 100.12.
              </p>
            </div>

            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <label className="text-sm font-semibold text-slate-800 block">
                Thermographic Survey (Khảo sát nhiệt Table 100.18)
              </label>
              <input
                type="text"
                value={data.visual_inspection.thermographic_survey || ''}
                onChange={(e) =>
                  setData({
                    ...data,
                    visual_inspection: {
                      ...data.visual_inspection,
                      thermographic_survey: e.target.value
                    }
                  })
                }
                className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md text-sm font-medium text-slate-800"
              />
              <p className="text-xs text-slate-500">
                Kết quả khảo sát nhiệt hồng ngoại tuân thủ Section 9 và Table 100.18.
              </p>
            </div>
          </div>

          {/* SITE INFO META */}
          <div className="pt-6 border-t border-slate-100">
            <h4 className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-4">
              Thông Tin Tuyến Cáp & Hiện Trường
            </h4>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 text-xs">
              <div>
                <label className="font-semibold text-slate-600 block mb-1">Mã Tuyến Cáp (Cable Tag)</label>
                <input
                  type="text"
                  value={data.site_info.cable_tag}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, cable_tag: e.target.value }
                    })
                  }
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-md font-bold text-slate-900"
                />
              </div>

              <div>
                <label className="font-semibold text-slate-600 block mb-1">Chủng Loại / Tiết Diện (Rating)</label>
                <input
                  type="text"
                  value={data.site_info.cable_rating}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, cable_rating: e.target.value }
                    })
                  }
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-md font-medium text-slate-900"
                />
              </div>

              <div>
                <label className="font-semibold text-slate-600 block mb-1">Chiều Dài Tuyến (Mét)</label>
                <input
                  type="number"
                  value={data.site_info.cable_length_meters}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, cable_length_meters: parseFloat(e.target.value) || 0 }
                    })
                  }
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-md font-medium text-slate-900"
                />
              </div>

              <div>
                <label className="font-semibold text-slate-600 block mb-1">Kỹ Sư FSE & Ngày Test</label>
                <div className="flex gap-2">
                  <input
                    type="text"
                    value={data.site_info.fse_name}
                    onChange={(e) =>
                      setData({
                        ...data,
                        site_info: { ...data.site_info, fse_name: e.target.value }
                      })
                    }
                    className="w-1/2 px-2 py-2 bg-slate-50 border border-slate-300 rounded-md font-medium text-slate-900"
                  />
                  <input
                    type="date"
                    value={data.site_info.test_date}
                    onChange={(e) =>
                      setData({
                        ...data,
                        site_info: { ...data.site_info, test_date: e.target.value }
                      })
                    }
                    className="w-1/2 px-2 py-2 bg-slate-50 border border-slate-300 rounded-md font-medium text-slate-900"
                  />
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* TAB 3: BÁO CÁO PHÂN TÍCH CHUYÊN GIA AI */}
      {activeTab === 'ai-report' && (
        <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm space-y-6">
          {analysisResult ? (
            <div className="space-y-6">
              {/* Header Status Card */}
              <div
                className={`p-5 rounded-xl border flex flex-col md:flex-row md:items-center justify-between gap-4 ${
                  analysisResult.overall_status === 'PASS'
                    ? 'bg-emerald-50 border-emerald-200 text-emerald-950'
                    : analysisResult.overall_status === 'INVESTIGATE'
                    ? 'bg-amber-50 border-amber-200 text-amber-950'
                    : 'bg-rose-50 border-rose-200 text-rose-950'
                }`}
              >
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    {analysisResult.overall_status === 'PASS' ? (
                      <ShieldCheck size={24} className="text-emerald-600" />
                    ) : (
                      <ShieldAlert size={24} className="text-rose-600" />
                    )}
                    <h3 className="text-lg font-extrabold tracking-tight">
                      {analysisResult.status_title}
                    </h3>
                  </div>
                  <p className="text-xs opacity-90 max-w-2xl">{analysisResult.status_reason}</p>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    onClick={() => {
                      navigator.clipboard.writeText(analysisResult.markdown_report);
                      alert('Đã sao chép báo cáo Markdown vào bộ nhớ tạm!');
                    }}
                    className="px-3.5 py-2 bg-white border border-slate-300 hover:bg-slate-50 text-slate-800 rounded-lg text-xs font-semibold flex items-center gap-1.5 shadow-sm"
                  >
                    <Copy size={14} />
                    Copy Markdown
                  </button>

                  <button
                    onClick={handleSaveToCmms}
                    className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-sm"
                  >
                    <Save size={14} />
                    Lưu Báo Cáo CMMS
                  </button>
                </div>
              </div>

              {/* Markdown Content */}
              <div className="prose prose-slate max-w-none bg-slate-50/70 p-6 rounded-xl border border-slate-200 text-sm leading-relaxed whitespace-pre-wrap font-sans">
                {analysisResult.markdown_report}
              </div>
            </div>
          ) : (
            <div className="text-center py-12 space-y-4">
              <div className="w-16 h-16 bg-emerald-100 text-emerald-700 rounded-full flex items-center justify-center mx-auto">
                <Sparkles size={32} />
              </div>
              <h3 className="text-lg font-bold text-slate-800">Chưa có kết quả phân tích AI</h3>
              <p className="text-slate-500 text-xs max-w-md mx-auto">
                Nhấn nút &ldquo;Phân Tích AI NETA ATS-2025&rdquo; ở thanh tiêu đề để AI chuyên gia tự động đối soát toàn bộ giá trị đo đạc với tiêu chuẩn ANSI/NETA ATS-2025 Mục 7.3.2.
              </p>
              <button
                onClick={handleRunAiAnalysis}
                disabled={isAnalyzing}
                className="px-6 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl text-sm transition-all shadow-md inline-flex items-center gap-2"
              >
                <Sparkles size={16} />
                Bắt đầu phân tích ngay
              </button>
            </div>
          )}
        </div>
      )}

      {/* TAB 4: SYSTEM INSTRUCTION & PROMPT CHO GOOGLE AI STUDIO */}
      {activeTab === 'ai-prompt' && (
        <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm space-y-6">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-100">
            <div>
              <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                <Code2 size={18} className="text-purple-600" />
                Cấu Hình System Instruction & JSON Prompt cho Google AI Studio
              </h3>
              <p className="text-xs text-slate-500 mt-0.5">
                Dán System Instruction này vào ô System Instructions của Google AI Studio (Model: Gemini 2.5 Flash), sau đó đưa dữ liệu JSON mẫu vào User Prompt để thử nghiệm.
              </p>
            </div>
            <button
              onClick={handleCopyPrompt}
              className="px-4 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 transition-colors self-start sm:self-auto"
            >
              {copiedPrompt ? <Check size={14} /> : <Copy size={14} />}
              {copiedPrompt ? 'Đã sao chép prompt!' : 'Sao chép System Instruction'}
            </button>
          </div>

          <div className="space-y-4">
            <div>
              <span className="text-xs font-bold text-purple-900 uppercase tracking-wider block mb-2">
                1. System Instruction (Dán vào Google AI Studio):
              </span>
              <pre className="p-4 bg-slate-900 text-emerald-400 rounded-xl text-xs font-mono overflow-x-auto whitespace-pre-wrap leading-relaxed border border-purple-500/20 max-h-96">
                {SYSTEM_INSTRUCTION_PROMPT}
              </pre>
            </div>

            <div>
              <div className="flex items-center justify-between mb-2">
                <span className="text-xs font-bold text-slate-700 uppercase tracking-wider block">
                  2. User Prompt Mẫu (JSON Đầu Vào từ FSE):
                </span>
                <button
                  onClick={handleCopyJson}
                  className="text-xs text-purple-600 hover:text-purple-700 font-semibold flex items-center gap-1"
                >
                  <Copy size={12} /> Copy JSON
                </button>
              </div>
              <pre className="p-4 bg-slate-900 text-blue-300 rounded-xl text-xs font-mono overflow-x-auto whitespace-pre-wrap leading-relaxed border border-slate-700 max-h-80">
                {JSON.stringify(data, null, 2)}
              </pre>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: NHẬP JSON TỪ SITE */}
      {showJsonModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-2xl w-full p-6 shadow-2xl space-y-4 border border-slate-200">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <h3 className="font-bold text-base text-slate-900 flex items-center gap-2">
                <Upload size={18} className="text-emerald-600" />
                Nhập Dữ Liệu Kiểm Tra JSON (Cáp Hạ Áp NETA ATS-2025)
              </h3>
              <button
                onClick={() => setShowJsonModal(false)}
                className="text-slate-400 hover:text-slate-600"
              >
                ✕
              </button>
            </div>
            <p className="text-xs text-slate-500">
              Dán chuỗi JSON nhận được từ thiết bị đo hiện trường hoặc file lưu trữ:
            </p>
            <textarea
              rows={12}
              value={jsonText}
              onChange={(e) => setJsonText(e.target.value)}
              placeholder="Dán nội dung JSON vào đây..."
              className="w-full p-3 font-mono text-xs bg-slate-50 border border-slate-300 rounded-xl focus:bg-white focus:outline-none focus:ring-2 focus:ring-emerald-500"
            />
            <div className="flex items-center justify-end gap-3 pt-3 border-t border-slate-100">
              <button
                onClick={() => setShowJsonModal(false)}
                className="px-4 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-lg"
              >
                Hủy
              </button>
              <button
                onClick={handleImportJson}
                className="px-5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg shadow-sm"
              >
                Xác Nhận Nhập Dữ Liệu
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: PROMPT GOOGLE AI STUDIO CHI TIẾT */}
      {showPromptModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-3xl w-full p-6 shadow-2xl space-y-4 border border-slate-200 max-h-[90vh] flex flex-col">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div className="flex items-center gap-2">
                <Code2 size={20} className="text-purple-600" />
                <h3 className="font-bold text-base text-slate-900">
                  System Instruction Prompt cho Google AI Studio
                </h3>
              </div>
              <button
                onClick={() => setShowPromptModal(false)}
                className="text-slate-400 hover:text-slate-600"
              >
                ✕
              </button>
            </div>

            <div className="overflow-y-auto space-y-4 flex-1 pr-1 text-xs">
              <p className="text-slate-600">
                Sao chép System Instruction này và dán vào phần <strong>System Instructions</strong> trên Google AI Studio. Sau đó dùng JSON mẫu bên dưới làm User Prompt:
              </p>
              <div className="relative">
                <button
                  onClick={handleCopyPrompt}
                  className="absolute right-3 top-3 px-3 py-1.5 bg-purple-600 hover:bg-purple-700 text-white rounded-md text-xs font-bold flex items-center gap-1 shadow"
                >
                  {copiedPrompt ? <Check size={12} /> : <Copy size={12} />}
                  {copiedPrompt ? 'Đã chép' : 'Sao chép'}
                </button>
                <pre className="p-4 bg-slate-950 text-emerald-400 rounded-xl font-mono overflow-x-auto whitespace-pre-wrap leading-relaxed border border-slate-800">
                  {SYSTEM_INSTRUCTION_PROMPT}
                </pre>
              </div>
            </div>

            <div className="pt-3 border-t border-slate-100 flex justify-end">
              <button
                onClick={() => setShowPromptModal(false)}
                className="px-5 py-2 bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold rounded-lg"
              >
                Đóng
              </button>
            </div>
          </div>
        </div>
      )}

      {/* PRINT / EXPORT PDF MODAL */}
      {showPrintModal && (
        <TevReportPrintModal
          isOpen={showPrintModal}
          onClose={() => setShowPrintModal(false)}
          equipmentName={`${data.site_info.cable_tag} (${data.site_info.cable_rating})`}
          equipmentId={data.site_info.cable_tag}
          testDate={data.site_info.test_date}
          technician={data.site_info.fse_name}
          standard="ANSI/NETA ATS-2025 Mục 7.3.2"
          overallStatus={analysisResult?.overall_status || liveStats.verdict}
          reportTitle="BIÊN BẢN KIỂM ĐỊNH CÁP ĐIỆN HẠ ÁP (NETA ATS-2025)"
          evaluations={[
            {
              name: 'Kiểm Tra Thị Giác & Cơ Khí (Sec 7.3.2.A)',
              standard: 'NETA ATS-2025 Mục 7.3.2.A',
              actual: data.visual_inspection.physical_condition_terminals,
              status: liveStats.isVisualPass ? 'PASS' : 'FAIL',
              note: data.visual_inspection.bolt_torque_check
            },
            {
              name: 'Thử Nghiệm Thông Mạch (Continuity Test)',
              standard: 'Mục 7.3.2.B.2',
              actual: data.electrical_tests.continuity_test,
              status: liveStats.isContinuityPass ? 'PASS' : 'FAIL',
              note: 'Bắt buộc thông mạch liên tục 100%'
            },
            {
              name: 'Điện Trở Cách Điện Pha C - Đất (1000V DC)',
              standard: 'Table 100.1 (Min ≥ 100 MΩ)',
              actual: `${data.electrical_tests.insulation_resistance_1000v_1min_megohms.phase_c_to_ground} MΩ`,
              status: liveStats.isPhaseCPass ? 'PASS' : 'FAIL',
              note: `So với năm ngoái (${liveStats.lastIrC} MΩ): sụt -${liveStats.irCDropPct}%`
            },
            {
              name: 'Điện Trở Cách Điện Pha A, B & N',
              standard: 'Table 100.1 (Min ≥ 100 MΩ)',
              actual: `Pha A: ${data.electrical_tests.insulation_resistance_1000v_1min_megohms.phase_a_to_ground} MΩ, Pha B: ${data.electrical_tests.insulation_resistance_1000v_1min_megohms.phase_b_to_ground} MΩ`,
              status: 'PASS',
              note: 'Cách điện tốt'
            },
            {
              name: 'Điện Trở Dây Song Song Pha C (Sợi 1 vs Sợi 2)',
              standard: 'Mục 7.3.2.B.3 (Uniform Resistance)',
              actual: `Run 1: ${data.electrical_tests.parallel_conductors_dc_resistance_milli_ohms?.phase_c_run1} mΩ | Run 2: ${data.electrical_tests.parallel_conductors_dc_resistance_milli_ohms?.phase_c_run2} mΩ`,
              status: liveStats.isParallelPass ? 'PASS' : 'INVESTIGATE',
              note: `Lệch +${liveStats.phaseCDelta}% (Cao gấp 2 lần, nguy cơ quá tải nhiệt Sợi 1)`
            }
          ]}
          recommendations={
            analysisResult
              ? [analysisResult.status_reason]
              : [
                  'Tách cô lập lộ cáp CBL-LV-MAIN-FEEDER-01, cấm đóng điện.',
                  'Dò tìm điểm dập nứt vỏ cáp Pha C và sấy khô nước ngấm trong mương cáp.',
                  'Bấm lại đầu cos hoặc siết bu-lông Pha C Sợi 2 để cân bằng điện trở dây song song.'
                ]
          }
        />
      )}
    </div>
  );
};
