import React, { useState, useMemo } from 'react';
import {
  ShieldAlert,
  ShieldCheck,
  CheckCircle2,
  AlertTriangle,
  Sparkles,
  Printer,
  Save,
  Activity,
  Copy,
  Check,
  Eye,
  Info,
  Code2,
  Download,
  Upload,
  ArrowRight,
  TrendingUp,
  MapPin,
  Layers,
  Scale,
  Flame,
  Grid
} from 'lucide-react';
import { NetaGroundingInputPayload, NetaGroundingEvaluationResult } from '../../server/netaGroundingAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

const SAMPLE_GROUNDING_DEFECT_DATA: NetaGroundingInputPayload = {
  site_info: {
    project_name: "Tram Bien Ap TEV 110kV Factory Substation",
    grounding_system_tag: "GRID-EARTH-MAIN-01",
    system_type: "Copper Mesh Grounding Grid with Copper-Clad Steel Rods",
    fse_name: "Nguyen Van X",
    test_date: "2026-09-26",
    facility_type: "substation",
    substation_voltage_kv: 110,
    target_resistance_ohms: 1.0,
    soil_resistivity_ohm_meters: 145
  },
  visual_inspection: {
    grounding_layout_match: true,
    physical_condition_and_corrosion: "Pass (Good condition, no severed cables)",
    exothermic_welds_and_bolted_joints: "Pass (Welds solid, Cadweld joints uniform)",
    conductor_and_bus_sizing: "Pass (240mm2 Cu bare cable & 50x5mm copper bus)",
    ground_wells_and_test_links_accessible: true,
    bolt_torque_check: "Pass (Torqued per Table 100.12)"
  },
  electrical_tests: {
    fall_of_potential_ground_resistance_ieee_81_ohms: 7.8,
    point_to_point_continuity_milli_ohms: {
      main_bus_to_transformer_frame: 12.5,
      main_bus_to_switchgear_earth_bar: 14.2,
      main_bus_to_control_building_structure: 15.0,
      main_bus_to_substation_fence_gate: 680.0
    },
    fence_bonding_resistance_ohms: 0.68,
    test_method: "IEEE Std 81 Fall-of-Potential (62% Distance Rule with 3 Pockets)"
  },
  previous_test_data: {
    last_test_date: "2025-09-25",
    last_ground_resistance_ohms: 4.2,
    last_fence_gate_milli_ohms: 25.0
  }
};

const SAMPLE_GROUNDING_PASS_DATA: NetaGroundingInputPayload = {
  site_info: {
    project_name: "Nha May TEV Solar Power Plant 50MW",
    grounding_system_tag: "GRID-EARTH-SOLAR-MAIN",
    system_type: "Continuous Copper Ground Grid with Deep Drilled Rods & GEM",
    fse_name: "Nguyen Van X",
    test_date: "2026-09-26",
    facility_type: "substation",
    substation_voltage_kv: 110,
    target_resistance_ohms: 1.0,
    soil_resistivity_ohm_meters: 65
  },
  visual_inspection: {
    grounding_layout_match: true,
    physical_condition_and_corrosion: "Pass (Clean, dry and no corrosion)",
    exothermic_welds_and_bolted_joints: "Pass (Cadweld 100% solid)",
    conductor_and_bus_sizing: "Pass (240mm2 Cu bare conductor)",
    ground_wells_and_test_links_accessible: true,
    bolt_torque_check: "Pass"
  },
  electrical_tests: {
    fall_of_potential_ground_resistance_ieee_81_ohms: 0.65,
    point_to_point_continuity_milli_ohms: {
      main_bus_to_transformer_frame: 8.5,
      main_bus_to_switchgear_earth_bar: 9.2,
      main_bus_to_control_building_structure: 10.4,
      main_bus_to_substation_fence_gate: 18.2
    },
    fence_bonding_resistance_ohms: 0.018,
    test_method: "IEEE Std 81 Fall-of-Potential"
  },
  previous_test_data: {
    last_test_date: "2025-09-20",
    last_ground_resistance_ohms: 0.72,
    last_fence_gate_milli_ohms: 19.5
  }
};

export const NetaGroundingChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaGroundingInputPayload>(SAMPLE_GROUNDING_DEFECT_DATA);
  const [activeTab, setActiveTab] = useState<'electrical' | 'visual' | 'ai-report' | 'ai-prompt'>('electrical');
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaGroundingEvaluationResult | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showPromptModal, setShowPromptModal] = useState(false);
  const [jsonText, setJsonText] = useState('');
  const [copiedPrompt, setCopiedPrompt] = useState(false);
  const [copiedJson, setCopiedJson] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live Calculations according to ANSI/NETA ATS-2025 Section 7.13 & IEEE Std 81
  const liveStats = useMemo(() => {
    const measuredOhms = data.electrical_tests.fall_of_potential_ground_resistance_ieee_81_ohms || 0;
    const isSubstation =
      data.site_info.project_name.toLowerCase().includes('substation') ||
      data.site_info.project_name.toLowerCase().includes('trạm') ||
      data.site_info.facility_type === 'substation' ||
      (data.site_info.substation_voltage_kv && data.site_info.substation_voltage_kv >= 22);

    const targetMaxOhms = isSubstation ? 1.0 : 5.0;
    const isGridPass = measuredOhms <= 5.0 && (!isSubstation || measuredOhms <= 1.0);

    const prevOhms = data.previous_test_data?.last_ground_resistance_ohms || 4.2;
    const gridIncreasePct = prevOhms > 0 ? (((measuredOhms - prevOhms) / prevOhms) * 100).toFixed(1) : '0';

    // Fence Gate & Continuity
    const pts = data.electrical_tests.point_to_point_continuity_milli_ohms;
    const fenceGateMilliOhms = pts.main_bus_to_substation_fence_gate || 0;
    const fenceGateOhms = fenceGateMilliOhms / 1000;
    const isFenceGatePass = fenceGateOhms <= 0.5;

    // Equip continuity
    const trfMilliOhms = pts.main_bus_to_transformer_frame || 0;
    const swgMilliOhms = pts.main_bus_to_switchgear_earth_bar || 0;
    const bldgMilliOhms = pts.main_bus_to_control_building_structure || 0;
    const isEquipContinuityPass = trfMilliOhms <= 500 && swgMilliOhms <= 500 && bldgMilliOhms <= 500;

    // Visual
    const v = data.visual_inspection;
    const isVisualPass =
      v.grounding_layout_match &&
      v.ground_wells_and_test_links_accessible &&
      v.physical_condition_and_corrosion.toLowerCase().includes('pass') &&
      v.exothermic_welds_and_bolted_joints.toLowerCase().includes('pass') &&
      v.conductor_and_bus_sizing.toLowerCase().includes('pass') &&
      v.bolt_torque_check.toLowerCase().includes('pass');

    // Overall verdict
    let verdict: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
    if (!isGridPass || !isVisualPass) {
      verdict = 'FAIL';
    } else if (!isFenceGatePass) {
      verdict = 'INVESTIGATE';
    }

    return {
      measuredOhms,
      targetMaxOhms,
      isSubstation,
      isGridPass,
      prevOhms,
      gridIncreasePct,
      fenceGateMilliOhms,
      fenceGateOhms,
      isFenceGatePass,
      trfMilliOhms,
      swgMilliOhms,
      bldgMilliOhms,
      isEquipContinuityPass,
      isVisualPass,
      verdict
    };
  }, [data]);

  const SYSTEM_INSTRUCTION_PROMPT = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách Hệ thống Nối đất / Tiếp địa (Grounding Systems) theo tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.13) và IEEE Std 81.

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.13 & IEEE Std 81):

1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Layout & Design Match: Đối soát cấu trúc lưới tiếp địa, cọc đất, thanh cái tiếp địa chính (Main Ground Bus) khớp 100% bản vẽ thiết kế và quy chuẩn NFPA 70 / NEC.
- Physical Condition & Welds: Dây dẫn và thanh cái không bị nứt đứt, oxy hóa; mối hàn hóa nhiệt (exothermic welds) ngấu đều, không rỗng xốp; mối nối bu-lông chắc chắn.
- Conductor Sizing: Tiết diện dây cáp đồng và thanh cái đạt theo thiết kế và tính toán dòng ngắn mạch sự cố.
- Ground Wells & Test Links: Giếng kiểm tra sạch sẽ, không bị lấp đất; điểm tháo thắt cầu nối thử nghiệm (test links) tháo lắp dễ dàng.
- Bolt Torque: Mối nối bu-lông siết đạt lực theo dữ liệu nhà sản xuất hoặc Bảng Table 100.12.

2. PHÉP ĐO ĐIỆN VÀ TIÊU CHUẨN ĐÁNH GIÁ (ELECTRICAL TEST VALUES):
- Điện trở Nối đất Điện cực / Lưới đất (Ground System Resistance to Earth - IEEE Std 81 Fall-of-Potential):
  + Điện trở nối đất tổng thể xuống lòng đất BẮT BUỘC <= 5.0 Ohms đối với hệ thống công nghiệp/thương mại thông thường.
  + Đối với trạm biến áp trung/cao áp, trung tâm dữ liệu hoặc công trình điện lực chuyên dụng, điện trở BẮT BUỘC <= 1.0 Ohm (hoặc theo quy định cụ thể của dự án).
- Điện trở Liên kết Điểm-đến-Điểm (Point-to-Point Ground Continuity / Resistance):
  + Đo bằng thiết bị đo điện trở thấp (Low-Resistance Ohmmeter) từ Main Ground Bus đến vỏ tủ điện, khung máy biến áp, kết cấu thép.
  + Giá trị đo BẮT BUỘC < 0.5 Ohm. CẢNH BÁO "INVESTIGATE" nếu giá trị đo giữa các điểm tương tự lệch quá 50% so với giá trị nhỏ nhất.
- Tiếp địa Hàng rào & Khung vỏ (Fence & Structure Bonding): Hàng rào kim loại và kết cấu thép trạm phải liên kết tiếp địa tin cậy với điện trở liên kết <= 0.5 Ohm.

NHIỆM VỤ CỦA AI:
Khi nhận dữ liệu JSON từ FSE, hãy phân tích và trả về phản hồi định dạng Markdown gồm 4 phần:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE.
2. BẢNG PHÂN TÍCH CHI TIẾT HỆ THỐNG NỐI ĐẤT (Detailed Evaluation Table): So sánh thực tế với NETA ATS-2025 Section 7.13 (Table 100.12) và IEEE Std 81.
3. PHÂN TÍCH ĐIỆN TRỞ ĐẤT, LIÊN KẾT & NGUY CƠ ĐIỆN GIẬT (Ground Grid Resistance, Bonding Integrity & Touch/Step Potential Risk).
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations).`;

  const handleRunAiAnalysis = async () => {
    try {
      setIsAnalyzing(true);
      const res = await fetch('/api/field-service/neta-grounding-analyze', {
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
      alert('Không thể kết nối máy chủ AI: ' + err.message);
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
        equipmentId: data.site_info.grounding_system_tag,
        date: data.site_info.test_date,
        type: 'Grounding Systems (NETA ATS-2025 Mục 7.13 & IEEE Std 81)'
      });
      setSavedSuccessMessage(`Đã lưu thành công biên bản kiểm định tiếp địa ${data.site_info.grounding_system_tag} vào kho báo cáo CMMS!`);
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
    downloadAnchor.setAttribute('download', `${data.site_info.grounding_system_tag}_NETA_ATS_2025.json`);
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
        alert('Định dạng JSON không hợp lệ theo chuẩn NETA ATS-2025 Hệ thống tiếp địa.');
      }
    } catch (e: any) {
      alert('Lỗi cú pháp JSON: ' + e.message);
    }
  };

  return (
    <div className="space-y-6">
      {/* HEADER BANNER */}
      <div className="bg-gradient-to-r from-slate-950 via-teal-950 to-slate-900 rounded-2xl p-6 text-white shadow-xl border border-teal-800/40 relative overflow-hidden">
        <div className="absolute right-0 top-0 w-96 h-96 bg-teal-500/10 rounded-full blur-3xl pointer-events-none" />
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6 relative z-10">
          <div>
            <div className="flex flex-wrap items-center gap-2 mb-2">
              <span className="px-3 py-1 bg-teal-500/20 text-teal-300 border border-teal-400/30 rounded-full text-xs font-bold uppercase tracking-wider flex items-center gap-1.5">
                <Grid size={14} className="text-teal-400" />
                ANSI/NETA ATS-2025 Mục 7.13
              </span>
              <span className="px-3 py-1 bg-emerald-500/20 text-emerald-300 border border-emerald-400/30 rounded-full text-xs font-semibold">
                IEEE Std 81 (Fall-of-Potential 3-Point)
              </span>
              <span className="px-3 py-1 bg-blue-500/20 text-blue-300 border border-blue-400/30 rounded-full text-xs font-semibold">
                IEEE Std 80 (Touch & Step Voltage Safety)
              </span>
            </div>
            <h1 className="text-2xl lg:text-3xl font-extrabold text-white flex items-center gap-3">
              Hệ Thống Nối Đất & Tiếp Địa Trạm (Grounding Systems)
            </h1>
            <p className="text-teal-200/90 text-sm mt-1 max-w-3xl">
              Quy chuẩn đối soát chuyên sâu: Điện trở đất tổng thể Fall-of-Potential (≤ 1.0Ω trạm biến áp / ≤ 5.0Ω công nghiệp), liên kết đẳng thế điểm-đến-điểm (&lt; 0.5Ω) và an toàn điện áp chạm/bước hàng rào trạm.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => {
                setData(SAMPLE_GROUNDING_DEFECT_DATA);
                setAnalysisResult(null);
              }}
              className="px-3.5 py-2 bg-rose-600/30 hover:bg-rose-600/40 text-rose-200 border border-rose-500/40 rounded-xl text-xs font-semibold transition-all flex items-center gap-1.5"
              title="Tải mẫu tiếp địa sự cố: R_đất 7.8Ω & Cổng hàng rào 0.68Ω"
            >
              <AlertTriangle size={14} className="text-rose-400" />
              Mẫu Lỗi Sự Cố (Fail)
            </button>

            <button
              onClick={() => {
                setData(SAMPLE_GROUNDING_PASS_DATA);
                setAnalysisResult(null);
              }}
              className="px-3.5 py-2 bg-emerald-600/30 hover:bg-emerald-600/40 text-emerald-200 border border-emerald-500/40 rounded-xl text-xs font-semibold transition-all flex items-center gap-1.5"
              title="Tải mẫu tiếp địa đạt chuẩn nghiệm thu"
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
              className="px-5 py-2 bg-gradient-to-r from-teal-500 to-emerald-500 hover:from-teal-400 hover:to-emerald-400 text-slate-950 font-bold rounded-xl text-sm transition-all shadow-lg shadow-teal-500/30 flex items-center gap-2 disabled:opacity-50"
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
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mt-6 pt-6 border-t border-teal-800/40 text-xs">
          <div className="bg-slate-900/60 p-3 rounded-xl border border-teal-500/20">
            <span className="text-slate-400 block mb-1">Điện trở Lưới Đất (Fall-of-Potential):</span>
            <div className="flex items-center gap-2">
              <span className={`text-base font-bold ${liveStats.isGridPass ? 'text-emerald-400' : 'text-rose-400'}`}>
                {liveStats.measuredOhms} Ω
              </span>
              <span className={`text-[11px] px-1.5 py-0.5 rounded font-bold ${liveStats.isGridPass ? 'bg-emerald-500/20 text-emerald-300' : 'bg-rose-500/20 text-rose-300'}`}>
                {liveStats.isGridPass ? 'ĐẠT' : `FAIL (> ${liveStats.targetMaxOhms}Ω)`}
              </span>
            </div>
            <span className="text-[10px] text-slate-400 mt-1 block">
              Năm trước ({liveStats.prevOhms} Ω): <span className="text-rose-400 font-semibold">+{liveStats.gridIncreasePct}%</span>
            </span>
          </div>

          <div className="bg-slate-900/60 p-3 rounded-xl border border-teal-500/20">
            <span className="text-slate-400 block mb-1">Liên kết Cổng Hàng Rào Trạm:</span>
            <div className="flex items-center gap-2">
              <span className={`text-base font-bold ${liveStats.isFenceGatePass ? 'text-emerald-400' : 'text-amber-400'}`}>
                {liveStats.fenceGateOhms} Ω ({liveStats.fenceGateMilliOhms} mΩ)
              </span>
              <span className={`text-[11px] px-1.5 py-0.5 rounded font-bold ${liveStats.isFenceGatePass ? 'bg-emerald-500/20 text-emerald-300' : 'bg-amber-500/20 text-amber-300'}`}>
                {liveStats.isFenceGatePass ? 'ĐẠT' : 'INVESTIGATE'}
              </span>
            </div>
            <span className="text-[10px] text-slate-400 mt-1 block">
              Chuẩn NETA Sec 7.13.D.1: &lt; 0.50 Ω
            </span>
          </div>

          <div className="bg-slate-900/60 p-3 rounded-xl border border-teal-500/20">
            <span className="text-slate-400 block mb-1">Liên kết Vỏ MBA & Tủ Điện:</span>
            <div className="flex items-center gap-2">
              <span className="text-base font-bold text-emerald-400">
                {liveStats.trfMilliOhms} mΩ / {liveStats.swgMilliOhms} mΩ
              </span>
              <span className="text-[11px] px-1.5 py-0.5 rounded font-bold bg-emerald-500/20 text-emerald-300">
                PASS
              </span>
            </div>
            <span className="text-[10px] text-slate-400 mt-1 block">Rất tốt (&lt; 20 mΩ)</span>
          </div>

          <div className="bg-slate-900/60 p-3 rounded-xl border border-teal-500/20">
            <span className="text-slate-400 block mb-1">Trạng thái tổng quan NETA:</span>
            <div className="flex items-center gap-2">
              <span className={`text-base font-extrabold ${liveStats.verdict === 'PASS' ? 'text-emerald-400' : liveStats.verdict === 'INVESTIGATE' ? 'text-amber-400' : 'text-rose-400'}`}>
                {liveStats.verdict}
              </span>
              <span className={`text-[11px] px-1.5 py-0.5 rounded font-bold ${liveStats.verdict === 'PASS' ? 'bg-emerald-500/20 text-emerald-300' : liveStats.verdict === 'INVESTIGATE' ? 'bg-amber-500/20 text-amber-300' : 'bg-rose-500/20 text-rose-300'}`}>
                {liveStats.verdict === 'FAIL' ? 'CẤM ĐÓNG ĐIỆN' : liveStats.verdict === 'INVESTIGATE' ? 'CẦN KIỂM TRA' : 'AN TOÀN'}
              </span>
            </div>
            <span className="text-[10px] text-slate-400 mt-1 block">IEEE 81 & Sec 7.13</span>
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
              activeTab === 'electrical' ? 'bg-white text-teal-700 shadow-sm' : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Điện Trở Đất & Liên Kết (IEEE 81)
          </button>
          <button
            onClick={() => setActiveTab('visual')}
            className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-all ${
              activeTab === 'visual' ? 'bg-white text-teal-700 shadow-sm' : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Thị Giác, Mối Hàn & Giếng Đo (Sec 7.13.A)
          </button>
          <button
            onClick={() => setActiveTab('ai-report')}
            className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-all flex items-center gap-1 ${
              activeTab === 'ai-report' ? 'bg-white text-teal-700 shadow-sm' : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <Sparkles size={13} className="text-teal-600" />
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
          >
            {copiedJson ? <Check size={14} className="text-emerald-600" /> : <Copy size={14} />}
            {copiedJson ? 'Đã sao chép' : 'Copy JSON'}
          </button>

          <button
            onClick={handleExportJson}
            className="px-3 py-1.5 border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-lg text-xs font-medium flex items-center gap-1.5 transition-colors"
          >
            <Download size={14} />
            Xuất JSON
          </button>

          <button
            onClick={() => setShowJsonModal(true)}
            className="px-3 py-1.5 border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-lg text-xs font-medium flex items-center gap-1.5 transition-colors"
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
            className="px-3.5 py-1.5 bg-teal-600 hover:bg-teal-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-sm transition-colors"
          >
            <Save size={14} />
            Lưu CMMS
          </button>
        </div>
      </div>

      {/* TAB 1: PHÉP ĐO ĐIỆN VÀ ĐIỆN TRỞ ĐẤT */}
      {activeTab === 'electrical' && (
        <div className="space-y-6">
          {/* SECTION 1: FALL-OF-POTENTIAL GROUND RESISTANCE */}
          <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-100">
              <div>
                <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                  <Grid size={18} className="text-teal-600" />
                  1. Điện Trở Nối Đất Lưới Đất Tổng Thể (IEEE Std 81 Fall-of-Potential)
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  ANSI/NETA ATS-2025 Mục 7.13.D.2 & IEEE Std 81: Phương pháp đo 3 điểm (62% Rule). Ngưỡng tối đa: <strong className="text-slate-800">≤ 1.0 Ohm</strong> cho trạm biến áp / data center, và <strong className="text-slate-800">≤ 5.0 Ohms</strong> cho công nghiệp thông thường.
                </p>
              </div>
              <div className="flex items-center gap-3 text-xs bg-slate-50 px-3 py-2 rounded-lg border border-slate-200">
                <span className="text-slate-500">Phân loại trạm:</span>
                <span className="font-bold text-teal-800">Trạm 110kV Substation (Ngưỡng ≤ 1.0 Ω)</span>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mt-5">
              <div className={`p-4 rounded-xl border ${liveStats.isGridPass ? 'bg-slate-50 border-slate-200' : 'bg-rose-50/80 border-rose-300 ring-2 ring-rose-400/20'}`}>
                <div className="flex items-center justify-between mb-2">
                  <label className="text-xs font-bold text-rose-800">
                    Điện Trở Đất Đo Được (Ohms - Ω)
                  </label>
                  <span className={`text-[10px] px-1.5 py-0.5 rounded font-bold ${liveStats.isGridPass ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-200 text-rose-900'}`}>
                    {liveStats.isGridPass ? 'PASS' : 'FAIL (> 1.0 Ω)'}
                  </span>
                </div>
                <input
                  type="number"
                  step="0.01"
                  value={data.electrical_tests.fall_of_potential_ground_resistance_ieee_81_ohms}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        fall_of_potential_ground_resistance_ieee_81_ohms: parseFloat(e.target.value) || 0
                      }
                    })
                  }
                  className="w-full px-3 py-2.5 bg-white border border-rose-300 rounded-lg font-extrabold text-xl text-rose-600 focus:outline-none focus:ring-2 focus:ring-rose-500"
                />
                <div className="mt-2 text-xs text-rose-700 font-medium">
                  📈 Năm ngoái: <strong>{liveStats.prevOhms} Ω</strong> (Tăng +{liveStats.gridIncreasePct}%)
                </div>
              </div>

              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
                <label className="text-xs font-bold text-slate-700 block mb-2">
                  Phương Pháp Đo & Thiết Bị (Test Method)
                </label>
                <input
                  type="text"
                  value={data.electrical_tests.test_method || ''}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        test_method: e.target.value
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-xs font-semibold text-slate-800"
                />
                <span className="text-[11px] text-slate-500 mt-2 block">
                  Quy tắc 62% khoảng cách cọc dòng và cọc áp theo IEEE Std 81
                </span>
              </div>

              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
                <label className="text-xs font-bold text-slate-700 block mb-2">
                  Điện Trở Suất Của Đất (Soil Resistivity - Ω·m)
                </label>
                <input
                  type="number"
                  value={data.site_info.soil_resistivity_ohm_meters || 145}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: {
                        ...data.site_info,
                        soil_resistivity_ohm_meters: parseFloat(e.target.value) || 0
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-sm font-bold text-slate-800"
                />
                <span className="text-[11px] text-slate-500 mt-2 block">
                  Đo theo phương pháp 4 cực Wenner 4-Pin
                </span>
              </div>
            </div>
          </div>

          {/* SECTION 2: POINT-TO-POINT GROUND CONTINUITY */}
          <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-100">
              <div>
                <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                  <Activity size={18} className="text-teal-600" />
                  2. Điện Trở Liên Kết Tiếp Địa Điểm-đến-Điểm (Point-to-Point Ground Continuity)
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  ANSI/NETA ATS-2025 Mục 7.13.D.1: Đo bằng thiết bị đo điện trở thấp (DLRO) từ Main Ground Bus đến vỏ thiết bị. BẮT BUỘC <strong className="text-slate-800">&lt; 0.5 Ohm (500 mΩ)</strong>.
                </p>
              </div>
              <span className={`px-3 py-1 rounded-full text-xs font-bold border ${liveStats.isFenceGatePass ? 'bg-emerald-50 text-emerald-700 border-emerald-300' : 'bg-amber-50 text-amber-800 border-amber-300'}`}>
                {liveStats.isFenceGatePass ? 'Liên kết đẳng thế đạt' : 'Cảnh báo hở mạch hàng rào'}
              </span>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mt-5">
              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
                <div className="flex items-center justify-between mb-1">
                  <label className="text-xs font-bold text-slate-700">Main Bus $\rightarrow$ Vỏ MBA (mΩ)</label>
                  <span className="text-[10px] bg-emerald-100 text-emerald-800 px-1.5 py-0.5 rounded font-bold">PASS</span>
                </div>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.point_to_point_continuity_milli_ohms.main_bus_to_transformer_frame}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        point_to_point_continuity_milli_ohms: {
                          ...data.electrical_tests.point_to_point_continuity_milli_ohms,
                          main_bus_to_transformer_frame: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md font-bold text-slate-900"
                />
                <span className="text-[10px] text-slate-500 mt-1 block">Chuẩn: &lt; 500 mΩ (0.5 Ω)</span>
              </div>

              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
                <div className="flex items-center justify-between mb-1">
                  <label className="text-xs font-bold text-slate-700">Main Bus $\rightarrow$ Tủ MSB (mΩ)</label>
                  <span className="text-[10px] bg-emerald-100 text-emerald-800 px-1.5 py-0.5 rounded font-bold">PASS</span>
                </div>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.point_to_point_continuity_milli_ohms.main_bus_to_switchgear_earth_bar}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        point_to_point_continuity_milli_ohms: {
                          ...data.electrical_tests.point_to_point_continuity_milli_ohms,
                          main_bus_to_switchgear_earth_bar: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md font-bold text-slate-900"
                />
                <span className="text-[10px] text-slate-500 mt-1 block">Chuẩn: &lt; 500 mΩ (0.5 Ω)</span>
              </div>

              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
                <div className="flex items-center justify-between mb-1">
                  <label className="text-xs font-bold text-slate-700">Main Bus $\rightarrow$ Nhà ĐK (mΩ)</label>
                  <span className="text-[10px] bg-emerald-100 text-emerald-800 px-1.5 py-0.5 rounded font-bold">PASS</span>
                </div>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.point_to_point_continuity_milli_ohms.main_bus_to_control_building_structure}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        point_to_point_continuity_milli_ohms: {
                          ...data.electrical_tests.point_to_point_continuity_milli_ohms,
                          main_bus_to_control_building_structure: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md font-bold text-slate-900"
                />
                <span className="text-[10px] text-slate-500 mt-1 block">Chuẩn: &lt; 500 mΩ (0.5 Ω)</span>
              </div>

              <div className="p-4 bg-amber-50/80 rounded-xl border border-amber-300 ring-2 ring-amber-400/20">
                <div className="flex items-center justify-between mb-1">
                  <label className="text-xs font-extrabold text-amber-900">Cổng Hàng Rào Trạm (mΩ)</label>
                  <span className="text-[10px] bg-amber-200 text-amber-900 px-1.5 py-0.5 rounded font-bold">
                    INVESTIGATE ⚠️
                  </span>
                </div>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.point_to_point_continuity_milli_ohms.main_bus_to_substation_fence_gate}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        point_to_point_continuity_milli_ohms: {
                          ...data.electrical_tests.point_to_point_continuity_milli_ohms,
                          main_bus_to_substation_fence_gate: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-white border border-amber-400 rounded-md font-extrabold text-amber-700"
                />
                <span className="text-[10px] text-amber-800 font-medium mt-1 block">
                  Đo được {liveStats.fenceGateOhms} Ω (&gt; 0.50 Ω chuẩn tối đa)
                </span>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* TAB 2: THỊ GIÁC, MỐI HÀN VÀ GIẾNG KIỂM TRA */}
      {activeTab === 'visual' && (
        <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm space-y-6">
          <div className="pb-4 border-b border-slate-100">
            <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
              <Eye size={18} className="text-teal-600" />
              Kiểm Tra Thị Giác, Mối Hàn Hóa Nhiệt & Giếng Đất (Sec 7.13.A)
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Khớp bản vẽ thiết kế, kiểm tra mối hàn Cadweld, tiết diện cáp đồng và giếng kiểm tra tiếp địa.
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <div className="flex items-center justify-between">
                <label className="text-sm font-semibold text-slate-800">
                  Grounding Layout Match (Khớp bản vẽ thiết kế)
                </label>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.grounding_layout_match}
                  onChange={(e) =>
                    setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, grounding_layout_match: e.target.checked }
                    })
                  }
                  className="w-5 h-5 text-teal-600 rounded border-slate-300 focus:ring-teal-500"
                />
              </div>
              <p className="text-xs text-slate-500">
                Đối soát cấu trúc lưới tiếp địa, số lượng cọc đất, thanh cái chính Main Ground Bus khớp 100% bản vẽ thiết kế.
              </p>
            </div>

            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <div className="flex items-center justify-between">
                <label className="text-sm font-semibold text-slate-800">
                  Giếng Kiểm Tra & Test Links Tiếp Cận Tốt
                </label>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.ground_wells_and_test_links_accessible}
                  onChange={(e) =>
                    setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, ground_wells_and_test_links_accessible: e.target.checked }
                    })
                  }
                  className="w-5 h-5 text-teal-600 rounded border-slate-300 focus:ring-teal-500"
                />
              </div>
              <p className="text-xs text-slate-500">
                Giếng kiểm tra sạch sẽ, không bị lấp đất đá, cầu nối thử nghiệm (test links) tháo lắp đo đạc dễ dàng.
              </p>
            </div>

            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <label className="text-sm font-semibold text-slate-800 block">
                Mối Hàn Hóa Nhiệt Cadweld & Khớp Nối Bu-lông
              </label>
              <input
                type="text"
                value={data.visual_inspection.exothermic_welds_and_bolted_joints}
                onChange={(e) =>
                  setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, exothermic_welds_and_bolted_joints: e.target.value }
                  })
                }
                className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md text-sm font-medium text-slate-800"
              />
              <p className="text-xs text-slate-500">
                Mối hàn hóa nhiệt ngấu đều, không rỗ xốp hay nứt nẻ; mối nối bu-lông siết chắc chắn.
              </p>
            </div>

            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <label className="text-sm font-semibold text-slate-800 block">
                Tiết Diện Cáp Đồng & Thanh Cái (Conductor Sizing)
              </label>
              <input
                type="text"
                value={data.visual_inspection.conductor_and_bus_sizing}
                onChange={(e) =>
                  setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, conductor_and_bus_sizing: e.target.value }
                  })
                }
                className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md text-sm font-medium text-slate-800"
              />
              <p className="text-xs text-slate-500">
                Tiết diện dây cáp đồng trần và thanh cái đạt theo thiết kế và tính toán dòng ngắn mạch sự cố (NFPA 70 / NEC).
              </p>
            </div>
          </div>

          {/* SITE INFO META */}
          <div className="pt-6 border-t border-slate-100">
            <h4 className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-4">
              Thông Tin Hệ Thống Nối Đất Hiện Trường
            </h4>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 text-xs">
              <div>
                <label className="font-semibold text-slate-600 block mb-1">Mã Lưới Tiếp Địa (Tag)</label>
                <input
                  type="text"
                  value={data.site_info.grounding_system_tag}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, grounding_system_tag: e.target.value }
                    })
                  }
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-md font-bold text-slate-900"
                />
              </div>

              <div>
                <label className="font-semibold text-slate-600 block mb-1">Cấu Trúc Tiếp Địa (System Type)</label>
                <input
                  type="text"
                  value={data.site_info.system_type}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, system_type: e.target.value }
                    })
                  }
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-md font-medium text-slate-900"
                />
              </div>

              <div>
                <label className="font-semibold text-slate-600 block mb-1">Kỹ Sư FSE Đo Đạc</label>
                <input
                  type="text"
                  value={data.site_info.fse_name}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, fse_name: e.target.value }
                    })
                  }
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-md font-medium text-slate-900"
                />
              </div>

              <div>
                <label className="font-semibold text-slate-600 block mb-1">Ngày Kiểm Định</label>
                <input
                  type="date"
                  value={data.site_info.test_date}
                  onChange={(e) =>
                    setData({
                      ...data,
                      site_info: { ...data.site_info, test_date: e.target.value }
                    })
                  }
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-md font-medium text-slate-900"
                />
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
                    className="px-4 py-2 bg-teal-600 hover:bg-teal-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-sm"
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
              <div className="w-16 h-16 bg-teal-100 text-teal-700 rounded-full flex items-center justify-center mx-auto">
                <Sparkles size={32} />
              </div>
              <h3 className="text-lg font-bold text-slate-800">Chưa có kết quả phân tích AI</h3>
              <p className="text-slate-500 text-xs max-w-md mx-auto">
                Nhấn nút &ldquo;Phân Tích AI NETA ATS-2025&rdquo; ở góc trên để AI chuyên gia tự động đối soát giá trị đo Fall-of-Potential và tiếp địa đẳng thế theo chuẩn IEEE Std 81 & Section 7.13.
              </p>
              <button
                onClick={handleRunAiAnalysis}
                disabled={isAnalyzing}
                className="px-6 py-2.5 bg-teal-600 hover:bg-teal-700 text-white font-bold rounded-xl text-sm transition-all shadow-md inline-flex items-center gap-2"
              >
                <Sparkles size={16} />
                Bắt đầu phân tích ngay
              </button>
            </div>
          )}
        </div>
      )}

      {/* TAB 4: SYSTEM INSTRUCTION PROMPT CHO GOOGLE AI STUDIO */}
      {activeTab === 'ai-prompt' && (
        <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm space-y-6">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-100">
            <div>
              <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                <Code2 size={18} className="text-purple-600" />
                Cấu Hình System Instruction & JSON Prompt cho Google AI Studio
              </h3>
              <p className="text-xs text-slate-500 mt-0.5">
                Dán System Instruction này vào ô System Instructions của Google AI Studio (Model: Gemini 2.5 Flash), sau đó đưa dữ liệu JSON mẫu vào User Prompt để kiểm thử.
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
              <pre className="p-4 bg-slate-900 text-teal-400 rounded-xl text-xs font-mono overflow-x-auto whitespace-pre-wrap leading-relaxed border border-purple-500/20 max-h-96">
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
                <Upload size={18} className="text-teal-600" />
                Nhập Dữ Liệu Kiểm Tra JSON (Hệ Thống Nối Đất NETA ATS-2025)
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
              className="w-full p-3 font-mono text-xs bg-slate-50 border border-slate-300 rounded-xl focus:bg-white focus:outline-none focus:ring-2 focus:ring-teal-500"
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
                className="px-5 py-2 bg-teal-600 hover:bg-teal-700 text-white text-xs font-bold rounded-lg shadow-sm"
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
                <pre className="p-4 bg-slate-950 text-teal-400 rounded-xl font-mono overflow-x-auto whitespace-pre-wrap leading-relaxed border border-slate-800">
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
          equipmentName={`${data.site_info.grounding_system_tag} (${data.site_info.system_type})`}
          equipmentId={data.site_info.grounding_system_tag}
          testDate={data.site_info.test_date}
          technician={data.site_info.fse_name}
          standard="ANSI/NETA ATS-2025 Mục 7.13 & IEEE Std 81"
          overallStatus={analysisResult?.overall_status || liveStats.verdict}
          reportTitle="BIÊN BẢN KIỂM ĐỊNH HỆ THỐNG NỐI ĐẤT (NETA ATS-2025)"
          evaluations={[
            {
              name: 'Kiểm Tra Thị Giác, Mối Hàn & Giếng Đất (Sec 7.13.A)',
              standard: 'NETA ATS-2025 Mục 7.13.A',
              actual: data.visual_inspection.exothermic_welds_and_bolted_joints,
              status: liveStats.isVisualPass ? 'PASS' : 'FAIL',
              note: data.visual_inspection.conductor_and_bus_sizing
            },
            {
              name: 'Điện Trở Nối Đất Lưới Đất (IEEE Std 81 Fall-of-Potential)',
              standard: liveStats.isSubstation ? 'Max ≤ 1.0 Ω (Trạm 110kV)' : 'Max ≤ 5.0 Ω',
              actual: `${data.electrical_tests.fall_of_potential_ground_resistance_ieee_81_ohms} Ω`,
              status: liveStats.isGridPass ? 'PASS' : 'FAIL',
              note: `So với năm ngoái (${liveStats.prevOhms} Ω): tăng +${liveStats.gridIncreasePct}%`
            },
            {
              name: 'Liên Kết Vỏ Máy Biến Áp (Point-to-Point)',
              standard: 'Mục 7.13.D.1 (Max < 0.50 Ω)',
              actual: `${data.electrical_tests.point_to_point_continuity_milli_ohms.main_bus_to_transformer_frame} mΩ`,
              status: 'PASS',
              note: 'Tiếp địa đẳng thế xuất sắc'
            },
            {
              name: 'Liên Kết Cổng Hàng Rào Trạm (Fence Gate)',
              standard: 'Mục 7.13.D.1 & IEEE Std 80 (Max ≤ 0.50 Ω)',
              actual: `${liveStats.fenceGateOhms} Ω (${liveStats.fenceGateMilliOhms} mΩ)`,
              status: liveStats.isFenceGatePass ? 'PASS' : 'INVESTIGATE',
              note: 'Vượt quá 0.50 Ω, nguy cơ điện áp chạm Touch Potential nguy hiểm!'
            }
          ]}
          recommendations={
            analysisResult
              ? [analysisResult.status_reason]
              : [
                  'CẢNH BÁO: Điện trở đất toàn trạm đạt 7.8 Ω vượt ngưỡng an toàn cho phép.',
                  'Thay thế dây liên kết mềm tiếp địa cổng hàng rào trạm để hạ điện trở liên kết về < 0.5 Ω.',
                  'Bổ sung cọc tiếp địa đóng sâu hoặc bơm hợp chất giảm điện trở đất GEM để hạ điện trở lưới đất về < 1.0 Ω.'
                ]
          }
        />
      )}
    </div>
  );
};
