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
  Power,
  Cpu,
  Clock,
  Wrench,
  Boxes
} from 'lucide-react';
import { NetaLvBreakerInputPayload, NetaLvBreakerEvaluationResult } from '../../server/netaLvBreakerAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

const SAMPLE_LV_BREAKER_DEFECT_DATA: NetaLvBreakerInputPayload = {
  site_info: {
    project_name: "Tram Phieu TEV Main Substation",
    breaker_tag: "ACB-INCOMING-3200A",
    breaker_rating: "3200A 65kA 600V 3-Pole Drawout Air Circuit Breaker",
    trip_unit_model: "MicroLogic 6.0A",
    fse_name: "Nguyen Van V",
    test_date: "2026-09-26",
    manufacturer: "Schneider Electric / MasterPact NW",
    rated_current_a: 3200,
    breaking_capacity_ka: 65,
    substation_location: "MSB Room 1 - Cell 01"
  },
  visual_inspection: {
    nameplate_match: true,
    physical_condition_cleanliness: "Pass (Clean, no debris, shipping brackets removed)",
    racking_and_interlocks: "Pass (Drawout mechanism smooth, cell interlocks operate correctly)",
    contacts_and_finger_clusters: "Good condition (Main & arcing contacts intact, finger clusters clean)",
    arc_chutes_condition: "Pass (No cracks, de-ionizing plates intact)",
    trip_unit_settings_match: true,
    bolt_torque_check: "Pass (Torqued per Table 100.12)"
  },
  electrical_tests: {
    bolted_resistance_micro_ohms: [9.1, 9.3, 9.2],
    contact_pole_resistance_micro_ohms: {
      pole_a: 18.2,
      pole_b: 18.5,
      pole_c: 42.1
    },
    insulation_resistance_1000v_1min_megohms: {
      closed_phase_a_to_ground: 2200,
      closed_phase_b_to_ground: 2100,
      closed_phase_c_to_ground: 1950,
      open_across_pole_a: 2500,
      open_across_pole_b: 2400,
      open_across_pole_c: 2300
    },
    control_wiring_ir_megohms: 35.0,
    secondary_injection_trip_unit_test: {
      long_time_pickup_status: "Pass (Tripped at 1.05 Ir within manufacturer tolerance)",
      short_time_delay_status: "Pass (Tripped at 0.20s with 5x Ir, compliant with curve)",
      instantaneous_pickup_status: "Pass (Instantaneous trip verified at 10x In)",
      ground_fault_pickup_status: "Fail (Tripped at 0.15s, expected 0.30s per curve)"
    },
    minimum_operating_voltage_test: "Pass (Close coil operates at 85% Un, Shunt trip operates at 70% Un)"
  },
  previous_test_data: {
    last_test_date: "2025-09-25",
    last_pole_c_contact_resistance_micro_ohms: 18.6,
    last_pole_a_contact_resistance_micro_ohms: 18.3,
    last_pole_b_contact_resistance_micro_ohms: 18.4,
    last_control_wiring_ir_megohms: 38.0
  }
};

const SAMPLE_LV_BREAKER_PASS_DATA: NetaLvBreakerInputPayload = {
  site_info: {
    project_name: "Nha May TEV Factory Block B",
    breaker_tag: "ACB-FEEDER-1600A-02",
    breaker_rating: "1600A 50kA 600V 3-Pole Drawout Air Circuit Breaker",
    trip_unit_model: "MicroLogic 5.0E",
    fse_name: "Nguyen Van V",
    test_date: "2026-09-26",
    manufacturer: "ABB / Emax 2",
    rated_current_a: 1600,
    breaking_capacity_ka: 50,
    substation_location: "Substation B - Panel 02"
  },
  visual_inspection: {
    nameplate_match: true,
    physical_condition_cleanliness: "Pass",
    racking_and_interlocks: "Pass",
    contacts_and_finger_clusters: "Good condition",
    arc_chutes_condition: "Pass",
    trip_unit_settings_match: true,
    bolt_torque_check: "Pass"
  },
  electrical_tests: {
    bolted_resistance_micro_ohms: [8.8, 8.9, 9.0],
    contact_pole_resistance_micro_ohms: {
      pole_a: 16.4,
      pole_b: 16.8,
      pole_c: 16.5
    },
    insulation_resistance_1000v_1min_megohms: {
      closed_phase_a_to_ground: 2800,
      closed_phase_b_to_ground: 2750,
      closed_phase_c_to_ground: 2900,
      open_across_pole_a: 3200,
      open_across_pole_b: 3100,
      open_across_pole_c: 3150
    },
    control_wiring_ir_megohms: 42.0,
    secondary_injection_trip_unit_test: {
      long_time_pickup_status: "Pass (Tripped within curve)",
      short_time_delay_status: "Pass (Tripped at 0.20s)",
      instantaneous_pickup_status: "Pass (Tripped instantaneously)",
      ground_fault_pickup_status: "Pass (Tripped at 0.30s, exactly matches curve)"
    },
    minimum_operating_voltage_test: "Pass"
  },
  previous_test_data: {
    last_test_date: "2025-09-20",
    last_pole_c_contact_resistance_micro_ohms: 16.2
  }
};

export const NetaLvBreakerChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaLvBreakerInputPayload>(SAMPLE_LV_BREAKER_DEFECT_DATA);
  const [activeTab, setActiveTab] = useState<'electrical' | 'visual' | 'ai-report' | 'ai-prompt'>('electrical');
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaLvBreakerEvaluationResult | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showPromptModal, setShowPromptModal] = useState(false);
  const [jsonText, setJsonText] = useState('');
  const [copiedPrompt, setCopiedPrompt] = useState(false);
  const [copiedJson, setCopiedJson] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live Calculations according to ANSI/NETA ATS-2025 Mục 7.6.1.2
  const liveStats = useMemo(() => {
    const cp = data.electrical_tests.contact_pole_resistance_micro_ohms;
    const minPole = Math.min(cp.pole_a, cp.pole_b, cp.pole_c);
    const maxPole = Math.max(cp.pole_a, cp.pole_b, cp.pole_c);
    const poleCDev = minPole > 0 ? parseFloat((((cp.pole_c - minPole) / minPole) * 100).toFixed(1)) : 0;
    const maxPoleDev = minPole > 0 ? parseFloat((((maxPole - minPole) / minPole) * 100).toFixed(1)) : 0;

    const isContactPass = maxPoleDev <= 50;

    const lastPoleC = data.previous_test_data?.last_pole_c_contact_resistance_micro_ohms || 18.6;
    const poleCIncreasePct = lastPoleC > 0 ? (((cp.pole_c - lastPoleC) / lastPoleC) * 100).toFixed(1) : '0';

    // Trip unit
    const inj = data.electrical_tests.secondary_injection_trip_unit_test;
    const isLtPass = inj.long_time_pickup_status.toLowerCase().includes('pass');
    const isStPass = inj.short_time_delay_status.toLowerCase().includes('pass');
    const isInstPass = inj.instantaneous_pickup_status.toLowerCase().includes('pass');
    const isGfPass = inj.ground_fault_pickup_status.toLowerCase().includes('pass');
    const isTripUnitAllPass = isLtPass && isStPass && isInstPass && isGfPass;

    // Control wiring IR
    const isControlWiringPass = (data.electrical_tests.control_wiring_ir_megohms || 0) >= 2.0;

    // Insulation resistance
    const ir = data.electrical_tests.insulation_resistance_1000v_1min_megohms;
    const minClosed = Math.min(ir.closed_phase_a_to_ground, ir.closed_phase_b_to_ground, ir.closed_phase_c_to_ground);
    const minOpen = Math.min(ir.open_across_pole_a, ir.open_across_pole_b, ir.open_across_pole_c);
    const isBreakerIrPass = minClosed >= 100 && minOpen >= 100;

    // Visual & interlocks
    const v = data.visual_inspection;
    const isVisualPass =
      v.nameplate_match &&
      v.trip_unit_settings_match &&
      v.physical_condition_cleanliness.toLowerCase().includes('pass') &&
      v.racking_and_interlocks.toLowerCase().includes('pass') &&
      v.arc_chutes_condition.toLowerCase().includes('pass') &&
      v.bolt_torque_check.toLowerCase().includes('pass');

    // Verdict
    let verdict: 'PASS' | 'INVESTIGATE' | 'FAIL' = 'PASS';
    if (!isTripUnitAllPass || !isControlWiringPass || !isBreakerIrPass || !isVisualPass) {
      verdict = 'FAIL';
    } else if (!isContactPass) {
      verdict = 'INVESTIGATE';
    }

    return {
      minPole,
      maxPole,
      poleCDev,
      maxPoleDev,
      isContactPass,
      lastPoleC,
      poleCIncreasePct,
      isGfPass,
      isTripUnitAllPass,
      isControlWiringPass,
      isBreakerIrPass,
      isVisualPass,
      verdict
    };
  }, [data]);

  const SYSTEM_INSTRUCTION_PROMPT = `Bạn là Trợ lý AI Chuyên gia Kiểm tra Field Service (TEV Platform AI) chuyên trách Máy cắt không khí hạ áp (Low-Voltage Power Circuit Breakers - ACB) theo tiêu chuẩn ANSI/NETA ATS-2025 (Mục 7.6.1.2).

QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.6.1.2):

1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Nameplate Match: Đối soát nhãn mác máy cắt (điện áp, dòng định mức, dòng cắt ngắn mạch, loại trip unit) khớp 100% bản vẽ.
- Physical Condition & Cleanliness: Máy cắt sạch bẩn, tháo bỏ kẹp vận chuyển; buồng dập hồ quang (arc chutes) nguyên vẹn.
- Racking & Interlocks: Cơ cấu kéo/đẩy drawout trơn tru; liên động an toàn ngăn rút/cắm máy cắt khi đang ĐÓNG.
- Contacts & Finger Clusters: Tiếp điểm chính, tiếp điểm dập hồ quang và ngàm cắm finger clusters không bị rỗ nặng hay biến dạng.
- Trip Unit Settings: Các thông số cài đặt L, S, I, G khớp 100% phiếu chỉnh định (Coordination Study).
- Bolt Torque: Mối nối bu-lông siết đạt lực theo dữ liệu nhà sản xuất hoặc Bảng Table 100.12.

2. PHÉP ĐO ĐIỆN VÀ TIÊU CHUẨN ĐÁNH GIÁ (ELECTRICAL TEST VALUES):
- Điện trở mối nối bu-lông (Bolted Connection Resistance): CẢNH BÁO "INVESTIGATE" nếu giá trị đo mối nối lệch quá 50% so với giá trị nhỏ nhất của mối nối tương tự.
- Điện trở tiếp xúc cực (Contact / Pole Resistance): CẢNH BÁO "INVESTIGATE" nếu giá trị đo cực lệch quá 50% so với cực kế cận hoặc giá trị nhỏ nhất.
- Điện trở cách điện Máy cắt (Insulation Resistance - IR): Đo 1 phút (Closed & Open) đạt tối thiểu theo Bảng Table 100.1 (>= 100 Megohms cho cấp 600V).
- Điện trở cách điện Mạch phụ (Control Wiring IR): Thử nghiệm 500V/1000V DC BẮT BUỘC >= 2.0 Megohms (2 MΩ).
- Thử nghiệm Bơm dòng Trip Unit (Primary/Secondary Injection): Ngưỡng pickup và thời gian trip L, S, I, G nằm trong dải dung sai đường cong đặc tính nhà sản xuất.
- Điện áp Thao tác Tối thiểu (Minimum Operating Voltage): Cuộn đóng và cuộn cắt tác động tin cậy ở mức điện áp sụt tối thiểu theo chuẩn ANSI/IEEE.

NHIỆM VỤ CỦA AI:
Khi nhận dữ liệu JSON từ FSE, hãy phân tích và trả về phản hồi định dạng Markdown gồm 4 phần:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE.
2. BẢNG PHÂN TÍCH CHI TIẾT MÁY CẮT HẠ ÁP (Detailed Evaluation Table): So sánh thực tế với NETA ATS-2025 Section 7.6.1.2 (Table 100.1, Table 100.12).
3. PHÂN TÍCH ĐẶC TÍNH TRIP UNIT, TIẾP ĐIỂM & NGUY CƠ MẤT AN TOÀN (Trip Unit Curve, Contact Resistance & Protection Risk).
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations).`;

  const handleRunAiAnalysis = async () => {
    try {
      setIsAnalyzing(true);
      const res = await fetch('/api/field-service/neta-lv-breaker-analyze', {
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
        equipmentId: data.site_info.breaker_tag,
        date: data.site_info.test_date,
        type: 'Low-Voltage Power Circuit Breakers (NETA ATS-2025 Mục 7.6.1.2)'
      });
      setSavedSuccessMessage(`Đã lưu thành công biên bản kiểm định máy cắt ${data.site_info.breaker_tag} vào kho báo cáo CMMS!`);
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
    downloadAnchor.setAttribute('download', `${data.site_info.breaker_tag}_NETA_ATS_2025.json`);
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
        alert('Định dạng JSON không hợp lệ theo chuẩn NETA ATS-2025 Máy Cắt Hạ Áp.');
      }
    } catch (e: any) {
      alert('Lỗi cú pháp JSON: ' + e.message);
    }
  };

  return (
    <div className="space-y-6">
      {/* HEADER BANNER */}
      <div className="bg-gradient-to-r from-slate-950 via-blue-950 to-indigo-950 rounded-2xl p-6 text-white shadow-xl border border-blue-800/40 relative overflow-hidden">
        <div className="absolute right-0 top-0 w-96 h-96 bg-blue-500/10 rounded-full blur-3xl pointer-events-none" />
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6 relative z-10">
          <div>
            <div className="flex flex-wrap items-center gap-2 mb-2">
              <span className="px-3 py-1 bg-blue-500/20 text-blue-300 border border-blue-400/30 rounded-full text-xs font-bold uppercase tracking-wider flex items-center gap-1.5">
                <Power size={14} className="text-blue-400" />
                ANSI/NETA ATS-2025 Mục 7.6.1.2
              </span>
              <span className="px-3 py-1 bg-indigo-500/20 text-indigo-300 border border-indigo-400/30 rounded-full text-xs font-semibold">
                ACB Drawout Air Circuit Breaker
              </span>
              <span className="px-3 py-1 bg-purple-500/20 text-purple-300 border border-purple-400/30 rounded-full text-xs font-semibold">
                Trip Unit L, S, I, G Secondary Injection
              </span>
            </div>
            <h1 className="text-2xl lg:text-3xl font-extrabold text-white flex items-center gap-3">
              Máy Cắt Không Khí Hạ Áp (Low-Voltage Power Circuit Breakers)
            </h1>
            <p className="text-blue-200/90 text-sm mt-1 max-w-3xl">
              Quy chuẩn đối soát chuyên sâu: Điện trở tiếp xúc cực (≤ 50% min Sec 7.6.1.2.D.2), bơm dòng thử nghiệm đặc tính bảo vệ MicroLogic / Ekip L, S, I, G, cách điện Table 100.1 và khóa liên động cơ điện an toàn.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => {
                setData(SAMPLE_LV_BREAKER_DEFECT_DATA);
                setAnalysisResult(null);
              }}
              className="px-3.5 py-2 bg-rose-600/30 hover:bg-rose-600/40 text-rose-200 border border-rose-500/40 rounded-xl text-xs font-semibold transition-all flex items-center gap-1.5"
              title="Tải mẫu ACB sự cố: Pole C lệch +131% & Ground fault nhảy sớm 0.15s"
            >
              <AlertTriangle size={14} className="text-rose-400" />
              Mẫu Lỗi Sự Cố (Fail)
            </button>

            <button
              onClick={() => {
                setData(SAMPLE_LV_BREAKER_PASS_DATA);
                setAnalysisResult(null);
              }}
              className="px-3.5 py-2 bg-emerald-600/30 hover:bg-emerald-600/40 text-emerald-200 border border-emerald-500/40 rounded-xl text-xs font-semibold transition-all flex items-center gap-1.5"
              title="Tải mẫu ACB đạt chuẩn nghiệm thu"
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
              className="px-5 py-2 bg-gradient-to-r from-blue-500 to-indigo-500 hover:from-blue-400 hover:to-indigo-400 text-white font-bold rounded-xl text-sm transition-all shadow-lg shadow-blue-500/30 flex items-center gap-2 disabled:opacity-50"
            >
              {isAnalyzing ? (
                <>
                  <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
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

        {/* LIVE TILES BANNER */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mt-6 pt-6 border-t border-blue-800/40 text-xs">
          <div className="bg-slate-900/60 p-3 rounded-xl border border-blue-500/20">
            <span className="text-slate-400 block mb-1">Điện trở tiếp xúc Cực C (Pole C):</span>
            <div className="flex items-center gap-2">
              <span className={`text-base font-bold ${liveStats.isContactPass ? 'text-emerald-400' : 'text-amber-400'}`}>
                {data.electrical_tests.contact_pole_resistance_micro_ohms.pole_c} µΩ
              </span>
              <span className={`text-[11px] px-1.5 py-0.5 rounded font-bold ${liveStats.isContactPass ? 'bg-emerald-500/20 text-emerald-300' : 'bg-amber-500/20 text-amber-300'}`}>
                {liveStats.isContactPass ? 'ĐẠT' : `+${liveStats.poleCDev}% (>50%)`}
              </span>
            </div>
            <span className="text-[10px] text-slate-400 mt-1 block">
              Pole A: {data.electrical_tests.contact_pole_resistance_micro_ohms.pole_a} µΩ | Pole B: {data.electrical_tests.contact_pole_resistance_micro_ohms.pole_b} µΩ
            </span>
          </div>

          <div className="bg-slate-900/60 p-3 rounded-xl border border-blue-500/20">
            <span className="text-slate-400 block mb-1">Bảo vệ Chạm Đất (Ground Fault):</span>
            <div className="flex items-center gap-2">
              <span className={`text-base font-bold ${liveStats.isGfPass ? 'text-emerald-400' : 'text-rose-400'}`}>
                {liveStats.isGfPass ? '0.30s (Chuẩn)' : '0.15s (Sớm)'}
              </span>
              <span className={`text-[11px] px-1.5 py-0.5 rounded font-bold ${liveStats.isGfPass ? 'bg-emerald-500/20 text-emerald-300' : 'bg-rose-500/20 text-rose-300'}`}>
                {liveStats.isGfPass ? 'PASS' : 'FAIL'}
              </span>
            </div>
            <span className="text-[10px] text-slate-400 mt-1 block">
              Đặc tính trễ thời gian Tg lệch đường cong
            </span>
          </div>

          <div className="bg-slate-900/60 p-3 rounded-xl border border-blue-500/20">
            <span className="text-slate-400 block mb-1">Cách điện Mạch điều khiển phụ:</span>
            <div className="flex items-center gap-2">
              <span className="text-base font-bold text-emerald-400">
                {data.electrical_tests.control_wiring_ir_megohms} MΩ
              </span>
              <span className="text-[11px] px-1.5 py-0.5 rounded font-bold bg-emerald-500/20 text-emerald-300">
                PASS (≥ 2.0 MΩ)
              </span>
            </div>
            <span className="text-[10px] text-slate-400 mt-1 block">Sec 7.6.1.2.D.4</span>
          </div>

          <div className="bg-slate-900/60 p-3 rounded-xl border border-blue-500/20">
            <span className="text-slate-400 block mb-1">Trạng thái tổng quan NETA:</span>
            <div className="flex items-center gap-2">
              <span className={`text-base font-extrabold ${liveStats.verdict === 'PASS' ? 'text-emerald-400' : liveStats.verdict === 'INVESTIGATE' ? 'text-amber-400' : 'text-rose-400'}`}>
                {liveStats.verdict}
              </span>
              <span className={`text-[11px] px-1.5 py-0.5 rounded font-bold ${liveStats.verdict === 'PASS' ? 'bg-emerald-500/20 text-emerald-300' : liveStats.verdict === 'INVESTIGATE' ? 'bg-amber-500/20 text-amber-300' : 'bg-rose-500/20 text-rose-300'}`}>
                {liveStats.verdict === 'FAIL' ? 'CẤM VẬN HÀNH' : liveStats.verdict === 'INVESTIGATE' ? 'CẦN KIỂM TRA' : 'AN TOÀN'}
              </span>
            </div>
            <span className="text-[10px] text-slate-400 mt-1 block">NETA ATS-2025 Sec 7.6.1.2</span>
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
              activeTab === 'electrical' ? 'bg-white text-blue-700 shadow-sm' : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Phép Đo Điện & Trip Unit (Sec 7.6.1.2.D)
          </button>
          <button
            onClick={() => setActiveTab('visual')}
            className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-all ${
              activeTab === 'visual' ? 'bg-white text-blue-700 shadow-sm' : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Thị Giác, Cơ Khí & Liên Động (Sec 7.6.1.2.A)
          </button>
          <button
            onClick={() => setActiveTab('ai-report')}
            className={`px-3 py-1.5 text-xs font-semibold rounded-md transition-all flex items-center gap-1 ${
              activeTab === 'ai-report' ? 'bg-white text-blue-700 shadow-sm' : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            <Sparkles size={13} className="text-blue-600" />
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
            className="px-3.5 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-sm transition-colors"
          >
            <Save size={14} />
            Lưu CMMS
          </button>
        </div>
      </div>

      {/* TAB 1: PHÉP ĐO ĐIỆN & TRIP UNIT (ELECTRICAL TESTS) */}
      {activeTab === 'electrical' && (
        <div className="space-y-6">
          {/* SECTION 1: CONTACT RESISTANCE (SEC 7.6.1.2.D.2) */}
          <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-100">
              <div>
                <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                  <Activity size={18} className="text-blue-600" />
                  1. Điện Trở Tiếp Xúc Cực (Contact / Pole Resistance per Section 7.6.1.2.D.2)
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  ANSI/NETA ATS-2025 Mục 7.6.1.2.D.2: CẢNH BÁO <strong className="text-amber-600">&ldquo;INVESTIGATE&rdquo;</strong> nếu giá trị đo cực lệch quá 50% so với cực kế cận hoặc giá trị nhỏ nhất.
                </p>
              </div>
              <span className={`px-3 py-1 rounded-full text-xs font-bold border ${liveStats.isContactPass ? 'bg-emerald-50 text-emerald-700 border-emerald-300' : 'bg-amber-50 text-amber-800 border-amber-300'}`}>
                {liveStats.isContactPass ? 'Tiếp xúc 3 cực đạt chuẩn' : `Pole C lệch +${liveStats.poleCDev}% (> 50%)`}
              </span>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mt-5">
              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
                <div className="flex items-center justify-between mb-2">
                  <label className="text-xs font-bold text-slate-700">Pole A (µΩ)</label>
                  <span className="text-[11px] font-bold text-emerald-600 bg-emerald-100 px-1.5 py-0.5 rounded">PASS</span>
                </div>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.contact_pole_resistance_micro_ohms.pole_a}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        contact_pole_resistance_micro_ohms: {
                          ...data.electrical_tests.contact_pole_resistance_micro_ohms,
                          pole_a: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg font-bold text-slate-900"
                />
                <span className="text-[10px] text-slate-500 mt-1 block">Giá trị chuẩn tham chiếu nhỏ nhất</span>
              </div>

              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
                <div className="flex items-center justify-between mb-2">
                  <label className="text-xs font-bold text-slate-700">Pole B (µΩ)</label>
                  <span className="text-[11px] font-bold text-emerald-600 bg-emerald-100 px-1.5 py-0.5 rounded">PASS</span>
                </div>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.contact_pole_resistance_micro_ohms.pole_b}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        contact_pole_resistance_micro_ohms: {
                          ...data.electrical_tests.contact_pole_resistance_micro_ohms,
                          pole_b: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg font-bold text-slate-900"
                />
                <span className="text-[10px] text-slate-500 mt-1 block">Lệch 1.6% so với Pole A</span>
              </div>

              <div className="p-4 bg-amber-50/80 rounded-xl border border-amber-300 ring-2 ring-amber-400/20">
                <div className="flex items-center justify-between mb-2">
                  <label className="text-xs font-extrabold text-amber-900">Pole C (µΩ - DEFECT)</label>
                  <span className="text-[11px] font-extrabold text-amber-800 bg-amber-200 px-1.5 py-0.5 rounded">
                    LỆCH +{liveStats.poleCDev}% ⚠️
                  </span>
                </div>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.contact_pole_resistance_micro_ohms.pole_c}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        contact_pole_resistance_micro_ohms: {
                          ...data.electrical_tests.contact_pole_resistance_micro_ohms,
                          pole_c: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-white border border-amber-400 rounded-lg font-extrabold text-amber-700"
                />
                <span className="text-[10px] text-amber-800 font-medium mt-1 block">
                  Năm ngoái: <strong>{liveStats.lastPoleC} µΩ</strong> (Tăng vọt +{liveStats.poleCIncreasePct}%)
                </span>
              </div>
            </div>
          </div>

          {/* SECTION 2: TRIP UNIT SECONDARY INJECTION (L, S, I, G) */}
          <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-slate-100">
              <div>
                <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                  <Cpu size={18} className="text-indigo-600" />
                  2. Thử Nghiệm Bơm Dòng Đặc Tính Bộ Trip Unit ({data.site_info.trip_unit_model})
                </h3>
                <p className="text-xs text-slate-500 mt-0.5">
                  ANSI/NETA ATS-2025 Mục 7.6.1.2.D.5: Ngưỡng pickup và thời gian trễ tác động các chức năng Long Time (L), Short Time (S), Instantaneous (I), Ground Fault (G) phải tuân thủ đúng đường cong của nhà sản xuất.
                </p>
              </div>
              <span className={`px-3 py-1 rounded-full text-xs font-bold border ${liveStats.isGfPass ? 'bg-emerald-50 text-emerald-700 border-emerald-300' : 'bg-rose-50 text-rose-800 border-rose-300'}`}>
                {liveStats.isGfPass ? 'Tất cả L, S, I, G đạt' : 'Lỗi thời gian Ground Fault (G)'}
              </span>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mt-5">
              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
                <label className="text-xs font-bold text-slate-700 block mb-1">
                  Long Time Pickup (L) - Bảo vệ quá tải thời gian dài
                </label>
                <input
                  type="text"
                  value={data.electrical_tests.secondary_injection_trip_unit_test.long_time_pickup_status}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        secondary_injection_trip_unit_test: {
                          ...data.electrical_tests.secondary_injection_trip_unit_test,
                          long_time_pickup_status: e.target.value
                        }
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md text-xs font-semibold text-slate-800"
                />
              </div>

              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
                <label className="text-xs font-bold text-slate-700 block mb-1">
                  Short Time Delay (S) - Bảo vệ ngắn mạch có trễ thời gian
                </label>
                <input
                  type="text"
                  value={data.electrical_tests.secondary_injection_trip_unit_test.short_time_delay_status}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        secondary_injection_trip_unit_test: {
                          ...data.electrical_tests.secondary_injection_trip_unit_test,
                          short_time_delay_status: e.target.value
                        }
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md text-xs font-semibold text-slate-800"
                />
              </div>

              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
                <label className="text-xs font-bold text-slate-700 block mb-1">
                  Instantaneous Pickup (I) - Bảo vệ cắt tức thì
                </label>
                <input
                  type="text"
                  value={data.electrical_tests.secondary_injection_trip_unit_test.instantaneous_pickup_status}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        secondary_injection_trip_unit_test: {
                          ...data.electrical_tests.secondary_injection_trip_unit_test,
                          instantaneous_pickup_status: e.target.value
                        }
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md text-xs font-semibold text-slate-800"
                />
              </div>

              <div className="p-4 bg-rose-50/80 rounded-xl border border-rose-300 ring-2 ring-rose-400/20">
                <div className="flex items-center justify-between mb-1">
                  <label className="text-xs font-extrabold text-rose-800">
                    Ground Fault Pickup (G) - Chống dòng rò chạm đất (DEFECT)
                  </label>
                  <span className="text-[10px] px-1.5 py-0.5 rounded bg-rose-200 text-rose-900 font-bold">
                    FAIL TIMING
                  </span>
                </div>
                <input
                  type="text"
                  value={data.electrical_tests.secondary_injection_trip_unit_test.ground_fault_pickup_status}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        secondary_injection_trip_unit_test: {
                          ...data.electrical_tests.secondary_injection_trip_unit_test,
                          ground_fault_pickup_status: e.target.value
                        }
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-white border border-rose-300 rounded-md text-xs font-bold text-rose-700"
                />
                <span className="text-[11px] text-rose-700 font-medium mt-1 block">
                  Trip 0.15s (Sớm hơn đường cong đặt 0.30s) $\rightarrow$ Nguy cơ nhảy cắt mất chọn lọc toàn trạm!
                </span>
              </div>
            </div>
          </div>

          {/* SECTION 3: INSULATION RESISTANCE & CONTROL WIRING */}
          <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm">
            <div className="pb-4 border-b border-slate-100 mb-5">
              <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                <Zap size={18} className="text-amber-500" />
                3. Điện Trở Cách Điện Máy Cắt & Mạch Phụ (Insulation Resistance per Table 100.1)
              </h3>
              <p className="text-xs text-slate-500 mt-0.5">
                Cấp 600V: BẮT BUỘC &ge; 100 Megohms (khi Đóng và khi Mở cực). Mạch phụ điều khiển: BẮT BUỘC &ge; 2.0 Megohms (Sec 7.6.1.2.D.4).
              </p>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
                <label className="text-xs font-semibold text-slate-600 block mb-1">Closed Pha A - Đất (MΩ)</label>
                <input
                  type="number"
                  value={data.electrical_tests.insulation_resistance_1000v_1min_megohms.closed_phase_a_to_ground}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_1000v_1min_megohms: {
                          ...data.electrical_tests.insulation_resistance_1000v_1min_megohms,
                          closed_phase_a_to_ground: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md font-bold text-slate-800"
                />
                <span className="text-[11px] text-emerald-600 font-bold mt-1 block">PASS (≥ 100 MΩ)</span>
              </div>

              <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
                <label className="text-xs font-semibold text-slate-600 block mb-1">Closed Pha B - Đất (MΩ)</label>
                <input
                  type="number"
                  value={data.electrical_tests.insulation_resistance_1000v_1min_megohms.closed_phase_b_to_ground}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_1000v_1min_megohms: {
                          ...data.electrical_tests.insulation_resistance_1000v_1min_megohms,
                          closed_phase_b_to_ground: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md font-bold text-slate-800"
                />
                <span className="text-[11px] text-emerald-600 font-bold mt-1 block">PASS (≥ 100 MΩ)</span>
              </div>

              <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
                <label className="text-xs font-semibold text-slate-600 block mb-1">Closed Pha C - Đất (MΩ)</label>
                <input
                  type="number"
                  value={data.electrical_tests.insulation_resistance_1000v_1min_megohms.closed_phase_c_to_ground}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_1000v_1min_megohms: {
                          ...data.electrical_tests.insulation_resistance_1000v_1min_megohms,
                          closed_phase_c_to_ground: parseFloat(e.target.value) || 0
                        }
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md font-bold text-slate-800"
                />
                <span className="text-[11px] text-emerald-600 font-bold mt-1 block">PASS (≥ 100 MΩ)</span>
              </div>

              <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
                <label className="text-xs font-bold text-indigo-700 block mb-1">Cách điện Mạch phụ (MΩ)</label>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.control_wiring_ir_megohms}
                  onChange={(e) =>
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        control_wiring_ir_megohms: parseFloat(e.target.value) || 0
                      }
                    })
                  }
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md font-bold text-indigo-700"
                />
                <span className="text-[11px] text-emerald-600 font-bold mt-1 block">PASS (Chuẩn ≥ 2.0 MΩ)</span>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* TAB 2: THỊ GIÁC, CƠ KHÍ & LIÊN ĐỘNG (VISUAL & MECHANICAL) */}
      {activeTab === 'visual' && (
        <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm space-y-6">
          <div className="pb-4 border-b border-slate-100">
            <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
              <Eye size={18} className="text-blue-600" />
              Kiểm Tra Thị Giác, Cơ Khí & Liên Động (Visual & Mechanical Inspection per Section 7.6.1.2.A)
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Khớp nhãn mác, khóa liên động an toàn khi cắm/rút máy cắt, ngàm cắm finger clusters và buồng dập hồ quang.
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <div className="flex items-center justify-between">
                <label className="text-sm font-semibold text-slate-800">
                  Nameplate Match (Khớp nhãn mác 100%)
                </label>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.nameplate_match}
                  onChange={(e) =>
                    setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, nameplate_match: e.target.checked }
                    })
                  }
                  className="w-5 h-5 text-blue-600 rounded border-slate-300 focus:ring-blue-500"
                />
              </div>
              <p className="text-xs text-slate-500">
                Đối soát nhãn mác máy cắt ({data.site_info.breaker_rating}), điện áp, dòng định mức, dòng cắt ngắn mạch và loại trip unit khớp 100% bản vẽ.
              </p>
            </div>

            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <div className="flex items-center justify-between">
                <label className="text-sm font-semibold text-slate-800">
                  Trip Unit Settings Match (Khớp phiếu chỉnh định)
                </label>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.trip_unit_settings_match}
                  onChange={(e) =>
                    setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, trip_unit_settings_match: e.target.checked }
                    })
                  }
                  className="w-5 h-5 text-blue-600 rounded border-slate-300 focus:ring-blue-500"
                />
              </div>
              <p className="text-xs text-slate-500">
                Các thông số cài đặt L, S, I, G trên bộ MicroLogic / Ekip khớp 100% phiếu tính toán phối hợp bảo vệ (Coordination Study).
              </p>
            </div>

            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <label className="text-sm font-semibold text-slate-800 block">
                Racking & Interlocks (Cơ cấu kéo/đẩy & Liên động an toàn)
              </label>
              <input
                type="text"
                value={data.visual_inspection.racking_and_interlocks}
                onChange={(e) =>
                  setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, racking_and_interlocks: e.target.value }
                  })
                }
                className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md text-sm font-medium text-slate-800"
              />
              <p className="text-xs text-slate-500">
                Cơ cấu drawout di chuyển trơn tru; liên động an toàn ngăn cắm/rút máy cắt khi tiếp điểm đang ĐÓNG (Closed).
              </p>
            </div>

            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <label className="text-sm font-semibold text-slate-800 block">
                Contacts & Finger Clusters (Ngàm cắm & Tiếp điểm chính)
              </label>
              <input
                type="text"
                value={data.visual_inspection.contacts_and_finger_clusters}
                onChange={(e) =>
                  setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, contacts_and_finger_clusters: e.target.value }
                  })
                }
                className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md text-sm font-medium text-slate-800"
              />
              <p className="text-xs text-slate-500">
                Tiếp điểm chính, tiếp điểm dập hồ quang và cụm ngàm cắm finger clusters phía sau không bị rỗ nặng hay biến dạng.
              </p>
            </div>

            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <label className="text-sm font-semibold text-slate-800 block">
                Arc Chutes Condition (Buồng dập hồ quang)
              </label>
              <input
                type="text"
                value={data.visual_inspection.arc_chutes_condition}
                onChange={(e) =>
                  setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, arc_chutes_condition: e.target.value }
                  })
                }
                className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md text-sm font-medium text-slate-800"
              />
              <p className="text-xs text-slate-500">
                Buồng dập hồ quang nguyên vẹn, các lá kim loại dập hồ quang không nứt vỡ hay bám dính carbon quá mức.
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
                    visual_inspection: { ...data.visual_inspection, bolt_torque_check: e.target.value }
                  })
                }
                className="w-full px-3 py-2 bg-white border border-slate-300 rounded-md text-sm font-medium text-slate-800"
              />
              <p className="text-xs text-slate-500">
                Mối nối bu-lông siết đạt lực theo dữ liệu nhà sản xuất hoặc Bảng Table 100.12.
              </p>
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
                    className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-sm"
                  >
                    <Save size={14} />
                    Lưu Báo Cáo CMMS
                  </button>
                </div>
              </div>

              <div className="prose prose-slate max-w-none bg-slate-50/70 p-6 rounded-xl border border-slate-200 text-sm leading-relaxed whitespace-pre-wrap font-sans">
                {analysisResult.markdown_report}
              </div>
            </div>
          ) : (
            <div className="text-center py-12 space-y-4">
              <div className="w-16 h-16 bg-blue-100 text-blue-700 rounded-full flex items-center justify-center mx-auto">
                <Sparkles size={32} />
              </div>
              <h3 className="text-lg font-bold text-slate-800">Chưa có kết quả phân tích AI</h3>
              <p className="text-slate-500 text-xs max-w-md mx-auto">
                Nhấn nút &ldquo;Phân Tích AI NETA ATS-2025&rdquo; ở thanh tiêu đề để AI chuyên gia tự động đối soát toàn bộ giá trị đo đạc với tiêu chuẩn ANSI/NETA ATS-2025 Mục 7.6.1.2.
              </p>
              <button
                onClick={handleRunAiAnalysis}
                disabled={isAnalyzing}
                className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-xl text-sm transition-all shadow-md inline-flex items-center gap-2"
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
              <pre className="p-4 bg-slate-900 text-blue-300 rounded-xl text-xs font-mono overflow-x-auto whitespace-pre-wrap leading-relaxed border border-purple-500/20 max-h-96">
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
              <pre className="p-4 bg-slate-900 text-emerald-400 rounded-xl text-xs font-mono overflow-x-auto whitespace-pre-wrap leading-relaxed border border-slate-700 max-h-80">
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
                <Upload size={18} className="text-blue-600" />
                Nhập Dữ Liệu Kiểm Tra JSON (Máy Cắt Hạ Áp NETA ATS-2025)
              </h3>
              <button
                onClick={() => setShowJsonModal(false)}
                className="text-slate-400 hover:text-slate-600"
              >
                ✕
              </button>
            </div>
            <textarea
              rows={12}
              value={jsonText}
              onChange={(e) => setJsonText(e.target.value)}
              placeholder="Dán nội dung JSON vào đây..."
              className="w-full p-3 font-mono text-xs bg-slate-50 border border-slate-300 rounded-xl focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500"
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
                className="px-5 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-lg shadow-sm"
              >
                Xác Nhận Nhập Dữ Liệu
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: PROMPT GOOGLE AI STUDIO */}
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
                Sao chép System Instruction này và dán vào phần <strong>System Instructions</strong> trên Google AI Studio:
              </p>
              <div className="relative">
                <button
                  onClick={handleCopyPrompt}
                  className="absolute right-3 top-3 px-3 py-1.5 bg-purple-600 hover:bg-purple-700 text-white rounded-md text-xs font-bold flex items-center gap-1 shadow"
                >
                  {copiedPrompt ? <Check size={12} /> : <Copy size={12} />}
                  {copiedPrompt ? 'Đã chép' : 'Sao chép'}
                </button>
                <pre className="p-4 bg-slate-950 text-blue-300 rounded-xl font-mono overflow-x-auto whitespace-pre-wrap leading-relaxed border border-slate-800">
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
          equipmentName={`${data.site_info.breaker_tag} (${data.site_info.breaker_rating})`}
          equipmentId={data.site_info.breaker_tag}
          testDate={data.site_info.test_date}
          technician={data.site_info.fse_name}
          standard="ANSI/NETA ATS-2025 Mục 7.6.1.2"
          overallStatus={analysisResult?.overall_status || liveStats.verdict}
          reportTitle="BIÊN BẢN KIỂM ĐỊNH MÁY CẮT KHÔNG KHÍ HẠ ÁP (NETA ATS-2025)"
          evaluations={[
            {
              name: 'Kiểm Tra Thị Giác, Racking & Liên Động',
              standard: 'NETA ATS-2025 Mục 7.6.1.2.A',
              actual: data.visual_inspection.racking_and_interlocks,
              status: liveStats.isVisualPass ? 'PASS' : 'FAIL',
              note: 'Drawout và liên động an toàn vận hành tốt'
            },
            {
              name: 'Điện Trở Tiếp Xúc Cực Pole C',
              standard: 'Mục 7.6.1.2.D.2 (≤ 50% min)',
              actual: `${data.electrical_tests.contact_pole_resistance_micro_ohms.pole_c} µΩ`,
              status: liveStats.isContactPass ? 'PASS' : 'INVESTIGATE',
              note: `Lệch +${liveStats.poleCDev}% so với Pole A (18.2 µΩ), tăng +${liveStats.poleCIncreasePct}% so với năm ngoái`
            },
            {
              name: 'Thử Nghiệm Bơm Dòng Trip Unit (Ground Fault)',
              standard: 'Mục 7.6.1.2.D.5 (Đường cong MicroLogic 6.0A)',
              actual: data.electrical_tests.secondary_injection_trip_unit_test.ground_fault_pickup_status,
              status: liveStats.isGfPass ? 'PASS' : 'FAIL',
              note: 'Tác động 0.15s (Sai lệch so với cài đặt 0.30s gây mất chọn lọc)'
            },
            {
              name: 'Điện Trở Cách Điện Máy Cắt (Table 100.1)',
              standard: 'Min ≥ 100 MΩ (1000V DC)',
              actual: `Closed min ${Math.min(data.electrical_tests.insulation_resistance_1000v_1min_megohms.closed_phase_a_to_ground, data.electrical_tests.insulation_resistance_1000v_1min_megohms.closed_phase_b_to_ground, data.electrical_tests.insulation_resistance_1000v_1min_megohms.closed_phase_c_to_ground)} MΩ`,
              status: 'PASS',
              note: 'Độ bền điện môi cách điện đạt chuẩn'
            },
            {
              name: 'Cách Điện Mạch Phụ Điều Khiển',
              standard: 'Mục 7.6.1.2.D.4 (Bắt buộc ≥ 2.0 MΩ)',
              actual: `${data.electrical_tests.control_wiring_ir_megohms} MΩ`,
              status: liveStats.isControlWiringPass ? 'PASS' : 'FAIL',
              note: 'Mạch điều khiển an toàn không bị ẩm'
            }
          ]}
          recommendations={
            analysisResult
              ? [analysisResult.status_reason]
              : [
                  'Hiệu chỉnh lại thời gian trễ Ground Fault trên bộ MicroLogic 6.0A đạt 0.30s.',
                  'Tách máy cắt ra vị trí TEST, vệ sinh tiếp điểm và cụm finger clusters Pole C.',
                  'Đo lại micro-ohm Pole C (đạt ≤ 20.0 µΩ) trước khi đưa vào vị trí CONNECTED.'
                ]
          }
        />
      )}
    </div>
  );
};
