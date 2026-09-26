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
  Flame,
  ArrowRight
} from 'lucide-react';
import {
  NetaDcMotorInputPayload,
  NetaDcMotorEvaluationResult,
  performDeterministicDcMotorAnalysis
} from '../../server/netaDcMotorAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';
import { FseEngineerDropdown } from './FseEngineerDropdown';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

const SAMPLE_DC_MOTOR_DATA: NetaDcMotorInputPayload = {
  site_info: {
    project_name: "Nha May Thep TEV Heavy Industry",
    motor_tag: "MOT-DC-ROLLING-MILL-01",
    motor_rating: "250kW (335HP) 440V Armature, 220V Field, 1150RPM DC Shunt Motor",
    fse_name: "Nguyen Van Y",
    test_date: "2026-09-26",
    machine_type: "DC Motor",
    rated_kw: 250,
    armature_voltage_v: 440,
    armature_current_a: 615,
    field_voltage_v: 220,
    field_current_a: 5.2,
    rated_rpm: 1150
  },
  visual_inspection: {
    nameplate_match: true,
    physical_condition_cleanliness: "Pass",
    commutator_and_brush_rigging: "Pass (Smooth surface, correct brush angle)",
    tachometer_generator_check: "Pass",
    anchorage_alignment_grounding: "Pass",
    bolt_torque_check: "Pass",
    shipping_braces_removed: true,
    thermographic_survey_check: "Pass (No hotspot detected per Table 100.18)"
  },
  electrical_tests: {
    bolted_resistance_micro_ohms: [10.1, 10.4, 10.2],
    insulation_resistance_40c_megohms: {
      armature_1min_ir_1000v: 320,
      armature_10min_ir_1000v: 890,
      armature_pi_value: 2.78,
      armature_dar_value: 1.85,
      shunt_field_1min_ir_1000v: 410,
      shunt_field_10min_ir_1000v: 1150,
      shunt_field_pi_value: 2.80
    },
    field_pole_ac_voltage_drop_volts: {
      pole_1: 18.2,
      pole_2: 18.4,
      pole_3: 18.1,
      pole_4: 18.3,
      average_drop_volts: 18.25
    },
    armature_bar_to_bar_resistance_micro_ohms: {
      min_bar_resistance_micro_ohms: 242.0,
      max_bar_resistance_micro_ohms: 268.5,
      max_deviation_percent: 8.7
    },
    surge_comparison_test: "Pass",
    running_current_and_field_check: "Pass (No-load Armature 14.2A, Field 4.8A at 220V)",
    unloaded_vibration_table_100_10: {
      velocity_in_sec_pk: 0.08,
      nema_limit_in_sec_pk: 0.15
    }
  },
  previous_test_data: {
    last_test_date: "2025-09-25",
    last_armature_pi: 2.9,
    last_vibration_in_sec: 0.075,
    last_bar_deviation_percent: 3.2
  }
};

export const NetaDcMotorChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaDcMotorInputPayload>(SAMPLE_DC_MOTOR_DATA);
  const [activeTab, setActiveTab] = useState<'site' | 'visual' | 'electrical' | 'ai'>('electrical');
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaDcMotorEvaluationResult | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showPromptModal, setShowPromptModal] = useState(false);
  const [jsonText, setJsonText] = useState('');
  const [copied, setCopied] = useState(false);
  const [promptCopied, setPromptCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live Calculations according to ANSI/NETA ATS-2025 Section 7.15.3
  const liveStats = useMemo(() => {
    const analysis = performDeterministicDcMotorAnalysis(data);

    // Bolted connection
    const bolts = data.electrical_tests.bolted_resistance_micro_ohms || [];
    let boltDev = 0;
    if (bolts.length > 1) {
      const minB = Math.min(...bolts);
      const maxB = Math.max(...bolts);
      boltDev = minB > 0 ? parseFloat((((maxB - minB) / minB) * 100).toFixed(1)) : 0;
    }

    // Pole-to-pole drop
    const poleData = data.electrical_tests.field_pole_ac_voltage_drop_volts;
    const poles = [poleData.pole_1, poleData.pole_2, poleData.pole_3, poleData.pole_4].filter(v => typeof v === 'number');
    const avgPole = poleData.average_drop_volts || (poles.reduce((a, b) => a + b, 0) / (poles.length || 1));
    let maxPoleDev = 0;
    poles.forEach(p => {
      const dev = Math.abs(p - avgPole) / avgPole * 100;
      if (dev > maxPoleDev) maxPoleDev = dev;
    });

    // Bar to bar
    const barDev = data.electrical_tests.armature_bar_to_bar_resistance_micro_ohms.max_deviation_percent;
    const barExceeded = barDev > 5.0;

    // Vibration
    const vib = data.electrical_tests.unloaded_vibration_table_100_10;
    const vibExceeded = vib.velocity_in_sec_pk > vib.nema_limit_in_sec_pk;

    // PI / DAR
    const pi = data.electrical_tests.insulation_resistance_40c_megohms.armature_pi_value;
    const dar = data.electrical_tests.insulation_resistance_40c_megohms.armature_dar_value;
    const isLarge = (data.site_info.rated_kw || 250) > 150;
    const piExceeded = isLarge ? (pi !== undefined && pi < 2.0) : (dar !== undefined && dar < 1.4);

    return {
      analysis,
      boltDev,
      boltWarning: boltDev > 50.0,
      maxPoleDev: parseFloat(maxPoleDev.toFixed(2)),
      poleWarning: maxPoleDev > 10.0,
      barDev,
      barExceeded,
      vibVal: vib.velocity_in_sec_pk,
      vibLimit: vib.nema_limit_in_sec_pk,
      vibExceeded,
      pi,
      dar,
      isLarge,
      piExceeded,
      isFail: analysis.overall_status === 'FAIL',
      isInvestigate: analysis.overall_status === 'INVESTIGATE'
    };
  }, [data]);

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    setAnalysisResult(null);
    setSavedSuccessMessage(null);

    try {
      const response = await fetch('/api/field-service/neta-dc-motor-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });

      const resJson = await response.json();
      if (resJson.success) {
        setAnalysisResult(resJson.data);
        setActiveTab('ai');
      } else {
        alert('Lỗi phân tích: ' + resJson.error);
      }
    } catch (err: any) {
      console.error('Lỗi khi gọi API phân tích NETA DC Motor:', err);
      alert('Không thể kết nối đến máy chủ AI: ' + err.message);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleSaveToCmms = () => {
    const evalData = analysisResult || liveStats.analysis;
    if (onSaveReport) {
      onSaveReport({
        data,
        analysisResult: evalData,
        equipmentId: data.site_info.motor_tag,
        date: data.site_info.test_date,
        type: 'DC Motors & Generators (NETA ATS-2025 Sec 7.15.3)'
      });
      setSavedSuccessMessage(`Đã lưu biên bản kiểm định máy điện DC [${data.site_info.motor_tag}] vào kho báo cáo CMMS!`);
    }
  };

  const handleResetToSample = () => {
    setData(SAMPLE_DC_MOTOR_DATA);
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

  const googleStudioPrompt = useMemo(() => {
    return `QUY TRẮC ĐÁNH GIÁ VÀ TIÊU CHUẨN THAM CHIẾU (NETA ATS-2025 Section 7.15.3 - DC Motors & Generators):

1. KIỂM TRA THỊ GIÁC & CƠ KHÍ (VISUAL & MECHANICAL INSPECTION):
- Nameplate Match: Đối soát nhãn mác máy điện DC (kW/HP, điện áp/dòng phần ứng Armature, điện áp/dòng kích từ Field, RPM) khớp 100% bản vẽ thiết kế.
- Physical Condition & Cleanliness: Động cơ/máy phát DC sạch bẩn; cuộn dây phần ứng, cuộn cực từ (field poles), vách gió, phin lọc nguyên vẹn; tháo bỏ kẹp vận chuyển.
- Commutator & Brushes: Bề mặt cổ góp nhẵn phẳng, không rãnh sâu/cháy xém; độ chặt phiến cổ góp (bar tightness), chổi than, giá đỡ chổi (brush rigging) và máy phát tốc (tachometer generator) đạt dung sai nhà sản xuất.
- Anchorage & Alignment: Neo chân đế chắc chắn, căn chỉnh định tâm đồng trục và tiếp địa vỏ máy đầy đủ.
- Bolt Torque: Mối nối bu-lông siết đạt lực theo dữ liệu nhà sản xuất hoặc Bảng Table 100.12.
- Thermographic Survey: Kết quả khảo sát nhiệt tuân thủ Section 9 / Table 100.18.

2. PHÉP ĐO ĐIỆN VÀ TIÊU CHUẨN ĐÁNH GIÁ (ELECTRICAL TEST VALUES):
- Điện trở mối nối bu-lông (Bolted Connection Resistance): CẢNH BÁO "INVESTIGATE" nếu giá trị đo mối nối lệch quá 50% so với giá trị nhỏ nhất của mối nối tương tự.
- Điện trở cách điện Cuộn dây (Insulation Resistance per IEEE 43 / Table 100.11 quy đổi 40°C):
  + Cuộn phần ứng Armature và các cuộn kích từ Shunt/Series/Interpole Field đạt Bảng Table 100.11.
  + Máy > 150 kW (200 HP): Đo 10 phút -> Chỉ số Phân cực BẮT BUỘC PI >= 2.0.
  + Máy <= 150 kW (200 HP): Đo 1 phút -> Hệ số Hấp thụ BẮT BUỘC DAR >= 1.4.
- Thử sụt áp AC Cực Từ Kích từ (Pole-to-Pole AC Voltage Drop): Độ lệch điện áp sụt trên từng cực từ kích từ KHÔNG ĐƯỢC vượt quá 10% (> 10%) so với giá trị trung bình toàn bộ các cực.
- Điện trở Phiến-đến-Phiến Cổ góp (Armature Bar-to-Bar Resistance): Giá trị điện trở giữa các phiến đồng kế cận trên cổ góp phần ứng KHÔNG ĐƯỢC sai lệch vượt quá 5% (> 5%).
- Cực tính & Thử Xung (Surge Comparison & Polarity): Cực tính tương đối giữa cuộn Shunt Field & Series Field đúng thiết kế; dạng sóng thử xung lồng ghép hoàn toàn.
- Dòng điện Vận hành & Kích từ (Running & Field Currents): Dòng phần ứng không tải và thông số kích từ đo được so sánh tương đồng với dữ liệu nhãn máy.
- Thử nghiệm Độ rung (Unloaded Vibration Test): Biên độ rung động không tải tháo khớp nối tuân thủ Bảng Table 100.10.

NHIỆM VỤ CỦA AI:
Khi nhận dữ liệu JSON từ FSE, hãy phân tích và trả về phản hồi định dạng Markdown gồm 4 phần:
1. TRẠNG THÁI TỔNG QUAN (Overall Status): PASS, FAIL, hoặc INVESTIGATE.
2. BẢNG PHÂN TÍCH CHI TIẾT MÁY ĐIỆN DC (Detailed Evaluation Table): So sánh thực tế với NETA ATS-2025 Section 7.15.3 (Table 100.10, Table 100.11, Table 100.12).
3. PHÂN TÍCH CỔ GÓP, CÁCH ĐIỆN CỰC TỪ & NGUY CƠ PHÓNG TIA LỬA ĐIỆN (Commutator, Field Pole & Sparking Risk).
4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations).

DỮ LIỆU JSON ĐẦU VÀO ĐỂ MÔ PHỎNG:
${JSON.stringify(data, null, 2)}`;
  }, [data]);

  const handleCopyPrompt = () => {
    navigator.clipboard.writeText(googleStudioPrompt);
    setPromptCopied(true);
    setTimeout(() => setPromptCopied(false), 2000);
  };

  return (
    <div className="space-y-6">
      {/* HEADER BANNER */}
      <div className="bg-gradient-to-r from-slate-900 via-rose-950 to-slate-900 rounded-2xl p-6 text-white border border-rose-800/40 shadow-xl relative overflow-hidden">
        <div className="absolute -right-10 -bottom-10 opacity-10 pointer-events-none">
          <RotateCcw size={280} className="text-rose-400" />
        </div>

        <div className="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-4 relative z-10">
          <div className="space-y-2">
            <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-rose-500/20 text-rose-300 border border-rose-500/30 text-xs font-semibold uppercase tracking-wider">
              <Zap size={13} className="text-rose-400 animate-pulse" />
              ANSI/NETA ATS-2025 Mục 7.15.3 & IEEE Std 43
            </div>
            <h1 className="text-2xl lg:text-3xl font-bold tracking-tight text-white flex items-center gap-3">
              Động Cơ & Máy Phát Điện Một Chiều DC
              <span className="text-sm px-2.5 py-0.5 rounded-md bg-rose-500/30 text-rose-200 border border-rose-400/40 font-mono">
                {data.site_info.motor_tag}
              </span>
            </h1>
            <p className="text-sm text-slate-300 max-w-2xl leading-relaxed">
              Kiểm định chuyên sâu hệ thống máy điện DC (Rolling Mill Motors, DC Generators, Shunt/Series Fields): Cổ góp & chổi than, điện trở phiến Bar-to-Bar (≤ 5%), sụt áp cực từ kích từ (≤ 10%), cách điện IEEE 43 (PI ≥ 2.0 / DAR ≥ 1.4) và độ rung không tải Table 100.10.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2 self-stretch lg:self-auto justify-end">
            <button
              onClick={handleResetToSample}
              className="px-3 py-2 bg-slate-800/90 hover:bg-slate-700 text-slate-200 text-xs font-medium rounded-lg border border-slate-700 transition-colors flex items-center gap-1.5 shadow-sm"
              title="Khôi phục dữ liệu kiểm định mẫu chuẩn NETA ATS-2025"
            >
              <RotateCcw size={14} />
              Mẫu Chuẩn FSE
            </button>
            <button
              onClick={() => setShowPromptModal(true)}
              className="px-3 py-2 bg-purple-900/40 hover:bg-purple-800/50 text-purple-200 text-xs font-medium rounded-lg border border-purple-500/30 transition-colors flex items-center gap-1.5 shadow-sm"
              title="Xem và sao chép Prompt mô phỏng trên Google AI Studio"
            >
              <Sparkles size={14} className="text-purple-400" />
              Prompt AI Studio
            </button>
            <button
              onClick={handleOpenJsonModal}
              className="px-3 py-2 bg-slate-800/90 hover:bg-slate-700 text-slate-200 text-xs font-medium rounded-lg border border-slate-700 transition-colors flex items-center gap-1.5 shadow-sm"
              title="Sao chép / Nhập file JSON kiểm định"
            >
              <Layers size={14} />
              JSON Dữ Liệu
            </button>
            <button
              onClick={() => setShowPrintModal(true)}
              className="px-3.5 py-2 bg-slate-800/90 hover:bg-slate-700 text-slate-200 text-xs font-medium rounded-lg border border-slate-700 transition-colors flex items-center gap-1.5 shadow-sm"
            >
              <Printer size={14} />
              In Báo Cáo
            </button>
            <button
              onClick={handleSaveToCmms}
              className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded-lg shadow-md transition-colors flex items-center gap-1.5"
            >
              <Save size={14} />
              Lưu CMMS
            </button>
          </div>
        </div>

        {/* QUICK STATUS BAR */}
        <div className="mt-5 pt-4 border-t border-slate-800/80 grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3 text-xs">
          {/* Bar-to-Bar Deviation */}
          <div className={`p-2.5 rounded-xl border ${liveStats.barExceeded ? 'bg-rose-950/70 border-rose-500/60 text-rose-200' : 'bg-slate-800/60 border-slate-700 text-slate-300'}`}>
            <div className="text-[10px] text-slate-400 flex items-center justify-between">
              <span>Lệch Trở Bar-to-Bar:</span>
              <span className="font-mono text-[9px]">NETA ≤ 5.0%</span>
            </div>
            <div className="text-base font-bold font-mono mt-0.5 flex items-center gap-1.5">
              {liveStats.barDev.toFixed(1)}%
              {liveStats.barExceeded ? (
                <span className="px-1.5 py-0.2 rounded text-[10px] bg-rose-600 text-white font-sans font-bold">FAIL</span>
              ) : (
                <span className="px-1.5 py-0.2 rounded text-[10px] bg-emerald-600 text-white font-sans font-bold">PASS</span>
              )}
            </div>
            <div className="text-[10px] opacity-80 mt-0.5 truncate">
              {liveStats.barExceeded ? 'Hồ quang / hở riser!' : 'Phiến cổ góp đồng đều'}
            </div>
          </div>

          {/* Field Pole Drop Deviation */}
          <div className={`p-2.5 rounded-xl border ${liveStats.poleWarning ? 'bg-amber-950/70 border-amber-500/60 text-amber-200' : 'bg-slate-800/60 border-slate-700 text-slate-300'}`}>
            <div className="text-[10px] text-slate-400 flex items-center justify-between">
              <span>Sụt Áp Cực Từ:</span>
              <span className="font-mono text-[9px]">NETA ≤ 10.0%</span>
            </div>
            <div className="text-base font-bold font-mono mt-0.5 flex items-center gap-1.5">
              {liveStats.maxPoleDev.toFixed(1)}%
              {liveStats.poleWarning ? (
                <span className="px-1.5 py-0.2 rounded text-[10px] bg-amber-600 text-white font-sans font-bold">LỆCH</span>
              ) : (
                <span className="px-1.5 py-0.2 rounded text-[10px] bg-emerald-600 text-white font-sans font-bold">PASS</span>
              )}
            </div>
            <div className="text-[10px] opacity-80 mt-0.5 truncate">
              {liveStats.poleWarning ? 'Ngắn mạch vòng dây cực' : 'Cực từ đối xứng'}
            </div>
          </div>

          {/* Armature PI / DAR */}
          <div className={`p-2.5 rounded-xl border ${liveStats.piExceeded ? 'bg-rose-950/70 border-rose-500/60 text-rose-200' : 'bg-slate-800/60 border-slate-700 text-slate-300'}`}>
            <div className="text-[10px] text-slate-400 flex items-center justify-between">
              <span>{liveStats.isLarge ? 'Chỉ Số PI (>150kW):' : 'Hệ Số DAR (≤150kW):'}</span>
              <span className="font-mono text-[9px]">{liveStats.isLarge ? 'PI ≥ 2.0' : 'DAR ≥ 1.4'}</span>
            </div>
            <div className="text-base font-bold font-mono mt-0.5 flex items-center gap-1.5">
              {liveStats.isLarge ? (liveStats.pi ?? 'N/A') : (liveStats.dar ?? 'N/A')}
              {liveStats.piExceeded ? (
                <span className="px-1.5 py-0.2 rounded text-[10px] bg-rose-600 text-white font-sans font-bold">FAIL</span>
              ) : (
                <span className="px-1.5 py-0.2 rounded text-[10px] bg-emerald-600 text-white font-sans font-bold">PASS</span>
              )}
            </div>
            <div className="text-[10px] opacity-80 mt-0.5 truncate">
              {liveStats.piExceeded ? 'Cuộn dây ẩm/bẩn' : 'Cách điện khô ráo'}
            </div>
          </div>

          {/* Vibration Table 100.10 */}
          <div className={`p-2.5 rounded-xl border ${liveStats.vibExceeded ? 'bg-rose-950/70 border-rose-500/60 text-rose-200' : 'bg-slate-800/60 border-slate-700 text-slate-300'}`}>
            <div className="text-[10px] text-slate-400 flex items-center justify-between">
              <span>Độ Rung Không Tải:</span>
              <span className="font-mono text-[9px]">Table 100.10</span>
            </div>
            <div className="text-base font-bold font-mono mt-0.5 flex items-center gap-1.5">
              {liveStats.vibVal} <span className="text-[10px] font-sans font-normal">in/s</span>
              {liveStats.vibExceeded ? (
                <span className="px-1.5 py-0.2 rounded text-[10px] bg-rose-600 text-white font-sans font-bold">VƯỢT</span>
              ) : (
                <span className="px-1.5 py-0.2 rounded text-[10px] bg-emerald-600 text-white font-sans font-bold">PASS</span>
              )}
            </div>
            <div className="text-[10px] opacity-80 mt-0.5 truncate">
              Giới hạn: ≤ {liveStats.vibLimit} in/s pk
            </div>
          </div>

          {/* Bolted Resistance */}
          <div className={`p-2.5 rounded-xl border ${liveStats.boltWarning ? 'bg-amber-950/70 border-amber-500/60 text-amber-200' : 'bg-slate-800/60 border-slate-700 text-slate-300'}`}>
            <div className="text-[10px] text-slate-400 flex items-center justify-between">
              <span>Lệch Bu-lông DLRO:</span>
              <span className="font-mono text-[9px]">Table 100.12</span>
            </div>
            <div className="text-base font-bold font-mono mt-0.5 flex items-center gap-1.5">
              {liveStats.boltDev.toFixed(1)}%
              {liveStats.boltWarning ? (
                <span className="px-1.5 py-0.2 rounded text-[10px] bg-amber-600 text-white font-sans font-bold">LỆCH</span>
              ) : (
                <span className="px-1.5 py-0.2 rounded text-[10px] bg-emerald-600 text-white font-sans font-bold">PASS</span>
              )}
            </div>
            <div className="text-[10px] opacity-80 mt-0.5 truncate">
              Ngưỡng cho phép ≤ 50%
            </div>
          </div>
        </div>

        {savedSuccessMessage && (
          <div className="mt-4 p-3 bg-emerald-500/20 border border-emerald-400/40 rounded-xl text-emerald-300 text-xs flex items-center justify-between">
            <span className="flex items-center gap-2">
              <CheckCircle2 size={16} className="text-emerald-400" />
              {savedSuccessMessage}
            </span>
            {onNavigateToReports && (
              <button
                onClick={onNavigateToReports}
                className="underline hover:text-white font-semibold flex items-center gap-1"
              >
                Xem trong Kho Báo Cáo <ArrowRight size={12} />
              </button>
            )}
          </div>
        )}
      </div>

      {/* TABS NAVIGATION */}
      <div className="flex border-b border-slate-200 bg-white rounded-t-xl px-4 pt-2 gap-2 shadow-sm overflow-x-auto">
        <button
          onClick={() => setActiveTab('electrical')}
          className={`px-4 py-3 text-sm font-semibold border-b-2 transition-all flex items-center gap-2 whitespace-nowrap ${
            activeTab === 'electrical'
              ? 'border-rose-600 text-rose-600 bg-rose-50/50 rounded-t-lg'
              : 'border-transparent text-slate-600 hover:text-slate-900'
          }`}
        >
          <Zap size={16} />
          Phép Đo Điện & Thử Nghiệm DC (Mục 7.15.3.D)
        </button>
        <button
          onClick={() => setActiveTab('visual')}
          className={`px-4 py-3 text-sm font-semibold border-b-2 transition-all flex items-center gap-2 whitespace-nowrap ${
            activeTab === 'visual'
              ? 'border-rose-600 text-rose-600 bg-rose-50/50 rounded-t-lg'
              : 'border-transparent text-slate-600 hover:text-slate-900'
          }`}
        >
          <ShieldCheck size={16} />
          Kiểm Tra Thị Giác & Cổ Góp (Mục 7.15.3.A)
        </button>
        <button
          onClick={() => setActiveTab('site')}
          className={`px-4 py-3 text-sm font-semibold border-b-2 transition-all flex items-center gap-2 whitespace-nowrap ${
            activeTab === 'site'
              ? 'border-rose-600 text-rose-600 bg-rose-50/50 rounded-t-lg'
              : 'border-transparent text-slate-600 hover:text-slate-900'
          }`}
        >
          <Info size={16} />
          Thông Tin Máy Điện DC & FSE
        </button>
        <button
          onClick={() => setActiveTab('ai')}
          className={`px-4 py-3 text-sm font-semibold border-b-2 transition-all flex items-center gap-2 whitespace-nowrap ${
            activeTab === 'ai'
              ? 'border-purple-600 text-purple-600 bg-purple-50/50 rounded-t-lg'
              : 'border-transparent text-purple-700 hover:text-purple-900'
          }`}
        >
          <Sparkles size={16} className="text-purple-600" />
          Chẩn Đoán AI & Đánh Giá NETA ATS-2025
          {analysisResult && (
            <span className={`text-[10px] px-1.5 py-0.5 rounded font-bold ${
              analysisResult.overall_status === 'PASS'
                ? 'bg-emerald-100 text-emerald-800'
                : 'bg-rose-100 text-rose-800'
            }`}>
              {analysisResult.overall_status}
            </span>
          )}
        </button>
      </div>

      {/* TAB 1: PHÉP ĐO ĐIỆN VÀ THỬ NGHIỆM DC */}
      {activeTab === 'electrical' && (
        <div className="bg-white rounded-b-xl border border-slate-200 border-t-0 p-6 space-y-6 shadow-sm">
          {/* Card 1: Armature Bar-to-Bar Resistance (Adjacent Bars) */}
          <div className="p-5 rounded-xl border border-rose-200 bg-gradient-to-br from-rose-50/40 via-white to-orange-50/30 space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-rose-100 pb-3">
              <div className="flex items-center gap-2">
                <div className="p-2 rounded-lg bg-rose-600 text-white shadow-sm">
                  <Flame size={18} />
                </div>
                <div>
                  <h3 className="font-bold text-slate-800 text-base flex items-center gap-2">
                    1. Điện Trở Phiến-đến-Phiến Cổ Góp (Armature Bar-to-Bar Resistance)
                    <span className="text-xs font-normal text-rose-700 bg-rose-100 px-2 py-0.5 rounded border border-rose-300">
                      NETA Mục 7.15.3.D.4
                    </span>
                  </h3>
                  <p className="text-xs text-slate-500">
                    Đo điện trở giữa các phiến đồng kế cận trên cổ góp phần ứng bằng thiết bị micro-ohmmeter 4 dây. Quy chuẩn: Sai lệch không được vượt quá 5% (&gt; 5.0% BẮT BUỘC ĐÁNH GIÁ FAIL/INVESTIGATE).
                  </p>
                </div>
              </div>
              <div className={`px-3 py-1.5 rounded-lg text-xs font-bold font-mono flex items-center gap-1.5 self-start sm:self-auto ${
                liveStats.barExceeded ? 'bg-rose-600 text-white shadow' : 'bg-emerald-600 text-white shadow'
              }`}>
                {liveStats.barExceeded ? <AlertOctagon size={14} /> : <CheckCircle2 size={14} />}
                LỆCH CỰC ĐẠI: {data.electrical_tests.armature_bar_to_bar_resistance_micro_ohms.max_deviation_percent}% (TIÊU CHUẨN: ≤ 5.0%)
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Điện trở phiến Min đo được (µΩ):
                </label>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.armature_bar_to_bar_resistance_micro_ohms.min_bar_resistance_micro_ohms}
                  onChange={(e) => {
                    const val = parseFloat(e.target.value) || 0;
                    const maxVal = data.electrical_tests.armature_bar_to_bar_resistance_micro_ohms.max_bar_resistance_micro_ohms;
                    const dev = val > 0 ? parseFloat((((maxVal - val) / val) * 100).toFixed(1)) : 0;
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        armature_bar_to_bar_resistance_micro_ohms: {
                          ...data.electrical_tests.armature_bar_to_bar_resistance_micro_ohms,
                          min_bar_resistance_micro_ohms: val,
                          max_deviation_percent: dev
                        }
                      }
                    });
                  }}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-mono font-bold text-slate-900 focus:bg-white focus:border-rose-500 focus:ring-1 focus:ring-rose-500"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Điện trở phiến Max đo được (µΩ):
                </label>
                <input
                  type="number"
                  step="0.1"
                  value={data.electrical_tests.armature_bar_to_bar_resistance_micro_ohms.max_bar_resistance_micro_ohms}
                  onChange={(e) => {
                    const val = parseFloat(e.target.value) || 0;
                    const minVal = data.electrical_tests.armature_bar_to_bar_resistance_micro_ohms.min_bar_resistance_micro_ohms;
                    const dev = minVal > 0 ? parseFloat((((val - minVal) / minVal) * 100).toFixed(1)) : 0;
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        armature_bar_to_bar_resistance_micro_ohms: {
                          ...data.electrical_tests.armature_bar_to_bar_resistance_micro_ohms,
                          max_bar_resistance_micro_ohms: val,
                          max_deviation_percent: dev
                        }
                      }
                    });
                  }}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-mono font-bold text-slate-900 focus:bg-white focus:border-rose-500 focus:ring-1 focus:ring-rose-500"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Độ lệch phần trăm tính toán (%):
                </label>
                <div className={`px-3 py-2 rounded-lg font-mono font-bold text-sm border flex items-center justify-between ${
                  liveStats.barExceeded ? 'bg-rose-100 border-rose-300 text-rose-800' : 'bg-emerald-50 border-emerald-300 text-emerald-800'
                }`}>
                  <span>{data.electrical_tests.armature_bar_to_bar_resistance_micro_ohms.max_deviation_percent}%</span>
                  <span className="text-xs font-normal">
                    {liveStats.barExceeded ? 'VƯỢT NGƯỠNG (FAIL)' : 'ĐẠT (PASS)'}
                  </span>
                </div>
              </div>
            </div>

            {liveStats.barExceeded && (
              <div className="p-3 bg-rose-500/10 border border-rose-300 rounded-lg text-xs text-rose-900 space-y-1">
                <div className="font-bold flex items-center gap-1.5 text-rose-800">
                  <AlertTriangle size={15} />
                  CẢNH BÁO NGUY HIỂM: Sai lệch điện trở phiến cổ góp {data.electrical_tests.armature_bar_to_bar_resistance_micro_ohms.max_deviation_percent}% vượt ngưỡng 5.0%!
                </div>
                <p>
                  Khi động cơ quay ở tốc độ danh định, độ lệch này gây sốc điện áp giữa các phiến, sinh tia lửa điện hồ quang phá hủy chổi than, rỗ bề mặt đồng cổ góp và nguy cơ cháy nổ máy điện DC. Cần kiểm tra mối hàn lá phiến (riser joints) ngay lập tức!
                </p>
              </div>
            )}
          </div>

          {/* Card 2: Field Pole AC Voltage Drop */}
          <div className="p-5 rounded-xl border border-blue-200 bg-gradient-to-br from-blue-50/40 via-white to-indigo-50/30 space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-blue-100 pb-3">
              <div className="flex items-center gap-2">
                <div className="p-2 rounded-lg bg-blue-600 text-white shadow-sm">
                  <Activity size={18} />
                </div>
                <div>
                  <h3 className="font-bold text-slate-800 text-base flex items-center gap-2">
                    2. Thử Sụt Áp AC Từng Cực Từ Kích Từ (Pole-to-Pole AC Voltage Drop)
                    <span className="text-xs font-normal text-blue-700 bg-blue-100 px-2 py-0.5 rounded border border-blue-300">
                      NETA Mục 7.15.3.D.3
                    </span>
                  </h3>
                  <p className="text-xs text-slate-500">
                    Cấp điện áp xoay chiều AC phù hợp vào toàn bộ mạch kích từ. Độ lệch điện áp sụt trên từng cực từ không được lệch quá 10.0% so với giá trị trung bình toàn bộ các cực (phát hiện ngắn mạch vòng dây cực từ).
                  </p>
                </div>
              </div>
              <div className={`px-3 py-1.5 rounded-lg text-xs font-bold font-mono flex items-center gap-1.5 self-start sm:self-auto ${
                liveStats.poleWarning ? 'bg-amber-600 text-white shadow' : 'bg-emerald-600 text-white shadow'
              }`}>
                LỆCH CỰC ĐẠI: {liveStats.maxPoleDev.toFixed(1)}% (CHUẨN ≤ 10.0%)
              </div>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-5 gap-3">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Cực từ 1 (V):</label>
                <input
                  type="number"
                  step="0.05"
                  value={data.electrical_tests.field_pole_ac_voltage_drop_volts.pole_1}
                  onChange={(e) => {
                    const v = parseFloat(e.target.value) || 0;
                    const p = data.electrical_tests.field_pole_ac_voltage_drop_volts;
                    const avg = (v + p.pole_2 + p.pole_3 + p.pole_4) / 4;
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        field_pole_ac_voltage_drop_volts: {
                          ...p,
                          pole_1: v,
                          average_drop_volts: parseFloat(avg.toFixed(2))
                        }
                      }
                    });
                  }}
                  className="w-full px-3 py-1.5 bg-slate-50 border border-slate-300 rounded-lg text-sm font-mono font-bold text-slate-900"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Cực từ 2 (V):</label>
                <input
                  type="number"
                  step="0.05"
                  value={data.electrical_tests.field_pole_ac_voltage_drop_volts.pole_2}
                  onChange={(e) => {
                    const v = parseFloat(e.target.value) || 0;
                    const p = data.electrical_tests.field_pole_ac_voltage_drop_volts;
                    const avg = (p.pole_1 + v + p.pole_3 + p.pole_4) / 4;
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        field_pole_ac_voltage_drop_volts: {
                          ...p,
                          pole_2: v,
                          average_drop_volts: parseFloat(avg.toFixed(2))
                        }
                      }
                    });
                  }}
                  className="w-full px-3 py-1.5 bg-slate-50 border border-slate-300 rounded-lg text-sm font-mono font-bold text-slate-900"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Cực từ 3 (V):</label>
                <input
                  type="number"
                  step="0.05"
                  value={data.electrical_tests.field_pole_ac_voltage_drop_volts.pole_3}
                  onChange={(e) => {
                    const v = parseFloat(e.target.value) || 0;
                    const p = data.electrical_tests.field_pole_ac_voltage_drop_volts;
                    const avg = (p.pole_1 + p.pole_2 + v + p.pole_4) / 4;
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        field_pole_ac_voltage_drop_volts: {
                          ...p,
                          pole_3: v,
                          average_drop_volts: parseFloat(avg.toFixed(2))
                        }
                      }
                    });
                  }}
                  className="w-full px-3 py-1.5 bg-slate-50 border border-slate-300 rounded-lg text-sm font-mono font-bold text-slate-900"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Cực từ 4 (V):</label>
                <input
                  type="number"
                  step="0.05"
                  value={data.electrical_tests.field_pole_ac_voltage_drop_volts.pole_4}
                  onChange={(e) => {
                    const v = parseFloat(e.target.value) || 0;
                    const p = data.electrical_tests.field_pole_ac_voltage_drop_volts;
                    const avg = (p.pole_1 + p.pole_2 + p.pole_3 + v) / 4;
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        field_pole_ac_voltage_drop_volts: {
                          ...p,
                          pole_4: v,
                          average_drop_volts: parseFloat(avg.toFixed(2))
                        }
                      }
                    });
                  }}
                  className="w-full px-3 py-1.5 bg-slate-50 border border-slate-300 rounded-lg text-sm font-mono font-bold text-slate-900"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">Giá trị TB (V):</label>
                <div className="px-3 py-1.5 bg-slate-100 border border-slate-200 rounded-lg text-sm font-mono font-bold text-blue-900">
                  {data.electrical_tests.field_pole_ac_voltage_drop_volts.average_drop_volts} V
                </div>
              </div>
            </div>
          </div>

          {/* Card 3: Insulation Resistance (IEEE 43 / Table 100.11) */}
          <div className="p-5 rounded-xl border border-slate-200 bg-slate-50/50 space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-200 pb-3">
              <div className="flex items-center gap-2">
                <div className="p-2 rounded-lg bg-emerald-600 text-white shadow-sm">
                  <ShieldCheck size={18} />
                </div>
                <div>
                  <h3 className="font-bold text-slate-800 text-base flex items-center gap-2">
                    3. Điện Trở Cách Điện & Chỉ Số Phân Cực PI / DAR (IEEE Std 43)
                    <span className="text-xs font-normal text-emerald-800 bg-emerald-100 px-2 py-0.5 rounded border border-emerald-300">
                      Table 100.11 (40°C)
                    </span>
                  </h3>
                  <p className="text-xs text-slate-500">
                    Máy &gt; 150 kW (200 HP) đo 10 phút bắt buộc PI ≥ 2.0; máy ≤ 150 kW bắt buộc DAR ≥ 1.4. Điện trở tối thiểu 1 phút ≥ 100 MΩ.
                  </p>
                </div>
              </div>
              <span className="text-xs text-slate-500 font-mono">
                Loại máy: {liveStats.isLarge ? 'Công suất lớn (>150kW)' : 'Công suất nhỏ (≤150kW)'}
              </span>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Cách điện Phần ứng Armature 1-min (MΩ):
                </label>
                <input
                  type="number"
                  value={data.electrical_tests.insulation_resistance_40c_megohms.armature_1min_ir_1000v}
                  onChange={(e) => {
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_40c_megohms: {
                          ...data.electrical_tests.insulation_resistance_40c_megohms,
                          armature_1min_ir_1000v: parseFloat(e.target.value) || 0
                        }
                      }
                    });
                  }}
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-sm font-mono font-bold text-slate-900"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Chỉ số Phân cực Armature (PI = R10/R1):
                </label>
                <input
                  type="number"
                  step="0.01"
                  value={data.electrical_tests.insulation_resistance_40c_megohms.armature_pi_value || ''}
                  onChange={(e) => {
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_40c_megohms: {
                          ...data.electrical_tests.insulation_resistance_40c_megohms,
                          armature_pi_value: parseFloat(e.target.value) || undefined
                        }
                      }
                    });
                  }}
                  className={`w-full px-3 py-2 border rounded-lg text-sm font-mono font-bold ${
                    liveStats.piExceeded ? 'bg-rose-50 border-rose-300 text-rose-800' : 'bg-white border-slate-300 text-slate-900'
                  }`}
                  placeholder="≥ 2.0"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Cách điện Cuộn Kích từ Shunt 1-min (MΩ):
                </label>
                <input
                  type="number"
                  value={data.electrical_tests.insulation_resistance_40c_megohms.shunt_field_1min_ir_1000v || ''}
                  onChange={(e) => {
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_40c_megohms: {
                          ...data.electrical_tests.insulation_resistance_40c_megohms,
                          shunt_field_1min_ir_1000v: parseFloat(e.target.value) || undefined
                        }
                      }
                    });
                  }}
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-sm font-mono font-bold text-slate-900"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Chỉ số Phân cực Shunt Field (PI):
                </label>
                <input
                  type="number"
                  step="0.01"
                  value={data.electrical_tests.insulation_resistance_40c_megohms.shunt_field_pi_value || ''}
                  onChange={(e) => {
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        insulation_resistance_40c_megohms: {
                          ...data.electrical_tests.insulation_resistance_40c_megohms,
                          shunt_field_pi_value: parseFloat(e.target.value) || undefined
                        }
                      }
                    });
                  }}
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-sm font-mono font-bold text-slate-900"
                  placeholder="≥ 2.0"
                />
              </div>
            </div>
          </div>

          {/* Card 4: Bolted Connections, Surge, Running & Vibration */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {/* Bolted Resistance & Vibration */}
            <div className="p-4 rounded-xl border border-slate-200 bg-white space-y-3">
              <h4 className="font-bold text-slate-800 text-sm flex items-center justify-between border-b border-slate-100 pb-2">
                <span className="flex items-center gap-1.5">
                  <Gauge size={16} className="text-slate-600" />
                  Mối Nối Bu-lông & Độ Rung (Table 100.10 & 100.12)
                </span>
                <span className="text-xs font-mono text-slate-500">DLRO & Accel</span>
              </h4>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Điện trở mối nối bu-lông (µΩ, cách nhau dấu phẩy):
                </label>
                <input
                  type="text"
                  value={data.electrical_tests.bolted_resistance_micro_ohms.join(', ')}
                  onChange={(e) => {
                    const arr = e.target.value.split(',').map(s => parseFloat(s.trim())).filter(n => !isNaN(n));
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        bolted_resistance_micro_ohms: arr
                      }
                    });
                  }}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-mono font-bold text-slate-900"
                />
                <span className="text-[11px] text-slate-500">
                  Lệch hiện tại: {liveStats.boltDev.toFixed(1)}% (Ngưỡng cho phép ≤ 50% min)
                </span>
              </div>

              <div className="grid grid-cols-2 gap-2 pt-1">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Độ rung Vận tốc Peak (in/s pk):
                  </label>
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.unloaded_vibration_table_100_10.velocity_in_sec_pk}
                    onChange={(e) => {
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          unloaded_vibration_table_100_10: {
                            ...data.electrical_tests.unloaded_vibration_table_100_10,
                            velocity_in_sec_pk: parseFloat(e.target.value) || 0
                          }
                        }
                      });
                    }}
                    className={`w-full px-3 py-1.5 border rounded-lg text-sm font-mono font-bold ${
                      liveStats.vibExceeded ? 'bg-rose-50 border-rose-300 text-rose-800' : 'bg-slate-50 border-slate-300 text-slate-900'
                    }`}
                  />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Giới hạn Table 100.10:
                  </label>
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.unloaded_vibration_table_100_10.nema_limit_in_sec_pk}
                    onChange={(e) => {
                      setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          unloaded_vibration_table_100_10: {
                            ...data.electrical_tests.unloaded_vibration_table_100_10,
                            nema_limit_in_sec_pk: parseFloat(e.target.value) || 0.15
                          }
                        }
                      });
                    }}
                    className="w-full px-3 py-1.5 bg-slate-100 border border-slate-200 rounded-lg text-sm font-mono font-bold text-slate-700"
                  />
                </div>
              </div>
            </div>

            {/* Surge & Running Current */}
            <div className="p-4 rounded-xl border border-slate-200 bg-white space-y-3">
              <h4 className="font-bold text-slate-800 text-sm flex items-center justify-between border-b border-slate-100 pb-2">
                <span className="flex items-center gap-1.5">
                  <Activity size={16} className="text-slate-600" />
                  Thử Xung & Dòng Vận Hành (NETA Sec 7.15.3.D.5-6)
                </span>
                <span className="text-xs font-mono text-slate-500">Waveform & Excitation</span>
              </h4>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Thử Xung So Sánh (Surge Comparison Test):
                </label>
                <input
                  type="text"
                  value={data.electrical_tests.surge_comparison_test}
                  onChange={(e) => {
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        surge_comparison_test: e.target.value
                      }
                    });
                  }}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm text-slate-900"
                  placeholder="Pass (Waveforms nested completely)"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Dòng vận hành & Thông số kích từ (Running & Field Current):
                </label>
                <input
                  type="text"
                  value={data.electrical_tests.running_current_and_field_check}
                  onChange={(e) => {
                    setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        running_current_and_field_check: e.target.value
                      }
                    });
                  }}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm text-slate-900"
                  placeholder="Pass (No-load Armature 14.2A, Field 4.8A at 220V)"
                />
              </div>
            </div>
          </div>

          {/* Action Button */}
          <div className="pt-2 flex justify-end">
            <button
              onClick={handleAnalyze}
              disabled={isAnalyzing}
              className="px-6 py-3 bg-gradient-to-r from-rose-600 via-rose-700 to-slate-900 hover:from-rose-500 hover:to-slate-800 text-white font-bold rounded-xl shadow-lg transition-all flex items-center gap-2"
            >
              {isAnalyzing ? (
                <>
                  <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                  Đang phân tích chuyên gia AI...
                </>
              ) : (
                <>
                  <Sparkles size={18} className="text-yellow-300" />
                  Phân Tích AI Chuyên Sâu (NETA ATS-2025 Sec 7.15.3)
                </>
              )}
            </button>
          </div>
        </div>
      )}

      {/* TAB 2: KIỂM TRA THỊ GIÁC & CƠ KHÍ */}
      {activeTab === 'visual' && (
        <div className="bg-white rounded-b-xl border border-slate-200 border-t-0 p-6 space-y-5 shadow-sm">
          <div className="border-b border-slate-200 pb-3">
            <h3 className="font-bold text-slate-800 text-base flex items-center gap-2">
              <ShieldCheck className="text-rose-600" size={20} />
              Kiểm Tra Thị Giác & Cơ Khí (NETA ATS-2025 Mục 7.15.3.A)
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Đảm bảo máy điện một chiều được lắp đặt chính xác, bề mặt cổ góp nhẵn phẳng, chổi than tiếp xúc đúng góc, neo móng và siết lực bu-lông đạt chuẩn.
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="p-4 rounded-xl border border-slate-200 bg-slate-50/50 space-y-3">
              <div className="flex items-center justify-between">
                <label className="text-xs font-semibold text-slate-700">
                  Đối soát nhãn mác (Nameplate Match):
                </label>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.nameplate_match}
                  onChange={(e) => {
                    setData({
                      ...data,
                      visual_inspection: {
                        ...data.visual_inspection,
                        nameplate_match: e.target.checked
                      }
                    });
                  }}
                  className="w-4 h-4 text-rose-600 rounded focus:ring-rose-500"
                />
              </div>
              <span className="text-[11px] text-slate-500 block">
                Khớp 100% bản vẽ thiết kế (kW/HP, điện áp/dòng phần ứng Armature, điện áp/dòng kích từ Field, RPM).
              </span>

              <div className="pt-2">
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Tình trạng cơ khí & Vệ sinh (Physical Condition & Cleanliness):
                </label>
                <input
                  type="text"
                  value={data.visual_inspection.physical_condition_cleanliness}
                  onChange={(e) => {
                    setData({
                      ...data,
                      visual_inspection: {
                        ...data.visual_inspection,
                        physical_condition_cleanliness: e.target.value
                      }
                    });
                  }}
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-sm text-slate-900"
                />
              </div>

              <div className="pt-1">
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Cổ góp & Giá đỡ chổi than (Commutator & Brush Rigging):
                </label>
                <input
                  type="text"
                  value={data.visual_inspection.commutator_and_brush_rigging}
                  onChange={(e) => {
                    setData({
                      ...data,
                      visual_inspection: {
                        ...data.visual_inspection,
                        commutator_and_brush_rigging: e.target.value
                      }
                    });
                  }}
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-sm text-slate-900"
                />
              </div>
            </div>

            <div className="p-4 rounded-xl border border-slate-200 bg-slate-50/50 space-y-3">
              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Máy phát tốc độ (Tachometer Generator Check):
                </label>
                <input
                  type="text"
                  value={data.visual_inspection.tachometer_generator_check}
                  onChange={(e) => {
                    setData({
                      ...data,
                      visual_inspection: {
                        ...data.visual_inspection,
                        tachometer_generator_check: e.target.value
                      }
                    });
                  }}
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-sm text-slate-900"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Neo móng, đồng trục & tiếp địa (Anchorage, Alignment & Grounding):
                </label>
                <input
                  type="text"
                  value={data.visual_inspection.anchorage_alignment_grounding}
                  onChange={(e) => {
                    setData({
                      ...data,
                      visual_inspection: {
                        ...data.visual_inspection,
                        anchorage_alignment_grounding: e.target.value
                      }
                    });
                  }}
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-sm text-slate-900"
                />
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Lực siết bu-lông (Bolt Torque per Table 100.12):
                </label>
                <input
                  type="text"
                  value={data.visual_inspection.bolt_torque_check}
                  onChange={(e) => {
                    setData({
                      ...data,
                      visual_inspection: {
                        ...data.visual_inspection,
                        bolt_torque_check: e.target.value
                      }
                    });
                  }}
                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-sm text-slate-900"
                />
              </div>
            </div>
          </div>
        </div>
      )}

      {/* TAB 3: THÔNG TIN MÁY ĐIỆN DC & FSE */}
      {activeTab === 'site' && (
        <div className="bg-white rounded-b-xl border border-slate-200 border-t-0 p-6 space-y-5 shadow-sm">
          <div className="border-b border-slate-200 pb-3">
            <h3 className="font-bold text-slate-800 text-base flex items-center gap-2">
              <Info className="text-rose-600" size={20} />
              Thông Tin Nhãn Máy & Hiện Trường (Nameplate & Site Metadata)
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Thông số định mức động cơ/máy phát DC và thông tin kỹ sư kiểm định hiện trường.
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">Tên dự án / Nhà máy:</label>
              <input
                type="text"
                value={data.site_info.project_name}
                onChange={(e) => setData({ ...data, site_info: { ...data.site_info, project_name: e.target.value } })}
                className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold text-slate-900"
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">Mã thiết bị (Motor Tag):</label>
              <input
                type="text"
                value={data.site_info.motor_tag}
                onChange={(e) => setData({ ...data, site_info: { ...data.site_info, motor_tag: e.target.value } })}
                className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-mono font-bold text-slate-900"
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">Kỹ sư kiểm định FSE:</label>
              <FseEngineerDropdown
                value={data.site_info.fse_name}
                onChange={(val) => setData({ ...data, site_info: { ...data.site_info, fse_name: val } })}
              />
            </div>

            <div className="md:col-span-2">
              <label className="block text-xs font-semibold text-slate-700 mb-1">Thông số nhãn máy (Rating Specification):</label>
              <input
                type="text"
                value={data.site_info.motor_rating}
                onChange={(e) => setData({ ...data, site_info: { ...data.site_info, motor_rating: e.target.value } })}
                className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold text-slate-900"
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">Ngày kiểm định:</label>
              <input
                type="date"
                value={data.site_info.test_date}
                onChange={(e) => setData({ ...data, site_info: { ...data.site_info, test_date: e.target.value } })}
                className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-mono text-slate-900"
              />
            </div>
          </div>
        </div>
      )}

      {/* TAB 4: CHẨN ĐOÁN AI & ĐÁNH GIÁ NETA ATS-2025 */}
      {activeTab === 'ai' && (
        <div className="bg-white rounded-b-xl border border-slate-200 border-t-0 p-6 space-y-6 shadow-sm">
          {analysisResult ? (
            <div className="space-y-6">
              {/* Header Box */}
              <div className={`p-4 rounded-xl border flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${
                analysisResult.overall_status === 'PASS'
                  ? 'bg-emerald-50 border-emerald-300 text-emerald-950'
                  : analysisResult.overall_status === 'INVESTIGATE'
                  ? 'bg-amber-50 border-amber-300 text-amber-950'
                  : 'bg-rose-50 border-rose-300 text-rose-950'
              }`}>
                <div className="flex items-center gap-3">
                  <div className={`p-2.5 rounded-lg ${
                    analysisResult.overall_status === 'PASS'
                      ? 'bg-emerald-600 text-white'
                      : analysisResult.overall_status === 'INVESTIGATE'
                      ? 'bg-amber-600 text-white'
                      : 'bg-rose-600 text-white'
                  }`}>
                    {analysisResult.overall_status === 'PASS' ? <CheckCircle2 size={24} /> : <AlertTriangle size={24} />}
                  </div>
                  <div>
                    <div className="text-xs uppercase tracking-wider font-bold opacity-75">
                      KẾT QUẢ ĐÁNH GIÁ TỔNG QUAN NETA ATS-2025 SEC 7.15.3
                    </div>
                    <div className="text-xl font-bold font-mono">
                      {analysisResult.status_title}
                    </div>
                  </div>
                </div>

                <div className="flex items-center gap-2 self-end sm:self-auto">
                  <button
                    onClick={handleSaveToCmms}
                    className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold rounded-lg shadow transition-colors flex items-center gap-1.5"
                  >
                    <Save size={14} />
                    Lưu Báo Cáo CMMS
                  </button>
                  <button
                    onClick={() => setShowPrintModal(true)}
                    className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-white text-xs font-semibold rounded-lg shadow transition-colors flex items-center gap-1.5"
                  >
                    <Printer size={14} />
                    In Biên Bản
                  </button>
                </div>
              </div>

              {/* Render Markdown Report */}
              <div className="p-6 bg-slate-900 text-slate-100 rounded-xl border border-slate-800 font-sans text-sm leading-relaxed overflow-x-auto shadow-inner">
                <div className="prose prose-invert max-w-none space-y-4">
                  <pre className="whitespace-pre-wrap font-sans text-slate-200 leading-relaxed">
                    {analysisResult.markdown_report}
                  </pre>
                </div>
              </div>
            </div>
          ) : (
            <div className="text-center py-12 space-y-4">
              <div className="inline-flex p-4 rounded-full bg-rose-50 text-rose-600 border border-rose-200">
                <Sparkles size={36} className="animate-bounce" />
              </div>
              <h3 className="text-lg font-bold text-slate-800">
                Chưa có dữ liệu phân tích từ Trợ Lý AI Chuyên Gia
              </h3>
              <p className="text-xs text-slate-500 max-w-md mx-auto">
                Nhấn vào nút bên dưới để tiến hành chạy thuật toán đối soát 100% theo tiêu chuẩn ANSI/NETA ATS-2025 Mục 7.15.3, IEEE Std 43 và Table 100.10, 100.11, 100.12.
              </p>
              <button
                onClick={handleAnalyze}
                disabled={isAnalyzing}
                className="px-6 py-2.5 bg-rose-600 hover:bg-rose-500 text-white text-sm font-bold rounded-lg shadow-md transition-all inline-flex items-center gap-2"
              >
                {isAnalyzing ? (
                  <>
                    <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                    Đang phân tích...
                  </>
                ) : (
                  <>
                    <Sparkles size={16} />
                    Chạy Phân Tích Ngay
                  </>
                )}
              </button>
            </div>
          )}
        </div>
      )}

      {/* JSON MODAL */}
      {showJsonModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-3xl overflow-hidden shadow-2xl flex flex-col max-h-[85vh]">
            <div className="p-4 border-b border-slate-800 flex items-center justify-between text-white">
              <h3 className="text-sm font-bold flex items-center gap-2">
                <Layers size={16} className="text-rose-400" />
                Dữ Liệu JSON Kiểm Định Máy Điện DC (NETA ATS-2025 Mục 7.15.3)
              </h3>
              <button
                onClick={() => setShowJsonModal(false)}
                className="text-slate-400 hover:text-white text-sm p-1"
              >
                ✕
              </button>
            </div>

            <div className="p-4 flex-1 overflow-hidden flex flex-col">
              <textarea
                value={jsonText}
                onChange={(e) => setJsonText(e.target.value)}
                className="w-full flex-1 bg-slate-950 text-emerald-400 font-mono text-xs p-3 rounded-lg border border-slate-800 outline-none resize-none leading-relaxed"
                rows={16}
              />
            </div>

            <div className="p-4 border-t border-slate-800 flex items-center justify-between bg-slate-950/50">
              <button
                onClick={handleCopyJson}
                className="px-3.5 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-lg text-xs font-medium flex items-center gap-1.5 transition-colors"
              >
                {copied ? <Check size={14} className="text-emerald-400" /> : <Copy size={14} />}
                {copied ? 'Đã sao chép!' : 'Sao chép JSON'}
              </button>

              <div className="flex items-center gap-2">
                <button
                  onClick={() => setShowJsonModal(false)}
                  className="px-3.5 py-2 text-slate-400 hover:text-slate-200 text-xs font-medium"
                >
                  Hủy bỏ
                </button>
                <button
                  onClick={handleApplyJson}
                  className="px-4 py-2 bg-rose-600 hover:bg-rose-500 text-white rounded-lg text-xs font-bold transition-colors"
                >
                  Áp Dụng Dữ Liệu
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* PROMPT GOOGLE AI STUDIO MODAL */}
      {showPromptModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4">
          <div className="bg-slate-900 border border-purple-800/60 rounded-2xl w-full max-w-3xl overflow-hidden shadow-2xl flex flex-col max-h-[85vh]">
            <div className="p-4 border-b border-slate-800 flex items-center justify-between text-white bg-gradient-to-r from-purple-950/80 to-slate-900">
              <h3 className="text-sm font-bold flex items-center gap-2 text-purple-200">
                <Sparkles size={16} className="text-purple-400" />
                Prompt Mô Phỏng Trên Google AI Studio (User Request & JSON)
              </h3>
              <button
                onClick={() => setShowPromptModal(false)}
                className="text-slate-400 hover:text-white text-sm p-1"
              >
                ✕
              </button>
            </div>

            <div className="p-4 flex-1 overflow-hidden flex flex-col">
              <p className="text-xs text-slate-400 mb-2">
                Sao chép toàn bộ prompt này để paste trực tiếp vào giao diện Google AI Studio hoặc API:
              </p>
              <textarea
                readOnly
                value={googleStudioPrompt}
                className="w-full flex-1 bg-slate-950 text-purple-300 font-mono text-xs p-3 rounded-lg border border-purple-950 outline-none resize-none leading-relaxed"
                rows={16}
              />
            </div>

            <div className="p-4 border-t border-slate-800 flex items-center justify-between bg-slate-950/50">
              <span className="text-[11px] text-slate-500">
                ANSI/NETA ATS-2025 Section 7.15.3 (DC Motors & Generators)
              </span>

              <div className="flex items-center gap-2">
                <button
                  onClick={() => setShowPromptModal(false)}
                  className="px-3.5 py-2 text-slate-400 hover:text-slate-200 text-xs font-medium"
                >
                  Đóng
                </button>
                <button
                  onClick={handleCopyPrompt}
                  className="px-4 py-2 bg-purple-600 hover:bg-purple-500 text-white rounded-lg text-xs font-bold transition-colors flex items-center gap-1.5"
                >
                  {promptCopied ? <Check size={14} className="text-emerald-300" /> : <Copy size={14} />}
                  {promptCopied ? 'Đã sao chép Prompt!' : 'Sao Chép Prompt'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* PRINT REPORT MODAL */}
      {showPrintModal && (
        <TevReportPrintModal
          isOpen={showPrintModal}
          onClose={() => setShowPrintModal(false)}
          equipmentCode={data.site_info.motor_tag}
          equipmentName={data.site_info.motor_rating}
          siteName={data.site_info.project_name}
          inspectorName={data.site_info.fse_name}
          inspectionDate={data.site_info.test_date}
          healthStatus={liveStats.isFail ? 'critical' : liveStats.isInvestigate ? 'warning' : 'healthy'}
          overallScore={liveStats.isFail ? 42 : liveStats.isInvestigate ? 70 : 96}
          netaData={data}
          netaAnalysis={analysisResult || liveStats.analysis}
          testStandard="ANSI/NETA ATS-2025 Mục 7.15.3 & IEEE Std 43"
        />
      )}
    </div>
  );
};
