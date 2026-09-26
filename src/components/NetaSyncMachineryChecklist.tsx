import React, { useState, useMemo } from 'react';
import {
  RotateCcw,
  Zap,
  CheckCircle2,
  AlertTriangle,
  FileText,
  Printer,
  Sparkles,
  Save,
  Compass,
  ShieldCheck,
  Activity,
  Check,
  Eye,
  Camera,
  Upload,
  Gauge,
  Sliders,
  Wind
} from 'lucide-react';
import {
  NetaSyncMachineryInputPayload,
  NetaSyncMachineryEvaluationResult,
  performDeterministicSyncMachineryAnalysis
} from '../../server/netaSyncMachineryAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';
import { NameplateScannerModal } from './NameplateScannerModal';
import { ExtractedNameplateData } from '../../server/nameplateOcrAnalyzer';
import { FseEngineerDropdown } from './FseEngineerDropdown';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

export const SAMPLE_SYNC_MACHINERY_FSE_DATA: NetaSyncMachineryInputPayload = {
  site_info: {
    project_name: "Nha May Nhiet Dien TEV Generator Hall",
    generator_tag: "GEN-SYNC-20MW-01",
    generator_rating: "20MW 11kV 1500RPM Brushless Synchronous Generator",
    fse_name: "Võ Minh Tú (sgm1707@gmail.com)",
    test_date: "2026-09-23",
    manufacturer: "Siemens / GE / Andritz",
    serial_number: "SN-TEV-GEN-2026-018",
    rated_voltage_kv: 11,
    rated_speed_rpm: 1500,
    excitation_type: "Brushless"
  },
  visual_inspection: {
    nameplate_match: true,
    physical_cleanliness: "Pass",
    air_gap_uniformity: "Pass",
    brushless_exciter_condition: "Good",
    bolt_torque_check: "Pass",
    shaft_alignment_check: "Pass"
  },
  electrical_tests: {
    stator_ir_40c_megohms: {
      stator_1min_ir_2500v: 1800,
      stator_10min_ir_2500v: 5400,
      calculated_pi_value: 3.0
    },
    main_field_rotor_ir_1000v_megohms: 250,
    exciter_stator_rotor_ir_megohms: 320,
    pole_to_pole_ac_voltage_drop_volts: {
      pole_1: 24.2,
      pole_2: 24.5,
      pole_3: 18.1,
      pole_4: 24.4,
      average_voltage_drop_per_pole: 22.8
    },
    rotating_diodes_check: "Pass (All 6 diodes healthy)",
    stator_phase_resistance_ohms: {
      phase_ab: 0.082,
      phase_bc: 0.083,
      phase_ca: 0.082
    },
    uncoupled_vibration_in_sec_pk: 0.05,
    power_factor_relay_trip_test: "Pass"
  },
  previous_test_data: {
    last_test_date: "2025-09-20",
    last_pole_3_voltage_drop_volts: 24.1,
    last_stator_pi_value: 3.1,
    last_vibration_in_sec_pk: 0.04
  }
};

export const NetaSyncMachineryChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaSyncMachineryInputPayload>(SAMPLE_SYNC_MACHINERY_FSE_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaSyncMachineryEvaluationResult | null>(null);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showCameraModal, setShowCameraModal] = useState(false);
  const [jsonInput, setJsonInput] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live statistical calculations
  const liveStats = useMemo(() => {
    // 1. Stator IR & PI (Table 100.11)
    const stator = data.electrical_tests.stator_ir_40c_megohms;
    const computedPi = stator.stator_1min_ir_2500v > 0
      ? Number((stator.stator_10min_ir_2500v / stator.stator_1min_ir_2500v).toFixed(2))
      : stator.calculated_pi_value;
    const piPass = computedPi >= 2.0;
    const statorIrPass = stator.stator_1min_ir_2500v >= 100;

    // 2. Rotor Main Field IR
    const rotorIrPass = data.electrical_tests.main_field_rotor_ir_1000v_megohms >= 100;

    // 3. Pole-to-Pole AC Voltage Drop Test
    const drops = data.electrical_tests.pole_to_pole_ac_voltage_drop_volts;
    const poleList = [
      { id: 1, val: drops.pole_1 },
      { id: 2, val: drops.pole_2 },
      { id: 3, val: drops.pole_3 },
      { id: 4, val: drops.pole_4 },
    ];
    if (drops.pole_5 !== undefined) poleList.push({ id: 5, val: drops.pole_5 });
    if (drops.pole_6 !== undefined) poleList.push({ id: 6, val: drops.pole_6 });

    const avgDrop = poleList.reduce((acc, p) => acc + p.val, 0) / poleList.length;

    let hasCriticalPoleFail = false;
    let worstPole = 1;
    let worstDev = 0;

    const analyzedPoles = poleList.map((p) => {
      const dev = avgDrop > 0 ? ((p.val - avgDrop) / avgDrop) * 100 : 0;
      if (Math.abs(dev) > Math.abs(worstDev)) {
        worstDev = dev;
        worstPole = p.id;
      }
      if (Math.abs(dev) > 10.0) {
        hasCriticalPoleFail = true;
      }
      return {
        ...p,
        dev,
        pass: Math.abs(dev) <= 10.0
      };
    });

    // 4. Stator phase resistance balance
    const res = data.electrical_tests.stator_phase_resistance_ohms;
    const resArr = [res.phase_ab, res.phase_bc, res.phase_ca];
    const minRes = Math.min(...resArr);
    const maxRes = Math.max(...resArr);
    const resDev = minRes > 0 ? ((maxRes - minRes) / minRes) * 100 : 0;
    const resPass = resDev <= 5.0;

    // 5. Uncoupled Vibration (Table 100.10)
    const vib = data.electrical_tests.uncoupled_vibration_in_sec_pk;
    const vibPass = vib <= 0.15;

    // 6. Diodes check
    const diodesPass = data.electrical_tests.rotating_diodes_check.toLowerCase().includes('pass') ||
      data.electrical_tests.rotating_diodes_check.toLowerCase().includes('healthy');

    const overallSafe = !hasCriticalPoleFail && piPass && statorIrPass && rotorIrPass && resPass && vibPass && diodesPass;

    return {
      computedPi,
      piPass,
      statorIrPass,
      rotorIrPass,
      avgDrop,
      analyzedPoles,
      hasCriticalPoleFail,
      worstPole,
      worstDev,
      resDev,
      resPass,
      vibPass,
      diodesPass,
      overallSafe
    };
  }, [data]);

  const handleApplyNameplateData = (extracted: ExtractedNameplateData) => {
    setData((prev) => ({
      ...prev,
      site_info: {
        ...prev.site_info,
        generator_tag: extracted.equipment_tag || prev.site_info.generator_tag,
        generator_rating: extracted.equipment_type || extracted.equipment_name || prev.site_info.generator_rating,
        manufacturer: extracted.manufacturer || prev.site_info.manufacturer,
        serial_number: extracted.serial_number || prev.site_info.serial_number,
        rated_voltage_kv: extracted.primary_voltage_kv || prev.site_info.rated_voltage_kv
      },
      visual_inspection: {
        ...prev.visual_inspection,
        nameplate_match: true
      }
    }));
    setSavedSuccessMessage(`Đã quét nhãn thành công: ${extracted.equipment_tag || 'Máy điện đồng bộ'} - Tự động cập nhật vào Checklist!`);
    setTimeout(() => setSavedSuccessMessage(null), 5000);
  };

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    setAnalysisResult(null);
    try {
      const res = await fetch('/api/field-service/neta-sync-machinery-analysis', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });
      const json = await res.json();
      if (json.success && json.data) {
        setAnalysisResult(json.data);
      } else {
        const det = performDeterministicSyncMachineryAnalysis(data);
        setAnalysisResult(det);
      }
    } catch {
      const det = performDeterministicSyncMachineryAnalysis(data);
      setAnalysisResult(det);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleSaveToCmms = () => {
    const reportPayload = {
      type: 'NETA_ATS_2025_SYNC_MACHINERY',
      data: data,
      analysisResult: analysisResult,
      equipmentId: data.site_info.generator_tag,
      date: data.site_info.test_date,
      title: `Biên bản Máy Điện Đồng Bộ NETA ATS-2025: ${data.site_info.generator_tag}`,
      standard: 'ANSI/NETA ATS-2025 Section 7.15.2',
      status: liveStats.overallSafe ? 'PASS' : 'FAIL',
      timestamp: new Date().toISOString()
    };

    if (onSaveReport) {
      onSaveReport(reportPayload);
      setSavedSuccessMessage(`Đã lưu biên bản kiểm định máy điện đồng bộ ${data.site_info.generator_tag} vào kho báo cáo thành công!`);
      setTimeout(() => setSavedSuccessMessage(null), 5000);
    }
  };

  const electricalTableForPrint = [
    {
      item: 'Cách điện Stator 2500V & Chỉ số PI (Stator IR & PI @ 40°C)',
      measured: `IR 1min: ${data.electrical_tests.stator_ir_40c_megohms.stator_1min_ir_2500v} MΩ, IR 10min: ${data.electrical_tests.stator_ir_40c_megohms.stator_10min_ir_2500v} MΩ, PI = ${liveStats.computedPi}`,
      standard: 'IR 1min ≥ 100 MΩ; PI (10min/1min) ≥ 2.0 (Table 100.11)',
      reference: 'NETA ATS-2025 Sec 7.15.2.B.1 & Table 100.11',
      deviation: `PI = ${liveStats.computedPi}`,
      status: liveStats.piPass && liveStats.statorIrPass ? 'PASS' : 'FAIL'
    },
    {
      item: 'Cách điện Cuộn Kích từ Rô-to 1000V DC (Main Field Rotor IR)',
      measured: `${data.electrical_tests.main_field_rotor_ir_1000v_megohms} MΩ`,
      standard: 'Tối thiểu ≥ 100 MΩ',
      reference: 'NETA ATS-2025 Sec 7.15.2.B.2 & Table 100.11',
      deviation: '0%',
      status: liveStats.rotorIrPass ? 'PASS' : 'FAIL'
    },
    {
      item: 'Thử Sụt áp AC từng Cực Rô-to (Pole-to-Pole AC Voltage Drop)',
      measured: `TB: ${liveStats.avgDrop.toFixed(1)}V | ${liveStats.analyzedPoles.map(p => `P${p.id}: ${p.val}V (${p.dev > 0 ? '+' : ''}${p.dev.toFixed(1)}%)`).join(', ')}`,
      standard: 'Độ lệch sụt áp trên từng cực KHÔNG ĐƯỢC vượt quá 10% (≤ ±10%) so với trung bình',
      reference: 'NETA ATS-2025 Sec 7.15.2.B.3',
      deviation: `Pole ${liveStats.worstPole}: ${liveStats.worstDev > 0 ? '+' : ''}${liveStats.worstDev.toFixed(1)}%`,
      status: !liveStats.hasCriticalPoleFail ? 'PASS' : 'FAIL'
    },
    {
      item: 'Cụm Điốt Quay Kích Từ (Rotating Diodes Check)',
      measured: data.electrical_tests.rotating_diodes_check,
      standard: 'Tất cả các điốt mở thuận và khóa ngược tốt, không bị thủng/đứt',
      reference: 'NETA ATS-2025 Sec 7.15.2.B.4',
      deviation: '0%',
      status: liveStats.diodesPass ? 'PASS' : 'FAIL'
    },
    {
      item: 'Điện trở cuộn dây Stator giữa các pha (Stator Phase Resistance)',
      measured: `AB: ${data.electrical_tests.stator_phase_resistance_ohms.phase_ab} Ω, BC: ${data.electrical_tests.stator_phase_resistance_ohms.phase_bc} Ω, CA: ${data.electrical_tests.stator_phase_resistance_ohms.phase_ca} Ω (Lệch max: ${liveStats.resDev.toFixed(2)}%)`,
      standard: 'Độ lệch điện trở giữa các pha ≤ 5.0%',
      reference: 'NETA ATS-2025 Sec 7.15.2.B.5',
      deviation: `${liveStats.resDev.toFixed(2)}%`,
      status: liveStats.resPass ? 'PASS' : 'INVESTIGATE'
    },
    {
      item: 'Độ rung không tải tháo khớp (Uncoupled Vibration Table 100.10)',
      measured: `${data.electrical_tests.uncoupled_vibration_in_sec_pk} in/sec pk`,
      standard: '≤ 0.15 in/sec pk theo Table 100.10 (máy 1500 - 1800 RPM)',
      reference: 'NETA ATS-2025 Sec 7.15.2.B.6 & Table 100.10',
      deviation: '0%',
      status: liveStats.vibPass ? 'PASS' : 'FAIL'
    }
  ];

  const visualChecksForPrint = [
    { label: 'Đối soát nhãn Stator, Rotor cực lồi & Exciter (Nameplate Match)', value: data.visual_inspection.nameplate_match ? 'ĐẠT' : 'KHÔNG ĐẠT', pass: data.visual_inspection.nameplate_match },
    { label: 'Vệ sinh Stator, Rotor & Cuộn kích từ (Cleanliness)', value: data.visual_inspection.physical_cleanliness, pass: data.visual_inspection.physical_cleanliness.toLowerCase().includes('pass') },
    { label: 'Độ đồng đều khe hở không khí giữa các cực (Air-Gap Uniformity)', value: data.visual_inspection.air_gap_uniformity, pass: data.visual_inspection.air_gap_uniformity.toLowerCase().includes('pass') },
    { label: 'Tình trạng bộ kích từ không chổi than (Brushless Exciter)', value: data.visual_inspection.brushless_exciter_condition, pass: data.visual_inspection.brushless_exciter_condition.toLowerCase().includes('good') || data.visual_inspection.brushless_exciter_condition.toLowerCase().includes('pass') },
    { label: 'Kiểm tra lực siết bu-lông chân máy & khớp nối (Bolt Torque Table 100.12)', value: data.visual_inspection.bolt_torque_check, pass: data.visual_inspection.bolt_torque_check.toLowerCase().includes('pass') },
    { label: 'Căn chỉnh đồng trục máy phát - tuabin (Shaft Alignment)', value: data.visual_inspection.shaft_alignment_check || 'Pass', pass: true }
  ];

  return (
    <div className="space-y-6">
      {/* Top Banner & Header */}
      <div className="bg-gradient-to-r from-slate-950 via-indigo-950 to-blue-950 text-white p-5 sm:p-6 rounded-2xl shadow-lg border border-slate-700/50">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2 mb-1.5">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-400/20 text-amber-300 border border-amber-400/30 uppercase tracking-wide">
                ANSI/NETA ATS-2025 Section 7.15.2
              </span>
              <span className="text-xs text-slate-300">Rotating Machinery • Synchronous Motors and Generators</span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white flex items-center gap-2.5">
              <RotateCcw className="text-blue-400" size={26} />
              Checklist Động Cơ & Máy Phát Điện Đồng Bộ
            </h1>
            <p className="text-xs sm:text-sm text-slate-300 mt-1 max-w-2xl">
              Thẩm định chuyên sâu theo NETA ATS-2025: Thử sụt áp AC các cực rô-to cực lồi (≤10% dev), cách điện Stator & chỉ số PI (Table 100.11), kiểm tra điốt quay và đo độ rung máy không tải (Table 100.10).
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => setShowCameraModal(true)}
              className="px-3.5 py-2 rounded-xl text-xs font-bold bg-amber-400 hover:bg-amber-300 text-slate-950 flex items-center gap-1.5 transition-all shadow-md active:scale-95"
              title="Quét tem Nameplate máy điện đồng bộ bằng Camera AI"
            >
              <Camera size={14} />
              <span>Quét Nameplate (Camera AI)</span>
            </button>

            <button
              onClick={() => {
                setData(SAMPLE_SYNC_MACHINERY_FSE_DATA);
                setAnalysisResult(null);
              }}
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title="Tải kịch bản mẫu FSE Máy phát điện đồng bộ 20MW"
            >
              <RotateCcw size={14} />
              <span>Nạp Kịch Bản Mẫu FSE</span>
            </button>

            <button
              onClick={() => setShowPrintModal(true)}
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-blue-600 hover:bg-blue-500 text-white border border-blue-400/40 flex items-center gap-1.5 transition-all shadow-md"
            >
              <Printer size={15} />
              <span>In / Xuất Báo Cáo PDF (TEV)</span>
            </button>

            <button
              onClick={handleSaveToCmms}
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-emerald-700 hover:bg-emerald-600 text-white border border-emerald-500/40 flex items-center gap-1.5 transition-all shadow-md"
              title="Lưu biên bản kiểm tra vào kho báo cáo"
            >
              <Save size={15} />
              <span>Lưu vào Kho Báo Cáo</span>
            </button>

            <div className="flex gap-2">
              <button
                onClick={() => {
                  setJsonInput(JSON.stringify(data, null, 2));
                  setShowJsonModal(true);
                }}
                className="px-3 py-2 bg-white/10 hover:bg-white/20 border border-white/20 text-white rounded-xl text-xs font-medium flex items-center justify-center gap-1 transition-colors"
                title="Xem hoặc nhập mã JSON khảo sát hiện trường"
              >
                <Upload size={13} />
                <span>Mã JSON</span>
              </button>
            </div>

            <button
              onClick={handleAnalyze}
              disabled={isAnalyzing}
              className="px-4 py-2 rounded-xl text-xs font-bold bg-gradient-to-r from-amber-500 to-blue-500 hover:from-amber-400 hover:to-blue-400 text-slate-950 flex items-center gap-2 transition-all shadow-lg shadow-blue-500/20 disabled:opacity-50"
            >
              <Sparkles size={16} className={isAnalyzing ? 'animate-spin' : ''} />
              <span>{isAnalyzing ? 'Đang Phân Tích...' : '🤖 Phân Tích AI Chuyên Gia'}</span>
            </button>
          </div>
        </div>
      </div>

      {/* Save Notification Toast */}
      {savedSuccessMessage && (
        <div className="p-3 bg-emerald-500/10 border border-emerald-500/30 rounded-xl text-emerald-800 text-xs font-semibold flex items-center justify-between animate-in fade-in">
          <div className="flex items-center gap-2">
            <CheckCircle2 size={16} className="text-emerald-600" />
            <span>{savedSuccessMessage}</span>
          </div>
          {onNavigateToReports && (
            <button
              onClick={onNavigateToReports}
              className="text-xs font-bold text-emerald-700 underline hover:text-emerald-900"
            >
              Xem Kho Báo Cáo
            </button>
          )}
        </div>
      )}

      {/* CRITICAL ROTOR DEFECT BANNER */}
      {liveStats.hasCriticalPoleFail && (
        <div className="bg-red-50 border-2 border-red-500 rounded-xl p-4 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 shadow-md animate-pulse">
          <div className="flex items-start gap-3">
            <div className="w-10 h-10 rounded-xl bg-red-100 text-red-700 flex items-center justify-center shrink-0 mt-0.5">
              <AlertTriangle size={24} />
            </div>
            <div>
              <h4 className="text-xs sm:text-sm font-black text-red-950 uppercase tracking-wide">
                CẢNH BÁO TỐI CẤP: CRITICAL ROTOR DEFECT (NGẮN MẠCH VÒNG DÂY CỰC TỪ POLE {liveStats.worstPole})
              </h4>
              <p className="text-xs text-red-800 mt-1 leading-relaxed">
                Cực số {liveStats.worstPole} có mức sụt áp AC lệch {liveStats.worstDev > 0 ? '+' : ''}{liveStats.worstDev.toFixed(1)}% so với trung bình các cực (vượt quá xa ngưỡng cho phép 10% theo NETA Section 7.15.2.B.3).
                <strong className="block text-red-900 mt-0.5">
                  CẤM HÒA LƯỚI ĐÓNG TẢI! Nguy cơ phát sinh lực hút từ trường bất đối xứng (UMP) làm rung giật phá hủy gối trục và cháy bối dây cực từ!
                </strong>
              </p>
            </div>
          </div>
          <button
            onClick={handleAnalyze}
            className="px-4 py-2 bg-red-700 hover:bg-red-800 text-white rounded-lg text-xs font-bold shrink-0 transition-colors shadow-xs"
          >
            Xem Khuyến Nghị AI
          </button>
        </div>
      )}

      {/* Main Grid: Section 1, 2, 3 */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left Column: Section 1 & Section 2 */}
        <div className="space-y-6 lg:col-span-1">
          {/* Section 1: Site Info & Machine ID */}
          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-2">
              <h3 className="font-bold text-sm text-slate-900 flex items-center gap-2">
                <Compass size={16} className="text-blue-600" />
                1. Thông tin Hiện trường & Máy phát
              </h3>
              <button
                type="button"
                onClick={() => setShowCameraModal(true)}
                className="px-2.5 py-1 bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-300 rounded font-bold text-[11px] transition-all flex items-center gap-1 shadow-2xs"
              >
                <Camera size={13} className="text-amber-600" /> Quét Nameplate
              </button>
            </div>

            <div className="flex items-center justify-between bg-amber-50/70 border border-amber-200/80 rounded-lg p-2.5 px-3">
              <div className="flex items-center gap-2 text-xs text-amber-950 font-medium">
                <Camera size={15} className="text-amber-600 shrink-0" />
                <span>Chụp ảnh tem nhãn máy phát / động cơ đồng bộ</span>
              </div>
              <button
                type="button"
                onClick={() => setShowCameraModal(true)}
                className="px-3 py-1 bg-amber-600 hover:bg-amber-700 text-white rounded font-bold text-xs shadow-2xs transition-all shrink-0 ml-2"
              >
                Mở Camera
              </button>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="block text-slate-500 mb-1">Dự án / Nhà máy điện</label>
                <input
                  type="text"
                  value={data.site_info.project_name}
                  onChange={(e) => setData({ ...data, site_info: { ...data.site_info, project_name: e.target.value } })}
                  className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white focus:ring-2 focus:ring-blue-500 font-medium text-slate-800"
                />
              </div>
              <div>
                <label className="block text-slate-500 mb-1">Mã Tag Máy phát / Động cơ</label>
                <input
                  type="text"
                  value={data.site_info.generator_tag}
                  onChange={(e) => setData({ ...data, site_info: { ...data.site_info, generator_tag: e.target.value } })}
                  className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white focus:ring-2 focus:ring-blue-500 font-mono font-bold text-blue-900"
                />
              </div>
              <div>
                <label className="block text-slate-500 mb-1">Công suất, Điện áp & Tốc độ định mức</label>
                <input
                  type="text"
                  value={data.site_info.generator_rating}
                  onChange={(e) => setData({ ...data, site_info: { ...data.site_info, generator_rating: e.target.value } })}
                  className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white font-medium text-slate-800"
                />
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-slate-500 mb-1">Điện áp (kV)</label>
                  <input
                    type="number"
                    value={data.site_info.rated_voltage_kv || 11}
                    onChange={(e) => setData({ ...data, site_info: { ...data.site_info, rated_voltage_kv: Number(e.target.value) } })}
                    className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white text-slate-800 font-mono"
                  />
                </div>
                <div>
                  <label className="block text-slate-500 mb-1">Tốc độ (RPM)</label>
                  <input
                    type="number"
                    value={data.site_info.rated_speed_rpm || 1500}
                    onChange={(e) => setData({ ...data, site_info: { ...data.site_info, rated_speed_rpm: Number(e.target.value) } })}
                    className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white text-slate-800 font-mono"
                  />
                </div>
              </div>
              <div>
                <label className="block text-slate-500 mb-1">Kỹ sư kiểm tra (FSE TEV)</label>
                <FseEngineerDropdown
                  value={data.site_info.fse_name}
                  onChange={(val) => setData({ ...data, site_info: { ...data.site_info, fse_name: val } })}
                />
              </div>
              <div>
                <label className="block text-slate-500 mb-1">Ngày kiểm tra</label>
                <input
                  type="date"
                  value={data.site_info.test_date}
                  onChange={(e) => setData({ ...data, site_info: { ...data.site_info, test_date: e.target.value } })}
                  className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white font-medium text-slate-800"
                />
              </div>
            </div>
          </div>

          {/* Section 2: Visual & Mechanical Inspection (NETA 7.15.2.A) */}
          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs space-y-4">
            <h3 className="font-bold text-sm text-slate-900 flex items-center gap-2 border-b border-slate-100 pb-2">
              <ShieldCheck size={16} className="text-blue-600" />
              2. Kiểm Tra Thị Giác & Cơ Khí (NETA 7.15.2.A)
            </h3>

            <div className="space-y-3 text-xs">
              <label className="flex items-center gap-2.5 p-2 rounded-lg border border-slate-200 hover:bg-slate-50 cursor-pointer">
                <input
                  type="checkbox"
                  checked={data.visual_inspection.nameplate_match}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, nameplate_match: e.target.checked }
                  })}
                  className="rounded text-blue-600 focus:ring-blue-500 w-4 h-4"
                />
                <div className="flex-1">
                  <span className="font-bold text-slate-800 block">Khớp nhãn mác (Nameplate Match)</span>
                  <span className="text-[11px] text-slate-500">Stator, Rotor cực lồi và Exciter khớp 100% bản vẽ</span>
                </div>
              </label>

              <div className="grid grid-cols-2 gap-2 pt-1">
                <div>
                  <label className="block text-slate-500 mb-1">Vệ sinh & Bụi than</label>
                  <select
                    value={data.visual_inspection.physical_cleanliness}
                    onChange={(e) => setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, physical_cleanliness: e.target.value }
                    })}
                    className="w-full px-2.5 py-1.5 border rounded-lg bg-slate-50 font-medium"
                  >
                    <option value="Pass">Pass (Sạch sẽ)</option>
                    <option value="Fail">Fail (Nhiễm bẩn)</option>
                  </select>
                </div>
                <div>
                  <label className="block text-slate-500 mb-1">Khe hở không khí</label>
                  <select
                    value={data.visual_inspection.air_gap_uniformity}
                    onChange={(e) => setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, air_gap_uniformity: e.target.value }
                    })}
                    className="w-full px-2.5 py-1.5 border rounded-lg bg-slate-50 font-medium"
                  >
                    <option value="Pass">Pass (Đồng đều)</option>
                    <option value="Fail">Fail (Lệch cực)</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-slate-500 mb-1">Bộ kích từ (Exciter)</label>
                  <select
                    value={data.visual_inspection.brushless_exciter_condition}
                    onChange={(e) => setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, brushless_exciter_condition: e.target.value }
                    })}
                    className="w-full px-2.5 py-1.5 border rounded-lg bg-slate-50 font-medium"
                  >
                    <option value="Good">Good (Tốt)</option>
                    <option value="Needs Service">Cần bảo dưỡng</option>
                  </select>
                </div>
                <div>
                  <label className="block text-slate-500 mb-1">Lực siết bu-lông</label>
                  <select
                    value={data.visual_inspection.bolt_torque_check}
                    onChange={(e) => setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, bolt_torque_check: e.target.value }
                    })}
                    className="w-full px-2.5 py-1.5 border rounded-lg bg-slate-50 font-medium"
                  >
                    <option value="Pass">Pass (Table 100.12)</option>
                    <option value="Fail">Fail (Chưa đạt lực)</option>
                  </select>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Right 2 Columns: Section 3 Electrical & Pole-to-Pole AC Voltage Drop */}
        <div className="space-y-6 lg:col-span-2">
          {/* Section 3: Electrical Tests */}
          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs space-y-5">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h3 className="font-bold text-sm text-slate-900 flex items-center gap-2">
                <Zap size={16} className="text-amber-500" />
                3. Phép Đo Điện & Cực Rô-to (NETA 7.15.2.B)
              </h3>
              <span className={`px-2.5 py-1 rounded-full text-xs font-bold ${
                liveStats.overallSafe ? 'bg-emerald-100 text-emerald-800' : 'bg-red-100 text-red-800'
              }`}>
                {liveStats.overallSafe ? 'HỆ THỐNG ĐẠT CHUẨN' : 'CRITICAL ROTOR DEFECT'}
              </span>
            </div>

            {/* 3.1 CRITICAL: Pole-to-Pole AC Voltage Drop Test */}
            <div className={`p-4 rounded-xl border-2 transition-all space-y-3 ${
              liveStats.hasCriticalPoleFail ? 'bg-red-50/90 border-red-500' : 'bg-blue-50/70 border-blue-300'
            }`}>
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <Wind size={18} className={liveStats.hasCriticalPoleFail ? 'text-red-700' : 'text-blue-700'} />
                  <span className="font-extrabold text-xs sm:text-sm text-slate-900">
                    3.1 Thử Sụt Áp AC Từng Cực Rô-to (Pole-to-Pole AC Voltage Drop - NETA 7.15.2.B.3)
                  </span>
                </div>
                <span className={`px-2.5 py-0.5 rounded text-[11px] font-black uppercase tracking-wider ${
                  liveStats.hasCriticalPoleFail ? 'bg-red-200 text-red-900 border border-red-300' : 'bg-emerald-200 text-emerald-900'
                }`}>
                  {liveStats.hasCriticalPoleFail ? 'FAIL (NGẮN MẠCH VÒNG DÂY)' : 'PASS (ĐỒNG ĐỀU ≤10%)'}
                </span>
              </div>

              <p className="text-[11px] text-slate-600">
                Tiêu chuẩn NETA Section 7.15.2.B.3 quy định: Điện áp sụt trên từng cuộn cực từ (Field pole) không được lệch quá <strong>±10%</strong> so với giá trị sụt áp trung bình toàn bộ các cực.
              </p>

              {/* Pole Inputs Grid */}
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
                {liveStats.analyzedPoles.map((p) => {
                  const poleKey = `pole_${p.id}` as keyof typeof data.electrical_tests.pole_to_pole_ac_voltage_drop_volts;
                  return (
                    <div
                      key={p.id}
                      className={`p-3 rounded-xl border ${
                        !p.pass ? 'bg-red-100/90 border-red-400' : 'bg-white border-slate-200'
                      }`}
                    >
                      <div className="flex items-center justify-between mb-1">
                        <span className="font-bold text-slate-800">Cực {p.id} (Pole {p.id})</span>
                        <span className={`text-[10px] font-extrabold px-1.5 py-0.2 rounded ${
                          !p.pass ? 'bg-red-600 text-white' : 'bg-slate-100 text-slate-600'
                        }`}>
                          {p.dev > 0 ? '+' : ''}{p.dev.toFixed(1)}%
                        </span>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <input
                          type="number"
                          step="0.1"
                          value={p.val}
                          onChange={(e) => {
                            setData({
                              ...data,
                              electrical_tests: {
                                ...data.electrical_tests,
                                pole_to_pole_ac_voltage_drop_volts: {
                                  ...data.electrical_tests.pole_to_pole_ac_voltage_drop_volts,
                                  [poleKey]: Number(e.target.value)
                                }
                              }
                            });
                          }}
                          className={`w-full px-2 py-1 font-mono font-bold rounded text-sm border ${
                            !p.pass ? 'bg-red-50 border-red-500 text-red-900' : 'bg-slate-50 border-slate-300 text-slate-800'
                          }`}
                        />
                        <span className="text-xs font-bold text-slate-500">V</span>
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* Sụt áp trung bình */}
              <div className="p-2.5 bg-white/80 rounded-lg border border-slate-200 flex items-center justify-between text-xs">
                <span className="text-slate-600 font-medium">Sụt áp AC trung bình tính toán:</span>
                <span className="font-bold font-mono text-slate-900">{liveStats.avgDrop.toFixed(2)} V</span>
              </div>

              {liveStats.hasCriticalPoleFail && (
                <div className="p-2.5 bg-red-100 border border-red-300 rounded-lg text-xs font-medium text-red-950 flex items-start gap-2">
                  <AlertTriangle size={15} className="text-red-800 shrink-0 mt-0.5" />
                  <span>
                    <strong>Phân tích IEEE 115:</strong> Cực số {liveStats.worstPole} bị sụt áp thấp hơn {Math.abs(liveStats.worstDev).toFixed(1)}% chứng minh một số vòng dây trong bối cực đã bị chập ngắn mạch (Shorted turns). Cần tháo cực số {liveStats.worstPole} quấn lại cách điện.
                  </span>
                </div>
              )}
            </div>

            {/* 3.2 Stator IR & PI (Table 100.11) */}
            <div>
              <div className="flex items-center justify-between mb-2">
                <span className="font-bold text-xs text-slate-700 flex items-center gap-1.5">
                  <Activity size={14} className="text-blue-600" />
                  3.2 Cách điện Stator & Chỉ số Phân cực PI (Table 100.11)
                </span>
                <span className={`text-[11px] font-bold ${liveStats.piPass && liveStats.statorIrPass ? 'text-emerald-600' : 'text-red-600'}`}>
                  {liveStats.piPass && liveStats.statorIrPass ? `✓ ĐẠT (PI = ${liveStats.computedPi} ≥ 2.0)` : `⚠️ CHƯA ĐẠT (PI = ${liveStats.computedPi})`}
                </span>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
                <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
                  <label className="text-slate-600 block mb-1">IR 1 phút (2500V DC):</label>
                  <div className="flex items-center gap-2">
                    <input
                      type="number"
                      value={data.electrical_tests.stator_ir_40c_megohms.stator_1min_ir_2500v}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          stator_ir_40c_megohms: {
                            ...data.electrical_tests.stator_ir_40c_megohms,
                            stator_1min_ir_2500v: Number(e.target.value)
                          }
                        }
                      })}
                      className="w-full bg-white border rounded px-2.5 py-1.5 font-bold font-mono text-slate-800"
                    />
                    <span className="text-xs font-semibold text-slate-500">MΩ</span>
                  </div>
                </div>

                <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
                  <label className="text-slate-600 block mb-1">IR 10 phút (2500V DC):</label>
                  <div className="flex items-center gap-2">
                    <input
                      type="number"
                      value={data.electrical_tests.stator_ir_40c_megohms.stator_10min_ir_2500v}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          stator_ir_40c_megohms: {
                            ...data.electrical_tests.stator_ir_40c_megohms,
                            stator_10min_ir_2500v: Number(e.target.value)
                          }
                        }
                      })}
                      className="w-full bg-white border rounded px-2.5 py-1.5 font-bold font-mono text-slate-800"
                    />
                    <span className="text-xs font-semibold text-slate-500">MΩ</span>
                  </div>
                </div>

                <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
                  <label className="text-slate-600 block mb-1">Chỉ số Phân cực PI (10min/1min):</label>
                  <div className="px-2.5 py-1.5 bg-blue-50 border border-blue-200 rounded font-bold font-mono text-sm text-blue-900 flex items-center justify-between">
                    <span>{liveStats.computedPi}</span>
                    <span className="text-[10px] text-blue-700 font-normal">Tiêu chuẩn: ≥ 2.0</span>
                  </div>
                </div>
              </div>
            </div>

            {/* 3.3 Main Field Rotor IR & Exciter IR */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
                <label className="text-slate-600 font-bold block mb-1">Cách điện Cuộn Kích từ Rô-to (1000V DC):</label>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    value={data.electrical_tests.main_field_rotor_ir_1000v_megohms}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        main_field_rotor_ir_1000v_megohms: Number(e.target.value)
                      }
                    })}
                    className="w-full bg-white border rounded px-2.5 py-1.5 font-bold font-mono text-slate-800"
                  />
                  <span className="text-xs font-semibold text-slate-500">MΩ</span>
                </div>
                <span className="text-[10px] text-slate-500 mt-1 block">Tiêu chuẩn: ≥ 100 MΩ</span>
              </div>

              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
                <label className="text-slate-600 font-bold block mb-1">Cụm Điốt Quay Kích Từ (Diodes Check):</label>
                <input
                  type="text"
                  value={data.electrical_tests.rotating_diodes_check}
                  onChange={(e) => setData({
                    ...data,
                    electrical_tests: {
                      ...data.electrical_tests,
                      rotating_diodes_check: e.target.value
                    }
                  })}
                  className="w-full bg-white border rounded px-2.5 py-1.5 font-medium text-slate-800"
                />
                <span className="text-[10px] text-slate-500 mt-1 block">Kiểm tra thông mạch 6 điốt quay</span>
              </div>
            </div>

            {/* 3.4 Stator Phase Resistance Balance & Vibration */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              {/* Stator Phase Resistance */}
              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-xs text-slate-800">
                    3.4 Điện trở Stator giữa các pha (Độ lệch ≤ 5%)
                  </span>
                  <span className={`text-[11px] font-bold ${liveStats.resPass ? 'text-emerald-600' : 'text-amber-600'}`}>
                    {liveStats.resPass ? `✓ CÂN BẰNG (${liveStats.resDev.toFixed(2)}%)` : `⚠️ LỆCH (${liveStats.resDev.toFixed(2)}%)`}
                  </span>
                </div>
                <div className="grid grid-cols-3 gap-2 text-xs">
                  <div>
                    <span className="text-[10px] text-slate-500 block">Pha A-B (Ω)</span>
                    <input
                      type="number"
                      step="0.001"
                      value={data.electrical_tests.stator_phase_resistance_ohms.phase_ab}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          stator_phase_resistance_ohms: {
                            ...data.electrical_tests.stator_phase_resistance_ohms,
                            phase_ab: Number(e.target.value)
                          }
                        }
                      })}
                      className="w-full bg-white border rounded px-2 py-1 font-mono font-bold text-slate-800"
                    />
                  </div>
                  <div>
                    <span className="text-[10px] text-slate-500 block">Pha B-C (Ω)</span>
                    <input
                      type="number"
                      step="0.001"
                      value={data.electrical_tests.stator_phase_resistance_ohms.phase_bc}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          stator_phase_resistance_ohms: {
                            ...data.electrical_tests.stator_phase_resistance_ohms,
                            phase_bc: Number(e.target.value)
                          }
                        }
                      })}
                      className="w-full bg-white border rounded px-2 py-1 font-mono font-bold text-slate-800"
                    />
                  </div>
                  <div>
                    <span className="text-[10px] text-slate-500 block">Pha C-A (Ω)</span>
                    <input
                      type="number"
                      step="0.001"
                      value={data.electrical_tests.stator_phase_resistance_ohms.phase_ca}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          stator_phase_resistance_ohms: {
                            ...data.electrical_tests.stator_phase_resistance_ohms,
                            phase_ca: Number(e.target.value)
                          }
                        }
                      })}
                      className="w-full bg-white border rounded px-2 py-1 font-mono font-bold text-slate-800"
                    />
                  </div>
                </div>
              </div>

              {/* Vibration */}
              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-xs text-slate-800">
                    3.5 Độ Rung Không Tải Tháo Khớp (Table 100.10)
                  </span>
                  <span className={`text-[11px] font-bold ${liveStats.vibPass ? 'text-emerald-600' : 'text-red-600'}`}>
                    {liveStats.vibPass ? '≤ 0.15 in/sec pk (ĐẠT)' : 'VƯỢT NGƯỠNG RUNG'}
                  </span>
                </div>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.uncoupled_vibration_in_sec_pk}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        uncoupled_vibration_in_sec_pk: Number(e.target.value)
                      }
                    })}
                    className="w-full bg-white border rounded px-2.5 py-1.5 font-bold font-mono text-sm text-slate-800"
                  />
                  <span className="text-xs font-semibold text-slate-500">in/sec pk</span>
                </div>
                <span className="text-[10px] text-slate-500 block">
                  Đo tại gối trục truyền động và gối trục kích từ không tải.
                </span>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Section 4: AI Analysis Results */}
      {analysisResult && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-xl overflow-hidden animate-in fade-in duration-300">
          <div className="p-5 sm:p-6 bg-slate-900 text-white flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className={`w-12 h-12 rounded-xl flex items-center justify-center text-white shrink-0 ${
                analysisResult.overall_status === 'PASS' ? 'bg-emerald-600' : 'bg-red-600'
              }`}>
                {analysisResult.overall_status === 'PASS' ? <CheckCircle2 size={28} /> : <AlertTriangle size={28} />}
              </div>
              <div>
                <span className="text-[11px] font-bold text-amber-400 uppercase tracking-wider">
                  KẾT QUẢ ĐỐI SOÁT AI CHUYÊN GIA NETA ATS-2025
                </span>
                <h2 className="text-xl sm:text-2xl font-black">
                  Trạng thái: {analysisResult.overall_status === 'PASS' ? 'PASS (ĐẠT CHUẨN)' : 'FAIL / CRITICAL ROTOR DEFECT'}
                </h2>
              </div>
            </div>

            <button
              onClick={() => setShowPrintModal(true)}
              className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white text-xs font-bold rounded-xl shadow-md transition-all flex items-center justify-center gap-2 self-start sm:self-auto"
            >
              <Printer size={15} />
              <span>In Báo Cáo PDF Chính Thức</span>
            </button>
          </div>

          <div className="p-6 bg-slate-50 space-y-6">
            <div className={`p-4 rounded-xl border font-medium text-xs sm:text-sm leading-relaxed ${
              analysisResult.overall_status === 'PASS' ? 'bg-emerald-50 border-emerald-200 text-emerald-900' : 'bg-red-50 border-red-200 text-red-900'
            }`}>
              <strong>Tóm lược đánh giá:</strong> {analysisResult.status_reason}
            </div>

            <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-xs prose prose-slate max-w-none prose-headings:font-bold prose-h1:text-xl prose-h2:text-base prose-h3:text-sm prose-p:text-xs prose-p:leading-relaxed prose-table:text-xs prose-th:bg-slate-100 prose-th:p-2 prose-td:p-2">
              <div className="whitespace-pre-wrap font-sans text-xs leading-relaxed text-slate-800">
                {analysisResult.markdown_report}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* JSON Modal */}
      {showJsonModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-white w-full max-w-2xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[85vh]">
            <div className="p-4 border-b border-slate-100 flex items-center justify-between bg-slate-50">
              <h3 className="font-bold text-slate-800 text-sm flex items-center gap-2">
                <FileText size={16} className="text-blue-600" />
                Mã JSON Khảo Sát Hiện Trường Máy Phát Đồng Bộ
              </h3>
              <button onClick={() => setShowJsonModal(false)} className="text-slate-400 hover:text-slate-600 text-lg">
                &times;
              </button>
            </div>
            <div className="p-4 flex-1 overflow-y-auto">
              <p className="text-xs text-slate-500 mb-2">
                Dán chuỗi JSON kiểm tra máy điện đồng bộ từ hệ thống SCADA / Hãng sản xuất:
              </p>
              <textarea
                value={jsonInput}
                onChange={(e) => setJsonInput(e.target.value)}
                rows={14}
                className="w-full p-3 font-mono text-xs bg-slate-900 text-emerald-400 rounded-xl border border-slate-800 focus:outline-hidden focus:ring-2 focus:ring-blue-500"
              />
            </div>
            <div className="p-4 border-t border-slate-100 bg-slate-50 flex justify-between gap-2">
              <button
                onClick={() => {
                  navigator.clipboard.writeText(jsonInput);
                  setCopied(true);
                  setTimeout(() => setCopied(false), 2000);
                }}
                className="px-3 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 text-xs font-semibold rounded-lg flex items-center gap-1.5"
              >
                {copied ? <Check size={14} className="text-emerald-600" /> : <Eye size={14} />}
                <span>{copied ? 'Đã sao chép' : 'Sao chép JSON'}</span>
              </button>
              <div className="flex gap-2">
                <button
                  onClick={() => setShowJsonModal(false)}
                  className="px-4 py-2 text-xs font-medium text-slate-600 hover:bg-slate-200 rounded-lg"
                >
                  Đóng
                </button>
                <button
                  onClick={() => {
                    try {
                      const parsed = JSON.parse(jsonInput);
                      setData(parsed);
                      setShowJsonModal(false);
                      setSavedSuccessMessage('Đã tải dữ liệu JSON máy điện đồng bộ thành công!');
                      setTimeout(() => setSavedSuccessMessage(null), 3000);
                    } catch {
                      alert('Chuỗi JSON không đúng định dạng!');
                    }
                  }}
                  className="px-4 py-2 text-xs font-bold text-white bg-blue-600 hover:bg-blue-500 rounded-lg shadow-md"
                >
                  Áp dụng Dữ Liệu
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Printable TEV Report Modal */}
      <TevReportPrintModal
        isOpen={showPrintModal}
        onClose={() => setShowPrintModal(false)}
        title="BIÊN BẢN THỬ NGHIỆM ĐỘNG CƠ & MÁY PHÁT ĐIỆN ĐỒNG BỘ"
        standard="ANSI/NETA ATS-2025 Section 7.15.2"
        siteInfo={data.site_info}
        overallStatus={analysisResult?.overall_status || (liveStats.overallSafe ? 'PASS' : 'FAIL')}
        statusReason={analysisResult?.status_reason || (liveStats.overallSafe ? 'Đạt chuẩn NETA ATS-2025 Section 7.15.2.' : `Cực từ số ${liveStats.worstPole} bị ngắn mạch giữa các vòng dây. CẤM HÒA LƯỚI ĐÓNG TẢI!`)}
        visualChecks={visualChecksForPrint}
        electricalTable={electricalTableForPrint}
        warnings={
          liveStats.overallSafe
            ? []
            : [
                liveStats.hasCriticalPoleFail ? `Cực số ${liveStats.worstPole} có sụt áp AC lệch ${liveStats.worstDev > 0 ? '+' : ''}${liveStats.worstDev.toFixed(1)}% vượt quá 10% (Chỉ thị ngắn mạch cuộn cực từ).` : '',
                !liveStats.resPass ? `Điện trở Stator giữa các pha lệch ${liveStats.resDev.toFixed(2)}% vượt quá 5.0%.` : '',
                !liveStats.piPass ? `Chỉ số phân cực PI = ${liveStats.computedPi} < 2.0.` : '',
                !liveStats.vibPass ? `Độ rung ${data.electrical_tests.uncoupled_vibration_in_sec_pk} in/sec pk vượt ngưỡng Table 100.10.` : ''
              ].filter(Boolean)
        }
        recommendations={[
          'CẢNH BÁO TỐI CẤP: Không cho máy phát hòa lưới đóng tải để tránh gây mất cân bằng từ trường và rung giật hỏng bạc đạn.',
          `Cô lập rô-to, tháo cuộn dây cực số ${liveStats.worstPole} để kiểm tra cách điện giữa các vòng dây (turn-to-turn insulation) hoặc quấn lại cuộn kích từ Pole ${liveStats.worstPole}.`,
          'Đo lại sụt áp AC cực rô-to sau khi xử lý, đảm bảo độ lệch giữa tất cả các cực nằm trong phạm vi ≤ ±10%.'
        ]}
      />

      {/* Camera Nameplate OCR Modal */}
      {showCameraModal && (
        <NameplateScannerModal
          isOpen={showCameraModal}
          onClose={() => setShowCameraModal(false)}
          onApplyData={handleApplyNameplateData}
          currentTag={data.site_info.generator_tag}
        />
      )}
    </div>
  );
};
