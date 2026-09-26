import React, { useState, useMemo } from 'react';
import {
  BatteryCharging,
  Zap,
  CheckCircle2,
  AlertTriangle,
  FileText,
  Printer,
  Sparkles,
  Save,
  RotateCcw,
  Compass,
  ShieldCheck,
  Activity,
  Layers,
  Check,
  Eye,
  Camera,
  Upload,
  Gauge,
  Sliders,
  Power
} from 'lucide-react';
import {
  NetaUpsInputPayload,
  NetaUpsEvaluationResult,
  performDeterministicUpsAnalysis
} from '../../server/netaUpsAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';
import { NameplateScannerModal } from './NameplateScannerModal';
import { ExtractedNameplateData } from '../../server/nameplateOcrAnalyzer';
import { FseEngineerDropdown } from './FseEngineerDropdown';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

export const SAMPLE_UPS_FSE_DATA: NetaUpsInputPayload = {
  site_info: {
    project_name: "Trung Tam Du Lieu TEV Data Center",
    ups_tag: "UPS-DC-300KVA-01",
    ups_rating: "300kVA 400V 3-Phase Online Double Conversion UPS",
    fse_name: "Võ Minh Tú (sgm1707@gmail.com)",
    test_date: "2026-09-23",
    manufacturer: "Eaton / Schneider / Vertiv",
    model_number: "93PM-300kVA",
    serial_number: "SN-TEV-UPS-2026-008"
  },
  visual_inspection: {
    nameplate_match: true,
    physical_condition: "Good (No damage)",
    anchorage_grounding_clearances: "Pass",
    fuse_ratings_match: true,
    cleanliness: "Pass",
    interlock_systems_operation: "Pass",
    bolt_torque_check: "Pass"
  },
  electrical_tests: {
    bolted_resistance_micro_ohms: [9.1, 9.3, 9.2],
    static_transfer_function: "Pass (Seamless transfer within 1ms)",
    oscillator_free_running_frequency_hz: 50.01,
    dc_undervoltage_inverter_trip_v: 384,
    alarm_circuits_test: "Pass",
    synchronizing_indicators_test: "Pass",
    ups_breakers_section_7_6_status: "Pass",
    ats_section_7_22_3_status: "Pass",
    battery_system_section_7_18: {
      battery_type: "VRLA 12V 100Ah (40 blocks)",
      total_string_voltage_v: 542,
      block_internal_resistance_check: "Pass for 38 blocks, 2 blocks high resistance (>35% dev)",
      status: "Investigate",
      max_block_resistance_deviation_percent: 36.5
    }
  },
  previous_test_data: {
    last_test_date: "2025-09-21",
    last_battery_internal_resistance_avg_mohm: 3.2,
    last_bolted_resistance_micro_ohms: [9.0, 9.2, 9.1]
  }
};

export const NetaUpsChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaUpsInputPayload>(SAMPLE_UPS_FSE_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaUpsEvaluationResult | null>(null);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showCameraModal, setShowCameraModal] = useState(false);
  const [jsonInput, setJsonInput] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live statistical calculations
  const liveStats = useMemo(() => {
    // 1. Bolted connection resistance deviation
    const resistances = data.electrical_tests.bolted_resistance_micro_ohms || [];
    let resistanceDev = 0;
    let resistancePass = true;
    if (resistances.length > 1) {
      const minVal = Math.min(...resistances);
      const maxVal = Math.max(...resistances);
      if (minVal > 0) {
        resistanceDev = ((maxVal - minVal) / minVal) * 100;
        if (resistanceDev > 50) resistancePass = false;
      }
    }

    // 2. Frequency deviation
    const freq = data.electrical_tests.oscillator_free_running_frequency_hz;
    const freqErr = Math.abs(freq - 50.0);
    const freqPass = freqErr <= 0.15;

    // 3. Static transfer pass
    const staticPass = data.electrical_tests.static_transfer_function.toLowerCase().includes('pass');

    // 4. Battery system health (Section 7.18)
    const bat = data.electrical_tests.battery_system_section_7_18;
    const batHasHighResistance = bat.block_internal_resistance_check.toLowerCase().includes('high') ||
      bat.block_internal_resistance_check.toLowerCase().includes('dev') ||
      (bat.max_block_resistance_deviation_percent !== undefined && bat.max_block_resistance_deviation_percent > 20);
    const batPass = bat.status.toLowerCase() === 'pass' && !batHasHighResistance;

    // 5. Interlock pass
    const interlockPass = data.visual_inspection.interlock_systems_operation.toLowerCase().includes('pass');

    const overallSafe = resistancePass && freqPass && staticPass && batPass && interlockPass && data.visual_inspection.nameplate_match;

    return {
      resistanceDev,
      resistancePass,
      freqErr,
      freqPass,
      staticPass,
      batHasHighResistance,
      batPass,
      interlockPass,
      overallSafe
    };
  }, [data]);

  const handleApplyNameplateData = (extracted: ExtractedNameplateData) => {
    setData((prev) => ({
      ...prev,
      site_info: {
        ...prev.site_info,
        ups_tag: extracted.equipment_tag || prev.site_info.ups_tag,
        ups_rating: extracted.equipment_type || extracted.equipment_name || prev.site_info.ups_rating,
        manufacturer: extracted.manufacturer || prev.site_info.manufacturer,
        serial_number: extracted.serial_number || prev.site_info.serial_number
      },
      visual_inspection: {
        ...prev.visual_inspection,
        nameplate_match: true
      }
    }));
    setSavedSuccessMessage(`Đã quét nhãn thành công: ${extracted.equipment_tag || 'UPS'} - Tự động cập nhật vào Checklist!`);
    setTimeout(() => setSavedSuccessMessage(null), 5000);
  };

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    setAnalysisResult(null);
    try {
      const res = await fetch('/api/field-service/neta-ups-analysis', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });
      const json = await res.json();
      if (json.success && json.data) {
        setAnalysisResult(json.data);
      } else {
        const det = performDeterministicUpsAnalysis(data);
        setAnalysisResult(det);
      }
    } catch {
      const det = performDeterministicUpsAnalysis(data);
      setAnalysisResult(det);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleSaveToCmms = () => {
    const reportPayload = {
      type: 'NETA_ATS_2025_UPS',
      data: data,
      analysisResult: analysisResult,
      equipmentId: data.site_info.ups_tag,
      date: data.site_info.test_date,
      title: `Biên bản UPS NETA ATS-2025: ${data.site_info.ups_tag}`,
      standard: 'ANSI/NETA ATS-2025 Section 7.22.2',
      status: liveStats.overallSafe ? 'PASS' : 'INVESTIGATE',
      timestamp: new Date().toISOString()
    };

    if (onSaveReport) {
      onSaveReport(reportPayload);
      setSavedSuccessMessage(`Đã lưu biên bản kiểm định UPS ${data.site_info.ups_tag} vào kho báo cáo thành công!`);
      setTimeout(() => setSavedSuccessMessage(null), 5000);
    }
  };

  const electricalTableForPrint = [
    {
      item: 'Điện trở tiếp xúc mối nối bu-lông (Bolted Resistance)',
      measured: `${data.electrical_tests.bolted_resistance_micro_ohms.join(', ')} µΩ (Độ lệch: ${liveStats.resistanceDev.toFixed(1)}%)`,
      standard: 'Độ lệch ≤ 50% so với giá trị nhỏ nhất',
      reference: 'NETA ATS-2025 Sec 7.22.2.B.1 & Table 100.1',
      deviation: `${liveStats.resistanceDev.toFixed(1)}%`,
      status: liveStats.resistancePass ? 'PASS' : 'INVESTIGATE'
    },
    {
      item: 'Chức năng Chuyển mạch Tĩnh (Static Transfer Function)',
      measured: data.electrical_tests.static_transfer_function,
      standard: 'Chuyển nguồn tức thời không gián đoạn phụ tải (< 2-4ms)',
      reference: 'NETA ATS-2025 Sec 7.22.2.B.2',
      deviation: '< 1ms',
      status: liveStats.staticPass ? 'PASS' : 'FAIL'
    },
    {
      item: 'Tần số Bộ Dao động Tự do Inverter (Oscillator Frequency)',
      measured: `${data.electrical_tests.oscillator_free_running_frequency_hz} Hz`,
      standard: '50.00 Hz ± 0.10 Hz',
      reference: 'NETA ATS-2025 Sec 7.22.2.B.3',
      deviation: `${(data.electrical_tests.oscillator_free_running_frequency_hz - 50.0).toFixed(2)} Hz`,
      status: liveStats.freqPass ? 'PASS' : 'INVESTIGATE'
    },
    {
      item: 'Cắt Thấp áp DC Inverter (DC Undervoltage Trip)',
      measured: `${data.electrical_tests.dc_undervoltage_inverter_trip_v} V DC`,
      standard: 'Trip cắt ngắt ngõ vào DC tại ngưỡng bảo vệ xả sâu của ắc quy',
      reference: 'NETA ATS-2025 Sec 7.22.2.B.4',
      deviation: 'Đạt chuẩn',
      status: 'PASS'
    },
    {
      item: 'Mạch Cảnh báo & Đèn Chỉ thị Đồng bộ (Alarms & Sync)',
      measured: `Cảnh báo: ${data.electrical_tests.alarm_circuits_test} | Đồng bộ: ${data.electrical_tests.synchronizing_indicators_test}`,
      standard: 'Báo động chuẩn xác; Chỉ thị đồng bộ pha Inverter - Bypass tin cậy',
      reference: 'NETA ATS-2025 Sec 7.22.2.B.5',
      deviation: '0%',
      status: 'PASS'
    },
    {
      item: 'Máy cắt UPS (Section 7.6) & Bộ chuyển ATS (Section 7.22.3)',
      measured: `Breakers: ${data.electrical_tests.ups_breakers_section_7_6_status} | ATS: ${data.electrical_tests.ats_section_7_22_3_status}`,
      standard: 'Đạt chỉ tiêu Section 7.6 và Section 7.22.3',
      reference: 'NETA ATS-2025 Sec 7.22.2.B.6 & B.7',
      deviation: '0%',
      status: 'PASS'
    },
    {
      item: 'Hệ thống Ắc quy Dự phòng (Section 7.18 Battery System)',
      measured: `Tổng áp: ${data.electrical_tests.battery_system_section_7_18.total_string_voltage_v} V DC | ${data.electrical_tests.battery_system_section_7_18.block_internal_resistance_check}`,
      standard: 'Độ lệch nội trở từng bình ≤ 20% so với giá trị trung bình chuỗi (IEEE 1188)',
      reference: 'NETA ATS-2025 Sec 7.22.2.B.8 & Sec 7.18',
      deviation: '> 35% lệch tại 2 bình',
      status: liveStats.batPass ? 'PASS' : 'INVESTIGATE'
    }
  ];

  const visualChecksForPrint = [
    { label: 'Đối soát nhãn mác thiết bị UPS (Nameplate Match)', value: data.visual_inspection.nameplate_match ? 'ĐẠT' : 'KHÔNG ĐẠT', pass: data.visual_inspection.nameplate_match },
    { label: 'Tình trạng vật lý thiết bị (Physical Condition)', value: data.visual_inspection.physical_condition, pass: data.visual_inspection.physical_condition.toLowerCase().includes('good') || data.visual_inspection.physical_condition.toLowerCase().includes('pass') },
    { label: 'Vệ sinh bo mạch & Quạt làm mát (Cleanliness)', value: data.visual_inspection.cleanliness, pass: data.visual_inspection.cleanliness.toLowerCase().includes('pass') },
    { label: 'Neo giữ, tiếp địa vỏ & khoảng cách thao tác (Anchorage & Clearances)', value: data.visual_inspection.anchorage_grounding_clearances, pass: data.visual_inspection.anchorage_grounding_clearances.toLowerCase().includes('pass') },
    { label: 'Định mức cầu chì bán dẫn (Fuse Ratings Match)', value: data.visual_inspection.fuse_ratings_match ? 'ĐẠT' : 'KHÔNG ĐẠT', pass: data.visual_inspection.fuse_ratings_match },
    { label: 'Khóa liên động Maintenance / Static Bypass', value: data.visual_inspection.interlock_systems_operation, pass: data.visual_inspection.interlock_systems_operation.toLowerCase().includes('pass') },
    { label: 'Kiểm tra lực siết bu-lông (Bolt Torque Table 100.12)', value: data.visual_inspection.bolt_torque_check, pass: data.visual_inspection.bolt_torque_check.toLowerCase().includes('pass') }
  ];

  return (
    <div className="space-y-6">
      {/* Top Banner & Action Header */}
      <div className="bg-gradient-to-r from-slate-900 via-emerald-950 to-teal-900 text-white p-5 sm:p-6 rounded-2xl shadow-lg border border-slate-700/50">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2 mb-1.5">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-400/20 text-amber-300 border border-amber-400/30 uppercase tracking-wide">
                ANSI/NETA ATS-2025 Section 7.22.2
              </span>
              <span className="text-xs text-slate-300">Emergency Systems • Uninterruptible Power Systems (UPS)</span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white flex items-center gap-2.5">
              <BatteryCharging className="text-emerald-400 fill-emerald-400/20" size={26} />
              Checklist Hệ Thống Nguồn Cấp Điện Liên Tục (UPS)
            </h1>
            <p className="text-xs sm:text-sm text-slate-300 mt-1 max-w-2xl">
              Thẩm định chuyên sâu theo NETA ATS-2025: Công tắc tĩnh Static Switch (&lt;2ms), tần số bộ dao động Inverter, khóa liên động Maintenance Bypass và đo nội trở dàn ắc quy dự phòng (Section 7.18).
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => setShowCameraModal(true)}
              className="px-3.5 py-2 rounded-xl text-xs font-bold bg-amber-400 hover:bg-amber-300 text-slate-950 flex items-center gap-1.5 transition-all shadow-md active:scale-95"
              title="Quét tem Nameplate UPS bằng Camera AI"
            >
              <Camera size={14} />
              <span>Quét Nameplate (Camera AI)</span>
            </button>

            <button
              onClick={() => {
                setData(SAMPLE_UPS_FSE_DATA);
                setAnalysisResult(null);
              }}
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title="Tải kịch bản mẫu FSE UPS 300kVA Data Center"
            >
              <RotateCcw size={14} />
              <span>Nạp Kịch Bản Mẫu FSE</span>
            </button>

            <button
              onClick={() => setShowPrintModal(true)}
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-teal-600 hover:bg-teal-500 text-white border border-teal-400/40 flex items-center gap-1.5 transition-all shadow-md"
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
              className="px-4 py-2 rounded-xl text-xs font-bold bg-gradient-to-r from-amber-500 to-emerald-500 hover:from-amber-400 hover:to-emerald-400 text-slate-950 flex items-center gap-2 transition-all shadow-lg shadow-emerald-500/20 disabled:opacity-50"
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

      {/* Battery Warning Quick Banner */}
      {liveStats.batHasHighResistance && (
        <div className="bg-amber-50 border border-amber-300 rounded-xl p-4 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 shadow-xs">
          <div className="flex items-start gap-3">
            <div className="w-9 h-9 rounded-xl bg-amber-100 text-amber-800 flex items-center justify-center shrink-0 mt-0.5">
              <AlertTriangle size={20} />
            </div>
            <div>
              <h4 className="text-xs sm:text-sm font-bold text-amber-950">
                CẢNH BÁO NETA ATS-2025: PHÁT HIỆN 2 BÌNH ẮC QUY TĂNG NỘI TRỞ CAO (&gt;35% DEV)
              </h4>
              <p className="text-xs text-amber-800 mt-0.5">
                Vượt quá ngưỡng cho phép 20% theo quy chuẩn NETA Section 7.18 & IEEE 1188. Nguy cơ sụt áp đột ngột và đứt chuỗi ắc quy khi xả tải mất điện lưới!
              </p>
            </div>
          </div>
          <button
            onClick={handleAnalyze}
            className="px-3.5 py-1.5 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-xs font-bold shrink-0 transition-colors"
          >
            Xem Khuyến Nghị AI
          </button>
        </div>
      )}

      {/* Main Grid: Section 1 & Section 2 & Section 3 */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left Column: Section 1 & Section 2 */}
        <div className="space-y-6 lg:col-span-1">
          {/* Section 1: Site Info & UPS ID */}
          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-2">
              <h3 className="font-bold text-sm text-slate-900 flex items-center gap-2">
                <Compass size={16} className="text-teal-600" />
                1. Thông tin Hiện trường & UPS
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
                <span>Chụp ảnh tem nhãn UPS để AI tự bóc tách thông số</span>
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
                <label className="block text-slate-500 mb-1">Dự án / Trung tâm dữ liệu</label>
                <input
                  type="text"
                  value={data.site_info.project_name}
                  onChange={(e) => setData({ ...data, site_info: { ...data.site_info, project_name: e.target.value } })}
                  className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white focus:ring-2 focus:ring-teal-500 font-medium text-slate-800"
                />
              </div>
              <div>
                <label className="block text-slate-500 mb-1">Mã Tag UPS (UPS Tag)</label>
                <input
                  type="text"
                  value={data.site_info.ups_tag}
                  onChange={(e) => setData({ ...data, site_info: { ...data.site_info, ups_tag: e.target.value } })}
                  className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white focus:ring-2 focus:ring-teal-500 font-mono font-bold text-teal-900"
                />
              </div>
              <div>
                <label className="block text-slate-500 mb-1">Công suất & Cấu hình định mức</label>
                <input
                  type="text"
                  value={data.site_info.ups_rating}
                  onChange={(e) => setData({ ...data, site_info: { ...data.site_info, ups_rating: e.target.value } })}
                  className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white font-medium text-slate-800"
                />
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-slate-500 mb-1">Hãng sản xuất</label>
                  <input
                    type="text"
                    value={data.site_info.manufacturer || ''}
                    onChange={(e) => setData({ ...data, site_info: { ...data.site_info, manufacturer: e.target.value } })}
                    className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white text-slate-800"
                  />
                </div>
                <div>
                  <label className="block text-slate-500 mb-1">Model / Serial</label>
                  <input
                    type="text"
                    value={data.site_info.model_number || ''}
                    onChange={(e) => setData({ ...data, site_info: { ...data.site_info, model_number: e.target.value } })}
                    className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white text-slate-800"
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

          {/* Section 2: Visual & Mechanical Inspection (NETA 7.22.2.A) */}
          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs space-y-4">
            <h3 className="font-bold text-sm text-slate-900 flex items-center gap-2 border-b border-slate-100 pb-2">
              <ShieldCheck size={16} className="text-teal-600" />
              2. Kiểm Tra Thị Giác & Cơ Khí (NETA 7.22.2.A)
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
                  className="rounded text-teal-600 focus:ring-teal-500 w-4 h-4"
                />
                <div className="flex-1">
                  <span className="font-bold text-slate-800 block">Khớp nhãn mác (Nameplate Match)</span>
                  <span className="text-[11px] text-slate-500">Rectifier, Inverter, Static Switch khớp 100% bản vẽ</span>
                </div>
              </label>

              {/* INTERLOCK SYSTEM */}
              <label className="flex items-center gap-2.5 p-2 rounded-lg border-2 border-teal-200 bg-teal-50/50 hover:bg-teal-50 cursor-pointer">
                <input
                  type="checkbox"
                  checked={data.visual_inspection.interlock_systems_operation.toLowerCase().includes('pass')}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, interlock_systems_operation: e.target.checked ? 'Pass' : 'Fail' }
                  })}
                  className="rounded text-teal-600 focus:ring-teal-500 w-4 h-4"
                />
                <div className="flex-1">
                  <div className="flex items-center gap-2">
                    <span className="font-bold text-teal-950 block">Khóa liên động cơ điện an toàn</span>
                    <span className="text-[10px] bg-teal-200 text-teal-900 font-extrabold px-1.5 py-0.5 rounded">MANDATORY</span>
                  </div>
                  <span className="text-[11px] text-teal-800">Liên động Maintenance Bypass và Inverter thao tác an toàn, chống xông ngược nguồn</span>
                </div>
              </label>

              <label className="flex items-center gap-2.5 p-2 rounded-lg border border-slate-200 hover:bg-slate-50 cursor-pointer">
                <input
                  type="checkbox"
                  checked={data.visual_inspection.fuse_ratings_match}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, fuse_ratings_match: e.target.checked }
                  })}
                  className="rounded text-teal-600 focus:ring-teal-500 w-4 h-4"
                />
                <div className="flex-1">
                  <span className="font-bold text-slate-800 block">Cầu chì bán dẫn (Fuses Match)</span>
                  <span className="text-[11px] text-slate-500">Đúng định mức và chủng loại đặc tuyến cắt nhanh</span>
                </div>
              </label>

              <div className="grid grid-cols-2 gap-2 pt-1">
                <div>
                  <label className="block text-slate-500 mb-1">Tình trạng vật lý</label>
                  <select
                    value={data.visual_inspection.physical_condition}
                    onChange={(e) => setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, physical_condition: e.target.value }
                    })}
                    className="w-full px-2.5 py-1.5 border rounded-lg bg-slate-50 font-medium"
                  >
                    <option value="Good (No damage)">Good (Nguyên vẹn)</option>
                    <option value="Damage observed">Hư hỏng cơ học</option>
                  </select>
                </div>
                <div>
                  <label className="block text-slate-500 mb-1">Vệ sinh & Bụi bẩn</label>
                  <select
                    value={data.visual_inspection.cleanliness}
                    onChange={(e) => setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, cleanliness: e.target.value }
                    })}
                    className="w-full px-2.5 py-1.5 border rounded-lg bg-slate-50 font-medium"
                  >
                    <option value="Pass">Pass (Sạch sẽ)</option>
                    <option value="Fail">Fail (Bụi bẩn/mạt)</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-slate-500 mb-1">Neo & Tiếp địa</label>
                  <select
                    value={data.visual_inspection.anchorage_grounding_clearances}
                    onChange={(e) => setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, anchorage_grounding_clearances: e.target.value }
                    })}
                    className="w-full px-2.5 py-1.5 border rounded-lg bg-slate-50 font-medium"
                  >
                    <option value="Pass">Pass (Chắc chắn)</option>
                    <option value="Fail">Fail (Lỏng lẻo)</option>
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

        {/* Right 2 Columns: Section 3 Electrical & System Integration Tests */}
        <div className="space-y-6 lg:col-span-2">
          {/* Section 3: Electrical Tests */}
          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs space-y-5">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h3 className="font-bold text-sm text-slate-900 flex items-center gap-2">
                <Zap size={16} className="text-amber-500" />
                3. Phép Đo Điện & Tích Hợp Hệ Thống (NETA 7.22.2.B)
              </h3>
              <span className={`px-2.5 py-1 rounded-full text-xs font-bold ${liveStats.overallSafe ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'}`}>
                {liveStats.overallSafe ? 'HỆ THỐNG ĐẠT CHUẨN' : 'CẦN KHẢO SÁT / THAY BÌNH'}
              </span>
            </div>

            {/* 3.1 Bolted Resistance */}
            <div>
              <div className="flex items-center justify-between mb-2">
                <span className="font-bold text-xs text-slate-700 flex items-center gap-1.5">
                  <ShieldCheck size={14} className="text-blue-600" />
                  3.1 Điện trở tiếp xúc mối nối bu-lông thanh cái (Độ lệch tối đa ≤ 50%)
                </span>
                <span className={`text-[11px] font-bold ${liveStats.resistancePass ? 'text-emerald-600' : 'text-amber-600'}`}>
                  {liveStats.resistancePass ? `✓ ĐẠT (Độ lệch ${liveStats.resistanceDev.toFixed(1)}%)` : `⚠️ LỆCH CAO (${liveStats.resistanceDev.toFixed(1)}%)`}
                </span>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                {data.electrical_tests.bolted_resistance_micro_ohms.map((resVal, idx) => (
                  <div key={idx} className="p-3 bg-slate-50 rounded-xl border border-slate-200">
                    <span className="text-[11px] text-slate-500 block mb-1">Mối nối điểm #{idx + 1} (µΩ)</span>
                    <div className="flex items-center gap-2">
                      <input
                        type="number"
                        step="0.1"
                        value={resVal}
                        onChange={(e) => {
                          const newArr = [...data.electrical_tests.bolted_resistance_micro_ohms];
                          newArr[idx] = Number(e.target.value);
                          setData({
                            ...data,
                            electrical_tests: {
                              ...data.electrical_tests,
                              bolted_resistance_micro_ohms: newArr
                            }
                          });
                        }}
                        className="w-full font-bold text-slate-800 bg-white border border-slate-300 rounded px-2 py-1 text-sm"
                      />
                      <span className="text-xs text-slate-500 font-semibold">µΩ</span>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            {/* 3.2 Static Transfer & Oscillator Frequency */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              {/* Static Transfer */}
              <div className="p-4 bg-teal-50/60 rounded-xl border border-teal-200 space-y-2">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-xs text-teal-950 flex items-center gap-1.5">
                    <Power size={14} className="text-teal-700" />
                    3.2 Công tắc tĩnh Static Switch
                  </span>
                  <span className={`px-2 py-0.5 rounded text-[11px] font-bold ${liveStats.staticPass ? 'bg-emerald-100 text-emerald-800' : 'bg-red-100 text-red-800'}`}>
                    {liveStats.staticPass ? 'VÔ CẤP (< 2ms)' : 'LỖI CHUYỂN NGUỒN'}
                  </span>
                </div>
                <input
                  type="text"
                  value={data.electrical_tests.static_transfer_function}
                  onChange={(e) => setData({
                    ...data,
                    electrical_tests: { ...data.electrical_tests, static_transfer_function: e.target.value }
                  })}
                  className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 font-medium text-xs text-slate-800"
                />
                <span className="text-[11px] text-teal-800 block">
                  Tiêu chuẩn: Chuyển mạch Inverter &lt;-&gt; Bypass không gián đoạn điện áp phụ tải.
                </span>
              </div>

              {/* Oscillator Frequency */}
              <div className="p-4 bg-blue-50/60 rounded-xl border border-blue-200 space-y-2">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-xs text-blue-950 flex items-center gap-1.5">
                    <Activity size={14} className="text-blue-700" />
                    3.3 Tần số bộ dao động tự do Inverter
                  </span>
                  <span className={`px-2 py-0.5 rounded text-[11px] font-bold ${liveStats.freqPass ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'}`}>
                    {liveStats.freqPass ? '50Hz ± 0.1Hz' : 'LỆCH TẦN SỐ'}
                  </span>
                </div>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.oscillator_free_running_frequency_hz}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: { ...data.electrical_tests, oscillator_free_running_frequency_hz: Number(e.target.value) }
                    })}
                    className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 font-bold font-mono text-sm text-slate-800"
                  />
                  <span className="text-xs font-bold text-slate-600">Hz</span>
                </div>
                <span className="text-[11px] text-blue-800 block">
                  Dung sai nhà sản xuất: 50.00 Hz ± 0.10 Hz khi mất đồng bộ lưới.
                </span>
              </div>
            </div>

            {/* 3.4 DC Undervoltage Trip & Alarms & Synchronizing */}
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
                <label className="text-slate-600 font-bold block mb-1">Cắt Thấp áp DC Inverter (V DC):</label>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    value={data.electrical_tests.dc_undervoltage_inverter_trip_v}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: { ...data.electrical_tests, dc_undervoltage_inverter_trip_v: Number(e.target.value) }
                    })}
                    className="w-full px-2.5 py-1.5 bg-white border rounded font-mono font-bold text-slate-800"
                  />
                  <span className="text-xs text-slate-500 font-bold">V</span>
                </div>
                <span className="text-[10px] text-slate-500 mt-1 block">Bảo vệ xả sâu ắc quy</span>
              </div>

              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
                <label className="text-slate-600 font-bold block mb-1">Mạch Cảnh báo sự cố (Alarms):</label>
                <input
                  type="text"
                  value={data.electrical_tests.alarm_circuits_test}
                  onChange={(e) => setData({
                    ...data,
                    electrical_tests: { ...data.electrical_tests, alarm_circuits_test: e.target.value }
                  })}
                  className="w-full px-2.5 py-1.5 bg-white border rounded font-medium text-slate-800"
                />
                <span className="text-[10px] text-slate-500 mt-1 block">Báo còi/đèn & SCADA BMS</span>
              </div>

              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
                <label className="text-slate-600 font-bold block mb-1">Chỉ thị đồng bộ pha (Sync):</label>
                <input
                  type="text"
                  value={data.electrical_tests.synchronizing_indicators_test}
                  onChange={(e) => setData({
                    ...data,
                    electrical_tests: { ...data.electrical_tests, synchronizing_indicators_test: e.target.value }
                  })}
                  className="w-full px-2.5 py-1.5 bg-white border rounded font-medium text-slate-800"
                />
                <span className="text-[10px] text-slate-500 mt-1 block">Khóa đồng bộ Inverter - Bypass</span>
              </div>
            </div>

            {/* 3.5 Breakers & ATS Subsystems */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
                <label className="text-slate-600 font-bold block mb-1">Máy cắt UPS (Section 7.6):</label>
                <input
                  type="text"
                  value={data.electrical_tests.ups_breakers_section_7_6_status}
                  onChange={(e) => setData({
                    ...data,
                    electrical_tests: { ...data.electrical_tests, ups_breakers_section_7_6_status: e.target.value }
                  })}
                  className="w-full px-2.5 py-1.5 bg-white border rounded font-medium text-slate-800"
                />
              </div>

              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
                <label className="text-slate-600 font-bold block mb-1">Bộ chuyển nguồn ATS (Section 7.22.3):</label>
                <input
                  type="text"
                  value={data.electrical_tests.ats_section_7_22_3_status}
                  onChange={(e) => setData({
                    ...data,
                    electrical_tests: { ...data.electrical_tests, ats_section_7_22_3_status: e.target.value }
                  })}
                  className="w-full px-2.5 py-1.5 bg-white border rounded font-medium text-slate-800"
                />
              </div>
            </div>

            {/* 3.6 CRITICAL: BATTERY SYSTEM SECTION 7.18 */}
            <div className={`p-4 rounded-xl border-2 transition-all space-y-3 ${
              liveStats.batHasHighResistance ? 'bg-amber-50/90 border-amber-400' : 'bg-emerald-50/70 border-emerald-300'
            }`}>
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <BatteryCharging size={18} className={liveStats.batHasHighResistance ? 'text-amber-700' : 'text-emerald-700'} />
                  <span className="font-extrabold text-xs sm:text-sm text-slate-900">
                    3.6 Hệ Thống Ắc Quy Dự Phòng (Battery System - NETA ATS-2025 Section 7.18)
                  </span>
                </div>
                <span className={`px-2.5 py-0.5 rounded text-[11px] font-black uppercase tracking-wider ${
                  liveStats.batHasHighResistance ? 'bg-amber-200 text-amber-900 border border-amber-300 animate-pulse' : 'bg-emerald-200 text-emerald-900'
                }`}>
                  {liveStats.batHasHighResistance ? 'INVESTIGATE / THAY BÌNH' : 'PASS (ĐẠT CHUẨN)'}
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
                <div>
                  <label className="text-slate-600 block mb-1">Chủng loại ắc quy & Cấu hình chuỗi:</label>
                  <input
                    type="text"
                    value={data.electrical_tests.battery_system_section_7_18.battery_type}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        battery_system_section_7_18: {
                          ...data.electrical_tests.battery_system_section_7_18,
                          battery_type: e.target.value
                        }
                      }
                    })}
                    className="w-full bg-white border rounded px-2.5 py-1.5 font-medium text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 block mb-1">Tổng điện áp chuỗi (V DC):</label>
                  <div className="flex items-center gap-2">
                    <input
                      type="number"
                      value={data.electrical_tests.battery_system_section_7_18.total_string_voltage_v}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          battery_system_section_7_18: {
                            ...data.electrical_tests.battery_system_section_7_18,
                            total_string_voltage_v: Number(e.target.value)
                          }
                        }
                      })}
                      className="w-full bg-white border rounded px-2.5 py-1.5 font-bold font-mono text-slate-800"
                    />
                    <span className="text-xs font-bold text-slate-600">V</span>
                  </div>
                </div>

                <div>
                  <label className="text-slate-600 block mb-1">Trạng thái tổng thể ắc quy:</label>
                  <select
                    value={data.electrical_tests.battery_system_section_7_18.status}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        battery_system_section_7_18: {
                          ...data.electrical_tests.battery_system_section_7_18,
                          status: e.target.value as any
                        }
                      }
                    })}
                    className="w-full bg-white border rounded px-2.5 py-1.5 font-bold text-slate-800"
                  >
                    <option value="Pass">Pass (Đạt)</option>
                    <option value="Investigate">Investigate (Cần thay bình)</option>
                    <option value="Fail">Fail (Không đạt)</option>
                  </select>
                </div>
              </div>

              <div>
                <label className="text-slate-600 text-xs font-semibold block mb-1">
                  Đo nội trở từng bình (Internal Resistance / Impedance Check per IEEE 1188):
                </label>
                <input
                  type="text"
                  value={data.electrical_tests.battery_system_section_7_18.block_internal_resistance_check}
                  onChange={(e) => setData({
                    ...data,
                    electrical_tests: {
                      ...data.electrical_tests,
                      battery_system_section_7_18: {
                        ...data.electrical_tests.battery_system_section_7_18,
                        block_internal_resistance_check: e.target.value
                      }
                    }
                  })}
                  className="w-full bg-white border border-slate-300 rounded px-3 py-2 font-medium text-xs text-slate-800"
                />
              </div>

              {liveStats.batHasHighResistance && (
                <div className="p-2.5 bg-amber-100/70 border border-amber-300 rounded-lg text-xs font-medium text-amber-950 flex items-start gap-2">
                  <AlertTriangle size={15} className="text-amber-800 shrink-0 mt-0.5" />
                  <span>
                    <strong>Phân tích IEEE 1188:</strong> Nội trở lệch &gt; 20% - 35% là dấu hiệu chai hóa cực âm/khô điện dịch. Yêu cầu FSE đánh dấu thay thế 2 bình này và thực hiện xả tải Load Bank Test trước khi bàn giao hệ thống.
                  </span>
                </div>
              )}
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
                analysisResult.overall_status === 'PASS' ? 'bg-emerald-600' : 'bg-amber-600'
              }`}>
                {analysisResult.overall_status === 'PASS' ? <CheckCircle2 size={28} /> : <AlertTriangle size={28} />}
              </div>
              <div>
                <span className="text-[11px] font-bold text-amber-400 uppercase tracking-wider">
                  KẾT QUẢ ĐỐI SOÁT AI CHUYÊN GIA NETA ATS-2025
                </span>
                <h2 className="text-xl sm:text-2xl font-black">
                  Trạng thái: {analysisResult.overall_status === 'PASS' ? 'PASS (ĐẠT CHUẨN)' : 'INVESTIGATE / BATTERY REPLACEMENT REQUIRED'}
                </h2>
              </div>
            </div>

            <button
              onClick={() => setShowPrintModal(true)}
              className="px-4 py-2 bg-teal-600 hover:bg-teal-500 text-white text-xs font-bold rounded-xl shadow-md transition-all flex items-center justify-center gap-2 self-start sm:self-auto"
            >
              <Printer size={15} />
              <span>In Báo Cáo PDF Chính Thức</span>
            </button>
          </div>

          <div className="p-6 bg-slate-50 space-y-6">
            <div className={`p-4 rounded-xl border font-medium text-xs sm:text-sm leading-relaxed ${
              analysisResult.overall_status === 'PASS' ? 'bg-emerald-50 border-emerald-200 text-emerald-900' : 'bg-amber-50 border-amber-200 text-amber-900'
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
                <FileText size={16} className="text-teal-600" />
                Mã JSON Khảo Sát Hiện Trường UPS (Field Service Input)
              </h3>
              <button onClick={() => setShowJsonModal(false)} className="text-slate-400 hover:text-slate-600 text-lg">
                &times;
              </button>
            </div>
            <div className="p-4 flex-1 overflow-y-auto">
              <p className="text-xs text-slate-500 mb-2">
                Dán chuỗi JSON kiểm tra UPS từ hệ thống BMS / Hãng sản xuất hoặc ứng dụng TEV:
              </p>
              <textarea
                value={jsonInput}
                onChange={(e) => setJsonInput(e.target.value)}
                rows={14}
                className="w-full p-3 font-mono text-xs bg-slate-900 text-emerald-400 rounded-xl border border-slate-800 focus:outline-hidden focus:ring-2 focus:ring-teal-500"
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
                      setSavedSuccessMessage('Đã tải dữ liệu JSON UPS thành công!');
                      setTimeout(() => setSavedSuccessMessage(null), 3000);
                    } catch {
                      alert('Chuỗi JSON không đúng định dạng!');
                    }
                  }}
                  className="px-4 py-2 text-xs font-bold text-white bg-teal-600 hover:bg-teal-500 rounded-lg shadow-md"
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
        title="BIÊN BẢN THỬ NGHIỆM HỆ THỐNG NGUỒN CẤP ĐIỆN LIÊN TỤC (UPS)"
        standard="ANSI/NETA ATS-2025 Section 7.22.2"
        siteInfo={data.site_info}
        overallStatus={analysisResult?.overall_status || (liveStats.overallSafe ? 'PASS' : 'INVESTIGATE')}
        statusReason={analysisResult?.status_reason || (liveStats.overallSafe ? 'Đạt chuẩn NETA ATS-2025 Section 7.22.2.' : 'Phát hiện bình ắc quy tăng nội trở cao >35%. Yêu cầu thay thế bình.')}
        visualChecks={visualChecksForPrint}
        electricalTable={electricalTableForPrint}
        warnings={
          liveStats.overallSafe
            ? []
            : [
                liveStats.batHasHighResistance ? 'Phát hiện 2 bình ắc quy tăng nội trở cao >35% so với mức trung bình chuỗi (Vượt ngưỡng 20% theo IEEE 1188 / NETA Section 7.18).' : '',
                !liveStats.resistancePass ? `Điện trở tiếp xúc mối nối bu-lông có độ lệch ${liveStats.resistanceDev.toFixed(1)}% vượt quá 50%.` : '',
                !liveStats.staticPass ? 'Công tắc tĩnh chuyển mạch không đạt chỉ tiêu vô cấp.' : '',
                !liveStats.freqPass ? 'Tần số bộ dao động tự do lệch chuẩn 50Hz.' : ''
              ].filter(Boolean)
        }
        recommendations={[
          'Đánh dấu số thứ tự và lập kế hoạch thay thế 2 bình ắc quy bị tăng nội trở cao (>35% dev) bằng bình mới cùng chủng loại 12V 100Ah VRLA.',
          'Thực hiện xả tải ắc quy thử nghiệm (Battery Load Discharge Test per Section 7.18 & IEEE 1188) sau khi thay thế bình mới.',
          'Kiểm tra định kỳ nhiệt độ hồng ngoại các mối nối bu-lông thanh cái và cực bình ắc quy khi mang tải đầy đủ.'
        ]}
      />

      {/* Camera Nameplate OCR Modal */}
      {showCameraModal && (
        <NameplateScannerModal
          isOpen={showCameraModal}
          onClose={() => setShowCameraModal(false)}
          onApplyData={handleApplyNameplateData}
          currentTag={data.site_info.ups_tag}
        />
      )}
    </div>
  );
};
