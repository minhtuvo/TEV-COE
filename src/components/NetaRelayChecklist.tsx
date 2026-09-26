import React, { useState, useMemo } from 'react';
import {
  Zap,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  FileText,
  Printer,
  Sparkles,
  Save,
  RotateCcw,
  Info,
  ShieldCheck,
  Compass,
  Cpu,
  Clock,
  Radio,
  Sliders,
  Check,
  Eye,
  Camera,
  Upload,
  Download
} from 'lucide-react';
import {
  NetaRelayInputPayload,
  NetaRelayEvaluationResult,
  performDeterministicRelayAnalysis
} from '../../server/netaRelayAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';
import { NameplateScannerModal } from './NameplateScannerModal';
import { ExtractedNameplateData } from '../../server/nameplateOcrAnalyzer';
import { FseEngineerDropdown } from './FseEngineerDropdown';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

export const SAMPLE_RELAY_FSE_DATA: NetaRelayInputPayload = {
  site_info: {
    project_name: "Tram 110kV TEV Substation Feeder Bay",
    relay_tag: "F01-RELAY-SEL751",
    relay_model: "SEL-751A Feeder Protection Relay",
    firmware_version: "R115-V0-Z001002",
    control_voltage_vdc: "110V DC",
    fse_name: "Võ Minh Tú (sgm1707@gmail.com)",
    test_date: "2026-09-23"
  },
  visual_inspection: {
    identification_match: true,
    display_and_leds_test: "Pass",
    cleanliness_and_connections_tight: true,
    frame_grounding_check: "Pass",
    settings_match_coordination_study: true,
    clock_date_verified: "2026-09-23 10:15:00 (Pass)",
    ct_shorting_blocks_check: "Pass"
  },
  electrical_tests: {
    insulation_resistance_1000v_megohms: {
      ac_current_inputs_to_ground: 250,
      ac_voltage_inputs_to_ground: 220,
      dc_power_supply_to_ground: 180
    },
    analog_metering_accuracy: {
      injected_ia_5_00a: "Relay displays 5.01A (Pass)",
      injected_va_66_4v: "Relay displays 66.3V (Pass)"
    },
    protection_elements_test: {
      ansi_51_time_overcurrent: {
        setting_pickup_a: 5.0,
        measured_pickup_a: 5.02,
        setting_time_dial: 1.0,
        expected_trip_time_sec_at_3x: 1.35,
        measured_trip_time_sec_at_3x: 1.36,
        status: "Pass"
      },
      ansi_50_instantaneous_overcurrent: {
        setting_pickup_a: 25.0,
        measured_pickup_a: 28.5,
        expected_trip_time_sec: 0.02,
        measured_trip_time_sec: 0.021,
        status: "Fail (Pickup error +14% exceeds +/-5% tolerance)"
      }
    },
    digital_inputs_test: "Pass (All 8 active DIs verified)",
    output_contacts_trip_test: "Pass (Breaker tripped successfully via OUT101)",
    internal_logic_verification: "Pass",
    trip_coil_monitoring_tcs: "Pass",
    arc_energy_reduction_sensor: "Pass (Optical sensor threshold verified above ambient)",
    event_records_cleared_after_test: false
  },
  previous_test_data: {
    last_test_date: "2025-09-21",
    last_ansi_50_pickup_a: 25.1
  }
};

export const NetaRelayChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaRelayInputPayload>(SAMPLE_RELAY_FSE_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaRelayEvaluationResult | null>(null);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showCameraModal, setShowCameraModal] = useState(false);
  const [jsonInput, setJsonInput] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live Calculations based on NETA ATS-2025 Section 7.9.2
  const liveStats = useMemo(() => {
    // 1. Insulation Resistance (>= 100 MΩ)
    const ir = data.electrical_tests.insulation_resistance_1000v_megohms;
    const irCurrentPass = ir.ac_current_inputs_to_ground >= 100;
    const irVoltagePass = ir.ac_voltage_inputs_to_ground >= 100;
    const irPowerPass = ir.dc_power_supply_to_ground >= 100;
    const irAllPass = irCurrentPass && irVoltagePass && irPowerPass;

    // 2. ANSI 51 Time-Overcurrent
    const ansi51 = data.electrical_tests.protection_elements_test.ansi_51_time_overcurrent;
    const ansi51PickupErr = ansi51.setting_pickup_a > 0
      ? ((ansi51.measured_pickup_a - ansi51.setting_pickup_a) / ansi51.setting_pickup_a) * 100
      : 0;
    const ansi51TimeErr = ansi51.expected_trip_time_sec_at_3x > 0
      ? ((ansi51.measured_trip_time_sec_at_3x - ansi51.expected_trip_time_sec_at_3x) / ansi51.expected_trip_time_sec_at_3x) * 100
      : 0;
    const ansi51Pass = Math.abs(ansi51PickupErr) <= 5.0 && Math.abs(ansi51TimeErr) <= 5.0;

    // 3. ANSI 50 Instantaneous Overcurrent (Critical +/-5% tolerance)
    const ansi50 = data.electrical_tests.protection_elements_test.ansi_50_instantaneous_overcurrent;
    const ansi50PickupErr = ansi50.setting_pickup_a > 0
      ? ((ansi50.measured_pickup_a - ansi50.setting_pickup_a) / ansi50.setting_pickup_a) * 100
      : 0;
    const ansi50Pass = Math.abs(ansi50PickupErr) <= 5.0 && ansi50.measured_trip_time_sec <= 0.035;

    // 4. Mandatory Event Records Cleared
    const recordsCleared = data.electrical_tests.event_records_cleared_after_test;

    // 5. Visual & Settings Match
    const visualPass = data.visual_inspection.identification_match &&
      data.visual_inspection.settings_match_coordination_study &&
      data.visual_inspection.cleanliness_and_connections_tight;

    const overallSafe = irAllPass && ansi51Pass && ansi50Pass && recordsCleared && visualPass;

    return {
      irAllPass,
      ansi51PickupErr,
      ansi51TimeErr,
      ansi51Pass,
      ansi50PickupErr,
      ansi50Pass,
      recordsCleared,
      visualPass,
      overallSafe
    };
  }, [data]);

  const handleApplyNameplateData = (extracted: ExtractedNameplateData) => {
    setData((prev) => ({
      ...prev,
      site_info: {
        ...prev.site_info,
        relay_tag: extracted.equipment_tag || prev.site_info.relay_tag,
        relay_model: extracted.equipment_name || extracted.equipment_type || prev.site_info.relay_model,
      },
      visual_inspection: {
        ...prev.visual_inspection,
        identification_match: true,
      }
    }));
    setSavedSuccessMessage(`Đã quét nhãn thành công: ${extracted.equipment_tag} (${extracted.equipment_name || 'Rơ-le'}) - Tự động cập nhật vào Checklist!`);
    setTimeout(() => setSavedSuccessMessage(null), 5000);
  };

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    setAnalysisResult(null);
    try {
      const res = await fetch('/api/field-service/neta-relay-analysis', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });
      const json = await res.json();
      if (json.success && json.data) {
        setAnalysisResult(json.data);
      } else {
        // Fallback to client-side deterministic evaluation
        const det = performDeterministicRelayAnalysis(data);
        setAnalysisResult(det);
      }
    } catch {
      const det = performDeterministicRelayAnalysis(data);
      setAnalysisResult(det);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleSaveToCmms = () => {
    const reportPayload = {
      type: 'NETA_ATS_2025_RELAY',
      data: data,
      analysisResult: analysisResult,
      equipmentId: data.site_info.relay_tag,
      date: data.site_info.test_date,
      title: `Biên bản Rơ-le Số NETA ATS-2025: ${data.site_info.relay_tag}`,
      standard: 'ANSI/NETA ATS-2025 Section 7.9.2',
      status: liveStats.overallSafe ? 'PASS' : 'FAIL',
      timestamp: new Date().toISOString()
    };

    if (onSaveReport) {
      onSaveReport(reportPayload);
      setSavedSuccessMessage(`Đã lưu biên bản kiểm định rơ-le ${data.site_info.relay_tag} vào kho báo cáo thành công!`);
      setTimeout(() => setSavedSuccessMessage(null), 5000);
    }
  };

  const electricalTableForPrint = [
    {
      item: 'Điện trở cách điện 1000V DC (Mạch dòng AC đến đất)',
      measured: `${data.electrical_tests.insulation_resistance_1000v_megohms.ac_current_inputs_to_ground} MΩ`,
      standard: 'Tối thiểu ≥ 100 MΩ',
      reference: 'NETA ATS-2025 Sec 7.9.2.B.1',
      deviation: '0%',
      status: data.electrical_tests.insulation_resistance_1000v_megohms.ac_current_inputs_to_ground >= 100 ? 'PASS' : 'FAIL'
    },
    {
      item: 'Điện trở cách điện 1000V DC (Mạch áp AC đến đất)',
      measured: `${data.electrical_tests.insulation_resistance_1000v_megohms.ac_voltage_inputs_to_ground} MΩ`,
      standard: 'Tối thiểu ≥ 100 MΩ',
      reference: 'NETA ATS-2025 Sec 7.9.2.B.1',
      deviation: '0%',
      status: data.electrical_tests.insulation_resistance_1000v_megohms.ac_voltage_inputs_to_ground >= 100 ? 'PASS' : 'FAIL'
    },
    {
      item: 'Điện trở cách điện 1000V DC (Nguồn DC đến đất)',
      measured: `${data.electrical_tests.insulation_resistance_1000v_megohms.dc_power_supply_to_ground} MΩ`,
      standard: 'Tối thiểu ≥ 100 MΩ',
      reference: 'NETA ATS-2025 Sec 7.9.2.B.1',
      deviation: '0%',
      status: data.electrical_tests.insulation_resistance_1000v_megohms.dc_power_supply_to_ground >= 100 ? 'PASS' : 'FAIL'
    },
    {
      item: 'Đo lường tương tự (Analog Metering IA & VA)',
      measured: `IA: ${data.electrical_tests.analog_metering_accuracy.injected_ia_5_00a} | VA: ${data.electrical_tests.analog_metering_accuracy.injected_va_66_4v}`,
      standard: 'Sai số đo lường ≤ ±0.5% dải đo',
      reference: 'NETA ATS-2025 Sec 7.9.2.B.2',
      deviation: '≤ 0.2%',
      status: 'PASS'
    },
    {
      item: 'Quá dòng có thời gian (ANSI 51 Pickup)',
      measured: `${data.electrical_tests.protection_elements_test.ansi_51_time_overcurrent.measured_pickup_a} A (Cài: ${data.electrical_tests.protection_elements_test.ansi_51_time_overcurrent.setting_pickup_a} A)`,
      standard: 'Dung sai khởi động ≤ ±5.0%',
      reference: 'NETA ATS-2025 Sec 7.9.2.B.3.a',
      deviation: `${liveStats.ansi51PickupErr.toFixed(2)}%`,
      status: Math.abs(liveStats.ansi51PickupErr) <= 5.0 ? 'PASS' : 'FAIL'
    },
    {
      item: 'Quá dòng có thời gian (ANSI 51 Trip Time @3xI)',
      measured: `${data.electrical_tests.protection_elements_test.ansi_51_time_overcurrent.measured_trip_time_sec_at_3x} s (Lý thuyết: ${data.electrical_tests.protection_elements_test.ansi_51_time_overcurrent.expected_trip_time_sec_at_3x} s)`,
      standard: 'Dung sai thời gian ≤ ±5.0%',
      reference: 'NETA ATS-2025 Sec 7.9.2.B.3.a',
      deviation: `${liveStats.ansi51TimeErr.toFixed(2)}%`,
      status: Math.abs(liveStats.ansi51TimeErr) <= 5.0 ? 'PASS' : 'FAIL'
    },
    {
      item: 'Quá dòng cắt nhanh tức thời (ANSI 50 Instantaneous Pickup)',
      measured: `${data.electrical_tests.protection_elements_test.ansi_50_instantaneous_overcurrent.measured_pickup_a} A (Cài: ${data.electrical_tests.protection_elements_test.ansi_50_instantaneous_overcurrent.setting_pickup_a} A)`,
      standard: 'Dung sai khởi động ≤ ±5.0%',
      reference: 'NETA ATS-2025 Sec 7.9.2.B.3.b',
      deviation: `${liveStats.ansi50PickupErr > 0 ? '+' : ''}${liveStats.ansi50PickupErr.toFixed(1)}% (Vượt dung sai)`,
      status: liveStats.ansi50Pass ? 'PASS' : 'FAIL'
    },
    {
      item: 'Thử nghiệm Trip máy cắt qua ngõ ra DO',
      measured: data.electrical_tests.output_contacts_trip_test,
      standard: 'Tiếp điểm lực đóng ngắt dứt khoát, máy cắt trip tin cậy',
      reference: 'NETA ATS-2025 Sec 7.9.2.B.4',
      deviation: '0%',
      status: 'PASS'
    },
    {
      item: 'Giám sát mạch cắt (Trip Coil Monitoring - TCS)',
      measured: data.electrical_tests.trip_coil_monitoring_tcs,
      standard: 'Báo động chính xác khi đứt/hở mạch cuộn cắt máy cắt',
      reference: 'NETA ATS-2025 Sec 7.9.2.B.6',
      deviation: 'Khớp',
      status: 'PASS'
    },
    {
      item: 'Xóa nhật ký sự cố giả sau thử nghiệm (Event Records Cleared)',
      measured: data.electrical_tests.event_records_cleared_after_test ? 'ĐÃ XÓA SẠCH (Pass)' : 'CHƯA XÓA (Vi phạm quy trình)',
      standard: 'BẮT BUỘC xóa sạch Fault logs, SOE, Event records',
      reference: 'NETA ATS-2025 Sec 7.9.2.B.8',
      deviation: data.electrical_tests.event_records_cleared_after_test ? '0%' : 'Chưa thực hiện',
      status: data.electrical_tests.event_records_cleared_after_test ? 'PASS' : 'FAIL'
    }
  ];

  const visualChecksForPrint = [
    { label: 'Nhận diện & Bản vẽ thiết kế (Identification)', value: data.visual_inspection.identification_match ? 'ĐẠT' : 'KHÔNG ĐẠT', pass: data.visual_inspection.identification_match },
    { label: 'Kiểm tra màn hình & đèn LED chỉ thị (Display & LEDs)', value: data.visual_inspection.display_and_leds_test, pass: data.visual_inspection.display_and_leds_test.toLowerCase().includes('pass') },
    { label: 'Vệ sinh & Siết chặt ốc đấu nối (Cleanliness & Terminals)', value: data.visual_inspection.cleanliness_and_connections_tight ? 'ĐẠT' : 'CHƯA ĐẠT', pass: data.visual_inspection.cleanliness_and_connections_tight },
    { label: 'Tiếp địa vỏ khung rơ-le (Frame Grounding)', value: data.visual_inspection.frame_grounding_check, pass: data.visual_inspection.frame_grounding_check.toLowerCase().includes('pass') },
    { label: 'Đối soát file cài đặt khớp 100% Phiếu chỉnh định', value: data.visual_inspection.settings_match_coordination_study ? 'ĐẠT' : 'KHÔNG ĐẠT', pass: data.visual_inspection.settings_match_coordination_study },
    { label: 'Đồng hồ thời gian thực (RTC Clock & Date Sync)', value: data.visual_inspection.clock_date_verified, pass: data.visual_inspection.clock_date_verified.toLowerCase().includes('pass') },
    { label: 'Cơ cấu ngắn mạch mạch dòng CT (CT Shorting Blocks)', value: data.visual_inspection.ct_shorting_blocks_check, pass: data.visual_inspection.ct_shorting_blocks_check.toLowerCase().includes('pass') }
  ];

  return (
    <div className="space-y-6">
      {/* Top Banner & Action Header */}
      <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-blue-900 text-white p-5 sm:p-6 rounded-2xl shadow-lg border border-slate-700/50">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2 mb-1.5">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-400/20 text-amber-300 border border-amber-400/30 uppercase tracking-wide">
                ANSI/NETA ATS-2025 Section 7.9.2
              </span>
              <span className="text-xs text-slate-300">Rơ-le Số • SEL, ABB, Siemens, MiCOM, GE</span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white flex items-center gap-2.5">
              <Cpu className="text-amber-400 fill-amber-400/20" size={26} />
              Checklist Rơ-le Bảo Vệ Vi Xử Lý (Protective Relays)
            </h1>
            <p className="text-xs sm:text-sm text-slate-300 mt-1 max-w-2xl">
              Quy chuẩn kiểm tra toàn diện: Đối soát file cài đặt 100%, bảo vệ ANSI 50/51/87/21, đo lường tương tự, logic trip cuộn cắt, giám sát TCS và bắt buộc xóa sạch nhật ký sự cố.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => setShowCameraModal(true)}
              className="px-3.5 py-2 rounded-xl text-xs font-bold bg-amber-400 hover:bg-amber-300 text-slate-950 flex items-center gap-1.5 transition-all shadow-md active:scale-95"
              title="Quét nhãn Nameplate bằng Camera AI"
            >
              <Camera size={14} />
              <span>Quét Nameplate (AI OCR)</span>
            </button>

            <button
              onClick={() => {
                setData(SAMPLE_RELAY_FSE_DATA);
                setAnalysisResult(null);
              }}
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title="Tải kịch bản mẫu FSE SEL-751A"
            >
              <RotateCcw size={14} />
              <span>Nạp Kịch Bản Mẫu FSE</span>
            </button>

            <button
              onClick={() => setShowPrintModal(true)}
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-indigo-600 hover:bg-indigo-500 text-white border border-indigo-400/40 flex items-center gap-1.5 transition-all shadow-md"
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
              className="px-4 py-2 rounded-xl text-xs font-bold bg-gradient-to-r from-amber-500 to-orange-500 hover:from-amber-400 hover:to-orange-400 text-slate-950 flex items-center gap-2 transition-all shadow-lg shadow-orange-500/20 disabled:opacity-50"
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

      {/* Critical Status Quick Banner */}
      {!liveStats.overallSafe && (
        <div className="bg-red-50 border border-red-200 rounded-xl p-4 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 shadow-xs">
          <div className="flex items-start gap-3">
            <div className="w-9 h-9 rounded-xl bg-red-100 text-red-600 flex items-center justify-center shrink-0 mt-0.5">
              <AlertTriangle size={20} />
            </div>
            <div>
              <h4 className="text-xs sm:text-sm font-bold text-red-900">
                CẢNH BÁO NETA ATS-2025: PHÁT HIỆN LỖI KHÔNG ĐẠT TIÊU CHUẨN ĐÓNG ĐIỆN!
              </h4>
              <p className="text-xs text-red-700 mt-0.5">
                {!liveStats.ansi50Pass && `• Bảo vệ cắt nhanh ANSI 50 có sai số khởi động ${liveStats.ansi50PickupErr > 0 ? '+' : ''}${liveStats.ansi50PickupErr.toFixed(1)}% vượt quá dung sai cho phép ±5.0%. `}
                {!liveStats.recordsCleared && `• Chưa xóa nhật ký sự cố giả sau thử nghiệm (Vi phạm NETA 7.9.2.B.8). `}
                {!liveStats.irAllPass && `• Điện trở cách điện suy giảm < 100 MΩ. `}
              </p>
            </div>
          </div>
          <button
            onClick={handleAnalyze}
            className="px-3.5 py-1.5 bg-red-600 hover:bg-red-700 text-white rounded-lg text-xs font-bold shrink-0 transition-colors"
          >
            Xem Khuyến Nghị AI
          </button>
        </div>
      )}

      {/* Main Grid: Section 1 & Section 2 & Section 3 */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left Column: Section 1 & Section 2 */}
        <div className="space-y-6 lg:col-span-1">
          {/* Section 1: Site Info & Relay ID */}
          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-2">
              <h3 className="font-bold text-sm text-slate-900 flex items-center gap-2">
                <Compass size={16} className="text-blue-600" />
                1. Thông tin Hiện trường & Rơ-le
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
                <span>Chụp ảnh tem rơ-le để AI tự trích xuất thông số</span>
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
                <label className="block text-slate-500 mb-1">Dự án / Trạm biến áp</label>
                <input
                  type="text"
                  value={data.site_info.project_name}
                  onChange={(e) => setData({ ...data, site_info: { ...data.site_info, project_name: e.target.value } })}
                  className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white focus:ring-2 focus:ring-blue-500 font-medium text-slate-800"
                />
              </div>
              <div>
                <label className="block text-slate-500 mb-1">Mã Tag Rơ-le (Relay Tag)</label>
                <input
                  type="text"
                  value={data.site_info.relay_tag}
                  onChange={(e) => setData({ ...data, site_info: { ...data.site_info, relay_tag: e.target.value } })}
                  className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white focus:ring-2 focus:ring-blue-500 font-mono font-bold text-blue-900"
                />
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-slate-500 mb-1">Model Rơ-le</label>
                  <input
                    type="text"
                    value={data.site_info.relay_model}
                    onChange={(e) => setData({ ...data, site_info: { ...data.site_info, relay_model: e.target.value } })}
                    className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white font-medium text-slate-800"
                  />
                </div>
                <div>
                  <label className="block text-slate-500 mb-1">Firmware Version</label>
                  <input
                    type="text"
                    value={data.site_info.firmware_version}
                    onChange={(e) => setData({ ...data, site_info: { ...data.site_info, firmware_version: e.target.value } })}
                    className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white font-mono text-slate-800"
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
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-slate-500 mb-1">Nguồn nuôi DC</label>
                  <input
                    type="text"
                    value={data.site_info.control_voltage_vdc || '110V DC'}
                    onChange={(e) => setData({ ...data, site_info: { ...data.site_info, control_voltage_vdc: e.target.value } })}
                    className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white font-medium text-slate-800"
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
          </div>

          {/* Section 2: Visual & Mechanical Inspection (NETA 7.9.2.A) */}
          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs space-y-4">
            <h3 className="font-bold text-sm text-slate-900 flex items-center gap-2 border-b border-slate-100 pb-2">
              <ShieldCheck size={16} className="text-emerald-600" />
              2. Kiểm Tra Thị Giác & Cơ Khí (NETA 7.9.2.A)
            </h3>

            <div className="space-y-3 text-xs">
              <label className="flex items-center gap-2.5 p-2 rounded-lg border border-slate-200 hover:bg-slate-50 cursor-pointer">
                <input
                  type="checkbox"
                  checked={data.visual_inspection.identification_match}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, identification_match: e.target.checked }
                  })}
                  className="rounded text-blue-600 focus:ring-blue-500 w-4 h-4"
                />
                <div className="flex-1">
                  <span className="font-bold text-slate-800 block">Nhận diện khớp bản vẽ (Identification)</span>
                  <span className="text-[11px] text-slate-500">Model, Style, Serial, Firmware trùng khớp hồ sơ thiết kế</span>
                </div>
              </label>

              {/* CRITICAL SETTINGS VERIFICATION */}
              <label className="flex items-center gap-2.5 p-2 rounded-lg border-2 border-indigo-200 bg-indigo-50/50 hover:bg-indigo-50 cursor-pointer">
                <input
                  type="checkbox"
                  checked={data.visual_inspection.settings_match_coordination_study}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, settings_match_coordination_study: e.target.checked }
                  })}
                  className="rounded text-indigo-600 focus:ring-indigo-500 w-4 h-4"
                />
                <div className="flex-1">
                  <div className="flex items-center gap-2">
                    <span className="font-bold text-indigo-900 block">Đối soát file cài đặt trùng khớp 100%</span>
                    <span className="text-[10px] bg-indigo-200 text-indigo-800 font-extrabold px-1.5 py-0.5 rounded">MANDATORY</span>
                  </div>
                  <span className="text-[11px] text-indigo-700">Tải file settings & logic ra và so sánh khớp 100% phiếu chỉnh định (Coordination study)</span>
                </div>
              </label>

              <label className="flex items-center gap-2.5 p-2 rounded-lg border border-slate-200 hover:bg-slate-50 cursor-pointer">
                <input
                  type="checkbox"
                  checked={data.visual_inspection.cleanliness_and_connections_tight}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, cleanliness_and_connections_tight: e.target.checked }
                  })}
                  className="rounded text-blue-600 focus:ring-blue-500 w-4 h-4"
                />
                <div className="flex-1">
                  <span className="font-bold text-slate-800 block">Vệ sinh & Siết ốc đầu nối</span>
                  <span className="text-[11px] text-slate-500">Rơ-le sạch bụi bẩn, các vít cầu đấu siết chặt</span>
                </div>
              </label>

              <div className="grid grid-cols-2 gap-2 pt-1">
                <div>
                  <label className="block text-slate-500 mb-1">Màn hình & Đèn LED</label>
                  <select
                    value={data.visual_inspection.display_and_leds_test}
                    onChange={(e) => setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, display_and_leds_test: e.target.value }
                    })}
                    className="w-full px-2.5 py-1.5 border rounded-lg bg-slate-50 font-medium"
                  >
                    <option value="Pass">Pass (Sáng rõ)</option>
                    <option value="Fail">Fail (Lỗi màn/LED)</option>
                  </select>
                </div>
                <div>
                  <label className="block text-slate-500 mb-1">Tiếp địa vỏ rơ-le</label>
                  <select
                    value={data.visual_inspection.frame_grounding_check}
                    onChange={(e) => setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, frame_grounding_check: e.target.value }
                    })}
                    className="w-full px-2.5 py-1.5 border rounded-lg bg-slate-50 font-medium"
                  >
                    <option value="Pass">Pass (Nối đất tốt)</option>
                    <option value="Fail">Fail (Chưa tiếp địa)</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-slate-500 mb-1">Khối ngắn mạch CT</label>
                  <select
                    value={data.visual_inspection.ct_shorting_blocks_check}
                    onChange={(e) => setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, ct_shorting_blocks_check: e.target.value }
                    })}
                    className="w-full px-2.5 py-1.5 border rounded-lg bg-slate-50 font-medium"
                  >
                    <option value="Pass">Pass (Tin cậy)</option>
                    <option value="Fail">Fail (Nguy cơ hở CT)</option>
                  </select>
                </div>
                <div>
                  <label className="block text-slate-500 mb-1">Đồng hồ RTC</label>
                  <input
                    type="text"
                    value={data.visual_inspection.clock_date_verified}
                    onChange={(e) => setData({
                      ...data,
                      visual_inspection: { ...data.visual_inspection, clock_date_verified: e.target.value }
                    })}
                    className="w-full px-2 py-1.5 border rounded-lg bg-slate-50 font-mono text-[11px]"
                  />
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Right 2 Columns: Section 3 Electrical Tests & Protection Elements */}
        <div className="space-y-6 lg:col-span-2">
          {/* Section 3: Electrical Tests */}
          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs space-y-5">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h3 className="font-bold text-sm text-slate-900 flex items-center gap-2">
                <Zap size={16} className="text-amber-500" />
                3. Phép Đo Điện & Chức Năng Bảo Vệ (NETA 7.9.2.B)
              </h3>
              <span className={`px-2.5 py-1 rounded-full text-xs font-bold ${liveStats.overallSafe ? 'bg-emerald-100 text-emerald-800' : 'bg-red-100 text-red-800'}`}>
                {liveStats.overallSafe ? 'TẤT CẢ PHÉP ĐO ĐẠT' : 'PHÁT HIỆN SAI SỐ / LỖI'}
              </span>
            </div>

            {/* 3.1 Insulation Resistance 1000V DC */}
            <div>
              <div className="flex items-center justify-between mb-2">
                <span className="font-bold text-xs text-slate-700 flex items-center gap-1.5">
                  <ShieldCheck size={14} className="text-blue-600" />
                  3.1 Điện trở cách điện @ 1000V DC (Insulation Resistance - Tiêu chuẩn ≥ 100 MΩ)
                </span>
                <span className={`text-[11px] font-bold ${liveStats.irAllPass ? 'text-emerald-600' : 'text-red-600'}`}>
                  {liveStats.irAllPass ? '✓ ĐẠT CÁCH ĐIỆN' : '✗ KHÔNG ĐẠT CÁCH ĐIỆN'}
                </span>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
                  <span className="text-[11px] text-slate-500 block mb-1">Mạch dòng AC đến vỏ (MΩ)</span>
                  <div className="flex items-center gap-2">
                    <input
                      type="number"
                      value={data.electrical_tests.insulation_resistance_1000v_megohms.ac_current_inputs_to_ground}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1000v_megohms: {
                            ...data.electrical_tests.insulation_resistance_1000v_megohms,
                            ac_current_inputs_to_ground: Number(e.target.value)
                          }
                        }
                      })}
                      className="w-full font-bold text-slate-800 bg-white border border-slate-300 rounded px-2 py-1 text-sm"
                    />
                    <span className="text-xs text-slate-500 font-semibold">MΩ</span>
                  </div>
                </div>

                <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
                  <span className="text-[11px] text-slate-500 block mb-1">Mạch áp AC đến vỏ (MΩ)</span>
                  <div className="flex items-center gap-2">
                    <input
                      type="number"
                      value={data.electrical_tests.insulation_resistance_1000v_megohms.ac_voltage_inputs_to_ground}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1000v_megohms: {
                            ...data.electrical_tests.insulation_resistance_1000v_megohms,
                            ac_voltage_inputs_to_ground: Number(e.target.value)
                          }
                        }
                      })}
                      className="w-full font-bold text-slate-800 bg-white border border-slate-300 rounded px-2 py-1 text-sm"
                    />
                    <span className="text-xs text-slate-500 font-semibold">MΩ</span>
                  </div>
                </div>

                <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
                  <span className="text-[11px] text-slate-500 block mb-1">Nguồn nuôi DC đến vỏ (MΩ)</span>
                  <div className="flex items-center gap-2">
                    <input
                      type="number"
                      value={data.electrical_tests.insulation_resistance_1000v_megohms.dc_power_supply_to_ground}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1000v_megohms: {
                            ...data.electrical_tests.insulation_resistance_1000v_megohms,
                            dc_power_supply_to_ground: Number(e.target.value)
                          }
                        }
                      })}
                      className="w-full font-bold text-slate-800 bg-white border border-slate-300 rounded px-2 py-1 text-sm"
                    />
                    <span className="text-xs text-slate-500 font-semibold">MΩ</span>
                  </div>
                </div>
              </div>
            </div>

            {/* 3.2 Analog Metering Accuracy */}
            <div>
              <span className="font-bold text-xs text-slate-700 block mb-2">
                3.2 Đo lường Tương tự (Analog Input Metering - Sai số cho phép ≤ ±0.5%)
              </span>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
                  <label className="text-slate-500 block mb-1">Bơm dòng IA = 5.00 A:</label>
                  <input
                    type="text"
                    value={data.electrical_tests.analog_metering_accuracy.injected_ia_5_00a || ''}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        analog_metering_accuracy: {
                          ...data.electrical_tests.analog_metering_accuracy,
                          injected_ia_5_00a: e.target.value
                        }
                      }
                    })}
                    className="w-full px-2.5 py-1.5 bg-white border rounded font-mono text-slate-800"
                  />
                </div>
                <div className="p-3 bg-slate-50 rounded-xl border border-slate-200">
                  <label className="text-slate-500 block mb-1">Bơm áp VA = 66.4 V:</label>
                  <input
                    type="text"
                    value={data.electrical_tests.analog_metering_accuracy.injected_va_66_4v || ''}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        analog_metering_accuracy: {
                          ...data.electrical_tests.analog_metering_accuracy,
                          injected_va_66_4v: e.target.value
                        }
                      }
                    })}
                    className="w-full px-2.5 py-1.5 bg-white border rounded font-mono text-slate-800"
                  />
                </div>
              </div>
            </div>

            {/* 3.3 Protection Elements: ANSI 51 Time-Overcurrent */}
            <div className="p-4 bg-blue-50/60 rounded-xl border border-blue-200 space-y-3">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <Clock size={16} className="text-blue-700" />
                  <span className="font-bold text-xs text-blue-950">
                    3.3 Bảo vệ Quá dòng có thời gian (ANSI 51 Time-Overcurrent)
                  </span>
                </div>
                <span className={`px-2 py-0.5 rounded text-[11px] font-bold ${liveStats.ansi51Pass ? 'bg-emerald-100 text-emerald-800' : 'bg-red-100 text-red-800'}`}>
                  {liveStats.ansi51Pass ? 'ĐẠT DUNG SAI ±5%' : 'LỖI SAI SỐ'}
                </span>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
                <div>
                  <label className="text-slate-600 block mb-1">Cài đặt Pickup (A)</label>
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.protection_elements_test.ansi_51_time_overcurrent.setting_pickup_a}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        protection_elements_test: {
                          ...data.electrical_tests.protection_elements_test,
                          ansi_51_time_overcurrent: {
                            ...data.electrical_tests.protection_elements_test.ansi_51_time_overcurrent,
                            setting_pickup_a: Number(e.target.value)
                          }
                        }
                      }
                    })}
                    className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 font-bold text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 block mb-1">Đo thực tế Pickup (A)</label>
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.protection_elements_test.ansi_51_time_overcurrent.measured_pickup_a}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        protection_elements_test: {
                          ...data.electrical_tests.protection_elements_test,
                          ansi_51_time_overcurrent: {
                            ...data.electrical_tests.protection_elements_test.ansi_51_time_overcurrent,
                            measured_pickup_a: Number(e.target.value)
                          }
                        }
                      }
                    })}
                    className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 font-bold text-blue-900"
                  />
                  <span className="text-[10px] text-slate-500 mt-0.5 block">
                    Sai số: {liveStats.ansi51PickupErr > 0 ? '+' : ''}{liveStats.ansi51PickupErr.toFixed(2)}%
                  </span>
                </div>

                <div>
                  <label className="text-slate-600 block mb-1">Thời gian lý thuyết @3x (s)</label>
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.protection_elements_test.ansi_51_time_overcurrent.expected_trip_time_sec_at_3x}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        protection_elements_test: {
                          ...data.electrical_tests.protection_elements_test,
                          ansi_51_time_overcurrent: {
                            ...data.electrical_tests.protection_elements_test.ansi_51_time_overcurrent,
                            expected_trip_time_sec_at_3x: Number(e.target.value)
                          }
                        }
                      }
                    })}
                    className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 font-bold text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 block mb-1">Đo thực tế @3x (s)</label>
                  <input
                    type="number"
                    step="0.01"
                    value={data.electrical_tests.protection_elements_test.ansi_51_time_overcurrent.measured_trip_time_sec_at_3x}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        protection_elements_test: {
                          ...data.electrical_tests.protection_elements_test,
                          ansi_51_time_overcurrent: {
                            ...data.electrical_tests.protection_elements_test.ansi_51_time_overcurrent,
                            measured_trip_time_sec_at_3x: Number(e.target.value)
                          }
                        }
                      }
                    })}
                    className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 font-bold text-blue-900"
                  />
                  <span className="text-[10px] text-slate-500 mt-0.5 block">
                    Sai số: {liveStats.ansi51TimeErr > 0 ? '+' : ''}{liveStats.ansi51TimeErr.toFixed(2)}%
                  </span>
                </div>
              </div>
            </div>

            {/* 3.4 Protection Elements: ANSI 50 Instantaneous Overcurrent */}
            <div className={`p-4 rounded-xl border space-y-3 ${liveStats.ansi50Pass ? 'bg-emerald-50/60 border-emerald-200' : 'bg-red-50/80 border-red-300'}`}>
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <Zap size={16} className={liveStats.ansi50Pass ? 'text-emerald-700' : 'text-red-600'} />
                  <span className={`font-bold text-xs ${liveStats.ansi50Pass ? 'text-emerald-950' : 'text-red-950'}`}>
                    3.4 Bảo vệ Cắt nhanh tức thời (ANSI 50 Instantaneous Overcurrent)
                  </span>
                </div>
                <span className={`px-2.5 py-0.5 rounded text-[11px] font-black ${liveStats.ansi50Pass ? 'bg-emerald-200 text-emerald-900' : 'bg-red-200 text-red-900 animate-pulse'}`}>
                  {liveStats.ansi50Pass ? 'PASS (≤ ±5%)' : `FAIL: SAI SỐ ${liveStats.ansi50PickupErr > 0 ? '+' : ''}${liveStats.ansi50PickupErr.toFixed(1)}%`}
                </span>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
                <div>
                  <label className="text-slate-600 block mb-1">Cài đặt Pickup (A)</label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.protection_elements_test.ansi_50_instantaneous_overcurrent.setting_pickup_a}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        protection_elements_test: {
                          ...data.electrical_tests.protection_elements_test,
                          ansi_50_instantaneous_overcurrent: {
                            ...data.electrical_tests.protection_elements_test.ansi_50_instantaneous_overcurrent,
                            setting_pickup_a: Number(e.target.value)
                          }
                        }
                      }
                    })}
                    className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 font-bold text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 block mb-1">Đo thực tế Pickup (A)</label>
                  <input
                    type="number"
                    step="0.1"
                    value={data.electrical_tests.protection_elements_test.ansi_50_instantaneous_overcurrent.measured_pickup_a}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        protection_elements_test: {
                          ...data.electrical_tests.protection_elements_test,
                          ansi_50_instantaneous_overcurrent: {
                            ...data.electrical_tests.protection_elements_test,
                            ansi_50_instantaneous_overcurrent: Number(e.target.value)
                          }
                        }
                      }
                    })}
                    className={`w-full bg-white border rounded px-2.5 py-1.5 font-black ${liveStats.ansi50Pass ? 'text-emerald-900 border-slate-300' : 'text-red-700 border-red-400'}`}
                  />
                  <span className={`text-[10px] font-bold mt-0.5 block ${liveStats.ansi50Pass ? 'text-emerald-700' : 'text-red-700'}`}>
                    Lệch: {liveStats.ansi50PickupErr > 0 ? '+' : ''}{liveStats.ansi50PickupErr.toFixed(1)}% (Dung sai: ±5%)
                  </span>
                </div>

                <div>
                  <label className="text-slate-600 block mb-1">Thời gian lý thuyết (s)</label>
                  <input
                    type="number"
                    step="0.001"
                    value={data.electrical_tests.protection_elements_test.ansi_50_instantaneous_overcurrent.expected_trip_time_sec}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        protection_elements_test: {
                          ...data.electrical_tests.protection_elements_test,
                          ansi_50_instantaneous_overcurrent: {
                            ...data.electrical_tests.protection_elements_test.ansi_50_instantaneous_overcurrent,
                            expected_trip_time_sec: Number(e.target.value)
                          }
                        }
                      }
                    })}
                    className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 font-mono text-slate-800"
                  />
                </div>

                <div>
                  <label className="text-slate-600 block mb-1">Đo thực tế Trip time (s)</label>
                  <input
                    type="number"
                    step="0.001"
                    value={data.electrical_tests.protection_elements_test.ansi_50_instantaneous_overcurrent.measured_trip_time_sec}
                    onChange={(e) => setData({
                      ...data,
                      electrical_tests: {
                        ...data.electrical_tests,
                        protection_elements_test: {
                          ...data.electrical_tests.protection_elements_test,
                          ansi_50_instantaneous_overcurrent: {
                            ...data.electrical_tests.protection_elements_test.ansi_50_instantaneous_overcurrent,
                            measured_trip_time_sec: Number(e.target.value)
                          }
                        }
                      }
                    })}
                    className="w-full bg-white border border-slate-300 rounded px-2.5 py-1.5 font-mono text-slate-800"
                  />
                </div>
              </div>
            </div>

            {/* 3.5 Digital I/O, Trip Contacts, Logic & TCS */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs pt-1">
              <div>
                <label className="block text-slate-600 font-bold mb-1">Ngõ vào số (Digital Inputs):</label>
                <input
                  type="text"
                  value={data.electrical_tests.digital_inputs_test}
                  onChange={(e) => setData({
                    ...data,
                    electrical_tests: { ...data.electrical_tests, digital_inputs_test: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border rounded-lg font-medium"
                />
              </div>

              <div>
                <label className="block text-slate-600 font-bold mb-1">Ngõ ra Trip máy cắt (Output Contacts):</label>
                <input
                  type="text"
                  value={data.electrical_tests.output_contacts_trip_test}
                  onChange={(e) => setData({
                    ...data,
                    electrical_tests: { ...data.electrical_tests, output_contacts_trip_test: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border rounded-lg font-medium"
                />
              </div>

              <div>
                <label className="block text-slate-600 font-bold mb-1">Logic nội bộ & Liên động (PSL/GOOSE):</label>
                <input
                  type="text"
                  value={data.electrical_tests.internal_logic_verification}
                  onChange={(e) => setData({
                    ...data,
                    electrical_tests: { ...data.electrical_tests, internal_logic_verification: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border rounded-lg font-medium"
                />
              </div>

              <div>
                <label className="block text-slate-600 font-bold mb-1">Giám sát mạch cắt (TCS):</label>
                <input
                  type="text"
                  value={data.electrical_tests.trip_coil_monitoring_tcs}
                  onChange={(e) => setData({
                    ...data,
                    electrical_tests: { ...data.electrical_tests, trip_coil_monitoring_tcs: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border rounded-lg font-medium"
                />
              </div>
            </div>

            {/* 3.6 MANDATORY NETA REQUIREMENT: EVENT RECORDS CLEARED */}
            <div className={`p-4 rounded-xl border-2 transition-all ${data.electrical_tests.event_records_cleared_after_test ? 'bg-emerald-50/70 border-emerald-300' : 'bg-amber-50/90 border-amber-400'}`}>
              <label className="flex items-start gap-3 cursor-pointer">
                <input
                  type="checkbox"
                  checked={data.electrical_tests.event_records_cleared_after_test}
                  onChange={(e) => setData({
                    ...data,
                    electrical_tests: { ...data.electrical_tests, event_records_cleared_after_test: e.target.checked }
                  })}
                  className="rounded text-emerald-600 focus:ring-emerald-500 w-5 h-5 mt-0.5 shrink-0"
                />
                <div className="flex-1">
                  <div className="flex items-center gap-2">
                    <span className="font-extrabold text-sm text-slate-900">
                      BẮT BUỘC: Đã Xóa Sạch Nhật Ký Sự Cố Giả Sau Thử Nghiệm (Reset Event Records)
                    </span>
                    <span className="px-2 py-0.5 rounded text-[10px] font-black uppercase tracking-wider bg-amber-200 text-amber-900 border border-amber-300">
                      NETA 7.9.2.B.8
                    </span>
                  </div>
                  <p className="text-xs text-slate-700 mt-1 leading-relaxed">
                    Theo tiêu chuẩn ANSI/NETA ATS-2025 Mục 7.9.2, sau khi kết thúc mọi phép thử bơm dòng, FSE <strong>BẮT BUỘC</strong> phải truy cập phần mềm hoặc mặt máy rơ-le để xóa sạch toàn bộ các bản ghi sự cố giả (Fault logs, SER/SOE, Event reports) phát sinh trong quá trình thử nghiệm.
                  </p>
                  {!data.electrical_tests.event_records_cleared_after_test && (
                    <div className="mt-2 text-xs font-bold text-red-700 flex items-center gap-1.5">
                      <AlertTriangle size={14} />
                      <span>CẢNH BÁO: Chưa tích chọn xác nhận xóa nhật ký sự cố. Không đủ điều kiện nghiệm thu đóng điện!</span>
                    </div>
                  )}
                </div>
              </label>
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
                analysisResult.overall_status === 'PASS' ? 'bg-emerald-600' : 'bg-red-600 animate-pulse'
              }`}>
                {analysisResult.overall_status === 'PASS' ? <CheckCircle2 size={28} /> : <AlertTriangle size={28} />}
              </div>
              <div>
                <span className="text-[11px] font-bold text-amber-400 uppercase tracking-wider">
                  KẾT QUẢ ĐỐI SOÁT AI CHUYÊN GIA NETA ATS-2025
                </span>
                <h2 className="text-xl sm:text-2xl font-black">
                  Trạng thái: {analysisResult.overall_status === 'PASS' ? 'PASS (ĐẠT CHUẨN ĐÓNG ĐIỆN)' : 'FAIL / CALIBRATION & LOG CLEARING REQUIRED'}
                </h2>
              </div>
            </div>

            <button
              onClick={() => setShowPrintModal(true)}
              className="px-4 py-2 bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold rounded-xl shadow-md transition-all flex items-center justify-center gap-2 self-start sm:self-auto"
            >
              <Printer size={15} />
              <span>In Báo Cáo PDF Chính Thức</span>
            </button>
          </div>

          <div className="p-6 bg-slate-50 space-y-6">
            {/* Quick Status Reason Banner */}
            <div className={`p-4 rounded-xl border font-medium text-xs sm:text-sm leading-relaxed ${
              analysisResult.overall_status === 'PASS' ? 'bg-emerald-50 border-emerald-200 text-emerald-900' : 'bg-red-50 border-red-200 text-red-900'
            }`}>
              <strong>Tóm lược đánh giá:</strong> {analysisResult.status_reason}
            </div>

            {/* Markdown Report Render */}
            <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-xs prose prose-slate max-w-none prose-headings:font-bold prose-h1:text-xl prose-h2:text-base prose-h3:text-sm prose-p:text-xs prose-p:leading-relaxed prose-table:text-xs prose-th:bg-slate-100 prose-th:p-2 prose-td:p-2">
              <div className="whitespace-pre-wrap font-sans text-xs leading-relaxed text-slate-800">
                {analysisResult.markdown_report}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* JSON Import/Export Modal */}
      {showJsonModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-white w-full max-w-2xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[85vh]">
            <div className="p-4 border-b border-slate-100 flex items-center justify-between bg-slate-50">
              <h3 className="font-bold text-slate-800 text-sm flex items-center gap-2">
                <FileText size={16} className="text-blue-600" />
                Mã JSON Khảo Sát Hiện Trường Rơ-le (Field Service Input)
              </h3>
              <button onClick={() => setShowJsonModal(false)} className="text-slate-400 hover:text-slate-600 text-lg">
                &times;
              </button>
            </div>
            <div className="p-4 flex-1 overflow-y-auto">
              <p className="text-xs text-slate-500 mb-2">
                Dán chuỗi JSON kiểm tra rơ-le số từ hợp bộ thử nghiệm (Omicron, Ponovo) hoặc hệ thống TEV:
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
                      setSavedSuccessMessage('Đã tải dữ liệu JSON rơ-le thành công!');
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
        title="BIÊN BẢN THỬ NGHIỆM RƠ-LE BẢO VỆ VI XỬ LÝ"
        standard="ANSI/NETA ATS-2025 Section 7.9.2"
        siteInfo={data.site_info}
        overallStatus={analysisResult?.overall_status || (liveStats.overallSafe ? 'PASS' : 'FAIL')}
        statusReason={analysisResult?.status_reason || (liveStats.overallSafe ? 'Đạt chuẩn NETA ATS-2025 Section 7.9.2.' : 'Chưa đạt do sai số cắt nhanh ANSI 50 hoặc chưa xóa Event records.')}
        visualChecks={visualChecksForPrint}
        electricalTable={electricalTableForPrint}
        warnings={
          liveStats.overallSafe
            ? []
            : [
                !liveStats.ansi50Pass ? `Bảo vệ cắt nhanh ANSI 50 có sai số khởi động ${liveStats.ansi50PickupErr > 0 ? '+' : ''}${liveStats.ansi50PickupErr.toFixed(1)}% vượt quá dung sai cho phép ±5.0% của nhà sản xuất SEL.` : '',
                !liveStats.recordsCleared ? `Chưa xóa nhật ký sự cố (Fault, Event records, SOE) sau thử nghiệm -> Vi phạm NETA ATS-2025 Sec 7.9.2.B.8.` : '',
                !liveStats.irAllPass ? `Điện trở cách điện suy giảm dưới 100 MΩ.` : '',
                !liveStats.ansi51Pass ? `Bảo vệ quá dòng có thời gian ANSI 51 có sai số vượt quá ±5.0%.` : ''
              ].filter(Boolean)
        }
        recommendations={[
          'Kiểm tra lại cấu hình tỉ số biến dòng CT ratio (CTR) trong menu cài đặt rơ-le SEL-751A và hiệu chuẩn lại kênh dòng Analog Input.',
          'BẮT BUỘC kết nối phần mềm SEL AcSELerator QuickSet để thực hiện lệnh RESET EVENT RECORDS và CLEAR FAULT LOGS trước khi đóng điện.',
          'Thử nghiệm liên động Trip máy cắt qua tiếp điểm lực OUT101 và kiểm tra tín hiệu cảnh báo hở mạch TCS.'
        ]}
      />

      {/* Camera Nameplate OCR Modal */}
      {showCameraModal && (
        <NameplateScannerModal
          isOpen={showCameraModal}
          onClose={() => setShowCameraModal(false)}
          onApplyData={handleApplyNameplateData}
          currentTag={data.site_info.relay_tag}
        />
      )}
    </div>
  );
};
