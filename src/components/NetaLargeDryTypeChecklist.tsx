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
  Sliders,
  Activity,
  Gauge,
  ShieldCheck,
  Compass,
  Thermometer,
  Wind,
  Camera
} from 'lucide-react';
import { NetaLargeInputPayload, NetaLargeEvaluationResult } from '../../server/netaLargeAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';
import { NameplateScannerModal } from './NameplateScannerModal';
import { ExtractedNameplateData } from '../../server/nameplateOcrAnalyzer';
import { FseEngineerDropdown } from './FseEngineerDropdown';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

const SAMPLE_LARGE_FSE_DATA: NetaLargeInputPayload = {
  site_info: {
    project_name: "Tram Bien Ap Trung Tam TEV",
    transformer_tag: "TR-MV-01",
    fse_name: "Võ Minh Tú (sgm1707@gmail.com)",
    test_date: "2026-09-22"
  },
  visual_inspection: {
    nameplate_match: true,
    shipping_brackets_removed: true,
    temp_indicator_settings_check: "Pass",
    cooling_fans_and_overcurrent_check: "Pass",
    bolt_torque_check: "Pass",
    surge_arresters_present: true
  },
  electrical_tests: {
    bolted_resistance_micro_ohms: [8.1, 8.3, 8.5],
    insulation_resistance_1min_megohms: {
      test_voltage_dc: 2500,
      pri_to_ground: 3500,
      sec_to_ground: 1200,
      pri_to_sec: 4100
    },
    pi_value: 1.45,
    turns_ratio_all_taps: {
      max_error_percent: 0.12
    },
    winding_resistance_all_taps: {
      max_deviation_from_factory_percent: 2.3
    },
    excitation_current_milliamp: {
      phase_a: 42.1,
      phase_b: 28.5,
      phase_c: 41.8
    },
    core_insulation_resistance_500v_dc_megohms: 0.45
  },
  previous_test_data: {
    last_test_date: "2025-09-10",
    last_pri_ground_megohms: 3600
  }
};

export const NetaLargeDryTypeChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaLargeInputPayload>(SAMPLE_LARGE_FSE_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaLargeEvaluationResult | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showCameraModal, setShowCameraModal] = useState(false);
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  const handleApplyNameplateData = (extracted: ExtractedNameplateData) => {
    setData((prev) => ({
      ...prev,
      site_info: {
        ...prev.site_info,
        transformer_tag: extracted.equipment_tag || prev.site_info.transformer_tag,
      },
      visual_inspection: {
        ...prev.visual_inspection,
        nameplate_match: true,
      }
    }));
    setSavedSuccessMessage(`Đã quét nhãn thành công: ${extracted.equipment_tag} (${extracted.manufacturer || 'MBA'}) - Tự động cập nhật vào Checklist!`);
    setTimeout(() => setSavedSuccessMessage(null), 5000);
  };

  // Live Calculations
  const liveStats = useMemo(() => {
    // 1. Bolted resistance deviation
    const bolts = data.electrical_tests.bolted_resistance_micro_ohms;
    const minBolt = bolts.length > 0 ? Math.min(...bolts) : 0;
    const maxBolt = bolts.length > 0 ? Math.max(...bolts) : 0;
    const boltDev = minBolt > 0 ? ((maxBolt - minBolt) / minBolt) * 100 : 0;
    const boltPass = boltDev <= 50;

    // 2. Insulation Resistance Trend
    const irPri = data.electrical_tests.insulation_resistance_1min_megohms.pri_to_ground;
    const prevIr = data.previous_test_data?.last_pri_ground_megohms || 0;
    const irDrop = prevIr > 0 ? ((prevIr - irPri) / prevIr) * 100 : 0;
    const irTrendWarning = irDrop > 20;

    // 3. PI Check
    const pi = data.electrical_tests.pi_value;
    const piPass = pi >= 1.0;

    // 4. Turns ratio error
    const ttrErr = data.electrical_tests.turns_ratio_all_taps.max_error_percent;
    const ttrPass = ttrErr <= 0.5;

    // 5. Winding resistance deviation
    const wrDev = data.electrical_tests.winding_resistance_all_taps.max_deviation_from_factory_percent;
    const wrPass = wrDev <= 1.0;

    // 6. Excitation current pattern (3-legged core: A & C outer legs higher, B middle lower)
    const exc = data.electrical_tests.excitation_current_milliamp;
    const excVals = [exc.phase_a, exc.phase_b, exc.phase_c];
    const excMin = Math.min(...excVals);
    const excPatternPass = exc.phase_b === excMin && Math.abs(exc.phase_a - exc.phase_c) / Math.max(exc.phase_a, exc.phase_c) <= 0.15;

    // 7. Core Insulation Resistance (>= 1.0 MΩ @ 500V DC)
    const coreIr = data.electrical_tests.core_insulation_resistance_500v_dc_megohms;
    const coreIrPass = coreIr >= 1.0;

    return {
      minBolt,
      maxBolt,
      boltDev: parseFloat(boltDev.toFixed(1)),
      boltPass,
      irDrop: parseFloat(irDrop.toFixed(1)),
      irTrendWarning,
      piPass,
      ttrPass,
      wrPass,
      excPatternPass,
      coreIrPass
    };
  }, [data]);

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    try {
      const response = await fetch('/api/field-service/neta-large-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });
      const res = await response.json();
      if (res.success) {
        setAnalysisResult(res.data);
      } else {
        alert('Lỗi phân tích: ' + res.error);
      }
    } catch (err: any) {
      alert('Không thể kết nối đến máy chủ: ' + err.message);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const loadSample = () => {
    setData(SAMPLE_LARGE_FSE_DATA);
    setAnalysisResult(null);
    setSavedSuccessMessage(null);
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
        equipmentId: data.site_info.transformer_tag,
        date: data.site_info.test_date,
        type: 'Transformers, Dry-Type, Air-Cooled, Large (NETA ATS-2025 Sec 7.2.1.2)'
      });
      setSavedSuccessMessage(`Đã lưu biên bản kiểm định MBA lớn [${data.site_info.transformer_tag}] vào kho báo cáo CMMS!`);
    }
  };

  // Prepare table data for print modal
  const electricalTableForPrint = [
    {
      item: 'Điện trở mối nối bu-lông',
      measured: `[${data.electrical_tests.bolted_resistance_micro_ohms.join(', ')}] µΩ`,
      standard: 'Lệch ≤ 50% so với mối nối nhỏ nhất',
      reference: `Min: ${liveStats.minBolt} µΩ`,
      deviation: `Lệch lớn nhất: ${liveStats.boltDev}%`,
      status: liveStats.boltPass ? 'PASS' : 'INVESTIGATE' as const
    },
    {
      item: `Điện trở cách điện Sơ-Đất (${data.electrical_tests.insulation_resistance_1min_megohms.test_voltage_dc}V DC)`,
      measured: `${data.electrical_tests.insulation_resistance_1min_megohms.pri_to_ground} MΩ`,
      standard: 'Table 100.5 (≥ 1000 MΩ)',
      reference: `${data.previous_test_data?.last_pri_ground_megohms || 'N/A'} MΩ`,
      deviation: liveStats.irDrop > 0 ? `Giảm ${liveStats.irDrop}% so với trước` : 'Ổn định',
      status: liveStats.irTrendWarning ? 'INVESTIGATE' : 'PASS' as const
    },
    {
      item: 'Chỉ số phân cực (Polarization Index - PI)',
      measured: `${data.electrical_tests.pi_value}`,
      standard: 'Bắt buộc ≥ 1.0 (NETA ATS-2025)',
      reference: 'Chuẩn ≥ 1.0',
      deviation: `PI = ${data.electrical_tests.pi_value}`,
      status: liveStats.piPass ? 'PASS' : 'FAIL' as const
    },
    {
      item: 'Tỷ số biến áp (Turns-Ratio All Taps)',
      measured: `Sai số max: ${data.electrical_tests.turns_ratio_all_taps.max_error_percent}%`,
      standard: 'Sai số ≤ 0.5% so với tính toán/kế cận',
      reference: 'Tất cả nấc tap',
      deviation: `${data.electrical_tests.turns_ratio_all_taps.max_error_percent}%`,
      status: liveStats.ttrPass ? 'PASS' : 'FAIL' as const
    },
    {
      item: 'Điện trở cuộn dây (Winding Resistance)',
      measured: `Lệch: ${data.electrical_tests.winding_resistance_all_taps.max_deviation_from_factory_percent}%`,
      standard: 'Quy đổi nhiệt độ ≤ 1.0% (within 1%)',
      reference: 'Dữ liệu xuất xưởng',
      deviation: `Lệch ${data.electrical_tests.winding_resistance_all_taps.max_deviation_from_factory_percent}% (> 1.0%)`,
      status: liveStats.wrPass ? 'PASS' : 'INVESTIGATE' as const
    },
    {
      item: 'Dòng kích thích (Excitation Current)',
      measured: `A: ${data.electrical_tests.excitation_current_milliamp.phase_a}mA, B: ${data.electrical_tests.excitation_current_milliamp.phase_b}mA, C: ${data.electrical_tests.excitation_current_milliamp.phase_c}mA`,
      standard: 'Lõi 3 trụ: 2 pha cao tương đương, 1 pha thấp',
      reference: 'Mẫu chuẩn 2 cao 1 thấp',
      deviation: liveStats.excPatternPass ? 'Đúng mẫu chuẩn lõi 3 trụ' : 'Không chuẩn mẫu',
      status: liveStats.excPatternPass ? 'PASS' : 'INVESTIGATE' as const
    },
    {
      item: 'Điện trở cách điện lõi thép (Core IR @ 500V DC)',
      measured: `${data.electrical_tests.core_insulation_resistance_500v_dc_megohms} MΩ`,
      standard: 'Tối thiểu ≥ 1.0 Megohm (1 MΩ)',
      reference: 'NETA ATS-2025 Sec 7.2.1.2.2',
      deviation: `${data.electrical_tests.core_insulation_resistance_500v_dc_megohms} MΩ < 1.0 MΩ (Lỗi chạm đất lõi)`,
      status: liveStats.coreIrPass ? 'PASS' : 'FAIL' as const
    }
  ];

  const visualChecksForPrint = [
    { label: 'Nameplate khớp bản vẽ thiết kế', value: data.visual_inspection.nameplate_match ? 'ĐẠT' : 'KHÔNG ĐẠT', pass: data.visual_inspection.nameplate_match },
    { label: 'Khung kẹp vận chuyển (Shipping brackets) đã tháo rời', value: data.visual_inspection.shipping_brackets_removed ? 'ĐÃ THÁO RỜI' : 'CHƯA THÁO', pass: data.visual_inspection.shipping_brackets_removed },
    { label: 'Cài đặt đồng hồ nhiệt độ (Alarm/Trip)', value: data.visual_inspection.temp_indicator_settings_check, pass: data.visual_inspection.temp_indicator_settings_check === 'Pass' },
    { label: 'Hệ thống quạt làm mát & Rơ-le quá dòng', value: data.visual_inspection.cooling_fans_and_overcurrent_check, pass: data.visual_inspection.cooling_fans_and_overcurrent_check === 'Pass' },
    { label: 'Lực siết bu-lông (Bolt Torque Table 100.12)', value: data.visual_inspection.bolt_torque_check, pass: data.visual_inspection.bolt_torque_check === 'Pass' },
    { label: 'Chống sét van (Surge Arresters) lắp đặt đầy đủ', value: data.visual_inspection.surge_arresters_present ? 'ĐÃ LẮP ĐẶT' : 'CHƯA LẮP ĐẶT', pass: data.visual_inspection.surge_arresters_present }
  ];

  return (
    <div className="space-y-6">
      {/* Top Banner & Action Header */}
      <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-blue-900 text-white p-5 sm:p-6 rounded-2xl shadow-lg border border-slate-700/50">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2 mb-1.5">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-400/20 text-amber-300 border border-amber-400/30 uppercase tracking-wide">
                NETA ATS-2025 Section 7.2.1.2
              </span>
              <span className="text-xs text-slate-300">Công suất lớn • &gt;600V hoặc &gt;500 kVA</span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white flex items-center gap-2.5">
              <Zap className="text-amber-400 fill-amber-400/20" size={24} />
              Checklist MBA Khô Dung Lượng Lớn (Dry-Type Large)
            </h1>
            <p className="text-xs sm:text-sm text-slate-300 mt-1 max-w-2xl">
              Quy chuẩn kiểm tra toàn diện: Cách điện lõi thép (Core IR), Điện trở cuộn dây all taps, Dòng kích thích lõi 3 trụ, Quạt làm mát và Chống sét van.
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
              onClick={loadSample}
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title="Tải kịch bản mẫu TR-MV-01"
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

        {/* Live Status Indicators Bar */}
        <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-2.5 mt-5 pt-4 border-t border-white/10 text-xs">
          <div className="bg-white/5 rounded-lg p-2 border border-white/10">
            <div className="text-[10px] text-slate-400">Độ lệch Bu-lông</div>
            <div className={`font-mono font-bold text-sm ${liveStats.boltPass ? 'text-emerald-400' : 'text-amber-400'}`}>
              {liveStats.boltDev}%
            </div>
            <div className="text-[9px] text-slate-400">Chuẩn ≤ 50%</div>
          </div>

          <div className="bg-white/5 rounded-lg p-2 border border-white/10">
            <div className="text-[10px] text-slate-400">Điện trở Sơ-Đất</div>
            <div className="font-mono font-bold text-sm text-blue-300">
              {data.electrical_tests.insulation_resistance_1min_megohms.pri_to_ground} MΩ
            </div>
            <div className={`text-[9px] ${liveStats.irTrendWarning ? 'text-amber-400 font-bold' : 'text-slate-400'}`}>
              {liveStats.irDrop > 0 ? `Giảm ${liveStats.irDrop}%` : 'Tốt'}
            </div>
          </div>

          <div className="bg-white/5 rounded-lg p-2 border border-white/10">
            <div className="text-[10px] text-slate-400">Chỉ số PI</div>
            <div className={`font-mono font-bold text-sm ${liveStats.piPass ? 'text-emerald-400' : 'text-rose-400'}`}>
              {data.electrical_tests.pi_value}
            </div>
            <div className="text-[9px] text-slate-400">Chuẩn ≥ 1.0</div>
          </div>

          <div className="bg-white/5 rounded-lg p-2 border border-white/10">
            <div className="text-[10px] text-slate-400">Sai số TTR</div>
            <div className={`font-mono font-bold text-sm ${liveStats.ttrPass ? 'text-emerald-400' : 'text-rose-400'}`}>
              {data.electrical_tests.turns_ratio_all_taps.max_error_percent}%
            </div>
            <div className="text-[9px] text-slate-400">Chuẩn ≤ 0.5%</div>
          </div>

          <div className="bg-white/5 rounded-lg p-2 border border-white/10">
            <div className="text-[10px] text-slate-400">Lệch Cuộn Dây (WR)</div>
            <div className={`font-mono font-bold text-sm ${liveStats.wrPass ? 'text-emerald-400' : 'text-amber-400'}`}>
              {data.electrical_tests.winding_resistance_all_taps.max_deviation_from_factory_percent}%
            </div>
            <div className="text-[9px] text-slate-400">Chuẩn ≤ 1.0%</div>
          </div>

          <div className="bg-white/5 rounded-lg p-2 border border-white/10">
            <div className="text-[10px] text-slate-400">Dòng Kích Thích</div>
            <div className={`font-mono font-bold text-xs ${liveStats.excPatternPass ? 'text-emerald-400' : 'text-amber-400'}`}>
              {liveStats.excPatternPass ? 'Chuẩn (2 Cao 1 Thấp)' : 'Lệch Dạng Mẫu'}
            </div>
            <div className="text-[9px] text-slate-400">Lõi 3 trụ</div>
          </div>

          <div className="bg-white/5 rounded-lg p-2 border border-white/10">
            <div className="text-[10px] text-slate-400">Core IR (Lõi thép)</div>
            <div className={`font-mono font-bold text-sm ${liveStats.coreIrPass ? 'text-emerald-400' : 'text-rose-400 animate-pulse'}`}>
              {data.electrical_tests.core_insulation_resistance_500v_dc_megohms} MΩ
            </div>
            <div className="text-[9px] text-slate-400">Chuẩn ≥ 1.0 MΩ</div>
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

      {/* Main Form Sections */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left Column: Site Info & Visual Inspection */}
        <div className="space-y-6 lg:col-span-1">
          {/* Site Info */}
          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-2">
              <h3 className="font-bold text-sm text-slate-900 flex items-center gap-2">
                <Compass size={16} className="text-blue-600" />
                Thông tin Hiện trường & Nhãn MBA
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
                <span>Quét bảng tên kim loại để AI tự điền thông số</span>
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
                <label className="block text-slate-500 mb-1">Mã Tag Máy Biến Áp</label>
                <input
                  type="text"
                  value={data.site_info.transformer_tag}
                  onChange={(e) => setData({ ...data, site_info: { ...data.site_info, transformer_tag: e.target.value } })}
                  className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white focus:ring-2 focus:ring-blue-500 font-mono font-bold text-blue-900"
                />
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

          {/* Visual & Mechanical Inspection (NETA 7.2.1.2.1) */}
          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs space-y-4">
            <h3 className="font-bold text-sm text-slate-900 flex items-center gap-2 border-b border-slate-100 pb-2">
              <ShieldCheck size={16} className="text-emerald-600" />
              1. Kiểm Tra Thị Giác & Cơ Khí
            </h3>

            <div className="space-y-3 text-xs">
              <div className="flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-100">
                <div>
                  <div className="font-semibold text-slate-800">Khớp nhãn Nameplate</div>
                  <div className="text-[10px] text-slate-400">Khớp bản vẽ thiết kế</div>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.nameplate_match}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, nameplate_match: e.target.checked }
                  })}
                  className="w-4 h-4 text-blue-600 rounded"
                />
              </div>

              <div className="flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-100">
                <div>
                  <div className="font-semibold text-slate-800">Khung kẹp vận chuyển</div>
                  <div className="text-[10px] text-slate-400">BẮT BUỘC đã tháo rời hoàn toàn</div>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.shipping_brackets_removed}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, shipping_brackets_removed: e.target.checked }
                  })}
                  className="w-4 h-4 text-blue-600 rounded"
                />
              </div>

              <div className="flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-100">
                <div>
                  <div className="font-semibold text-slate-800">Đồng hồ nhiệt độ</div>
                  <div className="text-[10px] text-slate-400">Cài đặt Alarm / Trip đúng thông số</div>
                </div>
                <select
                  value={data.visual_inspection.temp_indicator_settings_check}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, temp_indicator_settings_check: e.target.value }
                  })}
                  className="px-2 py-1 border rounded bg-white font-medium"
                >
                  <option value="Pass">Pass (Đạt)</option>
                  <option value="Fail">Fail (Lỗi)</option>
                </select>
              </div>

              <div className="flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-100">
                <div>
                  <div className="font-semibold text-slate-800">Quạt làm mát & Rơ-le quá dòng</div>
                  <div className="text-[10px] text-slate-400">Quạt quay êm, bảo vệ đúng dòng</div>
                </div>
                <select
                  value={data.visual_inspection.cooling_fans_and_overcurrent_check}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, cooling_fans_and_overcurrent_check: e.target.value }
                  })}
                  className="px-2 py-1 border rounded bg-white font-medium"
                >
                  <option value="Pass">Pass (Đạt)</option>
                  <option value="Fail">Fail (Lỗi)</option>
                </select>
              </div>

              <div className="flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-100">
                <div>
                  <div className="font-semibold text-slate-800">Lực siết bu-lông (Bolt Torque)</div>
                  <div className="text-[10px] text-slate-400">Theo Table 100.12 hoặc NSX</div>
                </div>
                <select
                  value={data.visual_inspection.bolt_torque_check}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, bolt_torque_check: e.target.value }
                  })}
                  className="px-2 py-1 border rounded bg-white font-medium"
                >
                  <option value="Pass">Pass (Đạt)</option>
                  <option value="Fail">Fail (Lỗi)</option>
                </select>
              </div>

              <div className="flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-100">
                <div>
                  <div className="font-semibold text-slate-800">Chống sét van (Surge Arresters)</div>
                  <div className="text-[10px] text-slate-400">Đã lắp đặt & đấu nối tiếp địa</div>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.surge_arresters_present}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, surge_arresters_present: e.target.checked }
                  })}
                  className="w-4 h-4 text-blue-600 rounded"
                />
              </div>
            </div>
          </div>
        </div>

        {/* Center & Right Column: Electrical Tests (NETA 7.2.1.2.2) */}
        <div className="space-y-6 lg:col-span-2">
          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs space-y-4">
            <h3 className="font-bold text-sm text-slate-900 flex items-center justify-between border-b border-slate-100 pb-2">
              <span className="flex items-center gap-2">
                <Gauge size={16} className="text-amber-500" />
                2. Phép Đo Điện & Tiêu Chuẩn NETA ATS-2025 Mục 7.2.1.2
              </span>
              <span className="text-[11px] text-slate-400 font-normal">Bảng so sánh giá trị tham chiếu</span>
            </h3>

            {/* Electrical Tests Grid */}
            <div className="space-y-4 text-xs">
              {/* Row 1: Bolted Connections Resistance */}
              <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/60 space-y-2">
                <div className="flex items-center justify-between">
                  <div>
                    <span className="font-bold text-slate-800">Điện trở mối nối bu-lông (DLRO - µΩ)</span>
                    <span className="text-[11px] text-slate-500 ml-2">NETA: Độ lệch ≤ 50% so với giá trị nhỏ nhất</span>
                  </div>
                  <span className={`px-2 py-0.5 rounded font-bold text-[10px] ${liveStats.boltPass ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'}`}>
                    {liveStats.boltPass ? 'ĐẠT (≤ 50%)' : `CẢNH BÁO (Lệch ${liveStats.boltDev}%)`}
                  </span>
                </div>
                <div className="flex items-center gap-2">
                  {data.electrical_tests.bolted_resistance_micro_ohms.map((val, idx) => (
                    <div key={idx} className="flex-1">
                      <label className="text-[10px] text-slate-400 block mb-0.5">Mối nối #{idx + 1}</label>
                      <input
                        type="number"
                        step="0.1"
                        value={val}
                        onChange={(e) => {
                          const newB = [...data.electrical_tests.bolted_resistance_micro_ohms];
                          newB[idx] = parseFloat(e.target.value) || 0;
                          setData({ ...data, electrical_tests: { ...data.electrical_tests, bolted_resistance_micro_ohms: newB } });
                        }}
                        className="w-full px-2.5 py-1.5 border rounded-lg bg-white font-mono font-medium text-slate-800"
                      />
                    </div>
                  ))}
                </div>
              </div>

              {/* Row 2: Insulation Resistance & PI */}
              <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/60 space-y-3">
                <div className="flex items-center justify-between">
                  <div>
                    <span className="font-bold text-slate-800">Điện trở cách điện (IR) & Chỉ số phân cực (PI)</span>
                    <span className="text-[11px] text-slate-500 ml-2">Table 100.5 • PI bắt buộc ≥ 1.0</span>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="text-[10px] text-slate-500">Điện áp thử:</span>
                    <select
                      value={data.electrical_tests.insulation_resistance_1min_megohms.test_voltage_dc}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1min_megohms: {
                            ...data.electrical_tests.insulation_resistance_1min_megohms,
                            test_voltage_dc: parseInt(e.target.value) || 2500
                          }
                        }
                      })}
                      className="px-2 py-0.5 border rounded bg-white font-bold text-blue-900"
                    >
                      <option value={1000}>1000V DC</option>
                      <option value={2500}>2500V DC</option>
                      <option value={5000}>5000V DC</option>
                    </select>
                  </div>
                </div>

                <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                  <div>
                    <label className="text-[10px] text-slate-400 block mb-0.5">Sơ cấp - Đất (MΩ)</label>
                    <input
                      type="number"
                      value={data.electrical_tests.insulation_resistance_1min_megohms.pri_to_ground}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1min_megohms: {
                            ...data.electrical_tests.insulation_resistance_1min_megohms,
                            pri_to_ground: parseFloat(e.target.value) || 0
                          }
                        }
                      })}
                      className="w-full px-2.5 py-1.5 border rounded-lg bg-white font-mono font-bold text-slate-900"
                    />
                  </div>
                  <div>
                    <label className="text-[10px] text-slate-400 block mb-0.5">Thứ cấp - Đất (MΩ)</label>
                    <input
                      type="number"
                      value={data.electrical_tests.insulation_resistance_1min_megohms.sec_to_ground}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1min_megohms: {
                            ...data.electrical_tests.insulation_resistance_1min_megohms,
                            sec_to_ground: parseFloat(e.target.value) || 0
                          }
                        }
                      })}
                      className="w-full px-2.5 py-1.5 border rounded-lg bg-white font-mono font-medium text-slate-800"
                    />
                  </div>
                  <div>
                    <label className="text-[10px] text-slate-400 block mb-0.5">Sơ cấp - Thứ cấp (MΩ)</label>
                    <input
                      type="number"
                      value={data.electrical_tests.insulation_resistance_1min_megohms.pri_to_sec}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1min_megohms: {
                            ...data.electrical_tests.insulation_resistance_1min_megohms,
                            pri_to_sec: parseFloat(e.target.value) || 0
                          }
                        }
                      })}
                      className="w-full px-2.5 py-1.5 border rounded-lg bg-white font-mono font-medium text-slate-800"
                    />
                  </div>
                  <div>
                    <label className="text-[10px] text-slate-400 block mb-0.5">Chỉ số phân cực (PI)</label>
                    <input
                      type="number"
                      step="0.01"
                      value={data.electrical_tests.pi_value}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: { ...data.electrical_tests, pi_value: parseFloat(e.target.value) || 0 }
                      })}
                      className={`w-full px-2.5 py-1.5 border rounded-lg bg-white font-mono font-bold ${liveStats.piPass ? 'text-emerald-700' : 'text-rose-700'}`}
                    />
                  </div>
                </div>

                {/* Trend Comparison with previous measurement */}
                <div className="flex items-center justify-between p-2 rounded-lg bg-amber-50/60 border border-amber-200 text-slate-700">
                  <div className="flex items-center gap-1.5">
                    <Activity size={14} className="text-amber-600" />
                    <span>Lần đo gần nhất (10/09/2025): <strong>{data.previous_test_data?.last_pri_ground_megohms} MΩ</strong></span>
                  </div>
                  <div className={`font-bold text-xs ${liveStats.irTrendWarning ? 'text-amber-700' : 'text-emerald-700'}`}>
                    {liveStats.irDrop > 0 ? `Sụt giảm: ${liveStats.irDrop}% (Ngưỡng cảnh báo > 20%)` : 'Không suy giảm'}
                  </div>
                </div>
              </div>

              {/* Row 3: Turns-Ratio & Winding Resistance */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/60 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-slate-800">Tỷ số biến áp (TTR - All Taps)</span>
                    <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${liveStats.ttrPass ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'}`}>
                      {liveStats.ttrPass ? 'ĐẠT (≤ 0.5%)' : 'LỖI (> 0.5%)'}
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-500">Sai số lớn nhất đo ở tất cả các nấc tap so với tính toán:</p>
                  <div className="flex items-center gap-2">
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
                      className="w-full px-2.5 py-1.5 border rounded-lg bg-white font-mono font-bold text-slate-900"
                    />
                    <span className="font-bold text-slate-600">%</span>
                  </div>
                </div>

                <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/60 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-slate-800">Điện trở cuộn dây (Winding Res.)</span>
                    <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${liveStats.wrPass ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'}`}>
                      {liveStats.wrPass ? 'ĐẠT (≤ 1.0%)' : `LỆCH ${liveStats.wrDev}%`}
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-500">Độ lệch quy đổi nhiệt độ so với xuất xưởng (chuẩn ≤ 1.0%):</p>
                  <div className="flex items-center gap-2">
                    <input
                      type="number"
                      step="0.1"
                      value={data.electrical_tests.winding_resistance_all_taps.max_deviation_from_factory_percent}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          winding_resistance_all_taps: { max_deviation_from_factory_percent: parseFloat(e.target.value) || 0 }
                        }
                      })}
                      className="w-full px-2.5 py-1.5 border rounded-lg bg-white font-mono font-bold text-slate-900"
                    />
                    <span className="font-bold text-slate-600">%</span>
                  </div>
                </div>
              </div>

              {/* Row 4: Excitation Current (Lõi 3 trụ) & Core Insulation Resistance */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                {/* Excitation Current */}
                <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/60 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-slate-800">Dòng kích thích (Excitation Current - mA)</span>
                    <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${liveStats.excPatternPass ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'}`}>
                      {liveStats.excPatternPass ? 'ĐÚNG MẪU (2 Cao 1 Thấp)' : 'LỆCH MẪU CHUẨN'}
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-500">Lõi 3 trụ: Phase A & C tương đương, Phase B thấp hơn</p>
                  <div className="grid grid-cols-3 gap-2">
                    <div>
                      <label className="text-[10px] text-slate-400 block mb-0.5">Pha A (mA)</label>
                      <input
                        type="number"
                        step="0.1"
                        value={data.electrical_tests.excitation_current_milliamp.phase_a}
                        onChange={(e) => setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            excitation_current_milliamp: {
                              ...data.electrical_tests.excitation_current_milliamp,
                              phase_a: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-2 py-1.5 border rounded-lg bg-white font-mono font-medium text-slate-900"
                      />
                    </div>
                    <div>
                      <label className="text-[10px] text-slate-400 block mb-0.5">Pha B (mA)</label>
                      <input
                        type="number"
                        step="0.1"
                        value={data.electrical_tests.excitation_current_milliamp.phase_b}
                        onChange={(e) => setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            excitation_current_milliamp: {
                              ...data.electrical_tests.excitation_current_milliamp,
                              phase_b: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-2 py-1.5 border rounded-lg bg-white font-mono font-medium text-slate-900"
                      />
                    </div>
                    <div>
                      <label className="text-[10px] text-slate-400 block mb-0.5">Pha C (mA)</label>
                      <input
                        type="number"
                        step="0.1"
                        value={data.electrical_tests.excitation_current_milliamp.phase_c}
                        onChange={(e) => setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            excitation_current_milliamp: {
                              ...data.electrical_tests.excitation_current_milliamp,
                              phase_c: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-2 py-1.5 border rounded-lg bg-white font-mono font-medium text-slate-900"
                      />
                    </div>
                  </div>
                </div>

                {/* Core Insulation Resistance */}
                <div className={`p-3.5 rounded-xl border space-y-2 ${liveStats.coreIrPass ? 'border-slate-200 bg-slate-50/60' : 'border-rose-300 bg-rose-50/60'}`}>
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-slate-800">Cách điện lõi thép (Core IR @ 500V DC)</span>
                    <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${liveStats.coreIrPass ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800 animate-pulse'}`}>
                      {liveStats.coreIrPass ? 'ĐẠT (≥ 1.0 MΩ)' : 'FAIL (< 1.0 MΩ)'}
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-500">NETA ATS-2025: Điện trở lõi thép tối thiểu phải đạt ≥ 1.0 MΩ</p>
                  <div className="flex items-center gap-2">
                    <input
                      type="number"
                      step="0.01"
                      value={data.electrical_tests.core_insulation_resistance_500v_dc_megohms}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          core_insulation_resistance_500v_dc_megohms: parseFloat(e.target.value) || 0
                        }
                      })}
                      className={`w-full px-2.5 py-1.5 border rounded-lg font-mono font-bold text-sm ${liveStats.coreIrPass ? 'bg-white text-slate-900' : 'bg-white border-rose-400 text-rose-700'}`}
                    />
                    <span className="font-bold text-slate-600">MΩ</span>
                  </div>
                  {!liveStats.coreIrPass && (
                    <div className="text-[11px] text-rose-700 font-semibold flex items-center gap-1">
                      <AlertTriangle size={12} />
                      Nguy cơ chạm đất lõi từ (core ground fault) hoặc ẩm bẩn cách điện lõi!
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* AI Analysis Result Section */}
      {analysisResult && (
        <div className="bg-white rounded-2xl border-2 border-indigo-200 p-6 shadow-xl space-y-6">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-200 pb-4">
            <div>
              <div className="flex items-center gap-2">
                <Sparkles size={20} className="text-amber-500" />
                <h2 className="text-lg font-black text-slate-900">
                  BÁO CÁO PHÂN TÍCH CHUYÊN GIA AI (TEV PLATFORM AI)
                </h2>
              </div>
              <p className="text-xs text-slate-500 mt-0.5">
                Tiêu chuẩn đối chiếu: ANSI/NETA ATS-2025 Mục 7.2.1.2 (Transformers, Dry-Type, Air-Cooled, Large)
              </p>
            </div>

            <div className="flex items-center gap-3">
              <span className={`px-4 py-1.5 rounded-full font-black text-xs uppercase tracking-wider border shadow-xs ${
                analysisResult.overall_status === 'PASS'
                  ? 'bg-emerald-100 text-emerald-800 border-emerald-300'
                  : analysisResult.overall_status === 'INVESTIGATE'
                  ? 'bg-amber-100 text-amber-800 border-amber-300'
                  : 'bg-rose-100 text-rose-800 border-rose-300'
              }`}>
                {analysisResult.overall_status === 'FAIL' ? '🔴 FAIL / INVESTIGATE' : analysisResult.overall_status}
              </span>

              <button
                onClick={() => setShowPrintModal(true)}
                className="px-3.5 py-1.5 bg-blue-600 hover:bg-blue-500 text-white text-xs font-semibold rounded-lg flex items-center gap-1.5 shadow-xs transition-colors"
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

          {/* Render Markdown report body */}
          <div className="prose prose-sm max-w-none text-slate-800 leading-relaxed font-sans bg-slate-50/50 p-6 rounded-xl border border-slate-200">
            <div
              className="space-y-4"
              dangerouslySetInnerHTML={{
                __html: analysisResult.markdown_report
                  .replace(/### (.*?)\n/g, '<h3 class="text-base font-bold text-slate-900 border-b border-slate-200 pb-1 mt-4 mb-2">$1</h3>')
                  .replace(/## (.*?)\n/g, '<h2 class="text-lg font-black text-blue-950 mt-4 mb-2">$1</h2>')
                  .replace(/\*\*(.*?)\*\*/g, '<strong class="font-bold text-slate-900">$1</strong>')
                  .replace(/\*(.*?)\*/g, '<em class="text-slate-700">$1</em>')
                  .replace(/\| (.*?) \|/g, (match) => {
                    return match;
                  })
                  .replace(/\n\n/g, '<br/><br/>')
              }}
            />
          </div>
        </div>
      )}

      {/* Printable TEV Report Modal */}
      <TevReportPrintModal
        isOpen={showPrintModal}
        onClose={() => setShowPrintModal(false)}
        title="BIÊN BẢN KIỂM ĐỊNH MÁY BIẾN ÁP KHÔ DUNG LƯỢNG LỚN"
        standard="ANSI/NETA ATS-2025 Section 7.2.1.2"
        siteInfo={data.site_info}
        overallStatus={analysisResult?.overall_status || (liveStats.coreIrPass && liveStats.ttrPass && liveStats.wrPass ? 'PASS' : 'INVESTIGATE')}
        statusReason={analysisResult?.status_reason || 'Đã đối chiếu với quy chuẩn kỹ thuật NETA ATS-2025 Mục 7.2.1.2.'}
        visualChecks={visualChecksForPrint}
        electricalTable={electricalTableForPrint}
        warnings={
          liveStats.coreIrPass
            ? []
            : [
                `Điện trở cách điện lõi thép (Core IR) đạt ${data.electrical_tests.core_insulation_resistance_500v_dc_megohms} MΩ < 1.0 MΩ tiêu chuẩn NETA ATS-2025 -> FAIL (Có hiện tượng chạm chập tiếp địa lõi thép hoặc bẩn/ẩm tại điểm cách điện lõi).`,
                `Điện trở cuộn dây lệch ${data.electrical_tests.winding_resistance_all_taps.max_deviation_from_factory_percent}% > 1.0% cho phép so với xuất xưởng -> INVESTIGATE.`,
                liveStats.irTrendWarning ? `Điện trở cách điện Sơ-Đất sụt giảm ${liveStats.irDrop}% so với lần đo trước (${data.previous_test_data?.last_pri_ground_megohms} MΩ).` : ''
              ].filter(Boolean)
        }
        recommendations={[
          'Kiểm tra và tháo thanh nối đất lõi thép (core ground strap) để vệ sinh/sấy khô vị trí cách điện lõi và đo lại Core IR ở 500V DC xem có khôi phục trên 1.0 MΩ hay không.',
          'Bôi mỡ tiếp xúc chuyên dụng và chuyển nấc tap nhiều lần để làm sạch bề mặt tiếp xúc trước khi đo lại điện trở cuộn dây.',
          'Kiểm tra kỹ tình trạng cách điện và bu-lông siết trước khi ký biên bản đóng điện.'
        ]}
      />

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
