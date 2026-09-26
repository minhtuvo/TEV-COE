import React, { useState, useMemo } from 'react';
import {
  BatteryCharging,
  Flame,
  ShieldCheck,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  Sparkles,
  Printer,
  RotateCcw,
  Upload,
  Download,
  Save,
  Activity,
  Zap,
  Gauge,
  Sliders,
  Compass,
  FileCheck2,
  Copy,
  Check
} from 'lucide-react';
import { NetaBessInputPayload, NetaBessEvaluationResult } from '../../server/netaBessAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

const SAMPLE_BESS_DATA: NetaBessInputPayload = {
  site_info: {
    project_name: "Tram Luu Tru Nang Luong BESS TEV 5MW",
    bess_container_tag: "BESS-CONT-01",
    fse_name: "Hoang Van E",
    test_date: "2026-09-22"
  },
  visual_inspection: {
    battery_location_clearance: "Pass",
    fire_suppression_installed: true,
    eyewash_station_present: true,
    nameplate_match: true,
    physical_condition: "Good (No leakage or damage)",
    rack_mounting_and_grounding: "Pass",
    cleanliness: "Pass",
    bolt_torque_check: "Pass"
  },
  electrical_and_subsystem_tests: {
    transformer_section_7_2_status: "Pass",
    lv_breakers_section_7_6_1_status: "Pass",
    lv_cables_section_7_3_2_status: "Pass",
    battery_system_polarity_correct: true,
    measured_total_dc_voltage_v: 815,
    expected_total_dc_voltage_v: 816,
    load_test_results: {
      discharge_power_kw: 1000,
      status: "Pass (Capacity meets manufacturer specs)"
    },
    power_quality_at_poi_ieee_1547: {
      voltage_thd_percent: 6.2,
      current_thd_percent: 2.1
    }
  },
  previous_test_data: {
    last_test_date: "2025-09-20",
    last_measured_dc_voltage_v: 816
  }
};

export const NetaBessChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaBessInputPayload>(SAMPLE_BESS_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaBessEvaluationResult | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [jsonInput, setJsonInput] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live Calculations
  const liveStats = useMemo(() => {
    // 1. Safety Mandatory
    const fireSafe = data.visual_inspection.fire_suppression_installed && data.visual_inspection.eyewash_station_present;

    // 2. DC Voltage Deviation
    const exp = data.electrical_and_subsystem_tests.expected_total_dc_voltage_v;
    const meas = data.electrical_and_subsystem_tests.measured_total_dc_voltage_v;
    const diff = Math.abs(meas - exp);
    const voltDev = exp > 0 ? parseFloat(((diff / exp) * 100).toFixed(2)) : 0;
    const voltDevPass = voltDev <= 1.5;

    // 3. Polarity
    const polarityPass = data.electrical_and_subsystem_tests.battery_system_polarity_correct;

    // 4. Power Quality IEEE 1547 (Limit < 5%)
    const vThd = data.electrical_and_subsystem_tests.power_quality_at_poi_ieee_1547.voltage_thd_percent;
    const iThd = data.electrical_and_subsystem_tests.power_quality_at_poi_ieee_1547.current_thd_percent;
    const vThdPass = vThd <= 5.0;
    const iThdPass = iThd <= 5.0;

    // 5. Subsystems
    const transPass = data.electrical_and_subsystem_tests.transformer_section_7_2_status.toLowerCase().includes('pass');
    const breakerPass = data.electrical_and_subsystem_tests.lv_breakers_section_7_6_1_status.toLowerCase().includes('pass');
    const cablePass = data.electrical_and_subsystem_tests.lv_cables_section_7_3_2_status.toLowerCase().includes('pass');
    const subPass = transPass && breakerPass && cablePass;

    return {
      fireSafe,
      voltDev,
      voltDevPass,
      polarityPass,
      vThd,
      iThd,
      vThdPass,
      iThdPass,
      subPass
    };
  }, [data]);

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    try {
      const response = await fetch('/api/field-service/neta-bess-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });
      const res = await response.json();
      if (res.success) {
        setAnalysisResult(res.data);
      } else {
        alert('Lỗi phân tích BESS: ' + res.error);
      }
    } catch (err: any) {
      alert('Không thể kết nối đến máy chủ: ' + err.message);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const loadSample = () => {
    setData(SAMPLE_BESS_DATA);
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
        equipmentId: data.site_info.bess_container_tag,
        date: data.site_info.test_date,
        type: 'Battery Energy Storage Systems (ANSI/NETA ATS-2025 Sec 7.28)'
      });
      setSavedSuccessMessage(`Đã lưu biên bản kiểm định BESS [${data.site_info.bess_container_tag}] vào kho báo cáo CMMS!`);
    }
  };

  const handleImportJson = () => {
    try {
      const parsed = JSON.parse(jsonInput);
      if (parsed.site_info && parsed.visual_inspection && parsed.electrical_and_subsystem_tests) {
        setData(parsed);
        setShowJsonModal(false);
        setAnalysisResult(null);
        alert('Đã nhập thành công dữ liệu kiểm định BESS từ JSON!');
      } else {
        alert('JSON không đúng định dạng chuẩn của BESS NETA ATS-2025.');
      }
    } catch (e: any) {
      alert('JSON không hợp lệ: ' + e.message);
    }
  };

  const handleCopyMarkdown = () => {
    if (analysisResult?.markdown_report) {
      navigator.clipboard.writeText(analysisResult.markdown_report);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }
  };

  // Prepare table data for print modal
  const electricalTableForPrint = [
    {
      item: 'Máy biến áp BESS (Transformer)',
      measured: `Trạng thái: ${data.electrical_and_subsystem_tests.transformer_section_7_2_status}`,
      standard: 'Tuân thủ toàn bộ NETA Section 7.2',
      reference: 'Sec 7.2 (Dry/Liquid)',
      deviation: 'Đạt cách điện & tỷ số',
      status: data.electrical_and_subsystem_tests.transformer_section_7_2_status.toLowerCase().includes('pass') ? 'PASS' as const : 'FAIL' as const
    },
    {
      item: 'Máy cắt hạ áp (LV Circuit Breakers)',
      measured: `Trạng thái: ${data.electrical_and_subsystem_tests.lv_breakers_section_7_6_1_status}`,
      standard: 'NETA Section 7.6.1 (7.6.1.1.1, 7.6.1.1.2, 7.6.1.2)',
      reference: 'Section 7.6.1',
      deviation: 'Đạt ngắt và tiếp xúc',
      status: data.electrical_and_subsystem_tests.lv_breakers_section_7_6_1_status.toLowerCase().includes('pass') ? 'PASS' as const : 'FAIL' as const
    },
    {
      item: 'Cáp kết nối hạ áp (LV Interconnecting Cables)',
      measured: `Trạng thái: ${data.electrical_and_subsystem_tests.lv_cables_section_7_3_2_status}`,
      standard: 'NETA Section 7.3.2 (Table 100.1)',
      reference: 'Section 7.3.2',
      deviation: 'Cách điện cáp DC/AC đạt chuẩn',
      status: data.electrical_and_subsystem_tests.lv_cables_section_7_3_2_status.toLowerCase().includes('pass') ? 'PASS' as const : 'FAIL' as const
    },
    {
      item: 'Cực tính hệ thống Pin (Polarity)',
      measured: data.electrical_and_subsystem_tests.battery_system_polarity_correct ? 'Chính xác (+/-)' : 'SAI CỰC TÍNH',
      standard: 'Khớp 100% tài liệu nhà sản xuất',
      reference: 'Bản vẽ đấu nối (+/-)',
      deviation: data.electrical_and_subsystem_tests.battery_system_polarity_correct ? 'Chính xác' : 'ĐẢO CỰC TÍNH NGUY HIỂM',
      status: data.electrical_and_subsystem_tests.battery_system_polarity_correct ? 'PASS' as const : 'FAIL' as const
    },
    {
      item: 'Điện áp DC tổng (Total DC Voltage)',
      measured: `${data.electrical_and_subsystem_tests.measured_total_dc_voltage_v} V DC`,
      standard: 'Lệch ≤ 1.5% so với định mức',
      reference: `Định mức: ${data.electrical_and_subsystem_tests.expected_total_dc_voltage_v} V`,
      deviation: `Lệch ${liveStats.voltDev}% (${data.electrical_and_subsystem_tests.measured_total_dc_voltage_v}V / ${data.electrical_and_subsystem_tests.expected_total_dc_voltage_v}V)`,
      status: liveStats.voltDevPass ? 'PASS' as const : 'INVESTIGATE' as const
    },
    {
      item: 'Thử nghiệm mang tải (Load Test)',
      measured: `${data.electrical_and_subsystem_tests.load_test_results.discharge_power_kw} kW (${data.electrical_and_subsystem_tests.load_test_results.status})`,
      standard: 'Đạt công suất nạp/xả & dung lượng thiết kế',
      reference: 'Thông số nhà sản xuất',
      deviation: 'Xả tải ổn định',
      status: 'PASS' as const
    },
    {
      item: 'Chất lượng điện năng tại POI (IEEE 1547)',
      measured: `THDu = ${data.electrical_and_subsystem_tests.power_quality_at_poi_ieee_1547.voltage_thd_percent}%, THDi = ${data.electrical_and_subsystem_tests.power_quality_at_poi_ieee_1547.current_thd_percent}%`,
      standard: 'Sóng hài điện áp THDu < 5.0%, THDi < 5.0%',
      reference: 'IEEE 1547 Standard',
      deviation: liveStats.vThdPass ? 'Đạt chuẩn hòa lưới' : `THDu = ${liveStats.vThd}% > 5.0% (Vượt ngưỡng)`,
      status: liveStats.vThdPass ? 'PASS' as const : 'INVESTIGATE' as const
    }
  ];

  const visualChecksForPrint = [
    { label: 'Vị trí đặt pin & Khoảng cách bảo trì', value: data.visual_inspection.battery_location_clearance, pass: data.visual_inspection.battery_location_clearance === 'Pass' },
    { label: 'Hệ thống PCCC chuyên dụng cho Container', value: data.visual_inspection.fire_suppression_installed ? 'ĐÃ LẮP ĐẶT' : 'CHƯA LẮP ĐẶT (FAIL)', pass: data.visual_inspection.fire_suppression_installed },
    { label: 'Trạm rửa mắt khẩn cấp (Eyewash station)', value: data.visual_inspection.eyewash_station_present ? 'ĐÃ TRANG BỊ' : 'THIẾU TRẠNG BỊ (FAIL)', pass: data.visual_inspection.eyewash_station_present },
    { label: 'Khớp nhãn Nameplate (Cell, Rack, PCS, TR)', value: data.visual_inspection.nameplate_match ? 'ĐẠT 100%' : 'KHÔNG KHỚP', pass: data.visual_inspection.nameplate_match },
    { label: 'Tình trạng cơ học & Rò rỉ điện dịch', value: data.visual_inspection.physical_condition, pass: data.visual_inspection.physical_condition.toLowerCase().includes('good') },
    { label: 'Neo giá đỡ (Racks) & Tiếp địa vỏ tủ', value: data.visual_inspection.rack_mounting_and_grounding, pass: data.visual_inspection.rack_mounting_and_grounding === 'Pass' },
    { label: 'Độ sạch sẽ khoang pin & PCS', value: data.visual_inspection.cleanliness, pass: data.visual_inspection.cleanliness === 'Pass' },
    { label: 'Lực siết bu-lông (Bolt Torque NSX)', value: data.visual_inspection.bolt_torque_check, pass: data.visual_inspection.bolt_torque_check === 'Pass' }
  ];

  return (
    <div className="space-y-6">
      {/* Top Banner & Action Header */}
      <div className="bg-gradient-to-r from-emerald-950 via-teal-950 to-slate-900 text-white p-5 sm:p-6 rounded-2xl shadow-lg border border-teal-800/40">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2 mb-1.5">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-teal-400/20 text-teal-300 border border-teal-400/30 uppercase tracking-wide flex items-center gap-1">
                <BatteryCharging size={13} />
                ANSI/NETA ATS-2025 Section 7.28
              </span>
              <span className="text-xs text-slate-300">Container BESS • PCS • Lõi Pin Lithium / Flow • IEEE 1547</span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white flex items-center gap-2.5">
              <BatteryCharging className="text-teal-400 fill-teal-400/20" size={24} />
              Checklist Hệ Thống Pin Lưu Trữ Năng Lượng (BESS)
            </h1>
            <p className="text-xs sm:text-sm text-slate-300 mt-1 max-w-2xl">
              Nghiệm thu toàn diện: An toàn PCCC container & Trạm rửa mắt, Thử nghiệm phân hệ (MBA Sec 7.2, Máy cắt Sec 7.6.1, Cáp Sec 7.3.2), Cực tính/Điện áp DC và Sóng hài hòa lưới IEEE 1547.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={loadSample}
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title="Tải kịch bản mẫu Container BESS 5MW"
            >
              <RotateCcw size={14} />
              <span>Nạp Kịch Bản Mẫu FSE</span>
            </button>

            <button
              onClick={() => setShowJsonModal(true)}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
            >
              <Upload size={14} />
              <span>Nhập JSON</span>
            </button>

            <button
              onClick={() => {
                const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `NETA_ATS_2025_BESS_${data.site_info.bess_container_tag || 'CONT'}.json`;
                a.click();
              }}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
            >
              <Download size={14} />
              <span>Xuất JSON</span>
            </button>

            <button
              onClick={() => setShowPrintModal(true)}
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-teal-700 hover:bg-teal-600 text-white border border-teal-500/40 flex items-center gap-1.5 transition-all shadow-md"
            >
              <Printer size={15} />
              <span>In / Xuất Báo Cáo PDF (TEV)</span>
            </button>

            <button
              onClick={handleAnalyze}
              disabled={isAnalyzing}
              className="px-4 py-2 rounded-xl text-xs font-bold bg-gradient-to-r from-teal-400 to-emerald-400 hover:from-teal-300 hover:to-emerald-300 text-slate-950 flex items-center gap-2 transition-all shadow-lg shadow-teal-500/20 disabled:opacity-50"
            >
              <Sparkles size={16} className={isAnalyzing ? 'animate-spin' : ''} />
              <span>{isAnalyzing ? 'Đang Phân Tích...' : '🤖 Phân Tích AI Chuyên Gia'}</span>
            </button>
          </div>
        </div>

        {/* Live Status Indicators Bar */}
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2.5 mt-5 pt-4 border-t border-white/10 text-xs">
          <div className="bg-white/5 rounded-lg p-2.5 border border-white/10">
            <div className="text-[10px] text-slate-400 flex items-center gap-1">
              <Flame size={12} className={liveStats.fireSafe ? 'text-teal-400' : 'text-rose-400'} />
              <span>PCCC & Rửa Mắt</span>
            </div>
            <div className={`font-bold text-xs mt-0.5 ${liveStats.fireSafe ? 'text-teal-300' : 'text-rose-400 animate-pulse'}`}>
              {liveStats.fireSafe ? 'ĐẠT (BẮT BUỘC)' : 'THIẾU AN TOÀN'}
            </div>
            <div className="text-[9px] text-slate-400">NETA Sec 7.28.1</div>
          </div>

          <div className="bg-white/5 rounded-lg p-2.5 border border-white/10">
            <div className="text-[10px] text-slate-400">Điện áp DC Tổng</div>
            <div className={`font-mono font-bold text-sm mt-0.5 ${liveStats.voltDevPass ? 'text-teal-300' : 'text-amber-400'}`}>
              {data.electrical_and_subsystem_tests.measured_total_dc_voltage_v} V
            </div>
            <div className="text-[9px] text-slate-400">Định mức: {data.electrical_and_subsystem_tests.expected_total_dc_voltage_v}V (Lệch {liveStats.voltDev}%)</div>
          </div>

          <div className="bg-white/5 rounded-lg p-2.5 border border-white/10">
            <div className="text-[10px] text-slate-400">Cực tính (+/-)</div>
            <div className={`font-bold text-xs mt-0.5 ${liveStats.polarityPass ? 'text-emerald-400' : 'text-rose-400 animate-bounce'}`}>
              {liveStats.polarityPass ? 'CHÍNH XÁC (+/-)' : 'SAI CỰC TÍNH (NGUY HIỂM)'}
            </div>
            <div className="text-[9px] text-slate-400">Bảo vệ Cell/PCS</div>
          </div>

          <div className="bg-white/5 rounded-lg p-2.5 border border-white/10">
            <div className="text-[10px] text-slate-400">Sóng Hài THDu (POI)</div>
            <div className={`font-mono font-bold text-sm mt-0.5 ${liveStats.vThdPass ? 'text-teal-300' : 'text-amber-400'}`}>
              {liveStats.vThd}%
            </div>
            <div className="text-[9px] text-slate-400">IEEE 1547 (Chuẩn &lt; 5.0%)</div>
          </div>

          <div className="bg-white/5 rounded-lg p-2.5 border border-white/10">
            <div className="text-[10px] text-slate-400">Sóng Hài THDi (POI)</div>
            <div className={`font-mono font-bold text-sm mt-0.5 ${liveStats.iThdPass ? 'text-teal-300' : 'text-amber-400'}`}>
              {liveStats.iThd}%
            </div>
            <div className="text-[9px] text-slate-400">IEEE 1547 (Chuẩn &lt; 5.0%)</div>
          </div>

          <div className="bg-white/5 rounded-lg p-2.5 border border-white/10">
            <div className="text-[10px] text-slate-400">Phân hệ Hạ tầng</div>
            <div className={`font-bold text-xs mt-0.5 ${liveStats.subPass ? 'text-teal-300' : 'text-rose-400'}`}>
              {liveStats.subPass ? 'ĐẠT (MBA/CB/Cáp)' : 'CÓ LỖI'}
            </div>
            <div className="text-[9px] text-slate-400">Sec 7.2 / 7.6.1 / 7.3.2</div>
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
        {/* Left Column: Site Info & Visual / Safety Inspection */}
        <div className="space-y-6 lg:col-span-1">
          {/* Site Info */}
          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs space-y-4">
            <h3 className="font-bold text-sm text-slate-900 flex items-center gap-2 border-b border-slate-100 pb-2">
              <Compass size={16} className="text-teal-600" />
              Thông tin Container BESS & Trạm Site
            </h3>
            <div className="space-y-3 text-xs">
              <div>
                <label className="block text-slate-500 mb-1">Dự án BESS</label>
                <input
                  type="text"
                  value={data.site_info.project_name}
                  onChange={(e) => setData({ ...data, site_info: { ...data.site_info, project_name: e.target.value } })}
                  className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white focus:ring-2 focus:ring-teal-500 font-medium text-slate-800"
                />
              </div>
              <div>
                <label className="block text-slate-500 mb-1">Mã Tag Container BESS</label>
                <input
                  type="text"
                  value={data.site_info.bess_container_tag}
                  onChange={(e) => setData({ ...data, site_info: { ...data.site_info, bess_container_tag: e.target.value } })}
                  className="w-full px-3 py-2 border rounded-lg bg-slate-50 focus:bg-white focus:ring-2 focus:ring-teal-500 font-mono font-bold text-teal-900"
                />
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-slate-500 mb-1">Kỹ sư FSE</label>
                  <input
                    type="text"
                    value={data.site_info.fse_name}
                    onChange={(e) => setData({ ...data, site_info: { ...data.site_info, fse_name: e.target.value } })}
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

          {/* Visual & Mechanical Inspection (NETA 7.28.1) */}
          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs space-y-4">
            <h3 className="font-bold text-sm text-slate-900 flex items-center gap-2 border-b border-slate-100 pb-2">
              <ShieldCheck size={16} className="text-teal-600" />
              1. Kiểm Tra Thị Giác & An Toàn BESS
            </h3>

            <div className="space-y-3 text-xs">
              <div className="flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-100">
                <div>
                  <div className="font-semibold text-slate-800">Khoảng cách & Vị trí đặt pin</div>
                  <div className="text-[10px] text-slate-400">Thông thoáng, đúng hành lang an toàn</div>
                </div>
                <select
                  value={data.visual_inspection.battery_location_clearance}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, battery_location_clearance: e.target.value }
                  })}
                  className="px-2 py-1 border rounded bg-white font-medium"
                >
                  <option value="Pass">Pass (Đạt)</option>
                  <option value="Fail">Fail (Lỗi)</option>
                </select>
              </div>

              {/* Fire Suppression (Mandatory) */}
              <div className={`flex items-center justify-between p-2.5 rounded-lg border ${data.visual_inspection.fire_suppression_installed ? 'bg-teal-50/60 border-teal-200' : 'bg-rose-50 border-rose-300'}`}>
                <div>
                  <div className="font-bold text-slate-900 flex items-center gap-1.5">
                    <Flame size={14} className={data.visual_inspection.fire_suppression_installed ? 'text-teal-600' : 'text-rose-600'} />
                    <span>Hệ thống PCCC chuyên dụng</span>
                  </div>
                  <div className="text-[10px] text-slate-500">BẮT BUỘC có PCCC khí/aerosol cho pin</div>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.fire_suppression_installed}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, fire_suppression_installed: e.target.checked }
                  })}
                  className="w-4 h-4 text-teal-600 rounded"
                />
              </div>

              {/* Eyewash Station (Mandatory) */}
              <div className={`flex items-center justify-between p-2.5 rounded-lg border ${data.visual_inspection.eyewash_station_present ? 'bg-teal-50/60 border-teal-200' : 'bg-rose-50 border-rose-300'}`}>
                <div>
                  <div className="font-bold text-slate-900">Trạm rửa mắt khẩn cấp (Eyewash)</div>
                  <div className="text-[10px] text-slate-500">BẮT BUỘC cho khu vực lưu trữ pin</div>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.eyewash_station_present}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, eyewash_station_present: e.target.checked }
                  })}
                  className="w-4 h-4 text-teal-600 rounded"
                />
              </div>

              <div className="flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-100">
                <div>
                  <div className="font-semibold text-slate-800">Khớp nhãn Nameplate</div>
                  <div className="text-[10px] text-slate-400">Cell, Rack, PCS, MBA khớp 100% bản vẽ</div>
                </div>
                <input
                  type="checkbox"
                  checked={data.visual_inspection.nameplate_match}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, nameplate_match: e.target.checked }
                  })}
                  className="w-4 h-4 text-teal-600 rounded"
                />
              </div>

              <div className="flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-100">
                <div>
                  <div className="font-semibold text-slate-800">Tình trạng cơ học & Rò rỉ</div>
                  <div className="text-[10px] text-slate-400">Không nứt vỡ, không rò rỉ điện dịch</div>
                </div>
                <input
                  type="text"
                  value={data.visual_inspection.physical_condition}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, physical_condition: e.target.value }
                  })}
                  className="w-44 px-2 py-1 border rounded bg-white font-medium text-right"
                />
              </div>

              <div className="flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-100">
                <div>
                  <div className="font-semibold text-slate-800">Neo khung giá đỡ (Racks) & Tiếp địa</div>
                  <div className="text-[10px] text-slate-400">Bắt bu-lông sàn chắc chắn & tiếp địa vỏ</div>
                </div>
                <select
                  value={data.visual_inspection.rack_mounting_and_grounding}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, rack_mounting_and_grounding: e.target.value }
                  })}
                  className="px-2 py-1 border rounded bg-white font-medium"
                >
                  <option value="Pass">Pass (Đạt)</option>
                  <option value="Fail">Fail (Lỗi)</option>
                </select>
              </div>

              <div className="flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-100">
                <div>
                  <div className="font-semibold text-slate-800">Độ sạch sẽ khoang tủ</div>
                  <div className="text-[10px] text-slate-400">Không bụi kim loại, không đọng ẩm</div>
                </div>
                <select
                  value={data.visual_inspection.cleanliness}
                  onChange={(e) => setData({
                    ...data,
                    visual_inspection: { ...data.visual_inspection, cleanliness: e.target.value }
                  })}
                  className="px-2 py-1 border rounded bg-white font-medium"
                >
                  <option value="Pass">Pass (Đạt)</option>
                  <option value="Fail">Fail (Lỗi)</option>
                </select>
              </div>

              <div className="flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-100">
                <div>
                  <div className="font-semibold text-slate-800">Lực siết bu-lông mối nối (Bolt Torque)</div>
                  <div className="text-[10px] text-slate-400">Theo thông số xuất bản của NSX</div>
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
            </div>
          </div>
        </div>

        {/* Center & Right Column: Subsystems & Electrical Tests (NETA 7.28.2) */}
        <div className="space-y-6 lg:col-span-2">
          <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-xs space-y-4">
            <h3 className="font-bold text-sm text-slate-900 flex items-center justify-between border-b border-slate-100 pb-2">
              <span className="flex items-center gap-2">
                <Gauge size={16} className="text-teal-600" />
                2. Phép Đo Điện & Phân Hệ BESS (NETA ATS-2025 Mục 7.28.2)
              </span>
              <span className="text-[11px] text-slate-400 font-normal">Tiêu chuẩn NETA & IEEE 1547</span>
            </h3>

            {/* Subsystems Status Table */}
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/60 space-y-3 text-xs">
              <div className="flex items-center justify-between">
                <div>
                  <span className="font-bold text-slate-800">Thử nghiệm các phân hệ hạ tầng tích hợp</span>
                  <span className="text-[11px] text-slate-500 ml-2">NETA ATS-2025 Sec 7.2, 7.6.1, 7.3.2</span>
                </div>
                <span className={`px-2 py-0.5 rounded font-bold text-[10px] ${liveStats.subPass ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'}`}>
                  {liveStats.subPass ? 'TẤT CẢ ĐẠT' : 'CÓ PHÂN HỆ LỖI'}
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div className="p-2.5 rounded-lg bg-white border border-slate-200 space-y-1">
                  <div className="text-slate-500 font-medium">Máy biến áp (Section 7.2)</div>
                  <select
                    value={data.electrical_and_subsystem_tests.transformer_section_7_2_status}
                    onChange={(e) => setData({
                      ...data,
                      electrical_and_subsystem_tests: {
                        ...data.electrical_and_subsystem_tests,
                        transformer_section_7_2_status: e.target.value
                      }
                    })}
                    className="w-full px-2 py-1 border rounded bg-slate-50 font-bold text-slate-800"
                  >
                    <option value="Pass">Pass (Đạt tiêu chuẩn)</option>
                    <option value="Fail">Fail (Không đạt)</option>
                    <option value="N/A">N/A</option>
                  </select>
                </div>

                <div className="p-2.5 rounded-lg bg-white border border-slate-200 space-y-1">
                  <div className="text-slate-500 font-medium">Máy cắt hạ áp (Sec 7.6.1)</div>
                  <select
                    value={data.electrical_and_subsystem_tests.lv_breakers_section_7_6_1_status}
                    onChange={(e) => setData({
                      ...data,
                      electrical_and_subsystem_tests: {
                        ...data.electrical_and_subsystem_tests,
                        lv_breakers_section_7_6_1_status: e.target.value
                      }
                    })}
                    className="w-full px-2 py-1 border rounded bg-slate-50 font-bold text-slate-800"
                  >
                    <option value="Pass">Pass (Đạt tiêu chuẩn)</option>
                    <option value="Fail">Fail (Không đạt)</option>
                  </select>
                </div>

                <div className="p-2.5 rounded-lg bg-white border border-slate-200 space-y-1">
                  <div className="text-slate-500 font-medium">Cáp liên kết hạ áp (Sec 7.3.2)</div>
                  <select
                    value={data.electrical_and_subsystem_tests.lv_cables_section_7_3_2_status}
                    onChange={(e) => setData({
                      ...data,
                      electrical_and_subsystem_tests: {
                        ...data.electrical_and_subsystem_tests,
                        lv_cables_section_7_3_2_status: e.target.value
                      }
                    })}
                    className="w-full px-2 py-1 border rounded bg-slate-50 font-bold text-slate-800"
                  >
                    <option value="Pass">Pass (Đạt tiêu chuẩn)</option>
                    <option value="Fail">Fail (Không đạt)</option>
                  </select>
                </div>
              </div>
            </div>

            {/* DC Voltage & Polarity */}
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/60 space-y-3 text-xs">
              <div className="flex items-center justify-between">
                <div>
                  <span className="font-bold text-slate-800">Điện áp & Cực tính chuỗi Pin (DC Voltage & Polarity)</span>
                  <span className="text-[11px] text-slate-500 ml-2">NETA: Cực tính chính xác, điện áp khớp định mức NSX</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <input
                    type="checkbox"
                    id="polarity_check"
                    checked={data.electrical_and_subsystem_tests.battery_system_polarity_correct}
                    onChange={(e) => setData({
                      ...data,
                      electrical_and_subsystem_tests: {
                        ...data.electrical_and_subsystem_tests,
                        battery_system_polarity_correct: e.target.checked
                      }
                    })}
                    className="w-4 h-4 text-teal-600 rounded"
                  />
                  <label htmlFor="polarity_check" className="font-bold text-slate-700 cursor-pointer">
                    Cực tính (+/-) chính xác
                  </label>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div>
                  <label className="text-[10px] text-slate-400 block mb-0.5">Điện áp DC thực đo (V)</label>
                  <input
                    type="number"
                    value={data.electrical_and_subsystem_tests.measured_total_dc_voltage_v}
                    onChange={(e) => setData({
                      ...data,
                      electrical_and_subsystem_tests: {
                        ...data.electrical_and_subsystem_tests,
                        measured_total_dc_voltage_v: parseFloat(e.target.value) || 0
                      }
                    })}
                    className="w-full px-2.5 py-1.5 border rounded-lg bg-white font-mono font-bold text-slate-900"
                  />
                </div>
                <div>
                  <label className="text-[10px] text-slate-400 block mb-0.5">Điện áp định mức thiết kế (V)</label>
                  <input
                    type="number"
                    value={data.electrical_and_subsystem_tests.expected_total_dc_voltage_v}
                    onChange={(e) => setData({
                      ...data,
                      electrical_and_subsystem_tests: {
                        ...data.electrical_and_subsystem_tests,
                        expected_total_dc_voltage_v: parseFloat(e.target.value) || 0
                      }
                    })}
                    className="w-full px-2.5 py-1.5 border rounded-lg bg-white font-mono font-medium text-slate-800"
                  />
                </div>
                <div>
                  <label className="text-[10px] text-slate-400 block mb-0.5">Lần đo trước (20/09/2025)</label>
                  <input
                    type="number"
                    value={data.previous_test_data?.last_measured_dc_voltage_v || 0}
                    onChange={(e) => setData({
                      ...data,
                      previous_test_data: {
                        ...data.previous_test_data,
                        last_measured_dc_voltage_v: parseFloat(e.target.value) || 0
                      }
                    })}
                    className="w-full px-2.5 py-1.5 border rounded-lg bg-slate-100 font-mono text-slate-600"
                  />
                </div>
              </div>

              <div className="flex items-center justify-between p-2 rounded-lg bg-teal-50/70 border border-teal-200 text-teal-900">
                <span className="font-semibold">Độ lệch điện áp đo so với thiết kế: <strong>{liveStats.voltDev}%</strong></span>
                <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${liveStats.voltDevPass ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'}`}>
                  {liveStats.voltDevPass ? 'ĐẠT (≤ 1.5%)' : 'LỆCH CAO'}
                </span>
              </div>
            </div>

            {/* Load Test & Power Quality at POI (IEEE 1547) */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs">
              {/* Load Test */}
              <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/60 space-y-2">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-slate-800">Thử nghiệm mang tải (Load Test)</span>
                  <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800">
                    PASS
                  </span>
                </div>
                <div>
                  <label className="text-[10px] text-slate-400 block mb-0.5">Công suất xả thử nghiệm (kW)</label>
                  <input
                    type="number"
                    value={data.electrical_and_subsystem_tests.load_test_results.discharge_power_kw}
                    onChange={(e) => setData({
                      ...data,
                      electrical_and_subsystem_tests: {
                        ...data.electrical_and_subsystem_tests,
                        load_test_results: {
                          ...data.electrical_and_subsystem_tests.load_test_results,
                          discharge_power_kw: parseFloat(e.target.value) || 0
                        }
                      }
                    })}
                    className="w-full px-2.5 py-1.5 border rounded-lg bg-white font-mono font-bold text-slate-900"
                  />
                </div>
                <div>
                  <label className="text-[10px] text-slate-400 block mb-0.5">Trạng thái dung lượng & BMS</label>
                  <input
                    type="text"
                    value={data.electrical_and_subsystem_tests.load_test_results.status}
                    onChange={(e) => setData({
                      ...data,
                      electrical_and_subsystem_tests: {
                        ...data.electrical_and_subsystem_tests,
                        load_test_results: {
                          ...data.electrical_and_subsystem_tests.load_test_results,
                          status: e.target.value
                        }
                      }
                    })}
                    className="w-full px-2.5 py-1.5 border rounded-lg bg-white text-slate-800"
                  />
                </div>
              </div>

              {/* Power Quality at MV POI (IEEE 1547) */}
              <div className={`p-3.5 rounded-xl border space-y-2 ${liveStats.vThdPass ? 'border-slate-200 bg-slate-50/60' : 'border-amber-300 bg-amber-50/60'}`}>
                <div className="flex items-center justify-between">
                  <span className="font-bold text-slate-800">Chất lượng điện tại POI (IEEE 1547)</span>
                  <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${liveStats.vThdPass ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'}`}>
                    {liveStats.vThdPass ? 'ĐẠT (THD < 5%)' : 'INVESTIGATE'}
                  </span>
                </div>
                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="text-[10px] text-slate-400 block mb-0.5">Sóng hài áp THDu (%)</label>
                    <input
                      type="number"
                      step="0.1"
                      value={data.electrical_and_subsystem_tests.power_quality_at_poi_ieee_1547.voltage_thd_percent}
                      onChange={(e) => setData({
                        ...data,
                        electrical_and_subsystem_tests: {
                          ...data.electrical_and_subsystem_tests,
                          power_quality_at_poi_ieee_1547: {
                            ...data.electrical_and_subsystem_tests.power_quality_at_poi_ieee_1547,
                            voltage_thd_percent: parseFloat(e.target.value) || 0
                          }
                        }
                      })}
                      className={`w-full px-2.5 py-1.5 border rounded-lg font-mono font-bold ${liveStats.vThdPass ? 'bg-white text-slate-900' : 'bg-white border-amber-400 text-amber-900'}`}
                    />
                  </div>
                  <div>
                    <label className="text-[10px] text-slate-400 block mb-0.5">Sóng hài dòng THDi (%)</label>
                    <input
                      type="number"
                      step="0.1"
                      value={data.electrical_and_subsystem_tests.power_quality_at_poi_ieee_1547.current_thd_percent}
                      onChange={(e) => setData({
                        ...data,
                        electrical_and_subsystem_tests: {
                          ...data.electrical_and_subsystem_tests,
                          power_quality_at_poi_ieee_1547: {
                            ...data.electrical_and_subsystem_tests.power_quality_at_poi_ieee_1547,
                            current_thd_percent: parseFloat(e.target.value) || 0
                          }
                        }
                      })}
                      className="w-full px-2.5 py-1.5 border rounded-lg bg-white font-mono font-bold text-slate-900"
                    />
                  </div>
                </div>
                {!liveStats.vThdPass && (
                  <div className="text-[11px] text-amber-800 font-semibold flex items-center gap-1 mt-1">
                    <AlertTriangle size={12} />
                    THDu = {liveStats.vThd}% vượt ngưỡng 5.0% tiêu chuẩn IEEE 1547!
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* AI Analysis Result Section */}
      {analysisResult && (
        <div className="bg-white rounded-2xl border-2 border-teal-200 p-6 shadow-xl space-y-6">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-200 pb-4">
            <div>
              <div className="flex items-center gap-2">
                <Sparkles size={20} className="text-teal-600" />
                <h2 className="text-lg font-black text-slate-900">
                  BÁO CÁO PHÂN TÍCH CHUYÊN GIA AI BESS (TEV PLATFORM AI)
                </h2>
              </div>
              <p className="text-xs text-slate-500 mt-0.5">
                Tiêu chuẩn đối chiếu: ANSI/NETA ATS-2025 Mục 7.28 (Battery Energy Storage Systems) & IEEE 1547
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-2.5">
              <span className={`px-4 py-1.5 rounded-full font-black text-xs uppercase tracking-wider border shadow-xs ${
                analysisResult.overall_status === 'PASS'
                  ? 'bg-emerald-100 text-emerald-800 border-emerald-300'
                  : analysisResult.overall_status === 'INVESTIGATE'
                  ? 'bg-amber-100 text-amber-800 border-amber-300'
                  : 'bg-rose-100 text-rose-800 border-rose-300'
              }`}>
                {analysisResult.overall_status === 'FAIL' ? '🔴 FAIL' : analysisResult.overall_status}
              </span>

              <button
                onClick={handleCopyMarkdown}
                className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-lg flex items-center gap-1 border border-slate-300 transition-colors"
              >
                {copied ? <Check size={14} className="text-emerald-600" /> : <Copy size={14} />}
                <span>{copied ? 'Đã chép' : 'Sao chép Markdown'}</span>
              </button>

              <button
                onClick={() => setShowPrintModal(true)}
                className="px-3.5 py-1.5 bg-teal-700 hover:bg-teal-600 text-white text-xs font-semibold rounded-lg flex items-center gap-1.5 shadow-xs transition-colors"
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
                  .replace(/## (.*?)\n/g, '<h2 class="text-lg font-black text-teal-950 mt-4 mb-2">$1</h2>')
                  .replace(/\*\*(.*?)\*\*/g, '<strong class="font-bold text-slate-900">$1</strong>')
                  .replace(/\*(.*?)\*/g, '<em class="text-slate-700">$1</em>')
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
        title="BIÊN BẢN KIỂM ĐỊNH HỆ THỐNG PIN LƯU TRỮ NĂNG LƯỢNG (BESS)"
        standard="ANSI/NETA ATS-2025 Section 7.28 & IEEE 1547"
        siteInfo={{
          project_name: data.site_info.project_name,
          transformer_tag: data.site_info.bess_container_tag,
          fse_name: data.site_info.fse_name,
          test_date: data.site_info.test_date
        }}
        overallStatus={analysisResult?.overall_status || (liveStats.fireSafe && liveStats.vThdPass ? 'PASS' : 'INVESTIGATE')}
        statusReason={analysisResult?.status_reason || 'Đã đối chiếu với quy chuẩn kỹ thuật NETA ATS-2025 Mục 7.28 và IEEE 1547.'}
        visualChecks={visualChecksForPrint}
        electricalTable={electricalTableForPrint}
        warnings={
          liveStats.vThdPass
            ? []
            : [
                `Sóng hài điện áp THDu = ${liveStats.vThd}% vượt quá ngưỡng 5.0% tiêu chuẩn IEEE 1547 -> INVESTIGATE.`,
                !liveStats.fireSafe ? 'Chưa lắp đặt đủ hệ thống PCCC hoặc thiếu trạm rửa mắt khẩn cấp cho Container BESS -> FAIL.' : ''
              ].filter(Boolean)
        }
        recommendations={[
          'Kiểm tra bộ lọc bộ biến đổi công suất (PCS/Inverter Harmonic Filters) và tụ lọc đầu ra AC.',
          'Cấu hình lại tham số điều khiển vòng lặp dòng điện (Current Control Loop) của Inverter để giảm độ méo sóng hài điện áp THDu tại POI xuống dưới 5.0% trước khi hòa lưới chính thức.',
          'Giám sát độ lệch điện áp giữa các Cell (Cell Delta Voltage < 20mV) và nhiệt độ khối pin trong chu trình xả tải đầy công suất.'
        ]}
      />

      {/* JSON Import Modal */}
      {showJsonModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs">
          <div className="bg-white rounded-2xl p-6 max-w-lg w-full space-y-4 shadow-2xl">
            <h3 className="font-bold text-slate-900 text-sm">Nhập Dữ Liệu Kiểm Định BESS JSON</h3>
            <textarea
              rows={10}
              value={jsonInput}
              onChange={(e) => setJsonInput(e.target.value)}
              placeholder="Dán mã JSON kiểm định BESS NETA ATS-2025..."
              className="w-full p-3 font-mono text-xs border rounded-lg bg-slate-50 focus:bg-white"
            />
            <div className="flex justify-end gap-2">
              <button
                onClick={() => setShowJsonModal(false)}
                className="px-4 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-lg"
              >
                Hủy
              </button>
              <button
                onClick={handleImportJson}
                className="px-4 py-2 text-xs font-bold bg-teal-600 hover:bg-teal-500 text-white rounded-lg"
              >
                Nhập Dữ Liệu
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
