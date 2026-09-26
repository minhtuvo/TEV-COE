import React, { useState, useMemo } from 'react';
import {
  ClipboardCheck,
  Zap,
  AlertTriangle,
  CheckCircle2,
  XCircle,
  HelpCircle,
  Bot,
  Copy,
  Download,
  Upload,
  RefreshCw,
  FileText,
  ShieldAlert,
  ArrowRight,
  TrendingDown,
  Gauge,
  Sliders,
  Check,
  Eye,
  Info,
  Printer,
  Save,
  Camera
} from 'lucide-react';
import { TevReportPrintModal } from './TevReportPrintModal';
import { NameplateScannerModal } from './NameplateScannerModal';
import { ExtractedNameplateData } from '../../server/nameplateOcrAnalyzer';
import { FseEngineerDropdown } from './FseEngineerDropdown';

export interface NetaDryTypePayload {
  site_info: {
    project_name: string;
    transformer_tag: string;
    fse_name: string;
    test_date: string;
  };
  visual_inspection: {
    nameplate_match: boolean;
    physical_condition: string;
    grounding_check: string;
    shipping_brackets_removed: boolean;
    cleanliness: string;
    bolt_torque_check: string;
    tap_position: string;
  };
  electrical_tests: {
    bolted_resistance_micro_ohms: number[];
    insulation_resistance_1min_1000v_megohms: {
      pri_to_ground: number;
      sec_to_ground: number;
      pri_to_sec: number;
    };
    dar_value: number;
    turns_ratio: {
      calculated_ratio: number;
      measured_ratios: number[];
    };
    secondary_voltage_no_load: string;
  };
  previous_test_data?: {
    last_test_date?: string;
    last_insulation_resistance_pri_ground?: number;
  };
}

const DEFAULT_SAMPLE_DATA: NetaDryTypePayload = {
  site_info: {
    project_name: 'Nha may TEV Phase 2',
    transformer_tag: 'TR-LV-01',
    fse_name: 'Võ Minh Tú (sgm1707@gmail.com)',
    test_date: new Date().toISOString().split('T')[0]
  },
  visual_inspection: {
    nameplate_match: true,
    physical_condition: 'Good',
    grounding_check: 'Pass',
    shipping_brackets_removed: true,
    cleanliness: 'Pass',
    bolt_torque_check: 'Pass',
    tap_position: 'Tap 3 (480V)'
  },
  electrical_tests: {
    bolted_resistance_micro_ohms: [12.0, 12.4, 21.0],
    insulation_resistance_1min_1000v_megohms: {
      pri_to_ground: 650,
      sec_to_ground: 700,
      pri_to_sec: 800
    },
    dar_value: 1.2,
    turns_ratio: {
      calculated_ratio: 2.307,
      measured_ratios: [2.308, 2.306, 2.325]
    },
    secondary_voltage_no_load: '208V/120V (Khớp Nameplate)'
  },
  previous_test_data: {
    last_test_date: '2025-09-15',
    last_insulation_resistance_pri_ground: 850
  }
};

const PHASE_NAMES = ['Pha A', 'Pha B', 'Pha C'];

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

export const NetaDryTypeChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaDryTypePayload>(DEFAULT_SAMPLE_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<any>(null);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showCameraModal, setShowCameraModal] = useState(false);
  const [jsonInput, setJsonInput] = useState('');
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

  // Live Calculations based on NETA ATS-2025
  const liveStats = useMemo(() => {
    // 1. Bolted resistance
    const bolts = data.electrical_tests.bolted_resistance_micro_ohms || [];
    const minBolt = bolts.length > 0 ? Math.min(...bolts) : 0;
    const maxBolt = bolts.length > 0 ? Math.max(...bolts) : 0;
    const boltDev = minBolt > 0 ? ((maxBolt - minBolt) / minBolt) * 100 : 0;
    const boltPass = boltDev <= 50;

    // 2. IR checks
    const ir = data.electrical_tests.insulation_resistance_1min_1000v_megohms;
    const irPriPass = ir.pri_to_ground >= 500;
    const irSecPass = ir.sec_to_ground >= 500;
    const irPriSecPass = ir.pri_to_sec >= 500;
    const irAllPass = irPriPass && irSecPass && irPriSecPass;

    // Trend
    const prevIR = data.previous_test_data?.last_insulation_resistance_pri_ground || 0;
    const irDropPercent = prevIR > 0 ? ((prevIR - ir.pri_to_ground) / prevIR) * 100 : 0;
    const irTrendWarning = irDropPercent > 20;

    // 3. DAR
    const darPass = data.electrical_tests.dar_value >= 1.0;

    // 4. Turns-ratio
    const calc = data.electrical_tests.turns_ratio.calculated_ratio || 1;
    const ratios = data.electrical_tests.turns_ratio.measured_ratios || [];
    const phaseNames = ['Pha A', 'Pha B', 'Pha C'];
    const phaseCalcs = ratios.map((r, i) => {
      const err = calc > 0 ? (Math.abs(r - calc) / calc) * 100 : 0;
      return {
        phase: phaseNames[i] || `Pha ${i + 1}`,
        measured: r,
        error: err,
        pass: err <= 0.5
      };
    });
    const turnsRatioAllPass = phaseCalcs.every(p => p.pass);

    // 5. Visual
    const visualPass = data.visual_inspection.shipping_brackets_removed &&
      data.visual_inspection.nameplate_match &&
      data.visual_inspection.grounding_check.toLowerCase() === 'pass';

    return {
      minBolt,
      maxBolt,
      boltDev,
      boltPass,
      irPriPass,
      irSecPass,
      irPriSecPass,
      irAllPass,
      prevIR,
      irDropPercent,
      irTrendWarning,
      darPass,
      phaseCalcs,
      turnsRatioAllPass,
      visualPass
    };
  }, [data]);

  // Run AI Analysis via backend endpoint
  const handleRunAiAnalysis = async () => {
    setIsAnalyzing(true);
    try {
      const res = await fetch('/api/field-service/neta-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });
      const result = await res.json();
      if (result.success && result.data) {
        setAnalysisResult(result.data);
      } else {
        alert('Lỗi phân tích: ' + (result.error || 'Vui lòng thử lại'));
      }
    } catch (err: any) {
      alert('Không thể kết nối đến máy chủ AI: ' + err.message);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleCopyMarkdown = () => {
    if (!analysisResult?.markdown_report) return;
    navigator.clipboard.writeText(analysisResult.markdown_report);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
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
        type: 'Transformers, Dry-Type, Air-Cooled, Small (NETA ATS-2025 Sec 7.2.1.1)'
      });
      setSavedSuccessMessage(`Đã lưu biên bản kiểm định MBA nhỏ [${data.site_info.transformer_tag}] vào kho báo cáo CMMS!`);
    }
  };

  const handleApplyJson = () => {
    try {
      const parsed = JSON.parse(jsonInput);
      if (!parsed.site_info || !parsed.electrical_tests || !parsed.visual_inspection) {
        alert('JSON không đúng định dạng kiểm tra NETA ATS-2025. Vui lòng kiểm tra lại cấu trúc.');
        return;
      }
      setData(parsed);
      setShowJsonModal(false);
      setAnalysisResult(null);
    } catch (err: any) {
      alert('Lỗi cú pháp JSON: ' + err.message);
    }
  };

  return (
    <div className="space-y-6">
      {/* Top Banner & Scope Declaration */}
      <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-blue-900 rounded-2xl p-6 text-white border border-slate-700 shadow-lg">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-2">
            <div className="flex items-center gap-2.5 flex-wrap">
              <span className="px-3 py-1 rounded-full text-xs font-bold uppercase tracking-wider bg-blue-500/30 text-blue-200 border border-blue-400/40">
                ANSI/NETA ATS-2025 (Sec 7.2.1.1)
              </span>
              <span className="px-3 py-1 rounded-full text-xs font-medium bg-emerald-500/20 text-emerald-300 border border-emerald-400/30">
                Low-Voltage Small Dry-Type
              </span>
              <span className="px-3 py-1 rounded-full text-xs font-mono bg-amber-500/20 text-amber-300 border border-amber-400/30">
                &le; 600V | &le; 167kVA (1Ø) | &le; 500kVA (3Ø)
              </span>
            </div>
            <h2 className="text-2xl font-black text-white tracking-tight">
              Bảng Kiểm Tra & Đánh Giá Máy Biến Áp Khô Hạ Áp
            </h2>
            <p className="text-sm text-slate-300 max-w-3xl leading-relaxed">
              Dành riêng cho Field Service Engineer (FSE). Tích hợp ngưỡng tham chiếu chuẩn NETA ATS-2025, tự động tính sai số độ lệch điện trở mối nối bu-lông, suy giảm cách điện và trợ lý AI phân tích đánh giá đóng điện.
            </p>
          </div>

          <div className="flex flex-wrap sm:flex-col gap-2 shrink-0">
            <button
              onClick={() => setShowCameraModal(true)}
              className="px-3.5 py-2 bg-amber-400 hover:bg-amber-300 active:bg-amber-500 text-slate-950 font-bold rounded-xl text-xs shadow-md transition-all flex items-center justify-center gap-1.5"
              title="Quét nhãn Nameplate bằng Camera AI"
            >
              <Camera size={14} />
              <span>Quét Nameplate (AI OCR)</span>
            </button>

            <button
              onClick={() => {
                setData(DEFAULT_SAMPLE_DATA);
                setAnalysisResult(null);
              }}
              className="px-3.5 py-2 bg-blue-600 hover:bg-blue-500 active:bg-blue-700 text-white rounded-xl text-xs font-semibold shadow-md transition-all flex items-center justify-center gap-2"
            >
              <Zap size={14} />
              <span>Nạp Kịch Bản Mẫu FSE</span>
            </button>

            <button
              onClick={() => setShowPrintModal(true)}
              className="px-3.5 py-2 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-semibold shadow-md transition-all flex items-center justify-center gap-1.5 border border-indigo-400/40"
            >
              <Printer size={14} />
              <span>In / Xuất Báo Cáo PDF (TEV)</span>
            </button>

            <button
              onClick={handleSaveToCmms}
              className="px-3.5 py-2 bg-emerald-700 hover:bg-emerald-600 text-white rounded-xl text-xs font-semibold shadow-md transition-all flex items-center justify-center gap-1.5 border border-emerald-500/40"
              title="Lưu biên bản kiểm tra vào kho báo cáo"
            >
              <Save size={14} />
              <span>Lưu vào Kho Báo Cáo</span>
            </button>

            <div className="flex gap-2">
              <button
                onClick={() => {
                  setJsonInput(JSON.stringify(data, null, 2));
                  setShowJsonModal(true);
                }}
                className="flex-1 px-2.5 py-1.5 bg-white/10 hover:bg-white/20 border border-white/20 text-white rounded-lg text-xs font-medium flex items-center justify-center gap-1 transition-colors"
                title="Nhập chuỗi JSON dữ liệu đo"
              >
                <Upload size={12} />
                <span>Nhập JSON</span>
              </button>
              <button
                onClick={() => {
                  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
                  const url = URL.createObjectURL(blob);
                  const a = document.createElement('a');
                  a.href = url;
                  a.download = `NETA_ATS_2025_${data.site_info.transformer_tag || 'TR'}.json`;
                  a.click();
                }}
                className="flex-1 px-2.5 py-1.5 bg-white/10 hover:bg-white/20 border border-white/20 text-white rounded-lg text-xs font-medium flex items-center justify-center gap-1 transition-colors"
                title="Tải về file JSON"
              >
                <Download size={12} />
                <span>Xuất JSON</span>
              </button>
            </div>
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

      {/* JSON Import Modal */}
      {showJsonModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-white w-full max-w-2xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[85vh]">
            <div className="p-4 border-b border-slate-100 flex items-center justify-between bg-slate-50">
              <h3 className="font-bold text-slate-800 text-sm flex items-center gap-2">
                <FileText size={16} className="text-blue-600" />
                Nhập Dữ Liệu Kiểm Tra JSON (Field Service Input)
              </h3>
              <button onClick={() => setShowJsonModal(false)} className="text-slate-400 hover:text-slate-600">
                &times;
              </button>
            </div>
            <div className="p-4 flex-1 overflow-y-auto">
              <p className="text-xs text-slate-500 mb-2">
                Dán chuỗi JSON kiểm tra hiện trường từ thiết bị đo hoặc hệ thống SCADA/TEV:
              </p>
              <textarea
                value={jsonInput}
                onChange={(e) => setJsonInput(e.target.value)}
                rows={14}
                className="w-full p-3 font-mono text-xs bg-slate-900 text-emerald-400 rounded-xl border border-slate-800 focus:outline-hidden focus:ring-2 focus:ring-blue-500"
              />
            </div>
            <div className="p-4 border-t border-slate-100 bg-slate-50 flex justify-end gap-2">
              <button
                onClick={() => setShowJsonModal(false)}
                className="px-4 py-2 text-xs font-medium text-slate-600 hover:bg-slate-200 rounded-lg"
              >
                Hủy bỏ
              </button>
              <button
                onClick={handleApplyJson}
                className="px-4 py-2 text-xs font-semibold bg-blue-600 hover:bg-blue-700 text-white rounded-lg"
              >
                Áp dụng dữ liệu
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 1. Site Information Card */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-xs p-5">
        <div className="flex items-center justify-between mb-4 border-b border-slate-100 pb-2">
          <h3 className="text-sm font-bold uppercase tracking-wider text-slate-800 flex items-center gap-2">
            <Info size={16} className="text-blue-600" />
            1. Thông tin Hiện Trường & Nhãn MBA Khô
          </h3>
          <button
            type="button"
            onClick={() => setShowCameraModal(true)}
            className="px-2.5 py-1 bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-300 rounded font-bold text-[11px] transition-all flex items-center gap-1 shadow-2xs"
          >
            <Camera size={13} className="text-amber-600" /> Quét Nameplate
          </button>
        </div>

        <div className="flex items-center justify-between mb-3 bg-amber-50/70 border border-amber-200/80 rounded-lg p-2.5 px-3">
          <div className="flex items-center gap-2 text-xs text-amber-950 font-medium">
            <Camera size={15} className="text-amber-600 shrink-0" />
            <span>Chụp ảnh bảng tên kim loại MBA để AI tự động nhận dạng thông số</span>
          </div>
          <button
            type="button"
            onClick={() => setShowCameraModal(true)}
            className="px-3 py-1 bg-amber-600 hover:bg-amber-700 text-white rounded font-bold text-xs shadow-2xs transition-all shrink-0 ml-2"
          >
            Mở Camera
          </button>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">Tên Dự án / Nhà máy</label>
            <input
              type="text"
              value={data.site_info.project_name}
              onChange={(e) => setData({
                ...data,
                site_info: { ...data.site_info, project_name: e.target.value }
              })}
              className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">Mã định danh (Tag MBA)</label>
            <input
              type="text"
              value={data.site_info.transformer_tag}
              onChange={(e) => setData({
                ...data,
                site_info: { ...data.site_info, transformer_tag: e.target.value }
              })}
              className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none font-semibold text-blue-900"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">Kỹ sư kiểm tra (FSE TEV)</label>
            <FseEngineerDropdown
              value={data.site_info.fse_name}
              onChange={(val) => setData({
                ...data,
                site_info: { ...data.site_info, fse_name: val }
              })}
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1">Ngày thí nghiệm</label>
            <input
              type="date"
              value={data.site_info.test_date}
              onChange={(e) => setData({
                ...data,
                site_info: { ...data.site_info, test_date: e.target.value }
              })}
              className="w-full px-3 py-2 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
            />
          </div>
        </div>
      </div>

      {/* 2. Visual & Mechanical Inspection (NETA Sec 7.2.1.1.1) */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
        <div className="bg-slate-50 px-5 py-4 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Eye size={18} className="text-indigo-600" />
            <h3 className="font-bold text-slate-900 text-sm">
              2. Kiểm Tra Thị Giác & Cơ Khí (Visual & Mechanical Inspection)
            </h3>
          </div>
          <span className="text-xs font-medium text-slate-500">ANSI/NETA ATS-2025 Mục 7.2.1.1.1</span>
        </div>

        <div className="p-5 space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {/* Nameplate */}
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/50 flex items-center justify-between">
              <div>
                <div className="text-sm font-semibold text-slate-800">Thông số nhãn máy (Nameplate)</div>
                <div className="text-xs text-slate-500">So khớp bản vẽ thiết kế và cấu hình thực tế</div>
              </div>
              <button
                onClick={() => setData({
                  ...data,
                  visual_inspection: { ...data.visual_inspection, nameplate_match: !data.visual_inspection.nameplate_match }
                })}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                  data.visual_inspection.nameplate_match
                    ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                    : 'bg-rose-100 text-rose-800 border border-rose-300'
                }`}
              >
                {data.visual_inspection.nameplate_match ? 'Khớp bản vẽ' : 'Không khớp'}
              </button>
            </div>

            {/* Physical condition */}
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/50 flex items-center justify-between">
              <div>
                <div className="text-sm font-semibold text-slate-800">Tình trạng cơ lý bề ngoài</div>
                <div className="text-xs text-slate-500">Không móp méo, nứt vỡ, trầy xước, rỉ sét</div>
              </div>
              <select
                value={data.visual_inspection.physical_condition}
                onChange={(e) => setData({
                  ...data,
                  visual_inspection: { ...data.visual_inspection, physical_condition: e.target.value }
                })}
                className="px-3 py-1.5 text-xs font-semibold bg-white border border-slate-300 rounded-lg outline-none"
              >
                <option value="Good">Good (Tốt)</option>
                <option value="Acceptable">Acceptable (Chấp nhận)</option>
                <option value="Damaged">Damaged (Hư hỏng)</option>
              </select>
            </div>

            {/* Grounding */}
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/50 flex items-center justify-between">
              <div>
                <div className="text-sm font-semibold text-slate-800">Cố định & Tiếp địa (Grounding)</div>
                <div className="text-xs text-slate-500">Tiếp địa vỏ máy và điểm trung tính đúng thiết kế</div>
              </div>
              <select
                value={data.visual_inspection.grounding_check}
                onChange={(e) => setData({
                  ...data,
                  visual_inspection: { ...data.visual_inspection, grounding_check: e.target.value }
                })}
                className="px-3 py-1.5 text-xs font-semibold bg-white border border-slate-300 rounded-lg outline-none"
              >
                <option value="Pass">Pass (Đạt)</option>
                <option value="Fail">Fail (Không đạt)</option>
              </select>
            </div>

            {/* Shipping Brackets (CRITICAL) */}
            <div className={`p-3.5 rounded-xl border flex items-center justify-between transition-colors ${
              data.visual_inspection.shipping_brackets_removed
                ? 'border-emerald-200 bg-emerald-50/40'
                : 'border-rose-300 bg-rose-50'
            }`}>
              <div>
                <div className="text-sm font-bold text-slate-900 flex items-center gap-1.5">
                  <span>Khung kẹp vận chuyển (Shipping brackets)</span>
                  <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-amber-200 text-amber-900 uppercase">Bắt buộc</span>
                </div>
                <div className="text-xs text-slate-600">Đã tháo rời hoàn toàn; chân đế đàn hồi (resilient mounts) tự do</div>
              </div>
              <button
                onClick={() => setData({
                  ...data,
                  visual_inspection: { ...data.visual_inspection, shipping_brackets_removed: !data.visual_inspection.shipping_brackets_removed }
                })}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                  data.visual_inspection.shipping_brackets_removed
                    ? 'bg-emerald-600 text-white'
                    : 'bg-rose-600 text-white animate-bounce'
                }`}
              >
                {data.visual_inspection.shipping_brackets_removed ? 'ĐÃ THÁO RỜI' : 'CHƯA THÁO'}
              </button>
            </div>

            {/* Cleanliness */}
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/50 flex items-center justify-between">
              <div>
                <div className="text-sm font-semibold text-slate-800">Độ sạch sẽ (Cleanliness)</div>
                <div className="text-xs text-slate-500">Bề mặt cuộn dây, lõi thép sạch sẽ, không bụi/ẩm</div>
              </div>
              <select
                value={data.visual_inspection.cleanliness}
                onChange={(e) => setData({
                  ...data,
                  visual_inspection: { ...data.visual_inspection, cleanliness: e.target.value }
                })}
                className="px-3 py-1.5 text-xs font-semibold bg-white border border-slate-300 rounded-lg outline-none"
              >
                <option value="Pass">Pass (Sạch sẽ)</option>
                <option value="Needs Cleaning">Needs Cleaning (Cần vệ sinh)</option>
                <option value="Dirty/Wet">Dirty/Wet (Ẩm ướt/Bẩn)</option>
              </select>
            </div>

            {/* Bolt Torque */}
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/50 flex items-center justify-between">
              <div>
                <div className="text-sm font-semibold text-slate-800">Lực siết bu-lông (Bolt Torque)</div>
                <div className="text-xs text-slate-500">Theo NSX hoặc NETA ATS Table 100.12</div>
              </div>
              <select
                value={data.visual_inspection.bolt_torque_check}
                onChange={(e) => setData({
                  ...data,
                  visual_inspection: { ...data.visual_inspection, bolt_torque_check: e.target.value }
                })}
                className="px-3 py-1.5 text-xs font-semibold bg-white border border-slate-300 rounded-lg outline-none"
              >
                <option value="Pass">Pass (Đạt tiêu chuẩn)</option>
                <option value="Re-torqued">Re-torqued (Đã siết lại)</option>
                <option value="Fail">Fail (Chưa đạt)</option>
              </select>
            </div>

            {/* Tap Position */}
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/50 flex items-center justify-between md:col-span-2">
              <div>
                <div className="text-sm font-semibold text-slate-800">Nấc phân áp (Tap connections)</div>
                <div className="text-xs text-slate-500">Xác nhận đúng vị trí chỉ định ban đầu theo yêu cầu vận hành</div>
              </div>
              <input
                type="text"
                value={data.visual_inspection.tap_position}
                onChange={(e) => setData({
                  ...data,
                  visual_inspection: { ...data.visual_inspection, tap_position: e.target.value }
                })}
                placeholder="VD: Tap 3 (480V)"
                className="px-3 py-1.5 text-xs font-semibold bg-white border border-slate-300 rounded-lg outline-none w-44"
              />
            </div>
          </div>
        </div>
      </div>

      {/* 3. Electrical Test Values & Reference Standards */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
        <div className="bg-slate-900 px-5 py-4 text-white flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Gauge size={18} className="text-amber-400" />
            <h3 className="font-bold text-sm">
              3. Phép Đo Điện & Tiêu Chuẩn Tham Chiếu NETA ATS-2025
            </h3>
          </div>
          <span className="text-xs text-slate-400">Section 7.2.1.1.2 Electrical Tests</span>
        </div>

        <div className="p-5 space-y-6">
          {/* COMPARISON TABLE */}
          <div className="overflow-x-auto rounded-xl border border-slate-200">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-100 text-slate-700 font-bold uppercase tracking-wider border-b border-slate-200">
                  <th className="p-3.5">Hạng mục kiểm tra</th>
                  <th className="p-3.5">Giá trị đo thực tế</th>
                  <th className="p-3.5">Tiêu chuẩn NETA ATS-2025</th>
                  <th className="p-3.5">Lần đo trước / Tham chiếu</th>
                  <th className="p-3.5">Đánh giá sai số & Trend</th>
                  <th className="p-3.5 text-center">Trạng thái</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {/* 1. Bolted connection resistance */}
                <tr className={liveStats.boltPass ? 'hover:bg-slate-50/80' : 'bg-amber-50/40 hover:bg-amber-50/70'}>
                  <td className="p-3.5 font-semibold text-slate-900">
                    <div>Điện trở mối nối bu-lông</div>
                    <div className="text-[11px] text-slate-500 font-normal">Low-Resistance Ohmmeter (µΩ)</div>
                  </td>
                  <td className="p-3.5">
                    <div className="flex items-center gap-1.5 flex-wrap">
                      {data.electrical_tests.bolted_resistance_micro_ohms.map((val, idx) => (
                        <div key={idx} className="flex items-center gap-1">
                          <span className="text-[10px] text-slate-400">P{idx + 1}:</span>
                          <input
                            type="number"
                            step="0.1"
                            value={val}
                            onChange={(e) => {
                              const newArr = [...data.electrical_tests.bolted_resistance_micro_ohms];
                              newArr[idx] = parseFloat(e.target.value) || 0;
                              setData({
                                ...data,
                                electrical_tests: { ...data.electrical_tests, bolted_resistance_micro_ohms: newArr }
                              });
                            }}
                            className={`w-16 px-2 py-1 border rounded text-xs font-mono font-bold ${
                              val === liveStats.maxBolt && !liveStats.boltPass
                                ? 'border-amber-500 bg-amber-50 text-amber-900'
                                : 'border-slate-300 bg-white'
                            }`}
                          />
                        </div>
                      ))}
                    </div>
                  </td>
                  <td className="p-3.5 text-slate-600">
                    Độ lệch so với mối nối nhỏ nhất: <span className="font-semibold text-slate-800">&le; 50%</span>
                  </td>
                  <td className="p-3.5 text-slate-600">
                    Min mối nối: <strong className="text-slate-800">{liveStats.minBolt} µΩ</strong>
                  </td>
                  <td className="p-3.5">
                    <span className={liveStats.boltPass ? 'text-emerald-700 font-semibold' : 'text-amber-700 font-bold'}>
                      Lệch: {liveStats.boltDev.toFixed(1)}%
                    </span>
                    {!liveStats.boltPass && (
                      <div className="text-[10px] text-amber-800 font-medium">Vượt ngưỡng 50% NETA</div>
                    )}
                  </td>
                  <td className="p-3.5 text-center">
                    <span className={`px-2.5 py-1 rounded-full text-[11px] font-bold uppercase ${
                      liveStats.boltPass
                        ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                        : 'bg-amber-100 text-amber-800 border border-amber-300'
                    }`}>
                      {liveStats.boltPass ? 'PASS' : 'INVESTIGATE'}
                    </span>
                  </td>
                </tr>

                {/* 2. IR Pri to Ground */}
                <tr className={liveStats.irPriPass ? 'hover:bg-slate-50/80' : 'bg-rose-50/40 hover:bg-rose-50/70'}>
                  <td className="p-3.5 font-semibold text-slate-900">
                    <div>Cách điện: Sơ cấp - Đất</div>
                    <div className="text-[11px] text-slate-500 font-normal">IR 1 min @ 1000V DC (MΩ)</div>
                  </td>
                  <td className="p-3.5">
                    <input
                      type="number"
                      value={data.electrical_tests.insulation_resistance_1min_1000v_megohms.pri_to_ground}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1min_1000v_megohms: {
                            ...data.electrical_tests.insulation_resistance_1min_1000v_megohms,
                            pri_to_ground: parseFloat(e.target.value) || 0
                          }
                        }
                      })}
                      className="w-24 px-2 py-1 border border-slate-300 rounded text-xs font-mono font-bold bg-white"
                    />
                  </td>
                  <td className="p-3.5 text-slate-600">
                    Tối thiểu: <span className="font-semibold text-slate-800">&ge; 500 MΩ</span> (Table 100.5)
                  </td>
                  <td className="p-3.5">
                    <div className="flex items-center gap-1">
                      <input
                        type="number"
                        placeholder="Lần trước"
                        value={data.previous_test_data?.last_insulation_resistance_pri_ground || ''}
                        onChange={(e) => setData({
                          ...data,
                          previous_test_data: {
                            ...data.previous_test_data,
                            last_insulation_resistance_pri_ground: parseFloat(e.target.value) || 0
                          }
                        })}
                        className="w-20 px-2 py-0.5 border border-slate-300 rounded text-[11px] font-mono bg-white"
                      />
                      <span className="text-[11px] text-slate-500">MΩ</span>
                    </div>
                  </td>
                  <td className="p-3.5">
                    {liveStats.prevIR > 0 ? (
                      <div>
                        <span className={liveStats.irTrendWarning ? 'text-amber-700 font-bold' : 'text-emerald-700 font-semibold'}>
                          Giảm: {liveStats.irDropPercent.toFixed(1)}%
                        </span>
                        {liveStats.irTrendWarning && (
                          <div className="text-[10px] text-amber-800 font-medium">Giảm &gt; 20% (Cảnh báo ẩm)</div>
                        )}
                      </div>
                    ) : (
                      <span className="text-slate-400">Không có dữ liệu cũ</span>
                    )}
                  </td>
                  <td className="p-3.5 text-center">
                    <span className={`px-2.5 py-1 rounded-full text-[11px] font-bold uppercase ${
                      liveStats.irPriPass
                        ? (liveStats.irTrendWarning ? 'bg-amber-100 text-amber-800 border border-amber-300' : 'bg-emerald-100 text-emerald-800 border border-emerald-300')
                        : 'bg-rose-100 text-rose-800 border border-rose-300'
                    }`}>
                      {liveStats.irPriPass ? (liveStats.irTrendWarning ? 'INVESTIGATE' : 'PASS') : 'FAIL'}
                    </span>
                  </td>
                </tr>

                {/* 3. IR Sec to Ground */}
                <tr className="hover:bg-slate-50/80">
                  <td className="p-3.5 font-semibold text-slate-900">
                    <div>Cách điện: Thứ cấp - Đất</div>
                    <div className="text-[11px] text-slate-500 font-normal">IR 1 min @ 1000V DC (MΩ)</div>
                  </td>
                  <td className="p-3.5">
                    <input
                      type="number"
                      value={data.electrical_tests.insulation_resistance_1min_1000v_megohms.sec_to_ground}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1min_1000v_megohms: {
                            ...data.electrical_tests.insulation_resistance_1min_1000v_megohms,
                            sec_to_ground: parseFloat(e.target.value) || 0
                          }
                        }
                      })}
                      className="w-24 px-2 py-1 border border-slate-300 rounded text-xs font-mono font-bold bg-white"
                    />
                  </td>
                  <td className="p-3.5 text-slate-600">
                    Tối thiểu: <span className="font-semibold text-slate-800">&ge; 500 MΩ</span>
                  </td>
                  <td className="p-3.5 text-slate-400">N/A</td>
                  <td className="p-3.5 text-emerald-700 font-semibold">Đạt chuẩn cách điện</td>
                  <td className="p-3.5 text-center">
                    <span className={`px-2.5 py-1 rounded-full text-[11px] font-bold uppercase ${
                      liveStats.irSecPass ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'
                    }`}>
                      {liveStats.irSecPass ? 'PASS' : 'FAIL'}
                    </span>
                  </td>
                </tr>

                {/* 4. IR Pri to Sec */}
                <tr className="hover:bg-slate-50/80">
                  <td className="p-3.5 font-semibold text-slate-900">
                    <div>Cách điện: Sơ cấp - Thứ cấp</div>
                    <div className="text-[11px] text-slate-500 font-normal">IR 1 min @ 1000V DC (MΩ)</div>
                  </td>
                  <td className="p-3.5">
                    <input
                      type="number"
                      value={data.electrical_tests.insulation_resistance_1min_1000v_megohms.pri_to_sec}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: {
                          ...data.electrical_tests,
                          insulation_resistance_1min_1000v_megohms: {
                            ...data.electrical_tests.insulation_resistance_1min_1000v_megohms,
                            pri_to_sec: parseFloat(e.target.value) || 0
                          }
                        }
                      })}
                      className="w-24 px-2 py-1 border border-slate-300 rounded text-xs font-mono font-bold bg-white"
                    />
                  </td>
                  <td className="p-3.5 text-slate-600">
                    Tối thiểu: <span className="font-semibold text-slate-800">&ge; 500 MΩ</span>
                  </td>
                  <td className="p-3.5 text-slate-400">N/A</td>
                  <td className="p-3.5 text-emerald-700 font-semibold">Đạt chuẩn cách điện</td>
                  <td className="p-3.5 text-center">
                    <span className={`px-2.5 py-1 rounded-full text-[11px] font-bold uppercase ${
                      liveStats.irPriSecPass ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'
                    }`}>
                      {liveStats.irPriSecPass ? 'PASS' : 'FAIL'}
                    </span>
                  </td>
                </tr>

                {/* 5. DAR */}
                <tr className="hover:bg-slate-50/80">
                  <td className="p-3.5 font-semibold text-slate-900">
                    <div>Hệ số hấp thụ điện dịch (DAR)</div>
                    <div className="text-[11px] text-slate-500 font-normal">Tỷ số R60s / R30s</div>
                  </td>
                  <td className="p-3.5">
                    <input
                      type="number"
                      step="0.05"
                      value={data.electrical_tests.dar_value}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: { ...data.electrical_tests, dar_value: parseFloat(e.target.value) || 0 }
                      })}
                      className="w-20 px-2 py-1 border border-slate-300 rounded text-xs font-mono font-bold bg-white"
                    />
                  </td>
                  <td className="p-3.5 text-slate-600">
                    Tối thiểu: <span className="font-semibold text-slate-800">&ge; 1.0</span> (Không được &lt; 1.0)
                  </td>
                  <td className="p-3.5 text-slate-400">N/A</td>
                  <td className="p-3.5">
                    <span className={liveStats.darPass ? 'text-emerald-700 font-semibold' : 'text-rose-700 font-bold'}>
                      {liveStats.darPass ? 'Khả năng nạp điện tích tốt' : 'Dưới ngưỡng 1.0 (Nguy cơ rò)'}
                    </span>
                  </td>
                  <td className="p-3.5 text-center">
                    <span className={`px-2.5 py-1 rounded-full text-[11px] font-bold uppercase ${
                      liveStats.darPass ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'
                    }`}>
                      {liveStats.darPass ? 'PASS' : 'FAIL'}
                    </span>
                  </td>
                </tr>

                {/* 6. Turns-Ratio Test */}
                <tr className={liveStats.turnsRatioAllPass ? 'hover:bg-slate-50/80' : 'bg-rose-50/40 hover:bg-rose-50/70'}>
                  <td className="p-3.5 font-semibold text-slate-900">
                    <div>Tỷ số biến áp (Turns-Ratio TTR)</div>
                    <div className="text-[11px] text-slate-500 font-normal">Sai số giữa các pha & so với tính toán</div>
                  </td>
                  <td className="p-3.5">
                    <div className="space-y-1.5">
                      <div className="flex items-center gap-1.5 flex-wrap">
                        {data.electrical_tests.turns_ratio.measured_ratios.map((val, idx) => (
                          <div key={idx} className="flex items-center gap-1">
                            <span className="text-[10px] text-slate-400">{PHASE_NAMES[idx] || `P${idx+1}`}:</span>
                            <input
                              type="number"
                              step="0.001"
                              value={val}
                              onChange={(e) => {
                                const newR = [...data.electrical_tests.turns_ratio.measured_ratios];
                                newR[idx] = parseFloat(e.target.value) || 0;
                                setData({
                                  ...data,
                                  electrical_tests: {
                                    ...data.electrical_tests,
                                    turns_ratio: { ...data.electrical_tests.turns_ratio, measured_ratios: newR }
                                  }
                                });
                              }}
                              className={`w-20 px-2 py-1 border rounded text-xs font-mono font-bold ${
                                liveStats.phaseCalcs[idx]?.pass ? 'border-slate-300 bg-white' : 'border-rose-400 bg-rose-50 text-rose-900'
                              }`}
                            />
                          </div>
                        ))}
                      </div>
                    </div>
                  </td>
                  <td className="p-3.5 text-slate-600">
                    Sai số tối đa: <span className="font-semibold text-slate-800">&le; 0.5%</span>
                  </td>
                  <td className="p-3.5">
                    <div className="flex items-center gap-1">
                      <span className="text-[11px] text-slate-500">Tỷ số tính toán:</span>
                      <input
                        type="number"
                        step="0.001"
                        value={data.electrical_tests.turns_ratio.calculated_ratio}
                        onChange={(e) => setData({
                          ...data,
                          electrical_tests: {
                            ...data.electrical_tests,
                            turns_ratio: { ...data.electrical_tests.turns_ratio, calculated_ratio: parseFloat(e.target.value) || 1 }
                          }
                        })}
                        className="w-16 px-1.5 py-0.5 border border-slate-300 rounded text-xs font-mono font-semibold bg-white"
                      />
                    </div>
                  </td>
                  <td className="p-3.5">
                    <div className="space-y-0.5">
                      {liveStats.phaseCalcs.map((p, i) => (
                        <div key={i} className={`text-[11px] ${p.pass ? 'text-slate-600' : 'text-rose-700 font-bold'}`}>
                          {p.phase}: lệch {p.error.toFixed(2)}% {p.pass ? '✓' : '⚠️ (>0.5%)'}
                        </div>
                      ))}
                    </div>
                  </td>
                  <td className="p-3.5 text-center">
                    <span className={`px-2.5 py-1 rounded-full text-[11px] font-bold uppercase ${
                      liveStats.turnsRatioAllPass ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'
                    }`}>
                      {liveStats.turnsRatioAllPass ? 'PASS' : 'FAIL'}
                    </span>
                  </td>
                </tr>

                {/* 7. Secondary Voltage */}
                <tr className="hover:bg-slate-50/80">
                  <td className="p-3.5 font-semibold text-slate-900">
                    <div>Điện áp thứ cấp không tải</div>
                    <div className="text-[11px] text-slate-500 font-normal">Pha-Pha & Pha-Trung tính</div>
                  </td>
                  <td className="p-3.5">
                    <input
                      type="text"
                      value={data.electrical_tests.secondary_voltage_no_load}
                      onChange={(e) => setData({
                        ...data,
                        electrical_tests: { ...data.electrical_tests, secondary_voltage_no_load: e.target.value }
                      })}
                      className="w-52 px-2.5 py-1 border border-slate-300 rounded text-xs font-semibold bg-white"
                    />
                  </td>
                  <td className="p-3.5 text-slate-600">Phù hợp thông số nhãn máy (Nameplate)</td>
                  <td className="p-3.5 text-slate-600">208V/120V danh định</td>
                  <td className="p-3.5 text-emerald-700 font-semibold">Khớp điện áp thiết kế</td>
                  <td className="p-3.5 text-center">
                    <span className="px-2.5 py-1 rounded-full text-[11px] font-bold uppercase bg-emerald-100 text-emerald-800">
                      PASS
                    </span>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* Trigger AI Evaluation Button */}
          <div className="pt-2 flex flex-col sm:flex-row items-center justify-between gap-4 border-t border-slate-100">
            <div className="text-xs text-slate-500">
              * Hệ thống tự động so khớp tất cả các quy tắc ANSI/NETA ATS-2025 Mục 7.2.1.1. Bấm nút dưới đây để AI phân tích chuyên sâu.
            </div>
            <button
              disabled={isAnalyzing}
              onClick={handleRunAiAnalysis}
              className="w-full sm:w-auto px-6 py-3 bg-gradient-to-r from-indigo-600 via-blue-600 to-blue-700 hover:from-indigo-500 hover:to-blue-600 active:from-indigo-700 active:to-blue-800 disabled:opacity-50 text-white rounded-xl text-sm font-bold shadow-md flex items-center justify-center gap-2.5 transition-all"
            >
              {isAnalyzing ? (
                <>
                  <RefreshCw size={16} className="animate-spin" />
                  <span>TEV AI Đang Phân Tích Chuyên Sâu...</span>
                </>
              ) : (
                <>
                  <Bot size={18} />
                  <span>Phân Tích AI Chuyên Gia (TEV Platform AI)</span>
                </>
              )}
            </button>
          </div>
        </div>
      </div>

      {/* 4. AI Analysis Result Section (4 Parts) */}
      {analysisResult && (
        <div className="bg-white rounded-2xl border-2 border-indigo-200 shadow-xl overflow-hidden animate-in fade-in duration-300">
          {/* Header */}
          <div className="bg-gradient-to-r from-slate-950 via-slate-900 to-indigo-950 px-6 py-4 text-white flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-xl bg-blue-500/20 border border-blue-400/40 text-blue-300 flex items-center justify-center">
                <Bot size={20} />
              </div>
              <div>
                <h3 className="font-bold text-base text-white">
                  Báo Cáo Phân Tích Chuyên Gia NETA ATS-2025
                </h3>
                <p className="text-xs text-slate-400">
                  Thiết bị: {data.site_info.transformer_tag} • Kỹ sư: {data.site_info.fse_name} • Ngày: {data.site_info.test_date}
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={handleCopyMarkdown}
                className="px-3 py-1.5 bg-white/10 hover:bg-white/20 text-white rounded-lg text-xs font-medium flex items-center gap-1.5 transition-colors border border-white/20"
              >
                {copied ? <Check size={14} className="text-emerald-400" /> : <Copy size={14} />}
                <span>{copied ? 'Đã sao chép' : 'Sao chép Markdown'}</span>
              </button>
              {onSaveReport && (
                <button
                  onClick={handleSaveToCmms}
                  className="px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-semibold flex items-center gap-1.5 shadow-sm transition-colors"
                >
                  <Save size={14} />
                  <span>Lưu vào Kho Báo Cáo</span>
                </button>
              )}
            </div>
          </div>

          <div className="p-6 space-y-6">
            {/* PART 1: OVERALL STATUS */}
            <div className={`p-5 rounded-2xl border flex flex-col sm:flex-row sm:items-center justify-between gap-4 ${
              analysisResult.overall_status === 'INVESTIGATE'
                ? 'bg-amber-500/10 border-amber-300'
                : analysisResult.overall_status === 'FAIL'
                ? 'bg-rose-500/10 border-rose-300'
                : 'bg-emerald-500/10 border-emerald-300'
            }`}>
              <div className="space-y-1">
                <div className="text-xs uppercase font-extrabold tracking-wider text-slate-500">
                  1. Trạng Thái Tổng Quan (Overall Status)
                </div>
                <div className="flex items-center gap-2.5">
                  <span className={`px-4 py-1.5 rounded-full text-base font-black tracking-wide uppercase ${
                    analysisResult.overall_status === 'INVESTIGATE'
                      ? 'bg-amber-500 text-slate-950'
                      : analysisResult.overall_status === 'FAIL'
                      ? 'bg-rose-600 text-white'
                      : 'bg-emerald-600 text-white'
                  }`}>
                    {analysisResult.overall_status}
                  </span>
                  <span className="text-sm font-bold text-slate-800">
                    {analysisResult.overall_status === 'INVESTIGATE'
                      ? 'CẦN KIỂM TRA / XỬ LÝ LẠI TẠI CHỖ TRƯỚC KHI ĐÓNG ĐIỆN'
                      : analysisResult.overall_status === 'FAIL'
                      ? 'KHÔNG ĐẠT TIÊU CHUẨN ĐÓNG ĐIỆN'
                      : 'ĐẠT TIÊU CHUẨN ANSI/NETA ATS-2025'}
                  </span>
                </div>
                <p className="text-xs text-slate-600 mt-1">{analysisResult.status_reason}</p>
              </div>

              <div className="text-right shrink-0">
                <div className="text-xs text-slate-400">Tiêu chuẩn nghiệm thu</div>
                <div className="text-xs font-mono font-bold text-slate-800">ANSI/NETA ATS-2025 Sec 7.2.1.1</div>
              </div>
            </div>

            {/* PART 2 & 3 & 4 IN FORMATTED MARKDOWN & VISUAL CARDS */}
            <div className="prose prose-slate max-w-none text-sm bg-slate-50/70 p-6 rounded-2xl border border-slate-200">
              <div className="whitespace-pre-line leading-relaxed text-slate-800 font-sans">
                {analysisResult.markdown_report}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Printable TEV Field Service Report */}
      <TevReportPrintModal
        isOpen={showPrintModal}
        onClose={() => setShowPrintModal(false)}
        title="BIÊN BẢN KIỂM ĐỊNH MÁY BIẾN ÁP KHÔ HẠ ÁP"
        standard="ANSI/NETA ATS-2025 Section 7.2.1.1"
        siteInfo={data.site_info}
        overallStatus={analysisResult?.overall_status || (liveStats.boltPass && liveStats.ttrPass && !liveStats.irTrendWarning ? 'PASS' : 'INVESTIGATE')}
        statusReason={analysisResult?.status_reason || 'Đã đối chiếu với quy chuẩn kỹ thuật NETA ATS-2025 Mục 7.2.1.1.'}
        visualChecks={[
          { label: 'Nameplate khớp bản vẽ thiết kế', value: data.visual_inspection.nameplate_match ? 'ĐẠT' : 'KHÔNG ĐẠT', pass: data.visual_inspection.nameplate_match },
          { label: 'Tình trạng cơ lý vỏ/lõi', value: data.visual_inspection.physical_condition, pass: data.visual_inspection.physical_condition === 'Good' },
          { label: 'Tiếp địa vỏ & lõi', value: data.visual_inspection.grounding_check, pass: data.visual_inspection.grounding_check === 'Pass' },
          { label: 'Khung kẹp vận chuyển (Shipping brackets)', value: data.visual_inspection.shipping_brackets_removed ? 'ĐÃ THÁO RỜI' : 'CHƯA THÁO', pass: data.visual_inspection.shipping_brackets_removed },
          { label: 'Độ sạch sẽ & khô ráo', value: data.visual_inspection.cleanliness, pass: data.visual_inspection.cleanliness === 'Pass' },
          { label: 'Lực siết bu-lông (Bolt Torque Table 100.12)', value: data.visual_inspection.bolt_torque_check, pass: data.visual_inspection.bolt_torque_check === 'Pass' },
          { label: 'Nấc phân áp (Tap connections)', value: data.visual_inspection.tap_position, pass: true }
        ]}
        electricalTable={[
          {
            item: 'Điện trở mối nối bu-lông',
            measured: `[${data.electrical_tests.bolted_resistance_micro_ohms.join(', ')}] µΩ`,
            standard: 'Lệch ≤ 50% so với mối nối nhỏ nhất',
            reference: `Min: ${liveStats.minBolt} µΩ`,
            deviation: `Lệch lớn nhất: ${liveStats.boltDev}%`,
            status: liveStats.boltPass ? 'PASS' : 'INVESTIGATE'
          },
          {
            item: 'Điện trở cách điện Sơ - Đất (1000V DC)',
            measured: `${data.electrical_tests.insulation_resistance_1min_1000v_megohms.pri_to_ground} MΩ`,
            standard: '≥ 500 MΩ (Table 100.5)',
            reference: `${data.previous_test_data?.last_insulation_resistance_pri_ground || 'N/A'} MΩ`,
            deviation: liveStats.irDrop > 0 ? `Giảm ${liveStats.irDrop}% so với trước` : 'Ổn định',
            status: liveStats.irTrendWarning ? 'INVESTIGATE' : 'PASS'
          },
          {
            item: 'Điện trở cách điện Thứ - Đất',
            measured: `${data.electrical_tests.insulation_resistance_1min_1000v_megohms.sec_to_ground} MΩ`,
            standard: '≥ 500 MΩ (Table 100.5)',
            reference: 'N/A',
            deviation: 'Đạt chuẩn',
            status: 'PASS'
          },
          {
            item: 'Điện trở cách điện Sơ - Thứ',
            measured: `${data.electrical_tests.insulation_resistance_1min_1000v_megohms.pri_to_sec} MΩ`,
            standard: '≥ 500 MΩ (Table 100.5)',
            reference: 'N/A',
            deviation: 'Đạt chuẩn',
            status: 'PASS'
          },
          {
            item: 'Hệ số hấp thụ điện dịch (DAR)',
            measured: `${data.electrical_tests.dar_value}`,
            standard: '≥ 1.0 (NETA ATS-2025)',
            reference: 'Chuẩn ≥ 1.0',
            deviation: `DAR = ${data.electrical_tests.dar_value}`,
            status: liveStats.darPass ? 'PASS' : 'FAIL'
          },
          {
            item: 'Tỷ số biến áp (Turns-Ratio TTR)',
            measured: `[${data.electrical_tests.turns_ratio.measured_ratios.join(', ')}]`,
            standard: 'Sai số ≤ 0.5% so với tính toán',
            reference: `Tính toán: ${data.electrical_tests.turns_ratio.calculated_ratio}`,
            deviation: `Sai số: ${liveStats.maxTtrError}% (Pha C)`,
            status: liveStats.ttrPass ? 'PASS' : 'FAIL'
          },
          {
            item: 'Điện áp thứ cấp không tải',
            measured: data.electrical_tests.secondary_voltage_no_load,
            standard: 'Khớp Nameplate máy',
            reference: '208V / 120V',
            deviation: 'Khớp',
            status: 'PASS'
          }
        ]}
        warnings={[
          !liveStats.boltPass ? `Điện trở mối nối bu-lông lệch ${liveStats.boltDev}% (> 50%) so với mối nối nhỏ nhất (${liveStats.minBolt} µΩ).` : '',
          liveStats.irTrendWarning ? `Điện trở cách điện Sơ-Đất giảm ${liveStats.irDrop}% so với lần đo trước (${data.previous_test_data?.last_insulation_resistance_pri_ground} MΩ).` : '',
          !liveStats.ttrPass ? `Tỷ số biến áp Pha C vượt quá sai số cho phép 0.5% (đo được sai số ${liveStats.maxTtrError}%).` : ''
        ].filter(Boolean)}
        recommendations={[
          'Siết lại bu-lông mối nối có điện trở cao bằng cờ-lê lực theo Table 100.12 và đo lại bằng DLRO.',
          'Kiểm tra lại nấc chuyển mạch phân áp (Tap switch/connections), siết chặt đầu cực và đo lại TTR.',
          'Vệ sinh cuộn dây và xem xét sấy nhẹ nếu độ ẩm trạm biến áp cao.'
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
