import React, { useState, useMemo } from 'react';
import {
  ShieldCheck,
  ShieldAlert,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  Sparkles,
  Printer,
  Save,
  Copy,
  Check,
  Info,
  Layers,
  Gauge,
  Flame,
  Zap,
  Activity,
  Globe,
  Share2,
  FileText,
  FileCheck2,
  Code2,
  ChevronDown,
  ArrowRight,
  Database
} from 'lucide-react';
import {
  NETA_ATS_23_MASTER_SPECS,
  NetaAtsSpecificationDef,
  generateAtsMarkdownReport,
  getAllAtsSpecsList,
} from '../data/netaAtsMasterSpecifications';
import { TevReportPrintModal } from './TevReportPrintModal';
import {
  buildDynamicAtsSpec,
  getAtsPayload,
  saveAtsPayload,
  resetAtsPayload,
} from '../services/netaAtsDataService';
import { CustomerSelectDropdown } from './CustomerSelectDropdown';
import { ProjectSelectDropdown } from './ProjectSelectDropdown';
import { FseEngineerDropdown } from './FseEngineerDropdown';

interface Props {
  initialSpecId?: string;
  initialLang?: 'vi' | 'en';
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
  onSelectSpecId?: (specId: string) => void;
  onSwitchToForm?: () => void;
  displayMode?: 'tev-report' | 'form';
}

export const NetaAtsUniversalReportView: React.FC<Props> = ({
  initialSpecId = 'neta-liquid-transformer',
  initialLang = 'vi',
  onSaveReport,
  onNavigateToReports,
  onSelectSpecId,
  onSwitchToForm,
  displayMode = 'tev-report',
}) => {
  const [selectedId, setSelectedId] = useState<string>(initialSpecId);
  const [lang, setLang] = useState<'vi' | 'en'>(initialLang);
  const [payloadVersion, setPayloadVersion] = useState(0);
  const [copied, setCopied] = useState(false);
  const [copiedJson, setCopiedJson] = useState(false);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [editableJsonText, setEditableJsonText] = useState('');
  const [jsonParseError, setJsonParseError] = useState<string | null>(null);
  const [isSimulatingAi, setIsSimulatingAi] = useState(false);
  const [saveSuccessMsg, setSaveSuccessMsg] = useState<string | null>(null);
  const [isEditingSiteInfo, setIsEditingSiteInfo] = useState(false);

  const handleUpdateSiteField = (field: string, val: string) => {
    const existing = getAtsPayload(selectedId) || { ...currentSpec.samplePayloadJson };
    const currentSite = existing.site_info || { ...site };
    const updated = {
      ...existing,
      site_info: {
        ...currentSite,
        [field]: val
      }
    };
    saveAtsPayload(selectedId, updated);
    setPayloadVersion((v) => v + 1);
  };

  // Sync if initialSpecId changes externally or when switching to tev-report
  React.useEffect(() => {
    if (initialSpecId && NETA_ATS_23_MASTER_SPECS[initialSpecId]) {
      setSelectedId(initialSpecId);
      setPayloadVersion((v) => v + 1);
    }
  }, [initialSpecId, displayMode]);

  // Listen for field entry updates from checklists
  React.useEffect(() => {
    const handlePayloadUpdated = (e: any) => {
      setPayloadVersion((v) => v + 1);
    };
    window.addEventListener('tev_ats_payload_updated', handlePayloadUpdated);
    return () => {
      window.removeEventListener('tev_ats_payload_updated', handlePayloadUpdated);
    };
  }, []);

  const hasCustomInput = useMemo(() => {
    return Boolean(getAtsPayload(selectedId));
  }, [selectedId, payloadVersion]);

  const currentSpec: NetaAtsSpecificationDef = useMemo(() => {
    return buildDynamicAtsSpec(selectedId);
  }, [selectedId, payloadVersion]);

  const allSpecs = useMemo(() => getAllAtsSpecsList(), []);

  const isVi = lang === 'vi';
  const site = currentSpec.defaultSiteInfo;

  // Generate markdown report on demand
  const markdownReport = useMemo(() => {
    return generateAtsMarkdownReport(currentSpec, lang);
  }, [currentSpec, lang]);

  const handleCopyMarkdown = () => {
    navigator.clipboard.writeText(markdownReport);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleCopyJson = () => {
    navigator.clipboard.writeText(JSON.stringify(currentSpec.samplePayloadJson, null, 2));
    setCopiedJson(true);
    setTimeout(() => setCopiedJson(false), 2000);
  };

  const handleOpenJsonModal = () => {
    setEditableJsonText(JSON.stringify(currentSpec.samplePayloadJson, null, 2));
    setJsonParseError(null);
    setShowJsonModal(true);
  };

  const handleApplyJson = () => {
    try {
      const parsed = JSON.parse(editableJsonText);
      saveAtsPayload(selectedId, parsed);
      setJsonParseError(null);
      setShowJsonModal(false);
      setSaveSuccessMsg(
        isVi
          ? 'Đã cập nhật và đồng bộ dữ liệu JSON mới vào Báo cáo & Biểu mẫu!'
          : 'Successfully updated and synchronized JSON dataset into Report & Form!'
      );
      setTimeout(() => setSaveSuccessMsg(null), 4000);
    } catch (err: any) {
      setJsonParseError(err?.message || 'Lỗi cú pháp JSON');
    }
  };

  const handleResetToDefault = () => {
    resetAtsPayload(selectedId);
    setSaveSuccessMsg(
      isVi
        ? 'Đã đặt lại dữ liệu về kịch bản tiêu chuẩn ban đầu của ANSI/NETA ATS-2025.'
        : 'Reset back to standard ANSI/NETA ATS-2025 benchmark dataset.'
    );
    setTimeout(() => setSaveSuccessMsg(null), 4000);
  };

  const handleSelectSpec = (id: string) => {
    setSelectedId(id);
    if (onSelectSpecId) {
      onSelectSpecId(id);
    }
  };

  const handleSaveToCmms = () => {
    if (onSaveReport) {
      onSaveReport({
        payloadType: 'NETA_ATS_UNIVERSAL',
        testStandard: currentSpec.standardCode,
        equipTag: site.equipment_tag,
        equipName: isVi ? currentSpec.equipmentTitleVi : currentSpec.equipmentTitleEn,
        overallStatus: currentSpec.overallStatus,
        reportNotes: `${currentSpec.standardCode} - ${isVi ? currentSpec.statusTitleVi : currentSpec.statusTitleEn}`,
        data: currentSpec.samplePayloadJson,
        evaluation: {
          overall_status: currentSpec.overallStatus,
          status_title: isVi ? currentSpec.statusTitleVi : currentSpec.statusTitleEn,
          status_reason: isVi ? currentSpec.statusReasonVi : currentSpec.statusReasonEn,
          markdown_report: markdownReport,
        },
      });
      setSaveSuccessMsg(
        isVi
          ? `Đã lưu thành công biên bản "${currentSpec.equipmentTitleVi}" vào Kho Báo Cáo CMMS!`
          : `Successfully saved "${currentSpec.equipmentTitleEn}" report into CMMS!`
      );
      setTimeout(() => setSaveSuccessMsg(null), 5000);
    }
  };

  const handleSimulateAi = async () => {
    setIsSimulatingAi(true);
    try {
      const activePayload = getAtsPayload(selectedId) || currentSpec.samplePayloadJson;
      if (selectedId === 'neta-liquid-transformer') {
        try {
          const res = await fetch('/api/field-service/neta-liquid-transformer-analyze', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(activePayload),
          });
          const json = await res.json();
          if (json.success && json.data) {
            saveAtsPayload(selectedId, {
              ...activePayload,
              analysis_result: json.data,
            });
          }
        } catch {
          // fallback to client dynamic evaluator
        }
      }
      setPayloadVersion((v) => v + 1);
      setSaveSuccessMsg(
        isVi
          ? 'Đã đối soát hoàn tất: Trạng thái, Bảng phân tích chi tiết, Chẩn đoán rủi ro chuyên sâu & Khuyến nghị đã được cập nhật 100% theo số đo mới nhất!'
          : 'AI evaluation completed: Status, Detailed table, Anomaly risk cards & Recommendations fully re-evaluated per latest measurements!'
      );
      setTimeout(() => setSaveSuccessMsg(null), 5000);
    } finally {
      setIsSimulatingAi(false);
    }
  };

  return (
    <div className="space-y-6">
      {/* 23 SPECIFICATIONS QUICK SELECTOR STRIP */}
      <div className="bg-slate-900 rounded-2xl p-4 border border-slate-800 shadow-md">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-3">
          <div className="flex items-center gap-2">
            <Layers size={18} className="text-amber-400" />
            <h3 className="text-xs font-black text-white uppercase tracking-wider">
              {isVi
                ? 'DANH MỤC 23 BIÊN BẢN KIỂM ĐỊNH CHUYÊN SÂU NETA ATS-2025'
                : 'CATALOG OF 23 ANSI/NETA ATS-2025 ACCEPTANCE TEST SPECIFICATIONS'}
            </h3>
            <span className="px-2 py-0.5 rounded-full text-[10px] font-black bg-blue-500/20 text-blue-300 border border-blue-400/30">
              23 Standards
            </span>
          </div>

          <div className="flex items-center gap-2">
            <span className="text-[11px] text-slate-400 hidden md:inline">
              {isVi ? 'Chọn biên bản nhanh:' : 'Quick Select:'}
            </span>
            <select
              value={selectedId}
              onChange={(e) => handleSelectSpec(e.target.value)}
              className="bg-slate-950 text-white text-xs font-bold px-3 py-1.5 rounded-xl border border-slate-700 focus:outline-none focus:ring-2 focus:ring-blue-500"
            >
              {allSpecs.map((s, idx) => (
                <option key={s.id} value={s.id}>
                  {idx + 1}. {s.sectionCode} - {isVi ? s.equipmentTitleVi.split('(')[0] : s.equipmentTitleEn.split('(')[0]}
                </option>
              ))}
            </select>
          </div>
        </div>

        {/* Scrollable Badges for all 23 */}
        <div className="flex items-center gap-2 overflow-x-auto pb-1 text-xs scrollbar-thin">
          {allSpecs.map((spec, index) => {
            const isActive = spec.id === selectedId;
            return (
              <button
                key={spec.id}
                type="button"
                onClick={() => handleSelectSpec(spec.id)}
                className={`px-3 py-1.5 rounded-xl font-bold whitespace-nowrap transition-all flex items-center gap-1.5 shrink-0 text-[11px] ${
                  isActive
                    ? 'bg-blue-600 text-white shadow-md shadow-blue-600/30 scale-102 border border-blue-400'
                    : 'bg-slate-950/80 hover:bg-slate-800 text-slate-300 border border-slate-800'
                }`}
              >
                <span className={`w-1.5 h-1.5 rounded-full ${spec.overallStatus === 'PASS' ? 'bg-emerald-400' : 'bg-rose-400'}`} />
                <span className="opacity-75 font-mono text-[10px]">{spec.sectionCode}</span>
                <span>{isVi ? spec.equipmentTitleVi.split('(')[0].trim() : spec.equipmentTitleEn.split('(')[0].trim()}</span>
              </button>
            );
          })}
        </div>
      </div>

      {/* Save Success Notification */}
      {saveSuccessMsg && (
        <div className="p-3.5 bg-emerald-50 border border-emerald-300 rounded-xl text-emerald-950 text-xs font-bold flex items-center justify-between gap-3 animate-fade-in shadow-sm">
          <div className="flex items-center gap-2">
            <CheckCircle2 size={16} className="text-emerald-600 shrink-0" />
            <span>{saveSuccessMsg}</span>
          </div>
          {onNavigateToReports && (
            <button
              onClick={onNavigateToReports}
              className="px-3 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-[11px] font-bold flex items-center gap-1 shrink-0"
            >
              <span>{isVi ? 'Xem Kho CMMS' : 'View CMMS'}</span>
              <ArrowRight size={12} />
            </button>
          )}
        </div>
      )}

      {/* MAIN TEV-BRANDED REPORT CONTAINER */}
      <div className="bg-white rounded-2xl border border-slate-200 shadow-xl overflow-hidden">
        {/* OFFICIAL TEV HEADER BANNER */}
        <div className="bg-gradient-to-r from-slate-950 via-slate-900 to-blue-950 p-6 sm:p-7 text-white border-b border-slate-800 space-y-4">
          {/* Row 1: TEV Brand, Category & Standard Badge */}
          <div className="flex items-center justify-between gap-4 flex-wrap">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-blue-600 flex items-center justify-center font-black text-base text-white shadow-md tracking-wider border border-blue-400/40 shrink-0">
                TEV
              </div>
              <div className="flex items-center gap-2.5 flex-wrap">
                <span className="text-xs sm:text-sm font-black uppercase tracking-wider text-blue-400">
                  TEV FLATFORM AI • FIELD SERVICE EVALUATION REPORT
                </span>
                <span className="px-2.5 py-0.5 rounded text-[11px] font-black uppercase tracking-wider bg-white/10 text-slate-200 border border-white/20">
                  {currentSpec.standardCode}
                </span>
              </div>
            </div>

            <div className="hidden sm:flex items-center gap-1.5 px-3 py-1 bg-blue-900/40 border border-blue-400/30 rounded-full text-xs text-blue-200 font-bold">
              <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
              <span>Báo Cáo Nghiệm Thu NETA ATS</span>
            </div>
          </div>

          {/* Row 2: Full-Width Horizontal Equipment Title & Subtitle */}
          <div className="pt-2.5 border-t border-slate-800/80">
            <h2 className="text-xl sm:text-2xl lg:text-3xl font-black text-white tracking-tight leading-snug">
              {isVi ? currentSpec.equipmentTitleVi : currentSpec.equipmentTitleEn}
            </h2>
            <p className="text-xs sm:text-sm text-slate-300 mt-1 max-w-5xl leading-relaxed font-medium">
              {isVi ? currentSpec.subtitleVi : currentSpec.subtitleEn}
            </p>
          </div>

          {/* Row 3: Dedicated Action Toolbar Beneath Equipment Title */}
          <div className="pt-3 border-t border-slate-800 flex items-center gap-2.5 flex-wrap">
            {/* Language Switcher */}
            <div className="flex items-center bg-slate-950 p-1 rounded-xl border border-slate-800">
              <Globe size={13} className="text-slate-400 ml-1.5 mr-1" />
              <button
                type="button"
                onClick={() => setLang('vi')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                  isVi ? 'bg-amber-500 text-slate-950 font-black shadow-xs' : 'text-slate-400 hover:text-white'
                }`}
              >
                🇻🇳 Tiếng Việt
              </button>
              <button
                type="button"
                onClick={() => setLang('en')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                  !isVi ? 'bg-amber-500 text-slate-950 font-black shadow-xs' : 'text-slate-400 hover:text-white'
                }`}
              >
                🇺🇸 English
              </button>
            </div>

            {onSwitchToForm && (
              <button
                type="button"
                onClick={onSwitchToForm}
                className="px-3.5 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-xs font-black flex items-center gap-1.5 shadow-md shadow-blue-600/30 transition-all border border-blue-400"
                title={isVi ? 'Quay lại biểu mẫu nhập liệu chi tiết' : 'Switch to Field Entry Form'}
              >
                <Layers size={14} className="text-amber-300" />
                <span>{isVi ? 'Biểu Mẫu Nhập Liệu 📝' : 'Field Entry Form 📝'}</span>
              </button>
            )}

            <button
              type="button"
              onClick={() => setShowPrintModal(true)}
              className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md shadow-blue-600/30 transition-all"
            >
              <Printer size={15} />
              <span>{isVi ? 'In / Xuất PDF' : 'Print / Export PDF'}</span>
            </button>

            <button
              type="button"
              onClick={handleCopyMarkdown}
              className="px-3.5 py-2 bg-slate-800 hover:bg-slate-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 border border-slate-700 transition-all shadow-sm"
              title="Sao chép toàn bộ báo cáo Markdown"
            >
              {copied ? <Check size={14} className="text-emerald-400" /> : <Copy size={14} />}
              <span>{copied ? (isVi ? 'Đã chép!' : 'Copied!') : (isVi ? 'Sao chép MD' : 'Copy MD')}</span>
            </button>

            <button
              type="button"
              onClick={handleOpenJsonModal}
              className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white rounded-xl text-xs font-bold flex items-center gap-1.5 border border-slate-700 transition-all shadow-sm"
              title="Xem và chỉnh sửa dữ liệu JSON"
            >
              <Code2 size={14} className="text-amber-400" />
              <span>JSON</span>
            </button>

            <button
              type="button"
              onClick={handleSaveToCmms}
              className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md shadow-emerald-600/30 transition-all"
            >
              <Save size={15} />
              <span>{isVi ? 'Lưu Báo Cáo CMMS' : 'Save CMMS'}</span>
            </button>

            <button
              type="button"
              onClick={handleSimulateAi}
              disabled={isSimulatingAi}
              className="px-4 py-2 bg-gradient-to-r from-amber-500 to-yellow-500 hover:from-amber-400 hover:to-yellow-400 text-slate-950 rounded-xl text-xs font-black flex items-center gap-1.5 shadow-md shadow-amber-500/20 transition-all"
            >
              <ShieldCheck size={14} />
              <span>{isSimulatingAi ? (isVi ? 'Đang thẩm định...' : 'Analyzing...') : (isVi ? 'Thẩm Định Kỹ Thuật NETA ATS' : 'NETA ATS Technical Evaluation')}</span>
            </button>
          </div>
        </div>

        {/* REAL-TIME FIELD SYNC NOTIFICATION BAR */}
        {hasCustomInput ? (
          <div className="bg-emerald-50 border-b border-emerald-200 px-6 py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
            <div className="flex items-center gap-2.5 text-emerald-900">
              <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse shrink-0" />
              <div>
                <span className="font-black uppercase tracking-wider text-emerald-800 text-[11px] block">
                  {isVi ? 'ĐÃ ĐỒNG BỘ DỮ LIỆU TỪ BIỂU MẪU NHẬP LIỆU' : 'SYNCHRONIZED WITH FIELD ENTRY FORM'}
                </span>
                <span className="text-slate-700 font-medium">
                  {isVi
                    ? `Báo cáo đang thể hiện số liệu mới nhất vừa nhập ngoài site (${site.equipment_tag || 'Thiết bị'}, ${site.test_date || 'Hôm nay'}).`
                    : `Report is dynamically reflecting your latest site measurements (${site.equipment_tag || 'Equipment'}, ${site.test_date || 'Today'}).`}
                </span>
              </div>
            </div>
            <div className="flex items-center gap-2 shrink-0">
              {onSwitchToForm && (
                <button
                  type="button"
                  onClick={onSwitchToForm}
                  className="px-3 py-1.5 bg-emerald-700 hover:bg-emerald-800 text-white font-bold rounded-lg text-xs flex items-center gap-1 shadow-xs"
                >
                  <span>{isVi ? 'Tiếp tục nhập liệu 📝' : 'Edit Form 📝'}</span>
                </button>
              )}
              <button
                type="button"
                onClick={handleResetToDefault}
                className="px-3 py-1.5 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 font-bold rounded-lg text-xs transition-colors"
                title={isVi ? 'Khôi phục về mẫu ban đầu' : 'Reset to default sample dataset'}
              >
                <span>{isVi ? 'Đặt Lại Mẫu Mặc Định ↺' : 'Reset Sample ↺'}</span>
              </button>
            </div>
          </div>
        ) : (
          <div className="bg-slate-50 border-b border-slate-200 px-6 py-2.5 flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-xs text-slate-600">
            <div className="flex items-center gap-2">
              <Info size={15} className="text-blue-500 shrink-0" />
              <span>
                {isVi
                  ? 'Đang hiển thị kịch bản mẫu tiêu chuẩn ANSI/NETA ATS-2025. Nhập số liệu tại Biểu Mẫu để báo cáo tự động cập nhật số đo thực tế.'
                  : 'Displaying ANSI/NETA ATS-2025 benchmark sample. Enter measurements in the Field Form to auto-populate report.'}
              </span>
            </div>
            {onSwitchToForm && (
              <button
                type="button"
                onClick={onSwitchToForm}
                className="text-blue-600 hover:text-blue-800 font-bold flex items-center gap-1 text-xs shrink-0 self-start sm:self-auto"
              >
                <span>{isVi ? 'Mở Biểu Mẫu Nhập Liệu 📝 →' : 'Open Field Form 📝 →'}</span>
              </button>
            )}
          </div>
        )}

        {/* SECTION 1: OVERALL STATUS */}
        <div
          className={`p-6 border-b text-xs ${
            currentSpec.overallStatus === 'FAIL'
              ? 'bg-rose-50/70 border-rose-200'
              : currentSpec.overallStatus === 'INVESTIGATE'
              ? 'bg-amber-50/70 border-amber-200'
              : 'bg-emerald-50/70 border-emerald-200'
          }`}
        >
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div className="space-y-2">
              <span className="text-[11px] uppercase font-black tracking-wider text-slate-500 block">
                {isVi ? '1. TRẠNG THÁI TỔNG QUAN (Overall Status)' : '1. OVERALL STATUS'}
              </span>
              <div className="flex items-center gap-3">
                <span
                  className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl font-black text-sm tracking-wide uppercase shadow-sm ${
                    currentSpec.overallStatus === 'FAIL'
                      ? 'bg-rose-600 text-white'
                      : currentSpec.overallStatus === 'INVESTIGATE'
                      ? 'bg-amber-600 text-white'
                      : 'bg-emerald-600 text-white'
                  }`}
                >
                  {currentSpec.overallStatus === 'FAIL' && <XCircle size={16} />}
                  {currentSpec.overallStatus === 'INVESTIGATE' && <AlertTriangle size={16} />}
                  {currentSpec.overallStatus === 'PASS' && <CheckCircle2 size={16} />}
                  <span>{isVi ? currentSpec.statusTitleVi : currentSpec.statusTitleEn}</span>
                </span>
              </div>
              <p className="text-xs text-slate-700 leading-relaxed max-w-3xl pt-1">
                {isVi ? currentSpec.statusReasonVi : currentSpec.statusReasonEn}
              </p>
            </div>

            {/* Equipment Site Metadata Box */}
            <div className="text-xs space-y-1.5 bg-white/90 p-3.5 rounded-xl border border-slate-200/90 shadow-xs shrink-0 min-w-[280px]">
              <div className="flex items-center justify-between border-b border-slate-100 pb-1 mb-1">
                <span className="font-bold text-slate-800 flex items-center gap-1">
                  <Database size={13} className="text-blue-600" />
                  {isVi ? 'Thông Tin Site / Master' : 'Site & Master Info'}
                </span>
                <button
                  type="button"
                  onClick={() => setIsEditingSiteInfo(!isEditingSiteInfo)}
                  className="px-2 py-0.5 rounded text-[10px] font-bold bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 transition-colors"
                >
                  {isEditingSiteInfo ? (isVi ? '✕ Đóng' : '✕ Close') : (isVi ? '✏️ Chỉnh Sửa' : '✏️ Edit')}
                </button>
              </div>

              {isEditingSiteInfo ? (
                <div className="space-y-2 pt-1 animate-in fade-in-50 duration-150">
                  <CustomerSelectDropdown
                    label={isVi ? 'Tên Khách Hàng (Master)' : 'Customer (Master)'}
                    value={(site as any).customer_name || ''}
                    onChange={(cName) => handleUpdateSiteField('customer_name', cName)}
                    placeholder={isVi ? 'Chọn hoặc thêm Khách Hàng...' : 'Select or add Customer...'}
                  />

                  <ProjectSelectDropdown
                    label={isVi ? 'Tên Dự Án / Nhà Máy' : 'Project / Plant'}
                    value={site.project_name}
                    customerName={(site as any).customer_name}
                    onChange={(pName) => handleUpdateSiteField('project_name', pName)}
                    placeholder={isVi ? 'Chọn hoặc thêm Dự Án...' : 'Select or add Project...'}
                    standard="ats"
                  />

                  <div>
                    <label className="block text-slate-500 font-semibold mb-0.5 text-[11px]">
                      {isVi ? 'Kỹ Sư Nghiệm Thu (FSE):' : 'Commissioning Engineer:'}
                    </label>
                    <FseEngineerDropdown
                      value={site.fse_name}
                      onChange={(eng) => handleUpdateSiteField('fse_name', eng)}
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-1.5 pt-1">
                    <div>
                      <label className="block text-slate-500 font-semibold mb-0.5 text-[10px]">
                        {isVi ? 'Mã thiết bị:' : 'Tag:'}
                      </label>
                      <input
                        type="text"
                        value={site.equipment_tag}
                        onChange={(e) => handleUpdateSiteField('equipment_tag', e.target.value)}
                        className="w-full text-xs font-mono font-bold px-2 py-1 bg-slate-50 border rounded"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-500 font-semibold mb-0.5 text-[10px]">
                        {isVi ? 'Ngày đo:' : 'Date:'}
                      </label>
                      <input
                        type="date"
                        value={site.test_date}
                        onChange={(e) => handleUpdateSiteField('test_date', e.target.value)}
                        className="w-full text-xs px-2 py-1 bg-slate-50 border rounded"
                      />
                    </div>
                  </div>
                </div>
              ) : (
                <>
                  <div>
                    <strong className="text-slate-500">{isVi ? 'Mã thiết bị:' : 'Equipment Tag:'}</strong>{' '}
                    <span className="font-mono font-bold text-slate-900">{site.equipment_tag}</span>
                  </div>
                  <div>
                    <strong className="text-slate-500">{isVi ? 'Dự án / Nhà máy:' : 'Substation / Plant:'}</strong>{' '}
                    <span className="font-semibold text-slate-900">{site.project_name}</span>
                  </div>
                  <div>
                    <strong className="text-slate-500">{isVi ? 'Công suất / Cấp áp:' : 'Rating / Voltage:'}</strong>{' '}
                    <span className="font-semibold text-slate-900">{site.equipment_rating}</span>
                  </div>
                  <div>
                    <strong className="text-slate-500">{isVi ? 'Kỹ sư Nghiệm thu:' : 'Field Engineer:'}</strong>{' '}
                    <span className="font-semibold text-slate-900">{site.fse_name}</span>
                  </div>
                  <div>
                    <strong className="text-slate-500">{isVi ? 'Ngày thử nghiệm:' : 'Test Date:'}</strong>{' '}
                    <span className="font-semibold text-slate-900">{site.test_date}</span>
                  </div>
                </>
              )}
            </div>
          </div>
        </div>

        {/* SECTION 2: DETAILED EVALUATION TABLE */}
        <div className="p-6 border-b border-slate-200">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-3">
            <h3 className="font-black text-slate-900 text-sm uppercase tracking-wide flex items-center gap-2">
              <span className="w-2.5 h-2.5 rounded-full bg-blue-600"></span>
              <span>
                {isVi
                  ? `2. BẢNG PHÂN TÍCH CHI TIẾT THEO TIÊU CHUẨN ANSI/NETA ATS-2025 (${currentSpec.sectionCode})`
                  : `2. DETAILED EVALUATION TABLE (ANSI/NETA ATS-2025 ${currentSpec.sectionCode})`}
              </span>
            </h3>
            <span className="text-xs text-slate-500 font-semibold">
              {currentSpec.standardCode} & Reference Standards
            </span>
          </div>

          <div className="overflow-x-auto rounded-xl border border-slate-200">
            <table className="w-full text-left text-xs border-collapse min-w-[1050px]">
              <thead>
                <tr className="bg-slate-100 text-slate-800 font-bold border-b border-slate-200 text-xs">
                  <th className="p-4 w-[42%] min-w-[380px]">{isVi ? 'Hạng mục kiểm tra & thử nghiệm' : 'Inspection / Test Item'}</th>
                  <th className="p-4 w-[16%] min-w-[150px]">{isVi ? 'Giá trị đo thực tế' : 'Measured Site Value'}</th>
                  <th className="p-4 w-[22%] min-w-[210px]">{isVi ? 'Tiêu chuẩn NETA ATS-2025' : 'NETA ATS-2025 Standard'}</th>
                  <th className="p-4 w-[11%] min-w-[110px]">{isVi ? 'Tham chiếu kỹ thuật' : 'Reference Standard'}</th>
                  <th className="p-4 w-[9%] min-w-[95px] text-center">{isVi ? 'Kết luận' : 'Status'}</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {currentSpec.evaluationItems.map((item, idx) => (
                  <tr
                    key={idx}
                    className={
                      item.status === 'FAIL'
                        ? 'bg-rose-50/40 hover:bg-rose-50/60'
                        : item.status === 'INVESTIGATE'
                        ? 'bg-amber-50/40 hover:bg-amber-50/60'
                        : 'hover:bg-slate-50/80 transition-colors'
                    }
                  >
                    <td className="p-4 font-bold text-slate-900 align-top">
                      <div className="text-xs sm:text-sm text-slate-900 font-bold leading-snug break-words">
                        {isVi ? item.item : item.itemEn}
                      </div>
                      <div className="text-[11px] text-slate-600 font-normal pt-1.5 leading-relaxed break-words">
                        {isVi ? item.detail : item.detailEn}
                      </div>
                    </td>
                    <td className="p-4 font-mono font-bold text-slate-900 whitespace-nowrap align-top">
                      <span className="bg-slate-100 px-2 py-1 rounded border border-slate-200/80 inline-block">
                        {item.measured}
                      </span>
                    </td>
                    <td className="p-4 text-slate-800 align-top leading-relaxed break-words text-xs">
                      {isVi ? item.standard : item.standardEn}
                    </td>
                    <td className="p-4 text-slate-600 font-mono text-[11px] align-top whitespace-nowrap">
                      <span className="bg-blue-50/60 text-blue-900 px-2 py-0.5 rounded border border-blue-200/50">
                        {item.reference}
                      </span>
                    </td>
                    <td className="p-4 text-center whitespace-nowrap align-top">
                      <span
                        className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full font-black text-[11px] uppercase shadow-2xs ${
                          item.status === 'PASS'
                            ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                            : item.status === 'INVESTIGATE'
                            ? 'bg-amber-100 text-amber-800 border border-amber-300'
                            : 'bg-rose-100 text-rose-800 border border-rose-300'
                        }`}
                      >
                        {item.status === 'PASS' && <CheckCircle2 size={13} />}
                        {item.status === 'INVESTIGATE' && <AlertTriangle size={13} />}
                        {item.status === 'FAIL' && <XCircle size={13} />}
                        <span>{item.status}</span>
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* SECTION 3: IN-DEPTH ANOMALY & RISK DIAGNOSTICS */}
        <div className="p-6 border-b border-slate-200 bg-slate-50/50 space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="font-black text-slate-900 text-sm uppercase tracking-wide flex items-center gap-2">
              <span className="w-2.5 h-2.5 rounded-full bg-amber-500"></span>
              <span>
                {isVi
                  ? '3. PHÂN TÍCH CHUYÊN SÂU & RỦI RO BẤT THƯỜNG (In-Depth Anomaly & Risk Analysis)'
                  : '3. IN-DEPTH ANOMALY & RISK DIAGNOSTIC ANALYSIS'}
              </span>
            </h3>
            <span className="text-xs text-slate-500 font-semibold hidden sm:inline">
              4 Diagnostic Dimensions per IEEE / ASTM / NFPA
            </span>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
            {currentSpec.anomalyCards.map((card, idx) => (
              <div
                key={idx}
                className={`p-4 rounded-xl border space-y-2 shadow-xs bg-white ${
                  card.severity === 'critical'
                    ? 'border-rose-200 border-l-4 border-l-rose-600'
                    : card.severity === 'warning'
                    ? 'border-amber-200 border-l-4 border-l-amber-500'
                    : 'border-slate-200 border-l-4 border-l-emerald-600'
                }`}
              >
                <div className="flex items-center justify-between">
                  <div
                    className={`font-bold text-xs flex items-center gap-1.5 ${
                      card.severity === 'critical'
                        ? 'text-rose-700'
                        : card.severity === 'warning'
                        ? 'text-amber-700'
                        : 'text-slate-900'
                    }`}
                  >
                    {card.category === 'insulation' && <Zap size={15} />}
                    {card.category === 'contact' && <Gauge size={15} />}
                    {card.category === 'thermal' && <Flame size={15} />}
                    {card.category === 'operational' && <Activity size={15} />}
                    <span>{isVi ? card.titleVi : card.titleEn}</span>
                  </div>

                  <span
                    className={`px-2 py-0.5 rounded text-[9px] font-black uppercase tracking-wider ${
                      card.severity === 'critical'
                        ? 'bg-rose-100 text-rose-800'
                        : card.severity === 'warning'
                        ? 'bg-amber-100 text-amber-800'
                        : 'bg-emerald-100 text-emerald-800'
                    }`}
                  >
                    {card.severity}
                  </span>
                </div>

                <ul className="space-y-1.5 text-slate-700 leading-relaxed list-disc pl-4 text-[11px]">
                  {(isVi ? card.pointsVi : card.pointsEn).map((pt, pIdx) => (
                    <li key={pIdx}>{pt}</li>
                  ))}
                </ul>
              </div>
            ))}
          </div>
        </div>

        {/* SECTION 4: ACTIONABLE RECOMMENDATIONS */}
        <div className="p-6 bg-white space-y-3">
          <div className="flex items-center justify-between">
            <h3 className="font-black text-slate-900 text-sm uppercase tracking-wide flex items-center gap-2">
              <span className="w-2.5 h-2.5 rounded-full bg-emerald-600"></span>
              <span>
                {isVi
                  ? '4. KHUYẾN NGHỊ HÀNH ĐỘNG CHI TIẾT CHO FSE (Actionable Recommendations)'
                  : '4. ACTIONABLE FIELD SERVICE ENGINEERING RECOMMENDATIONS'}
              </span>
            </h3>
            <span className="text-xs text-slate-500 font-semibold">
              Field Execution Protocol
            </span>
          </div>

          <div
            className={`p-4 rounded-xl border ${
              currentSpec.overallStatus === 'FAIL'
                ? 'bg-rose-50/40 border-rose-200'
                : currentSpec.overallStatus === 'INVESTIGATE'
                ? 'bg-amber-50/40 border-amber-200'
                : 'bg-emerald-50/40 border-emerald-200'
            }`}
          >
            <ol className="space-y-2.5 text-xs text-slate-800 list-decimal pl-5 leading-relaxed font-medium">
              {(isVi ? currentSpec.recommendationsVi : currentSpec.recommendationsEn).map((rec, rIdx) => (
                <li key={rIdx}>
                  <span>{rec}</span>
                </li>
              ))}
            </ol>
          </div>
        </div>

        {/* FOOTER BAR WITH CERTIFICATE SEAL */}
        <div className="p-4 bg-slate-50 border-t border-slate-200 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-500">
          <div className="flex items-center gap-2">
            <ShieldCheck size={16} className="text-blue-600" />
            <span className="font-mono text-[11px]">
              CERTIFICATE: TEV-ATS-{site.equipment_tag}-{site.test_date.replace(/-/g, '')}
            </span>
          </div>

          <div className="flex items-center gap-3">
            <span>{isVi ? 'Hệ thống định dạng chuẩn hóa NETA ATS 4 Phần' : 'Standardized NETA ATS 4-Part Architecture'}</span>
            <button
              type="button"
              onClick={() => setShowPrintModal(true)}
              className="text-blue-600 hover:text-blue-700 font-bold flex items-center gap-1"
            >
              <Printer size={13} />
              <span>{isVi ? 'Xuất bản in chính thức' : 'Export Official Printout'}</span>
            </button>
          </div>
        </div>
      </div>

      {/* MODAL 1: TEV OFFICIAL PRINT & PDF MODAL */}
      {showPrintModal && (
        <TevReportPrintModal
          isOpen={showPrintModal}
          onClose={() => setShowPrintModal(false)}
          title={currentSpec.equipmentTitleVi}
          titleEn={currentSpec.equipmentTitleEn}
          standard={currentSpec.standardCode}
          siteInfo={{
            project_name: site.project_name,
            equipment_tag: site.equipment_tag,
            equipment_rating: site.equipment_rating,
            fse_name: site.fse_name,
            test_date: site.test_date,
          }}
          overallStatus={currentSpec.overallStatus}
          statusReason={currentSpec.statusReasonVi}
          statusReasonEn={currentSpec.statusReasonEn}
          electricalTable={currentSpec.evaluationItems.map((item) => ({
            item: item.item,
            itemEn: item.itemEn,
            measured: item.measured,
            standard: item.standard,
            standardEn: item.standardEn,
            reference: item.reference,
            status: item.status,
          }))}
          warnings={currentSpec.anomalyCards.flatMap((c) => c.pointsVi)}
          warningsEn={currentSpec.anomalyCards.flatMap((c) => c.pointsEn)}
          recommendations={currentSpec.recommendationsVi}
          recommendationsEn={currentSpec.recommendationsEn}
          lang={lang}
        />
      )}

      {/* MODAL 2: JSON PAYLOAD MODAL */}
      {showJsonModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/70 backdrop-blur-xs">
          <div className="bg-white rounded-2xl max-w-2xl w-full p-6 shadow-2xl space-y-4 max-h-[90vh] flex flex-col">
            <div className="flex items-center justify-between border-b border-slate-200 pb-3">
              <div className="flex items-center gap-2">
                <Code2 size={18} className="text-amber-600" />
                <h3 className="text-sm font-bold text-slate-900">
                  {isVi ? 'Dữ Liệu JSON Chuẩn Đầu Vào từ FSE' : 'FSE Field Input JSON Payload'}
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setShowJsonModal(false)}
                className="text-slate-400 hover:text-slate-600 font-bold text-lg"
              >
                ✕
              </button>
            </div>

            <p className="text-xs text-slate-500">
              {isVi
                ? 'Xem hoặc chỉnh sửa trực tiếp thông số hiện trường dạng JSON. Nhấn "Áp Dụng Dữ Liệu" để cập nhật ngay vào Báo cáo & Biểu mẫu:'
                : 'View or edit field parameters directly in JSON. Click "Apply Dataset" to update Report & Form immediately:'}
            </p>

            {jsonParseError && (
              <div className="p-2.5 bg-rose-50 border border-rose-300 rounded-lg text-rose-800 text-xs font-bold flex items-center gap-2">
                <AlertTriangle size={15} className="text-rose-600 shrink-0" />
                <span>{jsonParseError}</span>
              </div>
            )}

            <textarea
              value={editableJsonText}
              onChange={(e) => setEditableJsonText(e.target.value)}
              rows={14}
              className="p-4 bg-slate-950 text-emerald-400 rounded-xl text-xs font-mono w-full border border-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500 leading-relaxed resize-y"
              spellCheck={false}
            />

            <div className="flex items-center justify-between pt-2 flex-wrap gap-2">
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handleCopyJson}
                  className="px-3.5 py-2 bg-slate-800 hover:bg-slate-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-sm"
                >
                  {copiedJson ? <Check size={14} /> : <Copy size={14} />}
                  <span>{copiedJson ? (isVi ? 'Đã sao chép!' : 'Copied!') : (isVi ? 'Sao chép JSON' : 'Copy JSON')}</span>
                </button>
                <button
                  type="button"
                  onClick={handleResetToDefault}
                  className="px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold"
                  title="Khôi phục về mẫu ban đầu"
                >
                  <span>{isVi ? 'Khôi Phục Mẫu ↺' : 'Reset Sample ↺'}</span>
                </button>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setShowJsonModal(false)}
                  className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-800 rounded-xl text-xs font-bold"
                >
                  {isVi ? 'Hủy' : 'Cancel'}
                </button>
                <button
                  type="button"
                  onClick={handleApplyJson}
                  className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-xs font-black shadow-md shadow-blue-600/30 flex items-center gap-1.5"
                >
                  <Sparkles size={14} className="text-amber-300" />
                  <span>{isVi ? 'Áp Dụng Dữ Liệu Vào Báo Cáo ⚡' : 'Apply Dataset to Report ⚡'}</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
