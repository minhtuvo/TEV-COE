import React, { useState, useMemo } from 'react';
import {
  ShieldAlert,
  ShieldCheck,
  AlertTriangle,
  AlertOctagon,
  CheckCircle2,
  Sparkles,
  Download,
  Copy,
  Check,
  RotateCcw,
  Info,
  Layers,
  FileCheck2,
  Save,
  Zap,
  Activity,
  Gauge,
  Thermometer,
  Wrench,
  FileText,
  Sliders,
  Share2,
  ArrowRight,
  ExternalLink,
  ChevronDown,
  Calendar,
  Building,
  UserCheck,
  HelpCircle,
  Radio,
  Split
} from 'lucide-react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  NetaMtsBuswayInputPayload,
  NetaMtsBuswayEvaluationResult,
  performDeterministicMtsBuswayAnalysis,
  MTS_BUSWAY_SYSTEM_INSTRUCTION,
  SAMPLE_MTS_BUSWAY_DEFECT_PAYLOAD,
  SAMPLE_MTS_BUSWAY_PASS_PAYLOAD,
  SAMPLE_MTS_BUSWAY_CAUTION_PAYLOAD
} from '../../server/netaMtsBuswayAnalyzer';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
  initialLanguage?: 'vi' | 'en';
}

export const NetaMtsBuswayChecklist: React.FC<Props> = ({
  onSaveReport,
  onNavigateToReports,
  initialLanguage = 'vi'
}) => {
  const [lang, setLang] = useState<'vi' | 'en'>(initialLanguage);
  const [activeTab, setActiveTab] = useState<'checklist' | 'evaluation' | 'prompt' | 'json'>('checklist');
  const [activePhase, setActivePhase] = useState<'phase1' | 'phase2' | 'phase3'>('phase1');
  const [isEvaluating, setIsEvaluating] = useState(false);
  const [evaluationResult, setEvaluationResult] = useState<NetaMtsBuswayEvaluationResult | null>(null);
  const [copySuccess, setCopySuccess] = useState(false);
  const [jsonCopySuccess, setJsonCopySuccess] = useState(false);
  const [saveSuccess, setSaveSuccess] = useState(false);

  // Form State initialized with Defect scenario
  const [formData, setFormData] = useState<NetaMtsBuswayInputPayload>(SAMPLE_MTS_BUSWAY_DEFECT_PAYLOAD);

  const isVi = lang === 'vi';

  // Handler to load presets
  const handleLoadPreset = (scenario: 'defect' | 'pass' | 'caution') => {
    let payload = SAMPLE_MTS_BUSWAY_DEFECT_PAYLOAD;
    if (scenario === 'pass') payload = SAMPLE_MTS_BUSWAY_PASS_PAYLOAD;
    if (scenario === 'caution') payload = SAMPLE_MTS_BUSWAY_CAUTION_PAYLOAD;

    setFormData(JSON.parse(JSON.stringify(payload)));
    setEvaluationResult(null);
  };

  // Run AI / Deterministic Evaluation
  const handleRunEvaluation = async () => {
    setIsEvaluating(true);
    try {
      const response = await fetch('/api/field-service/neta-mts-busway-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(formData),
      });

      if (response.ok) {
        const result = await response.json();
        if (result.success && result.data) {
          setEvaluationResult(result.data);
          setActiveTab('evaluation');
          return;
        }
      }
      // Fallback to client-side deterministic evaluation
      const localEval = performDeterministicMtsBuswayAnalysis(formData);
      setEvaluationResult(localEval);
      setActiveTab('evaluation');
    } catch (err) {
      console.warn('API error, using local deterministic engine:', err);
      const localEval = performDeterministicMtsBuswayAnalysis(formData);
      setEvaluationResult(localEval);
      setActiveTab('evaluation');
    } finally {
      setIsEvaluating(false);
    }
  };

  // Copy AI Studio System Instruction Prompt
  const handleCopyPrompt = () => {
    navigator.clipboard.writeText(MTS_BUSWAY_SYSTEM_INSTRUCTION);
    setCopySuccess(true);
    setTimeout(() => setCopySuccess(false), 2000);
  };

  // Copy JSON Input
  const handleCopyJson = () => {
    navigator.clipboard.writeText(JSON.stringify(formData, null, 2));
    setJsonCopySuccess(true);
    setTimeout(() => setJsonCopySuccess(false), 2000);
  };

  // Save Report
  const handleSaveReport = () => {
    if (!evaluationResult) {
      const computed = performDeterministicMtsBuswayAnalysis(formData);
      setEvaluationResult(computed);
    }
    const reportData = {
      id: `NETA-MTS-BUSWAY-${Date.now()}`,
      moduleType: 'MTS-BUSWAY-7.4',
      title: `${formData.tev_coe_header.equipment_tag} - NETA MTS-2023 Sec 7.4`,
      equipmentTag: formData.tev_coe_header.equipment_tag,
      projectName: formData.tev_coe_header.project_name,
      fseName: formData.tev_coe_header.fse_name,
      testDate: formData.tev_coe_header.test_date,
      overallStatus: evaluationResult?.overall_status || 'CRITICAL',
      data: formData,
      evaluation: evaluationResult,
      createdAt: new Date().toISOString()
    };

    if (onSaveReport) {
      onSaveReport(reportData);
    }
    setSaveSuccess(true);
    setTimeout(() => setSaveSuccess(false), 2500);
  };

  // PDF Export
  const handleExportPdf = () => {
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const currentEval = evaluationResult || performDeterministicMtsBuswayAnalysis(formData);

    // Header
    doc.setFillColor(15, 23, 42);
    doc.rect(0, 0, 210, 26, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(13);
    doc.text('TEV COE - ELECTRICAL TESTING & COMMISSIONING PLATFORM', 14, 11);
    doc.setFontSize(9.5);
    doc.setTextColor(148, 163, 184);
    doc.text('ANSI/NETA MTS-2023 SECTION 7.4 - METAL-ENCLOSED BUSWAYS (LV & MV)', 14, 18);

    // Info Box
    doc.setFillColor(248, 250, 252);
    doc.rect(14, 30, 182, 32, 'F');
    doc.setDrawColor(203, 213, 225);
    doc.rect(14, 30, 182, 32, 'S');

    doc.setTextColor(15, 23, 42);
    doc.setFontSize(8.5);
    doc.text(`Project: ${formData.tev_coe_header.project_name}`, 18, 38);
    doc.text(`Equipment Tag: ${formData.tev_coe_header.equipment_tag}`, 18, 44);
    doc.text(`Rating: ${formData.tev_coe_header.voltage_rating_kV} kV | ${formData.tev_coe_header.current_rating_A} A | ${formData.tev_coe_header.busbar_material}`, 18, 50);
    doc.text(`Enclosure: ${formData.tev_coe_header.enclosure_type || 'Aluminum'} (${formData.tev_coe_header.ip_rating || 'IP54'})`, 18, 56);

    doc.text(`FSE Engineer: ${formData.tev_coe_header.fse_name}`, 115, 38);
    doc.text(`Test Date: ${formData.tev_coe_header.test_date}`, 115, 44);
    doc.text(`Location: ${formData.tev_coe_header.location_bay || 'Main Feeder'}`, 115, 50);
    doc.text(`Standard: ANSI/NETA MTS-2023 (Sec 7.4) & IEEE C37.23`, 115, 56);

    // Severity Box
    let sevColor: [number, number, number] = [16, 185, 129];
    if (currentEval.overall_status === 'CRITICAL') sevColor = [225, 29, 72];
    if (currentEval.overall_status === 'CAUTION') sevColor = [245, 158, 11];

    doc.setFillColor(...sevColor);
    doc.rect(14, 66, 182, 9, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(9.5);
    const sevTitle = isVi ? currentEval.overall_title_vi : currentEval.overall_title_en;
    doc.text(`OVERALL STATUS: ${currentEval.overall_status} - ${sevTitle}`, 18, 72.5);

    // Evaluation Table
    const tableRows = currentEval.evaluation_table.map((row) => [
      isVi ? row.item_vn : row.item,
      row.standard_ref,
      row.measured_value,
      row.status,
      isVi ? row.notes_vn : row.notes
    ]);

    autoTable(doc, {
      startY: 79,
      head: [['Inspection & Test Item', 'Standard Ref', 'Measured / Observed Value', 'Status', 'Technical Evaluation']],
      body: tableRows,
      theme: 'grid',
      headStyles: { fillColor: [30, 41, 59], textColor: [255, 255, 255], fontSize: 8 },
      bodyStyles: { fontSize: 7.5, cellPadding: 2 },
      columnStyles: {
        0: { cellWidth: 42 },
        1: { cellWidth: 28 },
        2: { cellWidth: 45 },
        3: { cellWidth: 18, halign: 'center' },
        4: { cellWidth: 49 },
      },
      didParseCell: (data) => {
        if (data.column.index === 3 && data.section === 'body') {
          const val = data.cell.raw;
          if (val === 'CRITICAL' || val === 'FAIL') {
            data.cell.styles.textColor = [225, 29, 72];
            data.cell.styles.fontStyle = 'bold';
          } else if (val === 'PASS') {
            data.cell.styles.textColor = [16, 185, 129];
            data.cell.styles.fontStyle = 'bold';
          } else if (val === 'CAUTION') {
            data.cell.styles.textColor = [217, 119, 6];
            data.cell.styles.fontStyle = 'bold';
          }
        }
      }
    });

    const finalY = (doc as any).lastAutoTable.finalY + 6;

    // Recommendations
    doc.setTextColor(15, 23, 42);
    doc.setFontSize(9);
    doc.setFont(undefined, 'bold');
    doc.text(
      isVi ? 'KHUYẾN NGHỊ HÀNH ĐỘNG DÀNH CHO KỸ SƯ HIỆN TRƯỜNG (FSE):' : 'ACTIONABLE TECHNICAL RECOMMENDATIONS FOR FIELD ENGINEERS (FSE):',
      14,
      finalY
    );

    doc.setFont(undefined, 'normal');
    doc.setFontSize(8);
    const recs = isVi ? currentEval.recommendations_vi : currentEval.recommendations_en;
    let textY = finalY + 5;
    recs.slice(0, 4).forEach((rec, idx) => {
      const splitText = doc.splitTextToSize(`${idx + 1}. ${rec}`, 182);
      doc.text(splitText, 14, textY);
      textY += splitText.length * 4.2;
    });

    // Signature Area
    const sigY = Math.max(textY + 6, 260);
    doc.setDrawColor(203, 213, 225);
    doc.line(14, sigY, 196, sigY);

    doc.setFontSize(8);
    doc.setTextColor(100, 116, 139);
    doc.text(`Lead Testing Engineer: ${formData.tev_coe_header.fse_name}`, 14, sigY + 6);
    doc.text('Authorized Commissioning Lead Signature: __________________________', 95, sigY + 6);
    doc.text(`Generated: ${new Date().toLocaleDateString()} | TEV COE Analytical Engine`, 14, sigY + 11);

    doc.save(`NETA_MTS_BUSWAY_${formData.tev_coe_header.equipment_tag.replace(/[^a-zA-Z0-9_-]/g, '_')}.pdf`);
  };

  // Helper for quick stats
  const stats = useMemo(() => {
    const computed = evaluationResult || performDeterministicMtsBuswayAnalysis(formData);
    const total = computed.evaluation_table.length;
    const passCount = computed.evaluation_table.filter(t => t.status === 'PASS').length;
    const failCount = computed.evaluation_table.filter(t => t.status === 'FAIL' || t.status === 'CRITICAL').length;
    const cautionCount = computed.evaluation_table.filter(t => t.status === 'CAUTION').length;
    return { total, passCount, failCount, cautionCount, status: computed.overall_status };
  }, [formData, evaluationResult]);

  return (
    <div className="bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden">
      {/* Top Header Banner */}
      <div className="bg-gradient-to-r from-slate-900 via-teal-950 to-slate-900 px-6 py-5 text-white border-b border-teal-800/40">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="space-y-1">
            <div className="flex items-center gap-2">
              <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-teal-500/20 text-teal-300 border border-teal-400/30">
                <Layers className="w-3.5 h-3.5 text-teal-400" />
                ANSI/NETA MTS-2023 Sec 7.4
              </span>
              <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-medium bg-emerald-500/10 text-emerald-300 border border-emerald-500/20">
                <CheckCircle2 className="w-3 h-3 text-emerald-400" />
                Tables 100.1, 100.12, 100.17, 100.18, 100.23
              </span>
            </div>
            <h2 className="text-xl md:text-2xl font-bold tracking-tight text-white flex items-center gap-2">
              {isVi
                ? 'Bảo Trì & Thử Nghiệm Thanh Cái Vỏ Kim Loại (Metal-Enclosed Busways)'
                : 'Maintenance & Testing of Metal-Enclosed Busways (LV & MV)'}
            </h2>
            <p className="text-xs md:text-sm text-slate-300 max-w-3xl">
              {isVi
                ? 'Đo điện trở sụt áp mối nối Bus Joint (độ lệch ≤ 50%), điện trở cách điện Megohm, thử nghiệm cao áp Hi-Pot, thứ tự pha và khảo sát phóng điện cục bộ online theo NETA MTS-2023 Mục 7.4.'
                : 'Bus joint contact resistance (deviation ≤ 50%), insulation resistance megohms, dielectric withstand test, phasing verification, and online PD survey per NETA MTS-2023 Section 7.4.'}
            </p>
          </div>

          {/* Quick Action Tools */}
          <div className="flex flex-wrap items-center gap-2 self-start md:self-center">
            {/* Language toggle */}
            <div className="flex rounded-lg bg-slate-800/80 p-0.5 border border-slate-700">
              <button
                type="button"
                onClick={() => setLang('vi')}
                className={`px-2.5 py-1 text-xs font-medium rounded-md transition-all ${
                  lang === 'vi' ? 'bg-teal-600 text-white shadow-sm' : 'text-slate-400 hover:text-white'
                }`}
              >
                VN
              </button>
              <button
                type="button"
                onClick={() => setLang('en')}
                className={`px-2.5 py-1 text-xs font-medium rounded-md transition-all ${
                  lang === 'en' ? 'bg-teal-600 text-white shadow-sm' : 'text-slate-400 hover:text-white'
                }`}
              >
                EN
              </button>
            </div>

            {/* Export PDF */}
            <button
              type="button"
              onClick={handleExportPdf}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-xs font-medium text-white border border-slate-700 shadow-sm transition"
              title="Xuất Báo Cáo PDF Thử Nghiệm Busway Chuẩn ANSI/NETA"
            >
              <Download className="w-3.5 h-3.5 text-teal-400" />
              <span>{isVi ? 'Xuất PDF' : 'PDF Report'}</span>
            </button>

            {/* Save Report */}
            <button
              type="button"
              onClick={handleSaveReport}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium transition shadow-sm ${
                saveSuccess
                  ? 'bg-emerald-600 text-white border border-emerald-500'
                  : 'bg-teal-700 hover:bg-teal-600 text-white border border-teal-500'
              }`}
            >
              {saveSuccess ? <Check className="w-3.5 h-3.5" /> : <Save className="w-3.5 h-3.5" />}
              <span>{saveSuccess ? (isVi ? 'Đã Lưu!' : 'Saved!') : (isVi ? 'Lưu Báo Cáo' : 'Save Report')}</span>
            </button>
          </div>
        </div>

        {/* Quick Scenario Preset Selector Bar */}
        <div className="mt-4 pt-3 border-t border-teal-800/40 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex items-center gap-2">
            <span className="text-slate-300 font-medium">{isVi ? 'Kịch bản mẫu:' : 'Simulation Presets:'}</span>
            <div className="inline-flex rounded-lg bg-slate-800/90 p-0.5 border border-slate-700">
              <button
                type="button"
                onClick={() => handleLoadPreset('defect')}
                className="px-2.5 py-1 rounded-md text-rose-300 hover:bg-rose-500/20 font-medium transition flex items-center gap-1"
              >
                <AlertOctagon className="w-3 h-3 text-rose-400" />
                <span>{isVi ? 'Tòa Nhà TEV (Sự cố Joint 05 & PD 32dB)' : 'TEV Tower Defect (Joint 05 High-R)'}</span>
              </button>
              <button
                type="button"
                onClick={() => handleLoadPreset('pass')}
                className="px-2.5 py-1 rounded-md text-emerald-300 hover:bg-emerald-500/20 font-medium transition flex items-center gap-1"
              >
                <CheckCircle2 className="w-3 h-3 text-emerald-400" />
                <span>{isVi ? 'Nhà Máy Cleanroom 24kV (Đạt Chuẩn)' : '24kV Cleanroom (Pass)'}</span>
              </button>
              <button
                type="button"
                onClick={() => handleLoadPreset('caution')}
                className="px-2.5 py-1 rounded-md text-amber-300 hover:bg-amber-500/20 font-medium transition flex items-center gap-1"
              >
                <AlertTriangle className="w-3 h-3 text-amber-400" />
                <span>{isVi ? 'Tuyến Hạ Thế 0.4kV (Lệch Áp 43% & Mất Sấy)' : '0.4kV LV Riser (Caution)'}</span>
              </button>
            </div>
          </div>

          {/* Quick Header Summary Pill */}
          <div className="flex items-center gap-2">
            <span className="text-slate-400">
              {formData.tev_coe_header.equipment_tag} ({formData.tev_coe_header.voltage_rating_kV}kV - {formData.tev_coe_header.current_rating_A}A)
            </span>
            <span
              className={`px-2 py-0.5 rounded text-xs font-bold uppercase tracking-wider ${
                stats.status === 'CRITICAL'
                  ? 'bg-rose-500 text-white'
                  : stats.status === 'CAUTION'
                  ? 'bg-amber-500 text-slate-900'
                  : 'bg-emerald-500 text-white'
              }`}
            >
              {stats.status}
            </span>
          </div>
        </div>
      </div>

      {/* Main Navigation Tabs */}
      <div className="border-b border-slate-200 bg-slate-50 px-6 flex items-center justify-between">
        <div className="flex space-x-1 sm:space-x-4">
          <button
            type="button"
            onClick={() => setActiveTab('checklist')}
            className={`py-3.5 px-3 border-b-2 font-medium text-xs sm:text-sm flex items-center gap-2 transition ${
              activeTab === 'checklist'
                ? 'border-teal-600 text-teal-700 bg-white shadow-sm'
                : 'border-transparent text-slate-600 hover:text-slate-900 hover:border-slate-300'
            }`}
          >
            <Layers className="w-4 h-4" />
            <span>{isVi ? '1. Phiếu Khảo Sát & Nhập Liệu' : '1. Checklist & Data Entry'}</span>
          </button>

          <button
            type="button"
            onClick={() => setActiveTab('evaluation')}
            className={`py-3.5 px-3 border-b-2 font-medium text-xs sm:text-sm flex items-center gap-2 transition ${
              activeTab === 'evaluation'
                ? 'border-teal-600 text-teal-700 bg-white shadow-sm'
                : 'border-transparent text-slate-600 hover:text-slate-900 hover:border-slate-300'
            }`}
          >
            <Sparkles className="w-4 h-4 text-teal-600" />
            <span>{isVi ? '2. Kết Quả Chẩn Đoán AI & NETA' : '2. AI & NETA Diagnostic Evaluation'}</span>
            {stats.failCount > 0 && (
              <span className="px-1.5 py-0.2 rounded-full text-xs font-bold bg-rose-100 text-rose-700">
                {stats.failCount}
              </span>
            )}
          </button>

          <button
            type="button"
            onClick={() => setActiveTab('prompt')}
            className={`py-3.5 px-3 border-b-2 font-medium text-xs sm:text-sm flex items-center gap-2 transition ${
              activeTab === 'prompt'
                ? 'border-teal-600 text-teal-700 bg-white shadow-sm'
                : 'border-transparent text-slate-600 hover:text-slate-900 hover:border-slate-300'
            }`}
          >
            <FileText className="w-4 h-4" />
            <span>{isVi ? '3. System Prompt AI Studio' : '3. AI System Prompt'}</span>
          </button>

          <button
            type="button"
            onClick={() => setActiveTab('json')}
            className={`py-3.5 px-3 border-b-2 font-medium text-xs sm:text-sm flex items-center gap-2 transition ${
              activeTab === 'json'
                ? 'border-teal-600 text-teal-700 bg-white shadow-sm'
                : 'border-transparent text-slate-600 hover:text-slate-900 hover:border-slate-300'
            }`}
          >
            <Sliders className="w-4 h-4" />
            <span>{isVi ? '4. JSON Payload' : '4. JSON Payload'}</span>
          </button>
        </div>

        {/* Big Trigger Evaluation Button */}
        <button
          type="button"
          onClick={handleRunEvaluation}
          disabled={isEvaluating}
          className="inline-flex items-center gap-2 px-4 py-2 rounded-xl bg-gradient-to-r from-teal-600 to-emerald-600 hover:from-teal-700 hover:to-emerald-700 text-white text-xs sm:text-sm font-semibold shadow-md transition transform active:scale-95 disabled:opacity-50"
        >
          <Sparkles className={`w-4 h-4 ${isEvaluating ? 'animate-spin' : ''}`} />
          <span>
            {isEvaluating
              ? isVi ? 'Đang Phân Tích AI...' : 'Analyzing...'
              : isVi ? 'Chẩn Đoán AI NETA' : 'Run AI Evaluation'}
          </span>
        </button>
      </div>

      {/* TAB CONTENT 1: CHECKLIST & DATA ENTRY */}
      {activeTab === 'checklist' && (
        <div className="p-6 space-y-6">
          {/* Header Metadata Section */}
          <div className="bg-slate-50 rounded-xl p-4 border border-slate-200">
            <h3 className="text-xs font-bold uppercase tracking-wider text-slate-600 mb-3 flex items-center gap-2">
              <Building className="w-4 h-4 text-teal-600" />
              {isVi ? 'Thông Tin Nhãn Mác Tuyến Thanh Cái Vỏ Kim Loại (Busway Parameters)' : 'Metal-Enclosed Busway Technical Parameters'}
            </h3>
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 text-xs">
              <div>
                <label className="block font-medium text-slate-700 mb-1">{isVi ? 'Tên Dự Án / Tòa Nhà' : 'Project / Building'}</label>
                <input
                  type="text"
                  value={formData.tev_coe_header.project_name}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      tev_coe_header: { ...formData.tev_coe_header, project_name: e.target.value }
                    })
                  }
                  className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-teal-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-medium text-slate-700 mb-1">{isVi ? 'Mã Ký Hiệu Tuyến Busway (Tag)' : 'Busway Tag'}</label>
                <input
                  type="text"
                  value={formData.tev_coe_header.equipment_tag}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      tev_coe_header: { ...formData.tev_coe_header, equipment_tag: e.target.value }
                    })
                  }
                  className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-teal-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-medium text-slate-700 mb-1">{isVi ? 'Điện Áp Định Mức (kV)' : 'Voltage Rating (kV)'}</label>
                <input
                  type="number"
                  step={0.1}
                  value={formData.tev_coe_header.voltage_rating_kV}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      tev_coe_header: { ...formData.tev_coe_header, voltage_rating_kV: Number(e.target.value) }
                    })
                  }
                  className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 font-bold text-teal-700 focus:ring-2 focus:ring-teal-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-medium text-slate-700 mb-1">{isVi ? 'Dòng Điện Định Mức (A)' : 'Current Rating (A)'}</label>
                <input
                  type="number"
                  value={formData.tev_coe_header.current_rating_A}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      tev_coe_header: { ...formData.tev_coe_header, current_rating_A: Number(e.target.value) }
                    })
                  }
                  className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 font-bold text-teal-700 focus:ring-2 focus:ring-teal-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-medium text-slate-700 mb-1">{isVi ? 'Vật Liệu Thanh Cái' : 'Busbar Material'}</label>
                <input
                  type="text"
                  value={formData.tev_coe_header.busbar_material}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      tev_coe_header: { ...formData.tev_coe_header, busbar_material: e.target.value }
                    })
                  }
                  className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-teal-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-medium text-slate-700 mb-1">{isVi ? 'Kiểu Vỏ & Cấp Bảo Vệ IP' : 'Enclosure Type & IP'}</label>
                <input
                  type="text"
                  value={`${formData.tev_coe_header.enclosure_type || 'Aluminum'} (${formData.tev_coe_header.ip_rating || 'IP54'})`}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      tev_coe_header: { ...formData.tev_coe_header, enclosure_type: e.target.value }
                    })
                  }
                  className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-teal-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-medium text-slate-700 mb-1">{isVi ? 'Vị Trí / Đoạn Tuyến' : 'Location / Routing Bay'}</label>
                <input
                  type="text"
                  value={formData.tev_coe_header.location_bay || ''}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      tev_coe_header: { ...formData.tev_coe_header, location_bay: e.target.value }
                    })
                  }
                  className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-teal-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-medium text-slate-700 mb-1">{isVi ? 'Kỹ Sư Thử Nghiệm (FSE)' : 'Lead FSE Engineer'}</label>
                <input
                  type="text"
                  value={formData.tev_coe_header.fse_name}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      tev_coe_header: { ...formData.tev_coe_header, fse_name: e.target.value }
                    })
                  }
                  className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-teal-500 focus:outline-none"
                />
              </div>
            </div>
          </div>

          {/* Sub-phase Tabs for NETA MTS Stages */}
          <div className="flex border-b border-slate-200">
            <button
              type="button"
              onClick={() => setActivePhase('phase1')}
              className={`py-2 px-4 font-semibold text-xs transition border-b-2 ${
                activePhase === 'phase1'
                  ? 'border-teal-600 text-teal-700'
                  : 'border-transparent text-slate-500 hover:text-slate-800'
              }`}
            >
              {isVi ? 'Giai Đoạn 1: Thị Giác & Cơ Khí (7.4.A)' : 'Phase 1: Visual & Mechanical (7.4.A)'}
            </button>
            <button
              type="button"
              onClick={() => setActivePhase('phase2')}
              className={`py-2 px-4 font-semibold text-xs transition border-b-2 ${
                activePhase === 'phase2'
                  ? 'border-teal-600 text-teal-700'
                  : 'border-transparent text-slate-500 hover:text-slate-800'
              }`}
            >
              {isVi ? 'Giai Đoạn 2: Thử Nghiệm Điện (7.4.B)' : 'Phase 2: Electrical Tests (7.4.B)'}
            </button>
            <button
              type="button"
              onClick={() => setActivePhase('phase3')}
              className={`py-2 px-4 font-semibold text-xs transition border-b-2 ${
                activePhase === 'phase3'
                  ? 'border-teal-600 text-teal-700'
                  : 'border-transparent text-slate-500 hover:text-slate-800'
              }`}
            >
              {isVi ? 'Giai Đoạn 3: Hồng Ngoại & Phóng Điện Cục Bộ' : 'Phase 3: Thermography & PD'}
            </button>
          </div>

          {/* PHASE 1: VISUAL & MECHANICAL */}
          {activePhase === 'phase1' && (
            <div className="space-y-4 text-xs">
              <div className="border border-slate-200 rounded-xl overflow-hidden">
                <div className="bg-slate-100 px-4 py-2 font-bold text-slate-700 uppercase tracking-wider flex items-center gap-2">
                  <Wrench className="w-3.5 h-3.5 text-teal-600" />
                  {isVi ? 'Kiểm Tra Thị Giác, Vệ Sinh & Cơ Khí (ANSI/NETA MTS-2023 Mục 7.4.A)' : 'Visual & Mechanical Inspection (MTS-2023 Sec 7.4.A)'}
                </div>

                <div className="p-4 space-y-4">
                  <div>
                    <label className="block font-medium text-slate-700 mb-1">
                      {isVi ? 'A1-A2. Tình Trạng Cơ Lý Vỏ Kim Loại, Căn Chỉnh Tuyến & Giá Treo Đỡ' : 'A1-A2. Physical Condition, Alignment & Anchorage'}
                    </label>
                    <textarea
                      rows={2}
                      value={formData.visual_and_mechanical_inspection.physical_condition}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          visual_and_mechanical_inspection: {
                            ...formData.visual_and_mechanical_inspection,
                            physical_condition: e.target.value
                          }
                        })
                      }
                      className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-teal-500 focus:outline-none"
                    />
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                      <label className="block font-medium text-slate-700 mb-1">
                        {isVi ? 'A2. Tính Liên Tục Của Mạch Nối Đất Vỏ Kim Loại (Enclosure Grounding)' : 'A2. Enclosure Grounding & Bonding Continuity'}
                      </label>
                      <input
                        type="text"
                        value={formData.visual_and_mechanical_inspection.anchorage_and_grounding}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            visual_and_mechanical_inspection: {
                              ...formData.visual_and_mechanical_inspection,
                              anchorage_and_grounding: e.target.value
                            }
                          })
                        }
                        className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-teal-500 focus:outline-none"
                      />
                    </div>

                    <div>
                      <label className="block font-medium text-slate-700 mb-1">
                        {isVi ? 'A3. Độ Chặt Bu-lông Mối Nối Bus Joint (NETA Table 100.12 Torque)' : 'A3. Bolted Connection Bolt Torque (NETA Table 100.12)'}
                      </label>
                      <input
                        type="text"
                        value={formData.visual_and_mechanical_inspection.bolted_connection_torque_table_100_12}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            visual_and_mechanical_inspection: {
                              ...formData.visual_and_mechanical_inspection,
                              bolted_connection_torque_table_100_12: e.target.value
                            }
                          })
                        }
                        className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-teal-500 focus:outline-none"
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                      <label className="block font-medium text-slate-700 mb-1">
                        {isVi ? 'A6. Hệ Thống Sấy Kháng Ẩm & Điều Nhiệt (Space Heaters & Thermostat)' : 'A6. Space Heaters & Moisture Control Operation (7.4.A.6)'}
                      </label>
                      <input
                        type="text"
                        value={formData.visual_and_mechanical_inspection.space_heaters_operation}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            visual_and_mechanical_inspection: {
                              ...formData.visual_and_mechanical_inspection,
                              space_heaters_operation: e.target.value
                            }
                          })
                        }
                        className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-teal-500 focus:outline-none"
                      />
                    </div>

                    <div>
                      <label className="block font-medium text-slate-700 mb-1">
                        {isVi ? 'A10. Tấm Cách Điện Pha, Giá Đỡ Epoxy & Vách Ngăn Cháy (Fire Stops)' : 'A10. Phase Barriers, Insulators & Fire Stops (7.4.A.10)'}
                      </label>
                      <input
                        type="text"
                        value={formData.visual_and_mechanical_inspection.insulators_and_phase_barriers || ''}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            visual_and_mechanical_inspection: {
                              ...formData.visual_and_mechanical_inspection,
                              insulators_and_phase_barriers: e.target.value
                            }
                          })
                        }
                        className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-teal-500 focus:outline-none"
                      />
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* PHASE 2: ELECTRICAL TESTS */}
          {activePhase === 'phase2' && (
            <div className="space-y-4 text-xs">
              {/* 1. Bus Joint Microhm Resistance Test (7.4.D.1 & 7.4.D.4) */}
              <div className="border border-teal-200 bg-teal-50/20 rounded-xl p-4">
                <div className="flex items-center justify-between mb-3 border-b border-teal-200/60 pb-2">
                  <h4 className="text-xs font-bold uppercase tracking-wider text-teal-900 flex items-center gap-2">
                    <Activity className="w-4 h-4 text-teal-600" />
                    {isVi ? '1. Đo Điện Trở Mối Nối Thanh Cái (Bus Joint Resistance - Độ Lệch ≤ 50%)' : '1. Bus Joint Microhm Resistance Test (Deviation ≤ 50%)'}
                  </h4>
                  <span className="text-[11px] font-semibold text-teal-800 bg-teal-100 px-2 py-0.5 rounded">
                    {isVi ? 'Tiêu chuẩn NETA MTS 7.4.D.1 & 7.4.D.4' : 'NETA MTS 7.4.D.1 & 7.4.D.4'}
                  </span>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                  <div className="bg-white p-3 rounded-lg border border-slate-200">
                    <label className="text-slate-600 block mb-1">{isVi ? 'Trung bình Joint 01 - 04 (µΩ)' : 'Avg Joint 01-04 (µΩ)'}</label>
                    <input
                      type="number"
                      step={0.1}
                      value={formData.electrical_tests_7_4_b.bus_joint_microhm_resistance_test.joint_01_to_04_avg_uOhm}
                      onChange={(e) => {
                        const ref = Number(e.target.value);
                        const j5 = formData.electrical_tests_7_4_b.bus_joint_microhm_resistance_test.joint_05_phase_b_uOhm;
                        const dev = ref > 0 ? ((j5 - ref) / ref) * 100 : 0;
                        setFormData({
                          ...formData,
                          electrical_tests_7_4_b: {
                            ...formData.electrical_tests_7_4_b,
                            bus_joint_microhm_resistance_test: {
                              ...formData.electrical_tests_7_4_b.bus_joint_microhm_resistance_test,
                              joint_01_to_04_avg_uOhm: ref,
                              max_resistance_deviation_percent: Number(dev.toFixed(1))
                            }
                          }
                        });
                      }}
                      className="w-full px-2 py-1.5 border border-slate-300 rounded font-semibold text-slate-800"
                    />
                  </div>

                  <div className="bg-white p-3 rounded-lg border-2 border-teal-400">
                    <label className="text-teal-900 font-bold block mb-1">{isVi ? 'Joint 05 Pha B Giám Sát (µΩ)' : 'Joint 05 Phase B (µΩ)'}</label>
                    <input
                      type="number"
                      step={0.1}
                      value={formData.electrical_tests_7_4_b.bus_joint_microhm_resistance_test.joint_05_phase_b_uOhm}
                      onChange={(e) => {
                        const j5 = Number(e.target.value);
                        const ref = formData.electrical_tests_7_4_b.bus_joint_microhm_resistance_test.joint_01_to_04_avg_uOhm || 12.0;
                        const dev = ref > 0 ? ((j5 - ref) / ref) * 100 : 0;
                        setFormData({
                          ...formData,
                          electrical_tests_7_4_b: {
                            ...formData.electrical_tests_7_4_b,
                            bus_joint_microhm_resistance_test: {
                              ...formData.electrical_tests_7_4_b.bus_joint_microhm_resistance_test,
                              joint_05_phase_b_uOhm: j5,
                              max_resistance_deviation_percent: Number(dev.toFixed(1))
                            }
                          }
                        });
                      }}
                      className="w-full px-2 py-1.5 border border-teal-500 rounded font-bold text-teal-900 text-sm"
                    />
                  </div>

                  <div className="bg-white p-3 rounded-lg border border-slate-200">
                    <label className="text-slate-600 block mb-1">{isVi ? 'Độ Lệch Tối Đa (%) (Chuẩn ≤ 50%)' : 'Max Deviation (%) (Std ≤ 50%)'}</label>
                    <input
                      type="number"
                      step={0.1}
                      value={formData.electrical_tests_7_4_b.bus_joint_microhm_resistance_test.max_resistance_deviation_percent}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          electrical_tests_7_4_b: {
                            ...formData.electrical_tests_7_4_b,
                            bus_joint_microhm_resistance_test: {
                              ...formData.electrical_tests_7_4_b.bus_joint_microhm_resistance_test,
                              max_resistance_deviation_percent: Number(e.target.value)
                            }
                          }
                        })
                      }
                      className={`w-full px-2 py-1.5 border rounded font-bold text-sm ${
                        formData.electrical_tests_7_4_b.bus_joint_microhm_resistance_test.max_resistance_deviation_percent > 50
                          ? 'bg-rose-50 border-rose-400 text-rose-700'
                          : 'bg-emerald-50 border-emerald-400 text-emerald-700'
                      }`}
                    />
                  </div>

                  <div className="bg-white p-3 rounded-lg border border-slate-200">
                    <label className="text-slate-600 block mb-1">{isVi ? 'Đánh Giá / Trạng Thái Mối Nối' : 'Joint Status Evaluation'}</label>
                    <input
                      type="text"
                      value={formData.electrical_tests_7_4_b.bus_joint_microhm_resistance_test.status}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          electrical_tests_7_4_b: {
                            ...formData.electrical_tests_7_4_b,
                            bus_joint_microhm_resistance_test: {
                              ...formData.electrical_tests_7_4_b.bus_joint_microhm_resistance_test,
                              status: e.target.value
                            }
                          }
                        })
                      }
                      className="w-full px-2 py-1.5 border border-slate-300 rounded text-slate-800"
                    />
                  </div>
                </div>
              </div>

              {/* 2. Insulation Resistance Test (Table 100.1 & 7.4.D.2) */}
              <div className="border border-slate-200 rounded-xl p-4">
                <div className="flex items-center justify-between mb-3 border-b border-slate-200 pb-2">
                  <h4 className="text-xs font-bold uppercase tracking-wider text-slate-700 flex items-center gap-2">
                    <Gauge className="w-4 h-4 text-teal-600" />
                    {isVi ? '2. Thử Nghiệm Điện Trở Cách Điện Megohm (NETA Table 100.1 & 7.4.D.2)' : '2. Insulation Resistance Megohms Test (Table 100.1)'}
                  </h4>
                  <span className="text-[11px] text-slate-500 font-medium">
                    {isVi
                      ? formData.tev_coe_header.voltage_rating_kV >= 1.0
                        ? 'Hệ Trung Thế: Min ≥ 1,000 MΩ (thử 5000V DC)'
                        : 'Hệ Hạ Thế: Min ≥ 2.0 MΩ (thử 1000V DC)'
                      : formData.tev_coe_header.voltage_rating_kV >= 1.0
                      ? 'MV Busway: Min ≥ 1,000 MΩ at 5kV DC'
                      : 'LV Busway: Min ≥ 2.0 MΩ at 1kV DC'}
                  </span>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                  <div>
                    <label className="text-slate-600 block mb-1">{isVi ? 'Pha A Đối Đất (MΩ)' : 'Phase A to Ground (MΩ)'}</label>
                    <input
                      type="number"
                      value={formData.electrical_tests_7_4_b.insulation_resistance_test_table_100_1.phase_a_to_ground_Mohm}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          electrical_tests_7_4_b: {
                            ...formData.electrical_tests_7_4_b,
                            insulation_resistance_test_table_100_1: {
                              ...formData.electrical_tests_7_4_b.insulation_resistance_test_table_100_1,
                              phase_a_to_ground_Mohm: Number(e.target.value)
                            }
                          }
                        })
                      }
                      className="w-full px-2.5 py-1.5 border border-slate-300 rounded font-semibold text-slate-900"
                    />
                  </div>

                  <div>
                    <label className="text-slate-600 block mb-1">{isVi ? 'Pha B Đối Đất (MΩ)' : 'Phase B to Ground (MΩ)'}</label>
                    <input
                      type="number"
                      value={formData.electrical_tests_7_4_b.insulation_resistance_test_table_100_1.phase_b_to_ground_Mohm}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          electrical_tests_7_4_b: {
                            ...formData.electrical_tests_7_4_b,
                            insulation_resistance_test_table_100_1: {
                              ...formData.electrical_tests_7_4_b.insulation_resistance_test_table_100_1,
                              phase_b_to_ground_Mohm: Number(e.target.value)
                            }
                          }
                        })
                      }
                      className={`w-full px-2.5 py-1.5 border rounded font-bold ${
                        formData.electrical_tests_7_4_b.insulation_resistance_test_table_100_1.phase_b_to_ground_Mohm < (formData.tev_coe_header.voltage_rating_kV >= 1 ? 1000 : 2)
                          ? 'border-rose-400 bg-rose-50 text-rose-700'
                          : 'border-slate-300 text-slate-900'
                      }`}
                    />
                  </div>

                  <div>
                    <label className="text-slate-600 block mb-1">{isVi ? 'Pha C Đối Đất (MΩ)' : 'Phase C to Ground (MΩ)'}</label>
                    <input
                      type="number"
                      value={formData.electrical_tests_7_4_b.insulation_resistance_test_table_100_1.phase_c_to_ground_Mohm}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          electrical_tests_7_4_b: {
                            ...formData.electrical_tests_7_4_b,
                            insulation_resistance_test_table_100_1: {
                              ...formData.electrical_tests_7_4_b.insulation_resistance_test_table_100_1,
                              phase_c_to_ground_Mohm: Number(e.target.value)
                            }
                          }
                        })
                      }
                      className="w-full px-2.5 py-1.5 border border-slate-300 rounded font-semibold text-slate-900"
                    />
                  </div>

                  <div>
                    <label className="text-slate-600 block mb-1">{isVi ? 'Điện Áp Thử DC (V)' : 'DC Test Voltage (V)'}</label>
                    <input
                      type="number"
                      value={formData.electrical_tests_7_4_b.insulation_resistance_test_table_100_1.test_voltage_dc_V}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          electrical_tests_7_4_b: {
                            ...formData.electrical_tests_7_4_b,
                            insulation_resistance_test_table_100_1: {
                              ...formData.electrical_tests_7_4_b.insulation_resistance_test_table_100_1,
                              test_voltage_dc_V: Number(e.target.value)
                            }
                          }
                        })
                      }
                      className="w-full px-2.5 py-1.5 border border-slate-300 rounded font-semibold text-slate-900"
                    />
                  </div>
                </div>
              </div>

              {/* 3. Dielectric Withstand & Phasing Test */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="border border-slate-200 rounded-xl p-4">
                  <h4 className="font-bold text-slate-700 mb-2 flex items-center gap-2">
                    <Zap className="w-4 h-4 text-teal-600" />
                    {isVi ? '3. Thử Chịu Đựng Điện Áp Cao Hi-Pot (NETA Table 100.17)' : '3. Dielectric Withstand Voltage Test (Table 100.17)'}
                  </h4>
                  <div className="grid grid-cols-2 gap-3 mb-2">
                    <div>
                      <label className="text-slate-600 block mb-0.5">{isVi ? 'Điện áp thử (kV AC)' : 'Test Voltage (kV AC)'}</label>
                      <input
                        type="number"
                        step={0.1}
                        value={formData.electrical_tests_7_4_b.dielectric_withstand_test_table_100_17.test_voltage_kV_ac}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            electrical_tests_7_4_b: {
                              ...formData.electrical_tests_7_4_b,
                              dielectric_withstand_test_table_100_17: {
                                ...formData.electrical_tests_7_4_b.dielectric_withstand_test_table_100_17,
                                test_voltage_kV_ac: Number(e.target.value)
                              }
                            }
                          })
                        }
                        className="w-full px-2 py-1 border border-slate-300 rounded text-slate-900"
                      />
                    </div>
                    <div>
                      <label className="text-slate-600 block mb-0.5">{isVi ? 'Thời gian thử (Phút)' : 'Duration (Min)'}</label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_4_b.dielectric_withstand_test_table_100_17.duration_minutes}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            electrical_tests_7_4_b: {
                              ...formData.electrical_tests_7_4_b,
                              dielectric_withstand_test_table_100_17: {
                                ...formData.electrical_tests_7_4_b.dielectric_withstand_test_table_100_17,
                                duration_minutes: Number(e.target.value)
                              }
                            }
                          })
                        }
                        className="w-full px-2 py-1 border border-slate-300 rounded text-slate-900"
                      />
                    </div>
                  </div>
                  <div>
                    <label className="text-slate-600 block mb-0.5">{isVi ? 'Kết quả thử cao áp' : 'High Potential Result'}</label>
                    <input
                      type="text"
                      value={formData.electrical_tests_7_4_b.dielectric_withstand_test_table_100_17.phase_b_status}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          electrical_tests_7_4_b: {
                            ...formData.electrical_tests_7_4_b,
                            dielectric_withstand_test_table_100_17: {
                              ...formData.electrical_tests_7_4_b.dielectric_withstand_test_table_100_17,
                              phase_b_status: e.target.value
                            }
                          }
                        })
                      }
                      className="w-full px-2.5 py-1 border border-slate-300 rounded text-slate-900"
                    />
                  </div>
                </div>

                <div className="border border-slate-200 rounded-xl p-4">
                  <h4 className="font-bold text-slate-700 mb-2 flex items-center gap-2">
                    <Split className="w-4 h-4 text-teal-600" />
                    {isVi ? '4. Thử Thứ Tự Pha & Trùng Pha (NETA MTS 7.4.D.5)' : '4. Phasing Verification Test (7.4.D.5)'}
                  </h4>
                  <p className="text-slate-500 mb-2">
                    {isVi
                      ? 'Xác nhận thứ tự pha A-B-C tại điểm đầu và điểm cuối của tuyến Busway trước khi đóng điện hòa mạng.'
                      : 'Verify correct phase sequence A-B-C matching system single-line diagram.'}
                  </p>
                  <input
                    type="text"
                    value={formData.electrical_tests_7_4_b.phasing_verification_test}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        electrical_tests_7_4_b: {
                          ...formData.electrical_tests_7_4_b,
                          phasing_verification_test: e.target.value
                        }
                      })
                    }
                    className="w-full px-2.5 py-1.5 border border-slate-300 rounded text-slate-900 font-semibold"
                  />
                </div>
              </div>
            </div>
          )}

          {/* PHASE 3: THERMOGRAPHY & PD */}
          {activePhase === 'phase3' && (
            <div className="space-y-4 text-xs">
              {/* Online Partial Discharge */}
              <div className="border border-purple-200 bg-purple-50/20 rounded-xl p-4">
                <div className="flex items-center justify-between mb-3 border-b border-purple-200/60 pb-2">
                  <h4 className="text-xs font-bold uppercase tracking-wider text-purple-900 flex items-center gap-2">
                    <Radio className="w-4 h-4 text-purple-600" />
                    {isVi ? 'Khảo Sát Phóng Điện Cục Bộ Online (NETA Table 100.23)' : 'Online Partial Discharge Survey (Table 100.23)'}
                  </h4>
                  <span className="text-[11px] font-semibold text-purple-800 bg-purple-100 px-2 py-0.5 rounded">
                    {isVi ? 'TEV < 20 dB (An toàn), Acoustic < 3 dB' : 'TEV < 20 dB (Safe), Acoustic < 3 dB'}
                  </span>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div>
                    <label className="text-slate-600 block mb-1">{isVi ? 'Mức TEV Cực Đại (dB)' : 'Max TEV Level (dB)'}</label>
                    <input
                      type="number"
                      value={formData.electrical_tests_7_4_b.online_pd_survey_tables_100_23.tev_max_dB}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          electrical_tests_7_4_b: {
                            ...formData.electrical_tests_7_4_b,
                            online_pd_survey_tables_100_23: {
                              ...formData.electrical_tests_7_4_b.online_pd_survey_tables_100_23,
                              tev_max_dB: Number(e.target.value)
                            }
                          }
                        })
                      }
                      className={`w-full px-2.5 py-1.5 border rounded font-bold text-sm ${
                        formData.electrical_tests_7_4_b.online_pd_survey_tables_100_23.tev_max_dB >= 30
                          ? 'border-rose-400 bg-rose-50 text-rose-700'
                          : 'border-slate-300 text-slate-900'
                      }`}
                    />
                  </div>

                  <div>
                    <label className="text-slate-600 block mb-1">{isVi ? 'Mức Siêu Âm Acoustic (dB)' : 'Max Acoustic Level (dB)'}</label>
                    <input
                      type="number"
                      step={0.1}
                      value={formData.electrical_tests_7_4_b.online_pd_survey_tables_100_23.acoustic_max_dB}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          electrical_tests_7_4_b: {
                            ...formData.electrical_tests_7_4_b,
                            online_pd_survey_tables_100_23: {
                              ...formData.electrical_tests_7_4_b.online_pd_survey_tables_100_23,
                              acoustic_max_dB: Number(e.target.value)
                            }
                          }
                        })
                      }
                      className={`w-full px-2.5 py-1.5 border rounded font-bold text-sm ${
                        formData.electrical_tests_7_4_b.online_pd_survey_tables_100_23.acoustic_max_dB > 6.0
                          ? 'border-rose-400 bg-rose-50 text-rose-700'
                          : 'border-slate-300 text-slate-900'
                      }`}
                    />
                  </div>

                  <div>
                    <label className="text-slate-600 block mb-1">{isVi ? 'Vị Trí Khoang Mối Nối Khảo Sát' : 'Joint Location Surveyed'}</label>
                    <input
                      type="text"
                      value={formData.electrical_tests_7_4_b.online_pd_survey_tables_100_23.location}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          electrical_tests_7_4_b: {
                            ...formData.electrical_tests_7_4_b,
                            online_pd_survey_tables_100_23: {
                              ...formData.electrical_tests_7_4_b.online_pd_survey_tables_100_23,
                              location: e.target.value
                            }
                          }
                        })
                      }
                      className="w-full px-2.5 py-1.5 border border-slate-300 rounded text-slate-900"
                    />
                  </div>
                </div>
              </div>

              {/* Thermography */}
              <div className="border border-slate-200 rounded-xl p-4">
                <h4 className="font-bold text-slate-700 mb-2 flex items-center gap-2">
                  <Thermometer className="w-4 h-4 text-rose-600" />
                  {isVi ? 'Khảo Sát Hồng Ngoại Nhiệt Độ Tiếp Xúc (NETA MTS Table 100.18)' : 'Thermographic Survey & Thermal Inspection (NETA Table 100.18)'}
                </h4>
                <p className="text-slate-500 mb-2">
                  {isVi
                    ? 'Chụp ảnh nhiệt hồng ngoại các mối nối bu-lông Bus Joint khi mang tải ≥ 40% định mức. Sai lệch nhiệt độ ΔT > 4°C yêu cầu bảo dưỡng, ΔT > 15°C yêu cầu cô lập tức thì.'
                    : 'Thermographic survey across bolted bus joints under minimum 40% rated load per NETA Table 100.18.'}
                </p>
                <textarea
                  rows={2}
                  value={formData.post_test_survey.thermographic_survey_table_100_18}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      post_test_survey: {
                        ...formData.post_test_survey,
                        thermographic_survey_table_100_18: e.target.value
                      }
                    })
                  }
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-teal-500 focus:outline-none"
                />
              </div>
            </div>
          )}
        </div>
      )}

      {/* TAB CONTENT 2: AI & NETA DIAGNOSTIC EVALUATION */}
      {activeTab === 'evaluation' && (
        <div className="p-6 space-y-6">
          {(() => {
            const currentEval = evaluationResult || performDeterministicMtsBuswayAnalysis(formData);
            const isCritical = currentEval.overall_status === 'CRITICAL';
            const isCaution = currentEval.overall_status === 'CAUTION';

            return (
              <>
                {/* Overall Severity Banner */}
                <div
                  className={`p-5 rounded-2xl border flex flex-col md:flex-row md:items-center justify-between gap-4 ${
                    isCritical
                      ? 'bg-rose-50 border-rose-300 text-rose-950'
                      : isCaution
                      ? 'bg-amber-50 border-amber-300 text-amber-950'
                      : 'bg-emerald-50 border-emerald-300 text-emerald-950'
                  }`}
                >
                  <div className="flex items-start gap-3.5">
                    {isCritical ? (
                      <AlertOctagon className="w-8 h-8 text-rose-600 shrink-0 mt-0.5" />
                    ) : isCaution ? (
                      <AlertTriangle className="w-8 h-8 text-amber-600 shrink-0 mt-0.5" />
                    ) : (
                      <ShieldCheck className="w-8 h-8 text-emerald-600 shrink-0 mt-0.5" />
                    )}
                    <div>
                      <div className="flex items-center gap-2">
                        <span
                          className={`px-2.5 py-0.5 rounded-full text-xs font-bold uppercase tracking-wider ${
                            isCritical
                              ? 'bg-rose-600 text-white'
                              : isCaution
                              ? 'bg-amber-600 text-white'
                              : 'bg-emerald-600 text-white'
                          }`}
                        >
                          {currentEval.overall_status}
                        </span>
                        <h3 className="text-base sm:text-lg font-bold">
                          {isVi ? currentEval.overall_title_vi : currentEval.overall_title_en}
                        </h3>
                      </div>
                      <p className="text-xs sm:text-sm mt-1 text-slate-700">
                        {isVi ? currentEval.summary_vi : currentEval.summary_en}
                      </p>
                    </div>
                  </div>

                  <div className="flex items-center gap-2 self-start md:self-center shrink-0">
                    <button
                      type="button"
                      onClick={handleExportPdf}
                      className="px-3 py-1.5 rounded-lg bg-white shadow-sm border border-slate-300 hover:bg-slate-50 text-xs font-semibold text-slate-800 flex items-center gap-1.5 transition"
                    >
                      <Download className="w-3.5 h-3.5 text-teal-600" />
                      <span>{isVi ? 'Xuất PDF Báo Cáo' : 'Export PDF'}</span>
                    </button>
                  </div>
                </div>

                {/* Gemini AI Detailed Markdown Output (if available) */}
                {currentEval.ai_generated_markdown && (
                  <div className="bg-gradient-to-br from-teal-50/50 via-white to-emerald-50/30 rounded-2xl border border-teal-200 p-5 shadow-sm">
                    <div className="flex items-center gap-2 mb-3 pb-2 border-b border-teal-200/60">
                      <Sparkles className="w-4 h-4 text-teal-600" />
                      <h4 className="text-xs font-bold uppercase tracking-wider text-teal-900">
                        {isVi ? 'Phân Tích Chuyên Sâu Tự Động từ AI Studio Engine (Gemini 2.5 Flash)' : 'Deep Analytical Reasoning (Gemini 2.5 Flash)'}
                      </h4>
                    </div>
                    <div className="prose prose-sm max-w-none text-xs text-slate-700 whitespace-pre-line leading-relaxed">
                      {currentEval.ai_generated_markdown}
                    </div>
                  </div>
                )}

                {/* Detailed Compliance Evaluation Table */}
                <div className="border border-slate-200 rounded-xl overflow-hidden shadow-sm">
                  <div className="bg-slate-900 px-4 py-3 text-white flex items-center justify-between">
                    <span className="text-xs font-bold uppercase tracking-wider flex items-center gap-2">
                      <FileCheck2 className="w-4 h-4 text-teal-400" />
                      {isVi
                        ? 'Bảng Đối Soát Thông Số Thử Nghiệm Busway Theo ANSI/NETA MTS-2023 Mục 7.4'
                        : 'Busway Maintenance Test Parameter Verification Table (NETA MTS-2023 Sec 7.4)'}
                    </span>
                    <span className="text-[11px] text-slate-400">
                      {currentEval.evaluation_table.length} {isVi ? 'Hạng mục kiểm tra' : 'Test items'}
                    </span>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-slate-100 text-slate-700 font-semibold border-b border-slate-200">
                          <th className="py-2.5 px-3.5">{isVi ? 'Hạng Mục Kiểm Tra & Thử Nghiệm' : 'Inspection & Test Item'}</th>
                          <th className="py-2.5 px-3">{isVi ? 'Tiêu Chuẩn Tham Chiếu' : 'Standard Ref'}</th>
                          <th className="py-2.5 px-3.5">{isVi ? 'Giá Trị Đo Thực Tế' : 'Measured / Observed Value'}</th>
                          <th className="py-2.5 px-3 text-center">{isVi ? 'Kết Luận' : 'Status'}</th>
                          <th className="py-2.5 px-3.5">{isVi ? 'Tiêu Chí & Đánh Giá Kỹ Thuật' : 'Technical Evaluation'}</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {currentEval.evaluation_table.map((row, idx) => {
                          const isRowCrit = row.status === 'CRITICAL' || row.status === 'FAIL';
                          const isRowCaut = row.status === 'CAUTION';

                          return (
                            <tr key={idx} className={isRowCrit ? 'bg-rose-50/40' : isRowCaut ? 'bg-amber-50/40' : 'hover:bg-slate-50/80'}>
                              <td className="py-2.5 px-3.5 font-semibold text-slate-800">
                                {isVi ? row.item_vn : row.item}
                              </td>
                              <td className="py-2.5 px-3 text-slate-500 font-mono text-[11px]">
                                {row.standard_ref}
                              </td>
                              <td className="py-2.5 px-3.5 font-medium text-slate-900">
                                {row.measured_value}
                              </td>
                              <td className="py-2.5 px-3 text-center">
                                <span
                                  className={`inline-flex px-2 py-0.5 rounded text-[11px] font-bold ${
                                    isRowCrit
                                      ? 'bg-rose-100 text-rose-700 border border-rose-300'
                                      : isRowCaut
                                      ? 'bg-amber-100 text-amber-800 border border-amber-300'
                                      : 'bg-emerald-100 text-emerald-700 border border-emerald-300'
                                  }`}
                                >
                                  {row.status}
                                </span>
                              </td>
                              <td className="py-2.5 px-3.5 text-slate-600 text-[11px] leading-relaxed">
                                {isVi ? row.notes_vn : row.notes}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* Technical Analysis & Recommendations Grid */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  {/* Technical Analysis */}
                  <div className="bg-slate-50 rounded-xl p-4 border border-slate-200">
                    <h4 className="text-xs font-bold uppercase tracking-wider text-slate-700 mb-3 flex items-center gap-2">
                      <Activity className="w-4 h-4 text-teal-600" />
                      {isVi ? 'Phân Tích Kỹ Thuật & Nguy Cơ Quá Nhiệt Nổ Mối Nối' : 'Technical Analysis & Hazard Breakdown'}
                    </h4>
                    <ul className="space-y-2 text-xs text-slate-700">
                      {(isVi ? currentEval.technical_analysis_vi : currentEval.technical_analysis_en).map((item, idx) => (
                        <li key={idx} className="flex items-start gap-2">
                          <span className="w-1.5 h-1.5 rounded-full bg-teal-600 mt-1.5 shrink-0" />
                          <span className="leading-relaxed">{item}</span>
                        </li>
                      ))}
                    </ul>
                  </div>

                  {/* Recommendations */}
                  <div className="bg-teal-50/40 rounded-xl p-4 border border-teal-200">
                    <h4 className="text-xs font-bold uppercase tracking-wider text-teal-900 mb-3 flex items-center gap-2">
                      <Wrench className="w-4 h-4 text-teal-600" />
                      {isVi ? 'Khuyến Nghị Hành Động Dành Cho FSE Hiện Trường' : 'Actionable Recommendations for FSE'}
                    </h4>
                    <ul className="space-y-2 text-xs text-slate-800">
                      {(isVi ? currentEval.recommendations_vi : currentEval.recommendations_en).map((rec, idx) => (
                        <li key={idx} className="flex items-start gap-2">
                          <span className="font-bold text-teal-700 shrink-0">{idx + 1}.</span>
                          <span className="leading-relaxed">{rec}</span>
                        </li>
                      ))}
                    </ul>
                  </div>
                </div>
              </>
            );
          })()}
        </div>
      )}

      {/* TAB CONTENT 3: SYSTEM PROMPT AI STUDIO */}
      {activeTab === 'prompt' && (
        <div className="p-6 space-y-4">
          <div className="flex items-center justify-between">
            <div>
              <h3 className="text-sm font-bold text-slate-800">
                {isVi ? 'System Instruction Chuẩn Hóa Cho Google AI Studio' : 'Standardized System Instruction Prompt'}
              </h3>
              <p className="text-xs text-slate-500">
                {isVi
                  ? 'Sao chép đoạn System Instruction này để nạp vào AI Studio hoặc Gemini API backend phục vụ chẩn đoán tự động.'
                  : 'Copy this System Instruction prompt for Google AI Studio or backend Gemini model integration.'}
              </p>
            </div>
            <button
              type="button"
              onClick={handleCopyPrompt}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-teal-600 hover:bg-teal-700 text-white text-xs font-medium shadow transition"
            >
              {copySuccess ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
              <span>{copySuccess ? (isVi ? 'Đã Sao Chép!' : 'Copied!') : (isVi ? 'Sao Chép Prompt' : 'Copy Prompt')}</span>
            </button>
          </div>

          <pre className="p-4 bg-slate-900 text-teal-200 text-xs font-mono rounded-xl overflow-x-auto whitespace-pre-wrap max-h-[500px] border border-slate-800">
            {MTS_BUSWAY_SYSTEM_INSTRUCTION}
          </pre>
        </div>
      )}

      {/* TAB CONTENT 4: JSON PAYLOAD */}
      {activeTab === 'json' && (
        <div className="p-6 space-y-4">
          <div className="flex items-center justify-between">
            <div>
              <h3 className="text-sm font-bold text-slate-800">
                {isVi ? 'Mẫu Dữ Liệu JSON Đầu Vào (Input Payload)' : 'Input JSON Payload'}
              </h3>
              <p className="text-xs text-slate-500">
                {isVi
                  ? 'Cấu trúc dữ liệu JSON đầy đủ phục vụ gọi API /api/field-service/neta-mts-busway-analyze.'
                  : 'Complete JSON payload structure for API integration.'}
              </p>
            </div>
            <button
              type="button"
              onClick={handleCopyJson}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-white text-xs font-medium shadow transition"
            >
              {jsonCopySuccess ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
              <span>{jsonCopySuccess ? (isVi ? 'Đã Sao Chép!' : 'Copied!') : (isVi ? 'Sao Chép JSON' : 'Copy JSON')}</span>
            </button>
          </div>

          <pre className="p-4 bg-slate-900 text-emerald-300 text-xs font-mono rounded-xl overflow-x-auto whitespace-pre-wrap max-h-[500px] border border-slate-800">
            {JSON.stringify(formData, null, 2)}
          </pre>
        </div>
      )}
    </div>
  );
};
