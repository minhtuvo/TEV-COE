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
  Radio
} from 'lucide-react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  NetaMtsCcvtInputPayload,
  NetaMtsCcvtEvaluationResult,
  performDeterministicMtsCcvtAnalysis,
  MTS_CCVT_SYSTEM_INSTRUCTION,
  SAMPLE_MTS_CCVT_DEFECT_PAYLOAD,
  SAMPLE_MTS_CCVT_PASS_PAYLOAD
} from '../../server/netaMtsCcvtAnalyzer';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
  initialLanguage?: 'vi' | 'en';
}

export const SAMPLE_MTS_CCVT_CAUTION_PAYLOAD: NetaMtsCcvtInputPayload = {
  tev_coe_header: {
    module: "CCVT-MAINTENANCE-MTS-7.10.3",
    project_name: "Tram 110kV KCN Yen Phong - TEV Substation",
    equipment_tag: "110KV-OUT-CCVT-SET (Bay 171 Line Exit)",
    voltage_rating_kV: 110,
    ccvt_ratio_rating: "110000/sqrt(3) : 110/sqrt(3) V",
    accuracy_class_rating: "0.5 3P 150VA (Line Relaying & Synchrocheck)",
    c1_nameplate_capacitance_pF: 6500,
    c2_nameplate_capacitance_pF: 95000,
    fse_name: "Hoang Van Duc",
    test_date: "2026-10-18",
    station_name: "110kV Substation Yen Phong",
    bay_location: "Line 171 Exit Structure"
  },
  visual_and_mechanical_inspection: {
    physical_condition: "Bề mặt sứ cách điện bám bụi nhẹ, đã lau chùi sạch bằng cồn công nghiệp trước khi thử; mức dầu IVT bình thường",
    grounding_switch_operation: "Pass (Dao tiếp địa vận hành cơ khí đóng/cắt nhẹ nhàng, khóa liên động an toàn)",
    carrier_drain_coil_status: "Pass (Cuộn cản xả bẫy sóng PLC carrier drain coil thông mạch tốt, điện trở đo đạt 0.45 ohm)",
    bolted_connection_torque_table_100_12: "Pass (Bu lông siết lực đạt tiêu chuẩn NETA Table 100.12)",
    secondary_grounding_point_ieee_c57_13_3: {
      ground_points_detected: 2,
      ground_location: "Tiếp địa cả ở hộp đầu cực CCVT và tại thanh cái đất tủ Rơ-le bảo vệ (Loop đất kép)",
      status: "Fail (Phát hiện 2 điểm tiếp địa mạch thứ cấp vi phạm IEEE C57.13.3 - Nguy cơ loop đất gây sai số đo)"
    },
    fluid_level_check: "Pass (Mắt dầu trong suốt, mức dầu nằm ở vạch 25°C)"
  },
  electrical_tests_7_10_3_b: {
    insulation_resistance_megohms_table_100_5: {
      primary_ivt_to_ground_5000v_dc_Mohm: 68000,
      secondary_to_ground_1000v_dc_Mohm: 1850,
      status: "Pass (Đạt tiêu chuẩn NETA Table 100.5 > 5000 MΩ và > 100 MΩ)"
    },
    capacitance_measurements_ieee_c93_1: {
      phase_a_c1_pF: 6515,
      phase_a_c2_pF: 95100,
      phase_b_c1_pF: 6490,
      phase_b_c2_pF: 94850,
      phase_c_c1_pF: 6520,
      phase_c_c2_pF: 95200,
      c1_deviation_phase_b_percent: -0.15,
      status: "Pass (Độ lệch điện dung C1 nhỏ hơn 0.5%, nằm trong dung sai ±5% IEEE C93.1)"
    },
    power_factor_tan_delta_test: {
      phase_a_c1_pf_percent: 0.32,
      phase_b_c1_pf_percent: 0.34,
      phase_c_c1_pf_percent: 0.30,
      status: "Pass (Hệ số tổn hao cách điện Tan delta <= 0.34% < 0.5% theo NETA MTS 7.10.3.D.8)"
    },
    polarity_test_ieee_c93_1: "Pass (Cực tính H1-X1 chuẩn trừ thuận)",
    turns_ratio_verification_test_ieee_c93_1: {
      applied_primary_voltage_V: 63508,
      measured_secondary_voltage_phase_a_V: 63.48,
      measured_secondary_voltage_phase_b_V: 63.45,
      measured_secondary_voltage_phase_c_V: 63.50,
      ratio_error_phase_b_percent: -0.09,
      status: "Pass (Sai số tỉ số biến < 0.1% đạt cấp chính xác 0.5 cho bảo vệ)"
    },
    secondary_burden_measurement: {
      measured_burden_phase_b_VA: 48.0,
      ccvt_rated_burden_VA: 150,
      status: "Pass (Phụ tải thứ cấp 48 VA nằm an toàn trong dải định mức 150 VA)"
    },
    dielectric_withstand_test_table_100_9: {
      test_voltage_kV_ac: 28,
      duration_minutes: 1,
      status: "Pass (Thử nghiệm điện áp tăng cao 28kV AC 1 phút đạt yêu cầu)"
    }
  },
  post_test_survey: {
    thermographic_survey_table_100_18: "Pass (Nhiệt độ mối nối tiếp xúc đồng đều, ΔT < 1.2°C per Table 100.18)"
  }
};

export const NetaMtsCcvtChecklist: React.FC<Props> = ({
  onSaveReport,
  onNavigateToReports,
  initialLanguage = 'vi'
}) => {
  const [lang, setLang] = useState<'vi' | 'en'>(initialLanguage);
  const [activeTab, setActiveTab] = useState<'checklist' | 'evaluation' | 'prompt' | 'json'>('checklist');
  const [activePhase, setActivePhase] = useState<'phase1' | 'phase2' | 'phase3'>('phase1');
  const [isEvaluating, setIsEvaluating] = useState(false);
  const [evaluationResult, setEvaluationResult] = useState<NetaMtsCcvtEvaluationResult | null>(null);
  const [copySuccess, setCopySuccess] = useState(false);
  const [jsonCopySuccess, setJsonCopySuccess] = useState(false);
  const [saveSuccess, setSaveSuccess] = useState(false);

  // Form State
  const [formData, setFormData] = useState<NetaMtsCcvtInputPayload>(SAMPLE_MTS_CCVT_DEFECT_PAYLOAD);

  const isVi = lang === 'vi';

  // Handler to load presets
  const handleLoadPreset = (scenario: 'defect' | 'pass' | 'caution') => {
    let payload = SAMPLE_MTS_CCVT_DEFECT_PAYLOAD;
    if (scenario === 'pass') payload = SAMPLE_MTS_CCVT_PASS_PAYLOAD;
    if (scenario === 'caution') payload = SAMPLE_MTS_CCVT_CAUTION_PAYLOAD;

    setFormData(JSON.parse(JSON.stringify(payload)));
    setEvaluationResult(null);
  };

  // Run AI / Deterministic Evaluation
  const handleRunEvaluation = async () => {
    setIsEvaluating(true);
    try {
      const response = await fetch('/api/field-service/neta-mts-ccvt-analyze', {
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
      const localEval = performDeterministicMtsCcvtAnalysis(formData);
      setEvaluationResult(localEval);
      setActiveTab('evaluation');
    } catch (err) {
      console.warn('API error, using local deterministic engine:', err);
      const localEval = performDeterministicMtsCcvtAnalysis(formData);
      setEvaluationResult(localEval);
      setActiveTab('evaluation');
    } finally {
      setIsEvaluating(false);
    }
  };

  // Copy AI Studio System Instruction Prompt
  const handleCopyPrompt = () => {
    navigator.clipboard.writeText(MTS_CCVT_SYSTEM_INSTRUCTION);
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
      const computed = performDeterministicMtsCcvtAnalysis(formData);
      setEvaluationResult(computed);
    }
    const reportData = {
      id: `NETA-MTS-CCVT-${Date.now()}`,
      moduleType: 'MTS-CCVT-7.10.3',
      title: `${formData.tev_coe_header.equipment_tag} - NETA MTS-2023 Sec 7.10.3`,
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
    const currentEval = evaluationResult || performDeterministicMtsCcvtAnalysis(formData);

    // Header
    doc.setFillColor(15, 23, 42);
    doc.rect(0, 0, 210, 26, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(13);
    doc.text('TEV COE - ELECTRICAL TESTING & COMMISSIONING PLATFORM', 14, 11);
    doc.setFontSize(9.5);
    doc.setTextColor(148, 163, 184);
    doc.text('ANSI/NETA MTS-2023 SECTION 7.10.3 - COUPLING-CAPACITOR VOLTAGE TRANSFORMERS (CCVT)', 14, 18);

    // Info Box
    doc.setFillColor(248, 250, 252);
    doc.rect(14, 30, 182, 32, 'F');
    doc.setDrawColor(203, 213, 225);
    doc.rect(14, 30, 182, 32, 'S');

    doc.setTextColor(15, 23, 42);
    doc.setFontSize(8.5);
    doc.text(`Project: ${formData.tev_coe_header.project_name}`, 18, 38);
    doc.text(`Equipment Tag: ${formData.tev_coe_header.equipment_tag}`, 18, 44);
    doc.text(`Ratio: ${formData.tev_coe_header.ccvt_ratio_rating} | Class: ${formData.tev_coe_header.accuracy_class_rating}`, 18, 50);
    doc.text(`Capacitance: C1=${formData.tev_coe_header.c1_nameplate_capacitance_pF}pF, C2=${formData.tev_coe_header.c2_nameplate_capacitance_pF}pF | Voltage: ${formData.tev_coe_header.voltage_rating_kV} kV`, 18, 56);

    doc.text(`FSE Engineer: ${formData.tev_coe_header.fse_name}`, 115, 38);
    doc.text(`Test Date: ${formData.tev_coe_header.test_date}`, 115, 44);
    doc.text(`Location: ${formData.tev_coe_header.bay_location || 'Line Bay Feeder'}`, 115, 50);
    doc.text(`Standard: ANSI/NETA MTS-2023 (Sec 7.10.3) & IEEE C93.1`, 115, 56);

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
      head: [['Inspection & Test Item', 'Standard Ref', 'Measured / Observed Value', 'Status', 'Technical Criteria']],
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

    // Technical Summary & Recommendations
    doc.setTextColor(15, 23, 42);
    doc.setFontSize(9);
    doc.setFont(undefined, 'bold');
    doc.text(isVi ? 'KHUYẾN NGHỊ KỸ THUẬT HÀNH ĐỘNG DÀNH CHO KỸ SƯ HIỆN TRƯỜNG (FSE):' : 'ACTIONABLE TECHNICAL RECOMMENDATIONS FOR FIELD ENGINEERS (FSE):', 14, finalY);

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
    doc.text('Station Manager / Commissioning Authority Signature: __________________________', 85, sigY + 6);
    doc.text(`Generated: ${new Date().toLocaleDateString()} | TEV COE Analytical Engine`, 14, sigY + 11);

    doc.save(`NETA_MTS_CCVT_${formData.tev_coe_header.equipment_tag.replace(/[^a-zA-Z0-9_-]/g, '_')}.pdf`);
  };

  // Helper for quick stats
  const stats = useMemo(() => {
    const computed = evaluationResult || performDeterministicMtsCcvtAnalysis(formData);
    const total = computed.evaluation_table.length;
    const passCount = computed.evaluation_table.filter(t => t.status === 'PASS').length;
    const failCount = computed.evaluation_table.filter(t => t.status === 'FAIL' || t.status === 'CRITICAL').length;
    const cautionCount = computed.evaluation_table.filter(t => t.status === 'CAUTION').length;
    return { total, passCount, failCount, cautionCount, status: computed.overall_status };
  }, [formData, evaluationResult]);

  return (
    <div className="bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden">
      {/* Top Header Banner */}
      <div className="bg-gradient-to-r from-slate-900 via-purple-950 to-slate-900 px-6 py-5 text-white border-b border-purple-800/40">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="space-y-1">
            <div className="flex items-center gap-2">
              <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-purple-500/20 text-purple-300 border border-purple-400/30">
                <Radio className="w-3.5 h-3.5 text-purple-400" />
                ANSI/NETA MTS-2023 Sec 7.10.3
              </span>
              <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-medium bg-emerald-500/10 text-emerald-300 border border-emerald-500/20">
                <CheckCircle2 className="w-3 h-3 text-emerald-400" />
                IEEE C93.1 & IEEE C57.13.3
              </span>
            </div>
            <h2 className="text-xl md:text-2xl font-bold tracking-tight text-white flex items-center gap-2">
              {isVi
                ? 'Bảo Trì & Thử Nghiệm Biến Điện Áp Kiểu Tụ (CCVT)'
                : 'Maintenance & Testing of Coupling-Capacitor Voltage Transformers (CCVT)'}
            </h2>
            <p className="text-xs md:text-sm text-slate-300 max-w-3xl">
              {isVi
                ? 'Đo điện dung C1/C2 (dung sai ≤ ±5%), tổn hao cách điện Tan delta (≤ 0.5%), tỉ số biến, dao tiếp địa và điểm nối đất duy nhất theo NETA MTS-2023 Mục 7.10.3.'
                : 'Capacitance C1/C2 measurement (tolerance ≤ ±5%), dissipation factor Tan Delta (≤ 0.5%), ratio verification, grounding switch, and single ground point per NETA MTS-2023 Sec 7.10.3.'}
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
                  lang === 'vi' ? 'bg-purple-600 text-white shadow-sm' : 'text-slate-400 hover:text-white'
                }`}
              >
                VN
              </button>
              <button
                type="button"
                onClick={() => setLang('en')}
                className={`px-2.5 py-1 text-xs font-medium rounded-md transition-all ${
                  lang === 'en' ? 'bg-purple-600 text-white shadow-sm' : 'text-slate-400 hover:text-white'
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
              title="Xuất Báo Cáo PDF Thử Nghiệm CCVT Chuẩn ANSI/NETA"
            >
              <Download className="w-3.5 h-3.5 text-purple-400" />
              <span>{isVi ? 'Xuất PDF' : 'PDF Report'}</span>
            </button>

            {/* Save Report */}
            <button
              type="button"
              onClick={handleSaveReport}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium transition shadow-sm ${
                saveSuccess
                  ? 'bg-emerald-600 text-white border border-emerald-500'
                  : 'bg-purple-700 hover:bg-purple-600 text-white border border-purple-500'
              }`}
            >
              {saveSuccess ? <Check className="w-3.5 h-3.5" /> : <Save className="w-3.5 h-3.5" />}
              <span>{saveSuccess ? (isVi ? 'Đã Lưu!' : 'Saved!') : (isVi ? 'Lưu Báo Cáo' : 'Save Report')}</span>
            </button>
          </div>
        </div>

        {/* Quick Scenario Preset Selector Bar */}
        <div className="mt-4 pt-3 border-t border-purple-800/40 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex items-center gap-2">
            <span className="text-slate-300 font-medium">{isVi ? 'Kịch bản mẫu:' : 'Simulation Presets:'}</span>
            <div className="inline-flex rounded-lg bg-slate-800/90 p-0.5 border border-slate-700">
              <button
                type="button"
                onClick={() => handleLoadPreset('defect')}
                className="px-2.5 py-1 rounded-md text-rose-300 hover:bg-rose-500/20 font-medium transition flex items-center gap-1"
              >
                <AlertOctagon className="w-3 h-3 text-rose-400" />
                <span>{isVi ? 'Trạm 230kV (Sự cố Chập Tụ Pha B)' : '230kV Defect (C1 Shorted)'}</span>
              </button>
              <button
                type="button"
                onClick={() => handleLoadPreset('pass')}
                className="px-2.5 py-1 rounded-md text-emerald-300 hover:bg-emerald-500/20 font-medium transition flex items-center gap-1"
              >
                <CheckCircle2 className="w-3 h-3 text-emerald-400" />
                <span>{isVi ? 'Trạm 500kV (Đạt Chuẩn 100%)' : '500kV Standard Pass'}</span>
              </button>
              <button
                type="button"
                onClick={() => handleLoadPreset('caution')}
                className="px-2.5 py-1 rounded-md text-amber-300 hover:bg-amber-500/20 font-medium transition flex items-center gap-1"
              >
                <AlertTriangle className="w-3 h-3 text-amber-400" />
                <span>{isVi ? 'Trạm 110kV (Lặp Đất Mạch Thứ Cấp)' : '110kV Dual Ground Loop'}</span>
              </button>
            </div>
          </div>

          {/* Quick Header Summary Pill */}
          <div className="flex items-center gap-2">
            <span className="text-slate-400">
              {formData.tev_coe_header.equipment_tag} ({formData.tev_coe_header.voltage_rating_kV}kV)
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
                ? 'border-purple-600 text-purple-700 bg-white shadow-sm'
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
                ? 'border-purple-600 text-purple-700 bg-white shadow-sm'
                : 'border-transparent text-slate-600 hover:text-slate-900 hover:border-slate-300'
            }`}
          >
            <Sparkles className="w-4 h-4 text-purple-600" />
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
                ? 'border-purple-600 text-purple-700 bg-white shadow-sm'
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
                ? 'border-purple-600 text-purple-700 bg-white shadow-sm'
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
          className="inline-flex items-center gap-2 px-4 py-2 rounded-xl bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-700 hover:to-indigo-700 text-white text-xs sm:text-sm font-semibold shadow-md transition transform active:scale-95 disabled:opacity-50"
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
              <Building className="w-4 h-4 text-purple-600" />
              {isVi ? 'Thông Tin Hồ Sơ & Nhãn Mác Biến Điện Áp Tụ CCVT' : 'CCVT Nameplate & Substation Parameters'}
            </h3>
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 text-xs">
              <div>
                <label className="block font-medium text-slate-700 mb-1">{isVi ? 'Tên Dự Án / Trạm' : 'Project / Substation'}</label>
                <input
                  type="text"
                  value={formData.tev_coe_header.project_name}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      tev_coe_header: { ...formData.tev_coe_header, project_name: e.target.value }
                    })
                  }
                  className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-purple-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-medium text-slate-700 mb-1">{isVi ? 'Mã Ký Hiệu Thiết Bị (Tag)' : 'Equipment Tag'}</label>
                <input
                  type="text"
                  value={formData.tev_coe_header.equipment_tag}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      tev_coe_header: { ...formData.tev_coe_header, equipment_tag: e.target.value }
                    })
                  }
                  className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-purple-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-medium text-slate-700 mb-1">{isVi ? 'Cấp Điện Áp Hệ Thống (kV)' : 'Voltage Rating (kV)'}</label>
                <input
                  type="number"
                  value={formData.tev_coe_header.voltage_rating_kV}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      tev_coe_header: { ...formData.tev_coe_header, voltage_rating_kV: Number(e.target.value) }
                    })
                  }
                  className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-purple-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-medium text-slate-700 mb-1">{isVi ? 'Tỉ Số Biến CCVT' : 'CCVT Ratio Rating'}</label>
                <input
                  type="text"
                  value={formData.tev_coe_header.ccvt_ratio_rating}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      tev_coe_header: { ...formData.tev_coe_header, ccvt_ratio_rating: e.target.value }
                    })
                  }
                  className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-purple-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-medium text-slate-700 mb-1">{isVi ? 'Cấp Chính Xác / Burden' : 'Accuracy Class / Burden'}</label>
                <input
                  type="text"
                  value={formData.tev_coe_header.accuracy_class_rating}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      tev_coe_header: { ...formData.tev_coe_header, accuracy_class_rating: e.target.value }
                    })
                  }
                  className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-purple-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-medium text-slate-700 mb-1">{isVi ? 'Điện Dung C1 Định Mức (pF)' : 'C1 Nameplate (pF)'}</label>
                <input
                  type="number"
                  value={formData.tev_coe_header.c1_nameplate_capacitance_pF}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      tev_coe_header: { ...formData.tev_coe_header, c1_nameplate_capacitance_pF: Number(e.target.value) }
                    })
                  }
                  className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 font-bold text-purple-700 focus:ring-2 focus:ring-purple-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-medium text-slate-700 mb-1">{isVi ? 'Điện Dung C2 Định Mức (pF)' : 'C2 Nameplate (pF)'}</label>
                <input
                  type="number"
                  value={formData.tev_coe_header.c2_nameplate_capacitance_pF}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      tev_coe_header: { ...formData.tev_coe_header, c2_nameplate_capacitance_pF: Number(e.target.value) }
                    })
                  }
                  className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 font-bold text-purple-700 focus:ring-2 focus:ring-purple-500 focus:outline-none"
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
                  className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-purple-500 focus:outline-none"
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
                  ? 'border-purple-600 text-purple-700'
                  : 'border-transparent text-slate-500 hover:text-slate-800'
              }`}
            >
              {isVi ? 'Giai Đoạn 1: Thị Giác & Cơ Khí (7.10.3.A)' : 'Phase 1: Visual & Mechanical (7.10.3.A)'}
            </button>
            <button
              type="button"
              onClick={() => setActivePhase('phase2')}
              className={`py-2 px-4 font-semibold text-xs transition border-b-2 ${
                activePhase === 'phase2'
                  ? 'border-purple-600 text-purple-700'
                  : 'border-transparent text-slate-500 hover:text-slate-800'
              }`}
            >
              {isVi ? 'Giai Đoạn 2: Thử Nghiệm Điện (7.10.3.B)' : 'Phase 2: Electrical Tests (7.10.3.B)'}
            </button>
            <button
              type="button"
              onClick={() => setActivePhase('phase3')}
              className={`py-2 px-4 font-semibold text-xs transition border-b-2 ${
                activePhase === 'phase3'
                  ? 'border-purple-600 text-purple-700'
                  : 'border-transparent text-slate-500 hover:text-slate-800'
              }`}
            >
              {isVi ? 'Giai Đoạn 3: Hồng Ngoại & Sau Thử Nghiệm' : 'Phase 3: Thermography & Post-Test'}
            </button>
          </div>

          {/* PHASE 1: VISUAL & MECHANICAL */}
          {activePhase === 'phase1' && (
            <div className="space-y-4">
              <div className="border border-slate-200 rounded-xl overflow-hidden">
                <div className="bg-slate-100 px-4 py-2 text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-2">
                  <Wrench className="w-3.5 h-3.5 text-purple-600" />
                  {isVi ? 'Kiểm Tra Thị Giác, Vệ Sinh & Cơ Khí (ANSI/NETA MTS-2023 Mục 7.10.3.A)' : 'Visual & Mechanical Inspection (MTS-2023 Sec 7.10.3.A)'}
                </div>

                <div className="p-4 space-y-4 text-xs">
                  {/* Physical condition & Oil level */}
                  <div>
                    <label className="block font-medium text-slate-700 mb-1">
                      {isVi ? 'A1-A2. Tình Trạng Cơ Lý, Sứ Cách Điện & Mức Dầu Hộp Trung Gian (IVT Tank)' : 'A1-A2. Physical Condition, Porcelain Stacks & IVT Oil Level'}
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
                      className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-purple-500 focus:outline-none"
                    />
                  </div>

                  {/* Grounding Switch Operation */}
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                      <label className="block font-medium text-slate-700 mb-1">
                        {isVi ? 'A6. Vận Hành Cơ Khí Dao Tiếp Địa (Grounding Switch Operation - 7.10.3.A.6)' : 'A6. CCVT Grounding Switch Mechanical Operation (7.10.3.A.6)'}
                      </label>
                      <input
                        type="text"
                        value={formData.visual_and_mechanical_inspection.grounding_switch_operation}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            visual_and_mechanical_inspection: {
                              ...formData.visual_and_mechanical_inspection,
                              grounding_switch_operation: e.target.value
                            }
                          })
                        }
                        className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-purple-500 focus:outline-none"
                      />
                      <span className="text-slate-400 text-[11px]">
                        {isVi ? 'Dao đóng cắt nhẹ nhàng, dứt khoát, liên động an toàn' : 'Smooth closing/opening, positive latching, interlock verified'}
                      </span>
                    </div>

                    <div>
                      <label className="block font-medium text-slate-700 mb-1">
                        {isVi ? 'A5. Cuộn Cản Xả Bẫy Sóng PLC & Chống Sét Phụ (Carrier Drain Coil & Gap)' : 'A5. Carrier Drain Coil & Surge Gap Status (7.10.3.A.5)'}
                      </label>
                      <input
                        type="text"
                        value={formData.visual_and_mechanical_inspection.carrier_drain_coil_status}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            visual_and_mechanical_inspection: {
                              ...formData.visual_and_mechanical_inspection,
                              carrier_drain_coil_status: e.target.value
                            }
                          })
                        }
                        className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-purple-500 focus:outline-none"
                      />
                    </div>
                  </div>

                  {/* Bolted Connections */}
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                      <label className="block font-medium text-slate-700 mb-1">
                        {isVi ? 'A4. Độ Chặt Bu-lông Kết Nối (NETA Table 100.12 Torque Verification)' : 'A4. Bolted Connection Torque Tightness (NETA Table 100.12)'}
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
                        className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-purple-500 focus:outline-none"
                      />
                    </div>

                    {/* Secondary Grounding IEEE C57.13.3 */}
                    <div className="p-3 bg-purple-50/60 border border-purple-200 rounded-lg">
                      <div className="flex items-center justify-between mb-1">
                        <label className="font-semibold text-purple-900 flex items-center gap-1.5">
                          <ShieldCheck className="w-3.5 h-3.5 text-purple-600" />
                          {isVi ? 'B9. Điểm Nối Đất Mạch Thứ Cấp (IEEE C57.13.3 - MANDATORY SINGLE POINT)' : 'B9. Secondary Grounding Point (IEEE C57.13.3)'}
                        </label>
                        <span className="text-[11px] font-bold text-purple-800">
                          {isVi ? 'Chỉ duy nhất 1 điểm nối đất!' : 'Single Grounding Point ONLY'}
                        </span>
                      </div>
                      <div className="grid grid-cols-2 gap-2 mt-2">
                        <div>
                          <label className="block text-[11px] text-slate-600 mb-0.5">{isVi ? 'Số điểm đất phát hiện' : 'Points Detected'}</label>
                          <input
                            type="number"
                            min={0}
                            max={5}
                            value={formData.visual_and_mechanical_inspection.secondary_grounding_point_ieee_c57_13_3.ground_points_detected}
                            onChange={(e) =>
                              setFormData({
                                ...formData,
                                visual_and_mechanical_inspection: {
                                  ...formData.visual_and_mechanical_inspection,
                                  secondary_grounding_point_ieee_c57_13_3: {
                                    ...formData.visual_and_mechanical_inspection.secondary_grounding_point_ieee_c57_13_3,
                                    ground_points_detected: Number(e.target.value)
                                  }
                                }
                              })
                            }
                            className={`w-full px-2.5 py-1 text-center font-bold text-sm bg-white border rounded ${
                              formData.visual_and_mechanical_inspection.secondary_grounding_point_ieee_c57_13_3.ground_points_detected === 1
                                ? 'border-emerald-400 text-emerald-700'
                                : 'border-rose-400 text-rose-700'
                            }`}
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] text-slate-600 mb-0.5">{isVi ? 'Vị trí tiếp địa' : 'Grounding Location'}</label>
                          <input
                            type="text"
                            value={formData.visual_and_mechanical_inspection.secondary_grounding_point_ieee_c57_13_3.ground_location}
                            onChange={(e) =>
                              setFormData({
                                ...formData,
                                visual_and_mechanical_inspection: {
                                  ...formData.visual_and_mechanical_inspection,
                                  secondary_grounding_point_ieee_c57_13_3: {
                                    ...formData.visual_and_mechanical_inspection.secondary_grounding_point_ieee_c57_13_3,
                                    ground_location: e.target.value
                                  }
                                }
                              })
                            }
                            className="w-full px-2.5 py-1 bg-white border border-slate-300 rounded text-slate-900"
                          />
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* PHASE 2: ELECTRICAL TESTS */}
          {activePhase === 'phase2' && (
            <div className="space-y-4">
              {/* 1. Capacitance & Dissipation Factor C1/C2 (IEEE C93.1) */}
              <div className="border border-purple-200 bg-purple-50/20 rounded-xl p-4">
                <div className="flex items-center justify-between mb-3 border-b border-purple-200/60 pb-2">
                  <h4 className="text-xs font-bold uppercase tracking-wider text-purple-900 flex items-center gap-2">
                    <Activity className="w-4 h-4 text-purple-600" />
                    {isVi ? '1. Thử Nghiệm Điện Dung C1 & C2 (IEEE C93.1 - Độ Lệch Chuẩn ≤ ±5%)' : '1. C1 & C2 Capacitance Measurement (IEEE C93.1)'}
                  </h4>
                  <span className="text-[11px] font-semibold text-purple-700 bg-purple-100 px-2 py-0.5 rounded">
                    {isVi ? 'Dung sai cho phép: ≤ ±5% Nameplate' : 'Tolerance: ≤ ±5% of Nameplate'}
                  </span>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                  {/* Phase A */}
                  <div className="bg-white p-3 rounded-lg border border-slate-200 shadow-sm space-y-2">
                    <span className="font-bold text-slate-800 text-xs flex items-center gap-1.5">
                      <span className="w-2 h-2 rounded-full bg-amber-500" />
                      {isVi ? 'Pha A (Phase A)' : 'Phase A'}
                    </span>
                    <div>
                      <label className="text-[11px] text-slate-500">{isVi ? 'Điện dung C1 (pF)' : 'C1 Capacitance (pF)'}</label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_3_b.capacitance_measurements_ieee_c93_1.phase_a_c1_pF}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            electrical_tests_7_10_3_b: {
                              ...formData.electrical_tests_7_10_3_b,
                              capacitance_measurements_ieee_c93_1: {
                                ...formData.electrical_tests_7_10_3_b.capacitance_measurements_ieee_c93_1,
                                phase_a_c1_pF: Number(e.target.value)
                              }
                            }
                          })
                        }
                        className="w-full px-2 py-1 border border-slate-300 rounded font-semibold text-slate-800"
                      />
                    </div>
                    <div>
                      <label className="text-[11px] text-slate-500">{isVi ? 'Tổn hao Tan Delta C1 (%)' : 'C1 Tan Delta / PF (%)'}</label>
                      <input
                        type="number"
                        step={0.01}
                        value={formData.electrical_tests_7_10_3_b.power_factor_tan_delta_test.phase_a_c1_pf_percent}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            electrical_tests_7_10_3_b: {
                              ...formData.electrical_tests_7_10_3_b,
                              power_factor_tan_delta_test: {
                                ...formData.electrical_tests_7_10_3_b.power_factor_tan_delta_test,
                                phase_a_c1_pf_percent: Number(e.target.value)
                              }
                            }
                          })
                        }
                        className="w-full px-2 py-1 border border-slate-300 rounded font-semibold text-slate-800"
                      />
                    </div>
                  </div>

                  {/* Phase B */}
                  <div className="bg-white p-3 rounded-lg border-2 border-purple-300 shadow-sm space-y-2">
                    <span className="font-bold text-purple-900 text-xs flex items-center justify-between">
                      <span className="flex items-center gap-1.5">
                        <span className="w-2 h-2 rounded-full bg-purple-600" />
                        {isVi ? 'Pha B (Phase B - Giám Sát Chuyên Sâu)' : 'Phase B (Monitored Phase)'}
                      </span>
                      {Math.abs(formData.electrical_tests_7_10_3_b.capacitance_measurements_ieee_c93_1.c1_deviation_phase_b_percent) > 5.0 && (
                        <span className="text-[10px] text-rose-600 font-extrabold uppercase">FAIL &gt; 5%</span>
                      )}
                    </span>
                    <div>
                      <label className="text-[11px] text-slate-500">{isVi ? 'Điện dung C1 (pF)' : 'C1 Capacitance (pF)'}</label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_3_b.capacitance_measurements_ieee_c93_1.phase_b_c1_pF}
                        onChange={(e) => {
                          const val = Number(e.target.value);
                          const np = formData.tev_coe_header.c1_nameplate_capacitance_pF || 4500;
                          const dev = np > 0 ? ((val - np) / np) * 100 : 0;
                          setFormData({
                            ...formData,
                            electrical_tests_7_10_3_b: {
                              ...formData.electrical_tests_7_10_3_b,
                              capacitance_measurements_ieee_c93_1: {
                                ...formData.electrical_tests_7_10_3_b.capacitance_measurements_ieee_c93_1,
                                phase_b_c1_pF: val,
                                c1_deviation_phase_b_percent: Number(dev.toFixed(2))
                              }
                            }
                          });
                        }}
                        className="w-full px-2 py-1 border border-purple-400 rounded font-bold text-purple-900"
                      />
                    </div>
                    <div>
                      <label className="text-[11px] text-slate-500">{isVi ? 'Độ lệch C1 so với Nhãn Mác (%)' : 'C1 Deviation vs Nameplate (%)'}</label>
                      <input
                        type="number"
                        step={0.01}
                        value={formData.electrical_tests_7_10_3_b.capacitance_measurements_ieee_c93_1.c1_deviation_phase_b_percent}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            electrical_tests_7_10_3_b: {
                              ...formData.electrical_tests_7_10_3_b,
                              capacitance_measurements_ieee_c93_1: {
                                ...formData.electrical_tests_7_10_3_b.capacitance_measurements_ieee_c93_1,
                                c1_deviation_phase_b_percent: Number(e.target.value)
                              }
                            }
                          })
                        }
                        className={`w-full px-2 py-1 border rounded font-bold ${
                          Math.abs(formData.electrical_tests_7_10_3_b.capacitance_measurements_ieee_c93_1.c1_deviation_phase_b_percent) > 5.0
                            ? 'bg-rose-50 border-rose-400 text-rose-700'
                            : 'bg-emerald-50 border-emerald-400 text-emerald-700'
                        }`}
                      />
                    </div>
                    <div>
                      <label className="text-[11px] text-slate-500">{isVi ? 'Tổn hao Tan Delta C1 (%) (Chuẩn ≤ 0.5%)' : 'C1 Tan Delta / PF (%) (Std ≤ 0.5%)'}</label>
                      <input
                        type="number"
                        step={0.01}
                        value={formData.electrical_tests_7_10_3_b.power_factor_tan_delta_test.phase_b_c1_pf_percent}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            electrical_tests_7_10_3_b: {
                              ...formData.electrical_tests_7_10_3_b,
                              power_factor_tan_delta_test: {
                                ...formData.electrical_tests_7_10_3_b.power_factor_tan_delta_test,
                                phase_b_c1_pf_percent: Number(e.target.value)
                              }
                            }
                          })
                        }
                        className={`w-full px-2 py-1 border rounded font-bold ${
                          formData.electrical_tests_7_10_3_b.power_factor_tan_delta_test.phase_b_c1_pf_percent > 0.50
                            ? 'bg-rose-50 border-rose-400 text-rose-700'
                            : 'bg-emerald-50 border-emerald-400 text-emerald-700'
                        }`}
                      />
                    </div>
                  </div>

                  {/* Phase C */}
                  <div className="bg-white p-3 rounded-lg border border-slate-200 shadow-sm space-y-2">
                    <span className="font-bold text-slate-800 text-xs flex items-center gap-1.5">
                      <span className="w-2 h-2 rounded-full bg-sky-500" />
                      {isVi ? 'Pha C (Phase C)' : 'Phase C'}
                    </span>
                    <div>
                      <label className="text-[11px] text-slate-500">{isVi ? 'Điện dung C1 (pF)' : 'C1 Capacitance (pF)'}</label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_3_b.capacitance_measurements_ieee_c93_1.phase_c_c1_pF || 4500}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            electrical_tests_7_10_3_b: {
                              ...formData.electrical_tests_7_10_3_b,
                              capacitance_measurements_ieee_c93_1: {
                                ...formData.electrical_tests_7_10_3_b.capacitance_measurements_ieee_c93_1,
                                phase_c_c1_pF: Number(e.target.value)
                              }
                            }
                          })
                        }
                        className="w-full px-2 py-1 border border-slate-300 rounded font-semibold text-slate-800"
                      />
                    </div>
                    <div>
                      <label className="text-[11px] text-slate-500">{isVi ? 'Tổn hao Tan Delta C1 (%)' : 'C1 Tan Delta / PF (%)'}</label>
                      <input
                        type="number"
                        step={0.01}
                        value={formData.electrical_tests_7_10_3_b.power_factor_tan_delta_test.phase_c_c1_pf_percent}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            electrical_tests_7_10_3_b: {
                              ...formData.electrical_tests_7_10_3_b,
                              power_factor_tan_delta_test: {
                                ...formData.electrical_tests_7_10_3_b.power_factor_tan_delta_test,
                                phase_c_c1_pf_percent: Number(e.target.value)
                              }
                            }
                          })
                        }
                        className="w-full px-2 py-1 border border-slate-300 rounded font-semibold text-slate-800"
                      />
                    </div>
                  </div>
                </div>
              </div>

              {/* 2. Insulation Resistance (Table 100.5) */}
              <div className="border border-slate-200 rounded-xl p-4">
                <h4 className="text-xs font-bold uppercase tracking-wider text-slate-700 mb-3 flex items-center gap-2">
                  <Gauge className="w-4 h-4 text-purple-600" />
                  {isVi ? '2. Điện Trở Cách Điện Megohm (NETA MTS Table 100.5)' : '2. Insulation Resistance Megohm (Table 100.5)'}
                </h4>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
                  <div>
                    <label className="block text-slate-700 font-medium mb-1">
                      {isVi ? 'Sơ Cấp / IVT Đối Đất (5000V DC - Chuẩn ≥ 5,000 MΩ)' : 'Primary IVT to Ground (5000V DC - Std ≥ 5,000 MΩ)'}
                    </label>
                    <input
                      type="number"
                      value={formData.electrical_tests_7_10_3_b.insulation_resistance_megohms_table_100_5.primary_ivt_to_ground_5000v_dc_Mohm}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          electrical_tests_7_10_3_b: {
                            ...formData.electrical_tests_7_10_3_b,
                            insulation_resistance_megohms_table_100_5: {
                              ...formData.electrical_tests_7_10_3_b.insulation_resistance_megohms_table_100_5,
                              primary_ivt_to_ground_5000v_dc_Mohm: Number(e.target.value)
                            }
                          }
                        })
                      }
                      className="w-full px-3 py-1.5 border border-slate-300 rounded-lg text-slate-900 font-semibold"
                    />
                  </div>

                  <div>
                    <label className="block text-slate-700 font-medium mb-1">
                      {isVi ? 'Thứ Cấp Đối Đất (1000V DC - Chuẩn ≥ 100 MΩ)' : 'Secondary to Ground (1000V DC - Std ≥ 100 MΩ)'}
                    </label>
                    <input
                      type="number"
                      value={formData.electrical_tests_7_10_3_b.insulation_resistance_megohms_table_100_5.secondary_to_ground_1000v_dc_Mohm}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          electrical_tests_7_10_3_b: {
                            ...formData.electrical_tests_7_10_3_b,
                            insulation_resistance_megohms_table_100_5: {
                              ...formData.electrical_tests_7_10_3_b.insulation_resistance_megohms_table_100_5,
                              secondary_to_ground_1000v_dc_Mohm: Number(e.target.value)
                            }
                          }
                        })
                      }
                      className="w-full px-3 py-1.5 border border-slate-300 rounded-lg text-slate-900 font-semibold"
                    />
                  </div>
                </div>
              </div>

              {/* 3. Turns Ratio & Secondary Voltage Verification */}
              <div className="border border-slate-200 rounded-xl p-4">
                <h4 className="text-xs font-bold uppercase tracking-wider text-slate-700 mb-3 flex items-center gap-2">
                  <Zap className="w-4 h-4 text-purple-600" />
                  {isVi ? '3. Thử Nghiệm Tỉ Số Biến Áp & Cực Tính (Turns-Ratio & Polarity - IEEE C93.1)' : '3. Turns Ratio Verification & Polarity (IEEE C93.1)'}
                </h4>
                <div className="grid grid-cols-1 md:grid-cols-4 gap-4 text-xs">
                  <div>
                    <label className="block text-slate-600 mb-1">{isVi ? 'Điện áp sơ cấp đặt vào (V)' : 'Applied Primary Voltage (V)'}</label>
                    <input
                      type="number"
                      value={formData.electrical_tests_7_10_3_b.turns_ratio_verification_test_ieee_c93_1.applied_primary_voltage_V}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          electrical_tests_7_10_3_b: {
                            ...formData.electrical_tests_7_10_3_b,
                            turns_ratio_verification_test_ieee_c93_1: {
                              ...formData.electrical_tests_7_10_3_b.turns_ratio_verification_test_ieee_c93_1,
                              applied_primary_voltage_V: Number(e.target.value)
                            }
                          }
                        })
                      }
                      className="w-full px-2.5 py-1.5 border border-slate-300 rounded text-slate-900"
                    />
                  </div>

                  <div>
                    <label className="block text-slate-600 mb-1">{isVi ? 'Áp thứ cấp Pha B đo được (V)' : 'Sec Voltage Phase B (V)'}</label>
                    <input
                      type="number"
                      step={0.1}
                      value={formData.electrical_tests_7_10_3_b.turns_ratio_verification_test_ieee_c93_1.measured_secondary_voltage_phase_b_V}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          electrical_tests_7_10_3_b: {
                            ...formData.electrical_tests_7_10_3_b,
                            turns_ratio_verification_test_ieee_c93_1: {
                              ...formData.electrical_tests_7_10_3_b.turns_ratio_verification_test_ieee_c93_1,
                              measured_secondary_voltage_phase_b_V: Number(e.target.value)
                            }
                          }
                        })
                      }
                      className="w-full px-2.5 py-1.5 border border-slate-300 rounded font-bold text-slate-900"
                    />
                  </div>

                  <div>
                    <label className="block text-slate-600 mb-1">{isVi ? 'Sai số tỉ số biến Pha B (%)' : 'Ratio Error Phase B (%)'}</label>
                    <input
                      type="number"
                      step={0.01}
                      value={formData.electrical_tests_7_10_3_b.turns_ratio_verification_test_ieee_c93_1.ratio_error_phase_b_percent}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          electrical_tests_7_10_3_b: {
                            ...formData.electrical_tests_7_10_3_b,
                            turns_ratio_verification_test_ieee_c93_1: {
                              ...formData.electrical_tests_7_10_3_b.turns_ratio_verification_test_ieee_c93_1,
                              ratio_error_phase_b_percent: Number(e.target.value)
                            }
                          }
                        })
                      }
                      className={`w-full px-2.5 py-1.5 border rounded font-bold ${
                        Math.abs(formData.electrical_tests_7_10_3_b.turns_ratio_verification_test_ieee_c93_1.ratio_error_phase_b_percent) > 1.2
                          ? 'bg-rose-50 border-rose-400 text-rose-700'
                          : 'bg-emerald-50 border-emerald-400 text-emerald-700'
                      }`}
                    />
                  </div>

                  <div>
                    <label className="block text-slate-600 mb-1">{isVi ? 'Thử Cực Tính H1-X1' : 'Polarity Test'}</label>
                    <input
                      type="text"
                      value={formData.electrical_tests_7_10_3_b.polarity_test_ieee_c93_1}
                      onChange={(e) =>
                        setFormData({
                          ...formData,
                          electrical_tests_7_10_3_b: {
                            ...formData.electrical_tests_7_10_3_b,
                            polarity_test_ieee_c93_1: e.target.value
                          }
                        })
                      }
                      className="w-full px-2.5 py-1.5 border border-slate-300 rounded text-slate-900"
                    />
                  </div>
                </div>
              </div>

              {/* 4. Secondary Burden & Dielectric Withstand */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
                <div className="border border-slate-200 rounded-xl p-4">
                  <h4 className="font-bold text-slate-700 mb-2">{isVi ? 'Đo Phụ Tải Thứ Cấp (Burden Measurement)' : 'Secondary Burden Measurement'}</h4>
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-slate-600 mb-1">{isVi ? 'Phụ tải đo được (VA)' : 'Measured Burden (VA)'}</label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_3_b.secondary_burden_measurement.measured_burden_phase_b_VA}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            electrical_tests_7_10_3_b: {
                              ...formData.electrical_tests_7_10_3_b,
                              secondary_burden_measurement: {
                                ...formData.electrical_tests_7_10_3_b.secondary_burden_measurement,
                                measured_burden_phase_b_VA: Number(e.target.value)
                              }
                            }
                          })
                        }
                        className="w-full px-2.5 py-1.5 border border-slate-300 rounded text-slate-900 font-semibold"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-600 mb-1">{isVi ? 'Burden định mức (VA)' : 'Rated Burden (VA)'}</label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_3_b.secondary_burden_measurement.ccvt_rated_burden_VA}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            electrical_tests_7_10_3_b: {
                              ...formData.electrical_tests_7_10_3_b,
                              secondary_burden_measurement: {
                                ...formData.electrical_tests_7_10_3_b.secondary_burden_measurement,
                                ccvt_rated_burden_VA: Number(e.target.value)
                              }
                            }
                          })
                        }
                        className="w-full px-2.5 py-1.5 border border-slate-300 rounded text-slate-900"
                      />
                    </div>
                  </div>
                </div>

                <div className="border border-slate-200 rounded-xl p-4">
                  <h4 className="font-bold text-slate-700 mb-2">{isVi ? 'Điện Áp Chịu Đựng Tần Số Nguồn (NETA Table 100.9)' : 'Dielectric Withstand Test (Table 100.9)'}</h4>
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-slate-600 mb-1">{isVi ? 'Điện áp thử (kV AC)' : 'Test Voltage (kV AC)'}</label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_3_b.dielectric_withstand_test_table_100_9.test_voltage_kV_ac}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            electrical_tests_7_10_3_b: {
                              ...formData.electrical_tests_7_10_3_b,
                              dielectric_withstand_test_table_100_9: {
                                ...formData.electrical_tests_7_10_3_b.dielectric_withstand_test_table_100_9,
                                test_voltage_kV_ac: Number(e.target.value)
                              }
                            }
                          })
                        }
                        className="w-full px-2.5 py-1.5 border border-slate-300 rounded text-slate-900"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-600 mb-1">{isVi ? 'Thời gian duy trì (Phút)' : 'Duration (Minutes)'}</label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_3_b.dielectric_withstand_test_table_100_9.duration_minutes}
                        onChange={(e) =>
                          setFormData({
                            ...formData,
                            electrical_tests_7_10_3_b: {
                              ...formData.electrical_tests_7_10_3_b,
                              dielectric_withstand_test_table_100_9: {
                                ...formData.electrical_tests_7_10_3_b.dielectric_withstand_test_table_100_9,
                                duration_minutes: Number(e.target.value)
                              }
                            }
                          })
                        }
                        className="w-full px-2.5 py-1.5 border border-slate-300 rounded text-slate-900"
                      />
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* PHASE 3: THERMOGRAPHY */}
          {activePhase === 'phase3' && (
            <div className="space-y-4 text-xs">
              <div className="border border-slate-200 rounded-xl p-4">
                <h4 className="font-bold text-slate-700 mb-2 flex items-center gap-2">
                  <Thermometer className="w-4 h-4 text-rose-600" />
                  {isVi ? 'Khảo Sát Hồng Ngoại Nhiệt Độ Tiếp Xúc (NETA MTS Table 100.18)' : 'Thermographic Survey & Thermal Inspection (NETA Table 100.18)'}
                </h4>
                <p className="text-slate-500 mb-3">
                  {isVi
                    ? 'Kiểm tra chênh lệch nhiệt độ giữa các điểm đầu cực cao thế H1, hộp đấu dây trung gian IVT và mối nối bu-lông. Sai lệch ΔT > 4°C yêu cầu bảo dưỡng; ΔT > 15°C yêu cầu cô lập tức thì.'
                    : 'Survey temperature gradients across high-voltage terminals, intermediate transformer casing, and bolted joints per NETA Table 100.18.'}
                </p>
                <textarea
                  rows={3}
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
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg text-slate-900 focus:ring-2 focus:ring-purple-500 focus:outline-none"
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
            const currentEval = evaluationResult || performDeterministicMtsCcvtAnalysis(formData);
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
                      <Download className="w-3.5 h-3.5 text-purple-600" />
                      <span>{isVi ? 'Xuất PDF Báo Cáo' : 'Export PDF'}</span>
                    </button>
                  </div>
                </div>

                {/* Gemini AI Detailed Markdown Output (if available) */}
                {currentEval.ai_generated_markdown && (
                  <div className="bg-gradient-to-br from-purple-50/50 via-white to-indigo-50/30 rounded-2xl border border-purple-200 p-5 shadow-sm">
                    <div className="flex items-center gap-2 mb-3 pb-2 border-b border-purple-200/60">
                      <Sparkles className="w-4 h-4 text-purple-600" />
                      <h4 className="text-xs font-bold uppercase tracking-wider text-purple-900">
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
                      <FileCheck2 className="w-4 h-4 text-purple-400" />
                      {isVi
                        ? 'Bảng Đối Soát Thông Số Thử Nghiệm CCVT Theo ANSI/NETA MTS-2023 Mục 7.10.3'
                        : 'CCVT Maintenance Test Parameter Verification Table (NETA MTS-2023 Sec 7.10.3)'}
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
                      <Activity className="w-4 h-4 text-purple-600" />
                      {isVi ? 'Nhận Định & Phân Tích Kỹ Thuật Hệ Thống' : 'Technical & Diagnostic Analysis'}
                    </h4>
                    <ul className="space-y-2 text-xs text-slate-700">
                      {(isVi ? currentEval.technical_analysis_vi : currentEval.technical_analysis_en).map((item, idx) => (
                        <li key={idx} className="flex items-start gap-2">
                          <span className="w-1.5 h-1.5 rounded-full bg-purple-600 mt-1.5 shrink-0" />
                          <span>{item}</span>
                        </li>
                      ))}
                    </ul>
                  </div>

                  {/* Recommendations */}
                  <div className="bg-purple-50/40 rounded-xl p-4 border border-purple-200">
                    <h4 className="text-xs font-bold uppercase tracking-wider text-purple-900 mb-3 flex items-center gap-2">
                      <Wrench className="w-4 h-4 text-purple-600" />
                      {isVi ? 'Khuyến Nghị Hành Động Dành Cho FSE Hiện Trường' : 'Actionable Recommendations for FSE'}
                    </h4>
                    <ul className="space-y-2 text-xs text-slate-800">
                      {(isVi ? currentEval.recommendations_vi : currentEval.recommendations_en).map((rec, idx) => (
                        <li key={idx} className="flex items-start gap-2">
                          <span className="font-bold text-purple-700 shrink-0">{idx + 1}.</span>
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
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-purple-600 hover:bg-purple-700 text-white text-xs font-medium shadow transition"
            >
              {copySuccess ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
              <span>{copySuccess ? (isVi ? 'Đã Sao Chép!' : 'Copied!') : (isVi ? 'Sao Chép Prompt' : 'Copy Prompt')}</span>
            </button>
          </div>

          <pre className="p-4 bg-slate-900 text-purple-200 text-xs font-mono rounded-xl overflow-x-auto whitespace-pre-wrap max-h-[500px] border border-slate-800">
            {MTS_CCVT_SYSTEM_INSTRUCTION}
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
                  ? 'Cấu trúc dữ liệu JSON đầy đủ phục vụ gọi API /api/field-service/neta-mts-ccvt-analyze.'
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
