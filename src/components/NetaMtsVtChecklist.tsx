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
  HelpCircle
} from 'lucide-react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  NetaMtsVtInputPayload,
  NetaMtsVtEvaluationResult,
  performDeterministicMtsVtAnalysis,
  MTS_VT_SYSTEM_INSTRUCTION
} from '../../server/netaMtsVtAnalyzer';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
  initialLanguage?: 'vi' | 'en';
}

export const SAMPLE_MTS_VT_DEFECT_PAYLOAD: NetaMtsVtInputPayload = {
  tev_coe_header: {
    module: "VT-MAINTENANCE-MTS-7.10.2",
    project_name: "Tram Bien Ap Trung The TEV Commercial Substation",
    equipment_tag: "22KV-BUS-VT-SET (Set of 3 Single-Phase VTs)",
    voltage_rating_kV: 22,
    vt_ratio_rating: "22000/sqrt(3) : 110/sqrt(3) V",
    accuracy_class_rating: "0.2 / 3P 50VA",
    fse_name: "Nguyen Van Q",
    test_date: "2026-10-19",
    station_name: "Substation TEV Commercial",
    bay_location: "Busbar VT Cubicle Bay 02"
  },
  visual_and_mechanical_inspection: {
    physical_condition: "Good condition, clean resin surface, no cracks or oil leakage observed",
    primary_fuse_status: "Pass (Current-limiting primary fuses intact and verified)",
    bolted_connection_torque_table_100_12: "Pass (Torque wrench verified per Table 100.12)",
    secondary_grounding_point_ieee_c57_13_3: {
      ground_points_detected: 1,
      ground_location: "Grounded at Protection Relay Panel N-bus star point",
      status: "Pass (Single grounding point verified per IEEE C57.13.3)"
    },
    fluid_or_gas_level_check: "Pass (Cast resin dry-type insulation intact)"
  },
  electrical_tests_7_10_2_b: {
    insulation_resistance_megohms_table_100_5: {
      primary_to_ground_2500v_dc_Mohm: {
        phase_a: 8500.0,
        phase_b: 42.0,
        phase_c: 8200.0
      },
      secondary_to_ground_1000v_dc_Mohm: 380.0,
      status: "Fail (Phase B primary insulation 42.0 MΩ fails NETA Table 100.5 minimum 1,000 MΩ threshold)"
    },
    polarity_test_ieee_c57_13_1: "Pass (H1-X1 correct polarity verified on all 3 phases)",
    turns_ratio_verification_test_ieee_c57_13: {
      applied_primary_voltage_V: 12702,
      measured_secondary_voltage_phase_a_V: 63.51,
      measured_secondary_voltage_phase_b_V: 61.20,
      measured_secondary_voltage_phase_c_V: 63.50,
      ratio_error_phase_b_percent: -3.64,
      status: "Fail (Phase B ratio error -3.64% exceeds 0.2 accuracy class limit)"
    },
    secondary_burden_measurement: {
      measured_burden_phase_a_VA: 14.2,
      measured_burden_phase_b_VA: 14.5,
      measured_burden_phase_c_VA: 14.1,
      vt_rated_burden_VA: 50.0,
      status: "Pass (Measured burden well within 50VA VT rating)"
    },
    dielectric_withstand_test_table_100_9: {
      test_voltage_kV_ac: 32.5,
      duration_minutes: 1,
      phase_b_status: "Fail (Dielectric breakdown / tracking observed on Phase B at 21kV AC)"
    },
    power_factor_test_percent: {
      measured_pf_percent: 0.42,
      status: "Pass"
    }
  },
  post_test_survey: {
    thermographic_survey_table_100_18: "Pass (Normal thermal distribution, no hotspot detected)"
  }
};

export const SAMPLE_MTS_VT_PASS_PAYLOAD: NetaMtsVtInputPayload = {
  tev_coe_header: {
    module: "VT-MAINTENANCE-MTS-7.10.2",
    project_name: "Tram 110kV KCN Song Khoai TEV Substation",
    equipment_tag: "110KV-BUS-VT-CVT-SET",
    voltage_rating_kV: 110,
    vt_ratio_rating: "110000/sqrt(3) : 110/sqrt(3) V",
    accuracy_class_rating: "0.2 / 3P 75VA",
    fse_name: "Tran Minh T",
    test_date: "2026-10-18",
    station_name: "110kV Song Khoai",
    bay_location: "Busbar 110kV C11"
  },
  visual_and_mechanical_inspection: {
    physical_condition: "Pristine porcelain bushing, oil level normal, no weeping or crack",
    primary_fuse_status: "Pass (Solidly grounded CVT connection per design)",
    bolted_connection_torque_table_100_12: "Pass (Calibrated torque wrench verified per Table 100.12)",
    secondary_grounding_point_ieee_c57_13_3: {
      ground_points_detected: 1,
      ground_location: "Grounded at Protection Relay Panel MK cabinet ONLY",
      status: "Pass (Single grounding point verified per IEEE C57.13.3)"
    },
    fluid_or_gas_level_check: "Pass (Oil level indicator in normal green zone at 25°C)"
  },
  electrical_tests_7_10_2_b: {
    insulation_resistance_megohms_table_100_5: {
      primary_to_ground_2500v_dc_Mohm: {
        phase_a: 19500.0,
        phase_b: 18900.0,
        phase_c: 19200.0
      },
      secondary_to_ground_1000v_dc_Mohm: 890.0,
      status: "Pass (Exceeds Table 100.5 thresholds >= 5,000 MΩ)"
    },
    polarity_test_ieee_c57_13_1: "Pass (H1-X1 polarity verified with DC pulse test)",
    turns_ratio_verification_test_ieee_c57_13: {
      applied_primary_voltage_V: 63508,
      measured_secondary_voltage_phase_a_V: 63.52,
      measured_secondary_voltage_phase_b_V: 63.50,
      measured_secondary_voltage_phase_c_V: 63.51,
      ratio_error_phase_b_percent: -0.01,
      status: "Pass (Ratio error well within +/- 0.2% class)"
    },
    secondary_burden_measurement: {
      measured_burden_phase_a_VA: 22.1,
      measured_burden_phase_b_VA: 21.8,
      measured_burden_phase_c_VA: 22.0,
      vt_rated_burden_VA: 75.0,
      status: "Pass (Burden well within rated 75VA)"
    },
    dielectric_withstand_test_table_100_9: {
      test_voltage_kV_ac: 149.0,
      duration_minutes: 1,
      phase_b_status: "Pass (Withstand passed without breakdown)"
    },
    power_factor_test_percent: {
      measured_pf_percent: 0.35,
      status: "Pass (Power factor 0.35% < 0.5% limit at 20°C)"
    }
  },
  post_test_survey: {
    thermographic_survey_table_100_18: "Pass (Thermal gradient < 1.0°C per Table 100.18)"
  }
};

export const SAMPLE_MTS_VT_CAUTION_PAYLOAD: NetaMtsVtInputPayload = {
  tev_coe_header: {
    module: "VT-MAINTENANCE-MTS-7.10.2",
    project_name: "Tram Phap Van TEV Distribution Center",
    equipment_tag: "24KV-BUS-VT-01-SET",
    voltage_rating_kV: 24,
    vt_ratio_rating: "24000/sqrt(3) : 120/sqrt(3) V",
    accuracy_class_rating: "0.5 / 3P 30VA",
    fse_name: "Le Hoang D",
    test_date: "2026-10-17",
    station_name: "Dist Substation D-03",
    bay_location: "VT Compartment 24kV"
  },
  visual_and_mechanical_inspection: {
    physical_condition: "Minor dust accumulation on resin surface, cleaned prior to test",
    primary_fuse_status: "Caution (Phase C secondary fuse holder oxidized, contact resistance elevated)",
    bolted_connection_torque_table_100_12: "Pass",
    secondary_grounding_point_ieee_c57_13_3: {
      ground_points_detected: 1,
      ground_location: "Relay Panel single point",
      status: "Pass"
    },
    fluid_or_gas_level_check: "Pass"
  },
  electrical_tests_7_10_2_b: {
    insulation_resistance_megohms_table_100_5: {
      primary_to_ground_2500v_dc_Mohm: {
        phase_a: 5400.0,
        phase_b: 5100.0,
        phase_c: 4800.0
      },
      secondary_to_ground_1000v_dc_Mohm: 160.0,
      status: "Pass (Meets Table 100.5 thresholds)"
    },
    polarity_test_ieee_c57_13_1: "Pass (H1-X1 correct)",
    turns_ratio_verification_test_ieee_c57_13: {
      applied_primary_voltage_V: 13856,
      measured_secondary_voltage_phase_a_V: 69.25,
      measured_secondary_voltage_phase_b_V: 69.20,
      measured_secondary_voltage_phase_c_V: 68.90,
      ratio_error_phase_b_percent: -0.12,
      status: "Caution (Phase C ratio error -0.55% borderline for 0.5 class)"
    },
    secondary_burden_measurement: {
      measured_burden_phase_a_VA: 28.5,
      measured_burden_phase_b_VA: 29.2,
      measured_burden_phase_c_VA: 31.8,
      vt_rated_burden_VA: 30.0,
      status: "Fail (Phase C burden 31.8 VA exceeds 30.0 VA rated capacity, causing extra voltage drop)"
    },
    dielectric_withstand_test_table_100_9: {
      test_voltage_kV_ac: 37.0,
      duration_minutes: 1,
      phase_b_status: "Pass"
    }
  },
  post_test_survey: {
    thermographic_survey_table_100_18: "Pass"
  }
};

export const NetaMtsVtChecklist: React.FC<Props> = ({
  onSaveReport,
  onNavigateToReports,
  initialLanguage = 'vi'
}) => {
  const [lang, setLang] = useState<'vi' | 'en'>(initialLanguage);
  const [activeTab, setActiveTab] = useState<'checklist' | 'evaluation' | 'prompt' | 'json'>('checklist');
  const [activePhase, setActivePhase] = useState<'phase1' | 'phase2' | 'phase3'>('phase1');
  const [isEvaluating, setIsEvaluating] = useState(false);
  const [evaluationResult, setEvaluationResult] = useState<NetaMtsVtEvaluationResult | null>(null);
  const [copySuccess, setCopySuccess] = useState(false);
  const [jsonCopySuccess, setJsonCopySuccess] = useState(false);
  const [saveSuccess, setSaveSuccess] = useState(false);

  // Form State
  const [formData, setFormData] = useState<NetaMtsVtInputPayload>(SAMPLE_MTS_VT_DEFECT_PAYLOAD);

  const isVi = lang === 'vi';

  // Handler to load presets
  const handleLoadPreset = (scenario: 'defect' | 'pass' | 'caution') => {
    let payload = SAMPLE_MTS_VT_DEFECT_PAYLOAD;
    if (scenario === 'pass') payload = SAMPLE_MTS_VT_PASS_PAYLOAD;
    if (scenario === 'caution') payload = SAMPLE_MTS_VT_CAUTION_PAYLOAD;

    setFormData(JSON.parse(JSON.stringify(payload)));
    setEvaluationResult(null);
  };

  // Run AI / Deterministic Evaluation
  const handleRunEvaluation = async () => {
    setIsEvaluating(true);
    try {
      const response = await fetch('/api/field-service/neta-mts-vt-analyze', {
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
      const localEval = performDeterministicMtsVtAnalysis(formData);
      setEvaluationResult(localEval);
      setActiveTab('evaluation');
    } catch (err) {
      console.warn('API error, using local deterministic engine:', err);
      const localEval = performDeterministicMtsVtAnalysis(formData);
      setEvaluationResult(localEval);
      setActiveTab('evaluation');
    } finally {
      setIsEvaluating(false);
    }
  };

  // Copy AI Studio System Instruction Prompt
  const handleCopyPrompt = () => {
    navigator.clipboard.writeText(MTS_VT_SYSTEM_INSTRUCTION);
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
      const computed = performDeterministicMtsVtAnalysis(formData);
      setEvaluationResult(computed);
    }
    const reportData = {
      id: `NETA-MTS-VT-${Date.now()}`,
      moduleType: 'MTS-VT-7.10.2',
      title: `${formData.tev_coe_header.equipment_tag} - NETA MTS-2023 Sec 7.10.2`,
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
    const currentEval = evaluationResult || performDeterministicMtsVtAnalysis(formData);

    // Header
    doc.setFillColor(15, 23, 42);
    doc.rect(0, 0, 210, 26, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(14);
    doc.text('TEV COE - ELECTRICAL TESTING & COMMISSIONING PLATFORM', 14, 11);
    doc.setFontSize(10);
    doc.setTextColor(148, 163, 184);
    doc.text('ANSI/NETA MTS-2023 SECTION 7.10.2 - INSTRUMENT TRANSFORMERS (VT/PT)', 14, 18);

    // Info Box
    doc.setFillColor(248, 250, 252);
    doc.rect(14, 32, 182, 30, 'F');
    doc.setDrawColor(203, 213, 225);
    doc.rect(14, 32, 182, 30, 'S');

    doc.setTextColor(15, 23, 42);
    doc.setFontSize(9);
    doc.text(`Project: ${formData.tev_coe_header.project_name}`, 18, 40);
    doc.text(`Equipment Tag: ${formData.tev_coe_header.equipment_tag}`, 18, 46);
    doc.text(`Ratio: ${formData.tev_coe_header.vt_ratio_rating} | Class: ${formData.tev_coe_header.accuracy_class_rating}`, 18, 52);
    doc.text(`Voltage: ${formData.tev_coe_header.voltage_rating_kV} kV`, 18, 58);

    doc.text(`FSE Engineer: ${formData.tev_coe_header.fse_name}`, 115, 40);
    doc.text(`Test Date: ${formData.tev_coe_header.test_date}`, 115, 46);
    doc.text(`Location: ${formData.tev_coe_header.bay_location || 'VT Cubicle'}`, 115, 52);
    doc.text(`Standard: ANSI/NETA MTS-2023 (Sec 7.10.2)`, 115, 58);

    // Severity Box
    let sevColor: [number, number, number] = [16, 185, 129];
    if (currentEval.overall_status === 'CRITICAL') sevColor = [225, 29, 72];
    if (currentEval.overall_status === 'CAUTION') sevColor = [245, 158, 11];

    doc.setFillColor(...sevColor);
    doc.rect(14, 66, 182, 10, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(10);
    const sevTitle = isVi ? currentEval.overall_title_vi : currentEval.overall_title_en;
    doc.text(`OVERALL STATUS: ${currentEval.overall_status} - ${sevTitle}`, 18, 73);

    // Evaluation Table
    const tableRows = currentEval.evaluation_table.map((row) => [
      isVi ? row.item_vn : row.item,
      row.standard_ref,
      row.measured_value,
      row.status,
      isVi ? row.notes_vn : row.notes
    ]);

    autoTable(doc, {
      startY: 80,
      head: [[
        isVi ? 'Hạng mục kiểm tra' : 'Test Item',
        isVi ? 'Tiêu chuẩn tham chiếu' : 'Standard Ref',
        isVi ? 'Giá trị đo thực tế' : 'Measured Value',
        isVi ? 'Kết quả' : 'Status',
        isVi ? 'Ghi chú kỹ thuật' : 'Technical Diagnostic'
      ]],
      body: tableRows,
      theme: 'grid',
      headStyles: { fillColor: [30, 41, 59], textColor: [255, 255, 255], fontSize: 8 },
      bodyStyles: { fontSize: 8 },
      columnStyles: {
        0: { cellWidth: 38 },
        1: { cellWidth: 32 },
        2: { cellWidth: 40 },
        3: { cellWidth: 20 },
        4: { cellWidth: 52 },
      },
      didParseCell: (data) => {
        if (data.section === 'body' && data.column.index === 3) {
          const val = data.cell.raw;
          if (val === 'PASS') data.cell.styles.textColor = [16, 185, 129];
          if (val === 'FAIL' || val === 'CRITICAL') data.cell.styles.textColor = [225, 29, 72];
          if (val === 'CAUTION') data.cell.styles.textColor = [245, 158, 11];
        }
      }
    });

    const finalY = (doc as any).lastAutoTable.finalY + 8;
    if (finalY < 240) {
      doc.setFontSize(9);
      doc.setTextColor(15, 23, 42);
      doc.text(isVi ? 'Khuyến nghị hành động kỹ sư FSE:' : 'Actionable Recommendations for FSE:', 14, finalY);
      const recs = isVi ? currentEval.recommendations_vi : currentEval.recommendations_en;
      recs.slice(0, 3).forEach((r, idx) => {
        doc.setFontSize(8);
        doc.setTextColor(51, 65, 85);
        const splitText = doc.splitTextToSize(`${idx + 1}. ${r}`, 180);
        doc.text(splitText, 14, finalY + 6 + (idx * 11));
      });
    }

    doc.setFontSize(8);
    doc.setTextColor(100, 116, 139);
    doc.text('Tested by (FSE Signature): _______________________', 14, 280);
    doc.text('Approved by (Lead Commissioning): _______________________', 110, 280);

    doc.save(`NETA_MTS_VT_${formData.tev_coe_header.equipment_tag}_${formData.tev_coe_header.test_date}.pdf`);
  };

  return (
    <div className="space-y-6">
      {/* Top Banner */}
      <div className="bg-gradient-to-r from-slate-900 via-sky-950 to-slate-900 rounded-2xl p-6 text-white shadow-xl border border-sky-900/40 relative overflow-hidden">
        <div className="absolute right-0 top-0 w-96 h-96 bg-sky-500/10 rounded-full blur-3xl -mr-20 -mt-20 pointer-events-none" />

        <div className="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-6 relative z-10">
          <div>
            <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full text-xs font-semibold bg-sky-500/20 text-sky-300 border border-sky-400/30 mb-3">
              <Zap size={14} className="text-sky-400" />
              <span>ANSI/NETA MTS-2023 Mục 7.10.2 • Bảng 100.5, 100.9, 100.12, 100.18</span>
            </div>
            <h1 className="text-2xl lg:text-3xl font-extrabold tracking-tight">
              {isVi
                ? 'Kiểm Tra & Chẩn Đoán Biến Điện Áp Đo Lường & Bảo Vệ (Voltage Transformers - VT/PT)'
                : 'Instrument Transformers: Voltage Transformers (VT/PT) Maintenance Testing & Diagnostics'}
            </h1>
            <p className="text-slate-300 text-sm mt-1 max-w-3xl">
              {isVi
                ? 'Quy trình thử nghiệm bảo trì biến điện áp VT/PT chuẩn hóa theo ANSI/NETA MTS-2023 Section 7.10.2, IEEE C57.13 (Cực tính, Tỉ số biến áp, Phụ tải Burden) và IEEE C57.13.3 (Quy chuẩn DUY NHẤT 1 điểm nối đất mạch thứ cấp).'
                : 'VT/PT maintenance testing protocol standardized per ANSI/NETA MTS-2023 Sec 7.10.2, IEEE C57.13 (Polarity, Turns Ratio, Burden) and IEEE C57.13.3 (MANDATORY single secondary grounding point).'}
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => setLang(isVi ? 'en' : 'vi')}
              className="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-xs font-medium border border-slate-700 text-slate-200 transition"
            >
              {isVi ? 'English (EN)' : 'Tiếng Việt (VN)'}
            </button>

            <button
              onClick={handleRunEvaluation}
              disabled={isEvaluating}
              className="flex items-center gap-2 px-4 py-2 rounded-xl bg-gradient-to-r from-sky-500 to-sky-600 hover:from-sky-600 hover:to-sky-700 text-white font-semibold shadow-lg shadow-sky-500/30 text-sm transition"
            >
              {isEvaluating ? (
                <>
                  <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                  <span>{isVi ? 'Đang Chẩn Đoán AI...' : 'Analyzing with AI...'}</span>
                </>
              ) : (
                <>
                  <Sparkles size={16} />
                  <span>{isVi ? 'Chẩn Đoán NETA MTS AI' : 'Run NETA MTS AI Analysis'}</span>
                </>
              )}
            </button>

            <button
              onClick={handleExportPdf}
              className="flex items-center gap-1.5 px-3 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 text-sm font-medium transition"
            >
              <Download size={15} />
              <span>PDF</span>
            </button>

            <button
              onClick={handleSaveReport}
              className="flex items-center gap-1.5 px-3 py-2 rounded-xl bg-emerald-600/90 hover:bg-emerald-500 text-white text-sm font-medium transition shadow-sm"
            >
              <Save size={15} />
              <span>{saveSuccess ? (isVi ? 'Đã Lưu!' : 'Saved!') : (isVi ? 'Lưu Biên Bản' : 'Save Report')}</span>
            </button>
          </div>
        </div>

        {/* Preset Toolbar */}
        <div className="mt-5 pt-4 border-t border-slate-800 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex items-center gap-2 text-slate-300 font-medium">
            <Sliders size={14} className="text-sky-400" />
            <span>{isVi ? 'Kịch bản mẫu kiểm thử nhanh (Field Scenarios):' : 'Quick Field Scenarios:'}</span>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => handleLoadPreset('defect')}
              className="flex items-center gap-1.5 px-2.5 py-1 rounded-md bg-rose-950/80 hover:bg-rose-900 text-rose-300 border border-rose-800/80 transition"
            >
              <AlertOctagon size={13} className="text-rose-400" />
              <span>{isVi ? 'Sự cố: Suy Giảm Cách Điện Pha B & Phóng Điện 21kV (Mẫu Prompt)' : 'Defect: Phase B Insulation Breakdown & Tracking'}</span>
            </button>

            <button
              onClick={() => handleLoadPreset('pass')}
              className="flex items-center gap-1.5 px-2.5 py-1 rounded-md bg-emerald-950/80 hover:bg-emerald-900 text-emerald-300 border border-emerald-800/80 transition"
            >
              <CheckCircle2 size={13} className="text-emerald-400" />
              <span>{isVi ? 'Đạt Chuẩn: Bộ VT 110kV CVT Đạt 100% (Pristine Pass)' : 'Pristine: 110kV CVT 100% Compliant'}</span>
            </button>

            <button
              onClick={() => handleLoadPreset('caution')}
              className="flex items-center gap-1.5 px-2.5 py-1 rounded-md bg-amber-950/80 hover:bg-amber-900 text-amber-300 border border-amber-800/80 transition"
            >
              <AlertTriangle size={13} className="text-amber-400" />
              <span>{isVi ? 'Cảnh Báo: Quá Tải Burden Pha C & Cầu Chì Oxy Hóa' : 'Caution: Burden Overload & Fuse Defect'}</span>
            </button>
          </div>
        </div>
      </div>

      {/* Main Tab Bar */}
      <div className="flex border-b border-slate-200 dark:border-slate-800 space-x-2 text-sm font-medium">
        <button
          onClick={() => setActiveTab('checklist')}
          className={`flex items-center gap-2 py-3 px-4 border-b-2 font-semibold transition-colors ${
            activeTab === 'checklist'
              ? 'border-sky-600 text-sky-600 dark:text-sky-400'
              : 'border-transparent text-slate-500 hover:text-slate-800 dark:text-slate-400 dark:hover:text-slate-200'
          }`}
        >
          <Layers size={16} />
          <span>{isVi ? '1. Biên Bản Đo Đạc Site (Checklist NETA MTS 7.10.2)' : '1. Field Test Checklist'}</span>
        </button>

        <button
          onClick={() => setActiveTab('evaluation')}
          className={`flex items-center gap-2 py-3 px-4 border-b-2 font-semibold transition-colors relative ${
            activeTab === 'evaluation'
              ? 'border-sky-600 text-sky-600 dark:text-sky-400'
              : 'border-transparent text-slate-500 hover:text-slate-800 dark:text-slate-400 dark:hover:text-slate-200'
          }`}
        >
          <Sparkles size={16} />
          <span>{isVi ? '2. Chẩn Đoán AI & Đánh Giá Kỹ Thuật' : '2. AI Diagnostic & Evaluation'}</span>
          {evaluationResult && (
            <span className={`w-2 h-2 rounded-full ${evaluationResult.overall_status === 'CRITICAL' ? 'bg-rose-500' : evaluationResult.overall_status === 'CAUTION' ? 'bg-amber-500' : 'bg-emerald-500'}`} />
          )}
        </button>

        <button
          onClick={() => setActiveTab('prompt')}
          className={`flex items-center gap-2 py-3 px-4 border-b-2 font-semibold transition-colors ${
            activeTab === 'prompt'
              ? 'border-sky-600 text-sky-600 dark:text-sky-400'
              : 'border-transparent text-slate-500 hover:text-slate-800 dark:text-slate-400 dark:hover:text-slate-200'
          }`}
        >
          <FileText size={16} />
          <span>{isVi ? '3. System Prompt Google AI Studio' : '3. Google AI Studio System Prompt'}</span>
        </button>

        <button
          onClick={() => setActiveTab('json')}
          className={`flex items-center gap-2 py-3 px-4 border-b-2 font-semibold transition-colors ${
            activeTab === 'json'
              ? 'border-sky-600 text-sky-600 dark:text-sky-400'
              : 'border-transparent text-slate-500 hover:text-slate-800 dark:text-slate-400 dark:hover:text-slate-200'
          }`}
        >
          <FileCheck2 size={16} />
          <span>{isVi ? '4. Mẫu JSON Input' : '4. JSON Input Schema'}</span>
        </button>
      </div>

      {/* TAB 1: FIELD TEST CHECKLIST */}
      {activeTab === 'checklist' && (
        <div className="space-y-6">
          {/* Header Metadata Card */}
          <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm">
            <h3 className="text-base font-bold text-slate-800 dark:text-slate-100 mb-4 flex items-center gap-2">
              <Building size={18} className="text-sky-600 dark:text-sky-400" />
              <span>{isVi ? 'Thông Tin Dự Án & Bộ Biến Điện Áp VT/PT' : 'Project & Instrument VT Information'}</span>
            </h3>

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 text-sm">
              <div>
                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">
                  {isVi ? 'Tên Dự Án / Nhà Máy' : 'Project / Plant Name'}
                </label>
                <input
                  type="text"
                  value={formData.tev_coe_header.project_name}
                  onChange={(e) => setFormData({
                    ...formData,
                    tev_coe_header: { ...formData.tev_coe_header, project_name: e.target.value }
                  })}
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-100 text-sm focus:ring-2 focus:ring-sky-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">
                  {isVi ? 'Mã Nhãn Thiết Bị (Tag)' : 'Equipment Tag'}
                </label>
                <input
                  type="text"
                  value={formData.tev_coe_header.equipment_tag}
                  onChange={(e) => setFormData({
                    ...formData,
                    tev_coe_header: { ...formData.tev_coe_header, equipment_tag: e.target.value }
                  })}
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-100 text-sm focus:ring-2 focus:ring-sky-500 focus:outline-none font-semibold text-sky-600 dark:text-sky-400"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">
                  {isVi ? 'Tỉ Số Biến Điện Áp (VT Ratio)' : 'VT Ratio Rating'}
                </label>
                <input
                  type="text"
                  value={formData.tev_coe_header.vt_ratio_rating}
                  onChange={(e) => setFormData({
                    ...formData,
                    tev_coe_header: { ...formData.tev_coe_header, vt_ratio_rating: e.target.value }
                  })}
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-100 text-sm focus:ring-2 focus:ring-sky-500 focus:outline-none font-mono"
                  placeholder="e.g. 22000/sqrt(3) : 110/sqrt(3) V"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">
                  {isVi ? 'Cấp Chính Xác & Dung Lượng (Accuracy & VA)' : 'Accuracy Class & Rated VA'}
                </label>
                <input
                  type="text"
                  value={formData.tev_coe_header.accuracy_class_rating}
                  onChange={(e) => setFormData({
                    ...formData,
                    tev_coe_header: { ...formData.tev_coe_header, accuracy_class_rating: e.target.value }
                  })}
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-100 text-sm focus:ring-2 focus:ring-sky-500 focus:outline-none font-mono"
                  placeholder="e.g. 0.2 / 3P 50VA"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">
                  {isVi ? 'Điện Áp Định Mức Hệ Thống (kV)' : 'System Voltage (kV)'}
                </label>
                <input
                  type="number"
                  value={formData.tev_coe_header.voltage_rating_kV}
                  onChange={(e) => setFormData({
                    ...formData,
                    tev_coe_header: { ...formData.tev_coe_header, voltage_rating_kV: parseFloat(e.target.value) || 22 }
                  })}
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-100 text-sm focus:ring-2 focus:ring-sky-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">
                  {isVi ? 'Kỹ Sư Thử Nghiệm (FSE)' : 'Field Service Engineer'}
                </label>
                <input
                  type="text"
                  value={formData.tev_coe_header.fse_name}
                  onChange={(e) => setFormData({
                    ...formData,
                    tev_coe_header: { ...formData.tev_coe_header, fse_name: e.target.value }
                  })}
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-100 text-sm focus:ring-2 focus:ring-sky-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">
                  {isVi ? 'Ngày Thử Nghiệm (Test Date)' : 'Test Date'}
                </label>
                <input
                  type="date"
                  value={formData.tev_coe_header.test_date}
                  onChange={(e) => setFormData({
                    ...formData,
                    tev_coe_header: { ...formData.tev_coe_header, test_date: e.target.value }
                  })}
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-100 text-sm focus:ring-2 focus:ring-sky-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">
                  {isVi ? 'Vị Trí Ngăn Lộ / Tủ Điện' : 'Bay / Cubicle Location'}
                </label>
                <input
                  type="text"
                  value={formData.tev_coe_header.bay_location || ''}
                  onChange={(e) => setFormData({
                    ...formData,
                    tev_coe_header: { ...formData.tev_coe_header, bay_location: e.target.value }
                  })}
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-100 text-sm focus:ring-2 focus:ring-sky-500 focus:outline-none"
                />
              </div>
            </div>
          </div>

          {/* Phase Navigation */}
          <div className="flex border-b border-slate-200 dark:border-slate-800 space-x-4 text-xs font-semibold">
            <button
              onClick={() => setActivePhase('phase1')}
              className={`pb-2 transition border-b-2 ${activePhase === 'phase1' ? 'border-sky-600 text-sky-600 dark:text-sky-400' : 'border-transparent text-slate-500 hover:text-slate-800'}`}
            >
              {isVi ? 'Giai Đoạn 1: Kiểm Tra Thị Giác & Cầu Chì (NETA MTS 7.10.2.A)' : 'Phase 1: Visual & Mechanical (7.10.2.A)'}
            </button>
            <button
              onClick={() => setActivePhase('phase2')}
              className={`pb-2 transition border-b-2 ${activePhase === 'phase2' ? 'border-sky-600 text-sky-600 dark:text-sky-400' : 'border-transparent text-slate-500 hover:text-slate-800'}`}
            >
              {isVi ? 'Giai Đoạn 2: Thử Nghiệm Điện & Tỉ Số Biến (NETA MTS 7.10.2.B)' : 'Phase 2: Electrical Tests (7.10.2.B)'}
            </button>
            <button
              onClick={() => setActivePhase('phase3')}
              className={`pb-2 transition border-b-2 ${activePhase === 'phase3' ? 'border-sky-600 text-sky-600 dark:text-sky-400' : 'border-transparent text-slate-500 hover:text-slate-800'}`}
            >
              {isVi ? 'Giai Đoạn 3: Khảo Sát Hồng Ngoại & Sau Thử Nghiệm' : 'Phase 3: Post-Test Survey & Thermography'}
            </button>
          </div>

          {/* PHASE 1: VISUAL & MECHANICAL */}
          {activePhase === 'phase1' && (
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm space-y-6">
              <div className="flex items-center justify-between border-b border-slate-200 dark:border-slate-800 pb-3">
                <div className="flex items-center gap-2">
                  <div className="w-8 h-8 rounded-lg bg-sky-100 dark:bg-sky-900/40 text-sky-600 dark:text-sky-400 flex items-center justify-center font-bold">
                    A
                  </div>
                  <div>
                    <h4 className="font-bold text-slate-900 dark:text-slate-100 text-sm">
                      {isVi ? 'Kiểm Tra Thị Giác, Cầu Chì & Cơ Khí (NETA MTS Mục 7.10.2.A)' : 'Visual & Mechanical Inspection per NETA MTS Sec 7.10.2.A'}
                    </h4>
                    <p className="text-xs text-slate-500">
                      {isVi ? 'Đối soát thông số nhãn mác, tình trạng cơ lý, cầu chì sơ cấp, mức dầu/khí SF6 và điểm tiếp địa mạch nhị thứ.' : 'Verify nameplate, physical condition, primary fuses, fluid/gas levels, and secondary grounding.'}
                    </p>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-5 text-sm">
                {/* A2: Physical Condition */}
                <div className="p-4 rounded-lg bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-semibold text-slate-800 dark:text-slate-200">
                      A2. {isVi ? 'Kiểm Tra Cơ Lý & Vỏ Đúc Cách Điện' : 'Physical Condition & Insulation Integrity'}
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-2">
                    {isVi ? 'Kiểm tra nứt vỡ cách điện, rò rỉ keo đúc/dầu, vết phóng điện bề mặt và vệ sinh sứ sạch sẽ.' : 'Inspect for physical damage, cracked resin/porcelain, fluid leakage, and clean surfaces.'}
                  </p>
                  <input
                    type="text"
                    value={formData.visual_and_mechanical_inspection.physical_condition}
                    onChange={(e) => setFormData({
                      ...formData,
                      visual_and_mechanical_inspection: {
                        ...formData.visual_and_mechanical_inspection,
                        physical_condition: e.target.value
                      }
                    })}
                    className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-xs"
                  />
                </div>

                {/* A3: Primary Fuses */}
                <div className="p-4 rounded-lg bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-semibold text-slate-800 dark:text-slate-200">
                      A3. {isVi ? 'Kiểm Tra Cầu Chì Sơ Cấp (7.10.2.A.4)' : 'Primary & Secondary Fuses (7.10.2.A.4)'}
                    </span>
                    <span className="text-xs px-2 py-0.5 rounded bg-blue-100 text-blue-800 dark:bg-blue-900/40 dark:text-blue-300">
                      Current-Limiting
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-2">
                    {isVi ? 'Kiểm tra thông mạch, loại cầu chì hạn dòng (Current-limiting), tiếp xúc ngàm kẹp và aptomat MCB thứ cấp.' : 'Verify primary fuse continuity, current-limiting rating, holder contacts, and secondary MCB.'}
                  </p>
                  <input
                    type="text"
                    value={formData.visual_and_mechanical_inspection.primary_fuse_status}
                    onChange={(e) => setFormData({
                      ...formData,
                      visual_and_mechanical_inspection: {
                        ...formData.visual_and_mechanical_inspection,
                        primary_fuse_status: e.target.value
                      }
                    })}
                    className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-xs"
                  />
                </div>

                {/* A4: Bolted Connections */}
                <div className="p-4 rounded-lg bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-semibold text-slate-800 dark:text-slate-200">
                      A4. {isVi ? 'Lực Siết Bu-lông Kết Nối (Bảng 100.12)' : 'Bolted Electrical Connections (Table 100.12)'}
                    </span>
                    <span className="text-xs px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 dark:bg-emerald-900/40 dark:text-emerald-300">
                      Table 100.12
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-2">
                    {isVi ? 'Kiểm tra lực siết bu-lông mạch sơ & thứ cấp bằng cờ-lê lực hoặc đo điện trở tiếp xúc DLRO (< 50% lệch).' : 'Verify bolted connection tightness per Table 100.12 or low-resistance ohmmeter.'}
                  </p>
                  <input
                    type="text"
                    value={formData.visual_and_mechanical_inspection.bolted_connection_torque_table_100_12}
                    onChange={(e) => setFormData({
                      ...formData,
                      visual_and_mechanical_inspection: {
                        ...formData.visual_and_mechanical_inspection,
                        bolted_connection_torque_table_100_12: e.target.value
                      }
                    })}
                    className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-xs"
                  />
                </div>

                {/* A5: Grounding Point Check */}
                <div className="p-4 rounded-lg bg-gradient-to-br from-sky-50 to-slate-50 dark:from-sky-950/30 dark:to-slate-800/60 border-2 border-sky-400/50 shadow-sm">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-bold text-slate-800 dark:text-slate-200 flex items-center gap-1.5">
                      <ShieldAlert className="text-sky-600 dark:text-sky-400" size={17} />
                      <span>A5. {isVi ? 'Điểm Nối Đất Mạch Thứ Cấp (IEEE C57.13.3)' : 'Secondary Grounding (IEEE C57.13.3)'}</span>
                    </span>
                    <span className="text-xs px-2 py-0.5 rounded bg-sky-600 text-white font-bold">
                      Duy nhất 1 điểm
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-2">
                    {isVi ? 'Bắt buộc chỉ có duy nhất 1 điểm tiếp địa trên toàn bộ vòng lặp mạch áp (thường tại Tủ Rơ-le).' : 'MANDATORY exactly one grounding point on secondary voltage loop.'}
                  </p>

                  <div className="flex gap-2">
                    <select
                      value={formData.visual_and_mechanical_inspection.secondary_grounding_point_ieee_c57_13_3.ground_points_detected}
                      onChange={(e) => {
                        const pts = parseInt(e.target.value) || 1;
                        setFormData({
                          ...formData,
                          visual_and_mechanical_inspection: {
                            ...formData.visual_and_mechanical_inspection,
                            secondary_grounding_point_ieee_c57_13_3: {
                              ...formData.visual_and_mechanical_inspection.secondary_grounding_point_ieee_c57_13_3,
                              ground_points_detected: pts,
                              status: pts === 1 ? 'Pass' : `Fail (${pts} points detected)`
                            }
                          }
                        });
                      }}
                      className={`w-1/3 px-2 py-1.5 rounded border text-xs font-bold ${
                        formData.visual_and_mechanical_inspection.secondary_grounding_point_ieee_c57_13_3.ground_points_detected === 1
                          ? 'border-emerald-500 text-emerald-700 bg-emerald-50'
                          : 'border-rose-500 text-rose-700 bg-rose-50 animate-pulse'
                      }`}
                    >
                      <option value={1}>1 Điểm Đất (Pass)</option>
                      <option value={2}>2 Điểm (Fail)</option>
                      <option value={0}>0 Điểm (Fail)</option>
                    </select>
                    <input
                      type="text"
                      value={formData.visual_and_mechanical_inspection.secondary_grounding_point_ieee_c57_13_3.ground_location}
                      onChange={(e) => setFormData({
                        ...formData,
                        visual_and_mechanical_inspection: {
                          ...formData.visual_and_mechanical_inspection,
                          secondary_grounding_point_ieee_c57_13_3: {
                            ...formData.visual_and_mechanical_inspection.secondary_grounding_point_ieee_c57_13_3,
                            ground_location: e.target.value
                          }
                        }
                      })}
                      className="w-2/3 px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-xs font-mono"
                    />
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* PHASE 2: ELECTRICAL TESTS */}
          {activePhase === 'phase2' && (
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm space-y-6">
              <div className="flex items-center justify-between border-b border-slate-200 dark:border-slate-800 pb-3">
                <div className="flex items-center gap-2">
                  <div className="w-8 h-8 rounded-lg bg-sky-100 dark:bg-sky-900/40 text-sky-600 dark:text-sky-400 flex items-center justify-center font-bold">
                    B
                  </div>
                  <div>
                    <h4 className="font-bold text-slate-900 dark:text-slate-100 text-sm">
                      {isVi ? 'Thử Nghiệm Điện & Đo Đạc Thông Số (NETA MTS Mục 7.10.2.B)' : 'Electrical Tests per ANSI/NETA MTS-2023 Sec 7.10.2.B'}
                    </h4>
                    <p className="text-xs text-slate-500">
                      {isVi ? 'Cách điện sơ & thứ cấp Bảng 100.5, cực tính H1-X1, thử tỉ số biến áp 3 pha, đo phụ tải Burden và chịu đựng điện áp cao Bảng 100.9.' : 'Insulation Table 100.5, polarity, turns-ratio verification, burden, and dielectric withstand Table 100.9.'}
                    </p>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5 text-sm">
                {/* B1: Primary Insulation Resistance 3 Phases */}
                <div className="p-4 rounded-xl bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60 lg:col-span-3">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-bold text-slate-800 dark:text-slate-200">
                      B1. {isVi ? 'Điện Trở Cách Điện Cuộn Sơ Cấp 3 Pha & Thứ Cấp (Bảng 100.5)' : 'Insulation Resistance Primary (3 Phases) & Secondary (Table 100.5)'}
                    </span>
                    <span className="text-xs font-mono text-sky-600 dark:text-sky-400">
                      Min ≥ 1,000 MΩ (Sơ cấp) / ≥ 100 MΩ (Thứ cấp)
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-3">
                    {isVi ? 'Đo điện trở cách điện 1 phút ở 2500V DC cho cuộn sơ cấp (Pha A, B, C) so với thứ cấp + đất, và 1000V DC cho cuộn thứ cấp.' : 'Measure 1-minute insulation resistance at 2500V DC (Primary) and 1000V DC (Secondary).'}
                  </p>

                  <div className="grid grid-cols-1 md:grid-cols-4 gap-4 text-xs">
                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Pha A - Sơ cấp (MΩ, 2500V):' : 'Phase A Primary (MΩ):'}
                      </label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_2_b.insulation_resistance_megohms_table_100_5.primary_to_ground_2500v_dc_Mohm.phase_a}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_2_b: {
                            ...formData.electrical_tests_7_10_2_b,
                            insulation_resistance_megohms_table_100_5: {
                              ...formData.electrical_tests_7_10_2_b.insulation_resistance_megohms_table_100_5,
                              primary_to_ground_2500v_dc_Mohm: {
                                ...formData.electrical_tests_7_10_2_b.insulation_resistance_megohms_table_100_5.primary_to_ground_2500v_dc_Mohm,
                                phase_a: parseFloat(e.target.value) || 0
                              }
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Pha B - Sơ cấp (MΩ, 2500V):' : 'Phase B Primary (MΩ):'}
                      </label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_2_b.insulation_resistance_megohms_table_100_5.primary_to_ground_2500v_dc_Mohm.phase_b}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_2_b: {
                            ...formData.electrical_tests_7_10_2_b,
                            insulation_resistance_megohms_table_100_5: {
                              ...formData.electrical_tests_7_10_2_b.insulation_resistance_megohms_table_100_5,
                              primary_to_ground_2500v_dc_Mohm: {
                                ...formData.electrical_tests_7_10_2_b.insulation_resistance_megohms_table_100_5.primary_to_ground_2500v_dc_Mohm,
                                phase_b: parseFloat(e.target.value) || 0
                              }
                            }
                          }
                        })}
                        className={`w-full px-3 py-1.5 rounded border font-mono font-bold ${
                          formData.electrical_tests_7_10_2_b.insulation_resistance_megohms_table_100_5.primary_to_ground_2500v_dc_Mohm.phase_b >= 1000
                            ? 'border-emerald-500 text-emerald-700 bg-emerald-50'
                            : 'border-rose-500 text-rose-700 bg-rose-50 animate-pulse'
                        }`}
                      />
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Pha C - Sơ cấp (MΩ, 2500V):' : 'Phase C Primary (MΩ):'}
                      </label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_2_b.insulation_resistance_megohms_table_100_5.primary_to_ground_2500v_dc_Mohm.phase_c}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_2_b: {
                            ...formData.electrical_tests_7_10_2_b,
                            insulation_resistance_megohms_table_100_5: {
                              ...formData.electrical_tests_7_10_2_b.insulation_resistance_megohms_table_100_5,
                              primary_to_ground_2500v_dc_Mohm: {
                                ...formData.electrical_tests_7_10_2_b.insulation_resistance_megohms_table_100_5.primary_to_ground_2500v_dc_Mohm,
                                phase_c: parseFloat(e.target.value) || 0
                              }
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Thứ cấp - Đất (MΩ, 1000V):' : 'Secondary to Ground (MΩ):'}
                      </label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_2_b.insulation_resistance_megohms_table_100_5.secondary_to_ground_1000v_dc_Mohm}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_2_b: {
                            ...formData.electrical_tests_7_10_2_b,
                            insulation_resistance_megohms_table_100_5: {
                              ...formData.electrical_tests_7_10_2_b.insulation_resistance_megohms_table_100_5,
                              secondary_to_ground_1000v_dc_Mohm: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>
                  </div>
                </div>

                {/* B2: Polarity Test */}
                <div className="p-4 rounded-xl bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-bold text-slate-800 dark:text-slate-200">
                      B2. {isVi ? 'Thử Cực Tính Cuộn Dây (H1-X1)' : 'Polarity Test (IEEE C57.13.1)'}
                    </span>
                    <span className="text-xs px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 dark:bg-emerald-900/40 dark:text-emerald-300">
                      H1-X1
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-3">
                    {isVi ? 'Xác nhận cực tính H1-X1 cuộn sơ/thứ cấp bằng xung DC hoặc AC.' : 'Verify polarity agreement with H1-X1 markings.'}
                  </p>
                  <select
                    value={formData.electrical_tests_7_10_2_b.polarity_test_ieee_c57_13_1}
                    onChange={(e) => setFormData({
                      ...formData,
                      electrical_tests_7_10_2_b: {
                        ...formData.electrical_tests_7_10_2_b,
                        polarity_test_ieee_c57_13_1: e.target.value
                      }
                    })}
                    className="w-full px-3 py-2 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-xs font-semibold"
                  >
                    <option value="Pass (H1-X1 correct polarity verified on all 3 phases)">Pass - Đúng cực tính H1-X1 cả 3 pha</option>
                    <option value="Fail (Reversed polarity detected on Phase B)">Fail - Ngược cực tính Pha B (Nguy hiểm bảo vệ hướng)</option>
                    <option value="Fail (Reversed polarity detected on Phase A)">Fail - Ngược cực tính Pha A</option>
                  </select>
                </div>

                {/* B4: Burden Measurement */}
                <div className="p-4 rounded-xl bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60 lg:col-span-2">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-bold text-slate-800 dark:text-slate-200">
                      B4. {isVi ? 'Đo Phụ Tải Mạch Áp Thứ Cấp (Burden VA)' : 'Secondary Voltage Burden (7.10.2.D.5)'}
                    </span>
                    <span className="text-xs font-mono text-sky-600 dark:text-sky-400">
                      Burden ≤ {formData.electrical_tests_7_10_2_b.secondary_burden_measurement.vt_rated_burden_VA} VA
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-3">
                    {isVi ? 'Đo phụ tải thực tế (VA) của các rơ-le và đồng hồ đo. Không được vượt công suất định mức của VT.' : 'Measure actual circuit burden (VA). Must not exceed rated VA.'}
                  </p>

                  <div className="grid grid-cols-3 gap-3 text-xs">
                    <div>
                      <label className="block text-slate-500 mb-1">{isVi ? 'Pha A (VA):' : 'Phase A (VA):'}</label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_2_b.secondary_burden_measurement.measured_burden_phase_a_VA}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_2_b: {
                            ...formData.electrical_tests_7_10_2_b,
                            secondary_burden_measurement: {
                              ...formData.electrical_tests_7_10_2_b.secondary_burden_measurement,
                              measured_burden_phase_a_VA: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-2 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-500 mb-1">{isVi ? 'Pha B (VA):' : 'Phase B (VA):'}</label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_2_b.secondary_burden_measurement.measured_burden_phase_b_VA}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_2_b: {
                            ...formData.electrical_tests_7_10_2_b,
                            secondary_burden_measurement: {
                              ...formData.electrical_tests_7_10_2_b.secondary_burden_measurement,
                              measured_burden_phase_b_VA: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-2 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-500 mb-1">{isVi ? 'Pha C (VA):' : 'Phase C (VA):'}</label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_2_b.secondary_burden_measurement.measured_burden_phase_c_VA}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_2_b: {
                            ...formData.electrical_tests_7_10_2_b,
                            secondary_burden_measurement: {
                              ...formData.electrical_tests_7_10_2_b.secondary_burden_measurement,
                              measured_burden_phase_c_VA: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-2 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>
                  </div>
                </div>

                {/* B3: Turns-Ratio Verification Test */}
                <div className="p-4 rounded-xl bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60 lg:col-span-3">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-bold text-slate-800 dark:text-slate-200">
                      B3. {isVi ? 'Thử Tỉ Số Biến Điện Áp (Turns-Ratio per IEEE C57.13)' : 'Turns-Ratio Test (IEEE C57.13)'}
                    </span>
                    <span className="text-xs px-2 py-0.5 rounded bg-blue-100 text-blue-800 dark:bg-blue-900/40 dark:text-blue-300">
                      Cấp 0.2 / Sai số ≤ ±0.2%
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-3">
                    {isVi ? 'Đặt điện áp sơ cấp thử nghiệm và đo điện áp thứ cấp các pha A, B, C. Sai số tỉ số phải nằm trong cấp chính xác định mức (0.2%).' : 'Apply test primary voltage and measure secondary output on all phases. Ratio error must be within class 0.2.'}
                  </p>

                  <div className="grid grid-cols-1 md:grid-cols-4 gap-4 text-xs">
                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Áp sơ cấp đặt vào (V):' : 'Applied Primary (V):'}
                      </label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_2_b.turns_ratio_verification_test_ieee_c57_13.applied_primary_voltage_V}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_2_b: {
                            ...formData.electrical_tests_7_10_2_b,
                            turns_ratio_verification_test_ieee_c57_13: {
                              ...formData.electrical_tests_7_10_2_b.turns_ratio_verification_test_ieee_c57_13,
                              applied_primary_voltage_V: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Áp thứ cấp Pha A (V):' : 'Sec Voltage Phase A (V):'}
                      </label>
                      <input
                        type="number"
                        step="0.01"
                        value={formData.electrical_tests_7_10_2_b.turns_ratio_verification_test_ieee_c57_13.measured_secondary_voltage_phase_a_V}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_2_b: {
                            ...formData.electrical_tests_7_10_2_b,
                            turns_ratio_verification_test_ieee_c57_13: {
                              ...formData.electrical_tests_7_10_2_b.turns_ratio_verification_test_ieee_c57_13,
                              measured_secondary_voltage_phase_a_V: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Áp thứ cấp Pha B (V) & Sai số %:' : 'Sec Phase B (V) & Error %:'}
                      </label>
                      <div className="flex gap-2">
                        <input
                          type="number"
                          step="0.01"
                          value={formData.electrical_tests_7_10_2_b.turns_ratio_verification_test_ieee_c57_13.measured_secondary_voltage_phase_b_V}
                          onChange={(e) => {
                            const val = parseFloat(e.target.value) || 0;
                            const nominalSec = 63.51; // 110/sqrt(3)
                            const err = ((val - nominalSec) / nominalSec) * 100;
                            setFormData({
                              ...formData,
                              electrical_tests_7_10_2_b: {
                                ...formData.electrical_tests_7_10_2_b,
                                turns_ratio_verification_test_ieee_c57_13: {
                                  ...formData.electrical_tests_7_10_2_b.turns_ratio_verification_test_ieee_c57_13,
                                  measured_secondary_voltage_phase_b_V: val,
                                  ratio_error_phase_b_percent: parseFloat(err.toFixed(2))
                                }
                              }
                            });
                          }}
                          className="w-1/2 px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                        />
                        <div className={`w-1/2 px-2 py-1.5 rounded text-center font-bold font-mono ${
                          Math.abs(formData.electrical_tests_7_10_2_b.turns_ratio_verification_test_ieee_c57_13.ratio_error_phase_b_percent) <= 0.2
                            ? 'bg-emerald-100 text-emerald-800 dark:bg-emerald-950/60 dark:text-emerald-300'
                            : 'bg-rose-100 text-rose-800 dark:bg-rose-950/60 dark:text-rose-300'
                        }`}>
                          {formData.electrical_tests_7_10_2_b.turns_ratio_verification_test_ieee_c57_13.ratio_error_phase_b_percent}%
                        </div>
                      </div>
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Áp thứ cấp Pha C (V):' : 'Sec Voltage Phase C (V):'}
                      </label>
                      <input
                        type="number"
                        step="0.01"
                        value={formData.electrical_tests_7_10_2_b.turns_ratio_verification_test_ieee_c57_13.measured_secondary_voltage_phase_c_V}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_2_b: {
                            ...formData.electrical_tests_7_10_2_b,
                            turns_ratio_verification_test_ieee_c57_13: {
                              ...formData.electrical_tests_7_10_2_b.turns_ratio_verification_test_ieee_c57_13,
                              measured_secondary_voltage_phase_c_V: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>
                  </div>
                </div>

                {/* B5: Dielectric Withstand Test */}
                <div className="p-4 rounded-xl bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60 lg:col-span-3">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-bold text-slate-800 dark:text-slate-200">
                      B5. {isVi ? 'Thử Chịu Đựng Điện Áp Tăng Cao (Dielectric Withstand Table 100.9)' : 'Dielectric Withstand Test (Table 100.9)'}
                    </span>
                    <span className="text-xs font-mono text-sky-600 dark:text-sky-400">
                      Table 100.9 (32.5kV AC 1 min)
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-3">
                    {isVi ? 'Áp dụng điện áp thử tần số công nghiệp lên cuộn sơ cấp trong 1 phút (ví dụ hệ 22kV thử 32.5kV AC). Không được có hiện tượng đánh thủng cách điện.' : 'Apply rated test voltage for 1 minute to primary winding without breakdown.'}
                  </p>

                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Điện áp thử nghiệm (kV AC):' : 'Test Voltage (kV AC):'}
                      </label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_2_b.dielectric_withstand_test_table_100_9.test_voltage_kV_ac}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_2_b: {
                            ...formData.electrical_tests_7_10_2_b,
                            dielectric_withstand_test_table_100_9: {
                              ...formData.electrical_tests_7_10_2_b.dielectric_withstand_test_table_100_9,
                              test_voltage_kV_ac: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Thời gian duy trì (Phút):' : 'Duration (Minutes):'}
                      </label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_2_b.dielectric_withstand_test_table_100_9.duration_minutes}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_2_b: {
                            ...formData.electrical_tests_7_10_2_b,
                            dielectric_withstand_test_table_100_9: {
                              ...formData.electrical_tests_7_10_2_b.dielectric_withstand_test_table_100_9,
                              duration_minutes: parseInt(e.target.value) || 1
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Hiện trạng Pha B:' : 'Phase B Withstand Status:'}
                      </label>
                      <input
                        type="text"
                        value={formData.electrical_tests_7_10_2_b.dielectric_withstand_test_table_100_9.phase_b_status}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_2_b: {
                            ...formData.electrical_tests_7_10_2_b,
                            dielectric_withstand_test_table_100_9: {
                              ...formData.electrical_tests_7_10_2_b.dielectric_withstand_test_table_100_9,
                              phase_b_status: e.target.value
                            }
                          }
                        })}
                        className={`w-full px-3 py-1.5 rounded border text-xs font-semibold ${
                          formData.electrical_tests_7_10_2_b.dielectric_withstand_test_table_100_9.phase_b_status.toLowerCase().includes('pass')
                            ? 'border-emerald-500 bg-emerald-50 text-emerald-800'
                            : 'border-rose-500 bg-rose-50 text-rose-800 animate-pulse'
                        }`}
                      />
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* PHASE 3: POST-TEST & THERMOGRAPHY */}
          {activePhase === 'phase3' && (
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm space-y-6">
              <div className="flex items-center justify-between border-b border-slate-200 dark:border-slate-800 pb-3">
                <div className="flex items-center gap-2">
                  <div className="w-8 h-8 rounded-lg bg-sky-100 dark:bg-sky-900/40 text-sky-600 dark:text-sky-400 flex items-center justify-center font-bold">
                    C
                  </div>
                  <div>
                    <h4 className="font-bold text-slate-900 dark:text-slate-100 text-sm">
                      {isVi ? 'Khảo Sát Hồng Ngoại & Hoàn Tất Chạy Thử (Section 9 / Table 100.18)' : 'Post-Test Survey & Thermography (Table 100.18)'}
                    </h4>
                    <p className="text-xs text-slate-500">
                      {isVi ? 'Khảo sát ảnh nhiệt IR dưới tải và kiểm tra kết nối hoàn tất bảo trì.' : 'Infrared thermographic survey under load and final sign-off.'}
                    </p>
                  </div>
                </div>
              </div>

              <div className="p-4 rounded-xl bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60">
                <div className="flex items-center justify-between mb-2">
                  <span className="font-bold text-slate-800 dark:text-slate-200 flex items-center gap-2">
                    <Thermometer size={16} className="text-rose-500" />
                    <span>{isVi ? 'Khảo Sát Nhiệt IR Mối Nối & Thân VT (NETA Section 9 / Table 100.18)' : 'Thermographic Survey (Table 100.18)'}</span>
                  </span>
                </div>
                <p className="text-xs text-slate-500 mb-3">
                  {isVi ? 'Kiểm tra phát hiện điểm quá nhiệt bất thường (Hotspot) tại đế cầu chì sơ cấp, đầu cốt đấu nối và thân nhựa đúc/sứ.' : 'Check for abnormal hotspots at primary fuse clips, terminal blocks, and insulation body.'}
                </p>
                <input
                  type="text"
                  value={formData.post_test_survey.thermographic_survey_table_100_18}
                  onChange={(e) => setFormData({
                    ...formData,
                    post_test_survey: {
                      ...formData.post_test_survey,
                      thermographic_survey_table_100_18: e.target.value
                    }
                  })}
                  className="w-full px-3 py-2 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-xs font-semibold"
                />
              </div>

              <div className="pt-4 flex justify-end">
                <button
                  onClick={handleRunEvaluation}
                  disabled={isEvaluating}
                  className="flex items-center gap-2 px-5 py-2.5 rounded-xl bg-sky-600 hover:bg-sky-700 text-white font-bold text-sm shadow-md transition"
                >
                  <Sparkles size={16} />
                  <span>{isVi ? 'Tiến Hành Chẩn Đoán AI & Xuất Đánh Giá' : 'Run AI Diagnostic Evaluation'}</span>
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {/* TAB 2: AI DIAGNOSTIC EVALUATION */}
      {activeTab === 'evaluation' && (
        <div className="space-y-6">
          {isEvaluating ? (
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-12 text-center shadow-sm">
              <div className="w-12 h-12 border-4 border-sky-600 border-t-transparent rounded-full animate-spin mx-auto mb-4" />
              <h3 className="text-lg font-bold text-slate-900 dark:text-slate-100">
                {isVi ? 'Đang Phân Tích Dữ Liệu Theo ANSI/NETA MTS-2023 Mục 7.10.2...' : 'Analyzing VT Data per ANSI/NETA MTS-2023 Sec 7.10.2...'}
              </h3>
              <p className="text-sm text-slate-500 mt-1">
                {isVi ? 'Đối soát Bảng 100.5, 100.9, 100.12, 100.18, IEEE C57.13 và IEEE C57.13.3' : 'Cross-referencing Tables 100.5, 100.9, 100.12, 100.18, IEEE C57.13 & C57.13.3'}
              </p>
            </div>
          ) : evaluationResult ? (
            <div className="space-y-6">
              {/* Overall Severity Banner */}
              <div className={`rounded-xl p-5 border text-white shadow-lg ${
                evaluationResult.overall_status === 'CRITICAL'
                  ? 'bg-gradient-to-r from-rose-900 via-rose-800 to-slate-900 border-rose-700'
                  : evaluationResult.overall_status === 'CAUTION'
                  ? 'bg-gradient-to-r from-amber-900 via-amber-800 to-slate-900 border-amber-700'
                  : 'bg-gradient-to-r from-emerald-900 via-emerald-800 to-slate-900 border-emerald-700'
              }`}>
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div className="flex items-center gap-3">
                    <div className="p-3 rounded-lg bg-white/10 backdrop-blur-sm">
                      {evaluationResult.overall_status === 'CRITICAL' && <AlertOctagon size={28} className="text-rose-300" />}
                      {evaluationResult.overall_status === 'CAUTION' && <AlertTriangle size={28} className="text-amber-300" />}
                      {evaluationResult.overall_status === 'PASS' && <CheckCircle2 size={28} className="text-emerald-300" />}
                    </div>
                    <div>
                      <div className="text-xs uppercase font-extrabold tracking-wider text-white/80">
                        {isVi ? 'TRẠNG THÁI TỔNG QUAN / OVERALL SEVERITY STATUS' : 'OVERALL SEVERITY STATUS'}
                      </div>
                      <h2 className="text-xl font-black mt-0.5">
                        {evaluationResult.overall_status} — {isVi ? evaluationResult.overall_title_vi : evaluationResult.overall_title_en}
                      </h2>
                    </div>
                  </div>

                  <div className="flex items-center gap-2">
                    <button
                      onClick={handleExportPdf}
                      className="px-3.5 py-1.5 rounded-lg bg-white/20 hover:bg-white/30 text-white text-xs font-semibold backdrop-blur-sm transition flex items-center gap-1.5"
                    >
                      <Download size={14} />
                      <span>{isVi ? 'Xuất PDF Báo Cáo' : 'Export PDF'}</span>
                    </button>
                  </div>
                </div>

                <div className="mt-3 pt-3 border-t border-white/20 text-xs text-white/90">
                  <p className="font-medium">
                    {isVi ? evaluationResult.summary_vi : evaluationResult.summary_en}
                  </p>
                </div>
              </div>

              {/* Detailed Evaluation Table */}
              <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm">
                <h3 className="text-base font-bold text-slate-900 dark:text-slate-100 mb-4 flex items-center gap-2">
                  <FileCheck2 size={18} className="text-sky-600 dark:text-sky-400" />
                  <span>
                    {isVi
                      ? 'Bảng Phân Tích Chi Tiết Bảo Trì Biến Điện Áp VT (NETA MTS Mục 7.10.2 & NETA Tables)'
                      : 'Detailed VT Maintenance Evaluation Table (NETA MTS Sec 7.10.2 & Tables)'}
                  </span>
                </h3>

                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs">
                    <thead>
                      <tr className="border-b border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-800/60 text-slate-600 dark:text-slate-300 font-semibold uppercase tracking-wider">
                        <th className="py-3 px-3">{isVi ? 'Hạng Mục Thử Nghiệm / Test Item' : 'Test Item'}</th>
                        <th className="py-3 px-3">{isVi ? 'Tiêu Chuẩn Tham Chiếu' : 'Standard Ref'}</th>
                        <th className="py-3 px-3">{isVi ? 'Giá Trị Đo Thực Tế' : 'Measured Value'}</th>
                        <th className="py-3 px-3 text-center">{isVi ? 'Đánh Giá' : 'Evaluation'}</th>
                        <th className="py-3 px-3">{isVi ? 'Ghi Chú Kỹ Thuật Chẩn Đoán' : 'Diagnostic Notes'}</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 dark:divide-slate-800">
                      {evaluationResult.evaluation_table.map((row, idx) => (
                        <tr key={idx} className="hover:bg-slate-50 dark:hover:bg-slate-800/40 transition">
                          <td className="py-3 px-3 font-semibold text-slate-800 dark:text-slate-200">
                            {isVi ? row.item_vn : row.item}
                          </td>
                          <td className="py-3 px-3 font-mono text-slate-500 dark:text-slate-400">
                            {row.standard_ref}
                          </td>
                          <td className="py-3 px-3 font-mono font-medium text-slate-700 dark:text-slate-300">
                            {row.measured_value}
                          </td>
                          <td className="py-3 px-3 text-center">
                            <span className={`inline-block px-2.5 py-0.5 rounded-full text-xs font-bold ${
                              row.status === 'PASS'
                                ? 'bg-emerald-100 text-emerald-800 dark:bg-emerald-950/60 dark:text-emerald-300'
                                : row.status === 'CRITICAL'
                                ? 'bg-rose-100 text-rose-800 dark:bg-rose-950/60 dark:text-rose-300 animate-pulse'
                                : row.status === 'FAIL'
                                ? 'bg-rose-100 text-rose-800 dark:bg-rose-950/60 dark:text-rose-300'
                                : 'bg-amber-100 text-amber-800 dark:bg-amber-950/60 dark:text-amber-300'
                            }`}>
                              {row.status}
                            </span>
                          </td>
                          <td className="py-3 px-3 text-slate-600 dark:text-slate-400">
                            {isVi ? row.notes_vn : row.notes}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Technical Analysis & Hazard Deep Dive */}
              <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm space-y-4">
                <h3 className="text-base font-bold text-slate-900 dark:text-slate-100 flex items-center gap-2">
                  <Activity size={18} className="text-sky-600 dark:text-sky-400" />
                  <span>
                    {isVi
                      ? 'Phân Tích Kỹ Thuật & Nguy Cơ Nổ Biến Điện Áp Pha B (Technical Analysis & Hazard)'
                      : 'Technical Analysis & Phase B VT Dielectric Flashover Hazard'}
                  </span>
                </h3>

                <div className="space-y-3 text-sm">
                  {(isVi ? evaluationResult.technical_analysis_vi : evaluationResult.technical_analysis_en).map((item, idx) => (
                    <div key={idx} className="p-4 rounded-lg bg-slate-50 dark:bg-slate-800/60 border border-slate-200 dark:border-slate-700/60 text-slate-700 dark:text-slate-300 leading-relaxed">
                      <p>{item}</p>
                    </div>
                  ))}
                </div>
              </div>

              {/* Actionable Recommendations for FSE */}
              <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm space-y-4">
                <h3 className="text-base font-bold text-slate-900 dark:text-slate-100 flex items-center gap-2">
                  <Wrench size={18} className="text-emerald-600 dark:text-emerald-400" />
                  <span>
                    {isVi
                      ? 'Khuyến Nghị Hành Động Chi Tiết Cho Kỹ Sư FSE (Actionable Recommendations)'
                      : 'Actionable Step-by-Step Recommendations for FSE'}
                  </span>
                </h3>

                <div className="space-y-3">
                  {(isVi ? evaluationResult.recommendations_vi : evaluationResult.recommendations_en).map((rec, idx) => (
                    <div key={idx} className="flex items-start gap-3 p-3.5 rounded-lg bg-emerald-50/60 dark:bg-emerald-950/20 border border-emerald-200 dark:border-emerald-800/40 text-sm text-slate-800 dark:text-slate-200">
                      <div className="w-6 h-6 rounded-full bg-emerald-600 text-white font-bold flex items-center justify-center text-xs shrink-0 mt-0.5">
                        {idx + 1}
                      </div>
                      <p className="leading-relaxed">{rec}</p>
                    </div>
                  ))}
                </div>
              </div>

              {/* AI Generated Markdown if available */}
              {evaluationResult.ai_generated_markdown && (
                <div className="bg-slate-950 border border-sky-900/60 rounded-xl p-5 text-slate-200 text-xs font-mono space-y-2 shadow-inner">
                  <div className="flex items-center justify-between text-sky-400 font-semibold border-b border-slate-800 pb-2">
                    <span className="flex items-center gap-1.5">
                      <Sparkles size={14} />
                      <span>Gemini 2.5 Flash Analytical Output (NETA MTS 7.10.2)</span>
                    </span>
                  </div>
                  <pre className="whitespace-pre-wrap leading-relaxed max-h-96 overflow-y-auto">
                    {evaluationResult.ai_generated_markdown}
                  </pre>
                </div>
              )}
            </div>
          ) : (
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-12 text-center shadow-sm">
              <Sparkles size={40} className="text-sky-400 mx-auto mb-3" />
              <h3 className="text-base font-bold text-slate-800 dark:text-slate-200">
                {isVi ? 'Chưa Có Kết Quả Chẩn Đoán AI' : 'No AI Diagnostic Results Yet'}
              </h3>
              <p className="text-xs text-slate-500 mt-1 max-w-md mx-auto">
                {isVi
                  ? 'Bấm nút "Chẩn Đoán NETA MTS AI" ở trên để phân tích dữ liệu kiểm tra biến điện áp VT/PT theo đúng chuẩn NETA MTS Mục 7.10.2.'
                  : 'Click "Run NETA MTS AI Analysis" above to evaluate voltage transformer maintenance data.'}
              </p>
              <button
                onClick={handleRunEvaluation}
                className="mt-4 px-4 py-2 rounded-xl bg-sky-600 hover:bg-sky-700 text-white text-xs font-bold transition"
              >
                {isVi ? 'Chạy Chẩn Đoán Ngay' : 'Run Analysis Now'}
              </button>
            </div>
          )}
        </div>
      )}

      {/* TAB 3: SYSTEM INSTRUCTION PROMPT */}
      {activeTab === 'prompt' && (
        <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-200 dark:border-slate-800 pb-3">
            <div>
              <h3 className="font-bold text-slate-900 dark:text-slate-100 text-sm flex items-center gap-2">
                <FileText size={18} className="text-sky-600 dark:text-sky-400" />
                <span>{isVi ? 'System Instruction Prompt Dành Cho Google AI Studio (TEV COE 7.10.2 VT)' : 'Google AI Studio System Instruction Prompt'}</span>
              </h3>
              <p className="text-xs text-slate-500">
                {isVi ? 'Copy đoạn văn bản dưới đây và dán vào mục System Instructions trên Google AI Studio' : 'Copy and paste into System Instructions field in Google AI Studio'}
              </p>
            </div>

            <button
              onClick={handleCopyPrompt}
              className="flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg bg-sky-600 hover:bg-sky-700 text-white text-xs font-semibold shadow transition"
            >
              {copySuccess ? <Check size={14} /> : <Copy size={14} />}
              <span>{copySuccess ? (isVi ? 'Đã Copy!' : 'Copied!') : (isVi ? 'Copy System Prompt' : 'Copy System Prompt')}</span>
            </button>
          </div>

          <div className="bg-slate-950 text-slate-200 p-4 rounded-xl text-xs font-mono max-h-[500px] overflow-y-auto border border-slate-800 leading-relaxed">
            <pre className="whitespace-pre-wrap">{MTS_VT_SYSTEM_INSTRUCTION}</pre>
          </div>
        </div>
      )}

      {/* TAB 4: JSON INPUT SCHEMA */}
      {activeTab === 'json' && (
        <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-200 dark:border-slate-800 pb-3">
            <div>
              <h3 className="font-bold text-slate-900 dark:text-slate-100 text-sm flex items-center gap-2">
                <FileCheck2 size={18} className="text-sky-600 dark:text-sky-400" />
                <span>{isVi ? 'Mẫu JSON User Prompt Đầu Vào Dùng Cho Google AI Studio' : 'JSON User Prompt Payload for Google AI Studio'}</span>
              </h3>
              <p className="text-xs text-slate-500">
                {isVi ? 'Dữ liệu JSON mô phỏng đầy đủ các phép đo theo ANSI/NETA MTS-2023 Mục 7.10.2' : 'Complete JSON structure per ANSI/NETA MTS-2023 Section 7.10.2'}
              </p>
            </div>

            <button
              onClick={handleCopyJson}
              className="flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg bg-sky-600 hover:bg-sky-700 text-white text-xs font-semibold shadow transition"
            >
              {jsonCopySuccess ? <Check size={14} /> : <Copy size={14} />}
              <span>{jsonCopySuccess ? (isVi ? 'Đã Copy JSON!' : 'Copied JSON!') : (isVi ? 'Copy JSON Payload' : 'Copy JSON Payload')}</span>
            </button>
          </div>

          <div className="bg-slate-950 text-slate-200 p-4 rounded-xl text-xs font-mono max-h-[500px] overflow-y-auto border border-slate-800 leading-relaxed">
            <pre className="whitespace-pre-wrap">{JSON.stringify(formData, null, 2)}</pre>
          </div>
        </div>
      )}
    </div>
  );
};
