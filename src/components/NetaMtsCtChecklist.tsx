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
  NetaMtsCtInputPayload,
  NetaMtsCtEvaluationResult,
  performDeterministicMtsCtAnalysis,
  MTS_CT_SYSTEM_INSTRUCTION
} from '../../server/netaMtsCtAnalyzer';
import { CustomerSelectDropdown } from './CustomerSelectDropdown';
import { ProjectSelectDropdown } from './ProjectSelectDropdown';
import { FseEngineerDropdown } from './FseEngineerDropdown';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
  initialLanguage?: 'vi' | 'en';
}

export const SAMPLE_MTS_CT_DEFECT_PAYLOAD: NetaMtsCtInputPayload = {
  tev_coe_header: {
    module: "CT-MAINTENANCE-MTS-7.10.1",
    project_name: "Tram Bien Ap Trung The TEV Industrial Substation",
    equipment_tag: "15KV-BUS-INCOMER-CT-SET (Set of 3 CTs)",
    voltage_rating_kV: 15,
    ct_ratio_rating: "1200:5 A",
    accuracy_class_rating: "C400 / 0.3B-1.8",
    fse_name: "Nguyen Van P",
    test_date: "2026-10-19",
    station_name: "Substation SS-01",
    bay_location: "Incomer Bay 15kV-01"
  },
  visual_and_mechanical_inspection: {
    nameplate_verification: {
      matched: true,
      notes: "Trùng khớp 100% hồ sơ thiết kế và sơ đồ nhị thứ rơ-le"
    },
    physical_condition: "Good condition, clean resin surfaces, no cracks or oil leakage",
    bolted_connection_torque_table_100_12: "Pass (Torque wrench verified per Table 100.12)",
    clearance_verification: "Pass (Adequate clearance maintained per 15kV class)",
    secondary_grounding_point_ieee_c57_13_3: {
      ground_points_detected: 2,
      ground_locations: "Grounded at CT Terminal Box AND grounded at Protection Relay Panel",
      status: "Fail (Multiple ground points detected on secondary loop violating IEEE C57.13.3)"
    }
  },
  electrical_tests_7_10_1_b: {
    insulation_resistance_megohms_table_100_5: {
      primary_to_ground_2500v_dc_Mohm: 12500.0,
      secondary_to_ground_1000v_dc_Mohm: 450.0,
      status: "Pass (Exceeds NETA Table 100.5 minimum thresholds)"
    },
    polarity_test_ieee_c57_13_1: "Pass (H1-X1 correct polarity verified on all 3 phases)",
    ratio_verification_test_ieee_c57_13: {
      applied_primary_current_A: 1200,
      measured_secondary_current_phase_a_A: 5.001,
      measured_secondary_current_phase_b_A: 4.998,
      measured_secondary_current_phase_c_A: 4.902,
      ratio_error_phase_c_percent: -1.96,
      status: "Fail (Phase C ratio error -1.96% exceeds 0.3% accuracy class limit)"
    },
    excitation_saturation_test_ieee_c57_13_1: {
      phase_a_knee_point_voltage_V: 420,
      phase_b_knee_point_voltage_V: 418,
      phase_c_knee_point_voltage_V: 185,
      status: "Fail (Phase C knee-point voltage 185V severely depressed compared to C400 rating)"
    },
    winding_dc_resistance_ohms: {
      phase_a_ohms: 0.42,
      phase_b_ohms: 0.41,
      phase_c_ohms: 0.28,
      status: "Fail (Phase C resistance 0.28 ohm deviates > 30% indicating shorted turns)"
    },
    secondary_burden_measurement: {
      measured_burden_phase_c_VA: 12.5,
      ct_rated_burden_VA: 45.0,
      status: "Pass (Measured burden well within CT rating)"
    },
    dielectric_withstand_test_table_100_9: {
      test_voltage_kV_ac: 27.0,
      duration_minutes: 1,
      status: "Pass (No dielectric breakdown observed)"
    },
    power_factor_test_percent: {
      measured_pf_percent: 0.45,
      status: "Pass (Normal dissipation factor)"
    }
  },
  post_test_survey: {
    thermographic_survey_table_100_18: "Pass (Normal thermal distribution under rated load)"
  }
};

export const SAMPLE_MTS_CT_PASS_PAYLOAD: NetaMtsCtInputPayload = {
  tev_coe_header: {
    module: "CT-MAINTENANCE-MTS-7.10.1",
    project_name: "Tram 110kV KCN Song Khoai TEV Substation",
    equipment_tag: "110KV-OUTGOING-LINE-CT-01",
    voltage_rating_kV: 110,
    ct_ratio_rating: "2000:5 A",
    accuracy_class_rating: "C800 / 0.2S",
    fse_name: "Tran Minh T",
    test_date: "2026-10-18",
    station_name: "110kV Song Khoai",
    bay_location: "Bay 171"
  },
  visual_and_mechanical_inspection: {
    nameplate_verification: {
      matched: true,
      notes: "Hoàn toàn trùng khớp bản vẽ sơ đồ nguyên lý"
    },
    physical_condition: "Pristine porcelain bushing, oil level normal, no weeping or corona traces",
    bolted_connection_torque_table_100_12: "Pass (Calibrated torque wrench verified per Table 100.12)",
    clearance_verification: "Pass (Full phase-to-earth clearances per 110kV standard)",
    secondary_grounding_point_ieee_c57_13_3: {
      ground_points_detected: 1,
      ground_locations: "Grounded at Protection Relay Panel MK Cabinet ONLY",
      status: "Pass (Single grounding point verified compliant with IEEE C57.13.3)"
    }
  },
  electrical_tests_7_10_1_b: {
    insulation_resistance_megohms_table_100_5: {
      primary_to_ground_2500v_dc_Mohm: 18500.0,
      secondary_to_ground_1000v_dc_Mohm: 850.0,
      status: "Pass (Exceeds Table 100.5 thresholds)"
    },
    polarity_test_ieee_c57_13_1: "Pass (H1-X1 polarity verified with DC pulse test)",
    ratio_verification_test_ieee_c57_13: {
      applied_primary_current_A: 2000,
      measured_secondary_current_phase_a_A: 5.002,
      measured_secondary_current_phase_b_A: 4.999,
      measured_secondary_current_phase_c_A: 5.001,
      ratio_error_phase_c_percent: 0.02,
      status: "Pass (Ratio error well within +/- 0.2% class)"
    },
    excitation_saturation_test_ieee_c57_13_1: {
      phase_a_knee_point_voltage_V: 825,
      phase_b_knee_point_voltage_V: 820,
      phase_c_knee_point_voltage_V: 822,
      status: "Pass (Knee-point voltages meet C800 rating >= 800V)"
    },
    winding_dc_resistance_ohms: {
      phase_a_ohms: 0.84,
      phase_b_ohms: 0.835,
      phase_c_ohms: 0.842,
      status: "Pass (Phase-to-phase deviation < 1.0% compliant with IEEE C57.13.1)"
    },
    secondary_burden_measurement: {
      measured_burden_phase_c_VA: 18.2,
      ct_rated_burden_VA: 60.0,
      status: "Pass (Burden well within rated 60VA)"
    },
    dielectric_withstand_test_table_100_9: {
      test_voltage_kV_ac: 140.0,
      duration_minutes: 1,
      status: "Pass (Withstand passed without breakdown)"
    },
    power_factor_test_percent: {
      measured_pf_percent: 0.38,
      status: "Pass (Power factor 0.38% < 0.5% limit at 20°C)"
    }
  },
  post_test_survey: {
    thermographic_survey_table_100_18: "Pass (Thermal gradient < 1.5°C per Table 100.18)"
  }
};

export const SAMPLE_MTS_CT_CAUTION_PAYLOAD: NetaMtsCtInputPayload = {
  tev_coe_header: {
    module: "CT-MAINTENANCE-MTS-7.10.1",
    project_name: "Tram Phap Van TEV Distribution Center",
    equipment_tag: "24KV-FEEDER-03-CT-SET",
    voltage_rating_kV: 24,
    ct_ratio_rating: "600:5 A",
    accuracy_class_rating: "C200 / 0.5B-0.9",
    fse_name: "Le Hoang D",
    test_date: "2026-10-17",
    station_name: "Dist Substation D-03",
    bay_location: "Feeder Bay 03"
  },
  visual_and_mechanical_inspection: {
    nameplate_verification: {
      matched: true,
      notes: "Khớp hồ sơ thiết kế"
    },
    physical_condition: "Minor dust accumulation, cleaned before electrical testing",
    bolted_connection_torque_table_100_12: "Caution (Primary busbar H1 bolt torque loose, re-torqued to Table 100.12)",
    clearance_verification: "Pass",
    secondary_grounding_point_ieee_c57_13_3: {
      ground_points_detected: 1,
      ground_locations: "Relay Panel single point",
      status: "Pass"
    }
  },
  electrical_tests_7_10_1_b: {
    insulation_resistance_megohms_table_100_5: {
      primary_to_ground_2500v_dc_Mohm: 6200.0,
      secondary_to_ground_1000v_dc_Mohm: 180.0,
      status: "Pass (Meets Table 100.5 thresholds)"
    },
    polarity_test_ieee_c57_13_1: "Pass (H1-X1 correct)",
    ratio_verification_test_ieee_c57_13: {
      applied_primary_current_A: 600,
      measured_secondary_current_phase_a_A: 4.975,
      measured_secondary_current_phase_b_A: 4.968,
      measured_secondary_current_phase_c_A: 4.970,
      ratio_error_phase_c_percent: -0.60,
      status: "Caution (Ratio error -0.60% borderline for 0.5 class)"
    },
    excitation_saturation_test_ieee_c57_13_1: {
      phase_a_knee_point_voltage_V: 205,
      phase_b_knee_point_voltage_V: 198,
      phase_c_knee_point_voltage_V: 202,
      status: "Pass (Meets C200 rating >= 190V)"
    },
    winding_dc_resistance_ohms: {
      phase_a_ohms: 0.22,
      phase_b_ohms: 0.23,
      phase_c_ohms: 0.225,
      status: "Pass (Deviation < 4.5%)"
    },
    secondary_burden_measurement: {
      measured_burden_phase_c_VA: 26.5,
      ct_rated_burden_VA: 22.5,
      status: "Fail (Burden 26.5 VA exceeds rated 22.5 VA, risks early saturation under heavy faults)"
    },
    dielectric_withstand_test_table_100_9: {
      test_voltage_kV_ac: 37.0,
      duration_minutes: 1,
      status: "Pass"
    }
  },
  post_test_survey: {
    thermographic_survey_table_100_18: "Pass"
  }
};

export const NetaMtsCtChecklist: React.FC<Props> = ({
  onSaveReport,
  onNavigateToReports,
  initialLanguage = 'vi'
}) => {
  const [lang, setLang] = useState<'vi' | 'en'>(initialLanguage);
  const [activeTab, setActiveTab] = useState<'checklist' | 'evaluation' | 'prompt' | 'json'>('checklist');
  const [activePhase, setActivePhase] = useState<'phase1' | 'phase2' | 'phase3'>('phase1');
  const [isEvaluating, setIsEvaluating] = useState(false);
  const [evaluationResult, setEvaluationResult] = useState<NetaMtsCtEvaluationResult | null>(null);
  const [copySuccess, setCopySuccess] = useState(false);
  const [jsonCopySuccess, setJsonCopySuccess] = useState(false);
  const [saveSuccess, setSaveSuccess] = useState(false);

  // Form State
  const [formData, setFormData] = useState<NetaMtsCtInputPayload>(SAMPLE_MTS_CT_DEFECT_PAYLOAD);

  const isVi = lang === 'vi';

  // Handler to load presets
  const handleLoadPreset = (scenario: 'defect' | 'pass' | 'caution') => {
    let payload = SAMPLE_MTS_CT_DEFECT_PAYLOAD;
    if (scenario === 'pass') payload = SAMPLE_MTS_CT_PASS_PAYLOAD;
    if (scenario === 'caution') payload = SAMPLE_MTS_CT_CAUTION_PAYLOAD;

    setFormData(JSON.parse(JSON.stringify(payload)));
    setEvaluationResult(null);
  };

  // Run AI / Deterministic Evaluation
  const handleRunEvaluation = async () => {
    setIsEvaluating(true);
    try {
      const response = await fetch('/api/field-service/neta-mts-ct-analyze', {
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
      const localEval = performDeterministicMtsCtAnalysis(formData);
      setEvaluationResult(localEval);
      setActiveTab('evaluation');
    } catch (err) {
      console.warn('API error, using local deterministic engine:', err);
      const localEval = performDeterministicMtsCtAnalysis(formData);
      setEvaluationResult(localEval);
      setActiveTab('evaluation');
    } finally {
      setIsEvaluating(false);
    }
  };

  // Copy AI Studio System Instruction Prompt
  const handleCopyPrompt = () => {
    navigator.clipboard.writeText(MTS_CT_SYSTEM_INSTRUCTION);
    setCopySuccess(true);
    setTimeout(() => setCopySuccess(false), 2000);
  };

  // Copy JSON Input
  const handleCopyJson = () => {
    navigator.clipboard.writeText(JSON.stringify(formData, null, 2));
    setJsonCopySuccess(true);
    setTimeout(() => setJsonCopySuccess(false), 2000);
  };

  // Save Report to CMMS / App
  const handleSaveReport = () => {
    if (!evaluationResult) {
      const computed = performDeterministicMtsCtAnalysis(formData);
      setEvaluationResult(computed);
    }
    const reportData = {
      id: `NETA-MTS-CT-${Date.now()}`,
      moduleType: 'MTS-CT-7.10.1',
      title: `${formData.tev_coe_header.equipment_tag} - NETA MTS-2023 Sec 7.10.1`,
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
    const currentEval = evaluationResult || performDeterministicMtsCtAnalysis(formData);

    // Title & Header
    doc.setFillColor(15, 23, 42); // slate-900
    doc.rect(0, 0, 210, 26, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(14);
    doc.text('TEV COE - ELECTRICAL TESTING & COMMISSIONING PLATFORM', 14, 11);
    doc.setFontSize(10);
    doc.setTextColor(148, 163, 184);
    doc.text('ANSI/NETA MTS-2023 SECTION 7.10.1 - INSTRUMENT TRANSFORMERS (CT)', 14, 18);

    // Metadata Box
    doc.setFillColor(248, 250, 252);
    doc.rect(14, 32, 182, 30, 'F');
    doc.setDrawColor(203, 213, 225);
    doc.rect(14, 32, 182, 30, 'S');

    doc.setTextColor(15, 23, 42);
    doc.setFontSize(9);
    doc.text(`Project: ${formData.tev_coe_header.project_name}`, 18, 40);
    doc.text(`Equipment Tag: ${formData.tev_coe_header.equipment_tag}`, 18, 46);
    doc.text(`Rating: ${formData.tev_coe_header.ct_ratio_rating} | Class: ${formData.tev_coe_header.accuracy_class_rating}`, 18, 52);
    doc.text(`Voltage: ${formData.tev_coe_header.voltage_rating_kV} kV`, 18, 58);

    doc.text(`FSE Engineer: ${formData.tev_coe_header.fse_name}`, 115, 40);
    doc.text(`Test Date: ${formData.tev_coe_header.test_date}`, 115, 46);
    doc.text(`Location: ${formData.tev_coe_header.bay_location || 'Switchgear Bay'}`, 115, 52);
    doc.text(`Standard: ANSI/NETA MTS-2023 (Sec 7.10.1)`, 115, 58);

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

    // Sign off block at bottom
    doc.setFontSize(8);
    doc.setTextColor(100, 116, 139);
    doc.text('Tested by (FSE Signature): _______________________', 14, 280);
    doc.text('Approved by (Lead Commissioning): _______________________', 110, 280);

    doc.save(`NETA_MTS_CT_${formData.tev_coe_header.equipment_tag}_${formData.tev_coe_header.test_date}.pdf`);
  };

  // Calculations for display
  const phaseDeviation = useMemo(() => {
    const rA = formData.electrical_tests_7_10_1_b.winding_dc_resistance_ohms.phase_a_ohms;
    const rB = formData.electrical_tests_7_10_1_b.winding_dc_resistance_ohms.phase_b_ohms;
    const rC = formData.electrical_tests_7_10_1_b.winding_dc_resistance_ohms.phase_c_ohms;
    const avg = (rA + rB + rC) / 3;
    if (avg === 0) return 0;
    const maxDev = Math.max(Math.abs(rA - avg), Math.abs(rB - avg), Math.abs(rC - avg));
    return (maxDev / avg) * 100;
  }, [formData]);

  return (
    <div className="space-y-6">
      {/* Top Banner */}
      <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 rounded-2xl p-6 text-white shadow-xl border border-indigo-900/40 relative overflow-hidden">
        <div className="absolute right-0 top-0 w-96 h-96 bg-indigo-500/10 rounded-full blur-3xl -mr-20 -mt-20 pointer-events-none" />
        
        <div className="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-6 relative z-10">
          <div>
            <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full text-xs font-semibold bg-indigo-500/20 text-indigo-300 border border-indigo-400/30 mb-3">
              <Gauge size={14} className="text-indigo-400" />
              <span>ANSI/NETA MTS-2023 Mục 7.10.1 • Bảng 100.5, 100.9, 100.12, 100.18</span>
            </div>
            <h1 className="text-2xl lg:text-3xl font-extrabold tracking-tight">
              {isVi 
                ? 'Kiểm Tra & Chẩn Đoán Biến Dòng Điện Đo Lường & Bảo Vệ (Current Transformers - CT)'
                : 'Instrument Transformers: Current Transformers (CT) Maintenance Testing & Diagnostics'}
            </h1>
            <p className="text-slate-300 text-sm mt-1 max-w-3xl">
              {isVi
                ? 'Quy trình thử nghiệm bảo trì biến dòng điện CT chuẩn hóa theo ANSI/NETA MTS-2023 Section 7.10.1, IEEE C57.13 (Cực tính, Tỉ số, Điểm bão hòa Knee-point) và IEEE C57.13.3 (Quy chuẩn DUY NHẤT 1 điểm nối đất mạch thứ cấp).'
                : 'CT maintenance testing protocol standardized per ANSI/NETA MTS-2023 Sec 7.10.1, IEEE C57.13 (Polarity, Ratio, Saturation Knee-Point) and IEEE C57.13.3 (MANDATORY single secondary grounding point).'}
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {/* Language Switch */}
            <button
              onClick={() => setLang(isVi ? 'en' : 'vi')}
              className="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-xs font-medium border border-slate-700 text-slate-200 transition"
            >
              {isVi ? 'English (EN)' : 'Tiếng Việt (VN)'}
            </button>

            {/* Run AI Analysis */}
            <button
              onClick={handleRunEvaluation}
              disabled={isEvaluating}
              className="flex items-center gap-2 px-4 py-2 rounded-xl bg-gradient-to-r from-indigo-500 to-indigo-600 hover:from-indigo-600 hover:to-indigo-700 text-white font-semibold shadow-lg shadow-indigo-500/30 text-sm transition"
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

            {/* Export PDF */}
            <button
              onClick={handleExportPdf}
              className="flex items-center gap-1.5 px-3 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 text-sm font-medium transition"
            >
              <Download size={15} />
              <span>PDF</span>
            </button>

            {/* Save Report */}
            <button
              onClick={handleSaveReport}
              className="flex items-center gap-1.5 px-3 py-2 rounded-xl bg-emerald-600/90 hover:bg-emerald-500 text-white text-sm font-medium transition shadow-sm"
            >
              <Save size={15} />
              <span>{saveSuccess ? (isVi ? 'Đã Lưu!' : 'Saved!') : (isVi ? 'Lưu Biên Bản' : 'Save Report')}</span>
            </button>
          </div>
        </div>

        {/* Preset Scenarios Toolbar */}
        <div className="mt-5 pt-4 border-t border-slate-800 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex items-center gap-2 text-slate-300 font-medium">
            <Sliders size={14} className="text-indigo-400" />
            <span>{isVi ? 'Kịch bản mẫu kiểm thử nhanh (Field Scenarios):' : 'Quick Field Scenarios:'}</span>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => handleLoadPreset('defect')}
              className="flex items-center gap-1.5 px-2.5 py-1 rounded-md bg-rose-950/80 hover:bg-rose-900 text-rose-300 border border-rose-800/80 transition"
            >
              <AlertOctagon size={13} className="text-rose-400" />
              <span>{isVi ? 'Sự cố: Chập Vòng Pha C & Nối Đất 2 Điểm (Mẫu Prompt)' : 'Defect: Phase C Shorted & Dual Ground'}</span>
            </button>

            <button
              onClick={() => handleLoadPreset('pass')}
              className="flex items-center gap-1.5 px-2.5 py-1 rounded-md bg-emerald-950/80 hover:bg-emerald-900 text-emerald-300 border border-emerald-800/80 transition"
            >
              <CheckCircle2 size={13} className="text-emerald-400" />
              <span>{isVi ? 'Đạt Chuẩn: Bộ CT 110kV Đạt 100% (Pristine Pass)' : 'Pristine: 110kV CT 100% Compliant'}</span>
            </button>

            <button
              onClick={() => handleLoadPreset('caution')}
              className="flex items-center gap-1.5 px-2.5 py-1 rounded-md bg-amber-950/80 hover:bg-amber-900 text-amber-300 border border-amber-800/80 transition"
            >
              <AlertTriangle size={13} className="text-amber-400" />
              <span>{isVi ? 'Cảnh Báo: Quá Tải Burden & Lỏng Bu-lông H1' : 'Caution: Burden Overload & Torque Defect'}</span>
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
              ? 'border-indigo-600 text-indigo-600 dark:text-indigo-400'
              : 'border-transparent text-slate-500 hover:text-slate-800 dark:text-slate-400 dark:hover:text-slate-200'
          }`}
        >
          <Layers size={16} />
          <span>{isVi ? '1. Biên Bản Đo Đạc Site (Checklist NETA MTS 7.10.1)' : '1. Field Test Checklist'}</span>
        </button>

        <button
          onClick={() => setActiveTab('evaluation')}
          className={`flex items-center gap-2 py-3 px-4 border-b-2 font-semibold transition-colors relative ${
            activeTab === 'evaluation'
              ? 'border-indigo-600 text-indigo-600 dark:text-indigo-400'
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
              ? 'border-indigo-600 text-indigo-600 dark:text-indigo-400'
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
              ? 'border-indigo-600 text-indigo-600 dark:text-indigo-400'
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
              <Building size={18} className="text-indigo-600 dark:text-indigo-400" />
              <span>{isVi ? 'Thông Tin Dự Án & Thiết Bị Biến Dòng CT' : 'Project & Instrument CT Information'}</span>
            </h3>

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 text-sm">
              <div>
                <CustomerSelectDropdown
                  label={isVi ? 'Tên Khách Hàng (Master)' : 'Customer Name (Master)'}
                  value={formData.tev_coe_header.customer_name || ''}
                  onChange={(cName) => setFormData({
                    ...formData,
                    tev_coe_header: { ...formData.tev_coe_header, customer_name: cName }
                  })}
                  placeholder={isVi ? 'Chọn hoặc thêm Khách Hàng...' : 'Select or add Customer...'}
                />
              </div>

              <div>
                <ProjectSelectDropdown
                  label={isVi ? 'Tên Dự Án / Nhà Máy' : 'Project / Plant Name'}
                  value={formData.tev_coe_header.project_name}
                  customerName={formData.tev_coe_header.customer_name}
                  onChange={(pName) => setFormData({
                    ...formData,
                    tev_coe_header: { ...formData.tev_coe_header, project_name: pName }
                  })}
                  placeholder={isVi ? 'Chọn hoặc thêm Dự Án...' : 'Select or add Project...'}
                  standard="mts"
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
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-100 text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none font-semibold text-indigo-600 dark:text-indigo-400"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">
                  {isVi ? 'Tỉ Số Biến Dòng (CT Ratio)' : 'CT Ratio Rating'}
                </label>
                <input
                  type="text"
                  value={formData.tev_coe_header.ct_ratio_rating}
                  onChange={(e) => setFormData({
                    ...formData,
                    tev_coe_header: { ...formData.tev_coe_header, ct_ratio_rating: e.target.value }
                  })}
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-100 text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none font-mono"
                  placeholder="e.g. 1200:5 A"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">
                  {isVi ? 'Cấp Chính Xác & C-Class (Accuracy)' : 'Accuracy Class & C-Class'}
                </label>
                <input
                  type="text"
                  value={formData.tev_coe_header.accuracy_class_rating}
                  onChange={(e) => setFormData({
                    ...formData,
                    tev_coe_header: { ...formData.tev_coe_header, accuracy_class_rating: e.target.value }
                  })}
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-100 text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none font-mono"
                  placeholder="e.g. C400 / 0.3B-1.8"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">
                  {isVi ? 'Điện Áp Định Mức Hệ Thống (kV)' : 'Voltage Rating (kV)'}
                </label>
                <input
                  type="number"
                  value={formData.tev_coe_header.voltage_rating_kV}
                  onChange={(e) => setFormData({
                    ...formData,
                    tev_coe_header: { ...formData.tev_coe_header, voltage_rating_kV: parseFloat(e.target.value) || 15 }
                  })}
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-100 text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">
                  {isVi ? 'Kỹ Sư Nghiệm Thu / FSE' : 'Commissioning & FSE Engineer'}
                </label>
                <FseEngineerDropdown
                  value={formData.tev_coe_header.fse_name}
                  onChange={(engName) => setFormData({
                    ...formData,
                    tev_coe_header: { ...formData.tev_coe_header, fse_name: engName }
                  })}
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
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-100 text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">
                  {isVi ? 'Vị Trí Ngăn Lộ / Tủ Điện' : 'Bay / Panel Location'}
                </label>
                <input
                  type="text"
                  value={formData.tev_coe_header.bay_location || ''}
                  onChange={(e) => setFormData({
                    ...formData,
                    tev_coe_header: { ...formData.tev_coe_header, bay_location: e.target.value }
                  })}
                  className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-100 text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none"
                  placeholder="e.g. Incomer Bay 15kV-01"
                />
              </div>
            </div>
          </div>

          {/* Checklist Phase Navigation */}
          <div className="flex border-b border-slate-200 dark:border-slate-800 space-x-4 text-xs font-semibold">
            <button
              onClick={() => setActivePhase('phase1')}
              className={`pb-2 transition border-b-2 ${activePhase === 'phase1' ? 'border-indigo-600 text-indigo-600 dark:text-indigo-400' : 'border-transparent text-slate-500 hover:text-slate-800'}`}
            >
              {isVi ? 'Giai Đoạn 1: Kiểm Tra Thị Giác & Cơ Khí (NETA MTS 7.10.1.A)' : 'Phase 1: Visual & Mechanical (7.10.1.A)'}
            </button>
            <button
              onClick={() => setActivePhase('phase2')}
              className={`pb-2 transition border-b-2 ${activePhase === 'phase2' ? 'border-indigo-600 text-indigo-600 dark:text-indigo-400' : 'border-transparent text-slate-500 hover:text-slate-800'}`}
            >
              {isVi ? 'Giai Đoạn 2: Thử Nghiệm Điện & Đặc Tuyến (NETA MTS 7.10.1.B)' : 'Phase 2: Electrical Tests (7.10.1.B)'}
            </button>
            <button
              onClick={() => setActivePhase('phase3')}
              className={`pb-2 transition border-b-2 ${activePhase === 'phase3' ? 'border-indigo-600 text-indigo-600 dark:text-indigo-400' : 'border-transparent text-slate-500 hover:text-slate-800'}`}
            >
              {isVi ? 'Giai Đoạn 3: Khảo Sát Hồng Ngoại & Sau Thử Nghiệm' : 'Phase 3: Post-Test Survey & Thermography'}
            </button>
          </div>

          {/* PHASE 1: VISUAL & MECHANICAL */}
          {activePhase === 'phase1' && (
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm space-y-6">
              <div className="flex items-center justify-between border-b border-slate-200 dark:border-slate-800 pb-3">
                <div className="flex items-center gap-2">
                  <div className="w-8 h-8 rounded-lg bg-indigo-100 dark:bg-indigo-900/40 text-indigo-600 dark:text-indigo-400 flex items-center justify-center font-bold">
                    A
                  </div>
                  <div>
                    <h4 className="font-bold text-slate-900 dark:text-slate-100 text-sm">
                      {isVi ? 'Kiểm Tra Thị Giác & Cơ Khí Theo ANSI/NETA MTS-2023 Mục 7.10.1.A' : 'Visual & Mechanical Inspection per ANSI/NETA MTS-2023 Sec 7.10.1.A'}
                    </h4>
                    <p className="text-xs text-slate-500">
                      {isVi ? 'Đối soát nhãn mác, tình trạng cơ lý, lực siết bu-lông Bảng 100.12 và điểm tiếp địa mạch nhị thứ.' : 'Nameplate verification, physical condition, bolt torque per Table 100.12, and secondary grounding.'}
                    </p>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-5 text-sm">
                {/* A1: Nameplate */}
                <div className="p-4 rounded-lg bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-semibold text-slate-800 dark:text-slate-200">
                      A1. {isVi ? 'Rà soát Nhãn Mác & Bản Vẽ (7.10.1.A.1)' : 'Nameplate & Drawing Verification'}
                    </span>
                    <span className="text-xs px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 dark:bg-emerald-900/40 dark:text-emerald-300">
                      Trùng khớp 100%
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-2">
                    {isVi ? 'Đối soát tỉ số biến, cấp chính xác, dung lượng Burden, BIL với bản vẽ sơ đồ nguyên lý và rơ-le.' : 'Verify ratio, accuracy class, burden rating, BIL against drawings and relay settings.'}
                  </p>
                  <input
                    type="text"
                    value={formData.visual_and_mechanical_inspection.nameplate_verification.notes}
                    onChange={(e) => setFormData({
                      ...formData,
                      visual_and_mechanical_inspection: {
                        ...formData.visual_and_mechanical_inspection,
                        nameplate_verification: {
                          ...formData.visual_and_mechanical_inspection.nameplate_verification,
                          notes: e.target.value
                        }
                      }
                    })}
                    className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-xs"
                  />
                </div>

                {/* A2: Physical Condition */}
                <div className="p-4 rounded-lg bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-semibold text-slate-800 dark:text-slate-200">
                      A2. {isVi ? 'Kiểm Tra Cơ Lý & Vệ Sinh (7.10.1.A.2 & 3)' : 'Physical Condition & Cleanliness'}
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-2">
                    {isVi ? 'Inspect tình trạng cơ học vỏ đúc/sứ cách điện, nứt vỡ, rò rỉ dầu/keo, vết phóng điện bề mặt.' : 'Inspect for physical damage, oil leakage, cracked resin, corona tracking, and clean surfaces.'}
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

                {/* A3: Bolted Connections */}
                <div className="p-4 rounded-lg bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-semibold text-slate-800 dark:text-slate-200">
                      A3. {isVi ? 'Lực Siết Bu-lông Kết Nối (Bảng 100.12)' : 'Bolted Connections & Torque (Table 100.12)'}
                    </span>
                    <span className="text-xs px-2 py-0.5 rounded bg-blue-100 text-blue-800 dark:bg-blue-900/40 dark:text-blue-300">
                      Table 100.12
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-2">
                    {isVi ? 'Kiểm tra lực siết bu-lông mạch sơ & thứ cấp bằng cờ-lê lực hoặc đo điện trở tiếp xúc DLRO (< 50% lệch).' : 'Verify bolted connection tightness per Table 100.12 or low-resistance ohmmeter (< 50% deviation).'}
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

                {/* A4: Clearance */}
                <div className="p-4 rounded-lg bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-semibold text-slate-800 dark:text-slate-200">
                      A4. {isVi ? 'Khoảng Cách Cách Điện Sơ / Thứ Cấp (7.10.1.A.4)' : 'Clearance Verification (7.10.1.A.4)'}
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-2">
                    {isVi ? 'Đảm bảo khoảng cách an toàn điện giữa mạch sơ cấp (HV/MV) và mạch thứ cấp/vỏ tủ đất.' : 'Verify electrical clearance between primary terminals and secondary circuit/ground.'}
                  </p>
                  <input
                    type="text"
                    value={formData.visual_and_mechanical_inspection.clearance_verification}
                    onChange={(e) => setFormData({
                      ...formData,
                      visual_and_mechanical_inspection: {
                        ...formData.visual_and_mechanical_inspection,
                        clearance_verification: e.target.value
                      }
                    })}
                    className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-xs"
                  />
                </div>

                {/* A5: CRITICAL SECONDARY GROUNDING POINT CHECK */}
                <div className="md:col-span-2 p-5 rounded-xl bg-gradient-to-br from-indigo-50 to-slate-50 dark:from-indigo-950/30 dark:to-slate-800/60 border-2 border-indigo-400/50 shadow-sm">
                  <div className="flex items-start justify-between gap-4 mb-3">
                    <div className="flex items-center gap-2">
                      <ShieldAlert className="text-indigo-600 dark:text-indigo-400" size={20} />
                      <span className="font-bold text-slate-900 dark:text-slate-100 text-sm">
                        A5. {isVi ? 'Quy Chuẩn Nối Đất Mạch Thứ Cấp Biến Dòng (IEEE C57.13.3 & NETA 7.10.1.B.10)' : 'CT Secondary Circuit Grounding Point Verification (IEEE C57.13.3)'}
                      </span>
                    </div>
                    <span className="px-2.5 py-1 rounded-full text-xs font-bold uppercase tracking-wider bg-indigo-600 text-white shadow-sm">
                      {isVi ? 'BẮT BUỘC DUY NHẤT 1 ĐIỂM ĐẤT' : 'MANDATORY SINGLE POINT ONLY'}
                    </span>
                  </div>

                  <p className="text-xs text-slate-600 dark:text-slate-300 mb-4">
                    {isVi
                      ? 'Tiêu chuẩn IEEE C57.13.3 và NETA MTS 7.10.1.B.10 quy định BẮT BUỘC toàn bộ vòng lặp mạch dòng CT chỉ được phép nối đất tại DUY NHẤT 1 VỊ TRÍ (thường tại Tủ Rơ-le bảo vệ). Nối đất 2 điểm hoặc không nối đất là LỖI NGUY HIỂM CẤP ĐỘ CRITICAL, tạo dòng tuần hoàn rò đất gây tác động nhầm rơ-le.'
                      : 'IEEE C57.13.3 & NETA MTS 7.10.1.B.10 mandate that CT secondary circuits MUST have EXACTLY ONE grounding point (typically at relay panel). Multiple grounding points create parallel earth loops inducing false trip currents.'}
                  </p>

                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1">
                        {isVi ? 'Số điểm nối đất phát hiện được:' : 'Ground points detected:'}
                      </label>
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
                                status: pts === 1 
                                  ? 'Pass (Single point compliant with IEEE C57.13.3)'
                                  : `Fail (${pts} ground points detected violating IEEE C57.13.3)`
                              }
                            }
                          });
                        }}
                        className={`w-full px-3 py-2 rounded-lg border font-bold text-sm focus:outline-none ${
                          formData.visual_and_mechanical_inspection.secondary_grounding_point_ieee_c57_13_3.ground_points_detected === 1
                            ? 'border-emerald-500 bg-emerald-50 dark:bg-emerald-950/40 text-emerald-700 dark:text-emerald-300'
                            : 'border-rose-500 bg-rose-50 dark:bg-rose-950/40 text-rose-700 dark:text-rose-300 animate-pulse'
                        }`}
                      >
                        <option value={1}>1 Điểm Nối Đất Duy Nhất (Pass / Chuẩn IEEE C57.13.3)</option>
                        <option value={2}>2 Điểm Nối Đất (CRITICAL FAIL / Nối đất tại Hộp CT + Tủ Rơ-le)</option>
                        <option value={0}>0 Điểm (CRITICAL FAIL / Mất tiếp địa mạch dòng)</option>
                        <option value={3}>&gt; 2 Điểm (CRITICAL FAIL / Đa điểm nối đất)</option>
                      </select>
                    </div>

                    <div className="md:col-span-2">
                      <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1">
                        {isVi ? 'Vị trí điểm nối đất thực tế:' : 'Ground locations description:'}
                      </label>
                      <input
                        type="text"
                        value={formData.visual_and_mechanical_inspection.secondary_grounding_point_ieee_c57_13_3.ground_locations}
                        onChange={(e) => setFormData({
                          ...formData,
                          visual_and_mechanical_inspection: {
                            ...formData.visual_and_mechanical_inspection,
                            secondary_grounding_point_ieee_c57_13_3: {
                              ...formData.visual_and_mechanical_inspection.secondary_grounding_point_ieee_c57_13_3,
                              ground_locations: e.target.value
                            }
                          }
                        })}
                        className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-xs font-mono"
                        placeholder="e.g. Grounded at CT Terminal Box AND at Protection Relay Panel"
                      />
                    </div>
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
                  <div className="w-8 h-8 rounded-lg bg-indigo-100 dark:bg-indigo-900/40 text-indigo-600 dark:text-indigo-400 flex items-center justify-center font-bold">
                    B
                  </div>
                  <div>
                    <h4 className="font-bold text-slate-900 dark:text-slate-100 text-sm">
                      {isVi ? 'Thử Nghiệm Điện & Đo Đạc Đặc Tuyến (NETA MTS Mục 7.10.1.B)' : 'Electrical Tests per ANSI/NETA MTS-2023 Sec 7.10.1.B'}
                    </h4>
                    <p className="text-xs text-slate-500">
                      {isVi ? 'Cách điện Bảng 100.5, cực tính H1-X1, sai số tỉ số biến, đường cong Knee-point C400, điện trở cuộn dây, burden và điện áp tăng cao Bảng 100.9.' : 'Insulation Table 100.5, polarity, ratio verification, excitation knee-point curve, winding resistance, burden, and dielectric Table 100.9.'}
                    </p>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5 text-sm">
                {/* B1: Insulation Resistance */}
                <div className="p-4 rounded-xl bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-bold text-slate-800 dark:text-slate-200">
                      B1. {isVi ? 'Điện Trở Cách Điện (Bảng 100.5)' : 'Insulation Resistance (Table 100.5)'}
                    </span>
                    <span className="text-xs font-mono text-indigo-600 dark:text-indigo-400">
                      Table 100.5
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-3">
                    {isVi ? 'Sơ cấp vs Đất (2500V/5000V DC) ≥ 1,000 MΩ; Thứ cấp vs Đất (1000V DC) ≥ 100 MΩ.' : 'Primary to ground >= 1,000 MΩ (MV) / 5,000 MΩ (HV); Secondary >= 100 MΩ.'}
                  </p>

                  <div className="space-y-2 text-xs">
                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Sơ cấp - Đất (MΩ, 2500V DC):' : 'Primary to Ground (MΩ):'}
                      </label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_1_b.insulation_resistance_megohms_table_100_5.primary_to_ground_2500v_dc_Mohm}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_1_b: {
                            ...formData.electrical_tests_7_10_1_b,
                            insulation_resistance_megohms_table_100_5: {
                              ...formData.electrical_tests_7_10_1_b.insulation_resistance_megohms_table_100_5,
                              primary_to_ground_2500v_dc_Mohm: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Thứ cấp - Đất (MΩ, 1000V DC):' : 'Secondary to Ground (MΩ):'}
                      </label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_1_b.insulation_resistance_megohms_table_100_5.secondary_to_ground_1000v_dc_Mohm}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_1_b: {
                            ...formData.electrical_tests_7_10_1_b,
                            insulation_resistance_megohms_table_100_5: {
                              ...formData.electrical_tests_7_10_1_b.insulation_resistance_megohms_table_100_5,
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
                      B2. {isVi ? 'Thử Cực Tính Cuộn Dây' : 'Polarity Test (IEEE C57.13.1)'}
                    </span>
                    <span className="text-xs px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 dark:bg-emerald-900/40 dark:text-emerald-300">
                      H1-X1
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-3">
                    {isVi ? 'Xác nhận cực tính cuộn dây sơ/thứ cấp bằng xung DC hoặc AC (Trùng khớp 100% nhãn H1-X1).' : 'Verify polarity agreement with H1-X1 transformer terminal markings.'}
                  </p>

                  <select
                    value={formData.electrical_tests_7_10_1_b.polarity_test_ieee_c57_13_1}
                    onChange={(e) => setFormData({
                      ...formData,
                      electrical_tests_7_10_1_b: {
                        ...formData.electrical_tests_7_10_1_b,
                        polarity_test_ieee_c57_13_1: e.target.value
                      }
                    })}
                    className="w-full px-3 py-2 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-xs font-semibold"
                  >
                    <option value="Pass (H1-X1 correct polarity verified on all 3 phases)">Pass - Đúng cực tính H1-X1 cả 3 pha</option>
                    <option value="Fail (Reversed polarity detected on Phase C)">Fail - Ngược cực tính pha C (Nguy hiểm rơ-le hướng)</option>
                    <option value="Fail (Reversed polarity detected on Phase A)">Fail - Ngược cực tính pha A</option>
                  </select>
                </div>

                {/* B6: Secondary Burden */}
                <div className="p-4 rounded-xl bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-bold text-slate-800 dark:text-slate-200">
                      B6. {isVi ? 'Đo Phụ Tải Mạch Dòng (Burden)' : 'Secondary Burden (7.10.1.D.6)'}
                    </span>
                    <span className="text-xs font-mono text-indigo-600 dark:text-indigo-400">
                      VA ≤ Rated
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-3">
                    {isVi ? 'Phụ tải thực tế (VA hoặc Ohms) của rơ-le và cáp mạch dòng KHÔNG ĐƯỢC vượt công suất định mức.' : 'Measured loop burden (VA) MUST NOT exceed CT rated burden capacity.'}
                  </p>

                  <div className="space-y-2 text-xs">
                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Burden đo thực tế Pha C (VA):' : 'Measured Burden (VA):'}
                      </label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_1_b.secondary_burden_measurement.measured_burden_phase_c_VA}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_1_b: {
                            ...formData.electrical_tests_7_10_1_b,
                            secondary_burden_measurement: {
                              ...formData.electrical_tests_7_10_1_b.secondary_burden_measurement,
                              measured_burden_phase_c_VA: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Công suất định mức CT (Rated VA):' : 'Rated Burden Capacity (VA):'}
                      </label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_1_b.secondary_burden_measurement.ct_rated_burden_VA}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_1_b: {
                            ...formData.electrical_tests_7_10_1_b,
                            secondary_burden_measurement: {
                              ...formData.electrical_tests_7_10_1_b.secondary_burden_measurement,
                              ct_rated_burden_VA: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>
                  </div>
                </div>

                {/* B3: Ratio Verification (3 Phases) */}
                <div className="p-4 rounded-xl bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60 lg:col-span-3">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-bold text-slate-800 dark:text-slate-200">
                      B3. {isVi ? 'Thử Tỉ Số Biến Dòng (Ratio-Verification per IEEE C57.13)' : 'Ratio-Verification Test (IEEE C57.13)'}
                    </span>
                    <span className="text-xs px-2 py-0.5 rounded bg-blue-100 text-blue-800 dark:bg-blue-900/40 dark:text-blue-300">
                      Cấp 0.3 / Sai số ≤ ±0.3%
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-3">
                    {isVi ? 'Bơm dòng sơ cấp định mức và đo dòng thứ cấp các pha A, B, C. Sai số tỉ số (Ratio error) phải nằm trong cấp chính xác định mức.' : 'Apply primary current and measure secondary outputs on all phases. Ratio error must stay within accuracy class.'}
                  </p>

                  <div className="grid grid-cols-1 md:grid-cols-4 gap-4 text-xs">
                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Dòng sơ cấp áp vào (A):' : 'Applied Primary (A):'}
                      </label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_1_b.ratio_verification_test_ieee_c57_13.applied_primary_current_A}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_1_b: {
                            ...formData.electrical_tests_7_10_1_b,
                            ratio_verification_test_ieee_c57_13: {
                              ...formData.electrical_tests_7_10_1_b.ratio_verification_test_ieee_c57_13,
                              applied_primary_current_A: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Dòng thứ cấp Pha A (A):' : 'Sec Current Phase A (A):'}
                      </label>
                      <input
                        type="number"
                        step="0.001"
                        value={formData.electrical_tests_7_10_1_b.ratio_verification_test_ieee_c57_13.measured_secondary_current_phase_a_A}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_1_b: {
                            ...formData.electrical_tests_7_10_1_b,
                            ratio_verification_test_ieee_c57_13: {
                              ...formData.electrical_tests_7_10_1_b.ratio_verification_test_ieee_c57_13,
                              measured_secondary_current_phase_a_A: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Dòng thứ cấp Pha B (A):' : 'Sec Current Phase B (A):'}
                      </label>
                      <input
                        type="number"
                        step="0.001"
                        value={formData.electrical_tests_7_10_1_b.ratio_verification_test_ieee_c57_13.measured_secondary_current_phase_b_A}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_1_b: {
                            ...formData.electrical_tests_7_10_1_b,
                            ratio_verification_test_ieee_c57_13: {
                              ...formData.electrical_tests_7_10_1_b.ratio_verification_test_ieee_c57_13,
                              measured_secondary_current_phase_b_A: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Dòng thứ cấp Pha C (A) & Sai số %:' : 'Sec Phase C (A) & Error %:'}
                      </label>
                      <div className="flex gap-2">
                        <input
                          type="number"
                          step="0.001"
                          value={formData.electrical_tests_7_10_1_b.ratio_verification_test_ieee_c57_13.measured_secondary_current_phase_c_A}
                          onChange={(e) => {
                            const val = parseFloat(e.target.value) || 0;
                            // auto calculate error assuming 5A rated
                            const err = ((val - 5.0) / 5.0) * 100;
                            setFormData({
                              ...formData,
                              electrical_tests_7_10_1_b: {
                                ...formData.electrical_tests_7_10_1_b,
                                ratio_verification_test_ieee_c57_13: {
                                  ...formData.electrical_tests_7_10_1_b.ratio_verification_test_ieee_c57_13,
                                  measured_secondary_current_phase_c_A: val,
                                  ratio_error_phase_c_percent: parseFloat(err.toFixed(2))
                                }
                              }
                            });
                          }}
                          className="w-1/2 px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                        />
                        <div className={`w-1/2 px-2 py-1.5 rounded text-center font-bold font-mono ${
                          Math.abs(formData.electrical_tests_7_10_1_b.ratio_verification_test_ieee_c57_13.ratio_error_phase_c_percent) <= 0.3
                            ? 'bg-emerald-100 text-emerald-800 dark:bg-emerald-950/60 dark:text-emerald-300'
                            : 'bg-rose-100 text-rose-800 dark:bg-rose-950/60 dark:text-rose-300'
                        }`}>
                          {formData.electrical_tests_7_10_1_b.ratio_verification_test_ieee_c57_13.ratio_error_phase_c_percent}%
                        </div>
                      </div>
                    </div>
                  </div>
                </div>

                {/* B4: Excitation Knee-Point Test */}
                <div className="p-4 rounded-xl bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60 lg:col-span-3">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-bold text-slate-800 dark:text-slate-200">
                      B4. {isVi ? 'Thử Đặc Tính Từ Hóa & Điểm Bão Hòa Knee-Point (IEEE C57.13.1)' : 'Excitation Saturation & Knee-Point Test (IEEE C57.13.1)'}
                    </span>
                    <span className="text-xs px-2 py-0.5 rounded bg-purple-100 text-purple-800 dark:bg-purple-900/40 dark:text-purple-300">
                      C400 (Knee-point ≥ 400V)
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-3">
                    {isVi ? 'Bơm điện áp AC thứ cấp tăng dần để xác định điện áp Knee-point. Nếu điểm Knee-point sụt giảm sâu (ví dụ < 200V đối với C400), CT bị bão hòa sớm và chập vòng nội bộ.' : 'Increase AC secondary voltage to locate knee-point. Severely depressed knee-point indicates shorted turns and premature saturation.'}
                  </p>

                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Pha A - Điện áp Knee-point (V):' : 'Phase A Knee-Point (V):'}
                      </label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_1_b.excitation_saturation_test_ieee_c57_13_1.phase_a_knee_point_voltage_V}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_1_b: {
                            ...formData.electrical_tests_7_10_1_b,
                            excitation_saturation_test_ieee_c57_13_1: {
                              ...formData.electrical_tests_7_10_1_b.excitation_saturation_test_ieee_c57_13_1,
                              phase_a_knee_point_voltage_V: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Pha B - Điện áp Knee-point (V):' : 'Phase B Knee-Point (V):'}
                      </label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_1_b.excitation_saturation_test_ieee_c57_13_1.phase_b_knee_point_voltage_V}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_1_b: {
                            ...formData.electrical_tests_7_10_1_b,
                            excitation_saturation_test_ieee_c57_13_1: {
                              ...formData.electrical_tests_7_10_1_b.excitation_saturation_test_ieee_c57_13_1,
                              phase_b_knee_point_voltage_V: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Pha C - Điện áp Knee-point (V):' : 'Phase C Knee-Point (V):'}
                      </label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_1_b.excitation_saturation_test_ieee_c57_13_1.phase_c_knee_point_voltage_V}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_1_b: {
                            ...formData.electrical_tests_7_10_1_b,
                            excitation_saturation_test_ieee_c57_13_1: {
                              ...formData.electrical_tests_7_10_1_b.excitation_saturation_test_ieee_c57_13_1,
                              phase_c_knee_point_voltage_V: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className={`w-full px-3 py-1.5 rounded border font-mono font-bold ${
                          formData.electrical_tests_7_10_1_b.excitation_saturation_test_ieee_c57_13_1.phase_c_knee_point_voltage_V >= 380
                            ? 'border-emerald-500 text-emerald-700 bg-emerald-50 dark:bg-emerald-950/40'
                            : 'border-rose-500 text-rose-700 bg-rose-50 dark:bg-rose-950/40'
                        }`}
                      />
                    </div>
                  </div>
                </div>

                {/* B5: Winding DC Resistance */}
                <div className="p-4 rounded-xl bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60 lg:col-span-3">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-bold text-slate-800 dark:text-slate-200">
                      B5. {isVi ? 'Điện Trở Một Chiều Cuộn Dây (Winding DC Resistance per IEEE C57.13.1)' : 'Winding DC Resistance (IEEE C57.13.1)'}
                    </span>
                    <span className={`text-xs px-2 py-0.5 rounded font-bold ${
                      phaseDeviation <= 5.0
                        ? 'bg-emerald-100 text-emerald-800 dark:bg-emerald-900/40 dark:text-emerald-300'
                        : 'bg-rose-100 text-rose-800 dark:bg-rose-900/40 dark:text-rose-300'
                    }`}>
                      Độ lệch pha: {phaseDeviation.toFixed(1)}% (Chuẩn ≤ 5%)
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-3">
                    {isVi ? 'Đo điện trở cuộn dây thứ cấp bằng Micro-ohmmeter. Độ lệch giữa các pha KHÔNG ĐƯỢC vượt quá 5%. Nếu sụt giảm > 15-30% là dấu hiệu chập vòng dây nội bộ.' : 'Measure secondary winding resistance. Phase-to-phase deviation must be <= 5%. A drop > 15-30% confirms internal shorted turns.'}
                  </p>

                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Pha A (Ohms):' : 'Phase A (Ohms):'}
                      </label>
                      <input
                        type="number"
                        step="0.001"
                        value={formData.electrical_tests_7_10_1_b.winding_dc_resistance_ohms.phase_a_ohms}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_1_b: {
                            ...formData.electrical_tests_7_10_1_b,
                            winding_dc_resistance_ohms: {
                              ...formData.electrical_tests_7_10_1_b.winding_dc_resistance_ohms,
                              phase_a_ohms: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Pha B (Ohms):' : 'Phase B (Ohms):'}
                      </label>
                      <input
                        type="number"
                        step="0.001"
                        value={formData.electrical_tests_7_10_1_b.winding_dc_resistance_ohms.phase_b_ohms}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_1_b: {
                            ...formData.electrical_tests_7_10_1_b,
                            winding_dc_resistance_ohms: {
                              ...formData.electrical_tests_7_10_1_b.winding_dc_resistance_ohms,
                              phase_b_ohms: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Pha C (Ohms):' : 'Phase C (Ohms):'}
                      </label>
                      <input
                        type="number"
                        step="0.001"
                        value={formData.electrical_tests_7_10_1_b.winding_dc_resistance_ohms.phase_c_ohms}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_1_b: {
                            ...formData.electrical_tests_7_10_1_b,
                            winding_dc_resistance_ohms: {
                              ...formData.electrical_tests_7_10_1_b.winding_dc_resistance_ohms,
                              phase_c_ohms: parseFloat(e.target.value) || 0
                            }
                          }
                        })}
                        className={`w-full px-3 py-1.5 rounded border font-mono font-bold ${
                          formData.electrical_tests_7_10_1_b.winding_dc_resistance_ohms.phase_c_ohms >= 0.38
                            ? 'border-emerald-500 text-emerald-700 bg-emerald-50 dark:bg-emerald-950/40'
                            : 'border-rose-500 text-rose-700 bg-rose-50 dark:bg-rose-950/40'
                        }`}
                      />
                    </div>
                  </div>
                </div>

                {/* B7: Dielectric Withstand */}
                <div className="p-4 rounded-xl bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60 lg:col-span-3">
                  <div className="flex items-center justify-between mb-2">
                    <span className="font-bold text-slate-800 dark:text-slate-200">
                      B7. {isVi ? 'Thử Chịu Đựng Điện Áp Tăng Cao (Dielectric Withstand Table 100.9)' : 'Dielectric Withstand Test (Table 100.9)'}
                    </span>
                    <span className="text-xs font-mono text-indigo-600 dark:text-indigo-400">
                      Table 100.9 (27kV AC 1 min)
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mb-3">
                    {isVi ? 'Áp dụng điện áp thử tần số công nghiệp lên cuộn sơ cấp trong 1 phút (ví dụ hệ 15kV thử 27kV AC). Không được có hiện tượng đánh thủng cách điện.' : 'Apply rated test voltage for 1 minute to primary winding without breakdown.'}
                  </p>

                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Điện áp thử nghiệm (kV AC):' : 'Test Voltage (kV AC):'}
                      </label>
                      <input
                        type="number"
                        value={formData.electrical_tests_7_10_1_b.dielectric_withstand_test_table_100_9.test_voltage_kV_ac}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_1_b: {
                            ...formData.electrical_tests_7_10_1_b,
                            dielectric_withstand_test_table_100_9: {
                              ...formData.electrical_tests_7_10_1_b.dielectric_withstand_test_table_100_9,
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
                        value={formData.electrical_tests_7_10_1_b.dielectric_withstand_test_table_100_9.duration_minutes}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_1_b: {
                            ...formData.electrical_tests_7_10_1_b,
                            dielectric_withstand_test_table_100_9: {
                              ...formData.electrical_tests_7_10_1_b.dielectric_withstand_test_table_100_9,
                              duration_minutes: parseInt(e.target.value) || 1
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-slate-500 dark:text-slate-400 mb-1">
                        {isVi ? 'Kết quả thử nghiệm:' : 'Result Status:'}
                      </label>
                      <input
                        type="text"
                        value={formData.electrical_tests_7_10_1_b.dielectric_withstand_test_table_100_9.status}
                        onChange={(e) => setFormData({
                          ...formData,
                          electrical_tests_7_10_1_b: {
                            ...formData.electrical_tests_7_10_1_b,
                            dielectric_withstand_test_table_100_9: {
                              ...formData.electrical_tests_7_10_1_b.dielectric_withstand_test_table_100_9,
                              status: e.target.value
                            }
                          }
                        })}
                        className="w-full px-3 py-1.5 rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-xs font-semibold"
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
                  <div className="w-8 h-8 rounded-lg bg-indigo-100 dark:bg-indigo-900/40 text-indigo-600 dark:text-indigo-400 flex items-center justify-center font-bold">
                    C
                  </div>
                  <div>
                    <h4 className="font-bold text-slate-900 dark:text-slate-100 text-sm">
                      {isVi ? 'Khảo Sát Hồng Ngoại & Hoàn Tất Chạy Thử (Section 9 / Table 100.18)' : 'Post-Test Survey & Thermography (Table 100.18)'}
                    </h4>
                    <p className="text-xs text-slate-500">
                      {isVi ? 'Khảo sát ảnh nhiệt IR dưới tải định mức và hoàn tất thủ tục bàn giao đóng điện.' : 'Infrared thermographic survey under load and final sign-off.'}
                    </p>
                  </div>
                </div>
              </div>

              <div className="p-4 rounded-xl bg-slate-50 dark:bg-slate-800/50 border border-slate-200 dark:border-slate-700/60">
                <div className="flex items-center justify-between mb-2">
                  <span className="font-bold text-slate-800 dark:text-slate-200 flex items-center gap-2">
                    <Thermometer size={16} className="text-rose-500" />
                    <span>{isVi ? 'Khảo Sát Nhiệt IR Mối Nối & Thân CT (NETA Section 9 / Table 100.18)' : 'Thermographic Survey (Table 100.18)'}</span>
                  </span>
                </div>
                <p className="text-xs text-slate-500 mb-3">
                  {isVi ? 'Kiểm tra phát hiện điểm quá nhiệt bất thường (Hotspot) tại đầu cốt sơ cấp H1, cầu đấu thứ cấp X1-X2 và bề mặt cách điện.' : 'Check for abnormal hotspots at primary H1 terminals, secondary X1-X2 blocks, and resin body.'}
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

              {/* Action bar inside tab */}
              <div className="pt-4 flex justify-end">
                <button
                  onClick={handleRunEvaluation}
                  disabled={isEvaluating}
                  className="flex items-center gap-2 px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-sm shadow-md transition"
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
              <div className="w-12 h-12 border-4 border-indigo-600 border-t-transparent rounded-full animate-spin mx-auto mb-4" />
              <h3 className="text-lg font-bold text-slate-900 dark:text-slate-100">
                {isVi ? 'Đang Phân Tích Dữ Liệu Theo ANSI/NETA MTS-2023 Mục 7.10.1...' : 'Analyzing CT Data per ANSI/NETA MTS-2023 Sec 7.10.1...'}
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
                  <FileCheck2 size={18} className="text-indigo-600 dark:text-indigo-400" />
                  <span>
                    {isVi
                      ? 'Bảng Phân Tích Chi Tiết Bảo Trì Biến Dòng Điện CT (NETA MTS Mục 7.10.1 & NETA Tables)'
                      : 'Detailed CT Maintenance Evaluation Table (NETA MTS Sec 7.10.1 & Tables)'}
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
                  <Activity size={18} className="text-indigo-600 dark:text-indigo-400" />
                  <span>
                    {isVi
                      ? 'Phân Tích Kỹ Thuật & Nguy Cơ Tự Mất Bảo Vệ Rơ-le (Technical Analysis & Hazard)'
                      : 'Technical Analysis & Protection Relay Maloperation Hazard'}
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
                <div className="bg-slate-950 border border-indigo-900/60 rounded-xl p-5 text-slate-200 text-xs font-mono space-y-2 shadow-inner">
                  <div className="flex items-center justify-between text-indigo-400 font-semibold border-b border-slate-800 pb-2">
                    <span className="flex items-center gap-1.5">
                      <Sparkles size={14} />
                      <span>Gemini 2.5 Flash Raw Analytical Output</span>
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
              <Sparkles size={40} className="text-indigo-400 mx-auto mb-3" />
              <h3 className="text-base font-bold text-slate-800 dark:text-slate-200">
                {isVi ? 'Chưa Có Kết Quả Chẩn Đoán AI' : 'No AI Diagnostic Results Yet'}
              </h3>
              <p className="text-xs text-slate-500 mt-1 max-w-md mx-auto">
                {isVi
                  ? 'Bấm nút "Chẩn Đoán NETA MTS AI" ở trên để phân tích dữ liệu kiểm tra biến dòng điện CT theo đúng thuật toán NETA MTS Mục 7.10.1.'
                  : 'Click "Run NETA MTS AI Analysis" above to evaluate instrument current transformer data.'}
              </p>
              <button
                onClick={handleRunEvaluation}
                className="mt-4 px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition"
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
                <FileText size={18} className="text-indigo-600 dark:text-indigo-400" />
                <span>{isVi ? 'System Instruction Prompt Dành Cho Google AI Studio (TEV COE 7.10.1 CT)' : 'Google AI Studio System Instruction Prompt'}</span>
              </h3>
              <p className="text-xs text-slate-500">
                {isVi ? 'Copy đoạn văn bản dưới đây và dán vào mục System Instructions trên Google AI Studio' : 'Copy and paste into System Instructions field in Google AI Studio'}
              </p>
            </div>

            <button
              onClick={handleCopyPrompt}
              className="flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold shadow transition"
            >
              {copySuccess ? <Check size={14} /> : <Copy size={14} />}
              <span>{copySuccess ? (isVi ? 'Đã Copy!' : 'Copied!') : (isVi ? 'Copy System Prompt' : 'Copy System Prompt')}</span>
            </button>
          </div>

          <div className="bg-slate-950 text-slate-200 p-4 rounded-xl text-xs font-mono max-h-[500px] overflow-y-auto border border-slate-800 leading-relaxed">
            <pre className="whitespace-pre-wrap">{MTS_CT_SYSTEM_INSTRUCTION}</pre>
          </div>
        </div>
      )}

      {/* TAB 4: JSON INPUT SCHEMA */}
      {activeTab === 'json' && (
        <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-5 shadow-sm space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-200 dark:border-slate-800 pb-3">
            <div>
              <h3 className="font-bold text-slate-900 dark:text-slate-100 text-sm flex items-center gap-2">
                <FileCheck2 size={18} className="text-indigo-600 dark:text-indigo-400" />
                <span>{isVi ? 'Mẫu JSON User Prompt Đầu Vào Dùng Cho Google AI Studio' : 'JSON User Prompt Payload for Google AI Studio'}</span>
              </h3>
              <p className="text-xs text-slate-500">
                {isVi ? 'Dữ liệu JSON mô phỏng đầy đủ các phép đo theo ANSI/NETA MTS-2023 Mục 7.10.1' : 'Complete JSON structure per ANSI/NETA MTS-2023 Section 7.10.1'}
              </p>
            </div>

            <button
              onClick={handleCopyJson}
              className="flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold shadow transition"
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
