import React, { useState, useMemo } from 'react';
import {
  ShieldAlert,
  ShieldCheck,
  AlertTriangle,
  AlertOctagon,
  CheckCircle2,
  XCircle,
  Sparkles,
  Download,
  Copy,
  Check,
  RotateCcw,
  Info,
  Layers,
  FileCheck2,
  ChevronDown,
  ChevronUp,
  Clock,
  ArrowRight,
  Eye,
  SlidersHorizontal,
  Save,
  Radio,
  Zap,
  Gauge,
  Thermometer,
  Flame,
  Globe,
  ClipboardCheck,
  Search,
  CheckSquare,
  Square
} from 'lucide-react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  NetaEcsLvInputPayload,
  NetaEcsLvEvaluationResult,
  performDeterministicEcsLvAnalysis,
  ECS_LV_SYSTEM_INSTRUCTION
} from '../../server/netaEcsLowVoltageAnalyzer';
import { FseEngineerDropdown } from './FseEngineerDropdown';
import { CustomerSelectDropdown } from './CustomerSelectDropdown';
import { ProjectSelectDropdown } from './ProjectSelectDropdown';
import { EquipmentSelectDropdown } from './EquipmentSelectDropdown';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
  initialLanguage?: 'vi' | 'en';
}

// ----------------------------------------------------------------------------
// PRESET COMMISSIONING SCENARIOS
// ----------------------------------------------------------------------------
interface PresetScenario {
  id: string;
  name: string;
  nameEn: string;
  badge: string;
  badgeColor: string;
  systemTag: string;
  insA: number;
  insB: number;
  insC: number;
  groundingR: number;
  isTorquePass: boolean;
  isGroundsRemoved: boolean;
  isPhaseSeqPass: boolean;
  isThermalPass: boolean;
  loadA: number;
  loadB: number;
  loadC: number;
  isGfciPass: boolean;
  descriptionVn: string;
}

const PRESET_SCENARIOS: PresetScenario[] = [
  {
    id: 'preset-phase-c-fail',
    name: 'Sự cố Cách điện Pha C Tủ MSB-01 (0.45 MΩ < 1.0 MΩ - Đề bài)',
    nameEn: 'Phase C Low Insulation Hazard (0.45 MΩ < 1.0 MΩ - Sample Defect)',
    badge: 'CRITICAL / NGUY HẠI ĐÓNG ĐIỆN',
    badgeColor: 'bg-rose-500/20 text-rose-300 border-rose-400/40',
    systemTag: 'LV-SWITCHBOARD-MSB-01 (< 1 kV System)',
    insA: 185.0,
    insB: 190.0,
    insC: 0.45,
    groundingR: 0.85,
    isTorquePass: true,
    isGroundsRemoved: true,
    isPhaseSeqPass: true,
    isThermalPass: true,
    loadA: 420,
    loadB: 415,
    loadC: 418,
    isGfciPass: true,
    descriptionVn: 'Cách điện Pha C chỉ đạt 0.45 MΩ, dưới ngưỡng 1.0 MΩ (Table 100.1). TẠM DỪNG ĐÓNG ĐIỆN NGAY LẬP TỨC để khoanh vùng xử lý ẩm/sứ đỡ!'
  },
  {
    id: 'preset-all-pass',
    name: 'Nghiệm thu Đạt Chuẩn 100% (All Pass - Ready for Commercial Operation)',
    nameEn: '100% Compliance Verified (All Pass - Ready for Commercial Operation)',
    badge: 'COMMISSIONED / SẴN SÀNG',
    badgeColor: 'bg-emerald-500/20 text-emerald-300 border-emerald-400/40',
    systemTag: 'LV-MCC-01 (Motor Control Center 400V)',
    insA: 250.0,
    insB: 240.0,
    insC: 245.0,
    groundingR: 0.65,
    isTorquePass: true,
    isGroundsRemoved: true,
    isPhaseSeqPass: true,
    isThermalPass: true,
    loadA: 310,
    loadB: 305,
    loadC: 308,
    isGfciPass: true,
    descriptionVn: 'Tất cả 13 hạng mục NETA ECS 7.1 đều đạt xuất sắc (Cách điện > 200 MΩ, tiếp địa 0.65Ω, tải cân bằng, không có điểm nóng IR).'
  },
  {
    id: 'preset-caution-torque-unbalance',
    name: 'Cảnh báo Lực Siết & Mất Cân Bằng Tải Pha (Punch-List Required)',
    nameEn: 'Caution: Torque Deficiency & Load Unbalance (Punch-List Required)',
    badge: 'CAUTION / CẦN KHẮC PHỤC',
    badgeColor: 'bg-amber-500/20 text-amber-300 border-amber-400/40',
    systemTag: 'LV-DB-PRODUCTION-02 (Distribution Board 400V)',
    insA: 85.0,
    insB: 92.0,
    insC: 78.0,
    groundingR: 1.2,
    isTorquePass: false,
    isGroundsRemoved: true,
    isPhaseSeqPass: true,
    isThermalPass: false,
    loadA: 450,
    loadB: 320,
    loadC: 430,
    isGfciPass: true,
    descriptionVn: 'Chưa siết đủ lực bu-lông, độ lệch dòng pha 23.5% (> 15%) và xuất hiện phát nhiệt nhẹ. Cần lập Punch-List hoàn thiện trong 48h.'
  }
];

export const NetaEcsLowVoltageChecklist: React.FC<Props> = ({
  onSaveReport,
  onNavigateToReports,
  initialLanguage = 'vi'
}) => {
  const [lang, setLang] = useState<'vi' | 'en'>(initialLanguage);
  const isVi = lang === 'vi';

  // --------------------------------------------------------------------------
  // HEADER & METADATA STATE
  // --------------------------------------------------------------------------
  const [reportNo, setReportNo] = useState(`TEV-ECS-2024-LV-${Math.floor(100 + Math.random() * 900)}`);
  const [customerName, setCustomerName] = useState('Công ty Cổ phần Năng lượng TEV');
  const [projectName, setProjectName] = useState('Nha May Cong Nghe Cao TEV High-Tech Plant');
  const [systemTag, setSystemTag] = useState('LV-SWITCHBOARD-MSB-01 (< 1 kV System)');
  const [voltageRatingV, setVoltageRatingV] = useState(400);
  const [fseName, setFseName] = useState('Nguyen Van H');
  const [commissioningDate, setCommissioningDate] = useState(new Date().toISOString().split('T')[0]);

  // --------------------------------------------------------------------------
  // SECTION 7.1.A: PRE-ENERGIZATION STATE
  // --------------------------------------------------------------------------
  const [acceptanceTestReportsReviewed, setAcceptanceTestReportsReviewed] = useState(true);
  const [insulationPhaseA, setInsulationPhaseA] = useState(185.0);
  const [insulationPhaseB, setInsulationPhaseB] = useState(190.0);
  const [insulationPhaseC, setInsulationPhaseC] = useState(0.45);
  const [boltedTorqueStatus, setBoltedTorqueStatus] = useState<'Pass' | 'Fail'>('Pass');
  const [groundingResistanceOhms, setGroundingResistanceOhms] = useState(0.85);
  const [safetyInterlocksEstop, setSafetyInterlocksEstop] = useState<'Pass' | 'Fail'>('Pass');
  const [writtenEnergizationPlanApproved, setWrittenEnergizationPlanApproved] = useState(true);
  const [dcControlBatteryReady, setDcControlBatteryReady] = useState(true);

  // --------------------------------------------------------------------------
  // SECTION 7.1.B: ENERGIZATION STATE
  // --------------------------------------------------------------------------
  const [tempProtectiveGroundsRemoved, setTempProtectiveGroundsRemoved] = useState(true);
  const [safetyBarricadesRestricted, setSafetyBarricadesRestricted] = useState(true);
  const [phaseSequenceCheck, setPhaseSequenceCheck] = useState<'Pass' | 'Fail'>('Pass');
  const [motorRotationCheck, setMotorRotationCheck] = useState<'Pass' | 'Fail'>('Pass');
  const [alarmsAndMeteringStatus, setAlarmsAndMeteringStatus] = useState<'Pass' | 'Fail'>('Pass');

  // --------------------------------------------------------------------------
  // SECTION 7.1.C: POST-ENERGIZATION STATE
  // --------------------------------------------------------------------------
  const [thermographicSurveyStatus, setThermographicSurveyStatus] = useState<'Pass' | 'Investigate'>('Pass');
  const [loadCurrentPhaseA, setLoadCurrentPhaseA] = useState(420);
  const [loadCurrentPhaseB, setLoadCurrentPhaseB] = useState(415);
  const [loadCurrentPhaseC, setLoadCurrentPhaseC] = useState(418);
  const [gfciReceptaclesTested, setGfciReceptaclesTested] = useState<'Pass' | 'Fail'>('Pass');

  // Active accordion tabs
  const [activeStageTab, setActiveStageTab] = useState<'all' | 'stageA' | 'stageB' | 'stageC'>('all');
  const [showPromptModal, setShowPromptModal] = useState(false);
  const [copiedPrompt, setCopiedPrompt] = useState(false);
  const [saveSuccessMessage, setSaveSuccessMessage] = useState<string | null>(null);
  const [isAiRunning, setIsAiRunning] = useState(false);

  // Computed phase unbalance %
  const phaseUnbalance = useMemo(() => {
    const avg = (loadCurrentPhaseA + loadCurrentPhaseB + loadCurrentPhaseC) / 3;
    if (avg <= 0) return 0;
    const maxDiff = Math.max(
      Math.abs(loadCurrentPhaseA - avg),
      Math.abs(loadCurrentPhaseB - avg),
      Math.abs(loadCurrentPhaseC - avg)
    );
    return Number(((maxDiff / avg) * 100).toFixed(1));
  }, [loadCurrentPhaseA, loadCurrentPhaseB, loadCurrentPhaseC]);

  // Current payload
  const currentPayload: NetaEcsLvInputPayload = useMemo(() => {
    const minIns = Math.min(insulationPhaseA, insulationPhaseB, insulationPhaseC);
    return {
      tev_coe_header: {
        module: 'LV-COMMISSIONING-ECS-7.1',
        project_name: projectName,
        system_tag: systemTag,
        voltage_rating_V: voltageRatingV,
        fse_name: fseName,
        commissioning_date: commissioningDate
      },
      pre_energization_7_1_a: {
        acceptance_test_reports_reviewed: acceptanceTestReportsReviewed,
        busbar_insulation_resistance_1000v_megohms: {
          phase_a_to_ground: insulationPhaseA,
          phase_b_to_ground: insulationPhaseB,
          phase_c_to_ground: insulationPhaseC,
          status: minIns < 1.0 ? `Fail (Min ${minIns} MΩ < 1.0 MΩ NETA Table 100.1 threshold)` : 'Pass (>= 1.0 MΩ)'
        },
        bolted_connection_torque_table_100_12: boltedTorqueStatus === 'Pass'
          ? 'Pass (Calibrated torque wrench verified per Table 100.12)'
          : 'Fail (Torque verification missing or deficient)',
        grounding_resistance_ohms: groundingResistanceOhms,
        safety_interlocks_and_estop: safetyInterlocksEstop === 'Pass'
          ? 'Pass (E-stop trips main breaker)'
          : 'Fail (Interlocks or E-stop failure)',
        written_energization_plan_approved: writtenEnergizationPlanApproved,
        dc_control_power_battery_ready: dcControlBatteryReady
      },
      energization_7_1_b: {
        temporary_protective_grounds_removed: tempProtectiveGroundsRemoved,
        safety_barricades_access_restricted: safetyBarricadesRestricted,
        phase_sequence_check: phaseSequenceCheck === 'Pass'
          ? 'Pass (A-B-C correct phase sequence)'
          : 'Fail (Reverse phase sequence)',
        motor_rotation_check: motorRotationCheck === 'Pass'
          ? 'Pass (Correct rotation)'
          : 'Fail (Reverse rotation)',
        alarms_and_metering_status: alarmsAndMeteringStatus === 'Pass'
          ? 'Pass (0 active fault alarms)'
          : 'Fail (Active alarms)'
      },
      post_energization_7_1_c: {
        thermographic_survey_table_100_18: thermographicSurveyStatus === 'Pass'
          ? 'Pass (Normal temperature distribution under load)'
          : 'Investigate (Hotspot observed per Table 100.18)',
        phase_load_balance_amps: {
          phase_a: loadCurrentPhaseA,
          phase_b: loadCurrentPhaseB,
          phase_c: loadCurrentPhaseC
        },
        gfci_receptacles_tested: gfciReceptaclesTested === 'Pass'
          ? 'Pass (100% trip functional)'
          : 'Fail (GFCI failure detected)'
      }
    };
  }, [
    projectName, systemTag, voltageRatingV, fseName, commissioningDate,
    acceptanceTestReportsReviewed, insulationPhaseA, insulationPhaseB, insulationPhaseC,
    boltedTorqueStatus, groundingResistanceOhms, safetyInterlocksEstop, writtenEnergizationPlanApproved, dcControlBatteryReady,
    tempProtectiveGroundsRemoved, safetyBarricadesRestricted, phaseSequenceCheck, motorRotationCheck, alarmsAndMeteringStatus,
    thermographicSurveyStatus, loadCurrentPhaseA, loadCurrentPhaseB, loadCurrentPhaseC, gfciReceptaclesTested
  ]);

  // Evaluated result
  const evaluationResult: NetaEcsLvEvaluationResult = useMemo(() => {
    return performDeterministicEcsLvAnalysis(currentPayload);
  }, [currentPayload]);

  // --------------------------------------------------------------------------
  // PRESET LOADER
  // --------------------------------------------------------------------------
  const handleLoadPreset = (scenario: PresetScenario) => {
    setSystemTag(scenario.systemTag);
    setInsulationPhaseA(scenario.insA);
    setInsulationPhaseB(scenario.insB);
    setInsulationPhaseC(scenario.insC);
    setGroundingResistanceOhms(scenario.groundingR);
    setBoltedTorqueStatus(scenario.isTorquePass ? 'Pass' : 'Fail');
    setTempProtectiveGroundsRemoved(scenario.isGroundsRemoved);
    setPhaseSequenceCheck(scenario.isPhaseSeqPass ? 'Pass' : 'Fail');
    setThermographicSurveyStatus(scenario.isThermalPass ? 'Pass' : 'Investigate');
    setLoadCurrentPhaseA(scenario.loadA);
    setLoadCurrentPhaseB(scenario.loadB);
    setLoadCurrentPhaseC(scenario.loadC);
    setGfciReceptaclesTested(scenario.isGfciPass ? 'Pass' : 'Fail');
  };

  // --------------------------------------------------------------------------
  // AI DIAGNOSTIC HANDLER
  // --------------------------------------------------------------------------
  const handleRunAiEvaluation = async () => {
    setIsAiRunning(true);
    try {
      const res = await fetch('/api/field-service/neta-ecs-lv-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(currentPayload)
      });
      const json = await res.json();
      if (json.success) {
        setSaveSuccessMessage(isVi ? 'Đã hoàn tất thẩm định AI theo chuẩn ANSI/NETA ECS-2024!' : 'AI Commissioning Evaluation Completed!');
        setTimeout(() => setSaveSuccessMessage(null), 3500);
      }
    } catch (e) {
      console.warn('AI analysis call:', e);
    } finally {
      setIsAiRunning(false);
    }
  };

  // --------------------------------------------------------------------------
  // SAVE REPORT TO CMMS / ARCHIVE
  // --------------------------------------------------------------------------
  const handleSaveToReports = () => {
    if (onSaveReport) {
      onSaveReport({
        data: currentPayload,
        analysisResult: evaluationResult,
        equipmentId: systemTag,
        date: commissioningDate,
        type: 'ANSI/NETA ECS-2024 LV Commissioning Report'
      });
    }
    setSaveSuccessMessage(isVi ? `Đã lưu hồ sơ nghiệm thu chạy thử [${systemTag}] vào CMMS!` : `Commissioning record for [${systemTag}] archived!`);
    setTimeout(() => setSaveSuccessMessage(null), 3500);
  };

  // --------------------------------------------------------------------------
  // A4 PDF EXPORT (Chuẩn TEV COE)
  // --------------------------------------------------------------------------
  const handleExportA4Pdf = () => {
    const doc = new jsPDF('p', 'mm', 'a4');
    const pageWidth = doc.internal.pageSize.getWidth();

    // TEV Corporate Header
    doc.setFillColor(15, 23, 42); // slate-900
    doc.rect(0, 0, pageWidth, 28, 'F');

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(14);
    doc.setFont('helvetica', 'bold');
    doc.text('CÔNG TY CỔ PHẦN KỸ NGHỆ NĂNG LƯỢNG VIỆT NAM (TEV)', 14, 11);

    doc.setFontSize(8);
    doc.setFont('helvetica', 'normal');
    doc.text('TEV CENTER OF EXCELLENCE (COE) • ELECTRICAL COMMISSIONING SPECIFICATIONS (ECS)', 14, 17);
    doc.text(`REPORT NO: ${reportNo} • DATE: ${commissioningDate} • STANDARD: ANSI/NETA ECS-2024 (SEC 7.1)`, 14, 23);

    // Title
    doc.setTextColor(15, 23, 42);
    doc.setFontSize(13);
    doc.setFont('helvetica', 'bold');
    doc.text('BIÊN BẢN NGHIỆM THU CHẠY THỬ & ĐÓNG ĐIỆN HỆ THỐNG HẠ THẾ (< 1 kV)', 14, 37);

    doc.setFontSize(8.5);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(71, 85, 105);
    doc.text('General Inspection and Test Procedures: Low-Voltage Systems (< 1 kV) - NETA ECS-2024 & Tables 100.1, 100.12', 14, 43);

    // Section 1: Project Metadata
    const metaTable = [
      [`Dự án: ${projectName}`, `Kỹ sư FSE (NETA III): ${fseName}`],
      [`Thiết bị / Tủ điện: ${systemTag}`, `Điện áp định mức: ${voltageRatingV} V (Hệ < 1 kV)`],
      [`Ngày chạy thử: ${commissioningDate}`, `Số hiệu biên bản: ${reportNo}`]
    ];

    autoTable(doc, {
      startY: 47,
      body: metaTable,
      theme: 'grid',
      styles: { fontSize: 8, cellPadding: 2, textColor: [30, 41, 59] },
      columnStyles: {
        0: { fontStyle: 'bold', cellWidth: 100 },
        1: { cellWidth: 82 }
      }
    });

    let currentY = (doc as any).lastAutoTable?.finalY || 65;

    // Severity Status Banner in PDF
    const isCritical = evaluationResult.overall_status === 'CRITICAL';
    const isCaution = evaluationResult.overall_status === 'CAUTION';

    doc.setFillColor(isCritical ? 254 : isCaution ? 254 : 236, isCritical ? 226 : isCaution ? 243 : 253, isCritical ? 226 : isCaution ? 199 : 245);
    doc.rect(14, currentY + 3, pageWidth - 28, 14, 'F');
    doc.setDrawColor(isCritical ? 239 : isCaution ? 245 : 16, isCritical ? 68 : isCaution ? 158 : 185, isCritical ? 68 : isCaution ? 11 : 129);
    doc.rect(14, currentY + 3, pageWidth - 28, 14);

    doc.setFontSize(9);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(isCritical ? 185 : isCaution ? 180 : 4, isCritical ? 28 : isCaution ? 83 : 120, isCritical ? 28 : isCaution ? 9 : 87);
    doc.text(`TRẠNG THÁI NGHIỆM THU: ${evaluationResult.summary_badge}`, 18, currentY + 9);
    doc.setFontSize(7.5);
    doc.setFont('helvetica', 'normal');
    doc.text(evaluationResult.status_title_vn, 18, currentY + 14);

    currentY += 21;

    // Evaluation Table
    const tableRows = evaluationResult.evaluation_table.map((row) => [
      row.item_code,
      row.item_name_vn,
      row.neta_standard_ref,
      row.measured_value,
      row.status
    ]);

    autoTable(doc, {
      startY: currentY,
      head: [['Mục', 'Hạng Mục Kiểm Tra & Nghiệm Thu', 'Tiêu Chuẩn Tham Chiếu', 'Giá Trị Đo / Hiện Trường', 'Đánh Giá']],
      body: tableRows,
      theme: 'grid',
      styles: { fontSize: 7, cellPadding: 2, textColor: [30, 41, 59] },
      headStyles: { fillColor: [15, 23, 42], textColor: [255, 255, 255], fontStyle: 'bold' },
      columnStyles: {
        0: { cellWidth: 10, fontStyle: 'bold', halign: 'center' },
        1: { cellWidth: 62 },
        2: { cellWidth: 42, fontSize: 6.5 },
        3: { cellWidth: 50 },
        4: { cellWidth: 18, fontStyle: 'bold', halign: 'center' }
      },
      didParseCell: (data) => {
        if (data.section === 'body' && data.column.index === 4) {
          const val = String(data.cell.raw);
          if (val === 'FAIL') {
            data.cell.styles.fillColor = [254, 226, 226];
            data.cell.styles.textColor = [185, 28, 28];
          } else if (val === 'INVESTIGATE') {
            data.cell.styles.fillColor = [254, 243, 199];
            data.cell.styles.textColor = [180, 83, 9];
          } else {
            data.cell.styles.fillColor = [236, 253, 245];
            data.cell.styles.textColor = [4, 120, 87];
          }
        }
      }
    });

    currentY = (doc as any).lastAutoTable?.finalY || 180;

    // Technical Analysis & Action Recommendations Box
    if (currentY + 60 > 260) {
      doc.addPage();
      currentY = 20;
    }

    doc.setFillColor(248, 250, 252);
    doc.rect(14, currentY + 3, pageWidth - 28, 36, 'F');
    doc.setDrawColor(203, 213, 225);
    doc.rect(14, currentY + 3, pageWidth - 28, 36);

    doc.setFontSize(8.5);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(isCritical ? 185 : 15, isCritical ? 28 : 23, isCritical ? 28 : 42);
    doc.text(`PHÂN TÍCH KỸ THUẬT & HÀNH ĐỘNG KHUYẾN NGHỊ FSE:`, 18, currentY + 9);

    doc.setFontSize(7);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(51, 65, 85);

    const splitDesc = doc.splitTextToSize(evaluationResult.technical_analysis.description_vn, pageWidth - 36);
    doc.text(splitDesc.slice(0, 3), 18, currentY + 15);

    evaluationResult.actionable_recommendations.steps.slice(0, 2).forEach((step, idx) => {
      doc.setFont('helvetica', 'bold');
      doc.text(`• Bước ${step.step_number}: ${step.title_vn}`, 18, currentY + 26 + idx * 5);
      doc.setFont('helvetica', 'normal');
      doc.text(`  ${step.details_vn.slice(0, 115)}...`, 18, currentY + 29 + idx * 5);
    });

    // 3-Party Signature Block
    const sigY = 262;
    doc.setDrawColor(203, 213, 225);
    doc.line(14, sigY, pageWidth - 14, sigY);

    doc.setFontSize(8);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(15, 23, 42);
    doc.text('KỸ SƯ NGHIỆM THU FSE (NETA III)', 20, sigY + 6);
    doc.text('GIÁM SÁT KỸ THUẬT NHÀ MÁY', 85, sigY + 6);
    doc.text('ĐẠI DIỆN CHỦ ĐẦU TƯ / BAN QLDA', 150, sigY + 6);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.setTextColor(100, 116, 139);
    doc.text(fseName, 20, sigY + 20);
    doc.text('Kỹ sư Giám sát Nghiệm thu', 85, sigY + 20);
    doc.text(projectName, 150, sigY + 20);

    doc.save(`TEV_ECS_Report_${systemTag.replace(/[^a-zA-Z0-9]/g, '_')}_${commissioningDate}.pdf`);
  };

  return (
    <div className="space-y-6">
      {/* -------------------------------------------------------------------- */}
      {/* 1. TOP HEADER & COMMISSIONING ACTION BAR */}
      {/* -------------------------------------------------------------------- */}
      <div className="bg-gradient-to-r from-slate-950 via-teal-950 to-slate-900 text-white p-5 sm:p-6 rounded-2xl shadow-xl border border-teal-500/30">
        <div className="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-4">
          <div className="space-y-1.5">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="px-2.5 py-1 rounded-md text-[11px] font-black uppercase tracking-wider bg-teal-500/30 text-teal-200 border border-teal-400/30">
                NETA ECS-2024 • MỤC 7.1
              </span>
              <span className="px-2 py-0.5 rounded-md text-[10px] font-bold bg-amber-500/20 text-amber-300 border border-amber-500/30">
                Low-Voltage Systems (&lt; 1 kV)
              </span>
              <span className="px-2 py-0.5 rounded-md text-[10px] font-mono text-slate-300 bg-slate-800/80 border border-slate-700">
                Tables 100.1, 100.12, 100.18
              </span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white flex items-center gap-2.5">
              <ClipboardCheck className="text-teal-400" size={26} />
              <span>
                {isVi
                  ? 'Quy Trình Kiểm Tra & Chạy Thử Hệ Thống Điện Hạ Thế (< 1 kV)'
                  : 'General Inspection & Test Procedures - Low-Voltage Systems (< 1 kV)'}
              </span>
            </h1>
            <p className="text-xs text-slate-300 max-w-3xl leading-relaxed">
              {isVi
                ? 'Quy chuẩn đóng điện nghiệm thu (Commissioning & Acceptance) 3 giai đoạn theo ANSI/NETA ECS-2024: Tiền đóng điện (Pre-Energization), Trong khi đóng điện (Energization) và Sau khi đóng điện (Post-Energization).'
                : '3-stage electrical commissioning & acceptance procedure according to ANSI/NETA ECS-2024: Pre-Energization, Energization, and Post-Energization verification.'}
            </p>
          </div>

          <div className="flex items-center gap-2 flex-wrap self-stretch sm:self-auto">
            {/* Language toggle */}
            <div className="flex items-center gap-1 bg-slate-900 p-1 rounded-xl border border-slate-800">
              <button
                type="button"
                onClick={() => setLang('vi')}
                className={`px-2 py-1 rounded-lg text-xs font-bold transition-all ${
                  isVi ? 'bg-teal-500 text-slate-950 font-black' : 'text-slate-400 hover:text-white'
                }`}
              >
                🇻🇳 VN
              </button>
              <button
                type="button"
                onClick={() => setLang('en')}
                className={`px-2 py-1 rounded-lg text-xs font-bold transition-all ${
                  !isVi ? 'bg-teal-500 text-slate-950 font-black' : 'text-slate-400 hover:text-white'
                }`}
              >
                🇺🇸 EN
              </button>
            </div>

            <button
              type="button"
              onClick={() => setShowPromptModal(true)}
              className="px-3 py-2 rounded-xl text-xs font-bold bg-slate-900 hover:bg-slate-800 text-teal-300 border border-teal-500/30 flex items-center gap-1.5 transition-all"
              title="Xem và sao chép System Instruction Prompt chuẩn cho Google AI Studio"
            >
              <Sparkles size={14} className="text-amber-400" />
              <span>Prompt AI Studio</span>
            </button>

            <button
              type="button"
              onClick={handleRunAiEvaluation}
              disabled={isAiRunning}
              className="px-3.5 py-2 rounded-xl text-xs font-black bg-gradient-to-r from-teal-500 to-emerald-600 hover:from-teal-400 hover:to-emerald-500 text-slate-950 shadow-md shadow-teal-500/20 flex items-center gap-1.5 transition-all disabled:opacity-50"
            >
              <Sparkles size={14} />
              <span>{isAiRunning ? 'Đang thẩm định...' : 'Thẩm Định AI'}</span>
            </button>

            <button
              type="button"
              onClick={handleSaveToReports}
              className="px-3 py-2 rounded-xl text-xs font-bold bg-slate-800 hover:bg-slate-700 text-white border border-slate-700 flex items-center gap-1.5 transition-all"
            >
              <Save size={14} />
              <span>Lưu Hồ Sơ</span>
            </button>

            <button
              type="button"
              onClick={handleExportA4Pdf}
              className="px-3.5 py-2 rounded-xl text-xs font-black bg-gradient-to-r from-indigo-500 to-indigo-600 hover:from-indigo-400 hover:to-indigo-500 text-white shadow-md shadow-indigo-600/30 flex items-center gap-1.5 transition-all"
            >
              <Download size={14} />
              <span>Xuất Biên Bản PDF</span>
            </button>
          </div>
        </div>

        {/* 3 PRESET TEST SCENARIOS */}
        <div className="mt-4 pt-3 border-t border-slate-800/80 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2.5">
          <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1.5 shrink-0">
            <Layers size={13} className="text-teal-400" />
            <span>Kịch bản mẫu kiểm thử nhanh:</span>
          </span>
          <div className="flex items-center gap-2 overflow-x-auto w-full sm:w-auto scrollbar-none py-0.5">
            {PRESET_SCENARIOS.map((sc) => (
              <button
                key={sc.id}
                type="button"
                onClick={() => handleLoadPreset(sc)}
                className="px-3 py-1.5 rounded-xl text-xs font-bold bg-slate-900/90 hover:bg-slate-800 text-slate-200 border border-slate-800 hover:border-slate-700 flex items-center gap-2 shrink-0 transition-all text-left"
                title={sc.descriptionVn}
              >
                <span className={`w-2 h-2 rounded-full ${
                  sc.id.includes('fail') ? 'bg-rose-500 animate-pulse' : sc.id.includes('caution') ? 'bg-amber-400' : 'bg-emerald-400'
                }`} />
                <span>{isVi ? sc.name : sc.nameEn}</span>
                <span className={`text-[9px] px-1.5 py-0.2 rounded font-black border ${sc.badgeColor}`}>
                  {sc.badge}
                </span>
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* Success Banner */}
      {saveSuccessMessage && (
        <div className="p-3 bg-emerald-500/10 border border-emerald-500/30 text-emerald-700 text-xs font-bold rounded-xl flex items-center gap-2 animate-in fade-in duration-200">
          <CheckCircle2 size={16} className="text-emerald-600" />
          <span>{saveSuccessMessage}</span>
        </div>
      )}

      {/* -------------------------------------------------------------------- */}
      {/* 2. OVERALL SEVERITY BANNER (CRITICAL / CAUTION / PASS) */}
      {/* -------------------------------------------------------------------- */}
      <div className={`p-4 sm:p-5 rounded-2xl border shadow-sm flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 ${
        evaluationResult.overall_status === 'CRITICAL'
          ? 'bg-rose-50 border-rose-200 text-rose-950'
          : evaluationResult.overall_status === 'CAUTION'
          ? 'bg-amber-50 border-amber-200 text-amber-950'
          : 'bg-emerald-50 border-emerald-200 text-emerald-950'
      }`}>
        <div className="flex items-start gap-3">
          <div className={`w-10 h-10 rounded-xl flex items-center justify-center shrink-0 ${
            evaluationResult.overall_status === 'CRITICAL'
              ? 'bg-rose-600 text-white shadow-md shadow-rose-600/30'
              : evaluationResult.overall_status === 'CAUTION'
              ? 'bg-amber-500 text-slate-950 shadow-md shadow-amber-500/30'
              : 'bg-emerald-600 text-white shadow-md shadow-emerald-600/30'
          }`}>
            {evaluationResult.overall_status === 'CRITICAL' ? (
              <AlertOctagon size={22} />
            ) : evaluationResult.overall_status === 'CAUTION' ? (
              <AlertTriangle size={22} />
            ) : (
              <ShieldCheck size={22} />
            )}
          </div>
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <span className={`text-[10px] font-black uppercase tracking-wider px-2 py-0.5 rounded-full border ${
                evaluationResult.overall_status === 'CRITICAL'
                  ? 'bg-rose-100 text-rose-900 border-rose-300'
                  : evaluationResult.overall_status === 'CAUTION'
                  ? 'bg-amber-100 text-amber-900 border-amber-300'
                  : 'bg-emerald-100 text-emerald-900 border-emerald-300'
              }`}>
                {evaluationResult.summary_badge}
              </span>
              <span className="text-xs text-slate-500 font-mono">
                NETA ECS-2024 Sec 7.1
              </span>
            </div>
            <h2 className="text-base sm:text-lg font-black tracking-tight mt-0.5">
              {isVi ? evaluationResult.status_title_vn : evaluationResult.status_title_en}
            </h2>
            <p className="text-xs text-slate-600 mt-1 leading-relaxed max-w-4xl">
              {isVi ? evaluationResult.technical_analysis.description_vn : evaluationResult.technical_analysis.description_en}
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 self-end sm:self-center shrink-0">
          <div className="text-right">
            <span className="text-[10px] text-slate-400 block uppercase font-bold">Cách điện Min</span>
            <span className={`text-base font-black font-mono ${
              evaluationResult.technical_analysis.insulation_min_val < 1.0 ? 'text-rose-600' : 'text-emerald-700'
            }`}>
              {evaluationResult.technical_analysis.insulation_min_val} MΩ
            </span>
          </div>
          <div className="h-8 w-px bg-slate-200 mx-1 hidden sm:block" />
          <div className="text-right">
            <span className="text-[10px] text-slate-400 block uppercase font-bold">Lệch Tải Pha</span>
            <span className={`text-base font-black font-mono ${
              evaluationResult.technical_analysis.load_unbalance_pct > 15 ? 'text-amber-600' : 'text-slate-800'
            }`}>
              {evaluationResult.technical_analysis.load_unbalance_pct}%
            </span>
          </div>
        </div>
      </div>

      {/* -------------------------------------------------------------------- */}
      {/* 3. METADATA & 3-STAGE COMMISSIONING CHECKLIST INPUTS */}
      {/* -------------------------------------------------------------------- */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left Column: Identification Metadata & Quick Actions */}
        <div className="lg:col-span-1 space-y-5">
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h3 className="text-sm font-black text-slate-900 flex items-center gap-2">
                <FileCheck2 size={16} className="text-teal-600" />
                <span>Định Danh Tủ &amp; Hệ Thống Hạ Thế</span>
              </h3>
              <span className="text-[10px] px-2 py-0.5 rounded bg-slate-100 font-mono text-slate-600">
                &lt; 1 kV
              </span>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <CustomerSelectDropdown
                  label="Tên Khách Hàng (Master):"
                  value={customerName}
                  onChange={(cName) => setCustomerName(cName)}
                  placeholder="Chọn hoặc thêm Khách Hàng mới..."
                />
              </div>

              <div>
                <ProjectSelectDropdown
                  label="Tên Dự án / Nhà máy:"
                  value={projectName}
                  customerName={customerName}
                  onCustomerAutoSelect={(cName) => {
                    if (!customerName) setCustomerName(cName);
                  }}
                  onChange={(pName) => setProjectName(pName)}
                  placeholder="Chọn hoặc thêm Dự Án / Nhà Máy..."
                  standard="ecs"
                />
              </div>

              <div>
                <label className="block text-slate-600 font-semibold mb-1">Mã Ngăn Lộ / Tủ Điện Hạ Thế:</label>
                <input
                  type="text"
                  value={systemTag}
                  onChange={(e) => setSystemTag(e.target.value)}
                  className="w-full p-2 rounded-lg border border-slate-200 text-xs font-bold text-slate-900 focus:ring-1 focus:ring-teal-500 focus:outline-hidden"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-slate-600 font-semibold mb-1">Điện Áp Định Mức (V):</label>
                  <input
                    type="number"
                    value={voltageRatingV}
                    onChange={(e) => setVoltageRatingV(Number(e.target.value))}
                    className="w-full p-2 rounded-lg border border-slate-200 text-xs font-mono font-bold focus:ring-1 focus:ring-teal-500 focus:outline-hidden"
                  />
                </div>
                <div>
                  <label className="block text-slate-600 font-semibold mb-1">Ngày Chạy Thử:</label>
                  <input
                    type="date"
                    value={commissioningDate}
                    onChange={(e) => setCommissioningDate(e.target.value)}
                    className="w-full p-2 rounded-lg border border-slate-200 text-xs focus:ring-1 focus:ring-teal-500 focus:outline-hidden"
                  />
                </div>
              </div>

              <div>
                <label className="block text-slate-600 font-semibold mb-1">Kỹ Sư Nghiệm Thu FSE (NETA III):</label>
                <FseEngineerDropdown
                  value={fseName}
                  onChange={(eng) => setFseName(eng)}
                />
              </div>

              <div>
                <label className="block text-slate-600 font-semibold mb-1">Số Hiệu Biên Bản (Report No):</label>
                <input
                  type="text"
                  value={reportNo}
                  onChange={(e) => setReportNo(e.target.value)}
                  className="w-full p-2 rounded-lg border border-slate-200 text-xs font-mono text-slate-700 focus:ring-1 focus:ring-teal-500 focus:outline-hidden"
                />
              </div>
            </div>
          </div>

          {/* Standards Summary Card */}
          <div className="bg-gradient-to-br from-slate-900 to-teal-950 text-white p-5 rounded-2xl border border-teal-500/30 shadow-sm space-y-3 text-xs">
            <h4 className="font-bold text-teal-300 flex items-center gap-2">
              <Info size={15} />
              <span>Tiêu Chuẩn Tham Chiếu Nghiệm Thu Hạ Thế</span>
            </h4>
            <ul className="space-y-2 text-slate-300">
              <li className="flex items-start gap-2">
                <span className="text-teal-400 font-bold shrink-0">•</span>
                <span><strong>NETA Table 100.1</strong>: Điện trở cách điện khô hệ &lt; 1 kV tối thiểu <strong className="text-white">≥ 1.0 MΩ</strong> ở 1000V DC.</span>
              </li>
              <li className="flex items-start gap-2">
                <span className="text-teal-400 font-bold shrink-0">•</span>
                <span><strong>NETA Table 100.12</strong>: Kiểm tra lực siết bu-lông thanh cái bằng cờ-lê lực đã hiệu chuẩn.</span>
              </li>
              <li className="flex items-start gap-2">
                <span className="text-teal-400 font-bold shrink-0">•</span>
                <span><strong>NETA Sec 7.13 / IEEE 81</strong>: Điện trở nối đất cọc tiếp địa <strong className="text-white">≤ 5.0 Ω</strong>.</span>
              </li>
              <li className="flex items-start gap-2">
                <span className="text-teal-400 font-bold shrink-0">•</span>
                <span><strong>NETA Table 100.18</strong>: Khảo sát hồng ngoại mang tải, không có điểm quá nhiệt bất thường.</span>
              </li>
              <li className="flex items-start gap-2">
                <span className="text-teal-400 font-bold shrink-0">•</span>
                <span><strong>NETA ECS 7.1.B.1</strong>: Tuyệt đối tháo dỡ 100% tiếp địa tạm trước khi đóng aptomat tổng!</span>
              </li>
            </ul>
          </div>
        </div>

        {/* Right Column: 3 Stages Checklist Form */}
        <div className="lg:col-span-2 space-y-5">
          {/* Stage Tab Filters */}
          <div className="bg-white p-1.5 rounded-xl border border-slate-200 flex items-center gap-1 overflow-x-auto text-xs shadow-xs">
            <button
              type="button"
              onClick={() => setActiveStageTab('all')}
              className={`px-3 py-1.5 rounded-lg font-bold transition-all ${
                activeStageTab === 'all'
                  ? 'bg-slate-900 text-white shadow-2xs'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-slate-50'
              }`}
            >
              Tất Cả 3 Giai Đoạn (13 Mục)
            </button>
            <button
              type="button"
              onClick={() => setActiveStageTab('stageA')}
              className={`px-3 py-1.5 rounded-lg font-bold transition-all ${
                activeStageTab === 'stageA'
                  ? 'bg-teal-600 text-white shadow-2xs'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-slate-50'
              }`}
            >
              GĐ 1: Tiền Đóng Điện (7.1.A)
            </button>
            <button
              type="button"
              onClick={() => setActiveStageTab('stageB')}
              className={`px-3 py-1.5 rounded-lg font-bold transition-all ${
                activeStageTab === 'stageB'
                  ? 'bg-teal-600 text-white shadow-2xs'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-slate-50'
              }`}
            >
              GĐ 2: Trong Đóng Điện (7.1.B)
            </button>
            <button
              type="button"
              onClick={() => setActiveStageTab('stageC')}
              className={`px-3 py-1.5 rounded-lg font-bold transition-all ${
                activeStageTab === 'stageC'
                  ? 'bg-teal-600 text-white shadow-2xs'
                  : 'text-slate-600 hover:text-slate-900 hover:bg-slate-50'
              }`}
            >
              GĐ 3: Sau Đóng Điện (7.1.C)
            </button>
          </div>

          {/* GIAI ĐOẠN 1: TIỀN ĐÓNG ĐIỆN (7.1.A) */}
          {(activeStageTab === 'all' || activeStageTab === 'stageA') && (
            <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <h3 className="text-sm font-black text-slate-900 flex items-center gap-2">
                  <span className="w-6 h-6 rounded-lg bg-teal-100 text-teal-800 flex items-center justify-center font-bold text-xs">
                    A
                  </span>
                  <span>GIAI ĐOẠN 1: TIỀN CHẠY THỬ / TRƯỚC KHI ĐÓNG ĐIỆN (NETA ECS 7.1.A)</span>
                </h3>
                <span className="text-[11px] font-bold text-teal-700 bg-teal-50 px-2.5 py-0.5 rounded-md border border-teal-200">
                  Pre-Energization
                </span>
              </div>

              <div className="space-y-4 text-xs">
                {/* A1: Documentation */}
                <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                  <div>
                    <span className="font-bold text-slate-900 block">A1. Rà soát Hồ sơ &amp; Nhãn mác (OPR / BOD / Drawings)</span>
                    <span className="text-[11px] text-slate-500">NETA ECS Section 7.1.A.1 - 7.1.A.6: Khớp 100% hồ sơ thiết kế và thông số định mức.</span>
                  </div>
                  <label className="flex items-center gap-2 cursor-pointer shrink-0">
                    <input
                      type="checkbox"
                      checked={acceptanceTestReportsReviewed}
                      onChange={(e) => setAcceptanceTestReportsReviewed(e.target.checked)}
                      className="w-4 h-4 accent-teal-600 rounded"
                    />
                    <span className="font-bold text-slate-800">Đã duyệt Báo cáo Acceptance</span>
                  </label>
                </div>

                {/* A2: Busbar Insulation Resistance (Table 100.1) */}
                <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
                  <div className="flex items-center justify-between flex-wrap gap-2">
                    <div>
                      <span className="font-bold text-slate-900 block">
                        A2. Thử nghiệm Acceptance Testing: Điện trở Cách điện Thanh cái (1000V DC)
                      </span>
                      <span className="text-[11px] text-slate-500">
                        NETA ATS Section 7.1.A.3 &amp; <strong>Table 100.1</strong>: Ngưỡng tối thiểu <strong className="text-indigo-900">≥ 1.0 MΩ</strong> cho hệ hạ thế &lt; 1 kV.
                      </span>
                    </div>
                    {Math.min(insulationPhaseA, insulationPhaseB, insulationPhaseC) < 1.0 && (
                      <span className="px-2 py-0.5 rounded text-[10px] font-black bg-rose-100 text-rose-800 border border-rose-300 animate-pulse">
                        CẢNH BÁO: CÓ PHA &lt; 1.0 MΩ (NGUY CƠ NGẮN MẠCH)
                      </span>
                    )}
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-1">
                    <div className="p-2 rounded-lg bg-white border border-slate-200">
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">Pha A - Đất (MΩ):</label>
                      <input
                        type="number"
                        step="0.01"
                        value={insulationPhaseA}
                        onChange={(e) => setInsulationPhaseA(Number(e.target.value))}
                        className={`w-full p-1.5 rounded border text-xs font-mono font-bold ${
                          insulationPhaseA < 1.0 ? 'border-rose-400 bg-rose-50 text-rose-800' : 'border-slate-200'
                        }`}
                      />
                    </div>
                    <div className="p-2 rounded-lg bg-white border border-slate-200">
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">Pha B - Đất (MΩ):</label>
                      <input
                        type="number"
                        step="0.01"
                        value={insulationPhaseB}
                        onChange={(e) => setInsulationPhaseB(Number(e.target.value))}
                        className={`w-full p-1.5 rounded border text-xs font-mono font-bold ${
                          insulationPhaseB < 1.0 ? 'border-rose-400 bg-rose-50 text-rose-800' : 'border-slate-200'
                        }`}
                      />
                    </div>
                    <div className="p-2 rounded-lg bg-white border border-slate-200">
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">Pha C - Đất (MΩ):</label>
                      <input
                        type="number"
                        step="0.01"
                        value={insulationPhaseC}
                        onChange={(e) => setInsulationPhaseC(Number(e.target.value))}
                        className={`w-full p-1.5 rounded border text-xs font-mono font-bold ${
                          insulationPhaseC < 1.0 ? 'border-rose-400 bg-rose-50 text-rose-800' : 'border-slate-200'
                        }`}
                      />
                    </div>
                  </div>
                </div>

                {/* A3 & A4: Torque & Grounding */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-1.5">
                    <span className="font-bold text-slate-900 block">A3. Độ Chặt Bu-lông (Bolt Torque)</span>
                    <span className="text-[11px] text-slate-500 block">ANSI/NETA Table 100.12 (Torque Wrench).</span>
                    <select
                      value={boltedTorqueStatus}
                      onChange={(e) => setBoltedTorqueStatus(e.target.value as any)}
                      className="w-full p-2 rounded-lg border border-slate-200 text-xs font-bold"
                    >
                      <option value="Pass">Pass - Đạt lực siết Table 100.12</option>
                      <option value="Fail">Fail - Chưa siết lực đủ tiêu chuẩn</option>
                    </select>
                  </div>

                  <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-1.5">
                    <span className="font-bold text-slate-900 block">A4. Tiếp Địa &amp; Đất (IEEE 81 / Sec 7.13)</span>
                    <span className="text-[11px] text-slate-500 block">Điện trở nối đất cọc tiếp địa: Chuẩn ≤ 5.0 Ω.</span>
                    <div className="flex items-center gap-2">
                      <input
                        type="number"
                        step="0.01"
                        value={groundingResistanceOhms}
                        onChange={(e) => setGroundingResistanceOhms(Number(e.target.value))}
                        className={`w-full p-2 rounded-lg border text-xs font-mono font-bold ${
                          groundingResistanceOhms > 5.0 ? 'border-rose-400 bg-rose-50 text-rose-800' : 'border-slate-200'
                        }`}
                      />
                      <span className="font-bold text-slate-700">Ω</span>
                    </div>
                  </div>
                </div>

                {/* A5 & A6: Interlocks & Energization Plan */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-1.5">
                    <span className="font-bold text-slate-900 block">A5. Khóa An toàn &amp; Liên động (Interlocks / E-Stop)</span>
                    <span className="text-[11px] text-slate-500 block">NETA ECS Section 7.1.A.30 &amp; 7.1.A.32.</span>
                    <select
                      value={safetyInterlocksEstop}
                      onChange={(e) => setSafetyInterlocksEstop(e.target.value as any)}
                      className="w-full p-2 rounded-lg border border-slate-200 text-xs font-bold"
                    >
                      <option value="Pass">Pass - Liên động tốt, E-Stop ngắt máy cắt</option>
                      <option value="Fail">Fail - Lỗi cơ cấu khóa hoặc E-Stop</option>
                    </select>
                  </div>

                  <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-1.5">
                    <span className="font-bold text-slate-900 block">A6. Kế Hoạch Đóng Điện &amp; Nguồn DC Phụ Trợ</span>
                    <span className="text-[11px] text-slate-500 block">NETA ECS 7.1.A.41 &amp; 7.1.A.45 (Written Plan).</span>
                    <label className="flex items-center gap-2 cursor-pointer pt-1">
                      <input
                        type="checkbox"
                        checked={writtenEnergizationPlanApproved}
                        onChange={(e) => setWrittenEnergizationPlanApproved(e.target.checked)}
                        className="w-4 h-4 accent-teal-600 rounded"
                      />
                      <span className="font-bold text-slate-800">Kế hoạch đóng điện đã được ký duyệt</span>
                    </label>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* GIAI ĐOẠN 2: TRONG KHI ĐÓNG ĐIỆN (7.1.B) */}
          {(activeStageTab === 'all' || activeStageTab === 'stageB') && (
            <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <h3 className="text-sm font-black text-slate-900 flex items-center gap-2">
                  <span className="w-6 h-6 rounded-lg bg-blue-100 text-blue-800 flex items-center justify-center font-bold text-xs">
                    B
                  </span>
                  <span>GIAI ĐOẠN 2: TRONG KHI ĐÓNG ĐIỆN (NETA ECS 7.1.B)</span>
                </h3>
                <span className="text-[11px] font-bold text-blue-700 bg-blue-50 px-2.5 py-0.5 rounded-md border border-blue-200">
                  Energization Phase
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                {/* B1: Temporary Grounds Removal */}
                <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-slate-900 block">B1. Tháo dỡ Tiếp Địa Tạm Thời</span>
                    <span className="text-[10px] text-slate-500">7.1.B.1 &amp; 7.1.B.2</span>
                  </div>
                  <label className="flex items-center gap-2 cursor-pointer p-2 rounded-lg bg-white border border-slate-200">
                    <input
                      type="checkbox"
                      checked={tempProtectiveGroundsRemoved}
                      onChange={(e) => setTempProtectiveGroundsRemoved(e.target.checked)}
                      className="w-4 h-4 accent-teal-600 rounded"
                    />
                    <span className={`font-bold ${!tempProtectiveGroundsRemoved ? 'text-rose-600' : 'text-slate-800'}`}>
                      {tempProtectiveGroundsRemoved
                        ? '✔ Đã tháo bỏ 100% tiếp địa di động'
                        : '❌ CẢNH BÁO: CÒN TIẾP ĐỊA TẠM!'}
                    </span>
                  </label>
                </div>

                {/* B2: Phase Sequence */}
                <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-slate-900 block">B2. Thứ Tự Pha &amp; Góc Pha (Phasing)</span>
                    <span className="text-[10px] text-slate-500">7.1.B.8 &amp; 7.1.B.9</span>
                  </div>
                  <select
                    value={phaseSequenceCheck}
                    onChange={(e) => setPhaseSequenceCheck(e.target.value as any)}
                    className="w-full p-2 rounded-lg border border-slate-200 text-xs font-bold"
                  >
                    <option value="Pass">Pass - Thứ tự pha chuẩn A-B-C (U-V-W)</option>
                    <option value="Fail">Fail - Ngược thứ tự pha (A-C-B)</option>
                  </select>
                </div>

                {/* B3: Motor Rotation */}
                <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-slate-900 block">B3. Chiều Quay Động Cơ &amp; Tải</span>
                    <span className="text-[10px] text-slate-500">7.1.B.10 &amp; 7.1.B.11</span>
                  </div>
                  <select
                    value={motorRotationCheck}
                    onChange={(e) => setMotorRotationCheck(e.target.value as any)}
                    className="w-full p-2 rounded-lg border border-slate-200 text-xs font-bold"
                  >
                    <option value="Pass">Pass - Động cơ quay đúng 100% chiều mũi tên</option>
                    <option value="Fail">Fail - Phát hiện động cơ quay ngược</option>
                  </select>
                </div>

                {/* B4: Alarms & Metering */}
                <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="font-bold text-slate-900 block">B4. Cảnh Báo &amp; Đo Lường Metering</span>
                    <span className="text-[10px] text-slate-500">7.1.B.7 &amp; 7.1.B.13</span>
                  </div>
                  <select
                    value={alarmsAndMeteringStatus}
                    onChange={(e) => setAlarmsAndMeteringStatus(e.target.value as any)}
                    className="w-full p-2 rounded-lg border border-slate-200 text-xs font-bold"
                  >
                    <option value="Pass">Pass - 0 Alarm, đồng hồ hiển thị chuẩn</option>
                    <option value="Fail">Fail - Còn cảnh báo sự cố Fault/Alarm</option>
                  </select>
                </div>
              </div>
            </div>
          )}

          {/* GIAI ĐOẠN 3: SAU KHI ĐÓNG ĐIỆN (7.1.C) */}
          {(activeStageTab === 'all' || activeStageTab === 'stageC') && (
            <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <h3 className="text-sm font-black text-slate-900 flex items-center gap-2">
                  <span className="w-6 h-6 rounded-lg bg-amber-100 text-amber-800 flex items-center justify-center font-bold text-xs">
                    C
                  </span>
                  <span>GIAI ĐOẠN 3: SAU KHI ĐÓNG ĐIỆN MANG TẢI (NETA ECS 7.1.C)</span>
                </h3>
                <span className="text-[11px] font-bold text-amber-700 bg-amber-50 px-2.5 py-0.5 rounded-md border border-amber-200">
                  Post-Energization
                </span>
              </div>

              <div className="space-y-4 text-xs">
                {/* C1: Thermographic Survey Table 100.18 */}
                <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <span className="font-bold text-slate-900 block">
                      C1. Khảo Sát Nhiệt Hồng Ngoại Mang Tải (Table 100.18 / NFPA 70B)
                    </span>
                    <span className="text-[11px] text-slate-500">
                      NETA ECS Section 7.1.C.3: Chụp ảnh nhiệt toàn bộ aptomat, thanh cái và đầu cáp dưới tải vận hành.
                    </span>
                  </div>
                  <select
                    value={thermographicSurveyStatus}
                    onChange={(e) => setThermographicSurveyStatus(e.target.value as any)}
                    className="p-2 rounded-lg border border-slate-200 text-xs font-bold shrink-0"
                  >
                    <option value="Pass">Pass - Nhiệt độ đồng đều, 0 Hotspot</option>
                    <option value="Investigate">Investigate - Xuất hiện Hotspot bất thường</option>
                  </select>
                </div>

                {/* C2: Load Balance & Power Quality */}
                <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
                  <div className="flex items-center justify-between flex-wrap gap-2">
                    <div>
                      <span className="font-bold text-slate-900 block">
                        C2. Cân Bằng Tải &amp; Chất Lượng Điện Năng (7.1.C.6 &amp; 7.1.C.7)
                      </span>
                      <span className="text-[11px] text-slate-500">
                        Đo dòng điện tải 3 pha. Tiêu chuẩn khuyến nghị độ lệch mất cân bằng ≤ 15%.
                      </span>
                    </div>
                    <span className={`px-2 py-0.5 rounded text-[10px] font-black ${
                      phaseUnbalance > 15 ? 'bg-amber-100 text-amber-800 border border-amber-300' : 'bg-emerald-100 text-emerald-800'
                    }`}>
                      Độ Lệch Pha: {phaseUnbalance}% {phaseUnbalance > 15 ? '(Vượt 15%)' : '(Đạt)'}
                    </span>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-1">
                    <div className="p-2 rounded-lg bg-white border border-slate-200">
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">Dòng Tải Pha A (Amps):</label>
                      <input
                        type="number"
                        value={loadCurrentPhaseA}
                        onChange={(e) => setLoadCurrentPhaseA(Number(e.target.value))}
                        className="w-full p-1.5 rounded border border-slate-200 text-xs font-mono font-bold"
                      />
                    </div>
                    <div className="p-2 rounded-lg bg-white border border-slate-200">
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">Dòng Tải Pha B (Amps):</label>
                      <input
                        type="number"
                        value={loadCurrentPhaseB}
                        onChange={(e) => setLoadCurrentPhaseB(Number(e.target.value))}
                        className="w-full p-1.5 rounded border border-slate-200 text-xs font-mono font-bold"
                      />
                    </div>
                    <div className="p-2 rounded-lg bg-white border border-slate-200">
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">Dòng Tải Pha C (Amps):</label>
                      <input
                        type="number"
                        value={loadCurrentPhaseC}
                        onChange={(e) => setLoadCurrentPhaseC(Number(e.target.value))}
                        className="w-full p-1.5 rounded border border-slate-200 text-xs font-mono font-bold"
                      />
                    </div>
                  </div>
                </div>

                {/* C3: GFCI Receptacle Test */}
                <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <span className="font-bold text-slate-900 block">
                      C3. Thử Nghiệm Ổ Cắm Chống Giật GFCI (NETA ECS 7.1.C.10)
                    </span>
                    <span className="text-[11px] text-slate-500">
                      Tạo dòng rò mô phỏng thử nghiệm tác động ngắt của 100% ổ cắm GFCI hạ thế trong khu vực.
                    </span>
                  </div>
                  <select
                    value={gfciReceptaclesTested}
                    onChange={(e) => setGfciReceptaclesTested(e.target.value as any)}
                    className="p-2 rounded-lg border border-slate-200 text-xs font-bold shrink-0"
                  >
                    <option value="Pass">Pass - 100% ổ cắm GFCI nhảy chính xác</option>
                    <option value="Fail">Fail - Có ổ cắm không tác động</option>
                  </select>
                </div>
              </div>
            </div>
          )}
        </div>
      </div>

      {/* -------------------------------------------------------------------- */}
      {/* 4. DETAILED EVALUATION TABLE (NETA ECS SECTION 7.1 & TABLES) */}
      {/* -------------------------------------------------------------------- */}
      <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
        <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 border-b border-slate-100 pb-3">
          <div>
            <h3 className="text-sm font-black text-slate-900 flex items-center gap-2">
              <Layers size={17} className="text-teal-600" />
              <span>BẢNG PHÂN TÍCH CHI TIẾT NGHIỆM THU HẠ THẾ (NETA ECS SECTION 7.1)</span>
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Đối chiếu 13 hạng mục chạy thử theo tiêu chuẩn ANSI/NETA ECS-2024 &amp; các bảng NETA Tables 100.1, 100.12, 100.18
            </p>
          </div>
          <span className="text-xs font-bold text-slate-600 bg-slate-100 px-2.5 py-1 rounded-lg">
            Tổng cộng: {evaluationResult.evaluation_table.length} Hạng mục
          </span>
        </div>

        <div className="overflow-x-auto border border-slate-200 rounded-xl">
          <table className="w-full text-left text-xs border-collapse">
            <thead>
              <tr className="bg-slate-100 text-slate-700 text-[11px]">
                <th className="p-2.5 border-b border-slate-200 font-bold text-center">STT</th>
                <th className="p-2.5 border-b border-slate-200 font-bold">Hạng Mục Chạy Thử / Commissioning Item</th>
                <th className="p-2.5 border-b border-slate-200 font-bold">Tiêu Chuẩn Tham Chiếu</th>
                <th className="p-2.5 border-b border-slate-200 font-bold">Giá Trị Đo Thực Tế</th>
                <th className="p-2.5 border-b border-slate-200 font-bold text-center">Đánh Giá</th>
              </tr>
            </thead>
            <tbody>
              {evaluationResult.evaluation_table.map((row) => (
                <tr key={row.item_code} className="hover:bg-slate-50/90 transition-colors border-b border-slate-100 last:border-b-0">
                  <td className="p-2.5 text-center font-bold font-mono text-slate-500">{row.item_code}</td>
                  <td className="p-2.5 font-semibold text-slate-900">
                    <div>{isVi ? row.item_name_vn : row.item_name_en}</div>
                    <div className="text-[10px] text-slate-500 mt-0.5">{isVi ? row.evaluation_vn : row.evaluation_en}</div>
                  </td>
                  <td className="p-2.5 text-slate-600 font-mono text-[11px]">{row.neta_standard_ref}</td>
                  <td className="p-2.5 font-bold font-mono text-slate-800">{row.measured_value}</td>
                  <td className="p-2.5 text-center">
                    <span className={`px-2 py-0.5 rounded text-[10px] font-black border ${
                      row.status === 'FAIL'
                        ? 'bg-rose-100 text-rose-800 border-rose-300'
                        : row.status === 'INVESTIGATE'
                        ? 'bg-amber-100 text-amber-800 border-amber-300'
                        : 'bg-emerald-100 text-emerald-800 border-emerald-300'
                    }`}>
                      {row.status === 'PASS' ? 'PASS / ĐẠT' : row.status === 'FAIL' ? 'FAIL / KHÔNG ĐẠT' : 'INVESTIGATE'}
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* -------------------------------------------------------------------- */}
      {/* 5. TECHNICAL HAZARD ANALYSIS & ACTIONABLE RECOMMENDATIONS */}
      {/* -------------------------------------------------------------------- */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Technical Hazard Box */}
        <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-3">
          <div className="flex items-center gap-2 border-b border-slate-100 pb-3">
            <AlertOctagon size={18} className={evaluationResult.technical_analysis.hazard_identified ? 'text-rose-600' : 'text-emerald-600'} />
            <h3 className="text-sm font-black text-slate-900">
              {isVi ? 'Phân Tích Kỹ Thuật & Nguy Cơ Tiềm Ẩn' : 'Technical Analysis & Hazard Assessment'}
            </h3>
          </div>
          <div>
            <h4 className="font-bold text-xs text-slate-900 mb-1">
              {isVi ? evaluationResult.technical_analysis.title_vn : evaluationResult.technical_analysis.title_en}
            </h4>
            <p className="text-xs text-slate-600 leading-relaxed">
              {isVi ? evaluationResult.technical_analysis.description_vn : evaluationResult.technical_analysis.description_en}
            </p>
          </div>
          <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 space-y-1.5 text-xs">
            <span className="font-bold text-slate-700 block">Tóm tắt Thông số Đo Nòng Cốt:</span>
            <div className="grid grid-cols-3 gap-2 text-center pt-1 font-mono">
              <div className="p-1.5 bg-white rounded border border-slate-200">
                <span className="text-[10px] text-slate-400 block font-sans">Cách điện Min</span>
                <span className={`font-black ${evaluationResult.technical_analysis.insulation_min_val < 1.0 ? 'text-rose-600' : 'text-emerald-700'}`}>
                  {evaluationResult.technical_analysis.insulation_min_val} MΩ
                </span>
              </div>
              <div className="p-1.5 bg-white rounded border border-slate-200">
                <span className="text-[10px] text-slate-400 block font-sans">Tiếp địa</span>
                <span className={`font-black ${evaluationResult.technical_analysis.grounding_val > 5.0 ? 'text-rose-600' : 'text-emerald-700'}`}>
                  {evaluationResult.technical_analysis.grounding_val} Ω
                </span>
              </div>
              <div className="p-1.5 bg-white rounded border border-slate-200">
                <span className="text-[10px] text-slate-400 block font-sans">Lệch Tải</span>
                <span className={`font-black ${evaluationResult.technical_analysis.load_unbalance_pct > 15 ? 'text-amber-600' : 'text-emerald-700'}`}>
                  {evaluationResult.technical_analysis.load_unbalance_pct}%
                </span>
              </div>
            </div>
          </div>
        </div>

        {/* Actionable Recommendations Box */}
        <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-3">
          <div className="flex items-center gap-2 border-b border-slate-100 pb-3">
            <ShieldCheck size={18} className="text-teal-600" />
            <h3 className="text-sm font-black text-slate-900">
              {isVi ? 'Khuyến Nghị Hành Động Chi Tiết Cho FSE' : 'Actionable Recommendations for FSE'}
            </h3>
          </div>
          <div className="space-y-3">
            {evaluationResult.actionable_recommendations.steps.map((st) => (
              <div key={st.step_number} className="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-1 text-xs">
                <div className="font-bold text-slate-900 flex items-center gap-2">
                  <span className="w-5 h-5 rounded-full bg-slate-900 text-white flex items-center justify-center text-[10px]">
                    {st.step_number}
                  </span>
                  <span>{isVi ? st.title_vn : st.title_en}</span>
                </div>
                <p className="text-[11px] text-slate-600 pl-7 leading-relaxed">
                  {isVi ? st.details_vn : st.details_en}
                </p>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* -------------------------------------------------------------------- */}
      {/* 6. MODAL: SYSTEM INSTRUCTION PROMPT FOR GOOGLE AI STUDIO */}
      {/* -------------------------------------------------------------------- */}
      {showPromptModal && (
        <div className="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-slate-900 text-white border border-teal-500/40 rounded-2xl max-w-3xl w-full max-h-[85vh] flex flex-col shadow-2xl">
            <div className="p-4 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Sparkles size={18} className="text-amber-400" />
                <h3 className="font-bold text-sm text-white">
                  Master System Instruction Prompt (ANSI/NETA ECS-2024 Section 7.1)
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setShowPromptModal(false)}
                className="text-slate-400 hover:text-white p-1"
              >
                ✕
              </button>
            </div>

            <div className="p-4 overflow-y-auto flex-1 font-mono text-xs bg-slate-950 text-teal-200 whitespace-pre-wrap leading-relaxed select-all">
              {ECS_LV_SYSTEM_INSTRUCTION}
            </div>

            <div className="p-4 border-t border-slate-800 flex items-center justify-between">
              <span className="text-xs text-slate-400">
                Nhấp nút Copy rồi dán trực tiếp vào mục "System Instructions" trên Google AI Studio.
              </span>
              <button
                type="button"
                onClick={() => {
                  navigator.clipboard.writeText(ECS_LV_SYSTEM_INSTRUCTION);
                  setCopiedPrompt(true);
                  setTimeout(() => setCopiedPrompt(false), 3000);
                }}
                className="flex items-center gap-2 px-4 py-2 bg-teal-600 hover:bg-teal-500 text-slate-950 font-black rounded-xl text-xs shadow-sm"
              >
                {copiedPrompt ? <Check size={14} className="text-white" /> : <Copy size={14} />}
                <span>{copiedPrompt ? 'Đã sao chép vào Clipboard!' : 'Sao chép Prompt'}</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
