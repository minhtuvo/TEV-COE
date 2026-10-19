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
  Radio,
  Zap,
  Gauge,
  Thermometer,
  Flame,
  Globe,
  ClipboardCheck,
  Search,
  Activity,
  Droplets,
  Wind
} from 'lucide-react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  NetaEcsMvInputPayload,
  NetaEcsMvEvaluationResult,
  performDeterministicEcsMvAnalysis,
  ECS_MV_SYSTEM_INSTRUCTION
} from '../../server/netaEcsMediumVoltageAnalyzer';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
  initialLanguage?: 'vi' | 'en';
}

interface PresetScenario {
  id: string;
  name: string;
  nameEn: string;
  badge: string;
  badgeColor: string;
  systemTag: string;
  nominalVoltageKv: number;
  insA: number;
  insB: number;
  insC: number;
  isVlfPass: boolean;
  vlfKv: number;
  tanDelta: number;
  isTorquePass: boolean;
  sf6Pressure: number;
  isDiffProtPass: boolean;
  isInterlocksPass: boolean;
  isGroundsRemoved: boolean;
  isPhasingPass: boolean;
  phasingDeltaV: number;
  isThermalPass: boolean;
  tevMaxDb: number;
  acousticMaxDb: number;
  descriptionVn: string;
}

const PRESET_SCENARIOS: PresetScenario[] = [
  {
    id: 'preset-mv-all-pass',
    name: 'Nghiệm thu Đạt Chuẩn Toàn diện (22kV Bus A - Mẫu Đề Bài)',
    nameEn: 'Fully Commissioned & Approved (22kV Bus A - Benchmark Scenario)',
    badge: 'PASS / ĐỦ ĐIỀU KIỆN ĐÓNG ĐIỆN',
    badgeColor: 'bg-emerald-500/20 text-emerald-300 border-emerald-400/40',
    systemTag: 'MV-SWITCHGEAR-22KV-BUS-A (22 kV System)',
    nominalVoltageKv: 22,
    insA: 4500.0,
    insB: 4800.0,
    insC: 4600.0,
    isVlfPass: true,
    vlfKv: 28.0,
    tanDelta: 0.0012,
    isTorquePass: true,
    sf6Pressure: 5.8,
    isDiffProtPass: true,
    isInterlocksPass: true,
    isGroundsRemoved: true,
    isPhasingPass: true,
    phasingDeltaV: 12,
    isThermalPass: true,
    tevMaxDb: 14,
    acousticMaxDb: 1.2,
    descriptionVn: 'Cách điện > 4500 MΩ (Table 100.1), thử VLF 28kV Tan Delta 0.0012, SF6 5.8 bar, đồng pha máy cắt liên lạc ΔV=12V (<= 15V), PD TEV 14 dB an toàn.'
  },
  {
    id: 'preset-mv-critical-hazard',
    name: 'Sự cố Cách điện Thanh cái 22kV (< 1,000 MΩ) & Lệch Pha Hòa Nguồn',
    nameEn: 'Critical Dielectric Breakdown & Out-of-Phase Condition',
    badge: 'CRITICAL / NGUY HẠI ĐÓNG ĐIỆN',
    badgeColor: 'bg-rose-500/20 text-rose-300 border-rose-400/40',
    systemTag: 'MV-SWITCHGEAR-22KV-BUS-B (22 kV Outgoing)',
    nominalVoltageKv: 22,
    insA: 320.0,
    insB: 4100.0,
    insC: 4200.0,
    isVlfPass: false,
    vlfKv: 28.0,
    tanDelta: 0.0085,
    isTorquePass: true,
    sf6Pressure: 4.2,
    isDiffProtPass: false,
    isInterlocksPass: true,
    isGroundsRemoved: true,
    isPhasingPass: false,
    phasingDeltaV: 180,
    isThermalPass: false,
    tevMaxDb: 34,
    acousticMaxDb: 18.0,
    descriptionVn: 'Cách điện Pha A chỉ 320 MΩ (< 1,000 MΩ Table 100.1), cáp VLF Tan Delta cao, chênh áp đồng pha 180V (LỆCH PHA NGUY CƠ NỔ), PD TEV 34 dB. CẤM ĐÓNG ĐIỆN!'
  },
  {
    id: 'preset-mv-caution-punchlist',
    name: 'Cảnh báo Áp suất Khí SF6 Tụt Giảm & Tín hiệu PD Chớm Xuất Hiện',
    nameEn: 'Caution: Low SF6 Gas Pressure & Elevated Online PD Warning',
    badge: 'CAUTION / CẦN KHẮC PHỤC',
    badgeColor: 'bg-amber-500/20 text-amber-300 border-amber-400/40',
    systemTag: 'MV-GIS-22KV-FEEDER-03 (Gas Insulated Switchgear)',
    nominalVoltageKv: 22,
    insA: 2800.0,
    insB: 2900.0,
    insC: 2750.0,
    isVlfPass: true,
    vlfKv: 28.0,
    tanDelta: 0.0028,
    isTorquePass: false,
    sf6Pressure: 5.1,
    isDiffProtPass: true,
    isInterlocksPass: true,
    isGroundsRemoved: true,
    isPhasingPass: true,
    phasingDeltaV: 14,
    isThermalPass: false,
    tevMaxDb: 24,
    acousticMaxDb: 5.5,
    descriptionVn: 'Áp suất SF6 ở mức 5.1 bar (thấp hơn định mức 5.8 bar), phát hiện TEV 24 dBmV vượt ngưỡng Warning (20 dB). Cần lập Punch-List xử lý trong 48h.'
  }
];

export const NetaEcsMediumVoltageChecklist: React.FC<Props> = ({
  onSaveReport,
  onNavigateToReports,
  initialLanguage = 'vi'
}) => {
  const [lang, setLang] = useState<'vi' | 'en'>(initialLanguage);
  const isVi = lang === 'vi';

  // Metadata
  const [reportNo, setReportNo] = useState(`TEV-ECS-2024-MV-${Math.floor(100 + Math.random() * 900)}`);
  const [projectName, setProjectName] = useState('Tram Bien Ap Trung The TEV Substation');
  const [systemTag, setSystemTag] = useState('MV-SWITCHGEAR-22KV-BUS-A (22 kV System)');
  const [nominalVoltageKv, setNominalVoltageKv] = useState(22);
  const [fseName, setFseName] = useState('Nguyen Van I');
  const [commissioningDate, setCommissioningDate] = useState(new Date().toISOString().split('T')[0]);

  // Stage 1: Pre-Energization (7.2.A)
  const [acceptanceTestReportsReviewed, setAcceptanceTestReportsReviewed] = useState(true);
  const [insulationPhaseA, setInsulationPhaseA] = useState(4500.0);
  const [insulationPhaseB, setInsulationPhaseB] = useState(4800.0);
  const [insulationPhaseC, setInsulationPhaseC] = useState(4600.0);
  const [vlfTestVoltageKv, setVlfTestVoltageKv] = useState(28.0);
  const [vlfTanDeltaMean, setVlfTanDeltaMean] = useState(0.0012);
  const [vlfStatus, setVlfStatus] = useState<'Pass' | 'Fail'>('Pass');
  const [boltedTorqueStatus, setBoltedTorqueStatus] = useState<'Pass' | 'Fail'>('Pass');
  const [sf6MeasuredPressure, setSf6MeasuredPressure] = useState(5.8);
  const [sf6NominalRating, setSf6NominalRating] = useState(5.8);
  const [differentialProtectionStatus, setDifferentialProtectionStatus] = useState<'Pass' | 'Fail'>('Pass');
  const [safetyInterlocksKeyStatus, setSafetyInterlocksKeyStatus] = useState<'Pass' | 'Fail'>('Pass');
  const [writtenEnergizationPlanApproved, setWrittenEnergizationPlanApproved] = useState(true);
  const [dcBatteryReady, setDcBatteryReady] = useState(true);

  // Stage 2: Energization (7.2.B)
  const [tempProtectiveGroundsRemoved, setTempProtectiveGroundsRemoved] = useState(true);
  const [phasingVoltageA, setPhasingVoltageA] = useState(12);
  const [phasingVoltageB, setPhasingVoltageB] = useState(15);
  const [phasingVoltageC, setPhasingVoltageC] = useState(10);
  const [relayMeteringInputsStatus, setRelayMeteringInputsStatus] = useState<'Pass' | 'Fail'>('Pass');

  // Stage 3: Post-Energization (7.2.C)
  const [thermographicSurveyStatus, setThermographicSurveyStatus] = useState<'Pass' | 'Investigate'>('Pass');
  const [onlinePdTevMaxDb, setOnlinePdTevMaxDb] = useState(14);
  const [onlinePdAcousticMaxDb, setOnlinePdAcousticMaxDb] = useState(1.2);
  const [postEnergizationReportCompleted, setPostEnergizationReportCompleted] = useState(true);

  // UI state
  const [activeStageTab, setActiveStageTab] = useState<'all' | 'stageA' | 'stageB' | 'stageC'>('all');
  const [showPromptModal, setShowPromptModal] = useState(false);
  const [copiedPrompt, setCopiedPrompt] = useState(false);
  const [saveSuccessMessage, setSaveSuccessMessage] = useState<string | null>(null);
  const [isAiRunning, setIsAiRunning] = useState(false);

  // Payload computation
  const currentPayload: NetaEcsMvInputPayload = useMemo(() => {
    const minIns = Math.min(insulationPhaseA, insulationPhaseB, insulationPhaseC);
    const maxPhasing = Math.max(phasingVoltageA, phasingVoltageB, phasingVoltageC);

    return {
      tev_coe_header: {
        module: 'MV-COMMISSIONING-ECS-7.2',
        project_name: projectName,
        system_tag: systemTag,
        nominal_voltage_kV: nominalVoltageKv,
        fse_name: fseName,
        commissioning_date: commissioningDate
      },
      pre_energization_7_2_a: {
        acceptance_test_reports_reviewed: acceptanceTestReportsReviewed,
        mv_busbar_insulation_resistance_5000v_megohms: {
          phase_a_to_ground: insulationPhaseA,
          phase_b_to_ground: insulationPhaseB,
          phase_c_to_ground: insulationPhaseC,
          status: minIns >= 1000 ? 'Pass (Exceeds NETA Table 100.1 minimum threshold)' : `Fail (${minIns} MΩ < 1,000 MΩ Table 100.1)`
        },
        mv_cable_vlf_withstand_table_100_2: {
          test_voltage_kV_rms: vlfTestVoltageKv,
          frequency_hz: 0.1,
          duration_minutes: 30,
          tan_delta_mean_rad: vlfTanDeltaMean,
          status: vlfStatus === 'Pass' ? 'Pass (Complies with NETA Table 100.2 / IEEE 400.2)' : 'Fail'
        },
        bolted_connection_torque_table_100_12: boltedTorqueStatus === 'Pass'
          ? 'Pass (Calibrated torque wrench verified per Table 100.12)'
          : 'Fail',
        sf6_gas_pressure_bar: {
          measured_pressure: sf6MeasuredPressure,
          nominal_rating: sf6NominalRating,
          status: sf6MeasuredPressure >= sf6NominalRating * 0.95 ? 'Pass' : 'Fail (Low SF6 Pressure)'
        },
        differential_protection_primary_injection: differentialProtectionStatus === 'Pass'
          ? 'Pass (Through-fault primary injection test verified)'
          : 'Fail',
        safety_interlocks_and_key_keys: safetyInterlocksKeyStatus === 'Pass'
          ? 'Pass (Key interlocks 100% operational)'
          : 'Fail',
        written_energization_plan_approved: writtenEnergizationPlanApproved,
        dc_auxiliary_power_battery_ok: dcBatteryReady
      },
      energization_7_2_b: {
        temporary_protective_grounds_removed: tempProtectiveGroundsRemoved,
        phasing_check_bus_tie_breaker: {
          phase_a1_a2_voltage_volts: phasingVoltageA,
          phase_b1_b2_voltage_volts: phasingVoltageB,
          phase_c1_c2_voltage_volts: phasingVoltageC,
          status: maxPhasing <= 15 ? 'Pass (In-phase condition confirmed across dual sources)' : 'Fail (Out-of-phase)'
        },
        relay_metering_inputs_status: relayMeteringInputsStatus === 'Pass'
          ? 'Pass (Vector angle and secondary currents verified)'
          : 'Fail'
      },
      post_energization_7_2_c: {
        thermographic_survey_table_100_18: thermographicSurveyStatus === 'Pass'
          ? 'Pass (Normal temperature distribution under load)'
          : 'Investigate',
        online_pd_survey_tables_100_23: {
          tev_max_dB: onlinePdTevMaxDb,
          acoustic_max_dB: onlinePdAcousticMaxDb,
          status: (onlinePdTevMaxDb < 20 && onlinePdAcousticMaxDb < 6) ? 'Pass (Below Warning threshold per Table 100.23.1)' : 'Fail'
        },
        post_energization_gas_fluid_report: postEnergizationReportCompleted
      }
    };
  }, [
    projectName, systemTag, nominalVoltageKv, fseName, commissioningDate,
    acceptanceTestReportsReviewed, insulationPhaseA, insulationPhaseB, insulationPhaseC,
    vlfTestVoltageKv, vlfTanDeltaMean, vlfStatus, boltedTorqueStatus, sf6MeasuredPressure, sf6NominalRating,
    differentialProtectionStatus, safetyInterlocksKeyStatus, writtenEnergizationPlanApproved, dcBatteryReady,
    tempProtectiveGroundsRemoved, phasingVoltageA, phasingVoltageB, phasingVoltageC, relayMeteringInputsStatus,
    thermographicSurveyStatus, onlinePdTevMaxDb, onlinePdAcousticMaxDb, postEnergizationReportCompleted
  ]);

  const evaluationResult: NetaEcsMvEvaluationResult = useMemo(() => {
    return performDeterministicEcsMvAnalysis(currentPayload);
  }, [currentPayload]);

  const handleLoadPreset = (sc: PresetScenario) => {
    setSystemTag(sc.systemTag);
    setNominalVoltageKv(sc.nominalVoltageKv);
    setInsulationPhaseA(sc.insA);
    setInsulationPhaseB(sc.insB);
    setInsulationPhaseC(sc.insC);
    setVlfStatus(sc.isVlfPass ? 'Pass' : 'Fail');
    setVlfTestVoltageKv(sc.vlfKv);
    setVlfTanDeltaMean(sc.tanDelta);
    setBoltedTorqueStatus(sc.isTorquePass ? 'Pass' : 'Fail');
    setSf6MeasuredPressure(sc.sf6Pressure);
    setDifferentialProtectionStatus(sc.isDiffProtPass ? 'Pass' : 'Fail');
    setSafetyInterlocksKeyStatus(sc.isInterlocksPass ? 'Pass' : 'Fail');
    setTempProtectiveGroundsRemoved(sc.isGroundsRemoved);
    setPhasingVoltageA(sc.phasingDeltaV);
    setPhasingVoltageB(sc.phasingDeltaV);
    setPhasingVoltageC(sc.phasingDeltaV);
    setThermographicSurveyStatus(sc.isThermalPass ? 'Pass' : 'Investigate');
    setOnlinePdTevMaxDb(sc.tevMaxDb);
    setOnlinePdAcousticMaxDb(sc.acousticMaxDb);
  };

  const handleRunAiEvaluation = async () => {
    setIsAiRunning(true);
    try {
      const res = await fetch('/api/field-service/neta-ecs-mv-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(currentPayload)
      });
      const json = await res.json();
      if (json.success) {
        setSaveSuccessMessage(isVi ? 'Đã hoàn tất thẩm định AI theo chuẩn ANSI/NETA ECS-2024 Mục 7.2!' : 'AI MV Commissioning Evaluation Completed!');
        setTimeout(() => setSaveSuccessMessage(null), 3500);
      }
    } catch (e) {
      console.warn('AI analysis call:', e);
    } finally {
      setIsAiRunning(false);
    }
  };

  const handleSaveToReports = () => {
    if (onSaveReport) {
      onSaveReport({
        data: currentPayload,
        analysisResult: evaluationResult,
        equipmentId: systemTag,
        date: commissioningDate,
        type: 'ANSI/NETA ECS-2024 MV Commissioning Report'
      });
    }
    setSaveSuccessMessage(isVi ? `Đã lưu hồ sơ nghiệm thu trung thế [${systemTag}] vào CMMS!` : `MV Commissioning record for [${systemTag}] saved!`);
    setTimeout(() => setSaveSuccessMessage(null), 3500);
  };

  const handleExportA4Pdf = () => {
    const doc = new jsPDF('p', 'mm', 'a4');
    const pageWidth = doc.internal.pageSize.getWidth();

    doc.setFillColor(15, 23, 42);
    doc.rect(0, 0, pageWidth, 28, 'F');

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(14);
    doc.setFont('helvetica', 'bold');
    doc.text('CÔNG TY CỔ PHẦN KỸ NGHỆ NĂNG LƯỢNG VIỆT NAM (TEV)', 14, 11);

    doc.setFontSize(8);
    doc.setFont('helvetica', 'normal');
    doc.text('TEV CENTER OF EXCELLENCE (COE) • ELECTRICAL COMMISSIONING SPECIFICATIONS (ECS)', 14, 17);
    doc.text(`REPORT NO: ${reportNo} • DATE: ${commissioningDate} • STANDARD: ANSI/NETA ECS-2024 (SEC 7.2)`, 14, 23);

    doc.setTextColor(15, 23, 42);
    doc.setFontSize(13);
    doc.setFont('helvetica', 'bold');
    doc.text('BIÊN BẢN NGHIỆM THU CHẠY THỬ & ĐÓNG ĐIỆN HỆ THỐNG TRUNG THẾ (> 1 kV - < 100 kV)', 14, 37);

    doc.setFontSize(8.5);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(71, 85, 105);
    doc.text('General Inspection and Test Procedures: Medium-Voltage Systems - NETA ECS-2024 & Tables 100.1, 100.2, 100.12, 100.18, 100.23', 14, 43);

    const metaTable = [
      [`Dự án: ${projectName}`, `Kỹ sư FSE (NETA III): ${fseName}`],
      [`Thiết bị / Tủ trung thế: ${systemTag}`, `Điện áp định mức: ${nominalVoltageKv} kV`],
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

    doc.save(`TEV_ECS_MV_Report_${systemTag.replace(/[^a-zA-Z0-9]/g, '_')}_${commissioningDate}.pdf`);
  };

  return (
    <div className="space-y-6">
      {/* 1. TOP HEADER & COMMISSIONING ACTION BAR */}
      <div className="bg-gradient-to-r from-slate-950 via-teal-950 to-slate-900 text-white p-5 sm:p-6 rounded-2xl shadow-xl border border-teal-500/30">
        <div className="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-4">
          <div className="space-y-1.5">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="px-2.5 py-1 rounded-md text-[11px] font-black uppercase tracking-wider bg-teal-500/30 text-teal-200 border border-teal-400/30">
                NETA ECS-2024 • MỤC 7.2
              </span>
              <span className="px-2 py-0.5 rounded-md text-[10px] font-bold bg-amber-500/20 text-amber-300 border border-amber-500/30">
                Medium-Voltage Systems (&gt; 1 kV &amp; &lt; 100 kV)
              </span>
              <span className="px-2 py-0.5 rounded-md text-[10px] font-mono text-slate-300 bg-slate-800/80 border border-slate-700">
                Tables 100.1, 100.2, 100.12, 100.18, 100.23
              </span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white flex items-center gap-2.5">
              <Activity className="text-teal-400" size={26} />
              <span>
                {isVi
                  ? 'Quy Trình Kiểm Tra & Chạy Thử Hệ Thống Điện Trung Thế (1 kV - 100 kV)'
                  : 'General Inspection & Test Procedures - Medium-Voltage Systems (> 1 kV & < 100 kV)'}
              </span>
            </h1>
            <p className="text-xs text-slate-300 max-w-3xl leading-relaxed">
              {isVi
                ? 'Quy chuẩn đóng điện nghiệm thu trung thế 3 giai đoạn theo ANSI/NETA ECS-2024: Kiểm tra cách điện 5000V (Table 100.1), cáp VLF/Tan Delta (Table 100.2), đồng pha hai nguồn máy cắt liên lạc, chụp nhiệt IR và đo phóng điện cục bộ online (Table 100.23).'
                : '3-stage MV commissioning procedure per ANSI/NETA ECS-2024 Section 7.2: 5000V insulation, VLF cable withstand, bus tie phasing, under-load IR thermography, and online PD monitoring.'}
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
            <span>Kịch bản mẫu kiểm thử trung thế:</span>
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
                  sc.id.includes('critical') ? 'bg-rose-500 animate-pulse' : sc.id.includes('caution') ? 'bg-amber-400' : 'bg-emerald-400'
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

      {/* 2. OVERALL SEVERITY BANNER */}
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
                NETA ECS-2024 Sec 7.2
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
              evaluationResult.technical_analysis.insulation_min_val < 1000 ? 'text-rose-600' : 'text-emerald-700'
            }`}>
              {evaluationResult.technical_analysis.insulation_min_val} MΩ
            </span>
          </div>
          <div className="h-8 w-px bg-slate-200 mx-1 hidden sm:block" />
          <div className="text-right">
            <span className="text-[10px] text-slate-400 block uppercase font-bold">Đồng Pha ΔV</span>
            <span className={`text-base font-black font-mono ${
              evaluationResult.technical_analysis.phasing_delta_v > 15 ? 'text-rose-600' : 'text-emerald-700'
            }`}>
              {evaluationResult.technical_analysis.phasing_delta_v} V
            </span>
          </div>
          <div className="h-8 w-px bg-slate-200 mx-1 hidden sm:block" />
          <div className="text-right">
            <span className="text-[10px] text-slate-400 block uppercase font-bold">PD TEV</span>
            <span className={`text-base font-black font-mono ${
              evaluationResult.technical_analysis.tev_val >= 20 ? 'text-amber-600' : 'text-slate-800'
            }`}>
              {evaluationResult.technical_analysis.tev_val} dB
            </span>
          </div>
        </div>
      </div>

      {/* 3. INPUT FORM GRID */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left Col: Identification */}
        <div className="lg:col-span-1 space-y-5">
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h3 className="text-sm font-black text-slate-900 flex items-center gap-2">
                <FileCheck2 size={16} className="text-teal-600" />
                <span>Định Danh Tủ &amp; Thiết Bị Trung Thế</span>
              </h3>
              <span className="text-[10px] px-2 py-0.5 rounded bg-slate-100 font-mono text-slate-600">
                1kV - 100kV
              </span>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="block text-slate-600 font-semibold mb-1">Tên Dự án / Trạm biến áp:</label>
                <input
                  type="text"
                  value={projectName}
                  onChange={(e) => setProjectName(e.target.value)}
                  className="w-full p-2 rounded-lg border border-slate-200 text-xs focus:ring-1 focus:ring-teal-500 focus:outline-hidden"
                />
              </div>

              <div>
                <label className="block text-slate-600 font-semibold mb-1">Mã Ngăn Lộ / Tủ Trung Thế (MV Switchgear):</label>
                <input
                  type="text"
                  value={systemTag}
                  onChange={(e) => setSystemTag(e.target.value)}
                  className="w-full p-2 rounded-lg border border-slate-200 text-xs font-bold text-slate-900 focus:ring-1 focus:ring-teal-500 focus:outline-hidden"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-slate-600 font-semibold mb-1">Điện Áp Định Mức (kV):</label>
                  <input
                    type="number"
                    value={nominalVoltageKv}
                    onChange={(e) => setNominalVoltageKv(Number(e.target.value))}
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
                <input
                  type="text"
                  value={fseName}
                  onChange={(e) => setFseName(e.target.value)}
                  className="w-full p-2 rounded-lg border border-slate-200 text-xs focus:ring-1 focus:ring-teal-500 focus:outline-hidden"
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

          {/* Reference Standards */}
          <div className="bg-gradient-to-br from-slate-900 to-teal-950 text-white p-5 rounded-2xl border border-teal-500/30 shadow-sm space-y-3 text-xs">
            <h4 className="font-bold text-teal-300 flex items-center gap-2">
              <Info size={15} />
              <span>Tiêu Chuẩn NETA ECS Mục 7.2 Trung Thế</span>
            </h4>
            <ul className="space-y-2 text-slate-300">
              <li className="flex items-start gap-2">
                <span className="text-teal-400 font-bold shrink-0">•</span>
                <span><strong>NETA Table 100.1</strong>: Điện trở cách điện thanh cái 15kV/22kV thử ở 5000V DC đạt tối thiểu <strong className="text-white">≥ 1,000 MΩ</strong>.</span>
              </li>
              <li className="flex items-start gap-2">
                <span className="text-teal-400 font-bold shrink-0">•</span>
                <span><strong>NETA Table 100.2 / IEEE 400.2</strong>: Thử nghiệm cao áp tần số thấp VLF 0.1Hz và tổn hao Tan Delta cho cáp ngầm trung thế.</span>
              </li>
              <li className="flex items-start gap-2">
                <span className="text-teal-400 font-bold shrink-0">•</span>
                <span><strong>NETA Sec 7.2.B.8</strong>: Đồng pha 2 nguồn máy cắt liên lạc tie breaker: Chênh áp <strong className="text-white">≤ 15V</strong>.</span>
              </li>
              <li className="flex items-start gap-2">
                <span className="text-teal-400 font-bold shrink-0">•</span>
                <span><strong>NETA Table 100.23</strong>: Khảo sát PD Online bằng TEV &amp; Siêu âm trước khi bàn giao.</span>
              </li>
            </ul>
          </div>
        </div>

        {/* Right Col: 3 Stages Form */}
        <div className="lg:col-span-2 space-y-5">
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
              Tất Cả 3 Giai Đoạn (14 Mục)
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
              GĐ 1: Tiền Đóng Điện (7.2.A)
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
              GĐ 2: Trong Đóng Điện (7.2.B)
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
              GĐ 3: Sau Đóng Điện (7.2.C)
            </button>
          </div>

          {/* GIAI ĐOẠN 1: TIỀN ĐÓNG ĐIỆN 7.2.A */}
          {(activeStageTab === 'all' || activeStageTab === 'stageA') && (
            <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <h3 className="text-sm font-black text-slate-900 flex items-center gap-2">
                  <span className="w-6 h-6 rounded-lg bg-teal-100 text-teal-800 flex items-center justify-center font-bold text-xs">
                    A
                  </span>
                  <span>GIAI ĐOẠN 1: TIỀN CHẠY THỬ TRUNG THẾ (NETA ECS 7.2.A)</span>
                </h3>
                <span className="text-[11px] font-bold text-teal-700 bg-teal-50 px-2.5 py-0.5 rounded-md border border-teal-200">
                  Pre-Energization MV
                </span>
              </div>

              <div className="space-y-4 text-xs">
                {/* A1: Documentation */}
                <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                  <div>
                    <span className="font-bold text-slate-900 block">A1. Rà soát Hồ sơ, Bản vẽ &amp; Nhãn mác MV</span>
                    <span className="text-[11px] text-slate-500">NETA ECS Section 7.2.A.1 - 7.2.A.4: Khớp sơ đồ 1 sợi, catalog và Acceptance Testing.</span>
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

                {/* A2: MV Busbar Insulation (5000V DC Table 100.1) */}
                <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
                  <div className="flex items-center justify-between flex-wrap gap-2">
                    <div>
                      <span className="font-bold text-slate-900 block">
                        A2. Điện trở Cách điện Thanh cái &amp; Thiết bị MV (5000V DC)
                      </span>
                      <span className="text-[11px] text-slate-500">
                        NETA ATS Table 100.1: Tiêu chuẩn bắt buộc tối thiểu <strong className="text-indigo-900">≥ 1,000 MΩ</strong> đối với hệ 15kV–35kV.
                      </span>
                    </div>
                    {Math.min(insulationPhaseA, insulationPhaseB, insulationPhaseC) < 1000 && (
                      <span className="px-2 py-0.5 rounded text-[10px] font-black bg-rose-100 text-rose-800 border border-rose-300 animate-pulse">
                        CẢNH BÁO: CÁCH ĐIỆN &lt; 1,000 MΩ (NGUY CƠ PHÓNG ĐIỆN)
                      </span>
                    )}
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-1">
                    <div className="p-2 rounded-lg bg-white border border-slate-200">
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">Pha A - Đất (MΩ):</label>
                      <input
                        type="number"
                        value={insulationPhaseA}
                        onChange={(e) => setInsulationPhaseA(Number(e.target.value))}
                        className={`w-full p-1.5 rounded border text-xs font-mono font-bold ${
                          insulationPhaseA < 1000 ? 'border-rose-400 bg-rose-50 text-rose-800' : 'border-slate-200'
                        }`}
                      />
                    </div>
                    <div className="p-2 rounded-lg bg-white border border-slate-200">
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">Pha B - Đất (MΩ):</label>
                      <input
                        type="number"
                        value={insulationPhaseB}
                        onChange={(e) => setInsulationPhaseB(Number(e.target.value))}
                        className={`w-full p-1.5 rounded border text-xs font-mono font-bold ${
                          insulationPhaseB < 1000 ? 'border-rose-400 bg-rose-50 text-rose-800' : 'border-slate-200'
                        }`}
                      />
                    </div>
                    <div className="p-2 rounded-lg bg-white border border-slate-200">
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">Pha C - Đất (MΩ):</label>
                      <input
                        type="number"
                        value={insulationPhaseC}
                        onChange={(e) => setInsulationPhaseC(Number(e.target.value))}
                        className={`w-full p-1.5 rounded border text-xs font-mono font-bold ${
                          insulationPhaseC < 1000 ? 'border-rose-400 bg-rose-50 text-rose-800' : 'border-slate-200'
                        }`}
                      />
                    </div>
                  </div>
                </div>

                {/* A3: VLF Cable Test Table 100.2 */}
                <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
                  <div className="flex items-center justify-between flex-wrap gap-2">
                    <div>
                      <span className="font-bold text-slate-900 block">
                        A3. Thử nghiệm Cáp Ngầm Trung thế VLF &amp; Tan Delta (NETA Table 100.2)
                      </span>
                      <span className="text-[11px] text-slate-500">
                        IEEE 400.2: Thử chịu đựng điện áp cao tần số thấp 0.1Hz trong 30-60 phút.
                      </span>
                    </div>
                    <select
                      value={vlfStatus}
                      onChange={(e) => setVlfStatus(e.target.value as any)}
                      className="p-1.5 rounded-lg border border-slate-200 text-xs font-bold"
                    >
                      <option value="Pass">Pass - Đạt điện áp VLF &amp; Tan Delta</option>
                      <option value="Fail">Fail - Phóng điện đánh thủng / Tan Delta cao</option>
                    </select>
                  </div>

                  <div className="grid grid-cols-2 gap-3 pt-1">
                    <div className="p-2 rounded-lg bg-white border border-slate-200">
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">Điện Áp Thử VLF (kV RMS):</label>
                      <input
                        type="number"
                        step="0.5"
                        value={vlfTestVoltageKv}
                        onChange={(e) => setVlfTestVoltageKv(Number(e.target.value))}
                        className="w-full p-1.5 rounded border border-slate-200 text-xs font-mono font-bold"
                      />
                    </div>
                    <div className="p-2 rounded-lg bg-white border border-slate-200">
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">Hệ Số Tổn Hao Tan Delta (rad):</label>
                      <input
                        type="number"
                        step="0.0001"
                        value={vlfTanDeltaMean}
                        onChange={(e) => setVlfTanDeltaMean(Number(e.target.value))}
                        className="w-full p-1.5 rounded border border-slate-200 text-xs font-mono font-bold"
                      />
                    </div>
                  </div>
                </div>

                {/* A4 & A5: Torque & SF6 Pressure */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-1.5">
                    <span className="font-bold text-slate-900 block">A4. Lực Siết Bu-lông Thanh cái &amp; Đầu cáp</span>
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
                    <span className="font-bold text-slate-900 block">A5. Áp Suất Khí SF6 Máy Cắt / GIS</span>
                    <span className="text-[11px] text-slate-500 block">NETA ECS Section 7.2.A.35 (Định mức: {sf6NominalRating} bar).</span>
                    <div className="flex items-center gap-2">
                      <input
                        type="number"
                        step="0.1"
                        value={sf6MeasuredPressure}
                        onChange={(e) => setSf6MeasuredPressure(Number(e.target.value))}
                        className={`w-full p-2 rounded-lg border text-xs font-mono font-bold ${
                          sf6MeasuredPressure < sf6NominalRating * 0.95 ? 'border-amber-400 bg-amber-50 text-amber-800' : 'border-slate-200'
                        }`}
                      />
                      <span className="font-bold text-slate-700">bar</span>
                    </div>
                  </div>
                </div>

                {/* A6, A7, A8 */}
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-1.5">
                    <span className="font-bold text-slate-900 block">A6. Bảo Vệ So Lệch &amp; CT/VT</span>
                    <span className="text-[11px] text-slate-500 block">Bơm dòng sơ cấp Primary injection.</span>
                    <select
                      value={differentialProtectionStatus}
                      onChange={(e) => setDifferentialProtectionStatus(e.target.value as any)}
                      className="w-full p-2 rounded-lg border border-slate-200 text-xs font-bold"
                    >
                      <option value="Pass">Pass - Bơm dòng sơ cấp đạt</option>
                      <option value="Fail">Fail - Lỗi cực tính CT/So lệch</option>
                    </select>
                  </div>

                  <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-1.5">
                    <span className="font-bold text-slate-900 block">A7. Liên Động Khóa Kirk-Key</span>
                    <span className="text-[11px] text-slate-500 block">Khóa cơ/điện &amp; Arc Energy.</span>
                    <select
                      value={safetyInterlocksKeyStatus}
                      onChange={(e) => setSafetyInterlocksKeyStatus(e.target.value as any)}
                      className="w-full p-2 rounded-lg border border-slate-200 text-xs font-bold"
                    >
                      <option value="Pass">Pass - Liên động 100% tin cậy</option>
                      <option value="Fail">Fail - Kẹt liên động an toàn</option>
                    </select>
                  </div>

                  <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-1.5">
                    <span className="font-bold text-slate-900 block">A8. Kế Hoạch Đóng Điện</span>
                    <span className="text-[11px] text-slate-500 block">Duyệt văn bản &amp; Nguồn DC.</span>
                    <label className="flex items-center gap-2 cursor-pointer pt-1">
                      <input
                        type="checkbox"
                        checked={writtenEnergizationPlanApproved}
                        onChange={(e) => setWrittenEnergizationPlanApproved(e.target.checked)}
                        className="w-4 h-4 accent-teal-600 rounded"
                      />
                      <span className="font-bold text-slate-800">Đã ký duyệt kế hoạch</span>
                    </label>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* GIAI ĐOẠN 2: TRONG KHI ĐÓNG ĐIỆN 7.2.B */}
          {(activeStageTab === 'all' || activeStageTab === 'stageB') && (
            <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <h3 className="text-sm font-black text-slate-900 flex items-center gap-2">
                  <span className="w-6 h-6 rounded-lg bg-blue-100 text-blue-800 flex items-center justify-center font-bold text-xs">
                    B
                  </span>
                  <span>GIAI ĐOẠN 2: TRONG KHI ĐÓNG ĐIỆN TRUNG THẾ (NETA ECS 7.2.B)</span>
                </h3>
                <span className="text-[11px] font-bold text-blue-700 bg-blue-50 px-2.5 py-0.5 rounded-md border border-blue-200">
                  Energization MV
                </span>
              </div>

              <div className="space-y-4 text-xs">
                {/* B1: Temporary Grounds Removal */}
                <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <span className="font-bold text-slate-900 block">
                      B1. Xác Nhận Tháo Dỡ 100% Tiếp Địa Di Động Tạm Thời (7.2.B.1 &amp; 7.2.B.3)
                    </span>
                    <span className="text-[11px] text-slate-500">
                      Tuyệt đối không còn dây tiếp địa tạm trên thanh cái, máy cắt hay đầu cáp trung thế.
                    </span>
                  </div>
                  <label className="flex items-center gap-2 cursor-pointer p-2 rounded-lg bg-white border border-slate-200 shrink-0">
                    <input
                      type="checkbox"
                      checked={tempProtectiveGroundsRemoved}
                      onChange={(e) => setTempProtectiveGroundsRemoved(e.target.checked)}
                      className="w-4 h-4 accent-teal-600 rounded"
                    />
                    <span className={`font-bold ${!tempProtectiveGroundsRemoved ? 'text-rose-600' : 'text-slate-800'}`}>
                      {tempProtectiveGroundsRemoved ? '✔ Đã tháo bỏ 100% tiếp địa tạm' : '❌ NGUY HIỂM: CÒN TIẾP ĐỊA!'}
                    </span>
                  </label>
                </div>

                {/* B2: Phasing Checks */}
                <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
                  <div className="flex items-center justify-between flex-wrap gap-2">
                    <div>
                      <span className="font-bold text-slate-900 block">
                        B2. Đo Đồng Pha Hai Nguồn - Máy Cắt Liên Lạc (NETA ECS 7.2.B.8 &amp; 7.2.B.9)
                      </span>
                      <span className="text-[11px] text-slate-500">
                        Đo điện áp chênh lệch giữa hai nguồn độc lập tại máy cắt liên lạc (Tie Breaker). Tiêu chuẩn: <strong className="text-indigo-900">ΔV ≤ 15V</strong>.
                      </span>
                    </div>
                    {Math.max(phasingVoltageA, phasingVoltageB, phasingVoltageC) > 15 && (
                      <span className="px-2 py-0.5 rounded text-[10px] font-black bg-rose-100 text-rose-800 border border-rose-300 animate-pulse">
                        CẢNH BÁO: LỆCH PHA ({Math.max(phasingVoltageA, phasingVoltageB, phasingVoltageC)}V &gt; 15V) - CẤM ĐÓNG HÒA
                      </span>
                    )}
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-1">
                    <div className="p-2 rounded-lg bg-white border border-slate-200">
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">Chênh Áp Pha A1-A2 (V):</label>
                      <input
                        type="number"
                        value={phasingVoltageA}
                        onChange={(e) => setPhasingVoltageA(Number(e.target.value))}
                        className={`w-full p-1.5 rounded border text-xs font-mono font-bold ${
                          phasingVoltageA > 15 ? 'border-rose-400 bg-rose-50 text-rose-800' : 'border-slate-200'
                        }`}
                      />
                    </div>
                    <div className="p-2 rounded-lg bg-white border border-slate-200">
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">Chênh Áp Pha B1-B2 (V):</label>
                      <input
                        type="number"
                        value={phasingVoltageB}
                        onChange={(e) => setPhasingVoltageB(Number(e.target.value))}
                        className={`w-full p-1.5 rounded border text-xs font-mono font-bold ${
                          phasingVoltageB > 15 ? 'border-rose-400 bg-rose-50 text-rose-800' : 'border-slate-200'
                        }`}
                      />
                    </div>
                    <div className="p-2 rounded-lg bg-white border border-slate-200">
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">Chênh Áp Pha C1-C2 (V):</label>
                      <input
                        type="number"
                        value={phasingVoltageC}
                        onChange={(e) => setPhasingVoltageC(Number(e.target.value))}
                        className={`w-full p-1.5 rounded border text-xs font-mono font-bold ${
                          phasingVoltageC > 15 ? 'border-rose-400 bg-rose-50 text-rose-800' : 'border-slate-200'
                        }`}
                      />
                    </div>
                  </div>
                </div>

                {/* B3: Relay inputs */}
                <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <span className="font-bold text-slate-900 block">
                      B3. Tín Hiệu Nhị Thứ Rơ-le &amp; Đồng Hồ Đo Lường (7.2.B.7 &amp; 7.2.B.10)
                    </span>
                    <span className="text-[11px] text-slate-500">
                      Góc vector pha dòng/áp nhị thứ khớp 100% chiều công suất và dòng tải thực tế.
                    </span>
                  </div>
                  <select
                    value={relayMeteringInputsStatus}
                    onChange={(e) => setRelayMeteringInputsStatus(e.target.value as any)}
                    className="p-2 rounded-lg border border-slate-200 text-xs font-bold shrink-0"
                  >
                    <option value="Pass">Pass - Vector dòng/áp khớp tải</option>
                    <option value="Fail">Fail - Nhầm cực tính CT/VT</option>
                  </select>
                </div>
              </div>
            </div>
          )}

          {/* GIAI ĐOẠN 3: SAU KHI ĐÓNG ĐIỆN 7.2.C */}
          {(activeStageTab === 'all' || activeStageTab === 'stageC') && (
            <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <h3 className="text-sm font-black text-slate-900 flex items-center gap-2">
                  <span className="w-6 h-6 rounded-lg bg-amber-100 text-amber-800 flex items-center justify-center font-bold text-xs">
                    C
                  </span>
                  <span>GIAI ĐOẠN 3: SAU KHI ĐÓNG ĐIỆN MANG TẢI TRUNG THẾ (NETA ECS 7.2.C)</span>
                </h3>
                <span className="text-[11px] font-bold text-amber-700 bg-amber-50 px-2.5 py-0.5 rounded-md border border-amber-200">
                  Post-Energization MV
                </span>
              </div>

              <div className="space-y-4 text-xs">
                {/* C1: IR Thermography Table 100.18 */}
                <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <span className="font-bold text-slate-900 block">
                      C1. Khảo Sát Nhiệt Hồng Ngoại Mang Tải (NETA Table 100.18)
                    </span>
                    <span className="text-[11px] text-slate-500">
                      NETA ECS 7.2.C.3: Chụp ảnh nhiệt khoang thanh cái, máy cắt và sứ xuyên dưới 70% tải.
                    </span>
                  </div>
                  <select
                    value={thermographicSurveyStatus}
                    onChange={(e) => setThermographicSurveyStatus(e.target.value as any)}
                    className="p-2 rounded-lg border border-slate-200 text-xs font-bold shrink-0"
                  >
                    <option value="Pass">Pass - Nhiệt độ đều, 0 Hotspot</option>
                    <option value="Investigate">Investigate - Xuất hiện Hotspot bất thường</option>
                  </select>
                </div>

                {/* C2: Online Partial Discharge Survey Table 100.23 */}
                <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
                  <div className="flex items-center justify-between flex-wrap gap-2">
                    <div>
                      <span className="font-bold text-slate-900 block">
                        C2. Khảo Sát Phóng Điện Cục Bộ Online (NETA MTS Tables 100.23.1 - 100.23.5)
                      </span>
                      <span className="text-[11px] text-slate-500">
                        Đo cảm biến TEV (Chuẩn cảnh báo &ge; 20 dBmV) và Siêu âm Acoustic (Chuẩn cảnh báo &ge; 6 dBµV).
                      </span>
                    </div>
                    <span className={`px-2 py-0.5 rounded text-[10px] font-black ${
                      onlinePdTevMaxDb >= 20 || onlinePdAcousticMaxDb >= 6
                        ? 'bg-amber-100 text-amber-800 border border-amber-300'
                        : 'bg-emerald-100 text-emerald-800'
                    }`}>
                      {onlinePdTevMaxDb >= 20 || onlinePdAcousticMaxDb >= 6 ? 'CẢNH BÁO PD CAO' : 'PD AN TOÀN'}
                    </span>
                  </div>

                  <div className="grid grid-cols-2 gap-3 pt-1">
                    <div className="p-2 rounded-lg bg-white border border-slate-200">
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">TEV Cực Đại (dBmV):</label>
                      <input
                        type="number"
                        value={onlinePdTevMaxDb}
                        onChange={(e) => setOnlinePdTevMaxDb(Number(e.target.value))}
                        className={`w-full p-1.5 rounded border text-xs font-mono font-bold ${
                          onlinePdTevMaxDb >= 20 ? 'border-amber-400 bg-amber-50 text-amber-800' : 'border-slate-200'
                        }`}
                      />
                    </div>
                    <div className="p-2 rounded-lg bg-white border border-slate-200">
                      <label className="block text-[11px] font-bold text-slate-600 mb-1">Siêu Âm Acoustic (dBµV):</label>
                      <input
                        type="number"
                        step="0.1"
                        value={onlinePdAcousticMaxDb}
                        onChange={(e) => setOnlinePdAcousticMaxDb(Number(e.target.value))}
                        className={`w-full p-1.5 rounded border text-xs font-mono font-bold ${
                          onlinePdAcousticMaxDb >= 6 ? 'border-amber-400 bg-amber-50 text-amber-800' : 'border-slate-200'
                        }`}
                      />
                    </div>
                  </div>
                </div>

                {/* C3: Post-Energization Report */}
                <div className="p-3 rounded-xl bg-slate-50 border border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <span className="font-bold text-slate-900 block">
                      C3. Lấy Mẫu Khí/Dầu Sau 24h &amp; Phát Hành Báo Cáo Nghiệm Thu (Section 5.4)
                    </span>
                    <span className="text-[11px] text-slate-500">
                      Hoàn thiện bộ hồ sơ bàn giao chính thức cho Chủ đầu tư và đơn vị vận hành.
                    </span>
                  </div>
                  <label className="flex items-center gap-2 cursor-pointer shrink-0">
                    <input
                      type="checkbox"
                      checked={postEnergizationReportCompleted}
                      onChange={(e) => setPostEnergizationReportCompleted(e.target.checked)}
                      className="w-4 h-4 accent-teal-600 rounded"
                    />
                    <span className="font-bold text-slate-800">Báo cáo nghiệm thu hoàn tất</span>
                  </label>
                </div>
              </div>
            </div>
          )}
        </div>
      </div>

      {/* 4. DETAILED EVALUATION TABLE */}
      <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
        <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 border-b border-slate-100 pb-3">
          <div>
            <h3 className="text-sm font-black text-slate-900 flex items-center gap-2">
              <Layers size={17} className="text-teal-600" />
              <span>BẢNG PHÂN TÍCH CHI TIẾT NGHIỆM THU TRUNG THẾ (NETA ECS SECTION 7.2)</span>
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Đối chiếu 14 hạng mục chạy thử trung thế theo ANSI/NETA ECS-2024 &amp; các bảng Tables 100.1, 100.2, 100.12, 100.18, 100.23
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

      {/* 5. TECHNICAL HAZARD ASSESSMENT & RECOMMENDATIONS */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-3">
          <div className="flex items-center gap-2 border-b border-slate-100 pb-3">
            <AlertOctagon size={18} className={evaluationResult.technical_analysis.hazard_identified ? 'text-rose-600' : 'text-emerald-600'} />
            <h3 className="text-sm font-black text-slate-900">
              {isVi ? 'Phân Tích Kỹ Thuật & Đánh Giá Hệ Thống Trung Thế' : 'Technical Analysis & Hazard Assessment'}
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
            <span className="font-bold text-slate-700 block">Thông Số Kiểm Soát Nòng Cốt:</span>
            <div className="grid grid-cols-3 gap-2 text-center pt-1 font-mono">
              <div className="p-1.5 bg-white rounded border border-slate-200">
                <span className="text-[10px] text-slate-400 block font-sans">Cách điện 5kV</span>
                <span className={`font-black ${evaluationResult.technical_analysis.insulation_min_val < 1000 ? 'text-rose-600' : 'text-emerald-700'}`}>
                  {evaluationResult.technical_analysis.insulation_min_val} MΩ
                </span>
              </div>
              <div className="p-1.5 bg-white rounded border border-slate-200">
                <span className="text-[10px] text-slate-400 block font-sans">Đồng Pha ΔV</span>
                <span className={`font-black ${evaluationResult.technical_analysis.phasing_delta_v > 15 ? 'text-rose-600' : 'text-emerald-700'}`}>
                  {evaluationResult.technical_analysis.phasing_delta_v} V
                </span>
              </div>
              <div className="p-1.5 bg-white rounded border border-slate-200">
                <span className="text-[10px] text-slate-400 block font-sans">PD Online</span>
                <span className={`font-black ${evaluationResult.technical_analysis.tev_val >= 20 ? 'text-amber-600' : 'text-emerald-700'}`}>
                  {evaluationResult.technical_analysis.tev_val} dB
                </span>
              </div>
            </div>
          </div>
        </div>

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

      {/* 6. MODAL: SYSTEM INSTRUCTION PROMPT */}
      {showPromptModal && (
        <div className="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-slate-900 text-white border border-teal-500/40 rounded-2xl max-w-3xl w-full max-h-[85vh] flex flex-col shadow-2xl">
            <div className="p-4 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Sparkles size={18} className="text-amber-400" />
                <h3 className="font-bold text-sm text-white">
                  Master System Instruction Prompt (ANSI/NETA ECS-2024 Section 7.2 MV)
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
              {ECS_MV_SYSTEM_INSTRUCTION}
            </div>

            <div className="p-4 border-t border-slate-800 flex items-center justify-between">
              <span className="text-xs text-slate-400">
                Nhấp nút Copy rồi dán trực tiếp vào mục "System Instructions" trên Google AI Studio.
              </span>
              <button
                type="button"
                onClick={() => {
                  navigator.clipboard.writeText(ECS_MV_SYSTEM_INSTRUCTION);
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
