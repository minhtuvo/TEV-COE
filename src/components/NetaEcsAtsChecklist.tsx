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
  Gauge,
  Thermometer,
  Flame,
  Globe,
  ClipboardCheck,
  Search,
  Activity,
  GitBranch,
  Radio,
  Clock,
  Sliders,
  Cable,
  Power
} from 'lucide-react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  NetaEcsAtsInputPayload,
  NetaEcsAtsEvaluationResult,
  performDeterministicEcsAtsAnalysis,
  ECS_ATS_SYSTEM_INSTRUCTION
} from '../../server/netaEcsAtsAnalyzer';

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
  voltageRatingV: number;
  currentRatingA: number;
  controlIr: number;
  s1Poles: { a: number; b: number; c: number; n: number };
  s2Poles: { a: number; b: number; c: number; n: number };
  isInterlocksPass: boolean;
  isPhasingPass: boolean;
  totalTransferTimeSec: number;
  isRetransferPass: boolean;
  isSinglePhasePass: boolean;
  isTorquePass: boolean;
  isPlanApproved: boolean;
  isThermalPass: boolean;
  thermalDeltaT?: number;
  description: string;
}

const PRESET_SCENARIOS: PresetScenario[] = [
  {
    id: 'critical-contact-anomaly',
    name: '🔴 NGUY HẠI ĐÓNG ĐIỆN: Điện Trở Tiếp Điểm Nguồn 2 Pha B Quá Cao (185 µΩ - Mẫu Chuẩn AI Studio)',
    nameEn: 'CRITICAL HAZARD: High Contact Resistance on Source 2 Phase B (185 µOhm - Benchmark Prompt)',
    badge: 'CRITICAL / NGUY HIỂM',
    badgeColor: 'bg-rose-500/20 text-rose-300 border-rose-500/40',
    systemTag: 'ATS-CRITICAL-LOAD-01 (400V 800A 4-Pole ATS)',
    voltageRatingV: 400,
    currentRatingA: 800,
    controlIr: 125.0,
    s1Poles: { a: 32, b: 34, c: 31, n: 33 },
    s2Poles: { a: 35, b: 185, c: 34, n: 36 }, // 185 µΩ vs 34 µΩ -> > 400% deviation!
    isInterlocksPass: true,
    isPhasingPass: true,
    totalTransferTimeSec: 8.8, // <= 10s NFPA 110 Class 10
    isRetransferPass: true,
    isSinglePhasePass: true,
    isTorquePass: true,
    isPlanApproved: true,
    isThermalPass: true,
    description: 'Pha B Nguồn 2 đo được 185 µΩ (lệch > 400% so với 34 µΩ). Cần khóa ATS ở Nguồn 1, không cho chuyển sang máy phát.'
  },
  {
    id: 'all-pass-datacenter',
    name: '🟢 ĐẠT CHUẨN 100%: Tủ ATS Data Center 400V 800A (NFPA 110 Class 10 Chuyển Nguồn 8.2s)',
    nameEn: '100% PASS: Data Center 400V 800A 4-Pole ATS (NFPA 110 Class 10 Compliant)',
    badge: 'PASS / ĐẠT CHUẨN',
    badgeColor: 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40',
    systemTag: 'ATS-DATACENTER-MAIN-02',
    voltageRatingV: 400,
    currentRatingA: 800,
    controlIr: 240.0,
    s1Poles: { a: 31, b: 33, c: 32, n: 32 },
    s2Poles: { a: 33, b: 35, c: 32, n: 34 },
    isInterlocksPass: true,
    isPhasingPass: true,
    totalTransferTimeSec: 8.2,
    isRetransferPass: true,
    isSinglePhasePass: true,
    isTorquePass: true,
    isPlanApproved: true,
    isThermalPass: true,
    description: 'Tiếp điểm đồng đều lệch < 10%, cách điện 240 MΩ, chuyển nguồn 8.2s, trả nguồn 300s, làm mát 240s, đạt chuẩn NFPA 110 & NETA ATS.'
  },
  {
    id: 'critical-interlock-failure',
    name: '🔴 NGUY HẠI: Hỏng Khóa Liên Động Cơ Điện & Không Nhận Biết Mất Pha (IEEE 446)',
    nameEn: 'CRITICAL: Mechanical Interlock Failure & Single-Phase Loss Detection Blindspot',
    badge: 'CRITICAL / NGUY HIỂM',
    badgeColor: 'bg-rose-500/20 text-rose-300 border-rose-500/40',
    systemTag: 'ATS-HOSPITAL-EMERGENCY-01',
    voltageRatingV: 400,
    currentRatingA: 1200,
    controlIr: 45.0,
    s1Poles: { a: 28, b: 29, c: 27, n: 28 },
    s2Poles: { a: 30, b: 31, c: 29, n: 30 },
    isInterlocksPass: false, // Critical failure
    isPhasingPass: true,
    totalTransferTimeSec: 9.5,
    isRetransferPass: true,
    isSinglePhasePass: false, // Missed single-phase loss
    isTorquePass: true,
    isPlanApproved: false,
    isThermalPass: true,
    description: 'Hỏng cơ cấu khóa liên động cơ học (nguy cơ đóng song song 2 nguồn) và bộ điều khiển không nhận biết khi mất 1 pha Nguồn 1.'
  },
  {
    id: 'caution-hotspot-torque',
    name: '🟡 CẢNH BÁO: Điểm Nóng Tiếp Điểm ATS (ΔT = 16.5°C) & Bu-lông Chưa Đủ Lực Siết',
    nameEn: 'CAUTION: Thermal Hotspot Anomaly & Deficient Bolted Torque (Table 100.12)',
    badge: 'CAUTION / CẢNH BÁO',
    badgeColor: 'bg-amber-500/20 text-amber-300 border-amber-500/40',
    systemTag: 'ATS-COMMERCIAL-BUILDING-03',
    voltageRatingV: 400,
    currentRatingA: 630,
    controlIr: 85.0,
    s1Poles: { a: 34, b: 36, c: 33, n: 35 },
    s2Poles: { a: 38, b: 42, c: 37, n: 39 },
    isInterlocksPass: true,
    isPhasingPass: true,
    totalTransferTimeSec: 9.1,
    isRetransferPass: true,
    isSinglePhasePass: true,
    isTorquePass: false, // Deficient torque
    isPlanApproved: true,
    isThermalPass: false, // Hotspot
    thermalDeltaT: 16.5,
    description: 'Tiếp điểm cọc đấu cáp phát nhiệt ΔT = 16.5°C theo NETA MTS Table 100.18 và bu-lông chưa đạt lực siết theo Bảng Table 100.12.'
  }
];

export const NetaEcsAtsChecklist: React.FC<Props> = ({
  onSaveReport,
  onNavigateToReports,
  initialLanguage = 'vi'
}) => {
  const [lang, setLang] = useState<'vi' | 'en'>(initialLanguage);
  const isVi = lang === 'vi';

  // Header metadata
  const [projectName, setProjectName] = useState('Toa Nha Trung Tam Du Lieu TEV Data Center');
  const [systemTag, setSystemTag] = useState('ATS-CRITICAL-LOAD-01 (400V 800A 4-Pole ATS)');
  const [voltageRatingV, setVoltageRatingV] = useState<number>(400);
  const [currentRatingA, setCurrentRatingA] = useState<number>(800);
  const [fseName, setFseName] = useState('Nguyen Van L');
  const [commissioningDate, setCommissioningDate] = useState('2026-10-17');

  // Phase 1: Pre-Energization (8.3.A)
  const [isAtsReportsReviewed, setIsAtsReportsReviewed] = useState(true);
  const [controlIr, setControlIr] = useState<number>(125.0);
  const [s1PhaseA, setS1PhaseA] = useState<number>(32);
  const [s1PhaseB, setS1PhaseB] = useState<number>(34);
  const [s1PhaseC, setS1PhaseC] = useState<number>(31);
  const [s1Neutral, setS1Neutral] = useState<number>(33);

  const [s2PhaseA, setS2PhaseA] = useState<number>(35);
  const [s2PhaseB, setS2PhaseB] = useState<number>(185);
  const [s2PhaseC, setS2PhaseC] = useState<number>(34);
  const [s2Neutral, setS2Neutral] = useState<number>(36);

  const [isInterlocksPass, setIsInterlocksPass] = useState(true);
  const [engineStartDelaySec, setEngineStartDelaySec] = useState<number>(3.0);
  const [transferDelaySec, setTransferDelaySec] = useState<number>(5.0);
  const [retransferDelaySec, setRetransferDelaySec] = useState<number>(300.0);
  const [coolDownDelaySec, setCoolDownDelaySec] = useState<number>(240.0);

  const [isTorquePass, setIsTorquePass] = useState(true);
  const [isGroundingPass, setIsGroundingPass] = useState(true);
  const [isPlanApproved, setIsPlanApproved] = useState(true);

  // Phase 2: Energization & Transfer Testing (8.3.B)
  const [s1Volts, setS1Volts] = useState<number>(401);
  const [s1Freq, setS1Freq] = useState<number>(50.0);
  const [s2Volts, setS2Volts] = useState<number>(402);
  const [s2Freq, setS2Freq] = useState<number>(50.1);
  const [isPhasingPass, setIsPhasingPass] = useState(true);

  const [engineStartSignalTime, setEngineStartSignalTime] = useState<number>(3.0);
  const [genReadyTime, setGenReadyTime] = useState<number>(6.2);
  const [atsTransferCompleteTime, setAtsTransferCompleteTime] = useState<number>(8.8);
  const [totalTransferTime, setTotalTransferTime] = useState<number>(8.8);

  const [isRetransferSuccess, setIsRetransferSuccess] = useState(true);
  const [isGenShutdownSuccess, setIsGenShutdownSuccess] = useState(true);
  const [isSinglePhasePass, setIsSinglePhasePass] = useState(true);

  // Phase 3: Post-Energization (8.3.C)
  const [isThermalPass, setIsThermalPass] = useState(true);
  const [thermalDeltaT, setThermalDeltaT] = useState<number>(3.8);
  const [isScadaPass, setIsScadaPass] = useState(true);
  const [isTrainingCompleted, setIsTrainingCompleted] = useState(true);

  // Analysis State
  const [evaluationResult, setEvaluationResult] = useState<NetaEcsAtsEvaluationResult | null>(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [activeResultTab, setActiveResultTab] = useState<'summary' | 'table' | 'recommendations' | 'ai-markdown'>('summary');

  // Modals
  const [showSystemPromptModal, setShowSystemPromptModal] = useState(false);
  const [showJsonPromptModal, setShowJsonPromptModal] = useState(false);
  const [copiedPrompt, setCopiedPrompt] = useState(false);
  const [copiedJson, setCopiedJson] = useState(false);

  // Live Contact Resistance Deviation Calculation
  const contactDeviationStats = useMemo(() => {
    const s1 = [s1PhaseA, s1PhaseB, s1PhaseC];
    const s2 = [s2PhaseA, s2PhaseB, s2PhaseC];

    const minS1 = Math.min(...s1);
    const maxS1 = Math.max(...s1);
    const devS1 = minS1 > 0 ? ((maxS1 - minS1) / minS1) * 100 : 0;

    const minS2 = Math.min(...s2);
    const maxS2 = Math.max(...s2);
    const devS2 = minS2 > 0 ? ((maxS2 - minS2) / minS2) * 100 : 0;

    const maxDev = Math.max(devS1, devS2);
    const isPass = maxDev <= 50.0;

    return {
      devS1: parseFloat(devS1.toFixed(1)),
      devS2: parseFloat(devS2.toFixed(1)),
      maxDev: parseFloat(maxDev.toFixed(1)),
      maxS2,
      isPass
    };
  }, [s1PhaseA, s1PhaseB, s1PhaseC, s2PhaseA, s2PhaseB, s2PhaseC]);

  // Build Payload
  const currentPayload: NetaEcsAtsInputPayload = useMemo(() => {
    return {
      tev_coe_header: {
        module: 'ATS-COMMISSIONING-ECS-8.3',
        project_name: projectName,
        system_tag: systemTag,
        voltage_rating_V: voltageRatingV,
        current_rating_A: currentRatingA,
        fse_name: fseName,
        commissioning_date: commissioningDate
      },
      pre_energization_8_3_a: {
        acceptance_test_reports_reviewed_ats_7_22_3: isAtsReportsReviewed,
        control_wiring_insulation_1000v_megohms: controlIr,
        main_contact_resistance_micro_ohms: {
          source_1_normal_poles_uOhm: {
            phase_a: s1PhaseA,
            phase_b: s1PhaseB,
            phase_c: s1PhaseC,
            neutral: s1Neutral
          },
          source_2_emergency_poles_uOhm: {
            phase_a: s2PhaseA,
            phase_b: s2PhaseB,
            phase_c: s2PhaseC,
            neutral: s2Neutral
          },
          max_deviation_percent: contactDeviationStats.maxDev,
          status: contactDeviationStats.isPass ? 'Pass' : `Fail (Contact resistance deviation ${contactDeviationStats.maxDev}% > 50% limit)`
        },
        mechanical_and_electrical_interlocks: isInterlocksPass ? 'Pass (100% interlock verification successful)' : 'Fail (Interlock failure)',
        control_logic_and_timers_configured: {
          engine_start_delay_sec: engineStartDelaySec,
          transfer_time_delay_sec: transferDelaySec,
          retransfer_time_delay_sec: retransferDelaySec,
          engine_cool_down_delay_sec: coolDownDelaySec,
          status: 'Pass (Timers comply with NFPA 110)'
        },
        bolted_connection_torque_table_100_12: isTorquePass ? 'Pass (Calibrated torque wrench verified per Table 100.12)' : 'Fail (Insufficient torque)',
        grounding_and_neutral_bonding: isGroundingPass ? 'Pass' : 'Fail',
        written_energization_plan_approved: isPlanApproved
      },
      energization_8_3_b: {
        source_voltage_and_frequency: {
          source_1_utility_volts: s1Volts,
          source_1_freq_hz: s1Freq,
          source_2_generator_volts: s2Volts,
          source_2_freq_hz: s2Freq,
          status: 'Pass'
        },
        phase_sequence_matching: isPhasingPass ? 'Pass (A-B-C phase sequence identical between Source 1 and Source 2)' : 'Fail (Phase rotation mismatch)',
        automatic_load_transfer_test_nfpa_110: {
          simulated_event: 'Loss of Normal Power (Source 1 open)',
          engine_start_signal_time_sec: engineStartSignalTime,
          generator_voltage_ready_time_sec: genReadyTime,
          ats_transfer_completed_time_sec: atsTransferCompleteTime,
          total_emergency_transfer_time_sec: totalTransferTime,
          status: totalTransferTime <= 10.0 ? `Pass (Total transfer time ${totalTransferTime}s <= 10s NFPA 110 Class 10 threshold)` : 'Fail'
        },
        retransfer_and_cool_down_test: {
          simulated_event: 'Restoration of Normal Power (Source 1 closed)',
          retransfer_timer_active_sec: retransferDelaySec,
          ats_retransfer_to_source_1_completed: isRetransferSuccess,
          engine_cool_down_run_sec: coolDownDelaySec,
          generator_auto_shutdown_successful: isGenShutdownSuccess,
          status: isRetransferSuccess && isGenShutdownSuccess ? 'Pass' : 'Fail'
        },
        single_phase_failure_test: isSinglePhasePass ? 'Pass (ATS detected Phase A loss on Source 1 and initiated transfer)' : 'Fail'
      },
      post_energization_8_3_c: {
        thermographic_survey_table_100_18: {
          max_delta_t_celsius: thermalDeltaT,
          status: isThermalPass ? 'Pass (Normal thermal distribution)' : 'Investigate (Hotspot detected per Table 100.18)'
        },
        scada_bms_telemetry_verified: isScadaPass ? 'Pass' : 'Fail',
        owner_manual_bypass_training_completed: isTrainingCompleted
      }
    };
  }, [
    projectName, systemTag, voltageRatingV, currentRatingA, fseName, commissioningDate,
    isAtsReportsReviewed, controlIr, s1PhaseA, s1PhaseB, s1PhaseC, s1Neutral,
    s2PhaseA, s2PhaseB, s2PhaseC, s2Neutral, contactDeviationStats, isInterlocksPass,
    engineStartDelaySec, transferDelaySec, retransferDelaySec, coolDownDelaySec,
    isTorquePass, isGroundingPass, isPlanApproved,
    s1Volts, s1Freq, s2Volts, s2Freq, isPhasingPass,
    engineStartSignalTime, genReadyTime, atsTransferCompleteTime, totalTransferTime,
    isRetransferSuccess, isGenShutdownSuccess, isSinglePhasePass,
    isThermalPass, thermalDeltaT, isScadaPass, isTrainingCompleted
  ]);

  // Apply Preset Scenario
  const applyPresetScenario = (scenario: PresetScenario) => {
    setSystemTag(scenario.systemTag);
    setVoltageRatingV(scenario.voltageRatingV);
    setCurrentRatingA(scenario.currentRatingA);
    setControlIr(scenario.controlIr);
    setS1PhaseA(scenario.s1Poles.a);
    setS1PhaseB(scenario.s1Poles.b);
    setS1PhaseC(scenario.s1Poles.c);
    setS1Neutral(scenario.s1Poles.n);

    setS2PhaseA(scenario.s2Poles.a);
    setS2PhaseB(scenario.s2Poles.b);
    setS2PhaseC(scenario.s2Poles.c);
    setS2Neutral(scenario.s2Poles.n);

    setIsInterlocksPass(scenario.isInterlocksPass);
    setIsPhasingPass(scenario.isPhasingPass);
    setTotalTransferTime(scenario.totalTransferTimeSec);
    setAtsTransferCompleteTime(scenario.totalTransferTimeSec);
    setIsRetransferSuccess(scenario.isRetransferPass);
    setIsGenShutdownSuccess(scenario.isRetransferPass);
    setIsSinglePhasePass(scenario.isSinglePhasePass);
    setIsTorquePass(scenario.isTorquePass);
    setIsPlanApproved(scenario.isPlanApproved);
    setIsThermalPass(scenario.isThermalPass);
    if (scenario.thermalDeltaT !== undefined) setThermalDeltaT(scenario.thermalDeltaT);

    const updatedPayload: NetaEcsAtsInputPayload = {
      ...currentPayload,
      tev_coe_header: {
        ...currentPayload.tev_coe_header,
        system_tag: scenario.systemTag,
        voltage_rating_V: scenario.voltageRatingV,
        current_rating_A: scenario.currentRatingA
      },
      pre_energization_8_3_a: {
        ...currentPayload.pre_energization_8_3_a,
        control_wiring_insulation_1000v_megohms: scenario.controlIr,
        main_contact_resistance_micro_ohms: {
          source_1_normal_poles_uOhm: {
            phase_a: scenario.s1Poles.a,
            phase_b: scenario.s1Poles.b,
            phase_c: scenario.s1Poles.c,
            neutral: scenario.s1Poles.n
          },
          source_2_emergency_poles_uOhm: {
            phase_a: scenario.s2Poles.a,
            phase_b: scenario.s2Poles.b,
            phase_c: scenario.s2Poles.c,
            neutral: scenario.s2Poles.n
          }
        },
        mechanical_and_electrical_interlocks: scenario.isInterlocksPass ? 'Pass' : 'Fail',
        bolted_connection_torque_table_100_12: scenario.isTorquePass ? 'Pass' : 'Fail',
        written_energization_plan_approved: scenario.isPlanApproved
      },
      energization_8_3_b: {
        ...currentPayload.energization_8_3_b,
        phase_sequence_matching: scenario.isPhasingPass ? 'Pass' : 'Fail',
        automatic_load_transfer_test_nfpa_110: {
          engine_start_signal_time_sec: 3.0,
          generator_voltage_ready_time_sec: 6.2,
          ats_transfer_completed_time_sec: scenario.totalTransferTimeSec,
          total_emergency_transfer_time_sec: scenario.totalTransferTimeSec,
          status: scenario.totalTransferTimeSec <= 10.0 ? 'Pass' : 'Fail'
        },
        single_phase_failure_test: scenario.isSinglePhasePass ? 'Pass' : 'Fail'
      },
      post_energization_8_3_c: {
        ...currentPayload.post_energization_8_3_c,
        thermographic_survey_table_100_18: {
          max_delta_t_celsius: scenario.thermalDeltaT ?? 3.5,
          status: scenario.isThermalPass ? 'Pass' : 'Investigate'
        }
      }
    };

    const res = performDeterministicEcsAtsAnalysis(updatedPayload);
    setEvaluationResult(res);
  };

  // Run Deterministic Analysis
  const handleRunDeterministic = () => {
    setIsAnalyzing(true);
    setTimeout(() => {
      const res = performDeterministicEcsAtsAnalysis(currentPayload);
      setEvaluationResult(res);
      setIsAnalyzing(false);
    }, 250);
  };

  // Run Gemini AI Analysis
  const handleRunAiAnalysis = async () => {
    setIsAnalyzing(true);
    try {
      const res = await fetch('/api/field-service/neta-ecs-ats-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(currentPayload)
      });
      const data = await res.json();
      if (data.success && data.data) {
        setEvaluationResult(data.data);
      } else {
        setEvaluationResult(performDeterministicEcsAtsAnalysis(currentPayload));
      }
    } catch (e) {
      console.error('Failed to run AI analysis, falling back to deterministic:', e);
      setEvaluationResult(performDeterministicEcsAtsAnalysis(currentPayload));
    } finally {
      setIsAnalyzing(false);
    }
  };

  // Export A4 PDF Report
  const handleExportPdf = () => {
    const doc = new jsPDF('p', 'mm', 'a4');
    const result = evaluationResult || performDeterministicEcsAtsAnalysis(currentPayload);

    // Header Background
    doc.setFillColor(15, 23, 42); // slate-900
    doc.rect(0, 0, 210, 36, 'F');

    // Title
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(14.5);
    doc.setFont('helvetica', 'bold');
    doc.text('TEV COE - BIEN BAN CHAY THU & NGHIEM THU TU ATS', 14, 14);

    doc.setFontSize(9);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(203, 213, 225);
    doc.text('ANSI/NETA ECS-2024 Sec 8.3 | NETA ATS-2025 Sec 7.22.3 | NFPA 110 | IEEE 446 | Table 100.1, 100.12, 100.18', 14, 21);

    doc.setTextColor(56, 189, 248); // Sky blue
    doc.text(`TRANG THAI: ${result.overall_status} - ${result.status_title_en}`, 14, 29);

    // Meta block
    doc.setFillColor(248, 250, 252);
    doc.rect(14, 40, 182, 22, 'F');
    doc.setDrawColor(226, 232, 240);
    doc.rect(14, 40, 182, 22, 'S');

    doc.setTextColor(51, 65, 85);
    doc.setFontSize(8.5);
    doc.setFont('helvetica', 'bold');
    doc.text('DU AN (PROJECT):', 18, 46);
    doc.setFont('helvetica', 'normal');
    doc.text(projectName, 58, 46);

    doc.setFont('helvetica', 'bold');
    doc.text('MA THIET BI ATS:', 18, 52);
    doc.setFont('helvetica', 'normal');
    doc.text(`${systemTag} (${voltageRatingV}V, ${currentRatingA}A)`, 58, 52);

    doc.setFont('helvetica', 'bold');
    doc.text('KY SU FSE:', 18, 58);
    doc.setFont('helvetica', 'normal');
    doc.text(`${fseName} | Ngay: ${commissioningDate}`, 58, 58);

    doc.setFont('helvetica', 'bold');
    doc.text('CONTACT RESISTANCE:', 125, 46);
    doc.setFont('helvetica', 'normal');
    doc.text(`Dev: ${contactDeviationStats.maxDev}% (Max <= 50%)`, 165, 46);

    doc.setFont('helvetica', 'bold');
    doc.text('CONTROL WIRING IR:', 125, 52);
    doc.setFont('helvetica', 'normal');
    doc.text(`${controlIr} M (Min >= 1.0 M)`, 165, 52);

    doc.setFont('helvetica', 'bold');
    doc.text('TRANSFER TIME:', 125, 58);
    doc.setFont('helvetica', 'normal');
    doc.text(`${totalTransferTime}s (NFPA 110 <= 10.0s)`, 165, 58);

    // Evaluation Table
    const tableData = result.evaluation_table.map((row) => [
      row.item_code,
      row.item_name_vn.split('/')[0].trim(),
      row.neta_standard_ref,
      row.measured_value,
      row.status,
      row.evaluation_vn
    ]);

    autoTable(doc, {
      startY: 66,
      head: [['Muc', 'Hang Muc Kiem Tra', 'Tieu Chuan Tham Chieu', 'Gia Tri Do', 'Ket Qua', 'Danh Gia Ky Thuat']],
      body: tableData,
      theme: 'grid',
      headStyles: {
        fillColor: [30, 41, 59],
        textColor: [255, 255, 255],
        fontSize: 7.5,
        fontStyle: 'bold'
      },
      styles: {
        fontSize: 7,
        cellPadding: 1.8,
        lineColor: [226, 232, 240]
      },
      columnStyles: {
        0: { cellWidth: 10, halign: 'center' },
        1: { cellWidth: 38 },
        2: { cellWidth: 38 },
        3: { cellWidth: 34 },
        4: { cellWidth: 16, halign: 'center', fontStyle: 'bold' },
        5: { cellWidth: 46 }
      },
      didParseCell: (data) => {
        if (data.section === 'body' && data.column.index === 4) {
          const val = String(data.cell.raw);
          if (val === 'PASS') {
            data.cell.styles.textColor = [16, 185, 129];
          } else if (val === 'FAIL') {
            data.cell.styles.textColor = [225, 29, 72];
          } else {
            data.cell.styles.textColor = [217, 119, 6];
          }
        }
      }
    });

    const finalY = (doc as any).lastAutoTable.finalY + 6;

    // Recommendations Box
    doc.setFillColor(241, 245, 249);
    doc.rect(14, finalY, 182, 38, 'F');
    doc.setDrawColor(203, 213, 225);
    doc.rect(14, finalY, 182, 38, 'S');

    doc.setTextColor(15, 23, 42);
    doc.setFontSize(8.5);
    doc.setFont('helvetica', 'bold');
    doc.text('KHUYEN NGHI KY THUAT CHO FSE (ACTIONABLE RECOMMENDATIONS):', 18, finalY + 6);

    doc.setFontSize(7.5);
    doc.setFont('helvetica', 'normal');
    let recY = finalY + 12;
    result.actionable_recommendations.steps.slice(0, 3).forEach((step) => {
      doc.setFont('helvetica', 'bold');
      doc.text(`${step.step_number}. ${step.title_vn.slice(0, 75)}`, 18, recY);
      recY += 4.5;
      doc.setFont('helvetica', 'normal');
      doc.text(step.details_vn.slice(0, 110), 22, recY);
      recY += 4.5;
    });

    // Signature blocks
    const sigY = finalY + 44;
    doc.setFontSize(8);
    doc.setFont('helvetica', 'bold');
    doc.text('KY SU CHAY THU (FSE)', 26, sigY);
    doc.text('DON VI GIAM SAT (OWNER/PM)', 88, sigY);
    doc.text('DON VI VAN HANH / O&M', 152, sigY);

    doc.setFont('helvetica', 'italic');
    doc.setFontSize(7);
    doc.setTextColor(100, 116, 139);
    doc.text('(Ky & ghi ro ho ten)', 28, sigY + 14);
    doc.text('(Ky & ghi ro ho ten)', 92, sigY + 14);
    doc.text('(Ky & ghi ro ho ten)', 155, sigY + 14);

    doc.save(`TEV_COE_NETA_ECS_8_3_ATS_${systemTag.replace(/[^a-zA-Z0-9]/g, '_')}_${commissioningDate}.pdf`);
  };

  const copyToClipboard = (text: string, type: 'prompt' | 'json') => {
    navigator.clipboard.writeText(text);
    if (type === 'prompt') {
      setCopiedPrompt(true);
      setTimeout(() => setCopiedPrompt(false), 2000);
    } else {
      setCopiedJson(true);
      setTimeout(() => setCopiedJson(false), 2000);
    }
  };

  return (
    <div className="space-y-6">
      {/* 1. TOP HEADER BANNER */}
      <div className="bg-gradient-to-r from-slate-900 via-blue-950 to-slate-900 rounded-2xl p-5 md:p-6 text-white shadow-xl border border-blue-500/30">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="space-y-1.5">
            <div className="flex items-center gap-2.5 flex-wrap">
              <div className="w-10 h-10 rounded-xl bg-blue-500/20 text-blue-400 border border-blue-400/30 flex items-center justify-center font-black">
                <GitBranch size={24} />
              </div>
              <div>
                <div className="flex items-center gap-2 flex-wrap">
                  <h2 className="text-xl font-black tracking-tight">
                    {isVi
                      ? 'Nghiệm Thu & Chạy Thử Tủ Chuyển Nguồn Tự Động (ATS)'
                      : 'Source Specific Systems Commissioning: Automatic Transfer Switch (ATS)'}
                  </h2>
                  <span className="px-2 py-0.5 rounded-full text-[10px] font-black uppercase tracking-wider bg-blue-500 text-white shadow-sm">
                    NETA ECS-2024 Mục 8.3
                  </span>
                </div>
                <p className="text-xs text-blue-200/80">
                  {isVi
                    ? 'Chuẩn hóa theo ANSI/NETA ECS-2024 Mục 8.3, ANSI/NETA ATS-2025 Mục 7.22.3, NFPA 110 & IEEE 446'
                    : 'Standardized per ANSI/NETA ECS-2024 Sec 8.3, NETA ATS-2025 Sec 7.22.3, NFPA 110 & IEEE 446'}
                </p>
              </div>
            </div>
          </div>

          {/* Quick Actions & AI Studio Buttons */}
          <div className="flex items-center gap-2 flex-wrap">
            <button
              type="button"
              onClick={() => setShowSystemPromptModal(true)}
              className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold bg-blue-500/20 text-blue-300 hover:bg-blue-500/30 border border-blue-400/40 transition-all shadow-xs"
              title="Sao chép System Instructions cho Google AI Studio"
            >
              <Copy size={14} />
              <span>Prompt AI Studio</span>
            </button>

            <button
              type="button"
              onClick={() => setShowJsonPromptModal(true)}
              className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold bg-slate-800 text-slate-200 hover:bg-slate-700 border border-slate-700 transition-all shadow-xs"
              title="Xem mẫu JSON Payload"
            >
              <FileCheck2 size={14} />
              <span>JSON Input</span>
            </button>

            <button
              type="button"
              onClick={handleExportPdf}
              className="flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-xs font-black bg-gradient-to-r from-blue-500 to-indigo-600 text-white hover:from-blue-400 hover:to-indigo-500 transition-all shadow-md shadow-blue-500/20"
            >
              <Download size={14} />
              <span>{isVi ? 'Xuất PDF A4' : 'Export PDF'}</span>
            </button>
          </div>
        </div>

        {/* 2. PRESET SCENARIO SELECTOR */}
        <div className="mt-5 pt-4 border-t border-blue-500/20">
          <div className="text-xs font-bold text-blue-300 mb-2 flex items-center gap-1.5">
            <Sparkles size={14} />
            <span>{isVi ? 'Kịch bản thực địa mẫu (1-Click Presets):' : 'Sample Field Scenarios:'}</span>
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2">
            {PRESET_SCENARIOS.map((scenario) => (
              <button
                key={scenario.id}
                type="button"
                onClick={() => applyPresetScenario(scenario)}
                className="text-left p-2.5 rounded-xl bg-slate-950/70 hover:bg-slate-900 border border-slate-800 hover:border-blue-400/50 transition-all group"
              >
                <div className="flex items-center justify-between mb-1">
                  <span className={`text-[9px] font-black px-1.5 py-0.5 rounded border ${scenario.badgeColor}`}>
                    {scenario.badge}
                  </span>
                  <span className="text-[10px] font-mono text-slate-400">{scenario.currentRatingA}A 4-Pole</span>
                </div>
                <div className="text-xs font-bold text-slate-200 group-hover:text-blue-300 line-clamp-1">
                  {isVi ? scenario.name : scenario.nameEn}
                </div>
                <div className="text-[10px] text-slate-400 line-clamp-2 mt-0.5">
                  {scenario.description}
                </div>
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* 3. INPUT FORMS (3 GIAI ĐOẠN ĐÓNG ĐIỆN & CHUYỂN NGUỒN) */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* PHASE 1: PRE-ENERGIZATION (NETA ECS 8.3.A) */}
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 space-y-4">
          <div className="flex items-center justify-between border-b border-slate-100 pb-3">
            <div className="flex items-center gap-2">
              <div className="w-7 h-7 rounded-lg bg-blue-50 text-blue-700 flex items-center justify-center font-bold text-xs">
                A
              </div>
              <h3 className="font-black text-sm text-slate-900">
                {isVi ? 'GIAI ĐOẠN 1: TIỀN CHẠY THỬ' : 'PHASE 1: PRE-ENERGIZATION'}
              </h3>
            </div>
            <span className="text-[10px] font-mono font-bold px-2 py-0.5 rounded bg-slate-100 text-slate-600">
              NETA ECS 8.3.A
            </span>
          </div>

          {/* Voltage & Current Rating */}
          <div className="grid grid-cols-2 gap-2">
            <div>
              <label className="text-[11px] font-semibold text-slate-600 block mb-1">
                {isVi ? 'Điện áp định mức:' : 'Rated Voltage:'}
              </label>
              <input
                type="number"
                value={voltageRatingV}
                onChange={(e) => setVoltageRatingV(parseFloat(e.target.value) || 0)}
                className="w-full text-xs font-bold px-2 py-1.5 bg-slate-50 border border-slate-300 rounded-lg text-slate-800"
              />
            </div>
            <div>
              <label className="text-[11px] font-semibold text-slate-600 block mb-1">
                {isVi ? 'Dòng điện định mức (A):' : 'Rated Current (A):'}
              </label>
              <input
                type="number"
                value={currentRatingA}
                onChange={(e) => setCurrentRatingA(parseFloat(e.target.value) || 0)}
                className="w-full text-xs font-bold px-2 py-1.5 bg-slate-50 border border-slate-300 rounded-lg text-slate-800"
              />
            </div>
          </div>

          {/* Control Wiring Insulation (Table 100.1) */}
          <div className="space-y-1 bg-slate-50 p-2.5 rounded-xl border border-slate-200">
            <div className="flex items-center justify-between text-xs font-bold text-slate-800">
              <span>{isVi ? 'Cách điện mạch điều khiển (1000V DC):' : 'Control Wiring IR (1000V DC):'}</span>
              <span className={`text-[10px] font-black px-1.5 rounded ${
                controlIr >= 1.0 ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'
              }`}>
                Min ≥ 1.0 MΩ
              </span>
            </div>
            <div className="flex items-center gap-2">
              <input
                type="number"
                value={controlIr}
                onChange={(e) => setControlIr(parseFloat(e.target.value) || 0)}
                className="w-full text-xs font-black px-2 py-1 bg-white border border-slate-300 rounded-lg text-slate-800"
              />
              <span className="text-xs font-bold text-slate-600">MΩ</span>
            </div>
          </div>

          {/* Contact Resistance Source 1 vs Source 2 */}
          <div className="space-y-2 bg-slate-50 p-2.5 rounded-xl border border-slate-200">
            <div className="flex items-center justify-between text-xs font-bold text-slate-800">
              <span>{isVi ? 'Điện trở tiếp điểm cực (µΩ):' : 'Contact Resistance (µΩ):'}</span>
              <span className={`text-[10px] font-black px-1.5 rounded ${
                contactDeviationStats.isPass ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800 animate-pulse'
              }`}>
                Δ Lệch: {contactDeviationStats.maxDev}% (Max ≤ 50%)
              </span>
            </div>

            {/* Source 1 (Normal) */}
            <div className="text-[11px] font-semibold text-slate-700">
              Nguồn 1 (Ưu tiên/Lưới): A, B, C (µΩ)
            </div>
            <div className="grid grid-cols-3 gap-1.5">
              <input
                type="number"
                placeholder="A"
                value={s1PhaseA}
                onChange={(e) => setS1PhaseA(parseFloat(e.target.value) || 0)}
                className="text-xs font-bold px-1.5 py-1 text-center bg-white border border-slate-300 rounded"
              />
              <input
                type="number"
                placeholder="B"
                value={s1PhaseB}
                onChange={(e) => setS1PhaseB(parseFloat(e.target.value) || 0)}
                className="text-xs font-bold px-1.5 py-1 text-center bg-white border border-slate-300 rounded"
              />
              <input
                type="number"
                placeholder="C"
                value={s1PhaseC}
                onChange={(e) => setS1PhaseC(parseFloat(e.target.value) || 0)}
                className="text-xs font-bold px-1.5 py-1 text-center bg-white border border-slate-300 rounded"
              />
            </div>

            {/* Source 2 (Emergency) */}
            <div className="text-[11px] font-semibold text-slate-700 pt-1">
              Nguồn 2 (Dự phòng/Máy phát): A, B, C (µΩ)
            </div>
            <div className="grid grid-cols-3 gap-1.5">
              <input
                type="number"
                placeholder="A"
                value={s2PhaseA}
                onChange={(e) => setS2PhaseA(parseFloat(e.target.value) || 0)}
                className="text-xs font-bold px-1.5 py-1 text-center bg-white border border-slate-300 rounded"
              />
              <input
                type="number"
                placeholder="B"
                value={s2PhaseB}
                onChange={(e) => setS2PhaseB(parseFloat(e.target.value) || 0)}
                className={`text-xs font-black px-1.5 py-1 text-center bg-white border rounded ${
                  s2PhaseB > 50 ? 'border-rose-400 text-rose-600 bg-rose-50/50' : 'border-slate-300 text-slate-800'
                }`}
              />
              <input
                type="number"
                placeholder="C"
                value={s2PhaseC}
                onChange={(e) => setS2PhaseC(parseFloat(e.target.value) || 0)}
                className="text-xs font-bold px-1.5 py-1 text-center bg-white border border-slate-300 rounded"
              />
            </div>
          </div>

          {/* Mechanical & Electrical Interlocks */}
          <div className="space-y-1">
            <label className="text-xs font-bold text-slate-700">
              {isVi ? 'Khóa liên động cơ & điện (NFPA 110):' : 'Mechanical & Electrical Interlocks:'}
            </label>
            <button
              type="button"
              onClick={() => setIsInterlocksPass(!isInterlocksPass)}
              className={`w-full py-1.5 px-2 text-xs font-bold rounded-lg border flex items-center justify-center gap-1.5 ${
                isInterlocksPass
                  ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                  : 'bg-rose-50 text-rose-800 border-rose-300 font-black'
              }`}
            >
              {isInterlocksPass ? <CheckCircle2 size={13} /> : <AlertOctagon size={13} />}
              <span>{isInterlocksPass ? (isVi ? 'Liên động tin cậy 100%' : 'Interlocks Verified') : (isVi ? 'LỖI LIÊN ĐỘNG CƠ ĐIỆN' : 'INTERLOCK FAILURE')}</span>
            </button>
          </div>

          {/* Bolt Torque & Plan */}
          <div className="grid grid-cols-2 gap-2 pt-1">
            <div>
              <label className="text-[11px] font-semibold text-slate-600 block mb-0.5">{isVi ? 'Lực siết bu-lông:' : 'Bolt Torque:'}</label>
              <button
                type="button"
                onClick={() => setIsTorquePass(!isTorquePass)}
                className={`w-full py-1 text-xs font-bold rounded-lg border ${
                  isTorquePass ? 'bg-emerald-50 text-emerald-800 border-emerald-300' : 'bg-rose-50 text-rose-800 border-rose-300'
                }`}
              >
                {isTorquePass ? (isVi ? 'Đạt Bảng 100.12' : 'Table 100.12 OK') : (isVi ? 'Chưa đủ lực' : 'Deficient')}
              </button>
            </div>
            <div>
              <label className="text-[11px] font-semibold text-slate-600 block mb-0.5">{isVi ? 'Kế hoạch đóng điện:' : 'Written Plan:'}</label>
              <button
                type="button"
                onClick={() => setIsPlanApproved(!isPlanApproved)}
                className={`w-full py-1 text-xs font-bold rounded-lg border ${
                  isPlanApproved ? 'bg-emerald-50 text-emerald-800 border-emerald-300' : 'bg-rose-50 text-rose-800 border-rose-300'
                }`}
              >
                {isPlanApproved ? (isVi ? 'Đã ký duyệt' : 'Approved') : (isVi ? 'Chưa duyệt' : 'Missing')}
              </button>
            </div>
          </div>
        </div>

        {/* PHASE 2: ENERGIZATION & TRANSFER TESTING (NETA ECS 8.3.B) */}
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 space-y-4">
          <div className="flex items-center justify-between border-b border-slate-100 pb-3">
            <div className="flex items-center gap-2">
              <div className="w-7 h-7 rounded-lg bg-teal-50 text-teal-700 flex items-center justify-center font-bold text-xs">
                B
              </div>
              <h3 className="font-black text-sm text-slate-900">
                {isVi ? 'GIAI ĐOẠN 2: THỬ CHUYỂN NGUỒN' : 'PHASE 2: LOAD TRANSFER'}
              </h3>
            </div>
            <span className="text-[10px] font-mono font-bold px-2 py-0.5 rounded bg-slate-100 text-slate-600">
              NETA ECS 8.3.B
            </span>
          </div>

          {/* Phase Sequence Matching */}
          <div className="space-y-1">
            <label className="text-xs font-bold text-slate-700">
              {isVi ? 'B2. Trùng thứ tự pha hai nguồn (A-B-C):' : 'B2. Phase Sequence Matching:'}
            </label>
            <button
              type="button"
              onClick={() => setIsPhasingPass(!isPhasingPass)}
              className={`w-full py-2 px-3 text-xs font-bold rounded-xl border flex items-center justify-between ${
                isPhasingPass
                  ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                  : 'bg-rose-50 text-rose-800 border-rose-300 font-black'
              }`}
            >
              <span>{isPhasingPass ? (isVi ? 'Thứ tự pha trùng khớp 100%' : 'Phase Sequence Matches') : (isVi ? 'NGƯỢC PHA HAI NGUỒN' : 'PHASE ROTATION MISMATCH')}</span>
              {isPhasingPass ? <CheckCircle2 size={15} /> : <AlertOctagon size={15} />}
            </button>
          </div>

          {/* Automatic Load Transfer Timing (NFPA 110 Class 10 <= 10s) */}
          <div className="space-y-2 bg-slate-50 p-3 rounded-xl border border-slate-200">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <Clock size={14} className="text-blue-600" />
                <span>{isVi ? 'B3. Thời gian chuyển sang máy phát:' : 'B3. Emergency Transfer Time:'}</span>
              </span>
              <span className={`text-[10px] font-black px-1.5 py-0.5 rounded ${
                totalTransferTime <= 10.0 ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800 animate-pulse'
              }`}>
                {totalTransferTime <= 10.0 ? 'ĐẠT (≤ 10.0s NFPA 110)' : 'FAIL (> 10.0s)'}
              </span>
            </div>
            <div>
              <div className="flex items-center justify-between text-[11px] text-slate-600 mb-1">
                <span>{isVi ? 'Tổng thời gian mất điện đến khi cấp lại:' : 'Total Outage-to-Transfer Time:'}</span>
                <span className="font-mono font-bold text-slate-900">{totalTransferTime} giây (s)</span>
              </div>
              <input
                type="range"
                min="3.0"
                max="25.0"
                step="0.2"
                value={totalTransferTime}
                onChange={(e) => {
                  const val = parseFloat(e.target.value);
                  setTotalTransferTime(val);
                  setAtsTransferCompleteTime(val);
                }}
                className="w-full accent-blue-600"
              />
            </div>
          </div>

          {/* Retransfer & Engine Cool-down Test */}
          <div className="space-y-1.5 bg-slate-50 p-2.5 rounded-xl border border-slate-200">
            <div className="text-xs font-bold text-slate-800">
              {isVi ? 'B4. Trả nguồn lưới & Làm mát máy phát:' : 'B4. Retransfer & Cool-down Sequence:'}
            </div>
            <div className="grid grid-cols-2 gap-2 text-[11px]">
              <div className="bg-white p-2 rounded border border-slate-200">
                <span className="text-slate-500 block">Trễ trả nguồn:</span>
                <span className="font-bold text-slate-800">{retransferDelaySec} giây (5 phút)</span>
              </div>
              <div className="bg-white p-2 rounded border border-slate-200">
                <span className="text-slate-500 block">Chạy làm mát:</span>
                <span className="font-bold text-slate-800">{coolDownDelaySec} giây (4 phút)</span>
              </div>
            </div>
            <button
              type="button"
              onClick={() => {
                setIsRetransferSuccess(!isRetransferSuccess);
                setIsGenShutdownSuccess(!isGenShutdownSuccess);
              }}
              className={`w-full py-1 text-xs font-bold rounded border ${
                isRetransferSuccess && isGenShutdownSuccess
                  ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                  : 'bg-rose-50 text-rose-800 border-rose-300'
              }`}
            >
              {isRetransferSuccess ? (isVi ? 'Trả nguồn êm & Dừng Gen tự động' : 'Retransfer & Auto Shutdown OK') : (isVi ? 'Lỗi quy trình trả nguồn/làm mát' : 'Retransfer failure')}
            </button>
          </div>

          {/* Single-Phase Failure Test (IEEE 446) */}
          <div className="space-y-1">
            <label className="text-xs font-bold text-slate-700">
              {isVi ? 'B5. Thử nghiệm mất 1 pha (IEEE 446):' : 'B5. Single-Phase Loss Test (IEEE 446):'}
            </label>
            <button
              type="button"
              onClick={() => setIsSinglePhasePass(!isSinglePhasePass)}
              className={`w-full py-2 px-3 text-xs font-bold rounded-xl border flex items-center justify-between ${
                isSinglePhasePass
                  ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                  : 'bg-rose-50 text-rose-800 border-rose-300 font-black'
              }`}
            >
              <span>{isSinglePhasePass ? (isVi ? 'Phát hiện mất Pha A và kích hoạt ATS' : 'Phase Loss Detected') : (isVi ? 'BỎ SÓT LỖI MẤT PHA' : 'PHASE LOSS MISSED')}</span>
              {isSinglePhasePass ? <CheckCircle2 size={15} /> : <AlertOctagon size={15} />}
            </button>
          </div>
        </div>

        {/* PHASE 3: POST-ENERGIZATION (NETA ECS 8.3.C) */}
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 space-y-4">
          <div className="flex items-center justify-between border-b border-slate-100 pb-3">
            <div className="flex items-center gap-2">
              <div className="w-7 h-7 rounded-lg bg-indigo-50 text-indigo-700 flex items-center justify-center font-bold text-xs">
                C
              </div>
              <h3 className="font-black text-sm text-slate-900">
                {isVi ? 'GIAI ĐOẠN 3: SAU ĐÓNG ĐIỆN' : 'PHASE 3: POST-ENERGIZATION'}
              </h3>
            </div>
            <span className="text-[10px] font-mono font-bold px-2 py-0.5 rounded bg-slate-100 text-slate-600">
              NETA ECS 8.3.C
            </span>
          </div>

          {/* C1: Thermographic Survey Table 100.18 */}
          <div className="space-y-2 bg-slate-50 p-3 rounded-xl border border-slate-200">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <Flame size={14} className="text-rose-500" />
                <span>{isVi ? 'C1. Chụp nhiệt IR tiếp điểm (Table 100.18):' : 'C1. IR Survey (Table 100.18):'}</span>
              </span>
              <button
                type="button"
                onClick={() => setIsThermalPass(!isThermalPass)}
                className={`text-[10px] font-bold px-2 py-0.5 rounded ${
                  isThermalPass ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'
                }`}
              >
                {isThermalPass ? (isVi ? 'Nhiệt độ đều' : 'Normal') : (isVi ? 'CẢNH BÁO HOTSPOT' : 'HOTSPOT DETECTED')}
              </button>
            </div>
            <div>
              <div className="flex items-center justify-between text-[11px] text-slate-600 mb-1">
                <span>{isVi ? 'Độ chênh nhiệt cực đại (ΔT):' : 'Max Delta T:'}</span>
                <span className="font-mono font-bold text-slate-900">{thermalDeltaT}°C</span>
              </div>
              <input
                type="range"
                min="1"
                max="30"
                step="0.5"
                value={thermalDeltaT}
                onChange={(e) => {
                  const val = parseFloat(e.target.value);
                  setThermalDeltaT(val);
                  setIsThermalPass(val <= 10.0);
                }}
                className="w-full accent-rose-500"
              />
            </div>
          </div>

          {/* C2: SCADA/BMS Telemetry */}
          <div className="space-y-1">
            <label className="text-xs font-bold text-slate-700">
              {isVi ? 'C2. Truyền thông giám sát SCADA / BMS:' : 'C2. SCADA / BMS Telemetry:'}
            </label>
            <button
              type="button"
              onClick={() => setIsScadaPass(!isScadaPass)}
              className={`w-full py-2 px-3 text-xs font-bold rounded-xl border flex items-center justify-between ${
                isScadaPass
                  ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                  : 'bg-amber-50 text-amber-800 border-amber-300'
              }`}
            >
              <span>{isScadaPass ? (isVi ? 'Tín hiệu Status/Alarm truyền chuẩn' : 'BMS Logging Verified') : (isVi ? 'Lỗi cổng truyền thông Modbus/BACnet' : 'BMS communication error')}</span>
              <Activity size={15} />
            </button>
          </div>

          {/* C3: Owner Manual Bypass Training */}
          <div className="space-y-1">
            <label className="text-xs font-bold text-slate-700">
              {isVi ? 'C3. Đào tạo thao tác tay & Bàn giao:' : 'C3. Manual Bypass Training & Handover:'}
            </label>
            <button
              type="button"
              onClick={() => setIsTrainingCompleted(!isTrainingCompleted)}
              className={`w-full py-2 px-3 text-xs font-bold rounded-xl border flex items-center justify-between ${
                isTrainingCompleted
                  ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                  : 'bg-slate-50 text-slate-700 border-slate-300'
              }`}
            >
              <span>{isTrainingCompleted ? (isVi ? 'Đã đào tạo cán bộ vận hành tay' : 'Training Verified') : (isVi ? 'Chưa đào tạo thao tác khẩn cấp' : 'Training Pending')}</span>
              <CheckCircle2 size={15} />
            </button>
          </div>

          {/* Action Trigger Buttons */}
          <div className="pt-2 space-y-2">
            <button
              type="button"
              onClick={handleRunDeterministic}
              disabled={isAnalyzing}
              className="w-full py-2.5 px-4 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-black transition-all flex items-center justify-center gap-2 shadow-sm"
            >
              <Zap size={14} className="text-blue-400" />
              <span>{isVi ? 'Chẩn Đoán NETA ECS (Tất Định)' : 'Run Deterministic Audit'}</span>
            </button>

            <button
              type="button"
              onClick={handleRunAiAnalysis}
              disabled={isAnalyzing}
              className="w-full py-2.5 px-4 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white rounded-xl text-xs font-black transition-all flex items-center justify-center gap-2 shadow-md shadow-blue-500/20"
            >
              <Sparkles size={14} className="text-amber-300" />
              <span>{isAnalyzing ? (isVi ? 'AI Đang Thẩm Định...' : 'AI Evaluating...') : (isVi ? 'Phân Tích Gemini AI Chuyên Sâu' : 'Run Gemini AI Specialist')}</span>
            </button>
          </div>
        </div>
      </div>

      {/* 4. RESULTS DISPLAY SECTION */}
      {evaluationResult && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-lg overflow-hidden space-y-0">
          {/* Severity Banner */}
          <div className={`p-4 sm:p-5 flex flex-col md:flex-row md:items-center justify-between gap-3 border-b ${
            evaluationResult.overall_status === 'PASS'
              ? 'bg-gradient-to-r from-emerald-950 via-slate-900 to-emerald-950 text-white border-emerald-500/40'
              : evaluationResult.overall_status === 'CAUTION'
              ? 'bg-gradient-to-r from-amber-950 via-slate-900 to-amber-950 text-white border-amber-500/40'
              : 'bg-gradient-to-r from-rose-950 via-slate-900 to-rose-950 text-white border-rose-500/40'
          }`}>
            <div className="flex items-center gap-3">
              <div className={`w-12 h-12 rounded-2xl flex items-center justify-center font-black ${
                evaluationResult.overall_status === 'PASS'
                  ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-400/40'
                  : evaluationResult.overall_status === 'CAUTION'
                  ? 'bg-amber-500/20 text-amber-400 border border-amber-400/40'
                  : 'bg-rose-500/20 text-rose-400 border border-rose-400/40 animate-pulse'
              }`}>
                {evaluationResult.overall_status === 'PASS' ? <ShieldCheck size={28} /> : <AlertOctagon size={28} />}
              </div>
              <div>
                <div className="flex items-center gap-2 flex-wrap">
                  <span className={`text-xs font-black px-2 py-0.5 rounded-full border ${evaluationResult.summary_color}`}>
                    {evaluationResult.summary_badge}
                  </span>
                  <span className="text-xs font-mono text-slate-300">{systemTag}</span>
                </div>
                <h3 className="text-base sm:text-lg font-black tracking-tight mt-0.5">
                  {isVi ? evaluationResult.status_title_vn : evaluationResult.status_title_en}
                </h3>
              </div>
            </div>

            {/* Quick action buttons in banner */}
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={handleExportPdf}
                className="px-3.5 py-2 rounded-xl bg-white/10 hover:bg-white/20 text-white text-xs font-bold transition-all flex items-center gap-1.5"
              >
                <Download size={14} />
                <span>PDF</span>
              </button>
            </div>
          </div>

          {/* Results Navigation Tabs */}
          <div className="bg-slate-100 px-4 py-2 border-b border-slate-200 flex items-center gap-2 overflow-x-auto">
            <button
              type="button"
              onClick={() => setActiveResultTab('summary')}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                activeResultTab === 'summary' ? 'bg-white text-slate-900 shadow-xs' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              {isVi ? 'Tổng quan Chẩn đoán' : 'Diagnostic Summary'}
            </button>
            <button
              type="button"
              onClick={() => setActiveResultTab('table')}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                activeResultTab === 'table' ? 'bg-white text-slate-900 shadow-xs' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              {isVi ? 'Bảng Đối Soát 17 Hạng Mục (NETA ECS 8.3)' : 'Evaluation Table (17 Items)'}
            </button>
            <button
              type="button"
              onClick={() => setActiveResultTab('recommendations')}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                activeResultTab === 'recommendations' ? 'bg-white text-slate-900 shadow-xs' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              {isVi ? 'Khuyến nghị Hành động FSE' : 'Actionable Steps'}
            </button>
            {evaluationResult.raw_markdown_report && (
              <button
                type="button"
                onClick={() => setActiveResultTab('ai-markdown')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 ${
                  activeResultTab === 'ai-markdown' ? 'bg-blue-600 text-white font-black shadow-xs' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Sparkles size={12} className="text-amber-300" />
                <span>{isVi ? 'Báo Cáo Gemini AI' : 'Gemini AI Report'}</span>
              </button>
            )}
          </div>

          {/* TAB 1: SUMMARY & HAZARDS */}
          {activeResultTab === 'summary' && (
            <div className="p-5 space-y-4">
              <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
                <h4 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                  <Activity size={16} className="text-blue-600" />
                  <span>{isVi ? evaluationResult.technical_analysis.title_vn : evaluationResult.technical_analysis.title_en}</span>
                </h4>
                <p className="text-xs text-slate-700 leading-relaxed">
                  {isVi ? evaluationResult.technical_analysis.description_vn : evaluationResult.technical_analysis.description_en}
                </p>
              </div>

              {/* Key KPI Chips */}
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                <div className="p-3 rounded-xl bg-slate-50 border border-slate-200">
                  <div className="text-[10px] text-slate-500 font-bold uppercase">Cách điện Mạch Điều khiển</div>
                  <div className={`text-base font-black ${
                    evaluationResult.technical_analysis.control_ir_megohms < 1.0 ? 'text-rose-600' : 'text-emerald-600'
                  }`}>
                    {evaluationResult.technical_analysis.control_ir_megohms} MΩ
                  </div>
                  <div className="text-[9px] text-slate-400">Chuẩn ≥ 1.0 MΩ</div>
                </div>

                <div className="p-3 rounded-xl bg-slate-50 border border-slate-200">
                  <div className="text-[10px] text-slate-500 font-bold uppercase">Độ lệch Tiếp điểm (Max)</div>
                  <div className={`text-base font-black ${
                    evaluationResult.technical_analysis.contact_deviation_percent > 50.0 ? 'text-rose-600' : 'text-emerald-600'
                  }`}>
                    {evaluationResult.technical_analysis.contact_deviation_percent}%
                  </div>
                  <div className="text-[9px] text-slate-400">Giới hạn ≤ 50%</div>
                </div>

                <div className="p-3 rounded-xl bg-slate-50 border border-slate-200">
                  <div className="text-[10px] text-slate-500 font-bold uppercase">Tiếp điểm Max Nguồn 2</div>
                  <div className={`text-base font-black ${
                    evaluationResult.technical_analysis.source_2_max_contact_uohm > 50 ? 'text-rose-600' : 'text-emerald-600'
                  }`}>
                    {evaluationResult.technical_analysis.source_2_max_contact_uohm} µΩ
                  </div>
                  <div className="text-[9px] text-slate-400">Pha B Nguồn 2</div>
                </div>

                <div className="p-3 rounded-xl bg-slate-50 border border-slate-200">
                  <div className="text-[10px] text-slate-500 font-bold uppercase">Thời gian Chuyển nguồn</div>
                  <div className={`text-base font-black ${
                    evaluationResult.technical_analysis.total_transfer_time_sec > 10.0 ? 'text-rose-600' : 'text-emerald-600'
                  }`}>
                    {evaluationResult.technical_analysis.total_transfer_time_sec}s
                  </div>
                  <div className="text-[9px] text-slate-400">NFPA 110 ≤ 10.0s</div>
                </div>
              </div>
            </div>
          )}

          {/* TAB 2: DETAILED EVALUATION TABLE */}
          {activeResultTab === 'table' && (
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-900 text-white">
                    <th className="py-2.5 px-3 font-bold w-12 text-center">Mục</th>
                    <th className="py-2.5 px-3 font-bold">Hạng Mục Kiểm Tra (NETA ECS 8.3)</th>
                    <th className="py-2.5 px-3 font-bold">Tiêu Chuẩn & Bảng Tham Chiếu</th>
                    <th className="py-2.5 px-3 font-bold">Giá Trị Đo Thực Tế</th>
                    <th className="py-2.5 px-3 font-bold text-center w-20">Kết Quả</th>
                    <th className="py-2.5 px-3 font-bold">Đánh Giá Kỹ Thuật</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {evaluationResult.evaluation_table.map((row) => (
                    <tr key={row.item_code} className="hover:bg-slate-50 transition-colors">
                      <td className="py-2 px-3 font-bold text-center text-slate-500">{row.item_code}</td>
                      <td className="py-2 px-3 font-semibold text-slate-900">
                        {isVi ? row.item_name_vn : row.item_name_en}
                      </td>
                      <td className="py-2 px-3 font-mono text-[11px] text-slate-600">{row.neta_standard_ref}</td>
                      <td className="py-2 px-3 font-medium text-slate-800">{row.measured_value}</td>
                      <td className="py-2 px-3 text-center">
                        <span className={`inline-block px-2 py-0.5 rounded font-black text-[10px] ${
                          row.status === 'PASS'
                            ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                            : row.status === 'FAIL'
                            ? 'bg-rose-100 text-rose-800 border border-rose-300'
                            : 'bg-amber-100 text-amber-800 border border-amber-300'
                        }`}>
                          {row.status}
                        </span>
                      </td>
                      <td className="py-2 px-3 text-slate-700">
                        {isVi ? row.evaluation_vn : row.evaluation_en}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* TAB 3: ACTIONABLE RECOMMENDATIONS */}
          {activeResultTab === 'recommendations' && (
            <div className="p-5 space-y-4">
              <div className="flex items-center justify-between">
                <h4 className="text-sm font-black text-slate-900">
                  {isVi ? 'Quy trình Khắc phục & Khuyến nghị Thao tác Hiện trường:' : 'FSE Actionable Procedures:'}
                </h4>
                {evaluationResult.actionable_recommendations.halt_transfer && (
                  <span className="px-2.5 py-1 rounded-full text-xs font-black bg-rose-600 text-white animate-pulse">
                    {isVi ? '⚠️ KHÔNG CHUYỂN SANG MÁY PHÁT' : '⚠️ HOLD TRANSFER TO GENERATOR'}
                  </span>
                )}
              </div>

              <div className="space-y-3">
                {evaluationResult.actionable_recommendations.steps.map((step) => (
                  <div key={step.step_number} className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 flex gap-3">
                    <div className="w-7 h-7 rounded-lg bg-blue-600 text-white font-black flex items-center justify-center shrink-0 text-xs shadow-xs">
                      {step.step_number}
                    </div>
                    <div className="space-y-1">
                      <div className="text-xs font-black text-slate-900">
                        {isVi ? step.title_vn : step.title_en}
                      </div>
                      <div className="text-xs text-slate-700">
                        {isVi ? step.details_vn : step.details_en}
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* TAB 4: GEMINI AI MARKDOWN */}
          {activeResultTab === 'ai-markdown' && evaluationResult.raw_markdown_report && (
            <div className="p-5 bg-slate-950 text-slate-200 overflow-x-auto text-xs leading-relaxed font-mono">
              <pre className="whitespace-pre-wrap">{evaluationResult.raw_markdown_report}</pre>
            </div>
          )}
        </div>
      )}

      {/* MODAL 1: SYSTEM INSTRUCTION PROMPT */}
      {showSystemPromptModal && (
        <div className="fixed inset-0 z-50 bg-black/70 flex items-center justify-center p-4">
          <div className="bg-slate-900 text-white rounded-2xl max-w-2xl w-full border border-slate-800 shadow-2xl overflow-hidden flex flex-col max-h-[85vh]">
            <div className="p-4 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Sparkles size={18} className="text-amber-400" />
                <h3 className="font-bold text-sm">System Instruction Prompt (Google AI Studio - ATS ECS 8.3)</h3>
              </div>
              <button
                type="button"
                onClick={() => setShowSystemPromptModal(false)}
                className="text-slate-400 hover:text-white text-sm"
              >
                ✕
              </button>
            </div>
            <div className="p-4 overflow-y-auto font-mono text-xs text-slate-300 whitespace-pre-wrap bg-slate-950 flex-1">
              {ECS_ATS_SYSTEM_INSTRUCTION}
            </div>
            <div className="p-3 bg-slate-900 border-t border-slate-800 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => copyToClipboard(ECS_ATS_SYSTEM_INSTRUCTION, 'prompt')}
                className="px-4 py-2 rounded-xl text-xs font-black bg-blue-600 text-white hover:bg-blue-500 flex items-center gap-1.5"
              >
                {copiedPrompt ? <Check size={14} /> : <Copy size={14} />}
                <span>{copiedPrompt ? 'Đã sao chép!' : 'Sao chép Prompt'}</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 2: JSON INPUT PROMPT */}
      {showJsonPromptModal && (
        <div className="fixed inset-0 z-50 bg-black/70 flex items-center justify-center p-4">
          <div className="bg-slate-900 text-white rounded-2xl max-w-2xl w-full border border-slate-800 shadow-2xl overflow-hidden flex flex-col max-h-[85vh]">
            <div className="p-4 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <FileCheck2 size={18} className="text-amber-400" />
                <h3 className="font-bold text-sm">JSON User Prompt Payload (ATS ECS 8.3)</h3>
              </div>
              <button
                type="button"
                onClick={() => setShowJsonPromptModal(false)}
                className="text-slate-400 hover:text-white text-sm"
              >
                ✕
              </button>
            </div>
            <div className="p-4 overflow-y-auto font-mono text-xs text-slate-300 whitespace-pre-wrap bg-slate-950 flex-1">
              {JSON.stringify(currentPayload, null, 2)}
            </div>
            <div className="p-3 bg-slate-900 border-t border-slate-800 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => copyToClipboard(JSON.stringify(currentPayload, null, 2), 'json')}
                className="px-4 py-2 rounded-xl text-xs font-black bg-blue-600 text-white hover:bg-blue-500 flex items-center gap-1.5"
              >
                {copiedJson ? <Check size={14} /> : <Copy size={14} />}
                <span>{copiedJson ? 'Đã sao chép!' : 'Sao chép JSON'}</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
