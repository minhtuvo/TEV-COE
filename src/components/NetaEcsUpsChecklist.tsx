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
  BatteryCharging,
  Radio,
  Clock,
  Sliders,
  Cable,
  Power,
  Cpu,
  TrendingUp
} from 'lucide-react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  NetaEcsUpsInputPayload,
  NetaEcsUpsEvaluationResult,
  performDeterministicEcsUpsAnalysis,
  ECS_UPS_SYSTEM_INSTRUCTION
} from '../../server/netaEcsUpsAnalyzer';
import { UpsBatteryResistanceTrendChart } from './UpsBatteryResistanceTrendChart';

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
  upsRatingKva: number;
  acVoltageRatingV: number;
  dcBusVoltageV: number;
  batteryType: string;
  batteryCapacityAh: number;
  powerIr: number;
  controlIr: number;
  batteryDeviation: number;
  batteryTorquePass: boolean;
  batteryUniformPass: boolean;
  isInterlocksPass: boolean;
  isPhasingPass: boolean;
  dcRipple: number;
  syncPhaseAngle: number;
  syncVoltageDiff: number;
  staticTransferMs: number;
  isLossOfAcPass: boolean;
  fullLoadDischargePass: boolean;
  thdPercent: number;
  isManualBypassPass: boolean;
  thermalDeltaT: number;
  isScadaPass: boolean;
  isTrainingPass: boolean;
  isPlanApproved: boolean;
  description: string;
}

const PRESET_SCENARIOS: PresetScenario[] = [
  {
    id: 'critical-battery-impedance',
    name: '🔴 NGUY HẠI ĐÓNG ĐIỆN: Nội Trở Bình Ắc-quy Tăng Vọt (Lệch +45% IEEE 1188) & Quá Nhiệt Cọc Bình (ΔT = 18.2°C)',
    nameEn: 'CRITICAL HAZARD: Severe Battery Impedance Anomaly (+45% IEEE 1188) & Terminal Hotspot (Delta-T = 18.2°C)',
    badge: 'CRITICAL / NGUY HIỂM',
    badgeColor: 'bg-rose-500/20 text-rose-300 border-rose-500/40',
    systemTag: 'UPS-CRITICAL-DATACENTER-01 (400kVA 3-Phase Online UPS)',
    upsRatingKva: 400,
    acVoltageRatingV: 400,
    dcBusVoltageV: 540,
    batteryType: 'VRLA (12V 150Ah x 40 bình)',
    batteryCapacityAh: 150,
    powerIr: 250.0,
    controlIr: 45.0,
    batteryDeviation: 45.2, // > 20% limit IEEE 1188
    batteryTorquePass: false, // Deficient torque
    batteryUniformPass: false,
    isInterlocksPass: true,
    isPhasingPass: true,
    dcRipple: 0.65,
    syncPhaseAngle: 1.2,
    syncVoltageDiff: 2.1,
    staticTransferMs: 0.0,
    isLossOfAcPass: true,
    fullLoadDischargePass: false, // Premature collapse
    thdPercent: 1.8,
    isManualBypassPass: true,
    thermalDeltaT: 18.2, // > 15°C critical Table 100.18
    isScadaPass: true,
    isTrainingPass: true,
    isPlanApproved: true,
    description: 'Đo nội trở chuỗi ắc-quy phát hiện Cell #18 lệch +45.2% so với baseline và bu-lông cọc bình lỏng gây phát nhiệt ΔT = 18.2°C. Nguy cơ nổ bình và sụt áp toàn chuỗi khi phóng tải.'
  },
  {
    id: 'all-pass-datacenter-ups',
    name: '🟢 ĐẠT CHUẨN 100%: Hệ Thống UPS 3 Pha 400kVA Data Center (Chuyển Mạch Tĩnh 0ms, THD-v 1.6%, Tải Giả 100% Đạt Chuẩn)',
    nameEn: '100% PASS: Mission-Critical 400kVA Data Center UPS (0ms Static Transfer, THD-v 1.6%, Full Load Bank Cleared)',
    badge: 'PASS / ĐẠT CHUẨN',
    badgeColor: 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40',
    systemTag: 'UPS-DATACENTER-TIER3-A',
    upsRatingKva: 400,
    acVoltageRatingV: 400,
    dcBusVoltageV: 540,
    batteryType: 'VRLA High-Rate (12V 150Ah x 40 bình)',
    batteryCapacityAh: 150,
    powerIr: 480.0,
    controlIr: 120.0,
    batteryDeviation: 4.8,
    batteryTorquePass: true,
    batteryUniformPass: true,
    isInterlocksPass: true,
    isPhasingPass: true,
    dcRipple: 0.35,
    syncPhaseAngle: 0.8,
    syncVoltageDiff: 1.4,
    staticTransferMs: 0.0,
    isLossOfAcPass: true,
    fullLoadDischargePass: true,
    thdPercent: 1.6,
    isManualBypassPass: true,
    thermalDeltaT: 1.8,
    isScadaPass: true,
    isTrainingPass: true,
    isPlanApproved: true,
    description: 'Chuyển mạch tĩnh 0ms không gián đoạn, cách điện 480 MΩ, lệch nội trở bình 4.8%, méo sóng hài THD-v 1.6%, xả tải 100% trong 30 phút đạt chuẩn NFPA 111 & NETA ECS.'
  },
  {
    id: 'critical-mbb-interlock-defect',
    name: '🔴 NGUY HẠI: Hỏng Khóa Liên Động Cơ Học Cắt MBB Bảo Trì (Nguy Cơ Chập Pha Khi Thao Tác Chuyển Nguồn)',
    nameEn: 'CRITICAL: Maintenance Bypass (MBB) Mechanical Interlock Defect (Risk of Live Inverter Short-Circuit)',
    badge: 'CRITICAL / NGUY HIỂM',
    badgeColor: 'bg-rose-500/20 text-rose-300 border-rose-500/40',
    systemTag: 'UPS-HOSPITAL-EMERGENCY-02',
    upsRatingKva: 250,
    acVoltageRatingV: 400,
    dcBusVoltageV: 480,
    batteryType: 'VRLA (12V 100Ah x 36 bình)',
    batteryCapacityAh: 100,
    powerIr: 310.0,
    controlIr: 80.0,
    batteryDeviation: 6.5,
    batteryTorquePass: true,
    batteryUniformPass: true,
    isInterlocksPass: false, // Critical failure
    isPhasingPass: true,
    dcRipple: 0.45,
    syncPhaseAngle: 1.0,
    syncVoltageDiff: 2.0,
    staticTransferMs: 0.0,
    isLossOfAcPass: true,
    fullLoadDischargePass: true,
    thdPercent: 2.1,
    isManualBypassPass: false,
    thermalDeltaT: 2.4,
    isScadaPass: true,
    isTrainingPass: false,
    isPlanApproved: false,
    description: 'Cơ cấu chốt khóa liên động cơ khí giữa máy cắt MBB và Inverter bị kẹt, nguy cơ nối tắt nguồn không đồng bộ gây phá hủy tầng công suất IGBT.'
  },
  {
    id: 'caution-harmonics-scada',
    name: '🟡 CẢNH BÁO: Độ Méo Sóng Hài Điện Áp THD-v Cao (4.8%) & Chưa Hoàn Thành Ánh Xạ Cảnh Báo DCIM/BMS',
    nameEn: 'CAUTION: Elevated Voltage Harmonics (THD-v 4.8%) & Incomplete DCIM/BMS Alarm Telemetry',
    badge: 'CAUTION / CẢNH BÁO',
    badgeColor: 'bg-amber-500/20 text-amber-300 border-amber-500/40',
    systemTag: 'UPS-FINANCIAL-TRADING-03',
    upsRatingKva: 160,
    acVoltageRatingV: 400,
    dcBusVoltageV: 480,
    batteryType: 'Lithium Iron Phosphate (LFP Rack 480V 100Ah)',
    batteryCapacityAh: 100,
    powerIr: 190.0,
    controlIr: 35.0,
    batteryDeviation: 8.2,
    batteryTorquePass: true,
    batteryUniformPass: true,
    isInterlocksPass: true,
    isPhasingPass: true,
    dcRipple: 0.85,
    syncPhaseAngle: 2.4,
    syncVoltageDiff: 3.8,
    staticTransferMs: 1.5,
    isLossOfAcPass: true,
    fullLoadDischargePass: true,
    thdPercent: 4.8, // > 3.0% IEEE 519 warning
    isManualBypassPass: true,
    thermalDeltaT: 4.2, // 4-15°C Priority 2 warning
    isScadaPass: false, // Scada issue
    isTrainingPass: true,
    isPlanApproved: true,
    description: 'Độ méo sóng hài THD-v ngõ ra đạt 4.8% (gần ngưỡng trần 5.0% IEEE 519) và một số thanh ghi cảnh báo ngắt DC chưa liên kết về hệ thống BMS trung tâm.'
  }
];

export const NetaEcsUpsChecklist: React.FC<Props> = ({
  onSaveReport,
  onNavigateToReports,
  initialLanguage = 'vi'
}) => {
  const [lang, setLang] = useState<'vi' | 'en'>(initialLanguage);
  const [selectedScenarioId, setSelectedScenarioId] = useState<string>('critical-battery-impedance');
  const [activeTab, setActiveTab] = useState<'checklist' | 'evaluation' | 'trend' | 'prompt' | 'json'>('checklist');
  const [isCopiedPrompt, setIsCopiedPrompt] = useState<boolean>(false);
  const [isCopiedJson, setIsCopiedJson] = useState<boolean>(false);
  const [isAnalyzing, setIsAnalyzing] = useState<boolean>(false);
  const [analysisError, setAnalysisError] = useState<string | null>(null);

  // Form State
  const [projectName, setProjectName] = useState('Dự Án Trung Tâm Dữ Liệu TEV Hyperscale DC-01');
  const [systemTag, setSystemTag] = useState('UPS-CRITICAL-DATACENTER-01 (400kVA 3-Phase Online UPS)');
  const [upsRatingKva, setUpsRatingKva] = useState(400);
  const [acVoltageRatingV, setAcVoltageRatingV] = useState(400);
  const [dcBusVoltageV, setDcBusVoltageV] = useState(540);
  const [batteryType, setBatteryType] = useState('VRLA (12V 150Ah x 40 bình)');
  const [batteryCapacityAh, setBatteryCapacityAh] = useState(150);
  const [fseName, setFseName] = useState('Võ Minh Tú (sgm1707@gmail.com)');
  const [commissioningDate, setCommissioningDate] = useState('2026-10-18');

  // Pre-energization 8.2.A
  const [isDocApproved, setIsDocApproved] = useState(true);
  const [isAtsReviewed, setIsAtsReviewed] = useState(true);
  const [powerIr, setPowerIr] = useState(250.0);
  const [controlIr, setControlIr] = useState(45.0);
  const [batteryDeviation, setBatteryDeviation] = useState(45.2);
  const [isBatteryTorquePass, setIsBatteryTorquePass] = useState(false);
  const [isBatteryUniformPass, setIsBatteryUniformPass] = useState(false);
  const [isInterlocksPass, setIsInterlocksPass] = useState(true);
  const [isGroundingPass, setIsGroundingPass] = useState(true);
  const [isPlanApproved, setIsPlanApproved] = useState(true);

  // Energization 8.2.B
  const [rectifierInputVoltage, setRectifierInputVoltage] = useState(402);
  const [bypassInputVoltage, setBypassInputVoltage] = useState(401);
  const [isPhasingPass, setIsPhasingPass] = useState(true);
  const [dcRipple, setDcRipple] = useState(0.65);
  const [inverterOutputVoltage, setInverterOutputVoltage] = useState(400);
  const [phaseBalanceDeviation, setPhaseBalanceDeviation] = useState(0.5);
  const [syncPhaseAngle, setSyncPhaseAngle] = useState(1.2);
  const [syncVoltageDiff, setSyncVoltageDiff] = useState(2.1);
  const [staticTransferMs, setStaticTransferMs] = useState(0.0);
  const [isLossOfAcPass, setIsLossOfAcPass] = useState(true);
  const [fullLoadDischargePass, setFullLoadDischargePass] = useState(false);
  const [thdPercent, setThdPercent] = useState(1.8);
  const [isManualBypassPass, setIsManualBypassPass] = useState(true);

  // Post-energization 8.2.C
  const [thermalDeltaT, setThermalDeltaT] = useState(18.2);
  const [isThermalPass, setIsThermalPass] = useState(false);
  const [isScadaPass, setIsScadaPass] = useState(true);
  const [isTrainingPass, setIsTrainingPass] = useState(true);

  // Current Payload
  const currentPayload: NetaEcsUpsInputPayload = useMemo(() => {
    return {
      tev_coe_header: {
        module: 'SOURCE SPECIFIC SYSTEMS COMMISSIONING - UNINTERRUPTIBLE POWER SUPPLY (UPS)',
        project_name: projectName,
        system_tag: systemTag,
        ups_rating_kVA: upsRatingKva,
        ac_voltage_rating_V: acVoltageRatingV,
        dc_bus_voltage_V: dcBusVoltageV,
        battery_type: batteryType,
        battery_capacity_Ah: batteryCapacityAh,
        fse_name: fseName,
        commissioning_date: commissioningDate
      },
      pre_energization_8_2_a: {
        documentation_and_drawings_approved: isDocApproved,
        acceptance_test_reports_reviewed_ats_7_22_2: isAtsReviewed,
        power_circuit_insulation_1000v_megohms: powerIr,
        control_wiring_insulation_500v_megohms: controlIr,
        battery_bank_testing: {
          number_of_cells: 40,
          max_cell_impedance_deviation_percent: batteryDeviation,
          battery_float_voltage_V: dcBusVoltageV,
          cell_voltage_uniformity_pass: isBatteryUniformPass,
          bolted_connections_torque_table_100_12: isBatteryTorquePass ? 'Pass' : 'Fail',
          status: isBatteryTorquePass && batteryDeviation <= 20 ? 'Pass' : 'Deficient'
        },
        mechanical_and_electrical_interlocks: isInterlocksPass ? 'Pass' : 'Fail',
        grounding_and_neutral_bonding: isGroundingPass ? 'Pass' : 'Fail',
        written_energization_plan_approved: isPlanApproved
      },
      energization_8_2_b: {
        input_voltage_and_phase_rotation: {
          rectifier_input_voltage_V: rectifierInputVoltage,
          bypass_input_voltage_V: bypassInputVoltage,
          phase_sequence_match: isPhasingPass ? 'Pass' : 'Fail',
          status: isPhasingPass ? 'Pass' : 'Fail'
        },
        rectifier_startup_and_dc_bus: {
          dc_bus_float_voltage_V: dcBusVoltageV,
          dc_ripple_voltage_percent: dcRipple,
          status: dcRipple <= 1.0 ? 'Pass' : 'Investigate'
        },
        inverter_startup_and_sync: {
          inverter_output_voltage_V: inverterOutputVoltage,
          phase_balance_deviation_percent: phaseBalanceDeviation,
          sync_phase_angle_deg: syncPhaseAngle,
          sync_voltage_diff_V: syncVoltageDiff,
          status: syncPhaseAngle <= 3.0 ? 'Pass' : 'Investigate'
        },
        static_bypass_transfer_test: {
          transfer_time_ms: staticTransferMs,
          transient_voltage_deviation_percent: 2.5,
          status: staticTransferMs <= 4.0 ? 'Pass' : 'Fail'
        },
        loss_of_ac_input_power_test: {
          transfer_to_battery_seamless: isLossOfAcPass,
          output_voltage_drop_V: isLossOfAcPass ? 1.5 : 25.0,
          status: isLossOfAcPass ? 'Pass' : 'Fail'
        },
        battery_discharge_and_full_load_test: {
          load_percentage: 100,
          discharge_duration_minutes: 30,
          end_of_discharge_voltage_V: fullLoadDischargePass ? 420 : 375,
          full_load_capacity_met: fullLoadDischargePass,
          status: fullLoadDischargePass ? 'Pass' : 'Fail'
        },
        output_thd_harmonics: {
          thd_v_percent: thdPercent,
          status: thdPercent <= 3.5 ? 'Pass' : 'Investigate'
        },
        manual_maintenance_bypass_transfer_test: isManualBypassPass ? 'Pass' : 'Fail'
      },
      post_energization_8_2_c: {
        thermographic_survey_table_100_18: {
          max_delta_t_celsius: thermalDeltaT,
          status: isThermalPass ? 'Pass' : 'Investigate',
          hotspot_location: !isThermalPass ? 'Battery String Cell #18 Terminal Lug' : undefined
        },
        scada_bms_telemetry_verified: isScadaPass ? 'Pass' : 'Fail',
        owner_training_and_handover_completed: isTrainingPass
      }
    };
  }, [
    projectName, systemTag, upsRatingKva, acVoltageRatingV, dcBusVoltageV, batteryType,
    batteryCapacityAh, fseName, commissioningDate, isDocApproved, isAtsReviewed, powerIr,
    controlIr, batteryDeviation, isBatteryTorquePass, isBatteryUniformPass, isInterlocksPass,
    isGroundingPass, isPlanApproved, rectifierInputVoltage, bypassInputVoltage, isPhasingPass,
    dcRipple, inverterOutputVoltage, phaseBalanceDeviation, syncPhaseAngle, syncVoltageDiff,
    staticTransferMs, isLossOfAcPass, fullLoadDischargePass, thdPercent, isManualBypassPass,
    thermalDeltaT, isThermalPass, isScadaPass, isTrainingPass
  ]);

  // Evaluation Result State
  const [evaluationResult, setEvaluationResult] = useState<NetaEcsUpsEvaluationResult>(() =>
    performDeterministicEcsUpsAnalysis(currentPayload)
  );

  // Scenario Loader
  const handleSelectScenario = (scenario: PresetScenario) => {
    setSelectedScenarioId(scenario.id);
    setSystemTag(scenario.systemTag);
    setUpsRatingKva(scenario.upsRatingKva);
    setAcVoltageRatingV(scenario.acVoltageRatingV);
    setDcBusVoltageV(scenario.dcBusVoltageV);
    setBatteryType(scenario.batteryType);
    setBatteryCapacityAh(scenario.batteryCapacityAh);
    setPowerIr(scenario.powerIr);
    setControlIr(scenario.controlIr);
    setBatteryDeviation(scenario.batteryDeviation);
    setIsBatteryTorquePass(scenario.batteryTorquePass);
    setIsBatteryUniformPass(scenario.batteryUniformPass);
    setIsInterlocksPass(scenario.isInterlocksPass);
    setIsPhasingPass(scenario.isPhasingPass);
    setDcRipple(scenario.dcRipple);
    setSyncPhaseAngle(scenario.syncPhaseAngle);
    setSyncVoltageDiff(scenario.syncVoltageDiff);
    setStaticTransferMs(scenario.staticTransferMs);
    setIsLossOfAcPass(scenario.isLossOfAcPass);
    setFullLoadDischargePass(scenario.fullLoadDischargePass);
    setThdPercent(scenario.thdPercent);
    setIsManualBypassPass(scenario.isManualBypassPass);
    setThermalDeltaT(scenario.thermalDeltaT);
    setIsThermalPass(scenario.thermalDeltaT <= 3.0);
    setIsScadaPass(scenario.isScadaPass);
    setIsTrainingPass(scenario.isTrainingPass);
    setIsPlanApproved(scenario.isPlanApproved);

    // Immediate local recalculation
    const updatedPayload: NetaEcsUpsInputPayload = {
      ...currentPayload,
      tev_coe_header: {
        ...currentPayload.tev_coe_header,
        system_tag: scenario.systemTag,
        ups_rating_kVA: scenario.upsRatingKva,
        ac_voltage_rating_V: scenario.acVoltageRatingV,
        dc_bus_voltage_V: scenario.dcBusVoltageV,
        battery_type: scenario.batteryType,
        battery_capacity_Ah: scenario.batteryCapacityAh
      },
      pre_energization_8_2_a: {
        ...currentPayload.pre_energization_8_2_a,
        power_circuit_insulation_1000v_megohms: scenario.powerIr,
        control_wiring_insulation_500v_megohms: scenario.controlIr,
        battery_bank_testing: {
          ...currentPayload.pre_energization_8_2_a.battery_bank_testing,
          max_cell_impedance_deviation_percent: scenario.batteryDeviation,
          bolted_connections_torque_table_100_12: scenario.batteryTorquePass ? 'Pass' : 'Fail',
          cell_voltage_uniformity_pass: scenario.batteryUniformPass
        },
        mechanical_and_electrical_interlocks: scenario.isInterlocksPass ? 'Pass' : 'Fail',
        written_energization_plan_approved: scenario.isPlanApproved
      },
      energization_8_2_b: {
        ...currentPayload.energization_8_2_b,
        input_voltage_and_phase_rotation: {
          ...currentPayload.energization_8_2_b.input_voltage_and_phase_rotation,
          phase_sequence_match: scenario.isPhasingPass ? 'Pass' : 'Fail'
        },
        rectifier_startup_and_dc_bus: {
          ...currentPayload.energization_8_2_b.rectifier_startup_and_dc_bus,
          dc_ripple_voltage_percent: scenario.dcRipple
        },
        inverter_startup_and_sync: {
          ...currentPayload.energization_8_2_b.inverter_startup_and_sync,
          sync_phase_angle_deg: scenario.syncPhaseAngle,
          sync_voltage_diff_V: scenario.syncVoltageDiff
        },
        static_bypass_transfer_test: {
          ...currentPayload.energization_8_2_b.static_bypass_transfer_test,
          transfer_time_ms: scenario.staticTransferMs
        },
        loss_of_ac_input_power_test: {
          ...currentPayload.energization_8_2_b.loss_of_ac_input_power_test,
          transfer_to_battery_seamless: scenario.isLossOfAcPass
        },
        battery_discharge_and_full_load_test: {
          ...currentPayload.energization_8_2_b.battery_discharge_and_full_load_test,
          full_load_capacity_met: scenario.fullLoadDischargePass
        },
        output_thd_harmonics: {
          ...currentPayload.energization_8_2_b.output_thd_harmonics,
          thd_v_percent: scenario.thdPercent
        },
        manual_maintenance_bypass_transfer_test: scenario.isManualBypassPass ? 'Pass' : 'Fail'
      },
      post_energization_8_2_c: {
        ...currentPayload.post_energization_8_2_c,
        thermographic_survey_table_100_18: {
          ...currentPayload.post_energization_8_2_c.thermographic_survey_table_100_18,
          max_delta_t_celsius: scenario.thermalDeltaT,
          status: scenario.thermalDeltaT <= 3.0 ? 'Pass' : 'Investigate'
        },
        scada_bms_telemetry_verified: scenario.isScadaPass ? 'Pass' : 'Fail',
        owner_training_and_handover_completed: scenario.isTrainingPass
      }
    };
    setEvaluationResult(performDeterministicEcsUpsAnalysis(updatedPayload));
  };

  // Run AI & Deterministic Analysis
  const handleRunAnalysis = async () => {
    setIsAnalyzing(true);
    setAnalysisError(null);
    try {
      const response = await fetch('/api/field-service/neta-ecs-ups-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(currentPayload)
      });
      const resData = await response.json();
      if (resData.success && resData.data) {
        setEvaluationResult(resData.data);
      } else {
        // Fallback to client-side deterministic evaluation
        setEvaluationResult(performDeterministicEcsUpsAnalysis(currentPayload));
      }
      setActiveTab('evaluation');
    } catch (err: any) {
      console.warn('API error, falling back to deterministic local evaluation:', err);
      setEvaluationResult(performDeterministicEcsUpsAnalysis(currentPayload));
      setActiveTab('evaluation');
    } finally {
      setIsAnalyzing(false);
    }
  };

  // Copy System Instruction
  const handleCopyPrompt = () => {
    navigator.clipboard.writeText(ECS_UPS_SYSTEM_INSTRUCTION);
    setIsCopiedPrompt(true);
    setTimeout(() => setIsCopiedPrompt(false), 2000);
  };

  // Copy JSON Payload
  const handleCopyJson = () => {
    navigator.clipboard.writeText(JSON.stringify(currentPayload, null, 2));
    setIsCopiedJson(true);
    setTimeout(() => setIsCopiedJson(false), 2000);
  };

  // Export PDF Report
  const handleExportPdf = () => {
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const isVn = lang === 'vi';

    // Header styling
    doc.setFillColor(15, 23, 42); // slate-900
    doc.rect(0, 0, 210, 36, 'F');

    doc.setTextColor(248, 250, 252);
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text(
      isVn
        ? 'BÁO CÁO NGHIỆM THU CHẠY THỬ HỆ THỐNG UPS (NETA ECS-2024 MỤC 8.2)'
        : 'UPS COMMISSIONING ACCEPTANCE REPORT (ANSI/NETA ECS-2024 SEC 8.2)',
      14,
      14
    );

    doc.setFontSize(9);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(148, 163, 184);
    doc.text(
      `Standards: ANSI/NETA ECS-2024 Section 8.2 | NETA ATS 7.22.2/7.18 | IEEE 450/1188 | NFPA 111 | NETA Tables 100.1, 100.12, 100.18`,
      14,
      21
    );
    doc.text(
      `Project: ${projectName} | System: ${systemTag} | Rating: ${upsRatingKva}kVA (${acVoltageRatingV}V AC, ${dcBusVoltageV}V DC)`,
      14,
      27
    );
    doc.text(
      `Lead FSE: ${fseName} | Date: ${commissioningDate} | Battery: ${batteryType}`,
      14,
      32
    );

    // Status Banner
    const isCritical = evaluationResult.overall_status === 'CRITICAL';
    const isCaution = evaluationResult.overall_status === 'CAUTION';
    if (isCritical) {
      doc.setFillColor(244, 63, 94); // rose-500
    } else if (isCaution) {
      doc.setFillColor(245, 158, 11); // amber-500
    } else {
      doc.setFillColor(16, 185, 129); // emerald-500
    }
    doc.rect(14, 40, 182, 12, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(11);
    doc.setFont('helvetica', 'bold');
    doc.text(
      isVn ? evaluationResult.status_title_vn : evaluationResult.status_title_en,
      18,
      48
    );

    // Technical Analysis Summary Box
    doc.setFillColor(241, 245, 249);
    doc.rect(14, 56, 182, 26, 'F');
    doc.setTextColor(30, 41, 59);
    doc.setFontSize(10);
    doc.setFont('helvetica', 'bold');
    doc.text(
      isVn
        ? evaluationResult.technical_analysis.title_vn
        : evaluationResult.technical_analysis.title_en,
      18,
      62
    );

    doc.setFontSize(8);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(51, 65, 85);
    const splitDesc = doc.splitTextToSize(
      isVn
        ? evaluationResult.technical_analysis.description_vn
        : evaluationResult.technical_analysis.description_en,
      174
    );
    doc.text(splitDesc, 18, 68);

    // KPI Metrics bar
    doc.setFontSize(8);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(15, 23, 42);
    const kpiText = `Static Transfer: ${evaluationResult.technical_analysis.static_transfer_time_ms} ms (0ms target) | Battery Dev: ${evaluationResult.technical_analysis.battery_impedance_deviation_percent}% (<=20%) | Power IR: ${evaluationResult.technical_analysis.power_ir_megohms} M Ohm (>=100M) | THD-v: ${evaluationResult.technical_analysis.thd_v_percent}% (<=3%) | Delta-T: ${evaluationResult.technical_analysis.thermal_delta_t}°C`;
    doc.text(kpiText, 18, 80);

    // Evaluation Table
    const tableData = evaluationResult.evaluation_table.map((row) => [
      row.item_code,
      isVn ? row.item_name_vn : row.item_name_en,
      row.neta_standard_ref,
      row.measured_value,
      row.status,
      isVn ? row.evaluation_vn : row.evaluation_en
    ]);

    autoTable(doc, {
      startY: 86,
      head: [
        [
          'Item',
          isVn ? 'Hạng Mục Kiểm Tra Chạy Thử' : 'Commissioning Test Item',
          isVn ? 'Tiêu Chuẩn NETA' : 'Standard Ref',
          isVn ? 'Giá Trị Đo Tại Site' : 'Measured Value',
          'Status',
          isVn ? 'Đánh Giá Kỹ Thuật TEV COE' : 'Engineering Finding'
        ]
      ],
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
        cellPadding: 2,
        overflow: 'linebreak'
      },
      columnStyles: {
        0: { cellWidth: 16 },
        1: { cellWidth: 46 },
        2: { cellWidth: 34 },
        3: { cellWidth: 30 },
        4: { cellWidth: 16, fontStyle: 'bold' },
        5: { cellWidth: 40 }
      },
      didParseCell: function (data) {
        if (data.column.index === 4) {
          if (data.cell.raw === 'FAIL') {
            data.cell.styles.textColor = [225, 29, 72];
          } else if (data.cell.raw === 'PASS') {
            data.cell.styles.textColor = [16, 185, 129];
          } else {
            data.cell.styles.textColor = [217, 119, 6];
          }
        }
      }
    });

    // Recommendations section
    const finalY = (doc as any).lastAutoTable?.finalY || 240;
    if (finalY < 255) {
      doc.setFontSize(9);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(15, 23, 42);
      doc.text(
        isVn
          ? 'CÁC BƯỚC KHẮC PHỤC & KHUYẾN NGHỊ BÀN GIAO (ACTIONABLE STEPS):'
          : 'ACTIONABLE PROTOCOL RECOMMENDATIONS:',
        14,
        finalY + 8
      );

      doc.setFontSize(7.5);
      doc.setFont('helvetica', 'normal');
      evaluationResult.actionable_recommendations.steps.forEach((st, idx) => {
        const title = isVn ? st.title_vn : st.title_en;
        const detail = isVn ? st.details_vn : st.details_en;
        doc.setFont('helvetica', 'bold');
        doc.text(`${st.step_number}. ${title}`, 16, finalY + 14 + idx * 9);
        doc.setFont('helvetica', 'normal');
        doc.text(detail, 18, finalY + 18 + idx * 9, { maxWidth: 178 });
      });
    }

    doc.save(`NETA_ECS_8.2_UPS_${systemTag.replace(/[^a-zA-Z0-9]/g, '_')}_Report.pdf`);
  };

  // Save to CMMS
  const handleSaveToCmms = () => {
    if (onSaveReport) {
      onSaveReport({
        id: `ECS-UPS-${Date.now()}`,
        standard: 'ANSI/NETA ECS-2024 Section 8.2',
        title: `UPS Commissioning: ${systemTag}`,
        status: evaluationResult.overall_status,
        date: commissioningDate,
        fse: fseName,
        payload: currentPayload,
        result: evaluationResult
      });
    }
  };

  return (
    <div className="space-y-6 text-slate-100">
      {/* Top Banner */}
      <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 p-6 rounded-2xl border border-indigo-500/30 shadow-2xl relative overflow-hidden">
        <div className="absolute top-0 right-0 p-8 opacity-10 pointer-events-none">
          <Zap size={180} className="text-indigo-400" />
        </div>

        <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4 relative z-10">
          <div>
            <div className="flex items-center gap-2 flex-wrap mb-2">
              <span className="px-2.5 py-1 text-xs font-bold rounded-md bg-indigo-500/20 text-indigo-300 border border-indigo-400/30">
                ANSI/NETA ECS-2024 Mục 8.2
              </span>
              <span className="px-2.5 py-1 text-xs font-semibold rounded-md bg-slate-800 text-slate-300 border border-slate-700">
                NETA ATS 7.22.2 & 7.18
              </span>
              <span className="px-2.5 py-1 text-xs font-semibold rounded-md bg-slate-800 text-slate-300 border border-slate-700">
                IEEE 450 / 1188 / 1184
              </span>
              <span className="px-2.5 py-1 text-xs font-semibold rounded-md bg-slate-800 text-slate-300 border border-slate-700">
                NFPA 111 (SEPSS)
              </span>
              <span className="px-2.5 py-1 text-xs font-semibold rounded-md bg-slate-800 text-slate-300 border border-slate-700">
                NETA Tables 100.1, 100.12, 100.18
              </span>
            </div>

            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white flex items-center gap-2">
              <Zap className="text-indigo-400 shrink-0" size={26} />
              <span>
                {lang === 'vi'
                  ? 'Quy Trình Chạy Thử & Nghiệm Thu Hệ Thống UPS (Source Specific Systems)'
                  : 'Source Specific Systems: Uninterruptible Power Supply (UPS) Commissioning'}
              </span>
            </h1>

            <p className="text-sm text-slate-300 mt-1 max-w-4xl">
              {lang === 'vi'
                ? 'Nền tảng kiểm định và chạy thử hệ thống nguồn liên tục (UPS) trung tâm dữ liệu, y tế theo ANSI/NETA ECS-2024 Mục 8.2. Đánh giá tự động chuyển mạch tĩnh 0ms, nội trở ắc-quy IEEE 1188, cách điện Table 100.1, xả tải giả 100% NFPA 111 và chụp nhiệt Table 100.18.'
                : 'Commissioning & acceptance testing engine for mission-critical UPS systems per ANSI/NETA ECS-2024 Section 8.2. Automated validation of 0ms static transfer, IEEE 1188 ohmic health, Table 100.1 insulation, NFPA 111 100% load discharge, and Table 100.18 infrared scans.'}
            </p>
          </div>

          {/* Action buttons */}
          <div className="flex items-center gap-2 flex-wrap">
            <button
              onClick={() => setLang(lang === 'vi' ? 'en' : 'vi')}
              className="px-3 py-1.5 rounded-lg bg-slate-800/80 hover:bg-slate-700 text-xs font-medium border border-slate-700 flex items-center gap-1.5 transition"
            >
              <Globe size={14} className="text-indigo-400" />
              <span>{lang === 'vi' ? 'EN' : 'VN'}</span>
            </button>

            <button
              onClick={handleExportPdf}
              className="px-3.5 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold shadow-lg shadow-indigo-600/30 flex items-center gap-1.5 transition"
            >
              <Download size={14} />
              <span>{lang === 'vi' ? 'Xuất PDF A4' : 'Export PDF'}</span>
            </button>

            {onSaveReport && (
              <button
                onClick={handleSaveToCmms}
                className="px-3.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold shadow-lg shadow-emerald-600/30 flex items-center gap-1.5 transition"
              >
                <Save size={14} />
                <span>{lang === 'vi' ? 'Lưu CMMS' : 'Save CMMS'}</span>
              </button>
            )}
          </div>
        </div>

        {/* Live KPI Grid */}
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3 mt-6">
          <div className="bg-slate-950/70 p-3 rounded-xl border border-slate-800">
            <div className="text-xs text-slate-400 flex items-center gap-1">
              <Clock size={12} className="text-indigo-400" />
              <span>Chuyển Mạch Tĩnh</span>
            </div>
            <div className="text-lg font-bold text-white mt-1">
              {evaluationResult.technical_analysis.static_transfer_time_ms} ms
            </div>
            <div className="text-[10px] text-emerald-400 mt-0.5">Mục tiêu: 0 ms (No Break)</div>
          </div>

          <div className="bg-slate-950/70 p-3 rounded-xl border border-slate-800">
            <div className="text-xs text-slate-400 flex items-center gap-1">
              <BatteryCharging size={12} className="text-amber-400" />
              <span>Lệch Trở Kháng Pin</span>
            </div>
            <div
              className={`text-lg font-bold mt-1 ${
                evaluationResult.technical_analysis.battery_impedance_deviation_percent > 20
                  ? 'text-rose-400'
                  : 'text-emerald-400'
              }`}
            >
              {evaluationResult.technical_analysis.battery_impedance_deviation_percent}%
            </div>
            <div className="text-[10px] text-slate-400 mt-0.5">IEEE 1188: ≤ 20%</div>
          </div>

          <div className="bg-slate-950/70 p-3 rounded-xl border border-slate-800">
            <div className="text-xs text-slate-400 flex items-center gap-1">
              <ShieldCheck size={12} className="text-cyan-400" />
              <span>Cách Điện Động Lực</span>
            </div>
            <div
              className={`text-lg font-bold mt-1 ${
                evaluationResult.technical_analysis.power_ir_megohms >= 100
                  ? 'text-emerald-400'
                  : 'text-rose-400'
              }`}
            >
              {evaluationResult.technical_analysis.power_ir_megohms} MΩ
            </div>
            <div className="text-[10px] text-slate-400 mt-0.5">Table 100.1: ≥ 100 MΩ</div>
          </div>

          <div className="bg-slate-950/70 p-3 rounded-xl border border-slate-800">
            <div className="text-xs text-slate-400 flex items-center gap-1">
              <Activity size={12} className="text-purple-400" />
              <span>Méo Sóng Hài THD-v</span>
            </div>
            <div
              className={`text-lg font-bold mt-1 ${
                evaluationResult.technical_analysis.thd_v_percent > 3.0
                  ? 'text-amber-400'
                  : 'text-emerald-400'
              }`}
            >
              {evaluationResult.technical_analysis.thd_v_percent}%
            </div>
            <div className="text-[10px] text-slate-400 mt-0.5">IEEE 519: ≤ 3.0%</div>
          </div>

          <div className="bg-slate-950/70 p-3 rounded-xl border border-slate-800">
            <div className="text-xs text-slate-400 flex items-center gap-1">
              <Thermometer size={12} className="text-rose-400" />
              <span>Chụp Nhiệt ΔT</span>
            </div>
            <div
              className={`text-lg font-bold mt-1 ${
                (evaluationResult.technical_analysis.thermal_delta_t ?? 0) > 3.0
                  ? 'text-rose-400'
                  : 'text-emerald-400'
              }`}
            >
              {(evaluationResult.technical_analysis.thermal_delta_t ?? 0).toFixed(1)}°C
            </div>
            <div className="text-[10px] text-slate-400 mt-0.5">Table 100.18: ≤ 3.0°C</div>
          </div>
        </div>
      </div>

      {/* Preset Scenarios Selector */}
      <div className="bg-slate-900/90 p-4 rounded-xl border border-slate-800">
        <div className="flex items-center justify-between gap-2 mb-3">
          <span className="text-xs font-bold uppercase tracking-wider text-slate-400 flex items-center gap-1.5">
            <Sliders size={14} className="text-indigo-400" />
            <span>Kịch Bản Nghiệm Thu Thực Địa Mẫu (1-Click Benchmarks):</span>
          </span>
          <span className="text-[11px] text-slate-500">
            ANSI/NETA ECS-2024 Mục 8.2 Testing Scenarios
          </span>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-2.5">
          {PRESET_SCENARIOS.map((sc) => {
            const isSelected = selectedScenarioId === sc.id;
            return (
              <button
                key={sc.id}
                onClick={() => handleSelectScenario(sc)}
                className={`text-left p-3 rounded-lg border transition ${
                  isSelected
                    ? 'bg-indigo-950/60 border-indigo-500/80 shadow-md ring-1 ring-indigo-500/40'
                    : 'bg-slate-950/40 border-slate-800 hover:border-slate-700 hover:bg-slate-900/60'
                }`}
              >
                <div className="flex items-center justify-between gap-2">
                  <span className="text-xs font-semibold text-white">
                    {lang === 'vi' ? sc.name : sc.nameEn}
                  </span>
                  <span className={`text-[10px] px-2 py-0.5 rounded border ${sc.badgeColor}`}>
                    {sc.badge}
                  </span>
                </div>
                <p className="text-[11px] text-slate-400 mt-1 line-clamp-2">{sc.description}</p>
              </button>
            );
          })}
        </div>
      </div>

      {/* Sub Tabs */}
      <div className="flex items-center gap-2 border-b border-slate-800 pb-2">
        <button
          onClick={() => setActiveTab('checklist')}
          className={`px-3.5 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1.5 transition ${
            activeTab === 'checklist'
              ? 'bg-indigo-600 text-white font-semibold shadow'
              : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800'
          }`}
        >
          <ClipboardCheck size={14} />
          <span>{lang === 'vi' ? 'Bảng Checklist Chạy Thử (8.2)' : 'Commissioning Checklist'}</span>
        </button>

        <button
          onClick={() => setActiveTab('evaluation')}
          className={`px-3.5 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1.5 transition ${
            activeTab === 'evaluation'
              ? 'bg-indigo-600 text-white font-semibold shadow'
              : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800'
          }`}
        >
          <Activity size={14} />
          <span>
            {lang === 'vi' ? 'Kết Quả Thẩm Định & Khuyến Nghị' : 'Diagnostic Evaluation'}
          </span>
          <span
            className={`ml-1 text-[10px] px-1.5 py-0.2 rounded font-bold ${
              evaluationResult.overall_status === 'CRITICAL'
                ? 'bg-rose-500/20 text-rose-300'
                : evaluationResult.overall_status === 'CAUTION'
                ? 'bg-amber-500/20 text-amber-300'
                : 'bg-emerald-500/20 text-emerald-300'
            }`}
          >
            {evaluationResult.overall_status}
          </span>
        </button>

        <button
          onClick={() => setActiveTab('trend')}
          className={`px-3.5 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1.5 transition ${
            activeTab === 'trend'
              ? 'bg-indigo-600 text-white font-semibold shadow'
              : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800'
          }`}
        >
          <TrendingUp size={14} className="text-amber-400" />
          <span>
            {lang === 'vi' ? 'Xu Hướng Nội Trở (IEEE 1188)' : 'Ohmic Trend Chart'}
          </span>
        </button>

        <button
          onClick={() => setActiveTab('prompt')}
          className={`px-3.5 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1.5 transition ${
            activeTab === 'prompt'
              ? 'bg-indigo-600 text-white font-semibold shadow'
              : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800'
          }`}
        >
          <Sparkles size={14} className="text-amber-400" />
          <span>System Prompt AI Studio</span>
        </button>

        <button
          onClick={() => setActiveTab('json')}
          className={`px-3.5 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1.5 transition ${
            activeTab === 'json'
              ? 'bg-indigo-600 text-white font-semibold shadow'
              : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800'
          }`}
        >
          <Cpu size={14} />
          <span>Mẫu JSON Đầu Vào</span>
        </button>
      </div>

      {/* TAB 1: CHECKLIST & MEASUREMENT INPUT FORM */}
      {activeTab === 'checklist' && (
        <div className="space-y-6">
          {/* Section Header Inputs */}
          <div className="bg-slate-900/80 p-4 rounded-xl border border-slate-800 space-y-4">
            <h3 className="text-xs font-bold uppercase tracking-wider text-indigo-400 flex items-center gap-2">
              <Layers size={14} />
              <span>Thông Tin Nhãn Mác & Cấu Hình Hệ Thống UPS (TEV COE Header)</span>
            </h3>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
              <div>
                <label className="text-[11px] text-slate-400 block mb-1">Dự Án (Project Name)</label>
                <input
                  type="text"
                  value={projectName}
                  onChange={(e) => setProjectName(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-white focus:border-indigo-500 outline-none"
                />
              </div>

              <div>
                <label className="text-[11px] text-slate-400 block mb-1">Mã Thiết Bị (System Tag)</label>
                <input
                  type="text"
                  value={systemTag}
                  onChange={(e) => setSystemTag(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-white focus:border-indigo-500 outline-none"
                />
              </div>

              <div>
                <label className="text-[11px] text-slate-400 block mb-1">Công Suất UPS (kVA)</label>
                <input
                  type="number"
                  value={upsRatingKva}
                  onChange={(e) => setUpsRatingKva(Number(e.target.value))}
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-white focus:border-indigo-500 outline-none"
                />
              </div>

              <div>
                <label className="text-[11px] text-slate-400 block mb-1">Điện Áp AC Định Mức (V)</label>
                <input
                  type="number"
                  value={acVoltageRatingV}
                  onChange={(e) => setAcVoltageRatingV(Number(e.target.value))}
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-white focus:border-indigo-500 outline-none"
                />
              </div>

              <div>
                <label className="text-[11px] text-slate-400 block mb-1">Điện Áp DC Bus Nạp Float (V)</label>
                <input
                  type="number"
                  value={dcBusVoltageV}
                  onChange={(e) => setDcBusVoltageV(Number(e.target.value))}
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-white focus:border-indigo-500 outline-none"
                />
              </div>

              <div>
                <label className="text-[11px] text-slate-400 block mb-1">Loại Ắc-quy (Battery Type)</label>
                <input
                  type="text"
                  value={batteryType}
                  onChange={(e) => setBatteryType(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-white focus:border-indigo-500 outline-none"
                />
              </div>

              <div>
                <label className="text-[11px] text-slate-400 block mb-1">Dung Lượng Dàn Bình (Ah)</label>
                <input
                  type="number"
                  value={batteryCapacityAh}
                  onChange={(e) => setBatteryCapacityAh(Number(e.target.value))}
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-white focus:border-indigo-500 outline-none"
                />
              </div>

              <div>
                <label className="text-[11px] text-slate-400 block mb-1">Kỹ Sư Trưởng Site (FSE)</label>
                <input
                  type="text"
                  value={fseName}
                  onChange={(e) => setFseName(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-white focus:border-indigo-500 outline-none"
                />
              </div>
            </div>
          </div>

          {/* Giai đoạn 1: Tiền chạy thử (8.2.A) */}
          <div className="bg-slate-900/80 p-4 rounded-xl border border-slate-800 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-800 pb-2">
              <h3 className="text-xs font-bold uppercase tracking-wider text-teal-400 flex items-center gap-2">
                <ShieldCheck size={14} />
                <span>Giai Đoạn 1: Tiền Chạy Thử / Trước Khi Đóng Điện (Pre-Energization 8.2.A)</span>
              </h3>
              <span className="text-[11px] text-slate-500">ANSI/NETA ECS-2024 Mục 8.2.A</span>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
              {/* A1 */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800 flex items-center justify-between">
                <div>
                  <div className="text-xs font-semibold text-white">8.2.A.1 - Hồ Sơ & Bản Vẽ</div>
                  <div className="text-[10px] text-slate-400">OPR, BOD, sơ đồ nguyên lý DC/AC</div>
                </div>
                <button
                  type="button"
                  onClick={() => setIsDocApproved(!isDocApproved)}
                  className={`px-3 py-1 text-xs font-bold rounded ${
                    isDocApproved
                      ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'
                      : 'bg-rose-500/20 text-rose-300 border border-rose-500/40'
                  }`}
                >
                  {isDocApproved ? 'Pass / Đã Duyệt' : 'Fail / Thiếu'}
                </button>
              </div>

              {/* A2 */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800 flex items-center justify-between">
                <div>
                  <div className="text-xs font-semibold text-white">8.2.A.3 - Báo Cáo ATS 7.22.2/7.18</div>
                  <div className="text-[10px] text-slate-400">Nghiệm thu chấp nhận tủ UPS & bình</div>
                </div>
                <button
                  type="button"
                  onClick={() => setIsAtsReviewed(!isAtsReviewed)}
                  className={`px-3 py-1 text-xs font-bold rounded ${
                    isAtsReviewed
                      ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'
                      : 'bg-rose-500/20 text-rose-300 border border-rose-500/40'
                  }`}
                >
                  {isAtsReviewed ? 'Pass / Đạt Báo Cáo' : 'Fail / Thiếu'}
                </button>
              </div>

              {/* A3 - Power IR */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800">
                <div className="flex justify-between items-center mb-1">
                  <span className="text-xs font-semibold text-white">8.2.A.15 - Cách Điện Động Lực</span>
                  <span className="text-[10px] text-slate-400">≥ 100 MΩ (Table 100.1)</span>
                </div>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    value={powerIr}
                    onChange={(e) => setPowerIr(Number(e.target.value))}
                    className="w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-white"
                  />
                  <span className="text-xs text-slate-400">MΩ</span>
                </div>
              </div>

              {/* A3 - Control IR */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800">
                <div className="flex justify-between items-center mb-1">
                  <span className="text-xs font-semibold text-white">8.2.A.16 - Cách Điện Mạch Điều Khiển</span>
                  <span className="text-[10px] text-slate-400">≥ 1.0 MΩ (500V DC)</span>
                </div>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    value={controlIr}
                    onChange={(e) => setControlIr(Number(e.target.value))}
                    className="w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-white"
                  />
                  <span className="text-xs text-slate-400">MΩ</span>
                </div>
              </div>

              {/* A4 - Battery Deviation */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800">
                <div className="flex justify-between items-center mb-1">
                  <span className="text-xs font-semibold text-white">8.2.A.4 - Lệch Nội Trở Ắc-quy</span>
                  <span className="text-[10px] text-slate-400">IEEE 1188: ≤ 20%</span>
                </div>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    value={batteryDeviation}
                    onChange={(e) => setBatteryDeviation(Number(e.target.value))}
                    className="w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-white"
                  />
                  <span className="text-xs text-slate-400">%</span>
                </div>
                <div className="flex justify-between items-center mt-1.5 pt-1.5 border-t border-slate-800/80">
                  <span className="text-[10px] text-slate-500">IEEE 1188 Trend</span>
                  <button
                    type="button"
                    onClick={() => setActiveTab('trend')}
                    className="text-[10px] text-indigo-400 hover:text-indigo-300 font-semibold flex items-center gap-1 transition"
                  >
                    <TrendingUp size={11} />
                    <span>Xem Biểu Đồ Lịch Sử</span>
                  </button>
                </div>
              </div>

              {/* A4 - Battery Torque & Uniformity */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800 flex items-center justify-between">
                <div>
                  <div className="text-xs font-semibold text-white">8.2.A.4 - Lực Siết Cọc Bình 100.12</div>
                  <div className="text-[10px] text-slate-400">Cờ-lê lực kiểm định</div>
                </div>
                <button
                  type="button"
                  onClick={() => setIsBatteryTorquePass(!isBatteryTorquePass)}
                  className={`px-3 py-1 text-xs font-bold rounded ${
                    isBatteryTorquePass
                      ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'
                      : 'bg-rose-500/20 text-rose-300 border border-rose-500/40'
                  }`}
                >
                  {isBatteryTorquePass ? 'Pass / Đủ Lực' : 'Fail / Lỏng'}
                </button>
              </div>

              {/* A5 - Interlocks */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800 flex items-center justify-between">
                <div>
                  <div className="text-xs font-semibold text-white">8.2.A.18 - Liên Động MBB vs Inverter</div>
                  <div className="text-[10px] text-slate-400">Khóa cơ & khóa điện tránh chập</div>
                </div>
                <button
                  type="button"
                  onClick={() => setIsInterlocksPass(!isInterlocksPass)}
                  className={`px-3 py-1 text-xs font-bold rounded ${
                    isInterlocksPass
                      ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'
                      : 'bg-rose-500/20 text-rose-300 border border-rose-500/40'
                  }`}
                >
                  {isInterlocksPass ? 'Pass / Hoàn Hảo' : 'Fail / Lỗi Khóa'}
                </button>
              </div>

              {/* A6 - Grounding */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800 flex items-center justify-between">
                <div>
                  <div className="text-xs font-semibold text-white">8.2.A.11 - Tiếp Địa & Trung Tính SDS</div>
                  <div className="text-[10px] text-slate-400">Điện trở tiếp địa ≤ 5.0 Ω</div>
                </div>
                <button
                  type="button"
                  onClick={() => setIsGroundingPass(!isGroundingPass)}
                  className={`px-3 py-1 text-xs font-bold rounded ${
                    isGroundingPass
                      ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'
                      : 'bg-rose-500/20 text-rose-300 border border-rose-500/40'
                  }`}
                >
                  {isGroundingPass ? 'Pass / Đạt' : 'Fail / Cao'}
                </button>
              </div>

              {/* A7 - Written Plan */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800 flex items-center justify-between">
                <div>
                  <div className="text-xs font-semibold text-white">8.2.A.38 - Kế Hoạch Đóng Điện & BHLĐ</div>
                  <div className="text-[10px] text-slate-400">Trang bị Arc Flash PPE</div>
                </div>
                <button
                  type="button"
                  onClick={() => setIsPlanApproved(!isPlanApproved)}
                  className={`px-3 py-1 text-xs font-bold rounded ${
                    isPlanApproved
                      ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'
                      : 'bg-rose-500/20 text-rose-300 border border-rose-500/40'
                  }`}
                >
                  {isPlanApproved ? 'Pass / Đã Duyệt' : 'Fail / Chưa Có'}
                </button>
              </div>
            </div>
          </div>

          {/* Giai đoạn 2: Trong khi đóng điện & Chạy thử chức năng (8.2.B) */}
          <div className="bg-slate-900/80 p-4 rounded-xl border border-slate-800 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-800 pb-2">
              <h3 className="text-xs font-bold uppercase tracking-wider text-indigo-400 flex items-center gap-2">
                <Zap size={14} />
                <span>Giai Đoạn 2: Trong Đóng Điện & Chạy Thử Chức Năng (Energization Tests 8.2.B)</span>
              </h3>
              <span className="text-[11px] text-slate-500">ANSI/NETA ECS-2024 Mục 8.2.B</span>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
              {/* B1 - Phase Rotation */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800 flex items-center justify-between">
                <div>
                  <div className="text-xs font-semibold text-white">8.2.B.3 - Thứ Tự Pha AC Ngõ Vào</div>
                  <div className="text-[10px] text-slate-400">Khớp Rectifier & Bypass</div>
                </div>
                <button
                  type="button"
                  onClick={() => setIsPhasingPass(!isPhasingPass)}
                  className={`px-3 py-1 text-xs font-bold rounded ${
                    isPhasingPass
                      ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'
                      : 'bg-rose-500/20 text-rose-300 border border-rose-500/40'
                  }`}
                >
                  {isPhasingPass ? 'Pass / Đồng Pha A-B-C' : 'Fail / Ngược Pha'}
                </button>
              </div>

              {/* B2 - DC Ripple */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800">
                <div className="flex justify-between items-center mb-1">
                  <span className="text-xs font-semibold text-white">8.2.B.8 - Độ Nhấp Nhô DC Ripple</span>
                  <span className="text-[10px] text-slate-400">Mục tiêu: ≤ 1.0%</span>
                </div>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    step="0.05"
                    value={dcRipple}
                    onChange={(e) => setDcRipple(Number(e.target.value))}
                    className="w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-white"
                  />
                  <span className="text-xs text-slate-400">%</span>
                </div>
              </div>

              {/* B3 - Inverter Sync */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800">
                <div className="flex justify-between items-center mb-1">
                  <span className="text-xs font-semibold text-white">8.2.B.9 - Lệch Góc Pha Đồng Bộ</span>
                  <span className="text-[10px] text-slate-400">Khóa pha Bypass: ≤ 3.0°</span>
                </div>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    step="0.1"
                    value={syncPhaseAngle}
                    onChange={(e) => setSyncPhaseAngle(Number(e.target.value))}
                    className="w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-white"
                  />
                  <span className="text-xs text-slate-400">độ</span>
                </div>
              </div>

              {/* B4 - Static Bypass Transfer */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800">
                <div className="flex justify-between items-center mb-1">
                  <span className="text-xs font-semibold text-white">8.2.B.10 - Chuyển Mạch Tĩnh 0ms</span>
                  <span className="text-[10px] text-slate-400">No Break (≤ 4.0 ms)</span>
                </div>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    step="0.5"
                    value={staticTransferMs}
                    onChange={(e) => setStaticTransferMs(Number(e.target.value))}
                    className="w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-white"
                  />
                  <span className="text-xs text-slate-400">ms</span>
                </div>
              </div>

              {/* B5 - Loss of AC */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800 flex items-center justify-between">
                <div>
                  <div className="text-xs font-semibold text-white">8.2.B.11 - Cắt Mất Lưới Đột Ngột</div>
                  <div className="text-[10px] text-slate-400">Chuyển sang ắc-quy tức thời</div>
                </div>
                <button
                  type="button"
                  onClick={() => setIsLossOfAcPass(!isLossOfAcPass)}
                  className={`px-3 py-1 text-xs font-bold rounded ${
                    isLossOfAcPass
                      ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'
                      : 'bg-rose-500/20 text-rose-300 border border-rose-500/40'
                  }`}
                >
                  {isLossOfAcPass ? 'Pass / Không Sụt Áp' : 'Fail / Sập Nguồn'}
                </button>
              </div>

              {/* B6 - Full Load Battery Discharge */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800 flex items-center justify-between">
                <div>
                  <div className="text-xs font-semibold text-white">8.2.B.12 - Xả 100% Tải Giả Load Bank</div>
                  <div className="text-[10px] text-slate-400">NFPA 111 / IEEE 450</div>
                </div>
                <button
                  type="button"
                  onClick={() => setFullLoadDischargePass(!fullLoadDischargePass)}
                  className={`px-3 py-1 text-xs font-bold rounded ${
                    fullLoadDischargePass
                      ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'
                      : 'bg-rose-500/20 text-rose-300 border border-rose-500/40'
                  }`}
                >
                  {fullLoadDischargePass ? 'Pass / Đạt Thời Gian' : 'Fail / Sụt Sớm'}
                </button>
              </div>

              {/* B7 - THD-v */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800">
                <div className="flex justify-between items-center mb-1">
                  <span className="text-xs font-semibold text-white">8.2.B.13 - Méo Sóng Hài THD-v</span>
                  <span className="text-[10px] text-slate-400">IEEE 519: ≤ 3.0%</span>
                </div>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    step="0.1"
                    value={thdPercent}
                    onChange={(e) => setThdPercent(Number(e.target.value))}
                    className="w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-white"
                  />
                  <span className="text-xs text-slate-400">%</span>
                </div>
              </div>

              {/* B8 - Manual MBB */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800 flex items-center justify-between">
                <div>
                  <div className="text-xs font-semibold text-white">8.2.B.14 - Chuyển MBB Bảo Trì Bằng Tay</div>
                  <div className="text-[10px] text-slate-400">Cô lập UPS mà không gián đoạn</div>
                </div>
                <button
                  type="button"
                  onClick={() => setIsManualBypassPass(!isManualBypassPass)}
                  className={`px-3 py-1 text-xs font-bold rounded ${
                    isManualBypassPass
                      ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'
                      : 'bg-rose-500/20 text-rose-300 border border-rose-500/40'
                  }`}
                >
                  {isManualBypassPass ? 'Pass / Không Mất Tải' : 'Fail / Mất Tải'}
                </button>
              </div>
            </div>
          </div>

          {/* Giai đoạn 3: Sau đóng điện & Nghiệm thu bàn giao (8.2.C) */}
          <div className="bg-slate-900/80 p-4 rounded-xl border border-slate-800 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-800 pb-2">
              <h3 className="text-xs font-bold uppercase tracking-wider text-rose-400 flex items-center gap-2">
                <Thermometer size={14} />
                <span>Giai Đoạn 3: Sau Đóng Điện & Bàn Giao (Post-Energization 8.2.C)</span>
              </h3>
              <span className="text-[11px] text-slate-500">ANSI/NETA ECS-2024 Mục 8.2.C</span>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
              {/* C1 - Thermographic Survey */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800">
                <div className="flex justify-between items-center mb-1">
                  <span className="text-xs font-semibold text-white">8.2.C.1 - Chụp Nhiệt Hồng Ngoại IR</span>
                  <span className="text-[10px] text-slate-400">Table 100.18: ΔT ≤ 3.0°C</span>
                </div>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    step="0.5"
                    value={thermalDeltaT}
                    onChange={(e) => {
                      const val = Number(e.target.value);
                      setThermalDeltaT(val);
                      setIsThermalPass(val <= 3.0);
                    }}
                    className="w-full bg-slate-900 border border-slate-700 rounded px-2 py-1 text-xs text-white"
                  />
                  <span className="text-xs text-slate-400">°C</span>
                </div>
              </div>

              {/* C2 - SCADA */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800 flex items-center justify-between">
                <div>
                  <div className="text-xs font-semibold text-white">8.2.C.2 - Giám Sát SCADA / BMS / DCIM</div>
                  <div className="text-[10px] text-slate-400">Modbus/SNMP cảnh báo</div>
                </div>
                <button
                  type="button"
                  onClick={() => setIsScadaPass(!isScadaPass)}
                  className={`px-3 py-1 text-xs font-bold rounded ${
                    isScadaPass
                      ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'
                      : 'bg-rose-500/20 text-rose-300 border border-rose-500/40'
                  }`}
                >
                  {isScadaPass ? 'Pass / Kết Nối Đủ' : 'Fail / Mất Tín Hiệu'}
                </button>
              </div>

              {/* C3 - Training */}
              <div className="bg-slate-950 p-3 rounded-lg border border-slate-800 flex items-center justify-between">
                <div>
                  <div className="text-xs font-semibold text-white">8.2.C.3 - Đào Tạo Vận Hành & Hoàn Công</div>
                  <div className="text-[10px] text-slate-400">Bàn giao quy trình MBB</div>
                </div>
                <button
                  type="button"
                  onClick={() => setIsTrainingPass(!isTrainingPass)}
                  className={`px-3 py-1 text-xs font-bold rounded ${
                    isTrainingPass
                      ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'
                      : 'bg-rose-500/20 text-rose-300 border border-rose-500/40'
                  }`}
                >
                  {isTrainingPass ? 'Pass / Hoàn Thành' : 'Fail / Chưa Đào Tạo'}
                </button>
              </div>
            </div>
          </div>

          {/* Trigger Analysis Button */}
          <div className="flex items-center justify-end gap-3 pt-2">
            <button
              onClick={() => {
                setEvaluationResult(performDeterministicEcsUpsAnalysis(currentPayload));
                setActiveTab('evaluation');
              }}
              className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-medium text-slate-300 transition"
            >
              Chạy Phân Tích Tất Định (Deterministic)
            </button>

            <button
              onClick={handleRunAnalysis}
              disabled={isAnalyzing}
              className="px-5 py-2.5 rounded-xl bg-gradient-to-r from-indigo-600 to-indigo-500 hover:from-indigo-500 hover:to-indigo-400 text-white text-xs font-bold shadow-lg shadow-indigo-600/30 flex items-center gap-2 transition disabled:opacity-50"
            >
              <Sparkles size={16} className={isAnalyzing ? 'animate-spin' : ''} />
              <span>
                {isAnalyzing
                  ? 'Đang Phân Tích Gemini AI Studio...'
                  : 'Phân Tích Chuyên Sâu AI & Xuất Thẩm Định'}
              </span>
            </button>
          </div>
        </div>
      )}

      {/* TAB 2: DIAGNOSTIC EVALUATION RESULTS */}
      {activeTab === 'evaluation' && (
        <div className="space-y-6">
          {/* Severity Banner */}
          <div
            className={`p-5 rounded-2xl border shadow-xl flex flex-col md:flex-row items-start md:items-center justify-between gap-4 ${
              evaluationResult.overall_status === 'CRITICAL'
                ? 'bg-rose-950/40 border-rose-500/60 text-rose-200'
                : evaluationResult.overall_status === 'CAUTION'
                ? 'bg-amber-950/40 border-amber-500/60 text-amber-200'
                : 'bg-emerald-950/40 border-emerald-500/60 text-emerald-200'
            }`}
          >
            <div className="flex items-center gap-3">
              {evaluationResult.overall_status === 'CRITICAL' ? (
                <ShieldAlert size={36} className="text-rose-400 shrink-0" />
              ) : evaluationResult.overall_status === 'CAUTION' ? (
                <AlertTriangle size={36} className="text-amber-400 shrink-0" />
              ) : (
                <ShieldCheck size={36} className="text-emerald-400 shrink-0" />
              )}
              <div>
                <div className="text-[11px] uppercase tracking-wider font-bold opacity-80">
                  {lang === 'vi' ? 'Kết Quả Thẩm Định Tổng Thể' : 'Overall Severity Assessment'}
                </div>
                <div className="text-lg font-black tracking-tight text-white">
                  {lang === 'vi'
                    ? evaluationResult.status_title_vn
                    : evaluationResult.status_title_en}
                </div>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <span className={`px-3 py-1 rounded-full text-xs font-black border ${evaluationResult.summary_color}`}>
                {evaluationResult.summary_badge}
              </span>
            </div>
          </div>

          {/* Technical Analysis Box */}
          <div className="bg-slate-900/90 p-5 rounded-xl border border-slate-800 space-y-3">
            <h3 className="text-xs font-bold uppercase tracking-wider text-indigo-400 flex items-center gap-2">
              <FileCheck2 size={16} />
              <span>
                {lang === 'vi'
                  ? 'Bản Báo Cáo Phân Tích Kỹ Thuật Chuyên Sâu (Engineering Technical Analysis)'
                  : 'Engineering Technical Analysis'}
              </span>
            </h3>

            <div className="text-sm font-bold text-white">
              {lang === 'vi'
                ? evaluationResult.technical_analysis.title_vn
                : evaluationResult.technical_analysis.title_en}
            </div>

            <p className="text-xs text-slate-300 leading-relaxed">
              {lang === 'vi'
                ? evaluationResult.technical_analysis.description_vn
                : evaluationResult.technical_analysis.description_en}
            </p>
          </div>

          {/* Detailed Item Checklist Table */}
          <div className="bg-slate-900/90 rounded-xl border border-slate-800 overflow-hidden">
            <div className="p-4 border-b border-slate-800 flex justify-between items-center">
              <h3 className="text-xs font-bold uppercase tracking-wider text-slate-300 flex items-center gap-2">
                <ClipboardCheck size={16} className="text-indigo-400" />
                <span>
                  {lang === 'vi'
                    ? 'Bảng Đánh Giá Chi Tiết Hạng Mục Nghiệm Thu (NETA ECS Section 8.2)'
                    : 'Detailed Commissioning Acceptance Table'}
                </span>
              </h3>
              <span className="text-[11px] text-slate-500">
                18 Hạng Mục Tiền / Trong / Sau Đóng Điện
              </span>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-950 text-slate-400 uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="py-3 px-3">Mã Mục</th>
                    <th className="py-3 px-3">Hạng Mục Kiểm Tra</th>
                    <th className="py-3 px-3">Tiêu Chuẩn NETA</th>
                    <th className="py-3 px-3">Giá Trị Đo Thực Tế</th>
                    <th className="py-3 px-3 text-center">Trạng Thái</th>
                    <th className="py-3 px-3">Kết Luận Kỹ Thuật TEV COE</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800">
                  {evaluationResult.evaluation_table.map((row) => (
                    <tr key={row.item_code} className="hover:bg-slate-800/40 transition">
                      <td className="py-2.5 px-3 font-mono font-bold text-indigo-300">
                        {row.item_code}
                      </td>
                      <td className="py-2.5 px-3 font-medium text-white">
                        <div>{lang === 'vi' ? row.item_name_vn : row.item_name_en}</div>
                      </td>
                      <td className="py-2.5 px-3 text-slate-400 font-mono text-[11px]">
                        {row.neta_standard_ref}
                      </td>
                      <td className="py-2.5 px-3 font-mono text-slate-200">
                        {row.measured_value}
                      </td>
                      <td className="py-2.5 px-3 text-center">
                        <span
                          className={`px-2 py-0.5 rounded text-[10px] font-black ${
                            row.status === 'PASS'
                              ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/40'
                              : row.status === 'FAIL'
                              ? 'bg-rose-500/20 text-rose-300 border border-rose-500/40'
                              : 'bg-amber-500/20 text-amber-300 border border-amber-500/40'
                          }`}
                        >
                          {row.status}
                        </span>
                      </td>
                      <td className="py-2.5 px-3 text-slate-300 text-[11px]">
                        {lang === 'vi' ? row.evaluation_vn : row.evaluation_en}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Actionable Protocol Steps */}
          <div className="bg-slate-900/90 p-5 rounded-xl border border-slate-800 space-y-4">
            <h3 className="text-xs font-bold uppercase tracking-wider text-indigo-400 flex items-center gap-2">
              <ShieldAlert size={16} />
              <span>
                {lang === 'vi'
                  ? 'Quy Trình Khắc Phục Khẩn Cấp & Kế Hoạch Nghiệm Thu (Actionable Protocol)'
                  : 'Actionable Engineering Protocol Recommendations'}
              </span>
            </h3>

            <div className="space-y-3">
              {evaluationResult.actionable_recommendations.steps.map((st) => (
                <div
                  key={st.step_number}
                  className="bg-slate-950/80 p-3.5 rounded-lg border border-slate-800/80"
                >
                  <div className="flex items-center gap-2 mb-1">
                    <span className="w-5 h-5 rounded-full bg-indigo-500/20 text-indigo-300 font-bold text-xs flex items-center justify-center border border-indigo-400/40">
                      {st.step_number}
                    </span>
                    <span className="text-xs font-bold text-white">
                      {lang === 'vi' ? st.title_vn : st.title_en}
                    </span>
                  </div>
                  <p className="text-xs text-slate-300 pl-7">
                    {lang === 'vi' ? st.details_vn : st.details_en}
                  </p>
                </div>
              ))}
            </div>
          </div>

          {/* Gemini AI Markdown Report (if generated) */}
          {evaluationResult.raw_markdown_report && (
            <div className="bg-slate-900/90 p-5 rounded-xl border border-indigo-500/40 space-y-3">
              <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                <span className="text-xs font-bold uppercase tracking-wider text-indigo-400 flex items-center gap-2">
                  <Sparkles size={16} />
                  <span>Báo Cáo Phân Tích Độc Lập Từ Gemini AI Specialist</span>
                </span>
                <span className="text-[10px] text-slate-400">Gemini 2.5 Flash</span>
              </div>
              <div className="prose prose-invert max-w-none text-xs text-slate-300 whitespace-pre-wrap leading-relaxed font-mono bg-slate-950 p-4 rounded-lg border border-slate-800 max-h-96 overflow-y-auto">
                {evaluationResult.raw_markdown_report}
              </div>
            </div>
          )}
        </div>
      )}

      {/* TAB: BATTERY INTERNAL RESISTANCE HISTORICAL TREND (IEEE 1188 / NETA MTS) */}
      {activeTab === 'trend' && (
        <UpsBatteryResistanceTrendChart
          systemTag={systemTag}
          batteryType={batteryType}
          batteryCapacityAh={batteryCapacityAh}
          currentDeviationPercent={batteryDeviation}
          currentWorstCell="Cell #18"
          lang={lang}
          onUpdateCurrentValue={(dev) => setBatteryDeviation(dev)}
        />
      )}

      {/* TAB 3: SYSTEM INSTRUCTION PROMPT FOR GOOGLE AI STUDIO */}
      {activeTab === 'prompt' && (
        <div className="bg-slate-900/90 p-5 rounded-xl border border-slate-800 space-y-4">
          <div className="flex items-center justify-between border-b border-slate-800 pb-3">
            <div>
              <h3 className="text-sm font-bold text-white flex items-center gap-2">
                <Sparkles size={16} className="text-amber-400" />
                <span>System Instruction Prompt Chuẩn Hóa Cho Google AI Studio</span>
              </h3>
              <p className="text-xs text-slate-400 mt-0.5">
                Sao chép System Instruction này vào Google AI Studio khi kiểm thử mô hình với dữ liệu chạy thử thực địa UPS.
              </p>
            </div>

            <button
              onClick={handleCopyPrompt}
              className="px-3.5 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold flex items-center gap-1.5 shadow transition"
            >
              {isCopiedPrompt ? <Check size={14} /> : <Copy size={14} />}
              <span>{isCopiedPrompt ? 'Đã Sao Chép!' : 'Sao Chép Prompt'}</span>
            </button>
          </div>

          <pre className="p-4 bg-slate-950 rounded-lg text-xs font-mono text-slate-300 overflow-x-auto whitespace-pre-wrap border border-slate-800 max-h-[500px]">
            {ECS_UPS_SYSTEM_INSTRUCTION}
          </pre>
        </div>
      )}

      {/* TAB 4: SAMPLE INPUT JSON PAYLOAD */}
      {activeTab === 'json' && (
        <div className="bg-slate-900/90 p-5 rounded-xl border border-slate-800 space-y-4">
          <div className="flex items-center justify-between border-b border-slate-800 pb-3">
            <div>
              <h3 className="text-sm font-bold text-white flex items-center gap-2">
                <Cpu size={16} className="text-indigo-400" />
                <span>Mẫu JSON Đầu Vào Nghiệm Thu UPS Chuẩn NETA ECS-2024 Mục 8.2</span>
              </h3>
              <p className="text-xs text-slate-400 mt-0.5">
                Cấu trúc JSON đầu vào chuẩn hóa tích hợp API và xuất báo cáo lưu trữ CMMS.
              </p>
            </div>

            <button
              onClick={handleCopyJson}
              className="px-3.5 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold flex items-center gap-1.5 shadow transition"
            >
              {isCopiedJson ? <Check size={14} /> : <Copy size={14} />}
              <span>{isCopiedJson ? 'Đã Sao Chép JSON!' : 'Sao Chép JSON'}</span>
            </button>
          </div>

          <pre className="p-4 bg-slate-950 rounded-lg text-xs font-mono text-emerald-400 overflow-x-auto whitespace-pre-wrap border border-slate-800 max-h-[500px]">
            {JSON.stringify(currentPayload, null, 2)}
          </pre>
        </div>
      )}
    </div>
  );
};
