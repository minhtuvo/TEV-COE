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
  Sun,
  Radio,
  Sliders,
  Cable,
  TrendingDown
} from 'lucide-react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  NetaEcsSolarPvInputPayload,
  NetaEcsSolarPvEvaluationResult,
  performDeterministicEcsSolarPvAnalysis,
  ECS_SOLAR_PV_SYSTEM_INSTRUCTION
} from '../../server/netaEcsSolarPvAnalyzer';

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
  string0103DryIr: number;
  string04DryIr: number;
  isPolarityPass: boolean;
  string0103Voc: number;
  string04Voc: number;
  pmaxStcString04: number;
  groundingR: number;
  isTorquePass: boolean;
  isPlanApproved: boolean;
  antiIslandingTripTime: number;
  isThermalPass: boolean;
  thermalDeltaT?: number;
  isNercPass: boolean;
  description: string;
}

const PRESET_SCENARIOS: PresetScenario[] = [
  {
    id: 'critical-ground-fault',
    name: '🔴 NGUY HẠI ĐÓNG ĐIỆN: Sự Cố Chạm Đất DC & Sụt Áp Voc String 04 (Mẫu Chuẩn AI Studio)',
    nameEn: 'CRITICAL HAZARD: DC Ground Fault & Voc Drop on String 04 (Benchmark Prompt)',
    badge: 'CRITICAL / NGUY HIỂM',
    badgeColor: 'bg-rose-500/20 text-rose-300 border-rose-500/40',
    systemTag: 'BLOCK-01-1500V-PV-SYSTEM',
    voltageRatingV: 1500,
    string0103DryIr: 45.0,
    string04DryIr: 0.85, // < 5.0 MΩ (Table 100.1)
    isPolarityPass: true,
    string0103Voc: 1485,
    string04Voc: 1390, // Deviation = 6.4% (> 5%)
    pmaxStcString04: 72.0, // < 95%
    groundingR: 0.65,
    isTorquePass: true,
    isPlanApproved: true,
    antiIslandingTripTime: 0.15,
    isThermalPass: true,
    isNercPass: true,
    description: 'String 04 đo cách điện chỉ 0.85 MΩ (< 5.0 MΩ), Voc lệch 6.4%, I-V méo nấc Pmax = 72%. Cần cô lập rút cầu chì String 04.'
  },
  {
    id: 'all-pass-utility',
    name: '🟢 ĐẠT CHUẨN 100%: Nhà Máy Điện Mặt Trời Utility 1500V DC (Sẵn Sàng COD)',
    nameEn: '100% PASS: 1500V DC Utility PV Solar Farm (COD Ready)',
    badge: 'PASS / ĐẠT CHUẨN',
    badgeColor: 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40',
    systemTag: 'SOLAR-INVERTER-SKID-02',
    voltageRatingV: 1500,
    string0103DryIr: 52.0,
    string04DryIr: 48.5,
    isPolarityPass: true,
    string0103Voc: 1490,
    string04Voc: 1478, // Deviation = 0.8%
    pmaxStcString04: 98.4,
    groundingR: 0.45,
    isTorquePass: true,
    isPlanApproved: true,
    antiIslandingTripTime: 0.14,
    isThermalPass: true,
    isNercPass: true,
    description: 'Cách điện > 45 MΩ, Voc lệch 0.8%, Pmax 98.4%, chống đảo lưới 0.14s, không có điểm nóng, đạt chuẩn NERC MOD-025-2.'
  },
  {
    id: 'critical-reverse-polarity',
    name: '🔴 NGUY HẠI: Ngược Cực Tính DC & Hỏng Bảo Vệ Chống Đảo Lưới IEEE 1547',
    nameEn: 'CRITICAL: DC Reverse Polarity & Anti-Islanding Protection Trip Failure',
    badge: 'CRITICAL / NGUY HIỂM',
    badgeColor: 'bg-rose-500/20 text-rose-300 border-rose-500/40',
    systemTag: 'ROOFTOP-PV-1000V-ARRAY',
    voltageRatingV: 1000,
    string0103DryIr: 12.0,
    string04DryIr: 15.0,
    isPolarityPass: false, // Reverse polarity!
    string0103Voc: 980,
    string04Voc: 975,
    pmaxStcString04: 96.0,
    groundingR: 1.2,
    isTorquePass: true,
    isPlanApproved: false,
    antiIslandingTripTime: 2.8, // > 2.0s fail!
    isThermalPass: true,
    isNercPass: true,
    description: 'Ngược cực tính DC (+/-) tại ngõ vào tủ gom và biến tần không ngắt phát điện khi mất lưới (> 2.8s), vi phạm IEEE 1547.'
  },
  {
    id: 'caution-thermal-hotspot',
    name: '🟡 CẢNH BÁO: Điểm Nóng Hồng Ngoại MC4 & Bu-lông Chưa Đủ Lực Siết',
    nameEn: 'CAUTION: Combiner MC4 Hotspot Anomaly & Deficient Bolted Torque',
    badge: 'CAUTION / CẢNH BÁO',
    badgeColor: 'bg-amber-500/20 text-amber-300 border-amber-500/40',
    systemTag: 'COMBINER-BOX-C14-1500V',
    voltageRatingV: 1500,
    string0103DryIr: 38.0,
    string04DryIr: 35.0,
    isPolarityPass: true,
    string0103Voc: 1480,
    string04Voc: 1465, // Deviation = 1.0%
    pmaxStcString04: 96.5,
    groundingR: 2.1,
    isTorquePass: false, // Deficient torque
    isPlanApproved: true,
    antiIslandingTripTime: 0.16,
    isThermalPass: false, // Hotspot Table 100.18
    thermalDeltaT: 18.5,
    isNercPass: true,
    description: 'Phát hiện điểm nóng chênh lệch nhiệt độ ΔT = 18.5°C tại cọc đấu dây tủ gom theo Bảng 100.18 và bu-lông chưa đạt lực siết.'
  }
];

export const NetaEcsSolarPvChecklist: React.FC<Props> = ({
  onSaveReport,
  onNavigateToReports,
  initialLanguage = 'vi'
}) => {
  const [lang, setLang] = useState<'vi' | 'en'>(initialLanguage);
  const isVi = lang === 'vi';

  // Header metadata
  const [projectName, setProjectName] = useState('Nha May Dien Mat Troi TEV Solar Farm');
  const [systemTag, setSystemTag] = useState('BLOCK-01-1500V-PV-SYSTEM');
  const [voltageRatingV, setVoltageRatingV] = useState<number>(1500);
  const [fseName, setFseName] = useState('Nguyen Van K');
  const [commissioningDate, setCommissioningDate] = useState('2026-10-17');

  // Phase A: Pre-Energization (8.1.A)
  const [isAtsReviewed, setIsAtsReviewed] = useState(true);
  const [dcPolarityStatus, setDcPolarityStatus] = useState<'pass' | 'fail'>('pass');
  const [string0103DryIr, setString0103DryIr] = useState<number>(45.0);
  const [string04DryIr, setString04DryIr] = useState<number>(0.85);
  const [string0103Voc, setString0103Voc] = useState<number>(1485);
  const [string04Voc, setString04Voc] = useState<number>(1390);
  const [pmaxStcString0103, setPmaxStcString0103] = useState<number>(98.1);
  const [pmaxStcString04, setPmaxStcString04] = useState<number>(72.0);
  const [groundingR, setGroundingR] = useState<number>(0.65);
  const [isTorquePass, setIsTorquePass] = useState(true);
  const [isInverterSettingsPass, setIsInverterSettingsPass] = useState(true);
  const [isPlanApproved, setIsPlanApproved] = useState(true);

  // Phase B: Energization (8.1.B)
  const [isSafetyRestricted, setIsSafetyRestricted] = useState(true);
  const [isAcPhasingPass, setIsAcPhasingPass] = useState(true);
  const [antiIslandingTripTime, setAntiIslandingTripTime] = useState<number>(0.15);
  const [isGridCodePass, setIsGridCodePass] = useState(true);

  // Phase C: Post-Energization (8.1.C)
  const [isDasBaselinePass, setIsDasBaselinePass] = useState(true);
  const [isThermalPass, setIsThermalPass] = useState(true);
  const [thermalDeltaT, setThermalDeltaT] = useState<number>(4.2);
  const [isNercPass, setIsNercPass] = useState(true);

  // Evaluation & Results State
  const [evaluationResult, setEvaluationResult] = useState<NetaEcsSolarPvEvaluationResult | null>(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [activeResultTab, setActiveResultTab] = useState<'summary' | 'table' | 'recommendations' | 'ai-markdown'>('summary');

  // Modals
  const [showSystemPromptModal, setShowSystemPromptModal] = useState(false);
  const [showJsonPromptModal, setShowJsonPromptModal] = useState(false);
  const [copiedPrompt, setCopiedPrompt] = useState(false);
  const [copiedJson, setCopiedJson] = useState(false);

  // Live calculations
  const vocDeviationPercent = useMemo(() => {
    if (!string0103Voc || !string04Voc) return 0;
    const diff = Math.abs(string0103Voc - string04Voc);
    const avg = (string0103Voc + string04Voc) / 2;
    return parseFloat(((diff / avg) * 100).toFixed(2));
  }, [string0103Voc, string04Voc]);

  const minRequiredDryIr = voltageRatingV >= 1500 ? 5.0 : 1.0;

  // Build Payload
  const currentPayload: NetaEcsSolarPvInputPayload = useMemo(() => {
    return {
      tev_coe_header: {
        module: 'PV-COMMISSIONING-ECS-8.1',
        project_name: projectName,
        system_tag: systemTag,
        system_dc_voltage_rating_V: voltageRatingV,
        fse_name: fseName,
        commissioning_date: commissioningDate
      },
      pre_energization_8_1_a: {
        acceptance_test_reports_reviewed_ats_7_29: isAtsReviewed,
        dc_polarity_and_blocking_diodes: dcPolarityStatus === 'pass' ? 'Pass (100% correct DC polarity verified)' : 'Fail (DC reverse polarity hazard)',
        pv_string_dry_insulation_resistance_megohms: {
          voltage_rating_v: voltageRatingV,
          string_01_to_03_avg: string0103DryIr,
          string_04: string04DryIr,
          status: string04DryIr >= minRequiredDryIr ? 'Pass' : `Fail (String 04 dry IR < ${minRequiredDryIr} MΩ requirement per NETA Table 100.1)`
        },
        string_voc_measurements_volts: {
          string_01_to_03_avg_voc: string0103Voc,
          string_04_voc: string04Voc,
          max_voc_deviation_percent: vocDeviationPercent,
          status: vocDeviationPercent <= 5.0 ? 'Pass' : `Fail (String 04 Voc deviation ${vocDeviationPercent}% > 5% allowable limit)`
        },
        iv_curve_tracing_performance: {
          string_01_to_03_stc_pmax_percent: pmaxStcString0103,
          string_04_stc_pmax_percent: pmaxStcString04,
          status: pmaxStcString04 >= 95.0 ? 'Pass' : 'Fail (String 04 I-V curve distorted with step-notch)'
        },
        grounding_resistance_ohms: groundingR,
        bolted_connection_torque_table_100_12: isTorquePass ? 'Pass (Calibrated torque wrench verified per Table 100.12)' : 'Fail (Insufficient torque)',
        inverter_gateway_das_settings: isInverterSettingsPass ? 'Pass (Engineered setting files loaded)' : 'Fail (Configuration discrepancy)',
        written_energization_plan_approved: isPlanApproved
      },
      energization_8_1_b: {
        safety_access_restricted: isSafetyRestricted,
        ac_phase_sequence_check: isAcPhasingPass ? 'Pass (Correct AC sequence)' : 'Fail (Reverse AC phase sequence)',
        anti_islanding_protection_test_ieee_1547: {
          trip_time_seconds: antiIslandingTripTime,
          status: antiIslandingTripTime <= 2.0 ? `Pass (Inverter trips within ${antiIslandingTripTime}s upon AC breaker open)` : 'Fail (Anti-islanding timeout)'
        },
        grid_code_and_power_factor_tuning: isGridCodePass ? 'Pass (Cos phi = 0.99 leading/lagging verified)' : 'Fail (Grid code non-compliance)'
      },
      post_energization_8_1_c: {
        das_weather_sensors_baseline: isDasBaselinePass ? 'Pass (Pyranometer and RTD calibrated)' : 'Fail (Sensors calibration error)',
        thermographic_survey_table_100_18: {
          max_delta_t_celsius: thermalDeltaT,
          status: isThermalPass ? 'Pass (Normal temperature distribution)' : 'Investigate (Hotspot detected per Table 100.18)'
        },
        nerc_mod_025_2_reactive_power_verification: isNercPass ? 'Pass (Reactive power capability curve verified)' : 'Fail (Insufficient Q-capability)'
      }
    };
  }, [
    projectName, systemTag, voltageRatingV, fseName, commissioningDate,
    isAtsReviewed, dcPolarityStatus, string0103DryIr, string04DryIr, minRequiredDryIr,
    string0103Voc, string04Voc, vocDeviationPercent, pmaxStcString0103, pmaxStcString04,
    groundingR, isTorquePass, isInverterSettingsPass, isPlanApproved,
    isSafetyRestricted, isAcPhasingPass, antiIslandingTripTime, isGridCodePass,
    isDasBaselinePass, isThermalPass, thermalDeltaT, isNercPass
  ]);

  // Handle Preset Selection
  const applyPresetScenario = (scenario: PresetScenario) => {
    setSystemTag(scenario.systemTag);
    setVoltageRatingV(scenario.voltageRatingV);
    setString0103DryIr(scenario.string0103DryIr);
    setString04DryIr(scenario.string04DryIr);
    setDcPolarityStatus(scenario.isPolarityPass ? 'pass' : 'fail');
    setString0103Voc(scenario.string0103Voc);
    setString04Voc(scenario.string04Voc);
    setPmaxStcString04(scenario.pmaxStcString04);
    setGroundingR(scenario.groundingR);
    setIsTorquePass(scenario.isTorquePass);
    setIsPlanApproved(scenario.isPlanApproved);
    setAntiIslandingTripTime(scenario.antiIslandingTripTime);
    setIsThermalPass(scenario.isThermalPass);
    if (scenario.thermalDeltaT !== undefined) setThermalDeltaT(scenario.thermalDeltaT);
    setIsNercPass(scenario.isNercPass);

    // Auto run analysis
    const updatedPayload: NetaEcsSolarPvInputPayload = {
      ...currentPayload,
      tev_coe_header: {
        ...currentPayload.tev_coe_header,
        system_tag: scenario.systemTag,
        system_dc_voltage_rating_V: scenario.voltageRatingV
      },
      pre_energization_8_1_a: {
        ...currentPayload.pre_energization_8_1_a,
        dc_polarity_and_blocking_diodes: scenario.isPolarityPass ? 'Pass (100% correct DC polarity verified)' : 'Fail (DC reverse polarity hazard)',
        pv_string_dry_insulation_resistance_megohms: {
          voltage_rating_v: scenario.voltageRatingV,
          string_01_to_03_avg: scenario.string0103DryIr,
          string_04: scenario.string04DryIr,
          status: scenario.string04DryIr >= (scenario.voltageRatingV >= 1500 ? 5.0 : 1.0) ? 'Pass' : 'Fail'
        },
        string_voc_measurements_volts: {
          string_01_to_03_avg_voc: scenario.string0103Voc,
          string_04_voc: scenario.string04Voc,
          max_voc_deviation_percent: parseFloat(((Math.abs(scenario.string0103Voc - scenario.string04Voc) / ((scenario.string0103Voc + scenario.string04Voc) / 2)) * 100).toFixed(2))
        },
        iv_curve_tracing_performance: {
          string_01_to_03_stc_pmax_percent: 98.1,
          string_04_stc_pmax_percent: scenario.pmaxStcString04
        },
        grounding_resistance_ohms: scenario.groundingR,
        bolted_connection_torque_table_100_12: scenario.isTorquePass ? 'Pass' : 'Fail',
        written_energization_plan_approved: scenario.isPlanApproved
      },
      energization_8_1_b: {
        ...currentPayload.energization_8_1_b,
        anti_islanding_protection_test_ieee_1547: {
          trip_time_seconds: scenario.antiIslandingTripTime,
          status: scenario.antiIslandingTripTime <= 2.0 ? 'Pass' : 'Fail'
        }
      },
      post_energization_8_1_c: {
        ...currentPayload.post_energization_8_1_c,
        thermographic_survey_table_100_18: {
          max_delta_t_celsius: scenario.thermalDeltaT ?? 4.0,
          status: scenario.isThermalPass ? 'Pass' : 'Investigate'
        },
        nerc_mod_025_2_reactive_power_verification: scenario.isNercPass ? 'Pass' : 'Fail'
      }
    };

    const res = performDeterministicEcsSolarPvAnalysis(updatedPayload);
    setEvaluationResult(res);
  };

  // Run Deterministic Analysis
  const handleRunDeterministic = () => {
    setIsAnalyzing(true);
    setTimeout(() => {
      const res = performDeterministicEcsSolarPvAnalysis(currentPayload);
      setEvaluationResult(res);
      setIsAnalyzing(false);
    }, 250);
  };

  // Run Gemini AI Analysis
  const handleRunAiAnalysis = async () => {
    setIsAnalyzing(true);
    try {
      const res = await fetch('/api/field-service/neta-ecs-solar-pv-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(currentPayload)
      });
      const data = await res.json();
      if (data.success && data.data) {
        setEvaluationResult(data.data);
      } else {
        // Fallback to deterministic
        setEvaluationResult(performDeterministicEcsSolarPvAnalysis(currentPayload));
      }
    } catch (e) {
      console.error('Failed to run AI analysis, falling back to deterministic:', e);
      setEvaluationResult(performDeterministicEcsSolarPvAnalysis(currentPayload));
    } finally {
      setIsAnalyzing(false);
    }
  };

  // Export PDF Report (A4 Landscape or Portrait)
  const handleExportPdf = () => {
    const doc = new jsPDF('p', 'mm', 'a4');
    const result = evaluationResult || performDeterministicEcsSolarPvAnalysis(currentPayload);

    // Header Background
    doc.setFillColor(15, 23, 42); // slate-900
    doc.rect(0, 0, 210, 36, 'F');

    // Title
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(15);
    doc.setFont('helvetica', 'bold');
    doc.text('TEV COE - BIEN BAN CHAY THU & NGHIEM THU DIEN MAT TROI', 14, 14);

    doc.setFontSize(9);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(203, 213, 225);
    doc.text('ANSI/NETA ECS-2024 Sec 8.1 | NETA ATS-2025 Sec 7.29 | IEC 62446-1 | IEEE 1547 | NERC MOD-025-2', 14, 21);

    doc.setTextColor(245, 158, 11); // Amber
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
    doc.text('MA KHOI / HE THONG:', 18, 52);
    doc.setFont('helvetica', 'normal');
    doc.text(`${systemTag} (${voltageRatingV}V DC)`, 58, 52);

    doc.setFont('helvetica', 'bold');
    doc.text('KY SU FSE:', 18, 58);
    doc.setFont('helvetica', 'normal');
    doc.text(`${fseName} | Ngay: ${commissioningDate}`, 58, 58);

    doc.setFont('helvetica', 'bold');
    doc.text('VOC DEVIATION:', 125, 46);
    doc.setFont('helvetica', 'normal');
    doc.text(`${vocDeviationPercent}% (Max <= 5%)`, 160, 46);

    doc.setFont('helvetica', 'bold');
    doc.text('DRY IR STRING 04:', 125, 52);
    doc.setFont('helvetica', 'normal');
    doc.text(`${string04DryIr} M (Min >= ${minRequiredDryIr} M)`, 160, 52);

    doc.setFont('helvetica', 'bold');
    doc.text('ANTI-ISLANDING:', 125, 58);
    doc.setFont('helvetica', 'normal');
    doc.text(`${antiIslandingTripTime}s (IEEE 1547 <= 2.0s)`, 160, 58);

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

    doc.save(`TEV_COE_NETA_ECS_8_1_SolarPV_${systemTag}_${commissioningDate}.pdf`);
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
      <div className="bg-gradient-to-r from-slate-900 via-amber-950 to-slate-900 rounded-2xl p-5 md:p-6 text-white shadow-xl border border-amber-500/30">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="space-y-1.5">
            <div className="flex items-center gap-2.5 flex-wrap">
              <div className="w-10 h-10 rounded-xl bg-amber-500/20 text-amber-400 border border-amber-400/30 flex items-center justify-center font-black">
                <Sun size={24} />
              </div>
              <div>
                <div className="flex items-center gap-2 flex-wrap">
                  <h2 className="text-xl font-black tracking-tight">
                    {isVi
                      ? 'Nghiệm Thu & Chạy Thử Hệ Thống Điện Mặt Trời (Solar PV)'
                      : 'Source Specific Systems Commissioning: Solar PV Procedures'}
                  </h2>
                  <span className="px-2 py-0.5 rounded-full text-[10px] font-black uppercase tracking-wider bg-amber-500 text-slate-950 shadow-sm">
                    NETA ECS-2024 Mục 8.1
                  </span>
                </div>
                <p className="text-xs text-amber-200/80">
                  {isVi
                    ? 'Chuẩn hóa theo ANSI/NETA ECS-2024 Mục 8.1, NETA ATS 7.29, IEC 62446-1, IEEE 1547 & NERC MOD-025-2'
                    : 'Standardized per ANSI/NETA ECS-2024 Sec 8.1, NETA ATS 7.29, IEC 62446-1, IEEE 1547 & NERC MOD-025-2'}
                </p>
              </div>
            </div>
          </div>

          {/* Quick Actions & AI Studio Buttons */}
          <div className="flex items-center gap-2 flex-wrap">
            <button
              type="button"
              onClick={() => setShowSystemPromptModal(true)}
              className="flex items-center gap-1.5 px-3 py-2 rounded-xl text-xs font-bold bg-amber-500/20 text-amber-300 hover:bg-amber-500/30 border border-amber-400/40 transition-all shadow-xs"
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
              className="flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-xs font-black bg-gradient-to-r from-amber-500 to-amber-400 text-slate-950 hover:from-amber-400 hover:to-amber-300 transition-all shadow-md shadow-amber-500/20"
            >
              <Download size={14} />
              <span>{isVi ? 'Xuất PDF A4' : 'Export PDF'}</span>
            </button>
          </div>
        </div>

        {/* 2. PRESET SCENARIO SELECTOR */}
        <div className="mt-5 pt-4 border-t border-amber-500/20">
          <div className="text-xs font-bold text-amber-300 mb-2 flex items-center gap-1.5">
            <Sparkles size={14} />
            <span>{isVi ? 'Kịch bản thực địa mẫu (1-Click Presets):' : 'Sample Field Scenarios:'}</span>
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2">
            {PRESET_SCENARIOS.map((scenario) => (
              <button
                key={scenario.id}
                type="button"
                onClick={() => applyPresetScenario(scenario)}
                className="text-left p-2.5 rounded-xl bg-slate-950/70 hover:bg-slate-900 border border-slate-800 hover:border-amber-400/50 transition-all group"
              >
                <div className="flex items-center justify-between mb-1">
                  <span className={`text-[9px] font-black px-1.5 py-0.5 rounded border ${scenario.badgeColor}`}>
                    {scenario.badge}
                  </span>
                  <span className="text-[10px] font-mono text-slate-400">{scenario.voltageRatingV}V DC</span>
                </div>
                <div className="text-xs font-bold text-slate-200 group-hover:text-amber-300 line-clamp-1">
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

      {/* 3. INPUT FORMS (3 GIAI ĐOẠN ĐÓNG ĐIỆN) */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* PHASE 1: PRE-ENERGIZATION (NETA ECS 8.1.A) */}
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 space-y-4">
          <div className="flex items-center justify-between border-b border-slate-100 pb-3">
            <div className="flex items-center gap-2">
              <div className="w-7 h-7 rounded-lg bg-amber-50 text-amber-700 flex items-center justify-center font-bold text-xs">
                A
              </div>
              <h3 className="font-black text-sm text-slate-900">
                {isVi ? 'GIAI ĐOẠN 1: TIỀN ĐÓNG ĐIỆN' : 'PHASE 1: PRE-ENERGIZATION'}
              </h3>
            </div>
            <span className="text-[10px] font-mono font-bold px-2 py-0.5 rounded bg-slate-100 text-slate-600">
              NETA ECS 8.1.A
            </span>
          </div>

          {/* System DC Voltage Rating */}
          <div className="space-y-1">
            <label className="text-xs font-bold text-slate-700">
              {isVi ? 'Cấp điện áp DC hệ thống:' : 'System DC Voltage Rating:'}
            </label>
            <div className="grid grid-cols-2 gap-2">
              <button
                type="button"
                onClick={() => setVoltageRatingV(1000)}
                className={`py-1.5 text-xs font-bold rounded-lg border transition-all ${
                  voltageRatingV === 1000
                    ? 'bg-amber-500 text-slate-950 border-amber-400 font-black shadow-xs'
                    : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                }`}
              >
                1000V DC (Min ≥ 1.0 MΩ)
              </button>
              <button
                type="button"
                onClick={() => setVoltageRatingV(1500)}
                className={`py-1.5 text-xs font-bold rounded-lg border transition-all ${
                  voltageRatingV === 1500
                    ? 'bg-amber-500 text-slate-950 border-amber-400 font-black shadow-xs'
                    : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                }`}
              >
                1500V DC (Min ≥ 5.0 MΩ)
              </button>
            </div>
          </div>

          {/* Acceptance Testing & Polarity */}
          <div className="grid grid-cols-2 gap-2 pt-1">
            <div>
              <label className="text-[11px] font-semibold text-slate-600 block mb-1">
                {isVi ? 'Báo cáo ATS 7.29:' : 'ATS 7.29 Reports:'}
              </label>
              <button
                type="button"
                onClick={() => setIsAtsReviewed(!isAtsReviewed)}
                className={`w-full py-1.5 px-2 text-xs font-bold rounded-lg border flex items-center justify-center gap-1.5 ${
                  isAtsReviewed
                    ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                    : 'bg-rose-50 text-rose-800 border-rose-300'
                }`}
              >
                {isAtsReviewed ? <CheckCircle2 size={13} /> : <AlertTriangle size={13} />}
                <span>{isAtsReviewed ? (isVi ? 'Đã duyệt' : 'Reviewed') : (isVi ? 'Thiếu' : 'Missing')}</span>
              </button>
            </div>

            <div>
              <label className="text-[11px] font-semibold text-slate-600 block mb-1">
                {isVi ? 'Thử cực tính DC:' : 'DC Polarity:'}
              </label>
              <button
                type="button"
                onClick={() => setDcPolarityStatus(dcPolarityStatus === 'pass' ? 'fail' : 'pass')}
                className={`w-full py-1.5 px-2 text-xs font-bold rounded-lg border flex items-center justify-center gap-1.5 ${
                  dcPolarityStatus === 'pass'
                    ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                    : 'bg-rose-50 text-rose-800 border-rose-300 font-black'
                }`}
              >
                {dcPolarityStatus === 'pass' ? <CheckCircle2 size={13} /> : <AlertOctagon size={13} />}
                <span>{dcPolarityStatus === 'pass' ? (isVi ? 'Chuẩn (+/-)' : 'Correct') : (isVi ? 'NGƯỢC CỰC' : 'REVERSED')}</span>
              </button>
            </div>
          </div>

          {/* Dry Insulation Resistance Table 100.1 */}
          <div className="space-y-1.5 bg-slate-50 p-2.5 rounded-xl border border-slate-200">
            <div className="flex items-center justify-between text-xs font-bold text-slate-800">
              <span>{isVi ? 'Cách điện Khô String (MΩ):' : 'String Dry IR (MΩ):'}</span>
              <span className="text-[10px] text-amber-700">Yêu cầu ≥ {minRequiredDryIr} MΩ</span>
            </div>
            <div className="grid grid-cols-2 gap-2">
              <div>
                <span className="text-[10px] text-slate-500 block">String 01-03 TB:</span>
                <input
                  type="number"
                  value={string0103DryIr}
                  onChange={(e) => setString0103DryIr(parseFloat(e.target.value) || 0)}
                  className="w-full text-xs font-bold px-2 py-1 bg-white border border-slate-300 rounded-lg text-slate-800"
                />
              </div>
              <div>
                <span className="text-[10px] text-slate-500 block">String 04:</span>
                <input
                  type="number"
                  step="0.05"
                  value={string04DryIr}
                  onChange={(e) => setString04DryIr(parseFloat(e.target.value) || 0)}
                  className={`w-full text-xs font-black px-2 py-1 bg-white border rounded-lg ${
                    string04DryIr < minRequiredDryIr ? 'border-rose-400 text-rose-600 bg-rose-50/50' : 'border-slate-300 text-slate-800'
                  }`}
                />
              </div>
            </div>
          </div>

          {/* Voc Open-Circuit Voltage & I-V Curve */}
          <div className="space-y-1.5 bg-slate-50 p-2.5 rounded-xl border border-slate-200">
            <div className="flex items-center justify-between text-xs font-bold text-slate-800">
              <span>{isVi ? 'Đo Voc (V) & Độ lệch %:' : 'Voc Measurements:'}</span>
              <span className={`text-[10px] font-black px-1.5 rounded ${
                vocDeviationPercent > 5.0 ? 'bg-rose-100 text-rose-700' : 'bg-emerald-100 text-emerald-700'
              }`}>
                Δ = {vocDeviationPercent}% (Max ≤ 5%)
              </span>
            </div>
            <div className="grid grid-cols-2 gap-2">
              <div>
                <span className="text-[10px] text-slate-500 block">String 01-03 (V):</span>
                <input
                  type="number"
                  value={string0103Voc}
                  onChange={(e) => setString0103Voc(parseFloat(e.target.value) || 0)}
                  className="w-full text-xs font-bold px-2 py-1 bg-white border border-slate-300 rounded-lg text-slate-800"
                />
              </div>
              <div>
                <span className="text-[10px] text-slate-500 block">String 04 (V):</span>
                <input
                  type="number"
                  value={string04Voc}
                  onChange={(e) => setString04Voc(parseFloat(e.target.value) || 0)}
                  className="w-full text-xs font-bold px-2 py-1 bg-white border border-slate-300 rounded-lg text-slate-800"
                />
              </div>
            </div>

            <div className="pt-1 flex items-center justify-between">
              <span className="text-[11px] text-slate-600 font-semibold">{isVi ? 'Pmax STC String 04 (%):' : 'String 04 STC Pmax:'}</span>
              <div className="flex items-center gap-1.5">
                <input
                  type="number"
                  step="0.5"
                  value={pmaxStcString04}
                  onChange={(e) => setPmaxStcString04(parseFloat(e.target.value) || 0)}
                  className="w-16 text-xs font-bold px-1.5 py-0.5 text-right bg-white border border-slate-300 rounded"
                />
                <span className="text-[10px] text-slate-500">(≥95%)</span>
              </div>
            </div>
          </div>

          {/* Grounding & Torque */}
          <div className="grid grid-cols-2 gap-2">
            <div>
              <span className="text-[11px] text-slate-600 font-semibold block mb-0.5">{isVi ? 'Tiếp địa trạm (Ω):' : 'Grounding (Ω):'}</span>
              <input
                type="number"
                step="0.1"
                value={groundingR}
                onChange={(e) => setGroundingR(parseFloat(e.target.value) || 0)}
                className={`w-full text-xs font-bold px-2 py-1 bg-white border rounded-lg ${
                  groundingR > 5.0 ? 'border-rose-400 text-rose-600' : 'border-slate-300 text-slate-800'
                }`}
              />
            </div>
            <div>
              <span className="text-[11px] text-slate-600 font-semibold block mb-0.5">{isVi ? 'Lực siết bu-lông:' : 'Bolt Torque:'}</span>
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
          </div>

          {/* Inverter Settings & Plan */}
          <div className="grid grid-cols-2 gap-2 pt-1">
            <button
              type="button"
              onClick={() => setIsInverterSettingsPass(!isInverterSettingsPass)}
              className={`py-1.5 px-2 text-[11px] font-bold rounded-lg border text-center ${
                isInverterSettingsPass ? 'bg-slate-50 text-slate-800 border-slate-300' : 'bg-rose-50 text-rose-800 border-rose-300'
              }`}
            >
              {isVi ? 'Cài đặt Inverter & DAS' : 'Inverter & DAS'}
            </button>
            <button
              type="button"
              onClick={() => setIsPlanApproved(!isPlanApproved)}
              className={`py-1.5 px-2 text-[11px] font-bold rounded-lg border text-center ${
                isPlanApproved ? 'bg-emerald-50 text-emerald-800 border-emerald-300' : 'bg-rose-50 text-rose-800 border-rose-300'
              }`}
            >
              {isPlanApproved ? (isVi ? 'Kế hoạch đã duyệt' : 'Plan Approved') : (isVi ? 'Chưa duyệt KH' : 'Plan Missing')}
            </button>
          </div>
        </div>

        {/* PHASE 2: ENERGIZATION (NETA ECS 8.1.B) */}
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 space-y-4">
          <div className="flex items-center justify-between border-b border-slate-100 pb-3">
            <div className="flex items-center gap-2">
              <div className="w-7 h-7 rounded-lg bg-teal-50 text-teal-700 flex items-center justify-center font-bold text-xs">
                B
              </div>
              <h3 className="font-black text-sm text-slate-900">
                {isVi ? 'GIAI ĐOẠN 2: TRONG ĐÓNG ĐIỆN' : 'PHASE 2: ENERGIZATION'}
              </h3>
            </div>
            <span className="text-[10px] font-mono font-bold px-2 py-0.5 rounded bg-slate-100 text-slate-600">
              NETA ECS 8.1.B
            </span>
          </div>

          {/* Safety & Access Control */}
          <div className="space-y-1">
            <label className="text-xs font-bold text-slate-700">
              {isVi ? 'B1. Khoanh vùng an toàn & PPE hồ quang:' : 'B1. Access Control & Arc Flash PPE:'}
            </label>
            <button
              type="button"
              onClick={() => setIsSafetyRestricted(!isSafetyRestricted)}
              className={`w-full py-2 px-3 text-xs font-bold rounded-xl border flex items-center justify-between ${
                isSafetyRestricted
                  ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                  : 'bg-rose-50 text-rose-800 border-rose-300'
              }`}
            >
              <span>{isSafetyRestricted ? (isVi ? 'Đã rào chắn khoanh vùng an toàn 100%' : 'Perimeter Secured') : (isVi ? 'Chưa rào chắn an toàn!' : 'Perimeter Unsecured')}</span>
              {isSafetyRestricted ? <CheckCircle2 size={15} /> : <AlertOctagon size={15} />}
            </button>
          </div>

          {/* AC Phasing Check */}
          <div className="space-y-1">
            <label className="text-xs font-bold text-slate-700">
              {isVi ? 'B2. Thứ tự pha AC & Đồng pha đầu ra:' : 'B2. AC Phase Rotation & Phasing:'}
            </label>
            <button
              type="button"
              onClick={() => setIsAcPhasingPass(!isAcPhasingPass)}
              className={`w-full py-2 px-3 text-xs font-bold rounded-xl border flex items-center justify-between ${
                isAcPhasingPass
                  ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                  : 'bg-rose-50 text-rose-800 border-rose-300'
              }`}
            >
              <span>{isAcPhasingPass ? (isVi ? 'Trùng thứ tự pha lưới (A-B-C thuận)' : 'Correct Phase Sequence') : (isVi ? 'NGƯỢC THỨ TỰ PHA AC' : 'REVERSE AC PHASING')}</span>
              {isAcPhasingPass ? <CheckCircle2 size={15} /> : <AlertTriangle size={15} />}
            </button>
          </div>

          {/* Anti-Islanding Protection (IEEE 1547) */}
          <div className="space-y-2 bg-slate-50 p-3 rounded-xl border border-slate-200">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <Radio size={14} className="text-amber-600" />
                <span>{isVi ? 'B3. Chống đảo lưới (IEEE 1547):' : 'B3. Anti-Islanding (IEEE 1547):'}</span>
              </span>
              <span className={`text-[10px] font-black px-1.5 py-0.5 rounded ${
                antiIslandingTripTime <= 2.0 ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'
              }`}>
                {antiIslandingTripTime <= 2.0 ? 'ĐẠT (≤ 2.0s)' : 'FAIL (> 2.0s)'}
              </span>
            </div>
            <div>
              <div className="flex items-center justify-between text-[11px] text-slate-600 mb-1">
                <span>{isVi ? 'Thời gian Inverter ngắt khi mất lưới:' : 'Trip Time upon AC Loss:'}</span>
                <span className="font-mono font-bold text-slate-900">{antiIslandingTripTime} giây (s)</span>
              </div>
              <input
                type="range"
                min="0.05"
                max="3.0"
                step="0.05"
                value={antiIslandingTripTime}
                onChange={(e) => setAntiIslandingTripTime(parseFloat(e.target.value))}
                className="w-full accent-amber-500"
              />
            </div>
          </div>

          {/* Grid Code & Power Factor Tuning */}
          <div className="space-y-1">
            <label className="text-xs font-bold text-slate-700">
              {isVi ? 'B4. Hòa lưới & Hệ số Cosφ (Grid Code):' : 'B4. Grid Code & Power Factor Tuning:'}
            </label>
            <button
              type="button"
              onClick={() => setIsGridCodePass(!isGridCodePass)}
              className={`w-full py-2 px-3 text-xs font-bold rounded-xl border flex items-center justify-between ${
                isGridCodePass
                  ? 'bg-slate-50 text-slate-800 border-slate-300'
                  : 'bg-rose-50 text-rose-800 border-rose-300'
              }`}
            >
              <span>{isGridCodePass ? 'Cosφ = 0.99 lead/lag (Đạt Grid Code)' : 'Lỗi điều độ / Cosφ'}</span>
              <Gauge size={15} className="text-slate-500" />
            </button>
          </div>
        </div>

        {/* PHASE 3: POST-ENERGIZATION (NETA ECS 8.1.C) */}
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
              NETA ECS 8.1.C
            </span>
          </div>

          {/* C1: Baseline Parameters & DAS */}
          <div className="space-y-1">
            <label className="text-xs font-bold text-slate-700">
              {isVi ? 'C1. Dữ liệu nền cảm biến bức xạ & DAS:' : 'C1. Baseline DAS & Weather Station:'}
            </label>
            <button
              type="button"
              onClick={() => setIsDasBaselinePass(!isDasBaselinePass)}
              className={`w-full py-2 px-3 text-xs font-bold rounded-xl border flex items-center justify-between ${
                isDasBaselinePass
                  ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                  : 'bg-amber-50 text-amber-800 border-amber-300'
              }`}
            >
              <span>{isDasBaselinePass ? (isVi ? 'Pyranometer & RTD chuẩn xác' : 'Calibrated & Logged') : (isVi ? 'Sai số cảm biến bức xạ' : 'Sensor error')}</span>
              <Activity size={15} />
            </button>
          </div>

          {/* C2: Thermographic Survey Table 100.18 */}
          <div className="space-y-2 bg-slate-50 p-3 rounded-xl border border-slate-200">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <Flame size={14} className="text-rose-500" />
                <span>{isVi ? 'C2. Khảo sát nhiệt IR (Table 100.18):' : 'C2. Thermography (Table 100.18):'}</span>
              </span>
              <button
                type="button"
                onClick={() => setIsThermalPass(!isThermalPass)}
                className={`text-[10px] font-bold px-2 py-0.5 rounded ${
                  isThermalPass ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'
                }`}
              >
                {isThermalPass ? (isVi ? 'Không Hotspot' : 'Normal') : (isVi ? 'CẢNH BÁO HOTSPOT' : 'HOTSPOT DETECTED')}
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
                max="35"
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

          {/* C3: NERC MOD-025-2 Reactive Power Verification */}
          <div className="space-y-1">
            <label className="text-xs font-bold text-slate-700">
              {isVi ? 'C3. NERC MOD-025-2 Công suất phản kháng:' : 'C3. NERC MOD-025-2 Reactive Power:'}
            </label>
            <button
              type="button"
              onClick={() => setIsNercPass(!isNercPass)}
              className={`w-full py-2 px-3 text-xs font-bold rounded-xl border flex items-center justify-between ${
                isNercPass
                  ? 'bg-emerald-50 text-emerald-800 border-emerald-300'
                  : 'bg-rose-50 text-rose-800 border-rose-300'
              }`}
            >
              <span>{isNercPass ? 'Đạt 100% biểu đồ P-Q (NERC MOD-025-2)' : 'Thiếu dải công suất phản kháng'}</span>
              <ShieldCheck size={15} className="text-emerald-600" />
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
              <Zap size={14} className="text-amber-400" />
              <span>{isVi ? 'Chẩn Đoán NETA ECS (Tất Định)' : 'Run Deterministic Audit'}</span>
            </button>

            <button
              type="button"
              onClick={handleRunAiAnalysis}
              disabled={isAnalyzing}
              className="w-full py-2.5 px-4 bg-gradient-to-r from-amber-500 to-amber-400 hover:from-amber-400 hover:to-amber-300 text-slate-950 rounded-xl text-xs font-black transition-all flex items-center justify-center gap-2 shadow-md shadow-amber-500/20"
            >
              <Sparkles size={14} />
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
              {isVi ? 'Bảng Đối Soát 16 Hạng Mục (NETA ECS 8.1)' : 'Evaluation Table (16 Items)'}
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
                  activeResultTab === 'ai-markdown' ? 'bg-amber-500 text-slate-950 font-black shadow-xs' : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Sparkles size={12} />
                <span>{isVi ? 'Báo Cáo Gemini AI' : 'Gemini AI Report'}</span>
              </button>
            )}
          </div>

          {/* TAB 1: SUMMARY & HAZARDS */}
          {activeResultTab === 'summary' && (
            <div className="p-5 space-y-4">
              <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
                <h4 className="text-sm font-bold text-slate-900 flex items-center gap-2">
                  <Activity size={16} className="text-amber-600" />
                  <span>{isVi ? evaluationResult.technical_analysis.title_vn : evaluationResult.technical_analysis.title_en}</span>
                </h4>
                <p className="text-xs text-slate-700 leading-relaxed">
                  {isVi ? evaluationResult.technical_analysis.description_vn : evaluationResult.technical_analysis.description_en}
                </p>
              </div>

              {/* Key KPI Chips */}
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                <div className="p-3 rounded-xl bg-slate-50 border border-slate-200">
                  <div className="text-[10px] text-slate-500 font-bold uppercase">Cách điện Khô Min</div>
                  <div className={`text-base font-black ${
                    evaluationResult.technical_analysis.dry_insulation_min_val < minRequiredDryIr ? 'text-rose-600' : 'text-emerald-600'
                  }`}>
                    {evaluationResult.technical_analysis.dry_insulation_min_val} MΩ
                  </div>
                  <div className="text-[9px] text-slate-400">Chuẩn ≥ {minRequiredDryIr} MΩ</div>
                </div>

                <div className="p-3 rounded-xl bg-slate-50 border border-slate-200">
                  <div className="text-[10px] text-slate-500 font-bold uppercase">Độ lệch Voc (Max)</div>
                  <div className={`text-base font-black ${
                    evaluationResult.technical_analysis.voc_deviation_percent > 5.0 ? 'text-rose-600' : 'text-emerald-600'
                  }`}>
                    {evaluationResult.technical_analysis.voc_deviation_percent}%
                  </div>
                  <div className="text-[9px] text-slate-400">Giới hạn ≤ 5.0%</div>
                </div>

                <div className="p-3 rounded-xl bg-slate-50 border border-slate-200">
                  <div className="text-[10px] text-slate-500 font-bold uppercase">Pmax STC String 04</div>
                  <div className={`text-base font-black ${
                    evaluationResult.technical_analysis.pmax_stc_percent < 95.0 ? 'text-rose-600' : 'text-emerald-600'
                  }`}>
                    {evaluationResult.technical_analysis.pmax_stc_percent}%
                  </div>
                  <div className="text-[9px] text-slate-400">Yêu cầu ≥ 95.0%</div>
                </div>

                <div className="p-3 rounded-xl bg-slate-50 border border-slate-200">
                  <div className="text-[10px] text-slate-500 font-bold uppercase">Chống Đảo Lưới</div>
                  <div className={`text-base font-black ${
                    evaluationResult.technical_analysis.anti_islanding_time > 2.0 ? 'text-rose-600' : 'text-emerald-600'
                  }`}>
                    {evaluationResult.technical_analysis.anti_islanding_time}s
                  </div>
                  <div className="text-[9px] text-slate-400">IEEE 1547 ≤ 2.0s</div>
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
                    <th className="py-2.5 px-3 font-bold">Hạng Mục Kiểm Tra (NETA ECS 8.1)</th>
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
                {evaluationResult.actionable_recommendations.halt_energization && (
                  <span className="px-2.5 py-1 rounded-full text-xs font-black bg-rose-600 text-white animate-pulse">
                    {isVi ? '⚠️ DỪNG ĐÓNG ĐIỆN STRING 04' : '⚠️ HALT STRING 04 ENERGIZATION'}
                  </span>
                )}
              </div>

              <div className="space-y-3">
                {evaluationResult.actionable_recommendations.steps.map((step) => (
                  <div key={step.step_number} className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 flex gap-3">
                    <div className="w-7 h-7 rounded-lg bg-amber-500 text-slate-950 font-black flex items-center justify-center shrink-0 text-xs">
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
                <h3 className="font-bold text-sm">System Instruction Prompt (Google AI Studio)</h3>
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
              {ECS_SOLAR_PV_SYSTEM_INSTRUCTION}
            </div>
            <div className="p-3 bg-slate-900 border-t border-slate-800 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => copyToClipboard(ECS_SOLAR_PV_SYSTEM_INSTRUCTION, 'prompt')}
                className="px-4 py-2 rounded-xl text-xs font-black bg-amber-500 text-slate-950 hover:bg-amber-400 flex items-center gap-1.5"
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
                <h3 className="font-bold text-sm">JSON User Prompt Payload</h3>
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
                className="px-4 py-2 rounded-xl text-xs font-black bg-amber-500 text-slate-950 hover:bg-amber-400 flex items-center gap-1.5"
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
