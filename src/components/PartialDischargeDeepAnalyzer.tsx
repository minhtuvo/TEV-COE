import React, { useState, useEffect, useRef, useMemo } from 'react';
import {
  AudioWaveform,
  Zap,
  Thermometer,
  ShieldAlert,
  ShieldCheck,
  AlertTriangle,
  AlertOctagon,
  CheckCircle2,
  XCircle,
  Sparkles,
  Download,
  Printer,
  Copy,
  Check,
  RotateCcw,
  Sliders,
  Info,
  Layers,
  Camera,
  Upload,
  Image as ImageIcon,
  FileText,
  FileCheck2,
  ChevronDown,
  ChevronUp,
  Maximize2,
  Clock,
  ArrowRight,
  Eye,
  SlidersHorizontal,
  Save,
  Radio,
  Flame,
  Volume2,
  Search,
  Plus,
  Trash2
} from 'lucide-react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as htmlToImage from 'html-to-image';
import {
  EaPdDefectType,
  PrpdVisionAnalysisResult,
  ThermalOcrResult,
  MultiModalPdPayload,
  CrossModalityDiagnosticResult,
  evaluateCrossModalityEngine,
  EA_SENSOR_MATRIX,
  SAMPLE_EA_PD_PAYLOAD
} from '../../server/eaTechnologyPdAnalyzer';
import { FseEngineerDropdown } from './FseEngineerDropdown';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

// ----------------------------------------------------------------------------
// PRESET REALISTIC SCENARIOS WITH BASE64 MOCK PLOTS & FIELD CONDITIONS
// ----------------------------------------------------------------------------
interface PresetScenario {
  id: string;
  name: string;
  defectType: EaPdDefectType;
  badge: string;
  badgeColor: string;
  cubicleId: string;
  voltage: string;
  tev: number;
  us: number;
  tMax: number;
  tRef: number;
  rh: number;
  descriptionVn: string;
}

const PRESET_SCENARIOS: PresetScenario[] = [
  {
    id: 'surface-spout',
    name: 'Phóng điện Bề mặt Sứ Spout Bushing (Vết than Carbon Treeing)',
    defectType: 'SURFACE_DISCHARGE',
    badge: 'HIGH CONFIRMED / CẦN CẮT ĐIỆN',
    badgeColor: 'bg-rose-500/20 text-rose-300 border-rose-400/40',
    cubicleId: 'J02 - Feeder 22kV',
    voltage: '22 kV',
    tev: 36,
    us: 28,
    tMax: 54.2,
    tRef: 32.0,
    rh: 86.0,
    descriptionVn: '2 cụm đám mây hình vòm/đồi rộng (Parabolic mound) trên sườn lên và đỉnh sin. Siêu âm 28 dBµV rất lớn, tiếng rít lách tách.'
  },
  {
    id: 'internal-void-ct',
    name: 'Phóng điện Lỗ rỗng trong Biến dòng CT Khô Epoxy (Internal Void PD)',
    defectType: 'INTERNAL_VOID',
    badge: 'HIGH / CẦN ĐẠI TU',
    badgeColor: 'bg-rose-500/20 text-rose-300 border-rose-400/40',
    cubicleId: 'J01 - Incomer 22kV',
    voltage: '22 kV',
    tev: 34,
    us: 4,
    tMax: 46.5,
    tRef: 31.0,
    rh: 64.0,
    descriptionVn: '2 cụm hình tam giác nghiêng trái cách nhau 180° (sườn lên dốc đứng, sườn xuống thoải dần). TEV 34 dBmV, siêu âm = 0.'
  },
  {
    id: 'poor-contact-busbar',
    name: 'Tiếp xúc xấu Mối nối Bu-lông Thanh cái / Ngàm Máy cắt (Poor Contact)',
    defectType: 'POOR_CONTACT_FLOATING',
    badge: 'CỘNG HƯỞNG NGUY HIỂM / ΔT > 15°C',
    badgeColor: 'bg-red-600 text-white border-red-400',
    cubicleId: 'J04 - Bus Coupler 22kV',
    voltage: '22 kV',
    tev: 31,
    us: 22,
    tMax: 68.4,
    tRef: 33.2,
    rh: 72.0,
    descriptionVn: '2 dải nằm ngang đỉnh phẳng (Flat-top band) ở biên độ cố định, khoảng dưới rỗng. Kèm ΔT = 35.2°C phát nhiệt ngàm cực kỳ nguy hiểm.'
  },
  {
    id: 'corona-busbar-bolt',
    name: 'Phóng điện Vầng quang Đầu bu-lông nhọn Thanh cái (Corona Discharge)',
    defectType: 'CORONA',
    badge: 'MEDIUM / THEO DÕI 3-6 THÁNG',
    badgeColor: 'bg-amber-500/20 text-amber-300 border-amber-400/40',
    cubicleId: 'J03 - Outgoing 6.6kV',
    voltage: '6.6 kV',
    tev: 9,
    us: 24,
    tMax: 36.0,
    tRef: 32.5,
    rh: 89.0,
    descriptionVn: 'Chỉ có 1 cụm duy nhất tập trung quanh đỉnh âm 270°, bán kỳ dương trống. Độ ẩm 89% kích thích ion hóa không khí.'
  },
  {
    id: 'spark-gap-cable',
    name: 'Phóng tia lửa Điện cực Khe hở Cổ cáp (Spark Gap Discharge)',
    defectType: 'SPARK_GAP',
    badge: 'HIGH CONFIRMED / CÁP ĐỘNG LỰC',
    badgeColor: 'bg-rose-500/20 text-rose-300 border-rose-400/40',
    cubicleId: 'J05 - Cable Incomer 22kV',
    voltage: '22 kV',
    tev: 38,
    us: 26,
    tMax: 51.0,
    tRef: 33.0,
    rh: 78.0,
    descriptionVn: 'Các cụm hạt cô đặc lơ lửng ở biên độ cao (1-2 chùm dương, 2-3 chùm âm). Phát hiện trên cả EM-TEV và Siêu âm.'
  }
];

export const PartialDischargeDeepAnalyzer: React.FC<Props> = ({
  onSaveReport,
  onNavigateToReports
}) => {
  // --------------------------------------------------------------------------
  // STATE 1: METADATA FORM (Phân hệ 1)
  // --------------------------------------------------------------------------
  const [reportNo, setReportNo] = useState(`TEV-COE-2026-PD-${Math.floor(100 + Math.random() * 900)}`);
  const [projectCode, setProjectCode] = useState('2026-TEV-COE-CBM-01');
  const [customerName, setCustomerName] = useState('Nhà máy Nhiệt điện TEV Vũng Áng');
  const [substationName, setSubstationName] = useState('Trạm Phân Phối 22kV Switchgear Room');
  const [surveyDate, setSurveyDate] = useState(new Date().toISOString().split('T')[0]);
  const [inspectorName, setInspectorName] = useState('Võ Minh Tú');

  const [cubicleId, setCubicleId] = useState('J02 - Incoming 22kV Feeder');
  const [cubicleType, setCubicleType] = useState('Metal-Clad');
  const [testPosition, setTestPosition] = useState('Cáp đầu vào/ra - Cable Box');
  const [manufacturer, setManufacturer] = useState('ABB UniGear ZS1');
  const [ratedVoltage, setRatedVoltage] = useState('22 kV');
  const [operatingVoltage, setOperatingVoltage] = useState(22.0);
  const [loadCurrent, setLoadCurrent] = useState(480);
  const [loadPercentage, setLoadPercentage] = useState(75);

  const [ambientTemp, setAmbientTemp] = useState(32.5);
  const [relativeHumidity, setRelativeHumidity] = useState(86.0);
  const [backgroundNoiseTEV, setBackgroundNoiseTEV] = useState(8);
  const [backgroundNoiseUS, setBackgroundNoiseUS] = useState(3);

  const [thermalMaxTemp, setThermalMaxTemp] = useState(54.2);
  const [thermalRefTemp, setThermalRefTemp] = useState(32.0);
  const thermalDeltaT = useMemo(() => {
    return Math.max(0, Number((thermalMaxTemp - thermalRefTemp).toFixed(1)));
  }, [thermalMaxTemp, thermalRefTemp]);

  const [sensoryChecks, setSensoryChecks] = useState({
    ozoneSmell: true,
    audibleSound: true,
    heatAnomaly: true,
    visualTracking: true
  });

  // --------------------------------------------------------------------------
  // STATE 2: 3-SLOT MEDIA UPLOADER & IMAGES (Phân hệ 2)
  // --------------------------------------------------------------------------
  const [prpdImage, setPrpdImage] = useState<string | null>(null);
  const [cubicleImage, setCubicleImage] = useState<string | null>(null);
  const [thermalImage, setThermalImage] = useState<string | null>(null);

  const [cubicleSensorLocation, setCubicleSensorLocation] = useState<'Top' | 'Middle' | 'Bottom' | 'Cable Box'>('Cable Box');
  const [cubicleCaption, setCubicleCaption] = useState('Cảm biến TEV đặt tại vách khoang cáp, đầu dò siêu âm rà qua khe tản nhiệt');

  const [isAiVisionRunning, setIsAiVisionRunning] = useState(false);
  const [isThermalOcrRunning, setIsThermalOcrRunning] = useState(false);
  const [aiVisionResult, setAiVisionResult] = useState<PrpdVisionAnalysisResult | null>(null);
  const handleImagePaste = (e: React.ClipboardEvent, setter: (val: string) => void) => {
    const items = e.clipboardData?.items;
    if (!items) return;
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (item.type && item.type.startsWith('image/')) {
        const file = item.getAsFile();
        if (file) {
          const reader = new FileReader();
          reader.onload = () => setter(reader.result as string);
          reader.readAsDataURL(file);
        }
        break;
      }
    }
  };

  // --------------------------------------------------------------------------
  // STATE 3: SIMULATOR SLIDERS & CANVAS (Phân hệ 3)
  // --------------------------------------------------------------------------
  const [simDefectType, setSimDefectType] = useState<EaPdDefectType>('SURFACE_DISCHARGE');
  const [simPeakMagnitude, setSimPeakMagnitude] = useState(36);
  const [simNoiseFloor, setSimNoiseFloor] = useState(6);
  const [simPPC, setSimPPC] = useState(18);
  const [simPhaseShift, setSimPhaseShift] = useState(0);
  const [simIsBipolar, setSimIsBipolar] = useState(true);
  const [simIsPhaseLocked, setSimIsPhaseLocked] = useState(true);

  const prpdCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const ultrasonicCanvasRef = useRef<HTMLCanvasElement | null>(null);

  // --------------------------------------------------------------------------
  // STATE 4: SUMMARY & MODALS
  // --------------------------------------------------------------------------
  const [showPdfModal, setShowPdfModal] = useState(false);
  const [showPromptModal, setShowPromptModal] = useState(false);
  const [copiedPrompt, setCopiedPrompt] = useState(false);
  const [saveSuccessMessage, setSaveSuccessMessage] = useState<string | null>(null);

  // Summary history table of measured cubicles
  interface HistoryRecord {
    id: string;
    cubicleId: string;
    date: string;
    defect: EaPdDefectType;
    tev: number;
    us: number;
    deltaT: number;
    status: string;
    inspector: string;
  }

  const DEFAULT_HISTORY: HistoryRecord[] = [
    {
      id: 'REC-01',
      cubicleId: 'J01 - Incomer 22kV',
      date: '2026-10-15',
      defect: 'INTERNAL_VOID',
      tev: 34,
      us: 4,
      deltaT: 15.5,
      status: 'HIGH',
      inspector: 'Võ Minh Tú'
    },
    {
      id: 'REC-02',
      cubicleId: 'J02 - Feeder 22kV',
      date: '2026-10-16',
      defect: 'SURFACE_DISCHARGE',
      tev: 36,
      us: 28,
      deltaT: 22.2,
      status: 'CRITICAL',
      inspector: 'Võ Minh Tú'
    },
    {
      id: 'REC-03',
      cubicleId: 'J03 - Outgoing 6.6kV',
      date: '2026-10-16',
      defect: 'CORONA',
      tev: 9,
      us: 24,
      deltaT: 3.5,
      status: 'MEDIUM',
      inspector: 'Võ Minh Tú'
    },
    {
      id: 'REC-04',
      cubicleId: 'J04 - Bus Coupler 22kV',
      date: '2026-10-16',
      defect: 'POOR_CONTACT_FLOATING',
      tev: 31,
      us: 22,
      deltaT: 35.2,
      status: 'CRITICAL',
      inspector: 'Võ Minh Tú'
    },
    {
      id: 'REC-05',
      cubicleId: 'J05 - Feeder 22kV',
      date: '2026-10-16',
      defect: 'NOISE',
      tev: 6,
      us: 2,
      deltaT: 1.2,
      status: 'LOW',
      inspector: 'Võ Minh Tú'
    }
  ];

  const [cubicleHistory, setCubicleHistory] = useState<HistoryRecord[]>(() => {
    try {
      const saved = localStorage.getItem('tev_coe_pd_cubicle_history');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch (e) {
      // ignore
    }
    return DEFAULT_HISTORY;
  });

  useEffect(() => {
    try {
      localStorage.setItem('tev_coe_pd_cubicle_history', JSON.stringify(cubicleHistory));
    } catch (e) {
      // ignore
    }
  }, [cubicleHistory]);

  const [historySearch, setHistorySearch] = useState('');
  const [historyStatusFilter, setHistoryStatusFilter] = useState<'ALL' | 'CRITICAL' | 'HIGH' | 'MEDIUM' | 'LOW'>('ALL');

  // --------------------------------------------------------------------------
  // COMPUTE CROSS-MODALITY SEVERITY (Phân hệ 4)
  // --------------------------------------------------------------------------
  const currentPayload: MultiModalPdPayload = useMemo(() => {
    return {
      site_info: {
        reportNo,
        projectCode,
        customerName,
        substationName,
        surveyDate,
        inspectorName,
        cubicleId,
        cubicleType,
        testPosition,
        manufacturer,
        ratedVoltage,
        operatingVoltage,
        loadCurrent,
        loadPercentage
      },
      environmental_conditions: {
        ambientTemp,
        relativeHumidity,
        backgroundNoiseTEV,
        backgroundNoiseUS
      },
      thermal_data: {
        thermalMaxTemp,
        thermalRefTemp,
        thermalDeltaT
      },
      sensory_checks: sensoryChecks,
      pd_measurements: {
        maxTev_dBmV: simPeakMagnitude,
        maxUltrasonic_dBuV: simDefectType === 'INTERNAL_VOID' ? 4 : simDefectType === 'CORONA' ? 24 : 28,
        pulsesPerCycle: simPPC,
        firstCountAntennaPercent: simDefectType === 'NOISE' ? 88 : 12,
        firstCountInternalPercent: simDefectType === 'NOISE' ? 12 : 88,
        isPhaseLocked: simIsPhaseLocked
      },
      prpd_characteristics: {
        defectType: simDefectType,
        phaseDistribution: simDefectType === 'CORONA' ? 'Quanh 270°' : '45°-135° & 225°-315°',
        patternShape: simDefectType,
        bipolar: simIsBipolar
      },
      image_annotations: {
        cubicleSensorLocation,
        cubicleCaption
      }
    };
  }, [
    reportNo, projectCode, customerName, substationName, surveyDate, inspectorName,
    cubicleId, cubicleType, testPosition, manufacturer, ratedVoltage, operatingVoltage, loadCurrent, loadPercentage,
    ambientTemp, relativeHumidity, backgroundNoiseTEV, backgroundNoiseUS,
    thermalMaxTemp, thermalRefTemp, thermalDeltaT, sensoryChecks,
    simDefectType, simPeakMagnitude, simPPC, simIsPhaseLocked, simIsBipolar,
    cubicleSensorLocation, cubicleCaption
  ]);

  const diagnosticResult = useMemo(() => {
    return evaluateCrossModalityEngine(currentPayload);
  }, [currentPayload]);

  // --------------------------------------------------------------------------
  // RENDER HTML5 CANVASES (Phân hệ 3)
  // --------------------------------------------------------------------------
  useEffect(() => {
    // 1. Render PRPD Cloud (Canvas 1: 360° x 60dB)
    const canvas = prpdCanvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const width = canvas.width;
    const height = canvas.height;

    // Background
    ctx.fillStyle = '#090d16';
    ctx.fillRect(0, 0, width, height);

    // Grid lines
    ctx.strokeStyle = '#1e293b';
    ctx.lineWidth = 1;

    // 0°, 90°, 180°, 270°, 360° phase lines
    [0, 90, 180, 270, 360].forEach((deg) => {
      const x = (deg / 360) * width;
      ctx.beginPath();
      ctx.moveTo(x, 0);
      ctx.lineTo(x, height);
      ctx.stroke();

      ctx.fillStyle = '#64748b';
      ctx.font = '10px monospace';
      ctx.fillText(`${deg}°`, x + 2, height - 6);
    });

    // Amplitude grid lines (0, 20, 40, 60 dB)
    [0, 20, 40, 60].forEach((db) => {
      const y = height - (db / 60) * height;
      ctx.beginPath();
      ctx.moveTo(0, y);
      ctx.lineTo(width, y);
      ctx.stroke();

      if (db > 0 && db < 60) {
        ctx.fillStyle = '#64748b';
        ctx.font = '10px monospace';
        ctx.fillText(`${db} dB`, 6, y - 3);
      }
    });

    // Reference 50Hz Sine Wave (in Red per EA Tech)
    ctx.strokeStyle = '#ef4444';
    ctx.lineWidth = 2;
    ctx.beginPath();
    const midY = height / 2;
    for (let x = 0; x <= width; x++) {
      const angle = (x / width) * 2 * Math.PI + (simPhaseShift * Math.PI) / 180;
      const y = midY - Math.sin(angle) * (height * 0.38);
      if (x === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    }
    ctx.stroke();

    // Pulse Density Points Generation based on simDefectType
    // Color Palette: Green (<2), Blue (2-6), Red (6-12), Dark Brown/Black (>12)
    const points: Array<{ x: number; y: number; density: number }> = [];
    const count = Math.min(600, Math.max(80, simPPC * 25));

    for (let i = 0; i < count; i++) {
      let phaseDeg = 0;
      let ampDb = 0;

      if (simDefectType === 'INTERNAL_VOID') {
        // 2 triangular clusters tilted left (rising edge steep, falling slope gradual)
        const isPositive = Math.random() > 0.5;
        const center = isPositive ? 65 : 245;
        // Triangular skew: random with left-tilt
        const u = Math.random();
        const spread = u < 0.4 ? (Math.random() - 0.5) * 25 : (Math.random() * 45);
        phaseDeg = center + spread;
        const peak = simPeakMagnitude;
        ampDb = Math.max(simNoiseFloor, peak * (0.4 + 0.6 * (1 - Math.abs(spread) / 45)));
      } else if (simDefectType === 'SURFACE_DISCHARGE') {
        // 2 wide parabolic mounds covering 45-135 and 225-315
        const isPositive = Math.random() > 0.5;
        const center = isPositive ? 90 : 270;
        const spread = (Math.random() - 0.5) * 80;
        phaseDeg = center + spread;
        ampDb = Math.max(simNoiseFloor, simPeakMagnitude * (0.3 + 0.7 * Math.cos((spread / 80) * (Math.PI / 2))));
      } else if (simDefectType === 'CORONA') {
        // 1 single narrow cluster around negative peak 270
        const spread = (Math.random() - 0.5) * 40;
        phaseDeg = 270 + spread;
        ampDb = simPeakMagnitude * (0.8 + 0.2 * (Math.random() - 0.5));
      } else if (simDefectType === 'POOR_CONTACT_FLOATING') {
        // 2 flat-top horizontal bands at fixed amplitude across phases
        const isPositive = Math.random() > 0.5;
        phaseDeg = Math.random() * 360;
        const baseBand = isPositive ? simPeakMagnitude : simPeakMagnitude * 0.85;
        ampDb = baseBand + (Math.random() - 0.5) * 3;
      } else if (simDefectType === 'SPARK_GAP') {
        // Discrete tight clusters hovering at high amplitude
        const clusters = [55, 85, 235, 265, 295];
        const pick = clusters[Math.floor(Math.random() * clusters.length)];
        phaseDeg = pick + (Math.random() - 0.5) * 12;
        ampDb = simPeakMagnitude * 0.9 + (Math.random() - 0.5) * 5;
      } else {
        // NOISE: Random 0-360
        phaseDeg = Math.random() * 360;
        ampDb = simNoiseFloor + Math.random() * (simPeakMagnitude * 0.3);
      }

      phaseDeg = ((phaseDeg % 360) + 360) % 360;
      const x = (phaseDeg / 360) * width;
      const y = height - (Math.min(60, Math.max(0, ampDb)) / 60) * height;

      // Assign density score for coloring
      const density = Math.floor(Math.random() * 15);
      points.push({ x, y, density });
    }

    // Draw points with EA Technology color gradient
    points.forEach((p) => {
      let color = '#22c55e'; // Green < 2
      if (p.density >= 12) color = '#78350f'; // Dark Brown > 12
      else if (p.density >= 6) color = '#ef4444'; // Red 6-12
      else if (p.density >= 2) color = '#3b82f6'; // Blue 2-6

      ctx.fillStyle = color;
      ctx.beginPath();
      ctx.arc(p.x, p.y, 2.5, 0, 2 * Math.PI);
      ctx.fill();
    });

    // 2. Render Ultrasonic Envelope (Canvas 2: 360° x envelope)
    const usCanvas = ultrasonicCanvasRef.current;
    if (!usCanvas) return;
    const usCtx = usCanvas.getContext('2d');
    if (!usCtx) return;

    const uWidth = usCanvas.width;
    const uHeight = usCanvas.height;

    usCtx.fillStyle = '#0f172a';
    usCtx.fillRect(0, 0, uWidth, uHeight);

    usCtx.strokeStyle = '#334155';
    usCtx.lineWidth = 1;
    [0, 90, 180, 270, 360].forEach((deg) => {
      const x = (deg / 360) * uWidth;
      usCtx.beginPath();
      usCtx.moveTo(x, 0);
      usCtx.lineTo(x, uHeight);
      usCtx.stroke();
    });

    // Draw simulated ultrasonic envelope
    usCtx.strokeStyle = '#84cc16'; // Lime green
    usCtx.lineWidth = 2.5;
    usCtx.beginPath();

    const baseline = uHeight * 0.75;

    for (let x = 0; x <= uWidth; x++) {
      const deg = (x / uWidth) * 360;
      let amp = 0;

      if (simDefectType === 'INTERNAL_VOID') {
        // Flat noise with slight random variation
        amp = (Math.random() - 0.5) * 4;
      } else if (simDefectType === 'SURFACE_DISCHARGE') {
        // 2 distinct rounded humps/mounds at 90 and 270
        const d1 = Math.abs(deg - 90);
        const d2 = Math.abs(deg - 270);
        if (d1 < 45) amp = Math.cos((d1 / 45) * (Math.PI / 2)) * 32;
        if (d2 < 45) amp = Math.cos((d2 / 45) * (Math.PI / 2)) * 32;
        amp += (Math.random() - 0.5) * 3;
      } else if (simDefectType === 'CORONA') {
        // 1 single rounded peak at 270
        const d = Math.abs(deg - 270);
        if (d < 45) amp = Math.cos((d / 45) * (Math.PI / 2)) * 35;
        amp += (Math.random() - 0.5) * 2;
      } else if (simDefectType === 'POOR_CONTACT_FLOATING') {
        // Flat-top square wave step pulses
        if ((deg >= 45 && deg <= 135) || (deg >= 225 && deg <= 315)) {
          amp = 28 + (Math.random() - 0.5) * 3;
        } else {
          amp = (Math.random() - 0.5) * 3;
        }
      } else if (simDefectType === 'SPARK_GAP') {
        // Sharp discrete needle bursts / spikes
        const spikes = [50, 75, 100, 230, 260, 290];
        spikes.forEach((s) => {
          if (Math.abs(deg - s) < 3) amp = 34;
        });
        amp += (Math.random() - 0.5) * 3;
      } else {
        amp = (Math.random() - 0.5) * 5;
      }

      const y = Math.max(10, Math.min(uHeight - 5, baseline - amp));
      if (x === 0) usCtx.moveTo(x, y);
      else usCtx.lineTo(x, y);
    }
    usCtx.stroke();
  }, [simDefectType, simPeakMagnitude, simNoiseFloor, simPPC, simPhaseShift, simIsBipolar]);

  // --------------------------------------------------------------------------
  // AI VISION HANDLER: PRPD IMAGE ANALYSIS (Phân hệ 2)
  // --------------------------------------------------------------------------
  const handleRunAiVision = async () => {
    if (!prpdImage) return;
    setIsAiVisionRunning(true);

    try {
      const res = await fetch('/api/deep-analysis/pd-vision-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          imageBase64: prpdImage,
          contextMetadata: { cubicleId, voltage: ratedVoltage }
        })
      });

      const data = await res.json();
      if (data.success && data.data) {
        const result: PrpdVisionAnalysisResult = data.data;
        setAiVisionResult(result);

        // Auto-sync into simulator
        setSimDefectType(result.primaryPDType);
        setSimPeakMagnitude(result.estimatedPeakMagnitude);
        setSimNoiseFloor(result.estimatedNoiseFloor);
        setSimPPC(result.estimatedPPC);
        setSimPhaseShift(result.estimatedPhaseShift);
        setSimIsBipolar(result.isBipolar);
        setSimIsPhaseLocked(result.isPhaseLocked);
      }
    } catch (err) {
      console.error('Vision analysis error:', err);
    } finally {
      setIsAiVisionRunning(false);
    }
  };

  // --------------------------------------------------------------------------
  // AI THERMAL OCR HANDLER: READ T_MAX, T_REF FROM IR IMAGE (Phân hệ 2)
  // --------------------------------------------------------------------------
  const handleRunThermalOcr = async () => {
    if (!thermalImage) return;
    setIsThermalOcrRunning(true);

    try {
      const res = await fetch('/api/deep-analysis/pd-thermal-ocr', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ imageBase64: thermalImage })
      });

      const data = await res.json();
      if (data.success && data.data) {
        const ocr: ThermalOcrResult = data.data;
        setThermalMaxTemp(ocr.maxTemp);
        setThermalRefTemp(ocr.refTemp);
      }
    } catch (err) {
      console.error('Thermal OCR error:', err);
    } finally {
      setIsThermalOcrRunning(false);
    }
  };

  // --------------------------------------------------------------------------
  // FAST SCENARIO PRESET LOADER
  // --------------------------------------------------------------------------
  const handleLoadPreset = (scenario: PresetScenario) => {
    setSimDefectType(scenario.defectType);
    setSimPeakMagnitude(scenario.tev);
    setCubicleId(scenario.cubicleId);
    setRatedVoltage(scenario.voltage);
    setThermalMaxTemp(scenario.tMax);
    setThermalRefTemp(scenario.tRef);
    setRelativeHumidity(scenario.rh);
  };

  // --------------------------------------------------------------------------
  // SAVE TO HISTORY & CMMS
  // --------------------------------------------------------------------------
  const handleSaveToHistory = () => {
    const newRecord: HistoryRecord = {
      id: `REC-${Date.now().toString().slice(-4)}`,
      cubicleId,
      date: surveyDate,
      defect: simDefectType,
      tev: simPeakMagnitude,
      us: simDefectType === 'INTERNAL_VOID' ? 4 : simDefectType === 'CORONA' ? 24 : 28,
      deltaT: thermalDeltaT,
      status: diagnosticResult.overallStatus,
      inspector: inspectorName
    };

    setCubicleHistory(prev => [newRecord, ...prev]);

    if (onSaveReport) {
      onSaveReport({
        data: currentPayload,
        analysisResult: diagnosticResult,
        equipmentId: cubicleId,
        date: surveyDate,
        type: `TEV COE Multi-Modal PD & Thermal IR Report`
      });
    }

    setSaveSuccessMessage(`Đã lưu kết quả chẩn đoán ngăn lộ [${cubicleId}] vào danh sách tổng hợp!`);
    setTimeout(() => setSaveSuccessMessage(null), 4000);
  };

  // --------------------------------------------------------------------------
  // MASTER SYSTEM INSTRUCTION PROMPT FOR GOOGLE AI STUDIO
  // --------------------------------------------------------------------------
  const MASTER_SYSTEM_INSTRUCTION = `You are the TEV Platform AI Engine — a Master Diagnostic Expert Assistant specializing in Partial Discharge (PD) Analysis for Medium Voltage (MV) Switchgear (6.6kV - 36kV), operating strictly according to EA Technology methodologies, ANSI/NETA ATS/MTS (Table 100.23), and DNO CNAIM standards.

Your primary function is to process JSON input containing PD field test data (TEV dBmV, Airborne Ultrasonic dBµV, HFCT pC, %First Count, Pulses Per Cycle - PPC, and PRPD Phase Patterns) and output an expert-level diagnostic report.

1. CORE KNOWLEDGE BASE & PRPD PATTERN RECOGNITION MATRIX (EA Technology):
1.1 INTERNAL VOID PD: 2 symmetrical triangular clusters aligned with AC cycles (45°-90° and 225°-270°). TEV = YES, Ultrasonic = NO.
1.2 SURFACE TRACKING PD: 2 parabolic mounds with "Rabbit Ears" pointing vertically to peaks. Ultrasonic = YES, TEV = YES. Accompanied by hiss/crackling sound.
1.3 CORONA DISCHARGE: Flat narrow band around negative voltage peak (~270°). Ultrasonic = YES, TEV = NO. Strongly humidity dependent.
1.4 POOR CONTACT / FLOATING METAL: 2 flat-top horizontal parallel bars across phases at fixed amplitude with empty space below. High PPC (>10-50). TEV = YES, Ultrasonic = YES.
1.5 SPARK GAP: Discrete tight clusters hovering at high amplitude. TEV = YES, Ultrasonic = YES.
1.6 NOISE: Flat random distribution across 0°-360°. %FC Aerial Antenna > %FC Internal Sensors.

2. SEVERITY EVALUATION MATRIX (EA TECHNOLOGY & NETA TABLE 100.23):
- GREEN (LOW): TEV < 20 dBmV & Ultrasonic < 6 dBµV. Routine 12-month monitoring.
- AMBER (MEDIUM): TEV 20–29 dBmV OR Ultrasonic 6–15 dBµV. Increase monitoring to 3–6 months.
- RED (HIGH): TEV 30–39 dBmV OR Ultrasonic 16–29 dBµV. Schedule outage within 15–30 days.
- CRITICAL (DANGER): TEV >= 40 dBmV OR Ultrasonic >= 30 dBµV OR confirmed Poor Contact + ΔT >= 10°C. EMERGENCY OUTAGE (< 24-48 Hours).
`;

  // --------------------------------------------------------------------------
  // DIRECT PDF EXPORT (Phân hệ 5: A4 Portrait with 2x2 Image Grid)
  // --------------------------------------------------------------------------
  const handleExportA4Pdf = () => {
    const doc = new jsPDF('p', 'mm', 'a4');
    const pageWidth = doc.internal.pageSize.getWidth();

    // TEV Header
    doc.setFillColor(15, 23, 42); // slate-900
    doc.rect(0, 0, pageWidth, 28, 'F');

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(14);
    doc.setFont('helvetica', 'bold');
    doc.text('CÔNG TY CỔ PHẦN KỸ NGHỆ NĂNG LƯỢNG VIỆT NAM (TEV)', 14, 11);

    doc.setFontSize(8);
    doc.setFont('helvetica', 'normal');
    doc.text('TEV CENTER OF EXCELLENCE (COE) • HIGH-VOLTAGE ASSET MANAGEMENT & PD CBM', 14, 17);
    doc.text(`REPORT NO: ${reportNo} • DATE: ${surveyDate}`, 14, 23);

    // Document Title
    doc.setTextColor(15, 23, 42);
    doc.setFontSize(13);
    doc.setFont('helvetica', 'bold');
    doc.text('BIÊN BẢN KHẢO SÁT & CHẨN ĐOÁN PHÓNG ĐIỆN CỤC BỘ (PD) & NHIỆT ẢNH', 14, 37);

    doc.setFontSize(9);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(71, 85, 105);
    doc.text(`Tiêu chuẩn: EA Technology UltraTEV Plus² • ANSI/NETA MTS-2025 (Table 100.23) • DNO CNAIM`, 14, 43);

    // Section 1: Metadata Table
    const metaTable = [
      [`Khách hàng / Nhà máy: ${customerName}`, `Mã dự án: ${projectCode}`],
      [`Trạm biến áp / Phòng: ${substationName}`, `Kỹ sư khảo sát: ${inspectorName} (NETA III)`],
      [`Ngăn lộ / Tủ điện: ${cubicleId} (${cubicleType})`, `Hãng SX: ${manufacturer} • Điện áp: ${ratedVoltage}`],
      [`Vị trí đặt cảm biến: ${testPosition} (${cubicleSensorLocation})`, `Dòng tải: ${loadCurrent} A (${loadPercentage}% I_rate)`],
      [`Nhiệt độ phòng / Độ ẩm: ${ambientTemp}°C / ${relativeHumidity}% RH`, `Nhiệt độ IR: Tmax=${thermalMaxTemp}°C, Tref=${thermalRefTemp}°C (ΔT=${thermalDeltaT}°C)`]
    ];

    autoTable(doc, {
      startY: 47,
      body: metaTable,
      theme: 'grid',
      styles: { fontSize: 8, cellPadding: 2, textColor: [30, 41, 59] },
      columnStyles: {
        0: { fontStyle: 'bold', cellWidth: 95 },
        1: { cellWidth: 87 }
      }
    });

    let currentY = (doc as any).lastAutoTable?.finalY || 80;

    // Section 2: 2x2 Image Grid
    doc.setFontSize(10);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(15, 23, 42);
    doc.text('MỤC II: MINH CHỨNG HÌNH ẢNH HIỆN TRƯỜNG & MÔ PHỎNG ĐÁM MÂY PRPD (LƯỚI 2x2)', 14, currentY + 7);

    const cellW = 85;
    const cellH = 46;
    const startX1 = 14;
    const startX2 = 106;
    const startY1 = currentY + 10;
    const startY2 = startY1 + cellH + 6;

    // Slot 1: Cubicle Photo
    doc.setDrawColor(203, 213, 225);
    doc.rect(startX1, startY1, cellW, cellH);
    doc.setFontSize(7.5);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(51, 65, 85);
    doc.text('1. ẢNH MẶT TRƯỚC TỦ ĐIỆN / ĐIỂM ĐO', startX1 + 2, startY1 + 4);
    if (cubicleImage) {
      doc.addImage(cubicleImage, 'JPEG', startX1 + 2, startY1 + 6, cellW - 4, cellH - 10);
    } else {
      doc.setFont('helvetica', 'italic');
      doc.setTextColor(148, 163, 184);
      doc.text(`[Vị trí cảm biến: ${cubicleSensorLocation}]`, startX1 + 15, startY1 + 22);
    }

    // Slot 2: Thermal IR
    doc.rect(startX2, startY1, cellW, cellH);
    doc.setFontSize(7.5);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(51, 65, 85);
    doc.text(`2. ẢNH NHIỆT HỒNG NGOẠI (ΔT = ${thermalDeltaT}°C)`, startX2 + 2, startY1 + 4);
    if (thermalImage) {
      doc.addImage(thermalImage, 'JPEG', startX2 + 2, startY1 + 6, cellW - 4, cellH - 10);
    } else {
      doc.setFont('helvetica', 'italic');
      doc.setTextColor(148, 163, 184);
      doc.text(`[Tmax: ${thermalMaxTemp}°C • Tref: ${thermalRefTemp}°C]`, startX2 + 15, startY1 + 22);
    }

    // Slot 3: Original PRPD Image
    doc.rect(startX1, startY2, cellW, cellH);
    doc.setFontSize(7.5);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(51, 65, 85);
    doc.text('3. ẢNH ĐỒ THỊ PRPD THỰC TẾ (TỪ THIẾT BỊ ĐO)', startX1 + 2, startY2 + 4);
    if (prpdImage) {
      doc.addImage(prpdImage, 'PNG', startX1 + 2, startY2 + 6, cellW - 4, cellH - 10);
    } else {
      doc.setFont('helvetica', 'italic');
      doc.setTextColor(148, 163, 184);
      doc.text('[Chụp ảnh màn hình UltraTEV Plus²]', startX1 + 15, startY2 + 22);
    }

    // Slot 4: Canvas Simulated PRPD Plot
    doc.rect(startX2, startY2, cellW, cellH);
    doc.setFontSize(7.5);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(51, 65, 85);
    doc.text(`4. MÔ PHỎNG ĐÁM MÂY PRPD (${simDefectType})`, startX2 + 2, startY2 + 4);
    if (prpdCanvasRef.current) {
      const canvasImg = prpdCanvasRef.current.toDataURL('image/png');
      doc.addImage(canvasImg, 'PNG', startX2 + 2, startY2 + 6, cellW - 4, cellH - 10);
    }

    currentY = startY2 + cellH + 8;

    // Section 3: Diagnostic Results & Sensor Compatibility Matrix
    doc.setFontSize(10);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(15, 23, 42);
    doc.text('MỤC III: KẾT QUẢ CHẨN ĐOÁN CHUYÊN SÂU & HỆ SỐ TÌNH TRẠNG CBRM', 14, currentY);

    const diagTable = [
      [
        `Trạng thái: ${diagnosticResult.statusLabel}`,
        `CBRM Factor: ${diagnosticResult.cbrmConditionFactor} • Collar: ${diagnosticResult.cbrmCollar}`,
        `Dạng lỗi: ${simDefectType}`
      ],
      [
        `TEV cực đại: ${simPeakMagnitude} dBmV`,
        `Siêu âm cực đại: ${currentPayload.pd_measurements.maxUltrasonic_dBuV} dBµV`,
        `Phát hiện: EM-TEV = ${diagnosticResult.sensorCompatibility.emTevUhfVpisHfct ? 'YES' : 'NO'} | US = ${diagnosticResult.sensorCompatibility.ultrasonic ? 'YES' : 'NO'}`
      ]
    ];

    autoTable(doc, {
      startY: currentY + 3,
      body: diagTable,
      theme: 'plain',
      styles: { fontSize: 8, cellPadding: 2, textColor: [30, 41, 59], fontStyle: 'bold' }
    });

    currentY = (doc as any).lastAutoTable?.finalY || currentY + 20;

    // Action Plan Box
    doc.setFillColor(248, 250, 252);
    doc.rect(14, currentY + 2, pageWidth - 28, 26, 'F');
    doc.setDrawColor(203, 213, 225);
    doc.rect(14, currentY + 2, pageWidth - 28, 26);

    doc.setFontSize(8.5);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(185, 28, 28);
    doc.text(`TIẾN ĐỘ KHẮC PHỤC: ${diagnosticResult.actionPlan.urgencyTimelineVn}`, 18, currentY + 8);

    doc.setFontSize(7.5);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(51, 65, 85);
    diagnosticResult.actionPlan.netaActionSteps.slice(0, 3).forEach((step, idx) => {
      doc.text(`• ${step}`, 18, currentY + 14 + idx * 4);
    });

    // Signature Area
    const sigY = 262;
    doc.setDrawColor(203, 213, 225);
    doc.line(14, sigY, pageWidth - 14, sigY);

    doc.setFontSize(8);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(15, 23, 42);
    doc.text('CHUYÊN GIA CHẨN ĐOÁN TEV COE', 20, sigY + 6);
    doc.text('KỸ SƯ KHẢO SÁT HIỆN TRƯỜNG', 85, sigY + 6);
    doc.text('ĐẠI DIỆN QUẢN LÝ VẬN HÀNH', 150, sigY + 6);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.setTextColor(100, 116, 139);
    doc.text('Võ Minh Tú (Senior Specialist)', 20, sigY + 20);
    doc.text(inspectorName, 85, sigY + 20);
    doc.text(customerName, 150, sigY + 20);

    const fileName = `TEV_COE_PD_Report_${cubicleId.replace(/[^a-zA-Z0-9]/g, '_')}_${surveyDate}.pdf`;
    doc.save(fileName);
  };

  // --------------------------------------------------------------------------
  // EXPORT SUBSTATION SUMMARY CSV & PDF (Phân hệ 5: Bảng Tổng Hợp Trạm)
  // --------------------------------------------------------------------------
  const handleExportSummaryCsv = () => {
    if (cubicleHistory.length === 0) return;
    const headers = ['STT', 'Mã Ngăn Lộ', 'Ngày Khảo Sát', 'Dạng Phóng Điện', 'TEV (dBmV)', 'Siêu Âm (dBµV)', 'Nhiệt ΔT (°C)', 'Mức Rủi Ro', 'Kỹ Sư FSE'];
    const rows = cubicleHistory.map((rec, index) => [
      index + 1,
      `"${rec.cubicleId.replace(/"/g, '""')}"`,
      rec.date,
      rec.defect,
      rec.tev,
      rec.us,
      rec.deltaT,
      rec.status,
      `"${rec.inspector.replace(/"/g, '""')}"`
    ]);
    const csvContent = '\uFEFF' + [headers.join(','), ...rows.map(r => r.join(','))].join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Bang_Tong_Hop_PD_${substationName.replace(/[^a-zA-Z0-9]/g, '_')}_${surveyDate}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  };

  const handleExportSummaryPdf = () => {
    if (cubicleHistory.length === 0) return;
    const doc = new jsPDF('l', 'mm', 'a4'); // Landscape for rich table layout
    const pageWidth = doc.internal.pageSize.getWidth();

    // TEV Corporate Header
    doc.setFillColor(15, 23, 42); // slate-900
    doc.rect(0, 0, pageWidth, 28, 'F');

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(14);
    doc.setFont('helvetica', 'bold');
    doc.text('CÔNG TY CỔ PHẦN KỸ NGHỆ NĂNG LƯỢNG VIỆT NAM (VIETNAM ENERGY ENGINEERING JSC)', 14, 11);

    doc.setFontSize(8);
    doc.setFont('helvetica', 'normal');
    doc.text('TEV CENTER OF EXCELLENCE (COE) • BẢNG TỔNG HỢP KHẢO SÁT PHÓNG ĐIỆN CỤC BỘ (PD) & NHIỆT ẢNH TRẠM', 14, 17);
    doc.text(`TRẠM BIẾN ÁP: ${substationName.toUpperCase()} • NGÀY BÁO CÁO: ${surveyDate} • MÃ DỰ ÁN: ${projectCode}`, 14, 23);

    // Title
    doc.setTextColor(15, 23, 42);
    doc.setFontSize(13);
    doc.setFont('helvetica', 'bold');
    doc.text('BẢNG TỔNG HỢP ĐÁNH GIÁ TÌNH TRẠNG CÁC NGĂN LỘ THEO EA TECHNOLOGY & NETA MTS', 14, 36);

    // KPI Summary
    const totalCount = cubicleHistory.length;
    const criticalCount = cubicleHistory.filter(r => r.status === 'CRITICAL').length;
    const redCount = cubicleHistory.filter(r => r.status === 'HIGH' || r.status === 'RED').length;
    const amberCount = cubicleHistory.filter(r => r.status === 'MEDIUM' || r.status === 'AMBER').length;
    const greenCount = cubicleHistory.filter(r => r.status === 'LOW' || r.status === 'GREEN').length;

    const kpiY = 41;
    const boxW = 50;
    const boxH = 14;

    // Total
    doc.setFillColor(241, 245, 249);
    doc.rect(14, kpiY, boxW, boxH, 'F');
    doc.setFontSize(7.5);
    doc.setTextColor(71, 85, 105);
    doc.text('TỔNG SỐ NGĂN LỘ', 17, kpiY + 5);
    doc.setFontSize(12);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(15, 23, 42);
    doc.text(`${totalCount} tủ`, 17, kpiY + 11);

    // Normal Green
    doc.setFillColor(236, 253, 245);
    doc.rect(68, kpiY, boxW, boxH, 'F');
    doc.setFontSize(7.5);
    doc.setTextColor(5, 150, 105);
    doc.text('BÌNH THƯỜNG (XANH)', 71, kpiY + 5);
    doc.setFontSize(12);
    doc.setFont('helvetica', 'bold');
    doc.text(`${greenCount} tủ`, 71, kpiY + 11);

    // Amber
    doc.setFillColor(254, 243, 199);
    doc.rect(122, kpiY, boxW, boxH, 'F');
    doc.setFontSize(7.5);
    doc.setTextColor(217, 119, 6);
    doc.text('THEO DÕI (VÀNG)', 125, kpiY + 5);
    doc.setFontSize(12);
    doc.setFont('helvetica', 'bold');
    doc.text(`${amberCount} tủ`, 125, kpiY + 11);

    // High Red
    doc.setFillColor(254, 226, 226);
    doc.rect(176, kpiY, boxW, boxH, 'F');
    doc.setFontSize(7.5);
    doc.setTextColor(220, 38, 38);
    doc.text('NGHIÊM TRỌNG (ĐỎ)', 179, kpiY + 5);
    doc.setFontSize(12);
    doc.setFont('helvetica', 'bold');
    doc.text(`${redCount} tủ`, 179, kpiY + 11);

    // Critical Purple
    doc.setFillColor(243, 232, 255);
    doc.rect(230, kpiY, boxW, boxH, 'F');
    doc.setFontSize(7.5);
    doc.setTextColor(147, 51, 234);
    doc.text('KHẨN CẤP / CẮT ĐIỆN (TÍM)', 233, kpiY + 5);
    doc.setFontSize(12);
    doc.setFont('helvetica', 'bold');
    doc.text(`${criticalCount} tủ`, 233, kpiY + 11);

    // Table Data
    const tableBody = cubicleHistory.map((rec, index) => {
      let timeline = 'Giám sát 12 tháng';
      if (rec.status === 'CRITICAL') timeline = 'Khẩn cấp < 24-48 giờ';
      else if (rec.status === 'HIGH' || rec.status === 'RED') timeline = 'Cắt điện trong 15-30 ngày';
      else if (rec.status === 'MEDIUM' || rec.status === 'AMBER') timeline = 'Rút ngắn chu kỳ 3-6 tháng';

      return [
        index + 1,
        rec.cubicleId,
        rec.date,
        rec.defect,
        `${rec.tev} dBmV`,
        `${rec.us} dBµV`,
        `${rec.deltaT}°C`,
        rec.status,
        timeline,
        rec.inspector
      ];
    });

    autoTable(doc, {
      startY: 59,
      head: [['STT', 'Mã Ngăn Lộ / Tủ Điện', 'Ngày Đo', 'Dạng Phóng Điện (EA Tech)', 'TEV Cực Đại', 'Siêu Âm', 'Nhiệt ΔT', 'Mức Rủi Ro', 'Thời Hạn Xử Lý', 'Kỹ Sư FSE']],
      body: tableBody,
      theme: 'grid',
      styles: { fontSize: 8, cellPadding: 2.5, textColor: [30, 41, 59] },
      headStyles: { fillColor: [15, 23, 42], textColor: [255, 255, 255], fontStyle: 'bold' },
      didParseCell: (data) => {
        if (data.section === 'body' && data.column.index === 7) {
          const val = String(data.cell.raw);
          if (val === 'CRITICAL') {
            data.cell.styles.fillColor = [243, 232, 255];
            data.cell.styles.textColor = [126, 34, 206];
            data.cell.styles.fontStyle = 'bold';
          } else if (val === 'HIGH' || val === 'RED') {
            data.cell.styles.fillColor = [254, 226, 226];
            data.cell.styles.textColor = [185, 28, 28];
            data.cell.styles.fontStyle = 'bold';
          } else if (val === 'MEDIUM' || val === 'AMBER') {
            data.cell.styles.fillColor = [254, 243, 199];
            data.cell.styles.textColor = [180, 83, 9];
            data.cell.styles.fontStyle = 'bold';
          } else {
            data.cell.styles.fillColor = [236, 253, 245];
            data.cell.styles.textColor = [4, 120, 87];
            data.cell.styles.fontStyle = 'bold';
          }
        }
      }
    });

    const finalY = (doc as any).lastAutoTable?.finalY || 150;
    const sigY = Math.min(185, Math.max(finalY + 12, 160));

    doc.setDrawColor(203, 213, 225);
    doc.line(14, sigY, pageWidth - 14, sigY);

    doc.setFontSize(8);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(15, 23, 42);
    doc.text('KỸ SƯ KHẢO SÁT HIỆN TRƯỜNG (FSE)', 25, sigY + 6);
    doc.text('GIÁM SÁT KỸ THUẬT NHÀ MÁY / TRẠM', 115, sigY + 6);
    doc.text('ĐẠI DIỆN CHỦ ĐẦU TƯ / BAN VẬN HÀNH', 205, sigY + 6);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.5);
    doc.setTextColor(100, 116, 139);
    doc.text(inspectorName, 25, sigY + 16);
    doc.text('Kỹ sư Giám sát Trạm', 115, sigY + 16);
    doc.text(customerName, 205, sigY + 16);

    const fileName = `Bang_Tong_Hop_PD_${substationName.replace(/[^a-zA-Z0-9]/g, '_')}_${surveyDate}.pdf`;
    doc.save(fileName);
  };

  const handleLoadRecordIntoForm = (rec: HistoryRecord) => {
    setCubicleId(rec.cubicleId);
    setSimDefectType(rec.defect);
    setSimPeakMagnitude(rec.tev);
    setThermalMaxTemp(32 + rec.deltaT);
    setThermalRefTemp(32);
    setSurveyDate(rec.date);
    setSaveSuccessMessage(`Đã nạp số liệu ngăn lộ [${rec.cubicleId}] lên form phân tích!`);
    setTimeout(() => setSaveSuccessMessage(null), 3500);
  };

  const handleDeleteRecord = (id: string) => {
    setCubicleHistory(prev => prev.filter(r => r.id !== id));
  };

  const handleClearAllHistory = () => {
    if (window.confirm('Bạn có chắc chắn muốn xóa toàn bộ danh sách lịch sử đo trạm này?')) {
      setCubicleHistory([]);
    }
  };

  const handleResetDefaultHistory = () => {
    setCubicleHistory(DEFAULT_HISTORY);
    setSaveSuccessMessage('Đã nạp mẫu 5 ngăn lộ trạm điện hoàn chỉnh!');
    setTimeout(() => setSaveSuccessMessage(null), 3000);
  };

  return (
    <div className="space-y-6">
      {/* -------------------------------------------------------------------- */}
      {/* 1. TOP HEADER & ACTION BAR */}
      {/* -------------------------------------------------------------------- */}
      <div className="bg-gradient-to-r from-slate-950 via-slate-900 to-indigo-950 text-white p-5 sm:p-6 rounded-2xl shadow-xl border border-indigo-500/30">
        <div className="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-4">
          <div className="space-y-1.5">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="px-2.5 py-1 rounded-md text-[11px] font-black uppercase tracking-wider bg-indigo-500/30 text-indigo-200 border border-indigo-400/30">
                TEV COE Platform • EA Technology UltraTEV Plus²
              </span>
              <span className="px-2 py-0.5 rounded-md text-[10px] font-bold bg-amber-500/20 text-amber-300 border border-amber-500/30">
                ANSI/NETA MTS-2025 Table 100.23
              </span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white flex items-center gap-2.5">
              <AudioWaveform size={24} className="text-indigo-400" />
              <span>Chẩn Đoán Đa Phương Thức Phóng Điện Cục Bộ (AI Vision PRPD + Nhiệt Ảnh IR)</span>
            </h1>
            <p className="text-xs sm:text-sm text-slate-300 max-w-4xl leading-relaxed">
              Tích hợp nhận dạng thị giác Gemini 2.5 Flash, bộ mô phỏng đám mây xung 360° theo bảng màu EA Technology,
              thuật toán cộng hưởng nhiệt ảnh IR &amp; độ ẩm, tự động lập Biên bản Khảo sát Hiện trường khổ A4 chuẩn TEV.
            </p>
          </div>

          <div className="flex items-center gap-2 flex-wrap w-full lg:w-auto">
            <button
              type="button"
              onClick={() => setShowPromptModal(true)}
              className="flex-1 lg:flex-none flex items-center justify-center gap-1.5 px-3.5 py-2 rounded-xl text-xs font-bold bg-slate-800 hover:bg-slate-700 text-indigo-300 border border-indigo-500/30 shadow-sm"
              title="Xem và sao chép Master System Instruction Prompt để dán vào Google AI Studio"
            >
              <Sparkles size={14} className="text-amber-400" />
              <span>Prompt AI Studio</span>
            </button>

            <button
              type="button"
              onClick={handleSaveToHistory}
              className="flex-1 lg:flex-none flex items-center justify-center gap-1.5 px-3.5 py-2 rounded-xl text-xs font-bold bg-slate-800 hover:bg-slate-700 text-emerald-300 border border-emerald-500/30 shadow-sm"
            >
              <Save size={14} />
              <span>Lưu Lịch Sử ({cubicleHistory.length})</span>
            </button>

            <button
              type="button"
              onClick={handleExportA4Pdf}
              className="w-full sm:w-auto flex items-center justify-center gap-2 px-4 py-2 rounded-xl text-xs font-black bg-gradient-to-r from-indigo-500 to-indigo-600 hover:from-indigo-400 hover:to-indigo-500 text-white shadow-lg shadow-indigo-600/30"
            >
              <Download size={15} />
              <span>Xuất Biên Bản PDF (A4)</span>
            </button>
          </div>
        </div>

        {/* 5 Preset Fast Scenarios Strip */}
        <div className="mt-5 pt-4 border-t border-slate-800/80 flex items-center gap-2 overflow-x-auto scrollbar-none pb-1">
          <span className="text-[11px] font-bold text-slate-400 shrink-0 flex items-center gap-1">
            <Sliders size={13} className="text-indigo-400" />
            <span>Kịch bản mẫu:</span>
          </span>
          {PRESET_SCENARIOS.map((p) => (
            <button
              key={p.id}
              type="button"
              onClick={() => handleLoadPreset(p)}
              className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all shrink-0 border flex items-center gap-1.5 ${
                simDefectType === p.defectType
                  ? 'bg-indigo-600 text-white border-indigo-400 shadow-sm'
                  : 'bg-slate-900/90 text-slate-300 hover:text-white hover:bg-slate-800 border-slate-800'
              }`}
            >
              <span>{p.name.split(' (')[0]}</span>
              <span className={`text-[9px] px-1.5 py-0.2 rounded font-black ${p.badgeColor}`}>
                {p.defectType}
              </span>
            </button>
          ))}
        </div>
      </div>

      {saveSuccessMessage && (
        <div className="p-3.5 bg-emerald-500/10 border border-emerald-500/30 rounded-xl text-xs font-bold text-emerald-800 flex items-center gap-2">
          <CheckCircle2 size={16} className="text-emerald-600 shrink-0" />
          <span>{saveSuccessMessage}</span>
        </div>
      )}

      {/* -------------------------------------------------------------------- */}
      {/* 2. MAIN 2-COLUMN WORKSPACE: LEFT (FORM & MEDIA), RIGHT (CANVAS & DIAG) */}
      {/* -------------------------------------------------------------------- */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* LEFT COLUMN: PHÂN HỆ 1 (METADATA) + PHÂN HỆ 2 (3-SLOT MEDIA) */}
        <div className="lg:col-span-5 space-y-6">
          {/* PHÂN HỆ 2: 3-SLOT MEDIA UPLOADER */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-black text-slate-900 flex items-center gap-2">
                <Camera size={16} className="text-indigo-600" />
                <span>PHÂN HỆ 2: TRÌNH QUẢN LÝ 3 NGUỒN ẢNH HIỆN TRƯỜNG</span>
              </h2>
              <span className="text-[10px] font-bold text-slate-500">Kéo thả / Chọn file / Ctrl+V</span>
            </div>

            {/* 3 SLOTS */}
            <div className="space-y-4">
              {/* SLOT 1: PRPD PLOT IMAGE */}
              <div className="p-3.5 rounded-xl border border-indigo-200 bg-indigo-50/30 space-y-2">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-bold text-indigo-950 flex items-center gap-1.5">
                    <AudioWaveform size={14} className="text-indigo-600" />
                    <span>Slot 1: Ảnh Đồ thị PRPD (Bắt buộc cho AI Vision)</span>
                  </span>
                  {prpdImage && (
                    <button
                      type="button"
                      onClick={() => setPrpdImage(null)}
                      className="text-[11px] text-rose-600 hover:underline font-bold"
                    >
                      Xóa ảnh
                    </button>
                  )}
                </div>

                <div
                  onPaste={(e) => handleImagePaste(e, setPrpdImage)}
                  className="relative border-2 border-dashed border-indigo-300 rounded-xl p-3 text-center hover:bg-indigo-50/60 transition-all bg-white"
                >
                  {prpdImage ? (
                    <div className="space-y-2">
                      <img src={prpdImage} alt="PRPD Plot" className="max-h-36 mx-auto rounded-lg shadow-sm object-contain" />
                      <button
                        type="button"
                        onClick={handleRunAiVision}
                        disabled={isAiVisionRunning}
                        className="w-full flex items-center justify-center gap-2 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs font-black shadow-sm disabled:opacity-50"
                      >
                        <Sparkles size={14} className="text-amber-300" />
                        <span>{isAiVisionRunning ? 'Đang phân tích thị giác...' : 'Phân tích PRPD bằng AI Vision'}</span>
                      </button>
                    </div>
                  ) : (
                    <label className="cursor-pointer block py-4">
                      <Upload size={24} className="mx-auto text-indigo-400 mb-1" />
                      <span className="text-xs font-bold text-indigo-900 block">Tải lên hoặc Click rồi bấm Ctrl+V để dán ảnh PRPD</span>
                      <span className="text-[10px] text-slate-500">Hỗ trợ ảnh chụp màn hình UltraTEV Plus²</span>
                      <input
                        type="file"
                        accept="image/*"
                        className="hidden"
                        onChange={(e) => {
                          const file = e.target.files?.[0];
                          if (file) {
                            const reader = new FileReader();
                            reader.onload = () => setPrpdImage(reader.result as string);
                            reader.readAsDataURL(file);
                          }
                        }}
                      />
                    </label>
                  )}
                </div>

                {/* AI Vision Result Pill */}
                {aiVisionResult && (
                  <div className="p-2.5 rounded-lg bg-emerald-50 border border-emerald-200 text-xs text-emerald-900 space-y-1">
                    <div className="flex items-center justify-between font-bold">
                      <span>AI Vision phát hiện: {aiVisionResult.primaryPDType}</span>
                      <span className="text-[10px] px-1.5 py-0.2 rounded bg-emerald-200 text-emerald-950 font-black">
                        Độ tin cậy {aiVisionResult.confidenceScore}%
                      </span>
                    </div>
                    <p className="text-[11px] text-slate-700">{aiVisionResult.visualAnalysisDetails}</p>
                  </div>
                )}
              </div>

              {/* SLOT 2: CUBICLE PHOTO & SENSOR POSITION */}
              <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/50 space-y-2">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                    <ImageIcon size={14} className="text-slate-600" />
                    <span>Slot 2: Ảnh Mặt trước Tủ điện / Điểm đo</span>
                  </span>
                  {cubicleImage && (
                    <button
                      type="button"
                      onClick={() => setCubicleImage(null)}
                      className="text-[11px] text-rose-600 hover:underline font-bold"
                    >
                      Xóa ảnh
                    </button>
                  )}
                </div>

                <div
                  onPaste={(e) => handleImagePaste(e, setCubicleImage)}
                  className="relative border-2 border-dashed border-slate-300 rounded-xl p-2.5 text-center bg-white"
                >
                  {cubicleImage ? (
                    <img src={cubicleImage} alt="Cubicle Front" className="max-h-28 mx-auto rounded-lg shadow-sm object-contain" />
                  ) : (
                    <label className="cursor-pointer block py-2.5">
                      <Camera size={20} className="mx-auto text-slate-400 mb-1" />
                      <span className="text-xs font-bold text-slate-700 block">Kéo thả / Chọn ảnh mặt tủ hoặc Ctrl+V</span>
                      <input
                        type="file"
                        accept="image/*"
                        className="hidden"
                        onChange={(e) => {
                          const file = e.target.files?.[0];
                          if (file) {
                            const reader = new FileReader();
                            reader.onload = () => setCubicleImage(reader.result as string);
                            reader.readAsDataURL(file);
                          }
                        }}
                      />
                    </label>
                  )}
                </div>

                {/* Vị trí đặt cảm biến */}
                <div className="grid grid-cols-2 gap-2 pt-1 text-xs">
                  <div>
                    <label className="text-[10px] font-bold text-slate-600 block mb-1">Vị trí gắn cảm biến:</label>
                    <select
                      value={cubicleSensorLocation}
                      onChange={(e) => setCubicleSensorLocation(e.target.value as any)}
                      className="w-full px-2 py-1 bg-white border border-slate-200 rounded-lg text-xs font-bold"
                    >
                      <option value="Cable Box">Cable Box (Khoang Cáp)</option>
                      <option value="Middle">Middle (Khoang Máy Cắt)</option>
                      <option value="Top">Top (Khoang Thanh Cái)</option>
                      <option value="Bottom">Bottom (Đáy Tủ/Tiếp Địa)</option>
                    </select>
                  </div>
                  <div>
                    <label className="text-[10px] font-bold text-slate-600 block mb-1">Ghi chú ảnh hiện trường:</label>
                    <input
                      type="text"
                      value={cubicleCaption}
                      onChange={(e) => setCubicleCaption(e.target.value)}
                      className="w-full px-2 py-1 bg-white border border-slate-200 rounded-lg text-xs"
                      placeholder="Ghi chú vị trí đo..."
                    />
                  </div>
                </div>
              </div>

              {/* SLOT 3: THERMAL IR IMAGE */}
              <div className="p-3.5 rounded-xl border border-rose-200 bg-rose-50/30 space-y-2">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-bold text-rose-950 flex items-center gap-1.5">
                    <Flame size={14} className="text-rose-600" />
                    <span>Slot 3: Ảnh Nhiệt Hồng ngoại (Thermal IR Image)</span>
                  </span>
                  {thermalImage && (
                    <button
                      type="button"
                      onClick={() => setThermalImage(null)}
                      className="text-[11px] text-rose-600 hover:underline font-bold"
                    >
                      Xóa ảnh
                    </button>
                  )}
                </div>

                <div
                  onPaste={(e) => handleImagePaste(e, setThermalImage)}
                  className="relative border-2 border-dashed border-rose-300 rounded-xl p-2.5 text-center bg-white"
                >
                  {thermalImage ? (
                    <div className="space-y-2">
                      <img src={thermalImage} alt="Thermal IR" className="max-h-28 mx-auto rounded-lg shadow-sm object-contain" />
                      <button
                        type="button"
                        onClick={handleRunThermalOcr}
                        disabled={isThermalOcrRunning}
                        className="w-full flex items-center justify-center gap-1.5 py-1.5 bg-rose-600 hover:bg-rose-700 text-white rounded-lg text-xs font-bold shadow-sm disabled:opacity-50"
                      >
                        <Thermometer size={14} />
                        <span>{isThermalOcrRunning ? 'Đang đọc nhiệt độ...' : 'Đọc Nhiệt độ từ Ảnh IR (AI Thermal OCR)'}</span>
                      </button>
                    </div>
                  ) : (
                    <label className="cursor-pointer block py-2.5">
                      <Thermometer size={20} className="mx-auto text-rose-400 mb-1" />
                      <span className="text-xs font-bold text-rose-900 block">Tải ảnh nhiệt FLIR/Fluke hoặc Ctrl+V</span>
                      <input
                        type="file"
                        accept="image/*"
                        className="hidden"
                        onChange={(e) => {
                          const file = e.target.files?.[0];
                          if (file) {
                            const reader = new FileReader();
                            reader.onload = () => setThermalImage(reader.result as string);
                            reader.readAsDataURL(file);
                          }
                        }}
                      />
                    </label>
                  )}
                </div>
              </div>
            </div>
          </div>

          {/* PHÂN HỆ 1: METADATA FORM */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="border-b border-slate-100 pb-3">
              <h2 className="text-sm font-black text-slate-900 flex items-center gap-2">
                <FileText size={16} className="text-indigo-600" />
                <span>PHÂN HỆ 1: FORM ĐỊNH DANH TỦ &amp; THAM SỐ VẬN HÀNH</span>
              </h2>
            </div>

            {/* Group 1: Project & Substation */}
            <div className="space-y-2.5">
              <h3 className="text-xs font-black text-slate-700 uppercase tracking-wider">1. Dự án &amp; Trạm Điện</h3>
              <div className="grid grid-cols-2 gap-2 text-xs">
                <div>
                  <label className="text-[10px] text-slate-500 block mb-0.5">Mã Biên Bản</label>
                  <input
                    type="text"
                    value={reportNo}
                    onChange={(e) => setReportNo(e.target.value)}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg font-mono font-bold"
                  />
                </div>
                <div>
                  <label className="text-[10px] text-slate-500 block mb-0.5">Ngày Khảo Sát</label>
                  <input
                    type="date"
                    value={surveyDate}
                    onChange={(e) => setSurveyDate(e.target.value)}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg font-bold"
                  />
                </div>
                <div className="col-span-2">
                  <label className="text-[10px] text-slate-500 block mb-0.5">Khách Hàng / Nhà Máy</label>
                  <input
                    type="text"
                    value={customerName}
                    onChange={(e) => setCustomerName(e.target.value)}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg font-bold"
                  />
                </div>
                <div>
                  <label className="text-[10px] text-slate-500 block mb-0.5">Trạm / Phòng Phân Phối</label>
                  <input
                    type="text"
                    value={substationName}
                    onChange={(e) => setSubstationName(e.target.value)}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg"
                  />
                </div>
                <div>
                  <label className="text-[10px] text-slate-500 block mb-0.5">Kỹ Sư Hiện Trường (FSE)</label>
                  <input
                    type="text"
                    value={inspectorName}
                    onChange={(e) => setInspectorName(e.target.value)}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg font-bold text-indigo-900"
                  />
                </div>
              </div>
            </div>

            {/* Group 2: Cubicle & Test Point */}
            <div className="space-y-2.5 pt-2 border-t border-slate-100">
              <h3 className="text-xs font-black text-slate-700 uppercase tracking-wider">2. Định danh Tủ &amp; Điểm Đo</h3>
              <div className="grid grid-cols-2 gap-2 text-xs">
                <div>
                  <label className="text-[10px] text-slate-500 block mb-0.5">Mã Ngăn Lộ (Cubicle ID)</label>
                  <input
                    type="text"
                    value={cubicleId}
                    onChange={(e) => setCubicleId(e.target.value)}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg font-bold"
                  />
                </div>
                <div>
                  <label className="text-[10px] text-slate-500 block mb-0.5">Loại Tủ Điện</label>
                  <select
                    value={cubicleType}
                    onChange={(e) => setCubicleType(e.target.value)}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg"
                  >
                    <option value="Metal-Clad">Metal-Clad</option>
                    <option value="Metal-Enclosed">Metal-Enclosed</option>
                    <option value="GIS">GIS</option>
                    <option value="RMU">RMU</option>
                    <option value="Cable Termination Box">Cable Termination Box</option>
                  </select>
                </div>
                <div>
                  <label className="text-[10px] text-slate-500 block mb-0.5">Điện Áp Định Mức / Đo</label>
                  <select
                    value={ratedVoltage}
                    onChange={(e) => setRatedVoltage(e.target.value)}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg font-bold"
                  >
                    <option value="3.3 kV">3.3 kV</option>
                    <option value="6.6 kV">6.6 kV</option>
                    <option value="11 kV">11 kV</option>
                    <option value="15 kV">15 kV</option>
                    <option value="22 kV">22 kV</option>
                    <option value="35 kV">35 kV</option>
                  </select>
                </div>
                <div>
                  <label className="text-[10px] text-slate-500 block mb-0.5">Dòng Tải Lúc Đo (A / %)</label>
                  <div className="flex gap-1.5">
                    <input
                      type="number"
                      value={loadCurrent}
                      onChange={(e) => setLoadCurrent(Number(e.target.value))}
                      className="w-20 px-2 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs"
                    />
                    <input
                      type="number"
                      value={loadPercentage}
                      onChange={(e) => setLoadPercentage(Number(e.target.value))}
                      className="w-16 px-2 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-xs font-bold"
                      placeholder="%"
                    />
                  </div>
                </div>
              </div>
            </div>

            {/* Group 3: Environment, Thermography & Sensory */}
            <div className="space-y-2.5 pt-2 border-t border-slate-100">
              <h3 className="text-xs font-black text-slate-700 uppercase tracking-wider">3. Môi Trường, Nhiệt Ảnh &amp; Cảm Quan</h3>
              <div className="grid grid-cols-2 gap-2 text-xs">
                <div>
                  <label className="text-[10px] text-slate-500 block mb-0.5">Nhiệt độ phòng (°C)</label>
                  <input
                    type="number"
                    value={ambientTemp}
                    onChange={(e) => setAmbientTemp(Number(e.target.value))}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg font-bold"
                  />
                </div>
                <div>
                  <label className="text-[10px] text-slate-500 block mb-0.5">Độ ẩm tương đối (% RH)</label>
                  <input
                    type="number"
                    value={relativeHumidity}
                    onChange={(e) => setRelativeHumidity(Number(e.target.value))}
                    className={`w-full px-2.5 py-1.5 bg-slate-50 border rounded-lg font-bold ${
                      relativeHumidity > 70 ? 'border-amber-400 text-amber-900 bg-amber-50/50' : 'border-slate-200'
                    }`}
                  />
                </div>
                <div>
                  <label className="text-[10px] text-slate-500 block mb-0.5">Nhiệt độ T_max (°C)</label>
                  <input
                    type="number"
                    value={thermalMaxTemp}
                    onChange={(e) => setThermalMaxTemp(Number(e.target.value))}
                    className="w-full px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg font-bold text-rose-700"
                  />
                </div>
                <div>
                  <label className="text-[10px] text-slate-500 block mb-0.5">
                    Độ chênh ΔT = Tmax - Tref (°C)
                  </label>
                  <div className={`px-2.5 py-1.5 rounded-lg font-black text-xs border flex items-center justify-between ${
                    thermalDeltaT >= 10 ? 'bg-rose-100 border-rose-300 text-rose-900' : 'bg-slate-100 border-slate-200 text-slate-800'
                  }`}>
                    <span>{thermalDeltaT} °C</span>
                    <span className="text-[9px] font-bold text-slate-500">(Tref: {thermalRefTemp}°C)</span>
                  </div>
                </div>
              </div>

              {/* Sensory Checkboxes */}
              <div className="pt-1.5 grid grid-cols-2 gap-2 text-xs text-slate-700">
                <label className="flex items-center gap-1.5 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={sensoryChecks.ozoneSmell}
                    onChange={(e) => setSensoryChecks(prev => ({ ...prev, ozoneSmell: e.target.checked }))}
                    className="rounded text-indigo-600"
                  />
                  <span>Mùi Ozone (Smell)</span>
                </label>
                <label className="flex items-center gap-1.5 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={sensoryChecks.audibleSound}
                    onChange={(e) => setSensoryChecks(prev => ({ ...prev, audibleSound: e.target.checked }))}
                    className="rounded text-indigo-600"
                  />
                  <span>Tiếng rít lách tách</span>
                </label>
                <label className="flex items-center gap-1.5 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={sensoryChecks.heatAnomaly}
                    onChange={(e) => setSensoryChecks(prev => ({ ...prev, heatAnomaly: e.target.checked }))}
                    className="rounded text-indigo-600"
                  />
                  <span>Phát nhiệt bất thường</span>
                </label>
                <label className="flex items-center gap-1.5 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={sensoryChecks.visualTracking}
                    onChange={(e) => setSensoryChecks(prev => ({ ...prev, visualTracking: e.target.checked }))}
                    className="rounded text-indigo-600"
                  />
                  <span>Bột trắng / Vết rãnh</span>
                </label>
              </div>
            </div>
          </div>
        </div>

        {/* RIGHT COLUMN: PHÂN HỆ 3 (CANVAS SIMULATOR) + PHÂN HỆ 4 (CROSS-MODALITY) */}
        <div className="lg:col-span-7 space-y-6">
          {/* PHÂN HỆ 3: INTERACTIVE PRPD CLOUD & ULTRASONIC SIMULATOR (HTML5 CANVAS) */}
          <div className="bg-slate-950 p-5 rounded-2xl border border-slate-800 text-white shadow-xl space-y-4">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <div className="flex items-center gap-2">
                <AudioWaveform size={18} className="text-cyan-400" />
                <h2 className="text-sm font-black text-white">
                  PHÂN HỆ 3: BỘ MÔ PHỎNG ĐÁM MÂY PRPD &amp; ĐƯỜNG BAO SIÊU ÂM (HTML5 CANVAS)
                </h2>
              </div>
              <span className="text-[10px] font-mono px-2 py-0.5 rounded bg-cyan-950 border border-cyan-800 text-cyan-300 font-bold">
                {simDefectType}
              </span>
            </div>

            {/* Canvas 1: PRPD Density Cloud */}
            <div className="space-y-1.5">
              <div className="flex items-center justify-between text-xs text-slate-300">
                <span className="font-bold flex items-center gap-1.5">
                  <span>1. Đồ thị Đám mây PRPD (0° - 360° theo Phase Angle &amp; dBmV):</span>
                </span>
                {/* EA Tech Legend */}
                <div className="flex items-center gap-2 text-[10px] font-mono">
                  <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-emerald-500"></span> &lt;2</span>
                  <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-blue-500"></span> 2-6</span>
                  <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-red-500"></span> 6-12</span>
                  <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-amber-900"></span> &gt;12</span>
                </div>
              </div>

              <div className="rounded-xl overflow-hidden border border-slate-800 shadow-inner bg-slate-900/90">
                <canvas
                  ref={prpdCanvasRef}
                  width={680}
                  height={220}
                  className="w-full h-44 sm:h-52 block"
                />
              </div>
            </div>

            {/* Canvas 2: Ultrasonic Envelope Waveform */}
            <div className="space-y-1.5 pt-1">
              <div className="flex items-center justify-between text-xs text-slate-300">
                <span className="font-bold">2. Đường bao Sóng Siêu âm (Ultrasonic Envelope &amp; Miền Thời gian):</span>
                <span className="text-[10px] text-slate-400">Tần số đồng bộ 50 Hz</span>
              </div>
              <div className="rounded-xl overflow-hidden border border-slate-800 bg-slate-900/90">
                <canvas
                  ref={ultrasonicCanvasRef}
                  width={680}
                  height={75}
                  className="w-full h-16 sm:h-20 block"
                />
              </div>
            </div>

            {/* Simulator Controls & Sliders */}
            <div className="pt-2 border-t border-slate-800/80 grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
              <div>
                <div className="flex justify-between text-[10px] text-slate-400 mb-1">
                  <span>Dạng Đám Mây:</span>
                </div>
                <select
                  value={simDefectType}
                  onChange={(e) => setSimDefectType(e.target.value as any)}
                  className="w-full px-2 py-1.5 bg-slate-900 border border-slate-700 rounded-lg text-white font-bold"
                >
                  <option value="INTERNAL_VOID">INTERNAL_VOID</option>
                  <option value="SURFACE_DISCHARGE">SURFACE_DISCHARGE</option>
                  <option value="CORONA">CORONA</option>
                  <option value="POOR_CONTACT_FLOATING">POOR_CONTACT</option>
                  <option value="SPARK_GAP">SPARK_GAP</option>
                  <option value="NOISE">NOISE</option>
                </select>
              </div>

              <div>
                <div className="flex justify-between text-[10px] text-slate-400 mb-1">
                  <span>Biên độ Max:</span>
                  <span className="font-bold text-cyan-300">{simPeakMagnitude} dB</span>
                </div>
                <input
                  type="range"
                  min="5"
                  max="60"
                  value={simPeakMagnitude}
                  onChange={(e) => setSimPeakMagnitude(Number(e.target.value))}
                  className="w-full accent-cyan-400"
                />
              </div>

              <div>
                <div className="flex justify-between text-[10px] text-slate-400 mb-1">
                  <span>Xung/Chu kỳ (PPC):</span>
                  <span className="font-bold text-cyan-300">{simPPC}</span>
                </div>
                <input
                  type="range"
                  min="2"
                  max="40"
                  value={simPPC}
                  onChange={(e) => setSimPPC(Number(e.target.value))}
                  className="w-full accent-cyan-400"
                />
              </div>

              <div>
                <div className="flex justify-between text-[10px] text-slate-400 mb-1">
                  <span>Lệch pha (Phase):</span>
                  <span className="font-bold text-cyan-300">{simPhaseShift}°</span>
                </div>
                <input
                  type="range"
                  min="-90"
                  max="90"
                  value={simPhaseShift}
                  onChange={(e) => setSimPhaseShift(Number(e.target.value))}
                  className="w-full accent-cyan-400"
                />
              </div>
            </div>
          </div>

          {/* PHÂN HỆ 4: CROSS-MODALITY SEVERITY ENGINE (PD + NHIỆT + MÔI TRƯỜNG) */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-black text-slate-900 flex items-center gap-2">
                <ShieldAlert size={16} className="text-rose-600" />
                <span>PHÂN HỆ 4: THANG ĐO NGHIÊM TRỌNG ĐA PHƯƠNG THỨC &amp; CỘNG HƯỞNG</span>
              </h2>
              <span className={`text-xs px-2.5 py-1 rounded-full font-black border ${
                diagnosticResult.overallStatus === 'CRITICAL'
                  ? 'bg-rose-100 border-rose-300 text-rose-900'
                  : diagnosticResult.overallStatus === 'RED'
                  ? 'bg-red-50 border-red-200 text-red-800'
                  : diagnosticResult.overallStatus === 'AMBER'
                  ? 'bg-amber-50 border-amber-200 text-amber-800'
                  : 'bg-emerald-50 border-emerald-200 text-emerald-800'
              }`}>
                {diagnosticResult.statusLabel}
              </span>
            </div>

            {/* EA Technology Condition Factors Strip */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5 text-xs">
              <div className="p-2.5 rounded-xl bg-slate-50 border border-slate-200">
                <span className="text-[10px] text-slate-500 block">CBRM Condition Factor</span>
                <span className="text-sm font-black text-indigo-950 font-mono">
                  {diagnosticResult.cbrmConditionFactor.toFixed(2)}
                </span>
              </div>
              <div className="p-2.5 rounded-xl bg-slate-50 border border-slate-200">
                <span className="text-[10px] text-slate-500 block">CBRM Collar</span>
                <span className="text-sm font-black text-indigo-950 font-mono">
                  {diagnosticResult.cbrmCollar.toFixed(1)}
                </span>
              </div>
              <div className="p-2.5 rounded-xl bg-slate-50 border border-slate-200">
                <span className="text-[10px] text-slate-500 block">Cảm biến Điện từ (EM)</span>
                <span className={`text-xs font-black ${
                  diagnosticResult.sensorCompatibility.emTevUhfVpisHfct ? 'text-emerald-700' : 'text-slate-400'
                }`}>
                  {diagnosticResult.sensorCompatibility.emTevUhfVpisHfct ? '✔ Tích cực' : '❌ Không nhạy'}
                </span>
              </div>
              <div className="p-2.5 rounded-xl bg-slate-50 border border-slate-200">
                <span className="text-[10px] text-slate-500 block">Cảm biến Siêu âm (US)</span>
                <span className={`text-xs font-black ${
                  diagnosticResult.sensorCompatibility.ultrasonic ? 'text-emerald-700' : 'text-slate-400'
                }`}>
                  {diagnosticResult.sensorCompatibility.ultrasonic ? '✔ Tích cực' : '❌ Không nhạy'}
                </span>
              </div>
            </div>

            {/* Sensor Rule Description */}
            <div className="p-3 bg-indigo-50/60 rounded-xl border border-indigo-100 text-xs text-indigo-950 flex items-start gap-2">
              <Info size={16} className="text-indigo-600 shrink-0 mt-0.5" />
              <div>
                <strong>Quy tắc nhận dạng cảm biến EA Technology:</strong>{' '}
                <span>{diagnosticResult.sensorCompatibility.ruleVn}</span>
              </div>
            </div>

            {/* Resonance Alerts (Cộng hưởng nhiệt ảnh + độ ẩm) */}
            {diagnosticResult.resonanceAlerts.length > 0 && (
              <div className="space-y-2">
                {diagnosticResult.resonanceAlerts.map((alert, idx) => (
                  <div
                    key={idx}
                    className="p-3.5 bg-rose-50 border border-rose-300 rounded-xl text-xs text-rose-950 space-y-1 shadow-xs"
                  >
                    <div className="font-black flex items-center gap-1.5 text-rose-900">
                      <AlertOctagon size={15} className="text-rose-600 shrink-0" />
                      <span>{alert.title}</span>
                    </div>
                    <p className="text-[11px] leading-relaxed text-slate-800">{alert.messageVn}</p>
                  </div>
                ))}
              </div>
            )}

            {/* Matrix Comparison Table */}
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="bg-slate-100 text-slate-700 text-[11px]">
                    <th className="p-2 border border-slate-200 font-bold">Tham Số Khảo Sát</th>
                    <th className="p-2 border border-slate-200 font-bold">Giá Trị Đo</th>
                    <th className="p-2 border border-slate-200 font-bold">Ngưỡng Chuẩn NETA &amp; EA Tech</th>
                    <th className="p-2 border border-slate-200 font-bold text-center">Đánh Giá</th>
                  </tr>
                </thead>
                <tbody>
                  {diagnosticResult.matrixComparison.map((row, idx) => (
                    <tr key={idx} className="hover:bg-slate-50/80">
                      <td className="p-2 border border-slate-200 font-semibold text-slate-800">
                        {row.parameterVn}
                      </td>
                      <td className="p-2 border border-slate-200 font-mono font-bold text-indigo-950">
                        {row.measuredValue}
                      </td>
                      <td className="p-2 border border-slate-200 text-slate-600 text-[11px]">
                        {row.netaThreshold}
                      </td>
                      <td className="p-2 border border-slate-200 text-center">
                        <span className={`px-2 py-0.5 rounded text-[10px] font-black ${
                          row.status === 'CRITICAL' || row.status === 'RED'
                            ? 'bg-rose-100 text-rose-800'
                            : row.status === 'AMBER'
                            ? 'bg-amber-100 text-amber-800'
                            : 'bg-emerald-100 text-emerald-800'
                        }`}>
                          {row.status}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* Action Steps */}
            <div className="p-3.5 bg-slate-900 text-white rounded-xl space-y-2">
              <div className="flex items-center justify-between text-xs">
                <span className="font-bold text-amber-300">
                  TIẾN ĐỘ XỬ LÝ KHUYẾN NGHỊ: {diagnosticResult.actionPlan.urgencyTimelineVn}
                </span>
                <span className="text-[10px] text-slate-400">ANSI/NETA MTS-2025</span>
              </div>
              <ul className="text-xs space-y-1 text-slate-300 list-disc list-inside">
                {diagnosticResult.actionPlan.netaActionSteps.map((step, idx) => (
                  <li key={idx} className="leading-relaxed">{step}</li>
                ))}
              </ul>
            </div>
          </div>
        </div>
      </div>

      {/* -------------------------------------------------------------------- */}
      {/* 3. MEASURED CUBICLES SUMMARY HISTORY TABLE (Bảng Tổng Hợp Trạm) */}
      {/* -------------------------------------------------------------------- */}
      <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
        <div className="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-3 border-b border-slate-100 pb-3">
          <div>
            <h3 className="text-sm font-black text-slate-900 flex items-center gap-2">
              <Layers size={17} className="text-indigo-600" />
              <span>BẢNG TỔNG HỢP CÁC NGĂN LỘ ĐÃ ĐO TRONG TRẠM (SUBSTATION SUMMARY TABLE)</span>
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Theo dõi toàn diện các tủ điện trong trạm {substationName} • Đánh giá rủi ro theo tiêu chuẩn EA Technology &amp; ANSI/NETA MTS
            </p>
          </div>
          
          <div className="flex items-center gap-2 flex-wrap self-stretch sm:self-auto">
            <button
              type="button"
              onClick={handleSaveToHistory}
              className="px-3 py-1.5 rounded-lg text-xs font-bold bg-indigo-600 hover:bg-indigo-700 text-white flex items-center gap-1.5 shadow-sm transition-all"
              title="Lưu ngăn lộ đang khảo sát vào bảng tổng hợp"
            >
              <Save size={13} />
              <span>Lưu Ngăn Lộ Hiện Tại</span>
            </button>

            <button
              type="button"
              onClick={handleExportSummaryCsv}
              className="px-3 py-1.5 rounded-lg text-xs font-bold bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 flex items-center gap-1.5 transition-all"
              title="Xuất bảng số liệu ra file Excel CSV"
            >
              <Download size={13} />
              <span>Xuất CSV</span>
            </button>

            <button
              type="button"
              onClick={handleExportSummaryPdf}
              className="px-3 py-1.5 rounded-lg text-xs font-bold bg-gradient-to-r from-emerald-600 to-teal-700 hover:from-emerald-500 hover:to-teal-600 text-white flex items-center gap-1.5 shadow-sm transition-all"
              title="Xuất báo cáo tổng hợp toàn trạm dạng A4 Landscape kèm chữ ký 3 bên"
            >
              <FileCheck2 size={13} />
              <span>Báo Cáo Tổng Hợp PDF</span>
            </button>

            <button
              type="button"
              onClick={handleResetDefaultHistory}
              className="px-2.5 py-1.5 rounded-lg text-xs font-medium text-slate-600 hover:text-slate-900 hover:bg-slate-100 border border-slate-200 flex items-center gap-1 transition-all"
              title="Nạp lại bộ dữ liệu 5 ngăn lộ trạm mẫu"
            >
              <RotateCcw size={12} />
              <span>Nạp Mẫu</span>
            </button>

            {cubicleHistory.length > 0 && (
              <button
                type="button"
                onClick={handleClearAllHistory}
                className="p-1.5 rounded-lg text-xs font-medium text-rose-600 hover:bg-rose-50 border border-rose-200 transition-all"
                title="Xóa toàn bộ danh sách"
              >
                <Trash2 size={13} />
              </button>
            )}
          </div>
        </div>

        {/* Quick KPI Risk Counters */}
        <div className="grid grid-cols-2 sm:grid-cols-5 gap-2.5">
          <div className="p-3 rounded-xl bg-slate-50 border border-slate-200">
            <span className="text-[10px] text-slate-500 font-bold block uppercase">Tổng Số Ngăn Lộ</span>
            <span className="text-lg font-black text-slate-900">{cubicleHistory.length}</span>
          </div>
          <div className="p-3 rounded-xl bg-emerald-50/70 border border-emerald-200">
            <span className="text-[10px] text-emerald-700 font-bold block uppercase">Bình Thường (Xanh)</span>
            <span className="text-lg font-black text-emerald-800">
              {cubicleHistory.filter(r => r.status === 'LOW' || r.status === 'GREEN').length}
            </span>
          </div>
          <div className="p-3 rounded-xl bg-amber-50/70 border border-amber-200">
            <span className="text-[10px] text-amber-700 font-bold block uppercase">Theo Dõi 3-6T (Vàng)</span>
            <span className="text-lg font-black text-amber-800">
              {cubicleHistory.filter(r => r.status === 'MEDIUM' || r.status === 'AMBER').length}
            </span>
          </div>
          <div className="p-3 rounded-xl bg-rose-50/70 border border-rose-200">
            <span className="text-[10px] text-rose-700 font-bold block uppercase">Nghiêm Trọng (Đỏ)</span>
            <span className="text-lg font-black text-rose-800">
              {cubicleHistory.filter(r => r.status === 'HIGH' || r.status === 'RED').length}
            </span>
          </div>
          <div className="p-3 rounded-xl bg-purple-50/70 border border-purple-200">
            <span className="text-[10px] text-purple-700 font-bold block uppercase">Khẩn Cấp (Tím)</span>
            <span className="text-lg font-black text-purple-800">
              {cubicleHistory.filter(r => r.status === 'CRITICAL').length}
            </span>
          </div>
        </div>

        {/* Filter and Search Bar */}
        <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-2.5 pt-1">
          <div className="relative flex-1 max-w-sm">
            <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              type="text"
              placeholder="Tìm theo mã tủ, dạng phóng điện, kỹ sư..."
              value={historySearch}
              onChange={(e) => setHistorySearch(e.target.value)}
              className="w-full pl-8 pr-3 py-1.5 rounded-lg border border-slate-200 text-xs focus:ring-1 focus:ring-indigo-500 focus:outline-hidden"
            />
          </div>

          <div className="flex items-center gap-1.5 overflow-x-auto text-xs">
            <span className="text-[11px] text-slate-400 font-medium">Lọc:</span>
            {(['ALL', 'CRITICAL', 'HIGH', 'MEDIUM', 'LOW'] as const).map(st => (
              <button
                key={st}
                type="button"
                onClick={() => setHistoryStatusFilter(st)}
                className={`px-2 py-1 rounded-md text-[11px] font-bold transition-all border ${
                  historyStatusFilter === st
                    ? 'bg-slate-900 text-white border-slate-900 shadow-2xs'
                    : 'bg-white text-slate-600 border-slate-200 hover:bg-slate-50'
                }`}
              >
                {st === 'ALL' ? 'Tất cả' : st === 'CRITICAL' ? 'Khẩn cấp' : st === 'HIGH' ? 'Cao/Đỏ' : st === 'MEDIUM' ? 'Theo dõi' : 'Bình thường'}
              </button>
            ))}
          </div>
        </div>

        {/* History Table */}
        <div className="overflow-x-auto border border-slate-200 rounded-xl">
          <table className="w-full text-left text-xs border-collapse">
            <thead>
              <tr className="bg-slate-100 text-slate-700 text-[11px]">
                <th className="p-2.5 border-b border-slate-200 font-bold">Mã Ngăn Lộ</th>
                <th className="p-2.5 border-b border-slate-200 font-bold">Ngày Khảo Sát</th>
                <th className="p-2.5 border-b border-slate-200 font-bold">Dạng Phóng Điện (EA Tech)</th>
                <th className="p-2.5 border-b border-slate-200 font-bold text-center">TEV (dBmV)</th>
                <th className="p-2.5 border-b border-slate-200 font-bold text-center">Siêu Âm (dBµV)</th>
                <th className="p-2.5 border-b border-slate-200 font-bold text-center">Nhiệt ΔT (°C)</th>
                <th className="p-2.5 border-b border-slate-200 font-bold text-center">Mức Rủi Ro</th>
                <th className="p-2.5 border-b border-slate-200 font-bold">Kỹ Sư FSE</th>
                <th className="p-2.5 border-b border-slate-200 font-bold text-center">Thao Tác</th>
              </tr>
            </thead>
            <tbody>
              {cubicleHistory
                .filter(rec => {
                  const matchSearch =
                    rec.cubicleId.toLowerCase().includes(historySearch.toLowerCase()) ||
                    rec.defect.toLowerCase().includes(historySearch.toLowerCase()) ||
                    rec.inspector.toLowerCase().includes(historySearch.toLowerCase());
                  const matchStatus =
                    historyStatusFilter === 'ALL' ||
                    (historyStatusFilter === 'CRITICAL' && rec.status === 'CRITICAL') ||
                    (historyStatusFilter === 'HIGH' && (rec.status === 'HIGH' || rec.status === 'RED')) ||
                    (historyStatusFilter === 'MEDIUM' && (rec.status === 'MEDIUM' || rec.status === 'AMBER')) ||
                    (historyStatusFilter === 'LOW' && (rec.status === 'LOW' || rec.status === 'GREEN'));
                  return matchSearch && matchStatus;
                })
                .map((rec) => (
                  <tr key={rec.id} className="hover:bg-slate-50/90 transition-colors border-b border-slate-100 last:border-b-0">
                    <td className="p-2.5 font-bold text-slate-900 flex items-center gap-1.5">
                      <span className="w-1.5 h-1.5 rounded-full bg-indigo-500"></span>
                      <span>{rec.cubicleId}</span>
                    </td>
                    <td className="p-2.5 text-slate-600">{rec.date}</td>
                    <td className="p-2.5 font-mono font-bold text-indigo-900">
                      <span className="px-1.5 py-0.5 rounded bg-indigo-50 border border-indigo-100">
                        {rec.defect}
                      </span>
                    </td>
                    <td className="p-2.5 text-center font-bold text-slate-800">{rec.tev}</td>
                    <td className="p-2.5 text-center font-bold text-slate-800">{rec.us}</td>
                    <td className="p-2.5 text-center font-bold text-rose-700">{rec.deltaT}°C</td>
                    <td className="p-2.5 text-center">
                      <span className={`px-2 py-0.5 rounded text-[10px] font-black border ${
                        rec.status === 'CRITICAL'
                          ? 'bg-purple-100 text-purple-800 border-purple-300'
                          : rec.status === 'HIGH' || rec.status === 'RED'
                          ? 'bg-rose-100 text-rose-800 border-rose-300'
                          : rec.status === 'AMBER' || rec.status === 'MEDIUM'
                          ? 'bg-amber-100 text-amber-800 border-amber-300'
                          : 'bg-emerald-100 text-emerald-800 border-emerald-300'
                      }`}>
                        {rec.status}
                      </span>
                    </td>
                    <td className="p-2.5 text-slate-700 font-semibold">{rec.inspector}</td>
                    <td className="p-2.5 text-center">
                      <div className="flex items-center justify-center gap-1.5">
                        <button
                          type="button"
                          onClick={() => handleLoadRecordIntoForm(rec)}
                          className="px-2 py-1 rounded bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold text-[10px] border border-indigo-200 transition-colors"
                          title="Nạp số liệu ngăn lộ này lên form chẩn đoán"
                        >
                          Nạp Form
                        </button>
                        <button
                          type="button"
                          onClick={() => handleDeleteRecord(rec.id)}
                          className="p-1 rounded text-slate-400 hover:text-rose-600 hover:bg-rose-50 transition-colors"
                          title="Xóa ngăn lộ này"
                        >
                          <Trash2 size={13} />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* -------------------------------------------------------------------- */}
      {/* 4. MODAL: MASTER SYSTEM INSTRUCTION PROMPT FOR GOOGLE AI STUDIO */}
      {/* -------------------------------------------------------------------- */}
      {showPromptModal && (
        <div className="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-slate-900 text-white border border-indigo-500/40 rounded-2xl max-w-3xl w-full max-h-[85vh] flex flex-col shadow-2xl">
            <div className="p-4 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Sparkles size={18} className="text-amber-400" />
                <h3 className="font-bold text-sm text-white">
                  Master System Instruction Prompt (Google AI Studio)
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

            <div className="p-4 overflow-y-auto flex-1 font-mono text-xs bg-slate-950 text-indigo-200 whitespace-pre-wrap leading-relaxed select-all">
              {MASTER_SYSTEM_INSTRUCTION}
            </div>

            <div className="p-4 border-t border-slate-800 flex items-center justify-between">
              <span className="text-xs text-slate-400">
                Nhấp nút Copy rồi dán trực tiếp vào mục "System Instructions" trên Google AI Studio.
              </span>
              <button
                type="button"
                onClick={() => {
                  navigator.clipboard.writeText(MASTER_SYSTEM_INSTRUCTION);
                  setCopiedPrompt(true);
                  setTimeout(() => setCopiedPrompt(false), 3000);
                }}
                className="flex items-center gap-2 px-4 py-2 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-bold shadow-sm"
              >
                {copiedPrompt ? <Check size={14} className="text-emerald-300" /> : <Copy size={14} />}
                <span>{copiedPrompt ? 'Đã sao chép vào Clipboard!' : 'Sao chép Prompt'}</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
