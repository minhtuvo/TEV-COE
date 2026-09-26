import React, { useState, useMemo, useRef } from 'react';
import {
  Sun,
  Flame,
  ShieldAlert,
  ShieldCheck,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  Sparkles,
  Camera,
  Upload,
  RefreshCw,
  Printer,
  Copy,
  Check,
  FileCode,
  Eye,
  Thermometer,
  Layers,
  Activity,
  ArrowRight,
  Info,
  Sliders,
  ChevronRight,
  Crosshair,
  Maximize2,
  Minimize2,
  Zap,
  HelpCircle,
  Clock,
  Gauge
} from 'lucide-react';
import {
  PvSeverityLevel,
  PvLikelyCause,
  PvRecommendedAction,
  PvThermalInputPayload,
  PvCellAnalysisResponse,
  PvCellAnalysisJsonData
} from '../../server/pvCellThermalAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

// Pre-configured Defect Catalog based on user's Sitemark / Volateq inspection specification
export interface DefectPatternItem {
  id: string;
  name: string;
  vietnameseName: string;
  type: 'thermal' | 'visual';
  category: string;
  sampleTmax: number;
  sampleTref: number;
  sampleSeverity: PvSeverityLevel;
  sampleCause: PvLikelyCause;
  sampleAction: PvRecommendedAction;
  description: string;
  visualNote: string;
  svgGraphicType: string;
}

export const SITEMARK_THERMAL_ANOMALIES: DefectPatternItem[] = [
  {
    id: 'hotspot',
    name: 'Hotspot',
    vietnameseName: 'Hotspot đơn cell (Single Cell Hotspot)',
    type: 'thermal',
    category: 'Cell Level',
    sampleTmax: 56.4,
    sampleTref: 41.2,
    sampleSeverity: '3 — Significant',
    sampleCause: 'Module defect',
    sampleAction: 'Field check',
    description: 'Một cell đơn lẻ nóng cục bộ do nứt vi mô (micro-crack), đứt đường hàn ribbon hoặc lỗi nội bộ tế bào quang điện.',
    visualNote: 'Vết ố mờ nhỏ ở cell góc dưới bên phải, không có vật cản ngoại lai.',
    svgGraphicType: 'single_hotspot'
  },
  {
    id: 'multi-hotspot',
    name: 'Multi hotspot',
    vietnameseName: 'Đa Hotspot (Nhiều cell nóng trên cùng module)',
    type: 'thermal',
    category: 'Cell Level',
    sampleTmax: 51.8,
    sampleTref: 41.5,
    sampleSeverity: '3 — Significant',
    sampleCause: 'Module defect',
    sampleAction: 'Field check',
    description: 'Nhiều cell nằm rải rác trên mô-đun phát nhiệt đồng thời, dấu hiệu suy thoái tế bào quang điện hoặc áp lực cơ học.',
    visualNote: 'Không có bóng che, nhiều cell có vết biến màu sẫm hơn bình thường.',
    svgGraphicType: 'multi_hotspot'
  },
  {
    id: 'single-bypassed-substring',
    name: 'Single bypassed substring',
    vietnameseName: 'Kích hoạt Diode Chuỗi con 1/3 (Single Substring)',
    type: 'thermal',
    category: 'Substring Level',
    sampleTmax: 58.2,
    sampleTref: 42.0,
    sampleSeverity: '3 — Significant',
    sampleCause: 'Module defect',
    sampleAction: 'Quick fix',
    description: 'Nóng đồng đều một dải 1/3 mô-đun (thường 20 hoặc 24 cells) do Diode bypass đã kích hoạt để bảo vệ dải bị hỏng.',
    visualNote: 'Dải substring số 2 hoạt động ở chế độ phân cực ngược, nóng liên tục.',
    svgGraphicType: 'single_substring'
  },
  {
    id: 'double-bypassed-substring',
    name: 'Double bypassed substring',
    vietnameseName: 'Kích hoạt Diode Chuỗi con 2/3 (Double Substring)',
    type: 'thermal',
    category: 'Substring Level',
    sampleTmax: 61.5,
    sampleTref: 42.1,
    sampleSeverity: '3 — Significant',
    sampleCause: 'Module defect',
    sampleAction: 'Quick fix',
    description: '2 trên 3 dải Substring bị ngắt và nóng lên đồng đều, làm suy giảm hơn 66% công suất của tấm pin.',
    visualNote: '2 dải substring dọc nóng đều, tấm pin sụt áp nặng nề.',
    svgGraphicType: 'double_substring'
  },
  {
    id: 'single-diode',
    name: 'Single diode',
    vietnameseName: 'Diode đơn kích hoạt / chập hỏng',
    type: 'thermal',
    category: 'Diode / Substring',
    sampleTmax: 54.0,
    sampleTref: 42.0,
    sampleSeverity: '3 — Significant',
    sampleCause: 'Module defect',
    sampleAction: 'Quick fix',
    description: 'Diode bypass đơn bị kích hoạt do một phần cell bị suy giảm hoặc diode bị thủng chập.',
    visualNote: 'Kiểm tra hộp nối diode phía sau tấm pin.',
    svgGraphicType: 'single_diode'
  },
  {
    id: 'multi-diode',
    name: 'Multi diode',
    vietnameseName: 'Nhiều Diode bypass kích hoạt',
    type: 'thermal',
    category: 'Diode / Substring',
    sampleTmax: 59.8,
    sampleTref: 42.0,
    sampleSeverity: '3 — Significant',
    sampleCause: 'Module defect',
    sampleAction: 'Quick fix',
    description: 'Nhiều diode kích hoạt đồng thời làm toàn bộ module suy giảm nghiêm trọng.',
    visualNote: 'Nhiệt độ dải cell cao đều, tổn hao công suất đỉnh Pmax.',
    svgGraphicType: 'multi_diode'
  },
  {
    id: 'pid',
    name: 'PID',
    vietnameseName: 'Suy thoái thế năng cảm ứng (PID Degradation)',
    type: 'thermal',
    category: 'Degradation',
    sampleTmax: 48.2,
    sampleTref: 42.0,
    sampleSeverity: '2 — Mild',
    sampleCause: 'Module defect',
    sampleAction: 'Monitor',
    description: 'Nhiệt độ nóng dạng hoa văn bậc thang tăng dần từ mép khung cực âm sang cực dương dọc theo string do dòng rò cao thế.',
    visualNote: 'Mô-đun nằm ở cuối chuỗi DC gần cực âm biến tần biến đổi hoa văn nhiệt.',
    svgGraphicType: 'pid_pattern'
  },
  {
    id: 'string-open-circuit',
    name: 'String (open circuit)',
    vietnameseName: 'Chuỗi pin hở mạch (Open Circuit String)',
    type: 'thermal',
    category: 'String Level',
    sampleTmax: 45.8,
    sampleTref: 41.5,
    sampleSeverity: '3 — Significant',
    sampleCause: 'Wiring / installation',
    sampleAction: 'Repair',
    description: 'Toàn bộ chuỗi pin nóng hơn các chuỗi lân cận từ 2°C - 5°C do toàn bộ năng lượng mặt trời chuyển thành nhiệt mà không phát ra điện.',
    visualNote: 'Không có dòng điện DC chạy qua chuỗi; kiểm tra cầu chì hoặc công tắc DC.',
    svgGraphicType: 'open_circuit_string'
  },
  {
    id: 'string-reversed-polarity',
    name: 'String (reversed polarity)',
    vietnameseName: 'Đảo cực tính chuỗi (Reversed Polarity - Mất An Toàn)',
    type: 'thermal',
    category: 'String Level',
    sampleTmax: 78.5,
    sampleTref: 41.0,
    sampleSeverity: '4 — Safety-relevant',
    sampleCause: 'Wiring / installation',
    sampleAction: 'Repair',
    description: 'CỰC KỲ NGUY HIỂM: Chuỗi pin bị đấu ngược cực tính (+/-), nhận dòng điện ngắn mạch ngược từ các chuỗi song song khỏe gây phát nhiệt dữ dội.',
    visualNote: 'Toàn bộ chuỗi phát nhiệt nóng rực, nguy cơ bốc khói hộp gom DC.',
    svgGraphicType: 'reversed_polarity'
  },
  {
    id: 'heated-junction-box',
    name: 'Heated Junction Box',
    vietnameseName: 'Hộp đấu nối quá nhiệt (Junction Box Overheating)',
    type: 'thermal',
    category: 'Electrical Connection',
    sampleTmax: 68.4,
    sampleTref: 42.0,
    sampleSeverity: '4 — Safety-relevant',
    sampleCause: 'Module defect',
    sampleAction: 'Field check',
    description: 'Nhiệt độ tại vị trí Junction Box phía sau tấm pin vượt ngưỡng an toàn (ΔT = 26.4°C ≥ 10°C), nguy cơ cháy thủng màng lưng.',
    visualNote: 'Vị trí hộp đấu nối mặt lưng phát nhiệt điểm đỏ rực trên ảnh nhiệt IR.',
    svgGraphicType: 'junction_box'
  },
  {
    id: 'single-cross-cell',
    name: 'Single Cross Cell',
    vietnameseName: 'Nứt chéo cell đơn (Cross Cell Defect)',
    type: 'thermal',
    category: 'Cell Level',
    sampleTmax: 53.5,
    sampleTref: 41.8,
    sampleSeverity: '3 — Significant',
    sampleCause: 'Module defect',
    sampleAction: 'Field check',
    description: 'Nứt chéo bề mặt silicon làm dòng điện tập trung qua khe hẹp sinh nhiệt thành vệt sáng chéo.',
    visualNote: 'Vết nứt chân chim cắt ngang cell dưới ánh sáng chéo.',
    svgGraphicType: 'cross_cell'
  },
  {
    id: 'multi-cross-cell',
    name: 'Multi Cross Cell',
    vietnameseName: 'Nứt chéo nhiều cell (Multi Cross Cell)',
    type: 'thermal',
    category: 'Cell Level',
    sampleTmax: 57.2,
    sampleTref: 41.8,
    sampleSeverity: '3 — Significant',
    sampleCause: 'Module defect',
    sampleAction: 'Field check',
    description: 'Nhiều cell bị nứt chéo do va đập cơ học hoặc ứng suất gió giật.',
    visualNote: 'Nhiều tế bào quang điện xuất hiện vệt nứt và phát nhiệt.',
    svgGraphicType: 'multi_cross'
  },
  {
    id: 'shaded-module',
    name: 'Shaded Module',
    vietnameseName: 'Mô-đun bị che bóng râm (Shaded Module)',
    type: 'thermal',
    category: 'Environmental',
    sampleTmax: 47.5,
    sampleTref: 42.0,
    sampleSeverity: '2 — Mild',
    sampleCause: 'Shading-induced',
    sampleAction: 'Move / remove the object',
    description: 'Bóng râm làm các cell bị che nhận ít ánh sáng và tiêu thụ điện từ các cell lành, phát nhiệt nhẹ.',
    visualNote: 'Bóng đổ từ lan can mái hoặc cột thu lôi lân cận đè lên tấm pin.',
    svgGraphicType: 'shaded_module'
  },
  {
    id: 'module-open-circuit',
    name: 'Module (open circuit)',
    vietnameseName: 'Tấm pin đơn hở mạch (Module Open Circuit)',
    type: 'thermal',
    category: 'Module Level',
    sampleTmax: 46.2,
    sampleTref: 41.8,
    sampleSeverity: '3 — Significant',
    sampleCause: 'Wiring / installation',
    sampleAction: 'Repair',
    description: 'Một tấm pin bị ngắt kết nối trong chuỗi, nóng hơn toàn bộ các tấm pin xung quanh 3-4°C.',
    visualNote: 'Giắc nối MC4 giữa 2 tấm pin bị lỏng hoặc tuột chốt khóa.',
    svgGraphicType: 'module_open'
  },
  {
    id: 'combiner-box',
    name: 'Combiner box',
    vietnameseName: 'Tủ gom chuỗi DC quá nhiệt (Combiner Box Anomaly)',
    type: 'thermal',
    category: 'Balance of System',
    sampleTmax: 65.0,
    sampleTref: 40.0,
    sampleSeverity: '4 — Safety-relevant',
    sampleCause: 'Wiring / installation',
    sampleAction: 'Field check',
    description: 'Quá nhiệt tại khối cầu chì DC, thanh cái đồng hoặc tiếp điểm rơ-le trong tủ gom chuỗi Combiner Box.',
    visualNote: 'Vết xém nhiệt tại cầu chì DC chuỗi 03.',
    svgGraphicType: 'combiner_box'
  },
  {
    id: 'inverter',
    name: 'Inverter',
    vietnameseName: 'Biến tần DC/AC quá nhiệt cục bộ (Inverter Anomaly)',
    type: 'thermal',
    category: 'Balance of System',
    sampleTmax: 72.0,
    sampleTref: 42.0,
    sampleSeverity: '4 — Safety-relevant',
    sampleCause: 'Module defect',
    sampleAction: 'Field check',
    description: 'Nhiệt độ khối IGBT hoặc cuộn kháng biến tần cao bất thường, quạt tản nhiệt bị kẹt hoặc nghẹt lọc gió.',
    visualNote: 'Cảnh báo nhiệt độ Inverter trên màn hình giám sát SCADA/DAS.',
    svgGraphicType: 'inverter_hot'
  }
];

export const SITEMARK_VISUAL_ANOMALIES: DefectPatternItem[] = [
  {
    id: 'soiling',
    name: 'Soiling',
    vietnameseName: 'Bụi bẩn bề mặt (Soiling / Dust)',
    type: 'visual',
    category: 'Surface Contamination',
    sampleTmax: 46.8,
    sampleTref: 42.0,
    sampleSeverity: '2 — Mild',
    sampleCause: 'Soiling-induced',
    sampleAction: 'Clean',
    description: 'Bụi công nghiệp, phấn hoa hoặc cát bám dày làm giảm bức xạ quang học truyền vào cell.',
    visualNote: 'Lớp bụi mờ vàng nhạt bám đều trên bề mặt kính tấm pin.',
    svgGraphicType: 'soiling_visual'
  },
  {
    id: 'broken-glass',
    name: 'Broken Glass',
    vietnameseName: 'Kính cường lực bị nứt vỡ (Broken Glass)',
    type: 'visual',
    category: 'Physical Damage',
    sampleTmax: 57.0,
    sampleTref: 42.0,
    sampleSeverity: '3 — Significant',
    sampleCause: 'Module defect',
    sampleAction: 'Replace',
    description: 'Mặt kính cường lực bảo vệ phía trước bị nứt vỡ rạn chân chim do va đập cơ học hoặc mưa đá.',
    visualNote: 'Vết rạn nứt kính kiểu hoa băng vỡ nát, hơi nước có nguy cơ thẩm thấu vào bên trong.',
    svgGraphicType: 'broken_glass_visual'
  },
  {
    id: 'delamination',
    name: 'Delamination',
    vietnameseName: 'Bong tróc màng film (Delamination)',
    type: 'visual',
    category: 'Material Degradation',
    sampleTmax: 49.5,
    sampleTref: 42.0,
    sampleSeverity: '2 — Mild',
    sampleCause: 'Module defect',
    sampleAction: 'Field check',
    description: 'Màng bao gói EVA bị bong tách khỏi kính hoặc lớp màng lưng Tedlar do lão hóa thời tiết và tia UV.',
    visualNote: 'Vệt bong màu trắng sữa hoặc bong bóng khí xuất hiện ở rìa tấm pin.',
    svgGraphicType: 'delamination_visual'
  },
  {
    id: 'self-shading',
    name: 'Self Shading',
    vietnameseName: 'Tự che bóng giữa các giàn (Self Shading)',
    type: 'visual',
    category: 'Design / Shading',
    sampleTmax: 46.2,
    sampleTref: 42.0,
    sampleSeverity: '2 — Mild',
    sampleCause: 'Shading-induced',
    sampleAction: 'Monitor',
    description: 'Hàng pin phía trước che bóng hàng pin phía sau vào đầu giờ sáng hoặc cuối giờ chiều do khoảng cách giàn quá hẹp.',
    visualNote: 'Dải bóng đen ngang đáy tấm pin khi mặt trời góc thấp.',
    svgGraphicType: 'self_shading_visual'
  },
  {
    id: 'vegetation',
    name: 'Vegetation',
    vietnameseName: 'Cỏ cây mọc che khuất (Vegetation Overgrowth)',
    type: 'visual',
    category: 'Environmental',
    sampleTmax: 48.0,
    sampleTref: 42.0,
    sampleSeverity: '2 — Mild',
    sampleCause: 'Shading-induced',
    sampleAction: 'Move / remove the object',
    description: 'Cỏ dại, cây bụi xung quanh mọc cao vượt khung giá đỡ che bóng các cell hàng dưới.',
    visualNote: 'Lá cây và cành cỏ dại phủ mép dưới tấm pin.',
    svgGraphicType: 'vegetation_visual'
  },
  {
    id: 'shadowing',
    name: 'Shadowing',
    vietnameseName: 'Bóng râm kết cấu (Structural Shadowing)',
    type: 'visual',
    category: 'Environmental',
    sampleTmax: 48.5,
    sampleTref: 42.0,
    sampleSeverity: '2 — Mild',
    sampleCause: 'Shading-induced',
    sampleAction: 'Move / remove the object',
    description: 'Bóng râm cố định từ cột điện, ống thông gió, bồn nước hoặc lan can chiếu lên chuỗi pin.',
    visualNote: 'Bóng râm hình trụ vắt ngang 3 tấm pin kế cận.',
    svgGraphicType: 'shadowing_visual'
  },
  {
    id: 'dropping',
    name: 'Dropping',
    vietnameseName: 'Phân chim bám cục bộ (Bird Droppings)',
    type: 'visual',
    category: 'Surface Contamination',
    sampleTmax: 55.4,
    sampleTref: 41.5,
    sampleSeverity: '3 — Significant',
    sampleCause: 'Soiling-induced',
    sampleAction: 'Clean',
    description: 'Phân chim che kín hoàn toàn 1 cell đơn lẻ, biến cell này thành tải tiêu tán nhiệt và sinh ra Hotspot nghiêm trọng.',
    visualNote: 'Vệt phân chim trắng bám dính dày đặc ngay giữa cell số 7.',
    svgGraphicType: 'dropping_visual'
  },
  {
    id: 'birds-alive',
    name: 'Birds (alive)',
    vietnameseName: 'Chim chóc đậu trên bề mặt pin',
    type: 'visual',
    category: 'Fauna',
    sampleTmax: 46.0,
    sampleTref: 42.0,
    sampleSeverity: '2 — Mild',
    sampleCause: 'Shading-induced',
    sampleAction: 'Move / remove the object',
    description: 'Chim đậu hoặc làm tổ trên khung giá hoặc bề mặt tấm pin gây che bóng tạm thời.',
    visualNote: 'Phát hiện chim đậu trên mép khung pin.',
    svgGraphicType: 'birds_visual'
  },
  {
    id: 'removable-object',
    name: 'Removable object',
    vietnameseName: 'Dị vật có thể gỡ bỏ (Lá cây, rác thi công)',
    type: 'visual',
    category: 'Foreign Object',
    sampleTmax: 51.0,
    sampleTref: 42.0,
    sampleSeverity: '2 — Mild',
    sampleCause: 'Shading-induced',
    sampleAction: 'Move / remove the object',
    description: 'Lá cây khô, băng keo thi công hoặc mảnh ni lông rơi đè lên bề mặt tấm pin.',
    visualNote: 'Lá cây khô mục che khuất 1 phần cell.',
    svgGraphicType: 'object_visual'
  },
  {
    id: 'physical-internal',
    name: 'Physical Internal',
    vietnameseName: 'Lỗi vật lý bên trong tế bào (Internal Cell Defect)',
    type: 'visual',
    category: 'Manufacturing Defect',
    sampleTmax: 54.8,
    sampleTref: 42.0,
    sampleSeverity: '3 — Significant',
    sampleCause: 'Module defect',
    sampleAction: 'Field check',
    description: 'Vết ố ốc sên (Snail trails), đứt ngậm thanh busbar hoặc oxi hóa chân chì bên trong tế bào quang điện.',
    visualNote: 'Vết sọc sẫm màu chạy dọc đường hàn bên dưới lớp kính.',
    svgGraphicType: 'internal_visual'
  },
  {
    id: 'visible-unknown',
    name: 'Visible Unknown',
    vietnameseName: 'Bất thường chưa phân loại (Unknown Visual Anomaly)',
    type: 'visual',
    category: 'Unclassified',
    sampleTmax: 47.0,
    sampleTref: 42.0,
    sampleSeverity: '2 — Mild',
    sampleCause: 'Unknown',
    sampleAction: 'Field check',
    description: 'Dạng biến dạng hoặc đổi màu lạ mắt chưa có trong danh mục nhận diện tiêu chuẩn.',
    visualNote: 'Vệt loang lổ quang học bất thường.',
    svgGraphicType: 'unknown_visual'
  },
  {
    id: 'object',
    name: 'Object',
    vietnameseName: 'Vật thể lạ rơi trên tấm pin (Foreign Object)',
    type: 'visual',
    category: 'Foreign Object',
    sampleTmax: 52.0,
    sampleTref: 41.5,
    sampleSeverity: '3 — Significant',
    sampleCause: 'Shading-induced',
    sampleAction: 'Move / remove the object',
    description: 'Vật cản cơ học che khuất hoàn toàn đường dẫn quang học.',
    visualNote: 'Dụng cụ hoặc mảnh xà gồ rơi trên mặt tấm pin.',
    svgGraphicType: 'object_visual'
  }
];

export const PvCellDeepThermalAnalyzer: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  // Active catalog tab
  const [catalogTab, setCatalogTab] = useState<'thermal' | 'visual'>('thermal');
  const [selectedPattern, setSelectedPattern] = useState<DefectPatternItem>(SITEMARK_THERMAL_ANOMALIES[0]);

  // Operational & Field inputs
  const [moduleTag, setModuleTag] = useState('PV-MOD-STRING03-12');
  const [projectName, setProjectName] = useState('Nha May Dien Mat Troi TEV Roof 1MW');
  const [fseName, setFseName] = useState('Võ Minh Tú (sgm1707@gmail.com)');
  const [testDate, setTestDate] = useState('2026-09-23');
  const [irradiance, setIrradiance] = useState<number>(850);
  const [ambientTemp, setAmbientTemp] = useState<number>(32.0);
  const [tMax, setTMax] = useState<number>(56.4);
  const [tRef, setTRef] = useState<number>(41.2);
  const [visualNotes, setVisualNotes] = useState(
    'Phát hiện điểm phát nhiệt cục bộ trên cell số 8 thuộc Substring 2. Bề mặt không có rác che phủ.'
  );

  // Uploaded images states
  const [irImageBase64, setIrImageBase64] = useState<string | null>(null);
  const [rgbImageBase64, setRgbImageBase64] = useState<string | null>(null);
  const [irImageName, setIrImageName] = useState<string>('');
  const [rgbImageName, setRgbImageName] = useState<string>('');

  // AI & Analysis state
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<PvCellAnalysisResponse | null>(null);
  const [activeStepTab, setActiveStepTab] = useState<number>(2); // 1-5 steps
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Hidden file input refs
  const irFileInputRef = useRef<HTMLInputElement>(null);
  const rgbFileInputRef = useRef<HTMLInputElement>(null);

  // Dynamic Delta T calculation
  const calculatedDeltaT = useMemo(() => {
    return parseFloat((tMax - tRef).toFixed(1));
  }, [tMax, tRef]);

  // Quick select a pattern from catalog
  const handleSelectPattern = (item: DefectPatternItem) => {
    setSelectedPattern(item);
    setTMax(item.sampleTmax);
    setTRef(item.sampleTref);
    setVisualNotes(item.visualNote);
    setSavedSuccessMessage(null);
  };

  // Handle file uploads
  const handleImageUpload = (
    e: React.ChangeEvent<HTMLInputElement>,
    channel: 'ir' | 'rgb'
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const base64 = event.target?.result as string;
      if (channel === 'ir') {
        setIrImageBase64(base64);
        setIrImageName(file.name);
      } else {
        setRgbImageBase64(base64);
        setRgbImageName(file.name);
      }
    };
    reader.readAsDataURL(file);
  };

  // Perform AI Multimodal Analysis
  const handleRunAnalysis = async () => {
    setIsAnalyzing(true);
    setAnalysisResult(null);
    setSavedSuccessMessage(null);

    const payload: PvThermalInputPayload = {
      project_name: projectName,
      module_tag: moduleTag,
      fse_name: fseName,
      test_date: testDate,
      irradiance_w_m2: irradiance,
      ambient_temperature_c: ambientTemp,
      t_max_c: tMax,
      t_ref_c: tRef,
      pattern_type: selectedPattern.name,
      visual_defect_type: selectedPattern.type === 'visual' ? selectedPattern.name : 'None',
      visual_notes: visualNotes,
      image_ir_base64: irImageBase64 || undefined,
      image_rgb_base64: rgbImageBase64 || undefined,
      mime_type: 'image/jpeg'
    };

    try {
      const response = await fetch('/api/field-service/pv-cell-thermal-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });

      const resJson = await response.json();
      if (resJson.success) {
        setAnalysisResult(resJson);
      } else {
        alert('Lỗi phân tích: ' + (resJson.error || 'Không rõ nguyên nhân'));
      }
    } catch (err: any) {
      console.error('Lỗi khi gọi API phân tích nhiệt PV Cell:', err);
      alert('Không thể kết nối đến máy chủ AI: ' + err.message);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleSaveToCmms = () => {
    if (!analysisResult) {
      alert('Vui lòng thực hiện phân tích AI trước khi lưu báo cáo.');
      return;
    }
    if (onSaveReport) {
      onSaveReport({
        data: {
          project_name: projectName,
          module_tag: moduleTag,
          fse_name: fseName,
          test_date: testDate,
          t_max: tMax,
          t_ref: tRef,
          delta_t: calculatedDeltaT,
          pattern: selectedPattern.name,
          analysis: analysisResult
        },
        analysisResult: {
          overall_status: analysisResult.severity.includes('Safety') ? 'FAIL' : analysisResult.severity.includes('Significant') ? 'INVESTIGATE' : 'PASS',
          status_reason: `Phân tích nhiệt PV Cell: ${selectedPattern.vietnameseName} (ΔT = ${calculatedDeltaT}°C)`,
          markdown_report: analysisResult.markdown_report,
          evaluations: analysisResult.json_data
        },
        equipmentId: moduleTag,
        date: testDate,
        type: 'PV Cell Thermal IR & Visual Inspection (IEC 62446-3 & Sitemark)'
      });
      setSavedSuccessMessage(`Đã lưu biên bản kiểm định PV Cell [${moduleTag}] vào kho báo cáo CMMS!`);
    }
  };

  const handleCopyJson = () => {
    if (analysisResult?.json_data) {
      navigator.clipboard.writeText(JSON.stringify(analysisResult.json_data, null, 2));
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }
  };

  // Severity badge color helper
  const getSeverityBadge = (level: string) => {
    if (level.includes('4') || level.includes('Safety')) {
      return {
        bg: 'bg-red-500/20 text-red-400 border-red-500/40',
        text: 'Level 4 — Safety-relevant (Mất an toàn / Nguy cơ cháy nổ)',
        dot: 'bg-red-500 animate-ping',
        indicator: '🔴 Nguy hiểm cấp 4'
      };
    }
    if (level.includes('3') || level.includes('Significant')) {
      return {
        bg: 'bg-orange-500/20 text-orange-400 border-orange-500/40',
        text: 'Level 3 — Significant (Đáng chú ý / Cần sửa chữa)',
        dot: 'bg-orange-500',
        indicator: '🟠 Đáng chú ý cấp 3'
      };
    }
    if (level.includes('2') || level.includes('Mild')) {
      return {
        bg: 'bg-yellow-500/20 text-yellow-300 border-yellow-500/40',
        text: 'Level 2 — Mild (Mức độ nhẹ / Tiếp tục theo dõi)',
        dot: 'bg-yellow-400',
        indicator: '🟡 Mức độ nhẹ cấp 2'
      };
    }
    return {
      bg: 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40',
      text: 'Level 1 — No anomaly (Hoàn toàn bình thường)',
      dot: 'bg-emerald-400',
      indicator: '🟢 Bình thường cấp 1'
    };
  };

  const getActionBadge = (action: string) => {
    switch (action) {
      case 'Replace':
        return 'bg-red-600/30 text-red-200 border-red-500';
      case 'Repair':
        return 'bg-amber-600/30 text-amber-200 border-amber-500';
      case 'Quick fix':
        return 'bg-blue-600/30 text-blue-200 border-blue-500';
      case 'Field check':
        return 'bg-purple-600/30 text-purple-200 border-purple-500';
      case 'Move / remove the object':
        return 'bg-indigo-600/30 text-indigo-200 border-indigo-500';
      case 'Clean':
        return 'bg-cyan-600/30 text-cyan-200 border-cyan-500';
      default:
        return 'bg-slate-700 text-slate-300 border-slate-600';
    }
  };

  return (
    <div className="space-y-6">
      {/* Hidden file inputs */}
      <input
        type="file"
        ref={irFileInputRef}
        accept="image/*"
        className="hidden"
        onChange={(e) => handleImageUpload(e, 'ir')}
      />
      <input
        type="file"
        ref={rgbFileInputRef}
        accept="image/*"
        className="hidden"
        onChange={(e) => handleImageUpload(e, 'rgb')}
      />

      {/* HEADER BANNER */}
      <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-blue-950 text-white p-5 sm:p-6 rounded-2xl shadow-xl border border-indigo-500/30">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-2">
            <div className="flex flex-wrap items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-extrabold uppercase tracking-wider bg-indigo-500/20 text-indigo-300 border border-indigo-400/30 flex items-center gap-1.5">
                <Flame size={13} className="text-amber-400" />
                Tiêu Chuẩn IEC 62446-3 & Volateq / Sitemark
              </span>
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-blue-900/40 text-blue-200 border border-blue-400/30 flex items-center gap-1">
                <Sparkles size={12} className="text-yellow-300" />
                AI Multimodal Vision Phân Tích Ảnh Nhiệt PV Cell
              </span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white flex items-center gap-2">
              <Sun className="text-amber-400" size={26} />
              Chuyên Sâu Tấm Pin & Cell Mặt Trời (PV Cell Thermal IR & RGB)
            </h1>
            <p className="text-xs sm:text-sm text-slate-300 max-w-3xl leading-relaxed">
              Thuật toán 5 bước: <strong>Cắt phân vùng (Segmentation)</strong> ➔ <strong>Đo nhiệt độ ΔT & Phân loại</strong> ➔ <strong>Đánh giá Mức độ Nghiêm trọng (Level 1-4)</strong> ➔ <strong>Suy luận Nguyên nhân Khả thi</strong> ➔ <strong>Đề xuất Hành động</strong> dựa theo tiêu chuẩn phân tích ảnh nhiệt IEC 62446-3 và khung chuẩn hóa Volateq / Sitemark.
            </p>
          </div>

          {/* Quick Actions */}
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => {
                irFileInputRef.current?.click();
              }}
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-amber-500 hover:bg-amber-600 text-slate-950 flex items-center gap-1.5 transition-all shadow-md active:scale-95"
            >
              <Upload size={14} />
              <span>Tải Ảnh Nhiệt (IR)</span>
            </button>

            <button
              onClick={() => {
                rgbFileInputRef.current?.click();
              }}
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-blue-600 hover:bg-blue-700 text-white flex items-center gap-1.5 transition-all shadow-md active:scale-95"
            >
              <Camera size={14} />
              <span>Tải Ảnh Thực Tế (RGB)</span>
            </button>

            <button
              onClick={() => setShowPrintModal(true)}
              disabled={!analysisResult}
              className={`px-3 py-2 rounded-xl text-xs font-semibold flex items-center gap-1.5 transition-all border ${
                analysisResult
                  ? 'bg-white/10 hover:bg-white/20 text-white border-white/20 cursor-pointer'
                  : 'bg-white/5 text-slate-500 border-white/10 cursor-not-allowed'
              }`}
            >
              <Printer size={14} />
              <span>In Báo Cáo PDF</span>
            </button>
          </div>
        </div>

        {/* 5-STEP PIPELINE TRACKER */}
        <div className="mt-5 pt-4 border-t border-indigo-900/50">
          <div className="text-[11px] font-bold text-indigo-300 uppercase tracking-wider mb-2 flex items-center gap-1.5">
            <Layers size={13} />
            Thuật Toán 5 Bước Xử Lý Tuần Tự (AI Pipeline)
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-2 text-xs">
            {/* Step 1 */}
            <div
              onClick={() => setActiveStepTab(1)}
              className={`p-2.5 rounded-xl border transition-all cursor-pointer ${
                activeStepTab === 1
                  ? 'bg-indigo-600/30 border-indigo-400 text-white'
                  : 'bg-slate-800/60 border-slate-700/60 text-slate-300 hover:bg-slate-800'
              }`}
            >
              <div className="flex items-center justify-between font-bold text-[11px] mb-1 text-indigo-300">
                <span>BƯỚC 1</span>
                <span className="text-[10px] text-slate-400">IR / RGB</span>
              </div>
              <div className="font-semibold text-slate-200">Nhận diện & Cắt hình</div>
              <div className="text-[11px] text-slate-400 mt-0.5">String ➔ Module ➔ Substring ➔ Cell</div>
            </div>

            {/* Step 2 */}
            <div
              onClick={() => setActiveStepTab(2)}
              className={`p-2.5 rounded-xl border transition-all cursor-pointer ${
                activeStepTab === 2
                  ? 'bg-amber-600/30 border-amber-400 text-white'
                  : 'bg-slate-800/60 border-slate-700/60 text-slate-300 hover:bg-slate-800'
              }`}
            >
              <div className="flex items-center justify-between font-bold text-[11px] mb-1 text-amber-300">
                <span>BƯỚC 2</span>
                <span className="text-[10px] font-mono text-amber-200">ΔT: {calculatedDeltaT}°C</span>
              </div>
              <div className="font-semibold text-slate-200">Đo Nhiệt Độ ΔT & Phân Loại</div>
              <div className="text-[11px] text-slate-400 mt-0.5">Hotspot, Substring, PID, JB...</div>
            </div>

            {/* Step 3 */}
            <div
              onClick={() => setActiveStepTab(3)}
              className={`p-2.5 rounded-xl border transition-all cursor-pointer ${
                activeStepTab === 3
                  ? 'bg-orange-600/30 border-orange-400 text-white'
                  : 'bg-slate-800/60 border-slate-700/60 text-slate-300 hover:bg-slate-800'
              }`}
            >
              <div className="flex items-center justify-between font-bold text-[11px] mb-1 text-orange-300">
                <span>BƯỚC 3</span>
                <span className="text-[10px] text-slate-400">Level 1 - 4</span>
              </div>
              <div className="font-semibold text-slate-200">Mức Độ Nghiêm Trọng</div>
              <div className="text-[11px] text-slate-400 mt-0.5">Thang đánh giá tích lũy</div>
            </div>

            {/* Step 4 */}
            <div
              onClick={() => setActiveStepTab(4)}
              className={`p-2.5 rounded-xl border transition-all cursor-pointer ${
                activeStepTab === 4
                  ? 'bg-purple-600/30 border-purple-400 text-white'
                  : 'bg-slate-800/60 border-slate-700/60 text-slate-300 hover:bg-slate-800'
              }`}
            >
              <div className="flex items-center justify-between font-bold text-[11px] mb-1 text-purple-300">
                <span>BƯỚC 4</span>
                <span className="text-[10px] text-slate-400">Precedence</span>
              </div>
              <div className="font-semibold text-slate-200">Suy Luận Nguyên Nhân</div>
              <div className="text-[11px] text-slate-400 mt-0.5">Điện/Cấu trúc ➔ Nhiệt ➔ Ngoại sinh</div>
            </div>

            {/* Step 5 */}
            <div
              onClick={() => setActiveStepTab(5)}
              className={`p-2.5 rounded-xl border transition-all cursor-pointer ${
                activeStepTab === 5
                  ? 'bg-cyan-600/30 border-cyan-400 text-white'
                  : 'bg-slate-800/60 border-slate-700/60 text-slate-300 hover:bg-slate-800'
              }`}
            >
              <div className="flex items-center justify-between font-bold text-[11px] mb-1 text-cyan-300">
                <span>BƯỚC 5</span>
                <span className="text-[10px] text-slate-400">First-match</span>
              </div>
              <div className="font-semibold text-slate-200">Đề Xuất Hành Động</div>
              <div className="text-[11px] text-slate-400 mt-0.5">Replace ➔ Repair ➔ Clean...</div>
            </div>
          </div>
        </div>
      </div>

      {savedSuccessMessage && (
        <div className="p-3.5 rounded-xl bg-emerald-950/60 border border-emerald-500/50 text-emerald-200 text-xs sm:text-sm flex items-center justify-between">
          <div className="flex items-center gap-2">
            <CheckCircle2 size={16} className="text-emerald-400" />
            <span>{savedSuccessMessage}</span>
          </div>
          {onNavigateToReports && (
            <button
              onClick={onNavigateToReports}
              className="text-xs font-bold underline hover:text-white"
            >
              Xem Kho Báo Cáo
            </button>
          )}
        </div>
      )}

      {/* MAIN TWO-COLUMN WORKSPACE */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* LEFT COLUMN: Image Input, Parameters & Catalog (5 cols) */}
        <div className="lg:col-span-5 space-y-6">
          {/* 1. DUAL SPECTRUM IMAGE DROPZONE */}
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4 shadow-lg space-y-4">
            <div className="flex items-center justify-between">
              <h2 className="text-sm font-bold text-white flex items-center gap-2">
                <Camera size={16} className="text-amber-400" />
                Ảnh Chụp Hồng Ngoại (IR) & Ảnh Thực Tế (RGB)
              </h2>
              <span className="text-[11px] text-slate-400">Kênh kép (Dual-Channel)</span>
            </div>

            <div className="grid grid-cols-2 gap-3">
              {/* IR Thermal Channel */}
              <div className="space-y-2">
                <div className="text-xs font-semibold text-amber-300 flex items-center justify-between">
                  <span>Kênh Nhiệt (IR)</span>
                  {irImageBase64 && (
                    <button
                      onClick={() => {
                        setIrImageBase64(null);
                        setIrImageName('');
                      }}
                      className="text-[10px] text-rose-400 hover:underline"
                    >
                      Xóa ảnh
                    </button>
                  )}
                </div>

                <div
                  onClick={() => irFileInputRef.current?.click()}
                  className="aspect-video w-full rounded-xl border-2 border-dashed border-amber-500/40 hover:border-amber-400 bg-amber-950/20 hover:bg-amber-950/30 flex flex-col items-center justify-center p-2 text-center cursor-pointer transition-all overflow-hidden relative group"
                >
                  {irImageBase64 ? (
                    <>
                      <img
                        src={irImageBase64}
                        alt="IR Thermal"
                        className="w-full h-full object-cover rounded-lg"
                      />
                      <div className="absolute inset-0 bg-slate-950/70 opacity-0 group-hover:opacity-100 flex items-center justify-center text-xs text-amber-200 transition-all font-semibold">
                        Bấm để đổi ảnh
                      </div>
                    </>
                  ) : (
                    <div className="space-y-1">
                      <Flame size={24} className="mx-auto text-amber-400 opacity-80" />
                      <div className="text-[11px] font-bold text-amber-200">Tải Ảnh Nhiệt IR</div>
                      <div className="text-[10px] text-slate-400">FLIR, DJI Thermal, v.v.</div>
                    </div>
                  )}
                </div>
                {irImageName && (
                  <div className="text-[10px] text-slate-400 truncate text-center">
                    {irImageName}
                  </div>
                )}
              </div>

              {/* RGB Visual Channel */}
              <div className="space-y-2">
                <div className="text-xs font-semibold text-blue-300 flex items-center justify-between">
                  <span>Kênh Thực Tế (RGB)</span>
                  {rgbImageBase64 && (
                    <button
                      onClick={() => {
                        setRgbImageBase64(null);
                        setRgbImageName('');
                      }}
                      className="text-[10px] text-rose-400 hover:underline"
                    >
                      Xóa ảnh
                    </button>
                  )}
                </div>

                <div
                  onClick={() => rgbFileInputRef.current?.click()}
                  className="aspect-video w-full rounded-xl border-2 border-dashed border-blue-500/40 hover:border-blue-400 bg-blue-950/20 hover:bg-blue-950/30 flex flex-col items-center justify-center p-2 text-center cursor-pointer transition-all overflow-hidden relative group"
                >
                  {rgbImageBase64 ? (
                    <>
                      <img
                        src={rgbImageBase64}
                        alt="RGB Visual"
                        className="w-full h-full object-cover rounded-lg"
                      />
                      <div className="absolute inset-0 bg-slate-950/70 opacity-0 group-hover:opacity-100 flex items-center justify-center text-xs text-blue-200 transition-all font-semibold">
                        Bấm để đổi ảnh
                      </div>
                    </>
                  ) : (
                    <div className="space-y-1">
                      <Eye size={24} className="mx-auto text-blue-400 opacity-80" />
                      <div className="text-[11px] font-bold text-blue-200">Tải Ảnh Kính RGB</div>
                      <div className="text-[10px] text-slate-400">Drone RGB, Điện thoại</div>
                    </div>
                  )}
                </div>
                {rgbImageName && (
                  <div className="text-[10px] text-slate-400 truncate text-center">
                    {rgbImageName}
                  </div>
                )}
              </div>
            </div>

            {/* Synthetic Thermal Preview if no user image uploaded */}
            {!irImageBase64 && (
              <div className="p-3 rounded-xl bg-slate-950/80 border border-slate-800 text-xs space-y-2">
                <div className="flex items-center justify-between text-slate-300 font-semibold text-[11px]">
                  <span className="flex items-center gap-1.5 text-amber-300">
                    <Thermometer size={13} />
                    Mô Phỏng Phổ Nhiệt (Ironbow Thermal Map)
                  </span>
                  <span className="text-[10px] font-mono text-slate-400">
                    ΔT = {calculatedDeltaT}°C
                  </span>
                </div>

                {/* Interactive Simulated PV Module with Hotspot */}
                <div className="relative w-full h-24 rounded-lg bg-gradient-to-r from-indigo-950 via-purple-950 to-blue-950 border border-slate-700 p-2 overflow-hidden flex flex-col justify-between">
                  <div className="grid grid-cols-6 gap-1 w-full h-full opacity-60">
                    {Array.from({ length: 18 }).map((_, idx) => {
                      const isHotCell =
                        (selectedPattern.id === 'hotspot' && idx === 11) ||
                        (selectedPattern.id === 'multi-hotspot' && (idx === 4 || idx === 11 || idx === 14)) ||
                        (selectedPattern.id === 'single-bypassed-substring' && idx >= 6 && idx <= 11) ||
                        (selectedPattern.id === 'double-bypassed-substring' && idx <= 11) ||
                        (selectedPattern.id === 'heated-junction-box' && idx === 2);

                      return (
                        <div
                          key={idx}
                          className={`rounded-xs border border-white/10 transition-all ${
                            isHotCell
                              ? 'bg-gradient-to-br from-yellow-300 via-amber-500 to-red-600 shadow-lg shadow-amber-500/50 animate-pulse'
                              : 'bg-blue-900/40 hover:bg-blue-800/40'
                          }`}
                        />
                      );
                    })}
                  </div>

                  <div className="absolute bottom-1 right-2 bg-slate-950/80 px-2 py-0.5 rounded-md text-[10px] text-amber-300 font-mono border border-slate-700">
                    T_max: {tMax}°C | T_ref: {tRef}°C
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* 2. FIELD TEMPERATURES & ENVIRONMENTAL PARAMETERS */}
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4 shadow-lg space-y-4">
            <h2 className="text-sm font-bold text-white flex items-center justify-between">
              <span className="flex items-center gap-2">
                <Sliders size={16} className="text-indigo-400" />
                Thông Số Đo Thực Địa Ngoài Site
              </span>
              <span className="text-[11px] text-indigo-300 font-mono">
                ΔT = {calculatedDeltaT}°C
              </span>
            </h2>

            <div className="grid grid-cols-2 gap-3 text-xs">
              <div>
                <label className="text-[11px] font-semibold text-slate-300 block mb-1">
                  Mã Tấm Pin / Vị Trí
                </label>
                <input
                  type="text"
                  value={moduleTag}
                  onChange={(e) => setModuleTag(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-700 rounded-xl px-2.5 py-1.5 text-white font-mono focus:border-indigo-500 focus:outline-hidden"
                />
              </div>

              <div>
                <label className="text-[11px] font-semibold text-slate-300 block mb-1">
                  Bức Xạ Mặt Trời (W/m²)
                </label>
                <input
                  type="number"
                  value={irradiance}
                  onChange={(e) => setIrradiance(Number(e.target.value))}
                  className="w-full bg-slate-950 border border-slate-700 rounded-xl px-2.5 py-1.5 text-white font-mono focus:border-indigo-500 focus:outline-hidden"
                />
                {irradiance < 600 && (
                  <span className="text-[10px] text-amber-400 block mt-0.5">
                    ⚠️ IEC 62446-3 khuyến nghị ≥ 600 W/m²
                  </span>
                )}
              </div>

              <div>
                <label className="text-[11px] font-semibold text-amber-300 block mb-1">
                  Nhiệt độ T_max (°C)
                </label>
                <input
                  type="number"
                  step="0.1"
                  value={tMax}
                  onChange={(e) => setTMax(Number(e.target.value))}
                  className="w-full bg-slate-950 border border-amber-500/50 rounded-xl px-2.5 py-1.5 text-amber-300 font-bold font-mono focus:border-amber-400 focus:outline-hidden"
                />
              </div>

              <div>
                <label className="text-[11px] font-semibold text-blue-300 block mb-1">
                  Nhiệt độ T_ref (°C)
                </label>
                <input
                  type="number"
                  step="0.1"
                  value={tRef}
                  onChange={(e) => setTRef(Number(e.target.value))}
                  className="w-full bg-slate-950 border border-blue-500/50 rounded-xl px-2.5 py-1.5 text-blue-300 font-bold font-mono focus:border-blue-400 focus:outline-hidden"
                />
              </div>
            </div>

            {/* Visual Notes */}
            <div>
              <label className="text-[11px] font-semibold text-slate-300 block mb-1">
                Ghi Chú Thị Giác (Bụi bẩn, Nứt vỡ, Che bóng, Phân chim)
              </label>
              <textarea
                value={visualNotes}
                onChange={(e) => setVisualNotes(e.target.value)}
                rows={2}
                className="w-full bg-slate-950 border border-slate-700 rounded-xl p-2 text-xs text-slate-200 focus:border-indigo-500 focus:outline-hidden resize-none"
              />
            </div>

            {/* Run Analysis Action Button */}
            <button
              onClick={handleRunAnalysis}
              disabled={isAnalyzing}
              className={`w-full py-3 rounded-xl font-bold text-xs sm:text-sm flex items-center justify-center gap-2 shadow-lg transition-all active:scale-[0.98] ${
                isAnalyzing
                  ? 'bg-slate-700 text-slate-400 cursor-not-allowed'
                  : 'bg-gradient-to-r from-amber-500 via-orange-500 to-indigo-600 hover:from-amber-400 hover:to-indigo-500 text-slate-950 font-black cursor-pointer'
              }`}
            >
              {isAnalyzing ? (
                <>
                  <RefreshCw size={16} className="animate-spin text-slate-400" />
                  <span>Đang Phân Tích Vision AI (IEC 62446-3 & Sitemark)...</span>
                </>
              ) : (
                <>
                  <Sparkles size={16} className="text-slate-950" />
                  <span>Phân Tích Chuyên Sâu Tấm Pin & Cell (AI Multimodal)</span>
                </>
              )}
            </button>
          </div>

          {/* 3. THƯ VIỆN MẪU LỖI CHUẨN HÓA SITEMARK / VOLATEQ */}
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4 shadow-lg space-y-4">
            <div className="flex items-center justify-between">
              <h2 className="text-sm font-bold text-white flex items-center gap-2">
                <Flame size={16} className="text-orange-400" />
                Thư Viện Dạng Mẫu Lỗi Chuẩn Hóa
              </h2>
              <span className="text-[11px] text-slate-400">Sitemark / Volateq</span>
            </div>

            {/* Tab switch between Thermal Anomalies & Visual Anomalies */}
            <div className="flex rounded-xl bg-slate-950 p-1 border border-slate-800">
              <button
                onClick={() => setCatalogTab('thermal')}
                className={`flex-1 py-1.5 text-xs font-semibold rounded-lg transition-all flex items-center justify-center gap-1.5 ${
                  catalogTab === 'thermal'
                    ? 'bg-amber-600 text-white shadow-xs'
                    : 'text-slate-400 hover:text-white'
                }`}
              >
                <Flame size={13} />
                <span>Thermal Anomalies ({SITEMARK_THERMAL_ANOMALIES.length})</span>
              </button>

              <button
                onClick={() => setCatalogTab('visual')}
                className={`flex-1 py-1.5 text-xs font-semibold rounded-lg transition-all flex items-center justify-center gap-1.5 ${
                  catalogTab === 'visual'
                    ? 'bg-blue-600 text-white shadow-xs'
                    : 'text-slate-400 hover:text-white'
                }`}
              >
                <Eye size={13} />
                <span>Visual Anomalies ({SITEMARK_VISUAL_ANOMALIES.length})</span>
              </button>
            </div>

            {/* Pattern cards grid */}
            <div className="grid grid-cols-2 sm:grid-cols-2 gap-2 max-h-72 overflow-y-auto pr-1">
              {(catalogTab === 'thermal' ? SITEMARK_THERMAL_ANOMALIES : SITEMARK_VISUAL_ANOMALIES).map(
                (item) => {
                  const isSelected = selectedPattern.id === item.id;
                  return (
                    <div
                      key={item.id}
                      onClick={() => handleSelectPattern(item)}
                      className={`p-2.5 rounded-xl border text-left cursor-pointer transition-all flex flex-col justify-between ${
                        isSelected
                          ? 'bg-indigo-950/80 border-indigo-400 shadow-md ring-1 ring-indigo-400/50'
                          : 'bg-slate-950/60 border-slate-800 hover:border-slate-700 hover:bg-slate-950'
                      }`}
                    >
                      <div className="space-y-1">
                        <div className="flex items-center justify-between">
                          <span
                            className={`text-[9px] font-extrabold uppercase px-1.5 py-0.5 rounded-sm ${
                              item.type === 'thermal'
                                ? 'bg-amber-500/20 text-amber-300'
                                : 'bg-blue-500/20 text-blue-300'
                            }`}
                          >
                            {item.category}
                          </span>
                          <span className="text-[10px] font-mono text-slate-400">
                            {item.sampleTmax}°C
                          </span>
                        </div>
                        <div className="text-xs font-bold text-white truncate">
                          {item.name}
                        </div>
                        <div className="text-[10px] text-slate-400 line-clamp-1">
                          {item.vietnameseName}
                        </div>
                      </div>

                      <div className="mt-2 pt-1 border-t border-slate-800/80 flex items-center justify-between text-[10px]">
                        <span
                          className={`font-semibold ${
                            item.sampleSeverity.includes('4')
                              ? 'text-red-400'
                              : item.sampleSeverity.includes('3')
                              ? 'text-orange-400'
                              : 'text-yellow-400'
                          }`}
                        >
                          {item.sampleSeverity.split('—')[1]?.trim() || item.sampleSeverity}
                        </span>
                        <ChevronRight size={12} className="text-slate-500" />
                      </div>
                    </div>
                  );
                }
              )}
            </div>
            <p className="text-[11px] text-slate-400 italic">
              💡 Bấm chọn mẫu lỗi để nạp nhanh thông số nhiệt mẫu, ghi chú và cấu hình AI chẩn đoán.
            </p>
          </div>
        </div>

        {/* RIGHT COLUMN: 3-Dimension Matrix & AI Analysis Report (7 cols) */}
        <div className="lg:col-span-7 space-y-6">
          {/* Active Pattern Card Summary */}
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4 shadow-lg space-y-3">
            <div className="flex items-center justify-between">
              <span className="text-xs text-slate-400 font-semibold">
                Dạng mẫu đang chọn kiểm tra:
              </span>
              <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-indigo-500/20 text-indigo-300 border border-indigo-400/30">
                {selectedPattern.type === 'thermal' ? 'Nhiệt Hồng Ngoại (IR)' : 'Thực Tế (RGB)'}
              </span>
            </div>

            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-800 pb-3">
              <div>
                <h3 className="text-lg font-black text-white">
                  {selectedPattern.name}
                </h3>
                <p className="text-xs text-slate-300">
                  {selectedPattern.vietnameseName}
                </p>
              </div>
              <div className="text-right">
                <div className="text-xs text-slate-400">Chênh lệch nhiệt độ mẫu</div>
                <div className="text-lg font-mono font-bold text-amber-400">
                  ΔT = {(selectedPattern.sampleTmax - selectedPattern.sampleTref).toFixed(1)}°C
                </div>
              </div>
            </div>

            <p className="text-xs text-slate-300 leading-relaxed">
              {selectedPattern.description}
            </p>

            {/* 3-DIMENSION CLASSIFICATION MATRIX BANNER */}
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5 pt-2">
              {/* Dim 1: Severity */}
              <div className="p-3 rounded-xl bg-slate-950 border border-slate-800 space-y-1">
                <div className="text-[10px] font-extrabold uppercase tracking-wider text-slate-400">
                  1. Mức Độ Nghiêm Trọng
                </div>
                <div className="text-xs font-bold text-white flex items-center gap-1.5">
                  <span
                    className={`w-2 h-2 rounded-full ${
                      (analysisResult?.severity || selectedPattern.sampleSeverity).includes('4')
                        ? 'bg-red-500 animate-ping'
                        : (analysisResult?.severity || selectedPattern.sampleSeverity).includes('3')
                        ? 'bg-orange-500'
                        : 'bg-yellow-400'
                    }`}
                  />
                  <span>
                    {analysisResult?.severity || selectedPattern.sampleSeverity}
                  </span>
                </div>
              </div>

              {/* Dim 2: Likely Cause */}
              <div className="p-3 rounded-xl bg-slate-950 border border-slate-800 space-y-1">
                <div className="text-[10px] font-extrabold uppercase tracking-wider text-slate-400">
                  2. Nguyên Nhân Khả Thi
                </div>
                <div className="text-xs font-bold text-indigo-300">
                  {analysisResult?.likely_cause || selectedPattern.sampleCause}
                </div>
              </div>

              {/* Dim 3: Recommended Action */}
              <div className="p-3 rounded-xl bg-slate-950 border border-slate-800 space-y-1">
                <div className="text-[10px] font-extrabold uppercase tracking-wider text-slate-400">
                  3. Hành Động Khuyến Nghị
                </div>
                <div>
                  <span
                    className={`inline-block px-2 py-0.5 rounded-md text-[11px] font-extrabold border ${getActionBadge(
                      analysisResult?.recommended_action || selectedPattern.sampleAction
                    )}`}
                  >
                    {analysisResult?.recommended_action || selectedPattern.sampleAction}
                  </span>
                </div>
              </div>
            </div>
          </div>

          {/* CRITICAL SAFETY BANNER IF LEVEL 4 */}
          {((analysisResult?.severity || selectedPattern.sampleSeverity).includes('4') ||
            (analysisResult?.severity || selectedPattern.sampleSeverity).includes('Safety')) && (
            <div className="p-4 rounded-2xl bg-gradient-to-r from-red-950 via-rose-900 to-red-950 border-2 border-red-500 text-white shadow-xl animate-pulse space-y-2">
              <div className="flex items-center gap-2 text-sm sm:text-base font-black text-red-200">
                <ShieldAlert size={20} className="text-red-300" />
                <span>CẢNH BÁO TỐI CẤP: LEVEL 4 — SAFETY-RELEVANT (NGUY CƠ CHÁY NỔ MẢNG PIN)</span>
              </div>
              <p className="text-xs text-red-100 leading-relaxed">
                Mức chênh lệch nhiệt độ ΔT = {analysisResult?.delta_t ?? calculatedDeltaT}°C hoặc xuất hiện lỗi quá nhiệt Junction Box / Ngược cực tính chuỗi. Tách cô lập ngay chuỗi pin tại tủ Combiner Box để ngăn ngừa đánh thủng màng lưng cách điện và sinh hồ quang DC (Arc Fault)!
              </p>
            </div>
          )}

          {/* ANALYSIS REPORT DISPLAY */}
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 shadow-lg space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-800 pb-3">
              <div className="space-y-0.5">
                <h3 className="text-base font-bold text-white flex items-center gap-2">
                  <Sparkles size={16} className="text-amber-400" />
                  Kết Quả Đánh Giá Chuyên Gia (AI Report & Matrix)
                </h3>
                <p className="text-xs text-slate-400">
                  Chuẩn hóa theo IEC 62446-3, NETA ATS-2025 & Volateq / Sitemark
                </p>
              </div>

              <div className="flex items-center gap-2">
                {analysisResult && (
                  <>
                    <button
                      onClick={() => setShowJsonModal(true)}
                      className="px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 flex items-center gap-1 transition-all"
                      title="Xem dữ liệu JSON 2 chiều"
                    >
                      <FileCode size={13} />
                      <span>JSON</span>
                    </button>

                    <button
                      onClick={handleSaveToCmms}
                      className="px-3 py-1.5 rounded-lg text-xs font-semibold bg-emerald-600 hover:bg-emerald-500 text-white flex items-center gap-1 transition-all shadow-xs"
                    >
                      <CheckCircle2 size={13} />
                      <span>Lưu Báo Cáo</span>
                    </button>
                  </>
                )}
              </div>
            </div>

            {/* If analyzed, show the full markdown and structured view */}
            {analysisResult ? (
              <div className="space-y-4">
                {/* Structured Overview Card */}
                <div className="p-4 rounded-xl bg-slate-950 border border-slate-800 space-y-3">
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 text-center">
                    <div className="p-2 rounded-lg bg-slate-900 border border-slate-800">
                      <div className="text-[10px] text-slate-400">Mức Chênh Lệch ΔT</div>
                      <div className="text-base font-mono font-bold text-amber-400">
                        {analysisResult.delta_t}°C
                      </div>
                    </div>

                    <div className="p-2 rounded-lg bg-slate-900 border border-slate-800">
                      <div className="text-[10px] text-slate-400">Dạng Bất Thường</div>
                      <div className="text-xs font-bold text-white truncate">
                        {analysisResult.json_data?.anomaly_type || selectedPattern.name}
                      </div>
                    </div>

                    <div className="p-2 rounded-lg bg-slate-900 border border-slate-800">
                      <div className="text-[10px] text-slate-400">Vị Trí Ảnh Hưởng</div>
                      <div className="text-[11px] font-semibold text-slate-200 truncate">
                        {analysisResult.json_data?.technical_details?.affected_component || 'Cell / Substring'}
                      </div>
                    </div>

                    <div className="p-2 rounded-lg bg-slate-900 border border-slate-800">
                      <div className="text-[10px] text-slate-400">Hành Động O&M</div>
                      <div className="text-xs font-bold text-cyan-300 truncate">
                        {analysisResult.recommended_action}
                      </div>
                    </div>
                  </div>

                  {/* Technical Risk Description */}
                  {analysisResult.json_data?.technical_details?.risk_description && (
                    <div className="p-3 rounded-lg bg-slate-900/80 border border-slate-800/80 text-xs text-slate-300 leading-relaxed">
                      <strong className="text-white">Đánh giá rủi ro: </strong>
                      {analysisResult.json_data.technical_details.risk_description}
                    </div>
                  )}

                  {/* Field Instructions */}
                  {analysisResult.json_data?.technical_details?.field_instructions && (
                    <div className="p-3 rounded-lg bg-indigo-950/30 border border-indigo-500/30 text-xs text-indigo-200 leading-relaxed space-y-1">
                      <strong className="text-indigo-300 flex items-center gap-1.5">
                        <Activity size={14} />
                        Hướng dẫn hành động cho Kỹ sư Field Service (FSE):
                      </strong>
                      <div className="whitespace-pre-line text-slate-300">
                        {analysisResult.json_data.technical_details.field_instructions}
                      </div>
                    </div>
                  )}
                </div>

                {/* Raw Markdown Report Body */}
                <div className="p-4 rounded-xl bg-slate-950 border border-slate-800 prose prose-invert max-w-none text-xs leading-relaxed overflow-x-auto">
                  <div className="whitespace-pre-line text-slate-200 font-sans">
                    {analysisResult.markdown_report}
                  </div>
                </div>
              </div>
            ) : (
              <div className="py-12 px-4 text-center rounded-xl bg-slate-950/50 border border-dashed border-slate-800 space-y-3">
                <Flame size={36} className="mx-auto text-amber-500 opacity-60" />
                <div className="space-y-1">
                  <h4 className="text-sm font-bold text-slate-200">
                    Sẵn Sàng Phân Tích Chuyên Sâu Tấm Pin & Cell Mặt Trời
                  </h4>
                  <p className="text-xs text-slate-400 max-w-md mx-auto">
                    Tải ảnh nhiệt IR / RGB hoặc chọn một dạng mẫu lỗi từ thư viện bên trái, sau đó bấm <strong>"Phân Tích Chuyên Sâu Tấm Pin & Cell"</strong> để AI đánh giá 3 chiều (Severity, Likely Cause, Recommended Action) theo chuẩn IEC 62446-3 & Sitemark.
                  </p>
                </div>
                <button
                  onClick={handleRunAnalysis}
                  className="px-4 py-2 rounded-xl text-xs font-bold bg-amber-500 hover:bg-amber-400 text-slate-950 transition-all shadow-md"
                >
                  Chạy Phân Tích Mẫu Hiện Tại
                </button>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* JSON MODAL */}
      {showJsonModal && (
        <div className="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-700 rounded-2xl max-w-2xl w-full p-5 space-y-4 shadow-2xl">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <h3 className="text-base font-bold text-white flex items-center gap-2">
                <FileCode size={18} className="text-amber-400" />
                Dữ Liệu JSON Chuẩn Hóa Hệ Thống (IEC 62446-3 & Volateq)
              </h3>
              <button
                onClick={() => setShowJsonModal(false)}
                className="text-slate-400 hover:text-white text-sm"
              >
                ✕ Đóng
              </button>
            </div>

            <div className="bg-slate-950 rounded-xl p-3 border border-slate-800 max-h-96 overflow-y-auto">
              <pre className="text-xs font-mono text-emerald-400 whitespace-pre-wrap">
                {JSON.stringify(analysisResult?.json_data, null, 2)}
              </pre>
            </div>

            <div className="flex items-center justify-between pt-2">
              <button
                onClick={handleCopyJson}
                className="px-3 py-2 rounded-xl text-xs font-semibold bg-indigo-600 hover:bg-indigo-500 text-white flex items-center gap-1.5 transition-all"
              >
                {copied ? <Check size={14} /> : <Copy size={14} />}
                <span>{copied ? 'Đã Sao Chép!' : 'Sao Chép JSON'}</span>
              </button>

              <button
                onClick={() => setShowJsonModal(false)}
                className="px-4 py-2 rounded-xl text-xs font-semibold bg-slate-800 hover:bg-slate-700 text-slate-300"
              >
                Đóng
              </button>
            </div>
          </div>
        </div>
      )}

      {/* PRINT REPORT MODAL */}
      {showPrintModal && analysisResult && (
        <TevReportPrintModal
          isOpen={showPrintModal}
          onClose={() => setShowPrintModal(false)}
          equipmentName={`Mô-đun Pin Mặt Trời (PV Cell): ${moduleTag}`}
          equipmentTag={moduleTag}
          projectTitle={projectName}
          clientName="TEV Energy & Solar Operation Systems"
          standardReference="IEC 62446-3 & NETA ATS-2025 Section 7.29 (Volateq / Sitemark)"
          inspectorName={fseName}
          inspectionDate={testDate}
          reportSummary={{
            overallStatus: analysisResult.severity.includes('Safety')
              ? 'FAIL'
              : analysisResult.severity.includes('Significant')
              ? 'INVESTIGATE'
              : 'PASS',
            statusReason: `Đánh giá nhiệt chuyên sâu PV Cell: ${selectedPattern.name} (ΔT = ${analysisResult.delta_t}°C, Mức độ: ${analysisResult.severity})`,
            evaluationItems: [
              {
                title: 'Mức chênh lệch nhiệt độ (Delta T)',
                measuredValue: `ΔT = ${analysisResult.delta_t}°C (T_max = ${tMax}°C, T_ref = ${tRef}°C)`,
                standardReference: 'IEC 62446-3 (Ngưỡng Hotspot đơn ΔT > 10°C, Safety ΔT ≥ 40°C)',
                status: analysisResult.severity.includes('Safety')
                  ? 'FAIL'
                  : analysisResult.severity.includes('Significant')
                  ? 'INVESTIGATE'
                  : 'PASS',
                notes: `Dạng bất thường: ${selectedPattern.vietnameseName}`
              },
              {
                title: 'Chiều 1: Mức Độ Nghiêm Trọng (Severity)',
                measuredValue: analysisResult.severity,
                standardReference: 'Volateq / Sitemark Cumulative Scale (Level 1 to 4)',
                status: analysisResult.severity.includes('Safety')
                  ? 'FAIL'
                  : analysisResult.severity.includes('Significant')
                  ? 'INVESTIGATE'
                  : 'PASS',
                notes: analysisResult.severity.includes('Safety')
                  ? 'Mất an toàn / Nguy cơ cháy nổ cao'
                  : 'Đáng chú ý / Cần sửa chữa'
              },
              {
                title: 'Chiều 2: Nguyên Nhân Khả Thi (Likely Cause)',
                measuredValue: analysisResult.likely_cause,
                standardReference: 'Precedence: Not installed ➔ Wiring ➔ Module ➔ Shading ➔ Soiling ➔ Unknown',
                status: 'PASS',
                notes: `Cơ chế: ${analysisResult.json_data?.technical_details?.affected_component || 'Cell defect'}`
              },
              {
                title: 'Chiều 3: Hành Động Khuyến Nghị (Recommended Action)',
                measuredValue: analysisResult.recommended_action,
                standardReference: 'First-match: Replace ➔ Repair ➔ Quick fix ➔ Field check ➔ Clean ➔ Monitor',
                status: 'PASS',
                notes: analysisResult.json_data?.technical_details?.field_instructions || 'Xem chi tiết báo cáo'
              }
            ]
          }}
          markdownReport={analysisResult.markdown_report}
        />
      )}
    </div>
  );
};
