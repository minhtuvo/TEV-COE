import React, { useState } from 'react';
import {
  Droplets,
  Thermometer,
  Flame,
  Search,
  Filter,
  CheckCircle2,
  Sparkles,
  ChevronRight,
  ShieldCheck,
  Layers,
  Wrench,
  Clock,
  ArrowRight,
  Zap,
  Activity,
  Sliders,
  Radio,
  Cpu,
  Cable,
  Power,
  ChevronDown,
  LayoutGrid,
  Gauge
} from 'lucide-react';

export type MtsChecklistMode = 'transformer' | 'instrument_ct' | 'instrument_vt' | 'instrument_ccvt' | 'busway' | 'thermography' | 'switchgear' | 'cables' | 'grounding' | 'relays' | 'motors';

export interface MtsEquipmentItem {
  id: MtsChecklistMode;
  category: 'transformers' | 'thermography' | 'switchgear' | 'cables' | 'grounding' | 'relays' | 'motors';
  title: string;
  titleEn: string;
  shortTitle: string;
  shortTitleEn: string;
  subtitle: string;
  scope: string;
  standardCode: string;
  standardName: string;
  aiEngine: string;
  icon: React.ComponentType<{ size?: number; className?: string }>;
  colorGradient: string;
  borderColor: string;
  tagColor: string;
  features: string[];
  status: 'active' | 'coming_soon';
}

export const MTS_EQUIPMENT_CATALOG: MtsEquipmentItem[] = [
  {
    id: 'transformer',
    category: 'transformers',
    title: 'Máy Biến Áp Ngâm Chất Lỏng / Dầu (Liquid-Filled)',
    titleEn: 'Liquid-Filled Power & Distribution Transformers',
    shortTitle: 'Máy Biến Áp Dầu',
    shortTitleEn: 'Liquid Transformer',
    subtitle: 'Đánh Thủng D1816 • Ẩm D1533 • DGA Khí C2H2 (IEEE C57.104) • TTR (≤ 0.5%) • PI (≥ 1.0)',
    scope: 'Máy biến áp ngâm dầu khoáng, Silicone, Ester; trạm 110kV/22kV, công nghiệp',
    standardCode: 'ANSI/NETA MTS-2023 Mục 7.2.2',
    standardName: 'NETA MTS-2023 Sec 7.2.2 & Tables 100.3, 100.4, IEEE C57.104',
    aiEngine: 'Liquid-Filled Transformer MTS AI Specialist (Sec 7.2.2 & DGA)',
    icon: Droplets,
    colorGradient: 'from-amber-600 via-amber-700 to-slate-900',
    borderColor: 'hover:border-amber-400 focus:border-amber-500',
    tagColor: 'bg-amber-500/20 text-amber-800 border-amber-300',
    status: 'active',
    features: [
      'Điện Áp Đánh Thủng Dầu ASTM D1816 (Ngưỡng Bắt Buộc ≥ 40kV)',
      'Hàm Lượng Nước Hòa Tan Karl Fischer ASTM D1533 (≤ 25 ppm)',
      'Phân Tích 7 Khí Cháy DGA IEEE C57.104 (C2H2, H2, CH4, C2H4, C2H6, CO, CO2)',
      'Tỷ Số Biến Áp TTR (Lệch ≤ 0.5%) & Hệ Số Phân Cực PI (≥ 1.0)'
    ]
  },
  {
    id: 'instrument_ct',
    category: 'transformers',
    title: 'Biến Dòng Điện Đo Lường & Bảo Vệ (Current Transformers - CT)',
    titleEn: 'Instrument Transformers: Current Transformers (CT)',
    shortTitle: 'Biến Dòng CT',
    shortTitleEn: 'Current Transformers',
    subtitle: 'NETA MTS Sec 7.10.1 • Bảng 100.5, 100.9, 100.12, 100.18 • IEEE C57.13 • Cực Tính • Knee-point • Nối Đất 1 Điểm',
    scope: 'Biến dòng trung thế 15kV/24kV/35kV, biến dòng cao thế 110kV/220kV, tủ máy cắt, ngăn lộ xuất tuyến',
    standardCode: 'ANSI/NETA MTS-2023 Mục 7.10.1',
    standardName: 'NETA MTS-2023 Sec 7.10.1 & IEEE C57.13 / C57.13.1 / C57.13.3',
    aiEngine: 'Instrument CT MTS AI Specialist (NETA MTS 7.10.1)',
    icon: Gauge,
    colorGradient: 'from-indigo-600 via-indigo-700 to-slate-900',
    borderColor: 'hover:border-indigo-400 focus:border-indigo-500',
    tagColor: 'bg-indigo-500/20 text-indigo-800 border-indigo-300',
    status: 'active',
    features: [
      'Đo Điện Trở Cách Điện Cuộn Sơ & Thứ Cấp (NETA Table 100.5)',
      'Thử Cực Tính H1-X1 & Sai Số Tỉ Số Biến Dòng (IEEE C57.13)',
      'Đường Cong Từ Hóa Excitation & Điện Áp Knee-Point (IEEE C57.13.1)',
      'Xác Nhận DUY NHẤT 1 Điểm Nối Đất Mạch Thứ Cấp (IEEE C57.13.3)'
    ]
  },
  {
    id: 'instrument_vt',
    category: 'transformers',
    title: 'Biến Điện Áp Đo Lường & Bảo Vệ (Voltage Transformers - VT/PT)',
    titleEn: 'Instrument Transformers: Voltage Transformers (VT/PT)',
    shortTitle: 'Biến Điện Áp (VT/PT)',
    shortTitleEn: 'Voltage Transformers',
    subtitle: 'NETA MTS Sec 7.10.2 • Bảng 100.5, 100.9, 100.12, 100.18 • IEEE C57.13 • Tỉ Số Vòng Dây • Cực Tính • Cầu Chì Sơ Cấp',
    scope: 'Biến điện áp trung thế 15kV/22kV/35kV, biến điện áp tụ CVT cao thế 110kV/220kV/500kV',
    standardCode: 'ANSI/NETA MTS-2023 Mục 7.10.2',
    standardName: 'NETA MTS-2023 Sec 7.10.2 & IEEE C57.13 / C57.13.1 / C57.13.3',
    aiEngine: 'Instrument VT/PT MTS AI Specialist (NETA MTS 7.10.2)',
    icon: Zap,
    colorGradient: 'from-sky-600 via-sky-700 to-slate-900',
    borderColor: 'hover:border-sky-400 focus:border-sky-500',
    tagColor: 'bg-sky-500/20 text-sky-800 border-sky-300',
    status: 'active',
    features: [
      'Đo Điện Trở Cách Điện Cuộn Sơ & Thứ Cấp (NETA Table 100.5)',
      'Thử Nghiệm Tỉ Số Biến Turns-Ratio & Cực Tính H1-X1 (IEEE C57.13)',
      'Kiểm Tra Cầu Chì Sơ Cấp Current-Limiting & Tiếp Điểm Ngàm',
      'Xác Nhận DUY NHẤT 1 Điểm Nối Đất Mạch Thứ Cấp (IEEE C57.13.3)'
    ]
  },
  {
    id: 'instrument_ccvt',
    category: 'transformers',
    title: 'Biến Điện Áp Kiểu Tụ (Coupling-Capacitor Voltage Transformers - CCVT)',
    titleEn: 'Instrument Transformers: Coupling-Capacitor Voltage Transformers (CCVT)',
    shortTitle: 'Biến Điện Áp Tụ (CCVT)',
    shortTitleEn: 'Capacitor Voltage Transformers',
    subtitle: 'NETA MTS Sec 7.10.3 • Bảng 100.5, 100.9, 100.12, 100.18 • IEEE C93.1 • Đo C1/C2 (≤ ±5%) • Tan Delta (≤ 0.5%) • Bẫy Sóng PLC',
    scope: 'Trạm biến áp cao thế & siêu cao thế 110kV/220kV/500kV, xuất tuyến đường dây truyền tải, ghép tín hiệu truyền thông PLC',
    standardCode: 'ANSI/NETA MTS-2023 Mục 7.10.3',
    standardName: 'NETA MTS-2023 Sec 7.10.3 & IEEE C93.1 / C57.13.3',
    aiEngine: 'CCVT MTS AI Specialist (NETA MTS 7.10.3 & IEEE C93.1)',
    icon: Radio,
    colorGradient: 'from-violet-600 via-purple-700 to-slate-900',
    borderColor: 'hover:border-purple-400 focus:border-purple-500',
    tagColor: 'bg-purple-500/20 text-purple-800 border-purple-300',
    status: 'active',
    features: [
      'Đo Điện Dung Tầng Tụ C1 & C2 (Dung Sai ≤ ±5% Giá Trị Danh Định, IEEE C93.1)',
      'Tổn Hao Cách Điện Power Factor Tan Delta C1/C2/IVT (≤ 0.5% ở 20°C)',
      'Kiểm Tra Cơ Học Dao Tiếp Địa CCVT & Cuộn Cản Xả Bẫy Sóng PLC (Drain Coil)',
      'Xác Nhận DUY NHẤT 1 Điểm Nối Đất Mạch Thứ Cấp (IEEE C57.13.3)'
    ]
  },
  {
    id: 'thermography',
    category: 'thermography',
    title: 'Khảo Sát & Kiểm Tra Hồng Ngoại Thiết Bị Điện (Thermographic Survey)',
    titleEn: 'Electrical Equipment Thermographic Survey & Infrared Inspection',
    shortTitle: 'Khảo Sát Hồng Ngoại (IR)',
    shortTitleEn: 'Thermographic Survey',
    subtitle: 'ANSI/NETA MTS Section 9 • Bảng NETA Table 100.18 • NFPA 70B • Phân Cấp 4 Mức Sai Lệch (Level 1 - 4)',
    scope: 'Tủ phân phối tổng MSB, thanh cái busbar, mối nối bu-lông, đầu cốt cáp động lực, thiết bị đóng cắt mang tải, máy biến áp',
    standardCode: 'ANSI/NETA MTS-2023 Mục 9 & Bảng 100.18',
    standardName: 'NETA MTS-2023 Section 9 & Table 100.18 (NFPA 70B)',
    aiEngine: 'NETA MTS Thermographic Survey AI Specialist (Table 100.18)',
    icon: Thermometer,
    colorGradient: 'from-rose-600 via-rose-700 to-slate-900',
    borderColor: 'hover:border-rose-400 focus:border-rose-500',
    tagColor: 'bg-rose-500/20 text-rose-800 border-rose-300',
    status: 'active',
    features: [
      'Chênh Lệch Nhiệt Pha Tương Tự ΔT_comp (Level 1: 1-3°C; Level 2: 4-15°C; Level 4: > 15°C)',
      'Chênh Lệch Nhiệt Môi Trường ΔT_amb (Level 1: 1-10°C; Level 2: 11-20°C; Level 3: 21-40°C; Level 4: > 40°C)',
      'Phụ Tải Mạch Tối Thiểu ≥ 40% Định Mức (NETA MTS Section 9.3.3 & NFPA 70B)',
      'Khuyến Nghị Cấp Bách NETA Table 100.18 (Immediate Isolation & Repair)'
    ]
  },
  {
    id: 'switchgear',
    category: 'switchgear',
    title: 'Tủ Phân Phối & Máy Cắt Đóng Cắt (Switchgear & Circuit Breakers)',
    titleEn: 'Switchgear Assemblies & Power Circuit Breakers',
    shortTitle: 'Tủ Đóng Cắt & Máy Cắt',
    shortTitleEn: 'Switchgear & Breakers',
    subtitle: 'Điện Trở Tiếp Xúc DLRO (Table 100.12) • Thời Gian Đóng Cắt Timing • Cách Điện Thanh Cái',
    scope: 'Tủ trung thế 24kV, tủ MSB hạ thế, máy cắt chân không VCB, máy cắt không khí ACB',
    standardCode: 'ANSI/NETA MTS-2023 Mục 7.1 & 7.6',
    standardName: 'NETA MTS-2023 Sec 7.1, 7.6 & Table 100.1',
    aiEngine: 'Switchgear & Breaker MTS AI Engine',
    icon: Power,
    colorGradient: 'from-cyan-600 via-cyan-700 to-slate-900',
    borderColor: 'hover:border-cyan-400 focus:border-cyan-500',
    tagColor: 'bg-cyan-500/20 text-cyan-800 border-cyan-300',
    status: 'active',
    features: [
      'Điện Trở Tiếp Xúc Mối Nối DLRO Micro-ohm (NETA Table 100.12)',
      'Điện Trở Cách Điện Megohm Thanh Cái & Cực Tiếp Điểm',
      'Đặc Tuyến Thời Gian Đóng Cắt (Open/Close/Trip Timing per Table 100.7)'
    ]
  },
  {
    id: 'busway',
    category: 'switchgear',
    title: 'Thanh Cái Dẫn Điện Vỏ Kim Loại (Metal-Enclosed Busways - LV/MV)',
    titleEn: 'Metal-Enclosed Busways (Low- & Medium-Voltage, < 1 kV & 1-35 kV)',
    shortTitle: 'Thanh Cái Busway',
    shortTitleEn: 'Busways (LV/MV)',
    subtitle: 'NETA MTS Sec 7.4 • Bảng 100.1, 100.12, 100.17, 100.18, 100.23 • Điện Trở Mối Nối Joint (≤ 50%) • Thử Cao Áp Hi-Pot • Khảo Sát PD Online',
    scope: 'Tuyến Busway hạ thế MSB < 1kV, Busway trung thế 1kV - 35kV (15kV/24kV), trục đứng riser tòa nhà, xuất tuyến máy biến áp',
    standardCode: 'ANSI/NETA MTS-2023 Mục 7.4',
    standardName: 'NETA MTS-2023 Sec 7.4, IEEE C37.23, NEMA BU1',
    aiEngine: 'Metal-Enclosed Busways MTS AI Specialist (NETA MTS 7.4)',
    icon: Layers,
    colorGradient: 'from-teal-600 via-teal-700 to-slate-900',
    borderColor: 'hover:border-teal-400 focus:border-teal-500',
    tagColor: 'bg-teal-500/20 text-teal-800 border-teal-300',
    status: 'active',
    features: [
      'Đo Điện Trở Mối Nối Bus Joint Bằng Micro-ohmmeter Dòng Cao (Độ Lệch ≤ 50%, NETA MTS 7.4.D.1)',
      'Đo Điện Trở Cách Điện Megohm (NETA Table 100.1, Min ≥ 1,000 MΩ cho Hệ 15kV)',
      'Thử Nghiệm Chịu Đựng Điện Áp Tăng Cao Hi-Pot AC/DC 1 Phút (NETA Table 100.17)',
      'Khảo Sát Phóng Điện Cục Bộ Online TEV / Acoustic Mối Nối Bus Joint (NETA Table 100.23)'
    ]
  },
  {
    id: 'cables',
    category: 'cables',
    title: 'Cáp Điện Lực Trung & Hạ Thế (Power Cables)',
    titleEn: 'Medium & Low Voltage Power Cables',
    shortTitle: 'Cáp Điện Lực',
    shortTitleEn: 'Power Cables',
    subtitle: 'VLF Tan Delta • Điện Trở Cách Điện Megohm • Thử Điện Áp Cao VLF Withstand (IEEE 400.2)',
    scope: 'Cáp ngầm trung thế 22kV/35kV Cu/XLPE, cáp hạ thế 0.6/1kV',
    standardCode: 'ANSI/NETA MTS-2023 Mục 7.3.2 & 7.3.3',
    standardName: 'NETA MTS-2023 Sec 7.3.2, 7.3.3 & IEEE 400.2',
    aiEngine: 'Power Cables MTS AI Specialist',
    icon: Cable,
    colorGradient: 'from-blue-600 via-blue-700 to-slate-900',
    borderColor: 'hover:border-blue-400 focus:border-blue-500',
    tagColor: 'bg-blue-500/20 text-blue-800 border-blue-300',
    status: 'active',
    features: [
      'Đo Điện Trở Cách Điện Megohm 1 Phút / 10 Phút',
      'Thử Nghiệm Tổn Hao Điện Môi VLF Tan Delta (IEEE 400.2)',
      'Kiểm Tra Vỏ Bọc Ngoại Shield Continuity'
    ]
  },
  {
    id: 'grounding',
    category: 'grounding',
    title: 'Hệ Thống Tiếp Địa & Nối Đất Trạm (Grounding Systems)',
    titleEn: 'Substation & Facility Grounding Systems',
    shortTitle: 'Tiếp Địa Trạm',
    shortTitleEn: 'Grounding Systems',
    subtitle: 'Đo Điện Trở Nối Đất 3 Cọc Fall-of-Potential (IEEE Std 81) • Đẳng Thế Equipotential Bonding',
    scope: 'Lưới tiếp địa trạm biến áp 110kV/22kV, bãi tiếp địa công nghiệp, cột thu lôi',
    standardCode: 'ANSI/NETA MTS-2023 Mục 7.13',
    standardName: 'NETA MTS-2023 Sec 7.13 & IEEE Std 81',
    aiEngine: 'Grounding Systems MTS AI Engine',
    icon: Zap,
    colorGradient: 'from-emerald-600 via-emerald-700 to-slate-900',
    borderColor: 'hover:border-emerald-400 focus:border-emerald-500',
    tagColor: 'bg-emerald-500/20 text-emerald-800 border-emerald-300',
    status: 'active',
    features: [
      'Điện Trở Nối Đất Fall-of-Potential (Trạm Biến Áp ≤ 1.0 Ω, Công Nghiệp ≤ 5.0 Ω)',
      'Điện Trở Đẳng Thế Điểm Nối Tiếp Địa (Bắt Buộc ≤ 0.50 Ω)',
      'Kiểm Tra Ăn Mòn Cọc Tiếp Địa & Mối Hàn Hóa Nhiệt Cadweld'
    ]
  },
  {
    id: 'relays',
    category: 'relays',
    title: 'Rơ-le Bảo Vệ Kỹ Thuật Số & Điện Cơ (Protective Relays)',
    titleEn: 'Digital & Electromechanical Protective Relays',
    shortTitle: 'Rơ-le Bảo Vệ',
    shortTitleEn: 'Protective Relays',
    subtitle: 'Thử Nghiệm Tiêm Dòng Thứ Cấp • Quá Dòng 50/51 • So Lệch 87 • Chạm Đất 51N • Khoảng Cách 21',
    scope: 'Rơ-le số SEL, ABB, Schneider, Siemens bảo vệ máy biến áp, lộ xuất tuyến, thanh cái',
    standardCode: 'ANSI/NETA MTS-2023 Mục 7.9',
    standardName: 'NETA MTS-2023 Sec 7.9 & Table 100.9',
    aiEngine: 'Protective Relays MTS AI Specialist',
    icon: Cpu,
    colorGradient: 'from-purple-600 via-purple-700 to-slate-900',
    borderColor: 'hover:border-purple-400 focus:border-purple-500',
    tagColor: 'bg-purple-500/20 text-purple-800 border-purple-300',
    status: 'active',
    features: [
      'Kiểm Tra Ngưỡng Tác Động Pickup & Thời Gian Trễ Timing',
      'Đặc Tuyến Thời Gian Ngược Inverse Time Curve (NETA Table 100.9)',
      'Thử Nghiệm Trip Test Cắt Trực Tiếp Máy Cắt Liên Động'
    ]
  },
  {
    id: 'motors',
    category: 'motors',
    title: 'Động Cơ Điện & Máy Điện Quay (Rotating Machinery & Motors)',
    titleEn: 'AC Motors & Rotating Electrical Machinery',
    shortTitle: 'Động Cơ Điện',
    shortTitleEn: 'Motors & Rotating',
    subtitle: 'Đo Điện Trở Cách Điện Cuộn Dây • Hệ Số Hấp Thụ DAR • Hệ Số Phân Cực PI • Điện Trở Một Chiều',
    scope: 'Động cơ trung thế 3.3kV/6.6kV, động cơ bơm quạt hạ thế 380V công suất lớn',
    standardCode: 'ANSI/NETA MTS-2023 Mục 7.15',
    standardName: 'NETA MTS-2023 Sec 7.15 & Table 100.11',
    aiEngine: 'Rotating Machinery MTS AI Specialist',
    icon: Activity,
    colorGradient: 'from-orange-600 via-orange-700 to-slate-900',
    borderColor: 'hover:border-orange-400 focus:border-orange-500',
    tagColor: 'bg-orange-500/20 text-orange-800 border-orange-300',
    status: 'active',
    features: [
      'Điện Trở Cách Điện Megohm & Chỉ Số Phân Cực PI (≥ 2.0 per IEEE 43)',
      'Cân Bằng Điện Trở Cuộn Dây Stator 3 Pha (Độ Lệch ≤ 2%)',
      'Kiểm Tra Độ Rung Cơ Học & Nhiệt Độ Vòng Bi Bạc Đạn'
    ]
  }
];

interface Props {
  currentMode: MtsChecklistMode;
  onSelectMode: (mode: MtsChecklistMode) => void;
  lang?: 'vi' | 'en';
  isCollapsed?: boolean;
  onToggleCollapse?: () => void;
}

export const MtsEquipmentChecklistSelector: React.FC<Props> = ({
  currentMode,
  onSelectMode,
  lang = 'vi',
  isCollapsed: externalIsCollapsed,
  onToggleCollapse: externalOnToggleCollapse
}) => {
  const [internalIsCatalogOpen, setInternalIsCatalogOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedCategory, setSelectedCategory] = useState<string>('all');

  const isVi = lang === 'vi';

  // Toggle state
  const isCatalogOpen = externalIsCollapsed !== undefined ? !externalIsCollapsed : internalIsCatalogOpen;
  const toggleCatalog = () => {
    if (externalOnToggleCollapse) {
      externalOnToggleCollapse();
    } else {
      setInternalIsCatalogOpen(!internalIsCatalogOpen);
    }
  };

  const categories = [
    { id: 'all', label: isVi ? 'Tất cả thiết bị MTS' : 'All MTS Equipment', count: MTS_EQUIPMENT_CATALOG.length },
    { id: 'transformers', label: isVi ? 'Biến Áp, Biến Dòng, VT & Tụ CCVT' : 'Transformers, CTs, VTs & CCVTs', count: 4 },
    { id: 'thermography', label: isVi ? 'Khảo Sát Nhiệt IR' : 'Thermography', count: 1 },
    { id: 'switchgear', label: isVi ? 'Đóng Cắt & Busway' : 'Switchgear & Busways', count: 2 },
    { id: 'cables', label: isVi ? 'Cáp Điện Lực' : 'Power Cables', count: 1 },
    { id: 'grounding', label: isVi ? 'Tiếp Địa Trạm' : 'Grounding', count: 1 },
    { id: 'relays', label: isVi ? 'Rơ-le Bảo Vệ' : 'Relays', count: 1 },
    { id: 'motors', label: isVi ? 'Động Cơ Điện' : 'Motors', count: 1 },
  ];

  const filteredItems = MTS_EQUIPMENT_CATALOG.filter(item => {
    const matchesSearch =
      item.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
      item.titleEn.toLowerCase().includes(searchTerm.toLowerCase()) ||
      item.subtitle.toLowerCase().includes(searchTerm.toLowerCase()) ||
      item.standardCode.toLowerCase().includes(searchTerm.toLowerCase()) ||
      item.scope.toLowerCase().includes(searchTerm.toLowerCase());

    const matchesCategory = selectedCategory === 'all' || item.category === selectedCategory;

    return matchesSearch && matchesCategory;
  });

  const currentItem = MTS_EQUIPMENT_CATALOG.find(i => i.id === currentMode) || MTS_EQUIPMENT_CATALOG[0];

  return (
    <div className="bg-white rounded-2xl border border-slate-200 shadow-md overflow-hidden transition-all">
      {/* ------------------------------------------------------------------ */}
      {/* 1. HORIZONTAL TABS BAR: FAST 1-CLICK EQUIPMENT SWITCHER & TAB "ĐỔI THIẾT BỊ BẢO TRÌ" */}
      {/* ------------------------------------------------------------------ */}
      <div className="bg-slate-950 p-2 sm:p-2.5 border-b border-amber-500/30 flex items-center justify-between gap-2 overflow-x-auto scrollbar-none">
        <div className="flex items-center gap-1.5 shrink-0">
          <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-amber-500 to-amber-600 text-slate-950 flex items-center justify-center font-black text-xs shrink-0 shadow-sm">
            MTS
          </div>
          <span className="text-[11px] font-black text-amber-300 uppercase tracking-wider px-2 py-1 bg-amber-500/10 rounded-md border border-amber-500/20 hidden md:inline">
            NETA MTS-2023
          </span>
        </div>

        {/* The Fast Horizontal Tabs */}
        <div className="flex items-center gap-1.5 overflow-x-auto scrollbar-none py-0.5">
          {MTS_EQUIPMENT_CATALOG.map((item) => {
            const isSelected = currentMode === item.id;
            const IconComp = item.icon;
            const isThermal = item.id === 'thermography';

            return (
              <button
                key={item.id}
                type="button"
                onClick={() => onSelectMode(item.id)}
                className={`flex items-center gap-2 px-3 py-2 rounded-xl text-xs font-bold transition-all shrink-0 border ${
                  isSelected
                    ? 'bg-gradient-to-r from-amber-500 to-amber-400 text-slate-950 font-black border-amber-300 shadow-md shadow-amber-500/30 ring-2 ring-amber-400/50'
                    : 'bg-slate-900/90 text-slate-300 hover:text-white hover:bg-slate-800 border-slate-800'
                }`}
                title={`${isVi ? item.title : item.titleEn} (${item.standardCode})`}
              >
                <IconComp size={15} className={isSelected ? 'text-slate-950' : isThermal ? 'text-rose-400' : 'text-amber-400'} />
                <span>{isVi ? item.shortTitle : item.shortTitleEn}</span>
                {isThermal && (
                  <span className={`text-[9px] px-1.5 py-0.2 rounded font-black tracking-wide ${
                    isSelected ? 'bg-slate-950 text-rose-300' : 'bg-rose-500/20 text-rose-300 border border-rose-400/30'
                  }`}>
                    Mục 9
                  </span>
                )}
                {isSelected && <CheckCircle2 size={13} className="text-slate-950 ml-0.5" />}
              </button>
            );
          })}
        </div>

        {/* Tab "Đổi Thiết Bị Bảo Trì" (Expand/Collapse Full Catalog) */}
        <button
          type="button"
          onClick={toggleCatalog}
          className={`px-3 py-2 rounded-xl text-xs font-black transition-all border flex items-center gap-1.5 shrink-0 shadow-xs ${
            isCatalogOpen
              ? 'bg-amber-500 text-slate-950 border-amber-300 shadow-md shadow-amber-500/20'
              : 'bg-slate-800 hover:bg-slate-700 text-amber-300 border-amber-500/40 hover:text-white'
          }`}
          title={isVi ? "Xem danh mục và bổ sung thiết bị bảo trì MTS" : "Browse & add maintenance equipment"}
        >
          <LayoutGrid size={14} />
          <span className="hidden sm:inline">
            {isCatalogOpen ? (isVi ? 'Đóng Danh Mục' : 'Close Catalog') : (isVi ? '🔄 Đổi Thiết Bị Bảo Trì' : '🔄 Change Equipment')}
          </span>
          <span className="sm:hidden">{isVi ? 'Đổi TB' : 'Catalog'}</span>
          <span className="text-[10px] px-1.5 py-0.2 rounded-full bg-slate-950/40 font-bold">
            {MTS_EQUIPMENT_CATALOG.length}
          </span>
        </button>
      </div>

      {/* ------------------------------------------------------------------ */}
      {/* 2. ACTIVE EQUIPMENT SUMMARY STRIP */}
      {/* ------------------------------------------------------------------ */}
      <div className="px-5 py-3.5 bg-gradient-to-r from-slate-900 via-slate-800 to-slate-900 text-white flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3 border-b border-slate-700/60">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-xl bg-amber-500/20 text-amber-400 border border-amber-400/30 flex items-center justify-center shrink-0">
            {React.createElement(currentItem.icon, { size: 18 })}
          </div>
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-xs font-black text-white">
                {isVi ? currentItem.title : currentItem.titleEn}
              </span>
              <span className="text-[10px] font-mono px-2 py-0.5 rounded bg-amber-500/20 text-amber-300 border border-amber-400/30 font-bold">
                {currentItem.standardCode}
              </span>
              <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-500/20 text-emerald-300 border border-emerald-400/30">
                {isVi ? 'Đang Thực Hiện Thử Nghiệm' : 'Active Testing Protocol'}
              </span>
            </div>
            <p className="text-[11px] text-slate-300 mt-0.5 font-medium line-clamp-1">
              {currentItem.subtitle}
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 shrink-0">
          <button
            type="button"
            onClick={toggleCatalog}
            className="text-xs font-bold text-amber-400 hover:text-amber-300 flex items-center gap-1 transition-colors px-2.5 py-1.5 rounded-lg bg-slate-800/80 hover:bg-slate-700/80 border border-slate-700"
          >
            <span>{isCatalogOpen ? (isVi ? '▲ Ẩn danh mục mở rộng' : '▲ Hide Catalog') : (isVi ? '▼ Mở rộng danh mục thiết bị' : '▼ Expand Catalog')}</span>
          </button>
        </div>
      </div>

      {/* ------------------------------------------------------------------ */}
      {/* 3. EXPANDABLE MTS CATALOG DRAWER (SEARCH & FILTERING) */}
      {/* ------------------------------------------------------------------ */}
      {isCatalogOpen && (
        <div className="p-5 space-y-4 bg-slate-50 border-t border-slate-100 animate-in fade-in slide-in-from-top-2">
          {/* Header Notice for Roadmap / Equipment expansion */}
          <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3 bg-amber-50 border border-amber-200 p-3 rounded-xl text-xs text-amber-900">
            <div className="flex items-center gap-2">
              <Sparkles size={16} className="text-amber-600 shrink-0" />
              <span>
                <strong>{isVi ? 'Trung tâm thiết bị thử nghiệm bảo trì NETA MTS-2023:' : 'NETA MTS-2023 Maintenance Catalog Hub:'}</strong>{' '}
                {isVi
                  ? 'Sẵn sàng bổ sung và mở rộng không giới hạn các thiết bị bảo dưỡng (Máy biến áp, Khảo sát nhiệt IR, Tủ đóng cắt, Cáp, Nối đất, Rơ-le, Động cơ...)'
                  : 'Ready to expand and configure additional maintenance equipment checklists pursuant to ANSI/NETA MTS.'}
              </span>
            </div>
            <span className="text-[11px] font-bold text-amber-700 px-2 py-0.5 rounded bg-amber-100 border border-amber-300 shrink-0 self-start md:self-auto">
              {MTS_EQUIPMENT_CATALOG.length} {isVi ? 'Thiết Bị Đã Tích Hợp' : 'Equipment Integrated'}
            </span>
          </div>

          {/* Search & Category Filter */}
          <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3">
            {/* Search Input */}
            <div className="relative flex-1">
              <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                type="text"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                placeholder={isVi ? "Tìm kiếm thiết bị bảo trì MTS (Ví dụ: biến áp, hồng ngoại, cáp, rơ-le...)" : "Search MTS maintenance checklists..."}
                className="w-full pl-9 pr-4 py-2 bg-white border border-slate-200 rounded-xl text-xs font-medium text-slate-800 focus:outline-hidden focus:ring-2 focus:ring-amber-500 shadow-xs"
              />
            </div>

            {/* Category Filter Pills */}
            <div className="flex items-center gap-1.5 overflow-x-auto pb-1 sm:pb-0 scrollbar-none">
              {categories.map((cat) => (
                <button
                  key={cat.id}
                  onClick={() => setSelectedCategory(cat.id)}
                  className={`px-3 py-1.5 rounded-lg text-[11px] font-bold whitespace-nowrap transition-all flex items-center gap-1 shrink-0 ${
                    selectedCategory === cat.id
                      ? 'bg-slate-900 text-white shadow-xs'
                      : 'bg-white hover:bg-slate-100 text-slate-600 border border-slate-200'
                  }`}
                >
                  <span>{cat.label}</span>
                  <span className={`text-[10px] px-1.5 py-0.2 rounded-full ${
                    selectedCategory === cat.id ? 'bg-amber-400 text-slate-950 font-black' : 'bg-slate-100 text-slate-500'
                  }`}>
                    {cat.count}
                  </span>
                </button>
              ))}
            </div>
          </div>

          {/* Cards Grid */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3.5 pt-1">
            {filteredItems.map((item) => {
              const isSelected = currentMode === item.id;
              const IconComp = item.icon;

              return (
                <div
                  key={item.id}
                  onClick={() => {
                    onSelectMode(item.id);
                    toggleCatalog();
                  }}
                  className={`p-4 rounded-2xl border transition-all relative flex flex-col justify-between text-left cursor-pointer ${
                    isSelected
                      ? 'bg-gradient-to-br from-amber-50 to-white border-amber-400 ring-2 ring-amber-400/40 shadow-md'
                      : 'bg-white hover:bg-slate-50 border-slate-200 hover:border-slate-300 shadow-xs hover:shadow-md'
                  }`}
                >
                  <div className="space-y-2.5">
                    {/* Top Row: Icon, Standard Code & Status */}
                    <div className="flex items-start justify-between gap-2">
                      <div className="flex items-center gap-2.5">
                        <div className={`w-9 h-9 rounded-xl flex items-center justify-center shrink-0 ${
                          isSelected
                            ? 'bg-amber-500 text-slate-950 shadow-sm'
                            : 'bg-slate-900 text-white'
                        }`}>
                          <IconComp size={18} />
                        </div>
                        <div>
                          <span className="text-[10px] font-bold text-amber-700 bg-amber-50 px-2 py-0.5 rounded border border-amber-200/60 font-mono block">
                            {item.standardCode}
                          </span>
                        </div>
                      </div>

                      {isSelected ? (
                        <span className="px-2 py-0.5 rounded-full text-[10px] font-black bg-amber-500 text-slate-950 flex items-center gap-1 shadow-xs">
                          <CheckCircle2 size={11} />
                          {isVi ? 'Đang mở' : 'Selected'}
                        </span>
                      ) : (
                        <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-50 text-emerald-700 border border-emerald-200">
                          {isVi ? 'Sẵn sàng' : 'Active'}
                        </span>
                      )}
                    </div>

                    {/* Title & Subtitle */}
                    <div>
                      <h3 className="font-black text-xs sm:text-sm text-slate-900 leading-snug">
                        {isVi ? item.title : item.titleEn}
                      </h3>
                      <p className="text-[11px] text-slate-500 mt-1 line-clamp-2 leading-relaxed">
                        {item.subtitle}
                      </p>
                    </div>

                    {/* Key Tests */}
                    <div className="space-y-1 pt-1.5 border-t border-slate-100 text-[10px] text-slate-600">
                      {item.features.slice(0, 2).map((feat, idx) => (
                        <div key={idx} className="flex items-start gap-1.5">
                          <span className="w-1.5 h-1.5 rounded-full bg-amber-500 mt-1 shrink-0" />
                          <span className="line-clamp-1">{feat}</span>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Bottom Action Footer */}
                  <div className="pt-3 mt-3 border-t border-slate-100 flex items-center justify-between text-[11px]">
                    <span className="text-slate-400 font-mono text-[10px]">
                      {item.aiEngine.split('(')[0].trim()}
                    </span>
                    <span className={`font-bold flex items-center gap-1 ${
                      isSelected ? 'text-amber-700 font-black' : 'text-slate-600 group-hover:text-slate-900'
                    }`}>
                      <span>{isSelected ? (isVi ? 'Đang thực hiện' : 'Active') : (isVi ? 'Mở biên bản' : 'Open')}</span>
                      <ArrowRight size={12} />
                    </span>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
};
