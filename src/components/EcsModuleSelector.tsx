import React, { useState, useMemo } from 'react';
import {
  Zap,
  Activity,
  Layers,
  CheckCircle2,
  Sparkles,
  Info,
  ShieldCheck,
  ChevronRight,
  ChevronDown,
  LayoutGrid,
  Sun,
  Wind,
  BatteryCharging,
  Power,
  Cpu,
  Search,
  Filter,
  ArrowRight,
  Sliders,
  Radio,
  Clock,
  Flame,
  Gauge,
  ClipboardList,
  AlertCircle,
  GitBranch
} from 'lucide-react';

export type EcsModuleId =
  | 'low-voltage'
  | 'medium-voltage'
  | 'high-voltage'
  | 'solar-pv'
  | 'ats-switch'
  | 'wind-turbine'
  | 'bess'
  | 'engine-generator'
  | 'ups-system'
  | 'integrated-systems';

export type EcsCategory = 'all' | 'section-7' | 'section-8' | 'renewables' | 'backup-storage' | 'integrated';

export interface EcsModuleItem {
  id: EcsModuleId;
  category: EcsCategory;
  sectionCode: string;
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
  badge: string;
  badgeColor: string;
  features: string[];
  isAvailable: boolean;
}

export const ECS_MODULE_CATALOG: EcsModuleItem[] = [
  {
    id: 'low-voltage',
    category: 'section-7',
    sectionCode: 'Section 7.1',
    title: 'Hệ Thống Phân Phối Hạ Thế (< 1 kV)',
    titleEn: 'General Inspection & Test Procedures: Low-Voltage Systems (< 1 kV)',
    shortTitle: 'Hạ Thế (< 1 kV)',
    shortTitleEn: 'LV Systems (< 1 kV)',
    subtitle: 'Cách Điện 1000V (≥ 1MΩ) • Lực Siết Bu-lông Table 100.12 • Tiếp Địa ≤ 5.0Ω • Cân Bằng Pha • Rò Điện GFCI',
    scope: 'Tủ phân phối hạ thế MSB, tủ động lực MCC, tủ chiếu sáng DB, Aptomat ACB/MCCB, Cáp ngầm hạ thế, Ổ cắm GFCI (< 1 kV)',
    standardCode: 'ANSI/NETA ECS-2024 Mục 7.1',
    standardName: 'NETA ECS-2024 Sec 7.1 & Tables 100.1, 100.12, 100.18',
    aiEngine: 'Low-Voltage Commissioning AI Specialist (NETA Sec 7.1)',
    icon: Zap,
    colorGradient: 'from-slate-950 via-teal-950 to-slate-900',
    borderColor: 'hover:border-teal-400 focus:border-teal-500',
    tagColor: 'bg-teal-500/20 text-teal-800 border-teal-300',
    badge: 'Mục 7.1 (< 1kV)',
    badgeColor: 'bg-teal-500/20 text-teal-300 border-teal-400/40',
    features: [
      'Tiền đóng điện: Đo cách điện 1000V DC bắt buộc ≥ 1.0 MΩ (Table 100.1) & kiểm tra liên tục mạch tiếp địa ≤ 5.0 Ω',
      'Siết cờ-lê lực bu-lông kết nối thanh cái đồng/nhôm theo Bảng chuẩn NETA ATS/MTS Table 100.12',
      'Trong đóng điện: Tháo tiếp địa tạm thời, đối soát thứ tự pha A-B-C và kiểm tra điện áp dây/pha',
      'Sau đóng điện: Khảo sát ảnh nhiệt IR mang tải Table 100.18, cân bằng dòng tải các pha & test tác động GFCI'
    ],
    isAvailable: true
  },
  {
    id: 'medium-voltage',
    category: 'section-7',
    sectionCode: 'Section 7.2',
    title: 'Hệ Thống Đóng Cắt & Phân Phối Trung Thế (1 kV - 100 kV)',
    titleEn: 'General Inspection & Test Procedures: Medium-Voltage Systems (> 1 kV & < 100 kV)',
    shortTitle: 'Trung Thế (1kV - 100kV)',
    shortTitleEn: 'MV Systems (1kV - 100kV)',
    subtitle: 'Cách Điện 5000V (≥ 1,000 MΩ) • Cáp VLF 0.1Hz & Tan Delta • Đồng Pha Bus Tie (ΔV ≤ 15V) • Đo PD Online (Table 100.23)',
    scope: 'Tủ phân phối trung thế Metal-Clad 22kV/35kV, tủ GIS, Máy cắt chân không VCB/SF6, Cáp ngầm MV, Máy biến áp lực',
    standardCode: 'ANSI/NETA ECS-2024 Mục 7.2',
    standardName: 'NETA ECS-2024 Sec 7.2 & Tables 100.1, 100.2, 100.12, 100.18, 100.23',
    aiEngine: 'Medium-Voltage Commissioning AI Engine (NETA Sec 7.2)',
    icon: Activity,
    colorGradient: 'from-slate-950 via-emerald-950 to-slate-900',
    borderColor: 'hover:border-emerald-400 focus:border-emerald-500',
    tagColor: 'bg-emerald-500/20 text-emerald-800 border-emerald-300',
    badge: 'Mục 7.2 (1kV-100kV)',
    badgeColor: 'bg-emerald-500/20 text-emerald-300 border-emerald-400/40',
    features: [
      'Tiền đóng điện: Đo cách điện 5000V DC (≥ 1,000 MΩ Table 100.1) & Thử cao áp VLF 0.1Hz Tan Delta (Table 100.2)',
      'Bơm dòng sơ cấp thử nghiệm bảo vệ so lệch 87, quá dòng 50/51 & kiểm tra áp suất khí SF6 / mức chân không VCB',
      'Trong đóng điện: Kiểm tra đồng pha hai nguồn máy cắt liên lạc Bus Tie (ΔV ≤ 15V, góc lệch pha ≤ 5°)',
      'Sau đóng điện: Khảo sát nhiệt hồng ngoại IR Table 100.18 & Đo phóng điện cục bộ trực tuyến Online TEV/Ultrasonic Table 100.23'
    ],
    isAvailable: true
  },
  {
    id: 'high-voltage',
    category: 'section-7',
    sectionCode: 'Section 7.3',
    title: 'Hệ Thống Cao Thế & Siêu Cao Thế (> 100 kV & < 1,000 kV)',
    titleEn: 'General Inspection & Test Procedures: High-Voltage and Extra-High Voltage Systems (> 100 kV & < 1,000 kV)',
    shortTitle: 'Cao Thế & EHV (> 100kV)',
    shortTitleEn: 'HV/EHV Systems (> 100kV)',
    subtitle: 'Cách Điện 10kV (≥ 10,000 MΩ) • Khí SF6 Table 100.19 • Dầu & DGA MBA C57.104 • Rơ-le NERC PRC • Tiếp Địa Trạm ≤ 1.0Ω',
    scope: 'Trạm biến áp 110kV/220kV/500kV, Trạm GIS EHV, Máy cắt khí SF6, Máy biến áp lực EHV 250MVA - 900MVA, Cáp ngầm cao thế',
    standardCode: 'ANSI/NETA ECS-2024 Mục 7.3',
    standardName: 'NETA ECS-2024 Sec 7.3, NETA Tables 100.1-100.23, IEEE 81, IEEE C57.104, NERC PRC',
    aiEngine: 'HV/EHV Commissioning AI Engine (NETA Sec 7.3 & NERC PRC)',
    icon: Zap,
    colorGradient: 'from-slate-950 via-purple-950 to-slate-900',
    borderColor: 'hover:border-purple-400 focus:border-purple-500',
    tagColor: 'bg-purple-500/20 text-purple-800 border-purple-300',
    badge: 'Mục 7.3 (> 100kV)',
    badgeColor: 'bg-purple-500/20 text-purple-300 border-purple-400/40',
    features: [
      'Tiền đóng điện: Đo cách điện 10,000V DC (≥ 10,000 MΩ Table 100.1), kiểm tra chất lượng khí SF6 (độ ẩm ≤ 99 ppmv Table 100.19) & DGA MBA Condition 1',
      'Thử nghiệm rơ-le so lệch 87L, bảo vệ khoảng cách 21 và tuân thủ NERC PRC-002/005/019/024; kiểm tra điện trở lưới đất trạm ≤ 1.0 Ω (IEEE 81)',
      'Trong đóng điện: Xác nhận tháo 100% tiếp địa di động trên thanh cái EHV, kiểm tra đồng pha hai nguồn 230kV tại Bus Tie (ΔV ≤ 60V, lệch pha ≤ 5°)',
      'Sau đóng điện: Khảo sát nhiệt hồng ngoại IR mang tải Table 100.18, đo phóng điện cục bộ trực tuyến Online PD trạm GIS (UHF/Acoustic Table 100.23) & lấy mẫu sau 24h'
    ],
    isAvailable: true
  },
  {
    id: 'solar-pv',
    category: 'renewables',
    sectionCode: 'Section 8.1',
    title: 'Hệ Thống Năng Lượng Mặt Trời (Solar Photovoltaic - PV)',
    titleEn: 'Source Specific Systems Commissioning: Solar Photovoltaic (PV) Procedures',
    shortTitle: 'Điện Mặt Trời (Solar PV)',
    shortTitleEn: 'Solar PV Systems',
    subtitle: 'Cách Điện Khô/Ẩm 1500V DC • Cực Tính & Diode • Độ Lệch Voc (≤ 5%) • Quét I-V STC (≥ 95%) • Chống Đảo Lưới IEEE 1547',
    scope: 'Chuỗi pin PV String 1000V/1500V DC, Cáp DC năng lượng mặt trời, Tủ gom Combiner Box, Inverter Biến Tần hòa lưới, Hệ thống DAS/SCADA',
    standardCode: 'ANSI/NETA ECS-2024 Mục 8.1',
    standardName: 'NETA ECS-2024 Sec 8.1, ATS 7.29, IEC 62446-1, IEEE 1547, NERC MOD-025-2',
    aiEngine: 'Solar PV Commissioning AI Engine (NETA Sec 8.1 & IEC 62446-1)',
    icon: Sun,
    colorGradient: 'from-slate-950 via-amber-950 to-slate-900',
    borderColor: 'hover:border-amber-400 focus:border-amber-500',
    tagColor: 'bg-amber-500/20 text-amber-800 border-amber-300',
    badge: 'Mục 8.1 (Solar PV)',
    badgeColor: 'bg-amber-500/20 text-amber-300 border-amber-400/40',
    features: [
      'Tiền đóng điện: Kiểm tra cực tính DC (+/-) 100%, Cách điện khô/ẩm (1000V ≥ 1.0 MΩ, 1500V ≥ 5.0 MΩ Table 100.1)',
      'Đo điện áp Voc (độ lệch giữa các string ≤ 5%), quét đường cong I-V (Pmax STC ≥ 95% ASTM E1036) & tiếp địa ≤ 5.0 Ω',
      'Trong đóng điện: Kiểm tra thứ tự pha AC ngõ ra Inverter, Thử nghiệm bảo vệ chống đảo lưới tự ngắt ≤ 2.0s (IEEE 1547)',
      'Sau đóng điện: Khảo sát nhiệt hồng ngoại IR cell pin & jack MC4 Table 100.18, Thử nghiệm công suất phản kháng NERC MOD-025-2'
    ],
    isAvailable: true
  },
  {
    id: 'ats-switch',
    category: 'backup-storage',
    sectionCode: 'Section 8.3',
    title: 'Tủ Chuyển Nguồn Tự Động (Automatic Transfer Switch - ATS)',
    titleEn: 'Source Specific Systems Commissioning: Automatic Transfer Switch (ATS) Procedures',
    shortTitle: 'Tủ ATS (Transfer Switch)',
    shortTitleEn: 'ATS Switch',
    subtitle: 'Cách Điện Mạch Điều Khiển (≥ 1MΩ) • Điện Trở Tiếp Điểm (≤ 50%) • Khóa Liên Động Cơ/Điện • Chuyển Nguồn NFPA 110 (≤ 10s)',
    scope: 'Tủ chuyển nguồn tự động ATS 3 pha / 4 cực 100A - 4000A, Bộ điều khiển Controller, Cơ cấu chuyển mạch Open/Closed Transition',
    standardCode: 'ANSI/NETA ECS-2024 Mục 8.3',
    standardName: 'NETA ECS-2024 Sec 8.3, ATS 7.22.3, NFPA 110, IEEE 446',
    aiEngine: 'ATS Commissioning Specialist AI (NETA Sec 8.3 & NFPA 110)',
    icon: GitBranch,
    colorGradient: 'from-slate-950 via-blue-950 to-slate-900',
    borderColor: 'hover:border-blue-400 focus:border-blue-500',
    tagColor: 'bg-blue-500/20 text-blue-800 border-blue-300',
    badge: 'Mục 8.3 (Tủ ATS)',
    badgeColor: 'bg-blue-500/20 text-blue-300 border-blue-400/40',
    features: [
      'Tiền đóng điện: Đo cách điện mạch điều khiển 1000V DC (≥ 1.0 MΩ Table 100.1), điện trở tiếp điểm động lực (độ lệch ≤ 50%)',
      'Thử nghiệm liên động cơ học và liên động điện ngăn đóng đồng thời hai nguồn (NFPA 110) & lực siết bu-lông Table 100.12',
      'Trong đóng điện: Kiểm tra thứ tự pha & đồng pha hai nguồn, thử nghiệm chuyển nguồn mang tải tự động (≤ 10s NFPA 110 Class 10)',
      'Sau đóng điện: Thử trả nguồn lưới & làm mát máy phát (Engine cool-down), thử mất 1 pha (IEEE 446) & khảo sát nhiệt IR Table 100.18'
    ],
    isAvailable: true
  },
  {
    id: 'wind-turbine',
    category: 'renewables',
    sectionCode: 'Section 8.2',
    title: 'Hệ Thống Tua-bin Điện Gió (Wind Turbine Generation Systems)',
    titleEn: 'Source Specific Systems Commissioning: Wind Turbine Systems',
    shortTitle: 'Tua-bin Gió (Wind)',
    shortTitleEn: 'Wind Turbine',
    subtitle: 'Máy Phát DFIG/PMSG • Cáp Treo Tháp Gió • Cánh Quạt Blade Lightning (IEC 61400-24) • Rung Động ISO 10816 • Hòa Lưới',
    scope: 'Tua-bin gió Onshore / Offshore 2MW - 12MW, Tháp gió, Máy phát DFIG/PMSG, Bộ biến đổi Converter, Máy biến áp Nacelle',
    standardCode: 'ANSI/NETA ECS-2024 Mục 8.2 & IEC 61400',
    standardName: 'NETA ECS-2024 Sec 8.2 & IEC 61400-1, IEC 61400-24, DNVGL-ST-0376',
    aiEngine: 'Wind Turbine Commissioning Specialist AI',
    icon: Wind,
    colorGradient: 'from-slate-950 via-cyan-950 to-slate-900',
    borderColor: 'hover:border-cyan-400 focus:border-cyan-500',
    tagColor: 'bg-cyan-500/20 text-cyan-800 border-cyan-300',
    badge: 'Mục 8.2 (Lộ trình ECS)',
    badgeColor: 'bg-cyan-500/20 text-cyan-300 border-cyan-400/40',
    features: [
      'Tiền đóng điện: Tiếp địa hệ thống chống sét cánh quạt Lightning Receptor IEC 61400-24 (≤ 0.5 Ω) & Cách điện Stator/Rotor',
      'Kiểm tra mô-men siết bu-lông móng tháp, khớp xoay Yaw/Pitch theo tiêu chuẩn nhà sản xuất & NETA Table 100.12',
      'Trong đóng điện: Thử nghiệm đồng tốc khởi động, điều khiển góc pitch, hòa lưới êm dịu & đáp ứng tần số FFR',
      'Sau đóng điện: Phân tích rung động vòng bi Drive-train (ISO 10816-21) & Khảo sát nhiệt hồng ngoại tủ Converter'
    ],
    isAvailable: false
  },
  {
    id: 'bess',
    category: 'backup-storage',
    sectionCode: 'Section 8.3',
    title: 'Hệ Thống Pin Lưu Trữ Năng Lượng (Battery Energy Storage - BESS)',
    titleEn: 'Source Specific Systems Commissioning: Battery Energy Storage Systems (BESS)',
    shortTitle: 'Pin BESS (Storage)',
    shortTitleEn: 'Battery BESS',
    subtitle: 'Module Pin LFP/NMC • Hệ Thống BMS • Cách Điện DC Rack • Ngắt Khẩn Cấp E-Stop • Chống Cháy Nổ NFPA 855',
    scope: 'Hệ thống lưu trữ pin Lithium-ion BESS container, Rack pin, Hệ thống quản lý pin BMS, Biến tần PCS 2 chiều',
    standardCode: 'ANSI/NETA ECS-2024 Mục 8.3 & NFPA 855',
    standardName: 'NETA ECS-2024 Sec 8.3, NFPA 855, UL 9540A, IEEE 1679.1',
    aiEngine: 'BESS Commissioning Specialist AI',
    icon: BatteryCharging,
    colorGradient: 'from-slate-950 via-purple-950 to-slate-900',
    borderColor: 'hover:border-purple-400 focus:border-purple-500',
    tagColor: 'bg-purple-500/20 text-purple-800 border-purple-300',
    badge: 'Mục 8.3 (Lộ trình ECS)',
    badgeColor: 'bg-purple-500/20 text-purple-300 border-purple-400/40',
    features: [
      'Tiền đóng điện: Kiểm tra điện trở nội cell pin, độ đồng đều điện áp OCV giữa các module & cách điện DC Rack (≥ 10 MΩ)',
      'Thử nghiệm liên động BMS báo động nhiệt độ cao, sụt áp, ngắt công tắc tơ DC bảo vệ & khóa liên động chữa cháy NFPA 855',
      'Trong đóng điện: Nạp thử nghiệm bước dòng thấp, kiểm tra cân bằng tế bào (Cell Balancing) & thứ tự pha PCS',
      'Sau đóng điện: Thử nghiệm chu kỳ sạc/xả đầy tải (Round-Trip Efficiency), khảo sát nhiệt IR tủ rack & đo sóng hài PCS'
    ],
    isAvailable: false
  },
  {
    id: 'engine-generator',
    category: 'backup-storage',
    sectionCode: 'Section 8.4',
    title: 'Hệ Thống Tổ Máy Phát Điện Khẩn Cấp & Dự Phòng (Emergency Engine Generators)',
    titleEn: 'Source Specific Systems Commissioning: Engine Generator Systems',
    shortTitle: 'Máy Phát Điện (Genset)',
    shortTitleEn: 'Engine Generator',
    subtitle: 'NFPA 110 Class 10 (≤ 10s) • Cách Điện Stator/Rotor • Tải Giả Load Bank 100% • Tắt Máy Quá Tốc 115% • Hòa Đồng Bộ',
    scope: 'Tổ máy phát điện Diesel khẩn cấp 500kVA - 3500kVA Bệnh viện, Data Center, Tòa nhà cao tầng, Nhà máy công nghiệp',
    standardCode: 'ANSI/NETA ECS-2024 Mục 8.4 & NFPA 110',
    standardName: 'NETA ECS-2024 Sec 8.4, NFPA 110, IEEE 115',
    aiEngine: 'Emergency Genset Commissioning Specialist AI',
    icon: Power,
    colorGradient: 'from-slate-950 via-rose-950 to-slate-900',
    borderColor: 'hover:border-rose-400 focus:border-rose-500',
    tagColor: 'bg-rose-500/20 text-rose-800 border-rose-300',
    badge: 'Mục 8.4 (Lộ trình ECS)',
    badgeColor: 'bg-rose-500/20 text-rose-300 border-rose-400/40',
    features: [
      'Tiền đóng điện: Đo cách điện cuộn dây Stator/Rotor (PI ≥ 2.0 IEEE 43), điện trở một chiều DLRO & kiểm tra dầu/nước làm mát',
      'Thử nghiệm cơ cấu bảo vệ dừng máy khẩn cấp: Quá tốc 115%, áp lực nhớt thấp, nhiệt độ nước cao & dừng khẩn cấp E-Stop',
      'Trong đóng điện: Thử nghiệm khởi động tự động nhận tải trong 10 giây (NFPA 110 Class 10 Type 10) & đồng bộ hòa nguồn',
      'Sau đóng điện: Chạy thử tải bậc thang Load Bank (25%, 50%, 75%, 100% trong 4 giờ) & khảo sát nhiệt hồng ngoại máy phát'
    ],
    isAvailable: false
  },
  {
    id: 'ups-system',
    category: 'backup-storage',
    sectionCode: 'Section 8.2',
    title: 'Hệ Thống Nguồn Cung Cấp Liên Tục (Uninterruptible Power Supply - UPS)',
    titleEn: 'Source Specific Systems Commissioning: Uninterruptible Power Supply (UPS) Procedures',
    shortTitle: 'Nguồn UPS (UPS System)',
    shortTitleEn: 'UPS Systems',
    subtitle: 'Chuyển Mạch Static 0ms • Trở Kháng Ắc Quy IEEE 1188 • Cách Điện 1000V ≥ 100MΩ • Sóng Hài THD-v ≤ 3% • Tải Giả 100%',
    scope: 'Hệ thống UPS tĩnh 3 pha 50kVA - 1600kVA cho Trung tâm dữ liệu (Data Center), Bệnh viện, Phòng điều khiển SCADA, Ngân hàng',
    standardCode: 'ANSI/NETA ECS-2024 Mục 8.2',
    standardName: 'NETA ECS-2024 Sec 8.2, ATS 7.22.2 & 7.18, NFPA 111, IEEE 450/1188, Tables 100.1, 100.12, 100.18',
    aiEngine: 'UPS Commissioning Specialist AI (NETA Sec 8.2 & NFPA 111)',
    icon: Zap,
    colorGradient: 'from-slate-950 via-indigo-950 to-slate-900',
    borderColor: 'hover:border-indigo-400 focus:border-indigo-500',
    tagColor: 'bg-indigo-500/20 text-indigo-800 border-indigo-300',
    badge: 'Mục 8.2 (UPS)',
    badgeColor: 'bg-indigo-500/20 text-indigo-300 border-indigo-400/40',
    features: [
      'Tiền đóng điện: Đo cách điện mạch động lực 1000V DC (≥ 100 MΩ Table 100.1), kiểm tra nội trở tế bào bình ắc-quy (IEEE 1188 độ lệch ≤ 20%) & lực siết Table 100.12',
      'Kiểm tra khóa liên động cơ học và liên động điện giữa Inverter/Static Switch và máy cắt Bypass bảo trì MBB ngăn ngừa sự cố chập pha',
      'Trong đóng điện: Kiểm tra thứ tự pha AC ngõ vào, điện áp nạp DC Float/Equalize, thử chuyển mạch Static Bypass 0 ms (No Break) & thử ngắt AC ngõ vào tức thời',
      'Sau đóng điện: Xả ắc-quy 100% tải giả Load Bank (NFPA 111), đo méo sóng hài THD-v ≤ 3.0% (IEEE 519) & khảo sát nhiệt hồng ngoại IR Table 100.18'
    ],
    isAvailable: true
  },
  {
    id: 'integrated-systems',
    category: 'integrated',
    sectionCode: 'Section 9',
    title: 'Thử Nghiệm Tích Hợp Hệ Thống & Trình Tự Chức Năng (System Functional Sequence)',
    titleEn: 'Verification & System Functional Sequence Testing',
    shortTitle: 'Tích Hợp Hệ Thống',
    shortTitleEn: 'Integrated Systems',
    subtitle: 'Mất Lưới Tổng (Blackout Test) • Sa Thải Phụ Tải • Tự Động Đảo Nguồn ATS • Khóa Liên Động Liên Trạm • SCADA',
    scope: 'Thử nghiệm tích hợp toàn diện hệ thống điện: Trạm biến áp, Máy phát, UPS, Tủ ATS, Hệ thống PCCC & Tự động hóa',
    standardCode: 'ANSI/NETA ECS-2024 Mục 9',
    standardName: 'NETA ECS-2024 Sec 9, NFPA 3, NFPA 4, ASHRAE Guideline 0',
    aiEngine: 'Integrated System Commissioning Specialist AI',
    icon: ClipboardList,
    colorGradient: 'from-slate-950 via-slate-800 to-slate-900',
    borderColor: 'hover:border-slate-400 focus:border-slate-500',
    tagColor: 'bg-slate-500/20 text-slate-800 border-slate-300',
    badge: 'Mục 9 (Lộ trình ECS)',
    badgeColor: 'bg-slate-500/20 text-slate-300 border-slate-400/40',
    features: [
      'Mô phỏng sự cố mất nguồn lưới điện chính (Integrated Blackout Test) và đo đếm thời gian chuyển tải hệ thống',
      'Xác nhận chuỗi khởi động tự động: Máy phát khởi động ≤ 10s, ATS đóng nguồn khẩn cấp, UPS duy trì liên tục',
      'Thử nghiệm liên động hệ thống báo cháy PCCC: Ngắt thông gió, ngắt điện chiếu sáng và ưu tiên nguồn thang máy cứu hỏa',
      'Xác thực đường truyền truyền thông SCADA/BMS thời gian thực đối với tất cả các thông số đo lường và tín hiệu cảnh báo'
    ],
    isAvailable: false
  }
];

interface Props {
  currentMode: EcsModuleId;
  onSelectMode: (mode: EcsModuleId) => void;
  lang?: 'vi' | 'en';
}

export const EcsModuleSelector: React.FC<Props> = ({
  currentMode,
  onSelectMode,
  lang = 'vi'
}) => {
  const isVi = lang === 'vi';
  const [isExpanded, setIsExpanded] = useState(false);
  const [selectedCategory, setSelectedCategory] = useState<EcsCategory>('all');
  const [searchQuery, setSearchQuery] = useState('');

  const currentItem = ECS_MODULE_CATALOG.find(i => i.id === currentMode) || ECS_MODULE_CATALOG[0];

  // Category counts
  const categoryCounts = useMemo(() => {
    return {
      all: ECS_MODULE_CATALOG.length,
      'section-7': ECS_MODULE_CATALOG.filter(i => i.category === 'section-7').length,
      'section-8': ECS_MODULE_CATALOG.filter(i => i.sectionCode.startsWith('Section 8')).length,
      renewables: ECS_MODULE_CATALOG.filter(i => i.category === 'renewables').length,
      'backup-storage': ECS_MODULE_CATALOG.filter(i => i.category === 'backup-storage').length,
      integrated: ECS_MODULE_CATALOG.filter(i => i.category === 'integrated').length,
    };
  }, []);

  // Filtered items based on Category & Search Query
  const filteredItems = useMemo(() => {
    return ECS_MODULE_CATALOG.filter(item => {
      // Category filter
      let matchesCat = true;
      if (selectedCategory === 'section-7') {
        matchesCat = item.category === 'section-7';
      } else if (selectedCategory === 'section-8') {
        matchesCat = item.sectionCode.startsWith('Section 8');
      } else if (selectedCategory === 'renewables') {
        matchesCat = item.category === 'renewables';
      } else if (selectedCategory === 'backup-storage') {
        matchesCat = item.category === 'backup-storage';
      } else if (selectedCategory === 'integrated') {
        matchesCat = item.category === 'integrated';
      }

      // Search filter
      const q = searchQuery.toLowerCase().trim();
      const matchesSearch =
        !q ||
        item.title.toLowerCase().includes(q) ||
        item.titleEn.toLowerCase().includes(q) ||
        item.standardCode.toLowerCase().includes(q) ||
        item.scope.toLowerCase().includes(q) ||
        item.sectionCode.toLowerCase().includes(q) ||
        item.shortTitle.toLowerCase().includes(q);

      return matchesCat && matchesSearch;
    });
  }, [selectedCategory, searchQuery]);

  return (
    <div className="bg-white rounded-2xl border border-slate-200 shadow-md overflow-hidden transition-all">
      {/* 1. TOP HEADER & QUICK TOGGLE STRIP */}
      <div className="bg-slate-950 p-3 sm:p-3.5 border-b border-teal-500/30 flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3">
        <div className="flex items-center gap-3 shrink-0">
          <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-teal-500 via-teal-600 to-emerald-800 text-slate-950 flex items-center justify-center font-black text-xs shrink-0 shadow-sm shadow-teal-500/40">
            ECS
          </div>
          <div>
            <div className="text-xs font-black text-white tracking-wide uppercase flex items-center gap-1.5 flex-wrap">
              <span>{isVi ? 'Đặc Tả Kỹ Thuật Chạy Thử & Đóng Điện' : 'Electrical Commissioning Specifications'}</span>
              <span className="text-[10px] px-1.5 py-0.2 rounded font-mono font-bold bg-teal-500/20 text-teal-300 border border-teal-500/30">
                ANSI/NETA ECS-2024
              </span>
            </div>
            <div className="text-[11px] text-slate-400 mt-0.5">
              {isVi ? 'Hệ thống đối soát & thẩm định AI quy trình chạy thử theo cấp điện áp & nguồn phát' : 'Multi-voltage & source-specific electrical commissioning suite with AI verification'}
            </div>
          </div>
        </div>

        {/* Right side: Active Module Quick Button & Directory Toggle */}
        <div className="flex items-center gap-2 flex-wrap self-end md:self-auto">
          {/* Quick Active Pills for Ready Modules */}
          <div className="hidden sm:flex items-center gap-1.5 bg-slate-900/90 p-1 rounded-xl border border-slate-800">
            {ECS_MODULE_CATALOG.filter(i => i.isAvailable).map((item) => {
              const isSelected = currentMode === item.id;
              const IconComp = item.icon;
              return (
                <button
                  key={item.id}
                  type="button"
                  onClick={() => onSelectMode(item.id)}
                  className={`flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-bold transition-all ${
                    isSelected
                      ? 'bg-gradient-to-r from-teal-500 to-emerald-600 text-slate-950 font-black shadow-sm shadow-teal-500/30 ring-1 ring-teal-400/50'
                      : 'text-slate-300 hover:text-white hover:bg-slate-800'
                  }`}
                >
                  <IconComp size={13} className={isSelected ? 'text-slate-950' : 'text-teal-400'} />
                  <span>{isVi ? item.shortTitle : item.shortTitleEn}</span>
                  {isSelected && <CheckCircle2 size={12} className="text-slate-950" />}
                </button>
              );
            })}
          </div>

          {/* Directory Hub Button (Like ATS "Đổi Thiết Bị Kiểm Định") */}
          <button
            type="button"
            onClick={() => setIsExpanded(!isExpanded)}
            className={`px-3.5 py-2 rounded-xl text-xs font-black flex items-center gap-2 transition-all shadow-md ${
              isExpanded
                ? 'bg-teal-500 text-slate-950 shadow-teal-500/20 ring-2 ring-teal-400'
                : 'bg-slate-800 hover:bg-slate-700 text-white border border-slate-700'
            }`}
          >
            <Layers size={14} className={isExpanded ? 'text-slate-950' : 'text-teal-400'} />
            <span>{isExpanded ? (isVi ? 'Đóng Danh Mục' : 'Close Directory') : (isVi ? 'Đổi Quy Trình ECS (8)' : 'Change ECS Module (8)')}</span>
            <ChevronRight size={14} className={`transform transition-transform ${isExpanded ? 'rotate-90' : ''}`} />
          </button>
        </div>
      </div>

      {/* 2. ACTIVE MODULE INFORMATION STRIP (When collapsed or expanded) */}
      <div className="px-5 py-3 bg-gradient-to-r from-slate-900 via-slate-800 to-slate-900 text-white flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3 border-b border-slate-700/60">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-teal-500/20 text-teal-400 border border-teal-400/30 flex items-center justify-center shrink-0 shadow-sm">
            {React.createElement(currentItem.icon, { size: 20 })}
          </div>
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-sm font-black text-white">
                {isVi ? currentItem.title : currentItem.titleEn}
              </span>
              <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full border ${currentItem.badgeColor}`}>
                {currentItem.standardCode}
              </span>
            </div>
            <div className="text-xs text-slate-300 flex items-center gap-2 flex-wrap mt-0.5">
              <span className="font-semibold text-slate-400">{currentItem.standardName}</span>
              <span className="text-slate-500">•</span>
              <span className="text-teal-300/90 text-[11px]">{currentItem.subtitle}</span>
            </div>
          </div>
        </div>

        <div className="flex items-center gap-2 shrink-0 self-end sm:self-auto">
          <span className="px-2.5 py-1 rounded-full text-[10px] font-bold bg-teal-500/20 text-teal-300 border border-teal-400/30 flex items-center gap-1">
            <CheckCircle2 size={12} className="text-teal-400" />
            <span>{isVi ? 'Đang chọn áp dụng' : 'Active Module'}</span>
          </span>
        </div>
      </div>

      {/* 3. EXPANDED DIRECTORY & SELECTION HUB (DESIGNED ACCORDING TO ATS STANDARD) */}
      {isExpanded && (
        <div className="p-4 sm:p-6 bg-slate-50 border-t border-slate-200 space-y-4 animate-in slide-in-from-top-2 duration-200">
          {/* Top Bar: Category Filter Pills + Search Box (EXACTLY LIKE ATS IN SCREENSHOT) */}
          <div className="flex flex-col lg:flex-row items-stretch lg:items-center justify-between gap-3">
            {/* Category Pills */}
            <div className="flex flex-wrap items-center gap-1.5 text-xs font-semibold">
              <button
                type="button"
                onClick={() => setSelectedCategory('all')}
                className={`px-3 py-1.5 rounded-lg transition-all ${
                  selectedCategory === 'all'
                    ? 'bg-blue-600 text-white shadow-xs font-bold'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                {isVi ? 'Tất cả quy trình' : 'All Procedures'} ({categoryCounts.all})
              </button>

              <button
                type="button"
                onClick={() => setSelectedCategory('section-7')}
                className={`px-3 py-1.5 rounded-lg transition-all flex items-center gap-1.5 ${
                  selectedCategory === 'section-7'
                    ? 'bg-teal-700 text-white shadow-xs font-bold'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <Zap size={13} />
                <span>{isVi ? 'Hệ thống chung Mục 7' : 'General Systems Sec 7'} ({categoryCounts['section-7']})</span>
              </button>

              <button
                type="button"
                onClick={() => setSelectedCategory('renewables')}
                className={`px-3 py-1.5 rounded-lg transition-all flex items-center gap-1.5 ${
                  selectedCategory === 'renewables'
                    ? 'bg-amber-600 text-white shadow-xs font-bold'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <Sun size={13} />
                <span>{isVi ? 'Điện mặt trời & Gió' : 'Solar & Wind'} ({categoryCounts.renewables})</span>
              </button>

              <button
                type="button"
                onClick={() => setSelectedCategory('backup-storage')}
                className={`px-3 py-1.5 rounded-lg transition-all flex items-center gap-1.5 ${
                  selectedCategory === 'backup-storage'
                    ? 'bg-purple-700 text-white shadow-xs font-bold'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <BatteryCharging size={13} />
                <span>{isVi ? 'Lưu trữ BESS, Genset & UPS' : 'BESS, Genset & UPS'} ({categoryCounts['backup-storage']})</span>
              </button>

              <button
                type="button"
                onClick={() => setSelectedCategory('integrated')}
                className={`px-3 py-1.5 rounded-lg transition-all flex items-center gap-1.5 ${
                  selectedCategory === 'integrated'
                    ? 'bg-slate-800 text-white shadow-xs font-bold'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <ClipboardList size={13} />
                <span>{isVi ? 'Tích hợp hệ thống Mục 9' : 'Integrated Testing Sec 9'} ({categoryCounts.integrated})</span>
              </button>
            </div>

            {/* Search Box on the right (matching screenshot placeholder) */}
            <div className="relative max-w-xs w-full">
              <Search size={14} className="absolute left-3 top-2.5 text-slate-400" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder={isVi ? 'Tìm checklist theo tên hoặc NETA code...' : 'Search checklist by name or NETA code...'}
                className="w-full pl-8 pr-3 py-1.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-teal-500 font-medium text-slate-800"
              />
            </div>
          </div>

          {/* Cards Grid (Identical Layout to ATS Checklists) */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3.5">
            {filteredItems.map((item) => {
              const ItemIcon = item.icon;
              const isSelected = item.id === currentMode;

              return (
                <div
                  key={item.id}
                  onClick={() => {
                    if (item.isAvailable) {
                      onSelectMode(item.id);
                      setIsExpanded(false);
                    }
                  }}
                  className={`p-4 rounded-xl border-2 transition-all flex flex-col justify-between ${
                    isSelected
                      ? 'bg-white border-teal-600 shadow-md ring-2 ring-teal-500/20 cursor-pointer'
                      : item.isAvailable
                      ? 'bg-white border-slate-200 hover:border-teal-300 hover:shadow-xs cursor-pointer'
                      : 'bg-slate-100/70 border-slate-200 opacity-80 cursor-not-allowed'
                  }`}
                >
                  <div>
                    {/* Card Header: Icon, Tags, Title & Action */}
                    <div className="flex items-start justify-between gap-3">
                      <div className="flex items-center gap-2.5">
                        <div className={`w-9 h-9 rounded-lg bg-gradient-to-br ${item.colorGradient} flex items-center justify-center text-white shrink-0 shadow-xs`}>
                          <ItemIcon size={18} />
                        </div>
                        <div>
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full border ${item.tagColor}`}>
                              {item.standardCode}
                            </span>
                            <span className="text-[10px] font-mono text-slate-500 font-semibold">
                              {item.sectionCode}
                            </span>
                          </div>
                          <h3 className="font-bold text-sm text-slate-900 mt-1">
                            {isVi ? item.title : item.titleEn}
                          </h3>
                        </div>
                      </div>

                      {/* State Button */}
                      {isSelected ? (
                        <span className="px-2.5 py-1 rounded-full text-[10px] font-bold bg-teal-600 text-white flex items-center gap-1 shadow-xs shrink-0">
                          <CheckCircle2 size={12} />
                          <span>{isVi ? 'Đang chọn' : 'Selected'}</span>
                        </span>
                      ) : item.isAvailable ? (
                        <span className="px-2.5 py-1 rounded-full text-[10px] font-semibold bg-slate-100 text-slate-700 hover:bg-teal-50 hover:text-teal-800 flex items-center gap-1 shrink-0 border border-slate-200 transition-colors">
                          <span>{isVi ? 'Chọn mẫu' : 'Select'}</span>
                          <ArrowRight size={11} />
                        </span>
                      ) : (
                        <span className="px-2 py-0.5 rounded-full text-[9px] font-bold bg-slate-200 text-slate-600 flex items-center gap-1 shrink-0 border border-slate-300">
                          <Clock size={10} />
                          <span>{isVi ? 'Lộ trình ECS' : 'Roadmap'}</span>
                        </span>
                      )}
                    </div>

                    {/* Subtitle & Scope */}
                    <div className="mt-2.5 space-y-1">
                      <div className="text-[11px] font-semibold text-teal-800 bg-teal-50/70 px-2 py-1 rounded-md border border-teal-100/60">
                        {item.subtitle}
                      </div>
                      <p className="text-xs text-slate-600 line-clamp-2">
                        <span className="font-semibold text-slate-700">{isVi ? 'Phạm vi: ' : 'Scope: '}</span>
                        {item.scope}
                      </p>
                    </div>

                    {/* 4 Feature Bullet Points (Like ATS Catalog) */}
                    <div className="mt-3 space-y-1.5 pt-2.5 border-t border-slate-100">
                      {item.features.map((feat, idx) => (
                        <div key={idx} className="flex items-start gap-1.5 text-[11px] text-slate-600 leading-tight">
                          <div className="w-1.5 h-1.5 rounded-full bg-teal-500 mt-1 shrink-0" />
                          <span>{feat}</span>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Card Footer: Standard reference & AI Engine */}
                  <div className="mt-3.5 pt-2 border-t border-slate-100 flex items-center justify-between text-[10px] text-slate-500">
                    <span className="font-semibold text-slate-600">{item.standardName}</span>
                    <span className="flex items-center gap-1 text-teal-700 font-bold">
                      <Sparkles size={11} />
                      <span>{item.aiEngine}</span>
                    </span>
                  </div>
                </div>
              );
            })}
          </div>

          {filteredItems.length === 0 && (
            <div className="text-center py-8 bg-white rounded-xl border border-slate-200 text-slate-500">
              <Search size={24} className="mx-auto mb-2 text-slate-400" />
              <p className="text-sm font-semibold">{isVi ? 'Không tìm thấy quy trình ECS phù hợp' : 'No matching ECS procedures found'}</p>
              <p className="text-xs text-slate-400 mt-0.5">{isVi ? 'Vui lòng thử lại với từ khóa khác hoặc xóa bộ lọc.' : 'Please try another search keyword or reset filters.'}</p>
            </div>
          )}
        </div>
      )}
    </div>
  );
};
