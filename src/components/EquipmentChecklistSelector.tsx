import React, { useState } from 'react';
import {
  Sun,
  BatteryCharging,
  Zap,
  PlugZap,
  Cable,
  Wind,
  Thermometer,
  Radio,
  RotateCcw,
  Droplets,
  ClipboardList,
  Search,
  Filter,
  CheckCircle2,
  Sparkles,
  ChevronRight,
  ShieldCheck,
  FolderArchive,
  ArrowRight,
  Layers,
  Flame,
  Gauge,
  Camera,
  Cpu,
  GitBranch,
  Power
} from 'lucide-react';
import { NameplateScannerModal } from './NameplateScannerModal';
import { ExtractedNameplateData } from '../../server/nameplateOcrAnalyzer';

export type FieldEntryMode = 'neta-dc-motor' | 'neta-grounding' | 'neta-lv-breaker' | 'neta-cable-lv' | 'neta-switchgear' | 'neta-generator' | 'neta-ats' | 'neta-battery-vrla' | 'neta-battery-flooded' | 'neta-sync-machinery' | 'neta-ups' | 'neta-relay' | 'neta-liquid-transformer' | 'neta-motor' | 'neta-pd' | 'neta-thermography' | 'neta-sf6-switch' | 'neta-cable' | 'neta-evse' | 'neta-pv' | 'neta-bess' | 'neta-dry-large' | 'neta-dry-type' | 'standard';

export interface EquipmentChecklistItem {
  id: FieldEntryMode;
  category: 'grounding' | 'batteries' | 'ups' | 'relays' | 'motors' | 'pd' | 'thermography' | 'switches' | 'cables' | 'evse' | 'pv' | 'transformers' | 'bess' | 'standard';
  title: string;
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
}

export const CHECKLIST_CATALOG: EquipmentChecklistItem[] = [
  {
    id: 'neta-grounding',
    category: 'grounding',
    title: 'Hệ Thống Nối Đất & Tiếp Địa Trạm (Grounding Systems)',
    subtitle: 'Fall-of-Potential IEEE 81 (≤ 1.0Ω / 5.0Ω) • Liên Kết Điểm-Điểm (< 0.5Ω) • Hàng Rào Trạm • Mối Hàn Cadweld',
    scope: 'Lưới tiếp địa trạm biến áp 110kV/22kV, bãi cọc tiếp địa nhà máy công nghiệp, hệ thống chống sét & tiếp địa đẳng thế',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.13 & IEEE Std 81',
    standardName: 'NETA ATS-2025 Sec 7.13 & IEEE Std 81, IEEE Std 80',
    aiEngine: 'Grounding Systems Specialist AI (NETA Sec 7.13 & IEEE 81)',
    icon: ShieldCheck,
    colorGradient: 'from-slate-950 via-teal-950 to-slate-900',
    borderColor: 'hover:border-teal-400 focus:border-teal-500',
    tagColor: 'bg-teal-500/20 text-teal-800 border-teal-300',
    features: [
      'Điện Trở Lưới Đất Fall-of-Potential: Bắt Buộc ≤ 5.0 Ω (Trạm Biến Áp BẮT BUỘC ≤ 1.0 Ω)',
      'Điện Trở Liên Kết Điểm-đến-Điểm (Point-to-Point) Bắt Buộc < 0.5 Ω (Mục 7.13.D.1)',
      'Tiếp Địa Hàng Rào & Cổng Trạm (Fence & Gate Bonding) Bắt Buộc ≤ 0.5 Ω Chống Điện Áp Chạm',
      'Kiểm Tra Mối Hàn Hóa Nhiệt Cadweld, Tiết Diện Cáp Đồng 240mm² & Giếng Kiểm Tra'
    ]
  },
  {
    id: 'neta-switchgear',
    category: 'switches',
    title: 'Tủ Điện Phân Phối & Tủ Đóng Cắt (Switchgear & Switchboard Assemblies)',
    subtitle: 'Khóa Liên Động • Bu-lông DLRO (≤ 50%) • Cách Điện Mạch Điều Khiển (≥ 2.0 MΩ) • MBA CPT (≤ 0.5%) • TEV PD',
    scope: 'Tủ phân phối trung thế Metal-Clad 24kV, tủ MSB hạ thế, tủ đóng cắt phân đoạn, trạm biến áp nhà máy & tòa nhà',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.1.1',
    standardName: 'NETA ATS-2025 Sec 7.1.1 & Table 100.1, 100.12, 100.23',
    aiEngine: 'Switchgear & Switchboard Specialist AI (Section 7.1.1)',
    icon: Power,
    colorGradient: 'from-slate-950 via-indigo-950 to-slate-900',
    borderColor: 'hover:border-indigo-400 focus:border-indigo-500',
    tagColor: 'bg-indigo-500/20 text-indigo-800 border-indigo-300',
    features: [
      'Cách Điện Mạch Điều Khiển Nhị Thứ BẮT BUỘC ≥ 2.0 MΩ (NETA 7.1.1.D.4)',
      'Độ Lệch Điện Trở Mối Nối Bu-lông DLRO Bắt Buộc ≤ 50% Min (Mục 7.1.1.D.1)',
      'Cách Điện Thanh Cái Busbar (Table 100.1) & Sai Số Tỷ Số CPT (≤ 0.5%)',
      'Khóa Liên Động Cơ Điện (Key Exchange, Cửa Sập) & Xung TEV PD Online'
    ]
  },
  {
    id: 'neta-lv-breaker',
    category: 'switches',
    title: 'Máy Cắt Không Khí Hạ Áp (Low-Voltage Power Circuit Breakers - ACB)',
    subtitle: 'Tiếp Điểm Cực (≤ 50% Sec 7.6.1.2.D.2) • Bơm Dòng Trip Unit L, S, I, G • Cách Điện Table 100.1 • Khóa Liên Động',
    scope: 'Máy cắt tổng hạ thế ACB 630A-6300A tủ MSB, máy cắt liên lạc Bus Tie, máy cắt lộ ra tủ phân phối động lực',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.6.1.2',
    standardName: 'NETA ATS-2025 Sec 7.6.1.2 & Table 100.1, 100.12',
    aiEngine: 'LV Power Circuit Breaker Specialist AI (NETA Sec 7.6.1.2)',
    icon: Power,
    colorGradient: 'from-slate-950 via-blue-950 to-indigo-950',
    borderColor: 'hover:border-blue-400 focus:border-blue-500',
    tagColor: 'bg-blue-500/20 text-blue-800 border-blue-300',
    features: [
      'Độ Lệch Điện Trở Tiếp Xúc Cực Bắt Buộc ≤ 50% Min (Mục 7.6.1.2.D.2)',
      'Bơm Dòng Thứ Cấp Thử Nghiệm Trip Unit L, S, I, G Đúng Đường Cong Đặc Tính',
      'Cách Điện Máy Cắt Table 100.1 (≥ 100 MΩ) & Mạch Phụ Bắt Buộc ≥ 2.0 MΩ',
      'Cơ Cấu Kéo/Đẩy Drawout, Khóa Liên Động Vị Trí Cell & Ngàm Cắm Finger Clusters'
    ]
  },
  {
    id: 'neta-generator',
    category: 'motors',
    title: 'Tổ Máy Phát Điện Khẩn Cấp / Dự Phòng (Emergency Engine Generators)',
    subtitle: 'Bảo Vệ Tắt Máy (Quá Tốc 115%) • Cách Điện Stator IEEE 43 (PI ≥ 2.0) • Nhận Tải NFPA 110 Class 10 (≤ 10s)',
    scope: 'Tổ máy phát điện Diesel khẩn cấp 200kW-3000kW Bệnh viện, Data Center, Tòa nhà cao tầng, Nhà máy công nghiệp',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.22.1 & NFPA 110',
    standardName: 'NETA ATS-2025 Sec 7.22.1 & NFPA 110',
    aiEngine: 'Engine Generator Specialist AI (Section 7.22.1 & NFPA 110)',
    icon: Power,
    colorGradient: 'from-red-950 via-slate-900 to-amber-950',
    borderColor: 'hover:border-amber-400 focus:border-amber-500',
    tagColor: 'bg-amber-500/20 text-amber-800 border-amber-300',
    features: [
      'Bảo Vệ Tắt Máy Quá Tốc Độ (Overspeed Trip Bắt Buộc Cắt Tức Thì 115%)',
      'Chỉ Số Phân Cực Stator PI ≥ 2.0 (IEEE 43 & Table 100.11)',
      'Thời Gian Khởi Động & Cấp Tải Khẩn Cấp NFPA 110 Class 10 (≤ 10 Giây)',
      'Thử Nghiệm Tải Giả 100% Load Bank & Độ Ổn Định Điện Áp AVR (≤ ±1%)'
    ]
  },
  {
    id: 'neta-ats',
    category: 'switches',
    title: 'Bộ Chuyển Nguồn Tự Động (Automatic Transfer Switches - ATS)',
    subtitle: 'Khóa Liên Động Cơ Khí • Điện Trở Tiếp Điểm (≤ 50%) • Cách Điện Điều Khiển (≥ 2 MΩ) • Trình Tự Tự Động',
    scope: 'Tủ ATS hạ thế 400V/800A-3200A tòa nhà, bệnh viện, data center, trạm bơm, hệ thống nguồn khẩn cấp máy phát',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.22.3',
    standardName: 'NETA ATS-2025 Sec 7.22.3',
    aiEngine: 'Automatic Transfer Switch Specialist AI (Section 7.22.3)',
    icon: GitBranch,
    colorGradient: 'from-orange-950 via-slate-900 to-amber-950',
    borderColor: 'hover:border-orange-400 focus:border-orange-500',
    tagColor: 'bg-orange-500/20 text-orange-800 border-orange-300',
    features: [
      'Khóa Liên Động Cơ Khí (Mechanical Interlock) Chống Đóng Trùng',
      'Độ Lệch Điện Trở Tiếp Điểm Cực (Bắt Buộc ≤ 50% Sec 7.22.3.D.4)',
      'Cách Điện Mạch Điều Khiển Thử 500V/1000V (Bắt Buộc ≥ 2.0 MΩ)',
      'Trình Tự Chuyển Mạch Tự Động & Bộ Định Thời (Engine Start/Delays)'
    ]
  },
  {
    id: 'neta-battery-vrla',
    category: 'batteries',
    title: 'Ắc Quy Chì-Axit Kín Khí Van Điều Áp (VRLA AGM/GEL)',
    subtitle: 'Nhiệt Độ Cực Âm & Trôi Nhiệt • Nội Trở Ôm (≤ 25%) • Cầu Nối Intercell (≤ 50%) • Cân Bằng Đất DC',
    scope: 'Giàn pin kín khí 220V/110V/48V DC trạm 110kV/220kV, phòng UPS Data Center, viễn thông và công nghiệp',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.18.1.3',
    standardName: 'NETA ATS-2025 Sec 7.18.1.3 & IEEE 1188',
    aiEngine: 'VRLA Battery Specialist AI (Section 7.18.1.3 & IEEE 1188)',
    icon: BatteryCharging,
    colorGradient: 'from-teal-950 via-slate-900 to-emerald-950',
    borderColor: 'hover:border-teal-400 focus:border-teal-500',
    tagColor: 'bg-teal-500/20 text-teal-800 border-teal-300',
    features: [
      'Giám Sát Nhiệt Độ Cực Âm & Cảnh Báo Trôi Nhiệt (Thermal Runaway)',
      'Nội Trở Ôm Monoblock Biến Động (≤ 25% per NETA 7.18.1.3.D.7)',
      'Độ Lệch Điện Trở Cầu Nối Intercell (≤ 50% Min tương tự)',
      'Kiểm Tra Độ Cân Bằng Điện Áp DC Đối Đất (+/-) & Khảo Sát Nhiệt'
    ]
  },
  {
    id: 'neta-battery-flooded',
    category: 'batteries',
    title: 'Hệ Thống Điện Một Chiều DC & Ắc Quy Chì-Axit Ngập Dung Dịch (Flooded Lead-Acid)',
    subtitle: 'An Toàn Hydro • Tỷ Trọng SG • Lệch Áp Float (≤ 0.05V) • Nội Trở Ôm (≤ 25%) • Chạm Đất DC (+/-)',
    scope: 'Giàn pin 110V/220V/125V/48V DC trạm biến áp 110kV-500kV, nhà máy điện, tủ nạp ắc quy ngập dung dịch',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.18.1.1',
    standardName: 'NETA ATS-2025 Sec 7.18.1.1 & IEEE 450',
    aiEngine: 'Flooded Lead-Acid Battery Specialist AI (Section 7.18.1.1)',
    icon: BatteryCharging,
    colorGradient: 'from-amber-950 via-slate-900 to-amber-900',
    borderColor: 'hover:border-amber-400 focus:border-amber-500',
    tagColor: 'bg-amber-500/20 text-amber-800 border-amber-300',
    features: [
      'Thông Gió Chống Nổ Hydro H₂ & Trạm Rửa Mắt (Mục 7.18.1.1.A)',
      'Lệch Điện Trở Cầu Nối Intercell (≤ 50% Min tương tự)',
      'Độ Lệch Điện Áp Cell Chế Độ Float (≤ 0.05 Volt)',
      'Biến Động Nội Trở Ôm Cell (≤ 25%) & Phát Hiện Chạm Đất DC'
    ]
  },
  {
    id: 'neta-sync-machinery',
    category: 'motors',
    title: 'Động Cơ & Máy Phát Điện Đồng Bộ (Synchronous Motors & Generators)',
    subtitle: 'Sụt Áp Cực Rô-to (≤10%) • Ngắn Mạch Vòng Dây UMP • Stator IR/PI Table 100.11 • Điốt Quay Kích Từ',
    scope: 'Máy phát điện đồng bộ (Hydro/Thermal/Gas Turbine/Diesel 11kV-22kV), Động cơ đồng bộ công suất lớn',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.15.2',
    standardName: 'NETA ATS-2025 Sec 7.15.2 & IEEE 115',
    aiEngine: 'Synchronous Machinery Specialist AI (Section 7.15.2)',
    icon: RotateCcw,
    colorGradient: 'from-slate-950 via-indigo-950 to-blue-950',
    borderColor: 'hover:border-indigo-400 focus:border-indigo-500',
    tagColor: 'bg-indigo-500/20 text-indigo-800 border-indigo-300',
    features: [
      'Thử Sụt Áp AC Các Cực Rô-to (Pole-to-Pole Drop ≤ 10%)',
      'Phát Hiện Ngắn Mạch Vòng Dây & Lực Từ Bất Đối Xứng UMP',
      'Cách Điện Stator & Chỉ Số Phân Cực PI ≥ 2.0 (Table 100.11)',
      'Kiểm Tra Cầu Điốt Quay Kích Từ & Độ Rung (Table 100.10)'
    ]
  },
  {
    id: 'neta-ups',
    category: 'ups',
    title: 'Hệ Thống Nguồn Cấp Điện Liên Tục (Emergency Systems, UPS)',
    subtitle: 'Công Tắc Tĩnh (< 2ms) • Khóa Liên Động Bypass • Tần Số Dao Động Inverter • Dàn Ắc Quy 7.18',
    scope: 'UPS Online Double Conversion 3-Phase Data Center, Bệnh viện, Công nghiệp, Tủ phân phối nguồn khẩn cấp',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.22.2',
    standardName: 'NETA ATS-2025 Sec 7.22.2 & IEEE 1188',
    aiEngine: 'Uninterruptible Power Systems Specialist AI (Section 7.22.2)',
    icon: BatteryCharging,
    colorGradient: 'from-emerald-950 via-teal-900 to-slate-900',
    borderColor: 'hover:border-emerald-400 focus:border-emerald-500',
    tagColor: 'bg-emerald-500/20 text-emerald-800 border-emerald-300',
    features: [
      'Chuyển Mạch Tĩnh Static Switch Vô Cấp (< 2ms)',
      'Khóa Liên Động Cơ Điện Maintenance Bypass Chống Chập',
      'Độ Lệch Điện Trở Tiếp Xúc Mối Nối Bu-lông (≤ 50%)',
      'Đo Nội Trở Từng Bình Ắc Quy (IEEE 1188 / Section 7.18)'
    ]
  },
  {
    id: 'neta-relay',
    category: 'relays',
    title: 'Rơ-le Bảo Vệ Vi Xử Lý (Protective Relays, Microprocessor-Based)',
    subtitle: 'Đối Soát Settings 100% • ANSI 50/51/87/21 • Đo Lường Tương Tự • Xóa Sạch Event Logs',
    scope: 'Rơ-le số SEL, ABB, Siemens, Schneider MiCOM, GE Multilin... bảo vệ đường dây, MBA, thanh cái, động cơ',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.9.2',
    standardName: 'NETA ATS-2025 Sec 7.9.2 & IEEE C37.90',
    aiEngine: 'Protective Relays Specialist AI (Section 7.9.2)',
    icon: Cpu,
    colorGradient: 'from-amber-950 via-slate-900 to-indigo-950',
    borderColor: 'hover:border-amber-400 focus:border-amber-500',
    tagColor: 'bg-amber-500/20 text-amber-800 border-amber-300',
    features: [
      'Đối Soát File Cài Đặt Trùng Khớp 100% Phiếu Chỉnh Định',
      'Thử Nghiệm Ngưỡng & Thời Gian ANSI 50, 51, 87, 21',
      'Độ Chính Xác Đo Lường Tương Tự (Sai Số ≤ ±0.5%)',
      'BẮT BUỘC Xóa Sạch Nhật Ký Sự Cố Giả (NETA 7.9.2.B.8)'
    ]
  },
  {
    id: 'neta-dc-motor',
    category: 'motors',
    title: 'Động Cơ & Máy Phát Điện Một Chiều DC (DC Motors & Generators)',
    subtitle: 'Cổ Góp & Chổi Than • Sụt Áp Cực Từ (≤ 10%) • Lệch Trở Phiến Bar-to-Bar (≤ 5%) • PI (≥ 2.0) / DAR (≥ 1.4)',
    scope: 'Động cơ cán thép DC, máy phát điện DC, động cơ kéo tàu/xe lửa, máy nâng tời mỏ, động cơ DC kích từ độc lập & Shunt/Series',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.15.3',
    standardName: 'NETA ATS-2025 Sec 7.15.3 & Tables 100.10, 100.11, 100.12',
    aiEngine: 'DC Rotating Machinery Specialist AI (Section 7.15.3)',
    icon: RotateCcw,
    colorGradient: 'from-slate-950 via-rose-950 to-slate-900',
    borderColor: 'hover:border-rose-400 focus:border-rose-500',
    tagColor: 'bg-rose-500/20 text-rose-800 border-rose-300',
    features: [
      'Điện Trở Phiến Cổ Góp (Bar-to-Bar Resistance): Sai Lệch KHÔNG ĐƯỢC Vượt Quá 5.0%',
      'Thử Sụt Áp AC Từng Cực Từ Kích Từ (Pole-to-Pole Drop): Sai Lệch BẮT BUỘC ≤ 10.0%',
      'Cách Điện IEEE 43 & Table 100.11: Máy > 150 kW Bắt Buộc PI ≥ 2.0 (≤ 150 kW DAR ≥ 1.4)',
      'Kiểm Tra Bề Mặt Cổ Góp, Độ Chặt Phiến & Độ Rung Không Tải (Table 100.10)'
    ]
  },
  {
    id: 'neta-motor',
    category: 'motors',
    title: 'Động Cơ & Máy Phát Điện Cảm Ứng AC (Rotating Machinery)',
    subtitle: 'Đo Phổ Dòng MCSA • Rung Động (Table 100.10) • Lệch Pha Stator (≤ 5%) • PI (≥ 2.0)',
    scope: 'Động cơ hạ áp/trung áp, máy phát điện xoay chiều, quạt thông gió, máy bơm công nghiệp',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.15.1',
    standardName: 'NETA ATS-2025 Sec 7.15.1 & Tables 100.10, 100.11',
    aiEngine: 'Rotating Machinery Specialist AI (Section 7.15.1)',
    icon: RotateCcw,
    colorGradient: 'from-blue-950 via-slate-900 to-indigo-950',
    borderColor: 'hover:border-blue-400 focus:border-blue-500',
    tagColor: 'bg-blue-500/20 text-blue-800 border-blue-300',
    features: [
      'Phân Tích Phổ Dòng MCSA (< 40 dB: Gãy Thanh Rotor)',
      'Đo Rung Động Vận Tốc Peak (NETA Table 100.10)',
      'Độ Lệch Điện Trở Stator Giữa Các Pha (≤ 5%)',
      'Chỉ Số Phân Cực PI ≥ 2.0 & Thử Chịu Áp IEEE 95'
    ]
  },
  {
    id: 'neta-pd',
    category: 'pd',
    title: 'Khảo Sát Phóng Điện Cục Bộ Online (Online Partial Discharge)',
    subtitle: 'Cảm Biến TEV, HFCT, UHF, Siêu Âm • Lọc Nhiễu Góc Pha (Table 100.23)',
    scope: 'Tủ điện trung thế Metal-Clad, máy cắt MV, đầu cáp, thanh cái, máy biến áp đang vận hành',
    standardCode: 'ANSI/NETA ATS-2025 Mục 11 & Bảng 100.23',
    standardName: 'NETA ATS-2025 Sec 11 & Table 100.23',
    aiEngine: 'Partial Discharge Survey Specialist AI (Table 100.23)',
    icon: Radio,
    colorGradient: 'from-violet-950 via-indigo-900 to-slate-900',
    borderColor: 'hover:border-violet-400 focus:border-violet-500',
    tagColor: 'bg-violet-500/20 text-violet-800 border-violet-300',
    features: [
      'Đồng Bộ Góc Pha PRPD & Lọc Nhiễu Ngoài (Ghi Chú 1)',
      'Phân Cấp TEV Table 100.23.1 (> 30 dB: Locate and Repair ASAP)',
      'Phân Cấp HFCT Table 100.23.2 (> 500 pC: Locate and Repair ASAP)',
      'Phân Cấp Airborne/Contact Acoustic (Table 100.23.5)'
    ]
  },
  {
    id: 'neta-thermography',
    category: 'thermography',
    title: 'Khảo Sát Nhiệt Độ Hồng Ngoại (Thermographic Survey)',
    subtitle: 'Ảnh Nhiệt Camera • Phân Cấp 4 Mức Mối Nguy • Thanh Cái, Đầu Cos Cáp, Cực Aptomat (Table 100.18)',
    scope: 'Tủ phân phối tổng MSB, thanh cái busbar, mối nối bu-lông, đầu cốt cáp động lực, thiết bị đóng cắt mang tải',
    standardCode: 'ANSI/NETA ATS-2025 Mục 9 & Bảng 100.18',
    standardName: 'NETA ATS-2025 Sec 9 & Table 100.18',
    aiEngine: 'Thermographic Survey Specialist AI (Table 100.18)',
    icon: Thermometer,
    colorGradient: 'from-red-900 via-orange-800 to-slate-900',
    borderColor: 'hover:border-red-400 focus:border-red-500',
    tagColor: 'bg-red-500/20 text-red-800 border-red-300',
    features: [
      'Chênh Lệch Nhiệt Độ Cùng Tải ΔT_comp (Ngưỡng Mức 4: > 15°C)',
      'Chênh Lệch Nhiệt Độ Môi Trường ΔT_amb (Ngưỡng Mức 4: > 40°C)',
      'Phân Cấp Khuyến Nghị 4 Mức (Table 100.18: Repair Immediately)',
      'Kiểm Tra Điều Kiện Tải Vận Hành Bình Thường (Mục 9.1 - 9.3)'
    ]
  },
  {
    id: 'neta-sf6-switch',
    category: 'switches',
    title: 'Cầu Dao Ngắt Mạch Trung Áp Khí SF6 (Switches, SF6, MV)',
    subtitle: 'Tủ RMU • Dao Cắt Phụ Tải • Chất Lượng Khí SF6 (Table 100.13) • Tiếp Điểm Cực • Khóa Liên Động',
    scope: 'Tủ trung thế Ring Main Unit (RMU), phân đoạn lộ đường dây 24kV, trạm biến áp phân phối',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.5.4',
    standardName: 'NETA ATS-2025 Sec 7.5.4',
    aiEngine: 'SF6 Switch Specialist AI Engine (NETA Sec 7.5.4)',
    icon: Wind,
    colorGradient: 'from-teal-800 via-emerald-700 to-slate-900',
    borderColor: 'hover:border-teal-400 focus:border-teal-500',
    tagColor: 'bg-teal-500/20 text-teal-800 border-teal-300',
    features: [
      'Chất Lượng Khí SF6 (Độ tinh khiết > 98.5% Table 100.13)',
      'Nồng Độ Khí Phân Hủy SO2 (< 5 ppmv - Báo Hiệu Hồ Quang)',
      'Độ Lệch Tiếp Điểm Cực (Bắt Buộc ≤ 50% Sec 7.5.4.D.2)',
      'Kiểm Tra Rơ-le Mật Độ / Áp Suất & Khóa Liên Động RMU'
    ]
  },
  {
    id: 'neta-cable-lv',
    category: 'cables',
    title: 'Cáp Điện Hạ Áp (Cables, Low-Voltage, 1,000V Max)',
    subtitle: 'Cách Điện Table 100.1 (≥ 100 MΩ) • Cân Bằng Dây Song Song • Thông Mạch 100% • Lực Siết Table 100.12',
    scope: 'Tuyến cáp hạ thế 0.6/1kV tổng lộ MSB, cáp phân phối động lực nhà xưởng, cáp trục trung tâm thương mại & data center',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.3.2',
    standardName: 'NETA ATS-2025 Sec 7.3.2 & Table 100.1, 100.12',
    aiEngine: 'Low-Voltage Cable Specialist AI (NETA Sec 7.3.2)',
    icon: Cable,
    colorGradient: 'from-emerald-950 via-teal-900 to-slate-900',
    borderColor: 'hover:border-emerald-400 focus:border-emerald-500',
    tagColor: 'bg-emerald-500/20 text-emerald-800 border-emerald-300',
    features: [
      'Cách Điện Table 100.1: Cáp 600V/1000V BẮT BUỘC ≥ 100 MΩ (Cáp 300V ≥ 25 MΩ)',
      'Độ Đều Điện Trở Dây Song Song (Uniform Resistance - Báo Động INVESTIGATE)',
      'Thử Nghiệm Thông Mạch Liên Tục 100% Giữa Hai Đầu Tuyến (Sec 7.3.2.B.2)',
      'Lực Siết Bu-lông Điện Table 100.12 & Khảo Sát Nhiệt Table 100.18'
    ]
  },
  {
    id: 'neta-cable',
    category: 'cables',
    title: 'Cáp Trung & Cao Áp Màn Chắn Kim Loại (Shielded MV/HV Cables)',
    subtitle: 'Cáp Ngầm • Đầu Cáp & Hộp Nối • Thông Mạch Màn Chắn (≤10Ω/1k ft) • VLF 0.1Hz • Tan-Delta',
    scope: 'Tuyến cáp xuất tuyến trạm 110kV/22kV, cáp lực ngầm nhà máy, cáp lộ trung thế',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.3.3',
    standardName: 'NETA ATS-2025 Sec 7.3.3',
    aiEngine: 'Shielded Cable Specialist AI Engine (NETA Sec 7.3.3)',
    icon: Cable,
    colorGradient: 'from-purple-700 via-violet-800 to-indigo-900',
    borderColor: 'hover:border-purple-400 focus:border-purple-500',
    tagColor: 'bg-purple-500/20 text-purple-800 border-purple-300',
    features: [
      'Thông Mạch Màn Chắn (Bắt Buộc ≤ 10 Ω / 1,000 ft)',
      'Thử Chịu Áp VLF 0.1Hz 30 Phút (Table 100.6.3)',
      'Tổn Hao Điện Môi Tan-Delta & Tip-Up (Table 100.6.7.1)',
      'Tiếp Địa Lộn Ngược Qua Lòng Biến Dòng Cửa Sổ CT'
    ]
  },
  {
    id: 'neta-evse',
    category: 'evse',
    title: 'Trạm Sạc Xe Điện (EV Charging Systems - EVSE)',
    subtitle: 'DC Fast Charger • AC Level 2 • Súng Sạc Kép • Giao tiếp CP/PP • Tiếp Địa PE',
    scope: 'Trạm sạc nhanh cao tốc, trạm sạc xe bus điện, hub sạc thương mại, bãi đỗ xe',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.26',
    standardName: 'NETA ATS-2025 Sec 7.26',
    aiEngine: 'EVSE Specialist AI Engine (NETA Sec 7.26)',
    icon: PlugZap,
    colorGradient: 'from-cyan-600 via-blue-600 to-indigo-800',
    borderColor: 'hover:border-cyan-400 focus:border-cyan-500',
    tagColor: 'bg-cyan-500/20 text-cyan-800 border-cyan-300',
    features: [
      'Điện Trở Dây Bảo Vệ PE (Bắt Buộc ≤ 0.50 Ω)',
      'Tín Hiệu CP (PWM) & PP (Dòng Định Mức Súng)',
      'Bảo Vệ Chống Dòng Rò RCD/CCID & Điện Áp Chạm',
      'Độ Gợn Sóng DC Ripple (≤ 2.0%) & Khóa Liên Động'
    ]
  },
  {
    id: 'neta-pv',
    category: 'pv',
    title: 'Hệ Thống Điện Mặt Trời & PV Cell (Solar PV & Cell Thermal IR)',
    subtitle: 'Chuyên Sâu PV Cell (IR & RGB AI) • IEC 62446-3 & Sitemark • Chuỗi Pin NETA ATS-2025 Mục 7.29',
    scope: 'Trạm điện mặt trời mặt đất, áp mái công nghiệp (C&I Rooftop), Solar Farm, Giám sát Nhiệt Drone/IR',
    standardCode: 'IEC 62446-3 & NETA ATS-2025 Mục 7.29',
    standardName: 'IEC 62446-3 & NETA ATS Sec 7.29 (Volateq / Sitemark)',
    aiEngine: 'PV Cell Thermal & Solar Specialist AI (IEC 62446-3)',
    icon: Sun,
    colorGradient: 'from-amber-500 via-orange-500 to-amber-700',
    borderColor: 'hover:border-amber-400 focus:border-amber-500',
    tagColor: 'bg-amber-500/20 text-amber-800 border-amber-300',
    features: [
      'Phân Tích Ảnh Nhiệt IR & Thực Tế RGB (IEC 62446-3)',
      'Thư Viện Mẫu Lỗi Sitemark / Volateq (Hotspot, Bypass, PID, JB)',
      'Ma Trận 3 Chiều: Severity (Level 1-4), Cause & Action',
      'Đo Lệch Voc Song Song (≤ 5%), I-V Curve & Tiếp Địa (≤ 5Ω)'
    ]
  },
  {
    id: 'neta-bess',
    category: 'bess',
    title: 'Hệ Thống Pin Lưu Trữ Năng Lượng (BESS)',
    subtitle: 'Container BESS • PCS • Lõi Pin Lithium / Flow • MV POI',
    scope: 'Container BESS, Khối Pin, Racks, Inverter/PCS, Trạm POI Trung Áp',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.28',
    standardName: 'NETA ATS-2025 Sec 7.28 & IEEE 1547',
    aiEngine: 'BESS Specialist AI Engine (NETA & IEEE 1547)',
    icon: BatteryCharging,
    colorGradient: 'from-teal-600 via-emerald-600 to-teal-800',
    borderColor: 'hover:border-teal-400 focus:border-teal-500',
    tagColor: 'bg-teal-500/20 text-teal-800 border-teal-300',
    features: [
      'PCCC Khí/Aerosol & Trạm Rửa Mắt (Bắt Buộc)',
      'Phân Hệ MBA 7.2, Máy Cắt 7.6.1, Cáp 7.3.2',
      'Đo Cực Tính (+/-) & Điện Áp DC Tổng',
      'Đánh Giá Sóng Hài POI IEEE 1547 (THDu < 5%)'
    ]
  },
  {
    id: 'neta-liquid-transformer',
    category: 'transformers',
    title: 'Máy Biến Áp Ngâm Chất Lỏng / Dầu (Liquid-Filled)',
    subtitle: 'Đánh Thủng D1816 • Ẩm D1533 • DGA Khí C2H2 (IEEE C57.104) • TTR (≤ 0.5%) • PI (≥ 1.0)',
    scope: 'Máy biến áp ngâm dầu khoáng (Mineral Oil), Silicone, Ester; trạm 110kV/22kV, công nghiệp',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.2.2',
    standardName: 'NETA ATS-2025 Sec 7.2.2 & Tables 100.3, 100.4, 100.5, IEEE C57.104',
    aiEngine: 'Liquid-Filled Transformer Specialist AI (Sec 7.2.2 & DGA)',
    icon: Droplets,
    colorGradient: 'from-amber-900 via-yellow-900 to-slate-900',
    borderColor: 'hover:border-amber-400 focus:border-amber-500',
    tagColor: 'bg-amber-500/20 text-amber-800 border-amber-300',
    features: [
      'Điện Áp Đánh Thủng Dầu ASTM D1816 (Table 100.4.1)',
      'Hàm Lượng Ẩm ASTM D1533 (Giới Hạn ≤ 20 ppm)',
      'DGA IEEE C57.104: Khí C2H2 Cảnh Báo Hồ Quang Điện',
      'Chỉ Số PI ≥ 1.0 & Tổn Hao Cuộn Dây PF ≤ 0.5%'
    ]
  },
  {
    id: 'neta-dry-large',
    category: 'transformers',
    title: 'Máy Biến Áp Khô Dung Lượng Lớn (Dry-Type Large)',
    subtitle: 'Công suất lớn (>600V hoặc >500 kVA) • Hệ thống làm mát cưỡng bức',
    scope: 'Trạm điện trung áp, nhà máy công nghiệp nặng, trung tâm dữ liệu',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.2.1.2',
    standardName: 'NETA ATS-2025 Sec 7.2.1.2',
    aiEngine: 'Large Transformer AI Engine (NETA ATS-2025)',
    icon: Zap,
    colorGradient: 'from-indigo-600 via-blue-600 to-indigo-800',
    borderColor: 'hover:border-indigo-400 focus:border-indigo-500',
    tagColor: 'bg-indigo-500/20 text-indigo-800 border-indigo-300',
    features: [
      'Cách Điện Lõi Thép (Core IR ≥ 1.0 MΩ @ 500V DC)',
      'Điện Trở Cuộn Dây All Taps (Lệch Quy Đổi ≤ 1.0%)',
      'Dòng Kích Thích Lõi 3 Trụ (2 Cao 1 Thấp)',
      'Chỉ Số Phân Cực PI ≥ 1.0 & Chống Sét Van'
    ]
  },
  {
    id: 'neta-dry-type',
    category: 'transformers',
    title: 'Máy Biến Áp Khô Hạ Áp Nhỏ (Dry-Type Small)',
    subtitle: 'Hạ áp phân phối (≤600V và ≤500 kVA) • Tự làm mát bằng không khí',
    scope: 'Tủ trạm kiosk hạ áp, tòa nhà thương mại, xưởng sản xuất',
    standardCode: 'ANSI/NETA ATS-2025 Mục 7.2.1.1',
    standardName: 'NETA ATS-2025 Sec 7.2.1.1',
    aiEngine: 'Small Dry-Type AI Engine (NETA ATS-2025)',
    icon: Zap,
    colorGradient: 'from-blue-600 via-cyan-600 to-blue-800',
    borderColor: 'hover:border-blue-400 focus:border-blue-500',
    tagColor: 'bg-blue-500/20 text-blue-800 border-blue-300',
    features: [
      'Điện Trở Cách Điện 1000V DC (Table 100.5)',
      'Hấp Thụ Điện Môi (DAR ≥ 1.0)',
      'Tỷ Số Biến Áp TTR (Sai Số ≤ 0.5%)',
      'Độ Lệch Mối Nối Bu-lông DLRO (≤ 50%)'
    ]
  },
  {
    id: 'standard',
    category: 'standard',
    title: 'Nhập Liệu Hiện Trường Tiêu Chuẩn (Standard CMMS)',
    subtitle: 'Biên bản tổng hợp: Động cơ, Dầu DGA, Máy cắt, Cáp ngầm...',
    scope: 'Mọi thiết bị điện trong cơ sở dữ liệu tài sản CMMS',
    standardCode: 'CMMS Standard Format',
    standardName: 'Hồ Sơ Bảo Trì Tiêu Chuẩn',
    aiEngine: 'General CMMS Health Index Engine',
    icon: ClipboardList,
    colorGradient: 'from-slate-700 via-slate-800 to-slate-900',
    borderColor: 'hover:border-slate-400 focus:border-slate-500',
    tagColor: 'bg-slate-500/20 text-slate-800 border-slate-300',
    features: [
      'Quét Mã QR Nhận Diện Thiết Bị Nhanh',
      'Nhập Liệu Đa Dạng Thông Số Kỹ Thuật',
      'Tạo Phiếu Đề Xuất Công Việc Tự Động',
      'Lưu Trực Tiếp Vào Lịch Sử Thiết Bị'
    ]
  }
];

interface Props {
  currentMode: FieldEntryMode;
  onSelectMode: (mode: FieldEntryMode) => void;
  reportsCount?: number;
  onNavigateToReports?: () => void;
}

export const EquipmentChecklistSelector: React.FC<Props> = ({
  currentMode,
  onSelectMode,
  reportsCount = 0,
  onNavigateToReports
}) => {
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedCategory, setSelectedCategory] = useState<'all' | 'grounding' | 'batteries' | 'ups' | 'relays' | 'motors' | 'pd' | 'thermography' | 'switches' | 'cables' | 'evse' | 'pv' | 'transformers' | 'bess' | 'standard'>('all');
  const [isExpanded, setIsExpanded] = useState(false);
  const [showUniversalCamera, setShowUniversalCamera] = useState(false);
  const [scanNotification, setScanNotification] = useState<string | null>(null);

  const activeItem = CHECKLIST_CATALOG.find(item => item.id === currentMode) || CHECKLIST_CATALOG[0];
  const ActiveIcon = activeItem.icon;

  const handleUniversalScanApply = (extracted: ExtractedNameplateData) => {
    const text = `${extracted.equipment_tag || ''} ${extracted.equipment_name || ''} ${extracted.equipment_type || ''} ${extracted.oil_type || ''} ${extracted.full_nameplate_text || ''}`.toLowerCase();
    let targetMode: FieldEntryMode = 'neta-liquid-transformer';

    if (text.includes('ground') || text.includes('tiếp địa') || text.includes('nối đất') || text.includes('earth') || text.includes('grid')) {
      targetMode = 'neta-grounding';
    } else if (text.includes('ats') || text.includes('transfer switch') || text.includes('chuyển nguồn') || text.includes('socomec') || text.includes('asco') || text.includes('vitzro')) {
      targetMode = 'neta-ats';
    } else if (text.includes('genset') || text.includes('generator set') || text.includes('máy phát khẩn') || text.includes('diesel generator') || text.includes('cummins') || text.includes('perkins') || text.includes('kohler') || text.includes('doosan')) {
      targetMode = 'neta-generator';
    } else if (text.includes('vrla') || text.includes('agm') || text.includes('gel') || text.includes('kín khí') || text.includes('valve-regulated')) {
      targetMode = 'neta-battery-vrla';
    } else if (text.includes('flooded') || text.includes('lead-acid') || text.includes('chì-axit') || text.includes('cell 2v') || text.includes('opzs') || (text.includes('battery') && !text.includes('bess')) || (text.includes('ắc quy') && !text.includes('ups'))) {
      targetMode = 'neta-battery-flooded';
    } else if (text.includes('sync') || text.includes('đồng bộ') || text.includes('synchronous') || text.includes('cực lồi') || text.includes('máy phát') || text.includes('generator') || text.includes('exciter')) {
      targetMode = 'neta-sync-machinery';
    } else if (text.includes('ups') || text.includes('uninterruptible') || text.includes('nguồn liên tục') || text.includes('static bypass') || text.includes('rectifier/inverter')) {
      targetMode = 'neta-ups';
    } else if (text.includes('relay') || text.includes('rơ-le') || text.includes('sel-') || text.includes('751') || text.includes('micom') || text.includes('siprotec') || text.includes('multilin')) {
      targetMode = 'neta-relay';
    } else if (text.includes('dầu') || text.includes('oil') || text.includes('liquid') || extracted.oil_type || extracted.oil_volume_liters) {
      targetMode = 'neta-liquid-transformer';
    } else if (text.includes('motor') || text.includes('động cơ') || text.includes('rpm') || text.includes('pole')) {
      targetMode = 'neta-motor';
    } else if (text.includes('inverter') || text.includes('solar') || text.includes('pv') || text.includes('quang điện')) {
      targetMode = 'neta-pv';
    } else if (extracted.rated_power_kva && extracted.rated_power_kva > 500) {
      targetMode = 'neta-dry-large';
    } else {
      targetMode = 'neta-dry-type';
    }

    onSelectMode(targetMode);
    setScanNotification(`AI đã nhận diện: "${extracted.equipment_tag || extracted.manufacturer || 'Thiết bị'}" -> Đang chuyển sang biểu mẫu tương ứng!`);
    setTimeout(() => setScanNotification(null), 6000);
  };

  const filteredItems = CHECKLIST_CATALOG.filter(item => {
    const matchesCategory = selectedCategory === 'all' || item.category === selectedCategory;
    const matchesSearch =
      item.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
      item.subtitle.toLowerCase().includes(searchQuery.toLowerCase()) ||
      item.standardCode.toLowerCase().includes(searchQuery.toLowerCase()) ||
      item.standardName.toLowerCase().includes(searchQuery.toLowerCase());
    return matchesCategory && matchesSearch;
  });

  return (
    <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden transition-all">
      {/* Universal Camera OCR Banner */}
      <div className="bg-gradient-to-r from-amber-500 via-amber-400 to-yellow-500 text-slate-950 p-3 sm:p-4 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 border-b border-amber-300 shadow-xs">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-slate-950 text-amber-300 flex items-center justify-center shrink-0 shadow-md">
            <Camera size={20} className="animate-pulse" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <span className="text-[10px] font-black uppercase tracking-wider bg-slate-950 text-amber-300 px-2 py-0.5 rounded shadow-2xs">
                QUÉT CAMERA NAMEPLATE
              </span>
              <span className="text-xs font-black text-slate-950">
                Nhận Diện Bảng Tên Thiết Bị Tự Động Điền Checklist
              </span>
            </div>
            <p className="text-xs text-slate-900 font-medium mt-0.5">
              Dùng camera điện thoại chụp bảng tên MBA, động cơ, inverter... AI tự động bóc tách thông số kỹ thuật & điền vào biểu mẫu.
            </p>
          </div>
        </div>

        <button
          type="button"
          onClick={() => setShowUniversalCamera(true)}
          className="px-4 py-2 bg-slate-950 hover:bg-slate-900 text-amber-300 font-black rounded-xl text-xs flex items-center gap-2 shadow-lg shrink-0 transition-all active:scale-95 border border-amber-400/30"
        >
          <Camera size={15} />
          <span>Mở Camera Quét Nhãn</span>
        </button>
      </div>

      {/* Notification Toast */}
      {scanNotification && (
        <div className="p-3 bg-emerald-600 text-white text-xs font-bold flex items-center justify-between">
          <div className="flex items-center gap-2">
            <CheckCircle2 size={16} />
            <span>{scanNotification}</span>
          </div>
          <button onClick={() => setScanNotification(null)} className="text-white/80 hover:text-white">
            ✕
          </button>
        </div>
      )}

      {/* Active Mode Compact Strip */}
      <div className="p-3 sm:p-4 bg-slate-900 text-white flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3">
        <div className="flex items-center gap-3">
          <div className={`w-10 h-10 rounded-xl bg-gradient-to-br ${activeItem.colorGradient} flex items-center justify-center text-white shadow-md shrink-0`}>
            <ActiveIcon size={20} />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <span className="text-[10px] font-extrabold uppercase tracking-wider px-2 py-0.5 rounded-full bg-white/10 text-white border border-white/20">
                {activeItem.standardCode}
              </span>
              <span className="text-[11px] text-emerald-400 font-bold flex items-center gap-1">
                <Sparkles size={12} />
                {activeItem.aiEngine}
              </span>
            </div>
            <h2 className="text-sm sm:text-base font-black text-white mt-0.5">
              {activeItem.title}
            </h2>
          </div>
        </div>

        <div className="flex items-center gap-2 self-end md:self-center">
          {onNavigateToReports && (
            <button
              onClick={onNavigateToReports}
              className="px-3 py-1.5 rounded-xl text-xs font-bold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-colors shadow-xs"
              title="Xem tất cả báo cáo đã lưu"
            >
              <FolderArchive size={14} className="text-amber-400" />
              <span>Kho Báo Cáo ({reportsCount})</span>
            </button>
          )}

          <button
            onClick={() => setIsExpanded(!isExpanded)}
            className="px-3.5 py-1.5 rounded-xl text-xs font-bold bg-blue-600 hover:bg-blue-500 text-white flex items-center gap-1.5 transition-all shadow-md"
          >
            <Layers size={14} />
            <span>{isExpanded ? 'Đóng Danh Mục' : 'Đổi Thiết Bị Kiểm Định'}</span>
            <ChevronRight size={14} className={`transform transition-transform ${isExpanded ? 'rotate-90' : ''}`} />
          </button>
        </div>
      </div>

      {/* Expanded Directory & Selection Hub */}
      {isExpanded && (
        <div className="p-4 sm:p-6 bg-slate-50 border-t border-slate-200 space-y-4 animate-in slide-in-from-top-2 duration-200">
          <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3">
            {/* Category Pills */}
            <div className="flex flex-wrap items-center gap-1.5 text-xs font-semibold">
              <button
                onClick={() => setSelectedCategory('all')}
                className={`px-3 py-1.5 rounded-lg transition-colors ${
                  selectedCategory === 'all'
                    ? 'bg-blue-600 text-white shadow-xs'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                Tất cả thiết bị ({CHECKLIST_CATALOG.length})
              </button>
              <button
                onClick={() => setSelectedCategory('grounding')}
                className={`px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1 ${
                  selectedCategory === 'grounding'
                    ? 'bg-teal-700 text-white shadow-xs font-bold'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <ShieldCheck size={13} />
                Tiếp địa & Nối đất (1)
              </button>
              <button
                onClick={() => setSelectedCategory('batteries')}
                className={`px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1 ${
                  selectedCategory === 'batteries'
                    ? 'bg-amber-700 text-white shadow-xs font-bold'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <BatteryCharging size={13} />
                Ắc quy DC (Flooded & VRLA) (2)
              </button>
              <button
                onClick={() => setSelectedCategory('ups')}
                className={`px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1 ${
                  selectedCategory === 'ups'
                    ? 'bg-emerald-700 text-white shadow-xs font-bold'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <BatteryCharging size={13} />
                Nguồn UPS (1)
              </button>
              <button
                onClick={() => setSelectedCategory('relays')}
                className={`px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1 ${
                  selectedCategory === 'relays'
                    ? 'bg-amber-600 text-white shadow-xs font-bold'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <Cpu size={13} />
                Rơ-le bảo vệ (1)
              </button>
              <button
                onClick={() => setSelectedCategory('motors')}
                className={`px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1 ${
                  selectedCategory === 'motors'
                    ? 'bg-blue-900 text-white shadow-xs'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <Power size={13} />
                Động cơ & Máy phát (3)
              </button>
              <button
                onClick={() => setSelectedCategory('pd')}
                className={`px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1 ${
                  selectedCategory === 'pd'
                    ? 'bg-violet-800 text-white shadow-xs'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <Radio size={13} />
                Phóng điện PD (1)
              </button>
              <button
                onClick={() => setSelectedCategory('thermography')}
                className={`px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1 ${
                  selectedCategory === 'thermography'
                    ? 'bg-red-700 text-white shadow-xs'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <Thermometer size={13} />
                Ảnh nhiệt IR (1)
              </button>
              <button
                onClick={() => setSelectedCategory('switches')}
                className={`px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1 ${
                  selectedCategory === 'switches'
                    ? 'bg-teal-700 text-white shadow-xs'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <GitBranch size={13} />
                Cầu dao & ATS (2)
              </button>
              <button
                onClick={() => setSelectedCategory('cables')}
                className={`px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1 ${
                  selectedCategory === 'cables'
                    ? 'bg-purple-700 text-white shadow-xs'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <Cable size={13} />
                Cáp điện MV/HV (1)
              </button>
              <button
                onClick={() => setSelectedCategory('evse')}
                className={`px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1 ${
                  selectedCategory === 'evse'
                    ? 'bg-cyan-700 text-white shadow-xs'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <PlugZap size={13} />
                Sạc xe điện EVSE (1)
              </button>
              <button
                onClick={() => setSelectedCategory('pv')}
                className={`px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1 ${
                  selectedCategory === 'pv'
                    ? 'bg-amber-600 text-white shadow-xs'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <Sun size={13} />
                Điện mặt trời PV (1)
              </button>
              <button
                onClick={() => setSelectedCategory('bess')}
                className={`px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1 ${
                  selectedCategory === 'bess'
                    ? 'bg-teal-700 text-white shadow-xs'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <BatteryCharging size={13} />
                Pin BESS (1)
              </button>
              <button
                onClick={() => setSelectedCategory('transformers')}
                className={`px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1 ${
                  selectedCategory === 'transformers'
                    ? 'bg-indigo-700 text-white shadow-xs'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <Zap size={13} />
                Máy biến áp (3)
              </button>
              <button
                onClick={() => setSelectedCategory('standard')}
                className={`px-3 py-1.5 rounded-lg transition-colors flex items-center gap-1 ${
                  selectedCategory === 'standard'
                    ? 'bg-slate-800 text-white shadow-xs'
                    : 'bg-white text-slate-600 hover:bg-slate-200 border border-slate-200'
                }`}
              >
                <ClipboardList size={13} />
                Biểu mẫu chung (1)
              </button>
            </div>

            {/* Search Box */}
            <div className="relative max-w-xs w-full">
              <Search size={14} className="absolute left-3 top-2.5 text-slate-400" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Tìm checklist theo tên hoặc NETA code..."
                className="w-full pl-8 pr-3 py-1.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-blue-500 font-medium text-slate-800"
              />
            </div>
          </div>

          {/* Cards Grid */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3.5">
            {filteredItems.map((item) => {
              const ItemIcon = item.icon;
              const isSelected = item.id === currentMode;

              return (
                <div
                  key={item.id}
                  onClick={() => {
                    onSelectMode(item.id);
                    setIsExpanded(false);
                  }}
                  className={`p-4 rounded-xl border-2 cursor-pointer transition-all flex flex-col justify-between ${
                    isSelected
                      ? 'bg-white border-blue-600 shadow-md ring-2 ring-blue-500/20'
                      : 'bg-white border-slate-200 hover:border-slate-300 hover:shadow-xs'
                  }`}
                >
                  <div>
                    <div className="flex items-start justify-between gap-3">
                      <div className="flex items-center gap-2.5">
                        <div className={`w-9 h-9 rounded-lg bg-gradient-to-br ${item.colorGradient} flex items-center justify-center text-white shrink-0 shadow-xs`}>
                          <ItemIcon size={18} />
                        </div>
                        <div>
                          <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full border ${item.tagColor}`}>
                            {item.standardCode}
                          </span>
                          <h3 className="font-bold text-sm text-slate-900 mt-1">
                            {item.title}
                          </h3>
                        </div>
                      </div>

                      {isSelected ? (
                        <span className="px-2.5 py-1 rounded-full text-[10px] font-bold bg-blue-600 text-white flex items-center gap-1 shadow-xs shrink-0">
                          <CheckCircle2 size={12} />
                          Đang chọn
                        </span>
                      ) : (
                        <span className="px-2.5 py-1 rounded-full text-[10px] font-semibold bg-slate-100 text-slate-600 hover:bg-slate-200 flex items-center gap-1 shrink-0">
                          Chọn mẫu
                          <ArrowRight size={11} />
                        </span>
                      )}
                    </div>

                    <p className="text-xs text-slate-500 mt-2 line-clamp-2">
                      {item.subtitle}
                    </p>

                    <div className="grid grid-cols-2 gap-1.5 mt-3 pt-3 border-t border-slate-100">
                      {item.features.map((feat, idx) => (
                        <div key={idx} className="text-[10px] text-slate-600 flex items-center gap-1 truncate">
                          <span className="w-1.5 h-1.5 rounded-full bg-blue-500 shrink-0" />
                          <span className="truncate">{feat}</span>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div className="mt-3 pt-2 text-[10px] text-slate-400 flex items-center justify-between border-t border-slate-50">
                    <span className="flex items-center gap-1 text-blue-600 font-semibold">
                      <Sparkles size={11} />
                      {item.aiEngine}
                    </span>
                    <span>Chuẩn: {item.standardName}</span>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Universal Nameplate Scanner Modal */}
      {showUniversalCamera && (
        <NameplateScannerModal
          isOpen={showUniversalCamera}
          onClose={() => setShowUniversalCamera(false)}
          onApplyData={handleUniversalScanApply}
          currentTag={activeItem.title}
        />
      )}
    </div>
  );
};
