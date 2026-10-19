import React, { useState } from 'react';
import {
  Activity,
  Zap,
  Flame,
  Wind,
  Layers,
  Search,
  Filter,
  CheckCircle2,
  ChevronRight,
  Sparkles,
  Cpu,
  Clock,
  AudioWaveform,
  Thermometer,
  ShieldAlert,
  Sliders,
  ChevronDown,
  LayoutGrid,
  ExternalLink,
  ShieldCheck
} from 'lucide-react';

export type DeepAnalysisModuleId =
  | 'partial-discharge'
  | 'dga'
  | 'pv-cell'
  | 'wind-turbine'
  | 'vlf-cable'
  | 'motor-vibration'
  | 'sf6-gas';

export interface DeepAnalysisModuleItem {
  id: DeepAnalysisModuleId;
  title: string;
  titleEn: string;
  shortTitle: string;
  subLabel: string;
  category: 'switchgear_pd' | 'transformer' | 'renewable' | 'upcoming';
  standard: string;
  aiTech: string;
  badge: string;
  badgeColor: string;
  icon: React.ComponentType<{ size?: number; className?: string }>;
  description: string;
  isAvailable: boolean;
  accentColor: string;
  activeBorder: string;
  activeRing: string;
  activeGradient: string;
}

export const DEEP_ANALYSIS_CATALOG: DeepAnalysisModuleItem[] = [
  {
    id: 'partial-discharge',
    title: 'Chẩn đoán Phóng điện Cục bộ Đa phương thức (PD UltraTEV Plus² & Nhiệt ảnh IR)',
    titleEn: 'Multi-Modal Partial Discharge (AI Vision PRPD & Thermal IR)',
    shortTitle: '⚡ Phóng Điện PD (EA Tech)',
    subLabel: 'UltraTEV² • PRPD Vision • Siêu Âm',
    category: 'switchgear_pd',
    standard: 'EA Technology • NETA MTS Table 100.23 • DNO CNAIM',
    aiTech: 'Gemini 2.5 Flash Vision + Canvas PRPD Cloud Simulator',
    badge: 'AI Vision & Simulator',
    badgeColor: 'bg-indigo-500/20 text-indigo-300 border-indigo-400/40',
    icon: AudioWaveform,
    description: 'Nhận diện 5 dạng đám mây PRPD (Void, Surface, Corona, Floating, Spark Gap), đồng bộ mô phỏng Canvas, tương quan nhiệt ảnh IR và xuất biên bản PDF chuẩn TEV A4.',
    isAvailable: true,
    accentColor: 'text-indigo-400',
    activeBorder: 'border-indigo-400',
    activeRing: 'ring-2 ring-indigo-400/40',
    activeGradient: 'from-indigo-950 via-slate-900 to-indigo-900/60'
  },
  {
    id: 'dga',
    title: 'Phân tích Khí Hòa tan DGA Máy Biến Áp (Dissolved Gas Analysis)',
    titleEn: 'Transformer DGA & Gas Rate of Generation Analysis',
    shortTitle: '🌡️ DGA MBA Dầu',
    subLabel: 'Khí Cháy • Tam Giác Duval 1,4,5 • Rogers',
    category: 'transformer',
    standard: 'IEEE C57.104-2019 • IEC 60599 • Tam giác Duval',
    aiTech: 'Duval Coordinate Engine & Historical Trend Predictive AI',
    badge: 'IEEE C57.104',
    badgeColor: 'bg-blue-500/20 text-blue-300 border-blue-400/40',
    icon: Thermometer,
    description: 'Chẩn đoán 7 khí cháy (C2H2, H2, CH4, C2H4, C2H6, CO, CO2), đối soát nồng độ cơ sở dữ liệu lịch sử và phân loại dạng sự cố phóng điện/quá nhiệt.',
    isAvailable: true,
    accentColor: 'text-blue-400',
    activeBorder: 'border-blue-400',
    activeRing: 'ring-2 ring-blue-400/40',
    activeGradient: 'from-blue-950 via-slate-900 to-blue-900/60'
  },
  {
    id: 'pv-cell',
    title: 'Phân tích Tấm Pin & Cell Điện Mặt Trời (Photovoltaic Thermal IR & RGB)',
    titleEn: 'Solar PV Module & Cell Deep Thermal Defect Analyzer',
    shortTitle: '🔥 PV Solar Cells',
    subLabel: 'Nhiệt Ảnh Cell • Hotspot • Grading',
    category: 'renewable',
    standard: 'IEC 62446-3 • Volateq / Sitemark • NETA ATS 7.25',
    aiTech: 'Computer Vision Defect Segmentation & String Heatmap',
    badge: 'IEC 62446-3',
    badgeColor: 'bg-amber-500/20 text-amber-300 border-amber-400/40',
    icon: Flame,
    description: 'Nhận diện điểm nóng Hotspot, nứt cell micro-crack, đứt diode bypass, phân loại mức độ tổn hao công suất và đề xuất khắc phục.',
    isAvailable: true,
    accentColor: 'text-amber-400',
    activeBorder: 'border-amber-400',
    activeRing: 'ring-2 ring-amber-400/40',
    activeGradient: 'from-amber-950 via-slate-900 to-amber-900/60'
  },
  {
    id: 'wind-turbine',
    title: 'Phân tích Khảo sát & Hư hại Cánh Quạt Điện Gió (Wind Blade Inspection)',
    titleEn: 'Wind Turbine Blade Damage Severity Classifier',
    shortTitle: '💨 Wind Turbine Blade',
    subLabel: 'Cánh Quạt Gió • Nứt Rạn • Sét Đánh',
    category: 'renewable',
    standard: 'EPRI 3002019669 • DNV GL • IEC 61400-5 Severity 1-5',
    aiTech: 'Multi-Perspective Blade Visual AI & Work Order Link',
    badge: 'EPRI 3002019669',
    badgeColor: 'bg-teal-500/20 text-teal-300 border-teal-400/40',
    icon: Wind,
    description: 'Đánh giá xói mòn mép đón gió (Leading Edge), bong tách lớp sợi thủy tinh (Delamination), nứt rạn mép thoát và hư hại sét đánh.',
    isAvailable: true,
    accentColor: 'text-teal-400',
    activeBorder: 'border-teal-400',
    activeRing: 'ring-2 ring-teal-400/40',
    activeGradient: 'from-teal-950 via-slate-900 to-teal-900/60'
  },
  {
    id: 'vlf-cable',
    title: 'Đo Tổn hao Điện môi Cáp Ngầm VLF Tan Delta & PD Offline (Coming Soon)',
    titleEn: 'VLF Tan Delta & Offline Cable Diagnostic Engine',
    shortTitle: 'Cáp Ngầm VLF & TD',
    subLabel: 'Tan Delta 0.1Hz • Cây Nước XLPE',
    category: 'upcoming',
    standard: 'IEEE 400.2 • IEEE 400.3 • ICEA S-94-649',
    aiTech: 'Tan Delta Stability & Voltage Withstand Differential',
    badge: 'Lộ trình 2026',
    badgeColor: 'bg-slate-700 text-slate-300 border-slate-600',
    icon: Activity,
    description: 'Đánh giá cây nước (Water Treeing), thoái hóa cách điện XLPE/EPR và tổn hao điện môi theo tần số rất thấp 0.1Hz.',
    isAvailable: false,
    accentColor: 'text-slate-400',
    activeBorder: 'border-slate-600',
    activeRing: 'ring-1 ring-slate-600',
    activeGradient: 'from-slate-900 to-slate-950'
  },
  {
    id: 'motor-vibration',
    title: 'Phân tích Phổ Rung Động Động cơ & Bạc đạn Motor (Coming Soon)',
    titleEn: 'Rotating Machinery Vibration FFT & Bearing Health',
    shortTitle: 'Rung Động Động Cơ',
    subLabel: 'Phổ FFT • Vòng Bi Bạc Đạn',
    category: 'upcoming',
    standard: 'ISO 10816-3 • ISO 20816 • ANSI/NETA MTS 7.15',
    aiTech: 'FFT Harmonics Peak Detector & Bearing Fault Frequency (BPFO/BPFI)',
    badge: 'Lộ trình 2026',
    badgeColor: 'bg-slate-700 text-slate-300 border-slate-600',
    icon: Cpu,
    description: 'Phân tích phổ rung động theo trục ngang/dọc/dọc trục để phát hiện lệch tâm, mất cân bằng động và hư hại ổ bi.',
    isAvailable: false,
    accentColor: 'text-slate-400',
    activeBorder: 'border-slate-600',
    activeRing: 'ring-1 ring-slate-600',
    activeGradient: 'from-slate-900 to-slate-950'
  }
];

interface Props {
  activeModuleId: DeepAnalysisModuleId;
  onSelectModule: (id: DeepAnalysisModuleId) => void;
  lang?: 'vi' | 'en';
}

export const DeepAnalysisModuleSelector: React.FC<Props> = ({
  activeModuleId,
  onSelectModule,
  lang = 'vi'
}) => {
  const [isCatalogExpanded, setIsCatalogExpanded] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedCategory, setSelectedCategory] = useState<string>('all');

  const isVi = lang === 'vi';
  const currentModule =
    DEEP_ANALYSIS_CATALOG.find((m) => m.id === activeModuleId) || DEEP_ANALYSIS_CATALOG[0];

  const availableModules = DEEP_ANALYSIS_CATALOG.filter((m) => m.isAvailable);

  const categories = [
    { id: 'all', label: isVi ? 'Tất cả Module' : 'All Modules', count: DEEP_ANALYSIS_CATALOG.length },
    { id: 'switchgear_pd', label: isVi ? 'Tủ Đóng Cắt & PD' : 'Switchgear & PD', count: 1 },
    { id: 'transformer', label: isVi ? 'Máy Biến Áp' : 'Transformers', count: 1 },
    { id: 'renewable', label: isVi ? 'Năng Lượng Tái Tạo' : 'Renewable', count: 2 },
    { id: 'upcoming', label: isVi ? 'Lộ Trình Mở Rộng' : 'Upcoming', count: 2 }
  ];

  const filteredModules = DEEP_ANALYSIS_CATALOG.filter((m) => {
    const matchSearch =
      m.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
      m.titleEn.toLowerCase().includes(searchQuery.toLowerCase()) ||
      m.standard.toLowerCase().includes(searchQuery.toLowerCase()) ||
      m.description.toLowerCase().includes(searchQuery.toLowerCase());
    const matchCat = selectedCategory === 'all' || m.category === selectedCategory;
    return matchSearch && matchCat;
  });

  return (
    <div className="bg-slate-950 rounded-2xl border border-indigo-500/30 shadow-2xl overflow-hidden transition-all">
      {/* 1. TOP HEADER BRAND & ACTION STRIP */}
      <div className="p-3.5 sm:p-4 bg-gradient-to-r from-slate-950 via-indigo-950/80 to-slate-950 border-b border-indigo-500/20 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-indigo-500 via-indigo-600 to-indigo-800 text-white flex items-center justify-center font-black text-sm shrink-0 shadow-lg shadow-indigo-600/30 border border-indigo-400/40">
            COE
          </div>
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <h2 className="text-sm sm:text-base font-black text-white tracking-wide uppercase">
                {isVi ? 'Phân Tích Chuyên Sâu' : 'Deep Diagnostics Center of Excellence'}
              </h2>
              <span className="text-[10px] px-2 py-0.5 rounded font-mono font-bold bg-indigo-500/20 text-indigo-300 border border-indigo-500/40">
                TEV AI ENGINE
              </span>
              <span className="text-[10px] px-2 py-0.5 rounded font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 flex items-center gap-1">
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse"></span>
                <span>{availableModules.length} Module Sẵn Sàng</span>
              </span>
            </div>
            <p className="text-xs text-slate-400 mt-0.5">
              {isVi
                ? 'Trung tâm chẩn đoán đa phương thức & mô phỏng chuẩn IEEE, IEC, ANSI/NETA cho Kỹ sư Hiện trường (FSE)'
                : 'Multi-modal diagnostic & simulation suite conforming to IEEE, IEC, and ANSI/NETA standards.'}
            </p>
          </div>
        </div>

        {/* Catalog Drawer Toggle */}
        <button
          type="button"
          onClick={() => setIsCatalogExpanded(!isCatalogExpanded)}
          className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all border flex items-center gap-1.5 shrink-0 ${
            isCatalogExpanded
              ? 'bg-indigo-600 text-white border-indigo-400 shadow-md shadow-indigo-600/30'
              : 'bg-slate-900 hover:bg-slate-850 text-indigo-300 hover:text-white border-indigo-500/30'
          }`}
          title={isVi ? 'Mở danh mục tra cứu tất cả các module' : 'Open full module catalog'}
        >
          <LayoutGrid size={14} />
          <span>
            {isCatalogExpanded
              ? isVi
                ? 'Đóng Danh Mục'
                : 'Close Catalog'
              : isVi
              ? '📚 Thư Viện Đầy Đủ'
              : '📚 Full Catalog'}
          </span>
          <span className="text-[10px] px-1.5 py-0.2 rounded-full bg-slate-950 font-mono font-black text-indigo-200">
            {DEEP_ANALYSIS_CATALOG.length}
          </span>
        </button>
      </div>

      {/* 2. FULL-WIDTH EXPANDED MODULE SELECTOR BAR (ZERO SCROLLBAR, 100% VISIBLE FOR FSE) */}
      <div className="p-3 sm:p-4 bg-slate-950/95 border-b border-slate-800">
        <div className="text-[11px] font-bold uppercase tracking-wider text-slate-400 mb-2.5 flex items-center justify-between">
          <span className="flex items-center gap-1.5">
            <Sliders size={13} className="text-indigo-400" />
            <span>{isVi ? 'Chọn Phân Hệ Phân Tích Chuyên Sâu:' : 'Select Diagnostic Engine:'}</span>
          </span>
          <span className="text-[10px] text-slate-500 hidden sm:inline">
            {isVi ? 'Bấm để chuyển đổi tức thời giữa các module' : 'Instant 1-click switcher for field engineers'}
          </span>
        </div>

        {/* Responsive Grid with No Horizontal Scrollbar */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-2.5 sm:gap-3">
          {availableModules.map((m) => {
            const isSelected = activeModuleId === m.id;
            const IconComp = m.icon;

            return (
              <button
                key={m.id}
                type="button"
                onClick={() => onSelectModule(m.id)}
                className={`relative p-3 sm:p-3.5 rounded-xl border text-left transition-all duration-150 flex flex-col justify-between group cursor-pointer ${
                  isSelected
                    ? `bg-gradient-to-br ${m.activeGradient} ${m.activeBorder} ${m.activeRing} shadow-lg shadow-black/40`
                    : 'bg-slate-900/90 hover:bg-slate-900 border-slate-800 hover:border-slate-700 text-slate-300'
                }`}
              >
                {/* Top row of card */}
                <div className="flex items-start justify-between gap-2 mb-2">
                  <div
                    className={`w-8 h-8 rounded-lg flex items-center justify-center shrink-0 border transition-all ${
                      isSelected
                        ? 'bg-slate-950 text-white border-white/20 shadow-sm'
                        : 'bg-slate-950/80 text-slate-400 border-slate-800 group-hover:text-white'
                    }`}
                  >
                    <IconComp size={16} className={isSelected ? m.accentColor : 'text-slate-400'} />
                  </div>

                  {isSelected ? (
                    <span className="flex items-center gap-1 text-[10px] font-black px-1.5 py-0.5 rounded bg-emerald-500/20 text-emerald-300 border border-emerald-500/40">
                      <CheckCircle2 size={12} className="text-emerald-400" />
                      <span>{isVi ? 'Đang Mở' : 'Active'}</span>
                    </span>
                  ) : (
                    <span className={`text-[9px] px-1.5 py-0.2 rounded font-bold border ${m.badgeColor}`}>
                      {m.badge}
                    </span>
                  )}
                </div>

                {/* Title & SubLabel */}
                <div>
                  <div
                    className={`text-xs sm:text-sm font-black tracking-tight line-clamp-1 ${
                      isSelected ? 'text-white' : 'text-slate-200 group-hover:text-white'
                    }`}
                  >
                    {m.shortTitle}
                  </div>
                  <div className="text-[10px] text-slate-400 mt-0.5 line-clamp-1 font-medium">
                    {m.subLabel}
                  </div>
                </div>

                {/* Standard Code Footer */}
                <div className="mt-2 pt-2 border-t border-slate-800/80 flex items-center justify-between text-[10px]">
                  <span className="font-mono text-slate-500 font-semibold truncate max-w-[140px]">
                    {m.standard.split('•')[0].trim()}
                  </span>
                  <span
                    className={`font-bold transition-transform group-hover:translate-x-0.5 ${
                      isSelected ? m.accentColor : 'text-slate-600 group-hover:text-slate-300'
                    }`}
                  >
                    →
                  </span>
                </div>
              </button>
            );
          })}
        </div>
      </div>

      {/* 3. ACTIVE MODULE STATUS STRIP */}
      <div className="px-4 py-3 bg-gradient-to-r from-slate-950 via-slate-900 to-slate-950 text-white flex flex-col md:flex-row items-start md:items-center justify-between gap-3 border-b border-slate-800/80">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-xl bg-indigo-500/10 text-indigo-400 border border-indigo-500/30 flex items-center justify-center shrink-0">
            {React.createElement(currentModule.icon, { size: 18 })}
          </div>
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-xs sm:text-sm font-black text-white">
                {isVi ? currentModule.title : currentModule.titleEn}
              </span>
              <span className={`text-[10px] font-bold px-2 py-0.5 rounded-full border ${currentModule.badgeColor}`}>
                {currentModule.badge}
              </span>
            </div>
            <div className="text-[11px] text-slate-400 flex items-center gap-2 flex-wrap mt-0.5">
              <span className="font-semibold text-slate-300">{currentModule.standard}</span>
              <span className="text-slate-600 hidden sm:inline">•</span>
              <span className="text-indigo-300 font-medium hidden sm:inline">{currentModule.aiTech}</span>
            </div>
          </div>
        </div>

        <div className="flex items-center gap-2 self-start md:self-auto shrink-0">
          <span className="text-[11px] font-bold px-2.5 py-1 rounded-lg bg-emerald-500/20 text-emerald-300 border border-emerald-400/30 flex items-center gap-1.5 shadow-sm">
            <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
            <span>{isVi ? 'Đang kích hoạt' : 'Active Engine'}</span>
          </span>
        </div>
      </div>

      {/* 4. EXPANDABLE FULL CATALOG DRAWER (SEARCH & ROADMAP MODULES) */}
      {isCatalogExpanded && (
        <div className="p-4 sm:p-6 bg-slate-950 border-t border-slate-800 space-y-4 animate-in fade-in duration-200">
          <div className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3 pb-3 border-b border-slate-800">
            <div>
              <h3 className="text-sm font-black text-white flex items-center gap-2">
                <LayoutGrid size={16} className="text-indigo-400" />
                <span>{isVi ? 'Thư Viện Đầy Đủ Phân Tích Chuyên Sâu TEV COE' : 'Deep Diagnostic Module Catalog'}</span>
              </h3>
              <p className="text-xs text-slate-400 mt-0.5">
                {isVi
                  ? 'Tra cứu toàn bộ module đã phát hành và các module trong lộ trình công nghệ 2026'
                  : 'Search all available modules and technology roadmap modules.'}
              </p>
            </div>

            {/* Search Input */}
            <div className="relative w-full sm:w-64">
              <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder={isVi ? 'Tìm kiếm tiêu chuẩn, lỗi...' : 'Search module or standard...'}
                className="w-full pl-8 pr-3 py-1.5 bg-slate-900 border border-slate-700 rounded-xl text-xs text-white placeholder-slate-500 focus:outline-hidden focus:border-indigo-400"
              />
            </div>
          </div>

          {/* Category Filter Pills */}
          <div className="flex items-center gap-1.5 overflow-x-auto scrollbar-none pb-1">
            {categories.map((c) => (
              <button
                key={c.id}
                type="button"
                onClick={() => setSelectedCategory(c.id)}
                className={`px-3 py-1 rounded-lg text-xs font-bold transition-all shrink-0 ${
                  selectedCategory === c.id
                    ? 'bg-indigo-600 text-white shadow-xs'
                    : 'bg-slate-900 text-slate-300 hover:text-white hover:bg-slate-800 border border-slate-800'
                }`}
              >
                {c.label} ({c.count})
              </button>
            ))}
          </div>

          {/* Module Cards Grid */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3.5 pt-1">
            {filteredModules.map((item) => {
              const isSelected = activeModuleId === item.id;
              const IconComp = item.icon;

              return (
                <div
                  key={item.id}
                  onClick={() => {
                    if (item.isAvailable) {
                      onSelectModule(item.id);
                      setIsCatalogExpanded(false);
                    }
                  }}
                  className={`p-4 rounded-xl border text-left transition-all relative flex flex-col justify-between ${
                    isSelected
                      ? 'bg-indigo-950/70 border-indigo-400 shadow-lg shadow-indigo-950/50 ring-2 ring-indigo-400/30'
                      : item.isAvailable
                      ? 'bg-slate-900/80 hover:bg-slate-900 border-slate-800 hover:border-slate-700 cursor-pointer'
                      : 'bg-slate-950/40 border-slate-900 opacity-60 cursor-not-allowed'
                  }`}
                >
                  <div className="space-y-2">
                    <div className="flex items-start justify-between gap-2">
                      <div className="flex items-center gap-2.5">
                        <div
                          className={`w-9 h-9 rounded-lg flex items-center justify-center shrink-0 ${
                            isSelected ? 'bg-indigo-600 text-white' : 'bg-slate-800 text-slate-300'
                          }`}
                        >
                          <IconComp size={18} />
                        </div>
                        <div>
                          <h4 className="text-xs font-black text-white line-clamp-1">
                            {isVi ? item.title : item.titleEn}
                          </h4>
                          <span className="text-[10px] font-mono text-indigo-300 font-semibold">
                            {item.standard}
                          </span>
                        </div>
                      </div>

                      <span className={`text-[9px] px-2 py-0.5 rounded font-black tracking-wide border shrink-0 ${item.badgeColor}`}>
                        {item.badge}
                      </span>
                    </div>

                    <p className="text-[11px] text-slate-400 leading-relaxed line-clamp-2">
                      {item.description}
                    </p>
                  </div>

                  <div className="pt-3 mt-3 border-t border-slate-800/80 flex items-center justify-between">
                    <span className="text-[10px] text-slate-500 font-medium">
                      {item.aiTech}
                    </span>

                    {item.isAvailable ? (
                      <span
                        className={`text-[11px] font-bold flex items-center gap-1 ${
                          isSelected ? 'text-indigo-400 font-black' : 'text-slate-300'
                        }`}
                      >
                        {isSelected ? (isVi ? 'Đang mở' : 'Selected') : (isVi ? 'Kích hoạt →' : 'Launch →')}
                      </span>
                    ) : (
                      <span className="text-[10px] text-slate-500 font-bold bg-slate-900 px-2 py-0.5 rounded border border-slate-800">
                        {isVi ? 'Sắp ra mắt' : 'Roadmap 2026'}
                      </span>
                    )}
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
