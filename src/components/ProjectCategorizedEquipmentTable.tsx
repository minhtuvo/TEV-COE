import React, { useState, useMemo } from 'react';
import {
  Search,
  Filter,
  CheckCircle2,
  Clock,
  Play,
  Trash2,
  Plus,
  Layers,
  Copy,
  Check,
  Tag,
  ArrowRight,
  ShieldCheck,
  Sparkles,
  Zap
} from 'lucide-react';
import { ProjectMasterRecord, ProjectEquipmentItem } from '../services/masterDirectoryService';

export type EquipmentCategoryKey =
  | 'all'
  | 'transformer'
  | 'cable'
  | 'switchgear'
  | 'motor'
  | 'grounding'
  | 'dc_power'
  | 'other';

export interface CategoryInfo {
  key: EquipmentCategoryKey;
  label: string;
  icon: string;
  badgeBg: string;
  badgeText: string;
  border: string;
}

export const CATEGORY_DEFINITIONS: Record<EquipmentCategoryKey, CategoryInfo> = {
  all: {
    key: 'all',
    label: 'Tất cả thiết bị',
    icon: '⚡',
    badgeBg: 'bg-slate-100',
    badgeText: 'text-slate-800',
    border: 'border-slate-300'
  },
  transformer: {
    key: 'transformer',
    label: 'Máy biến áp (Dầu & Khô)',
    icon: '🔄',
    badgeBg: 'bg-amber-100',
    badgeText: 'text-amber-900',
    border: 'border-amber-300'
  },
  cable: {
    key: 'cable',
    label: 'Cáp điện & Busway',
    icon: '🔌',
    badgeBg: 'bg-blue-100',
    badgeText: 'text-blue-900',
    border: 'border-blue-300'
  },
  switchgear: {
    key: 'switchgear',
    label: 'Tủ điện & Máy cắt (ACB/VCB)',
    icon: '🛡️',
    badgeBg: 'bg-indigo-100',
    badgeText: 'text-indigo-900',
    border: 'border-indigo-300'
  },
  motor: {
    key: 'motor',
    label: 'Động cơ & Máy phát',
    icon: '⚙️',
    badgeBg: 'bg-emerald-100',
    badgeText: 'text-emerald-900',
    border: 'border-emerald-300'
  },
  grounding: {
    key: 'grounding',
    label: 'Hệ tiếp địa & Chống sét',
    icon: '⚡',
    badgeBg: 'bg-cyan-100',
    badgeText: 'text-cyan-900',
    border: 'border-cyan-300'
  },
  dc_power: {
    key: 'dc_power',
    label: 'Pin, UPS, BESS & Solar',
    icon: '🔋',
    badgeBg: 'bg-purple-100',
    badgeText: 'text-purple-900',
    border: 'border-purple-300'
  },
  other: {
    key: 'other',
    label: 'Rơ le, Chẩn đoán & Khác',
    icon: '🔍',
    badgeBg: 'bg-rose-100',
    badgeText: 'text-rose-900',
    border: 'border-rose-300'
  }
};

export function classifyEquipment(type?: string, name?: string): EquipmentCategoryKey {
  const text = `${type || ''} ${name || ''}`.toLowerCase();
  if (
    text.includes('biến áp') ||
    text.includes('transformer') ||
    text.includes('mba') ||
    text.includes('dầu') ||
    text.includes('dry-type') ||
    text.includes('liquid')
  ) {
    return 'transformer';
  }
  if (
    text.includes('cáp') ||
    text.includes('cable') ||
    text.includes('busway') ||
    text.includes('thanh dẫn') ||
    text.includes('đường dây') ||
    text.includes('busduct')
  ) {
    return 'cable';
  }
  if (
    text.includes('máy cắt') ||
    text.includes('breaker') ||
    text.includes('switchgear') ||
    text.includes('tủ điện') ||
    text.includes('acb') ||
    text.includes('vcb') ||
    text.includes('mccb') ||
    text.includes('rmu') ||
    text.includes('cầu dao') ||
    text.includes('dao cách ly') ||
    text.includes('switch')
  ) {
    return 'switchgear';
  }
  if (
    text.includes('động cơ') ||
    text.includes('motor') ||
    text.includes('máy phát') ||
    text.includes('generator') ||
    text.includes('tổ máy') ||
    text.includes('bơm') ||
    text.includes('quạt')
  ) {
    return 'motor';
  }
  if (
    text.includes('tiếp địa') ||
    text.includes('nối đất') ||
    text.includes('ground') ||
    text.includes('chống sét') ||
    text.includes('arrester') ||
    text.includes('hệ thống đất')
  ) {
    return 'grounding';
  }
  if (
    text.includes('ắc quy') ||
    text.includes('battery') ||
    text.includes('ups') ||
    text.includes('bess') ||
    text.includes('solar') ||
    text.includes('quang điện') ||
    text.includes('pv') ||
    text.includes('sạc') ||
    text.includes('evse') ||
    text.includes('dc')
  ) {
    return 'dc_power';
  }
  return 'other';
}

interface ProjectCategorizedEquipmentTableProps {
  project: ProjectMasterRecord;
  equipmentList?: ProjectEquipmentItem[];
  equipmentStandardMap: Record<string, 'ats' | 'mts'>;
  onSetStandard: (key: string, std: 'ats' | 'mts') => void;
  onNavigateToNeta: (
    project: ProjectMasterRecord,
    equipment: ProjectEquipmentItem,
    standard: 'ats' | 'mts'
  ) => void;
  onRemoveEquipment?: (kksCode: string) => void;
  onOpenAddEquipmentModal?: () => void;
  readOnlyActions?: boolean;
}

export const ProjectCategorizedEquipmentTable: React.FC<ProjectCategorizedEquipmentTableProps> = ({
  project,
  equipmentList = [],
  equipmentStandardMap,
  onSetStandard,
  onNavigateToNeta,
  onRemoveEquipment,
  onOpenAddEquipmentModal,
  readOnlyActions = false
}) => {
  const [selectedCategory, setSelectedCategory] = useState<EquipmentCategoryKey>('all');
  const [searchQuery, setSearchQuery] = useState('');
  const [copiedKks, setCopiedKks] = useState<string | null>(null);

  // Group & Count
  const categoryCounts = useMemo(() => {
    const counts: Record<EquipmentCategoryKey, number> = {
      all: equipmentList.length,
      transformer: 0,
      cable: 0,
      switchgear: 0,
      motor: 0,
      grounding: 0,
      dc_power: 0,
      other: 0
    };
    equipmentList.forEach((eq) => {
      const cat = classifyEquipment(eq.type, eq.name);
      counts[cat] = (counts[cat] || 0) + 1;
    });
    return counts;
  }, [equipmentList]);

  // Filtered List
  const filteredEquipment = useMemo(() => {
    let list = equipmentList;
    if (selectedCategory !== 'all') {
      list = list.filter((eq) => classifyEquipment(eq.type, eq.name) === selectedCategory);
    }
    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      list = list.filter(
        (eq) =>
          eq.kksCode.toLowerCase().includes(q) ||
          eq.name.toLowerCase().includes(q) ||
          (eq.type && eq.type.toLowerCase().includes(q)) ||
          (eq.location && eq.location.toLowerCase().includes(q))
      );
    }
    return list;
  }, [equipmentList, selectedCategory, searchQuery]);

  const handleCopy = (kks: string) => {
    navigator.clipboard.writeText(kks);
    setCopiedKks(kks);
    setTimeout(() => setCopiedKks(null), 2000);
  };

  return (
    <div className="space-y-3 bg-white rounded-xl border border-slate-200/90 shadow-2xs overflow-hidden">
      {/* Category Tabs & Filter Header */}
      <div className="p-3 bg-gradient-to-r from-slate-50 via-indigo-50/30 to-slate-50 border-b border-slate-200 flex flex-col md:flex-row md:items-center justify-between gap-3">
        {/* Category Pills Strip */}
        <div className="flex items-center gap-1.5 overflow-x-auto pb-1 text-xs scrollbar-thin">
          {(
            [
              'all',
              'transformer',
              'cable',
              'switchgear',
              'motor',
              'grounding',
              'dc_power',
              'other'
            ] as EquipmentCategoryKey[]
          ).map((catKey) => {
            const def = CATEGORY_DEFINITIONS[catKey];
            const count = categoryCounts[catKey] || 0;
            const isActive = selectedCategory === catKey;

            // Don't hide 'all' or active; hide categories with 0 items if we have items
            if (count === 0 && catKey !== 'all' && !isActive) return null;

            return (
              <button
                key={catKey}
                type="button"
                onClick={() => setSelectedCategory(catKey)}
                className={`px-2.5 py-1.5 rounded-lg font-bold text-xs flex items-center gap-1.5 transition-all whitespace-nowrap shrink-0 ${
                  isActive
                    ? 'bg-indigo-600 text-white shadow-xs ring-2 ring-indigo-400'
                    : 'bg-white text-slate-700 hover:bg-slate-100 border border-slate-200'
                }`}
              >
                <span>{def.icon}</span>
                <span>{def.label.split('(')[0].trim()}</span>
                <span
                  className={`px-1.5 py-0.2 rounded-full text-[10px] font-mono font-extrabold ${
                    isActive ? 'bg-white/20 text-white' : 'bg-slate-100 text-slate-700'
                  }`}
                >
                  {count}
                </span>
              </button>
            );
          })}
        </div>

        {/* Quick Search & Add Button */}
        <div className="flex items-center gap-2 shrink-0">
          <div className="relative min-w-[190px]">
            <Search size={13} className="absolute left-2.5 top-2.5 text-slate-400" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Lọc KKS, tên, loại..."
              className="w-full pl-8 pr-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500 font-medium"
            />
          </div>

          {onOpenAddEquipmentModal && (
            <button
              type="button"
              onClick={onOpenAddEquipmentModal}
              className="px-2.5 py-1.5 bg-amber-500 hover:bg-amber-600 text-slate-950 font-bold text-xs rounded-lg flex items-center gap-1 shadow-2xs whitespace-nowrap"
            >
              <Plus size={13} />
              <span>+ Thêm Thiết Bị</span>
            </button>
          )}
        </div>
      </div>

      {/* Main Table View */}
      {filteredEquipment.length > 0 ? (
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs border-collapse">
            <thead className="bg-slate-100/80 text-slate-700 font-bold border-b border-slate-200">
              <tr>
                <th className="py-2.5 px-3 w-12 text-center">STT</th>
                <th className="py-2.5 px-3 w-40">Mã KKS Code</th>
                <th className="py-2.5 px-3 min-w-[200px]">Tên Thiết Bị / Ngăn Lộ</th>
                <th className="py-2.5 px-3 w-44">Phân Loại NETA</th>
                <th className="py-2.5 px-3 w-28 text-center">Trạng Thái PMP</th>
                <th className="py-2.5 px-3 text-center min-w-[260px]">
                  Thử Nghiệm & Nhập Form NETA
                </th>
                {!readOnlyActions && onRemoveEquipment && (
                  <th className="py-2.5 px-3 w-14 text-center">Xóa</th>
                )}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filteredEquipment.map((eq, idx) => {
                const stdKey = `${project.id}_${eq.kksCode}`;
                const currentStd = equipmentStandardMap[stdKey] || eq.standard || 'ats';
                const catKey = classifyEquipment(eq.type, eq.name);
                const catDef = CATEGORY_DEFINITIONS[catKey];
                const isTested = eq.status === 'đã_test';

                return (
                  <tr
                    key={eq.kksCode || idx}
                    className="hover:bg-indigo-50/30 transition-colors group"
                  >
                    {/* STT */}
                    <td className="py-2.5 px-3 text-center font-mono font-bold text-slate-400">
                      {idx + 1}
                    </td>

                    {/* KKS Code */}
                    <td className="py-2.5 px-3 whitespace-nowrap">
                      <div className="flex items-center gap-1.5">
                        <span className="font-mono font-black text-indigo-700 text-xs bg-indigo-50 px-2 py-0.5 rounded border border-indigo-200/60">
                          {eq.kksCode}
                        </span>
                        <button
                          type="button"
                          onClick={() => handleCopy(eq.kksCode)}
                          className="opacity-0 group-hover:opacity-100 text-slate-400 hover:text-indigo-600 transition-opacity p-0.5 rounded"
                          title="Sao chép mã KKS"
                        >
                          {copiedKks === eq.kksCode ? (
                            <Check size={12} className="text-emerald-600" />
                          ) : (
                            <Copy size={12} />
                          )}
                        </button>
                      </div>
                    </td>

                    {/* Equipment Name & Details */}
                    <td className="py-2.5 px-3">
                      <div className="font-bold text-slate-900 text-xs flex items-center gap-1.5">
                        <span>{catDef.icon}</span>
                        <span>{eq.name}</span>
                      </div>
                      <div className="text-[11px] text-slate-500 flex items-center gap-2 mt-0.5">
                        {eq.location && <span>Vị trí: {eq.location}</span>}
                        {eq.rating && (
                          <>
                            <span>•</span>
                            <span className="font-mono text-slate-600">{eq.rating}</span>
                          </>
                        )}
                      </div>
                    </td>

                    {/* Category Badge */}
                    <td className="py-2.5 px-3 whitespace-nowrap">
                      <span
                        className={`inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-bold border ${catDef.badgeBg} ${catDef.badgeText} ${catDef.border}`}
                      >
                        <span>{catDef.icon}</span>
                        <span>{eq.type || catDef.label.split('(')[0].trim()}</span>
                      </span>
                    </td>

                    {/* PMP Status */}
                    <td className="py-2.5 px-3 text-center whitespace-nowrap">
                      <span
                        className={`inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[10px] font-bold ${
                          isTested
                            ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                            : 'bg-amber-100 text-amber-900 border border-amber-300'
                        }`}
                      >
                        {isTested ? <CheckCircle2 size={11} /> : <Clock size={11} />}
                        <span>{isTested ? 'Đã test' : 'Chờ test'}</span>
                      </span>
                    </td>

                    {/* Standard Selector & Direct Action */}
                    <td className="py-2.5 px-3 text-center whitespace-nowrap">
                      <div className="flex items-center justify-center gap-2">
                        {/* ATS / MST Toggle Button */}
                        <div className="flex items-center bg-white p-0.5 rounded-lg border border-slate-300 text-[10px] font-bold shrink-0 shadow-2xs">
                          <button
                            type="button"
                            onClick={() => onSetStandard(stdKey, 'ats')}
                            className={`px-2 py-1 rounded transition-all ${
                              currentStd === 'ats'
                                ? 'bg-blue-600 text-white font-black shadow-xs'
                                : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
                            }`}
                            title="ANSI/NETA ATS-2025: Thí nghiệm nghiệm thu đóng điện"
                          >
                            ATS
                          </button>
                          <button
                            type="button"
                            onClick={() => onSetStandard(stdKey, 'mts')}
                            className={`px-2 py-1 rounded transition-all ${
                              currentStd === 'mts'
                                ? 'bg-amber-500 text-slate-950 font-black shadow-xs'
                                : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
                            }`}
                            title="ANSI/NETA MTS-2023 (MST): Thí nghiệm bảo dưỡng định kỳ"
                          >
                            MST
                          </button>
                        </div>

                        {/* Direct Form Entry Button */}
                        <button
                          type="button"
                          onClick={() => onNavigateToNeta(project, eq, currentStd)}
                          className={`px-3 py-1.5 font-black text-xs rounded-lg transition-all flex items-center gap-1.5 shadow-xs active:scale-95 whitespace-nowrap ${
                            currentStd === 'mts'
                              ? 'bg-gradient-to-r from-amber-500 to-amber-600 hover:from-amber-600 hover:to-amber-700 text-slate-950 border border-amber-400/50'
                              : 'bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white'
                          }`}
                          title={`Mở biểu mẫu NETA ${currentStd === 'mts' ? 'MST' : 'ATS'} cho ${eq.name}`}
                        >
                          <Play
                            size={11}
                            className={currentStd === 'mts' ? 'fill-slate-950' : 'fill-white'}
                          />
                          <span>
                            Nhập Form NETA ({currentStd === 'mts' ? 'MST' : 'ATS'})
                          </span>
                        </button>
                      </div>
                    </td>

                    {/* Delete Action (if enabled) */}
                    {!readOnlyActions && onRemoveEquipment && (
                      <td className="py-2.5 px-3 text-center">
                        <button
                          type="button"
                          onClick={() => onRemoveEquipment(eq.kksCode)}
                          className="text-rose-600 hover:text-rose-800 p-1 hover:bg-rose-50 rounded transition-colors"
                          title="Xóa thiết bị khỏi dự án"
                        >
                          <Trash2 size={13} />
                        </button>
                      </td>
                    )}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      ) : (
        <div className="p-6 text-center text-slate-500 space-y-2">
          <Layers size={28} className="mx-auto text-slate-300" />
          <p className="text-xs font-semibold">
            {searchQuery
              ? `Không tìm thấy thiết bị nào khớp với từ khóa "${searchQuery}".`
              : selectedCategory !== 'all'
              ? `Chưa có thiết bị nào trong nhóm "${CATEGORY_DEFINITIONS[selectedCategory].label}".`
              : 'Dự án này chưa có thiết bị nào được soạn.'}
          </p>
          {onOpenAddEquipmentModal && (
            <button
              type="button"
              onClick={onOpenAddEquipmentModal}
              className="text-xs font-bold text-indigo-600 hover:underline inline-flex items-center gap-1"
            >
              <Plus size={13} /> Thêm thiết bị KKS ngay
            </button>
          )}
        </div>
      )}

      {/* Summary Footer */}
      <div className="p-2.5 bg-slate-50 border-t border-slate-200/80 flex items-center justify-between text-[11px] text-slate-500 px-4">
        <span>
          Hiển thị <strong>{filteredEquipment.length}</strong> / <strong>{equipmentList.length}</strong> thiết bị KKS thuộc dự án <strong>{project.name}</strong>
        </span>
        <span className="flex items-center gap-2">
          <span className="inline-flex items-center gap-1 text-emerald-700 font-bold">
            <CheckCircle2 size={11} /> {equipmentList.filter((e) => e.status === 'đã_test').length} Đã test
          </span>
          <span>•</span>
          <span className="inline-flex items-center gap-1 text-amber-700 font-bold">
            <Clock size={11} /> {equipmentList.filter((e) => e.status !== 'đã_test').length} Chờ test
          </span>
        </span>
      </div>
    </div>
  );
};
