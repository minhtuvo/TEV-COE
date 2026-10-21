import React, { useState, useEffect, useRef } from 'react';
import {
  UserCheck,
  ChevronDown,
  UserPlus,
  Mail,
  ShieldCheck,
  Search,
  Check,
  X,
  Phone,
  Building2
} from 'lucide-react';

export interface FseEngineer {
  id: string;
  name: string;
  email: string;
  role: string;
  phone?: string;
  company?: string;
}

export const DEFAULT_TEV_FSE_LIST: FseEngineer[] = [
  {
    id: 'fse-vo-minh-tu',
    name: 'Võ Minh Tú',
    email: 'sgm1707@gmail.com',
    role: 'Trưởng nhóm FSE / Chuyên gia NETA',
    phone: '0908 123 456',
    company: 'TEV Field Service'
  },
  {
    id: 'fse-ks-nguyen-van-s',
    name: 'KS. Nguyễn Văn S',
    email: 'nguyenvans@tev.vn',
    role: 'Kỹ sư Bảo trì Cơ điện Nhà máy',
    phone: '0918 234 567',
    company: 'TEV Plant Service'
  },
  {
    id: 'fse-ks-tran-van-h',
    name: 'KS. Trần Văn H',
    email: 'tranvanh@tev.vn',
    role: 'Kỹ sư Trưởng Vận hành & Bảo trì',
    phone: '0919 345 678',
    company: 'TEV Plant Service'
  },
  {
    id: 'fse-ks-le-hoang-d',
    name: 'KS. Lê Hoàng D',
    email: 'lehoangd@tev.vn',
    role: 'Kỹ sư Tự động hóa & SCADA',
    phone: '0917 456 789',
    company: 'TEV Plant Service'
  },
  {
    id: 'fse-ks-pham-minh-t',
    name: 'KS. Phạm Minh T',
    email: 'phamminht@tev.vn',
    role: 'Kỹ sư Thiết bị Động lực & MBA',
    phone: '0916 567 890',
    company: 'TEV Plant Service'
  },
  {
    id: 'fse-nguyen-van-long',
    name: 'Nguyễn Văn Long',
    email: 'long.nv@tev.vn',
    role: 'Kỹ sư Cao thế NETA Level 3',
    phone: '0912 345 678',
    company: 'TEV Field Service'
  },
  {
    id: 'fse-tran-van-binh',
    name: 'Trần Văn Bình',
    email: 'binh.tv@tev.vn',
    role: 'Kỹ sư Thí nghiệm Relay & MBA',
    phone: '0987 654 321',
    company: 'TEV Field Service'
  },
  {
    id: 'fse-le-hoang-nam',
    name: 'Lê Hoàng Nam',
    email: 'nam.lh@tev.vn',
    role: 'Kỹ sư Chẩn đoán DGA & Phóng điện PD',
    phone: '0934 567 890',
    company: 'TEV Field Service'
  },
  {
    id: 'fse-pham-quoc-dung',
    name: 'Phạm Quốc Dũng',
    email: 'dung.pq@tev.vn',
    role: 'Kỹ sư Đo nhiệt Hồng ngoại NETA L2',
    phone: '0945 678 901',
    company: 'TEV Field Service'
  },
  {
    id: 'fse-dang-minh-quan',
    name: 'Đặng Minh Quân',
    email: 'quan.dm@tev.vn',
    role: 'Kỹ sư Thí nghiệm Cáp & Tiếp địa',
    phone: '0923 456 789',
    company: 'TEV Field Service'
  },
  {
    id: 'fse-ks-le-minh-tuan',
    name: 'KS. Lê Minh Tuấn',
    email: 'tuan.lm@tev.vn',
    role: 'Kỹ sư NETA Thí nghiệm Điện',
    phone: '0903 112 233',
    company: 'TEV Field Service'
  },
  {
    id: 'fse-ks-pham-quoc-hung',
    name: 'KS. Phạm Quốc Hưng',
    email: 'hung.pq@tev.vn',
    role: 'Kỹ sư Rơ-le & Bảo vệ Hệ thống',
    phone: '0904 223 344',
    company: 'TEV Field Service'
  },
  {
    id: 'fse-ks-tran-hoang-long',
    name: 'KS. Trần Hoàng Long',
    email: 'long.th@tev.vn',
    role: 'Kỹ sư Hệ thống Điện & Máy phát',
    phone: '0905 334 455',
    company: 'TEV Field Service'
  },
  {
    id: 'fse-ks-dang-van-nam',
    name: 'KS. Đặng Văn Nam',
    email: 'nam.dv@tev.vn',
    role: 'Kỹ sư Khí cụ & Máy cắt Cao thế',
    phone: '0906 445 566',
    company: 'TEV Field Service'
  }
];

const STORAGE_KEY = 'tev_registered_fse_engineers_v3';
export const EVT_FSE_UPDATED = 'tev_fse_updated';

interface Props {
  value: string;
  onChange: (fseName: string, engineer?: FseEngineer) => void;
  className?: string;
  label?: string;
  required?: boolean;
}

export const FseEngineerDropdown: React.FC<Props> = ({
  value,
  onChange,
  className = '',
  label,
  required = false
}) => {
  const [engineers, setEngineers] = useState<FseEngineer[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          // Merge default list to ensure plant engineers exist
          const existingIds = new Set(parsed.map((p: FseEngineer) => p.name.toLowerCase()));
          const missingDefaults = DEFAULT_TEV_FSE_LIST.filter(
            d => !existingIds.has(d.name.toLowerCase())
          );
          return [...parsed, ...missingDefaults];
        }
      }
    } catch {
      // Fallback
    }
    return DEFAULT_TEV_FSE_LIST;
  });

  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState('');
  const [quickNewName, setQuickNewName] = useState('');
  const [showAddModal, setShowAddModal] = useState(false);
  const [newEngineer, setNewEngineer] = useState({
    name: '',
    email: '',
    role: 'Kỹ sư Thí nghiệm Hiện trường NETA',
    phone: ''
  });

  const dropdownRef = useRef<HTMLDivElement>(null);

  // Sync from external events
  useEffect(() => {
    const handleFseUpdate = (e: any) => {
      if (Array.isArray(e.detail) && e.detail.length > 0) {
        setEngineers(e.detail);
      }
    };
    window.addEventListener(EVT_FSE_UPDATED, handleFseUpdate);
    return () => window.removeEventListener(EVT_FSE_UPDATED, handleFseUpdate);
  }, []);

  // Sync to localStorage and broadcast
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(engineers));
      window.dispatchEvent(new CustomEvent(EVT_FSE_UPDATED, { detail: engineers }));
    } catch {
      // ignore
    }
  }, [engineers]);

  // Close when clicking outside
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Determine current active engineer
  const matchedEngineer = engineers.find(
    (e) =>
      value &&
      (value.includes(e.name) ||
        value.includes(e.email) ||
        e.name.toLowerCase() === value.toLowerCase() ||
        value.toLowerCase().includes(e.name.toLowerCase()))
  );

  const activeDisplayName = matchedEngineer 
    ? matchedEngineer.name 
    : (value && value.trim() ? value.trim() : (engineers[0]?.name || 'Võ Minh Tú'));
  
  const activeDisplaySub = matchedEngineer 
    ? matchedEngineer.email 
    : (value && value.trim() ? (matchedEngineer ? matchedEngineer.role : 'Người thực hiện / PIC') : (engineers[0]?.email || 'sgm1707@gmail.com'));

  const handleSelect = (eng: FseEngineer) => {
    // Send clean name to maintain consistency in Work Orders & Google Sheets
    onChange(eng.name, eng);
    setIsOpen(false);
    setSearch('');
  };

  const handleQuickAdd = () => {
    const nameToAdd = quickNewName.trim() || search.trim();
    if (!nameToAdd) return;

    // Check if already exists
    const existing = engineers.find(e => e.name.toLowerCase() === nameToAdd.toLowerCase());
    if (existing) {
      handleSelect(existing);
      setQuickNewName('');
      return;
    }

    const created: FseEngineer = {
      id: `fse-${Date.now()}`,
      name: nameToAdd,
      email: `${nameToAdd.toLowerCase().replace(/[^a-z0-9]/g, '') || 'fse'}@tev.vn`,
      role: 'Kỹ sư Hiện trường / FSE',
      phone: '',
      company: 'TEV Service'
    };

    const updated = [created, ...engineers];
    setEngineers(updated);
    handleSelect(created);
    setQuickNewName('');
    setSearch('');
  };

  const handleAddNew = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newEngineer.name.trim()) return;

    const created: FseEngineer = {
      id: `fse-${Date.now()}`,
      name: newEngineer.name.trim(),
      email: newEngineer.email.trim() || `${newEngineer.name.toLowerCase().replace(/[^a-z0-9]/g, '')}@tev.vn`,
      role: newEngineer.role.trim() || 'Kỹ sư Hiện trường / FSE',
      phone: newEngineer.phone.trim(),
      company: 'TEV Field Service'
    };

    const updated = [created, ...engineers];
    setEngineers(updated);
    handleSelect(created);
    setShowAddModal(false);
    setNewEngineer({
      name: '',
      email: '',
      role: 'Kỹ sư Thí nghiệm Hiện trường NETA',
      phone: ''
    });
  };

  const filtered = engineers.filter(
    (e) =>
      e.name.toLowerCase().includes(search.toLowerCase()) ||
      e.email.toLowerCase().includes(search.toLowerCase()) ||
      e.role.toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div className={`relative ${className}`} ref={dropdownRef}>
      {label && (
        <label className="block text-xs font-semibold text-slate-700 mb-1 flex items-center justify-between">
          <span className="flex items-center gap-1.5">
            <UserCheck size={13} className="text-amber-600" />
            {label} {required && <span className="text-red-500">*</span>}
          </span>
          <span className="text-[10px] font-mono text-amber-700 bg-amber-50 px-1.5 py-0.2 rounded border border-amber-200">
            Sheet FSE
          </span>
        </label>
      )}

      {/* Trigger Button */}
      <button
        type="button"
        onClick={() => setIsOpen(!isOpen)}
        className="w-full text-left bg-white border border-slate-300 hover:border-amber-400 rounded-lg px-2.5 py-1.5 focus:ring-2 focus:ring-amber-500 focus:outline-none transition-all flex items-center justify-between shadow-2xs group"
      >
        <div className="flex items-center gap-2 min-w-0 pr-1">
          <div className="w-6 h-6 rounded-full bg-gradient-to-tr from-amber-600 to-yellow-500 text-white flex items-center justify-center text-[10px] font-bold shrink-0 shadow-2xs">
            {activeDisplayName ? activeDisplayName.slice(0, 2).toUpperCase() : 'PIC'}
          </div>
          <div className="truncate">
            <div className="text-xs font-bold text-slate-800 flex items-center gap-1.5 truncate">
              <span>{activeDisplayName}</span>
              <span className="text-[10px] font-normal text-amber-700 bg-amber-50 border border-amber-200/80 px-1.5 py-0.2 rounded shrink-0">
                PIC
              </span>
            </div>
            <div className="text-[10px] text-slate-500 font-mono truncate">
              {activeDisplaySub}
            </div>
          </div>
        </div>
        <ChevronDown
          size={14}
          className={`text-slate-400 group-hover:text-amber-600 shrink-0 transition-transform ${
            isOpen ? 'rotate-180' : ''
          }`}
        />
      </button>

      {/* Floating Dropdown */}
      {isOpen && (
        <div className="absolute left-0 right-0 mt-1.5 bg-white border border-slate-200 rounded-xl shadow-2xl z-50 overflow-hidden animate-in fade-in-50 zoom-in-95 duration-100 min-w-[300px]">
          {/* Header & Search */}
          <div className="p-2.5 bg-slate-50 border-b border-slate-200 space-y-2">
            <div className="flex items-center justify-between text-xs font-bold text-slate-700">
              <span className="flex items-center gap-1.5">
                <UserCheck size={14} className="text-amber-600" />
                Người Thực Hiện / PIC ({engineers.length})
              </span>
              <button
                type="button"
                onClick={() => {
                  setShowAddModal(true);
                  setIsOpen(false);
                }}
                className="text-[11px] font-bold text-amber-700 hover:text-amber-800 flex items-center gap-1 hover:underline"
              >
                <UserPlus size={12} /> Thêm FSE Đầy Đủ
              </button>
            </div>
            <div className="relative">
              <Search size={12} className="absolute left-2.5 top-2.5 text-slate-400" />
              <input
                type="text"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Tìm tên kỹ sư, mã hoặc chức danh..."
                className="w-full text-xs pl-8 pr-2.5 py-1.5 border border-slate-200 rounded-lg focus:outline-none focus:ring-1 focus:ring-amber-500 bg-white"
                autoFocus
              />
            </div>

            {/* Quick Add Person Inline */}
            <div className="flex items-center gap-1.5 pt-1">
              <input
                type="text"
                value={quickNewName}
                onChange={(e) => setQuickNewName(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    handleQuickAdd();
                  }
                }}
                placeholder="+ Nhập tên người mới nhanh..."
                className="flex-1 text-[11px] px-2 py-1 border border-dashed border-amber-300 rounded-md bg-amber-50/50 focus:bg-white focus:border-amber-500 focus:outline-none text-slate-800 placeholder:text-amber-700/60 font-medium"
              />
              <button
                type="button"
                onClick={handleQuickAdd}
                disabled={!quickNewName.trim() && !search.trim()}
                className="px-2.5 py-1 bg-amber-600 hover:bg-amber-700 disabled:opacity-40 text-white font-bold text-[11px] rounded-md transition-colors shrink-0 flex items-center gap-1"
                title="Thêm và chọn ngay người này"
              >
                <UserPlus size={12} /> Thêm
              </button>
            </div>
          </div>

          {/* List items */}
          <div className="max-h-60 overflow-y-auto p-1 divide-y divide-slate-100">
            {filtered.length === 0 ? (
              <div className="p-4 text-center text-xs space-y-2">
                <p className="text-slate-400">Không tìm thấy kỹ sư "{search}"</p>
                {search.trim() && (
                  <button
                    type="button"
                    onClick={handleQuickAdd}
                    className="w-full py-1.5 px-3 bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 rounded-lg font-bold text-xs flex items-center justify-center gap-1.5 transition-colors"
                  >
                    <UserPlus size={13} /> Thêm "{search.trim()}" làm người thực hiện mới
                  </button>
                )}
              </div>
            ) : (
              filtered.map((eng) => {
                const isSelected = 
                  activeDisplayName?.toLowerCase() === eng.name.toLowerCase() ||
                  matchedEngineer?.id === eng.id ||
                  (value && value.toLowerCase().includes(eng.name.toLowerCase()));

                return (
                  <button
                    key={eng.id}
                    type="button"
                    onClick={() => handleSelect(eng)}
                    className={`w-full text-left p-2 rounded-lg flex items-center justify-between text-xs transition-colors ${
                      isSelected
                        ? 'bg-amber-50 text-amber-950 font-semibold'
                        : 'hover:bg-slate-50 text-slate-700'
                    }`}
                  >
                    <div className="flex items-center gap-2.5 min-w-0">
                      <div
                        className={`w-7 h-7 rounded-full flex items-center justify-center text-[10px] font-black shrink-0 ${
                          isSelected
                            ? 'bg-amber-600 text-white'
                            : 'bg-slate-100 text-slate-600'
                        }`}
                      >
                        {eng.name.slice(0, 2).toUpperCase()}
                      </div>
                      <div className="truncate">
                        <div className="font-bold text-slate-900 flex items-center gap-1.5 truncate">
                          <span>{eng.name}</span>
                          {eng.email === 'sgm1707@gmail.com' && (
                            <span className="text-[9px] bg-amber-500 text-slate-950 font-bold px-1.5 py-0.2 rounded">
                              Mặc định
                            </span>
                          )}
                        </div>
                        <div className="text-[10px] text-slate-500 font-mono truncate">
                          {eng.email}
                        </div>
                        <div className="text-[10px] text-amber-700 truncate">
                          {eng.role}
                        </div>
                      </div>
                    </div>
                    {isSelected && (
                      <Check size={16} className="text-amber-600 shrink-0 ml-2" />
                    )}
                  </button>
                );
              })
            )}
          </div>

          {/* Bottom Action */}
          <div className="p-2 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-[11px]">
            <span className="text-slate-500 font-medium">
              Đang chọn: <strong className="text-slate-800">{activeDisplayName}</strong>
            </span>
            <button
              type="button"
              onClick={() => {
                setShowAddModal(true);
                setIsOpen(false);
              }}
              className="text-amber-700 font-bold hover:underline flex items-center gap-1"
            >
              <UserPlus size={12} /> Thêm đầy đủ
            </button>
          </div>
        </div>
      )}

      {/* Modal Add New FSE Engineer */}
      {showAddModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-md w-full p-5 border border-slate-200 shadow-2xl space-y-4 animate-in fade-in-50 zoom-in-95">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div className="flex items-center gap-2">
                <span className="p-2 bg-amber-100 text-amber-700 rounded-lg">
                  <UserPlus size={18} />
                </span>
                <div>
                  <h3 className="text-sm font-bold text-slate-900">
                    Thêm Kỹ Sư FSE Vào Danh Sách TEV
                  </h3>
                  <p className="text-[11px] text-slate-500">
                    Kỹ sư mới sẽ được lưu vào danh sách chọn nhanh ngoài hiện trường
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setShowAddModal(false)}
                className="text-slate-400 hover:text-slate-600 p-1"
              >
                <X size={18} />
              </button>
            </div>

            <form onSubmit={handleAddNew} className="space-y-3 text-xs">
              <div>
                <label className="block font-semibold text-slate-700 mb-1">
                  Họ và tên Kỹ sư FSE *
                </label>
                <input
                  type="text"
                  required
                  placeholder="Ví dụ: Hoàng Văn D"
                  value={newEngineer.name}
                  onChange={(e) =>
                    setNewEngineer({ ...newEngineer, name: e.target.value })
                  }
                  className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-amber-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">
                  Địa chỉ Email *
                </label>
                <input
                  type="email"
                  required
                  placeholder="email.fse@tev.vn"
                  value={newEngineer.email}
                  onChange={(e) =>
                    setNewEngineer({ ...newEngineer, email: e.target.value })
                  }
                  className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-amber-500 focus:outline-none font-mono"
                />
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">
                  Chức danh / Chứng chỉ chuyên môn
                </label>
                <input
                  type="text"
                  placeholder="Ví dụ: Kỹ sư Cao thế NETA L2 / Chẩn đoán MBA"
                  value={newEngineer.role}
                  onChange={(e) =>
                    setNewEngineer({ ...newEngineer, role: e.target.value })
                  }
                  className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-amber-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">
                  Số điện thoại liên hệ
                </label>
                <input
                  type="tel"
                  placeholder="09xx xxx xxx"
                  value={newEngineer.phone}
                  onChange={(e) =>
                    setNewEngineer({ ...newEngineer, phone: e.target.value })
                  }
                  className="w-full border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-amber-500 focus:outline-none"
                />
              </div>

              <div className="pt-2 flex items-center justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setShowAddModal(false)}
                  className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg font-semibold"
                >
                  Hủy
                </button>
                <button
                  type="submit"
                  className="px-4 py-1.5 bg-amber-600 hover:bg-amber-700 text-white rounded-lg font-bold shadow-xs flex items-center gap-1.5"
                >
                  <Check size={14} /> Thêm Vào Danh Sách
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};
