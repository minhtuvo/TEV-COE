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
  }
];

const STORAGE_KEY = 'tev_registered_fse_engineers_v2';

interface Props {
  value: string;
  onChange: (fseName: string, engineer?: FseEngineer) => void;
  className?: string;
}

export const FseEngineerDropdown: React.FC<Props> = ({
  value,
  onChange,
  className = ''
}) => {
  const [engineers, setEngineers] = useState<FseEngineer[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          // Ensure Võ Minh Tú is always first
          const tu = parsed.find((e: FseEngineer) => e.email === 'sgm1707@gmail.com') || DEFAULT_TEV_FSE_LIST[0];
          const others = parsed.filter((e: FseEngineer) => e.email !== 'sgm1707@gmail.com');
          return [tu, ...others];
        }
      }
    } catch {
      // Fallback
    }
    return DEFAULT_TEV_FSE_LIST;
  });

  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState('');
  const [showAddModal, setShowAddModal] = useState(false);
  const [newEngineer, setNewEngineer] = useState({
    name: '',
    email: '',
    role: 'Kỹ sư Thí nghiệm Hiện trường NETA',
    phone: ''
  });

  const dropdownRef = useRef<HTMLDivElement>(null);

  // Sync to localStorage
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(engineers));
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
  const activeEngineer = engineers.find(
    (e) =>
      value?.includes(e.name) ||
      value?.includes(e.email) ||
      e.name.toLowerCase() === value?.toLowerCase()
  ) || engineers[0]; // Default is Vo Minh Tu

  const handleSelect = (eng: FseEngineer) => {
    // Standard format for FSE: "Võ Minh Tú (sgm1707@gmail.com)"
    const formatted = `${eng.name} (${eng.email})`;
    onChange(formatted, eng);
    setIsOpen(false);
  };

  const handleAddNew = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newEngineer.name.trim() || !newEngineer.email.trim()) return;

    const created: FseEngineer = {
      id: `fse-${Date.now()}`,
      name: newEngineer.name.trim(),
      email: newEngineer.email.trim(),
      role: newEngineer.role.trim() || 'Kỹ sư FSE',
      phone: newEngineer.phone.trim(),
      company: 'TEV Field Service'
    };

    const updated = [...engineers, created];
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
      {/* Trigger Button */}
      <button
        type="button"
        onClick={() => setIsOpen(!isOpen)}
        className="w-full text-left bg-white border border-slate-300 hover:border-amber-400 rounded-lg px-2.5 py-1.5 focus:ring-2 focus:ring-amber-500 focus:outline-none transition-all flex items-center justify-between shadow-2xs group"
      >
        <div className="flex items-center gap-2 min-w-0 pr-1">
          <div className="w-6 h-6 rounded-full bg-gradient-to-tr from-amber-600 to-yellow-500 text-white flex items-center justify-center text-[10px] font-bold shrink-0 shadow-2xs">
            {activeEngineer?.name?.slice(0, 2).toUpperCase() || 'VT'}
          </div>
          <div className="truncate">
            <div className="text-xs font-bold text-slate-800 flex items-center gap-1.5 truncate">
              <span>{activeEngineer?.name || 'Võ Minh Tú'}</span>
              <span className="text-[10px] font-normal text-amber-700 bg-amber-50 border border-amber-200/80 px-1.5 py-0.2 rounded shrink-0">
                FSE TEV
              </span>
            </div>
            <div className="text-[10px] text-slate-500 font-mono truncate">
              {activeEngineer?.email || 'sgm1707@gmail.com'}
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
        <div className="absolute left-0 right-0 mt-1.5 bg-white border border-slate-200 rounded-xl shadow-2xl z-50 overflow-hidden animate-in fade-in-50 zoom-in-95 duration-100 min-w-[280px]">
          {/* Header & Search */}
          <div className="p-2.5 bg-slate-50 border-b border-slate-200 space-y-2">
            <div className="flex items-center justify-between text-xs font-bold text-slate-700">
              <span className="flex items-center gap-1.5">
                <UserCheck size={14} className="text-amber-600" />
                Danh Sách Kỹ Sư FSE TEV ({engineers.length})
              </span>
              <button
                type="button"
                onClick={() => {
                  setShowAddModal(true);
                  setIsOpen(false);
                }}
                className="text-[11px] font-bold text-amber-700 hover:text-amber-800 flex items-center gap-1 hover:underline"
              >
                <UserPlus size={12} /> Thêm FSE
              </button>
            </div>
            <div className="relative">
              <Search size={12} className="absolute left-2.5 top-2.5 text-slate-400" />
              <input
                type="text"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Tìm tên, email kỹ sư FSE..."
                className="w-full text-xs pl-8 pr-2.5 py-1.5 border border-slate-200 rounded-lg focus:outline-none focus:ring-1 focus:ring-amber-500 bg-white"
                autoFocus
              />
            </div>
          </div>

          {/* List items */}
          <div className="max-h-56 overflow-y-auto p-1 divide-y divide-slate-100">
            {filtered.length === 0 ? (
              <div className="p-4 text-center text-xs text-slate-400">
                Không tìm thấy kỹ sư phù hợp.
              </div>
            ) : (
              filtered.map((eng) => {
                const isSelected = activeEngineer?.id === eng.id;
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
              Mặc định: <strong className="text-slate-800">Võ Minh Tú</strong>
            </span>
            <button
              type="button"
              onClick={() => {
                setShowAddModal(true);
                setIsOpen(false);
              }}
              className="text-amber-700 font-bold hover:underline flex items-center gap-1"
            >
              <UserPlus size={12} /> Thêm mới
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
