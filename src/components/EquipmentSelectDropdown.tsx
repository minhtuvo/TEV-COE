import React, { useState, useEffect, useRef } from 'react';
import {
  Cpu,
  ChevronDown,
  Plus,
  Search,
  Check,
  X,
  MapPin,
  Tag,
  Layers,
  FileSpreadsheet,
  Sparkles,
  Zap
} from 'lucide-react';
import { EquipmentMasterRecord, EquipmentCategoryType } from '../utils/fieldSyncService';
import {
  getRegisteredEquipment,
  saveEquipmentRecord,
  EVT_EQUIPMENT_UPDATED
} from '../services/masterDirectoryService';

interface EquipmentSelectDropdownProps {
  value: string;
  onChange: (equipmentValue: string, equipment?: EquipmentMasterRecord) => void;
  selectedCustomer?: string;
  filterType?: EquipmentCategoryType | string;
  displayMode?: 'id' | 'name' | 'both';
  placeholder?: string;
  className?: string;
  label?: string;
  required?: boolean;
}

const EQUIPMENT_CATEGORIES: EquipmentCategoryType[] = [
  'Máy biến áp',
  'Máy cắt',
  'Động cơ',
  'Máy phát',
  'Tủ điện',
  'Cáp điện',
  'Pin & UPS',
  'Rơ le bảo vệ',
  'Inverter'
];

export const EquipmentSelectDropdown: React.FC<EquipmentSelectDropdownProps> = ({
  value,
  onChange,
  selectedCustomer,
  filterType,
  displayMode = 'id',
  placeholder = 'Chọn hoặc thêm Mã / Tên Thiết Bị...',
  className = '',
  label,
  required = false
}) => {
  const [equipmentList, setEquipmentList] = useState<EquipmentMasterRecord[]>(() =>
    getRegisteredEquipment()
  );
  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState('');
  const [showAllCustomerEq, setShowAllCustomerEq] = useState(false);
  const [showAddModal, setShowAddModal] = useState(false);

  // New equipment form state
  const [newEq, setNewEq] = useState({
    id: '',
    name: '',
    type: (filterType as EquipmentCategoryType) || 'Máy biến áp',
    customer: selectedCustomer || 'Công ty Cổ phần Năng lượng TEV',
    factory: 'Nhà máy chính',
    location: '',
    manufacturer: '',
    model: '',
    yearOfManufacture: new Date().getFullYear().toString(),
    notes: ''
  });

  const dropdownRef = useRef<HTMLDivElement>(null);

  // Listen for updates across app
  useEffect(() => {
    const handleUpdate = (e: any) => {
      if (Array.isArray(e.detail)) {
        setEquipmentList(e.detail);
      } else {
        setEquipmentList(getRegisteredEquipment());
      }
    };
    window.addEventListener(EVT_EQUIPMENT_UPDATED, handleUpdate);
    return () => window.removeEventListener(EVT_EQUIPMENT_UPDATED, handleUpdate);
  }, []);

  // Sync default customer into newEq when selectedCustomer prop changes
  useEffect(() => {
    if (selectedCustomer) {
      setNewEq((prev) => ({ ...prev, customer: selectedCustomer }));
    }
  }, [selectedCustomer]);

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

  const cleanVal = (value || '').trim();
  const activeEquipment = equipmentList.find(
    (e) =>
      e.id.toLowerCase() === cleanVal.toLowerCase() ||
      e.name.toLowerCase() === cleanVal.toLowerCase() ||
      cleanVal.toLowerCase().includes(e.id.toLowerCase())
  );

  const handleSelect = (eq: EquipmentMasterRecord) => {
    let result = eq.id;
    if (displayMode === 'name') result = eq.name;
    else if (displayMode === 'both') result = `${eq.id} - ${eq.name}`;
    onChange(result, eq);
    setIsOpen(false);
    setSearch('');
  };

  const handleAddNewSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const name = newEq.name.trim();
    if (!name) return;

    let id = newEq.id.trim();
    if (!id) {
      // Auto-generate Tag ID based on type
      let prefix = 'EQ';
      if (newEq.type === 'Máy biến áp') prefix = 'TR';
      else if (newEq.type === 'Máy cắt') prefix = 'VCB';
      else if (newEq.type === 'Động cơ') prefix = 'MTR';
      else if (newEq.type === 'Máy phát') prefix = 'GEN';
      else if (newEq.type === 'Tủ điện') prefix = 'SWG';
      else if (newEq.type === 'Cáp điện') prefix = 'CAB';
      else if (newEq.type === 'Pin & UPS') prefix = 'BAT';
      else if (newEq.type === 'Rơ le bảo vệ') prefix = 'REL';
      id = `${prefix}-${Math.floor(100 + Math.random() * 900)}`;
    }

    const created: EquipmentMasterRecord = {
      id,
      name,
      type: newEq.type as EquipmentCategoryType,
      customer: newEq.customer.trim() || selectedCustomer || 'Công ty Cổ phần Năng lượng TEV',
      factory: newEq.factory.trim() || 'Nhà máy chính',
      location: newEq.location.trim() || 'Trạm phân phối',
      manufacturer: newEq.manufacturer.trim(),
      model: newEq.model.trim(),
      yearOfManufacture: newEq.yearOfManufacture.trim(),
      criticality: 'B',
      status: 'healthy',
      health: 95,
      lastCheck: new Date().toLocaleDateString('vi-VN'),
      pmCycleMonths: 6,
      notes: newEq.notes.trim(),
      source: 'app'
    };

    const updatedList = saveEquipmentRecord(created);
    setEquipmentList(updatedList);
    handleSelect(created);
    setShowAddModal(false);
    setNewEq({
      id: '',
      name: '',
      type: (filterType as EquipmentCategoryType) || 'Máy biến áp',
      customer: selectedCustomer || 'Công ty Cổ phần Năng lượng TEV',
      factory: 'Nhà máy chính',
      location: '',
      manufacturer: '',
      model: '',
      yearOfManufacture: new Date().getFullYear().toString(),
      notes: ''
    });
  };

  // Filter items
  const filtered = equipmentList.filter((e) => {
    // Customer filter if applicable and not overridden
    if (selectedCustomer && !showAllCustomerEq) {
      const matchCustomer =
        e.customer.toLowerCase().includes(selectedCustomer.toLowerCase()) ||
        selectedCustomer.toLowerCase().includes(e.customer.toLowerCase());
      if (!matchCustomer) return false;
    }

    // Type filter if applicable
    if (filterType && e.type !== filterType) {
      // Allow general match
    }

    const s = search.toLowerCase();
    return (
      e.id.toLowerCase().includes(s) ||
      e.name.toLowerCase().includes(s) ||
      e.type.toLowerCase().includes(s) ||
      (e.location && e.location.toLowerCase().includes(s)) ||
      (e.customer && e.customer.toLowerCase().includes(s))
    );
  });

  return (
    <div className={`relative ${className}`} ref={dropdownRef}>
      {label && (
        <label className="block text-xs font-semibold text-slate-700 mb-1 flex items-center justify-between">
          <span className="flex items-center gap-1.5">
            <Cpu size={13} className="text-cyan-600" />
            {label} {required && <span className="text-red-500">*</span>}
          </span>
          <span className="text-[10px] font-mono text-cyan-700 bg-cyan-50 px-1.5 py-0.2 rounded border border-cyan-200">
            Sheet thietbi
          </span>
        </label>
      )}

      {/* Trigger Button */}
      <div className="relative">
        <button
          type="button"
          onClick={() => setIsOpen(!isOpen)}
          className={`w-full text-left bg-white border rounded-lg px-3 py-2 text-xs font-medium focus:ring-2 focus:ring-cyan-500 focus:outline-none transition-all flex items-center justify-between shadow-2xs group ${
            activeEquipment ? 'border-cyan-300 bg-cyan-50/20' : 'border-slate-300 hover:border-cyan-400'
          }`}
        >
          <div className="flex items-center gap-2 min-w-0 pr-1">
            <div className="w-5 h-5 rounded-md bg-cyan-100 text-cyan-800 flex items-center justify-center text-[10px] font-bold shrink-0">
              <Zap size={11} />
            </div>
            <div className="truncate">
              {activeEquipment ? (
                <div className="flex items-center gap-1.5 truncate">
                  <span className="text-[10px] font-mono bg-cyan-100 text-cyan-900 px-1.5 py-0.2 rounded font-bold shrink-0">
                    {activeEquipment.id}
                  </span>
                  <span className="font-bold text-slate-900 truncate">{activeEquipment.name}</span>
                  <span className="text-[10px] text-slate-400 shrink-0">({activeEquipment.type})</span>
                </div>
              ) : (
                <span className={cleanVal ? 'font-mono font-bold text-slate-800' : 'text-slate-400'}>
                  {cleanVal || placeholder}
                </span>
              )}
            </div>
          </div>
          <ChevronDown
            size={14}
            className={`text-slate-400 group-hover:text-cyan-600 shrink-0 transition-transform ${
              isOpen ? 'rotate-180' : ''
            }`}
          />
        </button>

        {cleanVal && (
          <button
            type="button"
            onClick={(e) => {
              e.stopPropagation();
              onChange('');
            }}
            className="absolute right-7 top-1/2 -translate-y-1/2 p-1 text-slate-400 hover:text-slate-600"
            title="Xóa lựa chọn"
          >
            <X size={12} />
          </button>
        )}
      </div>

      {/* Floating Dropdown */}
      {isOpen && (
        <div className="absolute left-0 right-0 mt-1.5 bg-white border border-slate-200 rounded-xl shadow-2xl z-50 overflow-hidden animate-in fade-in-50 zoom-in-95 duration-100 min-w-[340px]">
          {/* Header & Search */}
          <div className="p-2.5 bg-slate-50 border-b border-slate-200 space-y-2">
            <div className="flex items-center justify-between text-xs font-bold text-slate-700">
              <span className="flex items-center gap-1.5 text-cyan-900">
                <FileSpreadsheet size={13} className="text-cyan-600" />
                Danh Mục Thiết Bị ({equipmentList.length})
              </span>
              <button
                type="button"
                onClick={() => {
                  setShowAddModal(true);
                  setIsOpen(false);
                }}
                className="text-[11px] font-bold text-cyan-700 hover:text-cyan-900 flex items-center gap-1 bg-cyan-100/70 hover:bg-cyan-100 px-2 py-0.5 rounded border border-cyan-200 transition-colors"
              >
                <Plus size={12} /> Bổ sung mới
              </button>
            </div>

            {selectedCustomer && (
              <div className="flex items-center justify-between text-[11px] bg-slate-100 px-2 py-1 rounded">
                <span className="text-slate-600 truncate">
                  Đang lọc theo: <strong>{selectedCustomer}</strong>
                </span>
                <button
                  type="button"
                  onClick={() => setShowAllCustomerEq(!showAllCustomerEq)}
                  className="text-cyan-700 font-bold hover:underline shrink-0 ml-2"
                >
                  {showAllCustomerEq ? 'Chỉ hiện của khách' : 'Hiện tất cả'}
                </button>
              </div>
            )}

            <div className="relative">
              <Search size={13} className="absolute left-2.5 top-2.5 text-slate-400" />
              <input
                type="text"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Tìm mã Tag, tên thiết bị, chủng loại, vị trí..."
                className="w-full text-xs pl-8 pr-2.5 py-1.5 border border-slate-200 rounded-lg focus:outline-none focus:ring-1 focus:ring-cyan-500 bg-white"
                autoFocus
              />
            </div>
          </div>

          {/* List items */}
          <div className="max-h-60 overflow-y-auto p-1 divide-y divide-slate-100">
            {filtered.length === 0 ? (
              <div className="p-4 text-center">
                <p className="text-xs text-slate-500 mb-2">Không tìm thấy thiết bị nào với "{search}"</p>
                <button
                  type="button"
                  onClick={() => {
                    setNewEq({ ...newEq, id: search, name: `Thiết bị ${search}` });
                    setShowAddModal(true);
                    setIsOpen(false);
                  }}
                  className="inline-flex items-center gap-1 text-xs font-bold text-cyan-700 hover:text-cyan-900 bg-cyan-50 hover:bg-cyan-100 px-2.5 py-1 rounded-lg border border-cyan-200"
                >
                  <Plus size={12} /> Tạo thiết bị "{search}" ngay
                </button>
              </div>
            ) : (
              filtered.map((eq) => {
                const isSelected = activeEquipment?.id === eq.id;
                return (
                  <button
                    key={eq.id}
                    type="button"
                    onClick={() => handleSelect(eq)}
                    className={`w-full text-left p-2 rounded-lg flex items-center justify-between text-xs transition-colors ${
                      isSelected
                        ? 'bg-cyan-50 text-cyan-950 font-semibold'
                        : 'hover:bg-slate-50 text-slate-700'
                    }`}
                  >
                    <div className="min-w-0 pr-2">
                      <div className="flex items-center gap-1.5 truncate">
                        <span className="text-[10px] font-mono bg-slate-100 text-slate-800 px-1.5 py-0.2 rounded font-bold shrink-0">
                          {eq.id}
                        </span>
                        <span className="font-bold text-slate-900 truncate">{eq.name}</span>
                        <span className="text-[10px] bg-cyan-50 text-cyan-800 px-1.5 py-0.2 rounded border border-cyan-200 shrink-0 font-medium">
                          {eq.type}
                        </span>
                      </div>
                      <div className="text-[10px] text-slate-500 flex items-center gap-2 mt-0.5 truncate">
                        <span className="truncate text-slate-600">{eq.customer}</span>
                        {eq.location && (
                          <span className="flex items-center gap-0.5 text-slate-400 truncate">
                            <MapPin size={10} className="shrink-0" />
                            {eq.location}
                          </span>
                        )}
                      </div>
                    </div>
                    {isSelected && <Check size={16} className="text-cyan-600 shrink-0 ml-1" />}
                  </button>
                );
              })
            )}
          </div>

          {/* Footer Action */}
          <div className="p-2 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-[11px]">
            <span className="text-slate-500 font-medium flex items-center gap-1">
              <Sparkles size={11} className="text-amber-500" /> Đồng bộ 2 chiều sheet <strong>thietbi</strong>
            </span>
            <button
              type="button"
              onClick={() => {
                setShowAddModal(true);
                setIsOpen(false);
              }}
              className="text-cyan-700 font-bold hover:underline flex items-center gap-1"
            >
              <Plus size={12} /> Bổ sung mới
            </button>
          </div>
        </div>
      )}

      {/* QUICK ADD MODAL */}
      {showAddModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-md w-full border border-slate-200 overflow-hidden animate-in fade-in-50 zoom-in-95 duration-150">
            <div className="p-4 bg-gradient-to-r from-cyan-700 to-blue-800 text-white flex items-center justify-between">
              <div className="flex items-center gap-2">
                <div className="p-1.5 bg-white/20 rounded-lg">
                  <Cpu size={18} />
                </div>
                <div>
                  <h3 className="font-bold text-sm">Bổ Sung Thiết Bị Mới</h3>
                  <p className="text-[11px] text-cyan-100">
                    Tự động lưu vào danh mục & đồng bộ Sheet <code>thietbi</code>
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setShowAddModal(false)}
                className="p-1 hover:bg-white/20 rounded-lg text-white/80 hover:text-white transition-colors"
              >
                <X size={16} />
              </button>
            </div>

            <form onSubmit={handleAddNewSubmit} className="p-4 space-y-3 text-xs">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    Mã Thiết Bị / Tag ID <span className="text-red-500">*</span>
                  </label>
                  <input
                    type="text"
                    required
                    placeholder="VD: TR-03, MVCB-05"
                    value={newEq.id}
                    onChange={(e) => setNewEq({ ...newEq, id: e.target.value.toUpperCase() })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:outline-none font-mono font-bold"
                  />
                </div>

                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    Phân Loại Thiết Bị
                  </label>
                  <select
                    value={newEq.type}
                    onChange={(e) =>
                      setNewEq({ ...newEq, type: e.target.value as EquipmentCategoryType })
                    }
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:outline-none"
                  >
                    {EQUIPMENT_CATEGORIES.map((cat) => (
                      <option key={cat} value={cat}>
                        {cat}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">
                  Tên Thiết Bị Đầy Đủ <span className="text-red-500">*</span>
                </label>
                <input
                  type="text"
                  required
                  placeholder="VD: Máy biến áp dầu T1 2500kVA 22/0.4kV..."
                  value={newEq.name}
                  onChange={(e) => setNewEq({ ...newEq, name: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:outline-none font-medium"
                />
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">
                  Thuộc Khách Hàng / Đơn Vị
                </label>
                <input
                  type="text"
                  placeholder="VD: Công ty Cổ phần Năng lượng TEV"
                  value={newEq.customer}
                  onChange={(e) => setNewEq({ ...newEq, customer: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:outline-none"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    Nhà Máy / Site
                  </label>
                  <input
                    type="text"
                    placeholder="VD: Nhà máy Bắc Ninh"
                    value={newEq.factory}
                    onChange={(e) => setNewEq({ ...newEq, factory: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:outline-none"
                  />
                </div>

                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Vị Trí / Ngăn Lộ</label>
                  <input
                    type="text"
                    placeholder="VD: Trạm TBA-01 Ngăn F02"
                    value={newEq.location}
                    onChange={(e) => setNewEq({ ...newEq, location: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:outline-none"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Hãng Sản Xuất</label>
                  <input
                    type="text"
                    placeholder="VD: ABB, Schneider, Thibidi"
                    value={newEq.manufacturer}
                    onChange={(e) => setNewEq({ ...newEq, manufacturer: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:outline-none"
                  />
                </div>

                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Năm Chế Tạo</label>
                  <input
                    type="text"
                    placeholder="VD: 2023"
                    value={newEq.yearOfManufacture}
                    onChange={(e) => setNewEq({ ...newEq, yearOfManufacture: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:outline-none font-mono"
                  />
                </div>
              </div>

              <div className="pt-2 flex items-center justify-end gap-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setShowAddModal(false)}
                  className="px-3 py-1.5 border border-slate-300 text-slate-700 font-medium rounded-lg hover:bg-slate-50 transition-colors"
                >
                  Hủy
                </button>
                <button
                  type="submit"
                  className="px-4 py-1.5 bg-cyan-700 hover:bg-cyan-800 text-white font-bold rounded-lg shadow-sm transition-colors flex items-center gap-1.5"
                >
                  <Check size={14} /> Lưu Thiết Bị
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};
