import React, { useState, useEffect, useRef } from 'react';
import {
  Building2,
  ChevronDown,
  Plus,
  Search,
  Check,
  X,
  MapPin,
  Mail,
  Phone,
  Factory,
  FileSpreadsheet,
  Sparkles
} from 'lucide-react';
import { CustomerRecord } from '../utils/fieldSyncService';
import {
  getRegisteredCustomers,
  saveCustomerRecord,
  EVT_CUSTOMERS_UPDATED
} from '../services/masterDirectoryService';
import { getNextCustomerId } from '../services/netaCmmsBridgeService';

interface CustomerSelectDropdownProps {
  value: string;
  onChange: (customerValue: string, customer?: CustomerRecord) => void;
  displayMode?: 'name' | 'id' | 'both';
  placeholder?: string;
  className?: string;
  label?: string;
  required?: boolean;
}

export const CustomerSelectDropdown: React.FC<CustomerSelectDropdownProps> = ({
  value,
  onChange,
  displayMode = 'name',
  placeholder = 'Chọn hoặc thêm Mã / Tên Khách Hàng...',
  className = '',
  label,
  required = false
}) => {
  const [customers, setCustomers] = useState<CustomerRecord[]>(() => getRegisteredCustomers());
  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState('');
  const [showAddModal, setShowAddModal] = useState(false);

  // New customer form state
  const [newCust, setNewCust] = useState({
    id: '',
    name: '',
    factories: '',
    email: '',
    phone: '',
    address: ''
  });

  const dropdownRef = useRef<HTMLDivElement>(null);

  // Listen for external updates across app
  useEffect(() => {
    const handleUpdate = (e: any) => {
      if (Array.isArray(e.detail)) {
        setCustomers(e.detail);
      } else {
        setCustomers(getRegisteredCustomers());
      }
    };
    window.addEventListener(EVT_CUSTOMERS_UPDATED, handleUpdate);
    return () => window.removeEventListener(EVT_CUSTOMERS_UPDATED, handleUpdate);
  }, []);

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

  // Determine current active customer
  const cleanVal = (value || '').trim();
  const activeCustomer = customers.find(
    (c) =>
      c.id.toLowerCase() === cleanVal.toLowerCase() ||
      c.name.toLowerCase() === cleanVal.toLowerCase() ||
      cleanVal.toLowerCase().includes(c.id.toLowerCase()) ||
      cleanVal.toLowerCase().includes(c.name.toLowerCase())
  );

  const handleSelect = (cust: CustomerRecord) => {
    let result = cust.name;
    if (displayMode === 'id') result = cust.id;
    else if (displayMode === 'both') result = `${cust.id} - ${cust.name}`;
    onChange(result, cust);
    setIsOpen(false);
    setSearch('');
  };

  const handleOpenAddModal = () => {
    const nextCode = getNextCustomerId(customers);
    setNewCust({
      id: nextCode,
      name: search.trim(),
      factories: '',
      email: '',
      phone: '',
      address: ''
    });
    setShowAddModal(true);
    setIsOpen(false);
  };

  const handleAddNewSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const name = newCust.name.trim();
    if (!name) return;

    let id = newCust.id.trim();
    if (!id) {
      id = getNextCustomerId(customers);
    } else if (!id.toUpperCase().startsWith('KH')) {
      const digits = id.replace(/\D/g, '');
      id = `KH${(digits || '1').padStart(3, '0')}`;
    } else {
      id = id.toUpperCase();
    }

    const factoriesList = newCust.factories
      .split(',')
      .map((f) => f.trim())
      .filter((f) => f.length > 0);

    const created: CustomerRecord = {
      id,
      name,
      factories: factoriesList.length > 0 ? factoriesList : ['Nhà máy chính'],
      email: newCust.email.trim(),
      phone: newCust.phone.trim(),
      address: newCust.address.trim(),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    const updatedList = saveCustomerRecord(created);
    setCustomers(updatedList);
    handleSelect(created);
    setShowAddModal(false);
    setNewCust({ id: '', name: '', factories: '', email: '', phone: '', address: '' });
  };

  const filtered = customers.filter(
    (c) =>
      c.name.toLowerCase().includes(search.toLowerCase()) ||
      c.id.toLowerCase().includes(search.toLowerCase()) ||
      c.factories?.some((f) => f.toLowerCase().includes(search.toLowerCase())) ||
      (c.address && c.address.toLowerCase().includes(search.toLowerCase()))
  );

  const displayLabel = activeCustomer
    ? displayMode === 'id'
      ? activeCustomer.id
      : displayMode === 'both'
      ? `${activeCustomer.id} - ${activeCustomer.name}`
      : activeCustomer.name
    : cleanVal || placeholder;

  return (
    <div className={`relative ${className}`} ref={dropdownRef}>
      {label && (
        <label className="block text-xs font-semibold text-slate-700 mb-1 flex items-center justify-between">
          <span className="flex items-center gap-1.5">
            <Building2 size={13} className="text-blue-600" />
            {label} {required && <span className="text-red-500">*</span>}
          </span>
          <span className="text-[10px] font-mono text-blue-600 bg-blue-50 px-1.5 py-0.2 rounded border border-blue-200">
            Sheet khachhang
          </span>
        </label>
      )}

      {/* Trigger Button */}
      <div className="relative">
        <button
          type="button"
          onClick={() => setIsOpen(!isOpen)}
          className={`w-full text-left bg-white border rounded-lg px-3 py-2 text-xs font-medium focus:ring-2 focus:ring-blue-500 focus:outline-none transition-all flex items-center justify-between shadow-2xs group ${
            activeCustomer ? 'border-blue-300 bg-blue-50/20' : 'border-slate-300 hover:border-blue-400'
          }`}
        >
          <div className="flex items-center gap-2 min-w-0 pr-1">
            <div className="w-5 h-5 rounded-md bg-blue-100 text-blue-700 flex items-center justify-center text-[10px] font-bold shrink-0">
              <Building2 size={12} />
            </div>
            <div className="truncate">
              {activeCustomer ? (
                <div className="flex items-center gap-1.5 truncate">
                  <span className="font-bold text-slate-900 truncate">{activeCustomer.name}</span>
                  <span className="text-[10px] font-mono bg-blue-100 text-blue-800 px-1 py-0.2 rounded font-semibold shrink-0">
                    {activeCustomer.id}
                  </span>
                </div>
              ) : (
                <span className={cleanVal ? 'font-medium text-slate-800' : 'text-slate-400'}>
                  {cleanVal || placeholder}
                </span>
              )}
            </div>
          </div>
          <ChevronDown
            size={14}
            className={`text-slate-400 group-hover:text-blue-600 shrink-0 transition-transform ${
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
        <div className="absolute left-0 right-0 mt-1.5 bg-white border border-slate-200 rounded-xl shadow-2xl z-50 overflow-hidden animate-in fade-in-50 zoom-in-95 duration-100 min-w-[320px]">
          {/* Header & Search */}
          <div className="p-2.5 bg-slate-50 border-b border-slate-200 space-y-2">
            <div className="flex items-center justify-between text-xs font-bold text-slate-700">
              <span className="flex items-center gap-1.5 text-blue-800">
                <FileSpreadsheet size={13} className="text-blue-600" />
                Danh Mục Khách Hàng ({customers.length})
              </span>
              <button
                type="button"
                onClick={handleOpenAddModal}
                className="text-[11px] font-bold text-blue-700 hover:text-blue-900 flex items-center gap-1 bg-blue-100/70 hover:bg-blue-100 px-2 py-0.5 rounded border border-blue-200 transition-colors"
              >
                <Plus size={12} /> Bổ sung mới
              </button>
            </div>
            <div className="relative">
              <Search size={13} className="absolute left-2.5 top-2.5 text-slate-400" />
              <input
                type="text"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Tìm mã KH, tên doanh nghiệp, nhà máy..."
                className="w-full text-xs pl-8 pr-2.5 py-1.5 border border-slate-200 rounded-lg focus:outline-none focus:ring-1 focus:ring-blue-500 bg-white"
                autoFocus
              />
            </div>
          </div>

          {/* List items */}
          <div className="max-h-60 overflow-y-auto p-1 divide-y divide-slate-100">
            {filtered.length === 0 ? (
              <div className="p-4 text-center">
                <p className="text-xs text-slate-500 mb-2">Không tìm thấy khách hàng nào với "{search}"</p>
                <button
                  type="button"
                  onClick={handleOpenAddModal}
                  className="inline-flex items-center gap-1 text-xs font-bold text-blue-600 hover:text-blue-800 bg-blue-50 hover:bg-blue-100 px-2.5 py-1 rounded-lg border border-blue-200"
                >
                  <Plus size={12} /> Tạo khách hàng "{search}" ngay
                </button>
              </div>
            ) : (
              filtered.map((cust) => {
                const isSelected = activeCustomer?.id === cust.id;
                return (
                  <button
                    key={cust.id}
                    type="button"
                    onClick={() => handleSelect(cust)}
                    className={`w-full text-left p-2 rounded-lg flex items-center justify-between text-xs transition-colors ${
                      isSelected
                        ? 'bg-blue-50 text-blue-950 font-semibold'
                        : 'hover:bg-slate-50 text-slate-700'
                    }`}
                  >
                    <div className="min-w-0 pr-2">
                      <div className="flex items-center gap-1.5 truncate">
                        <span className="font-bold text-slate-900 truncate">{cust.name}</span>
                        <span className="text-[10px] font-mono bg-blue-100 text-blue-800 px-1.5 py-0.2 rounded font-semibold shrink-0">
                          {cust.id}
                        </span>
                      </div>
                      <div className="text-[10px] text-slate-500 flex items-center gap-2 mt-0.5 truncate">
                        {cust.factories && cust.factories.length > 0 && (
                          <span className="flex items-center gap-0.5 text-slate-600 truncate">
                            <Factory size={10} className="shrink-0 text-slate-400" />
                            {cust.factories.join(', ')}
                          </span>
                        )}
                        {cust.address && (
                          <span className="flex items-center gap-0.5 text-slate-400 truncate">
                            <MapPin size={10} className="shrink-0" />
                            {cust.address}
                          </span>
                        )}
                      </div>
                    </div>
                    {isSelected && <Check size={16} className="text-blue-600 shrink-0 ml-1" />}
                  </button>
                );
              })
            )}
          </div>

          {/* Footer Action */}
          <div className="p-2 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-[11px]">
            <span className="text-slate-500 font-medium flex items-center gap-1">
              <Sparkles size={11} className="text-amber-500" /> Đồng bộ 2 chiều sheet <strong>khachhang</strong>
            </span>
            <button
              type="button"
              onClick={handleOpenAddModal}
              className="text-blue-700 font-bold hover:underline flex items-center gap-1"
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
            <div className="p-4 bg-gradient-to-r from-blue-700 to-indigo-800 text-white flex items-center justify-between">
              <div className="flex items-center gap-2">
                <div className="p-1.5 bg-white/20 rounded-lg">
                  <Building2 size={18} />
                </div>
                <div>
                  <h3 className="font-bold text-sm">Bổ Sung Khách Hàng Mới</h3>
                  <p className="text-[11px] text-blue-100">
                    Tự động lưu vào danh mục & đồng bộ Sheet <code>khachhang</code>
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
              <div>
                <label className="block font-semibold text-slate-700 mb-1">
                  Tên Khách Hàng / Doanh Nghiệp <span className="text-red-500">*</span>
                </label>
                <input
                  type="text"
                  required
                  placeholder="VD: Tập đoàn Dệt may Việt Nam, Samsung SEVT..."
                  value={newCust.name}
                  onChange={(e) => setNewCust({ ...newCust, name: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none font-medium"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    Mã Khách Hàng (Định dạng KHXXX)
                  </label>
                  <input
                    type="text"
                    placeholder="VD: KH006"
                    value={newCust.id}
                    onChange={(e) => setNewCust({ ...newCust, id: e.target.value.toUpperCase() })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none font-mono font-bold text-blue-700"
                  />
                  <p className="text-[10px] text-slate-400 mt-0.5">Tự động lấy STT tiếp theo dạng KHXXX</p>
                </div>

                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Số Điện Thoại</label>
                  <input
                    type="text"
                    placeholder="VD: 024 3888 xxxx"
                    value={newCust.phone}
                    onChange={(e) => setNewCust({ ...newCust, phone: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                  />
                </div>
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">
                  Nhà Máy / Phân Xưởng (Phân cách bằng dấu phẩy)
                </label>
                <input
                  type="text"
                  placeholder="VD: Nhà máy Bắc Ninh, Trạm 110kV KCN Yên Phong"
                  value={newCust.factories}
                  onChange={(e) => setNewCust({ ...newCust, factories: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">Email Liên Hệ</label>
                <input
                  type="email"
                  placeholder="contact@company.com"
                  value={newCust.email}
                  onChange={(e) => setNewCust({ ...newCust, email: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">Địa Chỉ Trụ Sở</label>
                <input
                  type="text"
                  placeholder="VD: Lô CN-01, KCN Quế Võ, Bắc Ninh"
                  value={newCust.address}
                  onChange={(e) => setNewCust({ ...newCust, address: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                />
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
                  className="px-4 py-1.5 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-lg shadow-sm transition-colors flex items-center gap-1.5"
                >
                  <Check size={14} /> Lưu Khách Hàng
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};
