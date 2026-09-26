import React, { useState, useMemo } from 'react';
import {
  Users,
  Search,
  Plus,
  Camera,
  Layers,
  Zap,
  Power,
  RotateCcw,
  CheckCircle2,
  AlertTriangle,
  FileText,
  Save,
  RefreshCw,
  ExternalLink,
  ShieldCheck,
  Building2,
  MapPin,
  Tag,
  Gauge,
  Activity,
  Calendar,
  Sparkles,
  ArrowRight,
  Database,
  Cpu,
  Check,
  Info,
  ChevronDown,
  Printer
} from 'lucide-react';
import {
  CustomerRecord,
  EquipmentMasterRecord,
  EquipmentCategoryType,
  customerToSheetRow,
  equipmentToSheetRow,
  inspectionReportToSheetRow
} from '../utils/fieldSyncService';
import { NameplateScannerModal } from './NameplateScannerModal';
import { ExtractedNameplateData } from '../../server/nameplateOcrAnalyzer';

interface FieldDataEntryHubProps {
  customers: CustomerRecord[];
  allEquipment: EquipmentMasterRecord[];
  isGoogleConnected: boolean;
  onSyncAll: () => Promise<void>;
  onSaveCustomer: (customer: CustomerRecord) => Promise<void>;
  onSaveEquipment: (equipment: EquipmentMasterRecord) => Promise<void>;
  onSaveReport: (report: any) => Promise<void>;
  onNavigateToChecklist?: (categoryMode: string) => void;
  currentUser?: any;
}

export const FieldDataEntryHub: React.FC<FieldDataEntryHubProps> = ({
  customers,
  allEquipment,
  isGoogleConnected,
  onSyncAll,
  onSaveCustomer,
  onSaveEquipment,
  onSaveReport,
  onNavigateToChecklist,
  currentUser
}) => {
  // --- STATE ---
  // Step 1: Customer
  const [selectedCustomerId, setSelectedCustomerId] = useState<string>(customers[0]?.id || '');
  const [customerSearch, setCustomerSearch] = useState('');
  const [showAddCustomerModal, setShowAddCustomerModal] = useState(false);
  const [newCustomerForm, setNewCustomerForm] = useState<Partial<CustomerRecord>>({
    id: `KH-${Date.now().toString().slice(-4)}`,
    name: '',
    factories: ['Nhà máy chính'],
    email: '',
    phone: '',
    address: '',
    contactPerson: '',
    notes: ''
  });

  // Step 2: Equipment Selection Mode
  const [eqSelectionMode, setEqSelectionMode] = useState<'preset' | 'scan' | 'manual'>('preset');
  const [selectedEquipmentId, setSelectedEquipmentId] = useState<string>('');
  const [showScannerModal, setShowScannerModal] = useState(false);
  const [showAddEquipmentModal, setShowAddEquipmentModal] = useState(false);

  // New Equipment Form (for manual entry or after scan)
  const [newEqForm, setNewEqForm] = useState<Partial<EquipmentMasterRecord>>({
    id: `EQ-${Date.now().toString().slice(-4)}`,
    name: '',
    customer: '',
    factory: '',
    location: '',
    type: 'Máy biến áp',
    manufacturer: '',
    model: '',
    serialNumber: '',
    yearOfManufacture: new Date().getFullYear().toString(),
    criticality: 'B',
    status: 'healthy',
    health: 95
  });

  // Step 3: Equipment Active Type & Dynamic Field Measurements
  const [activeCategory, setActiveCategory] = useState<EquipmentCategoryType>('Máy biến áp');
  const [measurements, setMeasurements] = useState<Record<string, any>>({});
  const [fseNotes, setFseNotes] = useState('');
  const [isSaving, setIsSaving] = useState(false);
  const [isSyncing, setIsSyncing] = useState(false);
  const [saveSuccessMsg, setSaveSuccessMsg] = useState<string | null>(null);

  // Derived selected customer object
  const currentCustomer = useMemo(() => {
    return customers.find(c => c.id === selectedCustomerId) || customers[0] || null;
  }, [customers, selectedCustomerId]);

  // Equipment list filtered by selected customer
  const filteredEquipment = useMemo(() => {
    if (!currentCustomer) return allEquipment;
    return allEquipment.filter(eq => 
      eq.customer.toLowerCase().includes(currentCustomer.name.toLowerCase()) ||
      currentCustomer.name.toLowerCase().includes(eq.customer.toLowerCase()) ||
      eq.customer === currentCustomer.id
    );
  }, [allEquipment, currentCustomer]);

  // Derived selected equipment object
  const currentEquipment = useMemo(() => {
    return allEquipment.find(eq => eq.id === selectedEquipmentId) || filteredEquipment[0] || null;
  }, [allEquipment, filteredEquipment, selectedEquipmentId]);

  // When equipment changes, sync active category and load previous measurements
  const handleSelectEquipment = (eq: EquipmentMasterRecord) => {
    setSelectedEquipmentId(eq.id);
    setActiveCategory(eq.type || 'Máy biến áp');
    if (eq.measurements) {
      setMeasurements({ ...eq.measurements });
    } else {
      setMeasurements({});
    }
  };

  // Measurement input helper
  const handleMeasurementChange = (key: string, value: any) => {
    setMeasurements(prev => ({ ...prev, [key]: value }));
  };

  // Handle OCR scan applied data
  const handleApplyScanData = (extracted: ExtractedNameplateData) => {
    setShowScannerModal(false);
    
    // Determine category from OCR
    let cat: EquipmentCategoryType = 'Máy biến áp';
    if (extracted.equipment_category) {
      cat = extracted.equipment_category;
    } else {
      const lower = (extracted.equipment_name + ' ' + extracted.equipment_type).toLowerCase();
      if (lower.includes('cắt') || lower.includes('breaker') || lower.includes('acb') || lower.includes('vcb')) cat = 'Máy cắt';
      else if (lower.includes('động cơ') || lower.includes('motor')) cat = 'Động cơ';
      else if (lower.includes('máy phát') || lower.includes('generator')) cat = 'Máy phát';
      else if (lower.includes('tủ') || lower.includes('switchgear')) cat = 'Tủ điện';
      else cat = 'Máy biến áp';
    }

    setActiveCategory(cat);

    const generatedId = extracted.equipment_tag || `EQ-SCAN-${Date.now().toString().slice(-4)}`;
    const newRecord: EquipmentMasterRecord = {
      id: generatedId,
      name: extracted.equipment_name || `${cat} Scan từ Site`,
      customer: currentCustomer ? currentCustomer.name : 'Khách hàng hiện trường',
      factory: currentCustomer?.factories?.[0] || 'Phân xưởng chính',
      location: 'Hiện trường kiểm tra',
      type: cat,
      manufacturer: extracted.manufacturer || '',
      model: extracted.serial_number || '',
      serialNumber: extracted.serial_number || '',
      yearOfManufacture: extracted.manufacturing_year || new Date().getFullYear().toString(),
      status: 'healthy',
      health: 95,
      criticality: 'B',
      source: 'ocr',
      specs: {
        rawText: extracted.full_nameplate_text,
        transformer: cat === 'Máy biến áp' ? {
          ratedPowerKva: extracted.rated_power_kva || undefined,
          primaryVoltageKv: extracted.primary_voltage_kv || undefined,
          secondaryVoltageV: extracted.secondary_voltage_v || undefined,
          vectorGroup: extracted.vector_group || 'Dyn11',
          insulationType: extracted.oil_type || 'Mineral Oil',
          impedanceZ: extracted.impedance_percent_z || undefined
        } : undefined,
        breaker: cat === 'Máy cắt' ? {
          ratedVoltageKv: extracted.primary_voltage_kv || 0.4,
          ratedCurrentA: extracted.rated_current_a || 1600,
          breakingCapacityKa: extracted.breaking_capacity_ka || 50,
          breakerType: extracted.equipment_type || 'ACB'
        } : undefined,
        motor: cat === 'Động cơ' ? {
          ratedPowerKw: extracted.rated_power_kw || (extracted.rated_power_kva ? Math.round(extracted.rated_power_kva * 0.8) : 250),
          ratedVoltageV: extracted.secondary_voltage_v || 380,
          ratedCurrentA: extracted.rated_current_a || undefined
        } : undefined,
        generator: cat === 'Máy phát' ? {
          ratedPowerKva: extracted.rated_power_kva || 1000,
          ratedVoltageV: extracted.secondary_voltage_v || 400,
          engineManufacturer: extracted.manufacturer || ''
        } : undefined
      }
    };

    // Save equipment to state & Google Sheet
    onSaveEquipment(newRecord);
    setSelectedEquipmentId(generatedId);
    setEqSelectionMode('preset');
    setSaveSuccessMsg(`Đã nhận dạng nameplate thành công: ${newRecord.name} (${newRecord.id})!`);
    setTimeout(() => setSaveSuccessMsg(null), 4000);
  };

  // Submit New Customer
  const handleCreateCustomer = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newCustomerForm.name) return;

    const customerToSave: CustomerRecord = {
      id: newCustomerForm.id || `KH-${Date.now().toString().slice(-4)}`,
      name: newCustomerForm.name,
      factories: typeof newCustomerForm.factories === 'string' 
        ? (newCustomerForm.factories as string).split(',').map(s => s.trim()) 
        : (newCustomerForm.factories || ['Nhà máy chính']),
      email: newCustomerForm.email || '',
      phone: newCustomerForm.phone || '',
      address: newCustomerForm.address || '',
      contactPerson: newCustomerForm.contactPerson || '',
      notes: newCustomerForm.notes || '',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    await onSaveCustomer(customerToSave);
    setSelectedCustomerId(customerToSave.id);
    setShowAddCustomerModal(false);
    setNewCustomerForm({
      id: `KH-${Date.now().toString().slice(-4)}`,
      name: '',
      factories: ['Nhà máy chính'],
      email: '',
      phone: '',
      address: '',
      contactPerson: '',
      notes: ''
    });

    setSaveSuccessMsg(`Đã tạo khách hàng "${customerToSave.name}" & tự động đồng bộ sang Google Sheet khachhang!`);
    setTimeout(() => setSaveSuccessMsg(null), 4000);
  };

  // Submit New Equipment Manually
  const handleCreateEquipmentManual = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newEqForm.name) return;

    const eqToSave: EquipmentMasterRecord = {
      id: newEqForm.id || `EQ-${Date.now().toString().slice(-4)}`,
      name: newEqForm.name,
      customer: currentCustomer ? currentCustomer.name : (newEqForm.customer || 'Khách hàng'),
      factory: newEqForm.factory || currentCustomer?.factories?.[0] || 'Nhà máy',
      location: newEqForm.location || 'Trạm chính',
      type: newEqForm.type || activeCategory,
      manufacturer: newEqForm.manufacturer || '',
      model: newEqForm.model || '',
      serialNumber: newEqForm.serialNumber || '',
      yearOfManufacture: newEqForm.yearOfManufacture || new Date().getFullYear().toString(),
      criticality: newEqForm.criticality || 'B',
      status: 'healthy',
      health: 95,
      lastCheck: new Date().toLocaleDateString('vi-VN'),
      source: 'app'
    };

    await onSaveEquipment(eqToSave);
    setSelectedEquipmentId(eqToSave.id);
    setActiveCategory(eqToSave.type);
    setShowAddEquipmentModal(false);
    setEqSelectionMode('preset');

    setSaveSuccessMsg(`Đã thêm thiết bị "${eqToSave.name}" (${eqToSave.id}) & đồng bộ sang Google Sheet thietbi!`);
    setTimeout(() => setSaveSuccessMsg(null), 4000);
  };

  // Real-time Health Index Evaluation based on dynamic inputs
  const calculatedHealth = useMemo(() => {
    let score = 95;
    let status: 'healthy' | 'warning' | 'critical' = 'healthy';

    if (activeCategory === 'Máy biến áp') {
      const oilTemp = parseFloat(measurements.oilTemp);
      const c2h2 = parseFloat(measurements.c2h2);
      const dielectric = parseFloat(measurements.dielectricStrength);
      if (oilTemp > 85 || c2h2 > 5 || (dielectric && dielectric < 30)) {
        status = 'critical';
        score = 55;
      } else if (oilTemp > 70 || c2h2 > 1 || (dielectric && dielectric < 45)) {
        status = 'warning';
        score = 75;
      }
    } else if (activeCategory === 'Máy cắt') {
      const contactRes = parseFloat(measurements.contactResPhaseA || measurements.contactResPhaseB || measurements.contactResPhaseC);
      const sf6 = parseFloat(measurements.sf6Pressure);
      if (contactRes > 80 || (sf6 && sf6 < 0.4)) {
        status = 'critical';
        score = 50;
      } else if (contactRes > 40 || (sf6 && sf6 < 0.5)) {
        status = 'warning';
        score = 72;
      }
    } else if (activeCategory === 'Động cơ') {
      const vib = parseFloat(measurements.vibrationDeRms);
      const temp = parseFloat(measurements.statorTemp);
      const ir = parseFloat(measurements.irStatorMegohms);
      if (vib > 7.1 || temp > 110 || (ir && ir < 5)) {
        status = 'critical';
        score = 52;
      } else if (vib > 4.5 || temp > 90 || (ir && ir < 20)) {
        status = 'warning';
        score = 74;
      }
    } else if (activeCategory === 'Máy phát') {
      const atsTime = parseFloat(measurements.atsTransferTimeSec);
      const coolantTemp = parseFloat(measurements.coolantTempC);
      if (atsTime > 15 || coolantTemp > 98) {
        status = 'critical';
        score = 58;
      } else if (atsTime > 10 || coolantTemp > 90) {
        status = 'warning';
        score = 76;
      }
    } else if (activeCategory === 'Tủ điện') {
      const tev = parseFloat(measurements.tevDbMax);
      const ultra = parseFloat(measurements.ultrasonicDbuv);
      if (tev > 29 || ultra > 20) {
        status = 'critical';
        score = 48;
      } else if (tev > 20 || ultra > 10) {
        status = 'warning';
        score = 70;
      }
    }

    return { score, status };
  }, [activeCategory, measurements]);

  // Handle Save Inspection Report to App & Google Sheets
  const handleSaveInspection = async () => {
    if (!currentEquipment) {
      alert('Vui lòng chọn hoặc thêm thiết bị trước khi lưu.');
      return;
    }

    setIsSaving(true);
    try {
      const reportId = `REP-${Date.now().toString().slice(-6)}`;
      const reportPayload = {
        id: reportId,
        date: new Date().toLocaleDateString('vi-VN'),
        customer: currentCustomer ? currentCustomer.name : currentEquipment.customer,
        factory: currentEquipment.factory || currentCustomer?.factories?.[0] || '',
        location: currentEquipment.location || '',
        equipmentId: currentEquipment.id,
        equipmentName: currentEquipment.name,
        type: activeCategory,
        inspector: currentUser?.displayName || 'FSE Engineer',
        status: calculatedHealth.status,
        health: calculatedHealth.score,
        notes: fseNotes || `Biên bản kiểm tra kỹ thuật hiện trường cho ${currentEquipment.name}.`,
        measurements: { ...measurements },
        createdAt: new Date().toISOString()
      };

      await onSaveReport(reportPayload);

      // Also auto-append to sheet 'TEV Service Flatform' via server
      try {
        await fetch('/api/sheets/append-inspection', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ report: reportPayload })
        });
      } catch (err) {
        console.warn('Silent append to sheet error:', err);
      }

      setSaveSuccessMsg(`Đã lưu biên bản ${reportId} và đồng bộ 2 chiều lên Google Sheet TEV Service Flatform!`);
      setTimeout(() => setSaveSuccessMsg(null), 5000);
    } catch (error: any) {
      alert('Lỗi lưu biên bản: ' + error.message);
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="space-y-6">
      {/* 2-WAY SYNC CONTROL BANNER */}
      <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4 sm:p-5 text-white shadow-xl relative overflow-hidden">
        <div className="absolute right-0 top-0 w-96 h-96 bg-blue-500/10 rounded-full blur-3xl pointer-events-none" />
        <div className="relative z-10 flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1.5">
            <div className="flex flex-wrap items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-xs font-bold bg-blue-500/20 text-blue-300 border border-blue-400/30 flex items-center gap-1.5">
                <Database size={13} />
                Đồng Bộ 2 Chiều Google Sheet
              </span>
              <span className={`px-2.5 py-0.5 rounded-full text-xs font-semibold flex items-center gap-1 ${
                isGoogleConnected 
                  ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-400/30' 
                  : 'bg-amber-500/20 text-amber-300 border border-amber-400/30'
              }`}>
                <span className={`w-2 h-2 rounded-full ${isGoogleConnected ? 'bg-emerald-400 animate-pulse' : 'bg-amber-400'}`} />
                {isGoogleConnected ? 'File: TEV Service Flatform (Connected)' : 'Chưa kết nối Google Drive'}
              </span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white flex items-center gap-2">
              <Activity className="text-blue-400" size={24} />
              Nhập Liệu Hiện Trường & Kiến Trúc Đồng Nhất Dữ Liệu
            </h1>
            <p className="text-xs sm:text-sm text-slate-300 max-w-3xl">
              FSE chọn từ danh mục chuẩn hóa sẵn có (sheets <code className="text-blue-300 font-mono">khachhang</code>, <code className="text-blue-300 font-mono">thietbi</code>) hoặc quét OCR nameplate / tạo mới ngay tại site. Mọi dữ liệu tự động đồng bộ 2 chiều theo đúng schema kỹ thuật từng loại thiết bị.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2.5 self-start lg:self-center shrink-0">
            <button
              onClick={async () => {
                setIsSyncing(true);
                try {
                  await onSyncAll();
                  setSaveSuccessMsg('Đã hoàn thành đồng bộ 2 chiều App <-> Google Sheets!');
                  setTimeout(() => setSaveSuccessMsg(null), 4000);
                } finally {
                  setIsSyncing(false);
                }
              }}
              disabled={isSyncing}
              className="px-4 py-2.5 rounded-xl font-bold text-xs bg-blue-600 hover:bg-blue-500 text-white flex items-center gap-2 transition-all shadow-lg shadow-blue-900/40 cursor-pointer disabled:opacity-50"
            >
              <RefreshCw size={14} className={isSyncing ? 'animate-spin' : ''} />
              <span>{isSyncing ? 'Đang Đồng Bộ...' : 'Đồng Bộ 2 Chiều Ngay'}</span>
            </button>
          </div>
        </div>

        {/* Sync Stats Pill Bar */}
        <div className="mt-4 pt-3 border-t border-slate-800 grid grid-cols-2 sm:grid-cols-4 gap-2 text-xs">
          <div className="bg-slate-800/60 rounded-lg p-2 flex items-center justify-between border border-slate-700/50">
            <span className="text-slate-400">Sheet `khachhang`:</span>
            <span className="font-bold text-blue-400">{customers.length} khách hàng</span>
          </div>
          <div className="bg-slate-800/60 rounded-lg p-2 flex items-center justify-between border border-slate-700/50">
            <span className="text-slate-400">Sheet `thietbi`:</span>
            <span className="font-bold text-emerald-400">{allEquipment.length} thiết bị</span>
          </div>
          <div className="bg-slate-800/60 rounded-lg p-2 flex items-center justify-between border border-slate-700/50">
            <span className="text-slate-400">Thiết bị theo KH này:</span>
            <span className="font-bold text-amber-400">{filteredEquipment.length} thiết bị</span>
          </div>
          <div className="bg-slate-800/60 rounded-lg p-2 flex items-center justify-between border border-slate-700/50">
            <span className="text-slate-400">Trạng thái ghi đè:</span>
            <span className="font-bold text-teal-300">An toàn 100%</span>
          </div>
        </div>
      </div>

      {saveSuccessMsg && (
        <div className="p-3 bg-emerald-50 border border-emerald-300 text-emerald-900 text-xs sm:text-sm font-semibold rounded-xl flex items-center gap-2 animate-in fade-in shadow-xs">
          <CheckCircle2 size={18} className="text-emerald-600 shrink-0" />
          <span>{saveSuccessMsg}</span>
        </div>
      )}

      {/* 2-COLUMN MAIN LAYOUT */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* LEFT COLUMN: Customer Selection & Equipment Identification (Step 1 & 2) */}
        <div className="lg:col-span-5 space-y-6">
          
          {/* STEP 1: KHÁCH HÀNG (CUSTOMER) */}
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <span className="w-6 h-6 rounded-full bg-blue-100 text-blue-700 font-black text-xs flex items-center justify-center">1</span>
                <h2 className="font-bold text-slate-900 text-sm sm:text-base flex items-center gap-1.5">
                  <Users size={18} className="text-blue-600" />
                  Đơn Vị Khách Hàng (Sheet khachhang)
                </h2>
              </div>
              <button
                type="button"
                onClick={() => setShowAddCustomerModal(true)}
                className="px-2.5 py-1 text-xs font-bold text-blue-700 bg-blue-50 hover:bg-blue-100 border border-blue-200 rounded-lg flex items-center gap-1 transition-colors"
              >
                <Plus size={13} />
                <span>+ Tạo mới KH</span>
              </button>
            </div>

            {/* Dropdown Customer Selector */}
            <div>
              <label className="block text-xs font-bold text-slate-600 mb-1.5 uppercase tracking-wider">
                Chọn khách hàng từ danh sách đồng bộ
              </label>
              <div className="relative">
                <select
                  value={selectedCustomerId}
                  onChange={(e) => {
                    setSelectedCustomerId(e.target.value);
                    setSelectedEquipmentId(''); // reset equipment selection for new customer
                  }}
                  className="w-full pl-3 pr-8 py-2.5 text-sm bg-slate-50 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 font-semibold text-slate-900 outline-none appearance-none"
                >
                  {customers.map((c) => (
                    <option key={c.id} value={c.id}>
                      {c.name} ({c.id})
                    </option>
                  ))}
                </select>
                <ChevronDown size={16} className="absolute right-3 top-3 text-slate-400 pointer-events-none" />
              </div>
            </div>

            {/* Customer Details Summary Card */}
            {currentCustomer && (
              <div className="p-3.5 bg-blue-50/60 rounded-xl border border-blue-100 text-xs space-y-2">
                <div className="flex items-start justify-between gap-2">
                  <span className="font-extrabold text-blue-950 text-sm">{currentCustomer.name}</span>
                  <span className="px-2 py-0.5 bg-blue-200/60 text-blue-800 rounded font-mono font-bold text-[10px]">
                    {currentCustomer.id}
                  </span>
                </div>
                <div className="grid grid-cols-1 gap-1 text-slate-600">
                  <div className="flex items-center gap-1.5">
                    <Building2 size={13} className="text-blue-500 shrink-0" />
                    <span>Nhà máy: <strong>{currentCustomer.factories?.join(', ') || 'Chính'}</strong></span>
                  </div>
                  {currentCustomer.address && (
                    <div className="flex items-center gap-1.5">
                      <MapPin size={13} className="text-slate-400 shrink-0" />
                      <span className="truncate">{currentCustomer.address}</span>
                    </div>
                  )}
                  {currentCustomer.contactPerson && (
                    <div className="flex items-center gap-1.5">
                      <Users size={13} className="text-slate-400 shrink-0" />
                      <span>Liên hệ: {currentCustomer.contactPerson} {currentCustomer.phone ? `(${currentCustomer.phone})` : ''}</span>
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>

          {/* STEP 2: THIẾT BỊ KIỂM TRA (EQUIPMENT SELECTION - 3 MODES) */}
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <span className="w-6 h-6 rounded-full bg-emerald-100 text-emerald-700 font-black text-xs flex items-center justify-center">2</span>
                <h2 className="font-bold text-slate-900 text-sm sm:text-base flex items-center gap-1.5">
                  <Zap size={18} className="text-emerald-600" />
                  Thiết Bị Kiểm Tra Hiện Trường
                </h2>
              </div>
            </div>

            {/* 3 Selection Mode Pills */}
            <div className="grid grid-cols-3 gap-1.5 p-1 bg-slate-100 rounded-xl text-xs font-bold text-slate-700">
              <button
                type="button"
                onClick={() => setEqSelectionMode('preset')}
                className={`py-2 px-1 text-center rounded-lg transition-all flex flex-col sm:flex-row items-center justify-center gap-1 ${
                  eqSelectionMode === 'preset'
                    ? 'bg-white text-blue-700 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Database size={13} />
                <span>Sheet thietbi ({filteredEquipment.length})</span>
              </button>
              <button
                type="button"
                onClick={() => setShowScannerModal(true)}
                className={`py-2 px-1 text-center rounded-lg transition-all flex flex-col sm:flex-row items-center justify-center gap-1 ${
                  eqSelectionMode === 'scan'
                    ? 'bg-white text-emerald-700 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Camera size={13} />
                <span>Scan Nameplate</span>
              </button>
              <button
                type="button"
                onClick={() => setShowAddEquipmentModal(true)}
                className={`py-2 px-1 text-center rounded-lg transition-all flex flex-col sm:flex-row items-center justify-center gap-1 ${
                  eqSelectionMode === 'manual'
                    ? 'bg-white text-indigo-700 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Plus size={13} />
                <span>Nhập tại site</span>
              </button>
            </div>

            {/* MODE 1: Chọn từ danh sách có sẵn (sheet thietbi) */}
            <div>
              <label className="block text-xs font-bold text-slate-600 mb-1.5 uppercase tracking-wider">
                Chọn thiết bị thuộc khách hàng đã chọn
              </label>
              {filteredEquipment.length > 0 ? (
                <div className="relative">
                  <select
                    value={selectedEquipmentId}
                    onChange={(e) => {
                      const found = allEquipment.find(eq => eq.id === e.target.value);
                      if (found) handleSelectEquipment(found);
                    }}
                    className="w-full pl-3 pr-8 py-2.5 text-sm bg-slate-50 border border-slate-300 rounded-xl focus:ring-2 focus:ring-emerald-500 font-semibold text-slate-900 outline-none appearance-none"
                  >
                    <option value="">-- Chọn thiết bị kiểm định --</option>
                    {filteredEquipment.map((eq) => (
                      <option key={eq.id} value={eq.id}>
                        [{eq.type}] {eq.id} - {eq.name} ({eq.location || eq.factory})
                      </option>
                    ))}
                  </select>
                  <ChevronDown size={16} className="absolute right-3 top-3 text-slate-400 pointer-events-none" />
                </div>
              ) : (
                <div className="p-4 bg-amber-50 rounded-xl border border-amber-200 text-center space-y-2">
                  <p className="text-xs text-amber-800">
                    Chưa có thiết bị nào chuẩn bị trước cho <strong>{currentCustomer?.name}</strong> trong sheet <code className="font-mono">thietbi</code>.
                  </p>
                  <div className="flex items-center justify-center gap-2">
                    <button
                      type="button"
                      onClick={() => setShowScannerModal(true)}
                      className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-xs"
                    >
                      <Camera size={13} />
                      Quét nhãn nameplate AI
                    </button>
                    <button
                      type="button"
                      onClick={() => setShowAddEquipmentModal(true)}
                      className="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-xs"
                    >
                      <Plus size={13} />
                      Nhập thủ công
                    </button>
                  </div>
                </div>
              )}
            </div>

            {/* Selected Equipment Card Preview */}
            {currentEquipment && (
              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 text-xs space-y-2.5">
                <div className="flex items-start justify-between gap-2">
                  <div>
                    <span className="font-extrabold text-slate-900 text-sm block">
                      {currentEquipment.name}
                    </span>
                    <span className="text-[11px] font-mono text-slate-500">Mã: {currentEquipment.id}</span>
                  </div>
                  <span className="px-2.5 py-1 rounded-full font-bold text-[10px] bg-emerald-100 text-emerald-800 border border-emerald-200">
                    {currentEquipment.type}
                  </span>
                </div>

                <div className="grid grid-cols-2 gap-2 text-slate-600 pt-1 border-t border-slate-200">
                  <div>
                    <span className="text-slate-400 block text-[10px]">Hãng & Model:</span>
                    <strong className="text-slate-800">{currentEquipment.manufacturer || 'N/A'} {currentEquipment.model}</strong>
                  </div>
                  <div>
                    <span className="text-slate-400 block text-[10px]">Vị trí / Ngăn lộ:</span>
                    <strong className="text-slate-800">{currentEquipment.location || currentEquipment.factory}</strong>
                  </div>
                  <div>
                    <span className="text-slate-400 block text-[10px]">Số chế tạo (Serial):</span>
                    <strong className="font-mono text-slate-800">{currentEquipment.serialNumber || 'N/A'}</strong>
                  </div>
                  <div>
                    <span className="text-slate-400 block text-[10px]">Kiểm tra gần nhất:</span>
                    <strong className="text-slate-800">{currentEquipment.lastCheck || 'Chưa kiểm tra'}</strong>
                  </div>
                </div>

                {currentEquipment.notes && (
                  <div className="p-2 bg-white rounded border border-slate-200 text-[11px] text-slate-600">
                    <span className="font-bold text-slate-700">Thông số kỹ thuật: </span>
                    {currentEquipment.notes}
                  </div>
                )}
              </div>
            )}
          </div>
        </div>

        {/* RIGHT COLUMN: Dynamic Equipment Technical Fields & Field Inspection Measurements (Step 3) */}
        <div className="lg:col-span-7 space-y-6">
          <div className="bg-white rounded-2xl border border-slate-200 p-5 sm:p-6 shadow-xs space-y-5">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-4">
              <div className="flex items-center gap-2">
                <span className="w-6 h-6 rounded-full bg-indigo-100 text-indigo-700 font-black text-xs flex items-center justify-center">3</span>
                <div>
                  <h2 className="font-bold text-slate-900 text-base sm:text-lg flex items-center gap-2">
                    <Gauge className="text-indigo-600" size={20} />
                    Thông Số Kỹ Thuật Động Theo Chủng Loại
                  </h2>
                  <p className="text-xs text-slate-500">
                    Các trường dữ liệu thích ứng chính xác theo schema của từng loại thiết bị
                  </p>
                </div>
              </div>

              {/* Real-time Health Badge */}
              <div className="flex items-center gap-2 self-start sm:self-center">
                <div className={`px-3 py-1.5 rounded-xl font-black text-xs flex items-center gap-1.5 ${
                  calculatedHealth.status === 'healthy'
                    ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                    : calculatedHealth.status === 'warning'
                    ? 'bg-amber-100 text-amber-800 border border-amber-300'
                    : 'bg-rose-100 text-rose-800 border border-rose-300'
                }`}>
                  <Activity size={14} />
                  <span>Sức khỏe HI: {calculatedHealth.score}% ({calculatedHealth.status.toUpperCase()})</span>
                </div>
              </div>
            </div>

            {/* Equipment Type Tabs */}
            <div className="flex flex-wrap items-center gap-1.5">
              {(['Máy biến áp', 'Máy cắt', 'Động cơ', 'Máy phát', 'Tủ điện'] as EquipmentCategoryType[]).map((cat) => {
                const isActive = activeCategory === cat;
                return (
                  <button
                    key={cat}
                    type="button"
                    onClick={() => setActiveCategory(cat)}
                    className={`px-3 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 ${
                      isActive
                        ? 'bg-indigo-600 text-white shadow-md'
                        : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                    }`}
                  >
                    {cat === 'Máy biến áp' && <Zap size={14} />}
                    {cat === 'Máy cắt' && <Power size={14} />}
                    {cat === 'Động cơ' && <RotateCcw size={14} />}
                    {cat === 'Máy phát' && <Cpu size={14} />}
                    {cat === 'Tủ điện' && <Layers size={14} />}
                    <span>{cat}</span>
                  </button>
                );
              })}
            </div>

            {/* --- DYNAMIC FORM FIELDS ACCORDING TO EQUIPMENT TYPE --- */}
            
            {/* TYPE 1: MÁY BIẾN ÁP */}
            {activeCategory === 'Máy biến áp' && (
              <div className="space-y-4 animate-in fade-in duration-200">
                <div className="p-3 bg-indigo-50/70 border border-indigo-200 rounded-xl text-xs text-indigo-900 flex items-center justify-between">
                  <span className="font-bold">Schema Máy biến áp (Transformer Standards IEEE / IEC 60076 / NETA Sec 7.2)</span>
                  {onNavigateToChecklist && (
                    <button
                      type="button"
                      onClick={() => onNavigateToChecklist('neta-liquid-transformer')}
                      className="text-indigo-700 hover:text-indigo-900 underline font-semibold flex items-center gap-1"
                    >
                      Mở Checklist NETA 7.2 đầy đủ <ArrowRight size={12} />
                    </button>
                  )}
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Nhiệt độ dầu (°C)</label>
                    <input
                      type="number"
                      value={measurements.oilTemp || ''}
                      onChange={(e) => handleMeasurementChange('oilTemp', e.target.value)}
                      placeholder="VD: 55"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-indigo-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">Tiêu chuẩn: ≤ 65°C</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Nhiệt độ cuộn dây (°C)</label>
                    <input
                      type="number"
                      value={measurements.windingTemp || ''}
                      onChange={(e) => handleMeasurementChange('windingTemp', e.target.value)}
                      placeholder="VD: 68"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-indigo-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">Tiêu chuẩn: ≤ 80°C</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Độ bền điện môi dầu (kV)</label>
                    <input
                      type="number"
                      value={measurements.dielectricStrength || ''}
                      onChange={(e) => handleMeasurementChange('dielectricStrength', e.target.value)}
                      placeholder="VD: 65"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-indigo-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">Đạt: ≥ 40kV (IEC 60156)</span>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">IR Cao-Hạ H-L (MΩ)</label>
                    <input
                      type="number"
                      value={measurements.irHighLow || ''}
                      onChange={(e) => handleMeasurementChange('irHighLow', e.target.value)}
                      placeholder="VD: 3200"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-indigo-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">NETA Table 100.5 (≥ 1000 MΩ)</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">IR Cao-Vỏ H-E (MΩ)</label>
                    <input
                      type="number"
                      value={measurements.irHighEarth || ''}
                      onChange={(e) => handleMeasurementChange('irHighEarth', e.target.value)}
                      placeholder="VD: 3000"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-indigo-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">NETA Table 100.5 (≥ 1000 MΩ)</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Độ ẩm dầu (ppm)</label>
                    <input
                      type="number"
                      value={measurements.oilMoisture || ''}
                      onChange={(e) => handleMeasurementChange('oilMoisture', e.target.value)}
                      placeholder="VD: 12"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-indigo-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">Tiêu chuẩn: ≤ 25 ppm</span>
                  </div>
                </div>

                {/* DGA Gases row */}
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1.5 uppercase">
                    Khí hòa tan DGA (ppm) - C2H2, H2, CH4, CO, CO2
                  </label>
                  <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 text-xs">
                    <div>
                      <span className="text-[10px] text-slate-500">C2H2 (Acetylene)</span>
                      <input
                        type="number"
                        value={measurements.c2h2 || ''}
                        onChange={(e) => handleMeasurementChange('c2h2', e.target.value)}
                        placeholder="0"
                        className="w-full p-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-bold text-rose-700"
                      />
                    </div>
                    <div>
                      <span className="text-[10px] text-slate-500">H2 (Hydrogen)</span>
                      <input
                        type="number"
                        value={measurements.h2 || ''}
                        onChange={(e) => handleMeasurementChange('h2', e.target.value)}
                        placeholder="15"
                        className="w-full p-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold text-slate-800"
                      />
                    </div>
                    <div>
                      <span className="text-[10px] text-slate-500">CH4 (Methane)</span>
                      <input
                        type="number"
                        value={measurements.ch4 || ''}
                        onChange={(e) => handleMeasurementChange('ch4', e.target.value)}
                        placeholder="12"
                        className="w-full p-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold text-slate-800"
                      />
                    </div>
                    <div>
                      <span className="text-[10px] text-slate-500">CO (Carbon Monoxide)</span>
                      <input
                        type="number"
                        value={measurements.co || ''}
                        onChange={(e) => handleMeasurementChange('co', e.target.value)}
                        placeholder="180"
                        className="w-full p-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold text-slate-800"
                      />
                    </div>
                    <div>
                      <span className="text-[10px] text-slate-500">CO2 (Carbon Dioxide)</span>
                      <input
                        type="number"
                        value={measurements.co2 || ''}
                        onChange={(e) => handleMeasurementChange('co2', e.target.value)}
                        placeholder="1650"
                        className="w-full p-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold text-slate-800"
                      />
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* TYPE 2: MÁY CẮT (CIRCUIT BREAKER) */}
            {activeCategory === 'Máy cắt' && (
              <div className="space-y-4 animate-in fade-in duration-200">
                <div className="p-3 bg-teal-50/70 border border-teal-200 rounded-xl text-xs text-teal-900 flex items-center justify-between">
                  <span className="font-bold">Schema Máy cắt ACB / VCB / SF6 (NETA ATS Mục 7.6.1.2 & IEC 60947-2)</span>
                  {onNavigateToChecklist && (
                    <button
                      type="button"
                      onClick={() => onNavigateToChecklist('neta-lv-breaker')}
                      className="text-teal-700 hover:text-teal-900 underline font-semibold flex items-center gap-1"
                    >
                      Mở Checklist Máy Cắt NETA 7.6 <ArrowRight size={12} />
                    </button>
                  )}
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Điện trở tiếp điểm Pha A (µΩ)</label>
                    <input
                      type="number"
                      value={measurements.contactResPhaseA || ''}
                      onChange={(e) => handleMeasurementChange('contactResPhaseA', e.target.value)}
                      placeholder="VD: 18.2"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-teal-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">NETA 7.6 (≤ 50 µΩ)</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Điện trở tiếp điểm Pha B (µΩ)</label>
                    <input
                      type="number"
                      value={measurements.contactResPhaseB || ''}
                      onChange={(e) => handleMeasurementChange('contactResPhaseB', e.target.value)}
                      placeholder="VD: 18.5"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-teal-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">Độ lệch 3 pha ≤ 50%</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Điện trở tiếp điểm Pha C (µΩ)</label>
                    <input
                      type="number"
                      value={measurements.contactResPhaseC || ''}
                      onChange={(e) => handleMeasurementChange('contactResPhaseC', e.target.value)}
                      placeholder="VD: 18.4"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-teal-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">Độ lệch 3 pha ≤ 50%</span>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Thời gian đóng cắt (ms)</label>
                    <input
                      type="number"
                      value={measurements.closingTimeMs || ''}
                      onChange={(e) => handleMeasurementChange('closingTimeMs', e.target.value)}
                      placeholder="VD: 45"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-teal-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">Tiêu chuẩn: ≤ 60ms</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Áp suất khí SF6 / Độ chân không</label>
                    <input
                      type="text"
                      value={measurements.sf6Pressure || ''}
                      onChange={(e) => handleMeasurementChange('sf6Pressure', e.target.value)}
                      placeholder="VD: 0.55 MPa hoặc Pass"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-teal-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">SF6 ≥ 0.50 MPa hoặc VCB Pass</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Test Trip Unit (Long/Short/Inst/GF)</label>
                    <select
                      value={measurements.tripUnitTest || 'Pass'}
                      onChange={(e) => handleMeasurementChange('tripUnitTest', e.target.value)}
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-teal-500 outline-none"
                    >
                      <option value="Pass">Đạt (Pass toàn bộ đường đặc tính)</option>
                      <option value="GF Warning">Cảnh báo Ground Fault lệch thời gian</option>
                      <option value="Fail">Không tác động (Fail)</option>
                    </select>
                  </div>
                </div>
              </div>
            )}

            {/* TYPE 3: ĐỘNG CƠ ĐIỆN (ELECTRIC MOTOR) */}
            {activeCategory === 'Động cơ' && (
              <div className="space-y-4 animate-in fade-in duration-200">
                <div className="p-3 bg-blue-50/70 border border-blue-200 rounded-xl text-xs text-blue-900 flex items-center justify-between">
                  <span className="font-bold">Schema Động cơ điện (NETA Sec 7.15 & ISO 10816-3 / IEEE 43)</span>
                  {onNavigateToChecklist && (
                    <button
                      type="button"
                      onClick={() => onNavigateToChecklist('neta-motor')}
                      className="text-blue-700 hover:text-blue-900 underline font-semibold flex items-center gap-1"
                    >
                      Mở Checklist Động Cơ NETA 7.15 <ArrowRight size={12} />
                    </button>
                  )}
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Độ rung đầu tải DE (mm/s RMS)</label>
                    <input
                      type="number"
                      step="0.01"
                      value={measurements.vibrationDeRms || ''}
                      onChange={(e) => handleMeasurementChange('vibrationDeRms', e.target.value)}
                      placeholder="VD: 1.85"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">ISO 10816: Đạt ≤ 2.8 mm/s</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Nhiệt độ Stator (°C)</label>
                    <input
                      type="number"
                      value={measurements.statorTemp || ''}
                      onChange={(e) => handleMeasurementChange('statorTemp', e.target.value)}
                      placeholder="VD: 75"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">Class F: ≤ 105°C</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Nhiệt độ Vòng bi DE/NDE (°C)</label>
                    <input
                      type="number"
                      value={measurements.bearingTempDe || ''}
                      onChange={(e) => handleMeasurementChange('bearingTempDe', e.target.value)}
                      placeholder="VD: 62"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">Tiêu chuẩn: ≤ 80°C</span>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Điện trở cách điện Stator (MΩ)</label>
                    <input
                      type="number"
                      value={measurements.irStatorMegohms || ''}
                      onChange={(e) => handleMeasurementChange('irStatorMegohms', e.target.value)}
                      placeholder="VD: 850"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">IEEE 43: ≥ 100 MΩ</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Chỉ số phân cực PI (10min/1min)</label>
                    <input
                      type="number"
                      step="0.1"
                      value={measurements.polarizationIndexPi || ''}
                      onChange={(e) => handleMeasurementChange('polarizationIndexPi', e.target.value)}
                      placeholder="VD: 3.2"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">Chuẩn Class F: ≥ 2.0</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Mất cân bằng dòng điện (%)</label>
                    <input
                      type="number"
                      step="0.1"
                      value={measurements.currentImbalance || ''}
                      onChange={(e) => handleMeasurementChange('currentImbalance', e.target.value)}
                      placeholder="VD: 1.2"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">Tiêu chuẩn: ≤ 5.0%</span>
                  </div>
                </div>
              </div>
            )}

            {/* TYPE 4: MÁY PHÁT ĐIỆN (ENGINE GENERATOR) */}
            {activeCategory === 'Máy phát' && (
              <div className="space-y-4 animate-in fade-in duration-200">
                <div className="p-3 bg-amber-50/70 border border-amber-200 rounded-xl text-xs text-amber-900 flex items-center justify-between">
                  <span className="font-bold">Schema Máy phát điện dự phòng (NETA Sec 7.22 & NFPA 110)</span>
                  {onNavigateToChecklist && (
                    <button
                      type="button"
                      onClick={() => onNavigateToChecklist('neta-generator')}
                      className="text-amber-700 hover:text-amber-900 underline font-semibold flex items-center gap-1"
                    >
                      Mở Checklist Máy Phát NETA 7.22 <ArrowRight size={12} />
                    </button>
                  )}
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Thời gian chuyển đổi ATS (giây)</label>
                    <input
                      type="number"
                      step="0.1"
                      value={measurements.atsTransferTimeSec || ''}
                      onChange={(e) => handleMeasurementChange('atsTransferTimeSec', e.target.value)}
                      placeholder="VD: 8.5"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-amber-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">NFPA 110 Class 10: ≤ 10.0s</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Áp suất dầu nhớt (bar)</label>
                    <input
                      type="number"
                      step="0.1"
                      value={measurements.oilPressureBar || ''}
                      onChange={(e) => handleMeasurementChange('oilPressureBar', e.target.value)}
                      placeholder="VD: 4.8"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-amber-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">Đạt: 3.5 - 6.0 bar</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Nhiệt độ nước làm mát (°C)</label>
                    <input
                      type="number"
                      value={measurements.coolantTempC || ''}
                      onChange={(e) => handleMeasurementChange('coolantTempC', e.target.value)}
                      placeholder="VD: 82"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-amber-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">Bình thường: 75°C - 90°C</span>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">IR Stator đầu phát (MΩ)</label>
                    <input
                      type="number"
                      value={measurements.insulationStator || ''}
                      onChange={(e) => handleMeasurementChange('insulationStator', e.target.value)}
                      placeholder="VD: 1250"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-amber-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">NETA Table 100.1 (≥ 100 MΩ)</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Tần số vận hành đầy tải (Hz)</label>
                    <input
                      type="number"
                      step="0.01"
                      value={measurements.frequencyHz || ''}
                      onChange={(e) => handleMeasurementChange('frequencyHz', e.target.value)}
                      placeholder="VD: 50.05"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-amber-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">Chuẩn: 49.5 - 50.5 Hz</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Điện áp ắc quy đề (VDC)</label>
                    <input
                      type="number"
                      step="0.1"
                      value={measurements.batteryVoltageV || ''}
                      onChange={(e) => handleMeasurementChange('batteryVoltageV', e.target.value)}
                      placeholder="VD: 26.8"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-amber-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">Hệ 24V: ≥ 25.5 VDC Float</span>
                  </div>
                </div>
              </div>
            )}

            {/* TYPE 5: TỦ ĐIỆN PHÂN PHỐI & ĐÓNG CẮT (SWITCHGEAR) */}
            {activeCategory === 'Tủ điện' && (
              <div className="space-y-4 animate-in fade-in duration-200">
                <div className="p-3 bg-violet-50/70 border border-violet-200 rounded-xl text-xs text-violet-900 flex items-center justify-between">
                  <span className="font-bold">Schema Tủ điện đóng cắt Switchgear (NETA Mục 7.1.1 & TEV Partial Discharge)</span>
                  {onNavigateToChecklist && (
                    <button
                      type="button"
                      onClick={() => onNavigateToChecklist('neta-switchgear')}
                      className="text-violet-700 hover:text-violet-900 underline font-semibold flex items-center gap-1"
                    >
                      Mở Checklist Tủ Điện NETA 7.1 <ArrowRight size={12} />
                    </button>
                  )}
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Xung TEV PD Cực Đại (dB)</label>
                    <input
                      type="number"
                      value={measurements.tevDbMax || ''}
                      onChange={(e) => handleMeasurementChange('tevDbMax', e.target.value)}
                      placeholder="VD: 12"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-violet-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">An toàn: &lt; 20 dB (Cảnh báo: 20-29dB)</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Sóng siêu âm bề mặt (dBµV)</label>
                    <input
                      type="number"
                      value={measurements.ultrasonicDbuv || ''}
                      onChange={(e) => handleMeasurementChange('ultrasonicDbuv', e.target.value)}
                      placeholder="VD: 8"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-violet-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">An toàn: &lt; 10 dBµV</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Nhiệt độ mối nối thanh cái (°C)</label>
                    <input
                      type="number"
                      value={measurements.maxBusbarTemp || ''}
                      onChange={(e) => handleMeasurementChange('maxBusbarTemp', e.target.value)}
                      placeholder="VD: 48.5"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-violet-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">ΔT so với môi trường ≤ 25°C</span>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Điện trở mối nối Bu-lông Busbar (µΩ)</label>
                    <input
                      type="number"
                      value={measurements.contactResBusbar || ''}
                      onChange={(e) => handleMeasurementChange('contactResBusbar', e.target.value)}
                      placeholder="VD: 16.5"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-violet-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">NETA 7.1.1.D.1 (Độ lệch ≤ 50%)</span>
                  </div>
                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">Cách điện mạch điều khiển nhị thứ (MΩ)</label>
                    <input
                      type="number"
                      value={measurements.controlWiringIr || ''}
                      onChange={(e) => handleMeasurementChange('controlWiringIr', e.target.value)}
                      placeholder="VD: 38"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold focus:bg-white focus:ring-2 focus:ring-violet-500 outline-none"
                    />
                    <span className="text-[10px] text-slate-400">NETA 7.1.1.D.4: Bắt buộc ≥ 2.0 MΩ</span>
                  </div>
                </div>
              </div>
            )}

            {/* FSE Inspection Notes */}
            <div>
              <label className="block text-xs font-bold text-slate-700 mb-1.5 uppercase tracking-wider">
                Nhận xét & Khuyến nghị kỹ thuật của FSE
              </label>
              <textarea
                rows={2}
                value={fseNotes}
                onChange={(e) => setFseNotes(e.target.value)}
                placeholder="Nhập ghi chú hiện trường: Tình trạng ngoại quan, bu-lông, độ ẩm, kiến nghị lịch bảo dưỡng tiếp theo..."
                className="w-full p-3 bg-slate-50 border border-slate-300 rounded-xl text-xs sm:text-sm text-slate-800 outline-none focus:ring-2 focus:ring-blue-500"
              />
            </div>

            {/* SUBMIT BUTTON BAR */}
            <div className="pt-3 border-t border-slate-200 flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3">
              <div className="text-xs text-slate-500 flex items-center gap-1.5">
                <CheckCircle2 size={15} className="text-emerald-500 shrink-0" />
                <span>
                  Sẽ ghi vào Sheet <strong className="text-slate-800">TEV Service Flatform</strong> & cập nhật sheet <strong className="text-slate-800">thietbi</strong>
                </span>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handleSaveInspection}
                  disabled={isSaving}
                  className="w-full sm:w-auto px-5 py-2.5 rounded-xl font-bold text-xs sm:text-sm bg-emerald-600 hover:bg-emerald-500 text-white flex items-center justify-center gap-2 shadow-md transition-all cursor-pointer disabled:opacity-50"
                >
                  <Save size={16} />
                  <span>{isSaving ? 'Đang Lưu & Đồng Bộ...' : 'Lưu Biên Bản & Đồng Bộ Lên Sheets'}</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* MODAL 1: TẠO MỚI KHÁCH HÀNG (QUICK CREATE CUSTOMER) */}
      {showAddCustomerModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/60 backdrop-blur-xs animate-in fade-in">
          <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-lg overflow-hidden animate-in zoom-in-95">
            <div className="p-4 bg-blue-600 text-white flex items-center justify-between">
              <h3 className="font-bold text-sm sm:text-base flex items-center gap-2">
                <Plus size={18} />
                Thêm Khách Hàng Mới (Tự động ghi vào sheet khachhang)
              </h3>
              <button 
                type="button"
                onClick={() => setShowAddCustomerModal(false)}
                className="text-white/80 hover:text-white"
              >
                ✕
              </button>
            </div>

            <form onSubmit={handleCreateCustomer} className="p-5 space-y-3.5 text-xs">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-slate-700 mb-1">Mã khách hàng</label>
                  <input
                    type="text"
                    required
                    value={newCustomerForm.id || ''}
                    onChange={(e) => setNewCustomerForm({ ...newCustomerForm, id: e.target.value })}
                    placeholder="VD: KH-SEVT-03"
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-semibold"
                  />
                </div>
                <div>
                  <label className="block font-bold text-slate-700 mb-1">Số điện thoại</label>
                  <input
                    type="text"
                    value={newCustomerForm.phone || ''}
                    onChange={(e) => setNewCustomerForm({ ...newCustomerForm, phone: e.target.value })}
                    placeholder="VD: 024.3888.9999"
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg"
                  />
                </div>
              </div>

              <div>
                <label className="block font-bold text-slate-700 mb-1">Tên khách hàng / Đơn vị *</label>
                <input
                  type="text"
                  required
                  value={newCustomerForm.name || ''}
                  onChange={(e) => setNewCustomerForm({ ...newCustomerForm, name: e.target.value })}
                  placeholder="VD: Nhà máy Bia Heineken Tiền Giang"
                  className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-bold text-slate-900"
                />
              </div>

              <div>
                <label className="block font-bold text-slate-700 mb-1">Danh sách Nhà máy / Phân xưởng (phân cách bằng dấu phẩy)</label>
                <input
                  type="text"
                  value={Array.isArray(newCustomerForm.factories) ? newCustomerForm.factories.join(', ') : (newCustomerForm.factories || '')}
                  onChange={(e) => setNewCustomerForm({ ...newCustomerForm, factories: e.target.value as any })}
                  placeholder="VD: Nhà máy 1, Phân xưởng Động lực, Trạm 110kV"
                  className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-slate-700 mb-1">Email liên hệ</label>
                  <input
                    type="email"
                    value={newCustomerForm.email || ''}
                    onChange={(e) => setNewCustomerForm({ ...newCustomerForm, email: e.target.value })}
                    placeholder="electrical@client.com"
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg"
                  />
                </div>
                <div>
                  <label className="block font-bold text-slate-700 mb-1">Người phụ trách / Chức vụ</label>
                  <input
                    type="text"
                    value={newCustomerForm.contactPerson || ''}
                    onChange={(e) => setNewCustomerForm({ ...newCustomerForm, contactPerson: e.target.value })}
                    placeholder="KS Trưởng Nguyễn Văn B"
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg"
                  />
                </div>
              </div>

              <div>
                <label className="block font-bold text-slate-700 mb-1">Địa chỉ nhà máy</label>
                <input
                  type="text"
                  value={newCustomerForm.address || ''}
                  onChange={(e) => setNewCustomerForm({ ...newCustomerForm, address: e.target.value })}
                  placeholder="KCN Mỹ Tho, Tỉnh Tiền Giang"
                  className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg"
                />
              </div>

              <div className="pt-3 border-t border-slate-200 flex items-center justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setShowAddCustomerModal(false)}
                  className="px-4 py-2 border border-slate-300 rounded-lg font-bold text-slate-600 hover:bg-slate-50"
                >
                  Hủy
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-bold flex items-center gap-1.5 shadow-sm"
                >
                  <Save size={14} />
                  Lưu & Đồng bộ Sheet khachhang
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL 2: TẠO MỚI THIẾT BỊ THỦ CÔNG TẠI SITE (QUICK CREATE EQUIPMENT) */}
      {showAddEquipmentModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/60 backdrop-blur-xs animate-in fade-in">
          <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-lg overflow-hidden animate-in zoom-in-95">
            <div className="p-4 bg-emerald-700 text-white flex items-center justify-between">
              <h3 className="font-bold text-sm sm:text-base flex items-center gap-2">
                <Plus size={18} />
                Thêm Thiết Bị Mới Tại Site (Ghi vào sheet thietbi)
              </h3>
              <button 
                type="button"
                onClick={() => setShowAddEquipmentModal(false)}
                className="text-white/80 hover:text-white"
              >
                ✕
              </button>
            </div>

            <form onSubmit={handleCreateEquipmentManual} className="p-5 space-y-3.5 text-xs">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-slate-700 mb-1">Mã thiết bị (Tag) *</label>
                  <input
                    type="text"
                    required
                    value={newEqForm.id || ''}
                    onChange={(e) => setNewEqForm({ ...newEqForm, id: e.target.value })}
                    placeholder="VD: TR-110-02"
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-mono font-bold"
                  />
                </div>
                <div>
                  <label className="block font-bold text-slate-700 mb-1">Chủng loại thiết bị *</label>
                  <select
                    value={newEqForm.type || activeCategory}
                    onChange={(e) => setNewEqForm({ ...newEqForm, type: e.target.value as any })}
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-semibold"
                  >
                    <option value="Máy biến áp">Máy biến áp</option>
                    <option value="Máy cắt">Máy cắt (ACB / VCB / SF6)</option>
                    <option value="Động cơ">Động cơ điện</option>
                    <option value="Máy phát">Máy phát điện</option>
                    <option value="Tủ điện">Tủ điện (Switchgear)</option>
                  </select>
                </div>
              </div>

              <div>
                <label className="block font-bold text-slate-700 mb-1">Tên thiết bị *</label>
                <input
                  type="text"
                  required
                  value={newEqForm.name || ''}
                  onChange={(e) => setNewEqForm({ ...newEqForm, name: e.target.value })}
                  placeholder="VD: Máy biến áp phân phối T2 2000kVA 22/0.4kV"
                  className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-bold text-slate-900"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-slate-700 mb-1">Nhà máy / Phân xưởng</label>
                  <input
                    type="text"
                    value={newEqForm.factory || currentCustomer?.factories?.[0] || ''}
                    onChange={(e) => setNewEqForm({ ...newEqForm, factory: e.target.value })}
                    placeholder="VD: Phân xưởng Hàn"
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg"
                  />
                </div>
                <div>
                  <label className="block font-bold text-slate-700 mb-1">Vị trí / Ngăn lộ</label>
                  <input
                    type="text"
                    value={newEqForm.location || ''}
                    onChange={(e) => setNewEqForm({ ...newEqForm, location: e.target.value })}
                    placeholder="VD: Trạm 110kV Ngăn lộ 171"
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg"
                  />
                </div>
              </div>

              <div className="grid grid-cols-3 gap-3">
                <div>
                  <label className="block font-bold text-slate-700 mb-1">Hãng sản xuất</label>
                  <input
                    type="text"
                    value={newEqForm.manufacturer || ''}
                    onChange={(e) => setNewEqForm({ ...newEqForm, manufacturer: e.target.value })}
                    placeholder="ABB / Schneider"
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg"
                  />
                </div>
                <div>
                  <label className="block font-bold text-slate-700 mb-1">Model / Kiểu</label>
                  <input
                    type="text"
                    value={newEqForm.model || ''}
                    onChange={(e) => setNewEqForm({ ...newEqForm, model: e.target.value })}
                    placeholder="VD: NW32 H1"
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg"
                  />
                </div>
                <div>
                  <label className="block font-bold text-slate-700 mb-1">Số Serial</label>
                  <input
                    type="text"
                    value={newEqForm.serialNumber || ''}
                    onChange={(e) => setNewEqForm({ ...newEqForm, serialNumber: e.target.value })}
                    placeholder="VD: 2024-8841"
                    className="w-full p-2.5 bg-slate-50 border border-slate-300 rounded-lg font-mono"
                  />
                </div>
              </div>

              <div className="pt-3 border-t border-slate-200 flex items-center justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setShowAddEquipmentModal(false)}
                  className="px-4 py-2 border border-slate-300 rounded-lg font-bold text-slate-600 hover:bg-slate-50"
                >
                  Hủy
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg font-bold flex items-center gap-1.5 shadow-sm"
                >
                  <Save size={14} />
                  Lưu & Đồng bộ Sheet thietbi
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL 3: NAMEPLATE CAMERA AI SCANNER */}
      {showScannerModal && (
        <NameplateScannerModal
          isOpen={showScannerModal}
          onClose={() => setShowScannerModal(false)}
          onApplyData={handleApplyScanData}
          currentTag={currentEquipment?.name || activeCategory}
        />
      )}
    </div>
  );
};
