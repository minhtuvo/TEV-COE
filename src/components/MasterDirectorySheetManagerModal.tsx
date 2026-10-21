import React, { useState, useEffect } from 'react';
import {
  FileSpreadsheet,
  X,
  Plus,
  Users,
  Cpu,
  UserCheck,
  Check,
  Download,
  Upload,
  Sparkles,
  Search,
  ExternalLink,
  ShieldCheck,
  Layers,
  Building2,
  RefreshCw,
  Mail,
  Phone,
  Tag,
  CheckCircle2,
  FolderPlus,
  Table,
  Sliders,
  HelpCircle,
  Briefcase,
  Edit3,
  Trash2,
  Calendar,
  MapPin,
  ListChecks,
  AlertCircle,
  CheckSquare,
  Square,
  ChevronRight,
  ClipboardList,
  Maximize2,
  Minimize2,
  Zap,
  Play,
  ArrowRight,
  Clock,
  Droplets
} from 'lucide-react';
import { CustomerRecord, EquipmentMasterRecord } from '../utils/fieldSyncService';
import { FseEngineer } from './FseEngineerDropdown';
import {
  getRegisteredCustomers,
  saveCustomerRecord,
  deleteCustomerRecord,
  getNextCustomerId,
  getRegisteredEquipment,
  saveEquipmentRecord,
  deleteEquipmentRecord,
  getRegisteredFseList,
  saveFseEngineerRecord,
  deleteFseEngineerRecord,
  getRegisteredProjects,
  saveProjectRecord,
  deleteProjectRecord,
  getNextProjectId,
  ProjectMasterRecord,
  ProjectEquipmentItem,
  SHEET_TAB_NAMES,
  SHEET_FSE_HEADERS,
  SHEET_PROJECT_HEADERS,
  fseToSheetRow,
  projectToSheetRow,
  EVT_CUSTOMERS_UPDATED,
  EVT_EQUIPMENT_UPDATED,
  EVT_FSE_UPDATED,
  EVT_PROJECTS_UPDATED,
  EVT_NAVIGATE_TO_NETA_FORM,
  resolveEquipmentTypeToSpecId
} from '../services/masterDirectoryService';
import { ProjectCategorizedEquipmentTable } from './ProjectCategorizedEquipmentTable';
import { ProjectPortfolioGanttTimeline } from './ProjectPortfolioGanttTimeline';
import { calculatePmpProjectHealth } from '../utils/pmpStandard';

interface MasterDirectorySheetManagerModalProps {
  isOpen: boolean;
  onClose: () => void;
  isGoogleConnected: boolean;
  googleSheetUrl?: string;
  onSyncAll?: () => void;
  initialTab?: 'fse' | 'khachhang' | 'thietbi' | 'duan' | 'create_tab';
  onNavigateToNetaForm?: (project: ProjectMasterRecord, equipment: ProjectEquipmentItem, standard: 'ats' | 'mts' | 'ecs') => void;
  onSelectProjectForWorkOrder?: (project: ProjectMasterRecord) => void;
}

export const MasterDirectorySheetManagerModal: React.FC<MasterDirectorySheetManagerModalProps> = ({
  isOpen,
  onClose,
  isGoogleConnected,
  googleSheetUrl,
  onSyncAll,
  initialTab = 'duan',
  onNavigateToNetaForm,
  onSelectProjectForWorkOrder
}) => {
  const [activeTab, setActiveTab] = useState<'fse' | 'khachhang' | 'thietbi' | 'duan' | 'create_tab'>(initialTab);
  const [customers, setCustomers] = useState<CustomerRecord[]>(() => getRegisteredCustomers());
  const [equipment, setEquipment] = useState<EquipmentMasterRecord[]>(() => getRegisteredEquipment());
  const [engineers, setEngineers] = useState<FseEngineer[]>(() => getRegisteredFseList());
  const [projects, setProjects] = useState<ProjectMasterRecord[]>(() => getRegisteredProjects());

  const [search, setSearch] = useState('');
  const [isCreatingTab, setIsCreatingTab] = useState(false);
  const [statusMsg, setStatusMsg] = useState<{ type: 'success' | 'error' | 'info'; text: string } | null>(null);

  // Fullscreen Workspace Mode (default true for spacious spreadsheet management)
  const [isFullscreen, setIsFullscreen] = useState<boolean>(true);

  // Track ATS vs MTS vs ECS standard selection per equipment
  const [equipmentStandardMap, setEquipmentStandardMap] = useState<Record<string, 'ats' | 'mts' | 'ecs'>>({});

  const handleNavigateToNeta = (
    project: ProjectMasterRecord,
    equipment: ProjectEquipmentItem,
    overrideStandard?: 'ats' | 'mts' | 'ecs'
  ) => {
    const key = `${project.id}_${equipment.kksCode}`;
    const standard = overrideStandard || equipmentStandardMap[key] || project.standard || 'ats';

    if (onNavigateToNetaForm) {
      onNavigateToNetaForm(project, equipment, standard);
    } else {
      window.dispatchEvent(
        new CustomEvent(EVT_NAVIGATE_TO_NETA_FORM, {
          detail: { project, equipment, standard }
        })
      );
      onClose();
    }
  };

  // Custom sheet tab creation states
  const [customTabName, setCustomTabName] = useState('');
  const [customTabHeaderInput, setCustomTabHeaderInput] = useState('');
  const [selectedPreset, setSelectedPreset] = useState<'fse' | 'khachhang' | 'thietbi' | 'duan' | 'custom'>('fse');
  const [createdTabsList, setCreatedTabsList] = useState<string[]>(() => {
    try {
      const saved = localStorage.getItem('tev_created_custom_sheet_tabs');
      return saved ? JSON.parse(saved) : ['khachhang', 'thietbi', 'FSE', 'Dự án'];
    } catch {
      return ['khachhang', 'thietbi', 'FSE', 'Dự án'];
    }
  });

  // Modal display states for Add / Edit
  const [showAddCustomer, setShowAddCustomer] = useState(false);
  const [editingCustomer, setEditingCustomer] = useState<CustomerRecord | null>(null);

  const [showAddEquipment, setShowAddEquipment] = useState(false);
  const [editingEquipment, setEditingEquipment] = useState<EquipmentMasterRecord | null>(null);

  const [showAddFse, setShowAddFse] = useState(false);
  const [editingFse, setEditingFse] = useState<FseEngineer | null>(null);

  const [showAddProject, setShowAddProject] = useState(false);
  const [editingProject, setEditingProject] = useState<ProjectMasterRecord | null>(null);
  const [managingProjectEquipment, setManagingProjectEquipment] = useState<ProjectMasterRecord | null>(null);

  // Form states for new customer (strictly KHXXX)
  const [newCust, setNewCust] = useState({ id: '', name: '', factories: '', email: '', phone: '', address: '' });
  
  // Form states for new equipment (with customer & factory dropdowns)
  const [newEq, setNewEq] = useState({
    id: '',
    name: '',
    type: 'Máy biến áp',
    customer: '',
    factory: '',
    customFactory: '',
    location: '',
    specs: ''
  });

  // Form states for new FSE engineer
  const [newFse, setNewFse] = useState({
    name: '',
    email: '',
    role: 'Kỹ sư Thí nghiệm Hiện trường NETA',
    phone: '',
    company: 'TEV Field Service'
  });

  // Form states for project
  const [newPrj, setNewPrj] = useState({
    id: '',
    name: '',
    customerId: '',
    customerName: '',
    factory: '',
    leadFse: '',
    standard: 'ats' as 'ats' | 'mts' | 'ecs',
    startDate: new Date().toISOString().split('T')[0],
    endDate: '',
    planningStartDate: new Date().toISOString().split('T')[0],
    planningEndDate: '',
    notes: '',
    status: 'chuẩn_bị' as 'chuẩn_bị' | 'đang_thực_hiện' | 'hoàn_thành'
  });

  // Project equipment quick add inside project modal
  const [newPrjEq, setNewPrjEq] = useState({
    kksCode: '',
    name: '',
    type: 'Máy biến áp',
    specs: '',
    location: ''
  });
  const [showMasterEqPicker, setShowMasterEqPicker] = useState(false);

  // Sync state with cross-component events
  useEffect(() => {
    const handleCustUpdate = () => setCustomers(getRegisteredCustomers());
    const handleEqUpdate = () => setEquipment(getRegisteredEquipment());
    const handleFseUpdate = () => setEngineers(getRegisteredFseList());
    const handlePrjUpdate = () => setProjects(getRegisteredProjects());

    window.addEventListener(EVT_CUSTOMERS_UPDATED, handleCustUpdate);
    window.addEventListener(EVT_EQUIPMENT_UPDATED, handleEqUpdate);
    window.addEventListener(EVT_FSE_UPDATED, handleFseUpdate);
    window.addEventListener(EVT_PROJECTS_UPDATED, handlePrjUpdate);

    return () => {
      window.removeEventListener(EVT_CUSTOMERS_UPDATED, handleCustUpdate);
      window.removeEventListener(EVT_EQUIPMENT_UPDATED, handleEqUpdate);
      window.removeEventListener(EVT_FSE_UPDATED, handleFseUpdate);
      window.removeEventListener(EVT_PROJECTS_UPDATED, handlePrjUpdate);
    };
  }, []);

  // Set default customer / factory for equipment add modal
  useEffect(() => {
    if (showAddEquipment && customers.length > 0 && !newEq.customer) {
      const firstCust = customers[0];
      setNewEq(prev => ({
        ...prev,
        customer: `${firstCust.id} - ${firstCust.name}`,
        factory: (firstCust.factories && firstCust.factories[0]) || 'Nhà máy chính'
      }));
    }
  }, [showAddEquipment, customers]);

  if (!isOpen) return null;

  // Selected customer object for equipment add modal factory resolution
  const selectedCustObj = customers.find(c => 
    newEq.customer.includes(c.id) || newEq.customer.includes(c.name)
  );

  // ---------------------------------------------------------------------------
  // 1. GOOGLE SHEET TAB CREATION
  // ---------------------------------------------------------------------------
  const handleCreateSheetTab = async (tabName: string, customHeaders?: string[]) => {
    if (!isGoogleConnected) {
      setStatusMsg({
        type: 'info',
        text: 'Vui lòng kết nối Google Drive trước để tạo Sheet tab trực tiếp trên file Google Sheets.'
      });
      return;
    }

    const cleanTab = tabName.trim();
    if (!cleanTab) {
      setStatusMsg({ type: 'error', text: 'Vui lòng nhập tên sheet tab muốn tạo' });
      return;
    }

    setIsCreatingTab(true);
    setStatusMsg(null);

    try {
      const spreadsheetId = localStorage.getItem('tev_spreadsheet_id') || '';
      let headers: string[] = customHeaders || [];
      let initialRows: any[][] = [];

      if (cleanTab.toLowerCase() === 'fse' || cleanTab.toLowerCase().includes('kỹ sư')) {
        headers = SHEET_FSE_HEADERS;
        initialRows = engineers.map((eng) => fseToSheetRow(eng));
      } else if (cleanTab.toLowerCase() === 'khachhang' || cleanTab.toLowerCase().includes('khách hàng')) {
        headers = ['Mã khách hàng', 'Tên khách hàng', 'Nhà máy / Trạm', 'Email', 'Số điện thoại', 'Địa chỉ', 'Ngày tạo', 'Ngày cập nhật'];
        initialRows = customers.map((c) => [
          c.id,
          c.name,
          Array.isArray(c.factories) ? c.factories.join(', ') : (c.factories || ''),
          c.email || '',
          c.phone || '',
          c.address || '',
          c.createdAt || new Date().toISOString(),
          new Date().toISOString()
        ]);
      } else if (cleanTab.toLowerCase() === 'thietbi' || cleanTab.toLowerCase().includes('thiết bị')) {
        headers = ['Mã thiết bị (Tag ID)', 'Tên thiết bị', 'Loại thiết bị', 'Khách hàng', 'Nhà máy', 'Vị trí ngăn lộ', 'Trạng thái', 'Sức khỏe HI (%)', 'Thông số kỹ thuật', 'Thông tin nhãn Nameplate', 'Kiểm tra gần nhất', 'Ngày tạo', 'Ngày cập nhật'];
        initialRows = equipment.map((eq) => [
          eq.id,
          eq.name,
          eq.type,
          eq.customer,
          eq.factory,
          eq.location || '',
          eq.status || 'healthy',
          eq.health !== undefined ? eq.health : 95,
          typeof eq.specs === 'object' ? JSON.stringify(eq.specs) : (eq.specs || eq.notes || ''),
          typeof eq.nameplate === 'object' ? JSON.stringify(eq.nameplate) : (eq.nameplate || ''),
          eq.lastCheck || new Date().toLocaleDateString('vi-VN'),
          new Date().toISOString(),
          new Date().toISOString()
        ]);
      } else if (cleanTab.toLowerCase() === 'duan' || cleanTab.toLowerCase().includes('dự án') || cleanTab.toLowerCase() === 'project') {
        headers = SHEET_PROJECT_HEADERS;
        initialRows = projects.map(p => projectToSheetRow(p));
      } else if (!headers || headers.length === 0) {
        headers = ['Mã định danh', 'Tên bản ghi', 'Phân loại', 'Người thực hiện / FSE', 'Khách hàng', 'Ghi chú', 'Ngày tạo'];
      }

      const response = await fetch('/api/sheets/create-tab', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${localStorage.getItem('google_tokens') || ''}`,
          ...(spreadsheetId ? { 'x-spreadsheet-id': spreadsheetId } : {})
        },
        body: JSON.stringify({
          tabName: cleanTab,
          headers,
          initialRows
        })
      });

      const resData = await response.json();
      if (!response.ok) {
        throw new Error(resData.error || 'Lỗi khi tạo sheet tab');
      }

      // Add to created list
      if (!createdTabsList.includes(cleanTab)) {
        const nextList = [...createdTabsList, cleanTab];
        setCreatedTabsList(nextList);
        try {
          localStorage.setItem('tev_created_custom_sheet_tabs', JSON.stringify(nextList));
        } catch {}
      }

      setStatusMsg({
        type: 'success',
        text: resData.message || `Đã tạo thành công Sheet tab "${cleanTab}" trên Google Sheets!`
      });
      setCustomTabName('');
      setCustomTabHeaderInput('');
    } catch (err: any) {
      setStatusMsg({
        type: 'error',
        text: `Lỗi: ${err.message}`
      });
    } finally {
      setIsCreatingTab(false);
    }
  };

  // ---------------------------------------------------------------------------
  // 2. FSE ENGINEERS (Add, Edit, Delete)
  // ---------------------------------------------------------------------------
  const handleOpenAddFse = () => {
    setNewFse({
      name: '',
      email: '',
      role: 'Kỹ sư Thí nghiệm Hiện trường NETA',
      phone: '',
      company: 'TEV Field Service'
    });
    setEditingFse(null);
    setShowAddFse(true);
  };

  const handleEditFseClick = (eng: FseEngineer) => {
    setEditingFse(eng);
    setNewFse({
      name: eng.name,
      email: eng.email,
      role: eng.role || 'Kỹ sư Thí nghiệm Hiện trường NETA',
      phone: eng.phone || '',
      company: eng.company || 'TEV Field Service'
    });
    setShowAddFse(true);
  };

  const handleDeleteFseClick = (eng: FseEngineer) => {
    if (eng.email === 'sgm1707@gmail.com') {
      alert('Không thể xóa tài khoản Kỹ sư trưởng hệ thống (sgm1707@gmail.com)');
      return;
    }
    if (confirm(`Bạn có chắc chắn muốn xóa Kỹ sư FSE "${eng.name}" (${eng.email}) khỏi danh bạ?`)) {
      const updated = deleteFseEngineerRecord(eng.id || eng.email);
      setEngineers(updated);
      setStatusMsg({ type: 'success', text: `Đã xóa Kỹ sư FSE "${eng.name}" khỏi danh bạ.` });
    }
  };

  const handleSaveFse = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newFse.name.trim() || !newFse.email.trim()) return;

    const engineerPayload: FseEngineer = {
      id: editingFse ? editingFse.id : `fse-${Date.now()}`,
      name: newFse.name.trim(),
      email: newFse.email.trim().toLowerCase(),
      role: newFse.role.trim() || 'Kỹ sư Thí nghiệm Hiện trường NETA',
      phone: newFse.phone.trim(),
      company: newFse.company.trim() || 'TEV Field Service'
    };

    const updated = saveFseEngineerRecord(engineerPayload);
    setEngineers(updated);
    setShowAddFse(false);
    setEditingFse(null);
    setStatusMsg({
      type: 'success',
      text: editingFse 
        ? `Đã cập nhật thông tin Kỹ Sư FSE "${engineerPayload.name}" thành công!`
        : `Đã bổ sung Kỹ Sư FSE "${engineerPayload.name}" (${engineerPayload.email}) vào danh mục!`
    });
  };

  // ---------------------------------------------------------------------------
  // 3. CUSTOMER (KHXXX strict format, Add, Edit, Delete)
  // ---------------------------------------------------------------------------
  const handleOpenAddCustomer = () => {
    const nextCode = getNextCustomerId(customers);
    setNewCust({
      id: nextCode,
      name: '',
      factories: '',
      email: '',
      phone: '',
      address: ''
    });
    setEditingCustomer(null);
    setShowAddCustomer(true);
  };

  const handleEditCustomerClick = (cust: CustomerRecord) => {
    setEditingCustomer(cust);
    setNewCust({
      id: cust.id,
      name: cust.name,
      factories: Array.isArray(cust.factories) ? cust.factories.join(', ') : (cust.factories || ''),
      email: cust.email || '',
      phone: cust.phone || '',
      address: cust.address || ''
    });
    setShowAddCustomer(true);
  };

  const handleDeleteCustomerClick = (cust: CustomerRecord) => {
    if (confirm(`Bạn có chắc chắn muốn xóa Khách hàng "${cust.name}" (Mã: ${cust.id})?`)) {
      const updated = deleteCustomerRecord(cust.id);
      setCustomers(updated);
      setStatusMsg({ type: 'success', text: `Đã xóa Khách hàng "${cust.name}" (${cust.id}).` });
    }
  };

  const handleSaveCustomer = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newCust.name.trim()) return;

    let cleanId = newCust.id.trim();
    if (!cleanId) {
      cleanId = getNextCustomerId(customers);
    } else {
      cleanId = cleanId.toUpperCase();
      if (!cleanId.startsWith('KH')) {
        const digits = cleanId.replace(/\D/g, '');
        cleanId = `KH${(digits || '1').padStart(3, '0')}`;
      }
    }

    const factoriesList = newCust.factories.split(',').map(f => f.trim()).filter(f => f.length > 0);
    const customerPayload: CustomerRecord = {
      id: cleanId,
      name: newCust.name.trim(),
      factories: factoriesList.length > 0 ? factoriesList : ['Nhà máy chính'],
      email: newCust.email.trim(),
      phone: newCust.phone.trim(),
      address: newCust.address.trim(),
      createdAt: editingCustomer?.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    const updated = saveCustomerRecord(customerPayload);
    setCustomers(updated);
    setShowAddCustomer(false);
    setEditingCustomer(null);
    setStatusMsg({
      type: 'success',
      text: editingCustomer
        ? `Đã cập nhật thông tin Khách Hàng "${customerPayload.name}" (${customerPayload.id})!`
        : `Đã bổ sung Khách Hàng "${customerPayload.name}" (${customerPayload.id}) vào danh mục!`
    });
  };

  // ---------------------------------------------------------------------------
  // 4. EQUIPMENT (Add with Customer & Factory dropdowns, Edit, Delete)
  // ---------------------------------------------------------------------------
  const handleOpenAddEquipment = () => {
    const defaultCust = customers[0];
    const custStr = defaultCust ? `${defaultCust.id} - ${defaultCust.name}` : 'Công ty Cổ phần Năng lượng TEV';
    const factoryStr = (defaultCust && defaultCust.factories && defaultCust.factories[0]) || 'Nhà máy chính';

    setNewEq({
      id: '',
      name: '',
      type: 'Máy biến áp',
      customer: custStr,
      factory: factoryStr,
      customFactory: '',
      location: '',
      specs: ''
    });
    setEditingEquipment(null);
    setShowAddEquipment(true);
  };

  const handleEditEquipmentClick = (eq: EquipmentMasterRecord) => {
    setEditingEquipment(eq);
    setNewEq({
      id: eq.id,
      name: eq.name,
      type: eq.type || 'Máy biến áp',
      customer: eq.customer || '',
      factory: eq.factory || '',
      customFactory: '',
      location: eq.location || '',
      specs: typeof eq.specs === 'object' ? JSON.stringify(eq.specs) : (eq.specs || eq.notes || '')
    });
    setShowAddEquipment(true);
  };

  const handleDeleteEquipmentClick = (eq: EquipmentMasterRecord) => {
    if (confirm(`Bạn có chắc muốn xóa Thiết bị "${eq.name}" (Tag: ${eq.id}) khỏi danh mục master?`)) {
      const updated = deleteEquipmentRecord(eq.id);
      setEquipment(updated);
      setStatusMsg({ type: 'success', text: `Đã xóa thiết bị "${eq.name}" (${eq.id}).` });
    }
  };

  const handleSaveEquipment = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newEq.id.trim() || !newEq.name.trim()) return;

    // Resolve customer and factory
    let targetCustomer = newEq.customer.trim();
    if (selectedCustObj) {
      targetCustomer = selectedCustObj.name;
    }
    const targetFactory = (newEq.factory === '__custom__' ? newEq.customFactory : newEq.factory) || 'Nhà máy chính';

    const eqPayload: EquipmentMasterRecord = {
      id: newEq.id.trim().toUpperCase(),
      name: newEq.name.trim(),
      type: newEq.type as any,
      customer: targetCustomer || 'Công ty Cổ phần Năng lượng TEV',
      factory: targetFactory,
      location: newEq.location.trim() || 'Trạm phân phối',
      status: editingEquipment?.status || 'healthy',
      health: editingEquipment?.health !== undefined ? editingEquipment.health : 95,
      specs: newEq.specs ? newEq.specs : editingEquipment?.specs,
      lastCheck: editingEquipment?.lastCheck || new Date().toLocaleDateString('vi-VN')
    };

    const updated = saveEquipmentRecord(eqPayload);
    setEquipment(updated);
    setShowAddEquipment(false);
    setEditingEquipment(null);
    setStatusMsg({
      type: 'success',
      text: editingEquipment
        ? `Đã cập nhật Thiết Bị "${eqPayload.name}" (${eqPayload.id})!`
        : `Đã bổ sung Thiết Bị "${eqPayload.name}" (${eqPayload.id}) thuộc KH "${eqPayload.customer}"!`
    });
  };

  // ---------------------------------------------------------------------------
  // 5. PROJECTS & FIELD PREPARATION (Sheet 'duan')
  // ---------------------------------------------------------------------------
  const handleOpenAddProject = () => {
    const nextPrjId = getNextProjectId(projects);
    const firstCust = customers[0];
    const defaultStart = new Date().toISOString().split('T')[0];
    const defaultEnd = new Date(Date.now() + 25 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
    setNewPrj({
      id: nextPrjId,
      name: '',
      customerId: firstCust?.id || 'KH001',
      customerName: firstCust?.name || '',
      factory: (firstCust?.factories && firstCust.factories[0]) || 'Nhà máy chính',
      leadFse: engineers[0]?.name || 'Võ Minh Tú',
      standard: 'ats',
      startDate: defaultStart,
      endDate: defaultEnd,
      planningStartDate: defaultStart,
      planningEndDate: defaultEnd,
      notes: '',
      status: 'chuẩn_bị'
    });
    setEditingProject(null);
    setShowAddProject(true);
  };

  const handleEditProjectClick = (prj: ProjectMasterRecord) => {
    setEditingProject(prj);
    const pStart = prj.planningStartDate || prj.startDate || '';
    const pEnd = prj.planningEndDate || prj.endDate || '';
    const std = prj.standard || (prj.name?.toLowerCase().includes('ecs') ? 'ecs' : prj.name?.toLowerCase().includes('mts') ? 'mts' : 'ats');
    setNewPrj({
      id: prj.id,
      name: prj.name,
      customerId: prj.customerId,
      customerName: prj.customerName,
      factory: prj.factory,
      leadFse: prj.leadFse || '',
      standard: std,
      startDate: prj.startDate || pStart,
      endDate: prj.endDate || pEnd,
      planningStartDate: pStart,
      planningEndDate: pEnd,
      notes: prj.notes || '',
      status: prj.status || 'chuẩn_bị'
    });
    setShowAddProject(true);
  };

  const handleDeleteProjectClick = (prj: ProjectMasterRecord) => {
    if (confirm(`Bạn có chắc muốn xóa Dự án "${prj.name}" (Mã: ${prj.id})?`)) {
      const updated = deleteProjectRecord(prj.id);
      setProjects(updated);
      setStatusMsg({ type: 'success', text: `Đã xóa Dự án "${prj.name}" (${prj.id}).` });
    }
  };

  const handleSaveProject = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newPrj.name.trim()) return;

    const custObj = customers.find(c => c.id === newPrj.customerId || c.name === newPrj.customerName);
    const customerId = custObj ? custObj.id : (newPrj.customerId || 'KH001');
    const customerName = custObj ? custObj.name : (newPrj.customerName || 'Khách hàng');

    const prjPayload: ProjectMasterRecord = {
      id: newPrj.id.trim() || getNextProjectId(projects),
      name: newPrj.name.trim(),
      customerId,
      customerName,
      factory: newPrj.factory.trim() || 'Nhà máy chính',
      leadFse: newPrj.leadFse.trim(),
      standard: newPrj.standard,
      startDate: newPrj.startDate || newPrj.planningStartDate,
      endDate: newPrj.endDate || newPrj.planningEndDate,
      planningStartDate: newPrj.planningStartDate || newPrj.startDate,
      planningEndDate: newPrj.planningEndDate || newPrj.endDate,
      notes: newPrj.notes,
      status: newPrj.status,
      equipmentList: editingProject ? editingProject.equipmentList : [],
      createdAt: editingProject?.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    const updated = saveProjectRecord(prjPayload);
    setProjects(updated);
    setShowAddProject(false);
    setEditingProject(null);
    setStatusMsg({
      type: 'success',
      text: editingProject
        ? `Đã cập nhật Dự Án "${prjPayload.name}" (${prjPayload.id})!`
        : `Đã tạo mới Dự Án "${prjPayload.name}" (${prjPayload.id})! Bạn có thể nhấn 'Soạn Thiết Bị' để gán danh sách KKS code.`
    });
  };

  // Manage equipment within a project
  const handleOpenManageProjectEquipment = (prj: ProjectMasterRecord) => {
    setManagingProjectEquipment(prj);
    setNewPrjEq({
      kksCode: '',
      name: '',
      type: 'Máy biến áp',
      specs: '',
      location: ''
    });
  };

  const handleAddEquipmentToProject = (e: React.FormEvent) => {
    e.preventDefault();
    if (!managingProjectEquipment || !newPrjEq.name.trim()) return;

    const cleanKks = newPrjEq.kksCode.trim() || `EQ-${Date.now().toString().slice(-4)}`;
    const newItem: ProjectEquipmentItem = {
      kksCode: cleanKks,
      name: newPrjEq.name.trim(),
      type: newPrjEq.type,
      specs: newPrjEq.specs.trim(),
      location: newPrjEq.location.trim() || managingProjectEquipment.factory,
      status: 'chờ_test'
    };

    const updatedPrj: ProjectMasterRecord = {
      ...managingProjectEquipment,
      equipmentList: [...(managingProjectEquipment.equipmentList || []), newItem],
      updatedAt: new Date().toISOString()
    };

    saveProjectRecord(updatedPrj);
    setManagingProjectEquipment(updatedPrj);
    setProjects(getRegisteredProjects());
    setNewPrjEq({ kksCode: '', name: '', type: 'Máy biến áp', specs: '', location: '' });
  };

  const handleRemoveEquipmentFromProject = (kksCode: string) => {
    if (!managingProjectEquipment) return;
    const nextList = (managingProjectEquipment.equipmentList || []).filter(item => item.kksCode !== kksCode);
    const updatedPrj: ProjectMasterRecord = {
      ...managingProjectEquipment,
      equipmentList: nextList,
      updatedAt: new Date().toISOString()
    };
    saveProjectRecord(updatedPrj);
    setManagingProjectEquipment(updatedPrj);
    setProjects(getRegisteredProjects());
  };

  const handlePickFromMasterEquipment = (masterEq: EquipmentMasterRecord) => {
    if (!managingProjectEquipment) return;
    // Check if already in project
    const exists = (managingProjectEquipment.equipmentList || []).some(
      item => item.kksCode.toLowerCase() === masterEq.id.toLowerCase()
    );
    if (exists) {
      alert(`Thiết bị ${masterEq.id} đã có trong dự án.`);
      return;
    }

    const newItem: ProjectEquipmentItem = {
      kksCode: masterEq.id,
      name: masterEq.name,
      type: masterEq.type,
      specs: typeof masterEq.specs === 'object' ? JSON.stringify(masterEq.specs) : (masterEq.specs || masterEq.notes || ''),
      location: masterEq.location || masterEq.factory || '',
      status: 'chờ_test'
    };

    const updatedPrj: ProjectMasterRecord = {
      ...managingProjectEquipment,
      equipmentList: [...(managingProjectEquipment.equipmentList || []), newItem],
      updatedAt: new Date().toISOString()
    };

    saveProjectRecord(updatedPrj);
    setManagingProjectEquipment(updatedPrj);
    setProjects(getRegisteredProjects());
  };

  return (
    <div className={`fixed inset-0 z-50 bg-slate-900/70 backdrop-blur-xs flex items-center justify-center ${isFullscreen ? 'p-0' : 'p-2 sm:p-4'}`}>
      <div className={`bg-white shadow-2xl flex flex-col border border-slate-200 overflow-hidden transition-all duration-150 ${
        isFullscreen
          ? 'w-screen h-screen max-w-none max-h-none rounded-none'
          : 'max-w-7xl w-full max-h-[95vh] rounded-2xl'
      }`}>
        
        {/* Header */}
        <div className="p-4 bg-gradient-to-r from-slate-900 via-indigo-950 to-blue-950 text-white flex items-center justify-between shrink-0">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-blue-600/30 border border-blue-400/30 rounded-xl">
              <FileSpreadsheet size={24} className="text-blue-400" />
            </div>
            <div>
              <h2 className="text-base sm:text-lg font-bold flex items-center gap-2">
                Quản Lý Dự Án (Chuẩn PMP)
                <span className="text-[10px] bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 px-2 py-0.5 rounded-full font-mono font-medium">
                  2-Way Sync
                </span>
                {isFullscreen && (
                  <span className="text-[10px] bg-indigo-500/30 text-indigo-200 border border-indigo-400/30 px-2 py-0.5 rounded-full font-mono font-bold hidden sm:inline-block">
                    Toàn Màn Hình
                  </span>
                )}
              </h2>
              <p className="text-xs text-slate-300">
                Chuẩn hóa quy trình PMP: <strong>4. Soạn Dự Án & Thiết Bị Chờ Test (ATS / MST)</strong>, <strong>1. Kỹ Sư FSE</strong>, <strong>2. Mã Khách Hàng (KHXXX)</strong>, <strong>3. Thiết Bị Master</strong>.
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => setIsFullscreen(!isFullscreen)}
              className="px-2.5 py-1.5 text-slate-300 hover:text-white hover:bg-white/10 rounded-lg transition-colors flex items-center gap-1.5 text-xs font-semibold border border-slate-700/60"
              title={isFullscreen ? 'Thu nhỏ cửa sổ' : 'Mở toàn màn hình'}
            >
              {isFullscreen ? <Minimize2 size={16} /> : <Maximize2 size={16} />}
              <span className="hidden sm:inline">{isFullscreen ? 'Thu nhỏ' : 'Toàn màn hình'}</span>
            </button>
            <button
              onClick={onClose}
              className="p-1.5 text-slate-400 hover:text-white hover:bg-white/10 rounded-lg transition-colors"
              title="Đóng"
            >
              <X size={20} />
            </button>
          </div>
        </div>

        {/* Status Notification */}
        {statusMsg && (
          <div
            className={`px-4 py-2 text-xs font-semibold flex items-center justify-between shrink-0 ${
              statusMsg.type === 'success'
                ? 'bg-emerald-50 text-emerald-900 border-b border-emerald-200'
                : statusMsg.type === 'error'
                ? 'bg-rose-50 text-rose-900 border-b border-rose-200'
                : 'bg-blue-50 text-blue-900 border-b border-blue-200'
            }`}
          >
            <span className="flex items-center gap-2">
              <Sparkles size={15} />
              {statusMsg.text}
            </span>
            <button onClick={() => setStatusMsg(null)} className="text-slate-400 hover:text-slate-700">
              <X size={14} />
            </button>
          </div>
        )}

        {/* 5 Navigation Tabs */}
        <div className="p-2.5 bg-slate-100 border-b border-slate-200 flex flex-wrap items-center justify-between gap-2.5 shrink-0">
          <div className="flex items-center gap-1.5 bg-white p-1 rounded-xl border border-slate-200 shadow-2xs overflow-x-auto max-w-full">
            
            {/* Tab 1: FSE */}
            <button
              onClick={() => { setActiveTab('fse'); setSearch(''); }}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition-all whitespace-nowrap ${
                activeTab === 'fse' ? 'bg-amber-500 text-white shadow-xs' : 'text-slate-600 hover:bg-slate-100'
              }`}
            >
              <UserCheck size={14} />
              <span>1. Kỹ Sư FSE</span>
              <span className="text-[10px] px-1.5 py-0.2 rounded-full font-mono bg-white/20 text-white font-bold">
                {engineers.length}
              </span>
              <span className="text-[9px] bg-amber-100 text-amber-950 font-mono px-1 rounded">
                Sheet FSE
              </span>
            </button>

            {/* Tab 2: Khách Hàng (KHXXX) */}
            <button
              onClick={() => { setActiveTab('khachhang'); setSearch(''); }}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition-all whitespace-nowrap ${
                activeTab === 'khachhang' ? 'bg-blue-600 text-white shadow-xs' : 'text-slate-600 hover:bg-slate-100'
              }`}
            >
              <Building2 size={14} />
              <span>2. Khách Hàng (KHXXX)</span>
              <span className="text-[10px] px-1.5 py-0.2 rounded-full font-mono bg-white/20 text-white font-bold">
                {customers.length}
              </span>
              <span className="text-[9px] bg-blue-100 text-blue-900 font-mono px-1 rounded">
                Sheet khachhang
              </span>
            </button>

            {/* Tab 3: Thiết Bị */}
            <button
              onClick={() => { setActiveTab('thietbi'); setSearch(''); }}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition-all whitespace-nowrap ${
                activeTab === 'thietbi' ? 'bg-cyan-600 text-white shadow-xs' : 'text-slate-600 hover:bg-slate-100'
              }`}
            >
              <Cpu size={14} />
              <span>3. Thiết Bị Master</span>
              <span className="text-[10px] px-1.5 py-0.2 rounded-full font-mono bg-white/20 text-white font-bold">
                {equipment.length}
              </span>
              <span className="text-[9px] bg-cyan-100 text-cyan-950 font-mono px-1 rounded">
                Sheet thietbi
              </span>
            </button>

            {/* Tab 4: Dự Án & Soạn Thiết Bị Hiện Trường */}
            <button
              onClick={() => { setActiveTab('duan'); setSearch(''); }}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition-all whitespace-nowrap ${
                activeTab === 'duan' ? 'bg-indigo-600 text-white shadow-xs ring-2 ring-indigo-400' : 'text-slate-700 hover:bg-indigo-50 font-semibold'
              }`}
            >
              <Briefcase size={14} className="text-amber-300" />
              <span>4. Soạn Dự Án & KKS Hiện Trường</span>
              <span className="text-[10px] px-1.5 py-0.2 rounded-full font-mono bg-white/20 text-white font-bold">
                {projects.length}
              </span>
              <span className="text-[9px] bg-amber-400 text-slate-950 font-black px-1 rounded uppercase">
                Mới
              </span>
            </button>

            {/* Tab 5: Tạo Thêm Sheet Tab Tùy Chọn */}
            <button
              onClick={() => { setActiveTab('create_tab'); setSearch(''); }}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition-all whitespace-nowrap ${
                activeTab === 'create_tab' ? 'bg-emerald-600 text-white shadow-xs' : 'text-slate-600 hover:bg-slate-100'
              }`}
            >
              <FolderPlus size={14} />
              <span>5. Tạo Sheet Tab Mới</span>
            </button>
          </div>

          {/* Global Actions */}
          <div className="flex items-center gap-2">
            {googleSheetUrl && (
              <a
                href={googleSheetUrl}
                target="_blank"
                rel="noreferrer"
                className="px-2.5 py-1.5 text-xs text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 font-semibold flex items-center gap-1.5 shadow-2xs"
              >
                <ExternalLink size={13} />
                <span>Mở Google Sheet</span>
              </a>
            )}
            {onSyncAll && (
              <button
                onClick={onSyncAll}
                className="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-lg flex items-center gap-1.5 shadow-xs transition-colors"
              >
                <RefreshCw size={13} />
                <span>Đồng bộ 2 Chiều</span>
              </button>
            )}
          </div>
        </div>

        {/* Content Area */}
        <div className="p-4 overflow-y-auto flex-1 space-y-4">
          
          {/* ================================================================= */}
          {/* TAB 1: KỸ SƯ FSE (Với nút Sửa & Xóa) */}
          {/* ================================================================= */}
          {activeTab === 'fse' && (
            <div className="space-y-4">
              {/* Action Banner */}
              <div className="bg-gradient-to-r from-amber-50 to-orange-50 border border-amber-200 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-2xs">
                <div className="space-y-1">
                  <div className="flex items-center gap-2">
                    <span className="font-bold text-slate-900 text-sm">Sheet Tab `FSE` (Danh mục Kỹ sư Hiện trường NETA)</span>
                    <span className="text-[10px] bg-amber-200 text-amber-950 px-2 py-0.5 rounded font-mono font-bold">
                      {engineers.length} kỹ sư
                    </span>
                  </div>
                  <p className="text-xs text-slate-600">
                    Bổ sung, chỉnh sửa thông tin hoặc xóa kỹ sư FSE. Click <strong>Tạo thêm Sheet tab FSE</strong> để đồng bộ tự động lên Google Sheets.
                  </p>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <button
                    onClick={() => handleCreateSheetTab('FSE')}
                    disabled={isCreatingTab}
                    className="px-3.5 py-2 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-xs transition-all disabled:opacity-50"
                  >
                    <FolderPlus size={14} />
                    <span>{isCreatingTab ? 'Đang tạo tab...' : 'Tạo thêm Sheet tab FSE'}</span>
                  </button>
                  <button
                    onClick={handleOpenAddFse}
                    className="px-3 py-2 bg-white hover:bg-slate-50 text-slate-800 border border-slate-300 rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-2xs"
                  >
                    <Plus size={14} className="text-amber-600" />
                    <span>Bổ sung Kỹ Sư FSE</span>
                  </button>
                </div>
              </div>

              {/* Search */}
              <div className="relative">
                <Search size={14} className="absolute left-3 top-2.5 text-slate-400" />
                <input
                  type="text"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="Tìm kiếm kỹ sư FSE theo tên, email, chức danh..."
                  className="w-full pl-9 pr-3 py-2 text-xs border border-slate-200 rounded-xl focus:ring-2 focus:ring-amber-500 focus:outline-none bg-slate-50/50"
                />
              </div>

              {/* Table with Edit and Delete */}
              <div className="border border-slate-200 rounded-xl overflow-hidden shadow-2xs">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-50 border-b border-slate-200 font-bold text-slate-700">
                    <tr>
                      <th className="p-3">Họ và tên Kỹ sư</th>
                      <th className="p-3">Email</th>
                      <th className="p-3">Chức danh / Vai trò</th>
                      <th className="p-3">Số điện thoại</th>
                      <th className="p-3">Đơn vị</th>
                      <th className="p-3 text-center">Thao tác</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {engineers
                      .filter(
                        (e) =>
                          e.name.toLowerCase().includes(search.toLowerCase()) ||
                          e.email.toLowerCase().includes(search.toLowerCase()) ||
                          (e.role && e.role.toLowerCase().includes(search.toLowerCase()))
                      )
                      .map((eng) => (
                        <tr key={eng.id || eng.email} className="hover:bg-amber-50/30 transition-colors">
                          <td className="p-3">
                            <div className="flex items-center gap-2">
                              <div className="w-6 h-6 rounded-full bg-amber-600 text-white flex items-center justify-center text-[10px] font-bold">
                                {eng.name.slice(0, 2).toUpperCase()}
                              </div>
                              <span className="font-bold text-slate-900">{eng.name}</span>
                              {eng.email === 'sgm1707@gmail.com' && (
                                <span className="text-[9px] bg-amber-100 text-amber-900 px-1.5 py-0.2 rounded font-bold">
                                  Trưởng nhóm
                                </span>
                              )}
                            </div>
                          </td>
                          <td className="p-3 font-mono text-slate-600">{eng.email}</td>
                          <td className="p-3 text-slate-700 font-medium">{eng.role}</td>
                          <td className="p-3 text-slate-600">{eng.phone || '—'}</td>
                          <td className="p-3 text-slate-600">{eng.company || 'TEV Field Service'}</td>
                          <td className="p-3 text-center">
                            <div className="flex items-center justify-center gap-1.5">
                              <button
                                onClick={() => handleEditFseClick(eng)}
                                className="p-1.5 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded-lg transition-colors"
                                title="Chỉnh sửa thông tin kỹ sư FSE"
                              >
                                <Edit3 size={14} />
                              </button>
                              {eng.email !== 'sgm1707@gmail.com' ? (
                                <button
                                  onClick={() => handleDeleteFseClick(eng)}
                                  className="p-1.5 text-rose-600 hover:text-rose-800 hover:bg-rose-50 rounded-lg transition-colors"
                                  title="Xóa kỹ sư FSE khỏi danh bạ"
                                >
                                  <Trash2 size={14} />
                                </button>
                              ) : (
                                <span className="text-[10px] text-slate-400 italic">Mặc định</span>
                              )}
                            </div>
                          </td>
                        </tr>
                      ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* ================================================================= */}
          {/* TAB 2: KHÁCH HÀNG (Chuẩn KHXXX, với nút Sửa & Xóa) */}
          {/* ================================================================= */}
          {activeTab === 'khachhang' && (
            <div className="space-y-4">
              {/* Action Banner */}
              <div className="bg-gradient-to-r from-blue-50 to-indigo-50 border border-blue-200 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-2xs">
                <div className="space-y-1">
                  <div className="flex items-center gap-2">
                    <span className="font-bold text-slate-900 text-sm">Sheet Tab `khachhang` (Quy ước chuẩn mã: KHXXX)</span>
                    <span className="text-[10px] bg-blue-200 text-blue-950 px-2 py-0.5 rounded font-mono font-bold">
                      {customers.length} khách hàng
                    </span>
                  </div>
                  <p className="text-xs text-slate-600">
                    Mã khách hàng tự động định dạng chuẩn <strong>KHXXX</strong> (KH001, KH002...). Có đầy đủ chức năng Thêm mới, Chỉnh sửa thông tin và Xóa khách hàng.
                  </p>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <button
                    onClick={() => handleCreateSheetTab('khachhang')}
                    disabled={isCreatingTab}
                    className="px-3.5 py-2 bg-blue-100 hover:bg-blue-200 text-blue-900 border border-blue-300 rounded-lg text-xs font-bold flex items-center gap-1.5 transition-all disabled:opacity-50"
                  >
                    <RefreshCw size={13} className={isCreatingTab ? 'animate-spin' : ''} />
                    <span>Đồng bộ / Tạo Sheet khachhang</span>
                  </button>
                  <button
                    onClick={handleOpenAddCustomer}
                    className="px-3.5 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-sm transition-all"
                  >
                    <Plus size={14} />
                    <span>Bổ sung Khách Hàng mới</span>
                  </button>
                </div>
              </div>

              {/* Search */}
              <div className="relative">
                <Search size={14} className="absolute left-3 top-2.5 text-slate-400" />
                <input
                  type="text"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="Tìm kiếm mã KH (KH001...), tên doanh nghiệp, địa chỉ..."
                  className="w-full pl-9 pr-3 py-2 text-xs border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 focus:outline-none bg-slate-50/50"
                />
              </div>

              {/* Table with Edit and Delete */}
              <div className="border border-slate-200 rounded-xl overflow-hidden shadow-2xs">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-50 border-b border-slate-200 font-bold text-slate-700">
                    <tr>
                      <th className="p-3">Mã Khách Hàng</th>
                      <th className="p-3">Tên Khách Hàng</th>
                      <th className="p-3">Nhà máy / Phân xưởng</th>
                      <th className="p-3">Liên hệ</th>
                      <th className="p-3">Địa chỉ</th>
                      <th className="p-3 text-center">Thao tác</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {customers
                      .filter(
                        (c) =>
                          c.id.toLowerCase().includes(search.toLowerCase()) ||
                          c.name.toLowerCase().includes(search.toLowerCase()) ||
                          (c.address && c.address.toLowerCase().includes(search.toLowerCase()))
                      )
                      .map((cust) => (
                        <tr key={cust.id} className="hover:bg-blue-50/30 transition-colors">
                          <td className="p-3 font-mono font-black text-blue-700">{cust.id}</td>
                          <td className="p-3 font-bold text-slate-900">{cust.name}</td>
                          <td className="p-3 text-slate-600">
                            {cust.factories?.map((f, i) => (
                              <span key={i} className="inline-block bg-slate-100 text-slate-700 px-1.5 py-0.2 rounded mr-1 text-[10px]">
                                {f}
                              </span>
                            ))}
                          </td>
                          <td className="p-3 text-slate-600">
                            {cust.email && <div className="text-[11px] font-mono">{cust.email}</div>}
                            {cust.phone && <div className="text-[10px] text-slate-500">{cust.phone}</div>}
                          </td>
                          <td className="p-3 text-slate-500">{cust.address || '—'}</td>
                          <td className="p-3 text-center">
                            <div className="flex items-center justify-center gap-1.5">
                              <button
                                onClick={() => handleEditCustomerClick(cust)}
                                className="p-1.5 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded-lg transition-colors"
                                title="Chỉnh sửa thông tin khách hàng"
                              >
                                <Edit3 size={14} />
                              </button>
                              <button
                                onClick={() => handleDeleteCustomerClick(cust)}
                                className="p-1.5 text-rose-600 hover:text-rose-800 hover:bg-rose-50 rounded-lg transition-colors"
                                title="Xóa khách hàng"
                              >
                                <Trash2 size={14} />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* ================================================================= */}
          {/* TAB 3: THIẾT BỊ MASTER (Có Dropdown Khách Hàng, Nhà Máy, Sửa & Xóa) */}
          {/* ================================================================= */}
          {activeTab === 'thietbi' && (
            <div className="space-y-4">
              {/* Action Banner */}
              <div className="bg-gradient-to-r from-cyan-50 to-teal-50 border border-cyan-200 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-2xs">
                <div className="space-y-1">
                  <div className="flex items-center gap-2">
                    <span className="font-bold text-slate-900 text-sm">Sheet Tab `thietbi` (Danh mục Thiết bị & Trạm điện)</span>
                    <span className="text-[10px] bg-cyan-200 text-cyan-950 px-2 py-0.5 rounded font-mono font-bold">
                      {equipment.length} thiết bị
                    </span>
                  </div>
                  <p className="text-xs text-slate-600">
                    Khi thêm thiết bị mới, bạn có thể <strong>chọn trực tiếp Khách hàng và Nhà máy / Phân xưởng từ dropdown</strong> liên kết với danh mục Master.
                  </p>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <button
                    onClick={() => handleCreateSheetTab('thietbi')}
                    disabled={isCreatingTab}
                    className="px-3.5 py-2 bg-cyan-100 hover:bg-cyan-200 text-cyan-950 border border-cyan-300 rounded-lg text-xs font-bold flex items-center gap-1.5 transition-all disabled:opacity-50"
                  >
                    <RefreshCw size={13} className={isCreatingTab ? 'animate-spin' : ''} />
                    <span>Đồng bộ / Tạo Sheet thietbi</span>
                  </button>
                  <button
                    onClick={handleOpenAddEquipment}
                    className="px-3.5 py-2 bg-cyan-700 hover:bg-cyan-800 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-sm transition-all"
                  >
                    <Plus size={14} />
                    <span>Bổ sung Thiết Bị mới</span>
                  </button>
                </div>
              </div>

              {/* Search */}
              <div className="relative">
                <Search size={14} className="absolute left-3 top-2.5 text-slate-400" />
                <input
                  type="text"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="Tìm kiếm mã thiết bị, tên máy, khách hàng, nhà máy..."
                  className="w-full pl-9 pr-3 py-2 text-xs border border-slate-200 rounded-xl focus:ring-2 focus:ring-cyan-500 focus:outline-none bg-slate-50/50"
                />
              </div>

              {/* Table with Edit and Delete */}
              <div className="border border-slate-200 rounded-xl overflow-hidden shadow-2xs">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-50 border-b border-slate-200 font-bold text-slate-700">
                    <tr>
                      <th className="p-3">Mã Tag ID</th>
                      <th className="p-3">Tên Thiết Bị</th>
                      <th className="p-3">Phân loại</th>
                      <th className="p-3">Khách hàng / Nhà máy</th>
                      <th className="p-3">Vị trí / Ngăn lộ</th>
                      <th className="p-3 text-right">HI (%)</th>
                      <th className="p-3 text-center">Thao tác</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {equipment
                      .filter(
                        (e) =>
                          e.id.toLowerCase().includes(search.toLowerCase()) ||
                          e.name.toLowerCase().includes(search.toLowerCase()) ||
                          e.type.toLowerCase().includes(search.toLowerCase()) ||
                          (e.customer && e.customer.toLowerCase().includes(search.toLowerCase())) ||
                          (e.factory && e.factory.toLowerCase().includes(search.toLowerCase()))
                      )
                      .map((eq) => (
                        <tr key={eq.id} className="hover:bg-cyan-50/30 transition-colors">
                          <td className="p-3 font-mono font-bold text-cyan-800">{eq.id}</td>
                          <td className="p-3 font-bold text-slate-900">{eq.name}</td>
                          <td className="p-3">
                            <span className="bg-slate-100 text-slate-800 px-2 py-0.5 rounded font-medium text-[10px]">
                              {eq.type}
                            </span>
                          </td>
                          <td className="p-3 text-slate-700">
                            <div className="font-semibold text-slate-900 truncate max-w-[180px]">{eq.customer}</div>
                            <div className="text-[10px] text-slate-500 truncate max-w-[180px]">{eq.factory}</div>
                          </td>
                          <td className="p-3 text-slate-500 truncate max-w-[150px]">{eq.location || '—'}</td>
                          <td className="p-3 text-right font-bold text-emerald-700">{eq.health || 95}%</td>
                          <td className="p-3 text-center">
                            <div className="flex items-center justify-center gap-1.5">
                              <button
                                onClick={() => handleEditEquipmentClick(eq)}
                                className="p-1.5 text-cyan-700 hover:text-cyan-900 hover:bg-cyan-50 rounded-lg transition-colors"
                                title="Chỉnh sửa thông tin thiết bị"
                              >
                                <Edit3 size={14} />
                              </button>
                              <button
                                onClick={() => handleDeleteEquipmentClick(eq)}
                                className="p-1.5 text-rose-600 hover:text-rose-800 hover:bg-rose-50 rounded-lg transition-colors"
                                title="Xóa thiết bị"
                              >
                                <Trash2 size={14} />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* ================================================================= */}
          {/* TAB 4: SOẠN DỰ ÁN & THIẾT BỊ HIỆN TRƯỜNG (KKS CODE CHO FSE) */}
          {/* ================================================================= */}
          {activeTab === 'duan' && (
            <div className="space-y-4">
              {/* TIMELINE GANTT CHART OVERVIEW FOR PROJECT MANAGERS */}
              <ProjectPortfolioGanttTimeline
                projects={projects}
                onSelectProject={(selectedPrj) => {
                  const el = document.getElementById(`project_card_${selectedPrj.id}`);
                  if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
                }}
                onOpenManageEquipment={handleOpenManageProjectEquipment}
                onOpenAddProject={handleOpenAddProject}
              />

              {/* Sheet Tab Action Card */}
              <div className="bg-gradient-to-r from-indigo-50 via-purple-50 to-blue-50 border border-indigo-200 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-2xs">
                <div className="space-y-1">
                  <div className="flex items-center gap-2">
                    <span className="font-bold text-slate-900 text-sm">Sheet Tab `Dự án` (Soạn Dự Án & Danh Mục Thiết Bị / KKS Code)</span>
                    <span className="text-[10px] bg-indigo-200 text-indigo-950 px-2 py-0.5 rounded font-mono font-bold">
                      {projects.length} dự án đã chuẩn bị
                    </span>
                  </div>
                  <p className="text-xs text-slate-600">
                    Bộ phận Field Service chuẩn bị trước Dự án, gắn Khách hàng, kế hoạch Planning (Bắt đầu - Kết thúc), tiêu chuẩn kiểm định <strong>(ATS / MTS / ECS)</strong> và soạn danh sách thiết bị KKS code. 
                    Khi kỹ sư ra site chỉ cần <strong>nhập hoặc chọn Mã dự án</strong> là hệ thống tự động đổ đầy đủ Khách hàng và dropdown các thiết bị KKS tương ứng để kiểm tra test ngay!
                  </p>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <button
                    onClick={() => handleCreateSheetTab('Dự án')}
                    disabled={isCreatingTab}
                    className="px-3.5 py-2 bg-indigo-700 hover:bg-indigo-800 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-xs transition-all disabled:opacity-50"
                  >
                    <FolderPlus size={14} />
                    <span>{isCreatingTab ? 'Đang tạo tab...' : 'Tạo Sheet tab `Dự án` Google Sheets'}</span>
                  </button>
                  <button
                    onClick={handleOpenAddProject}
                    className="px-3.5 py-2 bg-white hover:bg-slate-50 text-indigo-900 border border-indigo-300 rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-2xs"
                  >
                    <Plus size={14} className="text-indigo-600" />
                    <span>+ Tạo Dự Án Mới</span>
                  </button>
                </div>
              </div>

              {/* Search Projects */}
              <div className="relative">
                <Search size={14} className="absolute left-3 top-2.5 text-slate-400" />
                <input
                  type="text"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="Tìm kiếm mã dự án (DA-2026-001...), tên dự án, khách hàng, mã KKS..."
                  className="w-full pl-9 pr-3 py-2 text-xs border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:outline-none bg-slate-50/50"
                />
              </div>

              {/* Projects List Card View */}
              <div className="space-y-3">
                {projects
                  .filter(
                    (p) =>
                      p.id.toLowerCase().includes(search.toLowerCase()) ||
                      p.name.toLowerCase().includes(search.toLowerCase()) ||
                      p.customerName.toLowerCase().includes(search.toLowerCase()) ||
                      (p.factory && p.factory.toLowerCase().includes(search.toLowerCase())) ||
                      (p.equipmentList && p.equipmentList.some(eq => 
                        eq.kksCode.toLowerCase().includes(search.toLowerCase()) || 
                        eq.name.toLowerCase().includes(search.toLowerCase())
                      ))
                  )
                  .map((prj) => {
                    const std = prj.standard || (prj.name?.toLowerCase().includes('ecs') ? 'ecs' : prj.name?.toLowerCase().includes('mts') ? 'mts' : 'ats');
                    const planningStart = prj.planningStartDate || prj.startDate || '2026-03-01';
                    const planningEnd = prj.planningEndDate || prj.endDate || '2026-03-25';
                    const durationDays = (() => {
                      const s = new Date(planningStart).getTime();
                      const e = new Date(planningEnd).getTime();
                      return Math.max(1, Math.round((e - s) / (1000 * 60 * 60 * 24)) + 1);
                    })();
                    const pmpHealth = calculatePmpProjectHealth(prj);

                    return (
                    <div
                      key={prj.id}
                      id={`project_card_${prj.id}`}
                      className={`bg-white border rounded-xl p-4 shadow-2xs transition-all space-y-3 ${
                        pmpHealth.isOverdue
                          ? 'border-rose-300 ring-1 ring-rose-200/50 hover:border-rose-400'
                          : pmpHealth.status === 'at_risk'
                          ? 'border-amber-300 ring-1 ring-amber-200/50 hover:border-amber-400'
                          : 'border-slate-200 hover:border-indigo-300'
                      }`}
                    >
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-100 pb-3">
                        <div className="flex items-center gap-2.5">
                          <span className="px-2.5 py-1 rounded-lg bg-slate-100 border border-slate-200 text-slate-800 font-mono font-black text-xs">
                            {prj.id}
                          </span>
                          <div>
                            <div className="flex items-center gap-2 flex-wrap">
                              <h4 className="font-bold text-slate-900 text-sm">{prj.name}</h4>
                              {/* Standard Badge */}
                              <span className={`px-2 py-0.2 rounded font-mono font-extrabold text-[10px] ${
                                std === 'ecs'
                                  ? 'bg-teal-100 text-teal-800 border border-teal-300'
                                  : std === 'mts'
                                  ? 'bg-amber-100 text-amber-900 border border-amber-300'
                                  : 'bg-blue-100 text-blue-800 border border-blue-300'
                              }`}>
                                {std.toUpperCase()}
                              </span>
                            </div>
                            <div className="text-[11px] text-slate-500 flex flex-wrap items-center gap-2 mt-0.5">
                              <span className="font-semibold text-blue-700">KH: {prj.customerId} - {prj.customerName}</span>
                              <span>•</span>
                              <span>Nhà máy / Trạm: <strong className="text-slate-700">{prj.factory}</strong></span>
                              {prj.leadFse && (
                                <>
                                  <span>•</span>
                                  <span>FSE phụ trách: <strong className="text-amber-800">{prj.leadFse}</strong></span>
                                </>
                              )}
                            </div>
                          </div>
                        </div>

                        <div className="flex items-center gap-2 self-end sm:self-center flex-wrap">
                          {/* PMP Standard RAG Health Badge */}
                          <span
                            className={`px-2.5 py-1 rounded-full text-[10px] font-bold border flex items-center gap-1.5 ${pmpHealth.colorScheme.badgeBg} ${pmpHealth.colorScheme.badgeText} ${pmpHealth.colorScheme.badgeBorder}`}
                            title={`Chỉ số tiến độ SPI: ${pmpHealth.spi} | Kế hoạch: ${pmpHealth.plannedProgressPercent}% | Thực tế: ${pmpHealth.actualProgressPercent}%`}
                          >
                            <span className={`w-2 h-2 rounded-full ${pmpHealth.colorScheme.dotColor}`} />
                            <span>PMP: {pmpHealth.label}</span>
                            {pmpHealth.isOverdue && (
                              <span className="font-mono text-[9px] bg-rose-600 text-white px-1 py-0.2 rounded font-extrabold uppercase ml-0.5">
                                Trễ {Math.abs(pmpHealth.varianceDays)}d
                              </span>
                            )}
                          </span>

                          <button
                            type="button"
                            onClick={() => {
                              if (onSelectProjectForWorkOrder) {
                                onSelectProjectForWorkOrder(prj);
                              } else {
                                window.dispatchEvent(new CustomEvent('tev_select_project_for_wo', { detail: prj }));
                                onClose();
                              }
                            }}
                            className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-2xs transition-all active:scale-95"
                            title="Chọn dự án này và chuyển đến Trạm FSE để lập phiếu Work Order"
                          >
                            <ClipboardList size={14} />
                            <span>Vào Trạm FSE (Tạo WO)</span>
                          </button>

                          <button
                            onClick={() => handleOpenManageProjectEquipment(prj)}
                            className="px-3 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-800 border border-indigo-200 rounded-lg text-xs font-bold flex items-center gap-1.5 transition-colors"
                          >
                            <ListChecks size={14} />
                            <span>Soạn Thiết Bị ({prj.equipmentList?.length || 0})</span>
                          </button>

                          <button
                            onClick={() => handleEditProjectClick(prj)}
                            className="p-1.5 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded-lg transition-colors"
                            title="Sửa thông tin dự án"
                          >
                            <Edit3 size={14} />
                          </button>

                          <button
                            onClick={() => handleDeleteProjectClick(prj)}
                            className="p-1.5 text-rose-600 hover:text-rose-800 hover:bg-rose-50 rounded-lg transition-colors"
                            title="Xóa dự án"
                          >
                            <Trash2 size={14} />
                          </button>
                        </div>
                      </div>

                      {/* PMP EVM Progress & Schedule Variance Bar */}
                      <div className="space-y-1">
                        <div className="flex items-center justify-between text-[11px] text-slate-500 font-medium">
                          <span className="flex items-center gap-1.5">
                            <span className="font-bold text-slate-700">Tiến độ KKS Hiện Trường:</span>
                            <span className="font-mono font-black text-slate-900">{pmpHealth.actualProgressPercent}%</span>
                            <span>({prj.equipmentList?.filter(e => e.status === 'đã_test').length || 0}/{prj.equipmentList?.length || 0} thiết bị)</span>
                          </span>
                          <span className="font-mono text-[10px] text-slate-500">
                            Tiến độ lý thuyết: <strong className="text-slate-700">{pmpHealth.plannedProgressPercent}%</strong> • SPI: <strong className={pmpHealth.spi >= 0.95 ? 'text-emerald-700' : pmpHealth.spi >= 0.8 ? 'text-amber-700' : 'text-rose-700'}>{pmpHealth.spi}</strong>
                          </span>
                        </div>
                        <div className="w-full bg-slate-100 rounded-full h-2 overflow-hidden flex border border-slate-200">
                          <div
                            className={`h-full bg-gradient-to-r ${pmpHealth.colorScheme.barGradient} transition-all rounded-full`}
                            style={{ width: `${pmpHealth.actualProgressPercent}%` }}
                          />
                        </div>
                      </div>

                      {/* PLANNING TIME & STANDARDS SUMMARY ROW - PMP COLOR FORMATTED */}
                      <div className="grid grid-cols-1 sm:grid-cols-4 gap-2 bg-slate-50/90 border border-slate-200/90 rounded-xl p-2.5 text-xs">
                        <div className="flex items-center gap-2">
                          <div className="p-1.5 rounded-lg bg-blue-100 text-blue-700">
                            <Calendar size={14} />
                          </div>
                          <div>
                            <div className="text-[10px] text-slate-500 font-bold uppercase tracking-wider">Planning Bắt Đầu</div>
                            <div className="font-mono font-bold text-slate-900">{planningStart}</div>
                          </div>
                        </div>

                        <div className="flex items-center gap-2">
                          <div className={`p-1.5 rounded-lg ${
                            pmpHealth.isOverdue
                              ? 'bg-rose-100 text-rose-700 animate-pulse'
                              : pmpHealth.isEndingSoon
                              ? 'bg-amber-100 text-amber-700'
                              : 'bg-blue-100 text-blue-700'
                          }`}>
                            <Clock size={14} />
                          </div>
                          <div>
                            <div className="text-[10px] text-slate-500 font-bold uppercase tracking-wider flex items-center gap-1">
                              <span>Planning Kết Thúc</span>
                              {pmpHealth.isOverdue && (
                                <span className="text-[9px] text-rose-600 font-extrabold uppercase">● Quá hạn</span>
                              )}
                            </div>
                            <div className={`font-mono font-bold ${
                              pmpHealth.isOverdue
                                ? 'text-rose-700'
                                : pmpHealth.isEndingSoon
                                ? 'text-amber-800'
                                : 'text-slate-900'
                            }`}>{planningEnd}</div>
                          </div>
                        </div>

                        <div className="flex items-center gap-2">
                          <div className="p-1.5 rounded-lg bg-slate-200 text-slate-700">
                            <Layers size={14} />
                          </div>
                          <div>
                            <div className="text-[10px] text-slate-500 font-bold uppercase tracking-wider">Thời Lượng & Hạn</div>
                            <div className="font-mono font-bold text-slate-900 flex items-center gap-1.5">
                              <span>{durationDays} ngày</span>
                              <span className={`text-[10px] font-semibold px-1 rounded ${
                                pmpHealth.status === 'completed'
                                  ? 'bg-emerald-100 text-emerald-800'
                                  : pmpHealth.isOverdue
                                  ? 'bg-rose-100 text-rose-800'
                                  : pmpHealth.varianceDays <= 5
                                  ? 'bg-amber-100 text-amber-900'
                                  : 'bg-slate-100 text-slate-600'
                              }`}>
                                {pmpHealth.status === 'completed'
                                  ? 'Đã xong'
                                  : pmpHealth.isOverdue
                                  ? `Trễ ${Math.abs(pmpHealth.varianceDays)}d`
                                  : `Còn ${pmpHealth.varianceDays}d`}
                              </span>
                            </div>
                          </div>
                        </div>

                        <div className="flex items-center gap-2">
                          <div className={`p-1.5 rounded-lg ${
                            std === 'ecs' ? 'bg-teal-100 text-teal-700' : std === 'mts' ? 'bg-amber-100 text-amber-800' : 'bg-blue-100 text-blue-700'
                          }`}>
                            <ShieldCheck size={14} />
                          </div>
                          <div>
                            <div className="text-[10px] text-slate-500 font-bold uppercase tracking-wider">Tiêu Chuẩn NETA</div>
                            <div className="flex items-center gap-1 font-bold">
                              <span className={`px-2 py-0.2 rounded text-[10px] font-mono font-extrabold ${
                                std === 'ecs'
                                  ? 'bg-teal-600 text-white'
                                  : std === 'mts'
                                  ? 'bg-amber-500 text-slate-950'
                                  : 'bg-blue-600 text-white'
                              }`}>
                                {std.toUpperCase()}
                              </span>
                              <span className="text-[10px] text-slate-600">
                                {std === 'ecs' ? 'Commissioning' : std === 'mts' ? 'Bảo dưỡng' : 'Nghiệm thu'}
                              </span>
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* Quick Field Entry Launcher for the Project */}
                      <div className="bg-gradient-to-r from-indigo-50 via-blue-50 to-amber-50 border border-indigo-200/90 rounded-xl p-3 flex flex-col md:flex-row md:items-center justify-between gap-3 shadow-2xs">
                        <div className="flex items-center gap-2.5">
                          <div className="p-2 bg-indigo-600 text-white rounded-lg shadow-2xs shrink-0">
                            <Zap size={16} />
                          </div>
                          <div>
                            <div className="text-xs font-bold text-slate-900 flex items-center gap-2 flex-wrap">
                              <span>Nhập Liệu Hiện Trường Dự Án:</span>
                              <span className="font-mono text-indigo-700 font-extrabold">{prj.id}</span>
                              <span className="text-slate-300">•</span>
                              <span className="text-slate-700 font-medium">{prj.name}</span>
                            </div>
                            <p className="text-[11px] text-slate-500 mt-0.5">
                              Chọn thiết bị KKS & tiêu chuẩn (ATS: nghiệm thu đưa vào vận hành | MTS: bảo dưỡng định kỳ | ECS: commissioning vận hành) để mở form NETA chuẩn
                            </p>
                          </div>
                        </div>

                        <div className="flex items-center gap-2 flex-wrap">
                          {prj.equipmentList && prj.equipmentList.length > 0 ? (
                            <div className="flex items-center gap-1.5 bg-white p-1 rounded-xl border border-indigo-200 shadow-2xs">
                              <select
                                id={`quick_eq_${prj.id}`}
                                className="text-xs bg-slate-50 border border-slate-300 rounded-lg px-2 py-1 font-semibold text-slate-700 max-w-[170px] sm:max-w-[210px] truncate outline-none"
                                defaultValue={prj.equipmentList[0].kksCode}
                              >
                                {prj.equipmentList.map((eq, eidx) => (
                                  <option key={eidx} value={eq.kksCode}>
                                    {eq.kksCode} - {eq.name} ({eq.type})
                                  </option>
                                ))}
                              </select>

                              <select
                                id={`quick_std_${prj.id}`}
                                className="text-xs bg-slate-50 border border-slate-300 rounded-lg px-2 py-1 font-bold text-indigo-900 outline-none"
                                defaultValue={std}
                              >
                                <option value="ats">ANSI/NETA ATS (Nghiệm thu)</option>
                                <option value="mts">ANSI/NETA MTS (Bảo dưỡng)</option>
                                <option value="ecs">ANSI/NETA ECS (Commissioning)</option>
                              </select>

                              <button
                                type="button"
                                onClick={() => {
                                  const eqSelect = document.getElementById(`quick_eq_${prj.id}`) as HTMLSelectElement;
                                  const stdSelect = document.getElementById(`quick_std_${prj.id}`) as HTMLSelectElement;
                                  const chosenKks = eqSelect?.value || prj.equipmentList![0].kksCode;
                                  const chosenStd = (stdSelect?.value || std || 'ats') as 'ats' | 'mts' | 'ecs';
                                  const targetEq = prj.equipmentList!.find(e => e.kksCode === chosenKks) || prj.equipmentList![0];
                                  handleNavigateToNeta(prj, targetEq, chosenStd);
                                }}
                                className="px-3 py-1.5 bg-indigo-600 hover:bg-indigo-700 active:scale-95 text-white font-bold text-xs rounded-lg transition-all flex items-center gap-1.5 shadow-xs whitespace-nowrap"
                                title="Đi đến trạm nhập liệu và mở Form NETA chuẩn ngay lập tức"
                              >
                                <Play size={12} className="fill-white" />
                                <span>Đi Đến Nhập Liệu (Form NETA)</span>
                              </button>
                            </div>
                          ) : (
                            <button
                              type="button"
                              onClick={() => handleOpenManageProjectEquipment(prj)}
                              className="px-3 py-1.5 bg-amber-500 hover:bg-amber-600 active:scale-95 text-slate-950 font-bold text-xs rounded-lg transition-all flex items-center gap-1.5 shadow-xs"
                            >
                              <Plus size={13} />
                              <span>+ Soạn Thiết Bị KKS Cho Dự Án</span>
                            </button>
                          )}
                        </div>
                      </div>

                      {/* Professional Categorized Equipment Table */}
                      <div className="space-y-2">
                        <div className="text-xs font-bold text-slate-700 flex items-center justify-between">
                          <span className="flex items-center gap-1.5">
                            <Layers size={14} className="text-indigo-600" />
                            <span>Bảng Thiết Bị KKS & Phân Loại Kiểm Định Chuẩn PMP ({prj.equipmentList?.length || 0} thiết bị)</span>
                          </span>
                          <button
                            type="button"
                            onClick={() => handleOpenManageProjectEquipment(prj)}
                            className="text-xs font-bold text-indigo-700 hover:text-indigo-800 hover:underline flex items-center gap-1"
                          >
                            <Plus size={13} /> Soạn / Thêm thiết bị KKS
                          </button>
                        </div>

                        <ProjectCategorizedEquipmentTable
                          project={prj}
                          equipmentList={prj.equipmentList || []}
                          equipmentStandardMap={equipmentStandardMap}
                          onSetStandard={(key, std) => setEquipmentStandardMap(prev => ({ ...prev, [key]: std }))}
                          onNavigateToNeta={handleNavigateToNeta}
                          onOpenAddEquipmentModal={() => handleOpenManageProjectEquipment(prj)}
                          readOnlyActions={true}
                        />
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* ================================================================= */}
          {/* TAB 5: TẠO THÊM SHEET TAB TÙY CHỌN */}
          {/* ================================================================= */}
          {activeTab === 'create_tab' && (
            <div className="space-y-5">
              <div className="bg-gradient-to-r from-emerald-50 via-teal-50 to-blue-50 border border-emerald-200 rounded-xl p-5 shadow-2xs">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div className="space-y-1">
                    <h3 className="font-bold text-slate-900 text-base flex items-center gap-2">
                      <FolderPlus className="text-emerald-600" size={20} />
                      Tạo Thêm Sheet Tab Trực Tiếp Vào Google Sheets
                    </h3>
                    <p className="text-xs text-slate-600">
                      Chúng ta đã có sẵn: <strong>khachhang</strong>, <strong>thietbi</strong>, <strong>FSE</strong>, và <strong>Dự án</strong>. Bạn có thể tạo thêm bất kỳ sheet tab mới nào phục vụ quản lý hiện trường.
                    </p>
                  </div>
                  {googleSheetUrl && (
                    <a
                      href={googleSheetUrl}
                      target="_blank"
                      rel="noreferrer"
                      className="px-3.5 py-2 bg-white hover:bg-slate-50 border border-slate-300 text-slate-800 rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-2xs shrink-0 self-start md:self-auto"
                    >
                      <ExternalLink size={14} className="text-emerald-600" />
                      <span>Xem File Google Sheets</span>
                    </a>
                  )}
                </div>
              </div>

              {/* Form Tạo Sheet Tab Mới */}
              <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-2xs space-y-4">
                <h4 className="font-bold text-slate-900 text-sm flex items-center gap-2">
                  <Table size={16} className="text-blue-600" />
                  Cấu Hình & Khởi Tạo Sheet Tab Mới
                </h4>

                {/* Preset Selector */}
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1.5">
                    1. Chọn mẫu tab nhanh hoặc tạo tab tùy chỉnh:
                  </label>
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5">
                    <button
                      type="button"
                      onClick={() => {
                        setSelectedPreset('duan');
                        setCustomTabName('Dự án');
                        setCustomTabHeaderInput(SHEET_PROJECT_HEADERS.join(', '));
                      }}
                      className={`p-3 rounded-xl border text-left transition-all ${
                        selectedPreset === 'duan'
                          ? 'bg-indigo-50 border-indigo-300 ring-2 ring-indigo-400 text-indigo-950'
                          : 'bg-slate-50 border-slate-200 hover:bg-slate-100 text-slate-700'
                      }`}
                    >
                      <div className="font-bold text-xs flex items-center gap-1.5">
                        <Briefcase size={14} className="text-indigo-600" /> Tab `Dự án`
                      </div>
                      <div className="text-[10px] text-slate-500 mt-1">Soạn KKS thiết bị hiện trường</div>
                    </button>

                    <button
                      type="button"
                      onClick={() => {
                        setSelectedPreset('fse');
                        setCustomTabName('FSE');
                        setCustomTabHeaderInput(SHEET_FSE_HEADERS.join(', '));
                      }}
                      className={`p-3 rounded-xl border text-left transition-all ${
                        selectedPreset === 'fse'
                          ? 'bg-amber-50 border-amber-300 ring-2 ring-amber-400 text-amber-950'
                          : 'bg-slate-50 border-slate-200 hover:bg-slate-100 text-slate-700'
                      }`}
                    >
                      <div className="font-bold text-xs flex items-center gap-1.5">
                        <UserCheck size={14} className="text-amber-600" /> Tab `FSE`
                      </div>
                      <div className="text-[10px] text-slate-500 mt-1">Danh bạ kỹ sư hiện trường</div>
                    </button>

                    <button
                      type="button"
                      onClick={() => {
                        setSelectedPreset('khachhang');
                        setCustomTabName('khachhang');
                        setCustomTabHeaderInput('Mã khách hàng, Tên khách hàng, Nhà máy / Trạm, Email, Số điện thoại, Địa chỉ, Ngày tạo, Ngày cập nhật');
                      }}
                      className={`p-3 rounded-xl border text-left transition-all ${
                        selectedPreset === 'khachhang'
                          ? 'bg-blue-50 border-blue-300 ring-2 ring-blue-400 text-blue-950'
                          : 'bg-slate-50 border-slate-200 hover:bg-slate-100 text-slate-700'
                      }`}
                    >
                      <div className="font-bold text-xs flex items-center gap-1.5">
                        <Building2 size={14} className="text-blue-600" /> Tab `khachhang`
                      </div>
                      <div className="text-[10px] text-slate-500 mt-1">Danh mục khách hàng KHXXX</div>
                    </button>

                    <button
                      type="button"
                      onClick={() => {
                        setSelectedPreset('custom');
                        setCustomTabName('');
                        setCustomTabHeaderInput('Mã bản ghi, Tên hạng mục, Phân loại, Người thực hiện, Khách hàng, Ghi chú, Ngày tạo');
                      }}
                      className={`p-3 rounded-xl border text-left transition-all ${
                        selectedPreset === 'custom'
                          ? 'bg-emerald-50 border-emerald-300 ring-2 ring-emerald-400 text-emerald-950'
                          : 'bg-slate-50 border-slate-200 hover:bg-slate-100 text-slate-700'
                      }`}
                    >
                      <div className="font-bold text-xs flex items-center gap-1.5">
                        <Plus size={14} className="text-emerald-600" /> Tự đặt tên Tab
                      </div>
                      <div className="text-[10px] text-slate-500 mt-1">Sheet tab tùy chỉnh tự do</div>
                    </button>
                  </div>
                </div>

                {/* Input Fields */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs">
                  <div>
                    <label className="block font-bold text-slate-700 mb-1">
                      2. Tên Sheet Tab Muốn Tạo *
                    </label>
                    <input
                      type="text"
                      required
                      value={customTabName}
                      onChange={(e) => setCustomTabName(e.target.value)}
                      placeholder="VD: Dự án, BienBanKiemDinh, NhatKyHienTruong..."
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-emerald-500 font-bold text-slate-900"
                    />
                  </div>

                  <div>
                    <label className="block font-bold text-slate-700 mb-1">
                      3. Danh Sách Cột Tiêu Đề (Phân cách bằng dấu phẩy)
                    </label>
                    <input
                      type="text"
                      value={customTabHeaderInput}
                      onChange={(e) => setCustomTabHeaderInput(e.target.value)}
                      placeholder="Cột 1, Cột 2, Cột 3, Cột 4..."
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl focus:bg-white focus:ring-2 focus:ring-emerald-500 text-slate-800"
                    />
                  </div>
                </div>

                {/* Submit Action */}
                <div className="pt-2 flex items-center justify-between border-t border-slate-100">
                  <div className="text-xs text-slate-500">
                    Trạng thái kết nối Google: {isGoogleConnected ? (
                      <span className="text-emerald-600 font-bold">● Đã kết nối sẵn sàng</span>
                    ) : (
                      <span className="text-amber-600 font-bold">○ Chưa kết nối (Cần đăng nhập Google)</span>
                    )}
                  </div>

                  <button
                    type="button"
                    onClick={() => {
                      const headersList = customTabHeaderInput
                        .split(',')
                        .map((h) => h.trim())
                        .filter((h) => h.length > 0);
                      handleCreateSheetTab(customTabName, headersList);
                    }}
                    disabled={isCreatingTab || !customTabName.trim()}
                    className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold flex items-center gap-2 shadow-sm transition-all disabled:opacity-50"
                  >
                    <FolderPlus size={16} />
                    <span>{isCreatingTab ? 'Đang tạo tab lên Google Sheets...' : '+ Tạo Sheet Tab Ngay Lên Google Sheets'}</span>
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="p-3 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-500 shrink-0">
          <div className="flex items-center gap-2 flex-wrap">
            <span className="font-semibold text-slate-700">4 Sheet Cốt Lõi:</span>
            <span className="bg-amber-100 text-amber-950 px-2 py-0.5 rounded font-mono font-medium">Sheet FSE</span>
            <span className="bg-blue-100 text-blue-900 px-2 py-0.5 rounded font-mono font-medium">Sheet khachhang (KHXXX)</span>
            <span className="bg-cyan-100 text-cyan-900 px-2 py-0.5 rounded font-mono font-medium">Sheet thietbi</span>
            <span className="bg-indigo-100 text-indigo-900 px-2 py-0.5 rounded font-mono font-medium">Sheet Dự án (KKS Code)</span>
          </div>
          <button
            onClick={onClose}
            className="px-4 py-1.5 bg-slate-800 hover:bg-slate-900 text-white rounded-lg font-bold transition-colors"
          >
            Đóng
          </button>
        </div>
      </div>

      {/* ===================================================================== */}
      {/* MODAL 1: ADD / EDIT FSE ENGINEER */}
      {/* ===================================================================== */}
      {showAddFse && (
        <div className="fixed inset-0 z-60 bg-slate-900/60 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl max-w-sm w-full p-4 border border-slate-200 animate-in fade-in-50 zoom-in-95 duration-100">
            <div className="flex items-center justify-between pb-2 border-b border-slate-100 mb-3">
              <h3 className="font-bold text-sm text-slate-900 flex items-center gap-1.5">
                <UserCheck size={16} className="text-amber-600" />
                {editingFse ? 'Chỉnh Sửa Thông Tin Kỹ Sư FSE' : 'Bổ Sung Kỹ Sư FSE'}
              </h3>
              <button onClick={() => setShowAddFse(false)} className="text-slate-400 hover:text-slate-600">
                <X size={15} />
              </button>
            </div>
            <form onSubmit={handleSaveFse} className="space-y-3 text-xs">
              <div>
                <label className="block font-semibold text-slate-700 mb-1">Họ và Tên Kỹ Sư *</label>
                <input
                  type="text"
                  required
                  placeholder="VD: Nguyễn Văn Nam"
                  value={newFse.name}
                  onChange={(e) => setNewFse({ ...newFse, name: e.target.value })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-amber-500 focus:outline-none"
                />
              </div>
              <div>
                <label className="block font-semibold text-slate-700 mb-1">Email Kỹ Sư *</label>
                <input
                  type="email"
                  required
                  disabled={editingFse?.email === 'sgm1707@gmail.com'}
                  placeholder="nam.nv@tev.vn"
                  value={newFse.email}
                  onChange={(e) => setNewFse({ ...newFse, email: e.target.value })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-amber-500 focus:outline-none disabled:bg-slate-100 disabled:text-slate-500"
                />
              </div>
              <div>
                <label className="block font-semibold text-slate-700 mb-1">Chức Danh / Vai Trò</label>
                <input
                  type="text"
                  placeholder="Kỹ sư Thí nghiệm Hiện trường NETA"
                  value={newFse.role}
                  onChange={(e) => setNewFse({ ...newFse, role: e.target.value })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-amber-500 focus:outline-none"
                />
              </div>
              <div>
                <label className="block font-semibold text-slate-700 mb-1">Số Điện Thoại</label>
                <input
                  type="text"
                  placeholder="0912 345 678"
                  value={newFse.phone}
                  onChange={(e) => setNewFse({ ...newFse, phone: e.target.value })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-amber-500 focus:outline-none"
                />
              </div>
              <div>
                <label className="block font-semibold text-slate-700 mb-1">Công Ty / Đơn Vị</label>
                <input
                  type="text"
                  placeholder="TEV Field Service"
                  value={newFse.company}
                  onChange={(e) => setNewFse({ ...newFse, company: e.target.value })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-amber-500 focus:outline-none"
                />
              </div>
              <div className="pt-2 flex justify-end gap-2 border-t border-slate-100">
                <button type="button" onClick={() => setShowAddFse(false)} className="px-3 py-1.5 border rounded-lg text-slate-600 font-semibold">Hủy</button>
                <button type="submit" className="px-3.5 py-1.5 bg-amber-600 hover:bg-amber-700 text-white font-bold rounded-lg shadow-xs">
                  {editingFse ? 'Cập Nhật FSE' : 'Lưu FSE'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ===================================================================== */}
      {/* MODAL 2: ADD / EDIT CUSTOMER (KHXXX) */}
      {/* ===================================================================== */}
      {showAddCustomer && (
        <div className="fixed inset-0 z-60 bg-slate-900/60 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl max-w-sm w-full p-4 border border-slate-200 animate-in fade-in-50 zoom-in-95 duration-100">
            <div className="flex items-center justify-between pb-2 border-b border-slate-100 mb-3">
              <h3 className="font-bold text-sm text-slate-900 flex items-center gap-1.5">
                <Building2 size={16} className="text-blue-600" />
                {editingCustomer ? 'Chỉnh Sửa Khách Hàng' : 'Bổ Sung Khách Hàng Mới'}
              </h3>
              <button onClick={() => setShowAddCustomer(false)} className="text-slate-400 hover:text-slate-600">
                <X size={15} />
              </button>
            </div>
            <form onSubmit={handleSaveCustomer} className="space-y-3 text-xs">
              <div>
                <label className="block font-semibold text-slate-700 mb-1">Mã Khách Hàng (Chuẩn KHXXX) *</label>
                <input
                  type="text"
                  required
                  placeholder="KH001, KH002..."
                  value={newCust.id}
                  onChange={(e) => setNewCust({ ...newCust, id: e.target.value.toUpperCase() })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none font-mono font-bold text-blue-700 uppercase"
                />
                <p className="text-[10px] text-slate-500 mt-0.5">Tự động tăng theo quy ước KHXXX (KH001, KH002...).</p>
              </div>
              <div>
                <label className="block font-semibold text-slate-700 mb-1">Tên Doanh Nghiệp / Khách Hàng *</label>
                <input
                  type="text"
                  required
                  placeholder="VD: Công ty Cổ phần Năng lượng TEV"
                  value={newCust.name}
                  onChange={(e) => setNewCust({ ...newCust, name: e.target.value })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none font-semibold text-slate-900"
                />
              </div>
              <div>
                <label className="block font-semibold text-slate-700 mb-1">Nhà Máy / Trạm / Phân Xưởng (Phân cách bằng dấu phẩy)</label>
                <input
                  type="text"
                  placeholder="Nhà máy Bắc Ninh, Trạm 110kV..."
                  value={newCust.factories}
                  onChange={(e) => setNewCust({ ...newCust, factories: e.target.value })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                />
              </div>
              <div>
                <label className="block font-semibold text-slate-700 mb-1">Email Liên Hệ</label>
                <input
                  type="email"
                  placeholder="contact@doanhnghiep.vn"
                  value={newCust.email}
                  onChange={(e) => setNewCust({ ...newCust, email: e.target.value })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                />
              </div>
              <div>
                <label className="block font-semibold text-slate-700 mb-1">Số Điện Thoại</label>
                <input
                  type="text"
                  placeholder="024.3988.7766"
                  value={newCust.phone}
                  onChange={(e) => setNewCust({ ...newCust, phone: e.target.value })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                />
              </div>
              <div>
                <label className="block font-semibold text-slate-700 mb-1">Địa Chỉ</label>
                <input
                  type="text"
                  placeholder="KCN Quế Võ, Bắc Ninh"
                  value={newCust.address}
                  onChange={(e) => setNewCust({ ...newCust, address: e.target.value })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                />
              </div>
              <div className="pt-2 flex justify-end gap-2 border-t border-slate-100">
                <button type="button" onClick={() => setShowAddCustomer(false)} className="px-3 py-1.5 border rounded-lg text-slate-600 font-semibold">Hủy</button>
                <button type="submit" className="px-3.5 py-1.5 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-lg shadow-xs">
                  {editingCustomer ? 'Cập Nhật KH' : 'Lưu Khách Hàng'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ===================================================================== */}
      {/* MODAL 3: ADD / EDIT EQUIPMENT (VỚI DROPDOWN KHÁCH HÀNG & NHÀ MÁY) */}
      {/* ===================================================================== */}
      {showAddEquipment && (
        <div className="fixed inset-0 z-60 bg-slate-900/60 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl max-w-md w-full p-4 border border-slate-200 animate-in fade-in-50 zoom-in-95 duration-100">
            <div className="flex items-center justify-between pb-2 border-b border-slate-100 mb-3">
              <h3 className="font-bold text-sm text-slate-900 flex items-center gap-1.5">
                <Cpu size={16} className="text-cyan-600" />
                {editingEquipment ? 'Chỉnh Sửa Thiết Bị Master' : 'Bổ Sung Thiết Bị Mới'}
              </h3>
              <button onClick={() => setShowAddEquipment(false)} className="text-slate-400 hover:text-slate-600">
                <X size={15} />
              </button>
            </div>
            <form onSubmit={handleSaveEquipment} className="space-y-3 text-xs">
              
              {/* TRƯỜNG DROPDOWN KHÁCH HÀNG (YÊU CẦU CỦA USER) */}
              <div>
                <label className="block font-bold text-slate-800 mb-1 flex items-center gap-1 text-cyan-900">
                  <Building2 size={13} className="text-blue-600" />
                  Khách Hàng Sở Hữu Thiết Bị *
                </label>
                <select
                  value={newEq.customer}
                  onChange={(e) => {
                    const selectedVal = e.target.value;
                    const found = customers.find(c => `${c.id} - ${c.name}` === selectedVal || c.name === selectedVal);
                    const defaultFac = (found && found.factories && found.factories[0]) || 'Nhà máy chính';
                    setNewEq({
                      ...newEq,
                      customer: selectedVal,
                      factory: defaultFac
                    });
                  }}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:outline-none bg-slate-50 font-semibold text-slate-900"
                >
                  {customers.map((c) => (
                    <option key={c.id} value={`${c.id} - ${c.name}`}>
                      {c.id} - {c.name}
                    </option>
                  ))}
                </select>
              </div>

              {/* TRƯỜNG DROPDOWN NHÀ MÁY / PHÂN XƯỞNG (TỰ LOAD THEO KHÁCH HÀNG ĐÃ CHỌN) */}
              <div>
                <label className="block font-bold text-slate-800 mb-1 flex items-center gap-1 text-cyan-900">
                  <MapPin size={13} className="text-cyan-600" />
                  Nhà Máy / Phân Xưởng / Trạm *
                </label>
                <select
                  value={newEq.factory}
                  onChange={(e) => setNewEq({ ...newEq, factory: e.target.value })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:outline-none bg-slate-50 font-semibold text-slate-900"
                >
                  {selectedCustObj && selectedCustObj.factories && selectedCustObj.factories.length > 0 ? (
                    selectedCustObj.factories.map((fac, idx) => (
                      <option key={idx} value={fac}>{fac}</option>
                    ))
                  ) : (
                    <>
                      <option value="Nhà máy chính">Nhà máy chính</option>
                      <option value="Trạm biến áp chính TBA-01">Trạm biến áp chính TBA-01</option>
                    </>
                  )}
                  <option value="__custom__">+ Nhập tên nhà máy mới khác...</option>
                </select>

                {newEq.factory === '__custom__' && (
                  <input
                    type="text"
                    required
                    placeholder="Gõ tên nhà máy / trạm biến áp mới..."
                    value={newEq.customFactory}
                    onChange={(e) => setNewEq({ ...newEq, customFactory: e.target.value })}
                    className="w-full mt-1.5 px-3 py-1.5 border border-cyan-400 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:outline-none bg-cyan-50/50"
                  />
                )}
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Mã Tag ID / KKS *</label>
                  <input
                    type="text"
                    required
                    placeholder="TR-05, MVCB-02"
                    value={newEq.id}
                    onChange={(e) => setNewEq({ ...newEq, id: e.target.value.toUpperCase() })}
                    className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:outline-none font-mono font-bold text-cyan-800 uppercase"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Phân Loại NETA *</label>
                  <select
                    value={newEq.type}
                    onChange={(e) => setNewEq({ ...newEq, type: e.target.value })}
                    className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:outline-none"
                  >
                    <option value="Máy biến áp">Máy biến áp</option>
                    <option value="Máy cắt">Máy cắt</option>
                    <option value="Động cơ">Động cơ</option>
                    <option value="Máy phát">Máy phát</option>
                    <option value="Tủ điện">Tủ điện</option>
                    <option value="Cáp điện">Cáp điện</option>
                    <option value="Pin & UPS">Pin & UPS</option>
                    <option value="Rơ le bảo vệ">Rơ le bảo vệ</option>
                    <option value="Tiếp địa">Tiếp địa</option>
                  </select>
                </div>
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">Tên Thiết Bị Đầy Đủ *</label>
                <input
                  type="text"
                  required
                  placeholder="VD: Máy biến áp T3 2500kVA 22/0.4kV..."
                  value={newEq.name}
                  onChange={(e) => setNewEq({ ...newEq, name: e.target.value })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:outline-none font-semibold text-slate-900"
                />
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">Vị Trí / Ngăn Lộ (Bay / Cubicle)</label>
                <input
                  type="text"
                  placeholder="VD: Ngăn lộ 131 TBA 110kV, Tủ phân phối MV-02"
                  value={newEq.location}
                  onChange={(e) => setNewEq({ ...newEq, location: e.target.value })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">Thông Số Kỹ Thuật (Specs tóm tắt)</label>
                <input
                  type="text"
                  placeholder="2500kVA 22/0.4kV Thibidi, Dầu khoáng ONAN"
                  value={newEq.specs}
                  onChange={(e) => setNewEq({ ...newEq, specs: e.target.value })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-cyan-500 focus:outline-none"
                />
              </div>

              <div className="pt-2 flex justify-end gap-2 border-t border-slate-100">
                <button type="button" onClick={() => setShowAddEquipment(false)} className="px-3 py-1.5 border rounded-lg text-slate-600 font-semibold">Hủy</button>
                <button type="submit" className="px-3.5 py-1.5 bg-cyan-700 hover:bg-cyan-800 text-white font-bold rounded-lg shadow-xs">
                  {editingEquipment ? 'Cập Nhật Thiết Bị' : 'Lưu Thiết Bị'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ===================================================================== */}
      {/* MODAL 4: TẠO / SỬA DỰ ÁN MỚI */}
      {/* ===================================================================== */}
      {showAddProject && (
        <div className="fixed inset-0 z-60 bg-slate-900/60 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl max-w-lg w-full p-4 border border-slate-200 animate-in fade-in-50 zoom-in-95 duration-100">
            <div className="flex items-center justify-between pb-2 border-b border-slate-100 mb-3">
              <h3 className="font-bold text-sm text-slate-900 flex items-center gap-1.5">
                <Briefcase size={16} className="text-indigo-600" />
                {editingProject ? 'Chỉnh Sửa Dự Án' : 'Tạo Dự Án Mới Chuẩn Bị Hiện Trường'}
              </h3>
              <button onClick={() => setShowAddProject(false)} className="text-slate-400 hover:text-slate-600">
                <X size={15} />
              </button>
            </div>
            <form onSubmit={handleSaveProject} className="space-y-3 text-xs">
              <div className="grid grid-cols-3 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Mã Dự Án *</label>
                  <input
                    type="text"
                    required
                    value={newPrj.id}
                    onChange={(e) => setNewPrj({ ...newPrj, id: e.target.value.toUpperCase() })}
                    className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 font-mono font-bold text-indigo-700"
                  />
                </div>
                <div className="col-span-2">
                  <label className="block font-semibold text-slate-700 mb-1">Tên Dự Án Hiện Trường *</label>
                  <input
                    type="text"
                    required
                    placeholder="VD: Thí nghiệm định kỳ TBA 110kV Super Energy"
                    value={newPrj.name}
                    onChange={(e) => setNewPrj({ ...newPrj, name: e.target.value })}
                    className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 font-semibold text-slate-900"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Khách Hàng (Master) *</label>
                  <select
                    value={newPrj.customerId}
                    onChange={(e) => {
                      const cid = e.target.value;
                      const cust = customers.find(c => c.id === cid);
                      setNewPrj({
                        ...newPrj,
                        customerId: cid,
                        customerName: cust?.name || '',
                        factory: (cust?.factories && cust.factories[0]) || 'Nhà máy chính'
                      });
                    }}
                    className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 bg-slate-50 font-semibold"
                  >
                    {customers.map((c) => (
                      <option key={c.id} value={c.id}>
                        {c.id} - {c.name}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Nhà Máy / Trạm *</label>
                  {(() => {
                    const currentCust = customers.find(c => c.id === newPrj.customerId);
                    return (
                      <select
                        value={newPrj.factory}
                        onChange={(e) => setNewPrj({ ...newPrj, factory: e.target.value })}
                        className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 bg-slate-50 font-semibold"
                      >
                        {currentCust?.factories && currentCust.factories.length > 0 ? (
                          currentCust.factories.map((fac, i) => (
                            <option key={i} value={fac}>{fac}</option>
                          ))
                        ) : (
                          <option value="Nhà máy chính">Nhà máy chính</option>
                        )}
                      </select>
                    );
                  })()}
                </div>
              </div>

              <div className="grid grid-cols-3 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Tiêu Chuẩn Chính (NETA)</label>
                  <select
                    value={newPrj.standard}
                    onChange={(e) => setNewPrj({ ...newPrj, standard: e.target.value as any })}
                    className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 font-bold text-indigo-900 bg-indigo-50/50"
                  >
                    <option value="ats">ANSI/NETA ATS (Nghiệm thu đóng điện)</option>
                    <option value="mts">ANSI/NETA MTS (Bảo dưỡng định kỳ)</option>
                    <option value="ecs">ANSI/NETA ECS (Commissioning vận hành)</option>
                  </select>
                </div>

                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Kỹ Sư Trưởng Phụ Trách</label>
                  <select
                    value={newPrj.leadFse}
                    onChange={(e) => setNewPrj({ ...newPrj, leadFse: e.target.value })}
                    className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 bg-slate-50"
                  >
                    {engineers.map((eng) => (
                      <option key={eng.email} value={eng.name}>
                        {eng.name} ({eng.role || 'FSE'})
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Trạng Thái (PMP Phase)</label>
                  <select
                    value={newPrj.status}
                    onChange={(e) => setNewPrj({ ...newPrj, status: e.target.value as any })}
                    className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 font-semibold"
                  >
                    <option value="chuẩn_bị">Chuẩn bị trước (PMP: Planning / Baseline)</option>
                    <option value="đang_thực_hiện">Đang thực hiện ngoài site (PMP: Executing)</option>
                    <option value="hoàn_thành">Đã hoàn thành (PMP: Closing / Completed)</option>
                  </select>
                </div>
              </div>

              {/* Planning Dates Section */}
              <div className="p-2.5 bg-gradient-to-r from-indigo-50/60 to-purple-50/60 rounded-xl border border-indigo-200/80 space-y-2">
                <div className="font-bold text-slate-800 flex items-center justify-between">
                  <span className="flex items-center gap-1.5">
                    <Calendar size={13} className="text-indigo-700" />
                    <span>Planning Thời Gian Bắt Đầu & Kết Thúc (Gantt Portfolio)</span>
                  </span>
                  <span className="text-[10px] text-indigo-700 font-mono font-semibold">
                    Kế hoạch triển khai PM/FSE
                  </span>
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block font-semibold text-slate-700 mb-1 text-[11px]">
                      Planning Bắt Đầu *
                    </label>
                    <input
                      type="date"
                      required
                      value={newPrj.planningStartDate}
                      onChange={(e) => setNewPrj({ ...newPrj, planningStartDate: e.target.value, startDate: e.target.value })}
                      className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 font-mono font-bold bg-white"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold text-slate-700 mb-1 text-[11px]">
                      Planning Kết Thúc *
                    </label>
                    <input
                      type="date"
                      required
                      value={newPrj.planningEndDate}
                      onChange={(e) => setNewPrj({ ...newPrj, planningEndDate: e.target.value, endDate: e.target.value })}
                      className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 font-mono font-bold bg-white"
                    />
                  </div>
                </div>
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">Ghi Chú Yêu Cầu Kỹ Thuật Dự Án</label>
                <textarea
                  rows={2}
                  placeholder="VD: Kiểm định theo NETA ATS-2025, đo cách điện cáp 24kV và bảo dưỡng tiếp địa trạm..."
                  value={newPrj.notes}
                  onChange={(e) => setNewPrj({ ...newPrj, notes: e.target.value })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500"
                />
              </div>

              <div className="pt-2 flex justify-end gap-2 border-t border-slate-100">
                <button type="button" onClick={() => setShowAddProject(false)} className="px-3 py-1.5 border rounded-lg text-slate-600 font-semibold">Hủy</button>
                <button type="submit" className="px-3.5 py-1.5 bg-indigo-700 hover:bg-indigo-800 text-white font-bold rounded-lg shadow-xs">
                  {editingProject ? 'Cập Nhật Dự Án' : 'Lưu Dự Án'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ===================================================================== */}
      {/* MODAL 5: SOẠN DANH SÁCH THIẾT BỊ / MÃ KKS CODE CHO DỰ ÁN */}
      {/* ===================================================================== */}
      {managingProjectEquipment && (
        <div className="fixed inset-0 z-60 bg-slate-900/60 flex items-center justify-center p-3 sm:p-5">
          <div className="bg-white rounded-2xl shadow-2xl max-w-4xl w-full max-h-[90vh] flex flex-col border border-slate-200 animate-in fade-in-50 zoom-in-95 duration-150">
            {/* Header */}
            <div className="p-4 bg-indigo-950 text-white flex items-center justify-between shrink-0 rounded-t-2xl">
              <div>
                <h3 className="font-bold text-sm sm:text-base flex items-center gap-2">
                  <ListChecks className="text-amber-300" size={18} />
                  Soạn Danh Sách Thiết Bị & KKS Code: {managingProjectEquipment.name}
                </h3>
                <p className="text-xs text-indigo-200 mt-0.5">
                  Mã dự án: <span className="font-mono font-bold text-amber-300">{managingProjectEquipment.id}</span> • Khách hàng: <span className="font-semibold text-white">{managingProjectEquipment.customerName}</span> ({managingProjectEquipment.factory})
                </p>
              </div>
              <button
                onClick={() => setManagingProjectEquipment(null)}
                className="p-1 text-slate-400 hover:text-white rounded-lg"
              >
                <X size={18} />
              </button>
            </div>

            <div className="p-4 overflow-y-auto flex-1 space-y-4">
              {/* Add form */}
              <div className="bg-indigo-50/60 border border-indigo-200 rounded-xl p-3.5 space-y-3">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-slate-900 text-xs flex items-center gap-1.5">
                    <Plus size={14} className="text-indigo-600" />
                    Thêm Thiết Bị Mới Vào Dự Án
                  </span>
                  <button
                    type="button"
                    onClick={() => setShowMasterEqPicker(!showMasterEqPicker)}
                    className="px-2.5 py-1 bg-white hover:bg-slate-50 border border-indigo-300 text-indigo-900 text-xs font-bold rounded-lg flex items-center gap-1 shadow-2xs"
                  >
                    <Cpu size={13} className="text-cyan-600" />
                    <span>{showMasterEqPicker ? 'Đóng bộ chọn Master' : 'Chọn từ Danh Mục Thiết Bị Master'}</span>
                  </button>
                </div>

                {/* Master equipment quick picker accordion */}
                {showMasterEqPicker && (
                  <div className="bg-white border border-indigo-200 rounded-xl p-3 max-h-48 overflow-y-auto space-y-2">
                    <div className="text-[11px] font-bold text-slate-700">Click thiết bị để nạp nhanh vào dự án:</div>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-1.5">
                      {equipment.map(meq => (
                        <div
                          key={meq.id}
                          onClick={() => handlePickFromMasterEquipment(meq)}
                          className="p-2 border border-slate-200 hover:border-indigo-400 hover:bg-indigo-50/50 rounded-lg cursor-pointer flex items-center justify-between text-xs transition-colors"
                        >
                          <div>
                            <span className="font-mono font-bold text-indigo-700">{meq.id}</span>
                            <span className="text-slate-800 ml-1.5 font-medium">{meq.name}</span>
                          </div>
                          <span className="text-[10px] bg-slate-100 text-slate-600 px-1 rounded">{meq.type}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {/* Manual KKS equipment input */}
                <form onSubmit={handleAddEquipmentToProject} className="grid grid-cols-1 sm:grid-cols-12 gap-2 text-xs">
                  <div className="sm:col-span-3">
                    <label className="block font-semibold text-slate-700 mb-0.5">Mã KKS / Tag ID *</label>
                    <input
                      type="text"
                      required
                      placeholder="=10EAA01GS001"
                      value={newPrjEq.kksCode}
                      onChange={(e) => setNewPrjEq({ ...newPrjEq, kksCode: e.target.value })}
                      className="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg font-mono font-bold text-indigo-700"
                    />
                  </div>
                  <div className="sm:col-span-4">
                    <label className="block font-semibold text-slate-700 mb-0.5">Tên Thiết Bị *</label>
                    <input
                      type="text"
                      required
                      placeholder="Máy biến áp T1 110/22kV 63MVA"
                      value={newPrjEq.name}
                      onChange={(e) => setNewPrjEq({ ...newPrjEq, name: e.target.value })}
                      className="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg font-medium"
                    />
                  </div>
                  <div className="sm:col-span-2">
                    <label className="block font-semibold text-slate-700 mb-0.5">Phân loại NETA *</label>
                    <select
                      value={newPrjEq.type}
                      onChange={(e) => setNewPrjEq({ ...newPrjEq, type: e.target.value })}
                      className="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg"
                    >
                      <option value="Máy biến áp">Máy biến áp</option>
                      <option value="Máy cắt">Máy cắt</option>
                      <option value="Cáp điện">Cáp điện</option>
                      <option value="Động cơ">Động cơ</option>
                      <option value="Tủ điện">Tủ điện</option>
                      <option value="Máy phát">Máy phát</option>
                      <option value="Pin & UPS">Pin & UPS</option>
                      <option value="Tiếp địa">Tiếp địa</option>
                      <option value="Rơ le bảo vệ">Rơ le bảo vệ</option>
                    </select>
                  </div>
                  <div className="sm:col-span-2">
                    <label className="block font-semibold text-slate-700 mb-0.5">Vị Trí / Ngăn Lộ</label>
                    <input
                      type="text"
                      placeholder="Ngăn 131"
                      value={newPrjEq.location}
                      onChange={(e) => setNewPrjEq({ ...newPrjEq, location: e.target.value })}
                      className="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg"
                    />
                  </div>
                  <div className="sm:col-span-1 flex items-end">
                    <button
                      type="submit"
                      className="w-full py-1.5 bg-indigo-700 hover:bg-indigo-800 text-white rounded-lg font-bold flex items-center justify-center shadow-xs"
                      title="Thêm thiết bị vào dự án"
                    >
                      <Plus size={16} />
                    </button>
                  </div>
                </form>
              </div>

              {/* PMP Workflow Guide Strip */}
              <div className="bg-gradient-to-r from-blue-50 via-indigo-50 to-slate-50 border border-indigo-200 rounded-xl p-3 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 text-xs shadow-2xs">
                <div className="flex items-center gap-2.5">
                  <div className="p-2 bg-indigo-700 text-white rounded-lg shrink-0">
                    <ShieldCheck size={16} />
                  </div>
                  <div>
                    <div className="font-bold text-indigo-950 flex items-center gap-1.5">
                      <span>Quy Trình Hiện Trường Chuẩn PMP Cho Kỹ Sư FSE</span>
                      <span className="text-[10px] bg-indigo-100 text-indigo-800 font-bold px-1.5 py-0.2 rounded">NETA ATS / MST</span>
                    </div>
                    <p className="text-[11px] text-slate-600 mt-0.5">
                      1. Chọn chuẩn kiểm định (ATS hoặc MST) → 2. Nhấn nút <strong>"Nhập Form NETA"</strong> để điền biểu mẫu (khách hàng & thiết bị tự động nạp sẵn) → 3. Bên dưới biểu mẫu có nút chuyển tức thì sang xem <strong>Báo Cáo Chuẩn TEV AI</strong>.
                    </p>
                  </div>
                </div>
                <div className="text-[11px] bg-amber-100 text-amber-950 border border-amber-300 font-bold px-2.5 py-1 rounded-lg shrink-0 flex items-center gap-1">
                  <Droplets size={13} className="text-amber-700" />
                  <span>MTS hiện có 01 form MBA ngâm dầu</span>
                </div>
              </div>

              {/* Categorized Professional Equipment Table for Project */}
              <ProjectCategorizedEquipmentTable
                project={managingProjectEquipment}
                equipmentList={managingProjectEquipment.equipmentList || []}
                equipmentStandardMap={equipmentStandardMap}
                onSetStandard={(key, std) => setEquipmentStandardMap(prev => ({ ...prev, [key]: std }))}
                onNavigateToNeta={handleNavigateToNeta}
                onRemoveEquipment={handleRemoveEquipmentFromProject}
                readOnlyActions={false}
              />
            </div>

            {/* Footer */}
            <div className="p-3 bg-slate-50 border-t border-slate-200 flex items-center justify-between shrink-0 rounded-b-2xl">
              <span className="text-xs text-slate-500">
                Kỹ sư tại hiện trường chỉ cần nhập mã <strong>{managingProjectEquipment.id}</strong> là hệ thống tự động đổ {managingProjectEquipment.equipmentList?.length || 0} thiết bị này vào dropdown.
              </span>
              <button
                type="button"
                onClick={() => setManagingProjectEquipment(null)}
                className="px-4 py-1.5 bg-indigo-700 hover:bg-indigo-800 text-white rounded-lg font-bold text-xs"
              >
                Xong & Đóng
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
