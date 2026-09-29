import React, { useState, useEffect, useMemo } from 'react';
import {
  LayoutDashboard,
  Server,
  FileText,
  Settings,
  Bell,
  Search,
  AlertTriangle,
  CheckCircle,
  Activity,
  MapPin,
  Factory,
  ChevronDown,
  ChevronRight,
  MoreVertical,
  TrendingUp,
  TrendingDown,
  Clock,
  AlertCircle,
  ClipboardList,
  QrCode,
  Camera,
  UploadCloud,
  Save,
  Send,
  Thermometer,
  Zap,
  History,
  BarChart2,
  ArrowUpRight,
  ArrowDownRight,
  Minus,
  Plus,
  X,
  Download,
  Upload,
  Eye,
  Calendar,
  PieChart as PieChartIcon,
  BarChart as BarChartIcon,
  Wind,
  Sun,
  Droplets,
  Gauge,
  RefreshCw,
  RotateCcw,
  LogOut,
  User,
  Users,
  Paperclip,
  Info,
  Database,
  Menu,
  Copy,
  Package,
  Trash2,
  ShieldCheck,
  ShieldAlert,
  ExternalLink,
  Mail,
  BarChart3,
  Globe,
  Navigation,
  Layers,
  FileSpreadsheet,
  ClipboardCheck
} from 'lucide-react';
import {
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  ReferenceLine,
  Area,
  AreaChart,
  BarChart,
  Bar,
  PieChart,
  Pie,
  Cell,
  Legend,
  Radar,
  RadarChart,
  PolarGrid,
  PolarAngleAxis,
  PolarRadiusAxis
} from 'recharts';
import { MapContainer, TileLayer, Marker, Popup, Tooltip as LeafletTooltip, useMap } from 'react-leaflet';
import L from 'leaflet';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as htmlToImage from 'html-to-image';
import { auth, db, signInWithGoogle, logOut } from './firebase';
import { onAuthStateChanged } from 'firebase/auth';
import { doc, getDoc, setDoc, collection, query, where, onSnapshot, orderBy, limit, addDoc, updateDoc, deleteDoc } from 'firebase/firestore';
import { Html5QrcodeScanner } from 'html5-qrcode';
import { QRCodeSVG } from 'qrcode.react';
import { 
  calculateSwitchgearHealth, 
  calculateTransformerHealth, 
  calculateMotorHealth,
  SwitchgearParams,
  TransformerParams,
  MotorDiagnosticTests,
  MotorPhaseData
} from './lib/healthCalculator';
import { NetaDryTypeChecklist } from './components/NetaDryTypeChecklist';
import { NetaLargeDryTypeChecklist } from './components/NetaLargeDryTypeChecklist';
import { NetaBessChecklist } from './components/NetaBessChecklist';
import { NetaPvChecklist } from './components/NetaPvChecklist';
import { PvCellDeepThermalAnalyzer } from './components/PvCellDeepThermalAnalyzer';
import { NetaEvseChecklist } from './components/NetaEvseChecklist';
import { NetaCableChecklist } from './components/NetaCableChecklist';
import { NetaCableLvChecklist } from './components/NetaCableLvChecklist';
import { NetaLvBreakerChecklist } from './components/NetaLvBreakerChecklist';
import { NetaGroundingChecklist } from './components/NetaGroundingChecklist';
import { NetaDcMotorChecklist } from './components/NetaDcMotorChecklist';
import { NetaSf6SwitchChecklist } from './components/NetaSf6SwitchChecklist';
import { NetaThermographyChecklist } from './components/NetaThermographyChecklist';
import { NetaPartialDischargeChecklist } from './components/NetaPartialDischargeChecklist';
import { NetaMotorChecklist } from './components/NetaMotorChecklist';
import { NetaLiquidTransformerChecklist } from './components/NetaLiquidTransformerChecklist';
import { DgaHistoricalTrendChart } from './components/DgaHistoricalTrendChart';
import { DgaAdvancedDiagnosticsDashboard } from './components/DgaAdvancedDiagnosticsDashboard';
import { DgaParamTrendInput } from './components/dga/DgaParamTrendInput';
import { getLastRecordedDgaValues } from './utils/dgaHistoryService';
import { NetaRelayChecklist } from './components/NetaRelayChecklist';
import { NetaUpsChecklist } from './components/NetaUpsChecklist';
import { NetaSyncMachineryChecklist } from './components/NetaSyncMachineryChecklist';
import { NetaBatteryFloodedChecklist } from './components/NetaBatteryFloodedChecklist';
import { NetaBatteryVrlaChecklist } from './components/NetaBatteryVrlaChecklist';
import { NetaAtsChecklist } from './components/NetaAtsChecklist';
import { NetaEngineGeneratorChecklist } from './components/NetaEngineGeneratorChecklist';
import { NetaSwitchgearChecklist } from './components/NetaSwitchgearChecklist';
import { PmScheduleDashboard } from './components/PmScheduleDashboard';
import { EquipmentChecklistSelector, FieldEntryMode } from './components/EquipmentChecklistSelector';
import { 
  FseFieldDataEntryStation,
  CustomerItem,
  EquipmentItem,
  WorkOrderItem 
} from './components/FseFieldDataEntryStation';
import {
  compareVersionsBetweenFirestoreAndSheets,
  ReconciliationReport,
  ReconciliationItem,
  formatVersionDate,
  parseVersionTimestamp
} from './utils/syncReconciliationService';
import { SyncReconciliationModal } from './components/SyncReconciliationModal';
import {
  DEFAULT_SIMULATION_REPORTS,
  DEFAULT_CUSTOMERS,
  DEFAULT_EQUIPMENT,
  DEFAULT_WORK_ORDERS,
  DEFAULT_INVENTORY,
  getInitialEquipmentList,
  getInitialReportsList
} from './utils/defaultDataService';

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId: string | undefined;
    email: string | null | undefined;
    emailVerified: boolean | undefined;
    isAnonymous: boolean | undefined;
    tenantId: string | null | undefined;
    providerInfo: {
      providerId: string;
      displayName: string | null;
      email: string | null;
      photoUrl: string | null;
    }[];
  }
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId,
      providerInfo: auth.currentUser?.providerData.map(provider => ({
        providerId: provider.providerId,
        displayName: provider.displayName,
        email: provider.email,
        photoUrl: provider.photoURL
      })) || []
    },
    operationType,
    path
  }
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

// --- Mock Data for Maintenance Analytics ---
const mtbfData = [
  { month: 'Jan', hours: 320, target: 350 },
  { month: 'Feb', hours: 330, target: 350 },
  { month: 'Mar', hours: 340, target: 350 },
  { month: 'Apr', hours: 345, target: 350 },
  { month: 'May', hours: 300, target: 350 },
  { month: 'Jun', hours: 310, target: 350 },
  { month: 'Jul', hours: 330, target: 350 },
  { month: 'Aug', hours: 340, target: 350 },
  { month: 'Sep', hours: 345, target: 350 },
  { month: 'Oct', hours: 370, target: 350 },
  { month: 'Nov', hours: 365, target: 350 },
  { month: 'Dec', hours: 355, target: 350 },
];

const mttrData = [
  { month: 'Jan', hours: 4.0, target: 3.5 },
  { month: 'Feb', hours: 3.8, target: 3.5 },
  { month: 'Mar', hours: 3.9, target: 3.5 },
  { month: 'Apr', hours: 3.8, target: 3.5 },
  { month: 'May', hours: 3.5, target: 3.5 },
  { month: 'Jun', hours: 3.6, target: 3.5 },
  { month: 'Jul', hours: 3.2, target: 3.5 },
  { month: 'Aug', hours: 2.5, target: 3.5 },
  { month: 'Sep', hours: 2.4, target: 3.5 },
  { month: 'Oct', hours: 2.7, target: 3.5 },
  { month: 'Nov', hours: 2.7, target: 3.5 },
  { month: 'Dec', hours: 2.6, target: 3.5 },
];

const anomaliesData = [
  { name: 'Damaged Lens', value: 85 },
  { name: 'Cooling Failure', value: 45 },
  { name: 'Valve Failure', value: 35 },
  { name: 'Limit Switch', value: 25 },
  { name: 'Fan Failure', value: 18 },
  { name: 'Connectivity', value: 12 },
];

const maintenanceCostData = [
  { month: 'Jan', cost: 6.0, target: 5.0 },
  { month: 'Feb', cost: 6.0, target: 5.0 },
  { month: 'Mar', cost: 5.0, target: 5.0 },
  { month: 'Apr', cost: 5.0, target: 5.0 },
  { month: 'May', cost: 3.7, target: 5.0 },
  { month: 'Jun', cost: 3.3, target: 5.0 },
  { month: 'Jul', cost: 2.4, target: 5.0 },
  { month: 'Aug', cost: 2.2, target: 5.0 },
  { month: 'Sep', cost: 1.6, target: 5.0 },
  { month: 'Oct', cost: 1.5, target: 5.0 },
  { month: 'Nov', cost: 1.3, target: 5.0 },
  { month: 'Dec', cost: 0.8, target: 5.0 },
];

const adherenceData = [
  { name: 'Adherence', value: 82, color: '#10b981' },
  { name: 'Fail', value: 18, color: '#ef4444' },
];

const MaintenanceCharts = ({ equipment, workOrders, reports }: { equipment: any, workOrders: any[], reports: any[] }) => {
  // Filter data for this specific equipment
  const eqWorkOrders = workOrders.filter(wo => wo.equipmentId.includes(equipment.id));
  const eqReports = reports.filter(r => r.equipmentId === equipment.id);

  // Helper to get month name
  const getMonthName = (dateStr: string) => {
    try {
      if (!dateStr) return '---';
      const parts = dateStr.split('/');
      if (parts.length < 2) return '---';
      const monthIdx = parseInt(parts[1]) - 1;
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      return months[monthIdx] || '---';
    } catch {
      return '---';
    }
  };

  // 1. Calculate OEE Score (Mocked based on health but slightly dynamic)
  const oeeScore = equipment.health ? Math.min(100, Math.max(0, equipment.health - 5 + Math.floor(Math.random() * 10))) : 76;

  // 2. MTBF Evolution
  const monthsList = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const currentMonthIdx = new Date().getMonth();
  
  const mtbfEvolution = monthsList.slice(0, Math.max(6, currentMonthIdx + 1)).map((month, idx) => {
    const monthOrders = eqWorkOrders.filter(wo => getMonthName(wo.dueDate) === month && (wo.type === 'corrective' || wo.type === 'emergency'));
    let hours = monthOrders.length > 0 ? 350 - (monthOrders.length * 50) : 340 + (idx * 5);
    return { month, hours: Math.max(100, hours), target: 350 };
  });

  // 3. MTTR Evolution
  const mttrEvolution = monthsList.slice(0, Math.max(6, currentMonthIdx + 1)).map((month, idx) => {
    const monthOrders = eqWorkOrders.filter(wo => getMonthName(wo.dueDate) === month && wo.status === 'completed');
    let avgTime = monthOrders.length > 0 
      ? monthOrders.reduce((acc, curr) => acc + (curr.actualTimeSpent || 0), 0) / monthOrders.length 
      : 3.5 - (idx * 0.1);
    return { month, hours: parseFloat(avgTime.toFixed(1)), target: 3.5 };
  });

  // 4. Top Anomalies
  const failureCounts: Record<string, number> = {};
  eqWorkOrders.forEach(wo => {
    if (wo.failureCode) {
      failureCounts[wo.failureCode] = (failureCounts[wo.failureCode] || 0) + 1;
    }
  });
  
  let anomaliesDataList = Object.entries(failureCounts).map(([name, count]) => ({
    name: name.charAt(0).toUpperCase() + name.slice(1),
    value: count * 20
  })).sort((a, b) => b.value - a.value);

  if (anomaliesDataList.length === 0) {
    anomaliesDataList = [
      { name: 'Nhiệt độ dầu', value: equipment.status === 'critical' ? 85 : 15 },
      { name: 'Rò rỉ nhẹ', value: 30 },
      { name: 'Độ ẩm dầu', value: 20 }
    ];
  }

  // 5. Maintenance Cost
  const costEvolution = monthsList.slice(0, Math.max(6, currentMonthIdx + 1)).map((month, idx) => {
    const monthOrders = eqWorkOrders.filter(wo => getMonthName(wo.dueDate) === month);
    let monthCost = monthOrders.reduce((acc, curr) => acc + (curr.laborCost || 0) + (curr.partCost || 0), 0) / 1000;
    if (monthCost === 0) monthCost = 0.5 + (idx * 0.1); 
    return { month, cost: parseFloat(monthCost.toFixed(2)), target: 5.0 };
  });

  // 6. Adherence to Schedule
  const completedOrders = eqWorkOrders.filter(wo => wo.status === 'completed').length;
  const totalOrders = eqWorkOrders.length;
  const adherenceScore = totalOrders > 0 ? Math.round((completedOrders / totalOrders) * 100) : 85;

  return (
    <div className="space-y-8">
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {/* OEE Gauge */}
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
          <h4 className="text-sm font-bold text-slate-700 mb-4 text-center uppercase tracking-wider">OEE Performance</h4>
          <div className="h-40 flex flex-col items-center justify-center relative">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie
                  data={[
                    { value: oeeScore, fill: oeeScore > 70 ? '#10b981' : oeeScore > 40 ? '#f59e0b' : '#ef4444' },
                    { value: 100 - oeeScore, fill: '#f1f5f9' }
                  ]}
                  cx="50%"
                  cy="100%"
                  startAngle={180}
                  endAngle={0}
                  innerRadius={60}
                  outerRadius={80}
                  paddingAngle={0}
                  dataKey="value"
                >
                </Pie>
              </PieChart>
            </ResponsiveContainer>
            <div className="absolute bottom-2 text-center">
              <div className="text-2xl font-bold text-slate-800">{oeeScore}%</div>
              <div className="text-[10px] text-slate-500 font-medium uppercase">OEE SCORE</div>
            </div>
          </div>
          <div className="grid grid-cols-3 gap-2 mt-2 border-t border-slate-100 pt-3">
            <div className="text-center">
              <div className="text-[10px] text-slate-500 uppercase">MTBF</div>
              <div className="text-xs font-bold text-emerald-600">
                {mtbfEvolution.length > 0 ? mtbfEvolution[mtbfEvolution.length - 1].hours : '---'} hrs
              </div>
            </div>
            <div className="text-center border-x border-slate-100">
              <div className="text-[10px] text-slate-500 uppercase">MTTR</div>
              <div className="text-xs font-bold text-rose-600">
                {mttrEvolution.length > 0 ? mttrEvolution[mttrEvolution.length - 1].hours : '---'} hrs
              </div>
            </div>
            <div className="text-center">
              <div className="text-[10px] text-slate-500 uppercase">AM STEP</div>
              <div className="text-xs font-bold text-blue-600">3</div>
            </div>
          </div>
        </div>

        {/* MTBF Evolution */}
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
          <h4 className="text-sm font-bold text-slate-700 mb-4 uppercase tracking-wider">MTBF Evolution (hrs)</h4>
          <div className="h-48">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={mtbfEvolution}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                <XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#94a3b8' }} />
                <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#94a3b8' }} />
                <Tooltip 
                  contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                  cursor={{ fill: '#f8fafc' }}
                />
                <Bar dataKey="hours" radius={[4, 4, 0, 0]}>
                  {mtbfEvolution.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={entry.hours >= entry.target ? '#10b981' : entry.hours >= entry.target * 0.9 ? '#f59e0b' : '#ef4444'} />
                  ))}
                </Bar>
                <ReferenceLine y={350} stroke="#94a3b8" strokeDasharray="3 3" label={{ position: 'right', value: 'Target', fill: '#94a3b8', fontSize: 10 }} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* MTTR Evolution */}
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
          <h4 className="text-sm font-bold text-slate-700 mb-4 uppercase tracking-wider">MTTR Evolution (hrs)</h4>
          <div className="h-48">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={mttrEvolution}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                <XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#94a3b8' }} />
                <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#94a3b8' }} />
                <Tooltip 
                  contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                  cursor={{ fill: '#f8fafc' }}
                />
                <Bar dataKey="hours" radius={[4, 4, 0, 0]}>
                  {mttrEvolution.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={entry.hours <= entry.target ? '#10b981' : entry.hours <= entry.target * 1.1 ? '#f59e0b' : '#ef4444'} />
                  ))}
                </Bar>
                <ReferenceLine y={3.5} stroke="#94a3b8" strokeDasharray="3 3" label={{ position: 'right', value: 'Target', fill: '#94a3b8', fontSize: 10 }} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Top Anomalies */}
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
          <h4 className="text-sm font-bold text-slate-700 mb-4 uppercase tracking-wider">Top Anomalies</h4>
          <div className="h-48">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={anomaliesDataList} layout="vertical" margin={{ left: 40 }}>
                <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                <XAxis type="number" hide />
                <YAxis dataKey="name" type="category" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} width={80} />
                <Tooltip 
                  contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                />
                <Bar dataKey="value" fill="#0369a1" radius={[0, 4, 4, 0]} barSize={15} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Maintenance Cost */}
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
          <h4 className="text-sm font-bold text-slate-700 mb-4 uppercase tracking-wider">Maintenance Cost (k$)</h4>
          <div className="h-48">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={costEvolution}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                <XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#94a3b8' }} />
                <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#94a3b8' }} />
                <Tooltip 
                  contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                  cursor={{ fill: '#f8fafc' }}
                />
                <Bar dataKey="cost" radius={[4, 4, 0, 0]}>
                  {costEvolution.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={entry.cost <= entry.target ? '#10b981' : '#f59e0b'} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Adherence to Schedule */}
        <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
          <h4 className="text-sm font-bold text-slate-700 mb-4 uppercase tracking-wider text-center">Adherence to Schedule</h4>
          <div className="h-48 relative">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie
                  data={[
                    { name: 'Completed', value: adherenceScore, fill: '#10b981' },
                    { name: 'Remaining', value: 100 - adherenceScore, fill: '#f1f5f9' }
                  ]}
                  cx="50%"
                  cy="50%"
                  innerRadius={50}
                  outerRadius={70}
                  paddingAngle={5}
                  dataKey="value"
                >
                </Pie>
              </PieChart>
            </ResponsiveContainer>
            <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 text-center pointer-events-none">
              <div className="text-xl font-bold text-slate-800">{adherenceScore}%</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

// --- PDF Helper ---
let robotoRegularBase64: string | null = null;
let robotoBoldBase64: string | null = null;

const getConfiguredJsPDF = async (orientation: 'p' | 'l' = 'p') => {
  const pdf = new jsPDF(orientation, 'mm', 'a4');
  
  try {
    if (!robotoRegularBase64) {
      const resReg = await fetch('https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Regular.ttf');
      if (resReg.ok) {
        const bufReg = await resReg.arrayBuffer();
        robotoRegularBase64 = btoa(new Uint8Array(bufReg).reduce((data, byte) => data + String.fromCharCode(byte), ''));
      }
    }
    
    if (!robotoBoldBase64) {
      const resBold = await fetch('https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Medium.ttf');
      if (resBold.ok) {
        const bufBold = await resBold.arrayBuffer();
        robotoBoldBase64 = btoa(new Uint8Array(bufBold).reduce((data, byte) => data + String.fromCharCode(byte), ''));
      }
    }
    
    if (robotoRegularBase64) {
      pdf.addFileToVFS('Roboto-Regular.ttf', robotoRegularBase64);
      pdf.addFont('Roboto-Regular.ttf', 'Roboto', 'normal');
    }
    
    if (robotoBoldBase64) {
      pdf.addFileToVFS('Roboto-Bold.ttf', robotoBoldBase64);
      pdf.addFont('Roboto-Bold.ttf', 'Roboto', 'bold');
    }
    
    if (robotoRegularBase64 || robotoBoldBase64) {
      pdf.setFont('Roboto');
    }
  } catch (error) {
    console.error('Failed to load custom fonts:', error);
    // Fallback to default font
  }
  
  return pdf;
};

const paramLabels: Record<string, string> = {
  oilTemp: 'Nhiệt độ dầu (°C)',
  windingTemp: 'Nhiệt độ cuộn dây (°C)',
  irHighLow: 'Điện trở cách điện Cao-Thấp (MΩ)',
  irHighEarth: 'Điện trở cách điện Cao-Đất (MΩ)',
  oilLeak: 'Rò rỉ dầu',
  dga: 'Phân tích khí hòa tan (DGA)',
  dielectricStrength: 'Độ bền điện môi (kV)',
  furan: 'Hàm lượng Furan (ppm)',
  oilMoisture: 'Độ ẩm trong dầu (ppm)',
  thermography: 'Nhiệt độ tiếp xúc (°C)',
  contactRes: 'Điện trở tiếp xúc (μΩ)',
  tev: 'Phóng điện cục bộ TEV (dBmV)',
  ultrasonic: 'Phóng điện cục bộ Siêu âm (dBμV)',
  tevPulses: 'Số xung TEV/chu kỳ',
  humidity: 'Độ ẩm môi trường (%)',
  sf6Pressure: 'Áp suất khí SF6 (bar)',
  vibration: 'Độ rung (mm/s)',
  statorTemp: 'Nhiệt độ Stator (°C)',
  ir: 'Điện trở cách điện (MΩ)',
  pd: 'Phóng điện cục bộ (pC)',
  voltageImbalance: 'Mất cân bằng điện áp (%)',
  pi: 'Chỉ số phân cực (PI)',
  bearingTemp: 'Nhiệt độ ổ trục (°C)',
  tanDelta: 'Tổn hao điện môi (Tan Delta)'
};

const paramStandards: Record<string, string> = {
  oilTemp: '< 85 °C',
  windingTemp: '< 95 °C',
  irHighLow: '> 1000 MΩ',
  irHighEarth: '> 1000 MΩ',
  dielectricStrength: '> 30 kV',
  furan: '< 1.0 ppm',
  oilMoisture: '< 20 ppm',
  thermography: '< 70 °C',
  contactRes: '< 100 μΩ',
  tev: '< 20 dBmV',
  ultrasonic: '< 10 dBμV',
  vibration: '< 2.8 mm/s',
  statorTemp: '< 120 °C',
  ir: '> 100 MΩ',
  pd: '< 500 pC',
  voltageImbalance: '< 2.0 %',
  pi: '> 2.0',
  bearingTemp: '< 80 °C',
  tanDelta: '< 0.5 %'
};

const evaluateParam = (type: string, param: string, value: any): string => {
  const status = evaluateEquipmentParam(type, param, value);
  if (!status) return 'N/A';
  return status === 'critical' ? 'Nguy hiểm' : status === 'warning' ? 'Cảnh báo' : 'Bình thường';
};

const generateIndividualReportPDF = async (reportData: any, saveAsFile = false, historicalReports: any[] = []) => {
  const pdf = await getConfiguredJsPDF();
  
  // Header
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'bold');
  pdf.setFontSize(24);
  pdf.setTextColor(30, 58, 138); // blue-900
  pdf.text('TEV', 20, 25);
  
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'normal');
  pdf.setFontSize(10);
  pdf.setTextColor(29, 78, 216); // blue-700
  pdf.text('ASSET INSPECTION', 20, 32);
  
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'bold');
  pdf.setFontSize(18);
  pdf.setTextColor(30, 41, 59); // slate-800
  pdf.text('BÁO CÁO KIỂM TRA', 190, 25, { align: 'right' });
  
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'normal');
  pdf.setFontSize(12);
  pdf.setTextColor(100, 116, 139); // slate-500
  pdf.text('INSPECTION REPORT', 190, 32, { align: 'right' });
  
  // Line separator
  pdf.setDrawColor(30, 58, 138);
  pdf.setLineWidth(0.5);
  pdf.line(20, 38, 190, 38);
  
  // Info Grid
  let yPos = 50;
  pdf.setFontSize(10);
  pdf.setTextColor(100, 116, 139);
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'bold');
  
  pdf.text('Mã báo cáo / Report ID', 20, yPos);
  pdf.text('Ngày / Date', 105, yPos);
  yPos += 5;
  
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'normal');
  pdf.setFontSize(12);
  pdf.setTextColor(15, 23, 42); // slate-900
  
  pdf.text(reportData.id || '', 20, yPos);
  pdf.text(reportData.date || '', 105, yPos);
  yPos += 5;
  
  pdf.setFontSize(8);
  pdf.setTextColor(148, 163, 184); // slate-400
  pdf.text(`Thời gian tạo: ${new Date().toLocaleString('vi-VN')}`, 20, yPos);
  yPos += 5;
  
  pdf.setFontSize(10);
  pdf.setTextColor(100, 116, 139);
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'bold');
  pdf.text('Thiết bị / Equipment', 20, yPos);
  pdf.text('Nhà máy / Site', 105, yPos);
  yPos += 5;
  
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'normal');
  pdf.setFontSize(12);
  pdf.setTextColor(15, 23, 42); // slate-900
  
  const equipmentText = `${reportData.equipmentName || ''} (${reportData.equipmentId || ''})`;
  const splitEquipment = pdf.splitTextToSize(equipmentText, 80);
  const factoryText = reportData.factory || '';
  const splitFactory = pdf.splitTextToSize(factoryText, 80);
  
  pdf.text(splitEquipment, 20, yPos);
  pdf.text(splitFactory, 105, yPos);
  
  const maxLines = Math.max(splitEquipment.length, splitFactory.length);
  yPos += (maxLines * 5) + 5;
  
  pdf.setFontSize(10);
  pdf.setTextColor(100, 116, 139);
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'bold');
  pdf.text('Người thực hiện / Inspector', 20, yPos);
  pdf.text('Đánh giá / Condition', 105, yPos);
  yPos += 5;
  
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'normal');
  pdf.setFontSize(12);
  pdf.setTextColor(15, 23, 42); // slate-900
  pdf.text(reportData.inspector || '', 20, yPos);
  
  // Status Badge
  const status = reportData.status;
  let statusText = 'Bình thường / Normal';
  let statusColor = [16, 185, 129]; // emerald-500
  
  if (status === 'warning') {
    statusText = 'Cảnh báo / Warning';
    statusColor = [245, 158, 11]; // amber-500
  } else if (status === 'critical' || status === 'danger') {
    statusText = 'Nguy hiểm / Severe';
    statusColor = [225, 29, 72]; // rose-600
  }
  
  pdf.setTextColor(statusColor[0], statusColor[1], statusColor[2]);
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'bold');
  pdf.text(statusText, 105, yPos);
  yPos += 15;
  
  // Notes
  pdf.setFontSize(10);
  pdf.setTextColor(100, 116, 139);
  pdf.text('Ghi chú / Notes', 20, yPos);
  yPos += 7;
  
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'normal');
  pdf.setFontSize(11);
  pdf.setTextColor(15, 23, 42);
  
  const notes = reportData.notes || '';
  const splitNotes = pdf.splitTextToSize(notes, 170);
  pdf.text(splitNotes, 20, yPos);
  
  yPos += (splitNotes.length * 5) + 10;

  // Health Index Section
  pdf.setFontSize(10);
  pdf.setTextColor(100, 116, 139);
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'bold');
  pdf.text('Chỉ số sức khỏe / Health Index', 20, yPos);
  yPos += 8;
  
  const health = reportData.health || 0;
  pdf.setDrawColor(226, 232, 240); // slate-200
  pdf.setFillColor(248, 250, 252); // slate-50
  pdf.roundedRect(20, yPos, 170, 20, 3, 3, 'FD');
  
  pdf.setFontSize(24);
  pdf.setTextColor(statusColor[0], statusColor[1], statusColor[2]);
  pdf.text(`${health}%`, 35, yPos + 14);
  
  pdf.setFontSize(10);
  pdf.setTextColor(100, 116, 139);
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'normal');
  pdf.text('Tình trạng thiết bị dựa trên phân tích dữ liệu hiện tại và lịch sử vận hành.', 75, yPos + 12);
  
  yPos += 30;
  
  // Measurements
  if (reportData.measurements && Object.keys(reportData.measurements).length > 0) {
    pdf.setFontSize(10);
    pdf.setTextColor(100, 116, 139);
    pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'bold');
    pdf.text('Thông số đo lường / Measurements', 20, yPos);
    yPos += 8;
    
    const tableData = Object.entries(reportData.measurements)
      .filter(([k, v]) => v !== '' && v !== undefined && v !== null && k !== 'age' && k !== 'dutyFactor')
      .map(([k, v]) => {
        const standard = paramStandards[k] || 'N/A';
        const status = evaluateEquipmentParam(reportData.type, k, v);
        const evalResult = status === 'critical' ? 'Nguy hiểm' : status === 'warning' ? 'Cảnh báo' : status === 'healthy' ? 'Bình thường' : 'N/A';
        return [paramLabels[k] || String(k), String(v), standard, evalResult];
      });
      
    if (tableData.length > 0) {
      autoTable(pdf, {
        startY: yPos,
        head: [['Thông số / Parameter', 'Giá trị / Value', 'Tiêu chuẩn / Standard', 'Đánh giá / Evaluation']],
        body: tableData,
        theme: 'grid',
        styles: { font: robotoRegularBase64 ? 'Roboto' : 'helvetica', fontSize: 10 },
        headStyles: { fillColor: [241, 245, 249], textColor: [71, 85, 105], fontStyle: 'bold' },
        columnStyles: {
          3: { fontStyle: 'bold' } // Make evaluation column bold
        },
        didParseCell: function(data) {
          if (data.section === 'body' && data.column.index === 3) {
            if (data.cell.raw === 'Nguy hiểm') {
              data.cell.styles.textColor = [220, 38, 38]; // red-600
            } else if (data.cell.raw === 'Cảnh báo') {
              data.cell.styles.textColor = [217, 119, 6]; // amber-600
            } else if (data.cell.raw === 'Bình thường') {
              data.cell.styles.textColor = [5, 150, 105]; // emerald-600
            }
          }
        },
        margin: { left: 20, right: 20 }
      });
      yPos = (pdf as any).lastAutoTable.finalY + 15;
    }
  }

  // Trending Section (Upgraded)
  if (historicalReports && historicalReports.length > 1 && reportData.measurements) {
    // Filter out the current report and any newer reports
    const currentReportDate = reportData.date ? new Date(reportData.date.split('/').reverse().join('-')).getTime() : Date.now();
    const pastReports = historicalReports.filter(hr => {
      if (hr.id === reportData.id) return false;
      const hrDate = hr.date ? new Date(hr.date.split('/').reverse().join('-')).getTime() : 0;
      return hrDate <= currentReportDate;
    });
    
    const paramsWithHistory = Object.keys(reportData.measurements).filter(k => 
      k !== 'age' && k !== 'dutyFactor' &&
      pastReports.some(hr => hr.measurements && hr.measurements[k] !== undefined)
    );

    if (paramsWithHistory.length > 0) {
      if (yPos > 240) { pdf.addPage(); yPos = 20; }
      
      pdf.setFontSize(10);
      pdf.setTextColor(100, 116, 139);
      pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'bold');
      pdf.text('Phân tích xu hướng / Trend Analysis', 20, yPos);
      yPos += 8;

      const trendTableData = paramsWithHistory.map(k => {
        const currentVal = parseFloat(reportData.measurements[k]);
        
        // Sort historical reports by date (oldest to newest)
        const sortedHistory = [...pastReports].sort((a, b) => {
          const [d1, m1, y1] = a.date.split('/');
          const [d2, m2, y2] = b.date.split('/');
          return new Date(`${y1}-${m1}-${d1}`).getTime() - new Date(`${y2}-${m2}-${d2}`).getTime();
        });

        // Get the most recent previous value
        let prevVal = 'N/A';
        let trend = 'Ổn định';
        let trendIcon = '→';
        
        for (let i = sortedHistory.length - 1; i >= 0; i--) {
          if (sortedHistory[i].measurements && sortedHistory[i].measurements[k] !== undefined) {
            const val = parseFloat(sortedHistory[i].measurements[k]);
            if (!isNaN(val)) {
              prevVal = String(val);
              if (!isNaN(currentVal)) {
                const diff = currentVal - val;
                if (diff > 0.05 * val) { trend = 'Tăng'; trendIcon = '↑'; }
                else if (diff < -0.05 * val) { trend = 'Giảm'; trendIcon = '↓'; }
              }
              break;
            }
          }
        }
        
        return [paramLabels[k] || String(k), prevVal, String(reportData.measurements[k]), `${trendIcon} ${trend}`];
      });

      autoTable(pdf, {
        startY: yPos,
        head: [['Thông số / Parameter', 'Lần trước / Previous', 'Hiện tại / Current', 'Xu hướng / Trend']],
        body: trendTableData,
        theme: 'striped',
        styles: { font: robotoRegularBase64 ? 'Roboto' : 'helvetica', fontSize: 10 },
        headStyles: { fillColor: [248, 250, 252], textColor: [71, 85, 105], fontStyle: 'bold' },
        didParseCell: function(data) {
          if (data.section === 'body' && data.column.index === 3) {
            const rawValue = String(data.cell.raw || '');
            if (rawValue.includes('↑')) {
              data.cell.styles.textColor = [220, 38, 38]; // red-600
            } else if (rawValue.includes('↓')) {
              data.cell.styles.textColor = [5, 150, 105]; // emerald-600
            }
          }
        },
        margin: { left: 20, right: 20 }
      });
      yPos = (pdf as any).lastAutoTable.finalY + 15;
    }
  }

  // Recommendations
  pdf.setFontSize(10);
  pdf.setTextColor(100, 116, 139);
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'bold');
  pdf.text('Khuyến cáo / Recommendations', 20, yPos);
  yPos += 8;
  
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'normal');
  pdf.setFontSize(11);
  pdf.setTextColor(15, 23, 42);
  
  let recommendationText = 'Tiếp tục vận hành bình thường. Thực hiện kiểm tra định kỳ theo kế hoạch.';
  if (status === 'warning') {
    recommendationText = 'Cần theo dõi chặt chẽ các thông số bất thường. Lên kế hoạch kiểm tra chuyên sâu trong vòng 1-3 tháng tới.';
  } else if (status === 'critical' || status === 'danger') {
    recommendationText = 'NGUY HIỂM: Cần tách thiết bị khỏi lưới điện để kiểm tra và sửa chữa ngay lập tức. Nguy cơ sự cố cao.';
  }
  
  const splitRecs = pdf.splitTextToSize(recommendationText, 170);
  pdf.text(splitRecs, 20, yPos);
  yPos += (splitRecs.length * 5) + 10;
  
  // Footer
  pdf.setFontSize(9);
  pdf.setTextColor(148, 163, 184);
  pdf.text('Được tạo tự động bởi hệ thống TEV Asset Management', 105, 285, { align: 'center' });
  pdf.text('Automatically generated by TEV Asset Management System', 105, 290, { align: 'center' });
  
  if (saveAsFile) {
    pdf.save(`${reportData.id}.pdf`);
    return null;
  } else {
    return pdf.output('blob');
  }
};

const generateListReportPDF = async (reports: any[], user: any) => {
  if (!reports || reports.length === 0) {
    alert('Không có dữ liệu để xuất báo cáo.');
    return;
  }
  
  const pdf = await getConfiguredJsPDF('l'); // Landscape for list
  
  // Header
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'bold');
  pdf.setFontSize(24);
  pdf.setTextColor(30, 58, 138); // blue-900
  pdf.text('TEV', 14, 20);
  
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'normal');
  pdf.setFontSize(10);
  pdf.setTextColor(29, 78, 216); // blue-700
  pdf.text('ASSET MANAGEMENT', 14, 27);
  
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'bold');
  pdf.setFontSize(18);
  pdf.setTextColor(15, 23, 42); // slate-900
  pdf.text('BÁO CÁO DANH SÁCH THIẾT BỊ', 280, 20, { align: 'right' });
  
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'normal');
  pdf.setFontSize(12);
  pdf.setTextColor(100, 116, 139); // slate-500
  pdf.text('ASSET INVENTORY REPORT', 280, 27, { align: 'right' });
  
  // Line separator
  pdf.setDrawColor(30, 58, 138);
  pdf.setLineWidth(0.5);
  pdf.line(14, 32, 280, 32);
  
  // Info
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'bold');
  pdf.setFontSize(10);
  pdf.setTextColor(100, 116, 139); // slate-500
  pdf.text('Người xuất / Exported By:', 14, 42);
  pdf.text('Ngày xuất / Export Date:', 200, 42);
  
  pdf.setFont(robotoRegularBase64 ? 'Roboto' : 'helvetica', 'normal');
  pdf.setTextColor(15, 23, 42); // slate-900
  pdf.text(user?.displayName || 'System', 60, 42);
  pdf.text(new Date().toLocaleDateString('vi-VN'), 245, 42);
  
  // Group reports by factory
  const groupedReports = reports.reduce((acc, report) => {
    if (!acc[report.factory]) acc[report.factory] = [];
    acc[report.factory].push(report);
    return acc;
  }, {} as Record<string, any[]>);
  
  const tableBody: any[] = [];
  
  Object.entries(groupedReports).forEach(([factory, factoryReports]) => {
    // Add grouping row
    tableBody.push([
      {
        content: `Location Name: ${factory}`,
        colSpan: 6,
        styles: { fontStyle: 'bold', fillColor: [255, 255, 255], textColor: [15, 23, 42], lineWidth: 0 }
      }
    ]);
    
    // Add data rows
    (factoryReports as any[]).forEach(r => {
      let statusText = 'Normal';
      if (r.status === 'warning') statusText = 'Caution';
      else if (r.status === 'critical' || r.status === 'danger') statusText = 'Severe';
      
      tableBody.push([
        String(r.equipmentId || ''),
        String(r.equipmentName || ''),
        String(r.type || ''),
        statusText,
        String(r.date || ''),
        String(r.inspector || '')
      ]);
    });
  });
  
  autoTable(pdf, {
    startY: 50,
    head: [['Asset ID / Mã TB', 'Equipment Name / Tên TB', 'Type / Loại', 'Condition / Đánh giá', 'Last Inspected / Ngày KT', 'Inspector / Người KT']],
    body: tableBody,
    theme: 'plain',
    styles: { font: robotoRegularBase64 ? 'Roboto' : 'helvetica', fontSize: 10, cellPadding: 3 },
    headStyles: { 
      fontStyle: 'bold', 
      textColor: [15, 23, 42], 
      lineWidth: { top: 1, bottom: 1 }, 
      lineColor: [15, 23, 42] 
    },
    bodyStyles: { 
      lineWidth: { bottom: 0.1 }, 
      lineColor: [203, 213, 225] 
    },
    columnStyles: {
      0: { cellWidth: 35 },
      1: { cellWidth: 65 },
      2: { cellWidth: 40 },
      3: { cellWidth: 35 },
      4: { cellWidth: 40 },
      5: { cellWidth: 50 }
    },
    didParseCell: function(data) {
      // Style grouping rows
      if (data.section === 'body' && data.row.raw[0] && data.row.raw[0].colSpan === 6) {
        data.cell.styles.lineWidth = { bottom: 1 };
        data.cell.styles.lineColor = [15, 23, 42];
      }
      
      // Style Condition column
      if (data.section === 'body' && data.column.index === 3 && !data.row.raw[0]?.colSpan) {
        data.cell.styles.fontStyle = 'bold';
        data.cell.styles.textColor = [255, 255, 255];
        data.cell.styles.halign = 'center';
        
        if (data.cell.raw === 'Normal') {
          data.cell.styles.fillColor = [0, 255, 0]; // Bright green like image
          data.cell.styles.textColor = [15, 23, 42]; // Dark text
        } else if (data.cell.raw === 'Caution') {
          data.cell.styles.fillColor = [255, 255, 0]; // Bright yellow like image
          data.cell.styles.textColor = [15, 23, 42]; // Dark text
        } else if (data.cell.raw === 'Severe') {
          data.cell.styles.fillColor = [255, 0, 0]; // Bright red like image
        } else if (data.cell.raw === 'Observe') {
          data.cell.styles.fillColor = [0, 112, 192]; // Blue like image
        }
      }
    }
  });
  
  pdf.save(`Asset_Inventory_Report_${new Date().toISOString().split('T')[0]}.pdf`);
};

// --- MOCK DATA ---
// (Removed to use real data from Google Sheets)

const detailedRiskData: Record<string, any> = {};

const statusDistribution: any[] = [];

const healthDistribution: any[] = [];


const createCustomIcon = (status: string, count: number) => {
  let bgColor = 'bg-emerald-500';
  let shadowColor = 'shadow-emerald-500/50';
  let ringColor = 'ring-emerald-500/30';
  
  if (status === 'critical') { 
    bgColor = 'bg-rose-500'; 
    shadowColor = 'shadow-rose-500/50'; 
    ringColor = 'ring-rose-500/30';
  } else if (status === 'warning') { 
    bgColor = 'bg-amber-500'; 
    shadowColor = 'shadow-amber-500/50'; 
    ringColor = 'ring-amber-500/30';
  }

  return L.divIcon({
    html: `<div class="${bgColor} text-white rounded-full w-10 h-10 flex items-center justify-center font-bold shadow-lg ${shadowColor} border-2 border-white ring-4 ${ringColor} animate-pulse-slow">${count}</div>`,
    className: 'custom-leaflet-icon',
    iconSize: [40, 40],
    iconAnchor: [20, 20],
  });
};

const MapBoundsHandler = ({ 
  sites, 
  viewTrigger 
}: { 
  sites: any[]; 
  viewTrigger: { type: 'auto' | 'fit-all' | 'global' | 'sea' | 'vn'; timestamp: number } 
}) => {
  const map = useMap();

  useEffect(() => {
    if (!map) return;
    if (viewTrigger.type === 'global') {
      map.setView([20, 15], 2);
    } else if (viewTrigger.type === 'sea') {
      map.setView([13.0, 105.0], 5);
    } else if (viewTrigger.type === 'vn') {
      map.setView([16.047079, 108.206230], 5.5);
    } else if (viewTrigger.type === 'fit-all' || viewTrigger.type === 'auto') {
      const validPoints = (sites || []).filter(s => typeof s.lat === 'number' && typeof s.lng === 'number' && !isNaN(s.lat) && !isNaN(s.lng));
      if (validPoints.length === 1) {
        map.setView([validPoints[0].lat, validPoints[0].lng], 8);
      } else if (validPoints.length > 1) {
        const bounds = L.latLngBounds(validPoints.map(s => [s.lat, s.lng]));
        map.fitBounds(bounds, { padding: [50, 50], maxZoom: 12 });
      } else {
        map.setView([16.047079, 108.206230], 5);
      }
    }
  }, [sites, viewTrigger, map]);

  return null;
};

const initialAllEquipment: any[] = getInitialEquipmentList();
const initialAllReports: any[] = getInitialReportsList();

// --- COMPONENTS ---

const StatusBadge = ({ status }: { status: string }) => {
  switch (status) {
    case 'healthy':
      return <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-emerald-100 text-emerald-800 border border-emerald-200">Đạt</span>;
    case 'warning':
      return <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-amber-100 text-amber-800 border border-amber-200">Cảnh báo</span>;
    case 'critical':
      return <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-rose-100 text-rose-800 border border-rose-200">Nguy hiểm</span>;
    default:
      return null;
  }
};

const TestStatusBadge = ({ status }: { status: string }) => {
  switch (status) {
    case 'Excellent':
      return <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-emerald-100 text-emerald-800">Tốt (Excellent)</span>;
    case 'Good':
      return <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-blue-100 text-blue-800">Khá (Good)</span>;
    case 'Moderate':
      return <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-amber-100 text-amber-800">Trung bình (Moderate)</span>;
    case 'Critical':
      return <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-rose-100 text-rose-800">Nguy hiểm (Critical)</span>;
    default:
      return null;
  }
};

const HealthBar = ({ value, status }: { value: number, status?: string }) => {
  let color = 'bg-emerald-500';
  if (status) {
    if (status === 'critical') color = 'bg-rose-500';
    else if (status === 'warning') color = 'bg-amber-500';
  } else {
    if (value < 60) color = 'bg-rose-500';
    else if (value < 80) color = 'bg-amber-500';
  }

  return (
    <div className="flex items-center gap-2">
      <div className="w-full bg-slate-200 rounded-full h-2.5">
        <div className={`${color} h-2.5 rounded-full`} style={{ width: `${value}%` }}></div>
      </div>
      <span className="text-xs font-medium text-slate-600 w-8">{value}%</span>
    </div>
  );
};

const getPoint = (top: number, right: number, left: number) => {
  const sum = top + right + left;
  if (sum === 0) return '50,5';
  const pz = top / sum;
  const py = right / sum;
  const px = left / sum;
  const x = 6.7 * px + 93.3 * py + 50 * pz;
  const y = 80 * px + 80 * py + 5 * pz;
  return `${x},${y}`;
};

const getPolygon = (points: number[][]) => {
  return points.map(p => getPoint(p[0], p[1], p[2])).join(' ');
};

const getLabelPos = (points: number[][]) => {
  let sumTop = 0, sumRight = 0, sumLeft = 0;
  points.forEach(p => {
    sumTop += p[0];
    sumRight += p[1];
    sumLeft += p[2];
  });
  const n = points.length;
  const pt = getPoint(sumTop/n, sumRight/n, sumLeft/n);
  return pt.split(',').map(Number);
};

export const regionsT1: { id: string, label?: string, color: string, points: number[][] }[] = [
  { id: 'PD', color: '#bbf7d0', points: [[100,0,0], [98,2,0], [98,0,2]] },
  { id: 'T1', color: '#fde047', points: [[98,2,0], [80,20,0], [76,20,4], [96,0,4], [98,0,2]] },
  { id: 'T2', color: '#f97316', points: [[80,20,0], [50,50,0], [46,50,4], [76,20,4]] },
  { id: 'T3', color: '#e11d48', points: [[50,50,0], [0,100,0], [0,85,15], [35,50,15]] },
  { id: 'DT', color: '#a5f3fc', points: [[96,0,4], [76,20,4], [46,50,4], [37,50,13], [64,23,13], [87,0,13]] },
  { id: 'D1', color: '#60a5fa', points: [[87,0,13], [64,23,13], [0,23,77], [0,0,100]] },
  { id: 'D2', color: '#e879f9', points: [[64,23,13], [37,50,13], [35,50,15], [0,85,15], [0,23,77]] }
];

const regionsT4: { id: string, label?: string, color: string, points: number[][] }[] = [
  { id: 'PD', color: '#bbf7d0', points: [[100,0,0], [98,2,0], [98,0,2]] },
  { id: 'O', color: '#60a5fa', points: [[9,91,0], [9,0,91], [0,0,100], [0,100,0]] },
  { id: 'C', color: '#e879f9', points: [[9,91,0], [9,15,76], [14,15,71], [14,40,46], [60,40,0]] },
  { id: 'S', color: '#a5f3fc', points: [[98,2,0], [60,40,0], [14,40,46], [14,15,71], [9,15,76], [9,0,91], [98,0,2]] }
];

const regionsT5: { id: string, label?: string, color: string, points: number[][] }[] = [
  { id: 'PD', color: '#bbf7d0', points: [[100,0,0], [98,2,0], [98,0,2]] },
  { id: 'O', color: '#60a5fa', points: [[98,2,0], [96,4,0], [82,4,14], [86,0,14], [98,0,2]] },
  { id: 'O2', label: 'O', color: '#60a5fa', points: [[10,0,90], [10,15,75], [0,15,85], [0,0,100]] },
  { id: 'T2', color: '#f97316', points: [[96,4,0], [50,50,0], [36,50,14], [82,4,14]] },
  { id: 'T3', color: '#e879f9', points: [[50,50,0], [0,100,0], [0,50,50], [36,50,14]] },
  { id: 'C', color: '#fde047', points: [[71,15,14], [36,50,14], [0,50,50], [0,15,85]] },
  { id: 'S', color: '#a5f3fc', points: [[86,0,14], [82,4,14], [71,15,14], [10,15,75], [10,0,90]] }
];

export const regionsP1 = [
  { name: 'PD', color: '#bbf7d0', polygon: [[0, 33], [-1, 33], [-1, 24.5], [0, 24.5]] },
  { name: 'D1', color: '#60a5fa', polygon: [[0, 40], [38, 12], [32, -6.1], [4, 16], [0, 1.5]] },
  { name: 'D2', color: '#e879f9', polygon: [[4, 16], [32, -6.1], [24.3, -30], [0, -3], [0, 1.5]] },
  { name: 'T3', color: '#e11d48', polygon: [[0, -3], [24.3, -30], [23.5, -32.4], [1, -32], [-6, -4]] },
  { name: 'T2', color: '#f97316', polygon: [[-6, -4], [1, -32.4], [-22.5, -32.4]] },
  { name: 'T1', color: '#fde047', polygon: [[-6, -4], [-22.5, -32.4], [-23.5, -32.4], [-35, 3], [0, 1.5], [0, -3]] },
  { name: 'S', color: '#a5f3fc', polygon: [[0, 1.5], [-35, 3.1], [-38, 12.4], [0, 40], [0, 33], [-1, 33], [-1, 24.5], [0, 24.5]] }
];

export const regionsP2 = [
  { name: 'PD', color: '#bbf7d0', polygon: [[0, 33], [-1, 33], [-1, 24.5], [0, 24.5]] },
  { name: 'D1', color: '#60a5fa', polygon: [[0, 40], [38, 12], [32, -6.1], [4, 16], [0, 1.5]] },
  { name: 'D2', color: '#e879f9', polygon: [[4, 16], [32, -6.1], [24.3, -30], [0, -3], [0, 1.5]] },
  { name: 'S', color: '#a5f3fc', polygon: [[0, 1.5], [-35, 3.1], [-38, 12.4], [0, 40], [0, 33], [-1, 33], [-1, 24.5], [0, 24.5]] },
  { name: 'T3', color: '#e11d48', polygon: [[0, -3], [24.3, -30], [23.5, -32.4], [2.5, -32.4], [-3.5, -3]] },
  { name: 'C', color: '#fde047', polygon: [[-3.5, -3], [2.5, -32.4], [-21.5, -32.4], [-11, -8]] },
  { name: 'O', color: '#93c5fd', polygon: [[-3.5, -3], [-11, -8], [-21.5, -32.4], [-23.5, -32.4], [-35, 3.1], [0, 1.5], [0, -3]] }
];

const DuvalPentagon = ({ title, labels, data, type }: { title: string, labels: string[], data: number[], type: 1 | 2 }) => {
  const sum = data[0] + data[1] + data[2] + data[3] + data[4];
  const p1 = sum > 0 ? data[0] / sum : 0; // H2
  const p2 = sum > 0 ? data[1] / sum : 0; // C2H6
  const p3 = sum > 0 ? data[2] / sum : 0; // CH4
  const p4 = sum > 0 ? data[3] / sum : 0; // C2H4
  const p5 = sum > 0 ? data[4] / sum : 0; // C2H2

  const mathX = p1 * 0 + p2 * (-38) + p3 * (-23.5) + p4 * 23.5 + p5 * 38;
  const mathY = p1 * 40 + p2 * 12.4 + p3 * (-32.4) + p4 * (-32.4) + p5 * 12.4;

  const x = 50 - mathX;
  const y = 50 - mathY;

  const regions = type === 1 ? regionsP1 : regionsP2;

  const legends = {
    1: [
      { id: 'PD', color: '#bbf7d0', text: 'Phóng điện cục bộ (Partial Discharge)' },
      { id: 'D1', color: '#60a5fa', text: 'Phóng điện năng lượng thấp (Low energy discharge)' },
      { id: 'D2', color: '#e879f9', text: 'Phóng điện năng lượng cao (High energy discharge)' },
      { id: 'T1', color: '#fde047', text: 'Quá nhiệt < 300°C (Thermal fault < 300°C)' },
      { id: 'T2', color: '#f97316', text: 'Quá nhiệt 300°C - 700°C (Thermal fault 300-700°C)' },
      { id: 'T3', color: '#e11d48', text: 'Quá nhiệt > 700°C (Thermal fault > 700°C)' },
      { id: 'S', color: '#a5f3fc', text: 'Lỗi có thể xảy ra trong dầu (Stray gassing of oil)' }
    ],
    2: [
      { id: 'PD', color: '#bbf7d0', text: 'Phóng điện cục bộ (Partial Discharge)' },
      { id: 'D1', color: '#60a5fa', text: 'Phóng điện năng lượng thấp (Low energy discharge)' },
      { id: 'D2', color: '#e879f9', text: 'Phóng điện năng lượng cao (High energy discharge)' },
      { id: 'S', color: '#a5f3fc', text: 'Lỗi có thể xảy ra trong dầu (Stray gassing of oil)' },
      { id: 'T3', color: '#e11d48', text: 'Quá nhiệt > 700°C (Thermal fault > 700°C)' },
      { id: 'C', color: '#fde047', text: 'Lỗi nhiệt có thể liên quan đến giấy (Thermal fault with paper involvement)' },
      { id: 'O', color: '#93c5fd', text: 'Quá nhiệt < 250°C (Overheating < 250°C)' }
    ]
  };

  const getPolygon = (points: number[][]) => {
    return points.map(p => `${50 - p[0]},${50 - p[1]}`).join(' ');
  };

  const getLabelPos = (points: number[][]) => {
    let cx = 0, cy = 0;
    points.forEach(p => { cx += p[0]; cy += p[1]; });
    return [50 - (cx / points.length), 50 - (cy / points.length)];
  };

  return (
    <div className="flex flex-col md:flex-row items-center gap-8 w-full">
      <div className="flex flex-col items-center flex-shrink-0">
        <h4 className="text-sm font-bold mb-2 text-slate-700">{title}</h4>
        <svg viewBox="0 0 100 100" className="w-full max-w-[240px] aspect-square overflow-visible">
          {regions.map((region, i) => {
            const [lx, ly] = getLabelPos(region.polygon);
            return (
              <g key={region.name + i}>
                <polygon points={getPolygon(region.polygon)} fill={region.color} stroke="#475569" strokeWidth="0.2" />
                {region.name !== 'PD' && (
                  <text x={lx} y={ly} fontSize="3" fontWeight="bold" textAnchor="middle" dominantBaseline="middle" fill="#1e293b">
                    {region.name}
                  </text>
                )}
              </g>
            );
          })}
          <polygon points="50,10 88,37.6 73.5,82.4 26.5,82.4 12,37.6" fill="none" stroke="#1e293b" strokeWidth="0.5" />
          <text x="50" y="8" fontSize="4" fontWeight="bold" textAnchor="middle" fill="#1e293b">{labels[0]}</text>
          <text x="90" y="37.6" fontSize="4" fontWeight="bold" textAnchor="start" fill="#1e293b">{labels[1]}</text>
          <text x="75" y="86" fontSize="4" fontWeight="bold" textAnchor="start" fill="#1e293b">{labels[2]}</text>
          <text x="25" y="86" fontSize="4" fontWeight="bold" textAnchor="end" fill="#1e293b">{labels[3]}</text>
          <text x="10" y="37.6" fontSize="4" fontWeight="bold" textAnchor="end" fill="#1e293b">{labels[4]}</text>
          {sum > 0 && (
            <g>
              <circle cx={x} cy={y} r="1.5" fill="#ef4444" stroke="#ffffff" strokeWidth="0.5" />
            </g>
          )}
        </svg>
      </div>
      <div className="flex-1 w-full">
        <h5 className="font-bold text-slate-800 mb-3 text-sm border-b border-slate-200 pb-2">Chú thích {title}</h5>
        <ul className="text-sm space-y-2">
          {legends[type].map((item) => (
            <li key={item.id} className="flex items-start gap-2">
              <span className="inline-block w-4 h-4 rounded-sm flex-shrink-0 mt-0.5 border border-slate-300" style={{ backgroundColor: item.color }}></span>
              <span className="font-medium text-slate-700 min-w-[30px]">{item.id}:</span>
              <span className="text-slate-600">{item.text}</span>
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
};

const DuvalTriangle = ({ title, labels, data, type }: { title: string, labels: string[], data: number[], type: 1 | 4 | 5 }) => {
  const sum = data[0] + data[1] + data[2];
  const px = sum > 0 ? data[2] / sum * 100 : 0; // bottom-left
  const py = sum > 0 ? data[1] / sum * 100 : 0; // bottom-right
  const pz = sum > 0 ? data[0] / sum * 100 : 0; // top

  const x = 6.7 * (px / 100) + 93.3 * (py / 100) + 50 * (pz / 100);
  const y = 80 * (px / 100) + 80 * (py / 100) + 5 * (pz / 100);

  const regions = type === 1 ? regionsT1 : type === 4 ? regionsT4 : regionsT5;

  const legends = {
    1: [
      { id: 'PD', color: '#bbf7d0', text: 'Phóng điện cục bộ (Partial Discharge)' },
      { id: 'D1', color: '#60a5fa', text: 'Phóng điện năng lượng thấp (Low energy discharge)' },
      { id: 'D2', color: '#e879f9', text: 'Phóng điện năng lượng cao (High energy discharge)' },
      { id: 'T1', color: '#fde047', text: 'Quá nhiệt < 300°C (Thermal fault < 300°C)' },
      { id: 'T2', color: '#f97316', text: 'Quá nhiệt 300°C - 700°C (Thermal fault 300-700°C)' },
      { id: 'T3', color: '#e11d48', text: 'Quá nhiệt > 700°C (Thermal fault > 700°C)' },
      { id: 'DT', color: '#a5f3fc', text: 'Hỗn hợp nhiệt và điện (Thermal and electrical faults)' }
    ],
    4: [
      { id: 'PD', color: '#bbf7d0', text: 'Phóng điện cục bộ (Partial Discharge)' },
      { id: 'S', color: '#a5f3fc', text: 'Lỗi có thể xảy ra trong dầu (Stray gassing of oil)' },
      { id: 'O', color: '#60a5fa', text: 'Quá nhiệt < 250°C (Overheating < 250°C)' },
      { id: 'C', color: '#e879f9', text: 'Lỗi nhiệt có thể liên quan đến giấy (Thermal fault with paper involvement)' },
      { id: 'ND', color: '#cbd5e1', text: 'Không xác định (Not Determined)' }
    ],
    5: [
      { id: 'PD', color: '#bbf7d0', text: 'Phóng điện cục bộ (Partial Discharge)' },
      { id: 'T1', color: '#fde047', text: 'Quá nhiệt < 300°C (Thermal fault < 300°C)' },
      { id: 'T2', color: '#f97316', text: 'Quá nhiệt 300°C - 700°C (Thermal fault 300-700°C)' },
      { id: 'T3', color: '#e879f9', text: 'Quá nhiệt > 700°C (Thermal fault > 700°C)' },
      { id: 'C', color: '#fde047', text: 'Lỗi nhiệt có thể liên quan đến giấy (Thermal fault with paper involvement)' },
      { id: 'O', color: '#60a5fa', text: 'Quá nhiệt < 250°C (Overheating < 250°C)' },
      { id: 'S', color: '#a5f3fc', text: 'Lỗi có thể xảy ra trong dầu (Stray gassing of oil)' },
      { id: 'ND', color: '#cbd5e1', text: 'Không xác định (Not Determined)' }
    ]
  };

  return (
    <div className="flex flex-col md:flex-row items-center gap-8 w-full">
      <div className="flex flex-col items-center flex-shrink-0">
        <h4 className="text-sm font-bold mb-2 text-slate-700">{title}</h4>
        <svg viewBox="0 0 100 100" className="w-full max-w-[240px] aspect-square overflow-visible">
          {regions.map((region, i) => {
            const [lx, ly] = getLabelPos(region.points);
            return (
              <g key={region.id + i}>
                <polygon points={getPolygon(region.points)} fill={region.color} stroke="#475569" strokeWidth="0.2" />
                {region.id !== 'PD' && (
                  <text x={lx} y={ly} fontSize="3" fontWeight="bold" textAnchor="middle" dominantBaseline="middle" fill="#1e293b">
                    {region.label || region.id}
                  </text>
                )}
              </g>
            );
          })}
          <polygon points="50,5 93.3,80 6.7,80" fill="none" stroke="#1e293b" strokeWidth="0.5" />
          <text x="50" y="2" fontSize="4" fontWeight="bold" textAnchor="middle" fill="#1e293b">{labels[0]}</text>
          <text x="96" y="84" fontSize="4" fontWeight="bold" textAnchor="start" fill="#1e293b">{labels[1]}</text>
          <text x="4" y="84" fontSize="4" fontWeight="bold" textAnchor="end" fill="#1e293b">{labels[2]}</text>
          {sum > 0 && (
            <g>
              <circle cx={x} cy={y} r="1.5" fill="#ef4444" stroke="#ffffff" strokeWidth="0.5" />
            </g>
          )}
        </svg>
      </div>
      <div className="flex-1 w-full">
        <h5 className="font-bold text-slate-800 mb-3 text-sm border-b border-slate-200 pb-2">Chú thích {title}</h5>
        <ul className="text-sm space-y-2">
          {legends[type].map((item) => (
            <li key={item.id} className="flex items-start gap-2">
              <span className="inline-block w-4 h-4 rounded-sm flex-shrink-0 mt-0.5 border border-slate-300" style={{ backgroundColor: item.color }}></span>
              <span className="font-medium text-slate-700 min-w-[30px]">{item.id}:</span>
              <span className="text-slate-600">{item.text}</span>
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
};

const DgaMatrix = ({ matrix }: { matrix: Record<string, string> }) => {
  const methods = ['IEEE (Dornenburg)', 'IEEE (Rogers)', 'IEC Ratio', 'Duval T1', 'Duval T4', 'Duval T5', 'Duval P1', 'Duval P2', 'KeyGas', 'ETRA', 'CO2/CO'];
  const statuses = ['ND', 'OK', 'PD', 'S', 'T1', 'O', 'C', 'T2', 'T3', 'DT', 'D2', 'D1'];
  
  const statusColors: Record<string, string> = {
    'ND': 'bg-slate-400', 'OK': 'bg-emerald-500', 'PD': 'bg-blue-600', 'S': 'bg-sky-400',
    'T1': 'bg-orange-200', 'O': 'bg-orange-300', 'C': 'bg-orange-400', 'T2': 'bg-orange-600',
    'T3': 'bg-orange-800', 'DT': 'bg-amber-500', 'D2': 'bg-red-600', 'D1': 'bg-red-800',
    'PD/Corona': 'bg-blue-600', 'Overheating (Oil)': 'bg-orange-600', 'Overheating (Paper)': 'bg-orange-800',
    'Arcing': 'bg-red-600', 'Low Temp Thermal': 'bg-orange-200', 'Normal': 'bg-emerald-500',
    'Condition 2': 'bg-yellow-400', 'Condition 3': 'bg-orange-500', 'Condition 4': 'bg-red-600'
  };

  return (
    <div className="overflow-x-auto">
      <table className="w-full border-collapse text-xs text-center">
        <thead>
          <tr>
            <th className="border border-slate-300 p-2 bg-slate-100 text-left">Methods</th>
            {statuses.map(s => (
              <th key={s} className={`border border-slate-300 p-1 text-white ${statusColors[s] || 'bg-slate-500'}`}>
                {s}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {methods.map(m => (
            <tr key={m}>
              <td className="border border-slate-300 p-2 text-left font-medium">{m}</td>
              {statuses.map(s => (
                <td key={s} className={`border border-slate-300 p-0`}>
                  {matrix[m] === s && <div className={`w-full h-6 ${statusColors[s]}`}></div>}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

const evaluateEquipmentParam = (type: string, param: string, value: any) => {
  if (value === '' || value === undefined || value === null) return null;
  
  // Robust number parsing (handles units like "65 °C" or "500 ppm")
  const parseVal = (val: any) => {
    if (typeof val === 'number') return val;
    const s = val.toString().replace(/[^0-9.,-]/g, '').replace(',', '.');
    const n = parseFloat(s);
    return isNaN(n) ? null : n;
  };

  const numVal = parseVal(value);
  
  if (type === 'Máy biến áp') {
    if (numVal !== null) {
      if (param === 'oilTemp') return numVal <= 80 ? 'healthy' : numVal <= 90 ? 'warning' : 'critical';
      if (param === 'windingTemp') return numVal <= 90 ? 'healthy' : numVal <= 105 ? 'warning' : 'critical';
      if (param === 'irHighLow') return numVal >= 2000 ? 'healthy' : numVal >= 1000 ? 'warning' : 'critical';
      if (param === 'irHighEarth') return numVal >= 2000 ? 'healthy' : numVal >= 1000 ? 'warning' : 'critical';
      if (param === 'dga') return numVal <= 1000 ? 'healthy' : numVal <= 2500 ? 'warning' : 'critical';
      if (param === 'dielectricStrength') return numVal >= 50 ? 'healthy' : numVal >= 40 ? 'warning' : 'critical';
      if (param === 'furan') return numVal <= 1 ? 'healthy' : numVal <= 5 ? 'warning' : 'critical';
      if (param === 'oilMoisture') return numVal <= 15 ? 'healthy' : numVal <= 25 ? 'warning' : 'critical';
    }
    
    if (param === 'oilLeak') {
      const s = value?.toString().toLowerCase().trim() || '';
      if (s === 'none' || s === 'normal' || s === 'bình thường' || s === 'không' || s === 'tốt' || s === 'đạt' || s === '') return 'healthy';
      if (s === 'light' || s === 'nhẹ' || s === 'ít' || s === 'theo dõi') return 'warning';
      return 'critical';
    }
  }
  if (type === 'Động cơ') {
    if (param === 'statorTemp') return numVal <= 110 ? 'healthy' : numVal <= 130 ? 'warning' : 'critical';
    if (param === 'bearingTemp') return numVal <= 80 ? 'healthy' : numVal <= 95 ? 'warning' : 'critical';
    if (param === 'vibration') return numVal <= 2.3 ? 'healthy' : numVal <= 4.5 ? 'warning' : 'critical';
    if (param === 'tanDelta') return numVal < 0.04 ? 'healthy' : numVal <= 0.07 ? 'warning' : 'critical';
    if (param === 'tipUp') return numVal < 0.004 ? 'healthy' : numVal <= 0.006 ? 'warning' : 'critical';
    if (param === 'pd') return numVal < 10000 ? 'healthy' : numVal <= 15000 ? 'warning' : 'critical';
    if (param === 'ir') return numVal >= 10 ? 'healthy' : numVal >= 1 ? 'warning' : 'critical';
    if (param === 'pi') return numVal >= 1.5 ? 'healthy' : numVal >= 1.0 ? 'warning' : 'critical';
    if (param === 'dd') return numVal < 4 ? 'healthy' : numVal <= 8 ? 'warning' : 'critical';
    if (param === 'elcid') return numVal < 200 ? 'healthy' : numVal <= 300 ? 'warning' : 'critical';
    if (param === 'voltageImbalance') return numVal <= 3 ? 'healthy' : 'warning';
  }
  if (type === 'Tủ điện trung thế' || type === 'Tủ điện') {
    if (param === 'thermography') return numVal <= 60 ? 'healthy' : numVal <= 75 ? 'warning' : 'critical';
    if (param === 'contactRes') return numVal <= 30 ? 'healthy' : numVal <= 50 ? 'warning' : 'critical';
    if (param === 'tev') return numVal < 20 ? 'healthy' : numVal <= 30 ? 'warning' : 'critical';
    if (param === 'tevPulses') return numVal < 5 ? 'healthy' : numVal <= 20 ? 'warning' : 'critical';
    if (param === 'ultrasonic') return numVal <= 5 ? 'healthy' : numVal <= 10 ? 'warning' : 'critical';
    if (param === 'humidity') return numVal <= 60 ? 'healthy' : numVal <= 80 ? 'warning' : 'critical';
    if (param === 'sf6Pressure') return numVal >= 5.5 ? 'healthy' : numVal >= 5.0 ? 'warning' : 'critical';
  }
  if (type === 'Inverter') {
    if (numVal !== null) {
      if (param === 'acOutputVoltage') return (numVal >= 207 && numVal <= 244) || (numVal >= 360 && numVal <= 424) ? 'healthy' : 'warning';
      if (param === 'acFrequency') return (numVal >= 49.7 && numVal <= 50.3) ? 'healthy' : 'warning';
      if (param === 'harmonicDistortion') return numVal <= 1 ? 'healthy' : numVal <= 3 ? 'warning' : 'critical';
      if (param === 'maxEfficiency') return numVal >= 95 ? 'healthy' : numVal >= 90 ? 'warning' : 'critical';
      if (param === 'euroEfficiency') return numVal >= 94 ? 'healthy' : numVal >= 89 ? 'warning' : 'critical';
      if (param === 'operatingTemp') return numVal <= 40 ? 'healthy' : numVal <= 50 ? 'warning' : 'critical';
    }
    
    // Boolean/Status checks
    const s = value?.toString().toLowerCase().trim() || '';
    if (param === 'antiIslanding') return s === 'đạt' || s === 'pass' || s === 'ok' || s === 'yes' ? 'healthy' : 'critical';
    if (param === 'rcdIsolation') return s === 'đạt' || s === 'pass' || s === 'ok' || s === 'yes' ? 'healthy' : 'critical';
    if (param === 'physicalDisconnects') return s === 'đạt' || s === 'pass' || s === 'ok' || s === 'yes' ? 'healthy' : 'critical';
    if (param === 'airFilters') return s === 'sạch' || s === 'clean' || s === 'ok' ? 'healthy' : 'warning';
  }
  return null;
};

const getColumnIndex = (header: string[], possibleNames: string[]) => {
  if (!header) return -1;
  return header.findIndex(col => {
    if (col === undefined || col === null) return false;
    const colStr = col.toString().toLowerCase();
    return possibleNames.some(name => colStr.includes(name.toLowerCase()));
  });
};

const findHeaderRow = (rows: any[][]) => {
  if (!rows || rows.length === 0) return -1;
  // Look for a row that contains common keywords like "Mã thiết bị" or "ID"
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const row = rows[i];
    if (row && row.some(cell => {
      const s = cell?.toString().toLowerCase() || '';
      return s.includes('mã thiết bị') || s.includes('equipment id') || s.includes('id') || s.includes('tên thiết bị');
    })) {
      return i;
    }
  }
  return 0; // Default to first row
};

const mapStatusFromSheet = (status: any) => {
  const s = status?.toString().toLowerCase() || '';
  if (s.includes('bình thường') || s.includes('healthy') || s.includes('tốt') || s.includes('normal') || s.includes('đạt')) return 'healthy';
  if (s.includes('cảnh báo') || s.includes('warning') || s.includes('theo dõi')) return 'warning';
  if (s.includes('nguy hiểm') || s.includes('critical') || s.includes('xấu') || s.includes('đang hỏng') || s.includes('không đạt')) return 'critical';
  return 'healthy'; // Default
};

const calculateNextPMDate = (lastDateStr: string, frequency: string) => {
  if (!lastDateStr || !frequency) return null;
  const date = new Date(lastDateStr);
  if (isNaN(date.getTime())) return null;

  switch (frequency) {
    case 'daily': date.setDate(date.getDate() + 1); break;
    case 'weekly': date.setDate(date.getDate() + 7); break;
    case 'bi-monthly': date.setDate(date.getDate() + 15); break;
    case 'monthly': date.setMonth(date.getMonth() + 1); break;
    case '3-months': date.setMonth(date.getMonth() + 3); break;
    case '6-months': date.setMonth(date.getMonth() + 6); break;
    case '1-year': date.setFullYear(date.getFullYear() + 1); break;
    case '3-years': date.setFullYear(date.getFullYear() + 3); break;
    case '6-years': date.setFullYear(date.getFullYear() + 6); break;
    default: return null;
  }
  return date;
};

export default function App() {
  const [user, setUser] = useState<any>(null);
  const [userRole, setUserRole] = useState<string | null>(null);
  const [userCustomerName, setUserCustomerName] = useState<string | null>(null);
  const [userFactory, setUserFactory] = useState<string | null>(null);
  const [isAuthLoading, setIsAuthLoading] = useState(true);
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);

  const [activeTab, setActiveTab] = useState('dashboard');
  const [mapLayer, setMapLayer] = useState<'streets' | 'light' | 'satellite' | 'topo'>('streets');
  const [mapViewTrigger, setMapViewTrigger] = useState<{ type: 'auto' | 'fit-all' | 'global' | 'sea' | 'vn'; timestamp: number }>({ type: 'auto', timestamp: Date.now() });
  const [equipmentSearch, setEquipmentSearch] = useState('');
  const [showEqDropdown, setShowEqDropdown] = useState(false);
  const [workOrders, setWorkOrders] = useState<any[]>(() => DEFAULT_WORK_ORDERS);
  const [isWorkOrdersLoading, setIsWorkOrdersLoading] = useState(false);
  const [customers, setCustomers] = useState<any[]>(() => DEFAULT_CUSTOMERS);
  const [isCustomersLoading, setIsCustomersLoading] = useState(false);
  const [inventory, setInventory] = useState<any[]>(() => DEFAULT_INVENTORY);
  const [isInventoryLoading, setIsInventoryLoading] = useState(false);

  const [showHistoryModal, setShowHistoryModal] = useState(false);
  const [showWorkOrderModal, setShowWorkOrderModal] = useState(false);
  const [showInventoryModal, setShowInventoryModal] = useState(false);
  const [showCustomerModal, setShowCustomerModal] = useState(false);
  const [selectedWorkOrder, setSelectedWorkOrder] = useState<any>(null);
  const [selectedInventoryItem, setSelectedInventoryItem] = useState<any>(null);
  const [editingCustomer, setEditingCustomer] = useState<any>(null);
  const [newWorkOrder, setNewWorkOrder] = useState({
    title: '',
    description: '',
    equipmentId: [] as string[],
    customerId: '',
    factory: '',
    priority: 'medium',
    type: 'preventive',
    status: 'initiated',
    assignedTo: '',
    dueDate: '',
    usedMaterials: [] as any[],
    responsibleApprove: 'TM',
    responsibleDo: 'Everybody',
    blockingRequired: false,
    workPermitId: '',
    isUnplanned: false,
    failureCode: '',
    rootCause: '',
    actualTimeSpent: 0,
    pmFrequency: '',
    estimatedTime: 0,
    laborCount: 0,
    laborCost: 0,
    partCost: 0,
    attachments: [] as string[],
    downtimeStart: '',
    repairStart: '',
    repairEnd: '',
    restartTime: ''
  });
  const [newInventoryItem, setNewInventoryItem] = useState({
    name: '',
    sku: '',
    category: '',
    quantity: 0,
    unit: 'pcs',
    minStock: 5,
    location: '',
    price: 0
  });
  const [newCustomer, setNewCustomer] = useState({
    name: '',
    email: '',
    phone: '',
    address: '',
    factories: [] as string[]
  });
  
  const [woTypeFilter, setWoTypeFilter] = useState('');
  const [woPriorityFilter, setWoPriorityFilter] = useState('');
  const [woStatusFilter, setWoStatusFilter] = useState('');
  const [woAssigneeFilter, setWoAssigneeFilter] = useState('');
  const [woSearchQuery, setWoSearchQuery] = useState('');
  const [woError, setWoError] = useState('');
  const [confirmDialog, setConfirmDialog] = useState<{isOpen: boolean, message: string, onConfirm: () => void} | null>(null);
  
  // Dashboard CMMS States
  const [dashboardWoFilter, setDashboardWoFilter] = useState<'all' | 'corrective' | 'preventive' | 'in-progress' | 'initiated' | 'overdue'>('all');
  const [dashboardWoSearch, setDashboardWoSearch] = useState('');
  
  const [selectedParamHistory, setSelectedParamHistory] = useState('');
  const [showEquipmentProfile, setShowEquipmentProfile] = useState(false);
  const [selectedEqType, setSelectedEqType] = useState('Máy biến áp');
  const [equipmentCode, setEquipmentCode] = useState('TRF-01');
  const [equipmentName, setEquipmentName] = useState('Máy biến áp T1');
  const [siteName, setSiteName] = useState('Nhà máy Bắc Ninh');
  const [customerName, setCustomerName] = useState('Công ty Điện lực A');
  const [locationName, setLocationName] = useState('Trạm biến áp 110kV');
  const [wpNumber, setWpNumber] = useState('');
  const [showEqSuggestions, setShowEqSuggestions] = useState(false);
  const [showQRScanner, setShowQRScanner] = useState(false);
  const [isNewEquipment, setIsNewEquipment] = useState(false);
  const [formData, setFormData] = useState<Record<string, any>>({});
  const [healthResult, setHealthResult] = useState<{index: number, status: string} | null>(null);
  const [selectedRiskDetail, setSelectedRiskDetail] = useState<string | null>(null);
  const [selectedEqForQR, setSelectedEqForQR] = useState<any>(null);
  const [showQRModal, setShowQRModal] = useState(false);
  const [allEquipment, setAllEquipment] = useState(initialAllEquipment);

  // Tự động đồng bộ các thiết bị được tạo từ phiếu CMMS vào allEquipment nếu chưa có (ví dụ: MVCB-001)
  useEffect(() => {
    if (workOrders.length > 0) {
      setAllEquipment(prevEq => {
        let hasNew = false;
        const next = [...prevEq];
        workOrders.forEach(wo => {
          const rawEqIds = Array.isArray(wo.equipmentId) ? wo.equipmentId : (wo.equipmentId ? [wo.equipmentId] : []);
          rawEqIds.forEach((id: string) => {
            const cleanId = String(id).trim();
            if (cleanId && !next.some(e => e.id?.toLowerCase() === cleanId.toLowerCase())) {
              const cust = wo.customer || wo.customerId || 'Super Energy';
              next.push({
                id: cleanId,
                name: wo.equipmentName || (cleanId.startsWith('MVCB') ? `Tủ máy cắt trung thế ${cleanId} 24kV` : `Thiết bị ${cleanId}`),
                customer: cust,
                factory: `Nhà máy Điện Mặt Trời ${cust}`,
                location: 'Trạm phân phối 24kV - Ngăn lộ F01',
                type: cleanId.startsWith('MVCB') || cleanId.startsWith('MC') || cleanId.startsWith('SWG') ? 'Tủ điện' : 'Máy biến áp',
                status: wo.priority === 'Critical' || wo.priority === 'High' ? 'warning' : 'healthy',
                health: wo.priority === 'Critical' ? 65 : 72,
                lastCheck: new Date().toLocaleDateString('vi-VN'),
                criticality: 'A',
                notes: `Thiết bị liên kết từ Phiếu CMMS ${wo.id} (${wo.title || ''})`
              });
              hasNew = true;
            }
          });
        });
        return hasNew ? next : prevEq;
      });
    }
  }, [workOrders]);
  const [showPMNotificationModal, setShowPMNotificationModal] = useState(false);
  const [pmNotificationLoading, setPmNotificationLoading] = useState(false);
  const [pmNotificationResult, setPmNotificationResult] = useState<any>(null);
  const [pmUpcomingTasks, setPmUpcomingTasks] = useState<any[]>([]);
  const [testEmailRecipient, setTestEmailRecipient] = useState('sgm1707@gmail.com');
  const [pmNotificationConfig, setPmNotificationConfig] = useState<any>(null);
  const [pmSubTab, setPmSubTab] = useState<'dashboard' | 'schedule' | 'alerts'>('dashboard');
  const [cmmsSubTab, setCmmsSubTab] = useState<'station' | 'table' | 'pm'>('station');

  // Sync Reconciliation State (Kiểm tra phiên bản & Đối soát Firestore ↔ Google Sheets)
  const [isReconciling, setIsReconciling] = useState<boolean>(false);
  const [showReconciliationModal, setShowReconciliationModal] = useState<boolean>(false);
  const [reconciliationReport, setReconciliationReport] = useState<ReconciliationReport | null>(null);

  // 2-Way Automated Sync States (Luồng dữ liệu 2 chiều tự động App ↔ Google Sheets 'TEV Service Flatform')
  const [isAutoSyncEnabled, setIsAutoSyncEnabled] = useState<boolean>(true);
  const [autoSyncStatus, setAutoSyncStatus] = useState<'idle' | 'syncing' | 'synced' | 'error'>('idle');
  const [lastAutoSyncTime, setLastAutoSyncTime] = useState<string>('');
  const [twoWaySyncNotice, setTwoWaySyncNotice] = useState<{ type: 'success' | 'error' | 'info'; message: string } | null>(null);
  const [googleSheetUrl, setGoogleSheetUrl] = useState<string>(() => {
    const id = localStorage.getItem('tev_spreadsheet_id');
    return id ? `https://docs.google.com/spreadsheets/d/${id}/edit` : '';
  });

  const handleQuickUpdateWorkOrderStatus = async (woId: string, newStatus: string) => {
    const now = new Date().toISOString();
    setWorkOrders(prev => prev.map(wo => wo.id === woId ? { ...wo, status: newStatus, completedAt: newStatus === 'completed' ? now : wo.completedAt } : wo));
    try {
      const docRef = doc(db, 'workOrders', woId);
      await updateDoc(docRef, {
        status: newStatus,
        updatedAt: now,
        completedAt: newStatus === 'completed' ? now : null
      });
    } catch (e) {
      console.warn('Updated local work order status:', woId, newStatus);
    }
  };

  const checkUpcomingPMAlerts = async () => {
    try {
      setPmNotificationLoading(true);
      const res = await fetch('/api/notifications/pm-upcoming', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ workOrders, allEquipment, targetDays: 3 })
      });
      const data = await res.json();
      if (data.success) {
        setPmUpcomingTasks(data.tasks || []);
      }
    } catch (err) {
      console.error('Failed to check upcoming PM tasks:', err);
    } finally {
      setPmNotificationLoading(false);
    }
  };

  const fetchPMNotificationConfig = async () => {
    try {
      const res = await fetch('/api/notifications/pm-config');
      const data = await res.json();
      setPmNotificationConfig(data);
    } catch (err) {
      console.error('Failed to load PM notification config:', err);
    }
  };

  const handleSendPMAlerts = async (force: boolean = false) => {
    try {
      setPmNotificationLoading(true);
      const res = await fetch('/api/notifications/send-pm-alerts', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ workOrders, allEquipment, forceSend: force })
      });
      const data = await res.json();
      setPmNotificationResult(data);
      if (data.success) {
        alert(data.message || 'Đã gửi thông báo bảo trì PM thành công!');
        checkUpcomingPMAlerts();
      } else {
        alert('Lỗi: ' + (data.error || 'Không thể gửi thông báo.'));
      }
    } catch (err: any) {
      alert('Lỗi gửi thông báo: ' + err.message);
    } finally {
      setPmNotificationLoading(false);
    }
  };

  const handleSendTestEmail = async () => {
    try {
      setPmNotificationLoading(true);
      const res = await fetch('/api/notifications/test-email', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ recipient: testEmailRecipient })
      });
      const data = await res.json();
      setPmNotificationResult(data);
      if (data.success) {
        alert(data.message || 'Đã gửi email thử nghiệm!');
      } else {
        alert('Lỗi: ' + (data.error || 'Gửi thất bại.'));
      }
    } catch (err: any) {
      alert('Lỗi gửi email thử nghiệm: ' + err.message);
    } finally {
      setPmNotificationLoading(false);
    }
  };

  useEffect(() => {
    if (activeTab === 'pm-schedule') {
      checkUpcomingPMAlerts();
      fetchPMNotificationConfig();
    }
  }, [activeTab, workOrders.length]);

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      try {
        if (currentUser) {
          setUser(currentUser);
          // Check if user exists in Firestore
          const userDocRef = doc(db, 'users', currentUser.uid);
          const userDoc = await getDoc(userDocRef);
          
          if (userDoc.exists()) {
            const userData = userDoc.data();
            // Normalize role: handle both 'customer' and 'Khách hàng'
            const rawRole = userData.role?.toLowerCase() || 'customer';
            const normalizedRole = (rawRole === 'admin' || rawRole === 'quản trị viên' || rawRole === 'quan tri vien') ? 'admin' : 'customer';
            
            setUserRole(normalizedRole);
            if (userData.assignedFactory) {
              setUserFactory(userData.assignedFactory.trim());
            }
            if (normalizedRole === 'customer' && userData.customerId) {
              // Fetch customer name
              const customerDoc = await getDoc(doc(db, 'customers', userData.customerId));
              if (customerDoc.exists()) {
                setUserCustomerName(customerDoc.data().name);
              }
            }
            // Ensure customer database and CMMS work orders are loaded for all users (to populate dropdowns and dashboard)
            fetchCustomers();
            fetchWorkOrders();
          } else {
            // Create default user document
            // If email matches the default admin email, set as admin
            const role = currentUser.email === 'sgm1707@gmail.com' ? 'admin' : 'customer';
            await setDoc(userDocRef, {
              email: currentUser.email,
              displayName: currentUser.displayName,
              role: role,
              createdAt: new Date().toISOString()
            });
            setUserRole(role);
            fetchCustomers();
            fetchWorkOrders();
          }
        } else {
          setUser(null);
          setUserRole(null);
          setUserCustomerName(null);
          setUserFactory(null);
        }
      } catch (error) {
        console.error('Error in auth state change:', error);
        // Fallback state
        setUser(currentUser);
        setUserRole(currentUser?.email === 'sgm1707@gmail.com' ? 'admin' : 'customer');
        fetchCustomers();
        fetchWorkOrders();
      } finally {
        setIsAuthLoading(false);
      }
    });

    return () => unsubscribe();
  }, []);

  useEffect(() => {
    // Initial fetch to ensure Dashboard immediately displays latest CMMS work orders and customers
    fetchWorkOrders();
    fetchCustomers();
  }, []);

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const eqId = params.get('eqId');
    if (eqId && allEquipment.length > 0) {
      const foundEq = allEquipment.find(e => e.id.toLowerCase() === eqId.toLowerCase());
      if (foundEq) {
        setEquipmentCode(foundEq.id);
        setEquipmentName(foundEq.name);
        setCustomerName(foundEq.customer);
        setSiteName(foundEq.factory);
        setLocationName(foundEq.location);
        setSelectedEqType(foundEq.type === 'Tủ điện' ? 'Tủ điện trung thế' : foundEq.type);
        populateFormData(foundEq);
        setShowEquipmentProfile(true);
        // Clear the param from URL without refreshing
        const newUrl = window.location.pathname;
        window.history.replaceState({}, '', newUrl);
      }
    }
  }, [allEquipment]);

  const handleLogin = async () => {
    try {
      await signInWithGoogle();
    } catch (error) {
      console.error(error);
    }
  };

  const fetchWorkOrders = () => {
    if (!auth.currentUser) return;
    
    setIsWorkOrdersLoading(true);
    const q = query(
      collection(db, 'workOrders'),
      orderBy('createdAt', 'desc'),
      limit(50)
    );

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const orders = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      }));
      setWorkOrders(orders);
      setIsWorkOrdersLoading(false);
    }, (error) => {
      console.error("Error fetching work orders:", error);
      setIsWorkOrdersLoading(false);
    });

    return unsubscribe;
  };

  const handleViewEquipmentProfile = (eq: any) => {
    setEquipmentCode(eq.id);
    setEquipmentName(eq.name);
    setCustomerName(eq.customer);
    setSiteName(eq.factory);
    setLocationName(eq.location);
    setSelectedEqType(eq.type === 'Tủ điện' ? 'Tủ điện trung thế' : eq.type);
    populateFormData(eq);
    setShowEquipmentProfile(true);
  };

  const fetchCustomers = () => {
    if (!auth.currentUser) return;
    setIsCustomersLoading(true);
    const q = query(collection(db, 'customers'), orderBy('name', 'asc'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const items = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      setCustomers(items);
      setIsCustomersLoading(false);
    }, (error) => {
      console.error("Error fetching customers:", error);
      setIsCustomersLoading(false);
    });
    return unsubscribe;
  };

  const fetchInventory = () => {
    if (!auth.currentUser) return;
    setIsInventoryLoading(true);
    const q = query(collection(db, 'inventory'), orderBy('name', 'asc'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const items = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      setInventory(items);
      setIsInventoryLoading(false);
    }, (error) => {
      console.error("Error fetching inventory:", error);
      setIsInventoryLoading(false);
    });
    return unsubscribe;
  };

  const handleTabChange = (tab: string) => {
    setActiveTab(tab);
    setIsSidebarOpen(false); // Close sidebar on mobile when tab changes
    if (tab === 'dashboard' && isGoogleConnected) {
      handleFetchFromSheets();
    }
    if (tab === 'dashboard' || tab === 'cmms' || tab === 'customers') {
      fetchWorkOrders();
      fetchCustomers();
    }
    if (tab === 'inventory') {
      fetchInventory();
    }
  };

  const updateThreePhase = (param: string, phase: string, val: any) => {
    const current = formData[param] || { R: '', Y: '', B: '' };
    setFormData({
      ...formData,
      [param]: { ...current, [phase]: val }
    });
  };

  const populateFormData = (eq: any) => {
    if (!eq) return;
    
    let linksStr = '';
    
    if (eq.measurements) {
      // Use pre-parsed measurements if available (most robust)
      setFormData(eq.measurements);
      
      // Handle links separately as they are not in measurements
      if (eq.rawData) {
        if (eq.type === 'Máy biến áp') linksStr = eq.rawData[20] || '';
        else if (eq.type === 'Tủ điện trung thế' || eq.type === 'Tủ điện') linksStr = eq.rawData[18] || '';
        else if (eq.type === 'Động cơ') linksStr = eq.rawData[19] || '';
        else if (eq.type === 'Inverter') linksStr = eq.rawData[23] || '';
      }
    } else if (eq.rawData) {
      // Fallback to rawData (less robust, but kept for compatibility)
      if (eq.type === 'Máy biến áp') {
        setFormData({
          oilTemp: eq.rawData[9] || '',
          windingTemp: eq.rawData[10] || '',
          irHighLow: eq.rawData[11] || '',
          irHighEarth: eq.rawData[12] || '',
          oilLeak: eq.rawData[13] || '',
          dga: eq.rawData[14] || '',
          dielectricStrength: eq.rawData[15] || '',
          furan: eq.rawData[16] || '',
          oilMoisture: eq.rawData[17] || '',
          age: eq.rawData[18] || '',
          dutyFactor: eq.rawData[19] || ''
        });
        linksStr = eq.rawData[20] || '';
      } else if (eq.type === 'Tủ điện trung thế' || eq.type === 'Tủ điện') {
        setFormData({
          thermography: eq.rawData[9] || '',
          contactRes: eq.rawData[10] || '',
          tev: eq.rawData[11] || '',
          ultrasonic: eq.rawData[12] || '',
          tevPulses: eq.rawData[13] || '',
          humidity: eq.rawData[14] || '',
          sf6Pressure: eq.rawData[15] || '',
          age: eq.rawData[16] || '',
          dutyFactor: eq.rawData[17] || ''
        });
        linksStr = eq.rawData[18] || '';
      } else if (eq.type === 'Động cơ') {
        setFormData({
          vibration: eq.rawData[9] || '',
          statorTemp: eq.rawData[10] || '',
          ir: eq.rawData[11] || '',
          pd: eq.rawData[12] || '',
          voltageImbalance: eq.rawData[13] || '',
          pi: eq.rawData[14] || '',
          bearingTemp: eq.rawData[15] || '',
          tanDelta: eq.rawData[16] || '',
          age: eq.rawData[17] || '',
          dutyFactor: eq.rawData[18] || ''
        });
        linksStr = eq.rawData[19] || '';
      } else if (eq.type === 'Inverter') {
        setFormData({
          dcInputVoltage: eq.rawData[9] || '',
          mpptVoltageRange: eq.rawData[10] || '',
          dcInputCurrent: eq.rawData[11] || '',
          acOutputVoltage: eq.rawData[12] || '',
          acFrequency: eq.rawData[13] || '',
          harmonicDistortion: eq.rawData[14] || '',
          maxEfficiency: eq.rawData[15] || '',
          euroEfficiency: eq.rawData[16] || '',
          antiIslanding: eq.rawData[17] || '',
          rcdIsolation: eq.rawData[18] || '',
          dielectricVoltage: eq.rawData[19] || '',
          physicalDisconnects: eq.rawData[20] || '',
          operatingTemp: eq.rawData[21] || '',
          airFilters: eq.rawData[22] || ''
        });
        linksStr = eq.rawData[23] || '';
      }
    }

    if (linksStr) {
      try {
        const links = linksStr.split(',').map((l: string) => {
          const parts = l.split('|');
          return { name: parts[0].trim(), url: (parts[1] || parts[0]).trim() };
        });
        setAttachedFiles(links);
      } catch (e) {
        console.error("Error parsing links", e);
      }
    } else {
      setAttachedFiles([]);
    }
  };

  const generateWorkOrderPDF = async (order: any) => {
    const pdf = await getConfiguredJsPDF();
    const pageWidth = pdf.internal.pageSize.getWidth();
    
    // Add TEV Logo (using a placeholder or stylized text if image is not available)
    // We will draw a simple stylized TEV logo
    pdf.setFillColor(30, 64, 175); // Blue-800
    pdf.rect(20, 15, 30, 15, 'F');
    pdf.setTextColor(255, 255, 255);
    pdf.setFontSize(14);
    pdf.setFont('Roboto', 'bold');
    pdf.text('TEV', 35, 25, { align: 'center' });
    
    // Header
    pdf.setFontSize(20);
    pdf.setTextColor(30, 64, 175); // Blue-800
    pdf.text('PHIẾU CÔNG VIỆC / WORK ORDER', pageWidth / 2, 25, { align: 'center' });
    
    pdf.setFontSize(10);
    pdf.setTextColor(100, 116, 139); // Slate-500
    pdf.setFont('Roboto', 'normal');
    pdf.text(`Mã phiếu / WO ID: ${order.workPermitId || order.id.substring(0, 8)}`, pageWidth - 20, 20, { align: 'right' });
    pdf.text(`Ngày tạo / Created: ${order.createdAt ? new Date(order.createdAt).toLocaleDateString('vi-VN') : 'N/A'}`, pageWidth - 20, 25, { align: 'right' });

    // Content
    pdf.setFontSize(12);
    pdf.setTextColor(30, 41, 59); // Slate-800
    pdf.setFont('Roboto', 'bold');
    pdf.text('THÔNG TIN CHUNG / GENERAL INFO', 20, 45);
    pdf.line(20, 47, pageWidth - 20, 47);

    pdf.setFont('Roboto', 'normal');
    autoTable(pdf, {
      startY: 50,
      head: [['Hạng mục / Item', 'Chi tiết / Details']],
      body: [
        ['Tiêu đề / Title', order.title],
        ['Mô tả / Description', order.description || 'Không có mô tả / No description'],
        ['Thiết bị / Equipment', order.equipmentId],
        ['Khách hàng / Customer', order.customerId],
        ['Loại công việc / Type', order.type],
        ['Mức độ ưu tiên / Priority', order.priority],
        ['Trạng thái / Status', order.status],
        ['Người thực hiện / PIC', order.assignedTo || 'Chưa phân công / Unassigned'],
        ['Hạn hoàn thành / Due Date', order.dueDate ? new Date(order.dueDate).toLocaleDateString('vi-VN') : 'N/A'],
      ],
      theme: 'striped',
      headStyles: { fillColor: [51, 65, 85], font: 'Roboto', fontStyle: 'bold' },
      bodyStyles: { font: 'Roboto' },
    });

    // Materials
    if (order.usedMaterials && order.usedMaterials.length > 0) {
      const finalY = (pdf as any).lastAutoTable.finalY || 150;
      pdf.setFont('Roboto', 'bold');
      pdf.text('VẬT TƯ SỬ DỤNG / USED MATERIALS', 20, finalY + 15);
      pdf.line(20, finalY + 17, pageWidth - 20, finalY + 17);

      autoTable(pdf, {
        startY: finalY + 20,
        head: [['Tên vật tư / Material Name', 'Số lượng / Qty', 'Đơn vị / Unit']],
        body: order.usedMaterials.map((m: any) => [m.name, m.quantity, m.unit]),
        theme: 'grid',
        headStyles: { fillColor: [71, 85, 105], font: 'Roboto', fontStyle: 'bold' },
        bodyStyles: { font: 'Roboto' },
      });
    }

    // Signatures
    const finalY2 = (pdf as any).lastAutoTable.finalY || 200;
    pdf.setFontSize(10);
    pdf.setFont('Roboto', 'bold');
    pdf.text('Người thực hiện / Technician', 40, finalY2 + 30);
    pdf.setFont('Roboto', 'normal');
    pdf.text('(Ký và ghi rõ họ tên / Sign & Name)', 35, finalY2 + 35);
    
    pdf.setFont('Roboto', 'bold');
    pdf.text('Xác nhận khách hàng / Customer', pageWidth - 80, finalY2 + 30);
    pdf.setFont('Roboto', 'normal');
    pdf.text('(Ký và ghi rõ họ tên / Sign & Name)', pageWidth - 85, finalY2 + 35);

    pdf.save(`WorkOrder_${order.workPermitId || order.id.substring(0, 8)}.pdf`);
  };

  const handleSaveInventoryItem = async () => {
    if (!newInventoryItem.name || !newInventoryItem.sku) {
      alert('Vui lòng nhập tên và mã SKU');
      return;
    }

    try {
      let invId = '';
      const now = new Date().toISOString();
      if (selectedInventoryItem) {
        invId = selectedInventoryItem.id;
        await updateDoc(doc(db, 'inventory', invId), {
          ...newInventoryItem,
          updatedAt: now
        });
      } else {
        const docRef = await addDoc(collection(db, 'inventory'), {
          ...newInventoryItem,
          createdAt: now,
          updatedAt: now
        });
        invId = docRef.id;
      }

      // Sync to Google Sheets
      await syncToSheet('QuanLyKho', [[
        invId,
        newInventoryItem.name,
        newInventoryItem.sku,
        newInventoryItem.category,
        newInventoryItem.quantity,
        newInventoryItem.unit,
        newInventoryItem.minStock,
        newInventoryItem.location,
        newInventoryItem.price,
        selectedInventoryItem ? selectedInventoryItem.createdAt : now,
        now
      ]]);

      setShowInventoryModal(false);
      setNewInventoryItem({
        name: '',
        sku: '',
        category: '',
        quantity: 0,
        unit: 'pcs',
        minStock: 5,
        location: '',
        price: 0
      });
      setSelectedInventoryItem(null);
      alert('Đã lưu vật tư thành công!');
    } catch (error) {
      console.error("Error saving inventory item:", error);
      alert('Lỗi khi lưu vật tư');
    }
  };

  const handleSaveCustomer = async () => {
    if (!newCustomer.name) {
      alert('Vui lòng nhập tên khách hàng');
      return;
    }

    try {
      let custId = '';
      const now = new Date().toISOString();
      if (editingCustomer) {
        custId = editingCustomer.id;
        await updateDoc(doc(db, 'customers', custId), {
          ...newCustomer,
          updatedAt: now
        });
      } else {
        custId = generateNextCustomerId();
        await setDoc(doc(db, 'customers', custId), {
          ...newCustomer,
          createdAt: now,
          updatedAt: now
        });
      }

      setShowCustomerModal(false);
      setNewCustomer({ name: '', email: '', phone: '', address: '', factories: [] });
      setEditingCustomer(null);
    } catch (error) {
      console.error("Error saving customer:", error);
      alert('Lỗi khi lưu khách hàng');
    }
  };

  const handleDeleteCustomer = async (id: string) => {
    setConfirmDialog({
      isOpen: true,
      message: 'Bạn có chắc chắn muốn xóa khách hàng này?',
      onConfirm: async () => {
        try {
          await deleteDoc(doc(db, 'customers', id));
          setConfirmDialog(null);
        } catch (error) {
          console.error("Error deleting customer:", error);
          alert('Lỗi khi xóa khách hàng');
        }
      }
    });
  };

  const handleDeleteWorkOrder = async (id: string) => {
    setConfirmDialog({
      isOpen: true,
      message: 'Bạn có chắc chắn muốn xóa phiếu công việc này?',
      onConfirm: async () => {
        try {
          await deleteDoc(doc(db, 'workOrders', id));
          setConfirmDialog(null);
        } catch (error) {
          console.error("Error deleting work order:", error);
          setWoError('Lỗi khi xóa phiếu công việc');
        }
      }
    });
  };

  const generateWorkPermitId = () => {
    const currentYear = new Date().getFullYear();
    const wpPrefix = `WP-${currentYear}-`;
    const currentYearWPs = workOrders
      .map(wo => wo.workPermitId)
      .filter(id => id && id.startsWith(wpPrefix));
    
    let maxNum = 0;
    currentYearWPs.forEach(id => {
      const numStr = id.replace(wpPrefix, '');
      const num = parseInt(numStr, 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    });
    
    const nextNum = maxNum + 1;
    return `${wpPrefix}${nextNum.toString().padStart(3, '0')}`;
  };

  const generateNextCustomerId = () => {
    const prefix = 'KH-';
    const existingIds = customers
      .map(c => c.id)
      .filter(id => id && id.startsWith(prefix));
    
    let maxNum = 0;
    existingIds.forEach(id => {
      const numStr = id.replace(prefix, '');
      const num = parseInt(numStr, 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    });
    
    const nextNum = maxNum + 1;
    return `${prefix}${nextNum.toString().padStart(3, '0')}`;
  };

  const handleSaveWorkOrder = async () => {
    setWoError('');
    if (!newWorkOrder.title || !newWorkOrder.equipmentId || (Array.isArray(newWorkOrder.equipmentId) && newWorkOrder.equipmentId.length === 0) || !newWorkOrder.customerId) {
      setWoError('Vui lòng nhập tiêu đề, chọn ít nhất một thiết bị và chọn khách hàng');
      return;
    }

    try {
      let woId = '';
      const now = new Date().toISOString();
      
      // Auto-populate timestamps based on status
      const updatedWO = { ...newWorkOrder };
      
      if (updatedWO.status === 'in-progress' && !updatedWO.repairStart) {
        updatedWO.repairStart = now;
      }
      
      if (updatedWO.status === 'completed') {
        if (!updatedWO.repairEnd) updatedWO.repairEnd = now;
        if (!updatedWO.restartTime) updatedWO.restartTime = now;
      }

      if (updatedWO.isUnplanned && !updatedWO.downtimeStart) {
        updatedWO.downtimeStart = selectedWorkOrder ? (selectedWorkOrder.downtimeStart || now) : now;
      }

      if (selectedWorkOrder) {
        // Update
        woId = selectedWorkOrder.id;
        await updateDoc(doc(db, 'workOrders', woId), {
          ...updatedWO,
          updatedAt: now,
          completedAt: updatedWO.status === 'completed' ? now : (selectedWorkOrder?.completedAt || null)
        });
      } else {
        // Create
        const docRef = await addDoc(collection(db, 'workOrders'), {
          ...updatedWO,
          createdAt: now,
          updatedAt: now,
          completedAt: updatedWO.status === 'completed' ? now : null
        });
        woId = docRef.id;
      }

      // Sync to Google Sheets
      const customerNameStr = customers.find(c => c.id === updatedWO.customerId)?.name || updatedWO.customerId;
      const attachmentsStr = updatedWO.attachments ? updatedWO.attachments.join(', ') : '';
      const equipmentIdsStr = Array.isArray(updatedWO.equipmentId) ? updatedWO.equipmentId.join(', ') : updatedWO.equipmentId;
      await syncToSheet('CMMS', [[
        woId,
        updatedWO.workPermitId,
        updatedWO.title,
        updatedWO.description,
        equipmentIdsStr,
        customerNameStr,
        updatedWO.factory || '',
        updatedWO.type,
        updatedWO.priority,
        updatedWO.status,
        updatedWO.isUnplanned ? 'Unplanned' : 'Planned',
        updatedWO.assignedTo,
        updatedWO.responsibleApprove,
        updatedWO.dueDate,
        selectedWorkOrder ? selectedWorkOrder.createdAt : now,
        now,
        updatedWO.pmFrequency || '',
        updatedWO.failureCode || '',
        updatedWO.rootCause || '',
        updatedWO.actualTimeSpent || 0,
        updatedWO.estimatedTime || 0,
        updatedWO.laborCount || 0,
        updatedWO.laborCost || 0,
        updatedWO.partCost || 0,
        attachmentsStr,
        updatedWO.downtimeStart || '',
        updatedWO.repairStart || '',
        updatedWO.repairEnd || '',
        updatedWO.restartTime || ''
      ]]);

      setShowWorkOrderModal(false);
      setNewWorkOrder({
        title: '',
        description: '',
        equipmentId: [],
        customerId: '',
        factory: '',
        priority: 'medium',
        type: 'preventive',
        status: 'initiated',
        assignedTo: '',
        dueDate: '',
        usedMaterials: [],
        responsibleApprove: 'TM',
        responsibleDo: 'Everybody',
        blockingRequired: false,
        workPermitId: '',
        isUnplanned: false,
        failureCode: '',
        rootCause: '',
        actualTimeSpent: 0,
        pmFrequency: '',
        estimatedTime: 0,
        laborCount: 0,
        laborCost: 0,
        partCost: 0,
        attachments: [],
        downtimeStart: '',
        repairStart: '',
        repairEnd: '',
        restartTime: ''
      });
      setSelectedWorkOrder(null);
    } catch (error) {
      console.error("Error saving work order:", error);
      setWoError('Lỗi: ' + (error instanceof Error ? error.message : String(error)));
    }
  };
  const handleQRScan = (decodedText: string) => {
    // Stop the scanner
    setShowQRScanner(false);
    
    // Search for equipment by ID or Name
    const foundEq = allEquipment.find(eq => eq.id.toLowerCase() === decodedText.toLowerCase() || eq.name.toLowerCase() === decodedText.toLowerCase());
    
    if (foundEq) {
      setEquipmentCode(foundEq.id);
      setEquipmentName(foundEq.name);
      setCustomerName(foundEq.customer);
      setSiteName(foundEq.factory);
      setLocationName(foundEq.location);
      setSelectedEqType(foundEq.type === 'Tủ điện' ? 'Tủ điện trung thế' : foundEq.type);
      setIsNewEquipment(false);
      populateFormData(foundEq);
      setShowEquipmentProfile(true);
      alert(`Đã tìm thấy thiết bị: ${foundEq.name} (${foundEq.id}). Đang hiển thị hồ sơ chi tiết.`);
    } else {
      setEquipmentCode(decodedText);
      setIsNewEquipment(true);
      alert(`Không tìm thấy thiết bị với mã: ${decodedText}. Bạn có thể tạo mới.`);
    }
  };

  useEffect(() => {
    if (showQRScanner) {
      const html5QrcodeScanner = new Html5QrcodeScanner(
        "qr-reader",
        { fps: 10, qrbox: { width: 250, height: 250 } },
        /* verbose= */ false
      );
      html5QrcodeScanner.render(
        (decodedText) => {
          handleQRScan(decodedText);
          html5QrcodeScanner.clear();
        },
        (error) => {
          // Ignore scanning errors (happens constantly when no QR is in view)
        }
      );

      return () => {
        html5QrcodeScanner.clear().catch(error => {
          console.error("Failed to clear html5QrcodeScanner. ", error);
        });
      };
    }
  }, [showQRScanner, allEquipment]);

  const currentDetailedRiskData = useMemo(() => {
    if (!selectedRiskDetail) return null;
    
    // Generate mock detailed data based on the equipment
    const eq = allEquipment.find(e => e.id === selectedRiskDetail);
    if (!eq) return null;

    const isCritical = eq.status === 'critical';
    const isWarning = eq.status === 'warning';
    const baseRisk = 100 - eq.health;
    const getStatus = (score: number) => {
      if (score >= 8) return 'Excellent';
      if (score >= 6) return 'Good';
      if (score >= 4) return 'Moderate';
      return 'Critical';
    };
    const getRec = (score: number, critRec: string, modRec: string) => {
      if (score < 4) return critRec;
      if (score < 6) return modRec;
      return 'Bình thường';
    };

    let generalRecommendation = '';
    let failureProfile: any[] = [];
    let generatedDefects: any[] = [];
    let tests: any[] = [];

    if (eq.type === 'Máy biến áp') {
      const raw = eq.rawData || [];
      // Generate realistic values based on DNO Common Network Asset Indices Methodology
      const moisture = raw[17] ? parseFloat(raw[17].toString().replace(',', '.')) : (isCritical ? 35 + Math.random() * 15 : isWarning ? 20 + Math.random() * 15 : 5 + Math.random() * 10);
      const acidity = isCritical ? 0.2 + Math.random() * 0.15 : isWarning ? 0.12 + Math.random() * 0.08 : 0.02 + Math.random() * 0.08;
      const bdv = raw[15] ? parseFloat(raw[15].toString().replace(',', '.')) : (isCritical ? 25 + Math.random() * 10 : isWarning ? 35 + Math.random() * 10 : 55 + Math.random() * 15);
      const ffa = raw[16] ? parseFloat(raw[16].toString().replace(',', '.')) : (isCritical ? 5.5 + Math.random() * 2 : isWarning ? 3.5 + Math.random() * 2 : 0.5 + Math.random() * 2);
      
      // Scoring based on PDF calibration tables (Inverted: 10 is Good, 0 is Bad)
      const moistureScore = moisture > 45 ? 0 : moisture > 35 ? 2 : moisture > 25 ? 6 : moisture > 15 ? 8 : 10;
      const acidityScore = acidity > 0.3 ? 0 : acidity > 0.2 ? 2 : acidity > 0.15 ? 6 : acidity > 0.1 ? 8 : 10;
      const bdvScore = bdv < 30 ? 0 : bdv < 40 ? 6 : bdv < 50 ? 8 : 10;
      const ffaScore = ffa > 7 ? 0 : ffa > 6 ? 2 : ffa > 5 ? 4 : ffa > 4 ? 6 : 10;
      const dgaScore = isCritical ? 1 : isWarning ? 5 : 10;
      const pdScore = isCritical ? 2 : isWarning ? 6 : 10;

      generalRecommendation = isCritical 
        ? 'Cảnh báo mức độ cao. Cần tiến hành kiểm tra DGA, đo PD và lên kế hoạch bảo dưỡng/thay thế ngay lập tức.' 
        : isWarning 
          ? 'Thiết bị đang ở trạng thái cảnh báo. Cần tăng cường tần suất theo dõi tình trạng dầu và nhiệt độ.'
          : 'Thiết bị hoạt động bình thường. Tiếp tục bảo dưỡng định kỳ theo khuyến cáo của nhà sản xuất.';

      failureProfile = [
        { subject: 'Tình trạng bên ngoài', score: Math.min(100, 100 - (baseRisk + Math.random() * 20)) },
        { subject: 'Phóng điện cục bộ (PD)', score: Math.min(100, pdScore * 10) },
        { subject: 'Chất lượng dầu', score: Math.min(100, Math.min(moistureScore, acidityScore, bdvScore) * 10) },
        { subject: 'Khí hòa tan (DGA)', score: Math.min(100, dgaScore * 10) },
        { subject: 'Lão hóa giấy (FFA)', score: Math.min(100, ffaScore * 10) },
      ];

      generatedDefects = [
        { name: 'Suy giảm chất lượng dầu (Oil Degradation)', warnings: (10 - Math.min(moistureScore, acidityScore)) * 2.5, risk: (10 - Math.min(moistureScore, acidityScore)) * 10 },
        { name: 'Phóng điện cục bộ (Partial Discharge)', warnings: (10 - pdScore) * 2, risk: (10 - pdScore) * 10 },
        { name: 'Lão hóa giấy cách điện (Paper Ageing)', warnings: (10 - ffaScore) * 2, risk: (10 - ffaScore) * 10 },
      ].filter(d => d.risk > 0);

      tests = [
        { name: 'Khí hòa tan (DGA)', shortName: 'DGA', value: raw[14] || (isCritical ? 'Mức cao' : isWarning ? 'Cảnh báo' : 'Bình thường'), unit: '', score: dgaScore, weight: 3, status: getStatus(dgaScore), rec: getRec(dgaScore, 'Lọc/thay dầu ngay', 'Theo dõi thêm') },
        { name: 'Hàm lượng Furan (FFA)', shortName: 'FFA', value: ffa.toFixed(2), unit: 'ppm', score: ffaScore, weight: 2, status: getStatus(ffaScore), rec: getRec(ffaScore, 'Kiểm tra cách điện giấy', 'Theo dõi xu hướng') },
        { name: 'Độ ẩm trong dầu', shortName: 'Moisture', value: moisture.toFixed(1), unit: 'ppm', score: moistureScore, weight: 2, status: getStatus(moistureScore), rec: getRec(moistureScore, 'Xử lý lọc dầu', 'Kiểm tra độ kín') },
        { name: 'Độ axit', shortName: 'Acidity', value: acidity.toFixed(3), unit: 'mg KOH/g', score: acidityScore, weight: 1, status: getStatus(acidityScore), rec: getRec(acidityScore, 'Thay dầu', 'Theo dõi') },
        { name: 'Điện áp đánh thủng', shortName: 'BDV', value: bdv.toFixed(1), unit: 'kV', score: bdvScore, weight: 2, status: getStatus(bdvScore), rec: getRec(bdvScore, 'Xử lý lọc dầu', 'Theo dõi thêm') },
        { name: 'Phóng điện cục bộ', shortName: 'PD', value: isCritical ? 'Cao' : isWarning ? 'Trung bình' : 'Thấp', unit: '', score: pdScore, weight: 3, status: getStatus(pdScore), rec: getRec(pdScore, 'Kiểm tra siêu âm/TEV', 'Theo dõi định kỳ') },
      ];
    } else if (eq.type === 'Tủ điện trung thế' || eq.type === 'Tủ điện') {
      const raw = eq.rawData || [];
      const tevLevel = raw[11] ? parseFloat(raw[11].toString().replace(',', '.')) : (isCritical ? 30 + Math.random() * 20 : isWarning ? 20 + Math.random() * 9 : 5 + Math.random() * 14);
      const pulses = raw[13] ? parseFloat(raw[13].toString().replace(',', '.')) : (isCritical ? Math.floor(1 + Math.random() * 28) : isWarning ? Math.floor(1 + Math.random() * 5) : 0);
      const ultrasonic = raw[12] ? parseFloat(raw[12].toString().replace(',', '.')) : (isCritical ? 20 + Math.random() * 20 : isWarning ? 5 + Math.random() * 15 : Math.random() * 5);

      const tevScore = tevLevel > 29 ? 1 : tevLevel >= 20 ? 4 : 10;
      const pulsesScore = pulses >= 7 && pulses <= 29 ? 1 : pulses >= 1 && pulses <= 6 ? 4 : pulses >= 30 ? 6 : 10; 
      const ultrasonicScore = ultrasonic > 15 ? 1 : ultrasonic > 5 ? 4 : 10;

      generalRecommendation = isCritical 
          ? 'Khả năng rất cao có hiện tượng phóng điện cục bộ. Yêu cầu kiểm tra vào lần dừng máy kế tiếp hoặc dừng máy ngay để xác định nguồn.' 
          : isWarning 
            ? 'Có khả năng xảy ra phóng điện cục bộ. Kiểm tra số xung/chu kỳ và theo dõi định kỳ.'
            : 'Không có dấu hiệu phóng điện cục bộ. Kiểm tra lại trong vòng 12 tháng.';

      failureProfile = [
          { subject: 'Tình trạng bên ngoài', score: Math.min(100, 100 - (baseRisk + Math.random() * 20)) },
          { subject: 'Phóng điện TEV', score: Math.min(100, tevScore * 10) },
          { subject: 'Phóng điện siêu âm', score: Math.min(100, ultrasonicScore * 10) },
          { subject: 'Mật độ xung', score: Math.min(100, pulsesScore * 10) },
      ];

      generatedDefects = [
        { name: 'Phóng điện bề mặt (Surface PD)', warnings: (10 - ultrasonicScore) * 2, risk: (10 - ultrasonicScore) * 10 },
        { name: 'Phóng điện bên trong (Internal PD)', warnings: (10 - tevScore) * 2, risk: (10 - tevScore) * 10 },
      ].filter(d => d.risk > 10);

      tests = [
          { name: 'Mức TEV', shortName: 'TEV', value: tevLevel.toFixed(1), unit: 'dB', score: tevScore, weight: 3, status: getStatus(tevScore), rec: getRec(tevScore, 'Dừng máy kiểm tra', 'Theo dõi') },
          { name: 'Số xung/chu kỳ', shortName: 'Pulses', value: pulses.toString(), unit: 'xung', score: pulsesScore, weight: 2, status: getStatus(pulsesScore), rec: getRec(pulsesScore, 'Định vị nguồn PD', 'Kiểm tra lại') },
          { name: 'Siêu âm', shortName: 'Ultrasonic', value: ultrasonic.toFixed(1), unit: 'dBµV', score: ultrasonicScore, weight: 2, status: getStatus(ultrasonicScore), rec: getRec(ultrasonicScore, 'Kiểm tra phóng điện bề mặt', 'Theo dõi') },
      ];
    } else if (eq.type === 'Động cơ') {
      const raw = eq.rawData || [];
      const tanDelta = raw[13] ? parseFloat(raw[13].toString().replace(',', '.')) : (isCritical ? 0.08 + Math.random() * 0.05 : isWarning ? 0.04 + Math.random() * 0.03 : 0.01 + Math.random() * 0.02);
      const pd = raw[19] ? parseFloat(raw[19].toString().replace(',', '.')) : (isCritical ? 18000 + Math.random() * 5000 : isWarning ? 8000 + Math.random() * 7000 : 2000 + Math.random() * 3000);
      const ir = raw[22] ? parseFloat(raw[22].toString().replace(',', '.')) : (isCritical ? 0.5 + Math.random() * 0.5 : isWarning ? 5 + Math.random() * 5 : 60 + Math.random() * 50);
      const pi = raw[25] ? parseFloat(raw[25].toString().replace(',', '.')) : (isCritical ? 0.8 + Math.random() * 0.2 : isWarning ? 1.2 + Math.random() * 0.8 : 2.5 + Math.random() * 1.5);
      const elcid = raw[31] ? parseFloat(raw[31].toString().replace(',', '.')) : (isCritical ? 250 + Math.random() * 100 : isWarning ? 120 + Math.random() * 80 : 40 + Math.random() * 30);

      const tanDeltaScore = tanDelta >= 0.1 ? 1 : tanDelta >= 0.07 ? 3 : tanDelta >= 0.04 ? 6 : 10;
      const pdScore = pd >= 20000 ? 1 : pd >= 15000 ? 3 : pd >= 10000 ? 6 : 10;
      const irScore = ir < 1 ? 1 : ir < 10 ? 3 : ir < 50 ? 6 : 10;
      const piScore = pi < 1 ? 1 : pi < 2 ? 5 : 10;
      const elcidScore = elcid > 200 ? 1 : elcid > 110 ? 3 : elcid > 70 ? 6 : 10;

      generalRecommendation = isCritical 
          ? 'Tình trạng cách điện suy giảm nghiêm trọng. Cần lên kế hoạch bảo dưỡng, sấy cuộn dây hoặc quấn lại.' 
          : isWarning 
            ? 'Có dấu hiệu lão hóa cách điện. Tăng cường theo dõi PD và điện trở cách điện.'
            : 'Cách điện động cơ ở trạng thái tốt. Tiếp tục vận hành và bảo dưỡng định kỳ.';

      failureProfile = [
          { subject: 'Tổn hao điện môi (Tan-delta)', score: Math.min(100, tanDeltaScore * 10) },
          { subject: 'Phóng điện cục bộ (PD)', score: Math.min(100, pdScore * 10) },
          { subject: 'Điện trở cách điện (IR)', score: Math.min(100, irScore * 10) },
          { subject: 'Chỉ số phân cực (PI)', score: Math.min(100, piScore * 10) },
          { subject: 'Lõi thép (ELCID)', score: Math.min(100, elcidScore * 10) },
      ];

      generatedDefects = [
        { name: 'Lão hóa cách điện cuộn dây', warnings: (10 - irScore) * 2, risk: (10 - irScore) * 10 },
        { name: 'Phóng điện cục bộ stator', warnings: (10 - pdScore) * 2, risk: (10 - pdScore) * 10 },
        { name: 'Hỏng hóc lõi thép', warnings: (10 - elcidScore) * 2, risk: (10 - elcidScore) * 10 },
      ].filter(d => d.risk > 10);

      tests = [
          { name: 'Tan-delta', shortName: 'Tan-δ', value: tanDelta.toFixed(3), unit: '', score: tanDeltaScore, weight: 2, status: getStatus(tanDeltaScore), rec: getRec(tanDeltaScore, 'Kiểm tra toàn diện', 'Theo dõi') },
          { name: 'Phóng điện cục bộ', shortName: 'PD', value: Math.round(pd).toString(), unit: 'pC', score: pdScore, weight: 3, status: getStatus(pdScore), rec: getRec(pdScore, 'Xử lý cách điện', 'Theo dõi') },
          { name: 'Điện trở cách điện', shortName: 'IR', value: ir.toFixed(1), unit: 'GΩ', score: irScore, weight: 1, status: getStatus(irScore), rec: getRec(irScore, 'Sấy cuộn dây', 'Kiểm tra định kỳ') },
          { name: 'Chỉ số phân cực', shortName: 'PI', value: pi.toFixed(2), unit: '', score: piScore, weight: 1, status: getStatus(piScore), rec: getRec(piScore, 'Vệ sinh, sấy', 'Theo dõi') },
          { name: 'Kiểm tra lõi thép', shortName: 'ELCID', value: Math.round(elcid).toString(), unit: 'mA', score: elcidScore, weight: 1, status: getStatus(elcidScore), rec: getRec(elcidScore, 'Sửa chữa lõi từ', 'Theo dõi') },
      ];
    }

    if (generatedDefects.length === 0) {
      generatedDefects.push({ name: 'Hao mòn thông thường (Normal Wear)', warnings: 1, risk: 5 });
    }

    return {
      id: eq.id,
      name: eq.name,
      type: eq.type,
      health: eq.health,
      status: eq.status,
      factory: eq.factory,
      location: eq.location,
      generalRecommendation,
      failureProfile,
      defects: generatedDefects,
      tests
    };
  }, [selectedRiskDetail, allEquipment]);

  const [selectedCustomer, setSelectedCustomer] = useState('all');
  const [selectedFactory, setSelectedFactory] = useState('all');
  const [isGoogleConnected, setIsGoogleConnected] = useState(false);
  const [spreadsheetId, setSpreadsheetId] = useState('');
  const [isSaving, setIsSaving] = useState(false);
  const [saveSuccess, setSaveSuccess] = useState(false);
  
  // New states for improvements
  const [currentPage, setCurrentPage] = useState(1);
  const [statusFilter, setStatusFilter] = useState('all');
  const [filterName, setFilterName] = useState('');
  const [filterLocation, setFilterLocation] = useState('');
  const [filterType, setFilterType] = useState('all');
  const [filterHealthMin, setFilterHealthMin] = useState('');
  const [filterHealthMax, setFilterHealthMax] = useState('');
  const [showAddEqModal, setShowAddEqModal] = useState(false);
  const [attachedFiles, setAttachedFiles] = useState<File[]>([]);
  const [newEqData, setNewEqData] = useState({ id: '', name: '', customer: '', factory: '', location: '', type: 'Máy biến áp' });
  const [trendingChartType, setTrendingChartType] = useState('temp'); // 'temp' or 'tev'
  const [trendingEquipment, setTrendingEquipment] = useState('TRF-01');
  const [isSyncing, setIsSyncing] = useState(false);
  const [syncSuccess, setSyncSuccess] = useState(false);
  
  // Reports State
  const [allReports, setAllReports] = useState(initialAllReports);
  const [reportFilterSearch, setReportFilterSearch] = useState('');
  const [reportFilterFactory, setReportFilterFactory] = useState('all');
  const [reportFilterType, setReportFilterType] = useState('all');
  const [reportFilterStatus, setReportFilterStatus] = useState('all');
  const [reportFilterStartDate, setReportFilterStartDate] = useState('');
  const [reportFilterEndDate, setReportFilterEndDate] = useState('');
  const [reportCurrentPage, setReportCurrentPage] = useState(1);
  const reportsPerPage = 15;

  // Field Entry Mode: NETA ATS-2025 Liquid Transformers, Solar PV, BESS, Dry-Type, vs Standard
  const [fieldEntryMode, setFieldEntryMode] = useState<FieldEntryMode>('neta-liquid-transformer');
  const [fieldEntrySubTab, setFieldEntrySubTab] = useState<'station' | 'checklists'>('station');

  const handleSaveNetaReport = async (reportPayload: any) => {
    try {
      const { data: netaData, analysisResult, equipmentId, date, type: payloadType } = reportPayload;
      const reportId = `REP-NETA-${Date.now()}`;
      const statusMap = {
        PASS: 'healthy' as const,
        INVESTIGATE: 'warning' as const,
        REPAIR_SCHEDULED: 'warning' as const,
        MONITOR: 'warning' as const,
        REPAIR_PRACTICAL: 'warning' as const,
        REPAIR_IMMEDIATELY: 'critical' as const,
        REPAIR_URGENT: 'critical' as const,
        NORMAL_TRENDING: 'healthy' as const,
        FAIL: 'critical' as const,
      };
      const calculatedStatus = statusMap[analysisResult?.overall_status as keyof typeof statusMap] || 'warning';

      const isGenerator = payloadType?.includes('ENGINE_GENERATOR') || payloadType?.includes('GENERATOR') || Boolean(netaData?.site_info?.generator_tag);
      const isSwitchgear = !isGenerator && (payloadType?.includes('Switchgear') || payloadType?.includes('Sec 7.1.1') || Boolean(netaData?.site_info?.switchgear_tag));
      const isAts = !isGenerator && !isSwitchgear && (payloadType?.includes('AUTOMATIC_TRANSFER_SWITCH') || payloadType?.includes('ATS') || Boolean(netaData?.site_info?.ats_tag));
      const isBatteryVrla = payloadType?.includes('BATTERY_VRLA') || payloadType?.includes('VRLA') || (Boolean(netaData?.site_info?.battery_bank_tag) && (netaData?.electrical_tests?.negative_post_temperature_c !== undefined || netaData?.site_info?.battery_type?.includes('VRLA')));
      const isBatteryFlooded = !isBatteryVrla && (payloadType?.includes('BATTERY_FLOODED') || payloadType?.includes('Flooded') || Boolean(netaData?.site_info?.battery_bank_tag));
      const isLiquidTransformer = payloadType?.includes('Liquid-Filled') || payloadType?.includes('Sec 7.2.2') || Boolean(netaData?.electrical_tests?.oil_sample_tests_astm_d923);
      const isDcMotor = payloadType?.includes('DC') || payloadType?.includes('Sec 7.15.3') || Boolean(netaData?.electrical_tests?.armature_bar_to_bar_resistance_micro_ohms);
      const isMotor = !isDcMotor && (payloadType?.includes('Rotating') || payloadType?.includes('Motor') || Boolean(netaData?.site_info?.motor_tag));
      const isPd = payloadType?.includes('Partial Discharge') || payloadType?.includes('PD') || Boolean(netaData?.pd_sensor_measurements);
      const isThermography = payloadType?.includes('Thermograph') || Boolean(netaData?.site_info?.survey_tag);
      const isSf6Switch = payloadType?.includes('SF6') || Boolean(netaData?.site_info?.switch_tag);
      const isGrounding = payloadType?.includes('Grounding') || payloadType?.includes('Sec 7.13') || Boolean(netaData?.site_info?.grounding_system_tag) || Boolean(netaData?.electrical_tests?.fall_of_potential_ground_resistance_ieee_81_ohms);
      const isLvBreaker = payloadType?.includes('Circuit Breakers') || payloadType?.includes('ACB') || payloadType?.includes('Sec 7.6.1.2') || (Boolean(netaData?.site_info?.breaker_tag) && Boolean(netaData?.electrical_tests?.secondary_injection_trip_unit_test));
      const isCableLv = payloadType?.includes('Low-Voltage') || payloadType?.includes('Sec 7.3.2') || (Boolean(netaData?.site_info?.cable_tag) && Boolean(netaData?.electrical_tests?.parallel_conductors_dc_resistance_milli_ohms));
      const isCable = !isCableLv && (payloadType?.includes('Cable') || Boolean(netaData?.site_info?.cable_tag));
      const isEvse = payloadType?.includes('Electric Vehicle') || payloadType?.includes('EVSE') || Boolean(netaData?.site_info?.evse_tag);
      const isPv = payloadType?.includes('Solar') || Boolean(netaData?.site_info?.pv_array_tag);
      const isBess = payloadType?.includes('BESS') || Boolean(netaData?.site_info?.bess_container_tag);
      const isLarge = payloadType?.includes('Large') || Boolean(netaData?.electrical_tests?.core_insulation_resistance_500v_dc_megohms !== undefined);

      const equipTag = equipmentId || netaData?.site_info?.grounding_system_tag || netaData?.site_info?.breaker_tag || netaData?.site_info?.switchgear_tag || netaData?.site_info?.generator_tag || netaData?.site_info?.ats_tag || netaData?.site_info?.battery_bank_tag || netaData?.site_info?.motor_tag || netaData?.site_info?.equipment_tag || netaData?.site_info?.survey_tag || netaData?.site_info?.switch_tag || netaData?.site_info?.cable_tag || netaData?.site_info?.evse_tag || netaData?.site_info?.pv_array_tag || netaData?.site_info?.bess_container_tag || netaData?.site_info?.transformer_tag || 'EQ-01';
      let equipName = `Thiết bị ${equipTag}`;
      let equipType = 'Máy biến áp';
      let testStandard = 'ANSI/NETA ATS-2025';
      let reportNotes = `[${testStandard}] ${analysisResult?.status_reason || 'Đã phân tích kỹ thuật'}`;

      if (isGrounding) {
        equipName = `Hệ thống tiếp địa & nối đất ${equipTag}`;
        equipType = 'Grounding Systems (NETA ATS-2025 Mục 7.13 & IEEE Std 81)';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.13 & IEEE Std 81';
        reportNotes = `[Tiếp Địa Sec 7.13 & IEEE 81] ${analysisResult?.status_reason || 'Đã phân tích'}. Fall-of-Potential: ${netaData?.electrical_tests?.fall_of_potential_ground_resistance_ieee_81_ohms || 7.8} Ω (Chuẩn Trạm ≤ 1.0 Ω / CN ≤ 5.0 Ω), Cổng hàng rào: ${netaData?.electrical_tests?.point_to_point_continuity_milli_ohms?.main_bus_to_substation_fence_gate || 680} mΩ (Chuẩn ≤ 0.5 Ω)`;
      } else if (isLvBreaker) {
        equipName = `Máy cắt không khí hạ áp ${equipTag}`;
        equipType = 'Low-Voltage Power Circuit Breakers (ACB)';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.6.1.2';
        reportNotes = `[ACB Sec 7.6.1.2] ${analysisResult?.status_reason || 'Đã phân tích'}. Tiếp điểm Pole C: ${netaData?.electrical_tests?.contact_pole_resistance_micro_ohms?.pole_c} µΩ (Lệch ${analysisResult?.evaluations?.contact_resistance?.max_deviation_percent || 131.3}%), Trip Unit Ground Fault: ${netaData?.electrical_tests?.secondary_injection_trip_unit_test?.ground_fault_pickup_status}, Cách điện ĐK: ${netaData?.electrical_tests?.control_wiring_ir_megohms} MΩ`;
      } else if (isCableLv) {
        equipName = `Cáp điện hạ áp ${equipTag}`;
        equipType = 'Low-Voltage Cable (1,000V Max)';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.3.2';
        reportNotes = `[Cáp Hạ Áp Sec 7.3.2] ${analysisResult?.status_reason || 'Đã phân tích'}. Cách điện Pha C: ${netaData?.electrical_tests?.insulation_resistance_1000v_1min_megohms?.phase_c_to_ground} MΩ (Chuẩn ≥ 100 MΩ), Dây song song Pha C: Run 1 ${netaData?.electrical_tests?.parallel_conductors_dc_resistance_milli_ohms?.phase_c_run1} mΩ / Run 2 ${netaData?.electrical_tests?.parallel_conductors_dc_resistance_milli_ohms?.phase_c_run2} mΩ, Thông mạch: Đạt`;
      } else if (isSwitchgear) {
        equipName = `Tủ điện phân phối & tủ đóng cắt ${equipTag}`;
        equipType = 'Switchgear & Switchboard Assemblies';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.1.1';
        reportNotes = `[Switchgear Sec 7.1.1] ${analysisResult?.status_reason || 'Đã phân tích'}. Cách điện ĐK: ${netaData?.electrical_tests?.control_wiring_ir_1000v_megohms} MΩ (Chuẩn ≥ 2.0 MΩ), Thanh cái IR: ${netaData?.electrical_tests?.bus_insulation_resistance_2500v_1min_megohms?.phase_a_to_ground} MΩ, CPT Turns Ratio: ±${netaData?.electrical_tests?.cpt_tests?.turns_ratio_error_percent}%, TEV PD: ${netaData?.electrical_tests?.online_partial_discharge_tev_db} dB`;
      } else if (isGenerator) {
        equipName = `Tổ máy phát điện khẩn cấp ${equipTag}`;
        equipType = 'Emergency Engine Generator';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.22.1 & NFPA 110';
        reportNotes = `[Generator Sec 7.22.1 & NFPA 110] ${analysisResult?.status_reason || 'Đã phân tích'}. Quá tốc độ: ${analysisResult?.overspeed_status || 'Fail'}, PI Stator: ${analysisResult?.insulation_analysis?.pi_value || 1.7} (Chuẩn ≥ 2.0), Thời gian nhận tải: ${analysisResult?.nfpa_110_analysis?.start_time_sec || 8.5}s (NFPA 110 Class 10 ≤ 10s)`;
      } else if (isAts) {
        equipName = `Bộ chuyển nguồn tự động ATS ${equipTag}`;
        equipType = 'Automatic Transfer Switch (ATS)';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.22.3';
        reportNotes = `[ATS Sec 7.22.3] ${analysisResult?.status_reason || 'Đã phân tích'}. Tiếp điểm Normal C: ${netaData?.electrical_tests?.contact_pole_resistance_micro_ohms?.normal_source_pole_c || 68.4} µΩ (Lệch ${analysisResult?.contact_resistance_analysis?.normal_max_deviation_percent || 209.5}%), Cách điện ĐK: ${netaData?.electrical_tests?.control_wiring_ir_megohms || 45} MΩ, Trình tự tự động: Đạt`;
      } else if (isBatteryVrla) {
        equipName = `Giàn ắc quy VRLA AGM/GEL ${equipTag}`;
        equipType = 'Valve-Regulated Lead-Acid (VRLA) Battery';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.18.1.3 & IEEE 1188';
        reportNotes = `[Battery VRLA Sec 7.18.1.3] ${analysisResult?.status_reason || 'Đã phân tích'}. Nhiệt độ cực âm: ${netaData?.electrical_tests?.negative_post_temperature_c?.max_temp_monoblock_12_c || 41.2}°C (Avg ${netaData?.electrical_tests?.negative_post_temperature_c?.avg_temp_c || 26.5}°C), Biến động ôm: ${netaData?.electrical_tests?.internal_ohmic_measurement_resistance_mohm?.variance_percent || 45.3}%, Cân bằng đất: (+) ${netaData?.electrical_tests?.system_voltage_to_ground_v?.positive_to_ground_v}V / (-) ${netaData?.electrical_tests?.system_voltage_to_ground_v?.negative_to_ground_v}V`;
      } else if (isBatteryFlooded) {
        equipName = `Giàn ắc quy Flooded Lead-Acid ${equipTag}`;
        equipType = 'Flooded Lead-Acid Battery';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.18.1.1 & IEEE 450';
        reportNotes = `[Battery Flooded Sec 7.18.1.1] ${analysisResult?.status_reason || 'Đã phân tích'}. Lệch áp float: ${netaData?.electrical_tests?.cell_voltages_float_mode_v?.voltage_spread_v || 0.08}V, Biến động ôm: ${netaData?.electrical_tests?.internal_ohmic_measurement_resistance_mohm?.variance_percent || 37.7}%, Điện áp đất: (+) ${netaData?.electrical_tests?.system_voltage_to_ground_v?.positive_to_ground_v}V / (-) ${netaData?.electrical_tests?.system_voltage_to_ground_v?.negative_to_ground_v}V`;
      } else if (isLiquidTransformer) {
        equipName = `Máy biến áp ngâm dầu ${equipTag}`;
        equipType = 'Liquid-Filled Transformer';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.2.2';
        reportNotes = `[Liquid-Filled Sec 7.2.2] ${analysisResult?.status_reason || 'Đã phân tích'}. Đánh thủng dầu: ${netaData?.electrical_tests?.oil_sample_tests_astm_d923?.dielectric_breakdown_d1816_1mm_kv || 21}kV, Ẩm: ${netaData?.electrical_tests?.oil_sample_tests_astm_d923?.water_content_d1533_ppm || 32}ppm, DGA C2H2: ${netaData?.electrical_tests?.dga_test_ieee_c57_104?.acetylene_c2h2_ppm || 6.8}ppm`;
      } else if (isDcMotor) {
        equipName = `Động cơ / Máy phát điện một chiều DC ${equipTag}`;
        equipType = 'DC Motors & Generators (NETA ATS-2025 Mục 7.15.3)';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.15.3 & IEEE Std 43';
        reportNotes = `[DC Motor Sec 7.15.3] ${analysisResult?.status_reason || 'Đã phân tích'}. Rating: ${netaData?.site_info?.motor_rating || '250kW 440V DC'}, Bar-to-Bar: ${netaData?.electrical_tests?.armature_bar_to_bar_resistance_micro_ohms?.max_deviation_percent || 8.7}% (Chuẩn ≤ 5.0%), Sụt áp cực từ: ${analysisResult?.evaluations?.field_pole_voltage_drop?.max_deviation_percent || 0.8}% (Chuẩn ≤ 10%), Armature PI: ${netaData?.electrical_tests?.insulation_resistance_40c_megohms?.armature_pi_value || 2.78} (Chuẩn ≥ 2.0)`;
      } else if (isMotor) {
        equipName = `Động cơ / Máy phát điện quay ${equipTag}`;
        equipType = 'Rotating Machinery (AC Induction Motor)';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.15.1';
        reportNotes = `[Rotating Machinery Sec 7.15.1] ${analysisResult?.status_reason || 'Đã phân tích'}. Rating: ${netaData?.site_info?.motor_rating || '4160V 500HP'}, Rung: ${netaData?.electrical_tests?.vibration_test_table_100_10?.velocity_in_sec_pk || 0.22} in./s, MCSA: ${netaData?.electrical_tests?.current_signature_analysis_mcsa?.broken_bar_sideband_db || 35.2} dB`;
      } else if (isPd) {
        equipName = `Khảo sát phóng điện cục bộ online ${equipTag}`;
        equipType = 'Online Partial Discharge';
        testStandard = 'ANSI/NETA ATS-2025 Mục 11 & Bảng 100.23';
        reportNotes = `[Online PD Table 100.23] ${analysisResult?.status_reason || 'Đã phân tích'}. Điện áp: ${netaData?.survey_conditions?.system_voltage_kv || 22}kV, Dòng: ${netaData?.survey_conditions?.operating_current_amp || 410}A, Số cảm biến: ${netaData?.pd_sensor_measurements?.length || 3}`;
      } else if (isThermography) {
        equipName = `Khảo sát nhiệt độ hồng ngoại ${equipTag}`;
        equipType = 'Thermographic Survey';
        testStandard = 'ANSI/NETA ATS-2025 Mục 9 & Bảng 100.18';
        reportNotes = `[Thermography Table 100.18] ${analysisResult?.status_reason || 'Đã phân tích'}. Tải: ${netaData?.survey_conditions?.load_percentage || 78}%, Môi trường: ${netaData?.survey_conditions?.ambient_temperature_c || 32}°C`;
      } else if (isSf6Switch) {
        equipName = `Cầu dao ngắt mạch trung áp SF6 ${equipTag}`;
        equipType = 'SF6 Switch';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.5.4';
        reportNotes = `[SF6 Switch Sec 7.5.4] ${analysisResult?.status_reason || 'Đã phân tích'}. Khí SO2: ${netaData?.electrical_tests?.sf6_gas_quality_test?.so2_decomposition_ppmv || 18.5} ppmv, Purity: ${netaData?.electrical_tests?.sf6_gas_quality_test?.sf6_purity_percent || 96.2}%, Pole C: ${netaData?.electrical_tests?.contact_pole_resistance_micro_ohms?.pole_c || 82.5}µΩ`;
      } else if (isCable) {
        equipName = `Cáp trung/cao áp có màn chắn ${equipTag}`;
        equipType = 'Shielded Cable';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.3.3';
        reportNotes = `[Cable Sec 7.3.3] ${analysisResult?.status_reason || 'Đã phân tích'}. Màn chắn Pha A: ${netaData?.electrical_tests?.shield_resistance_ohms?.phase_a_shield || 8.2}Ω, Pha B: ${netaData?.electrical_tests?.shield_resistance_ohms?.phase_b_shield || 8.5}Ω, Pha C: ${netaData?.electrical_tests?.shield_resistance_ohms?.phase_c_shield || 24.6}Ω, VLF: ${netaData?.electrical_tests?.vlf_withstand_test_0_1hz_30min?.test_voltage_kv_rms || 28}kV`;
      } else if (isEvse) {
        equipName = `Trạm sạc xe điện EVSE ${equipTag}`;
        equipType = 'EVSE';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.26';
        reportNotes = `[EVSE Sec 7.26] ${analysisResult?.status_reason || 'Đã phân tích'}. Súng 1 PE: ${netaData?.electrical_tests?.protective_conductor_resistance_ohms?.gun_1_pe_resistance || 0.18}Ω, Súng 2 PE: ${netaData?.electrical_tests?.protective_conductor_resistance_ohms?.gun_2_pe_resistance || 0.68}Ω, V_DC: ${netaData?.electrical_tests?.charger_output_voltage_v_dc || 402}V, Ripple: ${netaData?.electrical_tests?.dc_voltage_ripple_percent || 0.8}%`;
      } else if (isPv) {
        equipName = `Hệ thống Điện mặt trời PV ${equipTag}`;
        equipType = 'Solar PV';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.29';
        reportNotes = `[Solar PV Sec 7.29] ${analysisResult?.status_reason || 'Đã phân tích'}. Voc Chuỗi 3: ${netaData?.electrical_tests?.open_circuit_voltage_voc?.measured_voc_v?.[2] || 850}V, IR Chuỗi 3: ${netaData?.electrical_tests?.dry_insulation_resistance_megohms?.string_03_to_ground || 12}MΩ, Nối đất: ${netaData?.electrical_tests?.ground_resistance_section_7_13_ohms || 2.1}Ω`;
      } else if (isBess) {
        equipName = `Hệ thống Pin Lưu Trữ BESS ${equipTag}`;
        equipType = 'BESS';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.28 & IEEE 1547';
        reportNotes = `[BESS Sec 7.28] ${analysisResult?.status_reason || 'Đã phân tích'}. V_DC: ${netaData?.electrical_and_subsystem_tests?.measured_total_dc_voltage_v}V, THDu: ${netaData?.electrical_and_subsystem_tests?.power_quality_at_poi_ieee_1547?.voltage_thd_percent}%, PCCC: ${netaData?.visual_inspection?.fire_suppression_installed ? 'Đạt' : 'Lỗi'}`;
      } else if (isLarge) {
        equipName = `Máy biến áp khô dung lượng lớn ${equipTag}`;
        equipType = 'Máy biến áp';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.2.1.2';
        reportNotes = `[MBA Khô Lớn Sec 7.2.1.2] ${analysisResult?.status_reason || 'Đã phân tích'}. Core IR: ${netaData?.electrical_tests?.core_insulation_resistance_500v_dc_megohms}MΩ, PI: ${netaData?.electrical_tests?.pi_value}`;
      } else {
        equipName = `Máy biến áp khô hạ áp ${equipTag}`;
        equipType = 'Máy biến áp';
        testStandard = 'ANSI/NETA ATS-2025 Mục 7.2.1.1';
        reportNotes = `[MBA Khô Nhỏ Sec 7.2.1.1] ${analysisResult?.status_reason || 'Đã phân tích'}. IR: ${netaData?.electrical_tests?.insulation_resistance_1min_1000v_megohms?.pri_to_ground}MΩ, DAR: ${netaData?.electrical_tests?.dar_value}`;
      }

      const newReport = {
        id: reportId,
        equipmentId: equipTag,
        equipmentName: equipName,
        date: date || netaData?.site_info?.test_date || new Date().toISOString().split('T')[0],
        inspector: netaData?.site_info?.fse_name || 'Kỹ sư FSE',
        status: calculatedStatus,
        notes: reportNotes,
        factory: netaData?.site_info?.project_name || 'Nhà máy TEV Phase 2',
        type: equipType,
        testStandard: testStandard,
        details: {
          netaData,
          analysisResult
        }
      };

      setAllReports(prev => [newReport, ...prev]);

      // Update equipment status if found in allEquipment
      setAllEquipment(prev => prev.map(eq => {
        if (eq.id === newReport.equipmentId) {
          return {
            ...eq,
            lastCheck: newReport.date,
            status: calculatedStatus
          };
        }
        return eq;
      }));

      // Persist to firestore if authenticated
      if (auth.currentUser) {
        try {
          await setDoc(doc(db, 'reports', reportId), newReport);
        } catch (err) {
          console.warn('Could not save to firestore:', err);
        }
      }

      // Auto-sync field test result directly to Google Sheet "TEV Service Flatform" & the corresponding Equipment Sheet
      try {
        const tokens = localStorage.getItem('google_tokens');
        const spreadsheetId = localStorage.getItem('tev_spreadsheet_id') || '';
        const headers: Record<string, string> = { 'Content-Type': 'application/json' };
        if (tokens) headers['Authorization'] = `Bearer ${tokens}`;
        if (spreadsheetId) headers['x-spreadsheet-id'] = spreadsheetId;

        let targetRecordType = 'transformer';
        if (isLvBreaker || equipType.includes('Breaker') || equipType.includes('máy cắt')) {
          targetRecordType = 'breaker';
        } else if (isMotor || isDcMotor || equipType.includes('Motor') || equipType.includes('động cơ')) {
          targetRecordType = 'motor';
        } else if (isGenerator || equipType.includes('Generator') || equipType.includes('máy phát')) {
          targetRecordType = 'generator';
        } else if (isSwitchgear || equipType.includes('Switchgear') || equipType.includes('tủ điện')) {
          targetRecordType = 'switchgear';
        } else if (isCable || isCableLv || equipType.includes('Cable') || equipType.includes('cáp')) {
          targetRecordType = 'cable';
        } else if (isBatteryVrla || isBatteryFlooded || isBess || equipType.includes('Battery') || equipType.includes('pin') || equipType.includes('UPS')) {
          targetRecordType = 'battery';
        } else if (equipType.includes('Relay') || equipType.includes('rơ le')) {
          targetRecordType = 'relay';
        }

        // 1. Sync to Equipment-Specific Sheet (Máy cắt, Động cơ, Máy phát, Tủ điện, Máy biến áp, Cáp điện, Pin & UPS, Rơ le)
        await fetch('/api/sheets/sync-record', {
          method: 'POST',
          headers,
          body: JSON.stringify({
            recordType: targetRecordType,
            spreadsheetId,
            data: {
              woId: reportId,
              equipmentId: equipTag,
              equipmentName: equipName,
              customer: customerName || 'Khách hàng hiện trường',
              location: newReport.factory || siteName || 'Hiện trường',
              assessment: `${analysisResult?.overall_status || 'Hoàn thành'} - ${analysisResult?.status_reason || 'Đạt tiêu chuẩn'}`,
              testDate: newReport.date,
              inspector: newReport.inspector || auth.currentUser?.email || 'Kỹ sư FSE',
              // Breaker
              contactResistanceR: netaData?.electrical_tests?.contact_pole_resistance_micro_ohms?.pole_a || '',
              contactResistanceY: netaData?.electrical_tests?.contact_pole_resistance_micro_ohms?.pole_b || '',
              contactResistanceB: netaData?.electrical_tests?.contact_pole_resistance_micro_ohms?.pole_c || '',
              insulationIrMo: netaData?.electrical_tests?.insulation_resistance_megohms?.pole_a_to_ground || '',
              controlCircuitIrMo: netaData?.electrical_tests?.control_wiring_ir_megohms || '',
              // Motor
              motorIrMo: netaData?.electrical_tests?.insulation_resistance_1min_megohms || '',
              motorPi: netaData?.electrical_tests?.polarization_index_pi || '',
              vibrationDeMmS: netaData?.mechanical_tests?.vibration_de_mm_s || '',
              vibrationNdeMmS: netaData?.mechanical_tests?.vibration_nde_mm_s || '',
              // Generator
              genPowerKva: netaData?.site_info?.rated_power_kva || '',
              genStartSec: netaData?.operational_tests?.cranking_time_seconds || '',
              genOilPressureBar: netaData?.operational_tests?.oil_pressure_psi || '',
              // Switchgear
              swgTevDbmv: netaData?.electrical_tests?.online_partial_discharge_tev_db || '',
              swgUltrasonicDbuv: netaData?.electrical_tests?.airborne_ultrasound_dbuv || '',
              // Transformer
              trfPowerKva: netaData?.site_info?.rated_mva ? Number(netaData.site_info.rated_mva) * 1000 : '',
              trfPriVoltageKv: netaData?.site_info?.primary_voltage_kv || '',
              trfSecVoltageV: netaData?.site_info?.secondary_voltage_v || '',
              trfOilTempC: netaData?.site_info?.top_oil_temperature_c || '',
              trfDgaH2: netaData?.electrical_tests?.oil_sample_tests_astm_d923?.dga_h2_ppm || '',
              trfDgaCh4: netaData?.electrical_tests?.oil_sample_tests_astm_d923?.dga_ch4_ppm || '',
              trfDgaC2h2: netaData?.electrical_tests?.oil_sample_tests_astm_d923?.dga_c2h2_ppm || '',
              trfBreakdownKv: netaData?.electrical_tests?.oil_sample_tests_astm_d923?.dielectric_breakdown_d1816_kv || '',
              trfMoisturePpm: netaData?.electrical_tests?.oil_sample_tests_astm_d923?.water_content_d1533_ppm || '',
              // Cable
              irPhaseAGroundMo: netaData?.electrical_tests?.insulation_resistance_1000v_1min_megohms?.phase_a_to_ground || '',
              irPhaseBGroundMo: netaData?.electrical_tests?.insulation_resistance_1000v_1min_megohms?.phase_b_to_ground || '',
              irPhaseCGroundMo: netaData?.electrical_tests?.insulation_resistance_1000v_1min_megohms?.phase_c_to_ground || '',
              // Battery
              totalVoltageV: netaData?.electrical_tests?.battery_bank_voltage_v || '',
              floatCurrentA: netaData?.electrical_tests?.float_current_ma ? (Number(netaData.electrical_tests.float_current_ma) / 1000).toFixed(2) : ''
            }
          })
        });

        // 2. Also sync to TEV Service Flatform (Work Order 23 columns)
        await fetch('/api/sheets/sync-record', {
          method: 'POST',
          headers,
          body: JSON.stringify({
            recordType: 'workOrder',
            spreadsheetId,
            data: {
              id: reportId,
              workPermitId: `WP-${Date.now().toString().slice(-6)}`,
              title: `Kiểm tra ${equipName} (${testStandard})`,
              description: reportNotes,
              equipmentId: equipTag,
              failureCode: calculatedStatus === 'critical' ? 'ELEC-01' : calculatedStatus === 'warning' ? 'THERM-01' : 'PM-ROUTINE',
              customer: customerName || 'Khách hàng hiện trường',
              type: 'Kiểm định thử nghiệm (Testing)',
              isUnplanned: false,
              priority: calculatedStatus === 'critical' ? 'urgent' : calculatedStatus === 'warning' ? 'high' : 'medium',
              status: 'Hoàn thành',
              assignedTo: newReport.inspector || auth.currentUser?.email || 'Kỹ sư FSE',
              responsibleApprove: 'FSE Lead',
              responsibleDo: 'FSE Onsite',
              blockingRequired: false,
              dueDate: newReport.date,
              usedMaterials: '',
              createdAt: new Date().toISOString(),
              updatedAt: new Date().toISOString(),
              downtimeStart: '',
              repairStart: newReport.date,
              repairEnd: newReport.date,
              restartTime: ''
            }
          })
        });
      } catch (sheetSyncErr) {
        console.warn('Google Sheet auto-sync error for NETA report:', sheetSyncErr);
      }

      alert(`✅ Đã lưu thành công biên bản kiểm tra ${reportId} cho thiết bị ${newReport.equipmentId} vào Kho Báo Cáo CMMS và đồng bộ lên file Google Sheet "TEV Service Flatform"!`);
    } catch (err: any) {
      alert('Lỗi lưu biên bản: ' + err.message);
    }
  };

  // Handlers for FseFieldDataEntryStation with accurate version timestamps
  const handleSaveCustomerFromField = async (customer: CustomerItem) => {
    const nowIso = new Date().toISOString();
    const customerWithTimestamp: CustomerItem = {
      ...customer,
      createdAt: customer.createdAt || nowIso,
      updatedAt: nowIso
    };

    setCustomers(prev => {
      const idx = prev.findIndex(c => c.id === customerWithTimestamp.id);
      if (idx >= 0) {
        const next = [...prev];
        next[idx] = customerWithTimestamp;
        return next;
      }
      return [customerWithTimestamp, ...prev];
    });

    if (auth.currentUser) {
      try {
        await setDoc(doc(db, 'customers', customerWithTimestamp.id), customerWithTimestamp, { merge: true });
      } catch (e) {
        console.warn('Firestore customer save error:', e);
      }
    }

    try {
      const tokens = localStorage.getItem('google_tokens');
      const spreadsheetId = localStorage.getItem('tev_spreadsheet_id') || '';
      const headers: Record<string, string> = { 'Content-Type': 'application/json' };
      if (tokens) headers['Authorization'] = `Bearer ${tokens}`;
      if (spreadsheetId) headers['x-spreadsheet-id'] = spreadsheetId;

      await fetch('/api/sheets/sync-record', {
        method: 'POST',
        headers,
        body: JSON.stringify({
          recordType: 'customer',
          spreadsheetId,
          data: customerWithTimestamp
        })
      });
    } catch (e) {
      console.warn('Sheet customer sync error:', e);
    }
  };

  const handleSaveEquipmentFromField = async (equipment: EquipmentItem) => {
    const nowIso = new Date().toISOString();
    const equipmentWithTimestamp: EquipmentItem = {
      ...equipment,
      createdAt: equipment.createdAt || nowIso,
      updatedAt: nowIso
    };

    setAllEquipment(prev => {
      const idx = prev.findIndex(e => e.id === equipmentWithTimestamp.id);
      if (idx >= 0) {
        const next = [...prev];
        next[idx] = equipmentWithTimestamp;
        return next;
      }
      return [equipmentWithTimestamp, ...prev];
    });

    if (auth.currentUser) {
      try {
        await setDoc(doc(db, 'equipment', equipmentWithTimestamp.id), equipmentWithTimestamp, { merge: true });
      } catch (e) {
        console.warn('Firestore equipment save error:', e);
      }
    }

    try {
      const tokens = localStorage.getItem('google_tokens');
      const spreadsheetId = localStorage.getItem('tev_spreadsheet_id') || '';
      const headers: Record<string, string> = { 'Content-Type': 'application/json' };
      if (tokens) headers['Authorization'] = `Bearer ${tokens}`;
      if (spreadsheetId) headers['x-spreadsheet-id'] = spreadsheetId;

      await fetch('/api/sheets/sync-record', {
        method: 'POST',
        headers,
        body: JSON.stringify({
          recordType: 'equipment',
          spreadsheetId,
          data: equipmentWithTimestamp
        })
      });
    } catch (e) {
      console.warn('Sheet equipment sync error:', e);
    }
  };

  const handleSaveWorkOrderFromField = async (wo: WorkOrderItem, syncToSheetsImmediately: boolean = true) => {
    const nowIso = new Date().toISOString();
    const woWithTimestamp: WorkOrderItem = {
      ...wo,
      createdAt: wo.createdAt || nowIso,
      updatedAt: nowIso
    };

    setWorkOrders(prev => {
      const idx = prev.findIndex(w => w.id === woWithTimestamp.id);
      if (idx >= 0) {
        const next = [...prev];
        next[idx] = woWithTimestamp;
        return next;
      }
      return [woWithTimestamp, ...prev];
    });

    if (auth.currentUser) {
      try {
        await setDoc(doc(db, 'workOrders', woWithTimestamp.id), woWithTimestamp, { merge: true });
      } catch (e) {
        console.warn('Firestore WO save error:', e);
      }
    }

    if (syncToSheetsImmediately) {
      try {
        const tokens = localStorage.getItem('google_tokens');
        const spreadsheetId = localStorage.getItem('tev_spreadsheet_id') || '';
        const headers: Record<string, string> = { 'Content-Type': 'application/json' };
        if (tokens) headers['Authorization'] = `Bearer ${tokens}`;
        if (spreadsheetId) headers['x-spreadsheet-id'] = spreadsheetId;

        await fetch('/api/sheets/sync-record', {
          method: 'POST',
          headers,
          body: JSON.stringify({
            recordType: 'workOrder',
            spreadsheetId,
            data: woWithTimestamp
          })
        });
      } catch (e) {
        console.warn('Sheet workOrder sync error:', e);
      }
    }
  };

  // Deep Analysis State
  const [dgaData, setDgaData] = useState({
    h2: '',
    ch4: '',
    c2h6: '',
    c2h4: '',
    c2h2: '',
    co: '',
    co2: '',
    o2: '',
    n2: '',
    moisture: '',
    bdStrength: '',
    acidity: '',
    ffa: '',
    estDp: '',
    age: '20',
    loadFactor: '50'
  });
  const [dgaAnalysisResult, setDgaAnalysisResult] = useState<any>(null);
  const [dgaViewMode, setDgaViewMode] = useState<'dashboard' | 'both' | 'trend' | 'diagnostic'>('dashboard');

  const handleLoadDgaSampleToForm = (sample: any) => {
    setDgaData(prev => ({
      ...prev,
      h2: sample.h2 !== undefined ? String(sample.h2) : prev.h2,
      ch4: sample.ch4 !== undefined ? String(sample.ch4) : prev.ch4,
      c2h6: sample.c2h6 !== undefined ? String(sample.c2h6) : prev.c2h6,
      c2h4: sample.c2h4 !== undefined ? String(sample.c2h4) : prev.c2h4,
      c2h2: sample.c2h2 !== undefined ? String(sample.c2h2) : prev.c2h2,
      co: sample.co !== undefined ? String(sample.co) : prev.co,
      co2: sample.co2 !== undefined ? String(sample.co2) : prev.co2
    }));
  };

  // Last recorded DGA values from historical database for trend comparison
  const lastDgaRecord = useMemo(() => {
    return getLastRecordedDgaValues(equipmentCode, allReports);
  }, [equipmentCode, allReports]);

  const handleLoadLastRecordedDga = () => {
    setDgaData({
      h2: String(lastDgaRecord.gases.h2 ?? '68'),
      ch4: String(lastDgaRecord.gases.ch4 ?? '82'),
      c2h6: String(lastDgaRecord.gases.c2h6 ?? '48'),
      c2h4: String(lastDgaRecord.gases.c2h4 ?? '35'),
      c2h2: String(lastDgaRecord.gases.c2h2 ?? '0.7'),
      co: String(lastDgaRecord.gases.co ?? '420'),
      co2: String(lastDgaRecord.gases.co2 ?? '3250'),
      o2: String(lastDgaRecord.gases.o2 ?? '520'),
      n2: String(lastDgaRecord.gases.n2 ?? '44000'),
      moisture: String(lastDgaRecord.otherParams.moisture ?? '14'),
      bdStrength: String(lastDgaRecord.otherParams.bdStrength ?? '52'),
      acidity: String(lastDgaRecord.otherParams.acidity ?? '0.06'),
      ffa: String(lastDgaRecord.otherParams.ffa ?? '1.5'),
      estDp: String(lastDgaRecord.otherParams.estDp ?? '680'),
      age: String(lastDgaRecord.otherParams.age ?? '20'),
      loadFactor: String(lastDgaRecord.otherParams.loadFactor ?? '65'),
    });
  };

  const handleLoadElevatedDgaSample = () => {
    setDgaData({
      h2: '125',
      ch4: '148',
      c2h6: '75',
      c2h4: '62',
      c2h2: '2.6',
      co: '520',
      co2: '3850',
      o2: '490',
      n2: '43000',
      moisture: '24',
      bdStrength: '35',
      acidity: '0.15',
      ffa: '2.9',
      estDp: '440',
      age: '20',
      loadFactor: '85'
    });
  };

  const handleLoadDecreasingDgaSample = () => {
    setDgaData({
      h2: '22',
      ch4: '28',
      c2h6: '16',
      c2h4: '11',
      c2h2: '0.1',
      co: '195',
      co2: '1850',
      o2: '560',
      n2: '45500',
      moisture: '8',
      bdStrength: '68',
      acidity: '0.02',
      ffa: '0.7',
      estDp: '780',
      age: '20',
      loadFactor: '50'
    });
  };

  const [deepAnalysisSubTab, setDeepAnalysisSubTab] = useState<'dga' | 'pv-cell' | 'wind-turbine'>('dga');

  const [pvCellData, setPvCellData] = useState({
    temp: '45',
    irradiance: '800',
    voc: '40',
    isc: '9',
    ff: '0.75',
    efficiency: '18'
  });
  const [pvAnalysisResult, setPvAnalysisResult] = useState<any>(null);

  const [windBladeData, setWindBladeData] = useState({
    vibration: '0.5',
    acoustic: '20',
    rotationSpeed: '15',
    windSpeed: '12',
    visualNotes: ''
  });
  const [windAnalysisResult, setWindAnalysisResult] = useState<any>(null);

  const dynamicSiteData = useMemo(() => {
    const siteMap = new Map();
    
    // Filter allEquipment first if user is a customer with an assigned factory
    const equipmentToProcess = (userRole === 'customer' && userFactory)
      ? allEquipment.filter(eq => eq.factory?.trim().toLowerCase() === userFactory.trim().toLowerCase())
      : allEquipment;
    
    // Helper to detect country & flag based on lat/lng or factory name
    const getCountryInfo = (lat: number, lng: number, factoryName: string) => {
      const lower = factoryName.toLowerCase();
      if (lower.includes('laos') || lower.includes('lào') || lower.includes('nam theun') || lower.includes('xekaman') || lower.includes('vientiane') || lower.includes('nam ngum')) {
        return { country: 'Lào', flag: '🇱🇦' };
      }
      if (lower.includes('cambodia') || lower.includes('campuchia') || lower.includes('phnom penh') || lower.includes('sesan 2') || lower.includes('bavet')) {
        return { country: 'Campuchia', flag: '🇰🇭' };
      }
      if (lower.includes('thái lan') || lower.includes('thailand') || lower.includes('bangkok') || lower.includes('chonburi') || lower.includes('rayong')) {
        return { country: 'Thái Lan', flag: '🇹🇭' };
      }
      if (lower.includes('singapore') || lower.includes('jurong')) {
        return { country: 'Singapore', flag: '🇸🇬' };
      }
      if (lower.includes('indonesia') || lower.includes('jakarta') || lower.includes('cikarang')) {
        return { country: 'Indonesia', flag: '🇮🇩' };
      }
      if (lower.includes('malaysia') || lower.includes('kuala lumpur') || lower.includes('penang')) {
        return { country: 'Malaysia', flag: '🇲🇾' };
      }
      if (lower.includes('nhật') || lower.includes('japan') || lower.includes('tokyo') || lower.includes('osaka')) {
        return { country: 'Nhật Bản', flag: '🇯🇵' };
      }
      if (lower.includes('đức') || lower.includes('germany') || lower.includes('frankfurt') || lower.includes('munich')) {
        return { country: 'Đức', flag: '🇩🇪' };
      }
      if (lower.includes('mỹ') || lower.includes('usa') || lower.includes('california') || lower.includes('texas')) {
        return { country: 'Hoa Kỳ', flag: '🇺🇸' };
      }
      if (lower.includes('úc') || lower.includes('australia') || lower.includes('sydney')) {
        return { country: 'Úc', flag: '🇦🇺' };
      }
      if (lat >= 8.0 && lat <= 24.0 && lng >= 102.0 && lng <= 110.0) {
        return { country: 'Việt Nam', flag: '🇻🇳' };
      }
      return { country: 'Quốc tế', flag: '🌐' };
    };

    // Known coordinates for nice map display (Domestic and International sites)
    const knownCoords: Record<string, {lat: number, lng: number, country?: string, flag?: string}> = {
      // VIETNAM SITES
      'Thủy điện Sơn La': { lat: 21.496, lng: 103.995, country: 'Việt Nam', flag: '🇻🇳' },
      'Thủy điện Lai Châu': { lat: 22.140, lng: 102.980, country: 'Việt Nam', flag: '🇻🇳' },
      'Thủy điện Hòa Bình': { lat: 20.808, lng: 105.328, country: 'Việt Nam', flag: '🇻🇳' },
      'Nhiệt điện Phả Lại': { lat: 21.111, lng: 106.315, country: 'Việt Nam', flag: '🇻🇳' },
      'Nhiệt điện Mông Dương': { lat: 21.070, lng: 107.350, country: 'Việt Nam', flag: '🇻🇳' },
      'Nhiệt điện Nghi Sơn': { lat: 19.330, lng: 105.780, country: 'Việt Nam', flag: '🇻🇳' },
      'Thủy điện Bản Vẽ': { lat: 19.320, lng: 104.480, country: 'Việt Nam', flag: '🇻🇳' },
      'Nhiệt điện Vũng Áng': { lat: 18.120, lng: 106.350, country: 'Việt Nam', flag: '🇻🇳' },
      'Thủy điện Quảng Trị': { lat: 16.650, lng: 106.750, country: 'Việt Nam', flag: '🇻🇳' },
      'Thủy điện A Lưới': { lat: 16.250, lng: 107.250, country: 'Việt Nam', flag: '🇻🇳' },
      'Thủy điện Sông Tranh 2': { lat: 15.350, lng: 108.150, country: 'Việt Nam', flag: '🇻🇳' },
      'Thủy điện Ialy': { lat: 14.220, lng: 107.750, country: 'Việt Nam', flag: '🇻🇳' },
      'Thủy điện Sê San 4': { lat: 13.950, lng: 107.550, country: 'Việt Nam', flag: '🇻🇳' },
      'Điện gió Phương Mai': { lat: 13.850, lng: 109.250, country: 'Việt Nam', flag: '🇻🇳' },
      'Điện mặt trời Trung Nam': { lat: 11.650, lng: 108.950, country: 'Việt Nam', flag: '🇻🇳' },
      'Nhiệt điện Vĩnh Tân': { lat: 11.320, lng: 108.850, country: 'Việt Nam', flag: '🇻🇳' },
      'Thủy điện Trị An': { lat: 11.120, lng: 107.020, country: 'Việt Nam', flag: '🇻🇳' },
      'Nhiệt điện Phú Mỹ': { lat: 10.580, lng: 107.050, country: 'Việt Nam', flag: '🇻🇳' },
      'Điện gió Bạc Liêu': { lat: 9.250, lng: 105.820, country: 'Việt Nam', flag: '🇻🇳' },
      'Nhiệt điện Cà Mau': { lat: 9.180, lng: 104.920, country: 'Việt Nam', flag: '🇻🇳' },
      'Nhà máy Điện Gió Mũi Né': { lat: 10.933, lng: 108.287, country: 'Việt Nam', flag: '🇻🇳' },
      'Tòa nhà TEV Office Tower': { lat: 10.776, lng: 106.700, country: 'Việt Nam', flag: '🇻🇳' },
      'Bệnh Viện TEV General Hospital': { lat: 10.755, lng: 106.665, country: 'Việt Nam', flag: '🇻🇳' },
      'Trung Tâm Dữ Liệu TEV Tier-III': { lat: 10.840, lng: 106.810, country: 'Việt Nam', flag: '🇻🇳' },
      'Khu Công Nghiệp VSIP 1': { lat: 10.925, lng: 106.710, country: 'Việt Nam', flag: '🇻🇳' },
      'Nhà máy Sản xuất TEV Tân Uyên': { lat: 11.080, lng: 106.790, country: 'Việt Nam', flag: '🇻🇳' },
      'Trạm Sạc Cao Tốc Long Thành': { lat: 10.740, lng: 106.980, country: 'Việt Nam', flag: '🇻🇳' },

      // INTERNATIONAL SITES (Lào, Campuchia, Thái Lan, Singapore, Nhật Bản, Đức, Mỹ,...)
      'Thủy điện Nam Theun 2 (Lào)': { lat: 17.830, lng: 104.970, country: 'Lào', flag: '🇱🇦' },
      'Thủy điện Xekaman 1 (Lào)': { lat: 14.980, lng: 107.150, country: 'Lào', flag: '🇱🇦' },
      'Thủy điện Xekaman 3 (Lào)': { lat: 15.220, lng: 107.450, country: 'Lào', flag: '🇱🇦' },
      'Trạm biến áp Vientiane (Lào)': { lat: 17.970, lng: 102.630, country: 'Lào', flag: '🇱🇦' },
      'Thủy điện Lower Sesan 2 (Campuchia)': { lat: 13.550, lng: 106.270, country: 'Campuchia', flag: '🇰🇭' },
      'Trạm biến áp Phnom Penh (Campuchia)': { lat: 11.550, lng: 104.920, country: 'Campuchia', flag: '🇰🇭' },
      'Điện mặt trời Bavet (Campuchia)': { lat: 11.080, lng: 106.140, country: 'Campuchia', flag: '🇰🇭' },
      'Bangkok Energy & Data Hub (Thái Lan)': { lat: 13.750, lng: 100.520, country: 'Thái Lan', flag: '🇹🇭' },
      'Rayong Petrochemical Complex (Thái Lan)': { lat: 12.680, lng: 101.280, country: 'Thái Lan', flag: '🇹🇭' },
      'Singapore Jurong Island Substation': { lat: 1.270, lng: 103.710, country: 'Singapore', flag: '🇸🇬' },
      'Singapore Tier-IV Data Center': { lat: 1.350, lng: 103.820, country: 'Singapore', flag: '🇸🇬' },
      'Kuala Lumpur High-Tech Plant (Malaysia)': { lat: 3.140, lng: 101.690, country: 'Malaysia', flag: '🇲🇾' },
      'Jakarta Cikarang Industrial Power (Indonesia)': { lat: -6.300, lng: 107.160, country: 'Indonesia', flag: '🇮🇩' },
      'Tokyo Advanced Energy Plant (Nhật Bản)': { lat: 35.680, lng: 139.760, country: 'Nhật Bản', flag: '🇯🇵' },
      'Frankfurt Mission Critical Substation (Đức)': { lat: 50.110, lng: 8.680, country: 'Đức', flag: '🇩🇪' },
      'California Solar & BESS Hub (Hoa Kỳ)': { lat: 34.050, lng: -118.250, country: 'Hoa Kỳ', flag: '🇺🇸' },
      'Texas Wind & Substation (Hoa Kỳ)': { lat: 31.960, lng: -99.900, country: 'Hoa Kỳ', flag: '🇺🇸' },
      'Sydney Renewable Power Site (Úc)': { lat: -33.860, lng: 151.200, country: 'Úc', flag: '🇦🇺' }
    };

    equipmentToProcess.forEach(eq => {
      const factory = eq.factory?.trim() || 'Chưa xác định';
      const customer = eq.customer?.trim() || 'Chưa xác định';
      const id = factory.toLowerCase().replace(/\s+/g, '-');
      
      if (!siteMap.has(id)) {
        let lat: number | null = null;
        let lng: number | null = null;

        // 1. Direct coordinates on equipment if present
        if (typeof eq.lat === 'number' && typeof eq.lng === 'number' && !isNaN(eq.lat) && !isNaN(eq.lng)) {
          lat = eq.lat;
          lng = eq.lng;
        } else if (typeof eq.latitude === 'number' && typeof eq.longitude === 'number' && !isNaN(eq.latitude) && !isNaN(eq.longitude)) {
          lat = eq.latitude;
          lng = eq.longitude;
        } else if (eq.coordinates && typeof eq.coordinates.lat === 'number' && typeof eq.coordinates.lng === 'number') {
          lat = eq.coordinates.lat;
          lng = eq.coordinates.lng;
        } else if (typeof eq.locationCoord === 'string' && eq.locationCoord.includes(',')) {
          const parts = eq.locationCoord.split(',').map((p: string) => parseFloat(p.trim()));
          if (!isNaN(parts[0]) && !isNaN(parts[1])) {
            lat = parts[0];
            lng = parts[1];
          }
        }

        // 2. Case-insensitive coordinate matching against known locations (domestic & global)
        const normalizedFactory = factory.toLowerCase();
        const coordKey = Object.keys(knownCoords).find(k => 
          k.toLowerCase() === normalizedFactory || 
          normalizedFactory.includes(k.toLowerCase()) || 
          k.toLowerCase().includes(normalizedFactory)
        );
        
        if (lat === null && coordKey) {
          lat = knownCoords[coordKey].lat;
          lng = knownCoords[coordKey].lng;
        }

        // 3. Fallback: Smart regional coordinate assignment if site name mentions country
        if (lat === null || lng === null) {
          const hash = id.split('').reduce((acc, char) => acc + char.charCodeAt(0), 0);
          if (normalizedFactory.includes('laos') || normalizedFactory.includes('lào')) {
            lat = 17.5 + (hash % 10) * 0.1;
            lng = 104.5 + (hash % 10) * 0.1;
          } else if (normalizedFactory.includes('cambodia') || normalizedFactory.includes('campuchia')) {
            lat = 11.5 + (hash % 10) * 0.1;
            lng = 104.5 + (hash % 10) * 0.1;
          } else if (normalizedFactory.includes('thái lan') || normalizedFactory.includes('thailand')) {
            lat = 13.5 + (hash % 10) * 0.1;
            lng = 100.5 + (hash % 10) * 0.1;
          } else if (normalizedFactory.includes('singapore')) {
            lat = 1.30 + (hash % 10) * 0.01;
            lng = 103.80 + (hash % 10) * 0.01;
          } else if (normalizedFactory.includes('japan') || normalizedFactory.includes('nhật')) {
            lat = 35.5 + (hash % 10) * 0.1;
            lng = 139.5 + (hash % 10) * 0.1;
          } else if (normalizedFactory.includes('đức') || normalizedFactory.includes('germany')) {
            lat = 50.0 + (hash % 10) * 0.1;
            lng = 9.0 + (hash % 10) * 0.1;
          } else if (normalizedFactory.includes('mỹ') || normalizedFactory.includes('usa')) {
            lat = 34.0 + (hash % 10) * 0.1;
            lng = -118.0 + (hash % 10) * 0.1;
          } else {
            // General VN pseudo-random coordinate
            lat = 10 + (hash % 10) + (hash % 100) / 100;
            lng = 105 + (hash % 5) + (hash % 100) / 100;
          }
        }

        const countryInfo = getCountryInfo(lat, lng, factory);

        siteMap.set(id, {
          id,
          name: factory,
          customer: customer,
          lat,
          lng,
          country: countryInfo.country,
          flag: countryInfo.flag,
          count: 0,
          healthy: 0,
          warning: 0,
          critical: 0,
          status: 'healthy'
        });
      }
      
      const site = siteMap.get(id);
      site.count += 1;
      if (eq.status === 'healthy') site.healthy += 1;
      if (eq.status === 'warning') site.warning += 1;
      if (eq.status === 'critical') site.critical += 1;
      
      if (site.critical > 0) site.status = 'critical';
      else if (site.warning > 0 && site.status !== 'critical') site.status = 'warning';
    });
    
    return Array.from(siteMap.values());
  }, [allEquipment, userRole, userFactory]);

  const uniqueCustomers = useMemo(() => {
    const set = new Set<string>();
    dynamicSiteData.forEach(site => { if (site.customer) set.add(site.customer); });
    customers.forEach(c => { if (c.name) set.add(c.name); });
    workOrders.forEach(wo => {
      if (wo.customer) set.add(wo.customer);
      if (wo.customerId && !wo.customerId.startsWith('KH-')) set.add(wo.customerId);
    });
    return Array.from(set);
  }, [dynamicSiteData, customers, workOrders]);

  const filteredSiteData = dynamicSiteData.filter(site => {
    const matchCustomer = selectedCustomer === 'all' || site.customer === selectedCustomer;
    const matchFactory = selectedFactory === 'all' || site.id === selectedFactory;
    return matchCustomer && matchFactory;
  });

  const dynamicTrendData = useMemo(() => {
    const eq = allEquipment.find(e => e.id === trendingEquipment);
    if (!eq) return { temp: [], tev: [], dga: [], winding_temp: [], ir: [], winding_res: [], pd: [], elcid: [], pi: [] };

    const reportsForEq = allReports.filter(r => r.equipmentId === trendingEquipment);
    const sortedReports = [...reportsForEq].sort((a, b) => {
      const [d1, m1, y1] = a.date.split('/');
      const [d2, m2, y2] = b.date.split('/');
      return new Date(`${y1}-${m1}-${d1}`).getTime() - new Date(`${y2}-${m2}-${d2}`).getTime();
    });

    const parseVal = (val: any) => {
      if (typeof val === 'string') return parseFloat(val.replace(',', '.'));
      return val;
    };

    const tempTrend = sortedReports.map(r => ({
      time: r.date,
      temp: parseVal(r.measurements?.oilTemp || r.measurements?.statorTemp) || 0
    })).filter(d => d.temp > 0);

    const windingTempTrend = sortedReports.map(r => ({
      time: r.date,
      temp: parseVal(r.measurements?.windingTemp || r.measurements?.bearingTemp) || 0
    })).filter(d => d.temp > 0);

    const tevTrend = sortedReports.map(r => ({
      time: r.date,
      tev: parseVal(r.measurements?.tev) || 0,
      ultrasonic: parseVal(r.measurements?.ultrasonic) || 0
    })).filter(d => d.tev > 0 || d.ultrasonic > 0);

    const dgaTrend = sortedReports.map(r => {
      return {
        time: r.date,
        h2: parseVal(r.measurements?.h2) || 0,
        o2: parseVal(r.measurements?.o2) || 0,
        n2: parseVal(r.measurements?.n2) || 0,
        ch4: parseVal(r.measurements?.ch4) || 0,
        co: parseVal(r.measurements?.co) || 0,
        co2: parseVal(r.measurements?.co2) || 0,
        c2h4: parseVal(r.measurements?.c2h4) || 0,
        c2h6: parseVal(r.measurements?.c2h6) || 0,
        c2h2: parseVal(r.measurements?.c2h2) || 0
      };
    }).filter(d => d.h2 > 0 || d.ch4 > 0 || d.c2h2 > 0 || d.co > 0);

    const irTrend = sortedReports.map(r => ({
      time: r.date,
      ir: parseVal(r.measurements?.irHighLow || r.measurements?.irR) || 0
    })).filter(d => d.ir > 0);

    const windingResTrend = sortedReports.map(r => ({
      time: r.date,
      res: parseVal(r.measurements?.contactRes) || 0
    })).filter(d => d.res > 0);

    const pdTrend = sortedReports.map(r => ({
      time: r.date,
      pd: parseVal(r.measurements?.pdR) || 0
    })).filter(d => d.pd > 0);

    const elcidTrend = sortedReports.map(r => ({
      time: r.date,
      elcid: parseVal(r.measurements?.elcidR) || 0
    })).filter(d => d.elcid > 0);

    const piTrend = sortedReports.map(r => ({
      time: r.date,
      pi: parseVal(r.measurements?.piR) || 0
    })).filter(d => d.pi > 0);

    return { 
      temp: tempTrend, 
      tev: tevTrend, 
      dga: dgaTrend,
      winding_temp: windingTempTrend,
      ir: irTrend,
      winding_res: windingResTrend,
      pd: pdTrend,
      elcid: elcidTrend,
      pi: piTrend
    };
  }, [trendingEquipment, allEquipment, allReports, userRole, userFactory]);

  const dynamicParamHistoryData = useMemo(() => {
    if (!selectedRiskDetail || !selectedParamHistory || !allReports) return null;
    
    const paramKeyMap: Record<string, string> = {
      'Tuổi thiết bị': 'age',
      'Hệ số tải': 'dutyFactor',
      'Nhiệt độ dầu': 'oilTemp',
      'Nhiệt độ cuộn dây': 'windingTemp',
      'Điện trở cách điện (Cao-Hạ)': 'irHighLow',
      'Điện trở cách điện (Cao-Vỏ)': 'irHighEarth',
      'Tình trạng rò rỉ dầu': 'oilLeak',
      'DGA': 'dga',
      'Độ bền điện môi': 'dielectricStrength',
      'Furan': 'furan',
      'Độ ẩm dầu': 'oilMoisture',
      'Độ rung': 'vibration',
      'Mất cân bằng điện áp': 'voltageImbalance',
      'Nhiệt độ Stator': 'statorTemp',
      'Nhiệt độ vòng bi': 'bearingTemp',
      'Nhiệt độ tiếp xúc': 'thermography',
      'Điện trở tiếp xúc': 'contactRes',
      'TEV': 'tev',
      'Số xung TEV': 'tevPulses',
      'Ultrasonic': 'ultrasonic',
      'Độ ẩm môi trường': 'humidity',
      'Áp suất SF6': 'sf6Pressure'
    };

    const key = paramKeyMap[selectedParamHistory];
    if (!key) return null;

    const reportsForEq = allReports.filter(r => r.equipmentId === selectedRiskDetail);
    
    // Sort by date ascending for chart
    const sortedReports = [...reportsForEq].sort((a, b) => {
      const [d1, m1, y1] = a.date.split('/');
      const [d2, m2, y2] = b.date.split('/');
      return new Date(`${y1}-${m1}-${d1}`).getTime() - new Date(`${y2}-${m2}-${d2}`).getTime();
    });

    const dataPoints = sortedReports.map(r => {
      let val = r.measurements?.[key];
      if (typeof val === 'string') {
        val = parseFloat(val.replace(',', '.'));
      }
      return {
        date: r.date,
        value: isNaN(val) ? null : val
      };
    }).filter(d => d.value !== null);

    return dataPoints.length > 0 ? dataPoints : null;
  }, [allReports, selectedRiskDetail, selectedParamHistory]);

  const selectedFactoryName = selectedFactory === 'all' 
    ? 'all' 
    : dynamicSiteData.find(s => s.id === selectedFactory)?.name;

  const filteredEquipment = allEquipment.filter(eq => {
    // If user is a customer and has an assigned factory, restrict to that factory
    if (userRole === 'customer' && userFactory) {
      return eq.factory?.trim().toLowerCase() === userFactory.trim().toLowerCase();
    }
    
    const matchCustomer = selectedCustomer === 'all' || eq.customer === selectedCustomer;
    const matchFactory = selectedFactoryName === 'all' || eq.factory === selectedFactoryName;
    return matchCustomer && matchFactory;
  });

  const filteredKpiData = {
    total: filteredEquipment.length,
    healthy: filteredEquipment.filter(e => e.status === 'healthy').length,
    warning: filteredEquipment.filter(e => e.status === 'warning').length,
    critical: filteredEquipment.filter(e => e.status === 'critical').length,
  };

  const displayedEquipment = filteredEquipment.filter(eq => {
    const matchStatus = statusFilter === 'all' || eq.status === statusFilter;
    const matchName = filterName === '' || eq.name.toLowerCase().includes(filterName.toLowerCase()) || eq.id.toLowerCase().includes(filterName.toLowerCase());
    const matchLocation = filterLocation === '' || eq.factory.toLowerCase().includes(filterLocation.toLowerCase()) || eq.location.toLowerCase().includes(filterLocation.toLowerCase());
    const matchType = filterType === 'all' || eq.type === filterType;
    const matchHealthMin = filterHealthMin === '' || eq.health >= parseInt(filterHealthMin);
    const matchHealthMax = filterHealthMax === '' || eq.health <= parseInt(filterHealthMax);
    return matchStatus && matchName && matchLocation && matchType && matchHealthMin && matchHealthMax;
  });

  const displayedReports = allReports.filter(report => {
    // If user is a customer and has an assigned factory, restrict to that factory
    if (userRole === 'customer' && userFactory) {
      if (report.factory?.trim().toLowerCase() !== userFactory.trim().toLowerCase()) return false;
    }
    
    const matchSearch = reportFilterSearch === '' || 
      report.id.toLowerCase().includes(reportFilterSearch.toLowerCase()) || 
      report.equipmentId.toLowerCase().includes(reportFilterSearch.toLowerCase()) ||
      report.equipmentName.toLowerCase().includes(reportFilterSearch.toLowerCase());
    const matchFactory = reportFilterFactory === 'all' || report.factory === reportFilterFactory;
    const matchType = reportFilterType === 'all' || report.type === reportFilterType;
    const matchStatus = reportFilterStatus === 'all' || report.status === reportFilterStatus;
    
    let matchDate = true;
    if (reportFilterStartDate || reportFilterEndDate) {
      const [d, m, y] = report.date.split('/');
      const reportDate = new Date(`${y}-${m}-${d}`);
      if (reportFilterStartDate) {
        matchDate = matchDate && reportDate >= new Date(reportFilterStartDate);
      }
      if (reportFilterEndDate) {
        matchDate = matchDate && reportDate <= new Date(reportFilterEndDate);
      }
    }
    
    return matchSearch && matchFactory && matchType && matchStatus && matchDate;
  });

  const totalReportPages = Math.ceil(displayedReports.length / reportsPerPage);
  const currentReports = displayedReports.slice(
    (reportCurrentPage - 1) * reportsPerPage,
    reportCurrentPage * reportsPerPage
  );

  // Pagination
  const itemsPerPage = 20;
  const totalPages = Math.ceil(displayedEquipment.length / itemsPerPage);
  const paginatedEquipment = displayedEquipment.slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage);

  const topRiskCount = Math.max(5, Math.ceil(filteredEquipment.length * 0.1));
  const filteredTopRiskEquipment = filteredEquipment
    .filter(eq => eq.status === 'critical' || eq.status === 'warning')
    .sort((a, b) => a.health - b.health)
    .slice(0, topRiskCount);

  const filteredStatusDistribution = [
    { name: 'Khỏe mạnh', value: filteredKpiData.healthy, color: '#10b981' },
    { name: 'Cảnh báo', value: filteredKpiData.warning, color: '#f59e0b' },
    { name: 'Nguy hiểm', value: filteredKpiData.critical, color: '#f43f5e' },
  ];

  const filteredHealthDistribution = [
    { range: '90-100%', count: filteredEquipment.filter(eq => eq.health >= 90).length },
    { range: '70-89%', count: filteredEquipment.filter(eq => eq.health >= 70 && eq.health < 90).length },
    { range: '50-69%', count: filteredEquipment.filter(eq => eq.health >= 50 && eq.health < 70).length },
    { range: '<50%', count: filteredEquipment.filter(eq => eq.health < 50).length },
  ];

  const averageHealth = filteredEquipment.length > 0
    ? Math.round(filteredEquipment.reduce((sum, eq) => sum + eq.health, 0) / filteredEquipment.length)
    : 0;

  const reliabilityKpis = useMemo(() => {
    const filteredWOs = workOrders.filter(wo => {
      const woCust = (wo.customer || wo.customerId || '').toLowerCase().trim();
      const selCust = selectedCustomer.toLowerCase().trim();
      if (selectedCustomer !== 'all') {
        const match = woCust === selCust || woCust.includes(selCust) || selCust.includes(woCust) || (wo.customerId && wo.customerId.toLowerCase() === selCust);
        if (!match) return false;
      }
      if (userRole === 'customer' && userFactory) {
        const matchUser = wo.customerId === userFactory || (wo.customer && userCustomerName && wo.customer.toLowerCase().includes(userCustomerName.toLowerCase()));
        if (!matchUser) return false;
      }
      return true;
    });

    const isCorrective = (wo: any) => {
      const t = (wo.type || '').toLowerCase();
      return t.includes('corrective') || t.includes('emergency') || t.includes('đột xuất') || t.includes('khẩn cấp') || t.includes('cm') || wo.isUnplanned;
    };
    const isPreventive = (wo: any) => {
      const t = (wo.type || '').toLowerCase();
      return t.includes('preventive') || t.includes('predictive') || t.includes('định kỳ') || t.includes('pm') || t.includes('bảo dưỡng');
    };
    const isCompleted = (wo: any) => {
      const s = (wo.status || '').toLowerCase();
      return s.includes('complete') || s.includes('hoàn thành') || s.includes('đã đóng');
    };

    const completedCorrective = filteredWOs.filter(wo => isCorrective(wo) && isCompleted(wo));

    // 1. MTTR: Total Repair Time / Total No. of Repairs
    const totalRepairTime = completedCorrective.reduce((sum, wo) => sum + (wo.actualTimeSpent || 4), 0);
    const mttr = completedCorrective.length > 0 
      ? (totalRepairTime / completedCorrective.length).toFixed(1) 
      : (filteredWOs.length > 0 ? '3.5' : '0.0');

    // 2. MTBF: Total Operational Hours / Total No. of Failures
    const totalAssets = Math.max(1, filteredEquipment.length);
    const totalOpHours = totalAssets * 30 * 24; 
    const totalFailures = Math.max(1, completedCorrective.length + filteredWOs.filter(isCorrective).length);
    const mtbf = Math.round(totalOpHours / totalFailures);

    // 3. MTTF: Total Hours of Operation / Total Assets
    const mttf = Math.round(totalOpHours / totalAssets);

    // 4. MWT: Average time waiting for maintenance
    const assignedWOs = filteredWOs.filter(wo => !((wo.status || '').toLowerCase().includes('init') || (wo.status || '').toLowerCase().includes('mới')));
    const totalWaitTime = assignedWOs.reduce((sum, wo) => {
      const created = new Date(wo.createdAt || Date.now()).getTime();
      const updated = new Date(wo.updatedAt || Date.now()).getTime();
      return sum + Math.max(0, (updated - created) / (1000 * 60 * 60)); 
    }, 0);
    const mwt = assignedWOs.length > 0 ? (totalWaitTime / assignedWOs.length).toFixed(1) : (filteredWOs.length > 0 ? '1.5' : '0.0');

    const totalBreakdownTime = totalRepairTime || (filteredWOs.filter(isCorrective).length * 4);
    const totalCost = filteredWOs.reduce((sum, wo) => sum + (wo.laborCost || 0) + (wo.partCost || 0), 0) || (filteredWOs.length * 150);
    const backlog = filteredWOs.filter(wo => !isCompleted(wo) && !(wo.status || '').toLowerCase().includes('cancel')).length;
    
    const preventive = filteredWOs.filter(isPreventive);
    const completedPreventive = preventive.filter(isCompleted);
    const compliance = preventive.length > 0 ? Math.round((completedPreventive.length / preventive.length) * 100) : (filteredWOs.length > 0 ? 85 : 100);

    // Maintenance Type Distribution
    const prevCount = filteredWOs.filter(isPreventive).length;
    const corrCount = filteredWOs.filter(isCorrective).length;
    const otherCount = Math.max(0, filteredWOs.length - prevCount - corrCount);

    const typeDist = [
      { name: 'Phòng ngừa (PM)', value: prevCount || (filteredWOs.length === 0 ? 1 : 0), color: '#3b82f6' },
      { name: 'Khắc phục (CM)', value: corrCount || (filteredWOs.length === 0 ? 1 : 0), color: '#f43f5e' },
      { name: 'Khác / Khảo sát', value: otherCount, color: '#94a3b8' }
    ];

    const initiatedCount = filteredWOs.filter(wo => {
      const s = (wo.status || '').toLowerCase();
      return s.includes('init') || s.includes('mới') || s.includes('open') || s.includes('chờ');
    }).length;

    const inProgressCount = filteredWOs.filter(wo => {
      const s = (wo.status || '').toLowerCase();
      return s.includes('progress') || s.includes('đang') || s.includes('assigned');
    }).length;

    const completedCount = filteredWOs.filter(isCompleted).length;

    const criticalCount = filteredWOs.filter(wo => {
      const p = (wo.priority || '').toLowerCase();
      return p === 'critical' || p === 'high' || p.includes('cao') || p.includes('khẩn');
    }).length;

    return { 
      mttr, 
      mtbf, 
      mttf, 
      mwt, 
      totalBreakdownTime, 
      totalCost, 
      backlog, 
      compliance, 
      typeDist,
      totalWOs: filteredWOs.length,
      initiatedCount,
      inProgressCount,
      completedCount,
      criticalCount
    };
  }, [workOrders, filteredEquipment, selectedCustomer, userRole, userFactory, userCustomerName]);

  const dashboardFilteredWorkOrders = useMemo(() => {
    return workOrders.filter(wo => {
      // 1. Role and Customer filtering
      const woCust = (wo.customer || wo.customerId || '').toLowerCase().trim();
      const selCust = selectedCustomer.toLowerCase().trim();
      if (selectedCustomer !== 'all') {
        const match = woCust === selCust || woCust.includes(selCust) || selCust.includes(woCust) || (wo.customerId && wo.customerId.toLowerCase() === selCust);
        if (!match) return false;
      }
      if (userRole === 'customer' && userFactory) {
        const matchUser = wo.customerId === userFactory || (wo.customer && userCustomerName && wo.customer.toLowerCase().includes(userCustomerName.toLowerCase()));
        if (!matchUser) return false;
      }

      // 2. Type/Status filter tabs
      if (dashboardWoFilter === 'corrective') {
        const t = (wo.type || '').toLowerCase();
        if (!t.includes('corrective') && !t.includes('đột xuất') && !t.includes('cm') && !t.includes('khẩn cấp') && !wo.isUnplanned) return false;
      } else if (dashboardWoFilter === 'preventive') {
        const t = (wo.type || '').toLowerCase();
        if (!t.includes('preventive') && !t.includes('predictive') && !t.includes('định kỳ') && !t.includes('pm') && !t.includes('bảo dưỡng')) return false;
      } else if (dashboardWoFilter === 'in-progress') {
        const s = (wo.status || '').toLowerCase();
        if (!s.includes('progress') && !s.includes('đang') && !s.includes('assigned')) return false;
      } else if (dashboardWoFilter === 'initiated') {
        const s = (wo.status || '').toLowerCase();
        if (!s.includes('init') && !s.includes('mới') && !s.includes('open') && !s.includes('chờ')) return false;
      } else if (dashboardWoFilter === 'overdue') {
        if (!wo.dueDate) return false;
        const isOver = new Date(wo.dueDate).getTime() < Date.now();
        const s = (wo.status || '').toLowerCase();
        const isComp = s.includes('complete') || s.includes('hoàn thành');
        if (!isOver || isComp) return false;
      }

      // 3. Search query
      if (dashboardWoSearch.trim()) {
        const q = dashboardWoSearch.toLowerCase().trim();
        const eqStr = Array.isArray(wo.equipmentId) ? wo.equipmentId.join(' ') : String(wo.equipmentId || '');
        const matchSearch = (
          (wo.id || '').toLowerCase().includes(q) ||
          (wo.title || '').toLowerCase().includes(q) ||
          (wo.description || '').toLowerCase().includes(q) ||
          (wo.customer || '').toLowerCase().includes(q) ||
          (wo.assignedTo || '').toLowerCase().includes(q) ||
          eqStr.toLowerCase().includes(q) ||
          (wo.equipmentName || '').toLowerCase().includes(q)
        );
        if (!matchSearch) return false;
      }

      return true;
    });
  }, [workOrders, selectedCustomer, userRole, userFactory, userCustomerName, dashboardWoFilter, dashboardWoSearch]);

  let healthColor = '#10b981'; // emerald-500
  let healthText = 'Khá Tốt';
  let healthTextColor = 'text-emerald-600';
  if (averageHealth < 50) {
    healthColor = '#f43f5e'; // rose-500
    healthText = 'Nguy hiểm';
    healthTextColor = 'text-rose-600';
  } else if (averageHealth < 70) {
    healthColor = '#f59e0b'; // amber-500
    healthText = 'Cảnh báo';
    healthTextColor = 'text-amber-600';
  } else if (averageHealth >= 90) {
    healthText = 'Rất Tốt';
  }

  // Reset form when changing equipment type
  useEffect(() => {
    setFormData({});
    setHealthResult(null);
    setAttachedFiles([]);
  }, [selectedEqType]);

  // Google Auth Effect
  useEffect(() => {
    let retryCount = 0;
    const maxRetries = 3;

    const checkAuthStatus = async () => {
      try {
        const response = await fetch('/api/auth/status', {
          headers: getAuthHeaders(false)
        });
        if (response.ok) {
          const data = await response.json();
          setIsGoogleConnected(data.isAuthenticated);
          retryCount = 0; // Reset on success
        }
      } catch (error: any) {
        // Only log error if we've exhausted retries or it's not a fetch error
        if (error.message === 'Failed to fetch' && retryCount < maxRetries) {
          retryCount++;
          setTimeout(checkAuthStatus, 2000 * retryCount); // Exponential backoff
        } else {
          console.error('Failed to check auth status', error);
        }
      }
    };
    checkAuthStatus();

    // Poll for status every 30 seconds to catch expiration
    const interval = setInterval(checkAuthStatus, 30000);

    const handleMessage = (event: MessageEvent) => {
      if (event.origin !== window.location.origin) {
        return;
      }
      if (event.data?.type === 'OAUTH_AUTH_SUCCESS') {
        if (event.data.tokens) {
          localStorage.setItem('google_tokens', JSON.stringify(event.data.tokens));
        }
        setIsGoogleConnected(true);
      }
    };
    window.addEventListener('message', handleMessage);
    return () => {
      window.removeEventListener('message', handleMessage);
      clearInterval(interval);
    };
  }, []);

  // Firestore connection test
  useEffect(() => {
    const testFirestore = async () => {
      try {
        const { doc, getDocFromCache, getDocFromServer } = await import('firebase/firestore');
        // Try cache first, then server
        try {
          await getDocFromCache(doc(db, '_connection_test_', 'ping'));
        } catch (e) {
          // Ignore cache errors
        }
        await getDocFromServer(doc(db, '_connection_test_', 'ping')).catch(err => {
          if (err.code === 'unavailable') {
            console.warn('Firestore is unavailable (offline mode). This is expected in some network environments.');
          } else if (err.code !== 'not-found') {
            console.error('Firestore connection test failed:', err);
          }
        });
      } catch (error) {
        console.error('Firestore test error:', error);
      }
    };
    testFirestore();
  }, []);

  const getAuthHeaders = (isJson = true) => {
    const tokens = localStorage.getItem('google_tokens');
    const headers: Record<string, string> = {};
    if (tokens) {
      headers['Authorization'] = `Bearer ${tokens}`;
    }
    if (isJson) {
      headers['Content-Type'] = 'application/json';
    }
    return headers;
  };

  const syncToSheet = async (sheetName: string, values: any[]) => {
    if (!isGoogleConnected) {
      throw new Error('Chưa kết nối Google Drive.');
    }
    try {
      const response = await fetch('/api/sheets/append', {
        method: 'POST',
        headers: getAuthHeaders(),
        body: JSON.stringify({
          range: `${sheetName}!A:Z`,
          values
        })
      });
      if (!response.ok) {
        const err = await response.json().catch(() => ({ error: 'Lỗi không xác định' }));
        if (response.status === 401) {
          setIsGoogleConnected(false);
          throw new Error('Phiên đăng nhập đã hết hạn. Vui lòng kết nối lại Google Drive.');
        }
        throw new Error(err.error || 'Lỗi đồng bộ Google Sheets');
      }
      return true;
    } catch (error: any) {
      console.error('Lỗi đồng bộ Google Sheets:', error);
      throw error;
    }
  };

  const handleConnectGoogle = async () => {
    try {
      const redirectUri = `${window.location.origin}/auth/callback`;
      const response = await fetch(`/api/auth/google/url?redirectUri=${encodeURIComponent(redirectUri)}`);
      
      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error || 'Failed to get auth URL');
      }
      
      const { url } = await response.json();
      
      const authWindow = window.open(url, 'oauth_popup', 'width=600,height=700');
      if (!authWindow) {
        alert('Please allow popups for this site to connect your account.');
      }
    } catch (error: any) {
      console.error('OAuth error:', error);
      alert(`Lỗi kết nối Google Drive: ${error.message}`);
    }
  };

  const analyzeDGA = () => {
    const h2 = parseFloat(dgaData.h2) || 0;
    const ch4 = parseFloat(dgaData.ch4) || 0;
    const c2h6 = parseFloat(dgaData.c2h6) || 0;
    const c2h4 = parseFloat(dgaData.c2h4) || 0;
    const c2h2 = parseFloat(dgaData.c2h2) || 0;
    const co = parseFloat(dgaData.co) || 0;
    const co2 = parseFloat(dgaData.co2) || 0;
    
    const matrix: Record<string, string> = {};
    const methods = ['IEEE (Dornenburg)', 'IEEE (Rogers)', 'IEC Ratio', 'Duval T1', 'Duval T4', 'Duval T5', 'Duval P1', 'Duval P2', 'KeyGas', 'ETRA', 'CO2/CO'];
    methods.forEach(m => matrix[m] = 'ND');

    // Duval T1
    const sumT1 = ch4 + c2h4 + c2h2;
    if (sumT1 > 0) {
        const pCH4 = ch4 / sumT1 * 100;
        const pC2H4 = c2h4 / sumT1 * 100;
        const pC2H2 = c2h2 / sumT1 * 100;
        if (pCH4 >= 98) matrix['Duval T1'] = 'PD';
        else if (pC2H2 >= 29 || (pC2H2 >= 13 && pC2H4 >= 23 && pC2H4 < 50) || (pC2H2 >= 15 && pC2H4 >= 50)) matrix['Duval T1'] = 'D2';
        else if (pC2H2 >= 13 && pC2H2 < 29 && pC2H4 < 23) matrix['Duval T1'] = 'D1';
        else if (pC2H2 < 15 && pC2H4 >= 50) matrix['Duval T1'] = 'T3';
        else if (pC2H2 < 4 && pC2H4 >= 20 && pC2H4 < 50) matrix['Duval T1'] = 'T2';
        else if (pC2H2 < 4 && pC2H4 < 20) matrix['Duval T1'] = 'T1';
        else matrix['Duval T1'] = 'DT';
    }

    // Duval T4
    const sumT4 = h2 + ch4 + c2h6;
    if (sumT4 > 0) {
        const pH2 = h2 / sumT4 * 100;
        const pCH4 = ch4 / sumT4 * 100;
        const pC2H6 = c2h6 / sumT4 * 100;
        if (pCH4 >= 2 && pCH4 < 15 && pC2H6 < 1) matrix['Duval T4'] = 'PD';
        else if (pH2 < 9 && pC2H6 >= 30) matrix['Duval T4'] = 'O';
        else if (pCH4 >= 36 && pC2H6 >= 24) matrix['Duval T4'] = 'C';
        else if (pH2 < 15 && pC2H6 >= 24 && pC2H6 < 30) matrix['Duval T4'] = 'C';
        else if (pH2 >= 9 && pC2H6 >= 46) matrix['Duval T4'] = 'ND';
        else matrix['Duval T4'] = 'S';
    }

    // Duval T5
    const sumT5 = ch4 + c2h4 + c2h6;
    if (sumT5 > 0) {
        const pC2H4 = c2h4 / sumT5 * 100;
        const pC2H6 = c2h6 / sumT5 * 100;
        if (pC2H4 < 1 && pC2H6 >= 2 && pC2H6 < 14) matrix['Duval T5'] = 'PD';
        else if (pC2H4 >= 1 && pC2H4 < 10 && pC2H6 >= 2 && pC2H6 < 14) matrix['Duval T5'] = 'O';
        else if (pC2H4 < 1 && pC2H6 < 2) matrix['Duval T5'] = 'O';
        else if (pC2H4 < 10 && pC2H6 >= 54) matrix['Duval T5'] = 'O';
        else if (pC2H4 < 10 && pC2H6 >= 14 && pC2H6 < 54) matrix['Duval T5'] = 'S';
        else if (pC2H4 >= 10 && pC2H4 < 35 && pC2H6 < 12) matrix['Duval T5'] = 'T2';
        else if (pC2H4 >= 35 && pC2H6 < 12) matrix['Duval T5'] = 'T3';
        else if (pC2H4 >= 50 && pC2H6 >= 12 && pC2H6 < 14) matrix['Duval T5'] = 'T3';
        else if (pC2H4 >= 70 && pC2H6 >= 14) matrix['Duval T5'] = 'T3';
        else if (pC2H4 >= 35 && pC2H6 >= 30) matrix['Duval T5'] = 'T3';
        else if (pC2H4 >= 10 && pC2H4 < 50 && pC2H6 >= 12 && pC2H6 < 14) matrix['Duval T5'] = 'C';
        else if (pC2H4 >= 10 && pC2H4 < 70 && pC2H6 >= 14 && pC2H6 < 30) matrix['Duval T5'] = 'C';
        else if (pC2H4 >= 10 && pC2H4 < 35 && pC2H6 >= 30) matrix['Duval T5'] = 'ND';
        else matrix['Duval T5'] = 'ND';
    }

    // Duval Pentagons
    const sumP = h2 + c2h6 + ch4 + c2h4 + c2h2;
    if (sumP > 0) {
        const p1 = h2 / sumP;
        const p2 = c2h6 / sumP;
        const p3 = ch4 / sumP;
        const p4 = c2h4 / sumP;
        const p5 = c2h2 / sumP;

        const x = p1 * 0 + p2 * (-38) + p3 * (-23.5) + p4 * 23.5 + p5 * 38;
        const y = p1 * 40 + p2 * 12.4 + p3 * (-32.4) + p4 * (-32.4) + p5 * 12.4;
        const pt = [x, y];

        const pointInPolygon = (point: number[], vs: number[][]) => {
            let inside = false;
            for (let i = 0, j = vs.length - 1; i < vs.length; j = i++) {
                let xi = vs[i][0], yi = vs[i][1];
                let xj = vs[j][0], yj = vs[j][1];
                let intersect = ((yi > point[1]) !== (yj > point[1]))
                    && (point[0] < (xj - xi) * (point[1] - yi) / (yj - yi) + xi);
                if (intersect) inside = !inside;
            }
            return inside;
        };

        for (const zone of regionsP1) {
            if (pointInPolygon(pt, zone.polygon)) {
                matrix['Duval P1'] = zone.name;
                break;
            }
        }
        for (const zone of regionsP2) {
            if (pointInPolygon(pt, zone.polygon)) {
                matrix['Duval P2'] = zone.name;
                break;
            }
        }
    }

    // KeyGas
    let maxGas = '';
    let maxVal = 0;
    const gases = { h2, ch4, c2h6, c2h4, c2h2, co };
    for (const [gas, val] of Object.entries(gases)) {
        if (val > maxVal) {
            maxVal = val;
            maxGas = gas;
        }
    }
    if (maxGas === 'h2') {
        if (c2h2 > 0.1 * h2) matrix['KeyGas'] = 'Arcing';
        else matrix['KeyGas'] = 'PD/Corona';
    }
    else if (maxGas === 'c2h4') matrix['KeyGas'] = 'Overheating (Oil)';
    else if (maxGas === 'co') matrix['KeyGas'] = 'Overheating (Paper)';
    else if (maxGas === 'c2h2') matrix['KeyGas'] = 'Arcing';
    else if (maxGas === 'ch4' || maxGas === 'c2h6') matrix['KeyGas'] = 'Low Temp Thermal';
    else matrix['KeyGas'] = 'Normal';

    // ETRA
    const totalCombustible = h2 + ch4 + c2h6 + c2h4 + c2h2 + co;
    if (totalCombustible < 720) matrix['ETRA'] = 'Normal';
    else if (totalCombustible < 1920) matrix['ETRA'] = 'Condition 2';
    else if (totalCombustible < 4630) matrix['ETRA'] = 'Condition 3';
    else matrix['ETRA'] = 'Condition 4';

    // CO2/CO
    if (co > 500 && co2 > 5000) {
        const ratio = co2 / co;
        if (ratio < 3) matrix['CO2/CO'] = 'C';
        else if (ratio > 7 && ratio < 10) matrix['CO2/CO'] = 'OK';
        else if (ratio > 10) matrix['CO2/CO'] = 'O';
        else matrix['CO2/CO'] = 'ND';
    } else {
        matrix['CO2/CO'] = 'ND';
    }

    // IEC Ratio & Rogers Ratios
    const ratio2 = c2h4 !== 0 ? c2h2 / c2h4 : 0; // C2H2/C2H4
    const ratio1 = h2 !== 0 ? ch4 / h2 : 0;      // CH4/H2
    const ratio3 = c2h6 !== 0 ? c2h4 / c2h6 : 0; // C2H4/C2H6

    // IEC 60599
    if (ratio1 < 0.1 && ratio3 < 0.2) matrix['IEC Ratio'] = 'PD';
    else if (ratio2 > 1.0 && ratio1 >= 0.1 && ratio1 <= 0.5 && ratio3 > 1.0) matrix['IEC Ratio'] = 'D1';
    else if (ratio2 >= 0.6 && ratio2 <= 2.5 && ratio1 >= 0.1 && ratio1 <= 1.0 && ratio3 > 2.0) matrix['IEC Ratio'] = 'D2';
    else if (ratio2 < 0.1 && ratio1 > 1.0 && ratio3 < 1.0) matrix['IEC Ratio'] = 'T1';
    else if (ratio2 < 0.1 && ratio1 > 1.0 && ratio3 >= 1.0 && ratio3 <= 4.0) matrix['IEC Ratio'] = 'T2';
    else if (ratio2 < 0.2 && ratio1 > 1.0 && ratio3 > 4.0) matrix['IEC Ratio'] = 'T3';
    else matrix['IEC Ratio'] = 'ND';

    // Rogers Ratios (IEEE PC57.104)
    if (ratio2 < 0.1 && ratio1 >= 0.1 && ratio1 <= 1.0 && ratio3 < 1.0) matrix['IEEE (Rogers)'] = 'OK';
    else if (ratio2 < 0.1 && ratio1 < 0.1 && ratio3 < 1.0) matrix['IEEE (Rogers)'] = 'PD';
    else if (ratio2 >= 0.1 && ratio1 >= 0.1 && ratio1 <= 1.0 && ratio3 >= 1.0) matrix['IEEE (Rogers)'] = 'D2'; // Arcing
    else if (ratio2 < 0.1 && ratio1 > 1.0 && ratio3 < 1.0) matrix['IEEE (Rogers)'] = 'T1';
    else if (ratio2 < 0.1 && ratio1 > 1.0 && ratio3 >= 1.0 && ratio3 <= 3.0) matrix['IEEE (Rogers)'] = 'T2';
    else if (ratio2 < 0.1 && ratio1 > 1.0 && ratio3 > 3.0) matrix['IEEE (Rogers)'] = 'T3';
    else matrix['IEEE (Rogers)'] = 'ND';

    // Dornenburg Ratios
    const d_r1 = h2 !== 0 ? ch4 / h2 : 0;
    const d_r2 = c2h4 !== 0 ? c2h2 / c2h4 : 0;
    const d_r3 = ch4 !== 0 ? c2h2 / ch4 : 0;
    const d_r4 = c2h2 !== 0 ? c2h6 / c2h2 : 0;

    if (d_r1 > 1.0 && d_r2 < 0.75 && d_r3 < 0.3 && d_r4 > 0.4) matrix['IEEE (Dornenburg)'] = 'T1';
    else if (d_r1 < 0.1 && d_r3 < 0.3 && d_r4 > 0.4) matrix['IEEE (Dornenburg)'] = 'PD';
    else if (d_r1 > 0.1 && d_r1 < 1.0 && d_r2 > 0.75 && d_r3 > 0.3 && d_r4 < 0.4) matrix['IEEE (Dornenburg)'] = 'D2';
    else matrix['IEEE (Dornenburg)'] = 'ND';

    // DNO CNAIM Calculation
    const age = parseFloat(dgaData.age) || 20;
    const loadFactor = parseFloat(dgaData.loadFactor) || 50; // Default 50% load
    
    // Duty Factor based on load (Table 32/33)
    let dutyFactor = 1;
    if (loadFactor <= 50) dutyFactor = 0.9;
    else if (loadFactor <= 70) dutyFactor = 0.95;
    else if (loadFactor <= 100) dutyFactor = 1;
    else dutyFactor = 1.4;

    const normalExpectedLife = 60; // Assuming pre-1980 or standard
    const expectedLife = normalExpectedLife / dutyFactor;
    
    const beta1 = Math.log(5.5/0.5) / expectedLife;
    let initialHealthScore = 0.5 * Math.exp(beta1 * age);
    initialHealthScore = Math.min(5.5, initialHealthScore); // Capped at 5.5
    
    // DGA Factor (simplified based on Table 201-205)
    let dgaScore = 0;
    if (h2 > 100) dgaScore += 10 * 50; else if (h2 > 40) dgaScore += 4 * 50;
    if (ch4 > 150) dgaScore += 10 * 30; else if (ch4 > 50) dgaScore += 4 * 30;
    if (c2h4 > 150) dgaScore += 10 * 30; else if (c2h4 > 50) dgaScore += 4 * 30;
    if (c2h6 > 150) dgaScore += 10 * 30; else if (c2h6 > 50) dgaScore += 4 * 30;
    if (c2h2 > 20) dgaScore += 8 * 120; else if (c2h2 > 5) dgaScore += 4 * 120;
    
    const dgaTestCollar = dgaScore / 220;
    let dgaFactor = 1;
    if (dgaTestCollar > 7) dgaFactor = 1.5;
    else if (dgaTestCollar > 4) dgaFactor = 1.2;

    // Oil Factor (simplified based on Table 196-198)
    let oilScore = 0;
    const moisture = parseFloat(dgaData.moisture) || 0;
    const acidity = parseFloat(dgaData.acidity) || 0;
    const bdStrength = parseFloat(dgaData.bdStrength) || 60;
    
    if (moisture > 35) oilScore += 8 * 80; else if (moisture > 25) oilScore += 4 * 80;
    if (acidity > 0.2) oilScore += 8 * 125; else if (acidity > 0.15) oilScore += 4 * 125;
    if (bdStrength < 30) oilScore += 10 * 80; else if (bdStrength < 40) oilScore += 4 * 80;
    
    let oilFactor = 1;
    if (oilScore > 1000) oilFactor = 1.2;
    else if (oilScore > 500) oilFactor = 1.1;
    else if (oilScore > 200) oilFactor = 1.05;

    // MMI combination for Health Score Factor
    const factors = [dgaFactor, oilFactor];
    const maxFactor = Math.max(...factors);
    let healthScoreFactor = maxFactor;
    if (maxFactor > 1) {
        const otherFactors = factors.filter(f => f !== maxFactor && f > 1);
        const sumOther = otherFactors.reduce((a, b) => a + (b - 1), 0);
        healthScoreFactor = maxFactor + (sumOther / 1.5);
    }

    let currentHealthScore = initialHealthScore * healthScoreFactor;
    currentHealthScore = Math.max(currentHealthScore, dgaTestCollar); // Apply collar
    currentHealthScore = Math.min(10, currentHealthScore); // Capped at 10
    
    // Future Health Score (10 years)
    let beta2 = beta1;
    if (currentHealthScore > 0.5) {
        beta2 = Math.log(currentHealthScore / 0.5) / age;
        beta2 = Math.min(beta2, 2 * beta1); // Capped at 2 * beta1
    }
    
    let ageingReduction = 1;
    if (currentHealthScore > 5.5) ageingReduction = 1.5;
    else if (currentHealthScore > 2) ageingReduction = ((currentHealthScore - 2) / 7) + 1;
    
    const futureHealthScore = Math.min(15, currentHealthScore * Math.exp((beta2 / ageingReduction) * 10));

    const getHIBand = (score: number) => {
        if (score < 4) return 'HI1';
        if (score < 5.5) return 'HI2';
        if (score < 6.5) return 'HI3';
        if (score < 8) return 'HI4';
        return 'HI5';
    };

    const currentHI = getHIBand(currentHealthScore);
    const futureHI = getHIBand(futureHealthScore);

    // TDCG IEEE C57.104
    const tdcg = h2 + ch4 + c2h6 + c2h4 + c2h2 + co;
    let tdcgCondition = '';
    let tdcgColor = '';
    let tdcgRecommendation = '';

    if (tdcg <= 720) {
        tdcgCondition = 'Condition 1';
        tdcgColor = 'text-emerald-700 bg-emerald-100 border-emerald-300';
        tdcgRecommendation = 'TDCG ở mức bình thường. Tiếp tục vận hành bình thường và lấy mẫu định kỳ.';
    } else if (tdcg <= 1920) {
        tdcgCondition = 'Condition 2';
        tdcgColor = 'text-yellow-700 bg-yellow-100 border-yellow-300';
        tdcgRecommendation = 'TDCG cao hơn bình thường. Cần theo dõi sự gia tăng khí. Đề nghị lấy mẫu lại sau 3-6 tháng.';
    } else if (tdcg <= 4630) {
        tdcgCondition = 'Condition 3';
        tdcgColor = 'text-orange-700 bg-orange-100 border-orange-300';
        tdcgRecommendation = 'TDCG ở mức cảnh báo. Có thể có sự cố đang phát triển. Đề nghị lấy mẫu lại sau 1-2 tháng và xem xét các thử nghiệm khác.';
    } else {
        tdcgCondition = 'Condition 4';
        tdcgColor = 'text-red-700 bg-red-100 border-red-300';
        tdcgRecommendation = 'TDCG ở mức nguy hiểm. Sự cố có thể đang xảy ra. Đề nghị ngừng vận hành máy biến áp để kiểm tra ngay lập tức.';
    }

    const healthScoreData = [
      { 
        name: 'Hiện tại', 
        'Health Score': parseFloat(currentHealthScore.toFixed(2)),
        fill: currentHealthScore < 4 ? '#22c55e' : currentHealthScore < 5.5 ? '#84cc16' : currentHealthScore < 6.5 ? '#eab308' : currentHealthScore < 8 ? '#f97316' : '#ef4444'
      },
      { 
        name: 'Sau 10 năm', 
        'Health Score': parseFloat(futureHealthScore.toFixed(2)),
        fill: futureHealthScore < 4 ? '#22c55e' : futureHealthScore < 5.5 ? '#84cc16' : futureHealthScore < 6.5 ? '#eab308' : futureHealthScore < 8 ? '#f97316' : '#ef4444'
      }
    ];

    let condition = currentHealthScore < 4 ? 'Tốt' : currentHealthScore < 6.5 ? 'Trung bình' : 'Kém';
    let recommendations = [tdcgRecommendation];
    let recommendationColor = 'emerald';
    
    if (currentHealthScore >= 6.5) {
      recommendations.push('Cần lên kế hoạch bảo dưỡng hoặc thay thế trong thời gian ngắn do Health Index cao.');
      recommendationColor = 'red';
    } else if (currentHealthScore >= 4) {
      recommendations.push('Tăng cường tần suất lấy mẫu DGA và theo dõi chặt chẽ do Health Index có dấu hiệu suy giảm.');
      recommendationColor = 'yellow';
    }

    // DP Estimation (Chendong model approximation or similar heuristic based on CO/CO2)
    // A common heuristic: DP = 1000 - 150 * ln(CO/10 + 1)
    const dpEstimation = Math.max(200, Math.round(1000 - 150 * Math.log(co / 10 + 1)));

    // CNAIM POF and EOL
    const currentPOF = (0.00073 * Math.exp(0.5 * currentHealthScore)).toFixed(4);
    const futurePOF = (0.00073 * Math.exp(0.5 * futureHealthScore)).toFixed(4);
    const eolYears = currentHealthScore >= 7 ? 0 : Math.max(0, Math.round(Math.log(7.0 / currentHealthScore) / beta1));

    setDgaAnalysisResult({
      matrix,
      currentHealthScore: currentHealthScore.toFixed(2),
      futureHealthScore: futureHealthScore.toFixed(2),
      currentHI,
      futureHI,
      healthScoreData,
      tdcg,
      tdcgCondition,
      tdcgColor,
      condition,
      recommendations,
      recommendationColor,
      dpEstimation,
      currentPOF,
      futurePOF,
      eolYears,
      timestamp: new Date().toLocaleString('vi-VN')
    });
  };

  const analyzePVCell = () => {
    const eff = parseFloat(pvCellData.efficiency);
    const temp = parseFloat(pvCellData.temp);
    
    let condition = 'Tốt';
    let recommendations = ['Tiếp tục theo dõi định kỳ.'];
    let color = 'text-emerald-600';
    let bgColor = 'bg-emerald-50';

    if (eff < 15 || temp > 65) {
      condition = 'Nguy hiểm';
      recommendations = ['Kiểm tra điểm nóng (hotspot) ngay lập tức.', 'Vệ sinh bề mặt tấm pin.', 'Kiểm tra đấu nối inverter.'];
      color = 'text-red-600';
      bgColor = 'bg-red-50';
    } else if (eff < 17 || temp > 55) {
      condition = 'Cảnh báo';
      recommendations = ['Kiểm tra vệ sinh tấm pin.', 'Theo dõi nhiệt độ vận hành.'];
      color = 'text-amber-600';
      bgColor = 'bg-amber-50';
    }

    setPvAnalysisResult({
      condition,
      recommendations,
      color,
      bgColor,
      timestamp: new Date().toLocaleString('vi-VN')
    });
  };

  const analyzeWindBlade = () => {
    const vib = parseFloat(windBladeData.vibration);
    const acoustic = parseFloat(windBladeData.acoustic);
    
    let condition = 'Tốt';
    let recommendations = ['Bảo trì định kỳ theo kế hoạch.'];
    let color = 'text-emerald-600';
    let bgColor = 'bg-emerald-50';

    if (vib > 1.5 || acoustic > 60) {
      condition = 'Nguy hiểm';
      recommendations = ['Dừng turbine ngay lập tức để kiểm tra vết nứt.', 'Kiểm tra độ cân bằng cánh.', 'Sử dụng drone kiểm tra chi tiết bề mặt.'];
      color = 'text-red-600';
      bgColor = 'bg-red-50';
    } else if (vib > 0.8 || acoustic > 40) {
      condition = 'Cảnh báo';
      recommendations = ['Tăng tần suất giám sát rung động.', 'Lên kế hoạch kiểm tra bằng hình ảnh trong kỳ dừng máy tới.'];
      color = 'text-amber-600';
      bgColor = 'bg-amber-50';
    }

    setWindAnalysisResult({
      condition,
      recommendations,
      color,
      bgColor,
      timestamp: new Date().toLocaleString('vi-VN')
    });
  };

  const exportToPDF = async () => {
    const reportElement = document.getElementById('dga-report-content');
    if (!reportElement) return;

    try {
      // Save original styles
      const originalStyle = reportElement.style.cssText;
      
      // Temporarily modify styles for A4 format (approx 210x297mm)
      // We set a fixed width of 1024px to ensure the layout doesn't squish
      reportElement.style.width = '1024px';
      reportElement.style.maxWidth = 'none';
      reportElement.style.margin = '0';
      reportElement.style.padding = '40px';
      reportElement.style.boxShadow = 'none';
      reportElement.style.border = 'none';
      
      // Add a small delay to allow DOM to update
      await new Promise(resolve => setTimeout(resolve, 100));

      const imgData = await htmlToImage.toPng(reportElement, { 
        quality: 1.0,
        pixelRatio: 2,
        backgroundColor: '#ffffff'
      });
      
      // Restore original styles
      reportElement.style.cssText = originalStyle;
      
      const pdf = new jsPDF('p', 'mm', 'a4');
      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      
      // Calculate height based on element's aspect ratio
      const elementWidth = 1024;
      const elementHeight = reportElement.offsetHeight;
      const pdfHeight = (elementHeight * pdfWidth) / elementWidth;
      
      let heightLeft = pdfHeight;
      let position = 0;

      // Add first page
      pdf.addImage(imgData, 'PNG', 0, position, pdfWidth, pdfHeight);
      heightLeft -= pageHeight;

      // Add subsequent pages if the content is longer than one page
      while (heightLeft > 0) {
        position = heightLeft - pdfHeight;
        pdf.addPage();
        pdf.addImage(imgData, 'PNG', 0, position, pdfWidth, pdfHeight);
        heightLeft -= pageHeight;
      }

      pdf.save(`DGA_Report_${equipmentCode}_${new Date().toISOString().split('T')[0]}.pdf`);
    } catch (error) {
      console.error('Error generating PDF:', error);
      alert('Có lỗi xảy ra khi tạo PDF.');
    }
  };

  const handleSeedData = () => {
    // Removed to prevent overwriting real data with mock data
  };

  const executeSeedData = async () => {
    // Removed to prevent overwriting real data with mock data
  };

  /**
   * Kiểm tra phiên bản dữ liệu (updatedAt / version timestamp) giữa Firestore và Google Sheets qua Sheets API
   * Đảm bảo tính nhất quán (Sync Reconciliation) trước khi thực hiện ghi đè dữ liệu.
   */
  const checkDataVersionReconciliation = async (options?: { silent?: boolean }): Promise<ReconciliationReport | null> => {
    const silent = options?.silent ?? false;
    if (!isGoogleConnected) {
      if (!silent) alert('Vui lòng kết nối Google Drive trước khi kiểm tra đối soát.');
      return null;
    }

    setIsReconciling(true);
    try {
      const spreadsheetId = localStorage.getItem('tev_spreadsheet_id') || '';
      const response = await fetch('/api/sheets/get', {
        headers: {
          ...getAuthHeaders(),
          ...(spreadsheetId ? { 'x-spreadsheet-id': spreadsheetId } : {})
        }
      });

      if (!response.ok) {
        const err = await response.json();
        throw new Error(err.error || 'Failed to fetch from Google Sheets API');
      }

      const resJson = await response.json();
      const allSheets = resJson.data?.allSheets || {};

      // Đối soát giữa dữ liệu Firestore trong state với các hàng trong Google Sheets
      const report = compareVersionsBetweenFirestoreAndSheets({
        firestoreData: {
          workOrders,
          customers,
          equipment: allEquipment,
          inventory
        },
        sheetsData: {
          allSheets
        }
      });

      setReconciliationReport(report);
      return report;
    } catch (error: any) {
      console.error('Error in checkDataVersionReconciliation:', error);
      if (!silent) {
        alert('Lỗi kiểm tra phiên bản Google Sheets: ' + (error.message || String(error)));
      }
      return null;
    } finally {
      setIsReconciling(false);
    }
  };

  const handleOpenReconciliation = async () => {
    setShowReconciliationModal(true);
    await checkDataVersionReconciliation({ silent: true });
  };

  /**
   * Hòa giải phiên bản thông minh (Newer Wins):
   * - Bản ghi nào trên Google Sheets có timestamp mới hơn -> nạp về Cloud Firestore.
   * - Bản ghi nào trên Firestore có timestamp mới hơn -> đẩy lên Google Sheets.
   * - Tuyệt đối không làm mất dữ liệu của cả 2 phía.
   */
  const handleResolveNewerWins = async () => {
    if (!reconciliationReport) return;
    setIsReconciling(true);
    try {
      let pulledCount = 0;
      let pushedCount = 0;

      // 1. Kéo các bản ghi mà Sheets mới hơn hoặc chỉ có ở Sheets về Firestore
      const sheetsItemsToPull = reconciliationReport.items.filter(i => 
        i.status === 'sheets_newer' || i.status === 'only_in_sheets'
      );

      for (const item of sheetsItemsToPull) {
        if (item.entityType === 'customer' && item.sheetsData) {
          const row = item.sheetsData;
          const custId = String(row[0] || item.id).trim();
          const custName = String(row[1] || item.name).trim();
          const factories = row[2] ? String(row[2]).split(',').map((s: string) => s.trim()).filter(Boolean) : [];
          const email = String(row[3] || '');
          const phone = String(row[4] || '');
          const address = String(row[5] || '');
          const createdAt = String(row[6] || new Date().toISOString());
          const updatedAt = String(row[7] || row[6] || new Date().toISOString());
          
          const updatedCust = { id: custId, name: custName, factories, email, phone, address, createdAt, updatedAt };
          setCustomers(prev => {
            const idx = prev.findIndex(c => c.id.toLowerCase() === custId.toLowerCase());
            if (idx >= 0) {
              const next = [...prev];
              next[idx] = updatedCust;
              return next;
            }
            return [...prev, updatedCust];
          });
          if (auth.currentUser) {
            await setDoc(doc(db, 'customers', custId), updatedCust, { merge: true });
          }
          pulledCount++;
        } else if (item.entityType === 'workOrder' && item.sheetsData) {
          const row = item.sheetsData;
          const woId = String(row[0] || item.id).trim();
          const updatedWo: any = {
            id: woId,
            workPermitId: String(row[1] || ''),
            title: String(row[2] || item.name),
            description: String(row[3] || ''),
            equipmentId: String(row[4] || ''),
            failureCode: String(row[5] || ''),
            customer: String(row[6] || ''),
            type: String(row[7] || ''),
            isUnplanned: String(row[8] || '').toLowerCase() === 'có',
            priority: String(row[9] || 'Medium'),
            status: String(row[10] || 'open'),
            assignedTo: String(row[11] || ''),
            responsibleApprove: String(row[12] || ''),
            responsibleDo: String(row[13] || ''),
            blockingRequired: String(row[14] || '').toLowerCase() === 'có',
            dueDate: String(row[15] || ''),
            usedMaterials: String(row[16] || ''),
            createdAt: String(row[17] || new Date().toISOString()),
            updatedAt: String(row[18] || row[17] || new Date().toISOString()),
            downtimeStart: String(row[19] || ''),
            repairStart: String(row[20] || ''),
            repairEnd: String(row[21] || ''),
            restartTime: String(row[22] || '')
          };
          setWorkOrders(prev => {
            const idx = prev.findIndex(w => w.id.toLowerCase() === woId.toLowerCase());
            if (idx >= 0) {
              const next = [...prev];
              next[idx] = updatedWo;
              return next;
            }
            return [updatedWo, ...prev];
          });
          if (auth.currentUser) {
            await setDoc(doc(db, 'workOrders', woId), updatedWo, { merge: true });
          }
          pulledCount++;
        } else if (item.entityType === 'inventory' && item.sheetsData) {
          const row = item.sheetsData;
          const invId = String(row[0] || item.id).trim();
          const updatedInv = {
            id: invId,
            name: String(row[1] || item.name),
            sku: String(row[2] || ''),
            category: String(row[3] || ''),
            quantity: parseFloat(String(row[4] || '0').replace(',', '.')) || 0,
            unit: String(row[5] || ''),
            minStock: parseFloat(String(row[6] || '0').replace(',', '.')) || 0,
            location: String(row[7] || ''),
            price: parseFloat(String(row[8] || '0').replace(',', '.')) || 0,
            createdAt: String(row[9] || new Date().toISOString()),
            updatedAt: String(row[10] || row[9] || new Date().toISOString())
          };
          setInventory(prev => {
            const idx = prev.findIndex(i => i.id.toLowerCase() === invId.toLowerCase());
            if (idx >= 0) {
              const next = [...prev];
              next[idx] = updatedInv;
              return next;
            }
            return [...prev, updatedInv];
          });
          if (auth.currentUser) {
            await setDoc(doc(db, 'inventory', invId), updatedInv, { merge: true });
          }
          pulledCount++;
        }
      }

      // 2. Đẩy các bản ghi mà Firestore mới hơn lên Sheets
      const firestoreItemsToPush = reconciliationReport.items.filter(i => 
        i.status === 'firestore_newer' || i.status === 'only_in_firestore'
      );
      if (firestoreItemsToPush.length > 0) {
        await handleSyncToSheets(undefined, true, true);
        pushedCount = firestoreItemsToPush.length;
      }

      // 3. Cập nhật lại báo cáo đối soát
      await checkDataVersionReconciliation({ silent: true });
      alert(`✅ Hòa giải phiên bản (Newer Wins) hoàn tất!\n- Đã kéo ${pulledCount} mục mới hơn từ Google Sheets về Firestore.\n- Đã đồng bộ ${pushedCount} mục mới hơn từ Firestore lên Google Sheets.`);
    } catch (e: any) {
      console.error('Reconciliation error:', e);
      alert('Lỗi hòa giải phiên bản: ' + (e.message || String(e)));
    } finally {
      setIsReconciling(false);
    }
  };

  const handleForcePushToSheets = async () => {
    setShowReconciliationModal(false);
    await handleSyncToSheets(undefined, false, true);
  };

  const handlePullFromSheetsToFirestore = async () => {
    setShowReconciliationModal(false);
    await handleFetchFromSheets();
  };

  const handleSyncToSheets = async (equipmentListToSync?: any[], silent: boolean = false, force: boolean = false) => {
    if (!isGoogleConnected) {
      if (!silent) alert('Vui lòng kết nối Google Drive trước khi đồng bộ.');
      return;
    }

    // Kiểm tra phiên bản dữ liệu đối soát ngầm (không tự động mở popup gây phiền người dùng)
    if (!force) {
      try {
        await checkDataVersionReconciliation({ silent: true });
      } catch (recErr) {
        console.warn('Reconciliation audit check:', recErr);
      }
    }

    setIsSyncing(true);
    try {
      // Prepare data for all equipment
      const transformers: any[] = [];
      const switchgears: any[] = [];
      const motors: any[] = [];
      const inverters: any[] = [];
      const cmmsData: any[] = [];
      const inventoryData: any[] = [];
      const customersData: any[] = [];
      
      // Headers
      transformers.push(['Thời gian kiểm tra', 'Khách hàng', 'Nhà máy / Site', 'Vị trí / Khu vực', 'Mã thiết bị', 'Tên thiết bị', 'Loại thiết bị', 'Điểm sức khỏe (%)', 'Trạng thái', 'Nhiệt độ dầu (°C)', 'Nhiệt độ cuộn dây (°C)', 'IR Cao-Hạ (MΩ)', 'IR Cao-Vỏ (MΩ)', 'Tình trạng rò rỉ dầu', 'Khí hòa tan DGA (ppm)', 'Độ bền điện môi (kV)', 'Hàm lượng Furan (mg/kg)', 'Độ ẩm trong dầu (ppm)', 'Tuổi thọ (Age)', 'Hệ số làm việc (Duty Factor)', 'File đính kèm (Links)']);
      switchgears.push(['Thời gian kiểm tra', 'Khách hàng', 'Nhà máy / Site', 'Vị trí / Khu vực', 'Mã thiết bị', 'Tên thiết bị', 'Loại thiết bị', 'Điểm sức khỏe (%)', 'Trạng thái', 'Chụp ảnh nhiệt (°C)', 'Điện trở tiếp xúc (μΩ)', 'TEV (dBmV)', 'Siêu âm (dBμV)', 'Xung TEV (pps)', 'Độ ẩm (%)', 'Áp suất khí SF6 (bar)', 'Tuổi thọ (Age)', 'Hệ số làm việc (Duty Factor)', 'File đính kèm (Links)']);
      motors.push(['Thời gian kiểm tra', 'Khách hàng', 'Nhà máy / Site', 'Vị trí / Khu vực', 'Mã thiết bị', 'Tên thiết bị', 'Loại thiết bị', 'Điểm sức khỏe (%)', 'Trạng thái', 'Độ rung (mm/s)', 'Nhiệt độ Stator (°C)', 'Nhiệt độ vòng bi (°C)', 'Độ lệch điện áp (%)', 'Tan-delta R', 'Tan-delta Y', 'Tan-delta B', 'Tip-up R', 'Tip-up Y', 'Tip-up B', 'PD R', 'PD Y', 'PD B', 'IR R', 'IR Y', 'IR B', 'PI R', 'PI Y', 'PI B', 'DD R', 'DD Y', 'DD B', 'ELCID R', 'ELCID Y', 'ELCID B', 'Tuổi thọ (Age)', 'Hệ số làm việc (Duty Factor)', 'File đính kèm (Links)']);
      inverters.push(['Thời gian kiểm tra', 'Khách hàng', 'Nhà máy / Site', 'Vị trí / Khu vực', 'Mã thiết bị', 'Tên thiết bị', 'Loại thiết bị', 'Điểm sức khỏe (%)', 'Trạng thái', 'Điện áp DC (V)', 'Dải MPPT (V)', 'Dòng điện DC (A)', 'Điện áp AC (V)', 'Tần số AC (Hz)', 'Độ méo hài (%)', 'Hiệu suất tối đa (%)', 'Hiệu suất Châu Âu (%)', 'Chống hòa lưới', 'Dòng rò & Cách ly', 'Điện áp chịu đựng (kV)', 'Cách ly vật lý', 'Nhiệt độ vận hành (°C)', 'Bộ lọc khí', 'File đính kèm (Links)']);
      cmmsData.push([
        'Mã WO', 'Mã Work Permit', 'Tiêu đề', 'Mô tả', 'Thiết bị', 'Failure Code',
        'Khách hàng', 'Loại công việc', 'Ngoài kế hoạch', 'Mức độ ưu tiên', 'Trạng thái',
        'Người thực hiện (PIC)', 'Vai trò Phê duyệt', 'Vai trò Thực hiện', 'Yêu cầu Cô lập',
        'Hạn hoàn thành', 'Vật tư sử dụng', 'Ngày tạo', 'Ngày cập nhật',
        'Bắt đầu dừng máy', 'Bắt đầu sửa chữa', 'Kết thúc sửa chữa', 'Chạy lại máy'
      ]);
      inventoryData.push(['Mã vật tư', 'Tên vật tư', 'SKU', 'Danh mục', 'Số lượng', 'Đơn vị', 'Tồn kho tối thiểu', 'Vị trí', 'Đơn giá', 'Ngày tạo', 'Ngày cập nhật']);
      customersData.push(['Mã khách hàng', 'Tên khách hàng', 'Danh sách nhà máy', 'Email', 'Số điện thoại', 'Địa chỉ', 'Ngày tạo', 'Ngày cập nhật']);

      const equipmentData: any[] = [];
      equipmentData.push([
        'Mã thiết bị', 'Tên thiết bị', 'Loại thiết bị', 'Khách hàng', 'Nhà máy / Site', 'Vị trí / Khu vực',
        'Trạng thái', 'Điểm sức khỏe (HI %)', 'Thông số kỹ thuật (Specs)', 'Bảng tên Nameplate', 'Lần kiểm tra cuối', 'Ngày tạo', 'Ngày cập nhật'
      ]);
      allEquipment.forEach(eq => {
        equipmentData.push([
          eq.id || '',
          eq.name || '',
          eq.type || '',
          eq.customer || '',
          eq.factory || '',
          eq.location || '',
          eq.status || 'healthy',
          eq.health !== undefined ? eq.health : 100,
          typeof eq.technicalSpecs === 'object' ? JSON.stringify(eq.technicalSpecs) : (eq.technicalSpecs || ''),
          typeof eq.nameplate === 'object' ? JSON.stringify(eq.nameplate) : (eq.nameplate || ''),
          eq.lastCheck || '',
          eq.createdAt || '',
          eq.updatedAt || ''
        ]);
      });

      // Map existing data
      const listToSync = Array.isArray(equipmentListToSync) ? equipmentListToSync : allReports;
      
      let transformerRowIdx = 2;
      let switchgearRowIdx = 2;
      let motorRowIdx = 2;
      let inverterRowIdx = 2;

      listToSync.forEach(item => {
        const eqId = item.equipmentId || item.id;
        const eqName = item.equipmentName || item.name;
        const lastCheckDate = item.date || item.lastCheck || new Date().toLocaleDateString('vi-VN');
        const baseData = [lastCheckDate, item.customer || '', item.factory || '', item.location || '', eqId, eqName, item.type];
        const raw = item.rawData || [];
        
        if (item.type === 'Máy biến áp') {
          const rowIdx = transformerRowIdx++;
          const healthFormula = `=MAX(0, MIN(100, 100 - (IFS(I${rowIdx}="Nguy hiểm", 40, I${rowIdx}="Cảnh báo", 20, TRUE, 0)) - (IF(ISNUMBER(S${rowIdx}), S${rowIdx}, 0)/40)*10))`;
          const statusFormula = `=IFS(OR(AND(ISNUMBER(J${rowIdx}), J${rowIdx}>90), AND(ISNUMBER(K${rowIdx}), K${rowIdx}>105), AND(ISNUMBER(L${rowIdx}), L${rowIdx}<1000), AND(ISNUMBER(M${rowIdx}), M${rowIdx}<1000), N${rowIdx}="heavy", AND(ISNUMBER(O${rowIdx}), O${rowIdx}>2500), AND(ISNUMBER(P${rowIdx}), P${rowIdx}<40), AND(ISNUMBER(Q${rowIdx}), Q${rowIdx}>5), AND(ISNUMBER(R${rowIdx}), R${rowIdx}>25)), "Nguy hiểm", OR(AND(ISNUMBER(J${rowIdx}), J${rowIdx}>80), AND(ISNUMBER(K${rowIdx}), K${rowIdx}>90), AND(ISNUMBER(L${rowIdx}), L${rowIdx}<2000), AND(ISNUMBER(M${rowIdx}), M${rowIdx}<2000), N${rowIdx}="light", AND(ISNUMBER(O${rowIdx}), O${rowIdx}>1000), AND(ISNUMBER(P${rowIdx}), P${rowIdx}<50), AND(ISNUMBER(Q${rowIdx}), Q${rowIdx}>1), AND(ISNUMBER(R${rowIdx}), R${rowIdx}>15)), "Cảnh báo", TRUE, "Bình thường")`;
          
          const rowData = [...baseData, healthFormula, statusFormula];
          for (let i = 9; i < 27; i++) rowData.push(raw[i] !== undefined ? raw[i] : '');
          rowData.push(raw[27] !== undefined ? raw[27] : ''); // Age
          rowData.push(raw[28] !== undefined ? raw[28] : ''); // Duty Factor
          rowData.push(raw[29] !== undefined ? raw[29] : ''); // Links
          transformers.push(rowData);
        } else if (item.type === 'Tủ điện trung thế' || item.type === 'Tủ điện') {
          const rowIdx = switchgearRowIdx++;
          const healthFormula = `=MAX(0, MIN(100, 100 - (IFS(I${rowIdx}="Nguy hiểm", 40, I${rowIdx}="Cảnh báo", 20, TRUE, 0)) - (IF(ISNUMBER(Q${rowIdx}), Q${rowIdx}, 0)/40)*10))`;
          const statusFormula = `=IFS(OR(AND(ISNUMBER(J${rowIdx}), J${rowIdx}>75), AND(ISNUMBER(K${rowIdx}), K${rowIdx}>100), AND(ISNUMBER(L${rowIdx}), L${rowIdx}>30), AND(ISNUMBER(M${rowIdx}), M${rowIdx}>15), AND(ISNUMBER(N${rowIdx}), N${rowIdx}>50), AND(ISNUMBER(O${rowIdx}), O${rowIdx}>85), AND(ISNUMBER(P${rowIdx}), P${rowIdx}<5.0)), "Nguy hiểm", OR(AND(ISNUMBER(J${rowIdx}), J${rowIdx}>60), AND(ISNUMBER(K${rowIdx}), K${rowIdx}>50), AND(ISNUMBER(L${rowIdx}), L${rowIdx}>=20), AND(ISNUMBER(M${rowIdx}), M${rowIdx}>5), AND(ISNUMBER(N${rowIdx}), N${rowIdx}>=10), AND(ISNUMBER(O${rowIdx}), O${rowIdx}>70), AND(ISNUMBER(P${rowIdx}), P${rowIdx}<5.5)), "Cảnh báo", TRUE, "Bình thường")`;

          const rowData = [...baseData, healthFormula, statusFormula];
          for (let i = 9; i < 16; i++) rowData.push(raw[i] !== undefined ? raw[i] : '');
          rowData.push(raw[16] !== undefined ? raw[16] : ''); // Age
          rowData.push(raw[17] !== undefined ? raw[17] : ''); // Duty Factor
          rowData.push(raw[18] !== undefined ? raw[18] : ''); // Links
          switchgears.push(rowData);
        } else if (item.type === 'Động cơ') {
          const rowIdx = motorRowIdx++;
          // Motor Health Index based on Weighted Scoring Method (3-phase)
          // Columns: N,O,P (Tan-delta), Q,R,S (Tip-up), T,U,V (PD), W,X,Y (IR), Z,AA,AB (PI), AC,AD,AE (DD), AF,AG,AH (ELCID)
          
          const worstTanDelta = `MAX(N${rowIdx}, O${rowIdx}, P${rowIdx})`;
          const worstTipUp = `MAX(Q${rowIdx}, R${rowIdx}, S${rowIdx})`;
          const worstPD = `MAX(T${rowIdx}, U${rowIdx}, V${rowIdx})`;
          const worstIR = `MIN(W${rowIdx}, X${rowIdx}, Y${rowIdx})`;
          const worstPI = `MIN(Z${rowIdx}, AA${rowIdx}, AB${rowIdx})`;
          const worstDD = `MAX(AC${rowIdx}, AD${rowIdx}, AE${rowIdx})`;
          const worstELCID = `MAX(AF${rowIdx}, AG${rowIdx}, AH${rowIdx})`;

          const sTanDelta = `IFS(${worstTanDelta}<0.02, 10, ${worstTanDelta}<=0.04, 7, ${worstTanDelta}<=0.07, 5, TRUE, 1)`;
          const sTipUp = `IFS(${worstTipUp}<0.002, 10, ${worstTipUp}<=0.004, 7, ${worstTipUp}<=0.006, 5, TRUE, 1)`;
          const sPD = `IFS(${worstPD}<=5000, 10, ${worstPD}<=10000, 7, ${worstPD}<=15000, 5, TRUE, 1)`;
          const sIR = `IFS(${worstIR}>50, 10, ${worstIR}>=10, 7, ${worstIR}>=1, 5, TRUE, 1)`;
          const sPI = `IFS(${worstPI}>2, 10, ${worstPI}>=1.5, 7, ${worstPI}>=1.0, 5, TRUE, 1)`;
          const sDD = `IFS(${worstDD}<2, 10, ${worstDD}<=4, 7, ${worstDD}<=8, 5, TRUE, 1)`;
          const sELCID = `IFS(${worstELCID}<100, 10, ${worstELCID}<=200, 7, ${worstELCID}<=300, 5, TRUE, 1)`;
          
      const sumScoreWeight = `((${sTanDelta}*2) + (${sTipUp}*2) + (${sPD}*3) + (${sIR}*1) + (${sPI}*1) + (${sDD}*1) + (${sELCID}*1))`;
      const sumWeights = `11`;
          
          const healthFormula = `=ROUND((${sumScoreWeight} / (10 * ${sumWeights})) * 100, 1)`;
          const statusFormula = `=IFS(H${rowIdx}<40, "Nguy hiểm", H${rowIdx}<65, "Cảnh báo", TRUE, "Bình thường")`;

          const rowData = [...baseData, healthFormula, statusFormula];
          // Vibration, Stator Temp, Bearing Temp, Voltage Imbalance (9, 10, 11, 12)
          for (let i = 9; i < 13; i++) rowData.push(raw[i] !== undefined ? raw[i] : '');
          // 3-phase tests (13 to 33)
          for (let i = 13; i < 34; i++) rowData.push(raw[i] !== undefined ? raw[i] : '');
          rowData.push(raw[34] !== undefined ? raw[34] : ''); // Age
          rowData.push(raw[35] !== undefined ? raw[35] : ''); // Duty Factor
          rowData.push(raw[36] !== undefined ? raw[36] : ''); // Links
          motors.push(rowData);
        } else if (item.type === 'Inverter') {
          const rowIdx = inverterRowIdx++;
          const rowData = [...baseData, item.health || 0, item.status || 'healthy'];
          const raw = item.measurements || {};
          
          rowData.push(raw.dcInputVoltage || '');
          rowData.push(raw.mpptVoltageRange || '');
          rowData.push(raw.dcInputCurrent || '');
          rowData.push(raw.acOutputVoltage || '');
          rowData.push(raw.acFrequency || '');
          rowData.push(raw.harmonicDistortion || '');
          rowData.push(raw.maxEfficiency || '');
          rowData.push(raw.euroEfficiency || '');
          rowData.push(raw.antiIslanding || '');
          rowData.push(raw.rcdIsolation || '');
          rowData.push(raw.dielectricVoltage || '');
          rowData.push(raw.physicalDisconnects || '');
          rowData.push(raw.operatingTemp || '');
          rowData.push(raw.airFilters || '');
          rowData.push(item.fileUrl || '');
          inverters.push(rowData);
        }
      });

      // Map CMMS data (23 columns TEV Service Flatform)
      workOrders.forEach(wo => {
        const materialsStr = Array.isArray(wo.usedMaterials)
          ? wo.usedMaterials.map((m: any) => `${m.name} (${m.quantity} ${m.unit})`).join(', ')
          : (wo.usedMaterials || '');
        const customerNameStr = customers.find(c => c.id === wo.customerId)?.name || wo.customer || wo.customerId || '';
        cmmsData.push([
          wo.id || '',
          wo.workPermitId || '',
          wo.title || '',
          wo.description || '',
          (Array.isArray(wo.equipmentId) ? wo.equipmentId.join(', ') : (wo.equipmentId || wo.equipmentName || '')) || '',
          wo.failureCode || '',
          customerNameStr,
          wo.type || '',
          wo.isUnplanned ? 'Có' : 'Không',
          wo.priority || '',
          wo.status || '',
          wo.assignedTo || '',
          wo.responsibleApprove || '',
          wo.responsibleDo || '',
          wo.blockingRequired ? 'Có' : 'Không',
          wo.dueDate ? (wo.dueDate.includes('T') ? new Date(wo.dueDate).toLocaleDateString('vi-VN') : wo.dueDate) : '',
          materialsStr,
          wo.createdAt ? new Date(wo.createdAt).toLocaleString('vi-VN') : '',
          wo.updatedAt ? new Date(wo.updatedAt).toLocaleString('vi-VN') : '',
          wo.downtimeStart || '',
          wo.repairStart || '',
          wo.repairEnd || '',
          wo.restartTime || ''
        ]);
      });

      // Map Inventory data
      inventory.forEach(item => {
        inventoryData.push([
          item.id,
          item.name,
          item.sku || '',
          item.category || '',
          item.quantity || 0,
          item.unit || '',
          item.minStock || 0,
          item.location || '',
          item.price || 0,
          item.createdAt || new Date().toISOString(),
          new Date().toISOString()
        ]);
      });

      // Map Customer data
      customers.forEach(cust => {
        customersData.push([
          cust.id,
          cust.name,
          (cust.factories || []).join(', '),
          cust.email || '',
          cust.phone || '',
          cust.address || '',
          cust.createdAt || new Date().toISOString(),
          cust.updatedAt || new Date().toISOString()
        ]);
      });

      // Map Generators
      const generators: any[] = [];
      generators.push(['Thời gian kiểm tra', 'Khách hàng', 'Nhà máy / Site', 'Vị trí / Khu vực', 'Mã thiết bị', 'Tên thiết bị', 'Điểm sức khỏe (%)', 'Trạng thái', 'Điện áp L1-L2 (V)', 'Tần số (Hz)', 'Áp suất dầu (bar)', 'Nhiệt độ nước làm mát (°C)', 'Điện áp ắc quy (V)', 'Thời gian đóng ATS (s)', 'Điện trở cách điện Stator (MΩ)', 'Điện trở cách điện Rotor (MΩ)', 'File đính kèm']);
      listToSync.filter(item => item.type === 'Máy phát').forEach(item => {
        const raw = item.measurements || {};
        generators.push([
          item.date || new Date().toLocaleDateString('vi-VN'),
          item.customer || '',
          item.factory || '',
          item.location || '',
          item.equipmentId || item.id,
          item.equipmentName || item.name,
          item.health !== undefined ? item.health : 95,
          item.status || 'healthy',
          raw.voltageL1L2 || '401',
          raw.frequencyHz || '50.05',
          raw.oilPressureBar || '4.8',
          raw.coolantTempC || '82',
          raw.batteryVoltageV || '26.8',
          raw.atsTransferTimeSec || '8.5',
          raw.insulationStator || '1250',
          raw.insulationRotor || '680',
          item.fileUrl || ''
        ]);
      });

      // Map Breakers (Máy cắt)
      const breakersData: any[] = [];
      breakersData.push(['Thời gian kiểm tra', 'Khách hàng', 'Nhà máy / Site', 'Vị trí / Khu vực', 'Mã thiết bị', 'Tên thiết bị', 'Điểm sức khỏe (%)', 'Trạng thái', 'Điện trở tiếp xúc Pha A (µΩ)', 'Điện trở tiếp xúc Pha B (µΩ)', 'Điện trở tiếp xúc Pha C (µΩ)', 'Cách điện tiếp điểm mở (MΩ)', 'Cách điện xuống đất (MΩ)', 'Thời gian đóng (ms)', 'Thời gian cắt (ms)', 'Độ chân không VCB', 'File đính kèm']);
      listToSync.filter(item => item.type === 'Máy cắt').forEach(item => {
        const raw = item.measurements || {};
        breakersData.push([
          item.date || new Date().toLocaleDateString('vi-VN'),
          item.customer || '',
          item.factory || '',
          item.location || '',
          item.equipmentId || item.id,
          item.equipmentName || item.name,
          item.health !== undefined ? item.health : 98,
          item.status || 'healthy',
          raw.contactResPhaseA || '18.2',
          raw.contactResPhaseB || '18.5',
          raw.contactResPhaseC || '18.4',
          raw.insulationOpenPoles || '2400',
          raw.insulationToGround || '2100',
          raw.closingTimeMs || '48',
          raw.openingTimeMs || '32',
          raw.vacuumIntegrity || 'Đạt tiêu chuẩn 25kV',
          item.fileUrl || ''
        ]);
      });

      // Map Cables (Cáp điện)
      const cablesData: any[] = [];
      cablesData.push(['Thời gian kiểm tra', 'Khách hàng', 'Nhà máy / Site', 'Vị trí tuyến cáp', 'Mã thiết bị', 'Tên thiết bị', 'Điểm sức khỏe (%)', 'Trạng thái', 'Cách điện Pha A-Đất (MΩ)', 'Cách điện Pha B-Đất (MΩ)', 'Cách điện Pha C-Đất (MΩ)', 'Tan-delta VLF (x10^-3)', 'PD Cáp (pC)', 'File đính kèm']);
      listToSync.filter(item => item.type === 'Cáp điện').forEach(item => {
        const raw = item.measurements || {};
        cablesData.push([
          item.date || new Date().toLocaleDateString('vi-VN'), item.customer || '', item.factory || '', item.location || '',
          item.equipmentId || item.id, item.equipmentName || item.name, item.health || 95, item.status || 'healthy',
          raw.irA || '4500', raw.irB || '4600', raw.irC || '4450', raw.tanDelta || '1.2', raw.pd || '50', item.fileUrl || ''
        ]);
      });

      // Map Batteries & UPS (Pin & UPS)
      const batteriesData: any[] = [];
      batteriesData.push(['Thời gian kiểm tra', 'Khách hàng', 'Nhà máy / Site', 'Vị trí', 'Mã thiết bị', 'Tên thiết bị', 'Điểm sức khỏe (%)', 'Trạng thái', 'Điện áp dàn bình (VDC)', 'Nội trở bình TB (mΩ)', 'Nhiệt độ bình (°C)', 'Thử phóng tải (Ah)', 'File đính kèm']);
      listToSync.filter(item => item.type === 'Pin & UPS').forEach(item => {
        const raw = item.measurements || {};
        batteriesData.push([
          item.date || new Date().toLocaleDateString('vi-VN'), item.customer || '', item.factory || '', item.location || '',
          item.equipmentId || item.id, item.equipmentName || item.name, item.health || 96, item.status || 'healthy',
          raw.bankVoltage || '228.5', raw.internalRes || '0.85', raw.temp || '26.5', raw.capacityTest || 'Pass (100%)', item.fileUrl || ''
        ]);
      });

      // Map Protective Relays (Rơ le bảo vệ)
      const relaysData: any[] = [];
      relaysData.push(['Thời gian kiểm tra', 'Khách hàng', 'Nhà máy / Site', 'Vị trí', 'Mã thiết bị', 'Tên thiết bị', 'Điểm sức khỏe (%)', 'Trạng thái', 'Dòng khởi động 50/51 (A)', 'Thời gian tác động Trip (ms)', 'Tỷ số biến dòng CT', 'Tiếp điểm cắt Trip', 'File đính kèm']);
      listToSync.filter(item => item.type === 'Rơ le bảo vệ').forEach(item => {
        const raw = item.measurements || {};
        relaysData.push([
          item.date || new Date().toLocaleDateString('vi-VN'), item.customer || '', item.factory || '', item.location || '',
          item.equipmentId || item.id, item.equipmentName || item.name, item.health || 98, item.status || 'healthy',
          raw.pickupCurrent || '5.0', raw.tripTimeMs || '35', raw.ctRatio || '300/5A', raw.tripContact || 'Pass', item.fileUrl || ''
        ]);
      });

      const spreadsheetId = localStorage.getItem('tev_spreadsheet_id') || '';
      const response = await fetch('/api/sheets/sync-export', {
        method: 'POST',
        headers: {
          ...getAuthHeaders(),
          ...(spreadsheetId ? { 'x-spreadsheet-id': spreadsheetId } : {})
        },
        body: JSON.stringify({
          spreadsheetId,
          transformers,
          switchgears,
          motors,
          generators,
          breakersData,
          cablesData,
          batteriesData,
          relaysData,
          inverters,
          cmmsData,
          tevServiceFlatformData: cmmsData,
          equipmentData,
          inventoryData,
          customersData
        })
      });

      if (!response.ok) {
        const err = await response.json();
        if (response.status === 401) {
          setIsGoogleConnected(false);
          throw new Error('Phiên đăng nhập đã hết hạn. Vui lòng tải lại trang và kết nối lại Google Drive.');
        }
        if (err.error && err.error.includes('SPREADSHEET_ID')) {
          throw new Error('Chưa cấu hình SPREADSHEET_ID. Vui lòng vào Settings -> Secrets để thêm ID của file Google Sheets.');
        }
        if (err.error && (err.error.includes('not found') || err.error.includes('Requested entity was not found'))) {
          throw new Error('Không tìm thấy file Google Sheets. Vui lòng kiểm tra lại SPREADSHEET_ID trong Settings -> Secrets (chỉ lấy phần ID, không lấy cả đường link).');
        }
        throw new Error(err.error || 'Failed to sync');
      }

      const resJson = await response.json();
      if (resJson.spreadsheetId) {
        localStorage.setItem('tev_spreadsheet_id', resJson.spreadsheetId);
        setGoogleSheetUrl(`https://docs.google.com/spreadsheets/d/${resJson.spreadsheetId}/edit`);
      }

      setSyncSuccess(true);
      setTimeout(() => setSyncSuccess(false), 3000);
      if (!silent) alert('Đồng bộ toàn bộ dữ liệu & biểu đồ mô phỏng lên file Google Sheet "TEV Service Flatform" thành công!');
    } catch (error: any) {
      console.error('Sync error:', error);
      const errorMessage = error.message === 'Failed to fetch' 
        ? 'Không thể kết nối với máy chủ. Vui lòng kiểm tra kết nối mạng hoặc thử lại sau vài giây.' 
        : error.message;
      if (!silent) alert('Lỗi khi đồng bộ dữ liệu: ' + errorMessage);
    } finally {
      setIsSyncing(false);
    }
  };

  /**
   * FLOW 1: App ➔ Google Sheet ➔ Firestore
   * Đẩy toàn bộ dữ liệu từ App lên Google Sheet và lưu trữ an toàn vào Cloud Firestore
   */
  const handleSyncAppToSheetsAndFirestore = async (silent: boolean = false) => {
    if (!isGoogleConnected) {
      if (!silent) alert('Vui lòng kết nối Google Drive trước khi thực hiện đồng bộ.');
      return;
    }
    setIsSyncing(true);
    try {
      // 1. Đẩy dữ liệu lên Google Sheets (force: true để không bị gián đoạn bởi popup đối soát)
      await handleSyncToSheets(undefined, true, true);

      // 2. Lưu đồng bộ toàn bộ Work Orders & Khách hàng vào Cloud Firestore
      if (auth.currentUser) {
        const promises: Promise<any>[] = [];
        workOrders.forEach(wo => {
          if (wo.id) {
            promises.push(setDoc(doc(db, 'workOrders', wo.id), wo, { merge: true }));
          }
        });
        customers.forEach(c => {
          if (c.id) {
            promises.push(setDoc(doc(db, 'customers', c.id), c, { merge: true }));
          }
        });
        await Promise.allSettled(promises);
      }

      const nowTime = new Date().toLocaleTimeString('vi-VN');
      setLastAutoSyncTime(nowTime);
      setTwoWaySyncNotice({
        type: 'success',
        message: `Đã hoàn tất luồng [App ➔ Google Sheet ➔ Firestore] lúc ${nowTime}! Đã đồng bộ ${workOrders.length} phiếu WO và ${customers.length} khách hàng.`
      });
      setTimeout(() => setTwoWaySyncNotice(null), 6000);
    } catch (err: any) {
      console.error('Flow App -> Sheet -> Firestore error:', err);
      setTwoWaySyncNotice({
        type: 'error',
        message: `Lỗi đồng bộ [App ➔ Sheet ➔ Firestore]: ${err.message || String(err)}`
      });
      setTimeout(() => setTwoWaySyncNotice(null), 7000);
    } finally {
      setIsSyncing(false);
    }
  };

  /**
   * FLOW 2: Google Sheet ➔ Firestore ➔ App
   * Đọc dữ liệu từ Google Sheet, lưu vào Cloud Firestore và nạp vào state App
   */
  const handleSyncSheetsToFirestoreAndApp = async (silent: boolean = false) => {
    if (!isGoogleConnected) {
      if (!silent) alert('Vui lòng kết nối Google Drive trước khi đọc dữ liệu.');
      return;
    }
    setIsSyncing(true);
    try {
      // Nạp từ Sheet, tự động ghi vào Firestore và cập nhật App
      await handleFetchFromSheets(true, true);
      const nowTime = new Date().toLocaleTimeString('vi-VN');
      setLastAutoSyncTime(nowTime);
      setTwoWaySyncNotice({
        type: 'success',
        message: `Đã hoàn tất luồng [Google Sheet ➔ Firestore ➔ App] lúc ${nowTime}! Dữ liệu mới nhất đã được nạp về Firestore và cập nhật trên App.`
      });
      setTimeout(() => setTwoWaySyncNotice(null), 6000);
    } catch (err: any) {
      console.error('Flow Sheet -> Firestore -> App error:', err);
      setTwoWaySyncNotice({
        type: 'error',
        message: `Lỗi nạp [Sheet ➔ Firestore ➔ App]: ${err.message || String(err)}`
      });
      setTimeout(() => setTwoWaySyncNotice(null), 7000);
    } finally {
      setIsSyncing(false);
    }
  };

  /**
   * Thực hiện luồng dữ liệu 2 chiều tự động (Bidirectional 2-Way Sync) giữa App và Google Sheets 'TEV Service Flatform':
   * Chạy ngầm êm ái, bảo vệ dữ liệu với Newer Wins mà không làm phiền người dùng bằng popup
   */
  const handlePerformTwoWaySync = async (silent: boolean = false) => {
    if (!isGoogleConnected) {
      if (!silent) alert('Vui lòng kết nối Google Drive trước khi đồng bộ 2 chiều.');
      return;
    }
    setAutoSyncStatus('syncing');
    try {
      // Step 1: Export full app data to Google Sheets (force: true để không chặn ngầm)
      await handleSyncToSheets(undefined, true, true);

      // Step 2: Fetch latest changes from Google Sheets into App and save to Firestore
      await handleFetchFromSheets(true, true);

      const nowTime = new Date().toLocaleTimeString('vi-VN');
      setLastAutoSyncTime(nowTime);
      setAutoSyncStatus('synced');
      setTwoWaySyncNotice({
        type: 'success',
        message: `Đã đồng bộ 2 chiều tự động thành công giữa Ứng Dụng ↔ Firestore ↔ Google Sheet lúc ${nowTime}!`
      });
      setTimeout(() => setTwoWaySyncNotice(null), 5000);
    } catch (err: any) {
      console.error('Two-way sync error:', err);
      setAutoSyncStatus('error');
      if (!silent) {
        alert('Lỗi khi đồng bộ 2 chiều: ' + (err.message || 'Lỗi không xác định'));
      }
    }
  };

  // Tự động đồng bộ 2 chiều (Background Auto 2-Way Sync every 60s)
  useEffect(() => {
    if (!isGoogleConnected || !isAutoSyncEnabled) return;

    // Trigger initial 2-way sync 3 seconds after connect
    const initialTimer = setTimeout(() => {
      handlePerformTwoWaySync(true);
    }, 3000);

    const interval = setInterval(() => {
      handlePerformTwoWaySync(true);
    }, 60000); // 60 seconds cycle

    return () => {
      clearTimeout(initialTimer);
      clearInterval(interval);
    };
  }, [isGoogleConnected, isAutoSyncEnabled]);

  // Auto resolve spreadsheet & discover 'TEV Service Flatform'
  useEffect(() => {
    if (isGoogleConnected) {
      fetch('/api/sheets/resolve-or-create', {
        headers: getAuthHeaders()
      })
      .then(r => r.json())
      .then(res => {
        if (res.spreadsheetId) {
          localStorage.setItem('tev_spreadsheet_id', res.spreadsheetId);
          setGoogleSheetUrl(`https://docs.google.com/spreadsheets/d/${res.spreadsheetId}/edit`);
        }
      })
      .catch(err => console.warn('Auto resolve sheet error:', err));
    }
  }, [isGoogleConnected]);

  const handleFetchFromSheets = async (silent: boolean = false, forceFromSheets: boolean = false) => {
    setIsSyncing(true);
    try {
      const response = await fetch('/api/sheets/get', {
        headers: getAuthHeaders()
      });
      
      if (!response.ok) {
        const err = await response.json();
        if (response.status === 401) {
          setIsGoogleConnected(false);
          throw new Error('Phiên đăng nhập đã hết hạn. Vui lòng tải lại trang và kết nối lại Google Drive.');
        }
        throw new Error(err.error || 'Failed to fetch data');
      }

      const { data } = await response.json();
      console.log('Fetched data from server:', data);
      
      const newEquipmentList: any[] = [];
      const newReportsList: any[] = [];
      
      // Helper to map common fields
      const getCommonIndices = (header: string[]) => ({
        dateIdx: getColumnIndex(header, ['Ngày', 'Date', 'Last Check', 'Thời gian', 'Time', 'Ngày kiểm tra']),
        customerIdx: getColumnIndex(header, ['Khách hàng', 'Customer', 'Client', 'Đơn vị']),
        factoryIdx: getColumnIndex(header, ['Nhà máy', 'Factory', 'Site', 'Trạm', 'Khu vực']),
        locationIdx: getColumnIndex(header, ['Vị trí', 'Location', 'Khu vực', 'Area', 'Ngăn lộ']),
        idIdx: getColumnIndex(header, ['Mã thiết bị', 'Equipment ID', 'Mã TB', 'Tag']),
        nameIdx: getColumnIndex(header, ['Tên thiết bị', 'Equipment Name', 'Tên TB', 'Description']),
        typeIdx: getColumnIndex(header, ['Loại thiết bị', 'Type', 'Loại TB', 'Phân loại']),
        healthIdx: getColumnIndex(header, ['Chỉ số sức khỏe', 'Health', 'HI', 'Sức khỏe', 'Điểm']),
        statusIdx: getColumnIndex(header, ['Trạng thái', 'Status', 'Tình trạng', 'Kết quả', 'Đánh giá']),
        ageIdx: getColumnIndex(header, ['Tuổi thọ', 'Age', 'Năm vận hành', 'Năm SX']),
        dutyFactorIdx: getColumnIndex(header, ['Hệ số làm việc', 'Duty Factor', 'Hệ số tải']),
        criticalityIdx: getColumnIndex(header, ['Độ quan trọng', 'Criticality', 'Phân loại rủi ro', 'Mức độ quan trọng']),
        linksIdx: getColumnIndex(header, ['File đính kèm', 'Links', 'Tài liệu', 'Link'])
      });

      const allSheetsData = data.allSheets || {};
      const sheetNames = Object.keys(allSheetsData);
      
      if (sheetNames.length === 0) {
        alert('Không tìm thấy sheet nào trong file Google Sheets.');
        setIsSyncing(false);
        return;
      }

      let syncedCustomersCount = 0;
      let syncedInventoryCount = 0;
      let syncedWoCount = 0;
      let protectedRecordsCount = 0;

      for (const sheetName of sheetNames) {
        const rows = allSheetsData[sheetName];
        if (!rows || rows.length === 0) continue;

        const headerIdx = findHeaderRow(rows);
        if (headerIdx === -1) continue;

        const header = rows[headerIdx];
        const idx = getCommonIndices(header);
        const lowerSheetName = sheetName.toLowerCase();

        // 1. Đồng bộ Work Orders (TEV Service Flatform / CMMS) có kiểm tra phiên bản
        if (lowerSheetName.includes('tev service flatform') || lowerSheetName.includes('work order') || (lowerSheetName.includes('cmms') && !lowerSheetName.includes('dashboard'))) {
          const dataRows = rows.slice(headerIdx + 1);
          const newWoList: any[] = [];
          for (const row of dataRows) {
            if (!row || row.length < 3) continue;
            const woId = row[0]?.toString().trim();
            if (!woId || woId.toLowerCase().includes('mã wo')) continue;

            const sheetUpdatedAt = row[18] || row[17] || '';
            const sheetTimestamp = parseVersionTimestamp(sheetUpdatedAt);

            // Kiểm tra phiên bản với Firestore / Local state
            const existingWo = workOrders.find(w => w.id?.toLowerCase() === woId.toLowerCase());
            const fsTimestamp = parseVersionTimestamp(existingWo?.updatedAt || existingWo?.createdAt);

            if (!forceFromSheets && existingWo && sheetTimestamp > 0 && fsTimestamp > sheetTimestamp && (fsTimestamp - sheetTimestamp > 3000)) {
              // Phiên bản trên Firestore mới hơn -> Bảo vệ Firestore, không ghi đè dữ liệu cũ từ Sheet
              protectedRecordsCount++;
              continue;
            }

            const customerVal = row[6] || '';
            const rawEq = row[4] ? row[4].toString().split(',').map((s: string) => s.trim()) : [];
            const eqFirst = rawEq[0] || '';
            const eqTitle = eqFirst.includes('MVCB') 
              ? `Tủ máy cắt trung thế ${eqFirst} 24kV` 
              : (eqFirst ? `Thiết bị ${eqFirst}` : '');

            const woObj: any = {
              id: woId,
              workPermitId: row[1] || '',
              title: row[2] || `Phiếu ${woId}`,
              description: row[3] || '',
              equipmentId: rawEq,
              equipmentName: eqTitle,
              failureCode: row[5] || '',
              customer: customerVal,
              customerId: customerVal,
              type: row[7] || '',
              isUnplanned: String(row[8] || '').toLowerCase() === 'có' || String(row[8] || '').toLowerCase() === 'true',
              priority: row[9] || 'Medium',
              status: row[10] || 'initiated',
              assignedTo: row[11] || '',
              responsibleApprove: row[12] || '',
              responsibleDo: row[13] || '',
              blockingRequired: String(row[14] || '').toLowerCase() === 'có',
              dueDate: row[15] || '',
              usedMaterials: row[16] || '',
              createdAt: row[17] || new Date().toISOString(),
              updatedAt: sheetUpdatedAt || new Date().toISOString(),
              downtimeStart: row[19] || '',
              repairStart: row[20] || '',
              repairEnd: row[21] || '',
              restartTime: row[22] || ''
            };
            newWoList.push(woObj);
            if (auth.currentUser) {
              try {
                await setDoc(doc(db, 'workOrders', woId), woObj, { merge: true });
                syncedWoCount++;
              } catch (err) {
                console.warn(`Firestore sync error for WO ${woId}:`, err);
              }
            } else {
              syncedWoCount++;
            }
          }

          if (newWoList.length > 0) {
            setWorkOrders(prev => {
              const map = new Map<string, any>(prev.map(w => [w.id?.toLowerCase(), w]));
              newWoList.forEach(w => map.set(w.id?.toLowerCase(), w));
              const all: any[] = Array.from(map.values());
              return all.sort((a: any, b: any) => {
                const timeA = parseVersionTimestamp(a.updatedAt || a.createdAt);
                const timeB = parseVersionTimestamp(b.updatedAt || b.createdAt);
                if (timeB !== timeA) return timeB - timeA;
                return (b.id || '').localeCompare(a.id || '');
              });
            });
          }
          continue;
        }

        // 2. Skip non-equipment sheets - Khách hàng có kiểm tra phiên bản
        if (lowerSheetName.includes('khachhang') || lowerSheetName.includes('khách hàng')) {
          const dataRows = rows.slice(headerIdx + 1);
          for (const row of dataRows) {
            if (!row || row.length < 2) continue;
            const id = row[0]?.toString().trim();
            const name = row[1]?.toString().trim();
            if (!id || !name || id.toLowerCase().includes('mã khách hàng')) continue;

            const factories = row[2] ? row[2].toString().split(',').map((s: string) => s.trim()).filter((s: string) => s !== '') : [];
            const email = row[3] || '';
            const phone = row[4] || '';
            const address = row[5] || '';
            const createdAt = row[6] || new Date().toISOString();
            const updatedAt = row[7] || row[6] || new Date().toISOString();

            // Kiểm tra phiên bản đối soát: Nếu Firestore mới hơn thì bỏ qua ghi đè
            const sheetTimestamp = parseVersionTimestamp(updatedAt || createdAt);
            const existingCust = customers.find(c => c.id?.toLowerCase() === id.toLowerCase());
            const fsTimestamp = parseVersionTimestamp(existingCust?.updatedAt || existingCust?.createdAt);
            if (existingCust && fsTimestamp > sheetTimestamp && (fsTimestamp - sheetTimestamp > 3000)) {
              protectedRecordsCount++;
              continue;
            }

            if (auth.currentUser) {
              try {
                await setDoc(doc(db, 'customers', id), {
                  name, factories, email, phone, address, createdAt, updatedAt
                }, { merge: true });
                syncedCustomersCount++;
              } catch (err) {
                console.warn(`Firestore sync error for customer ${id}:`, err);
                syncedCustomersCount++;
              }
            } else {
              syncedCustomersCount++;
            }
          }
          continue;
        }

        // 3. Handle Inventory sheets có kiểm tra phiên bản
        if (lowerSheetName.includes('quanlykho') || lowerSheetName.includes('kho')) {
          const dataRows = rows.slice(headerIdx + 1);
          const newInvList: any[] = [];
          for (const row of dataRows) {
            if (!row || row.length < 2) continue;
            const id = row[0]?.toString().trim();
            const name = row[1]?.toString().trim();
            if (!id || !name || id.toLowerCase().includes('mã vật tư') || id.toLowerCase().includes('id')) continue;

            const sku = row[2] || '';
            const category = row[3] || '';
            const quantity = parseFloat(row[4]?.toString().replace(',', '.') || '0') || 0;
            const unit = row[5] || '';
            const minStock = parseFloat(row[6]?.toString().replace(',', '.') || '0') || 0;
            const location = row[7] || '';
            const price = parseFloat(row[8]?.toString().replace(',', '.') || '0') || 0;
            const createdAt = row[9] || new Date().toISOString();
            const updatedAt = row[10] || row[9] || new Date().toISOString();

            // Kiểm tra phiên bản: Nếu Firestore mới hơn thì bảo vệ, không ghi đè
            const sheetTimestamp = parseVersionTimestamp(updatedAt || createdAt);
            const existingInv = inventory.find(i => i.id?.toLowerCase() === id.toLowerCase());
            const fsTimestamp = parseVersionTimestamp(existingInv?.updatedAt || existingInv?.createdAt);
            if (existingInv && fsTimestamp > sheetTimestamp && (fsTimestamp - sheetTimestamp > 3000)) {
              protectedRecordsCount++;
              continue;
            }

            newInvList.push({ id, name, sku, category, quantity, unit, minStock, location, price, createdAt, updatedAt });

            if (auth.currentUser) {
              try {
                await setDoc(doc(db, 'inventory', id), {
                  name, sku, category, quantity, unit, minStock, location, price, createdAt, updatedAt
                }, { merge: true });
                syncedInventoryCount++;
              } catch (err) {
                console.warn(`Firestore sync error for inventory item ${id}:`, err);
                syncedInventoryCount++;
              }
            } else {
              syncedInventoryCount++;
            }
          }

          if (newInvList.length > 0) {
            setInventory(prev => {
              const map = new Map(prev.map(i => [i.id, i]));
              newInvList.forEach(i => map.set(i.id, i));
              return Array.from(map.values());
            });
          }
          continue;
        }

        if (lowerSheetName.includes('cmms') || lowerSheetName.includes('dashboard') || lowerSheetName.includes('nhân sự') || lowerSheetName.includes('nhansu')) {
          continue;
        }

        // Determine equipment type
        let eqType: 'Máy biến áp' | 'Tủ điện' | 'Động cơ' | 'Inverter' = 'Máy biến áp';
        if (lowerSheetName.includes('tủ điện') || lowerSheetName.includes('trung thế') || lowerSheetName.includes('hạ thế') || lowerSheetName.includes('switchgear') || lowerSheetName.includes('swg') || lowerSheetName.includes('mcc') || lowerSheetName.includes('tev') || lowerSheetName.includes('rmu')) {
          eqType = 'Tủ điện';
        } else if (lowerSheetName.includes('động cơ') || lowerSheetName.includes('motor') || lowerSheetName.includes('máy bơm') || lowerSheetName.includes('pump') || lowerSheetName.includes('fan') || lowerSheetName.includes('quạt')) {
          eqType = 'Động cơ';
        } else if (lowerSheetName.includes('biến áp') || lowerSheetName.includes('mba') || lowerSheetName.includes('transformer')) {
          eqType = 'Máy biến áp';
        } else if (lowerSheetName.includes('inverter') || lowerSheetName.includes('biến tần') || lowerSheetName.includes('solar') || lowerSheetName.includes('wind')) {
          eqType = 'Inverter';
        } else {
          const hasTransformerCols = getColumnIndex(header, ['Nhiệt độ dầu', 'DGA', 'Furan']) !== -1;
          const hasSwitchgearCols = getColumnIndex(header, ['TEV', 'Siêu âm', 'SF6']) !== -1;
          const hasMotorCols = getColumnIndex(header, ['Rung động', 'Stator', 'Bearing']) !== -1;
          const hasInverterCols = getColumnIndex(header, ['MPPT', 'DC Input', 'AC Output', 'Hiệu suất']) !== -1;
          const hasEqId = getColumnIndex(header, ['Mã thiết bị', 'Mã TB', 'Equipment ID']) !== -1;
          
          if (hasMotorCols) eqType = 'Động cơ';
          else if (hasSwitchgearCols) eqType = 'Tủ điện';
          else if (hasInverterCols) eqType = 'Inverter';
          else if (hasTransformerCols) eqType = 'Máy biến áp';
          else if (hasEqId) eqType = 'Máy biến áp';
          else continue; // Skip sheet if no equipment indicators found
        }

        const dataRows = rows.slice(headerIdx + 1);
        
        if (eqType === 'Máy biến áp') {
          const mIdx = {
            oilTemp: getColumnIndex(header, ['Nhiệt độ dầu', 'Oil Temp', 'Temp dầu']),
            windingTemp: getColumnIndex(header, ['Nhiệt độ cuộn dây', 'Winding Temp', 'Temp cuộn dây']),
            irHighLow: getColumnIndex(header, ['Điện trở cách điện H-L', 'IR High-Low', 'IR Cao-Hạ']),
            irHighEarth: getColumnIndex(header, ['Điện trở cách điện H-E', 'IR High-Earth', 'IR Cao-Vỏ']),
            oilLeak: getColumnIndex(header, ['Rò rỉ dầu', 'Oil Leak']),
            dga: getColumnIndex(header, ['Phân tích khí hòa tan', 'DGA', 'Khí hòa tan']),
            dielectricStrength: getColumnIndex(header, ['Độ bền điện môi', 'Dielectric Strength']),
            furan: getColumnIndex(header, ['Hàm lượng Furan', 'Furan']),
            oilMoisture: getColumnIndex(header, ['Độ ẩm trong dầu', 'Oil Moisture', 'Độ ẩm dầu']),
            h2: getColumnIndex(header, ['H2']),
            o2: getColumnIndex(header, ['O2']),
            n2: getColumnIndex(header, ['N2']),
            ch4: getColumnIndex(header, ['CH4']),
            co: getColumnIndex(header, ['CO']),
            co2: getColumnIndex(header, ['CO2']),
            c2h4: getColumnIndex(header, ['C2H4']),
            c2h6: getColumnIndex(header, ['C2H6']),
            c2h2: getColumnIndex(header, ['C2H2'])
          };

          dataRows.forEach((row: any[]) => {
            if (!row || row.length === 0) return;
            const id = idx.idIdx !== -1 ? row[idx.idIdx] : (row[4] || row[0]);
            const name = idx.nameIdx !== -1 ? row[idx.nameIdx] : (row[5] || row[1]);
            if (!id || !name || id.toString().trim() === '' || id.toString().toLowerCase().includes('mã thiết bị')) return;

            let healthVal = idx.healthIdx !== -1 ? parseFloat(row[idx.healthIdx]?.toString().replace(',', '.') || '0') : 0;
            let statusVal = mapStatusFromSheet(idx.statusIdx !== -1 ? row[idx.statusIdx] : '');
            let age = idx.ageIdx !== -1 ? parseFloat(row[idx.ageIdx]?.toString().replace(',', '.') || '10') : 10;
            let dutyFactor = idx.dutyFactorIdx !== -1 ? parseFloat(row[idx.dutyFactorIdx]?.toString().replace(',', '.') || '1.0') : 1.0;
            if (isNaN(age)) age = 10;
            if (isNaN(dutyFactor)) dutyFactor = 1.0;

            let hasCritical = false;
            let hasWarning = false;
            const paramsToCheck = [
              { key: 'oilTemp', val: mIdx.oilTemp !== -1 ? row[mIdx.oilTemp] : '' },
              { key: 'windingTemp', val: mIdx.windingTemp !== -1 ? row[mIdx.windingTemp] : '' },
              { key: 'irHighLow', val: mIdx.irHighLow !== -1 ? row[mIdx.irHighLow] : '' },
              { key: 'irHighEarth', val: mIdx.irHighEarth !== -1 ? row[mIdx.irHighEarth] : '' },
              { key: 'oilLeak', val: mIdx.oilLeak !== -1 ? row[mIdx.oilLeak] : '' },
              { key: 'dga', val: mIdx.dga !== -1 ? row[mIdx.dga] : '' },
              { key: 'dielectricStrength', val: mIdx.dielectricStrength !== -1 ? row[mIdx.dielectricStrength] : '' },
              { key: 'furan', val: mIdx.furan !== -1 ? row[mIdx.furan] : '' },
              { key: 'oilMoisture', val: mIdx.oilMoisture !== -1 ? row[mIdx.oilMoisture] : '' },
              { key: 'h2', val: mIdx.h2 !== -1 ? row[mIdx.h2] : '' },
              { key: 'o2', val: mIdx.o2 !== -1 ? row[mIdx.o2] : '' },
              { key: 'n2', val: mIdx.n2 !== -1 ? row[mIdx.n2] : '' },
              { key: 'ch4', val: mIdx.ch4 !== -1 ? row[mIdx.ch4] : '' },
              { key: 'co', val: mIdx.co !== -1 ? row[mIdx.co] : '' },
              { key: 'co2', val: mIdx.co2 !== -1 ? row[mIdx.co2] : '' },
              { key: 'c2h4', val: mIdx.c2h4 !== -1 ? row[mIdx.c2h4] : '' },
              { key: 'c2h6', val: mIdx.c2h6 !== -1 ? row[mIdx.c2h6] : '' },
              { key: 'c2h2', val: mIdx.c2h2 !== -1 ? row[mIdx.c2h2] : '' }
            ];
            paramsToCheck.forEach(p => {
              const status = evaluateEquipmentParam('Máy biến áp', p.key, p.val);
              if (status === 'critical') hasCritical = true;
              if (status === 'warning') hasWarning = true;
            });

            const params: TransformerParams = {
              mainTransformer: { age, normalExpectedLife: 40, dutyFactor, locationFactor: 1.0, healthScoreFactor: hasCritical ? 1.5 : (hasWarning ? 1.2 : 1.0), reliabilityFactor: 1.0, healthScoreCap: 10, healthScoreCollar: 0.5, reliabilityCollar: 0.5 },
              tapchanger: { age, normalExpectedLife: 40, dutyFactor, locationFactor: 1.0, healthScoreFactor: hasCritical ? 1.5 : (hasWarning ? 1.2 : 1.0), reliabilityFactor: 1.0, healthScoreCap: 10, healthScoreCollar: 0.5, reliabilityCollar: 0.5 }
            };
            const result = calculateTransformerHealth(params);
            let calcHealth = Math.max(0, Math.round(100 - (result.score / 10) * 100));
            
            if (hasCritical) {
              statusVal = 'critical';
              if (calcHealth >= 60) calcHealth = 59; // Ensure health reflects critical status
            } else if (hasWarning) {
              statusVal = 'warning';
              if (calcHealth >= 80) calcHealth = 79; // Ensure health reflects warning status
              else if (calcHealth < 60) calcHealth = 60; // Ensure health doesn't drop to critical if only warning
            } else {
              statusVal = 'healthy';
              if (calcHealth < 80) calcHealth = 85; // Ensure health reflects healthy status
            }

            if (healthVal === 0 || isNaN(healthVal)) healthVal = calcHealth;
            else {
              // If health is provided but doesn't match status, adjust it
              if (statusVal === 'critical' && healthVal >= 60) healthVal = 59;
              if (statusVal === 'warning' && (healthVal >= 80 || healthVal < 60)) healthVal = 75;
              if (statusVal === 'healthy' && healthVal < 80) healthVal = 90;
            }

            const eqData = {
              lastCheck: (idx.dateIdx !== -1 ? row[idx.dateIdx] : '') || '',
              customer: (idx.customerIdx !== -1 ? row[idx.customerIdx] : '') || '',
              factory: (idx.factoryIdx !== -1 ? row[idx.factoryIdx] : '') || '',
              location: (idx.locationIdx !== -1 ? row[idx.locationIdx] : '') || '',
              criticality: idx.criticalityIdx !== -1 ? (row[idx.criticalityIdx]?.toString().trim().toUpperCase() || 'B') : 'B',
              id, name, type: 'Máy biến áp', health: healthVal, status: statusVal,
              measurements: {
                oilTemp: paramsToCheck[0].val, windingTemp: paramsToCheck[1].val, irHighLow: paramsToCheck[2].val,
                irHighEarth: paramsToCheck[3].val, oilLeak: paramsToCheck[4].val, dga: paramsToCheck[5].val,
                dielectricStrength: paramsToCheck[6].val, furan: paramsToCheck[7].val, oilMoisture: paramsToCheck[8].val,
                h2: paramsToCheck[9].val, o2: paramsToCheck[10].val, n2: paramsToCheck[11].val,
                ch4: paramsToCheck[12].val, co: paramsToCheck[13].val, co2: paramsToCheck[14].val,
                c2h4: paramsToCheck[15].val, c2h6: paramsToCheck[16].val, c2h2: paramsToCheck[17].val,
                age: age.toString(), dutyFactor: dutyFactor.toString()
              },
              rawData: [...row]
            };
            newEquipmentList.push(eqData);
            newReportsList.push({
              id: `REP-${id}-${Math.floor(Math.random() * 10000)}`,
              date: (idx.dateIdx !== -1 ? row[idx.dateIdx] : '') || new Date().toLocaleDateString('vi-VN'),
              lastCheck: (idx.dateIdx !== -1 ? row[idx.dateIdx] : '') || new Date().toLocaleDateString('vi-VN'),
              customer: (idx.customerIdx !== -1 ? row[idx.customerIdx] : '') || '',
              location: (idx.locationIdx !== -1 ? row[idx.locationIdx] : '') || '',
              equipmentId: id, equipmentName: name, factory: (idx.factoryIdx !== -1 ? row[idx.factoryIdx] : '') || '',
              type: 'Máy biến áp', inspector: 'FSE', status: statusVal, notes: 'Dữ liệu đồng bộ từ Google Sheets',
              measurements: eqData.measurements, fileUrl: (idx.linksIdx !== -1 ? row[idx.linksIdx] : '') || '#',
              rawData: [...row]
            });
          });
        } else if (eqType === 'Tủ điện') {
          const mIdx = {
            thermography: getColumnIndex(header, ['Nhiệt hồng ngoại', 'Thermography', 'Nhiệt độ tiếp xúc']),
            contactRes: getColumnIndex(header, ['Điện trở tiếp xúc', 'Contact Res']),
            tev: getColumnIndex(header, ['TEV']),
            ultrasonic: getColumnIndex(header, ['Siêu âm', 'Ultrasonic']),
            tevPulses: getColumnIndex(header, ['Xung TEV', 'TEV Pulses', 'Số xung TEV']),
            humidity: getColumnIndex(header, ['Độ ẩm', 'Humidity', 'Độ ẩm môi trường']),
            sf6Pressure: getColumnIndex(header, ['Áp suất SF6', 'SF6 Pressure'])
          };

          dataRows.forEach((row: any[]) => {
            if (!row || row.length === 0) return;
            const id = idx.idIdx !== -1 ? row[idx.idIdx] : (row[4] || row[0]);
            const name = idx.nameIdx !== -1 ? row[idx.nameIdx] : (row[5] || row[1]);
            if (!id || !name || id.toString().trim() === '' || id.toString().toLowerCase().includes('mã thiết bị')) return;

            let healthVal = idx.healthIdx !== -1 ? parseFloat(row[idx.healthIdx]?.toString().replace(',', '.') || '0') : 0;
            let statusVal = mapStatusFromSheet(idx.statusIdx !== -1 ? row[idx.statusIdx] : '');
            let age = idx.ageIdx !== -1 ? parseFloat(row[idx.ageIdx]?.toString().replace(',', '.') || '10') : 10;
            let dutyFactor = idx.dutyFactorIdx !== -1 ? parseFloat(row[idx.dutyFactorIdx]?.toString().replace(',', '.') || '1.0') : 1.0;
            if (isNaN(age)) age = 10;
            if (isNaN(dutyFactor)) dutyFactor = 1.0;

            let hasCritical = false;
            let hasWarning = false;
            const paramsToCheck = [
              { key: 'thermography', val: mIdx.thermography !== -1 ? row[mIdx.thermography] : '' },
              { key: 'contactRes', val: mIdx.contactRes !== -1 ? row[mIdx.contactRes] : '' },
              { key: 'tev', val: mIdx.tev !== -1 ? row[mIdx.tev] : '' },
              { key: 'ultrasonic', val: mIdx.ultrasonic !== -1 ? row[mIdx.ultrasonic] : '' },
              { key: 'tevPulses', val: mIdx.tevPulses !== -1 ? row[mIdx.tevPulses] : '' },
              { key: 'humidity', val: mIdx.humidity !== -1 ? row[mIdx.humidity] : '' },
              { key: 'sf6Pressure', val: mIdx.sf6Pressure !== -1 ? row[mIdx.sf6Pressure] : '' }
            ];
            paramsToCheck.forEach(p => {
              const status = evaluateEquipmentParam('Tủ điện', p.key, p.val);
              if (status === 'critical') hasCritical = true;
              if (status === 'warning') hasWarning = true;
            });

            const params: SwitchgearParams = {
              age, normalExpectedLife: 40, dutyFactor, locationFactor: 1.0, healthScoreFactor: 1.0, reliabilityFactor: 1.0, healthScoreCap: 10, healthScoreCollar: 0.5, reliabilityCollar: 0.5, observedFactor: 1.0, measuredFactor: hasCritical ? 1.5 : (hasWarning ? 1.2 : 1.0)
            };
            const result = calculateSwitchgearHealth(params);
            let calcHealth = Math.max(0, Math.round(100 - (result.score / 10) * 100));
            
            if (hasCritical) {
              statusVal = 'critical';
              if (calcHealth >= 60) calcHealth = 59;
            } else if (hasWarning) {
              statusVal = 'warning';
              if (calcHealth >= 80) calcHealth = 79;
              else if (calcHealth < 60) calcHealth = 60;
            } else {
              statusVal = 'healthy';
              if (calcHealth < 80) calcHealth = 85;
            }

            if (healthVal === 0 || isNaN(healthVal)) healthVal = calcHealth;
            else {
              if (statusVal === 'critical' && healthVal >= 60) healthVal = 59;
              if (statusVal === 'warning' && (healthVal >= 80 || healthVal < 60)) healthVal = 75;
              if (statusVal === 'healthy' && healthVal < 80) healthVal = 90;
            }

            const eqData = {
              lastCheck: (idx.dateIdx !== -1 ? row[idx.dateIdx] : '') || '',
              customer: (idx.customerIdx !== -1 ? row[idx.customerIdx] : '') || '',
              factory: (idx.factoryIdx !== -1 ? row[idx.factoryIdx] : '') || '',
              location: (idx.locationIdx !== -1 ? row[idx.locationIdx] : '') || '',
              criticality: idx.criticalityIdx !== -1 ? (row[idx.criticalityIdx]?.toString().trim().toUpperCase() || 'B') : 'B',
              id, name, type: 'Tủ điện', health: healthVal, status: statusVal,
              measurements: {
                thermography: paramsToCheck[0].val, contactRes: paramsToCheck[1].val, tev: paramsToCheck[2].val,
                ultrasonic: paramsToCheck[3].val, tevPulses: paramsToCheck[4].val, humidity: paramsToCheck[5].val,
                sf6Pressure: paramsToCheck[6].val, age: age.toString(), dutyFactor: dutyFactor.toString()
              },
              rawData: [...row]
            };
            newEquipmentList.push(eqData);
            newReportsList.push({
              id: `REP-${id}-${Math.floor(Math.random() * 10000)}`,
              date: (idx.dateIdx !== -1 ? row[idx.dateIdx] : '') || new Date().toLocaleDateString('vi-VN'),
              lastCheck: (idx.dateIdx !== -1 ? row[idx.dateIdx] : '') || new Date().toLocaleDateString('vi-VN'),
              customer: (idx.customerIdx !== -1 ? row[idx.customerIdx] : '') || '',
              location: (idx.locationIdx !== -1 ? row[idx.locationIdx] : '') || '',
              equipmentId: id, equipmentName: name, factory: (idx.factoryIdx !== -1 ? row[idx.factoryIdx] : '') || '',
              type: 'Tủ điện', inspector: 'FSE', status: statusVal, notes: 'Dữ liệu đồng bộ từ Google Sheets',
              measurements: eqData.measurements, fileUrl: (idx.linksIdx !== -1 ? row[idx.linksIdx] : '') || '#',
              rawData: [...row]
            });
          });
        } else if (eqType === 'Động cơ') {
          const mIdx = {
            vibration: getColumnIndex(header, ['Rung động', 'Vibration', 'Độ rung']),
            statorTemp: getColumnIndex(header, ['Nhiệt độ Stator', 'Stator Temp']),
            bearingTemp: getColumnIndex(header, ['Nhiệt độ ổ bi', 'Bearing Temp', 'Nhiệt độ vòng bi']),
            voltageImbalance: getColumnIndex(header, ['Mất cân bằng điện áp', 'Voltage Imbalance', 'Độ lệch điện áp']),
            tanDelta: getColumnIndex(header, ['Tan Delta']),
            tipUp: getColumnIndex(header, ['Tip Up']),
            pd: getColumnIndex(header, ['Phóng điện cục bộ', 'PD']),
            ir: getColumnIndex(header, ['Điện trở cách điện', 'IR']),
            pi: getColumnIndex(header, ['Chỉ số phân cực', 'PI']),
            dd: getColumnIndex(header, ['Phóng điện điện môi', 'DD']),
            elcid: getColumnIndex(header, ['ELCID'])
          };

          dataRows.forEach((row: any[]) => {
            if (!row || row.length === 0) return;
            const id = idx.idIdx !== -1 ? row[idx.idIdx] : (row[4] || row[0]);
            const name = idx.nameIdx !== -1 ? row[idx.nameIdx] : (row[5] || row[1]);
            if (!id || !name || id.toString().trim() === '' || id.toString().toLowerCase().includes('mã thiết bị')) return;

            let healthVal = idx.healthIdx !== -1 ? parseFloat(row[idx.healthIdx]?.toString().replace(',', '.') || '0') : 0;
            let statusVal = mapStatusFromSheet(idx.statusIdx !== -1 ? row[idx.statusIdx] : '');
            
            let hasCritical = false;
            let hasWarning = false;

            const singleParams = [
              { key: 'vibration', val: mIdx.vibration !== -1 ? row[mIdx.vibration] : '' },
              { key: 'statorTemp', val: mIdx.statorTemp !== -1 ? row[mIdx.statorTemp] : '' },
              { key: 'bearingTemp', val: mIdx.bearingTemp !== -1 ? row[mIdx.bearingTemp] : '' },
              { key: 'voltageImbalance', val: mIdx.voltageImbalance !== -1 ? row[mIdx.voltageImbalance] : '' }
            ];

            singleParams.forEach(p => {
              const status = evaluateEquipmentParam('Động cơ', p.key, p.val);
              if (status === 'critical') hasCritical = true;
              if (status === 'warning') hasWarning = true;
            });

            const parsePhaseData = (baseIdx: number): MotorPhaseData => {
              if (baseIdx === -1) return { R: null, Y: null, B: null };
              return {
                R: parseFloat(row[baseIdx]?.toString().replace(',', '.') || '0') || null,
                Y: parseFloat(row[baseIdx + 1]?.toString().replace(',', '.') || '0') || null,
                B: parseFloat(row[baseIdx + 2]?.toString().replace(',', '.') || '0') || null
              };
            };

            const tests: MotorDiagnosticTests = {
              ratedKV: 6600,
              tanDelta: parsePhaseData(mIdx.tanDelta),
              tipUp: parsePhaseData(mIdx.tipUp),
              pd: parsePhaseData(mIdx.pd),
              ir: parsePhaseData(mIdx.ir),
              pi: parsePhaseData(mIdx.pi),
              dd: parsePhaseData(mIdx.dd),
              elcid: parsePhaseData(mIdx.elcid)
            };

            // Check phase-based tests for critical/warning
            const phaseTests = [
              { key: 'tanDelta', vals: tests.tanDelta },
              { key: 'tipUp', vals: tests.tipUp },
              { key: 'pd', vals: tests.pd },
              { key: 'ir', vals: tests.ir },
              { key: 'pi', vals: tests.pi },
              { key: 'dd', vals: tests.dd },
              { key: 'elcid', vals: tests.elcid }
            ];

            phaseTests.forEach(test => {
              if (test.vals) {
                [test.vals.R, test.vals.Y, test.vals.B].forEach(v => {
                  if (v !== null) {
                    const status = evaluateEquipmentParam('Động cơ', test.key, v);
                    if (status === 'critical') hasCritical = true;
                    if (status === 'warning') hasWarning = true;
                  }
                });
              }
            });

            const result = calculateMotorHealth(tests);
            let calcHealth = Math.max(0, Math.round(result.hiPercentage));
            
            if (hasCritical) {
              statusVal = 'critical';
              if (calcHealth >= 60) calcHealth = 59;
            } else if (hasWarning) {
              statusVal = 'warning';
              if (calcHealth >= 80) calcHealth = 79;
              else if (calcHealth < 60) calcHealth = 60;
            } else {
              statusVal = 'healthy';
              if (calcHealth < 80) calcHealth = 85;
            }

            if (healthVal === 0 || isNaN(healthVal)) healthVal = calcHealth;
            else {
              if (statusVal === 'critical' && healthVal >= 60) healthVal = 59;
              if (statusVal === 'warning' && (healthVal >= 80 || healthVal < 60)) healthVal = 75;
              if (statusVal === 'healthy' && healthVal < 80) healthVal = 90;
            }

            let age = idx.ageIdx !== -1 ? parseFloat(row[idx.ageIdx]?.toString().replace(',', '.') || '10') : 10;
            let dutyFactor = idx.dutyFactorIdx !== -1 ? parseFloat(row[idx.dutyFactorIdx]?.toString().replace(',', '.') || '1.0') : 1.0;
            if (isNaN(age)) age = 10;
            if (isNaN(dutyFactor)) dutyFactor = 1.0;

            const eqData = {
              lastCheck: (idx.dateIdx !== -1 ? row[idx.dateIdx] : '') || '',
              customer: (idx.customerIdx !== -1 ? row[idx.customerIdx] : '') || '',
              factory: (idx.factoryIdx !== -1 ? row[idx.factoryIdx] : '') || '',
              location: (idx.locationIdx !== -1 ? row[idx.locationIdx] : '') || '',
              criticality: idx.criticalityIdx !== -1 ? (row[idx.criticalityIdx]?.toString().trim().toUpperCase() || 'B') : 'B',
              id, name, type: 'Động cơ', health: healthVal, status: statusVal,
              measurements: {
                vibration: singleParams[0].val, statorTemp: singleParams[1].val, bearingTemp: singleParams[2].val,
                voltageImbalance: singleParams[3].val,
                tanDeltaR: tests.tanDelta?.R, tanDeltaY: tests.tanDelta?.Y, tanDeltaB: tests.tanDelta?.B,
                tipUpR: tests.tipUp?.R, tipUpY: tests.tipUp?.Y, tipUpB: tests.tipUp?.B,
                pdR: tests.pd?.R, pdY: tests.pd?.Y, pdB: tests.pd?.B,
                irR: tests.ir?.R, irY: tests.ir?.Y, irB: tests.ir?.B,
                piR: tests.pi?.R, piY: tests.pi?.Y, piB: tests.pi?.B,
                ddR: tests.dd?.R, ddY: tests.dd?.Y, ddB: tests.dd?.B,
                elcidR: tests.elcid?.R, elcidY: tests.elcid?.Y, elcidB: tests.elcid?.B,
                age: age.toString(), dutyFactor: dutyFactor.toString()
              },
              rawData: [...row]
            };
            newEquipmentList.push(eqData);
            newReportsList.push({
              id: `REP-${id}-${Math.floor(Math.random() * 10000)}`,
              date: (idx.dateIdx !== -1 ? row[idx.dateIdx] : '') || new Date().toLocaleDateString('vi-VN'),
              lastCheck: (idx.dateIdx !== -1 ? row[idx.dateIdx] : '') || new Date().toLocaleDateString('vi-VN'),
              customer: (idx.customerIdx !== -1 ? row[idx.customerIdx] : '') || '',
              location: (idx.locationIdx !== -1 ? row[idx.locationIdx] : '') || '',
              equipmentId: id, equipmentName: name, factory: (idx.factoryIdx !== -1 ? row[idx.factoryIdx] : '') || '',
              type: 'Động cơ', inspector: 'FSE', status: statusVal, notes: 'Dữ liệu đồng bộ từ Google Sheets',
              measurements: eqData.measurements, fileUrl: (idx.linksIdx !== -1 ? row[idx.linksIdx] : '') || '#',
              rawData: [...row]
            });
          });
        } else if (eqType === 'Inverter') {
          const mIdx = {
            dcInputVoltage: getColumnIndex(header, ['Điện áp DC', 'DC Input Voltage', 'Voc max']),
            mpptVoltageRange: getColumnIndex(header, ['Dải MPPT', 'MPPT Voltage Range']),
            dcInputCurrent: getColumnIndex(header, ['Dòng điện DC', 'DC Input Current', 'Isc']),
            acOutputVoltage: getColumnIndex(header, ['Điện áp AC', 'AC Output Voltage']),
            acFrequency: getColumnIndex(header, ['Tần số AC', 'AC Frequency']),
            harmonicDistortion: getColumnIndex(header, ['Độ méo hài', 'Harmonic Distortion', 'THD']),
            maxEfficiency: getColumnIndex(header, ['Hiệu suất tối đa', 'Max Efficiency']),
            euroEfficiency: getColumnIndex(header, ['Hiệu suất Châu Âu', 'Euro Efficiency']),
            antiIslanding: getColumnIndex(header, ['Chống hòa lưới', 'Anti-islanding']),
            rcdIsolation: getColumnIndex(header, ['Dòng rò & Cách ly', 'RCD & Isolation']),
            dielectricVoltage: getColumnIndex(header, ['Điện áp chịu đựng', 'Dielectric Voltage']),
            physicalDisconnects: getColumnIndex(header, ['Cách ly vật lý', 'Physical Disconnects']),
            operatingTemp: getColumnIndex(header, ['Nhiệt độ vận hành', 'Operating Temp']),
            airFilters: getColumnIndex(header, ['Bộ lọc khí', 'Air Filters'])
          };

          dataRows.forEach((row: any[]) => {
            if (!row || row.length === 0) return;
            const id = idx.idIdx !== -1 ? row[idx.idIdx] : (row[4] || row[0]);
            const name = idx.nameIdx !== -1 ? row[idx.nameIdx] : (row[5] || row[1]);
            if (!id || !name || id.toString().trim() === '' || id.toString().toLowerCase().includes('mã thiết bị')) return;

            let healthVal = idx.healthIdx !== -1 ? parseFloat(row[idx.healthIdx]?.toString().replace(',', '.') || '0') : 0;
            let statusVal = mapStatusFromSheet(idx.statusIdx !== -1 ? row[idx.statusIdx] : '');
            
            let hasCritical = false;
            let hasWarning = false;

            const measurements: any = {};
            Object.entries(mIdx).forEach(([key, colIdx]) => {
              if (colIdx !== -1) {
                const val = row[colIdx];
                measurements[key] = val;
                const status = evaluateEquipmentParam('Inverter', key, val);
                if (status === 'critical') hasCritical = true;
                if (status === 'warning') hasWarning = true;
              }
            });

            let calcHealth = 100;
            if (hasCritical) {
              statusVal = 'critical';
              calcHealth = 55;
            } else if (hasWarning) {
              statusVal = 'warning';
              calcHealth = 75;
            } else {
              statusVal = 'healthy';
              calcHealth = 95;
            }

            if (healthVal === 0 || isNaN(healthVal)) healthVal = calcHealth;

            const eqData = {
              lastCheck: (idx.dateIdx !== -1 ? row[idx.dateIdx] : '') || '',
              customer: (idx.customerIdx !== -1 ? row[idx.customerIdx] : '') || '',
              factory: (idx.factoryIdx !== -1 ? row[idx.factoryIdx] : '') || '',
              location: (idx.locationIdx !== -1 ? row[idx.locationIdx] : '') || '',
              criticality: idx.criticalityIdx !== -1 ? (row[idx.criticalityIdx]?.toString().trim().toUpperCase() || 'B') : 'B',
              id, name, type: 'Inverter', health: healthVal, status: statusVal,
              measurements,
              rawData: [...row]
            };
            newEquipmentList.push(eqData);
            newReportsList.push({
              id: `REP-${id}-${Math.floor(Math.random() * 10000)}`,
              date: (idx.dateIdx !== -1 ? row[idx.dateIdx] : '') || new Date().toLocaleDateString('vi-VN'),
              lastCheck: (idx.dateIdx !== -1 ? row[idx.dateIdx] : '') || new Date().toLocaleDateString('vi-VN'),
              customer: (idx.customerIdx !== -1 ? row[idx.customerIdx] : '') || '',
              location: (idx.locationIdx !== -1 ? row[idx.locationIdx] : '') || '',
              equipmentId: id, equipmentName: name, factory: (idx.factoryIdx !== -1 ? row[idx.factoryIdx] : '') || '',
              type: 'Inverter', inspector: 'FSE', status: statusVal, notes: 'Dữ liệu đồng bộ từ Google Sheets',
              measurements: eqData.measurements, fileUrl: (idx.linksIdx !== -1 ? row[idx.linksIdx] : '') || '#',
              rawData: [...row]
            });
          });
        }
      }


      if (newEquipmentList.length > 0) {
        // Deduplicate equipment list (keep latest by date)
        const uniqueEqMap = new Map();
        newEquipmentList.forEach(eq => {
          const existing = uniqueEqMap.get(eq.id);
          if (!existing) {
            uniqueEqMap.set(eq.id, eq);
          } else {
            // Compare dates (DD/MM/YYYY)
            const parseDate = (dateStr: string) => {
              if (!dateStr) return 0;
              const parts = dateStr.split('/');
              if (parts.length === 3) {
                return new Date(`${parts[2]}-${parts[1]}-${parts[0]}`).getTime();
              }
              return new Date(dateStr).getTime() || 0;
            };
            const date1 = parseDate(eq.lastCheck);
            const date2 = parseDate(existing.lastCheck);
            if (date1 > date2) {
              uniqueEqMap.set(eq.id, eq);
            }
          }
        });
        
        const uniqueEquipmentList = Array.from(uniqueEqMap.values());
        
        // Sort reports by date descending
        newReportsList.sort((a, b) => {
          const parseDate = (dateStr: string) => {
            if (!dateStr) return 0;
            const parts = dateStr.split('/');
            if (parts.length === 3) {
              return new Date(`${parts[2]}-${parts[1]}-${parts[0]}`).getTime();
            }
            return new Date(dateStr).getTime() || 0;
          };
          return parseDate(b.date) - parseDate(a.date);
        });

        setAllEquipment(uniqueEquipmentList);
        setAllReports(newReportsList);
        setSyncSuccess(true);
        
        // Update siteName and customerName from the first equipment found if currently default
        if (uniqueEquipmentList.length > 0 && (siteName === 'Nhà máy Bắc Ninh' || !siteName)) {
          setSiteName(uniqueEquipmentList[0].factory);
          setCustomerName(uniqueEquipmentList[0].customer);
        }
        setTimeout(() => setSyncSuccess(false), 3000);
        
        let successMsg = 'Đã tải dữ liệu từ Google Sheets thành công!';
        if (syncedWoCount > 0) successMsg += ` (${syncedWoCount} phiếu WO)`;
        if (syncedCustomersCount > 0) successMsg += ` (${syncedCustomersCount} KH)`;
        if (syncedInventoryCount > 0) successMsg += ` (${syncedInventoryCount} vật tư)`;
        if (protectedRecordsCount > 0) {
          successMsg += `\n🛡️ Đã bảo vệ an toàn ${protectedRecordsCount} bản ghi Firestore có phiên bản mới hơn, không bị ghi đè!`;
        }
        if (!silent) alert(successMsg);
      } else if (syncedCustomersCount > 0 || syncedInventoryCount > 0 || syncedWoCount > 0) {
        let msg = '';
        if (syncedWoCount > 0) msg += `Đã đồng bộ ${syncedWoCount} phiếu WO. `;
        if (syncedCustomersCount > 0) msg += `Đã đồng bộ ${syncedCustomersCount} khách hàng. `;
        if (syncedInventoryCount > 0) msg += `Đã đồng bộ ${syncedInventoryCount} vật tư kho. `;
        if (protectedRecordsCount > 0) {
          msg += `\n🛡️ Đã bảo vệ an toàn ${protectedRecordsCount} bản ghi Firestore có phiên bản mới hơn.`;
        }
        if (!silent) alert(msg.trim());
      } else {
        let msg = 'Không tìm thấy dữ liệu mới trên Google Sheets.';
        if (protectedRecordsCount > 0) {
          msg = `Dữ liệu trên Firestore đã là mới nhất. Đã bảo vệ ${protectedRecordsCount} bản ghi không bị ghi đè dữ liệu cũ từ Sheet.`;
        }
        if (!silent) alert(msg);
      }
    } catch (error: any) {
      console.error('Fetch error:', error);
      const errorMessage = error.message === 'Failed to fetch' 
        ? 'Không thể kết nối với máy chủ. Vui lòng kiểm tra kết nối mạng hoặc thử lại sau vài giây.' 
        : error.message;
      if (!silent) alert('Lỗi khi tải dữ liệu từ Sheets: ' + errorMessage);
    } finally {
      setIsSyncing(false);
    }
  };

  // Auto fetch when connected
  useEffect(() => {
    if (isGoogleConnected) {
      handleFetchFromSheets();
    }
  }, [isGoogleConnected]);

  const handleSaveToSheets = async () => {
    if (!isGoogleConnected) {
      alert('Vui lòng kết nối Google Drive trước khi lưu.');
      return;
    }

    setIsSaving(true);
    setSaveSuccess(false);

    try {
      // 0. Generate Report ID and Date
      const newReportId = `REP-${new Date().getFullYear()}-${(new Date().getMonth() + 1).toString().padStart(2, '0')}-${Math.floor(Math.random() * 1000).toString().padStart(3, '0')}`;
      const reportDate = new Date().toLocaleDateString('vi-VN');

      // 1. Generate PDF Report
      const reportData = {
        id: newReportId,
        date: reportDate,
        equipmentId: equipmentCode,
        equipmentName: equipmentName,
        factory: siteName,
        type: selectedEqType,
        inspector: user?.displayName || 'FSE',
        status: healthResult?.status || 'healthy',
        notes: `Báo cáo kiểm tra định kỳ. ${attachedFiles.length > 0 ? 'Có file đính kèm.' : ''}`,
        measurements: formData
      };
      
      const historicalReports = allReports.filter(r => r.equipmentId === equipmentCode);
      const pdfBlob = await generateIndividualReportPDF(reportData, false, historicalReports) as Blob;
      const pdfFile = new File([pdfBlob], `${newReportId}.pdf`, { type: 'application/pdf' });

      // 2. Upload files to Drive (including the generated PDF)
      let attachmentLinks = '';
      const uploadFormData = new FormData();
      
      // Append user attached files
      if (attachedFiles.length > 0) {
        attachedFiles.forEach(file => {
          uploadFormData.append('files', file);
        });
      }
      
      // Append auto-generated PDF
      uploadFormData.append('files', pdfFile);

      const uploadRes = await fetch('/api/drive/upload', {
        method: 'POST',
        headers: getAuthHeaders(false),
        body: uploadFormData
      });

      if (uploadRes.ok) {
        const uploadData = await uploadRes.json();
        if (uploadData.success && uploadData.links) {
          attachmentLinks = uploadData.links.join('\n');
        }
      } else {
        console.error('Failed to upload files');
        alert('Có lỗi khi upload file đính kèm, nhưng dữ liệu vẫn sẽ được lưu.');
      }

      // 3. Update local reports state
      setAllReports(prev => [reportData, ...prev]);

      // 4. Prepare data row based on equipment type
      const now = new Date().toLocaleString('vi-VN');
      const baseData = [now, customerName, siteName, locationName, equipmentCode, equipmentName, selectedEqType, healthResult?.index || 'N/A', healthResult?.status || 'N/A'];
      
      let values: any[] = [];
      let range = 'Sheet1!A:Z'; // Default
      
      if (selectedEqType === 'Máy biến áp') {
        values = [
          ...baseData, 
          formData.oilTemp || '', 
          formData.windingTemp || '', 
          formData.irHighLow || '', 
          formData.irHighEarth || '', 
          formData.oilLeak || '',
          formData.dga || '',
          formData.dielectricStrength || '',
          formData.furan || '',
          formData.oilMoisture || '',
          formData.h2 || '',
          formData.o2 || '',
          formData.n2 || '',
          formData.ch4 || '',
          formData.co || '',
          formData.co2 || '',
          formData.c2h4 || '',
          formData.c2h6 || '',
          formData.c2h2 || '',
          formData.age || '',
          formData.dutyFactor || '',
          attachmentLinks
        ];
        range = 'Máy biến áp!A:Z';
      } else if (selectedEqType === 'Tủ điện trung thế' || selectedEqType === 'Tủ điện') {
        values = [
          ...baseData, 
          formData.thermography || '', 
          formData.contactRes || '', 
          formData.tev || '', 
          formData.ultrasonic || '', 
          formData.tevPulses || '',
          formData.humidity || '',
          formData.sf6Pressure || '',
          formData.age || '',
          formData.dutyFactor || '',
          attachmentLinks
        ];
        range = 'Tủ điện trung thế!A:Z';
      } else if (selectedEqType === 'Động cơ') {
        values = [
          ...baseData, 
          formData.vibration || '', 
          formData.statorTemp || '', 
          formData.ir || '', 
          formData.pd || '', 
          formData.voltageImbalance || '', 
          formData.pi || '',
          formData.bearingTemp || '',
          formData.tanDelta || '',
          formData.age || '',
          formData.dutyFactor || '',
          attachmentLinks
        ];
        range = 'Động cơ!A:Z';
      } else if (selectedEqType === 'Inverter') {
        values = [
          ...baseData,
          formData.dcInputVoltage || '',
          formData.mpptVoltageRange || '',
          formData.dcInputCurrent || '',
          formData.acOutputVoltage || '',
          formData.acFrequency || '',
          formData.harmonicDistortion || '',
          formData.maxEfficiency || '',
          formData.euroEfficiency || '',
          formData.antiIslanding || '',
          formData.rcdIsolation || '',
          formData.dielectricVoltage || '',
          formData.physicalDisconnects || '',
          formData.operatingTemp || '',
          formData.airFilters || '',
          attachmentLinks
        ];
        range = 'Inverter!A:Z';
      }

      const response = await fetch('/api/sheets/append', {
        method: 'POST',
        headers: getAuthHeaders(),
        body: JSON.stringify({
          range,
          values
        })
      });

      if (!response.ok) {
        const err = await response.json();
        if (response.status === 401) {
          setIsGoogleConnected(false);
          throw new Error('Phiên đăng nhập đã hết hạn. Vui lòng tải lại trang và kết nối lại Google Drive.');
        }
        if (err.error && err.error.includes('SPREADSHEET_ID')) {
          throw new Error('Chưa cấu hình SPREADSHEET_ID. Vui lòng vào Settings -> Secrets để thêm ID của file Google Sheets.');
        }
        if (err.error && (err.error.includes('not found') || err.error.includes('Requested entity was not found'))) {
          throw new Error('Không tìm thấy file Google Sheets. Vui lòng kiểm tra lại SPREADSHEET_ID trong Settings -> Secrets (chỉ lấy phần ID, không lấy cả đường link).');
        }
        throw new Error(err.error || 'Failed to save');
      }

      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 3000);
      
      // Add a new report to the reports list
      const newReport = {
        id: newReportId,
        date: reportDate,
        equipmentId: equipmentCode,
        equipmentName: equipmentName,
        factory: siteName,
        type: selectedEqType,
        inspector: user?.displayName || 'FSE',
        status: healthResult?.status || 'healthy',
        notes: `Báo cáo kiểm tra định kỳ. ${attachmentLinks ? 'Có file đính kèm.' : ''}`,
        fileUrl: attachmentLinks || '#'
      };
      setAllReports(prev => [newReport, ...prev]);

      setAttachedFiles([]);
      setFormData({});
      setHealthResult(null);
      
      // Fetch updated data to reflect the new entry
      handleFetchFromSheets();
    } catch (error: any) {
      console.error('Save error:', error);
      alert('Lỗi khi lưu dữ liệu: ' + error.message);
    } finally {
      setIsSaving(false);
    }
  };

  // Calculate overall Health Index
  useEffect(() => {
    if (Object.keys(formData).length === 0) {
      setHealthResult(null);
      return;
    }

    let finalScore = 0;
    let finalStatus = 'healthy';

    if (selectedEqType === 'Máy biến áp') {
      const age = parseFloat(formData.age);
      const dutyFactor = parseFloat(formData.dutyFactor);
      if (!isNaN(age) && !isNaN(dutyFactor)) {
        let hasCritical = false;
        let hasWarning = false;
        Object.keys(formData).forEach(key => {
          if (key !== 'age' && key !== 'dutyFactor') {
            const status = evaluateEquipmentParam(selectedEqType, key, formData[key]);
            if (status === 'critical') hasCritical = true;
            if (status === 'warning') hasWarning = true;
          }
        });
        const healthScoreFactor = hasCritical ? 1.5 : (hasWarning ? 1.2 : 1.0);
        
        const params: TransformerParams = {
          mainTransformer: {
            age,
            normalExpectedLife: 40,
            dutyFactor,
            locationFactor: 1.0,
            healthScoreFactor: healthScoreFactor,
            reliabilityFactor: 1.0,
            healthScoreCap: 10,
            healthScoreCollar: 0.5,
            reliabilityCollar: 0.5
          },
          tapchanger: {
            age,
            normalExpectedLife: 40,
            dutyFactor,
            locationFactor: 1.0,
            healthScoreFactor: healthScoreFactor,
            reliabilityFactor: 1.0,
            healthScoreCap: 10,
            healthScoreCollar: 0.5,
            reliabilityCollar: 0.5
          }
        };
        const result = calculateTransformerHealth(params);
        finalScore = Math.max(0, Math.round(100 - (result.score / 10) * 100));
        finalStatus = (result.banding === 'HI1' || result.banding === 'HI2') ? 'healthy' : (result.banding === 'HI3' || result.banding === 'HI4') ? 'warning' : 'critical';
      } else {
        // Fallback to old logic if age/dutyFactor not provided
        let totalScore = 0;
        let count = 0;
        let hasCritical = false;
        let hasWarning = false;
        Object.keys(formData).forEach(key => {
          const status = evaluateParam(selectedEqType, key, formData[key]);
          if (status) {
            count++;
            if (status === 'healthy') totalScore += 100;
            if (status === 'warning') { totalScore += 60; hasWarning = true; }
            if (status === 'critical') { totalScore += 20; hasCritical = true; }
          }
        });
        if (count > 0) {
          finalScore = Math.round(totalScore / count);
          if (hasCritical || finalScore < 60) finalStatus = 'critical';
          else if (hasWarning || finalScore < 80) finalStatus = 'warning';
        } else {
          setHealthResult(null);
          return;
        }
      }
    } else if (selectedEqType === 'Tủ điện trung thế' || selectedEqType === 'Tủ điện') {
      const age = parseFloat(formData.age);
      const dutyFactor = parseFloat(formData.dutyFactor);
      if (!isNaN(age) && !isNaN(dutyFactor)) {
        let hasCritical = false;
        let hasWarning = false;
        Object.keys(formData).forEach(key => {
          if (key !== 'age' && key !== 'dutyFactor') {
            const status = evaluateEquipmentParam(selectedEqType, key, formData[key]);
            if (status === 'critical') hasCritical = true;
            if (status === 'warning') hasWarning = true;
          }
        });
        const measuredFactor = hasCritical ? 1.5 : (hasWarning ? 1.2 : 1.0);
        
        const params: SwitchgearParams = {
          age,
          normalExpectedLife: 40,
          dutyFactor,
          locationFactor: 1.0,
          healthScoreFactor: 1.0,
          reliabilityFactor: 1.0,
          healthScoreCap: 10,
          healthScoreCollar: 0.5,
          reliabilityCollar: 0.5,
          observedFactor: 1.0,
          measuredFactor: measuredFactor
        };
        const result = calculateSwitchgearHealth(params);
        finalScore = Math.max(0, Math.round(100 - (result.score / 10) * 100));
        finalStatus = (result.banding === 'HI1' || result.banding === 'HI2') ? 'healthy' : (result.banding === 'HI3' || result.banding === 'HI4') ? 'warning' : 'critical';
      } else {
        // Fallback
        let totalScore = 0;
        let count = 0;
        let hasCritical = false;
        let hasWarning = false;
        Object.keys(formData).forEach(key => {
          const status = evaluateParam(selectedEqType, key, formData[key]);
          if (status) {
            count++;
            if (status === 'healthy') totalScore += 100;
            if (status === 'warning') { totalScore += 60; hasWarning = true; }
            if (status === 'critical') { totalScore += 20; hasCritical = true; }
          }
        });
        if (count > 0) {
          finalScore = Math.round(totalScore / count);
          if (hasCritical || finalScore < 60) finalStatus = 'critical';
          else if (hasWarning || finalScore < 80) finalStatus = 'warning';
        } else {
          setHealthResult(null);
          return;
        }
      }
    } else if (selectedEqType === 'Động cơ') {
      const tests: MotorDiagnosticTests = {
        ratedKV: 6600,
        ir: formData.ir,
        pd: formData.pd,
        pi: formData.pi,
        tanDelta: formData.tanDelta,
        tipUp: formData.tipUp,
        dd: formData.dd,
        elcid: formData.elcid
      };
      
      const result = calculateMotorHealth(tests);
      if (result.hiPercentage > 0) {
        finalScore = Math.round(result.hiPercentage);
        finalStatus = (result.banding === 'HI1' || result.banding === 'HI2') ? 'healthy' : (result.banding === 'HI3' || result.banding === 'HI4') ? 'warning' : 'critical';
      } else {
        // Fallback for other parameters if diagnostic tests are missing
        let totalScore = 0;
        let count = 0;
        let hasCritical = false;
        let hasWarning = false;
        const otherParams = ['vibration', 'statorTemp', 'bearingTemp', 'voltageImbalance'];
        otherParams.forEach(key => {
          const status = evaluateParam(selectedEqType, key, formData[key]);
          if (status) {
            count++;
            if (status === 'healthy') totalScore += 100;
            if (status === 'warning') { totalScore += 60; hasWarning = true; }
            if (status === 'critical') { totalScore += 20; hasCritical = true; }
          }
        });
        if (count > 0) {
          finalScore = Math.round(totalScore / count);
          if (hasCritical || finalScore < 60) finalStatus = 'critical';
          else if (hasWarning || finalScore < 80) finalStatus = 'warning';
        } else {
          setHealthResult(null);
          return;
        }
      }
    }

    setHealthResult({ index: finalScore, status: finalStatus });
  }, [formData, selectedEqType]);

  const ParamInput = ({ 
    title, 
    unit, 
    icon: Icon, 
    iconColor, 
    standard, 
    prevValue, 
    prevTrend,
    onHistoryClick,
    value,
    onChange,
    evalStatus
  }: any) => {
    let inputBorder = 'border-slate-300 focus:border-blue-500 focus:ring-blue-200 bg-white';
    let textCol = 'text-slate-900';
    if (evalStatus === 'warning') {
      inputBorder = 'border-amber-400 focus:border-amber-500 focus:ring-amber-200 bg-amber-50';
      textCol = 'text-amber-900';
    } else if (evalStatus === 'critical') {
      inputBorder = 'border-rose-400 focus:border-rose-500 focus:ring-rose-200 bg-rose-50';
      textCol = 'text-rose-900';
    } else if (evalStatus === 'healthy') {
      inputBorder = 'border-emerald-400 focus:border-emerald-500 focus:ring-emerald-200 bg-emerald-50';
      textCol = 'text-emerald-900';
    }

    return (
      <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
        <div className="flex items-center justify-between mb-3">
          <label className="text-sm font-semibold text-slate-800">{title} {unit && `(${unit})`}</label>
          <button 
            onClick={onHistoryClick}
            className="text-blue-600 bg-blue-100/50 p-1.5 rounded hover:bg-blue-100 transition-colors flex items-center gap-1 text-xs font-medium"
          >
            <BarChartIcon size={14} />
            Lịch sử
          </button>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <input 
              type="number" 
              value={value || ''}
              onChange={(e) => onChange(e.target.value)}
              placeholder="Nhập giá trị..." 
              className={`w-full px-4 py-2.5 ${inputBorder} ${textCol} rounded-lg text-base outline-none font-medium transition-colors`} 
            />
            {evalStatus && (
              <div className={`text-xs mt-1.5 font-medium ${evalStatus === 'healthy' ? 'text-emerald-600' : evalStatus === 'warning' ? 'text-amber-600' : 'text-rose-600'}`}>
                {evalStatus === 'healthy' ? '✓ Đạt tiêu chuẩn' : evalStatus === 'warning' ? '⚠ Cảnh báo' : '⚠ Vượt ngưỡng nguy hiểm'}
              </div>
            )}
          </div>
          <div className="flex flex-col justify-center space-y-1.5 bg-white p-2.5 rounded-lg border border-slate-200 shadow-sm">
            <div className="flex justify-between items-center text-xs">
              <span className="text-slate-500">Tiêu chuẩn:</span>
              <span className="font-semibold text-emerald-600">{standard}</span>
            </div>
            <div className="flex justify-between items-center text-xs">
              <span className="text-slate-500">Lần trước:</span>
              <span className={`font-semibold flex items-center gap-0.5 ${prevTrend === 'up' ? 'text-rose-600' : prevTrend === 'down' ? 'text-emerald-600' : 'text-slate-700'}`}>
                {prevValue} {prevTrend === 'up' ? <ArrowUpRight size={14} /> : prevTrend === 'down' ? <ArrowDownRight size={14} /> : <Minus size={14} />}
              </span>
            </div>
          </div>
        </div>
      </div>
    );
  };

  const ThreePhaseParamInput = ({ 
    title, 
    unit, 
    standard, 
    prevValue, 
    prevTrend,
    onHistoryClick,
    values = {}, 
    onChange, 
    evalStatus = {} 
  }: any) => {
    const getStatusColor = (status: string) => {
      if (status === 'warning') return 'border-amber-400 bg-amber-50 text-amber-900';
      if (status === 'critical') return 'border-rose-400 bg-rose-50 text-rose-900';
      if (status === 'healthy') return 'border-emerald-400 bg-emerald-50 text-emerald-900';
      return 'border-slate-300 bg-white text-slate-900';
    };

    return (
      <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
        <div className="flex items-center justify-between mb-3">
          <label className="text-sm font-semibold text-slate-800">{title} {unit && `(${unit})`}</label>
          <button 
            onClick={onHistoryClick}
            className="text-blue-600 bg-blue-100/50 p-1.5 rounded hover:bg-blue-100 transition-colors flex items-center gap-1 text-xs font-medium"
          >
            <BarChartIcon size={14} />
            Lịch sử
          </button>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-3">
          {['R', 'Y', 'B'].map(phase => (
            <div key={phase}>
              <div className="text-[10px] font-bold text-slate-500 mb-1 uppercase">Pha {phase}</div>
              <input 
                type="number" 
                value={values[phase] || ''}
                onChange={(e) => onChange(phase, e.target.value)}
                placeholder={`${phase}...`} 
                className={`w-full px-3 py-2 ${getStatusColor(evalStatus[phase])} rounded-lg text-sm outline-none font-medium transition-colors border focus:ring-2 focus:ring-blue-200`} 
              />
            </div>
          ))}
        </div>
        <div className="flex flex-col justify-center space-y-1.5 bg-white p-2.5 rounded-lg border border-slate-200 shadow-sm">
          <div className="flex justify-between items-center text-xs">
            <span className="text-slate-500">Tiêu chuẩn:</span>
            <span className="font-semibold text-emerald-600">{standard}</span>
          </div>
          <div className="flex justify-between items-center text-xs">
            <span className="text-slate-500">Lần trước:</span>
            <span className={`font-semibold flex items-center gap-0.5 ${prevTrend === 'up' ? 'text-rose-600' : prevTrend === 'down' ? 'text-emerald-600' : 'text-slate-700'}`}>
              {prevValue} {prevTrend === 'up' ? <ArrowUpRight size={14} /> : prevTrend === 'down' ? <ArrowDownRight size={14} /> : <Minus size={14} />}
            </span>
          </div>
        </div>
      </div>
    );
  };

  const renderTransformerParams = () => (
    <>
      {/* Thông số cơ bản (Tuổi & Hệ số tải) */}
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Info size={16} className="text-slate-500" />
          Thông số cơ bản
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <ParamInput 
            title="Tuổi thiết bị (Age)" unit="năm" standard="< 40 năm" prevValue="10 năm" prevTrend="stable"
            value={formData.age} onChange={(val: any) => setFormData({...formData, age: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('Tuổi thiết bị'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Hệ số tải (Duty Factor)" unit="" standard="1.0" prevValue="1.0" prevTrend="stable"
            value={formData.dutyFactor} onChange={(val: any) => setFormData({...formData, dutyFactor: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('Hệ số tải'); setShowHistoryModal(true); }}
          />
        </div>
      </div>

      {/* Parameter Group 1 */}
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Thermometer size={16} className="text-rose-500" />
          Nhiệt độ & Môi trường
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <ParamInput 
            title="Nhiệt độ dầu" unit="°C" standard="≤ 90 °C" prevValue="85 °C" prevTrend="up"
            value={formData.oilTemp} onChange={(val: any) => setFormData({...formData, oilTemp: val})}
            evalStatus={evaluateParam('Máy biến áp', 'oilTemp', formData.oilTemp)}
            onHistoryClick={() => { setSelectedParamHistory('Nhiệt độ dầu'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Nhiệt độ cuộn dây" unit="°C" standard="≤ 105 °C" prevValue="92 °C" prevTrend="stable"
            value={formData.windingTemp} onChange={(val: any) => setFormData({...formData, windingTemp: val})}
            evalStatus={evaluateParam('Máy biến áp', 'windingTemp', formData.windingTemp)}
            onHistoryClick={() => { setSelectedParamHistory('Nhiệt độ cuộn dây'); setShowHistoryModal(true); }}
          />
        </div>
      </div>

      {/* Parameter Group 2 */}
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Zap size={16} className="text-amber-500" />
          Điện trở cách điện
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <ParamInput 
            title="Cao áp - Hạ áp" unit="MΩ" standard="≥ 1000 MΩ" prevValue="1250 MΩ" prevTrend="down"
            value={formData.irHighLow} onChange={(val: any) => setFormData({...formData, irHighLow: val})}
            evalStatus={evaluateParam('Máy biến áp', 'irHighLow', formData.irHighLow)}
            onHistoryClick={() => { setSelectedParamHistory('Điện trở cách điện (Cao-Hạ)'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Cao áp - Vỏ" unit="MΩ" standard="≥ 1000 MΩ" prevValue="1400 MΩ" prevTrend="stable"
            value={formData.irHighEarth} onChange={(val: any) => setFormData({...formData, irHighEarth: val})}
            evalStatus={evaluateParam('Máy biến áp', 'irHighEarth', formData.irHighEarth)}
            onHistoryClick={() => { setSelectedParamHistory('Điện trở cách điện (Cao-Vỏ)'); setShowHistoryModal(true); }}
          />
        </div>
      </div>

      {/* Visual Inspection */}
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <CheckCircle size={16} className="text-emerald-500" />
          Kiểm tra ngoại quan
        </h3>
        <div className="p-4 bg-slate-50 rounded-xl border border-slate-200">
          <div className="flex items-center justify-between mb-3">
            <label className="text-sm font-semibold text-slate-800">Tình trạng rò rỉ dầu</label>
            <button 
              onClick={() => { setSelectedParamHistory('Tình trạng rò rỉ dầu'); setShowHistoryModal(true); }}
              className="text-blue-600 bg-blue-100/50 p-1.5 rounded hover:bg-blue-100 transition-colors flex items-center gap-1 text-xs font-medium"
            >
              <History size={14} />
              Lịch sử
            </button>
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <select 
                value={formData.oilLeak || ''}
                onChange={(e) => setFormData({...formData, oilLeak: e.target.value})}
                className={`w-full px-4 py-2.5 bg-white border ${
                  evaluateParam('Máy biến áp', 'oilLeak', formData.oilLeak) === 'warning' ? 'border-amber-400 bg-amber-50 text-amber-900' :
                  evaluateParam('Máy biến áp', 'oilLeak', formData.oilLeak) === 'critical' ? 'border-rose-400 bg-rose-50 text-rose-900' :
                  evaluateParam('Máy biến áp', 'oilLeak', formData.oilLeak) === 'healthy' ? 'border-emerald-400 bg-emerald-50 text-emerald-900' :
                  'border-slate-300 text-slate-900'
                } focus:border-blue-500 focus:ring-2 focus:ring-blue-200 rounded-lg text-base outline-none appearance-none font-medium transition-colors`}
              >
                <option value="" disabled>Chọn tình trạng...</option>
                <option value="normal">Bình thường (Không rò rỉ)</option>
                <option value="light">Rò rỉ nhẹ (Thấm dầu)</option>
                <option value="heavy">Rò rỉ nặng (Nhỏ giọt)</option>
              </select>
              {formData.oilLeak && (
                <div className={`text-xs mt-1.5 font-medium ${
                  evaluateParam('Máy biến áp', 'oilLeak', formData.oilLeak) === 'healthy' ? 'text-emerald-600' : 
                  evaluateParam('Máy biến áp', 'oilLeak', formData.oilLeak) === 'warning' ? 'text-amber-600' : 'text-rose-600'
                }`}>
                  {evaluateParam('Máy biến áp', 'oilLeak', formData.oilLeak) === 'healthy' ? '✓ Đạt tiêu chuẩn' : 
                   evaluateParam('Máy biến áp', 'oilLeak', formData.oilLeak) === 'warning' ? '⚠ Cảnh báo' : '⚠ Vượt ngưỡng nguy hiểm'}
                </div>
              )}
            </div>
            <div className="flex flex-col justify-center space-y-1.5 bg-white p-2.5 rounded-lg border border-slate-200 shadow-sm">
              <div className="flex justify-between items-center text-xs">
                <span className="text-slate-500">Tiêu chuẩn:</span>
                <span className="font-semibold text-emerald-600">Bình thường</span>
              </div>
              <div className="flex justify-between items-center text-xs">
                <span className="text-slate-500">Lần trước:</span>
                <span className="font-semibold text-slate-700">Bình thường</span>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Parameter Group 3: Phân tích dầu (DGA & Hóa lý) */}
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Droplets size={16} className="text-blue-500" />
          Phân tích dầu (DGA & Hóa lý)
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <ParamInput 
            title="Khí hòa tan (DGA - TDCG)" unit="ppm" standard="≤ 1000 ppm" prevValue="850 ppm" prevTrend="up"
            value={formData.dga} onChange={(val: any) => setFormData({...formData, dga: val})}
            evalStatus={evaluateParam('Máy biến áp', 'dga', formData.dga)}
            onHistoryClick={() => { setSelectedParamHistory('DGA'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Độ bền điện môi" unit="kV" standard="≥ 50 kV" prevValue="55 kV" prevTrend="down"
            value={formData.dielectricStrength} onChange={(val: any) => setFormData({...formData, dielectricStrength: val})}
            evalStatus={evaluateParam('Máy biến áp', 'dielectricStrength', formData.dielectricStrength)}
            onHistoryClick={() => { setSelectedParamHistory('Độ bền điện môi'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Hàm lượng Furan" unit="mg/kg" standard="≤ 1 mg/kg" prevValue="0.5 mg/kg" prevTrend="stable"
            value={formData.furan} onChange={(val: any) => setFormData({...formData, furan: val})}
            evalStatus={evaluateParam('Máy biến áp', 'furan', formData.furan)}
            onHistoryClick={() => { setSelectedParamHistory('Furan'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Độ ẩm trong dầu" unit="ppm" standard="≤ 15 ppm" prevValue="12 ppm" prevTrend="up"
            value={formData.oilMoisture} onChange={(val: any) => setFormData({...formData, oilMoisture: val})}
            evalStatus={evaluateParam('Máy biến áp', 'oilMoisture', formData.oilMoisture)}
            onHistoryClick={() => { setSelectedParamHistory('Độ ẩm dầu'); setShowHistoryModal(true); }}
          />
        </div>
      </div>

      {/* Chi tiết khí hòa tan (DGA) */}
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Activity size={16} className="text-blue-500" />
          Chi tiết khí hòa tan (DGA)
        </h3>
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
          <ParamInput 
            title="H2" unit="ppm" standard="< 100" prevValue="15" prevTrend="stable"
            value={formData.h2} onChange={(val: any) => setFormData({...formData, h2: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('H2'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="CH4" unit="ppm" standard="< 120" prevValue="20" prevTrend="stable"
            value={formData.ch4} onChange={(val: any) => setFormData({...formData, ch4: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('CH4'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="C2H6" unit="ppm" standard="< 65" prevValue="10" prevTrend="stable"
            value={formData.c2h6} onChange={(val: any) => setFormData({...formData, c2h6: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('C2H6'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="C2H4" unit="ppm" standard="< 50" prevValue="5" prevTrend="stable"
            value={formData.c2h4} onChange={(val: any) => setFormData({...formData, c2h4: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('C2H4'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="C2H2" unit="ppm" standard="< 1" prevValue="0" prevTrend="stable"
            value={formData.c2h2} onChange={(val: any) => setFormData({...formData, c2h2: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('C2H2'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="CO" unit="ppm" standard="< 350" prevValue="150" prevTrend="stable"
            value={formData.co} onChange={(val: any) => setFormData({...formData, co: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('CO'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="CO2" unit="ppm" standard="< 2500" prevValue="1200" prevTrend="stable"
            value={formData.co2} onChange={(val: any) => setFormData({...formData, co2: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('CO2'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="O2" unit="ppm" standard="" prevValue="500" prevTrend="stable"
            value={formData.o2} onChange={(val: any) => setFormData({...formData, o2: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('O2'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="N2" unit="ppm" standard="" prevValue="45000" prevTrend="stable"
            value={formData.n2} onChange={(val: any) => setFormData({...formData, n2: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('N2'); setShowHistoryModal(true); }}
          />
        </div>
      </div>
    </>
  );

  const renderMotorParams = () => (
    <>
      {/* Thông số cơ bản (Tuổi & Hệ số tải) */}
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Info size={16} className="text-slate-500" />
          Thông số cơ bản
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <ParamInput 
            title="Tuổi thiết bị (Age)" unit="năm" standard="< 40 năm" prevValue="10 năm" prevTrend="stable"
            value={formData.age} onChange={(val: any) => setFormData({...formData, age: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('Tuổi thiết bị'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Hệ số tải (Duty Factor)" unit="" standard="1.0" prevValue="1.0" prevTrend="stable"
            value={formData.dutyFactor} onChange={(val: any) => setFormData({...formData, dutyFactor: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('Hệ số tải'); setShowHistoryModal(true); }}
          />
        </div>
      </div>

      {/* Thông số vận hành */}
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Activity size={16} className="text-emerald-500" />
          Thông số vận hành
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <ParamInput 
            title="Độ rung (Vibration)" unit="mm/s" standard="≤ 4.5 mm/s" prevValue="2.1 mm/s" prevTrend="up"
            value={formData.vibration} onChange={(val: any) => setFormData({...formData, vibration: val})}
            evalStatus={evaluateParam('Động cơ', 'vibration', formData.vibration)}
            onHistoryClick={() => { setSelectedParamHistory('Độ rung'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Mất cân bằng điện áp (Voltage Imbalance)" unit="%" standard="≤ 3%" prevValue="1.5%" prevTrend="stable"
            value={formData.voltageImbalance} onChange={(val: any) => setFormData({...formData, voltageImbalance: val})}
            evalStatus={evaluateParam('Động cơ', 'voltageImbalance', formData.voltageImbalance)}
            onHistoryClick={() => { setSelectedParamHistory('Mất cân bằng điện áp'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Nhiệt độ cuộn dây Stator" unit="°C" standard="≤ 130 °C" prevValue="110 °C" prevTrend="up"
            value={formData.statorTemp} onChange={(val: any) => setFormData({...formData, statorTemp: val})}
            evalStatus={evaluateParam('Động cơ', 'statorTemp', formData.statorTemp)}
            onHistoryClick={() => { setSelectedParamHistory('Nhiệt độ Stator'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Nhiệt độ vòng bi (DE/NDE)" unit="°C" standard="≤ 95 °C" prevValue="82 °C" prevTrend="stable"
            value={formData.bearingTemp} onChange={(val: any) => setFormData({...formData, bearingTemp: val})}
            evalStatus={evaluateParam('Động cơ', 'bearingTemp', formData.bearingTemp)}
            onHistoryClick={() => { setSelectedParamHistory('Nhiệt độ vòng bi'); setShowHistoryModal(true); }}
          />
        </div>
      </div>

      {/* Thử nghiệm chẩn đoán (Diagnostic Tests) */}
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Zap size={16} className="text-amber-500" />
          Thử nghiệm chẩn đoán (3 Pha R-Y-B)
        </h3>
        <div className="grid grid-cols-1 gap-4">
          <ThreePhaseParamInput 
            title="Tan-delta" unit="" standard="< 0.07"
            values={formData.tanDelta || { R: '', Y: '', B: '' }}
            onChange={(phase, val) => updateThreePhase('tanDelta', phase, val)}
            evalStatuses={{
              R: evaluateParam('Động cơ', 'tanDelta', formData.tanDelta?.R),
              Y: evaluateParam('Động cơ', 'tanDelta', formData.tanDelta?.Y),
              B: evaluateParam('Động cơ', 'tanDelta', formData.tanDelta?.B)
            }}
          />
          <ThreePhaseParamInput 
            title="Tip-up" unit="" standard="< 0.006"
            values={formData.tipUp || { R: '', Y: '', B: '' }}
            onChange={(phase, val) => updateThreePhase('tipUp', phase, val)}
            evalStatuses={{
              R: evaluateParam('Động cơ', 'tipUp', formData.tipUp?.R),
              Y: evaluateParam('Động cơ', 'tipUp', formData.tipUp?.Y),
              B: evaluateParam('Động cơ', 'tipUp', formData.tipUp?.B)
            }}
          />
          <ThreePhaseParamInput 
            title="Phóng điện cục bộ (PD)" unit="pC" standard="< 15000 pC"
            values={formData.pd || { R: '', Y: '', B: '' }}
            onChange={(phase, val) => updateThreePhase('pd', phase, val)}
            evalStatuses={{
              R: evaluateParam('Động cơ', 'pd', formData.pd?.R),
              Y: evaluateParam('Động cơ', 'pd', formData.pd?.Y),
              B: evaluateParam('Động cơ', 'pd', formData.pd?.B)
            }}
          />
          <ThreePhaseParamInput 
            title="Điện trở cách điện (IR)" unit="GΩ" standard="> 1 GΩ"
            values={formData.ir || { R: '', Y: '', B: '' }}
            onChange={(phase, val) => updateThreePhase('ir', phase, val)}
            evalStatuses={{
              R: evaluateParam('Động cơ', 'ir', formData.ir?.R),
              Y: evaluateParam('Động cơ', 'ir', formData.ir?.Y),
              B: evaluateParam('Động cơ', 'ir', formData.ir?.B)
            }}
          />
          <ThreePhaseParamInput 
            title="Chỉ số phân cực (PI)" unit="" standard="> 1.0"
            values={formData.pi || { R: '', Y: '', B: '' }}
            onChange={(phase, val) => updateThreePhase('pi', phase, val)}
            evalStatuses={{
              R: evaluateParam('Động cơ', 'pi', formData.pi?.R),
              Y: evaluateParam('Động cơ', 'pi', formData.pi?.Y),
              B: evaluateParam('Động cơ', 'pi', formData.pi?.B)
            }}
          />
          <ThreePhaseParamInput 
            title="Dielectric Discharge (DD)" unit="" standard="< 8"
            values={formData.dd || { R: '', Y: '', B: '' }}
            onChange={(phase, val) => updateThreePhase('dd', phase, val)}
            evalStatuses={{
              R: evaluateParam('Động cơ', 'dd', formData.dd?.R),
              Y: evaluateParam('Động cơ', 'dd', formData.dd?.Y),
              B: evaluateParam('Động cơ', 'dd', formData.dd?.B)
            }}
          />
          <ThreePhaseParamInput 
            title="ELCID" unit="mA" standard="< 300 mA"
            values={formData.elcid || { R: '', Y: '', B: '' }}
            onChange={(phase, val) => updateThreePhase('elcid', phase, val)}
            evalStatuses={{
              R: evaluateParam('Động cơ', 'elcid', formData.elcid?.R),
              Y: evaluateParam('Động cơ', 'elcid', formData.elcid?.Y),
              B: evaluateParam('Động cơ', 'elcid', formData.elcid?.B)
            }}
          />
        </div>
      </div>
    </>
  );

  const renderInverterParams = () => (
    <>
      {/* Thông số điện cơ bản */}
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Zap size={16} className="text-blue-500" />
          Thông số điện cơ bản (DC & AC)
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <ParamInput 
            title="Điện áp DC đầu vào (Voc max)" unit="V" standard="< Umax Inverter" prevValue="600V" prevTrend="stable"
            value={formData.dcInputVoltage} onChange={(val: any) => setFormData({...formData, dcInputVoltage: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('Điện áp DC đầu vào'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Dải điện áp MPPT" unit="V" standard="UMPPTmin ≤ V ≤ UMPPTmax" prevValue="450V" prevTrend="stable"
            value={formData.mpptVoltageRange} onChange={(val: any) => setFormData({...formData, mpptVoltageRange: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('Dải điện áp MPPT'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Dòng điện DC đầu vào (Isc)" unit="A" standard="≤ Imax Inverter" prevValue="15A" prevTrend="stable"
            value={formData.dcInputCurrent} onChange={(val: any) => setFormData({...formData, dcInputCurrent: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('Dòng điện DC đầu vào'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Điện áp đầu ra AC" unit="V" standard="230/400V (+6%/-10%)" prevValue="230V" prevTrend="stable"
            value={formData.acOutputVoltage} onChange={(val: any) => setFormData({...formData, acOutputVoltage: val})}
            evalStatus={evaluateParam('Inverter', 'acOutputVoltage', formData.acOutputVoltage)}
            onHistoryClick={() => { setSelectedParamHistory('Điện áp đầu ra AC'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Tần số AC" unit="Hz" standard="50Hz (±0.3 Hz)" prevValue="50Hz" prevTrend="stable"
            value={formData.acFrequency} onChange={(val: any) => setFormData({...formData, acFrequency: val})}
            evalStatus={evaluateParam('Inverter', 'acFrequency', formData.acFrequency)}
            onHistoryClick={() => { setSelectedParamHistory('Tần số AC'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Độ méo hài (THD)" unit="%" standard="< 3%" prevValue="1.2%" prevTrend="stable"
            value={formData.harmonicDistortion} onChange={(val: any) => setFormData({...formData, harmonicDistortion: val})}
            evalStatus={evaluateParam('Inverter', 'harmonicDistortion', formData.harmonicDistortion)}
            onHistoryClick={() => { setSelectedParamHistory('Độ méo hài'); setShowHistoryModal(true); }}
          />
        </div>
      </div>

      {/* Thông số Hiệu suất */}
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Activity size={16} className="text-emerald-500" />
          Thông số Hiệu suất
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <ParamInput 
            title="Hiệu suất tối đa (Max Efficiency)" unit="%" standard="95% - 97.5%" prevValue="97.2%" prevTrend="stable"
            value={formData.maxEfficiency} onChange={(val: any) => setFormData({...formData, maxEfficiency: val})}
            evalStatus={evaluateParam('Inverter', 'maxEfficiency', formData.maxEfficiency)}
            onHistoryClick={() => { setSelectedParamHistory('Hiệu suất tối đa'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Hiệu suất Châu Âu (Euro Efficiency)" unit="%" standard="94% - 96%" prevValue="95.5%" prevTrend="stable"
            value={formData.euroEfficiency} onChange={(val: any) => setFormData({...formData, euroEfficiency: val})}
            evalStatus={evaluateParam('Inverter', 'euroEfficiency', formData.euroEfficiency)}
            onHistoryClick={() => { setSelectedParamHistory('Hiệu suất Châu Âu'); setShowHistoryModal(true); }}
          />
        </div>
      </div>

      {/* Bảo vệ An toàn */}
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <ShieldCheck size={16} className="text-rose-500" />
          Bảo vệ An toàn & Tiêu chuẩn
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <div className="p-3 border border-slate-200 rounded-lg bg-slate-50">
            <label className="block text-xs font-bold text-slate-500 uppercase mb-2">Chống hòa lưới (Anti-islanding)</label>
            <select 
              value={formData.antiIslanding} 
              onChange={(e) => setFormData({...formData, antiIslanding: e.target.value})}
              className={`w-full px-3 py-2 border rounded-lg focus:outline-none focus:ring-2 ${
                evaluateParam('Inverter', 'antiIslanding', formData.antiIslanding) === 'Bình thường' ? 'border-emerald-200 focus:ring-emerald-500' : 'border-rose-200 focus:ring-rose-500'
              }`}
            >
              <option value="">-- Chọn trạng thái --</option>
              <option value="Đạt">Đạt (Ngắt &lt; 0.1s)</option>
              <option value="Không đạt">Không đạt</option>
            </select>
          </div>
          <div className="p-3 border border-slate-200 rounded-lg bg-slate-50">
            <label className="block text-xs font-bold text-slate-500 uppercase mb-2">Dòng rò & Cách ly (RCD Type B)</label>
            <select 
              value={formData.rcdIsolation} 
              onChange={(e) => setFormData({...formData, rcdIsolation: e.target.value})}
              className={`w-full px-3 py-2 border rounded-lg focus:outline-none focus:ring-2 ${
                evaluateParam('Inverter', 'rcdIsolation', formData.rcdIsolation) === 'Bình thường' ? 'border-emerald-200 focus:ring-emerald-500' : 'border-rose-200 focus:ring-rose-500'
              }`}
            >
              <option value="">-- Chọn trạng thái --</option>
              <option value="Đạt">Đạt tiêu chuẩn IEC 60364</option>
              <option value="Không đạt">Không đạt</option>
            </select>
          </div>
          <div className="p-3 border border-slate-200 rounded-lg bg-slate-50">
            <label className="block text-xs font-bold text-slate-500 uppercase mb-2">Cách ly vật lý (Disconnects)</label>
            <select 
              value={formData.physicalDisconnects} 
              onChange={(e) => setFormData({...formData, physicalDisconnects: e.target.value})}
              className={`w-full px-3 py-2 border rounded-lg focus:outline-none focus:ring-2 ${
                evaluateParam('Inverter', 'physicalDisconnects', formData.physicalDisconnects) === 'Bình thường' ? 'border-emerald-200 focus:ring-emerald-500' : 'border-rose-200 focus:ring-rose-500'
              }`}
            >
              <option value="">-- Chọn trạng thái --</option>
              <option value="Đạt">Đạt (Có khóa OFF)</option>
              <option value="Không đạt">Không đạt</option>
            </select>
          </div>
          <ParamInput 
            title="Điện áp chịu đựng (Dielectric)" unit="kV" standard="DC: 2.5-8kV, AC: 4kV" prevValue="4kV" prevTrend="stable"
            value={formData.dielectricVoltage} onChange={(val: any) => setFormData({...formData, dielectricVoltage: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('Điện áp chịu đựng'); setShowHistoryModal(true); }}
          />
        </div>
      </div>

      {/* Môi trường & Vật lý */}
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Thermometer size={16} className="text-amber-500" />
          Yêu cầu Vật lý & Môi trường
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <ParamInput 
            title="Nhiệt độ vận hành" unit="°C" standard="≤ 40 °C" prevValue="35 °C" prevTrend="stable"
            value={formData.operatingTemp} onChange={(val: any) => setFormData({...formData, operatingTemp: val})}
            evalStatus={evaluateParam('Inverter', 'operatingTemp', formData.operatingTemp)}
            onHistoryClick={() => { setSelectedParamHistory('Nhiệt độ vận hành'); setShowHistoryModal(true); }}
          />
          <div className="p-3 border border-slate-200 rounded-lg bg-slate-50">
            <label className="block text-xs font-bold text-slate-500 uppercase mb-2">Bộ lọc khí (Air Filters)</label>
            <select 
              value={formData.airFilters} 
              onChange={(e) => setFormData({...formData, airFilters: e.target.value})}
              className={`w-full px-3 py-2 border rounded-lg focus:outline-none focus:ring-2 ${
                evaluateParam('Inverter', 'airFilters', formData.airFilters) === 'Bình thường' ? 'border-emerald-200 focus:ring-emerald-500' : 'border-amber-200 focus:ring-amber-500'
              }`}
            >
              <option value="">-- Chọn trạng thái --</option>
              <option value="Sạch">Sạch / Thông thoáng</option>
              <option value="Bẩn">Bẩn / Cần vệ sinh</option>
            </select>
          </div>
        </div>
      </div>
    </>
  );

  const renderSwitchgearParams = () => (
    <>
      {/* Thông số cơ bản (Tuổi & Hệ số tải) */}
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Info size={16} className="text-slate-500" />
          Thông số cơ bản
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <ParamInput 
            title="Tuổi thiết bị (Age)" unit="năm" standard="< 40 năm" prevValue="10 năm" prevTrend="stable"
            value={formData.age} onChange={(val: any) => setFormData({...formData, age: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('Tuổi thiết bị'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Hệ số tải (Duty Factor)" unit="" standard="1.0" prevValue="1.0" prevTrend="stable"
            value={formData.dutyFactor} onChange={(val: any) => setFormData({...formData, dutyFactor: val})}
            evalStatus={null}
            onHistoryClick={() => { setSelectedParamHistory('Hệ số tải'); setShowHistoryModal(true); }}
          />
        </div>
      </div>
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Zap size={16} className="text-amber-500" />
          Kiểm tra điện & Nhiệt
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <ParamInput 
            title="Nhiệt độ tiếp xúc (Thermography)" unit="°C" standard="≤ 75 °C" prevValue="60 °C" prevTrend="stable"
            value={formData.thermography} onChange={(val: any) => setFormData({...formData, thermography: val})}
            evalStatus={evaluateParam('Tủ điện', 'thermography', formData.thermography)}
            onHistoryClick={() => { setSelectedParamHistory('Nhiệt độ tiếp xúc'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Điện trở tiếp xúc" unit="µΩ" standard="≤ 50 µΩ" prevValue="35 µΩ" prevTrend="up"
            value={formData.contactRes} onChange={(val: any) => setFormData({...formData, contactRes: val})}
            evalStatus={evaluateParam('Tủ điện', 'contactRes', formData.contactRes)}
            onHistoryClick={() => { setSelectedParamHistory('Điện trở tiếp xúc'); setShowHistoryModal(true); }}
          />
        </div>
      </div>
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Wind size={16} className="text-indigo-500" />
          Phóng điện cục bộ (PD)
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <ParamInput 
            title="TEV (Transient Earth Voltage)" unit="dBmV" standard="≤ 20 dBmV" prevValue="12 dBmV" prevTrend="stable"
            value={formData.tev} onChange={(val: any) => setFormData({...formData, tev: val})}
            evalStatus={evaluateParam('Tủ điện trung thế', 'tev', formData.tev)}
            onHistoryClick={() => { setSelectedParamHistory('TEV'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Số xung TEV/chu kỳ" unit="xung/chu kỳ" standard="≤ 1" prevValue="0" prevTrend="stable"
            value={formData.tevPulses} onChange={(val: any) => setFormData({...formData, tevPulses: val})}
            evalStatus={evaluateParam('Tủ điện trung thế', 'tevPulses', formData.tevPulses)}
            onHistoryClick={() => { setSelectedParamHistory('Số xung TEV'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Siêu âm (Ultrasonic)" unit="dBµV" standard="≤ 10 dBµV" prevValue="4 dBµV" prevTrend="down"
            value={formData.ultrasonic} onChange={(val: any) => setFormData({...formData, ultrasonic: val})}
            evalStatus={evaluateParam('Tủ điện trung thế', 'ultrasonic', formData.ultrasonic)}
            onHistoryClick={() => { setSelectedParamHistory('Ultrasonic'); setShowHistoryModal(true); }}
          />
        </div>
      </div>
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Droplets size={16} className="text-blue-500" />
          Môi trường & Khí SF6
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <ParamInput 
            title="Độ ẩm môi trường" unit="%" standard="≤ 60 %" prevValue="55 %" prevTrend="up"
            value={formData.humidity} onChange={(val: any) => setFormData({...formData, humidity: val})}
            evalStatus={evaluateParam('Tủ điện', 'humidity', formData.humidity)}
            onHistoryClick={() => { setSelectedParamHistory('Độ ẩm môi trường'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Áp suất khí SF6" unit="bar" standard="≥ 5.5 bar" prevValue="5.8 bar" prevTrend="down"
            value={formData.sf6Pressure} onChange={(val: any) => setFormData({...formData, sf6Pressure: val})}
            evalStatus={evaluateParam('Tủ điện', 'sf6Pressure', formData.sf6Pressure)}
            onHistoryClick={() => { setSelectedParamHistory('Áp suất SF6'); setShowHistoryModal(true); }}
          />
        </div>
      </div>
    </>
  );

  if (isAuthLoading) {
    return (
      <div className="flex h-screen items-center justify-center bg-slate-50">
        <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600"></div>
      </div>
    );
  }

  if (!user) {
    return (
      <div className="flex h-screen items-center justify-center bg-slate-50">
        <div className="bg-white p-8 rounded-2xl shadow-lg max-w-md w-full text-center">
          <div className="flex justify-center mb-6">
            <img src="https://static.wixstatic.com/media/dea24c_6207cba10c9e4e15939adcc4dbd42524~mv2.jpg" alt="TEV Logo" className="w-20 h-20 rounded-full object-contain border border-slate-200 p-1" />
          </div>
          <h1 className="text-2xl font-bold text-slate-900 mb-2">Welcome to TEV COE</h1>
          <p className="text-slate-500 mb-8">Vui lòng đăng nhập để tiếp tục sử dụng hệ thống.</p>
          <button 
            onClick={handleLogin}
            className="w-full flex items-center justify-center gap-3 bg-blue-600 hover:bg-blue-700 text-white py-3 px-4 rounded-xl font-medium transition-colors"
          >
            <svg className="w-5 h-5" viewBox="0 0 24 24">
              <path fill="currentColor" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" />
              <path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" />
              <path fill="#FBBC05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z" />
              <path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z" />
            </svg>
            Đăng nhập bằng Google
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="flex h-screen bg-slate-50 font-sans text-slate-900 overflow-hidden relative">
      
      {/* MOBILE SIDEBAR BACKDROP */}
      {isSidebarOpen && (
        <div 
          className="fixed inset-0 bg-slate-900/50 z-40 lg:hidden backdrop-blur-sm transition-opacity"
          onClick={() => setIsSidebarOpen(false)}
        />
      )}

      {/* SIDEBAR */}
      <aside className={`w-64 bg-slate-900 text-slate-300 flex flex-col flex-shrink-0 fixed inset-y-0 left-0 z-50 transform ${isSidebarOpen ? 'translate-x-0' : '-translate-x-full'} lg:translate-x-0 lg:relative transition-transform duration-300 ease-in-out`}>
        <div className="h-16 flex items-center justify-between px-6 border-b border-slate-800">
          <div className="flex items-center gap-2 text-white font-bold text-xl tracking-tight">
            <img src="https://static.wixstatic.com/media/dea24c_6207cba10c9e4e15939adcc4dbd42524~mv2.jpg" alt="TEV Logo" className="w-8 h-8 rounded-full bg-white object-contain p-0.5" />
            <span>TEV COE</span>
          </div>
          <button 
            className="lg:hidden text-slate-400 hover:text-white"
            onClick={() => setIsSidebarOpen(false)}
          >
            <X size={20} />
          </button>
        </div>
        
        <div className="px-6 py-4 border-b border-slate-800">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-full bg-slate-700 flex items-center justify-center text-white font-bold">
              {user?.displayName ? user.displayName.charAt(0).toUpperCase() : <User size={20} />}
            </div>
            <div className="flex-1 min-w-0">
              <p className="text-sm font-medium text-white truncate">{user?.displayName || user?.email}</p>
              <p className="text-xs text-slate-400 truncate capitalize">{userRole === 'admin' ? 'Quản trị viên' : userCustomerName || 'Khách hàng'}</p>
            </div>
          </div>
        </div>

        <nav className="flex-1 px-4 py-6 space-y-1 overflow-y-auto">
          <div className="px-3 py-2 text-xs font-semibold text-slate-500 uppercase tracking-wider">
            {userRole === 'admin' ? 'Quản lý (Admin)' : 'Khách hàng (Portal)'}
          </div>
          <button 
            onClick={() => handleTabChange('dashboard')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${activeTab === 'dashboard' ? 'bg-blue-600 text-white' : 'hover:bg-slate-800 hover:text-white'}`}
          >
            <LayoutDashboard size={18} />
            Dashboard Tổng quan
          </button>
          <button 
            onClick={() => handleTabChange('equipment')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${activeTab === 'equipment' ? 'bg-blue-600 text-white' : 'hover:bg-slate-800 hover:text-white'}`}
          >
            <Server size={18} />
            Quản lý Thiết bị
          </button>
          <button 
            onClick={() => handleTabChange('reports')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${activeTab === 'reports' ? 'bg-blue-600 text-white' : 'hover:bg-slate-800 hover:text-white'}`}
          >
            <FileText size={18} />
            Kho Báo cáo
          </button>

          <div className="px-3 py-2 mt-4 text-xs font-semibold text-slate-500 uppercase tracking-wider">
            Nội bộ TEV (Admin/FSE)
          </div>
          <button 
            onClick={() => handleTabChange('field-entry')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${activeTab === 'field-entry' ? 'bg-blue-600 text-white' : 'hover:bg-slate-800 hover:text-white'}`}
          >
            <ClipboardList size={18} />
            Nhập liệu hiện trường
          </button>
          <button 
            onClick={() => handleTabChange('deep-analysis')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${activeTab === 'deep-analysis' ? 'bg-blue-600 text-white' : 'hover:bg-slate-800 hover:text-white'}`}
          >
            <Activity size={18} />
            Phân tích chuyên sâu
          </button>
          <button 
            onClick={() => handleTabChange('cmms')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${activeTab === 'cmms' ? 'bg-blue-600 text-white' : 'hover:bg-slate-800 hover:text-white'}`}
          >
            <ClipboardList size={18} />
            Quản lý CMMS
          </button>
          <button 
            onClick={() => handleTabChange('pm-schedule')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${activeTab === 'pm-schedule' ? 'bg-blue-600 text-white' : 'hover:bg-slate-800 hover:text-white'}`}
          >
            <Calendar size={18} />
            Lịch bảo trì PM
          </button>
          <button 
            onClick={() => handleTabChange('customers')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${activeTab === 'customers' ? 'bg-blue-600 text-white' : 'hover:bg-slate-800 hover:text-white'}`}
          >
            <User size={18} />
            Quản lý Khách hàng
          </button>
          <button 
            onClick={() => handleTabChange('inventory')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${activeTab === 'inventory' ? 'bg-blue-600 text-white' : 'hover:bg-slate-800 hover:text-white'}`}
          >
            <Package size={18} />
            Quản lý Kho
          </button>
        </nav>

        <div className="p-4 border-t border-slate-800 space-y-1">
          <button 
            onClick={() => handleTabChange('settings')}
            className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${activeTab === 'settings' ? 'bg-blue-600 text-white' : 'hover:bg-slate-800 hover:text-white'}`}
          >
            <Settings size={18} />
            Cài đặt
          </button>
          <button 
            onClick={logOut}
            className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium text-red-400 hover:bg-red-500/10 hover:text-red-300 transition-colors"
          >
            <LogOut size={18} />
            Đăng xuất
          </button>
        </div>
      </aside>

      {/* MAIN CONTENT */}
      <main className="flex-1 flex flex-col min-w-0 overflow-hidden w-full">
        
        {/* HEADER */}
        <header className="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-4 md:px-8 flex-shrink-0 z-10">
          <div className="flex items-center gap-3 md:gap-4">
            <button 
              className="lg:hidden p-2 -ml-2 text-slate-600 hover:bg-slate-100 rounded-lg transition-colors"
              onClick={() => setIsSidebarOpen(true)}
            >
              <Menu size={24} />
            </button>
            <h1 className="text-lg md:text-xl font-semibold text-slate-800 truncate max-w-[180px] sm:max-w-none">
              {activeTab === 'dashboard' ? 'Dashboard Tổng quan' : 
               activeTab === 'equipment' ? 'Quản lý Thiết bị' : 
               activeTab === 'reports' ? 'Kho Báo cáo' : 
               activeTab === 'settings' ? 'Cài đặt hệ thống' :
               'Nhập liệu hiện trường'}
            </h1>
            {activeTab !== 'field-entry' && (
              <div className="hidden md:flex items-center gap-2 px-3 py-1.5 bg-slate-100 rounded-md text-sm text-slate-600 border border-slate-200 cursor-pointer hover:bg-slate-200 transition-colors">
                <Factory size={16} className="text-slate-400" />
                <span>Tất cả nhà máy</span>
                <ChevronDown size={14} />
              </div>
            )}
          </div>
          
          <div className="flex items-center gap-3">
            {!isGoogleConnected ? (
              <button 
                onClick={handleConnectGoogle}
                className="flex items-center gap-2 px-3 py-1.5 bg-blue-50 text-blue-600 rounded-md text-sm font-medium hover:bg-blue-100 transition-colors border border-blue-200"
              >
                <UploadCloud size={16} />
                Kết nối Google Drive
              </button>
            ) : (
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setIsAutoSyncEnabled(!isAutoSyncEnabled)}
                  title={isAutoSyncEnabled ? "Đang bật tự động đồng bộ 2 chiều mỗi 60s" : "Đã tạm dừng tự động đồng bộ"}
                  className={`hidden sm:flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold border transition-all ${
                    isAutoSyncEnabled
                      ? 'bg-emerald-50 text-emerald-700 border-emerald-300'
                      : 'bg-slate-100 text-slate-500 border-slate-300'
                  }`}
                >
                  <span className={`w-2 h-2 rounded-full ${isAutoSyncEnabled ? 'bg-emerald-500 animate-pulse' : 'bg-slate-400'}`}></span>
                  <span>{isAutoSyncEnabled ? 'Auto 2-Way Sync' : 'Sync Tắt'}</span>
                </button>

                <button
                  onClick={() => handlePerformTwoWaySync(false)}
                  disabled={isSyncing || autoSyncStatus === 'syncing'}
                  title="Đồng bộ toàn bộ dữ liệu 2 chiều ngay lập tức"
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-bold shadow-sm transition-all disabled:opacity-50"
                >
                  <RefreshCw size={13} className={isSyncing || autoSyncStatus === 'syncing' ? 'animate-spin' : ''} />
                  <span>{isSyncing || autoSyncStatus === 'syncing' ? 'Đang Sync...' : 'Sync 2 Chiều'}</span>
                </button>

                <button
                  onClick={handleOpenReconciliation}
                  title="Đối soát phiên bản dữ liệu (Reconciliation) giữa Firestore và Google Sheets"
                  className="hidden md:flex items-center gap-1 px-2.5 py-1.5 bg-amber-50 hover:bg-amber-100 text-amber-800 rounded-lg text-xs font-semibold border border-amber-200 transition-colors"
                >
                  <ShieldAlert size={14} className="text-amber-600" />
                  <span>Đối soát</span>
                </button>

                {googleSheetUrl && (
                  <a
                    href={googleSheetUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    title="Mở file Google Sheet 'TEV Service Flatform'"
                    className="hidden lg:flex items-center gap-1 px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-medium border border-slate-200 transition-colors"
                  >
                    <ExternalLink size={13} className="text-slate-500" />
                    <span>File Sheets</span>
                  </a>
                )}
              </div>
            )}
            <div className="relative hidden md:block">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
              <input 
                type="text" 
                placeholder="Tìm kiếm..." 
                className="pl-9 pr-4 py-1.5 bg-slate-100 border-transparent focus:bg-white focus:border-blue-500 focus:ring-2 focus:ring-blue-200 rounded-lg text-sm w-64 transition-all outline-none"
              />
            </div>
            <button className="relative p-2 text-slate-400 hover:text-slate-600 transition-colors">
              <Bell size={20} />
              <span className="absolute top-1.5 right-1.5 w-2 h-2 bg-rose-500 rounded-full border-2 border-white"></span>
            </button>
            <div className="w-8 h-8 rounded-full bg-blue-100 text-blue-700 flex items-center justify-center font-bold text-sm border border-blue-200 cursor-pointer">
              {activeTab === 'field-entry' ? 'FSE' : 'AD'}
            </div>
          </div>
        </header>

        {/* Real-time 2-Way Sync Notice Banner */}
        {twoWaySyncNotice && (
          <div className="px-4 md:px-8 pt-3 bg-slate-100">
            <div className={`p-3 rounded-xl border flex items-center justify-between shadow-xs transition-all ${
              twoWaySyncNotice.type === 'success'
                ? 'bg-emerald-50 border-emerald-300 text-emerald-900'
                : twoWaySyncNotice.type === 'error'
                ? 'bg-rose-50 border-rose-300 text-rose-900'
                : 'bg-blue-50 border-blue-300 text-blue-900'
            }`}>
              <div className="flex items-center gap-2 text-xs font-semibold">
                {twoWaySyncNotice.type === 'success' ? (
                  <CheckCircle size={16} className="text-emerald-600 shrink-0" />
                ) : (
                  <AlertTriangle size={16} className="text-rose-600 shrink-0" />
                )}
                <span>{twoWaySyncNotice.message}</span>
              </div>
              <button 
                onClick={() => setTwoWaySyncNotice(null)}
                className="text-slate-400 hover:text-slate-700 p-1"
              >
                <X size={14} />
              </button>
            </div>
          </div>
        )}

        {/* SCROLLABLE CONTENT */}
        <div className="flex-1 overflow-y-auto p-4 md:p-8 bg-slate-100">
          
          {/* --- FIELD ENTRY VIEW --- */}
          {activeTab === 'field-entry' && (
            <div className="max-w-6xl mx-auto space-y-6 pb-12">
              {/* Header Banner for NETA ATS Testing & Quick Link to CMMS */}
              <div className="bg-gradient-to-r from-slate-900 via-blue-950 to-slate-900 rounded-2xl p-6 text-white shadow-md flex flex-col md:flex-row md:items-center justify-between gap-4 border border-blue-900/40">
                <div>
                  <div className="flex items-center gap-2 mb-1.5">
                    <ClipboardCheck className="text-blue-400" size={24} />
                    <h2 className="text-xl font-bold">23 Biên Bản Kiểm Định Chuyên Sâu NETA ATS (Mỹ)</h2>
                    <span className="px-2 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-wider bg-blue-500/30 text-blue-300 border border-blue-400/30">
                      Chuẩn Quốc Tế
                    </span>
                  </div>
                  <p className="text-sm text-slate-300 max-w-2xl">
                    Hệ thống thí nghiệm nghiệm thu và bảo trì điện lực theo tiêu chuẩn ANSI/NETA ATS. Tự động tính toán dung lượng, điện trở cách điện, tỷ số biến dòng & xuất báo cáo PDF chuẩn.
                  </p>
                </div>
                <button
                  type="button"
                  onClick={() => handleTabChange('cmms')}
                  className="flex items-center gap-2 px-4 py-2.5 bg-blue-600 hover:bg-blue-500 text-white rounded-xl text-sm font-bold shadow-md shadow-blue-600/30 transition-all flex-shrink-0"
                  title="Chuyển đến phần Quản lý CMMS để nhập phiếu công việc, quản lý thiết bị & đồng bộ 2 chiều Google Sheet"
                >
                  <FileSpreadsheet size={17} />
                  <span>Quản lý CMMS & Đồng Bộ Sheets →</span>
                </button>
              </div>

              {/* Equipment Checklist Selection Hub */}
              <EquipmentChecklistSelector
                currentMode={fieldEntryMode}
                onSelectMode={(mode) => setFieldEntryMode(mode)}
                reportsCount={allReports.length}
                onNavigateToReports={() => setActiveTab('reports')}
              />

              {fieldEntryMode === 'neta-dc-motor' ? (
                <NetaDcMotorChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-grounding' ? (
                <NetaGroundingChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-lv-breaker' ? (
                <NetaLvBreakerChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-cable-lv' ? (
                <NetaCableLvChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-switchgear' ? (
                <NetaSwitchgearChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-generator' ? (
                <NetaEngineGeneratorChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-ats' ? (
                <NetaAtsChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-battery-vrla' ? (
                <NetaBatteryVrlaChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-battery-flooded' ? (
                <NetaBatteryFloodedChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-sync-machinery' ? (
                <NetaSyncMachineryChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-ups' ? (
                <NetaUpsChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-relay' ? (
                <NetaRelayChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-liquid-transformer' ? (
                <NetaLiquidTransformerChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-motor' ? (
                <NetaMotorChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-pd' ? (
                <NetaPartialDischargeChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-thermography' ? (
                <NetaThermographyChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-sf6-switch' ? (
                <NetaSf6SwitchChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-cable' ? (
                <NetaCableChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-evse' ? (
                <NetaEvseChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-pv' ? (
                <NetaPvChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-bess' ? (
                <NetaBessChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-dry-large' ? (
                <NetaLargeDryTypeChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : fieldEntryMode === 'neta-dry-type' ? (
                <NetaDryTypeChecklist
                  onSaveReport={handleSaveNetaReport}
                  onNavigateToReports={() => setActiveTab('reports')}
                />
              ) : (
                <div className="max-w-3xl mx-auto space-y-6">
              
              {/* Step 1: Identify Equipment */}
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="bg-blue-600 px-5 py-4 text-white flex items-center justify-between">
                  <h2 className="font-semibold text-lg flex items-center gap-2">
                    <Search size={20} />
                    1. Xác định thiết bị
                  </h2>
                  <button 
                    onClick={() => setShowQRScanner(true)}
                    className="bg-blue-500 hover:bg-blue-400 text-white px-3 py-1.5 rounded-lg text-sm font-medium flex items-center gap-2 transition-colors border border-blue-400"
                  >
                    <QrCode size={16} />
                    Quét QR Code
                  </button>
                </div>
                
                {/* QR Scanner Modal */}
                {showQRScanner && (
                  <div className="fixed inset-0 bg-slate-900/50 z-50 flex items-center justify-center p-4">
                    <div className="bg-white rounded-xl shadow-xl w-full max-w-md overflow-hidden">
                      <div className="flex items-center justify-between p-4 border-b border-slate-200">
                        <h3 className="font-semibold text-slate-800">Quét mã QR</h3>
                        <button onClick={() => setShowQRScanner(false)} className="text-slate-400 hover:text-slate-600">
                          <X size={20} />
                        </button>
                      </div>
                      <div className="p-4">
                        <div id="qr-reader" className="w-full"></div>
                        <p className="text-sm text-slate-500 text-center mt-4">Hướng camera vào mã QR trên thiết bị để quét.</p>
                      </div>
                    </div>
                  </div>
                )}

                <div className="p-5 space-y-4">
                  <div className="relative">
                    <label className="block text-sm font-medium text-slate-700 mb-1">Mã thiết bị / Tên thiết bị</label>
                    <div className="relative">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
                      <input 
                        type="text" 
                        value={equipmentCode}
                        onChange={(e) => {
                          setEquipmentCode(e.target.value);
                          setShowEqSuggestions(true);
                        }}
                        onFocus={() => setShowEqSuggestions(true)}
                        onBlur={() => setTimeout(() => setShowEqSuggestions(false), 200)}
                        placeholder="Nhập mã hoặc tên thiết bị..."
                        className="w-full pl-10 pr-4 py-3 bg-slate-50 border border-slate-300 focus:bg-white focus:border-blue-500 focus:ring-2 focus:ring-blue-200 rounded-lg text-base transition-all outline-none font-medium text-slate-900"
                      />
                    </div>
                    {showEqSuggestions && (
                      <div className="absolute z-10 w-full mt-1 bg-white border border-slate-200 rounded-lg shadow-lg max-h-60 overflow-y-auto">
                        {allEquipment
                          .filter(eq => eq.id.toLowerCase().includes(equipmentCode.toLowerCase()) || eq.name.toLowerCase().includes(equipmentCode.toLowerCase()))
                          .map(eq => (
                            <div 
                              key={eq.id} 
                              className="px-4 py-2 hover:bg-slate-50 cursor-pointer border-b border-slate-100 last:border-0"
                              onClick={() => {
                                setEquipmentCode(eq.id);
                                setEquipmentName(eq.name);
                                setCustomerName(eq.customer);
                                setSiteName(eq.factory);
                                setLocationName(eq.location);
                                setSelectedEqType(eq.type === 'Tủ điện' ? 'Tủ điện trung thế' : eq.type);
                                setIsNewEquipment(false);
                                setShowEqSuggestions(false);
                                populateFormData(eq);
                              }}
                            >
                              <div className="font-medium text-slate-900">{eq.id} - {eq.name}</div>
                              <div className="text-xs text-slate-500">{eq.customer} | {eq.factory} | {eq.location}</div>
                            </div>
                          ))}
                        <div 
                          className="px-4 py-2 hover:bg-blue-50 cursor-pointer text-blue-600 font-medium flex items-center gap-2"
                          onClick={() => {
                            setShowEqSuggestions(false);
                            setIsNewEquipment(true);
                            setEquipmentName('');
                            setCustomerName('');
                            setSiteName('');
                            setLocationName('');
                          }}
                        >
                          <Plus size={16} />
                          Tạo mới thiết bị: "{equipmentCode}"
                        </div>
                      </div>
                    )}
                  </div>
                  
                  <div className="mt-4">
                    <label className="block text-sm font-medium text-slate-700 mb-1">Số Work Permit (WP)</label>
                    <input 
                      type="text" 
                      value={wpNumber}
                      onChange={(e) => setWpNumber(e.target.value)}
                      placeholder="VD: WP-2026-001"
                      className="w-full px-4 py-2.5 bg-white border border-slate-300 focus:border-blue-500 focus:ring-2 focus:ring-blue-200 rounded-lg text-base outline-none font-medium text-slate-900"
                    />
                  </div>
                  
                  <div className="mt-4">
                    <label className="block text-sm font-medium text-slate-700 mb-1">Loại thiết bị</label>
                    <select 
                      value={selectedEqType}
                      onChange={(e) => setSelectedEqType(e.target.value)}
                      className="w-full px-4 py-2.5 bg-white border border-slate-300 focus:border-blue-500 focus:ring-2 focus:ring-blue-200 rounded-lg text-base outline-none font-medium text-slate-900"
                    >
                      <option value="Máy biến áp">Máy biến áp</option>
                      <option value="Động cơ">Động cơ điện</option>
                      <option value="Tủ điện trung thế">Tủ điện trung thế (Switchgear)</option>
                      <option value="Inverter">Inverter Solar/Wind</option>
                    </select>
                  </div>
                  
                  {/* Selected Equipment Info Card / New Equipment Form */}
                  {isNewEquipment ? (
                    <div className="bg-blue-50 p-4 rounded-lg border border-blue-200 space-y-4">
                      <h3 className="font-bold text-blue-900 text-sm flex items-center gap-2">
                        <Plus size={16} />
                        Thông tin thiết bị mới
                      </h3>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div>
                          <label className="block text-xs font-medium text-slate-700 mb-1">Tên thiết bị</label>
                          <input 
                            type="text" 
                            value={equipmentName}
                            onChange={(e) => setEquipmentName(e.target.value)}
                            placeholder="VD: Máy biến áp T2"
                            className="w-full px-3 py-2 bg-white border border-slate-300 focus:border-blue-500 focus:ring-1 focus:ring-blue-200 rounded-md text-sm outline-none"
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-medium text-slate-700 mb-1">Khách hàng</label>
                          <select 
                            value={customerName}
                            onChange={(e) => setCustomerName(e.target.value)}
                            className="w-full px-3 py-2 bg-white border border-slate-300 focus:border-blue-500 focus:ring-1 focus:ring-blue-200 rounded-md text-sm outline-none"
                          >
                            <option value="">-- Chọn khách hàng --</option>
                            {customers.map(c => (
                              <option key={c.id} value={c.name}>{c.name}</option>
                            ))}
                          </select>
                        </div>
                        <div>
                          <label className="block text-xs font-medium text-slate-700 mb-1">Nhà máy / Site</label>
                          <input 
                            type="text" 
                            value={siteName}
                            onChange={(e) => setSiteName(e.target.value)}
                            placeholder="VD: Nhà máy Bắc Ninh"
                            className="w-full px-3 py-2 bg-white border border-slate-300 focus:border-blue-500 focus:ring-1 focus:ring-blue-200 rounded-md text-sm outline-none"
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-medium text-slate-700 mb-1">Vị trí / Khu vực</label>
                          <input 
                            type="text" 
                            value={locationName}
                            onChange={(e) => setLocationName(e.target.value)}
                            placeholder="VD: Trạm biến áp 110kV"
                            className="w-full px-3 py-2 bg-white border border-slate-300 focus:border-blue-500 focus:ring-1 focus:ring-blue-200 rounded-md text-sm outline-none"
                          />
                        </div>
                      </div>
                    </div>
                  ) : (
                    <div className="bg-slate-50 p-4 rounded-lg border border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                      <div>
                        <h3 className="font-bold text-slate-900 text-lg">{equipmentName} ({equipmentCode})</h3>
                        <div className="flex flex-col sm:flex-row sm:items-center gap-2 sm:gap-4 mt-1 text-sm text-slate-600">
                          <span className="flex items-center gap-1 font-medium text-blue-700">{customerName}</span>
                          <span className="hidden sm:inline text-slate-300">|</span>
                          <span className="flex items-center gap-1"><Factory size={14} /> {siteName}</span>
                          <span className="flex items-center gap-1"><MapPin size={14} /> {locationName}</span>
                        </div>
                      </div>
                      <div className="flex flex-col sm:items-end gap-3">
                        <div className="text-right">
                          <div className="text-xs text-slate-500 mb-1">Lần kiểm tra cuối</div>
                          <div className="font-medium text-slate-900">
                            {allEquipment.find(e => e.id === equipmentCode)?.lastCheck || 'Chưa có dữ liệu'}
                          </div>
                        </div>
                        <button 
                          onClick={() => setShowEquipmentProfile(true)}
                          className="text-sm bg-white border border-slate-300 text-blue-600 px-3 py-1.5 rounded-md font-medium hover:bg-blue-50 transition-colors flex items-center gap-1.5 shadow-sm"
                        >
                          <History size={16} />
                          Hồ sơ & Báo cáo cũ
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              </div>

              {/* Step 2: Inspection History (If exists) */}
              {!isNewEquipment && equipmentCode && (
                <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                  <div className="bg-slate-100 px-5 py-3 border-b border-slate-200 flex items-center justify-between">
                    <h2 className="font-semibold text-slate-800 flex items-center gap-2">
                      <History size={18} className="text-slate-500" />
                      Lịch sử kiểm tra gần đây
                    </h2>
                  </div>
                  <div className="p-0 overflow-x-auto">
                    <table className="w-full text-sm text-left min-w-[600px]">
                      <thead className="bg-slate-50 text-slate-500 border-b border-slate-200">
                        <tr>
                          <th className="px-4 py-2 font-medium">Ngày</th>
                          <th className="px-4 py-2 font-medium">Người kiểm tra</th>
                          <th className="px-4 py-2 font-medium">Trạng thái</th>
                          <th className="px-4 py-2 font-medium">Ghi chú</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {allReports.filter(r => r.equipmentId === equipmentCode).map((report, idx) => (
                          <tr key={idx} className="hover:bg-slate-50">
                            <td className="px-4 py-3 text-slate-700">{report.date}</td>
                            <td className="px-4 py-3 text-slate-700">{report.inspector}</td>
                            <td className="px-4 py-3">
                              <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium ${
                                report.status === 'healthy' ? 'bg-emerald-100 text-emerald-700' :
                                report.status === 'warning' ? 'bg-amber-100 text-amber-700' :
                                'bg-rose-100 text-rose-700'
                              }`}>
                                {report.status === 'healthy' ? 'Bình thường' : report.status === 'warning' ? 'Cảnh báo' : 'Nguy hiểm'}
                              </span>
                            </td>
                            <td className="px-4 py-3 text-slate-600">{report.notes}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* Step 3: Input Parameters */}
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="bg-slate-800 px-5 py-4 text-white flex items-center justify-between">
                  <h2 className="font-semibold text-lg flex items-center gap-2">
                    <ClipboardList size={20} />
                    {isNewEquipment ? '2' : '3'}. Nhập thông số thí nghiệm
                  </h2>
                </div>
                <div className="p-5 space-y-6">
                  {selectedEqType === 'Máy biến áp' && renderTransformerParams()}
                  {selectedEqType === 'Động cơ' && renderMotorParams()}
                  {selectedEqType === 'Tủ điện' && renderSwitchgearParams()}
                  {selectedEqType === 'Inverter' && renderInverterParams()}

                  {/* Real-time Health Index Display */}
                  <div className="mt-8 p-5 bg-slate-50 border border-slate-200 rounded-xl flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                    <div>
                      <h4 className="text-base font-bold text-slate-800 flex items-center gap-2">
                        <Activity className="text-blue-500" size={18} />
                        Chỉ số sức khỏe (Dự kiến)
                      </h4>
                      <p className="text-sm text-slate-500 mt-1">Tự động tính toán dựa trên các thông số nhập liệu hiện tại</p>
                    </div>
                    {healthResult ? (
                      <div className="flex items-center gap-4 bg-white px-4 py-3 rounded-lg border border-slate-200 shadow-sm">
                        <div className="text-right">
                          <div className="text-3xl font-bold text-slate-900">{healthResult.index}%</div>
                        </div>
                        <div className="h-10 w-px bg-slate-200"></div>
                        <StatusBadge status={healthResult.status} />
                      </div>
                    ) : (
                      <div className="bg-white px-4 py-3 rounded-lg border border-slate-200 border-dashed text-sm text-slate-400 italic flex items-center gap-2">
                        <AlertCircle size={16} />
                        Đang chờ nhập liệu...
                      </div>
                    )}
                  </div>
                </div>
              </div>

              {/* Step 3: Uploads & Notes */}
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="p-5 space-y-6">
                  
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-2">Đính kèm hình ảnh / File báo cáo PDF</label>
                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                      <label className="h-24 border-2 border-dashed border-slate-300 rounded-xl flex flex-col items-center justify-center text-slate-500 hover:bg-slate-50 hover:border-blue-400 hover:text-blue-500 transition-colors cursor-pointer">
                        <Camera size={24} className="mb-1" />
                        <span className="text-xs font-medium">Chụp ảnh</span>
                        <input type="file" accept="image/*" capture="environment" className="hidden" onChange={(e) => {
                          if (e.target.files) setAttachedFiles([...attachedFiles, ...Array.from(e.target.files)]);
                        }} />
                      </label>
                      <label className="h-24 border-2 border-dashed border-slate-300 rounded-xl flex flex-col items-center justify-center text-slate-500 hover:bg-slate-50 hover:border-blue-400 hover:text-blue-500 transition-colors cursor-pointer">
                        <UploadCloud size={24} className="mb-1" />
                        <span className="text-xs font-medium">Tải PDF lên</span>
                        <input type="file" accept=".pdf" multiple className="hidden" onChange={(e) => {
                          if (e.target.files) setAttachedFiles([...attachedFiles, ...Array.from(e.target.files)]);
                        }} />
                      </label>
                      {attachedFiles.map((file, index) => (
                        <div key={index} className="h-24 border border-slate-200 rounded-xl flex flex-col items-center justify-center bg-slate-50 relative group">
                          <span className="text-xs font-medium text-slate-600 truncate w-full px-2 text-center">{file.name}</span>
                          <button 
                            onClick={() => setAttachedFiles(attachedFiles.filter((_, i) => i !== index))}
                            className="absolute -top-2 -right-2 bg-rose-500 text-white rounded-full p-1 opacity-0 group-hover:opacity-100 transition-opacity"
                          >
                            <X size={12} />
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-2">Ghi chú thêm (Nếu có)</label>
                    <textarea 
                      rows={3} 
                      placeholder="Nhập các bất thường phát hiện tại hiện trường..."
                      className="w-full px-4 py-3 bg-white border border-slate-300 focus:border-blue-500 focus:ring-2 focus:ring-blue-200 rounded-lg text-base outline-none resize-none"
                    ></textarea>
                  </div>



                </div>
                
                {/* Action Buttons */}
                <div className="bg-slate-50 p-5 border-t border-slate-200 flex flex-col sm:flex-row gap-3 justify-end items-center">
                  {saveSuccess && (
                    <span className="text-emerald-600 text-sm font-medium flex items-center gap-1 mr-auto">
                      <CheckCircle size={16} /> Đã lưu vào Google Sheets
                    </span>
                  )}
                  <button 
                    onClick={handleSaveToSheets}
                    disabled={isSaving || !isGoogleConnected}
                    className="px-6 py-3 bg-blue-600 text-white rounded-lg font-medium hover:bg-blue-700 disabled:bg-blue-300 disabled:cursor-not-allowed flex items-center justify-center gap-2 transition-colors shadow-sm"
                  >
                    {isSaving ? (
                      <div className="w-5 h-5 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
                    ) : (
                      <Send size={18} />
                    )}
                    {isSaving ? 'Đang lưu...' : (isGoogleConnected ? 'Lưu & Đồng bộ Sheets' : 'Cần kết nối Google Drive để Lưu')}
                  </button>
                </div>
              </div>
            </div>
          )}
    </div>
  )}

          {/* --- DASHBOARD VIEW (Original) --- */}
          {activeTab === 'dashboard' && (
            <div className="max-w-7xl mx-auto space-y-6">
              {/* DASHBOARD HEADER */}
              <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
                <div>
                  <h1 className="text-xl font-bold text-slate-800">Tổng quan hệ thống</h1>
                  <p className="text-sm text-slate-500 mt-1">Theo dõi trạng thái thiết bị theo thời gian thực</p>
                </div>
                <div className="flex flex-col sm:flex-row items-start sm:items-center gap-3 w-full sm:w-auto">
                  {!(userRole === 'customer' && userFactory) && (
                    <>
                      <div className="flex items-center gap-2 w-full sm:w-auto">
                        <span className="text-sm font-medium text-slate-600 whitespace-nowrap">Khách hàng:</span>
                        <select 
                          value={selectedCustomer}
                          onChange={(e) => {
                            setSelectedCustomer(e.target.value);
                            setSelectedFactory('all'); // Reset factory when customer changes
                          }}
                          className="w-full sm:w-auto border border-slate-300 rounded-lg text-slate-700 bg-slate-50 px-4 py-2 outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-all font-medium"
                        >
                          <option value="all">Tất cả khách hàng</option>
                          {uniqueCustomers.map(customer => (
                            <option key={customer} value={customer}>{customer}</option>
                          ))}
                        </select>
                      </div>
                      <div className="flex items-center gap-2 w-full sm:w-auto">
                        <span className="text-sm font-medium text-slate-600 whitespace-nowrap">Nhà máy:</span>
                        <select 
                          value={selectedFactory}
                          onChange={(e) => setSelectedFactory(e.target.value)}
                          className="w-full sm:w-auto border border-slate-300 rounded-lg text-slate-700 bg-slate-50 px-4 py-2 outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-all font-medium"
                        >
                          <option value="all">Tất cả nhà máy</option>
                          {dynamicSiteData
                            .filter(site => selectedCustomer === 'all' || site.customer === selectedCustomer)
                            .map(site => (
                            <option key={site.id} value={site.id}>{site.name}</option>
                          ))}
                        </select>
                      </div>
                    </>
                  )}
                  {userRole === 'customer' && userFactory && (
                    <div className="flex items-center gap-2 px-4 py-2 bg-blue-50 border border-blue-100 rounded-lg text-blue-700 font-medium text-sm">
                      <Factory size={16} />
                      <span>{userFactory}</span>
                    </div>
                  )}
                </div>
              </div>

              {/* KPI CARDS */}
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-sm flex items-center gap-4 cursor-pointer hover:border-blue-300 transition-colors" onClick={() => { setStatusFilter('all'); handleTabChange('equipment'); }}>
                  <div className="w-12 h-12 rounded-full bg-blue-50 flex items-center justify-center text-blue-600">
                    <Server size={24} />
                  </div>
                  <div>
                    <p className="text-sm font-medium text-slate-500">Tổng thiết bị</p>
                    <p className="text-2xl font-bold text-slate-900">{filteredKpiData.total}</p>
                  </div>
                </div>
                
                <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-sm flex items-center gap-4 cursor-pointer hover:border-emerald-300 transition-colors" onClick={() => { setStatusFilter('healthy'); handleTabChange('equipment'); }}>
                  <div className="w-12 h-12 rounded-full bg-emerald-50 flex items-center justify-center text-emerald-600">
                    <CheckCircle size={24} />
                  </div>
                  <div>
                    <p className="text-sm font-medium text-slate-500">Đạt tiêu chuẩn</p>
                    <p className="text-2xl font-bold text-slate-900">{filteredKpiData.healthy}</p>
                  </div>
                </div>

                <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-sm flex items-center gap-4 cursor-pointer hover:border-amber-300 transition-colors" onClick={() => { setStatusFilter('warning'); handleTabChange('equipment'); }}>
                  <div className="w-12 h-12 rounded-full bg-amber-50 flex items-center justify-center text-amber-600">
                    <AlertTriangle size={24} />
                  </div>
                  <div>
                    <p className="text-sm font-medium text-slate-500">Cảnh báo</p>
                    <p className="text-2xl font-bold text-slate-900">{filteredKpiData.warning}</p>
                  </div>
                </div>

                <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-sm flex items-center gap-4 cursor-pointer hover:border-rose-300 transition-colors" onClick={() => { setStatusFilter('critical'); handleTabChange('equipment'); }}>
                  <div className="w-12 h-12 rounded-full bg-rose-50 flex items-center justify-center text-rose-600">
                    <AlertCircle size={24} />
                  </div>
                  <div>
                    <p className="text-sm font-medium text-slate-500">Nguy hiểm</p>
                    <p className="text-2xl font-bold text-slate-900">{filteredKpiData.critical}</p>
                  </div>
                </div>
              </div>

              {/* RELIABILITY KPIs */}
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6 gap-4">
                <div className="bg-slate-900 text-white rounded-xl p-4 shadow-lg flex flex-col justify-between">
                  <div className="flex justify-between items-start">
                    <p className="text-slate-400 text-[10px] font-bold uppercase tracking-wider">MTBF</p>
                    <TrendingUp size={14} className="text-emerald-400" />
                  </div>
                  <div className="mt-1">
                    <p className="text-xl font-bold">{reliabilityKpis.mtbf.toLocaleString()} <span className="text-[10px] font-normal text-slate-400">giờ</span></p>
                    <p className="text-[9px] text-slate-400 mt-0.5">Thời gian giữa các sự cố</p>
                  </div>
                </div>
                
                <div className="bg-slate-900 text-white rounded-xl p-4 shadow-lg flex flex-col justify-between">
                  <div className="flex justify-between items-start">
                    <p className="text-slate-400 text-[10px] font-bold uppercase tracking-wider">MTTR</p>
                    <TrendingDown size={14} className="text-emerald-400" />
                  </div>
                  <div className="mt-1">
                    <p className="text-xl font-bold">{reliabilityKpis.mttr} <span className="text-[10px] font-normal text-slate-400">giờ</span></p>
                    <p className="text-[9px] text-slate-400 mt-0.5">Thời gian sửa chữa TB</p>
                  </div>
                </div>

                <div className="bg-slate-900 text-white rounded-xl p-4 shadow-lg flex flex-col justify-between">
                  <div className="flex justify-between items-start">
                    <p className="text-slate-400 text-[10px] font-bold uppercase tracking-wider">MTTF</p>
                    <Activity size={14} className="text-blue-400" />
                  </div>
                  <div className="mt-1">
                    <p className="text-xl font-bold">{reliabilityKpis.mttf.toLocaleString()} <span className="text-[10px] font-normal text-slate-400">giờ</span></p>
                    <p className="text-[9px] text-slate-400 mt-0.5">Thời gian đến khi hỏng</p>
                  </div>
                </div>

                <div className="bg-slate-900 text-white rounded-xl p-4 shadow-lg flex flex-col justify-between">
                  <div className="flex justify-between items-start">
                    <p className="text-slate-400 text-[10px] font-bold uppercase tracking-wider">MWT</p>
                    <Clock size={14} className="text-amber-400" />
                  </div>
                  <div className="mt-1">
                    <p className="text-xl font-bold">{reliabilityKpis.mwt} <span className="text-[10px] font-normal text-slate-400">giờ</span></p>
                    <p className="text-[9px] text-slate-400 mt-0.5">Thời gian chờ bảo trì</p>
                  </div>
                </div>

                <div className="bg-slate-900 text-white rounded-xl p-4 shadow-lg flex flex-col justify-between">
                  <div className="flex justify-between items-start">
                    <p className="text-slate-400 text-[10px] font-bold uppercase tracking-wider">Downtime</p>
                    <AlertTriangle size={14} className="text-rose-400" />
                  </div>
                  <div className="mt-1">
                    <p className="text-xl font-bold">{reliabilityKpis.totalBreakdownTime.toLocaleString()} <span className="text-[10px] font-normal text-slate-400">giờ</span></p>
                    <p className="text-[9px] text-slate-400 mt-0.5">Tổng thời gian dừng máy</p>
                  </div>
                </div>

                <div className="bg-slate-900 text-white rounded-xl p-4 shadow-lg flex flex-col justify-between">
                  <div className="flex justify-between items-start">
                    <p className="text-slate-400 text-[10px] font-bold uppercase tracking-wider">Cost</p>
                    <Zap size={14} className="text-emerald-400" />
                  </div>
                  <div className="mt-1">
                    <p className="text-xl font-bold">${reliabilityKpis.totalCost.toLocaleString()}</p>
                    <p className="text-[9px] text-slate-400 mt-0.5">Chi phí bảo trì tổng cộng</p>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="bg-white rounded-xl p-4 border border-slate-200 shadow-sm flex items-center justify-between">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-lg bg-amber-50 flex items-center justify-center text-amber-600">
                      <ClipboardList size={20} />
                    </div>
                    <div>
                      <p className="text-xs font-medium text-slate-500">Backlog</p>
                      <p className="text-lg font-bold text-slate-900">{reliabilityKpis.backlog} phiếu</p>
                    </div>
                  </div>
                  <div className="text-right">
                    <p className="text-[10px] text-slate-400">Công việc tồn đọng</p>
                  </div>
                </div>
                <div className="bg-white rounded-xl p-4 border border-slate-200 shadow-sm flex items-center justify-between">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-lg bg-blue-50 flex items-center justify-center text-blue-600">
                      <CheckCircle size={20} />
                    </div>
                    <div>
                      <p className="text-xs font-medium text-slate-500">Compliance</p>
                      <p className="text-lg font-bold text-slate-900">{reliabilityKpis.compliance}%</p>
                    </div>
                  </div>
                  <div className="text-right">
                    <p className="text-[10px] text-slate-400">Tuân thủ kế hoạch</p>
                  </div>
                </div>
              </div>

              {/* MIDDLE SECTION: MAP & RISK LIST */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                
                {/* Maintenance Strategy Chart */}
                <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                  <div className="p-4 border-b border-slate-100 flex items-center justify-between">
                    <h3 className="font-semibold text-slate-800 flex items-center gap-2">
                      <BarChart2 size={18} className="text-blue-500" />
                      Chiến lược bảo trì
                    </h3>
                  </div>
                  <div className="p-4 h-[300px]">
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart>
                        <Pie
                          data={reliabilityKpis.typeDist}
                          cx="50%"
                          cy="50%"
                          innerRadius={60}
                          outerRadius={80}
                          paddingAngle={5}
                          dataKey="value"
                        >
                          {reliabilityKpis.typeDist.map((entry, index) => (
                            <Cell key={`cell-${index}`} fill={entry.color} />
                          ))}
                        </Pie>
                        <Tooltip />
                        <Legend verticalAlign="bottom" height={36}/>
                      </PieChart>
                    </ResponsiveContainer>
                    <div className="mt-2 text-center">
                      <p className="text-xs text-slate-500 italic">* Phân bổ giữa Bảo trì phòng ngừa và Bảo trì khắc phục</p>
                    </div>
                  </div>
                </div>

                {/* GEOGRAPHICAL MAP */}
                <div className="lg:col-span-2 bg-white rounded-xl border border-slate-200 shadow-sm flex flex-col overflow-hidden">
                  <div className="p-4 border-b border-slate-100 flex flex-wrap items-center justify-between gap-3 z-10 bg-white">
                    <div className="flex items-center gap-2.5">
                      <div className="w-8 h-8 rounded-lg bg-blue-50 border border-blue-200 flex items-center justify-center text-blue-600">
                        <Globe size={18} />
                      </div>
                      <div>
                        <div className="flex items-center gap-2">
                          <h2 className="font-bold text-slate-800 text-sm md:text-base">
                            Bản đồ phân bố thiết bị & Dự án (Sites)
                          </h2>
                          <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-blue-100 text-blue-800 border border-blue-200">
                            {filteredSiteData.length} Sites
                          </span>
                        </div>
                        <p className="text-xs text-slate-500">
                          Hỗ trợ định vị toàn cầu cho các dự án trong nước & quốc tế
                        </p>
                      </div>
                    </div>

                    <div className="flex flex-wrap items-center gap-2">
                      {/* View shortcuts */}
                      <div className="flex items-center gap-1 bg-slate-100 p-0.5 rounded-lg text-xs font-medium">
                        <button
                          type="button"
                          onClick={() => setMapViewTrigger({ type: 'fit-all', timestamp: Date.now() })}
                          className="px-2 py-1 rounded-md text-slate-700 hover:bg-white hover:shadow-xs transition-all flex items-center gap-1 text-[11px]"
                          title="Căn chỉnh vừa vặn tất cả các site hiện có"
                        >
                          <Navigation size={12} className="text-blue-600" />
                          <span>Gom tất cả ({filteredSiteData.length})</span>
                        </button>
                        <button
                          type="button"
                          onClick={() => setMapViewTrigger({ type: 'global', timestamp: Date.now() })}
                          className="px-2 py-1 rounded-md text-slate-700 hover:bg-white hover:shadow-xs transition-all text-[11px]"
                          title="Thu phóng toàn cầu"
                        >
                          🌍 Toàn cầu
                        </button>
                        <button
                          type="button"
                          onClick={() => setMapViewTrigger({ type: 'sea', timestamp: Date.now() })}
                          className="px-2 py-1 rounded-md text-slate-700 hover:bg-white hover:shadow-xs transition-all text-[11px]"
                          title="Khu vực Đông Nam Á"
                        >
                          🌏 Đông Nam Á
                        </button>
                        <button
                          type="button"
                          onClick={() => setMapViewTrigger({ type: 'vn', timestamp: Date.now() })}
                          className="px-2 py-1 rounded-md text-slate-700 hover:bg-white hover:shadow-xs transition-all text-[11px]"
                          title="Khu vực Việt Nam"
                        >
                          🇻🇳 Việt Nam
                        </button>
                      </div>

                      {/* Map Tile Layers (OSM & Esri - No API Key Required) */}
                      <div className="flex items-center gap-1 bg-slate-100 p-0.5 rounded-lg text-xs font-medium">
                        <button
                          type="button"
                          onClick={() => setMapLayer('streets')}
                          className={`px-2 py-1 rounded-md transition-all text-[11px] ${
                            mapLayer === 'streets' || (mapLayer as string) === 'voyager'
                              ? 'bg-white text-blue-600 font-bold shadow-xs'
                              : 'text-slate-600 hover:text-slate-900'
                          }`}
                        >
                          Đường bộ (OSM)
                        </button>
                        <button
                          type="button"
                          onClick={() => setMapLayer('light')}
                          className={`px-2 py-1 rounded-md transition-all text-[11px] ${
                            mapLayer === 'light'
                              ? 'bg-white text-blue-600 font-bold shadow-xs'
                              : 'text-slate-600 hover:text-slate-900'
                          }`}
                        >
                          Nền sáng (Light)
                        </button>
                        <button
                          type="button"
                          onClick={() => setMapLayer('satellite')}
                          className={`px-2 py-1 rounded-md transition-all text-[11px] ${
                            mapLayer === 'satellite'
                              ? 'bg-white text-blue-600 font-bold shadow-xs'
                              : 'text-slate-600 hover:text-slate-900'
                          }`}
                        >
                          Vệ tinh (Satellite)
                        </button>
                        <button
                          type="button"
                          onClick={() => setMapLayer('topo')}
                          className={`px-2 py-1 rounded-md transition-all text-[11px] ${
                            mapLayer === 'topo'
                              ? 'bg-white text-blue-600 font-bold shadow-xs'
                              : 'text-slate-600 hover:text-slate-900'
                          }`}
                        >
                          Địa hình (Topo)
                        </button>
                      </div>
                    </div>
                  </div>
                  <div className="flex-1 min-h-[420px] relative z-0">
                    <MapContainer center={[16.047079, 108.206230]} zoom={5} style={{ height: '100%', width: '100%', zIndex: 0 }}>
                      <MapBoundsHandler sites={filteredSiteData} viewTrigger={mapViewTrigger} />
                      {mapLayer === 'satellite' ? (
                        <TileLayer
                          key="satellite"
                          attribution='Tiles &copy; Esri &mdash; Source: Esri, Maxar, Earthstar Geographics'
                          url="https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}"
                          maxZoom={18}
                        />
                      ) : mapLayer === 'light' ? (
                        <TileLayer
                          key="light"
                          attribution='Tiles &copy; Esri &mdash; Esri, DeLorme, NAVTEQ'
                          url="https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Base/MapServer/tile/{z}/{y}/{x}"
                          maxZoom={16}
                        />
                      ) : mapLayer === 'topo' ? (
                        <TileLayer
                          key="topo"
                          attribution='Tiles &copy; Esri &mdash; USGS, Esri'
                          url="https://server.arcgisonline.com/ArcGIS/rest/services/World_Topo_Map/MapServer/tile/{z}/{y}/{x}"
                          maxZoom={18}
                        />
                      ) : (
                        <TileLayer
                          key="streets"
                          attribution='&copy; <a href="https://www.openstreetmap.org/copyright" target="_blank" rel="noreferrer">OpenStreetMap</a> contributors'
                          url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
                          maxZoom={19}
                        />
                      )}
                      {filteredSiteData.map((site) => (
                        <Marker 
                          key={site.id} 
                          position={[site.lat, site.lng]} 
                          icon={createCustomIcon(site.status, site.count)}
                        >
                          <Popup className="custom-popup">
                            <div className="p-1 min-w-[210px]">
                              <div className="flex items-center gap-1.5 mb-1.5 border-b pb-1.5">
                                <span className="text-base">{site.flag || '🌐'}</span>
                                <div>
                                  <h3 className="font-bold text-slate-800 text-sm leading-tight">{site.name}</h3>
                                  <span className="text-[10px] text-slate-500 font-medium">{site.country || 'Quốc tế'}</span>
                                </div>
                              </div>
                              <div className="space-y-1.5 text-xs">
                                <div className="flex justify-between items-center text-slate-600">
                                  <span>Tổng thiết bị:</span>
                                  <span className="font-bold text-slate-900">{site.count}</span>
                                </div>
                                <div className="flex justify-between items-center text-emerald-700">
                                  <span className="flex items-center gap-1"><div className="w-2 h-2 rounded-full bg-emerald-500"></div> Khỏe mạnh:</span>
                                  <span className="font-semibold">{site.healthy}</span>
                                </div>
                                <div className="flex justify-between items-center text-amber-700">
                                  <span className="flex items-center gap-1"><div className="w-2 h-2 rounded-full bg-amber-500"></div> Cảnh báo:</span>
                                  <span className="font-semibold">{site.warning}</span>
                                </div>
                                <div className="flex justify-between items-center text-rose-700">
                                  <span className="flex items-center gap-1"><div className="w-2 h-2 rounded-full bg-rose-500"></div> Nguy hiểm:</span>
                                  <span className="font-semibold">{site.critical}</span>
                                </div>
                                <div className="pt-1.5 border-t border-slate-100 flex items-center justify-between text-[10px] text-slate-400 font-mono">
                                  <span>Tọa độ:</span>
                                  <span>{site.lat.toFixed(3)}, {site.lng.toFixed(3)}</span>
                                </div>
                              </div>
                            </div>
                          </Popup>
                          <LeafletTooltip direction="top" offset={[0, -20]} opacity={1}>
                            <span className="font-semibold">{site.flag || ''} {site.name}</span>
                          </LeafletTooltip>
                        </Marker>
                      ))}
                    </MapContainer>
                  </div>
                </div>

                {/* TOP RISK EQUIPMENT */}
                <div className="lg:col-span-1 bg-white rounded-xl border border-slate-200 shadow-sm flex flex-col">
                  <div className="p-5 border-b border-slate-100 flex items-center justify-between">
                    <h2 className="font-semibold text-slate-800 flex items-center gap-2">
                      <AlertTriangle className="text-rose-500" size={18} />
                      Top 10% rủi ro cao nhất
                    </h2>
                  </div>
                  <div className="p-2 flex-1 overflow-y-auto max-h-[400px]">
                    {filteredTopRiskEquipment.length > 0 ? filteredTopRiskEquipment.map((item) => (
                      <div 
                        key={item.id} 
                        onClick={() => setSelectedRiskDetail(item.id)}
                        className="p-3 hover:bg-slate-50 rounded-lg transition-colors cursor-pointer border-b border-slate-50 last:border-0"
                      >
                        <div className="flex justify-between items-start mb-1">
                          <span className="font-medium text-slate-900 text-sm">{item.name}</span>
                          <StatusBadge status={item.status} />
                        </div>
                        <div className="flex items-center gap-3 text-xs text-slate-500 mb-2">
                          <span className="flex items-center gap-1"><Factory size={12} /> {item.factory}</span>
                          <span className="flex items-center gap-1"><MapPin size={12} /> {item.location}</span>
                        </div>
                        <div className="flex items-center justify-between">
                          <span className="text-xs font-medium text-slate-700">Health Index:</span>
                          <div className="w-32">
                            <HealthBar value={item.health} status={item.status} />
                          </div>
                        </div>
                      </div>
                    )) : (
                      <div className="p-8 text-center text-slate-500 flex flex-col items-center justify-center h-full">
                        <CheckCircle size={32} className="text-emerald-400 mb-2" />
                        <p>Không có thiết bị rủi ro cao</p>
                      </div>
                    )}
                  </div>
                  <div className="p-3 border-t border-slate-100 text-center">
                    <button className="text-sm text-blue-600 font-medium hover:text-blue-800 transition-colors">
                      Xem tất cả cảnh báo
                    </button>
                  </div>
                </div>
              </div>

              {/* TRENDING & GAUGE SECTION */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                {/* TRENDING CHART */}
                <div className="lg:col-span-2 bg-white rounded-xl border border-slate-200 shadow-sm flex flex-col">
                  <div className="p-5 border-b border-slate-100 flex items-center justify-between">
                    <div>
                      <h2 className="font-semibold text-slate-800 flex items-center gap-2">
                        <TrendingUp className="text-blue-500" size={18} />
                        Biểu đồ xu hướng (Trending)
                      </h2>
                      <div className="flex items-center gap-3 mt-2">
                        <select 
                          value={trendingEquipment} 
                          onChange={(e) => {
                            setTrendingEquipment(e.target.value);
                            const eq = allEquipment.find(item => item.id === e.target.value);
                            if (eq) {
                              if (eq.type === 'Máy biến áp') {
                                setTrendingChartType('temp');
                              } else if (eq.type === 'Động cơ') {
                                setTrendingChartType('ir');
                              } else if (eq.type === 'Tủ điện') {
                                setTrendingChartType('tev');
                              } else {
                                setTrendingChartType('temp');
                              }
                            }
                          }}
                          className="text-xs border-slate-200 rounded-md text-slate-600 bg-slate-50 px-2 py-1.5 outline-none focus:ring-2 focus:ring-blue-500 font-medium"
                        >
                          {allEquipment.map(eq => (
                            <option key={eq.id} value={eq.id}>{eq.name} ({eq.id})</option>
                          ))}
                        </select>
                        <select 
                          value={trendingChartType} 
                          onChange={(e) => setTrendingChartType(e.target.value)}
                          className="text-xs border-slate-200 rounded-md text-slate-600 bg-slate-50 px-2 py-1.5 outline-none focus:ring-2 focus:ring-blue-500 font-medium"
                        >
                          {allEquipment.find(eq => eq.id === trendingEquipment)?.type === 'Máy biến áp' && (
                            <>
                              <option value="temp">Nhiệt độ dầu</option>
                              <option value="winding_temp">Nhiệt độ cuộn dây</option>
                              <option value="dga">Phân tích khí (DGA)</option>
                            </>
                          )}
                          {allEquipment.find(eq => eq.id === trendingEquipment)?.type === 'Động cơ' && (
                            <>
                              <option value="ir">Điện trở cách điện</option>
                              <option value="winding_res">Điện trở cuộn dây</option>
                              <option value="pd">PD</option>
                              <option value="elcid">ELCID</option>
                              <option value="pi">PI</option>
                            </>
                          )}
                          {allEquipment.find(eq => eq.id === trendingEquipment)?.type === 'Tủ điện' && (
                            <>
                              <option value="tev">PD (TEV)</option>
                            </>
                          )}
                          {!['Máy biến áp', 'Động cơ', 'Tủ điện'].includes(allEquipment.find(eq => eq.id === trendingEquipment)?.type || '') && (
                             <option value="temp">Nhiệt độ</option>
                          )}
                        </select>
                      </div>
                    </div>
                    <select className="text-sm border-slate-200 rounded-md text-slate-600 bg-slate-50 px-3 py-1.5 outline-none focus:ring-2 focus:ring-blue-500">
                      <option>Hôm nay</option>
                      <option>7 ngày qua</option>
                      <option>30 ngày qua</option>
                      <option>1 năm qua</option>
                    </select>
                  </div>
                  <div className="p-5 flex-1 min-h-[300px]">
                    <ResponsiveContainer width="100%" height="100%">
                      {trendingChartType === 'temp' || trendingChartType === 'winding_temp' ? (
                        <AreaChart data={trendingChartType === 'temp' ? dynamicTrendData.temp : dynamicTrendData.winding_temp} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
                          <defs>
                            <linearGradient id="colorTemp" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="5%" stopColor="#3b82f6" stopOpacity={0.3}/>
                              <stop offset="95%" stopColor="#3b82f6" stopOpacity={0}/>
                            </linearGradient>
                          </defs>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                          <XAxis dataKey="time" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dy={10} />
                          <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dx={-10} />
                          <Tooltip 
                            contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                            labelStyle={{ color: '#64748b', marginBottom: '4px' }}
                          />
                          <ReferenceLine y={90} label={{ position: 'top', value: 'Ngưỡng cảnh báo (90°C)', fill: '#ef4444', fontSize: 12 }} stroke="#ef4444" strokeDasharray="3 3" />
                          <Area type="monotone" dataKey="temp" name={trendingChartType === 'temp' ? "Nhiệt độ dầu (°C)" : "Nhiệt độ cuộn dây (°C)"} stroke="#3b82f6" strokeWidth={3} fillOpacity={1} fill="url(#colorTemp)" activeDot={{ r: 6, strokeWidth: 0, fill: '#3b82f6' }} />
                        </AreaChart>
                      ) : trendingChartType === 'tev' ? (
                        <AreaChart data={dynamicTrendData.tev} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
                          <defs>
                            <linearGradient id="colorTev" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="5%" stopColor="#6366f1" stopOpacity={0.3}/>
                              <stop offset="95%" stopColor="#6366f1" stopOpacity={0}/>
                            </linearGradient>
                            <linearGradient id="colorUltra" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="5%" stopColor="#10b981" stopOpacity={0.3}/>
                              <stop offset="95%" stopColor="#10b981" stopOpacity={0}/>
                            </linearGradient>
                          </defs>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                          <XAxis dataKey="time" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dy={10} />
                          <YAxis yAxisId="left" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dx={-10} />
                          <YAxis yAxisId="right" orientation="right" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dx={10} />
                          <Tooltip 
                            contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                            labelStyle={{ color: '#64748b', marginBottom: '4px' }}
                          />
                          <Legend verticalAlign="top" height={36} />
                          <ReferenceLine yAxisId="left" y={20} label={{ position: 'insideTopLeft', value: 'Ngưỡng TEV (20 dBmV)', fill: '#ef4444', fontSize: 10 }} stroke="#ef4444" strokeDasharray="3 3" />
                          <ReferenceLine yAxisId="right" y={10} label={{ position: 'insideTopRight', value: 'Ngưỡng Siêu âm (10 dBµV)', fill: '#f59e0b', fontSize: 10 }} stroke="#f59e0b" strokeDasharray="3 3" />
                          <Area yAxisId="left" type="monotone" dataKey="tev" name="TEV (dBmV)" stroke="#6366f1" strokeWidth={3} fillOpacity={1} fill="url(#colorTev)" activeDot={{ r: 6, strokeWidth: 0, fill: '#6366f1' }} />
                          <Area yAxisId="right" type="monotone" dataKey="ultrasonic" name="Siêu âm (dBµV)" stroke="#10b981" strokeWidth={3} fillOpacity={1} fill="url(#colorUltra)" activeDot={{ r: 6, strokeWidth: 0, fill: '#10b981' }} />
                        </AreaChart>
                      ) : trendingChartType === 'dga' ? (
                        <LineChart data={dynamicTrendData.dga} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                          <XAxis dataKey="time" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dy={10} />
                          <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dx={-10} />
                          <Tooltip 
                            contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                            labelStyle={{ color: '#64748b', marginBottom: '4px' }}
                          />
                          <Legend verticalAlign="top" height={36} />
                          <Line type="monotone" dataKey="h2" name="H2 (ppm)" stroke="#3b82f6" strokeWidth={2} dot={{ r: 4 }} activeDot={{ r: 6 }} />
                          <Line type="monotone" dataKey="ch4" name="CH4 (ppm)" stroke="#10b981" strokeWidth={2} dot={{ r: 4 }} activeDot={{ r: 6 }} />
                          <Line type="monotone" dataKey="c2h6" name="C2H6 (ppm)" stroke="#f59e0b" strokeWidth={2} dot={{ r: 4 }} activeDot={{ r: 6 }} />
                          <Line type="monotone" dataKey="c2h4" name="C2H4 (ppm)" stroke="#ef4444" strokeWidth={2} dot={{ r: 4 }} activeDot={{ r: 6 }} />
                          <Line type="monotone" dataKey="c2h2" name="C2H2 (ppm)" stroke="#8b5cf6" strokeWidth={2} dot={{ r: 4 }} activeDot={{ r: 6 }} />
                          <Line type="monotone" dataKey="co" name="CO (ppm)" stroke="#64748b" strokeWidth={2} dot={{ r: 4 }} activeDot={{ r: 6 }} />
                          <Line type="monotone" dataKey="co2" name="CO2 (ppm)" stroke="#0ea5e9" strokeWidth={2} dot={{ r: 4 }} activeDot={{ r: 6 }} />
                          <Line type="monotone" dataKey="o2" name="O2 (ppm)" stroke="#22c55e" strokeWidth={2} dot={{ r: 4 }} activeDot={{ r: 6 }} />
                          <Line type="monotone" dataKey="n2" name="N2 (ppm)" stroke="#94a3b8" strokeWidth={2} dot={{ r: 4 }} activeDot={{ r: 6 }} />
                        </LineChart>
                      ) : trendingChartType === 'ir' ? (
                        <AreaChart data={dynamicTrendData.ir} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
                          <defs>
                            <linearGradient id="colorIr" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="5%" stopColor="#10b981" stopOpacity={0.3}/>
                              <stop offset="95%" stopColor="#10b981" stopOpacity={0}/>
                            </linearGradient>
                          </defs>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                          <XAxis dataKey="time" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dy={10} />
                          <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dx={-10} />
                          <Tooltip 
                            contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                            labelStyle={{ color: '#64748b', marginBottom: '4px' }}
                          />
                          <ReferenceLine y={100} label={{ position: 'top', value: 'Ngưỡng cảnh báo (100 MΩ)', fill: '#ef4444', fontSize: 12 }} stroke="#ef4444" strokeDasharray="3 3" />
                          <Area type="monotone" dataKey="ir" name="Điện trở cách điện (MΩ)" stroke="#10b981" strokeWidth={3} fillOpacity={1} fill="url(#colorIr)" activeDot={{ r: 6, strokeWidth: 0, fill: '#10b981' }} />
                        </AreaChart>
                      ) : trendingChartType === 'winding_res' ? (
                        <AreaChart data={dynamicTrendData.winding_res} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
                          <defs>
                            <linearGradient id="colorRes" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="5%" stopColor="#f59e0b" stopOpacity={0.3}/>
                              <stop offset="95%" stopColor="#f59e0b" stopOpacity={0}/>
                            </linearGradient>
                          </defs>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                          <XAxis dataKey="time" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dy={10} />
                          <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dx={-10} />
                          <Tooltip 
                            contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                            labelStyle={{ color: '#64748b', marginBottom: '4px' }}
                          />
                          <Area type="monotone" dataKey="res" name="Điện trở cuộn dây (Ω)" stroke="#f59e0b" strokeWidth={3} fillOpacity={1} fill="url(#colorRes)" activeDot={{ r: 6, strokeWidth: 0, fill: '#f59e0b' }} />
                        </AreaChart>
                      ) : trendingChartType === 'pd' ? (
                        <AreaChart data={dynamicTrendData.pd} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
                          <defs>
                            <linearGradient id="colorPd" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="5%" stopColor="#8b5cf6" stopOpacity={0.3}/>
                              <stop offset="95%" stopColor="#8b5cf6" stopOpacity={0}/>
                            </linearGradient>
                          </defs>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                          <XAxis dataKey="time" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dy={10} />
                          <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dx={-10} />
                          <Tooltip 
                            contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                            labelStyle={{ color: '#64748b', marginBottom: '4px' }}
                          />
                          <ReferenceLine y={40} label={{ position: 'top', value: 'Ngưỡng cảnh báo (40 pC)', fill: '#ef4444', fontSize: 12 }} stroke="#ef4444" strokeDasharray="3 3" />
                          <Area type="monotone" dataKey="pd" name="PD (pC)" stroke="#8b5cf6" strokeWidth={3} fillOpacity={1} fill="url(#colorPd)" activeDot={{ r: 6, strokeWidth: 0, fill: '#8b5cf6' }} />
                        </AreaChart>
                      ) : trendingChartType === 'elcid' ? (
                        <AreaChart data={dynamicTrendData.elcid} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
                          <defs>
                            <linearGradient id="colorElcid" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="5%" stopColor="#ec4899" stopOpacity={0.3}/>
                              <stop offset="95%" stopColor="#ec4899" stopOpacity={0}/>
                            </linearGradient>
                          </defs>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                          <XAxis dataKey="time" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dy={10} />
                          <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dx={-10} />
                          <Tooltip 
                            contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                            labelStyle={{ color: '#64748b', marginBottom: '4px' }}
                          />
                          <ReferenceLine y={100} label={{ position: 'top', value: 'Ngưỡng cảnh báo (100 mA)', fill: '#ef4444', fontSize: 12 }} stroke="#ef4444" strokeDasharray="3 3" />
                          <Area type="monotone" dataKey="elcid" name="ELCID (mA)" stroke="#ec4899" strokeWidth={3} fillOpacity={1} fill="url(#colorElcid)" activeDot={{ r: 6, strokeWidth: 0, fill: '#ec4899' }} />
                        </AreaChart>
                      ) : (
                        <AreaChart data={dynamicTrendData.pi} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
                          <defs>
                            <linearGradient id="colorPi" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="5%" stopColor="#06b6d4" stopOpacity={0.3}/>
                              <stop offset="95%" stopColor="#06b6d4" stopOpacity={0}/>
                            </linearGradient>
                          </defs>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                          <XAxis dataKey="time" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dy={10} />
                          <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dx={-10} />
                          <Tooltip 
                            contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                            labelStyle={{ color: '#64748b', marginBottom: '4px' }}
                          />
                          <ReferenceLine y={2.0} label={{ position: 'top', value: 'Ngưỡng an toàn (2.0)', fill: '#10b981', fontSize: 12 }} stroke="#10b981" strokeDasharray="3 3" />
                          <Area type="monotone" dataKey="pi" name="Chỉ số phân cực (PI)" stroke="#06b6d4" strokeWidth={3} fillOpacity={1} fill="url(#colorPi)" activeDot={{ r: 6, strokeWidth: 0, fill: '#06b6d4' }} />
                        </AreaChart>
                      )}
                    </ResponsiveContainer>
                  </div>
                </div>

                {/* FACTORY HEALTH GAUGE */}
                <div className="lg:col-span-1 bg-white rounded-xl border border-slate-200 shadow-sm flex flex-col">
                  <div className="p-5 border-b border-slate-100 flex items-center justify-between">
                    <h2 className="font-semibold text-slate-800 flex items-center gap-2">
                      <Gauge className="text-indigo-500" size={18} />
                      Sức khỏe toàn nhà máy
                    </h2>
                  </div>
                  <div className="p-5 flex-1 min-h-[300px] flex flex-col items-center justify-center relative">
                    <div className="h-48 w-full">
                      <ResponsiveContainer width="100%" height="100%">
                        <PieChart>
                          <Pie
                            data={[
                              { name: 'Score', value: averageHealth, fill: healthColor },
                              { name: 'Remaining', value: 100 - averageHealth, fill: '#f1f5f9' }
                            ]}
                            cx="50%"
                            cy="100%"
                            startAngle={180}
                            endAngle={0}
                            innerRadius={100}
                            outerRadius={120}
                            dataKey="value"
                            stroke="none"
                          >
                            <Cell key="cell-0" fill={healthColor} />
                            <Cell key="cell-1" fill="#f1f5f9" />
                          </Pie>
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                    <div className="absolute bottom-12 text-center">
                      <div className="text-5xl font-black text-slate-800">{averageHealth}<span className="text-2xl text-slate-400">%</span></div>
                      <div className={`text-base font-medium mt-1 ${healthTextColor}`}>{healthText}</div>
                    </div>
                  </div>
                </div>
              </div>

              {/* CHARTS SECTION 2 */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                {/* STATUS PIE CHART */}
                <div className="lg:col-span-1 bg-white rounded-xl border border-slate-200 shadow-sm flex flex-col">
                  <div className="p-5 border-b border-slate-100 flex items-center justify-between">
                    <h2 className="font-semibold text-slate-800 flex items-center gap-2">
                      <PieChartIcon className="text-blue-500" size={18} />
                      Phân bổ trạng thái
                    </h2>
                  </div>
                  <div className="p-5 flex-1 min-h-[250px] flex items-center justify-center">
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart>
                        <Pie
                          data={filteredStatusDistribution}
                          cx="50%"
                          cy="50%"
                          innerRadius={60}
                          outerRadius={80}
                          paddingAngle={5}
                          dataKey="value"
                        >
                          {filteredStatusDistribution.map((entry, index) => (
                            <Cell key={`cell-${index}`} fill={entry.color} />
                          ))}
                        </Pie>
                        <Tooltip 
                          contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                        />
                        <Legend verticalAlign="bottom" height={36} iconType="circle" />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                {/* HEALTH DISTRIBUTION BAR CHART */}
                <div className="lg:col-span-2 bg-white rounded-xl border border-slate-200 shadow-sm flex flex-col">
                  <div className="p-5 border-b border-slate-100 flex items-center justify-between">
                    <h2 className="font-semibold text-slate-800 flex items-center gap-2">
                      <BarChartIcon className="text-emerald-500" size={18} />
                      Phân bố Health Index
                    </h2>
                  </div>
                  <div className="p-5 flex-1 min-h-[250px]">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={filteredHealthDistribution} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                        <XAxis dataKey="range" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dy={10} />
                        <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dx={-10} />
                        <Tooltip 
                          cursor={{fill: '#f8fafc'}}
                          contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                        />
                        <Bar dataKey="count" fill="#3b82f6" radius={[4, 4, 0, 0]} maxBarSize={40}>
                          {filteredHealthDistribution.map((entry, index) => (
                            <Cell key={`cell-${index}`} fill={
                              entry.range === '<50%' ? '#f43f5e' : 
                              entry.range === '50-69%' ? '#f59e0b' : 
                              entry.range === '70-89%' ? '#3b82f6' : '#10b981'
                            } />
                          ))}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </div>

              {/* CMMS WORK ORDERS MANAGEMENT & REAL-TIME STATUS */}
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="p-5 border-b border-slate-100 flex flex-col gap-4 bg-gradient-to-r from-slate-50 via-white to-blue-50/30">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                    <div>
                      <div className="flex items-center gap-2.5">
                        <div className="w-9 h-9 rounded-lg bg-blue-600 text-white flex items-center justify-center shadow-xs">
                          <ClipboardList size={20} />
                        </div>
                        <div>
                          <div className="flex items-center gap-2">
                            <h2 className="font-bold text-slate-800 text-base">
                              Phiếu công việc CMMS &amp; Bảo trì hiện trường
                            </h2>
                            <span className="px-2 py-0.5 rounded-full text-xs font-bold bg-blue-100 text-blue-800 border border-blue-200">
                              {dashboardFilteredWorkOrders.length} / {workOrders.length} Phiếu
                            </span>
                          </div>
                          <p className="text-xs text-slate-500 mt-0.5">
                            Theo dõi tiến độ xử lý sự cố, bảo dưỡng định kỳ và dữ liệu CMMS đồng bộ Google Sheets
                          </p>
                        </div>
                      </div>
                    </div>
                    
                    <div className="flex flex-wrap items-center gap-2">
                      <button
                        onClick={() => {
                          setCmmsSubTab('table');
                          handleTabChange('cmms');
                        }}
                        className="text-xs sm:text-sm bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 px-3 py-1.5 rounded-md font-semibold transition-colors flex items-center gap-1.5 shadow-2xs"
                        title="Chuyển đến màn hình Quản lý CMMS chuyên sâu"
                      >
                        <span>Mở phân hệ CMMS đầy đủ</span>
                        <ArrowUpRight size={14} />
                      </button>

                      <button
                        onClick={() => {
                          setNewWorkOrder({
                            title: '',
                            description: '',
                            equipmentId: [],
                            customerId: selectedCustomer !== 'all' ? selectedCustomer : '',
                            factory: '',
                            priority: 'Medium',
                            type: 'corrective',
                            status: 'initiated',
                            assignedTo: user?.email || '',
                            blockingRequired: false,
                            workPermitId: '',
                            isUnplanned: true,
                            failureCode: '',
                            rootCause: '',
                            actualTimeSpent: 0,
                            pmFrequency: '',
                            estimatedTime: 2,
                            laborCount: 1,
                            laborCost: 0,
                            partCost: 0,
                            attachments: [],
                            downtimeStart: '',
                            repairStart: '',
                            repairEnd: '',
                            restartTime: ''
                          });
                          setShowWorkOrderModal(true);
                        }}
                        className="text-xs sm:text-sm bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded-md font-semibold transition-colors flex items-center gap-1.5 shadow-xs"
                      >
                        <Plus size={14} />
                        <span>Tạo phiếu WO</span>
                      </button>

                      <button
                        onClick={handleFetchFromSheets}
                        disabled={!isGoogleConnected || isSyncing}
                        className="text-xs sm:text-sm bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 px-2.5 py-1.5 rounded-md font-medium transition-colors flex items-center gap-1.5"
                        title="Tải lại phiếu từ Google Sheet TEV Service Flatform"
                      >
                        <RefreshCw size={13} className={isSyncing ? 'animate-spin text-blue-600' : 'text-slate-500'} />
                        <span className="hidden sm:inline">Làm mới Sheets</span>
                      </button>
                    </div>
                  </div>

                  {/* Quick KPI stats row for CMMS */}
                  <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 pt-2 border-t border-slate-200/60">
                    <button
                      onClick={() => setDashboardWoFilter('all')}
                      className={`p-2.5 rounded-lg border text-left transition-all ${
                        dashboardWoFilter === 'all'
                          ? 'bg-blue-50 border-blue-300 ring-1 ring-blue-400'
                          : 'bg-white border-slate-200 hover:bg-slate-50'
                      }`}
                    >
                      <p className="text-[11px] font-medium text-slate-500">Tất cả phiếu</p>
                      <p className="text-lg font-bold text-slate-900">{workOrders.length}</p>
                    </button>

                    <button
                      onClick={() => setDashboardWoFilter('initiated')}
                      className={`p-2.5 rounded-lg border text-left transition-all ${
                        dashboardWoFilter === 'initiated'
                          ? 'bg-sky-50 border-sky-300 ring-1 ring-sky-400'
                          : 'bg-white border-slate-200 hover:bg-slate-50'
                      }`}
                    >
                      <p className="text-[11px] font-medium text-sky-600 flex items-center gap-1">
                        <Clock size={12} /> Mới tạo / Chờ duyệt
                      </p>
                      <p className="text-lg font-bold text-sky-700">{reliabilityKpis.initiatedCount}</p>
                    </button>

                    <button
                      onClick={() => setDashboardWoFilter('in-progress')}
                      className={`p-2.5 rounded-lg border text-left transition-all ${
                        dashboardWoFilter === 'in-progress'
                          ? 'bg-amber-50 border-amber-300 ring-1 ring-amber-400'
                          : 'bg-white border-slate-200 hover:bg-slate-50'
                      }`}
                    >
                      <p className="text-[11px] font-medium text-amber-600 flex items-center gap-1">
                        <Activity size={12} /> Đang xử lý (Active)
                      </p>
                      <p className="text-lg font-bold text-amber-700">{reliabilityKpis.inProgressCount}</p>
                    </button>

                    <button
                      onClick={() => setDashboardWoFilter('corrective')}
                      className={`p-2.5 rounded-lg border text-left transition-all ${
                        dashboardWoFilter === 'corrective'
                          ? 'bg-rose-50 border-rose-300 ring-1 ring-rose-400'
                          : 'bg-white border-slate-200 hover:bg-slate-50'
                      }`}
                    >
                      <p className="text-[11px] font-medium text-rose-600 flex items-center gap-1">
                        <AlertTriangle size={12} /> Sự cố đột xuất (CM)
                      </p>
                      <p className="text-lg font-bold text-rose-700">
                        {workOrders.filter(wo => {
                          const t = (wo.type || '').toLowerCase();
                          return t.includes('corrective') || t.includes('đột xuất') || t.includes('cm') || wo.isUnplanned;
                        }).length}
                      </p>
                    </button>

                    <button
                      onClick={() => setDashboardWoFilter('preventive')}
                      className={`p-2.5 rounded-lg border text-left transition-all ${
                        dashboardWoFilter === 'preventive'
                          ? 'bg-emerald-50 border-emerald-300 ring-1 ring-emerald-400'
                          : 'bg-white border-slate-200 hover:bg-slate-50'
                      }`}
                    >
                      <p className="text-[11px] font-medium text-emerald-600 flex items-center gap-1">
                        <CheckCircle size={12} /> Định kỳ (PM)
                      </p>
                      <p className="text-lg font-bold text-emerald-700">
                        {workOrders.filter(wo => {
                          const t = (wo.type || '').toLowerCase();
                          return t.includes('preventive') || t.includes('định kỳ') || t.includes('pm');
                        }).length}
                      </p>
                    </button>
                  </div>

                  {/* Filter & Search Bar */}
                  <div className="flex flex-col sm:flex-row items-center justify-between gap-3 pt-1">
                    <div className="flex flex-wrap items-center gap-1.5 w-full sm:w-auto">
                      <span className="text-xs text-slate-500 font-medium mr-1">Lọc nhanh:</span>
                      {[
                        { id: 'all', label: 'Tất cả' },
                        { id: 'initiated', label: 'Mới tạo' },
                        { id: 'in-progress', label: 'Đang làm' },
                        { id: 'corrective', label: 'Sửa đột xuất (CM)' },
                        { id: 'preventive', label: 'Bảo trì định kỳ (PM)' },
                        { id: 'overdue', label: 'Quá hạn' },
                      ].map(f => (
                        <button
                          key={f.id}
                          onClick={() => setDashboardWoFilter(f.id as any)}
                          className={`px-2.5 py-1 rounded-md text-xs font-medium transition-colors ${
                            dashboardWoFilter === f.id
                              ? 'bg-slate-800 text-white shadow-2xs font-semibold'
                              : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                          }`}
                        >
                          {f.label}
                        </button>
                      ))}
                    </div>

                    <div className="relative w-full sm:w-72">
                      <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
                      <input
                        type="text"
                        placeholder="Tìm WO, thiết bị, khách hàng..."
                        value={dashboardWoSearch}
                        onChange={(e) => setDashboardWoSearch(e.target.value)}
                        className="w-full text-xs pl-8 pr-3 py-1.5 border border-slate-200 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-500"
                      />
                      {dashboardWoSearch && (
                        <button
                          onClick={() => setDashboardWoSearch('')}
                          className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 text-xs"
                        >
                          ✕
                        </button>
                      )}
                    </div>
                  </div>
                </div>

                {/* Table of Work Orders */}
                <div className="overflow-x-auto max-h-[460px]">
                  <table className="w-full text-left border-collapse">
                    <thead className="sticky top-0 bg-slate-100 border-b border-slate-200 text-[11px] uppercase tracking-wider text-slate-600 font-semibold z-10">
                      <tr>
                        <th className="p-3">Mã WO &amp; Tiêu đề</th>
                        <th className="p-3">Thiết bị &amp; Khách hàng</th>
                        <th className="p-3">Loại hình</th>
                        <th className="p-3">Ưu tiên</th>
                        <th className="p-3">Trạng thái</th>
                        <th className="p-3">Phụ trách</th>
                        <th className="p-3">Hạn chót</th>
                        <th className="p-3 text-right">Thao tác</th>
                      </tr>
                    </thead>
                    <tbody className="text-xs divide-y divide-slate-100">
                      {dashboardFilteredWorkOrders.length > 0 ? (
                        dashboardFilteredWorkOrders.map((wo) => {
                          const eqStr = Array.isArray(wo.equipmentId) ? wo.equipmentId.join(', ') : (wo.equipmentId || 'Chưa gán');
                          const isHigh = wo.priority === 'Critical' || wo.priority === 'High' || wo.priority === 'Khẩn cấp';
                          const statusLower = (wo.status || '').toLowerCase();
                          const isDone = statusLower.includes('complete') || statusLower.includes('hoàn thành');
                          const isInProg = statusLower.includes('progress') || statusLower.includes('đang');
                          
                          return (
                            <tr key={wo.id} className="hover:bg-slate-50 transition-colors">
                              <td className="p-3">
                                <div className="font-mono font-bold text-blue-700 flex items-center gap-1.5">
                                  <span>{wo.id}</span>
                                  {wo.workPermitId && (
                                    <span className="text-[10px] font-normal text-slate-400 font-sans">
                                      ({wo.workPermitId})
                                    </span>
                                  )}
                                </div>
                                <div className="font-medium text-slate-900 mt-0.5 line-clamp-1" title={wo.title}>
                                  {wo.title}
                                </div>
                                {wo.description && (
                                  <div className="text-[11px] text-slate-500 line-clamp-1" title={wo.description}>
                                    {wo.description}
                                  </div>
                                )}
                              </td>
                              <td className="p-3">
                                <div className="font-semibold text-slate-800 flex items-center gap-1">
                                  <span className="px-1.5 py-0.5 rounded bg-slate-100 border border-slate-200 text-slate-700 font-mono text-[10px]">
                                    {eqStr}
                                  </span>
                                  {wo.equipmentName && (
                                    <span className="text-[11px] text-slate-600 line-clamp-1">
                                      {wo.equipmentName}
                                    </span>
                                  )}
                                </div>
                                <div className="text-[11px] text-slate-500 mt-0.5 flex items-center gap-1">
                                  <Factory size={11} className="text-slate-400" />
                                  <span className="font-medium text-slate-700">{wo.customer || 'Chưa phân loại'}</span>
                                </div>
                              </td>
                              <td className="p-3">
                                <span className={`px-2 py-0.5 rounded text-[10px] font-semibold ${
                                  (wo.type || '').toLowerCase().includes('corrective') || (wo.type || '').toLowerCase().includes('đột xuất') || wo.isUnplanned
                                    ? 'bg-rose-50 text-rose-700 border border-rose-200'
                                    : (wo.type || '').toLowerCase().includes('preventive') || (wo.type || '').toLowerCase().includes('định kỳ')
                                    ? 'bg-blue-50 text-blue-700 border border-blue-200'
                                    : 'bg-slate-100 text-slate-700 border border-slate-200'
                                }`}>
                                  {wo.type || (wo.isUnplanned ? 'Sửa đột xuất (CM)' : 'Định kỳ (PM)')}
                                </span>
                              </td>
                              <td className="p-3">
                                <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                                  isHigh
                                    ? 'bg-rose-100 text-rose-800 border border-rose-300'
                                    : wo.priority === 'Medium' || wo.priority === 'Trung bình'
                                    ? 'bg-amber-100 text-amber-800 border border-amber-300'
                                    : 'bg-slate-100 text-slate-700 border border-slate-200'
                                }`}>
                                  {wo.priority || 'Medium'}
                                </span>
                              </td>
                              <td className="p-3">
                                <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold ${
                                  isDone
                                    ? 'bg-emerald-100 text-emerald-800 border border-emerald-200'
                                    : isInProg
                                    ? 'bg-amber-100 text-amber-800 border border-amber-200 animate-pulse'
                                    : 'bg-sky-100 text-sky-800 border border-sky-200'
                                }`}>
                                  <span className={`w-1.5 h-1.5 rounded-full ${
                                    isDone ? 'bg-emerald-600' : isInProg ? 'bg-amber-600' : 'bg-sky-600'
                                  }`}></span>
                                  {isDone ? 'Hoàn thành' : isInProg ? 'Đang xử lý' : (wo.status || 'Mới tạo')}
                                </span>
                              </td>
                              <td className="p-3 text-slate-600">
                                <div className="text-[11px] font-medium text-slate-800">{wo.assignedTo || 'Chưa phân công'}</div>
                                {wo.responsibleDo && <div className="text-[10px] text-slate-400">{wo.responsibleDo}</div>}
                              </td>
                              <td className="p-3 font-mono text-slate-600 text-[11px]">
                                {wo.dueDate ? (
                                  <span className={new Date(wo.dueDate).getTime() < Date.now() && !isDone ? 'text-rose-600 font-bold' : ''}>
                                    {wo.dueDate}
                                  </span>
                                ) : (
                                  <span className="text-slate-400">-</span>
                                )}
                              </td>
                              <td className="p-3 text-right">
                                <div className="flex items-center justify-end gap-1.5">
                                  {!isDone && (
                                    <button
                                      onClick={() => handleQuickUpdateWorkOrderStatus(wo.id, isInProg ? 'completed' : 'in_progress')}
                                      className="px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 text-[11px] font-medium rounded transition-colors"
                                      title={isInProg ? 'Đánh dấu hoàn thành' : 'Chuyển sang Đang xử lý'}
                                    >
                                      {isInProg ? '✓ Hoàn thành' : '▶ Bắt đầu'}
                                    </button>
                                  )}
                                  <button
                                    onClick={() => {
                                      setSelectedWorkOrder(wo);
                                      setCmmsSubTab('table');
                                      handleTabChange('cmms');
                                    }}
                                    className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded transition-colors"
                                    title="Xem chi tiết phiếu trong CMMS"
                                  >
                                    <Eye size={16} />
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })
                      ) : (
                        <tr>
                          <td colSpan={8} className="p-8 text-center text-slate-500">
                            <div className="flex flex-col items-center justify-center gap-2">
                              <ClipboardList size={32} className="text-slate-300" />
                              <p className="text-sm font-medium">Không tìm thấy phiếu CMMS nào phù hợp với bộ lọc</p>
                              <p className="text-xs text-slate-400">
                                Thử chọn lại khách hàng &quot;Tất cả&quot; hoặc xóa từ khóa tìm kiếm
                              </p>
                            </div>
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>

                <div className="p-3 border-t border-slate-100 bg-slate-50 flex flex-wrap items-center justify-between text-xs text-slate-500">
                  <div className="flex items-center gap-2">
                    <span>Hiển thị <strong>{dashboardFilteredWorkOrders.length}</strong> / <strong>{workOrders.length}</strong> phiếu công việc.</span>
                    {selectedCustomer !== 'all' && (
                      <span className="text-blue-600 font-semibold">(Đang lọc theo: {selectedCustomer})</span>
                    )}
                  </div>
                  <div className="flex items-center gap-3">
                    <button
                      onClick={() => {
                        setCmmsSubTab('station');
                        handleTabChange('cmms');
                      }}
                      className="text-blue-600 hover:text-blue-800 font-semibold transition-colors flex items-center gap-1"
                    >
                      <span>Trạm nhập liệu &amp; Thí nghiệm hiện trường</span>
                      <ArrowUpRight size={13} />
                    </button>
                  </div>
                </div>
              </div>

              {/* ALL EQUIPMENT TABLE */}
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="p-5 border-b border-slate-100 flex flex-col gap-4">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                    <div>
                      <div className="flex items-center gap-2">
                        <h2 className="font-semibold text-slate-800">Danh sách toàn bộ thiết bị & Mô phỏng</h2>
                        <span className="px-2 py-0.5 rounded-full text-xs font-bold bg-blue-100 text-blue-700">
                          {allEquipment.length} Thiết bị
                        </span>
                      </div>
                      <p className="text-xs text-slate-500 mt-0.5 flex items-center gap-1.5">
                        <span className={`w-2 h-2 rounded-full ${isGoogleConnected && isAutoSyncEnabled ? 'bg-emerald-500 animate-pulse' : 'bg-slate-400'}`}></span>
                        <span>Luồng 2 chiều với Google Sheet <strong>TEV Service Flatform</strong>: {isGoogleConnected ? (lastAutoSyncTime ? `Đã đồng bộ (${lastAutoSyncTime})` : 'Tự động đồng bộ mỗi 60s') : 'Chưa kết nối'}</span>
                      </p>
                    </div>
                    <div className="flex flex-wrap items-center gap-2">
                      <button 
                        onClick={() => handlePerformTwoWaySync(false)}
                        disabled={!isGoogleConnected || isSyncing || autoSyncStatus === 'syncing'}
                        className="text-xs sm:text-sm bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded-md font-bold transition-all shadow-sm disabled:opacity-50 flex items-center gap-1.5"
                        title="Đẩy toàn bộ dữ liệu & biểu đồ mô phỏng lên Google Sheet TEV Service Flatform và kéo cập nhật mới nhất về app"
                      >
                        <RefreshCw size={14} className={isSyncing || autoSyncStatus === 'syncing' ? 'animate-spin' : ''} />
                        <span>{isSyncing || autoSyncStatus === 'syncing' ? 'Đang đồng bộ...' : 'Đồng bộ 2 chiều ngay'}</span>
                      </button>

                      <button 
                        onClick={handleOpenReconciliation}
                        disabled={!isGoogleConnected || isReconciling}
                        className="text-xs sm:text-sm bg-amber-50 border border-amber-200 text-amber-800 hover:bg-amber-100 px-3 py-1.5 rounded-md font-medium transition-colors disabled:opacity-50 flex items-center gap-1.5"
                        title="Đối soát phiên bản dữ liệu (Reconciliation) giữa App và Sheets"
                      >
                        <ShieldAlert size={14} className="text-amber-600" />
                        <span>Đối soát phiên bản</span>
                      </button>

                      {googleSheetUrl && (
                        <a
                          href={googleSheetUrl}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="text-xs sm:text-sm bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 px-3 py-1.5 rounded-md font-medium transition-colors flex items-center gap-1.5"
                          title="Mở trực tiếp file Google Sheet 'TEV Service Flatform'"
                        >
                          <ExternalLink size={14} className="text-slate-500" />
                          <span>Mở file Sheets</span>
                        </a>
                      )}
                    </div>
                  </div>
                  
                  {/* Filters */}
                  <div className="grid grid-cols-1 md:grid-cols-5 gap-3 bg-slate-50 p-3 rounded-lg border border-slate-200">
                    <div>
                      <label className="block text-xs font-medium text-slate-500 mb-1">Tên thiết bị</label>
                      <input 
                        type="text" 
                        placeholder="Tìm tên..." 
                        value={filterName}
                        onChange={(e) => setFilterName(e.target.value)}
                        className="w-full text-sm border border-slate-300 rounded-md px-3 py-1.5 focus:outline-none focus:ring-2 focus:ring-blue-500"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-slate-500 mb-1">Nhà máy / Vị trí</label>
                      <input 
                        type="text" 
                        placeholder="Tìm vị trí..." 
                        value={filterLocation}
                        onChange={(e) => setFilterLocation(e.target.value)}
                        className="w-full text-sm border border-slate-300 rounded-md px-3 py-1.5 focus:outline-none focus:ring-2 focus:ring-blue-500"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-slate-500 mb-1">Loại thiết bị</label>
                      <select 
                        value={filterType}
                        onChange={(e) => setFilterType(e.target.value)}
                        className="w-full text-sm border border-slate-300 rounded-md px-3 py-1.5 focus:outline-none focus:ring-2 focus:ring-blue-500"
                      >
                        <option value="all">Tất cả</option>
                        <option value="Máy biến áp">Máy biến áp</option>
                        <option value="Động cơ">Động cơ</option>
                        <option value="Tủ điện">Tủ điện</option>
                        <option value="Máy phát">Máy phát</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-slate-500 mb-1">Trạng thái</label>
                      <select 
                        value={statusFilter}
                        onChange={(e) => setStatusFilter(e.target.value)}
                        className="w-full text-sm border border-slate-300 rounded-md px-3 py-1.5 focus:outline-none focus:ring-2 focus:ring-blue-500"
                      >
                        <option value="all">Tất cả</option>
                        <option value="healthy">Đạt tiêu chuẩn</option>
                        <option value="warning">Cảnh báo</option>
                        <option value="critical">Nguy hiểm</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-slate-500 mb-1">Chỉ số sức khỏe</label>
                      <div className="flex items-center gap-2">
                        <input 
                          type="number" 
                          placeholder="Min" 
                          value={filterHealthMin}
                          onChange={(e) => setFilterHealthMin(e.target.value)}
                          className="w-full text-sm border border-slate-300 rounded-md px-2 py-1.5 focus:outline-none focus:ring-2 focus:ring-blue-500"
                        />
                        <span className="text-slate-400">-</span>
                        <input 
                          type="number" 
                          placeholder="Max" 
                          value={filterHealthMax}
                          onChange={(e) => setFilterHealthMax(e.target.value)}
                          className="w-full text-sm border border-slate-300 rounded-md px-2 py-1.5 focus:outline-none focus:ring-2 focus:ring-blue-500"
                        />
                      </div>
                    </div>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="bg-slate-50 border-b border-slate-200 text-xs uppercase tracking-wider text-slate-500 font-semibold">
                        <th className="p-4">Mã TB</th>
                        <th className="p-4">Tên thiết bị</th>
                        <th className="p-4">Nhà máy / Vị trí</th>
                        <th className="p-4">Loại</th>
                        <th className="p-4">Trạng thái</th>
                        <th className="p-4 w-48">Chỉ số sức khỏe</th>
                        <th className="p-4 text-right">Thao tác</th>
                      </tr>
                    </thead>
                    <tbody className="text-sm divide-y divide-slate-100">
                      {paginatedEquipment.map((eq) => (
                        <tr key={eq.id} className="hover:bg-slate-50 transition-colors">
                          <td className="p-4 font-mono text-slate-500">{eq.id}</td>
                          <td className="p-4 font-medium text-slate-900">{eq.name}</td>
                          <td className="p-4">
                            <div className="text-slate-900">{eq.factory}</div>
                            <div className="text-xs text-slate-500">{eq.location}</div>
                          </td>
                          <td className="p-4 text-slate-600">{eq.type}</td>
                          <td className="p-4">
                            <span className={`px-2 py-1 rounded-md text-xs font-bold ${
                              eq.criticality === 'A' ? 'bg-rose-100 text-rose-700 border border-rose-200' :
                              eq.criticality === 'B' ? 'bg-amber-100 text-amber-700 border border-amber-200' :
                              'bg-slate-100 text-slate-700 border border-slate-200'
                            }`}>
                              {eq.criticality || 'B'}
                            </span>
                          </td>
                          <td className="p-4">
                            <StatusBadge status={eq.status} />
                          </td>
                          <td className="p-4">
                            <HealthBar value={eq.health} status={eq.status} />
                          </td>
                          <td className="p-4 text-right">
                            <div className="flex items-center justify-end gap-2">
                              <button 
                                onClick={() => handleViewEquipmentProfile(eq)}
                                className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded transition-colors" 
                                title="Xem chi tiết"
                              >
                                <Eye size={18} />
                              </button>
                              <button 
                                onClick={() => { setSelectedEqForQR(eq); setShowQRModal(true); }}
                                className="p-1.5 text-slate-400 hover:text-emerald-600 hover:bg-emerald-50 rounded transition-colors" 
                                title="Tạo mã QR"
                              >
                                <QrCode size={18} />
                              </button>
                              <button className="p-1.5 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded transition-colors">
                                <MoreVertical size={18} />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div className="p-4 border-t border-slate-100 flex items-center justify-between text-sm text-slate-500">
                  <span>Hiển thị {displayedEquipment.length > 0 ? (currentPage - 1) * itemsPerPage + 1 : 0}-{Math.min(currentPage * itemsPerPage, displayedEquipment.length)} trong số {displayedEquipment.length} thiết bị</span>
                  <div className="flex gap-1">
                    <button 
                      onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
                      disabled={currentPage === 1}
                      className="px-3 py-1 border border-slate-200 rounded hover:bg-slate-50 disabled:opacity-50"
                    >
                      Trước
                    </button>
                    <button className="px-3 py-1 border border-slate-200 rounded bg-blue-50 text-blue-600 font-medium">{currentPage}</button>
                    <button 
                      onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
                      disabled={currentPage === totalPages || totalPages === 0}
                      className="px-3 py-1 border border-slate-200 rounded hover:bg-slate-50 disabled:opacity-50"
                    >
                      Sau
                    </button>
                  </div>
                </div>
              </div>

            </div>
          )}
          {/* --- EQUIPMENT MANAGEMENT VIEW --- */}
          {activeTab === 'equipment' && (
            <div className="max-w-7xl mx-auto space-y-6">
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden flex flex-col min-h-[500px] lg:h-[calc(100vh-8rem)]">
                <div className="p-5 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                  <div className="flex flex-col gap-1">
                    <h2 className="font-semibold text-slate-800 flex items-center gap-2">
                      <Server className="text-blue-500" size={20} />
                      Quản lý Thiết bị
                    </h2>
                    <div className="flex items-center gap-2 text-xs text-slate-400">
                      <span>Hệ thống</span>
                      <ChevronRight size={12} />
                      <span className={selectedCustomer === 'all' ? 'text-blue-500 font-medium' : ''}>Tất cả khách hàng</span>
                      {selectedCustomer !== 'all' && (
                        <>
                          <ChevronRight size={12} />
                          <span className={selectedFactory === 'all' ? 'text-blue-500 font-medium' : ''}>{selectedCustomer}</span>
                        </>
                      )}
                      {selectedFactory !== 'all' && (
                        <>
                          <ChevronRight size={12} />
                          <span className="text-blue-500 font-medium">{dynamicSiteData.find(s => s.id === selectedFactory)?.name}</span>
                        </>
                      )}
                    </div>
                  </div>
                  <div className="flex items-center gap-3">
                    <div className="relative">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                      <input 
                        type="text" 
                        placeholder="Tìm mã thiết bị, tên..." 
                        className="pl-9 pr-4 py-2 bg-slate-50 border border-slate-200 focus:bg-white focus:border-blue-500 focus:ring-2 focus:ring-blue-200 rounded-lg text-sm w-full sm:w-64 transition-all outline-none"
                        value={filterName}
                        onChange={(e) => { setFilterName(e.target.value); setCurrentPage(1); }}
                      />
                    </div>
                    <button onClick={() => setShowAddEqModal(true)} className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg text-sm font-medium hover:bg-blue-700 transition-colors shadow-sm">
                      <span className="text-lg leading-none">+</span> Thêm thiết bị
                    </button>
                  </div>
                </div>
                
                {/* Filters */}
                <div className="px-5 py-3 bg-slate-50 border-b border-slate-100 flex flex-wrap items-center gap-4 text-sm">
                  <div className="flex items-center gap-2">
                    <span className="text-slate-500 font-medium">Lọc theo:</span>
                  </div>
                  {!(userRole === 'customer' && userFactory) && (
                    <select 
                      className="bg-white border border-slate-200 rounded-md px-3 py-1.5 text-slate-700 outline-none focus:border-blue-500"
                      value={filterLocation}
                      onChange={(e) => { setFilterLocation(e.target.value); setCurrentPage(1); }}
                    >
                      <option value="">Tất cả nhà máy</option>
                      {dynamicSiteData
                        .filter(site => selectedCustomer === 'all' || site.customer === selectedCustomer)
                        .map(site => (
                        <option key={site.id} value={site.name}>{site.name}</option>
                      ))}
                    </select>
                  )}
                  <select 
                    className="bg-white border border-slate-200 rounded-md px-3 py-1.5 text-slate-700 outline-none focus:border-blue-500"
                    value={filterType}
                    onChange={(e) => { setFilterType(e.target.value); setCurrentPage(1); }}
                  >
                    <option value="all">Tất cả loại thiết bị</option>
                    <option value="Máy biến áp">Máy biến áp</option>
                    <option value="Động cơ">Động cơ điện</option>
                    <option value="Máy phát">Máy phát điện</option>
                    <option value="Tủ điện">Tủ điện</option>
                    <option value="Inverter">Inverter Solar/Wind</option>
                  </select>
                  <select 
                    className="bg-white border border-slate-200 rounded-md px-3 py-1.5 text-slate-700 outline-none focus:border-blue-500"
                    value={statusFilter}
                    onChange={(e) => { setStatusFilter(e.target.value); setCurrentPage(1); }}
                  >
                    <option value="all">Tất cả trạng thái</option>
                    <option value="healthy">Khỏe mạnh</option>
                    <option value="warning">Cảnh báo</option>
                    <option value="critical">Nguy hiểm</option>
                  </select>
                </div>

                {/* Equipment Table */}
                <div className="flex-1 overflow-auto">
                  <table className="w-full text-left border-collapse min-w-[800px]">
                    <thead className="sticky top-0 bg-slate-50 z-10 shadow-sm">
                      <tr className="border-b border-slate-200 text-xs uppercase tracking-wider text-slate-500 font-semibold">
                        <th className="p-4">Mã TB</th>
                        <th className="p-4">Tên thiết bị</th>
                        <th className="p-4">Nhà máy / Vị trí</th>
                        <th className="p-4">Loại</th>
                        <th className="p-4">Criticality</th>
                        <th className="p-4">Trạng thái</th>
                        <th className="p-4 w-48">Chỉ số sức khỏe</th>
                        <th className="p-4">Kiểm tra lần cuối</th>
                        <th className="p-4 text-right">Thao tác</th>
                      </tr>
                    </thead>
                    <tbody className="text-sm divide-y divide-slate-100">
                      {paginatedEquipment.map((eq) => (
                        <tr key={eq.id} className="hover:bg-slate-50 transition-colors">
                          <td className="p-4 font-mono text-slate-500">{eq.id}</td>
                          <td className="p-4 font-medium text-slate-900">{eq.name}</td>
                          <td className="p-4">
                            <div className="text-slate-900">{eq.factory}</div>
                            <div className="text-xs text-slate-500">{eq.location}</div>
                          </td>
                          <td className="p-4 text-slate-600">{eq.type}</td>
                          <td className="p-4">
                            <span className={`px-2 py-1 rounded-md text-xs font-bold ${
                              eq.criticality === 'A' ? 'bg-rose-100 text-rose-700 border border-rose-200' :
                              eq.criticality === 'B' ? 'bg-amber-100 text-amber-700 border border-amber-200' :
                              'bg-slate-100 text-slate-700 border border-slate-200'
                            }`}>
                              {eq.criticality || 'B'}
                            </span>
                          </td>
                          <td className="p-4">
                            <StatusBadge status={eq.status} />
                          </td>
                          <td className="p-4">
                            <HealthBar value={eq.health} status={eq.status} />
                          </td>
                          <td className="p-4 text-slate-500">{eq.lastCheck}</td>
                          <td className="p-4 text-right">
                            <div className="flex items-center justify-end gap-2">
                              <button 
                                onClick={() => {
                                  const reportData = {
                                    id: `PREVIEW-${eq.id}`,
                                    date: new Date().toLocaleDateString('vi-VN'),
                                    equipmentId: eq.id,
                                    equipmentName: eq.name,
                                    factory: eq.factory,
                                    type: eq.type,
                                    inspector: user?.displayName || 'FSE',
                                    status: eq.status,
                                    health: eq.health,
                                    notes: 'Báo cáo xem trước từ dữ liệu hiện tại.',
                                    measurements: eq.rawData ? (
                                      eq.type === 'Máy biến áp' ? {
                                        oilTemp: eq.rawData[9], windingTemp: eq.rawData[10], irHighLow: eq.rawData[11],
                                        irHighEarth: eq.rawData[12], oilLeak: eq.rawData[13], dga: eq.rawData[14],
                                        dielectricStrength: eq.rawData[15], furan: eq.rawData[16], oilMoisture: eq.rawData[17]
                                      } : eq.type === 'Động cơ' ? {
                                        vibration: eq.rawData[9], statorTemp: eq.rawData[10], ir: eq.rawData[11],
                                        pd: eq.rawData[12], voltageImbalance: eq.rawData[13], pi: eq.rawData[14],
                                        bearingTemp: eq.rawData[15], tanDelta: eq.rawData[16]
                                      } : {
                                        thermography: eq.rawData[9], contactRes: eq.rawData[10], tev: eq.rawData[11],
                                        ultrasonic: eq.rawData[12], tevPulses: eq.rawData[13], humidity: eq.rawData[14],
                                        sf6Pressure: eq.rawData[15]
                                      }
                                    ) : {}
                                  };
                                  generateIndividualReportPDF(reportData, true, allReports.filter(r => r.equipmentId === eq.id));
                                }}
                                className="p-1.5 text-slate-400 hover:text-emerald-600 hover:bg-emerald-50 rounded transition-colors" title="Xem báo cáo hiện tại"
                              >
                                <FileText size={16} />
                              </button>
                              <button 
                                onClick={() => handleViewEquipmentProfile(eq)}
                                className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded transition-colors" title="Hồ sơ thiết bị"
                              >
                                <History size={16} />
                              </button>
                              <button 
                                onClick={() => { setSelectedEqForQR(eq); setShowQRModal(true); }}
                                className="p-1.5 text-slate-400 hover:text-emerald-600 hover:bg-emerald-50 rounded transition-colors" 
                                title="Tạo mã QR"
                              >
                                <QrCode size={16} />
                              </button>
                              <button className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded transition-colors" title="Chỉnh sửa">
                                <Settings size={16} />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                
                {/* Pagination */}
                <div className="p-4 border-t border-slate-100 flex items-center justify-between text-sm text-slate-500 bg-white">
                  <span>Hiển thị {displayedEquipment.length > 0 ? (currentPage - 1) * itemsPerPage + 1 : 0}-{Math.min(currentPage * itemsPerPage, displayedEquipment.length)} trong số {displayedEquipment.length} thiết bị</span>
                  <div className="flex gap-1">
                    <button 
                      onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
                      disabled={currentPage === 1}
                      className="px-3 py-1 border border-slate-200 rounded hover:bg-slate-50 disabled:opacity-50"
                    >
                      Trước
                    </button>
                    <button className="px-3 py-1 border border-slate-200 rounded bg-blue-50 text-blue-600 font-medium">{currentPage}</button>
                    <button 
                      onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
                      disabled={currentPage === totalPages || totalPages === 0}
                      className="px-3 py-1 border border-slate-200 rounded hover:bg-slate-50 disabled:opacity-50"
                    >
                      Sau
                    </button>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* --- REPORTS VIEW --- */}
          {activeTab === 'reports' && (
            <div className="max-w-7xl mx-auto space-y-6">
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden flex flex-col min-h-[500px] lg:h-[calc(100vh-8rem)]">
                <div className="p-5 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                  <h2 className="font-semibold text-slate-800 flex items-center gap-2">
                    <FileText className="text-blue-500" size={20} />
                    Kho Báo cáo & Tài liệu
                  </h2>
                  <div className="flex flex-col sm:flex-row items-start sm:items-center gap-3">
                    <p className="text-xs text-slate-500 italic mb-2 sm:mb-0">
                      * Báo cáo tại đây là dữ liệu lịch sử. Để xem báo cáo mới nhất, hãy dùng tab "Quản lý Thiết bị".
                    </p>
                    <button 
                      onClick={() => {
                        alert('Vui lòng chọn thiết bị trong tab "Quản lý Thiết bị" và nhấn biểu tượng Báo cáo để tạo báo cáo mới nhất.');
                        handleTabChange('equipment');
                      }}
                      className="flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-lg text-sm font-medium hover:bg-emerald-700 transition-colors shadow-sm"
                    >
                      <Plus size={16} />
                      Tạo báo cáo mới
                    </button>
                    <div className="relative">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                      <input 
                        type="text" 
                        value={reportFilterSearch}
                        onChange={(e) => {
                          setReportFilterSearch(e.target.value);
                          setReportCurrentPage(1);
                        }}
                        placeholder="Tìm mã báo cáo, thiết bị..." 
                        className="pl-9 pr-4 py-2 bg-slate-50 border border-slate-200 focus:bg-white focus:border-blue-500 focus:ring-2 focus:ring-blue-200 rounded-lg text-sm w-full sm:w-64 transition-all outline-none"
                      />
                    </div>
                    <button 
                      onClick={() => generateListReportPDF(displayedReports, user)}
                      className="flex items-center gap-2 px-4 py-2 bg-blue-50 text-blue-600 rounded-lg text-sm font-medium hover:bg-blue-100 transition-colors"
                    >
                      <Download size={16} />
                      Xuất danh sách
                    </button>
                  </div>
                </div>
                
                {/* Filters */}
                <div className="px-5 py-3 bg-slate-50 border-b border-slate-100 flex flex-wrap items-center gap-4 text-sm">
                  <div className="flex items-center gap-2">
                    <span className="text-slate-500 font-medium">Lọc theo:</span>
                  </div>
                  {!(userRole === 'customer' && userFactory) && (
                    <select 
                      value={reportFilterFactory}
                      onChange={(e) => {
                        setReportFilterFactory(e.target.value);
                        setReportCurrentPage(1);
                      }}
                      className="bg-white border border-slate-200 rounded-md px-3 py-1.5 text-slate-700 outline-none focus:border-blue-500"
                    >
                      <option value="all">Tất cả nhà máy</option>
                      {[...new Set(allEquipment.map(eq => eq.factory))].map(factory => (
                        <option key={factory} value={factory}>{factory}</option>
                      ))}
                    </select>
                  )}
                  <select 
                    value={reportFilterType}
                    onChange={(e) => {
                      setReportFilterType(e.target.value);
                      setReportCurrentPage(1);
                    }}
                    className="bg-white border border-slate-200 rounded-md px-3 py-1.5 text-slate-700 outline-none focus:border-blue-500"
                  >
                    <option value="all">Tất cả loại thiết bị</option>
                    <option value="Máy biến áp">Máy biến áp</option>
                    <option value="Động cơ">Động cơ điện</option>
                    <option value="Tủ điện">Tủ điện</option>
                    <option value="Inverter">Inverter Solar/Wind</option>
                  </select>
                  <select 
                    value={reportFilterStatus}
                    onChange={(e) => {
                      setReportFilterStatus(e.target.value);
                      setReportCurrentPage(1);
                    }}
                    className="bg-white border border-slate-200 rounded-md px-3 py-1.5 text-slate-700 outline-none focus:border-blue-500"
                  >
                    <option value="all">Tất cả trạng thái</option>
                    <option value="healthy">Khỏe mạnh</option>
                    <option value="warning">Cảnh báo</option>
                    <option value="critical">Nguy hiểm</option>
                  </select>
                  <div className="flex items-center gap-2 w-full lg:w-auto lg:ml-auto">
                    <span className="text-slate-500">Thời gian:</span>
                    <div className="flex items-center gap-2 flex-1 lg:flex-none">
                      <input 
                        type="date" 
                        value={reportFilterStartDate}
                        onChange={(e) => {
                          setReportFilterStartDate(e.target.value);
                          setReportCurrentPage(1);
                        }}
                        className="bg-white border border-slate-200 rounded-md px-2 py-1.5 text-slate-700 outline-none focus:border-blue-500 text-xs w-full" 
                      />
                      <span className="text-slate-400">-</span>
                      <input 
                        type="date" 
                        value={reportFilterEndDate}
                        onChange={(e) => {
                          setReportFilterEndDate(e.target.value);
                          setReportCurrentPage(1);
                        }}
                        className="bg-white border border-slate-200 rounded-md px-2 py-1.5 text-slate-700 outline-none focus:border-blue-500 text-xs w-full" 
                      />
                    </div>
                  </div>
                </div>

                {/* Reports Table */}
                <div className="flex-1 overflow-auto">
                  <table className="w-full text-left border-collapse min-w-[800px]">
                    <thead className="sticky top-0 bg-slate-50 z-10 shadow-sm">
                      <tr className="border-b border-slate-200 text-xs uppercase tracking-wider text-slate-500 font-semibold">
                        <th className="p-4">Mã Báo Cáo</th>
                        <th className="p-4">Ngày kiểm tra</th>
                        <th className="p-4">Thiết bị</th>
                        <th className="p-4">Người thực hiện</th>
                        <th className="p-4">Đánh giá</th>
                        <th className="p-4">Ghi chú</th>
                        <th className="p-4 text-right">Thao tác</th>
                      </tr>
                    </thead>
                    <tbody className="text-sm divide-y divide-slate-100">
                      {currentReports.length > 0 ? currentReports.map((report, index) => (
                        <tr key={`${report.id}-${index}`} className="hover:bg-slate-50 transition-colors">
                          <td className="p-4 font-medium text-blue-600">{report.id}</td>
                          <td className="p-4 flex items-center gap-2"><Calendar size={14} className="text-slate-400"/> {report.date}</td>
                          <td className="p-4">
                            <div className="font-medium text-slate-900">{report.equipmentName}</div>
                            <div className="text-xs text-slate-500 font-mono">{report.equipmentId}</div>
                          </td>
                          <td className="p-4 text-slate-700">{report.inspector}</td>
                          <td className="p-4"><StatusBadge status={report.status} /></td>
                          <td className="p-4 text-slate-500 max-w-[200px] truncate" title={report.notes}>{report.notes}</td>
                          <td className="p-4 text-right">
                            <div className="flex items-center justify-end gap-2">
                              <button className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded transition-colors" title="Xem chi tiết">
                                <Eye size={16} />
                              </button>
                              <button 
                                onClick={() => generateIndividualReportPDF(report, true, allReports.filter(r => r.equipmentId === report.equipmentId))}
                                className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded transition-colors" 
                                title="Tải PDF"
                              >
                                <Download size={16} />
                              </button>
                            </div>
                          </td>
                        </tr>
                      )) : (
                        <tr>
                          <td colSpan={7} className="p-8 text-center text-slate-500">
                            Không tìm thấy báo cáo nào phù hợp với bộ lọc.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
                
                {/* Pagination */}
                <div className="p-4 border-t border-slate-100 flex items-center justify-between text-sm text-slate-500 bg-white">
                  <span>
                    {displayedReports.length > 0 
                      ? `Hiển thị ${(reportCurrentPage - 1) * reportsPerPage + 1}-${Math.min(reportCurrentPage * reportsPerPage, displayedReports.length)} trong số ${displayedReports.length} báo cáo`
                      : 'Không có báo cáo nào'}
                  </span>
                  <div className="flex gap-1">
                    <button 
                      onClick={() => setReportCurrentPage(prev => Math.max(prev - 1, 1))}
                      disabled={reportCurrentPage === 1}
                      className="px-3 py-1 border border-slate-200 rounded hover:bg-slate-50 disabled:opacity-50"
                    >
                      Trước
                    </button>
                    {Array.from({ length: Math.min(5, totalReportPages) }, (_, i) => {
                      let pageNum = i + 1;
                      if (totalReportPages > 5 && reportCurrentPage > 3) {
                        pageNum = reportCurrentPage - 2 + i;
                        if (pageNum > totalReportPages) pageNum = totalReportPages - (4 - i);
                      }
                      return (
                        <button 
                          key={pageNum}
                          onClick={() => setReportCurrentPage(pageNum)}
                          className={`px-3 py-1 border rounded ${reportCurrentPage === pageNum ? 'bg-blue-50 text-blue-600 font-medium border-blue-200' : 'border-slate-200 hover:bg-slate-50'}`}
                        >
                          {pageNum}
                        </button>
                      );
                    })}
                    <button 
                      onClick={() => setReportCurrentPage(prev => Math.min(prev + 1, totalReportPages))}
                      disabled={reportCurrentPage === totalReportPages || totalReportPages === 0}
                      className="px-3 py-1 border border-slate-200 rounded hover:bg-slate-50 disabled:opacity-50"
                    >
                      Sau
                    </button>
                  </div>
                </div>
              </div>
            </div>
          )}
          {activeTab === 'deep-analysis' && (
            <div className="flex-1 overflow-y-auto p-8 bg-slate-50/50">
              <div className={`${deepAnalysisSubTab === 'pv-cell' || deepAnalysisSubTab === 'dga' ? 'max-w-7xl' : 'max-w-5xl'} mx-auto space-y-6`}>
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div>
                    <h2 className="text-2xl font-bold text-slate-900 tracking-tight">Phân tích chuyên sâu</h2>
                    <p className="text-slate-500 mt-1">
                      {deepAnalysisSubTab === 'dga' && 'Phân tích tình trạng & Xu hướng nồng độ khí hòa tan DGA (IEEE C57.104 & Duval Triangles)'}
                      {deepAnalysisSubTab === 'pv-cell' && 'Phân tích tình trạng tấm pin/cell mặt trời từ ảnh nhiệt (IR) & ảnh thực tế (RGB) theo IEC 62446-3 & Volateq / Sitemark'}
                      {deepAnalysisSubTab === 'wind-turbine' && 'Phân tích tình trạng cánh quạt điện gió (Wind Turbine Blade)'}
                    </p>
                  </div>
                  <div className="flex items-center gap-2 bg-white p-1 rounded-xl border border-slate-200 shadow-sm">
                    <button 
                      onClick={() => setDeepAnalysisSubTab('dga')}
                      className={`px-4 py-2 rounded-lg text-sm font-medium transition-all ${deepAnalysisSubTab === 'dga' ? 'bg-slate-900 text-white shadow-md' : 'text-slate-600 hover:bg-slate-50'}`}
                    >
                      DGA MBA
                    </button>
                    <button 
                      onClick={() => setDeepAnalysisSubTab('pv-cell')}
                      className={`px-4 py-2 rounded-lg text-sm font-medium transition-all ${deepAnalysisSubTab === 'pv-cell' ? 'bg-slate-900 text-white shadow-md' : 'text-slate-600 hover:bg-slate-50'}`}
                    >
                      PV Cell
                    </button>
                    <button 
                      onClick={() => setDeepAnalysisSubTab('wind-turbine')}
                      className={`px-4 py-2 rounded-lg text-sm font-medium transition-all ${deepAnalysisSubTab === 'wind-turbine' ? 'bg-slate-900 text-white shadow-md' : 'text-slate-600 hover:bg-slate-50'}`}
                    >
                      Wind Blade
                    </button>
                  </div>
                  {deepAnalysisSubTab === 'dga' && dgaAnalysisResult && (
                    <button 
                      onClick={exportToPDF}
                      className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg font-medium hover:bg-blue-700 transition-colors shadow-sm"
                    >
                      <Download size={18} />
                      Xuất báo cáo PDF
                    </button>
                  )}
                </div>

                {deepAnalysisSubTab === 'dga' && (
                  <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                  {/* Input Form */}
                  <div className="lg:col-span-1 bg-white rounded-xl shadow-sm border border-slate-200 p-5">
                    <div className="flex items-start justify-between gap-2 mb-3">
                      <div>
                        <h3 className="text-base font-bold text-slate-800 flex items-center gap-2">
                          <Thermometer size={18} className="text-blue-500" />
                          Nhập liệu &amp; Xu hướng DGA
                        </h3>
                        <p className="text-[11px] text-slate-500 mt-0.5">
                          Chỉ báo mũi tên (▲/▼) đối soát tức thời với số liệu kỳ đo trước trong cơ sở dữ liệu lịch sử
                        </p>
                      </div>
                    </div>

                    {/* Historical Baseline Reference Header Card */}
                    <div className="mb-4 p-3 bg-gradient-to-r from-blue-50/70 to-slate-50 rounded-xl border border-blue-100/80 text-xs">
                      <div className="flex items-center justify-between text-[11px] mb-1.5">
                        <span className="font-bold text-blue-900 flex items-center gap-1">
                          <History size={12} className="text-blue-600" />
                          Cơ sở dữ liệu lịch sử tham chiếu:
                        </span>
                        <span className="font-mono font-bold text-blue-700 bg-blue-100/60 px-1.5 py-0.5 rounded">
                          {lastDgaRecord.date}
                        </span>
                      </div>
                      <div className="text-[11px] text-slate-600 line-clamp-1 mb-2">
                        Thiết bị: <strong className="text-slate-800">{equipmentCode}</strong> • {lastDgaRecord.source}
                      </div>

                      {/* Quick Sample Presets for Testing & Engineering Validation */}
                      <div className="flex items-center gap-1.5 flex-wrap pt-1.5 border-t border-blue-200/60">
                        <button
                          type="button"
                          onClick={handleLoadLastRecordedDga}
                          className="px-2 py-1 bg-white hover:bg-blue-50 text-blue-700 border border-blue-200 rounded text-[10px] font-bold shadow-2xs transition-colors flex items-center gap-1"
                          title="Nạp số liệu đo của lần trước để so sánh (Độ lệch = 0)"
                        >
                          <RotateCcw size={10} />
                          Nạp mẫu kỳ trước
                        </button>
                        <button
                          type="button"
                          onClick={handleLoadElevatedDgaSample}
                          className="px-2 py-1 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 rounded text-[10px] font-bold shadow-2xs transition-colors flex items-center gap-1"
                          title="Mẫu mô phỏng nồng độ khí tăng (Mũi tên đỏ ▲)"
                        >
                          <ArrowUpRight size={10} />
                          Mẫu tăng khí (▲)
                        </button>
                        <button
                          type="button"
                          onClick={handleLoadDecreasingDgaSample}
                          className="px-2 py-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 rounded text-[10px] font-bold shadow-2xs transition-colors flex items-center gap-1"
                          title="Mẫu mô phỏng nồng độ khí giảm (Mũi tên xanh ▼)"
                        >
                          <ArrowDownRight size={10} />
                          Mẫu giảm khí (▼)
                        </button>
                      </div>
                    </div>

                    {/* Trend Indicator Legend */}
                    <div className="flex items-center justify-between text-[10px] mb-3 px-1 text-slate-500 font-medium">
                      <span className="flex items-center gap-1 text-rose-600 font-bold">
                        ▲ Tăng (Xấu hơn)
                      </span>
                      <span className="flex items-center gap-1 text-emerald-600 font-bold">
                        ▼ Giảm (Tốt hơn)
                      </span>
                      <span className="flex items-center gap-1 text-slate-500 font-bold">
                        → Ổn định (±1%)
                      </span>
                    </div>

                    {/* Gas Parameters with Visual Trend Indicators */}
                    <div className="space-y-4">
                      <div>
                        <div className="text-xs font-bold text-slate-800 uppercase tracking-wider mb-2 flex items-center justify-between">
                          <span>Nồng độ khí hòa tan (ppm)</span>
                          <span className="text-[10px] font-mono text-slate-400 font-normal">IEEE C57.104</span>
                        </div>
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                          <DgaParamTrendInput
                            paramKey="h2"
                            label="H₂"
                            subLabel="Hydrogen"
                            unit="ppm"
                            standard="≤ 100"
                            value={dgaData.h2}
                            onChange={(val) => setDgaData({ ...dgaData, h2: val })}
                            historicalValue={lastDgaRecord.gases.h2}
                            historicalDate={lastDgaRecord.date}
                            isHigherBetter={false}
                            placeholder="Nhập H₂"
                          />
                          <DgaParamTrendInput
                            paramKey="ch4"
                            label="CH₄"
                            subLabel="Methane"
                            unit="ppm"
                            standard="≤ 120"
                            value={dgaData.ch4}
                            onChange={(val) => setDgaData({ ...dgaData, ch4: val })}
                            historicalValue={lastDgaRecord.gases.ch4}
                            historicalDate={lastDgaRecord.date}
                            isHigherBetter={false}
                            placeholder="Nhập CH₄"
                          />
                          <DgaParamTrendInput
                            paramKey="c2h6"
                            label="C₂H₆"
                            subLabel="Ethane"
                            unit="ppm"
                            standard="≤ 65"
                            value={dgaData.c2h6}
                            onChange={(val) => setDgaData({ ...dgaData, c2h6: val })}
                            historicalValue={lastDgaRecord.gases.c2h6}
                            historicalDate={lastDgaRecord.date}
                            isHigherBetter={false}
                            placeholder="Nhập C₂H₆"
                          />
                          <DgaParamTrendInput
                            paramKey="c2h4"
                            label="C₂H₄"
                            subLabel="Ethylene"
                            unit="ppm"
                            standard="≤ 50"
                            value={dgaData.c2h4}
                            onChange={(val) => setDgaData({ ...dgaData, c2h4: val })}
                            historicalValue={lastDgaRecord.gases.c2h4}
                            historicalDate={lastDgaRecord.date}
                            isHigherBetter={false}
                            placeholder="Nhập C₂H₄"
                          />
                          <DgaParamTrendInput
                            paramKey="c2h2"
                            label="C₂H₂"
                            subLabel="Acetylene"
                            unit="ppm"
                            standard="≤ 1.0"
                            value={dgaData.c2h2}
                            onChange={(val) => setDgaData({ ...dgaData, c2h2: val })}
                            historicalValue={lastDgaRecord.gases.c2h2}
                            historicalDate={lastDgaRecord.date}
                            isHigherBetter={false}
                            placeholder="Nhập C₂H₂"
                          />
                          <DgaParamTrendInput
                            paramKey="co"
                            label="CO"
                            subLabel="Carbon Monoxide"
                            unit="ppm"
                            standard="≤ 350"
                            value={dgaData.co}
                            onChange={(val) => setDgaData({ ...dgaData, co: val })}
                            historicalValue={lastDgaRecord.gases.co}
                            historicalDate={lastDgaRecord.date}
                            isHigherBetter={false}
                            placeholder="Nhập CO"
                          />
                          <DgaParamTrendInput
                            paramKey="co2"
                            label="CO₂"
                            subLabel="Carbon Dioxide"
                            unit="ppm"
                            standard="≤ 2500"
                            value={dgaData.co2}
                            onChange={(val) => setDgaData({ ...dgaData, co2: val })}
                            historicalValue={lastDgaRecord.gases.co2}
                            historicalDate={lastDgaRecord.date}
                            isHigherBetter={false}
                            placeholder="Nhập CO₂"
                          />
                          <DgaParamTrendInput
                            paramKey="o2"
                            label="O₂"
                            subLabel="Oxygen"
                            unit="ppm"
                            standard="≤ 3500"
                            value={dgaData.o2}
                            onChange={(val) => setDgaData({ ...dgaData, o2: val })}
                            historicalValue={lastDgaRecord.gases.o2}
                            historicalDate={lastDgaRecord.date}
                            isHigherBetter={false}
                            placeholder="Nhập O₂"
                          />
                          <DgaParamTrendInput
                            paramKey="n2"
                            label="N₂"
                            subLabel="Nitrogen"
                            unit="ppm"
                            standard="≤ 50000"
                            value={dgaData.n2}
                            onChange={(val) => setDgaData({ ...dgaData, n2: val })}
                            historicalValue={lastDgaRecord.gases.n2}
                            historicalDate={lastDgaRecord.date}
                            isHigherBetter={false}
                            placeholder="Nhập N₂"
                            className="sm:col-span-2"
                          />
                        </div>
                      </div>

                      {/* Other Oil & Transformer Parameters */}
                      <div className="border-t border-slate-200 pt-3.5 mt-2">
                        <div className="text-xs font-bold text-slate-800 uppercase tracking-wider mb-2 flex items-center justify-between">
                          <span>Thông số lý hóa dầu &amp; Máy biến áp</span>
                          <span className="text-[10px] font-mono text-slate-400 font-normal">ASTM / IEC</span>
                        </div>
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                          <DgaParamTrendInput
                            paramKey="moisture"
                            label="Moisture"
                            subLabel="Độ ẩm dầu"
                            unit="ppm"
                            standard="≤ 20"
                            value={dgaData.moisture}
                            onChange={(val) => setDgaData({ ...dgaData, moisture: val })}
                            historicalValue={lastDgaRecord.otherParams.moisture}
                            historicalDate={lastDgaRecord.date}
                            isHigherBetter={false}
                            placeholder="Nhập Moisture"
                          />
                          <DgaParamTrendInput
                            paramKey="bdStrength"
                            label="BD Strength"
                            subLabel="Đánh thủng"
                            unit="kV"
                            standard="≥ 50"
                            value={dgaData.bdStrength}
                            onChange={(val) => setDgaData({ ...dgaData, bdStrength: val })}
                            historicalValue={lastDgaRecord.otherParams.bdStrength}
                            historicalDate={lastDgaRecord.date}
                            isHigherBetter={true}
                            placeholder="Nhập BD Strength"
                          />
                          <DgaParamTrendInput
                            paramKey="acidity"
                            label="Acidity"
                            subLabel="Chỉ số axit"
                            unit="mgKOH/g"
                            standard="≤ 0.10"
                            value={dgaData.acidity}
                            onChange={(val) => setDgaData({ ...dgaData, acidity: val })}
                            historicalValue={lastDgaRecord.otherParams.acidity}
                            historicalDate={lastDgaRecord.date}
                            isHigherBetter={false}
                            placeholder="Nhập Acidity"
                          />
                          <DgaParamTrendInput
                            paramKey="ffa"
                            label="FFA"
                            subLabel="Furan 2-FAL"
                            unit="ppm"
                            standard="≤ 1.5"
                            value={dgaData.ffa}
                            onChange={(val) => setDgaData({ ...dgaData, ffa: val })}
                            historicalValue={lastDgaRecord.otherParams.ffa}
                            historicalDate={lastDgaRecord.date}
                            isHigherBetter={false}
                            placeholder="Nhập FFA"
                          />
                          <DgaParamTrendInput
                            paramKey="estDp"
                            label="Est DP"
                            subLabel="Độ trùng hợp"
                            unit="DP"
                            standard="≥ 500"
                            value={dgaData.estDp}
                            onChange={(val) => setDgaData({ ...dgaData, estDp: val })}
                            historicalValue={lastDgaRecord.otherParams.estDp}
                            historicalDate={lastDgaRecord.date}
                            isHigherBetter={true}
                            placeholder="Nhập DP"
                          />
                          <DgaParamTrendInput
                            paramKey="age"
                            label="Tuổi MBA"
                            subLabel="Năm vận hành"
                            unit="năm"
                            standard="≤ 40"
                            value={dgaData.age}
                            onChange={(val) => setDgaData({ ...dgaData, age: val })}
                            historicalValue={lastDgaRecord.otherParams.age}
                            historicalDate={lastDgaRecord.date}
                            isHigherBetter={false}
                            placeholder="Nhập tuổi"
                          />
                          <DgaParamTrendInput
                            paramKey="loadFactor"
                            label="% Phụ tải"
                            subLabel="Tỷ lệ tải"
                            unit="%"
                            standard="≤ 100"
                            value={dgaData.loadFactor}
                            onChange={(val) => setDgaData({ ...dgaData, loadFactor: val })}
                            historicalValue={lastDgaRecord.otherParams.loadFactor}
                            historicalDate={lastDgaRecord.date}
                            isHigherBetter={false}
                            placeholder="Nhập % tải"
                            className="sm:col-span-2"
                          />
                        </div>
                      </div>

                      <button 
                        onClick={analyzeDGA}
                        className="w-full mt-4 py-2.5 bg-slate-900 hover:bg-slate-800 active:bg-slate-950 text-white rounded-lg font-bold text-xs shadow-sm transition-all flex items-center justify-center gap-2 cursor-pointer"
                      >
                        <Activity size={15} />
                        <span>Phân tích dữ liệu &amp; Cập nhật Xu hướng</span>
                      </button>
                    </div>
                  </div>

                  {/* Analysis Results */}
                  <div className="lg:col-span-2 space-y-6">
                    {/* View Mode Switcher for DGA Analysis */}
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-white p-2.5 rounded-xl border border-slate-200 shadow-sm">
                      <div className="flex items-center gap-1.5 flex-wrap">
                        <button
                          type="button"
                          onClick={() => setDgaViewMode('dashboard')}
                          className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 ${
                            dgaViewMode === 'dashboard' ? 'bg-blue-600 text-white shadow-sm' : 'text-slate-600 hover:bg-slate-100'
                          }`}
                        >
                          <Activity size={13} />
                          DGA Advanced Dashboard
                        </button>
                        <button
                          type="button"
                          onClick={() => setDgaViewMode('both')}
                          className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                            dgaViewMode === 'both' ? 'bg-slate-900 text-white shadow-sm' : 'text-slate-600 hover:bg-slate-100'
                          }`}
                        >
                          Toàn bộ (Biểu đồ &amp; Chẩn đoán)
                        </button>
                        <button
                          type="button"
                          onClick={() => setDgaViewMode('trend')}
                          className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 ${
                            dgaViewMode === 'trend' ? 'bg-blue-600 text-white shadow-sm' : 'text-slate-600 hover:bg-slate-100'
                          }`}
                        >
                          <TrendingUp size={13} />
                          Biểu đồ Xu hướng (LineChart)
                        </button>
                        <button
                          type="button"
                          onClick={() => setDgaViewMode('diagnostic')}
                          className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 ${
                            dgaViewMode === 'diagnostic' ? 'bg-indigo-600 text-white shadow-sm' : 'text-slate-600 hover:bg-slate-100'
                          }`}
                        >
                          <Activity size={13} />
                          Chẩn đoán Duval &amp; Ma trận
                        </button>
                      </div>

                      <div className="text-xs text-slate-500 font-medium px-2">
                        Tiêu chuẩn: <span className="font-semibold text-slate-700">IEEE C57.104 &amp; IEC 60599</span>
                      </div>
                    </div>

                    {/* DGA Advanced Diagnostics Dashboard (Matching images) */}
                    {dgaViewMode === 'dashboard' && (
                      <DgaAdvancedDiagnosticsDashboard
                        currentFormValues={dgaData}
                        equipmentCode={equipmentCode}
                        equipmentName={equipmentName}
                        customerName={customerName}
                        analysisResult={dgaAnalysisResult}
                        lastRecordedValues={lastDgaRecord}
                        onAnalyze={analyzeDGA}
                        onExportPdf={exportToPDF}
                        onUpdateAssetInfo={(info) => {
                          if (info.equipmentCode) setEquipmentCode(info.equipmentCode);
                          if (info.equipmentName) setEquipmentName(info.equipmentName);
                        }}
                        onUpdateFormValues={(vals) => {
                          setDgaData(prev => ({ ...prev, ...vals }));
                        }}
                      />
                    )}

                    {/* Historical Trend LineChart */}
                    {(dgaViewMode === 'both' || dgaViewMode === 'trend') && (
                      <DgaHistoricalTrendChart
                        currentFormValues={dgaData}
                        equipmentCode={equipmentCode}
                        equipmentName={equipmentName}
                        onLoadSampleToForm={handleLoadDgaSampleToForm}
                      />
                    )}

                    {/* Detailed Diagnostic Results */}
                    {(dgaViewMode === 'both' || dgaViewMode === 'diagnostic') && (
                      dgaAnalysisResult ? (
                      <div id="dga-report-content" className="bg-white rounded-xl shadow-sm border border-slate-200 p-8">
                        <div className="border-b border-slate-200 pb-6 mb-6">
                          <div className="flex justify-between items-start mb-4">
                            <div>
                              <h2 className="text-2xl font-bold text-slate-900">Báo cáo Phân tích DGA</h2>
                              <p className="text-slate-500 mt-1">Mã thiết bị: {equipmentCode} | Khách hàng: {customerName}</p>
                            </div>
                            <div className="text-right">
                              <div className="text-sm font-medium text-slate-500">Ngày phân tích</div>
                              <div className="text-slate-900 font-semibold">{dgaAnalysisResult.timestamp}</div>
                            </div>
                          </div>
                        </div>

                        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
                          <div className="bg-slate-50 p-5 rounded-xl border border-slate-100">
                            <h4 className="text-sm font-bold text-slate-500 uppercase tracking-wider mb-2">Tình trạng hiện tại</h4>
                            <div className={`text-xl font-bold ${
                              dgaAnalysisResult.condition === 'Tốt' ? 'text-emerald-600' : 
                              dgaAnalysisResult.condition === 'Trung bình' ? 'text-amber-500' : 'text-rose-600'
                            }`}>
                              {dgaAnalysisResult.condition} (HI: {dgaAnalysisResult.currentHealthScore})
                            </div>
                          </div>
                          <div className="bg-slate-50 p-5 rounded-xl border border-slate-100">
                            <h4 className="text-sm font-bold text-slate-500 uppercase tracking-wider mb-2">Dự báo 10 năm</h4>
                            <div className="text-xl font-bold text-blue-600">
                              HI: {dgaAnalysisResult.futureHealthScore} ({dgaAnalysisResult.futureHI})
                            </div>
                          </div>
                        </div>

                        <div className="mb-8">
                          <h4 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
                            <Activity size={20} className="text-blue-500" />
                            Evaluate your DGA Matrix
                          </h4>
                          <div className="mb-4 p-4 bg-slate-50 rounded-lg border border-slate-200 text-sm">
                            <h5 className="font-bold text-slate-700 mb-2">Chú thích mã lỗi (Fault Codes Legend):</h5>
                            <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
                              <div><span className="font-semibold text-slate-800">ND:</span> Not Determined</div>
                              <div><span className="font-semibold text-emerald-600">OK:</span> Normal</div>
                              <div><span className="font-semibold text-blue-600">PD:</span> Partial Discharge</div>
                              <div><span className="font-semibold text-sky-500">S:</span> Stray Gassing</div>
                              <div><span className="font-semibold text-orange-400">T1:</span> Thermal &lt; 300°C</div>
                              <div><span className="font-semibold text-orange-500">T2:</span> Thermal 300-700°C</div>
                              <div><span className="font-semibold text-orange-600">T3:</span> Thermal &gt; 700°C</div>
                              <div><span className="font-semibold text-orange-300">O:</span> Overheating &lt; 250°C</div>
                              <div><span className="font-semibold text-orange-400">C:</span> Thermal with paper</div>
                              <div><span className="font-semibold text-amber-500">DT:</span> Thermal &amp; Electrical</div>
                              <div><span className="font-semibold text-red-800">D1:</span> Low Energy Discharge</div>
                              <div><span className="font-semibold text-red-600">D2:</span> High Energy Discharge</div>
                            </div>
                          </div>
                          <DgaMatrix matrix={dgaAnalysisResult.matrix} />
                        </div>

                        {dgaAnalysisResult.tdcgCondition !== 'Condition 1' ? (
                          <>
                            <div className="mb-8">
                              <h4 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
                                <Activity size={20} className="text-indigo-500" />
                                Duval Triangles
                              </h4>
                              <div className="flex flex-col gap-8 bg-slate-50 p-6 rounded-xl border border-slate-100">
                                <DuvalTriangle 
                                  title="Triangle 1" 
                                  labels={['%CH4', '%C2H4', '%C2H2']} 
                                  data={[parseFloat(dgaData.ch4)||0, parseFloat(dgaData.c2h4)||0, parseFloat(dgaData.c2h2)||0]} 
                                  type={1}
                                />
                                <DuvalTriangle 
                                  title="Triangle 4" 
                                  labels={['%H2', '%CH4', '%C2H6']} 
                                  data={[parseFloat(dgaData.h2)||0, parseFloat(dgaData.ch4)||0, parseFloat(dgaData.c2h6)||0]} 
                                  type={4}
                                />
                                <DuvalTriangle 
                                  title="Triangle 5" 
                                  labels={['%CH4', '%C2H4', '%C2H6']} 
                                  data={[parseFloat(dgaData.ch4)||0, parseFloat(dgaData.c2h4)||0, parseFloat(dgaData.c2h6)||0]} 
                                  type={5}
                                />
                              </div>
                            </div>

                            <div className="mb-8">
                              <h4 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
                                <Activity size={20} className="text-purple-500" />
                                Duval Pentagons
                              </h4>
                              <div className="flex flex-col gap-8 bg-slate-50 p-6 rounded-xl border border-slate-100">
                                <DuvalPentagon 
                                  title="Pentagon 1" 
                                  labels={['%H2', '%C2H6', '%CH4', '%C2H4', '%C2H2']} 
                                  data={[parseFloat(dgaData.h2)||0, parseFloat(dgaData.c2h6)||0, parseFloat(dgaData.ch4)||0, parseFloat(dgaData.c2h4)||0, parseFloat(dgaData.c2h2)||0]} 
                                  type={1}
                                />
                                <DuvalPentagon 
                                  title="Pentagon 2" 
                                  labels={['%H2', '%C2H6', '%CH4', '%C2H4', '%C2H2']} 
                                  data={[parseFloat(dgaData.h2)||0, parseFloat(dgaData.c2h6)||0, parseFloat(dgaData.ch4)||0, parseFloat(dgaData.c2h4)||0, parseFloat(dgaData.c2h2)||0]} 
                                  type={2}
                                />
                              </div>
                            </div>
                          </>
                        ) : (
                          <div className="mb-8 p-6 bg-emerald-50 rounded-xl border border-emerald-200 text-center">
                            <div className="flex justify-center mb-3">
                              <CheckCircle size={32} className="text-emerald-500" />
                            </div>
                            <h4 className="text-lg font-bold text-emerald-800 mb-2">Tình trạng bình thường (Condition 1)</h4>
                            <p className="text-emerald-600 max-w-2xl mx-auto">
                              Theo tiêu chuẩn IEEE C57.104-2008, do tổng lượng khí hòa tan (TDCG) đang ở mức Condition 1, máy biến áp đang hoạt động bình thường. Không cần thiết phải áp dụng các phương pháp phân tích chẩn đoán lỗi (như Duval Triangle hay Duval Pentagon).
                            </p>
                          </div>
                        )}

                        <div className="mb-8">
                          <h4 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
                            <Activity size={20} className="text-blue-500" />
                            Các chỉ số khác
                          </h4>
                          <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                            <div className="bg-slate-50 p-4 rounded-xl border border-slate-100 text-center">
                              <div className="text-sm font-medium text-slate-500 mb-1">KeyGas</div>
                              <div className="text-lg font-bold text-slate-800">{dgaAnalysisResult.matrix['KeyGas']}</div>
                            </div>
                            <div className="bg-slate-50 p-4 rounded-xl border border-slate-100 text-center">
                              <div className="text-sm font-medium text-slate-500 mb-1">ETRA</div>
                              <div className="text-lg font-bold text-slate-800">{dgaAnalysisResult.matrix['ETRA']}</div>
                            </div>
                            <div className="bg-slate-50 p-4 rounded-xl border border-slate-100 text-center">
                              <div className="text-sm font-medium text-slate-500 mb-1">CO2/CO</div>
                              <div className="text-lg font-bold text-slate-800">{dgaAnalysisResult.matrix['CO2/CO']}</div>
                            </div>
                            <div className="bg-slate-50 p-4 rounded-xl border border-slate-100 text-center">
                              <div className="text-sm font-medium text-slate-500 mb-1">IEC Ratio</div>
                              <div className="text-lg font-bold text-slate-800">{dgaAnalysisResult.matrix['IEC Ratio']}</div>
                            </div>
                          </div>
                          
                          <div className="grid grid-cols-2 md:grid-cols-3 gap-4 mt-4">
                            <div className={`p-4 rounded-xl border text-center ${dgaAnalysisResult.dpEstimation < 250 ? 'bg-red-50 border-red-200 text-red-800' : dgaAnalysisResult.dpEstimation < 400 ? 'bg-amber-50 border-amber-200 text-amber-800' : 'bg-emerald-50 border-emerald-200 text-emerald-800'}`}>
                              <div className="text-sm font-medium mb-1 opacity-80">DP Estimation</div>
                              <div className="text-lg font-bold">{dgaAnalysisResult.dpEstimation}</div>
                            </div>
                            <div className={`p-4 rounded-xl border text-center ${parseFloat(dgaAnalysisResult.currentHealthScore) >= 6.5 ? 'bg-red-50 border-red-200 text-red-800' : parseFloat(dgaAnalysisResult.currentHealthScore) >= 4 ? 'bg-amber-50 border-amber-200 text-amber-800' : 'bg-emerald-50 border-emerald-200 text-emerald-800'}`}>
                              <div className="text-sm font-medium mb-1 opacity-80">Current Health Index</div>
                              <div className="text-lg font-bold">{dgaAnalysisResult.currentHealthScore}</div>
                            </div>
                            <div className={`p-4 rounded-xl border text-center ${parseFloat(dgaAnalysisResult.currentPOF) > 0.05 ? 'bg-red-50 border-red-200 text-red-800' : parseFloat(dgaAnalysisResult.currentPOF) > 0.01 ? 'bg-amber-50 border-amber-200 text-amber-800' : 'bg-emerald-50 border-emerald-200 text-emerald-800'}`}>
                              <div className="text-sm font-medium mb-1 opacity-80">Current POF</div>
                              <div className="text-lg font-bold">{dgaAnalysisResult.currentPOF}</div>
                            </div>
                            <div className={`p-4 rounded-xl border text-center ${parseFloat(dgaAnalysisResult.futureHealthScore) >= 6.5 ? 'bg-red-50 border-red-200 text-red-800' : parseFloat(dgaAnalysisResult.futureHealthScore) >= 4 ? 'bg-amber-50 border-amber-200 text-amber-800' : 'bg-emerald-50 border-emerald-200 text-emerald-800'}`}>
                              <div className="text-sm font-medium mb-1 opacity-80">Health Index Year 10</div>
                              <div className="text-lg font-bold">{dgaAnalysisResult.futureHealthScore}</div>
                            </div>
                            <div className={`p-4 rounded-xl border text-center ${parseFloat(dgaAnalysisResult.futurePOF) > 0.05 ? 'bg-red-50 border-red-200 text-red-800' : parseFloat(dgaAnalysisResult.futurePOF) > 0.01 ? 'bg-amber-50 border-amber-200 text-amber-800' : 'bg-emerald-50 border-emerald-200 text-emerald-800'}`}>
                              <div className="text-sm font-medium mb-1 opacity-80">Year 10 POF</div>
                              <div className="text-lg font-bold">{dgaAnalysisResult.futurePOF}</div>
                            </div>
                            <div className={`p-4 rounded-xl border text-center ${dgaAnalysisResult.eolYears < 5 ? 'bg-red-50 border-red-200 text-red-800' : dgaAnalysisResult.eolYears < 10 ? 'bg-amber-50 border-amber-200 text-amber-800' : 'bg-emerald-50 border-emerald-200 text-emerald-800'}`}>
                              <div className="text-sm font-medium mb-1 opacity-80">Estimated EOL (Years)</div>
                              <div className="text-lg font-bold">{dgaAnalysisResult.eolYears}</div>
                            </div>
                          </div>
                        </div>

                        <div className="mb-8">
                          <h4 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
                            <Activity size={20} className="text-blue-500" />
                            Đánh giá TDCG (IEEE C57.104)
                          </h4>
                          <div className={`p-4 rounded-xl border ${dgaAnalysisResult.tdcgColor} flex flex-col gap-2`}>
                            <div className="flex justify-between items-center">
                              <span className="font-semibold">Tổng khí hòa tan (TDCG):</span>
                              <span className="text-xl font-bold">{dgaAnalysisResult.tdcg} ppm</span>
                            </div>
                            <div className="flex justify-between items-center">
                              <span className="font-semibold">Đánh giá:</span>
                              <span className="text-lg font-bold">{dgaAnalysisResult.tdcgCondition}</span>
                            </div>
                          </div>
                        </div>

                        <div className="mb-8">
                          <h4 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
                            <BarChart2 size={20} className="text-purple-500" />
                            Chỉ số sức khỏe (Health Score) - Hiện tại & 10 năm tới
                          </h4>
                          <div className="h-64">
                            <ResponsiveContainer width="100%" height="100%">
                              <BarChart data={dgaAnalysisResult.healthScoreData} margin={{ top: 20, right: 30, left: 0, bottom: 5 }}>
                                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                                <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fill: '#64748b', fontSize: 14, fontWeight: 500 }} />
                                <YAxis domain={[0, 10]} axisLine={false} tickLine={false} tick={{ fill: '#64748b', fontSize: 12 }} />
                                <Tooltip cursor={{fill: 'transparent'}} contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
                                <Bar dataKey="Health Score" radius={[6, 6, 0, 0]} maxBarSize={80}>
                                  {
                                    dgaAnalysisResult.healthScoreData.map((entry: any, index: number) => (
                                      <Cell key={`cell-${index}`} fill={entry.fill} />
                                    ))
                                  }
                                </Bar>
                              </BarChart>
                            </ResponsiveContainer>
                          </div>
                        </div>

                        <div>
                          <h4 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
                            <CheckCircle size={20} className={`text-${dgaAnalysisResult.recommendationColor}-500`} />
                            Khuyến cáo cho Khách hàng
                          </h4>
                          <ul className="space-y-2">
                            {dgaAnalysisResult.recommendations.map((rec: string, idx: number) => (
                              <li key={idx} className={`flex items-start gap-2 text-slate-700 bg-${dgaAnalysisResult.recommendationColor}-50/50 p-3 rounded-lg border border-${dgaAnalysisResult.recommendationColor}-100`}>
                                <span className={`text-${dgaAnalysisResult.recommendationColor}-500 mt-0.5`}>•</span>
                                <span>{rec}</span>
                              </li>
                            ))}
                          </ul>
                        </div>
                      </div>
                    ) : (
                      dgaViewMode === 'diagnostic' ? (
                        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-12 flex flex-col items-center justify-center text-center h-full min-h-[400px]">
                          <div className="w-16 h-16 bg-blue-50 text-blue-500 rounded-full flex items-center justify-center mb-4">
                            <Activity size={32} />
                          </div>
                          <h3 className="text-xl font-bold text-slate-800 mb-2">Chưa có kết quả chẩn đoán chi tiết</h3>
                          <p className="text-slate-500 max-w-md">
                            Vui lòng nhập các giá trị khí hòa tan (DGA) ở cột bên trái hoặc chọn <strong>"Nạp vào Form"</strong> từ bảng lịch sử đo trên, sau đó nhấn <strong>"Phân tích dữ liệu"</strong> để xem ma trận chẩn đoán và dự báo tuổi thọ.
                          </p>
                        </div>
                      ) : (
                        <div className="bg-gradient-to-r from-blue-50 to-indigo-50 rounded-xl p-5 border border-blue-100 flex flex-col sm:flex-row items-center justify-between gap-4">
                          <div className="flex items-center gap-3">
                            <div className="w-10 h-10 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center flex-shrink-0">
                              <Activity size={20} />
                            </div>
                            <div>
                              <h4 className="text-sm font-bold text-slate-800">Chưa tạo báo cáo ma trận chẩn đoán cho lần đo hiện tại</h4>
                              <p className="text-xs text-slate-600 mt-0.5">
                                Nhấn nút <strong>"Phân tích dữ liệu"</strong> ở cột bên trái hoặc chọn <strong>"Nạp vào Form"</strong> từ bảng lịch sử để chạy ma trận Duval Triangles, Pentagons và đánh giá Health Index.
                              </p>
                            </div>
                          </div>
                          <button
                            type="button"
                            onClick={analyzeDGA}
                            className="px-4 py-2 bg-slate-900 text-white text-xs font-bold rounded-lg hover:bg-slate-800 whitespace-nowrap shadow-sm"
                          >
                            Phân tích ngay
                          </button>
                        </div>
                      )
                    )
                  )}
                  </div>
                </div>
              )}

                {deepAnalysisSubTab === 'pv-cell' && (
                  <PvCellDeepThermalAnalyzer
                    onSaveReport={handleSaveNetaReport}
                    onNavigateToReports={() => setActiveTab('reports')}
                  />
                )}

                {deepAnalysisSubTab === 'wind-turbine' && (
                  <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                    <div className="lg:col-span-1 bg-white rounded-xl shadow-sm border border-slate-200 p-6">
                      <h3 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
                        <Wind size={20} className="text-blue-500" />
                        Dữ liệu cảm biến
                      </h3>
                      <div className="space-y-4">
                        <div>
                          <label className="block text-xs font-medium text-slate-700 mb-1">Độ rung cánh (mm/s)</label>
                          <input type="number" value={windBladeData.vibration} onChange={(e) => setWindBladeData({...windBladeData, vibration: e.target.value})} className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" />
                        </div>
                        <div>
                          <label className="block text-xs font-medium text-slate-700 mb-1">Phát xạ âm thanh (dB)</label>
                          <input type="number" value={windBladeData.acoustic} onChange={(e) => setWindBladeData({...windBladeData, acoustic: e.target.value})} className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" />
                        </div>
                        <div>
                          <label className="block text-xs font-medium text-slate-700 mb-1">Tốc độ quay (RPM)</label>
                          <input type="number" value={windBladeData.rotationSpeed} onChange={(e) => setWindBladeData({...windBladeData, rotationSpeed: e.target.value})} className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" />
                        </div>
                        <div>
                          <label className="block text-xs font-medium text-slate-700 mb-1">Tốc độ gió (m/s)</label>
                          <input type="number" value={windBladeData.windSpeed} onChange={(e) => setWindBladeData({...windBladeData, windSpeed: e.target.value})} className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" />
                        </div>
                        <div>
                          <label className="block text-xs font-medium text-slate-700 mb-1">Ghi chú kiểm tra trực quan</label>
                          <textarea value={windBladeData.visualNotes} onChange={(e) => setWindBladeData({...windBladeData, visualNotes: e.target.value})} className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none h-24" placeholder="Nhập các quan sát về vết nứt, xói mòn..." />
                        </div>
                        <button 
                          onClick={analyzeWindBlade}
                          className="w-full mt-6 py-2.5 bg-slate-900 text-white rounded-lg font-medium hover:bg-slate-800 transition-colors shadow-sm"
                        >
                          Phân tích cánh quạt
                        </button>
                      </div>
                    </div>

                    <div className="lg:col-span-2 space-y-6">
                      {windAnalysisResult ? (
                        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-8">
                          <div className="flex justify-between items-start mb-8 border-b border-slate-100 pb-6">
                            <div>
                              <h3 className="text-xl font-bold text-slate-900">Kết quả phân tích Cánh quạt</h3>
                              <p className="text-slate-500 mt-1">Thời gian: {windAnalysisResult.timestamp}</p>
                            </div>
                            <div className={`px-4 py-2 rounded-full font-bold text-sm ${windAnalysisResult.bgColor} ${windAnalysisResult.color}`}>
                              {windAnalysisResult.condition}
                            </div>
                          </div>

                          <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
                            <div className="p-6 bg-slate-50 rounded-xl border border-slate-100">
                              <h4 className="text-sm font-bold text-slate-500 uppercase tracking-wider mb-4">Mức độ rung động</h4>
                              <div className="flex items-end gap-2">
                                <span className={`text-4xl font-bold ${parseFloat(windBladeData.vibration) > 0.8 ? 'text-amber-600' : 'text-slate-900'}`}>{windBladeData.vibration}</span>
                                <span className="text-slate-500 mb-1">mm/s RMS</span>
                              </div>
                            </div>
                            <div className="p-6 bg-slate-50 rounded-xl border border-slate-100">
                              <h4 className="text-sm font-bold text-slate-500 uppercase tracking-wider mb-4">Phát xạ âm thanh</h4>
                              <div className="flex items-end gap-2">
                                <span className={`text-4xl font-bold ${parseFloat(windBladeData.acoustic) > 40 ? 'text-amber-600' : 'text-slate-900'}`}>{windBladeData.acoustic}</span>
                                <span className="text-slate-500 mb-1">dB</span>
                              </div>
                            </div>
                          </div>

                          <div>
                            <h4 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
                              <AlertCircle size={20} className="text-blue-500" />
                              Kết luận & Hành động
                            </h4>
                            <ul className="space-y-3">
                              {windAnalysisResult.recommendations.map((rec: string, idx: number) => (
                                <li key={idx} className="flex items-start gap-3 p-4 bg-slate-50 rounded-lg border border-slate-100">
                                  <div className={`mt-1 w-2 h-2 rounded-full ${windAnalysisResult.bgColor.replace('bg-', 'bg-').replace('50', '500')}`} />
                                  <span className="text-slate-700">{rec}</span>
                                </li>
                              ))}
                            </ul>
                          </div>
                        </div>
                      ) : (
                        <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-12 flex flex-col items-center justify-center text-center h-full min-h-[400px]">
                          <Wind size={48} className="text-slate-200 mb-4" />
                          <h3 className="text-xl font-bold text-slate-800 mb-2">Sẵn sàng phân tích Cánh quạt</h3>
                          <p className="text-slate-500 max-w-md">
                            Nhập dữ liệu từ cảm biến rung động và âm thanh để phát hiện các hư hỏng cấu trúc tiềm ẩn trên cánh quạt.
                          </p>
                        </div>
                      )}
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}
          {/* --- CMMS VIEW --- */}
          {activeTab === 'cmms' && (
            <div className="space-y-6">
              {/* CMMS Mode / Sub-navigation Tabs */}
              <div className="bg-white rounded-2xl p-2 border border-slate-200 shadow-sm flex flex-col sm:flex-row gap-2">
                <button
                  type="button"
                  onClick={() => setCmmsSubTab('station')}
                  className={`flex-1 flex items-center justify-center gap-2.5 py-3 px-4 rounded-xl font-bold text-sm transition-all ${
                    cmmsSubTab === 'station'
                      ? 'bg-blue-600 text-white shadow-md shadow-blue-500/25 ring-2 ring-blue-600/20'
                      : 'text-slate-600 hover:bg-slate-100 hover:text-slate-900'
                  }`}
                >
                  <FileSpreadsheet size={18} className={cmmsSubTab === 'station' ? 'text-blue-200' : 'text-slate-400'} />
                  <span>Trạm Nhập Liệu & Đồng Bộ 2 Chiều Google Sheet</span>
                  <span className={`text-[10px] uppercase tracking-wider px-2 py-0.5 rounded-full font-bold ${
                    cmmsSubTab === 'station' ? 'bg-blue-500 text-white' : 'bg-slate-200 text-slate-700'
                  }`}>
                    TEV Flatform
                  </span>
                </button>

                <button
                  type="button"
                  onClick={() => setCmmsSubTab('table')}
                  className={`flex-1 flex items-center justify-center gap-2.5 py-3 px-4 rounded-xl font-bold text-sm transition-all ${
                    cmmsSubTab === 'table'
                      ? 'bg-blue-600 text-white shadow-md shadow-blue-500/25 ring-2 ring-blue-600/20'
                      : 'text-slate-600 hover:bg-slate-100 hover:text-slate-900'
                  }`}
                >
                  <ClipboardList size={18} className={cmmsSubTab === 'table' ? 'text-blue-200' : 'text-slate-400'} />
                  <span>Danh Sách & Bộ Lọc Phiếu Công Việc</span>
                  <span className={`text-[10px] uppercase tracking-wider px-2 py-0.5 rounded-full font-bold ${
                    cmmsSubTab === 'table' ? 'bg-blue-500 text-white' : 'bg-slate-200 text-slate-700'
                  }`}>
                    {workOrders.length} WO
                  </span>
                </button>
              </div>

              {/* View 1: Trạm Nhập Liệu & Đồng Bộ 2 Chiều Google Sheet */}
              {cmmsSubTab === 'station' && (
                <FseFieldDataEntryStation
                  customers={customers}
                  allEquipment={allEquipment}
                  workOrders={workOrders}
                  isGoogleConnected={isGoogleConnected}
                  onConnectGoogle={handleConnectGoogle}
                  onSaveCustomer={handleSaveCustomerFromField}
                  onSaveEquipment={handleSaveEquipmentFromField}
                  onSaveWorkOrder={handleSaveWorkOrderFromField}
                  onSyncAllToSheets={handleSyncToSheets}
                  onFetchFromSheets={handleFetchFromSheets}
                  onSyncAppToSheetsAndFirestore={handleSyncAppToSheetsAndFirestore}
                  onSyncSheetsToFirestoreAndApp={handleSyncSheetsToFirestoreAndApp}
                  isSyncing={isSyncing}
                  userEmail={auth.currentUser?.email || 'sgm1707@gmail.com'}
                />
              )}

              {/* View 2: Bảng Danh Sách & Thống Kê Phiếu CMMS */}
              {cmmsSubTab === 'table' && (
                <div className="space-y-6">
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                    <div>
                      <h2 className="text-2xl font-bold text-slate-900">Quản lý CMMS (Solar & Wind)</h2>
                      <p className="text-slate-500">Quản lý phiếu công việc, bảo trì định kỳ và sửa chữa.</p>
                    </div>
                <div className="flex items-center gap-3">
                  <button 
                    onClick={() => {
                      setNewWorkOrder({
                        title: '',
                        description: '',
                        equipmentId: [],
                        customerId: '',
                        priority: 'medium',
                        type: 'preventive',
                        status: 'initiated',
                        assignedTo: '',
                        dueDate: '',
                        usedMaterials: [],
                        responsibleApprove: 'TM',
                        responsibleDo: 'Everybody',
                        blockingRequired: false,
                        workPermitId: generateWorkPermitId(),
                        isUnplanned: false,
                        failureCode: '',
                        rootCause: '',
                        actualTimeSpent: 0,
                        pmFrequency: '',
                        estimatedTime: 0,
                        laborCount: 0,
                        laborCost: 0,
                        partCost: 0,
                        attachments: [],
                        downtimeStart: '',
                        repairStart: '',
                        repairEnd: '',
                        restartTime: ''
                      });
                      setSelectedWorkOrder(null);
                      setShowWorkOrderModal(true);
                    }}
                    className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg font-medium hover:bg-blue-700 transition-colors shadow-sm"
                  >
                    <Plus size={18} />
                    Tạo phiếu mới
                  </button>
                </div>
              </div>

              {/* Stats */}
              <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-xs font-bold text-slate-500 uppercase">Tổng số phiếu</span>
                    <ClipboardList className="text-blue-500" size={16} />
                  </div>
                  <div className="text-2xl font-bold text-slate-900">{workOrders.length}</div>
                </div>
                <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-xs font-bold text-slate-500 uppercase">Đang thực hiện</span>
                    <Activity className="text-amber-500" size={16} />
                  </div>
                  <div className="text-2xl font-bold text-slate-900">
                    {workOrders.filter(o => o.status === 'in-progress').length}
                  </div>
                </div>
                <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-xs font-bold text-slate-500 uppercase">Đã hoàn thành</span>
                    <CheckCircle className="text-emerald-500" size={16} />
                  </div>
                  <div className="text-2xl font-bold text-slate-900">
                    {workOrders.filter(o => o.status === 'completed').length}
                  </div>
                </div>
                <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-xs font-bold text-slate-500 uppercase">Quá hạn</span>
                    <AlertCircle className="text-rose-600" size={16} />
                  </div>
                  <div className="text-2xl font-bold text-rose-600">
                    {workOrders.filter(o => o.status === 'overdue' || (o.dueDate && new Date(o.dueDate) < new Date() && o.status !== 'completed')).length}
                  </div>
                </div>
              </div>

              {/* Work Orders Table */}
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="p-4 border-b border-slate-100 bg-slate-50 flex flex-col gap-4">
                  <div className="flex items-center justify-between">
                    <h3 className="font-bold text-slate-800">Danh sách Phiếu công việc</h3>
                    <div className="flex items-center gap-2">
                      <div className="relative">
                        <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                        <input 
                          type="text" 
                          placeholder="Tìm kiếm phiếu..." 
                          value={woSearchQuery}
                          onChange={(e) => setWoSearchQuery(e.target.value)}
                          className="pl-9 pr-4 py-1.5 bg-white border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 w-64"
                        />
                      </div>
                    </div>
                  </div>
                  
                  {/* Filters */}
                  <div className="flex flex-wrap items-center gap-3">
                    <select 
                      value={woTypeFilter} 
                      onChange={(e) => setWoTypeFilter(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      <option value="">Tất cả Loại</option>
                      <option value="preventive">Định kỳ</option>
                      <option value="corrective">Sửa chữa</option>
                      <option value="inspection">Kiểm tra</option>
                      <option value="emergency">Khẩn cấp</option>
                    </select>
                    <select 
                      value={woPriorityFilter} 
                      onChange={(e) => setWoPriorityFilter(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      <option value="">Tất cả Ưu tiên</option>
                      <option value="low">Thấp</option>
                      <option value="medium">Trung bình</option>
                      <option value="high">Cao</option>
                      <option value="urgent">Khẩn cấp</option>
                    </select>
                    <select 
                      value={woStatusFilter} 
                      onChange={(e) => setWoStatusFilter(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      <option value="">Tất cả Trạng thái</option>
                      <option value="initiated">Khởi tạo</option>
                      <option value="approved">Đã duyệt</option>
                      <option value="planned">Đã lên kế hoạch</option>
                      <option value="scheduled">Đã lên lịch</option>
                      <option value="in-progress">Đang thực hiện</option>
                      <option value="completed">Hoàn thành</option>
                      <option value="overdue">Quá hạn</option>
                      <option value="cancelled">Đã hủy</option>
                    </select>
                    <select 
                      value={woAssigneeFilter} 
                      onChange={(e) => setWoAssigneeFilter(e.target.value)}
                      className="px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      <option value="">Tất cả Người thực hiện</option>
                      {Array.from(new Set(workOrders.map(wo => wo.assignedTo).filter(Boolean))).map(assignee => (
                        <option key={assignee} value={assignee}>{assignee}</option>
                      ))}
                    </select>
                    {(woTypeFilter || woPriorityFilter || woStatusFilter || woAssigneeFilter || woSearchQuery) && (
                      <button 
                        onClick={() => {
                          setWoTypeFilter('');
                          setWoPriorityFilter('');
                          setWoStatusFilter('');
                          setWoAssigneeFilter('');
                          setWoSearchQuery('');
                        }}
                        className="text-xs text-blue-600 hover:text-blue-800 font-medium"
                      >
                        Xóa bộ lọc
                      </button>
                    )}
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="border-b border-slate-100 text-xs uppercase tracking-wider text-slate-500 font-semibold bg-slate-50/50">
                        <th className="p-4">Mã WP</th>
                        <th className="p-4">Tiêu đề / Thiết bị</th>
                        <th className="p-4">Loại</th>
                        <th className="p-4">Ưu tiên</th>
                        <th className="p-4">Trạng thái</th>
                        <th className="p-4">Người thực hiện</th>
                        <th className="p-4">Ngày tạo</th>
                        <th className="p-4 text-right">Thao tác</th>
                      </tr>
                    </thead>
                    <tbody className="text-sm divide-y divide-slate-50">
                      {isWorkOrdersLoading ? (
                        <tr>
                          <td colSpan={8} className="p-8 text-center text-slate-500">Đang tải dữ liệu...</td>
                        </tr>
                      ) : (() => {
                        const filteredWorkOrders = workOrders.filter(order => {
                          if (woTypeFilter && order.type !== woTypeFilter) return false;
                          if (woPriorityFilter && order.priority !== woPriorityFilter) return false;
                          if (woStatusFilter && order.status !== woStatusFilter) return false;
                          if (woAssigneeFilter && order.assignedTo !== woAssigneeFilter) return false;
                          if (woSearchQuery) {
                            const q = woSearchQuery.toLowerCase();
                            return (
                              (order.title && order.title.toLowerCase().includes(q)) ||
                              (Array.isArray(order.equipmentId) ? order.equipmentId.some(id => id.toLowerCase().includes(q)) : (order.equipmentId && order.equipmentId.toLowerCase().includes(q))) ||
                              (order.workPermitId && order.workPermitId.toLowerCase().includes(q))
                            );
                          }
                          return true;
                        });

                        if (filteredWorkOrders.length === 0) {
                          return (
                            <tr>
                              <td colSpan={8} className="p-8 text-center text-slate-500">Không tìm thấy phiếu công việc nào phù hợp.</td>
                            </tr>
                          );
                        }

                        return filteredWorkOrders.map((order) => (
                          <tr key={order.id} className="hover:bg-slate-50 transition-colors">
                            <td className="p-4 font-mono text-xs font-bold text-blue-600 bg-blue-50/50 rounded-lg">
                              {order.workPermitId || 'N/A'}
                            </td>
                            <td className="p-4">
                              <div className="font-bold text-slate-900">{order.title}</div>
                              <div className="text-xs text-slate-500 font-mono">
                                {Array.isArray(order.equipmentId) ? order.equipmentId.join(', ') : order.equipmentId}
                              </div>
                            </td>
                            <td className="p-4">
                              <span className="px-2 py-0.5 rounded-full text-[10px] font-bold uppercase bg-slate-100 text-slate-600 border border-slate-200">
                                {order.type === 'preventive' ? 'Định kỳ' : order.type === 'corrective' ? 'Sửa chữa' : order.type === 'inspection' ? 'Kiểm tra' : 'Khẩn cấp'}
                              </span>
                            </td>
                            <td className="p-4">
                              <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold uppercase ${
                                order.priority === 'urgent' ? 'bg-rose-100 text-rose-700 border border-rose-200' :
                                order.priority === 'high' ? 'bg-orange-100 text-orange-700 border border-orange-200' :
                                order.priority === 'medium' ? 'bg-blue-100 text-blue-700 border border-blue-200' :
                                'bg-slate-100 text-slate-600 border border-slate-200'
                              }`}>
                                {order.priority}
                              </span>
                            </td>
                            <td className="p-4">
                              <div className="flex items-center gap-1.5">
                                <span className={`w-2 h-2 rounded-full ${
                                  order.status === 'completed' ? 'bg-emerald-500' :
                                  order.status === 'in-progress' ? 'bg-amber-500' :
                                  order.status === 'overdue' ? 'bg-rose-600' :
                                  order.status === 'cancelled' ? 'bg-slate-400' : 'bg-blue-500'
                                }`}></span>
                                <span className={`font-medium capitalize ${
                                  order.status === 'overdue' ? 'text-rose-600 font-bold' : 'text-slate-700'
                                }`}>
                                  {order.status === 'overdue' ? 'Quá hạn' : 
                                   order.status === 'pending' ? 'Chờ xử lý' :
                                   order.status === 'in-progress' ? 'Đang làm' :
                                   order.status === 'completed' ? 'Hoàn thành' :
                                   order.status === 'initiated' ? 'Khởi tạo' :
                                   order.status === 'approved' ? 'Đã duyệt' :
                                   order.status === 'planned' ? 'Đã lên KH' :
                                   order.status === 'scheduled' ? 'Đã lên lịch' : 'Đã hủy'}
                                </span>
                              </div>
                            </td>
                            <td className="p-4 text-slate-600">{order.assignedTo || 'Chưa phân công'}</td>
                            <td className="p-4 text-slate-500 text-xs">
                              <div>Tạo: {order.createdAt ? new Date(order.createdAt).toLocaleDateString('vi-VN') : 'N/A'}</div>
                              {order.dueDate && (
                                <div className={`mt-1 font-medium ${new Date(order.dueDate) < new Date() && order.status !== 'completed' ? 'text-rose-500' : 'text-slate-400'}`}>
                                  Hạn: {new Date(order.dueDate).toLocaleDateString('vi-VN')}
                                </div>
                              )}
                            </td>
                            <td className="p-4 text-right">
                              <div className="flex items-center justify-end gap-2">
                                <button 
                                  onClick={() => generateWorkOrderPDF(order)}
                                  className="p-2 text-slate-500 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                                  title="Xuất PDF"
                                >
                                  <Download size={16} />
                                </button>
                                <button 
                                  onClick={() => handleDeleteWorkOrder(order.id)}
                                  className="p-2 text-slate-500 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-colors"
                                  title="Xóa"
                                >
                                  <Trash2 size={16} />
                                </button>
                                <button 
                                  onClick={() => {
                                    setSelectedWorkOrder(order);
                                    setNewWorkOrder({
                                      title: order.title,
                                      description: order.description || '',
                                      equipmentId: Array.isArray(order.equipmentId) ? order.equipmentId : (order.equipmentId ? [order.equipmentId] : []),
                                      customerId: order.customerId,
                                      priority: order.priority,
                                      type: order.type,
                                      status: order.status,
                                      assignedTo: order.assignedTo || '',
                                      dueDate: order.dueDate || '',
                                      usedMaterials: order.usedMaterials || [],
                                      responsibleApprove: order.responsibleApprove || 'TM',
                                      responsibleDo: order.responsibleDo || 'Everybody',
                                      blockingRequired: order.blockingRequired || false,
                                      workPermitId: order.workPermitId || '',
                                      isUnplanned: order.isUnplanned || false,
                                      failureCode: order.failureCode || '',
                                      rootCause: order.rootCause || '',
                                      actualTimeSpent: order.actualTimeSpent || 0,
                                      pmFrequency: order.pmFrequency || '',
                                      estimatedTime: order.estimatedTime || 0,
                                      laborCount: order.laborCount || 0,
                                      laborCost: order.laborCost || 0,
                                      partCost: order.partCost || 0,
                                      attachments: order.attachments || [],
                                      downtimeStart: order.downtimeStart || '',
                                      repairStart: order.repairStart || '',
                                      repairEnd: order.repairEnd || '',
                                      restartTime: order.restartTime || ''
                                    });
                                    setShowWorkOrderModal(true);
                                  }}
                                  className="p-2 text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                                  title="Chỉnh sửa"
                                >
                                  <Settings size={16} />
                                </button>
                              </div>
                            </td>
                          </tr>
                        ));
                      })()}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}
        </div>
      )}

          {/* --- PM SCHEDULE VIEW --- */}
          {activeTab === 'pm-schedule' && (
            <div className="space-y-6">
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div>
                  <h2 className="text-2xl font-bold text-slate-900">Lịch bảo trì định kỳ (PM) & Giám sát Tiến độ</h2>
                  <p className="text-slate-500">Bảng điều khiển theo dõi phiếu PM đã hoàn thành, quá hạn và sắp đến hạn trong tháng, kết hợp tự động cảnh báo Email.</p>
                </div>
                <div className="flex items-center gap-3">
                  <button
                    onClick={() => {
                      setNewWorkOrder({
                        title: '',
                        description: '',
                        equipmentId: [],
                        customerId: '',
                        factory: '',
                        priority: 'medium',
                        type: 'preventive',
                        status: 'initiated',
                        assignedTo: '',
                        dueDate: new Date().toISOString().split('T')[0],
                        usedMaterials: [],
                        responsibleApprove: 'TM',
                        responsibleDo: 'Everybody',
                        blockingRequired: false,
                        workPermitId: generateWorkPermitId(),
                        isUnplanned: false,
                        failureCode: '',
                        pmFrequency: 'monthly',
                        estimatedTime: 0,
                        laborCount: 0,
                        laborCost: 0,
                        partCost: 0,
                        attachments: [],
                        downtimeStart: '',
                        repairStart: '',
                        repairEnd: '',
                        restartTime: ''
                      });
                      setSelectedWorkOrder(null);
                      setShowWorkOrderModal(true);
                    }}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-sm font-semibold shadow-xs transition-all"
                  >
                    <Plus size={16} />
                    <span>Tạo Phiếu PM Mới</span>
                  </button>

                  <button
                    onClick={() => setShowPMNotificationModal(true)}
                    className="inline-flex items-center gap-2 px-4 py-2 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white rounded-xl text-sm font-semibold shadow-xs transition-all"
                  >
                    <Mail size={16} />
                    <span>Email Cảnh Báo (3 Ngày)</span>
                    {pmUpcomingTasks.length > 0 && (
                      <span className="ml-1 px-2 py-0.5 bg-amber-400 text-slate-950 text-xs font-bold rounded-full">
                        {pmUpcomingTasks.length}
                      </span>
                    )}
                  </button>
                </div>
              </div>

              {/* Sub-navigation Tabs */}
              <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-200 pb-3">
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    onClick={() => setPmSubTab('dashboard')}
                    className={`flex items-center gap-2 px-4 py-2 rounded-xl font-bold text-sm transition-all ${
                      pmSubTab === 'dashboard'
                        ? 'bg-blue-600 text-white shadow-sm shadow-blue-500/20'
                        : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-200'
                    }`}
                  >
                    <BarChart3 size={16} />
                    <span>Biểu Đồ & Dashboard Tiến Độ PM</span>
                    <span className={`px-1.5 py-0.2 rounded-full text-[10px] font-black uppercase tracking-wider ${
                      pmSubTab === 'dashboard' ? 'bg-blue-500 text-white' : 'bg-blue-100 text-blue-700'
                    }`}>
                      Mới
                    </span>
                  </button>

                  <button
                    onClick={() => setPmSubTab('schedule')}
                    className={`flex items-center gap-2 px-4 py-2 rounded-xl font-bold text-sm transition-all ${
                      pmSubTab === 'schedule'
                        ? 'bg-blue-600 text-white shadow-sm shadow-blue-500/20'
                        : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-200'
                    }`}
                  >
                    <Calendar size={16} />
                    <span>Danh Mục Lịch Thiết Bị PM ({allEquipment.length})</span>
                  </button>

                  <button
                    onClick={() => setPmSubTab('alerts')}
                    className={`flex items-center gap-2 px-4 py-2 rounded-xl font-bold text-sm transition-all ${
                      pmSubTab === 'alerts'
                        ? 'bg-blue-600 text-white shadow-sm shadow-blue-500/20'
                        : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-200'
                    }`}
                  >
                    <Mail size={16} />
                    <span>Cảnh Báo Tự Động (Cloud Functions)</span>
                    {pmUpcomingTasks.length > 0 && (
                      <span className="px-1.5 py-0.2 rounded-full text-[10px] font-bold bg-amber-400 text-slate-950">
                        {pmUpcomingTasks.length}
                      </span>
                    )}
                  </button>
                </div>
              </div>

              {/* Sub-tab 1: PM Dashboard & Charts */}
              {pmSubTab === 'dashboard' && (
                <PmScheduleDashboard
                  workOrders={workOrders}
                  allEquipment={allEquipment}
                  onOpenWorkOrderModal={(wo) => {
                    setSelectedWorkOrder(wo);
                    setShowWorkOrderModal(true);
                  }}
                  onSendAlert={(id) => handleSendPMAlerts(true)}
                  onOpenEmailConfig={() => setShowPMNotificationModal(true)}
                  onUpdateWorkOrderStatus={handleQuickUpdateWorkOrderStatus}
                />
              )}

              {/* Sub-tab 2: Equipment Schedule Table */}
              {pmSubTab === 'schedule' && (
                <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="border-b border-slate-100 text-xs uppercase tracking-wider text-slate-500 font-semibold bg-slate-50/50">
                          <th className="p-4">Thiết bị</th>
                          <th className="p-4">Tần suất PM</th>
                          <th className="p-4">Lần bảo trì cuối</th>
                          <th className="p-4">Lần bảo trì tiếp theo</th>
                          <th className="p-4">Trạng thái</th>
                          <th className="p-4 text-right">Thao tác</th>
                        </tr>
                      </thead>
                      <tbody className="text-sm divide-y divide-slate-50">
                        {allEquipment.map(eq => {
                          // Find the latest completed PM work order for this equipment
                          const eqPmOrders = workOrders.filter(wo => (Array.isArray(wo.equipmentId) ? wo.equipmentId.includes(eq.id) : wo.equipmentId === eq.id) && wo.type === 'preventive' && wo.status === 'completed');
                          eqPmOrders.sort((a, b) => new Date(b.updatedAt || b.createdAt).getTime() - new Date(a.updatedAt || a.createdAt).getTime());
                          const latestPm = eqPmOrders[0];
                          
                          // If no PM order exists, we might not know the frequency unless it's stored on the equipment.
                          // For now, we'll look for ANY PM order (even not completed) to get the frequency.
                          const anyPmOrder = workOrders.find(wo => (Array.isArray(wo.equipmentId) ? wo.equipmentId.includes(eq.id) : wo.equipmentId === eq.id) && wo.type === 'preventive' && wo.pmFrequency);
                          const frequency = anyPmOrder?.pmFrequency || '';
                          
                          if (!frequency) return null; // Skip equipment without PM schedule

                          const lastDateStr = latestPm ? (latestPm.updatedAt || latestPm.createdAt) : null;
                          const nextDate = lastDateStr ? calculateNextPMDate(lastDateStr, frequency) : null;
                          const isOverdue = nextDate ? nextDate < new Date() : false;
                          const isDueSoon = nextDate ? (nextDate.getTime() - new Date().getTime()) / (1000 * 3600 * 24) <= 7 : false;

                          return (
                            <tr key={eq.id} className="hover:bg-slate-50 transition-colors">
                              <td className="p-4">
                                <div className="font-bold text-slate-900">{eq.name}</div>
                                <div className="text-xs text-slate-500 font-mono">{eq.id}</div>
                              </td>
                              <td className="p-4 text-slate-600">
                                {frequency === 'daily' ? 'Hàng ngày' :
                                 frequency === 'weekly' ? 'Hàng tuần' :
                                 frequency === 'bi-monthly' ? 'Nửa tháng' :
                                 frequency === 'monthly' ? 'Hàng tháng' :
                                 frequency === '3-months' ? '3 tháng' :
                                 frequency === '6-months' ? '6 tháng' :
                                 frequency === '1-year' ? '1 năm' :
                                 frequency === '3-years' ? '3 năm' :
                                 frequency === '6-years' ? '6 năm' : frequency}
                              </td>
                              <td className="p-4 text-slate-600">
                                {lastDateStr ? new Date(lastDateStr).toLocaleDateString('vi-VN') : 'Chưa có dữ liệu'}
                              </td>
                              <td className="p-4 font-medium">
                                {nextDate ? (
                                  <span className={isOverdue ? 'text-rose-600 font-bold' : isDueSoon ? 'text-amber-600 font-bold' : 'text-slate-900'}>
                                    {nextDate.toLocaleDateString('vi-VN')}
                                  </span>
                                ) : 'Chưa xác định'}
                              </td>
                              <td className="p-4">
                                {isOverdue ? (
                                  <span className="px-2 py-0.5 rounded-full text-[10px] font-bold uppercase bg-rose-100 text-rose-700 border border-rose-200">Quá hạn</span>
                                ) : isDueSoon ? (
                                  <span className="px-2 py-0.5 rounded-full text-[10px] font-bold uppercase bg-amber-100 text-amber-700 border border-amber-200">Sắp đến hạn</span>
                                ) : (
                                  <span className="px-2 py-0.5 rounded-full text-[10px] font-bold uppercase bg-emerald-100 text-emerald-700 border border-emerald-200">Bình thường</span>
                                )}
                              </td>
                              <td className="p-4 text-right">
                                <button 
                                  onClick={() => {
                                    setNewWorkOrder({
                                      title: `Bảo trì định kỳ - ${eq.name}`,
                                      description: '',
                                      equipmentId: [eq.id],
                                      customerId: eq.customer || '',
                                      priority: 'medium',
                                      type: 'preventive',
                                      status: 'initiated',
                                      assignedTo: '',
                                      dueDate: nextDate ? nextDate.toISOString().split('T')[0] : '',
                                      usedMaterials: [],
                                      responsibleApprove: 'TM',
                                      responsibleDo: 'Everybody',
                                      blockingRequired: false,
                                      workPermitId: generateWorkPermitId(),
                                      isUnplanned: false,
                                      failureCode: '',
                                      pmFrequency: frequency,
                                      estimatedTime: 0,
                                      laborCount: 0,
                                      laborCost: 0,
                                      partCost: 0,
                                      attachments: [],
                                      downtimeStart: '',
                                      repairStart: '',
                                      repairEnd: '',
                                      restartTime: ''
                                    });
                                    setSelectedWorkOrder(null);
                                    setShowWorkOrderModal(true);
                                  }}
                                  className="px-3 py-1.5 bg-blue-50 text-blue-600 hover:bg-blue-100 rounded-lg text-xs font-medium transition-colors"
                                >
                                  Tạo phiếu PM
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                        {allEquipment.filter(eq => workOrders.some(wo => (Array.isArray(wo.equipmentId) ? wo.equipmentId.includes(eq.id) : wo.equipmentId === eq.id) && wo.type === 'preventive' && wo.pmFrequency)).length === 0 && (
                          <tr>
                            <td colSpan={6} className="p-8 text-center text-slate-500">
                              Chưa có thiết bị nào được thiết lập lịch bảo trì định kỳ (PM).<br/>
                              Hãy tạo một phiếu công việc loại "Bảo trì định kỳ" và chọn tần suất để thiết lập.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* Sub-tab 3: Automated Alert Center */}
              {pmSubTab === 'alerts' && (
                <div className="space-y-6">
                  {/* AUTOMATED EMAIL NOTIFICATION HIGHLIGHT BANNER */}
                  <div className="bg-gradient-to-br from-slate-900 via-slate-800 to-indigo-950 rounded-2xl p-6 text-white border border-slate-700 shadow-md">
                    <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6">
                      <div className="space-y-2 max-w-2xl">
                        <div className="flex items-center gap-3">
                          <span className="px-3 py-1 rounded-full text-xs font-bold uppercase tracking-wider bg-blue-500/20 text-blue-300 border border-blue-500/30">
                            ⚡ Firebase Cloud Functions
                          </span>
                          <span className="px-3 py-1 rounded-full text-xs font-medium bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 flex items-center gap-1.5">
                            <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
                            Tự động quét hằng ngày (08:00 AM)
                          </span>
                        </div>
                        <h3 className="text-xl font-bold text-white">
                          Cảnh báo Email Tự động Trước 3 Ngày Đến Hạn PM
                        </h3>
                        <p className="text-sm text-slate-300 leading-relaxed">
                          Hệ thống tự động lọc các phiếu bảo trì định kỳ có ngày đến hạn (<code className="text-blue-300 font-mono text-xs">dueDate</code>) còn đúng 3 ngày, gửi email HTML thông báo chi tiết đến kỹ thuật viên phụ trách và quản lý (<code className="text-amber-300 font-mono text-xs">sgm1707@gmail.com</code>).
                        </p>
                      </div>

                      <div className="flex flex-col sm:flex-row lg:flex-col gap-3 shrink-0">
                        <div className="bg-white/10 backdrop-blur-sm rounded-xl p-4 border border-white/10 flex items-center justify-between gap-4">
                          <div>
                            <div className="text-xs text-slate-400 font-medium">Phiếu PM sắp đến hạn (&le; 3 ngày):</div>
                            <div className="text-2xl font-black text-amber-400">
                              {pmUpcomingTasks.length} <span className="text-sm font-normal text-slate-300">thiết bị</span>
                            </div>
                          </div>
                          <button
                            disabled={pmNotificationLoading}
                            onClick={() => handleSendPMAlerts(false)}
                            className="px-3.5 py-2 bg-blue-600 hover:bg-blue-500 active:bg-blue-700 disabled:opacity-50 text-white text-xs font-semibold rounded-lg transition-colors flex items-center gap-1.5 shadow-sm"
                          >
                            {pmNotificationLoading ? <RefreshCw size={14} className="animate-spin" /> : <Send size={14} />}
                            <span>Quét & Gửi ngay</span>
                          </button>
                        </div>

                        <button
                          onClick={() => setShowPMNotificationModal(true)}
                          className="px-4 py-2 bg-white/15 hover:bg-white/25 text-white text-xs font-medium rounded-xl border border-white/20 transition-colors text-center"
                        >
                          Xem chi tiết danh sách & Cấu hình gửi mail
                        </button>
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* PM EMAIL NOTIFICATION MODAL */}
          {showPMNotificationModal && (
            <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200">
              <div className="bg-white w-full max-w-3xl rounded-2xl shadow-2xl border border-slate-200 max-h-[90vh] flex flex-col overflow-hidden">
                {/* Header */}
                <div className="p-6 border-b border-slate-100 flex items-center justify-between bg-slate-50/70">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-xl bg-blue-600 text-white flex items-center justify-center shadow-sm">
                      <Mail size={20} />
                    </div>
                    <div>
                      <h3 className="text-lg font-bold text-slate-900">
                        Cấu hình & Giám sát Email Cảnh báo PM
                      </h3>
                      <p className="text-xs text-slate-500">
                        Firebase Cloud Functions tự động kích hoạt trước ngày đến hạn 3 ngày
                      </p>
                    </div>
                  </div>
                  <button
                    onClick={() => setShowPMNotificationModal(false)}
                    className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-200/60 rounded-xl transition-colors"
                  >
                    <X size={18} />
                  </button>
                </div>

                {/* Content */}
                <div className="p-6 overflow-y-auto space-y-6 flex-1 text-sm">
                  {/* Status Banner */}
                  <div className="p-4 bg-blue-50/60 rounded-xl border border-blue-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                    <div>
                      <div className="font-semibold text-blue-950 flex items-center gap-2">
                        <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 animate-pulse"></span>
                        Firebase Cloud Function Scheduler: <code className="bg-blue-100 px-1.5 py-0.5 rounded text-blue-800 font-mono text-xs">checkUpcomingPMTasks</code>
                      </div>
                      <div className="text-xs text-blue-700 mt-1">
                        Chu kỳ: Tự động chạy mỗi ngày lúc 08:00 sáng (Asia/Ho_Chi_Minh). Lọc các phiếu có <code className="font-mono">dueDate == Hôm nay + 3 ngày</code>.
                      </div>
                    </div>
                    <span className="shrink-0 px-3 py-1 bg-emerald-100 text-emerald-800 text-xs font-semibold rounded-full border border-emerald-200">
                      Sẵn sàng hoạt động
                    </span>
                  </div>

                  {/* Upcoming PM List */}
                  <div>
                    <div className="flex items-center justify-between mb-3">
                      <h4 className="font-bold text-slate-900 flex items-center gap-2">
                        <span>Phiếu PM sắp đến hạn đúng 3 ngày</span>
                        <span className="px-2 py-0.5 text-xs bg-slate-100 text-slate-700 rounded-full font-semibold">
                          {pmUpcomingTasks.length}
                        </span>
                      </h4>
                      <div className="flex items-center gap-2">
                        <button
                          disabled={pmNotificationLoading}
                          onClick={checkUpcomingPMAlerts}
                          className="px-2.5 py-1 text-xs text-slate-600 hover:text-slate-900 hover:bg-slate-100 rounded-lg flex items-center gap-1 border border-slate-200"
                        >
                          <RefreshCw size={12} className={pmNotificationLoading ? 'animate-spin' : ''} />
                          <span>Làm mới</span>
                        </button>
                        {pmUpcomingTasks.length > 0 && (
                          <button
                            disabled={pmNotificationLoading}
                            onClick={() => handleSendPMAlerts(true)}
                            className="px-3 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-semibold flex items-center gap-1"
                          >
                            <Send size={12} />
                            <span>Gửi cảnh báo ngay</span>
                          </button>
                        )}
                      </div>
                    </div>

                    {pmUpcomingTasks.length === 0 ? (
                      <div className="p-6 bg-slate-50 rounded-xl border border-dashed border-slate-200 text-center text-slate-500">
                        <Calendar size={28} className="mx-auto mb-2 text-slate-400" />
                        <p className="font-medium text-slate-700">Hiện không có phiếu PM nào đến hạn trong đúng 3 ngày tới.</p>
                        <p className="text-xs text-slate-400 mt-1">
                          Khi có phiếu bảo trì với <code className="font-mono text-slate-600">dueDate</code> cách hiện tại 3 ngày, hệ thống sẽ tự động gửi email.
                        </p>
                      </div>
                    ) : (
                      <div className="space-y-3">
                        {pmUpcomingTasks.map((task) => (
                          <div key={task.id} className="p-4 bg-white rounded-xl border border-slate-200 shadow-xs flex flex-col md:flex-row md:items-center justify-between gap-3">
                            <div className="space-y-1">
                              <div className="flex items-center gap-2">
                                <span className="font-bold text-slate-900">{task.title}</span>
                                <span className="px-2 py-0.5 bg-amber-100 text-amber-800 text-[11px] font-bold rounded-md">
                                  Hạn: {task.dueDate}
                                </span>
                              </div>
                              <div className="text-xs text-slate-500 flex flex-wrap gap-x-4 gap-y-1">
                                <span>Thiết bị: <strong className="text-slate-700">{task.equipmentName || task.equipmentId}</strong></span>
                                <span>Khách hàng: <strong className="text-slate-700">{task.customerName || task.customerId}</strong></span>
                                <span>Người thực hiện: <strong className="text-slate-700">{task.assignedTo || 'Chưa gán'}</strong></span>
                              </div>
                            </div>
                            <div className="text-xs text-slate-500 shrink-0">
                              <span>Gửi tới: <code className="text-blue-600">{task.recipients?.join(', ') || 'sgm1707@gmail.com'}</code></span>
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>

                  {/* Send Test Email Card */}
                  <div className="p-4 rounded-xl border border-slate-200 bg-slate-50/50 space-y-3">
                    <h4 className="font-bold text-slate-900 flex items-center gap-2">
                      <Mail size={16} className="text-blue-600" />
                      <span>Thử nghiệm gửi Email Thông báo</span>
                    </h4>
                    <p className="text-xs text-slate-500">
                      Gửi một email kiểm tra với khuôn mẫu chuẩn (HTML format) cảnh báo PM để xem giao diện hộp thư của bạn.
                    </p>
                    <div className="flex flex-col sm:flex-row items-center gap-2">
                      <input
                        type="email"
                        value={testEmailRecipient}
                        onChange={(e) => setTestEmailRecipient(e.target.value)}
                        placeholder="Nhập địa chỉ email nhận..."
                        className="w-full sm:flex-1 px-3.5 py-2 text-xs border border-slate-300 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-blue-500 bg-white"
                      />
                      <button
                        disabled={pmNotificationLoading || !testEmailRecipient}
                        onClick={handleSendTestEmail}
                        className="w-full sm:w-auto px-4 py-2 bg-slate-900 hover:bg-slate-800 active:bg-slate-950 disabled:opacity-50 text-white text-xs font-semibold rounded-lg flex items-center justify-center gap-2 transition-colors shrink-0 shadow-sm"
                      >
                        {pmNotificationLoading ? <RefreshCw size={14} className="animate-spin" /> : <Send size={14} />}
                        <span>Gửi email test</span>
                      </button>
                    </div>

                    {pmNotificationResult && (
                      <div className={`p-3 rounded-lg text-xs ${pmNotificationResult.success ? 'bg-emerald-50 text-emerald-900 border border-emerald-200' : 'bg-rose-50 text-rose-900 border border-rose-200'}`}>
                        <div className="font-semibold">{pmNotificationResult.message || (pmNotificationResult.success ? 'Thành công' : 'Thất bại')}</div>
                        {pmNotificationResult.mode === 'simulated' && (
                          <div className="text-[11px] text-emerald-700 mt-1">
                            * Chế độ Mô phỏng (Simulated): Email đã được định dạng chuẩn HTML và ghi nhận an toàn. Để gửi qua hộp thư thực tế, cấu hình <code className="font-mono">SMTP_HOST</code>, <code className="font-mono">SMTP_USER</code>, <code className="font-mono">SMTP_PASS</code> trong môi trường hệ thống.
                          </div>
                        )}
                      </div>
                    )}
                  </div>

                  {/* Firebase Cloud Functions Deployment Info */}
                  <div className="p-4 rounded-xl border border-slate-200 bg-white space-y-2">
                    <h4 className="font-bold text-slate-900 flex items-center gap-2">
                      <FileText size={16} className="text-indigo-600" />
                      <span>Hướng dẫn Triển khai Cloud Functions (Firebase CLI)</span>
                    </h4>
                    <p className="text-xs text-slate-600">
                      Mã nguồn Firebase Cloud Functions đã được tạo sẵn trong thư mục <code className="font-mono bg-slate-100 px-1 py-0.5 rounded text-indigo-700">functions/src/index.ts</code>. Để triển khai lên dự án Firebase của bạn:
                    </p>
                    <div className="bg-slate-900 text-slate-100 p-3 rounded-lg font-mono text-xs overflow-x-auto select-all">
                      firebase deploy --only functions
                    </div>
                  </div>
                </div>

                {/* Footer */}
                <div className="p-4 border-t border-slate-100 bg-slate-50 flex justify-end">
                  <button
                    onClick={() => setShowPMNotificationModal(false)}
                    className="px-5 py-2 bg-slate-800 hover:bg-slate-900 text-white rounded-xl text-xs font-semibold transition-colors"
                  >
                    Đóng
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* --- INVENTORY VIEW --- */}
          {activeTab === 'inventory' && (
            <div className="space-y-6">
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div>
                  <h2 className="text-2xl font-bold text-slate-900">Quản lý Kho vật tư</h2>
                  <p className="text-slate-500">Theo dõi tồn kho phụ tùng, thiết bị thay thế.</p>
                </div>
                <button 
                  onClick={() => {
                    setNewInventoryItem({
                      name: '',
                      sku: '',
                      category: '',
                      quantity: 0,
                      unit: 'pcs',
                      minStock: 5,
                      location: '',
                      price: 0
                    });
                    setSelectedInventoryItem(null);
                    setShowInventoryModal(true);
                  }}
                  className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg font-medium hover:bg-blue-700 transition-colors shadow-sm"
                >
                  <Plus size={18} />
                  Thêm vật tư
                </button>
              </div>

              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="border-b border-slate-100 text-xs uppercase tracking-wider text-slate-500 font-semibold bg-slate-50/50">
                        <th className="p-4">Tên vật tư / SKU</th>
                        <th className="p-4">Danh mục</th>
                        <th className="p-4">Số lượng</th>
                        <th className="p-4">Đơn vị</th>
                        <th className="p-4">Vị trí kho</th>
                        <th className="p-4">Trạng thái</th>
                        <th className="p-4 text-right">Thao tác</th>
                      </tr>
                    </thead>
                    <tbody className="text-sm divide-y divide-slate-50">
                      {isInventoryLoading ? (
                        <tr><td colSpan={7} className="p-8 text-center text-slate-500">Đang tải...</td></tr>
                      ) : inventory.length === 0 ? (
                        <tr><td colSpan={7} className="p-8 text-center text-slate-500">Kho trống.</td></tr>
                      ) : inventory.map((item) => (
                        <tr key={item.id} className="hover:bg-slate-50 transition-colors">
                          <td className="p-4">
                            <div className="font-bold text-slate-900">{item.name}</div>
                            <div className="text-xs text-slate-500 font-mono">{item.sku}</div>
                          </td>
                          <td className="p-4 text-slate-600">{item.category || 'N/A'}</td>
                          <td className="p-4 font-bold text-slate-900">{item.quantity}</td>
                          <td className="p-4 text-slate-500">{item.unit}</td>
                          <td className="p-4 text-slate-500">{item.location || 'N/A'}</td>
                          <td className="p-4">
                            {item.quantity <= item.minStock ? (
                              <span className="px-2 py-0.5 rounded-full text-[10px] font-bold uppercase bg-rose-100 text-rose-700 border border-rose-200">
                                Sắp hết hàng
                              </span>
                            ) : (
                              <span className="px-2 py-0.5 rounded-full text-[10px] font-bold uppercase bg-emerald-100 text-emerald-700 border border-emerald-200">
                                Sẵn sàng
                              </span>
                            )}
                          </td>
                          <td className="p-4 text-right">
                            <button 
                              onClick={() => {
                                setSelectedInventoryItem(item);
                                setNewInventoryItem({ ...item });
                                setShowInventoryModal(true);
                              }}
                              className="p-2 text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                            >
                              <Settings size={16} />
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* --- CUSTOMERS VIEW --- */}
          {activeTab === 'customers' && (
            <div className="space-y-6">
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div>
                  <h2 className="text-2xl font-bold text-slate-900">Quản lý Khách hàng</h2>
                  <p className="text-slate-500">Danh sách khách hàng và đối tác.</p>
                </div>
                <div className="flex items-center gap-3">
                  <button 
                    onClick={() => handleFetchFromSheets()}
                    disabled={isSyncing}
                    className="flex items-center gap-2 px-4 py-2 bg-white border border-slate-200 text-slate-700 rounded-lg font-medium hover:bg-slate-50 transition-colors shadow-sm disabled:opacity-50"
                  >
                    <Download size={18} className={isSyncing ? 'animate-spin' : ''} />
                    Tải từ Sheets
                  </button>
                  <button 
                    onClick={() => handleSyncToSheets()}
                    disabled={isSyncing}
                    className="flex items-center gap-2 px-4 py-2 bg-white border border-slate-200 text-slate-700 rounded-lg font-medium hover:bg-slate-50 transition-colors shadow-sm disabled:opacity-50"
                  >
                    <Upload size={18} className={isSyncing ? 'animate-spin' : ''} />
                    Đồng bộ lên Sheets
                  </button>
                  <button 
                    onClick={() => {
                      setNewCustomer({ name: '', email: '', phone: '', address: '', factories: [] });
                      setEditingCustomer(null);
                      setShowCustomerModal(true);
                    }}
                    className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg font-medium hover:bg-blue-700 transition-colors shadow-sm"
                  >
                    <Plus size={18} />
                    Thêm khách hàng
                  </button>
                </div>
              </div>

              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="border-b border-slate-100 text-xs uppercase tracking-wider text-slate-500 font-semibold bg-slate-50/50">
                        <th className="p-4">Tên khách hàng</th>
                        <th className="p-4">Nhà máy</th>
                        <th className="p-4">Email</th>
                        <th className="p-4">Số điện thoại</th>
                        <th className="p-4">Địa chỉ</th>
                        <th className="p-4 text-right">Thao tác</th>
                      </tr>
                    </thead>
                    <tbody className="text-sm divide-y divide-slate-50">
                      {isCustomersLoading ? (
                        <tr><td colSpan={6} className="p-8 text-center text-slate-500">Đang tải...</td></tr>
                      ) : customers.length === 0 ? (
                        <tr><td colSpan={6} className="p-8 text-center text-slate-500">Chưa có khách hàng nào.</td></tr>
                      ) : customers.map((customer) => {
                        const customerFactories = customer.factories || [];
                        return (
                        <tr key={customer.id} className="hover:bg-slate-50 transition-colors">
                          <td className="p-4 font-bold text-slate-900">{customer.name}</td>
                          <td className="p-4">
                            {customerFactories.length > 0 ? (
                              <div className="flex flex-wrap gap-1">
                                {customerFactories.map((factory, idx) => (
                                  <span key={idx} className="px-2 py-0.5 bg-blue-50 text-blue-700 border border-blue-100 rounded-md text-xs">
                                    {factory}
                                  </span>
                                ))}
                              </div>
                            ) : (
                              <span className="text-slate-400 text-xs italic">Chưa có dữ liệu</span>
                            )}
                          </td>
                          <td className="p-4 text-slate-600">{customer.email || 'N/A'}</td>
                          <td className="p-4 text-slate-600">{customer.phone || 'N/A'}</td>
                          <td className="p-4 text-slate-500">{customer.address || 'N/A'}</td>
                          <td className="p-4 text-right space-x-2">
                            <button 
                              onClick={() => {
                                setEditingCustomer(customer);
                                setNewCustomer({ ...customer });
                                setShowCustomerModal(true);
                              }}
                              className="p-2 text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                            >
                              <Settings size={16} />
                            </button>
                            <button 
                              onClick={() => handleDeleteCustomer(customer.id)}
                              className="p-2 text-red-600 hover:bg-red-50 rounded-lg transition-colors"
                            >
                              <Trash2 size={16} />
                            </button>
                          </td>
                        </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* --- SETTINGS VIEW --- */}
          {activeTab === 'settings' && (
            <div className="max-w-4xl mx-auto space-y-6">
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="p-6 border-b border-slate-100">
                  <h2 className="text-xl font-bold text-slate-900 flex items-center gap-2">
                    <Settings className="text-blue-500" size={24} />
                    Cài đặt tài khoản & Hệ thống
                  </h2>
                </div>
                
                <div className="p-6 space-y-8">
                  {/* User Profile Section */}
                  <section>
                    <div className="flex items-center justify-between mb-4">
                      <h3 className="text-sm font-bold text-slate-500 uppercase tracking-wider">Thông tin tài khoản</h3>
                      <button 
                        onClick={async () => {
                          if (user) {
                            setIsAuthLoading(true);
                            const userDocRef = doc(db, 'users', user.uid);
                            const userDoc = await getDoc(userDocRef);
            if (userDoc.exists()) {
              const userData = userDoc.data();
              const rawRole = userData.role?.toLowerCase() || 'customer';
              const normalizedRole = (rawRole === 'admin' || rawRole === 'quản trị viên' || rawRole === 'quan tri vien') ? 'admin' : 'customer';
              
              setUserRole(normalizedRole);
              if (userData.assignedFactory) {
                setUserFactory(userData.assignedFactory.trim());
              }
              if (normalizedRole === 'customer' && userData.customerId) {
                const customerDoc = await getDoc(doc(db, 'customers', userData.customerId));
                if (customerDoc.exists()) {
                  setUserCustomerName(customerDoc.data().name);
                }
              }
            }
                            setIsAuthLoading(false);
                            alert('Đã cập nhật thông tin tài khoản!');
                          }
                        }}
                        className="flex items-center gap-1 text-xs font-medium text-blue-600 hover:text-blue-700 transition-colors"
                      >
                        <RefreshCw size={14} />
                        Làm mới dữ liệu
                      </button>
                    </div>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6 bg-slate-50 p-6 rounded-xl border border-slate-100">
                      <div>
                        <label className="block text-xs font-medium text-slate-500 mb-1">Email đăng nhập</label>
                        <p className="text-slate-900 font-medium">{user?.email}</p>
                      </div>
                      <div>
                        <label className="block text-xs font-medium text-slate-500 mb-1">Vai trò hệ thống</label>
                        <div className="flex items-center gap-2">
                          <span className={`px-2 py-0.5 rounded-full text-xs font-bold uppercase ${
                            userRole === 'admin' ? 'bg-blue-100 text-blue-700' : 'bg-emerald-100 text-emerald-700'
                          }`}>
                            {userRole === 'admin' ? 'Quản trị viên' : 'Khách hàng'}
                          </span>
                        </div>
                      </div>
                      <div>
                        <label className="block text-xs font-medium text-slate-500 mb-1">Nhà máy được chỉ định</label>
                        <p className="text-slate-900 font-medium">{userFactory || 'Tất cả (Admin)'}</p>
                      </div>
                      <div>
                        <label className="block text-xs font-medium text-slate-500 mb-1">User ID (UID)</label>
                        <div className="flex items-center gap-2">
                          <p className="text-xs font-mono text-slate-600 bg-white px-2 py-1 border border-slate-200 rounded select-all break-all">
                            {user?.uid || 'Đang tải...'}
                          </p>
                          {user?.uid && (
                            <button 
                              onClick={() => {
                                navigator.clipboard.writeText(user.uid);
                                alert('Đã sao chép UID vào bộ nhớ tạm!');
                              }}
                              className="p-1.5 text-blue-600 hover:bg-blue-50 rounded transition-colors"
                              title="Sao chép UID"
                            >
                              <Copy size={14} />
                            </button>
                          )}
                        </div>
                      </div>
                    </div>
                  </section>

                  {/* Admin Instructions Section */}
                  {userRole === 'admin' && (
                    <section className="bg-blue-50 border border-blue-100 p-6 rounded-xl">
                      <h3 className="text-blue-800 font-bold mb-3 flex items-center gap-2">
                        <Info size={18} />
                        Hướng dẫn phân quyền khách hàng
                      </h3>
                      <div className="text-sm text-blue-700 space-y-3">
                        <p>Để giới hạn một khách hàng chỉ xem được dữ liệu của một nhà máy cụ thể (ví dụ: <strong>Nhiệt điện Cà Mau</strong>), bạn cần thực hiện các bước sau trong Firebase Console:</p>
                        <ol className="list-decimal ml-5 space-y-2">
                          <li>Truy cập vào <strong>Firestore Database</strong>.</li>
                          <li>Tìm đến collection <code>users</code>.</li>
                          <li>Tìm document có ID trùng với <strong>UID</strong> của khách hàng đó.</li>
                          <li>Thêm trường (Field) mới:
                            <ul className="list-disc ml-5 mt-1">
                              <li>Field name: <code>assignedFactory</code></li>
                              <li>Type: <code>string</code></li>
                              <li>Value: <code>Nhiệt điện Cà Mau</code> (Phải khớp chính xác với tên nhà máy trong database)</li>
                            </ul>
                          </li>
                          <li>Đảm bảo trường <code>role</code> của người dùng đó là <code>customer</code>.</li>
                        </ol>
                        <p className="mt-4 font-medium italic">Hệ thống sẽ tự động lọc toàn bộ Dashboard, Thiết bị và Báo cáo dựa trên nhà máy này khi khách hàng đăng nhập.</p>
                      </div>
                    </section>
                  )}

                  {/* Google Sheets Synchronization Section */}
                  <section className="space-y-4">
                    <h3 className="text-sm font-bold text-slate-500 uppercase tracking-wider">Đồng bộ dữ liệu Google Sheets</h3>
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                      <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm space-y-3">
                        <div className="flex items-center gap-2 text-slate-900 font-bold">
                          <ClipboardList size={18} className="text-emerald-500" />
                          Phiếu công việc (CMMS)
                        </div>
                        <p className="text-xs text-slate-500">Đồng bộ toàn bộ lịch sử phiếu công việc và bảo trì.</p>
                        <button 
                          disabled={isSyncing}
                          onClick={async () => {
                            if (!isGoogleConnected) { alert('Vui lòng kết nối Google Drive trước.'); return; }
                            if (workOrders.length === 0) { alert('Không có dữ liệu phiếu công việc để đồng bộ.'); return; }
                            try {
                              setIsSyncing(true);
                              const allRows = workOrders.map(wo => [
                                wo.id, wo.workPermitId || '', wo.title, (Array.isArray(wo.equipmentId) ? wo.equipmentId.join(', ') : wo.equipmentId), wo.customer, wo.type, wo.status, wo.priority, wo.assignedTo, wo.description, wo.failureCode || '', wo.pmFrequency || '', wo.estimatedTime || 0, wo.laborCount || 0, wo.laborCost || 0, wo.partCost || 0, (wo.attachments || []).join(', '), wo.createdAt, wo.updatedAt
                              ]);
                              await syncToSheet('CMMS', allRows);
                              alert('Đồng bộ CMMS thành công!');
                            } catch (err: any) { 
                              alert(`Lỗi đồng bộ: ${err.message}`); 
                            } finally {
                              setIsSyncing(false);
                            }
                          }}
                          className={`w-full py-2 rounded-lg text-sm font-medium transition-colors flex items-center justify-center gap-2 ${
                            !isGoogleConnected 
                              ? 'bg-slate-100 text-slate-400 hover:bg-slate-200' 
                              : 'bg-slate-100 hover:bg-slate-200 text-slate-700'
                          } ${isSyncing ? 'opacity-50 cursor-not-allowed' : ''}`}
                        >
                          <RefreshCw size={14} className={isSyncing ? 'animate-spin' : ''} />
                          {isSyncing ? 'Đang đồng bộ...' : 'Đồng bộ ngay'}
                        </button>
                      </div>

                      <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm space-y-3">
                        <div className="flex items-center gap-2 text-slate-900 font-bold">
                          <Package size={18} className="text-amber-500" />
                          Quản lý Kho
                        </div>
                        <p className="text-xs text-slate-500">Đồng bộ toàn bộ danh mục vật tư và số lượng tồn kho.</p>
                        <button 
                          disabled={isSyncing}
                          onClick={async () => {
                            if (!isGoogleConnected) { alert('Vui lòng kết nối Google Drive trước.'); return; }
                            if (inventory.length === 0) { alert('Không có dữ liệu kho để đồng bộ.'); return; }
                            try {
                              setIsSyncing(true);
                              const allRows = inventory.map(item => [
                                item.id, item.name, item.sku, item.category, item.quantity, item.minStock, item.unit, item.location, item.price, item.lastRestocked || '', item.updatedAt
                              ]);
                              await syncToSheet('QuanLyKho', allRows);
                              alert('Đồng bộ Kho thành công!');
                            } catch (err: any) { 
                              alert(`Lỗi đồng bộ: ${err.message}`); 
                            } finally {
                              setIsSyncing(false);
                            }
                          }}
                          className={`w-full py-2 rounded-lg text-sm font-medium transition-colors flex items-center justify-center gap-2 ${
                            !isGoogleConnected 
                              ? 'bg-slate-100 text-slate-400 hover:bg-slate-200' 
                              : 'bg-slate-100 hover:bg-slate-200 text-slate-700'
                          } ${isSyncing ? 'opacity-50 cursor-not-allowed' : ''}`}
                        >
                          <RefreshCw size={14} className={isSyncing ? 'animate-spin' : ''} />
                          {isSyncing ? 'Đang đồng bộ...' : 'Đồng bộ ngay'}
                        </button>
                      </div>
                    </div>
                  </section>
                </div>
              </div>
            </div>
          )}
        </div>
      </main>

      {/* MODALS */}
      {selectedRiskDetail && currentDetailedRiskData && (
        <div className="fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-4 backdrop-blur-sm">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-5xl max-h-[90vh] flex flex-col overflow-hidden border border-slate-200">
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-white">
              <div className="flex items-center gap-3">
                <div className={`w-10 h-10 rounded-full flex items-center justify-center ${
                  currentDetailedRiskData.status === 'critical' ? 'bg-rose-100 text-rose-600' : 
                  currentDetailedRiskData.status === 'warning' ? 'bg-amber-100 text-amber-600' : 'bg-emerald-100 text-emerald-600'
                }`}>
                  <AlertTriangle size={20} />
                </div>
                <div>
                  <h2 className="text-lg font-bold text-slate-900">Chi tiết Chỉ số sức khỏe (Health Index)</h2>
                  <p className="text-sm text-slate-500 font-medium">{currentDetailedRiskData.name} ({currentDetailedRiskData.id})</p>
                </div>
              </div>
              <button onClick={() => setSelectedRiskDetail(null)} className="p-2 text-slate-400 hover:text-slate-600 hover:bg-slate-100 rounded-full transition-colors">
                <X size={20} />
              </button>
            </div>
            
            <div className="flex-1 overflow-y-auto p-6 bg-slate-50/50">
              {/* Top Overview */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">
                <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm flex flex-col justify-center items-center text-center">
                  <div className="text-sm font-semibold text-slate-500 uppercase tracking-wider mb-2">Health Index</div>
                  <div className="flex items-end gap-2 justify-center">
                    <span className={`text-5xl font-black ${
                      currentDetailedRiskData.status === 'critical' ? 'text-rose-600' : 
                      currentDetailedRiskData.status === 'warning' ? 'text-amber-500' : 'text-emerald-500'
                    }`}>
                      {currentDetailedRiskData.health}
                    </span>
                    <span className="text-xl font-bold text-slate-400 mb-1">/100</span>
                  </div>
                  <div className="mt-3">
                    <StatusBadge status={currentDetailedRiskData.status} />
                  </div>
                </div>

                <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm border-l-4 border-l-rose-500 flex flex-col justify-center">
                  <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-2 flex items-center gap-2">
                    <AlertCircle size={16} className="text-rose-500" />
                    Khuyến nghị chung
                  </h3>
                  <p className="text-slate-700 leading-relaxed text-sm">
                    {currentDetailedRiskData.generalRecommendation}
                  </p>
                </div>
              </div>

              {/* Health Matrix (Failure Profile & Defects) */}
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden mb-6">
                <div className="grid grid-cols-1 lg:grid-cols-5 divide-y lg:divide-y-0 lg:divide-x divide-slate-200">
                  {/* Failure Profile (Radar Chart) */}
                  <div className="p-6 lg:col-span-2 flex flex-col">
                    <h3 className="text-sm font-bold text-slate-500 uppercase tracking-wider mb-4">Failure Profile</h3>
                    <div className="flex-1 min-h-[250px] w-full relative">
                      <ResponsiveContainer width="100%" height="100%">
                        <RadarChart cx="50%" cy="50%" outerRadius="70%" data={currentDetailedRiskData.failureProfile}>
                          <PolarGrid stroke="#e2e8f0" />
                          <PolarAngleAxis dataKey="subject" tick={{ fill: '#475569', fontSize: 11, fontWeight: 500 }} />
                          <PolarRadiusAxis angle={30} domain={[0, 100]} tick={{ fill: '#94a3b8', fontSize: 10 }} axisLine={false} />
                          <Radar 
                            name="Health" 
                            dataKey="score" 
                            stroke={
                              currentDetailedRiskData.status === 'critical' ? '#f43f5e' : 
                              currentDetailedRiskData.status === 'warning' ? '#f59e0b' : '#10b981'
                            } 
                            strokeWidth={2} 
                            fill={
                              currentDetailedRiskData.status === 'critical' ? '#f43f5e' : 
                              currentDetailedRiskData.status === 'warning' ? '#f59e0b' : '#10b981'
                            } 
                            fillOpacity={0.3} 
                          />
                        </RadarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>

                  {/* Defects Table */}
                  <div className="p-6 lg:col-span-3 flex flex-col">
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse min-w-[500px]">
                        <thead>
                          <tr className="border-b border-slate-200">
                            <th className="pb-3 px-2 text-xs font-semibold text-slate-500 uppercase tracking-wider w-1/3">Defect Type</th>
                            <th className="pb-3 px-2 text-xs font-semibold text-slate-500 uppercase tracking-wider w-1/4 text-center">Warnings</th>
                            <th className="pb-3 px-2 text-xs font-semibold text-slate-500 uppercase tracking-wider">Risk %</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {currentDetailedRiskData.defects.map((defect: any, idx: number) => (
                            <tr key={idx} className="hover:bg-slate-50/50 transition-colors">
                              <td className="py-3 px-2 text-sm font-medium text-slate-700">{defect.name}</td>
                              <td className="py-3 px-2 text-center">
                                <span className="inline-block px-2 py-1 bg-slate-100 text-slate-700 rounded text-xs font-semibold border border-slate-200">
                                  {defect.warnings.toFixed(1)}%
                                </span>
                              </td>
                              <td className="py-3 px-2">
                                <div className="flex items-center gap-3">
                                  <div className="flex-1 h-2 bg-slate-100 rounded-full overflow-hidden relative border border-slate-200">
                                    {/* Ticks for 0, 25, 50, 75, 100 */}
                                    <div className="absolute top-0 bottom-0 left-1/4 w-px bg-white/50 z-10"></div>
                                    <div className="absolute top-0 bottom-0 left-2/4 w-px bg-white/50 z-10"></div>
                                    <div className="absolute top-0 bottom-0 left-3/4 w-px bg-white/50 z-10"></div>
                                    
                                    <div 
                                      className={`h-full rounded-full ${
                                        defect.risk > 70 ? 'bg-rose-500' : 
                                        defect.risk > 40 ? 'bg-amber-500' : 'bg-emerald-500'
                                      }`}
                                      style={{ width: `${defect.risk}%` }}
                                    ></div>
                                  </div>
                                  <span className="text-xs font-bold text-slate-700 w-8 text-right">
                                    {defect.risk.toFixed(1)}
                                  </span>
                                </div>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              </div>

              {/* Detailed Table */}
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="px-5 py-4 border-b border-slate-100 bg-white">
                  <h3 className="font-bold text-slate-800">Bảng phân tích thông số chi tiết</h3>
                  <p className="text-xs text-slate-500 mt-1">Dựa trên tiêu chuẩn đánh giá IEEE & IEC</p>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse min-w-[800px]">
                    <thead>
                      <tr className="bg-slate-50 border-b border-slate-200 text-xs uppercase tracking-wider text-slate-600 font-bold">
                        <th className="p-4">Hạng mục kiểm tra (Tests)</th>
                        <th className="p-4">Giá trị đo</th>
                        <th className="p-4 text-center">Trọng số</th>
                        <th className="p-4 text-center">Điểm (1-10)</th>
                        <th className="p-4">Đánh giá</th>
                        <th className="p-4 w-1/3">Khuyến nghị</th>
                      </tr>
                    </thead>
                    <tbody className="text-sm divide-y divide-slate-100">
                      {currentDetailedRiskData.tests.map((test: any, idx: number) => (
                        <tr key={idx} className="hover:bg-slate-50 transition-colors">
                          <td className="p-4 font-medium text-slate-800">{test.name}</td>
                          <td className="p-4 font-mono text-slate-700 font-semibold">
                            {test.value} <span className="text-xs text-slate-500 font-sans font-normal">{test.unit}</span>
                          </td>
                          <td className="p-4 text-center text-slate-500 font-medium">x{test.weight}</td>
                          <td className="p-4 text-center">
                            <span className={`inline-flex items-center justify-center w-8 h-8 rounded-full font-bold ${
                              test.score >= 8 ? 'bg-emerald-100 text-emerald-700' :
                              test.score >= 6 ? 'bg-blue-100 text-blue-700' :
                              test.score >= 4 ? 'bg-amber-100 text-amber-700' :
                              'bg-rose-100 text-rose-700'
                            }`}>
                              {test.score}
                            </span>
                          </td>
                          <td className="p-4">
                            <TestStatusBadge status={test.status} />
                          </td>
                          <td className="p-4 text-slate-600 leading-relaxed text-xs">
                            {test.rec}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
            
            <div className="px-6 py-4 border-t border-slate-200 bg-white flex justify-end">
              <button 
                onClick={() => setSelectedRiskDetail(null)}
                className="px-5 py-2.5 bg-slate-100 text-slate-700 rounded-lg font-medium hover:bg-slate-200 transition-colors"
              >
                Đóng
              </button>
            </div>
          </div>
        </div>
      )}

      {showQRModal && selectedEqForQR && (
        <div className="fixed inset-0 bg-slate-900/50 z-[60] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-sm overflow-hidden animate-in fade-in zoom-in duration-200">
            <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50">
              <h3 className="font-bold text-slate-800 flex items-center gap-2">
                <QrCode size={20} className="text-blue-600" />
                Mã QR Thiết bị
              </h3>
              <button onClick={() => setShowQRModal(false)} className="p-2 text-slate-400 hover:text-slate-600 hover:bg-slate-200 rounded-full transition-colors">
                <X size={20} />
              </button>
            </div>
            <div className="p-8 flex flex-col items-center text-center">
              <div className="bg-white p-4 rounded-2xl border-4 border-slate-100 shadow-inner mb-6">
                <QRCodeSVG 
                  value={`${window.location.origin}${window.location.pathname}?eqId=${selectedEqForQR.id}`}
                  size={200}
                  level="H"
                  includeMargin={true}
                />
              </div>
              <h4 className="text-lg font-bold text-slate-900 mb-1">{selectedEqForQR.name}</h4>
              <p className="text-sm text-slate-500 font-mono mb-6">{selectedEqForQR.id}</p>
              
              <div className="w-full grid grid-cols-2 gap-3">
                <button 
                  onClick={() => {
                    const svg = document.querySelector('.p-8 svg');
                    if (svg) {
                      const svgData = new XMLSerializer().serializeToString(svg);
                      const canvas = document.createElement('canvas');
                      const ctx = canvas.getContext('2d');
                      const img = new Image();
                      img.onload = () => {
                        canvas.width = img.width;
                        canvas.height = img.height;
                        ctx?.drawImage(img, 0, 0);
                        const pngFile = canvas.toDataURL('image/png');
                        const downloadLink = document.createElement('a');
                        downloadLink.download = `QR_${selectedEqForQR.id}.png`;
                        downloadLink.href = pngFile;
                        downloadLink.click();
                      };
                      img.src = 'data:image/svg+xml;base64,' + btoa(svgData);
                    }
                  }}
                  className="flex items-center justify-center gap-2 py-2.5 bg-blue-600 text-white rounded-xl text-sm font-semibold hover:bg-blue-700 transition-colors shadow-sm"
                >
                  <Download size={16} />
                  Tải ảnh
                </button>
                <button 
                  onClick={() => window.print()}
                  className="flex items-center justify-center gap-2 py-2.5 bg-slate-100 text-slate-700 rounded-xl text-sm font-semibold hover:bg-slate-200 transition-colors"
                >
                  <FileText size={16} />
                  In mã
                </button>
              </div>
              <p className="mt-6 text-[10px] text-slate-400 italic">
                Quét mã này để truy cập nhanh hồ sơ thiết bị và lịch sử bảo trì
              </p>
            </div>
          </div>
        </div>
      )}

      {showEquipmentProfile && (
        <div className="fixed inset-0 bg-slate-900/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-4xl max-h-[90vh] flex flex-col overflow-hidden">
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-full bg-blue-100 flex items-center justify-center text-blue-600">
                  <Server size={20} />
                </div>
                <div>
                  <h2 className="text-lg font-bold text-slate-900">Hồ sơ & Lịch sử kiểm tra</h2>
                  <p className="text-sm text-slate-500">Máy biến áp T1 (TRF-01)</p>
                </div>
              </div>
              <button onClick={() => setShowEquipmentProfile(false)} className="p-2 text-slate-400 hover:text-slate-600 hover:bg-slate-200 rounded-full transition-colors">
                <X size={20} />
              </button>
            </div>
            
            <div className="flex-1 overflow-y-auto p-6 space-y-8">
              {/* Health Overview */}
              <div>
                <h3 className="text-base font-bold text-slate-800 mb-4 flex items-center gap-2">
                  <Activity size={18} className="text-blue-500" />
                  Tổng quan sức khỏe thiết bị
                </h3>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-200">
                    <div className="text-sm text-slate-500 mb-1">Chỉ số sức khỏe hiện tại</div>
                    <div className="flex items-end gap-2">
                      <span className="text-3xl font-bold text-rose-600">45%</span>
                      <span className="text-sm font-medium text-rose-600 bg-rose-100 px-2 py-0.5 rounded-full mb-1">Nguy hiểm</span>
                    </div>
                  </div>
                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-200">
                    <div className="text-sm text-slate-500 mb-1">Số lần kiểm tra (12 tháng)</div>
                    <div className="text-3xl font-bold text-slate-800">4</div>
                  </div>
                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-200">
                    <div className="text-sm text-slate-500 mb-1">Vấn đề thường gặp</div>
                    <div className="text-base font-medium text-slate-800">Nhiệt độ dầu cao</div>
                  </div>
                </div>
              </div>

              {/* Maintenance Analytics */}
              <div>
                <h3 className="text-base font-bold text-slate-800 mb-4 flex items-center gap-2">
                  <TrendingUp size={18} className="text-blue-500" />
                  Phân tích & Xu hướng Bảo trì
                </h3>
                <MaintenanceCharts 
                  equipment={allEquipment.find(e => e.id === equipmentCode) || { id: equipmentCode, health: 50, status: 'healthy' }} 
                  workOrders={workOrders} 
                  reports={allReports} 
                />
              </div>

              {/* Past Reports */}
              <div>
                <h3 className="text-base font-bold text-slate-800 mb-4 flex items-center gap-2">
                  <FileText size={18} className="text-blue-500" />
                  Danh sách báo cáo cũ
                </h3>
                <div className="border border-slate-200 rounded-xl overflow-x-auto">
                  <table className="w-full text-left border-collapse min-w-[700px]">
                    <thead>
                      <tr className="bg-slate-50 border-b border-slate-200 text-xs uppercase tracking-wider text-slate-500 font-semibold">
                        <th className="p-4">Mã Báo Cáo</th>
                        <th className="p-4">Ngày kiểm tra</th>
                        <th className="p-4">Người thực hiện</th>
                        <th className="p-4">Đánh giá</th>
                        <th className="p-4">Ghi chú</th>
                        <th className="p-4 text-right">Thao tác</th>
                      </tr>
                    </thead>
                    <tbody className="text-sm divide-y divide-slate-100">
                      {allReports.filter(r => r.equipmentId === equipmentCode).map((report) => (
                        <tr key={report.id} className="hover:bg-slate-50 transition-colors">
                          <td className="p-4 font-medium text-blue-600">{report.id}</td>
                          <td className="p-4 flex items-center gap-2"><Calendar size={14} className="text-slate-400"/> {report.date}</td>
                          <td className="p-4 text-slate-700">{report.inspector}</td>
                          <td className="p-4"><StatusBadge status={report.status} /></td>
                          <td className="p-4 text-slate-500 max-w-[200px] truncate" title={report.notes}>{report.notes}</td>
                          <td className="p-4 text-right">
                            <div className="flex items-center justify-end gap-2">
                              <button className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded transition-colors" title="Xem chi tiết">
                                <Eye size={16} />
                              </button>
                              <button 
                                onClick={() => generateIndividualReportPDF(report, true, allReports.filter(r => r.equipmentId === report.equipmentId))}
                                className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded transition-colors" 
                                title="Tải PDF"
                              >
                                <Download size={16} />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
            
            <div className="px-6 py-4 border-t border-slate-200 bg-slate-50 flex justify-end">
              <button 
                onClick={() => setShowEquipmentProfile(false)}
                className="px-4 py-2 bg-white border border-slate-300 text-slate-700 rounded-lg font-medium hover:bg-slate-50 transition-colors"
              >
                Đóng
              </button>
            </div>
          </div>
        </div>
      )}

      {showAddEqModal && (
        <div className="fixed inset-0 bg-slate-900/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-md overflow-hidden">
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
              <h3 className="font-bold text-slate-800 text-lg">Thêm thiết bị mới</h3>
              <button onClick={() => setShowAddEqModal(false)} className="text-slate-400 hover:text-slate-600 transition-colors">
                <X size={20} />
              </button>
            </div>
            <div className="p-6 space-y-4">
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Mã thiết bị</label>
                <input type="text" value={newEqData.id} onChange={(e) => setNewEqData({...newEqData, id: e.target.value})} className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder="VD: TRF-02" />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Tên thiết bị</label>
                <input type="text" value={newEqData.name} onChange={(e) => setNewEqData({...newEqData, name: e.target.value})} className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder="VD: Máy biến áp T2" />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Khách hàng</label>
                {customers.length > 0 ? (
                  <div className="space-y-2">
                    <select 
                      value={customers.find(c => c.name === newEqData.customer) ? newEqData.customer : (newEqData.customer ? "other" : "")} 
                      onChange={(e) => {
                        if (e.target.value === "other") {
                          setNewEqData({...newEqData, customer: ""});
                        } else {
                          setNewEqData({...newEqData, customer: e.target.value});
                        }
                      }} 
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      <option value="">-- Chọn khách hàng --</option>
                      {customers.map(c => (
                        <option key={c.id} value={c.name}>{c.name}</option>
                      ))}
                      <option value="other">-- Khác (Nhập tay) --</option>
                    </select>
                    {(!customers.find(c => c.name === newEqData.customer) && newEqData.customer !== "" || !customers.length) && (
                      <input 
                        type="text" 
                        value={newEqData.customer} 
                        onChange={(e) => setNewEqData({...newEqData, customer: e.target.value})} 
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                        placeholder="Nhập tên khách hàng mới..." 
                      />
                    )}
                  </div>
                ) : (
                  <input 
                    type="text" 
                    value={newEqData.customer} 
                    onChange={(e) => setNewEqData({...newEqData, customer: e.target.value})} 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                    placeholder="Nhập tên khách hàng..." 
                  />
                )}
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Nhà máy</label>
                {newEqData.customer && customers.find(c => c.name === newEqData.customer)?.factories?.length > 0 ? (
                  <select 
                    value={newEqData.factory} 
                    onChange={(e) => setNewEqData({...newEqData, factory: e.target.value})} 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                  >
                    <option value="">-- Chọn nhà máy --</option>
                    {customers.find(c => c.name === newEqData.customer).factories.map((f, idx) => (
                      <option key={idx} value={f}>{f}</option>
                    ))}
                  </select>
                ) : (
                  <input type="text" value={newEqData.factory} onChange={(e) => setNewEqData({...newEqData, factory: e.target.value})} className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder="VD: Nhà máy Bắc Ninh" />
                )}
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Vị trí</label>
                <input type="text" value={newEqData.location} onChange={(e) => setNewEqData({...newEqData, location: e.target.value})} className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder="VD: Trạm 110kV" />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Loại thiết bị</label>
                <select value={newEqData.type} onChange={(e) => setNewEqData({...newEqData, type: e.target.value})} className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500">
                  <option value="Máy biến áp">Máy biến áp</option>
                  <option value="Tủ điện trung thế">Tủ điện trung thế</option>
                  <option value="Động cơ">Động cơ</option>
                  <option value="Inverter">Inverter Solar/Wind</option>
                </select>
              </div>
            </div>
            <div className="px-6 py-4 border-t border-slate-200 bg-slate-50 flex justify-end gap-3">
              <button onClick={() => setShowAddEqModal(false)} className="px-4 py-2 text-slate-600 font-medium hover:bg-slate-200 rounded-lg transition-colors">Hủy</button>
              <button 
                onClick={() => {
                  if (!newEqData.id || !newEqData.name) {
                    alert('Vui lòng nhập mã và tên thiết bị');
                    return;
                  }
                  setAllEquipment([...allEquipment, {
                    ...newEqData,
                    status: 'healthy',
                    health: 100,
                    lastCheck: new Date().toLocaleDateString('vi-VN'),
                    lat: 21.0285,
                    lng: 105.8542
                  }]);
                  setShowAddEqModal(false);
                  setNewEqData({ id: '', name: '', customer: '', factory: '', location: '', type: 'Máy biến áp' });
                  alert('Đã thêm thiết bị thành công!');
                }}
                className="px-4 py-2 bg-blue-600 text-white font-medium rounded-lg hover:bg-blue-700 transition-colors shadow-sm"
              >
                Lưu thiết bị
              </button>
            </div>
          </div>
        </div>
      )}

      {showHistoryModal && (
        <div className="fixed inset-0 bg-slate-900/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-2xl overflow-hidden">
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
              <div className="flex items-center gap-2">
                <BarChart2 className="text-blue-600" size={20} />
                <h2 className="text-lg font-bold text-slate-900">Lịch sử thông số: {selectedParamHistory}</h2>
              </div>
              <button onClick={() => setShowHistoryModal(false)} className="p-2 text-slate-400 hover:text-slate-600 hover:bg-slate-200 rounded-full transition-colors">
                <X size={20} />
              </button>
            </div>
            
            <div className="p-6">
              {dynamicParamHistoryData ? (
                <div className="h-64 w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={dynamicParamHistoryData} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                      <XAxis dataKey="date" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dy={10} />
                      <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dx={-10} />
                      <Tooltip 
                        contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                        labelStyle={{ color: '#64748b', marginBottom: '4px' }}
                      />
                      <Line type="monotone" dataKey="value" stroke="#3b82f6" strokeWidth={3} activeDot={{ r: 6, strokeWidth: 0, fill: '#3b82f6' }} />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              ) : (
                <div className="py-12 text-center text-slate-500">
                  <History size={48} className="mx-auto text-slate-300 mb-4" />
                  <p>Không có dữ liệu lịch sử dạng biểu đồ cho thông số này.</p>
                  <p className="text-sm mt-1">Vui lòng xem trong các báo cáo cũ.</p>
                </div>
              )}
            </div>
            
            <div className="px-6 py-4 border-t border-slate-200 bg-slate-50 flex justify-end">
              <button 
                onClick={() => setShowHistoryModal(false)}
                className="px-4 py-2 bg-white border border-slate-300 text-slate-700 rounded-lg font-medium hover:bg-slate-50 transition-colors"
              >
                Đóng
              </button>
            </div>
          </div>
        </div>
      )}

      {showWorkOrderModal && (
        <div className="fixed inset-0 bg-slate-900/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-2xl overflow-hidden">
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
              <h3 className="font-bold text-slate-800 text-lg">
                {selectedWorkOrder ? 'Chi tiết Phiếu công việc' : 'Tạo Phiếu công việc mới'}
              </h3>
              <button onClick={() => setShowWorkOrderModal(false)} className="text-slate-400 hover:text-slate-600 transition-colors">
                <X size={20} />
              </button>
            </div>
            <div className="p-6 overflow-y-auto max-h-[70vh]">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="md:col-span-2">
                  <label className="block text-sm font-medium text-slate-700 mb-1">Tiêu đề công việc</label>
                  <input 
                    type="text" 
                    value={newWorkOrder.title} 
                    onChange={(e) => setNewWorkOrder({...newWorkOrder, title: e.target.value})} 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                    placeholder="VD: Kiểm tra định kỳ Inverter 01" 
                  />
                </div>
                <div className="md:col-span-2">
                  <label className="block text-sm font-medium text-slate-700 mb-1">Mô tả chi tiết</label>
                  <textarea 
                    value={newWorkOrder.description} 
                    onChange={(e) => setNewWorkOrder({...newWorkOrder, description: e.target.value})} 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 h-24" 
                    placeholder="Mô tả các bước thực hiện..."
                  />
                </div>
                <div className="md:col-span-2 relative">
                  <label className="block text-sm font-medium text-slate-700 mb-1">Mã thiết bị (Asset ID) - Có thể chọn nhiều</label>
                  <div className="min-h-[42px] p-1.5 border border-slate-300 rounded-lg focus-within:ring-2 focus-within:ring-blue-500 bg-white flex flex-wrap gap-2">
                    {newWorkOrder.equipmentId.map(id => (
                      <div key={id} className="flex items-center gap-1 px-2 py-1 bg-blue-50 text-blue-700 rounded-md text-sm border border-blue-100">
                        <span>{id}</span>
                        <button 
                          onClick={() => setNewWorkOrder({
                            ...newWorkOrder, 
                            equipmentId: newWorkOrder.equipmentId.filter(item => item !== id)
                          })}
                          className="hover:text-blue-900"
                        >
                          <X size={14} />
                        </button>
                      </div>
                    ))}
                    <input 
                      type="text" 
                      value={equipmentSearch}
                      onFocus={() => setShowEqDropdown(true)}
                      onChange={(e) => {
                        setEquipmentSearch(e.target.value);
                        setShowEqDropdown(true);
                      }}
                      className="flex-1 min-w-[120px] outline-none text-sm py-1"
                      placeholder={newWorkOrder.equipmentId.length === 0 ? "Tìm và chọn thiết bị..." : ""}
                    />
                  </div>
                  
                  {showEqDropdown && (
                    <div className="absolute z-50 left-0 right-0 mt-1 bg-white border border-slate-200 rounded-lg shadow-lg max-h-60 overflow-y-auto">
                      <div className="p-2 border-b border-slate-100 sticky top-0 bg-white flex justify-between items-center">
                        <span className="text-xs font-bold text-slate-400 uppercase">Danh sách thiết bị</span>
                        <button onClick={() => setShowEqDropdown(false)} className="text-slate-400 hover:text-slate-600">
                          <X size={14} />
                        </button>
                      </div>
                      {allEquipment
                        .filter(eq => 
                          !newWorkOrder.equipmentId.includes(eq.id) && 
                          (eq.id.toLowerCase().includes(equipmentSearch.toLowerCase()) || 
                           eq.name.toLowerCase().includes(equipmentSearch.toLowerCase()))
                        )
                        .slice(0, 50)
                        .map(eq => (
                          <button
                            key={eq.id}
                            onClick={() => {
                              setNewWorkOrder({
                                ...newWorkOrder,
                                equipmentId: [...newWorkOrder.equipmentId, eq.id]
                              });
                              setEquipmentSearch('');
                              setShowEqDropdown(false);
                            }}
                            className="w-full text-left px-4 py-2 hover:bg-slate-50 flex flex-col transition-colors border-b border-slate-50 last:border-0"
                          >
                            <span className="font-bold text-slate-800 text-sm">{eq.id}</span>
                            <span className="text-xs text-slate-500">{eq.name} - {eq.factory}</span>
                          </button>
                        ))}
                      {allEquipment.filter(eq => 
                        !newWorkOrder.equipmentId.includes(eq.id) && 
                        (eq.id.toLowerCase().includes(equipmentSearch.toLowerCase()) || 
                         eq.name.toLowerCase().includes(equipmentSearch.toLowerCase()))
                      ).length === 0 && (
                        <div className="p-4 text-center text-slate-400 text-sm">Không tìm thấy thiết bị nào</div>
                      )}
                    </div>
                  )}
                </div>
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Khách hàng</label>
                  <select 
                    value={newWorkOrder.customerId} 
                    onChange={(e) => {
                      const custId = e.target.value;
                      setNewWorkOrder({...newWorkOrder, customerId: custId, factory: ''});
                    }} 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                  >
                    <option value="">-- Chọn khách hàng --</option>
                    {customers.map(customer => (
                      <option key={customer.id} value={customer.id}>{customer.name}</option>
                    ))}
                    {customers.length === 0 && !isCustomersLoading && (
                      <option disabled>Chưa có dữ liệu khách hàng</option>
                    )}
                  </select>
                </div>
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Nhà máy</label>
                  {newWorkOrder.customerId && customers.find(c => c.id === newWorkOrder.customerId)?.factories?.length > 0 ? (
                    <select 
                      value={newWorkOrder.factory} 
                      onChange={(e) => setNewWorkOrder({...newWorkOrder, factory: e.target.value})} 
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      <option value="">-- Chọn nhà máy --</option>
                      {customers.find(c => c.id === newWorkOrder.customerId).factories.map((f, idx) => (
                        <option key={idx} value={f}>{f}</option>
                      ))}
                    </select>
                  ) : (
                    <input 
                      type="text" 
                      value={newWorkOrder.factory} 
                      onChange={(e) => setNewWorkOrder({...newWorkOrder, factory: e.target.value})} 
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                      placeholder="VD: Nhà máy Bắc Ninh" 
                    />
                  )}
                </div>
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Loại công việc</label>
                  <select 
                    value={newWorkOrder.type} 
                    onChange={(e) => setNewWorkOrder({...newWorkOrder, type: e.target.value, isUnplanned: e.target.value === 'unplanned'})} 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                  >
                    <option value="preventive">Bảo trì định kỳ (PM)</option>
                    <option value="corrective">Sửa chữa khắc phục (CM)</option>
                    <option value="inspection">Kiểm tra hiện trường</option>
                    <option value="emergency">Xử lý khẩn cấp</option>
                    <option value="unplanned">Ngoài kế hoạch (Unplanned)</option>
                  </select>
                </div>
                {newWorkOrder.type === 'preventive' && (
                  <div>
                    <label className="block text-sm font-medium text-slate-700 mb-1">Tần suất PM</label>
                    <select 
                      value={newWorkOrder.pmFrequency} 
                      onChange={(e) => setNewWorkOrder({...newWorkOrder, pmFrequency: e.target.value})} 
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      <option value="">-- Chọn tần suất --</option>
                      <option value="daily">Hàng ngày (Daily)</option>
                      <option value="weekly">Hàng tuần (Weekly)</option>
                      <option value="bi-monthly">Nửa tháng (Bi-monthly)</option>
                      <option value="monthly">Hàng tháng (Monthly)</option>
                      <option value="3-months">3 tháng (3 months)</option>
                      <option value="6-months">6 tháng (6 months)</option>
                      <option value="1-year">1 năm (1 year)</option>
                      <option value="3-years">3 năm (3 years)</option>
                      <option value="6-years">6 năm (6 years)</option>
                    </select>
                  </div>
                )}
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Mã lỗi sự cố (Failure Code)</label>
                  <select 
                    value={newWorkOrder.failureCode} 
                    onChange={(e) => setNewWorkOrder({...newWorkOrder, failureCode: e.target.value})} 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                  >
                    <option value="">-- Chọn mã lỗi --</option>
                    <option value="connection error">Connection error</option>
                    <option value="control failure">Control failure</option>
                    <option value="corrosion">Corrosion</option>
                    <option value="crack">Crack</option>
                    <option value="deformation">Deformation</option>
                    <option value="display error">Display error</option>
                    <option value="leakage">Leakage</option>
                    <option value="looseen">Looseen</option>
                    <option value="low performanace">Low performanace</option>
                    <option value="noise">Noise</option>
                    <option value="overheat">Overheat</option>
                    <option value="power failure">Power failure</option>
                    <option value="shorted circuit">Shorted circuit</option>
                    <option value="vibration">Vibration</option>
                    <option value="worn out">Worn out</option>
                    <option value="jammed stuck">Jammed stuck</option>
                    <option value="other">Khác...</option>
                  </select>
                  {newWorkOrder.failureCode === 'other' && (
                    <input 
                      type="text" 
                      onChange={(e) => setNewWorkOrder({...newWorkOrder, failureCode: e.target.value})} 
                      className="w-full px-3 py-2 mt-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                      placeholder="Nhập mã lỗi khác" 
                    />
                  )}
                </div>
                {newWorkOrder.status === 'completed' && (
                  <>
                    <div>
                      <label className="block text-sm font-medium text-slate-700 mb-1">Nguyên nhân gốc rễ (Root Cause)</label>
                      <textarea 
                        value={newWorkOrder.rootCause} 
                        onChange={(e) => setNewWorkOrder({...newWorkOrder, rootCause: e.target.value})} 
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                        placeholder="Phân tích nguyên nhân gây ra sự cố..."
                        rows={2}
                      />
                    </div>
                    <div>
                      <label className="block text-sm font-medium text-slate-700 mb-1">Thời gian thực tế (Giờ)</label>
                      <input 
                        type="number" 
                        value={newWorkOrder.actualTimeSpent} 
                        onChange={(e) => setNewWorkOrder({...newWorkOrder, actualTimeSpent: parseFloat(e.target.value) || 0})} 
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                        placeholder="VD: 3.5" 
                      />
                    </div>
                  </>
                )}
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Mức độ ưu tiên</label>
                  <select 
                    value={newWorkOrder.priority} 
                    onChange={(e) => setNewWorkOrder({...newWorkOrder, priority: e.target.value})} 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                  >
                    <option value="low">Thấp</option>
                    <option value="medium">Trung bình</option>
                    <option value="high">Cao</option>
                    <option value="urgent">Khẩn cấp (P1)</option>
                  </select>
                </div>
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Trạng thái (Phase)</label>
                  <select 
                    value={newWorkOrder.status} 
                    onChange={(e) => setNewWorkOrder({...newWorkOrder, status: e.target.value})} 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                  >
                    <option value="initiated">1. Khởi tạo (Initiated/WR)</option>
                    <option value="approved">2. Đã duyệt (Approved)</option>
                    <option value="planned">3. Đã lập kế hoạch (Planned)</option>
                    <option value="scheduled">4. Đã lập lịch (Scheduled)</option>
                    <option value="in-progress">5. Đang thực hiện (In Progress)</option>
                    <option value="completed">6. Hoàn thành (Completed)</option>
                    <option value="overdue">Quá hạn (Overdue)</option>
                    <option value="cancelled">Đã hủy</option>
                  </select>
                </div>

                {/* Timestamps Section */}
                <div className="md:col-span-2 grid grid-cols-2 md:grid-cols-4 gap-4 bg-slate-50 p-3 rounded-lg border border-slate-200">
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Bắt đầu dừng máy</label>
                    <div className="text-xs font-medium text-slate-600">
                      {newWorkOrder.downtimeStart ? new Date(newWorkOrder.downtimeStart).toLocaleString('vi-VN') : '---'}
                    </div>
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Bắt đầu sửa chữa</label>
                    <div className="text-xs font-medium text-slate-600">
                      {newWorkOrder.repairStart ? new Date(newWorkOrder.repairStart).toLocaleString('vi-VN') : '---'}
                    </div>
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Kết thúc sửa chữa</label>
                    <div className="text-xs font-medium text-slate-600">
                      {newWorkOrder.repairEnd ? new Date(newWorkOrder.repairEnd).toLocaleString('vi-VN') : '---'}
                    </div>
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Chạy lại máy</label>
                    <div className="text-xs font-medium text-slate-600">
                      {newWorkOrder.restartTime ? new Date(newWorkOrder.restartTime).toLocaleString('vi-VN') : '---'}
                    </div>
                  </div>
                </div>
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Người thực hiện / PIC</label>
                  <select 
                    value={newWorkOrder.assignedTo} 
                    onChange={(e) => setNewWorkOrder({...newWorkOrder, assignedTo: e.target.value})} 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                  >
                    <option value="">-- Chọn người thực hiện --</option>
                    <option value="Nguyen Van A">Nguyen Van A</option>
                    <option value="Tran Van B">Tran Van B</option>
                    <option value="Le Thi C">Le Thi C</option>
                    <option value="Pham Van D">Pham Van D</option>
                    <option value="Vu Van E">Vu Van E</option>
                    <option value="Khác">Khác...</option>
                  </select>
                  {newWorkOrder.assignedTo === 'Khác' && (
                    <input 
                      type="text" 
                      onChange={(e) => setNewWorkOrder({...newWorkOrder, assignedTo: e.target.value})} 
                      className="w-full px-3 py-2 mt-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                      placeholder="Nhập tên người thực hiện" 
                    />
                  )}
                </div>
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Vai trò Phê duyệt (Approve)</label>
                  <select 
                    value={newWorkOrder.responsibleApprove} 
                    onChange={(e) => setNewWorkOrder({...newWorkOrder, responsibleApprove: e.target.value})} 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                  >
                    <option value="TM">TM (Technical Manager)</option>
                    <option value="PE">PE (Plant Engineer)</option>
                    <option value="MS">MS (Maintenance Supervisor)</option>
                    <option value="OSL">OSL (Operation Shift Leader)</option>
                    <option value="SO">SO (Safety Officer)</option>
                    <option value="DTM">DTM (Deputy Technical Manager)</option>
                  </select>
                </div>
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Vai trò Thực hiện (Do)</label>
                  <select 
                    value={newWorkOrder.responsibleDo} 
                    onChange={(e) => setNewWorkOrder({...newWorkOrder, responsibleDo: e.target.value})} 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                  >
                    <option value="Everybody">Everybody</option>
                    <option value="ME">ME (Maintenance Engineer)</option>
                    <option value="PE">PE (Plant Engineer)</option>
                    <option value="OSL">OSL (Operation Shift Leader)</option>
                    <option value="OSE">OSE (Operation Shift Engineer)</option>
                    <option value="Subcontractor">Subcontractor</option>
                  </select>
                </div>
                <div className="flex items-center gap-4 mt-2">
                  <label className="flex items-center gap-2 cursor-pointer">
                    <input 
                      type="checkbox" 
                      checked={newWorkOrder.blockingRequired} 
                      onChange={(e) => setNewWorkOrder({...newWorkOrder, blockingRequired: e.target.checked})}
                      className="w-4 h-4 text-blue-600 rounded"
                    />
                    <span className="text-sm font-medium text-slate-700">Yêu cầu Cô lập (Blocking)</span>
                  </label>
                </div>
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Mã Work Permit (nếu có)</label>
                  <input 
                    type="text" 
                    value={newWorkOrder.workPermitId} 
                    onChange={(e) => setNewWorkOrder({...newWorkOrder, workPermitId: e.target.value})} 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                    placeholder="VD: WP-2024-001" 
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Hạn hoàn thành</label>
                  <input 
                    type="date" 
                    value={newWorkOrder.dueDate} 
                    onChange={(e) => setNewWorkOrder({...newWorkOrder, dueDate: e.target.value})} 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                  />
                </div>
                
                <div className="md:col-span-2 border-t border-slate-100 pt-4 mt-2">
                  <h4 className="text-sm font-bold text-slate-700 mb-4">Nguồn lực & Chi phí</h4>
                  <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4">
                    <div>
                      <label className="block text-xs font-medium text-slate-700 mb-1">Thời gian (Giờ)</label>
                      <input 
                        type="number" 
                        value={newWorkOrder.estimatedTime} 
                        onChange={(e) => setNewWorkOrder({...newWorkOrder, estimatedTime: parseFloat(e.target.value) || 0})} 
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm" 
                        placeholder="VD: 2.5"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-slate-700 mb-1">Số người (Labor)</label>
                      <input 
                        type="number" 
                        value={newWorkOrder.laborCount} 
                        onChange={(e) => setNewWorkOrder({...newWorkOrder, laborCount: parseInt(e.target.value) || 0})} 
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm" 
                        placeholder="VD: 2"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-slate-700 mb-1">Chi phí nhân công ($)</label>
                      <input 
                        type="number" 
                        value={newWorkOrder.laborCost} 
                        onChange={(e) => setNewWorkOrder({...newWorkOrder, laborCost: parseFloat(e.target.value) || 0})} 
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm" 
                        placeholder="VD: 150"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-slate-700 mb-1">Chi phí vật tư ($)</label>
                      <input 
                        type="number" 
                        value={newWorkOrder.partCost} 
                        onChange={(e) => setNewWorkOrder({...newWorkOrder, partCost: parseFloat(e.target.value) || 0})} 
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm" 
                        placeholder="VD: 500"
                      />
                    </div>
                  </div>
                </div>

                <div className="md:col-span-2 border-t border-slate-100 pt-4 mt-2">
                  <div className="flex items-center justify-between mb-2">
                    <label className="block text-sm font-bold text-slate-700">Đính kèm tài liệu / Hình ảnh</label>
                  </div>
                  <div className="flex items-center gap-4">
                    <label className="flex items-center justify-center w-full h-24 px-4 transition bg-white border-2 border-slate-300 border-dashed rounded-xl appearance-none cursor-pointer hover:border-blue-400 focus:outline-none">
                      <span className="flex items-center space-x-2">
                        <svg xmlns="http://www.w3.org/2000/svg" className="w-6 h-6 text-slate-400" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2">
                          <path strokeLinecap="round" strokeLinejoin="round" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                        </svg>
                        <span className="font-medium text-slate-600">
                          Nhấn để tải lên hoặc kéo thả file
                        </span>
                      </span>
                      <input 
                        type="file" 
                        name="file_upload" 
                        className="hidden" 
                        accept="image/*,.pdf,.doc,.docx"
                        multiple
                        onChange={(e) => {
                          if (e.target.files && e.target.files.length > 0) {
                            const files = Array.from(e.target.files) as File[];
                            const newAttachments = files.map(f => URL.createObjectURL(f));
                            setNewWorkOrder(prev => ({
                              ...prev,
                              attachments: [...(prev.attachments || []), ...newAttachments]
                            }));
                          }
                        }}
                      />
                    </label>
                  </div>
                  {newWorkOrder.attachments && newWorkOrder.attachments.length > 0 && (
                    <div className="flex flex-wrap gap-2 mt-3">
                      {newWorkOrder.attachments.map((url, idx) => (
                        <div key={idx} className="relative w-16 h-16 rounded-lg overflow-hidden border border-slate-200">
                          {url.startsWith('blob:') ? (
                            <img src={url} alt="Attachment" className="w-full h-full object-cover" />
                          ) : (
                            <div className="w-full h-full flex items-center justify-center bg-slate-100 text-xs text-slate-500">File</div>
                          )}
                          <button 
                            className="absolute top-0 right-0 bg-red-500 text-white p-0.5 rounded-bl-lg"
                            onClick={() => {
                              const newAtt = [...newWorkOrder.attachments];
                              newAtt.splice(idx, 1);
                              setNewWorkOrder({...newWorkOrder, attachments: newAtt});
                            }}
                          >
                            <X size={12} />
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>

                <div className="md:col-span-2 border-t border-slate-100 pt-4 mt-2">
                  <div className="flex items-center justify-between mb-2">
                    <label className="block text-sm font-bold text-slate-700">Vật tư sử dụng</label>
                  </div>
                  <div className="flex gap-2 mb-4">
                    <input 
                      type="text" 
                      id="newMaterialName"
                      placeholder="Tên vật tư" 
                      className="flex-1 px-3 py-1.5 border border-slate-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                    />
                    <input 
                      type="number" 
                      id="newMaterialQty"
                      placeholder="SL" 
                      defaultValue="1"
                      className="w-20 px-3 py-1.5 border border-slate-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                    />
                    <button 
                      onClick={() => {
                        const nameInput = document.getElementById('newMaterialName') as HTMLInputElement;
                        const qtyInput = document.getElementById('newMaterialQty') as HTMLInputElement;
                        if (nameInput && qtyInput && nameInput.value) {
                          setNewWorkOrder({
                            ...newWorkOrder,
                            usedMaterials: [...(newWorkOrder.usedMaterials || []), { 
                              name: nameInput.value, 
                              quantity: parseFloat(qtyInput.value) || 1, 
                              unit: 'pcs' 
                            }]
                          });
                          nameInput.value = '';
                          qtyInput.value = '1';
                        }
                      }}
                      className="px-3 py-1.5 bg-slate-100 text-blue-600 font-medium rounded-lg hover:bg-blue-50 border border-slate-200 transition-colors flex items-center gap-1 text-sm"
                    >
                      <Plus size={14} /> Thêm
                    </button>
                  </div>
                  <div className="space-y-2">
                    {(newWorkOrder.usedMaterials || []).map((m, idx) => (
                      <div key={idx} className="flex items-center justify-between bg-slate-50 p-2 rounded-lg text-sm">
                        <span>{m.name} - {m.quantity} {m.unit}</span>
                        <button 
                          onClick={() => {
                            const updated = [...newWorkOrder.usedMaterials];
                            updated.splice(idx, 1);
                            setNewWorkOrder({...newWorkOrder, usedMaterials: updated});
                          }}
                          className="text-rose-500 hover:text-rose-700"
                        >
                          <X size={14} />
                        </button>
                      </div>
                    ))}
                    {(newWorkOrder.usedMaterials || []).length === 0 && (
                      <p className="text-xs text-slate-400 italic">Chưa có vật tư nào được thêm.</p>
                    )}
                  </div>
                </div>
              </div>
            </div>
            <div className="px-6 py-4 border-t border-slate-200 bg-slate-50 flex flex-col gap-3">
              {woError && (
                <div className="text-sm text-rose-600 font-medium bg-rose-50 p-2 rounded-lg border border-rose-100">
                  {woError}
                </div>
              )}
              <div className="flex justify-end gap-3">
                <button onClick={() => setShowWorkOrderModal(false)} className="px-4 py-2 text-slate-600 font-medium hover:bg-slate-200 rounded-lg transition-colors">Hủy</button>
                <button 
                  onClick={handleSaveWorkOrder}
                  className="px-4 py-2 bg-blue-600 text-white font-medium rounded-lg hover:bg-blue-700 transition-colors shadow-sm"
                >
                  {selectedWorkOrder ? 'Cập nhật phiếu' : 'Lưu phiếu mới'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {confirmDialog && (
        <div className="fixed inset-0 bg-slate-900/50 z-[60] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-sm overflow-hidden">
            <div className="p-6">
              <h3 className="font-bold text-slate-800 text-lg mb-2">Xác nhận</h3>
              <p className="text-slate-600">{confirmDialog.message}</p>
            </div>
            <div className="px-6 py-4 border-t border-slate-200 bg-slate-50 flex justify-end gap-3">
              <button 
                onClick={() => setConfirmDialog(null)} 
                className="px-4 py-2 text-slate-600 font-medium hover:bg-slate-200 rounded-lg transition-colors"
              >
                Hủy
              </button>
              <button 
                onClick={confirmDialog.onConfirm}
                className="px-4 py-2 bg-rose-600 text-white font-medium rounded-lg hover:bg-rose-700 transition-colors shadow-sm"
              >
                Đồng ý
              </button>
            </div>
          </div>
        </div>
      )}

      {showInventoryModal && (
        <div className="fixed inset-0 bg-slate-900/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-md overflow-hidden">
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
              <h3 className="font-bold text-slate-800 text-lg">
                {selectedInventoryItem ? 'Cập nhật vật tư' : 'Thêm vật tư mới'}
              </h3>
              <button onClick={() => setShowInventoryModal(false)} className="text-slate-400 hover:text-slate-600 transition-colors">
                <X size={20} />
              </button>
            </div>
            <div className="p-6 space-y-4">
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Tên vật tư</label>
                <input 
                  type="text" 
                  value={newInventoryItem.name} 
                  onChange={(e) => setNewInventoryItem({...newInventoryItem, name: e.target.value})} 
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                  placeholder="VD: Cầu chì 10A" 
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Mã SKU</label>
                <input 
                  type="text" 
                  value={newInventoryItem.sku} 
                  onChange={(e) => setNewInventoryItem({...newInventoryItem, sku: e.target.value})} 
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                  placeholder="VD: FUSE-10A-001" 
                />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Số lượng</label>
                  <input 
                    type="number" 
                    value={newInventoryItem.quantity} 
                    onChange={(e) => setNewInventoryItem({...newInventoryItem, quantity: parseInt(e.target.value) || 0})} 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Đơn vị</label>
                  <input 
                    type="text" 
                    value={newInventoryItem.unit} 
                    onChange={(e) => setNewInventoryItem({...newInventoryItem, unit: e.target.value})} 
                    className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                    placeholder="pcs, m, l..." 
                  />
                </div>
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Vị trí kho</label>
                <input 
                  type="text" 
                  value={newInventoryItem.location} 
                  onChange={(e) => setNewInventoryItem({...newInventoryItem, location: e.target.value})} 
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                  placeholder="VD: Kệ A1, Ngăn 2" 
                />
              </div>
            </div>
            <div className="px-6 py-4 border-t border-slate-200 bg-slate-50 flex justify-end gap-3">
              <button onClick={() => setShowInventoryModal(false)} className="px-4 py-2 text-slate-600 font-medium hover:bg-slate-200 rounded-lg transition-colors">Hủy</button>
              <button 
                onClick={handleSaveInventoryItem}
                className="px-4 py-2 bg-blue-600 text-white font-medium rounded-lg hover:bg-blue-700 transition-colors shadow-sm"
              >
                Lưu vật tư
              </button>
            </div>
          </div>
        </div>
      )}

      {/* --- CUSTOMER MODAL --- */}
      {showCustomerModal && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-md overflow-hidden">
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
              <h3 className="font-bold text-slate-800 text-lg">
                {editingCustomer ? 'Cập nhật khách hàng' : 'Thêm khách hàng mới'}
              </h3>
              <button onClick={() => setShowCustomerModal(false)} className="text-slate-400 hover:text-slate-600 transition-colors">
                <X size={20} />
              </button>
            </div>
            <div className="p-6 space-y-4">
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Mã khách hàng</label>
                <input 
                  type="text" 
                  value={editingCustomer ? editingCustomer.id : generateNextCustomerId()} 
                  disabled
                  className="w-full px-3 py-2 border border-slate-200 rounded-lg bg-slate-50 text-slate-500 cursor-not-allowed" 
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Tên khách hàng</label>
                <input 
                  type="text" 
                  value={newCustomer.name} 
                  onChange={(e) => setNewCustomer({...newCustomer, name: e.target.value})} 
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                  placeholder="VD: Công ty Điện lực A" 
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Danh sách nhà máy</label>
                <div className="flex gap-2 mb-2">
                  <input 
                    type="text" 
                    id="new-factory-input"
                    className="flex-1 px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                    placeholder="Nhập tên nhà máy..." 
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        const input = e.currentTarget;
                        if (input.value.trim()) {
                          setNewCustomer({...newCustomer, factories: [...(newCustomer.factories || []), input.value.trim()]});
                          input.value = '';
                        }
                      }
                    }}
                  />
                  <button 
                    onClick={() => {
                      const input = document.getElementById('new-factory-input') as HTMLInputElement;
                      if (input.value.trim()) {
                        setNewCustomer({...newCustomer, factories: [...(newCustomer.factories || []), input.value.trim()]});
                        input.value = '';
                      }
                    }}
                    className="px-3 py-2 bg-slate-100 text-slate-600 rounded-lg hover:bg-slate-200"
                  >
                    Thêm
                  </button>
                </div>
                <div className="flex flex-wrap gap-2">
                  {(newCustomer.factories || []).map((f, i) => (
                    <span key={i} className="flex items-center gap-1 px-2 py-1 bg-blue-50 text-blue-700 rounded-md text-xs border border-blue-100">
                      {f}
                      <button onClick={() => setNewCustomer({...newCustomer, factories: newCustomer.factories.filter((_, idx) => idx !== i)})}>
                        <X size={12} />
                      </button>
                    </span>
                  ))}
                </div>
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Email</label>
                <input 
                  type="email" 
                  value={newCustomer.email} 
                  onChange={(e) => setNewCustomer({...newCustomer, email: e.target.value})} 
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                  placeholder="VD: contact@company.com" 
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Số điện thoại</label>
                <input 
                  type="text" 
                  value={newCustomer.phone} 
                  onChange={(e) => setNewCustomer({...newCustomer, phone: e.target.value})} 
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                  placeholder="VD: 0901234567" 
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Địa chỉ</label>
                <textarea 
                  value={newCustomer.address} 
                  onChange={(e) => setNewCustomer({...newCustomer, address: e.target.value})} 
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" 
                  placeholder="VD: Số 1, Đường A, TP. HCM" 
                  rows={3}
                />
              </div>
            </div>
            <div className="px-6 py-4 border-t border-slate-200 bg-slate-50 flex justify-end gap-3">
              <button onClick={() => setShowCustomerModal(false)} className="px-4 py-2 text-slate-600 font-medium hover:bg-slate-200 rounded-lg transition-colors">Hủy</button>
              <button 
                onClick={handleSaveCustomer}
                className="px-4 py-2 bg-blue-600 text-white font-medium rounded-lg hover:bg-blue-700 transition-colors shadow-sm"
              >
                Lưu khách hàng
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal Đối soát & Hòa giải phiên bản dữ liệu (Sync Reconciliation Modal) */}
      {showReconciliationModal && (
        <SyncReconciliationModal
          isOpen={showReconciliationModal}
          onClose={() => setShowReconciliationModal(false)}
          report={reconciliationReport}
          isReconciling={isReconciling}
          onResolveNewerWins={handleResolveNewerWins}
          onForcePushToSheets={handleForcePushToSheets}
          onPullFromSheets={handlePullFromSheetsToFirestore}
          onRefresh={() => checkDataVersionReconciliation({ silent: true })}
        />
      )}

    </div>
  );
}

