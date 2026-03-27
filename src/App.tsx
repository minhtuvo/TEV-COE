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
  MoreVertical,
  TrendingUp,
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
  Eye,
  Calendar,
  PieChart as PieChartIcon,
  BarChart as BarChartIcon,
  Wind,
  Droplets,
  Gauge,
  RefreshCw,
  LogOut,
  User,
  Paperclip,
  X
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
import { MapContainer, TileLayer, Marker, Popup, Tooltip as LeafletTooltip } from 'react-leaflet';
import L from 'leaflet';
import jsPDF from 'jspdf';
import * as htmlToImage from 'html-to-image';
import { auth, db, signInWithGoogle, logOut } from './firebase';
import { onAuthStateChanged } from 'firebase/auth';
import { doc, getDoc, setDoc } from 'firebase/firestore';
import { Html5QrcodeScanner } from 'html5-qrcode';

// --- MOCK DATA ---
const kpiData = {
  total: 124,
  healthy: 105,
  warning: 14,
  critical: 5
};

const topRiskEquipment = [
  { id: 'TRF-01', name: 'Máy biến áp T1', factory: 'Nhà máy Bắc Ninh', location: 'Trạm biến áp 110kV', health: 45, status: 'critical', trend: 'down' },
  { id: 'MTR-05', name: 'Động cơ bơm làm mát', factory: 'Nhà máy Bình Dương', location: 'Khu vực Bơm', health: 52, status: 'critical', trend: 'down' },
  { id: 'GEN-02', name: 'Máy phát điện dự phòng', factory: 'Nhà máy Bắc Ninh', location: 'Phòng Máy Phát', health: 68, status: 'warning', trend: 'stable' },
  { id: 'TRF-03', name: 'Máy biến áp T3', factory: 'Nhà máy Đồng Nai', location: 'Trạm biến áp 22kV', health: 71, status: 'warning', trend: 'down' },
];

const detailedRiskData: Record<string, any> = {
  'TRF-01': {
    id: 'TRF-01',
    name: 'Máy biến áp T1',
    type: 'Máy biến áp',
    health: 45,
    status: 'critical',
    factory: 'Nhà máy Bắc Ninh',
    location: 'Trạm biến áp 110kV',
    generalRecommendation: 'Cảnh báo mức độ 4. Cần tiến hành lọc dầu hoặc thay dầu ngay lập tức. Lên kế hoạch cắt điện để kiểm tra tổng thể phần ruột máy do hàm lượng khí cháy và Furan tăng cao đột biến.',
    failureProfile: [
      { subject: 'Mag Circuit', score: 20 },
      { subject: 'Winding', score: 80 },
      { subject: 'Main Tank', score: 40 },
      { subject: 'Cooling', score: 10 },
      { subject: 'Liquid', score: 90 },
      { subject: 'Bushing', score: 30 },
      { subject: 'LTC', score: 50 },
    ],
    defects: [
      { name: 'Cooling Controls Failure', warnings: 25.0, risk: 75.0 },
      { name: 'Liquid Insulation Arcing', warnings: 12.0, risk: 50.0 },
      { name: 'Loose Connection', warnings: 10.0, risk: 42.0 },
      { name: 'Loss Of Core Ground', warnings: 20.0, risk: 80.0 },
      { name: 'Moisture Contamination', warnings: 10.0, risk: 42.0 },
      { name: 'Overheating', warnings: 2.0, risk: 5.0 },
    ],
    tests: [
      { name: 'Phân tích khí hòa tan (DGA - TDCG)', shortName: 'DGA', value: '5200', unit: 'ppm', score: 2, weight: 3, status: 'Critical', rec: 'Khí cháy hòa tan vượt ngưỡng nguy hiểm. Có hiện tượng phóng điện hoặc quá nhiệt cục bộ bên trong.' },
      { name: 'Độ bền điện môi của dầu', shortName: 'Điện môi', value: '35', unit: 'kV', score: 3, weight: 2, status: 'Critical', rec: 'Dầu suy giảm chất lượng nghiêm trọng, cần lọc hoặc thay mới.' },
      { name: 'Điện trở cách điện', shortName: 'Cách điện', value: '800', unit: 'MΩ', score: 5, weight: 1, status: 'Moderate', rec: 'Cách điện suy giảm, cần sấy lại.' },
      { name: 'Điện trở cuộn dây', shortName: 'Điện trở', value: '4.2', unit: '% lệch', score: 5, weight: 1, status: 'Moderate', rec: 'Kiểm tra lại các đầu nối và bộ điều áp dưới tải (OLTC).' },
      { name: 'Hàm lượng Furan', shortName: 'Furan', value: '6.5', unit: 'mg/kg', score: 2, weight: 2, status: 'Critical', rec: 'Giấy cách điện lão hóa nặng, tuổi thọ còn lại rất thấp.' },
    ]
  },
  'MTR-05': {
    id: 'MTR-05',
    name: 'Động cơ bơm làm mát',
    type: 'Động cơ',
    health: 52,
    status: 'critical',
    factory: 'Nhà máy Bình Dương',
    location: 'Khu vực Bơm',
    generalRecommendation: 'Cần lên kế hoạch bảo dưỡng khẩn cấp. Theo dõi sát sao hiện tượng phóng điện cục bộ và lên phương án quấn lại cuộn dây hoặc thay thế stator nếu tình trạng xấu đi.',
    failureProfile: [
      { subject: 'Stator', score: 85 },
      { subject: 'Rotor', score: 40 },
      { subject: 'Bearings', score: 60 },
      { subject: 'Cooling', score: 30 },
      { subject: 'Insulation', score: 90 },
      { subject: 'Vibration', score: 70 },
      { subject: 'Alignment', score: 20 },
    ],
    defects: [
      { name: 'Insulation Breakdown', warnings: 30.0, risk: 85.0 },
      { name: 'Bearing Wear', warnings: 15.0, risk: 60.0 },
      { name: 'Rotor Bar Damage', warnings: 5.0, risk: 40.0 },
      { name: 'Overheating', warnings: 20.0, risk: 70.0 },
      { name: 'Misalignment', warnings: 8.0, risk: 20.0 },
    ],
    tests: [
      { name: 'Tan-delta', shortName: 'Tan-delta', value: '0.08', unit: '', score: 4, weight: 2, status: 'Moderate', rec: 'Theo dõi định kỳ, lớp cách điện có dấu hiệu lão hóa.' },
      { name: 'Tan-delta Tip-up', shortName: 'Tip-up', value: '0.008', unit: '', score: 2, weight: 2, status: 'Critical', rec: 'Kiểm tra ngay, có lỗ hổng trong lớp cách điện.' },
      { name: 'Partial Discharge (PD)', shortName: 'PD', value: '22000', unit: 'pC', score: 2, weight: 3, status: 'Critical', rec: 'Nguy hiểm! Phóng điện cục bộ rất cao, nguy cơ đánh thủng cách điện.' },
      { name: 'Insulation Resistance', shortName: 'IR', value: '15', unit: 'GΩ', score: 6, weight: 1, status: 'Good', rec: 'Bình thường, tiếp tục theo dõi.' },
      { name: 'Polarization Index (PI)', shortName: 'PI', value: '1.5', unit: '', score: 4, weight: 1, status: 'Moderate', rec: 'Vệ sinh và sấy khô cuộn dây.' },
      { name: 'Dielectric Discharge', shortName: 'DD', value: '5', unit: '', score: 6, weight: 1, status: 'Good', rec: 'Bình thường.' },
      { name: 'ELCID', shortName: 'ELCID', value: '150', unit: 'mA', score: 4, weight: 1, status: 'Moderate', rec: 'Kiểm tra lõi từ, có dấu hiệu ngắn mạch cục bộ.' },
    ]
  },
  'GEN-02': {
    id: 'GEN-02',
    name: 'Máy phát điện dự phòng',
    type: 'Máy phát',
    health: 68,
    status: 'warning',
    factory: 'Nhà máy Bắc Ninh',
    location: 'Phòng Máy Phát',
    generalRecommendation: 'Thiết bị đang ở trạng thái cảnh báo. Cần bảo dưỡng hệ thống làm mát và kiểm tra lại độ rung của ổ trục.',
    failureProfile: [
      { subject: 'Exciter', score: 30 },
      { subject: 'Stator', score: 40 },
      { subject: 'Rotor', score: 50 },
      { subject: 'Bearings', score: 70 },
      { subject: 'Cooling', score: 60 },
      { subject: 'Controls', score: 20 },
    ],
    defects: [
      { name: 'High Vibration', warnings: 18.0, risk: 70.0 },
      { name: 'Cooling Inefficiency', warnings: 12.0, risk: 60.0 },
      { name: 'Rotor Imbalance', warnings: 10.0, risk: 50.0 },
      { name: 'Stator Winding Issue', warnings: 5.0, risk: 40.0 },
    ],
    tests: [
      { name: 'Độ rung (Vibration)', shortName: 'Độ rung', value: '5.2', unit: 'mm/s', score: 4, weight: 3, status: 'Moderate', rec: 'Độ rung cao, kiểm tra độ đồng tâm và vòng bi.' },
      { name: 'Nhiệt độ Stator', shortName: 'Nhiệt độ', value: '115', unit: '°C', score: 5, weight: 2, status: 'Moderate', rec: 'Nhiệt độ hơi cao, vệ sinh hệ thống tản nhiệt.' },
      { name: 'Điện trở cách điện', shortName: 'Cách điện', value: '50', unit: 'GΩ', score: 7, weight: 2, status: 'Good', rec: 'Cách điện tốt.' },
    ]
  },
  'TRF-03': {
    id: 'TRF-03',
    name: 'Máy biến áp T3',
    type: 'Máy biến áp',
    health: 71,
    status: 'warning',
    factory: 'Nhà máy Đồng Nai',
    location: 'Trạm biến áp 22kV',
    generalRecommendation: 'Thiết bị hoạt động ổn định nhưng có dấu hiệu suy giảm nhẹ ở chất lượng dầu. Cần lấy mẫu dầu kiểm tra lại sau 3 tháng.',
    failureProfile: [
      { subject: 'Mag Circuit', score: 10 },
      { subject: 'Winding', score: 20 },
      { subject: 'Main Tank', score: 15 },
      { subject: 'Cooling', score: 5 },
      { subject: 'Liquid', score: 60 },
      { subject: 'Bushing', score: 25 },
      { subject: 'LTC', score: 30 },
    ],
    defects: [
      { name: 'Moisture Contamination', warnings: 15.0, risk: 60.0 },
      { name: 'LTC Wear', warnings: 8.0, risk: 30.0 },
      { name: 'Bushing Degradation', warnings: 5.0, risk: 25.0 },
      { name: 'Winding Deformation', warnings: 2.0, risk: 20.0 },
    ],
    tests: [
      { name: 'Phân tích khí hòa tan (DGA)', shortName: 'DGA', value: '1500', unit: 'ppm', score: 6, weight: 3, status: 'Good', rec: 'Khí hòa tan trong giới hạn cho phép.' },
      { name: 'Độ ẩm trong dầu', shortName: 'Độ ẩm', value: '25', unit: 'ppm', score: 4, weight: 2, status: 'Moderate', rec: 'Độ ẩm hơi cao, theo dõi thêm.' },
      { name: 'Điện trở cách điện', shortName: 'Cách điện', value: '1200', unit: 'MΩ', score: 7, weight: 1, status: 'Good', rec: 'Cách điện tốt.' },
    ]
  }
};

const statusDistribution = [
  { name: 'Khỏe mạnh', value: 105, color: '#10b981' },
  { name: 'Cảnh báo', value: 14, color: '#f59e0b' },
  { name: 'Nguy hiểm', value: 5, color: '#f43f5e' },
];

const healthDistribution = [
  { range: '90-100%', count: 85 },
  { range: '70-89%', count: 20 },
  { range: '50-69%', count: 14 },
  { range: '<50%', count: 5 },
];


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

const generateMockData = () => {
  const factories = [
    { name: 'Nhà máy Nhiệt điện Phả Lại', customer: 'Tập đoàn Điện lực Việt Nam (EVN)' },
    { name: 'Nhà máy Thủy điện Hòa Bình', customer: 'Tập đoàn Điện lực Việt Nam (EVN)' },
    { name: 'Nhà máy Sữa Tiên Sơn', customer: 'Công ty CP Sữa Việt Nam (Vinamilk)' },
    { name: 'Nhà máy Sữa Bình Dương', customer: 'Công ty CP Sữa Việt Nam (Vinamilk)' },
    { name: 'Khu liên hợp Gang thép Hải Dương', customer: 'Tập đoàn Hòa Phát' },
    { name: 'Khu liên hợp Gang thép Dung Quất', customer: 'Tập đoàn Hòa Phát' },
    { name: 'Nhà máy Thaco Mazda', customer: 'Công ty CP Ô tô Trường Hải (Thaco)' },
    { name: 'Nhà máy Thaco Kia', customer: 'Công ty CP Ô tô Trường Hải (Thaco)' },
    { name: 'Nhà máy Xử lý Khí Dinh Cố', customer: 'Tổng Công ty Khí Việt Nam (PV Gas)' },
    { name: 'Trung tâm Dữ liệu Viettel Hòa Lạc', customer: 'Tập đoàn Công nghiệp Viễn thông Quân đội (Viettel)' },
    { name: 'Trung tâm Dữ liệu Viettel Bình Dương', customer: 'Tập đoàn Công nghiệp Viễn thông Quân đội (Viettel)' },
    { name: 'Nhà máy Masan Nutri-Science', customer: 'Công ty CP Tập đoàn Masan' },
    { name: 'Nhà máy Masan Consumer Bình Dương', customer: 'Công ty CP Tập đoàn Masan' },
    { name: 'Nhà máy VinFast Hải Phòng', customer: 'Tập đoàn Vingroup' },
    { name: 'FPT Tower', customer: 'Công ty CP FPT' },
    { name: 'Sân bay Quốc tế Nội Bài', customer: 'Tổng Công ty Cảng hàng không Việt Nam (ACV)' },
    { name: 'Sân bay Quốc tế Tân Sơn Nhất', customer: 'Tổng Công ty Cảng hàng không Việt Nam (ACV)' },
    { name: 'Nhà máy Thủy điện Sơn La', customer: 'Tập đoàn Điện lực Việt Nam (EVN)' },
    { name: 'Nhà máy Sữa Đà Nẵng', customer: 'Công ty CP Sữa Việt Nam (Vinamilk)' },
    { name: 'Khu liên hợp Gang thép Hưng Yên', customer: 'Tập đoàn Hòa Phát' },
  ];

  const types = ['Máy biến áp', 'Động cơ', 'Tủ điện'];
  const locations = ['Trạm biến áp chính', 'Phân xưởng 1', 'Phân xưởng 2', 'Trạm bơm', 'Phòng máy chủ', 'Khu vực làm mát'];
  
  const equipment = [];
  let idCounter = 1;

  factories.forEach(factory => {
    for (let i = 0; i < 5; i++) {
      const type = types[Math.floor(Math.random() * types.length)];
      const location = locations[Math.floor(Math.random() * locations.length)];
      const health = Math.floor(Math.random() * 60) + 40; // 40 to 100
      let status = 'healthy';
      if (health < 60) status = 'critical';
      else if (health < 80) status = 'warning';

      let prefix = 'EQ';
      if (type === 'Máy biến áp') prefix = 'TRF';
      if (type === 'Động cơ') prefix = 'MTR';
      if (type === 'Tủ điện') prefix = 'SWG';

      equipment.push({
        id: `${prefix}-${idCounter.toString().padStart(3, '0')}`,
        name: `${type} ${i + 1} - ${factory.name}`,
        customer: factory.customer,
        factory: factory.name,
        location: location,
        type: type,
        health: health,
        status: status,
        lastCheck: `1${Math.floor(Math.random() * 9)}/03/2026`
      });
      idCounter++;
    }
  });

  return equipment;
};

const initialAllEquipment = generateMockData();

const pastReports = [
  { id: 'REP-2026-02', date: '10/02/2026', inspector: 'Nguyễn Văn A', status: 'warning', notes: 'Nhiệt độ dầu hơi cao, cần theo dõi.' },
  { id: 'REP-2025-08', date: '15/08/2025', inspector: 'Trần Văn B', status: 'healthy', notes: 'Thiết bị hoạt động bình thường.' },
  { id: 'REP-2025-02', date: '20/02/2025', inspector: 'Nguyễn Văn A', status: 'healthy', notes: 'Bảo dưỡng định kỳ.' },
];

const paramHistoryData = {
  'Nhiệt độ dầu': [
    { date: '20/02/2025', value: 75 },
    { date: '15/08/2025', value: 80 },
    { date: '10/02/2026', value: 85 },
    { date: '17/03/2026', value: 92 },
  ],
  'Nhiệt độ cuộn dây': [
    { date: '20/02/2025', value: 85 },
    { date: '15/08/2025', value: 88 },
    { date: '10/02/2026', value: 92 },
    { date: '17/03/2026', value: 92 },
  ],
  'Điện trở cách điện (Cao-Hạ)': [
    { date: '20/02/2025', value: 1500 },
    { date: '15/08/2025', value: 1400 },
    { date: '10/02/2026', value: 1250 },
    { date: '17/03/2026', value: 1100 },
  ],
};

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

const HealthBar = ({ value }: { value: number }) => {
  let color = 'bg-emerald-500';
  if (value < 60) color = 'bg-rose-500';
  else if (value < 80) color = 'bg-amber-500';

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

const regionsT1: { id: string, label?: string, color: string, points: number[][] }[] = [
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

export default function App() {
  const [user, setUser] = useState<any>(null);
  const [userRole, setUserRole] = useState<string | null>(null);
  const [userCustomerName, setUserCustomerName] = useState<string | null>(null);
  const [isAuthLoading, setIsAuthLoading] = useState(true);

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
            setUserRole(userData.role);
            if (userData.role === 'customer' && userData.customerId) {
              // Fetch customer name
              const customerDoc = await getDoc(doc(db, 'customers', userData.customerId));
              if (customerDoc.exists()) {
                setUserCustomerName(customerDoc.data().name);
              }
            }
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
          }
        } else {
          setUser(null);
          setUserRole(null);
          setUserCustomerName(null);
        }
      } catch (error) {
        console.error('Error in auth state change:', error);
        // Fallback state
        setUser(currentUser);
        setUserRole(currentUser?.email === 'sgm1707@gmail.com' ? 'admin' : 'customer');
      } finally {
        setIsAuthLoading(false);
      }
    });

    return () => unsubscribe();
  }, []);

  const handleLogin = async () => {
    try {
      await signInWithGoogle();
    } catch (error) {
      console.error(error);
    }
  };

  const [activeTab, setActiveTab] = useState('dashboard');

  const handleTabChange = (tab: string) => {
    setActiveTab(tab);
    if (tab === 'dashboard' && isGoogleConnected) {
      handleFetchFromSheets();
    }
  };
  const [showHistoryModal, setShowHistoryModal] = useState(false);
  const [selectedParamHistory, setSelectedParamHistory] = useState('');
  const [showEquipmentProfile, setShowEquipmentProfile] = useState(false);
  const [selectedEqType, setSelectedEqType] = useState('Máy biến áp');
  const [equipmentCode, setEquipmentCode] = useState('TRF-01');
  const [equipmentName, setEquipmentName] = useState('Máy biến áp T1');
  const [siteName, setSiteName] = useState('Nhà máy Bắc Ninh');
  const [customerName, setCustomerName] = useState('Công ty Điện lực A');
  const [locationName, setLocationName] = useState('Trạm biến áp 110kV');
  const [showEqSuggestions, setShowEqSuggestions] = useState(false);
  const [showQRScanner, setShowQRScanner] = useState(false);
  const [isNewEquipment, setIsNewEquipment] = useState(false);
  const [formData, setFormData] = useState<Record<string, any>>({});
  const [healthResult, setHealthResult] = useState<{index: number, status: string} | null>(null);
  const [selectedRiskDetail, setSelectedRiskDetail] = useState<string | null>(null);
  const [allEquipment, setAllEquipment] = useState(initialAllEquipment);

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
      alert(`Đã tìm thấy thiết bị: ${foundEq.name} (${foundEq.id})`);
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
    const getStatus = (score: number) => score >= 8 ? 'Critical' : score >= 4 ? 'Moderate' : 'Good';
    const getRec = (score: number, critRec: string, modRec: string) => score >= 8 ? critRec : score >= 4 ? modRec : 'Bình thường';

    let generalRecommendation = '';
    let failureProfile: any[] = [];
    let generatedDefects: any[] = [];
    let tests: any[] = [];

    if (eq.type === 'Máy biến áp') {
      // Generate realistic values based on DNO Common Network Asset Indices Methodology
      const moisture = isCritical ? 35 + Math.random() * 15 : isWarning ? 20 + Math.random() * 15 : 5 + Math.random() * 10;
      const acidity = isCritical ? 0.2 + Math.random() * 0.15 : isWarning ? 0.12 + Math.random() * 0.08 : 0.02 + Math.random() * 0.08;
      const bdv = isCritical ? 25 + Math.random() * 10 : isWarning ? 35 + Math.random() * 10 : 55 + Math.random() * 15;
      const ffa = isCritical ? 5.5 + Math.random() * 2 : isWarning ? 3.5 + Math.random() * 2 : 0.5 + Math.random() * 2;
      
      // Scoring based on PDF calibration tables
      const moistureScore = moisture > 45 ? 10 : moisture > 35 ? 8 : moisture > 25 ? 4 : moisture > 15 ? 2 : 0;
      const acidityScore = acidity > 0.3 ? 10 : acidity > 0.2 ? 8 : acidity > 0.15 ? 4 : acidity > 0.1 ? 2 : 0;
      const bdvScore = bdv < 30 ? 10 : bdv < 40 ? 4 : bdv < 50 ? 2 : 0;
      const ffaScore = ffa > 7 ? 10 : ffa > 6 ? 8 : ffa > 5 ? 6 : ffa > 4 ? 4 : 0;
      const dgaScore = isCritical ? 9 : isWarning ? 5 : 1;
      const pdScore = isCritical ? 8 : isWarning ? 4 : 1;

      generalRecommendation = isCritical 
        ? 'Cảnh báo mức độ cao. Cần tiến hành kiểm tra DGA, đo PD và lên kế hoạch bảo dưỡng/thay thế ngay lập tức.' 
        : isWarning 
          ? 'Thiết bị đang ở trạng thái cảnh báo. Cần tăng cường tần suất theo dõi tình trạng dầu và nhiệt độ.'
          : 'Thiết bị hoạt động bình thường. Tiếp tục bảo dưỡng định kỳ theo khuyến cáo của nhà sản xuất.';

      failureProfile = [
        { subject: 'Tình trạng bên ngoài', score: Math.min(100, baseRisk + Math.random() * 20) },
        { subject: 'Phóng điện cục bộ (PD)', score: Math.min(100, pdScore * 10 + Math.random() * 10) },
        { subject: 'Chất lượng dầu', score: Math.min(100, Math.max(moistureScore, acidityScore, bdvScore) * 10 + Math.random() * 10) },
        { subject: 'Khí hòa tan (DGA)', score: Math.min(100, dgaScore * 10 + Math.random() * 10) },
        { subject: 'Lão hóa giấy (FFA)', score: Math.min(100, ffaScore * 10 + Math.random() * 10) },
      ];

      generatedDefects = [
        { name: 'Suy giảm chất lượng dầu (Oil Degradation)', warnings: Math.max(moistureScore, acidityScore) * 2.5, risk: Math.max(moistureScore, acidityScore) * 10 },
        { name: 'Phóng điện cục bộ (Partial Discharge)', warnings: pdScore * 2, risk: pdScore * 10 },
        { name: 'Lão hóa giấy cách điện (Paper Ageing)', warnings: ffaScore * 2, risk: ffaScore * 10 },
      ].filter(d => d.risk > 0);

      tests = [
        { name: 'Khí hòa tan (DGA)', shortName: 'DGA', value: isCritical ? 'Mức cao' : isWarning ? 'Cảnh báo' : 'Bình thường', unit: '', score: dgaScore, weight: 3, status: getStatus(dgaScore), rec: getRec(dgaScore, 'Lọc/thay dầu ngay', 'Theo dõi thêm') },
        { name: 'Hàm lượng Furan (FFA)', shortName: 'FFA', value: ffa.toFixed(2), unit: 'ppm', score: ffaScore, weight: 2, status: getStatus(ffaScore), rec: getRec(ffaScore, 'Kiểm tra cách điện giấy', 'Theo dõi xu hướng') },
        { name: 'Độ ẩm trong dầu', shortName: 'Moisture', value: moisture.toFixed(1), unit: 'ppm', score: moistureScore, weight: 2, status: getStatus(moistureScore), rec: getRec(moistureScore, 'Xử lý lọc dầu', 'Kiểm tra độ kín') },
        { name: 'Độ axit', shortName: 'Acidity', value: acidity.toFixed(3), unit: 'mg KOH/g', score: acidityScore, weight: 1, status: getStatus(acidityScore), rec: getRec(acidityScore, 'Thay dầu', 'Theo dõi') },
        { name: 'Điện áp đánh thủng', shortName: 'BDV', value: bdv.toFixed(1), unit: 'kV', score: bdvScore, weight: 2, status: getStatus(bdvScore), rec: getRec(bdvScore, 'Xử lý lọc dầu', 'Theo dõi thêm') },
        { name: 'Phóng điện cục bộ', shortName: 'PD', value: isCritical ? 'Cao' : isWarning ? 'Trung bình' : 'Thấp', unit: '', score: pdScore, weight: 3, status: getStatus(pdScore), rec: getRec(pdScore, 'Kiểm tra siêu âm/TEV', 'Theo dõi định kỳ') },
      ];
    } else if (eq.type === 'Tủ điện trung thế' || eq.type === 'Tủ điện') {
      const tevLevel = isCritical ? 30 + Math.random() * 20 : isWarning ? 20 + Math.random() * 9 : 5 + Math.random() * 14;
      const pulses = isCritical ? Math.floor(1 + Math.random() * 28) : isWarning ? Math.floor(1 + Math.random() * 5) : 0;
      const ultrasonic = isCritical ? 20 + Math.random() * 20 : isWarning ? 5 + Math.random() * 15 : Math.random() * 5;

      const tevScore = tevLevel > 29 ? 10 : tevLevel >= 20 ? 6 : 1;
      const pulsesScore = pulses >= 7 && pulses <= 29 ? 10 : pulses >= 1 && pulses <= 6 ? 8 : pulses >= 30 ? 4 : 1; // 30-50 is noise
      const ultrasonicScore = ultrasonic > 15 ? 10 : ultrasonic > 5 ? 6 : 1;

      generalRecommendation = isCritical 
          ? 'Khả năng rất cao có hiện tượng phóng điện cục bộ. Yêu cầu kiểm tra vào lần dừng máy kế tiếp hoặc dừng máy ngay để xác định nguồn.' 
          : isWarning 
            ? 'Có khả năng xảy ra phóng điện cục bộ. Kiểm tra số xung/chu kỳ và theo dõi định kỳ.'
            : 'Không có dấu hiệu phóng điện cục bộ. Kiểm tra lại trong vòng 12 tháng.';

      failureProfile = [
          { subject: 'Tình trạng bên ngoài', score: Math.min(100, baseRisk + Math.random() * 20) },
          { subject: 'Phóng điện TEV', score: Math.min(100, tevScore * 10 + Math.random() * 10) },
          { subject: 'Phóng điện siêu âm', score: Math.min(100, ultrasonicScore * 10 + Math.random() * 10) },
          { subject: 'Mật độ xung', score: Math.min(100, pulsesScore * 10 + Math.random() * 10) },
      ];

      generatedDefects = [
        { name: 'Phóng điện bề mặt (Surface PD)', warnings: ultrasonicScore * 2, risk: ultrasonicScore * 10 },
        { name: 'Phóng điện bên trong (Internal PD)', warnings: tevScore * 2, risk: tevScore * 10 },
      ].filter(d => d.risk > 10);

      tests = [
          { name: 'Mức TEV', shortName: 'TEV', value: tevLevel.toFixed(1), unit: 'dB', score: tevScore, weight: 3, status: getStatus(tevScore), rec: getRec(tevScore, 'Dừng máy kiểm tra', 'Theo dõi') },
          { name: 'Số xung/chu kỳ', shortName: 'Pulses', value: pulses.toString(), unit: 'xung', score: pulsesScore, weight: 2, status: getStatus(pulsesScore), rec: getRec(pulsesScore, 'Định vị nguồn PD', 'Kiểm tra lại') },
          { name: 'Siêu âm', shortName: 'Ultrasonic', value: ultrasonic.toFixed(1), unit: 'dBµV', score: ultrasonicScore, weight: 2, status: getStatus(ultrasonicScore), rec: getRec(ultrasonicScore, 'Kiểm tra phóng điện bề mặt', 'Theo dõi') },
      ];
    } else if (eq.type === 'Động cơ') {
      const tanDelta = isCritical ? 0.08 + Math.random() * 0.05 : isWarning ? 0.04 + Math.random() * 0.03 : 0.01 + Math.random() * 0.02;
      const pd = isCritical ? 18000 + Math.random() * 5000 : isWarning ? 8000 + Math.random() * 7000 : 2000 + Math.random() * 3000;
      const ir = isCritical ? 0.5 + Math.random() * 0.5 : isWarning ? 5 + Math.random() * 5 : 60 + Math.random() * 50;
      const pi = isCritical ? 0.8 + Math.random() * 0.2 : isWarning ? 1.2 + Math.random() * 0.8 : 2.5 + Math.random() * 1.5;
      const elcid = isCritical ? 250 + Math.random() * 100 : isWarning ? 120 + Math.random() * 80 : 40 + Math.random() * 30;

      const tanDeltaScore = tanDelta >= 0.1 ? 10 : tanDelta >= 0.07 ? 8 : tanDelta >= 0.04 ? 4 : 1;
      const pdScore = pd >= 20000 ? 10 : pd >= 15000 ? 8 : pd >= 10000 ? 4 : 1;
      const irScore = ir < 1 ? 10 : ir < 10 ? 8 : ir < 50 ? 4 : 1;
      const piScore = pi < 1 ? 10 : pi < 2 ? 6 : 1;
      const elcidScore = elcid > 200 ? 10 : elcid > 110 ? 8 : elcid > 70 ? 4 : 1;

      generalRecommendation = isCritical 
          ? 'Tình trạng cách điện suy giảm nghiêm trọng. Cần lên kế hoạch bảo dưỡng, sấy cuộn dây hoặc quấn lại.' 
          : isWarning 
            ? 'Có dấu hiệu lão hóa cách điện. Tăng cường theo dõi PD và điện trở cách điện.'
            : 'Cách điện động cơ ở trạng thái tốt. Tiếp tục vận hành và bảo dưỡng định kỳ.';

      failureProfile = [
          { subject: 'Tổn hao điện môi (Tan-delta)', score: Math.min(100, tanDeltaScore * 10 + Math.random() * 10) },
          { subject: 'Phóng điện cục bộ (PD)', score: Math.min(100, pdScore * 10 + Math.random() * 10) },
          { subject: 'Điện trở cách điện (IR)', score: Math.min(100, irScore * 10 + Math.random() * 10) },
          { subject: 'Chỉ số phân cực (PI)', score: Math.min(100, piScore * 10 + Math.random() * 10) },
          { subject: 'Lõi thép (ELCID)', score: Math.min(100, elcidScore * 10 + Math.random() * 10) },
      ];

      generatedDefects = [
        { name: 'Lão hóa cách điện cuộn dây', warnings: irScore * 2, risk: irScore * 10 },
        { name: 'Phóng điện cục bộ stator', warnings: pdScore * 2, risk: pdScore * 10 },
        { name: 'Hỏng hóc lõi thép', warnings: elcidScore * 2, risk: elcidScore * 10 },
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

  const dynamicSiteData = useMemo(() => {
    const siteMap = new Map();
    
    // Known coordinates for nice map display
    const knownCoords: Record<string, {lat: number, lng: number}> = {
      'Nhà máy Bắc Ninh': { lat: 21.1861, lng: 106.0763 },
      'Nhà máy Bình Dương': { lat: 11.0648, lng: 106.6615 },
      'Nhà máy Đồng Nai': { lat: 10.9487, lng: 106.8243 },
      'Nhà máy Nhiệt điện Phả Lại': { lat: 21.111, lng: 106.315 },
      'Nhà máy Thủy điện Hòa Bình': { lat: 20.808, lng: 105.328 },
      'Nhà máy Sữa Tiên Sơn': { lat: 21.125, lng: 106.012 },
      'Nhà máy Sữa Bình Dương': { lat: 11.064, lng: 106.661 },
      'Khu liên hợp Gang thép Hải Dương': { lat: 21.000, lng: 106.500 },
      'Khu liên hợp Gang thép Dung Quất': { lat: 15.385, lng: 108.794 },
      'Nhà máy Thaco Mazda': { lat: 15.441, lng: 108.647 },
      'Nhà máy Thaco Kia': { lat: 15.445, lng: 108.650 },
      'Nhà máy Xử lý Khí Dinh Cố': { lat: 10.420, lng: 107.195 },
      'Trung tâm Dữ liệu Viettel Hòa Lạc': { lat: 21.015, lng: 105.526 },
      'Trung tâm Dữ liệu Viettel Bình Dương': { lat: 11.050, lng: 106.650 },
      'Nhà máy Masan Nutri-Science': { lat: 20.550, lng: 105.950 },
      'Nhà máy Masan Consumer Bình Dương': { lat: 11.060, lng: 106.660 },
      'Nhà máy VinFast Hải Phòng': { lat: 20.816, lng: 106.852 },
      'FPT Tower': { lat: 21.031, lng: 105.782 },
      'Sân bay Quốc tế Nội Bài': { lat: 21.218, lng: 105.804 },
      'Sân bay Quốc tế Tân Sơn Nhất': { lat: 10.818, lng: 106.658 },
      'Nhà máy Thủy điện Sơn La': { lat: 21.496, lng: 103.995 },
      'Nhà máy Sữa Đà Nẵng': { lat: 16.030, lng: 108.180 },
      'Khu liên hợp Gang thép Formosa Hà Tĩnh': { lat: 18.040, lng: 106.330 },
    };

    allEquipment.forEach(eq => {
      const factory = eq.factory || 'Chưa xác định';
      const customer = eq.customer || 'Chưa xác định';
      const id = factory.toLowerCase().replace(/\s+/g, '-');
      
      if (!siteMap.has(id)) {
        // Generate some pseudo-random coordinates in VN if unknown
        const hash = id.split('').reduce((acc, char) => acc + char.charCodeAt(0), 0);
        const lat = knownCoords[factory]?.lat || (10 + (hash % 10) + (hash % 100) / 100);
        const lng = knownCoords[factory]?.lng || (105 + (hash % 5) + (hash % 100) / 100);

        siteMap.set(id, {
          id,
          name: factory,
          customer: customer,
          lat,
          lng,
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
  }, [allEquipment]);

  const uniqueCustomers = Array.from(new Set(dynamicSiteData.map(site => site.customer)));

  const filteredSiteData = dynamicSiteData.filter(site => {
    const matchCustomer = selectedCustomer === 'all' || site.customer === selectedCustomer;
    const matchFactory = selectedFactory === 'all' || site.id === selectedFactory;
    return matchCustomer && matchFactory;
  });

  const filteredKpiData = {
    total: filteredSiteData.reduce((sum, site) => sum + site.count, 0),
    healthy: filteredSiteData.reduce((sum, site) => sum + site.healthy, 0),
    warning: filteredSiteData.reduce((sum, site) => sum + site.warning, 0),
    critical: filteredSiteData.reduce((sum, site) => sum + site.critical, 0),
  };

  const dynamicTrendData = useMemo(() => {
    const eq = allEquipment.find(e => e.id === trendingEquipment);
    if (!eq) return { temp: [], tev: [], dga: [] };

    // Simple hash function to generate deterministic pseudo-random numbers based on equipment ID
    const seed = eq.id.split('').reduce((acc, char) => acc + char.charCodeAt(0), 0);
    const random = (offset: number) => {
      const x = Math.sin(seed + offset) * 10000;
      return x - Math.floor(x);
    };

    const isCritical = eq.status === 'critical';
    const isWarning = eq.status === 'warning';
    
    const baseTemp = isCritical ? 85 : isWarning ? 75 : 60;
    const tempTrend = Array.from({ length: 9 }).map((_, i) => ({
      time: `${(i + 8).toString().padStart(2, '0')}:00`,
      temp: Math.round(baseTemp + random(i) * 10 + (isCritical ? i * 1.5 : isWarning ? i * 0.5 : 0))
    }));

    const baseTev = isCritical ? 20 : isWarning ? 10 : 2;
    const baseUltra = isCritical ? 10 : isWarning ? 5 : 1;
    const tevTrend = Array.from({ length: 9 }).map((_, i) => ({
      time: `${(i + 8).toString().padStart(2, '0')}:00`,
      tev: Math.round(baseTev + random(i + 100) * 5 + (isCritical ? i * 2 : isWarning ? i * 1 : 0)),
      ultrasonic: Math.round(baseUltra + random(i + 200) * 3 + (isCritical ? i * 1.5 : isWarning ? i * 0.5 : 0))
    }));

    const baseH2 = isCritical ? 50 : isWarning ? 20 : 5;
    const dgaTrend = Array.from({ length: 6 }).map((_, i) => {
      const month = (i * 3 + 1) % 12 || 12;
      const year = 25 + Math.floor((i * 3 + 1) / 12);
      const multiplier = isCritical ? Math.pow(1.5, i) : isWarning ? Math.pow(1.2, i) : 1;
      return {
        time: `T${month}/${year}`,
        h2: Math.round((baseH2 + random(i + 300) * 10) * multiplier),
        ch4: Math.round((baseH2 * 0.7 + random(i + 400) * 5) * multiplier),
        c2h6: Math.round((baseH2 * 0.4 + random(i + 500) * 3) * multiplier),
        c2h4: Math.round((baseH2 * 0.3 + random(i + 600) * 2) * multiplier),
        c2h2: Math.round((isCritical ? 5 : 0) + random(i + 700) * 2 * multiplier),
        co: Math.round((150 + random(i + 800) * 50) * (multiplier > 1 ? multiplier * 0.5 : 1)),
        co2: Math.round((1200 + random(i + 900) * 200) * (multiplier > 1 ? multiplier * 0.5 : 1))
      };
    });

    return { temp: tempTrend, tev: tevTrend, dga: dgaTrend };
  }, [trendingEquipment, allEquipment]);

  const selectedFactoryName = selectedFactory === 'all' 
    ? 'all' 
    : dynamicSiteData.find(s => s.id === selectedFactory)?.name;

  const filteredEquipment = allEquipment.filter(eq => {
    const matchCustomer = selectedCustomer === 'all' || eq.customer === selectedCustomer;
    const matchFactory = selectedFactoryName === 'all' || eq.factory === selectedFactoryName;
    return matchCustomer && matchFactory;
  });

  const displayedEquipment = filteredEquipment.filter(eq => {
    const matchStatus = statusFilter === 'all' || eq.status === statusFilter;
    const matchName = filterName === '' || eq.name.toLowerCase().includes(filterName.toLowerCase());
    const matchLocation = filterLocation === '' || eq.factory.toLowerCase().includes(filterLocation.toLowerCase()) || eq.location.toLowerCase().includes(filterLocation.toLowerCase());
    const matchType = filterType === 'all' || eq.type === filterType;
    const matchHealthMin = filterHealthMin === '' || eq.health >= parseInt(filterHealthMin);
    const matchHealthMax = filterHealthMax === '' || eq.health <= parseInt(filterHealthMax);
    return matchStatus && matchName && matchLocation && matchType && matchHealthMin && matchHealthMax;
  });

  // Pagination
  const itemsPerPage = 20;
  const totalPages = Math.ceil(displayedEquipment.length / itemsPerPage);
  const paginatedEquipment = displayedEquipment.slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage);

  const filteredTopRiskEquipment = filteredEquipment
    .filter(eq => eq.status === 'critical' || eq.status === 'warning')
    .sort((a, b) => a.health - b.health)
    .slice(0, 5);

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
    const checkAuthStatus = async () => {
      try {
        const response = await fetch('/api/auth/status');
        if (response.ok) {
          const data = await response.json();
          setIsGoogleConnected(data.isAuthenticated);
        }
      } catch (error) {
        console.error('Failed to check auth status', error);
      }
    };
    checkAuthStatus();

    const handleMessage = (event: MessageEvent) => {
      const origin = event.origin;
      if (!origin.endsWith('.run.app') && !origin.includes('localhost')) {
        return;
      }
      if (event.data?.type === 'OAUTH_AUTH_SUCCESS') {
        setIsGoogleConnected(true);
      }
    };
    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, []);

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

  const handleSyncToSheets = async () => {
    if (!isGoogleConnected) {
      alert('Vui lòng kết nối Google Drive trước khi đồng bộ.');
      return;
    }

    setIsSyncing(true);
    try {
      // Prepare data for all equipment
      const transformers: any[] = [];
      const switchgears: any[] = [];
      const motors: any[] = [];
      
      // Headers
      transformers.push(['Thời gian kiểm tra', 'Khách hàng', 'Nhà máy / Site', 'Vị trí / Khu vực', 'Mã thiết bị', 'Tên thiết bị', 'Loại thiết bị', 'Điểm sức khỏe (%)', 'Trạng thái', 'Nhiệt độ dầu (°C)', 'Nhiệt độ cuộn dây (°C)', 'IR Cao-Hạ (MΩ)', 'IR Cao-Vỏ (MΩ)', 'Tình trạng rò rỉ dầu', 'Khí hòa tan DGA (ppm)', 'Độ bền điện môi (kV)', 'Hàm lượng Furan (mg/kg)', 'Độ ẩm trong dầu (ppm)']);
      switchgears.push(['Thời gian kiểm tra', 'Khách hàng', 'Nhà máy / Site', 'Vị trí / Khu vực', 'Mã thiết bị', 'Tên thiết bị', 'Loại thiết bị', 'Điểm sức khỏe (%)', 'Trạng thái', 'Chụp ảnh nhiệt (°C)', 'Điện trở tiếp xúc (μΩ)', 'TEV (dBmV)', 'Siêu âm (dBμV)', 'Xung TEV (pps)', 'Độ ẩm (%)', 'Áp suất khí SF6 (bar)']);
      motors.push(['Thời gian kiểm tra', 'Khách hàng', 'Nhà máy / Site', 'Vị trí / Khu vực', 'Mã thiết bị', 'Tên thiết bị', 'Loại thiết bị', 'Điểm sức khỏe (%)', 'Trạng thái', 'Độ rung (mm/s)', 'Nhiệt độ Stator (°C)', 'Điện trở cách điện (GΩ)', 'Phóng điện cục bộ PD (pC)', 'Độ lệch điện áp (%)', 'Chỉ số phân cực PI', 'Nhiệt độ vòng bi (°C)', 'Tan Delta']);

      // Map existing data
      allEquipment.forEach(eq => {
        const baseData = [eq.lastCheck || new Date().toLocaleString('vi-VN'), eq.customer, eq.factory, eq.location, eq.id, eq.name, eq.type, eq.health, eq.status];
        if (eq.type === 'Máy biến áp') {
          transformers.push([...baseData, '', '', '', '', '', '', '', '', '']);
        } else if (eq.type === 'Tủ điện trung thế' || eq.type === 'Tủ điện') {
          switchgears.push([...baseData, '', '', '', '', '', '', '']);
        } else if (eq.type === 'Động cơ') {
          motors.push([...baseData, '', '', '', '', '', '', '', '']);
        }
      });

      const response = await fetch('/api/sheets/sync-export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          transformers,
          switchgears,
          motors
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

      setSyncSuccess(true);
      setTimeout(() => setSyncSuccess(false), 3000);
      alert('Đồng bộ dữ liệu lên Google Sheets thành công!');
    } catch (error: any) {
      console.error('Sync error:', error);
      alert('Lỗi khi đồng bộ dữ liệu: ' + error.message);
    } finally {
      setIsSyncing(false);
    }
  };

  const handleRegenerateData = () => {
    setAllEquipment(generateMockData());
    alert('Đã tạo lại 100 thiết bị mẫu. Bạn có thể bấm "Đồng bộ lên Sheets" để lưu.');
  };

  const handleFetchFromSheets = async () => {
    if (!isGoogleConnected) {
      return;
    }

    setIsSyncing(true);
    try {
      const response = await fetch('/api/sheets/get');
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
        throw new Error(err.error || 'Failed to fetch');
      }

      const { data } = await response.json();
      
      const newEquipmentList: any[] = [];
      
      // Parse transformers
      if (data.transformers && data.transformers.length > 1) {
        const rows = data.transformers.slice(1); // Skip header
        rows.forEach((row: any[]) => {
          if (row[4] && row[5]) { // Must have ID and Name
            const healthVal = parseInt(row[7]) || 0;
            newEquipmentList.push({
              lastCheck: row[0] || '',
              customer: row[1] || '',
              factory: row[2] || '',
              location: row[3] || '',
              id: row[4],
              name: row[5],
              type: row[6] || 'Máy biến áp',
              health: healthVal,
              status: healthVal < 60 ? 'critical' : healthVal < 80 ? 'warning' : 'healthy'
            });
          }
        });
      }

      // Parse switchgears
      if (data.switchgears && data.switchgears.length > 1) {
        const rows = data.switchgears.slice(1);
        rows.forEach((row: any[]) => {
          if (row[4] && row[5]) {
            const healthVal = parseInt(row[7]) || 0;
            newEquipmentList.push({
              lastCheck: row[0] || '',
              customer: row[1] || '',
              factory: row[2] || '',
              location: row[3] || '',
              id: row[4],
              name: row[5],
              type: row[6] || 'Tủ điện',
              health: healthVal,
              status: healthVal < 60 ? 'critical' : healthVal < 80 ? 'warning' : 'healthy'
            });
          }
        });
      }

      // Parse motors
      if (data.motors && data.motors.length > 1) {
        const rows = data.motors.slice(1);
        rows.forEach((row: any[]) => {
          if (row[4] && row[5]) {
            const healthVal = parseInt(row[7]) || 0;
            newEquipmentList.push({
              lastCheck: row[0] || '',
              customer: row[1] || '',
              factory: row[2] || '',
              location: row[3] || '',
              id: row[4],
              name: row[5],
              type: row[6] || 'Động cơ',
              health: healthVal,
              status: healthVal < 60 ? 'critical' : healthVal < 80 ? 'warning' : 'healthy'
            });
          }
        });
      }

      if (newEquipmentList.length > 0) {
        setAllEquipment(newEquipmentList);
        setSyncSuccess(true);
        setTimeout(() => setSyncSuccess(false), 3000);
      }
    } catch (error: any) {
      console.error('Fetch error:', error);
      // Don't alert on auto-fetch to avoid annoying the user
    } finally {
      setIsSyncing(false);
    }
  };

  // Check auth status on mount
  useEffect(() => {
    const checkAuthStatus = async () => {
      try {
        const response = await fetch('/api/auth/status');
        const data = await response.json();
        if (data.isAuthenticated) {
          setIsGoogleConnected(true);
        }
      } catch (error) {
        console.error('Failed to check auth status:', error);
      }
    };
    checkAuthStatus();
  }, []);

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
      // 1. Upload files to Drive if any
      let attachmentLinks = '';
      if (attachedFiles.length > 0) {
        const uploadFormData = new FormData();
        attachedFiles.forEach(file => {
          uploadFormData.append('files', file);
        });

        const uploadRes = await fetch('/api/drive/upload', {
          method: 'POST',
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
      }

      // Prepare data row based on equipment type
      const now = new Date().toLocaleString('vi-VN');
      const baseData = [now, customerName, siteName, locationName, equipmentCode, equipmentName, selectedEqType, healthResult?.index || 'N/A', healthResult?.status || 'N/A', attachmentLinks];
      
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
          formData.oilMoisture || ''
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
          formData.sf6Pressure || ''
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
          formData.tanDelta || ''
        ];
        range = 'Động cơ!A:Z';
      }

      const response = await fetch('/api/sheets/append', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
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

  // Algorithm to evaluate single parameter against standards
  const evaluateParam = (type: string, param: string, value: any) => {
    if (value === '' || value === undefined) return null;
    const numVal = Number(value);
    
    if (type === 'Máy biến áp') {
      if (param === 'oilTemp') return numVal <= 80 ? 'healthy' : numVal <= 90 ? 'warning' : 'critical';
      if (param === 'windingTemp') return numVal <= 90 ? 'healthy' : numVal <= 105 ? 'warning' : 'critical';
      if (param === 'irHighLow') return numVal >= 2000 ? 'healthy' : numVal >= 1000 ? 'warning' : 'critical';
      if (param === 'irHighEarth') return numVal >= 2000 ? 'healthy' : numVal >= 1000 ? 'warning' : 'critical';
      if (param === 'oilLeak') return value === 'normal' ? 'healthy' : value === 'light' ? 'warning' : 'critical';
      if (param === 'dga') return numVal <= 1000 ? 'healthy' : numVal <= 2500 ? 'warning' : 'critical';
      if (param === 'dielectricStrength') return numVal >= 50 ? 'healthy' : numVal >= 40 ? 'warning' : 'critical';
      if (param === 'furan') return numVal <= 1 ? 'healthy' : numVal <= 5 ? 'warning' : 'critical';
      if (param === 'oilMoisture') return numVal <= 15 ? 'healthy' : numVal <= 25 ? 'warning' : 'critical';
    }
    if (type === 'Động cơ') {
      if (param === 'statorTemp') return numVal <= 110 ? 'healthy' : numVal <= 130 ? 'warning' : 'critical';
      if (param === 'bearingTemp') return numVal <= 80 ? 'healthy' : numVal <= 95 ? 'warning' : 'critical';
      if (param === 'vibration') return numVal <= 2.3 ? 'healthy' : numVal <= 4.5 ? 'warning' : 'critical';
      if (param === 'tanDelta') return numVal <= 0.02 ? 'healthy' : numVal <= 0.05 ? 'warning' : 'critical';
      if (param === 'pd') return numVal <= 1000 ? 'healthy' : numVal <= 5000 ? 'warning' : 'critical';
      if (param === 'ir') return numVal >= 100 ? 'healthy' : numVal >= 50 ? 'warning' : 'critical';
      if (param === 'pi') return numVal >= 2.0 ? 'healthy' : numVal >= 1.0 ? 'warning' : 'critical';
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
    return null;
  };

  // Calculate overall Health Index
  useEffect(() => {
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

    if (count === 0) {
      setHealthResult(null);
      return;
    }

    const avgScore = Math.round(totalScore / count);
    let finalStatus = 'healthy';
    if (hasCritical || avgScore < 60) finalStatus = 'critical';
    else if (hasWarning || avgScore < 80) finalStatus = 'warning';

    setHealthResult({ index: avgScore, status: finalStatus });
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

  const renderTransformerParams = () => (
    <>
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
    </>
  );

  const renderMotorParams = () => (
    <>
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Thermometer size={16} className="text-rose-500" />
          Nhiệt độ
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
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
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Activity size={16} className="text-blue-500" />
          Độ rung (Vibration)
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <ParamInput 
            title="Độ rung tổng thể" unit="mm/s" standard="≤ 4.5 mm/s" prevValue="2.1 mm/s" prevTrend="up"
            value={formData.vibration} onChange={(val: any) => setFormData({...formData, vibration: val})}
            evalStatus={evaluateParam('Động cơ', 'vibration', formData.vibration)}
            onHistoryClick={() => { setSelectedParamHistory('Độ rung'); setShowHistoryModal(true); }}
          />
        </div>
      </div>
      <div>
        <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-3 flex items-center gap-2 border-b border-slate-100 pb-2">
          <Zap size={16} className="text-amber-500" />
          Kiểm tra điện & Cách điện
        </h3>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <ParamInput 
            title="Mất cân bằng điện áp (Voltage Imbalance)" unit="%" standard="≤ 3%" prevValue="1.5%" prevTrend="stable"
            value={formData.voltageImbalance} onChange={(val: any) => setFormData({...formData, voltageImbalance: val})}
            evalStatus={evaluateParam('Động cơ', 'voltageImbalance', formData.voltageImbalance)}
            onHistoryClick={() => { setSelectedParamHistory('Mất cân bằng điện áp'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Điện trở cách điện (IR)" unit="GΩ" standard="≥ 100 GΩ" prevValue="120 GΩ" prevTrend="down"
            value={formData.ir} onChange={(val: any) => setFormData({...formData, ir: val})}
            evalStatus={evaluateParam('Động cơ', 'ir', formData.ir)}
            onHistoryClick={() => { setSelectedParamHistory('Điện trở cách điện'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Chỉ số phân cực (PI)" unit="" standard="≥ 2.0" prevValue="2.5" prevTrend="stable"
            value={formData.pi} onChange={(val: any) => setFormData({...formData, pi: val})}
            evalStatus={evaluateParam('Động cơ', 'pi', formData.pi)}
            onHistoryClick={() => { setSelectedParamHistory('Chỉ số phân cực (PI)'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Tổn hao điện môi (Tan-delta)" unit="" standard="≤ 0.02" prevValue="0.015" prevTrend="up"
            value={formData.tanDelta} onChange={(val: any) => setFormData({...formData, tanDelta: val})}
            evalStatus={evaluateParam('Động cơ', 'tanDelta', formData.tanDelta)}
            onHistoryClick={() => { setSelectedParamHistory('Tan-delta'); setShowHistoryModal(true); }}
          />
          <ParamInput 
            title="Phóng điện cục bộ (PD)" unit="pC" standard="≤ 1000 pC" prevValue="800 pC" prevTrend="up"
            value={formData.pd} onChange={(val: any) => setFormData({...formData, pd: val})}
            evalStatus={evaluateParam('Động cơ', 'pd', formData.pd)}
            onHistoryClick={() => { setSelectedParamHistory('Phóng điện cục bộ (PD)'); setShowHistoryModal(true); }}
          />
        </div>
      </div>
    </>
  );

  const renderSwitchgearParams = () => (
    <>
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
          <h1 className="text-2xl font-bold text-slate-900 mb-2">Hệ thống Quản lý DGA</h1>
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
    <div className="flex h-screen bg-slate-50 font-sans text-slate-900 overflow-hidden">
      
      {/* SIDEBAR */}
      <aside className="w-64 bg-slate-900 text-slate-300 flex flex-col flex-shrink-0">
        <div className="h-16 flex items-center px-6 border-b border-slate-800">
          <div className="flex items-center gap-2 text-white font-bold text-xl tracking-tight">
            <img src="https://static.wixstatic.com/media/dea24c_6207cba10c9e4e15939adcc4dbd42524~mv2.jpg" alt="TEV Logo" className="w-8 h-8 rounded-full bg-white object-contain p-0.5" />
            <span>TEV COE</span>
          </div>
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
        </nav>

        <div className="p-4 border-t border-slate-800 space-y-1">
          <button className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium hover:bg-slate-800 hover:text-white transition-colors">
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
      <main className="flex-1 flex flex-col min-w-0 overflow-hidden">
        
        {/* HEADER */}
        <header className="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-8 flex-shrink-0 z-10">
          <div className="flex items-center gap-4">
            <h1 className="text-xl font-semibold text-slate-800">
              {activeTab === 'dashboard' ? 'Dashboard Tổng quan' : 
               activeTab === 'equipment' ? 'Quản lý Thiết bị' : 
               activeTab === 'reports' ? 'Kho Báo cáo' : 
               'Nhập liệu hiện trường (FSE App)'}
            </h1>
            {activeTab !== 'field-entry' && (
              <div className="hidden md:flex items-center gap-2 px-3 py-1.5 bg-slate-100 rounded-md text-sm text-slate-600 border border-slate-200 cursor-pointer hover:bg-slate-200 transition-colors">
                <Factory size={16} className="text-slate-400" />
                <span>Tất cả nhà máy</span>
                <ChevronDown size={14} />
              </div>
            )}
          </div>
          
          <div className="flex items-center gap-4">
            {!isGoogleConnected ? (
              <button 
                onClick={handleConnectGoogle}
                className="hidden md:flex items-center gap-2 px-3 py-1.5 bg-blue-50 text-blue-600 rounded-md text-sm font-medium hover:bg-blue-100 transition-colors border border-blue-200"
              >
                <UploadCloud size={16} />
                Kết nối Google Drive
              </button>
            ) : (
              <div className="hidden md:flex items-center gap-2 px-3 py-1.5 bg-emerald-50 text-emerald-600 rounded-md text-sm font-medium border border-emerald-200">
                <CheckCircle size={16} />
                Đã kết nối Drive
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

        {/* SCROLLABLE CONTENT */}
        <div className="flex-1 overflow-y-auto p-4 md:p-8 bg-slate-100">
          
          {/* --- FIELD ENTRY VIEW --- */}
          {activeTab === 'field-entry' && (
            <div className="max-w-3xl mx-auto space-y-6 pb-12">
              
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
                    <label className="block text-sm font-medium text-slate-700 mb-1">Loại thiết bị</label>
                    <select 
                      value={selectedEqType}
                      onChange={(e) => setSelectedEqType(e.target.value)}
                      className="w-full px-4 py-2.5 bg-white border border-slate-300 focus:border-blue-500 focus:ring-2 focus:ring-blue-200 rounded-lg text-base outline-none font-medium text-slate-900"
                    >
                      <option value="Máy biến áp">Máy biến áp</option>
                      <option value="Động cơ">Động cơ điện</option>
                      <option value="Tủ điện trung thế">Tủ điện trung thế (Switchgear)</option>
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
                          <input 
                            type="text" 
                            value={customerName}
                            onChange={(e) => setCustomerName(e.target.value)}
                            placeholder="VD: Công ty Điện lực A"
                            className="w-full px-3 py-2 bg-white border border-slate-300 focus:border-blue-500 focus:ring-1 focus:ring-blue-200 rounded-md text-sm outline-none"
                          />
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
                  <div className="p-0">
                    <table className="w-full text-sm text-left">
                      <thead className="bg-slate-50 text-slate-500 border-b border-slate-200">
                        <tr>
                          <th className="px-4 py-2 font-medium">Ngày</th>
                          <th className="px-4 py-2 font-medium">Người kiểm tra</th>
                          <th className="px-4 py-2 font-medium">Trạng thái</th>
                          <th className="px-4 py-2 font-medium">Ghi chú</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {pastReports.map((report, idx) => (
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

              {/* MIDDLE SECTION: MAP & RISK LIST */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                
                {/* GEOGRAPHICAL MAP */}
                <div className="lg:col-span-2 bg-white rounded-xl border border-slate-200 shadow-sm flex flex-col overflow-hidden">
                  <div className="p-5 border-b border-slate-100 flex items-center justify-between z-10 bg-white">
                    <h2 className="font-semibold text-slate-800 flex items-center gap-2">
                      <MapPin className="text-blue-500" size={18} />
                      Bản đồ phân bố thiết bị
                    </h2>
                  </div>
                  <div className="flex-1 min-h-[400px] relative z-0">
                    <MapContainer center={[16.047079, 108.206230]} zoom={5} style={{ height: '100%', width: '100%', zIndex: 0 }}>
                      <TileLayer
                        attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
                        url="https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png"
                      />
                      {filteredSiteData.map((site) => (
                        <Marker 
                          key={site.id} 
                          position={[site.lat, site.lng]} 
                          icon={createCustomIcon(site.status, site.count)}
                        >
                          <Popup className="custom-popup">
                            <div className="p-1 min-w-[200px]">
                              <h3 className="font-bold text-slate-800 mb-2 border-b pb-2">{site.name}</h3>
                              <div className="space-y-2 text-sm">
                                <div className="flex justify-between items-center">
                                  <span className="text-slate-500">Tổng thiết bị:</span>
                                  <span className="font-bold">{site.count}</span>
                                </div>
                                <div className="flex justify-between items-center">
                                  <span className="flex items-center gap-1 text-emerald-600"><div className="w-2 h-2 rounded-full bg-emerald-500"></div> Khỏe mạnh:</span>
                                  <span className="font-medium">{site.healthy}</span>
                                </div>
                                <div className="flex justify-between items-center">
                                  <span className="flex items-center gap-1 text-amber-500"><div className="w-2 h-2 rounded-full bg-amber-500"></div> Cảnh báo:</span>
                                  <span className="font-medium">{site.warning}</span>
                                </div>
                                <div className="flex justify-between items-center">
                                  <span className="flex items-center gap-1 text-rose-500"><div className="w-2 h-2 rounded-full bg-rose-500"></div> Nguy hiểm:</span>
                                  <span className="font-medium">{site.critical}</span>
                                </div>
                              </div>
                            </div>
                          </Popup>
                          <LeafletTooltip direction="top" offset={[0, -20]} opacity={1}>
                            <span className="font-semibold">{site.name}</span>
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
                      Top rủi ro cao nhất
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
                            <HealthBar value={item.health} />
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
                          onChange={(e) => setTrendingEquipment(e.target.value)}
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
                          <option value="temp">Nhiệt độ dầu</option>
                          <option value="tev">TEV & Siêu âm (PD)</option>
                          <option value="dga">Khí hòa tan (DGA)</option>
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
                      {trendingChartType === 'temp' ? (
                        <AreaChart data={dynamicTrendData.temp} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
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
                          <Area type="monotone" dataKey="temp" name="Nhiệt độ dầu (°C)" stroke="#3b82f6" strokeWidth={3} fillOpacity={1} fill="url(#colorTemp)" activeDot={{ r: 6, strokeWidth: 0, fill: '#3b82f6' }} />
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
                      ) : (
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
                        </LineChart>
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

              {/* ALL EQUIPMENT TABLE */}
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="p-5 border-b border-slate-100 flex flex-col gap-4">
                  <div className="flex items-center justify-between">
                    <h2 className="font-semibold text-slate-800">Danh sách toàn bộ thiết bị</h2>
                    <div className="flex gap-2">
                      <button 
                        onClick={handleFetchFromSheets}
                        disabled={!isGoogleConnected || isSyncing}
                        className="text-sm bg-white border border-slate-300 text-slate-700 px-3 py-1.5 rounded-md font-medium hover:bg-slate-50 transition-colors disabled:opacity-50 flex items-center gap-2"
                      >
                        <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" /></svg>
                        Tải từ Sheets
                      </button>
                      <button 
                        onClick={handleSyncToSheets}
                        disabled={!isGoogleConnected || isSyncing}
                        className="text-sm bg-blue-50 text-blue-600 px-3 py-1.5 rounded-md font-medium hover:bg-blue-100 transition-colors disabled:opacity-50 flex items-center gap-2"
                      >
                        <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" /></svg>
                        Đồng bộ lên Sheets
                      </button>
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
                            <StatusBadge status={eq.status} />
                          </td>
                          <td className="p-4">
                            <HealthBar value={eq.health} />
                          </td>
                          <td className="p-4 text-right">
                            <button className="p-1 text-slate-400 hover:text-blue-600 transition-colors rounded">
                              <MoreVertical size={18} />
                            </button>
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
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden flex flex-col h-[calc(100vh-8rem)]">
                <div className="p-5 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                  <h2 className="font-semibold text-slate-800 flex items-center gap-2">
                    <Server className="text-blue-500" size={20} />
                    Quản lý Thiết bị
                  </h2>
                  <div className="flex items-center gap-3">
                    <div className="relative">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                      <input 
                        type="text" 
                        placeholder="Tìm mã thiết bị, tên..." 
                        className="pl-9 pr-4 py-2 bg-slate-50 border border-slate-200 focus:bg-white focus:border-blue-500 focus:ring-2 focus:ring-blue-200 rounded-lg text-sm w-full sm:w-64 transition-all outline-none"
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
                  <select className="bg-white border border-slate-200 rounded-md px-3 py-1.5 text-slate-700 outline-none focus:border-blue-500">
                    <option>Tất cả nhà máy</option>
                    <option>Nhà máy Bắc Ninh</option>
                    <option>Nhà máy Bình Dương</option>
                    <option>Nhà máy Đồng Nai</option>
                  </select>
                  <select className="bg-white border border-slate-200 rounded-md px-3 py-1.5 text-slate-700 outline-none focus:border-blue-500">
                    <option>Tất cả loại thiết bị</option>
                    <option>Máy biến áp</option>
                    <option>Động cơ điện</option>
                    <option>Tủ điện</option>
                  </select>
                  <select className="bg-white border border-slate-200 rounded-md px-3 py-1.5 text-slate-700 outline-none focus:border-blue-500">
                    <option>Tất cả trạng thái</option>
                    <option>Khỏe mạnh</option>
                    <option>Cảnh báo</option>
                    <option>Nguy hiểm</option>
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
                            <StatusBadge status={eq.status} />
                          </td>
                          <td className="p-4">
                            <HealthBar value={eq.health} />
                          </td>
                          <td className="p-4 text-slate-500">{eq.lastCheck}</td>
                          <td className="p-4 text-right">
                            <div className="flex items-center justify-end gap-2">
                              <button 
                                onClick={() => setShowEquipmentProfile(true)}
                                className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded transition-colors" title="Hồ sơ thiết bị"
                              >
                                <History size={16} />
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
              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden flex flex-col h-[calc(100vh-8rem)]">
                <div className="p-5 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                  <h2 className="font-semibold text-slate-800 flex items-center gap-2">
                    <FileText className="text-blue-500" size={20} />
                    Kho Báo cáo & Tài liệu
                  </h2>
                  <div className="flex items-center gap-3">
                    <div className="relative">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                      <input 
                        type="text" 
                        placeholder="Tìm mã báo cáo, thiết bị..." 
                        className="pl-9 pr-4 py-2 bg-slate-50 border border-slate-200 focus:bg-white focus:border-blue-500 focus:ring-2 focus:ring-blue-200 rounded-lg text-sm w-full sm:w-64 transition-all outline-none"
                      />
                    </div>
                    <button 
                      onClick={() => {
                        const headers = ['Mã TB', 'Tên thiết bị', 'Nhà máy', 'Loại', 'Trạng thái', 'Sức khỏe'];
                        const csvContent = [
                          headers.join(','),
                          ...displayedEquipment.map(eq => `${eq.id},${eq.name},${eq.factory},${eq.type},${eq.status},${eq.health}`)
                        ].join('\n');
                        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
                        const link = document.createElement('a');
                        link.href = URL.createObjectURL(blob);
                        link.download = 'danh_sach_thiet_bi.csv';
                        link.click();
                      }}
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
                  <select className="bg-white border border-slate-200 rounded-md px-3 py-1.5 text-slate-700 outline-none focus:border-blue-500">
                    <option>Tất cả nhà máy</option>
                    <option>Nhà máy Bắc Ninh</option>
                    <option>Nhà máy Bình Dương</option>
                    <option>Nhà máy Đồng Nai</option>
                  </select>
                  <select className="bg-white border border-slate-200 rounded-md px-3 py-1.5 text-slate-700 outline-none focus:border-blue-500">
                    <option>Tất cả loại thiết bị</option>
                    <option>Máy biến áp</option>
                    <option>Động cơ điện</option>
                    <option>Tủ điện</option>
                  </select>
                  <select className="bg-white border border-slate-200 rounded-md px-3 py-1.5 text-slate-700 outline-none focus:border-blue-500">
                    <option>Tất cả trạng thái</option>
                    <option>Khỏe mạnh</option>
                    <option>Cảnh báo</option>
                    <option>Nguy hiểm</option>
                  </select>
                  <div className="flex items-center gap-2 ml-auto">
                    <span className="text-slate-500">Thời gian:</span>
                    <input type="date" className="bg-white border border-slate-200 rounded-md px-3 py-1.5 text-slate-700 outline-none focus:border-blue-500" />
                    <span className="text-slate-400">-</span>
                    <input type="date" className="bg-white border border-slate-200 rounded-md px-3 py-1.5 text-slate-700 outline-none focus:border-blue-500" />
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
                      {[...pastReports, ...pastReports, ...pastReports].map((report, index) => (
                        <tr key={`${report.id}-${index}`} className="hover:bg-slate-50 transition-colors">
                          <td className="p-4 font-medium text-blue-600">{report.id}</td>
                          <td className="p-4 flex items-center gap-2"><Calendar size={14} className="text-slate-400"/> {report.date}</td>
                          <td className="p-4">
                            <div className="font-medium text-slate-900">Máy biến áp T1</div>
                            <div className="text-xs text-slate-500 font-mono">TRF-01</div>
                          </td>
                          <td className="p-4 text-slate-700">{report.inspector}</td>
                          <td className="p-4"><StatusBadge status={report.status} /></td>
                          <td className="p-4 text-slate-500 max-w-[200px] truncate" title={report.notes}>{report.notes}</td>
                          <td className="p-4 text-right">
                            <div className="flex items-center justify-end gap-2">
                              <button className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded transition-colors" title="Xem chi tiết">
                                <Eye size={16} />
                              </button>
                              <button className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded transition-colors" title="Tải PDF">
                                <Download size={16} />
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
                  <span>Hiển thị 1-9 trong số 45 báo cáo</span>
                  <div className="flex gap-1">
                    <button className="px-3 py-1 border border-slate-200 rounded hover:bg-slate-50 disabled:opacity-50" disabled>Trước</button>
                    <button className="px-3 py-1 border border-slate-200 rounded bg-blue-50 text-blue-600 font-medium">1</button>
                    <button className="px-3 py-1 border border-slate-200 rounded hover:bg-slate-50">2</button>
                    <button className="px-3 py-1 border border-slate-200 rounded hover:bg-slate-50">3</button>
                    <button className="px-3 py-1 border border-slate-200 rounded hover:bg-slate-50">...</button>
                    <button className="px-3 py-1 border border-slate-200 rounded hover:bg-slate-50">5</button>
                    <button className="px-3 py-1 border border-slate-200 rounded hover:bg-slate-50">Sau</button>
                  </div>
                </div>
              </div>
            </div>
          )}
          {activeTab === 'deep-analysis' && (
            <div className="flex-1 overflow-y-auto p-8 bg-slate-50/50">
              <div className="max-w-5xl mx-auto space-y-6">
                <div className="flex items-center justify-between">
                  <div>
                    <h2 className="text-2xl font-bold text-slate-900 tracking-tight">Phân tích chuyên sâu</h2>
                    <p className="text-slate-500 mt-1">Phân tích tình trạng máy biến áp dựa trên dữ liệu DGA</p>
                  </div>
                  {dgaAnalysisResult && (
                    <button 
                      onClick={exportToPDF}
                      className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg font-medium hover:bg-blue-700 transition-colors shadow-sm"
                    >
                      <Download size={18} />
                      Xuất báo cáo PDF
                    </button>
                  )}
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                  {/* Input Form */}
                  <div className="lg:col-span-1 bg-white rounded-xl shadow-sm border border-slate-200 p-6">
                    <h3 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
                      <Thermometer size={20} className="text-blue-500" />
                      Nhập liệu (ppm)
                    </h3>
                    <div className="space-y-4">
                      <div className="grid grid-cols-2 gap-2">
                        {['h2', 'ch4', 'c2h6', 'c2h4', 'c2h2', 'co', 'co2', 'o2', 'n2'].map((gas) => (
                          <div key={gas}>
                            <label className="block text-xs font-medium text-slate-700 mb-1 uppercase">{gas}</label>
                            <input
                              type="number"
                              value={dgaData[gas as keyof typeof dgaData]}
                              onChange={(e) => setDgaData({...dgaData, [gas]: e.target.value})}
                              className="w-full px-2 py-1.5 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all"
                              placeholder={`Nhập ${gas.toUpperCase()}`}
                            />
                          </div>
                        ))}
                      </div>
                      <div className="border-t border-slate-200 pt-4 mt-4">
                        <h4 className="text-sm font-bold text-slate-800 mb-3">Thông số khác</h4>
                        <div className="grid grid-cols-2 gap-2">
                          <div>
                            <label className="block text-xs font-medium text-slate-700 mb-1">Moisture (ppm)</label>
                            <input type="number" value={dgaData.moisture} onChange={(e) => setDgaData({...dgaData, moisture: e.target.value})} className="w-full px-2 py-1.5 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" />
                          </div>
                          <div>
                            <label className="block text-xs font-medium text-slate-700 mb-1">BD Strength (kV)</label>
                            <input type="number" value={dgaData.bdStrength} onChange={(e) => setDgaData({...dgaData, bdStrength: e.target.value})} className="w-full px-2 py-1.5 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" />
                          </div>
                          <div>
                            <label className="block text-xs font-medium text-slate-700 mb-1">Acidity (mgKOH/g)</label>
                            <input type="number" value={dgaData.acidity} onChange={(e) => setDgaData({...dgaData, acidity: e.target.value})} className="w-full px-2 py-1.5 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" />
                          </div>
                          <div>
                            <label className="block text-xs font-medium text-slate-700 mb-1">FFA (ppm)</label>
                            <input type="number" value={dgaData.ffa} onChange={(e) => setDgaData({...dgaData, ffa: e.target.value})} className="w-full px-2 py-1.5 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" />
                          </div>
                          <div>
                            <label className="block text-xs font-medium text-slate-700 mb-1">Est DP</label>
                            <input type="number" value={dgaData.estDp} onChange={(e) => setDgaData({...dgaData, estDp: e.target.value})} className="w-full px-2 py-1.5 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" />
                          </div>
                          <div>
                            <label className="block text-xs font-medium text-slate-700 mb-1">Tuổi MBA (năm)</label>
                            <input type="number" value={dgaData.age} onChange={(e) => setDgaData({...dgaData, age: e.target.value})} className="w-full px-2 py-1.5 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" />
                          </div>
                          <div>
                            <label className="block text-xs font-medium text-slate-700 mb-1">% Phụ tải</label>
                            <input type="number" value={dgaData.loadFactor} onChange={(e) => setDgaData({...dgaData, loadFactor: e.target.value})} className="w-full px-2 py-1.5 text-sm border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" />
                          </div>
                        </div>
                      </div>
                      <button 
                        onClick={analyzeDGA}
                        className="w-full mt-6 py-2.5 bg-slate-900 text-white rounded-lg font-medium hover:bg-slate-800 transition-colors shadow-sm"
                      >
                        Phân tích dữ liệu
                      </button>
                    </div>
                  </div>

                  {/* Analysis Results */}
                  <div className="lg:col-span-2 space-y-6">
                    {dgaAnalysisResult ? (
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
                      <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-12 flex flex-col items-center justify-center text-center h-full min-h-[400px]">
                        <div className="w-16 h-16 bg-blue-50 text-blue-500 rounded-full flex items-center justify-center mb-4">
                          <Activity size={32} />
                        </div>
                        <h3 className="text-xl font-bold text-slate-800 mb-2">Chưa có dữ liệu phân tích</h3>
                        <p className="text-slate-500 max-w-md">
                          Vui lòng nhập các giá trị khí hòa tan (DGA) ở cột bên trái và nhấn "Phân tích dữ liệu" để xem kết quả chẩn đoán và dự báo tuổi thọ.
                        </p>
                      </div>
                    )}
                  </div>
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
                  currentDetailedRiskData.status === 'critical' ? 'bg-rose-100 text-rose-600' : 'bg-amber-100 text-amber-600'
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
                      currentDetailedRiskData.status === 'critical' ? 'text-rose-600' : 'text-amber-500'
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
                          <Radar name="Risk" dataKey="score" stroke={currentDetailedRiskData.status === 'critical' ? '#f43f5e' : '#3b82f6'} strokeWidth={2} fill={currentDetailedRiskData.status === 'critical' ? '#f43f5e' : '#3b82f6'} fillOpacity={0.3} />
                        </RadarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>

                  {/* Defects Table */}
                  <div className="p-6 lg:col-span-3 flex flex-col">
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse">
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
                  <table className="w-full text-left border-collapse">
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

              {/* Past Reports */}
              <div>
                <h3 className="text-base font-bold text-slate-800 mb-4 flex items-center gap-2">
                  <FileText size={18} className="text-blue-500" />
                  Danh sách báo cáo cũ
                </h3>
                <div className="border border-slate-200 rounded-xl overflow-hidden">
                  <table className="w-full text-left border-collapse">
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
                      {pastReports.map((report) => (
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
                              <button className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded transition-colors" title="Tải PDF">
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
                <input type="text" value={newEqData.customer} onChange={(e) => setNewEqData({...newEqData, customer: e.target.value})} className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder="VD: Công ty Điện lực A" />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Nhà máy</label>
                <input type="text" value={newEqData.factory} onChange={(e) => setNewEqData({...newEqData, factory: e.target.value})} className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500" placeholder="VD: Nhà máy Bắc Ninh" />
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
              {paramHistoryData[selectedParamHistory as keyof typeof paramHistoryData] ? (
                <div className="h-64 w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={paramHistoryData[selectedParamHistory as keyof typeof paramHistoryData]} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
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
    </div>
  );
}

