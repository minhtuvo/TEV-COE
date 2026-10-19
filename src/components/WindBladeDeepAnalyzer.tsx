import React, { useState, useMemo } from 'react';
import {
  Wind,
  Shield,
  ShieldAlert,
  ShieldCheck,
  AlertTriangle,
  AlertCircle,
  XCircle,
  CheckCircle2,
  Sparkles,
  Download,
  Printer,
  Copy,
  Check,
  RotateCcw,
  Activity,
  Layers,
  Zap,
  Clock,
  ArrowRight,
  Eye,
  FileSpreadsheet,
  FileText,
  Sliders,
  ChevronDown,
  ChevronUp,
  Maximize2,
  Wrench,
  HelpCircle,
  PlusCircle,
  TrendingUp,
  MapPin,
  Calendar,
  Flame,
  Info
} from 'lucide-react';
import {
  EpriBladeCategory,
  BladeChecklistItem,
  WindBladeInputPayload,
  WindBladeAnalysisResponse,
  WindBladeAnalysisJsonData,
  evaluateWindBladeDeterministic
} from '../../server/windBladeAnalyzer';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
  onCreateWorkOrder?: (woData: any) => void;
}

// ----------------------------------------------------------------------------
// EPRI / DNV GL DAMAGE CATEGORIZATION REFERENCE DATA
// ----------------------------------------------------------------------------
export const EPRI_CATEGORIES_INFO: Record<
  EpriBladeCategory,
  {
    category: EpriBladeCategory;
    titleVn: string;
    titleEn: string;
    badgeClass: string;
    borderClass: string;
    bgClass: string;
    descVn: string;
    descEn: string;
    impactVn: string;
    impactEn: string;
    actionVn: string;
    actionEn: string;
  }
> = {
  'Category 1': {
    category: 'Category 1',
    titleVn: 'Sai lệch Thẩm mỹ Nhẹ',
    titleEn: 'Minor Cosmetic Variance',
    badgeClass: 'bg-emerald-100 text-emerald-800 border-emerald-300',
    borderClass: 'border-emerald-200',
    bgClass: 'bg-emerald-50/40',
    descVn: 'Bụi bẩn, dính dầu mỡ, côn trùng, trầy xước sơn nhẹ, thiếu nhãn dán mác.',
    descEn: 'Dirt, grease, light surface scratches, missing stickers.',
    impactVn: 'Không ảnh hưởng đến kết cấu và khí động học.',
    impactEn: 'No impact on structural integrity or aerodynamics.',
    actionVn: 'Vận hành bình thường; không cần sửa chữa.',
    actionEn: 'Continue operation; no repair required.'
  },
  'Category 2': {
    category: 'Category 2',
    titleVn: 'Tổn hại Bề mặt Nhẹ',
    titleEn: 'Minor Cosmetic & Surface Damage',
    badgeClass: 'bg-blue-100 text-blue-800 border-blue-300',
    borderClass: 'border-blue-200',
    bgClass: 'bg-blue-50/40',
    descVn: 'Tróc sơn nhẹ, xói mòn mép trước giai đoạn đầu, xước bề mặt gelcoat, rải rác mất một vài đĩa VGs.',
    descEn: 'Minor gelcoat cracks, early stage LE erosion, few missing VGs.',
    impactVn: 'Ảnh hưởng khí động học không đáng kể; không giảm tuổi thọ cánh.',
    impactEn: 'Minor aerodynamic impact; no reduction in blade life.',
    actionVn: 'Vận hành & Theo dõi trong các kỳ bảo dưỡng định kỳ (6 - 12 tháng).',
    actionEn: 'Continue operation; monitor during routine maintenance (6-12 months).'
  },
  'Category 3': {
    category: 'Category 3',
    titleVn: 'Hư hỏng Kết cấu Nhẹ / Trung bình',
    titleEn: 'Moderate Structural & Aerodynamic Damage',
    badgeClass: 'bg-amber-100 text-amber-900 border-amber-300',
    borderClass: 'border-amber-200',
    bgClass: 'bg-amber-50/50',
    descVn: 'Nứt nhẹ bề mặt, bong tróc dán keo mép trước/mép sau nhỏ, xói mòn lộ lớp sợi, mất nhiều VGs.',
    descEn: 'Small shell/bondline cracks, progressing LE erosion, missing VG strips.',
    impactVn: 'Có nguy cơ giảm tuổi thọ cánh nếu không xử lý; phát triển vết nứt trung bình.',
    impactEn: 'Life reduction expected if left unaddressed; moderate growth potential.',
    actionVn: 'Vận hành & Lập kế hoạch Sửa chữa trong vòng 6 - 12 tháng; kiểm tra định kỳ 6 tháng.',
    actionEn: 'Continue operation with 6-month monitoring; repair within 6-12 months.'
  },
  'Category 4': {
    category: 'Category 4',
    titleVn: 'Hư hỏng Kết cấu Nghiêm trọng',
    titleEn: 'Significant Structural Defect',
    badgeClass: 'bg-orange-100 text-orange-900 border-orange-300',
    borderClass: 'border-orange-300',
    bgClass: 'bg-orange-50/50',
    descVn: 'Nứt sâu sợi composite, tách mép sau lớn lộ cấu trúc bên trong, rỗng lớp foam lõi, đứt bu-lông gốc cánh, đĩa LPS hư hỏng nặng.',
    descEn: 'Deep structural cracks, extensive TE separation, core gaps, broken root bolts, severe LPS damage.',
    impactVn: 'Ảnh hưởng lớn đến khả năng chịu lực và hiệu suất khí động; nguy cơ không đạt tuổi thọ thiết kế.',
    impactEn: 'High impact on load capacity and aerodynamics; high risk of failing design life.',
    actionVn: 'Đánh giá Kỹ thuật khẩn cấp; bắt buộc sửa chữa trong vòng 1 - 3 tháng; tăng tần suất giám sát hàng tháng.',
    actionEn: 'Require engineering review; repair within 1-3 months with monthly monitoring.'
  },
  'Category 5': {
    category: 'Category 5',
    titleVn: 'Nguy cơ Sụp đổ Cấu trúc',
    titleEn: 'Severe Damage / Imminent Failure Risk',
    badgeClass: 'bg-rose-100 text-rose-900 border-rose-400 font-bold',
    borderClass: 'border-rose-400',
    bgClass: 'bg-rose-50/60',
    descVn: 'Nứt đứt dầm chịu lực (spar cap/shear web buckling), tách rời vách ngăn dầm, nứt gãy sườn cánh rộng, bong tách vỏ cánh lớn.',
    descEn: 'Buckled shear webs, spar cap delamination, major chordwise/spanwise structural cracks.',
    impactVn: 'Khả năng gãy cánh/sụp đổ cấu trúc RẤT CAO khi tiếp tục quay.',
    impactEn: 'High risk of imminent structural failure during operation.',
    actionVn: 'DỪNG MÁY KHẨN CẤP (SHUTDOWN IMMEDIATELY); ngắt vận hành cho đến khi hoàn tất sửa chữa hoặc thay cánh mới.',
    actionEn: 'Emergency Shutdown Immediately; do not operate until repaired/replaced.'
  }
};

// ----------------------------------------------------------------------------
// DEFAULT BENCHMARK SAMPLE (USER'S EXACT SPECIFICATION - CATEGORY 4 CASE)
// ----------------------------------------------------------------------------
export const BENCHMARK_CAT4_SAMPLE: WindBladeInputPayload = {
  turbineId: 'WTG-08',
  bladeTag: 'BLADE-A-75M',
  bladeModel: 'LM 75.2 P2 / Vestas V150',
  windFarmName: 'Nhà máy Điện Gió Mũi Né (Bình Thuận)',
  inspectionDate: '2026-10-16',
  inspectorName: 'KS. Hoàng Đình Trọng (Lead FSE NETA / Drone Pilot)',
  operatingHours: 18450,
  previousTeCrackLengthCm: 20, // 20cm in last year
  vibrationRmsMmS: 1.45,
  acousticDb: 48.5,
  rotationRpm: 11.2,
  windSpeedMs: 9.8,
  checklistItems: [
    {
      id: 'A1',
      nameVn: 'Mép Trước Cánh',
      nameEn: 'Leading Edge (LE) Condition',
      locationSpanM: '45m - 55m',
      defectSignVn: 'Xói mòn lớp phủ lộ sợi composite (LE erosion)',
      defectSignEn: 'Progressing erosion exposing fiberglass laminate',
      status: 'Investigate',
      severityCategory: 'Category 3',
      defectLengthCm: 450,
      detailsVn: 'Xói mòn lộ lớp sợi composite tại dải 45m-55m do mưa gió cát ngoài khơi, gelcoat bị bong tróc từng mảng.',
      detailsEn: 'Progressing leading edge erosion from 45m to 55m exposing fiberglass fabric, coating peeled off.'
    },
    {
      id: 'A2',
      nameVn: 'Mép Sau & Đường Dán Keo',
      nameEn: 'Trailing Edge (TE) & Bondline',
      locationSpanM: '60m',
      defectSignVn: 'Tách keo mép sau dài 1.8m lộ lõi xốp (Split TE Seam)',
      defectSignEn: 'Split TE bondline (1.8m) exposing structural core',
      status: 'Fail',
      severityCategory: 'Category 4',
      defectLengthCm: 180,
      detailsVn: 'Đường dán keo mép sau bị tách toác dài 1.8m tại vị trí 60m cách gốc, nước ẩm xâm nhập vào lõi balsa/foam.',
      detailsEn: '1.8-meter split trailing edge seam at 60m span, exposing interior sandwich core to moisture.'
    },
    {
      id: 'A3',
      nameVn: 'Vỏ Bán Cánh & Ma Sát Mặt',
      nameEn: 'Shell Surfaces & Laminate',
      locationSpanM: '35m (Pressure Side)',
      defectSignVn: 'Vết nứt dọc lực dài 85cm gần dầm Spar Cap',
      defectSignEn: 'Spanwise crack (85cm) near main Spar Cap region',
      status: 'Fail',
      severityCategory: 'Category 4',
      defectLengthCm: 85,
      detailsVn: 'Vết nứt dọc lực dài 85cm tại mặt nén (Pressure Side) cách gốc 35m, nằm sát mép dầm chính Spar Cap.',
      detailsEn: '85cm spanwise structural crack on pressure side at 35m span, adjacent to main spar cap.'
    },
    {
      id: 'A4',
      nameVn: 'Hệ Thống Chống Sét Cánh (LPS)',
      nameEn: 'Lightning Protection System (LPS)',
      locationSpanM: 'Tip Receptor #3',
      defectSignVn: 'Đĩa thu sét chảy nổ, thủng vỏ cánh 15cm',
      defectSignEn: 'Melted tip receptor with 15cm burn hole in shell',
      status: 'Fail',
      severityCategory: 'Category 4',
      defectLengthCm: 15,
      detailsVn: 'Đĩa thu sét số 3 tại mũi cánh bị phóng hồ quang chảy sém, thủng vỏ composite đường kính 15cm quanh đĩa.',
      detailsEn: 'Tip Receptor #3 shows severe arc melting damage with a 15cm puncture hole in surrounding shell.'
    },
    {
      id: 'A5',
      nameVn: 'Thiết Bị Khí Động Học Phụ Trợ',
      nameEn: 'Aerodynamic Devices',
      locationSpanM: 'Gốc cánh 10m - 22m',
      defectSignVn: 'Rơi mất 12 khối cánh tạo xoáy VG (Vortex Generators)',
      defectSignEn: '12 Vortex Generator (VG) blocks detached/missing',
      status: 'Investigate',
      severityCategory: 'Category 2',
      defectLengthCm: 0,
      detailsVn: 'Rơi rụng mất 12 khối VG tại dải gốc cánh từ 10m đến 22m, suy giảm lực nâng khí động.',
      detailsEn: '12 VG blocks detached along the root transition zone (10m-22m), causing minor aerodynamic drag.'
    }
  ],
  fseVisualNotes: 'Cánh có dấu hiệu chịu xung lực sét đánh và tải trọng gió mỏi cao. Cần khẩn cấp đánh giá derating giảm tải và lập phương án đu dây sửa chữa trong vòng 30-60 ngày.'
};

export const WindBladeDeepAnalyzer: React.FC<Props> = ({
  onSaveReport,
  onNavigateToReports,
  onCreateWorkOrder
}) => {
  // Language mode: 'bilingual' | 'vn' | 'en'
  const [langMode, setLangMode] = useState<'bilingual' | 'vn' | 'en'>('bilingual');

  // Active view: 'checklist-input' | 'analysis-result' | 'matrix-reference'
  const [activeTab, setActiveTab] = useState<'checklist-input' | 'analysis-result' | 'matrix-reference'>('checklist-input');

  // Input state initialized with user's exact specification
  const [formData, setFormData] = useState<WindBladeInputPayload>(BENCHMARK_CAT4_SAMPLE);

  // Analysis result state
  const [analysisResult, setAnalysisResult] = useState<WindBladeAnalysisResponse | null>(() => {
    return evaluateWindBladeDeterministic(BENCHMARK_CAT4_SAMPLE);
  });

  const [isLoadingAi, setIsLoadingAi] = useState(false);
  const [isCopied, setIsCopied] = useState(false);
  const [saveSuccessNotice, setSaveSuccessNotice] = useState<string | null>(null);

  // Active highlighted defect for blade diagram
  const [hoveredDefectId, setHoveredDefectId] = useState<string | null>(null);

  // Quick preset loader
  const handleLoadPreset = (type: 'cat4_benchmark' | 'cat5_shutdown' | 'cat3_erosion' | 'cat1_healthy') => {
    if (type === 'cat4_benchmark') {
      setFormData(BENCHMARK_CAT4_SAMPLE);
      setAnalysisResult(evaluateWindBladeDeterministic(BENCHMARK_CAT4_SAMPLE));
    } else if (type === 'cat5_shutdown') {
      const cat5Data: WindBladeInputPayload = {
        ...BENCHMARK_CAT4_SAMPLE,
        bladeTag: 'BLADE-B-75M',
        vibrationRmsMmS: 2.85,
        acousticDb: 64.0,
        checklistItems: [
          ...BENCHMARK_CAT4_SAMPLE.checklistItems.map(item => {
            if (item.id === 'A3') {
              return {
                ...item,
                defectSignVn: 'Nứt gãy dầm chịu lực Spar Cap & phồng rộp vỏ cánh (Buckling)',
                defectSignEn: 'Spar Cap structural failure and severe shear web buckling',
                status: 'Fail' as const,
                severityCategory: 'Category 5' as EpriBladeCategory,
                defectLengthCm: 240,
                detailsVn: 'Nứt toác dầm chính Spar Cap 2.4m, sập vách Shear Web bên trong, cánh có nguy cơ gãy đứt khi quay.',
                detailsEn: 'Major 2.4m fracture through main spar cap with buckled shear web. Catastrophic blade throw risk.'
              };
            }
            return item;
          })
        ],
        fseVisualNotes: 'Phát hiện dầm chính Spar Cap bị nứt toác 2.4m. Yêu cầu dừng máy khẩn cấp ngay lập tức!'
      };
      setFormData(cat5Data);
      setAnalysisResult(evaluateWindBladeDeterministic(cat5Data));
    } else if (type === 'cat3_erosion') {
      const cat3Data: WindBladeInputPayload = {
        ...BENCHMARK_CAT4_SAMPLE,
        bladeTag: 'BLADE-C-75M',
        vibrationRmsMmS: 0.95,
        acousticDb: 42.0,
        previousTeCrackLengthCm: 15,
        checklistItems: BENCHMARK_CAT4_SAMPLE.checklistItems.map(item => {
          if (item.id === 'A2') {
            return {
              ...item,
              defectLengthCm: 25,
              severityCategory: 'Category 3' as EpriBladeCategory,
              status: 'Investigate' as const,
              detailsVn: 'Vết tách mép sau nhỏ dài 25cm, chưa lộ sâu vào lõi foam.',
              detailsEn: 'Minor 25cm TE split seam, superficial bondline void.'
            };
          }
          if (item.id === 'A3') {
            return {
              ...item,
              severityCategory: 'Category 2' as EpriBladeCategory,
              status: 'Pass' as const,
              defectLengthCm: 0,
              detailsVn: 'Vỏ cánh bề mặt nguyên vẹn, trầy xước sơn nhẹ.',
              detailsEn: 'Intact composite shell, minor surface paint scuffs.'
            };
          }
          if (item.id === 'A4') {
            return {
              ...item,
              severityCategory: 'Category 2' as EpriBladeCategory,
              status: 'Pass' as const,
              defectLengthCm: 0,
              detailsVn: 'Đĩa thu sét có vết cháy xém nhẹ, dây dẫn sét liên tục.',
              detailsEn: 'Minor receptor scorching, electrical continuity verified.'
            };
          }
          return item;
        }),
        fseVisualNotes: 'Cánh bị xói mòn mép trước trung bình và tách keo nhỏ. Tiếp tục vận hành và sửa chữa trong vòng 6-12 tháng.'
      };
      setFormData(cat3Data);
      setAnalysisResult(evaluateWindBladeDeterministic(cat3Data));
    } else if (type === 'cat1_healthy') {
      const healthyData: WindBladeInputPayload = {
        ...BENCHMARK_CAT4_SAMPLE,
        bladeTag: 'BLADE-HEALTHY-01',
        vibrationRmsMmS: 0.42,
        acousticDb: 32.5,
        checklistItems: BENCHMARK_CAT4_SAMPLE.checklistItems.map(item => ({
          ...item,
          status: 'Pass' as const,
          severityCategory: 'Category 1' as EpriBladeCategory,
          defectLengthCm: 0,
          detailsVn: 'Bề mặt sạch sẽ, lớp gelcoat bóng, đĩa thu sét và cánh tạo xoáy nguyên vẹn.',
          detailsEn: 'Clean surface, glossy gelcoat, intact lightning receptors and vortex generators.'
        })),
        fseVisualNotes: 'Cánh quạt trong tình trạng hoàn hảo, không phát hiện vết nứt hay xói mòn kết cấu.'
      };
      setFormData(healthyData);
      setAnalysisResult(evaluateWindBladeDeterministic(healthyData));
    }
  };

  // Run AI analysis
  const handleRunAiAnalysis = async () => {
    setIsLoadingAi(true);
    try {
      const res = await fetch('/api/wind-blade-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(formData)
      });
      if (res.ok) {
        const data: WindBladeAnalysisResponse = await res.json();
        setAnalysisResult(data);
        setActiveTab('analysis-result');
      } else {
        // Fallback to client deterministic
        const fallback = evaluateWindBladeDeterministic(formData);
        setAnalysisResult(fallback);
        setActiveTab('analysis-result');
      }
    } catch (e) {
      console.warn('Network call failed, using deterministic evaluation:', e);
      const fallback = evaluateWindBladeDeterministic(formData);
      setAnalysisResult(fallback);
      setActiveTab('analysis-result');
    } finally {
      setIsLoadingAi(false);
    }
  };

  // Copy AI Studio prompt
  const handleCopyPrompt = () => {
    const promptText = `// SYSTEM INSTRUCTIONS DÀNH CHO GOOGLE AI STUDIO (EPRI 3002019669 / DNV GL WIND BLADE)
You are an expert wind turbine blade inspection AI adhering strictly to EPRI 3002019669, DNVGL-ST-0376, IEC 61400-5, and IEC 61400-24.
Evaluate the following blade checklist input and generate a bilingual EN/VN assessment:

${JSON.stringify(formData, null, 2)}`;
    navigator.clipboard.writeText(promptText);
    setIsCopied(true);
    setTimeout(() => setIsCopied(false), 2500);
  };

  // Save Report Handler
  const handleSaveToReports = () => {
    if (!analysisResult) return;
    const reportRecord = {
      id: `REP-WIND-${Date.now().toString().slice(-6)}`,
      date: formData.inspectionDate || new Date().toISOString().split('T')[0],
      customer: formData.windFarmName || 'Nhà máy Điện Gió Mũi Né',
      factory: formData.windFarmName || 'Trạm Điện Gió',
      location: `Trụ ${formData.turbineId || 'WTG-08'} - Cánh ${formData.bladeTag || 'BLADE-A-75M'}`,
      equipmentId: formData.bladeTag || 'WIND-BLADE-01',
      equipmentName: `Cánh quạt tua-bin gió (${formData.bladeModel || 'LM 75.2 P2 75m'})`,
      type: 'Tua-bin gió (Wind Blade)',
      inspector: formData.inspectorName || 'TEV FSE Team',
      status: analysisResult.overall_status === 'PASS' ? 'healthy' : analysisResult.overall_status === 'INVESTIGATE' ? 'warning' : 'critical',
      health: analysisResult.max_category === 'Category 1' ? 98 : analysisResult.max_category === 'Category 2' ? 88 : analysisResult.max_category === 'Category 3' ? 74 : analysisResult.max_category === 'Category 4' ? 45 : 15,
      notes: `${analysisResult.json_data.summary_vn} (EPRI ${analysisResult.max_category})`,
      measurements: {
        maxSeverity: analysisResult.max_category,
        overallStatus: analysisResult.overall_status,
        vibrationRms: formData.vibrationRmsMmS,
        acousticDb: formData.acousticDb,
        teSplitLengthCm: formData.checklistItems.find(i => i.id === 'A2')?.defectLengthCm || 0
      },
      markdownReport: analysisResult.markdown_report
    };

    if (onSaveReport) {
      onSaveReport(reportRecord);
    }
    setSaveSuccessNotice(`Đã lưu báo cáo "${reportRecord.id}" vào Kho Báo Cáo!`);
    setTimeout(() => setSaveSuccessNotice(null), 3500);
  };

  // Create Work Order directly from this inspection
  const handleCreateWorkOrder = () => {
    if (!analysisResult) return;
    const priCode = analysisResult.json_data.recommendations.recommended_work_order_priority || 'P1';
    const newWo = {
      title: `[EPRI ${analysisResult.max_category}] Khắc phục hư hỏng cánh quạt ${formData.bladeTag} - Trụ ${formData.turbineId}`,
      description: `Đội đu dây (Rope-Access): ${analysisResult.json_data.recommendations.uptower_repair_protocol_vn}`,
      equipmentId: formData.bladeTag || 'WIND-BLADE-01',
      equipmentName: `Cánh quạt ${formData.bladeModel || 'LM 75.2 P2 75m'} - Trụ ${formData.turbineId}`,
      customer: formData.windFarmName || 'Nhà máy Điện Gió Mũi Né',
      priorityCode: priCode,
      workStream: 'CM',
      workflowStage: priCode === 'P1' ? '1_wr_pending' : '2_daily_approved',
      blockingRequired: true,
      deEnergyRequired: true,
      failureCode: 'WIND-BLADE-FATIGUE'
    };

    if (onCreateWorkOrder) {
      onCreateWorkOrder(newWo);
    } else {
      setSaveSuccessNotice(`Đã kích hoạt tạo Work Order (Mức ưu tiên ${priCode}) cho đội sửa chữa trên cao!`);
      setTimeout(() => setSaveSuccessNotice(null), 3500);
    }
  };

  // Trailing edge growth calculation
  const prevTe = formData.previousTeCrackLengthCm || 20;
  const currentTe = formData.checklistItems.find(i => i.id === 'A2')?.defectLengthCm || 180;
  const growthRate = prevTe > 0 ? Math.round(((currentTe - prevTe) / prevTe) * 100) : 0;

  return (
    <div className="space-y-6">
      {/* ---------------------------------------------------------------------- */}
      {/* HEADER & CONTROLS BANNER */}
      {/* ---------------------------------------------------------------------- */}
      <div className="bg-gradient-to-r from-slate-900 via-slate-800 to-indigo-950 text-white rounded-2xl p-5 sm:p-7 shadow-lg border border-slate-700/60 relative overflow-hidden">
        {/* Subtle Wind Wave Ambient */}
        <div className="absolute right-0 top-0 bottom-0 w-96 opacity-10 pointer-events-none flex items-center justify-center">
          <Wind size={260} className="text-white transform translate-x-12" />
        </div>

        <div className="relative z-10 flex flex-col lg:flex-row lg:items-center justify-between gap-5">
          <div>
            <div className="flex items-center gap-2.5 flex-wrap mb-2">
              <span className="px-2.5 py-0.5 rounded-full text-xs font-bold bg-blue-500/30 text-blue-300 border border-blue-400/30 flex items-center gap-1.5">
                <Wind size={13} />
                EPRI 3002019669
              </span>
              <span className="px-2.5 py-0.5 rounded-full text-xs font-bold bg-indigo-500/30 text-indigo-300 border border-indigo-400/30">
                DNVGL-ST-0376
              </span>
              <span className="px-2.5 py-0.5 rounded-full text-xs font-bold bg-teal-500/30 text-teal-300 border border-teal-400/30">
                IEC 61400-5 / 24
              </span>
              <span className="px-2 py-0.5 rounded text-[11px] font-semibold bg-white/10 text-slate-200">
                Song ngữ EN / VN
              </span>
            </div>
            <h2 className="text-xl sm:text-2xl font-black tracking-tight text-white flex items-center gap-2.5">
              Phân Tích Chuyên Sâu Cánh Quạt Gió (Wind Turbine Blade)
            </h2>
            <p className="text-xs sm:text-sm text-slate-300 mt-1 max-w-2xl">
              Hệ thống đánh giá hư hại cánh quạt gió bằng Drone & Hình ảnh thị giác theo 5 cấp độ EPRI / DNV GL. Tự động tính toán tốc độ phát triển nứt mỏi và phác thảo biện pháp thi công trên cao (Uptower Repair).
            </p>
          </div>

          {/* Quick Actions & Language Mode Switch */}
          <div className="flex items-center gap-2 flex-wrap">
            {/* Language Switch */}
            <div className="bg-slate-800/90 p-1 rounded-xl border border-slate-700 flex items-center shadow-inner">
              <button
                type="button"
                onClick={() => setLangMode('bilingual')}
                className={`px-2.5 py-1 rounded-lg text-xs font-bold transition-all ${langMode === 'bilingual' ? 'bg-blue-600 text-white shadow' : 'text-slate-400 hover:text-white'}`}
                title="Hiển thị song ngữ Anh - Việt"
              >
                Song ngữ (EN/VN)
              </button>
              <button
                type="button"
                onClick={() => setLangMode('vn')}
                className={`px-2.5 py-1 rounded-lg text-xs font-bold transition-all ${langMode === 'vn' ? 'bg-blue-600 text-white shadow' : 'text-slate-400 hover:text-white'}`}
                title="Chỉ hiển thị tiếng Việt"
              >
                Tiếng Việt
              </button>
              <button
                type="button"
                onClick={() => setLangMode('en')}
                className={`px-2.5 py-1 rounded-lg text-xs font-bold transition-all ${langMode === 'en' ? 'bg-blue-600 text-white shadow' : 'text-slate-400 hover:text-white'}`}
                title="English Only"
              >
                English
              </button>
            </div>

            <button
              onClick={handleCopyPrompt}
              className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 rounded-xl text-xs font-semibold flex items-center gap-1.5 transition-colors shadow-sm"
              title="Sao chép Prompt & JSON đầu vào chuẩn hóa Google AI Studio"
            >
              {isCopied ? <Check size={14} className="text-emerald-400" /> : <Copy size={14} />}
              <span>{isCopied ? 'Đã sao chép!' : 'JSON / Prompt'}</span>
            </button>
          </div>
        </div>

        {/* Tab Navigation */}
        <div className="flex items-center gap-2 mt-6 pt-4 border-t border-slate-700/60 overflow-x-auto">
          <button
            onClick={() => setActiveTab('checklist-input')}
            className={`px-4 py-2 rounded-xl text-xs sm:text-sm font-bold flex items-center gap-2 transition-all ${
              activeTab === 'checklist-input'
                ? 'bg-white text-slate-900 shadow-md font-black'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <Sliders size={16} />
            <span>Phần A: Checklist Drone & Hiện trường (5 Hạng Mục)</span>
          </button>
          <button
            onClick={() => setActiveTab('analysis-result')}
            className={`px-4 py-2 rounded-xl text-xs sm:text-sm font-bold flex items-center gap-2 transition-all ${
              activeTab === 'analysis-result'
                ? 'bg-blue-600 text-white shadow-md font-black'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <Activity size={16} />
            <span>Kết Quả Phân Tích & Hành Động Sửa Chữa</span>
            {analysisResult && (
              <span className={`text-[10px] px-1.5 py-0.2 rounded font-black ${
                analysisResult.max_category === 'Category 5' ? 'bg-rose-500 text-white' :
                analysisResult.max_category === 'Category 4' ? 'bg-orange-500 text-white' : 'bg-emerald-500 text-white'
              }`}>
                {analysisResult.max_category}
              </span>
            )}
          </button>
          <button
            onClick={() => setActiveTab('matrix-reference')}
            className={`px-4 py-2 rounded-xl text-xs sm:text-sm font-bold flex items-center gap-2 transition-all ${
              activeTab === 'matrix-reference'
                ? 'bg-white text-slate-900 shadow-md font-black'
                : 'text-slate-300 hover:bg-slate-800/80 hover:text-white'
            }`}
          >
            <Layers size={16} />
            <span>Phần B: Ma Trận EPRI / DNV GL (Cat 1 - Cat 5)</span>
          </button>
        </div>
      </div>

      {/* Save Success Banner */}
      {saveSuccessNotice && (
        <div className="p-3.5 bg-emerald-50 border border-emerald-200 rounded-xl flex items-center justify-between text-xs sm:text-sm text-emerald-800 animate-in fade-in duration-200 shadow-sm">
          <div className="flex items-center gap-2">
            <CheckCircle2 size={18} className="text-emerald-600 shrink-0" />
            <span className="font-semibold">{saveSuccessNotice}</span>
          </div>
          {onNavigateToReports && (
            <button
              onClick={onNavigateToReports}
              className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-bold transition-colors"
            >
              Xem Kho Báo Cáo &rarr;
            </button>
          )}
        </div>
      )}

      {/* ---------------------------------------------------------------------- */}
      {/* TAB 1: CHECKLIST INPUT & BLADE ANATOMY MAP */}
      {/* ---------------------------------------------------------------------- */}
      {activeTab === 'checklist-input' && (
        <div className="space-y-6">
          {/* Preset Buttons & Quick Test Bar */}
          <div className="bg-white rounded-xl border border-slate-200 p-4 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-3">
            <div className="flex items-center gap-2">
              <span className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                <RotateCcw size={14} className="text-blue-600" />
                Mẫu thử nghiệm tiêu chuẩn (Presets):
              </span>
            </div>
            <div className="flex items-center gap-2 flex-wrap">
              <button
                type="button"
                onClick={() => handleLoadPreset('cat4_benchmark')}
                className="px-3 py-1.5 bg-orange-50 hover:bg-orange-100 text-orange-800 border border-orange-200 rounded-lg text-xs font-bold transition-colors flex items-center gap-1.5 shadow-2xs"
                title="Mẫu thử chuẩn của User: Cat 4 (Tách mép sau 1.8m, nứt vỏ 85cm, nổ đĩa sét 15cm)"
              >
                <AlertTriangle size={13} className="text-orange-600" />
                Mẫu Cat 4 Chuẩn (BLADE-A-75M)
              </button>
              <button
                type="button"
                onClick={() => handleLoadPreset('cat5_shutdown')}
                className="px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-800 border border-rose-200 rounded-lg text-xs font-bold transition-colors flex items-center gap-1.5 shadow-2xs"
                title="Mẫu khẩn cấp Cat 5: Gãy dầm Spar Cap / Buckling - Dừng máy khẩn cấp"
              >
                <Flame size={13} className="text-rose-600" />
                Mẫu Cat 5 (Dừng máy ngay)
              </button>
              <button
                type="button"
                onClick={() => handleLoadPreset('cat3_erosion')}
                className="px-3 py-1.5 bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 rounded-lg text-xs font-bold transition-colors flex items-center gap-1.5 shadow-2xs"
                title="Mẫu Cat 3: Xói mòn mép trước lộ sợi & tách keo nhỏ"
              >
                <AlertCircle size={13} className="text-amber-600" />
                Mẫu Cat 3 (Xói mòn mép)
              </button>
              <button
                type="button"
                onClick={() => handleLoadPreset('cat1_healthy')}
                className="px-3 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 rounded-lg text-xs font-bold transition-colors flex items-center gap-1.5 shadow-2xs"
                title="Mẫu Cat 1: Cánh vận hành hoàn hảo"
              >
                <CheckCircle2 size={13} className="text-emerald-600" />
                Mẫu Cat 1 (Tốt)
              </button>
            </div>
          </div>

          {/* Asset Metadata & Telemetry Card */}
          <div className="bg-white rounded-2xl border border-slate-200 p-5 sm:p-6 shadow-sm">
            <h3 className="text-base font-bold text-slate-900 mb-4 flex items-center gap-2">
              <MapPin size={18} className="text-blue-600" />
              Thông Tin Thiết Bị & Cảm Biến Giám Sát (Asset & Sensor Telemetry)
            </h3>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Mã Cánh / Blade Tag</label>
                <input
                  type="text"
                  value={formData.bladeTag}
                  onChange={(e) => setFormData({ ...formData, bladeTag: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 rounded-xl text-sm font-bold text-slate-900 focus:ring-2 focus:ring-blue-500 outline-none"
                  placeholder="BLADE-A-75M"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Trụ Turbine / WTG ID</label>
                <input
                  type="text"
                  value={formData.turbineId}
                  onChange={(e) => setFormData({ ...formData, turbineId: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 rounded-xl text-sm text-slate-900 focus:ring-2 focus:ring-blue-500 outline-none"
                  placeholder="WTG-08"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Model Cánh / Blade Model</label>
                <input
                  type="text"
                  value={formData.bladeModel}
                  onChange={(e) => setFormData({ ...formData, bladeModel: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 rounded-xl text-sm text-slate-900 focus:ring-2 focus:ring-blue-500 outline-none"
                  placeholder="LM 75.2 P2 / Vestas V150"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Nhà Máy Điện Gió / Wind Farm</label>
                <input
                  type="text"
                  value={formData.windFarmName}
                  onChange={(e) => setFormData({ ...formData, windFarmName: e.target.value })}
                  className="w-full px-3 py-2 border border-slate-300 rounded-xl text-sm text-slate-900 focus:ring-2 focus:ring-blue-500 outline-none"
                  placeholder="Nhà máy Điện Gió Mũi Né"
                />
              </div>
            </div>

            {/* Operational Telemetry Inputs */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 mt-4 pt-4 border-t border-slate-100">
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Độ rung cánh (mm/s RMS)</label>
                <input
                  type="number"
                  step="0.05"
                  value={formData.vibrationRmsMmS || ''}
                  onChange={(e) => setFormData({ ...formData, vibrationRmsMmS: parseFloat(e.target.value) || 0 })}
                  className={`w-full px-3 py-1.5 border rounded-lg text-sm font-mono font-bold ${
                    (formData.vibrationRmsMmS || 0) > 1.2 ? 'border-amber-400 bg-amber-50/50 text-amber-900' : 'border-slate-300 text-slate-900'
                  }`}
                />
                <span className="text-[10px] text-slate-400">Ngưỡng an toàn &lt; 1.0 mm/s</span>
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Phát xạ âm thanh (dB)</label>
                <input
                  type="number"
                  step="0.5"
                  value={formData.acousticDb || ''}
                  onChange={(e) => setFormData({ ...formData, acousticDb: parseFloat(e.target.value) || 0 })}
                  className={`w-full px-3 py-1.5 border rounded-lg text-sm font-mono font-bold ${
                    (formData.acousticDb || 0) > 45 ? 'border-amber-400 bg-amber-50/50 text-amber-900' : 'border-slate-300 text-slate-900'
                  }`}
                />
                <span className="text-[10px] text-slate-400">Ngưỡng &lt; 40 dB</span>
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Vết nứt mép sau kỳ trước (cm)</label>
                <input
                  type="number"
                  value={formData.previousTeCrackLengthCm || ''}
                  onChange={(e) => setFormData({ ...formData, previousTeCrackLengthCm: parseFloat(e.target.value) || 0 })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg text-sm font-mono text-slate-900"
                  placeholder="20"
                />
                <span className="text-[10px] text-blue-600 font-semibold">Để tính độ tăng nứt mỏi (%)</span>
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">Kỹ sư kiểm tra / Inspector</label>
                <input
                  type="text"
                  value={formData.inspectorName || ''}
                  onChange={(e) => setFormData({ ...formData, inspectorName: e.target.value })}
                  className="w-full px-3 py-1.5 border border-slate-300 rounded-lg text-sm text-slate-900"
                  placeholder="KS. Hoàng Đình Trọng"
                />
              </div>
            </div>
          </div>

          {/* Interactive Visual Blade Anatomy Map */}
          <div className="bg-white rounded-2xl border border-slate-200 p-5 sm:p-6 shadow-sm">
            <div className="flex items-center justify-between mb-3">
              <div>
                <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                  <Eye size={18} className="text-indigo-600" />
                  Sơ Đồ Giải Phẫu Cánh & Định Vị Khuyết Tật (Blade Anatomy & Defect Spatial Map)
                </h3>
                <p className="text-xs text-slate-500">
                  Rê chuột hoặc nhấp vào các điểm đánh dấu trên thân cánh (75m) để xem vị trí khuyết tật
                </p>
              </div>
              <div className="text-xs font-semibold text-indigo-700 bg-indigo-50 px-2.5 py-1 rounded-lg border border-indigo-100">
                Chiều dài: 75.2m
              </div>
            </div>

            {/* Blade SVG Interactive Canvas */}
            <div className="relative bg-slate-900 rounded-xl p-6 overflow-hidden border border-slate-800">
              {/* Length scale markers */}
              <div className="flex justify-between text-[11px] font-mono text-slate-400 mb-2 border-b border-slate-800 pb-1">
                <span>GỐC CÁNH (Root 0m)</span>
                <span>20m</span>
                <span>35m (Vỏ/Dầm)</span>
                <span>50m (Mép trước)</span>
                <span>60m (Mép sau)</span>
                <span>MŨI CÁNH (Tip 75m)</span>
              </div>

              {/* Graphic Blade Contour */}
              <div className="relative h-24 sm:h-28 flex items-center">
                {/* Blade Body Contour */}
                <svg className="w-full h-full drop-shadow-md" viewBox="0 0 1000 120" preserveAspectRatio="none">
                  <defs>
                    <linearGradient id="bladeGrad" x1="0%" y1="0%" x2="100%" y2="0%">
                      <stop offset="0%" stopColor="#334155" />
                      <stop offset="25%" stopColor="#475569" />
                      <stop offset="60%" stopColor="#64748b" />
                      <stop offset="100%" stopColor="#cbd5e1" />
                    </linearGradient>
                    <linearGradient id="sparCapGrad" x1="0%" y1="0%" x2="100%" y2="0%">
                      <stop offset="0%" stopColor="#1e293b" />
                      <stop offset="100%" stopColor="#334155" />
                    </linearGradient>
                  </defs>

                  {/* Main Aerodynamic Shell Profile */}
                  <path
                    d="M 20 60 C 50 15, 120 10, 220 20 C 350 32, 600 48, 970 58 C 980 60, 980 62, 970 64 C 600 74, 350 90, 220 102 C 120 112, 50 105, 20 60 Z"
                    fill="url(#bladeGrad)"
                    stroke="#94a3b8"
                    strokeWidth="2"
                  />

                  {/* Main Spar Cap Internal Spar Area (Chord Center 25m - 65m) */}
                  <path
                    d="M 180 42 L 750 54 L 750 68 L 180 80 Z"
                    fill="url(#sparCapGrad)"
                    opacity="0.75"
                    stroke="#cbd5e1"
                    strokeWidth="1"
                    strokeDasharray="4 2"
                  />
                  <text x="400" y="63" fill="#94a3b8" fontSize="11" fontFamily="sans-serif" textAnchor="middle" fontWeight="bold">
                    DẦM CHỊU LỰC CHÍNH / MAIN SPAR CAP REGION
                  </text>
                </svg>

                {/* Defect Markers Pin Overlay */}
                {/* A5: Vortex Generators at Root (15m) */}
                <div
                  onMouseEnter={() => setHoveredDefectId('A5')}
                  onMouseLeave={() => setHoveredDefectId(null)}
                  className="absolute left-[15%] top-[25%] -translate-x-1/2 -translate-y-1/2 cursor-pointer group"
                >
                  <div className={`w-6 h-6 rounded-full flex items-center justify-center font-bold text-xs shadow-lg transition-transform ${
                    hoveredDefectId === 'A5' ? 'scale-125 ring-4 ring-blue-400' : ''
                  } bg-blue-500 text-white`}>
                    A5
                  </div>
                  <div className="hidden group-hover:block absolute bottom-8 left-1/2 -translate-x-1/2 bg-slate-950 text-white text-[11px] p-2 rounded-lg shadow-xl border border-slate-700 whitespace-nowrap z-30">
                    <strong>A5: Rơi mất 12 khối VG</strong> (Gốc cánh 10-22m)<br />
                    Phân loại: Category 2
                  </div>
                </div>

                {/* A3: Pressure Shell Crack at 35m */}
                <div
                  onMouseEnter={() => setHoveredDefectId('A3')}
                  onMouseLeave={() => setHoveredDefectId(null)}
                  className="absolute left-[38%] top-[70%] -translate-x-1/2 -translate-y-1/2 cursor-pointer group"
                >
                  <div className={`w-7 h-7 rounded-full flex items-center justify-center font-black text-xs shadow-lg transition-transform ${
                    hoveredDefectId === 'A3' ? 'scale-125 ring-4 ring-orange-400' : ''
                  } bg-orange-500 text-white animate-pulse`}>
                    A3
                  </div>
                  <div className="hidden group-hover:block absolute bottom-9 left-1/2 -translate-x-1/2 bg-slate-950 text-white text-[11px] p-2 rounded-lg shadow-xl border border-orange-500/60 whitespace-nowrap z-30">
                    <strong className="text-orange-400">A3: Vết nứt dọc lực 85cm gần Spar Cap (35m)</strong><br />
                    Phân loại: Category 4 | Suy giảm độ bền uốn
                  </div>
                </div>

                {/* A1: Leading Edge Erosion at 45m-55m */}
                <div
                  onMouseEnter={() => setHoveredDefectId('A1')}
                  onMouseLeave={() => setHoveredDefectId(null)}
                  className="absolute left-[54%] top-[18%] -translate-x-1/2 -translate-y-1/2 cursor-pointer group"
                >
                  <div className={`w-6 h-6 rounded-full flex items-center justify-center font-bold text-xs shadow-lg transition-transform ${
                    hoveredDefectId === 'A1' ? 'scale-125 ring-4 ring-amber-400' : ''
                  } bg-amber-500 text-white`}>
                    A1
                  </div>
                  <div className="hidden group-hover:block absolute bottom-8 left-1/2 -translate-x-1/2 bg-slate-950 text-white text-[11px] p-2 rounded-lg shadow-xl border border-amber-500/60 whitespace-nowrap z-30">
                    <strong className="text-amber-400">A1: Xói mòn mép trước lộ sợi (45m-55m)</strong><br />
                    Phân loại: Category 3 | Suy giảm hiệu suất khí động
                  </div>
                </div>

                {/* A2: Trailing Edge Split Seam at 60m */}
                <div
                  onMouseEnter={() => setHoveredDefectId('A2')}
                  onMouseLeave={() => setHoveredDefectId(null)}
                  className="absolute left-[64%] top-[80%] -translate-x-1/2 -translate-y-1/2 cursor-pointer group"
                >
                  <div className={`w-7 h-7 rounded-full flex items-center justify-center font-black text-xs shadow-lg transition-transform ${
                    hoveredDefectId === 'A2' ? 'scale-125 ring-4 ring-red-400' : ''
                  } bg-red-600 text-white animate-bounce`}>
                    A2
                  </div>
                  <div className="hidden group-hover:block absolute bottom-9 left-1/2 -translate-x-1/2 bg-slate-950 text-white text-[11px] p-2 rounded-lg shadow-xl border border-red-500/60 whitespace-nowrap z-30">
                    <strong className="text-red-400">A2: Tách keo mép sau dài 1.8m lộ lõi xốp (60m)</strong><br />
                    Phân loại: Category 4 | Tăng trưởng &gt;800% so với năm trước
                  </div>
                </div>

                {/* A4: Lightning Strike Tip Receptor at Tip (74m) */}
                <div
                  onMouseEnter={() => setHoveredDefectId('A4')}
                  onMouseLeave={() => setHoveredDefectId(null)}
                  className="absolute left-[95%] top-[50%] -translate-x-1/2 -translate-y-1/2 cursor-pointer group"
                >
                  <div className={`w-7 h-7 rounded-full flex items-center justify-center font-black text-xs shadow-lg transition-transform ${
                    hoveredDefectId === 'A4' ? 'scale-125 ring-4 ring-purple-400' : ''
                  } bg-purple-600 text-white`}>
                    A4
                  </div>
                  <div className="hidden group-hover:block absolute bottom-9 right-0 bg-slate-950 text-white text-[11px] p-2 rounded-lg shadow-xl border border-purple-500/60 whitespace-nowrap z-30">
                    <strong className="text-purple-400">A4: Đĩa thu sét #3 nổ chảy, thủng vỏ 15cm (Tip)</strong><br />
                    Phân loại: Category 4 | Mất khả năng chống sét
                  </div>
                </div>
              </div>

              {/* Diagram Legend */}
              <div className="flex items-center justify-between text-[11px] text-slate-400 mt-3 pt-2 border-t border-slate-800 flex-wrap gap-2">
                <span className="flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-blue-500"></span> A5: Khí động (VGs)
                </span>
                <span className="flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-amber-500"></span> A1: Mép trước (LE)
                </span>
                <span className="flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-orange-500"></span> A3: Vỏ cánh/Dầm (Shell)
                </span>
                <span className="flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-red-600"></span> A2: Mép sau (TE)
                </span>
                <span className="flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-purple-600"></span> A4: Chống sét (LPS)
                </span>
              </div>
            </div>
          </div>

          {/* SECTION A: DETAILED 5-ITEM CHECKLIST TABLE */}
          <div className="bg-white rounded-2xl border border-slate-200 p-5 sm:p-6 shadow-sm">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4">
              <div>
                <h3 className="text-base font-bold text-slate-900 flex items-center gap-2">
                  <Sliders size={18} className="text-blue-600" />
                  PHẦN A: BẢNG CHECKLIST KIỂM TRA CÁNH QUẠT GIÓ (5 HẠNG MỤC CHUẨN EPRI / DNV GL)
                </h3>
                <p className="text-xs text-slate-500">
                  Dữ liệu ghi nhận từ hình ảnh Drone độ phân giải cao và kỹ sư hiện trường
                </p>
              </div>

              <button
                type="button"
                onClick={handleRunAiAnalysis}
                disabled={isLoadingAi}
                className="px-5 py-2.5 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white rounded-xl text-sm font-bold shadow-md hover:shadow-lg transition-all flex items-center gap-2 justify-center"
              >
                {isLoadingAi ? (
                  <>
                    <Sparkles size={16} className="animate-spin" />
                    <span>Đang phân tích AI...</span>
                  </>
                ) : (
                  <>
                    <Sparkles size={16} />
                    <span>Phân Tích AI Chuyên Sâu (EPRI / DNV GL)</span>
                  </>
                )}
              </button>
            </div>

            <div className="space-y-4">
              {formData.checklistItems.map((item, idx) => {
                const catInfo = EPRI_CATEGORIES_INFO[item.severityCategory];
                return (
                  <div
                    key={item.id}
                    onMouseEnter={() => setHoveredDefectId(item.id)}
                    onMouseLeave={() => setHoveredDefectId(null)}
                    className={`rounded-xl border p-4 transition-all ${
                      hoveredDefectId === item.id ? 'ring-2 ring-blue-500 shadow-md' : 'shadow-2xs'
                    } ${
                      item.severityCategory === 'Category 5' ? 'border-rose-300 bg-rose-50/20' :
                      item.severityCategory === 'Category 4' ? 'border-orange-300 bg-orange-50/20' :
                      item.severityCategory === 'Category 3' ? 'border-amber-300 bg-amber-50/20' :
                      'border-slate-200 bg-white'
                    }`}
                  >
                    <div className="flex flex-col md:flex-row md:items-start justify-between gap-4">
                      {/* Left: Item Info */}
                      <div className="flex-1">
                        <div className="flex items-center gap-2 flex-wrap mb-1">
                          <span className="px-2 py-0.5 rounded bg-slate-900 text-white text-xs font-mono font-bold">
                            {item.id}
                          </span>
                          <h4 className="text-sm font-bold text-slate-900">
                            {langMode === 'en' ? item.nameEn : langMode === 'vn' ? item.nameVn : `${item.nameVn} / ${item.nameEn}`}
                          </h4>
                          {item.locationSpanM && (
                            <span className="px-2 py-0.5 rounded-full text-[11px] font-semibold bg-indigo-50 text-indigo-700 border border-indigo-200">
                              Vị trí: {item.locationSpanM}
                            </span>
                          )}
                        </div>

                        {/* Standard Procedure & Sign */}
                        <div className="text-xs text-slate-600 mt-2 space-y-1">
                          <p>
                            <strong className="text-slate-800">Dấu hiệu phát hiện:</strong>{' '}
                            {langMode === 'en' ? item.defectSignEn : langMode === 'vn' ? item.defectSignVn : `${item.defectSignVn} (${item.defectSignEn})`}
                          </p>
                        </div>

                        {/* Finding Details Editable Input */}
                        <div className="mt-3">
                          <label className="block text-[11px] font-semibold text-slate-500 mb-1">
                            Mô tả chi tiết vết lỗi tại hiện trường (Field Finding Details):
                          </label>
                          <textarea
                            rows={2}
                            value={item.detailsVn || ''}
                            onChange={(e) => {
                              const newDetails = e.target.value;
                              setFormData({
                                ...formData,
                                checklistItems: formData.checklistItems.map(it => it.id === item.id ? { ...it, detailsVn: newDetails } : it)
                              });
                            }}
                            className="w-full px-3 py-1.5 text-xs border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
                            placeholder="Ghi chú chi tiết kích thước, dạng nứt, tình trạng lớp keo..."
                          />
                        </div>
                      </div>

                      {/* Right: Controls (Category, Status, Length) */}
                      <div className="w-full md:w-64 shrink-0 bg-slate-50 rounded-xl p-3 border border-slate-200 space-y-2.5">
                        {/* Status Toggle (Pass / Investigate / Fail) */}
                        <div>
                          <label className="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1">
                            Quy Tắc Đánh Giá (App Logic)
                          </label>
                          <div className="grid grid-cols-3 gap-1">
                            <button
                              type="button"
                              onClick={() => {
                                setFormData({
                                  ...formData,
                                  checklistItems: formData.checklistItems.map(it => it.id === item.id ? { ...it, status: 'Pass', severityCategory: 'Category 1' } : it)
                                });
                              }}
                              className={`py-1 rounded text-xs font-bold transition-all ${
                                item.status === 'Pass' ? 'bg-emerald-600 text-white shadow-2xs' : 'bg-white text-slate-700 hover:bg-slate-100'
                              }`}
                            >
                              Pass
                            </button>
                            <button
                              type="button"
                              onClick={() => {
                                setFormData({
                                  ...formData,
                                  checklistItems: formData.checklistItems.map(it => it.id === item.id ? { ...it, status: 'Investigate', severityCategory: 'Category 3' } : it)
                                });
                              }}
                              className={`py-1 rounded text-xs font-bold transition-all ${
                                item.status === 'Investigate' ? 'bg-amber-600 text-white shadow-2xs' : 'bg-white text-slate-700 hover:bg-slate-100'
                              }`}
                            >
                              Investigate
                            </button>
                            <button
                              type="button"
                              onClick={() => {
                                setFormData({
                                  ...formData,
                                  checklistItems: formData.checklistItems.map(it => it.id === item.id ? { ...it, status: 'Fail', severityCategory: 'Category 4' } : it)
                                });
                              }}
                              className={`py-1 rounded text-xs font-bold transition-all ${
                                item.status === 'Fail' ? 'bg-rose-600 text-white shadow-2xs' : 'bg-white text-slate-700 hover:bg-slate-100'
                              }`}
                            >
                              Fail
                            </button>
                          </div>
                        </div>

                        {/* Severity Category Selector */}
                        <div>
                          <label className="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1">
                            Phân Loại Mức Độ (EPRI Category)
                          </label>
                          <select
                            value={item.severityCategory}
                            onChange={(e) => {
                              const newCat = e.target.value as EpriBladeCategory;
                              setFormData({
                                ...formData,
                                checklistItems: formData.checklistItems.map(it => it.id === item.id ? { ...it, severityCategory: newCat } : it)
                              });
                            }}
                            className={`w-full px-2 py-1.5 rounded-lg text-xs font-bold border ${catInfo.badgeClass} outline-none cursor-pointer`}
                          >
                            <option value="Category 1">Category 1 (Thẩm mỹ nhẹ / Intact)</option>
                            <option value="Category 2">Category 2 (Bề mặt nhẹ / Minor)</option>
                            <option value="Category 3">Category 3 (Hư hỏng vừa / Moderate)</option>
                            <option value="Category 4">Category 4 (Nghiêm trọng / Significant)</option>
                            <option value="Category 5">Category 5 (Sụp đổ cấu trúc / Severe)</option>
                          </select>
                        </div>

                        {/* Defect Length in cm */}
                        <div>
                          <label className="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1">
                            Kích thước vết lỗi (cm / mm)
                          </label>
                          <input
                            type="number"
                            value={item.defectLengthCm || 0}
                            onChange={(e) => {
                              const len = parseFloat(e.target.value) || 0;
                              setFormData({
                                ...formData,
                                checklistItems: formData.checklistItems.map(it => it.id === item.id ? { ...it, defectLengthCm: len } : it)
                              });
                            }}
                            className="w-full px-2 py-1 border border-slate-300 rounded text-xs font-mono font-bold"
                            placeholder="cm"
                          />
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Bottom Action Bar */}
            <div className="mt-6 pt-4 border-t border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div className="text-xs text-slate-500 flex items-center gap-1.5">
                <Info size={14} className="text-blue-500" />
                <span>Tiêu chuẩn đánh giá tự động dựa theo thuật toán DNVGL-ST-0376 và tài liệu kỹ thuật EPRI 3002019669.</span>
              </div>
              <button
                type="button"
                onClick={handleRunAiAnalysis}
                disabled={isLoadingAi}
                className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-sm font-bold shadow-md transition-all flex items-center gap-2 justify-center"
              >
                {isLoadingAi ? <Sparkles size={16} className="animate-spin" /> : <Sparkles size={16} />}
                <span>Khởi Chạy Phân Tích &amp; Xuất Khuyến Nghị</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ---------------------------------------------------------------------- */}
      {/* TAB 2: DETAILED ANALYSIS RESULT & REPAIR ACTION PLAN */}
      {/* ---------------------------------------------------------------------- */}
      {activeTab === 'analysis-result' && analysisResult && (
        <div className="space-y-6">
          {/* Top Status Banner */}
          <div className={`p-6 rounded-2xl border shadow-sm ${
            analysisResult.overall_status === 'EMERGENCY_SHUTDOWN'
              ? 'bg-rose-50 border-rose-300 text-rose-950'
              : analysisResult.overall_status === 'FAIL'
              ? 'bg-orange-50 border-orange-300 text-orange-950'
              : analysisResult.overall_status === 'INVESTIGATE'
              ? 'bg-amber-50 border-amber-300 text-amber-950'
              : 'bg-emerald-50 border-emerald-300 text-emerald-950'
          }`}>
            <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
              <div className="flex items-start gap-4">
                <div className={`w-14 h-14 rounded-2xl flex items-center justify-center shrink-0 shadow-md ${
                  analysisResult.overall_status === 'EMERGENCY_SHUTDOWN'
                    ? 'bg-rose-600 text-white animate-pulse'
                    : analysisResult.overall_status === 'FAIL'
                    ? 'bg-orange-600 text-white'
                    : analysisResult.overall_status === 'INVESTIGATE'
                    ? 'bg-amber-500 text-white'
                    : 'bg-emerald-600 text-white'
                }`}>
                  {analysisResult.overall_status === 'EMERGENCY_SHUTDOWN' ? <Flame size={32} /> :
                   analysisResult.overall_status === 'FAIL' ? <AlertTriangle size={32} /> :
                   analysisResult.overall_status === 'INVESTIGATE' ? <AlertCircle size={32} /> :
                   <CheckCircle2 size={32} />}
                </div>
                <div>
                  <div className="flex items-center gap-2 flex-wrap mb-1">
                    <span className="text-xs font-bold uppercase tracking-wider px-2 py-0.5 rounded bg-white/80 border">
                      TRẠNG THÁI TỔNG QUAN / OVERALL STATUS
                    </span>
                    <span className="px-2.5 py-0.5 rounded text-xs font-black bg-slate-900 text-white">
                      {analysisResult.max_category} MAX
                    </span>
                  </div>
                  <h3 className="text-xl sm:text-2xl font-black tracking-tight">
                    {analysisResult.overall_status === 'EMERGENCY_SHUTDOWN'
                      ? 'DỪNG MÁY KHẨN CẤP / EMERGENCY SHUTDOWN (CATEGORY 5)'
                      : analysisResult.overall_status === 'FAIL'
                      ? 'FAIL / URGENT REPAIR REQUIRED (CATEGORY 4 MAX) (LỖI KẾT CẤU BẮT BUỘC SỬA CHỮA KHẨN CẤP)'
                      : analysisResult.overall_status === 'INVESTIGATE'
                      ? 'CẦN THEO DÕI & LẬP KẾ HOẠCH BẢO DƯỠNG (CATEGORY 2/3)'
                      : 'PASS / CÁNH VẬN HÀNH BÌNH THƯỜNG (CATEGORY 1/2)'}
                  </h3>
                  <p className="text-xs sm:text-sm mt-1 opacity-90">
                    Thiết bị: <strong>{formData.bladeTag}</strong> • Trụ <strong>{formData.turbineId}</strong> • {formData.windFarmName}
                  </p>
                </div>
              </div>

              {/* Action Buttons */}
              <div className="flex items-center gap-2 flex-wrap">
                <button
                  type="button"
                  onClick={handleCreateWorkOrder}
                  className="px-4 py-2 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold shadow transition-all flex items-center gap-1.5"
                >
                  <Wrench size={14} className="text-amber-400" />
                  <span>Tạo Phiếu WO / WR ({analysisResult.json_data.recommendations.recommended_work_order_priority})</span>
                </button>
                <button
                  type="button"
                  onClick={handleSaveToReports}
                  className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold shadow transition-all flex items-center gap-1.5"
                >
                  <Download size={14} />
                  <span>Lưu Kho Báo Cáo</span>
                </button>
              </div>
            </div>
          </div>

          {/* 1. BẢNG PHÂN TÍCH CHI TIẾT CÁNH QUẠT GIÓ (TABLE 1) */}
          <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-sm">
            <div className="p-5 border-b border-slate-100 flex items-center justify-between">
              <div>
                <h4 className="text-base font-bold text-slate-900 flex items-center gap-2">
                  <FileSpreadsheet size={18} className="text-blue-600" />
                  1. BẢNG PHÂN TÍCH CHI TIẾT CÁNH QUẠT GIÓ / DETAILED BLADE EVALUATION TABLE
                </h4>
                <p className="text-xs text-slate-500 mt-0.5">
                  Phân loại theo tiêu chuẩn EPRI 3002019669 & DNVGL-ST-0376
                </p>
              </div>
              <span className="text-xs font-mono font-bold text-slate-500">
                {analysisResult.json_data.evaluation_table.length} Hạng Mục
              </span>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="bg-slate-50 text-slate-700 border-b border-slate-200">
                    <th className="p-3.5 font-bold uppercase tracking-wider w-1/4">Vị trí Cánh / Blade Location</th>
                    <th className="p-3.5 font-bold uppercase tracking-wider w-1/3">Dạng Lỗi Nhận diện / Detected Finding</th>
                    <th className="p-3.5 font-bold uppercase tracking-wider text-center w-36">Phân loại EPRI</th>
                    <th className="p-3.5 font-bold uppercase tracking-wider">Tác động &amp; Rủi ro / Risk Impact</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {analysisResult.json_data.evaluation_table.map((row, idx) => {
                    const catInfo = EPRI_CATEGORIES_INFO[row.category];
                    return (
                      <tr key={idx} className="hover:bg-slate-50/70 transition-colors">
                        <td className="p-3.5 align-top">
                          <div className="font-bold text-slate-900">{row.locationVn}</div>
                          {langMode !== 'vn' && (
                            <div className="text-[11px] text-slate-500 italic">{row.locationEn}</div>
                          )}
                        </td>
                        <td className="p-3.5 align-top">
                          <div className="font-semibold text-slate-800">{row.findingVn}</div>
                          {langMode !== 'vn' && (
                            <div className="text-[11px] text-slate-500 italic mt-0.5">{row.findingEn}</div>
                          )}
                        </td>
                        <td className="p-3.5 align-top text-center">
                          <span className={`inline-block px-2.5 py-1 rounded-full text-xs font-bold border ${catInfo.badgeClass}`}>
                            {row.category}
                          </span>
                        </td>
                        <td className="p-3.5 align-top">
                          <div className="text-slate-800">{row.riskImpactVn}</div>
                          {langMode !== 'vn' && (
                            <div className="text-[11px] text-slate-500 italic mt-0.5">{row.riskImpactEn}</div>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          {/* 2. PHÂN TÍCH BỀN KẾT CẤU & NGUY CƠ KHÍ ĐỘNG HỌC (SECTION 2) */}
          <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-5">
            <div className="border-b border-slate-100 pb-3">
              <h4 className="text-base font-bold text-slate-900 flex items-center gap-2">
                <ShieldAlert size={18} className="text-orange-600" />
                2. PHÂN TÍCH BỀN KẾT CẤU &amp; NGUY CƠ KHÍ ĐỘNG HỌC / STRUCTURAL INTEGRITY &amp; AERODYNAMIC RISK ANALYSIS
              </h4>
              <p className="text-xs text-slate-500 mt-0.5">
                Đánh giá suy giảm tải trọng mỏi, ngấm ẩm lõi xốp composite và nguy cơ sét đánh
              </p>
            </div>

            {/* Severity Assessment */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200/80 space-y-2">
              <h5 className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-blue-600"></span>
                Mức độ Hư hỏng Tổng thể Cánh (Overall Blade Damage Severity):
              </h5>
              {langMode !== 'vn' && (
                <p className="text-xs text-slate-800 font-medium">
                  <strong className="text-slate-900">EN:</strong> {analysisResult.json_data.structural_integrity.severity_assessment_en}
                </p>
              )}
              {langMode !== 'en' && (
                <p className="text-xs text-slate-800 font-medium">
                  <strong className="text-slate-900">VN:</strong> {analysisResult.json_data.structural_integrity.severity_assessment_vn}
                </p>
              )}
            </div>

            {/* Defect Growth Rate & Fatigue Risk */}
            <div className="p-4 bg-orange-50/40 rounded-xl border border-orange-200/80 space-y-2">
              <div className="flex items-center justify-between flex-wrap gap-2">
                <h5 className="text-xs font-bold text-orange-900 uppercase tracking-wider flex items-center gap-1.5">
                  <TrendingUp size={14} className="text-orange-600" />
                  Nguy cơ Tác động &amp; Tốc độ Phát triển Lỗi (Impact &amp; Defect Growth Risk):
                </h5>
                {growthRate > 0 && (
                  <span className="px-2 py-0.5 rounded text-[11px] font-bold bg-orange-100 text-orange-800 border border-orange-300">
                    Tách mép sau tăng: +{growthRate}% so với năm trước ({prevTe}cm ➔ {currentTe}cm)
                  </span>
                )}
              </div>
              {langMode !== 'vn' && (
                <p className="text-xs text-slate-800 font-medium">
                  <strong className="text-slate-900">EN:</strong> {analysisResult.json_data.structural_integrity.growth_and_fatigue_risk_en}
                </p>
              )}
              {langMode !== 'en' && (
                <p className="text-xs text-slate-800 font-medium">
                  <strong className="text-slate-900">VN:</strong> {analysisResult.json_data.structural_integrity.growth_and_fatigue_risk_vn}
                </p>
              )}
            </div>
          </div>

          {/* 3. KHUYẾN NGHỊ HÀNH ĐỘNG & TIẾN ĐỘ SỬA CHỮA (SECTION 3) */}
          <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-5">
            <div className="border-b border-slate-100 pb-3 flex items-center justify-between flex-wrap gap-2">
              <div>
                <h4 className="text-base font-bold text-slate-900 flex items-center gap-2">
                  <Clock size={18} className="text-indigo-600" />
                  3. KHUYẾN NGHỊ HÀNH ĐỘNG &amp; TIẾN ĐỘ SỬA CHỮA / ACTIONABLE RECOMMENDATIONS &amp; REPAIR TIMELINE
                </h4>
                <p className="text-xs text-slate-500 mt-0.5">
                  Lộ trình xử lý kỹ thuật và phương án sửa chữa trên cao (Uptower Protocol)
                </p>
              </div>
              <div className="px-3 py-1 bg-indigo-50 border border-indigo-200 text-indigo-900 rounded-lg text-xs font-bold">
                Ưu tiên CMMS: {analysisResult.json_data.recommendations.recommended_work_order_priority}
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              {/* Item 1: Immediate Evaluation */}
              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
                <div className="flex items-center gap-2 text-slate-900 font-bold text-xs uppercase tracking-wider">
                  <AlertCircle size={14} className="text-blue-600" />
                  Đánh giá Kỹ thuật Khẩn cấp
                </div>
                {langMode !== 'vn' && (
                  <p className="text-xs text-slate-700">
                    <strong>EN:</strong> {analysisResult.json_data.recommendations.engineering_evaluation_en}
                  </p>
                )}
                {langMode !== 'en' && (
                  <p className="text-xs text-slate-800">
                    <strong>VN:</strong> {analysisResult.json_data.recommendations.engineering_evaluation_vn}
                  </p>
                )}
              </div>

              {/* Item 2: Uptower Repair Protocol */}
              <div className="p-4 bg-blue-50/50 rounded-xl border border-blue-200 space-y-2">
                <div className="flex items-center gap-2 text-blue-900 font-bold text-xs uppercase tracking-wider">
                  <Wrench size={14} className="text-blue-700" />
                  Sửa chữa Trên cao (30-60 Ngày)
                </div>
                {langMode !== 'vn' && (
                  <div className="text-xs text-slate-700 whitespace-pre-line">
                    <strong>EN:</strong> {analysisResult.json_data.recommendations.uptower_repair_protocol_en}
                  </div>
                )}
                {langMode !== 'en' && (
                  <div className="text-xs text-slate-800 whitespace-pre-line">
                    <strong>VN:</strong> {analysisResult.json_data.recommendations.uptower_repair_protocol_vn}
                  </div>
                )}
              </div>

              {/* Item 3: Monthly Monitoring */}
              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
                <div className="flex items-center gap-2 text-slate-900 font-bold text-xs uppercase tracking-wider">
                  <Eye size={14} className="text-indigo-600" />
                  Giám sát Định kỳ (Drone)
                </div>
                {langMode !== 'vn' && (
                  <p className="text-xs text-slate-700">
                    <strong>EN:</strong> {analysisResult.json_data.recommendations.monitoring_plan_en}
                  </p>
                )}
                {langMode !== 'en' && (
                  <p className="text-xs text-slate-800">
                    <strong>VN:</strong> {analysisResult.json_data.recommendations.monitoring_plan_vn}
                  </p>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ---------------------------------------------------------------------- */}
      {/* TAB 3: EPRI / DNV GL DAMAGE CATEGORIZATION MATRIX REFERENCE */}
      {/* ---------------------------------------------------------------------- */}
      {activeTab === 'matrix-reference' && (
        <div className="space-y-6">
          <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
            <div className="mb-6">
              <h3 className="text-lg font-bold text-slate-900 flex items-center gap-2">
                <Layers size={20} className="text-blue-600" />
                PHẦN B: PHÂN LOẠI MỨC ĐỘ HƯ HỎNG THEO CHUẨN EPRI / DNV GL (DAMAGE CATEGORIZATION MATRIX)
              </h3>
              <p className="text-xs sm:text-sm text-slate-500 mt-1">
                Tài liệu tham chiếu chuẩn hóa: EPRI 3002019669 | DNVGL-ST-0376 | IEC 61400-5
              </p>
            </div>

            <div className="space-y-4">
              {(Object.keys(EPRI_CATEGORIES_INFO) as EpriBladeCategory[]).map((catKey) => {
                const info = EPRI_CATEGORIES_INFO[catKey];
                return (
                  <div
                    key={catKey}
                    className={`rounded-xl border p-5 transition-all ${info.borderClass} ${info.bgClass}`}
                  >
                    <div className="flex flex-col lg:flex-row lg:items-start justify-between gap-4">
                      {/* Left Header */}
                      <div className="lg:w-1/4">
                        <span className={`inline-block px-3 py-1 rounded-full text-xs font-black border ${info.badgeClass} mb-1.5`}>
                          {info.category}
                        </span>
                        <h4 className="text-base font-bold text-slate-900">{info.titleVn}</h4>
                        <div className="text-xs text-slate-500 italic">{info.titleEn}</div>
                      </div>

                      {/* Middle: Description & Impact */}
                      <div className="lg:w-1/2 space-y-2 text-xs">
                        <div>
                          <strong className="text-slate-800">Mô tả đặc tính hư hỏng / Description:</strong>
                          <p className="text-slate-700 mt-0.5">{info.descVn}</p>
                          <p className="text-slate-500 italic">{info.descEn}</p>
                        </div>
                        <div className="pt-1.5 border-t border-slate-200/60">
                          <strong className="text-slate-800">Tác động kết cấu &amp; khí động / Impact:</strong>
                          <p className="text-slate-700 mt-0.5">{info.impactVn}</p>
                          <p className="text-slate-500 italic">{info.impactEn}</p>
                        </div>
                      </div>

                      {/* Right: Action & Timeline */}
                      <div className="lg:w-1/4 bg-white/80 rounded-xl p-3 border border-slate-200/80 text-xs space-y-1">
                        <strong className="text-slate-900 font-bold block text-[11px] uppercase tracking-wider">
                          Kế Hoạch &amp; Thời Gian Khắc Phục:
                        </strong>
                        <p className="text-slate-800 font-semibold">{info.actionVn}</p>
                        <p className="text-slate-500 italic text-[11px]">{info.actionEn}</p>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
