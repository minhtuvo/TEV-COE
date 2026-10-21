import React, { useState, useMemo } from 'react';
import {
  X,
  Activity,
  Clock,
  AlertTriangle,
  Zap,
  TrendingDown,
  TrendingUp,
  DollarSign,
  Copy,
  Check,
  Download,
  ShieldCheck,
  CheckCircle2,
  HelpCircle,
  FileCode,
  Layers,
  ChevronRight,
  Flame,
  Award,
  BarChart3
} from 'lucide-react';
import {
  ArpAnalysisResult,
  FleetArpSummary,
  generateMobiusArpJson
} from '../utils/arpReliabilityService';

interface Props {
  isOpen: boolean;
  onClose: () => void;
  fleetSummary: FleetArpSummary;
  selectedAssetId?: string;
  allEquipment: any[];
  workOrders: any[];
}

export const ArpReliabilityAnalysisModal: React.FC<Props> = ({
  isOpen,
  onClose,
  fleetSummary,
  selectedAssetId,
  allEquipment,
  workOrders
}) => {
  const [currentAssetId, setCurrentAssetId] = useState<string>(selectedAssetId || 'fleet');
  const [copied, setCopied] = useState<boolean>(false);
  const [viewMode, setViewMode] = useState<'dashboard' | 'json' | 'pareto'>('dashboard');

  // Select target asset or synthesized fleet asset
  const activeResult: ArpAnalysisResult = useMemo(() => {
    if (currentAssetId === 'fleet' || !fleetSummary.asset_results[currentAssetId]) {
      // Synthesize fleet-level ArpAnalysisResult
      return {
        asset_id: 'FLEET-TOTAL',
        asset_name: `Toàn bộ Fleet (${fleetSummary.total_assets} thiết bị)`,
        customer: 'Tất cả khách hàng',
        factory: 'Toàn bộ nhà máy',
        asset_type: 'Repairable',
        analysis_period_hours: fleetSummary.period_hours,
        total_failures: fleetSummary.total_failures,
        operating_hours: fleetSummary.total_operating_hours,
        reliability_metrics: {
          MTBF_hours: fleetSummary.fleet_mtbf,
          MTTF_hours: fleetSummary.fleet_mttf,
          MTTR_hours: fleetSummary.fleet_mttr,
          MWT_hours: fleetSummary.fleet_mwt,
          total_unscheduled_downtime_hours: fleetSummary.fleet_unscheduled_downtime,
          total_planned_downtime_hours: fleetSummary.fleet_planned_downtime,
          total_downtime_hours: fleetSummary.fleet_total_downtime,
          operational_availability_percent: fleetSummary.fleet_availability_percent,
          failure_rate_lambda: fleetSummary.total_operating_hours > 0 
            ? Number(((fleetSummary.total_failures / fleetSummary.total_operating_hours) * 1000).toFixed(3)) 
            : 0
        },
        cost_metrics: {
          total_direct_repair_cost: fleetSummary.fleet_total_direct_cost,
          MCTR_per_failure: fleetSummary.fleet_mctr,
          total_downtime_cost: fleetSummary.fleet_downtime_cost,
          total_quality_loss_cost: fleetSummary.fleet_quality_loss,
          COUR_total: fleetSummary.fleet_cour_total
        },
        reliability_warnings: fleetSummary.warnings,
        recommendations: [
          fleetSummary.mwt_exceeds_mttr 
            ? `Áp dụng Job Kitting (soạn trước vật tư theo danh mục bảo trì) và chuẩn hóa phê duyệt Work Permit để kéo MWT (${fleetSummary.fleet_mwt}h) xuống thấp hơn MTTR (${fleetSummary.fleet_mttr}h).`
            : 'Duy trì tỷ lệ MWT/MTTR trong mức kiểm soát an toàn.',
          'Thực hiện RCM (Reliability-Centered Maintenance) cho nhóm Bad Actors (20% thiết bị chiếm 80% COUR).'
        ],
        is_bad_actor: fleetSummary.bad_actors.length > 0
      };
    }

    return fleetSummary.asset_results[currentAssetId];
  }, [currentAssetId, fleetSummary]);

  // Formatted JSON String
  const jsonContent = useMemo(() => {
    return generateMobiusArpJson(activeResult);
  }, [activeResult]);

  if (!isOpen) return null;

  const handleCopyJson = () => {
    navigator.clipboard.writeText(jsonContent);
    setCopied(true);
    setTimeout(() => setCopied(false), 2500);
  };

  const handleDownloadJson = () => {
    const blob = new Blob([jsonContent], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `ARP_E_${activeResult.asset_id}_${new Date().toISOString().slice(0, 10)}.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const mwtMttrRatio = activeResult.reliability_metrics.MTTR_hours > 0 
    ? (activeResult.reliability_metrics.MWT_hours / activeResult.reliability_metrics.MTTR_hours).toFixed(1)
    : '1.0';
  const mwtExceeds = activeResult.reliability_metrics.MWT_hours > activeResult.reliability_metrics.MTTR_hours;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/70 backdrop-blur-sm animate-in fade-in duration-200">
      <div className="bg-white w-full max-w-5xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[92vh]">
        
        {/* Header */}
        <div className="bg-slate-900 text-white px-6 py-4 flex items-center justify-between border-b border-slate-800">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-blue-600/30 border border-blue-500/50 flex items-center justify-center text-blue-400">
              <Award size={22} />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="text-lg font-bold">Mô Hình Độ Tin Cậy & Chi Phí Mobius ARP-E</h3>
                <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-blue-500/20 text-blue-300 border border-blue-400/30">
                  ARP-E Certified Engine
                </span>
              </div>
              <p className="text-xs text-slate-400">
                Chuẩn Mobius Institute (Asset Reliability Practitioner – Engineer) · Phân tích MTBF, MTTF, MTTR, MWT, Ao & COUR
              </p>
            </div>
          </div>
          
          <button 
            onClick={onClose}
            className="w-8 h-8 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white flex items-center justify-center transition-colors"
          >
            <X size={18} />
          </button>
        </div>

        {/* Toolbar: Asset Selector & View Switcher */}
        <div className="bg-slate-50 px-6 py-3 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-2">
            <label className="text-xs font-bold text-slate-600 uppercase tracking-wider">Chọn Đối Tượng:</label>
            <select
              value={currentAssetId}
              onChange={(e) => setCurrentAssetId(e.target.value)}
              className="bg-white border border-slate-300 text-slate-800 text-xs font-semibold rounded-lg px-3 py-1.5 outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 shadow-2xs min-w-[280px]"
            >
              <option value="fleet">Toàn bộ Fleet ({fleetSummary.total_assets} thiết bị)</option>
              <optgroup label="Thiết bị cụ thể (104 Thiết bị)">
                {allEquipment.map(eq => (
                  <option key={eq.id} value={eq.id}>
                    [{eq.id}] {eq.name} ({eq.factory})
                  </option>
                ))}
              </optgroup>
            </select>
          </div>

          <div className="flex items-center gap-2">
            <div className="flex items-center bg-slate-200/80 p-0.5 rounded-lg text-xs font-medium text-slate-600">
              <button
                onClick={() => setViewMode('dashboard')}
                className={`px-3 py-1 rounded-md transition-all ${
                  viewMode === 'dashboard' ? 'bg-white text-blue-700 font-bold shadow-2xs' : 'hover:text-slate-900'
                }`}
              >
                Chỉ số ARP-E
              </button>
              <button
                onClick={() => setViewMode('pareto')}
                className={`px-3 py-1 rounded-md transition-all ${
                  viewMode === 'pareto' ? 'bg-white text-blue-700 font-bold shadow-2xs' : 'hover:text-slate-900'
                }`}
              >
                Pareto Bad Actors ({fleetSummary.bad_actors.length})
              </button>
              <button
                onClick={() => setViewMode('json')}
                className={`px-3 py-1 rounded-md transition-all flex items-center gap-1.5 ${
                  viewMode === 'json' ? 'bg-white text-blue-700 font-bold shadow-2xs' : 'hover:text-slate-900'
                }`}
              >
                <FileCode size={13} />
                JSON Output
              </button>
            </div>

            <button
              onClick={handleCopyJson}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 text-xs font-semibold rounded-lg transition-colors shadow-2xs"
              title="Sao chép JSON chuẩn hóa sang Clipboard"
            >
              {copied ? <Check size={14} className="text-emerald-600" /> : <Copy size={14} />}
              <span>{copied ? 'Đã sao chép' : 'Copy JSON'}</span>
            </button>

            <button
              onClick={handleDownloadJson}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold rounded-lg transition-colors shadow-2xs"
            >
              <Download size={14} />
              <span>Tải JSON</span>
            </button>
          </div>
        </div>

        {/* Content Body */}
        <div className="flex-1 overflow-y-auto p-6 space-y-6">
          
          {/* VIEW MODE 1: DASHBOARD OVERVIEW */}
          {viewMode === 'dashboard' && (
            <>
              {/* Asset Header Info */}
              <div className="bg-gradient-to-r from-slate-900 to-slate-800 text-white rounded-xl p-5 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div className="space-y-1">
                  <div className="flex items-center gap-2">
                    <span className="font-mono text-sm px-2 py-0.5 rounded bg-blue-500/20 text-blue-300 border border-blue-400/30 font-bold">
                      {activeResult.asset_id}
                    </span>
                    <span className={`px-2 py-0.5 rounded text-xs font-bold ${
                      activeResult.asset_type === 'Repairable' 
                        ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-400/30'
                        : 'bg-amber-500/20 text-amber-300 border border-amber-400/30'
                    }`}>
                      {activeResult.asset_type === 'Repairable' ? 'Hệ thống Sửa chữa được (Repairable)' : 'Linh kiện Thay thế 1 lần (Non-repairable)'}
                    </span>
                  </div>
                  <h4 className="text-lg font-bold text-white">{activeResult.asset_name}</h4>
                  <p className="text-xs text-slate-300">
                    Kỳ đánh giá: <span className="font-semibold text-white">{activeResult.analysis_period_hours} giờ</span> · Số lần sự cố (Nf): <span className="font-semibold text-rose-300">{activeResult.total_failures} lần</span> · Thời gian chạy (Top): <span className="font-semibold text-emerald-300">{activeResult.operating_hours.toLocaleString()} giờ</span>
                  </p>
                </div>

                <div className="flex items-center gap-4 bg-slate-800/80 border border-slate-700 p-3 rounded-xl shrink-0">
                  <div className="text-center px-2">
                    <p className="text-[10px] text-slate-400 uppercase font-bold tracking-wider">Độ Sẵn Sàng (Ao)</p>
                    <p className={`text-2xl font-black ${
                      activeResult.reliability_metrics.operational_availability_percent >= 95 
                        ? 'text-emerald-400' 
                        : 'text-rose-400'
                    }`}>
                      {activeResult.reliability_metrics.operational_availability_percent}%
                    </p>
                  </div>
                  <div className="w-px h-10 bg-slate-700" />
                  <div className="text-center px-2">
                    <p className="text-[10px] text-slate-400 uppercase font-bold tracking-wider">Tổng COUR</p>
                    <p className="text-2xl font-black text-amber-400">
                      ${activeResult.cost_metrics.COUR_total.toLocaleString()}
                    </p>
                  </div>
                </div>
              </div>

              {/* MWT vs MTTR Warning Banner if MWT > MTTR */}
              {mwtExceeds && (
                <div className="bg-amber-50 border-l-4 border-amber-500 p-4 rounded-r-xl shadow-2xs flex items-start gap-3">
                  <AlertTriangle size={20} className="text-amber-600 shrink-0 mt-0.5" />
                  <div className="space-y-1">
                    <p className="text-xs font-bold text-amber-900 uppercase tracking-wide">
                      Cảnh Báo ARP-E: MWT ({activeResult.reliability_metrics.MWT_hours}h) Lớn Hơn MTTR ({activeResult.reliability_metrics.MTTR_hours}h) Gấp {mwtMttrRatio} Lần!
                    </p>
                    <p className="text-xs text-amber-800 leading-relaxed">
                      Thời gian chờ đợi bảo trì (chờ phụ tùng, cấp phép LOTO, phân công) đang cao hơn thời gian kỹ thuật viên trực tiếp sửa chữa. Cần tối ưu khâu <strong>Quản lý Kho vật tư (Spares Management)</strong> và <strong>Lập kế hoạch & Chuẩn bị (Job Kitting)</strong>.
                    </p>
                  </div>
                </div>
              )}

              {/* GROUP A: TIME & RELIABILITY METRICS */}
              <div className="space-y-3">
                <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                  <h5 className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center gap-2">
                    <Activity size={15} className="text-blue-600" />
                    A. Nhóm Chỉ Số Thời Gian & Độ Tin Cậy (Time & Reliability Metrics)
                  </h5>
                  <span className="text-[11px] text-slate-400 font-mono">Công thức chuẩn Mobius ARP-E</span>
                </div>

                <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
                  
                  {/* MTBF */}
                  <div className="bg-white p-3.5 rounded-xl border border-slate-200 shadow-2xs">
                    <div className="flex justify-between items-start">
                      <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">MTBF</span>
                      <TrendingUp size={14} className="text-emerald-500" />
                    </div>
                    <p className="text-xl font-black text-slate-900 mt-1">
                      {activeResult.reliability_metrics.MTBF_hours !== null 
                        ? `${activeResult.reliability_metrics.MTBF_hours.toLocaleString()} h` 
                        : 'N/A'}
                    </p>
                    <p className="text-[10px] text-slate-400 mt-0.5">Top / Nf (Repairable)</p>
                  </div>

                  {/* MTTF */}
                  <div className="bg-white p-3.5 rounded-xl border border-slate-200 shadow-2xs">
                    <div className="flex justify-between items-start">
                      <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">MTTF</span>
                      <Activity size={14} className="text-blue-500" />
                    </div>
                    <p className="text-xl font-black text-slate-900 mt-1">
                      {activeResult.reliability_metrics.MTTF_hours !== null 
                        ? `${activeResult.reliability_metrics.MTTF_hours.toLocaleString()} h` 
                        : (activeResult.asset_type === 'Repairable' ? 'Theo cụm' : '---')}
                    </p>
                    <p className="text-[10px] text-slate-400 mt-0.5">Non-repairable item</p>
                  </div>

                  {/* MTTR */}
                  <div className="bg-white p-3.5 rounded-xl border border-slate-200 shadow-2xs">
                    <div className="flex justify-between items-start">
                      <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">MTTR</span>
                      <TrendingDown size={14} className="text-emerald-600" />
                    </div>
                    <p className="text-xl font-black text-slate-900 mt-1">
                      {activeResult.reliability_metrics.MTTR_hours} <span className="text-xs font-normal text-slate-400">h</span>
                    </p>
                    <p className="text-[10px] text-slate-400 mt-0.5">Thời gian sửa chữa TB</p>
                  </div>

                  {/* MWT */}
                  <div className={`p-3.5 rounded-xl border shadow-2xs ${
                    mwtExceeds 
                      ? 'bg-amber-50/70 border-amber-300 text-amber-950' 
                      : 'bg-white border-slate-200'
                  }`}>
                    <div className="flex justify-between items-start">
                      <span className="text-[10px] font-bold text-amber-700 uppercase tracking-wider">MWT</span>
                      <Clock size={14} className="text-amber-500" />
                    </div>
                    <p className="text-xl font-black text-amber-900 mt-1">
                      {activeResult.reliability_metrics.MWT_hours} <span className="text-xs font-normal text-amber-700">h</span>
                    </p>
                    <p className="text-[10px] text-amber-700/80 mt-0.5">Thời gian chờ đợi TB</p>
                  </div>

                  {/* Unscheduled Downtime */}
                  <div className="bg-white p-3.5 rounded-xl border border-slate-200 shadow-2xs">
                    <div className="flex justify-between items-start">
                      <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Downtime Sự Cố</span>
                      <AlertTriangle size={14} className="text-rose-500" />
                    </div>
                    <p className="text-xl font-black text-rose-600 mt-1">
                      {activeResult.reliability_metrics.total_unscheduled_downtime_hours} <span className="text-xs font-normal text-slate-400">h</span>
                    </p>
                    <p className="text-[10px] text-slate-400 mt-0.5">Nf × (MTTR + MWT)</p>
                  </div>

                  {/* Operational Availability */}
                  <div className="bg-white p-3.5 rounded-xl border border-slate-200 shadow-2xs">
                    <div className="flex justify-between items-start">
                      <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Availability (Ao)</span>
                      <ShieldCheck size={14} className="text-emerald-500" />
                    </div>
                    <p className="text-xl font-black text-emerald-600 mt-1">
                      {activeResult.reliability_metrics.operational_availability_percent}%
                    </p>
                    <p className="text-[10px] text-slate-400 mt-0.5">Top / (Top + Downtime)</p>
                  </div>

                </div>
              </div>

              {/* GROUP B: COST METRICS */}
              <div className="space-y-3">
                <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                  <h5 className="text-xs font-bold text-slate-800 uppercase tracking-wider flex items-center gap-2">
                    <DollarSign size={15} className="text-emerald-600" />
                    B. Nhóm Chỉ Số Chi Phí Bảo Trì & Tổn Thất (Cost Metrics)
                  </h5>
                  <span className="text-[11px] text-slate-400 font-mono">Đơn vị: USD ($)</span>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
                  
                  {/* Direct Repair Cost */}
                  <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
                    <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Chi Phí Trực Tiếp (C_direct)</span>
                    <p className="text-xl font-black text-slate-900 mt-1">
                      ${activeResult.cost_metrics.total_direct_repair_cost.toLocaleString()}
                    </p>
                    <p className="text-[10px] text-slate-400 mt-0.5">Nhân công + Phụ tùng + Thầu ngoài</p>
                  </div>

                  {/* MCTR */}
                  <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
                    <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">MCTR (Cost / Failure)</span>
                    <p className="text-xl font-black text-slate-900 mt-1">
                      ${activeResult.cost_metrics.MCTR_per_failure.toLocaleString()}
                    </p>
                    <p className="text-[10px] text-slate-400 mt-0.5">Chi phí trung bình mỗi sự cố</p>
                  </div>

                  {/* Downtime Cost */}
                  <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
                    <span className="text-[10px] font-bold text-rose-600 uppercase tracking-wider">Chi Phí Dừng Máy (C_dt)</span>
                    <p className="text-xl font-black text-rose-600 mt-1">
                      ${activeResult.cost_metrics.total_downtime_cost.toLocaleString()}
                    </p>
                    <p className="text-[10px] text-slate-400 mt-0.5">Unscheduled DT × $500/h</p>
                  </div>

                  {/* Quality Loss */}
                  <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs">
                    <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Tổn Thất Chất Lượng</span>
                    <p className="text-xl font-black text-amber-600 mt-1">
                      ${activeResult.cost_metrics.total_quality_loss_cost.toLocaleString()}
                    </p>
                    <p className="text-[10px] text-slate-400 mt-0.5">Phế phẩm & hạ cấp sản phẩm</p>
                  </div>

                  {/* COUR Total */}
                  <div className="bg-slate-900 text-white p-4 rounded-xl shadow-sm border border-slate-800">
                    <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">COUR (Cost of Unreliability)</span>
                    <p className="text-xl font-black text-amber-400 mt-1">
                      ${activeResult.cost_metrics.COUR_total.toLocaleString()}
                    </p>
                    <p className="text-[10px] text-slate-400 mt-0.5">Tổng thiệt hại không tin cậy</p>
                  </div>

                </div>
              </div>

              {/* GROUP C: VALIDATION RULES & RECOMMENDATIONS */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                
                {/* Warnings */}
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-2">
                  <h6 className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-2">
                    <AlertTriangle size={14} className="text-amber-500" />
                    Kiểm Định Ràng Buộc & Cảnh Báo Hệ Thống
                  </h6>
                  <ul className="text-xs text-slate-600 space-y-2">
                    {activeResult.reliability_warnings.map((w, idx) => (
                      <li key={idx} className="flex items-start gap-2 bg-white p-2.5 rounded-lg border border-slate-200">
                        <span className="text-amber-600 font-bold">•</span>
                        <span>{w}</span>
                      </li>
                    ))}
                  </ul>
                </div>

                {/* Recommendations */}
                <div className="bg-blue-50/60 border border-blue-200 rounded-xl p-4 space-y-2">
                  <h6 className="text-xs font-bold text-blue-900 uppercase tracking-wider flex items-center gap-2">
                    <CheckCircle2 size={14} className="text-blue-600" />
                    Khuyến Nghị Kỹ Thuật Độ Tin Cậy (ARP-E)
                  </h6>
                  <ul className="text-xs text-blue-900 space-y-2">
                    {activeResult.recommendations.map((r, idx) => (
                      <li key={idx} className="flex items-start gap-2 bg-white p-2.5 rounded-lg border border-blue-200">
                        <span className="text-blue-600 font-bold">•</span>
                        <span>{r}</span>
                      </li>
                    ))}
                  </ul>
                </div>

              </div>
            </>
          )}

          {/* VIEW MODE 2: PARETO BAD ACTORS */}
          {viewMode === 'pareto' && (
            <div className="space-y-4">
              <div className="bg-slate-900 text-white p-4 rounded-xl flex items-center justify-between">
                <div>
                  <h4 className="font-bold text-sm">Nhận Diện Bad Actors theo Nguyên Tắc Pareto 80/20</h4>
                  <p className="text-xs text-slate-400 mt-0.5">
                    20% thiết bị chiếm tới 80% tổng tổn thất COUR (Cost of Unreliability). Đây là các mục tiêu ưu tiên xử lý dứt điểm.
                  </p>
                </div>
                <span className="px-3 py-1 bg-rose-600 text-white rounded-lg text-xs font-bold">
                  {fleetSummary.bad_actors.length} Thiết Bị Bad Actor
                </span>
              </div>

              <div className="overflow-x-auto border border-slate-200 rounded-xl">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 uppercase font-bold tracking-wider">
                    <tr>
                      <th className="p-3">Mã TB</th>
                      <th className="p-3">Tên thiết bị</th>
                      <th className="p-3">Nhà máy</th>
                      <th className="p-3 text-center">Số sự cố (Nf)</th>
                      <th className="p-3 text-center">MTTR</th>
                      <th className="p-3 text-center">MWT</th>
                      <th className="p-3 text-right">Downtime Sự Cố</th>
                      <th className="p-3 text-right">Tổng COUR</th>
                      <th className="p-3 text-center">Thao tác</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {fleetSummary.bad_actors.map((item, idx) => (
                      <tr key={item.asset_id} className="hover:bg-slate-50 transition-colors">
                        <td className="p-3 font-mono font-bold text-blue-700">
                          {item.asset_id}
                        </td>
                        <td className="p-3 font-semibold text-slate-900">
                          {item.asset_name}
                        </td>
                        <td className="p-3 text-slate-600">
                          {item.factory}
                        </td>
                        <td className="p-3 text-center">
                          <span className="px-2 py-0.5 rounded font-bold bg-rose-100 text-rose-700 border border-rose-200">
                            {item.total_failures}
                          </span>
                        </td>
                        <td className="p-3 text-center font-mono">
                          {item.reliability_metrics.MTTR_hours}h
                        </td>
                        <td className="p-3 text-center font-mono font-bold text-amber-700">
                          {item.reliability_metrics.MWT_hours}h
                        </td>
                        <td className="p-3 text-right font-mono font-bold text-rose-600">
                          {item.reliability_metrics.total_unscheduled_downtime_hours}h
                        </td>
                        <td className="p-3 text-right font-mono font-bold text-slate-900">
                          ${item.cost_metrics.COUR_total.toLocaleString()}
                        </td>
                        <td className="p-3 text-center">
                          <button
                            onClick={() => {
                              setCurrentAssetId(item.asset_id);
                              setViewMode('dashboard');
                            }}
                            className="px-2.5 py-1 bg-blue-50 text-blue-700 hover:bg-blue-100 rounded text-xs font-semibold"
                          >
                            Chi tiết
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* VIEW MODE 3: STANDARD JSON OUTPUT */}
          {viewMode === 'json' && (
            <div className="space-y-3">
              <div className="flex items-center justify-between">
                <div>
                  <h4 className="text-sm font-bold text-slate-800">
                    Định Dạng JSON Chuẩn Hóa Theo Đặc Tả ARP-E (Mục 4)
                  </h4>
                  <p className="text-xs text-slate-500">
                    Đối tượng: <span className="font-mono font-semibold">{activeResult.asset_id}</span>
                  </p>
                </div>
                <button
                  onClick={handleCopyJson}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-bold transition-colors shadow-2xs"
                >
                  {copied ? <Check size={14} /> : <Copy size={14} />}
                  <span>{copied ? 'Đã sao chép' : 'Sao chép JSON'}</span>
                </button>
              </div>

              <div className="relative bg-slate-950 rounded-xl p-4 overflow-x-auto border border-slate-800">
                <pre className="text-emerald-400 font-mono text-xs leading-relaxed">
                  {jsonContent}
                </pre>
              </div>
            </div>
          )}

        </div>

        {/* Modal Footer */}
        <div className="bg-slate-50 px-6 py-3 border-t border-slate-200 flex items-center justify-between text-xs text-slate-500">
          <div className="flex items-center gap-2">
            <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
            <span>Mô hình tính toán tự động cập nhật theo dữ liệu CMMS Work Orders & Equipment thực tế</span>
          </div>
          <button
            onClick={onClose}
            className="px-4 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg font-semibold transition-colors"
          >
            Đóng
          </button>
        </div>

      </div>
    </div>
  );
};
