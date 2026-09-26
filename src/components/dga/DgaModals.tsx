import React, { useState } from 'react';
import {
  X,
  AlertTriangle,
  CheckCircle2,
  Clock,
  Download,
  Flame,
  ShieldAlert,
  Sliders,
  FileText,
  Activity,
  Printer,
  RotateCcw,
  Zap,
  Info,
  ChevronRight,
  ShieldCheck,
  Check,
  ArrowUp,
  ArrowDown,
  ArrowRight
} from 'lucide-react';
import { DgaDataInput } from '../DgaAdvancedDiagnosticsDashboard';
import { computeDgaParamTrend } from './DgaParamTrendInput';

// IEEE C57.104 90th Percentile Typical Gas Baseline Limits (ppm) for Condition 1
export const IEEE_LIMITS: Record<string, { name: string; limit90: number; unit: string; desc: string }> = {
  h2: { name: 'Hydrogen (H₂)', limit90: 100, unit: 'ppm', desc: 'Phóng điện cục bộ (PD) hoặc vầng quang điện' },
  ch4: { name: 'Methane (CH₄)', limit90: 120, unit: 'ppm', desc: 'Quá nhiệt nhiệt độ thấp trong dầu (< 300°C)' },
  c2h6: { name: 'Ethane (C₂H₆)', limit90: 65, unit: 'ppm', desc: 'Quá nhiệt nhiệt độ thấp đến trung bình (300-500°C)' },
  c2h4: { name: 'Ethylene (C₂H₄)', limit90: 50, unit: 'ppm', desc: 'Quá nhiệt dầu nhiệt độ cao (> 700°C)' },
  c2h2: { name: 'Acetylene (C₂H₂)', limit90: 1.0, unit: 'ppm', desc: 'Khí hồ quang điện năng lượng cao (Arcing)' },
  co: { name: 'Carbon Monoxide (CO)', limit90: 350, unit: 'ppm', desc: 'Lão hóa hoặc quá nhiệt giấy cách điện xenlulo' },
  co2: { name: 'Carbon Dioxide (CO₂)', limit90: 2500, unit: 'ppm', desc: 'Sản phẩm oxy hóa dầu và lão hóa tự nhiên của giấy' },
  o2: { name: 'Oxygen (O₂)', limit90: 3000, unit: 'ppm', desc: 'Khí hòa tan từ không khí khí quyển' },
  n2: { name: 'Nitrogen (N₂)', limit90: 50000, unit: 'ppm', desc: 'Khí đệm bảo vệ mặt thoáng' }
};

// 1. DGA DETAILS MODAL
export const DgaDetailsModal: React.FC<{
  isOpen: boolean;
  onClose: () => void;
  currentGases?: DgaDataInput;
  currentFormValues?: DgaDataInput;
  analysisResult?: any;
  equipmentCode?: string;
  equipmentName?: string;
  sampleDate?: string;
  engineerName?: string;
  onReanalyze?: () => void;
  onUpdateFormValues?: (values: Partial<DgaDataInput>) => void;
  lastRecordedValues?: any;
}> = ({
  isOpen,
  onClose,
  currentGases,
  currentFormValues,
  analysisResult,
  equipmentCode = 'TR-01',
  equipmentName = 'MBA T1 110kV Bắc Ninh',
  sampleDate = '14/05/2025 08:15',
  engineerName = 'Nguyễn Văn Tuấn',
  onReanalyze,
  lastRecordedValues
}) => {
  if (!isOpen) return null;

  const data = currentGases || currentFormValues || {};
  const h2 = parseFloat(String(data.h2 || 0)) || 0;
  const ch4 = parseFloat(String(data.ch4 || 0)) || 0;
  const c2h6 = parseFloat(String(data.c2h6 || 0)) || 0;
  const c2h4 = parseFloat(String(data.c2h4 || 0)) || 0;
  const c2h2 = parseFloat(String(data.c2h2 || 0)) || 0;
  const co = parseFloat(String(data.co || 0)) || 0;
  const co2 = parseFloat(String(data.co2 || 0)) || 0;
  const o2 = parseFloat(String(data.o2 || 0)) || 0;
  const n2 = parseFloat(String(data.n2 || 0)) || 0;

  const tdcg = h2 + ch4 + c2h6 + c2h4 + c2h2 + co;
  const isHealthyCondition1 = analysisResult?.faultDiagnosisSkipped ?? (
    tdcg <= 720 && h2 < 100 && ch4 < 120 && c2h6 < 65 && c2h4 < 50 && c2h2 < 35 && co < 350
  );

  const gases = [
    { key: 'h2', val: h2 },
    { key: 'ch4', val: ch4 },
    { key: 'c2h6', val: c2h6 },
    { key: 'c2h4', val: c2h4 },
    { key: 'c2h2', val: c2h2 },
    { key: 'co', val: co },
    { key: 'co2', val: co2 },
    { key: 'o2', val: o2 },
    { key: 'n2', val: n2 }
  ];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-4 overflow-y-auto">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-3xl max-h-[90vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        {/* Header */}
        <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/80">
          <div>
            <h3 className="text-lg font-bold text-slate-900 flex items-center gap-2">
              <Activity className="text-blue-600" size={20} />
              Chi tiết Phân tích Khí Hòa tan (DGA Details)
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Thiết bị: <span className="font-semibold text-slate-700">{equipmentCode} — {equipmentName}</span> | Ngày: <span className="font-medium text-slate-700">{sampleDate}</span> | Kỹ sư: <span className="font-medium text-slate-700">{engineerName}</span>
            </p>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-full hover:bg-slate-200 text-slate-400 hover:text-slate-700 flex items-center justify-center transition-colors cursor-pointer"
          >
            <X size={18} />
          </button>
        </div>

        {/* Body */}
        <div className="p-6 overflow-y-auto space-y-6">
          {/* IEEE C57.104 Condition Banner */}
          {isHealthyCondition1 ? (
            <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-4 flex items-start gap-3">
              <ShieldCheck className="text-emerald-600 flex-shrink-0 mt-0.5" size={22} />
              <div>
                <h4 className="text-sm font-bold text-emerald-900">
                  Tình trạng máy: Condition 1 (Bình thường / Khỏe mạnh) — Tiêu chuẩn IEEE C57.104
                </h4>
                <p className="text-xs text-emerald-800 mt-1 leading-relaxed">
                  Tổng lượng khí cháy <strong>TDCG = {tdcg.toFixed(1)} ppm</strong> (≤ 720 ppm) và toàn bộ các khí riêng lẻ đều nằm dưới ngưỡng giới hạn. Thuật toán tự động bỏ qua phân tích các lỗi (Duval, Rogers, IEC) nhằm tránh báo động giả do nhiễu tạp âm nồng độ thấp. Tiếp tục vận hành bình thường và lấy mẫu định kỳ hàng năm.
                </p>
              </div>
            </div>
          ) : (
            <div className="bg-amber-50 border border-amber-200 rounded-xl p-4 flex items-start gap-3">
              <AlertTriangle className="text-amber-600 flex-shrink-0 mt-0.5" size={22} />
              <div>
                <h4 className="text-sm font-bold text-amber-900">
                  Cảnh báo: {analysisResult?.conditionName || (tdcg > 1920 ? 'Condition 3 / 4' : 'Condition 2')} — Khí cháy vượt ngưỡng
                </h4>
                <p className="text-xs text-amber-800 mt-1 leading-relaxed">
                  Tổng khí cháy <strong>TDCG = {tdcg.toFixed(1)} ppm</strong> vượt ngưỡng Condition 1. Hệ thống đã kích hoạt toàn bộ các phương pháp chẩn đoán đa ma trận (Duval Triangles, Pentagons, Rogers, IEC 60599).
                </p>
              </div>
            </div>
          )}

          {/* Quick Stats */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
            <div className="bg-blue-50 border border-blue-100 rounded-xl p-3.5">
              <span className="text-[11px] font-bold text-blue-600 uppercase tracking-wider">Tổng khí cháy (TDCG)</span>
              <div className="text-2xl font-black text-blue-950 mt-1">{tdcg.toFixed(1)} <span className="text-xs font-normal text-blue-700">ppm</span></div>
              <span className="text-[11px] text-blue-700 mt-1 block">
                {tdcg <= 720 ? 'Condition 1 (Bình thường ≤ 720)' : tdcg <= 1920 ? 'Condition 2 (721 - 1920)' : 'Condition 3/4 (> 1920)'}
              </span>
            </div>

            <div className={`border rounded-xl p-3.5 ${
              c2h2 >= 1.0 ? 'bg-rose-50 border-rose-200 text-rose-900' : 'bg-emerald-50 border-emerald-100 text-emerald-900'
            }`}>
              <span className="text-[11px] font-bold uppercase tracking-wider opacity-80">Khí hồ quang C₂H₂</span>
              <div className="text-2xl font-black mt-1">{c2h2} <span className="text-xs font-normal">ppm</span></div>
              <span className="text-[11px] opacity-80 mt-1 block">
                {c2h2 >= 1.0 ? 'Vượt ngưỡng báo động hồ quang (≥ 1.0 ppm)' : 'Dưới ngưỡng cảnh báo hồ quang (< 1.0 ppm)'}
              </span>
            </div>

            <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5">
              <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider">Tiêu chuẩn kiểm định</span>
              <div className="text-sm font-bold text-slate-800 mt-2">IEEE C57.104 &amp; IEC 60599</div>
              <span className="text-[11px] text-slate-500 mt-1 block">Baseline 90th &amp; 95th Percentile</span>
            </div>
          </div>

          {/* Gas Concentrations Table */}
          <div>
            <h4 className="text-sm font-bold text-slate-800 mb-2">Bảng nồng độ khí và đối chiếu ngưỡng giới hạn (ppm)</h4>
            <div className="overflow-x-auto border border-slate-200 rounded-xl">
              <table className="w-full text-xs text-left">
                <thead className="bg-slate-100/80 text-slate-600 font-semibold border-b border-slate-200">
                  <tr>
                    <th className="py-2.5 px-3">Tên khí</th>
                    <th className="py-2.5 px-3">Nồng độ thực tế</th>
                    <th className="py-2.5 px-3">Xu hướng vs Lịch sử</th>
                    <th className="py-2.5 px-3">Ngưỡng chuẩn (90th)</th>
                    <th className="py-2.5 px-3">Tỷ lệ so với chuẩn</th>
                    <th className="py-2.5 px-3">Đánh giá trạng thái</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {gases.map(g => {
                    const info = IEEE_LIMITS[g.key] || { name: g.key, limit90: 100, unit: 'ppm', desc: '' };
                    const ratio = info.limit90 > 0 ? (g.val / info.limit90) : 0;
                    const isExceeded = ratio > 1.0;
                    const isCritical = ratio > 2.0;

                    const histVal = lastRecordedValues?.gases?.[g.key];
                    const tr = computeDgaParamTrend(g.val, histVal, false);

                    return (
                      <tr key={g.key} className="hover:bg-slate-50/70">
                        <td className="py-2.5 px-3 font-semibold text-slate-800">
                          {info.name}
                          <span className="block text-[10px] text-slate-400 font-normal">{info.desc}</span>
                        </td>
                        <td className="py-2.5 px-3 font-bold text-slate-900">{g.val} {info.unit}</td>
                        <td className="py-2.5 px-3">
                          {tr.hasHistorical ? (
                            <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded font-mono text-[11px] font-bold border ${
                              tr.severity === 'worse'
                                ? 'bg-rose-50 text-rose-700 border-rose-200'
                                : tr.severity === 'better'
                                ? 'bg-emerald-50 text-emerald-700 border-emerald-200'
                                : 'bg-slate-100 text-slate-600 border-slate-200'
                            }`} title={`Kỳ trước: ${histVal} ${info.unit}`}>
                              {tr.direction === 'up' && <ArrowUp size={12} className="stroke-[2.5]" />}
                              {tr.direction === 'down' && <ArrowDown size={12} className="stroke-[2.5]" />}
                              {tr.direction === 'flat' && <ArrowRight size={12} />}
                              <span>{tr.badgeText}</span>
                            </span>
                          ) : (
                            <span className="text-slate-400 text-[10px] italic">Chưa có mẫu</span>
                          )}
                        </td>
                        <td className="py-2.5 px-3 text-slate-500">{info.limit90} {info.unit}</td>
                        <td className="py-2.5 px-3">
                          <div className="flex items-center gap-2">
                            <div className="w-16 bg-slate-200 h-1.5 rounded-full overflow-hidden">
                              <div
                                className={`h-full ${isCritical ? 'bg-rose-500' : isExceeded ? 'bg-amber-500' : 'bg-emerald-500'}`}
                                style={{ width: `${Math.min(100, ratio * 50)}%` }}
                              />
                            </div>
                            <span className="text-[11px] font-mono text-slate-600">{(ratio * 100).toFixed(0)}%</span>
                          </div>
                        </td>
                        <td className="py-2.5 px-3">
                          <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold ${
                            isCritical ? 'bg-rose-100 text-rose-700' : isExceeded ? 'bg-amber-100 text-amber-700' : 'bg-emerald-100 text-emerald-700'
                          }`}>
                            {isCritical ? 'Báo động (Alarm)' : isExceeded ? 'Vượt ngưỡng (Alert)' : 'Bình thường (Normal)'}
                          </span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>

          {/* Oil Parameters */}
          <div className="bg-slate-50 rounded-xl p-4 border border-slate-200">
            <h4 className="text-xs font-bold text-slate-700 uppercase tracking-wider mb-2">Chỉ tiêu hóa lý dầu cách điện</h4>
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
              <div>
                <span className="text-slate-500 block">Hàm lượng ẩm:</span>
                <span className="font-bold text-slate-800">{data.moisture || '12'} ppm</span>
              </div>
              <div>
                <span className="text-slate-500 block">Điện áp đánh thủng:</span>
                <span className="font-bold text-slate-800">{data.bdStrength || '58'} kV</span>
              </div>
              <div>
                <span className="text-slate-500 block">Chỉ số axit:</span>
                <span className="font-bold text-slate-800">{data.acidity || '0.04'} mg KOH/g</span>
              </div>
              <div>
                <span className="text-slate-500 block">Độ trùng hợp (Est. DP):</span>
                <span className="font-bold text-slate-800">{data.estDp || '650'}</span>
              </div>
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="px-6 py-3 border-t border-slate-200 bg-slate-50 flex items-center justify-between">
          <div className="text-xs text-slate-500">
            Trạng thái máy: <span className="font-bold text-slate-700">{isHealthyCondition1 ? 'Condition 1 (Bình thường)' : analysisResult?.condition || 'Cần theo dõi'}</span>
          </div>
          <div className="flex items-center gap-2">
            {onReanalyze && (
              <button
                type="button"
                onClick={onReanalyze}
                className="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-colors shadow-xs cursor-pointer"
              >
                <RotateCcw size={13} />
                Chạy lại phân tích DGA
              </button>
            )}
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg text-xs font-semibold transition-colors cursor-pointer"
            >
              Đóng
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

// 2. RATIO SETTINGS MODAL
export const RatioSettingsModal: React.FC<{
  isOpen: boolean;
  onClose: () => void;
}> = ({ isOpen, onClose }) => {
  const [settings, setSettings] = useState([
    { id: 'ch4_h2', name: 'CH₄ / H₂', standard: 'IEC 60599 / IEEE', alert: '0.1', alarm: '1.0', desc: 'Chỉ số quá nhiệt dầu hoặc phóng điện cục bộ' },
    { id: 'c2h2_c2h4', name: 'C₂H₂ / C₂H₄', standard: 'IEC 60599', alert: '0.1', alarm: '1.0', desc: 'Chỉ số hồ quang điện hoặc phóng điện năng lượng cao' },
    { id: 'c2h4_c2h6', name: 'C₂H₄ / C₂H₆', standard: 'IEC 60599', alert: '1.0', alarm: '3.0', desc: 'Chỉ số bẻ gãy nhiệt phân dầu ở nhiệt độ cao' },
    { id: 'c2h2_ch4', name: 'C₂H₂ / CH₄', standard: 'IEEE C57.104', alert: '0.1', alarm: '0.3', desc: 'Phân định hồ quang so với phân hủy nhiệt thông thường' },
    { id: 'co2_co', name: 'CO₂ / CO', standard: 'IEEE C57.104', alert: '7.0', alarm: '3.0', desc: 'Lão hóa cách điện giấy xenlulo (< 3 là thoái hóa nghiêm trọng)' }
  ]);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-4 overflow-y-auto">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-2xl max-h-[90vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/80">
          <div>
            <h3 className="text-lg font-bold text-slate-900 flex items-center gap-2">
              <Sliders className="text-blue-600" size={20} />
              Cấu hình Ngưỡng Cảnh báo Tỷ số Khí (Ratio Settings)
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Thiết lập ngưỡng cảnh báo Level 1 (Alarm) và Level 2 (Alert) theo IEC 60599 &amp; IEEE C57.104
            </p>
          </div>
          <button onClick={onClose} className="w-8 h-8 rounded-full hover:bg-slate-200 text-slate-400 hover:text-slate-700 flex items-center justify-center cursor-pointer">
            <X size={18} />
          </button>
        </div>

        <div className="p-6 overflow-y-auto space-y-4">
          <div className="bg-blue-50 border border-blue-200 rounded-xl p-3 text-xs text-blue-900 flex items-start gap-2">
            <Info className="text-blue-600 flex-shrink-0 mt-0.5" size={16} />
            <span>
              <strong>Quy tắc chuẩn IEEE C57.104:</strong> Khi máy biến áp ở Condition 1 (TDCG ≤ 720 ppm), các tỷ số khí chỉ mang tính tham khảo và thuật toán sẽ tự động bỏ qua phân tích lỗi để tránh báo động giả do nhiễu ở nồng độ cực thấp.
            </span>
          </div>

          <div className="space-y-3">
            {settings.map((item, idx) => (
              <div key={item.id} className="border border-slate-200 rounded-xl p-3.5 bg-slate-50/50">
                <div className="flex items-center justify-between mb-2">
                  <div>
                    <span className="font-bold text-slate-800 text-sm">{item.name}</span>
                    <span className="ml-2 text-[10px] bg-slate-200 text-slate-700 px-2 py-0.5 rounded-full font-semibold">{item.standard}</span>
                  </div>
                  <span className="text-[11px] text-slate-500">{item.desc}</span>
                </div>
                <div className="grid grid-cols-2 gap-3 mt-2">
                  <div>
                    <label className="block text-[11px] font-bold text-amber-700 mb-1">Ngưỡng Cảnh báo (Alert Level 2)</label>
                    <input
                      type="text"
                      value={item.alert}
                      onChange={e => {
                        const copy = [...settings];
                        copy[idx].alert = e.target.value;
                        setSettings(copy);
                      }}
                      className="w-full text-xs font-mono font-bold px-2.5 py-1.5 bg-white border border-amber-300 rounded-lg focus:ring-2 focus:ring-amber-500 outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-[11px] font-bold text-rose-700 mb-1">Ngưỡng Báo động (Alarm Level 1)</label>
                    <input
                      type="text"
                      value={item.alarm}
                      onChange={e => {
                        const copy = [...settings];
                        copy[idx].alarm = e.target.value;
                        setSettings(copy);
                      }}
                      className="w-full text-xs font-mono font-bold px-2.5 py-1.5 bg-white border border-rose-300 rounded-lg focus:ring-2 focus:ring-rose-500 outline-none"
                    />
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>

        <div className="px-6 py-3 border-t border-slate-200 bg-slate-50 flex items-center justify-between">
          <button
            type="button"
            onClick={() => {
              setSettings([
                { id: 'ch4_h2', name: 'CH₄ / H₂', standard: 'IEC 60599 / IEEE', alert: '0.1', alarm: '1.0', desc: 'Chỉ số quá nhiệt dầu hoặc phóng điện cục bộ' },
                { id: 'c2h2_c2h4', name: 'C₂H₂ / C₂H₄', standard: 'IEC 60599', alert: '0.1', alarm: '1.0', desc: 'Chỉ số hồ quang điện hoặc phóng điện năng lượng cao' },
                { id: 'c2h4_c2h6', name: 'C₂H₄ / C₂H₆', standard: 'IEC 60599', alert: '1.0', alarm: '3.0', desc: 'Chỉ số bẻ gãy nhiệt phân dầu ở nhiệt độ cao' },
                { id: 'c2h2_ch4', name: 'C₂H₂ / CH₄', standard: 'IEEE C57.104', alert: '0.1', alarm: '0.3', desc: 'Phân định hồ quang so với phân hủy nhiệt thông thường' },
                { id: 'co2_co', name: 'CO₂ / CO', standard: 'IEEE C57.104', alert: '7.0', alarm: '3.0', desc: 'Lão hóa cách điện giấy xenlulo' }
              ]);
            }}
            className="text-xs text-slate-500 hover:text-slate-700 font-medium cursor-pointer"
          >
            Khôi phục mặc định tiêu chuẩn
          </button>
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-semibold transition-colors cursor-pointer"
          >
            Áp dụng và Đóng
          </button>
        </div>
      </div>
    </div>
  );
};

// 3. RISK ASSESSMENT MODAL
export const RiskAssessmentModal: React.FC<{
  isOpen: boolean;
  onClose: () => void;
  riskScore?: number;
  currentGases?: DgaDataInput;
  currentFormValues?: DgaDataInput;
  analysisResult?: any;
  equipmentCode?: string;
  equipmentName?: string;
  sampleDate?: string;
  engineerName?: string;
}> = ({
  isOpen,
  onClose,
  riskScore = 20,
  currentGases,
  currentFormValues,
  analysisResult,
  equipmentCode = 'TR-01',
  equipmentName = 'MBA T1 110kV Bắc Ninh',
  sampleDate = '14/05/2025 08:15',
  engineerName = 'Nguyễn Văn Tuấn'
}) => {
  if (!isOpen) return null;

  const data = currentGases || currentFormValues || {};
  const tdcg = (
    (parseFloat(String(data.h2 || 0)) || 0) +
    (parseFloat(String(data.ch4 || 0)) || 0) +
    (parseFloat(String(data.c2h6 || 0)) || 0) +
    (parseFloat(String(data.c2h4 || 0)) || 0) +
    (parseFloat(String(data.c2h2 || 0)) || 0) +
    (parseFloat(String(data.co || 0)) || 0)
  );

  const isCondition1 = analysisResult?.faultDiagnosisSkipped ?? (tdcg <= 720);
  const adjustedRiskScore = isCondition1 ? Math.min(riskScore, 25) : riskScore;
  const currentHI = analysisResult?.currentHealthScore || (isCondition1 ? '2.1' : '3.8');
  const futureHI = analysisResult?.futureHealthScore || (isCondition1 ? '4.2' : '6.5');
  const pof = (0.00073 * Math.exp(0.5 * parseFloat(currentHI))).toFixed(4);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-4 overflow-y-auto">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-2xl max-h-[90vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/80">
          <div>
            <h3 className="text-lg font-bold text-slate-900 flex items-center gap-2">
              <ShieldAlert className="text-blue-600" size={20} />
              Đánh giá Rủi ro &amp; Chỉ số Sức khỏe (Risk Assessment)
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Thiết bị: <span className="font-semibold text-slate-700">{equipmentCode} — {equipmentName}</span> | Mô hình CNAIM &amp; IEEE C57.104
            </p>
          </div>
          <button onClick={onClose} className="w-8 h-8 rounded-full hover:bg-slate-200 text-slate-400 hover:text-slate-700 flex items-center justify-center cursor-pointer">
            <X size={18} />
          </button>
        </div>

        <div className="p-6 overflow-y-auto space-y-6">
          {/* Risk Level Banner */}
          <div className="flex items-center gap-5 p-4 rounded-xl bg-slate-50 border border-slate-200">
            <div className="relative w-20 h-20 flex-shrink-0 flex items-center justify-center">
              <svg className="w-20 h-20 transform -rotate-90">
                <circle cx="40" cy="40" r="32" stroke="#e2e8f0" strokeWidth="6" fill="transparent" />
                <circle
                  cx="40"
                  cy="40"
                  r="32"
                  stroke={adjustedRiskScore > 65 ? '#ef4444' : adjustedRiskScore > 40 ? '#f59e0b' : '#10b981'}
                  strokeWidth="6"
                  fill="transparent"
                  strokeDasharray="201"
                  strokeDashoffset={201 - (201 * adjustedRiskScore) / 100}
                  strokeLinecap="round"
                />
              </svg>
              <div className="absolute flex flex-col items-center">
                <span className="text-xl font-black text-slate-900 leading-none">{adjustedRiskScore}</span>
                <span className="text-[9px] font-bold text-slate-400 uppercase">Điểm</span>
              </div>
            </div>

            <div>
              <div className="flex items-center gap-2">
                <h4 className="text-base font-bold text-slate-900">Mức độ rủi ro:</h4>
                <span className={`px-2.5 py-0.5 rounded-full text-xs font-bold ${
                  adjustedRiskScore > 65 ? 'bg-rose-100 text-rose-700' : adjustedRiskScore > 40 ? 'bg-amber-100 text-amber-700' : 'bg-emerald-100 text-emerald-700'
                }`}>
                  {adjustedRiskScore > 65 ? 'Rủi ro Cao (High Risk)' : adjustedRiskScore > 40 ? 'Rủi ro Trung bình (Moderate Risk)' : 'Rủi ro Thấp (Low Risk)'}
                </span>
              </div>
              <p className="text-xs text-slate-600 mt-1">
                {isCondition1
                  ? 'Máy biến áp ở Condition 1 (Khỏe mạnh). TDCG ≤ 720 ppm, không có hoạt tính sinh khí hoặc phóng điện.'
                  : `Tình trạng tổng thể: ${analysisResult?.condition || 'Cần theo dõi'}. Có khí cháy hoặc tỷ số vượt ngưỡng.`}
              </p>
            </div>
          </div>

          {/* Health Index Matrix */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div className="border border-slate-200 rounded-xl p-4 bg-white">
              <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">Chỉ số Sức khỏe Hiện tại (HI)</span>
              <div className="text-3xl font-black text-blue-600 mt-2">{currentHI} <span className="text-sm font-normal text-slate-400">/ 10</span></div>
              <span className="text-xs text-slate-600 mt-1 block">Xác suất sự cố hàng năm (PoF): <strong>{pof}</strong> ({(parseFloat(pof) * 100).toFixed(2)}%)</span>
            </div>

            <div className="border border-slate-200 rounded-xl p-4 bg-white">
              <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">Dự báo Sức khỏe Sau 10 năm</span>
              <div className="text-3xl font-black text-slate-800 mt-2">{futureHI} <span className="text-sm font-normal text-slate-400">/ 10</span></div>
              <span className="text-xs text-slate-600 mt-1 block">Theo mô hình suy biến nhiệt và lão hóa dầu CNAIM</span>
            </div>
          </div>

          {/* Action Recommendations */}
          <div>
            <h4 className="text-sm font-bold text-slate-800 mb-2">Khuyến nghị Hành động Kỹ thuật</h4>
            <ul className="space-y-2 text-xs text-slate-700">
              <li className="flex items-start gap-2 p-2.5 bg-slate-50 rounded-lg border border-slate-200">
                <CheckCircle2 size={16} className="text-emerald-600 flex-shrink-0 mt-0.5" />
                <span>
                  <strong>Chu kỳ lấy mẫu thí nghiệm:</strong> {isCondition1 ? 'Duy trì chu kỳ lấy mẫu định kỳ 12 tháng/lần.' : 'Rút ngắn chu kỳ lấy mẫu DGA định kỳ xuống còn 1 - 3 tháng/lần.'}
                </span>
              </li>
              <li className="flex items-start gap-2 p-2.5 bg-slate-50 rounded-lg border border-slate-200">
                <CheckCircle2 size={16} className="text-emerald-600 flex-shrink-0 mt-0.5" />
                <span><strong>Chế độ vận hành:</strong> Vận hành bình thường theo tải định mức, tiếp tục theo dõi nhiệt độ lớp dầu đỉnh và cuộn dây.</span>
              </li>
            </ul>
          </div>
        </div>

        <div className="px-6 py-3 border-t border-slate-200 bg-slate-50 flex items-center justify-end">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-1.5 bg-slate-800 hover:bg-slate-900 text-white rounded-lg text-xs font-semibold transition-colors cursor-pointer"
          >
            Đóng
          </button>
        </div>
      </div>
    </div>
  );
};

// 4. RATIO ANALYSIS DEEP-DIVE MODAL
export const RatioAnalysisModal: React.FC<{
  isOpen: boolean;
  onClose: () => void;
  currentGases?: DgaDataInput;
  currentFormValues?: DgaDataInput;
  analysisResult?: any;
  equipmentCode?: string;
  equipmentName?: string;
}> = ({
  isOpen,
  onClose,
  currentGases,
  currentFormValues,
  analysisResult,
  equipmentCode = 'TR-01',
  equipmentName = 'MBA T1 110kV Bắc Ninh'
}) => {
  if (!isOpen) return null;

  const data = currentGases || currentFormValues || {};
  const h2 = parseFloat(String(data.h2 || 0)) || 0;
  const ch4 = parseFloat(String(data.ch4 || 0)) || 0;
  const c2h6 = parseFloat(String(data.c2h6 || 0)) || 0;
  const c2h4 = parseFloat(String(data.c2h4 || 0)) || 0;
  const c2h2 = parseFloat(String(data.c2h2 || 0)) || 0;
  const co = parseFloat(String(data.co || 0)) || 0;
  const co2 = parseFloat(String(data.co2 || 0)) || 0;

  const r1 = h2 > 0 ? (ch4 / h2).toFixed(3) : '0';
  const r2 = c2h4 > 0 ? (c2h2 / c2h4).toFixed(3) : '0';
  const r3 = c2h6 > 0 ? (c2h4 / c2h6).toFixed(3) : '0';
  const r4 = ch4 > 0 ? (c2h2 / ch4).toFixed(3) : '0';
  const r5 = co > 0 ? (co2 / co).toFixed(2) : '0';

  const tdcg = h2 + ch4 + c2h6 + c2h4 + c2h2 + co;
  const isHealthyCondition1 = analysisResult?.faultDiagnosisSkipped ?? (
    tdcg <= 720 && h2 < 100 && ch4 < 120 && c2h6 < 65 && c2h4 < 50 && c2h2 < 35 && co < 350
  );

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-4 overflow-y-auto">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-3xl max-h-[90vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/80">
          <div>
            <h3 className="text-lg font-bold text-slate-900 flex items-center gap-2">
              <Zap className="text-blue-600" size={20} />
              Phân tích Tương quan &amp; Đối chiếu Tỷ số Khí (Gas Ratios Analysis)
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Thiết bị: <span className="font-semibold text-slate-700">{equipmentCode} — {equipmentName}</span> | Rogers, IEC 60599, Dornenburg, CO₂/CO
            </p>
          </div>
          <button onClick={onClose} className="w-8 h-8 rounded-full hover:bg-slate-200 text-slate-400 hover:text-slate-700 flex items-center justify-center cursor-pointer">
            <X size={18} />
          </button>
        </div>

        <div className="p-6 overflow-y-auto space-y-6">
          {/* Condition 1 Notice */}
          {isHealthyCondition1 && (
            <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3.5 flex items-start gap-2.5 text-xs text-emerald-900">
              <CheckCircle2 className="text-emerald-600 flex-shrink-0 mt-0.5" size={17} />
              <div>
                <strong>Lưu ý quan trọng theo IEEE C57.104 &amp; IEC 60599:</strong>
                <p className="mt-0.5 text-emerald-800">
                  Do thiết bị đang ở Condition 1 (TDCG = {tdcg.toFixed(1)} ppm ≤ 720 ppm), toàn bộ khí nằm trong vùng giới hạn an toàn. Các tỷ số bên dưới chỉ được tính toán nhằm mục đích theo dõi xu hướng, thuật toán bỏ qua suy luận lỗi để không tạo ra cảnh báo giả.
                </p>
              </div>
            </div>
          )}

          {/* Table of Ratios */}
          <div>
            <h4 className="text-sm font-bold text-slate-800 mb-2">Giá trị tỷ số khí thực tế của mẫu đo</h4>
            <div className="overflow-x-auto border border-slate-200 rounded-xl">
              <table className="w-full text-xs text-left">
                <thead className="bg-slate-100/80 text-slate-600 font-semibold border-b border-slate-200">
                  <tr>
                    <th className="py-2.5 px-3">Tỷ số khí</th>
                    <th className="py-2.5 px-3">Công thức</th>
                    <th className="py-2.5 px-3">Giá trị tính toán</th>
                    <th className="py-2.5 px-3">Phạm vi chuẩn</th>
                    <th className="py-2.5 px-3">Ý nghĩa vật lý</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  <tr className="hover:bg-slate-50">
                    <td className="py-2.5 px-3 font-bold text-slate-800">CH₄ / H₂</td>
                    <td className="py-2.5 px-3 font-mono text-slate-500">Methane / Hydrogen</td>
                    <td className="py-2.5 px-3 font-bold text-blue-600">{r1}</td>
                    <td className="py-2.5 px-3 text-slate-500">0.1 - 1.0 (Bình thường)</td>
                    <td className="py-2.5 px-3 text-slate-600">&lt; 0.1: Phóng điện PD | &gt; 1.0: Quá nhiệt dầu T1</td>
                  </tr>
                  <tr className="hover:bg-slate-50">
                    <td className="py-2.5 px-3 font-bold text-slate-800">C₂H₂ / C₂H₄</td>
                    <td className="py-2.5 px-3 font-mono text-slate-500">Acetylene / Ethylene</td>
                    <td className="py-2.5 px-3 font-bold text-blue-600">{r2}</td>
                    <td className="py-2.5 px-3 text-slate-500">&lt; 0.1 (Không hồ quang)</td>
                    <td className="py-2.5 px-3 text-slate-600">&gt; 0.1: Hồ quang / phóng điện D1-D2</td>
                  </tr>
                  <tr className="hover:bg-slate-50">
                    <td className="py-2.5 px-3 font-bold text-slate-800">C₂H₄ / C₂H₆</td>
                    <td className="py-2.5 px-3 font-mono text-slate-500">Ethylene / Ethane</td>
                    <td className="py-2.5 px-3 font-bold text-blue-600">{r3}</td>
                    <td className="py-2.5 px-3 text-slate-500">&lt; 1.0 (Nhiệt độ thấp)</td>
                    <td className="py-2.5 px-3 text-slate-600">&gt; 3.0: Quá nhiệt độ cao &gt; 700°C (T3)</td>
                  </tr>
                  <tr className="hover:bg-slate-50">
                    <td className="py-2.5 px-3 font-bold text-slate-800">C₂H₂ / CH₄</td>
                    <td className="py-2.5 px-3 font-mono text-slate-500">Acetylene / Methane</td>
                    <td className="py-2.5 px-3 font-bold text-blue-600">{r4}</td>
                    <td className="py-2.5 px-3 text-slate-500">&lt; 0.1</td>
                    <td className="py-2.5 px-3 text-slate-600">Phân định phóng điện so với nhiệt phân</td>
                  </tr>
                  <tr className="hover:bg-slate-50">
                    <td className="py-2.5 px-3 font-bold text-slate-800">CO₂ / CO</td>
                    <td className="py-2.5 px-3 font-mono text-slate-500">Carbon Dioxide / CO</td>
                    <td className="py-2.5 px-3 font-bold text-blue-600">{r5}</td>
                    <td className="py-2.5 px-3 text-slate-500">3.0 - 10.0 (Giấy tốt)</td>
                    <td className="py-2.5 px-3 text-slate-600">&lt; 3.0: Thoái hóa giấy xenlulo nghiêm trọng</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          {/* Methods interpretation results */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
            <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5">
              <h5 className="text-[11px] font-bold text-slate-700 uppercase tracking-wider mb-1">Rogers Ratio (IEEE)</h5>
              <div className="text-sm font-bold text-blue-700">
                {isHealthyCondition1 ? 'Normal (Bỏ qua do Condition 1)' : (analysisResult?.matrix?.['IEEE (Rogers)'] || 'Normal')}
              </div>
              <p className="text-[11px] text-slate-500 mt-1">Dựa trên mã bộ 3 tỷ số R1, R2, R5.</p>
            </div>

            <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5">
              <h5 className="text-[11px] font-bold text-slate-700 uppercase tracking-wider mb-1">Tỷ số IEC 60599</h5>
              <div className="text-sm font-bold text-blue-700">
                {isHealthyCondition1 ? 'Normal (Bỏ qua do Condition 1)' : (analysisResult?.matrix?.['IEC Ratio'] || 'Normal')}
              </div>
              <p className="text-[11px] text-slate-500 mt-1">Chuẩn hóa theo IEC 60599 Table 1.</p>
            </div>

            <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5">
              <h5 className="text-[11px] font-bold text-slate-700 uppercase tracking-wider mb-1">Tỷ số Dornenburg</h5>
              <div className="text-sm font-bold text-blue-700">
                {isHealthyCondition1 ? 'Normal (Bỏ qua do Condition 1)' : (analysisResult?.matrix?.['IEEE (Dornenburg)'] || 'Normal')}
              </div>
              <p className="text-[11px] text-slate-500 mt-1">Áp dụng ngưỡng L1 trước khi suy luận.</p>
            </div>
          </div>
        </div>

        <div className="px-6 py-3 border-t border-slate-200 bg-slate-50 flex items-center justify-end">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-1.5 bg-slate-800 hover:bg-slate-900 text-white rounded-lg text-xs font-semibold transition-colors cursor-pointer"
          >
            Đóng
          </button>
        </div>
      </div>
    </div>
  );
};

// 5. SAMPLE HISTORY MODAL
export const SampleHistoryModal: React.FC<{
  isOpen: boolean;
  onClose: () => void;
  equipmentCode?: string;
  equipmentName?: string;
  onLoadSample: (sample: any) => void;
}> = ({
  isOpen,
  onClose,
  equipmentCode = 'TR-01',
  equipmentName = 'MBA T1 110kV Bắc Ninh',
  onLoadSample
}) => {
  const historySamples = [
    { date: '14/05/2025 08:15', sampleId: 'DGA-2025-05-14', engineer: 'Nguyễn Văn Tuấn', h2: 24, ch4: 38, c2h6: 18, c2h4: 45, c2h2: 0.8, co: 290, co2: 2150, tdcg: 415.8, status: 'Normal' },
    { date: '12/05/2025 10:30', sampleId: 'DGA-2025-05-12', engineer: 'Nguyễn Văn Tuấn', h2: 22, ch4: 35, c2h6: 17, c2h4: 42, c2h2: 0.6, co: 280, co2: 2100, tdcg: 396.6, status: 'Normal' },
    { date: '10/05/2025 14:00', sampleId: 'DGA-2025-05-10', engineer: 'Trần Văn Hoàng', h2: 19, ch4: 30, c2h6: 15, c2h4: 38, c2h2: 0.4, co: 275, co2: 2050, tdcg: 377.4, status: 'Normal' },
    { date: '08/05/2025 09:15', sampleId: 'DGA-2025-05-08', engineer: 'Nguyễn Văn Tuấn', h2: 18, ch4: 28, c2h6: 14, c2h4: 35, c2h2: 0.3, co: 260, co2: 2000, tdcg: 355.3, status: 'Normal' },
    { date: '01/05/2025 15:45', sampleId: 'DGA-2025-05-01', engineer: 'Lê Thanh Bình', h2: 15, ch4: 25, c2h6: 12, c2h4: 30, c2h2: 0.2, co: 250, co2: 1950, tdcg: 332.2, status: 'Normal' },
    { date: '15/04/2025 08:30', sampleId: 'DGA-2025-04-15', engineer: 'Nguyễn Văn Tuấn', h2: 14, ch4: 22, c2h6: 11, c2h4: 28, c2h2: 0.1, co: 240, co2: 1900, tdcg: 315.1, status: 'Normal' }
  ];

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-4 overflow-y-auto">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-4xl max-h-[90vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/80">
          <div>
            <h3 className="text-lg font-bold text-slate-900 flex items-center gap-2">
              <Clock className="text-blue-600" size={20} />
              Lịch sử Lấy mẫu Thí nghiệm DGA (Sample History)
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Thiết bị: <span className="font-semibold text-slate-700">{equipmentCode} — {equipmentName}</span> | Chọn một mẫu để nạp nhanh vào phân tích
            </p>
          </div>
          <button onClick={onClose} className="w-8 h-8 rounded-full hover:bg-slate-200 text-slate-400 hover:text-slate-700 flex items-center justify-center cursor-pointer">
            <X size={18} />
          </button>
        </div>

        <div className="p-6 overflow-y-auto">
          <div className="overflow-x-auto border border-slate-200 rounded-xl">
            <table className="w-full text-xs text-left">
              <thead className="bg-slate-100/80 text-slate-600 font-semibold border-b border-slate-200">
                <tr>
                  <th className="py-2.5 px-3">Thời gian</th>
                  <th className="py-2.5 px-3">Mã mẫu</th>
                  <th className="py-2.5 px-3">Kỹ sư lấy mẫu</th>
                  <th className="py-2.5 px-2">H₂</th>
                  <th className="py-2.5 px-2">CH₄</th>
                  <th className="py-2.5 px-2">C₂H₄</th>
                  <th className="py-2.5 px-2">C₂H₂</th>
                  <th className="py-2.5 px-2">TDCG</th>
                  <th className="py-2.5 px-3 text-right">Thao tác</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {historySamples.map(sample => (
                  <tr key={sample.sampleId} className="hover:bg-slate-50/80">
                    <td className="py-2.5 px-3 font-semibold text-slate-800">{sample.date}</td>
                    <td className="py-2.5 px-3 font-mono text-slate-500">{sample.sampleId}</td>
                    <td className="py-2.5 px-3 text-slate-600">{sample.engineer}</td>
                    <td className="py-2.5 px-2 font-mono">{sample.h2}</td>
                    <td className="py-2.5 px-2 font-mono">{sample.ch4}</td>
                    <td className="py-2.5 px-2 font-mono">{sample.c2h4}</td>
                    <td className="py-2.5 px-2 font-mono font-bold text-amber-600">{sample.c2h2}</td>
                    <td className="py-2.5 px-2 font-mono font-bold text-blue-600">{sample.tdcg}</td>
                    <td className="py-2.5 px-3 text-right">
                      <button
                        type="button"
                        onClick={() => {
                          onLoadSample(sample);
                          onClose();
                        }}
                        className="px-2.5 py-1 bg-blue-50 hover:bg-blue-100 text-blue-700 rounded-lg text-[11px] font-bold transition-colors cursor-pointer"
                      >
                        Nạp vào Form
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        <div className="px-6 py-3 border-t border-slate-200 bg-slate-50 flex items-center justify-end">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-1.5 bg-slate-800 hover:bg-slate-900 text-white rounded-lg text-xs font-semibold transition-colors cursor-pointer"
          >
            Đóng
          </button>
        </div>
      </div>
    </div>
  );
};

// 6. TRANSFORMER OVERVIEW REPORT MODAL (Complete 4-Part Structure)
export const TransformerOverviewReportModal: React.FC<{
  isOpen: boolean;
  onClose: () => void;
  equipmentCode?: string;
  equipmentName?: string;
  customerName?: string;
  sampleDate?: string;
  engineerName?: string;
  currentGases?: DgaDataInput;
  currentFormValues?: DgaDataInput;
  analysisResult?: any;
  riskScore?: number;
  onDownloadPdf?: () => void;
}> = ({
  isOpen,
  onClose,
  equipmentCode = 'TR-01',
  equipmentName = 'MBA T1 110kV Bắc Ninh',
  customerName = 'EVN / NPC',
  sampleDate = '14/05/2025 08:15',
  engineerName = 'Nguyễn Văn Tuấn',
  currentGases,
  currentFormValues,
  analysisResult,
  riskScore = 20,
  onDownloadPdf
}) => {
  if (!isOpen) return null;

  const data = currentGases || currentFormValues || {};
  const h2 = parseFloat(String(data.h2 || 0)) || 0;
  const ch4 = parseFloat(String(data.ch4 || 0)) || 0;
  const c2h6 = parseFloat(String(data.c2h6 || 0)) || 0;
  const c2h4 = parseFloat(String(data.c2h4 || 0)) || 0;
  const c2h2 = parseFloat(String(data.c2h2 || 0)) || 0;
  const co = parseFloat(String(data.co || 0)) || 0;
  const co2 = parseFloat(String(data.co2 || 0)) || 0;

  const tdcg = (h2 + ch4 + c2h6 + c2h4 + c2h2 + co).toFixed(1);
  const isCondition1 = analysisResult?.faultDiagnosisSkipped ?? (parseFloat(tdcg) <= 720);

  // Structured reports from engine if available
  const report = analysisResult?.reportSections;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-4 overflow-y-auto">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-4xl max-h-[92vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/80">
          <div>
            <h3 className="text-lg font-bold text-slate-900 flex items-center gap-2">
              <FileText className="text-blue-600" size={20} />
              Báo cáo Chẩn đoán DGA Toàn diện (Comprehensive Diagnostic Report)
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Chuẩn đoán theo IEEE C57.104, IEC 60599, CIGRE TB 771 &amp; Phương pháp Duval
            </p>
          </div>
          <button onClick={onClose} className="w-8 h-8 rounded-full hover:bg-slate-200 text-slate-400 hover:text-slate-700 flex items-center justify-center cursor-pointer">
            <X size={18} />
          </button>
        </div>

        <div className="p-6 overflow-y-auto space-y-6" id="printable-overview-report">
          {/* Header section */}
          <div className="border-b border-slate-200 pb-4">
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2">
              <div>
                <span className="text-[11px] font-bold text-blue-600 uppercase tracking-wider">Hệ thống Giám sát &amp; Phân tích DGA Nâng cao</span>
                <h4 className="text-xl font-black text-slate-900 mt-0.5">{equipmentName}</h4>
                <p className="text-xs text-slate-500 font-mono mt-0.5">Mã thiết bị: <strong>{equipmentCode}</strong> | Đơn vị quản lý: <strong>{customerName}</strong></p>
              </div>
              <div className="text-left sm:text-right bg-slate-50 p-2.5 rounded-xl border border-slate-200">
                <span className="text-[11px] text-slate-500 block">Thời điểm lấy mẫu / thử nghiệm:</span>
                <span className="text-xs font-bold text-slate-800">{sampleDate}</span>
                <span className="text-[11px] text-slate-600 block mt-0.5">Kỹ sư thực hiện: <strong>{engineerName}</strong></span>
              </div>
            </div>
          </div>

          {/* PART 1: BẢNG TỔNG HỢP NỒNG ĐỘ VÀ TỐC ĐỘ TĂNG KHÍ */}
          <div className="space-y-3">
            <h5 className="text-xs font-black text-slate-800 uppercase tracking-wider flex items-center gap-2">
              <span className="w-5 h-5 rounded-full bg-blue-100 text-blue-700 flex items-center justify-center text-[10px]">1</span>
              BẢNG TỔNG HỢP NỒNG ĐỘ VÀ TỐC ĐỘ TĂNG KHÍ (TDCG &amp; GASSING RATE)
            </h5>

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5 text-xs">
              <div className="bg-slate-50 border border-slate-200 rounded-xl p-3">
                <span className="text-slate-500 block text-[11px]">Tổng khí cháy (TDCG):</span>
                <span className="text-lg font-black text-slate-900">{tdcg} ppm</span>
                <span className="text-[10px] text-slate-500 block mt-0.5">H₂+CH₄+C₂H₆+C₂H₄+C₂H₂+CO</span>
              </div>
              <div className={`border rounded-xl p-3 ${isCondition1 ? 'bg-emerald-50 border-emerald-200' : 'bg-amber-50 border-amber-200'}`}>
                <span className="text-slate-500 block text-[11px]">Phân loại Condition IEEE:</span>
                <span className={`text-lg font-black ${isCondition1 ? 'text-emerald-700' : 'text-amber-700'}`}>
                  {analysisResult?.conditionName || (parseFloat(tdcg) <= 720 ? 'Condition 1' : 'Condition 2/3')}
                </span>
                <span className="text-[10px] opacity-80 block mt-0.5">{isCondition1 ? 'Bình thường / Khỏe mạnh' : 'Cần giám sát / Cảnh báo'}</span>
              </div>
              <div className="bg-slate-50 border border-slate-200 rounded-xl p-3">
                <span className="text-slate-500 block text-[11px]">Tốc độ sinh khí (Gassing Rate):</span>
                <span className="text-sm font-bold text-slate-800">
                  {analysisResult?.gassingRate?.status === 'High' ? 'Cao (Warning/High)' : 'Bình thường (Normal)'}
                </span>
                <span className="text-[10px] text-slate-500 block mt-0.5">&lt; 10 ppm/tháng</span>
              </div>
              <div className="bg-slate-50 border border-slate-200 rounded-xl p-3">
                <span className="text-slate-500 block text-[11px]">Quyết định sơ bộ:</span>
                <span className={`text-sm font-bold ${isCondition1 ? 'text-emerald-700' : 'text-rose-700'}`}>
                  {isCondition1 ? 'NORMAL / HEALTHY' : 'FAULT DETECTED / ACTIVE'}
                </span>
                <span className="text-[10px] text-slate-500 block mt-0.5">{isCondition1 ? 'Bỏ qua phân tích lỗi' : 'Kích hoạt chẩn đoán lỗi'}</span>
              </div>
            </div>

            {/* Gas values table */}
            <div className="grid grid-cols-4 sm:grid-cols-7 gap-1.5 text-center text-xs">
              <div className="p-2 bg-slate-50 rounded-lg border border-slate-200">
                <span className="text-[10px] text-slate-500 block">H₂</span>
                <span className="font-bold text-slate-800">{h2}</span>
              </div>
              <div className="p-2 bg-slate-50 rounded-lg border border-slate-200">
                <span className="text-[10px] text-slate-500 block">CH₄</span>
                <span className="font-bold text-slate-800">{ch4}</span>
              </div>
              <div className="p-2 bg-slate-50 rounded-lg border border-slate-200">
                <span className="text-[10px] text-slate-500 block">C₂H₆</span>
                <span className="font-bold text-slate-800">{c2h6}</span>
              </div>
              <div className="p-2 bg-slate-50 rounded-lg border border-slate-200">
                <span className="text-[10px] text-slate-500 block">C₂H₄</span>
                <span className="font-bold text-slate-800">{c2h4}</span>
              </div>
              <div className="p-2 bg-slate-50 rounded-lg border border-slate-200">
                <span className="text-[10px] text-rose-600 font-bold block">C₂H₂</span>
                <span className="font-bold text-rose-600">{c2h2}</span>
              </div>
              <div className="p-2 bg-slate-50 rounded-lg border border-slate-200">
                <span className="text-[10px] text-slate-500 block">CO</span>
                <span className="font-bold text-slate-800">{co}</span>
              </div>
              <div className="p-2 bg-slate-50 rounded-lg border border-slate-200">
                <span className="text-[10px] text-slate-500 block">CO₂</span>
                <span className="font-bold text-slate-800">{co2}</span>
              </div>
            </div>
          </div>

          {/* PART 2: KẾT QUẢ CHẨN ĐOÁN CHI TIẾT TỪ CÁC PHƯƠNG PHÁP */}
          <div className="space-y-3">
            <h5 className="text-xs font-black text-slate-800 uppercase tracking-wider flex items-center gap-2">
              <span className="w-5 h-5 rounded-full bg-blue-100 text-blue-700 flex items-center justify-center text-[10px]">2</span>
              KẾT QUẢ CHẨN ĐOÁN CHI TIẾT TỪ CÁC PHƯƠNG PHÁP
            </h5>

            {isCondition1 ? (
              <div className="bg-emerald-50/70 border border-emerald-200 rounded-xl p-4 text-xs text-emerald-900 space-y-1">
                <div className="font-bold text-emerald-950 flex items-center gap-1.5">
                  <CheckCircle2 size={16} className="text-emerald-600" />
                  BỎ QUA CÁC PHƯƠNG PHÁP PHÂN TÍCH LỖI (Condition 1 Healthy Rule)
                </div>
                <p className="text-emerald-800 leading-relaxed">
                  Theo thuật toán chuẩn đoán IEEE C57.104, IEC 60599 và CIGRE TB 771: Do tổng lượng khí cháy TDCG ở mức <strong>Condition 1 (≤ 720 ppm)</strong> và toàn bộ nồng độ khí riêng lẻ đều dưới ngưỡng cảnh báo, hệ thống tự động <strong>bỏ qua các phương pháp ma trận và tam giác Duval</strong> (Duval Triangle T1/T4/T5, Duval Pentagon P1/P2, Rogers, IEC, Dornenburg). Quy tắc này ngăn ngừa phát sinh các chẩn đoán sai và báo động giả bắt nguồn từ tạp âm đo lường ở nồng độ cực thấp.
                </p>
              </div>
            ) : null}

            <div className="overflow-x-auto border border-slate-200 rounded-xl">
              <table className="w-full text-xs text-left">
                <thead className="bg-slate-100/80 text-slate-700 font-semibold border-b border-slate-200">
                  <tr>
                    <th className="py-2.5 px-3">Phương pháp chẩn đoán</th>
                    <th className="py-2.5 px-3">Tiêu chuẩn / Cơ sở</th>
                    <th className="py-2.5 px-3">Kết quả</th>
                    <th className="py-2.5 px-3">Ý nghĩa &amp; Đánh giá</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  <tr>
                    <td className="py-2 px-3 font-semibold text-slate-800">Tỷ số CO₂ / CO</td>
                    <td className="py-2 px-3 text-slate-500">IEEE C57.104</td>
                    <td className="py-2 px-3 font-mono font-bold text-blue-700">
                      {isCondition1 ? 'Bình thường' : (analysisResult?.detailedResults?.co2_co?.status || 'OK')}
                    </td>
                    <td className="py-2 px-3 text-slate-600">
                      {isCondition1 ? 'Không có thoái hóa giấy cách điện cấp tính' : (analysisResult?.detailedResults?.co2_co?.description || 'Bình thường')}
                    </td>
                  </tr>
                  <tr>
                    <td className="py-2 px-3 font-semibold text-slate-800">Phương pháp Khí chủ đạo (Key Gas)</td>
                    <td className="py-2 px-3 text-slate-500">IEEE C57.104</td>
                    <td className="py-2 px-3 font-mono font-bold text-blue-700">
                      {isCondition1 ? 'Normal' : (analysisResult?.detailedResults?.keyGas?.fault || 'Normal')}
                    </td>
                    <td className="py-2 px-3 text-slate-600">
                      {isCondition1 ? 'Không có khí chủ đạo gây sự cố' : (analysisResult?.detailedResults?.keyGas?.description || 'Bình thường')}
                    </td>
                  </tr>
                  <tr>
                    <td className="py-2 px-3 font-semibold text-slate-800">Tỷ số Rogers (IEEE PC57.104)</td>
                    <td className="py-2 px-3 text-slate-500">IEEE / IEC</td>
                    <td className="py-2 px-3 font-mono font-bold text-blue-700">
                      {isCondition1 ? 'Normal' : (analysisResult?.matrix?.['IEEE (Rogers)'] || 'Normal')}
                    </td>
                    <td className="py-2 px-3 text-slate-600">
                      {isCondition1 ? 'Tỷ số trong giới hạn bình thường' : (analysisResult?.detailedResults?.rogers?.description || 'Bình thường')}
                    </td>
                  </tr>
                  <tr>
                    <td className="py-2 px-3 font-semibold text-slate-800">Tỷ số IEC 60599</td>
                    <td className="py-2 px-3 text-slate-500">IEC 60599 Table 1</td>
                    <td className="py-2 px-3 font-mono font-bold text-blue-700">
                      {isCondition1 ? 'Normal' : (analysisResult?.matrix?.['IEC Ratio'] || 'Normal')}
                    </td>
                    <td className="py-2 px-3 text-slate-600">
                      {isCondition1 ? 'Bình thường (Condition 1)' : (analysisResult?.detailedResults?.iec60599?.description || 'Bình thường')}
                    </td>
                  </tr>
                  <tr>
                    <td className="py-2 px-3 font-semibold text-slate-800">Tỷ số Dornenburg</td>
                    <td className="py-2 px-3 text-slate-500">IEEE C57.104</td>
                    <td className="py-2 px-3 font-mono font-bold text-blue-700">
                      {isCondition1 ? 'Normal' : (analysisResult?.matrix?.['IEEE (Dornenburg)'] || 'Normal')}
                    </td>
                    <td className="py-2 px-3 text-slate-600">
                      {isCondition1 ? 'Chưa vượt ngưỡng L1 kích hoạt' : (analysisResult?.detailedResults?.dornenburg?.description || 'Bình thường')}
                    </td>
                  </tr>
                  <tr>
                    <td className="py-2 px-3 font-semibold text-slate-800">Tam giác Duval 1 (T1)</td>
                    <td className="py-2 px-3 text-slate-500">Duval Methodology</td>
                    <td className="py-2 px-3 font-mono font-bold text-blue-700">
                      {isCondition1 ? 'Normal' : (analysisResult?.matrix?.['Duval T1'] || 'Normal')}
                    </td>
                    <td className="py-2 px-3 text-slate-600">
                      {isCondition1 ? 'Bỏ qua do Condition 1 để tránh báo động giả' : (analysisResult?.detailedResults?.duvalT1?.faultName || 'Bình thường')}
                    </td>
                  </tr>
                  <tr>
                    <td className="py-2 px-3 font-semibold text-slate-800">Tam giác Duval 4 / 5 (Tinh chỉnh)</td>
                    <td className="py-2 px-3 text-slate-500">CIGRE TB 771</td>
                    <td className="py-2 px-3 font-mono font-bold text-blue-700">
                      {isCondition1 ? 'Normal' : `${analysisResult?.matrix?.['Duval T4'] || 'ND'} / ${analysisResult?.matrix?.['Duval T5'] || 'ND'}`}
                    </td>
                    <td className="py-2 px-3 text-slate-600">
                      {isCondition1 ? 'Không kích hoạt tinh chỉnh' : 'Tinh chỉnh quá nhiệt dầu / giấy'}
                    </td>
                  </tr>
                  <tr>
                    <td className="py-2 px-3 font-semibold text-slate-800">Ngũ giác Duval P1 &amp; P2</td>
                    <td className="py-2 px-3 text-slate-500">Duval Pentagons</td>
                    <td className="py-2 px-3 font-mono font-bold text-blue-700">
                      {isCondition1 ? 'Normal' : (analysisResult?.matrix?.['Duval P1'] || 'Normal')}
                    </td>
                    <td className="py-2 px-3 text-slate-600">
                      {isCondition1 ? 'Bình thường (Condition 1)' : (analysisResult?.detailedResults?.duvalP1?.faultName || 'Bình thường')}
                    </td>
                  </tr>
                  <tr>
                    <td className="py-2 px-3 font-semibold text-slate-800">Ma trận ETRA</td>
                    <td className="py-2 px-3 text-slate-500">ETRA Matrix</td>
                    <td className="py-2 px-3 font-mono font-bold text-blue-700">
                      {isCondition1 ? 'Normal' : (analysisResult?.matrix?.['ETRA'] || 'Normal')}
                    </td>
                    <td className="py-2 px-3 text-slate-600">
                      {isCondition1 ? 'Điều kiện vận hành bình thường' : 'Theo dõi theo cấp độ TDCG'}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          {/* PART 3: KẾT LUẬN CUỐI CÙNG (DGA MATRIX SYNTHESIS) */}
          <div className="space-y-3">
            <h5 className="text-xs font-black text-slate-800 uppercase tracking-wider flex items-center gap-2">
              <span className="w-5 h-5 rounded-full bg-blue-100 text-blue-700 flex items-center justify-center text-[10px]">3</span>
              KẾT LUẬN CUỐI CÙNG (DGA MATRIX SYNTHESIS)
            </h5>

            <div className="bg-slate-50 rounded-xl p-4 border border-slate-200 space-y-2 text-xs">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-1 border-b border-slate-200 pb-2">
                <span className="text-slate-500">Dạng sự cố chủ đạo được xác định:</span>
                <span className="font-extrabold text-blue-800 text-sm">
                  {isCondition1 ? 'NORMAL / HEALTHY (Không có sự cố)' : (analysisResult?.synthesis?.primaryFaultVi || 'Bình thường')}
                </span>
              </div>
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-1 border-b border-slate-200 pb-2">
                <span className="text-slate-500">Độ tin cậy kết luận:</span>
                <span className="font-bold text-slate-800">
                  {isCondition1 ? 'Cao (Khẳng định theo IEEE C57.104 Condition 1)' : (analysisResult?.synthesis?.confidence || 'Cao')}
                </span>
              </div>
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-1">
                <span className="text-slate-500">Đánh giá ảnh hưởng đến cách điện giấy:</span>
                <span className="font-semibold text-slate-800">
                  {isCondition1 ? 'Cách điện giấy xenlulo tốt, không phát hiện thoái hóa nhiệt.' : (analysisResult?.synthesis?.paperInvolvement || 'Bình thường')}
                </span>
              </div>
            </div>
          </div>

          {/* PART 4: KHUYẾN CÁO VẬN HÀNH BẢO TRÌ */}
          <div className="space-y-3">
            <h5 className="text-xs font-black text-slate-800 uppercase tracking-wider flex items-center gap-2">
              <span className="w-5 h-5 rounded-full bg-blue-100 text-blue-700 flex items-center justify-center text-[10px]">4</span>
              KHUYẾN CÁO VẬN HÀNH BẢO TRÌ
            </h5>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
              <div className="p-3.5 bg-blue-50/60 rounded-xl border border-blue-100 space-y-1.5">
                <span className="text-[11px] font-bold text-blue-900 uppercase">Chu kỳ lấy mẫu thí nghiệm tiếp theo:</span>
                <div className="text-base font-black text-blue-900">
                  {isCondition1 ? '12 tháng (Lấy mẫu hàng năm)' : (analysisResult?.nextSamplingInterval || '1 - 3 tháng')}
                </div>
                <p className="text-slate-600 text-[11px]">
                  {isCondition1 ? 'Duy trì chế độ giám sát định kỳ hàng năm theo quy trình vận hành trạm.' : 'Rút ngắn chu kỳ lấy mẫu để theo dõi sát diễn biến khí sinh.'}
                </p>
              </div>

              <div className="p-3.5 bg-slate-50 rounded-xl border border-slate-200 space-y-1.5">
                <span className="text-[11px] font-bold text-slate-700 uppercase">Biện pháp vận hành và thí nghiệm điện:</span>
                <ul className="space-y-1 text-slate-600 text-[11px]">
                  {isCondition1 ? (
                    <>
                      <li className="flex items-center gap-1.5"><Check size={12} className="text-emerald-600" /> Tiếp tục vận hành máy biến áp ở chế độ tải định mức.</li>
                      <li className="flex items-center gap-1.5"><Check size={12} className="text-emerald-600" /> Theo dõi nhiệt độ cuộn dây và dầu đỉnh qua hệ thống Scada.</li>
                    </>
                  ) : (
                    <>
                      <li className="flex items-center gap-1.5"><Check size={12} className="text-amber-600" /> Đo phóng điện cục bộ (PD) tại các đầu sứ bushing và thùng MBA.</li>
                      <li className="flex items-center gap-1.5"><Check size={12} className="text-amber-600" /> Chụp ảnh nhiệt hồng ngoại các mối nối thanh cái và cánh tản nhiệt.</li>
                    </>
                  )}
                </ul>
              </div>
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="px-6 py-3 border-t border-slate-200 bg-slate-50 flex items-center justify-between">
          <button
            type="button"
            onClick={() => window.print()}
            className="px-3 py-1.5 bg-white border border-slate-200 hover:bg-slate-100 text-slate-700 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-colors shadow-2xs cursor-pointer"
          >
            <Printer size={14} />
            In trực tiếp
          </button>
          <div className="flex items-center gap-2">
            {onDownloadPdf && (
              <button
                type="button"
                onClick={onDownloadPdf}
                className="px-4 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-colors shadow-xs cursor-pointer"
              >
                <Download size={14} />
                Tải file PDF
              </button>
            )}
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg text-xs font-semibold transition-colors cursor-pointer"
            >
              Đóng
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

// 7. EVENT LOG MODAL
export const EventLogModal: React.FC<{
  isOpen: boolean;
  onClose: () => void;
  equipmentCode?: string;
  equipmentName?: string;
}> = ({
  isOpen,
  onClose,
  equipmentCode = 'TR-01',
  equipmentName = 'MBA T1 110kV Bắc Ninh'
}) => {
  const [filter, setFilter] = useState<'all' | 'alarm' | 'warning' | 'info'>('all');

  const events = [
    { id: 1, type: 'info', time: '14/05/2025 08:15', title: 'Hoàn thành lấy mẫu DGA định kỳ', desc: 'Chỉ số khí hòa tan cập nhật vào hệ thống giám sát.', tag: 'Hệ thống' },
    { id: 2, type: 'warning', time: '13/05/2025 18:05', title: 'Cảnh báo nồng độ khí CO tăng', desc: 'Nồng độ CO đạt 290 ppm, tăng nhẹ so với tuần trước. Cần theo dõi thêm.', tag: 'DGA Warning' },
    { id: 3, type: 'info', time: '11/05/2025 10:20', title: 'Chuyển đổi phụ tải trạm (Load transfer)', desc: 'Tải tăng từ 45% lên 78% công suất định mức trong 4 giờ.', tag: 'Vận hành' },
    { id: 4, type: 'info', time: '09/05/2025 14:15', title: 'Bảo trì hệ thống DGA online', desc: 'Hiệu chuẩn cảm biến hồng ngoại đo khí CH₄ và photoacoustic H₂.', tag: 'Bảo trì' },
    { id: 5, type: 'alarm', time: '02/05/2025 09:00', title: 'Cảnh báo vi lượng khí C₂H₂ (0.8 ppm)', desc: 'Xuất hiện khí Acetylen ở nồng độ cận ngưỡng hồ quang điện nhẹ.', tag: 'DGA Alarm' },
    { id: 6, type: 'info', time: '28/04/2025 11:30', title: 'Lấy mẫu dầu thí nghiệm độc lập ASTM D3612', desc: 'Kết quả phòng thí nghiệm xác nhận tương đồng với hệ thống online.', tag: 'Thí nghiệm' }
  ];

  const filteredEvents = events.filter(e => filter === 'all' || e.type === filter);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xs p-4 overflow-y-auto">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-2xl max-h-[90vh] flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/80">
          <div>
            <h3 className="text-lg font-bold text-slate-900 flex items-center gap-2">
              <Activity className="text-blue-600" size={20} />
              Nhật ký Sự kiện &amp; Cảnh báo DGA (Active Events Log)
            </h3>
            <p className="text-xs text-slate-500 mt-0.5">
              Thiết bị: <span className="font-semibold text-slate-700">{equipmentCode} — {equipmentName}</span>
            </p>
          </div>
          <button onClick={onClose} className="w-8 h-8 rounded-full hover:bg-slate-200 text-slate-400 hover:text-slate-700 flex items-center justify-center cursor-pointer">
            <X size={18} />
          </button>
        </div>

        {/* Filter Bar */}
        <div className="px-6 py-2.5 bg-slate-100/60 border-b border-slate-200 flex items-center gap-2">
          {(['all', 'alarm', 'warning', 'info'] as const).map(tab => (
            <button
              key={tab}
              type="button"
              onClick={() => setFilter(tab)}
              className={`px-3 py-1 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                filter === tab ? 'bg-white text-blue-600 shadow-2xs' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              {tab === 'all' ? 'Tất cả' : tab === 'alarm' ? 'Báo động (Alarm)' : tab === 'warning' ? 'Cảnh báo (Warning)' : 'Thông tin (Info)'}
            </button>
          ))}
        </div>

        <div className="p-6 overflow-y-auto space-y-3">
          {filteredEvents.map(evt => (
            <div key={evt.id} className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/50 hover:bg-slate-50 transition-colors flex items-start gap-3">
              <div className={`w-8 h-8 rounded-full flex items-center justify-center flex-shrink-0 mt-0.5 ${
                evt.type === 'alarm' ? 'bg-rose-100 text-rose-600' : evt.type === 'warning' ? 'bg-amber-100 text-amber-600' : 'bg-blue-100 text-blue-600'
              }`}>
                {evt.type === 'alarm' ? <Flame size={16} /> : evt.type === 'warning' ? <AlertTriangle size={16} /> : <Info size={16} />}
              </div>
              <div className="flex-1 min-w-0">
                <div className="flex items-center justify-between gap-2">
                  <h5 className="text-xs font-bold text-slate-900">{evt.title}</h5>
                  <span className="text-[10px] text-slate-400 font-mono whitespace-nowrap">{evt.time}</span>
                </div>
                <p className="text-xs text-slate-600 mt-1">{evt.desc}</p>
                <span className="inline-block mt-2 text-[10px] font-semibold px-2 py-0.5 rounded-md bg-slate-200 text-slate-700">
                  {evt.tag}
                </span>
              </div>
            </div>
          ))}
        </div>

        <div className="px-6 py-3 border-t border-slate-200 bg-slate-50 flex items-center justify-end">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-1.5 bg-slate-800 hover:bg-slate-900 text-white rounded-lg text-xs font-semibold transition-colors cursor-pointer"
          >
            Đóng
          </button>
        </div>
      </div>
    </div>
  );
};
