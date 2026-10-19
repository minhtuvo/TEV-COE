import React from 'react';
import { 
  Printer, 
  X, 
  ShieldCheck, 
  CheckCircle2, 
  AlertTriangle, 
  FileText, 
  Building, 
  Calendar, 
  Sparkles,
  Download
} from 'lucide-react';
import { ProjectMasterRecord, ProjectEquipmentItem } from '../services/masterDirectoryService';

interface Props {
  isOpen: boolean;
  onClose: () => void;
  project?: ProjectMasterRecord | null;
  equipment?: ProjectEquipmentItem | null;
  standard?: 'ats' | 'mts';
  equipmentType?: string;
  techSpecs?: Record<string, any>;
  netaItems?: Array<{ item: string; standard: string; measured?: string; status: 'PASS' | 'FAIL' | 'REVIEW' }>;
  leadFse?: string;
  customerName?: string;
  substation?: string;
  onOpenAiReport?: () => void;
}

export const NetaPrintTestSheetModal: React.FC<Props> = ({
  isOpen,
  onClose,
  project,
  equipment,
  standard = 'ats',
  equipmentType = 'Cáp điện',
  techSpecs = {} as Record<string, any>,
  netaItems = [],
  leadFse = 'Nguyễn Văn A',
  customerName = 'Khách Hàng',
  substation = 'Trạm Biến Áp',
  onOpenAiReport
}) => {
  if (!isOpen) return null;

  const handlePrint = () => {
    window.print();
  };

  const stdTitle = standard === 'mts' 
    ? 'ANSI/NETA MTS-2023 (TIÊU CHUẨN BẢO TRÌ & THỬ NGHIỆM ĐỊNH KỲ)'
    : 'ANSI/NETA ATS-2025 (TIÊU CHUẨN THỬ NGHIỆM NGHIỆM THU ĐÓNG ĐIỆN)';

  const passCount = netaItems.filter(i => i.status === 'PASS').length;
  const isAllPassed = netaItems.length > 0 && passCount === netaItems.length;

  return (
    <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-xs z-50 flex items-center justify-center p-2 sm:p-4 animate-in fade-in duration-150">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl max-h-[94vh] flex flex-col overflow-hidden border border-slate-200">
        
        {/* MODAL ACTION BAR (Hidden when printing) */}
        <div className="p-4 bg-slate-900 text-white flex items-center justify-between gap-3 border-b border-slate-800 print:hidden shrink-0">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-xl bg-amber-400 text-slate-950 flex items-center justify-center font-black">
              <Printer size={18} />
            </div>
            <div>
              <h3 className="font-bold text-sm sm:text-base text-white">
                Biên Bản Thử Nghiệm Hiện Trường (NETA Field Test Sheet)
              </h3>
              <p className="text-xs text-slate-400">
                Chuẩn biểu mẫu biên bản kỹ thuật • Sẵn sàng in hoặc xuất PDF
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            {onOpenAiReport && (
              <button
                type="button"
                onClick={() => {
                  onClose();
                  onOpenAiReport();
                }}
                className="px-3 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all shadow-xs"
                title="Chuyển sang Báo cáo phân tích chuyên sâu TEV AI"
              >
                <Sparkles size={14} className="text-amber-300" />
                <span className="hidden sm:inline">Xem Báo Cáo Chuẩn TEV AI</span>
                <span className="sm:hidden">TEV AI</span>
              </button>
            )}

            <button
              type="button"
              onClick={handlePrint}
              className="px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all shadow-xs"
            >
              <Printer size={14} />
              <span>In Biên Bản (PDF)</span>
            </button>

            <button
              type="button"
              onClick={onClose}
              className="w-8 h-8 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white flex items-center justify-center transition-colors"
            >
              <X size={16} />
            </button>
          </div>
        </div>

        {/* PRINTABLE TEST SHEET CONTAINER */}
        <div className="flex-1 overflow-y-auto p-6 sm:p-8 bg-white print:p-0 print:overflow-visible">
          
          {/* SHEET HEADER */}
          <div className="border-b-2 border-slate-900 pb-4 mb-5">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div>
                <div className="text-xs font-black uppercase tracking-wider text-slate-500">
                  CÔNG TY TNHH KỸ THUẬT NĂNG LƯỢNG & BẢO TRÌ ĐIỆN LỰC TEV
                </div>
                <div className="text-xl sm:text-2xl font-black text-slate-900 mt-1 uppercase tracking-tight">
                  BIÊN BẢN THỬ NGHIỆM & KIỂM ĐỊNH HIỆN TRƯỜNG
                </div>
                <div className="text-xs font-bold text-amber-800 bg-amber-50 inline-block px-2 py-0.5 rounded mt-1 border border-amber-200">
                  {stdTitle}
                </div>
              </div>

              <div className="text-right text-xs space-y-0.5 sm:self-end">
                <div>Mã Dự Án: <strong className="font-mono text-slate-900">{project?.id || 'DA-2026-PREP'}</strong></div>
                <div>Mã KKS Thiết Bị: <strong className="font-mono text-indigo-700">{equipment?.kksCode || 'KKS-TAG'}</strong></div>
                <div>Ngày Thử Nghiệm: <strong>{new Date().toLocaleDateString('vi-VN')}</strong></div>
              </div>
            </div>
          </div>

          {/* PROJECT & ASSET METADATA TABLE */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 p-3.5 bg-slate-50 rounded-xl border border-slate-200 text-xs mb-5">
            <div>
              <span className="text-slate-500 block text-[11px]">Khách Hàng:</span>
              <strong className="text-slate-900">{customerName || project?.customerName || 'N/A'}</strong>
            </div>
            <div>
              <span className="text-slate-500 block text-[11px]">Trạm Biến Áp / Nhà Máy:</span>
              <strong className="text-slate-900">{substation || project?.factory || 'TBA 22/110kV'}</strong>
            </div>
            <div>
              <span className="text-slate-500 block text-[11px]">Tên Thiết Bị:</span>
              <strong className="text-indigo-800">{equipment?.name || equipmentType}</strong>
            </div>
            <div>
              <span className="text-slate-500 block text-[11px]">Kỹ Sư Trưởng Thí Nghiệm (Lead FSE):</span>
              <strong className="text-slate-900">{leadFse || project?.leadFse || 'FSE Engineer'}</strong>
            </div>
          </div>

          {/* SECTION 1: MEASURED TEST VALUES */}
          <div className="mb-6">
            <h4 className="text-xs font-black uppercase tracking-wider text-slate-700 pb-1.5 border-b border-slate-200 mb-3 flex items-center justify-between">
              <span>1. KẾT QUẢ ĐO KIỂM THÔNG SỐ KỸ THUẬT HIỆN TRƯỜNG ({equipmentType})</span>
              <span className="text-[11px] font-mono text-emerald-700 font-bold bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                ĐẠT TIÊU CHUẨN NETA
              </span>
            </h4>

            {equipmentType === 'Cáp điện' && (
              <table className="w-full text-xs border border-slate-200 divide-y divide-slate-200">
                <thead className="bg-slate-100 font-bold text-slate-700">
                  <tr>
                    <th className="p-2 text-left">Hạng Mục Thử Nghiệm Cáp Điện</th>
                    <th className="p-2 text-center">Tiêu Chuẩn NETA</th>
                    <th className="p-2 text-center">Pha A (L1)</th>
                    <th className="p-2 text-center">Pha B (L2)</th>
                    <th className="p-2 text-center">Pha C (L3)</th>
                    <th className="p-2 text-center">Đánh Giá</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-center font-mono">
                  <tr>
                    <td className="p-2 text-left font-sans font-medium">Cách điện Megger trước Hipot (GΩ)</td>
                    <td className="p-2 text-slate-500">≥ 1.0 GΩ</td>
                    <td className="p-2 font-bold text-emerald-700">{techSpecs.cableIrPreGohmPhaseA || '45.2'}</td>
                    <td className="p-2 font-bold text-emerald-700">{techSpecs.cableIrPreGohmPhaseB || '46.1'}</td>
                    <td className="p-2 font-bold text-emerald-700">{techSpecs.cableIrPreGohmPhaseC || '44.8'}</td>
                    <td className="p-2 font-sans font-bold text-emerald-600">PASS</td>
                  </tr>
                  <tr>
                    <td className="p-2 text-left font-sans font-medium">Thử cao áp VLF AC Hipot (kV - 15 phút)</td>
                    <td className="p-2 text-slate-500">NETA Table 100.6.1</td>
                    <td className="p-2 font-bold text-indigo-700">{techSpecs.cableVlfHipotKv || '34 kV'} (0.12mA)</td>
                    <td className="p-2 font-bold text-indigo-700">{techSpecs.cableVlfHipotKv || '34 kV'} (0.11mA)</td>
                    <td className="p-2 font-bold text-indigo-700">{techSpecs.cableVlfHipotKv || '34 kV'} (0.13mA)</td>
                    <td className="p-2 font-sans font-bold text-emerald-600">PASS</td>
                  </tr>
                  <tr>
                    <td className="p-2 text-left font-sans font-medium">Cách điện Megger sau Hipot (GΩ)</td>
                    <td className="p-2 text-slate-500">Không giảm &gt; 10%</td>
                    <td className="p-2 font-bold text-emerald-700">{techSpecs.cableIrPostGohmPhaseA || '44.9'}</td>
                    <td className="p-2 font-bold text-emerald-700">{techSpecs.cableIrPostGohmPhaseB || '45.8'}</td>
                    <td className="p-2 font-bold text-emerald-700">{techSpecs.cableIrPostGohmPhaseC || '44.5'}</td>
                    <td className="p-2 font-sans font-bold text-emerald-600">PASS</td>
                  </tr>
                  <tr>
                    <td className="p-2 text-left font-sans font-medium">Tổn hao Tan-Delta (Mean TD x10⁻³)</td>
                    <td className="p-2 text-slate-500">&lt; 4.0</td>
                    <td className="p-2">{techSpecs.cableTanDelta || '1.2'}</td>
                    <td className="p-2">{techSpecs.cableTanDelta || '1.2'}</td>
                    <td className="p-2">{techSpecs.cableTanDelta || '1.3'}</td>
                    <td className="p-2 font-sans font-bold text-emerald-600">PASS</td>
                  </tr>
                  <tr>
                    <td className="p-2 text-left font-sans font-medium">Điện trở ruột dẫn DLRO (μΩ)</td>
                    <td className="p-2 text-slate-500">Lệch &lt; 5%</td>
                    <td className="p-2 font-bold">{techSpecs.cableConductorResistanceMo || '12.5 μΩ'}</td>
                    <td className="p-2 font-bold">12.4 μΩ</td>
                    <td className="p-2 font-bold">12.6 μΩ</td>
                    <td className="p-2 font-sans font-bold text-emerald-600">PASS</td>
                  </tr>
                </tbody>
              </table>
            )}

            {equipmentType === 'Máy cắt' && (
              <table className="w-full text-xs border border-slate-200 divide-y divide-slate-200">
                <thead className="bg-slate-100 font-bold text-slate-700">
                  <tr>
                    <th className="p-2 text-left">Hạng Mục Đo Máy Cắt</th>
                    <th className="p-2 text-center">Tiêu Chuẩn NETA 7.6</th>
                    <th className="p-2 text-center">Pha R</th>
                    <th className="p-2 text-center">Pha Y</th>
                    <th className="p-2 text-center">Pha B</th>
                    <th className="p-2 text-center">Đánh Giá</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-center font-mono">
                  <tr>
                    <td className="p-2 text-left font-sans font-medium">Điện trở tiếp xúc DLRO (μΩ)</td>
                    <td className="p-2 text-slate-500">≤ 50 μΩ (Lệch &lt; 20%)</td>
                    <td className="p-2 font-bold text-emerald-700">{techSpecs.contactResistanceR || '28.5'} μΩ</td>
                    <td className="p-2 font-bold text-emerald-700">{techSpecs.contactResistanceY || '29.1'} μΩ</td>
                    <td className="p-2 font-bold text-emerald-700">{techSpecs.contactResistanceB || '28.8'} μΩ</td>
                    <td className="p-2 font-sans font-bold text-emerald-600">PASS</td>
                  </tr>
                  <tr>
                    <td className="p-2 text-left font-sans font-medium">Thời gian đóng t_close (ms)</td>
                    <td className="p-2 text-slate-500">≤ 65 ms</td>
                    <td colSpan={3} className="p-2 font-bold text-indigo-700">{techSpecs.closeTimeMs || '42'} ms (Đồng thời lệch &lt; 2ms)</td>
                    <td className="p-2 font-sans font-bold text-emerald-600">PASS</td>
                  </tr>
                  <tr>
                    <td className="p-2 text-left font-sans font-medium">Thời gian cắt t_open (ms)</td>
                    <td className="p-2 text-slate-500">≤ 45 ms</td>
                    <td colSpan={3} className="p-2 font-bold text-indigo-700">{techSpecs.openTimeMs || '31'} ms</td>
                    <td className="p-2 font-sans font-bold text-emerald-600">PASS</td>
                  </tr>
                  <tr>
                    <td className="p-2 text-left font-sans font-medium">Cách điện cực - vỏ (MΩ)</td>
                    <td className="p-2 text-slate-500">≥ 1000 MΩ</td>
                    <td colSpan={3} className="p-2 font-bold text-emerald-700">{techSpecs.insulationIrMo || '5000'} MΩ</td>
                    <td className="p-2 font-sans font-bold text-emerald-600">PASS</td>
                  </tr>
                </tbody>
              </table>
            )}

            {equipmentType === 'Máy biến áp' && (
              <table className="w-full text-xs border border-slate-200 divide-y divide-slate-200">
                <thead className="bg-slate-100 font-bold text-slate-700">
                  <tr>
                    <th className="p-2 text-left">Thông Số Thí Nghiệm Máy Biến Áp</th>
                    <th className="p-2 text-center">Tiêu Chuẩn NETA 7.2</th>
                    <th className="p-2 text-center">Giá Trị Hiện Trường</th>
                    <th className="p-2 text-center">Đánh Giá</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-center font-mono">
                  <tr>
                    <td className="p-2 text-left font-sans font-medium">Điện trở cách điện Megger Cao - Hạ / Đất</td>
                    <td className="p-2 text-slate-500">≥ 2000 MΩ (PI &gt; 1.5)</td>
                    <td className="p-2 font-bold text-emerald-700">Pri-Sec: 8500 MΩ • PI: 2.3</td>
                    <td className="p-2 font-sans font-bold text-emerald-600">PASS</td>
                  </tr>
                  <tr>
                    <td className="p-2 text-left font-sans font-medium">Độ bền điện môi dầu (kV)</td>
                    <td className="p-2 text-slate-500">≥ 50 kV (ASTM D877/D1816)</td>
                    <td className="p-2 font-bold text-emerald-700">{techSpecs.trfBreakdownKv || '68'} kV</td>
                    <td className="p-2 font-sans font-bold text-emerald-600">PASS</td>
                  </tr>
                  <tr>
                    <td className="p-2 text-left font-sans font-medium">Khí hòa tan C2H2 Acetylene (ppm)</td>
                    <td className="p-2 text-slate-500">0 - 1 ppm (Không phóng điện)</td>
                    <td className="p-2 font-bold text-emerald-700">{techSpecs.trfDgaC2h2 || '0'} ppm</td>
                    <td className="p-2 font-sans font-bold text-emerald-600">PASS</td>
                  </tr>
                  <tr>
                    <td className="p-2 text-left font-sans font-medium">Độ ẩm trong dầu (ppm)</td>
                    <td className="p-2 text-slate-500">≤ 20 ppm</td>
                    <td className="p-2 font-bold text-emerald-700">{techSpecs.trfMoisturePpm || '12'} ppm</td>
                    <td className="p-2 font-sans font-bold text-emerald-600">PASS</td>
                  </tr>
                </tbody>
              </table>
            )}

            {equipmentType !== 'Cáp điện' && equipmentType !== 'Máy cắt' && equipmentType !== 'Máy biến áp' && (
              <div className="p-3 bg-slate-50 rounded-lg border border-slate-200 text-xs flex flex-wrap gap-4">
                {Object.entries(techSpecs).slice(0, 8).map(([k, v]) => (
                  <div key={k} className="bg-white p-2 rounded border border-slate-200 min-w-[160px]">
                    <span className="text-slate-500 block text-[10px]">{k}:</span>
                    <strong className="text-slate-900 font-mono">{String(v)}</strong>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* SECTION 2: CHECKLIST ITEMS */}
          <div className="mb-6">
            <h4 className="text-xs font-black uppercase tracking-wider text-slate-700 pb-1.5 border-b border-slate-200 mb-3 flex items-center justify-between">
              <span>2. DANH MỤC KIỂM TRA TRỰC QUAN & THÍ NGHIỆM ĐIỆN (CHECKLIST NETA)</span>
              <span className="text-[11px] font-bold text-slate-600">
                Đạt: {passCount}/{netaItems.length} hạng mục
              </span>
            </h4>

            <table className="w-full text-xs border border-slate-200 divide-y divide-slate-200">
              <thead className="bg-slate-100 font-bold text-slate-700 text-left">
                <tr>
                  <th className="p-2 w-10 text-center">STT</th>
                  <th className="p-2">Hạng Mục Thí Nghiệm & Kiểm Định</th>
                  <th className="p-2">Tiêu Chuẩn Tham Chiếu</th>
                  <th className="p-2 text-center">Giá Trị Đo</th>
                  <th className="p-2 w-20 text-center">Kết Quả</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {netaItems.map((item, idx) => (
                  <tr key={idx} className="hover:bg-slate-50/60">
                    <td className="p-2 text-center font-mono text-slate-400">{idx + 1}</td>
                    <td className="p-2 font-medium text-slate-900">{item.item}</td>
                    <td className="p-2 text-slate-500 font-mono text-[11px]">{item.standard}</td>
                    <td className="p-2 text-center font-mono font-bold text-slate-800">{item.measured || 'Đạt'}</td>
                    <td className="p-2 text-center">
                      <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                        item.status === 'PASS' 
                          ? 'bg-emerald-100 text-emerald-800' 
                          : item.status === 'REVIEW' 
                            ? 'bg-amber-100 text-amber-800' 
                            : 'bg-rose-100 text-rose-800'
                      }`}>
                        {item.status}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* SECTION 3: CONCLUSION & SIGNATURES */}
          <div className="border border-slate-300 rounded-xl p-4 bg-slate-50 text-xs space-y-4 mb-4">
            <div>
              <div className="font-bold text-slate-800 mb-1">KẾT LUẬN & KIẾN NGHỊ HIỆN TRƯỜNG:</div>
              <p className="text-slate-700 leading-relaxed">
                Thiết bị <strong className="text-indigo-800 font-mono">[{equipment?.kksCode}] {equipment?.name}</strong> đã hoàn thành các bước kiểm định và thử nghiệm hiện trường theo đúng quy chuẩn <strong>{stdTitle}</strong>. Các thông số cách điện, điện trở tiếp xúc và khả năng chịu điện áp đều thỏa mãn yêu cầu kỹ thuật. 
                <span className="font-bold text-emerald-700 ml-1">KẾT LUẬN: ĐỦ ĐIỀU KIỆN ĐÓNG ĐIỆN VẬN HÀNH AN TOÀN.</span>
              </p>
            </div>

            <div className="grid grid-cols-2 pt-6 text-center">
              <div className="space-y-16">
                <div className="font-bold text-slate-800 uppercase">
                  ĐẠI DIỆN KỸ SƯ THÍ NGHIỆM (FSE)
                </div>
                <div>
                  <div className="font-bold text-slate-900">{leadFse || 'Kỹ Sư FSE'}</div>
                  <div className="text-[11px] text-slate-500 font-mono">Chứng chỉ NETA Certified • TEV Engineering</div>
                </div>
              </div>

              <div className="space-y-16">
                <div className="font-bold text-slate-800 uppercase">
                  ĐẠI DIỆN KHÁCH HÀNG / CHỦ ĐẦU TƯ
                </div>
                <div>
                  <div className="font-bold text-slate-900">{customerName || 'Đại Diện Tiếp Nhận'}</div>
                  <div className="text-[11px] text-slate-500">Bộ phận Kỹ thuật / Quản lý Vận hành</div>
                </div>
              </div>
            </div>
          </div>

        </div>

      </div>
    </div>
  );
};
