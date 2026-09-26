import React from 'react';
import { Printer, Download, X, CheckCircle2, AlertTriangle, XCircle, ShieldCheck } from 'lucide-react';

interface Props {
  isOpen: boolean;
  onClose: () => void;
  title: string;
  standard: string;
  siteInfo: {
    project_name: string;
    transformer_tag: string;
    fse_name: string;
    test_date: string;
  };
  overallStatus: 'PASS' | 'INVESTIGATE' | 'FAIL';
  statusReason: string;
  visualChecks: Array<{ label: string; value: string; pass: boolean; note?: string }>;
  electricalTable: Array<{
    item: string;
    measured: string;
    standard: string;
    reference: string;
    deviation: string;
    status: 'PASS' | 'INVESTIGATE' | 'FAIL';
  }>;
  warnings: string[];
  recommendations: string[];
}

export const TevReportPrintModal: React.FC<Props> = ({
  isOpen,
  onClose,
  title,
  standard,
  siteInfo,
  overallStatus,
  statusReason,
  visualChecks,
  electricalTable,
  warnings,
  recommendations,
}) => {
  if (!isOpen) return null;

  const handlePrint = () => {
    window.print();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-2 sm:p-4 bg-slate-900/70 backdrop-blur-xs overflow-y-auto">
      <div className="bg-white w-full max-w-4xl rounded-2xl shadow-2xl border border-slate-200 overflow-hidden flex flex-col my-auto max-h-[92vh]">
        {/* Action bar (Hidden when printing) */}
        <div className="p-4 bg-slate-900 text-white flex items-center justify-between print:hidden">
          <div className="flex items-center gap-2">
            <ShieldCheck size={20} className="text-blue-400" />
            <h3 className="font-bold text-sm">Xem Trước & In Báo Cáo Kỹ Thuật (TEV Standard Format)</h3>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={handlePrint}
              className="px-3.5 py-1.5 bg-blue-600 hover:bg-blue-500 text-white text-xs font-semibold rounded-lg flex items-center gap-1.5 shadow-xs transition-colors"
            >
              <Printer size={15} />
              <span>In / Lưu PDF</span>
            </button>
            <button
              onClick={onClose}
              className="p-1.5 text-slate-400 hover:text-white hover:bg-slate-800 rounded-lg transition-colors"
            >
              <X size={18} />
            </button>
          </div>
        </div>

        {/* Printable Document Body */}
        <div id="tev-printable-report" className="p-6 sm:p-8 overflow-y-auto flex-1 text-slate-800 bg-white font-sans">
          {/* Header */}
          <div className="border-b-2 border-slate-900 pb-4 mb-6 flex flex-col sm:flex-row sm:items-end justify-between gap-4">
            <div>
              <div className="flex items-center gap-2">
                <span className="text-2xl font-black tracking-tight text-blue-900">TEV</span>
                <span className="text-xs uppercase font-bold tracking-widest text-slate-500 border-l border-slate-300 pl-2">
                  Engineering Field Service
                </span>
              </div>
              <h1 className="text-xl sm:text-2xl font-black text-slate-900 mt-2 uppercase">
                {title}
              </h1>
              <p className="text-xs font-semibold text-blue-700 tracking-wide">
                TIÊU CHUẨN NGHIỆM THU: {standard}
              </p>
            </div>

            <div className="text-left sm:text-right text-xs space-y-1">
              <div><span className="text-slate-500">Mã biên bản:</span> <strong className="font-mono">TEV-FSR-{siteInfo.transformer_tag}-{siteInfo.test_date.replace(/-/g, '')}</strong></div>
              <div><span className="text-slate-500">Ngày lập:</span> <strong>{siteInfo.test_date}</strong></div>
              <div><span className="text-slate-500">Kỹ sư FSE:</span> <strong>{siteInfo.fse_name}</strong></div>
            </div>
          </div>

          {/* Project & Equipment Info */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 bg-slate-50 p-4 rounded-xl border border-slate-200 text-xs mb-6">
            <div>
              <div className="text-slate-500">Dự án / Nhà máy:</div>
              <div className="font-bold text-slate-900">{siteInfo.project_name}</div>
            </div>
            <div>
              <div className="text-slate-500">Mã Tag thiết bị:</div>
              <div className="font-bold text-blue-900 font-mono">{siteInfo.transformer_tag}</div>
            </div>
            <div>
              <div className="text-slate-500">Tiêu chuẩn:</div>
              <div className="font-semibold text-slate-900">{standard}</div>
            </div>
            <div>
              <div className="text-slate-500">Trạng thái tổng quan:</div>
              <div>
                <span className={`inline-block px-2.5 py-0.5 rounded-full text-[11px] font-black uppercase ${
                  overallStatus === 'PASS'
                    ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                    : overallStatus === 'INVESTIGATE'
                    ? 'bg-amber-100 text-amber-800 border border-amber-300'
                    : 'bg-rose-100 text-rose-800 border border-rose-300'
                }`}>
                  {overallStatus}
                </span>
              </div>
            </div>
          </div>

          {/* Overall Status Banner */}
          <div className={`p-4 rounded-xl border mb-6 text-xs ${
            overallStatus === 'PASS'
              ? 'bg-emerald-50 border-emerald-300 text-emerald-950'
              : overallStatus === 'INVESTIGATE'
              ? 'bg-amber-50 border-amber-300 text-amber-950'
              : 'bg-rose-50 border-rose-300 text-rose-950'
          }`}>
            <div className="font-bold text-sm mb-1 flex items-center gap-1.5">
              {overallStatus === 'PASS' && <CheckCircle2 size={16} className="text-emerald-600" />}
              {overallStatus === 'INVESTIGATE' && <AlertTriangle size={16} className="text-amber-600" />}
              {overallStatus === 'FAIL' && <XCircle size={16} className="text-rose-600" />}
              <span>KẾT LUẬN NGHIỆM THU: {overallStatus === 'PASS' ? 'ĐẠT TIÊU CHUẨN ĐÓNG ĐIỆN' : overallStatus === 'INVESTIGATE' ? 'CẦN XỬ LÝ & KIỂM TRA LẠI TẠI CHỖ' : 'KHÔNG ĐẠT TIÊU CHUẨN AN TOÀN'}</span>
            </div>
            <div>{statusReason}</div>
          </div>

          {/* Section 1: Visual & Mechanical */}
          <div className="mb-6">
            <h2 className="text-xs font-bold uppercase tracking-wider text-slate-900 mb-2 border-b border-slate-200 pb-1">
              1. Kiểm Tra Thị Giác & Cơ Khí (Visual & Mechanical Inspection)
            </h2>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 text-xs">
              {visualChecks.map((v, i) => (
                <div key={i} className="flex items-center justify-between p-2 rounded-lg border border-slate-100 bg-slate-50/50">
                  <span className="text-slate-700">{v.label}</span>
                  <span className={`px-2 py-0.5 rounded font-bold text-[10px] ${
                    v.pass ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-800'
                  }`}>
                    {v.value}
                  </span>
                </div>
              ))}
            </div>
          </div>

          {/* Section 2: Electrical Test Table */}
          <div className="mb-6">
            <h2 className="text-xs font-bold uppercase tracking-wider text-slate-900 mb-2 border-b border-slate-200 pb-1">
              2. Phép Đo Điện & Tiêu Chuẩn So Sánh (Electrical Test Values & Reference Standards)
            </h2>
            <div className="border border-slate-200 rounded-lg overflow-hidden">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="bg-slate-100 text-slate-700 font-bold border-b border-slate-200">
                    <th className="p-2.5">Hạng mục kiểm tra</th>
                    <th className="p-2.5">Giá trị đo thực tế</th>
                    <th className="p-2.5">Tiêu chuẩn NETA ATS-2025</th>
                    <th className="p-2.5">Lần đo trước / Tham chiếu</th>
                    <th className="p-2.5">Đánh giá & Độ lệch</th>
                    <th className="p-2.5 text-center">Trạng thái</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {electricalTable.map((row, idx) => (
                    <tr key={idx} className={row.status === 'FAIL' ? 'bg-rose-50/40' : row.status === 'INVESTIGATE' ? 'bg-amber-50/40' : ''}>
                      <td className="p-2.5 font-semibold text-slate-900">{row.item}</td>
                      <td className="p-2.5 font-mono">{row.measured}</td>
                      <td className="p-2.5 text-slate-600">{row.standard}</td>
                      <td className="p-2.5 text-slate-600">{row.reference}</td>
                      <td className="p-2.5 font-medium">{row.deviation}</td>
                      <td className="p-2.5 text-center">
                        <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                          row.status === 'PASS'
                            ? 'bg-emerald-100 text-emerald-800'
                            : row.status === 'INVESTIGATE'
                            ? 'bg-amber-100 text-amber-800'
                            : 'bg-rose-100 text-rose-800'
                        }`}>
                          {row.status}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Section 3: Risk Warnings & Recommendations */}
          <div className="mb-6 grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/50 text-xs">
              <div className="font-bold text-slate-900 mb-2 flex items-center gap-1.5">
                <AlertTriangle size={14} className="text-amber-600" />
                <span>Cảnh báo rủi ro & Xu hướng</span>
              </div>
              {warnings.length > 0 ? (
                <ul className="list-disc pl-4 space-y-1 text-slate-700">
                  {warnings.map((w, i) => (
                    <li key={i}>{w}</li>
                  ))}
                </ul>
              ) : (
                <div className="text-emerald-700">Tất cả thông số đều đạt tiêu chuẩn tối ưu, không có cảnh báo bất thường.</div>
              )}
            </div>

            <div className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/50 text-xs">
              <div className="font-bold text-slate-900 mb-2 flex items-center gap-1.5">
                <CheckCircle2 size={14} className="text-blue-600" />
                <span>Khuyến nghị xử lý cho FSE</span>
              </div>
              <ul className="list-decimal pl-4 space-y-1 text-slate-700">
                {recommendations.map((r, i) => (
                  <li key={i}>{r}</li>
                ))}
              </ul>
            </div>
          </div>

          {/* Section 4: Signatures */}
          <div className="pt-6 border-t border-slate-300 grid grid-cols-3 gap-4 text-center text-xs">
            <div>
              <div className="font-bold text-slate-800 mb-12">KỸ SƯ HIỆN TRƯỜNG (FSE)</div>
              <div className="font-semibold text-slate-900">{siteInfo.fse_name}</div>
              <div className="text-[10px] text-slate-400">Field Service Engineer</div>
            </div>
            <div>
              <div className="font-bold text-slate-800 mb-12">ĐẠI DIỆN CHỦ ĐẦU TƯ / KHÁCH HÀNG</div>
              <div className="border-b border-dashed border-slate-300 w-32 mx-auto mb-1"></div>
              <div className="text-[10px] text-slate-400">Ký & Ghi rõ họ tên</div>
            </div>
            <div>
              <div className="font-bold text-slate-800 mb-12">LÃNH ĐẠO KỸ THUẬT TEV (QA/LEAD)</div>
              <div className="border-b border-dashed border-slate-300 w-32 mx-auto mb-1"></div>
              <div className="text-[10px] text-slate-400">Technical Lead / Sign-off</div>
            </div>
          </div>

          {/* Footer */}
          <div className="mt-8 pt-3 border-t border-slate-100 text-[10px] text-slate-400 text-center flex justify-between">
            <span>TEV Platform AI • Renewable CMMS</span>
            <span>Báo cáo kỹ thuật được xuất tự động theo quy chuẩn ANSI/NETA ATS-2025</span>
          </div>
        </div>
      </div>
    </div>
  );
};
