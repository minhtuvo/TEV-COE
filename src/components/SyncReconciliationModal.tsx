import React, { useState, useMemo } from 'react';
import {
  ShieldAlert,
  ShieldCheck,
  RefreshCw,
  AlertTriangle,
  ArrowRight,
  Database,
  FileSpreadsheet,
  CheckCircle2,
  Clock,
  ExternalLink,
  Search,
  Filter,
  X,
  Layers,
  ArrowDownLeft,
  ArrowUpRight,
  HelpCircle
} from 'lucide-react';
import {
  ReconciliationReport,
  ReconciliationItem,
  formatVersionDate
} from '../utils/syncReconciliationService';

interface SyncReconciliationModalProps {
  isOpen: boolean;
  onClose: () => void;
  report: ReconciliationReport | null;
  isLoading: boolean;
  onRefresh: () => void;
  onResolveNewerWins: () => void;
  onForcePushToSheets: () => void;
  onPullFromSheetsToFirestore: () => void;
}

export const SyncReconciliationModal: React.FC<SyncReconciliationModalProps> = ({
  isOpen,
  onClose,
  report,
  isLoading,
  onRefresh,
  onResolveNewerWins,
  onForcePushToSheets,
  onPullFromSheetsToFirestore,
}) => {
  const [filterType, setFilterType] = useState<string>('all');
  const [filterStatus, setFilterStatus] = useState<string>('all');
  const [searchTerm, setSearchTerm] = useState<string>('');

  if (!isOpen) return null;

  const filteredItems = useMemo(() => {
    if (!report?.items) return [];
    return report.items.filter(item => {
      // Filter by entity type
      if (filterType !== 'all' && item.entityType !== filterType) return false;

      // Filter by status
      if (filterStatus === 'needs_action') {
        if (item.status === 'in_sync') return false;
      } else if (filterStatus === 'sheets_newer' && item.status !== 'sheets_newer') {
        return false;
      } else if (filterStatus === 'firestore_newer' && item.status !== 'firestore_newer') {
        return false;
      } else if (filterStatus === 'in_sync' && item.status !== 'in_sync') {
        return false;
      }

      // Filter by search term
      if (searchTerm.trim()) {
        const term = searchTerm.toLowerCase();
        const matchId = item.id.toLowerCase().includes(term);
        const matchName = item.name.toLowerCase().includes(term);
        const matchType = item.entityLabel.toLowerCase().includes(term);
        if (!matchId && !matchName && !matchType) return false;
      }

      return true;
    });
  }, [report, filterType, filterStatus, searchTerm]);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4 animate-in fade-in duration-200">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 w-full max-w-5xl max-h-[92vh] flex flex-col overflow-hidden">
        
        {/* Modal Header */}
        <div className="px-6 py-4 border-b border-slate-200 bg-gradient-to-r from-slate-900 to-slate-800 text-white flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="p-2.5 rounded-xl bg-blue-500/20 text-blue-400 border border-blue-500/30">
              <ShieldCheck size={24} />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="text-lg font-bold">Kiểm Tra Phiên Bản & Đối Soát Nhất Quán (Sync Reconciliation)</h3>
                <span className="px-2 py-0.5 rounded-full text-xs font-semibold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                  Sheets API v4 ↔ Firestore
                </span>
              </div>
              <p className="text-xs text-slate-300 mt-0.5">
                So khớp thời gian sửa đổi (version/updatedAt) trước khi ghi đè dữ liệu, ngăn ngừa mất dữ liệu chéo giữa Cloud và Bảng tính.
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={onRefresh}
              disabled={isLoading}
              className="p-2 text-slate-300 hover:text-white rounded-lg hover:bg-slate-700/60 transition-colors disabled:opacity-50"
              title="Kiểm tra lại phiên bản mới nhất"
            >
              <RefreshCw size={18} className={isLoading ? 'animate-spin' : ''} />
            </button>
            <button
              onClick={onClose}
              className="p-2 text-slate-400 hover:text-white rounded-lg hover:bg-slate-700/60 transition-colors"
            >
              <X size={20} />
            </button>
          </div>
        </div>

        {/* Status Alert Banner */}
        {report && (
          <div className={`px-6 py-3 border-b text-sm flex items-center justify-between gap-4 ${
            report.hasConflicts || report.sheetsNewerCount > 0
              ? 'bg-amber-50 border-amber-200 text-amber-900'
              : 'bg-emerald-50 border-emerald-200 text-emerald-900'
          }`}>
            <div className="flex items-center gap-2.5">
              {report.hasConflicts || report.sheetsNewerCount > 0 ? (
                <AlertTriangle size={20} className="text-amber-600 shrink-0" />
              ) : (
                <CheckCircle2 size={20} className="text-emerald-600 shrink-0" />
              )}
              <span className="font-medium">{report.summaryMessage}</span>
            </div>
            <div className="text-xs text-slate-500 shrink-0 flex items-center gap-1">
              <Clock size={13} />
              Đối soát lúc: {new Date(report.timestamp).toLocaleTimeString('vi-VN')}
            </div>
          </div>
        )}

        {/* Modal Body */}
        <div className="p-6 flex-1 overflow-y-auto space-y-5">
          {/* Quick Metrics Cards */}
          {report && (
            <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-6 gap-3">
              <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 text-center">
                <div className="text-xs text-slate-500 font-medium">Tổng mục kiểm tra</div>
                <div className="text-2xl font-black text-slate-900 mt-1">{report.totalChecked}</div>
              </div>

              <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3 text-center">
                <div className="text-xs text-emerald-700 font-medium flex items-center justify-center gap-1">
                  <CheckCircle2 size={13} /> Khớp phiên bản
                </div>
                <div className="text-2xl font-black text-emerald-800 mt-1">{report.inSyncCount}</div>
              </div>

              <div className="bg-blue-50 border border-blue-200 rounded-xl p-3 text-center">
                <div className="text-xs text-blue-700 font-medium flex items-center justify-center gap-1">
                  <Database size={13} /> Firestore mới hơn
                </div>
                <div className="text-2xl font-black text-blue-800 mt-1">{report.firestoreNewerCount}</div>
              </div>

              <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 text-center">
                <div className="text-xs text-amber-700 font-medium flex items-center justify-center gap-1">
                  <FileSpreadsheet size={13} /> Sheets mới hơn
                </div>
                <div className="text-2xl font-black text-amber-800 mt-1">{report.sheetsNewerCount}</div>
              </div>

              <div className="bg-purple-50 border border-purple-200 rounded-xl p-3 text-center">
                <div className="text-xs text-purple-700 font-medium">Chỉ ở Firestore</div>
                <div className="text-2xl font-black text-purple-800 mt-1">{report.onlyInFirestoreCount}</div>
              </div>

              <div className="bg-cyan-50 border border-cyan-200 rounded-xl p-3 text-center">
                <div className="text-xs text-cyan-700 font-medium">Chỉ ở Sheets</div>
                <div className="text-2xl font-black text-cyan-800 mt-1">{report.onlyInSheetsCount}</div>
              </div>
            </div>
          )}

          {/* Filters & Search */}
          <div className="flex flex-col sm:flex-row gap-3 items-center justify-between bg-slate-50 p-3 rounded-xl border border-slate-200">
            <div className="flex flex-wrap items-center gap-2 w-full sm:w-auto">
              <div className="relative flex-1 sm:w-56">
                <Search size={15} className="absolute left-3 top-2.5 text-slate-400" />
                <input
                  type="text"
                  placeholder="Tìm theo mã hoặc tên..."
                  value={searchTerm}
                  onChange={e => setSearchTerm(e.target.value)}
                  className="w-full pl-9 pr-3 py-1.5 text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
              </div>

              <select
                value={filterType}
                onChange={e => setFilterType(e.target.value)}
                className="text-xs py-1.5 px-3 bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
              >
                <option value="all">Tất cả đối tượng</option>
                <option value="workOrder">Phiếu WO (TEV Service)</option>
                <option value="customer">Khách hàng</option>
                <option value="equipment">Thiết bị</option>
                <option value="inventory">Vật tư kho</option>
              </select>

              <select
                value={filterStatus}
                onChange={e => setFilterStatus(e.target.value)}
                className="text-xs py-1.5 px-3 bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
              >
                <option value="all">Tất cả trạng thái</option>
                <option value="needs_action">⚠️ Cần hòa giải (Không khớp)</option>
                <option value="sheets_newer">Sheets mới hơn Firestore</option>
                <option value="firestore_newer">Firestore mới hơn Sheets</option>
                <option value="in_sync">Đã khớp hoàn toàn</option>
              </select>
            </div>

            <span className="text-xs text-slate-500 shrink-0">
              Hiển thị <strong>{filteredItems.length}</strong> / {report?.totalChecked || 0} mục
            </span>
          </div>

          {/* Comparison Table */}
          <div className="border border-slate-200 rounded-xl overflow-hidden shadow-sm">
            <div className="max-h-80 overflow-y-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead className="bg-slate-100 text-slate-700 sticky top-0 border-b border-slate-200 z-10 font-semibold">
                  <tr>
                    <th className="py-2.5 px-3">Phân loại & Mã</th>
                    <th className="py-2.5 px-3">Tên / Tiêu đề</th>
                    <th className="py-2.5 px-3">
                      <div className="flex items-center gap-1.5 text-blue-700">
                        <Database size={13} /> Sửa đổi Firestore
                      </div>
                    </th>
                    <th className="py-2.5 px-3">
                      <div className="flex items-center gap-1.5 text-emerald-700">
                        <FileSpreadsheet size={13} /> Sửa đổi Google Sheets
                      </div>
                    </th>
                    <th className="py-2.5 px-3 text-center">Trạng thái phiên bản</th>
                    <th className="py-2.5 px-3 text-right">Đề xuất hòa giải</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {isLoading ? (
                    <tr>
                      <td colSpan={6} className="py-12 text-center text-slate-500">
                        <RefreshCw size={24} className="animate-spin mx-auto text-blue-600 mb-2" />
                        Đang truy vấn Google Sheets API và Cloud Firestore để so khớp phiên bản...
                      </td>
                    </tr>
                  ) : filteredItems.length === 0 ? (
                    <tr>
                      <td colSpan={6} className="py-8 text-center text-slate-400 italic">
                        Không có mục nào phù hợp với bộ lọc hiện tại.
                      </td>
                    </tr>
                  ) : (
                    filteredItems.map(item => (
                      <tr key={`${item.entityType}-${item.id}`} className="hover:bg-slate-50/80 transition-colors">
                        <td className="py-2 px-3">
                          <span className="font-mono font-medium text-slate-900 bg-slate-100 px-1.5 py-0.5 rounded">
                            {item.id}
                          </span>
                          <div className="text-[11px] text-slate-400 mt-0.5">{item.entityLabel}</div>
                        </td>
                        <td className="py-2 px-3 font-medium text-slate-800 max-w-[200px] truncate" title={item.name}>
                          {item.name}
                        </td>
                        <td className="py-2 px-3 text-slate-600">
                          {formatVersionDate(item.firestoreTimestamp || item.firestoreUpdatedAt)}
                        </td>
                        <td className="py-2 px-3 text-slate-600">
                          {item.sheetsUpdatedAt || item.sheetsTimestamp ? (
                            formatVersionDate(item.sheetsTimestamp || item.sheetsUpdatedAt)
                          ) : (
                            <span className="text-slate-400 italic">Chưa có trên Sheet</span>
                          )}
                        </td>
                        <td className="py-2 px-3 text-center">
                          {item.status === 'in_sync' && (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-emerald-100 text-emerald-800">
                              <CheckCircle2 size={11} /> Khớp
                            </span>
                          )}
                          {item.status === 'firestore_newer' && (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-blue-100 text-blue-800">
                              <ArrowUpRight size={11} /> Firestore mới hơn
                            </span>
                          )}
                          {item.status === 'sheets_newer' && (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-amber-100 text-amber-800">
                              <ArrowDownLeft size={11} /> Sheets mới hơn
                            </span>
                          )}
                          {item.status === 'conflict' && (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-red-100 text-red-800">
                              <AlertTriangle size={11} /> Xung đột
                            </span>
                          )}
                          {item.status === 'only_in_firestore' && (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-purple-100 text-purple-800">
                              Chỉ ở Firestore
                            </span>
                          )}
                          {item.status === 'only_in_sheets' && (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-cyan-100 text-cyan-800">
                              Chỉ ở Sheets
                            </span>
                          )}
                        </td>
                        <td className="py-2 px-3 text-right">
                          {item.recommendedAction === 'push_to_sheets' && (
                            <span className="text-blue-600 font-medium">Ghi đè lên Sheets</span>
                          )}
                          {item.recommendedAction === 'pull_to_firestore' && (
                            <span className="text-amber-600 font-medium">Kéo về Firestore</span>
                          )}
                          {item.recommendedAction === 'no_action' && (
                            <span className="text-slate-400">Giữ nguyên</span>
                          )}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {/* Detailed Advice / Recommendation Box */}
          <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 text-xs text-slate-600 space-y-1.5">
            <div className="font-semibold text-slate-800 flex items-center gap-1.5">
              <HelpCircle size={15} className="text-blue-600" />
              Quy tắc hòa giải nhất quán (Sync Reconciliation Strategy):
            </div>
            <p>
              • <strong>Hòa giải thông minh (Newer Wins - Khuyên dùng):</strong> Hệ thống tự động so khớp từng bản ghi. Bản ghi nào có Firestore mới hơn sẽ được ghi đè lên Sheets; bản ghi nào Google Sheets mới hơn sẽ được nạp về Cloud Firestore. Không bản ghi nào bị mất.
            </p>
            <p>
              • <strong>Đẩy Firestore lên Google Sheets:</strong> Chỉ ghi đè từ Firestore sang Sheets. Nếu Sheets có dữ liệu mới hơn, bạn cần cân nhắc trước khi thực hiện.
            </p>
            <p>
              • <strong>Kéo Sheets về Firestore:</strong> Cập nhật Firestore theo toàn bộ phiên bản hiện tại trên bảng tính Google Sheets.
            </p>
          </div>
        </div>

        {/* Modal Footer with Actions */}
        <div className="px-6 py-4 border-t border-slate-200 bg-slate-50 flex flex-wrap items-center justify-between gap-3">
          <div className="text-xs text-slate-500">
            Hỗ trợ tự động chuẩn hóa 23 cột Work Order, Danh mục Khách hàng, Thiết bị & Vật tư kho.
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={onClose}
              className="px-4 py-2 text-xs font-medium text-slate-600 hover:text-slate-800 hover:bg-slate-200/60 rounded-lg transition-colors"
            >
              Đóng
            </button>

            <button
              onClick={onPullFromSheetsToFirestore}
              disabled={isLoading}
              className="px-3.5 py-2 text-xs font-semibold bg-white border border-amber-300 text-amber-800 hover:bg-amber-50 rounded-lg transition-colors flex items-center gap-1.5 shadow-sm disabled:opacity-50"
            >
              <ArrowDownLeft size={14} />
              Kéo Sheets về Firestore
            </button>

            <button
              onClick={onForcePushToSheets}
              disabled={isLoading}
              className="px-3.5 py-2 text-xs font-semibold bg-white border border-blue-300 text-blue-800 hover:bg-blue-50 rounded-lg transition-colors flex items-center gap-1.5 shadow-sm disabled:opacity-50"
            >
              <ArrowUpRight size={14} />
              Ghi đè Firestore lên Sheets
            </button>

            <button
              onClick={onResolveNewerWins}
              disabled={isLoading}
              className="px-4 py-2 text-xs font-semibold bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg transition-colors flex items-center gap-1.5 shadow-sm disabled:opacity-50"
            >
              <ShieldCheck size={15} />
              Hòa giải thông minh (Newer Wins)
            </button>
          </div>
        </div>

      </div>
    </div>
  );
};
