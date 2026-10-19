import React, { useState } from 'react';
import { 
  LayoutDashboard, 
  Activity, 
  ClipboardList, 
  Calendar, 
  Menu, 
  User, 
  Package, 
  Settings, 
  Smartphone, 
  LogOut, 
  X,
  Share2,
  ClipboardCheck,
  Zap,
  Wrench,
  FileSpreadsheet,
  FileText,
  Sparkles
} from 'lucide-react';

interface MobileBottomNavProps {
  activeTab: string;
  onTabChange: (tab: string) => void;
  onOpenInstallModal: () => void;
  onOpenMasterDirectory?: () => void;
  onOpenSidebar: () => void;
  onLogout: () => void;
  userRole?: string | null;
}

export const MobileBottomNav: React.FC<MobileBottomNavProps> = ({
  activeTab,
  onTabChange,
  onOpenInstallModal,
  onOpenMasterDirectory,
  onOpenSidebar,
  onLogout,
  userRole
}) => {
  const [showMoreMenu, setShowMoreMenu] = useState(false);

  const mainTabs = [
    { id: 'dashboard', label: 'Tổng quan', icon: LayoutDashboard },
    { id: 'field-entry', label: 'FSE Hiện trường', icon: ClipboardCheck, badge: 'FSE' },
    { id: 'equipment', label: 'Thiết bị', icon: Activity },
    { id: 'cmms', label: 'Phiếu WO', icon: ClipboardList },
    { id: 'pm-schedule', label: 'Lịch PM', icon: Calendar },
  ];

  const handleSelectTab = (tabId: string) => {
    onTabChange(tabId);
    setShowMoreMenu(false);
  };

  return (
    <>
      {/* MORE MENU BOTTOM SHEET ON MOBILE */}
      {showMoreMenu && (
        <div 
          className="fixed inset-0 z-50 bg-slate-950/60 backdrop-blur-sm lg:hidden flex flex-col justify-end animate-in fade-in"
          onClick={() => setShowMoreMenu(false)}
        >
          <div 
            className="bg-white rounded-t-3xl p-5 shadow-2xl border-t border-slate-200 space-y-4 max-h-[82vh] overflow-y-auto"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <div className="w-8 h-8 rounded-lg bg-blue-600 flex items-center justify-center text-white font-bold text-xs">
                  TEV
                </div>
                <div>
                  <h3 className="font-bold text-slate-900 text-sm">Danh Mục Chức Năng</h3>
                  <p className="text-[11px] text-slate-500">Hệ thống CMMS & FSE Năng Lượng Tái Tạo</p>
                </div>
              </div>
              <button 
                onClick={() => setShowMoreMenu(false)}
                className="p-1.5 text-slate-400 hover:text-slate-700 rounded-full hover:bg-slate-100"
              >
                <X size={18} />
              </button>
            </div>

            {/* Quick Actions Grid */}
            <div className="grid grid-cols-2 gap-2 text-xs">
              <button
                onClick={() => {
                  setShowMoreMenu(false);
                  if (onOpenMasterDirectory) onOpenMasterDirectory();
                }}
                className="p-3 rounded-xl border border-amber-300 bg-amber-50/80 text-amber-950 font-bold flex items-center gap-2.5 transition-colors shadow-2xs"
              >
                <FileSpreadsheet size={18} className="text-amber-600 flex-shrink-0" />
                <div className="text-left">
                  <div className="font-bold flex items-center gap-1">
                    <span>Quản lý dự án</span>
                    <span className="text-[9px] bg-amber-200 text-amber-900 px-1 rounded font-mono">PMP</span>
                  </div>
                  <div className="text-[10px] text-amber-800">Soạn dự án & Thiết bị</div>
                </div>
              </button>

              <button
                onClick={() => handleSelectTab('deep-analysis')}
                className={`p-3 rounded-xl border flex items-center gap-2.5 transition-colors ${
                  activeTab === 'deep-analysis' ? 'bg-indigo-50 border-indigo-400 text-indigo-900 font-bold' : 'bg-slate-50 border-slate-200 text-slate-800'
                }`}
              >
                <Activity size={18} className="text-indigo-600 flex-shrink-0" />
                <div className="text-left">
                  <div className="font-bold flex items-center gap-1">
                    <span>Phân tích Chuyên sâu</span>
                    <Sparkles size={11} className="text-amber-500" />
                  </div>
                  <div className="text-[10px] text-slate-500">DGA MBA, PV Cell, Wind</div>
                </div>
              </button>

              <button
                onClick={() => handleSelectTab('field-entry')}
                className={`p-3 rounded-xl border flex items-center gap-2.5 transition-colors ${
                  activeTab === 'field-entry' ? 'bg-amber-50 border-amber-300 text-amber-900 font-bold' : 'bg-slate-50 border-slate-200 text-slate-800'
                }`}
              >
                <ClipboardCheck size={18} className="text-amber-600 flex-shrink-0" />
                <div className="text-left">
                  <div className="font-bold">Kỹ sư FSE (NETA)</div>
                  <div className="text-[10px] text-slate-500">23 Biên bản thí nghiệm</div>
                </div>
              </button>

              <button
                onClick={() => handleSelectTab('cmms')}
                className={`p-3 rounded-xl border flex items-center gap-2.5 transition-colors ${
                  activeTab === 'cmms' ? 'bg-blue-50 border-blue-300 text-blue-700 font-bold' : 'bg-slate-50 border-slate-200 text-slate-800'
                }`}
              >
                <FileSpreadsheet size={18} className="text-blue-600 flex-shrink-0" />
                <div className="text-left">
                  <div className="font-bold">Trạm FSE & Sheets</div>
                  <div className="text-[10px] text-slate-500">Quản lý phiếu CMMS</div>
                </div>
              </button>

              <button
                onClick={() => handleSelectTab('reports')}
                className={`p-3 rounded-xl border flex items-center gap-2.5 transition-colors ${
                  activeTab === 'reports' ? 'bg-blue-50 border-blue-300 text-blue-700 font-bold' : 'bg-slate-50 border-slate-200 text-slate-800'
                }`}
              >
                <FileText size={18} className="text-blue-600 flex-shrink-0" />
                <div className="text-left">
                  <div className="font-bold">Kho Báo Cáo</div>
                  <div className="text-[10px] text-slate-500">PDF & Lưu trữ NETA</div>
                </div>
              </button>

              <button
                onClick={() => handleSelectTab('customers')}
                className={`p-3 rounded-xl border flex items-center gap-2.5 transition-colors ${
                  activeTab === 'customers' ? 'bg-blue-50 border-blue-300 text-blue-700 font-bold' : 'bg-slate-50 border-slate-200 text-slate-700'
                }`}
              >
                <User size={18} className="text-blue-600 flex-shrink-0" />
                <span>Khách hàng</span>
              </button>

              <button
                onClick={() => handleSelectTab('inventory')}
                className={`p-3 rounded-xl border flex items-center gap-2.5 transition-colors ${
                  activeTab === 'inventory' ? 'bg-blue-50 border-blue-300 text-blue-700 font-bold' : 'bg-slate-50 border-slate-200 text-slate-700'
                }`}
              >
                <Package size={18} className="text-amber-500 flex-shrink-0" />
                <span>Quản lý Kho</span>
              </button>

              <button
                onClick={() => handleSelectTab('settings')}
                className={`p-3 rounded-xl border flex items-center gap-2.5 transition-colors ${
                  activeTab === 'settings' ? 'bg-blue-50 border-blue-300 text-blue-700 font-bold' : 'bg-slate-50 border-slate-200 text-slate-700'
                }`}
              >
                <Settings size={18} className="text-slate-600 flex-shrink-0" />
                <span>Cài đặt hệ thống</span>
              </button>

              <button
                onClick={() => {
                  setShowMoreMenu(false);
                  onOpenInstallModal();
                }}
                className="p-3 rounded-xl border border-emerald-300 bg-emerald-50 text-emerald-800 font-bold flex items-center gap-2.5 transition-colors"
              >
                <Smartphone size={18} className="text-emerald-600 flex-shrink-0" />
                <span>Cài App Điện Thoại</span>
              </button>
            </div>

            {/* Logout button */}
            <button
              onClick={() => {
                setShowMoreMenu(false);
                onLogout();
              }}
              className="w-full py-2.5 rounded-xl border border-red-200 bg-red-50 text-red-600 font-medium text-xs flex items-center justify-center gap-2 hover:bg-red-100 transition-colors"
            >
              <LogOut size={16} />
              <span>Đăng xuất tài khoản</span>
            </button>
          </div>
        </div>
      )}

      {/* FIXED MOBILE BOTTOM NAVBAR */}
      <nav 
        className="lg:hidden fixed bottom-0 inset-x-0 bg-white/95 backdrop-blur-md border-t border-slate-200/90 z-40 px-1 py-1 flex items-center justify-around shadow-[0_-4px_16px_rgba(0,0,0,0.06)] pb-[max(0.5rem,env(safe-area-inset-bottom))]"
      >
        {mainTabs.map(tab => {
          const Icon = tab.icon;
          const isActive = activeTab === tab.id;
          return (
            <button
              key={tab.id}
              onClick={() => handleSelectTab(tab.id)}
              className={`flex-1 py-1 px-0.5 flex flex-col items-center justify-center gap-0.5 relative transition-all ${
                isActive ? 'text-blue-600 font-bold' : 'text-slate-500 hover:text-slate-800 font-medium'
              }`}
            >
              {isActive && (
                <span className="absolute top-0 w-8 h-1 bg-blue-600 rounded-full animate-in fade-in" />
              )}
              <div className="relative">
                <Icon size={19} className={isActive ? 'stroke-[2.5px]' : 'stroke-[1.8px]'} />
                {tab.badge && (
                  <span className="absolute -top-1 -right-2 px-1 bg-amber-500 text-white rounded text-[8px] font-black leading-tight">
                    {tab.badge}
                  </span>
                )}
              </div>
              <span className="text-[9px] leading-tight truncate">{tab.label}</span>
            </button>
          );
        })}

        {/* More Menu Trigger */}
        <button
          onClick={() => setShowMoreMenu(true)}
          className={`flex-1 py-1 px-0.5 flex flex-col items-center justify-center gap-0.5 transition-all ${
            showMoreMenu || ['deep-analysis', 'reports', 'customers', 'inventory', 'settings'].includes(activeTab)
              ? 'text-blue-600 font-bold'
              : 'text-slate-500 hover:text-slate-800 font-medium'
          }`}
        >
          <Menu size={19} className="stroke-[1.8px]" />
          <span className="text-[9px] leading-tight truncate">Thêm</span>
        </button>
      </nav>
    </>
  );
};
