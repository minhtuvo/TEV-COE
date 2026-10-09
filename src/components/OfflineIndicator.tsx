import React from 'react';
import { WifiOff } from 'lucide-react';
import { usePWAInstall } from '../hooks/usePWAInstall';

export const OfflineIndicator: React.FC = () => {
  const { isOnline } = usePWAInstall();

  if (isOnline) return null;

  return (
    <div className="fixed top-16 left-1/2 -translate-x-1/2 z-50 flex items-center gap-2 rounded-full bg-slate-900/90 text-white px-4 py-1.5 text-xs font-medium shadow-xl border border-slate-700 backdrop-blur-md animate-pulse">
      <WifiOff size={14} className="text-amber-400" />
      <span>Chế độ ngoại tuyến (Offline) — Dữ liệu đã lưu tạm được sử dụng.</span>
    </div>
  );
};
