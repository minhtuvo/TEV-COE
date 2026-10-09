import React, { useState } from 'react';
import { QRCodeSVG } from 'qrcode.react';
import { 
  Smartphone, 
  Download, 
  Share, 
  PlusSquare, 
  CheckCircle2, 
  X, 
  QrCode, 
  WifiOff, 
  Zap, 
  ShieldCheck,
  Copy,
  ExternalLink,
  Laptop,
  HelpCircle,
  Sparkles
} from 'lucide-react';
import { usePWAInstall } from '../hooks/usePWAInstall';

interface PWAInstallModalProps {
  isOpen: boolean;
  onClose: () => void;
}

export const PWAInstallModal: React.FC<PWAInstallModalProps> = ({ isOpen, onClose }) => {
  const { isInstallable, isInstalled, isIOS, isMobile, install } = usePWAInstall();
  const [activeTab, setActiveTab] = useState<'qr' | 'android' | 'ios' | 'faq'>('qr');
  const [isInstalling, setIsInstalling] = useState(false);
  const [installSuccess, setInstallSuccess] = useState(false);
  const [copiedLink, setCopiedLink] = useState(false);

  if (!isOpen) return null;

  const currentUrl = typeof window !== 'undefined' 
    ? (window.location.origin ? `${window.location.origin}/` : window.location.href)
    : 'https://ais-pre-eg22unub6oonsisjuxnl7k-723937582851.asia-southeast1.run.app';

  const handleCopyLink = async () => {
    try {
      await navigator.clipboard.writeText(currentUrl);
      setCopiedLink(true);
      setTimeout(() => setCopiedLink(false), 3000);
    } catch (e) {
      console.warn('Copy failed:', e);
    }
  };

  const handleInstallClick = async () => {
    setIsInstalling(true);
    try {
      const success = await install();
      if (success) {
        setInstallSuccess(true);
        setTimeout(() => {
          onClose();
        }, 2500);
      }
    } catch (e) {
      console.error('PWA install error:', e);
    } finally {
      setIsInstalling(false);
    }
  };

  return (
    <div 
      className="fixed inset-0 z-[9999] flex items-center justify-center bg-slate-950/75 p-3 sm:p-4 backdrop-blur-md animate-in fade-in duration-200"
      onClick={onClose}
    >
      <div 
        className="bg-white rounded-3xl shadow-2xl max-w-lg w-full overflow-hidden border border-slate-200 flex flex-col max-h-[92vh] animate-in zoom-in-95 duration-200"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div className="bg-gradient-to-r from-slate-900 via-blue-950 to-indigo-900 p-5 sm:p-6 text-white relative flex-shrink-0">
          <button 
            onClick={onClose}
            className="absolute top-4 right-4 p-2 text-slate-400 hover:text-white rounded-full hover:bg-white/10 transition-colors"
            title="Đóng"
          >
            <X size={20} />
          </button>
          
          <div className="flex items-center gap-3.5">
            <div className="w-13 h-13 rounded-2xl bg-white p-1 shadow-lg flex-shrink-0 flex items-center justify-center border border-blue-400/30">
              <img 
                src="https://static.wixstatic.com/media/dea24c_6207cba10c9e4e15939adcc4dbd42524~mv2.jpg" 
                alt="TEV Logo" 
                className="w-full h-full object-contain rounded-xl"
              />
            </div>
            <div>
              <div className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-blue-500/20 text-blue-300 text-[11px] font-semibold mb-0.5 border border-blue-400/30">
                <Sparkles size={11} className="text-amber-400" />
                <span>PWA Mobile App</span>
              </div>
              <h2 className="text-lg sm:text-xl font-bold tracking-tight">Cài Đặt App Trên Điện Thoại</h2>
              <p className="text-xs text-slate-300">Biến ứng dụng thành App trên màn hình chính di động</p>
            </div>
          </div>

          {/* Navigation tabs */}
          <div className="grid grid-cols-4 gap-1 mt-4 bg-black/30 p-1 rounded-xl text-[11px] sm:text-xs font-semibold">
            <button
              onClick={() => setActiveTab('qr')}
              className={`py-1.5 px-1 rounded-lg transition-all flex items-center justify-center gap-1 ${
                activeTab === 'qr' ? 'bg-white text-slate-900 shadow font-bold' : 'text-slate-300 hover:text-white'
              }`}
            >
              <QrCode size={13} />
              <span>Quét QR</span>
            </button>
            <button
              onClick={() => setActiveTab('android')}
              className={`py-1.5 px-1 rounded-lg transition-all flex items-center justify-center gap-1 ${
                activeTab === 'android' ? 'bg-white text-slate-900 shadow font-bold' : 'text-slate-300 hover:text-white'
              }`}
            >
              <Download size={13} />
              <span>Android</span>
            </button>
            <button
              onClick={() => setActiveTab('ios')}
              className={`py-1.5 px-1 rounded-lg transition-all flex items-center justify-center gap-1 ${
                activeTab === 'ios' ? 'bg-white text-slate-900 shadow font-bold' : 'text-slate-300 hover:text-white'
              }`}
            >
              <Share size={13} />
              <span>iPhone</span>
            </button>
            <button
              onClick={() => setActiveTab('faq')}
              className={`py-1.5 px-1 rounded-lg transition-all flex items-center justify-center gap-1 ${
                activeTab === 'faq' ? 'bg-white text-slate-900 shadow font-bold' : 'text-slate-300 hover:text-white'
              }`}
            >
              <HelpCircle size={13} />
              <span>Cách Dùng</span>
            </button>
          </div>
        </div>

        {/* Content Body */}
        <div className="p-5 sm:p-6 overflow-y-auto space-y-4">
          {installSuccess ? (
            <div className="py-6 text-center space-y-3">
              <div className="w-16 h-16 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mx-auto animate-bounce">
                <CheckCircle2 size={36} />
              </div>
              <h3 className="text-lg font-bold text-slate-900">Đã cài đặt thành công!</h3>
              <p className="text-sm text-slate-600">Biểu tượng TEV CMMS đã xuất hiện trên màn hình chính của bạn.</p>
            </div>
          ) : isInstalled ? (
            <div className="p-3.5 bg-emerald-50 border border-emerald-200 rounded-2xl text-emerald-900 flex items-start gap-3">
              <CheckCircle2 size={20} className="text-emerald-600 flex-shrink-0 mt-0.5" />
              <div>
                <h4 className="font-bold text-xs sm:text-sm">Ứng dụng đã được cài đặt (Chế độ App)</h4>
                <p className="text-[11px] sm:text-xs text-emerald-700 mt-0.5">
                  Bạn đang sử dụng ở chế độ Standalone. Lần sau bạn chỉ cần bấm icon TEV CMMS trên màn hình điện thoại là ứng dụng mở ngay lập tức!
                </p>
              </div>
            </div>
          ) : null}

          {/* TAB 1: QR CODE FOR MOBILE SCAN */}
          {activeTab === 'qr' && (
            <div className="text-center space-y-4">
              <div className="p-4 bg-slate-50 border border-slate-200 rounded-2xl inline-block shadow-inner mx-auto">
                <QRCodeSVG 
                  value={currentUrl}
                  size={180}
                  level="H"
                  includeMargin={true}
                  imageSettings={{
                    src: "https://static.wixstatic.com/media/dea24c_6207cba10c9e4e15939adcc4dbd42524~mv2.jpg",
                    x: undefined,
                    y: undefined,
                    height: 36,
                    width: 36,
                    excavate: true,
                  }}
                />
              </div>

              <div className="space-y-1">
                <h4 className="font-bold text-slate-900 text-sm flex items-center justify-center gap-1.5">
                  <Smartphone size={16} className="text-blue-600" />
                  <span>Dùng điện thoại quét mã QR</span>
                </h4>
                <p className="text-xs text-slate-500 max-w-xs mx-auto">
                  Mở ứng dụng <strong>Camera</strong> hoặc <strong>Zalo</strong> trên điện thoại quét mã này để mở ngay trên di động.
                </p>
              </div>

              {/* Copy Link Section */}
              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
                <div className="flex items-center justify-between text-xs text-slate-600">
                  <span className="font-medium">Link ứng dụng di động:</span>
                  <button
                    onClick={handleCopyLink}
                    className="inline-flex items-center gap-1 px-2.5 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-[11px] font-bold transition-all shadow-sm"
                  >
                    {copiedLink ? <CheckCircle2 size={13} /> : <Copy size={13} />}
                    <span>{copiedLink ? 'Đã sao chép!' : 'Sao chép link'}</span>
                  </button>
                </div>
                <div className="text-[11px] text-slate-700 font-mono bg-white p-2 rounded-lg border border-slate-200 break-all select-all text-left">
                  {currentUrl}
                </div>
              </div>
            </div>
          )}

          {/* TAB 2: ANDROID / CHROME INSTALL */}
          {activeTab === 'android' && (
            <div className="space-y-4">
              {isInstallable && (
                <div className="p-4 bg-blue-50 border border-blue-200 rounded-2xl text-center space-y-3">
                  <div className="w-12 h-12 bg-blue-600 text-white rounded-full flex items-center justify-center mx-auto shadow-md">
                    <Download size={24} />
                  </div>
                  <div>
                    <h4 className="font-bold text-slate-900 text-base">Cài đặt trực tiếp 1 chạm</h4>
                    <p className="text-xs text-slate-600 mt-1">Trình duyệt Android hỗ trợ cài trực tiếp thành biểu tượng App.</p>
                  </div>
                  <button
                    onClick={handleInstallClick}
                    disabled={isInstalling}
                    className="w-full py-3 px-4 bg-blue-600 hover:bg-blue-700 text-white rounded-xl font-bold text-sm shadow-md transition-all flex items-center justify-center gap-2 disabled:opacity-50"
                  >
                    {isInstalling ? (
                      <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
                    ) : (
                      <Download size={18} />
                    )}
                    <span>Thêm vào Màn hình chính ngay</span>
                  </button>
                </div>
              )}

              <div className="space-y-3 text-sm text-slate-700">
                <div className="p-3 bg-amber-50 border border-amber-200 rounded-xl text-xs text-amber-900 flex items-start gap-2">
                  <Smartphone size={16} className="text-amber-600 flex-shrink-0 mt-0.5" />
                  <span>Cách thêm trên Chrome / Cốc Cốc / Samsung Internet:</span>
                </div>

                <div className="space-y-2">
                  <div className="flex items-start gap-3 p-3 bg-slate-50 rounded-xl border border-slate-200">
                    <div className="w-6 h-6 rounded-full bg-blue-600 text-white text-xs font-bold flex items-center justify-center flex-shrink-0">1</div>
                    <p className="text-xs">Bấm vào dấu <strong>3 chấm (⋮)</strong> ở góc trên bên phải trình duyệt.</p>
                  </div>
                  <div className="flex items-start gap-3 p-3 bg-slate-50 rounded-xl border border-slate-200">
                    <div className="w-6 h-6 rounded-full bg-blue-600 text-white text-xs font-bold flex items-center justify-center flex-shrink-0">2</div>
                    <p className="text-xs">Chọn mục <strong>"Cài đặt ứng dụng"</strong> hoặc <strong>"Thêm vào Màn hình chính"</strong>.</p>
                  </div>
                  <div className="flex items-start gap-3 p-3 bg-slate-50 rounded-xl border border-slate-200">
                    <div className="w-6 h-6 rounded-full bg-blue-600 text-white text-xs font-bold flex items-center justify-center flex-shrink-0">3</div>
                    <p className="text-xs">Bấm <strong>"Cài đặt"</strong> để xác nhận. Icon TEV CMMS sẽ nằm ngoài màn hình điện thoại.</p>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* TAB 3: IPHONE / IPAD (SAFARI) INSTALL */}
          {activeTab === 'ios' && (
            <div className="space-y-3">
              <div className="p-3 bg-blue-50 border border-blue-200 rounded-xl text-xs text-blue-900 flex items-start gap-2">
                <Share size={16} className="text-blue-600 flex-shrink-0 mt-0.5" />
                <span>Trên iPhone, mở link bằng trình duyệt <strong>Safari</strong> và làm theo 3 bước:</span>
              </div>

              <div className="space-y-2.5">
                <div className="flex items-start gap-3 p-3 bg-slate-50 rounded-2xl border border-slate-200">
                  <div className="w-6 h-6 rounded-full bg-slate-900 text-white text-xs font-bold flex items-center justify-center flex-shrink-0 shadow-sm">1</div>
                  <div className="text-xs">
                    <p className="font-bold text-slate-900">Bấm nút "Chia sẻ" (Share)</p>
                    <p className="text-slate-600 mt-0.5">Biểu tượng hình vuông có mũi tên chỉ lên <span className="inline-block px-1.5 py-0.5 bg-white rounded border border-slate-300 font-bold text-blue-600">⎋</span> ở đáy màn hình Safari.</p>
                  </div>
                </div>

                <div className="flex items-start gap-3 p-3 bg-slate-50 rounded-2xl border border-slate-200">
                  <div className="w-6 h-6 rounded-full bg-slate-900 text-white text-xs font-bold flex items-center justify-center flex-shrink-0 shadow-sm">2</div>
                  <div className="text-xs">
                    <p className="font-bold text-slate-900">Chọn "Thêm vào MH chính" (Add to Home Screen)</p>
                    <p className="text-slate-600 mt-0.5">Cuộn menu xuống và bấm dòng có dấu cộng <span className="inline-block px-1.5 py-0.5 bg-white rounded border border-slate-300 font-bold text-slate-800">➕</span>.</p>
                  </div>
                </div>

                <div className="flex items-start gap-3 p-3 bg-slate-50 rounded-2xl border border-slate-200">
                  <div className="w-6 h-6 rounded-full bg-slate-900 text-white text-xs font-bold flex items-center justify-center flex-shrink-0 shadow-sm">3</div>
                  <div className="text-xs">
                    <p className="font-bold text-slate-900">Bấm "Thêm" (Add) ở góc trên bên phải</p>
                    <p className="text-slate-600 mt-0.5">Icon <strong>TEV CMMS</strong> sẽ xuất hiện trên màn hình chính iPhone của bạn.</p>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* TAB 4: HƯỚNG DẪN MỖI LẦN SỬ DỤNG LẠI */}
          {activeTab === 'faq' && (
            <div className="space-y-3 text-xs text-slate-700">
              <div className="p-3 bg-indigo-50 border border-indigo-200 rounded-xl text-indigo-900">
                <h4 className="font-bold flex items-center gap-1.5 text-sm">
                  <HelpCircle size={16} className="text-indigo-600" />
                  <span>Mỗi lần sử dụng lại tôi sẽ làm sao?</span>
                </h4>
                <p className="text-xs text-indigo-800 mt-1">
                  Bạn có 2 cách mở lại ứng dụng cực kỳ đơn giản bất cứ lúc nào:
                </p>
              </div>

              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 space-y-1.5">
                <h5 className="font-bold text-slate-900 text-xs flex items-center gap-1.5 text-blue-600">
                  <Smartphone size={15} />
                  <span>Cách 1: Mở từ Màn hình chính điện thoại (Khuyên dùng)</span>
                </h5>
                <p className="text-slate-600 leading-relaxed">
                  Sau khi đã thực hiện bước <strong>"Thêm vào Màn hình chính"</strong> ở tab Android hoặc iPhone, trên màn hình điện thoại của bạn đã có biểu tượng <strong>TEV CMMS</strong>. Mỗi lần cần kiểm tra hoặc nhập liệu, bạn chỉ cần <strong>chạm vào biểu tượng đó</strong>. Ứng dụng sẽ tự động mở toàn màn hình tức thì như Zalo hay Facebook, không cần gõ link.
                </p>
              </div>

              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 space-y-1.5">
                <h5 className="font-bold text-slate-900 text-xs flex items-center gap-1.5 text-emerald-600">
                  <ExternalLink size={15} />
                  <span>Cách 2: Lưu link vào Bookmark hoặc gửi qua Zalo</span>
                </h5>
                <p className="text-slate-600 leading-relaxed">
                  Bấm nút <strong>"Sao chép link"</strong> ở tab Quét QR và gửi vào mục <em>Cloud của tôi (Zalo)</em> hoặc Đánh dấu trang (Bookmark) trên trình duyệt. Lần sau bấm vào link là vào được ngay.
                </p>
              </div>

              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 space-y-1">
                <h5 className="font-bold text-slate-900 text-xs flex items-center gap-1.5 text-amber-600">
                  <ShieldCheck size={15} />
                  <span>Dữ liệu và tài khoản có bị mất không?</span>
                </h5>
                <p className="text-slate-600 leading-relaxed">
                  Toàn bộ dữ liệu thiết bị, lịch bảo trì và phiếu công việc được lưu trên Cloud Firestore và đồng bộ với Google Sheets. Khi đăng nhập 1 lần, hệ thống sẽ duy trì phiên làm việc cho bạn.
                </p>
              </div>
            </div>
          )}

          {/* Benefits Feature List */}
          <div className="pt-3 border-t border-slate-100 grid grid-cols-2 gap-2 text-[11px] text-slate-600">
            <div className="flex items-center gap-1.5 p-2 bg-slate-50 rounded-lg">
              <Zap size={13} className="text-amber-500" />
              <span>Toàn màn hình không viền</span>
            </div>
            <div className="flex items-center gap-1.5 p-2 bg-slate-50 rounded-lg">
              <Smartphone size={13} className="text-blue-500" />
              <span>Chạy mượt như App gốc</span>
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="p-3.5 bg-slate-50 border-t border-slate-200 flex items-center justify-between flex-shrink-0">
          <button
            onClick={handleCopyLink}
            className="text-xs text-blue-600 hover:text-blue-800 font-semibold flex items-center gap-1"
          >
            <Copy size={13} />
            <span>{copiedLink ? 'Đã sao chép link!' : 'Sao chép link web'}</span>
          </button>
          <button
            onClick={onClose}
            className="px-4 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-800 rounded-xl text-xs font-bold transition-colors"
          >
            Đóng
          </button>
        </div>
      </div>
    </div>
  );
};
