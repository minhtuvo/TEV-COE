import React, { useState, useRef } from 'react';
import {
  Camera,
  Upload,
  Sparkles,
  CheckCircle2,
  AlertTriangle,
  X,
  FileText,
  Zap,
  Droplets,
  Layers,
  ArrowRight,
  RefreshCw,
  Eye,
  SlidersHorizontal
} from 'lucide-react';
import { ExtractedNameplateData } from '../../server/nameplateOcrAnalyzer';

interface Props {
  isOpen: boolean;
  onClose: () => void;
  onApplyData: (extracted: ExtractedNameplateData) => void;
  currentTag?: string;
}

// Sample base64 or SVG generated demo nameplates for instant testing
const SAMPLE_NAMEPLATES = [
  {
    name: 'MBA Dầu Đông Anh EEMC 2000kVA 22/0.4kV',
    tag: 'TR-EEMC-2000KVA-01',
    svgText: `TỔNG CÔNG TY THIẾT BỊ ĐIỆN ĐÔNG ANH (EEMC)
MÁY BIẾN ÁP 3 PHA NGÂM DẦU
Số chế tạo: EEMC-2023-8869
Công suất định mức: 2000 kVA
Điện áp định mức: 22 ± 2x2.5% / 0.4 kV
Dòng điện định mức: 52.5 / 2886 A
Tần số: 50 Hz | Số pha: 3
Tổ đấu dây: Dyn-11
Dầu cách điện: Dầu khoáng (Mineral Oil)
Khối lượng dầu: 1250 kg | Tổng trọng lượng: 5800 kg
Điện áp ngắn mạch Uk%: 5.5% | Tiêu chuẩn: IEC 60076 / TCVN 6306`
  },
  {
    name: 'MBA ABB 2500kVA 35(22)/0.4kV Mineral Oil',
    tag: 'TR-ABB-2500KVA-02',
    svgText: `ABB POWER GRIDS - TRANSFORMER DIVISION
OIL-IMMERSED DISTRIBUTION TRANSFORMER
Serial No: ABB-VN-2024-4512
Rated Power: 2500 kVA | 50 Hz | 3 Phase
High Voltage: 22000 V (Taps ± 2 x 2.5%)
Low Voltage: 400 V | Vector Group: Dyn11
Insulating Fluid: Mineral Oil IEC 60296
Oil Volume: 1480 L | Total Mass: 6950 kg
Short Circuit Impedance: 6.0% | Ambient: 40°C`
  }
];

export const NameplateScannerModal: React.FC<Props> = ({
  isOpen,
  onClose,
  onApplyData,
  currentTag
}) => {
  const [selectedImage, setSelectedImage] = useState<string | null>(null);
  const [mimeType, setMimeType] = useState<string>('image/jpeg');
  const [isScanning, setIsScanning] = useState(false);
  const [scanResult, setScanResult] = useState<ExtractedNameplateData | null>(null);
  const [scanError, setScanError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const cameraInputRef = useRef<HTMLInputElement>(null);

  if (!isOpen) return null;

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setMimeType(file.type || 'image/jpeg');
    const reader = new FileReader();
    reader.onload = (event) => {
      const base64 = event.target?.result as string;
      setSelectedImage(base64);
      setScanResult(null);
      setScanError(null);
    };
    reader.readAsDataURL(file);
  };

  const handleUseSample = (sample: typeof SAMPLE_NAMEPLATES[0]) => {
    // Generate a visual nameplate canvas/image for this sample
    const canvas = document.createElement('canvas');
    canvas.width = 800;
    canvas.height = 500;
    const ctx = canvas.getContext('2d');
    if (ctx) {
      // Plate background (metallic brushed steel)
      ctx.fillStyle = '#1e293b';
      ctx.fillRect(0, 0, 800, 500);
      ctx.strokeStyle = '#94a3b8';
      ctx.lineWidth = 6;
      ctx.strokeRect(15, 15, 770, 470);

      // Inner header plate
      ctx.fillStyle = '#0f172a';
      ctx.fillRect(20, 20, 760, 80);

      // Plate rivets/screws
      ctx.fillStyle = '#cbd5e1';
      [ [35, 35], [765, 35], [35, 465], [765, 465] ].forEach(([x, y]) => {
        ctx.beginPath();
        ctx.arc(x, y, 8, 0, Math.PI * 2);
        ctx.fill();
        ctx.strokeStyle = '#64748b';
        ctx.lineWidth = 2;
        ctx.stroke();
      });

      // Text
      ctx.fillStyle = '#f8fafc';
      ctx.font = 'bold 22px sans-serif';
      ctx.fillText(sample.name, 40, 65);

      ctx.fillStyle = '#e2e8f0';
      ctx.font = '16px monospace';
      const lines = sample.svgText.split('\n');
      lines.forEach((line, idx) => {
        ctx.fillText(line, 40, 130 + idx * 30);
      });

      const dataUrl = canvas.toDataURL('image/jpeg', 0.9);
      setSelectedImage(dataUrl);
      setMimeType('image/jpeg');
      setScanResult(null);
      setScanError(null);
    }
  };

  const handlePerformOcr = async () => {
    if (!selectedImage) return;

    setIsScanning(true);
    setScanError(null);
    try {
      const res = await fetch('/api/field-service/scan-nameplate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          image_base64: selectedImage,
          mime_type: mimeType,
          equipment_category: 'transformer'
        })
      });

      const json = await res.json();
      if (json.success && json.data) {
        setScanResult(json.data);
      } else {
        setScanError(json.error || 'Không thể trích xuất thông số từ ảnh nhãn mác. Vui lòng thử lại với góc chụp sáng hơn.');
      }
    } catch (err: any) {
      setScanError(err.message || 'Lỗi mạng khi gọi dịch vụ nhận diện Gemini Vision.');
    } finally {
      setIsScanning(false);
    }
  };

  const handleApply = () => {
    if (!scanResult) return;
    onApplyData(scanResult);
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/80 backdrop-blur-xs flex items-center justify-center p-3 sm:p-5 animate-fade-in overflow-y-auto">
      <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl max-w-4xl w-full overflow-hidden flex flex-col my-auto max-h-[92vh]">
        {/* Header */}
        <div className="bg-gradient-to-r from-amber-800 via-yellow-900 to-slate-900 text-white p-4 sm:p-5 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <span className="p-2.5 bg-amber-500/20 rounded-xl border border-amber-400/30 text-amber-300">
              <Camera size={22} />
            </span>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-[10px] font-black uppercase tracking-wider bg-amber-500 text-slate-950 px-2 py-0.5 rounded">
                  AI VISION OCR
                </span>
                <span className="text-xs text-amber-200/80">Gemini 3.8 Flash Multimodal</span>
              </div>
              <h2 className="text-base sm:text-lg font-bold">
                Quét Nhãn Mác Nameplate Máy Biến Áp Bằng Camera
              </h2>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-300 hover:text-white hover:bg-white/10 transition-colors"
          >
            <X size={20} />
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-5 overflow-y-auto space-y-5 flex-1">
          {/* Instructions banner */}
          <div className="bg-amber-50 border border-amber-200/80 rounded-xl p-3.5 text-xs text-amber-900 flex items-start gap-2.5">
            <Sparkles size={18} className="text-amber-600 shrink-0 mt-0.5" />
            <div>
              <strong className="font-bold">Cách thức hoạt động ngoài Site:</strong>
              <p className="mt-0.5 text-amber-800">
                FSE dùng camera điện thoại hướng thẳng vào bảng tên kim loại của máy biến áp (hoặc chọn ảnh chụp từ thư viện). AI Vision sẽ tự động nhận diện số chế tạo, công suất kVA, cấp điện áp sơ/thứ cấp, loại dầu (Mineral Oil / Silicone / Ester) và dung tích dầu, sau đó tự điền trực tiếp vào biểu mẫu checklist.
              </p>
            </div>
          </div>

          {/* Action buttons to capture / select / try sample */}
          <div className="flex flex-wrap items-center justify-between gap-3 bg-slate-50 p-3 rounded-xl border border-slate-200">
            <div className="flex items-center gap-2 flex-wrap">
              {/* Native mobile camera trigger */}
              <input
                ref={cameraInputRef}
                type="file"
                accept="image/*"
                capture="environment"
                className="hidden"
                onChange={handleFileChange}
              />
              <button
                type="button"
                onClick={() => cameraInputRef.current?.click()}
                className="px-3.5 py-2 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-xs font-bold flex items-center gap-2 shadow-xs transition-all active:scale-95"
              >
                <Camera size={16} /> Chụp bằng Camera Điện thoại
              </button>

              {/* Upload file from gallery or PC */}
              <input
                ref={fileInputRef}
                type="file"
                accept="image/*"
                className="hidden"
                onChange={handleFileChange}
              />
              <button
                type="button"
                onClick={() => fileInputRef.current?.click()}
                className="px-3.5 py-2 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 rounded-lg text-xs font-bold flex items-center gap-2 transition-all"
              >
                <Upload size={16} /> Chọn ảnh từ Thư viện
              </button>
            </div>

            <div className="flex items-center gap-1.5 text-xs text-slate-500">
              <span className="font-medium">Hoặc thử mẫu:</span>
              {SAMPLE_NAMEPLATES.map((sample, idx) => (
                <button
                  key={idx}
                  type="button"
                  onClick={() => handleUseSample(sample)}
                  className="px-2.5 py-1 bg-amber-100/70 hover:bg-amber-200 text-amber-900 rounded font-semibold text-[11px] border border-amber-300/60 transition-all"
                >
                  {sample.tag}
                </button>
              ))}
            </div>
          </div>

          {/* Main Grid: Image Preview & Results */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
            {/* Left: Image Canvas / Viewfinder */}
            <div className="space-y-3">
              <div className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center justify-between">
                <span>Ảnh chụp nhãn Nameplate</span>
                {selectedImage && (
                  <button
                    onClick={() => {
                      setSelectedImage(null);
                      setScanResult(null);
                      setScanError(null);
                    }}
                    className="text-slate-400 hover:text-rose-600 text-[11px] flex items-center gap-1 font-normal"
                  >
                    <RefreshCw size={12} /> Chụp lại
                  </button>
                )}
              </div>

              <div className="border-2 border-dashed border-slate-300 rounded-xl bg-slate-900/5 min-h-[260px] flex flex-col items-center justify-center p-3 relative overflow-hidden group">
                {selectedImage ? (
                  <div className="relative w-full h-full max-h-[340px] flex items-center justify-center bg-slate-950 rounded-lg overflow-hidden">
                    <img
                      src={selectedImage}
                      alt="Nameplate Preview"
                      className="max-h-[320px] w-auto object-contain rounded"
                    />
                    {/* Viewfinder crosshairs overlay */}
                    <div className="absolute inset-4 border border-amber-400/40 pointer-events-none rounded flex flex-col justify-between p-2">
                      <div className="flex justify-between">
                        <div className="w-4 h-4 border-t-2 border-l-2 border-amber-400"></div>
                        <div className="w-4 h-4 border-t-2 border-r-2 border-amber-400"></div>
                      </div>
                      <div className="flex justify-between">
                        <div className="w-4 h-4 border-b-2 border-l-2 border-amber-400"></div>
                        <div className="w-4 h-4 border-b-2 border-r-2 border-amber-400"></div>
                      </div>
                    </div>
                  </div>
                ) : (
                  <div className="text-center py-10 px-4 space-y-3">
                    <div className="w-14 h-14 mx-auto rounded-full bg-amber-100 text-amber-600 flex items-center justify-center border border-amber-200">
                      <Camera size={26} />
                    </div>
                    <div>
                      <p className="text-xs font-bold text-slate-800">Chưa có ảnh nhãn mác</p>
                      <p className="text-[11px] text-slate-500 mt-0.5">
                        Bấm nút chụp camera hoặc tải ảnh nhãn kim loại của máy biến áp lên đây.
                      </p>
                    </div>
                  </div>
                )}
              </div>

              {selectedImage && (
                <button
                  type="button"
                  onClick={handlePerformOcr}
                  disabled={isScanning}
                  className="w-full py-2.5 bg-gradient-to-r from-amber-600 to-yellow-600 hover:from-amber-500 hover:to-yellow-500 text-slate-950 font-black rounded-xl text-xs flex items-center justify-center gap-2 shadow-md transition-all disabled:opacity-60"
                >
                  <Sparkles size={16} />
                  {isScanning ? 'Đang đọc và bóc tách thông số bằng AI...' : 'Phân tích Nhãn mác bằng AI Vision'}
                </button>
              )}
            </div>

            {/* Right: Extracted Parameters */}
            <div className="space-y-3">
              <div className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center justify-between">
                <span>Thông số bóc tách tự động (AI OCR)</span>
                {scanResult && (
                  <span className="text-[10px] text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded font-bold border border-emerald-200 flex items-center gap-1">
                    <CheckCircle2 size={12} /> Đã nhận diện thành công
                  </span>
                )}
              </div>

              {scanError && (
                <div className="p-3 bg-rose-50 border border-rose-200 rounded-xl text-xs text-rose-800 flex items-start gap-2">
                  <AlertTriangle size={16} className="text-rose-600 shrink-0 mt-0.5" />
                  <div>
                    <strong className="font-bold">Lỗi nhận diện:</strong> {scanError}
                  </div>
                </div>
              )}

              {scanResult ? (
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 space-y-3 max-h-[380px] overflow-y-auto text-xs">
                  <div className="grid grid-cols-2 gap-2.5">
                    <div className="p-2.5 bg-white rounded-lg border border-slate-200">
                      <span className="text-[10px] text-slate-500 uppercase block font-semibold">Mã hiệu thiết bị (Tag)</span>
                      <span className="font-bold text-slate-900 font-mono text-sm">{scanResult.equipment_tag}</span>
                    </div>
                    <div className="p-2.5 bg-white rounded-lg border border-slate-200">
                      <span className="text-[10px] text-slate-500 uppercase block font-semibold">Hãng chế tạo</span>
                      <span className="font-bold text-slate-900">{scanResult.manufacturer || 'N/A'}</span>
                    </div>
                    <div className="p-2.5 bg-white rounded-lg border border-slate-200">
                      <span className="text-[10px] text-slate-500 uppercase block font-semibold">Số chế tạo (Serial No)</span>
                      <span className="font-bold text-slate-900 font-mono">{scanResult.serial_number || 'N/A'}</span>
                    </div>
                    <div className="p-2.5 bg-white rounded-lg border border-slate-200">
                      <span className="text-[10px] text-slate-500 uppercase block font-semibold">Công suất định mức</span>
                      <span className="font-bold text-amber-700">{scanResult.rated_power_kva ? `${scanResult.rated_power_kva} kVA` : 'N/A'}</span>
                    </div>
                    <div className="p-2.5 bg-white rounded-lg border border-slate-200">
                      <span className="text-[10px] text-slate-500 uppercase block font-semibold">Điện áp Sơ / Thứ cấp</span>
                      <span className="font-bold text-slate-900">
                        {scanResult.primary_voltage_kv} kV / {scanResult.secondary_voltage_v} V
                      </span>
                    </div>
                    <div className="p-2.5 bg-white rounded-lg border border-slate-200">
                      <span className="text-[10px] text-slate-500 uppercase block font-semibold">Loại dầu / Dung tích</span>
                      <span className="font-bold text-slate-900">
                        {scanResult.oil_type} ({scanResult.oil_volume_liters ? `${scanResult.oil_volume_liters} L` : scanResult.oil_weight_kg ? `${scanResult.oil_weight_kg} kg` : 'N/A'})
                      </span>
                    </div>
                    <div className="p-2.5 bg-white rounded-lg border border-slate-200">
                      <span className="text-[10px] text-slate-500 uppercase block font-semibold">Tổ đấu dây & Tần số</span>
                      <span className="font-bold text-slate-900">
                        {scanResult.vector_group || 'Dyn11'} | {scanResult.rated_frequency_hz || 50} Hz
                      </span>
                    </div>
                    <div className="p-2.5 bg-white rounded-lg border border-slate-200">
                      <span className="text-[10px] text-slate-500 uppercase block font-semibold">Điện áp ngắn mạch %Uk</span>
                      <span className="font-bold text-slate-900">{scanResult.impedance_percent_z ? `${scanResult.impedance_percent_z}%` : '5.5%'}</span>
                    </div>
                  </div>

                  {/* Raw text preview */}
                  <div className="p-2.5 bg-white rounded-lg border border-slate-200">
                    <span className="text-[10px] text-slate-500 uppercase block font-semibold mb-1">Dữ liệu văn bản OCR đọc được</span>
                    <p className="text-[11px] text-slate-700 font-mono whitespace-pre-line leading-relaxed max-h-24 overflow-y-auto bg-slate-50 p-2 rounded">
                      {scanResult.full_nameplate_text}
                    </p>
                  </div>
                </div>
              ) : (
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-8 text-center text-slate-400 text-xs min-h-[260px] flex flex-col items-center justify-center space-y-2">
                  <FileText size={32} className="text-slate-300" />
                  <p className="font-medium text-slate-600">Chưa có dữ liệu trích xuất</p>
                  <p className="text-[11px] text-slate-400 max-w-xs">
                    Sau khi bạn chụp ảnh hoặc chọn ảnh mẫu, bấm "Phân tích Nhãn mác bằng AI Vision" để xem bảng dữ liệu kỹ thuật.
                  </p>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Modal Footer */}
        <div className="p-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between">
          <div className="text-xs text-slate-500">
            {scanResult ? (
              <span className="text-emerald-700 font-bold flex items-center gap-1">
                <CheckCircle2 size={14} /> Sẵn sàng điền vào biểu mẫu kiểm định
              </span>
            ) : (
              <span>Vui lòng chụp ảnh nhãn mác rõ nét để AI nhận diện tối ưu</span>
            )}
          </div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 rounded-lg text-xs font-semibold"
            >
              Hủy
            </button>
            <button
              type="button"
              disabled={!scanResult}
              onClick={handleApply}
              className="px-5 py-2 bg-amber-600 hover:bg-amber-700 text-white font-bold rounded-lg text-xs flex items-center gap-1.5 shadow-sm transition-all disabled:opacity-50"
            >
              <CheckCircle2 size={15} /> Tự Động Điền Vào Checklist
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
