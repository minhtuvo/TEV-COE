import React, { useState, useMemo } from 'react';
import {
  Zap,
  Activity,
  AlertTriangle,
  AlertOctagon,
  CheckCircle2,
  XCircle,
  Sparkles,
  Printer,
  RotateCcw,
  Save,
  Plus,
  Trash2,
  FileCheck2,
  Copy,
  Check,
  Eye,
  Info,
  Layers,
  Radio,
  Volume2,
  Wind,
  Gauge,
  TrendingUp,
  Cpu
} from 'lucide-react';
import {
  NetaPdInputPayload,
  NetaPdEvaluationResult,
  PdSensorMeasurementItem,
  PdSensorType,
  evaluatePdSensorItem
} from '../../server/netaPartialDischargeAnalyzer';
import { TevReportPrintModal } from './TevReportPrintModal';

interface Props {
  onSaveReport?: (reportData: any) => void;
  onNavigateToReports?: () => void;
}

const SAMPLE_PD_DATA: NetaPdInputPayload = {
  site_info: {
    project_name: "Tram Trung Ap TEV Substation 22kV",
    equipment_tag: "SWG-MV-CELL-03",
    equipment_type: "22kV Metal-Clad Switchgear",
    fse_name: "Nguyen Van I",
    test_date: "2026-09-23"
  },
  survey_conditions: {
    system_voltage_kv: 22.0,
    operating_current_amp: 410,
    audible_pd_sound_detected: true,
    ozone_odor_detected: false,
    background_noise_floor_tev_db: 8.0,
    background_noise_floor_hfct_pc: 45
  },
  pd_sensor_measurements: [
    {
      sensor_type: "TEV",
      location: "Cable Compartment Outer Shell",
      measured_value_db: 34.5,
      phase_angle_correlated: true,
      prpd_pattern: "Internal Discharge Pattern (Symmetric pulses in 1st & 3rd quadrants)"
    },
    {
      sensor_type: "HFCT",
      location: "Feeder Cable Earth Strap",
      measured_value_pc: 620,
      phase_angle_correlated: true,
      prpd_pattern: "Surface Tracking Pattern"
    },
    {
      sensor_type: "Airborne_Acoustic",
      location: "Busbar Compartment Air Vent",
      measured_value_db: 1.5,
      phase_angle_correlated: false,
      prpd_pattern: "Random noise / Non-synchronous"
    }
  ],
  previous_test_data: {
    last_test_date: "2025-09-20",
    last_tev_value_db: 18.0
  }
};

export const NetaPartialDischargeChecklist: React.FC<Props> = ({ onSaveReport, onNavigateToReports }) => {
  const [data, setData] = useState<NetaPdInputPayload>(SAMPLE_PD_DATA);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisResult, setAnalysisResult] = useState<NetaPdEvaluationResult | null>(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [jsonText, setJsonText] = useState('');
  const [copied, setCopied] = useState(false);
  const [savedSuccessMessage, setSavedSuccessMessage] = useState<string | null>(null);

  // Live Calculations according to ANSI/NETA ATS-2025 Table 100.23
  const liveStats = useMemo(() => {
    const items = data.pd_sensor_measurements.map(item => evaluatePdSensorItem(item));
    const maxSeverity = items.length > 0 ? Math.max(...items.map(i => i.severity_level)) : 0;

    const tevItem = data.pd_sensor_measurements.find(i => i.sensor_type === 'TEV');
    const hfctItem = data.pd_sensor_measurements.find(i => i.sensor_type === 'HFCT');
    const acousticItem = data.pd_sensor_measurements.find(i => i.sensor_type === 'Airborne_Acoustic' || i.sensor_type === 'Contact_Acoustic');

    const curTev = tevItem?.measured_value_db ?? 0;
    const lastTev = data.previous_test_data?.last_tev_value_db ?? 18.0;
    const tevDelta = parseFloat((curTev - lastTev).toFixed(1));

    return {
      items,
      maxSeverity,
      curTev,
      lastTev,
      tevDelta,
      curHfct: hfctItem?.measured_value_pc ?? 0,
      acousticDb: acousticItem?.measured_value_db ?? 0,
      acousticSync: acousticItem?.phase_angle_correlated ?? false,
      isUrgent: maxSeverity === 3,
      isPractical: maxSeverity === 2
    };
  }, [data]);

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    setAnalysisResult(null);
    setSavedSuccessMessage(null);

    try {
      const response = await fetch('/api/field-service/neta-pd-analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });

      const resJson = await response.json();
      if (resJson.success) {
        setAnalysisResult(resJson.data);
      } else {
        alert('Lỗi phân tích: ' + resJson.error);
      }
    } catch (err: any) {
      console.error('Lỗi khi gọi API phân tích NETA PD:', err);
      alert('Không thể kết nối đến máy chủ AI: ' + err.message);
    } finally {
      setIsAnalyzing(false);
    }
  };

  const handleSaveToCmms = () => {
    if (!analysisResult) {
      alert('Vui lòng bấm "Phân Tích AI Chuyên Gia" trước khi lưu vào kho báo cáo.');
      return;
    }
    if (onSaveReport) {
      onSaveReport({
        data,
        analysisResult,
        equipmentId: data.site_info.equipment_tag,
        date: data.site_info.test_date,
        type: 'Online Partial Discharge Survey (NETA ATS-2025 Sec 11 & Table 100.23)'
      });
      setSavedSuccessMessage(`Đã lưu biên bản khảo sát PD [${data.site_info.equipment_tag}] vào kho báo cáo CMMS!`);
    }
  };

  const handleResetToSample = () => {
    setData(SAMPLE_PD_DATA);
    setAnalysisResult(null);
    setSavedSuccessMessage(null);
  };

  const handleOpenJsonModal = () => {
    setJsonText(JSON.stringify(data, null, 2));
    setShowJsonModal(true);
  };

  const handleApplyJson = () => {
    try {
      const parsed = JSON.parse(jsonText);
      setData(parsed);
      setShowJsonModal(false);
      setAnalysisResult(null);
      setSavedSuccessMessage(null);
    } catch (err: any) {
      alert('Định dạng JSON không hợp lệ: ' + err.message);
    }
  };

  const handleCopyJson = () => {
    navigator.clipboard.writeText(JSON.stringify(data, null, 2));
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleAddSensor = () => {
    const newSensor: PdSensorMeasurementItem = {
      sensor_type: 'TEV',
      location: `Circuit Breaker Front Compartment`,
      measured_value_db: 15.0,
      phase_angle_correlated: true,
      prpd_pattern: 'Random background pulse'
    };
    setData({
      ...data,
      pd_sensor_measurements: [...data.pd_sensor_measurements, newSensor]
    });
  };

  const handleRemoveSensor = (index: number) => {
    setData({
      ...data,
      pd_sensor_measurements: data.pd_sensor_measurements.filter((_, i) => i !== index)
    });
  };

  const handleSensorChange = (index: number, field: keyof PdSensorMeasurementItem, value: any) => {
    const updated = [...data.pd_sensor_measurements];
    updated[index] = {
      ...updated[index],
      [field]: value
    };
    setData({
      ...data,
      pd_sensor_measurements: updated
    });
  };

  return (
    <div className="space-y-6">
      {/* Top Banner & Header */}
      <div className="bg-gradient-to-r from-violet-950 via-indigo-950 to-slate-900 text-white p-5 sm:p-6 rounded-2xl shadow-lg border border-indigo-500/30">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="space-y-1">
            <div className="flex flex-wrap items-center gap-2">
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-extrabold uppercase tracking-wider bg-white/20 text-white backdrop-blur-xs border border-white/30 flex items-center gap-1.5">
                <Radio size={13} className="text-violet-300" />
                ANSI/NETA ATS-2025 Mục 11 & Bảng Table 100.23
              </span>
              <span className="px-2.5 py-0.5 rounded-full text-[11px] font-semibold bg-violet-900/60 text-violet-200 border border-violet-400/40 flex items-center gap-1">
                <Zap size={12} />
                Online Partial Discharge • TEV • HFCT • UHF • Acoustic • Lọc Nhiễu Góc Pha
              </span>
            </div>
            <h1 className="text-xl sm:text-2xl font-black tracking-tight text-white">
              Khảo Sát Phóng Điện Cục Bộ Online (Online Partial Discharge Survey)
            </h1>
            <p className="text-xs sm:text-sm text-indigo-100 max-w-2xl">
              Phát hiện sớm khuyết tật cách điện trung/cao áp qua cảm biến TEV, HFCT, UHF và Siêu âm. Tự động kiểm tra điều kiện đồng bộ góc pha (Phase-Angle Correlation) theo NETA Table 100.23 Note 1.
            </p>
          </div>

          {/* Action Buttons */}
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={handleResetToSample}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title="Khôi phục dữ liệu mẫu hiện trường FSE"
            >
              <RotateCcw size={14} />
              <span>Dữ Liệu Mẫu FSE</span>
            </button>

            <button
              onClick={handleAddSensor}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title="Thêm cảm biến đo PD mới"
            >
              <Plus size={14} />
              <span>Thêm Cảm Biến</span>
            </button>

            <button
              onClick={handleOpenJsonModal}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title="Nhập hoặc sao chép mã JSON"
            >
              <FileCheck2 size={14} />
              <span>Mã JSON</span>
            </button>

            <button
              onClick={() => setShowPrintModal(true)}
              className="px-3 py-2 rounded-xl text-xs font-semibold bg-white/10 hover:bg-white/20 text-white border border-white/20 flex items-center gap-1.5 transition-all shadow-xs"
              title="In hoặc xuất biên bản PDF TEV"
            >
              <Printer size={14} />
              <span>In Biên Bản PDF</span>
            </button>

            <button
              onClick={handleSaveToCmms}
              className="px-3.5 py-2 rounded-xl text-xs font-semibold bg-indigo-700 hover:bg-indigo-600 text-white border border-indigo-400/40 flex items-center gap-1.5 transition-all shadow-md"
              title="Lưu biên bản kiểm tra vào kho báo cáo"
            >
              <Save size={15} />
              <span>Lưu vào Kho Báo Cáo</span>
            </button>

            <button
              onClick={handleAnalyze}
              disabled={isAnalyzing}
              className="px-4 py-2 rounded-xl text-xs font-bold bg-gradient-to-r from-amber-400 to-yellow-300 hover:from-amber-300 hover:to-yellow-200 text-slate-950 flex items-center gap-2 transition-all shadow-lg shadow-indigo-950/40 disabled:opacity-50"
            >
              <Sparkles size={15} className={isAnalyzing ? 'animate-spin' : 'text-slate-950'} />
              <span>{isAnalyzing ? 'Đang Phân Tích NETA...' : 'Phân Tích AI Chuyên Gia'}</span>
            </button>
          </div>
        </div>

        {/* Live NETA Diagnostic Strip */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5 mt-5 pt-4 border-t border-indigo-500/30 text-xs">
          {/* TEV Measurement */}
          <div className="bg-indigo-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-indigo-500/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-indigo-200 block font-medium">TEV (Vỏ Buồng Cáp)</span>
              <span className={`text-base font-black ${liveStats.curTev > 30 ? 'text-rose-300' : liveStats.curTev >= 20 ? 'text-amber-300' : 'text-emerald-300'}`}>
                {liveStats.curTev} dB {liveStats.curTev > 30 ? '(> 30 dB)' : ''}
              </span>
            </div>
            {liveStats.curTev > 30 ? (
              <AlertOctagon size={20} className="text-rose-400 animate-bounce shrink-0" />
            ) : (
              <Zap size={20} className="text-indigo-400 shrink-0" />
            )}
          </div>

          {/* HFCT Measurement */}
          <div className="bg-indigo-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-indigo-500/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-indigo-200 block font-medium">HFCT (Tiếp Địa Màn Chắn)</span>
              <span className={`text-base font-black ${liveStats.curHfct > 500 ? 'text-rose-300' : liveStats.curHfct >= 200 ? 'text-amber-300' : 'text-emerald-300'}`}>
                {liveStats.curHfct} pC {liveStats.curHfct > 500 ? '(> 500 pC)' : ''}
              </span>
            </div>
            {liveStats.curHfct > 500 ? (
              <AlertTriangle size={20} className="text-rose-400 animate-pulse shrink-0" />
            ) : (
              <Activity size={20} className="text-violet-400 shrink-0" />
            )}
          </div>

          {/* Acoustic & Phase Correlation */}
          <div className="bg-indigo-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-indigo-500/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-indigo-200 block font-medium">Acoustic & Góc Pha</span>
              <span className="text-base font-black text-amber-300">
                {liveStats.acousticDb} dB • {liveStats.acousticSync ? 'Đồng bộ' : 'Lọc nhiễu (Note 1)'}
              </span>
            </div>
            <Volume2 size={20} className="text-amber-400 shrink-0" />
          </div>

          {/* Historical Trend Jump */}
          <div className="bg-indigo-950/40 backdrop-blur-xs p-2.5 rounded-xl border border-indigo-500/30 flex items-center justify-between">
            <div>
              <span className="text-[10px] text-indigo-200 block font-medium">Biến Thiên So Kỳ Trước</span>
              <span className="text-base font-black text-rose-300 flex items-center gap-1">
                +{liveStats.tevDelta} dB <TrendingUp size={15} />
              </span>
            </div>
            <Gauge size={20} className="text-rose-400 shrink-0" />
          </div>
        </div>
      </div>

      {/* Save Feedback Toast */}
      {savedSuccessMessage && (
        <div className="bg-emerald-50 border border-emerald-300 p-4 rounded-xl flex items-center justify-between shadow-xs">
          <div className="flex items-center gap-2.5">
            <CheckCircle2 size={18} className="text-emerald-600" />
            <span className="text-xs sm:text-sm font-bold text-emerald-900">{savedSuccessMessage}</span>
          </div>
          {onNavigateToReports && (
            <button
              onClick={onNavigateToReports}
              className="px-3 py-1 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold rounded-lg shadow-xs transition-colors"
            >
              Mở Kho Báo Cáo
            </button>
          )}
        </div>
      )}

      {/* Main Form: Survey Conditions & Measurements */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left Column: Site Info & Survey Conditions */}
        <div className="space-y-6">
          {/* Site Information */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <Radio size={17} className="text-violet-600" />
                Thông Tin Trạm & Ngăn Lộ
              </h2>
              <span className="text-[10px] font-semibold text-slate-500 bg-slate-100 px-2 py-0.5 rounded-full">
                NETA Mục 11
              </span>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="text-slate-600 font-medium block mb-1">Tên Trạm Biến Áp / Dự Án:</label>
                <input
                  type="text"
                  value={data.site_info.project_name}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, project_name: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-violet-500 font-medium text-slate-800"
                />
              </div>

              <div>
                <label className="text-slate-600 font-medium block mb-1">Mã Ngăn Lộ / Tủ Điện (Equipment Tag):</label>
                <input
                  type="text"
                  value={data.site_info.equipment_tag}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, equipment_tag: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-violet-500 font-bold text-violet-700 font-mono"
                />
              </div>

              <div>
                <label className="text-slate-600 font-medium block mb-1">Loại Thiết Bị (Equipment Type):</label>
                <input
                  type="text"
                  value={data.site_info.equipment_type}
                  onChange={(e) => setData({
                    ...data,
                    site_info: { ...data.site_info, equipment_type: e.target.value }
                  })}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-violet-500 font-medium text-slate-800"
                />
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Kỹ Sư FSE Đo:</label>
                  <input
                    type="text"
                    value={data.site_info.fse_name}
                    onChange={(e) => setData({
                      ...data,
                      site_info: { ...data.site_info, fse_name: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-violet-500 font-medium text-slate-800"
                  />
                </div>
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Ngày Khảo Sát:</label>
                  <input
                    type="date"
                    value={data.site_info.test_date}
                    onChange={(e) => setData({
                      ...data,
                      site_info: { ...data.site_info, test_date: e.target.value }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-violet-500 font-medium text-slate-800"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Survey Conditions & Environmental Check (Section 11.1) */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <Gauge size={17} className="text-violet-600" />
                Điều Kiện Đo (NETA Mục 11.1)
              </h2>
            </div>

            <div className="space-y-3 text-xs">
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Điện Áp Làm Việc (kV):</label>
                  <input
                    type="number"
                    step="0.5"
                    value={data.survey_conditions.system_voltage_kv}
                    onChange={(e) => setData({
                      ...data,
                      survey_conditions: { ...data.survey_conditions, system_voltage_kv: parseFloat(e.target.value) || 0 }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-violet-500 font-bold text-slate-800"
                  />
                </div>
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Dòng Điện Tải (A):</label>
                  <input
                    type="number"
                    value={data.survey_conditions.operating_current_amp}
                    onChange={(e) => setData({
                      ...data,
                      survey_conditions: { ...data.survey_conditions, operating_current_amp: parseFloat(e.target.value) || 0 }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-violet-500 font-bold text-slate-800"
                  />
                </div>
              </div>

              {/* Audible sound and Ozone checks */}
              <div className="space-y-2 pt-1 border-t border-slate-100">
                <label className="flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-200 cursor-pointer">
                  <span className="flex items-center gap-2 font-medium text-slate-700">
                    <Volume2 size={15} className={data.survey_conditions.audible_pd_sound_detected ? 'text-rose-600' : 'text-slate-400'} />
                    Âm thanh rít/lách tách (Audible Sound)
                  </span>
                  <input
                    type="checkbox"
                    checked={data.survey_conditions.audible_pd_sound_detected}
                    onChange={(e) => setData({
                      ...data,
                      survey_conditions: { ...data.survey_conditions, audible_pd_sound_detected: e.target.checked }
                    })}
                    className="w-4 h-4 text-violet-600 rounded focus:ring-violet-500"
                  />
                </label>

                <label className="flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-200 cursor-pointer">
                  <span className="flex items-center gap-2 font-medium text-slate-700">
                    <Wind size={15} className={data.survey_conditions.ozone_odor_detected ? 'text-rose-600' : 'text-slate-400'} />
                    Mùi khí Ozone đặc trưng (Ozone Odor)
                  </span>
                  <input
                    type="checkbox"
                    checked={data.survey_conditions.ozone_odor_detected}
                    onChange={(e) => setData({
                      ...data,
                      survey_conditions: { ...data.survey_conditions, ozone_odor_detected: e.target.checked }
                    })}
                    className="w-4 h-4 text-violet-600 rounded focus:ring-violet-500"
                  />
                </label>
              </div>

              {/* Background Noise Floors */}
              <div className="grid grid-cols-2 gap-2 pt-1">
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Nhiễu nền TEV (dB):</label>
                  <input
                    type="number"
                    step="0.5"
                    value={data.survey_conditions.background_noise_floor_tev_db || 8}
                    onChange={(e) => setData({
                      ...data,
                      survey_conditions: { ...data.survey_conditions, background_noise_floor_tev_db: parseFloat(e.target.value) || 0 }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-violet-500 text-slate-800"
                  />
                </div>
                <div>
                  <label className="text-slate-600 font-medium block mb-1">Nhiễu nền HFCT (pC):</label>
                  <input
                    type="number"
                    value={data.survey_conditions.background_noise_floor_hfct_pc || 45}
                    onChange={(e) => setData({
                      ...data,
                      survey_conditions: { ...data.survey_conditions, background_noise_floor_hfct_pc: parseFloat(e.target.value) || 0 }
                    })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-violet-500 text-slate-800"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Historical Trend Card */}
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-3 text-xs">
            <h3 className="font-bold text-slate-800 flex items-center gap-1.5 border-b border-slate-100 pb-2">
              <TrendingUp size={15} className="text-violet-600" />
              Dữ Liệu Khảo Sát Kỳ Trước (Trending)
            </h3>
            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="text-slate-500 text-[10px] block mb-1">Ngày Đo Kỳ Trước:</label>
                <input
                  type="date"
                  value={data.previous_test_data?.last_test_date || '2025-09-20'}
                  onChange={(e) => setData({
                    ...data,
                    previous_test_data: {
                      ...data.previous_test_data,
                      last_test_date: e.target.value
                    }
                  })}
                  className="w-full px-2 py-1 bg-slate-50 border border-slate-200 rounded text-slate-700"
                />
              </div>
              <div>
                <label className="text-slate-500 text-[10px] block mb-1">TEV Kỳ Trước (dB):</label>
                <input
                  type="number"
                  step="0.5"
                  value={data.previous_test_data?.last_tev_value_db || 18.0}
                  onChange={(e) => setData({
                    ...data,
                    previous_test_data: {
                      ...data.previous_test_data,
                      last_tev_value_db: parseFloat(e.target.value) || 0
                    }
                  })}
                  className="w-full px-2 py-1 bg-slate-50 border border-slate-200 rounded font-bold text-slate-800"
                />
              </div>
            </div>
            <div className="p-2 rounded bg-violet-50 border border-violet-200 text-[11px] text-violet-900 font-medium">
              Độ tăng trưởng xung TEV: <strong className="text-rose-600">+{liveStats.tevDelta} dB</strong> (từ {liveStats.lastTev} dB lên {liveStats.curTev} dB).
            </div>
          </div>

          {/* Quick Guide: Table 100.23 */}
          <div className="bg-slate-900 text-white p-4 rounded-2xl border border-slate-800 shadow-sm space-y-2.5 text-xs">
            <h3 className="font-bold text-violet-300 flex items-center gap-1.5 text-xs">
              <Info size={14} />
              Quy Chuẩn NETA ATS-2025 Table 100.23
            </h3>
            <ul className="space-y-1.5 text-[11px] text-slate-300">
              <li className="flex items-start gap-1.5">
                <span className="w-2 h-2 rounded-full bg-violet-400 mt-1 shrink-0" />
                <span><strong>Ghi chú 1 (Quan Trọng):</strong> Không đồng bộ góc pha $\rightarrow$ <em>Không phải PD</em> (khảo sát sau 12 tháng).</span>
              </li>
              <li className="flex items-start gap-1.5">
                <span className="w-2 h-2 rounded-full bg-amber-400 mt-1 shrink-0" />
                <span><strong>TEV:</strong> &lt;20 dB (6 tháng) • 20-30 dB (Practical) • &gt;30 dB (Urgent).</span>
              </li>
              <li className="flex items-start gap-1.5">
                <span className="w-2 h-2 rounded-full bg-amber-400 mt-1 shrink-0" />
                <span><strong>HFCT:</strong> &lt;250 pC (6 tháng) • 200-500 pC (Practical) • &gt;500 pC (Urgent).</span>
              </li>
              <li className="flex items-start gap-1.5">
                <span className="w-2 h-2 rounded-full bg-amber-400 mt-1 shrink-0" />
                <span><strong>Acoustic:</strong> &lt;3 dB (6 tháng) • 3-9 dB (Practical) • &gt;9 dB (Urgent).</span>
              </li>
            </ul>
          </div>
        </div>

        {/* Center & Right Column: Sensor Measurements */}
        <div className="lg:col-span-2 space-y-6">
          <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-100 pb-3">
              <div>
                <h2 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                  <Activity size={17} className="text-violet-600" />
                  Kênh Cảm Biến Khảo Sát PD (TEV, HFCT, UHF, Acoustic)
                </h2>
                <span className="text-[11px] text-slate-500">
                  Tự động phân cấp theo Table 100.23.1 - 100.23.5 và lọc nhiễu theo Ghi chú 1
                </span>
              </div>

              <button
                onClick={handleAddSensor}
                className="px-3 py-1.5 bg-violet-50 hover:bg-violet-100 text-violet-700 text-xs font-bold rounded-lg border border-violet-200 flex items-center gap-1 transition-colors self-start sm:self-auto"
              >
                <Plus size={14} />
                <span>Thêm Kênh Đo</span>
              </button>
            </div>

            {/* Sensor Items List */}
            <div className="space-y-3.5">
              {data.pd_sensor_measurements.map((sensor, idx) => {
                const evalItem = liveStats.items[idx];
                return (
                  <div
                    key={idx}
                    className={`p-4 rounded-xl border transition-all ${
                      evalItem?.severity_level === 3
                        ? 'bg-rose-50/70 border-rose-300 ring-1 ring-rose-300'
                        : evalItem?.severity_level === 2
                        ? 'bg-amber-50/70 border-amber-300'
                        : evalItem?.severity_level === 1
                        ? 'bg-blue-50/70 border-blue-200'
                        : 'bg-slate-50 border-slate-200'
                    }`}
                  >
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-2.5 mb-2.5 border-b border-slate-200/60">
                      <div className="flex items-center gap-2 flex-1">
                        <span className="w-5 h-5 rounded-full bg-slate-200 text-slate-700 text-[10px] font-bold flex items-center justify-center shrink-0">
                          {idx + 1}
                        </span>

                        {/* Sensor Type Selector */}
                        <select
                          value={sensor.sensor_type}
                          onChange={(e) => handleSensorChange(idx, 'sensor_type', e.target.value as PdSensorType)}
                          className="font-bold text-xs bg-white border border-slate-300 rounded px-2 py-1 text-slate-800 focus:ring-1 focus:ring-violet-500"
                        >
                          <option value="TEV">TEV (Transient Earth Voltage)</option>
                          <option value="HFCT">HFCT (High-Frequency CT)</option>
                          <option value="UHF">UHF (Ultra-High Frequency)</option>
                          <option value="Airborne_Acoustic">Airborne Acoustic (Siêu âm không khí)</option>
                          <option value="Contact_Acoustic">Contact Acoustic (Siêu âm tiếp xúc)</option>
                        </select>

                        {/* Location */}
                        <input
                          type="text"
                          value={sensor.location}
                          onChange={(e) => handleSensorChange(idx, 'location', e.target.value)}
                          placeholder="Vị trí đặt cảm biến trên tủ..."
                          className="font-semibold text-xs text-slate-800 bg-white/80 border border-slate-200 rounded px-2 py-1 flex-1 focus:ring-1 focus:ring-violet-500"
                        />
                      </div>

                      <div className="flex items-center gap-2">
                        {/* Status Badge */}
                        <span
                          className={`px-2.5 py-0.5 rounded-full text-[10px] font-black border ${
                            evalItem?.severity_level === 3
                              ? 'bg-rose-600 text-white border-rose-700 animate-pulse'
                              : evalItem?.severity_level === 2
                              ? 'bg-amber-400 text-slate-900 border-amber-500'
                              : evalItem?.severity_level === 1
                              ? 'bg-blue-500 text-white border-blue-600'
                              : 'bg-slate-500 text-white border-slate-600'
                          }`}
                        >
                          {evalItem?.severity_level === 3
                            ? 'REPAIR_URGENT'
                            : evalItem?.severity_level === 2
                            ? 'REPAIR_PRACTICAL'
                            : evalItem?.severity_level === 1
                            ? 'TREND_6M'
                            : 'NOISE (Note 1)'}
                        </span>

                        {data.pd_sensor_measurements.length > 1 && (
                          <button
                            onClick={() => handleRemoveSensor(idx)}
                            className="p-1 text-slate-400 hover:text-rose-600 transition-colors"
                            title="Xóa kênh đo này"
                          >
                            <Trash2 size={15} />
                          </button>
                        )}
                      </div>
                    </div>

                    {/* Sensor Value & Phase Angle Correlation */}
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
                      {/* Measured Value Input */}
                      <div>
                        <label className="text-slate-500 text-[10px] font-medium block">
                          Giá Trị Đo ({sensor.sensor_type === 'HFCT' ? 'pC' : sensor.sensor_type === 'UHF' ? 'dBmV' : 'dB'}):
                        </label>
                        <div className="flex items-center gap-1 mt-0.5">
                          {sensor.sensor_type === 'HFCT' ? (
                            <input
                              type="number"
                              value={sensor.measured_value_pc ?? 620}
                              onChange={(e) => handleSensorChange(idx, 'measured_value_pc', parseFloat(e.target.value) || 0)}
                              className="w-full px-2 py-1 bg-white border border-slate-300 rounded font-black text-slate-900 focus:ring-1 focus:ring-violet-500"
                            />
                          ) : sensor.sensor_type === 'UHF' ? (
                            <input
                              type="number"
                              step="0.5"
                              value={sensor.measured_value_dbmv ?? 15}
                              onChange={(e) => handleSensorChange(idx, 'measured_value_dbmv', parseFloat(e.target.value) || 0)}
                              className="w-full px-2 py-1 bg-white border border-slate-300 rounded font-black text-slate-900 focus:ring-1 focus:ring-violet-500"
                            />
                          ) : (
                            <input
                              type="number"
                              step="0.5"
                              value={sensor.measured_value_db ?? 34.5}
                              onChange={(e) => handleSensorChange(idx, 'measured_value_db', parseFloat(e.target.value) || 0)}
                              className="w-full px-2 py-1 bg-white border border-slate-300 rounded font-black text-slate-900 focus:ring-1 focus:ring-violet-500"
                            />
                          )}
                          <span className="text-slate-500 font-bold shrink-0">
                            {sensor.sensor_type === 'HFCT' ? 'pC' : sensor.sensor_type === 'UHF' ? 'dBmV' : 'dB'}
                          </span>
                        </div>
                      </div>

                      {/* Phase Angle Correlated Toggle */}
                      <div>
                        <label className="text-slate-500 text-[10px] font-medium block">
                          Đồng Bộ Góc Pha (Phase-Angle Correlated):
                        </label>
                        <button
                          type="button"
                          onClick={() => handleSensorChange(idx, 'phase_angle_correlated', !sensor.phase_angle_correlated)}
                          className={`w-full mt-0.5 px-2 py-1 rounded text-xs font-bold border transition-colors flex items-center justify-center gap-1.5 ${
                            sensor.phase_angle_correlated
                              ? 'bg-rose-100 text-rose-800 border-rose-300'
                              : 'bg-slate-200 text-slate-700 border-slate-300'
                          }`}
                        >
                          {sensor.phase_angle_correlated ? (
                            <>
                              <CheckCircle2 size={13} className="text-rose-600" />
                              <span>YES (Có liên quan góc pha)</span>
                            </>
                          ) : (
                            <>
                              <XCircle size={13} className="text-slate-500" />
                              <span>NO (Không đồng bộ pha - Nhiễu)</span>
                            </>
                          )}
                        </button>
                      </div>

                      {/* PRPD Pattern */}
                      <div>
                        <label className="text-slate-500 text-[10px] font-medium block">Dạng Phổ PRPD:</label>
                        <input
                          type="text"
                          value={sensor.prpd_pattern}
                          onChange={(e) => handleSensorChange(idx, 'prpd_pattern', e.target.value)}
                          placeholder="Ví dụ: Internal PD, Surface Tracking, Corona..."
                          className="w-full mt-0.5 px-2 py-1 bg-white border border-slate-300 rounded text-slate-800 text-[11px] focus:ring-1 focus:ring-violet-500"
                        />
                      </div>
                    </div>

                    {/* Rule Evaluation Detail */}
                    <div className="mt-2.5 pt-2 border-t border-slate-200/60 text-[11px] flex flex-col sm:flex-row sm:items-center justify-between gap-1 text-slate-600">
                      <div>
                        <strong>Đánh giá NETA:</strong> {evalItem?.detail}
                      </div>
                      <div className="font-bold text-slate-700 shrink-0">
                        {evalItem?.table_reference}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      </div>

      {/* AI Analysis Output Section */}
      {analysisResult && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden animate-in fade-in slide-in-from-bottom-3 duration-300">
          <div className="p-5 bg-gradient-to-r from-slate-900 via-violet-950 to-slate-900 text-white flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-violet-500/20 border border-violet-400/30 flex items-center justify-center text-violet-300 shrink-0">
                <Sparkles size={20} />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className="text-[10px] uppercase tracking-wider font-extrabold px-2 py-0.5 rounded-full bg-white/10 text-white">
                    TEV Platform AI Evaluation
                  </span>
                  <span className={`text-[11px] font-black px-2 py-0.5 rounded-full border ${
                    analysisResult.overall_status === 'REPAIR_URGENT'
                      ? 'bg-rose-500/30 text-rose-300 border-rose-400/50'
                      : analysisResult.overall_status === 'REPAIR_PRACTICAL'
                      ? 'bg-amber-500/30 text-amber-300 border-amber-400/50'
                      : 'bg-emerald-500/30 text-emerald-300 border-emerald-400/50'
                  }`}>
                    {analysisResult.overall_status}
                  </span>
                </div>
                <h3 className="text-base font-bold text-white mt-0.5">
                  Báo Cáo Đánh Giá Khảo Sát Phóng Điện Cục Bộ Online (NETA ATS-2025 Table 100.23)
                </h3>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={() => setShowPrintModal(true)}
                className="px-3.5 py-1.5 bg-white/10 hover:bg-white/20 text-white text-xs font-semibold rounded-lg flex items-center gap-1.5 transition-colors border border-white/20"
              >
                <Printer size={14} />
                <span>In Báo Cáo</span>
              </button>
              <button
                onClick={handleSaveToCmms}
                className="px-4 py-1.5 bg-indigo-700 hover:bg-indigo-600 text-white text-xs font-bold rounded-lg flex items-center gap-1.5 shadow-md transition-colors"
              >
                <Save size={14} />
                <span>Lưu vào Kho Báo Cáo</span>
              </button>
            </div>
          </div>

          <div className="p-6 prose prose-slate max-w-none text-xs sm:text-sm">
            <div
              className="space-y-4"
              dangerouslySetInnerHTML={{
                __html: analysisResult.markdown_report
                  .replace(/### (.*?)\n/g, '<h3 class="text-base font-bold text-slate-900 border-b border-slate-200 pb-2 mt-4 mb-2 flex items-center gap-2">$1</h3>')
                  .replace(/\n\n/g, '<br/>')
              }}
            />
          </div>
        </div>
      )}

      {/* JSON Import/Export Modal */}
      {showJsonModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-white rounded-2xl max-w-2xl w-full p-6 shadow-2xl border border-slate-200 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <FileCheck2 size={18} className="text-violet-600" />
                <h3 className="font-bold text-sm text-slate-800">
                  Dữ Liệu JSON Khảo Sát PD (NETA Sec 11 & Table 100.23)
                </h3>
              </div>
              <button
                onClick={() => setShowJsonModal(false)}
                className="text-slate-400 hover:text-slate-600 text-sm font-bold"
              >
                ✕
              </button>
            </div>

            <p className="text-xs text-slate-500">
              Kỹ sư FSE có thể sao chép chuỗi JSON này để lưu trữ hoặc dán dữ liệu đo mới từ máy đo PD (UltraTEV, PDTech, Omicron) vào đây:
            </p>

            <textarea
              value={jsonText}
              onChange={(e) => setJsonText(e.target.value)}
              rows={14}
              className="w-full p-3 font-mono text-xs bg-slate-900 text-violet-200 rounded-xl border border-slate-700 focus:outline-hidden focus:ring-2 focus:ring-violet-500"
            />

            <div className="flex items-center justify-between pt-2">
              <button
                onClick={handleCopyJson}
                className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-lg flex items-center gap-1.5 transition-colors"
              >
                {copied ? <Check size={14} className="text-emerald-600" /> : <Copy size={14} />}
                <span>{copied ? 'Đã Sao Chép!' : 'Sao Chép JSON'}</span>
              </button>

              <div className="flex gap-2">
                <button
                  onClick={() => setShowJsonModal(false)}
                  className="px-3.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-medium rounded-lg"
                >
                  Hủy
                </button>
                <button
                  onClick={handleApplyJson}
                  className="px-4 py-1.5 bg-violet-600 hover:bg-violet-500 text-white text-xs font-bold rounded-lg shadow-sm"
                >
                  Áp Dụng Dữ Liệu
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* TEV Report Print Modal */}
      {showPrintModal && (
        <TevReportPrintModal
          isOpen={showPrintModal}
          onClose={() => setShowPrintModal(false)}
          title="BIÊN BẢN KHẢO SÁT PHÓNG ĐIỆN CỤC BỘ ONLINE (ONLINE PARTIAL DISCHARGE SURVEY)"
          standard="ANSI/NETA ATS-2025 Section 11 & Table 100.23"
          siteInfo={{
            project_name: data.site_info.project_name,
            transformer_tag: data.site_info.equipment_tag,
            fse_name: data.site_info.fse_name,
            test_date: data.site_info.test_date
          }}
          overallStatus={analysisResult?.overall_status === 'REPAIR_URGENT' ? 'CRITICAL' : analysisResult?.overall_status === 'REPAIR_PRACTICAL' ? 'INVESTIGATE' : 'PASS'}
          statusReason={analysisResult?.status_reason || 'Đã đối soát với tiêu chuẩn ANSI/NETA ATS-2025 Table 100.23.'}
          visualChecks={[
            { label: 'Điện áp hệ thống vận hành', value: `${data.survey_conditions.system_voltage_kv} kV`, pass: true },
            { label: 'Dòng điện tải mang tải', value: `${data.survey_conditions.operating_current_amp} A`, pass: true },
            { label: 'Âm thanh rít/lách tách (Audible Sound)', value: data.survey_conditions.audible_pd_sound_detected ? 'Phát hiện tiếng rít' : 'Không', pass: !data.survey_conditions.audible_pd_sound_detected },
            { label: 'Mùi khí Ozone (Ozone Odor)', value: data.survey_conditions.ozone_odor_detected ? 'Có mùi' : 'Không phát hiện', pass: !data.survey_conditions.ozone_odor_detected },
            { label: 'Nhiễu nền môi trường TEV & HFCT', value: `TEV: ${data.survey_conditions.background_noise_floor_tev_db || 8} dB, HFCT: ${data.survey_conditions.background_noise_floor_hfct_pc || 45} pC`, pass: true }
          ]}
          electricalTable={liveStats.items.map(item => ({
            item: `${item.sensor_type} (${item.location})`,
            measured: item.display_value,
            standard: item.sensor_type === 'TEV' ? '< 20 dB (Trending), 20-30 dB (Practical), > 30 dB (Urgent)' : item.sensor_type === 'HFCT' ? '< 250 pC (Trending), 200-500 pC, > 500 pC (Urgent)' : '< 3 dB (Trending), 3-9 dB, > 9 dB (Urgent)',
            reference: item.table_reference,
            deviation: `Đồng bộ pha: ${item.phase_angle_correlated ? 'YES' : 'NO (Nhiễu môi trường)'}`,
            status: item.severity_level === 3 ? 'FAIL' : item.severity_level === 2 ? 'INVESTIGATE' : 'PASS'
          }))}
          warnings={[
            liveStats.isUrgent ? 'CẢNH BÁO TỐI CẤP: Mức TEV hoặc HFCT vượt ngưỡng > 30 dB / > 500 pC có đồng bộ góc pha. Nguy cơ đánh thủng cách điện dẫn đến sự cố phóng hồ quang ngắn mạch 3 pha!' : '',
            liveStats.curTev > liveStats.lastTev ? `Tín hiệu TEV tăng thêm +${liveStats.tevDelta} dB so với kỳ trước (${liveStats.lastTev} dB), khuyết tật đang phát triển nhanh.` : ''
          ].filter(Boolean)}
          recommendations={[
            'Lập kế hoạch ngắt điện cô lập Cell 03 (SWG-MV-CELL-03) để sửa chữa CẤP TỐC trong thời gian sớm nhất (as soon as possible).',
            'Thực hiện vệ sinh, nội soi kiểm tra bề mặt cách điện đầu cáp và buồng cáp để tìm vết phóng điện vết (tracking mark) hoặc hư hỏng lớp cách điện.',
            'Thực hiện thử nghiệm Offline PD (Section 7.3.3 / Table 100.6.4) sau khi xử lý để xác nhận đã triệt tiêu hoàn toàn xung phóng điện.'
          ]}
        />
      )}
    </div>
  );
};
