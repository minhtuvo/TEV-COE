import React from 'react';
import {
  ArrowUp,
  ArrowDown,
  ArrowRight,
  TrendingUp,
  TrendingDown,
  Minus,
  AlertTriangle,
  CheckCircle2,
  History,
  Activity
} from 'lucide-react';

export interface DgaTrendComparison {
  hasCurrent: boolean;
  hasHistorical: boolean;
  curNum: number;
  histNum: number;
  diff: number;
  percent: number;
  direction: 'up' | 'down' | 'flat' | 'none';
  severity: 'worse' | 'better' | 'neutral' | 'none';
  badgeText: string;
}

export function computeDgaParamTrend(
  currentValue: string | number | undefined,
  historicalValue: number | undefined,
  isHigherBetter: boolean = false
): DgaTrendComparison {
  const curStr = String(currentValue ?? '').trim();
  const curNum = parseFloat(curStr);
  const hasCurrent = curStr !== '' && !isNaN(curNum);
  const hasHistorical = historicalValue !== undefined && !isNaN(historicalValue);
  const histNum = hasHistorical ? historicalValue! : 0;

  if (!hasCurrent || !hasHistorical) {
    return {
      hasCurrent,
      hasHistorical,
      curNum: hasCurrent ? curNum : 0,
      histNum,
      diff: 0,
      percent: 0,
      direction: 'none',
      severity: 'none',
      badgeText: hasHistorical ? `Lần trước: ${histNum}` : 'Mẫu đầu'
    };
  }

  const diff = curNum - histNum;
  // If baseline is 0, handle division safely
  const percent = histNum > 0 ? (diff / histNum) * 100 : (diff === 0 ? 0 : 100);

  // Consider within ±1% and absolute difference < 0.2 as flat/stable
  const isFlat = Math.abs(diff) < 0.05 || (Math.abs(percent) < 1.0 && Math.abs(diff) < 0.2);

  if (isFlat) {
    return {
      hasCurrent,
      hasHistorical,
      curNum,
      histNum,
      diff: 0,
      percent: 0,
      direction: 'flat',
      severity: 'neutral',
      badgeText: `→ 0.0 (Ổn định)`
    };
  }

  const direction: 'up' | 'down' = diff > 0 ? 'up' : 'down';
  let severity: 'worse' | 'better' | 'neutral' | 'none' = 'neutral';

  if (isHigherBetter) {
    // e.g. BD Strength (Dielectric Voltage) or DP: higher is better
    severity = direction === 'up' ? 'better' : 'worse';
  } else {
    // e.g. Combustible Fault Gases (H2, CH4, C2H4, C2H2, etc.): higher is worse
    severity = direction === 'up' ? 'worse' : 'better';
  }

  const sign = diff > 0 ? '+' : '';
  const diffStr = Math.abs(diff) >= 10 ? diff.toFixed(0) : diff.toFixed(1);
  const pctStr = Math.abs(percent) >= 10 ? percent.toFixed(0) : percent.toFixed(1);
  const badgeText = `${sign}${diffStr} (${sign}${pctStr}%)`;

  return {
    hasCurrent,
    hasHistorical,
    curNum,
    histNum,
    diff,
    percent,
    direction,
    severity,
    badgeText
  };
}

interface DgaParamTrendInputProps {
  paramKey: string;
  label: string;
  subLabel?: string;
  unit?: string;
  standard?: string;
  value: string | number;
  onChange: (val: string) => void;
  historicalValue?: number;
  historicalDate?: string;
  isHigherBetter?: boolean;
  placeholder?: string;
  className?: string;
}

export const DgaParamTrendInput: React.FC<DgaParamTrendInputProps> = ({
  paramKey,
  label,
  subLabel,
  unit = 'ppm',
  standard,
  value,
  onChange,
  historicalValue,
  historicalDate,
  isHigherBetter = false,
  placeholder,
  className = ''
}) => {
  const trend = computeDgaParamTrend(value, historicalValue, isHigherBetter);

  // Border & background based on trend severity and presence of value
  let inputBorderClass = 'border-slate-300 focus:border-blue-500 focus:ring-blue-100 bg-white';
  if (trend.hasCurrent && trend.hasHistorical) {
    if (trend.severity === 'worse') {
      inputBorderClass = 'border-rose-400 focus:border-rose-500 focus:ring-rose-100 bg-rose-50/30 text-rose-950 font-semibold';
    } else if (trend.severity === 'better') {
      inputBorderClass = 'border-emerald-400 focus:border-emerald-500 focus:ring-emerald-100 bg-emerald-50/30 text-emerald-950 font-semibold';
    } else if (trend.direction === 'flat') {
      inputBorderClass = 'border-slate-300 focus:border-blue-500 focus:ring-blue-100 bg-slate-50/50 text-slate-800';
    }
  }

  return (
    <div className={`p-2.5 rounded-xl border border-slate-200/90 bg-white shadow-2xs hover:shadow-xs transition-all flex flex-col justify-between ${className}`}>
      {/* Header with Title and Standard */}
      <div className="flex items-center justify-between gap-1 mb-1.5">
        <div className="flex items-center gap-1 min-w-0">
          <span className="text-xs font-bold text-slate-800 uppercase tracking-wide truncate">
            {label}
          </span>
          {subLabel && (
            <span className="text-[10px] text-slate-400 font-medium truncate hidden sm:inline">
              ({subLabel})
            </span>
          )}
        </div>
        {standard && (
          <span className="text-[10px] font-mono font-medium text-slate-400 bg-slate-50 px-1.5 py-0.5 rounded border border-slate-100 flex-shrink-0" title={`Tiêu chuẩn IEEE/IEC: ${standard}`}>
            {standard}
          </span>
        )}
      </div>

      {/* Input Field with Inline Trend Badge */}
      <div className="relative flex items-center">
        <input
          type="number"
          step="any"
          value={value === undefined || value === null ? '' : value}
          onChange={(e) => onChange(e.target.value)}
          placeholder={placeholder || `Nhập ${label}`}
          className={`w-full pl-2.5 pr-14 py-1.5 text-xs rounded-lg border outline-none transition-all shadow-2xs ${inputBorderClass}`}
        />
        <span className="absolute right-2 text-[10px] font-mono text-slate-400 pointer-events-none">
          {unit}
        </span>
      </div>

      {/* Historical Trend Indicator Footer */}
      <div className="mt-1.5 pt-1.5 border-t border-slate-100/90 flex items-center justify-between text-[11px] gap-1">
        {/* Previous Value Reference */}
        <div className="text-[10px] text-slate-500 truncate flex items-center gap-1" title={historicalDate ? `Ghi nhận lần trước vào ${historicalDate}: ${historicalValue} ${unit}` : undefined}>
          <History size={11} className="text-slate-400 flex-shrink-0" />
          <span className="font-medium text-slate-600">
            {historicalValue !== undefined ? `${historicalValue} ${unit}` : 'Chưa có mẫu'}
          </span>
        </div>

        {/* Visual Trend Badge: Up/Down arrow with Delta & Percent */}
        {trend.hasCurrent && trend.hasHistorical ? (
          <div
            className={`inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded-md font-mono text-[10px] font-bold border transition-colors flex-shrink-0 ${
              trend.severity === 'worse'
                ? 'bg-rose-50 text-rose-700 border-rose-200'
                : trend.severity === 'better'
                ? 'bg-emerald-50 text-emerald-700 border-emerald-200'
                : 'bg-slate-100 text-slate-600 border-slate-200'
            }`}
            title={`So với lần trước (${historicalValue} ${unit}): ${trend.diff > 0 ? '+' : ''}${trend.diff.toFixed(2)} ${unit} (${trend.percent > 0 ? '+' : ''}${trend.percent.toFixed(1)}%)`}
          >
            {trend.direction === 'up' && (
              <ArrowUp size={11} className={trend.severity === 'worse' ? 'text-rose-600 stroke-[2.5]' : 'text-emerald-600 stroke-[2.5]'} />
            )}
            {trend.direction === 'down' && (
              <ArrowDown size={11} className={trend.severity === 'better' ? 'text-emerald-600 stroke-[2.5]' : 'text-rose-600 stroke-[2.5]'} />
            )}
            {trend.direction === 'flat' && (
              <ArrowRight size={11} className="text-slate-500" />
            )}
            <span>{trend.badgeText}</span>
          </div>
        ) : (
          <span className="text-[10px] text-slate-400 font-mono italic">
            {historicalValue !== undefined ? 'Nhập giá trị' : 'Gốc'}
          </span>
        )}
      </div>
    </div>
  );
};
