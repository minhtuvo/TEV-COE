import React from 'react';
import { AlertOctagon, AlertTriangle, CheckCircle2, ShieldCheck, ShieldAlert, Info } from 'lucide-react';
import { evaluatePmpValue, PmpEvaluationResult } from '../utils/pmpStandardService';

export interface PmpHighlightedInputProps {
  label: string;
  subLabel?: string;
  value: any;
  onChange: (val: any) => void;
  ruleKey: string;
  unit?: string;
  overrideMinThreshold?: number;
  placeholder?: string;
  step?: string | number;
  type?: 'number' | 'text';
  disabled?: boolean;
  className?: string;
  size?: 'sm' | 'md' | 'lg';
}

export const PmpHighlightedInput: React.FC<PmpHighlightedInputProps> = ({
  label,
  subLabel,
  value,
  onChange,
  ruleKey,
  unit,
  overrideMinThreshold,
  placeholder,
  step = 'any',
  type = 'number',
  disabled = false,
  className = '',
  size = 'md'
}) => {
  const evalResult: PmpEvaluationResult = evaluatePmpValue(ruleKey, value, overrideMinThreshold);
  const effectiveUnit = unit || evalResult.unit;

  return (
    <div className={`space-y-1.5 ${className}`}>
      {/* Top Label & PMP Badge */}
      <div className="flex items-start justify-between gap-1 flex-wrap">
        <div>
          <label className="block text-xs font-bold text-slate-800 dark:text-slate-200">
            {label}
            {effectiveUnit ? <span className="text-slate-500 font-normal ml-1">({effectiveUnit})</span> : null}
          </label>
          {subLabel && (
            <div className="text-[11px] text-slate-500 dark:text-slate-400 font-normal">
              {subLabel}
            </div>
          )}
        </div>

        {/* Live PMP Evaluation Badge */}
        <span
          className={`inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-black uppercase tracking-wider shadow-2xs transition-all ${
            evalResult.status === 'fail'
              ? 'bg-rose-600 text-white animate-pulse-slow'
              : evalResult.status === 'warning'
              ? 'bg-amber-500 text-slate-950 font-black'
              : 'bg-emerald-600 text-white font-bold'
          }`}
          title={`Tiêu chuẩn PMP: ${evalResult.standardText}`}
        >
          {evalResult.status === 'fail' && <AlertOctagon size={11} className="text-white" />}
          {evalResult.status === 'warning' && <AlertTriangle size={11} className="text-slate-950" />}
          {evalResult.status === 'pass' && <CheckCircle2 size={11} className="text-white" />}
          <span>{evalResult.badgeLabel}</span>
        </span>
      </div>

      {/* Input Box with dynamic PMP highlighting */}
      <div className="relative">
        <input
          type={type}
          step={step}
          value={value ?? ''}
          disabled={disabled}
          placeholder={placeholder || (evalResult.threshold ? `Nhập giá trị (${effectiveUnit})` : '')}
          onChange={(e) => {
            if (type === 'number') {
              const parsed = parseFloat(e.target.value);
              onChange(isNaN(parsed) ? '' : parsed);
            } else {
              onChange(e.target.value);
            }
          }}
          className={`w-full rounded-lg font-mono transition-all outline-none ${
            size === 'lg' ? 'px-3.5 py-2.5 text-base' : size === 'sm' ? 'px-2 py-1 text-xs' : 'px-3 py-1.5 text-sm'
          } ${evalResult.inputBg} ${evalResult.inputBorder} ${evalResult.inputText} ${
            disabled ? 'opacity-60 cursor-not-allowed' : ''
          }`}
        />

        {/* Standard Threshold Indicator */}
        <div className="absolute right-2.5 top-1/2 -translate-y-1/2 pointer-events-none text-[11px] font-bold text-slate-400">
          {effectiveUnit}
        </div>
      </div>

      {/* Dynamic PMP Status Message & Benchmark Banner */}
      <div className="text-[11px] space-y-0.5">
        <div className="flex items-center justify-between text-slate-500">
          <span className="font-semibold text-slate-600 dark:text-slate-300">
            Chuẩn PMP / NETA: <span className="font-bold text-slate-800 dark:text-slate-100">{evalResult.standardText}</span>
          </span>
          <span className="text-[10px] text-slate-400 font-mono">{evalResult.standardCode}</span>
        </div>

        {/* Highlighting alert message if failed or warning */}
        {evalResult.status === 'fail' && (
          <div className="p-1.5 rounded-md bg-rose-100/80 border border-rose-300 text-rose-900 font-medium text-[11px] flex items-start gap-1.5 shadow-2xs">
            <AlertOctagon size={13} className="text-rose-600 shrink-0 mt-0.5" />
            <span>{evalResult.messageVi}</span>
          </div>
        )}

        {evalResult.status === 'warning' && (
          <div className="p-1.5 rounded-md bg-amber-100/70 border border-amber-300 text-amber-950 font-medium text-[11px] flex items-start gap-1.5 shadow-2xs">
            <AlertTriangle size={13} className="text-amber-700 shrink-0 mt-0.5" />
            <span>{evalResult.messageVi}</span>
          </div>
        )}
      </div>
    </div>
  );
};
