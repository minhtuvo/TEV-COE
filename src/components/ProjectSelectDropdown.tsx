import React, { useState, useEffect, useRef } from 'react';
import {
  Building2,
  ChevronDown,
  Plus,
  Search,
  Check,
  Factory,
  FolderGit2,
  MapPin,
  Sparkles
} from 'lucide-react';
import {
  getRegisteredProjects,
  saveProjectRecord,
  ProjectMasterRecord,
  EVT_PROJECTS_UPDATED
} from '../services/masterDirectoryService';

interface ProjectSelectDropdownProps {
  value: string;
  onChange: (projectName: string, project?: ProjectMasterRecord) => void;
  customerName?: string;
  onCustomerAutoSelect?: (customerName: string) => void;
  placeholder?: string;
  className?: string;
  label?: string;
  required?: boolean;
  darkTheme?: boolean;
  standard?: 'ats' | 'mts' | 'ecs' | 'all';
}

export const ProjectSelectDropdown: React.FC<ProjectSelectDropdownProps> = ({
  value,
  onChange,
  customerName,
  onCustomerAutoSelect,
  placeholder = 'Chọn hoặc thêm Dự Án / Nhà Máy...',
  className = '',
  label,
  required = false,
  darkTheme = false,
  standard
}) => {
  const [projects, setProjects] = useState<ProjectMasterRecord[]>(() => getRegisteredProjects());
  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState('');
  const [quickNewProject, setQuickNewProject] = useState('');

  const dropdownRef = useRef<HTMLDivElement>(null);

  // Sync with project updates across the app
  useEffect(() => {
    const handleUpdate = () => {
      setProjects(getRegisteredProjects());
    };
    window.addEventListener(EVT_PROJECTS_UPDATED, handleUpdate);
    return () => window.removeEventListener(EVT_PROJECTS_UPDATED, handleUpdate);
  }, []);

  // Close when clicking outside
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const cleanVal = (value || '').trim();
  const matchedProject = projects.find(
    (p) =>
      p.name.toLowerCase() === cleanVal.toLowerCase() ||
      (p.id && p.id.toLowerCase() === cleanVal.toLowerCase()) ||
      cleanVal.toLowerCase().includes(p.name.toLowerCase())
  );

  const displayName = matchedProject ? matchedProject.name : cleanVal;

  const handleSelect = (project: ProjectMasterRecord) => {
    onChange(project.name, project);
    if (onCustomerAutoSelect && project.customerName) {
      onCustomerAutoSelect(project.customerName);
    }
    setIsOpen(false);
    setSearch('');
  };

  const handleQuickAdd = () => {
    const nameToAdd = quickNewProject.trim() || search.trim();
    if (!nameToAdd) return;

    // Check if already exists
    const existing = projects.find(
      (p) => p.name.toLowerCase() === nameToAdd.toLowerCase()
    );
    if (existing) {
      handleSelect(existing);
      setQuickNewProject('');
      return;
    }

    const prjId = `PRJ-${Date.now().toString(36).toUpperCase()}`;
    const nowIso = new Date().toISOString();
    const newPrj: ProjectMasterRecord = {
      id: prjId,
      name: nameToAdd,
      customerId: 'KH-GEN',
      customerName: customerName || 'Khách hàng TEV',
      factory: 'Nhà máy chính',
      status: 'chuẩn_bị',
      standard: standard === 'all' || !standard ? 'ats' : standard,
      equipmentList: [],
      notes: '',
      createdAt: nowIso,
      updatedAt: nowIso
    };

    saveProjectRecord(newPrj);
    setProjects(getRegisteredProjects());
    onChange(newPrj.name, newPrj);
    if (onCustomerAutoSelect && newPrj.customerName) {
      onCustomerAutoSelect(newPrj.customerName);
    }
    setQuickNewProject('');
    setSearch('');
    setIsOpen(false);
  };

  const filteredProjects = projects.filter((p) => {
    const q = search.toLowerCase();
    return (
      p.name.toLowerCase().includes(q) ||
      (p.customerName && p.customerName.toLowerCase().includes(q)) ||
      (p.factory && p.factory.toLowerCase().includes(q)) ||
      (p.id && p.id.toLowerCase().includes(q))
    );
  });

  return (
    <div className={`relative ${className}`} ref={dropdownRef}>
      {label && (
        <label
          className={`block text-xs font-semibold mb-1 flex items-center justify-between ${
            darkTheme ? 'text-slate-300' : 'text-slate-700'
          }`}
        >
          <span className="flex items-center gap-1.5">
            <Building2 size={13} className="text-amber-500" />
            {label} {required && <span className="text-red-500">*</span>}
          </span>
          <span
            className={`text-[10px] font-mono px-1.5 py-0.2 rounded border ${
              darkTheme
                ? 'bg-slate-800 text-amber-300 border-slate-700'
                : 'bg-amber-50 text-amber-700 border-amber-200'
            }`}
          >
            Master Dự Án ({projects.length})
          </span>
        </label>
      )}

      {/* Trigger Button */}
      <button
        type="button"
        onClick={() => setIsOpen(!isOpen)}
        className={`w-full text-left rounded-lg px-3 py-2 flex items-center justify-between transition-all border ${
          darkTheme
            ? 'bg-slate-950 border-slate-700 text-white hover:border-amber-400 focus:border-amber-400'
            : 'bg-slate-50 border-slate-300 text-slate-800 hover:border-amber-500 focus:border-amber-500 hover:bg-white'
        } focus:ring-1 focus:ring-amber-500 focus:outline-hidden text-xs shadow-2xs group`}
      >
        <div className="flex items-center gap-2 min-w-0 pr-1">
          <div
            className={`w-6 h-6 rounded-md flex items-center justify-center text-[10px] font-bold shrink-0 ${
              darkTheme
                ? 'bg-amber-500/20 text-amber-400 border border-amber-500/30'
                : 'bg-amber-100 text-amber-800 border border-amber-200'
            }`}
          >
            <FolderGit2 size={13} />
          </div>
          <div className="truncate">
            {displayName ? (
              <div className="font-bold flex items-center gap-1.5 truncate">
                <span className="truncate">{displayName}</span>
                {matchedProject?.factory && (
                  <span
                    className={`text-[9px] px-1 py-0.2 rounded font-normal shrink-0 ${
                      darkTheme
                        ? 'bg-slate-800 text-slate-300 border border-slate-700'
                        : 'bg-slate-200/70 text-slate-600'
                    }`}
                  >
                    {matchedProject.factory}
                  </span>
                )}
              </div>
            ) : (
              <span className={darkTheme ? 'text-slate-500' : 'text-slate-400'}>
                {placeholder}
              </span>
            )}
          </div>
        </div>
        <ChevronDown
          size={14}
          className={`shrink-0 transition-transform ${
            darkTheme ? 'text-slate-400' : 'text-slate-500'
          } group-hover:text-amber-500 ${isOpen ? 'rotate-180' : ''}`}
        />
      </button>

      {/* Floating Dropdown */}
      {isOpen && (
        <div
          className={`absolute left-0 right-0 mt-1.5 rounded-xl shadow-2xl z-50 overflow-hidden border animate-in fade-in-50 zoom-in-95 duration-100 min-w-[320px] ${
            darkTheme
              ? 'bg-slate-900 border-slate-700 text-slate-200'
              : 'bg-white border-slate-200 text-slate-800'
          }`}
        >
          {/* Header & Search */}
          <div
            className={`p-2.5 border-b space-y-2 ${
              darkTheme ? 'bg-slate-950/80 border-slate-800' : 'bg-slate-50 border-slate-200'
            }`}
          >
            <div className="flex items-center justify-between text-xs font-bold">
              <span className="flex items-center gap-1.5">
                <Building2 size={13} className="text-amber-500" />
                Danh Sách Dự Án / Nhà Máy ({projects.length})
              </span>
            </div>

            <div className="relative">
              <Search
                size={12}
                className="absolute left-2.5 top-2.5 text-slate-400 pointer-events-none"
              />
              <input
                type="text"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Tìm theo tên dự án, nhà máy, khách hàng..."
                className={`w-full text-xs pl-8 pr-2.5 py-1.5 rounded-lg border focus:outline-hidden focus:ring-1 focus:ring-amber-500 ${
                  darkTheme
                    ? 'bg-slate-900 border-slate-700 text-white placeholder:text-slate-500'
                    : 'bg-white border-slate-200 text-slate-800 placeholder:text-slate-400'
                }`}
                autoFocus
              />
            </div>

            {/* Quick Add Project Inline */}
            <div className="flex items-center gap-1.5 pt-1">
              <input
                type="text"
                value={quickNewProject}
                onChange={(e) => setQuickNewProject(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    handleQuickAdd();
                  }
                }}
                placeholder="+ Thêm tên dự án / nhà máy mới..."
                className={`flex-1 text-[11px] px-2.5 py-1 rounded-md border border-dashed focus:outline-hidden ${
                  darkTheme
                    ? 'bg-slate-900/60 border-amber-500/50 text-white placeholder:text-amber-300/60 focus:border-amber-400'
                    : 'bg-amber-50/50 border-amber-300 text-slate-800 placeholder:text-amber-700/60 focus:border-amber-500'
                } font-medium`}
              />
              <button
                type="button"
                onClick={handleQuickAdd}
                disabled={!quickNewProject.trim() && !search.trim()}
                className="px-2.5 py-1 bg-amber-600 hover:bg-amber-500 disabled:opacity-40 text-white font-bold text-[11px] rounded-md transition-colors shrink-0 flex items-center gap-1 shadow-xs"
                title="Lưu vào danh mục và chọn ngay"
              >
                <Plus size={12} /> Thêm
              </button>
            </div>
          </div>

          {/* List items */}
          <div className="max-h-60 overflow-y-auto p-1 divide-y divide-slate-100 dark:divide-slate-800">
            {filteredProjects.length === 0 ? (
              <div className="p-4 text-center text-xs space-y-2">
                <p className={darkTheme ? 'text-slate-400' : 'text-slate-500'}>
                  Không tìm thấy dự án "{search}"
                </p>
                {search.trim() && (
                  <button
                    type="button"
                    onClick={handleQuickAdd}
                    className={`w-full py-1.5 px-3 rounded-lg font-bold text-xs flex items-center justify-center gap-1.5 border transition-colors ${
                      darkTheme
                        ? 'bg-amber-950/40 text-amber-300 border-amber-800 hover:bg-amber-900/40'
                        : 'bg-amber-50 text-amber-800 border-amber-200 hover:bg-amber-100'
                    }`}
                  >
                    <Plus size={13} /> Thêm "{search.trim()}" làm dự án / nhà máy mới
                  </button>
                )}
              </div>
            ) : (
              filteredProjects.map((prj) => {
                const isSelected =
                  cleanVal.toLowerCase() === prj.name.toLowerCase() ||
                  matchedProject?.id === prj.id;

                return (
                  <button
                    key={prj.id}
                    type="button"
                    onClick={() => handleSelect(prj)}
                    className={`w-full text-left p-2 rounded-lg flex items-center justify-between text-xs transition-colors ${
                      isSelected
                        ? darkTheme
                          ? 'bg-amber-950/50 text-amber-200 font-bold'
                          : 'bg-amber-50 text-amber-950 font-bold'
                        : darkTheme
                        ? 'hover:bg-slate-800/80 text-slate-200'
                        : 'hover:bg-slate-50 text-slate-700'
                    }`}
                  >
                    <div className="flex items-center gap-2.5 min-w-0">
                      <div
                        className={`w-7 h-7 rounded-md flex items-center justify-center text-[10px] font-black shrink-0 ${
                          isSelected
                            ? 'bg-amber-600 text-white'
                            : darkTheme
                            ? 'bg-slate-800 text-slate-400'
                            : 'bg-slate-100 text-slate-600'
                        }`}
                      >
                        <Building2 size={13} />
                      </div>
                      <div className="truncate">
                        <div className="font-bold flex items-center gap-1.5 truncate">
                          <span>{prj.name}</span>
                          {prj.standard && (
                            <span
                              className={`text-[9px] uppercase px-1 py-0.2 rounded font-mono ${
                                darkTheme
                                  ? 'bg-slate-800 text-indigo-300'
                                  : 'bg-indigo-50 text-indigo-700 border border-indigo-200'
                              }`}
                            >
                              {prj.standard}
                            </span>
                          )}
                        </div>
                        <div className="text-[10px] text-slate-400 flex items-center gap-2 truncate">
                          {prj.customerName && (
                            <span className="truncate">KH: {prj.customerName}</span>
                          )}
                          {prj.factory && (
                            <span className="truncate">• {prj.factory}</span>
                          )}
                        </div>
                      </div>
                    </div>
                    {isSelected && (
                      <Check size={16} className="text-amber-500 shrink-0 ml-2" />
                    )}
                  </button>
                );
              })
            )}
          </div>

          {/* Bottom info */}
          <div
            className={`p-2 border-t flex items-center justify-between text-[11px] ${
              darkTheme ? 'bg-slate-950 border-slate-800 text-slate-400' : 'bg-slate-50 border-slate-200 text-slate-500'
            }`}
          >
            <span>
              Đang chọn: <strong className={darkTheme ? 'text-amber-400' : 'text-slate-800'}>{displayName || 'Chưa chọn'}</strong>
            </span>
            <span className="text-[10px] text-amber-500 font-semibold">Tự động lưu Master</span>
          </div>
        </div>
      )}
    </div>
  );
};
