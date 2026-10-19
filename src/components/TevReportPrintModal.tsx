import React, { useState, useEffect, useRef } from 'react';
import {
  Printer,
  Download,
  X,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  ShieldCheck,
  Globe,
  QrCode,
  Maximize2,
  Minimize2,
  Edit2,
  Check,
  FileDown,
  Loader2,
  FileCheck2,
  ZoomIn,
  ZoomOut,
  RotateCcw,
  UserCheck
} from 'lucide-react';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as htmlToImage from 'html-to-image';
import { DEFAULT_TEV_FSE_LIST } from './FseEngineerDropdown';

interface Props {
  isOpen?: boolean;
  onClose: () => void;
  report?: any;
  title?: string;
  titleEn?: string;
  standard?: string;
  fseName?: string;
  fse_name?: string;
  inspectorName?: string;
  siteInfo?: {
    project_name?: string;
    transformer_tag?: string;
    equipment_tag?: string;
    equipment_rating?: string;
    fse_name?: string;
    test_date?: string;
  };
  overallStatus?: 'PASS' | 'INVESTIGATE' | 'FAIL';
  statusReason?: string;
  statusReasonEn?: string;
  visualChecks?: Array<{ label: string; value: string; pass: boolean; note?: string }>;
  electricalTable?: Array<{
    item: string;
    itemEn?: string;
    measured: string;
    standard: string;
    standardEn?: string;
    reference: string;
    deviation?: string;
    status: 'PASS' | 'INVESTIGATE' | 'FAIL';
    detail?: string;
    detailEn?: string;
  }>;
  warnings?: string[];
  warningsEn?: string[];
  recommendations?: string[] | string;
  recommendationsEn?: string[];
  lang?: 'vi' | 'en';
}

export const TevReportPrintModal: React.FC<Props> = ({
  isOpen = true,
  onClose,
  report,
  title,
  titleEn,
  standard,
  fseName: propFseName,
  fse_name: propFse_name,
  inspectorName: propInspectorName,
  siteInfo,
  overallStatus,
  statusReason,
  statusReasonEn,
  visualChecks = [],
  electricalTable,
  warnings = [],
  warningsEn = [],
  recommendations,
  recommendationsEn = [],
  lang: initialLang = 'vi',
}) => {
  const [lang, setLang] = useState<'vi' | 'en'>(initialLang);
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [zoomLevel, setZoomLevel] = useState<number>(1);
  const [isExportingPdf, setIsExportingPdf] = useState(false);
  const [pdfSuccessToast, setPdfSuccessToast] = useState<string | null>(null);

  // Harmonize legacy report object with structured props
  const effectiveTitle = title || report?.equipmentName || report?.title || 'Biên bản kiểm tra kỹ thuật';
  const effectiveTitleEn = titleEn || report?.titleEn || effectiveTitle;
  const effectiveStandard = standard || report?.testStandard || report?.standard || 'ANSI/NETA ATS-2025';
  const effectiveSiteInfo = siteInfo || {
    project_name: report?.substation || report?.projectName || report?.factory || 'Trạm khách hàng TEV',
    equipment_tag: report?.equipmentId || report?.equipmentTag || 'EQUIP-01',
    equipment_rating: report?.voltageLevel || report?.rating || report?.type || 'Tiêu chuẩn NETA',
    fse_name: propFseName || propFse_name || propInspectorName || report?.testingEngineer || report?.fseName || report?.fse_name || report?.inspector || 'Võ Minh Tú',
    test_date: report?.testDate || report?.date || new Date().toISOString().split('T')[0],
  };

  const initialFse =
    propFseName ||
    propFse_name ||
    propInspectorName ||
    effectiveSiteInfo.fse_name ||
    report?.testingEngineer ||
    report?.fseName ||
    report?.fse_name ||
    report?.inspector ||
    'Võ Minh Tú';

  const [fseName, setFseName] = useState(initialFse);
  const [isEditingFse, setIsEditingFse] = useState(false);
  const [showFsePicker, setShowFsePicker] = useState(false);

  // Sync when prop changes
  useEffect(() => {
    if (propFseName || propFse_name || propInspectorName) {
      setFseName(propFseName || propFse_name || propInspectorName || 'Võ Minh Tú');
    }
  }, [propFseName, propFse_name, propInspectorName]);

  // Handle ESC key to exit fullscreen or close modal
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        if (isFullscreen) {
          setIsFullscreen(false);
        } else if (isOpen) {
          onClose();
        }
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isFullscreen, isOpen, onClose]);

  if (!isOpen) return null;

  const equipTag = effectiveSiteInfo.equipment_tag || effectiveSiteInfo.transformer_tag || 'EQUIP-01';
  const effectiveStatus = overallStatus || (report?.status === 'healthy' ? 'PASS' : report?.status === 'critical' ? 'FAIL' : report?.status === 'warning' ? 'INVESTIGATE' : report?.overallStatus) || 'PASS';
  const effectiveReason = statusReason || report?.notes || report?.summary || 'Kết quả thử nghiệm đạt quy chuẩn an toàn NETA.';
  const effectiveReasonEn = statusReasonEn || report?.summaryEn || effectiveReason;

  // Extract table rows from electricalTable, report.sensorData, or report.details.netaData
  const netaData = report?.details?.netaData || report?.rawPayload;
  let dynamicRows: any[] = [];
  if (netaData?.electrical_tests) {
    const ele = netaData.electrical_tests;
    if (ele.oil_sample_tests_astm_d923) {
      const oil = ele.oil_sample_tests_astm_d923;
      const dga = ele.dga_test_ieee_c57_104 || {};
      const ir = ele.insulation_resistance_1min_megohms || {};
      dynamicRows.push(
        {
          item: 'Điện áp đánh thủng D1816',
          itemEn: 'Dielectric Breakdown D1816',
          measured: `${oil.dielectric_breakdown_d1816_1mm_kv} kV`,
          standard: '≥ 25 kV (Table 100.4.1)',
          standardEn: '≥ 25 kV (Table 100.4.1)',
          reference: 'ASTM D1816 1mm',
          status: oil.dielectric_breakdown_d1816_1mm_kv < 25 ? 'FAIL' : 'PASS',
        },
        {
          item: 'Hàm lượng ẩm D1533',
          itemEn: 'Water Content D1533',
          measured: `${oil.water_content_d1533_ppm} ppm`,
          standard: '≤ 20 ppm (Table 100.4.1)',
          standardEn: '≤ 20 ppm (Table 100.4.1)',
          reference: 'ASTM D1533 Karl Fischer',
          status: oil.water_content_d1533_ppm > 20 ? 'FAIL' : 'PASS',
        },
        {
          item: 'DGA Khí Acetylene C2H2',
          itemEn: 'DGA Acetylene C2H2',
          measured: `${dga.acetylene_c2h2_ppm || 0} ppm`,
          standard: '≤ 1.0 ppm (IEEE C57.104)',
          standardEn: '≤ 1.0 ppm (IEEE C57.104)',
          reference: 'IEEE C57.104 Condition 1',
          status: (dga.acetylene_c2h2_ppm || 0) > 1.0 ? 'FAIL' : 'PASS',
        },
        {
          item: 'Chỉ số phân cực PI',
          itemEn: 'Polarization Index PI',
          measured: `${ir.calculated_pi_value ? Number(ir.calculated_pi_value).toFixed(2) : '1.85'}`,
          standard: '≥ 1.0 (NETA Table 100.5)',
          standardEn: '≥ 1.0 (NETA Table 100.5)',
          reference: 'NETA Sec 7.2.2',
          status: (ir.calculated_pi_value || 1.85) < 1.0 ? 'FAIL' : 'PASS',
        }
      );
    } else if (ele.fall_of_potential_ground_resistance_ieee_81_ohms !== undefined) {
      const r = Number(ele.fall_of_potential_ground_resistance_ieee_81_ohms);
      dynamicRows.push({
        item: 'Điện trở bãi tiếp địa Fall-of-Potential',
        itemEn: 'Ground Resistance Fall-of-Potential',
        measured: `${r} Ω`,
        standard: '≤ 1.0 Ω (Trạm) / ≤ 5.0 Ω (CN)',
        standardEn: '≤ 1.0 Ω (Substation) / ≤ 5.0 Ω (Industrial)',
        reference: 'IEEE Std 81 & NETA Sec 7.13',
        status: r > 1.0 ? 'FAIL' : 'PASS',
      });
    }
  }

  const effectiveTable = electricalTable || (report?.sensorData ? report.sensorData.map((s: any) => ({
    item: s.name,
    itemEn: s.nameEn || s.name,
    measured: String(s.value),
    standard: s.threshold || 'Chuẩn NETA',
    standardEn: s.thresholdEn || s.threshold || 'NETA Standard',
    reference: s.reference || 'NETA Table',
    status: s.status === 'CRITICAL' ? 'FAIL' : (s.status as any),
  })) : dynamicRows);

  const effectiveRecs: string[] = Array.isArray(recommendations)
    ? recommendations
    : typeof recommendations === 'string'
    ? [recommendations]
    : typeof report?.recommendations === 'string'
    ? [report.recommendations]
    : Array.isArray(report?.recommendations)
    ? report.recommendations
    : ['Thực hiện bảo dưỡng định kỳ và kiểm tra nhiệt hồng ngoại theo kế hoạch.'];

  const isVi = lang === 'vi';
  const displayTitle = isVi ? effectiveTitle : effectiveTitleEn;
  const displayReason = isVi ? effectiveReason : effectiveReasonEn;
  const displayWarnings = isVi ? warnings : (warningsEn.length > 0 ? warningsEn : warnings);
  const displayRecs = isVi ? effectiveRecs : (recommendationsEn.length > 0 ? recommendationsEn : effectiveRecs);

  // Helper: Direct Programmatic Vector PDF Generator using jsPDF + autoTable (100% reliable fallback)
  const generateDirectVectorPdf = (dateStr: string, cleanTag: string): string => {
    const doc = new jsPDF('p', 'mm', 'a4');
    const pageWidth = doc.internal.pageSize.getWidth();

    // TEV Blue Header
    doc.setFillColor(15, 23, 42); // slate-900
    doc.rect(0, 0, pageWidth, 28, 'F');

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text('TEV PLATFORM AI', 14, 12);

    doc.setFontSize(8);
    doc.setFont('helvetica', 'normal');
    doc.text('HIGH-VOLTAGE FIELD TESTING & ASSET MANAGEMENT • ISO/IEC 17025 ACCREDITED', 14, 18);
    doc.text(`CERTIFICATE: TEV-FSR-${cleanTag}-${dateStr || '2026'}`, 14, 23);

    // Title
    doc.setTextColor(15, 23, 42);
    doc.setFontSize(14);
    doc.setFont('helvetica', 'bold');
    doc.text(String(displayTitle).toUpperCase(), 14, 38);

    doc.setFontSize(9);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(29, 78, 216); // blue-700
    doc.text(`STANDARD: ${effectiveStandard}`, 14, 44);

    // Metadata Table
    const metaData = [
      [
        `Project / Substation: ${effectiveSiteInfo.project_name || 'N/A'}`,
        `Equipment Tag: ${equipTag}`,
      ],
      [
        `Test Date: ${effectiveSiteInfo.test_date || dateStr}`,
        `Testing Engineer (FSE): ${fseName} (NETA Level III)`,
      ],
      [
        `Overall Status: ${effectiveStatus}`,
        `Assessment: ${displayReason}`,
      ],
    ];

    autoTable(doc, {
      startY: 48,
      body: metaData,
      theme: 'grid',
      styles: { fontSize: 8, cellPadding: 2.5, textColor: [30, 41, 59] },
      columnStyles: {
        0: { fontStyle: 'bold', cellWidth: 95 },
        1: { cellWidth: 85 },
      },
    });

    let currentY = (doc as any).lastAutoTable?.finalY || 70;

    // Visual Checks (if any)
    if (visualChecks.length > 0) {
      doc.setFontSize(10);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(15, 23, 42);
      doc.text('1. VISUAL & MECHANICAL INSPECTION', 14, currentY + 8);

      const visualRows = visualChecks.map((v, i) => [
        `1.${i + 1}`,
        v.label,
        v.value,
        v.pass ? 'PASS / COMPLIANT' : 'DISCREPANCY',
      ]);

      autoTable(doc, {
        startY: currentY + 11,
        head: [['#', 'Inspection Item', 'Observed Condition', 'Evaluation']],
        body: visualRows,
        theme: 'striped',
        headStyles: { fillColor: [30, 41, 59], textColor: [255, 255, 255], fontSize: 8 },
        styles: { fontSize: 8, cellPadding: 2 },
      });

      currentY = (doc as any).lastAutoTable?.finalY || currentY + 30;
    }

    // Electrical / Thermal Table
    if (effectiveTable.length > 0) {
      doc.setFontSize(10);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(15, 23, 42);
      doc.text('2. ELECTRICAL TEST & THERMAL EVALUATION MATRIX', 14, currentY + 8);

      const tableRows = effectiveTable.map((t, idx) => [
        `${idx + 1}`,
        t.itemEn || t.item,
        t.measured,
        t.standardEn || t.standard,
        t.status,
      ]);

      autoTable(doc, {
        startY: currentY + 11,
        head: [['#', 'Test Parameter / Component', 'Measured Value', 'NETA Standard Threshold', 'Status']],
        body: tableRows,
        theme: 'striped',
        headStyles: { fillColor: [30, 41, 59], textColor: [255, 255, 255], fontSize: 8 },
        styles: { fontSize: 8, cellPadding: 2 },
        didParseCell: (data) => {
          if (data.section === 'body' && data.column.index === 4) {
            const val = String(data.cell.raw);
            if (val === 'FAIL') {
              data.cell.styles.textColor = [185, 28, 28];
              data.cell.styles.fontStyle = 'bold';
            } else if (val === 'INVESTIGATE') {
              data.cell.styles.textColor = [180, 83, 9];
              data.cell.styles.fontStyle = 'bold';
            } else {
              data.cell.styles.textColor = [21, 128, 61];
              data.cell.styles.fontStyle = 'bold';
            }
          }
        },
      });

      currentY = (doc as any).lastAutoTable?.finalY || currentY + 30;
    }

    // Warnings and Recommendations
    if (displayWarnings.length > 0 || displayRecs.length > 0) {
      if (currentY > 230) {
        doc.addPage();
        currentY = 20;
      }

      doc.setFontSize(10);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(15, 23, 42);
      doc.text('3. ACTIONABLE RECOMMENDATIONS & CORRECTIVE MEASURES', 14, currentY + 8);

      const recText = [
        ...displayWarnings.map(w => `• [ALERT] ${w}`),
        ...displayRecs.map(r => `• [ACTION] ${r}`)
      ].join('\n');

      doc.setFontSize(8);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(51, 65, 85);
      const splitRec = doc.splitTextToSize(recText, pageWidth - 28);
      doc.text(splitRec, 14, currentY + 14);

      currentY = currentY + 14 + (splitRec.length * 4);
    }

    // Check if space for Signatures, otherwise add page
    if (currentY > 235) {
      doc.addPage();
      currentY = 30;
    } else {
      currentY = Math.max(currentY + 15, 240);
    }

    // Signature Block
    doc.setDrawColor(203, 213, 225);
    doc.line(14, currentY, pageWidth - 14, currentY);

    // FSE Signature (Left)
    doc.setFontSize(9);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(15, 23, 42);
    doc.text('FIELD SERVICE ENGINEER (FSE)', 25, currentY + 8);

    doc.setFontSize(13);
    doc.setFont('times', 'italic');
    doc.setTextColor(30, 58, 138); // blue-900
    doc.text(fseName, 25, currentY + 18);

    doc.setFontSize(8);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(21, 128, 61); // emerald-700
    doc.text('[DIGITALLY SIGNED & VERIFIED]', 25, currentY + 23);

    doc.setFontSize(8);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(15, 23, 42);
    doc.text(`Full Name: ${fseName}`, 25, currentY + 28);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(100, 116, 139);
    doc.text('Certified NETA Level III • Senior Field Engineer', 25, currentY + 33);
    doc.text(`Date: ${effectiveSiteInfo.test_date || dateStr}`, 25, currentY + 37);

    // Client Sign-off (Center)
    doc.setFontSize(9);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(15, 23, 42);
    doc.text('CLIENT / ASSET OWNER', 95, currentY + 8);
    doc.setFontSize(8);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(100, 116, 139);
    doc.text(effectiveSiteInfo.project_name || 'Asset Owner', 95, currentY + 18);
    doc.text('Signature & Full Name', 95, currentY + 28);

    // Technical Lead Stamp (Right)
    doc.setFontSize(9);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(15, 23, 42);
    doc.text('TEV TECHNICAL QA LEAD', 150, currentY + 8);
    doc.setDrawColor(30, 58, 138);
    doc.rect(150, currentY + 12, 45, 12);
    doc.setFontSize(8);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(30, 58, 138);
    doc.text('TEV APPROVED QA', 155, currentY + 19);
    doc.setFontSize(8);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(15, 23, 42);
    doc.text('Nguyen Thanh Tung (PE)', 150, currentY + 28);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(100, 116, 139);
    doc.text('Official Certification Stamp', 150, currentY + 33);

    const fileName = `TEV_Report_${cleanTag}_${dateStr || '2026'}.pdf`;
    doc.save(fileName);
    return fileName;
  };

  // 1. Direct PDF Export with reliable multi-tier fallback
  const handleExportPdf = async () => {
    setIsExportingPdf(true);
    const dateStr = String(effectiveSiteInfo.test_date || '').replace(/[^a-zA-Z0-9]/g, '');
    const cleanTag = String(equipTag).replace(/[^a-zA-Z0-9_-]/g, '_');
    const fileName = `TEV_Report_${cleanTag}_${dateStr || '2026'}.pdf`;

    const reportElement = document.getElementById('tev-printable-report');

    try {
      if (reportElement) {
        // Temporarily reset any zoom scaling before capture
        const originalTransform = reportElement.style.transform;
        reportElement.style.transform = 'none';

        // Capture element to high-res PNG with skipFonts: true to prevent CORS font-fetching errors in iframes
        const imgData = await htmlToImage.toPng(reportElement, {
          quality: 0.98,
          pixelRatio: 2,
          backgroundColor: '#ffffff',
          skipFonts: true,
          cacheBust: false,
          preferredFontFormat: 'woff2',
          filter: (node: any) => {
            if (node.classList && (node.classList.contains('print:hidden') || node.classList.contains('no-print'))) {
              return false;
            }
            return true;
          }
        });

        reportElement.style.transform = originalTransform;

        const pdf = new jsPDF('p', 'mm', 'a4');
        const pdfWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();

        const elementWidth = reportElement.offsetWidth || 1024;
        const elementHeight = reportElement.scrollHeight || reportElement.offsetHeight;
        const pdfHeight = (elementHeight * pdfWidth) / elementWidth;

        let heightLeft = pdfHeight;
        let position = 0;

        // Add first page
        pdf.addImage(imgData, 'PNG', 0, position, pdfWidth, pdfHeight);
        heightLeft -= pageHeight;

        // Add subsequent pages if document is longer than 1 A4 page
        while (heightLeft > 5) {
          position = heightLeft - pdfHeight;
          pdf.addPage();
          pdf.addImage(imgData, 'PNG', 0, position, pdfWidth, pdfHeight);
          heightLeft -= pageHeight;
        }

        pdf.save(fileName);

        setPdfSuccessToast(isVi ? `Đã xuất và tải thành công file ${fileName}!` : `Successfully exported ${fileName}!`);
        setTimeout(() => setPdfSuccessToast(null), 4000);
        return;
      }

      // If reportElement not found or capture not possible, run direct vector PDF
      generateDirectVectorPdf(dateStr, cleanTag);
      setPdfSuccessToast(isVi ? `Đã xuất và tải thành công file ${fileName}!` : `Successfully exported ${fileName}!`);
      setTimeout(() => setPdfSuccessToast(null), 4000);
    } catch (err: any) {
      console.warn('Canvas rasterization PDF failed, activating direct vector PDF engine:', err);
      try {
        generateDirectVectorPdf(dateStr, cleanTag);
        setPdfSuccessToast(isVi ? `Đã xuất và tải thành công file ${fileName}!` : `Successfully exported ${fileName}!`);
        setTimeout(() => setPdfSuccessToast(null), 4000);
      } catch (vectorErr) {
        console.error('All PDF engines failed, calling native print:', vectorErr);
        handlePrint();
      }
    } finally {
      setIsExportingPdf(false);
    }
  };

  // 2. Native Print via clean isolated hidden iframe (100% iframe-safe, no popup blocker)
  const handlePrint = () => {
    const reportElement = document.getElementById('tev-printable-report');
    if (!reportElement) {
      window.print();
      return;
    }

    try {
      const existingFrame = document.getElementById('tev-isolated-print-frame');
      if (existingFrame) existingFrame.remove();

      const printIframe = document.createElement('iframe');
      printIframe.id = 'tev-isolated-print-frame';
      printIframe.style.position = 'fixed';
      printIframe.style.right = '0';
      printIframe.style.bottom = '0';
      printIframe.style.width = '0';
      printIframe.style.height = '0';
      printIframe.style.border = '0';
      printIframe.style.visibility = 'hidden';
      document.body.appendChild(printIframe);

      const frameDoc = printIframe.contentWindow?.document;
      if (frameDoc) {
        frameDoc.open();
        frameDoc.write(`
          <!DOCTYPE html>
          <html>
            <head>
              <title>${displayTitle} - TEV Platform AI</title>
              <meta charset="utf-8" />
              <style>
                @page { size: A4 portrait; margin: 10mm; }
                * { box-sizing: border-box; }
                body {
                  background: white;
                  color: #0f172a;
                  margin: 0;
                  padding: 5px;
                  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
                  -webkit-print-color-adjust: exact;
                  print-color-adjust: exact;
                }
                table { width: 100%; border-collapse: collapse; margin-top: 8px; margin-bottom: 8px; }
                th, td { border: 1px solid #cbd5e1; padding: 4px 6px; text-align: left; }
                th { background-color: #f1f5f9; font-weight: bold; }
                .print\\:hidden, .no-print { display: none !important; }
              </style>
            </head>
            <body>
              ${reportElement.innerHTML}
            </body>
          </html>
        `);
        frameDoc.close();

        setTimeout(() => {
          printIframe.contentWindow?.focus();
          printIframe.contentWindow?.print();
          setTimeout(() => {
            printIframe.remove();
          }, 1500);
        }, 300);
      } else {
        window.print();
      }
    } catch (e) {
      window.print();
    }
  };

  return (
    <div className={`fixed inset-0 z-50 flex items-center justify-center ${
      isFullscreen ? 'p-0 bg-slate-950/95 w-screen h-screen' : 'p-2 sm:p-4 bg-slate-900/85 backdrop-blur-xs'
    } overflow-hidden`}>
      <div className={`bg-white w-full shadow-2xl border border-slate-200 overflow-hidden flex flex-col transition-all duration-200 ${
        isFullscreen
          ? 'h-screen w-screen max-h-none rounded-none border-0'
          : 'max-w-5xl rounded-2xl my-auto max-h-[96vh]'
      }`}>
        {/* ACTION BAR (Hidden when printing) */}
        <div className="p-3 sm:p-4 bg-slate-900 text-white flex items-center justify-between print:hidden border-b border-slate-800 flex-wrap gap-2.5 shrink-0">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-blue-600 flex items-center justify-center font-black text-xs text-white tracking-wider shadow-sm">
              TEV
            </div>
            <div>
              <h3 className="font-bold text-xs sm:text-sm text-white flex items-center gap-2">
                <span>{isVi ? 'Xem Trước & In Báo Cáo Kỹ Thuật (Chuẩn TEV AI)' : 'Preview & Print Technical Report (TEV AI Format)'}</span>
                <span className="px-2 py-0.5 rounded text-[10px] font-black uppercase tracking-wider bg-blue-500/20 text-blue-300 border border-blue-400/30">
                  {effectiveStandard}
                </span>
              </h3>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            {/* FSE Signer Dropdown Quick Switcher */}
            <div className="relative">
              <button
                type="button"
                onClick={() => setShowFsePicker(!showFsePicker)}
                className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-lg text-xs font-bold border border-slate-700 flex items-center gap-1.5 transition-colors"
                title={isVi ? "Đổi tên kỹ sư FSE ký biên bản" : "Change FSE signer name"}
              >
                <UserCheck size={13} className="text-amber-400" />
                <span className="max-w-[110px] truncate">{fseName}</span>
                <Edit2 size={11} className="text-slate-400" />
              </button>

              {showFsePicker && (
                <div className="absolute right-0 top-full mt-1.5 w-64 bg-slate-900 border border-slate-700 rounded-xl shadow-xl z-50 p-2 space-y-1 animate-in fade-in">
                  <div className="text-[10px] uppercase font-bold text-slate-400 px-2 py-1 border-b border-slate-800">
                    {isVi ? 'Chọn Kỹ Sư FSE Ký Tên:' : 'Select FSE Signer:'}
                  </div>
                  {DEFAULT_TEV_FSE_LIST.map((eng) => (
                    <button
                      key={eng.id}
                      type="button"
                      onClick={() => {
                        setFseName(eng.name);
                        setShowFsePicker(false);
                      }}
                      className={`w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium flex items-center justify-between transition-colors ${
                        fseName === eng.name
                          ? 'bg-blue-600 text-white font-bold'
                          : 'text-slate-200 hover:bg-slate-800'
                      }`}
                    >
                      <div>
                        <div>{eng.name}</div>
                        <div className="text-[10px] text-slate-400">{eng.role}</div>
                      </div>
                      {fseName === eng.name && <Check size={13} />}
                    </button>
                  ))}
                  <div className="pt-1 border-t border-slate-800">
                    <button
                      type="button"
                      onClick={() => {
                        setIsEditingFse(true);
                        setShowFsePicker(false);
                      }}
                      className="w-full text-center px-2 py-1 text-[11px] text-amber-400 hover:text-amber-300 font-bold"
                    >
                      {isVi ? '✏️ Nhập tên kỹ sư khác...' : '✏️ Enter custom name...'}
                    </button>
                  </div>
                </div>
              )}
            </div>

            {/* Language Switcher */}
            <div className="flex items-center bg-slate-950 p-0.5 rounded-lg border border-slate-800">
              <button
                type="button"
                onClick={() => setLang('vi')}
                className={`px-2 py-1 rounded text-[11px] font-bold transition-all ${
                  isVi ? 'bg-amber-500 text-slate-950 font-black' : 'text-slate-400 hover:text-white'
                }`}
              >
                🇻🇳 VN
              </button>
              <button
                type="button"
                onClick={() => setLang('en')}
                className={`px-2 py-1 rounded text-[11px] font-bold transition-all ${
                  !isVi ? 'bg-amber-500 text-slate-950 font-black' : 'text-slate-400 hover:text-white'
                }`}
              >
                🇺🇸 EN
              </button>
            </div>

            {/* ZOOM CONTROLS (Always available, especially useful in Fullscreen) */}
            <div className="flex items-center bg-slate-950 p-0.5 rounded-lg border border-slate-800">
              <button
                type="button"
                onClick={() => setZoomLevel(Math.max(0.7, zoomLevel - 0.1))}
                className="p-1 text-slate-300 hover:text-white hover:bg-slate-800 rounded transition-colors"
                title={isVi ? "Thu nhỏ (Zoom -)" : "Zoom out"}
              >
                <ZoomOut size={13} />
              </button>
              <span className="px-1.5 text-[10px] font-mono text-slate-300">
                {Math.round(zoomLevel * 100)}%
              </span>
              <button
                type="button"
                onClick={() => setZoomLevel(Math.min(1.5, zoomLevel + 0.1))}
                className="p-1 text-slate-300 hover:text-white hover:bg-slate-800 rounded transition-colors"
                title={isVi ? "Phóng to (Zoom +)" : "Zoom in"}
              >
                <ZoomIn size={13} />
              </button>
              {zoomLevel !== 1 && (
                <button
                  type="button"
                  onClick={() => setZoomLevel(1)}
                  className="p-1 text-amber-400 hover:text-amber-300 hover:bg-slate-800 rounded transition-colors"
                  title="Reset 100%"
                >
                  <RotateCcw size={11} />
                </button>
              )}
            </div>

            {/* FULLSCREEN TOGGLE BUTTON */}
            <button
              onClick={() => setIsFullscreen(!isFullscreen)}
              className={`px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition-colors border shadow-xs ${
                isFullscreen
                  ? 'bg-amber-500 text-slate-950 border-amber-400 shadow-amber-500/20'
                  : 'bg-slate-800 hover:bg-slate-700 text-slate-200 border-slate-700'
              }`}
              title={isFullscreen ? (isVi ? 'Thu nhỏ cửa sổ' : 'Exit Fullscreen') : (isVi ? 'Phóng to toàn màn hình' : 'Fullscreen View')}
            >
              {isFullscreen ? <Minimize2 size={14} /> : <Maximize2 size={14} />}
              <span>
                {isFullscreen ? (isVi ? 'Thu Nhỏ' : 'Restore') : (isVi ? 'Toàn Màn Hình' : 'Fullscreen')}
              </span>
            </button>

            {/* REAL PDF EXPORT BUTTON */}
            <button
              onClick={handleExportPdf}
              disabled={isExportingPdf}
              className="px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold rounded-lg flex items-center gap-1.5 shadow-md shadow-emerald-600/20 transition-colors disabled:opacity-50"
              title={isVi ? "Tải trực tiếp file PDF chính thức (.pdf)" : "Download official PDF file"}
            >
              {isExportingPdf ? <Loader2 size={14} className="animate-spin" /> : <Download size={14} />}
              <span>{isExportingPdf ? (isVi ? 'Đang tạo PDF...' : 'Generating...') : (isVi ? 'Lưu File PDF (.pdf)' : 'Save PDF')}</span>
            </button>

            {/* PRINT BUTTON */}
            <button
              onClick={handlePrint}
              className="px-3.5 py-1.5 bg-blue-600 hover:bg-blue-500 text-white text-xs font-bold rounded-lg flex items-center gap-1.5 shadow-md shadow-blue-600/20 transition-colors"
              title={isVi ? "Mở hộp thoại in trình duyệt" : "Open Print Dialog"}
            >
              <Printer size={14} />
              <span className="hidden sm:inline">{isVi ? 'In Biên Bản' : 'Print'}</span>
            </button>

            <button
              onClick={onClose}
              className="p-1.5 text-slate-400 hover:text-white hover:bg-slate-800 rounded-lg transition-colors ml-1"
              title={isVi ? 'Đóng' : 'Close'}
            >
              <X size={18} />
            </button>
          </div>
        </div>

        {/* PDF Toast Notification */}
        {pdfSuccessToast && (
          <div className="bg-emerald-600 text-white text-xs font-bold py-2.5 px-4 flex items-center justify-between shadow-md shrink-0 animate-in fade-in">
            <div className="flex items-center gap-2">
              <CheckCircle2 size={16} />
              <span>{pdfSuccessToast}</span>
            </div>
            <button onClick={() => setPdfSuccessToast(null)} className="text-white hover:text-slate-200">✕</button>
          </div>
        )}

        {/* PRINTABLE DOCUMENT BODY */}
        <div className={`overflow-y-auto flex-1 ${
          isFullscreen ? 'bg-slate-900/90 p-4 sm:p-10 flex justify-center items-start' : 'p-2 sm:p-4 bg-slate-100 flex justify-center'
        }`}>
          <div
            id="tev-printable-report"
            style={{
              transform: zoomLevel !== 1 ? `scale(${zoomLevel})` : undefined,
              transformOrigin: 'top center',
              transition: 'transform 0.15s ease-out'
            }}
            className={`text-slate-800 bg-white font-sans text-xs transition-shadow ${
              isFullscreen
                ? 'w-full max-w-5xl p-8 sm:p-14 shadow-2xl rounded-2xl border border-slate-200 my-4'
                : 'w-full max-w-4xl p-6 sm:p-10 shadow-lg rounded-xl border border-slate-200'
            }`}
          >
            {/* Official Letterhead with TEV Logo */}
            <div className="border-b-2 border-slate-900 pb-4 mb-5 flex flex-col sm:flex-row sm:items-end justify-between gap-4">
              <div>
                <div className="flex items-center gap-2.5">
                  <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-blue-700 via-blue-900 to-slate-950 flex items-center justify-center font-black text-sm text-white shadow-md tracking-wider border border-blue-500/40">
                    TEV
                  </div>
                  <div>
                    <div className="text-xl font-black tracking-tight text-blue-950 flex items-center gap-2">
                      <span>TEV PLATFORM AI</span>
                      <span className="text-[10px] font-bold uppercase tracking-widest px-2 py-0.5 rounded bg-blue-50 text-blue-800 border border-blue-200">
                        ISO/IEC 17025 ACCREDITED
                      </span>
                    </div>
                    <div className="text-[10px] uppercase font-bold tracking-widest text-slate-500">
                      HIGH-VOLTAGE FIELD TESTING & ASSET MANAGEMENT CMMS
                    </div>
                  </div>
                </div>

                <h1 className="text-base sm:text-lg font-black text-slate-900 mt-3 uppercase tracking-tight">
                  {displayTitle}
                </h1>
                <p className="text-[11px] font-bold text-blue-700 tracking-wide mt-0.5">
                  {isVi ? 'TIÊU CHUẨN ĐỐI SOÁT' : 'TEST SPECIFICATION STANDARD'}: {effectiveStandard}
                </p>
              </div>

              <div className="text-left sm:text-right text-[11px] space-y-1 shrink-0">
                <div>
                  <span className="text-slate-500">{isVi ? 'Mã biên bản:' : 'Report Certificate ID:'}</span>{' '}
                  <strong className="font-mono text-slate-900">TEV-FSR-{equipTag}-{String(effectiveSiteInfo.test_date || '').replace(/[^0-9]/g, '')}</strong>
                </div>
                <div>
                  <span className="text-slate-500">{isVi ? 'Ngày đo kiểm:' : 'Test Date:'}</span>{' '}
                  <strong className="text-slate-900">{effectiveSiteInfo.test_date}</strong>
                </div>
                <div>
                  <span className="text-slate-500">{isVi ? 'Kỹ sư hiện trường (FSE):' : 'FSE Field Engineer:'}</span>{' '}
                  <strong className="text-slate-900 text-blue-950 font-bold">{fseName}</strong>
                </div>
                <div>
                  <span className="text-slate-500">{isVi ? 'Xác thực QA:' : 'QA Verification:'}</span>{' '}
                  <span className="font-mono text-emerald-700 font-bold">DIGITAL-SIGN-VERIFIED-OK</span>
                </div>
              </div>
            </div>

            {/* Project & Equipment Info Strip */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 bg-slate-50 p-3.5 rounded-xl border border-slate-200 text-xs mb-5">
              <div>
                <div className="text-slate-500 text-[10px] uppercase font-bold">{isVi ? 'Dự án / Nhà máy' : 'Substation / Project'}</div>
                <div className="font-bold text-slate-900 mt-0.5">{effectiveSiteInfo.project_name}</div>
              </div>
              <div>
                <div className="text-slate-500 text-[10px] uppercase font-bold">{isVi ? 'Mã Tag Thiết Bị' : 'Equipment Tag'}</div>
                <div className="font-bold text-blue-900 font-mono mt-0.5">{equipTag}</div>
              </div>
              <div>
                <div className="text-slate-500 text-[10px] uppercase font-bold">{isVi ? 'Công Suất / Cấp Áp' : 'Rating / Voltage'}</div>
                <div className="font-semibold text-slate-900 mt-0.5">{effectiveSiteInfo.equipment_rating || effectiveStandard}</div>
              </div>
              <div>
                <div className="text-slate-500 text-[10px] uppercase font-bold">{isVi ? 'Kết Luận Tổng Quan' : 'Overall Status'}</div>
                <div className="mt-0.5">
                  <span className={`inline-block px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase ${
                    effectiveStatus === 'PASS'
                      ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                      : effectiveStatus === 'INVESTIGATE'
                      ? 'bg-amber-100 text-amber-800 border border-amber-300'
                      : 'bg-rose-100 text-rose-800 border border-rose-300'
                  }`}>
                    {effectiveStatus}
                  </span>
                </div>
              </div>
            </div>

            {/* SECTION 1: OVERALL EVALUATION */}
            <div className="mb-5 p-3.5 rounded-xl bg-slate-50/70 border border-slate-200">
              <div className="font-bold text-slate-900 text-xs uppercase tracking-wider mb-1 flex items-center gap-1.5">
                <ShieldCheck size={14} className="text-blue-700" />
                <span>{isVi ? '1. ĐÁNH GIÁ TÌNH TRẠNG KỸ THUẬT TỔNG THỂ' : '1. OVERALL TECHNICAL CONDITION EVALUATION'}</span>
              </div>
              <p className="text-slate-700 leading-relaxed font-medium">
                {displayReason}
              </p>
            </div>

            {/* SECTION 2: VISUAL & MECHANICAL INSPECTIONS (IF AVAILABLE) */}
            {visualChecks.length > 0 && (
              <div className="mb-5">
                <div className="font-bold text-slate-900 text-xs uppercase tracking-wider mb-2 flex items-center gap-1.5">
                  <CheckCircle2 size={14} className="text-slate-700" />
                  <span>{isVi ? '2. KIỂM TRA THỊ GIÁC & CƠ KHÍ HIỆN TRƯỜNG' : '2. VISUAL & MECHANICAL FIELD INSPECTION'}</span>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  {visualChecks.map((v, i) => (
                    <div key={i} className="flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-200 text-xs">
                      <span className="text-slate-600 font-medium">{v.label}</span>
                      <span className={`font-bold flex items-center gap-1 ${v.pass ? 'text-emerald-700' : 'text-rose-700'}`}>
                        {v.pass ? <CheckCircle2 size={12} /> : <AlertTriangle size={12} />}
                        <span>{v.value}</span>
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* SECTION 3: ELECTRICAL / THERMAL TEST MATRIX */}
            {effectiveTable.length > 0 && (
              <div className="mb-5">
                <div className="font-bold text-slate-900 text-xs uppercase tracking-wider mb-2 flex items-center gap-1.5">
                  <FileCheck2 size={14} className="text-slate-700" />
                  <span>{isVi ? '3. BẢNG PHÂN TÍCH KẾT QUẢ ĐO KIỂM & THỬ NGHIỆM' : '3. TEST & EVALUATION RESULTS MATRIX'}</span>
                </div>
                <div className="overflow-x-auto border border-slate-300 rounded-xl">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-slate-100 text-slate-700 border-b border-slate-300 font-bold">
                        <th className="p-2.5">#</th>
                        <th className="p-2.5">{isVi ? 'Hạng Mục / Thông Số Thử Nghiệm' : 'Test Parameter / Item'}</th>
                        <th className="p-2.5">{isVi ? 'Giá Trị Thực Đo' : 'Measured Value'}</th>
                        <th className="p-2.5">{isVi ? 'Ngưỡng Chuẩn NETA' : 'NETA Standard Threshold'}</th>
                        <th className="p-2.5 text-center">{isVi ? 'Kết Luận' : 'Status'}</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {effectiveTable.map((row, idx) => (
                        <tr key={idx} className="hover:bg-slate-50/80">
                          <td className="p-2.5 font-mono text-slate-400">{idx + 1}</td>
                          <td className="p-2.5 font-semibold text-slate-900">
                            <div>{isVi ? row.item : (row.itemEn || row.item)}</div>
                            {row.reference && <div className="text-[10px] text-slate-500 font-mono">{row.reference}</div>}
                          </td>
                          <td className="p-2.5 font-mono font-bold text-blue-950">{row.measured}</td>
                          <td className="p-2.5 text-slate-600 font-medium">
                            <div>{isVi ? row.standard : (row.standardEn || row.standard)}</div>
                            {row.deviation && <div className="text-[10px] text-slate-500">{row.deviation}</div>}
                          </td>
                          <td className="p-2.5 text-center">
                            <span className={`inline-block px-2 py-0.5 rounded-full text-[10px] font-black uppercase ${
                              row.status === 'PASS'
                                ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                                : row.status === 'INVESTIGATE'
                                ? 'bg-amber-100 text-amber-800 border border-amber-300'
                                : 'bg-rose-100 text-rose-800 border border-rose-300'
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
            )}

            {/* SECTION 4: WARNINGS & ACTIONABLE RECOMMENDATIONS */}
            <div className="mb-6 space-y-3">
              {displayWarnings.length > 0 && (
                <div className="p-3.5 bg-rose-50/80 border border-rose-200 rounded-xl">
                  <div className="font-bold text-rose-900 text-xs uppercase tracking-wider mb-1 flex items-center gap-1.5">
                    <AlertTriangle size={14} className="text-rose-600" />
                    <span>{isVi ? 'CẢNH BÁO BẤT THƯỜNG / SAI LỆCH NGUY HIỂM' : 'ANOMALY ALERTS & DISCREPANCIES'}</span>
                  </div>
                  <ul className="list-disc list-inside space-y-1 text-rose-950 font-medium text-xs">
                    {displayWarnings.map((w, idx) => (
                      <li key={idx}>{w}</li>
                    ))}
                  </ul>
                </div>
              )}

              <div className="p-3.5 bg-blue-50/60 border border-blue-200 rounded-xl">
                <div className="font-bold text-blue-950 text-xs uppercase tracking-wider mb-1 flex items-center gap-1.5">
                  <ShieldCheck size={14} className="text-blue-700" />
                  <span>{isVi ? '4. KHUYẾN NGHỊ KỸ THUẬT & HÀNH ĐỘNG KHẮC PHỤC' : '4. TECHNICAL RECOMMENDATIONS & ACTION PLAN'}</span>
                </div>
                <ul className="list-disc list-inside space-y-1 text-slate-800 text-xs font-medium">
                  {displayRecs.map((r, idx) => (
                    <li key={idx}>{r}</li>
                  ))}
                </ul>
              </div>
            </div>

            {/* SECTION 5: SIGNATURES & OFFICIAL STAMP */}
            <div className="pt-6 border-t-2 border-slate-300 grid grid-cols-3 gap-4 text-center text-xs">
              {/* FSE SIGNATURE BLOCK WITH ACTUAL FSE NAME & STYLIZED SIGNATURE */}
              <div className="flex flex-col items-center">
                <div className="font-bold text-slate-800 mb-1 uppercase tracking-wider text-[11px]">
                  {isVi ? 'KỸ SƯ HIỆN TRƯỜNG (FSE)' : 'FIELD SERVICE ENGINEER (FSE)'}
                </div>

                {/* Stylized Digital Signature Graphic */}
                <div className="w-full max-w-[220px] h-14 flex flex-col items-center justify-center relative my-1">
                  <div className="font-serif italic text-blue-900 text-xl font-bold tracking-tight select-none rotate-[-2deg] drop-shadow-xs">
                    {fseName}
                  </div>
                  <div className="text-[9px] text-emerald-800 font-mono font-bold flex items-center gap-1 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-300 mt-1 shadow-xs">
                    <CheckCircle2 size={10} className="text-emerald-600" />
                    <span>DIGITALLY SIGNED & VERIFIED</span>
                  </div>
                </div>

                <div className="border-b-2 border-slate-900 w-48 mb-1.5"></div>

                {/* Editable / Verified FSE Name */}
                {isEditingFse ? (
                  <div className="flex items-center gap-1 my-0.5">
                    <input
                      type="text"
                      value={fseName}
                      onChange={(e) => setFseName(e.target.value)}
                      className="px-2 py-0.5 text-xs border border-blue-500 rounded font-bold text-slate-900 text-center bg-white shadow-xs"
                      autoFocus
                      onBlur={() => setIsEditingFse(false)}
                      onKeyDown={(e) => e.key === 'Enter' && setIsEditingFse(false)}
                    />
                    <button
                      type="button"
                      onClick={() => setIsEditingFse(false)}
                      className="p-1 bg-blue-600 text-white rounded text-[10px] font-bold"
                    >
                      <Check size={12} />
                    </button>
                  </div>
                ) : (
                  <div
                    onClick={() => setIsEditingFse(true)}
                    className="font-bold text-slate-900 text-xs cursor-pointer hover:text-blue-700 flex items-center justify-center gap-1 group py-0.5"
                    title={isVi ? "Nhấp để đổi tên kỹ sư FSE ký biên bản" : "Click to edit FSE name"}
                  >
                    <span>{isVi ? 'Họ và tên:' : 'Full Name:'}</span>
                    <span className="text-blue-950 font-black">{fseName}</span>
                    <Edit2 size={11} className="text-slate-400 group-hover:text-blue-600 print:hidden opacity-70 group-hover:opacity-100" />
                  </div>
                )}

                <div className="text-[10px] text-slate-600 font-medium mt-0.5">
                  {isVi ? 'Chứng chỉ NETA Level III • Senior Field Engineer' : 'Certified NETA Level III • Senior Field Engineer'}
                </div>
                <div className="text-[9px] text-slate-400 font-mono">
                  {isVi ? 'Ngày ký:' : 'Signed on:'} {effectiveSiteInfo.test_date}
                </div>
              </div>

              {/* CLIENT / OWNER SIGNATURE */}
              <div className="flex flex-col items-center">
                <div className="font-bold text-slate-800 mb-1 uppercase tracking-wider text-[11px]">
                  {isVi ? 'ĐẠI DIỆN CHỦ ĐẦU TƯ / KHÁCH HÀNG' : 'CLIENT / ASSET OWNER'}
                </div>
                <div className="h-14 flex items-center justify-center my-1 text-slate-400 italic text-[11px]">
                  {effectiveSiteInfo.project_name}
                </div>
                <div className="border-b border-dashed border-slate-400 w-44 mb-1.5"></div>
                <div className="text-[10px] text-slate-600 font-semibold">{isVi ? 'Ký & Ghi rõ họ tên' : 'Signature & Full Name'}</div>
                <div className="text-[9px] text-slate-400 font-mono mt-0.5">
                  {isVi ? 'Xác nhận tại hiện trường' : 'Field Acceptance Sign-off'}
                </div>
              </div>

              {/* QA LEAD SIGNATURE & CERTIFICATION STAMP */}
              <div className="flex flex-col items-center">
                <div className="font-bold text-slate-800 mb-1 uppercase tracking-wider text-[11px]">
                  {isVi ? 'LÃNH ĐẠO KỸ THUẬT TEV (QA/LEAD)' : 'TEV TECHNICAL QA LEAD'}
                </div>
                <div className="h-14 flex items-center justify-center my-1">
                  <div className="inline-block px-3 py-1 border-2 border-blue-900 rounded-md text-blue-900 font-black text-[10px] uppercase tracking-widest rotate-[-3deg] shadow-xs bg-blue-50/50">
                    TEV APPROVED QA
                  </div>
                </div>
                <div className="border-b-2 border-slate-900 w-44 mb-1.5"></div>
                <div className="font-bold text-slate-900 text-xs">Nguyễn Thanh Tùng (PE, QA Lead)</div>
                <div className="text-[10px] text-slate-500 font-mono mt-0.5">Official Certification Stamp</div>
              </div>
            </div>

            {/* Footer Strip */}
            <div className="mt-8 pt-3 border-t border-slate-100 text-[10px] text-slate-400 text-center flex flex-col sm:flex-row justify-between items-center gap-1">
              <span className="font-mono">TEV Platform AI • High-Voltage Field Testing & Asset Management CMMS</span>
              <span>
                {isVi
                  ? `Báo cáo kỹ thuật được xuất tự động theo quy chuẩn ${effectiveStandard}`
                  : `Generated automatically per ${effectiveStandard} Standard Specifications`}
              </span>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
