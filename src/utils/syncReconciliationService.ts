/**
 * Sync Reconciliation Service
 * Kiểm tra phiên bản dữ liệu (timestamp/version) giữa Cloud Firestore và Google Sheets (qua Sheets API)
 * Đảm bảo tính nhất quán (Consistency), ngăn chặn ghi đè dữ liệu mới bởi dữ liệu cũ (Conflict Prevention).
 */

export type ReconciliationStatus = 
  | 'in_sync'             // Dữ liệu và phiên bản 2 bên khớp nhau
  | 'firestore_newer'     // Firestore mới cập nhật hơn -> an toàn ghi đè lên Sheets
  | 'sheets_newer'        // Google Sheets mới cập nhật hơn -> cần bảo vệ, kéo về Firestore
  | 'conflict'            // Cả 2 bên đều có sửa đổi chênh lệch cần người dùng quyết định
  | 'only_in_firestore'   // Chỉ có ở Firestore -> cần đẩy lên Sheets
  | 'only_in_sheets';     // Chỉ có ở Sheets -> cần kéo về Firestore

export interface ReconciliationItem {
  id: string;
  name: string;
  entityType: 'workOrder' | 'customer' | 'equipment' | 'inventory';
  entityLabel: string;
  firestoreUpdatedAt?: string;
  firestoreTimestamp: number;
  sheetsUpdatedAt?: string;
  sheetsTimestamp: number;
  status: ReconciliationStatus;
  statusText: string;
  sheetName: string;
  sheetRowIndex?: number;
  diffSummary?: string;
  firestoreData?: any;
  sheetsData?: any;
  recommendedAction: 'push_to_sheets' | 'pull_to_firestore' | 'no_action' | 'manual_review';
}

export interface ReconciliationReport {
  timestamp: string;
  totalChecked: number;
  inSyncCount: number;
  firestoreNewerCount: number;
  sheetsNewerCount: number;
  conflictsCount: number;
  onlyInFirestoreCount: number;
  onlyInSheetsCount: number;
  items: ReconciliationItem[];
  hasConflicts: boolean;
  canSafeOverwriteSheets: boolean;
  canSafeOverwriteFirestore: boolean;
  summaryMessage: string;
}

/**
 * Phân tích timestamp linh hoạt từ nhiều định dạng:
 * ISO 8601, 'DD/MM/YYYY, HH:mm:ss', Firestore Timestamp object, unix ms/seconds
 */
export const parseVersionTimestamp = (val: any): number => {
  if (val === null || val === undefined || val === '') return 0;
  
  if (typeof val === 'number') {
    return val > 1e11 ? val : val * 1000;
  }
  
  if (typeof val?.toMillis === 'function') {
    return val.toMillis();
  }
  
  if (typeof val?.toDate === 'function') {
    return val.toDate().getTime();
  }
  
  if (typeof val?.seconds === 'number') {
    return val.seconds * 1000;
  }
  
  if (typeof val === 'string') {
    const s = val.trim();
    if (!s) return 0;
    
    // Thử định dạng chuẩn ISO
    const directParsed = Date.parse(s);
    if (!isNaN(directParsed) && !s.includes('/')) {
      return directParsed;
    }
    
    // Thử định dạng Việt Nam DD/MM/YYYY hoặc DD/MM/YYYY HH:mm:ss hoặc DD/MM/YYYY, HH:mm:ss
    const match = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[,\s]+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
    if (match) {
      const day = parseInt(match[1], 10);
      const month = parseInt(match[2], 10) - 1;
      const year = parseInt(match[3], 10);
      const hours = match[4] ? parseInt(match[4], 10) : 0;
      const minutes = match[5] ? parseInt(match[5], 10) : 0;
      const seconds = match[6] ? parseInt(match[6], 10) : 0;
      const d = new Date(year, month, day, hours, minutes, seconds);
      if (!isNaN(d.getTime())) return d.getTime();
    }
    
    if (!isNaN(directParsed)) return directParsed;
  }
  
  return 0;
};

/**
 * Định dạng timestamp hiển thị cho người dùng
 */
export const formatVersionDate = (timestamp: number | string | undefined): string => {
  if (!timestamp) return 'Chưa có mốc thời gian';
  const ms = typeof timestamp === 'number' ? timestamp : parseVersionTimestamp(timestamp);
  if (ms === 0) return 'Chưa có mốc thời gian';
  return new Date(ms).toLocaleString('vi-VN', {
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    day: '2-digit',
    month: '2-digit',
    year: 'numeric'
  });
};

/**
 * Đối soát phiên bản dữ liệu chi tiết giữa Firestore và các sheet lấy từ Google Sheets API
 */
export const compareVersionsBetweenFirestoreAndSheets = (params: {
  firestoreData: {
    workOrders: any[];
    customers: any[];
    equipment: any[];
    inventory: any[];
  };
  sheetsData: {
    allSheets: Record<string, any[][]>;
  };
}): ReconciliationReport => {
  const { firestoreData, sheetsData } = params;
  const items: ReconciliationItem[] = [];
  const allSheets = sheetsData.allSheets || {};
  const sheetNames = Object.keys(allSheets);

  // 1. Phân loại sheets từ Google Sheets API
  const getSheetRows = (keywords: string[]): { name: string; rows: any[][] } | null => {
    for (const name of sheetNames) {
      const lower = name.toLowerCase();
      if (keywords.some(k => lower.includes(k))) {
        return { name, rows: allSheets[name] || [] };
      }
    }
    return null;
  };

  const woSheet = getSheetRows(['tev service flatform', 'work order', 'cmms', 'phiếu công tác']);
  const customerSheet = getSheetRows(['khachhang', 'khách hàng', 'customer']);
  const equipmentSheet = getSheetRows(['thietbi', 'thiết bị', 'equipment']);
  const inventorySheet = getSheetRows(['quanlykho', 'kho', 'inventory', 'vật tư']);

  // --- 1. Đối soát Work Orders (TEV Service Flatform) ---
  if (woSheet && woSheet.rows.length > 0) {
    const rows = woSheet.rows;
    // Tìm hàng header
    let headerIdx = 0;
    for (let r = 0; r < Math.min(5, rows.length); r++) {
      if (rows[r]?.some((cell: any) => String(cell).toLowerCase().includes('mã wo') || String(cell).toLowerCase().includes('tiêu đề'))) {
        headerIdx = r;
        break;
      }
    }

    const dataRows = rows.slice(headerIdx + 1);
    const sheetWoMap = new Map<string, { rowIndex: number; row: any[]; timestamp: number; dateStr: string; title: string }>();

    dataRows.forEach((row, idx) => {
      if (!row || row.length === 0) return;
      const id = String(row[0] || '').trim();
      if (!id || id.toLowerCase().includes('mã wo')) return;
      
      // Ngày cập nhật cột 19 (index 18), nếu trống thì lấy Ngày tạo cột 18 (index 17)
      const updatedAtStr = String(row[18] || row[17] || '');
      const timestamp = parseVersionTimestamp(updatedAtStr);
      sheetWoMap.set(id.toLowerCase(), {
        rowIndex: headerIdx + 1 + idx + 1,
        row,
        timestamp,
        dateStr: updatedAtStr,
        title: String(row[2] || id)
      });
    });

    const processedWoIds = new Set<string>();

    // So sánh từ phía Firestore
    firestoreData.workOrders.forEach(wo => {
      const id = (wo.id || '').trim();
      if (!id) return;
      processedWoIds.add(id.toLowerCase());

      const fsUpdatedAt = wo.updatedAt || wo.createdAt || '';
      const fsTimestamp = parseVersionTimestamp(fsUpdatedAt);
      const sheetRecord = sheetWoMap.get(id.toLowerCase());

      if (!sheetRecord) {
        // Chỉ có ở Firestore
        items.push({
          id,
          name: wo.title || `Phiếu ${id}`,
          entityType: 'workOrder',
          entityLabel: 'Phiếu Công Tác (WO)',
          firestoreUpdatedAt: fsUpdatedAt,
          firestoreTimestamp: fsTimestamp,
          sheetsUpdatedAt: undefined,
          sheetsTimestamp: 0,
          status: 'only_in_firestore',
          statusText: 'Chỉ có trên Firestore',
          sheetName: woSheet.name,
          recommendedAction: 'push_to_sheets',
          firestoreData: wo
        });
      } else {
        const diffMs = fsTimestamp - sheetRecord.timestamp;
        const absDiff = Math.abs(diffMs);
        
        let status: ReconciliationStatus = 'in_sync';
        let statusText = 'Đồng bộ nhất quán';
        let recommendedAction: ReconciliationItem['recommendedAction'] = 'no_action';

        if (absDiff <= 3000) {
          status = 'in_sync';
          statusText = 'Đồng bộ khớp phiên bản';
          recommendedAction = 'no_action';
        } else if (diffMs > 3000) {
          status = 'firestore_newer';
          statusText = 'Firestore mới hơn';
          recommendedAction = 'push_to_sheets';
        } else {
          status = 'sheets_newer';
          statusText = 'Google Sheets mới hơn';
          recommendedAction = 'pull_to_firestore';
        }

        items.push({
          id,
          name: wo.title || sheetRecord.title || `Phiếu ${id}`,
          entityType: 'workOrder',
          entityLabel: 'Phiếu Công Tác (WO)',
          firestoreUpdatedAt: fsUpdatedAt,
          firestoreTimestamp: fsTimestamp,
          sheetsUpdatedAt: sheetRecord.dateStr,
          sheetsTimestamp: sheetRecord.timestamp,
          status,
          statusText,
          sheetName: woSheet.name,
          sheetRowIndex: sheetRecord.rowIndex,
          diffSummary: `Chênh lệch: ${Math.round(absDiff / 1000)}s`,
          firestoreData: wo,
          sheetsData: sheetRecord.row,
          recommendedAction
        });
      }
    });

    // Các bản ghi chỉ có trên Sheets
    sheetWoMap.forEach((sheetRecord, lowerId) => {
      if (!processedWoIds.has(lowerId)) {
        items.push({
          id: String(sheetRecord.row[0] || lowerId),
          name: sheetRecord.title,
          entityType: 'workOrder',
          entityLabel: 'Phiếu Công Tác (WO)',
          firestoreUpdatedAt: undefined,
          firestoreTimestamp: 0,
          sheetsUpdatedAt: sheetRecord.dateStr,
          sheetsTimestamp: sheetRecord.timestamp,
          status: 'only_in_sheets',
          statusText: 'Chỉ có trên Google Sheets',
          sheetName: woSheet.name,
          sheetRowIndex: sheetRecord.rowIndex,
          recommendedAction: 'pull_to_firestore',
          sheetsData: sheetRecord.row
        });
      }
    });
  }

  // --- 2. Đối soát Khách Hàng (khachhang) ---
  if (customerSheet && customerSheet.rows.length > 0) {
    const rows = customerSheet.rows;
    let headerIdx = 0;
    for (let r = 0; r < Math.min(5, rows.length); r++) {
      if (rows[r]?.some((cell: any) => String(cell).toLowerCase().includes('mã khách hàng') || String(cell).toLowerCase().includes('tên khách hàng'))) {
        headerIdx = r;
        break;
      }
    }

    const dataRows = rows.slice(headerIdx + 1);
    const sheetCustMap = new Map<string, { rowIndex: number; row: any[]; timestamp: number; dateStr: string; name: string }>();

    dataRows.forEach((row, idx) => {
      if (!row || row.length === 0) return;
      const id = String(row[0] || '').trim();
      if (!id || id.toLowerCase().includes('mã khách hàng')) return;
      const updatedAtStr = String(row[7] || row[6] || '');
      const timestamp = parseVersionTimestamp(updatedAtStr);
      sheetCustMap.set(id.toLowerCase(), {
        rowIndex: headerIdx + 1 + idx + 1,
        row,
        timestamp,
        dateStr: updatedAtStr,
        name: String(row[1] || id)
      });
    });

    const processedCustIds = new Set<string>();

    firestoreData.customers.forEach(cust => {
      const id = (cust.id || '').trim();
      if (!id) return;
      processedCustIds.add(id.toLowerCase());

      const fsUpdatedAt = cust.updatedAt || cust.createdAt || '';
      const fsTimestamp = parseVersionTimestamp(fsUpdatedAt);
      const sheetRecord = sheetCustMap.get(id.toLowerCase());

      if (!sheetRecord) {
        items.push({
          id,
          name: cust.name || `KH ${id}`,
          entityType: 'customer',
          entityLabel: 'Khách hàng',
          firestoreUpdatedAt: fsUpdatedAt,
          firestoreTimestamp: fsTimestamp,
          sheetsUpdatedAt: undefined,
          sheetsTimestamp: 0,
          status: 'only_in_firestore',
          statusText: 'Chỉ có trên Firestore',
          sheetName: customerSheet.name,
          recommendedAction: 'push_to_sheets',
          firestoreData: cust
        });
      } else {
        const diffMs = fsTimestamp - sheetRecord.timestamp;
        const absDiff = Math.abs(diffMs);
        let status: ReconciliationStatus = 'in_sync';
        let statusText = 'Đồng bộ';
        let recommendedAction: ReconciliationItem['recommendedAction'] = 'no_action';

        if (absDiff <= 3000) {
          status = 'in_sync';
          statusText = 'Đồng bộ khớp phiên bản';
          recommendedAction = 'no_action';
        } else if (diffMs > 3000) {
          status = 'firestore_newer';
          statusText = 'Firestore mới hơn';
          recommendedAction = 'push_to_sheets';
        } else {
          status = 'sheets_newer';
          statusText = 'Google Sheets mới hơn';
          recommendedAction = 'pull_to_firestore';
        }

        items.push({
          id,
          name: cust.name || sheetRecord.name || `KH ${id}`,
          entityType: 'customer',
          entityLabel: 'Khách hàng',
          firestoreUpdatedAt: fsUpdatedAt,
          firestoreTimestamp: fsTimestamp,
          sheetsUpdatedAt: sheetRecord.dateStr,
          sheetsTimestamp: sheetRecord.timestamp,
          status,
          statusText,
          sheetName: customerSheet.name,
          sheetRowIndex: sheetRecord.rowIndex,
          diffSummary: `Chênh lệch: ${Math.round(absDiff / 1000)}s`,
          firestoreData: cust,
          sheetsData: sheetRecord.row,
          recommendedAction
        });
      }
    });

    sheetCustMap.forEach((sheetRecord, lowerId) => {
      if (!processedCustIds.has(lowerId)) {
        items.push({
          id: String(sheetRecord.row[0] || lowerId),
          name: sheetRecord.name,
          entityType: 'customer',
          entityLabel: 'Khách hàng',
          firestoreUpdatedAt: undefined,
          firestoreTimestamp: 0,
          sheetsUpdatedAt: sheetRecord.dateStr,
          sheetsTimestamp: sheetRecord.timestamp,
          status: 'only_in_sheets',
          statusText: 'Chỉ có trên Google Sheets',
          sheetName: customerSheet.name,
          sheetRowIndex: sheetRecord.rowIndex,
          recommendedAction: 'pull_to_firestore',
          sheetsData: sheetRecord.row
        });
      }
    });
  }

  // --- 3. Đối soát Thiết Bị (thietbi) ---
  if (equipmentSheet && equipmentSheet.rows.length > 0) {
    const rows = equipmentSheet.rows;
    let headerIdx = 0;
    for (let r = 0; r < Math.min(5, rows.length); r++) {
      if (rows[r]?.some((cell: any) => String(cell).toLowerCase().includes('mã thiết bị') || String(cell).toLowerCase().includes('tên thiết bị'))) {
        headerIdx = r;
        break;
      }
    }

    const dataRows = rows.slice(headerIdx + 1);
    const sheetEqMap = new Map<string, { rowIndex: number; row: any[]; timestamp: number; dateStr: string; name: string }>();

    dataRows.forEach((row, idx) => {
      if (!row || row.length === 0) return;
      const id = String(row[0] || '').trim();
      if (!id || id.toLowerCase().includes('mã thiết bị')) return;
      const updatedAtStr = String(row[12] || row[11] || row[10] || '');
      const timestamp = parseVersionTimestamp(updatedAtStr);
      sheetEqMap.set(id.toLowerCase(), {
        rowIndex: headerIdx + 1 + idx + 1,
        row,
        timestamp,
        dateStr: updatedAtStr,
        name: String(row[1] || id)
      });
    });

    const processedEqIds = new Set<string>();

    firestoreData.equipment.forEach(eq => {
      const id = (eq.id || '').trim();
      if (!id) return;
      processedEqIds.add(id.toLowerCase());

      const fsUpdatedAt = eq.updatedAt || eq.lastCheck || eq.createdAt || '';
      const fsTimestamp = parseVersionTimestamp(fsUpdatedAt);
      const sheetRecord = sheetEqMap.get(id.toLowerCase());

      if (!sheetRecord) {
        items.push({
          id,
          name: eq.name || `Thiết bị ${id}`,
          entityType: 'equipment',
          entityLabel: 'Thiết bị trạm',
          firestoreUpdatedAt: fsUpdatedAt,
          firestoreTimestamp: fsTimestamp,
          sheetsUpdatedAt: undefined,
          sheetsTimestamp: 0,
          status: 'only_in_firestore',
          statusText: 'Chỉ có trên Firestore',
          sheetName: equipmentSheet.name,
          recommendedAction: 'push_to_sheets',
          firestoreData: eq
        });
      } else {
        const diffMs = fsTimestamp - sheetRecord.timestamp;
        const absDiff = Math.abs(diffMs);
        let status: ReconciliationStatus = 'in_sync';
        let statusText = 'Đồng bộ';
        let recommendedAction: ReconciliationItem['recommendedAction'] = 'no_action';

        if (absDiff <= 3000) {
          status = 'in_sync';
          statusText = 'Đồng bộ khớp phiên bản';
          recommendedAction = 'no_action';
        } else if (diffMs > 3000) {
          status = 'firestore_newer';
          statusText = 'Firestore mới hơn';
          recommendedAction = 'push_to_sheets';
        } else {
          status = 'sheets_newer';
          statusText = 'Google Sheets mới hơn';
          recommendedAction = 'pull_to_firestore';
        }

        items.push({
          id,
          name: eq.name || sheetRecord.name || `Thiết bị ${id}`,
          entityType: 'equipment',
          entityLabel: 'Thiết bị trạm',
          firestoreUpdatedAt: fsUpdatedAt,
          firestoreTimestamp: fsTimestamp,
          sheetsUpdatedAt: sheetRecord.dateStr,
          sheetsTimestamp: sheetRecord.timestamp,
          status,
          statusText,
          sheetName: equipmentSheet.name,
          sheetRowIndex: sheetRecord.rowIndex,
          diffSummary: `Chênh lệch: ${Math.round(absDiff / 1000)}s`,
          firestoreData: eq,
          sheetsData: sheetRecord.row,
          recommendedAction
        });
      }
    });

    sheetEqMap.forEach((sheetRecord, lowerId) => {
      if (!processedEqIds.has(lowerId)) {
        items.push({
          id: String(sheetRecord.row[0] || lowerId),
          name: sheetRecord.name,
          entityType: 'equipment',
          entityLabel: 'Thiết bị trạm',
          firestoreUpdatedAt: undefined,
          firestoreTimestamp: 0,
          sheetsUpdatedAt: sheetRecord.dateStr,
          sheetsTimestamp: sheetRecord.timestamp,
          status: 'only_in_sheets',
          statusText: 'Chỉ có trên Google Sheets',
          sheetName: equipmentSheet.name,
          sheetRowIndex: sheetRecord.rowIndex,
          recommendedAction: 'pull_to_firestore',
          sheetsData: sheetRecord.row
        });
      }
    });
  }

  // --- 4. Đối soát Quản lý Kho (QuanLyKho) ---
  if (inventorySheet && inventorySheet.rows.length > 0) {
    const rows = inventorySheet.rows;
    let headerIdx = 0;
    for (let r = 0; r < Math.min(5, rows.length); r++) {
      if (rows[r]?.some((cell: any) => String(cell).toLowerCase().includes('mã vật tư') || String(cell).toLowerCase().includes('tên vật tư') || String(cell).toLowerCase().includes('mã vt'))) {
        headerIdx = r;
        break;
      }
    }

    const dataRows = rows.slice(headerIdx + 1);
    const sheetInvMap = new Map<string, { rowIndex: number; row: any[]; timestamp: number; dateStr: string; name: string }>();

    dataRows.forEach((row, idx) => {
      if (!row || row.length === 0) return;
      const id = String(row[0] || '').trim();
      if (!id || id.toLowerCase().includes('mã vật tư') || id.toLowerCase().includes('mã vt')) return;
      const updatedAtStr = String(row[10] || row[9] || '');
      const timestamp = parseVersionTimestamp(updatedAtStr);
      sheetInvMap.set(id.toLowerCase(), {
        rowIndex: headerIdx + 1 + idx + 1,
        row,
        timestamp,
        dateStr: updatedAtStr,
        name: String(row[1] || id)
      });
    });

    const processedInvIds = new Set<string>();

    firestoreData.inventory.forEach(item => {
      const id = (item.id || '').trim();
      if (!id) return;
      processedInvIds.add(id.toLowerCase());

      const fsUpdatedAt = item.updatedAt || item.createdAt || '';
      const fsTimestamp = parseVersionTimestamp(fsUpdatedAt);
      const sheetRecord = sheetInvMap.get(id.toLowerCase());

      if (!sheetRecord) {
        items.push({
          id,
          name: item.name || `Vật tư ${id}`,
          entityType: 'inventory',
          entityLabel: 'Vật tư kho',
          firestoreUpdatedAt: fsUpdatedAt,
          firestoreTimestamp: fsTimestamp,
          sheetsUpdatedAt: undefined,
          sheetsTimestamp: 0,
          status: 'only_in_firestore',
          statusText: 'Chỉ có trên Firestore',
          sheetName: inventorySheet.name,
          recommendedAction: 'push_to_sheets',
          firestoreData: item
        });
      } else {
        const diffMs = fsTimestamp - sheetRecord.timestamp;
        const absDiff = Math.abs(diffMs);
        let status: ReconciliationStatus = 'in_sync';
        let statusText = 'Đồng bộ';
        let recommendedAction: ReconciliationItem['recommendedAction'] = 'no_action';

        if (absDiff <= 3000) {
          status = 'in_sync';
          statusText = 'Đồng bộ khớp phiên bản';
          recommendedAction = 'no_action';
        } else if (diffMs > 3000) {
          status = 'firestore_newer';
          statusText = 'Firestore mới hơn';
          recommendedAction = 'push_to_sheets';
        } else {
          status = 'sheets_newer';
          statusText = 'Google Sheets mới hơn';
          recommendedAction = 'pull_to_firestore';
        }

        items.push({
          id,
          name: item.name || sheetRecord.name || `Vật tư ${id}`,
          entityType: 'inventory',
          entityLabel: 'Vật tư kho',
          firestoreUpdatedAt: fsUpdatedAt,
          firestoreTimestamp: fsTimestamp,
          sheetsUpdatedAt: sheetRecord.dateStr,
          sheetsTimestamp: sheetRecord.timestamp,
          status,
          statusText,
          sheetName: inventorySheet.name,
          sheetRowIndex: sheetRecord.rowIndex,
          diffSummary: `Chênh lệch: ${Math.round(absDiff / 1000)}s`,
          firestoreData: item,
          sheetsData: sheetRecord.row,
          recommendedAction
        });
      }
    });

    sheetInvMap.forEach((sheetRecord, lowerId) => {
      if (!processedInvIds.has(lowerId)) {
        items.push({
          id: String(sheetRecord.row[0] || lowerId),
          name: sheetRecord.name,
          entityType: 'inventory',
          entityLabel: 'Vật tư kho',
          firestoreUpdatedAt: undefined,
          firestoreTimestamp: 0,
          sheetsUpdatedAt: sheetRecord.dateStr,
          sheetsTimestamp: sheetRecord.timestamp,
          status: 'only_in_sheets',
          statusText: 'Chỉ có trên Google Sheets',
          sheetName: inventorySheet.name,
          sheetRowIndex: sheetRecord.rowIndex,
          recommendedAction: 'pull_to_firestore',
          sheetsData: sheetRecord.row
        });
      }
    });
  }

  // Tổng hợp thống kê
  const inSyncCount = items.filter(i => i.status === 'in_sync').length;
  const firestoreNewerCount = items.filter(i => i.status === 'firestore_newer').length;
  const sheetsNewerCount = items.filter(i => i.status === 'sheets_newer').length;
  const conflictsCount = items.filter(i => i.status === 'conflict').length;
  const onlyInFirestoreCount = items.filter(i => i.status === 'only_in_firestore').length;
  const onlyInSheetsCount = items.filter(i => i.status === 'only_in_sheets').length;

  const hasConflicts = conflictsCount > 0 || (sheetsNewerCount > 0 && firestoreNewerCount > 0);
  // An toàn ghi đè lên Sheets chỉ khi không có bản ghi nào trên Sheets mới hơn Firestore
  const canSafeOverwriteSheets = sheetsNewerCount === 0 && conflictsCount === 0;
  // An toàn ghi đè lên Firestore chỉ khi không có bản ghi nào trên Firestore mới hơn Sheets
  const canSafeOverwriteFirestore = firestoreNewerCount === 0 && conflictsCount === 0;

  let summaryMessage = 'Dữ liệu giữa Firestore và Google Sheets hoàn toàn nhất quán.';
  if (sheetsNewerCount > 0 && firestoreNewerCount > 0) {
    summaryMessage = `Phát hiện xung đột phiên bản: ${sheetsNewerCount} mục trên Sheets mới hơn và ${firestoreNewerCount} mục trên Firestore mới hơn! Cần hòa giải (reconcile) trước khi ghi đè.`;
  } else if (sheetsNewerCount > 0) {
    summaryMessage = `Có ${sheetsNewerCount} mục trên Google Sheets được sửa đổi gần đây hơn Firestore. Không nên ghi đè mù quáng từ Firestore lên Sheets!`;
  } else if (firestoreNewerCount > 0) {
    summaryMessage = `Có ${firestoreNewerCount} mục trên Firestore mới hơn Sheets. Có thể an toàn đẩy lên Google Sheets.`;
  }

  return {
    timestamp: new Date().toISOString(),
    totalChecked: items.length,
    inSyncCount,
    firestoreNewerCount,
    sheetsNewerCount,
    conflictsCount,
    onlyInFirestoreCount,
    onlyInSheetsCount,
    items,
    hasConflicts,
    canSafeOverwriteSheets,
    canSafeOverwriteFirestore,
    summaryMessage
  };
};
