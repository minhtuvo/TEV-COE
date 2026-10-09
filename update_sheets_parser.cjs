const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf8');

// 1. Update auto resolve check if (allEquipment.length <= 13)
code = code.replace(
  'if (allEquipment.length <= 13) {',
  'if (allEquipment.length < 104) {'
);

// 2. Update getCommonIndices
const oldIndices = `      // Helper to map common fields
      const getCommonIndices = (header: string[]) => ({
        dateIdx: getColumnIndex(header, ['Ngày', 'Date', 'Last Check', 'Thời gian', 'Time', 'Ngày kiểm tra']),
        customerIdx: getColumnIndex(header, ['Khách hàng', 'Customer', 'Client', 'Đơn vị']),
        factoryIdx: getColumnIndex(header, ['Nhà máy', 'Factory', 'Site', 'Trạm', 'Khu vực']),
        locationIdx: getColumnIndex(header, ['Vị trí', 'Location', 'Khu vực', 'Area', 'Ngăn lộ']),
        idIdx: getColumnIndex(header, ['Mã thiết bị', 'Equipment ID', 'Mã TB', 'Tag']),
        nameIdx: getColumnIndex(header, ['Tên thiết bị', 'Equipment Name', 'Tên TB', 'Description']),
        typeIdx: getColumnIndex(header, ['Loại thiết bị', 'Type', 'Loại TB', 'Phân loại']),
        healthIdx: getColumnIndex(header, ['Chỉ số sức khỏe', 'Health', 'HI', 'Sức khỏe', 'Điểm']),
        statusIdx: getColumnIndex(header, ['Trạng thái', 'Status', 'Tình trạng', 'Kết quả', 'Đánh giá']),
        ageIdx: getColumnIndex(header, ['Tuổi thọ', 'Age', 'Năm vận hành', 'Năm SX']),
        dutyFactorIdx: getColumnIndex(header, ['Hệ số làm việc', 'Duty Factor', 'Hệ số tải']),
        criticalityIdx: getColumnIndex(header, ['Độ quan trọng', 'Criticality', 'Phân loại rủi ro', 'Mức độ quan trọng']),
        linksIdx: getColumnIndex(header, ['File đính kèm', 'Links', 'Tài liệu', 'Link'])
      });`;

const newIndices = `      // Helper to map common fields (hỗ trợ toàn diện 104 thiết bị qua tất cả các sheet)
      const getCommonIndices = (header: string[]) => ({
        dateIdx: getColumnIndex(header, ['Ngày', 'Date', 'Last Check', 'Thời gian', 'Time', 'Ngày kiểm tra', 'Lần kiểm tra cuối', 'Kiểm tra gần nhất', 'Ngày tạo']),
        customerIdx: getColumnIndex(header, ['Khách hàng', 'Customer', 'Client', 'Đơn vị', 'Chủ đầu tư']),
        factoryIdx: getColumnIndex(header, ['Nhà máy', 'Factory', 'Site', 'Trạm', 'Khu vực', 'Nhà máy / Site', 'Cơ sở']),
        locationIdx: getColumnIndex(header, ['Vị trí', 'Location', 'Khu vực', 'Area', 'Ngăn lộ', 'Vị trí / Khu vực', 'Tuyến / Vị trí', 'Tủ']),
        idIdx: getColumnIndex(header, ['Mã thiết bị', 'Equipment ID', 'Mã TB', 'Tag', 'Mã máy cắt', 'Mã MC', 'Mã ĐC', 'Mã MBA', 'Mã máy phát', 'Mã cáp', 'Mã rơ le', 'Mã pin', 'Mã ID', 'ID']),
        nameIdx: getColumnIndex(header, ['Tên thiết bị', 'Equipment Name', 'Tên TB', 'Description', 'Tên máy cắt', 'Tên động cơ', 'Tên máy phát', 'Tên biến áp', 'Tên tủ điện', 'Tên tuyến cáp', 'Tên dàn pin/UPS', 'Tên rơ le', 'Tên', 'Thiết bị']),
        typeIdx: getColumnIndex(header, ['Loại thiết bị', 'Type', 'Loại TB', 'Phân loại', 'Chủng loại', 'Loại', 'Chủng loại (VCB/ACB/SF6)']),
        healthIdx: getColumnIndex(header, ['Chỉ số sức khỏe', 'Health', 'HI', 'Sức khỏe', 'Điểm', 'Điểm sức khỏe (HI %)', 'Điểm HI']),
        statusIdx: getColumnIndex(header, ['Trạng thái', 'Status', 'Tình trạng', 'Kết quả', 'Đánh giá']),
        ageIdx: getColumnIndex(header, ['Tuổi thọ', 'Age', 'Năm vận hành', 'Năm SX', 'Năm sản xuất']),
        dutyFactorIdx: getColumnIndex(header, ['Hệ số làm việc', 'Duty Factor', 'Hệ số tải']),
        criticalityIdx: getColumnIndex(header, ['Độ quan trọng', 'Criticality', 'Phân loại rủi ro', 'Mức độ quan trọng']),
        linksIdx: getColumnIndex(header, ['File đính kèm', 'Links', 'Tài liệu', 'Link', 'Báo cáo'])
      });`;

if (code.includes(oldIndices)) {
  code = code.replace(oldIndices, newIndices);
  console.log('getCommonIndices updated');
} else {
  console.warn('Could not find oldIndices exact string');
}

// 3. Extract equipment from Work Orders as well
const oldWoPush = `            newWoList.push(woObj);`;
const newWoPush = `            newWoList.push(woObj);
            // Trích xuất thiết bị liên kết từ Work Orders vào danh mục thiết bị
            rawEq.forEach((eqCode: string) => {
              const code = (eqCode || '').trim();
              if (code && !code.toLowerCase().includes('mã')) {
                const autoType = code.startsWith('MVCB') || code.startsWith('MC-') ? 'Máy cắt' :
                                 code.startsWith('SWG') ? 'Tủ điện' :
                                 code.startsWith('TRF') ? 'Máy biến áp' :
                                 code.startsWith('MTR') ? 'Động cơ' :
                                 code.startsWith('GEN') ? 'Máy phát' :
                                 code.startsWith('CAB') ? 'Cáp điện' :
                                 code.startsWith('BAT') ? 'Pin & UPS' :
                                 code.startsWith('REL') ? 'Rơ le bảo vệ' : 'Tủ điện';
                newEquipmentList.push({
                  id: code,
                  name: eqTitle || ('Thiết bị ' + code),
                  customer: customerVal || '',
                  factory: customerVal ? ('Nhà máy ' + customerVal) : 'Nhà máy Bắc Ninh',
                  location: 'Trạm điện phân phối',
                  type: autoType,
                  health: 92,
                  status: 'healthy',
                  criticality: 'B',
                  lastCheck: new Date().toLocaleDateString('vi-VN'),
                  measurements: {}
                });
              }
            });`;

if (code.includes(oldWoPush)) {
  code = code.replace(oldWoPush, newWoPush);
  console.log('WO equipment extraction added');
}

// 4. Replace the old equipment parsing block
const oldEqBlockStart = `        // Determine equipment type
        let eqType: 'Máy biến áp' | 'Tủ điện' | 'Động cơ' | 'Inverter' = 'Máy biến áp';`;

const oldEqBlockEnd = `        if (newEquipmentList.length > 0) {
        // Deduplicate equipment list (keep latest by date)`;

const idxStart = code.indexOf(oldEqBlockStart);
const idxEnd = code.indexOf(oldEqBlockEnd);

if (idxStart === -1 || idxEnd === -1) {
  console.error('Equipment block markers not found', { idxStart, idxEnd });
  process.exit(1);
}

const newEqBlock = `        // Xác định phân loại thiết bị mặc định theo tên Sheet
        let defaultSheetType: string = 'Máy biến áp';
        if (lowerSheetName.includes('máy cắt') || lowerSheetName.includes('breaker') || lowerSheetName.includes('vcb') || lowerSheetName.includes('acb')) {
          defaultSheetType = 'Máy cắt';
        } else if (lowerSheetName.includes('tủ điện') || lowerSheetName.includes('trung thế') || lowerSheetName.includes('hạ thế') || lowerSheetName.includes('switchgear') || lowerSheetName.includes('swg') || lowerSheetName.includes('rmu') || lowerSheetName.includes('mcc') || lowerSheetName.includes('tev')) {
          defaultSheetType = 'Tủ điện';
        } else if (lowerSheetName.includes('động cơ') || lowerSheetName.includes('motor') || lowerSheetName.includes('máy bơm') || lowerSheetName.includes('pump') || lowerSheetName.includes('quạt')) {
          defaultSheetType = 'Động cơ';
        } else if (lowerSheetName.includes('máy phát') || lowerSheetName.includes('generator')) {
          defaultSheetType = 'Máy phát';
        } else if (lowerSheetName.includes('biến áp') || lowerSheetName.includes('mba') || lowerSheetName.includes('transformer')) {
          defaultSheetType = 'Máy biến áp';
        } else if (lowerSheetName.includes('cáp điện') || lowerSheetName.includes('cable') || lowerSheetName.includes('cáp')) {
          defaultSheetType = 'Cáp điện';
        } else if (lowerSheetName.includes('pin') || lowerSheetName.includes('ups') || lowerSheetName.includes('bess') || lowerSheetName.includes('battery') || lowerSheetName.includes('ắc quy')) {
          defaultSheetType = 'Pin & UPS';
        } else if (lowerSheetName.includes('rơ le') || lowerSheetName.includes('relay') || lowerSheetName.includes('role')) {
          defaultSheetType = 'Rơ le bảo vệ';
        } else if (lowerSheetName.includes('inverter') || lowerSheetName.includes('biến tần') || lowerSheetName.includes('solar') || lowerSheetName.includes('wind')) {
          defaultSheetType = 'Inverter';
        }

        const dataRows = rows.slice(headerIdx + 1);
        for (const row of dataRows) {
          if (!row || row.length === 0) continue;
          const idVal = idx.idIdx !== -1 ? row[idx.idIdx] : (row[1] && String(row[1]).trim().length >= 3 && !String(row[1]).includes(' ') ? row[1] : (row[4] || row[0]));
          if (!idVal) continue;
          const idStr = String(idVal).trim();
          if (idStr === '' || idStr.toLowerCase().includes('mã ') || idStr.toLowerCase().includes('tag id') || idStr.toLowerCase().includes('equipment id')) continue;

          const nameVal = idx.nameIdx !== -1 ? row[idx.nameIdx] : (row[2] || row[1] || ('Thiết bị ' + idStr));
          const nameStr = String(nameVal || ('Thiết bị ' + idStr)).trim();

          // Xác định chính xác loại thiết bị
          let resolvedType = defaultSheetType;
          if (idx.typeIdx !== -1 && row[idx.typeIdx] && String(row[idx.typeIdx]).trim() !== '') {
            resolvedType = String(row[idx.typeIdx]).trim();
          } else if (idStr.startsWith('TRF')) {
            resolvedType = 'Máy biến áp';
          } else if (idStr.startsWith('MVCB') || idStr.startsWith('MC-')) {
            resolvedType = 'Máy cắt';
          } else if (idStr.startsWith('SWG')) {
            resolvedType = 'Tủ điện';
          } else if (idStr.startsWith('MTR')) {
            resolvedType = 'Động cơ';
          } else if (idStr.startsWith('GEN')) {
            resolvedType = 'Máy phát';
          } else if (idStr.startsWith('CAB')) {
            resolvedType = 'Cáp điện';
          } else if (idStr.startsWith('BAT')) {
            resolvedType = 'Pin & UPS';
          } else if (idStr.startsWith('REL')) {
            resolvedType = 'Rơ le bảo vệ';
          } else if (idStr.startsWith('INV')) {
            resolvedType = 'Inverter';
          }

          let healthVal = idx.healthIdx !== -1 ? parseFloat(row[idx.healthIdx]?.toString().replace(',', '.') || '92') : 92;
          if (isNaN(healthVal)) healthVal = 92;
          let statusVal = mapStatusFromSheet(idx.statusIdx !== -1 ? row[idx.statusIdx] : '');
          if (!statusVal) {
            statusVal = healthVal < 60 ? 'critical' : (healthVal < 80 ? 'warning' : 'healthy');
          }

          const customerVal = idx.customerIdx !== -1 ? String(row[idx.customerIdx] || '') : (row[3] || '');
          const factoryVal = idx.factoryIdx !== -1 ? String(row[idx.factoryIdx] || '') : (row[4] || '');
          const locationVal = idx.locationIdx !== -1 ? String(row[idx.locationIdx] || '') : (row[5] || '');
          const criticalityVal = idx.criticalityIdx !== -1 ? (row[idx.criticalityIdx]?.toString().trim().toUpperCase() || 'B') : 'B';
          const lastCheckVal = idx.dateIdx !== -1 ? (row[idx.dateIdx] || new Date().toLocaleDateString('vi-VN')) : new Date().toLocaleDateString('vi-VN');

          const measurementsObj: Record<string, any> = {};
          header.forEach((colName, cIdx) => {
            if (colName && row[cIdx] !== undefined && row[cIdx] !== '') {
              measurementsObj[colName] = row[cIdx];
            }
          });

          const eqData: any = {
            id: idStr,
            name: nameStr,
            customer: customerVal,
            factory: factoryVal,
            location: locationVal,
            type: resolvedType,
            health: healthVal,
            status: statusVal,
            criticality: criticalityVal,
            lastCheck: lastCheckVal,
            measurements: measurementsObj,
            rawData: [...row]
          };

          newEquipmentList.push(eqData);
          newReportsList.push({
            id: 'REP-' + idStr + '-' + Math.floor(Math.random() * 10000),
            date: lastCheckVal,
            lastCheck: lastCheckVal,
            customer: customerVal,
            location: locationVal,
            equipmentId: idStr,
            equipmentName: nameStr,
            factory: factoryVal,
            type: resolvedType,
            inspector: 'FSE',
            status: statusVal,
            notes: 'Đồng bộ từ Sheet ' + sheetName,
            measurements: measurementsObj,
            fileUrl: (idx.linksIdx !== -1 ? row[idx.linksIdx] : '') || '#',
            rawData: [...row]
          });
        }
      }

`;

code = code.slice(0, idxStart) + newEqBlock + code.slice(idxEnd);
console.log('Equipment parsing block replaced');

// 5. Update deduplication to merge with full 104 equipment fleet
const oldDedupStart = `      if (newEquipmentList.length > 0) {
        // Deduplicate equipment list (keep latest by date)
        const uniqueEqMap = new Map();`;

const oldDedupEnd = `        const uniqueEquipmentList = Array.from(uniqueEqMap.values());`;

const dStart = code.indexOf(oldDedupStart);
const dEnd = code.indexOf(oldDedupEnd);

if (dStart !== -1 && dEnd !== -1) {
  const newDedup = `      if (newEquipmentList.length > 0) {
        // Hợp nhất và khử trùng lặp thiết bị (giữ bản ghi chi tiết nhất và mới nhất)
        const uniqueEqMap = new Map();
        newEquipmentList.forEach(eq => {
          const key = (eq.id || '').toString().toLowerCase().trim();
          if (!key) return;
          const existing = uniqueEqMap.get(key);
          if (!existing) {
            uniqueEqMap.set(key, eq);
          } else {
            uniqueEqMap.set(key, {
              ...existing,
              ...eq,
              name: (eq.name && !eq.name.startsWith('Thiết bị ')) ? eq.name : existing.name,
              customer: eq.customer || existing.customer,
              factory: eq.factory || existing.factory,
              location: eq.location || existing.location,
              type: eq.type || existing.type,
              health: (typeof eq.health === 'number' && eq.health > 0) ? eq.health : existing.health,
              status: eq.status || existing.status,
              measurements: { ...(existing.measurements || {}), ...(eq.measurements || {}) }
            });
          }
        });

        // Bổ sung đầy đủ 104 thiết bị nếu Sheet chưa có đủ
        const masterDefaults = getInitialEquipmentList();
        masterDefaults.forEach((defEq: any) => {
          const key = (defEq.id || '').toString().toLowerCase().trim();
          if (!uniqueEqMap.has(key)) {
            uniqueEqMap.set(key, defEq);
          }
        });

        const uniqueEquipmentList = Array.from(uniqueEqMap.values());
        uniqueEquipmentList.sort((a: any, b: any) => (a.name || a.id || '').localeCompare(b.name || b.id || ''));`;

  code = code.slice(0, dStart) + newDedup + code.slice(dEnd + oldDedupEnd.length);
  console.log('Deduplication logic updated to maintain full 104 fleet');
} else {
  console.warn('Deduplication markers not found', { dStart, dEnd });
}

fs.writeFileSync('src/App.tsx', code, 'utf8');
console.log('Done updating handleFetchFromSheets in App.tsx!');
