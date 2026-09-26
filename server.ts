import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import { google } from 'googleapis';
import path from 'path';
import multer from 'multer';
import { Readable } from 'stream';
import {
  findUpcomingPMOrders,
  sendPMNotifications,
  getEmailTransporter,
  generatePMEmailHtml,
} from './server/pmNotifications';
import {
  generateNetaAiAnalysis,
  NetaInputPayload,
} from './server/netaAnalyzer';
import {
  generateNetaLargeAiAnalysis,
  NetaLargeInputPayload,
} from './server/netaLargeAnalyzer';
import {
  generateNetaBessAiAnalysis,
  NetaBessInputPayload,
} from './server/netaBessAnalyzer';
import {
  generateNetaPvAiAnalysis,
  NetaPvInputPayload,
} from './server/netaPvAnalyzer';
import {
  generateNetaEvseAiAnalysis,
  NetaEvseInputPayload,
} from './server/netaEvseAnalyzer';
import {
  generateNetaCableAiAnalysis,
  NetaCableInputPayload,
} from './server/netaCableAnalyzer';
import {
  generateNetaSf6SwitchAiAnalysis,
  NetaSf6SwitchInputPayload,
} from './server/netaSf6SwitchAnalyzer';
import {
  generateNetaThermographyAiAnalysis,
  NetaThermographyInputPayload,
} from './server/netaThermographyAnalyzer';
import {
  generateNetaPdAiAnalysis,
  NetaPdInputPayload,
} from './server/netaPartialDischargeAnalyzer';
import {
  generateNetaMotorAiAnalysis,
  NetaMotorInputPayload,
} from './server/netaMotorAnalyzer';
import {
  generateNetaLiquidTransformerAiAnalysis,
  NetaLiquidTransformerInputPayload,
} from './server/netaLiquidTransformerAnalyzer';
import {
  generateNetaRelayAiAnalysis,
  NetaRelayInputPayload,
} from './server/netaRelayAnalyzer';
import {
  generateNetaUpsAiAnalysis,
  NetaUpsInputPayload,
} from './server/netaUpsAnalyzer';
import {
  generateNetaSyncMachineryAiAnalysis,
  NetaSyncMachineryInputPayload,
} from './server/netaSyncMachineryAnalyzer';
import {
  generateNetaBatteryFloodedAiAnalysis,
  NetaBatteryFloodedInputPayload,
} from './server/netaBatteryFloodedAnalyzer';
import {
  generateNetaBatteryVrlaAiAnalysis,
  NetaBatteryVrlaInputPayload,
} from './server/netaBatteryVrlaAnalyzer';
import {
  generateNetaAtsAiAnalysis,
  NetaAtsInputPayload,
} from './server/netaAtsAnalyzer';
import {
  generateNetaGeneratorAiAnalysis,
  NetaGeneratorInputPayload,
} from './server/netaEngineGeneratorAnalyzer';
import {
  generateNetaSwitchgearAiAnalysis,
  NetaSwitchgearInputPayload,
} from './server/netaSwitchgearAnalyzer';
import {
  generateNetaCableLvAiAnalysis,
  NetaCableLvInputPayload,
} from './server/netaCableLvAnalyzer';
import {
  generateNetaLvBreakerAiAnalysis,
  NetaLvBreakerInputPayload,
} from './server/netaLvBreakerAnalyzer';
import {
  generateNetaGroundingAiAnalysis,
  NetaGroundingInputPayload,
} from './server/netaGroundingAnalyzer';
import {
  generateNetaDcMotorAiAnalysis,
  NetaDcMotorInputPayload,
} from './server/netaDcMotorAnalyzer';
import {
  extractNameplateDataWithAi,
  NameplateOcrInput,
} from './server/nameplateOcrAnalyzer';
import {
  analyzePvCellWithGemini,
  PvThermalInputPayload,
} from './server/pvCellThermalAnalyzer';

dotenv.config();

export const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

// Set up multer for file uploads (in-memory storage)
const upload = multer({ storage: multer.memoryStorage() });

// In-memory token storage for prototype (in production, use a database)
// We'll store tokens by a simple session ID or just globally for this single-user prototype
let globalTokens: any = null;

// Helper to get redirect URI based on the request origin
const getRedirectUri = (req: express.Request) => {
  if (process.env.APP_URL) {
    return `${process.env.APP_URL}/auth/callback`;
  }
  const host = req.get('host');
  const protocol = req.headers['x-forwarded-proto'] || req.protocol;
  const origin = req.get('origin') || `${protocol}://${host}`;
  return `${origin}/auth/callback`;
};

// 1. Get OAuth URL
app.get('/api/auth/google/url', (req, res) => {
  const redirectUri = req.query.redirectUri as string;
  
  if (!redirectUri) {
    return res.status(400).json({ error: 'redirectUri is required' });
  }

  if (!process.env.GOOGLE_CLIENT_ID || !process.env.GOOGLE_CLIENT_SECRET) {
    return res.status(500).json({ 
      error: 'Thiếu cấu hình GOOGLE_CLIENT_ID hoặc GOOGLE_CLIENT_SECRET trong Settings -> Secrets.' 
    });
  }

  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    redirectUri
  );

  const url = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive.file'
    ],
    state: redirectUri, // Pass the redirectUri in the state to use it in the callback
    redirect_uri: redirectUri // Explicitly pass redirect_uri
  });

  console.log('Generated OAuth URL:', url);
  console.log('Using redirectUri:', redirectUri);

  res.json({ url });
});

// 2. Handle OAuth Callback
app.get(['/auth/callback', '/auth/callback/', '/api/auth/callback', '/api/auth/callback/'], async (req, res) => {
  const { code, state } = req.query;
  
  try {
    const redirectUri = state as string;

    if (!redirectUri) {
      throw new Error('Missing state (redirectUri) in callback');
    }

    const oauth2Client = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET,
      redirectUri
    );

    const { tokens } = await oauth2Client.getToken(code as string);
    globalTokens = tokens; // Store globally for this prototype

    // Send success message to parent window and close popup
    res.send(`
      <html>
        <body>
          <script>
            if (window.opener) {
              window.opener.postMessage({ type: 'OAUTH_AUTH_SUCCESS', tokens: ${JSON.stringify(tokens)} }, '*');
              window.close();
            } else {
              window.location.href = '/';
            }
          </script>
          <p>Authentication successful. This window should close automatically.</p>
        </body>
      </html>
    `);
  } catch (error) {
    console.error('Error exchanging code for tokens:', error);
    res.status(500).send('Authentication failed');
  }
});

// 3. Check Auth Status
app.get('/api/auth/status', (req, res) => {
  try {
    console.log('Received auth status check request');
    const authHeader = req.headers.authorization;
    console.log('Auth status check. Header present:', !!authHeader);
    if (authHeader && authHeader.startsWith('Bearer ')) {
      return res.json({ isAuthenticated: true });
    }
    res.json({ isAuthenticated: !!globalTokens });
  } catch (error) {
    console.error('Error in auth status check:', error);
    res.status(500).json({ isAuthenticated: false, error: 'Internal server error' });
  }
});

// Helper to extract clean spreadsheet ID
const extractSpreadsheetId = (input?: string): string => {
  if (!input) return '';
  const trimmed = input.trim();
  if (trimmed.includes('docs.google.com/spreadsheets/d/')) {
    const match = trimmed.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match) return match[1];
  }
  return trimmed;
};

// Helper to get tokens from request
const getTokensFromRequest = (req: express.Request) => {
  const authHeader = req.headers.authorization;
  if (authHeader && authHeader.startsWith('Bearer ')) {
    const raw = authHeader.substring(7).trim();
    try {
      return JSON.parse(raw);
    } catch (e) {
      if (raw.length > 5) {
        return { access_token: raw };
      }
    }
  }
  return globalTokens;
};

// Standard column headers for TEV sheets
export const TEV_WO_HEADERS = [
  'Mã WO', 'Mã Work Permit', 'Tiêu đề', 'Mô tả', 'Thiết bị', 'Failure Code',
  'Khách hàng', 'Loại công việc', 'Ngoài kế hoạch', 'Mức độ ưu tiên', 'Trạng thái',
  'Người thực hiện (PIC)', 'Vai trò Phê duyệt', 'Vai trò Thực hiện', 'Yêu cầu Cô lập',
  'Hạn hoàn thành', 'Vật tư sử dụng', 'Ngày tạo', 'Ngày cập nhật',
  'Bắt đầu dừng máy', 'Bắt đầu sửa chữa', 'Kết thúc sửa chữa', 'Chạy lại máy'
];

export const TEV_CUSTOMER_HEADERS = [
  'Mã khách hàng', 'Tên khách hàng', 'Nhà máy / Trạm', 'Email', 'Số điện thoại', 'Địa chỉ', 'Ngày tạo', 'Ngày cập nhật'
];

export const TEV_EQUIPMENT_HEADERS = [
  'Mã thiết bị', 'Tên thiết bị', 'Loại thiết bị', 'Khách hàng', 'Nhà máy / Site', 'Vị trí / Khu vực',
  'Trạng thái', 'Điểm sức khỏe (HI %)', 'Thông số kỹ thuật (Specs)', 'Bảng tên Nameplate', 'Lần kiểm tra cuối', 'Ngày tạo', 'Ngày cập nhật'
];

// Equipment-Specific Detailed Sheets Headers
export const TEV_BREAKER_HEADERS = [
  'Mã WO', 'Mã thiết bị', 'Tên máy cắt', 'Khách hàng', 'Ngăn lộ / Tủ', 'Chủng loại (VCB/ACB/SF6)',
  'Điện áp định mức (kV)', 'Dòng định mức (A)', 'Dòng cắt ngắn mạch (kA)',
  'R_tiếp xúc Pha R (µΩ)', 'R_tiếp xúc Pha Y/S (µΩ)', 'R_tiếp xúc Pha B/T (µΩ)',
  'Thời gian đóng (ms)', 'Thời gian cắt (ms)', 'R_cách điện cực mở (MΩ)', 'R_mạch điều khiển (MΩ)',
  'Áp suất SF6 (bar)', 'Trip Unit (LSIG)', 'Cơ cấu nạp lò xo', 'Đánh giá FSE', 'Ngày kiểm tra', 'Kỹ sư FSE'
];

export const TEV_MOTOR_HEADERS = [
  'Mã WO', 'Mã thiết bị', 'Tên động cơ', 'Khách hàng', 'Vị trí lắp đặt', 'Công suất (kW)',
  'Điện áp (V)', 'Dòng định mức (A)', 'Tốc độ (RPM)',
  'R_cách điện Stator (MΩ)', 'R_cách điện Rotor (MΩ)', 'Chỉ số PI', 'Chỉ số phóng điện DD',
  'Tổn hao Tan-Delta (%)', 'Độ rung DE (mm/s)', 'Độ rung NDE (mm/s)',
  'Nhiệt độ Stator (°C)', 'Nhiệt độ ổ bi (°C)', 'Lệch pha dòng (%)', 'Đánh giá FSE', 'Ngày kiểm tra', 'Kỹ sư FSE'
];

export const TEV_GENERATOR_HEADERS = [
  'Mã WO', 'Mã thiết bị', 'Tên máy phát', 'Khách hàng', 'Vị trí', 'Công suất (kVA)',
  'Điện áp phát (V)', 'Dòng định mức (A)', 'Hệ số cosφ', 'Tần số (Hz)', 'Tốc độ (RPM)',
  'R_cách điện Stator (MΩ)', 'R_cách điện Rotor (MΩ)', 'Chỉ số PI', 'Thời gian khởi động (s)',
  'Áp suất dầu nhớt (bar)', 'Nhiệt độ nước làm mát (°C)', 'Thử tải Load Bank (%)', 'Độ ổn định AVR',
  'Cắt quá tốc độ', 'Đánh giá FSE', 'Ngày kiểm tra', 'Kỹ sư FSE'
];

export const TEV_SWITCHGEAR_HEADERS = [
  'Mã WO', 'Mã thiết bị', 'Tên tủ điện', 'Khách hàng', 'Vị trí / Trạm', 'Cấp điện áp (kV)',
  'Dòng định mức Busbar (A)', 'R_tiếp xúc Busbar R (µΩ)', 'R_tiếp xúc Busbar Y (µΩ)', 'R_tiếp xúc Busbar B (µΩ)',
  'Mức TEV PD (dBmV)', 'Siêu âm Airborne (dBµV)', 'Mật độ xung (pps)',
  'R_cách điện thanh cái (MΩ)', 'R_cách điện nhị thứ (MΩ)', 'Khóa liên động Interlock', 'Quá nhiệt hồng ngoại IR (°C)',
  'Đánh giá FSE', 'Ngày kiểm tra', 'Kỹ sư FSE'
];

export const TEV_TRANSFORMER_HEADERS = [
  'Mã WO', 'Mã thiết bị', 'Tên máy biến áp', 'Khách hàng', 'Trạm / Vị trí', 'Công suất (kVA)',
  'Điện áp sơ cấp (kV)', 'Điện áp thứ cấp (V)', 'Tổ đấu dây', 'Nhiệt độ dầu (°C)', 'Nhiệt độ cuộn dây (°C)',
  'R_cách điện Cao-Hạ (MΩ)', 'R_cách điện Cao-Đất (MΩ)', 'R_cách điện Hạ-Đất (MΩ)', 'Chỉ số PI',
  'H2 (ppm)', 'CH4 (ppm)', 'C2H2 (ppm)', 'C2H4 (ppm)', 'C2H6 (ppm)', 'CO (ppm)', 'CO2 (ppm)',
  'Điện áp đánh thủng dầu (kV)', 'Hàm lượng ẩm (ppm)', 'Đánh giá FSE', 'Ngày kiểm tra', 'Kỹ sư FSE'
];

export const TEV_CABLE_HEADERS = [
  'Mã WO', 'Mã cáp', 'Tên tuyến cáp', 'Khách hàng', 'Điểm đầu - Điểm cuối', 'Cấp điện áp (kV)',
  'Tiết diện (mm²)', 'Chiều dài (m)', 'R_cách điện Pha A (MΩ)', 'R_cách điện Pha B (MΩ)', 'R_cách điện Pha C (MΩ)',
  'Thử nghiệm VLF/Hipot (kV)', 'Dòng rò rỉ (µA)', 'Màn chắn tiếp địa', 'Đánh giá FSE', 'Ngày kiểm tra', 'Kỹ sư FSE'
];

export const TEV_BATTERY_HEADERS = [
  'Mã WO', 'Mã dàn pin / UPS', 'Tên hệ thống', 'Khách hàng', 'Vị trí', 'Loại pin (VRLA/Flooded/Lithium)',
  'Dung lượng (Ah)', 'Điện áp tổng giàn (V)', 'Dòng nạp Float (A)', 'Nội trở bình R_int (mΩ)',
  'Áp bình thấp nhất (V)', 'Nhiệt độ cực đại (°C)', 'Thời gian thử tải (phút)', 'Đánh giá FSE', 'Ngày kiểm tra', 'Kỹ sư FSE'
];

export const TEV_RELAY_HEADERS = [
  'Mã WO', 'Mã rơle', 'Tên bảo vệ', 'Khách hàng', 'Ngăn lộ', 'Model / Hãng',
  'Chức năng kích hoạt (50/51/87...)', 'Dòng khởi động Pickup (A)', 'Thời gian tác động Trip (ms)',
  'Tỷ số biến dòng CT', 'Tiếp điểm cắt Trip', 'Đánh giá FSE', 'Ngày kiểm tra', 'Kỹ sư FSE'
];

// Helper to get spreadsheet ID from any request
const resolveSpreadsheetId = (req: express.Request): string => {
  const fromQuery = req.query.spreadsheetId as string;
  const fromBody = req.body?.spreadsheetId as string;
  const fromHeader = req.headers['x-spreadsheet-id'] as string;
  const fromEnv = process.env.SPREADSHEET_ID;
  return extractSpreadsheetId(fromQuery || fromBody || fromHeader || fromEnv);
};

// 4. Save Data to Google Sheets (Append)
app.post('/api/sheets/append', async (req, res) => {
  const tokens = getTokensFromRequest(req);
  if (!tokens) {
    return res.status(401).json({ error: 'Not authenticated with Google' });
  }

  const { range, values } = req.body;
  const spreadsheetId = resolveSpreadsheetId(req);

  if (!spreadsheetId) {
    return res.status(400).json({ error: 'SPREADSHEET_ID is not configured. Vui lòng cung cấp Spreadsheet ID hoặc link Google Sheet.' });
  }

  try {
    const oauth2Client = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET
    );
    oauth2Client.setCredentials(tokens);

    const sheets = google.sheets({ version: 'v4', auth: oauth2Client });

    const response = await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: range || 'TEV Service Flatform!A:W',
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: Array.isArray(values[0]) ? values : [values]
      }
    });

    res.json({ success: true, data: response.data });
  } catch (error: any) {
    console.error('Error appending to sheet:', error);
    const isAuthError = error.code === 401 || error.status === 401 || (error.response && error.response.status === 401) || (error.message && error.message.includes('invalid_grant'));
    if (isAuthError) {
      globalTokens = null;
      return res.status(401).json({ error: 'Phiên đăng nhập đã hết hạn. Vui lòng kết nối lại Google Drive.' });
    }
    res.status(500).json({ error: error.message || 'Failed to save to Google Sheets' });
  }
});

// 5. Get Data from Google Sheets (Two-Way Sync Read)
app.get('/api/sheets/get', async (req, res) => {
  const tokens = getTokensFromRequest(req);
  if (!tokens) {
    return res.status(401).json({ error: 'Not authenticated with Google' });
  }

  const spreadsheetId = resolveSpreadsheetId(req);
  if (!spreadsheetId) {
    return res.status(400).json({ error: 'SPREADSHEET_ID is not configured. Vui lòng nhập link hoặc ID Google Sheet.' });
  }

  try {
    const oauth2Client = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET
    );
    oauth2Client.setCredentials(tokens);

    const sheets = google.sheets({ version: 'v4', auth: oauth2Client });

    // Fetch spreadsheet metadata to get actual sheet names
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetNames = spreadsheet.data.sheets?.map(s => s.properties?.title || '') || [];
    
    if (sheetNames.length === 0) {
      return res.json({ success: true, data: { transformers: [], switchgears: [], motors: [], allSheets: {} } });
    }

    const response = await sheets.spreadsheets.values.batchGet({
      spreadsheetId,
      ranges: sheetNames.map(name => `'${name}'!A:AZ`),
    });

    const allData: Record<string, any[][]> = {};
    sheetNames.forEach((name, index) => {
      allData[name] = response.data.valueRanges?.[index]?.values || [];
    });

    // Categorize data for backward compatibility and fast 2-way sync
    const categorizedData = {
      tevServiceFlatform: [] as any[][],
      customers: [] as any[][],
      equipment: [] as any[][],
      transformers: [] as any[][],
      switchgears: [] as any[][],
      motors: [] as any[][],
      generators: [] as any[][],
      allSheets: allData,
      sheetNames
    };

    sheetNames.forEach(name => {
      const lowerName = name.toLowerCase();
      const rows = allData[name];
      if (lowerName.includes('tev service flatform') || lowerName.includes('work order') || lowerName.includes('cmms')) {
        categorizedData.tevServiceFlatform = categorizedData.tevServiceFlatform.concat(rows);
      } else if (lowerName.includes('khachhang') || lowerName.includes('khách hàng') || lowerName.includes('customer')) {
        categorizedData.customers = categorizedData.customers.concat(rows);
      } else if (lowerName.includes('thietbi') || lowerName.includes('thiết bị') || lowerName.includes('equipment')) {
        categorizedData.equipment = categorizedData.equipment.concat(rows);
      } else if (lowerName.includes('biến áp') || lowerName.includes('mba') || lowerName.includes('transformer')) {
        categorizedData.transformers = categorizedData.transformers.concat(rows);
      } else if (lowerName.includes('tủ điện') || lowerName.includes('trung thế') || lowerName.includes('switchgear') || lowerName.includes('máy cắt') || lowerName.includes('breaker')) {
        categorizedData.switchgears = categorizedData.switchgears.concat(rows);
      } else if (lowerName.includes('động cơ') || lowerName.includes('motor')) {
        categorizedData.motors = categorizedData.motors.concat(rows);
      } else if (lowerName.includes('máy phát') || lowerName.includes('generator')) {
        categorizedData.generators = categorizedData.generators.concat(rows);
      }
    });

    res.json({ 
      success: true, 
      spreadsheetTitle: spreadsheet.data.properties?.title || 'TEV Service Flatform',
      spreadsheetId,
      data: categorizedData 
    });
  } catch (error: any) {
    console.error('Error getting data from sheet:', error);
    const isAuthError = error.code === 401 || error.status === 401 || (error.response && error.response.status === 401) || (error.message && error.message.includes('invalid_grant'));
    if (isAuthError) {
      globalTokens = null;
      return res.status(401).json({ error: 'Phiên đăng nhập đã hết hạn. Vui lòng tải lại trang và kết nối lại Google Drive.' });
    }
    res.status(500).json({ error: error.message || 'Failed to get data from Google Sheets' });
  }
});

// 5b. Atomic 2-Way Sync Record (Sync a single Work Order, Customer, or Equipment directly to Sheets)
app.post('/api/sheets/sync-record', async (req, res) => {
  const tokens = getTokensFromRequest(req);
  if (!tokens) {
    return res.status(401).json({ error: 'Not authenticated with Google' });
  }

  const spreadsheetId = resolveSpreadsheetId(req);
  if (!spreadsheetId) {
    return res.status(400).json({ error: 'SPREADSHEET_ID is not configured.' });
  }

  const { recordType, data } = req.body;
  if (!recordType || !data) {
    return res.status(400).json({ error: 'Missing recordType or data in request body' });
  }

  try {
    const oauth2Client = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET
    );
    oauth2Client.setCredentials(tokens);

    const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const existingSheetTitles = spreadsheet.data.sheets?.map(s => s.properties?.title) || [];

    let targetSheetTitle = 'TEV Service Flatform';
    let targetHeaders = TEV_WO_HEADERS;
    let rowValues: any[] = [];
    let recordId = '';

    if (recordType === 'workOrder') {
      targetSheetTitle = existingSheetTitles.find(t => t?.toLowerCase().includes('tev service flatform') || t?.toLowerCase().includes('cmms')) || 'TEV Service Flatform';
      targetHeaders = TEV_WO_HEADERS;
      recordId = data.id || data.woCode || '';
      rowValues = [
        data.id || data.woCode || '',
        data.workPermitId || '',
        data.title || '',
        data.description || '',
        Array.isArray(data.equipmentId) ? data.equipmentId.join(', ') : (data.equipmentId || data.equipmentName || ''),
        data.failureCode || '',
        data.customer || data.customerName || data.customerId || '',
        data.type || '',
        data.isUnplanned ? 'Có' : 'Không',
        data.priority || '',
        data.status || '',
        data.assignedTo || '',
        data.responsibleApprove || '',
        data.responsibleDo || '',
        data.blockingRequired ? 'Có' : 'Không',
        data.dueDate || '',
        Array.isArray(data.usedMaterials) ? data.usedMaterials.map((m: any) => `${m.name} (${m.quantity} ${m.unit})`).join(', ') : (data.usedMaterials || ''),
        data.createdAt || new Date().toISOString(),
        data.updatedAt || new Date().toISOString(),
        data.downtimeStart || '',
        data.repairStart || '',
        data.repairEnd || '',
        data.restartTime || ''
      ];
    } else if (recordType === 'customer') {
      targetSheetTitle = existingSheetTitles.find(t => t?.toLowerCase().includes('khachhang') || t?.toLowerCase().includes('khách hàng')) || 'khachhang';
      targetHeaders = TEV_CUSTOMER_HEADERS;
      recordId = data.id || data.customerCode || '';
      rowValues = [
        data.id || data.customerCode || '',
        data.name || '',
        Array.isArray(data.factories) ? data.factories.join(', ') : (data.factories || ''),
        data.email || '',
        data.phone || '',
        data.address || '',
        data.createdAt || new Date().toISOString(),
        data.updatedAt || new Date().toISOString()
      ];
    } else if (recordType === 'breaker' || (recordType === 'equipmentTest' && data.equipmentType?.toLowerCase().includes('cắt'))) {
      targetSheetTitle = existingSheetTitles.find(t => t?.toLowerCase().includes('máy cắt') || t?.toLowerCase().includes('breaker')) || 'Máy cắt';
      targetHeaders = TEV_BREAKER_HEADERS;
      recordId = data.woId || data.id || '';
      rowValues = [
        data.woId || data.id || '',
        data.equipmentId || '',
        data.equipmentName || '',
        data.customer || '',
        data.location || data.factory || '',
        data.breakerType || 'VCB',
        data.ratedVoltageKv || '24',
        data.ratedCurrentA || '630',
        data.shortCircuitKa || '25',
        data.contactResistanceR || '',
        data.contactResistanceY || data.contactResistanceS || '',
        data.contactResistanceB || data.contactResistanceT || '',
        data.closeTimeMs || '',
        data.openTimeMs || '',
        data.insulationIrMo || '',
        data.controlCircuitIrMo || '',
        data.sf6GasPressureBar || '',
        data.tripUnitLsig || 'Đạt',
        data.springChargeMechanism || 'Bình thường',
        data.assessment || 'Đạt tiêu chuẩn NETA ATS',
        data.testDate || new Date().toLocaleDateString('vi-VN'),
        data.inspector || data.assignedTo || 'Kỹ sư FSE'
      ];
    } else if (recordType === 'motor' || (recordType === 'equipmentTest' && data.equipmentType?.toLowerCase().includes('động cơ'))) {
      targetSheetTitle = existingSheetTitles.find(t => t?.toLowerCase().includes('động cơ') || t?.toLowerCase().includes('motor')) || 'Động cơ';
      targetHeaders = TEV_MOTOR_HEADERS;
      recordId = data.woId || data.id || '';
      rowValues = [
        data.woId || data.id || '',
        data.equipmentId || '',
        data.equipmentName || '',
        data.customer || '',
        data.location || data.factory || '',
        data.ratedPowerKw || '',
        data.motorVoltageV || '400',
        data.motorCurrentA || '',
        data.motorRpm || '',
        data.motorIrMo || '',
        data.motorIrRotorMo || '',
        data.motorPi || '',
        data.motorDd || '',
        data.motorTanDeltaPct || '',
        data.vibrationDeMmS || '',
        data.vibrationNdeMmS || '',
        data.statorTempC || '',
        data.bearingTempC || '',
        data.phaseImbalancePct || '',
        data.assessment || 'Đạt tiêu chuẩn vận hành',
        data.testDate || new Date().toLocaleDateString('vi-VN'),
        data.inspector || data.assignedTo || 'Kỹ sư FSE'
      ];
    } else if (recordType === 'generator' || (recordType === 'equipmentTest' && (data.equipmentType?.toLowerCase().includes('máy phát') || data.equipmentType?.toLowerCase().includes('generator')))) {
      targetSheetTitle = existingSheetTitles.find(t => t?.toLowerCase().includes('máy phát') || t?.toLowerCase().includes('generator')) || 'Máy phát';
      targetHeaders = TEV_GENERATOR_HEADERS;
      recordId = data.woId || data.id || '';
      rowValues = [
        data.woId || data.id || '',
        data.equipmentId || '',
        data.equipmentName || '',
        data.customer || '',
        data.location || data.factory || '',
        data.genPowerKva || '',
        data.genVoltageV || '400',
        data.genCurrentA || '',
        data.genCosPhi || '0.8',
        data.genHz || '50.0',
        data.genRpm || '1500',
        data.genIrStatorMo || '',
        data.genIrRotorMo || '',
        data.genPi || '',
        data.genStartSec || '',
        data.genOilPressureBar || '',
        data.genWaterTempC || '',
        data.genLoadBankPct || '',
        data.genAvrStability || 'Ổn định ±1%',
        data.genOverspeedTrip || 'Cắt 115% đạt',
        data.assessment || 'Sẵn sàng khởi động khẩn cấp (NFPA 110)',
        data.testDate || new Date().toLocaleDateString('vi-VN'),
        data.inspector || data.assignedTo || 'Kỹ sư FSE'
      ];
    } else if (recordType === 'switchgear' || (recordType === 'equipmentTest' && data.equipmentType?.toLowerCase().includes('tủ điện'))) {
      targetSheetTitle = existingSheetTitles.find(t => t?.toLowerCase().includes('tủ điện') || t?.toLowerCase().includes('switchgear')) || 'Tủ điện';
      targetHeaders = TEV_SWITCHGEAR_HEADERS;
      recordId = data.woId || data.id || '';
      rowValues = [
        data.woId || data.id || '',
        data.equipmentId || '',
        data.equipmentName || '',
        data.customer || '',
        data.location || data.factory || '',
        data.swgVoltageKv || '24',
        data.swgBusCurrentA || '2500',
        data.swgBusResistanceR || '',
        data.swgBusResistanceY || '',
        data.swgBusResistanceB || '',
        data.swgTevDbmv || '',
        data.swgUltrasonicDbuv || '',
        data.swgPulsePps || '',
        data.swgControlIrMo || '',
        data.swgSecondaryIrMo || '',
        data.swgInterlockStatus || 'Tốt / Đầy đủ',
        data.swgIrTempC || '',
        data.assessment || 'Đạt tiêu chuẩn phóng điện cục bộ TEV',
        data.testDate || new Date().toLocaleDateString('vi-VN'),
        data.inspector || data.assignedTo || 'Kỹ sư FSE'
      ];
    } else if (recordType === 'transformer' || (recordType === 'equipmentTest' && data.equipmentType?.toLowerCase().includes('biến áp'))) {
      targetSheetTitle = existingSheetTitles.find(t => t?.toLowerCase().includes('biến áp') || t?.toLowerCase().includes('transformer')) || 'Máy biến áp';
      targetHeaders = TEV_TRANSFORMER_HEADERS;
      recordId = data.woId || data.id || '';
      rowValues = [
        data.woId || data.id || '',
        data.equipmentId || '',
        data.equipmentName || '',
        data.customer || '',
        data.location || data.factory || '',
        data.trfPowerKva || '',
        data.trfPriVoltageKv || '22',
        data.trfSecVoltageV || '400',
        data.trfVectorGroup || 'Dyn11',
        data.trfOilTempC || '',
        data.trfWindingTempC || '',
        data.trfIrHighLowMo || '',
        data.trfIrHighEarthMo || '',
        data.trfIrLowEarthMo || '',
        data.trfPi || '',
        data.trfDgaH2 || '',
        data.trfDgaCh4 || '',
        data.trfDgaC2h2 || '',
        data.trfDgaC2h4 || '',
        data.trfDgaC2h6 || '',
        data.trfDgaCo || '',
        data.trfDgaCo2 || '',
        data.trfBreakdownKv || '',
        data.trfMoisturePpm || '',
        data.assessment || 'Đạt chuẩn NETA ATS / IEEE C57.104',
        data.testDate || new Date().toLocaleDateString('vi-VN'),
        data.inspector || data.assignedTo || 'Kỹ sư FSE'
      ];
    } else if (recordType === 'cable' || (recordType === 'equipmentTest' && data.equipmentType?.toLowerCase().includes('cáp'))) {
      targetSheetTitle = existingSheetTitles.find(t => t?.toLowerCase().includes('cáp') || t?.toLowerCase().includes('cable')) || 'Cáp điện';
      targetHeaders = TEV_CABLE_HEADERS;
      recordId = data.woId || data.id || '';
      rowValues = [
        data.woId || data.id || '',
        data.equipmentId || '',
        data.equipmentName || '',
        data.customer || '',
        data.location || data.route || '',
        data.voltageKv || '',
        data.crossSectionMm2 || '',
        data.cableLengthM || '',
        data.irPhaseAGroundMo || '',
        data.irPhaseBGroundMo || '',
        data.irPhaseCGroundMo || '',
        data.vlfHipotKv || '',
        data.leakageCurrentUa || '',
        data.shieldGrounding || 'Đạt chuẩn',
        data.assessment || 'Đạt chuẩn thử nghiệm VLF Mục 7.3.3',
        data.testDate || new Date().toLocaleDateString('vi-VN'),
        data.inspector || data.assignedTo || 'Kỹ sư FSE'
      ];
    } else if (recordType === 'battery' || (recordType === 'equipmentTest' && (data.equipmentType?.toLowerCase().includes('pin') || data.equipmentType?.toLowerCase().includes('ups') || data.equipmentType?.toLowerCase().includes('bess') || data.equipmentType?.toLowerCase().includes('ắc quy')))) {
      targetSheetTitle = existingSheetTitles.find(t => t?.toLowerCase().includes('pin') || t?.toLowerCase().includes('ups') || t?.toLowerCase().includes('bess')) || 'Pin & UPS';
      targetHeaders = TEV_BATTERY_HEADERS;
      recordId = data.woId || data.id || '';
      rowValues = [
        data.woId || data.id || '',
        data.equipmentId || '',
        data.equipmentName || '',
        data.customer || '',
        data.location || '',
        data.batteryType || 'VRLA',
        data.capacityAh || '',
        data.totalVoltageV || '',
        data.floatCurrentA || '',
        data.internalResistanceMohm || '',
        data.minCellVoltageV || '',
        data.maxTempC || '',
        data.dischargeTimeMin || '',
        data.assessment || 'Đạt chuẩn IEEE 1188 / 450',
        data.testDate || new Date().toLocaleDateString('vi-VN'),
        data.inspector || data.assignedTo || 'Kỹ sư FSE'
      ];
    } else if (recordType === 'relay' || (recordType === 'equipmentTest' && data.equipmentType?.toLowerCase().includes('rơ le'))) {
      targetSheetTitle = existingSheetTitles.find(t => t?.toLowerCase().includes('rơ le') || t?.toLowerCase().includes('relay')) || 'Rơ le bảo vệ';
      targetHeaders = TEV_RELAY_HEADERS;
      recordId = data.woId || data.id || '';
      rowValues = [
        data.woId || data.id || '',
        data.equipmentId || '',
        data.equipmentName || '',
        data.customer || '',
        data.location || '',
        data.model || '',
        data.functions || '50/51/50N/51N',
        data.pickupCurrentA || '',
        data.tripTimeMs || '',
        data.ctRatioResult || 'Đạt',
        data.tripContactStatus || 'Đạt',
        data.assessment || 'Đạt tiêu chuẩn bảo vệ rơle số',
        data.testDate || new Date().toLocaleDateString('vi-VN'),
        data.inspector || data.assignedTo || 'Kỹ sư FSE'
      ];
    }

    // Ensure target sheet exists
    if (!existingSheetTitles.includes(targetSheetTitle)) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [{
            addSheet: { properties: { title: targetSheetTitle } }
          }]
        }
      });
      // Add headers
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${targetSheetTitle}!A1`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [targetHeaders] }
      });
    }

    // Check if row already exists by record ID in column A
    const currentRowsRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${targetSheetTitle}!A:A`
    });
    const currentRows = currentRowsRes.data.values || [];
    let existingRowIndex = -1;

    if (recordId) {
      for (let r = 1; r < currentRows.length; r++) {
        const val = currentRows[r]?.[0]?.toString().trim();
        if (val && val.toLowerCase() === recordId.toLowerCase()) {
          existingRowIndex = r + 1; // 1-indexed row number in sheets
          break;
        }
      }
    }

    if (existingRowIndex > 0) {
      // Update existing row
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${targetSheetTitle}!A${existingRowIndex}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [rowValues] }
      });
      res.json({ success: true, action: 'updated', row: existingRowIndex, sheet: targetSheetTitle });
    } else {
      // If header is missing, write header first
      if (currentRows.length === 0) {
        await sheets.spreadsheets.values.update({
          spreadsheetId,
          range: `${targetSheetTitle}!A1`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [targetHeaders] }
        });
      }
      // Append row
      const appendRes = await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: `${targetSheetTitle}!A:Z`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [rowValues] }
      });
      res.json({ success: true, action: 'appended', sheet: targetSheetTitle, details: appendRes.data });
    }
  } catch (error: any) {
    console.error('Error syncing record to sheet:', error);
    res.status(500).json({ error: error.message || 'Failed to sync record to Google Sheets' });
  }
});

// 6. Sync/Export Full Data to Google Sheets
app.post('/api/sheets/sync-export', async (req, res) => {
  const tokens = getTokensFromRequest(req);
  if (!tokens) {
    return res.status(401).json({ error: 'Not authenticated with Google' });
  }

  const { 
    transformers, 
    switchgears, 
    motors, 
    generators,
    cmmsData, 
    tevServiceFlatformData,
    customersData, 
    equipmentData,
    inventoryData 
  } = req.body;
  
  const spreadsheetId = resolveSpreadsheetId(req);

  if (!spreadsheetId) {
    return res.status(400).json({ error: 'SPREADSHEET_ID is not configured. Vui lòng cung cấp link hoặc ID Google Sheet.' });
  }

  try {
    const oauth2Client = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET
    );
    oauth2Client.setCredentials(tokens);

    const sheets = google.sheets({ version: 'v4', auth: oauth2Client });

    // Target sheets - All equipment sheets corresponding to TEV Service Platform
    const coreSheets = [
      { key: 'tev', defaultTitle: 'TEV Service Flatform', keywords: ['tev service flatform', 'work order', 'cmms'], headers: TEV_WO_HEADERS },
      { key: 'khachhang', defaultTitle: 'khachhang', keywords: ['khachhang', 'khách hàng', 'customer'], headers: TEV_CUSTOMER_HEADERS },
      { key: 'thietbi', defaultTitle: 'thietbi', keywords: ['thietbi', 'thiết bị', 'equipment'], headers: TEV_EQUIPMENT_HEADERS },
      { key: 'maycat', defaultTitle: 'Máy cắt', keywords: ['máy cắt', 'breaker', 'vcb', 'acb'], headers: TEV_BREAKER_HEADERS },
      { key: 'dongco', defaultTitle: 'Động cơ', keywords: ['động cơ', 'motor'], headers: TEV_MOTOR_HEADERS },
      { key: 'mayphat', defaultTitle: 'Máy phát', keywords: ['máy phát', 'generator'], headers: TEV_GENERATOR_HEADERS },
      { key: 'tudien', defaultTitle: 'Tủ điện', keywords: ['tủ điện', 'switchgear', 'trung thế', 'rmu'], headers: TEV_SWITCHGEAR_HEADERS },
      { key: 'mba', defaultTitle: 'Máy biến áp', keywords: ['biến áp', 'mba', 'transformer'], headers: TEV_TRANSFORMER_HEADERS },
      { key: 'capdien', defaultTitle: 'Cáp điện', keywords: ['cáp điện', 'cable'], headers: TEV_CABLE_HEADERS },
      { key: 'pin_ups', defaultTitle: 'Pin & UPS', keywords: ['pin', 'ups', 'bess', 'battery', 'ắc quy'], headers: TEV_BATTERY_HEADERS },
      { key: 'role', defaultTitle: 'Rơ le bảo vệ', keywords: ['rơ le', 'relay'], headers: TEV_RELAY_HEADERS },
      { key: 'kho', defaultTitle: 'QuanLyKho', keywords: ['quanlykho', 'kho', 'inventory'], headers: ['Mã VT', 'Tên vật tư', 'Loại', 'Số lượng tồn', 'Đơn vị', 'Đơn giá', 'Vị trí kho', 'Tối thiểu', 'Ghi chú'] }
    ];

    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const existingSheetTitles = spreadsheet.data.sheets?.map(s => s.properties?.title || '') || [];

    const resolvedSheets = coreSheets.map(s => {
      const match = existingSheetTitles.find(t => s.keywords.some(k => t.toLowerCase().includes(k)));
      return { ...s, title: match || s.defaultTitle };
    });

    const missingSheets = resolvedSheets.filter(s => !existingSheetTitles.includes(s.title));
    if (missingSheets.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: missingSheets.map(s => ({
            addSheet: { properties: { title: s.title } }
          }))
        }
      });
    }

    // Prepare batch update values
    const dataToWrite: any[] = [];
    const actualWoData = tevServiceFlatformData || cmmsData;

    if (actualWoData && actualWoData.length > 0) {
      dataToWrite.push({ range: `${resolvedSheets[0].title}!A1`, values: actualWoData });
    }
    if (customersData && customersData.length > 0) {
      dataToWrite.push({ range: `${resolvedSheets[1].title}!A1`, values: customersData });
    }
    if (equipmentData && equipmentData.length > 0) {
      dataToWrite.push({ range: `${resolvedSheets[2].title}!A1`, values: equipmentData });
    }
    const { breakersData, cablesData, batteriesData, relaysData } = req.body;
    if (breakersData && breakersData.length > 0) {
      dataToWrite.push({ range: `${resolvedSheets[3].title}!A1`, values: breakersData });
    }
    if (motors && motors.length > 0) {
      dataToWrite.push({ range: `${resolvedSheets[4].title}!A1`, values: motors });
    }
    if (generators && generators.length > 0) {
      dataToWrite.push({ range: `${resolvedSheets[5].title}!A1`, values: generators });
    }
    if (switchgears && switchgears.length > 0) {
      dataToWrite.push({ range: `${resolvedSheets[6].title}!A1`, values: switchgears });
    }
    if (transformers && transformers.length > 0) {
      dataToWrite.push({ range: `${resolvedSheets[7].title}!A1`, values: transformers });
    }
    if (cablesData && cablesData.length > 0) {
      dataToWrite.push({ range: `${resolvedSheets[8].title}!A1`, values: cablesData });
    }
    if (batteriesData && batteriesData.length > 0) {
      dataToWrite.push({ range: `${resolvedSheets[9].title}!A1`, values: batteriesData });
    }
    if (relaysData && relaysData.length > 0) {
      dataToWrite.push({ range: `${resolvedSheets[10].title}!A1`, values: relaysData });
    }
    if (inventoryData && inventoryData.length > 0) {
      dataToWrite.push({ range: `${resolvedSheets[11].title}!A1`, values: inventoryData });
    }

    if (dataToWrite.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        requestBody: {
          valueInputOption: 'USER_ENTERED',
          data: dataToWrite
        }
      });
    }

    res.json({ 
      success: true, 
      message: 'Đồng bộ toàn diện lên Google Sheet thành công!',
      syncedSheets: dataToWrite.map(d => d.range.split('!')[0])
    });
  } catch (error: any) {
    console.error('Error syncing to sheet:', error);
    const isAuthError = error.code === 401 || error.status === 401 || (error.response && error.response.status === 401) || (error.message && error.message.includes('invalid_grant'));
    if (isAuthError) {
      globalTokens = null;
      return res.status(401).json({ error: 'Phiên đăng nhập đã hết hạn. Vui lòng kết nối lại Google Drive.' });
    }
    res.status(500).json({ error: error.message || 'Failed to sync to Google Sheets' });
  }
});

// 6b. Initialize All Equipment Sheets on TEV Service Flatform
app.post('/api/sheets/init-tev-sheets', async (req, res) => {
  const tokens = getTokensFromRequest(req);
  if (!tokens) {
    return res.status(401).json({ error: 'Not authenticated with Google' });
  }

  const spreadsheetId = resolveSpreadsheetId(req);
  if (!spreadsheetId) {
    return res.status(400).json({ error: 'SPREADSHEET_ID is not configured.' });
  }

  try {
    const oauth2Client = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET
    );
    oauth2Client.setCredentials(tokens);

    const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const existingSheets = spreadsheet.data.sheets || [];
    const existingTitles = existingSheets.map(s => s.properties?.title || '');

    const standardSheets = [
      { title: 'TEV Service Flatform', headers: TEV_WO_HEADERS },
      { title: 'khachhang', headers: TEV_CUSTOMER_HEADERS },
      { title: 'thietbi', headers: TEV_EQUIPMENT_HEADERS },
      { title: 'Máy cắt', headers: TEV_BREAKER_HEADERS },
      { title: 'Động cơ', headers: TEV_MOTOR_HEADERS },
      { title: 'Máy phát', headers: TEV_GENERATOR_HEADERS },
      { title: 'Tủ điện', headers: TEV_SWITCHGEAR_HEADERS },
      { title: 'Máy biến áp', headers: TEV_TRANSFORMER_HEADERS },
      { title: 'Cáp điện', headers: TEV_CABLE_HEADERS },
      { title: 'Pin & UPS', headers: TEV_BATTERY_HEADERS },
      { title: 'Rơ le bảo vệ', headers: TEV_RELAY_HEADERS },
      { title: 'QuanLyKho', headers: ['Mã VT', 'Tên vật tư', 'Loại', 'Số lượng tồn', 'Đơn vị', 'Đơn giá', 'Vị trí kho', 'Tối thiểu', 'Ghi chú'] }
    ];

    const missingSheets = standardSheets.filter(s => !existingTitles.includes(s.title));
    if (missingSheets.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: missingSheets.map(s => ({
            addSheet: { properties: { title: s.title } }
          }))
        }
      });
    }

    // Write headers for any sheets where row 1 is empty or initialize them
    const headerUpdates: any[] = [];
    for (const sheetDef of standardSheets) {
      headerUpdates.push({
        range: `${sheetDef.title}!A1`,
        values: [sheetDef.headers]
      });
    }

    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: headerUpdates
      }
    });

    res.json({
      success: true,
      message: `Đã khởi tạo & chuẩn hóa ${standardSheets.length} sheet thiết bị thành công trên file TEV Service Flatform!`,
      sheets: standardSheets.map(s => s.title)
    });
  } catch (error: any) {
    console.error('Error initializing sheets:', error);
    res.status(500).json({ error: error.message || 'Failed to initialize TEV sheets' });
  }
});

// 7. Upload files to Google Drive
app.post('/api/drive/upload', upload.array('files'), async (req, res) => {
  const tokens = getTokensFromRequest(req);
  if (!tokens) {
    return res.status(401).json({ error: 'Not authenticated with Google' });
  }

  try {
    const oauth2Client = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET
    );
    oauth2Client.setCredentials(tokens);

    const drive = google.drive({ version: 'v3', auth: oauth2Client });
    const folderName = 'TEV_Equipment_Reports';
    let folderId = '';

    // 1. Find or create the folder
    const resFolder = await drive.files.list({
      q: `mimeType='application/vnd.google-apps.folder' and name='${folderName}' and trashed=false`,
      fields: 'files(id, name)',
    });

    if (resFolder.data.files && resFolder.data.files.length > 0) {
      folderId = resFolder.data.files[0].id!;
    } else {
      const folderMetadata = {
        name: folderName,
        mimeType: 'application/vnd.google-apps.folder',
      };
      const folder = await drive.files.create({
        requestBody: folderMetadata,
        fields: 'id',
      });
      folderId = folder.data.id!;
    }

    // 2. Upload files
    const files = req.files as Express.Multer.File[];
    const uploadedLinks = [];

    for (const file of files) {
      const bufferStream = new Readable();
      bufferStream.push(file.buffer);
      bufferStream.push(null);

      const fileMetadata = {
        name: file.originalname,
        parents: [folderId],
      };
      const media = {
        mimeType: file.mimetype,
        body: bufferStream,
      };

      const uploadedFile = await drive.files.create({
        requestBody: fileMetadata,
        media: media,
        fields: 'id, webViewLink',
      });

      // Make it readable by anyone with link
      await drive.permissions.create({
        fileId: uploadedFile.data.id!,
        requestBody: {
          role: 'reader',
          type: 'anyone',
        },
      });

      uploadedLinks.push(uploadedFile.data.webViewLink);
    }

    res.json({ success: true, links: uploadedLinks });
  } catch (error: any) {
    console.error('Drive upload error:', error);
    const isAuthError = error.code === 401 || error.status === 401 || (error.response && error.response.status === 401) || (error.message && error.message.includes('invalid_grant'));
    if (isAuthError) {
      console.log('Clearing globalTokens due to auth error');
      globalTokens = null;
      return res.status(401).json({ error: 'Phiên đăng nhập đã hết hạn. Vui lòng tải lại trang và kết nối lại Google Drive.' });
    }
    res.status(500).json({ error: error.message || 'Failed to upload files to Google Drive' });
  }
});

// ==========================================
// PREVENTIVE MAINTENANCE (PM) EMAIL ALERTS
// ==========================================

// 1. Get PM Notification configuration & SMTP status
app.get('/api/notifications/pm-config', (req, res) => {
  const isSmtpConfigured = !!(process.env.SMTP_HOST && process.env.SMTP_USER && process.env.SMTP_PASS);
  res.json({
    smtpConfigured: isSmtpConfigured,
    smtpHost: process.env.SMTP_HOST || null,
    smtpPort: process.env.SMTP_PORT || '587',
    smtpUser: process.env.SMTP_USER ? `${process.env.SMTP_USER.slice(0, 3)}***` : null,
    defaultRecipient: process.env.DEFAULT_ALERT_EMAIL || 'sgm1707@gmail.com',
    senderAddress: process.env.NOTIFICATION_EMAIL_FROM || '"Renewable CMMS" <notifications@cmms-system.com>',
    targetDaysNotice: 3,
    cloudFunctionAvailable: true,
  });
});

// 2. Scan and return upcoming PM tasks (due in 3 days)
app.post('/api/notifications/pm-upcoming', (req, res) => {
  try {
    const { workOrders = [], allEquipment = [], targetDays = 3 } = req.body;
    const upcomingTasks = findUpcomingPMOrders(workOrders, allEquipment, Number(targetDays));
    
    res.json({
      success: true,
      targetDays,
      totalUpcoming: upcomingTasks.length,
      tasks: upcomingTasks,
    });
  } catch (error: any) {
    console.error('Error scanning upcoming PM tasks:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 3. Dispatch automated PM email notifications
app.post('/api/notifications/send-pm-alerts', async (req, res) => {
  try {
    const { workOrders = [], allEquipment = [], tasks = null, forceSend = false, customRecipient } = req.body;
    
    // Use provided tasks or calculate them on the fly
    const upcomingTasks = tasks || findUpcomingPMOrders(workOrders, allEquipment, 3);

    if (upcomingTasks.length === 0) {
      return res.json({
        success: true,
        message: 'Không có phiếu bảo trì định kỳ nào sắp đến hạn trong vòng 3 ngày.',
        totalTasks: 0,
        results: [],
      });
    }

    const dispatchSummary = await sendPMNotifications(upcomingTasks, {
      forceSend,
      customRecipient,
    });

    res.json({
      success: true,
      message: `Đã xử lý thông báo cho ${dispatchSummary.processedCount} phiếu bảo trì định kỳ.`,
      ...dispatchSummary,
    });
  } catch (error: any) {
    console.error('Error sending PM alerts:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 4. Test send a sample PM notification email
app.post('/api/notifications/test-email', async (req, res) => {
  try {
    const { recipient = 'sgm1707@gmail.com' } = req.body;
    const sampleHtml = generatePMEmailHtml({
      orderId: 'WO-PM-TEST-001',
      title: 'Bảo trì định kỳ Máy biến áp MBA-T1 (Test Notification)',
      equipmentName: 'Máy biến áp chính 110kV MBA-T1',
      equipmentId: 'EQ-TR-01',
      factory: 'Trạm biến áp 110kV Nhơn Hội',
      dueDateStr: new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toLocaleDateString('vi-VN'),
      daysRemaining: 3,
      priority: 'high',
      assignedTo: 'Kỹ sư Trưởng ca / SGM',
      description: 'Kiểm tra mức dầu, nhiệt độ cuộn dây, lấy mẫu dầu DGA định kỳ và siết lực bu lông đầu cực.',
      pmFrequency: '3-months',
    });

    const transporter = getEmailTransporter();
    const fromAddress = process.env.NOTIFICATION_EMAIL_FROM || '"Renewable CMMS" <notifications@cmms-system.com>';
    const subject = `[THỬ NGHIỆM HỆ THỐNG - CÒN 3 NGÀY] Bảo trì định kỳ MBA-T1 (Hạn: ${new Date(Date.now() + 3 * 24 * 60 * 60 * 1000).toLocaleDateString('vi-VN')})`;

    if (transporter) {
      const info = await transporter.sendMail({
        from: fromAddress,
        to: recipient,
        subject,
        html: sampleHtml,
      });

      res.json({
        success: true,
        mode: 'sent',
        messageId: info.messageId,
        recipient,
        message: `Đã gửi email thử nghiệm thành công tới ${recipient}!`,
      });
    } else {
      res.json({
        success: true,
        mode: 'simulated',
        recipient,
        subject,
        htmlPreview: sampleHtml,
        message: `Chế độ mô phỏng: Email mẫu đã được tạo thành công cho ${recipient}. Hãy cấu hình SMTP (SMTP_HOST, SMTP_USER, SMTP_PASS) để gửi trực tiếp vào hòm thư.`,
      });
    }
  } catch (error: any) {
    console.error('Error sending test email:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 5. NETA ATS-2025 Dry-Type Low-Voltage Transformer AI Evaluation
app.post('/api/field-service/neta-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({ success: false, error: 'Thiếu dữ liệu kiểm tra điện hoặc thị giác theo NETA ATS-2025.' });
    }

    const evaluation = await generateNetaAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 6. NETA ATS-2025 Dry-Type Large Transformer AI Evaluation (Section 7.2.1.2)
app.post('/api/field-service/neta-large-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaLargeInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({ success: false, error: 'Thiếu dữ liệu kiểm tra điện hoặc thị giác theo NETA ATS-2025 Mục 7.2.1.2.' });
    }

    const evaluation = await generateNetaLargeAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 Large Transformer data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 7. NETA ATS-2025 BESS (Battery Energy Storage Systems) AI Evaluation (Section 7.28)
app.post('/api/field-service/neta-bess-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaBessInputPayload;
    if (!payload || !payload.electrical_and_subsystem_tests || !payload.visual_inspection) {
      return res.status(400).json({ success: false, error: 'Thiếu dữ liệu kiểm tra hệ thống BESS theo NETA ATS-2025 Mục 7.28.' });
    }

    const evaluation = await generateNetaBessAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 BESS data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 8. NETA ATS-2025 Solar Photovoltaic (PV) Systems AI Evaluation (Section 7.29)
app.post('/api/field-service/neta-pv-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaPvInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({ success: false, error: 'Thiếu dữ liệu kiểm tra hệ thống Solar PV theo NETA ATS-2025 Mục 7.29.' });
    }

    const evaluation = await generateNetaPvAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 Solar PV data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 8b. Solar PV Cell Thermal IR & Visual RGB Analysis (IEC 62446-3 & Volateq / Sitemark)
app.post('/api/field-service/pv-cell-thermal-analyze', async (req, res) => {
  try {
    const payload = req.body as PvThermalInputPayload;
    if (!payload) {
      return res.status(400).json({ success: false, error: 'Thiếu dữ liệu phân tích nhiệt PV Cell.' });
    }

    const result = await analyzePvCellWithGemini(payload);
    res.json(result);
  } catch (error: any) {
    console.error('Error in PV Cell Thermal Analysis:', error);
    res.status(500).json({ success: false, error: error?.message || 'Lỗi server khi phân tích ảnh nhiệt PV Cell' });
  }
});

// 9. NETA ATS-2025 Electric Vehicle Charging Systems (EVSE) AI Evaluation (Section 7.26)
app.post('/api/field-service/neta-evse-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaEvseInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({ success: false, error: 'Thiếu dữ liệu kiểm tra hệ thống trạm sạc xe điện EVSE theo NETA ATS-2025 Mục 7.26.' });
    }

    const evaluation = await generateNetaEvseAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 EVSE data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 10. NETA ATS-2025 Shielded Cables, Medium- and High-Voltage AI Evaluation (Section 7.3.3)
app.post('/api/field-service/neta-cable-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaCableInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({ success: false, error: 'Thiếu dữ liệu kiểm tra hệ thống cáp trung/cao áp có màn chắn theo NETA ATS-2025 Mục 7.3.3.' });
    }

    const evaluation = await generateNetaCableAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 Cable data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 10b. NETA ATS-2025 Cables, Low-Voltage, 1,000-Volt Maximum AI Evaluation (Section 7.3.2)
app.post('/api/field-service/neta-cable-lv-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaCableLvInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({ success: false, error: 'Thiếu dữ liệu kiểm tra cáp điện hạ áp theo NETA ATS-2025 Mục 7.3.2.' });
    }

    const evaluation = await generateNetaCableLvAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 Low-Voltage Cable data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 10c. NETA ATS-2025 Low-Voltage Power Circuit Breakers AI Evaluation (Section 7.6.1.2)
app.post('/api/field-service/neta-lv-breaker-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaLvBreakerInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({ success: false, error: 'Thiếu dữ liệu kiểm tra máy cắt hạ áp theo NETA ATS-2025 Mục 7.6.1.2.' });
    }

    const evaluation = await generateNetaLvBreakerAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 LV Breaker data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 10d. NETA ATS-2025 Grounding Systems AI Evaluation (Section 7.13 & IEEE Std 81)
app.post('/api/field-service/neta-grounding-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaGroundingInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({ success: false, error: 'Thiếu dữ liệu kiểm tra hệ thống nối đất theo NETA ATS-2025 Mục 7.13 & IEEE Std 81.' });
    }

    const evaluation = await generateNetaGroundingAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 Grounding data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 11. NETA ATS-2025 Switches, SF6, Medium-Voltage AI Evaluation (Section 7.5.4)
app.post('/api/field-service/neta-sf6-switch-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaSf6SwitchInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({ success: false, error: 'Thiếu dữ liệu kiểm tra cầu dao ngắt mạch trung áp SF6 theo NETA ATS-2025 Mục 7.5.4.' });
    }

    const evaluation = await generateNetaSf6SwitchAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 SF6 Switch data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 12. NETA ATS-2025 Thermographic Survey AI Evaluation (Section 9 & Table 100.18)
app.post('/api/field-service/neta-thermography-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaThermographyInputPayload;
    if (!payload || !payload.survey_conditions || !payload.thermal_measurements) {
      return res.status(400).json({ success: false, error: 'Thiếu dữ liệu khảo sát nhiệt độ hồng ngoại theo NETA ATS-2025 Mục 9 & Table 100.18.' });
    }

    const evaluation = await generateNetaThermographyAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 Thermography data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 13. NETA ATS-2025 Online Partial Discharge Survey AI Evaluation (Section 11 & Table 100.23)
app.post('/api/field-service/neta-pd-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaPdInputPayload;
    if (!payload || !payload.survey_conditions || !payload.pd_sensor_measurements) {
      return res.status(400).json({ success: false, error: 'Thiếu dữ liệu khảo sát phóng điện cục bộ theo NETA ATS-2025 Mục 11 & Table 100.23.' });
    }

    const evaluation = await generateNetaPdAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 PD data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 14. NETA ATS-2025 Rotating Machinery AC Induction Motors & Generators (Section 7.15.1 & Tables 100.10, 100.11, 100.12)
app.post('/api/field-service/neta-motor-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaMotorInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({ success: false, error: 'Thiếu dữ liệu kiểm tra động cơ/máy phát theo NETA ATS-2025 Mục 7.15.1.' });
    }

    const evaluation = await generateNetaMotorAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 Motor data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 14b. NETA ATS-2025 Rotating Machinery DC Motors & Generators (Section 7.15.3 & Tables 100.10, 100.11, 100.12)
app.post('/api/field-service/neta-dc-motor-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaDcMotorInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({
        success: false,
        error: 'Thiếu dữ liệu kiểm tra động cơ/máy phát điện DC theo NETA ATS-2025 Mục 7.15.3.',
      });
    }

    const evaluation = await generateNetaDcMotorAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 DC Motor data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 13. NETA ATS-2025 Liquid-Filled Transformers AI Evaluation (Section 7.2.2)
app.post('/api/field-service/neta-liquid-transformer-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaLiquidTransformerInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({
        success: false,
        error: 'Thiếu dữ liệu kiểm tra máy biến áp ngâm dầu theo NETA ATS-2025 Mục 7.2.2.',
      });
    }

    const evaluation = await generateNetaLiquidTransformerAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 Liquid Transformer data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 14. NETA ATS-2025 Microprocessor-Based Protective Relay Analysis Endpoint
app.post('/api/field-service/neta-relay-analysis', async (req, res) => {
  try {
    const payload = req.body as NetaRelayInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({
        success: false,
        error: 'Thiếu dữ liệu kiểm tra rơ-le vi xử lý theo NETA ATS-2025 Mục 7.9.2.',
      });
    }

    const evaluation = await generateNetaRelayAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 Protective Relay data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 15. NETA ATS-2025 Uninterruptible Power Systems (UPS) Analysis Endpoint
app.post('/api/field-service/neta-ups-analysis', async (req, res) => {
  try {
    const payload = req.body as NetaUpsInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({
        success: false,
        error: 'Thiếu dữ liệu kiểm tra hệ thống nguồn liên tục UPS theo NETA ATS-2025 Mục 7.22.2.',
      });
    }

    const evaluation = await generateNetaUpsAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 UPS data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 16. NETA ATS-2025 Synchronous Motors and Generators Analysis Endpoint
app.post('/api/field-service/neta-sync-machinery-analysis', async (req, res) => {
  try {
    const payload = req.body as NetaSyncMachineryInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({
        success: false,
        error: 'Thiếu dữ liệu kiểm tra máy điện đồng bộ theo NETA ATS-2025 Mục 7.15.2.',
      });
    }

    const evaluation = await generateNetaSyncMachineryAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 Synchronous Machinery data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 17. NETA ATS-2025 Direct-Current Systems, Batteries, Flooded Lead-Acid Analysis Endpoint (Section 7.18.1.1 & IEEE 450)
app.post('/api/field-service/neta-battery-flooded-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaBatteryFloodedInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({
        success: false,
        error: 'Thiếu dữ liệu kiểm tra hệ thống DC & ắc quy Flooded Lead-Acid theo NETA ATS-2025 Mục 7.18.1.1.',
      });
    }

    const evaluation = await generateNetaBatteryFloodedAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 Battery Flooded Lead-Acid data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 18. NETA ATS-2025 Direct-Current Systems, Batteries, Valve-Regulated Lead-Acid (VRLA) Analysis Endpoint (Section 7.18.1.3 & IEEE 1188)
app.post('/api/field-service/neta-battery-vrla-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaBatteryVrlaInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({
        success: false,
        error: 'Thiếu dữ liệu kiểm tra hệ thống DC & ắc quy VRLA theo NETA ATS-2025 Mục 7.18.1.3 & IEEE 1188.',
      });
    }

    const evaluation = await generateNetaBatteryVrlaAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 Battery VRLA data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 19. NETA ATS-2025 Emergency Systems, Automatic Transfer Switches (ATS) Analysis Endpoint (Section 7.22.3)
app.post('/api/field-service/neta-ats-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaAtsInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({
        success: false,
        error: 'Thiếu dữ liệu kiểm tra bộ chuyển nguồn tự động ATS theo NETA ATS-2025 Mục 7.22.3.',
      });
    }

    const evaluation = await generateNetaAtsAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 Automatic Transfer Switch data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 20. NETA ATS-2025 Emergency Systems, Engine Generator Analysis Endpoint (Section 7.22.1 & NFPA 110)
app.post('/api/field-service/neta-generator-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaGeneratorInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({
        success: false,
        error: 'Thiếu dữ liệu kiểm định tổ máy phát điện khẩn cấp theo NETA ATS-2025 Mục 7.22.1 & NFPA 110.',
      });
    }

    const evaluation = await generateNetaGeneratorAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 Engine Generator data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 19. NETA ATS-2025 Switchgear and Switchboard Assemblies AI Evaluation (Section 7.1.1)
app.post('/api/field-service/neta-switchgear-analyze', async (req, res) => {
  try {
    const payload = req.body as NetaSwitchgearInputPayload;
    if (!payload || !payload.electrical_tests || !payload.visual_inspection) {
      return res.status(400).json({
        success: false,
        error: 'Thiếu dữ liệu kiểm định tủ điện phân phối và tủ đóng cắt theo NETA ATS-2025 Mục 7.1.1.',
      });
    }

    const evaluation = await generateNetaSwitchgearAiAnalysis(payload);
    res.json({
      success: true,
      data: evaluation,
    });
  } catch (error: any) {
    console.error('Error analyzing NETA ATS-2025 Switchgear data:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// 18. AI OCR Nameplate Scanner Endpoint (Gemini Vision Multimodal)
app.post('/api/field-service/scan-nameplate', async (req, res) => {
  try {
    const { image_base64, mime_type, equipment_category } = req.body;
    if (!image_base64) {
      return res.status(400).json({
        success: false,
        error: 'Thiếu dữ liệu ảnh chụp nhãn mác (image_base64 is required).',
      });
    }

    const result = await extractNameplateDataWithAi({
      image_base64,
      mime_type: mime_type || 'image/jpeg',
      equipment_category: equipment_category || 'transformer',
    });

    res.json(result);
  } catch (error: any) {
    console.error('Error during AI Nameplate OCR scan:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Vite middleware setup
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*all', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

// Export the app for Vercel serverless functions
export default app;

// Only start the server if we're not running in Vercel's serverless environment
if (process.env.VERCEL !== '1') {
  startServer();
}
