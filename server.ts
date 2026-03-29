import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import { google } from 'googleapis';
import path from 'path';
import multer from 'multer';
import { Readable } from 'stream';

dotenv.config();

export const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json());

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
  const authHeader = req.headers.authorization;
  if (authHeader && authHeader.startsWith('Bearer ')) {
    return res.json({ isAuthenticated: true });
  }
  res.json({ isAuthenticated: !!globalTokens });
});

// Helper to get tokens from request
const getTokensFromRequest = (req: express.Request) => {
  const authHeader = req.headers.authorization;
  if (authHeader && authHeader.startsWith('Bearer ')) {
    try {
      return JSON.parse(authHeader.substring(7));
    } catch (e) {
      console.error('Failed to parse tokens from auth header:', e);
      return null;
    }
  }
  return globalTokens;
};

// 4. Save Data to Google Sheets
app.post('/api/sheets/append', async (req, res) => {
  const tokens = getTokensFromRequest(req);
  if (!tokens) {
    return res.status(401).json({ error: 'Not authenticated with Google' });
  }

  const { range, values } = req.body;
  const spreadsheetId = process.env.SPREADSHEET_ID;

  if (!spreadsheetId) {
    return res.status(400).json({ error: 'SPREADSHEET_ID environment variable is not set' });
  }

  try {
    const oauth2Client = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET
    );
    oauth2Client.setCredentials(tokens);

    const sheets = google.sheets({ version: 'v4', auth: oauth2Client });

    // Ensure sheets exist
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const existingSheetTitles = spreadsheet.data.sheets?.map(s => s.properties?.title) || [];
    
    const requiredSheets = ['Máy biến áp', 'Tủ điện trung thế', 'Động cơ'];
    const missingSheets = requiredSheets.filter(title => !existingSheetTitles.includes(title));
    
    if (missingSheets.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: missingSheets.map(title => ({
            addSheet: {
              properties: { title }
            }
          }))
        }
      });
    }

    const response = await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: range || 'Máy biến áp!A:Z',
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [values]
      }
    });

    res.json({ success: true, data: response.data });
  } catch (error: any) {
    console.error('Error appending to sheet:', error);
    if (error.code === 401 || error.status === 401) {
      return res.status(401).json({ error: 'Phiên đăng nhập đã hết hạn. Vui lòng tải lại trang và kết nối lại Google Drive.' });
    }
    res.status(500).json({ error: error.message || 'Failed to save to Google Sheets' });
  }
});

// 5. Get Data from Google Sheets
app.get('/api/sheets/get', async (req, res) => {
  const tokens = getTokensFromRequest(req);
  if (!tokens) {
    return res.status(401).json({ error: 'Not authenticated with Google' });
  }

  const spreadsheetId = process.env.SPREADSHEET_ID;
  if (!spreadsheetId) {
    return res.status(400).json({ error: 'SPREADSHEET_ID environment variable is not set' });
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
    const sheetNames = spreadsheet.data.sheets?.map(s => s.properties?.title) || [];
    
    // Find sheets by keywords to be more robust
    const transformerSheet = sheetNames.find(n => n?.toLowerCase().includes('biến áp')) || sheetNames[0] || 'Máy biến áp';
    const switchgearSheet = sheetNames.find(n => n?.toLowerCase().includes('tủ điện')) || sheetNames[1] || 'Tủ điện trung thế';
    const motorSheet = sheetNames.find(n => n?.toLowerCase().includes('động cơ')) || sheetNames[2] || 'Động cơ';

    const ranges = [
      `${transformerSheet}!A:Z`,
      `${switchgearSheet}!A:Z`,
      `${motorSheet}!A:Z`
    ];

    let response;
    try {
      response = await sheets.spreadsheets.values.batchGet({
        spreadsheetId,
        ranges,
      });
    } catch (e: any) {
      console.log('Error fetching sheets (might not exist yet):', e.message);
      if (e.code === 401 || e.status === 401) {
        throw e; // Rethrow auth errors to be handled by the outer catch block
      }
      return res.json({
        success: true,
        data: { transformers: [], switchgears: [], motors: [] }
      });
    }

    const data = {
      transformers: response.data.valueRanges?.[0]?.values || [],
      switchgears: response.data.valueRanges?.[1]?.values || [],
      motors: response.data.valueRanges?.[2]?.values || [],
    };

    res.json({ success: true, data });
  } catch (error: any) {
    console.error('Error getting data from sheet:', error);
    if (error.code === 401 || error.status === 401) {
      return res.status(401).json({ error: 'Phiên đăng nhập đã hết hạn. Vui lòng tải lại trang và kết nối lại Google Drive.' });
    }
    res.status(500).json({ error: error.message || 'Failed to get data from Google Sheets' });
  }
});

// 6. Sync/Export Mock Data to Google Sheets
app.post('/api/sheets/sync-export', async (req, res) => {
  const tokens = getTokensFromRequest(req);
  if (!tokens) {
    return res.status(401).json({ error: 'Not authenticated with Google' });
  }

  const { transformers, switchgears, motors } = req.body;
  const spreadsheetId = process.env.SPREADSHEET_ID;

  if (!spreadsheetId) {
    return res.status(400).json({ error: 'SPREADSHEET_ID environment variable is not set' });
  }

  try {
    const oauth2Client = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET
    );
    oauth2Client.setCredentials(tokens);

    const sheets = google.sheets({ version: 'v4', auth: oauth2Client });

    // Ensure sheets exist
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const existingSheetTitles = spreadsheet.data.sheets?.map(s => s.properties?.title) || [];
    
    const requiredSheets = ['Máy biến áp', 'Tủ điện trung thế', 'Động cơ'];
    const missingSheets = requiredSheets.filter(title => !existingSheetTitles.includes(title));
    
    if (missingSheets.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: missingSheets.map(title => ({
            addSheet: {
              properties: { title }
            }
          }))
        }
      });
    }

    // Clear existing data first (optional, but good for a full sync)
    try {
      await sheets.spreadsheets.values.batchClear({
        spreadsheetId,
        requestBody: {
          ranges: ['Máy biến áp!A1:Z', 'Tủ điện trung thế!A1:Z', 'Động cơ!A1:Z']
        }
      });
    } catch (e: any) {
      console.log('Clear error (might be missing sheets or using tables):', e.message);
      if (e.message && e.message.toLowerCase().includes('table')) {
        throw new Error('File Google Sheets đang sử dụng tính năng "Table" (Bảng) nên không thể ghi đè. Vui lòng vào Google Sheets, click chuột phải vào bảng -> chọn "Convert to range" (Chuyển đổi thành dải ô) rồi thử lại.');
      }
    }

    // Update with new data
    const data = [];
    if (transformers && transformers.length > 0) {
      data.push({ range: 'Máy biến áp!A1', values: transformers });
    }
    if (switchgears && switchgears.length > 0) {
      data.push({ range: 'Tủ điện trung thế!A1', values: switchgears });
    }
    if (motors && motors.length > 0) {
      data.push({ range: 'Động cơ!A1', values: motors });
    }

    if (data.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        requestBody: {
          valueInputOption: 'USER_ENTERED',
          data: data
        }
      });
    }

    res.json({ success: true, message: 'Data synced successfully' });
  } catch (error: any) {
    console.error('Error syncing to sheet:', error);
    if (error.code === 401 || error.status === 401) {
      return res.status(401).json({ error: 'Phiên đăng nhập đã hết hạn. Vui lòng tải lại trang và kết nối lại Google Drive.' });
    }
    res.status(500).json({ error: error.message || 'Failed to sync to Google Sheets' });
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
    if (error.code === 401 || error.status === 401) {
      return res.status(401).json({ error: 'Phiên đăng nhập đã hết hạn. Vui lòng tải lại trang và kết nối lại Google Drive.' });
    }
    res.status(500).json({ error: error.message || 'Failed to upload files to Google Drive' });
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
