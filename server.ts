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

  let spreadsheetId = process.env.SPREADSHEET_ID;
  if (!spreadsheetId) {
    return res.status(400).json({ error: 'SPREADSHEET_ID environment variable is not set' });
  }

  // Extract ID if full URL is provided
  if (spreadsheetId.includes('docs.google.com/spreadsheets/d/')) {
    const match = spreadsheetId.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match) spreadsheetId = match[1];
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
      return res.json({ success: true, data: { transformers: [], switchgears: [], motors: [] } });
    }

    const response = await sheets.spreadsheets.values.batchGet({
      spreadsheetId,
      ranges: sheetNames.map(name => `'${name}'!A:AZ`),
    });

    const allData: Record<string, any[][]> = {};
    sheetNames.forEach((name, index) => {
      allData[name] = response.data.valueRanges?.[index]?.values || [];
    });

    // Categorize data for backward compatibility or easier frontend processing
    const categorizedData = {
      transformers: [] as any[][],
      switchgears: [] as any[][],
      motors: [] as any[][],
      allSheets: allData
    };

    sheetNames.forEach(name => {
      const lowerName = name.toLowerCase();
      const rows = allData[name];
      if (lowerName.includes('biến áp') || lowerName.includes('mba') || lowerName.includes('transformer')) {
        categorizedData.transformers = categorizedData.transformers.concat(rows);
      } else if (lowerName.includes('tủ điện') || lowerName.includes('trung thế') || lowerName.includes('hạ thế') || lowerName.includes('switchgear') || lowerName.includes('swg') || lowerName.includes('mcc') || lowerName.includes('tev') || lowerName.includes('rmu')) {
        categorizedData.switchgears = categorizedData.switchgears.concat(rows);
      } else if (lowerName.includes('động cơ') || lowerName.includes('motor') || lowerName.includes('máy bơm') || lowerName.includes('pump') || lowerName.includes('fan') || lowerName.includes('quạt')) {
        categorizedData.motors = categorizedData.motors.concat(rows);
      }
    });

    // Fallback: If a category is empty, try to find sheets that might contain that data but weren't caught by keywords
    // Or if all categories are empty, put all sheets into all categories to let the frontend filter by column headers
    if (categorizedData.transformers.length === 0 && categorizedData.switchgears.length === 0 && categorizedData.motors.length === 0) {
      sheetNames.forEach(name => {
        categorizedData.transformers = categorizedData.transformers.concat(allData[name]);
        categorizedData.switchgears = categorizedData.switchgears.concat(allData[name]);
        categorizedData.motors = categorizedData.motors.concat(allData[name]);
      });
    } else {
      // If some categories are still empty, use the first few sheets as fallback
      if (categorizedData.transformers.length === 0 && sheetNames.length > 0) categorizedData.transformers = allData[sheetNames[0]];
      if (categorizedData.switchgears.length === 0 && sheetNames.length > 1) categorizedData.switchgears = allData[sheetNames[1]];
      if (categorizedData.motors.length === 0 && sheetNames.length > 2) categorizedData.motors = allData[sheetNames[2]];
    }

    res.json({ success: true, data: categorizedData });
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

    // Ensure sheets exist or match by keyword
    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const existingSheetTitles = spreadsheet.data.sheets?.map(s => s.properties?.title) || [];
    
    const requiredSheets = ['Máy biến áp', 'Tủ điện trung thế', 'Động cơ'];
    const keywords = [
      ['biến áp'],
      ['tủ điện', 'tev'],
      ['động cơ']
    ];

    const finalSheetNames = requiredSheets.map((title, i) => {
      const match = existingSheetTitles.find(n => 
        keywords[i].some(k => n.toLowerCase().includes(k))
      );
      return match || title;
    });

    const missingSheets = finalSheetNames.filter(title => !existingSheetTitles.includes(title));
    
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

    // Clear existing data first
    try {
      await sheets.spreadsheets.values.batchClear({
        spreadsheetId,
        requestBody: {
          ranges: finalSheetNames.map(name => `${name}!A1:Z`)
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
      data.push({ range: `${finalSheetNames[0]}!A1`, values: transformers });
    }
    if (switchgears && switchgears.length > 0) {
      data.push({ range: `${finalSheetNames[1]}!A1`, values: switchgears });
    }
    if (motors && motors.length > 0) {
      data.push({ range: `${finalSheetNames[2]}!A1`, values: motors });
    }

    if (data.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId,
        requestBody: {
          valueInputOption: 'USER_ENTERED',
          data: data
        }
      });

      // Format Column A as Date (dd/mm/yyyy) and auto-resize columns
      const sheetMap = new Map(spreadsheet.data.sheets?.map(s => [s.properties?.title?.trim(), s.properties?.sheetId]) || []);
      const formattingRequests = [];
      
      for (const title of finalSheetNames) {
        // Try exact match first
        let id = sheetMap.get(title);
        if (id === undefined) {
          // Try case-insensitive partial match
          const entry = Array.from(sheetMap.entries()).find(([name]) => 
            name.toLowerCase().includes(title.toLowerCase()) || 
            title.toLowerCase().includes(name.toLowerCase())
          );
          if (entry) id = entry[1];
        }

        if (id !== undefined) {
          // Format Column A (Date)
          formattingRequests.push({
            repeatCell: {
              range: {
                sheetId: id,
                startRowIndex: 1,
                startColumnIndex: 0,
                endColumnIndex: 1
              },
              cell: {
                userEnteredFormat: {
                  numberFormat: {
                    type: 'DATE',
                    pattern: 'dd/mm/yyyy'
                  }
                }
              },
              fields: 'userEnteredFormat.numberFormat'
            }
          });
          
          // Auto-resize columns
          formattingRequests.push({
            autoResizeDimensions: {
              dimensions: {
                sheetId: id,
                dimension: 'COLUMNS',
                startIndex: 0,
                endIndex: 21
              }
            }
          });
        }
      }

      if (formattingRequests.length > 0) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: {
            requests: formattingRequests
          }
        });
      }
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
