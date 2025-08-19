const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
const XLSX = require('xlsx');
const http = require('http');              // 👈 add
const { Server } = require('socket.io');   // 👈 add

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static('public'));

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
  scopes: ['https://www.googleapis.com/auth/drive'],
});

// === Replace with your chat JSON file ID on Google Drive ===
const CHAT_FILE_ID = 'YOUR_CHAT_JSON_FILE_ID';

// === Helper: Read chat history from Drive ===
async function readChatHistory() {
  const drive = google.drive({ version: 'v3', auth: await auth.getClient() });
  try {
    const res = await drive.files.get(
      { fileId: CHAT_FILE_ID, alt: 'media' },
      { responseType: 'stream' }
    );

    let data = '';
    return new Promise((resolve, reject) => {
      res.data
        .on('data', chunk => (data += chunk))
        .on('end', () => {
          try {
            resolve(JSON.parse(data));
          } catch (err) {
            resolve([]); // Return empty array if JSON invalid
          }
        })
        .on('error', err => reject(err));
    });
  } catch (err) {
    console.error('Failed to read chat JSON:', err.message);
    return [];
  }
}

// === Helper: Write chat history to Drive ===
async function writeChatHistory(history) {
  const drive = google.drive({ version: 'v3', auth: await auth.getClient() });
  const buffer = Buffer.from(JSON.stringify(history, null, 2), 'utf-8');
  await drive.files.update({
    fileId: CHAT_FILE_ID,
    media: {
      mimeType: 'application/json',
      body: buffer,
    },
  });
}

// === List Files in Folder ===
app.get('/list-files', async (req, res) => {
  try {
    const drive = google.drive({ version: 'v3', auth: await auth.getClient() });

    const response = await drive.files.list({
      q: "'1dCYG_EF7m_ZjtQGvmWDcnNNT-OowL9V5' in parents and trashed=false and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'",
      fields: 'files(id, name)',
    });

    res.json(response.data.files);
  } catch (error) {
    console.error('Failed to list files:', error.message);
    res.status(500).json({ error: 'Failed to list files.' });
  }
});

// === Load XLSX as FULL MATRIX ===
app.get('/xlsx-to-json/:fileId', async (req, res) => {
  try {
    const drive = google.drive({ version: 'v3', auth: await auth.getClient() });

    const response = await drive.files.get(
      { fileId: req.params.fileId, alt: 'media' },
      { responseType: 'arraybuffer' }
    );

    const workbook = XLSX.read(response.data, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // 🧠 Return as 2D array to preserve blank rows/columns
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: true });

    res.json(matrix);
  } catch (error) {
    console.error('Failed to load Excel file:', error.message);
    res.status(500).json({ error: 'Failed to load file.' });
  }
});

// === Upload Edited JSON as XLSX ===
app.post('/upload-xlsx/:fileId', async (req, res) => {
  try {
    const matrix = req.body.data;

    // ✅ Filter out rows where columns 0 to 5 are all empty
    const filteredData = matrix.filter(row => {
      for (let col = 0; col <= 5; col++) {
        const cell = row[col];
        if (cell !== undefined && cell !== null && String(cell).trim() !== '') {
          return true; // Keep this row
        }
      }
      return false; // Remove this row
    });

    const worksheet = XLSX.utils.aoa_to_sheet(filteredData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    const drive = google.drive({ version: 'v3', auth: await auth.getClient() });

    await drive.files.update({
      fileId: req.params.fileId,
      media: {
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        body: Buffer.from(buffer)
      }
    });

    res.send('File updated successfully');
  } catch (error) {
    console.error('Failed to upload Excel:', error.message);
    res.status(500).json({ error: 'Failed to upload file.' });
  }
});

app.get('/', (req, res) => {
  res.sendFile(__dirname + '/public/index.html');
});

// === SOCKET.IO CHAT SETUP ===
const server = http.createServer(app);
const io = new Server(server, {
  cors: { origin: "*" }
});

io.on('connection', async (socket) => {
  console.log('🟢 A user connected:', socket.id);

  // Send existing chat history to newly connected user
  const history = await readChatHistory();
  history.forEach(msg => socket.emit('chat message', msg));

  socket.on('chat message', async (msg) => {
    console.log('💬 Message:', msg);

    // Save to Google Drive JSON
    const chatHistory = await readChatHistory();
    chatHistory.push(msg);

    // Optional: limit chat history size
    if (chatHistory.length > 1000) chatHistory.shift();

    await writeChatHistory(chatHistory);

    io.emit('chat message', msg); // broadcast to all clients
  });

  socket.on('disconnect', () => {
    console.log('🔴 A user disconnected:', socket.id);
  });
});

// === Serve borak.html ===
app.get('/borak', (req, res) => {
  res.sendFile(__dirname + '/public/borak.html');
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
  console.log(`🚀 Server running on http://localhost:${PORT}`);
});
