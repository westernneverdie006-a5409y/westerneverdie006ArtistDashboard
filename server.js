const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
const XLSX = require('xlsx');
const http = require('http');
const { Pool } = require('pg');
const { Server } = require('socket.io');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static('public'));

// === Google Drive Auth ===
const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
  scopes: ['https://www.googleapis.com/auth/drive'],
});

// === PostgreSQL Setup (for chat messages only) ===
const pool = new Pool({
  connectionString: process.env.DATABASE_URL || undefined,
  host: process.env.PGHOST,
  port: process.env.PGPORT,
  user: process.env.PGUSER,
  password: process.env.PGPASSWORD,
  database: process.env.PGDATABASE,
  ssl: {
    rejectUnauthorized: false
  }
});

// Ensure messages table exists
async function ensureMessagesTable() {
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS messages (
        id SERIAL PRIMARY KEY,
        text TEXT NOT NULL,
        sender VARCHAR(50) NOT NULL,
        time TIMESTAMP NOT NULL
      )
    `);
    console.log("✅ messages table is ready!");
  } catch (err) {
    console.error("❌ Failed to create messages table:", err);
  }
}

pool.connect()
  .then(() => {
    console.log("✅ Connected to PostgreSQL!");
    ensureMessagesTable();
  })
  .catch(err => console.error("❌ PostgreSQL connection error:", err));

// === Google Drive XLSX Routes ===
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
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: true });
    res.json(matrix);
  } catch (error) {
    console.error('Failed to load Excel file:', error.message);
    res.status(500).json({ error: 'Failed to load file.' });
  }
});

app.post('/upload-xlsx/:fileId', async (req, res) => {
  try {
    const matrix = req.body.data;
    const filteredData = matrix.filter(row => {
      for (let col = 0; col <= 5; col++) {
        const cell = row[col];
        if (cell !== undefined && cell !== null && String(cell).trim() !== '') {
          return true;
        }
      }
      return false;
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

// === Serve HTML ===
app.get('/', (req, res) => res.sendFile(__dirname + '/public/index.html'));
app.get('/borak', (req, res) => res.sendFile(__dirname + '/public/borak.html'));

// === SOCKET.IO CHAT SETUP ===
const server = http.createServer(app);
const io = new Server(server, { cors: { origin: "*" } });

io.on('connection', async (socket) => {
  console.log('🟢 A user connected:', socket.id);

  // Send last 50 messages from PostgreSQL
  try {
    const result = await pool.query(
      'SELECT text, sender, time FROM messages ORDER BY time ASC LIMIT 50'
    );
    result.rows.forEach(msg => socket.emit('chat message', msg));
  } catch (err) {
    console.error('Failed to fetch messages:', err);
  }

  socket.on('chat message', async (msg) => {
    io.emit('chat message', msg);

    // Save message to PostgreSQL
    try {
      await pool.query(
        'INSERT INTO messages (text, sender, time) VALUES ($1, $2, $3)',
        [msg.text, msg.sender, msg.time]
      );
    } catch (err) {
      console.error('Failed to save message:', err);
    }
  });

  socket.on('disconnect', () => {
    console.log('🔴 A user disconnected:', socket.id);
  });
});

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
  console.log(`🚀 Server running on http://localhost:${PORT}`);
});
