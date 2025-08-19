const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
const XLSX = require('xlsx');
const http = require('http');
const { Server } = require('socket.io');
const { MongoClient, ServerApiVersion } = require('mongodb'); // Added MongoDB

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static('public'));

// === Google Drive Auth ===
const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
  scopes: ['https://www.googleapis.com/auth/drive'],
});

// === MongoDB Setup ===
const mongoUri = process.env.MONGO_URI || "mongodb+srv://BorakApp:Cassey5409@westernneverdie006datab.saly3qq.mongodb.net/?retryWrites=true&w=majority";
const mongoClient = new MongoClient(mongoUri, {
  serverApi: { version: ServerApiVersion.v1, strict: true, deprecationErrors: true },
});

let chatCollection;

async function connectMongo() {
  if (chatCollection) return; // already connected
  try {
    console.log("⏳ Connecting to MongoDB...");
    await mongoClient.connect();
    const db = mongoClient.db("BorakChatDB");
    chatCollection = db.collection("messages");
    await chatCollection.createIndex({ time: 1 }); // Optional for faster queries
    console.log("✅ Connected to MongoDB!");
  } catch (err) {
    console.error("❌ MongoDB connection error:", err.message);
    process.exit(1); // Stop server if DB connection fails
  }
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

app.get('/', (req, res) => {
  res.sendFile(__dirname + '/public/index.html');
});

// === SOCKET.IO CHAT SETUP ===
const server = http.createServer(app);
const io = new Server(server, { cors: { origin: "*" } });

io.on('connection', async (socket) => {
  console.log('🟢 A user connected:', socket.id);

  // Ensure MongoDB is connected
  if (!chatCollection) await connectMongo();

  // Send last 50 messages from MongoDB to client
  if (chatCollection) {
    try {
      const recentMessages = await chatCollection.find().sort({ time: 1 }).limit(50).toArray();
      recentMessages.forEach(msg => socket.emit('chat message', msg));
    } catch (err) {
      console.error('Failed to fetch recent messages:', err);
    }
  }

  socket.on('chat message', async (msg) => {
    console.log('💬 Message:', msg);
    io.emit('chat message', msg); // broadcast to all clients

    // Save message to MongoDB
    if (chatCollection) {
      try {
        await chatCollection.insertOne({
          text: msg.text,
          sender: msg.sender,
          time: new Date(msg.time)
        });
      } catch (err) {
        console.error('Failed to save message:', err);
      }
    }
  });

  socket.on('disconnect', () => {
    console.log('🔴 A user disconnected:', socket.id);
  });
});

// === Serve borak.html ===
app.get('/borak', (req, res) => {
  res.sendFile(__dirname + '/public/borak.html');
});

// === Start Server after MongoDB is connected ===
(async () => {
  await connectMongo();
  const PORT = process.env.PORT || 3000;
  server.listen(PORT, () => {
    console.log(`🚀 Server running on http://localhost:${PORT}`);
  });
})();
