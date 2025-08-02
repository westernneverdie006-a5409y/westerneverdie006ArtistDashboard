const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');

const app = express();

// === Middleware ===
app.use(cors());
app.use(express.json());
app.use(express.static('public')); // Serves index.html and other assets

// === Google Drive Auth ===
const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
  scopes: ['https://www.googleapis.com/auth/drive.metadata.readonly'],
});

// === API: List Files in a Specific Folder ===
app.get('/list-files', async (req, res) => {
  try {
    const drive = google.drive({ version: 'v3', auth: await auth.getClient() });

    const response = await drive.files.list({
      q: "'1dCYG_EF7m_ZjtQGvmWDcnNNT-OowL9V5' in parents and trashed=false",
      fields: 'files(id, name, mimeType, modifiedTime)',
    });

    res.json(response.data.files);
  } catch (error) {
    console.error('Failed to list files:', error.message);
    res.status(500).json({ error: 'Failed to fetch files from Google Drive.' });
  }
});

// === Fallback Home Route ===
app.get('/', (req, res) => {
  res.sendFile(__dirname + '/public/index.html');
});

// === Start Server ===
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Server running on http://localhost:${PORT}`);
});
