const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
const XLSX = require('xlsx');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static('public'));

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
  scopes: ['https://www.googleapis.com/auth/drive'],
});

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

    if (!Array.isArray(matrix) || matrix.length === 0) {
      throw new Error("Invalid matrix data");
    }

    // Keep header row (row 0)
    const header = matrix[0];

    // Clean rows: keep only rows that have some content (excluding header)
    const filteredBody = matrix.slice(1).filter(
      row => row.some(cell => cell !== undefined && cell !== null && String(cell).trim() !== '')
    );

    const cleanedMatrix = [header, ...filteredBody];

    // Trim trailing empty rows
    while (
      cleanedMatrix.length > 1 &&
      cleanedMatrix[cleanedMatrix.length - 1].every(cell => cell === '' || cell === null || cell === undefined)
    ) {
      cleanedMatrix.pop();
    }

    // Generate worksheet
    const worksheet = XLSX.utils.aoa_to_sheet(cleanedMatrix);

    // Define sheet range manually to prevent 1000+ rows
    const numRows = cleanedMatrix.length;
    const numCols = Math.max(...cleanedMatrix.map(row => row.length));
    const endCol = XLSX.utils.encode_col(numCols - 1);
    const endRow = numRows;
    worksheet['!ref'] = `A1:${endCol}${endRow}`;

    // Create workbook
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    // Upload to Google Drive
    const drive = google.drive({ version: 'v3', auth: await auth.getClient() });

    await drive.files.update({
      fileId: req.params.fileId,
      media: {
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        body: Buffer.from(buffer),
      },
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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Server running on http://localhost:${PORT}`);
});
