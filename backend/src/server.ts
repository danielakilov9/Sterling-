import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import https from 'https';
import http from 'http';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { chatRouter } from './routes/chat';

dotenv.config();

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors({ origin: '*', methods: ['GET', 'POST'], allowedHeaders: ['Content-Type'] }));
app.use(express.json({ limit: '10mb' }));
app.use('/api/chat', chatRouter);
app.get('/health', (_req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// ── Local dev: run HTTPS so the add-in (https://localhost:3000) can call us ──
// The office-addin-dev-certs package installs certs in ~/.office-addin-dev-certs/
// We try HTTPS first, fall back to HTTP if certs aren't found (e.g. on Render).

const certDir = path.join(os.homedir(), '.office-addin-dev-certs');
const certPath = path.join(certDir, 'localhost.crt');
const keyPath  = path.join(certDir, 'localhost.key');

if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
  const httpsOptions = {
    cert: fs.readFileSync(certPath),
    key:  fs.readFileSync(keyPath),
  };
  https.createServer(httpsOptions, app).listen(PORT, () => {
    console.log(`✅ Backend running on https://localhost:${PORT} (HTTPS)`);
    console.log(`   Health check: https://localhost:${PORT}/health`);
  });
} else {
  // Fallback for Render deployment or if certs not installed yet
  http.createServer(app).listen(PORT, () => {
    console.log(`✅ Backend running on http://localhost:${PORT} (HTTP)`);
    console.log(`   Run: npx office-addin-dev-certs install   to enable HTTPS`);
  });
}