import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import multer from 'multer';
import { chatRouter } from './routes/chat';

dotenv.config();

const app = express();
const PORT = process.env.PORT || 3001;

// ─── Middleware ────────────────────────────────────────────────────────────────
app.use(cors({
  origin: process.env.FRONTEND_URL || '*',
  methods: ['GET', 'POST'],
  allowedHeaders: ['Content-Type']
}));
app.use(express.json({ limit: '10mb' })); // 10mb for large spreadsheet contexts

// ─── Routes ───────────────────────────────────────────────────────────────────
app.use('/api/chat', chatRouter);

// Health check endpoint — useful for Render deployment
app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// ─── Start ────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`✅ Financial AI Agent backend running on port ${PORT}`);
});