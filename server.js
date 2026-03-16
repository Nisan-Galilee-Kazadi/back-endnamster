import express from 'express';
import multer from 'multer';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import sharp from 'sharp';
import xlsx from 'xlsx';
import mammoth from 'mammoth';
import archiver from 'archiver';
import crypto from 'crypto';
import { PDFDocument } from 'pdf-lib';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

import cors from 'cors';

const app = express();
const port = process.env.PORT || 3001;

// Enable CORS and COOP for Google Auth
app.use(cors({
  origin: true,
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With', 'Accept'],
  credentials: true
}));

app.use((req, res, next) => {
  res.setHeader('Cross-Origin-Opener-Policy', 'same-origin-allow-popups');
  res.setHeader('Cross-Origin-Embedder-Policy', 'require-corp');
  next();
});

// Health check for Render
app.get('/', (req, res) => {
  res.json({ status: 'ok', service: 'Namster API' });
});

app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

const publicDir = path.join(__dirname, 'public');
const uploadsDir = path.join(__dirname, 'uploads');
const workDir = path.join(__dirname, 'work');

for (const d of [publicDir, uploadsDir, workDir]) {
  if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
}

app.use(express.static(publicDir));

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, uploadsDir);
  },
  filename: function (req, file, cb) {
    const unique = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, unique + '-' + file.originalname.replace(/\s+/g, '_'));
  },
});
const upload = multer({
  storage,
  limits: { fileSize: 50 * 1024 * 1024 }, // 50MB
});

const sessions = new Map();

function newSession() {
  const id = crypto.randomUUID();
  sessions.set(id, {
    id,
    createdAt: Date.now(),
    modelPath: null,
    listPath: null,
    names: [],
    cleanup: [],
  });
  return id;
}

function cleanupSession(id) {
  const s = sessions.get(id);
  if (!s) return;
  for (const p of s.cleanup) {
    try {
      if (p && fs.existsSync(p)) fs.rmSync(p, { recursive: true, force: true });
    } catch { }
  }
  sessions.delete(id);
}

function parseNamesFromText(text) {
  const lines = text.split(/\r?\n/).filter(line => line.trim() !== '');
  return lines.map(line => {
    const parts = line.split('=');
    if (parts.length >= 2) {
      return {
        name: parts[0].trim(),
        table: parts.slice(1).join('=').trim()
      };
    }
    const altParts = line.split(/[:\t]/);
    if (altParts.length >= 2) {
      return {
        name: altParts[0].trim(),
        table: altParts[1].trim()
      };
    }
    return { name: line.trim(), table: '' };
  }).filter(item => {
    const norm = item.name.toLowerCase().replace(/\s+/g, '');
    return norm !== 'liste' && norm !== '';
  });
}

async function extractNamesFromFile(filePath) {
  const ext = path.extname(filePath).toLowerCase();

  if (ext === '.csv') {
    return parseNamesFromText(fs.readFileSync(filePath, 'utf8'));
  }

  if (ext === '.xlsx' || ext === '.xls') {
    const wb = xlsx.readFile(filePath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(ws, { header: 1 });
    return data.map(row => {
      let name = String(row[0] || '').trim();
      let table = String(row[1] || '').trim();
      if (!table && name.includes('=')) {
        const parts = name.split('=');
        name = parts[0].trim();
        table = parts[1].trim();
      } else if (!table && name.includes(':')) {
        const parts = name.split(':');
        name = parts[0].trim();
        table = parts[1].trim();
      }
      return { name, table };
    }).filter(row => row.name);
  }

  if (ext === '.docx') {
    const { value } = await mammoth.extractRawText({ path: filePath });
    return parseNamesFromText(value || '');
  }

  if (ext === '.pdf') {
    try {
      const dataBuffer = fs.readFileSync(filePath);
      const { default: pdfParse } = await import('pdf-parse');
      const data = await pdfParse(dataBuffer);
      return parseNamesFromText(data.text || '');
    } catch (e) {
      console.warn('PDF parsing failed:', e?.message);
      return [];
    }
  }

  try {
    return parseNamesFromText(fs.readFileSync(filePath, 'utf8'));
  } catch {
    return [];
  }
}

function buildSVGOverlay(elements, width, height) {
  const allowedFonts = new Set([
    'Alex Brush',
    'Great Vibes',
    'Dancing Script',
    'Playfair Display',
    'Bodoni Moda',
    'Cinzel',
    'Cormorant Garamond',
    'Pinyon Script',
    'Rochester',
    'Sacramento',
    'Brush Script MT',
    'Monotype Corsiva',
    'Lucida Calligraphy',
    'Segoe Script',
    'Gabriola',
    'Palace Script MT',
    'Edwardian Script ITC',
    'Kunstler Script',
    'Vladimir Script',
    'Vivaldi',
    'Garamond',
    'Book Antiqua'
  ]);
  const safeStr = (s) => (s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const textItems = elements.map(el => {
    const requestedFont = (el.fontFamily || 'Arial').trim();
    const ff = allowedFonts.has(requestedFont) ? requestedFont : 'Arial';
    const fsPx = Number(el.fontSize) || 48;
    const fill = el.color || '#000000';
    const fw = el.fontWeight || 'normal';
    const fst = el.fontStyle || 'normal';
    const td = el.textDecoration || 'none';
    return `<text x="${el.x}" y="${el.y}" style="font-family: '${ff}'; font-size: ${fsPx}px; font-weight: ${fw}; font-style: ${fst}; text-decoration: ${td}; fill: ${fill}; dominant-baseline: hanging;">${safeStr(el.text)}</text>`;
  }).join('\n');
  return `<?xml version="1.0" encoding="UTF-8"?>
<svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" xmlns="http://www.w3.org/2000/svg">
  ${textItems}
</svg>`;
}

async function composeImageWithElements(modelPath, outPath, elements) {
  const img = sharp(fs.readFileSync(modelPath));
  const meta = await img.metadata();
  const width = meta.width || 2000;
  const height = meta.height || 1000;
  const svg = buildSVGOverlay(elements, width, height);
  const buffer = await img.composite([{ input: Buffer.from(svg), top: 0, left: 0 }]).png().toBuffer();
  await fs.promises.writeFile(outPath, buffer);
}

app.post('/api/upload', upload.fields([
  { name: 'model', maxCount: 1 },
  { name: 'list', maxCount: 1 },
]), async (req, res) => {
  try {
    const sid = newSession();
    const s = sessions.get(sid);
    const model = req.files['model']?.[0];
    const list = req.files['list']?.[0];
    if (!model || !list) return res.status(400).json({ error: 'Model image and list file are required.' });
    s.modelPath = model.path;
    s.listPath = list.path;
    s.cleanup.push(model.path, list.path);
    const names = await extractNamesFromFile(s.listPath);
    s.names = names;
    res.json({ sessionId: sid, namesTotal: names.length });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Upload failed.' });
  }
});

app.post('/api/test', express.json(), async (req, res) => {
  try {
    const { sessionId, x, y, tx, ty, useTable, fontFamily, fontSize, color, fontWeight, fontStyle, textDecoration } = req.body || {};
    const s = sessions.get(sessionId);
    if (!s) return res.status(400).json({ error: 'Invalid session' });
    const firstEntry = s.names[0] || { name: 'INVITE TEST', table: '01' };
    const elements = [{ text: firstEntry.name, x: Number(x) || 100, y: Number(y) || 100, fontFamily, fontSize, color, fontWeight, fontStyle, textDecoration }];
    if (useTable && tx !== undefined && ty !== undefined) {
      elements.push({ text: firstEntry.table || '01', x: Number(tx) || 100, y: Number(ty) || 100, fontFamily, fontSize, color, fontWeight, fontStyle, textDecoration });
    }
    const outDir = path.join(workDir, sessionId);
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
    const outPath = path.join(outDir, 'test.png');
    await composeImageWithElements(s.modelPath, outPath, elements);
    s.cleanup.push(outDir);
    const data = fs.readFileSync(outPath);
    res.json({ preview: 'data:image/png;base64,' + data.toString('base64') });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Test render failed.' });
  }
});

app.post('/api/generate', express.json(), async (req, res) => {
  try {
    const { sessionId, x, y, tx, ty, useTable, fontFamily, fontSize, color, fontWeight, fontStyle, textDecoration, offset, limit } = req.body || {};
    const s = sessions.get(sessionId);
    if (!s) return res.status(400).json({ error: 'Invalid session' });
    const outDir = path.join(workDir, sessionId, 'all');
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
    const start = Math.max(0, Number(offset) || 0);
    const batchSize = Math.min(50, Number(limit) || 50);
    const endExclusive = Math.min(s.names.length, start + batchSize);
    for (let idx = start; idx < endExclusive; idx++) {
      const entry = s.names[idx];
      const filename = `${String(idx + 1).padStart(3, '0')}-${entry.name.replace(/[^a-z0-9_-]+/gi, '_')}.png`;
      const outPath = path.join(outDir, filename);
      const elements = [{ text: entry.name, x: Number(x) || 100, y: Number(y) || 100, fontFamily, fontSize, color, fontWeight, fontStyle, textDecoration }];
      if (useTable && tx !== undefined && ty !== undefined) {
        elements.push({ text: entry.table, x: Number(tx) || 100, y: Number(ty) || 100, fontFamily, fontSize, color, fontWeight, fontStyle, textDecoration });
      }
      await composeImageWithElements(s.modelPath, outPath, elements);
    }
    const zipPath = path.join(workDir, sessionId, 'invitations.zip');
    await new Promise((resolve, reject) => {
      const output = fs.createWriteStream(zipPath);
      const archive = archiver('zip', { zlib: { level: 9 } });
      output.on('close', resolve);
      archive.on('error', reject);
      archive.pipe(output);
      archive.directory(outDir, false);
      archive.finalize();
    });
    s.cleanup.push(path.join(workDir, sessionId));
    res.json({ downloadUrl: `/api/download/${sessionId}`, processed: endExclusive - start, total: s.names.length });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Generation failed.' });
  }
});

app.get('/api/download/:sid', (req, res) => {
  const sid = req.params.sid;
  const s = sessions.get(sid);
  if (!s) return res.status(400).json({ error: 'Invalid session' });
  const zipPath = path.join(workDir, sid, 'invitations.zip');
  if (!fs.existsSync(zipPath)) return res.status(404).json({ error: 'ZIP not found' });
  res.download(zipPath, 'invitations.zip', (err) => {
    cleanupSession(sid);
  });
});

// --- Auth + User Routes (MongoDB Implementation) ---
import mongoose from 'mongoose';
import dotenv from 'dotenv';
dotenv.config();

const MONGO_URI = process.env.MONGO_URI || "mongodb://localhost:27017/namster";
mongoose.connect(MONGO_URI).then(() => {
  console.log("Connected to MongoDB via Mongoose");
}).catch(err => console.error("MongoDB connection error:", err));

const userSchema = new mongoose.Schema({
  id: { type: String, required: true, unique: true },
  email: { type: String, required: true, unique: true },
  password: { type: String },
  firstName: { type: String, default: '' },
  lastName: { type: String, default: '' },
  role: { type: String, default: 'user' },
  isPremium: { type: Boolean, default: false },
  isBanned: { type: Boolean, default: false },
  avatar: { type: String, default: '' },
  favoriteTemplateIds: { type: [Number], default: [] }
}, { timestamps: true });

const User = mongoose.model('User', userSchema);

const savedTemplateSchema = new mongoose.Schema({
  id: { type: String, required: true, unique: true },
  userId: { type: String, required: true },
  name: { type: String, required: true },
  templateId: { type: Number, required: true },
  customizationData: { type: mongoose.Schema.Types.Mixed, default: {} }
}, { timestamps: true });

const SavedTemplate = mongoose.model('SavedTemplate', savedTemplateSchema);

const historySchema = new mongoose.Schema({
  id: { type: String, required: true, unique: true },
  userId: { type: String, required: true },
  action: { type: String, required: true },
  details: { type: String, default: '' }
}, { timestamps: true });

const History = mongoose.model('History', historySchema);

const feedbackSchema = new mongoose.Schema({
  userId: { type: String },
  name: { type: String, default: 'Anonymous' },
  email: { type: String, default: 'noreply@namster.com' },
  message: { type: String, required: true },
  status: { type: String, default: 'pending' },
  response: { type: String, default: '' },
  respondedAt: { type: Date }
}, { timestamps: true });

const Feedback = mongoose.model('Feedback', feedbackSchema);

const activeTokens = new Map(); // token -> userId (Keep sessions fast in-memory)

const configuredAdminEmails = (process.env.ADMIN_EMAILS || '')
  .split(',')
  .map(e => e.trim().toLowerCase())
  .filter(Boolean);
const adminEmails = new Set([
  'galileokazadi45@gmail.com',
  ...configuredAdminEmails
]);

function resolveRoleByEmail(email) {
  const normalized = String(email || '').trim().toLowerCase();
  return adminEmails.has(normalized) ? 'admin' : 'user';
}

async function getPublicUser(user) {
  const history = await History.find({ userId: user.id }).sort({ createdAt: -1 }).lean();
  return {
    id: user.id,
    email: user.email,
    firstName: user.firstName || '',
    lastName: user.lastName || '',
    role: user.role || 'user',
    isPremium: Boolean(user.isPremium),
    avatar: user.avatar || '',
    favoriteTemplateIds: Array.isArray(user.favoriteTemplateIds) ? user.favoriteTemplateIds : [],
    history: history || []
  };
}

async function toAdminUser(user) {
  const history = await History.find({ userId: user.id }).sort({ createdAt: -1 }).lean();
  return {
    _id: user.id,
    id: user.id,
    email: user.email,
    firstName: user.firstName || '',
    lastName: user.lastName || '',
    role: user.role || 'user',
    isPremium: Boolean(user.isPremium),
    isBanned: Boolean(user.isBanned),
    createdAt: user.createdAt || new Date().toISOString(),
    updatedAt: user.updatedAt || user.createdAt || new Date().toISOString(),
    history: history || []
  };
}

function issueToken(userId) {
  const token = crypto.randomBytes(24).toString('hex');
  activeTokens.set(token, userId);
  return token;
}

function extractInvitationsCount(entry) {
  if (!entry || typeof entry.details !== 'string') return 0;
  const match = entry.details.match(/(\d+)/);
  return match ? Number(match[1]) || 0 : 0;
}

async function buildUserStats(userId) {
  const history = await History.find({ userId }).lean();
  const generationEntries = history.filter(h => h.action === 'generation');
  const invitationsGenerated = generationEntries.reduce((sum, entry) => sum + extractInvitationsCount(entry), 0);
  const generationOperations = generationEntries.length;
  const savedHours = Math.round(invitationsGenerated * 0.12);

  return {
    invitationsGenerees: invitationsGenerated,
    operationsGenerations: generationOperations,
    precision: 100,
    heuresGagnees: savedHours,
    ['invitationsGénérées']: invitationsGenerated,
    ['opérationsGénérations']: generationOperations,
    ['heuresGagnées']: savedHours
  };
}

async function authRequired(req, res, next) {
  const authHeader = req.headers.authorization || '';
  if (!authHeader.startsWith('Bearer ')) {
    return res.status(401).json({ error: 'Missing bearer token' });
  }

  const token = authHeader.slice(7).trim();
  const userId = activeTokens.get(token);
  if (!userId) {
    return res.status(401).json({ error: 'Invalid or expired token' });
  }

  const user = await User.findOne({ id: userId });
  if (!user) {
    return res.status(401).json({ error: 'User not found for token' });
  }

  req.user = user;
  req.token = token;
  next();
}

function adminRequired(req, res, next) {
  if ((req.user?.role || 'user') !== 'admin') {
    return res.status(403).json({ error: 'Admin access required' });
  }
  next();
}

app.post('/api/auth/register', async (req, res) => {
  const { email, password, firstName, lastName } = req.body;
  if (!email || !password) {
    return res.status(400).json({ error: 'Email and password required' });
  }

  const normalizedEmail = String(email).trim().toLowerCase();
  const existing = await User.findOne({ email: normalizedEmail });
  if (existing) {
    return res.status(409).json({ error: 'Email already registered' });
  }

  const user = new User({
    id: Date.now().toString(36) + Math.random().toString(36).slice(2, 8),
    email: normalizedEmail,
    password,
    firstName: firstName || '',
    lastName: lastName || '',
    role: resolveRoleByEmail(normalizedEmail)
  });
  await user.save();

  const token = issueToken(user.id);
  res.json({ token, user: await getPublicUser(user) });
});

app.post('/api/auth/login', async (req, res) => {
  const { email, password } = req.body;
  const normalizedEmail = String(email || '').trim().toLowerCase();
  const user = await User.findOne({ email: normalizedEmail, password });

  if (!user) {
    return res.status(401).json({ error: 'Invalid credentials' });
  }
  if (user.isBanned) {
    return res.status(403).json({ error: 'Account is banned' });
  }

  user.role = resolveRoleByEmail(user.email);
  await user.save();

  const token = issueToken(user.id);
  res.json({ token, user: await getPublicUser(user) });
});

app.post('/api/auth/google', async (req, res) => {
  const { accessToken, idToken } = req.body || {};
  if (!accessToken && !idToken) {
    return res.status(400).json({ error: 'Google token is required (accessToken or idToken)' });
  }

  try {
    let profile;

    if (accessToken) {
      const googleRes = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
        headers: { Authorization: `Bearer ${accessToken}` }
      });
      if (!googleRes.ok) {
        return res.status(401).json({ error: 'Invalid Google access token' });
      }
      const data = await googleRes.json();
      profile = {
        email: data.email,
        firstName: data.given_name || '',
        lastName: data.family_name || '',
        sub: data.sub
      };
    } else {
      const tokenInfoRes = await fetch(`https://oauth2.googleapis.com/tokeninfo?id_token=${encodeURIComponent(idToken)}`);
      if (!tokenInfoRes.ok) {
        return res.status(401).json({ error: 'Invalid Google id token' });
      }
      const data = await tokenInfoRes.json();
      profile = {
        email: data.email,
        firstName: data.given_name || '',
        lastName: data.family_name || '',
        sub: data.sub
      };
    }

    if (!profile?.email) {
      return res.status(400).json({ error: 'Google profile missing email' });
    }

    const email = String(profile.email).trim().toLowerCase();
    let user = await User.findOne({ email });
    if (!user) {
      user = new User({
        id: profile.sub ? `google-${profile.sub}` : ('google-' + Date.now()),
        email,
        password: null,
        firstName: profile.firstName,
        lastName: profile.lastName,
        role: resolveRoleByEmail(email)
      });
      await user.save();
    } else {
      user.firstName = profile.firstName || user.firstName;
      user.lastName = profile.lastName || user.lastName;
      user.role = resolveRoleByEmail(email);
      await user.save();
    }

    const token = issueToken(user.id);
    return res.json({ token, user: await getPublicUser(user) });
  } catch (error) {
    console.error('Google auth error:', error?.message || error);
    return res.status(500).json({ error: 'Google authentication failed' });
  }
});

app.get('/api/user/profile', authRequired, async (req, res) => {
  res.json(await getPublicUser(req.user));
});

app.put('/api/auth/profile', authRequired, async (req, res) => {
  const { firstName, lastName, email } = req.body || {};
  const nextEmail = email ? String(email).trim().toLowerCase() : req.user.email;

  const emailTaken = await User.findOne({ email: nextEmail, id: { $ne: req.user.id } });
  if (emailTaken) {
    return res.status(409).json({ error: 'Email already in use' });
  }

  req.user.firstName = firstName ?? req.user.firstName;
  req.user.lastName = lastName ?? req.user.lastName;
  req.user.email = nextEmail;
  req.user.role = resolveRoleByEmail(nextEmail);
  await req.user.save();

  res.json(await getPublicUser(req.user));
});

    res.status(500).json({ error: 'Failed to process image' });
  }
});

app.put('/api/auth/profile/avatar', authRequired, upload.single('avatar'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const buffer = await sharp(req.file.path)
      .resize(200, 200)
      .toFormat('webp')
      .toBuffer();
    const base64Avatar = `data:image/webp;base64,${buffer.toString('base64')}`;
    req.user.avatar = base64Avatar;
    await req.user.save();
    try { fs.unlinkSync(req.file.path); } catch (e) { }
    res.json({ avatar: base64Avatar });
  } catch (error) {
    console.error('Avatar upload error:', error);
    res.status(500).json({ error: 'Failed to process image' });
  }
});

app.get('/api/user/stats', authRequired, async (req, res) => {
  res.json(await buildUserStats(req.user.id));
});

app.post('/api/user/history', authRequired, async (req, res) => {
  const { action, details } = req.body || {};
  if (!action) {
    return res.status(400).json({ error: 'Action is required' });
  }

  const entry = new History({
    id: crypto.randomUUID(),
    userId: req.user.id,
    action: String(action),
    details: String(details || '')
  });
  await entry.save();

  res.status(201).json(entry);
});

// --- Saved templates (personnalisées nommées) ---
app.get('/api/user/saved-templates', authRequired, async (req, res) => {
  const list = await SavedTemplate.find({ userId: req.user.id }).sort({ updatedAt: -1 }).lean();
  res.json(list);
});

app.post('/api/user/saved-templates', authRequired, async (req, res) => {
  const { name, templateId, customizationData } = req.body || {};
  if (!name || typeof name !== 'string' || name.trim() === '') {
    return res.status(400).json({ error: 'Name is required' });
  }
  if (templateId == null || Number.isNaN(Number(templateId))) {
    return res.status(400).json({ error: 'templateId is required' });
  }
  const id = crypto.randomUUID();
  const doc = new SavedTemplate({
    id,
    userId: req.user.id,
    name: String(name).trim(),
    templateId: Number(templateId),
    customizationData: customizationData && typeof customizationData === 'object' ? customizationData : {}
  });
  await doc.save();
  res.status(201).json(doc);
});

app.patch('/api/user/saved-templates/:id', authRequired, async (req, res) => {
  const doc = await SavedTemplate.findOne({ id: req.params.id, userId: req.user.id });
  if (!doc) return res.status(404).json({ error: 'Saved template not found' });
  const { name } = req.body || {};
  if (typeof name === 'string' && name.trim() !== '') {
    doc.name = name.trim();
    await doc.save();
  }
  res.json(doc);
});

app.delete('/api/user/saved-templates/:id', authRequired, async (req, res) => {
  const result = await SavedTemplate.findOneAndDelete({ id: req.params.id, userId: req.user.id });
  if (!result) return res.status(404).json({ error: 'Saved template not found' });
  res.json({ success: true });
});

// --- Favoris (template de base) ---
app.get('/api/user/favorites', authRequired, async (req, res) => {
  const user = await User.findOne({ id: req.user.id }).lean();
  const ids = Array.isArray(user?.favoriteTemplateIds) ? user.favoriteTemplateIds : [];
  res.json({ favoriteTemplateIds: ids });
});

app.post('/api/user/favorites', authRequired, async (req, res) => {
  const { templateId } = req.body || {};
  const tid = Number(templateId);
  if (Number.isNaN(tid)) return res.status(400).json({ error: 'templateId is required' });
  const user = await User.findOne({ id: req.user.id });
  if (!user.favoriteTemplateIds) user.favoriteTemplateIds = [];
  if (!user.favoriteTemplateIds.includes(tid)) {
    user.favoriteTemplateIds.push(tid);
    await user.save();
  }
  res.json({ favoriteTemplateIds: user.favoriteTemplateIds });
});

app.delete('/api/user/favorites/:templateId', authRequired, async (req, res) => {
  const tid = Number(req.params.templateId);
  if (Number.isNaN(tid)) return res.status(400).json({ error: 'Invalid templateId' });
  const user = await User.findOne({ id: req.user.id });
  if (user.favoriteTemplateIds) {
    user.favoriteTemplateIds = user.favoriteTemplateIds.filter(id => id !== tid);
    await user.save();
  }
  res.json({ favoriteTemplateIds: user.favoriteTemplateIds || [] });
});

app.get('/api/contact/user/:userId', authRequired, async (req, res) => {
  const isAdmin = (req.user?.role || 'user') === 'admin';
  if (!isAdmin && String(req.user.id) !== String(req.params.userId)) {
    return res.status(403).json({ error: 'Forbidden' });
  }
  const rows = await Feedback.find({ userId: req.params.userId }).sort({ createdAt: -1 }).lean();
  res.json(rows);
});

app.get('/api/admin/stats', authRequired, adminRequired, async (req, res) => {
  const users = await User.countDocuments();
  const visitors = Math.max(users, 1);
  const feedbacksCount = await Feedback.countDocuments();
  const totalVisits = Math.max(users * 3 + feedbacksCount, visitors);
  const pendingFeedback = await Feedback.countDocuments({ status: 'pending' });
  res.json({ users, visitors, totalVisits, pendingFeedback });
});

app.get('/api/admin/users', authRequired, adminRequired, async (req, res) => {
  const users = await User.find().lean();
  const wrappedUsers = await Promise.all(users.map(u => toAdminUser(u)));
  res.json(wrappedUsers);
});

app.get('/api/admin/visitors', authRequired, adminRequired, async (req, res) => {
  const users = await User.find().lean();
  const visitors = users.map(u => ({
    _id: `visit-${u.id}`,
    userId: u.id,
    name: `${u.firstName || ''} ${u.lastName || ''}`.trim() || 'Utilisateur',
    email: u.email,
    ip: '127.0.0.1',
    userAgent: 'Browser',
    createdAt: u.updatedAt || u.createdAt || new Date().toISOString()
  }));
  res.json(visitors);
});

app.get('/api/admin/feedbacks', authRequired, adminRequired, async (req, res) => {
  const rows = await Feedback.find().sort({ createdAt: -1 }).lean();
  res.json(rows);
});

app.post('/api/admin/feedback/:id/respond', authRequired, adminRequired, async (req, res) => {
  const { response } = req.body || {};
  if (!response) return res.status(400).json({ error: 'Response is required' });

  const item = await Feedback.findById(req.params.id);
  if (!item) return res.status(404).json({ error: 'Feedback not found' });

  item.response = String(response);
  item.status = 'replied';
  item.respondedAt = new Date();
  await item.save();
  res.json(item);
});

app.patch('/api/admin/users/:id/status', authRequired, adminRequired, async (req, res) => {
  const target = await User.findOne({ id: req.params.id });
  if (!target) return res.status(404).json({ error: 'User not found' });

  const { isPremium, isBanned } = req.body || {};
  if (typeof isPremium === 'boolean') target.isPremium = isPremium;
  if (typeof isBanned === 'boolean') target.isBanned = isBanned;
  await target.save();

  res.json(await toAdminUser(target));
});

app.put('/api/admin/users/:id', authRequired, adminRequired, async (req, res) => {
  const target = await User.findOne({ id: req.params.id });
  if (!target) return res.status(404).json({ error: 'User not found' });

  const { firstName, lastName, email, role } = req.body || {};
  const nextEmail = email ? String(email).trim().toLowerCase() : target.email;
  const existing = await User.findOne({ email: nextEmail, id: { $ne: target.id } });
  if (existing) return res.status(409).json({ error: 'Email already in use' });

  target.firstName = firstName ?? target.firstName;
  target.lastName = lastName ?? target.lastName;
  target.email = nextEmail;
  target.role = resolveRoleByEmail(nextEmail);
  if (target.role !== 'admin' && (role === 'admin' || role === 'user')) {
    target.role = role;
  }
  await target.save();
  res.json(await toAdminUser(target));
});

app.delete('/api/admin/users/:id', authRequired, adminRequired, async (req, res) => {
  const targetId = String(req.params.id);
  if (String(req.user.id) === targetId) {
    return res.status(400).json({ error: 'You cannot delete your own account' });
  }

  const removed = await User.findOneAndDelete({ id: targetId });
  if (!removed) return res.status(404).json({ error: 'User not found' });

  await History.deleteMany({ userId: targetId });
  await SavedTemplate.deleteMany({ userId: targetId });
  for (const [token, uid] of activeTokens.entries()) {
    if (String(uid) === String(targetId)) activeTokens.delete(token);
  }
  res.json({ success: true });
});

app.post('/api/contact', express.json(), async (req, res) => {
  const { userId, name, email, message } = req.body || {};
  console.log('[Contact] Message from ' + (email || 'unknown-email'));

  if (!message) {
    return res.status(400).json({ error: 'Message is required.' });
  }

  const feedback = new Feedback({
    userId: userId || null,
    name: name || 'Anonymous',
    email: email || 'noreply@namster.com',
    message: String(message),
    status: 'pending',
  });
  await feedback.save();

  try {
    const key = process.env.WEB3FORMS_ACCESS_KEY || '00b59229-53c0-4777-9e71-8a937ab48a60';

    const response = await fetch('https://api.web3forms.com/submit', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      body: JSON.stringify({
        access_key: key,
        name: name || 'Anonymous',
        email: email || 'noreply@namster.com',
        message: String(message),
        subject: `Namster Contact: ${name || 'Anonymous'}`,
        from_name: 'Namster Premium'
      })
    });

    const responseText = await response.text();
    let data = null;
    try {
      data = JSON.parse(responseText);
    } catch {
      data = null;
    }

    if (data?.success) {
      return res.json({ success: true, message: 'Message sent successfully.', feedback });
    }

    return res.json({
      success: true,
      message: 'Message saved locally. External mail provider failed.',
      feedback,
      warning: data?.message || 'External provider error'
    });
  } catch (error) {
    console.error('[Contact] API error:', error.message);
    return res.json({
      success: true,
      message: 'Message saved locally.',
      feedback,
      warning: error.message
    });
  }
});

app.listen(port, () => {
  console.log(`Namster Premium server running at http://localhost:${port}`);
});
