import sys
import re

with open(r"d:\Namsterreact\backend\server.js", "r", encoding="utf-8") as f:
    text = f.read()

# Find the separator "// --- Auth + User Routes (Mock/In-memory Implementation) ---"
sep = "// --- Auth + User Routes (Mock/In-memory Implementation) ---"
parts = text.split(sep)
if len(parts) != 2:
    print("Separator not found or multiple found.")
    sys.exit(1)

top_code = parts[0]

# Now we append the Mongoose logic and rewritten routes
mongo_logic = """// --- Auth + User Routes (MongoDB Implementation) ---
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
  avatar: { type: String, default: '' }
}, { timestamps: true });

const User = mongoose.model('User', userSchema);

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
  const match = entry.details.match(/(\\d+)/);
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
"""

import codecs
with codecs.open(r"d:\Namsterreact\backend\server.js", "w", "utf-8") as f:
    f.write(top_code + mongo_logic)

print("Patch successful!")
