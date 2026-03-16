import mongoose from 'mongoose';
import crypto from 'crypto';
import dotenv from 'dotenv';

dotenv.config();

const MONGO_URI = process.env.MONGO_URI || "mongodb://localhost:27017/namster";

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

const ADMIN_DATA = {
  id: 'admin-' + Date.now(),
  email: 'galileokazadi45@gmail.com',
  password: crypto.randomBytes(16).toString('hex'), // Mot de passe aléatoire, à changer via login
  firstName: 'Nisan-Galiléé',
  lastName: 'Kazadi',
  role: 'admin',
  isPremium: true,
  isBanned: false
};

async function seedAdmin() {
  try {
    await mongoose.connect(MONGO_URI);
    console.log('Connected to MongoDB');

    // Check if admin already exists
    const existing = await User.findOne({ email: ADMIN_DATA.email });
    if (existing) {
      console.log('Admin already exists, updating role to admin...');
      existing.role = 'admin';
      existing.isPremium = true;
      existing.firstName = ADMIN_DATA.firstName;
      existing.lastName = ADMIN_DATA.lastName;
      await existing.save();
      console.log('Admin updated:', existing.email);
    } else {
      const admin = new User(ADMIN_DATA);
      await admin.save();
      console.log('Admin created successfully!');
      console.log('Email:', admin.email);
      console.log('Name:', admin.firstName, admin.lastName);
      console.log('Role:', admin.role);
    }

    await mongoose.disconnect();
    console.log('Done');
  } catch (err) {
    console.error('Error:', err);
    process.exit(1);
  }
}

seedAdmin();
