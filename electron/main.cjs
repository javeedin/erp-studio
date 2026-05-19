const {
  app, BrowserWindow, Tray, Menu, dialog, ipcMain,
  Notification, nativeImage, shell, safeStorage,
} = require('electron');
const path = require('path');
const fs   = require('fs');

let nodemailer = null;
try { nodemailer = require('nodemailer'); } catch (_) {}

const getUserDataPath = () => app.getPath('userData');

const APEX_BASE = 'https://g15d6279501ae08-buimerc.adb.me-dubai-1.oraclecloudapps.com/ords/bcldifc/reerp';

// ── SMTP config ─────────────────────────────────────────────────────────────
let _smtpCache = null;
let _smtpCacheAt = 0;
const SMTP_CACHE_TTL = 10 * 60 * 1000;

async function loadSmtpConfig() {
  if (_smtpCache && (Date.now() - _smtpCacheAt) < SMTP_CACHE_TTL) return _smtpCache;
  try {
    const res = await fetch(`${APEX_BASE}/config/emailsettings`);
    const data = await res.json();
    if (data.status === 'success') {
      _smtpCache = { host: data.host, port: data.port, secure: data.secure === true || data.secure === 'true', user: data.user, pass: data.pass, fromName: data.fromName };
      _smtpCacheAt = Date.now();
      return _smtpCache;
    }
  } catch (e) { console.warn('[email] APEX fetch failed:', e.message); }
  const candidates = [path.join(__dirname, 'email.config.json')];
  try { candidates.push(path.join(process.resourcesPath, 'email.config.json')); } catch (_) {}
  for (const p of candidates) {
    try {
      if (fs.existsSync(p)) {
        const raw = fs.readFileSync(p, 'utf8').replace(/^﻿/, '');
        return JSON.parse(raw);
      }
    } catch (_) {}
  }
  return null;
}

ipcMain.handle('send-otp-email', async (_event, { to, otp }) => {
  if (!nodemailer) return { success: false, error: 'nodemailer not available' };
  try {
    const cfg = await loadSmtpConfig();
    if (!cfg) return { success: false, error: 'No SMTP config found' };
    const transporter = nodemailer.createTransport({ host: cfg.host, port: cfg.port, secure: cfg.secure, auth: { user: cfg.user, pass: cfg.pass } });
    await transporter.sendMail({
      from: `"${cfg.fromName || 'ERP Studio'}" <${cfg.user}>`,
      to,
      subject: 'Your ERP Studio OTP Code',
      html: `<div style="font-family:sans-serif;max-width:400px"><h2>Your OTP Code</h2><p style="font-size:32px;font-weight:bold;color:#C74634;letter-spacing:8px">${otp}</p><p>This code expires in 10 minutes.</p></div>`,
    });
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// ── Notification ─────────────────────────────────────────────────────────────
ipcMain.on('show-notification', (_event, title, body) => {
  if (Notification.isSupported()) new Notification({ title, body }).show();
});

// ── Screen recording ─────────────────────────────────────────────────────────
ipcMain.handle('get-screen-sources', async () => {
  const { desktopCapturer } = require('electron');
  const sources = await desktopCapturer.getSources({
    types: ['screen', 'window'],
    thumbnailSize: { width: 300, height: 200 },
  });
  return sources.map(s => ({
    id:        s.id,
    name:      s.name,
    thumbnail: s.thumbnail.toDataURL(),
  }));
});

ipcMain.handle('save-recording', async (_event, { buffer, metadata }) => {
  const dir = path.join(getUserDataPath(), 'recordings');
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  const id       = `rec_${Date.now()}`;
  const fileName = `${id}.webm`;
  const filePath = path.join(dir, fileName);
  const metaPath = path.join(dir, `${id}.json`);
  fs.writeFileSync(filePath, Buffer.from(buffer));
  const meta = {
    id, fileName, filePath,
    title:       metadata.title       || 'Recording',
    description: metadata.description || '',
    category:    metadata.category    || 'General',
    duration:    metadata.duration    || 0,
    fileSize:    fs.statSync(filePath).size,
    createdAt:   new Date().toISOString(),
    createdBy:   metadata.createdBy   || '',
  };
  fs.writeFileSync(metaPath, JSON.stringify(meta, null, 2));
  return { success: true, id, filePath };
});

// ── Training library ─────────────────────────────────────────────────────────
ipcMain.handle('list-recordings', async () => {
  const dir = path.join(getUserDataPath(), 'recordings');
  if (!fs.existsSync(dir)) return [];
  const files = fs.readdirSync(dir).filter(f => f.endsWith('.json'));
  const recs = [];
  for (const f of files) {
    try { recs.push(JSON.parse(fs.readFileSync(path.join(dir, f), 'utf8'))); } catch (_) {}
  }
  return recs.sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
});

ipcMain.handle('delete-recording', async (_event, { id, filePath }) => {
  try {
    if (filePath && fs.existsSync(filePath)) fs.unlinkSync(filePath);
    const dir      = path.join(getUserDataPath(), 'recordings');
    const metaPath = path.join(dir, `${id}.json`);
    if (fs.existsSync(metaPath)) fs.unlinkSync(metaPath);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('open-recordings-folder', async () => {
  const dir = path.join(getUserDataPath(), 'recordings');
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  shell.openPath(dir);
});

ipcMain.handle('get-file-url', (_event, filePath) => {
  return `file://${filePath.replace(/\\/g, '/')}`;
});

// ── ERP session persistence ──────────────────────────────────────────────────
const SESSION_PATH = () => path.join(getUserDataPath(), 'erp-session.json');

ipcMain.handle('save-erp-session', (_event, { user, token }) => {
  try {
    fs.writeFileSync(SESSION_PATH(), JSON.stringify({ user, token }), 'utf8');
    return { success: true };
  } catch (e) { return { success: false }; }
});

ipcMain.handle('get-erp-session', () => {
  try {
    const p = SESSION_PATH();
    if (!fs.existsSync(p)) return null;
    return JSON.parse(fs.readFileSync(p, 'utf8'));
  } catch (_) { return null; }
});

ipcMain.handle('clear-erp-session', () => {
  try { fs.unlinkSync(SESSION_PATH()); } catch (_) {}
  return { success: true };
});

// ── Oracle Fusion credentials (OS keychain via safeStorage) ──────────────────
const CRED_PATH = () => path.join(getUserDataPath(), 'fusion-creds.bin');

ipcMain.handle('save-fusion-credentials', (_event, { username, password }) => {
  try {
    if (safeStorage.isEncryptionAvailable()) {
      const enc = safeStorage.encryptString(JSON.stringify({ username, password }));
      fs.writeFileSync(CRED_PATH(), enc);
    } else {
      fs.writeFileSync(CRED_PATH(), JSON.stringify({ username, password }), 'utf8');
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
});

ipcMain.handle('get-fusion-credentials', () => {
  try {
    const p = CRED_PATH();
    if (!fs.existsSync(p)) return null;
    const raw = fs.readFileSync(p);
    if (safeStorage.isEncryptionAvailable()) {
      return JSON.parse(safeStorage.decryptString(raw));
    }
    return JSON.parse(raw.toString('utf8'));
  } catch (_) { return null; }
});

ipcMain.handle('clear-fusion-credentials', () => {
  try { fs.unlinkSync(CRED_PATH()); } catch (_) {}
  return { success: true };
});

// ── Window setup ─────────────────────────────────────────────────────────────
let mainWindow = null;
let tray = null;

function createWindow() {
  mainWindow = new BrowserWindow({
    width:  1400,
    height: 900,
    minWidth:  900,
    minHeight: 600,
    icon: path.join(__dirname, '../public/icons/icon-512.png'),
    webPreferences: {
      preload:         path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration:  false,
      webviewTag:       true,   // required for Oracle Fusion embedded browser
      webSecurity:      false,  // required for cross-origin webview content
    },
    titleBarStyle: 'default',
    show: false,
  });

  const isDev = !app.isPackaged;
  if (isDev) {
    mainWindow.loadURL('http://localhost:5173');
    mainWindow.webContents.openDevTools({ mode: 'detach' });
  } else {
    mainWindow.loadFile(path.join(__dirname, '../dist/index.html'));
  }

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
    mainWindow.focus();
  });

  mainWindow.on('closed', () => { mainWindow = null; });
}

app.whenReady().then(() => {
  createWindow();

  // Tray icon
  try {
    const iconPath = path.join(__dirname, '../public/icons/icon-128.png');
    if (fs.existsSync(iconPath)) {
      tray = new Tray(nativeImage.createFromPath(iconPath));
      tray.setToolTip('ERP Studio');
      tray.setContextMenu(Menu.buildFromTemplate([
        { label: 'Open ERP Studio', click: () => { if (mainWindow) mainWindow.show(); else createWindow(); } },
        { type: 'separator' },
        { label: 'Quit', click: () => app.quit() },
      ]));
      tray.on('double-click', () => { if (mainWindow) mainWindow.show(); else createWindow(); });
    }
  } catch (_) {}
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
