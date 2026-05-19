const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  // Auth
  sendOtpEmail:     (to, otp) => ipcRenderer.invoke('send-otp-email', { to, otp }),
  showNotification: (title, body) => ipcRenderer.send('show-notification', title, body),

  // Screen recording
  getScreenSources: ()                 => ipcRenderer.invoke('get-screen-sources'),
  saveRecording:    (buffer, metadata) => ipcRenderer.invoke('save-recording', { buffer, metadata }),

  // Training library
  listRecordings:       ()             => ipcRenderer.invoke('list-recordings'),
  deleteRecording:      (id, filePath) => ipcRenderer.invoke('delete-recording', { id, filePath }),
  openRecordingsFolder: ()             => ipcRenderer.invoke('open-recordings-folder'),
  getFileUrl:           (filePath)     => ipcRenderer.invoke('get-file-url', filePath),

  // ERP session persistence
  saveErpSession:  (user, token) => ipcRenderer.invoke('save-erp-session', { user, token }),
  getErpSession:   ()            => ipcRenderer.invoke('get-erp-session'),
  clearErpSession: ()            => ipcRenderer.invoke('clear-erp-session'),

  // Oracle Fusion saved credentials
  saveFusionCredentials:  (username, password) => ipcRenderer.invoke('save-fusion-credentials', { username, password }),
  getFusionCredentials:   ()                   => ipcRenderer.invoke('get-fusion-credentials'),
  clearFusionCredentials: ()                   => ipcRenderer.invoke('clear-fusion-credentials'),

  // Platform
  isElectron: true,
  platform:   process.platform,
});
