const { app, BrowserWindow, ipcMain, shell } = require("electron");
const path = require("node:path");
const crypto = require("node:crypto");
const fs = require("node:fs/promises");
const { pathToFileURL } = require("node:url");

const DATA_FOLDER_NAME = "Mis Contactes Privats";
const CONTACTS_FILE_NAME = "contactes.json";
const BACKUPS_FOLDER_NAME = "backups";
const PHOTOS_FOLDER_NAME = "fotos";
const AUTO_BACKUP_FILE_NAME = "contactes-auto-backup.json";
let resolvedDataRoot = null;

app.disableHardwareAcceleration();

function getDataRoot() {
  return resolvedDataRoot || path.join(app.getPath("desktop"), DATA_FOLDER_NAME);
}

function getContactsPath() {
  return path.join(getDataRoot(), CONTACTS_FILE_NAME);
}

async function ensureDataLayout() {
  const preferredRoots = [
    path.join(app.getPath("documents"), DATA_FOLDER_NAME),
    path.join(app.getPath("desktop"), DATA_FOLDER_NAME),
    path.join(app.getPath("userData"), DATA_FOLDER_NAME),
  ];

  let lastError = null;
  for (const candidateRoot of preferredRoots) {
    try {
      const backupsDir = path.join(candidateRoot, BACKUPS_FOLDER_NAME);
      const photosDir = path.join(candidateRoot, PHOTOS_FOLDER_NAME);
      await fs.mkdir(backupsDir, { recursive: true });
      await fs.mkdir(photosDir, { recursive: true });
      const contactsPath = path.join(candidateRoot, CONTACTS_FILE_NAME);
      try {
        await fs.access(contactsPath);
      } catch {
        await fs.writeFile(contactsPath, "[]", "utf8");
      }
      resolvedDataRoot = candidateRoot;
      return {
        dataRoot: candidateRoot,
        contactsPath,
        backupsDir,
        photosDir,
      };
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error("No he pogut preparar cap carpeta local per als contactes.");
}

function toPhotoFileUrl(photoPath, photosDir) {
  if (!photoPath) return "";
  return pathToFileURL(path.join(photosDir, photoPath)).href;
}

async function hydrateDesktopContacts(contacts) {
  const { photosDir } = await ensureDataLayout();
  return (Array.isArray(contacts) ? contacts : []).map((contact) => {
    const nextContact = { ...contact };
    if (nextContact.photoPath) nextContact.photoData = toPhotoFileUrl(nextContact.photoPath, photosDir);
    return nextContact;
  });
}

async function readContactsFile() {
  const { contactsPath } = await ensureDataLayout();
  const raw = await fs.readFile(contactsPath, "utf8");
  const parsed = JSON.parse(raw);
  return hydrateDesktopContacts(parsed);
}

async function writeContactsFile(contacts) {
  const { contactsPath, backupsDir } = await ensureDataLayout();
  const normalizedContacts = (Array.isArray(contacts) ? contacts : []).map((contact) => {
    const nextContact = { ...contact };
    if (nextContact.photoPath) nextContact.photoData = toPhotoFileUrl(nextContact.photoPath, path.join(getDataRoot(), PHOTOS_FOLDER_NAME));
    return nextContact;
  });
  const payload = JSON.stringify(normalizedContacts, null, 2);
  const portableBackupPayload = JSON.stringify(await toPortableBackupContacts(normalizedContacts), null, 2);
  await fs.writeFile(contactsPath, payload, "utf8");
  await fs.writeFile(path.join(backupsDir, AUTO_BACKUP_FILE_NAME), portableBackupPayload, "utf8");
  return contactsPath;
}

async function saveBackupJson(payload) {
  const { backupsDir } = await ensureDataLayout();
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const backupPath = path.join(backupsDir, `contactes-backup-${stamp}.json`);
  let parsedPayload = [];
  try {
    parsedPayload = JSON.parse(String(payload || "[]"));
  } catch {
    parsedPayload = [];
  }
  const portablePayload = JSON.stringify(await toPortableBackupContacts(parsedPayload), null, 2);
  await fs.writeFile(backupPath, portablePayload, "utf8");
  return backupPath;
}

function extensionFromMime(mimeType) {
  const map = {
    "image/jpeg": ".jpg",
    "image/png": ".png",
    "image/gif": ".gif",
    "image/webp": ".webp",
    "image/bmp": ".bmp",
    "image/svg+xml": ".svg",
  };
  return map[mimeType] || ".png";
}

function mimeFromExtension(filePath) {
  const ext = path.extname(String(filePath || "")).toLowerCase();
  const map = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".gif": "image/gif",
    ".webp": "image/webp",
    ".bmp": "image/bmp",
    ".svg": "image/svg+xml",
  };
  return map[ext] || "application/octet-stream";
}

async function toPortableBackupContacts(contacts) {
  const { photosDir } = await ensureDataLayout();
  const portableContacts = [];
  for (const contact of Array.isArray(contacts) ? contacts : []) {
    const nextContact = { ...contact };
    if (nextContact.photoPath) {
      try {
        const absolutePhotoPath = path.join(photosDir, nextContact.photoPath);
        const buffer = await fs.readFile(absolutePhotoPath);
        nextContact.photoData = `data:${mimeFromExtension(nextContact.photoPath)};base64,${buffer.toString("base64")}`;
      } catch {
        nextContact.photoData = "";
      }
    }
    portableContacts.push(nextContact);
  }
  return portableContacts;
}

function safePhotoBaseName(name) {
  return String(name || "foto")
    .replace(/\.[a-z0-9]+$/i, "")
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .trim()
    .toLowerCase() || "foto";
}

async function savePhotoDataUrl(dataUrl, originalName) {
  const match = String(dataUrl || "").match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!match) throw new Error("La foto no te un format valid.");

  const [, mimeType, base64Payload] = match;
  const fileExt = extensionFromMime(mimeType);
  const fileName = `${safePhotoBaseName(originalName)}-${crypto.randomUUID()}${fileExt}`;
  const { photosDir } = await ensureDataLayout();
  const photoPath = path.join(photosDir, fileName);
  await fs.writeFile(photoPath, Buffer.from(base64Payload, "base64"));

  return {
    photoPath: fileName,
    photoUrl: pathToFileURL(photoPath).href,
  };
}

function createWindow() {
  const window = new BrowserWindow({
    width: 1540,
    height: 980,
    minWidth: 1180,
    minHeight: 760,
    backgroundColor: "#edf2f8",
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, "electron-preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  window.loadFile(path.join(__dirname, "index.html"));
  window.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url).catch(() => {});
    return { action: "deny" };
  });
}

ipcMain.handle("desktop:is-desktop-app", () => true);
ipcMain.handle("desktop:load-contacts", async () => readContactsFile());
ipcMain.handle("desktop:save-contacts", async (_event, contacts) => writeContactsFile(contacts));
ipcMain.handle("desktop:save-backup-json", async (_event, payload) => saveBackupJson(String(payload || "[]")));
ipcMain.handle("desktop:save-photo-data-url", async (_event, dataUrl, originalName) => savePhotoDataUrl(dataUrl, originalName));
ipcMain.handle("desktop:get-storage-info", async () => {
  const layout = await ensureDataLayout();
  return layout;
});

app.whenReady().then(async () => {
  await ensureDataLayout();
  createWindow();

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

process.on("unhandledRejection", (error) => {
  console.error("[main-unhandled-rejection]", error);
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
