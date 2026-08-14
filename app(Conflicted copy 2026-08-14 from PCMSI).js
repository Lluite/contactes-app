const STORAGE_KEY = "lluis-contactes-nova-app-v1";
const NOTIFY_DAY_KEY = "lluis-contactes-last-notify-day";
const DATE_KEYS = ["date1", "date2", "date3", "date4"];
const now = new Date();
const TODAY = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 12, 0, 0);
const MONTH_NAMES = [
  "enero",
  "febrero",
  "marzo",
  "abril",
  "mayo",
  "junio",
  "julio",
  "agosto",
  "septiembre",
  "octubre",
  "noviembre",
  "diciembre",
];
const MONTH_MAP = {
  enero: 1,
  febrero: 2,
  marzo: 3,
  abril: 4,
  mayo: 5,
  junio: 6,
  julio: 7,
  agosto: 8,
  septiembre: 9,
  setiembre: 9,
  octubre: 10,
  noviembre: 11,
  diciembre: 12,
  gener: 1,
  febrer: 2,
  marc: 3,
  abril: 4,
  maig: 5,
  juny: 6,
  juliol: 7,
  agost: 8,
  setembre: 9,
  octubre: 10,
  novembre: 11,
  desembre: 12,
};
const FIELD_ORDER = [
  "name",
  "group",
  "phoneWork",
  "phoneHome",
  "phoneMobile",
  "phoneMobileType",
  "date1",
  "date1Type",
  "date2",
  "date2Type",
  "date3",
  "date3Type",
  "date4",
  "date4Type",
  "mail1",
  "mail2",
  "address",
  "postalCode",
  "city",
  "gpsLat",
  "gpsLng",
];
const FIELD_LABELS = {
  name: "Nom",
  group: "Grup",
  phoneWork: "Telf. Treball",
  phoneHome: "Telf. Casa",
  phoneMobile: "Telf. Mobil",
  phoneMobileType: "Tipus telf. Mobil",
  date1: "Data 1",
  date1Type: "Tipus data 1",
  date2: "Data 2",
  date2Type: "Tipus data 2",
  date3: "Data 3",
  date3Type: "Tipus data 3",
  date4: "Data 4",
  date4Type: "Tipus data 4",
  mail1: "Mail 1",
  mail2: "Mail 2",
  address: "Adreca",
  postalCode: "C.P.",
  city: "Poblacio",
  gpsLat: "Latitud",
  gpsLng: "Longitud",
};
const IMPORT_ALIASES = {
  name: ["Nom", "Nombre", "Nombre formateado"],
  group: ["Grup", "Grupo", "Empresa"],
  phoneWork: ["Telf. Treball", "Telefono trabajo", "Telf. Trabajo"],
  phoneHome: ["Telf. Casa", "Telefono casa"],
  phoneMobile: ["Telf. Mobil", "Telf. Mòbil", "Telefono", "Teléfono", "Telf. 1"],
  phoneMobileType: ["Tipus telf. Mobil", "Tipus telf. Mòbil", "Tipo telefono movil", "Tipo teléfono móvil", "Tipo telf. 1"],
  date1: ["Data 1", "Cumpleaños", "Cumpleanos", "Fecha", "Fecha 1", "Fecha (Otro)"],
  date1Type: ["Tipus data 1", "Tipo fecha 1", "Fecha Tipo"],
  date2: ["Data 2", "Aniversario", "Fecha 2", "Fecha (1)"],
  date2Type: ["Tipus data 2", "Tipo fecha 2", "Fecha Tipo (1)"],
  date3: ["Data 3", "Fecha 3", "Fecha (2)"],
  date3Type: ["Tipus data 3", "Tipo fecha 3", "Fecha Tipo (2)"],
  date4: ["Data 4", "Fecha 4"],
  date4Type: ["Tipus data 4", "Tipo fecha 4"],
  mail1: ["Mail 1", "Email 1", "E-mail"],
  mail2: ["Mail 2", "Email 2"],
  address: ["Adreca", "Dirección", "Direccion"],
  postalCode: ["C.P.", "Codigo postal", "Código postal"],
  city: ["Poblacio", "Población", "Poblacion", "Ciudad"],
  gpsLat: ["Latitud", "GPS Lat"],
  gpsLng: ["Longitud", "GPS Lng"],
};
const REMINDER_WINDOW_DAYS = 30;

const state = {
  contacts: loadContacts(),
  selectedId: null,
  selectedIds: new Set(),
  activeLetter: "",
};

const contactList = document.getElementById("contactList");
const reminderList = document.getElementById("reminderList");
const contactCount = document.getElementById("contactCount");
const upcomingCount = document.getElementById("upcomingCount");
const searchInput = document.getElementById("searchInput");
const editorTitle = document.getElementById("editorTitle");
const contactForm = document.getElementById("contactForm");
const saveBtn = document.getElementById("saveBtn");
const newContactBtn = document.getElementById("newContactBtn");
const importXlsxBtn = document.getElementById("importXlsxBtn");
const exportXlsxBtn = document.getElementById("exportXlsxBtn");
const notificationsBtn = document.getElementById("notificationsBtn");
const xlsxInput = document.getElementById("xlsxInput");
const saveGpsBtn = document.getElementById("saveGpsBtn");
const openGpsBtn = document.getElementById("openGpsBtn");
const gpsStatus = document.getElementById("gpsStatus");
const backToListBtn = document.getElementById("backToListBtn");
const backupBtn = document.getElementById("backupBtn");
const selectFilteredBtn = document.getElementById("selectFilteredBtn");
const deleteSelectedBtn = document.getElementById("deleteSelectedBtn");
const alphabetNav = document.getElementById("alphabetNav");
const mail1Link = document.getElementById("mail1Link");
const mail2Link = document.getElementById("mail2Link");
const editorOverlay = document.getElementById("editorOverlay");
const editorBackdrop = document.getElementById("editorBackdrop");
const closeEditorBtn = document.getElementById("closeEditorBtn");
const deleteContactBtn = document.getElementById("deleteContactBtn");
const uploadPhotoBtn = document.getElementById("uploadPhotoBtn");
const photoInput = document.getElementById("photoInput");
const contactPhotoImage = document.getElementById("contactPhotoImage");
const contactPhotoInitials = document.getElementById("contactPhotoInitials");

function loadContacts() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveContacts() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.contacts));
}

function normalize(text) {
  return String(text || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function createEmptyContact() {
  const contact = { id: crypto.randomUUID(), createdAt: new Date().toISOString() };
  for (const key of FIELD_ORDER) {
    contact[key] = "";
  }
  contact.photoData = "";
  return contact;
}

function selectedContact() {
  return state.contacts.find((contact) => contact.id === state.selectedId) || null;
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function safeFileName(name) {
  return String(name || "contactes")
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function textCell(value) {
  const escaped = escapeHtml(String(value ?? ""));
  return `<c t="inlineStr"><is><t xml:space="preserve">${escaped}</t></is></c>`;
}

function buildSheetXml(rows) {
  const xmlRows = rows.map((row, index) => {
    const cells = row.map((value) => textCell(value)).join("");
    return `<row r="${index + 1}">${cells}</row>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>${xmlRows}</sheetData>
</worksheet>`;
}

async function exportContactsXlsx() {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`);
  zip.folder("_rels").file(".rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);
  zip.folder("xl").file("workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Contactes" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`);
  zip.folder("xl").folder("_rels").file("workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`);

  const rows = [
    FIELD_ORDER.map((key) => FIELD_LABELS[key]),
    ...state.contacts.map((contact) => FIELD_ORDER.map((key) => contact[key] || "")),
  ];
  zip.folder("xl").folder("worksheets").file("sheet1.xml", buildSheetXml(rows));

  const blob = await zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  downloadBlob(blob, `contactes-${safeFileName(formatDateForFile(new Date()))}.xlsx`);
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function exportBackupJson() {
  const blob = new Blob([JSON.stringify(state.contacts, null, 2)], {
    type: "application/json;charset=utf-8",
  });
  downloadBlob(blob, `contactes-backup-${safeFileName(formatDateForFile(new Date()))}.json`);
}

function formatDateForFile(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function excelSerialToDateText(serial) {
  const origin = new Date(Date.UTC(1899, 11, 30));
  origin.setUTCDate(origin.getUTCDate() + Number(serial));
  const year = origin.getUTCFullYear();
  const month = origin.getUTCMonth() + 1;
  const day = origin.getUTCDate();
  return `${String(day).padStart(2, "0")}/${String(month).padStart(2, "0")}/${year}`;
}

function parseWorksheetXml(xmlText, sharedStrings) {
  const parser = new DOMParser();
  const xml = parser.parseFromString(xmlText, "application/xml");
  const rows = [...xml.getElementsByTagName("row")];
  return rows.map((row) => {
    const result = [];
    const cells = [...row.getElementsByTagName("c")];
    for (const cell of cells) {
      const ref = cell.getAttribute("r") || "";
      const colLetters = ref.replace(/[0-9]/g, "");
      const index = columnLettersToIndex(colLetters);
      while (result.length < index) result.push("");
      result.push(readCellValue(cell, sharedStrings));
    }
    return result;
  });
}

function columnLettersToIndex(letters) {
  let total = 0;
  for (const char of letters) {
    total = total * 26 + (char.charCodeAt(0) - 64);
  }
  return Math.max(0, total - 1);
}

function readCellValue(cell, sharedStrings) {
  const type = cell.getAttribute("t");
  if (type === "inlineStr") {
    return [...cell.getElementsByTagName("t")].map((node) => node.textContent || "").join("");
  }
  const valueNode = cell.getElementsByTagName("v")[0];
  const raw = valueNode ? valueNode.textContent || "" : "";
  if (type === "s") {
    return sharedStrings[Number(raw)] || "";
  }
  if (type === "str") {
    return raw;
  }
  return raw;
}

function parseSharedStrings(xmlText) {
  const parser = new DOMParser();
  const xml = parser.parseFromString(xmlText, "application/xml");
  return [...xml.getElementsByTagName("si")].map((node) =>
    [...node.getElementsByTagName("t")].map((part) => part.textContent || "").join("")
  );
}

async function importContactsFromFile(file) {
  const buffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(buffer);
  const sharedStrings = zip.file("xl/sharedStrings.xml")
    ? parseSharedStrings(await zip.file("xl/sharedStrings.xml").async("string"))
    : [];

  const workbookXml = await zip.file("xl/workbook.xml").async("string");
  const parser = new DOMParser();
  const workbook = parser.parseFromString(workbookXml, "application/xml");
  const firstSheet = workbook.getElementsByTagName("sheet")[0];
  if (!firstSheet) throw new Error("No he trobat cap full dins del XLSX.");

  const relId = firstSheet.getAttribute("r:id");
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const relsDoc = parser.parseFromString(relsXml, "application/xml");
  const relation = [...relsDoc.getElementsByTagName("Relationship")].find((rel) => rel.getAttribute("Id") === relId);
  if (!relation) throw new Error("No he pogut localitzar el full principal del XLSX.");

  const target = relation.getAttribute("Target");
  const sheetPath = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
  const sheetXml = await zip.file(sheetPath).async("string");
  const rows = parseWorksheetXml(sheetXml, sharedStrings).filter((row) => row.some((value) => String(value || "").trim()));
  if (rows.length < 2) {
    state.contacts = [];
    state.selectedId = null;
    saveContacts();
    renderAll();
    return;
  }

  const headers = rows[0].map((value) => String(value || "").trim());
  const imported = rows.slice(1).map((row) => mapImportedRow(headers, row)).filter((contact) => contact.name);
  state.contacts = imported;
  state.selectedId = imported[0]?.id || null;
  saveContacts();
  renderAll();
}

function mapImportedRow(headers, row) {
  const contact = createEmptyContact();
  for (const key of FIELD_ORDER) {
    const value = findImportedValue(headers, row, key);
    contact[key] = normalizeImportedValue(key, value);
  }
  return contact;
}

function findImportedValue(headers, row, key) {
  const aliases = IMPORT_ALIASES[key] || [FIELD_LABELS[key]];
  for (let index = 0; index < headers.length; index += 1) {
    const header = headers[index];
    if (aliases.some((alias) => normalize(alias) === normalize(header))) {
      return row[index] ?? "";
    }
  }
  if ((key === "gpsLat" || key === "gpsLng")) {
    const combinedIndex = headers.findIndex((header) => normalize(header) === normalize("Ubicacio GPS"));
    if (combinedIndex !== -1) {
      const parsed = parseCombinedGps(row[combinedIndex] || "");
      return key === "gpsLat" ? parsed.lat : parsed.lng;
    }
  }
  return "";
}

function parseCombinedGps(value) {
  const text = String(value || "").trim();
  const match = text.match(/(-?\d+(?:[.,]\d+)?)\s*[,; ]\s*(-?\d+(?:[.,]\d+)?)/);
  if (!match) return { lat: "", lng: "" };
  return {
    lat: match[1].replace(",", "."),
    lng: match[2].replace(",", "."),
  };
}

function normalizeImportedValue(key, value) {
  const text = String(value ?? "").trim();
  if (!text) return "";
  if (DATE_KEYS.includes(key) && /^[0-9]+(?:\.[0-9]+)?$/.test(text)) {
    return formatDateLong(excelSerialToDateText(text));
  }
  if (DATE_KEYS.includes(key)) {
    return formatDateLong(text);
  }
  return text;
}

function renderAll() {
  ensureSelectedContact();
  renderAlphabetNav();
  renderContactList();
  renderReminders();
  if (!editorOverlay.classList.contains("hidden")) {
    renderSelectedContact();
  }
}

function contactInitials(name) {
  const words = String(name || "").trim().split(/\s+/).filter(Boolean);
  if (!words.length) return "?";
  return words.slice(0, 2).map((word) => word[0]).join("").toUpperCase();
}

function renderAlphabetNav() {
  const letters = ["·", ..."ABCDEFGHIJKLMNOPQRSTUVWXYZ"];
  alphabetNav.innerHTML = "";
  const fragment = document.createDocumentFragment();
  for (const letter of letters) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `alpha-btn${state.activeLetter === letter ? " active" : ""}`;
    button.textContent = letter;
    button.addEventListener("click", () => {
      state.activeLetter = letter === "·" ? "" : letter;
      renderAlphabetNav();
      renderContactList();
    });
    fragment.appendChild(button);
  }
  alphabetNav.appendChild(fragment);
}

function ensureSelectedContact() {
  if (!state.contacts.length) {
    state.selectedId = null;
    return;
  }
  if (!selectedContact()) {
    state.selectedId = state.contacts[0].id;
  }
}

function filteredContacts() {
  const query = normalize(searchInput.value);
  return state.contacts.filter((contact) => {
    if (state.activeLetter) {
      const first = normalize(contact.name).charAt(0).toUpperCase();
      if (first !== state.activeLetter) return false;
    }
    if (!query) return true;
    const haystack = normalize([
      contact.name,
      contact.group,
      contact.phoneMobile,
      contact.phoneWork,
      contact.phoneHome,
      contact.mail1,
      contact.mail2,
      contact.city,
    ].join(" "));
    return haystack.includes(query);
  });
}

function renderContactList() {
  const contacts = filteredContacts();
  contactList.innerHTML = "";
  contactCount.textContent = String(state.contacts.length);
  deleteSelectedBtn.textContent = state.selectedIds.size
    ? `Borrar seleccionados (${state.selectedIds.size})`
    : "Borrar seleccionados";

  if (!contacts.length) {
    contactList.innerHTML = `<div class="empty-state">No hi ha contactes amb aquest text.</div>`;
    return;
  }

  const fragment = document.createDocumentFragment();
  for (const contact of contacts) {
    const wrap = document.createElement("div");
    wrap.className = "contact-row-wrap";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.className = "contact-check";
    checkbox.checked = state.selectedIds.has(contact.id);
    checkbox.setAttribute("aria-label", `Seleccionar ${contact.name || "contacte"}`);
    checkbox.addEventListener("click", (event) => event.stopPropagation());
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) {
        state.selectedIds.add(contact.id);
      } else {
        state.selectedIds.delete(contact.id);
      }
      renderContactList();
    });

    const button = document.createElement("button");
    button.type = "button";
    button.className = `contact-row${contact.id === state.selectedId ? " active" : ""}`;
    button.innerHTML = `
      <p class="contact-name">${escapeHtml(contact.name || "Sense nom")}</p>
      <div class="contact-meta">${escapeHtml(contact.group || "Sense grup")} · ${escapeHtml(contact.phoneMobile || contact.phoneWork || contact.mail1 || "")}</div>
    `;
    button.addEventListener("click", () => {
      state.selectedId = contact.id;
      openEditor();
    });
    wrap.appendChild(checkbox);
    wrap.appendChild(button);
    fragment.appendChild(wrap);
  }
  contactList.appendChild(fragment);
}

function renderSelectedContact() {
  const contact = selectedContact();
  if (!contact) {
    editorTitle.textContent = "Nou contacte";
    contactForm.reset();
    gpsStatus.textContent = "";
    updateMailLinks();
    updateDateInfos();
    return;
  }

  editorTitle.textContent = contact.name || "Contacte sense nom";
  for (const key of FIELD_ORDER) {
    const input = contactForm.elements.namedItem(key);
    if (input) input.value = contact[key] || "";
  }
  gpsStatus.textContent = gpsLabel(contact);
  updateMailLinks();
  updateDateInfos();
}

function gpsLabel(contact) {
  if (!contact.gpsLat || !contact.gpsLng) return "";
  return `GPS guardat: ${contact.gpsLat}, ${contact.gpsLng}`;
}

function saveCurrentContact() {
  const formData = new FormData(contactForm);
  let contact = selectedContact();
  if (!contact) {
    contact = createEmptyContact();
    state.contacts.unshift(contact);
    state.selectedId = contact.id;
  }

  for (const key of FIELD_ORDER) {
    const raw = formData.get(key) || "";
    contact[key] = DATE_KEYS.includes(key) ? formatDateLong(raw) : String(raw).trim();
  }

  contact.updatedAt = new Date().toISOString();
  saveContacts();
  renderAll();
  closeEditor();
}

function updateMailLinks() {
  const mail1 = String(contactForm.elements.namedItem("mail1")?.value || "").trim();
  const mail2 = String(contactForm.elements.namedItem("mail2")?.value || "").trim();
  setMailLink(mail1Link, mail1);
  setMailLink(mail2Link, mail2);
}

function setMailLink(node, email) {
  if (!node) return;
  if (!email) {
    node.classList.add("hidden");
    node.removeAttribute("href");
    return;
  }
  node.href = `mailto:${email}`;
  node.classList.remove("hidden");
}

function newContact() {
  const contact = createEmptyContact();
  state.contacts.unshift(contact);
  state.selectedId = contact.id;
  saveContacts();
  renderAll();
  openEditor();
}

function parseDateText(value) {
  const text = String(value || "").trim();
  if (!text) return null;

  let match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) {
    return buildDateParts(Number(match[1]), Number(match[2]), Number(match[3]), true);
  }

  match = text.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (match) {
    return buildDateParts(Number(match[3]), Number(match[2]), Number(match[1]), true);
  }

  match = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (match) {
    return buildDateParts(Number(match[3]), Number(match[2]), Number(match[1]), true);
  }

  match = text.match(/^(\d{1,2})[-/.](\d{1,2})$/);
  if (match) {
    return buildDateParts(TODAY.getFullYear(), Number(match[2]), Number(match[1]), false);
  }

  match = normalize(text).match(/^(\d{1,2}) de ([a-z]+) de (\d{4})$/);
  if (match) {
    const month = MONTH_MAP[match[2]];
    return buildDateParts(Number(match[3]), month, Number(match[1]), true);
  }

  match = normalize(text).match(/^(\d{1,2}) ([a-z]+) (\d{4})$/);
  if (match) {
    const month = MONTH_MAP[match[2]];
    return buildDateParts(Number(match[3]), month, Number(match[1]), true);
  }

  match = normalize(text).match(/^(\d{1,2}) de ([a-z]+)$/);
  if (match) {
    const month = MONTH_MAP[match[2]];
    return buildDateParts(TODAY.getFullYear(), month, Number(match[1]), false);
  }

  match = normalize(text).match(/^(\d{1,2}) ([a-z]+)$/);
  if (match) {
    const month = MONTH_MAP[match[2]];
    return buildDateParts(TODAY.getFullYear(), month, Number(match[1]), false);
  }

  return null;
}

function buildDateParts(year, month, day, hasYear) {
  if (!year || !month || !day) return null;
  const date = new Date(year, month - 1, day);
  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }
  return { year, month, day, date, hasYear };
}

function formatDateLong(value) {
  const parts = parseDateText(value);
  if (!parts) return String(value || "").trim();
  if (!parts.hasYear) return `${parts.day} de ${MONTH_NAMES[parts.month - 1]}`;
  return `${parts.day} de ${MONTH_NAMES[parts.month - 1]} de ${parts.year}`;
}

function yearsMonthsDaysDiff(fromDate, toDate) {
  let years = toDate.getFullYear() - fromDate.getFullYear();
  let months = toDate.getMonth() - fromDate.getMonth();
  let days = toDate.getDate() - fromDate.getDate();

  if (days < 0) {
    months -= 1;
    const previousMonth = new Date(toDate.getFullYear(), toDate.getMonth(), 0).getDate();
    days += previousMonth;
  }
  if (months < 0) {
    years -= 1;
    months += 12;
  }
  return { years, months, days };
}

function nextOccurrence(parts, today = TODAY) {
  const candidate = new Date(today.getFullYear(), parts.month - 1, parts.day);
  if (candidate < startOfDay(today)) {
    candidate.setFullYear(candidate.getFullYear() + 1);
  }
  return candidate;
}

function startOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function diffDays(from, to) {
  const msPerDay = 24 * 60 * 60 * 1000;
  return Math.round((startOfDay(to) - startOfDay(from)) / msPerDay);
}

function describeDate(value, type) {
  const parts = parseDateText(value);
  if (!parts) return "";

  const next = nextOccurrence(parts, TODAY);
  const remaining = diffDays(TODAY, next);
  const kind = String(type || "").trim() || "Data";
  const remainingText =
    remaining === 0
      ? `avui`
      : remaining === 1
        ? `dema`
        : `${remaining} dies`;

  if (!parts.hasYear) {
    return `${kind}: ${remainingText}`;
  }

  const age = yearsMonthsDaysDiff(parts.date, TODAY);
  const nextAge = next.getFullYear() - parts.year;
  return `${kind}: ${age.years} anys, ${age.months} mesos, ${age.days} dies · ${remainingText} per ${nextAge} anys`;
}

function updateDateInfos() {
  for (let index = 1; index <= 4; index += 1) {
    const dateValue = contactForm.elements.namedItem(`date${index}`)?.value || "";
    const dateType = contactForm.elements.namedItem(`date${index}Type`)?.value || "";
    const target = document.getElementById(`date${index}Info`);
    if (target) target.textContent = describeDate(dateValue, dateType);
  }
}

function collectReminders() {
  const reminders = [];
  for (const contact of state.contacts) {
    for (let index = 1; index <= 4; index += 1) {
      const raw = contact[`date${index}`];
      const parts = parseDateText(raw);
      if (!parts) continue;
      const eventDate = nextOccurrence(parts, TODAY);
      const daysAway = diffDays(TODAY, eventDate);
      if (daysAway !== 0) continue;
      reminders.push({
        id: `${contact.id}-${index}`,
        contactName: contact.name || "Sense nom",
        type: contact[`date${index}Type`] || `Data ${index}`,
        label: formatDateLong(raw),
        daysAway,
        eventDate,
        age: eventDate.getFullYear() - parts.year,
      });
    }
  }
  return reminders.sort((a, b) => a.daysAway - b.daysAway || a.contactName.localeCompare(b.contactName));
}

function renderReminders() {
  const reminders = collectReminders();
  upcomingCount.textContent = String(reminders.length);
  reminderList.innerHTML = "";

  if (!reminders.length) {
    reminderList.innerHTML = `<div class="empty-state">Avui no hi ha cap aniversari ni data important.</div>`;
    return;
  }

  const fragment = document.createDocumentFragment();
  for (const reminder of reminders) {
    const item = document.createElement("article");
    item.className = `reminder-item${reminder.daysAway === 0 ? " today" : ""}`;
    const when =
      reminder.daysAway === 0
        ? "Avui"
        : reminder.daysAway === 1
          ? "Dema"
          : `D'aqui ${reminder.daysAway} dies`;
    item.innerHTML = `
      <h3>${escapeHtml(reminder.contactName)} · ${escapeHtml(reminder.type)}</h3>
      <p>${escapeHtml(reminder.label)} · ${escapeHtml(when)}${reminder.age > 0 ? ` · ${reminder.age} anys` : ""}</p>
    `;
    fragment.appendChild(item);
  }
  reminderList.appendChild(fragment);
}

async function notifyTodayReminders() {
  if (!("Notification" in window)) return;
  if (Notification.permission !== "granted") return;

  const dayKey = formatDateForFile(TODAY);
  if (localStorage.getItem(NOTIFY_DAY_KEY) === dayKey) return;

  const todayReminders = collectReminders().filter((reminder) => reminder.daysAway === 0);
  if (!todayReminders.length) return;

  const body = todayReminders
    .slice(0, 4)
    .map((reminder) => `${reminder.contactName}: ${reminder.type}`)
    .join(" · ");

  new Notification("Avui tens dates importants", { body });
  localStorage.setItem(NOTIFY_DAY_KEY, dayKey);
}

async function requestNotifications() {
  if (!("Notification" in window)) {
    alert("Aquest navegador no admet notificacions web.");
    return;
  }
  const permission = await Notification.requestPermission();
  if (permission === "granted") {
    notificationsBtn.textContent = "Avisos activats";
    await notifyTodayReminders();
  }
}

function openGpsLocation() {
  const lat = String(contactForm.elements.namedItem("gpsLat")?.value || "").trim();
  const lng = String(contactForm.elements.namedItem("gpsLng")?.value || "").trim();
  const address = String(contactForm.elements.namedItem("address")?.value || "").trim();
  const postalCode = String(contactForm.elements.namedItem("postalCode")?.value || "").trim();
  const city = String(contactForm.elements.namedItem("city")?.value || "").trim();

  let url = "";
  if (lat && lng) {
    url = `https://www.google.com/maps?q=${encodeURIComponent(`${lat},${lng}`)}`;
  } else {
    const query = [address, postalCode, city].filter(Boolean).join(", ");
    if (!query) {
      gpsStatus.textContent = "Falten coordenades o adreca per obrir el mapa.";
      return;
    }
    url = `https://www.google.com/maps?q=${encodeURIComponent(query)}`;
  }
  window.open(url, "_blank");
}

function captureGps() {
  if (!navigator.geolocation) {
    gpsStatus.textContent = "Aquest dispositiu no admet geolocalitzacio.";
    return;
  }
  gpsStatus.textContent = "Buscant la posicio actual...";
  navigator.geolocation.getCurrentPosition(
    (position) => {
      const lat = position.coords.latitude.toFixed(6);
      const lng = position.coords.longitude.toFixed(6);
      contactForm.elements.namedItem("gpsLat").value = lat;
      contactForm.elements.namedItem("gpsLng").value = lng;
      gpsStatus.textContent = `GPS capturat: ${lat}, ${lng}`;
      saveCurrentContact();
    },
    () => {
      gpsStatus.textContent = "No he pogut obtenir la posicio GPS.";
    },
    { enableHighAccuracy: true, timeout: 15000, maximumAge: 0 }
  );
}

function formChanged() {
  updateMailLinks();
  updateDateInfos();
}

function openEditor() {
  renderSelectedContact();
  editorOverlay.classList.remove("hidden");
  editorOverlay.setAttribute("aria-hidden", "false");
}

function closeEditor() {
  editorOverlay.classList.add("hidden");
  editorOverlay.setAttribute("aria-hidden", "true");
}

function deleteCurrentContact() {
  const contact = selectedContact();
  if (!contact) return;

  const name = contact.name || "este contacto";
  const confirmDelete = window.confirm(`¿Quieres borrar "${name}"?`);
  if (!confirmDelete) return;

  state.contacts = state.contacts.filter((item) => item.id !== contact.id);
  state.selectedIds.delete(contact.id);
  state.selectedId = state.contacts[0]?.id || null;
  saveContacts();
  renderAll();
  closeEditor();
}

function selectFilteredContacts() {
  const contacts = filteredContacts();
  for (const contact of contacts) {
    state.selectedIds.add(contact.id);
  }
  renderContactList();
}

function deleteSelectedContacts() {
  if (!state.selectedIds.size) {
    alert("No hay contactos marcados.");
    return;
  }
  const total = state.selectedIds.size;
  const confirmDelete = window.confirm(`¿Quieres borrar ${total} contactos seleccionados?`);
  if (!confirmDelete) return;

  state.contacts = state.contacts.filter((contact) => !state.selectedIds.has(contact.id));
  state.selectedIds.clear();
  if (state.selectedId && !state.contacts.some((contact) => contact.id === state.selectedId)) {
    state.selectedId = state.contacts[0]?.id || null;
    if (!state.selectedId) {
      closeEditor();
    }
  }
  saveContacts();
  renderAll();
}

function registerServiceWorker() {
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("./service-worker.js").then((registration) => {
      registration.update().catch(() => {});
    }).catch(() => {});
  }
}

function findStatus() {
  if (!("Notification" in window)) return "Avisos no disponibles";
  if (Notification.permission === "granted") return "Avisos activats";
  if (Notification.permission === "denied") return "Avisos bloquejats";
  return "Activar avisos";
}

function init() {
  notificationsBtn.textContent = findStatus();
  if (!state.contacts.length) {
    state.contacts = [];
  }
  renderAll();
  registerServiceWorker();
  notifyTodayReminders();
}

searchInput.addEventListener("input", renderContactList);
saveBtn.addEventListener("click", saveCurrentContact);
newContactBtn.addEventListener("click", newContact);
importXlsxBtn.addEventListener("click", () => xlsxInput.click());
exportXlsxBtn.addEventListener("click", exportContactsXlsx);
backupBtn.addEventListener("click", exportBackupJson);
selectFilteredBtn.addEventListener("click", selectFilteredContacts);
deleteSelectedBtn.addEventListener("click", deleteSelectedContacts);
notificationsBtn.addEventListener("click", requestNotifications);
saveGpsBtn.addEventListener("click", captureGps);
openGpsBtn.addEventListener("click", openGpsLocation);
backToListBtn.addEventListener("click", () => {
  closeEditor();
});
deleteContactBtn.addEventListener("click", deleteCurrentContact);
closeEditorBtn.addEventListener("click", closeEditor);
editorBackdrop.addEventListener("click", closeEditor);
contactForm.addEventListener("input", formChanged);
contactForm.addEventListener("submit", (event) => {
  event.preventDefault();
  saveCurrentContact();
});
xlsxInput.addEventListener("change", async (event) => {
  const [file] = event.target.files || [];
  if (!file) return;
  try {
    await importContactsFromFile(file);
  } catch (error) {
    alert(`No he pogut importar el fitxer: ${error.message}`);
  } finally {
    xlsxInput.value = "";
  }
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !editorOverlay.classList.contains("hidden")) {
    closeEditor();
  }
});

init();
