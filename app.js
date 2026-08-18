const STORAGE_KEY = "lluis-contactes-nova-app-v2";
const NOTIFY_DAY_KEY = "lluis-contactes-last-notify-day";
const DATE_KEYS = ["date1", "date2", "date3", "date4"];
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
  "notes",
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
  notes: "Notes",
  gpsLat: "Latitud",
  gpsLng: "Longitud",
};
const IMPORT_ALIASES = {
  name: ["Nom", "Nombre", "Nombre formateado"],
  group: ["Grup", "Grupo", "Empresa"],
  phoneWork: ["Telf. Treball", "Telefono trabajo", "Telf. Trabajo"],
  phoneHome: ["Telf. Casa", "Telefono casa"],
  phoneMobile: ["Telf. Mobil", "Telefono", "Teléfono", "Telf. 1"],
  phoneMobileType: ["Tipus telf. Mobil", "Tipo telefono movil", "Tipo teléfono móvil", "Tipo telf. 1"],
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
  notes: ["Notes", "Notas", "Observacions", "Observaciones"],
  gpsLat: ["Latitud", "GPS Lat"],
  gpsLng: ["Longitud", "GPS Lng"],
};
const MONTH_NAMES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
const MONTH_MAP = {
  enero: 1, febrero: 2, marzo: 3, abril: 4, mayo: 5, junio: 6, julio: 7, agosto: 8, septiembre: 9, setiembre: 9,
  octubre: 10, noviembre: 11, diciembre: 12, gener: 1, febrer: 2, marc: 3, maig: 5, juny: 6, juliol: 7, agost: 8,
  setembre: 9, novembre: 11, desembre: 12,
};
const now = new Date();
const TODAY = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 12, 0, 0);
var desktopAPI = window.desktopAPI || null;
var isDesktopApp = Boolean(desktopAPI?.isDesktopApp);

const state = {
  contacts: [],
  selectedId: null,
  selectedIds: new Set(),
  activeLetter: "",
  editorDirty: false,
  editorIsNew: false,
  editorOriginalSnapshot: "",
  returnToDuplicates: false,
};

const contactList = document.getElementById("contactList");
const alphabetNav = document.getElementById("alphabetNav");
const reminderList = document.getElementById("reminderList");
const contactCount = document.getElementById("contactCount");
const upcomingCount = document.getElementById("upcomingCount");
const searchInput = document.getElementById("searchInput");
const clearSearchBtn = document.getElementById("clearSearchBtn");
const contactForm = document.getElementById("contactForm");
const editorTitle = document.getElementById("editorTitle");
const saveBtn = document.getElementById("saveBtn");
const newContactBtn = document.getElementById("newContactBtn");
const newContactTopBtn = document.getElementById("newContactTopBtn");
const toolsMenuBtn = document.getElementById("toolsMenuBtn");
const toolsMenuPanel = document.getElementById("toolsMenuPanel");
const importXlsxBtn = document.getElementById("importXlsxBtn");
const importQuickBtn = document.getElementById("importQuickBtn");
const importJsonBtn = document.getElementById("importJsonBtn");
const importJsonQuickBtn = document.getElementById("importJsonQuickBtn");
const exportXlsxBtn = document.getElementById("exportXlsxBtn");
const exportQuickBtn = document.getElementById("exportQuickBtn");
const backupBtn = document.getElementById("backupBtn");
const backupQuickBtn = document.getElementById("backupQuickBtn");
const selectFilteredBtn = document.getElementById("selectFilteredBtn");
const selectQuickBtn = document.getElementById("selectQuickBtn");
const deleteSelectedBtn = document.getElementById("deleteSelectedBtn");
const deleteQuickBtn = document.getElementById("deleteQuickBtn");
const openDatesCalendarBtn = document.getElementById("openDatesCalendarBtn");
const openDatesCalendarQuickBtn = document.getElementById("openDatesCalendarQuickBtn");
const openDuplicatesBtn = document.getElementById("openDuplicatesBtn");
const openDuplicatesQuickBtn = document.getElementById("openDuplicatesQuickBtn");
const xlsxInput = document.getElementById("xlsxInput");
const jsonInput = document.getElementById("jsonInput");
const saveGpsBtn = document.getElementById("saveGpsBtn");
const openGpsBtn = document.getElementById("openGpsBtn");
const gpsStatus = document.getElementById("gpsStatus");
const mail1Link = document.getElementById("mail1Link");
const mail2Link = document.getElementById("mail2Link");
const phoneWorkLink = document.getElementById("phoneWorkLink");
const phoneHomeLink = document.getElementById("phoneHomeLink");
const phoneMobileLink = document.getElementById("phoneMobileLink");
const uploadPhotoBtn = document.getElementById("uploadPhotoBtn");
const deletePhotoBtn = document.getElementById("deletePhotoBtn");
const photoInput = document.getElementById("photoInput");
const contactPhotoPreview = document.getElementById("contactPhotoPreview");
const contactPhotoImage = document.getElementById("contactPhotoImage");
const contactPhotoInitials = document.getElementById("contactPhotoInitials");
const photoZoomInput = document.getElementById("photoZoomInput");
const photoPositionXInput = document.getElementById("photoPositionXInput");
const photoPositionYInput = document.getElementById("photoPositionYInput");
const editorOverlay = document.getElementById("editorOverlay");
const editorBackdrop = document.getElementById("editorBackdrop");
const backToListBtn = document.getElementById("backToListBtn");
const closeEditorBtn = document.getElementById("closeEditorBtn");
const deleteContactBtn = document.getElementById("deleteContactBtn");
const datesCalendarOverlay = document.getElementById("datesCalendarOverlay");
const datesCalendarBackdrop = document.getElementById("datesCalendarBackdrop");
const closeDatesCalendarBtn = document.getElementById("closeDatesCalendarBtn");
const datesCalendarContent = document.getElementById("datesCalendarContent");
const duplicatesOverlay = document.getElementById("duplicatesOverlay");
const duplicatesBackdrop = document.getElementById("duplicatesBackdrop");
const closeDuplicatesBtn = document.getElementById("closeDuplicatesBtn");
const duplicatesContent = document.getElementById("duplicatesContent");

let photoDragState = null;

async function loadContacts() {
  if (isDesktopApp) {
    try {
      const parsed = await desktopAPI.loadContacts();
      return Array.isArray(parsed) ? parsed : [];
    } catch (error) {
      console.error("No he pogut llegir els contactes locals del PC.", error);
    }
  }
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

async function saveContacts() {
  if (isDesktopApp) {
    try {
      await desktopAPI.saveContacts(state.contacts);
      return;
    } catch (error) {
      console.error("No he pogut guardar els contactes locals del PC.", error);
      alert("No he pogut guardar els contactes a la carpeta privada del PC.");
      return;
    }
  }
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.contacts));
}

function normalize(text) {
  return String(text || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
}

function escapeHtml(value) {
  return String(value || "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");
}

function safeFileName(name) {
  return String(name || "contactes").replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim();
}

function createEmptyContact() {
  const contact = {
    id: crypto.randomUUID(),
    createdAt: new Date().toISOString(),
    photoData: "",
    photoPath: "",
    photoZoom: 100,
    photoPositionX: 50,
    photoPositionY: 50,
  };
  for (const key of FIELD_ORDER) contact[key] = "";
  return contact;
}

function sanitizeImportedJsonContact(rawContact) {
  const contact = createEmptyContact();
  const source = rawContact && typeof rawContact === "object" ? rawContact : {};
  contact.id = String(source.id || crypto.randomUUID());
  contact.createdAt = String(source.createdAt || new Date().toISOString());
  contact.photoData = String(source.photoData || "");
  contact.photoPath = String(source.photoPath || "");
  contact.photoZoom = getPhotoZoomValue(source.photoZoom);
  contact.photoPositionX = getPhotoPositionValue(source.photoPositionX);
  contact.photoPositionY = getPhotoPositionValue(source.photoPositionY);
  for (const key of FIELD_ORDER) {
    const value = source[key] ?? "";
    contact[key] = DATE_KEYS.includes(key) ? formatDateLong(value) : String(value).trim();
  }
  return contact;
}

function selectedContact() {
  return state.contacts.find((contact) => contact.id === state.selectedId) || null;
}

function snapshotContact(contact) {
  const base = {};
  for (const key of FIELD_ORDER) base[key] = String(contact?.[key] || "");
  base.photoData = String(contact?.photoData || "");
  base.photoPath = String(contact?.photoPath || "");
  base.photoZoom = Number.isFinite(Number(contact?.photoZoom)) ? Number(contact.photoZoom) : 100;
  base.photoPositionX = Number.isFinite(Number(contact?.photoPositionX)) ? Number(contact.photoPositionX) : 50;
  base.photoPositionY = Number.isFinite(Number(contact?.photoPositionY)) ? Number(contact.photoPositionY) : 50;
  return JSON.stringify(base);
}

function formSnapshot() {
  const contact = selectedContact();
  const draft = {};
  for (const key of FIELD_ORDER) {
    const raw = contactForm.elements.namedItem(key)?.value || "";
    draft[key] = DATE_KEYS.includes(key) ? formatDateLong(raw) : String(raw).trim();
  }
  draft.photoData = String(contact?.photoData || "");
  draft.photoPath = String(contact?.photoPath || "");
  draft.photoZoom = Number(photoZoomInput?.value || 100);
  draft.photoPositionX = Number(photoPositionXInput?.value || 50);
  draft.photoPositionY = Number(photoPositionYInput?.value || 50);
  return JSON.stringify(draft);
}

function markEditorClean(contact) {
  state.editorOriginalSnapshot = snapshotContact(contact);
  state.editorDirty = false;
}

function syncEditorDirtyState() {
  if (editorOverlay.classList.contains("hidden")) {
    state.editorDirty = false;
    return false;
  }
  state.editorDirty = formSnapshot() !== state.editorOriginalSnapshot;
  return state.editorDirty;
}

function contactInitials(name) {
  const words = String(name || "").trim().split(/\s+/).filter(Boolean);
  if (!words.length) return "?";
  return words.slice(0, 2).map((word) => word[0]).join("").toUpperCase();
}

function getPhotoPositionValue(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return 50;
  return Math.max(0, Math.min(100, numeric));
}

function getPhotoZoomValue(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return 100;
  return Math.max(100, Math.min(220, numeric));
}

function buildPhotoObjectPosition(positionX = 50, positionY = 50) {
  const safeX = getPhotoPositionValue(positionX);
  const safeY = getPhotoPositionValue(positionY);
  return `${safeX}% ${safeY}%`;
}

function buildPhotoTransform(zoom = 100) {
  const safeZoom = getPhotoZoomValue(zoom);
  return `scale(${safeZoom / 100})`;
}

function buildEditorPhotoTransform(positionX = 50, positionY = 50, zoom = 100) {
  const safeX = getPhotoPositionValue(positionX);
  const safeY = getPhotoPositionValue(positionY);
  const safeZoom = getPhotoZoomValue(zoom);
  const moveX = ((safeX - 50) / 50) * 56;
  const moveY = ((safeY - 50) / 50) * 56;
  return `translate(${moveX}px, ${moveY}px) scale(${safeZoom / 100})`;
}

function buildListPhotoTransform(positionX = 50, positionY = 50, zoom = 100) {
  const safeX = getPhotoPositionValue(positionX);
  const safeY = getPhotoPositionValue(positionY);
  const safeZoom = Math.min(getPhotoZoomValue(zoom), 132);
  const moveX = ((safeX - 50) / 50) * 18;
  const moveY = ((safeY - 50) / 50) * 18;
  return `translate(${moveX}px, ${moveY}px) scale(${safeZoom / 100})`;
}

function applyPhotoPosition(positionX = 50, positionY = 50, zoom = 100) {
  const safeX = getPhotoPositionValue(positionX);
  const safeY = getPhotoPositionValue(positionY);
  const safeZoom = getPhotoZoomValue(zoom);
  if (photoZoomInput) photoZoomInput.value = String(safeZoom);
  if (photoPositionXInput) photoPositionXInput.value = String(safeX);
  if (photoPositionYInput) photoPositionYInput.value = String(safeY);
  contactPhotoImage.style.objectPosition = "50% 50%";
  contactPhotoImage.style.transform = buildEditorPhotoTransform(safeX, safeY, safeZoom);
}

function renderPhotoPreview(photoData, name = "", photoZoom = 100, photoPositionX = 50, photoPositionY = 50) {
  applyPhotoPosition(photoPositionX, photoPositionY, photoZoom);
  if (photoData) {
    contactPhotoImage.src = photoData;
    contactPhotoImage.classList.remove("hidden");
    contactPhotoInitials.classList.add("hidden");
  } else {
    contactPhotoImage.removeAttribute("src");
    contactPhotoImage.classList.add("hidden");
    contactPhotoInitials.textContent = contactInitials(name);
    contactPhotoInitials.classList.remove("hidden");
  }
}

function getPointerPoint(event) {
  if ("touches" in event && event.touches?.length) {
    return event.touches[0];
  }
  if ("changedTouches" in event && event.changedTouches?.length) {
    return event.changedTouches[0];
  }
  return event;
}

function beginPhotoDrag(event) {
  const contact = selectedContact();
  if (!contact || !contact.photoData) return;
  const point = getPointerPoint(event);
  photoDragState = {
    startX: point.clientX,
    startY: point.clientY,
    baseX: getPhotoPositionValue(photoPositionXInput?.value || contact.photoPositionX || 50),
    baseY: getPhotoPositionValue(photoPositionYInput?.value || contact.photoPositionY || 50),
  };
  contactPhotoPreview?.classList.add("is-dragging");
  event.preventDefault?.();
}

function updatePhotoDrag(event) {
  if (!photoDragState) return;
  const point = getPointerPoint(event);
  const deltaX = point.clientX - photoDragState.startX;
  const deltaY = point.clientY - photoDragState.startY;
  const nextX = getPhotoPositionValue(photoDragState.baseX + deltaX * 0.45);
  const nextY = getPhotoPositionValue(photoDragState.baseY + deltaY * 0.45);
  if (photoPositionXInput) photoPositionXInput.value = String(nextX);
  if (photoPositionYInput) photoPositionYInput.value = String(nextY);
  refreshPhotoPositionPreview();
  event.preventDefault?.();
}

function endPhotoDrag() {
  photoDragState = null;
  contactPhotoPreview?.classList.remove("is-dragging");
}

function handlePhotoWheel(event) {
  const contact = selectedContact();
  if (!contact || !contact.photoData) return;
  event.preventDefault();
  const currentZoom = getPhotoZoomValue(photoZoomInput?.value || contact.photoZoom || 100);
  const nextZoom = getPhotoZoomValue(currentZoom + (event.deltaY < 0 ? 8 : -8));
  if (photoZoomInput) photoZoomInput.value = String(nextZoom);
  refreshPhotoPositionPreview();
}

function openEditor() {
  try {
    renderSelectedContact();
  } catch (error) {
    console.error("No he pogut preparar la fitxa del contacte.", error);
    editorTitle.textContent = "Nou contacte";
    contactForm.reset();
    gpsStatus.textContent = "";
    renderPhotoPreview("", "", 100, 50, 50);
  }
  editorOverlay.classList.remove("hidden");
  editorOverlay.setAttribute("aria-hidden", "false");
  closeToolsMenu();
}

function restoreEditorOriginalState() {
  const contact = selectedContact();
  if (!contact || !state.editorOriginalSnapshot) return;
  const original = JSON.parse(state.editorOriginalSnapshot);
  for (const key of FIELD_ORDER) contact[key] = String(original[key] || "");
  contact.photoData = String(original.photoData || "");
  contact.photoPath = String(original.photoPath || "");
  contact.photoZoom = getPhotoZoomValue(original.photoZoom);
  contact.photoPositionX = getPhotoPositionValue(original.photoPositionX);
  contact.photoPositionY = getPhotoPositionValue(original.photoPositionY);
}

function discardEditorChanges() {
  const contact = selectedContact();
  if (state.editorIsNew) {
    if (contact) {
      state.contacts = state.contacts.filter((item) => item.id !== contact.id);
      state.selectedIds.delete(contact.id);
    }
    state.selectedId = state.contacts[0]?.id || null;
  } else {
    restoreEditorOriginalState();
  }
  state.editorDirty = false;
  state.editorIsNew = false;
}

async function shouldCloseEditor() {
  const hasUnsaved = syncEditorDirtyState();
  if (!hasUnsaved) {
    if (state.editorIsNew) discardEditorChanges();
    return true;
  }
  if (window.confirm("Has fet canvis. Vols guardar-los abans de sortir?")) {
    await saveCurrentContact(false);
    return true;
  }
  return window.confirm("Vols sortir sense guardar els canvis?");
}

async function closeEditor(force = false) {
  const reopenDuplicates = state.returnToDuplicates;
  if (!force && !(await shouldCloseEditor())) return;
  editorOverlay.classList.add("hidden");
  editorOverlay.setAttribute("aria-hidden", "true");
  if (force) {
    state.editorDirty = false;
    state.editorIsNew = false;
    state.editorOriginalSnapshot = "";
  } else {
    discardEditorChanges();
  }
  state.returnToDuplicates = false;
  renderAll();
  if (reopenDuplicates) openDuplicates();
}

function ensureSelectedContact() {
  if (!state.contacts.length) {
    state.selectedId = null;
    return;
  }
  if (!selectedContact()) state.selectedId = state.contacts[0].id;
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
      renderAll();
    });
    fragment.appendChild(button);
  }
  alphabetNav.appendChild(fragment);
}

function filteredContacts() {
  const query = normalize(searchInput.value);
  return state.contacts.filter((contact) => {
    if (state.activeLetter) {
      const first = normalize(contact.name).charAt(0).toUpperCase();
      if (first !== state.activeLetter) return false;
    }
    if (!query) return true;
    const haystack = normalize([contact.name, contact.group, contact.phoneMobile, contact.phoneWork, contact.phoneHome, contact.mail1, contact.mail2, contact.city, contact.notes].join(" "));
    return haystack.includes(query);
  }).sort((a, b) => {
    const nameA = normalize(a.name || "");
    const nameB = normalize(b.name || "");
    if (!nameA && !nameB) return 0;
    if (!nameA) return 1;
    if (!nameB) return -1;
    return nameA.localeCompare(nameB, "ca");
  });
}

function clearSearch() {
  searchInput.value = "";
  state.activeLetter = "";
  renderAll();
  searchInput.focus();
}

function renderContactList() {
  const contacts = filteredContacts();
  contactList.innerHTML = "";
  contactCount.textContent = String(state.contacts.length);
  deleteSelectedBtn.textContent = state.selectedIds.size ? `Borrar seleccionados (${state.selectedIds.size})` : "Borrar seleccionados";

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
      if (checkbox.checked) state.selectedIds.add(contact.id);
      else state.selectedIds.delete(contact.id);
      renderContactList();
    });

    const button = document.createElement("button");
    button.type = "button";
    button.className = `contact-row${contact.id === state.selectedId ? " active" : ""}`;
    button.innerHTML = `
      <div class="contact-avatar">
        ${contact.photoData ? `<img src="${escapeHtml(contact.photoData)}" alt="${escapeHtml(contact.name || "Foto")}" style="object-position: 50% 50%; transform: ${buildListPhotoTransform(contact.photoPositionX, contact.photoPositionY, contact.photoZoom)};">` : `<span>${escapeHtml(contactInitials(contact.name))}</span>`}
      </div>
      <div>
        <p class="contact-name">${escapeHtml(contact.name || "Sense nom")}</p>
        <div class="contact-meta">${escapeHtml(contact.group || "Sense grup")} · ${escapeHtml(contact.phoneMobile || contact.phoneWork || contact.mail1 || "")}</div>
      </div>
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
    renderPhotoPreview("", "", 100, 50, 50);
    updateMailLinks();
    updateDateInfos();
    return;
  }

  editorTitle.textContent = contact.name || "Contacte sense nom";
  for (const key of FIELD_ORDER) {
    const input = contactForm.elements.namedItem(key);
    if (input) input.value = contact[key] || "";
  }
  gpsStatus.textContent = contact.gpsLat && contact.gpsLng ? `GPS guardat: ${contact.gpsLat}, ${contact.gpsLng}` : "";
  renderPhotoPreview(contact.photoData || "", contact.name || "", contact.photoZoom, contact.photoPositionX, contact.photoPositionY);
  updateMailLinks();
  updateDateInfos();
  markEditorClean(contact);
}

function setMailLink(node, email) {
  if (!node) return;
  if (!email) {
    node.classList.add("hidden");
    node.removeAttribute("href");
    node.onclick = null;
    return;
  }
  node.href = `mailto:${email}`;
  node.onclick = () => {
    window.location.href = node.href;
  };
  node.classList.remove("hidden");
}

function normalizePhoneNumber(phone) {
  return String(phone || "").trim().replace(/[^\d+]/g, "");
}

function setPhoneLink(node, phone) {
  if (!node) return;
  const rawPhone = String(phone || "").trim();
  const cleanPhone = normalizePhoneNumber(rawPhone);
  if (!cleanPhone) {
    node.classList.add("hidden");
    node.removeAttribute("href");
    node.onclick = null;
    return;
  }
  node.href = `tel:${cleanPhone}`;
  node.onclick = () => {
    window.location.href = node.href;
  };
  node.classList.remove("hidden");
}

function updateMailLinks() {
  setMailLink(mail1Link, String(contactForm.elements.namedItem("mail1")?.value || "").trim());
  setMailLink(mail2Link, String(contactForm.elements.namedItem("mail2")?.value || "").trim());
  setPhoneLink(phoneWorkLink, contactForm.elements.namedItem("phoneWork")?.value || "");
  setPhoneLink(phoneHomeLink, contactForm.elements.namedItem("phoneHome")?.value || "");
  setPhoneLink(phoneMobileLink, contactForm.elements.namedItem("phoneMobile")?.value || "");
}

async function saveCurrentContact(closeAfterSave = true) {
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
  contact.photoData = contact.photoData || "";
  contact.photoZoom = getPhotoZoomValue(photoZoomInput?.value || contact.photoZoom || 100);
  contact.photoPositionX = getPhotoPositionValue(photoPositionXInput?.value || contact.photoPositionX || 50);
  contact.photoPositionY = getPhotoPositionValue(photoPositionYInput?.value || contact.photoPositionY || 50);
  contact.updatedAt = new Date().toISOString();
  await saveContacts();
  state.editorIsNew = false;
  markEditorClean(contact);
  renderAll();
  if (closeAfterSave) await closeEditor(true);
}

function newContact() {
  try {
    const contact = createEmptyContact();
    state.contacts.unshift(contact);
    state.selectedId = contact.id;
    state.editorIsNew = true;
    renderAll();
    openEditor();
  } catch (error) {
    console.error("No he pogut crear un nou contacte.", error);
    alert(`No he pogut obrir el nou registre: ${error.message}`);
    editorOverlay.classList.remove("hidden");
    editorOverlay.setAttribute("aria-hidden", "false");
  }
}

function deleteCurrentContact() {
  const contact = selectedContact();
  if (!contact) return;
  if (!window.confirm(`¿Quieres borrar "${contact.name || "este contacto"}"?`)) return;
  state.contacts = state.contacts.filter((item) => item.id !== contact.id);
  state.selectedIds.delete(contact.id);
  state.selectedId = state.contacts[0]?.id || null;
  void saveContacts();
  renderAll();
  closeEditor();
}

function selectFilteredContacts() {
  for (const contact of filteredContacts()) state.selectedIds.add(contact.id);
  renderContactList();
}

function deleteSelectedContacts() {
  if (!state.selectedIds.size) {
    alert("No hay contactos marcados.");
    return;
  }
  if (!window.confirm(`¿Quieres borrar ${state.selectedIds.size} contactos seleccionados?`)) return;
  state.contacts = state.contacts.filter((contact) => !state.selectedIds.has(contact.id));
  state.selectedIds.clear();
  if (state.selectedId && !state.contacts.some((contact) => contact.id === state.selectedId)) state.selectedId = state.contacts[0]?.id || null;
  void saveContacts();
  renderAll();
  if (!state.selectedId) closeEditor();
}

function uploadPhoto() {
  if (!selectedContact()) return;
  photoInput.click();
}

function handlePhotoSelected(event) {
  const [file] = event.target.files || [];
  if (!file) return;
  const contact = selectedContact();
  if (!contact) return;
  const reader = new FileReader();
  reader.onload = async () => {
    const loadedData = String(reader.result || "");
    try {
      if (isDesktopApp && desktopAPI?.savePhotoDataUrl) {
        const savedPhoto = await desktopAPI.savePhotoDataUrl(loadedData, file.name || contact.name || "foto");
        contact.photoData = savedPhoto.photoUrl;
        contact.photoPath = savedPhoto.photoPath;
      } else {
        contact.photoData = loadedData;
        contact.photoPath = "";
      }
      contact.photoZoom = 100;
      contact.photoPositionX = 50;
      contact.photoPositionY = 50;
      renderPhotoPreview(contact.photoData || "", contactForm.elements.namedItem("name")?.value || "", contact.photoZoom, contact.photoPositionX, contact.photoPositionY);
      renderContactList();
      syncEditorDirtyState();
    } catch (error) {
      alert(`No he pogut guardar la foto: ${error.message}`);
    }
  };
  reader.readAsDataURL(file);
  photoInput.value = "";
}

function deleteCurrentPhoto() {
  const contact = selectedContact();
  if (!contact || !contact.photoData) {
    return;
  }
  if (!window.confirm("Vols esborrar aquesta foto del contacte?")) {
    return;
  }
  contact.photoData = "";
  contact.photoPath = "";
  contact.photoZoom = 100;
  contact.photoPositionX = 50;
  contact.photoPositionY = 50;
  renderPhotoPreview("", contactForm.elements.namedItem("name")?.value || "", 100, 50, 50);
  renderContactList();
  syncEditorDirtyState();
}

function buildDateParts(year, month, day, hasYear) {
  if (!year || !month || !day) return null;
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
  return { year, month, day, date, hasYear };
}

function parseDateText(value) {
  const text = String(value || "").trim();
  if (!text) return null;
  let match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) return buildDateParts(Number(match[1]), Number(match[2]), Number(match[3]), true);
  match = text.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (match) return buildDateParts(Number(match[3]), Number(match[2]), Number(match[1]), true);
  match = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (match) return buildDateParts(Number(match[3]), Number(match[2]), Number(match[1]), true);
  match = text.match(/^(\d{1,2})[-/.](\d{1,2})$/);
  if (match) return buildDateParts(TODAY.getFullYear(), Number(match[2]), Number(match[1]), false);
  match = normalize(text).match(/^(\d{1,2}) de ([a-z]+) de (\d{4})$/);
  if (match) return buildDateParts(Number(match[3]), MONTH_MAP[match[2]], Number(match[1]), true);
  match = normalize(text).match(/^(\d{1,2}) ([a-z]+) (\d{4})$/);
  if (match) return buildDateParts(Number(match[3]), MONTH_MAP[match[2]], Number(match[1]), true);
  match = normalize(text).match(/^(\d{1,2}) de ([a-z]+)$/);
  if (match) return buildDateParts(TODAY.getFullYear(), MONTH_MAP[match[2]], Number(match[1]), false);
  match = normalize(text).match(/^(\d{1,2}) ([a-z]+)$/);
  if (match) return buildDateParts(TODAY.getFullYear(), MONTH_MAP[match[2]], Number(match[1]), false);
  return null;
}

function formatDateLong(value) {
  const parts = parseDateText(value);
  if (!parts) return String(value || "").trim();
  return parts.hasYear ? `${parts.day} de ${MONTH_NAMES[parts.month - 1]} de ${parts.year}` : `${parts.day} de ${MONTH_NAMES[parts.month - 1]}`;
}

function yearsMonthsDaysDiff(fromDate, toDate) {
  let years = toDate.getFullYear() - fromDate.getFullYear();
  let months = toDate.getMonth() - fromDate.getMonth();
  let days = toDate.getDate() - fromDate.getDate();
  if (days < 0) {
    months -= 1;
    days += new Date(toDate.getFullYear(), toDate.getMonth(), 0).getDate();
  }
  if (months < 0) {
    years -= 1;
    months += 12;
  }
  return { years, months, days };
}

function startOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function nextOccurrence(parts) {
  const candidate = new Date(TODAY.getFullYear(), parts.month - 1, parts.day);
  if (candidate < startOfDay(TODAY)) candidate.setFullYear(candidate.getFullYear() + 1);
  return candidate;
}

function diffDays(from, to) {
  return Math.round((startOfDay(to) - startOfDay(from)) / (24 * 60 * 60 * 1000));
}

function describeDate(value, type) {
  const parts = parseDateText(value);
  if (!parts) return "";
  const kind = String(type || "").trim() || "Data";
  const remaining = diffDays(TODAY, nextOccurrence(parts));
  const remainingText = remaining === 0 ? "avui" : remaining === 1 ? "dema" : `${remaining} dies`;
  if (!parts.hasYear) return `${kind}: ${remainingText}`;
  const age = yearsMonthsDaysDiff(parts.date, TODAY);
  const nextAge = nextOccurrence(parts).getFullYear() - parts.year;
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
      const parts = parseDateText(contact[`date${index}`]);
      if (!parts) continue;
      const daysAway = diffDays(TODAY, nextOccurrence(parts));
      if (daysAway < 0 || daysAway > 15) continue;
      reminders.push({
        id: `${contact.id}-${index}`,
        contactName: contact.name || "Sense nom",
        type: contact[`date${index}Type`] || `Data ${index}`,
        label: formatDateLong(contact[`date${index}`]),
        age: parts.hasYear ? nextOccurrence(parts).getFullYear() - parts.year : 0,
        daysAway,
      });
    }
  }
  return reminders.sort((a, b) => {
    if (a.daysAway !== b.daysAway) return a.daysAway - b.daysAway;
    return a.contactName.localeCompare(b.contactName, "ca");
  });
}

function renderReminders() {
  const reminders = collectReminders();
  upcomingCount.textContent = String(reminders.length);
  reminderList.innerHTML = "";
  if (!reminders.length) {
    reminderList.innerHTML = `<div class="empty-state">No hi ha cap data important per avui ni per als proxims 15 dies.</div>`;
    return;
  }
  const fragment = document.createDocumentFragment();
  for (const reminder of reminders) {
    const item = document.createElement("article");
    const whenText = reminder.daysAway === 0
      ? "Avui"
      : reminder.daysAway === 1
        ? "Dema"
        : `Falten ${reminder.daysAway} dies`;
    item.className = `reminder-item${reminder.daysAway === 0 ? " today" : " upcoming"}`;
    item.innerHTML = `
      <h3>${escapeHtml(reminder.contactName)} · ${escapeHtml(reminder.type)}</h3>
      <p>${escapeHtml(reminder.label)}${reminder.age > 0 ? ` · ${reminder.age} anys` : ""}</p>
      <span class="reminder-badge">${escapeHtml(whenText)}</span>
    `;
    fragment.appendChild(item);
  }
  reminderList.appendChild(fragment);
}

function collectAllDatesByMonth() {
  const grouped = new Map();
  for (const contact of state.contacts) {
    for (let index = 1; index <= 4; index += 1) {
      const rawDate = contact[`date${index}`];
      const parts = parseDateText(rawDate);
      if (!parts) continue;
      const month = parts.month;
      if (!grouped.has(month)) grouped.set(month, []);
      grouped.get(month).push({
        contactName: contact.name || "Sense nom",
        type: contact[`date${index}Type`] || `Data ${index}`,
        label: formatDateLong(rawDate),
        detail: describeDate(rawDate, ""),
        day: parts.day,
        month,
      });
    }
  }
  const currentMonth = TODAY.getMonth() + 1;
  return [...grouped.entries()]
    .sort((a, b) => {
      const offsetA = (a[0] - currentMonth + 12) % 12;
      const offsetB = (b[0] - currentMonth + 12) % 12;
      return offsetA - offsetB;
    })
    .map(([month, items]) => ({
      month,
      monthName: MONTH_NAMES[month - 1],
      items: items.sort((a, b) => {
        if (a.day !== b.day) return a.day - b.day;
        return a.contactName.localeCompare(b.contactName, "ca");
      }),
    }));
}

function dateTypeClass(type) {
  const normalized = normalize(type || "");
  if (normalized.includes("anivers")) return "anniversary";
  if (normalized.includes("santoral")) return "saint";
  if (normalized.includes("noces")) return "wedding";
  if (normalized.includes("traspas")) return "memorial";
  return "other";
}

function renderDatesCalendar() {
  const groups = collectAllDatesByMonth();
  datesCalendarContent.innerHTML = "";
  if (!groups.length) {
    datesCalendarContent.innerHTML = `<div class="empty-state">Encara no hi ha dates guardades per mostrar en el calendari.</div>`;
    return;
  }
  const fragment = document.createDocumentFragment();
  groups.forEach((group, index) => {
    const card = document.createElement("details");
    card.className = "calendar-month-card";
    if (index === 0) card.open = true;
    const listHtml = group.items.map((item) => `
      <article class="calendar-date-row">
        <strong>${escapeHtml(item.contactName)}</strong>
        <span>${escapeHtml(item.label)}</span>
        <span class="calendar-type-badge ${dateTypeClass(item.type)}"><strong>${escapeHtml(item.type)}</strong>${item.detail ? `<span>${escapeHtml(item.detail)}</span>` : ""}</span>
      </article>
    `).join("");
    card.innerHTML = `
      <summary class="calendar-month-summary">
        <h3>${escapeHtml(group.monthName)}</h3>
        <span class="calendar-month-count">${group.items.length}</span>
      </summary>
      <div class="calendar-month-list">${listHtml}</div>
    `;
    fragment.appendChild(card);
  });
  datesCalendarContent.appendChild(fragment);
}

function contactDisplayMeta(contact) {
  return escapeHtml(contact.group || contact.phoneMobile || contact.phoneWork || contact.mail1 || "Sense detall");
}

function collectDuplicateGroups() {
  const names = new Map();
  const phones = new Map();
  for (const contact of state.contacts) {
    const normalizedName = normalize(contact.name);
    if (normalizedName) {
      if (!names.has(normalizedName)) names.set(normalizedName, []);
      names.get(normalizedName).push(contact);
    }
    for (const key of ["phoneMobile", "phoneWork", "phoneHome"]) {
      const cleanPhone = normalizePhoneNumber(contact[key]);
      if (!cleanPhone) continue;
      if (!phones.has(cleanPhone)) phones.set(cleanPhone, []);
      phones.get(cleanPhone).push({ contact, source: key, raw: String(contact[key] || "").trim() });
    }
  }
  const duplicateNames = [...names.entries()]
    .filter(([, items]) => items.length > 1)
    .sort((a, b) => a[0].localeCompare(b[0], "ca"));
  const duplicatePhones = [...phones.entries()]
    .filter(([, items]) => new Set(items.map((item) => item.contact.id)).size > 1)
    .sort((a, b) => a[0].localeCompare(b[0], "ca"));
  return { duplicateNames, duplicatePhones };
}

function renderDuplicates() {
  const { duplicateNames, duplicatePhones } = collectDuplicateGroups();
  duplicatesContent.innerHTML = "";
  if (!duplicateNames.length && !duplicatePhones.length) {
    duplicatesContent.innerHTML = `<div class="empty-state">No he trobat noms ni telefons repetits.</div>`;
    return;
  }
  const blocks = [];
  if (duplicateNames.length) {
    const cards = duplicateNames.map(([name, items]) => `
      <article class="duplicate-card">
        <h3>Nom repetit: ${escapeHtml(items[0].name || name)}</h3>
        <div class="duplicate-list">
          ${items.map((contact) => `
            <button type="button" class="duplicate-row" data-contact-id="${escapeHtml(contact.id)}">
              <strong>${escapeHtml(contact.name || "Sense nom")}</strong>
              <span>${contactDisplayMeta(contact)}</span>
            </button>
          `).join("")}
        </div>
      </article>
    `).join("");
    blocks.push(`<section class="duplicate-section"><h2>Noms repetits</h2>${cards}</section>`);
  }
  if (duplicatePhones.length) {
    const cards = duplicatePhones.map(([phone, items]) => `
      <article class="duplicate-card">
        <h3>Telefon repetit: ${escapeHtml(items[0].raw || phone)}</h3>
        <div class="duplicate-list">
          ${items.map(({ contact, source, raw }) => `
            <button type="button" class="duplicate-row" data-contact-id="${escapeHtml(contact.id)}">
              <strong>${escapeHtml(contact.name || "Sense nom")}</strong>
              <span>${escapeHtml(FIELD_LABELS[source] || source)} · ${escapeHtml(raw || phone)} · ${contactDisplayMeta(contact)}</span>
            </button>
          `).join("")}
        </div>
      </article>
    `).join("");
    blocks.push(`<section class="duplicate-section"><h2>Telefons repetits</h2>${cards}</section>`);
  }
  duplicatesContent.innerHTML = blocks.join("");
  duplicatesContent.querySelectorAll("[data-contact-id]").forEach((node) => {
    node.addEventListener("click", () => {
      state.selectedId = node.getAttribute("data-contact-id");
      state.returnToDuplicates = true;
      closeDuplicates();
      openEditor();
    });
  });
}

function openDatesCalendar() {
  renderDatesCalendar();
  datesCalendarOverlay.classList.remove("hidden");
  datesCalendarOverlay.setAttribute("aria-hidden", "false");
}

function closeDatesCalendar() {
  datesCalendarOverlay.classList.add("hidden");
  datesCalendarOverlay.setAttribute("aria-hidden", "true");
}

function openDuplicates() {
  renderDuplicates();
  duplicatesOverlay.classList.remove("hidden");
  duplicatesOverlay.setAttribute("aria-hidden", "false");
  closeToolsMenu();
}

function closeDuplicates() {
  duplicatesOverlay.classList.add("hidden");
  duplicatesOverlay.setAttribute("aria-hidden", "true");
}

async function notifyTodayReminders() {
  if (!("Notification" in window) || Notification.permission !== "granted") return;
  const dayKey = formatDateForFile(TODAY);
  if (localStorage.getItem(NOTIFY_DAY_KEY) === dayKey) return;
  const todayReminders = collectReminders();
  if (!todayReminders.length) return;
  new Notification("Avui tens dates importants", {
    body: todayReminders.slice(0, 4).map((item) => `${item.contactName}: ${item.type}`).join(" · "),
  });
  localStorage.setItem(NOTIFY_DAY_KEY, dayKey);
}

function openGpsLocation() {
  const lat = String(contactForm.elements.namedItem("gpsLat")?.value || "").trim();
  const lng = String(contactForm.elements.namedItem("gpsLng")?.value || "").trim();
  const address = String(contactForm.elements.namedItem("address")?.value || "").trim();
  const postalCode = String(contactForm.elements.namedItem("postalCode")?.value || "").trim();
  const city = String(contactForm.elements.namedItem("city")?.value || "").trim();
  const url = lat && lng
    ? `https://www.google.com/maps?q=${encodeURIComponent(`${lat},${lng}`)}`
    : `https://www.google.com/maps?q=${encodeURIComponent([address, postalCode, city].filter(Boolean).join(", "))}`;
  if (!lat && !lng && !address && !postalCode && !city) {
    gpsStatus.textContent = "Falten coordenades o adreca per obrir el mapa.";
    return;
  }
  window.open(url, "_blank");
}

function captureGps() {
  if (!navigator.geolocation) {
    gpsStatus.textContent = "Aquest dispositiu no admet geolocalitzacio.";
    return;
  }
  gpsStatus.textContent = "Buscant la posicio actual...";
  navigator.geolocation.getCurrentPosition((position) => {
    contactForm.elements.namedItem("gpsLat").value = position.coords.latitude.toFixed(6);
    contactForm.elements.namedItem("gpsLng").value = position.coords.longitude.toFixed(6);
    gpsStatus.textContent = `GPS capturat: ${position.coords.latitude.toFixed(6)}, ${position.coords.longitude.toFixed(6)}`;
    updateMailLinks();
  }, () => {
    gpsStatus.textContent = "No he pogut obtenir la posicio GPS.";
  }, { enableHighAccuracy: true, timeout: 15000, maximumAge: 0 });
}

function formChanged() {
  updateMailLinks();
  updateDateInfos();
  const contact = selectedContact();
  if (contact) {
    renderPhotoPreview(
      contact.photoData || "",
      contactForm.elements.namedItem("name")?.value || "",
      photoZoomInput?.value || contact.photoZoom,
      photoPositionXInput?.value || contact.photoPositionX,
      photoPositionYInput?.value || contact.photoPositionY
    );
  }
  syncEditorDirtyState();
}

function formatDateForFile(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function exportBackupJson() {
  const payload = JSON.stringify(state.contacts, null, 2);
  if (isDesktopApp && desktopAPI?.saveBackupJson) {
    desktopAPI.saveBackupJson(payload).then((targetPath) => {
      alert(`Copia de seguretat guardada a:\n${targetPath}`);
    }).catch((error) => {
      alert(`No he pogut guardar la copia de seguretat: ${error.message}`);
    });
    return;
  }
  const blob = new Blob([payload], { type: "application/json;charset=utf-8" });
  downloadBlob(blob, `contactes-backup-${safeFileName(formatDateForFile(new Date()))}.json`);
}

function textCell(value) {
  return `<c t="inlineStr"><is><t xml:space="preserve">${escapeHtml(String(value ?? ""))}</t></is></c>`;
}

function buildSheetXml(rows) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows.map((row, index) => `<row r="${index + 1}">${row.map((value) => textCell(value)).join("")}</row>`).join("")}</sheetData></worksheet>`;
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

async function exportContactsXlsx() {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);
  zip.folder("_rels").file(".rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.folder("xl").file("workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Contactes" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.folder("xl").folder("_rels").file("workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
  const rows = [FIELD_ORDER.map((key) => FIELD_LABELS[key]), ...state.contacts.map((contact) => FIELD_ORDER.map((key) => contact[key] || ""))];
  zip.folder("xl").folder("worksheets").file("sheet1.xml", buildSheetXml(rows));
  const blob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  downloadBlob(blob, `contactes-${safeFileName(formatDateForFile(new Date()))}.xlsx`);
}

function excelSerialToDateText(serial) {
  const origin = new Date(Date.UTC(1899, 11, 30));
  origin.setUTCDate(origin.getUTCDate() + Number(serial));
  return `${String(origin.getUTCDate()).padStart(2, "0")}/${String(origin.getUTCMonth() + 1).padStart(2, "0")}/${origin.getUTCFullYear()}`;
}

function parseSharedStrings(xmlText) {
  const xml = new DOMParser().parseFromString(xmlText, "application/xml");
  return [...xml.getElementsByTagName("si")].map((node) => [...node.getElementsByTagName("t")].map((part) => part.textContent || "").join(""));
}

function readCellValue(cell, sharedStrings) {
  const type = cell.getAttribute("t");
  if (type === "inlineStr") return [...cell.getElementsByTagName("t")].map((node) => node.textContent || "").join("");
  const raw = cell.getElementsByTagName("v")[0]?.textContent || "";
  if (type === "s") return sharedStrings[Number(raw)] || "";
  return raw;
}

function columnLettersToIndex(letters) {
  let total = 0;
  for (const char of letters) total = total * 26 + (char.charCodeAt(0) - 64);
  return Math.max(0, total - 1);
}

function parseWorksheetXml(xmlText, sharedStrings) {
  const xml = new DOMParser().parseFromString(xmlText, "application/xml");
  return [...xml.getElementsByTagName("row")].map((row) => {
    const result = [];
    for (const cell of [...row.getElementsByTagName("c")]) {
      const index = columnLettersToIndex((cell.getAttribute("r") || "").replace(/[0-9]/g, ""));
      while (result.length < index) result.push("");
      result.push(readCellValue(cell, sharedStrings));
    }
    return result;
  });
}

function parseCombinedGps(value) {
  const match = String(value || "").trim().match(/(-?\d+(?:[.,]\d+)?)\s*[,; ]\s*(-?\d+(?:[.,]\d+)?)/);
  if (!match) return { lat: "", lng: "" };
  return { lat: match[1].replace(",", "."), lng: match[2].replace(",", ".") };
}

function findImportedValue(headers, row, key) {
  const aliases = IMPORT_ALIASES[key] || [FIELD_LABELS[key]];
  for (let index = 0; index < headers.length; index += 1) {
    if (aliases.some((alias) => normalize(alias) === normalize(headers[index]))) return row[index] ?? "";
  }
  if (key === "gpsLat" || key === "gpsLng") {
    const combinedIndex = headers.findIndex((header) => normalize(header) === normalize("Ubicacio GPS"));
    if (combinedIndex !== -1) {
      const parsed = parseCombinedGps(row[combinedIndex] || "");
      return key === "gpsLat" ? parsed.lat : parsed.lng;
    }
  }
  return "";
}

function normalizeImportedValue(key, value) {
  const text = String(value ?? "").trim();
  if (!text) return "";
  if (DATE_KEYS.includes(key) && /^[0-9]+(?:\.[0-9]+)?$/.test(text)) return formatDateLong(excelSerialToDateText(text));
  if (DATE_KEYS.includes(key)) return formatDateLong(text);
  return text;
}

function mapImportedRow(headers, row) {
  const contact = createEmptyContact();
  for (const key of FIELD_ORDER) contact[key] = normalizeImportedValue(key, findImportedValue(headers, row, key));
  return contact;
}

async function importContactsFromFile(file) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const sharedStrings = zip.file("xl/sharedStrings.xml") ? parseSharedStrings(await zip.file("xl/sharedStrings.xml").async("string")) : [];
  const workbook = new DOMParser().parseFromString(await zip.file("xl/workbook.xml").async("string"), "application/xml");
  const firstSheet = workbook.getElementsByTagName("sheet")[0];
  if (!firstSheet) throw new Error("No he trobat cap full dins del XLSX.");
  const relId = firstSheet.getAttribute("r:id");
  const relsDoc = new DOMParser().parseFromString(await zip.file("xl/_rels/workbook.xml.rels").async("string"), "application/xml");
  const relation = [...relsDoc.getElementsByTagName("Relationship")].find((rel) => rel.getAttribute("Id") === relId);
  if (!relation) throw new Error("No he pogut localitzar el full principal del XLSX.");
  const target = relation.getAttribute("Target");
  const sheetPath = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
  const rows = parseWorksheetXml(await zip.file(sheetPath).async("string"), sharedStrings).filter((row) => row.some((value) => String(value || "").trim()));
  if (rows.length < 2) {
    state.contacts = [];
    state.selectedId = null;
    await saveContacts();
    renderAll();
    return;
  }
  const headers = rows[0].map((value) => String(value || "").trim());
  state.contacts = rows.slice(1).map((row) => mapImportedRow(headers, row)).filter((contact) => contact.name);
  state.selectedId = state.contacts[0]?.id || null;
  state.selectedIds.clear();
  await saveContacts();
  renderAll();
}

async function importContactsFromJsonFile(file) {
  const text = await file.text();
  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch {
    throw new Error("El fitxer JSON no te un format valid.");
  }
  if (!Array.isArray(parsed)) {
    throw new Error("La copia JSON ha de contenir una llista de contactes.");
  }
  if (state.contacts.length) {
    const confirmed = window.confirm("Aquesta copia reemplaçara els contactes actuals. Vols continuar?");
    if (!confirmed) return;
  }
  state.contacts = parsed.map(sanitizeImportedJsonContact).filter((contact) => {
    return FIELD_ORDER.some((key) => String(contact[key] || "").trim()) || contact.photoData;
  });
  state.selectedId = state.contacts[0]?.id || null;
  state.selectedIds.clear();
  await saveContacts();
  renderAll();
  alert(`He importat ${state.contacts.length} contactes des de la copia JSON.`);
}

function registerServiceWorker() {
  if (!isDesktopApp && "serviceWorker" in navigator) {
    navigator.serviceWorker.register("./service-worker.js").then((registration) => registration.update().catch(() => {})).catch(() => {});
  }
}

function closeToolsMenu() {
  if (!toolsMenuPanel || !toolsMenuBtn) return;
  toolsMenuPanel.classList.add("hidden");
  toolsMenuBtn.classList.remove("is-open");
  toolsMenuBtn.setAttribute("aria-expanded", "false");
}

function toggleToolsMenu() {
  if (!toolsMenuPanel || !toolsMenuBtn) return;
  const isHidden = toolsMenuPanel.classList.contains("hidden");
  if (isHidden) {
    toolsMenuPanel.classList.remove("hidden");
    toolsMenuBtn.classList.add("is-open");
    toolsMenuBtn.setAttribute("aria-expanded", "true");
    return;
  }
  closeToolsMenu();
}

function renderAll() {
  ensureSelectedContact();
  renderAlphabetNav();
  renderContactList();
  renderReminders();
  if (!editorOverlay.classList.contains("hidden")) renderSelectedContact();
}

async function init() {
  state.contacts = await loadContacts();
  renderAll();
  registerServiceWorker();
  notifyTodayReminders();
}

searchInput.addEventListener("input", renderAll);
clearSearchBtn.addEventListener("click", clearSearch);
saveBtn.addEventListener("click", () => {
  void saveCurrentContact();
});
newContactBtn.addEventListener("click", newContact);
newContactTopBtn?.addEventListener("click", () => newContactBtn.click());
openDatesCalendarBtn?.addEventListener("click", openDatesCalendar);
openDatesCalendarQuickBtn?.addEventListener("click", openDatesCalendar);
openDuplicatesBtn?.addEventListener("click", openDuplicates);
openDuplicatesQuickBtn?.addEventListener("click", openDuplicates);
toolsMenuBtn?.addEventListener("click", (event) => {
  event.stopPropagation();
  toggleToolsMenu();
});
toolsMenuPanel?.addEventListener("click", (event) => event.stopPropagation());
importXlsxBtn.addEventListener("click", () => xlsxInput.click());
importQuickBtn?.addEventListener("click", () => importXlsxBtn.click());
importJsonBtn?.addEventListener("click", () => jsonInput.click());
importJsonQuickBtn?.addEventListener("click", () => importJsonBtn.click());
exportXlsxBtn.addEventListener("click", exportContactsXlsx);
exportQuickBtn?.addEventListener("click", () => exportXlsxBtn.click());
backupBtn.addEventListener("click", exportBackupJson);
backupQuickBtn?.addEventListener("click", () => backupBtn.click());
selectFilteredBtn.addEventListener("click", selectFilteredContacts);
selectQuickBtn?.addEventListener("click", () => selectFilteredBtn.click());
deleteSelectedBtn.addEventListener("click", deleteSelectedContacts);
deleteQuickBtn?.addEventListener("click", () => deleteSelectedBtn.click());
saveGpsBtn.addEventListener("click", captureGps);
openGpsBtn.addEventListener("click", openGpsLocation);
uploadPhotoBtn.addEventListener("click", uploadPhoto);
deletePhotoBtn?.addEventListener("click", deleteCurrentPhoto);
photoInput.addEventListener("change", handlePhotoSelected);
function refreshPhotoPositionPreview() {
  const contact = selectedContact();
  renderPhotoPreview(
    contact?.photoData || "",
    contactForm.elements.namedItem("name")?.value || "",
      photoZoomInput?.value || contact?.photoZoom || 100,
      photoPositionXInput?.value || contact?.photoPositionX || 50,
      photoPositionYInput?.value || contact?.photoPositionY || 50
    );
  syncEditorDirtyState();
}

photoZoomInput?.addEventListener("input", refreshPhotoPositionPreview);
photoPositionXInput?.addEventListener("input", refreshPhotoPositionPreview);
photoPositionYInput?.addEventListener("input", refreshPhotoPositionPreview);
contactPhotoPreview?.addEventListener("mousedown", beginPhotoDrag);
contactPhotoPreview?.addEventListener("touchstart", beginPhotoDrag, { passive: false });
contactPhotoPreview?.addEventListener("wheel", handlePhotoWheel, { passive: false });
window.addEventListener("mousemove", updatePhotoDrag);
window.addEventListener("touchmove", updatePhotoDrag, { passive: false });
window.addEventListener("mouseup", endPhotoDrag);
window.addEventListener("touchend", endPhotoDrag);
window.addEventListener("touchcancel", endPhotoDrag);
backToListBtn.addEventListener("click", () => {
  void closeEditor();
});
closeEditorBtn.addEventListener("click", () => {
  void closeEditor();
});
deleteContactBtn.addEventListener("click", deleteCurrentContact);
editorBackdrop.addEventListener("click", () => {
  void closeEditor();
});
closeDatesCalendarBtn?.addEventListener("click", closeDatesCalendar);
datesCalendarBackdrop?.addEventListener("click", closeDatesCalendar);
closeDuplicatesBtn?.addEventListener("click", closeDuplicates);
duplicatesBackdrop?.addEventListener("click", closeDuplicates);
contactForm.addEventListener("input", formChanged);
contactForm.addEventListener("submit", (event) => {
  event.preventDefault();
  void saveCurrentContact();
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
jsonInput?.addEventListener("change", async (event) => {
  const [file] = event.target.files || [];
  if (!file) return;
  try {
    await importContactsFromJsonFile(file);
  } catch (error) {
    alert(`No he pogut importar la copia JSON: ${error.message}`);
  } finally {
    jsonInput.value = "";
  }
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !editorOverlay.classList.contains("hidden")) closeEditor();
  if (event.key === "Escape" && !datesCalendarOverlay.classList.contains("hidden")) closeDatesCalendar();
  if (event.key === "Escape" && !duplicatesOverlay.classList.contains("hidden")) closeDuplicates();
  if (event.key === "Escape") closeToolsMenu();
});
document.addEventListener("click", () => closeToolsMenu());
window.addEventListener("beforeunload", (event) => {
  if (!editorOverlay.classList.contains("hidden") && syncEditorDirtyState()) {
    event.preventDefault();
    event.returnValue = "";
  }
});

init().catch((error) => {
  console.error(error);
  alert(`No he pogut iniciar l'app: ${error.message}`);
});
