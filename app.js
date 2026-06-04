const APP_CONFIG = {
  clientId: "",
  tenantId: "",
  redirectUri: `${window.location.origin}/`,
  sharepointSiteUrl: "https://lucalza.sharepoint.com/sites/prueba2",
  listName: "DescargasOperativas",
};

const defaultPilotos = [
  "Alexander Tojes",
  "Cristian Gomez",
  "Diego",
  "Abner",
  "Manuel Arias",
  "Abimael Maldonado",
];
const defaultUnitsByPilot = {
  "Manuel Arias": "TXT-22",
  "Abimael Maldonado": "TXT-17",
};
const productAliases = {
  PREMIUM: "4 PREMIUM",
};
const titles = {
  menuScreen: "Sistema de Despacho",
  cargaScreen: "Carga",
  checklistScreen: "Checklist Seguridad",
  despachoScreen: "Ingrese Despacho",
  historialScreen: "Historial Despachos",
  calificacionScreen: "Calificar Despacho",
  sharepointScreen: "Configuracion SharePoint",
};
const ratingOptions = [
  { value: "malo", face: "😟", label: "Malo" },
  { value: "regular", face: "😐", label: "Regular" },
  { value: "bueno", face: "🙂", label: "Bueno" },
  { value: "excelente", face: "😄", label: "Excelente" },
];

let state = {
  TablaDescargas: [],
  ChecklistPreSalida: [],
  checklistConfirmado: [],
};

let graphToken = "";
let graphUser = "";
let graphSiteId = "";
let graphListId = "";
let activeDispatchTripId = null;

const screens = document.querySelectorAll(".screen");
const screenTitle = document.querySelector("#screenTitle");
const toast = document.querySelector("#toast");
const loginBanner = document.createElement("div");
loginBanner.className = "helper-panel";
loginBanner.innerHTML = '<p class="helper-title">Sesion</p><p id="loginStatus" class="helper-text">Sin iniciar sesion</p>';

const cargaForm = document.querySelector("#cargaForm");
const cargaPilotoSelect = document.querySelector("#cargaPiloto");
const cargaUnidadInput = document.querySelector("#cargaUnidad");
const checklistForm = document.querySelector("#checklistForm");
const despachoForm = document.querySelector("#despachoForm");
const pilotoSelect = document.querySelector("#piloto");
const unidadInput = document.querySelector("#unidad");
const productoSelect = document.querySelector("#producto");
const galonesCargadosInput = document.querySelector("#galonesCargados");
const galonesRecibidosInput = document.querySelector("#galonesRecibidos");
const diferenciaPreview = document.querySelector("#diferenciaPreview");
const galonesPendientesOutput = document.querySelector("#galonesPendientes");
const historialList = document.querySelector("#historialList");
const calificacionList = document.querySelector("#calificacionList");
const filtroPiloto = document.querySelector("#filtroPiloto");
const filtroFecha = document.querySelector("#filtroFecha");
const filtroClienteCalificacion = document.querySelector("#filtroClienteCalificacion");
const filtroFechaCalificacion = document.querySelector("#filtroFechaCalificacion");
const sharepointForm = document.querySelector("#sharepointForm");
const storageModeSelect = document.querySelector("#storageMode");
const sharepointSiteUrlInput = document.querySelector("#sharepointSiteUrl");
const sharepointTenantIdInput = document.querySelector("#sharepointTenantId");
const sharepointClientIdInput = document.querySelector("#sharepointClientId");
const sharepointDescargasListInput = document.querySelector("#sharepointDescargasList");
const sharepointChecklistListInput = document.querySelector("#sharepointChecklistList");
const sharepointRatingsListInput = document.querySelector("#sharepointRatingsList");
const sharepointStatus = document.querySelector("#sharepointStatus");
const spTestButton = document.querySelector("#spTestButton");
const loginButton = document.createElement("button");
loginButton.type = "button";
loginButton.className = "secondary-button";
loginButton.textContent = "INICIAR SESION MICROSOFT 365";
loginButton.addEventListener("click", loginMicrosoft);

const sharepointPanel = document.querySelector("#sharepointScreen .form-card");
sharepointPanel.parentNode.insertBefore(loginBanner, sharepointPanel);
sharepointPanel.insertBefore(loginButton, sharepointPanel.firstChild);

document.querySelectorAll("[data-screen]").forEach((button) => {
  button.addEventListener("click", () => showScreen(button.dataset.screen));
});

cargaForm.addEventListener("submit", saveCarga);
cargaPilotoSelect.addEventListener("change", () => updateDefaultUnit(cargaPilotoSelect, cargaUnidadInput));
checklistForm.addEventListener("submit", saveChecklist);
despachoForm.addEventListener("submit", saveDespacho);
pilotoSelect.addEventListener("change", loadActiveTripForDispatch);
galonesCargadosInput.addEventListener("input", updateDifference);
galonesRecibidosInput.addEventListener("input", updateDifference);
filtroPiloto.addEventListener("change", renderHistory);
filtroFecha.addEventListener("change", renderHistory);
filtroClienteCalificacion.addEventListener("change", renderRatings);
filtroFechaCalificacion.addEventListener("change", renderRatings);
sharepointForm.addEventListener("submit", saveSharePointConfig);
spTestButton.addEventListener("click", testSharePointConnection);

setToday();
refreshDropdowns();
renderHistory();
updateDifference();
renderSharePointStatus();

initMicrosoftAuth().catch((error) => {
  console.warn("Microsoft 365 init failed:", error);
  updateLoginStatus("Sesion Microsoft no disponible aun.");
});

async function initMicrosoftAuth() {
  const config = getSharePointConfig();
  if (!config.clientId || !config.tenantId) {
    return;
  }

  let msal;
  try {
    msal = await ensureMsal();
  } catch (error) {
    updateLoginStatus("No se pudo cargar Microsoft 365.");
    throw error;
  }

  APP_CONFIG.clientId = config.clientId;
  APP_CONFIG.tenantId = config.tenantId;
  APP_CONFIG.redirectUri = window.location.origin + "/";

  const msalConfig = {
    auth: {
      clientId: APP_CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${APP_CONFIG.tenantId}`,
      redirectUri: APP_CONFIG.redirectUri,
    },
    cache: {
      cacheLocation: "sessionStorage",
    },
  };

  window.msalInstance = new msal.PublicClientApplication(msalConfig);
  await handleAuthRedirect();
  await hydrateSharePointSite();
  await loadAllData();
}

async function ensureMsal() {
  if (window.msal) return window.msal;
  await injectScript("https://alcdn.msauth.net/browser/2.39.0/js/msal-browser.min.js");
  if (!window.msal) {
    throw new Error("msal-browser no quedo disponible despues de cargar el script.");
  }
  return window.msal;
}

function injectScript(src) {
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.onload = resolve;
    script.onerror = () => reject(new Error(`No se pudo cargar ${src}`));
    document.head.appendChild(script);
  });
}

async function handleAuthRedirect() {
  const response = await window.msalInstance.handleRedirectPromise();
  if (response?.account) {
    window.msalInstance.setActiveAccount(response.account);
  }

  const account = window.msalInstance.getAllAccounts()[0] || null;
  if (account) {
    window.msalInstance.setActiveAccount(account);
    graphUser = account.username || account.name || "";
    updateLoginStatus(`Sesion activa: ${graphUser}`);
  }
}

async function loginMicrosoft() {
  const loginRequest = {
    scopes: ["User.Read", "Sites.ReadWrite.All", "offline_access"],
  };
  await window.msalInstance.loginRedirect(loginRequest);
}

function updateLoginStatus(text) {
  const status = document.querySelector("#loginStatus");
  status.textContent = text;
}

function persist() {
  // SharePoint es la fuente de verdad; no guardamos nada localmente.
}

function getSharePointConfig() {
  return {
    storageMode: storageModeSelect.value || "sharepoint",
    siteUrl: sharepointSiteUrlInput.value.trim() || APP_CONFIG.sharepointSiteUrl,
    tenantId: sharepointTenantIdInput.value.trim() || APP_CONFIG.tenantId,
    clientId: sharepointClientIdInput.value.trim() || APP_CONFIG.clientId,
    descargasList: sharepointDescargasListInput.value.trim() || APP_CONFIG.listName,
    checklistList: sharepointChecklistListInput.value.trim() || "ChecklistPreSalida",
    ratingsList: sharepointRatingsListInput.value.trim() || "CalificacionesDespacho",
  };
}

function saveSharePointConfig(event) {
  event.preventDefault();
  APP_CONFIG.sharepointSiteUrl = sharepointSiteUrlInput.value.trim() || APP_CONFIG.sharepointSiteUrl;
  APP_CONFIG.tenantId = sharepointTenantIdInput.value.trim() || APP_CONFIG.tenantId;
  APP_CONFIG.clientId = sharepointClientIdInput.value.trim() || APP_CONFIG.clientId;
  APP_CONFIG.listName = sharepointDescargasListInput.value.trim() || APP_CONFIG.listName;
  renderSharePointStatus();
  notify("Configuracion aplicada en esta sesion.");
}

function renderSharePointStatus() {
  const config = getSharePointConfig();
  if (config.storageMode !== "sharepoint") {
    sharepointStatus.textContent = "Modo local desactivado. La app esta preparada para usar SharePoint.";
    return;
  }

  if (!config.siteUrl || !config.tenantId || !config.clientId) {
    sharepointStatus.textContent = "Completa Site URL, Tenant ID y Client ID para iniciar Microsoft 365.";
    return;
  }

  sharepointStatus.textContent = `Conectando con ${config.siteUrl} y la lista ${config.descargasList}.`;
}

function showScreen(screenId) {
  screens.forEach((screen) => screen.classList.toggle("active", screen.id === screenId));
  screenTitle.textContent = titles[screenId] || titles.menuScreen;

  if (screenId === "historialScreen") renderHistory();
  if (screenId === "calificacionScreen") renderRatings();
  if (screenId === "checklistScreen") refreshDropdowns();

  window.scrollTo({ top: 0, behavior: "smooth" });
}

async function testSharePointConnection() {
  const config = getSharePointConfig();
  if (config.storageMode !== "sharepoint") {
    notify("Activa el modo SharePoint para probar la conexion.");
    return;
  }

  if (!config.siteUrl || !config.tenantId || !config.clientId) {
    notify("Completa Site URL, Tenant ID y Client ID.");
    return;
  }

  try {
    await initMicrosoftAuth();
    notify("SharePoint listo para conectar.");
  } catch (error) {
    console.error(error);
    notify("No se pudo conectar con Microsoft 365.");
  }
}

async function getGraphToken() {
  const account = window.msalInstance.getActiveAccount() || window.msalInstance.getAllAccounts()[0];
  if (!account) {
    await loginMicrosoft();
    return "";
  }

  try {
    const response = await window.msalInstance.acquireTokenSilent({
      account,
      scopes: ["User.Read", "Sites.ReadWrite.All", "offline_access"],
    });
    graphToken = response.accessToken;
    return graphToken;
  } catch {
    await window.msalInstance.acquireTokenRedirect({
      account,
      scopes: ["User.Read", "Sites.ReadWrite.All", "offline_access"],
    });
    return "";
  }
}

async function hydrateSharePointSite() {
  const config = getSharePointConfig();
  if (!config.siteUrl) return;
  const token = await getGraphToken();
  if (!token) return;

  const siteUrl = new URL(config.siteUrl);
  const hostname = siteUrl.hostname;
  const path = siteUrl.pathname.replace(/\/$/, "");
  const siteResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${hostname}:${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!siteResponse.ok) {
    throw new Error("No se pudo resolver el sitio de SharePoint.");
  }

  const site = await siteResponse.json();
  graphSiteId = site.id;

  const listsResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${graphSiteId}/lists`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const lists = listsResponse.ok ? await listsResponse.json() : { value: [] };
  const match = lists.value.find((item) => item.displayName === config.descargasList);
  graphListId = match ? match.id : "";
  renderSharePointStatus();
}

async function loadAllData() {
  const config = getSharePointConfig();
  if (config.storageMode !== "sharepoint" || !graphSiteId || !graphListId) return;
  const token = await getGraphToken();
  if (!token) return;

  const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${graphSiteId}/lists/${graphListId}/items?expand=fields`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!response.ok) {
    throw new Error("No se pudo leer la lista de SharePoint.");
  }

  const payload = await response.json();
  state.TablaDescargas = (payload.value || []).map((item) => mapSharePointDownload(item));
  state.ChecklistPreSalida = [];
  state.checklistConfirmado = [];
  normalizeLoadedState({ TablaDescargas: state.TablaDescargas, ChecklistPreSalida: state.ChecklistPreSalida });
  renderHistory();
  renderRatings();
  refreshFilterOptions();
  refreshRatingFilterOptions();
}

function mapSharePointDownload(item) {
  const fields = item.fields || {};
  return {
    id: item.id,
    Cliente: fields.Cliente || "",
    Piloto: fields.Piloto || "",
    Unidad: fields.Unidad || "",
    Producto: fields.Producto || "",
    FechaDescarga: fields.FechaDescarga || null,
    HoraInicio: fields.HoraInicio || "",
    HoraFin: fields.HoraFin || "",
    GalonesCargados: Number(fields.GalonesCargados || 0),
    GalonesRecibidos: Number(fields.GalonesRecibidos || 0),
    Diferencia: Number(fields.Diferencia || 0),
    Estado: fields.Estado || "",
    Operador: fields.Operador || "",
    Observaciones: fields.Observaciones || "",
    FechaRegistro: fields.FechaRegistro || "",
    Calificacion: fields.Calificacion || "",
  };
}

async function saveCarga(event) {
  event.preventDefault();
  const record = {
    FechaCarga: valueOf("#cargaFecha"),
    Estacion: valueOf("#cargaPlanta"),
    Piloto: valueOf("#cargaPiloto"),
    Unidad: valueOf("#cargaUnidad"),
    Producto: valueOf("#cargaProducto"),
    OdometroInicial: optionalNumberOf("#odometroInicial"),
    GalonesCargados: numberOf("#cargaGalones"),
    OdometroFinal: optionalNumberOf("#odometroFinal"),
    Estado: "EN RUTA",
    GalonesRestantes: numberOf("#cargaGalones"),
  };
  state.ChecklistPreSalida.push(record);
  await saveToSharePoint(record);
  cargaForm.reset();
  setToday();
  refreshDropdowns();
  notify("Carga guardada.");
  showScreen("menuScreen");
}

async function saveChecklist(event) {
  event.preventDefault();
  state.checklistConfirmado.push({ tripId: valueOf("#checklistViaje"), confirmadoEn: new Date().toISOString() });
  persist();
  checklistForm.reset();
  notify("Checklist guardado.");
  showScreen("menuScreen");
}

async function saveDespacho(event) {
  event.preventDefault();
  const activeTrip = findTripById(activeDispatchTripId);
  const galonesCargados = activeTrip ? Number(activeTrip.GalonesRestantes || 0) : numberOf("#galonesCargados");
  const galonesRecibidos = numberOf("#galonesRecibidos");
  const diferencia = Math.max(galonesCargados - galonesRecibidos, 0);
  const config = getSharePointConfig();
  const item = {
    FechaDescarga: valueOf("#fechaDescarga"),
    HoraInicio: valueOf("#horaInicio"),
    HoraFin: valueOf("#horaFin"),
    Estacion: activeTrip?.Estacion || valueOf("#cargaPlanta") || "",
    Piloto: valueOf("#piloto"),
    Unidad: valueOf("#unidad"),
    Producto: valueOf("#producto"),
    GalonesCargados: galonesCargados,
    GalonesRecibidos: galonesRecibidos,
    Diferencia: diferencia,
    Estado: diferencia === 0 ? "FINALIZADO" : "EN RUTA",
    Operador: graphUser || "",
    Observaciones: valueOf("#comentarioDespacho"),
    FechaRegistro: new Date().toISOString(),
  };

  state.TablaDescargas.push(item);
  updateTripStatus(item);
  await saveToSharePoint(item);
  despachoForm.reset();
  activeDispatchTripId = null;
  setToday();
  updateDifference();
  updatePendingIndicator(0);
  refreshDropdowns();
  notify("Despacho guardado en SharePoint.");
  showScreen("historialScreen");
}

async function saveToSharePoint(item) {
  const config = getSharePointConfig();
  if (config.storageMode !== "sharepoint") return;

  const token = await getGraphToken();
  if (!token) return;
  if (!graphSiteId || !graphListId) await hydrateSharePointSite();

  const fields = {
    FechaDescarga: item.FechaDescarga,
    HoraInicio: item.HoraInicio,
    HoraFin: item.HoraFin,
    Estacion: item.Estacion || item.Planta || "",
    Piloto: item.Piloto,
    Unidad: item.Unidad,
    Producto: item.Producto,
    GalonesCargados: item.GalonesCargados,
    GalonesRecibidos: item.GalonesRecibidos,
    Diferencia: item.Diferencia,
    Estado: item.Estado || "",
    Operador: item.Operador || graphUser || "",
    Observaciones: item.Observaciones || "",
    FechaRegistro: item.FechaRegistro || new Date().toISOString(),
    Cliente: item.Cliente || "",
    Calificacion: item.Calificacion || "",
    ViajeCargaId: item.ViajeCargaId || "",
    OdometroInicial: item.OdometroInicial ?? null,
    OdometroFinal: item.OdometroFinal ?? null,
    GalonesRestantes: item.GalonesRestantes ?? null,
    EstadoViaje: item.EstadoViaje || item.Estado || "",
    Planta: item.Planta || item.Estacion || "",
  };

  const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${graphSiteId}/lists/${graphListId}/items`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ fields }),
  });
  if (!response.ok) {
    throw new Error("No se pudo guardar en SharePoint.");
  }
}

function updateTripStatus(descarga) {
  const trip = state.ChecklistPreSalida.find((item) => {
    return item.Piloto === descarga.Piloto &&
      item.Unidad.trim().toUpperCase() === descarga.Unidad.trim().toUpperCase() &&
      item.Producto.trim().toUpperCase() === descarga.Producto.trim().toUpperCase() &&
      item.Estado !== "FINALIZADO";
  });
  if (!trip) return;
  const pendientesAntes = Number(trip.GalonesRestantes || trip.GalonesCargados || 0);
  trip.GalonesRestantes = Math.max(pendientesAntes - descarga.GalonesRecibidos, 0);
  trip.Estado = trip.GalonesRestantes === 0 ? "FINALIZADO" : "EN RUTA";
}

function renderHistory() {
  refreshFilterOptions();
  const selectedPilot = filtroPiloto.value;
  const selectedDate = filtroFecha.value;
  const rows = state.TablaDescargas
    .filter((row) => !selectedPilot || row.Piloto === selectedPilot)
    .filter((row) => !selectedDate || String(row.FechaDescarga).slice(0, 10) === selectedDate)
    .sort((a, b) => new Date(b.FechaDescarga) - new Date(a.FechaDescarga));
  historialList.innerHTML = "";
  if (!rows.length) {
    historialList.innerHTML = '<div class="empty-state">No hay despachos para los filtros seleccionados.</div>';
    return;
  }
  rows.forEach((row) => {
    const status = visualStatus(Number(row.Diferencia || 0));
    const card = document.createElement("article");
    card.className = `dispatch-card ${status.className}`;
    card.innerHTML = `
      <h2 class="card-title">${escapeHtml(row.Cliente || row.Piloto || "Despacho")}</h2>
      <p class="card-line"><strong>Fecha:</strong> ${formatDate(row.FechaDescarga)}</p>
      <p class="card-line"><strong>Piloto:</strong> ${escapeHtml(row.Piloto)}</p>
      <p class="card-line"><strong>Unidad:</strong> ${escapeHtml(row.Unidad)}</p>
      <p class="card-line"><strong>Producto:</strong> ${escapeHtml(row.Producto)}</p>
      <p class="card-line"><strong>Galones cargados:</strong> ${formatNumber(row.GalonesCargados)}</p>
      <p class="status-text">${status.label}</p>
      <p class="card-date">${formatDate(row.FechaRegistro || row.FechaDescarga)}</p>
    `;
    historialList.appendChild(card);
  });
}

function renderRatings() {
  refreshRatingFilterOptions();
  const selectedClient = filtroClienteCalificacion.value;
  const selectedDate = filtroFechaCalificacion.value;
  const rows = state.TablaDescargas
    .filter((row) => !selectedClient || row.Cliente === selectedClient)
    .filter((row) => !selectedDate || String(row.FechaDescarga).slice(0, 10) === selectedDate)
    .sort((a, b) => new Date(b.FechaDescarga) - new Date(a.FechaDescarga));
  calificacionList.innerHTML = "";
  if (!rows.length) {
    calificacionList.innerHTML = '<div class="empty-state">No hay despachos para los filtros seleccionados.</div>';
    return;
  }
  rows.forEach((row) => {
    const status = visualStatus(Number(row.Diferencia || 0));
    const card = document.createElement("article");
    card.className = "rating-card";
    card.innerHTML = `
      <div>
        <h2 class="card-title">${escapeHtml(row.Cliente || "Despacho")}</h2>
        <p class="card-line"><strong>Fecha:</strong> ${formatDate(row.FechaDescarga)}</p>
        <p class="card-line"><strong>Piloto:</strong> ${escapeHtml(row.Piloto)}</p>
        <p class="card-line"><strong>Unidad:</strong> ${escapeHtml(row.Unidad)}</p>
        <p class="card-line"><strong>Producto:</strong> ${escapeHtml(row.Producto)}</p>
        <p class="card-line"><strong>Entregado:</strong> ${formatNumber(row.GalonesRecibidos)} gal</p>
        <p class="status-text">${status.label}</p>
      </div>
      <div class="rating-actions">
        ${ratingOptions.map((rating) => `
          <button class="face-button ${row.Calificacion === rating.value ? "selected" : ""}" type="button" data-dispatch-id="${escapeHtml(row.id)}" data-rating="${rating.value}">
            <span class="face-icon">${rating.face}</span>
            <span>${rating.label}</span>
          </button>
        `).join("")}
      </div>
    `;
    calificacionList.appendChild(card);
  });
}

calificacionList.addEventListener("click", async (event) => {
  const button = event.target.closest("[data-dispatch-id][data-rating]");
  if (!button) return;
  const despacho = state.TablaDescargas.find((row) => row.id === button.dataset.dispatchId);
  if (!despacho) return;
  despacho.Calificacion = button.dataset.rating;
  await saveToSharePoint(despacho);
  renderRatings();
  notify("Calificacion guardada.");
});

function refreshDropdowns() {
  const pilotos = unique([
    ...defaultPilotos,
    ...state.ChecklistPreSalida.map((row) => row.Piloto),
    ...state.TablaDescargas.map((row) => row.Piloto),
  ]);
  fillSelect("#piloto", pilotos, "Seleccione piloto");
  fillSelect("#cargaPiloto", pilotos, "Seleccione piloto");

  if (pilotos.length) {
    if (!cargaPilotoSelect.value) {
      cargaPilotoSelect.value = pilotos[0];
    }
    if (!pilotoSelect.value) {
      pilotoSelect.value = pilotos[0];
    }
  }

  updateDefaultUnit(cargaPilotoSelect, cargaUnidadInput);
  loadActiveTripForDispatch();

  const activeTrips = state.ChecklistPreSalida.filter((row) => row.Estado !== "FINALIZADO");
  const tripSelect = document.querySelector("#checklistViaje");
  tripSelect.innerHTML = '<option value="">Seleccione viaje</option>';
  activeTrips.forEach((trip) => {
    const option = document.createElement("option");
    option.value = trip.id;
    option.textContent = `${trip.Piloto} | ${trip.Unidad} | ${trip.Producto} | ${formatNumber(trip.GalonesRestantes)} gal`;
    tripSelect.appendChild(option);
  });
}

function refreshFilterOptions() {
  const current = filtroPiloto.value;
  const pilotos = unique(state.TablaDescargas.map((row) => row.Piloto));
  filtroPiloto.innerHTML = '<option value="">Todos los pilotos</option>';
  pilotos.forEach((piloto) => {
    const option = document.createElement("option");
    option.value = piloto;
    option.textContent = piloto;
    filtroPiloto.appendChild(option);
  });
  filtroPiloto.value = pilotos.includes(current) ? current : "";
}

function refreshRatingFilterOptions() {
  const current = filtroClienteCalificacion.value;
  const clientes = unique(state.TablaDescargas.map((row) => row.Cliente));
  filtroClienteCalificacion.innerHTML = '<option value="">Todos los clientes</option>';
  clientes.forEach((cliente) => {
    const option = document.createElement("option");
    option.value = cliente;
    option.textContent = cliente;
    filtroClienteCalificacion.appendChild(option);
  });
  filtroClienteCalificacion.value = clientes.includes(current) ? current : "";
}

function fillSelect(selector, values, placeholder) {
  const select = document.querySelector(selector);
  const current = select.value;
  select.innerHTML = `<option value="">${placeholder}</option>`;
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
  select.value = values.includes(current) ? current : (values[0] || "");
}

function updateDefaultUnit(pilotSelect, unitInput) {
  const selectedPilot = pilotSelect.value;
  const defaultUnit = defaultUnitsByPilot[selectedPilot];
  if (defaultUnit) {
    unitInput.value = defaultUnit;
    unitInput.readOnly = true;
  } else {
    unitInput.value = "";
    unitInput.readOnly = false;
  }
}

function loadActiveTripForDispatch() {
  const selectedPilot = pilotoSelect.value;
  const activeTrip = getActiveTripForPilot(selectedPilot);
  if (!activeTrip) {
    activeDispatchTripId = null;
    updateDefaultUnit(pilotoSelect, unidadInput);
    productoSelect.value = "";
    galonesCargadosInput.value = "";
    updatePendingIndicator(0);
    updateDifference();
    return;
  }
  activeDispatchTripId = activeTrip.id;
  unidadInput.value = activeTrip.Unidad || "";
  unidadInput.readOnly = Boolean(defaultUnitsByPilot[selectedPilot]);
  productoSelect.value = activeTrip.Producto || "";
  galonesCargadosInput.value = Number(activeTrip.GalonesRestantes || activeTrip.GalonesCargados || 0);
  updatePendingIndicator(Number(activeTrip.GalonesRestantes || 0));
  updateDifference();
}

function getActiveTripForPilot(pilot) {
  if (!pilot) return null;
  return state.ChecklistPreSalida
    .filter((trip) => trip.Piloto === pilot && trip.Estado !== "FINALIZADO" && Number(trip.GalonesRestantes || 0) > 0)
    .sort((a, b) => new Date(a.FechaCarga || 0) - new Date(b.FechaCarga || 0))[0] || null;
}

function findTripById(id) {
  if (!id) return null;
  return state.ChecklistPreSalida.find((trip) => trip.id === id) || null;
}

function updatePendingIndicator(value) {
  galonesPendientesOutput.textContent = `GALONES PENDIENTES: ${formatNumber(value)}`;
}

function updateDifference() {
  const cargados = Number(galonesCargadosInput.value || 0);
  const recibidos = Number(galonesRecibidosInput.value || 0);
  diferenciaPreview.textContent = `Diferencia: ${formatNumber(Math.max(cargados - recibidos, 0))}`;
}

function visualStatus(difference) {
  if (difference === 0) return { className: "finalizado", label: "FINALIZADO" };
  if (difference <= 499) return { className: "pendiente-leve", label: `PENDIENTE GALONES: ${formatNumber(difference)}` };
  return { className: "pendiente-alto", label: `PENDIENTE GALONES: ${formatNumber(difference)}` };
}

function valueOf(selector) {
  return document.querySelector(selector).value.trim();
}

function numberOf(selector) {
  return Number(document.querySelector(selector).value || 0);
}

function optionalNumberOf(selector) {
  const value = document.querySelector(selector).value;
  return value === "" ? null : Number(value);
}

function toIsoDateTime(dateValue) {
  const [year, month, day] = dateValue.split("-").map(Number);
  return new Date(Date.UTC(year, month - 1, day)).toISOString();
}

function formatDate(isoValue) {
  return new Intl.DateTimeFormat("es-GT", {
    day: "numeric",
    month: "long",
    year: "numeric",
    timeZone: "UTC",
  }).format(new Date(isoValue));
}

function formatNumber(value) {
  return new Intl.NumberFormat("es-GT", { maximumFractionDigits: 2 }).format(value);
}

function shortId(id) {
  return String(id || "").slice(0, 8).toUpperCase();
}

function unique(values) {
  return [...new Set(values.filter(Boolean).map((value) => value.trim()))].sort((a, b) => a.localeCompare(b));
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function notify(message) {
  toast.textContent = message;
  toast.classList.add("show");
  window.setTimeout(() => toast.classList.remove("show"), 2400);
}
