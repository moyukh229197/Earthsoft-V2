const r3 = (v) => (Number.isFinite(v) ? Number(v).toFixed(3) : "0.000");
function formatVolume(v) {
  if (!Number.isFinite(v)) return "0.000 m³";
  if (v >= 10000000) {
    return `${(v / 10000000).toFixed(3)} Cr m³`;
  } else if (v >= 100000) {
    return `${(v / 100000).toFixed(3)} Lk m³`;
  }
  return `${Number(v).toLocaleString(undefined, { minimumFractionDigits: 3, maximumFractionDigits: 3 })} m³`;
}
const CROSS_SVG_W = 1700;
const CROSS_SVG_H = 980;
const CROSS_VIEW_MARGIN_X = 420;
const CROSS_VIEW_MARGIN_Y = 180;
const CROSS_DEFAULT_VIEWBOX = {
  x: -CROSS_VIEW_MARGIN_X,
  y: -CROSS_VIEW_MARGIN_Y,
  w: CROSS_SVG_W + (CROSS_VIEW_MARGIN_X * 2),
  h: CROSS_SVG_H + (CROSS_VIEW_MARGIN_Y * 2),
};

const BRIDGE_CATEGORIES = ["Minor", "Major", "Viaduct", "Important", "RoR", "Tunnel", "ROB", "MIBOR", "Aqueduct"];
const BRIDGE_TYPES = ["Box", "PSC Slab", "Composite Girder", "OWG", "Other"];
const BRIDGE_DEDUCT_RULES = ["Auto", "Always", "Never"];
const LOOP_LINE_TYPES = ["Loop", "TM Siding", "Ballast Siding", "Connecting Line", "Main Line", "Platform"];
const LOOP_SIDES = ["Left", "Right"];

function createEmptyLoopRow(index = 0, overrides = {}) {
  return {
    station: `LP-${index + 1}`,
    lineType: "",
    lineName: "",
    side: "",
    csb: null,
    tc: 0,
    loopStartCh: null,
    loopEndCh: null,
    pfWidth: 0,
    pfStartCh: null,
    pfEndCh: null,
    remarks: "",
    ...overrides,
  };
}
const state = {
  meta: null,
  rawRows: [],
  bridgeRows: [],
  curveRows: [],
  loopPlatformRows: [],
  seedRows: [],
  seedMeta: null,
  calcRows: [],
  settings: {},
  defaultSettings: {},
  seedDefaultSettings: {},
  project: {
    active: false,
    verified: false,
    name: "",
    uploads: {
      levels: false,
      curves: false,
      bridges: false,
      loops: false,
      kml: false,
    },
    profile: {
      corridorName: "",
      direction: "Up",
      chainageZeroRef: "",
    },
  },
  calcOverrides: [],
  slopeZones: [],
  importMappings: { client: "", templates: {} },
  mappingWizard: { aoa: null, headerRow: 0, headers: [], kind: "levels" },
  activeWorkPage: "overview",
  activeResultTab: "inputs",
  charts: { lSection: null, volume: null },
  currentCrossIndex: null,
  crossViewBox: { ...CROSS_DEFAULT_VIEWBOX },
  crossPan: { active: false, lastX: 0, lastY: 0 },
  crossTouch: { mode: "none", startDistance: 0, startViewBox: null, startCenter: null, lastX: 0, lastY: 0 },
  crossZoomFrame: null,
  kmlData: null,
  stationPlans: {},
  projectFileHandle: null,
  auth: {
    authenticated: false,
    user: "",
  },
};

const els = {
  loginScreen: document.getElementById("loginScreen"),
  loginForm: document.getElementById("loginForm"),
  loginUsername: document.getElementById("loginUsername"),
  loginPassword: document.getElementById("loginPassword"),
  loginError: document.getElementById("loginError"),
  loginSubmitBtn: document.getElementById("loginSubmitBtn"),
  logoutBtn: document.getElementById("logoutBtn"),
  loginUserChip: document.getElementById("loginUserChip"),
  projectMeta: document.getElementById("projectMeta"),
  rollDiagramCanvas: document.getElementById("rollDiagramCanvas"),
  rollDiagramWrap: document.getElementById("rollDiagramWrap"),
  rollDiagramEmpty: document.getElementById("rollDiagramEmpty"),
  sideViewCanvas: document.getElementById("sideViewCanvas"),
  sideViewWrap: document.getElementById("sideViewWrap"),
  rollPageHeader: document.getElementById("rollPageHeader"),
  rollSplitShell: document.getElementById("rollSplitShell"),
  rollFilterSelect: document.getElementById("rollFilterSelect"),
  rollZoomFit: document.getElementById("rollZoomFit"),
  rollFullscreenBtn: document.getElementById("rollFullscreenBtn"),
  rollFullscreenCloseBtn: document.getElementById("rollFullscreenCloseBtn"),
  rollFullscreenModal: document.getElementById("rollFullscreenModal"),
  rollFullscreenViewport: document.getElementById("rollFullscreenViewport"),
  tableBody: document.getElementById("tableBody"),
  exportExcelBtn: document.getElementById("exportExcelBtn"),
  totalFilling: document.getElementById("totalFilling"),
  totalCutting: document.getElementById("totalCutting"),
  fillLength: document.getElementById("fillLength"),
  cutLength: document.getElementById("cutLength"),
  bridgeImportInput: document.getElementById("bridgeImportInput"),
  curveImportInput: document.getElementById("curveImportInput"),
  loopImportInput: document.getElementById("loopImportInput"),
  importKmlBtn: document.getElementById("importKmlBtn"),
  importStationPlanBtn: document.getElementById("importStationPlanBtn"),
  kmlImportInput: document.getElementById("kmlImportInput"),
  stationPlanImportInput: document.getElementById("stationPlanImportInput"),
  importOptionsModal: document.getElementById("importOptionsModal"),
  closeImportOptionsBtn: document.getElementById("closeImportOptionsBtn"),
  importOptionButtons: Array.from(document.querySelectorAll("[data-import-kind]")),
  projectWizardModal: document.getElementById("projectWizardModal"),
  closeProjectWizardBtn: document.getElementById("closeProjectWizardBtn"),
  projectNameInput: document.getElementById("projectNameInput"),
  wizardStatus: document.getElementById("wizardStatus"),
  wizardTickLevels: document.getElementById("wizardTickLevels"),
  wizardTickCurves: document.getElementById("wizardTickCurves"),
  wizardTickBridges: document.getElementById("wizardTickBridges"),
  wizardTickLoops: document.getElementById("wizardTickLoops"),
  wizardTickKml: document.getElementById("wizardTickKml"),
  wizardSaveBtn: document.getElementById("wizardSaveBtn"),
  wizardCalculateBtn: document.getElementById("wizardCalculateBtn"),
  wizardUploadButtons: Array.from(document.querySelectorAll("[data-wizard-upload]")),
  bridgeAddBtn: document.getElementById("bridgeAddBtn"),
  bridgeApplyBtn: document.getElementById("bridgeApplyBtn"),
  bridgeTableBody: document.getElementById("bridgeTableBody"),
  bridgeMeta: document.getElementById("bridgeMeta"),
  curveTableBody: document.getElementById("curveTableBody"),
  curveMeta: document.getElementById("curveMeta"),
  loopTableBody: document.getElementById("loopTableBody"),
  loopMeta: document.getElementById("loopMeta"),
  curveAddBtn: document.getElementById("curveAddBtn"),
  curveApplyBtn: document.getElementById("curveApplyBtn"),
  loopAddBtn: document.getElementById("loopAddBtn"),
  loopApplyBtn: document.getElementById("loopApplyBtn"),
  workNav: document.getElementById("workNav"),
  workPageButtons: Array.from(document.querySelectorAll("[data-work-page-btn]")),
  workPages: Array.from(document.querySelectorAll("[data-work-page]")),
  resultTabs: document.getElementById("resultTabs"),
  resultTabButtons: Array.from(document.querySelectorAll("[data-result-tab]")),
  resultTabPanes: Array.from(document.querySelectorAll("[data-result-pane]")),
  resultInputBody: document.getElementById("resultInputBody"),
  resultFillBody: document.getElementById("resultFillBody"),
  resultCutBody: document.getElementById("resultCutBody"),
  resultQtyBody: document.getElementById("resultQtyBody"),
  importInput: document.getElementById("importInput"),
  projectImportInput: document.getElementById("projectImportInput"),
  importBtn: document.getElementById("importBtn"),
  createProjectBtn: document.getElementById("createProjectBtn"),
  themeToggleCheckbox: document.getElementById("themeToggleCheckbox"),
  importProjectBtn: document.getElementById("importProjectBtn"),
  saveProjectBtn: document.getElementById("saveProjectBtn"),
  resetProjectBtn: document.getElementById("resetProjectBtn"),
  snapshotsBtn: document.getElementById("snapshotsBtn"),
  snapshotsModal: document.getElementById("snapshotsModal"),
  closeSnapshotsBtn: document.getElementById("closeSnapshotsBtn"),
  takeSnapshotBtn: document.getElementById("takeSnapshotBtn"),
  snapshotNameInput: document.getElementById("snapshotNameInput"),
  snapshotList: document.getElementById("snapshotList"),
  openSettingsBtn: document.getElementById("openSettingsBtn"),
  settingsModal: document.getElementById("settingsModal"),
  settingsForm: document.getElementById("settingsForm"),
  closeSettingsBtn: document.getElementById("closeSettingsBtn"),
  settingsGrid: document.getElementById("settingsGrid"),
  profileGrid: document.getElementById("profileGrid"),
  standardsGrid: document.getElementById("standardsGrid"),
  qualityGrid: document.getElementById("qualityGrid"),
  materialProfileGrid: document.getElementById("materialProfileGrid"),
  visualGrid: document.getElementById("visualGrid"),
  mapGrid: document.getElementById("mapGrid"),
  boqGrid: document.getElementById("boqGrid"),
  downloadBoqBtn: document.getElementById("downloadBoqBtn"),
  openQualityBtn: document.getElementById("openQualityBtn"),
  qualityModal: document.getElementById("qualityModal"),
  qualityResults: document.getElementById("qualityResults"),
  closeQualityBtn: document.getElementById("closeQualityBtn"),
  openOverridesBtn: document.getElementById("openOverridesBtn"),
  openSlopeZonesBtn: document.getElementById("openSlopeZonesBtn"),
  openBermRulesBtn: document.getElementById("openBermRulesBtn"),
  openMappingWizardBtn: document.getElementById("openMappingWizardBtn"),
  mappingWizardModal: document.getElementById("mappingWizardModal"),
  closeMappingWizardBtn: document.getElementById("closeMappingWizardBtn"),
  mappingClientInput: document.getElementById("mappingClientInput"),
  mappingTypeSelect: document.getElementById("mappingTypeSelect"),
  mappingFileInput: document.getElementById("mappingFileInput"),
  mappingGrid: document.getElementById("mappingGrid"),
  saveMappingTemplateBtn: document.getElementById("saveMappingTemplateBtn"),
  applyMappingTemplateBtn: document.getElementById("applyMappingTemplateBtn"),
  openStationLayoutBtn: document.getElementById("openStationLayoutBtn"),
  openStationLayoutInlineBtn: document.getElementById("openStationLayoutInlineBtn"),
  stationLayoutModal: document.getElementById("stationLayoutModal"),
  closeStationLayoutBtn: document.getElementById("closeStationLayoutBtn"),
  stationLayoutSelect: document.getElementById("stationLayoutSelect"),
  stationLayoutList: document.getElementById("stationLayoutList"),
  saveStationLayoutBtn: document.getElementById("saveStationLayoutBtn"),
  overridesModal: document.getElementById("overridesModal"),
  closeOverridesBtn: document.getElementById("closeOverridesBtn"),
  overrideTableBody: document.getElementById("overrideTableBody"),
  addOverrideBtn: document.getElementById("addOverrideBtn"),
  applyOverridesBtn: document.getElementById("applyOverridesBtn"),
  slopeZonesModal: document.getElementById("slopeZonesModal"),
  closeSlopeZonesBtn: document.getElementById("closeSlopeZonesBtn"),
  slopeZoneTableBody: document.getElementById("slopeZoneTableBody"),
  addSlopeZoneBtn: document.getElementById("addSlopeZoneBtn"),
  applySlopeZonesBtn: document.getElementById("applySlopeZonesBtn"),
  snapshotCompareModal: document.getElementById("snapshotCompareModal"),
  closeSnapshotCompareBtn: document.getElementById("closeSnapshotCompareBtn"),
  snapshotCompareBody: document.getElementById("snapshotCompareBody"),
  reportLogoInput: document.getElementById("reportLogoInput"),
  reportSignatureInput: document.getElementById("reportSignatureInput"),
  reportSignName: document.getElementById("reportSignName"),
  reportSignTitle: document.getElementById("reportSignTitle"),
  resetDefaultsBtn: document.getElementById("resetDefaultsBtn"),
  lSectionChart: document.getElementById("lSectionChart"),
  volumeChart: document.getElementById("volumeChart"),
  crossSectionModal: document.getElementById("crossSectionModal"),
  crossTitle: document.getElementById("crossTitle"),
  crossPrevBtn: document.getElementById("crossPrevBtn"),
  crossNextBtn: document.getElementById("crossNextBtn"),
  closeCrossBtn: document.getElementById("closeCrossBtn"),
  crossSvg: document.getElementById("crossSvg"),
  rateClearing: document.getElementById("rateClearing"),
  rateBenching: document.getElementById("rateBenching"),
  rateFilling: document.getElementById("rateFilling"),
  rateBlanketing: document.getElementById("rateBlanketing"),
  rateCutSoil: document.getElementById("rateCutSoil"),
  rateCutSoft: document.getElementById("rateCutSoft"),
  rateCutHardBlast: document.getElementById("rateCutHardBlast"),
  rateCutHardChisel: document.getElementById("rateCutHardChisel"),
  rateExtraLead: document.getElementById("rateExtraLead"),
  rateTurfing: document.getElementById("rateTurfing"),
  pctSoil: document.getElementById("pctSoil"),
  pctSoftRock: document.getElementById("pctSoftRock"),
  pctHardBlast: document.getElementById("pctHardBlast"),
  pctHardChisel: document.getElementById("pctHardChisel"),
  pctReusableSpoil: document.getElementById("pctReusableSpoil"),
  leadKm: document.getElementById("leadKm"),
  recalcEstimatesBtn: document.getElementById("recalcEstimatesBtn"),
  estimatesBody: document.getElementById("estimatesBody"),
  estimateGrandTotal: document.getElementById("estimateGrandTotal"),
  crossGraphicWrap: document.getElementById("crossGraphicWrap"),
  zoomInBtn: document.getElementById("zoomInBtn"),
  zoomOutBtn: document.getElementById("zoomOutBtn"),
  zoomResetBtn: document.getElementById("zoomResetBtn"),
  layerTbody: document.getElementById("layerTbody"),
  dimTbody: document.getElementById("dimTbody"),
  actualFilling: document.getElementById("actualFilling"),
  reusableFilling: document.getElementById("reusableFilling"),
  fillReusableHatch: document.getElementById("fillReusableHatch"),
  fillReusablePctLabel: document.getElementById("fillReusablePctLabel"),
  actualCutting: document.getElementById("actualCutting"),
  crossMeta: document.getElementById("crossMeta"),
  graphModal: document.getElementById("graphModal"),
  graphModalTitle: document.getElementById("graphModalTitle"),
  closeGraphModalBtn: document.getElementById("closeGraphModalBtn"),
  expandedGraphCanvas: document.getElementById("expandedGraphCanvas"),
  fillWaterNode: document.getElementById("fillWaterNode"),
  cutWaterNode: document.getElementById("cutWaterNode"),
  sidebarToggle: document.getElementById("sidebarToggle"),
  appLayout: document.querySelector(".app-layout"),
  openExportModalBtn: document.getElementById("openExportModalBtn"),
  exportModal: document.getElementById("exportModal"),
  confirmExportBtn: document.getElementById("confirmExportBtn"),
  cancelExportBtn: document.getElementById("cancelExportBtn"),
};

async function loadAuthState() {
  const isLocalDev =
    window.location.hostname === "127.0.0.1" ||
    window.location.hostname === "localhost";

  if (isLocalDev) {
    state.auth = {
      authenticated: true,
      user: "Local Dev",
    };
    return;
  }

  try {
    const response = await fetch("/api/session", {
      credentials: "same-origin",
    });
    if (!response.ok) return;
    const session = await response.json();
    state.auth = {
      authenticated: Boolean(session?.authenticated),
      user: String(session?.user || "").trim(),
    };
  } catch (error) {
    console.warn("Failed to load auth session:", error);
  }
}

function updateAuthUI() {
  const authenticated = Boolean(state.auth?.authenticated);
  document.body.classList.toggle("is-authenticated", authenticated);

  if (els.logoutBtn) els.logoutBtn.style.display = authenticated ? "inline-flex" : "none";
  if (els.loginUserChip) {
    els.loginUserChip.style.display = authenticated ? "inline-flex" : "none";
    els.loginUserChip.textContent = authenticated && state.auth.user ? `Signed in: ${state.auth.user}` : "";
  }
  if (els.loginError) els.loginError.textContent = "";

  if (!authenticated) {
    requestAnimationFrame(() => {
      els.loginUsername?.focus();
    });
  }
}

function setAuthState(authenticated, user = "") {
  state.auth = {
    authenticated: Boolean(authenticated),
    user: authenticated ? String(user || "").trim() : "",
  };
  updateAuthUI();
}

async function logout() {
  try {
    await fetch("/api/logout", {
      method: "POST",
      credentials: "same-origin",
    });
  } catch (error) {
    console.warn("Failed to clear auth session:", error);
  }
  setAuthState(false, "");
  window.location.replace("/");
}

async function attemptLogin() {
  const username = String(els.loginUsername?.value || "").trim();
  const password = String(els.loginPassword?.value || "").trim();

  if (!username || !password) {
    if (els.loginError) els.loginError.textContent = "Enter both username and password to continue.";
    return;
  }

  try {
    const response = await fetch("/api/login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      credentials: "same-origin",
      body: JSON.stringify({
        username: username.toLowerCase(),
        password,
      }),
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      if (els.loginError) els.loginError.textContent = payload.error || "Invalid username or password.";
      if (els.loginPassword) els.loginPassword.value = "";
      els.loginPassword?.focus();
      return;
    }
    setAuthState(true, payload.user || username.toLowerCase());
    if (els.loginPassword) els.loginPassword.value = "";
    setWorkPage("overview");
    window.scrollTo({ top: 0, behavior: "smooth" });
  } catch (error) {
    if (els.loginError) els.loginError.textContent = "Login is unavailable right now.";
    if (els.loginPassword) els.loginPassword.value = "";
    els.loginPassword?.focus();
  }
}

function formatDashboardChainage(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return "-";
  const sign = n < 0 ? "-" : "";
  const abs = Math.abs(n);
  const km = Math.floor(abs / 1000);
  const m = Math.round(abs - (km * 1000));
  return `${sign}${km}+${String(m).padStart(3, "0")}`;
}

function formatCompactVolume(v) {
  return Number.isFinite(v)
    ? Number(v).toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 })
    : "0";
}

function getMedianInterval(rows) {
  const diffs = [];
  for (let i = 1; i < rows.length; i += 1) {
    const diff = safeNum(rows[i].chainage) - safeNum(rows[i - 1].chainage);
    if (diff > 0) diffs.push(diff);
  }
  if (!diffs.length) return NaN;
  diffs.sort((a, b) => a - b);
  const mid = Math.floor(diffs.length / 2);
  return diffs.length % 2 ? diffs[mid] : (diffs[mid - 1] + diffs[mid]) / 2;
}

function getGroupedStations(rows = state.loopPlatformRows || []) {
  const groups = new Map();
  rows.forEach((lp) => {
    const stationName = String(lp.station || "").trim();
    if (!stationName) return;
    const key = stationName.toLowerCase();
    const existing = groups.get(key) || {
      station: stationName,
      csb: Number.isFinite(parseChainage(lp.csb)) ? parseChainage(lp.csb) : NaN,
      tc: 0,
      pfWidth: 0,
      loopStartCh: Number.POSITIVE_INFINITY,
      loopEndCh: Number.NEGATIVE_INFINITY,
      remarks: [],
    };
    const loopStart = parseChainage(lp.loopStartCh);
    const loopEnd = parseChainage(lp.loopEndCh);
    if (Number.isFinite(loopStart)) existing.loopStartCh = Math.min(existing.loopStartCh, loopStart);
    if (Number.isFinite(loopEnd)) existing.loopEndCh = Math.max(existing.loopEndCh, loopEnd);
    existing.tc = Math.max(existing.tc, safeNum(lp.tc, 0));
    existing.pfWidth = Math.max(existing.pfWidth, safeNum(lp.pfWidth, 0));
    if (lp.remarks) existing.remarks.push(String(lp.remarks).trim());
    groups.set(key, existing);
  });
  return Array.from(groups.values()).map((station) => ({
    ...station,
    loopStartCh: Number.isFinite(station.loopStartCh) ? station.loopStartCh : NaN,
    loopEndCh: Number.isFinite(station.loopEndCh) ? station.loopEndCh : NaN,
  }));
}

function resolveStationNameAtChainage(chainage, groupedStations, tolerance = 25) {
  if (!Number.isFinite(chainage) || !Array.isArray(groupedStations) || !groupedStations.length) return "";
  const candidates = [];
  groupedStations.forEach((st) => {
    const loopStart = Number.isFinite(st.loopStartCh) ? st.loopStartCh : NaN;
    const loopEnd = Number.isFinite(st.loopEndCh) ? st.loopEndCh : NaN;
    const inLoop = Number.isFinite(loopStart) && Number.isFinite(loopEnd) && chainage >= loopStart && chainage <= loopEnd;
    const csb = Number.isFinite(st.csb) ? st.csb : NaN;
    const csbDist = Number.isFinite(csb) ? Math.abs(chainage - csb) : NaN;
    const nearCsb = Number.isFinite(csbDist) && csbDist <= tolerance;
    if (inLoop || nearCsb) {
      const span = Number.isFinite(loopStart) && Number.isFinite(loopEnd) ? Math.max(loopEnd - loopStart, 0) : Number.POSITIVE_INFINITY;
      candidates.push({ station: st.station, inLoop, csbDist, span });
    }
  });
  if (!candidates.length) return "";
  const inLoopCandidates = candidates.filter((c) => c.inLoop);
  const pool = inLoopCandidates.length ? inLoopCandidates : candidates;
  pool.sort((a, b) => {
    if (Number.isFinite(a.csbDist) && Number.isFinite(b.csbDist) && a.csbDist !== b.csbDist) return a.csbDist - b.csbDist;
    if (a.span !== b.span) return a.span - b.span;
    return String(a.station).localeCompare(String(b.station));
  });
  return String(pool[0].station || "").trim();
}

function normalizeStationKey(name) {
  return String(name || "").trim().toLowerCase();
}

function getStationRowsByNameOrChainage(chainage, stationName) {
  const allRows = Array.isArray(state.loopPlatformRows) ? state.loopPlatformRows : [];
  if (!allRows.length) return { rows: [], station: "" };
  const key = normalizeStationKey(stationName);
  let rows = key ? allRows.filter((r) => normalizeStationKey(r.station) === key) : [];
  let resolved = stationName;
  if (!rows.length && Number.isFinite(chainage)) {
    const groupedStations = getGroupedStations(allRows);
    const inferred = resolveStationNameAtChainage(chainage, groupedStations, 25);
    if (inferred) {
      const inferredKey = normalizeStationKey(inferred);
      rows = allRows.filter((r) => normalizeStationKey(r.station) === inferredKey);
      resolved = inferred;
    }
  }
  return { rows, station: resolved || (rows[0]?.station || "") };
}

function buildStationSequenceLayout(chainage, stationName, standardTc = 5.3, options = {}) {
  const useRanges = options.useRanges !== false;
  const { rows } = getStationRowsByNameOrChainage(chainage, stationName);
  if (!rows.length) return null;
  const ordered = rows
    .map((row, index) => ({
      row,
      order: Number.isFinite(row.order) ? row.order : index,
    }))
    .sort((a, b) => a.order - b.order);

  const activeItems = [];
  ordered.forEach((entry) => {
    const row = entry.row;
    const lineType = normalizeLoopLineType(row.lineType || row.lineName || "");
    const side = normalizeLoopSide(row.side || "");
    const isPlatform = lineType === "Platform" || safeNum(row.pfWidth, 0) > 0;
    const isMain = lineType === "Main Line";
    const loopStart = parseChainage(row.loopStartCh);
    const loopEnd = parseChainage(row.loopEndCh);
    const pfStart = parseChainage(row.pfStartCh);
    const pfEnd = parseChainage(row.pfEndCh);
    const hasLoopRange = Number.isFinite(loopStart) && Number.isFinite(loopEnd) && loopEnd >= loopStart;
    const hasPfRange = Number.isFinite(pfStart) && Number.isFinite(pfEnd) && pfEnd >= pfStart;
    const inLoop = hasLoopRange && Number.isFinite(chainage) && chainage >= loopStart && chainage <= loopEnd;
    const inPf = hasPfRange && Number.isFinite(chainage) && chainage >= pfStart && chainage <= pfEnd;
    const tc = safeNum(row.tc, 0);
    const pfWidth = safeNum(row.pfWidth, 0);

    const trackActive = isMain || (!useRanges
      ? (!isPlatform && (lineType || tc > 0))
      : (hasLoopRange ? inLoop : (tc > 0 && !isPlatform)));
    const platformActive = (isPlatform || pfWidth > 0)
      && hasPfRange
      && (!Number.isFinite(chainage) || inPf);

    if (trackActive && !isPlatform) {
      activeItems.push({
        kind: "track",
        row,
        order: entry.order,
        side,
        lineType,
        isMain,
      });
    }
    if (platformActive) {
      activeItems.push({
        kind: "platform",
        row,
        order: entry.order,
        side,
        lineType,
        isMain: false,
      });
    }
  });

  const trackItems = activeItems.filter((item) => item.kind === "track");
  if (!trackItems.length) return null;

  const mainTracks = trackItems.filter((item) => item.isMain);
  let refMain = mainTracks.find((item) => !item.side) || mainTracks[0] || trackItems[0];
  const refOrder = refMain.order;

  trackItems.forEach((item) => {
    if (!item.side) {
      item.side = item.order < refOrder ? "Left" : (item.order > refOrder ? "Right" : "");
    }
  });
  activeItems.forEach((item) => {
    if (item.kind !== "platform" || item.side) return;
    item.side = item.order < refOrder ? "Left" : (item.order > refOrder ? "Right" : "");
  });

  const leftTracks = trackItems
    .filter((item) => item !== refMain && item.side === "Left")
    .sort((a, b) => Math.abs(a.order - refOrder) - Math.abs(b.order - refOrder));
  const rightTracks = trackItems
    .filter((item) => item !== refMain && item.side === "Right")
    .sort((a, b) => Math.abs(a.order - refOrder) - Math.abs(b.order - refOrder));

  const offsetByItem = new Map();
  offsetByItem.set(refMain, 0);
  let acc = 0;
  leftTracks.forEach((item) => {
    const step = safeNum(item.row.tc, 0);
    const gap = step > 0 ? step : standardTc;
    acc += gap;
    offsetByItem.set(item, acc);
  });
  acc = 0;
  rightTracks.forEach((item) => {
    const step = safeNum(item.row.tc, 0);
    const gap = step > 0 ? step : standardTc;
    acc += gap;
    offsetByItem.set(item, -acc);
  });

  const orderedItems = [...activeItems].sort((a, b) => a.order - b.order);
  return {
    refMain,
    refOrder,
    trackItems,
    platformItems: activeItems.filter((item) => item.kind === "platform"),
    orderedItems,
    offsetByItem,
  };
}

const settingSchema = [
  ["formationWidthFill", "Formation Width (Fill) m"],
  ["cuttingWidth", "Formation Width (Cut) m"],
  ["blanketThickness", "Blanket Thickness m"],
  ["preparedSubgradeThickness", "Prepared Subgrade Thickness m"],
  ["bermWidth", "Berm Width m"],
  ["sideSlopeFactor", "Side Slope Factor (H:V)"],
  ["ballastCushionThickness", "Ballast Cushion Thickness m"],
  ["topLayerThickness", "Top Layer of Embankment Fill m"],
  ["activeSqCategory", "Soil Category (SQ1/SQ2/SQ3)"],
];

function parseChainage(value) {
  if (typeof value === "number") return value;
  if (typeof value !== "string") return NaN;
  const t = value.trim();
  if (!t) return NaN;
  if (t.startsWith("=")) {
    const expr = t.slice(1).trim();
    if (/^[0-9+\-*/().\s]+$/.test(expr)) {
      try {
        const v = Function(`"use strict"; return (${expr});`)();
        return Number.isFinite(v) ? Number(v) : NaN;
      } catch (_) {
        return NaN;
      }
    }
    return NaN;
  }
  if (t.includes("+")) {
    const [km, m] = t.split("+");
    const k = Number(km.replace(/[^\d.-]/g, ""));
    const mm = Number((m || "0").replace(/[^\d.-]/g, ""));
    if (Number.isFinite(k) && Number.isFinite(mm)) return (k * 1000) + mm;
  }
  const stripped = t.replace(/[^\d.-]/g, "");
  if (!stripped) return NaN;
  const n = Number(stripped);
  return Number.isFinite(n) ? n : NaN;
}

function safeNum(v, fallback = 0) {
  const n = Number(v);
  return Number.isFinite(n) ? n : fallback;
}

function parseLooseNumber(v, fallback = NaN) {
  if (typeof v === "number") return Number.isFinite(v) ? v : fallback;
  if (typeof v !== "string") return fallback;
  const t = v.trim().replace(/,/g, "");
  if (!t) return fallback;
  if (t.startsWith("=")) {
    const expr = t.slice(1).trim();
    if (/^[0-9+\-*/().\s]+$/.test(expr)) {
      try {
        const result = Function(`"use strict"; return (${expr});`)();
        return Number.isFinite(result) ? Number(result) : fallback;
      } catch (_) {
        return fallback;
      }
    }
    return fallback;
  }
  const n = Number(t);
  if (Number.isFinite(n)) return n;
  const m = t.match(/-?\d+(?:\.\d+)?/);
  if (!m) return fallback;
  const parsed = Number(m[0]);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function normalizeHeaderToken(v) {
  return String(v || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function normalizeLoopLineType(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  const token = normalizeHeaderToken(raw);
  if (token === "main" || token === "mainline") return "Main Line";
  if (token === "loop") return "Loop";
  if (token === "platform") return "Platform";
  if (token === "tmsiding") return "TM Siding";
  if (token === "ballastsiding") return "Ballast Siding";
  if (token === "connectingline") return "Connecting Line";
  return raw;
}

function normalizeLoopSide(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  const token = normalizeHeaderToken(raw);
  if (token === "left" || token === "lhs") return "Left";
  if (token === "right" || token === "rhs") return "Right";
  return raw;
}

function findColByAliases(headers, aliases) {
  for (const a of aliases) {
    const idx = headers.findIndex((h) => {
      const n = normalizeHeaderToken(h);
      return n === a || n.includes(a);
    });
    if (idx >= 0) return idx;
  }
  return -1;
}

function resolveImportColumns(headerCells) {
  const cols = {
    chainage: findColByAliases(headerCells, ["chainage", "ch", "km", "chainagem"]),
    groundLevel: findColByAliases(headerCells, ["groundlevel", "groundrl", "existinglevel", "existingrl", "naturalgroundlevel", "naturalground", "gl", "ngrl", "ngl"]),
    proposedLevel: findColByAliases(headerCells, ["proposedlevel", "proposedrl", "formationlevel", "formationrl", "frl", "finishedlevel", "finallevel", "fl", "pl", "prl"]),
    station: findColByAliases(headerCells, ["station", "stn"]),
    structureNo: findColByAliases(headerCells, ["structureno", "structure", "bridge", "strno"]),
  };
  return cols;
}

function detectImportHeaderRow(aoa) {
  const scanRows = Math.min(aoa.length, 30);
  let best = { rowIndex: -1, score: -1, cols: null };
  for (let i = 0; i < scanRows; i += 1) {
    const row = Array.isArray(aoa[i]) ? aoa[i] : [];
    if (!row.length) continue;
    const cols = resolveImportColumns(row);
    const score = (cols.groundLevel >= 0 ? 3 : 0)
      + (cols.proposedLevel >= 0 ? 3 : 0)
      + (cols.chainage >= 0 ? 2 : 0)
      + (cols.station >= 0 ? 1 : 0)
      + (cols.structureNo >= 0 ? 1 : 0);
    if (score > best.score) {
      best = { rowIndex: i, score, cols };
    }
  }
  if (best.rowIndex < 0) return null;
  if (best.cols.groundLevel < 0 || best.cols.proposedLevel < 0) return null;
  return best;
}

function updateDashboard() {
  const dashChRange = document.getElementById("dashChRange");
  const dashLength = document.getElementById("dashLength");
  const dashPoints = document.getElementById("dashPoints");
  const dashStructures = document.getElementById("dashStructures");
  const heroTitle = document.getElementById("overviewHeroTitle");
  const heroText = document.getElementById("overviewHeroText");
  const healthGrid = document.getElementById("projectHealthGrid");
  const criticalList = document.getElementById("criticalChainagesList");
  const corridor = document.getElementById("corridorIntelligence");
  const balanceGrid = document.getElementById("volumeBalanceGrid");
  const alertsList = document.getElementById("designAlertsList");

  const projectName = state.project?.name || "No active project";
  const active = Boolean(state.project?.active);
  const uploads = state.project?.uploads || {};
  const rawRows = Array.isArray(state.rawRows) ? state.rawRows : [];
  const calcRows = Array.isArray(state.calcRows) ? state.calcRows : [];
  const groupedStations = getGroupedStations();
  const stationPlanCount = Object.keys(state.stationPlans || {}).length;
  const validBridges = Array.isArray(state.bridgeRows) ? state.bridgeRows.length : 0;
  const validCurves = Array.isArray(state.curveRows) ? state.curveRows.length : 0;
  const savedAt = state.meta?.lastSavedAt ? new Date(state.meta.lastSavedAt) : null;
  const fillTotal = calcRows.reduce((s, r) => s + safeNum(r.fillVol, 0), 0);
  const cutTotal = calcRows.reduce((s, r) => s + safeNum(r.cutVol, 0), 0);
  const reusablePct = parseLooseNumber(els.pctReusableSpoil?.value, 60);
  const reusableSpoil = cutTotal * (safeNum(reusablePct, 60) / 100);
  const netBalance = reusableSpoil - fillTotal;

  if (!calcRows || calcRows.length === 0) {
    if (dashChRange) dashChRange.textContent = "0.000 to 0.000";
    if (dashLength) dashLength.textContent = "0.000 km";
    if (dashPoints) dashPoints.textContent = "0 Cross-Sections";
    if (dashStructures) dashStructures.textContent = "0 Bridges, 0 Curves";
  } else {
    const minCh = calcRows[0].chainage;
    const maxCh = calcRows[calcRows.length - 1].chainage;
    const totalL = Math.max(maxCh - minCh, 0);
    if (dashChRange) {
      dashChRange.textContent = `${formatDashboardChainage(minCh)} to ${formatDashboardChainage(maxCh)}`;
    }
    if (dashLength) {
      dashLength.textContent = `${r3(totalL / 1000)} km`;
    }
    if (dashPoints) {
      dashPoints.textContent = `${calcRows.length.toLocaleString()} Cross-Sections`;
    }
    if (dashStructures) {
      dashStructures.textContent = `${validBridges} Bridges, ${validCurves} Curves`;
    }
  }

  if (heroTitle) {
    heroTitle.textContent = active ? projectName : "No active project";
  }
  if (heroText) {
    if (!active) {
      heroText.textContent = "Create or open a project to begin importing levels, structures, loops, and map alignment.";
    } else if (!state.project.verified) {
      heroText.textContent = uploads.kml
        ? "Project is active with mapped alignment. Complete review and verify before export."
        : "Project is active. Core datasets can be verified now; add KML/KMZ to unlock geographic intelligence.";
    } else {
      heroText.textContent = uploads.kml
        ? "Project is verified and geographically mapped. Review chainage hotspots, alerts, and export readiness."
        : "Project is verified. Add KML/KMZ alignment to unlock full corridor intelligence and station mapping.";
    }
  }

  if (healthGrid) {
    const importCount = ["levels", "curves", "bridges", "loops"].filter((k) => uploads[k]).length;
    healthGrid.innerHTML = [
      ["Project", projectName, active ? (state.project.verified ? "Verified workspace" : "Draft workspace") : "No workspace loaded", active ? (state.project.verified ? "success" : "warning") : "danger"],
      ["Core Inputs", `${importCount}/4 loaded`, uploads.kml ? "KML/KMZ alignment loaded" : "KML/KMZ optional", importCount === 4 ? "success" : (importCount > 0 ? "warning" : "danger")],
      ["Map Readiness", uploads.kml ? "Ready" : "Pending", uploads.kml && state.kmlData?.totalDistance ? `${r3(state.kmlData.totalDistance / 1000)} km mapped` : "No geographic alignment", uploads.kml ? "info" : "warning"],
      ["Stations", `${groupedStations.length}`, `${stationPlanCount} plans attached`, groupedStations.length > 0 ? (stationPlanCount > 0 ? "info" : "warning") : "danger"],
      ["Verification", state.project?.verified ? "Passed" : "Pending", active ? "Use Verify before export" : "Create project first", state.project?.verified ? "success" : (active ? "warning" : "danger")],
      ["Last Saved", savedAt ? savedAt.toLocaleString() : "Local draft", savedAt ? "Persisted in browser/project file" : "Not saved in this session", savedAt ? "neutral" : "warning"],
    ].map(([label, value, sub, tone]) => `
      <div class="mission-chip mission-chip-${tone}">
        <div class="mission-chip-label">${label}</div>
        <div class="mission-chip-value">${value}</div>
        <div class="mission-chip-sub">${sub}</div>
      </div>
    `).join("");
  }

  if (criticalList) {
    const highestFill = calcRows.reduce((best, row) => safeNum(row.bank) > safeNum(best?.bank, -1) ? row : best, null);
    const highestCut = calcRows.reduce((best, row) => safeNum(row.cut) > safeNum(best?.cut, -1) ? row : best, null);
    const firstBridge = [...(state.bridgeRows || [])]
      .sort((a, b) => safeNum(a.startChainage) - safeNum(b.startChainage))[0];
    const sharpestCurve = [...(state.curveRows || [])]
      .filter((c) => safeNum(c.radius, NaN) > 0)
      .sort((a, b) => safeNum(a.radius) - safeNum(b.radius))[0];
    const stationComplexity = [...groupedStations]
      .sort((a, b) => (safeNum(b.tc) + safeNum(b.pfWidth)) - (safeNum(a.tc) + safeNum(a.pfWidth)))[0];
    const criticalItems = [
      highestFill ? ["Highest Fill", `Bank ${r3(highestFill.bank)} m`, formatDashboardChainage(highestFill.chainage)] : null,
      highestCut ? ["Highest Cut", `Cut ${r3(highestCut.cut)} m`, formatDashboardChainage(highestCut.chainage)] : null,
      firstBridge ? ["First Bridge Deduction", `${firstBridge.bridgeNo || "Bridge"} • ${firstBridge.bridgeType || "-"}`, formatDashboardChainage(firstBridge.startChainage)] : null,
      sharpestCurve ? ["Sharpest Curve", `${sharpestCurve.curve || "Curve"} • R=${r3(sharpestCurve.radius)} m`, formatDashboardChainage(sharpestCurve.chainage)] : null,
      stationComplexity ? ["Station Yard Complexity", `${stationComplexity.station} • TC ${r3(stationComplexity.tc)} m`, Number.isFinite(stationComplexity.csb) ? formatDashboardChainage(stationComplexity.csb) : "No CSB"] : null,
    ].filter(Boolean);
    criticalList.innerHTML = criticalItems.length
      ? criticalItems.map(([title, subtitle, value]) => `
        <div class="mission-list-item">
          <div>
            <strong>${title}</strong>
            <small>${subtitle}</small>
          </div>
          <div class="mission-list-value">${value}</div>
        </div>
      `).join("")
      : `<div class="mission-list-empty">Import and calculate project data to surface critical chainages and structural hotspots.</div>`;
  }

  if (corridor) {
    if (calcRows.length < 2) {
      corridor.innerHTML = `<div class="mission-list-empty">Import levels and calculate the project to generate the corridor intelligence strip.</div>`;
    } else {
      const minCh = safeNum(calcRows[0].chainage);
      const maxCh = safeNum(calcRows[calcRows.length - 1].chainage);
      const range = Math.max(maxCh - minCh, 1);
      const bandHtml = [];
      for (let i = 1; i < calcRows.length; i += 1) {
        const prev = calcRows[i - 1];
        const row = calcRows[i];
        const startPct = ((safeNum(prev.chainage) - minCh) / range) * 100;
        const widthPct = Math.max(((safeNum(row.chainage) - safeNum(prev.chainage)) / range) * 100, 0.3);
        if (safeNum(row.bank) > 0.001) {
          bandHtml.push(`<span class="corridor-band fill" style="left:${startPct}%; width:${widthPct}%;" data-tip="Fill zone: Bank ${r3(row.bank)} m @ ${formatDashboardChainage(row.chainage)}"></span>`);
        } else if (safeNum(row.cut) > 0.001) {
          bandHtml.push(`<span class="corridor-band cut" style="left:${startPct}%; width:${widthPct}%;" data-tip="Cut zone: Cut ${r3(row.cut)} m @ ${formatDashboardChainage(row.chainage)}"></span>`);
        }
      }
      const markerItems = [
        ...(state.bridgeRows || []).map((b) => ({ type: "bridge", at: safeNum(b.startChainage) })),
        ...(state.curveRows || []).map((c) => ({ type: "curve", at: safeNum(c.chainage) })),
        ...groupedStations.map((s) => ({ type: "station", at: Number.isFinite(s.csb) ? s.csb : safeNum(s.loopStartCh, NaN) })),
      ].filter((item) => Number.isFinite(item.at));
      const markersHtml = markerItems.map((item) => {
        const left = ((item.at - minCh) / range) * 100;
        const label = item.type === "bridge"
          ? `Bridge @ ${formatDashboardChainage(item.at)}`
          : item.type === "curve"
            ? `Curve @ ${formatDashboardChainage(item.at)}`
            : `Station @ ${formatDashboardChainage(item.at)}`;
        return `<span class="corridor-marker ${item.type}" style="left:${Math.min(100, Math.max(0, left))}%;" data-tip="${label}"></span>`;
      }).join("");
      corridor.innerHTML = `
        <div class="corridor-scale">${bandHtml.join("")}${markersHtml}</div>
        <div class="corridor-legend">
          <span><i style="background:#22c55e"></i> Fill zones</span>
          <span><i style="background:#f43f5e"></i> Cut zones</span>
          <span><i style="background:#3b82f6"></i> Bridges</span>
          <span><i style="background:#eab308"></i> Curves</span>
          <span><i style="background:#06b6d4"></i> Stations</span>
        </div>
        <div class="mission-chip-grid">
          <div class="mission-chip">
            <div class="mission-chip-label">Corridor Start</div>
            <div class="mission-chip-value">${formatDashboardChainage(minCh)}</div>
            <div class="mission-chip-sub">Start of calculated chainage</div>
          </div>
          <div class="mission-chip">
            <div class="mission-chip-label">Corridor End</div>
            <div class="mission-chip-value">${formatDashboardChainage(maxCh)}</div>
            <div class="mission-chip-sub">End of calculated chainage</div>
          </div>
        </div>
        <div class="corridor-tooltip" id="corridorTooltip"></div>
      `;
      const tip = corridor.querySelector("#corridorTooltip");
      const scale = corridor.querySelector(".corridor-scale");
      if (tip && scale) {
        scale.onmousemove = (e) => {
          const target = e.target.closest("[data-tip]");
          if (!target) {
            tip.style.opacity = "0";
            return;
          }
          tip.textContent = target.getAttribute("data-tip") || "";
          const rect = corridor.getBoundingClientRect();
          tip.style.left = `${e.clientX - rect.left + 12}px`;
          tip.style.top = `${e.clientY - rect.top + 12}px`;
          tip.style.opacity = "1";
        };
        scale.onmouseleave = () => {
          tip.style.opacity = "0";
        };
      }
    }
  }

  if (balanceGrid) {
    const ratioDen = fillTotal + cutTotal;
    const fillRatio = ratioDen > 0 ? Math.round((fillTotal / ratioDen) * 100) : 0;
    balanceGrid.innerHTML = [
      ["Total Fill", `${formatCompactVolume(fillTotal)} m³`, "Required embankment volume"],
      ["Total Cut", `${formatCompactVolume(cutTotal)} m³`, "Excavated volume"],
      ["Reusable Spoil", `${formatCompactVolume(reusableSpoil)} m³`, `${safeNum(reusablePct, 60).toFixed(0)}% of cut volume`],
      ["Net Balance", `${netBalance >= 0 ? "+" : ""}${formatCompactVolume(netBalance)} m³`, `${fillRatio}% of total volume is fill`],
    ].map(([label, value, sub]) => `
      <div class="balance-card">
        <small>${label}</small>
        <strong>${value}</strong>
        <small>${sub}</small>
      </div>
    `).join("");
  }

  if (alertsList) {
    const alerts = [];
    if (!active) alerts.push(["No active project", "Create or open a project to begin the calculation workflow.", "Action required"]);
    if (active && !uploads.levels) alerts.push(["Levels missing", "Import levels to unlock calculation, graphs, and corridor intelligence.", "Required"]);
    if (active && !uploads.curves) alerts.push(["Curve list missing", "Curve widening and curve intelligence are unavailable.", "Required"]);
    if (active && !uploads.bridges) alerts.push(["Bridge list missing", "Bridge deductions and structure highlights are incomplete.", "Required"]);
    if (active && !uploads.loops) alerts.push(["Loops & Platforms missing", "Station yard widths and platform offsets are not applied.", "Required"]);
    if (active && !uploads.kml) alerts.push(["KML/KMZ missing", "Map alignment and geographic overlays are unavailable.", "Optional"]);
    if (active && uploads.loops && stationPlanCount === 0) alerts.push(["No station plans attached", "Station conceptual plans can be attached from the Alignment Map popups.", "Optional"]);
    if (active && !state.project.verified) alerts.push(["Project not verified", "Run Verify before export to confirm structural calculations.", "Review"]);
    const medianGap = getMedianInterval(calcRows);
    if (Number.isFinite(medianGap)) {
      const suspicious = [];
      for (let i = 1; i < calcRows.length; i += 1) {
        const gap = safeNum(calcRows[i].chainage) - safeNum(calcRows[i - 1].chainage);
        if (gap > medianGap * 1.75) suspicious.push(`${formatDashboardChainage(calcRows[i - 1].chainage)}–${formatDashboardChainage(calcRows[i].chainage)}`);
      }
      if (suspicious.length) {
        alerts.push(["Suspicious chainage gaps", suspicious.slice(0, 3).join(", "), `${suspicious.length} segment${suspicious.length === 1 ? "" : "s"}`]);
      }
    }
    const badCsb = groupedStations.filter((station) => !Number.isFinite(station.csb)).length;
    if (badCsb > 0) alerts.push(["Stations missing CSB", `${badCsb} station${badCsb === 1 ? "" : "s"} do not have a usable CSB chainage.`, "Check loop import"]);
    const qualityIssues = runQualityChecks();
    if (qualityIssues.length) {
      alerts.push(["Quality checks", `${qualityIssues.length} issue${qualityIssues.length === 1 ? "" : "s"} detected.`, "Review"]);
    }

    alertsList.innerHTML = alerts.length
      ? alerts.map(([title, text, value]) => `
        <div class="mission-list-item">
          <div>
            <strong>${title}</strong>
            <small>${text}</small>
          </div>
          <div class="mission-list-value">${value}</div>
        </div>
      `).join("")
      : `<div class="mission-list-empty">No critical design alerts. The current dataset is complete enough for review and export.</div>`;
  }
}

function parseImportedRows(aoa, startCh, interval) {
  const header = detectImportHeaderRow(aoa);
  if (!header) {
    return { rows: [], error: "Ground Level and Proposed/Form. Level columns were not found in the import file." };
  }

  const rows = [];
  const cols = header.cols;
  let rowNo = 0;
  for (let i = header.rowIndex + 1; i < aoa.length; i += 1) {
    const row = Array.isArray(aoa[i]) ? aoa[i] : [];
    const groundLevel = parseLooseNumber(row[cols.groundLevel], NaN);
    const proposedLevel = parseLooseNumber(row[cols.proposedLevel], NaN);
    if (!Number.isFinite(groundLevel) || !Number.isFinite(proposedLevel)) continue;

    let chainage = NaN;
    if (cols.chainage >= 0) {
      chainage = parseChainage(row[cols.chainage]);
    }
    if (!Number.isFinite(chainage)) {
      chainage = startCh + (rowNo * interval);
    }
    if (!Number.isFinite(chainage)) continue;

    rows.push({
      station: cols.station >= 0 ? String(row[cols.station] || "") : "",
      structureNo: cols.structureNo >= 0 ? String(row[cols.structureNo] || "") : "",
      chainage,
      groundLevel,
      proposedLevel,
    });
    rowNo += 1;
  }

  return { rows, headerRow: header.rowIndex, cols };
}

function resolveBridgeColumns(headerCells) {
  return {
    bridgeNo: findColByAliases(headerCells, ["bridgeno", "bridgenumber", "tunnelno", "tunnelname", "structureno", "slno", "sl.no", "serialno", "sno", "sn", "structure", "bridge", "tunnel"]),
    startChainage: findColByAliases(headerCells, ["bridgestartchainage", "startchainage", "fromchainage", "chfrom", "startch", "fromkm"]),
    endChainage: findColByAliases(headerCells, ["bridgeendchainage", "endchainage", "tochainage", "chto", "endch", "tokm"]),
    length: findColByAliases(headerCells, ["totalspanlength", "bridgelength", "length", "spanm", "totallength", "len"]),
    chainage: findColByAliases(headerCells, ["chainage", "ch", "centerchainage", "location", "km"]),
    category: findColByAliases(headerCells, ["category", "purpose", "bridgecategory", "typeofobstruction", "typeofcrossing"]),
    type: findColByAliases(headerCells, ["bridgetype", "structuretype", "typeofspan", "typeofsuperstructure", "superstructure", "purpose"]),
    size: findColByAliases(headerCells, ["bridgesize", "size", "dimensions", "proposedspan", "proposedspanarrangement"]),
    spans: findColByAliases(headerCells, ["noofspans", "numberofspans", "spanarrangement", "arrangement"]),
    spanLength: findColByAliases(headerCells, ["spanlength", "individualspan", "unitsize", "spanarrangement", "arrangement"]),
    clearSpan: findColByAliases(headerCells, ["clearspan", "clear span", "clearwidth", "clear width"]),
    deductRule: findColByAliases(headerCells, ["deductrule", "deductionrule", "deduct", "deduction"]),
    autoDeduct: findColByAliases(headerCells, ["autodeduct", "auto deduct", "deductflag"]),
  };
}

// Parse size strings like "10x9.15", "1x7.5x5.5" to extract span count and individual span length
function parseBridgeSpanSize(sizeStr) {
  if (!sizeStr) return { spans: NaN, spanLength: NaN, totalLength: NaN };
  const s = String(sizeStr).trim();
  // Pattern: NxL or NxLxH (e.g. 10x9.15 or 1x7.5x5.5)
  const m = s.match(/^(\d+)\s*[xX×]\s*([\d.]+)/);
  if (!m) return { spans: NaN, spanLength: NaN, totalLength: NaN };
  const spans = parseInt(m[1], 10);
  const spanLength = parseFloat(m[2]);
  if (!Number.isFinite(spans) || !Number.isFinite(spanLength)) return { spans: NaN, spanLength: NaN, totalLength: NaN };
  return { spans, spanLength, totalLength: spans * spanLength };
}

function detectBridgeHeaderRow(aoa) {
  const scanRows = Math.min(aoa.length, 40);
  let best = { rowIndex: -1, score: -1, cols: null };
  for (let i = 0; i < scanRows; i += 1) {
    const row = Array.isArray(aoa[i]) ? aoa[i] : [];
    if (!row.length) continue;
    const cols = resolveBridgeColumns(row);
    const score = (cols.bridgeNo >= 0 ? 2 : 0)
      + (cols.startChainage >= 0 ? 3 : 0)
      + (cols.endChainage >= 0 ? 3 : 0)
      + (cols.length >= 0 ? 2 : 0)
      + (cols.chainage >= 0 ? 1 : 0)
      + (cols.size >= 0 ? 1 : 0);
    if (score > best.score) {
      best = { rowIndex: i, score, cols };
    }
  }
  if (best.rowIndex < 0) return null;
  const c = best.cols;
  const hasCore = (c.startChainage >= 0 && c.endChainage >= 0)
    || (c.chainage >= 0 && c.length >= 0)
    || (c.startChainage >= 0 && c.length >= 0)
    || (c.endChainage >= 0 && c.length >= 0)
    || (c.chainage >= 0 && c.size >= 0);  // chainage + size (length derived from span size)
  return hasCore ? best : null;
}

function normalizeBridgeEntry(raw, index = 0) {
  const bridgeNo = String(raw.bridgeNo || `BR-${index + 1}`).trim();
  let start = parseChainage(raw.startChainage);
  let end = parseChainage(raw.endChainage);
  let lengthRaw = parseLooseNumber(raw.length, NaN);
  const center = parseChainage(raw.chainage);
  const clearSpanRaw = parseLooseNumber(raw.clearSpan, NaN);
  const deductRule = normalizeDeductRule(raw.deductRule);
  const autoDeduct = normalizeBoolean(raw.autoDeduct, true);

  // Derive parameters from size string if possible
  const sizeParsed = parseBridgeSpanSize(raw.bridgeSize);

  // Spans
  let spansVal = String(raw.bridgeSpans || "");
  if (!spansVal || spansVal === "1") {
    if (Number.isFinite(sizeParsed.spans)) spansVal = String(sizeParsed.spans);
  }

  // Individual Span Length
  let spanLen = parseLooseNumber(raw.bridgeSpanLength, NaN);
  if (!Number.isFinite(spanLen)) {
    if (Number.isFinite(sizeParsed.spanLength)) spanLen = sizeParsed.spanLength;
  }

  // Total Length
  if (!Number.isFinite(lengthRaw)) {
    if (Number.isFinite(spanLen) && Number.isFinite(parseInt(spansVal, 10))) {
      lengthRaw = spanLen * parseInt(spansVal, 10);
    } else if (Number.isFinite(sizeParsed.totalLength)) {
      lengthRaw = sizeParsed.totalLength;
    }
  }

  if (!Number.isFinite(start) && Number.isFinite(end) && Number.isFinite(lengthRaw)) start = end - lengthRaw;
  if (!Number.isFinite(end) && Number.isFinite(start) && Number.isFinite(lengthRaw)) end = start + lengthRaw;
  if ((!Number.isFinite(start) || !Number.isFinite(end)) && Number.isFinite(center) && Number.isFinite(lengthRaw)) {
    start = center - (lengthRaw / 2);
    end = center + (lengthRaw / 2);
  }
  if (!Number.isFinite(start) || !Number.isFinite(end)) return null;
  if (end < start) [start, end] = [end, start];
  if (end <= start) return null;

  const length = Number.isFinite(lengthRaw) && lengthRaw > 0 ? lengthRaw : (end - start);

  // Derive Category/Type if missing or generic
  let cat = String(raw.bridgeCategory || "").trim();
  if (!cat || cat === "-" || cat === "Minor" || cat === "undefined") {
    const dCat = detectCategoryFromText(bridgeNo) || detectCategoryFromText(raw.bridgeType) || detectCategoryFromText(raw.bridgeSize);
    if (dCat) cat = dCat;
  }
  if (!cat || cat === "undefined") cat = "Minor";

  let bType = String(raw.bridgeType || "").trim();
  if (!bType || bType === "-" || bType === "Box" || bType === "undefined") {
    const dType = detectBridgeTypeFromText(cat) || detectBridgeTypeFromText(bridgeNo) || detectBridgeTypeFromText(raw.bridgeSize);
    if (dType) bType = dType;
  }
  if (!bType && cat === "Tunnel") bType = "Box";
  if (!bType || bType === "undefined") bType = "Box";

  const DEDUCTIBLE_CATEGORIES = ["Major", "Important", "RoR", "Viaduct", "Tunnel"];
  let shouldDeduct = false;
  if (deductRule === "Always") {
    shouldDeduct = true;
  } else if (deductRule === "Never") {
    shouldDeduct = false;
  } else {
    shouldDeduct = autoDeduct && DEDUCTIBLE_CATEGORIES.some(c => cat.toLowerCase() === c.toLowerCase());
  }

  const clearSpan = Number.isFinite(clearSpanRaw)
    ? clearSpanRaw
    : (Number.isFinite(spanLen) ? spanLen : "");

  return {
    bridgeNo,
    bridgeCategory: cat,
    bridgeType: bType,
    bridgeSize: String(raw.bridgeSize || "-"),
    bridgeSpans: spansVal || "1",
    bridgeSpanLength: Number.isFinite(spanLen) ? spanLen : "",
    clearSpan,
    deductRule,
    autoDeduct,
    startChainage: start,
    endChainage: end,
    length,
    shouldDeduct,
  };
}

function detectCategoryFromText(text) {
  if (!text) return null;
  const s = String(text).toLowerCase();
  if (s.includes("minor")) return "Minor";
  if (s.includes("major")) return "Major";
  if (s.includes("viaduct")) return "Viaduct";
  if (s.includes("tunnel") || s.includes("tinnel")) return "Tunnel";
  if (s.includes("important")) return "Important";
  if (s.includes("ror") || s.includes("road over rail")) return "RoR";
  if (s.includes("rob") || s.includes("road over bridge")) return "ROB";
  if (s.includes("mibor") || s.includes("missing bridge")) return "MIBOR";
  if (s.includes("aqueduct") || s.includes("aquiduct")) return "Aqueduct";
  return null;
}

function detectBridgeTypeFromText(text) {
  if (!text) return null;
  const s = String(text).toLowerCase();
  if (s.includes("box")) return "Box";
  if (s.includes("psc") || s.includes("slab")) return "PSC Slab";
  if (s.includes("composite") || s.includes("girder")) return "Composite Girder";
  if (s.includes("owg") || s.includes("open web")) return "OWG";
  return null;
}

function normalizeDeductRule(rule) {
  const s = String(rule || "").toLowerCase();
  if (!s) return "Auto";
  if (s.includes("always") || s === "yes" || s === "y") return "Always";
  if (s.includes("never") || s === "no" || s === "n") return "Never";
  return "Auto";
}

function normalizeBoolean(val, fallback = true) {
  if (val === null || val === undefined || val === "") return fallback;
  const s = String(val).toLowerCase().trim();
  if (["false", "no", "n", "0"].includes(s)) return false;
  if (["true", "yes", "y", "1"].includes(s)) return true;
  return fallback;
}

function parseBridgeRowsFromAoa(aoa, sheetName = '') {
  const header = detectBridgeHeaderRow(aoa);
  if (!header) return { rows: [], error: "Bridge sheet must include Start/End chainage or Chainage+Length columns." };

  // Clue from sheet name or header rows above
  let inferredCategory = detectCategoryFromText(sheetName);
  if (!inferredCategory) {
    // Scan rows above the header for categories (often in sheet title rows)
    for (let j = 0; j < header.rowIndex; j++) {
      const rowText = (aoa[j] || []).join(" ").trim();
      if (!rowText) continue;
      const found = detectCategoryFromText(rowText);
      if (found) { inferredCategory = found; break; }
    }
  }
  const defaultCategory = inferredCategory || "Minor";

  const rows = [];
  const cols = header.cols;
  for (let i = header.rowIndex + 1; i < aoa.length; i += 1) {
    const row = Array.isArray(aoa[i]) ? aoa[i] : [];

    // Read category from cell
    let cellCategory = cols.category >= 0 ? String(row[cols.category] || "").trim() : "";
    let cellType = cols.type >= 0 ? String(row[cols.type] || "").trim() : "";

    // Auto-detect category if possible
    const purpose = cols.category >= 0 ? String(row[cols.category] || "").toLowerCase() : "";
    const typeStr = cols.type >= 0 ? String(row[cols.type] || "").toLowerCase() : "";
    const bridgeNoStr = cols.bridgeNo >= 0 ? String(row[cols.bridgeNo] || "").toLowerCase() : "";

    if (!cellCategory) {
      cellCategory = detectCategoryFromText(purpose) || detectCategoryFromText(typeStr) || detectCategoryFromText(bridgeNoStr);
    }

    if (!cellCategory) cellCategory = defaultCategory;

    let detectedType = detectBridgeTypeFromText(cellType) || detectBridgeTypeFromText(purpose) || detectBridgeTypeFromText(bridgeNoStr);
    if (!detectedType) {
      if (cellCategory === "Tunnel") detectedType = "Box"; // Default for tunnel
    }
    const finalType = detectedType || cellType || "Box";

    const raw = {
      bridgeNo: cols.bridgeNo >= 0 ? row[cols.bridgeNo] : "",
      bridgeCategory: cellCategory,
      bridgeType: finalType,
      bridgeSize: cols.size >= 0 ? (String(row[cols.size] || "").trim() || "-") : "-",
      bridgeSpans: cols.spans >= 0 ? (String(row[cols.spans] || "").trim() || "1") : "1",
      bridgeSpanLength: cols.spanLength >= 0 ? row[cols.spanLength] : "",
      clearSpan: cols.clearSpan >= 0 ? row[cols.clearSpan] : "",
      deductRule: cols.deductRule >= 0 ? row[cols.deductRule] : "",
      autoDeduct: cols.autoDeduct >= 0 ? row[cols.autoDeduct] : "",
      startChainage: cols.startChainage >= 0 ? row[cols.startChainage] : "",
      endChainage: cols.endChainage >= 0 ? row[cols.endChainage] : "",
      length: cols.length >= 0 ? row[cols.length] : "",
      chainage: cols.chainage >= 0 ? row[cols.chainage] : "",
    };
    const normalized = normalizeBridgeEntry(raw, rows.length);
    if (normalized) rows.push(normalized);
  }
  return { rows };
}

function detectSimpleHeaderRow(aoa, requiredAliases = []) {
  const scanRows = Math.min(aoa.length, 30);
  let best = { rowIndex: -1, score: -1, header: [] };
  for (let i = 0; i < scanRows; i += 1) {
    const row = Array.isArray(aoa[i]) ? aoa[i] : [];
    if (!row.length) continue;
    const norm = row.map((c) => normalizeHeaderToken(c));
    const score = requiredAliases.reduce((acc, alias) => (
      acc + (norm.some((h) => h.includes(alias)) ? 1 : 0)
    ), 0);
    if (score > best.score) best = { rowIndex: i, score, header: row };
  }
  if (best.rowIndex < 0) return null;
  return best;
}

const IMPORT_MAPPING_FIELDS = {
  levels: [
    { key: "chainage", label: "Chainage" },
    { key: "groundLevel", label: "Ground RL" },
    { key: "proposedLevel", label: "Proposed RL" },
    { key: "station", label: "Station" },
    { key: "structureNo", label: "Structure No" },
  ],
  curves: [
    { key: "chainage", label: "Chainage" },
    { key: "curve", label: "Curve Name" },
    { key: "radius", label: "Radius" },
    { key: "length", label: "Length" },
  ],
  bridges: [
    { key: "bridgeNo", label: "Bridge No" },
    { key: "bridgeCategory", label: "Category" },
    { key: "bridgeType", label: "Type" },
    { key: "bridgeSize", label: "Size" },
    { key: "bridgeSpans", label: "Spans" },
    { key: "bridgeSpanLength", label: "Span Length" },
    { key: "clearSpan", label: "Clear Span" },
    { key: "deductRule", label: "Deduct Rule" },
    { key: "autoDeduct", label: "Auto Deduct" },
    { key: "startChainage", label: "Start Chainage" },
    { key: "endChainage", label: "End Chainage" },
    { key: "length", label: "Length" },
    { key: "chainage", label: "Chainage (Center)" },
  ],
  loops: [
    { key: "station", label: "Station" },
    { key: "csb", label: "CSB" },
    { key: "lineType", label: "Type" },
    { key: "side", label: "Side" },
    { key: "loopStartCh", label: "Loop Start Ch" },
    { key: "loopEndCh", label: "Loop End Ch" },
    { key: "tc", label: "T/C" },
    { key: "pfStartCh", label: "PF Start Ch" },
    { key: "pfEndCh", label: "PF End Ch" },
    { key: "pfWidth", label: "Width" },
    { key: "remarks", label: "Remarks" },
  ],
};

const MAPPING_CANON_HEADERS = {
  levels: ["Chainage", "Ground RL", "Proposed RL", "Station", "Structure No"],
  curves: ["Chainage", "Curve", "Radius", "Length"],
  bridges: ["Bridge No", "Category", "Type", "Size", "Spans", "Span Length", "Clear Span", "Deduct Rule", "Auto Deduct", "Start Chainage", "End Chainage", "Length", "Chainage"],
  loops: ["Stations", "CSB", "Type", "Side", "Loop Start Ch", "Loop End Ch", "T/C", "PF Start Ch", "PF End Ch", "Width", "Remarks"],
};

function buildMappedAoa(aoa, mapping, headerRow, kind) {
  if (!aoa || !aoa.length || !mapping) return aoa;
  const fields = IMPORT_MAPPING_FIELDS[kind] || [];
  const headers = MAPPING_CANON_HEADERS[kind] || [];
  const dataStart = Math.max(0, (Number.isFinite(headerRow) ? headerRow + 1 : 1));
  const newAoa = [headers];
  for (let i = dataStart; i < aoa.length; i += 1) {
    const row = Array.isArray(aoa[i]) ? aoa[i] : [];
    const newRow = new Array(headers.length).fill("");
    fields.forEach((f, idx) => {
      const colIdx = mapping[f.key];
      if (Number.isFinite(colIdx) && colIdx >= 0) {
        newRow[idx] = row[colIdx];
      }
    });
    newAoa.push(newRow);
  }
  return newAoa;
}

function getMappingTemplate(client, kind) {
  if (!client || !kind) return null;
  const templates = state.importMappings?.templates || {};
  return templates[client]?.[kind] || null;
}

function applyMappingTemplateToAoa(aoa, kind) {
  const client = String(state.importMappings?.client || "").trim();
  if (!client) return aoa;
  const preset = getMappingTemplate(client, kind);
  if (!preset) return aoa;
  const headerPick = detectSimpleHeaderRow(aoa, []);
  const headerRow = headerPick ? headerPick.rowIndex : 0;
  return buildMappedAoa(aoa, preset, headerRow, kind);
}

function renderMappingGrid(headers, kind, preset = null) {
  if (!els.mappingGrid) return;
  const fields = IMPORT_MAPPING_FIELDS[kind] || [];
  if (!headers || !headers.length) {
    els.mappingGrid.innerHTML = '<p class="muted">Upload a file to detect headers.</p>';
    return;
  }
  const options = headers.map((h, idx) => `<option value="${idx}">${escapeHtml(String(h))}</option>`).join("");
  els.mappingGrid.innerHTML = fields.map((f) => {
    const selectedIdx = preset && Number.isFinite(preset[f.key]) ? preset[f.key] : "";
    return `
      <label>${f.label}
        <select data-map-key="${f.key}">
          <option value="">-- Not Mapped --</option>
          ${options}
        </select>
      </label>
    `;
  }).join("");
  if (preset) {
    fields.forEach((f) => {
      const sel = els.mappingGrid.querySelector(`[data-map-key="${f.key}"]`);
      if (sel && Number.isFinite(preset[f.key])) {
        sel.value = String(preset[f.key]);
      }
    });
  }
}

function getStationChoices() {
  return [...new Set((state.loopPlatformRows || []).map((r) => String(r.station || "").trim()).filter(Boolean))];
}

function buildStationLayoutList(stationName) {
  if (!els.stationLayoutList) return;
  const key = normalizeStationKey(stationName);
  const rows = (state.loopPlatformRows || [])
    .map((row, index) => ({ row, index }))
    .filter((r) => normalizeStationKey(r.row.station) === key)
    .sort((a, b) => (Number.isFinite(a.row.order) ? a.row.order : a.index) - (Number.isFinite(b.row.order) ? b.row.order : b.index));
  els.stationLayoutList.innerHTML = rows.map((entry) => {
    const r = entry.row;
    const isPlatform = /platform/i.test(String(r.lineType || r.lineName || "")) || safeNum(r.pfWidth, 0) > 0;
    return `
      <div class="layout-item" draggable="true" data-row-index="${entry.index}">
        <div class="layout-meta">
          <strong>${escapeHtml(r.lineType || r.lineName || "Line")}</strong>
          <span>${escapeHtml(r.side || "-")} • TC ${r3(r.tc || 0)} • PF ${r3(r.pfWidth || 0)}</span>
          <span class="muted">${escapeHtml(r.remarks || "")}</span>
          <div class="layout-fields">
            <label>Side
              <select data-layout-side>
                ${["", ...LOOP_SIDES].map((side) => `<option value="${side}" ${String(r.side || "") === side ? "selected" : ""}>${side || "Auto"}</option>`).join("")}
              </select>
            </label>
            <label>Track Centre (m)
              <input data-layout-tc type="number" step="0.01" value="${r3(r.tc || 0)}" ${isPlatform ? "disabled" : ""} />
            </label>
            <label>Platform Width (m)
              <input data-layout-pf type="number" step="0.1" value="${r3(r.pfWidth || 0)}" ${isPlatform ? "" : "disabled"} />
            </label>
          </div>
        </div>
        <div class="layout-actions">
          <button type="button" class="btn btn-secondary" data-layout-up>↑</button>
          <button type="button" class="btn btn-secondary" data-layout-down>↓</button>
        </div>
      </div>
    `;
  }).join("") || '<p class="muted">No rows for this station.</p>';
}

function parseCurveRowsFromAoa(aoa) {
  const headerPick = detectSimpleHeaderRow(aoa, ["chainage", "curve", "radius", "length"]);
  const headerRow = headerPick ? headerPick.rowIndex : 0;
  const header = Array.isArray(aoa[headerRow]) ? aoa[headerRow] : [];
  const norm = header.map((h) => normalizeHeaderToken(h));
  const cChainage = norm.findIndex((h) => h.includes("chainage") || h === "ch");
  const cCurve = norm.findIndex((h) => h.includes("curve"));
  const cRadius = norm.findIndex((h) => h.includes("radius") || h === "r");
  const cLength = norm.findIndex((h) => h.includes("length") || h.includes("len"));
  const rows = [];
  for (let i = headerRow + 1; i < aoa.length; i += 1) {
    const row = Array.isArray(aoa[i]) ? aoa[i] : [];
    if (!row.length) continue;
    const chainage = cChainage >= 0 ? parseChainage(row[cChainage]) : NaN;
    const curve = cCurve >= 0 ? String(row[cCurve] || "").trim() : "";
    const radius = cRadius >= 0 ? parseLooseNumber(row[cRadius], NaN) : NaN;
    const length = cLength >= 0 ? parseLooseNumber(row[cLength], NaN) : NaN;
    if (!curve && !Number.isFinite(chainage) && !Number.isFinite(radius) && !Number.isFinite(length)) continue;
    rows.push({
      chainage: Number.isFinite(chainage) ? chainage : null,
      curve: curve || `Curve-${rows.length + 1}`,
      radius: Number.isFinite(radius) ? radius : null,
      length: Number.isFinite(length) ? length : null,
    });
  }
  return { rows };
}

function parseLoopPlatformRowsFromAoa(aoa) {
  const headerPick = detectSimpleHeaderRow(aoa, ["stations", "csb", "loopstart", "loopend", "pfstart", "pfend", "width"]);
  const headerRow = headerPick ? headerPick.rowIndex : 0;
  const header = Array.isArray(aoa[headerRow]) ? aoa[headerRow] : [];
  const norm = header.map((h) => normalizeHeaderToken(h));
  const cStation = norm.findIndex((h) => h === "stations" || h.includes("station"));
  const cLineName = findColByAliases(header, ["linetype", "line type", "linename", "type", "line"]);
  const cSide = findColByAliases(header, ["side", "lhside", "rhside", "lhs", "rhs"]);
  const cCsb = norm.findIndex((h) => h === "csb");
  const cLoopStart = norm.findIndex((h) => h.includes("loopstartch") || h.includes("loopstartchain"));
  const cLoopEnd = norm.findIndex((h) => h.includes("loopendch") || h.includes("loopendchain"));
  const cTc = norm.findIndex((h) => h === "tc" || h.includes("trackcentre") || h.includes("trackcenter"));
  const cPfStart = norm.findIndex((h) => h.includes("pfstartch") || h.includes("pfstartchain"));
  const cPfEnd = norm.findIndex((h) => h.includes("pfendch") || h.includes("pfendchain"));
  const cWidth = norm.findIndex((h) => h === "width" || h.includes("platformwidth"));
  const cRemarks = norm.findIndex((h) => h.includes("remark"));

  const findOffsetCol = (baseToken) => (
    norm.findIndex((h) => h.includes(baseToken) && !h.includes("ch") && !h.includes("chain"))
  );
  const cLoopStartOff = findOffsetCol("loopstart");
  const cLoopEndOff = findOffsetCol("loopend");
  const cPfStartOff = findOffsetCol("pfstart");
  const cPfEndOff = findOffsetCol("pfend");

  // Fallback to fixed template columns (legacy workbook layout).
  // Prefer header matches above; these are only used when headers are missing.
  const idx = {
    station: cStation >= 0 ? cStation : 1,
    csb: cCsb >= 0 ? cCsb : 2,
    tc: cTc >= 0 ? cTc : 5,
    loopStartCh: cLoopStart >= 0 ? cLoopStart : 6,
    loopEndCh: cLoopEnd >= 0 ? cLoopEnd : 7,
    pfStartCh: cPfStart >= 0 ? cPfStart : 11,
    pfEndCh: cPfEnd >= 0 ? cPfEnd : 12,
    width: cWidth >= 0 ? cWidth : 10,
    remarks: cRemarks >= 0 ? cRemarks : 13,
    loopStartOff: cLoopStartOff >= 0 ? cLoopStartOff : 3,
    loopEndOff: cLoopEndOff >= 0 ? cLoopEndOff : 4,
    pfStartOff: cPfStartOff >= 0 ? cPfStartOff : 8,
    pfEndOff: cPfEndOff >= 0 ? cPfEndOff : 9,
  };

  const rows = [];
  let currentStation = "";
  let currentCsb = NaN;
  for (let i = headerRow + 1; i < aoa.length; i += 1) {
    const row = Array.isArray(aoa[i]) ? aoa[i] : [];
    if (!row.length) continue;

    const stationRaw = String(row[idx.station] ?? "").trim();
    const rawLineType = cLineName >= 0 ? String(row[cLineName] ?? "").trim() : "";
    const lineType = normalizeLoopLineType(rawLineType);
    const side = normalizeLoopSide(cSide >= 0 ? row[cSide] : "");
    const isExampleRow = /example|target chainage|add here/i.test(stationRaw);
    if (stationRaw && !isExampleRow) currentStation = stationRaw;

    const csbCandidate = parseChainage(row[idx.csb]);
    if (Number.isFinite(csbCandidate)) currentCsb = csbCandidate;

    let loopStartCh = parseChainage(row[idx.loopStartCh]);
    let loopEndCh = parseChainage(row[idx.loopEndCh]);
    let pfStartCh = parseChainage(row[idx.pfStartCh]);
    let pfEndCh = parseChainage(row[idx.pfEndCh]);

    // If absolute chainage columns are unavailable, derive using CSB + offsets.
    if (!Number.isFinite(loopStartCh) && Number.isFinite(currentCsb)) {
      const off = parseLooseNumber(row[idx.loopStartOff], NaN);
      if (Number.isFinite(off)) loopStartCh = currentCsb + off;
    }
    if (!Number.isFinite(loopEndCh) && Number.isFinite(currentCsb)) {
      const off = parseLooseNumber(row[idx.loopEndOff], NaN);
      if (Number.isFinite(off)) loopEndCh = currentCsb + off;
    }
    if (!Number.isFinite(pfStartCh) && Number.isFinite(currentCsb)) {
      const off = parseLooseNumber(row[idx.pfStartOff], NaN);
      if (Number.isFinite(off)) pfStartCh = currentCsb + off;
    }
    if (!Number.isFinite(pfEndCh) && Number.isFinite(currentCsb)) {
      const off = parseLooseNumber(row[idx.pfEndOff], NaN);
      if (Number.isFinite(off)) pfEndCh = currentCsb + off;
    }

    if (Number.isFinite(loopStartCh) && Number.isFinite(loopEndCh) && loopEndCh < loopStartCh) {
      [loopStartCh, loopEndCh] = [loopEndCh, loopStartCh];
    }
    if (Number.isFinite(pfStartCh) && Number.isFinite(pfEndCh) && pfEndCh < pfStartCh) {
      [pfStartCh, pfEndCh] = [pfEndCh, pfStartCh];
    }

    const tc = parseLooseNumber(row[idx.tc], NaN);
    const width = parseLooseNumber(row[idx.width], NaN);
    const remarks = String(row[idx.remarks] ?? "").trim();

    const hasLoop = Number.isFinite(loopStartCh) && Number.isFinite(loopEndCh) && loopEndCh > loopStartCh && Number.isFinite(tc) && tc > 0;
    const hasPf = Number.isFinite(pfStartCh) && Number.isFinite(pfEndCh) && pfEndCh > pfStartCh && Number.isFinite(width) && width > 0;
    const hasLine = Boolean(lineType) || Boolean(side);
    const hasStationIdentity = (Boolean(stationRaw) && !isExampleRow) || Number.isFinite(csbCandidate) || Boolean(remarks);
    if (!hasLoop && !hasPf && !hasStationIdentity && !hasLine) continue;
    if (isExampleRow && !hasLoop && !hasPf) continue;

    rows.push({
      station: currentStation || `LP-${rows.length + 1}`,
      lineType,
      lineName: rawLineType,
      side,
      csb: Number.isFinite(currentCsb) ? currentCsb : null,
      tc: Number.isFinite(tc) ? tc : 0,
      loopStartCh: Number.isFinite(loopStartCh) ? loopStartCh : null,
      loopEndCh: Number.isFinite(loopEndCh) ? loopEndCh : null,
      pfWidth: Number.isFinite(width) ? width : 0,
      pfStartCh: Number.isFinite(pfStartCh) ? pfStartCh : null,
      pfEndCh: Number.isFinite(pfEndCh) ? pfEndCh : null,
      remarks,
      order: rows.length,
      _stationKey: String(currentStation || `LP-${rows.length + 1}`).toLowerCase(),
      _sideExplicit: Boolean(side),
    });
  }
  // Infer side for rows missing side using relative position to Main Line in each station block.
  if (rows.length) {
    const groups = new Map();
    rows.forEach((r, idx) => {
      const key = r._stationKey || String(r.station || "").toLowerCase();
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push({ row: r, idx });
    });
    groups.forEach((entries) => {
      const mainLines = entries.filter((e) => e.row.lineType === "Main Line");
      if (!mainLines.length) return;
      entries.forEach((entry) => {
        const r = entry.row;
        if (r._sideExplicit) return;
        const lineLabel = String(r.lineName || "").toLowerCase();
        const hasUp = /\bup\b/.test(lineLabel);
        const hasDn = /\bdn\b|\bdown\b/.test(lineLabel);
        if (hasUp && !hasDn) {
          r.side = "Left";
          return;
        }
        if (hasDn && !hasUp) {
          r.side = "Right";
          return;
        }
        if (!r.lineType || r.lineType === "Main Line") return;
        // If multiple main lines exist, use nearest main line row as divider.
        let nearest = mainLines[0];
        let bestDist = Math.abs(entry.idx - nearest.idx);
        for (const m of mainLines) {
          const dist = Math.abs(entry.idx - m.idx);
          if (dist < bestDist) {
            bestDist = dist;
            nearest = m;
          }
        }
        r.side = entry.idx < nearest.idx ? "Left" : "Right";
      });
    });
  }
  // Clean internal fields
  rows.forEach((r) => {
    delete r._stationKey;
    delete r._sideExplicit;
  });
  return { rows };
}

function getLoopPlatformAtChainage(chainage) {
  if (!Number.isFinite(chainage) || !state.loopPlatformRows.length) {
    return { loopTc: 0, loopCount: 0, platformWidth: 0 };
  }
  let loopTc = 0;
  let loopCount = 0;
  let platformWidth = 0;
  for (const r of state.loopPlatformRows) {
    if (Number.isFinite(r.loopStartCh) && Number.isFinite(r.loopEndCh) && chainage >= r.loopStartCh && chainage <= r.loopEndCh) {
      const tc = safeNum(r.tc, 0);
      loopTc += tc;
      if (tc > 0) {
        loopCount += Math.max(1, Math.round(tc / 5.3));
      }
    }
    if (Number.isFinite(r.pfStartCh) && Number.isFinite(r.pfEndCh) && chainage >= r.pfStartCh && chainage <= r.pfEndCh) {
      platformWidth += safeNum(r.pfWidth, 0);
    }
  }
  return { loopTc, loopCount, platformWidth };
}



async function readSheetAoaFromFile(file) {
  const ext = file.name.toLowerCase();
  if (!(ext.endsWith(".csv") || ext.endsWith(".xlsx") || ext.endsWith(".xls"))) {
    throw new Error("Supported files are .csv, .xlsx, .xls");
  }
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, { type: "array", cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false, blankrows: false });
  if (!aoa.length) throw new Error("File has no data rows.");
  return aoa;
}

// Read ALL sheets from an Excel file, returning [{name, aoa}]
async function readAllSheetsFromFile(file) {
  const ext = file.name.toLowerCase();
  if (!(ext.endsWith(".csv") || ext.endsWith(".xlsx") || ext.endsWith(".xls"))) {
    throw new Error("Supported files are .csv, .xlsx, .xls");
  }
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, { type: "array", cellDates: false });
  return wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false, blankrows: false });
    return { name, aoa };
  });
}

function buildBridgeIntervals() {
  return state.bridgeRows
    .map((b, i) => normalizeBridgeEntry(b, i))
    .filter(Boolean)
    .sort((a, b) => a.startChainage - b.startChainage);
}

function runQualityChecks() {
  const issues = [];
  const minTc = safeNum(state.settings.minTc, 5.3);
  const minPf = safeNum(state.settings.minPlatformWidth, 6.0);
  const minClear = safeNum(state.settings.minClearance, 1.5);

  const stationMap = new Map();
  (state.loopPlatformRows || []).forEach((row) => {
    const key = normalizeStationKey(row.station);
    if (!key) return;
    if (!stationMap.has(key)) stationMap.set(key, { station: row.station || key, rows: [] });
    stationMap.get(key).rows.push(row);
  });

  stationMap.forEach((group) => {
    const rows = group.rows;
    const hasMain = rows.some((r) => /main/i.test(String(r.lineType || r.lineName || "")) || Math.abs(safeNum(r.tc, 0)) < 0.001);
    if (!hasMain) {
      issues.push({ level: "error", title: "Missing Main Line", detail: `${group.station} has no main line entry.` });
    }

    const missingSide = rows.filter((r) => (safeNum(r.tc, 0) > 0 || safeNum(r.pfWidth, 0) > 0) && !String(r.side || "").trim());
    if (missingSide.length) {
      issues.push({ level: "warning", title: "Side Not Assigned", detail: `${group.station} has ${missingSide.length} row(s) without side.` });
    }

    const loopRanges = rows
      .filter((r) => Number.isFinite(r.loopStartCh) && Number.isFinite(r.loopEndCh))
      .map((r) => ({ start: r.loopStartCh, end: r.loopEndCh, lineType: r.lineType || r.lineName || "Loop", side: r.side || "-" }))
      .sort((a, b) => a.start - b.start);

    for (let i = 1; i < loopRanges.length; i += 1) {
      const prev = loopRanges[i - 1];
      const cur = loopRanges[i];
      if (cur.start < prev.end) {
        issues.push({ level: "error", title: "Loop Range Overlap", detail: `${group.station} ${prev.lineType} overlaps ${cur.lineType}.` });
      }
      if (cur.start > prev.end && prev.lineType === cur.lineType) {
        issues.push({ level: "warning", title: "Loop Gap", detail: `${group.station} ${prev.lineType} gap of ${r3(cur.start - prev.end)} m.` });
      }
    }

    const pfRanges = rows
      .filter((r) => Number.isFinite(r.pfStartCh) && Number.isFinite(r.pfEndCh))
      .map((r) => ({ start: r.pfStartCh, end: r.pfEndCh, width: safeNum(r.pfWidth, 0), remarks: r.remarks || "" }))
      .sort((a, b) => a.start - b.start);
    for (let i = 1; i < pfRanges.length; i += 1) {
      const prev = pfRanges[i - 1];
      const cur = pfRanges[i];
      if (cur.start < prev.end) {
        issues.push({ level: "warning", title: "Platform Range Overlap", detail: `${group.station} platform ranges overlap.` });
        break;
      }
    }
    pfRanges.forEach((pf) => {
      const within = loopRanges.some((lp) => pf.start >= lp.start && pf.end <= lp.end);
      if (!within) {
        issues.push({ level: "warning", title: "Platform Outside Loop", detail: `${group.station} platform range ${r3(pf.start)}–${r3(pf.end)} m outside loop span.` });
      }
      if (pf.width > 0 && pf.width < minPf) {
        issues.push({ level: "warning", title: "Platform Width Below Minimum", detail: `${group.station} platform width ${r3(pf.width)} m < ${r3(minPf)} m.` });
      }
    });

    // Clearance checks based on sequence layout
    const stationCh = rows.find((r) => Number.isFinite(r.csb))?.csb;
    const layout = buildStationSequenceLayout(stationCh, group.station, 5.3, { useRanges: false });
    if (layout) {
      layout.platformItems.forEach((pf) => {
        const width = safeNum(pf.row.pfWidth, 0);
        if (!width) return;
        const ordered = layout.orderedItems || [];
        const idx = ordered.indexOf(pf);
        const before = idx >= 0 ? ordered.slice(0, idx).reverse().find((i) => i.kind === "track") : null;
        const after = idx >= 0 ? ordered.slice(idx + 1).find((i) => i.kind === "track") : null;
        if (before && after) {
          const offA = layout.offsetByItem.get(before);
          const offB = layout.offsetByItem.get(after);
          if (Number.isFinite(offA) && Number.isFinite(offB)) {
            const gap = Math.abs(offA - offB);
            if (width + (2 * minClear) > gap) {
              issues.push({ level: "error", title: "Island Clearance Breach", detail: `${group.station} island PF width ${r3(width)} m exceeds clearance between tracks (${r3(gap)} m).` });
            }
          }
        } else {
          const anchor = before || after || layout.refMain;
          const anchorOffset = layout.offsetByItem.get(anchor);
          if (Number.isFinite(anchorOffset)) {
            const gap = Math.abs(anchorOffset);
            if (width + minClear > gap) {
              issues.push({ level: "warning", title: "Side Clearance Breach", detail: `${group.station} platform width ${r3(width)} m too close to track centre (${r3(gap)} m).` });
            }
          }
        }
      });
    }
  });

  (state.loopPlatformRows || []).forEach((r) => {
    const tc = safeNum(r.tc, 0);
    if (tc > 0 && tc < minTc) {
      issues.push({ level: "warning", title: "Track Centre Below Minimum", detail: `${r.station || "Station"} TC ${r3(tc)} m < ${r3(minTc)} m.` });
    }
  });

  return issues;
}

function renderQualityChecks() {
  if (!els.qualityResults) return;
  const issues = runQualityChecks();
  if (!issues.length) {
    els.qualityResults.innerHTML = `<div class="glass" style="padding:12px;">No critical issues detected.</div>`;
    return;
  }
  els.qualityResults.innerHTML = issues.map((issue) => `
    <div class="glass" style="padding:12px; border:1px solid rgba(255,255,255,0.12);">
      <div style="font-weight:700; color:${issue.level === "error" ? "#f87171" : "#fbbf24"};">${escapeHtml(issue.title)}</div>
      <div class="muted" style="margin-top:4px;">${escapeHtml(issue.detail)}</div>
    </div>
  `).join("");
}

function exportBoqCsv() {
  const qtyRows = Array.from(document.querySelectorAll("#resultQtyBody tr"));
  if (!qtyRows.length) {
    alert("No quantity summary rows available. Recalculate the project first.");
    return;
  }
  const mapping = state.settings.boqMapping || {};
  const headers = [
    mapping.range || "Range",
    mapping.prepared || "Prepared",
    mapping.blanket || "Blanket",
    mapping.fill || "Fill",
    mapping.cut || "Cut",
  ];
  const csvRows = [headers];
  qtyRows.forEach((tr) => {
    const cells = Array.from(tr.querySelectorAll("td")).map((td) => String(td.textContent || "").trim());
    if (!cells.length) return;
    const row = [
      cells[0] || "",
      cells[1] || "",
      cells[2] || "",
      cells[3] || "",
      cells[4] || "",
    ];
    csvRows.push(row);
  });
  const csv = csvRows.map((row) => row.map((cell) => `"${String(cell).replace(/\"/g, '\"\"')}"`).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${state.project.name || "earthwork"}-boq.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function exportCalculationExcel() {
  if (!state.calcRows || !state.calcRows.length) {
    alert("No calculation rows available. Recalculate the project first.");
    return;
  }
  const headers = [
    "Bridge No / Name",
    "Station",
    "Chainage (m)",
    "Diff (m)",
    "Ground RL",
    "Proposed RL",
    "Total Loop TC",
    "No. of Loops",
    "PF Width",
    "Deducted Len",
    "EW Diff",
    "RL Diff",
    "Bank (m)",
    "Cut (m)",
    "Top Width",
    "Fill Area",
    "Cut Area",
    "Fill Vol",
    "Cut Vol",
  ];
  const rows = state.calcRows.map((r) => {
    const bridgeRefs = (r.bridgeRefs && r.bridgeRefs.length) ? r.bridgeRefs.join(" | ") : "-";
    const station = (r.stationName || r.station) ? String(r.stationName || r.station).replace(/\n/g, " ") : "-";
    const chainageLabel = (r.chainage < 0 ? "-" : "") + Math.floor(Math.abs(r.chainage) / 1000) + "+" + (Math.abs(r.chainage) % 1000).toFixed(3).replace(/(\.\d*?[1-9])0+$|\.0+$/, "$1").padStart(3, "0");
    return [
      bridgeRefs,
      station,
      chainageLabel,
      r.diff ? r3(r.diff) : "—",
      r3(r.groundLevel),
      r3(r.proposedLevel),
      r.loopTc ? r3(r.loopTc) : "—",
      r.loopCount ? r.loopCount : "—",
      r.platformWidth ? r3(r.platformWidth) : "—",
      r.bridgeDeductLen ? r3(r.bridgeDeductLen) : "—",
      r.ewDiff ? r3(r.ewDiff) : "—",
      r.rlDiff ? r3(r.rlDiff) : "—",
      r.bank > 0.0001 ? r3(r.bank) : "—",
      r.cut > 0.0001 ? r3(r.cut) : "—",
      r.topWidth ? r3(r.topWidth) : "—",
      r.fillArea > 0.0001 ? r3(r.fillArea) : "—",
      r.cutArea > 0.0001 ? r3(r.cutArea) : "—",
      r.fillVol > 0.0001 ? r3(r.fillVol) : "—",
      r.cutVol > 0.0001 ? r3(r.cutVol) : "—",
    ];
  });
  if (typeof XLSX !== "undefined") {
    const data = [headers, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Calculation");
    XLSX.writeFile(wb, `${state.project.name || "earthwork"}-calculation.xlsx`);
    return;
  }
  const tableHtml = `
    <table>
      <thead><tr>${headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>
      <tbody>
        ${rows.map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("")}
      </tbody>
    </table>
  `;
  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body>${tableHtml}</body></html>`;
  const blob = new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${state.project.name || "earthwork"}-calculation.xls`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function getSegmentBridgeInfo(start, end, bridges) {
  if (!Number.isFinite(start) || !Number.isFinite(end) || end <= start || !bridges.length) {
    return { deductedLen: 0, refs: [] };
  }
  const overlaps = [];
  const refs = new Set();
  for (const b of bridges) {
    if (b.endChainage <= start || b.startChainage >= end) continue;
    const ovStart = Math.max(start, b.startChainage);
    const ovEnd = Math.min(end, b.endChainage);
    if (ovEnd > ovStart) {
      if (b.shouldDeduct) overlaps.push([ovStart, ovEnd]);
      refs.add(b.bridgeNo);
    }
  }
  if (!overlaps.length) return { deductedLen: 0, refs: [] };
  overlaps.sort((a, b) => a[0] - b[0]);
  let [curS, curE] = overlaps[0];
  let total = 0;
  for (let i = 1; i < overlaps.length; i += 1) {
    const [s, e] = overlaps[i];
    if (s <= curE) {
      curE = Math.max(curE, e);
    } else {
      total += (curE - curS);
      curS = s;
      curE = e;
    }
  }
  total += (curE - curS);
  return { deductedLen: total, refs: Array.from(refs) };
}

function getSlopeFactorForChainage(chainage) {
  if (!Number.isFinite(chainage) || !Array.isArray(state.slopeZones)) return safeNum(state.settings.sideSlopeFactor);
  const zone = state.slopeZones.find((z) => Number.isFinite(z.startCh) && Number.isFinite(z.endCh) && chainage >= z.startCh && chainage <= z.endCh);
  return zone ? safeNum(zone.slope, state.settings.sideSlopeFactor) : safeNum(state.settings.sideSlopeFactor);
}

function getOverrideForChainage(chainage) {
  if (!Number.isFinite(chainage) || !Array.isArray(state.calcOverrides)) return null;
  return state.calcOverrides.find((o) => Number.isFinite(o.startCh) && Number.isFinite(o.endCh) && chainage >= o.startCh && chainage <= o.endCh) || null;
}

function getBridgeRefsAtChainage(chainage, bridges) {
  if (!Number.isFinite(chainage) || !bridges.length) return [];
  return bridges
    .filter((b) => chainage >= b.startChainage && chainage <= b.endChainage)
    .map((b) => b.bridgeNo);
}

function getBridgeStyleInfo(category, type) {
  const typeLC = (String(category || "") + " " + String(type || "")).toLowerCase();
  if (typeLC.includes("tunnel")) return { bg: "rgba(168, 85, 247, 0.15)", col: "#d8b4fe", border: "1px solid rgba(168, 85, 247, 0.3)" };
  if (typeLC.includes("major")) return { bg: "rgba(239, 68, 68, 0.15)", col: "#fca5a5", border: "1px solid rgba(239, 68, 68, 0.3)" };
  if (typeLC.includes("minor")) return { bg: "rgba(16, 185, 129, 0.15)", col: "#6ee7b7", border: "1px solid rgba(16, 185, 129, 0.3)" };
  if (typeLC.includes("rub") || typeLC.includes("rob")) return { bg: "rgba(245, 158, 11, 0.15)", col: "#fcd34d", border: "1px solid rgba(245, 158, 11, 0.3)" };
  return { bg: "rgba(59, 130, 246, 0.15)", col: "#93c5fd", border: "1px solid rgba(59, 130, 246, 0.3)" };
}

function renderBridgeInputs() {
  if (!els.bridgeTableBody) return;
  if (!state.bridgeRows.length) {
    els.bridgeTableBody.innerHTML = `
      <tr>
        <td colspan="13" class="muted">No bridge rows loaded. Use Import Data > Import Bridge List or Add Bridge Row.</td>
      </tr>
    `;
    if (els.bridgeMeta) els.bridgeMeta.textContent = "No bridge deduction applied.";
    return;
  }
  els.bridgeTableBody.innerHTML = state.bridgeRows.map((b, i) => {
    const bColor = getBridgeStyleInfo(b.bridgeCategory, b.bridgeType);
    return `
    <tr data-bridge-row="${i}">
      <td><input data-bridge-field="bridgeNo" value="${String(b.bridgeNo || "")}" style="background: ${bColor.bg}; color: ${bColor.col}; border: ${bColor.border}; border-radius: 4px; font-weight: 600; padding: 4px 6px;" /></td>
      <td>
        <select data-bridge-field="bridgeCategory">
          ${BRIDGE_CATEGORIES.map(cat => `<option value="${cat}" ${b.bridgeCategory === cat ? 'selected' : ''}>${cat}</option>`).join("")}
        </select>
      </td>
      <td>
        <select data-bridge-field="bridgeType">
          ${BRIDGE_TYPES.map(type => `<option value="${type}" ${b.bridgeType === type ? 'selected' : ''}>${type}</option>`).join("")}
        </select>
      </td>
      <td><input data-bridge-field="bridgeSize" value="${String(b.bridgeSize || "")}" style="width: 80px;" placeholder="e.g. 6.1m" /></td>
      <td><input data-bridge-field="bridgeSpans" type="number" min="1" value="${String(b.bridgeSpans || "1")}" style="width: 50px;" /></td>
      <td><input data-bridge-field="bridgeSpanLength" value="${Number.isFinite(b.bridgeSpanLength) ? r3(b.bridgeSpanLength) : String(b.bridgeSpanLength ?? "")}" style="width: 70px;" /></td>
      <td><input data-bridge-field="clearSpan" value="${Number.isFinite(parseLooseNumber(b.clearSpan, NaN)) ? r3(parseLooseNumber(b.clearSpan, NaN)) : String(b.clearSpan ?? "")}" style="width: 70px;" /></td>
      <td>
        <select data-bridge-field="deductRule">
          ${BRIDGE_DEDUCT_RULES.map(rule => `<option value="${rule}" ${String(b.deductRule || "Auto") === rule ? "selected" : ""}>${rule}</option>`).join("")}
        </select>
      </td>
      <td style="text-align:center;">
        <input data-bridge-field="autoDeduct" type="checkbox" class="toggle-input sm" ${b.autoDeduct !== false ? "checked" : ""} />
      </td>
      <td><input data-bridge-field="startChainage" value="${Number.isFinite(parseChainage(b.startChainage)) ? r3(parseChainage(b.startChainage)) : String(b.startChainage ?? "")}" /></td>
      <td><input data-bridge-field="endChainage" value="${Number.isFinite(parseChainage(b.endChainage)) ? r3(parseChainage(b.endChainage)) : String(b.endChainage ?? "")}" /></td>
      <td><input data-bridge-field="length" value="${Number.isFinite(parseLooseNumber(b.length, NaN)) ? r3(parseLooseNumber(b.length, NaN)) : String(b.length ?? "")}" /></td>
      <td><button type="button" class="bridge-del" data-bridge-del="${i}">Delete</button></td>
    </tr>
  `;
  }).join("");
  const allValidBridges = buildBridgeIntervals();
  const deductibleBridges = allValidBridges.filter(b => b.shouldDeduct);
  const totalLength = allValidBridges.reduce((acc, b) => acc + safeNum(b.length), 0);
  const deductibleLength = deductibleBridges.reduce((acc, b) => acc + safeNum(b.length), 0);

  if (els.bridgeMeta) {
    const invalidCount = Math.max(state.bridgeRows.length - allValidBridges.length, 0);
    els.bridgeMeta.textContent = `Bridge rows: ${state.bridgeRows.length} | Valid: ${allValidBridges.length} | Total Length: ${r3(totalLength)}m | Deductible (Major/RoR/etc): ${r3(deductibleLength)}m${invalidCount ? ` | Invalid: ${invalidCount}` : ""}`;
  }
}

function renderCurveInputs() {
  if (!els.curveTableBody) return;
  const rows = Array.isArray(state.curveRows) ? state.curveRows : [];
  if (!rows.length) {
    els.curveTableBody.innerHTML = `<tr><td colspan="5" class="muted">No curve rows loaded. Use Import Data > Import Curve List.</td></tr>`;
    if (els.curveMeta) els.curveMeta.textContent = "No curve list imported.";
    return;
  }
  els.curveTableBody.innerHTML = rows.map((r, i) => `
    <tr data-curve-row="${i}">
      <td><input data-curve-field="curve" value="${String(r.curve || "")}" /></td>
      <td><input data-curve-field="chainage" value="${Number.isFinite(r.chainage) ? r3(r.chainage) : ""}" /></td>
      <td><input data-curve-field="radius" value="${Number.isFinite(r.radius) ? r3(r.radius) : ""}" /></td>
      <td><input data-curve-field="length" value="${Number.isFinite(r.length) ? r3(r.length) : ""}" /></td>
      <td><button type="button" class="bridge-del" data-curve-del="${i}">Delete</button></td>
    </tr>
  `).join("");
  const validLen = rows.reduce((acc, r) => acc + (Number.isFinite(r.length) ? r.length : 0), 0);
  if (els.curveMeta) els.curveMeta.textContent = `Curve rows: ${rows.length} | Total curve length: ${r3(validLen)} m`;
}

function renderLoopInputs() {
  if (!els.loopTableBody) return;
  const rows = Array.isArray(state.loopPlatformRows) ? state.loopPlatformRows : [];
  if (!rows.length) {
    els.loopTableBody.innerHTML = `<tr><td colspan="12" class="muted">No station/loop rows loaded. Use Import Data > Import Loops & Platforms.</td></tr>`;
    if (els.loopMeta) els.loopMeta.textContent = "No station/loops data imported.";
    return;
  }

  const stations = [...new Set(rows.map(r => r.station).filter(Boolean))];
  const stationColors = {};
  stations.forEach((s, idx) => {
    stationColors[s] = `hsl(${(idx * 137.508) % 360}, 65%, 50%)`;
  });

  els.loopTableBody.innerHTML = rows.map((r, i) => {
    const sColor = stationColors[r.station] || 'transparent';
    const bgStyle = sColor !== 'transparent' ? sColor.replace('hsl', 'hsla').replace(')', ', 0.08)') : 'transparent';
    const isMainLine = r.lineType === "Main Line";
    return `
    <tr data-loop-row="${i}" data-drag-row="${i}" draggable="true" style="border-left: 4px solid ${sColor}; background: ${bgStyle}; cursor: grab;">
      <td><input data-loop-field="station" value="${String(r.station || "")}" ${isMainLine ? "disabled" : ""} /></td>
      <td>
        <select data-loop-field="lineType">
          <option value=""></option>
          ${LOOP_LINE_TYPES.map((type) => `<option value="${type}" ${r.lineType === type ? "selected" : ""}>${type}</option>`).join("")}
        </select>
      </td>
      <td>
        <select data-loop-field="side" ${isMainLine ? "disabled" : ""}>
          <option value=""></option>
          ${LOOP_SIDES.map((side) => `<option value="${side}" ${r.side === side ? "selected" : ""}>${side}</option>`).join("")}
        </select>
      </td>
      <td><input data-loop-field="csb" value="${Number.isFinite(r.csb) ? r3(r.csb) : ""}" ${isMainLine ? "disabled" : ""} /></td>
      <td><input data-loop-field="tc" value="${Number.isFinite(r.tc) ? r3(r.tc) : ""}" /></td>
      <td><input data-loop-field="loopStartCh" value="${Number.isFinite(r.loopStartCh) ? r3(r.loopStartCh) : ""}" /></td>
      <td><input data-loop-field="loopEndCh" value="${Number.isFinite(r.loopEndCh) ? r3(r.loopEndCh) : ""}" /></td>
      <td><input data-loop-field="pfWidth" value="${Number.isFinite(r.pfWidth) ? r3(r.pfWidth) : ""}" ${isMainLine ? "disabled" : ""} /></td>
      <td><input data-loop-field="pfStartCh" value="${Number.isFinite(r.pfStartCh) ? r3(r.pfStartCh) : ""}" ${isMainLine ? "disabled" : ""} /></td>
      <td><input data-loop-field="pfEndCh" value="${Number.isFinite(r.pfEndCh) ? r3(r.pfEndCh) : ""}" ${isMainLine ? "disabled" : ""} /></td>
      <td><input data-loop-field="remarks" value="${String(r.remarks || "")}" ${isMainLine ? "disabled" : ""} /></td>
      <td>
        <div style="display:flex;gap:6px;align-items:center;">
          <button type="button" class="btn btn-secondary btn-sm icon-btn-sm" data-loop-up="${i}" title="Move Up" style="padding: 4px; border-radius: 6px; width: 26px; height: 26px;"><i class="ri-arrow-up-line"></i></button>
          <button type="button" class="btn btn-secondary btn-sm icon-btn-sm" data-loop-down="${i}" title="Move Down" style="padding: 4px; border-radius: 6px; width: 26px; height: 26px;"><i class="ri-arrow-down-line"></i></button>
          <button type="button" class="bridge-del" data-loop-del="${i}" style="margin-left:auto;">Delete</button>
        </div>
      </td>
    </tr>
    <tr class="loop-insert-row" data-loop-insert-row="${i}">
      <td colspan="12">
        <button type="button" class="loop-insert-btn" data-loop-insert="${i + 1}" aria-label="Add line between rows">+</button>
      </td>
    </tr>
  `}).join("");
  const loopRanges = rows.filter((r) => Number.isFinite(r.loopStartCh) && Number.isFinite(r.loopEndCh)).length;
  const pfRanges = rows.filter((r) => Number.isFinite(r.pfStartCh) && Number.isFinite(r.pfEndCh)).length;
  if (els.loopMeta) els.loopMeta.textContent = `Stations: ${stations.length} | Rows: ${rows.length} | Loop ranges: ${loopRanges} | Platform ranges: ${pfRanges}`;
}

function syncCurveStateFromTable() {
  if (!els.curveTableBody) return;
  const rows = Array.from(els.curveTableBody.querySelectorAll("tr"));
  const next = [];
  rows.forEach((tr, i) => {
    if (!tr.querySelector('[data-curve-field="curve"]')) return;
    const curve = String(tr.querySelector('[data-curve-field="curve"]')?.value || "").trim();
    const chainage = parseChainage(tr.querySelector('[data-curve-field="chainage"]')?.value ?? "");
    const radius = parseLooseNumber(tr.querySelector('[data-curve-field="radius"]')?.value ?? "", NaN);
    const length = parseLooseNumber(tr.querySelector('[data-curve-field="length"]')?.value ?? "", NaN);
    if (!curve && !Number.isFinite(chainage) && !Number.isFinite(radius) && !Number.isFinite(length)) return;
    next.push({
      curve: curve || `Curve-${i + 1}`,
      chainage: Number.isFinite(chainage) ? chainage : null,
      radius: Number.isFinite(radius) ? radius : null,
      length: Number.isFinite(length) ? length : null,
    });
  });
  state.curveRows = next;
}

function syncLoopStateFromTable() {
  if (!els.loopTableBody) return;
  const rows = Array.from(els.loopTableBody.querySelectorAll("tr"));
  const next = [];
  rows.forEach((tr) => {
    if (!tr.querySelector('[data-loop-field="station"]')) return;
    const station = String(tr.querySelector('[data-loop-field="station"]')?.value || "").trim();
    const lineType = String(tr.querySelector('[data-loop-field="lineType"]')?.value || "").trim();
    const side = String(tr.querySelector('[data-loop-field="side"]')?.value || "").trim();
    const csb = parseChainage(tr.querySelector('[data-loop-field="csb"]')?.value ?? "");
    const tc = parseLooseNumber(tr.querySelector('[data-loop-field="tc"]')?.value ?? "", NaN);
    const loopStartCh = parseChainage(tr.querySelector('[data-loop-field="loopStartCh"]')?.value ?? "");
    const loopEndCh = parseChainage(tr.querySelector('[data-loop-field="loopEndCh"]')?.value ?? "");
    const pfWidth = parseLooseNumber(tr.querySelector('[data-loop-field="pfWidth"]')?.value ?? "", NaN);
    const pfStartCh = parseChainage(tr.querySelector('[data-loop-field="pfStartCh"]')?.value ?? "");
    const pfEndCh = parseChainage(tr.querySelector('[data-loop-field="pfEndCh"]')?.value ?? "");
    const remarks = String(tr.querySelector('[data-loop-field="remarks"]')?.value || "").trim();
    if (!station
      && !lineType
      && !side
      && !Number.isFinite(csb)
      && !Number.isFinite(tc)
      && !Number.isFinite(loopStartCh)
      && !Number.isFinite(loopEndCh)
      && !Number.isFinite(pfWidth)
      && !Number.isFinite(pfStartCh)
      && !Number.isFinite(pfEndCh)
      && !remarks) return;
    next.push({
      station: station || `LP-${next.length + 1}`,
      lineType,
      lineName: lineType,
      side,
      csb: Number.isFinite(csb) ? csb : null,
      tc: Number.isFinite(tc) ? tc : 0,
      loopStartCh: Number.isFinite(loopStartCh) ? loopStartCh : null,
      loopEndCh: Number.isFinite(loopEndCh) ? loopEndCh : null,
      pfWidth: Number.isFinite(pfWidth) ? pfWidth : 0,
      pfStartCh: Number.isFinite(pfStartCh) ? pfStartCh : null,
      pfEndCh: Number.isFinite(pfEndCh) ? pfEndCh : null,
      remarks,
      order: next.length,
    });
  });
  state.loopPlatformRows = next;
}

function syncBridgeStateFromTable() {
  if (!els.bridgeTableBody) return;
  const rows = Array.from(els.bridgeTableBody.querySelectorAll("tr"));
  const parsed = [];
  rows.forEach((tr, i) => {
    if (tr.querySelectorAll("[data-bridge-field]").length === 0) return;
    const raw = {
      bridgeNo: tr.querySelector('[data-bridge-field="bridgeNo"]')?.value ?? "",
      bridgeCategory: tr.querySelector('[data-bridge-field="bridgeCategory"]')?.value ?? "Minor",
      bridgeType: tr.querySelector('[data-bridge-field="bridgeType"]')?.value ?? "Box",
      bridgeSize: tr.querySelector('[data-bridge-field="bridgeSize"]')?.value ?? "-",
      bridgeSpans: tr.querySelector('[data-bridge-field="bridgeSpans"]')?.value ?? "1",
      bridgeSpanLength: tr.querySelector('[data-bridge-field="bridgeSpanLength"]')?.value ?? "",
      clearSpan: tr.querySelector('[data-bridge-field="clearSpan"]')?.value ?? "",
      deductRule: tr.querySelector('[data-bridge-field="deductRule"]')?.value ?? "Auto",
      autoDeduct: tr.querySelector('[data-bridge-field="autoDeduct"]')?.checked ?? true,
      startChainage: tr.querySelector('[data-bridge-field="startChainage"]')?.value ?? "",
      endChainage: tr.querySelector('[data-bridge-field="endChainage"]')?.value ?? "",
      length: tr.querySelector('[data-bridge-field="length"]')?.value ?? "",
    };
    if (!raw.bridgeNo && !raw.startChainage && !raw.endChainage && !raw.length) return;
    if (!raw.bridgeNo) raw.bridgeNo = `BR-${i + 1}`;
    parsed.push(raw);
  });
  state.bridgeRows = parsed;
}

function syncMaterialProfileFromGrid() {
  if (!els.materialProfileGrid) return;
  const rows = Array.from(els.materialProfileGrid.querySelectorAll("tbody tr"));
  const next = [];
  rows.forEach((tr) => {
    const depth = parseLooseNumber(tr.querySelector('[data-mp-field="depth"]')?.value ?? "", NaN);
    const density = parseLooseNumber(tr.querySelector('[data-mp-field="density"]')?.value ?? "", NaN);
    const shrink = parseLooseNumber(tr.querySelector('[data-mp-field="shrink"]')?.value ?? "", NaN);
    const swell = parseLooseNumber(tr.querySelector('[data-mp-field="swell"]')?.value ?? "", NaN);
    if (!Number.isFinite(depth) && !Number.isFinite(density)) return;
    next.push({
      depth: Number.isFinite(depth) ? depth : 0,
      density: Number.isFinite(density) ? density : 0,
      shrink: Number.isFinite(shrink) ? shrink : 0,
      swell: Number.isFinite(swell) ? swell : 0,
    });
  });
  state.settings.materialProfile = next;
}

function renderOverrideTable() {
  if (!els.overrideTableBody) return;
  const rows = Array.isArray(state.calcOverrides) ? state.calcOverrides : [];
  els.overrideTableBody.innerHTML = rows.map((r, i) => `
    <tr>
      <td><input data-ov-field="startCh" value="${Number.isFinite(r.startCh) ? r3(r.startCh) : ""}" /></td>
      <td><input data-ov-field="endCh" value="${Number.isFinite(r.endCh) ? r3(r.endCh) : ""}" /></td>
      <td>
        <select data-ov-field="type">
          <option value="" ${!r.type ? "selected" : ""}>Auto</option>
          <option value="FILL" ${r.type === "FILL" ? "selected" : ""}>Fill</option>
          <option value="CUT" ${r.type === "CUT" ? "selected" : ""}>Cut</option>
          <option value="NEUTRAL" ${r.type === "NEUTRAL" ? "selected" : ""}>Neutral</option>
        </select>
      </td>
      <td><input data-ov-field="bank" value="${Number.isFinite(r.bank) ? r3(r.bank) : ""}" /></td>
      <td><input data-ov-field="cut" value="${Number.isFinite(r.cut) ? r3(r.cut) : ""}" /></td>
      <td><input data-ov-field="slope" value="${Number.isFinite(r.slope) ? r3(r.slope) : ""}" /></td>
      <td><button type="button" class="bridge-del" data-ov-del="${i}">Delete</button></td>
    </tr>
  `).join("");
}

function syncOverridesFromTable() {
  if (!els.overrideTableBody) return;
  const rows = Array.from(els.overrideTableBody.querySelectorAll("tr"));
  const next = [];
  rows.forEach((tr) => {
    const startCh = parseChainage(tr.querySelector('[data-ov-field="startCh"]')?.value ?? "");
    const endCh = parseChainage(tr.querySelector('[data-ov-field="endCh"]')?.value ?? "");
    const type = String(tr.querySelector('[data-ov-field="type"]')?.value || "").trim();
    const bank = parseLooseNumber(tr.querySelector('[data-ov-field="bank"]')?.value ?? "", NaN);
    const cut = parseLooseNumber(tr.querySelector('[data-ov-field="cut"]')?.value ?? "", NaN);
    const slope = parseLooseNumber(tr.querySelector('[data-ov-field="slope"]')?.value ?? "", NaN);
    if (!Number.isFinite(startCh) || !Number.isFinite(endCh)) return;
    next.push({ startCh, endCh, type, bank, cut, slope });
  });
  state.calcOverrides = next;
}

function renderSlopeZoneTable() {
  if (!els.slopeZoneTableBody) return;
  const rows = Array.isArray(state.slopeZones) ? state.slopeZones : [];
  els.slopeZoneTableBody.innerHTML = rows.map((r, i) => `
    <tr>
      <td><input data-sz-field="startCh" value="${Number.isFinite(r.startCh) ? r3(r.startCh) : ""}" /></td>
      <td><input data-sz-field="endCh" value="${Number.isFinite(r.endCh) ? r3(r.endCh) : ""}" /></td>
      <td><input data-sz-field="slope" value="${Number.isFinite(r.slope) ? r3(r.slope) : ""}" /></td>
      <td><button type="button" class="bridge-del" data-sz-del="${i}">Delete</button></td>
    </tr>
  `).join("");
}

function syncSlopeZonesFromTable() {
  if (!els.slopeZoneTableBody) return;
  const rows = Array.from(els.slopeZoneTableBody.querySelectorAll("tr"));
  const next = [];
  rows.forEach((tr) => {
    const startCh = parseChainage(tr.querySelector('[data-sz-field="startCh"]')?.value ?? "");
    const endCh = parseChainage(tr.querySelector('[data-sz-field="endCh"]')?.value ?? "");
    const slope = parseLooseNumber(tr.querySelector('[data-sz-field="slope"]')?.value ?? "", NaN);
    if (!Number.isFinite(startCh) || !Number.isFinite(endCh) || !Number.isFinite(slope)) return;
    next.push({ startCh, endCh, slope });
  });
  state.slopeZones = next;
}

function inferImportInterval() {
  if (!Array.isArray(state.rawRows) || state.rawRows.length < 2) return 20;
  const ch = state.rawRows
    .map((r) => Number(r.chainage))
    .filter((v) => Number.isFinite(v))
    .sort((a, b) => a - b);
  if (ch.length < 2) return 20;
  const diffs = [];
  for (let i = 1; i < ch.length; i += 1) {
    const d = ch[i] - ch[i - 1];
    if (d > 0) diffs.push(d);
  }
  if (!diffs.length) return 20;
  const avg = diffs.reduce((s, d) => s + d, 0) / diffs.length;
  return Number.isFinite(avg) && avg > 0 ? avg : 20;
}

function inferImportStartChainage() {
  if (Array.isArray(state.rawRows) && state.rawRows.length > 0) {
    const values = state.rawRows.map((r) => Number(r.chainage)).filter((v) => Number.isFinite(v));
    if (values.length) return Math.min(...values);
  }
  return 0;
}

function clamp(v, min, max) {
  return Math.min(max, Math.max(min, v));
}

function setCrossViewBox(vb) {
  const baseW = CROSS_DEFAULT_VIEWBOX.w;
  const baseH = CROSS_DEFAULT_VIEWBOX.h;
  const minW = baseW * 0.08;
  const maxW = baseW * 5.5;
  const minH = baseH * 0.08;
  const maxH = baseH * 5.5;
  const w = Math.min(maxW, Math.max(minW, vb.w));
  const h = Math.min(maxH, Math.max(minH, vb.h));
  const panPadX = w * 0.24;
  const panPadY = h * 0.24;
  const minX = CROSS_DEFAULT_VIEWBOX.x - panPadX;
  const minY = CROSS_DEFAULT_VIEWBOX.y - panPadY;
  const maxX = (CROSS_DEFAULT_VIEWBOX.x + CROSS_DEFAULT_VIEWBOX.w) - w + panPadX;
  const maxY = (CROSS_DEFAULT_VIEWBOX.y + CROSS_DEFAULT_VIEWBOX.h) - h + panPadY;
  const x = minX <= maxX ? clamp(vb.x, minX, maxX) : (minX + maxX) / 2;
  const y = minY <= maxY ? clamp(vb.y, minY, maxY) : (minY + maxY) / 2;
  state.crossViewBox = { x, y, w, h };
  if (els.crossSvg) {
    els.crossSvg.setAttribute("viewBox", `${state.crossViewBox.x} ${state.crossViewBox.y} ${state.crossViewBox.w} ${state.crossViewBox.h}`);
  }
}

function animateCrossViewBox(targetVb, duration = 140) {
  if (state.crossZoomFrame) {
    cancelAnimationFrame(state.crossZoomFrame);
    state.crossZoomFrame = null;
  }
  const startVb = { ...state.crossViewBox };
  const startTime = performance.now();
  const step = (now) => {
    const t = Math.min(1, (now - startTime) / duration);
    const eased = 1 - ((1 - t) * (1 - t) * (1 - t));
    setCrossViewBox({
      x: startVb.x + ((targetVb.x - startVb.x) * eased),
      y: startVb.y + ((targetVb.y - startVb.y) * eased),
      w: startVb.w + ((targetVb.w - startVb.w) * eased),
      h: startVb.h + ((targetVb.h - startVb.h) * eased),
    });
    if (t < 1) {
      state.crossZoomFrame = requestAnimationFrame(step);
    } else {
      state.crossZoomFrame = null;
    }
  };
  state.crossZoomFrame = requestAnimationFrame(step);
}

function resetCrossView() {
  animateCrossViewBox({ ...CROSS_DEFAULT_VIEWBOX }, 180);
}

function computeZoomedCrossViewBox(clientX, clientY, factor) {
  if (!els.crossSvg || !Number.isFinite(factor) || factor <= 0) return;
  const rect = els.crossSvg.getBoundingClientRect();
  if (!rect.width || !rect.height) return null;
  const rx = (clientX - rect.left) / rect.width;
  const ry = (clientY - rect.top) / rect.height;
  const curr = state.crossViewBox;
  const newW = curr.w * factor;
  const newH = curr.h * factor;
  const newX = curr.x + ((curr.w - newW) * rx);
  const newY = curr.y + ((curr.h - newH) * ry);
  return { x: newX, y: newY, w: newW, h: newH };
}

function zoomCrossAt(clientX, clientY, factor, mode = "animated") {
  const nextViewBox = computeZoomedCrossViewBox(clientX, clientY, factor);
  if (!nextViewBox) return;
  if (mode === "instant") {
    if (state.crossZoomFrame) {
      cancelAnimationFrame(state.crossZoomFrame);
      state.crossZoomFrame = null;
    }
    setCrossViewBox(nextViewBox);
    return;
  }
  animateCrossViewBox(nextViewBox);
}

function wheelToZoomFactor(e) {
  const modeScale = e.deltaMode === 1 ? 16 : (e.deltaMode === 2 ? 320 : 1);
  const normalized = clamp((e.deltaY * modeScale) / 120, -8, 8);
  const sensitivity = e.ctrlKey ? 0.12 : 0.085;
  return Math.exp(normalized * sensitivity);
}

function getTouchDistance(t0, t1) {
  return Math.hypot(t1.clientX - t0.clientX, t1.clientY - t0.clientY);
}

function getTouchCenter(t0, t1) {
  return {
    x: (t0.clientX + t1.clientX) / 2,
    y: (t0.clientY + t1.clientY) / 2,
  };
}

function bindCrossCanvasInteraction() {
  if (!els.crossSvg || !els.crossGraphicWrap) return;
  els.crossSvg.style.cursor = "grab";
  els.crossGraphicWrap.addEventListener("wheel", (e) => {
    e.preventDefault();
    const factor = wheelToZoomFactor(e);
    zoomCrossAt(e.clientX, e.clientY, factor, e.ctrlKey ? "instant" : "animated");
  }, { passive: false });
  const handleGestureEvent = (e) => {
    e.preventDefault();
    if (state.crossZoomFrame) {
      cancelAnimationFrame(state.crossZoomFrame);
      state.crossZoomFrame = null;
    }
    const gestureScale = Number.isFinite(e.scale) && e.scale > 0 ? e.scale : 1;
    const factor = 1 / gestureScale;
    zoomCrossAt(e.clientX, e.clientY, factor, "instant");
  };
  els.crossGraphicWrap.addEventListener("gesturestart", handleGestureEvent, { passive: false });
  els.crossGraphicWrap.addEventListener("gesturechange", handleGestureEvent, { passive: false });
  els.crossGraphicWrap.addEventListener("gestureend", (e) => {
    e.preventDefault();
  }, { passive: false });
  els.crossSvg.addEventListener("mousedown", (e) => {
    if (e.button !== 0 && e.button !== 1) return;
    e.preventDefault();
    if (state.crossZoomFrame) {
      cancelAnimationFrame(state.crossZoomFrame);
      state.crossZoomFrame = null;
    }
    state.crossPan.active = true;
    state.crossPan.lastX = e.clientX;
    state.crossPan.lastY = e.clientY;
    els.crossSvg.style.cursor = "grabbing";
  });
  window.addEventListener("mousemove", (e) => {
    if (!state.crossPan.active || !els.crossSvg) return;
    const rect = els.crossSvg.getBoundingClientRect();
    if (!rect.width || !rect.height) return;
    const dxPx = e.clientX - state.crossPan.lastX;
    const dyPx = e.clientY - state.crossPan.lastY;
    state.crossPan.lastX = e.clientX;
    state.crossPan.lastY = e.clientY;
    const dx = (dxPx / rect.width) * state.crossViewBox.w;
    const dy = (dyPx / rect.height) * state.crossViewBox.h;
    setCrossViewBox({
      x: state.crossViewBox.x - dx,
      y: state.crossViewBox.y - dy,
      w: state.crossViewBox.w,
      h: state.crossViewBox.h,
    });
  });
  window.addEventListener("mouseup", () => {
    if (!state.crossPan.active || !els.crossSvg) return;
    state.crossPan.active = false;
    els.crossSvg.style.cursor = "grab";
  });

  els.crossGraphicWrap.addEventListener("touchstart", (e) => {
    if (!els.crossSvg) return;
    if (state.crossZoomFrame) {
      cancelAnimationFrame(state.crossZoomFrame);
      state.crossZoomFrame = null;
    }
    if (e.touches.length === 1) {
      const touch = e.touches[0];
      state.crossTouch.mode = "pan";
      state.crossTouch.lastX = touch.clientX;
      state.crossTouch.lastY = touch.clientY;
    } else if (e.touches.length >= 2) {
      const t0 = e.touches[0];
      const t1 = e.touches[1];
      state.crossTouch.mode = "pinch";
      state.crossTouch.startDistance = getTouchDistance(t0, t1);
      state.crossTouch.startCenter = getTouchCenter(t0, t1);
      state.crossTouch.startViewBox = { ...state.crossViewBox };
    }
  }, { passive: true });

  els.crossGraphicWrap.addEventListener("touchmove", (e) => {
    if (!els.crossSvg) return;
    if (state.crossTouch.mode === "pinch" && e.touches.length >= 2) {
      e.preventDefault();
      const rect = els.crossSvg.getBoundingClientRect();
      if (!rect.width || !rect.height || !state.crossTouch.startViewBox || !state.crossTouch.startCenter) return;
      const t0 = e.touches[0];
      const t1 = e.touches[1];
      const center = getTouchCenter(t0, t1);
      const distance = getTouchDistance(t0, t1);
      if (!distance || !state.crossTouch.startDistance) return;
      const scale = state.crossTouch.startDistance / distance;
      const startVb = state.crossTouch.startViewBox;
      const rx = (center.x - rect.left) / rect.width;
      const ry = (center.y - rect.top) / rect.height;
      const newW = startVb.w * scale;
      const newH = startVb.h * scale;
      const newX = startVb.x + ((startVb.w - newW) * rx);
      const newY = startVb.y + ((startVb.h - newH) * ry);
      setCrossViewBox({ x: newX, y: newY, w: newW, h: newH });
      return;
    }
    if (state.crossTouch.mode === "pan" && e.touches.length === 1) {
      e.preventDefault();
      const rect = els.crossSvg.getBoundingClientRect();
      if (!rect.width || !rect.height) return;
      const touch = e.touches[0];
      const dxPx = touch.clientX - state.crossTouch.lastX;
      const dyPx = touch.clientY - state.crossTouch.lastY;
      state.crossTouch.lastX = touch.clientX;
      state.crossTouch.lastY = touch.clientY;
      const dx = (dxPx / rect.width) * state.crossViewBox.w;
      const dy = (dyPx / rect.height) * state.crossViewBox.h;
      setCrossViewBox({
        x: state.crossViewBox.x - dx,
        y: state.crossViewBox.y - dy,
        w: state.crossViewBox.w,
        h: state.crossViewBox.h,
      });
    }
  }, { passive: false });

  els.crossGraphicWrap.addEventListener("touchend", (e) => {
    if (e.touches.length >= 2) {
      const t0 = e.touches[0];
      const t1 = e.touches[1];
      state.crossTouch.mode = "pinch";
      state.crossTouch.startDistance = getTouchDistance(t0, t1);
      state.crossTouch.startCenter = getTouchCenter(t0, t1);
      state.crossTouch.startViewBox = { ...state.crossViewBox };
      return;
    }
    if (e.touches.length === 1) {
      const touch = e.touches[0];
      state.crossTouch.mode = "pan";
      state.crossTouch.lastX = touch.clientX;
      state.crossTouch.lastY = touch.clientY;
      state.crossTouch.startViewBox = null;
      return;
    }
    state.crossTouch.mode = "none";
    state.crossTouch.startDistance = 0;
    state.crossTouch.startViewBox = null;
    state.crossTouch.startCenter = null;
  });

  const tooltip = document.getElementById("svgTooltip");
  els.crossSvg.addEventListener("mousemove", (e) => {
    if (state.crossPan.active || !state.crossMatrix) {
      if (tooltip) tooltip.style.opacity = 0;
      return;
    }
    const rect = els.crossSvg.getBoundingClientRect();
    if (!rect.width || !rect.height) return;

    const ratioX = (e.clientX - rect.left) / rect.width;
    const ratioY = (e.clientY - rect.top) / rect.height;

    // Scale view coordinates
    const mx = state.crossMatrix;
    const svgX = state.crossViewBox.x + ratioX * state.crossViewBox.w;
    const svgY = state.crossViewBox.y + ratioY * state.crossViewBox.h;

    // Reverse Engineering map
    const dist = (svgX - mx.centerX) / mx.pxPerM;
    const elev = mx.fl + (mx.topY - svgY) / mx.pxPerM;

    if (tooltip) {
      tooltip.style.left = `${e.clientX + 15}px`;
      tooltip.style.top = `${e.clientY + 15}px`;
      const sign = dist > 0 ? "+" : "";
      tooltip.innerHTML = `<strong>Dist:</strong> ${sign}${dist.toFixed(2)}m<br/><strong>Elev:</strong> ${elev.toFixed(2)} RL`;
      tooltip.style.opacity = 1;
    }
  });

  els.crossSvg.addEventListener("mouseleave", () => {
    if (tooltip) tooltip.style.opacity = 0;
  });
}

function updateEstimates() {
  if (!els.estimatesBody) return;
  if (!state.project.verified) {
    els.estimatesBody.innerHTML = '<tr><td colspan="7" style="text-align: center; color: var(--muted); padding: 24px;">Please Verify & Calculate the project first.</td></tr>';
    if (els.estimateGrandTotal) els.estimateGrandTotal.textContent = "₹0.00";
    return;
  }

  const rates = {
    clearing: parseFloat(els.rateClearing?.value || 0),
    benching: parseFloat(els.rateBenching?.value || 0),
    fill: parseFloat(els.rateFilling?.value || 0),
    blanket: parseFloat(els.rateBlanketing?.value || 0),
    cutSoil: parseFloat(els.rateCutSoil?.value || 0),
    cutSoft: parseFloat(els.rateCutSoft?.value || 0),
    cutHardBlast: parseFloat(els.rateCutHardBlast?.value || 0),
    cutHardChisel: parseFloat(els.rateCutHardChisel?.value || 0),
    extraLead: parseFloat(els.rateExtraLead?.value || 0),
    turf: parseFloat(els.rateTurfing?.value || 0)
  };

  const pcts = {
    soil: parseFloat(els.pctSoil?.value || 0) / 100,
    soft: parseFloat(els.pctSoftRock?.value || 0) / 100,
    blast: parseFloat(els.pctHardBlast?.value || 0) / 100,
    chisel: parseFloat(els.pctHardChisel?.value || 0) / 100,
    reuse: parseFloat(els.pctReusableSpoil?.value || 0) / 100,
    lead: parseFloat(els.leadKm?.value || 3)
  };

  const rawFillTotal = state.calcRows.reduce((s, r) => s + r.fillVol, 0);
  const rawCutTotal = state.calcRows.reduce((s, r) => s + r.cutVol, 0);

  const cutBlanketTotal = state.calcRows.reduce((s, r) => s + (r.type === 'CUTTING' && r.blanketingVol ? parseLooseNumber(r.blanketingVol) : 0), 0);
  const fillBlanketTotal = state.calcRows.reduce((s, r) => s + (r.bank > 0 && r.blanketingVol ? parseLooseNumber(r.blanketingVol) : 0), 0);
  const cutTurfTotal = state.calcRows.reduce((s, r) => s + (r.type === 'CUTTING' && r.turfingSqM ? parseLooseNumber(r.turfingSqM) : 0), 0);
  const fillTurfTotal = state.calcRows.reduce((s, r) => s + (r.bank > 0 && r.turfingSqM ? parseLooseNumber(r.turfingSqM) : 0), 0);

  const fillLenM = state.calcRows.reduce((s, r) => s + ((r.bank > 0 ? safeNum(r.ewDiff) : 0)), 0);
  const cutLenM = state.calcRows.reduce((s, r) => s + ((r.type === 'CUTTING' ? safeNum(r.ewDiff) : 0)), 0);

  const topWidthFill = safeNum(state.settings?.formationWidthFill) || 7.85;
  const bottomWidthCut = safeNum(state.settings?.cuttingWidth) || 7.85;

  const clearingSqm = (fillLenM * topWidthFill) + (cutLenM * bottomWidthCut); // simplified footprint roughly
  const benchingKm = fillLenM / 1000;

  const cutSoil = rawCutTotal * pcts.soil;
  const reusableSpoil = cutSoil * pcts.reuse;
  const extraLeadQty = Math.max(0, cutSoil - reusableSpoil); // Assume unused soil taken offsite
  const extraLeadKm = Math.max(0, pcts.lead - 2); // IR generally gives 2km free lead

  const sq3Required = Math.max(0, rawFillTotal - reusableSpoil);
  const turfing100Sqm = (cutTurfTotal + fillTurfTotal) / 100;

  let html = "";
  let grandTotal = 0;

  const items = [
    { sn: "1", ussr: "011010", desc: "Preparation of foundation by clearing, grubbing & stripping", qty: clearingSqm / 100, unit: "100 Sqm", rate: rates.clearing },
    { sn: "2", ussr: "011020", desc: "Benching of existing embankment manually", qty: benchingKm, unit: "Km", rate: rates.benching },
    { sn: "3", ussr: "NS Item", desc: "Earthwork with contractor's earth in embankment SQ3 (after deducting reusable cut spoil)", qty: sq3Required, unit: "Cum", rate: rates.fill },
    { sn: "4", ussr: "012040", desc: "Providing blanketing on top of formation", qty: cutBlanketTotal + fillBlanketTotal, unit: "Cum", rate: rates.blanket },
    { sn: "5", ussr: "012011", desc: "Earthwork in cutting in formation in all soils", qty: cutSoil, unit: "Cum", rate: rates.cutSoil },
    { sn: "6", ussr: "012012", desc: "Earthwork in cutting in soft rock not required blasting", qty: rawCutTotal * pcts.soft, unit: "Cum", rate: rates.cutSoft },
    { sn: "7", ussr: "012013", desc: "In hard rock with Blasting", qty: rawCutTotal * pcts.blast, unit: "Cum", rate: rates.cutHardBlast },
    { sn: "8", ussr: "012014", desc: "In hard rock with hammer / chisel", qty: rawCutTotal * pcts.chisel, unit: "Cum", rate: rates.cutHardChisel },
    { sn: "9", ussr: "012015", desc: `Extra for leading of Cut spoil beyond 2km (${pcts.lead}km total - 2km base = ${extraLeadKm}km lead)`, qty: extraLeadQty * extraLeadKm, unit: "Cum-Km", rate: rates.extraLead },
    { sn: "10", ussr: "013023", desc: "Turfing / planting, including watering until properly rooted", qty: turfing100Sqm, unit: "100 Sqm", rate: rates.turf }
  ];

  items.forEach(item => {
    if (item.qty <= 0) return;
    const cost = item.qty * item.rate;
    grandTotal += cost;
    html += `
      <tr>
        <td style="text-align: center; color: var(--muted);">${item.sn}</td>
        <td style="color: var(--blue); font-family: monospace; font-weight: 500;">${item.ussr}</td>
        <td>${item.desc}</td>
        <td style="text-align: right;">${item.qty.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
        <td style="text-align: center; color: var(--muted);">${item.unit}</td>
        <td style="text-align: right;">₹${item.rate.toFixed(2)}</td>
        <td style="text-align: right; font-weight: 600;">₹${cost.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
      </tr>
    `;
  });

  if (html === "") {
    html = '<tr><td colspan="7" style="text-align: center; color: var(--muted); padding: 24px;">No quantities generated to estimate.</td></tr>';
  }

  els.estimatesBody.innerHTML = html;
  if (els.estimateGrandTotal) {
    els.estimateGrandTotal.textContent = "₹" + grandTotal.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
}

function recalculate() {
  const settings = state.settings;
  const sorted = [...state.rawRows].sort((a, b) => a.chainage - b.chainage);
  const bridgeIntervals = buildBridgeIntervals();
  const groupedStations = getGroupedStations();
  const medianInterval = getMedianInterval(sorted);
  const stationTolerance = Number.isFinite(medianInterval) ? Math.max(5, medianInterval * 0.75) : 25;
  let prev = null;
  const rows = [];

  for (const row of sorted) {
    const diff = prev ? Math.max(row.chainage - prev.chainage, 0) : 0;
    const segmentBridgeInfo = prev
      ? getSegmentBridgeInfo(prev.chainage, row.chainage, bridgeIntervals)
      : { deductedLen: 0, refs: [] };
    const bridgeRefsAtPoint = getBridgeRefsAtChainage(row.chainage, bridgeIntervals);
    const bridgeRefs = Array.from(new Set([...segmentBridgeInfo.refs, ...bridgeRefsAtPoint]));
    const bridgeDeductLen = Math.min(diff, segmentBridgeInfo.deductedLen);
    const ewDiff = Math.max(diff - bridgeDeductLen, 0);
    const gl = safeNum(row.groundLevel);
    const fl = safeNum(row.proposedLevel);
    const rlDiff = fl - gl;
    const lp = getLoopPlatformAtChainage(row.chainage);
    const loopTc = safeNum(lp.loopTc, 0);
    const loopCount = safeNum(lp.loopCount, 0);
    const platformWidth = safeNum(lp.platformWidth, 0);
    const effectiveFormationWidth = settings.formationWidthFill + loopTc + platformWidth;
    const stationFromLoops = resolveStationNameAtChainage(row.chainage, groupedStations, stationTolerance);
    const stationRaw = String(row.station || "").trim();
    const stationName = stationRaw || stationFromLoops;

    const bank = diff <= 0 ? 0 : Math.max(rlDiff, 0);
    const minCover = settings.blanketThickness + settings.preparedSubgradeThickness;
    const cut = diff <= 0 ? 0 : (rlDiff < 0 ? (-rlDiff + minCover) : Math.max(minCover - bank, 0));
    const slopeFactor = getSlopeFactorForChainage(row.chainage);

    let turfing = bank > 0 ? slopeFactor * bank : 0;
    const fillTop = effectiveFormationWidth;
    const fillBottom = fillTop + (2 * slopeFactor * bank);
    const cutTop = settings.cuttingWidth;
    const cutBottom = cutTop + (2 * slopeFactor * cut);
    const topWidth = bank > 0 ? fillTop : (cut > 0 ? cutTop : effectiveFormationWidth);

    let finalBank = bank;
    let finalCut = cut;
    let finalSlope = slopeFactor;
    const override = getOverrideForChainage(row.chainage);
    if (override) {
      if (override.type === "FILL") { finalBank = Number.isFinite(override.bank) ? override.bank : finalBank; finalCut = 0; }
      if (override.type === "CUT") { finalCut = Number.isFinite(override.cut) ? override.cut : finalCut; finalBank = 0; }
      if (override.type === "NEUTRAL") { finalBank = 0; finalCut = 0; }
      if (Number.isFinite(override.bank)) finalBank = override.bank;
      if (Number.isFinite(override.cut)) finalCut = override.cut;
      if (Number.isFinite(override.slope)) finalSlope = override.slope;
    }

    const finalFillBottom = fillTop + (2 * finalSlope * finalBank);
    const finalCutBottom = cutTop + (2 * finalSlope * finalCut);
    const finalTopWidth = finalBank > 0 ? fillTop : (finalCut > 0 ? cutTop : effectiveFormationWidth);
    turfing = finalBank > 0 ? finalSlope * finalBank : 0;

    const fillArea = finalBank > 0 ? ((fillTop + finalFillBottom) * 0.5 * finalBank) : 0;
    const cutArea = finalCut > 0 ? ((cutTop + finalCutBottom) * 0.5 * finalCut) : 0;

    const fillVol = prev ? ((prev.fillArea + fillArea) * 0.5 * ewDiff) : 0;
    const cutVol = prev ? ((prev.cutArea + cutArea) * 0.5 * ewDiff) : 0;

    let type = "NEUTRAL";
    if (ewDiff <= 0.001 && bridgeDeductLen > 0) type = "BRIDGE";
    else if (finalBank > 0.001) type = "FILLING";
    else if (finalCut > 0.001) type = "CUTTING";

    const calc = {
      ...row,
      diff,
      ewDiff,
      rlDiff,
      bank: finalBank,
      cut: finalCut,
      turfing,
      fillArea,
      cutArea,
      fillVol,
      cutVol,
      loopTc,
      loopCount,
      platformWidth,
      effectiveFormationWidth,
      bridgeRefs,
      bridgeDeductLen,
      type,
      topWidth: finalTopWidth,
      fillBottom: finalFillBottom,
      cutBottom: finalCutBottom,
      slopeFactor: finalSlope,
      stationName,
    };
    rows.push(calc);
    prev = calc;
  }

  state.calcRows = rows;
  renderBridgeInputs();
  renderCurveInputs();
  renderLoopInputs();
  renderSummary();
  renderTable();
  renderCharts();
  renderRollDiagram();
  renderSideView();
  updateEstimates();
  updateDashboard();

  if (state.kmlData && typeof drawAlignmentMap === "function") {
    requestAnimationFrame(() => drawAlignmentMap());
  }
}

function renderSummary() {
  const fillTotal = state.calcRows.reduce((s, r) => s + r.fillVol, 0);
  const cutTotal = state.calcRows.reduce((s, r) => s + r.cutVol, 0);
  const fillLen = state.calcRows.reduce((s, r) => s + ((r.bank > 0 ? safeNum(r.ewDiff) : 0)), 0) / 1000;
  const cutLen = state.calcRows.reduce((s, r) => s + ((r.type === "CUTTING" ? safeNum(r.ewDiff) : 0)), 0) / 1000;
  const reusablePct = parseLooseNumber(els.pctReusableSpoil?.value, 60);
  const reusableSpoil = cutTotal * (safeNum(reusablePct, 60) / 100);
  const grossFillEquivalent = fillTotal + reusableSpoil;

  els.totalFilling.textContent = formatVolume(fillTotal);
  els.totalCutting.textContent = formatVolume(cutTotal);

  if (els.actualFilling) {
    els.actualFilling.textContent = fillTotal >= 100000 ? `${Number(fillTotal).toLocaleString(undefined, { minimumFractionDigits: 3, maximumFractionDigits: 3 })} m³` : "";
    els.actualFilling.style.opacity = fillTotal >= 100000 ? "0.8" : "0";
  }
  if (els.reusableFilling) {
    els.reusableFilling.textContent = reusableSpoil > 0
      ? `Reusable Earth Deducted: ${Number(reusableSpoil).toLocaleString(undefined, { minimumFractionDigits: 3, maximumFractionDigits: 3 })} m³`
      : "";
    els.reusableFilling.style.opacity = reusableSpoil > 0 ? "0.92" : "0";
  }
  if (els.actualCutting) {
    els.actualCutting.textContent = cutTotal >= 100000 ? `${Number(cutTotal).toLocaleString(undefined, { minimumFractionDigits: 3, maximumFractionDigits: 3 })} m³` : "";
    els.actualCutting.style.opacity = cutTotal >= 100000 ? "0.8" : "0";
  }

  els.fillLength.textContent = `Length: ${r3(fillLen)} km`;
  els.cutLength.textContent = `Length: ${r3(cutLen)} km`;

  // Compute Fluid Animation Fill Levels
  const grandVol = fillTotal + cutTotal;
  const fillCapEl = document.getElementById("fillCapacity");
  const cutCapEl = document.getElementById("cutCapacity");
  if (grandVol > 0) {
    const fillPct = Math.round((fillTotal / grandVol) * 100);
    const cutPct = Math.round((cutTotal / grandVol) * 100);
    if (els.fillWaterNode) els.fillWaterNode.style.height = `${Math.min(fillPct, 100)}%`;
    if (els.fillReusableHatch) {
      const reusableShare = grossFillEquivalent > 0 ? (reusableSpoil / grossFillEquivalent) * 100 : 0;
      els.fillReusableHatch.style.height = `${Math.max(0, Math.min(reusableShare, 100))}%`;
      els.fillReusableHatch.style.opacity = reusableSpoil > 0 ? "0.95" : "0";
      if (els.fillReusablePctLabel) {
        els.fillReusablePctLabel.textContent = `${Math.round(reusableShare)}% reduced`;
        els.fillReusablePctLabel.style.display = reusableSpoil > 0 ? "inline-flex" : "none";
      }
    }
    if (els.cutWaterNode) els.cutWaterNode.style.height = `${Math.min(cutPct, 100)}%`;
    if (fillCapEl) fillCapEl.textContent = `${fillPct}% capacity`;
    if (cutCapEl) cutCapEl.textContent = `${cutPct}% capacity`;
  } else {
    if (els.fillWaterNode) els.fillWaterNode.style.height = "0%";
    if (els.fillReusableHatch) {
      els.fillReusableHatch.style.height = "0%";
      els.fillReusableHatch.style.opacity = "0";
      if (els.fillReusablePctLabel) els.fillReusablePctLabel.style.display = "none";
    }
    if (els.cutWaterNode) els.cutWaterNode.style.height = "0%";
    if (fillCapEl) fillCapEl.textContent = "0% capacity";
    if (cutCapEl) cutCapEl.textContent = "0% capacity";
  }
  renderFormulaSummary();

  // Enable/disable the global verify button
  const vBtn = document.getElementById("verifyCalcBtn");
  if (vBtn) vBtn.disabled = !state.calcRows.length;

  saveState();
}

function setResultTab(tabName) {
  const selected = tabName || "inputs";
  state.activeResultTab = selected;
  els.resultTabButtons.forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.resultTab === selected);
  });
  els.resultTabPanes.forEach((pane) => {
    pane.classList.toggle("active", pane.dataset.resultPane === selected);
  });
}

function setWorkPage(pageName) {
  const selected = pageName || "overview";
  state.activeWorkPage = selected;
  els.workPageButtons.forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.workPageBtn === selected);
  });
  els.workPages.forEach((page) => {
    page.classList.toggle("active", page.dataset.workPage === selected);
  });
  if (selected === "roll-diagram" && state.calcRows.length) {
    requestAnimationFrame(() => { renderRollDiagram(); renderSideView(); });
  }
  if (selected === "graphs" && state.calcRows.length) {
    requestAnimationFrame(() => renderCharts());
  }
  if (selected === "bridges") {
    renderBridgeInputs();
  }
  if (selected === "curves") {
    renderCurveInputs();
  }
  if (selected === "loops") {
    renderLoopInputs();
  }
}

function isProjectReadyForVerification() {
  const u = state.project.uploads;
  return Boolean(state.project.name && u.levels && u.curves && u.bridges && u.loops);
}

function updateWizardUI() {
  if (!els.projectNameInput) return;
  if (document.activeElement !== els.projectNameInput) {
    els.projectNameInput.value = state.project.name || "";
  }
  const tickMap = {
    levels: els.wizardTickLevels,
    curves: els.wizardTickCurves,
    bridges: els.wizardTickBridges,
    loops: els.wizardTickLoops,
    kml: els.wizardTickKml,
  };
  Object.entries(tickMap).forEach(([k, el]) => {
    if (!el) return;
    el.classList.toggle("done", Boolean(state.project.uploads[k]));
  });
  const ready = isProjectReadyForVerification();
  if (els.wizardCalculateBtn) els.wizardCalculateBtn.disabled = !ready;
  if (els.wizardStatus) {
    if (ready) {
      els.wizardStatus.textContent = state.project.uploads.kml
        ? "All core files and KML/KMZ alignment uploaded. Project is verified and ready for calculation."
        : "All four core files uploaded. Project is verified and ready for calculation. KML/KMZ alignment is optional.";
    } else {
      const missing = [];
      if (!state.project.name) missing.push("Project Name");
      if (!state.project.uploads.levels) missing.push("Levels");
      if (!state.project.uploads.curves) missing.push("Curve List");
      if (!state.project.uploads.bridges) missing.push("Bridge List");
      if (!state.project.uploads.loops) missing.push("Loops & Platforms");
      const optional = state.project.uploads.kml ? "" : " | Optional: KML/KMZ Alignment";
      els.wizardStatus.textContent = `Pending: ${missing.join(", ")}${optional}`;
    }
  }
}

function applyProjectGate() {
  const active = Boolean(state.project.active);
  if (els.importBtn) els.importBtn.disabled = !active;
  if (els.openSettingsBtn) els.openSettingsBtn.disabled = !active;
  if (els.saveProjectBtn) els.saveProjectBtn.disabled = !active;
  if (els.projectMeta) {
    els.projectMeta.textContent = "Earthwork Calculations";
  }
}

function collectProjectPayload() {
  return {
    version: "1.0",
    savedAt: new Date().toISOString(),
    project: {
      ...state.project,
      uploads: { ...state.project.uploads },
      profile: { ...state.project.profile },
    },
    settings: { ...state.settings },
    rawRows: state.rawRows,
    bridgeRows: state.bridgeRows,
    curveRows: state.curveRows,
    loopPlatformRows: state.loopPlatformRows,
    calcOverrides: state.calcOverrides,
    slopeZones: state.slopeZones,
    importMappings: state.importMappings,
    kmlData: state.kmlData,
    stationPlans: state.stationPlans,
  };
}

function loadProjectFromPayload(data, options = {}) {
  const { silent = false, fileHandle = null } = options;
  if (!data || typeof data !== "object") {
    alert("Project file is invalid.");
    return;
  }
  state.project = {
    active: true,
    verified: Boolean(data?.project?.verified),
    name: String(data?.project?.name || "Imported Project").trim(),
    uploads: {
      levels: Boolean(data?.project?.uploads?.levels),
      curves: Boolean(data?.project?.uploads?.curves),
      bridges: Boolean(data?.project?.uploads?.bridges),
      loops: Boolean(data?.project?.uploads?.loops),
      kml: Boolean(data?.project?.uploads?.kml) || Boolean(data?.kmlData),
    },
    profile: {
      corridorName: String(data?.project?.profile?.corridorName || ""),
      direction: String(data?.project?.profile?.direction || "Up"),
      chainageZeroRef: String(data?.project?.profile?.chainageZeroRef || ""),
    },
  };
  state.settings = { ...state.seedDefaultSettings, ...(data?.settings || {}) };
  if (!state.settings.reportBrand || typeof state.settings.reportBrand !== "object") {
    state.settings.reportBrand = { ...state.seedDefaultSettings.reportBrand };
  } else {
    state.settings.reportBrand = { ...state.seedDefaultSettings.reportBrand, ...state.settings.reportBrand };
  }
  if (!state.settings.boqMapping || typeof state.settings.boqMapping !== "object") {
    state.settings.boqMapping = { ...state.seedDefaultSettings.boqMapping };
  } else {
    state.settings.boqMapping = { ...state.seedDefaultSettings.boqMapping, ...state.settings.boqMapping };
  }
  state.rawRows = Array.isArray(data?.rawRows) && data.rawRows.length ? data.rawRows : [...state.seedRows];
  state.bridgeRows = Array.isArray(data?.bridgeRows) ? data.bridgeRows : [];
  state.curveRows = Array.isArray(data?.curveRows) ? data.curveRows : [];
  const rawLoopRows = Array.isArray(data?.loopPlatformRows) ? data.loopPlatformRows : [];
  state.loopPlatformRows = rawLoopRows.map((r, i) => ({
    station: String(r?.station || r?.name || `LP-${i + 1}`),
    lineType: normalizeLoopLineType(r?.lineType || r?.lineName || r?.line || ""),
    lineName: String(r?.lineName || r?.lineType || r?.line || ""),
    side: normalizeLoopSide(r?.side || ""),
    csb: Number.isFinite(parseChainage(r?.csb)) ? parseChainage(r.csb) : null,
    tc: Number.isFinite(parseLooseNumber(r?.tc, NaN)) ? parseLooseNumber(r.tc, NaN) : 0,
    loopStartCh: Number.isFinite(parseChainage(r?.loopStartCh))
      ? parseChainage(r.loopStartCh)
      : (Number.isFinite(parseChainage(r?.startChainage)) ? parseChainage(r.startChainage) : null),
    loopEndCh: Number.isFinite(parseChainage(r?.loopEndCh))
      ? parseChainage(r.loopEndCh)
      : (Number.isFinite(parseChainage(r?.endChainage)) ? parseChainage(r.endChainage) : null),
    pfWidth: Number.isFinite(parseLooseNumber(r?.pfWidth, NaN)) ? parseLooseNumber(r.pfWidth, NaN) : 0,
    pfStartCh: Number.isFinite(parseChainage(r?.pfStartCh)) ? parseChainage(r.pfStartCh) : null,
    pfEndCh: Number.isFinite(parseChainage(r?.pfEndCh)) ? parseChainage(r.pfEndCh) : null,
    remarks: String(r?.remarks || r?.type || ""),
    order: Number.isFinite(parseLooseNumber(r?.order, NaN)) ? parseLooseNumber(r.order, NaN) : i,
  }));
  state.kmlData = data?.kmlData && Array.isArray(data.kmlData?.points) ? data.kmlData : null;
  state.stationPlans = data?.stationPlans && typeof data.stationPlans === "object" ? data.stationPlans : {};
  state.calcOverrides = Array.isArray(data?.calcOverrides) ? data.calcOverrides : [];
  state.slopeZones = Array.isArray(data?.slopeZones) ? data.slopeZones : [];
  state.importMappings = data?.importMappings && typeof data.importMappings === "object"
    ? data.importMappings
    : { client: "", templates: {} };
  state.projectFileHandle = fileHandle;
  renderBridgeInputs();
  renderCurveInputs();
  renderLoopInputs();
  recalculate();
  updateWizardUI();
  applyProjectGate();
  if (!silent) alert(`Loaded project: ${state.project.name}`);
}

function escapeAttr(v) {
  return escapeHtml(v).replace(/'/g, "&#39;");
}

async function saveCurrentProject() {
  if (!state.project.active || !state.project.name) {
    alert("Create and name a project before saving.");
    return;
  }
  const payload = collectProjectPayload();
  const fileBase = String(state.project.name || "project")
    .replace(/[^\w.-]+/g, "_")
    .replace(/^_+|_+$/g, "")
    || "project";
  const jsonStr = JSON.stringify(payload, null, 2);

  try {
    if (state.projectFileHandle) {
      const writable = await state.projectFileHandle.createWritable();
      await writable.write(jsonStr);
      await writable.close();
      state.meta = { ...(state.meta || {}), lastSavedAt: new Date().toISOString() };
      alert(`Project updated: ${state.project.name}`);
    } else if ("showSaveFilePicker" in window) {
      const handle = await window.showSaveFilePicker({
        suggestedName: `${fileBase}.EW`,
        types: [{
          description: "Earthwork Project File",
          accept: { "application/json": [".EW"] },
        }],
      });
      state.projectFileHandle = handle;
      const writable = await handle.createWritable();
      await writable.write(jsonStr);
      await writable.close();
      state.meta = { ...(state.meta || {}), lastSavedAt: new Date().toISOString() };
      alert(`Project saved: ${state.project.name}`);
    } else {
      // Fallback for unsupported browsers
      const blob = new Blob([jsonStr], { type: "application/json" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `${fileBase}.EW`;
      a.click();
      URL.revokeObjectURL(a.href);
      state.meta = { ...(state.meta || {}), lastSavedAt: new Date().toISOString() };
      alert(`Project saved: ${state.project.name}`);
    }
  } catch (err) {
    if (err.name !== "AbortError") {
      console.error("Save project error:", err);
      alert("Failed to save the project. " + err.message);
    }
  }
}

function resetForNewProject() {
  localStorage.removeItem("earthsoft_saved_work");
  state.project = {
    active: false,
    verified: false,
    name: "",
    uploads: {
      levels: false,
      curves: false,
      bridges: false,
      loops: false,
      kml: false,
    },
  };
  state.rawRows = [];
  state.bridgeRows = [];
  state.curveRows = [];
  state.loopPlatformRows = [];
  state.kmlData = null;
  state.stationPlans = {};
  state.projectFileHandle = null;
  state.settings = { ...state.seedDefaultSettings };
  if (els.projectNameInput) els.projectNameInput.value = "";
  if (els.importInput) els.importInput.value = "";
  if (els.bridgeImportInput) els.bridgeImportInput.value = "";
  if (els.curveImportInput) els.curveImportInput.value = "";
  if (els.loopImportInput) els.loopImportInput.value = "";
  if (els.kmlImportInput) els.kmlImportInput.value = "";
  if (els.projectImportInput) els.projectImportInput.value = "";
  renderBridgeInputs();
  recalculate();
  updateWizardUI();
  applyProjectGate();
}

function escapeHtml(v) {
  return String(v ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}



function renderFormulaSummary() {
  if (!els.resultInputBody || !els.resultFillBody || !els.resultCutBody || !els.resultQtyBody) return;
  const rows = state.calcRows;
  const s = state.settings;
  const showRails = s.showRails !== false;
  const showPlatforms = s.showPlatforms !== false;
  const showDrains = s.showDrains !== false;
  const showLabels = s.showLabels !== false;
  if (!rows.length) {
    els.resultInputBody.innerHTML = "";
    els.resultFillBody.innerHTML = "";
    els.resultCutBody.innerHTML = "";
    els.resultQtyBody.innerHTML = "";
    return;
  }

  const minCover = s.blanketThickness + s.preparedSubgradeThickness;
  const sumLen = (pred) => rows.reduce((acc, r) => acc + ((safeNum(r.ewDiff) > 0 && pred(r)) ? safeNum(r.ewDiff) : 0), 0);
  const sumVol = (pred, key) => rows.reduce((acc, r) => acc + (pred(r) ? safeNum(r[key]) : 0), 0);
  const avg = (arr) => (arr.length ? (arr.reduce((a, b) => a + b, 0) / arr.length) : 0);
  const max = (arr) => (arr.length ? Math.max(...arr) : 0);
  const cutDepth = (r) => Math.max(r.cut - minCover, 0);

  const fillLenM = sumLen((r) => r.bank > 0);
  const fillLt5M = sumLen((r) => r.bank > 0 && r.bank <= 5);
  const fill5To10M = sumLen((r) => r.bank > 5 && r.bank <= 10);
  const fillGt10M = sumLen((r) => r.bank > 10);
  const fillValues = rows.filter((r) => safeNum(r.ewDiff) > 0 && r.bank > 0).map((r) => r.bank);

  const cutLenM = sumLen((r) => cutDepth(r) > 0);
  const cutLt5M = sumLen((r) => cutDepth(r) > 0 && cutDepth(r) <= 5);
  const cut5To10M = sumLen((r) => cutDepth(r) > 5 && cutDepth(r) <= 10);
  const cutGt10M = sumLen((r) => cutDepth(r) > 10);
  const cutValues = rows.filter((r) => safeNum(r.ewDiff) > 0).map((r) => cutDepth(r));

  const fillVolLt5 = sumVol((r) => r.bank > 0 && r.bank <= 5, "fillVol");
  const fillVol5To10 = sumVol((r) => r.bank > 5 && r.bank <= 10, "fillVol");
  const fillVolGt10 = sumVol((r) => r.bank > 10, "fillVol");
  const cutVolInCut = sumVol((r) => cutDepth(r) > 0, "cutVol");
  const cutVolLt5 = sumVol((r) => cutDepth(r) > 0 && cutDepth(r) <= 5, "cutVol");
  const cutVol5To10 = sumVol((r) => cutDepth(r) > 5 && cutDepth(r) <= 10, "cutVol");
  const cutVolGt10 = sumVol((r) => cutDepth(r) > 10, "cutVol");

  const blanketArea = s.formationWidthFill * s.blanketThickness;
  const preparedArea = s.formationWidthFill * (s.blanketThickness + s.preparedSubgradeThickness);
  const fmtSummary = (v) => Number(v).toLocaleString(undefined, { minimumFractionDigits: 3, maximumFractionDigits: 3 });
  const fmtM3 = (v) => fmtSummary(v);
  const fmtKm = (m) => fmtSummary(m / 1000);
  const firstCh = safeNum(rows[0]?.chainage, 0);
  const lastCh = safeNum(rows[rows.length - 1]?.chainage, firstCh);
  const totalLength = Math.max(lastCh - firstCh, 0);
  const exgTrack = parseLooseNumber(state.meta?.defaults?.trackBySideExgTrack, NaN);
  const validBridges = buildBridgeIntervals();
  const totalBridgeLen = validBridges.reduce((acc, b) => acc + safeNum(b.length), 0);
  const loopRows = Array.isArray(state.loopPlatformRows) ? state.loopPlatformRows : [];
  const totalLoopLen = loopRows.reduce((acc, r) => (
    acc + (Number.isFinite(r.loopStartCh) && Number.isFinite(r.loopEndCh) ? Math.max(r.loopEndCh - r.loopStartCh, 0) : 0)
  ), 0);
  const totalPfLen = loopRows.reduce((acc, r) => (
    acc + (Number.isFinite(r.pfStartCh) && Number.isFinite(r.pfEndCh) ? Math.max(r.pfEndCh - r.pfStartCh, 0) : 0)
  ), 0);

  const inputRows = [
    ["Formation Width (Fill)", fmtSummary(s.formationWidthFill), "m"],
    ["Formation Width (Cut)", fmtSummary(s.cuttingWidth), "m"],
    ["Blanket Thickness", fmtSummary(s.blanketThickness), "m"],
    ["Prepared Subgrade Thickness", fmtSummary(s.preparedSubgradeThickness), "m"],
    ["Berm Width", fmtSummary(s.bermWidth), "m"],
    ["Track by side of ExgTrack", Number.isFinite(exgTrack) ? fmtSummary(exgTrack) : "-", Number.isFinite(exgTrack) ? "m" : ""],
    ["Total Length", fmtSummary(totalLength), "m"],
    ["Bridge Deduction Length", fmtSummary(totalBridgeLen), "m"],
    ["Bridge Count", String(validBridges.length), ""],
    ["Loop/PF Rows", String(loopRows.length), ""],
    ["Total Loop Range", fmtSummary(totalLoopLen), "m"],
    ["Total Platform Range", fmtSummary(totalPfLen), "m"],
  ];
  const fillRows = [
    ["Fill Length", fmtKm(fillLenM), "km"],
    ["Fill Ht <5.0m", fmtKm(fillLt5M), "km"],
    ["Fill Ht 5.0m to 10.0m", fmtKm(fill5To10M), "km"],
    ["Fill Ht >10.0m", fmtKm(fillGt10M), "km"],
    ["Max Fill Ht", fmtSummary(max(fillValues)), "m"],
    ["Average Fill Height", fmtSummary(avg(fillValues)), "m"],
  ];
  const cutRows = [
    ["Cut Length", fmtKm(cutLenM), "km"],
    ["Cut depth <5.0m", fmtKm(cutLt5M), "km"],
    ["Cut 5.0m to 10.0m", fmtKm(cut5To10M), "km"],
    ["Cut >10.0m", fmtKm(cutGt10M), "km"],
    ["Max Cut depth", fmtSummary(max(cutValues)), "m"],
    ["Min Cover (Blanket+Subgrade)", fmtSummary(minCover), "m"],
  ];

  const qtyRows = [
    {
      label: "QTY in Cut",
      prepared: cutLenM * preparedArea,
      blanket: cutLenM * blanketArea,
      fill: 0,
      cut: cutVolInCut,
    },
    {
      label: "Bank <5.0m",
      prepared: fillLt5M * preparedArea,
      blanket: fillLt5M * blanketArea,
      fill: fillVolLt5,
      cut: cutVolLt5,
    },
    {
      label: "Bank 5.0m to 10.0m",
      prepared: fill5To10M * preparedArea,
      blanket: fill5To10M * blanketArea,
      fill: fillVol5To10,
      cut: cutVol5To10,
    },
    {
      label: "Bank >10.0m",
      prepared: fillGt10M * preparedArea,
      blanket: fillGt10M * blanketArea,
      fill: fillVolGt10,
      cut: cutVolGt10,
    },
  ];
  const qtyTotal = qtyRows.reduce((acc, r) => ({
    prepared: acc.prepared + r.prepared,
    blanket: acc.blanket + r.blanket,
    fill: acc.fill + r.fill,
    cut: acc.cut + r.cut,
  }), { prepared: 0, blanket: 0, fill: 0, cut: 0 });

  const renderRows = (list, colorClass = "") => list.map(([label, value, unit]) => `
    <tr class="${colorClass}">
      <td>${label}</td>
      <td>${value}</td>
      <td class="unit">${unit}</td>
    </tr>
  `).join("");
  els.resultInputBody.innerHTML = renderRows(inputRows);
  els.resultFillBody.innerHTML = renderRows(fillRows, "t-fill");
  els.resultCutBody.innerHTML = renderRows(cutRows, "t-cut");
  els.resultQtyBody.innerHTML = `
    ${qtyRows.map((r) => `
      <tr>
        <td>${r.label}</td>
        <td>${fmtM3(r.prepared)}</td>
        <td>${fmtM3(r.blanket)}</td>
        <td class="t-fill">${fmtM3(r.fill)}</td>
        <td class="t-cut">${fmtM3(r.cut)}</td>
      </tr>
    `).join("")}
    <tr>
      <td><strong>Total qty</strong></td>
      <td><strong>${fmtM3(qtyTotal.prepared)}</strong></td>
      <td><strong>${fmtM3(qtyTotal.blanket)}</strong></td>
      <td class="t-fill"><strong>${fmtM3(qtyTotal.fill)}</strong></td>
      <td class="t-cut"><strong>${fmtM3(qtyTotal.cut)}</strong></td>
    </tr>
  `;
}

function volumeCapsule(value, kind, active) {
  const v = r3(value);
  if (!active) return `<span>${v}</span>`;
  return `<span class="volume-pill ${kind}">${v}</span>`;
}

function renderTable() {
  const html = state.calcRows.map((r, idx) => {
    const structureNo = r.structureNo ? String(r.structureNo).replace(/\n/g, " ").trim() : "-";
    const station = (r.stationName || r.station) ? String(r.stationName || r.station).replace(/\n/g, " ") : "-";

    let bridgeRefs = "-";
    if (r.bridgeRefs && r.bridgeRefs.length) {
      bridgeRefs = r.bridgeRefs.map(ref => {
        const b = state.bridgeRows.find(br => String(br.bridgeNo) === String(ref));
        if (!b) return ref;
        const bColor = getBridgeStyleInfo(b.bridgeCategory, b.bridgeType);
        return `<span style="display:inline-block; padding: 2px 6px; border-radius: 4px; font-size: 0.8rem; font-weight: 600; color: ${bColor.col}; background: ${bColor.bg}; border: ${bColor.border}; white-space: nowrap; margin: 2px;">${ref}</span>`;
      }).join(" ");
    }
    const rowClass = (r.bridgeRefs && r.bridgeRefs.length) ? "bridge-row" : "";
    return `
      <tr class="${rowClass}" data-ch-index="${idx}">
        <td>${bridgeRefs}</td>
        <td>${station}</td>
        <td><button class="chainage-link theme-ch" data-cross-index="${idx}" title="Open cross-section">${(r.chainage < 0 ? "-" : "") + Math.floor(Math.abs(r.chainage) / 1000) + "+" + (Math.abs(r.chainage) % 1000).toFixed(3).replace(/(\.\d*?[1-9])0+$|\.0+$/, "$1").padStart(3, "0")}</button></td>
        <td>${r.diff ? r3(r.diff) : "—"}</td>
        <td>${r3(r.groundLevel)}</td>
        <td class="t-pro">${r3(r.proposedLevel)}</td>
        <td>${r.loopTc ? r3(r.loopTc) : "—"}</td>
        <td>${r.loopCount ? r.loopCount : "—"}</td>
        <td>${r.platformWidth ? r3(r.platformWidth) : "—"}</td>
        <td>${r.bridgeDeductLen ? r3(r.bridgeDeductLen) : "—"}</td>
        <td>${r.ewDiff ? r3(r.ewDiff) : "—"}</td>
        <td>${r.rlDiff ? r3(r.rlDiff) : "—"}</td>
        <td class="t-fill">${r.bank > 0.0001 ? r3(r.bank) : "—"}</td>
        <td class="t-cut">${r.cut > 0.0001 ? r3(r.cut) : "—"}</td>
        <td>${r.topWidth > 0.0001 ? r3(r.topWidth) : "—"}</td>
        <td class="t-fill">${r.fillArea > 0.0001 ? r3(r.fillArea) : "—"}</td>
        <td class="t-cut">${r.cutArea > 0.0001 ? r3(r.cutArea) : "—"}</td>
        <td class="t-fill">${r.fillVol > 0.0001 ? r3(r.fillVol) : "—"}</td>
        <td class="t-cut">${r.cutVol > 0.0001 ? r3(r.cutVol) : "—"}</td>
      </tr>
    `;
  }).join("");

  els.tableBody.innerHTML = html;
  
  if (typeof updateDiagnosticMinimap === 'function') {
    updateDiagnosticMinimap();
  }
}

function renderCharts() {
  if (typeof Chart === "undefined") {
    if (state.charts.lSection) {
      state.charts.lSection.destroy();
      state.charts.lSection = null;
    }
    if (state.charts.volume) {
      state.charts.volume.destroy();
      state.charts.volume = null;
    }
    return;
  }

  // Force destruction of any existing charts bound to the canvas via Chart.js globals.
  Chart.getChart("lSectionChart")?.destroy();
  Chart.getChart("volumeChart")?.destroy();

  if (!state.calcRows.length) {
    if (state.charts.lSection) {
      state.charts.lSection.destroy();
      state.charts.lSection = null;
    }
    if (state.charts.volume) {
      state.charts.volume.destroy();
      state.charts.volume = null;
    }
    return;
  }
  const labels = state.calcRows.map((r) => r3((r.chainage - state.calcRows[0].chainage) / 1000));
  const gl = state.calcRows.map((r) => Number(r3(r.groundLevel)));
  const fl = state.calcRows.map((r) => Number(r3(r.proposedLevel)));
  const diff = state.calcRows.map((r) => Number(r3(r.rlDiff)));

  // Lovable design colors
  const groundColor = "hsl(354, 72%, 55%)";   // red
  const proposedColor = "hsl(233, 100%, 57%)"; // blue
  const gridColor = "hsla(222, 20%, 25%, 0.3)";
  const tickColor = "hsl(215, 20%, 55%)";

  if (state.charts.lSection) state.charts.lSection.destroy();

  const lCtx = els.lSectionChart.getContext("2d");
  // Ground gradient fill
  const groundGrad = lCtx.createLinearGradient(0, 0, 0, els.lSectionChart.height);
  groundGrad.addColorStop(0, "hsla(354, 72%, 55%, 0.3)");
  groundGrad.addColorStop(1, "hsla(354, 72%, 55%, 0)");
  // Proposed gradient fill
  const proposedGrad = lCtx.createLinearGradient(0, 0, 0, els.lSectionChart.height);
  proposedGrad.addColorStop(0, "hsla(233, 100%, 57%, 0.3)");
  proposedGrad.addColorStop(1, "hsla(233, 100%, 57%, 0)");

  state.charts.lSection = new Chart(els.lSectionChart, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Ground RL",
          data: gl,
          borderColor: groundColor,
          backgroundColor: groundGrad,
          tension: 0.35,
          fill: true,
          pointRadius: 0,
          pointBackgroundColor: groundColor,
          pointBorderWidth: 0,
          pointHoverRadius: 4,
          pointHoverBackgroundColor: "hsl(222, 47%, 6%)",
          pointHoverBorderColor: groundColor,
          pointHoverBorderWidth: 1.5,
          borderWidth: 1.5,
        },
        {
          label: "Proposed RL",
          data: fl,
          borderColor: proposedColor,
          backgroundColor: proposedGrad,
          tension: 0.35,
          fill: true,
          pointRadius: 0,
          pointBackgroundColor: proposedColor,
          pointBorderWidth: 0,
          pointHoverRadius: 4,
          pointHoverBackgroundColor: "hsl(222, 47%, 6%)",
          pointHoverBorderColor: proposedColor,
          pointHoverBorderWidth: 1.5,
          borderWidth: 1.5,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: {
          labels: {
            boxWidth: 12,
            boxHeight: 2,
            usePointStyle: false,
            color: tickColor,
            font: { size: 10, family: "Outfit", weight: 500 },
            padding: 16,
          },
          position: "top",
          align: "end",
        },
        tooltip: {
          backgroundColor: "hsla(222, 40%, 12%, 0.9)",
          titleColor: "#fff",
          bodyColor: "hsl(215, 20%, 75%)",
          borderColor: "hsla(222, 20%, 30%, 0.5)",
          borderWidth: 1,
          padding: 10,
          titleFont: { family: "Outfit", weight: 600, size: 12 },
          bodyFont: { family: "Outfit", size: 11 },
          callbacks: {
            afterBody: (ctx) => {
              const i = ctx[0].dataIndex;
              return `RL Difference: ${r3(diff[i])} m`;
            },
          },
        },
        zoom: {
          pan: { enabled: true, mode: "x" },
          zoom: {
            wheel: { enabled: true, speed: 0.1 },
            pinch: { enabled: true },
            mode: "x",
          },
        },
      },
      scales: {
        x: {
          title: { display: false },
          ticks: { color: tickColor, font: { size: 10, family: "Outfit" }, maxTicksLimit: 10 },
          grid: { color: gridColor, drawBorder: false },
          border: { color: gridColor },
        },
        y: {
          title: { display: false },
          ticks: { color: tickColor, font: { size: 10, family: "Outfit" } },
          grid: { color: gridColor, drawBorder: false },
          border: { color: gridColor },
        },
      },
    },
  });

  // Double-click to reset zoom
  els.lSectionChart.ondblclick = () => {
    state.charts.lSection?.resetZoom();
  };

  if (state.charts.volume) state.charts.volume.destroy();
  state.charts.volume = new Chart(els.volumeChart, {
    data: {
      labels,
      datasets: [
        {
          type: "bar",
          label: "Fill Vol (m³)",
          data: state.calcRows.map((r) => Number(r3(r.fillVol))),
          backgroundColor: "hsla(152, 69%, 37%, 0.55)",
          borderRadius: 4,
        },
        {
          type: "bar",
          label: "Cut Vol (m³)",
          data: state.calcRows.map((r) => -Number(r3(r.cutVol))),
          backgroundColor: "hsla(354, 72%, 55%, 0.55)",
          borderRadius: 4,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "top",
          align: "end",
          labels: {
            boxWidth: 12,
            color: tickColor,
            font: { size: 10, family: "Outfit", weight: 500 },
          },
        },
        tooltip: {
          backgroundColor: "hsla(222, 40%, 12%, 0.9)",
          titleColor: "#fff",
          bodyColor: "hsl(215, 20%, 75%)",
          borderColor: "hsla(222, 20%, 30%, 0.5)",
          borderWidth: 1,
          padding: 10,
          titleFont: { family: "Outfit", weight: 600, size: 12 },
          bodyFont: { family: "Outfit", size: 11 },
        },
        zoom: {
          pan: { enabled: true, mode: "x" },
          zoom: {
            wheel: { enabled: true, speed: 0.1 },
            pinch: { enabled: true },
            mode: "x",
          },
        },
      },
      scales: {
        x: {
          ticks: { maxTicksLimit: 9, color: tickColor, font: { size: 10, family: "Outfit" } },
          grid: { color: gridColor, drawBorder: false },
          border: { color: gridColor },
        },
        y: {
          title: { display: false },
          ticks: { color: tickColor, font: { size: 10, family: "Outfit" } },
          grid: { color: gridColor, drawBorder: false },
          border: { color: gridColor },
        },
      },
    },
  });

  // Double-click to reset zoom
  els.volumeChart.ondblclick = () => {
    state.charts.volume?.resetZoom();
  };
}

function renderRollDiagram() {
  const canvas = els.rollDiagramCanvas;
  const wrap = els.rollDiagramWrap;
  const empty = els.rollDiagramEmpty;
  if (!canvas) return;

  const noData = !state.calcRows || state.calcRows.length === 0;
  if (empty) { empty.style.display = noData ? "flex" : "none"; }
  if (wrap) { wrap.style.display = noData ? "none" : "block"; }
  if (noData) return;

  const rows = state.calcRows;
  const minCh = rows[0].chainage;
  const maxCh = rows[rows.length - 1].chainage;
  const totalL = Math.max(maxCh - minCh, 1);
  const rollFilter = state.settings.rollFilter || "all";
  const lightMode = window._printModeLight === true;

  const PAD_L = 60, PAD_R = 40, PAD_T = 70, PAD_B = 45; // Reduced paddings
  const MAX_SAFE_W = 32000;
  const maxScaleLimit = (MAX_SAFE_W - PAD_L - PAD_R) / (totalL * 0.4);

  const baseScale = Math.max(0.05, Math.min(maxScaleLimit, window._planScale || 1));
  window._planScale = baseScale; // Update state inline

  const PX_PER_M_X = 0.4 * baseScale;
  const PX_PER_M_Y = 2.8; // Fixed vertical scale, height remains constant

  const maxHalfW = rows.reduce((m, r) => {
    const w = r.bank > 0 ? r.fillBottom : (r.cut > 0 ? r.cutBottom : (r.effectiveFormationWidth || 0));
    return Math.max(m, w / 2);
  }, 0);
  const loopMaxTc = (state.loopPlatformRows || []).reduce((m, lp) => Math.max(m, Math.abs(safeNum(lp.tc, 0))), 0);
  const loopPxH = (loopMaxTc + 10) * PX_PER_M_Y;
  const bodyHalf = Math.max(maxHalfW * PX_PER_M_Y, 28); // Reduced min from 40
  const canvasH = Math.ceil(PAD_T + bodyHalf * 2 + loopPxH + PAD_B);
  const canvasW = Math.ceil(PAD_L + totalL * PX_PER_M_X + PAD_R);
  const centerY = PAD_T + bodyHalf + loopPxH * 0.5;

  const exportScale = lightMode ? 2 : 1;
  canvas.width = Math.ceil(canvasW * exportScale);
  canvas.height = Math.ceil(canvasH * exportScale);
  canvas.style.width = `${canvasW}px`;
  canvas.style.height = `${canvasH}px`;
  const ctx = canvas.getContext("2d");
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  if (exportScale !== 1) ctx.scale(exportScale, exportScale);
  ctx.clearRect(0, 0, canvasW, canvasH);

  function getX(ch) {
    return Number.isFinite(ch) ? PAD_L + (ch - minCh) * PX_PER_M_X : null;
  }

  // ── Background ─────────────────────────────────────────────────────────
  const bgGrad = ctx.createLinearGradient(0, 0, 0, canvasH);
  if (lightMode) {
    bgGrad.addColorStop(0, "#f8fafc");
    bgGrad.addColorStop(1, "#ffffff");
  } else {
    bgGrad.addColorStop(0, "#0d1117");
    bgGrad.addColorStop(1, "#141d2e");
  }
  ctx.fillStyle = bgGrad;
  ctx.fillRect(0, 0, canvasW, canvasH);

  // ── X & Y Grid lines (Major & Minor) ───────────────────────────────────
  const minSpacPx = 25; // Minimum px spacing for minor grid lines

  // X-axis divisions (Chainage)
  let xMajor = 1000, xMinor = 500;
  if (PX_PER_M_X * 10 >= minSpacPx) { xMajor = 50; xMinor = 10; }
  else if (PX_PER_M_X * 25 >= minSpacPx) { xMajor = 100; xMinor = 25; }
  else if (PX_PER_M_X * 50 >= minSpacPx) { xMajor = 250; xMinor = 50; }
  else if (PX_PER_M_X * 100 >= minSpacPx) { xMajor = 500; xMinor = 100; }
  else if (PX_PER_M_X * 500 >= minSpacPx) { xMajor = 2000; xMinor = 500; }
  else { xMajor = 10000; xMinor = 2000; }

  ctx.lineWidth = 1;
  const startX = Math.ceil(minCh / xMinor) * xMinor;
  for (let ch = startX; ch <= maxCh; ch += xMinor) {
    const gx = getX(ch);
    if (gx == null) continue;
    const isMajor = (ch % xMajor === 0);
    ctx.strokeStyle = lightMode
      ? (isMajor ? "rgba(15,23,42,0.12)" : "rgba(15,23,42,0.05)")
      : (isMajor ? "rgba(255,255,255,0.12)" : "rgba(255,255,255,0.035)");
    ctx.beginPath(); ctx.moveTo(gx, PAD_T); ctx.lineTo(gx, canvasH - PAD_B); ctx.stroke();
  }

  // Y-axis divisions (Width from Centerline)
  let yMajor = 20, yMinor = 5;
  if (PX_PER_M_Y * 2 >= minSpacPx) { yMajor = 10; yMinor = 2; }
  else if (PX_PER_M_Y * 5 >= minSpacPx) { yMajor = 20; yMinor = 5; }
  else if (PX_PER_M_Y * 10 >= minSpacPx) { yMajor = 50; yMinor = 10; }
  else { yMajor = 100; yMinor = 20; }

  const axisExtentM = Math.max(yMinor, Math.ceil(maxHalfW / yMinor) * yMinor);
  ctx.fillStyle = lightMode ? "rgba(15,23,42,0.55)" : "rgba(255,255,255,0.4)";
  ctx.font = `9px Outfit,sans-serif`;
  ctx.textAlign = "left";

  for (let ym = 0; ym <= axisExtentM; ym += yMinor) {
    if (ym === 0) continue; // Skip centerline
    const isMajor = (ym % yMajor === 0);
    ctx.strokeStyle = lightMode
      ? (isMajor ? "rgba(15,23,42,0.12)" : "rgba(15,23,42,0.05)")
      : (isMajor ? "rgba(255,255,255,0.12)" : "rgba(255,255,255,0.035)");
    const dy = ym * PX_PER_M_Y;

    // Top side (negative Y from center)
    if (centerY - dy >= PAD_T) {
      ctx.beginPath(); ctx.moveTo(PAD_L, centerY - dy); ctx.lineTo(canvasW - PAD_R, centerY - dy); ctx.stroke();
      if (isMajor) ctx.fillText(`${ym}m`, PAD_L + 6, centerY - dy - 4);
    }
    // Bottom side (positive Y from center)
    if (centerY + dy <= canvasH - PAD_B) {
      ctx.beginPath(); ctx.moveTo(PAD_L, centerY + dy); ctx.lineTo(canvasW - PAD_R, centerY + dy); ctx.stroke();
      if (isMajor) ctx.fillText(`${ym}m`, PAD_L + 6, centerY + dy - 4);
    }
  }

  // ── Centreline ─────────────────────────────────────────────────────────
  ctx.setLineDash([6, 6]);
  ctx.strokeStyle = lightMode ? "rgba(71,85,105,0.45)" : "rgba(255,255,255,0.25)";
  ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(getX(minCh), centerY); ctx.lineTo(getX(maxCh), centerY); ctx.stroke();
  ctx.setLineDash([]);

  const stationRanges = (() => {
    if (rollFilter !== "stations") return [];
    const ranges = [];
    const stationMap = new Map();
    (state.loopPlatformRows || []).forEach((lp) => {
      const key = normalizeStationKey(lp.station);
      if (!key) return;
      if (!stationMap.has(key)) stationMap.set(key, { min: Number.POSITIVE_INFINITY, max: Number.NEGATIVE_INFINITY });
      const entry = stationMap.get(key);
      const loopStart = parseChainage(lp.loopStartCh);
      const loopEnd = parseChainage(lp.loopEndCh);
      const pfStart = parseChainage(lp.pfStartCh);
      const pfEnd = parseChainage(lp.pfEndCh);
      if (Number.isFinite(loopStart)) entry.min = Math.min(entry.min, loopStart);
      if (Number.isFinite(pfStart)) entry.min = Math.min(entry.min, pfStart);
      if (Number.isFinite(loopEnd)) entry.max = Math.max(entry.max, loopEnd);
      if (Number.isFinite(pfEnd)) entry.max = Math.max(entry.max, pfEnd);
    });
    stationMap.forEach((entry) => {
      if (Number.isFinite(entry.min) && Number.isFinite(entry.max) && entry.max > entry.min) {
        ranges.push({ start: entry.min, end: entry.max });
      }
    });
    ranges.sort((a, b) => a.start - b.start);
    const merged = [];
    ranges.forEach((r) => {
      const last = merged[merged.length - 1];
      if (!last || r.start > last.end) merged.push({ ...r });
      else last.end = Math.max(last.end, r.end);
    });
    return merged;
  })();

  const mainlineSegments = rollFilter === "main"
    ? [{ start: minCh, end: maxCh }]
    : (rollFilter === "stations" ? stationRanges : []);

  if (mainlineSegments.length) {
    ctx.strokeStyle = "rgba(239,68,68,0.95)";
    ctx.lineWidth = 2.5 * baseScale;
    mainlineSegments.forEach((seg) => {
      const x1 = getX(seg.start);
      const x2 = getX(seg.end);
      if (x1 == null || x2 == null) return;
      ctx.beginPath(); ctx.moveTo(x1, centerY); ctx.lineTo(x2, centerY); ctx.stroke();
    });
  }

  // ── Bridges ─────────────────────────────────────────────────────────────
  const bridges = buildBridgeIntervals();
  function isOnBridge(ch) {
    return bridges.some(b => ch >= b.startChainage && ch <= b.endChainage);
  }

  // ── Toe width — per-segment coloured trapezoids ─────────────────────────
  for (let i = 1; i < rows.length; i++) {
    const r0 = rows[i - 1], r1 = rows[i];
    const midCh = (r0.chainage + r1.chainage) / 2;
    if (isOnBridge(midCh)) continue;

    const x0 = getX(r0.chainage), x1g = getX(r1.chainage);
    if (x0 == null || x1g == null) continue;
    const w0 = r0.bank > 0 ? r0.fillBottom : (r0.cut > 0 ? r0.cutBottom : (r0.effectiveFormationWidth || 0));
    const w1 = r1.bank > 0 ? r1.fillBottom : (r1.cut > 0 ? r1.cutBottom : (r1.effectiveFormationWidth || 0));
    const h0 = (w0 / 2) * PX_PER_M_Y, h1 = (w1 / 2) * PX_PER_M_Y;
    const isFill = r1.bank > 0.001, isCut = r1.cut > 0.001;
  const fillCol = isFill
    ? (lightMode ? "rgba(34,197,94,0.2)" : "rgba(34,139,69,0.28)")
    : isCut
      ? (lightMode ? "rgba(239,68,68,0.2)" : "rgba(180,44,50,0.28)")
      : (lightMode ? "rgba(148,163,184,0.18)" : "rgba(100,110,130,0.18)");
  const strokeCol = isFill
    ? (lightMode ? "rgba(22,163,74,0.9)" : "rgba(34,200,80,0.75)")
    : isCut
      ? (lightMode ? "rgba(220,38,38,0.9)" : "rgba(240,70,70,0.75)")
      : (lightMode ? "rgba(100,116,139,0.6)" : "rgba(140,150,170,0.45)");
    // Top edge
    ctx.strokeStyle = strokeCol; ctx.lineWidth = 1.5;
    ctx.beginPath(); ctx.moveTo(x0, centerY - h0); ctx.lineTo(x1g, centerY - h1); ctx.stroke();
    // Bottom edge
    ctx.beginPath(); ctx.moveTo(x0, centerY + h0); ctx.lineTo(x1g, centerY + h1); ctx.stroke();
    // Fill trapezoid
    ctx.fillStyle = fillCol;
    ctx.beginPath(); ctx.moveTo(x0, centerY - h0); ctx.lineTo(x1g, centerY - h1);
    ctx.lineTo(x1g, centerY + h1); ctx.lineTo(x0, centerY + h0); ctx.closePath(); ctx.fill();
  }

  // ── Formation width dash ────────────────────────────────────────────────
  const fwH = (state.settings.formationWidthFill || 0) / 2 * PX_PER_M_Y;
  if (fwH > 0) {
    ctx.strokeStyle = lightMode ? "rgba(37,99,235,0.35)" : "rgba(120,200,255,0.35)"; ctx.lineWidth = 1; ctx.setLineDash([4, 4]);
    ctx.beginPath(); ctx.moveTo(getX(minCh), centerY - fwH); ctx.lineTo(getX(maxCh), centerY - fwH); ctx.stroke();
    ctx.beginPath(); ctx.moveTo(getX(minCh), centerY + fwH); ctx.lineTo(getX(maxCh), centerY + fwH); ctx.stroke();
    ctx.setLineDash([]);
  }

  // ── Draw Bridge Rectangles ──────────────────────────────────────────────
  for (const b of bridges) {
    const bx1 = getX(b.startChainage), bx2 = getX(b.endChainage);
    if (bx1 == null || bx2 == null) continue;
    const bHalf = 20 * PX_PER_M_Y;
    const bW = Math.max(bx2 - bx1, 3 * baseScale);
    ctx.fillStyle = lightMode ? "rgba(59,130,246,0.25)" : "rgba(37,99,235,0.50)";
    ctx.strokeStyle = lightMode ? "rgba(37,99,235,0.9)" : "rgba(99,163,255,0.9)";
    ctx.lineWidth = 1.5;
    ctx.fillRect(bx1, centerY - bHalf, bW, bHalf * 2);
    ctx.strokeRect(bx1, centerY - bHalf, bW, bHalf * 2);
    ctx.strokeStyle = lightMode ? "rgba(37,99,235,0.25)" : "rgba(99,163,255,0.3)"; ctx.lineWidth = 0.7;
    for (let hx = bx1; hx < bx2; hx += 8 * baseScale) {
      ctx.beginPath(); ctx.moveTo(hx, centerY - bHalf); ctx.lineTo(hx + 8 * baseScale, centerY + bHalf); ctx.stroke();
    }
    ctx.fillStyle = lightMode ? "#1d4ed8" : "#93c5fd"; ctx.font = `bold ${Math.max(9, 11 * baseScale)}px Outfit,sans-serif`;
    ctx.textAlign = "left";
    ctx.fillText(b.bridgeNo, bx1 + 2, centerY - bHalf - 5);
  }

  // ── Curves ──────────────────────────────────────────────────────────────
  if (state.curveRows && state.curveRows.length) {
    ctx.lineWidth = 2 * baseScale;
    state.curveRows.forEach((c, ci) => {
      if (!Number.isFinite(c.chainage) || c.chainage < minCh || c.chainage > maxCh) return;
      const cx1 = getX(c.chainage);
      if (cx1 == null) return;
      const len = safeNum(c.length);
      const cx2 = getX(Math.min(c.chainage + (len > 5 ? len : 200), maxCh));
      if (cx2 == null) return;
      const rad = safeNum(c.radius);
      const dir = (ci % 2 === 0) ? -1 : 1;
      const bulge = 25 * baseScale;
      ctx.strokeStyle = lightMode ? "rgba(217,119,6,0.85)" : "rgba(252,211,77,0.85)"; ctx.fillStyle = "transparent";
      ctx.beginPath();
      ctx.moveTo(cx1, centerY);
      ctx.quadraticCurveTo((cx1 + cx2) / 2, centerY + dir * bulge * 2, cx2, centerY);
      ctx.stroke();
      ctx.fillStyle = lightMode ? "#b45309" : "#fde68a"; ctx.font = `${Math.max(9, 10 * baseScale)}px Outfit,sans-serif`;
      ctx.textAlign = "center";
      const label = (c.curve || "Curve") + (rad > 0 ? ` R=${r3(rad)}m` : "");
      ctx.fillText(label, (cx1 + cx2) / 2, centerY + dir * (bulge * 2 + 14));
    });
  }

  // ── Loops & Platforms (Station Layout) ──────────────────────────────────
  if (state.loopPlatformRows && state.loopPlatformRows.length && rollFilter !== "main") {
    const stationMap = new Map();
    state.loopPlatformRows.forEach((lp) => {
      const key = normalizeStationKey(lp.station);
      if (!key) return;
      if (!stationMap.has(key)) {
        stationMap.set(key, { station: lp.station, rows: [], minCh: Number.POSITIVE_INFINITY, maxCh: Number.NEGATIVE_INFINITY, csb: Number.NaN });
      }
      const entry = stationMap.get(key);
      entry.rows.push(lp);
      const loopStart = parseChainage(lp.loopStartCh);
      const loopEnd = parseChainage(lp.loopEndCh);
      const pfStart = parseChainage(lp.pfStartCh);
      const pfEnd = parseChainage(lp.pfEndCh);
      if (Number.isFinite(loopStart)) entry.minCh = Math.min(entry.minCh, loopStart);
      if (Number.isFinite(pfStart)) entry.minCh = Math.min(entry.minCh, pfStart);
      if (Number.isFinite(loopEnd)) entry.maxCh = Math.max(entry.maxCh, loopEnd);
      if (Number.isFinite(pfEnd)) entry.maxCh = Math.max(entry.maxCh, pfEnd);
      if (!Number.isFinite(entry.csb)) {
        const csb = parseChainage(lp.csb);
        if (Number.isFinite(csb)) entry.csb = csb;
      }
    });

    stationMap.forEach((entry) => {
      if (!Number.isFinite(entry.minCh) || !Number.isFinite(entry.maxCh)) return;
      const midCh = Number.isFinite(entry.csb) ? entry.csb : ((entry.minCh + entry.maxCh) / 2);
      const layout = buildStationSequenceLayout(midCh, entry.station, 5.3, { useRanges: false });
      if (!layout) return;
      layout.trackItems.forEach((item) => {
        const offset = layout.offsetByItem.get(item);
        if (!Number.isFinite(offset)) return;
        let segStart = parseChainage(item.row.loopStartCh);
        let segEnd = parseChainage(item.row.loopEndCh);
        if (!Number.isFinite(segStart) || !Number.isFinite(segEnd) || segEnd <= segStart) {
          segStart = entry.minCh;
          segEnd = entry.maxCh;
        }
        const x1 = getX(segStart);
        const x2 = getX(segEnd);
        if (x1 == null || x2 == null) return;
        const y = centerY - (offset * PX_PER_M_Y);
        ctx.strokeStyle = item.isMain ? "rgba(239,68,68,0.95)" : (lightMode ? "rgba(14,116,144,0.85)" : "rgba(34,211,238,0.85)");
        ctx.lineWidth = item.isMain ? 2.5 * baseScale : 2 * baseScale;
        ctx.beginPath(); ctx.moveTo(x1, y); ctx.lineTo(x2, y); ctx.stroke();
      });

      layout.platformItems.forEach((item) => {
        const pfStart = parseChainage(item.row.pfStartCh);
        const pfEnd = parseChainage(item.row.pfEndCh);
        if (!Number.isFinite(pfStart) || !Number.isFinite(pfEnd) || pfEnd <= pfStart) return;
        const px1 = getX(pfStart), px2 = getX(pfEnd);
        if (px1 == null || px2 == null) return;
        const pfW = safeNum(item.row.pfWidth, 10);
        const pfH = pfW * PX_PER_M_Y * 0.5;
        const ordered = layout.orderedItems || [];
        const idx = ordered.indexOf(item);
        const findNeighborTrack = (startIndex, step) => {
          for (let i = startIndex + step; i >= 0 && i < ordered.length; i += step) {
            if (ordered[i].kind === "track") return ordered[i];
          }
          return null;
        };
        const before = idx >= 0 ? findNeighborTrack(idx, -1) : null;
        const after = idx >= 0 ? findNeighborTrack(idx, 1) : null;
        let offset = 0;
        if (before && after) {
          const offA = layout.offsetByItem.get(before);
          const offB = layout.offsetByItem.get(after);
          if (Number.isFinite(offA) && Number.isFinite(offB)) {
            offset = (offA + offB) / 2;
          }
        } else {
          const anchor = before || after || layout.refMain;
          const anchorOffset = layout.offsetByItem.get(anchor) || 0;
          const side = item.side || (anchorOffset >= 0 ? "Left" : "Right");
          const pushM = (pfW / 2) + 2;
          offset = anchorOffset + (side === "Left" ? pushM : -pushM);
        }
        const y = centerY - (offset * PX_PER_M_Y);
        const pfTop = y - (pfH / 2);
        ctx.fillStyle = lightMode ? "rgba(239,68,68,0.2)" : "rgba(239,68,68,0.45)";
        ctx.strokeStyle = lightMode ? "rgba(220,38,38,0.7)" : "rgba(252,165,165,0.8)";
        ctx.lineWidth = 1.5;
        ctx.fillRect(px1, pfTop, Math.max(px2 - px1, 4), pfH);
        ctx.strokeRect(px1, pfTop, Math.max(px2 - px1, 4), pfH);
      });

      const labelX = getX(entry.minCh);
      if (labelX != null) {
      ctx.fillStyle = lightMode ? "#0f172a" : "#67e8f9";
        ctx.font = `bold ${Math.max(9, 10 * baseScale)}px Outfit,sans-serif`;
        ctx.textAlign = "left";
        ctx.fillText("\u25CE " + (entry.station || "Station"), labelX, centerY - (fwH || 20) - 18);
      }
    });
  }

  // ── Scale Ruler ─────────────────────────────────────────────────────────
  const rulerY = canvasH - PAD_B + 16;
  const tkInt = totalL >= 50000 ? 5000 : totalL >= 20000 ? 2000 : totalL >= 5000 ? 1000 : 500;
  const startTk = Math.ceil(minCh / tkInt) * tkInt;
  ctx.strokeStyle = lightMode ? "rgba(71,85,105,0.45)" : "rgba(255,255,255,0.35)"; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(getX(minCh), rulerY); ctx.lineTo(getX(maxCh), rulerY); ctx.stroke();
  ctx.fillStyle = lightMode ? "rgba(15,23,42,0.7)" : "rgba(255,255,255,0.6)"; ctx.font = `9px Outfit,sans-serif`; ctx.textAlign = "center";
  for (let ch = startTk; ch <= maxCh; ch += tkInt) {
    const tx = getX(ch);
    if (tx == null) continue;
    ctx.beginPath(); ctx.moveTo(tx, rulerY - 4); ctx.lineTo(tx, rulerY + 4); ctx.stroke();
    ctx.fillText(ch >= 1000 ? `${r3(ch / 1000)} km` : `${ch} m`, tx, rulerY + 16);
  }
  ctx.textAlign = "left"; ctx.fillStyle = lightMode ? "rgba(15,23,42,0.5)" : "rgba(255,255,255,0.4)"; ctx.font = "10px Outfit,sans-serif";
  ctx.fillText(`CH ${minCh}m`, PAD_L, rulerY + 30);
  ctx.textAlign = "right"; ctx.fillText(`CH ${maxCh}m`, canvasW - PAD_R, rulerY + 30);

  // ── Legend ───────────────────────────────────────────────────────────────
  const legend = [
    { col: "rgba(34,200,80,0.75)", lbl: "Filling" },
    { col: "rgba(240,70,70,0.75)", lbl: "Cutting" },
    { col: "rgba(120,200,255,0.5)", lbl: "Formation" },
    { col: "rgba(99,163,255,0.9)", lbl: "Bridge" },
    { col: "rgba(252,211,77,0.85)", lbl: "Curve" },
    { col: "rgba(34,211,238,0.85)", lbl: "Loop/Station" },
    { col: "rgba(252,165,165,0.8)", lbl: "Platform" },
  ];
  ctx.font = `10px Outfit,sans-serif`;
  let lx = PAD_L;
  for (const item of legend) {
    ctx.fillStyle = item.col; ctx.fillRect(lx, 10, 12, 12);
    ctx.fillStyle = lightMode ? "rgba(15,23,42,0.7)" : "rgba(255,255,255,0.75)"; ctx.textAlign = "left";
    ctx.fillText(item.lbl, lx + 15, 21);
    lx += 15 + ctx.measureText(item.lbl).width + 14;
    if (lx > canvasW - 100) break;
  }

  // ── Tooltip Interaction ──────────────────────────────────────────────────
  const tooltip = document.getElementById("rollTooltip");
  ensureRollCrosshair("rollCrosshair", "rollDiagramWrap");
  ensureRollCrosshair("sideRollCrosshair", "sideViewWrap");

  if (tooltip) {
    // Store the last hovered row index so click can use it
    let _hoveredRowIdx = -1;

    canvas.onmousemove = (e) => {
      const rect = canvas.getBoundingClientRect();
      const scaleX = canvas.width / rect.width;
      const visualX = e.clientX - rect.left;
      const mouseX = visualX * scaleX;

      const chHit = minCh + (mouseX - PAD_L) / PX_PER_M_X;
      if (chHit >= minCh && chHit <= maxCh && mouseX >= PAD_L && mouseX <= canvasW - PAD_R) {
        // Find closest row
        const { row: closest, index: closestIdx } = findClosestCalcRowByChainage(rows, chHit);
        _hoveredRowIdx = closestIdx;

        // Construct Tooltip content
        const isOnBr = bridges.some(b => closest.chainage >= b.startChainage && closest.chainage <= b.endChainage);
        const w0 = closest.bank > 0 ? closest.fillBottom : (closest.cut > 0 ? closest.cutBottom : (closest.effectiveFormationWidth || 0));
        let content = `
          <div style="font-weight:700; border-bottom:1px solid hsla(222,20%,50%,0.3); padding-bottom:4px; margin-bottom:6px;">
            Chainage: ${closest.chainage.toFixed(3)} m
          </div>
          <div style="display:grid; grid-template-columns:auto auto; gap:4px 12px; font-size:0.8rem;">
            <span style="color:var(--muted)">Ground RL:</span> <span style="font-weight:600">${r3(closest.groundLevel)}</span>
            <span style="color:var(--muted)">Formation RL:</span> <span style="font-weight:600">${r3(closest.proposedLevel)}</span>
            <span style="color:var(--muted)">Toe Dist (Side):</span> <span style="font-weight:600">${isOnBr ? '0 (Bridge)' : r3(w0 / 2)} m</span>
            <span style="color:var(--green)">Fill Vol:</span> <span style="font-weight:600">${r3(closest.fillVol)} m³</span>
            <span style="color:var(--red)">Cut Vol:</span> <span style="font-weight:600">${r3(closest.cutVol)} m³</span>
          </div>
          <div style="margin-top:6px; padding-top:4px; border-top:1px solid hsla(222,20%,50%,0.2); font-size:0.7rem; color:hsla(233,100%,75%,0.9); font-weight:600; text-align:center;">
            Click to view Cross‑Section
          </div>
        `;

        tooltip.innerHTML = content;
        tooltip.style.display = 'block';
        tooltip.style.left = (e.pageX + 15) + 'px';
        tooltip.style.top = (e.pageY + 15) + 'px';

        // Ensure tooltip doesn't go off-screen
        const tRect = tooltip.getBoundingClientRect();
        if (tRect.right > window.innerWidth) {
          tooltip.style.left = (e.pageX - tRect.width - 15) + 'px';
        }
        if (tRect.bottom > window.innerHeight) {
          tooltip.style.top = (e.pageY - tRect.height - 15) + 'px';
        }

        setRollCrosshairPosition("rollCrosshair", canvas, visualX);
        const sideCanvas = els.sideViewCanvas;
        if (sideCanvas) {
          const closestX = getX(closest.chainage);
          if (closestX != null) {
            const sideVisualX = (closestX / sideCanvas.width) * sideCanvas.clientWidth;
            setRollCrosshairPosition("sideRollCrosshair", sideCanvas, sideVisualX);
          }
        }

        canvas.style.cursor = 'crosshair';
      } else {
        tooltip.style.display = 'none';
        hideRollCrosshairs();
        canvas.style.cursor = 'grab';
        _hoveredRowIdx = -1;
      }
    };

    canvas.onmouseleave = () => {
      tooltip.style.display = 'none';
      hideRollCrosshairs();
      canvas.style.cursor = 'grab';
      _hoveredRowIdx = -1;
    };

    // Distinguish click vs drag: only open cross-section if mouse didn't move
    let _downX = 0, _downY = 0;
    canvas.addEventListener('mousedown', (e) => {
      _downX = e.clientX;
      _downY = e.clientY;
    });

    canvas.addEventListener('mouseup', (e) => {
      const dx = Math.abs(e.clientX - _downX);
      const dy = Math.abs(e.clientY - _downY);
      // Only treat as a click if mouse moved less than 5px (not a drag)
      if (dx < 5 && dy < 5 && _hoveredRowIdx >= 0 && _hoveredRowIdx < state.calcRows.length) {
        openCrossSectionByIndex(_hoveredRowIdx);
        tooltip.style.display = 'none';
        hideRollCrosshairs();
      }
    });
  }
}

function ensureRollCrosshair(crosshairId, wrapId) {
  let crosshair = document.getElementById(crosshairId);
  const wrap = document.getElementById(wrapId);
  if (!crosshair && wrap) {
    crosshair = document.createElement("div");
    crosshair.id = crosshairId;
    crosshair.style.cssText = "display:none;position:absolute;top:0;width:1px;background:hsla(354, 72%, 55%, 0.85);pointer-events:none;z-index:90;box-shadow:0 0 3px hsla(354, 72%, 55%, 0.4);";
    wrap.appendChild(crosshair);
  }
  return crosshair;
}

function setRollCrosshairPosition(crosshairId, canvas, visualX) {
  const crosshair = ensureRollCrosshair(crosshairId, crosshairId === "rollCrosshair" ? "rollDiagramWrap" : "sideViewWrap");
  if (!crosshair || !canvas) return;
  crosshair.style.display = "block";
  crosshair.style.left = `${canvas.offsetLeft + visualX}px`;
  crosshair.style.height = `${canvas.clientHeight}px`;
}

function hideRollCrosshairs() {
  ["rollCrosshair", "sideRollCrosshair"].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.style.display = "none";
  });
}

function findClosestCalcRowByChainage(rows, chainage) {
  if (!Array.isArray(rows) || !rows.length || !Number.isFinite(chainage)) {
    return { row: null, index: -1 };
  }
  let closest = rows[0];
  let closestIdx = 0;
  let minDist = Infinity;
  for (let i = 0; i < rows.length; i += 1) {
    const d = Math.abs(rows[i].chainage - chainage);
    if (d < minDist) {
      minDist = d;
      closest = rows[i];
      closestIdx = i;
    }
  }
  return { row: closest, index: closestIdx };
}

function renderSideView() {
  const canvas = els.sideViewCanvas;
  if (!canvas) return;
  if (!state.calcRows || state.calcRows.length === 0) {
    canvas.width = 0; canvas.height = 0; return;
  }

  const rows = state.calcRows;
  const minCh = rows[0].chainage;
  const maxCh = rows[rows.length - 1].chainage;
  const totalL = Math.max(maxCh - minCh, 1);
  const lightMode = window._printModeLight === true;

  const PAD_L = 72;   // room for Y-axis labels
  const PAD_R = 30;
  const PAD_T = 60;   // legend + title
  const PAD_B = 50;   // chainage labels

  const MAX_SAFE_W = 32000;
  const maxScaleLimit = (MAX_SAFE_W - PAD_L - PAD_R) / (totalL * 0.4);

  const baseScale = Math.max(0.05, Math.min(maxScaleLimit, window._sideScale || 1));
  window._sideScale = baseScale; // Update state inline

  const PX_PER_M_X = 0.4 * baseScale;
  const PX_PER_M_Y = 1.5; // Fixed vertical scale

  // Elevation range from calcRows
  const allGLs = rows.map(r => safeNum(r.groundLevel));
  const allFLs = rows.map(r => safeNum(r.proposedLevel));
  const minElev = Math.min(...allGLs, ...allFLs) - 3;
  const maxElev = Math.max(...allGLs, ...allFLs) + 6;
  const elevRange = Math.max(maxElev - minElev, 1);

  const canvasW = Math.ceil(PAD_L + totalL * PX_PER_M_X + PAD_R);
  const bodyH = 400; // Elevation scale fixed height
  const canvasH = PAD_T + bodyH + PAD_B;

  const exportScale = lightMode ? 2 : 1;
  canvas.width = Math.ceil(canvasW * exportScale);
  canvas.height = Math.ceil(canvasH * exportScale);
  canvas.style.width = `${canvasW}px`;
  canvas.style.height = `${canvasH}px`;

  const ctx = canvas.getContext("2d");
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  if (exportScale !== 1) ctx.scale(exportScale, exportScale);
  ctx.clearRect(0, 0, canvasW, canvasH);

  // map chainage → canvas X
  function getX(ch) {
    return Number.isFinite(ch) ? PAD_L + (ch - minCh) * PX_PER_M_X : null;
  }
  // map elevation → canvas Y  (high RL = small Y)
  function getY(elev) {
    return PAD_T + bodyH - ((safeNum(elev) - minElev) / elevRange) * bodyH;
  }

  // ── Background ────────────────────────────────────────────────────────────
  const bg = ctx.createLinearGradient(0, 0, 0, canvasH);
  if (lightMode) {
    bg.addColorStop(0, "#f8fafc");
    bg.addColorStop(1, "#ffffff");
  } else {
    bg.addColorStop(0, "#0d1117");
    bg.addColorStop(1, "#141d2e");
  }
  ctx.fillStyle = bg;
  ctx.fillRect(0, 0, canvasW, canvasH);

  // ── Horizontal elevation grid ─────────────────────────────────────────────
  ctx.strokeStyle = lightMode ? "rgba(15,23,42,0.08)" : "rgba(255,255,255,0.06)";
  ctx.lineWidth = 1;
  const elevStep = elevRange > 200 ? 50 : elevRange > 50 ? 10 : elevRange > 20 ? 5 : 2;
  const startElev = Math.ceil(minElev / elevStep) * elevStep;
  for (let el = startElev; el <= maxElev; el += elevStep) {
    const gy = getY(el);
    ctx.beginPath(); ctx.moveTo(PAD_L, gy); ctx.lineTo(canvasW - PAD_R, gy); ctx.stroke();
    ctx.fillStyle = lightMode ? "rgba(15,23,42,0.6)" : "rgba(255,255,255,0.35)";
    ctx.font = `9px Outfit,sans-serif`;
    ctx.textAlign = "right";
    ctx.fillText(r3(el) + " m", PAD_L - 5, gy + 4);
  }

  // ── Vertical km grid ─────────────────────────────────────────────────────
  ctx.strokeStyle = lightMode ? "rgba(15,23,42,0.05)" : "rgba(255,255,255,0.05)";
  const gridKm = totalL >= 50000 ? 10000 : totalL >= 10000 ? 2000 : 1000;
  const startGrid = Math.ceil(minCh / gridKm) * gridKm;
  for (let ch = startGrid; ch <= maxCh; ch += gridKm) {
    const gx = getX(ch);
    if (gx == null) continue;
    ctx.beginPath(); ctx.moveTo(gx, PAD_T); ctx.lineTo(gx, PAD_T + bodyH); ctx.stroke();
  }

  // ── Build bridge intervals map for quick lookup ───────────────────────────
  const bridges = buildBridgeIntervals();
  // For each row, check if it's on a bridge
  function isOnBridge(ch) {
    return bridges.some(b => ch >= b.startChainage && ch <= b.endChainage);
  }
  function isTunnel(b) {
    return /tunnel/i.test(b.bridgeCategory || "") || /tunnel/i.test(b.bridgeType || "");
  }

  // ── Fill / Cut areas ─────────────────────────────────────────────────────
  // Draw fill (green) from formation down to ground
  for (let i = 1; i < rows.length; i++) {
    const r0 = rows[i - 1], r1 = rows[i];
    if (r1.bank <= 0.001) continue;
    const x0 = getX(r0.chainage), x1 = getX(r1.chainage);
    if (x0 == null || x1 == null) continue;
    ctx.beginPath();
    ctx.moveTo(x0, getY(r0.proposedLevel));
    ctx.lineTo(x1, getY(r1.proposedLevel));
    ctx.lineTo(x1, getY(r1.groundLevel));
    ctx.lineTo(x0, getY(r0.groundLevel));
    ctx.closePath();
    ctx.fillStyle = lightMode ? "rgba(34,197,94,0.2)" : "rgba(34,139,69,0.28)";
    ctx.fill();
  }
  // Draw cut (red) from ground down to cut-bottom
  for (let i = 1; i < rows.length; i++) {
    const r0 = rows[i - 1], r1 = rows[i];
    if (r1.cut <= 0.001) continue;
    const x0 = getX(r0.chainage), x1 = getX(r1.chainage);
    if (x0 == null || x1 == null) continue;
    ctx.beginPath();
    ctx.moveTo(x0, getY(r0.groundLevel));
    ctx.lineTo(x1, getY(r1.groundLevel));
    ctx.lineTo(x1, getY(r1.groundLevel - r1.cut));
    ctx.lineTo(x0, getY(r0.groundLevel - r0.cut));
    ctx.closePath();
    ctx.fillStyle = lightMode ? "rgba(239,68,68,0.18)" : "rgba(180,44,50,0.22)";
    ctx.fill();
  }

  // ── Ground level line ────────────────────────────────────────────────────
  ctx.beginPath();
  ctx.strokeStyle = lightMode ? "rgba(120,72,35,0.9)" : "rgba(180,120,60,0.85)";
  ctx.lineWidth = 2;
  let moved = false;
  for (const r of rows) {
    const x = getX(r.chainage), y = getY(r.groundLevel);
    if (x == null) { moved = false; continue; }
    if (!moved) { ctx.moveTo(x, y); moved = true; } else ctx.lineTo(x, y);
  }
  ctx.stroke();

  // ── Formation (proposed) RL line ─────────────────────────────────────────
  // Draw in segments — skip where bridge covers (will draw viaduct separately)
  ctx.lineWidth = 2.5;
  moved = false;
  let prevWasBridge = false;
  for (const r of rows) {
    const x = getX(r.chainage);
    const y = getY(r.proposedLevel);
    if (x == null) { moved = false; continue; }
    const onBr = isOnBridge(r.chainage);
    if (onBr) { moved = false; prevWasBridge = true; continue; }  // skip — drawn as viaduct
    ctx.strokeStyle = lightMode ? "rgba(37,99,235,0.9)" : "hsl(233,100%,65%)";
    if (!moved) { ctx.beginPath(); ctx.moveTo(x, y); moved = true; }
    else ctx.lineTo(x, y);
    if (prevWasBridge) { ctx.stroke(); ctx.beginPath(); ctx.moveTo(x, y); }
    prevWasBridge = false;
  }
  ctx.stroke();

  // ── Helper: interpolate formation/ground RL at any chainage ──────────────
  const flAt = (ch) => {
    for (let i = 0; i < rows.length - 1; i++) {
      if (ch >= rows[i].chainage && ch <= rows[i + 1].chainage) {
        const t = (ch - rows[i].chainage) / (rows[i + 1].chainage - rows[i].chainage);
        return rows[i].proposedLevel + t * (rows[i + 1].proposedLevel - rows[i].proposedLevel);
      }
    }
    return rows[rows.length - 1].proposedLevel;
  };
  const glAt = (ch) => {
    for (let i = 0; i < rows.length - 1; i++) {
      if (ch >= rows[i].chainage && ch <= rows[i + 1].chainage) {
        const t = (ch - rows[i].chainage) / (rows[i + 1].chainage - rows[i].chainage);
        return rows[i].groundLevel + t * (rows[i + 1].groundLevel - rows[i].groundLevel);
      }
    }
    return rows[rows.length - 1].groundLevel;
  };

  // ── Collect ALL annotation labels for de-overlap ──────────────────────────
  // Each label: { x, baseY, text, color, font, type }
  const allLabels = [];

  // ── Bridges / Viaducts / Tunnels ─────────────────────────────────────────
  for (const b of bridges) {
    const bx1 = getX(b.startChainage), bx2 = getX(b.endChainage);
    if (bx1 == null || bx2 == null) continue;

    const fl1 = flAt(b.startChainage), fl2 = flAt(b.endChainage);
    const gl1 = glAt(b.startChainage), gl2 = glAt(b.endChainage);
    const by1 = getY(fl1), by2 = getY(fl2);
    const gy1 = getY(gl1), gy2 = getY(gl2);
    const bW = Math.max(bx2 - bx1, 4);
    const tunnel = isTunnel(b);

    // All structures (bridges, tunnels, viaducts) drawn with the same style
    ctx.fillStyle = lightMode ? "rgba(59,130,246,0.2)" : "rgba(37,99,235,0.18)";
    ctx.strokeStyle = lightMode ? "rgba(37,99,235,0.9)" : "rgba(99,163,255,0.8)";
    ctx.lineWidth = 2.5;
    ctx.beginPath();
    ctx.moveTo(bx1, by1 - 6);
    ctx.lineTo(bx2, by2 - 6);
    ctx.lineTo(bx2, by2 + 6);
    ctx.lineTo(bx1, by1 + 6);
    ctx.closePath();
    ctx.fill(); ctx.stroke();

    ctx.strokeStyle = lightMode ? "rgba(37,99,235,0.85)" : "hsl(215,100%,65%)";
    ctx.lineWidth = 3;
    ctx.beginPath(); ctx.moveTo(bx1, by1); ctx.lineTo(bx2, by2); ctx.stroke();

    const pierSpacing = 60;
    const numPiers = Math.max(0, Math.floor(bW / pierSpacing) - 1);
    ctx.strokeStyle = lightMode ? "rgba(37,99,235,0.6)" : "rgba(99,163,255,0.6)"; ctx.lineWidth = 3;
    for (let p = 1; p <= numPiers; p++) {
      const t = p / (numPiers + 1);
      const px = bx1 + t * bW;
      const pyTop = by1 + t * (by2 - by1);
      const pierCh = b.startChainage + t * (b.endChainage - b.startChainage);
      const pierGl = glAt(pierCh);
      const pyBot = getY(pierGl);
      ctx.beginPath(); ctx.moveTo(px, pyTop + 6); ctx.lineTo(px, pyBot); ctx.stroke();
      ctx.strokeStyle = lightMode ? "rgba(37,99,235,0.4)" : "rgba(99,163,255,0.4)"; ctx.lineWidth = 8;
      ctx.beginPath(); ctx.moveTo(px, pyBot - 3); ctx.lineTo(px, pyBot); ctx.stroke();
      ctx.strokeStyle = lightMode ? "rgba(37,99,235,0.6)" : "rgba(99,163,255,0.6)"; ctx.lineWidth = 3;
    }

    ctx.strokeStyle = lightMode ? "rgba(37,99,235,0.55)" : "rgba(99,163,255,0.55)"; ctx.lineWidth = 4;
    ctx.beginPath(); ctx.moveTo(bx1, by1 - 10); ctx.lineTo(bx1, gy1); ctx.stroke();
    ctx.beginPath(); ctx.moveTo(bx2, by2 - 10); ctx.lineTo(bx2, gy2); ctx.stroke();

    // Collect label for de-overlap
    allLabels.push({
      x: (bx1 + bx2) / 2,
      baseY: Math.min(by1, by2) - 14,
      text: b.bridgeNo + (b.bridgeCategory ? ` (${b.bridgeCategory})` : ""),
      color: lightMode ? "#1d4ed8" : "#93c5fd",
      font: "bold 10px Outfit,sans-serif",
      type: tunnel ? "tunnel" : "bridge"
    });
  }

  // ── Station / Loop markers (vertical dashed) ──────────────────────────────
  if (state.loopPlatformRows) {
    for (const lp of state.loopPlatformRows) {
      const midCh = Number.isFinite(lp.loopStartCh) && Number.isFinite(lp.loopEndCh)
        ? (lp.loopStartCh + lp.loopEndCh) / 2
        : null;
      if (midCh == null) continue;
      const sx = getX(midCh);
      if (sx == null) continue;
      ctx.setLineDash([4, 4]);
      ctx.strokeStyle = lightMode ? "rgba(14,116,144,0.45)" : "rgba(34,211,238,0.5)"; ctx.lineWidth = 1.5;
      ctx.beginPath(); ctx.moveTo(sx, PAD_T); ctx.lineTo(sx, PAD_T + bodyH); ctx.stroke();
      ctx.setLineDash([]);

      // Collect station label for de-overlap
      allLabels.push({
        x: sx,
        baseY: PAD_T + 12,
        text: "◉ " + (lp.station || "Stn"),
        color: lightMode ? "#0e7490" : "#67e8f9",
        font: "bold 9px Outfit,sans-serif",
        type: "station"
      });
    }
  }

  // ── Curve markers (arc labels at top) ────────────────────────────────────
  if (state.curveRows) {
    state.curveRows.forEach((c) => {
      if (!Number.isFinite(c.chainage)) return;
      const cx = getX(c.chainage);
      if (cx == null) return;
      ctx.strokeStyle = lightMode ? "rgba(217,119,6,0.45)" : "rgba(252,211,77,0.5)"; ctx.lineWidth = 1; ctx.setLineDash([3, 4]);
      ctx.beginPath(); ctx.moveTo(cx, PAD_T + 20); ctx.lineTo(cx, PAD_T + bodyH); ctx.stroke();
      ctx.setLineDash([]);

      const rad = safeNum(c.radius);
      allLabels.push({
        x: cx,
        baseY: PAD_T + 8,
        text: (c.curve || "C") + (rad > 0 ? ` R=${r3(rad)}` : ""),
        color: lightMode ? "#b45309" : "#fde68a",
        font: "9px Outfit,sans-serif",
        type: "curve"
      });
    });
  }

  // ── De-overlap and draw all labels ────────────────────────────────────────
  // Sort by X position so we can compare nearest neighbors
  allLabels.sort((a, b) => a.x - b.x);

  // Measure widths for collision detection
  ctx.font = "bold 10px Outfit,sans-serif";
  for (const lb of allLabels) {
    ctx.font = lb.font;
    lb.width = ctx.measureText(lb.text).width;
    lb.finalY = lb.baseY;
  }

  // Push labels apart if they overlap (left-to-right sweep)
  const LABEL_H = 12; // approximate label height
  const MIN_GAP_X = 4; // minimum horizontal gap before treating as overlap
  for (let i = 1; i < allLabels.length; i++) {
    for (let j = i - 1; j >= Math.max(0, i - 8); j--) {
      const a = allLabels[j], b = allLabels[i];
      // Check horizontal overlap: do the label extents intersect?
      const aLeft = a.x - a.width / 2 - MIN_GAP_X;
      const aRight = a.x + a.width / 2 + MIN_GAP_X;
      const bLeft = b.x - b.width / 2;
      const bRight = b.x + b.width / 2;
      const hOverlap = aRight > bLeft && bRight > aLeft;
      if (!hOverlap) continue;

      // Check vertical overlap
      const vOverlap = Math.abs(b.finalY - a.finalY) < LABEL_H;
      if (vOverlap) {
        // Nudge the current label down (or up if already near bottom)
        b.finalY = a.finalY + LABEL_H;
        // Clamp to not exceed the chart body area
        if (b.finalY > PAD_T + bodyH - 10) {
          b.finalY = a.finalY - LABEL_H;
        }
      }
    }
  }

  // Now draw all labels at their de-overlapped positions
  for (const lb of allLabels) {
    ctx.fillStyle = lb.color;
    ctx.font = lb.font;
    ctx.textAlign = "center";

    // Draw a subtle connector line if label was pushed away from its base
    if (Math.abs(lb.finalY - lb.baseY) > 2) {
      ctx.strokeStyle = lb.color;
      ctx.globalAlpha = 0.25;
      ctx.lineWidth = 1;
      ctx.setLineDash([2, 2]);
      ctx.beginPath();
      ctx.moveTo(lb.x, lb.baseY);
      ctx.lineTo(lb.x, lb.finalY);
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.globalAlpha = 1.0;
    }

    ctx.fillText(lb.text, lb.x, lb.finalY);
  }

  // ── Y-axis border ────────────────────────────────────────────────────────
  ctx.strokeStyle = lightMode ? "rgba(15,23,42,0.25)" : "rgba(255,255,255,0.25)"; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(PAD_L, PAD_T); ctx.lineTo(PAD_L, PAD_T + bodyH); ctx.stroke();
  // Y-axis label (rotated)
  ctx.save();
  ctx.translate(14, PAD_T + bodyH / 2);
  ctx.rotate(-Math.PI / 2);
  ctx.fillStyle = lightMode ? "rgba(15,23,42,0.55)" : "rgba(255,255,255,0.4)"; ctx.font = "11px Outfit,sans-serif"; ctx.textAlign = "center";
  ctx.fillText("Elevation (m RL)", 0, 0);
  ctx.restore();

  // ── Scale ruler (X-axis) ─────────────────────────────────────────────────
  const rulerY = PAD_T + bodyH + 12;
  const tkInt = totalL >= 50000 ? 5000 : totalL >= 20000 ? 2000 : totalL >= 5000 ? 1000 : 500;
  const startTk = Math.ceil(minCh / tkInt) * tkInt;
  ctx.strokeStyle = lightMode ? "rgba(15,23,42,0.3)" : "rgba(255,255,255,0.3)"; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(PAD_L, rulerY); ctx.lineTo(canvasW - PAD_R, rulerY); ctx.stroke();
  ctx.fillStyle = lightMode ? "rgba(15,23,42,0.7)" : "rgba(255,255,255,0.55)"; ctx.font = `10px Outfit,sans-serif`;
  ctx.textAlign = "center";
  for (let ch = startTk; ch <= maxCh; ch += tkInt) {
    const tx = getX(ch);
    if (tx == null) continue;
    ctx.beginPath(); ctx.moveTo(tx, rulerY - 4); ctx.lineTo(tx, rulerY + 4); ctx.stroke();
    ctx.fillText(ch >= 1000 ? `${r3(ch / 1000)} km` : `${ch} m`, tx, rulerY + 16);
  }

  // ── Legend ────────────────────────────────────────────────────────────────
  const legend = [
    { col: lightMode ? "rgba(120,72,35,0.9)" : "rgba(180,120,60,0.85)", lbl: "Ground RL" },
    { col: lightMode ? "rgba(37,99,235,0.9)" : "hsl(233,100%,65%)", lbl: "Formation RL" },
    { col: lightMode ? "rgba(34,197,94,0.45)" : "rgba(34,139,69,0.5)", lbl: "Embankment (Fill)" },
    { col: lightMode ? "rgba(239,68,68,0.45)" : "rgba(180,44,50,0.5)", lbl: "Cut Section" },
    { col: lightMode ? "rgba(37,99,235,0.9)" : "rgba(99,163,255,0.8)", lbl: "Bridge / Viaduct / Tunnel" },
    { col: lightMode ? "rgba(14,116,144,0.7)" : "rgba(34,211,238,0.7)", lbl: "Station" },
  ];
  ctx.font = "10px Outfit,sans-serif"; let lx = PAD_L;
  for (const item of legend) {
    ctx.fillStyle = item.col; ctx.fillRect(lx, 8, 12, 12);
    ctx.fillStyle = lightMode ? "rgba(15,23,42,0.7)" : "rgba(255,255,255,0.7)"; ctx.textAlign = "left";
    ctx.fillText(item.lbl, lx + 15, 19);
    lx += 15 + ctx.measureText(item.lbl).width + 14;
    if (lx > canvasW - 80) break;
  }

  ensureRollCrosshair("rollCrosshair", "rollDiagramWrap");
  ensureRollCrosshair("sideRollCrosshair", "sideViewWrap");

  canvas.onmousemove = (e) => {
    const rect = canvas.getBoundingClientRect();
    const scaleX = canvas.width / rect.width;
    const visualX = e.clientX - rect.left;
    const mouseX = visualX * scaleX;
    const chHit = minCh + (mouseX - PAD_L) / PX_PER_M_X;

    if (chHit >= minCh && chHit <= maxCh && mouseX >= PAD_L && mouseX <= canvasW - PAD_R) {
      const { row: closest } = findClosestCalcRowByChainage(rows, chHit);
      if (!closest) return;
      const closestX = getX(closest.chainage);
      if (closestX == null) return;
      const planCanvas = els.rollDiagramCanvas;
      if (planCanvas) {
        const planVisualX = (closestX / planCanvas.width) * planCanvas.clientWidth;
        setRollCrosshairPosition("rollCrosshair", planCanvas, planVisualX);
      }
      setRollCrosshairPosition("sideRollCrosshair", canvas, visualX);
      canvas.style.cursor = "crosshair";
    } else {
      hideRollCrosshairs();
      canvas.style.cursor = "grab";
    }
  };

  canvas.onmouseleave = () => {
    hideRollCrosshairs();
    canvas.style.cursor = "grab";
  };
}

function syncRollFilterSelect() {
  if (!els.rollFilterSelect) return;
  els.rollFilterSelect.value = state.settings.rollFilter || "all";
}

function buildSettingsInputs() {
  if (els.profileGrid) {
    const profile = state.project?.profile || {};
    els.profileGrid.innerHTML = `
      <label>
        <span class="lbl-text">Corridor Name</span>
        <input type="text" name="corridorName" value="${escapeAttr(profile.corridorName || "")}" />
      </label>
      <label>
        <span class="lbl-text">Direction</span>
        <select name="direction">
          <option value="Up" ${profile.direction === "Up" ? "selected" : ""}>Up</option>
          <option value="Down" ${profile.direction === "Down" ? "selected" : ""}>Down</option>
        </select>
      </label>
      <label>
        <span class="lbl-text">Chainage Zero Reference</span>
        <input type="text" name="chainageZeroRef" value="${escapeAttr(profile.chainageZeroRef || "")}" placeholder="e.g. Nashik Yard 0+000" />
      </label>
    `;
  }

  if (els.standardsGrid) {
    const gauge = state.settings.gauge || "BG";
    const formationPreset = state.settings.formationPreset || "single";
    const slopePreset = state.settings.slopePreset || "default";
    els.standardsGrid.innerHTML = `
      <label>
        <span class="lbl-text">Gauge</span>
        <select name="gauge">
          <option value="BG" ${gauge === "BG" ? "selected" : ""}>Broad Gauge (1676 mm)</option>
          <option value="SG" ${gauge === "SG" ? "selected" : ""}>Standard Gauge (1435 mm)</option>
          <option value="MG" ${gauge === "MG" ? "selected" : ""}>Metre Gauge (1000 mm)</option>
        </select>
      </label>
      <label>
        <span class="lbl-text">Formation Preset</span>
        <select name="formationPreset">
          <option value="single" ${formationPreset === "single" ? "selected" : ""}>Single Line</option>
          <option value="double" ${formationPreset === "double" ? "selected" : ""}>Double Line</option>
          <option value="custom" ${formationPreset === "custom" ? "selected" : ""}>Custom</option>
        </select>
      </label>
      <label>
        <span class="lbl-text">Slope Preset (H:V)</span>
        <select name="slopePreset">
          <option value="default" ${slopePreset === "default" ? "selected" : ""}>Default (1:2.0)</option>
          <option value="1.5" ${slopePreset === "1.5" ? "selected" : ""}>1:1.5</option>
          <option value="2" ${slopePreset === "2" ? "selected" : ""}>1:2.0</option>
          <option value="2.5" ${slopePreset === "2.5" ? "selected" : ""}>1:2.5</option>
          <option value="3" ${slopePreset === "3" ? "selected" : ""}>1:3.0</option>
        </select>
      </label>
      <label>
        <span class="lbl-text">Berm 1 at Height (m)</span>
        <input type="number" step="0.1" name="bermRulePrimary" value="${r3(state.settings.bermRulePrimary)}" />
      </label>
      <label>
        <span class="lbl-text">Berm 2 at Height (m)</span>
        <input type="number" step="0.1" name="bermRuleSecondary" value="${r3(state.settings.bermRuleSecondary)}" />
      </label>
    `;
  }

  if (els.qualityGrid) {
    els.qualityGrid.innerHTML = `
      <label>
        <span class="lbl-text">Minimum Track Centre (m)</span>
        <input type="number" step="0.1" name="minTc" value="${r3(state.settings.minTc)}" />
      </label>
      <label>
        <span class="lbl-text">Minimum Platform Width (m)</span>
        <input type="number" step="0.1" name="minPlatformWidth" value="${r3(state.settings.minPlatformWidth)}" />
      </label>
      <label>
        <span class="lbl-text">Minimum Clearance (m)</span>
        <input type="number" step="0.1" name="minClearance" value="${r3(state.settings.minClearance)}" />
      </label>
    `;
  }

  if (els.settingsGrid) {
    els.settingsGrid.innerHTML = settingSchema.map(([key, label]) => {
      if (key === "activeSqCategory") {
        const current = Math.min(3, Math.max(1, Math.round(safeNum(state.settings[key], 3))));
        return `
          <label>
            <span class="lbl-text">${label}</span>
            <select name="${key}" required>
              <option value="1" ${current === 1 ? "selected" : ""}>SQ1</option>
              <option value="2" ${current === 2 ? "selected" : ""}>SQ2</option>
              <option value="3" ${current === 3 ? "selected" : ""}>SQ3</option>
            </select>
          </label>
        `;
      }
      return `
        <label>
          <span class="lbl-text">${label}</span>
          <input type="number" step="0.001" name="${key}" value="${r3(state.settings[key])}" required />
        </label>
      `;
    }).join("");
  }

  if (els.materialProfileGrid) {
    const rows = Array.isArray(state.settings.materialProfile) ? state.settings.materialProfile : [];
    els.materialProfileGrid.innerHTML = `
      <table class="bridge-table">
        <thead>
          <tr>
            <th>Layer Depth (m)</th>
            <th>Density (t/m³)</th>
            <th>Shrink (%)</th>
            <th>Swell (%)</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map((r, i) => `
            <tr>
              <td><input data-mp-field="depth" data-mp-index="${i}" value="${r3(r.depth)}" /></td>
              <td><input data-mp-field="density" data-mp-index="${i}" value="${r3(r.density)}" /></td>
              <td><input data-mp-field="shrink" data-mp-index="${i}" value="${r3(r.shrink)}" /></td>
              <td><input data-mp-field="swell" data-mp-index="${i}" value="${r3(r.swell)}" /></td>
              <td><button type="button" class="bridge-del" data-mp-del="${i}">Delete</button></td>
            </tr>
          `).join("")}
        </tbody>
      </table>
      <button type="button" class="btn btn-secondary" id="addMaterialRowBtn" style="margin-top:10px;">Add Layer</button>
    `;
  }

  const openStationLayout = () => {
    if (!els.stationLayoutModal) return;
    const stations = getStationChoices();
    if (els.stationLayoutSelect) {
      els.stationLayoutSelect.innerHTML = stations.map((s) => `<option value="${escapeAttr(s)}">${escapeHtml(s)}</option>`).join("");
    }
    if (stations.length) {
      buildStationLayoutList(stations[0]);
    } else {
      if (els.stationLayoutList) els.stationLayoutList.innerHTML = '<p class="muted">No stations found. Import loops first.</p>';
    }
    els.stationLayoutModal.showModal();
  };
  if (els.openStationLayoutBtn) {
    els.openStationLayoutBtn.addEventListener("click", openStationLayout);
  }
  if (els.openStationLayoutInlineBtn) {
    els.openStationLayoutInlineBtn.addEventListener("click", openStationLayout);
  }
  if (els.openOverridesBtn && els.overridesModal) {
    els.openOverridesBtn.addEventListener("click", () => {
      renderOverrideTable();
      els.overridesModal.showModal();
    });
  }
  if (els.closeOverridesBtn) {
    els.closeOverridesBtn.addEventListener("click", () => els.overridesModal?.close());
  }
  if (els.addOverrideBtn) {
    els.addOverrideBtn.addEventListener("click", () => {
      state.calcOverrides = [...(state.calcOverrides || []), { startCh: 0, endCh: 0, type: "", bank: NaN, cut: NaN, slope: NaN }];
      renderOverrideTable();
    });
  }
  if (els.overrideTableBody) {
    els.overrideTableBody.addEventListener("click", (e) => {
      const del = e.target.closest("[data-ov-del]");
      if (del) {
        const idx = Number(del.dataset.ovDel);
        if (Number.isFinite(idx)) {
          state.calcOverrides = (state.calcOverrides || []).filter((_, i) => i !== idx);
          renderOverrideTable();
        }
      }
    });
  }
  if (els.applyOverridesBtn) {
    els.applyOverridesBtn.addEventListener("click", () => {
      syncOverridesFromTable();
      recalculate();
      els.overridesModal?.close();
    });
  }
  if (els.openSlopeZonesBtn && els.slopeZonesModal) {
    els.openSlopeZonesBtn.addEventListener("click", () => {
      renderSlopeZoneTable();
      els.slopeZonesModal.showModal();
    });
  }
  if (els.openBermRulesBtn) {
    els.openBermRulesBtn.addEventListener("click", () => {
      els.standardsGrid?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  }
  if (els.openQualityBtn && els.qualityModal) {
    els.openQualityBtn.addEventListener("click", () => {
      renderQualityChecks();
      els.qualityModal.showModal();
    });
  }
  if (els.closeQualityBtn) {
    els.closeQualityBtn.addEventListener("click", () => els.qualityModal?.close());
  }
  if (els.downloadBoqBtn) {
    els.downloadBoqBtn.addEventListener("click", () => {
      exportBoqCsv();
    });
  }
  if (els.exportExcelBtn) {
    els.exportExcelBtn.addEventListener("click", () => {
      exportCalculationExcel();
    });
  }
  if (els.rollFilterSelect) {
    els.rollFilterSelect.addEventListener("change", () => {
      state.settings.rollFilter = String(els.rollFilterSelect.value || "all");
      renderRollDiagram();
      renderSideView();
    });
  }
  if (els.closeSlopeZonesBtn) {
    els.closeSlopeZonesBtn.addEventListener("click", () => els.slopeZonesModal?.close());
  }
  if (els.addSlopeZoneBtn) {
    els.addSlopeZoneBtn.addEventListener("click", () => {
      state.slopeZones = [...(state.slopeZones || []), { startCh: 0, endCh: 0, slope: state.settings.sideSlopeFactor }];
      renderSlopeZoneTable();
    });
  }
  if (els.slopeZoneTableBody) {
    els.slopeZoneTableBody.addEventListener("click", (e) => {
      const del = e.target.closest("[data-sz-del]");
      if (del) {
        const idx = Number(del.dataset.szDel);
        if (Number.isFinite(idx)) {
          state.slopeZones = (state.slopeZones || []).filter((_, i) => i !== idx);
          renderSlopeZoneTable();
        }
      }
    });
  }
  if (els.applySlopeZonesBtn) {
    els.applySlopeZonesBtn.addEventListener("click", () => {
      syncSlopeZonesFromTable();
      recalculate();
      els.slopeZonesModal?.close();
    });
  }
  if (els.closeStationLayoutBtn) {
    els.closeStationLayoutBtn.addEventListener("click", () => els.stationLayoutModal?.close());
  }
  if (els.stationLayoutSelect) {
    els.stationLayoutSelect.addEventListener("change", () => {
      buildStationLayoutList(els.stationLayoutSelect.value);
    });
  }
  if (els.stationLayoutList) {
    els.stationLayoutList.addEventListener("click", (e) => {
      const item = e.target.closest(".layout-item");
      if (!item) return;
      if (e.target?.hasAttribute("data-layout-up")) {
        const prev = item.previousElementSibling;
        if (prev) item.parentElement.insertBefore(item, prev);
      }
      if (e.target?.hasAttribute("data-layout-down")) {
        const next = item.nextElementSibling;
        if (next) item.parentElement.insertBefore(next, item);
      }
    });
    els.stationLayoutList.addEventListener("dragstart", (e) => {
      const item = e.target.closest(".layout-item");
      if (!item) return;
      item.classList.add("dragging");
      e.dataTransfer.setData("text/plain", item.dataset.rowIndex || "");
    });
    els.stationLayoutList.addEventListener("dragend", (e) => {
      const item = e.target.closest(".layout-item");
      if (item) item.classList.remove("dragging");
    });
    els.stationLayoutList.addEventListener("dragover", (e) => {
      e.preventDefault();
      const dragging = els.stationLayoutList.querySelector(".layout-item.dragging");
      const target = e.target.closest(".layout-item");
      if (!dragging || !target || dragging === target) return;
      const rect = target.getBoundingClientRect();
      const after = (e.clientY - rect.top) > rect.height / 2;
      els.stationLayoutList.insertBefore(dragging, after ? target.nextSibling : target);
    });
  }
  if (els.saveStationLayoutBtn) {
    els.saveStationLayoutBtn.addEventListener("click", () => {
      if (!els.stationLayoutList || !els.stationLayoutSelect) return;
      const items = Array.from(els.stationLayoutList.querySelectorAll(".layout-item"));
      items.forEach((el, idx) => {
        const rowIndex = Number(el.dataset.rowIndex);
        if (Number.isFinite(rowIndex) && state.loopPlatformRows[rowIndex]) {
          state.loopPlatformRows[rowIndex].order = idx;
          const sideSel = el.querySelector("[data-layout-side]");
          const tcInput = el.querySelector("[data-layout-tc]");
          const pfInput = el.querySelector("[data-layout-pf]");
          if (sideSel) state.loopPlatformRows[rowIndex].side = String(sideSel.value || "").trim();
          if (tcInput && !tcInput.disabled) {
            state.loopPlatformRows[rowIndex].tc = safeNum(tcInput.value, state.loopPlatformRows[rowIndex].tc);
          }
          if (pfInput && !pfInput.disabled) {
            state.loopPlatformRows[rowIndex].pfWidth = safeNum(pfInput.value, state.loopPlatformRows[rowIndex].pfWidth);
          }
        }
      });
      renderLoopInputs();
      recalculate();
      els.stationLayoutModal?.close();
    });
  }

  if (els.visualGrid) {
    els.visualGrid.innerHTML = `
      <label class="toggle-row"><span>Show Rails</span><input type="checkbox" class="toggle-input" name="showRails" ${state.settings.showRails ? "checked" : ""} /></label>
      <label class="toggle-row"><span>Show Platforms</span><input type="checkbox" class="toggle-input" name="showPlatforms" ${state.settings.showPlatforms ? "checked" : ""} /></label>
      <label class="toggle-row"><span>Show Drains</span><input type="checkbox" class="toggle-input" name="showDrains" ${state.settings.showDrains ? "checked" : ""} /></label>
      <label class="toggle-row"><span>Show Labels</span><input type="checkbox" class="toggle-input" name="showLabels" ${state.settings.showLabels ? "checked" : ""} /></label>
      <label>Roll Diagram Filter
        <select name="rollFilter">
          <option value="all" ${state.settings.rollFilter === "all" ? "selected" : ""}>All Lines</option>
          <option value="main" ${state.settings.rollFilter === "main" ? "selected" : ""}>Main Line Only</option>
          <option value="stations" ${state.settings.rollFilter === "stations" ? "selected" : ""}>Stations Only</option>
        </select>
      </label>
    `;
  }

  if (els.mapGrid) {
    els.mapGrid.innerHTML = `
      <label class="toggle-row"><span>Station Pins</span><input type="checkbox" class="toggle-input" name="showMapStations" ${state.settings.showMapStations ? "checked" : ""} /></label>
      <label class="toggle-row"><span>Bridge Pins</span><input type="checkbox" class="toggle-input" name="showMapBridges" ${state.settings.showMapBridges ? "checked" : ""} /></label>
      <label class="toggle-row"><span>Curve Pins</span><input type="checkbox" class="toggle-input" name="showMapCurves" ${state.settings.showMapCurves ? "checked" : ""} /></label>
    `;
  }

  if (els.boqGrid) {
    const mapping = state.settings.boqMapping || {};
    els.boqGrid.innerHTML = `
      <label>Range Column
        <input type="text" name="boqRange" value="${escapeAttr(mapping.range || "")}" />
      </label>
      <label>Prepared Column
        <input type="text" name="boqPrepared" value="${escapeAttr(mapping.prepared || "")}" />
      </label>
      <label>Blanket Column
        <input type="text" name="boqBlanket" value="${escapeAttr(mapping.blanket || "")}" />
      </label>
      <label>Fill Column
        <input type="text" name="boqFill" value="${escapeAttr(mapping.fill || "")}" />
      </label>
      <label>Cut Column
        <input type="text" name="boqCut" value="${escapeAttr(mapping.cut || "")}" />
      </label>
    `;
  }
  syncRollFilterSelect();
}

function applySettingsFromForm() {
  const fd = new FormData(els.settingsForm);
  state.project.profile = {
    corridorName: String(fd.get("corridorName") || "").trim(),
    direction: String(fd.get("direction") || "Up"),
    chainageZeroRef: String(fd.get("chainageZeroRef") || "").trim(),
  };
  const gauge = String(fd.get("gauge") || "BG");
  const formationPreset = String(fd.get("formationPreset") || "single");
  const slopePreset = String(fd.get("slopePreset") || "default");
  state.settings.gauge = gauge;
  state.settings.formationPreset = formationPreset;
  state.settings.slopePreset = slopePreset;
  state.settings.bermRulePrimary = safeNum(fd.get("bermRulePrimary"), state.settings.bermRulePrimary);
  state.settings.bermRuleSecondary = safeNum(fd.get("bermRuleSecondary"), state.settings.bermRuleSecondary);
  state.settings.minTc = safeNum(fd.get("minTc"), state.settings.minTc);
  state.settings.minPlatformWidth = safeNum(fd.get("minPlatformWidth"), state.settings.minPlatformWidth);
  state.settings.minClearance = safeNum(fd.get("minClearance"), state.settings.minClearance);
  for (const [k] of settingSchema) {
    state.settings[k] = safeNum(fd.get(k), state.settings[k]);
  }
  state.settings.showRails = fd.get("showRails") === "on";
  state.settings.showPlatforms = fd.get("showPlatforms") === "on";
  state.settings.showDrains = fd.get("showDrains") === "on";
  state.settings.showLabels = fd.get("showLabels") === "on";
  state.settings.rollFilter = String(fd.get("rollFilter") || "all");
  state.settings.showMapStations = fd.get("showMapStations") === "on";
  state.settings.showMapBridges = fd.get("showMapBridges") === "on";
  state.settings.showMapCurves = fd.get("showMapCurves") === "on";
  state.settings.boqMapping = {
    range: String(fd.get("boqRange") || state.settings.boqMapping?.range || "").trim(),
    prepared: String(fd.get("boqPrepared") || state.settings.boqMapping?.prepared || "").trim(),
    blanket: String(fd.get("boqBlanket") || state.settings.boqMapping?.blanket || "").trim(),
    fill: String(fd.get("boqFill") || state.settings.boqMapping?.fill || "").trim(),
    cut: String(fd.get("boqCut") || state.settings.boqMapping?.cut || "").trim(),
  };
  syncRollFilterSelect();

  if (formationPreset !== "custom") {
    if (formationPreset === "single") {
      state.settings.formationWidthFill = 7.85;
      state.settings.cuttingWidth = 12.05;
    } else if (formationPreset === "double") {
      state.settings.formationWidthFill = 12.05;
      state.settings.cuttingWidth = 16.5;
    }
  }
  if (slopePreset !== "default") {
    const slopeVal = parseLooseNumber(slopePreset, NaN);
    if (Number.isFinite(slopeVal)) state.settings.sideSlopeFactor = slopeVal;
  }

  syncMaterialProfileFromGrid();
  recalculate();
  if (typeof drawAlignmentMap === "function" && state.kmlData) {
    drawAlignmentMap();
  }
  renderRollDiagram();
  renderSideView();
}

function updateCrossSectionNav() {
  if (els.crossPrevBtn) {
    els.crossPrevBtn.disabled = !Number.isFinite(state.currentCrossIndex) || state.currentCrossIndex <= 0;
  }
  if (els.crossNextBtn) {
    els.crossNextBtn.disabled = !Number.isFinite(state.currentCrossIndex) || state.currentCrossIndex >= state.calcRows.length - 1;
  }
}

function ensureCrossSectionModalOpen() {
  if (!els.crossSectionModal) return;
  if (els.crossSectionModal.open) return;
  try {
    els.crossSectionModal.showModal();
  } catch (error) {
    els.crossSectionModal.setAttribute("open", "");
  }
}

function openCrossSectionByIndex(index) {
  if (!Number.isFinite(index)) return;
  const boundedIndex = Math.max(0, Math.min(Math.trunc(index), state.calcRows.length - 1));
  const row = state.calcRows[boundedIndex];
  if (!row) return;
  state.currentCrossIndex = boundedIndex;
  try {
    drawCrossSection(row);
  } catch (error) {
    console.error("Cross-section render failed:", error);
    if (els.crossSvg) {
      els.crossSvg.innerHTML = `
        <rect x="-120000" y="-120000" width="240000" height="240000" fill="#f8fcff" />
        <text x="${CROSS_SVG_W / 2}" y="${CROSS_SVG_H / 2}" text-anchor="middle" fill="#1f2e3a" font-size="16" font-weight="700">Cross-section render failed</text>
      `;
    }
    ensureCrossSectionModalOpen();
  }
}

function drawCrossSection(row, targetEl = els.crossSvg) {
  if (!targetEl) return;
  if (targetEl === els.crossSvg) {
    const rowIndex = state.calcRows.indexOf(row);
    if (rowIndex >= 0) {
      state.currentCrossIndex = rowIndex;
    }
    updateCrossSectionNav();
  }
  if (row.type === "BRIDGE") {
    const bridgeName = (row.bridgeRefs && row.bridgeRefs.length)
      ? row.bridgeRefs.join(", ")
      : "Bridge";
    const svgW = CROSS_SVG_W;
    const svgH = CROSS_SVG_H;
    targetEl.innerHTML = `
      <rect x="-120000" y="-120000" width="240000" height="240000" fill="#f8fcff" />
      <text x="${svgW / 2}" y="${svgH / 2 - 10}" text-anchor="middle" fill="#1f2e3a" font-size="22" font-weight="700">${escapeHtml(bridgeName)}</text>
      <text x="${svgW / 2}" y="${svgH / 2 + 18}" text-anchor="middle" fill="#4b5c6a" font-size="13">Cross-section deducted for bridge</text>
      <text x="42" y="40" fill="#1f2e3a" font-size="16" font-weight="700">Cross Section of Track</text>
    `;
    if (targetEl === els.crossSvg) {
      resetCrossView();
      try {
        if (!els.crossSectionModal.open) {
          els.crossSectionModal.showModal();
        }
      } catch (error) {
        console.error("Cross-section modal open failed:", error);
      }
    }
    return;
  }
  const s = state.settings;
  const showRails = s.showRails !== false;
  const showPlatforms = s.showPlatforms !== false;
  const showDrains = s.showDrains !== false;
  const showLabels = s.showLabels !== false;
  const fl = row.proposedLevel;
  const gl = row.groundLevel;
  const topWidthM = Math.max(safeNum(row.topWidth), 0) || safeNum(s.formationWidthFill);
  const standardTrackCenterM = 5.3;
  const sequenceLayout = buildStationSequenceLayout(
    row.chainage,
    row.stationName || row.station || "",
    standardTrackCenterM,
    { useRanges: false },
  );
  const activeLoopTrackOffsets = sequenceLayout ? [] : (state.loopPlatformRows || [])
    .filter((lp) =>
      Number.isFinite(lp.loopStartCh) &&
      Number.isFinite(lp.loopEndCh) &&
      row.chainage >= lp.loopStartCh &&
      row.chainage <= lp.loopEndCh &&
      Math.abs(safeNum(lp.tc, 0)) > 0,
    )
    .map((lp) => {
      const offset = safeNum(lp.tc, 0);
      if (!offset) return NaN;
      const side = normalizeLoopSide(lp.side);
      return side === "Right" ? -offset : offset;
    })
    .filter((offset) => Number.isFinite(offset))
    .sort((a, b) => a - b)
    .filter((offset, index, arr) => index === 0 || Math.abs(offset - arr[index - 1]) > 0.001);
  const rawTrackOffsets = sequenceLayout
    ? sequenceLayout.trackItems.map((item) => sequenceLayout.offsetByItem.get(item)).filter((v) => Number.isFinite(v))
    : [0, ...activeLoopTrackOffsets];
  const allTrackCenterOffsets = rawTrackOffsets
    .filter((v) => Number.isFinite(v))
    .sort((a, b) => a - b)
    .filter((offset, index, arr) => index === 0 || Math.abs(offset - arr[index - 1]) > 0.001);
  const trackOffsetMidpoint = allTrackCenterOffsets.length > 1
    ? (Math.min(...allTrackCenterOffsets) + Math.max(...allTrackCenterOffsets)) / 2
    : 0;
  const fmtDim = (value) => `${Number(safeNum(value)).toFixed(1)} m`;
  const ballastTopWidthM = 3.45;
  const ballastBottomWidthM = 6.1;
  const ballastHeightM = 0.35;
  const sqCategory = Math.min(3, Math.max(1, Math.round(s.activeSqCategory || 3)));
  const sqName = sqCategory === 1 ? "SQ1" : (sqCategory === 2 ? "SQ2" : "SQ3");
  const blanketBySq = { 1: 1.0, 2: 0.75, 3: 0.6 };
  const blanketRuleThickness = blanketBySq[sqCategory];
  const ballastThickness = ballastHeightM;
  const topLayerThickness = Math.max(0, s.topLayerThickness || 0.5);
  const topLayerEffectiveThickness = Math.min(topLayerThickness, Math.max(row.bank, 0));
  const embankmentCoreThickness = Math.max(row.bank - topLayerEffectiveThickness, 0);
  const layers = [
    { name: "Ballast Cushion", t: ballastThickness, top: fl + ballastThickness, bottom: fl, color: "#d5d8de" },
    { name: `Blanket (${sqName})`, t: blanketRuleThickness, top: fl, bottom: fl - blanketRuleThickness, color: "#f5eecf" },
    { name: "Top Layer Embankment Fill", t: topLayerEffectiveThickness, top: fl - blanketRuleThickness, bottom: fl - blanketRuleThickness - topLayerEffectiveThickness, color: "#edd6bf" },
    { name: "Embankment Fill (Core)", t: embankmentCoreThickness, top: fl - blanketRuleThickness - topLayerEffectiveThickness, bottom: gl, color: "#dfe9de" },
  ];

  const layerRows = layers;

  const drawTopWidthM = (() => {
    if (!sequenceLayout || !allTrackCenterOffsets.length) return topWidthM;
    const positives = allTrackCenterOffsets.filter((o) => o > 0);
    const negatives = allTrackCenterOffsets.filter((o) => o < 0).map((o) => Math.abs(o));
    const maxLeft = positives.length ? Math.max(...positives) : 0;
    const maxRight = negatives.length ? Math.max(...negatives) : 0;
    const edgeClearM = ballastBottomWidthM / 2;
    let leftExtra = 0;
    let rightExtra = 0;
    const ordered = sequenceLayout.orderedItems || [];
    const findNeighborTrack = (startIndex, step) => {
      for (let i = startIndex + step; i >= 0 && i < ordered.length; i += step) {
        if (ordered[i].kind === "track") return ordered[i];
      }
      return null;
    };
    ordered.forEach((item, idx) => {
      if (item.kind !== "platform") return;
      const width = safeNum(item.row.pfWidth, 0);
      if (!width) return;
      const before = findNeighborTrack(idx, -1);
      const after = findNeighborTrack(idx, 1);
      if (before && after) return; // island platform, no extra beyond tracks
      const anchor = before || after;
      if (!anchor) return;
      const anchorOffset = sequenceLayout.offsetByItem.get(anchor);
      const edgeExtra = (ballastTopWidthM / 2) + 0.3;
      if (Number.isFinite(anchorOffset)) {
        if (anchorOffset >= 0) leftExtra = Math.max(leftExtra, width + edgeExtra);
        else rightExtra = Math.max(rightExtra, width + edgeExtra);
      }
    });
    const requiredHalf = Math.max(maxLeft + edgeClearM + leftExtra, maxRight + edgeClearM + rightExtra, edgeClearM);
    return Math.max(topWidthM, requiredHalf * 2);
  })();

  if (targetEl === els.crossSvg) {
    try {
      els.layerTbody.innerHTML = layerRows.map((l) => `
        <tr>
          <td>${l.name}</td>
          <td>${r3(l.t)}</td>
          <td>${r3(l.top)}</td>
          <td>${r3(l.bottom)}</td>
        </tr>
      `).join("") + `
        <tr>
          <td><strong>Formation RL</strong></td>
          <td>-</td>
          <td>${r3(fl)}</td>
          <td>${r3(fl)}</td>
        </tr>
        <tr>
          <td><strong>Ground RL (GL)</strong></td>
          <td>-</td>
          <td>${r3(gl)}</td>
          <td>${r3(gl)}</td>
        </tr>
      `;
      els.dimTbody.innerHTML = `
        <tr><th>Formation Width (Top)</th><td>${r3(topWidthM)} m</td></tr>
        <tr><th>Tracks in Section</th><td>${allTrackCenterOffsets.length}</td></tr>
        <tr><th>Track Centres</th><td>${allTrackCenterOffsets.map((offset) => offset === 0 ? "Main Line" : `${r3(offset)} m Loop`).join(", ")}</td></tr>
        <tr><th>Ballast Top Width (Each Track)</th><td>3.45 m</td></tr>
        <tr><th>Ballast Bottom Width (Each Track)</th><td>6.10 m</td></tr>
        <tr><th>Ballast Height</th><td>0.35 m</td></tr>
        <tr><th>Berm Width (Each Berm)</th><td>${fmtDim(s.bermWidth)}</td></tr>
        <tr><th>Berms per Side (Drawing)</th><td>${row.bank >= safeNum(s.bermRuleSecondary, 8) ? 2 : (row.bank >= safeNum(s.bermRulePrimary, 4) ? 1 : 0)}</td></tr>
        <tr><th>Bottom Width (Fill)</th><td>${r3(row.fillBottom)} m</td></tr>
        <tr><th>Bottom Width (Cut)</th><td>${r3(row.cutBottom)} m</td></tr>
        <tr><th>Embankment Height</th><td>${r3(row.bank)} m</td></tr>
        <tr><th>Cut Depth</th><td>${r3(row.cut)} m</td></tr>
        <tr><th>Side Slope Factor</th><td>${r3(s.sideSlopeFactor)}</td></tr>
        <tr><th>Blanket Rule</th><td>SQ3=0.600 m, SQ2=0.750 m, SQ1=1.000 m</td></tr>
      `;

      els.crossTitle.textContent = `Cross-Section @ CH ${r3(row.chainage)} m`;
      els.crossMeta.textContent = `Type: ${row.type} | Ground RL: ${r3(gl)} | Proposed RL: ${r3(fl)} | Bank: ${r3(row.bank)} | Cut: ${r3(row.cut)}`;
    } catch (error) {
      console.error("Cross-section side panel update failed:", error);
    }
  }

  const svgW = CROSS_SVG_W;
  const svgH = CROSS_SVG_H;
  const marginX = 280;
  const centerX = svgW / 2;
  const layerTotalM = ballastThickness + blanketRuleThickness + topLayerThickness;
  const halfTopM = drawTopWidthM / 2;
  const slopeHV = safeNum(row.slopeFactor, s.sideSlopeFactor || 2);
  const fillRunM = row.bank * slopeHV;
  const cutRunM = row.cut * slopeHV;
  const bermPrimary = safeNum(s.bermRulePrimary, 4);
  const bermSecondary = safeNum(s.bermRuleSecondary, 8);
  const fillBermCount = row.bank >= bermSecondary ? 2 : (row.bank >= bermPrimary ? 1 : 0);
  const fillBermExtraM = row.bank > 0 ? (fillBermCount * safeNum(s.bermWidth, 0)) : 0;
  const halfBottomM = Math.max(halfTopM + fillRunM + fillBermExtraM, (row.fillBottom / 2) + fillBermExtraM, halfTopM);
  const halfCutM = Math.max(halfTopM + cutRunM, row.cutBottom / 2, s.cuttingWidth / 2);
  const maxHalfM = Math.max(halfBottomM + s.bermWidth + 1, halfCutM + 1.5, halfTopM + 1);
  const pxByWidth = ((svgW - (2 * marginX)) / 2) / Math.max(maxHalfM, 1);
  const maxVerticalM = Math.max(row.bank, row.cut, 0) + layerTotalM + 4;
  const pxByHeight = (svgH - 220) / Math.max(maxVerticalM, 1);
  const pxPerM = Math.max(18, Math.min(42, pxByWidth, pxByHeight));
  const halfTop = halfTopM * pxPerM;
  const halfBottom = halfBottomM * pxPerM;
  const halfCutBottom = halfCutM * pxPerM;
  const bermW = s.bermWidth * pxPerM;
  const fillH = row.bank * pxPerM;
  const cutH = row.cut * pxPerM;
  const datumY = svgH * 0.58;
  const topY = row.type === "CUTTING" ? (datumY - (Math.max(cutH, 120) * 0.45)) : (datumY - fillH);
  const toeY = topY + fillH;
  const cutBottomY = topY - cutH;
  const leftToeX = centerX - halfBottom;
  const rightToeX = centerX + halfBottom;
  const leftCutX = centerX - halfCutBottom;
  const rightCutX = centerX + halfCutBottom;
  let toeLeftX = row.type === "CUTTING" ? leftCutX : leftToeX;
  let toeRightX = row.type === "CUTTING" ? rightCutX : rightToeX;
  let toeBaseY = row.type === "CUTTING" ? cutBottomY : toeY;
  const groundY = row.type === "CUTTING" ? cutBottomY : toeY;
  const hflY = groundY - 48;
  const centerRef = row.type === "CUTTING" ? cutBottomY : toeY;

  state.crossMatrix = {
    centerX,
    topY,
    pxPerM,
    fl,
    gl,
  };

  const drawDim = (x1, x2, y, label) => `
    <line x1="${x1}" y1="${y}" x2="${x2}" y2="${y}" stroke="#2f4d6a" stroke-width="1.6" marker-start="url(#arrowSmall)" marker-end="url(#arrowSmall)" />
    <text x="${(x1 + x2) / 2}" y="${y - 6}" text-anchor="middle" fill="#2f4d6a" font-size="12" font-weight="700">${label}</text>
  `;
  const drawDimLine = (x1, x2, y) => `
    <line x1="${x1}" y1="${y}" x2="${x2}" y2="${y}" stroke="#2f4d6a" stroke-width="1.4" marker-start="url(#arrowSmall)" marker-end="url(#arrowSmall)" />
  `;

  const ballastH = ballastThickness * pxPerM;
  const blanketH = blanketRuleThickness * pxPerM;
  const topLayerH = topLayerEffectiveThickness * pxPerM;
  const sleeperThicknessPx = Math.max((s.sleeperThickness || 0.25) * pxPerM, 8);
  const railHeightPx = Math.max((s.railHeight || 0.18) * pxPerM, 5);
  const blanketHDraw = row.bank > 0 ? Math.min(blanketH, fillH) : blanketH;
  const topLayerHDraw = row.bank > 0 ? Math.min(topLayerH, Math.max(fillH - blanketHDraw, 0)) : topLayerH;
  const trackTopY = topY - ballastH;
  const layerCalloutStyles = {
    ballast: { stroke: "#7f8a95", text: "#6f7a86" },
    blanket: { stroke: "#90886c", text: "#7f775b" },
    topLayer: { stroke: "#9a856c", text: "#876f57" },
  };

  const layerRects = [];
  const ballastTopHalfPx = (ballastTopWidthM * pxPerM) / 2;
  const ballastBottomHalfPx = (ballastBottomWidthM * pxPerM) / 2;
  const byTop = trackTopY;
  const byBottom = trackTopY + ballastH;
  const buildBallastPolygon = (trackCenterX) => {
    const bx0 = trackCenterX - ballastBottomHalfPx;
    const bx1 = trackCenterX - ballastTopHalfPx;
    const bx2 = trackCenterX + ballastTopHalfPx;
    const bx3 = trackCenterX + ballastBottomHalfPx;
    return {
      leftBottom: bx0,
      leftTop: bx1,
      rightTop: bx2,
      rightBottom: bx3,
      svg: `<polygon points="${bx1},${byTop} ${bx2},${byTop} ${bx3},${byBottom} ${bx0},${byBottom}" fill="#cbd2d8" stroke="#7f8a95" />`,
    };
  };
  const trackCentersX = [];
  const trackCenterByOffset = new Map();
  const trackCenterByItem = new Map();
  if (sequenceLayout) {
    sequenceLayout.trackItems.forEach((item) => {
      const offset = sequenceLayout.offsetByItem.get(item);
      if (!Number.isFinite(offset)) return;
      const track = {
        item,
        offset,
        visualOffset: offset - trackOffsetMidpoint,
        x: centerX - ((offset - trackOffsetMidpoint) * pxPerM),
        isMain: Boolean(item.isMain),
      };
      trackCentersX.push(track);
      trackCenterByItem.set(item, track);
      trackCenterByOffset.set(offset, track);
    });
  } else {
    allTrackCenterOffsets.forEach((offset) => {
      if (!Number.isFinite(offset)) return;
      const track = {
        offset,
        visualOffset: offset - trackOffsetMidpoint,
        x: centerX - ((offset - trackOffsetMidpoint) * pxPerM),
        isMain: Math.abs(offset) < 0.001,
      };
      trackCentersX.push(track);
      trackCenterByOffset.set(offset, track);
    });
  }
  const mainTrackLayout = trackCentersX.find((track) => Math.abs(track.offset) < 0.001)
    || trackCentersX.find((track) => track.isMain)
    || { x: centerX, offset: 0 };
  const mainTrackBallast = buildBallastPolygon(mainTrackLayout.x);
  const ballastPolygons = trackCentersX
    .map((track) => buildBallastPolygon(track.x).svg)
    .join("");
  layerRects.push(ballastPolygons);

  const bltTopY0 = byBottom;
  const bltTopY1 = byBottom;
  const blanketLeft = `<polygon points="${mainTrackBallast.leftBottom},${bltTopY0} ${mainTrackBallast.leftTop},${bltTopY1} ${mainTrackBallast.leftTop},${bltTopY1 + blanketHDraw} ${mainTrackBallast.leftBottom},${bltTopY0 + blanketHDraw}" fill="#f5eecf" stroke="#90886c" />`;
  const blanketRight = `<polygon points="${mainTrackBallast.rightTop},${bltTopY1} ${mainTrackBallast.rightBottom},${bltTopY0} ${mainTrackBallast.rightBottom},${bltTopY0 + blanketHDraw} ${mainTrackBallast.rightTop},${bltTopY1 + blanketHDraw}" fill="#f5eecf" stroke="#90886c" />`;
  const blanketCenter = `<rect x="${mainTrackBallast.leftTop}" y="${bltTopY1}" width="${mainTrackBallast.rightTop - mainTrackBallast.leftTop}" height="${blanketHDraw}" fill="#f5eecf" stroke="#90886c" />`;
  layerRects.push(`${blanketLeft}${blanketCenter}${blanketRight}`);

  const tlTopY0 = bltTopY0 + blanketHDraw;
  const tlTopY1 = bltTopY1 + blanketHDraw;
  const topLayerLeft = `<polygon points="${mainTrackBallast.leftBottom},${tlTopY0} ${mainTrackBallast.leftTop},${tlTopY1} ${mainTrackBallast.leftTop},${tlTopY1 + topLayerHDraw} ${mainTrackBallast.leftBottom},${tlTopY0 + topLayerHDraw}" fill="#edd6bf" stroke="#9a856c" />`;
  const topLayerRight = `<polygon points="${mainTrackBallast.rightTop},${tlTopY1} ${mainTrackBallast.rightBottom},${tlTopY0} ${mainTrackBallast.rightBottom},${tlTopY0 + topLayerHDraw} ${mainTrackBallast.rightTop},${tlTopY1 + topLayerHDraw}" fill="#edd6bf" stroke="#9a856c" />`;
  const topLayerCenter = `<rect x="${mainTrackBallast.leftTop}" y="${tlTopY1}" width="${mainTrackBallast.rightTop - mainTrackBallast.leftTop}" height="${topLayerHDraw}" fill="#edd6bf" stroke="#9a856c" />`;
  layerRects.push(`${topLayerLeft}${topLayerCenter}${topLayerRight}`);

  const sleeperWidthPx = Math.max(2.75 * pxPerM, 60);
  const sleeperY = trackTopY - sleeperThicknessPx;
  const railGaugePx = 1.676 * pxPerM;
  const railHeadWidthPx = Math.max(0.09 * pxPerM, 8);
  const railWebWidthPx = Math.max(0.02 * pxPerM, 4);
  const railFootWidthPx = Math.max(0.15 * pxPerM, 12);
  const railTopY = sleeperY - railHeightPx;
  const railBaseY = sleeperY;
  const buildRail = (railCenterX) => `
    <rect x="${railCenterX - (railFootWidthPx / 2)}" y="${railBaseY - Math.max(railHeightPx * 0.22, 3)}" width="${railFootWidthPx}" height="${Math.max(railHeightPx * 0.22, 3)}" rx="1.5" fill="#747f8c" stroke="#42505d" stroke-width="1.2" />
    <rect x="${railCenterX - (railWebWidthPx / 2)}" y="${railTopY + Math.max(railHeightPx * 0.18, 2)}" width="${railWebWidthPx}" height="${Math.max(railHeightPx * 0.6, 6)}" fill="#697583" stroke="#42505d" stroke-width="0.8" />
    <rect x="${railCenterX - (railHeadWidthPx / 2)}" y="${railTopY}" width="${railHeadWidthPx}" height="${Math.max(railHeightPx * 0.2, 4)}" rx="1.5" fill="#b3bcc6" stroke="#42505d" stroke-width="1.2" />
  `;
  const buildTrackAssembly = (trackCenterX, isMain = false) => {
    const sleeperX = trackCenterX - (sleeperWidthPx / 2);
    const sleeperFill = isMain ? "#c0392b" : "#838d99";
    return `
      <rect x="${sleeperX}" y="${sleeperY}" width="${sleeperWidthPx}" height="${sleeperThicknessPx}" rx="4" fill="${sleeperFill}" stroke="#42505d" stroke-width="1.4" />
      ${buildRail(trackCenterX - (railGaugePx / 2))}
      ${buildRail(trackCenterX + (railGaugePx / 2))}
    `;
  };
  const trackAssembly = showRails
    ? trackCentersX.map((track) => buildTrackAssembly(track.x, track.isMain)).join("")
    : "";
  const getPlatformTopOffsetM = (remarks) => {
    const text = String(remarks || "").toLowerCase();
    if (/goods|goods\s*platform/.test(text)) return 1.065;
    if (/low\s*level|low\s*platform/.test(text)) return 0.455;
    if (/rail\s*level|rl\s*platform/.test(text)) return 0;
    return 0.76; // default high level platform
  };
  const buildPlatformRect = (x0, widthPx, label, platformTopY, platformHeightPx) => `
    <rect x="${x0}" y="${platformTopY}" width="${Math.max(widthPx, 8)}" height="${platformHeightPx}" rx="4" fill="url(#platformHatch)" stroke="#b45309" stroke-width="1.4" />
    ${showLabels ? `<text x="${x0 + 4}" y="${platformTopY - 6}" fill="#b45309" font-size="11" font-weight="700">${label}</text>` : ""}
  `;
  const platformShapes = showPlatforms ? (sequenceLayout ? (() => {
    const ordered = sequenceLayout.orderedItems || [];
    const shapes = [];
    const findNeighborTrack = (startIndex, step) => {
      for (let i = startIndex + step; i >= 0 && i < ordered.length; i += step) {
        if (ordered[i].kind === "track") return ordered[i];
      }
      return null;
    };
    ordered.forEach((item, idx) => {
      if (item.kind !== "platform") return;
      const width = safeNum(item.row.pfWidth, 0);
      if (!width) return;
      const before = findNeighborTrack(idx, -1);
      const after = findNeighborTrack(idx, 1);
      const remarks = String(item.row.remarks || "");
      const wantsIsland = /island|is-land|islnd|is\s*land/i.test(remarks);
      let isIsland = Boolean(before && after);
      if (wantsIsland && before && after) isIsland = true;
      if (isIsland && (!before || !after)) isIsland = false;
      if (isIsland && before && after) {
        const offA = sequenceLayout.offsetByItem.get(before);
        const offB = sequenceLayout.offsetByItem.get(after);
        if (!Number.isFinite(offA) || !Number.isFinite(offB)) return;
        const gapPx = Math.abs(offA - offB) * pxPerM;
        const edgeClearPx = ballastTopHalfPx + 6;
        const maxWidthPx = Math.max(gapPx - (2 * edgeClearPx), 0);
        if (maxWidthPx <= 8) return;
        const platformCenterOffset = (offA + offB) / 2;
        const platformCenterX = centerX - ((platformCenterOffset - trackOffsetMidpoint) * pxPerM);
        const widthPx = Math.min(Math.max(width * pxPerM, 8), maxWidthPx);
        const offsetM = getPlatformTopOffsetM(item.row.remarks);
        const platformTopY = railTopY - (offsetM * pxPerM);
        const platformHeightPx = Math.max(topY - platformTopY, 8);
        shapes.push(buildPlatformRect(platformCenterX - (widthPx / 2), widthPx, "ISLAND PF", platformTopY, platformHeightPx));
        return;
      }
      const anchor = before || after;
      if (!anchor) return;
      const anchorTrack = trackCenterByItem.get(anchor);
      if (!anchorTrack) return;
      const side = item.side || (anchorTrack.offset >= 0 ? "Left" : "Right");
      const edgeOffset = ballastTopHalfPx + 6;
      const widthPx = Math.max(width * pxPerM, 8);
      let x0 = anchorTrack.x + edgeOffset;
      if (side === "Left") x0 = anchorTrack.x - edgeOffset - widthPx;
      const offsetM = getPlatformTopOffsetM(item.row.remarks);
      const platformTopY = railTopY - (offsetM * pxPerM);
      const platformHeightPx = Math.max(topY - platformTopY, 8);
      shapes.push(buildPlatformRect(x0, widthPx, wantsIsland ? "ISLAND PF" : "PF", platformTopY, platformHeightPx));
    });
    return shapes.join("");
  })() : (() => {
    const activePlatforms = (state.loopPlatformRows || [])
      .filter((lp) =>
        Number.isFinite(lp.pfStartCh) &&
        Number.isFinite(lp.pfEndCh) &&
        row.chainage >= lp.pfStartCh &&
        row.chainage <= lp.pfEndCh &&
        safeNum(lp.pfWidth, 0) > 0,
      )
      .map((lp) => {
        const side = normalizeLoopSide(lp.side);
        const rawOffset = safeNum(lp.tc, 0);
        const signedOffset = rawOffset ? (side === "Right" ? -rawOffset : rawOffset) : 0;
        return {
          width: safeNum(lp.pfWidth, 0),
          side,
          isIsland: /island|is-land|islnd|is\s*land/i.test(String(lp.remarks || "")),
          offset: signedOffset,
          remarks: String(lp.remarks || ""),
        };
      })
      .filter((pf) => pf.side || pf.offset);
    const leftOffsets = allTrackCenterOffsets.filter((o) => o > 0).sort((a, b) => a - b);
    const rightOffsets = allTrackCenterOffsets.filter((o) => o < 0).sort((a, b) => b - a);
    const nearestLeft = leftOffsets[0] ?? 0;
    const nearestRight = rightOffsets[0] ?? 0;
    const outerLeft = leftOffsets.length ? leftOffsets[leftOffsets.length - 1] : 0;
    const outerRight = rightOffsets.length ? rightOffsets[rightOffsets.length - 1] : 0;
    const platformBySide = {
      Left: { island: null, side: null },
      Right: { island: null, side: null },
    };
    activePlatforms.forEach((pf) => {
      const side = pf.side === "Right" ? "Right" : "Left";
      const slot = pf.isIsland ? "island" : "side";
      const current = platformBySide[side][slot];
      if (!current || pf.width > current.width) {
        platformBySide[side][slot] = pf;
      }
    });
    return ["Left", "Right"].map((sideKey) => {
      const bucket = platformBySide[sideKey];
      if (!bucket) return "";
      const shapes = [];
      if (bucket.island && (sideKey === "Left" ? nearestLeft : nearestRight)) {
        const near = sideKey === "Left" ? nearestLeft : nearestRight;
        const gapPx = Math.abs(near - 0) * pxPerM;
        const edgeClearPx = ballastTopHalfPx + 6;
        const maxWidthPx = Math.max(gapPx - (2 * edgeClearPx), 0);
        if (maxWidthPx > 8) {
          const platformCenterOffset = (0 + near) / 2;
          const platformCenterX = centerX - ((platformCenterOffset - trackOffsetMidpoint) * pxPerM);
          const widthPx = Math.min(Math.max(bucket.island.width * pxPerM, 8), maxWidthPx);
          const offsetM = getPlatformTopOffsetM(bucket.island.remarks);
          const platformTopY = railTopY - (offsetM * pxPerM);
          const platformHeightPx = Math.max(topY - platformTopY, 8);
          shapes.push(buildPlatformRect(platformCenterX - (widthPx / 2), widthPx, "ISLAND PF", platformTopY, platformHeightPx));
        }
      }
      if (bucket.side) {
        const anchorOffset = sideKey === "Left" ? outerLeft : outerRight;
        const track = trackCenterByOffset.get(anchorOffset) || mainTrackLayout;
        const edgeOffset = ballastTopHalfPx + 6;
        const widthPx = Math.max(bucket.side.width * pxPerM, 8);
        let x0 = track.x + edgeOffset;
        if (sideKey === "Left") x0 = track.x - edgeOffset - widthPx;
        const offsetM = getPlatformTopOffsetM(bucket.side.remarks);
        const platformTopY = railTopY - (offsetM * pxPerM);
        const platformHeightPx = Math.max(topY - platformTopY, 8);
        shapes.push(buildPlatformRect(x0, widthPx, "PF", platformTopY, platformHeightPx));
      }
      return shapes.join("");
    }).join("");
  })()) : "";
  const trackCentersSorted = [...trackCentersX].sort((a, b) => a.x - b.x);
  const trackCenterDims = showLabels ? trackCentersSorted
    .slice(1)
    .map((track, index) => {
      const prev = trackCentersSorted[index];
      const dimY = trackTopY - 106 - (index * 22);
      return `
        <line x1="${prev.x}" y1="${trackTopY - 6}" x2="${prev.x}" y2="${dimY + 4}" stroke="#53718b" stroke-dasharray="4 4" />
        <line x1="${track.x}" y1="${trackTopY - 6}" x2="${track.x}" y2="${dimY + 4}" stroke="#53718b" stroke-dasharray="4 4" />
        ${drawDim(prev.x, track.x, dimY, `TC ${r3(Math.abs(track.offset - prev.offset))} m`)}
      `;
    })
    .join("") : "";

  if (row.bank > 0) {
    const halfBlanketBottom = halfTop + (blanketHDraw * slopeHV);
    const halfTopLayerBottom = halfTop + ((blanketHDraw + topLayerHDraw) * slopeHV);
    layerRects.push(`
      <polygon points="${centerX - halfTop},${topY} ${centerX + halfTop},${topY} ${centerX + halfBlanketBottom},${topY + blanketHDraw} ${centerX - halfBlanketBottom},${topY + blanketHDraw}"
      fill="#f5eecf" stroke="#90886c" />`);
    layerRects.push(`
      <polygon points="${centerX - halfBlanketBottom},${topY + blanketHDraw} ${centerX + halfBlanketBottom},${topY + blanketHDraw} ${centerX + halfTopLayerBottom},${topY + blanketHDraw + topLayerHDraw} ${centerX - halfTopLayerBottom},${topY + blanketHDraw + topLayerHDraw}"
      fill="#edd6bf" stroke="#9a856c" />`);
  } else {
    layerRects.push(`<rect x="${centerX - halfTop}" y="${topY}" width="${halfTop * 2}" height="${blanketH}" fill="#f5eecf" stroke="#90886c" />`);
    layerRects.push(`<rect x="${centerX - halfTop}" y="${topY + blanketH}" width="${halfTop * 2}" height="${topLayerH}" fill="#edd6bf" stroke="#9a856c" />`);
  }

  const bermLabel = fmtDim(s.bermWidth);
  let cutPoly = "";
  let fillPoly = "";
  let berms = "";
  let drainOverlays = "";

  if (row.type === "CUTTING" && row.cut > 0) {
    // Detailed Cutting profile with side drains and berms based on reference drawing
    const drainW = s.drainWidth ? (s.drainWidth * pxPerM) : (1.2 * pxPerM);
    const drainH = s.drainHeight ? (s.drainHeight * pxPerM) : (1.0 * pxPerM);
    const leftPts = [];
    const rightPts = [];
    const bermDimSnippets = [];

    // Formation edge and bottom drain
    const formLeft = centerX - halfTop;
    const formRight = centerX + halfTop;
    leftPts.push({ x: formLeft, y: topY });
    rightPts.push({ x: formRight, y: topY });

    // U-shaped side drain geometry (concrete overlay)
    const drainBotL_R = formLeft;          // inner vertical
    const drainBotL_L = formLeft - drainW; // outer vertical
    leftPts.push({ x: drainBotL_R, y: topY + drainH });
    leftPts.push({ x: drainBotL_L, y: topY + drainH });
    leftPts.push({ x: drainBotL_L, y: topY }); // outer top edge of drain

    const drainBotR_L = formRight;         // inner vertical
    const drainBotR_R = formRight + drainW;
    rightPts.push({ x: drainBotR_L, y: topY + drainH });
    rightPts.push({ x: drainBotR_R, y: topY + drainH });
    rightPts.push({ x: drainBotR_R, y: topY });

    const drainWall = Math.min(drainW * 0.25, 0.25 * pxPerM);
    const buildUDrain = (xInner, xOuter) => {
      const x0 = Math.min(xInner, xOuter);
      const x1 = Math.max(xInner, xOuter);
      const y0 = topY;
      const y1 = topY + drainH;
      const innerX0 = x0 + drainWall;
      const innerX1 = x1 - drainWall;
      const innerY0 = y0 + drainWall;
      const innerY1 = y1 - drainWall;
      return `
        <rect x="${x0}" y="${y0}" width="${drainWall}" height="${y1 - y0}" fill="#d9dee4" stroke="#8b98a6" stroke-width="1.2" />
        <rect x="${x1 - drainWall}" y="${y0}" width="${drainWall}" height="${y1 - y0}" fill="#d9dee4" stroke="#8b98a6" stroke-width="1.2" />
        <rect x="${innerX0}" y="${y1 - drainWall}" width="${innerX1 - innerX0}" height="${drainWall}" fill="#d9dee4" stroke="#8b98a6" stroke-width="1.2" />
      `;
    };
    drainOverlays = `
      ${buildUDrain(formLeft, formLeft - drainW)}
      ${buildUDrain(formRight, formRight + drainW)}
    `;
    if (!showDrains) drainOverlays = "";

    // Catch Water Drains on Berms (approx 0.5m x 0.5m based on image)
    const cwDrainW = 0.5 * pxPerM;
    const cwDrainH = 0.5 * pxPerM;

    // Calculate berms going UP the cut
    const bermCount = row.cut >= bermSecondary ? 2 : (row.cut >= bermPrimary ? 1 : 0);
    const bermFractions = bermCount === 2 ? [0.35, 0.72] : (bermCount === 1 ? [0.58] : []);

    let currentHeight = 0;
    let currentLeft = drainBotL_L;
    let currentRight = drainBotR_R;

    bermFractions.forEach((f, idx) => {
      const targetHeight = row.cut * f;
      const deltaHeight = targetHeight - currentHeight;
      const deltaRun = slopeHV * deltaHeight * pxPerM;
      const y = topY - targetHeight * pxPerM;
      const lx = currentLeft - deltaRun;
      const rx = currentRight + deltaRun;

      // Slope up to berm
      leftPts.push({ x: lx, y });
      rightPts.push({ x: rx, y });

      // Berm bench flat
      // Left Berm
      const lxBenchEdge = lx - bermW;
      leftPts.push({ x: lxBenchEdge, y }); // flat berm

      // Right Berm
      const rxBenchEdge = rx + bermW;
      rightPts.push({ x: rxBenchEdge, y });

      const dimY = y - 18;
      bermDimSnippets.push(drawDim(lxBenchEdge, lx, dimY, bermLabel));
      bermDimSnippets.push(drawDim(rx, rxBenchEdge, dimY, bermLabel));
      bermDimSnippets.push(`<text x="${(lxBenchEdge + lx) / 2}" y="${y - 28}" text-anchor="middle" fill="#385a48" font-size="11" font-weight="700">BERM</text>`);
      bermDimSnippets.push(`<text x="${(rx + rxBenchEdge) / 2}" y="${y - 28}" text-anchor="middle" fill="#385a48" font-size="11" font-weight="700">BERM</text>`);

      currentHeight = targetHeight;
      currentLeft = lxBenchEdge;
      currentRight = rxBenchEdge;
    });

    // Final slope to ground level (cut bottom is actually the top G.L. conceptually)
    const deltaHeight = row.cut - currentHeight;
    const deltaRun = slopeHV * deltaHeight * pxPerM;
    const xLeftTop = currentLeft - deltaRun;
    const xRightTop = currentRight + deltaRun;
    leftPts.push({ x: xLeftTop, y: cutBottomY });
    rightPts.push({ x: xRightTop, y: cutBottomY });
    toeLeftX = xLeftTop;
    toeRightX = xRightTop;
    toeBaseY = cutBottomY;

    // Close polygon
    const polyPts = [...leftPts.reverse(), ...rightPts].map((p) => `${p.x},${p.y}`).join(" ");
    cutPoly = `<polygon points="${polyPts}" fill="#f0e2e2" stroke="#8a6f72" />`;
    berms = showLabels ? bermDimSnippets.join("") : "";

  } else if (row.bank > 0) {
    const bermCount = row.bank >= bermSecondary ? 2 : (row.bank >= bermPrimary ? 1 : 0);
    const bermFractions = bermCount === 2 ? [0.35, 0.72] : (bermCount === 1 ? [0.58] : []);
    const leftPts = [{ x: centerX - halfTop, y: topY }];
    const rightPts = [{ x: centerX + halfTop, y: topY }];
    const bermDimSnippets = [];
    let currentDepth = 0;
    let currentLeft = centerX - halfTop;
    let currentRight = centerX + halfTop;
    bermFractions.forEach((f, idx) => {
      const targetDepth = row.bank * f;
      const deltaDepth = targetDepth - currentDepth;
      const deltaRun = slopeHV * deltaDepth * pxPerM;
      const y = topY + targetDepth * pxPerM;
      const lx = currentLeft - deltaRun;
      const rx = currentRight + deltaRun;
      // slope segment to berm elevation
      leftPts.push({ x: lx, y });
      rightPts.push({ x: rx, y });
      // berm bench
      const lxBench = lx - bermW;
      const rxBench = rx + bermW;
      leftPts.push({ x: lxBench, y });
      rightPts.push({ x: rxBench, y });
      const dimY = y + 26;
      bermDimSnippets.push(drawDim(lxBench, lx, dimY, bermLabel));
      bermDimSnippets.push(drawDim(rx, rxBench, dimY, bermLabel));
      bermDimSnippets.push(`<text x="${(lxBench + lx) / 2}" y="${y - 10}" text-anchor="middle" fill="#385a48" font-size="11" font-weight="700">BERM</text>`);
      bermDimSnippets.push(`<text x="${(rx + rxBench) / 2}" y="${y - 10}" text-anchor="middle" fill="#385a48" font-size="11" font-weight="700">BERM</text>`);
      currentDepth = targetDepth;
      currentLeft = lxBench;
      currentRight = rxBench;
    });
    // final slope to toe from last bench
    const deltaDepth = row.bank - currentDepth;
    const deltaRun = slopeHV * deltaDepth * pxPerM;
    const xLeftToe = currentLeft - deltaRun;
    const xRightToe = currentRight + deltaRun;
    leftPts.push({ x: xLeftToe, y: toeY });
    rightPts.push({ x: xRightToe, y: toeY });
    toeLeftX = xLeftToe;
    toeRightX = xRightToe;
    toeBaseY = toeY;
    const polyPts = [...leftPts, ...rightPts.reverse()].map((p) => `${p.x},${p.y}`).join(" ");
    fillPoly = `<polygon points="${polyPts}" fill="url(#embFillHatch)" stroke="#69786a" />`;
    berms = showLabels ? bermDimSnippets.join("") : "";
  }

  const segLeftOuter = 1.2;
  const segLeftShoulder = 0.35;
  const segRightShoulder = 0.35;
  const segRightOuter = 1.2;
  const segMiddle = Math.max(drawTopWidthM - (segLeftOuter + segLeftShoulder + segRightShoulder + segRightOuter), 0.5);
  const segs = [segLeftOuter, segLeftShoulder, segMiddle, segRightShoulder, segRightOuter];
  const segLabels = [`${r3(segLeftOuter)}m`, `${r3(segLeftShoulder)}m`, `${r3(segMiddle)}m`, `${r3(segRightShoulder)}m`, `${r3(segRightOuter)}m`];
  let cursorX = centerX - halfTop;
  const segDims = [];
  const segMidPoints = [];
  for (let i = 0; i < segs.length; i += 1) {
    const w = segs[i] * pxPerM;
    const nextX = cursorX + w;
    segDims.push(drawDimLine(cursorX, nextX, trackTopY - 42));
    segDims.push(`<line x1="${cursorX}" y1="${trackTopY - 24}" x2="${cursorX}" y2="${topY + 2}" stroke="#7c93a8" stroke-dasharray="3 3" />`);
    segMidPoints.push((cursorX + nextX) / 2);
    cursorX = nextX;
  }
  segDims.push(`<line x1="${centerX + halfTop}" y1="${trackTopY - 24}" x2="${centerX + halfTop}" y2="${topY + 2}" stroke="#7c93a8" stroke-dasharray="3 3" />`);
  const segLabelRows = segLabels.map((label, i) => {
    const y = (i % 2 === 0) ? (trackTopY - 54) : (trackTopY - 64);
    return `<text x="${segMidPoints[i]}" y="${y}" text-anchor="middle" fill="#2f4d6a" font-size="11" font-weight="700">${label}</text>`;
  }).join("");
  const bodyLabel = row.type === "CUTTING" ? "Cutting Profile" : "Embankment Fill";
  const bodySubLabel = row.type === "CUTTING" ? "" : `(Soil Category: ${sqName})`;
  const bodyYRef = row.type === "CUTTING" ? (topY + 120) : ((topY + toeY) / 2);
  const layerLegendX = 46;
  const layerLegendY = 74;
  const layerLegendItems = [
    { fill: "#cbd2d8", stroke: layerCalloutStyles.ballast.stroke, text: `${r3(ballastThickness)} m Ballast Cushion` },
    { fill: "#f5eecf", stroke: layerCalloutStyles.blanket.stroke, text: `Blanket (${sqName}) ${r3(blanketRuleThickness)} m` },
    { fill: "#edd6bf", stroke: layerCalloutStyles.topLayer.stroke, text: `Top Layer Fill ${r3(topLayerEffectiveThickness)} m` },
  ];
  const layerLegend = showLabels ? `
    <g>
      <rect x="${layerLegendX - 12}" y="${layerLegendY - 26}" width="270" height="${layerLegendItems.length * 28 + 18}" rx="14" fill="rgba(255,255,255,0.92)" stroke="#d7e0e8" />
      <text x="${layerLegendX}" y="${layerLegendY - 6}" fill="#223240" font-size="13" font-weight="700">Layer Legend</text>
      ${layerLegendItems.map((item, index) => {
        const itemY = layerLegendY + (index * 28);
        return `
          <rect x="${layerLegendX}" y="${itemY}" width="18" height="12" rx="3" fill="${item.fill}" stroke="${item.stroke}" />
          <text x="${layerLegendX + 28}" y="${itemY + 10}" fill="#33475a" font-size="12">${item.text}</text>
        `;
      }).join("")}
    </g>
  ` : "";
  const trackLegendX = 46;
  const trackLegendY = layerLegendY + (layerLegendItems.length * 28) + 34;
  const trackLegendItems = [
    { fill: "#b3bcc6", stroke: "#42505d", text: "Rail" },
    { fill: "#838d99", stroke: "#42505d", text: "Sleeper" },
    { fill: "#c0392b", stroke: "#42505d", text: "Main Line Sleeper" },
  ];
  const trackLegend = showLabels && showRails ? `
    <g>
      <rect x="${trackLegendX - 12}" y="${trackLegendY - 26}" width="180" height="${trackLegendItems.length * 28 + 18}" rx="14" fill="rgba(255,255,255,0.92)" stroke="#d7e0e8" />
      <text x="${trackLegendX}" y="${trackLegendY - 6}" fill="#223240" font-size="13" font-weight="700">Track Legend</text>
      ${trackLegendItems.map((item, index) => {
        const itemY = trackLegendY + (index * 28);
        return `
          <rect x="${trackLegendX}" y="${itemY}" width="18" height="12" rx="3" fill="${item.fill}" stroke="${item.stroke}" />
          <text x="${trackLegendX + 28}" y="${itemY + 10}" fill="#33475a" font-size="12">${item.text}</text>
        `;
      }).join("")}
    </g>
  ` : "";
  const slopeLabels = showLabels && row.bank > 0
    ? `
      <text x="${leftToeX - 46}" y="${topY + Math.max(fillH * 0.55, 24)}" fill="#2f4d6a" font-size="12" font-weight="700">1:${r3(slopeHV)}</text>
      <text x="${rightToeX + 16}" y="${topY + Math.max(fillH * 0.55, 24)}" fill="#2f4d6a" font-size="12" font-weight="700">1:${r3(slopeHV)}</text>
    `
    : "";
  const levelLabels = [
    `Formation RL: ${r3(fl)} m`,
    `Ballast Top RL: ${r3(fl + ballastThickness)} m`,
    `Blanket Bottom RL: ${r3(fl - blanketRuleThickness)} m`,
    `Top Layer Bottom RL: ${r3(fl - blanketRuleThickness - topLayerEffectiveThickness)} m`,
    `Ground RL (GL): ${r3(gl)} m`,
  ];
  const levelText = levelLabels.map((txt, i) => {
    const y = 80 + i * 22;
    const x = svgW - 260;
    return `<text x="${x}" y="${y}" fill="#2b3f52" font-size="12" font-weight="700">${txt}</text>`;
  }).join("");
  const levelCard = showLabels ? (() => {
    const cardX = svgW - 330;
    const cardY = 48;
    const padding = 12;
    const lineHeight = 18;
    const title = "RL Summary";
    const lines = levelLabels;
    const boxWidth = 260;
    const boxHeight = padding + 20 + (lines.length * lineHeight) + padding;
    const textX = cardX + padding;
    let cursorY = cardY + padding + 14;
    const lineTexts = lines.map((line) => {
      const y = cursorY;
      cursorY += lineHeight;
      return `<text x="${textX}" y="${y}" fill="#2b3f52" font-size="11">${escapeHtml(line)}</text>`;
    }).join("");
    return `
      <g>
        <rect x="${cardX}" y="${cardY}" width="${boxWidth}" height="${boxHeight}" rx="12" fill="rgba(255,255,255,0.92)" stroke="#ccd6df" />
        <text x="${textX}" y="${cardY + padding + 2}" fill="#203243" font-size="12" font-weight="700">${title}</text>
        ${lineTexts}
      </g>
    `;
  })() : "";
  const demarcRightX = svgW - 210;
  const glLeftX = centerX - 520;
  const hflLeftX = centerX - 410;
  const toeDimY = Math.max(groundY + 70, toeBaseY + 60);
  const toeSpanM = Math.abs(toeRightX - toeLeftX) / pxPerM;
  const toeDimLabel = `${r3(toeSpanM)} m`;
  const formationDim = showLabels ? drawDim(centerX - halfTop, centerX + halfTop, trackTopY - 52, `${r3(drawTopWidthM)} m`) : "";
  const toeDimLines = showLabels ? `
    <line x1="${toeLeftX}" y1="${toeBaseY}" x2="${toeLeftX}" y2="${toeDimY - 6}" stroke="#53718b" stroke-dasharray="4 4" />
    <line x1="${toeRightX}" y1="${toeBaseY}" x2="${toeRightX}" y2="${toeDimY - 6}" stroke="#53718b" stroke-dasharray="4 4" />
    ${drawDim(toeLeftX, toeRightX, toeDimY, toeDimLabel)}
  ` : "";
  const topFormationLabel = showLabels ? `<text x="${centerX + halfTop + 16}" y="${topY - 2}" fill="#253748" font-size="14" font-weight="700">Top of Formation</text>` : "";
  const naturalLabel = showLabels ? `<text x="${centerX - 95}" y="${groundY + 86}" fill="#3f4a55" font-size="13">Natural Ground / Subsoil</text>` : "";
  const crossTitleLabel = showLabels ? `<text x="42" y="40" fill="#1f2e3a" font-size="16" font-weight="700">Cross Section of Track</text>` : "";
  const bodyLabelBlock = showLabels ? `
    <text x="${centerX - 130}" y="${bodyYRef + 26}" fill="#344553" font-size="13" font-weight="700">${bodyLabel}</text>
    <text x="${centerX - 132}" y="${bodyYRef + 42}" fill="#344553" font-size="12">${bodySubLabel}</text>
  ` : "";
  const sequenceDebugOverlay = showLabels && sequenceLayout ? (() => {
    const stationLabel = row.stationName || row.station || "";
    const header = `Sequence Debug${stationLabel ? `: ${stationLabel}` : ""}`;
    const lines = (sequenceLayout.orderedItems || []).map((item) => {
      if (item.kind === "track") {
        const offset = sequenceLayout.offsetByItem.get(item);
        const tc = safeNum(item.row.tc, 0);
        const lineType = item.lineType || "Track";
        const offLabel = Number.isFinite(offset) ? r3(offset) : "-";
        return `#${item.order} ${lineType} ${item.side || "-"} tc=${r3(tc)} off=${offLabel}`;
      }
      const pfWidth = safeNum(item.row.pfWidth, 0);
      const remarks = String(item.row.remarks || "");
      const isIsland = /island|is-land|islnd|is\\s*land/i.test(remarks);
      return `#${item.order} ${isIsland ? "Island PF" : "Platform"} ${item.side || "-"} w=${r3(pfWidth)}`;
    });
    if (!lines.length) return "";
    const boxX = layerLegendX + 300;
    const boxY = layerLegendY - 26;
    const padding = 12;
    const lineHeight = 16;
    const titleHeight = 18;
    const boxWidth = 480;
    const boxHeight = padding + titleHeight + (lines.length * lineHeight) + padding;
    const textX = boxX + padding;
    let cursorY = boxY + padding + titleHeight;
    const lineTexts = lines.map((line) => {
      const y = cursorY;
      cursorY += lineHeight;
      return `<text x="${textX}" y="${y}" fill="#2a3a47" font-size="11">${escapeHtml(line)}</text>`;
    }).join("");
    return `
      <g>
        <rect x="${boxX}" y="${boxY}" width="${boxWidth}" height="${boxHeight}" rx="10" fill="rgba(255,255,255,0.92)" stroke="#ccd6df" />
        <text x="${textX}" y="${boxY + padding + 4}" fill="#203243" font-size="12" font-weight="700">${escapeHtml(header)}</text>
        ${lineTexts}
      </g>
    `;
  })() : "";

  targetEl.innerHTML = `
    <rect x="-120000" y="-120000" width="240000" height="240000" fill="#f8fcff" />
    <line x1="${glLeftX}" y1="${groundY}" x2="${demarcRightX}" y2="${groundY}" stroke="#5d6b77" stroke-dasharray="8 7" stroke-width="1.8" />
    ${showLabels ? `<text x="${demarcRightX + 18}" y="${groundY - 7}" fill="#374b5d" font-size="13" font-weight="700">G.L.</text>` : ""}
    <line x1="${hflLeftX}" y1="${hflY}" x2="${demarcRightX}" y2="${hflY}" stroke="#6d7680" stroke-dasharray="12 8" stroke-width="1.4" />
    ${showLabels ? `<text x="${demarcRightX + 18}" y="${hflY - 7}" fill="#4f5f6d" font-size="12">H.F.L.</text>` : ""}
    ${naturalLabel}
    ${fillPoly}
    ${cutPoly}
    ${drainOverlays}
    ${berms}
    ${layerRects.join("")}
    ${platformShapes}
    ${trackAssembly}
    <line x1="${centerX - halfTop}" y1="${topY}" x2="${centerX + halfTop}" y2="${topY}" stroke="#2e3b49" stroke-width="2.4" />
    ${topFormationLabel}
    ${formationDim}
    ${trackCenterDims}
    ${toeDimLines}
    
    ${layerLegend}
    ${trackLegend}
    ${showLabels ? `<line x1="${centerX}" y1="${topY}" x2="${centerX}" y2="${centerRef}" stroke="#2f4d6a" stroke-width="1.8" marker-start="url(#arrowSmall)" marker-end="url(#arrowSmall)" />` : ""}
    ${showLabels ? `<text x="${centerX + 10}" y="${(topY + centerRef) / 2}" fill="#2f4d6a" font-size="12" font-weight="700">${row.type === "CUTTING" ? `Cut = ${r3(row.cut)} m` : `e = ${r3(row.bank)} m`}</text>` : ""}
    ${slopeLabels}
    ${bodyLabelBlock}
    ${levelCard}
    ${sequenceDebugOverlay}
    ${crossTitleLabel}
    <defs>
      <pattern id="embFillHatch" patternUnits="userSpaceOnUse" width="10" height="10" patternTransform="rotate(20)">
        <rect width="10" height="10" fill="#dfe9de" />
        <path d="M0,0 L0,10" stroke="#9cb3a3" stroke-width="1" />
      </pattern>
      <pattern id="platformHatch" patternUnits="userSpaceOnUse" width="8" height="8" patternTransform="rotate(25)">
        <rect width="8" height="8" fill="#f6b54f" />
        <path d="M0,0 L0,8" stroke="#c97a14" stroke-width="1" />
      </pattern>
      <marker id="arrowSmall" markerWidth="7" markerHeight="7" refX="3.5" refY="3.5" orient="auto-start-reverse">
        <path d="M0,0 L7,3.5 L0,7 z" fill="#2f4d6a"/>
      </marker>
    </defs>
  `;
  if (targetEl === els.crossSvg) {
    resetCrossView();
    ensureCrossSectionModalOpen();
  }
}

function bindEvents() {
  if (els.loginForm) {
    els.loginForm.addEventListener("submit", (event) => {
      event.preventDefault();
      attemptLogin();
    });
  }

  if (els.loginSubmitBtn) {
    els.loginSubmitBtn.addEventListener("click", (event) => {
      event.preventDefault();
      attemptLogin();
    });
  }

  if (els.logoutBtn) {
    els.logoutBtn.addEventListener("click", () => {
      logout();
    });
  }

  if (els.workNav) {
    els.workNav.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-work-page-btn]");
      if (!btn) return;
      setWorkPage(btn.dataset.workPageBtn);
    });
  }

  if (els.themeToggleCheckbox) {
    els.themeToggleCheckbox.addEventListener("change", (e) => {
      // Dark is default (unchecked). Checked = light mode.
      document.documentElement.classList.toggle("light", e.target.checked);
    });
    // Set initial toggle state (unchecked = dark mode default)
    els.themeToggleCheckbox.checked = document.documentElement.classList.contains("light");
  }

  // --- Project Settings Tab Switching ---
  const settingsTabs = document.querySelectorAll(".settings-tab");
  const settingsSections = document.querySelectorAll(".settings-section");
  settingsTabs.forEach(tab => {
    tab.addEventListener("click", () => {
      // Remove active from all tabs
      settingsTabs.forEach(t => t.classList.remove("active"));
      // Hide all sections
      settingsSections.forEach(s => s.classList.remove("active"));
      
      // Activate clicked tab
      tab.classList.add("active");
      
      // Show corresponding section
      const targetId = `settings-tab-${tab.dataset.settingsTab}`;
      const targetSection = document.getElementById(targetId);
      if (targetSection) {
        targetSection.classList.add("active");
      }
    });
  });

  if (els.sidebarToggle && els.appLayout) {
    els.sidebarToggle.addEventListener("click", () => {
      els.appLayout.classList.add("sidebar-collapsed");
    });

    const brandIcon = document.querySelector(".brand-icon");
    if (brandIcon) {
      brandIcon.addEventListener("click", () => {
        if (els.appLayout.classList.contains("sidebar-collapsed")) {
          els.appLayout.classList.remove("sidebar-collapsed");
        }
      });
    }
  }

  if (els.openExportModalBtn) {
    els.openExportModalBtn.addEventListener("click", () => {
      if (!state.calcRows.length) {
        alert("No calculation data available to export.");
        return;
      }
      if (els.reportSignName) els.reportSignName.value = state.settings.reportBrand?.signName || "";
      if (els.reportSignTitle) els.reportSignTitle.value = state.settings.reportBrand?.signTitle || "";
      els.exportModal.showModal();
    });
  }

  if (els.confirmExportBtn) {
    els.confirmExportBtn.addEventListener("click", () => {
      const rowLimitInput = document.getElementById("exportRowLimit");
      const rowLimit = rowLimitInput ? parseInt(rowLimitInput.value) || 0 : 0;
      const rangeStartRaw = String(document.getElementById("exportRangeStart")?.value || "").trim();
      const rangeEndRaw = String(document.getElementById("exportRangeEnd")?.value || "").trim();
      const rangeStart = rangeStartRaw ? parseChainage(rangeStartRaw) : NaN;
      const rangeEnd = rangeEndRaw ? parseChainage(rangeEndRaw) : NaN;
      const exportRange = (Number.isFinite(rangeStart) || Number.isFinite(rangeEnd))
        ? { start: Number.isFinite(rangeStart) ? rangeStart : NaN, end: Number.isFinite(rangeEnd) ? rangeEnd : NaN }
        : null;

      const options = {
        calcSheet: document.getElementById("includeCalcSheet").checked,
        crossSections: document.getElementById("includeCrossSections").checked,
        rollDiagram: document.getElementById("includeRollDiagram").checked,
        includeProfileSection: document.getElementById("includeProfileSection")?.checked ?? true,
        includeStandardsSection: document.getElementById("includeStandardsSection")?.checked ?? true,
        includeMaterialSection: document.getElementById("includeMaterialSection")?.checked ?? false,
        includeQualitySection: document.getElementById("includeQualitySection")?.checked ?? true,
        rowLimit: rowLimit,
        fastMode: document.getElementById("exportFastMode")?.checked ?? false,
        exportRange,
      };
      els.exportModal.close();
      generateProjectReport(options);
    });
  }

  const exportRangeStart = document.getElementById("exportRangeStart");
  const exportRangeEnd = document.getElementById("exportRangeEnd");
  const exportRangeDistance = document.getElementById("exportRangeDistance");
  const updateExportRangeDistance = () => {
    if (!exportRangeDistance) return;
    const startVal = parseChainage(String(exportRangeStart?.value || ""));
    const endVal = parseChainage(String(exportRangeEnd?.value || ""));
    if (Number.isFinite(startVal) && Number.isFinite(endVal)) {
      const dist = Math.abs(endVal - startVal);
      const km = dist >= 1000 ? ` (${r3(dist / 1000)} km)` : "";
      exportRangeDistance.textContent = `Distance: ${r3(dist)} m${km}`;
    } else {
      exportRangeDistance.textContent = "Distance: —";
    }
  };
  if (exportRangeStart) exportRangeStart.addEventListener("input", updateExportRangeDistance);
  if (exportRangeEnd) exportRangeEnd.addEventListener("input", updateExportRangeDistance);

  if (els.cancelExportBtn) {
    els.cancelExportBtn.addEventListener("click", () => els.exportModal.close());
  }

  const readReportAsset = (inputEl, key) => {
    if (!inputEl) return;
    inputEl.addEventListener("change", () => {
      const file = inputEl.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        if (!state.settings.reportBrand) state.settings.reportBrand = { ...state.defaultSettings.reportBrand };
        state.settings.reportBrand[key] = reader.result;
      };
      reader.readAsDataURL(file);
    });
  };
  readReportAsset(els.reportLogoInput, "logo");
  readReportAsset(els.reportSignatureInput, "signature");
  if (els.reportSignName) {
    els.reportSignName.addEventListener("input", () => {
      if (!state.settings.reportBrand) state.settings.reportBrand = { ...state.defaultSettings.reportBrand };
      state.settings.reportBrand.signName = String(els.reportSignName.value || "");
    });
  }
  if (els.reportSignTitle) {
    els.reportSignTitle.addEventListener("input", () => {
      if (!state.settings.reportBrand) state.settings.reportBrand = { ...state.defaultSettings.reportBrand };
      state.settings.reportBrand.signTitle = String(els.reportSignTitle.value || "");
    });
  }

  // Populate date display
  const dateEl = document.getElementById("dateDisplay");
  if (dateEl) {
    const now = new Date();
    const days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    dateEl.textContent = `${days[now.getDay()]}, ${now.getDate()} ${months[now.getMonth()]} ${now.getFullYear()} `;
  }

  // --- Strict Verification Engine ---
  function validateProjectData() {
    const errors = [];
    if (!state.project.name) errors.push("- Project Name is missing.");
    if (!state.rawRows || state.rawRows.length === 0) errors.push("- No Earthwork Levels data uploaded.");

    // Check for calculation errors
    let nanFound = false;
    let logicError = null;
    let invalidChainage = null;

    if (state.calcRows && state.calcRows.length > 0) {
      for (let i = 0; i < state.calcRows.length; i++) {
        const row = state.calcRows[i];

        // Check essential numbers
        if (!Number.isFinite(row.chainage)) {
          invalidChainage = `Row ${i + 1} `;
        }

        // If cutVol or fillVol area/vol resulted in NaN
        if (Number.isNaN(row.cutVol) || Number.isNaN(row.fillVol) || Number.isNaN(row.bank) || Number.isNaN(row.ewArea) || Number.isNaN(row.areaDiff)) {
          nanFound = true;
          logicError = `Calculation Error at Chainage ${row.chainage} `;
          break;
        }
      }
    } else {
      errors.push("- No calculations generated. Please recalculate first.");
    }

    if (invalidChainage) errors.push(`- Invalid Chainage found near calculation ${invalidChainage}. Check input data.`);
    if (nanFound && logicError) errors.push(`- Engine encountered severe NaN / Logic calculation errors.${logicError}. Check cross - sectional parameters or unit rates.`);

    return errors;
  }

  // ── Roll Diagram pan / zoom / sync interaction ──────────────────────────
  (function setupRollInteractions() {
    const planWrap = document.getElementById("rollDiagramWrap");
    const sideWrap = document.getElementById("sideViewWrap");
    const syncBtn = document.getElementById("rollSyncBtn");
    const zoomIn = document.getElementById("rollZoomIn");
    const zoomOut = document.getElementById("rollZoomOut");
    const zoomReset = document.getElementById("rollZoomReset");
    const zoomFit = document.getElementById("rollZoomFit");
    const fullscreenBtn = document.getElementById("rollFullscreenBtn");
    const fullscreenCloseBtn = document.getElementById("rollFullscreenCloseBtn");
    const fullscreenModal = document.getElementById("rollFullscreenModal");
    const fullscreenViewport = document.getElementById("rollFullscreenViewport");
    const rollPageHeader = document.getElementById("rollPageHeader");
    const rollSplitShell = document.getElementById("rollSplitShell");

    let synced = false;
    let syncLock = false;
    let spacePanActive = false;
    let fullscreenOpen = false;
    let headerPlaceholder = null;
    let shellPlaceholder = null;

    function renderRollViews() {
      if (!state.calcRows.length) return;
      renderRollDiagram();
      renderSideView();
    }

    function setSyncButtonState() {
      if (!syncBtn) return;
      syncBtn.style.background = synced ? "rgba(34,211,238,0.18)" : "rgba(255,255,255,0.06)";
      syncBtn.style.borderColor = synced ? "rgba(34,211,238,0.6)" : "rgba(255,255,255,0.15)";
      syncBtn.style.color = synced ? "#67e8f9" : "rgba(255,255,255,0.7)";
      if (syncBtn.querySelector("span")) {
        syncBtn.querySelector("span").textContent = synced ? "Synced ✓" : "Sync Scrolling";
      }
    }

    // ── Sync button ─────────────────────────────────────────────────────────
    if (syncBtn) {
      syncBtn.addEventListener("click", () => {
        synced = !synced;
        setSyncButtonState();
        if (synced && planWrap && sideWrap) {
          window._sideScale = window._planScale || 1;
          if (state.calcRows.length) renderSideView();
          requestAnimationFrame(() => { sideWrap.scrollLeft = planWrap.scrollLeft; });
        }
      });
    }
    setSyncButtonState();

    function syncScrollLeft(from, to) {
      if (!synced || syncLock || !from || !to) return;
      syncLock = true;
      to.scrollLeft = from.scrollLeft;
      syncLock = false;
    }

    function getFitScaleForWrap(wrap, kind) {
      if (!wrap || !state.calcRows.length) return 1;
      const rows = state.calcRows;
      const minCh = rows[0].chainage;
      const maxCh = rows[rows.length - 1].chainage;
      const totalL = Math.max(maxCh - minCh, 1);
      const pads = kind === "plan" ? { left: 60, right: 40 } : { left: 72, right: 30 };
      const usableWidth = Math.max(wrap.clientWidth - pads.left - pads.right - 24, 240);
      return Math.max(0.05, Math.min(40, usableWidth / (totalL * 0.4)));
    }

    function fitRollViews() {
      if (!state.calcRows.length) return;
      const planFit = getFitScaleForWrap(planWrap, "plan");
      const sideFit = getFitScaleForWrap(sideWrap, "side");
      if (synced) {
        const sharedScale = Math.min(planFit, sideFit);
        window._planScale = sharedScale;
        window._sideScale = sharedScale;
      } else {
        window._planScale = planFit;
        window._sideScale = sideFit;
      }
      renderRollViews();
      requestAnimationFrame(() => {
        if (planWrap) {
          planWrap.scrollLeft = 0;
          planWrap.scrollTop = 0;
        }
        if (sideWrap) {
          sideWrap.scrollLeft = 0;
          sideWrap.scrollTop = 0;
        }
      });
    }

    // ── Drag-to-pan ─────────────────────────────────────────────────────────
    function addDragPan(wrap) {
      if (!wrap) return;
      let dragging = false, lx = 0, ly = 0;
      wrap.addEventListener("mousedown", (e) => {
        const shouldPan = e.button === 1 || e.button === 0 || (spacePanActive && e.button === 0);
        if (!shouldPan) return;
        dragging = true; lx = e.clientX; ly = e.clientY;
        wrap.classList.add("is-panning");
        e.preventDefault();
      });
      wrap.addEventListener("mousemove", (e) => {
        if (!dragging) return;
        const dx = e.clientX - lx, dy = e.clientY - ly;
        wrap.scrollLeft -= dx; wrap.scrollTop -= dy;
        lx = e.clientX; ly = e.clientY;
        if (synced) {
          const other = wrap === planWrap ? sideWrap : planWrap;
          syncScrollLeft(wrap, other);
        }
      });
      ["mouseup", "mouseleave"].forEach(ev => {
        wrap.addEventListener(ev, () => {
          dragging = false;
          wrap.classList.remove("is-panning");
        });
      });
      // Touch support
      let tlx = 0, tly = 0;
      wrap.addEventListener("touchstart", (e) => { tlx = e.touches[0].clientX; tly = e.touches[0].clientY; }, { passive: true });
      wrap.addEventListener("touchmove", (e) => {
        const dx = e.touches[0].clientX - tlx, dy = e.touches[0].clientY - tly;
        wrap.scrollLeft -= dx; wrap.scrollTop -= dy;
        tlx = e.touches[0].clientX; tly = e.touches[0].clientY;
        if (synced) { const other = wrap === planWrap ? sideWrap : planWrap; syncScrollLeft(wrap, other); }
        e.preventDefault();
      }, { passive: false });
      wrap.addEventListener("dblclick", () => {
        if (fullscreenOpen) fitRollViews();
      });
    }

    addDragPan(planWrap);
    addDragPan(sideWrap);

    // ── Scroll-based sync (for native scrollbar use) ─────────────────────────
    if (planWrap) planWrap.addEventListener("scroll", () => { if (!syncLock) syncScrollLeft(planWrap, sideWrap); });
    if (sideWrap) sideWrap.addEventListener("scroll", () => { if (!syncLock) syncScrollLeft(sideWrap, planWrap); });

    // ── Scroll-wheel zoom toward cursor (AutoCAD style) ───────────────────────
    function addWheelZoom(wrap) {
      if (!wrap) return;
      wrap.addEventListener("wheel", (e) => {
        e.preventDefault();
        if (!state.calcRows.length) return;
        const ZOOM_STEP = 1.12;
        const factor = e.deltaY < 0 ? ZOOM_STEP : 1 / ZOOM_STEP;   // scroll up = zoom in
        const isPlan = wrap === planWrap;
        const oldScale = isPlan ? (window._planScale || 1) : (window._sideScale || 1);
        const newScale = Math.max(0.05, oldScale * factor);
        if (newScale === oldScale) return;

        // Remember which chainage is under the cursor so we can restore it after re-render
        const rect = wrap.getBoundingClientRect();
        const cursorXpx = e.clientX - rect.left + wrap.scrollLeft;  // pixel in canvas coords
        const cursorYpx = e.clientY - rect.top + wrap.scrollTop;
        const ratio = newScale / oldScale;

        if (synced) {
          window._planScale = newScale;
          window._sideScale = newScale;
          renderRollViews();
        } else {
          if (isPlan) {
            window._planScale = newScale;
            renderRollDiagram();
          } else {
            window._sideScale = newScale;
            renderSideView();
          }
        }

        // Restore cursor focal point
        requestAnimationFrame(() => {
          wrap.scrollLeft = cursorXpx * ratio - (e.clientX - rect.left);
          wrap.scrollTop = cursorYpx * ratio - (e.clientY - rect.top);
          // Sync horizontal scroll to other panel if synced
          if (synced) {
            const other = wrap === planWrap ? sideWrap : planWrap;
            if (other) other.scrollLeft = wrap.scrollLeft;
          }
        });
      }, { passive: false });
    }

    addWheelZoom(planWrap);
    addWheelZoom(sideWrap);

    // ── +/−/Reset zoom buttons ───────────────────────────────────────────────
    function zoomButtons(factor, isReset, isFit) {
      if (!state.calcRows.length) return;
      if (isFit) {
        fitRollViews();
        return;
      }

      const oldPlan = window._planScale || 1, oldSide = window._sideScale || 1;
      const newPlan = isReset ? 1 : Math.max(0.05, oldPlan * factor);
      const newSide = isReset ? 1 : Math.max(0.05, oldSide * factor);

      const ratioPlan = newPlan / oldPlan;
      const ratioSide = newSide / oldSide;

      window._planScale = newPlan;
      window._sideScale = synced ? newPlan : newSide;

      // Centre points to zoom toward
      const pcx = planWrap ? planWrap.scrollLeft + planWrap.clientWidth / 2 : 0;
      const pcy = planWrap ? planWrap.scrollTop + planWrap.clientHeight / 2 : 0;
      const scx = sideWrap ? sideWrap.scrollLeft + sideWrap.clientWidth / 2 : 0;
      const scy = sideWrap ? sideWrap.scrollTop + sideWrap.clientHeight / 2 : 0;

      renderRollViews();

      requestAnimationFrame(() => {
        if (planWrap) {
          planWrap.scrollLeft = isReset ? 0 : pcx * ratioPlan - planWrap.clientWidth / 2;
          planWrap.scrollTop = isReset ? 0 : pcy * ratioPlan - planWrap.clientHeight / 2;
        }
        if (sideWrap) {
          if (synced && planWrap && !isReset) {
            sideWrap.scrollLeft = planWrap.scrollLeft;
          } else {
            sideWrap.scrollLeft = isReset ? 0 : scx * ratioSide - sideWrap.clientWidth / 2;
          }
          sideWrap.scrollTop = isReset ? 0 : scy * ratioSide - sideWrap.clientHeight / 2;
        }
      });
    }

    if (zoomIn) zoomIn.addEventListener("click", () => zoomButtons(1.4, false));
    if (zoomOut) zoomOut.addEventListener("click", () => zoomButtons(1 / 1.4, false));
    if (zoomReset) zoomReset.addEventListener("click", () => zoomButtons(1, true));
    if (zoomFit) zoomFit.addEventListener("click", () => zoomButtons(1, false, true));

    function restoreRollLayout() {
      if (!fullscreenOpen) return;
      if (document.fullscreenElement) {
        document.exitFullscreen().catch(() => {});
      }
      if (rollPageHeader && headerPlaceholder && headerPlaceholder.parentNode) {
        headerPlaceholder.parentNode.insertBefore(rollPageHeader, headerPlaceholder);
        headerPlaceholder.remove();
        headerPlaceholder = null;
      }
      if (rollSplitShell && shellPlaceholder && shellPlaceholder.parentNode) {
        shellPlaceholder.parentNode.insertBefore(rollSplitShell, shellPlaceholder);
        shellPlaceholder.remove();
        shellPlaceholder = null;
      }
      if (fullscreenCloseBtn) fullscreenCloseBtn.style.display = "none";
      if (fullscreenBtn) fullscreenBtn.style.display = "inline-flex";
      if (fullscreenModal && fullscreenModal.open) fullscreenModal.close();
      fullscreenOpen = false;
    }

    function openRollFullscreen() {
      if (fullscreenOpen || !fullscreenModal || !fullscreenViewport || !rollPageHeader || !rollSplitShell) return;
      headerPlaceholder = document.createComment("roll-page-header-placeholder");
      shellPlaceholder = document.createComment("roll-split-shell-placeholder");
      rollPageHeader.parentNode.insertBefore(headerPlaceholder, rollPageHeader);
      rollSplitShell.parentNode.insertBefore(shellPlaceholder, rollSplitShell);
      fullscreenViewport.appendChild(rollPageHeader);
      fullscreenViewport.appendChild(rollSplitShell);
      if (fullscreenBtn) fullscreenBtn.style.display = "none";
      if (fullscreenCloseBtn) fullscreenCloseBtn.style.display = "inline-flex";
      fullscreenModal.showModal();
      fullscreenOpen = true;
      requestAnimationFrame(() => {
        fitRollViews();
        const fullscreenTarget = fullscreenModal.querySelector(".roll-fullscreen-body");
        if (fullscreenTarget && fullscreenTarget.requestFullscreen) {
          fullscreenTarget.requestFullscreen().catch(() => {});
        }
      });
    }

    if (fullscreenBtn) fullscreenBtn.addEventListener("click", openRollFullscreen);
    if (fullscreenCloseBtn) fullscreenCloseBtn.addEventListener("click", restoreRollLayout);

    if (fullscreenModal) {
      fullscreenModal.addEventListener("cancel", (event) => {
        event.preventDefault();
        restoreRollLayout();
      });
      fullscreenModal.addEventListener("close", restoreRollLayout);
    }

    document.addEventListener("keydown", (event) => {
      if (event.code === "Space" && !event.repeat) {
        spacePanActive = true;
      }
      if (!fullscreenOpen) return;
      if (event.key === "0") {
        event.preventDefault();
        fitRollViews();
      } else if (event.key === "+" || event.key === "=") {
        event.preventDefault();
        zoomButtons(1.2, false);
      } else if (event.key === "-" || event.key === "_") {
        event.preventDefault();
        zoomButtons(1 / 1.2, false);
      } else if (event.key.toLowerCase() === "f") {
        event.preventDefault();
        fitRollViews();
      }
    });

    document.addEventListener("keyup", (event) => {
      if (event.code === "Space") {
        spacePanActive = false;
      }
    });
  })();

  // Verify Calculations button
  const verifyBtn = document.getElementById("verifyCalcBtn");
  if (verifyBtn) {
    verifyBtn.addEventListener("click", () => {
      verifyBtn.classList.add("verifying");
      verifyBtn.disabled = true;
      recalculate();

      setTimeout(() => {
        const errors = validateProjectData();
        verifyBtn.classList.remove("verifying");

        const modal = document.getElementById("verifyModal");
        const modalContent = document.getElementById("verifyModalContent");
        const modalIcon = document.getElementById("verifyModalIcon");
        const modalIconEl = document.getElementById("verifyModalIconEl");
        const modalTitle = document.getElementById("verifyModalTitle");

        if (errors.length === 0) {
          state.project.verified = true;
          updateWizardUI();
          applyProjectGate();
          updateEstimates();
          updateDashboard();
          verifyBtn.classList.add("verified");
          if (modalTitle) modalTitle.textContent = "Verification Successful";
          if (modalIcon) {
            modalIcon.style.background = "rgba(16, 185, 129, 0.15)";
            modalIcon.style.color = "#10b981";
            modalIcon.style.border = "1px solid rgba(16, 185, 129, 0.2)";
          }
          if (modalIconEl) modalIconEl.className = "ri-checkbox-circle-line";
          if (modalContent) {
            modalContent.innerHTML = `
               <div style="background: rgba(16, 185, 129, 0.05); border: 1px solid rgba(16, 185, 129, 0.2); padding: 16px; border-radius: 8px; margin-bottom: 12px; display: flex; align-items: flex-start; gap: 12px;">
                 <i class="ri-shield-check-fill" style="color: #10b981; font-size: 1.5rem;"></i>
                 <div>
                   <div style="color: #fff; font-weight: 600; margin-bottom: 4px;">Engine Verification Passed</div>
                   <div style="font-size: 0.85rem; color: rgba(255,255,255,0.7);">
                     All structural calculations, chainages, and cross-section parameters have been independently validated. You are clear to export.
                     <br><br>
                     <strong>Checked Points:</strong> ${state.calcRows.length} Cross-Sections
                     <br>
                     <strong>Timestamp:</strong> ${new Date().toLocaleTimeString()}
                   </div>
                 </div>
               </div>
             `;
          }

          if (modal) modal.showModal();

          setTimeout(() => {
            verifyBtn.classList.remove("verified");
            verifyBtn.disabled = false;
          }, 2000);

        } else {
          state.project.verified = false;
          updateWizardUI();
          applyProjectGate();
          updateEstimates();
          updateDashboard();
          verifyBtn.disabled = false;
          if (modalTitle) modalTitle.textContent = "Verification Failed";
          if (modalIcon) {
            modalIcon.style.background = "rgba(239, 68, 68, 0.15)";
            modalIcon.style.color = "#ef4444";
            modalIcon.style.border = "1px solid rgba(239, 68, 68, 0.2)";
          }
          if (modalIconEl) modalIconEl.className = "ri-error-warning-line";
          if (modalContent) {
            const errorList = errors.map(e => `<li style="margin-bottom: 8px;"><i class="ri-close-circle-fill" style="color: #ef4444; margin-right: 6px;"></i>${e}</li>`).join("");
            modalContent.innerHTML = `
               <div style="background: rgba(239, 68, 68, 0.05); border: 1px solid rgba(239, 68, 68, 0.2); padding: 16px; border-radius: 8px; margin-bottom: 12px;">
                 <div style="color: #ef4444; font-weight: 600; margin-bottom: 8px;">Critical Errors Found (${errors.length})</div>
                 <div style="font-size: 0.85rem; color: rgba(255,255,255,0.7); margin-bottom: 12px;">
                   The system has halted processing to prevent inaccurate data from being exported. Please resolve the following structural flaws:
                 </div>
                 <ul style="list-style: none; padding: 0; margin: 0; font-family: monospace; font-size: 13px; color: #f87171; background: rgba(0,0,0,0.3); padding: 12px; border-radius: 6px;">
                   ${errorList}
                 </ul>
               </div>
             `;
          }
          if (modal) modal.showModal();
        }
      }, 800);
    });
  }

  // --- Keyboard Shortcuts ---
  document.addEventListener("keydown", (e) => {
    // Save: Ctrl+S / Cmd+S
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
      e.preventDefault();
      if (els.saveProjectBtn) els.saveProjectBtn.click();
    }
    if (els.crossSectionModal?.open) {
      if (e.key === "ArrowLeft") {
        e.preventDefault();
        openCrossSectionByIndex((state.currentCrossIndex ?? 0) - 1);
        return;
      }
      if (e.key === "ArrowRight") {
        e.preventDefault();
        openCrossSectionByIndex((state.currentCrossIndex ?? -1) + 1);
        return;
      }
    }
    // Arrow Key Navigation for Chainages (Only active on Overview)
    const overviewPage = document.querySelector('.work-page[data-work-page="overview"]');
    if (overviewPage && overviewPage.classList.contains("active")) {
      if (e.key === "ArrowLeft") {
        document.getElementById("prevChainBtn")?.click();
      } else if (e.key === "ArrowRight") {
        document.getElementById("nextChainBtn")?.click();
      }
    }
  });

  if (els.resultTabs) {
    els.resultTabs.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-result-tab]");
      if (!btn) return;
      setResultTab(btn.dataset.resultTab);
    });
  }

  document.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-overview-action]");
    if (!btn) return;
    const action = btn.dataset.overviewAction;
    if (action === "create") els.createProjectBtn?.click();
    if (action === "open") els.importProjectBtn?.click();
    if (action === "import") els.importBtn?.click();
    if (action === "verify") document.getElementById("verifyCalcBtn")?.click();
    if (action === "bridge") setWorkPage("bridges");
    if (action === "station") setWorkPage("loops");
    if (action === "plans") setWorkPage("alignment-map");
    if (action === "map") {
      setWorkPage("alignment-map");
    }
    if (action === "export") els.openExportModalBtn?.click();
  });

  if (els.createProjectBtn && els.projectWizardModal) {
    els.createProjectBtn.addEventListener("click", () => {
      resetForNewProject();
      state.project.active = true;
      state.project.verified = false;
      updateWizardUI();
      applyProjectGate();
      els.projectWizardModal.showModal();
    });
  }

  if (els.importProjectBtn && els.projectImportInput) {
    els.importProjectBtn.addEventListener("click", async () => {
      if ("showOpenFilePicker" in window) {
        try {
          const [handle] = await window.showOpenFilePicker({
            multiple: false,
            types: [{
              description: "Earthwork Project File",
              accept: { "application/json": [".ew", ".EW"] },
            }],
          });
          if (!handle) return;
          const file = await handle.getFile();
          const txt = await file.text();
          const parsed = JSON.parse(txt);
          loadProjectFromPayload(parsed, { fileHandle: handle });
          return;
        } catch (err) {
          if (err?.name !== "AbortError") {
            console.error("Project open error:", err);
            alert(`Project import failed: ${err?.message || "Invalid .EW file."}`);
          }
          return;
        }
      }
      els.projectImportInput.click();
    });
  }

  if (els.closeProjectWizardBtn && els.projectWizardModal) {
    els.closeProjectWizardBtn.addEventListener("click", () => {
      els.projectWizardModal.close();
    });
  }

  if (els.projectNameInput) {
    els.projectNameInput.addEventListener("input", () => {
      state.project.name = String(els.projectNameInput.value || "").trim();
      updateWizardUI();
      applyProjectGate();
    });
  }

  if (els.wizardUploadButtons?.length) {
    els.wizardUploadButtons.forEach((btn) => {
      btn.addEventListener("click", () => {
        const kind = btn.dataset.wizardUpload;
        if (kind === "levels" && els.importInput) els.importInput.click();
        if (kind === "bridges" && els.bridgeImportInput) els.bridgeImportInput.click();
        if (kind === "curves" && els.curveImportInput) els.curveImportInput.click();
        if (kind === "loops" && els.loopImportInput) els.loopImportInput.click();
        if (kind === "kml" && els.kmlImportInput) els.kmlImportInput.click();
      });
    });
  }

  // --- Drag and Drop File Upload ---
  const dropZone = document.getElementById("wizardDropZone");
  const dropOverlay = document.getElementById("wizardDropOverlay");

  if (dropZone && dropOverlay) {
    dropZone.addEventListener("dragover", (e) => {
      e.preventDefault();
      dropZone.classList.add("drag-active");
    });

    dropZone.addEventListener("dragleave", (e) => {
      e.preventDefault();
      // Only remove if we really left the zone
      if (e.target === dropOverlay || e.target === dropZone) {
        dropZone.classList.remove("drag-active");
      }
    });

    dropZone.addEventListener("drop", async (e) => {
      e.preventDefault();
      dropZone.classList.remove("drag-active");

      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        const files = Array.from(e.dataTransfer.files).filter(f => f.name.match(/\.(xlsx|xls|csv|kml|kmz)$/i));

        // Show loading state on cursor 
        document.body.style.cursor = "wait";

        for (const file of files) {
          try {
            if (file.name.match(/\.(kml|kmz)$/i)) {
              if (els.kmlImportInput) {
                const dt = new DataTransfer();
                dt.items.add(file);
                els.kmlImportInput.files = dt.files;
                els.kmlImportInput.dispatchEvent(new Event("change"));
              }
              continue;
            }
            const data = await file.arrayBuffer();
            const wb = XLSX.read(data, { type: "array", cellDates: false });
            const headers = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 })[0] || [];
            const headerStr = headers.map(h => String(h).toLowerCase()).join(" ");

            if (headerStr.includes("radius") && headerStr.includes("length")) {
              await importCurveFile(file);
            } else if (headerStr.includes("bridge") || (headerStr.includes("start") && headerStr.includes("end"))) {
              await importBridgeFile(file);
            } else if (headerStr.includes("code") || headerStr.includes("platform") || wb.SheetNames[0].toLowerCase().includes("loop")) {
              await importLoopFile(file);
            } else if (headerStr.includes("chainage") || headerStr.includes("station") || headerStr.includes("level")) {
              // Fallback primarily for Earthwork Levels
              await importLevelsFile(file);
            } else {
              console.warn("Could not auto-route file by headers:", file.name, "- defaulting to Levels");
              // If completely ambiguous, guess levels because it's the most important
              await importLevelsFile(file);
            }
          } catch (err) {
            console.error("Auto-route error for", file.name, err);
            alert(`Could not process dropped file ${file.name}: ${err.message} `);
          }
        }

        document.body.style.cursor = "default";
      }
    });
  }

  if (els.wizardSaveBtn) {
    els.wizardSaveBtn.addEventListener("click", () => {
      saveCurrentProject();
    });
  }

  if (els.saveProjectBtn) {
    els.saveProjectBtn.addEventListener("click", () => {
      saveCurrentProject();
    });
  }

  if (els.wizardCalculateBtn) {
    els.wizardCalculateBtn.addEventListener("click", () => {
      if (!isProjectReadyForVerification()) {
        alert("Upload all files and enter project name to verify.");
        return;
      }

      // Start fluid loading animation
      els.wizardCalculateBtn.classList.add("is-loading");

      // Wait for wave animation to "fill"
      setTimeout(() => {
        // Switch to verified state
        els.wizardCalculateBtn.classList.add("is-verified");

        // Wait for checkmark to settle
        setTimeout(() => {
          // Execute actual logic
          state.project.verified = true;
          recalculate();
          updateWizardUI();
          applyProjectGate();
          els.projectWizardModal?.close();

          // Reset button for next time
          setTimeout(() => {
            els.wizardCalculateBtn.classList.remove("is-loading", "is-verified");
          }, 300);

        }, 600);
      }, 1200);
    });
  }

  if (els.resetProjectBtn) {
    els.resetProjectBtn.addEventListener("click", () => {
      if (confirm("Are you sure you want to reset the project? All unsaved progress will be lost.")) {
        resetForNewProject();
        alert("Workspace reset to default. Create a new project to continue.");
      }
    });
  }

  // --- Project Snapshots ---
  state.snapshots = state.snapshots || [];

  if (els.snapshotsBtn && els.snapshotsModal) {
    els.snapshotsBtn.addEventListener("click", () => {
      if (!state.project.active) {
        alert("Create or open a project first.");
        return;
      }
      renderSnapshotsList();
      els.snapshotsModal.showModal();
    });

    if (els.closeSnapshotsBtn) {
      els.closeSnapshotsBtn.addEventListener("click", () => els.snapshotsModal.close());
    }
    if (els.closeSnapshotCompareBtn) {
      els.closeSnapshotCompareBtn.addEventListener("click", () => els.snapshotCompareModal?.close());
    }

    if (els.takeSnapshotBtn) {
      els.takeSnapshotBtn.addEventListener("click", () => {
        const name = els.snapshotNameInput.value.trim() || `Snapshot ${new Date().toLocaleTimeString()} `;
        const snapshot = {
          id: Date.now().toString(),
          name,
          timestamp: new Date().toLocaleString(),
          rawRows: JSON.parse(JSON.stringify(state.rawRows || [])),
          curves: JSON.parse(JSON.stringify(state.curves || [])),
          bridges: JSON.parse(JSON.stringify(state.bridges || [])),
          loops: JSON.parse(JSON.stringify(state.loops || [])),
          levels: JSON.parse(JSON.stringify(state.levels || [])),
          settings: JSON.parse(JSON.stringify(state.settings || {})),
        };

        state.snapshots.push(snapshot);
        els.snapshotNameInput.value = "";
        renderSnapshotsList();
      });
    }
  }

  // --- Real-time Cost Estimation ---
  if (els.recalcEstimatesBtn) {
    els.recalcEstimatesBtn.addEventListener("click", updateEstimates);
  }
  [
    els.rateClearing, els.rateBenching, els.rateFilling, els.rateBlanketing,
    els.rateCutSoil, els.rateCutSoft, els.rateCutHardBlast, els.rateCutHardChisel,
    els.rateExtraLead, els.rateTurfing, els.pctSoil, els.pctSoftRock, els.pctHardBlast,
    els.pctHardChisel, els.pctReusableSpoil, els.leadKm
  ].forEach(input => {
    if (input) {
      input.addEventListener("input", updateEstimates);
      input.addEventListener("change", updateEstimates);
    }
  });

  // --- Expandable Graphs Logic ---
  let expandedChartInstance = null;

  function openExpandedGraph(graphType) {
    if (!els.graphModal) return;
    if (typeof Chart === "undefined") {
      alert("Chart library failed to load. Please check your internet connection.");
      return;
    }

    // Destroy previous expanded instance if it exists
    if (expandedChartInstance) {
      expandedChartInstance.destroy();
      expandedChartInstance = null;
    }

    const ctx = els.expandedGraphCanvas?.getContext("2d");
    if (!ctx) return;

    let sourceChart = null;
    let title = "";

    if (graphType === "lSection") {
      sourceChart = state.charts.lSection;
      title = "Expanded L-Section Profile";
    } else if (graphType === "volume") {
      sourceChart = state.charts.volume;
      title = "Expanded Volume Distribution";
    }

    if (!sourceChart || !sourceChart.data || sourceChart.data.datasets.length === 0) {
      alert("No data available to expand. Please generate a project first.");
      return;
    }

    els.graphModalTitle.textContent = title;

    // Clone data and config roughly (deep copy is safer for chart.js)
    const clonedData = JSON.parse(JSON.stringify(sourceChart.data));
    const clonedOptions = JSON.parse(JSON.stringify(sourceChart.options));

    // Adjust options for fullscreen view (disable relative sizing if needed, or adjust tooltips)
    clonedOptions.maintainAspectRatio = false;
    clonedOptions.plugins.title = { display: false };

    expandedChartInstance = new Chart(ctx, {
      type: sourceChart.config.type,
      data: clonedData,
      options: clonedOptions
    });

    els.graphModal.showModal();
  }

  document.addEventListener("click", (e) => {
    const btn = e.target.closest(".expand-graph-btn");
    if (btn) {
      openExpandedGraph(btn.dataset.graph);
    }
  });

  if (els.closeGraphModalBtn) {
    els.closeGraphModalBtn.addEventListener("click", () => {
      els.graphModal.close();
      if (expandedChartInstance) {
        expandedChartInstance.destroy();
        expandedChartInstance = null;
      }
    });
  }

  function renderSnapshotsList() {
    if (!els.snapshotList) return;
    if (state.snapshots.length === 0) {
      els.snapshotList.innerHTML = '<p class="muted">No snapshots taken yet.</p>';
      return;
    }

    els.snapshotList.innerHTML = state.snapshots.map(snap => `
    <div class="glass" style="padding: 12px; display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
        <div>
          <strong style="color: var(--text);">${snap.name}</strong><br/>
          <small class="muted">${snap.timestamp}</small>
        </div>
        <div style="display:flex; gap:8px;">
          <button class="btn btn-secondary" onclick="compareSnapshot('${snap.id}')">Compare</button>
          <button class="btn btn-secondary" onclick="restoreSnapshot('${snap.id}')">Restore</button>
        </div>
      </div>
    `).reverse().join("");
  }

  window.restoreSnapshot = (id) => {
    const snap = state.snapshots.find(s => s.id === id);
    if (!snap) return;
    if (!confirm(`Restore snapshot "${snap.name}" ? Current unsaved changes will be lost.`)) return;

    state.rawRows = JSON.parse(JSON.stringify(snap.rawRows));
    state.curves = JSON.parse(JSON.stringify(snap.curves));
    state.bridges = JSON.parse(JSON.stringify(snap.bridges));
    state.loops = JSON.parse(JSON.stringify(snap.loops));
    state.levels = JSON.parse(JSON.stringify(snap.levels));
    state.settings = JSON.parse(JSON.stringify(snap.settings));

    recalculate();
    els.snapshotsModal?.close();
  };

  const getSnapshotStats = (snap) => ({
    levels: snap?.rawRows?.length || 0,
    bridges: snap?.bridges?.length || 0,
    curves: snap?.curves?.length || 0,
    loops: snap?.loops?.length || 0,
    settings: snap?.settings || {},
  });

  const renderSnapshotCompare = (snap) => {
    if (!els.snapshotCompareBody) return;
    const current = {
      levels: state.rawRows.length,
      bridges: state.bridgeRows.length,
      curves: state.curveRows.length,
      loops: state.loopPlatformRows.length,
      settings: state.settings || {},
    };
    const prev = getSnapshotStats(snap);
    const cards = [
      ["Levels Rows", prev.levels, current.levels],
      ["Bridge Rows", prev.bridges, current.bridges],
      ["Curve Rows", prev.curves, current.curves],
      ["Loop Rows", prev.loops, current.loops],
      ["Formation Width", r3(prev.settings.formationWidthFill || 0), r3(current.settings.formationWidthFill || 0)],
      ["Slope Factor", r3(prev.settings.sideSlopeFactor || 0), r3(current.settings.sideSlopeFactor || 0)],
    ];
    els.snapshotCompareBody.innerHTML = cards.map(([label, oldVal, newVal]) => `
      <div class="glass" style="padding:12px;">
        <div style="font-weight:700; color:#e2e8f0;">${label}</div>
        <div class="muted" style="margin-top:4px;">Snapshot: ${oldVal}</div>
        <div class="muted">Current: ${newVal}</div>
      </div>
    `).join("");
  };

  window.compareSnapshot = (id) => {
    const snap = state.snapshots.find(s => s.id === id);
    if (!snap || !els.snapshotCompareModal) return;
    renderSnapshotCompare(snap);
    els.snapshotCompareModal.showModal();
  };

  const importLevelsFile = async (file) => {
    const ext = file.name.toLowerCase();
    if (!(ext.endsWith(".csv") || ext.endsWith(".xlsx") || ext.endsWith(".xls"))) {
      throw new Error("Supported files are .csv, .xlsx, .xls");
    }
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: "array", cellDates: false });

    const wsLevels = wb.Sheets[wb.SheetNames[0]];
    let aoaLevels = XLSX.utils.sheet_to_json(wsLevels, { header: 1, defval: "", raw: false, blankrows: false });
    if (!aoaLevels.length) throw new Error("File has no data rows.");
    aoaLevels = applyMappingTemplateToAoa(aoaLevels, "levels");

    // --- LOCAL PARSING FIRST (no AI required) ---
    const interval = inferImportInterval();
    const startChInput = inferImportStartChainage();
    const localResult = parseImportedRows(aoaLevels, startChInput, interval);
    let parsed = localResult.rows || [];

    // If local parsing failed, try AI as fallback
    if (!parsed.length) {
      // Strict validation
      const headerStr = aoaLevels.slice(0, 5).map(row => row.join(" ").toLowerCase()).join(" ");
      if (!(headerStr.includes("chainage") || headerStr.includes("station") || headerStr.includes("level"))) {
        throw new Error("Invalid File: The selected file does not appear to be an Earthwork Levels file. Check your column headers and ensure you are uploading to the correct field.");
      }

      // Try AI mapper as fallback
      let mapResult;
      try {
        const previewAoa = aoaLevels.slice(0, 50);
        const previewCsv = previewAoa.map(r => r.join(',')).join('\n');
        showAILoading("Mapping topographical Levels via AI...");
        const response = await fetch("/api/extract-data", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ text: previewCsv, dataType: 'levels' })
        });
        if (!response.ok) throw new Error(`Server Error: ${response.status} `);
        const resData = await response.json();
        mapResult = resData.data;
      } catch (e) {
        hideAILoading();
        throw new Error("Could not parse levels locally (no matching headers found), and AI fallback also failed: " + e.message);
      }
      hideAILoading();

      if (!mapResult || mapResult.chainageIndex < 0 || mapResult.groundLevelIndex < 0 || mapResult.proposedLevelIndex < 0) {
        throw new Error("Could not identify Chainage, Ground Level, and Proposed Level columns. Ensure your file has clear headers.");
      }

      let baseOffset = 0;
      const startIdx = Math.max(0, mapResult.dataStartRowIndex || 1);
      for (let i = startIdx; i < aoaLevels.length; i++) {
        const row = aoaLevels[i];
        if (!row || !row.length) continue;
        let chainageNum = parseChainage(String(row[mapResult.chainageIndex] || "").trim());
        let glNum = parseLooseNumber(row[mapResult.groundLevelIndex], NaN);
        let plNum = parseLooseNumber(row[mapResult.proposedLevelIndex], NaN);
        if (!Number.isFinite(glNum) || !Number.isFinite(plNum)) continue;
        if (!Number.isFinite(chainageNum)) {
          if (parsed.length === 0 && Number.isFinite(startChInput)) { baseOffset = startChInput; chainageNum = baseOffset; }
          else if (parsed.length > 0) { baseOffset += interval; chainageNum = baseOffset; }
          else continue;
        } else { baseOffset = chainageNum; }
        parsed.push({
          chainage: chainageNum,
          station: mapResult.stationIndex >= 0 ? String(row[mapResult.stationIndex] || "").trim() : "",
          structureNo: mapResult.structureNoIndex >= 0 ? String(row[mapResult.structureNoIndex] || "").trim() : "",
          groundLevel: glNum,
          proposedLevel: plNum
        });
      }
    }

    if (!parsed.length) {
      throw new Error("No valid rows found. Check your data file.");
    }
    state.rawRows = parsed;

    // Process additional sheets — LOCAL parsing first, AI fallback
    if (wb.SheetNames.length > 1) {
      for (let si = 1; si < wb.SheetNames.length; si++) {
        const sheetName = wb.SheetNames[si].toLowerCase();
        const ws = wb.Sheets[wb.SheetNames[si]];
        const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false, blankrows: false });
        if (!aoa.length) continue;

        if (sheetName.includes("bridge")) {
          const mappedAoa = applyMappingTemplateToAoa(aoa, "bridges");
          const localBridges = parseBridgeRowsFromAoa(mappedAoa, "Minor");
          if (localBridges.rows && localBridges.rows.length) {
            state.bridgeRows = localBridges.rows;
            state.project.uploads.bridges = true;
          } else {
            try {
              const csvContent = aoa.map(r => r.join(',')).join('\n');
              const res = await fetch("/api/extract-data", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ text: csvContent, dataType: 'bridges' }) }).then(r => r.json()).catch(() => null);
              if (res && res.success && res.data && res.data.length) { state.bridgeRows = res.data; state.project.uploads.bridges = true; }
            } catch (_) { /* AI fallback failed, skip */ }
          }
        }
        else if (sheetName.includes("curve")) {
          const mappedAoa = applyMappingTemplateToAoa(aoa, "curves");
          const localCurves = parseCurveRowsFromAoa(mappedAoa);
          if (localCurves.rows && localCurves.rows.length) {
            state.curveRows = localCurves.rows;
            state.project.uploads.curves = true;
          } else {
            try {
              const csvContent = aoa.map(r => r.join(',')).join('\n');
              const res = await fetch("/api/extract-data", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ text: csvContent, dataType: 'curves' }) }).then(r => r.json()).catch(() => null);
              if (res && res.success && res.data && res.data.length) { state.curveRows = res.data; state.project.uploads.curves = true; }
            } catch (_) { /* AI fallback failed, skip */ }
          }
        }
        else if (sheetName.includes("loop") || sheetName.includes("station") || sheetName.includes("platform")) {
          const mappedAoa = applyMappingTemplateToAoa(aoa, "loops");
          const localLoops = parseLoopPlatformRowsFromAoa(mappedAoa);
          if (localLoops.rows && localLoops.rows.length) {
            state.loopPlatformRows = localLoops.rows;
            state.project.uploads.loops = true;
          } else {
            try {
              const csvContent = aoa.map(r => r.join(',')).join('\n');
              const res = await fetch("/api/extract-data", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ text: csvContent, dataType: 'loops' }) }).then(r => r.json()).catch(() => null);
              if (res && res.success && res.data && res.data.length) { state.loopPlatformRows = res.data; state.project.uploads.loops = true; }
            } catch (_) { /* AI fallback failed, skip */ }
          }
        }
      }
    }

    const ch = parsed.map((r) => r.chainage).sort((a, b) => a - b);
    els.projectMeta.textContent = "Earthwork Calculations";
    state.project.uploads.levels = true;
    state.project.verified = false;
    updateWizardUI();
    applyProjectGate();

    renderBridgeInputs();
    renderCurveInputs();
    renderLoopInputs();
    recalculate();
  };

  // --- Universal AI File Processor Helpers ---
  const showAILoading = (text) => {
    const overlay = document.getElementById("aiLoadingOverlay");
    const textEl = document.getElementById("aiLoadingText");
    if (overlay && textEl) {
      textEl.textContent = text || "Extracting and formatting data structures.";
      overlay.classList.remove("hidden");
    }
    document.body.style.cursor = "wait";
  };

  const hideAILoading = () => {
    const overlay = document.getElementById("aiLoadingOverlay");
    if (overlay) overlay.classList.add("hidden");
    document.body.style.cursor = "default";
  };

  // --- Universal AI File Processor ---
  const processFileWithAI = async (file, dataType) => {
    // 1. Read file as an array of arrays
    const aoa = await readSheetAoaFromFile(file);
    if (!aoa || !aoa.length) throw new Error("File appears empty or unreadable.");

    // Strict validation: Verify the user uploaded the 'correct' file into the 'correct' field
    // Get headers (first 5 rows just in case they are shifted)
    const headerStr = aoa.slice(0, 5).map(row => row.join(" ").toLowerCase()).join(" ");

    if (dataType === "curves" && !(headerStr.includes("radius") && headerStr.includes("length") && (headerStr.includes("curve") || headerStr.includes("chainage")))) {
      throw new Error("Invalid File: The selected file does not appear to be a Curve list. Check your column headers and ensure you are uploading to the correct field.");
    }
    if (dataType === "bridges" && !(headerStr.includes("bridge") || (headerStr.includes("start") && headerStr.includes("end")) || (headerStr.includes("span") && headerStr.includes("length")))) {
      throw new Error("Invalid File: The selected file does not appear to be a Bridge list. Check your column headers and ensure you are uploading to the correct field.");
    }
    if (dataType === "loops" && !(headerStr.includes("code") || headerStr.includes("station") || headerStr.includes("platform") || headerStr.includes("loop") || headerStr.includes("width"))) {
      throw new Error("Invalid File: The selected file does not appear to be a Loops & Platforms list. Check your column headers and ensure you are uploading to the correct field.");
    }

    // 2. Convert to raw CSV string (to save tokens)
    const csvContent = aoa.map(row => row.join(',')).join('\n');

    // 3. Optional: display loading state inside the app (e.g., changing cursor)
    showAILoading(`Extracting ${dataType} details via AI...`);

    try {
      // 4. Send to Vercel API
      const response = await fetch("/api/extract-data", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ text: csvContent, dataType })
      });

      if (!response.ok) {
        const errData = await response.json().catch(() => ({}));
        throw new Error(errData.error || `Server Error: ${response.status} `);
      }

      const resData = await response.json();
      return resData;

    } finally {
      hideAILoading();
    }
  };

  const importBridgeFile = async (file) => {
    try {
      const sheets = await readAllSheetsFromFile(file);
      let importedRows = [];
      let combinedCsv = "";

      for (const sheet of sheets) {
        // Collect for potential AI fallback
        combinedCsv += `-- - SHEET: ${sheet.name} ---\n`;
        combinedCsv += sheet.aoa.map(row => row.map(c => `"${String(c || "").replace(/"/g, '""').replace(/\n/g, " ")}"`).join(", ")).join("\n") + "\n\n";

        // Try local parsing first using the existing bridge parser
        const mappedAoa = applyMappingTemplateToAoa(sheet.aoa, "bridges");
        const localResult = parseBridgeRowsFromAoa(mappedAoa, sheet.name);
        if (localResult.rows && localResult.rows.length > 0) {
          importedRows = importedRows.concat(localResult.rows);
        }
      }

      // If local parsing didn't find anything, try the AI (likely a complex/messy layout)
      if (importedRows.length === 0 && combinedCsv.trim().length > 10) {
        showAILoading(`No standard structure headers found. Trying AI extraction...`);
        try {
          const response = await fetch("/api/extract-data", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ text: combinedCsv, dataType: "bridges" })
          });

          if (response.ok) {
            const resData = await response.json();
            if (resData && resData.success && resData.data && resData.data.length > 0) {
              importedRows = resData.data.map((r, i) => normalizeBridgeEntry(r, i)).filter(Boolean);
            }
          }
        } catch (aiErr) {
          console.warn("AI Fallback failed:", aiErr);
        } finally {
          hideAILoading();
        }
      }

      if (importedRows.length > 0) {
        state.bridgeRows = importedRows;
        state.bridgeRows.sort((a, b) => (safeNum(a.startChainage) - safeNum(b.startChainage)));
        state.project.uploads.bridges = true;
        state.project.verified = false;
        updateWizardUI();
        applyProjectGate();
        renderBridgeInputs();
        recalculate();
        alert(`Successfully imported ${importedRows.length} bridge(s)/tunnel(s)!`);
      } else {
        alert("Could not identify any bridge structures in that file. Check your headers and sheet content.");
      }
    } catch (err) {
      console.error(err);
      alert(`Bridge import failed: ${err.message}`);
    }
  };

  const importCurveFile = async (file) => {
    try {
      // Try local parsing first
      let aoa = await readSheetAoaFromFile(file);
      aoa = applyMappingTemplateToAoa(aoa, "curves");
      const localResult = parseCurveRowsFromAoa(aoa);
      let importedRows = localResult.rows || [];

      // If local parsing yielded nothing, try AI fallback
      if (!importedRows.length) {
        try {
          const result = await processFileWithAI(file, "curves");
          if (result && result.success && result.data && result.data.length > 0) {
            importedRows = result.data;
          }
        } catch (aiErr) {
          console.warn("AI curve fallback failed:", aiErr.message);
        }
      }

      if (importedRows.length > 0) {
        state.curveRows = importedRows;
        state.project.uploads.curves = true;
        state.project.verified = false;
        updateWizardUI();
        applyProjectGate();
        recalculate();
        alert(`Successfully imported ${importedRows.length} curve(s)!`);
      } else {
        alert("Could not identify any curve data in that file. Check your headers.");
      }
    } catch (err) {
      console.error(err);
      alert(`Curve import failed: ${err.message}`);
    }
  };

  const importLoopFile = async (file) => {
    try {
      // Try local parsing first
      let aoa = await readSheetAoaFromFile(file);
      aoa = applyMappingTemplateToAoa(aoa, "loops");
      const localResult = parseLoopPlatformRowsFromAoa(aoa);
      let importedRows = localResult.rows || [];

      // If local parsing yielded nothing, try AI fallback
      if (!importedRows.length) {
        try {
          const result = await processFileWithAI(file, "loops");
          if (result && result.success && result.data && result.data.length > 0) {
            importedRows = result.data;
          }
        } catch (aiErr) {
          console.warn("AI loop fallback failed:", aiErr.message);
        }
      }

      if (importedRows.length > 0) {
        state.loopPlatformRows = importedRows;
        state.project.uploads.loops = true;
        state.project.verified = false;
        updateWizardUI();
        applyProjectGate();
        recalculate();
        const importedStations = [...new Set(importedRows.map((r) => String(r.station || "").trim()).filter(Boolean))];
        alert(`Successfully imported ${importedStations.length} station(s) across ${importedRows.length} station/loop row(s)!`);
      } else {
        alert("Could not identify any Loop or Station data in that file. Check your headers.");
      }
    } catch (err) {
      console.error(err);
      alert(`Loop/platform import failed: ${err.message}`);
    }
  };

  if (els.importBtn && els.importOptionsModal) {
    els.importBtn.addEventListener("click", () => els.importOptionsModal.showModal());
  }

  if (els.closeImportOptionsBtn && els.importOptionsModal) {
    els.closeImportOptionsBtn.addEventListener("click", () => els.importOptionsModal.close());
  }

  // --- AI Bridge Extraction Logic ---
  const aiBridgeModal = document.getElementById("aiBridgeModal");
  const bridgeAiBtn = document.getElementById("bridgeAiBtn");
  const closeAiModalBtn = document.getElementById("closeAiModalBtn");
  const startAiExtractBtn = document.getElementById("startAiExtractBtn");
  const aiTextInput = document.getElementById("aiTextInput");

  if (bridgeAiBtn && aiBridgeModal) {
    bridgeAiBtn.addEventListener("click", () => {
      aiTextInput.value = "";
      aiBridgeModal.classList.remove("hidden");
    });
  }

  if (closeAiModalBtn && aiBridgeModal) {
    closeAiModalBtn.addEventListener("click", () => {
      aiBridgeModal.classList.add("hidden");
    });
  }

  if (startAiExtractBtn && aiTextInput && aiBridgeModal) {
    startAiExtractBtn.addEventListener("click", async () => {
      const text = aiTextInput.value.trim();
      if (!text) return alert("Please paste some text first!");

      const originalBtnText = startAiExtractBtn.innerHTML;
      startAiExtractBtn.innerHTML = '<span class="spinner"></span> Extracting with AI...';
      startAiExtractBtn.disabled = true;

      try {
        const response = await fetch("/api/extract-data", {
          method: "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify({ text, dataType: "bridges" })
        });

        if (!response.ok) {
          const errData = await response.json().catch(() => ({}));
          throw new Error(errData.error || `Server Error: ${response.status}`);
        }

        const resData = await response.json();

        if (resData && resData.success && resData.data && resData.data.length > 0) {
          // Normalize the AI extracted bridges to add shouldDeduct flags and compute lengths
          const newRows = resData.data.map((r, i) => normalizeBridgeEntry(r, i)).filter(Boolean);
          state.bridgeRows = newRows;
          state.project.uploads.bridges = true;
          state.project.verified = false;

          updateWizardUI();
          applyProjectGate();
          renderBridgeInputs();
          recalculate();

          aiBridgeModal.classList.add("hidden");
          alert(`Successfully extracted and added ${resData.data.length} bridges using AI!`);
        } else {
          alert("The AI could not find any clear bridge data in that text.");
        }
      } catch (err) {
        console.error("AI Extraction Error:", err);
        alert(`Failed to extract bridges.\n\nError: ${err.message}`);
      } finally {
        startAiExtractBtn.innerHTML = originalBtnText;
        startAiExtractBtn.disabled = false;
      }
    });
  }
  // ----------------------------------

  if (els.importOptionsModal) {
    els.importOptionsModal.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-import-kind]");
      if (!btn) return;
      const kind = btn.dataset.importKind;
      if (kind === "levels" && els.importInput) {
        els.importOptionsModal.close();
        els.importInput.click();
      } else if (kind === "bridges" && els.bridgeImportInput) {
        els.importOptionsModal.close();
        els.bridgeImportInput.click();
      } else if (kind === "curves" && els.curveImportInput) {
        els.importOptionsModal.close();
        els.curveImportInput.click();
      } else if (kind === "loops" && els.loopImportInput) {
        els.importOptionsModal.close();
        els.loopImportInput.click();
      } else if (kind === "kml" && els.kmlImportInput) {
        els.importOptionsModal.close();
        els.kmlImportInput.click();
      }
    });
  }

  if (els.bridgeAddBtn) {
    els.bridgeAddBtn.addEventListener("click", () => {
      syncBridgeStateFromTable(); // Sync current inputs first
      const base = inferImportStartChainage();
      const nextNo = state.bridgeRows.length + 1;
      state.bridgeRows.push({
        bridgeNo: `BR-${nextNo}`,
        bridgeCategory: "Minor",
        bridgeType: "Box",
        bridgeSize: "6.1m",
        bridgeSpans: "1",
        clearSpan: "",
        deductRule: "Auto",
        autoDeduct: true,
        startChainage: r3(base),
        endChainage: r3(base + 20),
        length: r3(20),
      });
      state.project.uploads.bridges = true;
      state.project.verified = false;
      updateWizardUI();
      applyProjectGate();
      renderBridgeInputs();
    });
  }

  if (els.bridgeApplyBtn) {
    els.bridgeApplyBtn.addEventListener("click", () => {
      syncBridgeStateFromTable();
      state.project.uploads.bridges = true;
      state.project.verified = false;
      updateWizardUI();
      applyProjectGate();
      renderBridgeInputs();
      recalculate();
    });
  }

  if (els.bridgeTableBody) {
    els.bridgeTableBody.addEventListener("click", (e) => {
      const delBtn = e.target.closest("[data-bridge-del]");
      if (!delBtn) return;
      const i = Number(delBtn.dataset.bridgeDel);
      if (!Number.isFinite(i)) return;
      syncBridgeStateFromTable(); // Capture all current inputs first
      state.bridgeRows = state.bridgeRows.filter((_, idx) => idx !== i);
      state.project.uploads.bridges = state.bridgeRows.length > 0;
      state.project.verified = false;
      updateWizardUI();
      applyProjectGate();
      renderBridgeInputs();
      recalculate();
    });
  }

  if (els.curveAddBtn) {
    els.curveAddBtn.addEventListener("click", () => {
      state.curveRows.push({
        curve: `Curve-${state.curveRows.length + 1}`,
        chainage: null,
        radius: null,
        length: null,
      });
      state.project.uploads.curves = true;
      state.project.verified = false;
      updateWizardUI();
      applyProjectGate();
      renderCurveInputs();
    });
  }

  if (els.curveApplyBtn) {
    els.curveApplyBtn.addEventListener("click", () => {
      syncCurveStateFromTable();
      state.project.uploads.curves = state.curveRows.length > 0;
      state.project.verified = false;
      updateWizardUI();
      applyProjectGate();
      renderCurveInputs();
      recalculate();
    });
  }

  if (els.curveTableBody) {
    els.curveTableBody.addEventListener("click", (e) => {
      const delBtn = e.target.closest("[data-curve-del]");
      if (!delBtn) return;
      const i = Number(delBtn.dataset.curveDel);
      if (!Number.isFinite(i)) return;
      state.curveRows = state.curveRows.filter((_, idx) => idx !== i);
      state.project.uploads.curves = state.curveRows.length > 0;
      state.project.verified = false;
      updateWizardUI();
      applyProjectGate();
      renderCurveInputs();
      recalculate();
    });
  }

  if (els.loopAddBtn) {
    els.loopAddBtn.addEventListener("click", () => {
      state.loopPlatformRows.push(createEmptyLoopRow(state.loopPlatformRows.length));
      state.project.uploads.loops = true;
      state.project.verified = false;
      updateWizardUI();
      applyProjectGate();
      renderLoopInputs();
    });
  }

  if (els.loopApplyBtn) {
    els.loopApplyBtn.addEventListener("click", () => {
      syncLoopStateFromTable();
      state.project.uploads.loops = state.loopPlatformRows.length > 0;
      state.project.verified = false;
      updateWizardUI();
      applyProjectGate();
      renderLoopInputs();
      recalculate();
    });
  }

  if (els.loopTableBody) {
    els.loopTableBody.addEventListener("click", (e) => {
      const insertBtn = e.target.closest("[data-loop-insert]");
      if (insertBtn) {
        const insertIndex = Number(insertBtn.dataset.loopInsert);
        if (!Number.isFinite(insertIndex)) return;
        syncLoopStateFromTable();
        const aboveRow = state.loopPlatformRows[Math.max(0, insertIndex - 1)];
        state.loopPlatformRows.splice(
          insertIndex,
          0,
          createEmptyLoopRow(insertIndex, {
            station: String(aboveRow?.station || `LP-${insertIndex + 1}`),
          }),
        );
        state.project.uploads.loops = state.loopPlatformRows.length > 0;
        state.project.verified = false;
        updateWizardUI();
        applyProjectGate();
        renderLoopInputs();
        return;
      }
      const delBtn = e.target.closest("[data-loop-del]");
      if (delBtn) {
        const i = Number(delBtn.dataset.loopDel);
        if (!Number.isFinite(i)) return;
        state.loopPlatformRows = state.loopPlatformRows.filter((_, idx) => idx !== i);
        state.project.uploads.loops = state.loopPlatformRows.length > 0;
        state.project.verified = false;
        updateWizardUI();
        applyProjectGate();
        renderLoopInputs();
        recalculate();
        return;
      }

      const upBtn = e.target.closest("[data-loop-up]");
      if (upBtn) {
        const i = Number(upBtn.dataset.loopUp);
        if (Number.isFinite(i) && i > 0) {
          syncLoopStateFromTable();
          const targetRow = state.loopPlatformRows[i];
          const prevRow = state.loopPlatformRows[i - 1];
          state.loopPlatformRows.splice(i - 1, 2, targetRow, prevRow);
          state.project.verified = false;
          updateWizardUI();
          applyProjectGate();
          renderLoopInputs();
          recalculate();
        }
        return;
      }

      const downBtn = e.target.closest("[data-loop-down]");
      if (downBtn) {
        const i = Number(downBtn.dataset.loopDown);
        if (Number.isFinite(i) && i < state.loopPlatformRows.length - 1) {
          syncLoopStateFromTable();
          const targetRow = state.loopPlatformRows[i];
          const nextRow = state.loopPlatformRows[i + 1];
          state.loopPlatformRows.splice(i, 2, nextRow, targetRow);
          state.project.verified = false;
          updateWizardUI();
          applyProjectGate();
          renderLoopInputs();
          recalculate();
        }
        return;
      }
    });

    let draggedLoopIndex = null;
    els.loopTableBody.addEventListener("dragstart", (e) => {
      const row = e.target.closest("tr[data-drag-row]");
      if (!row) return;
      draggedLoopIndex = Number(row.dataset.dragRow);
      row.style.opacity = "0.4";
      e.dataTransfer.effectAllowed = "move";
    });

    els.loopTableBody.addEventListener("dragend", (e) => {
      const row = e.target.closest("tr[data-drag-row]");
      if (row) row.style.opacity = "1";
      draggedLoopIndex = null;
    });

    els.loopTableBody.addEventListener("dragover", (e) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      const row = e.target.closest("tr[data-drag-row]");
      if (row && Number(row.dataset.dragRow) !== draggedLoopIndex) {
        row.style.borderTop = "2px dashed var(--primary)";
      }
    });

    els.loopTableBody.addEventListener("dragleave", (e) => {
      const row = e.target.closest("tr[data-drag-row]");
      if (row) row.style.borderTop = "";
    });

    els.loopTableBody.addEventListener("drop", (e) => {
      e.preventDefault();
      const row = e.target.closest("tr[data-drag-row]");
      if (row) row.style.borderTop = "";
      
      if (draggedLoopIndex === null || !row) return;
      
      const targetIndex = Number(row.dataset.dragRow);
      if (draggedLoopIndex === targetIndex || !Number.isFinite(targetIndex)) return;

      syncLoopStateFromTable();
      const itemToMove = state.loopPlatformRows.splice(draggedLoopIndex, 1)[0];
      state.loopPlatformRows.splice(targetIndex, 0, itemToMove);

      state.project.verified = false;
      updateWizardUI();
      applyProjectGate();
      renderLoopInputs();
      recalculate();

      // subtle flash animation on the dropped row
      setTimeout(() => {
        const droppedRow = els.loopTableBody.querySelector(`tr[data-drag-row="${targetIndex}"]`);
        if (droppedRow) {
          droppedRow.style.transition = "background-color 0.8s ease";
          droppedRow.style.backgroundColor = "rgba(59, 130, 246, 0.3)";
          setTimeout(() => {
             // Let css restore the original via inline style cleanup if needed,
             // or just fade back
             droppedRow.style.backgroundColor = "";
          }, 800);
        }
      }, 50);
    });

    els.loopTableBody.addEventListener("change", (e) => {
      const lineTypeSelect = e.target.closest('[data-loop-field="lineType"]');
      if (!lineTypeSelect) return;
      syncLoopStateFromTable();
      renderLoopInputs();
    });
  }

  if (els.importInput) {
    els.importInput.addEventListener("change", async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        await importLevelsFile(file);
      } catch (err) {
        console.error("Level import error:", err);
        alert(`Import failed: ${err?.message || "Unable to parse levels file."}`);
      } finally {
        e.target.value = "";
      }
    });
  }

  if (els.bridgeImportInput) {
    els.bridgeImportInput.addEventListener("change", async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        await importBridgeFile(file);
      } catch (err) {
        console.error("Bridge import error:", err);
        alert(`Bridge import failed: ${err?.message || "Unable to parse bridge sheet."}`);
      } finally {
        e.target.value = "";
      }
    });
  }

  if (els.curveImportInput) {
    els.curveImportInput.addEventListener("change", async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        await importCurveFile(file);
      } catch (err) {
        console.error("Curve import error:", err);
        alert(`Curve import failed: ${err?.message || "Unable to parse curve list."}`);
      } finally {
        e.target.value = "";
      }
    });
  }

  if (els.loopImportInput) {
    els.loopImportInput.addEventListener("change", async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        await importLoopFile(file);
      } catch (err) {
        console.error("Loop/platform import error:", err);
        alert(`Loops/Platforms import failed: ${err?.message || "Unable to parse loop/platform list."}`);
      } finally {
        e.target.value = "";
      }
    });
  }

  if (els.projectImportInput) {
    els.projectImportInput.addEventListener("change", async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        const name = String(file.name || "");
        if (!/\.ew$/i.test(name)) {
          throw new Error("Only .EW project files are allowed.");
        }
        const txt = await file.text();
        const parsed = JSON.parse(txt);
        loadProjectFromPayload(parsed, { fileHandle: null });
      } catch (err) {
        console.error("Project import error:", err);
        alert(`Project import failed: ${err?.message || "Invalid .EW file."}`);
      } finally {
        e.target.value = "";
      }
    });
  }

  els.openSettingsBtn.addEventListener("click", () => {
    buildSettingsInputs();
    els.settingsModal.showModal();
  });
  els.closeSettingsBtn.addEventListener("click", () => els.settingsModal.close());

  els.settingsForm.addEventListener("submit", (e) => {
    e.preventDefault();
    applySettingsFromForm();
    els.settingsModal.close();
  });

  els.resetDefaultsBtn.addEventListener("click", () => {
    state.settings = { ...state.defaultSettings };
    buildSettingsInputs();
  });

  if (els.openMappingWizardBtn && els.mappingWizardModal) {
    els.openMappingWizardBtn.addEventListener("click", () => {
      if (els.mappingClientInput) els.mappingClientInput.value = state.importMappings?.client || "";
      if (els.mappingTypeSelect) els.mappingTypeSelect.value = "levels";
      if (els.mappingFileInput) els.mappingFileInput.value = "";
      state.mappingWizard = { aoa: null, headerRow: 0, headers: [], kind: "levels" };
      renderMappingGrid([], "levels");
      els.mappingWizardModal.showModal();
    });
  }
  if (els.closeMappingWizardBtn && els.mappingWizardModal) {
    els.closeMappingWizardBtn.addEventListener("click", () => els.mappingWizardModal.close());
  }
  if (els.mappingTypeSelect) {
    els.mappingTypeSelect.addEventListener("change", () => {
      const kind = String(els.mappingTypeSelect.value || "levels");
      state.mappingWizard.kind = kind;
      const client = String(els.mappingClientInput?.value || "").trim();
      const preset = getMappingTemplate(client, kind);
      renderMappingGrid(state.mappingWizard.headers, kind, preset);
    });
  }
  if (els.mappingFileInput) {
    els.mappingFileInput.addEventListener("change", async () => {
      const file = els.mappingFileInput.files?.[0];
      if (!file) return;
      try {
        const aoa = await readSheetAoaFromFile(file);
        const headerPick = detectSimpleHeaderRow(aoa, []);
        const headerRow = headerPick ? headerPick.rowIndex : 0;
        const headers = Array.isArray(aoa[headerRow]) ? aoa[headerRow] : [];
        state.mappingWizard = { aoa, headerRow, headers, kind: state.mappingWizard.kind || "levels" };
        const client = String(els.mappingClientInput?.value || "").trim();
        const preset = getMappingTemplate(client, state.mappingWizard.kind);
        renderMappingGrid(headers, state.mappingWizard.kind, preset);
      } catch (err) {
        console.error("Mapping wizard file error:", err);
        alert(`Mapping wizard failed: ${err.message}`);
      }
    });
  }
  if (els.saveMappingTemplateBtn) {
    els.saveMappingTemplateBtn.addEventListener("click", () => {
      const client = String(els.mappingClientInput?.value || "").trim();
      if (!client) {
        alert("Enter a client/template name.");
        return;
      }
      const kind = state.mappingWizard.kind || "levels";
      const mapping = {};
      const fields = IMPORT_MAPPING_FIELDS[kind] || [];
      fields.forEach((f) => {
        const sel = els.mappingGrid?.querySelector(`[data-map-key="${f.key}"]`);
        const val = sel ? Number(sel.value) : NaN;
        if (Number.isFinite(val)) mapping[f.key] = val;
      });
      if (!state.importMappings.templates[client]) state.importMappings.templates[client] = {};
      state.importMappings.templates[client][kind] = mapping;
      state.importMappings.client = client;
      alert("Mapping template saved.");
    });
  }
  if (els.applyMappingTemplateBtn) {
    els.applyMappingTemplateBtn.addEventListener("click", () => {
      const client = String(els.mappingClientInput?.value || "").trim();
      if (client) state.importMappings.client = client;
      els.mappingWizardModal?.close();
    });
  }

  if (els.materialProfileGrid) {
    els.materialProfileGrid.addEventListener("click", (e) => {
      const delBtn = e.target.closest("[data-mp-del]");
      if (delBtn) {
        const idx = Number(delBtn.dataset.mpDel);
        if (Number.isFinite(idx)) {
          state.settings.materialProfile = (state.settings.materialProfile || []).filter((_, i) => i !== idx);
          buildSettingsInputs();
        }
        return;
      }
      if (e.target?.id === "addMaterialRowBtn") {
        const next = Array.isArray(state.settings.materialProfile) ? [...state.settings.materialProfile] : [];
        next.push({ depth: 1, density: 1.9, shrink: 0, swell: 0 });
        state.settings.materialProfile = next;
        buildSettingsInputs();
      }
    });
  }

  els.tableBody.addEventListener("click", (e) => {
    const trigger = e.target.closest("[data-cross-index]");
    if (!trigger) return;
    e.preventDefault();
    e.stopPropagation();
    const i = Number(trigger.dataset.crossIndex);
    if (Number.isFinite(i)) openCrossSectionByIndex(i);
  });

  document.addEventListener("click", (e) => {
    const trigger = e.target.closest("[data-cross-index]");
    if (!trigger) return;
    if (els.tableBody && els.tableBody.contains(trigger)) return;
    e.preventDefault();
    const i = Number(trigger.dataset.crossIndex);
    if (Number.isFinite(i)) openCrossSectionByIndex(i);
  });

  if (els.crossPrevBtn) {
    els.crossPrevBtn.addEventListener("click", () => {
      openCrossSectionByIndex((state.currentCrossIndex ?? 0) - 1);
    });
  }
  if (els.crossNextBtn) {
    els.crossNextBtn.addEventListener("click", () => {
      openCrossSectionByIndex((state.currentCrossIndex ?? -1) + 1);
    });
  }
  els.closeCrossBtn.addEventListener("click", () => {
    els.crossSectionModal.close();
    state.currentCrossIndex = null;
    updateCrossSectionNav();
  });
  els.zoomInBtn.addEventListener("click", () => {
    const rect = els.crossSvg.getBoundingClientRect();
    zoomCrossAt(rect.left + (rect.width / 2), rect.top + (rect.height / 2), 0.86);
  });
  els.zoomOutBtn.addEventListener("click", () => {
    const rect = els.crossSvg.getBoundingClientRect();
    zoomCrossAt(rect.left + (rect.width / 2), rect.top + (rect.height / 2), 1.16);
  });
  els.zoomResetBtn.addEventListener("click", () => {
    resetCrossView();
  });
  
  /* NEW ADVANCED WORKFLOW UI LISTENERS */
  els.splitViewBtn = document.getElementById("splitViewBtn");
  els.mainSplitter = document.getElementById("mainSplitter");
  els.tableFab = document.getElementById("tableFab");
  els.fabJumpProblem = document.getElementById("fabJumpProblem");
  els.fabRecalc = document.getElementById("fabRecalc");
  els.tableWrapUI = document.querySelector('.work-page[data-work-page="table"] .table-wrap');
  els.diagnosticMinimap = document.getElementById("diagnosticMinimap");
  els.diagnosticBarcode = document.getElementById("diagnosticBarcode");
  els.xrayStationCanvas = document.getElementById("xrayStationCanvas");
  els.rollScrubber = document.getElementById("rollScrubber");
  els.rollPopupThumb = document.getElementById("rollPopupThumb");
  els.rollPopupThumbTitle = document.getElementById("rollPopupThumbTitle");
  els.rollPopupThumbCanvas = document.getElementById("rollPopupThumbCanvas");

  // 1. SPLIT VIEW
  if (els.splitViewBtn) {
    els.splitViewBtn.addEventListener("click", () => {
      const mainContent = document.querySelector(".main-content");
      if (mainContent.classList.contains("split-mode")) {
        mainContent.classList.remove("split-mode");
        els.splitViewBtn.classList.remove('active');
        if(els.mainSplitter) els.mainSplitter.classList.add("hidden");
        document.body.style.setProperty("--split-pos", "50%");
        setTimeout(() => window.dispatchEvent(new Event("resize")), 100);
      } else {
        mainContent.classList.add("split-mode");
        els.splitViewBtn.classList.add('active');
        if(els.mainSplitter) els.mainSplitter.classList.remove("hidden");
        setTimeout(() => window.dispatchEvent(new Event("resize")), 100);
      }
    });

    if (els.mainSplitter) {
      let isResizing = false;
      els.mainSplitter.addEventListener("mousedown", (e) => {
        isResizing = true;
        els.mainSplitter.classList.add("dragging");
      });
      document.addEventListener("mousemove", (e) => {
        if (!isResizing) return;
        const pct = (e.clientX / window.innerWidth) * 100;
        document.body.style.setProperty("--split-pos", `${pct}%`);
      });
      document.addEventListener("mouseup", () => {
        if(isResizing) window.dispatchEvent(new Event("resize"));
        isResizing = false;
        els.mainSplitter.classList.remove("dragging");
      });
    }
  }

  // 2. TABLE FAB
  if (els.tableWrapUI && els.tableFab) {
    els.tableWrapUI.addEventListener("scroll", () => {
      if (els.tableWrapUI.scrollTop > 200) {
        els.tableFab.classList.add("visible");
      } else {
        els.tableFab.classList.remove("visible");
      }
    });
    if (els.fabRecalc) {
      els.fabRecalc.addEventListener("click", () => {
        recalculate();
      });
    }
    if (els.fabJumpProblem) {
      els.fabJumpProblem.addEventListener("click", () => {
        const issues = els.tableBody.querySelectorAll(".bridge-row, [style*='color: red']");
        if (issues.length) {
          let nextIssue = Array.from(issues).find(el => el.getBoundingClientRect().top > (window.innerHeight/2 + 50));
          if(!nextIssue) nextIssue = issues[0];
          nextIssue.scrollIntoView({ behavior: 'smooth', block: 'center' });
          nextIssue.style.backgroundColor = "rgba(239, 68, 68, 0.2)";
          setTimeout(() => nextIssue.style.backgroundColor = "", 1000);
        } else {
          alert("No significant problem rows detected.");
        }
      });
    }
  }

  // 3. INLINE VALIDATION
  document.addEventListener("input", (e) => {
    if (e.target.tagName !== "INPUT" && e.target.tagName !== "SELECT") return;
    const val = Number(e.target.value);
    let isInvalid = false;
    if (e.target.dataset.loopField === 'tc' && val !== 0 && val < 4.72) isInvalid = true;
    if (e.target.dataset.bridgeField === 'startChainage' && val < 0) isInvalid = true;
    if (e.target.dataset.loopField === 'pfWidth' && val !== 0 && val < 4) isInvalid = true;
  
    if (isInvalid) {
      e.target.classList.add("cell-invalid");
    } else {
      e.target.classList.remove("cell-invalid");
    }
  });

  // 4. ROLL DIAGRAM SCRUBBER
  const rdWrap = document.getElementById("rollDiagramWrap");
  if (rdWrap && els.rollScrubber) {
    rdWrap.addEventListener("mousemove", (e) => {
      if(!state.calcRows || state.calcRows.length === 0) return;
      const rect = els.rollDiagramCanvas.getBoundingClientRect();
      const x = e.clientX - rect.left;
      if (x < 0 || x > rect.width) {
        els.rollScrubber.style.display = "none";
        els.rollPopupThumb.style.display = "none";
        return;
      }
      els.rollScrubber.style.display = "block";
      els.rollScrubber.style.left = `${x}px`;
      els.rollPopupThumb.style.display = "flex";
      els.rollPopupThumb.style.left = `${Math.min(x + 10, rect.width - 190)}px`;
      
      const pct = x / rect.width;
      const minCh = state.calcRows[0].chainage;
      const maxCh = state.calcRows[state.calcRows.length - 1].chainage;
      const scrubCh = minCh + (maxCh - minCh) * pct;
      const chText = (scrubCh < 0 ? "-" : "") + Math.floor(Math.abs(scrubCh) / 1000) + "+" + (Math.abs(scrubCh) % 1000).toFixed(0).padStart(3, "0");
      els.rollPopupThumbTitle.textContent = `CH ${chText}`;

      if (typeof drawMiniCrossSection === "function") {
         let bestRow = state.calcRows[0];
         let minDist = Infinity;
         for (let r of state.calcRows) {
            let dist = Math.abs(r.chainage - scrubCh);
            if(dist < minDist) { minDist = dist; bestRow = r; }
         }
         drawMiniCrossSection(els.rollPopupThumbCanvas, bestRow);
      }
    });
    rdWrap.addEventListener("mouseleave", () => {
       els.rollScrubber.style.display = "none";
       els.rollPopupThumb.style.display = "none";
    });
  }

  // 5. STATION LAYOUT X-RAY 
  if (els.stationLayoutSelect) {
    els.stationLayoutSelect.addEventListener("change", () => updateXRayStation());
  }

}

function updateDiagnosticMinimap() {
  if (!els.diagnosticBarcode || !els.diagnosticMinimap) return;
  if (!state.calcRows || state.calcRows.length === 0) {
    els.diagnosticMinimap.classList.remove('visible');
    return;
  }
  
  els.diagnosticMinimap.classList.add('visible');
  els.diagnosticBarcode.innerHTML = "";
  
  const total = state.calcRows.length;
  const maxBars = 100;
  const step = Math.ceil(total / maxBars);
  for(let i=0; i<total; i+=step) {
    const slice = state.calcRows.slice(i, i+step);
    const hasError = slice.some(r => r.cut > 8 || r.bank > 8 || (r.loopTc && Math.abs(r.loopTc) < 4.72));
    const div = document.createElement("div");
    div.className = "diag-bar" + (hasError ? " diag-error" : "");
    div.title = `Ch ${slice[0].chainage.toFixed(0)}`;
    div.onclick = () => {
       setWorkPage('table');
       const tableRow = els.tableBody.querySelector(`tr[data-ch-index="${i}"]`);
       if(tableRow) {
         tableRow.scrollIntoView({behavior: 'smooth', block: 'center'});
         tableRow.style.backgroundColor = "rgba(59, 130, 246, 0.3)";
         setTimeout(() => tableRow.style.backgroundColor = "", 1500);
       }
    };
    els.diagnosticBarcode.appendChild(div);
  }
}

function updateXRayStation() {
  if(!els.xrayStationCanvas) return;
  const ctx = els.xrayStationCanvas.getContext("2d");
  const cw = els.xrayStationCanvas.width;
  const ch = els.xrayStationCanvas.height;
  ctx.clearRect(0,0,cw,ch);
  
  if(!els.stationLayoutSelect) return;
  const stationName = els.stationLayoutSelect.value;
  if(!stationName) return;
  
  ctx.strokeStyle = "#94a3b8";
  ctx.lineWidth = 1.5;
  ctx.beginPath();
  ctx.moveTo(10, ch/2);
  ctx.lineTo(cw-10, ch/2);
  ctx.stroke();
  
  const stationRows = state.loopPlatformRows.filter(r => r.station === stationName && r.lineType !== "Main Line");
  stationRows.forEach((r, i) => {
    const side = r.side === "RHS" ? 1 : -1;
    const tcVal = Number.isFinite(r.tc) ? Math.abs(r.tc) : 5;
    const yOff = ch/2 + (side * (8 + (tcVal * 1.5)));
    
    ctx.strokeStyle = "#3b82f6";
    ctx.beginPath();
    ctx.moveTo(30 + (i*15), ch/2);
    ctx.lineTo(45 + (i*15), yOff);
    ctx.lineTo(cw - 45 - (i*15), yOff);
    ctx.lineTo(cw - 30 - (i*15), ch/2);
    ctx.stroke();
    
    const pfVal = Number.isFinite(r.pfWidth) ? r.pfWidth : 0;
    if (pfVal > 0) {
      ctx.fillStyle = "rgba(236, 72, 153, 0.8)";
      ctx.fillRect(cw/2 - 30, yOff + (side * 4) - 2.5, 60, pfVal);
    }
  });
}

function drawMiniCrossSection(canvas, row) {
  if(!canvas) return;
  const ctx = canvas.getContext("2d");
  const cw = canvas.width;
  const ch = canvas.height;
  ctx.clearRect(0,0,cw,ch);
  if(!row) return;

  const bVal = row.bank > 0 ? row.bank : 0;
  const cVal = row.cut > 0 ? row.cut : 0;
  
  ctx.strokeStyle = "#3b82f6";
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(cw/2 - 20, ch/2 - (bVal * 2));
  ctx.lineTo(cw/2 + 20, ch/2 - (bVal * 2)); 
  
  if (bVal > 0) {
     ctx.lineTo(cw/2 + 20 + bVal * 4, ch/2);
     ctx.lineTo(cw/2 - 20 - bVal * 4, ch/2);
     ctx.lineTo(cw/2 - 20, ch/2 - bVal * 2);
     ctx.fillStyle = "rgba(34, 197, 94, 0.4)";
     ctx.fill();
  } else if (cVal > 0) {
     ctx.lineTo(cw/2 + 20 + cVal * 3, ch/2 - cVal * 2);
     ctx.lineTo(cw/2 - 20 - cVal * 3, ch/2 - cVal * 2);
     ctx.lineTo(cw/2 - 20, ch/2);
     ctx.fillStyle = "rgba(244, 63, 94, 0.4)";
     ctx.fill();
  }
  ctx.stroke();
  
  ctx.fillStyle = "#f8fafc";
  ctx.font = "12px sans-serif";
  ctx.fillText(bVal > 0 ? `FILL ${r3(bVal)}m` : (cVal > 0 ? `CUT ${r3(cVal)}m` : "LEVEL"), cw/2 - 20, ch/2 + 15);
}

function saveState() {
  state.meta = state.meta || {};
  state.meta.lastSavedAt = new Date().toISOString();
  const data = {
    project: state.project,
    meta: state.meta,
    rawRows: state.rawRows,
    bridgeRows: state.bridgeRows,
    curveRows: state.curveRows,
    loopPlatformRows: state.loopPlatformRows,
    settings: state.settings,
    snapshots: state.snapshots,
    kmlData: state.kmlData,
    stationPlans: state.stationPlans,
  };
  localStorage.setItem("earthsoft_saved_work", JSON.stringify(data));
  
  // Push to Undo/Redo Stack
  if (!window._isUndoRedoOp && window._History) {
     window._History.push(data);
  }
}

function loadStoredState() {
  const saved = localStorage.getItem("earthsoft_saved_work");
  if (!saved) return;
  try {
    const data = JSON.parse(saved);
    if (!data) return;
    Object.assign(state, data);
  } catch (e) {
    console.error("Failed to load state:", e);
  }
}

// --- Undo / Redo Engine (Time Travel) ---
window._History = {
  stack: [],
  currentIndex: -1,
  push(stateSnapshot) {
     this.stack = this.stack.slice(0, this.currentIndex + 1);
     this.stack.push(JSON.stringify(stateSnapshot));
     if (this.stack.length > 20) this.stack.shift();
     else this.currentIndex++;
  },
  undo() {
     if (this.currentIndex > 0) {
        this.currentIndex--;
        return JSON.parse(this.stack[this.currentIndex]);
     }
     return null;
  },
  redo() {
     if (this.currentIndex < this.stack.length - 1) {
        this.currentIndex++;
        return JSON.parse(this.stack[this.currentIndex]);
     }
     return null;
  }
};

window.addEventListener("keydown", (e) => {
   if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
      const isShift = e.shiftKey;
      const prevState = isShift ? window._History.redo() : window._History.undo();
      if (prevState) {
         window._isUndoRedoOp = true;
         Object.assign(state, prevState);
         recalculate();
         updateWizardUI();
         buildSettingsInputs();
         buildTable();
         setTimeout(() => window._isUndoRedoOp = false, 100);
      }
   }
});

// --- 3D Fly-Through Viewer (Conceptual Demo) ---
let viewer3dScene, viewer3dCamera, viewer3dRenderer, viewer3dControls;
let animationId;
let flying = false;
let flyZ = 0;
const flyStep = 10;
let flySpeed = 0.5; // Default speed multiplier

function init3DViewer() {
   const container = document.getElementById("threeContainer");
   if(!container || viewer3dScene) return;

   viewer3dScene = new THREE.Scene();
   viewer3dScene.background = new THREE.Color(0x0b1020);
   viewer3dScene.fog = new THREE.FogExp2(0x0b1020, 0.002);

   viewer3dCamera = new THREE.PerspectiveCamera(60, container.clientWidth / container.clientHeight, 0.1, 2000);
   viewer3dCamera.position.set(20, 15, 40);

   viewer3dRenderer = new THREE.WebGLRenderer({ antialias: true, alpha: true });
   viewer3dRenderer.setSize(container.clientWidth, container.clientHeight);
   container.appendChild(viewer3dRenderer.domElement);

   viewer3dControls = new THREE.OrbitControls(viewer3dCamera, viewer3dRenderer.domElement);
   viewer3dControls.enableDamping = true;
   viewer3dControls.dampingFactor = 0.05;

   const ambientLight = new THREE.AmbientLight(0xffffff, 0.6);
   viewer3dScene.add(ambientLight);
   const dirLight = new THREE.DirectionalLight(0xffffff, 1.0);
   dirLight.position.set(100, 200, 100);
   viewer3dScene.add(dirLight);

   // Decorative sky sphere (glow effect)
   const skyGeo = new THREE.SphereGeometry(1500, 32, 32);
   const skyMat = new THREE.MeshBasicMaterial({ color: 0x0f172a, side: THREE.BackSide });
   const sky = new THREE.Mesh(skyGeo, skyMat);
   viewer3dScene.add(sky);

   // Grid base
   const gridHelper = new THREE.GridHelper(5000, 200, 0x1e293b, 0x0f172a);
   gridHelper.position.y = -10;
   viewer3dScene.add(gridHelper);

   generate3DMesh();

   function animate() {
     if (!viewer3dScene) return;
     animationId = requestAnimationFrame(animate);
     
     if (flying && state.calcRows.length > 0) {
        const speedMultiplier = parseFloat(document.getElementById("flySpeedSelect")?.value || "0.5");
        flyZ -= (flyStep * 0.1) * speedMultiplier;
        
        const totalLen = (state.calcRows.length - 1) * flyStep;
        if (Math.abs(flyZ) >= totalLen) {
           flyZ = 0; // Loop flight
        }
        
        const idx = Math.floor(Math.abs(flyZ) / flyStep);
        if (state.calcRows[idx]) {
           const row = state.calcRows[idx];
           const targetY = (row.proposedLevel / 10) + 2.5; // Fly slightly above track (lowered for intimacy)
           
           // Smooth camera movement (lerp)
           viewer3dCamera.position.z += (flyZ - viewer3dCamera.position.z) * 0.05;
           viewer3dCamera.position.y += (targetY - viewer3dCamera.position.y) * 0.05;
           viewer3dCamera.position.x += (0 - viewer3dCamera.position.x) * 0.05;
           
           // Look ahead logic
           const lookAheadIdx = Math.min(idx + 15, state.calcRows.length - 1);
           const aheadRow = state.calcRows[lookAheadIdx];
           const lookTarget = new THREE.Vector3(0, aheadRow.proposedLevel / 10, flyZ - 60);
           viewer3dCamera.lookAt(lookTarget);
        }
        viewer3dControls.enabled = false;
     } else {
        viewer3dControls.enabled = true;
        viewer3dControls.update();
     }
     
     viewer3dRenderer.render(viewer3dScene, viewer3dCamera);
   }
   animate();
   
   window.addEventListener('resize', () => {
      if(container.clientWidth === 0) return;
      viewer3dCamera.aspect = container.clientWidth / container.clientHeight;
      viewer3dCamera.updateProjectionMatrix();
      viewer3dRenderer.setSize(container.clientWidth, container.clientHeight);
   });
}

function generate3DMesh() {
   if(!viewer3dScene || state.calcRows.length === 0) return;
   
   // Clear old meshes
   viewer3dScene.children = viewer3dScene.children.filter(c => c.name !== 'trackCorridor');
   
   const trackGroup = new THREE.Group();
   trackGroup.name = 'trackCorridor';
   
   const matFill = new THREE.MeshLambertMaterial({ color: 0x22c55e, transparent: true, opacity: 0.7 });
   const matCut = new THREE.MeshLambertMaterial({ color: 0xef4444, transparent: true, opacity: 0.7 });
   const matTrack = new THREE.MeshPhongMaterial({ color: 0x334155, shininess: 30 });
   const matRail = new THREE.MeshStandardMaterial({ color: 0x94a3b8, metalness: 0.8, roughness: 0.2 });
   
   let zPos = 0;
   
   // Limit display for performance but keep long enough for flight
   const maxDisplay = Math.min(state.calcRows.length, 1000);
   
   for(let i=0; i < maxDisplay; i++) {
     const row = state.calcRows[i];
     const yPos = row.proposedLevel / 10;
     
     // Formation Bed
     const trackGeo = new THREE.BoxGeometry(7.85, 0.5, flyStep);
     const trackMesh = new THREE.Mesh(trackGeo, matTrack);
     trackMesh.position.set(0, yPos, zPos);
     trackGroup.add(trackMesh);

     // Simple Rails (Two lines)
     const railL = new THREE.Mesh(new THREE.BoxGeometry(0.1, 0.15, flyStep), matRail);
     railL.position.set(-0.835, yPos + 0.3, zPos);
     trackGroup.add(railL);

     const railR = new THREE.Mesh(new THREE.BoxGeometry(0.1, 0.15, flyStep), matRail);
     railR.position.set(0.835, yPos + 0.3, zPos);
     trackGroup.add(railR);
     
    // Embankment / Cutting visualization
    if (row.bank > 0.1) {
       const h = row.bank / 2;
       const fillGeo = new THREE.BoxGeometry(7.85 + (row.bank * 2), h, flyStep);
       const fillMesh = new THREE.Mesh(fillGeo, matFill);
       fillMesh.position.set(0, yPos - (1.2), zPos); // Fixed position below formation
       trackGroup.add(fillMesh);
    } else if (row.cut > 0.1) {
       const h = row.cut / 2;
       const cutGeo = new THREE.BoxGeometry(10.05 + (row.cut * 2), h, flyStep);
       const cutMesh = new THREE.Mesh(cutGeo, matCut);
       cutMesh.position.set(0, yPos + (1.2), zPos);
       trackGroup.add(cutMesh);
    }

    // --- Add Bridges ---
    const currentCh = row.chainage;
    const bridgeInRange = state.bridgeRows.find(b => {
       const start = parseChainage(b.startChainage);
       const end = parseChainage(b.endChainage);
       return currentCh >= start && currentCh <= end;
    });

    if (bridgeInRange) {
       const bridgeGeo = new THREE.BoxGeometry(10, 2, flyStep + 0.1);
       const bridgeMat = new THREE.MeshStandardMaterial({ color: 0x475569, metalness: 0.5 });
       const bridgeMesh = new THREE.Mesh(bridgeGeo, bridgeMat);
       bridgeMesh.position.set(0, yPos - 1.25, zPos);
       trackGroup.add(bridgeMesh);
       
       // Siderails for bridge
       const railBarGeo = new THREE.BoxGeometry(0.2, 1.2, flyStep);
       const railBarMat = new THREE.MeshStandardMaterial({ color: 0x94a3b8 });
       const leftBar = new THREE.Mesh(railBarGeo, railBarMat);
       leftBar.position.set(-5, yPos + 0.6, zPos);
       trackGroup.add(leftBar);
       const rightBar = new THREE.Mesh(railBarGeo, railBarMat);
       rightBar.position.set(5, yPos + 0.6, zPos);
       trackGroup.add(rightBar);
    }

    // --- Add Stations/Platforms ---
    const stationInRange = state.loopPlatformRows.find(s => {
       const start = parseChainage(s.pfStartCh);
       const end = parseChainage(s.pfEndCh);
       return (Number.isFinite(start) && Number.isFinite(end) && currentCh >= start && currentCh <= end);
    });

    if (stationInRange) {
       const pfSide = String(stationInRange.side).toLowerCase() === 'left' ? -1 : 1;
       const pfWidth = parseFloat(stationInRange.pfWidth) || 5;
       const pfGeo = new THREE.BoxGeometry(pfWidth, 0.4, flyStep);
       const pfMat = new THREE.MeshStandardMaterial({ color: 0x334155, roughness: 0.8 });
       const pfMesh = new THREE.Mesh(pfGeo, pfMat);
       pfMesh.position.set(pfSide * (3.925 + pfWidth/2), yPos + 0.2, zPos);
       trackGroup.add(pfMesh);
       
       // Station Sign/Pillar every 50m
       if (Math.abs(row.chainage % 50) < 5) {
          const pillarGeo = new THREE.BoxGeometry(0.5, 6, 0.5);
          const pillarMat = new THREE.MeshPhongMaterial({ color: 0x1e293b });
          const pillar = new THREE.Mesh(pillarGeo, pillarMat);
          pillar.position.set(pfSide * (3.925 + pfWidth - 0.5), yPos + 3, zPos);
          trackGroup.add(pillar);
          
          const boardGeo = new THREE.BoxGeometry(4, 1.5, 0.2);
          const boardMat = new THREE.MeshStandardMaterial({ color: 0xfacc15 });
          const board = new THREE.Mesh(boardGeo, boardMat);
          board.position.set(pfSide * (3.925 + pfWidth - 0.5), yPos + 5.5, zPos);
          trackGroup.add(board);
       }
    }
    
    zPos -= flyStep;
 }
 viewer3dScene.add(trackGroup);
}

document.addEventListener("DOMContentLoaded", () => {
   const viewerBtn = document.getElementById("viewer3dBtn");
   if(viewerBtn) {
       viewerBtn.addEventListener('click', () => {
           document.getElementById("viewer3dModal").showModal();
           setTimeout(() => {
               init3DViewer();
               generate3DMesh();
           }, 150);
       });
   }
   document.getElementById("close3dBtn")?.addEventListener('click', () => {
       document.getElementById("viewer3dModal").close();
       if (animationId) cancelAnimationFrame(animationId);
       viewer3dScene = null; // Reset for next open
   });
   
   const flyBtn = document.getElementById("playFlyBtn");
   flyBtn?.addEventListener('click', function() {
       flying = !flying;
       this.innerHTML = flying ? '<i class="ri-pause-fill" style="margin-right:4px;"></i> Pause Flight' : '<i class="ri-play-fill" style="margin-right:4px;"></i> Simulate Flight';
       if(flying) {
           flyZ = 0; // Restart from beginning
       }
   });
});

async function init() {
  const visualLayerDefaults = {
    // --- Core Earthwork Calculation Parameters ---
    formationWidthFill: 7.85,         // Formation width for filling (meters)
    blanketThickness: 0.3,            // Blanket layer thickness (meters)
    preparedSubgradeThickness: 0.0,   // Prepared subgrade thickness (meters)
    bermWidth: 3.0,                   // Berm width (meters)
    cuttingWidth: 12.05,              // Formation width for cutting (meters)
    sideSlopeFactor: 2.236,           // Side slope factor (1:2 slope = 2.236 hypotenuse factor)
    // --- Visual Layer Thicknesses ---
    sq3Thickness: 0.2,
    sq2Thickness: 0.15,
    aboveBallastThickness: 0.3,
    sleeperThickness: 0.25,
    railHeight: 0.18,
    ballastCushionThickness: 0.35,
    topLayerThickness: 0.5,
    activeSqCategory: 3,
    gauge: "BG",
    formationPreset: "single",
    slopePreset: "default",
    bermRulePrimary: 4,
    bermRuleSecondary: 8,
    showRails: true,
    showPlatforms: true,
    showDrains: true,
    showLabels: true,
    rollFilter: "all",
    minTc: 5.3,
    minPlatformWidth: 6.0,
    minClearance: 1.5,
    showMapStations: true,
    showMapBridges: true,
    showMapCurves: true,
    boqMapping: {
      range: "Range",
      prepared: "Prepared",
      blanket: "Blanket",
      fill: "Fill (m³)",
      cut: "Cut (m³)",
    },
    reportBrand: {
      logo: "",
      signature: "",
      signName: "",
      signTitle: "",
    },
    materialProfile: [
      { depth: 1.0, density: 1.9, shrink: 0.0, swell: 0.0 },
    ],
  };
  state.defaultSettings = { ...visualLayerDefaults };
  state.seedDefaultSettings = { ...state.defaultSettings };

  // Initialize defaults
  state.meta = {};
  state.settings = { ...state.defaultSettings };
  state.rawRows = [];
  state.seedRows = [];
  state.seedMeta = {};
  state.snapshots = [];

  loadStoredState();
  await loadAuthState();
  // Ensure settings always has all required keys (stored state may have stale/missing fields)
  state.settings = { ...state.defaultSettings, ...(state.settings || {}) };
  if (!Array.isArray(state.settings.materialProfile) || !state.settings.materialProfile.length) {
    state.settings.materialProfile = [...state.defaultSettings.materialProfile];
  }
  if (!state.settings.reportBrand || typeof state.settings.reportBrand !== "object") {
    state.settings.reportBrand = { ...state.defaultSettings.reportBrand };
  } else {
    state.settings.reportBrand = { ...state.defaultSettings.reportBrand, ...state.settings.reportBrand };
  }
  if (!state.settings.boqMapping || typeof state.settings.boqMapping !== "object") {
    state.settings.boqMapping = { ...state.defaultSettings.boqMapping };
  } else {
    state.settings.boqMapping = { ...state.defaultSettings.boqMapping, ...state.settings.boqMapping };
  }
  state.meta = state.meta || {};
  state.project = {
    active: Boolean(state.project?.active),
    verified: Boolean(state.project?.verified),
    name: String(state.project?.name || ""),
    uploads: {
      levels: Boolean(state.project?.uploads?.levels),
      curves: Boolean(state.project?.uploads?.curves),
      bridges: Boolean(state.project?.uploads?.bridges),
      loops: Boolean(state.project?.uploads?.loops),
      kml: Boolean(state.project?.uploads?.kml) || Boolean(state.kmlData),
    },
    profile: {
      corridorName: String(state.project?.profile?.corridorName || ""),
      direction: String(state.project?.profile?.direction || "Up"),
      chainageZeroRef: String(state.project?.profile?.chainageZeroRef || ""),
    },
  };
  state.calcOverrides = Array.isArray(state.calcOverrides) ? state.calcOverrides : [];
  state.slopeZones = Array.isArray(state.slopeZones) ? state.slopeZones : [];
  state.importMappings = state.importMappings && typeof state.importMappings === "object"
    ? state.importMappings
    : { client: "", templates: {} };

  bindEvents();
  bindCrossCanvasInteraction();
  syncRollFilterSelect();
  setWorkPage("overview");
  setResultTab("qty");
  updateAuthUI();
  updateWizardUI();
  applyProjectGate();
  recalculate();
}

init();

window._isPdfExportCancelled = false;

document.addEventListener('DOMContentLoaded', () => {
  const cancelBtn = document.getElementById('cancelPdfExportBtn');
  if (cancelBtn) {
    cancelBtn.addEventListener('click', () => {
      window._isPdfExportCancelled = true;
      const loadingText = document.getElementById('aiLoadingText');
      if (loadingText) loadingText.textContent = "Cancelling export safely... Please wait.";
    });
  }
});

function formatReportChainage(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return "-";
  const sign = n < 0 ? "-" : "";
  const abs = Math.abs(n);
  const km = Math.floor(abs / 1000);
  const m = Math.round(abs - (km * 1000));
  return `${sign}${km}+${String(m).padStart(3, "0")}`;
}

function createReportPage(title, subtitle = "") {
  const page = document.createElement("div");
  page.style.padding = "34px 42px";
  page.style.minHeight = "750px";
  page.style.boxSizing = "border-box";
  if (title) {
    const head = document.createElement("div");
    head.style.marginBottom = "24px";
    head.innerHTML = `
      <div style="display:flex; align-items:flex-end; justify-content:space-between; gap:16px; margin-bottom:12px;">
        <div>
          <div style="font-size:11px; letter-spacing:0.18em; text-transform:uppercase; color:#2563eb; font-weight:800;">Earthsoft Engineering Report</div>
          <h1 style="margin:8px 0 0; font-size:28px; letter-spacing:-0.03em; color:#0f172a;">${title}</h1>
        </div>
        <div style="text-align:right; font-size:11px; color:#64748b;">
          <div>${state.project.name || "Untitled Project"}</div>
          <div>${new Date().toLocaleDateString()}</div>
        </div>
      </div>
      ${subtitle ? `<p style="margin:0; color:#475569; font-size:13px; line-height:1.5;">${subtitle}</p>` : ""}
    `;
    page.appendChild(head);
  }
  return page;
}

function createReportSection(title, contentHtml) {
  const section = document.createElement("section");
  section.style.marginBottom = "18px";
  section.innerHTML = `
    <div style="margin-bottom:10px; display:flex; align-items:center; gap:10px;">
      <div style="width:22px; height:4px; border-radius:999px; background:#2563eb;"></div>
      <h2 style="margin:0; font-size:15px; color:#0f172a; letter-spacing:0.01em;">${title}</h2>
    </div>
    ${contentHtml}
  `;
  return section;
}

function reportMetricCard(label, value, tone = "slate", sub = "") {
  const toneMap = {
    slate: { bg: "#f8fafc", border: "#e2e8f0", value: "#0f172a" },
    blue: { bg: "#eff6ff", border: "#bfdbfe", value: "#1d4ed8" },
    green: { bg: "#ecfdf5", border: "#a7f3d0", value: "#047857" },
    red: { bg: "#fef2f2", border: "#fecaca", value: "#b91c1c" },
    amber: { bg: "#fffbeb", border: "#fde68a", value: "#b45309" },
  };
  const palette = toneMap[tone] || toneMap.slate;
  return `
    <div style="padding:14px 16px; border-radius:14px; background:${palette.bg}; border:1px solid ${palette.border};">
      <div style="font-size:10px; font-weight:800; letter-spacing:0.12em; text-transform:uppercase; color:#64748b;">${label}</div>
      <div style="margin-top:8px; font-size:22px; font-weight:800; color:${palette.value};">${value}</div>
      ${sub ? `<div style="margin-top:4px; font-size:11px; color:#64748b;">${sub}</div>` : ""}
    </div>
  `;
}

function buildReportContext(rows) {
  const first = rows[0] || {};
  const last = rows[rows.length - 1] || {};
  const fillTotal = rows.reduce((acc, row) => acc + safeNum(row.fillVol, 0), 0);
  const cutTotal = rows.reduce((acc, row) => acc + safeNum(row.cutVol, 0), 0);
  const totalLength = Math.max(safeNum(last.chainage) - safeNum(first.chainage), 0);
  const groupedStations = getGroupedStations();
  const reusablePct = parseLooseNumber(els.pctReusableSpoil?.value, 60);
  const reusableSpoil = cutTotal * (safeNum(reusablePct, 60) / 100);
  const balance = reusableSpoil - fillTotal;
  const uploads = state.project?.uploads || {};
  const alerts = [];
  if (!uploads.kml) alerts.push("KML/KMZ alignment not attached");
  if (!state.project?.verified) alerts.push("Project not verified");
  if (!uploads.loops) alerts.push("Loops & Platforms not imported");
  if (!uploads.bridges) alerts.push("Bridge deductions not imported");
  if (!uploads.curves) alerts.push("Curve list not imported");
  if (!Object.keys(state.stationPlans || {}).length) alerts.push("No station conceptual plans attached");
  return {
    first,
    last,
    fillTotal,
    cutTotal,
    totalLength,
    groupedStations,
    reusablePct,
    reusableSpoil,
    balance,
    alerts,
    validBridges: Array.isArray(state.bridgeRows) ? state.bridgeRows.length : 0,
    validCurves: Array.isArray(state.curveRows) ? state.curveRows.length : 0,
  };
}

async function generateProjectReport(options) {
  window._isPdfExportCancelled = false;

  const JsPdfCtor = (window.jspdf && window.jspdf.jsPDF) ? window.jspdf.jsPDF : window.jsPDF;
  if (!JsPdfCtor) {
    alert("jsPDF is not available. Please refresh and try again.");
    return;
  }

  const loading = document.getElementById("aiLoadingOverlay");
  const loadingText = document.getElementById("aiLoadingText");
  const loadingHeader = document.getElementById("aiLoadingHeader");
  const progressBarContainer = document.getElementById("aiProgressBarContainer");
  const progressBar = document.getElementById("aiProgressBar");

  if (loading) {
    loading.classList.remove("hidden");
    if (loadingHeader) loadingHeader.textContent = "System Processing...";
    if (loadingText) loadingText.textContent = "Preparing project report PDF...";
    if (progressBarContainer) progressBarContainer.style.display = "block";
    if (progressBar) progressBar.style.width = "0%";
  }

  try {
    let exportCalcRows = [...state.calcRows];
    if (options.exportRange) {
      const start = Number.isFinite(options.exportRange.start) ? options.exportRange.start : -Infinity;
      const end = Number.isFinite(options.exportRange.end) ? options.exportRange.end : Infinity;
      const minCh = Math.min(start, end);
      const maxCh = Math.max(start, end);
      exportCalcRows = exportCalcRows.filter((r) => Number.isFinite(r.chainage) && r.chainage >= minCh && r.chainage <= maxCh);
    }
    const exportRowsLimit = options.rowLimit && options.rowLimit > 0 ? options.rowLimit : exportCalcRows.length;
    exportCalcRows = exportCalcRows.slice(0, exportRowsLimit);

    if (exportCalcRows.length === 0) {
      throw new Error("No data to export.");
    }

    if (progressBar) progressBar.style.width = "5%";

    const cancelBtn = document.getElementById("cancelPdfExportBtn");
    if (cancelBtn) cancelBtn.style.display = "block";

    const reportCtx = buildReportContext(exportCalcRows);
    const reportBrand = state.settings.reportBrand || {};

    // --- Color palette ---
    const C = {
      primary: [37, 99, 235],
      dark: [15, 23, 42],
      text: [51, 65, 85],
      muted: [100, 116, 139],
      light: [241, 245, 249],
      white: [255, 255, 255],
      green: [4, 120, 87],
      red: [185, 28, 28],
      amber: [180, 83, 9],
      blue: [29, 78, 216],
      border: [226, 232, 240],
    };

    // Helper: draw a rounded rect
    function roundedRect(pdf, x, y, w, h, r, fillColor, strokeColor) {
      if (fillColor) { pdf.setFillColor(...fillColor); }
      if (strokeColor) { pdf.setDrawColor(...strokeColor); pdf.setLineWidth(0.3); }
      const mode = fillColor && strokeColor ? "FD" : fillColor ? "F" : "S";
      pdf.roundedRect(x, y, w, h, r, r, mode);
    }

    // Helper: metric card drawn directly
    function drawMetricCard(pdf, x, y, w, h, label, value, tone, sub) {
      const tones = {
        slate: { bg: [248, 250, 252], border: C.border, val: C.dark },
        blue: { bg: [239, 246, 255], border: [191, 219, 254], val: C.blue },
        green: { bg: [236, 253, 245], border: [167, 243, 208], val: C.green },
        red: { bg: [254, 242, 242], border: [254, 202, 202], val: C.red },
        amber: { bg: [255, 251, 235], border: [253, 230, 138], val: C.amber },
      };
      const t = tones[tone] || tones.slate;
      roundedRect(pdf, x, y, w, h, 2, t.bg, t.border);
      pdf.setFontSize(7); pdf.setFont("helvetica", "bold"); pdf.setTextColor(...C.muted);
      pdf.text(label.toUpperCase(), x + 3, y + 5);
      pdf.setFontSize(14); pdf.setFont("helvetica", "bold"); pdf.setTextColor(...t.val);
      pdf.text(String(value), x + 3, y + 13);
      if (sub) {
        pdf.setFontSize(7); pdf.setFont("helvetica", "normal"); pdf.setTextColor(...C.muted);
        pdf.text(String(sub), x + 3, y + 17);
      }
    }

    // Helper: section header
    function drawSectionHeader(pdf, y, title) {
      pdf.setDrawColor(...C.primary); pdf.setLineWidth(1);
      pdf.line(20, y, 24, y);
      pdf.setFontSize(12); pdf.setFont("helvetica", "bold"); pdf.setTextColor(...C.dark);
      pdf.text(title, 27, y + 1);
      return y + 6;
    }

    // Helper: page header/footer on content pages
    function drawPageHeaderFooter(pdf, pageTitle) {
      // Top bar
      pdf.setFontSize(7); pdf.setFont("helvetica", "bold"); pdf.setTextColor(...C.primary);
      pdf.text("EARTHSOFT ENGINEERING REPORT", 20, 12);
      pdf.setFont("helvetica", "normal"); pdf.setTextColor(...C.muted);
      pdf.text(state.project.name || "Untitled Project", 200, 12, { align: "right" });
      pdf.setDrawColor(...C.border); pdf.setLineWidth(0.3);
      pdf.line(20, 14, 200, 14);
      // Page title
      if (pageTitle) {
        pdf.setFontSize(18); pdf.setFont("helvetica", "bold"); pdf.setTextColor(...C.dark);
        pdf.text(pageTitle, 20, 24);
      }
      // Footer
      const pgH = pdf.internal.pageSize.getHeight();
      pdf.setFontSize(6); pdf.setTextColor(...C.muted); pdf.setFont("helvetica", "normal");
      pdf.text(`Generated ${new Date().toLocaleDateString()} • Earthsoft Professional`, 20, pgH - 8);
      pdf.text(`Page ${pdf.internal.getNumberOfPages()}`, 200, pgH - 8, { align: "right" });
    }

    // ===== 1. COVER PAGE (Portrait) =====
    if (loadingText) loadingText.textContent = "Building cover page...";
    const pdf = new JsPdfCtor({ unit: "mm", format: "a4", orientation: "portrait" });

    // Background gradient simulation
    roundedRect(pdf, 0, 0, 210, 297, 0, [245, 249, 255], null);
    roundedRect(pdf, 0, 0, 210, 80, 0, [239, 246, 255], null);

    // Badge
    roundedRect(pdf, 20, 30, 65, 7, 3, [239, 246, 255], [191, 219, 254]);
    pdf.setFontSize(7); pdf.setFont("helvetica", "bold"); pdf.setTextColor(...C.primary);
    pdf.text("EARTHSOFT PROFESSIONAL EXPORT", 23, 35);

    // Logo if exists
    if (reportBrand.logo) {
      try { pdf.addImage(reportBrand.logo, "PNG", 150, 25, 40, 15); } catch (e) { /* skip */ }
    }

    // Title
    pdf.setFontSize(32); pdf.setFont("helvetica", "bold"); pdf.setTextColor(...C.dark);
    pdf.text("Railway Earthwork", 20, 60);
    pdf.text("Report", 20, 72);

    // Project name
    pdf.setFontSize(16); pdf.setFont("helvetica", "normal"); pdf.setTextColor(...C.text);
    pdf.text(state.project.name || "Untitled Project", 20, 84);

    // Description
    pdf.setFontSize(10); pdf.setFont("helvetica", "normal"); pdf.setTextColor(100, 116, 139);
    const descLines = pdf.splitTextToSize("Executive engineering summary including project readiness, quantity balance, calculation sheets, roll diagrams, and cross-sectional detail drawings.", 170);
    pdf.text(descLines, 20, 94);

    // Metric cards row
    const cardW = 40; const cardH = 20; const cardY = 120;
    drawMetricCard(pdf, 20, cardY, cardW, cardH, "Chainage Range", `${formatReportChainage(reportCtx.first.chainage)} to ${formatReportChainage(reportCtx.last.chainage)}`, "blue");
    drawMetricCard(pdf, 63, cardY, cardW, cardH, "Total Length", `${r3(reportCtx.totalLength / 1000)} km`, "slate");
    drawMetricCard(pdf, 106, cardY, cardW, cardH, "Total Fill", formatVolume(reportCtx.fillTotal), "green");
    drawMetricCard(pdf, 149, cardY, cardW, cardH, "Total Cut", formatVolume(reportCtx.cutTotal), "red");

    // Snapshot box
    const snapY = 148;
    roundedRect(pdf, 20, snapY, 170, 55, 3, C.dark, null);
    pdf.setFontSize(8); pdf.setFont("helvetica", "bold"); pdf.setTextColor(125, 211, 252);
    pdf.text("PROJECT SNAPSHOT", 25, snapY + 8);

    const snapRows = [
      ["Generated", new Date().toLocaleString()],
      ["Verification", state.project?.verified ? "Verified" : "Draft"],
      ["Cross-Sections", String(exportCalcRows.length)],
      ["Bridges / Curves", `${reportCtx.validBridges} / ${reportCtx.validCurves}`],
      ["Mapped Stations", String(reportCtx.groupedStations.length)],
    ];
    pdf.setFontSize(9);
    snapRows.forEach(([label, value], i) => {
      const rowY = snapY + 16 + (i * 8);
      pdf.setFont("helvetica", "normal"); pdf.setTextColor(148, 163, 184);
      pdf.text(label, 25, rowY);
      pdf.setFont("helvetica", "bold"); pdf.setTextColor(226, 232, 240);
      pdf.text(value, 185, rowY, { align: "right" });
    });

    if (progressBar) progressBar.style.width = "10%";

    // ===== 2. EXECUTIVE SUMMARY =====
    if (window._isPdfExportCancelled) throw new Error("Export cancelled by user.");
    if (loadingText) loadingText.textContent = "Building executive summary...";
    pdf.addPage("a4", "portrait");
    drawPageHeaderFooter(pdf, "Executive Summary");

    let curY = 30;
    // Quantity Balance
    curY = drawSectionHeader(pdf, curY, "Quantity Balance");
    drawMetricCard(pdf, 20, curY, 40, 20, "Fill Volume", formatVolume(reportCtx.fillTotal), "green", "Total embankment");
    drawMetricCard(pdf, 63, curY, 40, 20, "Cut Volume", formatVolume(reportCtx.cutTotal), "red", "Total excavation");
    drawMetricCard(pdf, 106, curY, 40, 20, "Reusable Spoil", formatVolume(reportCtx.reusableSpoil), "amber", `${safeNum(reportCtx.reusablePct, 60).toFixed(0)}% reuse`);
    const balLabel = `${reportCtx.balance >= 0 ? "+" : ""}${formatCompactVolume(reportCtx.balance)} m³`;
    drawMetricCard(pdf, 149, curY, 40, 20, "Net Balance", balLabel, reportCtx.balance >= 0 ? "blue" : "red", reportCtx.balance >= 0 ? "Surplus" : "Borrow needed");
    curY += 26;

    // Project Readiness table
    curY = drawSectionHeader(pdf, curY, "Project Readiness");
    const readinessData = [
      ["Project Name", state.project.name || "Untitled Project"],
      ["Verification", state.project?.verified ? "Verified" : "Draft"],
      ["Levels Imported", state.project?.uploads?.levels ? "Yes" : "No"],
      ["Bridge List", state.project?.uploads?.bridges ? "Yes" : "No"],
      ["Curve List", state.project?.uploads?.curves ? "Yes" : "No"],
      ["Loops & Platforms", state.project?.uploads?.loops ? "Yes" : "No"],
      ["KML/KMZ Alignment", state.project?.uploads?.kml ? "Yes" : "No"],
      ["Station Plans", `${Object.keys(state.stationPlans || {}).length}`],
    ];
    pdf.autoTable({
      startY: curY,
      head: [],
      body: readinessData,
      theme: "grid",
      margin: { left: 20, right: 100 },
      styles: { fontSize: 8, cellPadding: 2, lineColor: C.border, lineWidth: 0.2 },
      columnStyles: { 0: { fontStyle: "bold", textColor: C.text, cellWidth: 40 }, 1: { textColor: C.dark } },
      alternateRowStyles: { fillColor: [248, 250, 252] },
    });
    curY = pdf.lastAutoTable.finalY + 6;

    // Design Notes
    if (reportCtx.alerts.length) {
      curY = drawSectionHeader(pdf, curY, "Design Notes");
      pdf.setFontSize(8); pdf.setFont("helvetica", "normal"); pdf.setTextColor(...C.text);
      reportCtx.alerts.forEach((alert, i) => {
        pdf.text(`• ${alert}`, 22, curY + (i * 5));
      });
      curY += reportCtx.alerts.length * 5 + 4;
    }

    // Field Intelligence
    const sharpestCurve = [...(state.curveRows || [])].filter((c) => safeNum(c.radius, NaN) > 0).sort((a, b) => safeNum(a.radius) - safeNum(b.radius))[0];
    const firstBridge = [...(state.bridgeRows || [])].sort((a, b) => safeNum(a.startChainage) - safeNum(b.startChainage))[0];
    const firstStation = reportCtx.groupedStations.filter((s) => Number.isFinite(s.csb)).sort((a, b) => safeNum(a.csb) - safeNum(b.csb))[0];
    curY = drawSectionHeader(pdf, curY, "Field Intelligence");
    drawMetricCard(pdf, 20, curY, 55, 20, "First Bridge", firstBridge ? `${firstBridge.bridgeNo || "Bridge"} @ ${formatReportChainage(firstBridge.startChainage)}` : "N/A", "blue");
    drawMetricCard(pdf, 78, curY, 55, 20, "Sharpest Curve", sharpestCurve ? `R=${r3(sharpestCurve.radius)} m` : "N/A", "amber");
    drawMetricCard(pdf, 136, curY, 55, 20, "First Station", firstStation ? firstStation.station : "N/A", "slate");

    if (progressBar) progressBar.style.width = "15%";

    // ===== OPTIONAL SECTIONS =====
    if (options.includeProfileSection) {
      if (window._isPdfExportCancelled) throw new Error("Export cancelled by user.");
      pdf.addPage("a4", "portrait"); drawPageHeaderFooter(pdf, "Project Profile");
      const profile = state.project?.profile || {};
      const profileData = [
        ["Project Name", state.project.name || "Untitled"],
        ["Corridor Name", profile.corridorName || "-"],
        ["Direction", profile.direction || "-"],
        ["Chainage Zero Ref", profile.chainageZeroRef || "-"],
      ];
      pdf.autoTable({
        startY: 30, head: [], body: profileData, theme: "grid",
        margin: { left: 20, right: 80 },
        styles: { fontSize: 9, cellPadding: 3, lineColor: C.border, lineWidth: 0.2 },
        columnStyles: { 0: { fontStyle: "bold", textColor: C.text, cellWidth: 50 }, 1: { textColor: C.dark } },
        alternateRowStyles: { fillColor: [248, 250, 252] },
      });
    }

    if (options.includeStandardsSection) {
      if (window._isPdfExportCancelled) throw new Error("Export cancelled by user.");
      pdf.addPage("a4", "portrait"); drawPageHeaderFooter(pdf, "Design Standards");
      const s = state.settings || {};
      const stdData = [
        ["Gauge", s.gauge || "BG"],
        ["Formation Preset", s.formationPreset || "single"],
        ["Formation Width (Fill)", `${r3(s.formationWidthFill)} m`],
        ["Cutting Width", `${r3(s.cuttingWidth)} m`],
        ["Side Slope (H:V)", r3(s.sideSlopeFactor)],
        ["Berm Rule 1", `${r3(s.bermRulePrimary)} m`],
        ["Berm Rule 2", `${r3(s.bermRuleSecondary)} m`],
        ["Min Track Centre", `${r3(s.minTc)} m`],
        ["Min Platform Width", `${r3(s.minPlatformWidth)} m`],
        ["Min Clearance", `${r3(s.minClearance)} m`],
      ];
      pdf.autoTable({
        startY: 30, head: [], body: stdData, theme: "grid",
        margin: { left: 20, right: 80 },
        styles: { fontSize: 9, cellPadding: 3, lineColor: C.border, lineWidth: 0.2 },
        columnStyles: { 0: { fontStyle: "bold", textColor: C.text, cellWidth: 50 }, 1: { textColor: C.dark } },
        alternateRowStyles: { fillColor: [248, 250, 252] },
      });
    }

    if (options.includeMaterialSection) {
      if (window._isPdfExportCancelled) throw new Error("Export cancelled by user.");
      pdf.addPage("a4", "portrait"); drawPageHeaderFooter(pdf, "Material Profile");
      const matRows = Array.isArray(state.settings.materialProfile) ? state.settings.materialProfile : [];
      const matBody = matRows.length ? matRows.map(r => [r3(r.depth), r3(r.density), r3(r.shrink), r3(r.swell)]) : [["No material layers defined", "", "", ""]];
      pdf.autoTable({
        startY: 30, head: [["Depth (m)", "Density (t/m³)", "Shrink (%)", "Swell (%)"]],
        body: matBody, theme: "grid", margin: { left: 20, right: 80 },
        styles: { fontSize: 9, cellPadding: 3, halign: "center", lineColor: C.border, lineWidth: 0.2 },
        headStyles: { fillColor: C.light, textColor: C.dark, fontStyle: "bold" },
      });
    }

    if (options.includeQualitySection) {
      if (window._isPdfExportCancelled) throw new Error("Export cancelled by user.");
      const issues = runQualityChecks();
      pdf.addPage("a4", "portrait"); drawPageHeaderFooter(pdf, "Quality Checks");
      if (!issues.length) {
        pdf.setFontSize(10); pdf.setFont("helvetica", "normal"); pdf.setTextColor(...C.text);
        pdf.text("No critical issues detected.", 20, 35);
      } else {
        const issueBody = issues.map(issue => [`${issue.title}: ${issue.detail}`]);
        pdf.autoTable({
          startY: 30, head: [["Issue"]], body: issueBody, theme: "grid",
          margin: { left: 20, right: 20 },
          styles: { fontSize: 8, cellPadding: 2, lineColor: C.border, lineWidth: 0.2 },
          headStyles: { fillColor: C.light, textColor: C.dark, fontStyle: "bold" },
        });
      }
    }

    if (progressBar) progressBar.style.width = "20%";

    // ===== 3. CALCULATION SHEET =====
    if (options.calcSheet) {
      if (window._isPdfExportCancelled) throw new Error("Export cancelled by user.");
      if (loadingText) loadingText.textContent = "Building calculation sheet...";

      // Summary page first
      pdf.addPage("a4", "landscape");
      drawPageHeaderFooter(pdf, "Calculation Sheet");

      // Quantity overview table from DOM
      let qtyData = [];
      const qtyRows = Array.from(document.querySelectorAll('#resultQtyBody tr'));
      if (qtyRows.length) {
        qtyData = qtyRows.map(tr => {
          const tds = Array.from(tr.querySelectorAll('td'));
          return tds.map(td => td.textContent.trim());
        });
      }
      if (qtyData.length) {
        pdf.autoTable({
          startY: 30,
          head: [["Range", "Prepared", "Blanket", "Fill (m³)", "Cut (m³)"]],
          body: qtyData, theme: "grid",
          margin: { left: 20, right: 20 },
          styles: { fontSize: 8, cellPadding: 2, halign: "center", lineColor: C.border, lineWidth: 0.2 },
          headStyles: { fillColor: C.light, textColor: C.dark, fontStyle: "bold" },
        });
      }

      // Full calculation table
      const calcHeaders = [
        "Bridge", "Station", "Chainage", "Diff", "Ground RL",
        "Proposed RL", "Loop TC", "Loops", "PF Width", "Deduct Len",
        "EW Diff", "RL Diff", "Bank", "Cut", "Top Width",
        "Fill Area", "Cut Area", "Fill Vol", "Cut Vol",
      ];

      const calcBody = exportCalcRows.map(r => {
        const bridgeRefs = (r.bridgeRefs && r.bridgeRefs.length) ? r.bridgeRefs.join("|") : "-";
        const station = (r.stationName || r.station) ? String(r.stationName || r.station) : "-";
        const ch = (r.chainage < 0 ? "-" : "") + Math.floor(Math.abs(r.chainage) / 1000) + "+" + (Math.abs(r.chainage) % 1000).toFixed(3).replace(/(\.\d*?[1-9])0+$|\.0+$/, "$1").padStart(3, "0");
        return [
          bridgeRefs, station, ch,
          r.diff ? r3(r.diff) : "—", r3(r.groundLevel), r3(r.proposedLevel),
          r.loopTc ? r3(r.loopTc) : "—", r.loopCount ? String(r.loopCount) : "—",
          r.platformWidth ? r3(r.platformWidth) : "—", r.bridgeDeductLen ? r3(r.bridgeDeductLen) : "—",
          r.ewDiff ? r3(r.ewDiff) : "—", r.rlDiff ? r3(r.rlDiff) : "—",
          r.bank > 0.0001 ? r3(r.bank) : "—", r.cut > 0.0001 ? r3(r.cut) : "—",
          r.topWidth ? r3(r.topWidth) : "—",
          r.fillArea > 0.0001 ? r3(r.fillArea) : "—", r.cutArea > 0.0001 ? r3(r.cutArea) : "—",
          r.fillVol > 0.0001 ? r3(r.fillVol) : "—", r.cutVol > 0.0001 ? r3(r.cutVol) : "—",
        ];
      });

      pdf.addPage("a4", "landscape");

      pdf.autoTable({
        startY: 18,
        head: [calcHeaders],
        body: calcBody,
        theme: "grid",
        margin: { left: 5, right: 5 },
        styles: {
          fontSize: 6.2,
          cellPadding: 1.2,
          halign: "center",
          valign: "middle",
          lineColor: [0, 0, 0],
          lineWidth: 0.15,
          textColor: [0, 0, 0],
          overflow: "linebreak",
        },
        headStyles: {
          fillColor: C.light,
          textColor: C.dark,
          fontStyle: "bold",
          fontSize: 6,
        },
        alternateRowStyles: { fillColor: [253, 253, 253] },
        didDrawPage: function (data) {
          // Header on each calc sheet page
          pdf.setFontSize(8); pdf.setFont("helvetica", "bold"); pdf.setTextColor(...C.dark);
          pdf.text("Detailed Calculation Sheet", 10, 10);
          pdf.setFontSize(6); pdf.setFont("helvetica", "normal"); pdf.setTextColor(...C.muted);
          pdf.text(`${state.project.name || "Project"} • Page ${pdf.internal.getNumberOfPages()}`, 287, 10, { align: "right" });
        },
        didParseCell: function (data) {
          if (data.section === 'body') {
            // Update progress
            if (data.row.index % 50 === 0 && progressBar) {
              progressBar.style.width = `${20 + ((data.row.index / calcBody.length) * 15)}%`;
            }
          }
        },
      });

      if (progressBar) progressBar.style.width = "35%";
    }

    // ===== 4. ROLL DIAGRAM =====
    if (options.rollDiagram) {
      if (window._isPdfExportCancelled) throw new Error("Export cancelled by user.");
      if (loadingText) loadingText.textContent = "Rendering roll diagrams...";

      const originalRollCanvas = els.rollDiagramCanvas;
      const originalSideCanvas = els.sideViewCanvas;
      const originalRows = state.calcRows;

      const CHUNK_LEN = 1000;
      const minOverallCh = exportCalcRows[0].chainage;
      const maxOverallCh = exportCalcRows[exportCalcRows.length - 1].chainage;
      let chunkIdx = 0;
      const totalChunks = Math.ceil((maxOverallCh - minOverallCh) / CHUNK_LEN);

      for (let chStart = minOverallCh; chStart <= maxOverallCh; chStart += CHUNK_LEN) {
        if (window._isPdfExportCancelled) throw new Error("Export cancelled by user.");
        const chunkRollRows = exportCalcRows.filter(r => r.chainage >= chStart && r.chainage <= chStart + CHUNK_LEN);
        if (chunkRollRows.length < 2) continue;

        // --- Roll diagram: own page ---
        pdf.addPage("a4", "landscape");
        drawPageHeaderFooter(pdf, `L-Section Roll Diagram (Ch ${r3(chStart)}m to ${r3(chStart + CHUNK_LEN)}m)`);

        const isFast = options.fastMode === true;
        const rdW = isFast ? 1200 : 2400;
        const rdH = isFast ? 300 : 600;
        const imgFmt = isFast ? "image/jpeg" : "image/png";
        const imgQual = isFast ? 0.6 : undefined;
        const pdfFmt = isFast ? "JPEG" : "PNG";

        const tempRollCanvas = document.createElement("canvas");
        tempRollCanvas.width = rdW; tempRollCanvas.height = rdH;
        els.rollDiagramCanvas = tempRollCanvas;
        state.calcRows = chunkRollRows;
        window._printModeLight = true;
        renderRollDiagram();

        if (tempRollCanvas.width > 0 && tempRollCanvas.height > 0) {
          try {
            const rollImg = tempRollCanvas.toDataURL(imgFmt, imgQual);
            const pgW = pdf.internal.pageSize.getWidth();
            const pgH = pdf.internal.pageSize.getHeight();
            const drawW = pgW - 16;
            const imgAspect = tempRollCanvas.width / tempRollCanvas.height;
            const drawH = drawW / imgAspect;
            const rollDrawH = Math.min(drawH, (pgH - 40) / 2);
            pdf.addImage(rollImg, pdfFmt, 8, 30, drawW, rollDrawH, undefined, "FAST");

            // --- Side view: same page below ---
            pdf.setFontSize(9); pdf.setFont("helvetica", "bold"); pdf.setTextColor(...C.dark);
            pdf.text("Cross-Sectional Toe Width Diagram", 8, 30 + rollDrawH + 6);

            const tempSideCanvas = document.createElement("canvas");
            tempSideCanvas.width = rdW; tempSideCanvas.height = rdH;
            els.sideViewCanvas = tempSideCanvas;
            renderSideView();

            if (tempSideCanvas.width > 0 && tempSideCanvas.height > 0) {
              const sideImg = tempSideCanvas.toDataURL(imgFmt, imgQual);
              const sideAspect = tempSideCanvas.width / tempSideCanvas.height;
              const sideDrawH = Math.min(drawW / sideAspect, (pgH - 40) / 2);
              pdf.addImage(sideImg, pdfFmt, 8, 30 + rollDrawH + 10, drawW, sideDrawH, undefined, "FAST");
            }
            tempSideCanvas.width = 0; tempSideCanvas.height = 0;
          } catch (e) { console.warn("Roll diagram image failed:", e); }
        }
        tempRollCanvas.width = 0; tempRollCanvas.height = 0;
        window._printModeLight = false;

        chunkIdx++;
        if (progressBar) progressBar.style.width = `${35 + ((chunkIdx / totalChunks) * 15)}%`;
        await new Promise(r => setTimeout(r, 10));
      }

      els.rollDiagramCanvas = originalRollCanvas;
      els.sideViewCanvas = originalSideCanvas;
      state.calcRows = originalRows;
    }

    if (progressBar) progressBar.style.width = "50%";

    // ===== 5. CROSS SECTIONS (1 per full landscape page) =====
    if (options.crossSections) {
      if (window._isPdfExportCancelled) throw new Error("Export cancelled by user.");
      if (loadingText) loadingText.textContent = "Rendering cross sections...";

      // Use the full viewBox with margins so RL summary on the right is included
      const fullVBW = CROSS_SVG_W + (CROSS_VIEW_MARGIN_X * 2); // 2540
      const fullVBH = CROSS_SVG_H + (CROSS_VIEW_MARGIN_Y * 2); // 1340

      for (let i = 0; i < exportCalcRows.length; i++) {
        if (window._isPdfExportCancelled) throw new Error("Export cancelled by user.");

        // One full landscape page per cross-section
        pdf.addPage("a4", "landscape");

        if (loadingText && i % 5 === 0) {
          loadingText.textContent = `Rendering cross section ${i + 1} / ${exportCalcRows.length}`;
        }
        if (progressBar && i % 5 === 0) {
          progressBar.style.width = `${50 + ((i / exportCalcRows.length) * 40)}%`;
        }

        const row = exportCalcRows[i];
        const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        svg.setAttribute("width", String(fullVBW));
        svg.setAttribute("height", String(fullVBH));
        // Include the margin area so RL summary text on the right side is captured
        svg.setAttribute("viewBox", `${-CROSS_VIEW_MARGIN_X} ${-CROSS_VIEW_MARGIN_Y} ${fullVBW} ${fullVBH}`);
        svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");

        const defs = els.crossSvg.querySelector("defs");
        if (defs) svg.appendChild(defs.cloneNode(true));
        drawCrossSection(row, svg);

        // Convert SVG to image via canvas at high resolution
        const svgStr = new XMLSerializer().serializeToString(svg);
        const svgBlob = new Blob([svgStr], { type: "image/svg+xml;charset=utf-8" });
        const svgUrl = URL.createObjectURL(svgBlob);

        try {
          const img = new Image();
          await new Promise((resolve, reject) => {
            img.onload = resolve;
            img.onerror = reject;
            img.src = svgUrl;
          });

          // Canvas matching SVG aspect ratio, adjusted for Fast Mode
          const isFast = options.fastMode === true;
          const cvsW = isFast ? 1270 : 2540;
          const cvsH = Math.round(cvsW * (fullVBH / fullVBW));
          const cvs = document.createElement("canvas");
          cvs.width = cvsW; cvs.height = cvsH;
          const ctx = cvs.getContext("2d");
          ctx.fillStyle = "#ffffff";
          ctx.fillRect(0, 0, cvsW, cvsH);
          ctx.drawImage(img, 0, 0, cvsW, cvsH);

          const jpegQuality = isFast ? 0.6 : 0.95;
          const imgData = cvs.toDataURL("image/jpeg", jpegQuality);

          const pgW = pdf.internal.pageSize.getWidth();  // 297mm landscape
          const pgH = pdf.internal.pageSize.getHeight(); // 210mm landscape

          // Header label
          pdf.setFontSize(10); pdf.setFont("helvetica", "bold"); pdf.setTextColor(...C.primary);
          pdf.text(`CROSS-SECTION`, 10, 10);
          pdf.setTextColor(...C.dark);
          pdf.text(`Chainage: ${r3(row.chainage)} m  •  Type: ${row.type}`, 52, 10);
          pdf.setFontSize(7); pdf.setFont("helvetica", "normal"); pdf.setTextColor(...C.muted);
          pdf.text(`${state.project.name || "Project"}`, pgW - 10, 10, { align: "right" });

          // Draw the image to fill the page, preserving aspect ratio
          const margin = 6;
          const availW = pgW - (margin * 2);
          const availH = pgH - 16 - margin; // 16mm top for header
          const imgAspect = cvsW / cvsH;
          let drawW = availW;
          let drawH = drawW / imgAspect;
          if (drawH > availH) {
            drawH = availH;
            drawW = drawH * imgAspect;
          }
          const offsetX = margin + (availW - drawW) / 2;
          pdf.addImage(imgData, "JPEG", offsetX, 14, drawW, drawH, undefined, "FAST");

          // Cleanup
          cvs.width = 0; cvs.height = 0;
          URL.revokeObjectURL(svgUrl);
        } catch (e) {
          console.warn(`Cross section ${i} render failed:`, e);
          URL.revokeObjectURL(svgUrl);
        }

        // Yield every 4 sections to keep browser responsive
        if (i % 4 === 0) await new Promise(r => setTimeout(r, 5));
      }
    }

    if (progressBar) progressBar.style.width = "92%";

    // ===== 6. SIGNATURE PAGE =====
    const signName = String(reportBrand.signName || "").trim();
    const signTitle = String(reportBrand.signTitle || "").trim();
    const signatureImg = reportBrand.signature;
    if (signatureImg || signName || signTitle) {
      pdf.addPage("a4", "portrait");
      drawPageHeaderFooter(pdf, "");
      const pgH = pdf.internal.pageSize.getHeight();
      const sigY = pgH - 80;

      pdf.setFontSize(8); pdf.setFont("helvetica", "bold"); pdf.setTextColor(...C.muted);
      pdf.text("SIGNATURE", 20, sigY);

      if (signatureImg) {
        try { pdf.addImage(signatureImg, "PNG", 20, sigY + 4, 60, 25); } catch (e) { /* skip */ }
      }
      pdf.setDrawColor(...C.border); pdf.setLineWidth(0.3);
      pdf.line(20, sigY + 32, 80, sigY + 32);

      pdf.setFontSize(11); pdf.setFont("helvetica", "bold"); pdf.setTextColor(...C.dark);
      pdf.text(signName || "Authorized Signatory", 20, sigY + 38);
      if (signTitle) {
        pdf.setFontSize(9); pdf.setFont("helvetica", "normal"); pdf.setTextColor(...C.text);
        pdf.text(signTitle, 20, sigY + 44);
      }
    }

    // ===== STAMP PAGE NUMBERS ON ALL PAGES =====
    if (loadingText) loadingText.textContent = "Adding page numbers...";
    const totalPages = pdf.internal.getNumberOfPages();
    for (let p = 1; p <= totalPages; p++) {
      pdf.setPage(p);
      const pgW = pdf.internal.pageSize.getWidth();
      const pgH = pdf.internal.pageSize.getHeight();
      pdf.setFontSize(7); pdf.setFont("helvetica", "normal"); pdf.setTextColor(148, 163, 184);
      pdf.text(`Page ${p} of ${totalPages}`, pgW - 10, pgH - 6, { align: "right" });
      pdf.text(`Earthsoft Professional • ${state.project.name || "Project"}`, 10, pgH - 6);
    }

    // ===== SAVE =====
    if (window._isPdfExportCancelled) throw new Error("Export cancelled by user.");
    if (progressBar) progressBar.style.width = "96%";
    if (loadingText) loadingText.textContent = "Saving PDF...";

    const filename = `${state.project.name || "Earthsoft_Report"}_${new Date().toISOString().split('T')[0]}.pdf`;
    pdf.save(filename);

    if (progressBar) progressBar.style.width = "100%";
    if (loadingText) loadingText.textContent = "Download complete!";
    await new Promise(r => setTimeout(r, 800));

  } catch (error) {
    if (error.message === "Export cancelled by user.") {
      console.log("PDF Export was cancelled by the user.");
      if (loadingText) loadingText.textContent = "Export cancelled.";
      return;
    }
    console.error("Report Generation Error:", error);
    alert("An error occurred during report generation: " + error.message);
  } finally {
    const wrapper = document.getElementById("pdf-export-wrapper");
    if (wrapper) wrapper.remove();

    const cancelBtn = document.getElementById("cancelPdfExportBtn");
    if (cancelBtn) cancelBtn.style.display = "none";

    if (loadingHeader) loadingHeader.textContent = "AI Processing...";
    if (progressBarContainer) progressBarContainer.style.display = "none";
    if (loading) loading.classList.add("hidden");
  }
}
