const r3 = (v) => (Number.isFinite(v) ? Number(v).toFixed(3) : "0.000");
const CROSS_SVG_W = 1700;
const CROSS_SVG_H = 980;

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
    },
  },
  activeWorkPage: "overview",
  activeResultTab: "inputs",
  charts: { lSection: null, volume: null },
  crossViewBox: { x: 0, y: 0, w: CROSS_SVG_W, h: CROSS_SVG_H },
  crossPan: { active: false, lastX: 0, lastY: 0 },
};

const els = {
  projectMeta: document.getElementById("projectMeta"),
  tableBody: document.getElementById("tableBody"),
  totalFilling: document.getElementById("totalFilling"),
  totalCutting: document.getElementById("totalCutting"),
  fillLength: document.getElementById("fillLength"),
  cutLength: document.getElementById("cutLength"),
  bridgeImportInput: document.getElementById("bridgeImportInput"),
  curveImportInput: document.getElementById("curveImportInput"),
  loopImportInput: document.getElementById("loopImportInput"),
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
  exportBtn: document.getElementById("exportBtn"),
  closePdfExportBtn: document.getElementById("closePdfExportBtn"),
  exportCalcPdfBtn: document.getElementById("exportCalcPdfBtn"),
  exportCrossPdfBtn: document.getElementById("exportCrossPdfBtn"),
  openSettingsBtn: document.getElementById("openSettingsBtn"),
  settingsModal: document.getElementById("settingsModal"),
  settingsForm: document.getElementById("settingsForm"),
  closeSettingsBtn: document.getElementById("closeSettingsBtn"),
  settingsGrid: document.getElementById("settingsGrid"),
  resetDefaultsBtn: document.getElementById("resetDefaultsBtn"),
  lSectionChart: document.getElementById("lSectionChart"),
  volumeChart: document.getElementById("volumeChart"),
  crossSectionModal: document.getElementById("crossSectionModal"),
  crossTitle: document.getElementById("crossTitle"),
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
  crossMeta: document.getElementById("crossMeta"),
  graphModal: document.getElementById("graphModal"),
  graphModalTitle: document.getElementById("graphModalTitle"),
  closeGraphModalBtn: document.getElementById("closeGraphModalBtn"),
  expandedGraphCanvas: document.getElementById("expandedGraphCanvas"),
  fillWaterNode: document.getElementById("fillWaterNode"),
  cutWaterNode: document.getElementById("cutWaterNode"),
};

const settingSchema = [
  ["formationWidthFill", "Formation Width (Fill) m"],
  ["cuttingWidth", "Formation Width (Cut) m"],
  ["blanketThickness", "Blanket Thickness m"],
  ["preparedSubgradeThickness", "Prepared Subgrade Thickness m"],
  ["bermWidth", "Berm Width m"],
  ["sideSlopeFactor", "Side Slope Factor (H:V)"],
  ["ballastCushionThickness", "Ballast Cushion Thickness m"],
  ["topLayerThickness", "Top Layer of Embankment Fill m"],
  ["activeSqCategory", "Active Soil Category (1=SQ1,2=SQ2,3=SQ3)"],
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
  return Number(t.replace(/[^\d.-]/g, ""));
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

function findColByAliases(headers, aliases) {
  return headers.findIndex((h) => {
    const n = normalizeHeaderToken(h);
    return aliases.some((a) => n === a || n.includes(a));
  });
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
    bridgeNo: findColByAliases(headerCells, ["bridgeno", "bridgenumber", "structureno", "structure", "bridge"]),
    startChainage: findColByAliases(headerCells, ["bridgestartchainage", "startchainage", "fromchainage", "chfrom", "startch", "from"]),
    endChainage: findColByAliases(headerCells, ["bridgeendchainage", "endchainage", "tochainage", "chto", "endch", "to"]),
    length: findColByAliases(headerCells, ["totalspanlength", "bridgelength", "spanlength", "length", "spanm"]),
    chainage: findColByAliases(headerCells, ["chainage", "ch", "centerchainage", "location"]),
  };
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
      + (cols.chainage >= 0 ? 1 : 0);
    if (score > best.score) {
      best = { rowIndex: i, score, cols };
    }
  }
  if (best.rowIndex < 0) return null;
  const c = best.cols;
  const hasCore = (c.startChainage >= 0 && c.endChainage >= 0)
    || (c.chainage >= 0 && c.length >= 0)
    || (c.startChainage >= 0 && c.length >= 0)
    || (c.endChainage >= 0 && c.length >= 0);
  return hasCore ? best : null;
}

function normalizeBridgeEntry(raw, index = 0) {
  const bridgeNo = String(raw.bridgeNo || `BR-${index + 1}`).trim();
  let start = parseChainage(raw.startChainage);
  let end = parseChainage(raw.endChainage);
  const lengthRaw = parseLooseNumber(raw.length, NaN);
  const center = parseChainage(raw.chainage);

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
  return {
    bridgeNo,
    startChainage: start,
    endChainage: end,
    length,
  };
}

function parseBridgeRowsFromAoa(aoa) {
  const header = detectBridgeHeaderRow(aoa);
  if (!header) return { rows: [], error: "Bridge sheet must include Start/End chainage or Chainage+Length columns." };
  const rows = [];
  const cols = header.cols;
  for (let i = header.rowIndex + 1; i < aoa.length; i += 1) {
    const row = Array.isArray(aoa[i]) ? aoa[i] : [];
    const raw = {
      bridgeNo: cols.bridgeNo >= 0 ? row[cols.bridgeNo] : "",
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
  const cCsb = norm.findIndex((h) => h === "csb");
  const cLoopStart = norm.findIndex((h) => h.includes("loopstartch") || h.includes("loopstartchain"));
  const cLoopEnd = norm.findIndex((h) => h.includes("loopendch") || h.includes("loopendchain"));
  const cTc = norm.findIndex((h) => h === "tc" || h.includes("trackcentre") || h.includes("trackcenter"));
  const cPfStart = norm.findIndex((h) => h.includes("pfstartch") || h.includes("pfstartchain"));
  const cPfEnd = norm.findIndex((h) => h.includes("pfendch") || h.includes("pfendchain"));
  const cWidth = norm.findIndex((h) => h === "width" || h.includes("platformwidth"));
  const cRemarks = norm.findIndex((h) => h.includes("remark"));

  // Fallback to fixed template columns from provided workbook:
  // B Station, C CSB, D Loop Start(off), E Loop End(off), F T/C, G Loop Start Ch, H Loop End Ch,
  // I PF Start(off), J PF End(off), K Width, L PF Start Ch, M PF End Ch, N Remarks
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
    loopStartOff: 3,
    loopEndOff: 4,
    pfStartOff: 8,
    pfEndOff: 9,
  };

  const rows = [];
  let currentStation = "";
  let currentCsb = NaN;
  for (let i = headerRow + 1; i < aoa.length; i += 1) {
    const row = Array.isArray(aoa[i]) ? aoa[i] : [];
    if (!row.length) continue;

    const stationRaw = String(row[idx.station] ?? "").trim();
    if (stationRaw && !/example|target chainage|add here/i.test(stationRaw)) currentStation = stationRaw;

    const csbCandidate = parseLooseNumber(row[idx.csb], NaN);
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
    if (!hasLoop && !hasPf) continue;

    rows.push({
      station: currentStation || `LP-${rows.length + 1}`,
      csb: Number.isFinite(currentCsb) ? currentCsb : null,
      tc: Number.isFinite(tc) ? tc : 0,
      loopStartCh: Number.isFinite(loopStartCh) ? loopStartCh : null,
      loopEndCh: Number.isFinite(loopEndCh) ? loopEndCh : null,
      pfWidth: Number.isFinite(width) ? width : 0,
      pfStartCh: Number.isFinite(pfStartCh) ? pfStartCh : null,
      pfEndCh: Number.isFinite(pfEndCh) ? pfEndCh : null,
      remarks,
    });
  }
  return { rows };
}

function getLoopPlatformAtChainage(chainage) {
  if (!Number.isFinite(chainage) || !state.loopPlatformRows.length) {
    return { loopTc: 0, platformWidth: 0 };
  }
  let loopTc = 0;
  let platformWidth = 0;
  for (const r of state.loopPlatformRows) {
    if (Number.isFinite(r.loopStartCh) && Number.isFinite(r.loopEndCh) && chainage >= r.loopStartCh && chainage <= r.loopEndCh) {
      loopTc += safeNum(r.tc, 0);
    }
    if (Number.isFinite(r.pfStartCh) && Number.isFinite(r.pfEndCh) && chainage >= r.pfStartCh && chainage <= r.pfEndCh) {
      platformWidth += safeNum(r.pfWidth, 0);
    }
  }
  return { loopTc, platformWidth };
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

function buildBridgeIntervals() {
  return state.bridgeRows
    .map((b, i) => normalizeBridgeEntry(b, i))
    .filter(Boolean)
    .sort((a, b) => a.startChainage - b.startChainage);
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
      overlaps.push([ovStart, ovEnd]);
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

function getBridgeRefsAtChainage(chainage, bridges) {
  if (!Number.isFinite(chainage) || !bridges.length) return [];
  return bridges
    .filter((b) => chainage >= b.startChainage && chainage <= b.endChainage)
    .map((b) => b.bridgeNo);
}

function renderBridgeInputs() {
  if (!els.bridgeTableBody) return;
  if (!state.bridgeRows.length) {
    els.bridgeTableBody.innerHTML = `
      <tr>
        <td colspan="5" class="muted">No bridge rows loaded. Use Import Data > Import Bridge List or Add Bridge Row.</td>
      </tr>
    `;
    if (els.bridgeMeta) els.bridgeMeta.textContent = "No bridge deduction applied.";
    return;
  }
  els.bridgeTableBody.innerHTML = state.bridgeRows.map((b, i) => `
    <tr data-bridge-row="${i}">
      <td><input data-bridge-field="bridgeNo" value="${String(b.bridgeNo || "")}" /></td>
      <td><input data-bridge-field="startChainage" value="${Number.isFinite(parseChainage(b.startChainage)) ? r3(parseChainage(b.startChainage)) : String(b.startChainage ?? "")}" /></td>
      <td><input data-bridge-field="endChainage" value="${Number.isFinite(parseChainage(b.endChainage)) ? r3(parseChainage(b.endChainage)) : String(b.endChainage ?? "")}" /></td>
      <td><input data-bridge-field="length" value="${Number.isFinite(parseLooseNumber(b.length, NaN)) ? r3(parseLooseNumber(b.length, NaN)) : String(b.length ?? "")}" /></td>
      <td><button type="button" class="bridge-del" data-bridge-del="${i}">Delete</button></td>
    </tr>
  `).join("");
  const validBridges = buildBridgeIntervals();
  const totalDeduction = validBridges.reduce((acc, b) => acc + safeNum(b.length), 0);
  if (els.bridgeMeta) {
    const invalidCount = Math.max(state.bridgeRows.length - validBridges.length, 0);
    els.bridgeMeta.textContent = `Bridge rows: ${state.bridgeRows.length} | Valid intervals: ${validBridges.length} | Total bridge length: ${r3(totalDeduction)} m${invalidCount ? ` | Invalid rows: ${invalidCount}` : ""}`;
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
    els.loopTableBody.innerHTML = `<tr><td colspan="10" class="muted">No station/loop rows loaded. Use Import Data > Import Loops & Platforms.</td></tr>`;
    if (els.loopMeta) els.loopMeta.textContent = "No station/loops data imported.";
    return;
  }
  els.loopTableBody.innerHTML = rows.map((r, i) => `
    <tr data-loop-row="${i}">
      <td><input data-loop-field="station" value="${String(r.station || "")}" /></td>
      <td><input data-loop-field="csb" value="${Number.isFinite(r.csb) ? r3(r.csb) : ""}" /></td>
      <td><input data-loop-field="tc" value="${Number.isFinite(r.tc) ? r3(r.tc) : ""}" /></td>
      <td><input data-loop-field="loopStartCh" value="${Number.isFinite(r.loopStartCh) ? r3(r.loopStartCh) : ""}" /></td>
      <td><input data-loop-field="loopEndCh" value="${Number.isFinite(r.loopEndCh) ? r3(r.loopEndCh) : ""}" /></td>
      <td><input data-loop-field="pfWidth" value="${Number.isFinite(r.pfWidth) ? r3(r.pfWidth) : ""}" /></td>
      <td><input data-loop-field="pfStartCh" value="${Number.isFinite(r.pfStartCh) ? r3(r.pfStartCh) : ""}" /></td>
      <td><input data-loop-field="pfEndCh" value="${Number.isFinite(r.pfEndCh) ? r3(r.pfEndCh) : ""}" /></td>
      <td><input data-loop-field="remarks" value="${String(r.remarks || "")}" /></td>
      <td><button type="button" class="bridge-del" data-loop-del="${i}">Delete</button></td>
    </tr>
  `).join("");
  const loopRanges = rows.filter((r) => Number.isFinite(r.loopStartCh) && Number.isFinite(r.loopEndCh)).length;
  const pfRanges = rows.filter((r) => Number.isFinite(r.pfStartCh) && Number.isFinite(r.pfEndCh)).length;
  if (els.loopMeta) els.loopMeta.textContent = `Rows: ${rows.length} | Loop ranges: ${loopRanges} | Platform ranges: ${pfRanges}`;
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
    const csb = parseLooseNumber(tr.querySelector('[data-loop-field="csb"]')?.value ?? "", NaN);
    const tc = parseLooseNumber(tr.querySelector('[data-loop-field="tc"]')?.value ?? "", NaN);
    const loopStartCh = parseChainage(tr.querySelector('[data-loop-field="loopStartCh"]')?.value ?? "");
    const loopEndCh = parseChainage(tr.querySelector('[data-loop-field="loopEndCh"]')?.value ?? "");
    const pfWidth = parseLooseNumber(tr.querySelector('[data-loop-field="pfWidth"]')?.value ?? "", NaN);
    const pfStartCh = parseChainage(tr.querySelector('[data-loop-field="pfStartCh"]')?.value ?? "");
    const pfEndCh = parseChainage(tr.querySelector('[data-loop-field="pfEndCh"]')?.value ?? "");
    const remarks = String(tr.querySelector('[data-loop-field="remarks"]')?.value || "").trim();
    if (!station
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
      csb: Number.isFinite(csb) ? csb : null,
      tc: Number.isFinite(tc) ? tc : 0,
      loopStartCh: Number.isFinite(loopStartCh) ? loopStartCh : null,
      loopEndCh: Number.isFinite(loopEndCh) ? loopEndCh : null,
      pfWidth: Number.isFinite(pfWidth) ? pfWidth : 0,
      pfStartCh: Number.isFinite(pfStartCh) ? pfStartCh : null,
      pfEndCh: Number.isFinite(pfEndCh) ? pfEndCh : null,
      remarks,
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
  const minW = CROSS_SVG_W * 0.08;
  const maxW = CROSS_SVG_W * 5.5;
  const minH = CROSS_SVG_H * 0.08;
  const maxH = CROSS_SVG_H * 5.5;
  const w = Math.min(maxW, Math.max(minW, vb.w));
  const h = Math.min(maxH, Math.max(minH, vb.h));
  const panPadX = w * 0.24;
  const panPadY = h * 0.24;
  const minX = -panPadX;
  const minY = -panPadY;
  const maxX = CROSS_SVG_W - w + panPadX;
  const maxY = CROSS_SVG_H - h + panPadY;
  const x = minX <= maxX ? clamp(vb.x, minX, maxX) : (minX + maxX) / 2;
  const y = minY <= maxY ? clamp(vb.y, minY, maxY) : (minY + maxY) / 2;
  state.crossViewBox = { x, y, w, h };
  if (els.crossSvg) {
    els.crossSvg.setAttribute("viewBox", `${state.crossViewBox.x} ${state.crossViewBox.y} ${state.crossViewBox.w} ${state.crossViewBox.h}`);
  }
}

function resetCrossView() {
  setCrossViewBox({ x: 0, y: 0, w: CROSS_SVG_W, h: CROSS_SVG_H });
}

function zoomCrossAt(clientX, clientY, factor) {
  if (!els.crossSvg || !Number.isFinite(factor) || factor <= 0) return;
  const rect = els.crossSvg.getBoundingClientRect();
  if (!rect.width || !rect.height) return;
  const rx = (clientX - rect.left) / rect.width;
  const ry = (clientY - rect.top) / rect.height;
  const curr = state.crossViewBox;
  const newW = curr.w * factor;
  const newH = curr.h * factor;
  const newX = curr.x + ((curr.w - newW) * rx);
  const newY = curr.y + ((curr.h - newH) * ry);
  setCrossViewBox({ x: newX, y: newY, w: newW, h: newH });
}

function wheelToZoomFactor(e) {
  const modeScale = e.deltaMode === 1 ? 16 : (e.deltaMode === 2 ? 320 : 1);
  const normalized = clamp((e.deltaY * modeScale) / 120, -8, 8);
  const sensitivity = e.ctrlKey ? 0.12 : 0.085;
  return Math.exp(normalized * sensitivity);
}

function bindCrossCanvasInteraction() {
  if (!els.crossSvg || !els.crossGraphicWrap) return;
  els.crossSvg.style.cursor = "grab";
  els.crossGraphicWrap.addEventListener("wheel", (e) => {
    e.preventDefault();
    const factor = wheelToZoomFactor(e);
    zoomCrossAt(e.clientX, e.clientY, factor);
  }, { passive: false });
  els.crossSvg.addEventListener("mousedown", (e) => {
    if (e.button !== 0 && e.button !== 1) return;
    e.preventDefault();
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
    const platformWidth = safeNum(lp.platformWidth, 0);
    const effectiveFormationWidth = settings.formationWidthFill + loopTc + platformWidth;

    const bank = diff <= 0 ? 0 : Math.max(rlDiff, 0);
    const minCover = settings.blanketThickness + settings.preparedSubgradeThickness;
    const cut = diff <= 0 ? 0 : (rlDiff < 0 ? (-rlDiff + minCover) : Math.max(minCover - bank, 0));

    const turfing = bank > 0 ? settings.sideSlopeFactor * bank : 0;
    const fillTop = effectiveFormationWidth;
    const fillBottom = fillTop + (2 * settings.sideSlopeFactor * bank);
    const cutTop = settings.cuttingWidth;
    const cutBottom = cutTop + (2 * settings.sideSlopeFactor * cut);

    const fillArea = bank > 0 ? ((fillTop + fillBottom) * 0.5 * bank) : 0;
    const cutArea = cut > 0 ? ((cutTop + cutBottom) * 0.5 * cut) : 0;

    const fillVol = prev ? ((prev.fillArea + fillArea) * 0.5 * ewDiff) : 0;
    const cutVol = prev ? ((prev.cutArea + cutArea) * 0.5 * ewDiff) : 0;

    let type = "NEUTRAL";
    if (ewDiff <= 0.001 && bridgeDeductLen > 0) type = "BRIDGE";
    else if (bank > 0.001) type = "FILLING";
    else if (cut > 0.001) type = "CUTTING";

    const calc = {
      ...row,
      diff,
      ewDiff,
      rlDiff,
      bank,
      cut,
      turfing,
      fillArea,
      cutArea,
      fillVol,
      cutVol,
      loopTc,
      platformWidth,
      effectiveFormationWidth,
      bridgeRefs,
      bridgeDeductLen,
      type,
      fillBottom,
      cutBottom,
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
  updateEstimates();
}

function renderSummary() {
  const fillTotal = state.calcRows.reduce((s, r) => s + r.fillVol, 0);
  const cutTotal = state.calcRows.reduce((s, r) => s + r.cutVol, 0);
  const fillLen = state.calcRows.reduce((s, r) => s + ((r.bank > 0 ? safeNum(r.ewDiff) : 0)), 0) / 1000;
  const cutLen = state.calcRows.reduce((s, r) => s + ((r.type === "CUTTING" ? safeNum(r.ewDiff) : 0)), 0) / 1000;

  els.totalFilling.textContent = `${r3(fillTotal)} m³`;
  els.totalCutting.textContent = `${r3(cutTotal)} m³`;
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
    if (els.cutWaterNode) els.cutWaterNode.style.height = `${Math.min(cutPct, 100)}%`;
    if (fillCapEl) fillCapEl.textContent = `${fillPct}% capacity`;
    if (cutCapEl) cutCapEl.textContent = `${cutPct}% capacity`;
  } else {
    if (els.fillWaterNode) els.fillWaterNode.style.height = "0%";
    if (els.cutWaterNode) els.cutWaterNode.style.height = "0%";
    if (fillCapEl) fillCapEl.textContent = "0% capacity";
    if (cutCapEl) cutCapEl.textContent = "0% capacity";
  }
  renderFormulaSummary();

  // Enable/disable the global verify button
  const vBtn = document.getElementById("verifyCalcBtn");
  if (vBtn) vBtn.disabled = !state.calcRows.length;
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
  };
  Object.entries(tickMap).forEach(([k, el]) => {
    if (!el) return;
    el.classList.toggle("done", Boolean(state.project.uploads[k]));
  });
  const ready = isProjectReadyForVerification();
  if (els.wizardCalculateBtn) els.wizardCalculateBtn.disabled = !ready;
  if (els.wizardStatus) {
    if (ready) {
      els.wizardStatus.textContent = "All files uploaded. Project is verified and ready for calculation.";
    } else {
      const missing = [];
      if (!state.project.name) missing.push("Project Name");
      if (!state.project.uploads.levels) missing.push("Levels");
      if (!state.project.uploads.curves) missing.push("Curve List");
      if (!state.project.uploads.bridges) missing.push("Bridge List");
      if (!state.project.uploads.loops) missing.push("Loops & Platforms");
      els.wizardStatus.textContent = `Pending: ${missing.join(", ")}`;
    }
  }
}

function applyProjectGate() {
  const active = Boolean(state.project.active);
  if (els.importBtn) els.importBtn.disabled = !active;
  if (els.exportBtn) els.exportBtn.disabled = !state.project.verified;
  if (els.openSettingsBtn) els.openSettingsBtn.disabled = !active;
  if (els.saveProjectBtn) els.saveProjectBtn.disabled = !active;
  const projectTitle = state.project.name ? `${state.project.name}` : "No active project";
  const status = active ? (state.project.verified ? "Verified" : "Draft") : "Create Project to begin";
  if (els.projectMeta) {
    const source = state.meta?.projectName ? ` | Base: ${state.meta.projectName}` : "";
    els.projectMeta.textContent = `${projectTitle} | ${status}${source}`;
  }
}

function collectProjectPayload() {
  return {
    version: "1.0",
    savedAt: new Date().toISOString(),
    project: {
      ...state.project,
      uploads: { ...state.project.uploads },
    },
    settings: { ...state.settings },
    rawRows: state.rawRows,
    bridgeRows: state.bridgeRows,
    curveRows: state.curveRows,
    loopPlatformRows: state.loopPlatformRows,
  };
}

function loadProjectFromPayload(data, options = {}) {
  const { silent = false } = options;
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
    },
  };
  state.settings = { ...state.seedDefaultSettings, ...(data?.settings || {}) };
  state.rawRows = Array.isArray(data?.rawRows) && data.rawRows.length ? data.rawRows : [...state.seedRows];
  state.bridgeRows = Array.isArray(data?.bridgeRows) ? data.bridgeRows : [];
  state.curveRows = Array.isArray(data?.curveRows) ? data.curveRows : [];
  const rawLoopRows = Array.isArray(data?.loopPlatformRows) ? data.loopPlatformRows : [];
  state.loopPlatformRows = rawLoopRows.map((r, i) => ({
    station: String(r?.station || r?.name || `LP-${i + 1}`),
    csb: Number.isFinite(parseLooseNumber(r?.csb, NaN)) ? parseLooseNumber(r.csb, NaN) : null,
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
  }));
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
    if ("showSaveFilePicker" in window) {
      const handle = await window.showSaveFilePicker({
        suggestedName: `${fileBase}.EW`,
        types: [{
          description: "Earthwork Project File",
          accept: { "application/json": [".EW"] },
        }],
      });
      const writable = await handle.createWritable();
      await writable.write(jsonStr);
      await writable.close();
      alert(`Project saved: ${state.project.name}`);
    } else {
      // Fallback for unsupported browsers
      const blob = new Blob([jsonStr], { type: "application/json" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `${fileBase}.EW`;
      a.click();
      URL.revokeObjectURL(a.href);
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
  state.project = {
    active: false,
    verified: false,
    name: "",
    uploads: {
      levels: false,
      curves: false,
      bridges: false,
      loops: false,
    },
  };
  state.rawRows = [];
  state.bridgeRows = [];
  state.curveRows = [];
  state.loopPlatformRows = [];
  state.settings = { ...state.seedDefaultSettings };
  if (els.projectNameInput) els.projectNameInput.value = "";
  if (els.importInput) els.importInput.value = "";
  if (els.bridgeImportInput) els.bridgeImportInput.value = "";
  if (els.curveImportInput) els.curveImportInput.value = "";
  if (els.loopImportInput) els.loopImportInput.value = "";
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

async function generateNativePdf(title, bodyHtml, orientation = "portrait") {
  if (typeof html2pdf === "undefined") {
    alert("PDF generation library failed to load. Please check your connection or try again.");
    return;
  }

  const container = document.createElement("div");
  container.innerHTML = `
    <div style="font-family: Arial, sans-serif; padding: 10px; color: #173045; background: #fff;">
      <style>
        h1 { margin: 0 0 6px; font-size: 20px; }
        h2 { margin: 16px 0 8px; font-size: 15px; }
        p { margin: 0 0 6px; font-size: 12px; color: #4a647d; }
        table { width: 100%; border-collapse: collapse; font-size: 10px; }
        th, td { border: 1px solid #c9dced; padding: 4px; text-align: right; }
        th:first-child, td:first-child { text-align: left; }
        thead th { background-color: #f0f7ff; }
        .card { page-break-inside: avoid; border: 1px solid #c9dced; border-radius: 8px; padding: 8px; margin-bottom: 20px; }
        .meta { font-size: 11px; margin-top: 6px; line-height: 1.45; }
        svg { width: 100%; height: 170px; border: 1px solid #d9e8f7; border-radius: 6px; background: #f7fbff; }
        .page-break { page-break-before: always; }
      </style>
      ${bodyHtml}
    </div>
  `;

  // We don't attach the container to the document body because html2pdf can render offscreen nodes.
  const fileBase = String(state.project.name || "project").replace(/[^\w.-]+/g, "_").replace(/^_+|_+$/g, "") || "project";
  const filename = `${fileBase}_${title.replace(/\s+/g, "_")}.pdf`;

  const opt = {
    margin: 10,
    filename: filename,
    image: { type: 'jpeg', quality: 0.98 },
    html2canvas: { scale: 2, useCORS: true },
    jsPDF: { unit: 'mm', format: 'a4', orientation: orientation }
  };

  try {
    els.exportBtn.textContent = "Generating PDF...";
    els.exportBtn.disabled = true;

    await html2pdf().set(opt).from(container).save();

    els.exportBtn.textContent = "Export PDF";
    els.exportBtn.disabled = false;
  } catch (err) {
    console.error("PDF generation failed:", err);
    alert("Failed to generate PDF. See console for details.");
    els.exportBtn.textContent = "Export PDF";
    els.exportBtn.disabled = false;
  }
}

function exportCalculationPdf() {
  if (!state.calcRows.length) {
    alert("No calculation rows available to export.");
    return;
  }
  const head = ["Structures", "Station", "Chainage", "Diff", "Ground RL", "Proposed RL", "Loop T/C", "PF Width", "Bridge No", "Deducted Len", "EW Diff", "RL Diff", "Bank", "Cut", "Fill Area", "Cut Area", "Fill Vol", "Cut Vol"];
  const rows = state.calcRows.map((r) => [
    (r.structureNo || "-"),
    (r.station || "-"),
    r3(r.chainage), r3(r.diff), r3(r.groundLevel), r3(r.proposedLevel), r3(r.loopTc), r3(r.platformWidth),
    (r.bridgeRefs && r.bridgeRefs.length) ? r.bridgeRefs.join(" | ") : "-",
    r3(r.bridgeDeductLen), r3(r.ewDiff), r3(r.rlDiff), r3(r.bank), r3(r.cut), r3(r.fillArea), r3(r.cutArea), r3(r.fillVol), r3(r.cutVol),
  ]);
  const html = `
    <h1>${escapeHtml(state.project.name || "Railway Earthwork Project")}</h1>
    <p>Calculation Sheet Export | Generated: ${escapeHtml(new Date().toLocaleString())}</p>
    <table>
      <thead><tr>${head.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>
      <tbody>
        ${rows.map((r) => `<tr>${r.map((c) => `<td>${escapeHtml(c)}</td>`).join("")}</tr>`).join("")}
      </tbody>
    </table>
  `;
  generateNativePdf("Calculation Sheet", html, "landscape");
}

function buildCrossSectionMiniSvg(row) {
  const bank = Math.max(safeNum(row.bank), 0);
  const cut = Math.max(safeNum(row.cut), 0);
  const topY = 56;
  const baseY = 140;
  const h = Math.min(70, Math.max(18, bank * 7));
  const leftTop = 230;
  const rightTop = 430;
  const leftToe = 90;
  const rightToe = 570;
  const crownY = topY;
  const toeY = baseY;
  const embankment = bank > 0
    ? `<polygon points="${leftTop},${crownY} ${rightTop},${crownY} ${rightToe},${toeY} ${leftToe},${toeY}" fill="#dfe9de" stroke="#738878" />`
    : `<polygon points="${leftTop},${crownY + 28} ${rightTop},${crownY + 28} ${rightToe},${toeY} ${leftToe},${toeY}" fill="#f2e5e6" stroke="#9c8082" />`;
  const layerTop = crownY - 10;
  return `
    <svg viewBox="0 0 660 180" role="img" aria-label="Cross Section">
      <line x1="50" y1="${toeY}" x2="610" y2="${toeY}" stroke="#5d6b77" stroke-dasharray="8 6" />
      <text x="615" y="${toeY - 4}" font-size="11" fill="#3b4d5b">G.L.</text>
      ${embankment}
      <rect x="${leftTop}" y="${layerTop}" width="${rightTop - leftTop}" height="10" fill="#ccd3d9" stroke="#7f8a95" />
      <text x="${rightTop + 8}" y="${layerTop + 8}" font-size="11" fill="#2f4d6a">Ballast 0.350 m</text>
      <text x="260" y="${crownY + 24}" font-size="12" fill="#2d435a" font-weight="700">${bank > 0 ? "Embankment Fill" : "Cutting Section"}</text>
      <text x="260" y="${crownY + 40}" font-size="11" fill="#2d435a">Bank: ${r3(bank)} m | Cut: ${r3(cut)} m</text>
    </svg>
  `;
}

function exportCrossSectionsPdf() {
  if (!state.calcRows.length) {
    alert("No cross-sections available to export.");
    return;
  }
  const cards = state.calcRows.map((r) => `
    <div class="card">
      <h2>CH ${r3(r.chainage)} m</h2>
      ${buildCrossSectionMiniSvg(r)}
      <div class="meta">
        Ground RL: ${r3(r.groundLevel)} m | Proposed RL: ${r3(r.proposedLevel)} m | RL Diff: ${r3(r.rlDiff)} m<br/>
        Loop T/C: ${r3(r.loopTc)} m | PF Width: ${r3(r.platformWidth)} m | Bridge Deduct: ${r3(r.bridgeDeductLen)} m
      </div>
    </div>
  `).join("");
  const html = `
    <h1>${escapeHtml(state.project.name || "Railway Earthwork Project")}</h1>
    <p>Cross-Sections Export | Generated: ${escapeHtml(new Date().toLocaleString())}</p>
    ${cards}
  `;
  generateNativePdf("Cross Sections", html, "portrait");
}

function renderFormulaSummary() {
  if (!els.resultInputBody || !els.resultFillBody || !els.resultCutBody || !els.resultQtyBody) return;
  const rows = state.calcRows;
  const s = state.settings;
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
  const fmtM3 = (v) => r3(v);
  const fmtKm = (m) => r3(m / 1000);
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
    ["Formation Width (Fill)", r3(s.formationWidthFill), "m"],
    ["Formation Width (Cut)", r3(s.cuttingWidth), "m"],
    ["Blanket Thickness", r3(s.blanketThickness), "m"],
    ["Prepared Subgrade Thickness", r3(s.preparedSubgradeThickness), "m"],
    ["Berm Width", r3(s.bermWidth), "m"],
    ["Track by side of ExgTrack", Number.isFinite(exgTrack) ? r3(exgTrack) : "-", Number.isFinite(exgTrack) ? "m" : ""],
    ["Total Length", r3(totalLength), "m"],
    ["Bridge Deduction Length", r3(totalBridgeLen), "m"],
    ["Bridge Count", String(validBridges.length), ""],
    ["Loop/PF Rows", String(loopRows.length), ""],
    ["Total Loop Range", r3(totalLoopLen), "m"],
    ["Total Platform Range", r3(totalPfLen), "m"],
  ];
  const fillRows = [
    ["Fill Length", fmtKm(fillLenM), "km"],
    ["Fill Ht <5.0m", fmtKm(fillLt5M), "km"],
    ["Fill Ht 5.0m to 10.0m", fmtKm(fill5To10M), "km"],
    ["Fill Ht >10.0m", fmtKm(fillGt10M), "km"],
    ["Max Fill Ht", r3(max(fillValues)), "m"],
    ["Average Fill Height", r3(avg(fillValues)), "m"],
  ];
  const cutRows = [
    ["Cut Length", fmtKm(cutLenM), "km"],
    ["Cut depth <5.0m", fmtKm(cutLt5M), "km"],
    ["Cut 5.0m to 10.0m", fmtKm(cut5To10M), "km"],
    ["Cut >10.0m", fmtKm(cutGt10M), "km"],
    ["Max Cut depth", r3(max(cutValues)), "m"],
    ["Min Cover (Blanket+Subgrade)", r3(minCover), "m"],
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

  const renderRows = (list) => list.map(([label, value, unit]) => `
    <tr>
      <td>${label}</td>
      <td>${value}</td>
      <td class="unit">${unit}</td>
    </tr>
  `).join("");
  els.resultInputBody.innerHTML = renderRows(inputRows);
  els.resultFillBody.innerHTML = renderRows(fillRows);
  els.resultCutBody.innerHTML = renderRows(cutRows);
  els.resultQtyBody.innerHTML = `
    ${qtyRows.map((r) => `
      <tr>
        <td>${r.label}</td>
        <td>${fmtM3(r.prepared)}</td>
        <td>${fmtM3(r.blanket)}</td>
        <td>${fmtM3(r.fill)}</td>
        <td>${fmtM3(r.cut)}</td>
      </tr>
    `).join("")}
    <tr>
      <td><strong>Total qty</strong></td>
      <td><strong>${fmtM3(qtyTotal.prepared)}</strong></td>
      <td><strong>${fmtM3(qtyTotal.blanket)}</strong></td>
      <td><strong>${fmtM3(qtyTotal.fill)}</strong></td>
      <td><strong>${fmtM3(qtyTotal.cut)}</strong></td>
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
    const station = r.station ? String(r.station).replace(/\n/g, " ") : "-";
    const bridgeRefs = (r.bridgeRefs && r.bridgeRefs.length) ? r.bridgeRefs.join(" | ") : "-";
    const rowClass = (r.bridgeRefs && r.bridgeRefs.length) ? "bridge-row" : "";
    return `
      <tr class="${rowClass}">
        <td>${structureNo}</td>
        <td>${station}</td>
        <td><button class="chainage-link" data-cross-index="${idx}" title="Open cross-section">${r3(r.chainage)}</button></td>
        <td>${r3(r.diff)}</td>
        <td>${r3(r.groundLevel)}</td>
        <td>${r3(r.proposedLevel)}</td>
        <td>${r3(r.loopTc)}</td>
        <td>${r3(r.platformWidth)}</td>
        <td>${bridgeRefs}</td>
        <td>${r3(r.bridgeDeductLen)}</td>
        <td>${r3(r.ewDiff)}</td>
        <td>${r3(r.rlDiff)}</td>
        <td>${r3(r.bank)}</td>
        <td>${r3(r.cut)}</td>
        <td>${r3(r.fillArea)}</td>
        <td>${r3(r.cutArea)}</td>
        <td>${volumeCapsule(r.fillVol, "fill", r.type === "FILLING" || r.fillVol > 0)}</td>
        <td>${volumeCapsule(r.cutVol, "cut", r.type === "CUTTING" || r.cutVol > 0)}</td>
      </tr>
    `;
  }).join("");

  els.tableBody.innerHTML = html;
}

function renderCharts() {
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
          pointRadius: 3,
          pointBackgroundColor: groundColor,
          pointBorderWidth: 0,
          pointHoverRadius: 5,
          pointHoverBackgroundColor: "hsl(222, 47%, 6%)",
          pointHoverBorderColor: groundColor,
          pointHoverBorderWidth: 2,
          borderWidth: 2,
        },
        {
          label: "Proposed RL",
          data: fl,
          borderColor: proposedColor,
          backgroundColor: proposedGrad,
          tension: 0.35,
          fill: true,
          pointRadius: 3,
          pointBackgroundColor: proposedColor,
          pointBorderWidth: 0,
          pointHoverRadius: 5,
          pointHoverBackgroundColor: "hsl(222, 47%, 6%)",
          pointHoverBorderColor: proposedColor,
          pointHoverBorderWidth: 2,
          borderWidth: 2,
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

function buildSettingsInputs() {
  els.settingsGrid.innerHTML = settingSchema.map(([key, label]) => `
    <label>
      ${label}
      <input type="number" step="0.001" name="${key}" value="${r3(state.settings[key])}" required />
    </label>
  `).join("");
}

function applySettingsFromForm() {
  const fd = new FormData(els.settingsForm);
  for (const [k] of settingSchema) {
    state.settings[k] = safeNum(fd.get(k), state.settings[k]);
  }
  recalculate();
}

function drawCrossSection(row) {
  const s = state.settings;
  const fl = row.proposedLevel;
  const gl = row.groundLevel;

  const sqCategory = Math.min(3, Math.max(1, Math.round(s.activeSqCategory || 3)));
  const sqName = sqCategory === 1 ? "SQ1" : (sqCategory === 2 ? "SQ2" : "SQ3");
  const blanketBySq = { 1: 1.0, 2: 0.75, 3: 0.6 };
  const blanketRuleThickness = blanketBySq[sqCategory];
  const ballastThickness = Math.max(0, s.ballastCushionThickness || 0.35);
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
    <tr><th>Formation Width (Top)</th><td>${r3(s.formationWidthFill)} m</td></tr>
    <tr><th>Berm Width (Each Berm)</th><td>${r3(s.bermWidth)} m (${Math.round(s.bermWidth * 1000)} mm)</td></tr>
    <tr><th>Berms per Side (Drawing)</th><td>${row.bank >= 8 ? 2 : (row.bank >= 4 ? 1 : 0)}</td></tr>
    <tr><th>Bottom Width (Fill)</th><td>${r3(row.fillBottom)} m</td></tr>
    <tr><th>Bottom Width (Cut)</th><td>${r3(row.cutBottom)} m</td></tr>
    <tr><th>Embankment Height</th><td>${r3(row.bank)} m</td></tr>
    <tr><th>Cut Depth</th><td>${r3(row.cut)} m</td></tr>
    <tr><th>Side Slope Factor</th><td>${r3(s.sideSlopeFactor)}</td></tr>
    <tr><th>Blanket Rule</th><td>SQ3=0.600 m, SQ2=0.750 m, SQ1=1.000 m</td></tr>
  `;

  els.crossTitle.textContent = `Cross-Section @ CH ${r3(row.chainage)} m`;
  els.crossMeta.textContent = `Type: ${row.type} | Ground RL: ${r3(gl)} | Proposed RL: ${r3(fl)} | Bank: ${r3(row.bank)} | Cut: ${r3(row.cut)}`;

  const svgW = CROSS_SVG_W;
  const svgH = CROSS_SVG_H;
  const marginX = 90;
  const centerX = svgW / 2;
  const layerTotalM = ballastThickness + blanketRuleThickness + topLayerThickness;
  const halfTopM = s.formationWidthFill / 2;
  const slopeHV = 2; // 2:1 slopes per requirement
  const fillRunM = row.bank * slopeHV;
  const cutRunM = row.cut * slopeHV;
  const halfBottomM = Math.max(halfTopM + fillRunM, row.fillBottom / 2, halfTopM);
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
  const blanketHDraw = row.bank > 0 ? Math.min(blanketH, fillH) : blanketH;
  const topLayerHDraw = row.bank > 0 ? Math.min(topLayerH, Math.max(fillH - blanketHDraw, 0)) : topLayerH;
  const trackTopY = topY - ballastH;
  const layerCalloutStyles = {
    ballast: { stroke: "#7f8a95", text: "#6f7a86" },
    blanket: { stroke: "#90886c", text: "#7f775b" },
    topLayer: { stroke: "#9a856c", text: "#876f57" },
  };

  const layerRects = [];
  // crowned ballast cushion with shoulders 0.60m each, 1:30 crossfall
  const shoulderWm = 0.6;
  const crownWm = Math.max(s.formationWidthFill - 2 * shoulderWm, 0.5);
  const shoulderPx = shoulderWm * pxPerM;
  const crownPx = crownWm * pxPerM;
  const dropPxRaw = (shoulderWm / 30) * pxPerM;
  const dropPx = Math.max(dropPxRaw, 2); // ensure visible crown drop
  const bx0 = centerX - halfTop;
  const bx1 = bx0 + shoulderPx;
  const bx2 = bx1 + crownPx;
  const bx3 = bx2 + shoulderPx;
  const byCenter = trackTopY;            // highest at center
  const byEdge = trackTopY + dropPx;     // lower at shoulders
  // ballast crown
  // ballast crown center (non-hatch), shoulders hatch, continuous fall outward
  const centerFill = "#cbd2d8";
  const shoulderLeft = `<polygon points="${bx0},${byEdge} ${bx1},${byCenter} ${bx1},${byCenter + ballastH} ${bx0},${byEdge + ballastH}" fill="url(#hatchFill)" stroke="#7f8a95" />`;
  const shoulderRight = `<polygon points="${bx2},${byCenter} ${bx3},${byEdge} ${bx3},${byEdge + ballastH} ${bx2},${byCenter + ballastH}" fill="url(#hatchFill)" stroke="#7f8a95" />`;
  const centerBallast = `<rect x="${bx1}" y="${byCenter}" width="${crownPx}" height="${ballastH}" fill="${centerFill}" stroke="#7f8a95" />`;
  layerRects.push(`${shoulderLeft}${centerBallast}${shoulderRight}`);
  // blanket follows same crossfall at ballast bottom (left/right hatch, center solid)
  const bltTopY0 = byEdge + ballastH;
  const bltTopY1 = byCenter + ballastH;
  const blanketLeft = `<polygon points="${bx0},${bltTopY0} ${bx1},${bltTopY1} ${bx1},${bltTopY1 + blanketHDraw} ${bx0},${bltTopY0 + blanketHDraw}" fill="#f5eecf" stroke="#90886c" />`;
  const blanketRight = `<polygon points="${bx2},${bltTopY1} ${bx3},${bltTopY0} ${bx3},${bltTopY0 + blanketHDraw} ${bx2},${bltTopY1 + blanketHDraw}" fill="#f5eecf" stroke="#90886c" />`;
  const blanketCenter = `<rect x="${bx1}" y="${bltTopY1}" width="${crownPx}" height="${blanketHDraw}" fill="#f5eecf" stroke="#90886c" />`;
  layerRects.push(`${blanketLeft}${blanketCenter}${blanketRight}`);
  // top layer with same pattern
  const tlTopY0 = bltTopY0 + blanketHDraw;
  const tlTopY1 = bltTopY1 + blanketHDraw;
  const topLayerLeft = `<polygon points="${bx0},${tlTopY0} ${bx1},${tlTopY1} ${bx1},${tlTopY1 + topLayerHDraw} ${bx0},${tlTopY0 + topLayerHDraw}" fill="#edd6bf" stroke="#9a856c" />`;
  const topLayerRight = `<polygon points="${bx2},${tlTopY1} ${bx3},${tlTopY0} ${bx3},${tlTopY0 + topLayerHDraw} ${bx2},${tlTopY1 + topLayerHDraw}" fill="#edd6bf" stroke="#9a856c" />`;
  const topLayerCenter = `<rect x="${bx1}" y="${tlTopY1}" width="${crownPx}" height="${topLayerHDraw}" fill="#edd6bf" stroke="#9a856c" />`;
  layerRects.push(`${topLayerLeft}${topLayerCenter}${topLayerRight}`);

  if (row.bank > 0) {
    const halfBlanketBottom = halfTop + (blanketHDraw * s.sideSlopeFactor);
    const halfTopLayerBottom = halfTop + ((blanketHDraw + topLayerHDraw) * s.sideSlopeFactor);
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

  const bermLabel = `${Math.round(s.bermWidth * 1000)} mm (${r3(s.bermWidth)} m)`;
  const cutPoly = row.cut > 0
    ? `<polygon points="${leftCutX},${cutBottomY} ${rightCutX},${cutBottomY} ${centerX + halfTop},${topY} ${centerX - halfTop},${topY}" fill="#f0e2e2" stroke="#8a6f72" />`
    : "";
  let fillPoly = "";
  let berms = "";
  if (row.bank > 0) {
    const bermCount = row.bank >= 8 ? 2 : (row.bank >= 4 ? 1 : 0);
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
    const polyPts = [...leftPts, ...rightPts.reverse()].map((p) => `${p.x},${p.y}`).join(" ");
    fillPoly = `<polygon points="${polyPts}" fill="url(#embFillHatch)" stroke="#69786a" />`;
    berms = bermDimSnippets.join("");
  }

  const segLeftOuter = 1.2;
  const segLeftShoulder = 0.35;
  const segRightShoulder = 0.35;
  const segRightOuter = 1.2;
  const segMiddle = Math.max(s.formationWidthFill - (segLeftOuter + segLeftShoulder + segRightShoulder + segRightOuter), 0.5);
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
  const bodyLabel = row.type === "CUTTING" ? "Cutting Section" : "Embankment Fill";
  const bodySubLabel = row.type === "CUTTING" ? "(Earth cutting profile)" : "(SQ1/SQ2/SQ3 Category Soils)";
  const bodyYRef = row.type === "CUTTING" ? ((topY + cutBottomY) / 2) : ((topY + toeY) / 2);
  const topLayerCalloutY = row.bank > 0 ? (topY + blanketHDraw) : (topY + blanketH);
  const calloutAnchorX = centerX + halfTop + 8;
  const calloutEndX = calloutAnchorX + 232;
  const calloutTextX = calloutEndX + 8;
  const calloutDy = -34;
  const ballastCalloutStartY = trackTopY + Math.max(ballastH * 0.48, 2);
  const blanketCalloutStartY = topY + Math.max(blanketHDraw * 0.62, 2);
  const topLayerCalloutStartY = topLayerCalloutY + Math.max(topLayerHDraw * 0.62, 2);
  const drawLayerCallout = (label, startY, style) => {
    const endY = startY + calloutDy;
    return `
      <line x1="${calloutAnchorX}" y1="${startY}" x2="${calloutEndX}" y2="${endY}" stroke="${style.stroke}" stroke-width="1.8" />
      <text x="${calloutTextX}" y="${endY - 2}" fill="${style.text}" font-size="12">${label}</text>
    `;
  };
  const layerCallouts = [
    drawLayerCallout(`${r3(ballastThickness)} m Ballast Cushion`, ballastCalloutStartY, layerCalloutStyles.ballast),
    drawLayerCallout(`Blanket (${sqName}) ${r3(blanketRuleThickness)} m`, blanketCalloutStartY, layerCalloutStyles.blanket),
    drawLayerCallout(`Top Layer of Embankment Fill ${r3(topLayerEffectiveThickness)} m`, topLayerCalloutStartY, layerCalloutStyles.topLayer),
  ].join("");
  const slopeLabels = row.bank > 0
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
    const y = topY + 40 + i * 20;
    const x = svgW - 240;
    return `<text x="${x}" y="${y}" fill="#2b3f52" font-size="12" font-weight="700">${txt}</text>`;
  }).join("");
  const demarcRightX = svgW - 210;
  const glLeftX = centerX - 520;
  const hflLeftX = centerX - 410;

  els.crossSvg.innerHTML = `
    <rect x="0" y="0" width="${svgW}" height="${svgH}" fill="#f8fcff" />
    <line x1="${glLeftX}" y1="${groundY}" x2="${demarcRightX}" y2="${groundY}" stroke="#5d6b77" stroke-dasharray="8 7" stroke-width="1.8" />
    <text x="${demarcRightX + 18}" y="${groundY - 7}" fill="#374b5d" font-size="13" font-weight="700">G.L.</text>
    <line x1="${hflLeftX}" y1="${hflY}" x2="${demarcRightX}" y2="${hflY}" stroke="#6d7680" stroke-dasharray="12 8" stroke-width="1.4" />
    <text x="${demarcRightX + 18}" y="${hflY - 7}" fill="#4f5f6d" font-size="12">H.F.L.</text>
    <text x="${centerX - 95}" y="${groundY + 86}" fill="#3f4a55" font-size="13">Natural Ground / Subsoil</text>
    ${fillPoly}
    ${cutPoly}
    ${berms}
    ${layerRects.join("")}
    <line x1="${centerX - halfTop}" y1="${topY}" x2="${centerX + halfTop}" y2="${topY}" stroke="#2e3b49" stroke-width="2.4" />
    <text x="${centerX + halfTop + 16}" y="${topY - 2}" fill="#253748" font-size="14" font-weight="700">Top of Formation</text>
    ${drawDim(centerX - halfTop, centerX + halfTop, trackTopY - 84, `${r3(s.formationWidthFill)} m`)}
    ${segDims.join("")}
    ${segLabelRows}
    ${layerCallouts}
    <line x1="${centerX}" y1="${topY}" x2="${centerX}" y2="${centerRef}" stroke="#2f4d6a" stroke-width="1.8" marker-start="url(#arrowSmall)" marker-end="url(#arrowSmall)" />
    <text x="${centerX + 10}" y="${(topY + centerRef) / 2}" fill="#2f4d6a" font-size="12" font-weight="700">${row.type === "CUTTING" ? `Cut = ${r3(row.cut)} m` : `e = ${r3(row.bank)} m`}</text>
    ${slopeLabels}
    <text x="${centerX - 130}" y="${bodyYRef + 26}" fill="#344553" font-size="13" font-weight="700">${bodyLabel}</text>
    <text x="${centerX - 132}" y="${bodyYRef + 42}" fill="#344553" font-size="12">${bodySubLabel}</text>
    ${levelText}
    <text x="42" y="40" fill="#1f2e3a" font-size="16" font-weight="700">Cross Section of Track</text>
    <defs>
      <pattern id="hatchFill" patternUnits="userSpaceOnUse" width="8" height="8">
        <rect width="8" height="8" fill="#dfe9de" />
        <path d="M-2,2 l4,-4 M0,8 l8,-8 M6,10 l4,-4" stroke="#8ca290" stroke-width="1" />
      </pattern>
      <pattern id="embFillHatch" patternUnits="userSpaceOnUse" width="10" height="10" patternTransform="rotate(20)">
        <rect width="10" height="10" fill="#dfe9de" />
        <path d="M0,0 L0,10" stroke="#9cb3a3" stroke-width="1" />
      </pattern>
      <marker id="arrowSmall" markerWidth="7" markerHeight="7" refX="3.5" refY="3.5" orient="auto-start-reverse">
        <path d="M0,0 L7,3.5 L0,7 z" fill="#2f4d6a"/>
      </marker>
    </defs>
  `;

  resetCrossView();
  els.crossSectionModal.showModal();
}

function bindEvents() {
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

  // Populate date display
  const dateEl = document.getElementById("dateDisplay");
  if (dateEl) {
    const now = new Date();
    const days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    dateEl.textContent = `${days[now.getDay()]}, ${now.getDate()} ${months[now.getMonth()]} ${now.getFullYear()}`;
  }

  // Verify Calculations button
  const verifyBtn = document.getElementById("verifyCalcBtn");
  if (verifyBtn) {
    verifyBtn.addEventListener("click", () => {
      verifyBtn.classList.add("verifying");
      verifyBtn.disabled = true;
      recalculate();
      setTimeout(() => {
        verifyBtn.classList.remove("verifying");
        verifyBtn.classList.add("verified");
        setTimeout(() => {
          verifyBtn.classList.remove("verified");
          verifyBtn.disabled = false;
        }, 2000);
      }, 1200);
    });
  }

  // --- Keyboard Shortcuts ---
  document.addEventListener("keydown", (e) => {
    // Save: Ctrl+S / Cmd+S
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
      e.preventDefault();
      if (els.saveProjectBtn) els.saveProjectBtn.click();
    }
    // Export PDF: Ctrl+E / Cmd+E
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "e") {
      e.preventDefault();
      if (els.exportBtn) els.exportBtn.click();
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
    els.importProjectBtn.addEventListener("click", () => {
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
        const files = Array.from(e.dataTransfer.files).filter(f => f.name.match(/\.(xlsx|xls|csv)$/i));

        // Show loading state on cursor 
        document.body.style.cursor = "wait";

        for (const file of files) {
          try {
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
            alert(`Could not process dropped file ${file.name}: ${err.message}`);
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
      resetForNewProject();
      alert("Workspace reset to default. Create a new project to continue.");
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

    if (els.takeSnapshotBtn) {
      els.takeSnapshotBtn.addEventListener("click", () => {
        const name = els.snapshotNameInput.value.trim() || `Snapshot ${new Date().toLocaleTimeString()}`;
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
        <button class="btn btn-secondary" onclick="restoreSnapshot('${snap.id}')">Restore</button>
      </div>
    `).reverse().join("");
  }

  window.restoreSnapshot = (id) => {
    const snap = state.snapshots.find(s => s.id === id);
    if (!snap) return;
    if (!confirm(`Restore snapshot "${snap.name}"? Current unsaved changes will be lost.`)) return;

    state.rawRows = JSON.parse(JSON.stringify(snap.rawRows));
    state.curves = JSON.parse(JSON.stringify(snap.curves));
    state.bridges = JSON.parse(JSON.stringify(snap.bridges));
    state.loops = JSON.parse(JSON.stringify(snap.loops));
    state.levels = JSON.parse(JSON.stringify(snap.levels));
    state.settings = JSON.parse(JSON.stringify(snap.settings));

    recalculate();
    els.snapshotsModal?.close();
  };

  const importLevelsFile = async (file) => {
    const ext = file.name.toLowerCase();
    if (!(ext.endsWith(".csv") || ext.endsWith(".xlsx") || ext.endsWith(".xls"))) {
      throw new Error("Supported files are .csv, .xlsx, .xls");
    }
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: "array", cellDates: false });

    const wsLevels = wb.Sheets[wb.SheetNames[0]];
    const aoaLevels = XLSX.utils.sheet_to_json(wsLevels, { header: 1, defval: "", raw: false, blankrows: false });
    if (!aoaLevels.length) throw new Error("File has no data rows.");

    const interval = inferImportInterval();
    const startCh = inferImportStartChainage();
    const parsedResult = parseImportedRows(aoaLevels, startCh, interval);
    if (parsedResult.error) throw new Error(parsedResult.error);
    const parsed = parsedResult.rows;
    if (!parsed.length) {
      throw new Error("No valid rows found after header detection. Check Ground Level / Proposed Level values.");
    }
    state.rawRows = parsed;

    if (wb.SheetNames.length > 1) {
      for (let i = 1; i < wb.SheetNames.length; i++) {
        const sheetName = wb.SheetNames[i].toLowerCase();
        const ws = wb.Sheets[wb.SheetNames[i]];
        const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false, blankrows: false });
        if (!aoa.length) continue;

        if (sheetName.includes("bridge") || detectBridgeHeaderRow(aoa)) {
          const bRes = parseBridgeRowsFromAoa(aoa);
          if (bRes && !bRes.error && bRes.rows.length) {
            state.bridgeRows = bRes.rows;
            state.project.uploads.bridges = true;
          }
        }
        else if (sheetName.includes("curve") || detectSimpleHeaderRow(aoa, ["chainage", "curve", "radius", "length"])) {
          const cRes = parseCurveRowsFromAoa(aoa);
          if (cRes && cRes.rows.length) {
            state.curveRows = cRes.rows;
            state.project.uploads.curves = true;
          }
        }
        else if (sheetName.includes("loop") || sheetName.includes("station") || sheetName.includes("platform") || detectSimpleHeaderRow(aoa, ["stations", "csb", "loopstart", "loopend", "pfstart", "pfend", "width"])) {
          const lRes = parseLoopPlatformRowsFromAoa(aoa);
          if (lRes && lRes.rows.length) {
            state.loopPlatformRows = lRes.rows;
            state.project.uploads.loops = true;
          }
        }
      }
    }

    const ch = parsed.map((r) => r.chainage).sort((a, b) => a - b);
    const chainageNote = Number.isFinite(ch[0]) && Number.isFinite(ch[ch.length - 1])
      ? ` | Range: ${r3(ch[0])} m to ${r3(ch[ch.length - 1])} m`
      : "";
    els.projectMeta.textContent = `Imported ${parsed.length} level rows${chainageNote}`;
    state.project.uploads.levels = true;
    state.project.verified = false;
    updateWizardUI();
    applyProjectGate();

    renderBridgeInputs();
    renderCurveInputs();
    renderLoopInputs();
    recalculate();
  };

  const importBridgeFile = async (file) => {
    const aoa = await readSheetAoaFromFile(file);
    const parsedResult = parseBridgeRowsFromAoa(aoa);
    if (parsedResult.error) throw new Error(parsedResult.error);
    if (!parsedResult.rows.length) throw new Error("No valid bridge rows found.");
    state.bridgeRows = parsedResult.rows;
    state.project.uploads.bridges = true;
    state.project.verified = false;
    updateWizardUI();
    applyProjectGate();
    renderBridgeInputs();
    recalculate();
    alert(`Bridge list imported: ${parsedResult.rows.length} rows.`);
  };

  const importCurveFile = async (file) => {
    const aoa = await readSheetAoaFromFile(file);
    const parsedResult = parseCurveRowsFromAoa(aoa);
    state.curveRows = parsedResult.rows;
    state.project.uploads.curves = true;
    state.project.verified = false;
    updateWizardUI();
    applyProjectGate();
    alert(`Curve list imported: ${state.curveRows.length} rows.`);
    recalculate();
  };

  const importLoopFile = async (file) => {
    const aoa = await readSheetAoaFromFile(file);
    const parsedResult = parseLoopPlatformRowsFromAoa(aoa);
    state.loopPlatformRows = parsedResult.rows;
    state.project.uploads.loops = true;
    state.project.verified = false;
    updateWizardUI();
    applyProjectGate();
    const loopRanges = state.loopPlatformRows.filter((r) => Number.isFinite(r.loopStartCh) && Number.isFinite(r.loopEndCh)).length;
    const pfRanges = state.loopPlatformRows.filter((r) => Number.isFinite(r.pfStartCh) && Number.isFinite(r.pfEndCh)).length;
    alert(`Loops & platforms imported: ${state.loopPlatformRows.length} rows (Loop ranges: ${loopRanges}, Platform ranges: ${pfRanges}).`);
    recalculate();
  };

  if (els.importBtn && els.importOptionsModal) {
    els.importBtn.addEventListener("click", () => els.importOptionsModal.showModal());
  }

  if (els.closeImportOptionsBtn && els.importOptionsModal) {
    els.closeImportOptionsBtn.addEventListener("click", () => els.importOptionsModal.close());
  }

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
      }
    });
  }

  if (els.bridgeAddBtn) {
    els.bridgeAddBtn.addEventListener("click", () => {
      const base = inferImportStartChainage();
      const nextNo = state.bridgeRows.length + 1;
      state.bridgeRows.push({
        bridgeNo: `BR-${nextNo}`,
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
      state.loopPlatformRows.push({
        station: `LP-${state.loopPlatformRows.length + 1}`,
        csb: null,
        tc: 0,
        loopStartCh: null,
        loopEndCh: null,
        pfWidth: 0,
        pfStartCh: null,
        pfEndCh: null,
        remarks: "",
      });
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
      const delBtn = e.target.closest("[data-loop-del]");
      if (!delBtn) return;
      const i = Number(delBtn.dataset.loopDel);
      if (!Number.isFinite(i)) return;
      state.loopPlatformRows = state.loopPlatformRows.filter((_, idx) => idx !== i);
      state.project.uploads.loops = state.loopPlatformRows.length > 0;
      state.project.verified = false;
      updateWizardUI();
      applyProjectGate();
      renderLoopInputs();
      recalculate();
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
        loadProjectFromPayload(parsed);
      } catch (err) {
        console.error("Project import error:", err);
        alert(`Project import failed: ${err?.message || "Invalid .EW file."}`);
      } finally {
        e.target.value = "";
      }
    });
  }

  els.exportBtn.addEventListener("click", () => {
    els.pdfExportModal?.showModal();
  });

  if (els.closePdfExportBtn && els.pdfExportModal) {
    els.closePdfExportBtn.addEventListener("click", () => els.pdfExportModal.close());
  }
  if (els.exportCalcPdfBtn) {
    els.exportCalcPdfBtn.addEventListener("click", () => {
      exportCalculationPdf();
      els.pdfExportModal?.close();
    });
  }
  if (els.exportCrossPdfBtn) {
    els.exportCrossPdfBtn.addEventListener("click", () => {
      exportCrossSectionsPdf();
      els.pdfExportModal?.close();
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

  els.tableBody.addEventListener("click", (e) => {
    const trigger = e.target.closest("[data-cross-index]");
    if (!trigger) return;
    const i = Number(trigger.dataset.crossIndex);
    if (Number.isFinite(i)) drawCrossSection(state.calcRows[i]);
  });

  els.closeCrossBtn.addEventListener("click", () => els.crossSectionModal.close());
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
}

async function init() {
  localStorage.clear();
  state.meta = {};
  const visualLayerDefaults = {
    ballastCushionThickness: 0.35,
    topLayerThickness: 0.5,
    activeSqCategory: 3,
  };
  state.defaultSettings = { ...visualLayerDefaults };
  state.seedDefaultSettings = { ...state.defaultSettings };
  state.settings = { ...state.defaultSettings };
  state.rawRows = [];
  state.seedRows = [];
  state.seedMeta = {};

  bindEvents();
  bindCrossCanvasInteraction();
  setWorkPage("overview");
  setResultTab("qty");
  updateWizardUI();
  applyProjectGate();
  recalculate();
}

init();
