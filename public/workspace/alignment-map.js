/**
 * alignment-map.js
 * Handles the Leaflet.js (Free Map) integration for geographic alignment and earthwork visual overlays.
 */

let alignmentMap = null;
let baseLayers = {};
let mapItems = null; // We'll use a FeatureGroup for easy clearing
let pendingStationPlanTarget = "";
let pendingStationPlanFile = null;
let googleEarthMap = null;
let googleEarthPolyline = null;
let googleEarthMarkers = [];
let googleEarthInfoWindow = null;
let googleMapsApiLoadPromise = null;
let googleMapsApiKeyInUse = "";
let googleEarthMapListenersBound = false;
let googleMapsAuthFailed = false;
const GOOGLE_MAPS_API_KEY_STORAGE_KEY = "earthsoft_google_maps_api_key";
const GOOGLE_MAPS_MAP_ID_STORAGE_KEY = "earthsoft_google_maps_map_id";
const MAP_PSEUDO_3D_DEFAULTS = {
    enabled: false,
    terrain: true,
    corridor: true,
    earthwork: true,
    structures: true,
    labels: true,
    boundary: true,
    ewStructures: true,
};

function getMapPseudo3DSettings() {
    const saved = state?.settings?.mapPseudo3d && typeof state.settings.mapPseudo3d === "object"
        ? state.settings.mapPseudo3d
        : {};
    return {
        enabled: saved.enabled === true,
        terrain: saved.terrain !== false,
        corridor: saved.corridor !== false,
        earthwork: saved.earthwork !== false,
        structures: saved.structures !== false,
        labels: saved.labels !== false,
        boundary: saved.boundary !== false,
        ewStructures: saved.ewStructures !== false,
    };
}

function syncMapPseudo3DControls() {
    const settings = getMapPseudo3DSettings();
    const mapping = {
        mapPseudo3dEnabledToggle: "enabled",
        mapOverlayTerrainToggle: "terrain",
        mapOverlayCorridorToggle: "corridor",
        mapOverlayEarthworkToggle: "earthwork",
        mapOverlayStructuresToggle: "structures",
        mapOverlayLabelsToggle: "labels",
        mapOverlayBoundaryToggle: "boundary",
        mapOverlayEWStructuresToggle: "ewStructures",
    };

    Object.entries(mapping).forEach(([id, key]) => {
        const input = document.getElementById(id);
        if (!input) return;
        input.checked = Boolean(settings[key]);
        if (key !== "enabled") {
            input.disabled = !settings.enabled;
            input.closest(".map-toggle-chip")?.classList.toggle("is-disabled", !settings.enabled);
        }
    });
}

function setMapPseudo3DSetting(key, value) {
    state.settings = state.settings || {};
    state.settings.mapPseudo3d = {
        ...MAP_PSEUDO_3D_DEFAULTS,
        ...getMapPseudo3DSettings(),
        [key]: value,
    };
    syncMapPseudo3DControls();
    saveState();
    if (state.kmlData) drawAlignmentMap();
}

// Initialize UI and Event Listeners when DOM is ready
document.addEventListener("DOMContentLoaded", () => {
    // Map control buttons
    const importKmlBtn = document.getElementById("importKmlBtn");
    const kmlImportInput = document.getElementById("kmlImportInput");
    const importStationPlanBtn = document.getElementById("importStationPlanBtn");
    const stationPlanImportInput = document.getElementById("stationPlanImportInput");
    const stationPlanModal = document.getElementById("stationPlanModal");
    const stationPlanStationSelect = document.getElementById("stationPlanStationSelect");
    const confirmStationPlanModalBtn = document.getElementById("confirmStationPlanModalBtn");
    const cancelStationPlanModalBtn = document.getElementById("cancelStationPlanModalBtn");
    const closeStationPlanModalBtn = document.getElementById("closeStationPlanModalBtn");
    const clearMapBtn = document.getElementById("clearMapBtn");
    const mapTypeSelect = document.getElementById("mapTypeSelect");
    const mapPseudo3dEnabledToggle = document.getElementById("mapPseudo3dEnabledToggle");
    const mapOverlayTerrainToggle = document.getElementById("mapOverlayTerrainToggle");
    const mapOverlayCorridorToggle = document.getElementById("mapOverlayCorridorToggle");
    const mapOverlayEarthworkToggle = document.getElementById("mapOverlayEarthworkToggle");
    const mapOverlayStructuresToggle = document.getElementById("mapOverlayStructuresToggle");
    const mapOverlayLabelsToggle = document.getElementById("mapOverlayLabelsToggle");
    const mapOverlayBoundaryToggle = document.getElementById("mapOverlayBoundaryToggle");
    const mapOverlayEWStructuresToggle = document.getElementById("mapOverlayEWStructuresToggle");
    const googleEarthImportKmlBtn = document.getElementById("googleEarthImportKmlBtn");
    const saveGoogleEarthApiKeyBtn = document.getElementById("saveGoogleEarthApiKeyBtn");
    const clearGoogleEarthApiKeyBtn = document.getElementById("clearGoogleEarthApiKeyBtn");
    const googleEarthApiKeyInput = document.getElementById("googleEarthApiKeyInput");
    const googleEarthMapIdInput = document.getElementById("googleEarthMapIdInput");
    const downloadAlignmentKmlBtn = document.getElementById("downloadAlignmentKmlBtn");
    const openGoogleEarthWebBtn = document.getElementById("openGoogleEarthWebBtn");
    const refreshGoogleEarthBtn = document.getElementById("refreshGoogleEarthBtn");
    const googleEarthTiltUpBtn = document.getElementById("googleEarthTiltUpBtn");
    const googleEarthTiltDownBtn = document.getElementById("googleEarthTiltDownBtn");
    const googleEarthRotateLeftBtn = document.getElementById("googleEarthRotateLeftBtn");
    const googleEarthRotateRightBtn = document.getElementById("googleEarthRotateRightBtn");
    const googleEarthResetCameraBtn = document.getElementById("googleEarthResetCameraBtn");
    const googleEarthAlignCameraBtn = document.getElementById("googleEarthAlignCameraBtn");

    if (importKmlBtn && kmlImportInput) {
        importKmlBtn.addEventListener("click", () => kmlImportInput.click());
        kmlImportInput.addEventListener("change", handleKmlImport);
    }
    if (googleEarthImportKmlBtn && kmlImportInput) {
        googleEarthImportKmlBtn.addEventListener("click", () => kmlImportInput.click());
    }
    if (importStationPlanBtn && stationPlanImportInput) {
        importStationPlanBtn.addEventListener("click", () => stationPlanImportInput.click());
        stationPlanImportInput.addEventListener("change", handleStationPlanImport);
    }
    if (confirmStationPlanModalBtn && stationPlanStationSelect) {
        confirmStationPlanModalBtn.addEventListener("click", () => {
            const selectedStation = String(stationPlanStationSelect.value || "").trim();
            if (!selectedStation) {
                alert("Select a station first.");
                return;
            }
            attachPendingStationPlan(selectedStation);
        });
    }
    if (cancelStationPlanModalBtn) {
        cancelStationPlanModalBtn.addEventListener("click", closeStationPlanDialog);
    }
    if (closeStationPlanModalBtn) {
        closeStationPlanModalBtn.addEventListener("click", closeStationPlanDialog);
    }
    if (stationPlanModal) {
        stationPlanModal.addEventListener("close", () => {
            if (stationPlanModal.returnValue !== "confirmed") {
                clearPendingStationPlan();
            }
            stationPlanModal.returnValue = "";
        });
    }
    if (clearMapBtn) {
        clearMapBtn.addEventListener("click", clearMapData);
    }
    if (mapTypeSelect) {
        mapTypeSelect.addEventListener("change", (e) => {
            if (alignmentMap) {
                const layerName = e.target.value;
                Object.values(baseLayers).forEach(layer => alignmentMap.removeLayer(layer));
                if (baseLayers[layerName]) {
                    alignmentMap.addLayer(baseLayers[layerName]);
                    // Toggle dark filter class for OSM
                    const container = alignmentMap.getContainer();
                    if (layerName === 'roadmap' || layerName === 'terrain') {
                        container.classList.remove('leaflet-satellite-layer');
                    } else {
                        container.classList.add('leaflet-satellite-layer');
                    }
                }
            }
        });
    }
    if (mapPseudo3dEnabledToggle) {
        mapPseudo3dEnabledToggle.addEventListener("change", (e) => setMapPseudo3DSetting("enabled", Boolean(e.target.checked)));
    }
    if (mapOverlayTerrainToggle) {
        mapOverlayTerrainToggle.addEventListener("change", (e) => setMapPseudo3DSetting("terrain", Boolean(e.target.checked)));
    }
    if (mapOverlayCorridorToggle) {
        mapOverlayCorridorToggle.addEventListener("change", (e) => setMapPseudo3DSetting("corridor", Boolean(e.target.checked)));
    }
    if (mapOverlayEarthworkToggle) {
        mapOverlayEarthworkToggle.addEventListener("change", (e) => setMapPseudo3DSetting("earthwork", Boolean(e.target.checked)));
    }
    if (mapOverlayStructuresToggle) {
        mapOverlayStructuresToggle.addEventListener("change", (e) => setMapPseudo3DSetting("structures", Boolean(e.target.checked)));
    }
    if (mapOverlayLabelsToggle) {
        mapOverlayLabelsToggle.addEventListener("change", (e) => setMapPseudo3DSetting("labels", Boolean(e.target.checked)));
    }
    if (mapOverlayBoundaryToggle) {
        mapOverlayBoundaryToggle.addEventListener("change", (e) => setMapPseudo3DSetting("boundary", Boolean(e.target.checked)));
    }
    if (mapOverlayEWStructuresToggle) {
        mapOverlayEWStructuresToggle.addEventListener("change", (e) => setMapPseudo3DSetting("ewStructures", Boolean(e.target.checked)));
    }
    if (googleEarthApiKeyInput) {
        googleEarthApiKeyInput.value = getStoredGoogleMapsApiKey();
        googleEarthApiKeyInput.addEventListener("keydown", (event) => {
            if (event.key === "Enter") {
                event.preventDefault();
                saveGoogleEarthApiKey();
            }
        });
    }
    if (googleEarthMapIdInput) {
        googleEarthMapIdInput.value = getStoredGoogleMapsMapId();
        googleEarthMapIdInput.addEventListener("keydown", (event) => {
            if (event.key === "Enter") {
                event.preventDefault();
                saveGoogleEarthApiKey();
            }
        });
    }
    if (saveGoogleEarthApiKeyBtn) {
        saveGoogleEarthApiKeyBtn.addEventListener("click", saveGoogleEarthApiKey);
    }
    if (clearGoogleEarthApiKeyBtn) {
        clearGoogleEarthApiKeyBtn.addEventListener("click", clearGoogleEarthApiKey);
    }
    if (downloadAlignmentKmlBtn) {
        downloadAlignmentKmlBtn.addEventListener("click", downloadAlignmentKml);
    }
    if (openGoogleEarthWebBtn) {
        openGoogleEarthWebBtn.addEventListener("click", () => {
            window.open("https://earth.google.com/web/", "_blank", "noopener");
        });
    }
    if (refreshGoogleEarthBtn) {
        refreshGoogleEarthBtn.addEventListener("click", () => renderGoogleEarthPage(true));
    }
    if (googleEarthTiltUpBtn) {
        googleEarthTiltUpBtn.addEventListener("click", () => adjustGoogleEarthTilt(12));
    }
    if (googleEarthTiltDownBtn) {
        googleEarthTiltDownBtn.addEventListener("click", () => adjustGoogleEarthTilt(-12));
    }
    if (googleEarthRotateLeftBtn) {
        googleEarthRotateLeftBtn.addEventListener("click", () => adjustGoogleEarthHeading(-20));
    }
    if (googleEarthRotateRightBtn) {
        googleEarthRotateRightBtn.addEventListener("click", () => adjustGoogleEarthHeading(20));
    }
    if (googleEarthResetCameraBtn) {
        googleEarthResetCameraBtn.addEventListener("click", resetGoogleEarthCamera);
    }
    if (googleEarthAlignCameraBtn) {
        googleEarthAlignCameraBtn.addEventListener("click", alignGoogleEarthCameraToCorridor);
    }
    syncMapPseudo3DControls();
    updateGoogleEarthSummary();
    syncGoogleEarthEmptyState();
    syncGoogleEarthControls();

    // Initialize Leaflet Map
    initLeafletMap();

    // Handle map resize on page reveal
    const workNavBtns = document.querySelectorAll(".work-nav-btn");
    workNavBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            if (btn.dataset.workPageBtn === "alignment-map") {
                setTimeout(() => {
                    if (alignmentMap) {
                        alignmentMap.invalidateSize();
                        if (state.kmlData) drawAlignmentMap();
                    }
                }, 300);
            }
            if (btn.dataset.workPageBtn === "google-earth") {
                setTimeout(() => {
                    renderGoogleEarthPage();
                }, 300);
            }
        });
    });

    document.addEventListener("click", (event) => {
        const uploadBtn = event.target.closest("[data-upload-station-plan]");
        if (!uploadBtn || !stationPlanImportInput) return;
        pendingStationPlanTarget = String(uploadBtn.dataset.uploadStationPlan || "").trim();
        if (!pendingStationPlanTarget) return;
        stationPlanImportInput.click();
    });

    if (state.activeWorkPage === "google-earth") {
        setTimeout(() => renderGoogleEarthPage(), 120);
    }
});

function initLeafletMap() {
    const mapContainer = document.getElementById("alignmentMapContainer");
    if (!mapContainer) return;

    alignmentMap = L.map(mapContainer, {
        center: [20.5937, 78.9629], // India
        zoom: 5,
        zoomControl: true,
        attributionControl: true
    });

    // Define Base Layers
    baseLayers = {
        roadmap: L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
            attribution: '&copy; OpenStreetMap contributors'
        }),
        satellite: L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
            attribution: 'Tiles &copy; Esri &mdash; Source: Esri, i-cubed, USDA, USGS, AEX, GeoEye, Getmapping, Aerogrid, IGN, IGP, UPR-EGP, and the GIS User Community'
        }),
        hybrid: L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
            attribution: 'Tiles &copy; Esri'
        }),
        terrain: L.tileLayer('https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png', {
            attribution: 'Tiles &copy; OpenTopoMap contributors'
        })
    };

    // Default to Hybrid
    alignmentMap.addLayer(baseLayers.hybrid);
    alignmentMap.getContainer().classList.add('leaflet-satellite-layer');

    [
        ["mapPseudo3dTerrainPane", 405],
        ["mapPseudo3dEarthworkPane", 415],
        ["mapPseudo3dCorridorPane", 425],
        ["mapPseudo3dStructurePane", 435],
        ["mapPseudo3dRailPane", 445],
        ["mapPseudo3dLabelPane", 455],
    ].forEach(([name, zIndex]) => {
        if (!alignmentMap.getPane(name)) {
            const pane = alignmentMap.createPane(name);
            pane.style.zIndex = String(zIndex);
        }
    });

    mapItems = L.featureGroup().addTo(alignmentMap);

    if (state.kmlData) {
        drawAlignmentMap();
    }
}

function clearMapData() {
    if (!confirm("Are you sure you want to clear all map data? This will remove the alignment and all station plans.")) return;

    state.kmlData = null;
    state.stationPlans = {};
    if (state?.project?.uploads) {
        state.project.uploads.kml = false;
    }
    saveState();
    if (typeof updateWizardUI === "function") updateWizardUI();
    if (typeof applyProjectGate === "function") applyProjectGate();

    if (mapItems) mapItems.clearLayers();

    document.getElementById("alignmentMapContainer").style.display = "none";
    document.getElementById("alignmentMapEmpty").style.display = "flex";
    clearGoogleEarthOverlays();
    syncGoogleEarthEmptyState();
    updateGoogleEarthSummary();

    console.log("Map data cleared.");
}

// ── KML/KMZ Import ─────────────────────────────────────────────────────────

function handleKmlImport(e) {
    const file = e.target.files[0];
    if (!file) return;

    const fileName = file.name.toLowerCase();

    if (fileName.endsWith('.kmz')) {
        const reader = new FileReader();
        reader.onload = async (evt) => {
            try {
                const jszip = new JSZip();
                const zip = await jszip.loadAsync(evt.target.result);
                const kmlFile = pickPreferredKmzKmlEntry(zip);

                if (!kmlFile) {
                    alert("Could not find a .kml file inside the KMZ archive.");
                    return;
                }

                const kmlText = await kmlFile.async("text");
                parseKMLData(kmlText);
            } catch (err) {
                console.error("KMZ Extraction Error:", err);
                alert("Failed to read KMZ file.");
            } finally {
                e.target.value = "";
            }
        };
        reader.readAsArrayBuffer(file);
    } else {
        const reader = new FileReader();
        reader.onload = (evt) => {
            parseKMLData(evt.target.result);
            e.target.value = "";
        };
        reader.readAsText(file);
    }
}

// Local fallbacks for utilities (in case app.js is still loading/parsing)
const _safeNum = (v, fallback = 0) => (typeof safeNum === 'function' ? safeNum(v, fallback) : (isNaN(parseFloat(v)) ? fallback : parseFloat(v)));

function parseKMLData(kmlText) {
    console.log("Parsing KML data...");

    const points = extractAlignmentPointsFromKml(kmlText);
    if (!points.length) {
        alert("Could not parse coordinates from the KML file.");
        return;
    }

    let cumDist = 0;
    points[0].ch = 0;
    for (let i = 1; i < points.length; i++) {
        cumDist += getDistanceInMeters(points[i - 1], points[i]);
        points[i].ch = cumDist;
    }

    state.kmlData = { points, totalDistance: cumDist };
    if (state?.project?.uploads) {
        state.project.uploads.kml = true;
        state.project.verified = false;
    }
    saveState();
    if (typeof updateWizardUI === "function") updateWizardUI();
    if (typeof applyProjectGate === "function") applyProjectGate();

    alert(`Successfully imported ${points.length} coordinates and mapped ${cumDist.toFixed(0)}m of alignment.`);

    const container = document.querySelector('.work-page[data-work-page="alignment-map"]');
    if (container && container.style.display !== 'none') {
        drawAlignmentMap();
    }
    updateGoogleEarthSummary();
    renderGoogleEarthPage();
}

function pickPreferredKmzKmlEntry(zip) {
    const candidates = [];
    zip.forEach((relativePath, zipEntry) => {
        if (!zipEntry.dir && relativePath.toLowerCase().endsWith(".kml")) {
            candidates.push(zipEntry);
        }
    });

    if (!candidates.length) return null;

    candidates.sort((a, b) => {
        const aPath = a.name.toLowerCase();
        const bPath = b.name.toLowerCase();
        const aScore = (aPath.endsWith("/doc.kml") || aPath === "doc.kml" ? 100 : 0) - aPath.split("/").length;
        const bScore = (bPath.endsWith("/doc.kml") || bPath === "doc.kml" ? 100 : 0) - bPath.split("/").length;
        return bScore - aScore;
    });

    return candidates[0];
}

function extractAlignmentPointsFromKml(kmlText) {
    const xmlPaths = extractPathsFromXmlKml(kmlText);
    const mergedXmlPath = buildAlignmentPath(xmlPaths);
    if (mergedXmlPath.length) {
        return dedupeConsecutivePoints(mergedXmlPath);
    }

    const regexPaths = extractPathsFromRawKml(kmlText);
    const mergedRegexPath = buildAlignmentPath(regexPaths);
    return dedupeConsecutivePoints(mergedRegexPath);
}

function extractPathsFromXmlKml(kmlText) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(kmlText, "application/xml");
    if (xmlDoc.querySelector("parsererror")) return [];

    const paths = [];
    const lineStrings = Array.from(xmlDoc.getElementsByTagName("LineString"));
    lineStrings.forEach((lineString) => {
        const coordsNode = lineString.getElementsByTagName("coordinates")[0];
        const points = parseCoordinateText(coordsNode?.textContent || "");
        if (points.length > 1) paths.push(points);
    });

    const trackNodes = Array.from(xmlDoc.getElementsByTagNameNS("*", "Track"));
    trackNodes.forEach((trackNode) => {
        const coordNodes = Array.from(trackNode.getElementsByTagNameNS("*", "coord"));
        const points = coordNodes
            .map((node) => parseGxCoordinate(node.textContent || ""))
            .filter(Boolean);
        if (points.length > 1) paths.push(points);
    });

    if (!paths.length) {
        const linearRings = Array.from(xmlDoc.getElementsByTagName("LinearRing"));
        linearRings.forEach((ring) => {
            const coordsNode = ring.getElementsByTagName("coordinates")[0];
            const points = parseCoordinateText(coordsNode?.textContent || "");
            if (points.length > 1) paths.push(points);
        });
    }

    return paths;
}

function extractPathsFromRawKml(kmlText) {
    const paths = [];
    const coordRegex = /<coordinates>([\s\S]*?)<\/coordinates>/gi;
    let match;
    while ((match = coordRegex.exec(kmlText)) !== null) {
        const points = parseCoordinateText(match[1] || "");
        if (points.length > 1) paths.push(points);
    }

    if (!paths.length) {
        const gxCoordRegex = /<gx:coord>([\s\S]*?)<\/gx:coord>/gi;
        const trackPoints = [];
        while ((match = gxCoordRegex.exec(kmlText)) !== null) {
            const point = parseGxCoordinate(match[1] || "");
            if (point) trackPoints.push(point);
        }
        if (trackPoints.length > 1) paths.push(trackPoints);
    }

    return paths;
}

function parseCoordinateText(text) {
    return String(text || "")
        .trim()
        .split(/[\s\r\n]+/)
        .map(parseStandardCoordinate)
        .filter(Boolean);
}

function parseStandardCoordinate(tuple) {
    const parts = String(tuple || "").split(",").map((s) => s.trim());
    if (parts.length < 2) return null;
    const lng = parseFloat(parts[0]);
    const lat = parseFloat(parts[1]);
    if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;
    return { lat, lng };
}

function parseGxCoordinate(tuple) {
    const parts = String(tuple || "").trim().split(/\s+/);
    if (parts.length < 2) return null;
    const lng = parseFloat(parts[0]);
    const lat = parseFloat(parts[1]);
    if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;
    return { lat, lng };
}

function dedupeConsecutivePoints(points) {
    return (points || []).reduce((acc, point) => {
        const last = acc[acc.length - 1];
        if (!last || Math.abs(last.lat - point.lat) > 1e-9 || Math.abs(last.lng - point.lng) > 1e-9) {
            acc.push({ lat: point.lat, lng: point.lng });
        }
        return acc;
    }, []);
}

function buildAlignmentPath(paths) {
    const validPaths = (paths || []).filter((path) => Array.isArray(path) && path.length > 1);
    if (!validPaths.length) return [];

    const mergedPaths = mergeConnectedPaths(validPaths);
    let bestPath = mergedPaths[0];
    let bestDistance = getPathDistance(bestPath);

    for (let i = 1; i < mergedPaths.length; i++) {
        const path = mergedPaths[i];
        const distance = getPathDistance(path);
        if (
            distance > bestDistance + 1 ||
            (Math.abs(distance - bestDistance) <= 1 && path.length > bestPath.length)
        ) {
            bestPath = path;
            bestDistance = distance;
        }
    }

    return bestPath;
}

function mergeConnectedPaths(paths) {
    const unused = paths
        .map((path) => dedupeConsecutivePoints(path))
        .filter((path) => path.length > 1);
    const merged = [];

    while (unused.length) {
        let current = unused.shift();
        let changed = true;

        while (changed) {
            changed = false;

            for (let i = 0; i < unused.length; i++) {
                const candidate = unused[i];
                const joined = tryJoinPaths(current, candidate);
                if (joined) {
                    current = joined;
                    unused.splice(i, 1);
                    changed = true;
                    break;
                }
            }
        }

        merged.push(current);
    }

    return merged;
}

function tryJoinPaths(a, b) {
    if (!a?.length || !b?.length) return null;

    const aStart = a[0];
    const aEnd = a[a.length - 1];
    const bStart = b[0];
    const bEnd = b[b.length - 1];

    if (pointsMatch(aEnd, bStart)) return a.concat(b.slice(1));
    if (pointsMatch(aEnd, bEnd)) return a.concat(b.slice(0, -1).reverse());
    if (pointsMatch(aStart, bEnd)) return b.concat(a.slice(1));
    if (pointsMatch(aStart, bStart)) return b.slice(1).reverse().concat(a);

    return null;
}

function pointsMatch(p1, p2, tolerance = 1e-6) {
    if (!p1 || !p2) return false;
    return Math.abs(p1.lat - p2.lat) <= tolerance && Math.abs(p1.lng - p2.lng) <= tolerance;
}

function getPathDistance(points) {
    let total = 0;
    for (let i = 1; i < points.length; i++) {
        total += getDistanceInMeters(points[i - 1], points[i]);
    }
    return total;
}

// ── Station Plan Import ────────────────────────────────────────────────────

function handleStationPlanImport(e) {
    const file = e.target.files[0];
    if (!file) return;

    if (!state.loopPlatformRows || state.loopPlatformRows.length === 0) {
        e.target.value = "";
        alert("Please add Loops & Station data first.");
        return;
    }

    const stationChoices = [...new Set(state.loopPlatformRows.map((r) => String(r.station || "").trim()).filter(Boolean))];
    if (!stationChoices.length) {
        e.target.value = "";
        alert("No station names are available yet. Import loop/station rows first.");
        return;
    }

    pendingStationPlanFile = file;
    e.target.value = "";
    openStationPlanDialog(stationChoices);
}

// ── Map Rendering Core ───────────────────────────────────────────────────

function drawAlignmentMap() {
    if (!alignmentMap || !mapItems) return;

    if (!state.kmlData || !state.kmlData.points || !state.kmlData.points.length) {
        document.getElementById("alignmentMapContainer").style.display = "none";
        document.getElementById("alignmentMapEmpty").style.display = "flex";
        updateMapLegend(0, false);
        return;
    }

    document.getElementById("alignmentMapContainer").style.display = "block";
    document.getElementById("alignmentMapEmpty").style.display = "none";
    alignmentMap.invalidateSize();

    mapItems.clearLayers();

    const points = state.kmlData.points;
    const path = points.map(p => [p.lat, p.lng]);
    const startChOffset = (state.calcRows && state.calcRows.length) ? _safeNum(state.calcRows[0].chainage) : 0;
    const mapPseudo3DSettings = getMapPseudo3DSettings();

    // Base Alignment Polyline
    L.polyline(path, {
        color: '#ffffff',
        weight: 3,
        opacity: 0.8
    }).addTo(mapItems);

    const hasLandBoundary = mapPseudo3DSettings.boundary ? drawLandBoundaryOverlay(points, startChOffset) : false;
    const showMapStations = state.settings?.showMapStations !== false;
    const showMapBridges = state.settings?.showMapBridges !== false;
    const showMapCurves = state.settings?.showMapCurves !== false;

    if (mapPseudo3DSettings.enabled) {
        drawPseudo3DOverlay(points, startChOffset, mapPseudo3DSettings);
    }

    // Earthwork Overlays
    if (state.calcRows && state.calcRows.length > 1) {
        for (let i = 1; i < state.calcRows.length; i++) {
            const r0 = state.calcRows[i - 1], r1 = state.calcRows[i];
            const isFill = r1.bank > 0.001, isCut = r1.cut > 0.001;
            if (!isFill && !isCut) continue;

            const p0 = getLatLngFromChainage(r0.chainage - startChOffset, points);
            const p1 = getLatLngFromChainage(r1.chainage - startChOffset, points);

            if (p0 && p1) {
                L.polyline([[p0.lat, p0.lng], [p1.lat, p1.lng]], {
                    color: isFill ? '#22c55e' : '#f43f5e',
                    weight: 7,
                    opacity: 0.9
                }).addTo(mapItems);
            }
        }
    }

    // Helper to add titled markers
    const addMarker = (midCh, title, content, color) => {
        const pt = getLatLngFromChainage(midCh - startChOffset, points);
        if (!pt) return null;

        return L.circleMarker([pt.lat, pt.lng], {
            radius: 8,
            fillColor: color,
            color: '#ffffff',
            weight: 1,
            opacity: 1,
            fillOpacity: 0.9
        }).bindPopup(content, { maxWidth: 350 }).addTo(mapItems);
    };

    // Bridges
    if (showMapBridges) {
    state.bridgeRows.forEach(br => {
        const mid = (_safeNum(br.startChainage) + _safeNum(br.endChainage)) / 2;
        const spanConfig = `${_safeNum(br.bridgeSpans, 1)} x ${br.bridgeSize || br.bridgeSpanLength || "-"}`;
        const content = `
            <div class="map-popup-title">Bridge ${br.bridgeNo}</div>
            <div class="map-popup-row"><span class="label">Category:</span> <span class="value">${br.bridgeCategory || "-"}</span></div>
            <div class="map-popup-row"><span class="label">Type:</span> <span class="value">${br.bridgeType}</span></div>
            <div class="map-popup-row"><span class="label">Span Config:</span> <span class="value">${spanConfig}</span></div>
            <div class="map-popup-row"><span class="label">Start:</span> <span class="value">${_safeNum(br.startChainage).toFixed(3)}</span></div>
            <div class="map-popup-row"><span class="label">End:</span> <span class="value">${_safeNum(br.endChainage).toFixed(3)}</span></div>
        `;
        addMarker(mid, `Bridge ${br.bridgeNo}`, content, "#3b82f6");
    });
    }

    // Curves
    if (showMapCurves) {
    state.curveRows.forEach((cr, index) => {
        const length = _safeNum(cr.length);
        const startCh = _safeNum(cr.chainage);
        const endCh = length > 0 ? startCh + length : startCh;
        const mid = length > 0 ? startCh + (length / 2) : startCh;
        if (!Number.isFinite(mid)) return;
        const direction = cr.direction || (((index % 2) === 0) ? "Left" : "Right");
        const content = `
            <div class="map-popup-title">${cr.curve || "Curve"}</div>
            <div class="map-popup-row"><span class="label">Chainage:</span> <span class="value">${startCh.toFixed(3)}</span></div>
            <div class="map-popup-row"><span class="label">Length:</span> <span class="value">${length.toFixed(3)} m</span></div>
            <div class="map-popup-row"><span class="label">Radius:</span> <span class="value">${_safeNum(cr.radius).toFixed(3)} m</span></div>
            <div class="map-popup-row"><span class="label">Direction:</span> <span class="value">${direction}</span></div>
        `;
        addMarker(mid, cr.curve || "Curve", content, "#eab308");
    });
    }

    // Stations
    const stationGroups = new Map();
    state.loopPlatformRows.forEach((lp) => {
        const stationName = String(lp.station || "").trim() || "Station";
        const key = stationName.toLowerCase();
        const group = stationGroups.get(key) || {
            station: stationName,
            csb: Number.isFinite(_safeNum(lp.csb, NaN)) ? _safeNum(lp.csb, NaN) : NaN,
            tc: 0,
            pfWidth: 0,
            loopStartCh: Number.POSITIVE_INFINITY,
            loopEndCh: Number.NEGATIVE_INFINITY,
            remarks: [],
        };

        const loopStartCh = _safeNum(lp.loopStartCh, NaN);
        const loopEndCh = _safeNum(lp.loopEndCh, NaN);
        if (Number.isFinite(loopStartCh)) group.loopStartCh = Math.min(group.loopStartCh, loopStartCh);
        if (Number.isFinite(loopEndCh)) group.loopEndCh = Math.max(group.loopEndCh, loopEndCh);
        group.tc = Math.max(group.tc, _safeNum(lp.tc, 0));
        group.pfWidth = Math.max(group.pfWidth, _safeNum(lp.pfWidth, 0));
        if (lp.remarks) group.remarks.push(String(lp.remarks).trim());

        stationGroups.set(key, group);
    });

    if (showMapStations) {
    stationGroups.forEach((station) => {
        const csbCh = Number.isFinite(station.csb) ? station.csb : NaN;
        const mid = Number.isFinite(csbCh)
            ? csbCh
            : (
                Number.isFinite(station.loopStartCh) && Number.isFinite(station.loopEndCh)
                    ? (station.loopStartCh + station.loopEndCh) / 2
                    : NaN
            );
        if (!Number.isFinite(mid)) return;
        let planHtml = '';
        const planKey = station.station.toLowerCase().trim();
        const plan = state.stationPlans[planKey];

        if (plan) {
            if (plan.startsWith('data:application/pdf')) {
                planHtml = `<iframe src="${plan}#toolbar=0" class="station-plan-pdf" style="width:100%;height:200px;border:none;"></iframe>`;
            } else {
                planHtml = `<img src="${plan}" class="station-plan-img" style="width:100%;max-height:150px;object-fit:contain;"/>`;
            }
        }

        const content = `
            ${planHtml}
            <div class="map-popup-title">${station.station}</div>
            <div class="map-popup-row"><span class="label">CSB:</span> <span class="value">${Number.isFinite(csbCh) ? csbCh.toFixed(3) : "-"}</span></div>
            <div class="map-popup-row"><span class="label">Loop CH:</span> <span class="value">${Number.isFinite(station.loopStartCh) ? station.loopStartCh.toFixed(3) : "-"} to ${Number.isFinite(station.loopEndCh) ? station.loopEndCh.toFixed(3) : "-"}</span></div>
            <div class="map-popup-row"><span class="label">Track Centre:</span> <span class="value">${_safeNum(station.tc).toFixed(3)} m</span></div>
            <div class="map-popup-row"><span class="label">Platform Width:</span> <span class="value">${_safeNum(station.pfWidth).toFixed(3)} m</span></div>
            <div class="map-popup-row"><span class="label">Remarks:</span> <span class="value">${station.remarks.length ? station.remarks.join("; ") : "-"}</span></div>
            <button type="button" class="station-plan-action" data-upload-station-plan="${escapeHtmlAttr(station.station)}">
                ${plan ? "Update Conceptual Plan" : "Upload Conceptual Plan"}
            </button>
        `;
        const stationMarker = addMarker(mid, station.station, content, "#06b6d4");
        const pt = getLatLngFromChainage(mid - startChOffset, points);
        if (pt) {
            L.marker([pt.lat, pt.lng], {
                interactive: true,
                icon: L.divIcon({
                    className: "station-label-wrap",
                    html: `<button type="button" class="station-label-chip">${escapeHtml(station.station)}</button>`,
                    iconSize: null,
                }),
            }).bindPopup(content, { maxWidth: 350 }).addTo(mapItems);
        }
        if (stationMarker) {
            stationMarker.bringToFront();
        }
    });
    }

    const bounds = mapItems.getBounds();
    if (bounds.isValid()) {
        alignmentMap.fitBounds(bounds, { padding: [50, 50] });
    }

    updateMapLegend(Object.keys(state.stationPlans || {}).length, hasLandBoundary, mapPseudo3DSettings.enabled);
}

function getStoredGoogleMapsApiKey() {
    try {
        const localValue = String(localStorage.getItem(GOOGLE_MAPS_API_KEY_STORAGE_KEY) || "").trim();
        if (localValue) return localValue;
    } catch (_) {
        // Fall back to local machine config below.
    }
    return getConfiguredGoogleMapsApiKey();
}

function getConfiguredGoogleMapsApiKey() {
    return String(window.EARTHSOFT_GOOGLE_MAPS_API_KEY || "").trim();
}

function getStoredGoogleMapsMapId() {
    try {
        const localValue = String(localStorage.getItem(GOOGLE_MAPS_MAP_ID_STORAGE_KEY) || "").trim();
        if (localValue) return localValue;
    } catch (_) {
        // Fall back to local machine config below.
    }
    return getConfiguredGoogleMapsMapId();
}

function getConfiguredGoogleMapsMapId() {
    return String(window.EARTHSOFT_GOOGLE_MAP_ID || "").trim();
}

function setStoredGoogleMapsApiKey(value) {
    try {
        if (value) localStorage.setItem(GOOGLE_MAPS_API_KEY_STORAGE_KEY, value);
        else localStorage.removeItem(GOOGLE_MAPS_API_KEY_STORAGE_KEY);
    } catch (_) {
        // Ignore storage failures and let the UI continue as a stateless view.
    }
}

function setStoredGoogleMapsMapId(value) {
    try {
        if (value) localStorage.setItem(GOOGLE_MAPS_MAP_ID_STORAGE_KEY, value);
        else localStorage.removeItem(GOOGLE_MAPS_MAP_ID_STORAGE_KEY);
    } catch (_) {
        // Ignore storage failures and let the UI continue as a stateless view.
    }
}

function setGoogleEarthStatus(message, tone = "") {
    const statusEl = document.getElementById("googleEarthApiStatus");
    if (!statusEl) return;
    statusEl.textContent = message;
    statusEl.classList.remove("is-ready", "is-warning");
    if (tone === "ready") statusEl.classList.add("is-ready");
    if (tone === "warning") statusEl.classList.add("is-warning");
}

function setGoogleEarthEmpty(title, text) {
    const titleEl = document.getElementById("googleEarthEmptyTitle");
    const textEl = document.getElementById("googleEarthEmptyText");
    if (titleEl) titleEl.textContent = title;
    if (textEl) textEl.textContent = text;
}

function syncGoogleEarthControls() {
    const controlDeck = document.getElementById("googleEarthControlDeck");
    const buttons = [
        "googleEarthTiltUpBtn",
        "googleEarthTiltDownBtn",
        "googleEarthRotateLeftBtn",
        "googleEarthRotateRightBtn",
        "googleEarthResetCameraBtn",
        "googleEarthAlignCameraBtn",
    ].map((id) => document.getElementById(id)).filter(Boolean);
    const hasInteractive3D = Boolean(googleEarthMap && getStoredGoogleMapsMapId());

    if (controlDeck) {
        controlDeck.style.display = googleEarthMap ? "flex" : "none";
    }
    buttons.forEach((button) => {
        button.disabled = !hasInteractive3D;
    });
    updateGoogleEarthViewHud();
}

function updateGoogleEarthViewHud() {
    const tiltEl = document.getElementById("googleEarthTiltValue");
    const headingEl = document.getElementById("googleEarthHeadingValue");
    const zoomEl = document.getElementById("googleEarthZoomValue");
    const modeEl = document.getElementById("googleEarthModeBadge");

    const tilt = Number(googleEarthMap?.getTilt?.() || 0);
    const heading = Number(googleEarthMap?.getHeading?.() || 0);
    const zoom = Number(googleEarthMap?.getZoom?.() || 0);
    const is3D = tilt > 0.1;

    if (tiltEl) tiltEl.textContent = `Tilt ${tilt.toFixed(0)}°`;
    if (headingEl) headingEl.textContent = `Heading ${(((heading % 360) + 360) % 360).toFixed(0)}°`;
    if (zoomEl) zoomEl.textContent = `Zoom ${zoom.toFixed(1)}`;
    if (modeEl) modeEl.textContent = is3D ? "3D" : "2D";
}

function bindGoogleEarthMapListeners() {
    if (!googleEarthMap || googleEarthMapListenersBound || !window.google?.maps?.event) return;
    ["tilt_changed", "heading_changed", "zoom_changed", "idle"].forEach((eventName) => {
        google.maps.event.addListener(googleEarthMap, eventName, updateGoogleEarthViewHud);
    });
    googleEarthMapListenersBound = true;
    updateGoogleEarthViewHud();
}

function normalizeHeading(value) {
    const normalized = Number(value || 0) % 360;
    return normalized < 0 ? normalized + 360 : normalized;
}

function adjustGoogleEarthTilt(delta) {
    if (!googleEarthMap || !getStoredGoogleMapsMapId()) return;
    const currentTilt = Number(googleEarthMap.getTilt?.() || 0);
    const nextTilt = Math.max(0, Math.min(67.5, currentTilt + delta));
    if (typeof googleEarthMap.setTilt === "function") {
        googleEarthMap.setTilt(nextTilt);
    }
    updateGoogleEarthViewHud();
}

function adjustGoogleEarthHeading(delta) {
    if (!googleEarthMap || !getStoredGoogleMapsMapId()) return;
    const currentHeading = Number(googleEarthMap.getHeading?.() || 0);
    const nextHeading = normalizeHeading(currentHeading + delta);
    if (typeof googleEarthMap.setHeading === "function") {
        googleEarthMap.setHeading(nextHeading);
    }
    updateGoogleEarthViewHud();
}

function getCorridorHeadingDegrees() {
    const points = state.kmlData?.points || [];
    if (points.length < 2) return 0;
    const first = points[0];
    const last = points[points.length - 1];
    const lat1 = first.lat * Math.PI / 180;
    const lat2 = last.lat * Math.PI / 180;
    const dLng = (last.lng - first.lng) * Math.PI / 180;
    const y = Math.sin(dLng) * Math.cos(lat2);
    const x = Math.cos(lat1) * Math.sin(lat2) - Math.sin(lat1) * Math.cos(lat2) * Math.cos(dLng);
    const bearing = Math.atan2(y, x) * 180 / Math.PI;
    return normalizeHeading(bearing);
}

function alignGoogleEarthCameraToCorridor() {
    if (!googleEarthMap || !getStoredGoogleMapsMapId()) return;
    if (typeof googleEarthMap.setHeading === "function") {
        googleEarthMap.setHeading(getCorridorHeadingDegrees());
    }
    if (typeof googleEarthMap.setTilt === "function") {
        googleEarthMap.setTilt(60);
    }
    updateGoogleEarthViewHud();
}

function resetGoogleEarthCamera() {
    if (!googleEarthMap) return;
    if (googleEarthPolyline && typeof window.google?.maps?.LatLngBounds === "function") {
        const bounds = new google.maps.LatLngBounds();
        googleEarthPolyline.getPath().forEach((point) => bounds.extend(point));
        if (!bounds.isEmpty()) googleEarthMap.fitBounds(bounds, 60);
    }
    if (typeof googleEarthMap.setHeading === "function") {
        googleEarthMap.setHeading(0);
    }
    if (typeof googleEarthMap.setTilt === "function") {
        googleEarthMap.setTilt(getStoredGoogleMapsMapId() ? 45 : 0);
    }
    updateGoogleEarthViewHud();
}

function syncGoogleEarthEmptyState() {
    const hasKml = Boolean(state.kmlData?.points?.length);
    const apiKey = getStoredGoogleMapsApiKey();
    const mapId = getStoredGoogleMapsMapId();
    if (!hasKml) {
        setGoogleEarthEmpty(
            "Import a KML to begin",
            "This tab will render the imported corridor on Google satellite imagery and generate an alignment KML for Google Earth.",
        );
        setGoogleEarthStatus("No alignment KML is loaded yet.", "warning");
        syncGoogleEarthControls();
        return;
    }
    if (!apiKey) {
        setGoogleEarthEmpty(
            "Google Maps API key needed",
            "Save a Google Maps JavaScript API key to render the corridor here. You can still download the active alignment KML and open it in Google Earth Web.",
        );
        setGoogleEarthStatus("API key missing. Download KML or add a key to render the satellite view in-app.", "warning");
        syncGoogleEarthControls();
        return;
    }
    if (mapId) {
        setGoogleEarthStatus("Google Maps API key and vector map ID detected. Tilt, rotate, and corridor alignment are available.", "ready");
    } else {
        setGoogleEarthStatus("API key detected. Add a vector map ID to unlock tilt and rotate controls.", "warning");
    }
    syncGoogleEarthControls();
}

function updateGoogleEarthSummary() {
    const alignmentStatusEl = document.getElementById("googleEarthAlignmentStatus");
    const lengthEl = document.getElementById("googleEarthLengthStat");
    const markerEl = document.getElementById("googleEarthMarkerStat");
    const hasKml = Boolean(state.kmlData?.points?.length);
    const stationCount = getStationGroupsForMap().length;
    const bridgeCount = Array.isArray(state.bridgeRows) ? state.bridgeRows.length : 0;
    const curveCount = Array.isArray(state.curveRows) ? state.curveRows.length : 0;
    const markerCount = stationCount + bridgeCount + curveCount;

    if (alignmentStatusEl) {
        alignmentStatusEl.textContent = hasKml
            ? `${state.kmlData.points.length} points imported`
            : "No KML imported";
    }
    if (lengthEl) {
        const totalDistance = Number.isFinite(state.kmlData?.totalDistance) ? state.kmlData.totalDistance : 0;
        lengthEl.textContent = `${(totalDistance / 1000).toFixed(3)} km`;
    }
    if (markerEl) {
        markerEl.textContent = `${markerCount} item${markerCount === 1 ? "" : "s"}`;
    }
}

function saveGoogleEarthApiKey() {
    const apiKeyInput = document.getElementById("googleEarthApiKeyInput");
    const mapIdInput = document.getElementById("googleEarthMapIdInput");
    const key = String(apiKeyInput?.value || "").trim();
    const mapId = String(mapIdInput?.value || "").trim();
    if (!key) {
        alert("Paste a Google Maps API key first.");
        return;
    }
    setStoredGoogleMapsApiKey(key);
    setStoredGoogleMapsMapId(mapId);
    setGoogleEarthStatus(
        mapId
            ? "API key and vector map ID saved in this browser. Loading the satellite viewer..."
            : "API key saved in this browser. Loading the satellite viewer...",
        "warning",
    );
    renderGoogleEarthPage(true);
}

function clearGoogleEarthApiKey() {
    setStoredGoogleMapsApiKey("");
    setStoredGoogleMapsMapId("");
    const apiKeyInput = document.getElementById("googleEarthApiKeyInput");
    const mapIdInput = document.getElementById("googleEarthMapIdInput");
    if (apiKeyInput) apiKeyInput.value = "";
    if (mapIdInput) mapIdInput.value = "";
    googleMapsApiLoadPromise = null;
    googleMapsApiKeyInUse = "";
    clearGoogleEarthOverlays();
    syncGoogleEarthEmptyState();
    renderGoogleEarthPage(true);
}

function ensureGoogleMapsLoaded(apiKey) {
    if (!apiKey) {
        return Promise.reject(new Error("Missing Google Maps API key"));
    }
    if (window.google?.maps) {
        return Promise.resolve(window.google.maps);
    }
    if (googleMapsApiLoadPromise && googleMapsApiKeyInUse === apiKey) {
        return googleMapsApiLoadPromise;
    }

    const existingScript = document.getElementById("earthsoftGoogleMapsScript");
    if (existingScript) existingScript.remove();

    googleMapsApiKeyInUse = apiKey;
    googleMapsAuthFailed = false;
    googleMapsApiLoadPromise = new Promise((resolve, reject) => {
        window.gm_authFailure = () => {
            googleMapsAuthFailed = true;
            reject(new Error("Google Maps authentication failed. Check billing, API restrictions, and allowed referrers."));
        };
        window.__earthsoftGoogleMapsInit = () => resolve(window.google.maps);
        const script = document.createElement("script");
        script.id = "earthsoftGoogleMapsScript";
        script.async = true;
        script.defer = true;
        script.onerror = () => reject(new Error("Failed to load the Google Maps JavaScript API."));
        script.src = `https://maps.googleapis.com/maps/api/js?key=${encodeURIComponent(apiKey)}&v=weekly&loading=async&callback=__earthsoftGoogleMapsInit`;
        document.head.appendChild(script);
    }).catch((error) => {
        googleMapsApiLoadPromise = null;
        throw error;
    });

    return googleMapsApiLoadPromise;
}

function initGoogleEarthMap(forceReload = false) {
    const mapContainer = document.getElementById("googleEarthMapContainer");
    if (!mapContainer || !window.google?.maps) return null;
    const mapId = getStoredGoogleMapsMapId();

    if (forceReload) {
        clearGoogleEarthOverlays();
        mapContainer.innerHTML = "";
        googleEarthMap = null;
        googleEarthInfoWindow = null;
        googleEarthMapListenersBound = false;
    }

    if (!googleEarthMap) {
        const mapOptions = {
            center: { lat: 20.5937, lng: 78.9629 },
            zoom: 5,
            mapTypeId: mapId ? "hybrid" : "satellite",
            streetViewControl: false,
            fullscreenControl: true,
            mapTypeControl: true,
            rotateControl: true,
            scaleControl: true,
            gestureHandling: "greedy",
            headingInteractionEnabled: Boolean(mapId),
            tiltInteractionEnabled: Boolean(mapId),
        };
        if (mapId) {
            mapOptions.mapId = mapId;
            if (window.google?.maps?.RenderingType?.VECTOR) {
                mapOptions.renderingType = google.maps.RenderingType.VECTOR;
            }
        }
        googleEarthMap = new google.maps.Map(mapContainer, {
            ...mapOptions,
        });
    }
    if (!googleEarthInfoWindow) {
        googleEarthInfoWindow = new google.maps.InfoWindow();
    }
    bindGoogleEarthMapListeners();
    syncGoogleEarthControls();
    return googleEarthMap;
}

function clearGoogleEarthOverlays() {
    if (googleEarthPolyline) {
        googleEarthPolyline.setMap(null);
        googleEarthPolyline = null;
    }
    googleEarthMarkers.forEach((marker) => marker.setMap(null));
    googleEarthMarkers = [];
    if (googleEarthInfoWindow) {
        googleEarthInfoWindow.close();
    }
    syncGoogleEarthControls();
}

function renderGoogleEarthPage(forceReload = false) {
    const mapContainer = document.getElementById("googleEarthMapContainer");
    const emptyEl = document.getElementById("googleEarthEmpty");
    if (!mapContainer || !emptyEl) return;

    updateGoogleEarthSummary();

    const points = state.kmlData?.points || [];
    if (!points.length) {
        clearGoogleEarthOverlays();
        mapContainer.style.display = "none";
        emptyEl.style.display = "flex";
        syncGoogleEarthEmptyState();
        return;
    }

    const apiKey = getStoredGoogleMapsApiKey();
    const mapId = getStoredGoogleMapsMapId();
    if (!apiKey) {
        clearGoogleEarthOverlays();
        mapContainer.style.display = "none";
        emptyEl.style.display = "flex";
        syncGoogleEarthEmptyState();
        return;
    }

    setGoogleEarthStatus(
        mapId ? "Loading Google vector-enabled satellite viewer..." : "Loading Google satellite viewer...",
        "warning",
    );
    ensureGoogleMapsLoaded(apiKey)
        .then(() => {
            if (googleMapsAuthFailed) {
                throw new Error("Google Maps authentication failed.");
            }
            initGoogleEarthMap(forceReload);
            drawGoogleEarthMap();
            emptyEl.style.display = "none";
            mapContainer.style.display = "block";
            setGoogleEarthStatus(
                mapId
                    ? "Google viewer is ready with the imported alignment KML and vector map ID."
                    : "Google satellite viewer is ready and synced with the imported alignment KML.",
                "ready",
            );
        })
        .catch((error) => {
            console.error("Google Earth tab failed to load:", error);
            clearGoogleEarthOverlays();
            mapContainer.style.display = "none";
            emptyEl.style.display = "flex";
            setGoogleEarthEmpty(
                "Google viewer could not load",
                "Google rejected the in-app map. Check that billing is enabled, Maps JavaScript API is enabled, and the API key allows this local origin such as http://127.0.0.1:* and http://localhost:*.",
            );
            setGoogleEarthStatus("Google viewer could not load. Allow localhost / 127.0.0.1 in the key referrers and enable Maps JavaScript API + billing.", "warning");
            syncGoogleEarthControls();
        });
}

function drawGoogleEarthMap() {
    if (!googleEarthMap || !window.google?.maps || !state.kmlData?.points?.length) return;

    const mapContainer = document.getElementById("googleEarthMapContainer");
    const emptyEl = document.getElementById("googleEarthEmpty");
    if (mapContainer) mapContainer.style.display = "block";
    if (emptyEl) emptyEl.style.display = "none";

    clearGoogleEarthOverlays();

    const points = state.kmlData.points;
    const path = points.map((point) => ({ lat: point.lat, lng: point.lng }));
    const bounds = new google.maps.LatLngBounds();
    path.forEach((point) => bounds.extend(point));

    googleEarthPolyline = new google.maps.Polyline({
        path,
        map: googleEarthMap,
        geodesic: true,
        strokeColor: "#ffffff",
        strokeOpacity: 0.92,
        strokeWeight: 4,
    });

    const startChOffset = (state.calcRows && state.calcRows.length) ? _safeNum(state.calcRows[0].chainage) : 0;

    (state.bridgeRows || []).forEach((bridge) => {
        const startCh = _safeNum(bridge?.startChainage, NaN);
        const endCh = _safeNum(bridge?.endChainage, NaN);
        if (!Number.isFinite(startCh) || !Number.isFinite(endCh)) return;
        const midCh = (startCh + endCh) / 2;
        const point = getLatLngFromChainage(midCh - startChOffset, points);
        if (!point) return;

        addGoogleEarthMarker(
            point,
            `Bridge ${bridge.bridgeNo || ""}`.trim(),
            "#3b82f6",
            `
                <div class="google-earth-popup-title">Bridge ${escapeHtml(bridge.bridgeNo || "-")}</div>
                <div class="google-earth-popup-row"><span class="label">Category</span><span class="value">${escapeHtml(bridge.bridgeCategory || "-")}</span></div>
                <div class="google-earth-popup-row"><span class="label">Type</span><span class="value">${escapeHtml(bridge.bridgeType || "-")}</span></div>
                <div class="google-earth-popup-row"><span class="label">Start</span><span class="value">${Number.isFinite(startCh) ? startCh.toFixed(3) : "-"}</span></div>
                <div class="google-earth-popup-row"><span class="label">End</span><span class="value">${Number.isFinite(endCh) ? endCh.toFixed(3) : "-"}</span></div>
            `,
        );
        bounds.extend(point);
    });

    (state.curveRows || []).forEach((curve, index) => {
        const length = _safeNum(curve?.length, 0);
        const startCh = _safeNum(curve?.chainage, NaN);
        if (!Number.isFinite(startCh)) return;
        const endCh = length > 0 ? startCh + length : startCh;
        const midCh = startCh + (Math.max(length, 0) / 2);
        const point = getLatLngFromChainage(midCh - startChOffset, points);
        if (!point) return;

        addGoogleEarthMarker(
            point,
            String(curve.curve || `Curve ${index + 1}`),
            "#eab308",
            `
                <div class="google-earth-popup-title">${escapeHtml(curve.curve || `Curve ${index + 1}`)}</div>
                <div class="google-earth-popup-row"><span class="label">Chainage</span><span class="value">${startCh.toFixed(3)}</span></div>
                <div class="google-earth-popup-row"><span class="label">Length</span><span class="value">${length.toFixed(3)} m</span></div>
                <div class="google-earth-popup-row"><span class="label">Radius</span><span class="value">${_safeNum(curve.radius, 0).toFixed(3)} m</span></div>
                <div class="google-earth-popup-row"><span class="label">Direction</span><span class="value">${escapeHtml(curve.direction || "-")}</span></div>
            `,
        );
        bounds.extend(point);
    });

    getStationGroupsForMap().forEach((station) => {
        const midCh = getStationMidChainage(station);
        if (!Number.isFinite(midCh)) return;
        const point = getLatLngFromChainage(midCh - startChOffset, points);
        if (!point) return;
        const planAttached = Boolean(state.stationPlans?.[String(station.station || "").toLowerCase().trim()]);

        addGoogleEarthMarker(
            point,
            station.station || "Station",
            "#06b6d4",
            `
                <div class="google-earth-popup-title">${escapeHtml(station.station || "Station")}</div>
                <div class="google-earth-popup-row"><span class="label">CSB</span><span class="value">${Number.isFinite(station.csb) ? station.csb.toFixed(3) : "-"}</span></div>
                <div class="google-earth-popup-row"><span class="label">Loop CH</span><span class="value">${Number.isFinite(station.loopStartCh) ? station.loopStartCh.toFixed(3) : "-"} to ${Number.isFinite(station.loopEndCh) ? station.loopEndCh.toFixed(3) : "-"}</span></div>
                <div class="google-earth-popup-row"><span class="label">Track Centre</span><span class="value">${_safeNum(station.tc, 0).toFixed(3)} m</span></div>
                <div class="google-earth-popup-row"><span class="label">Platform Width</span><span class="value">${_safeNum(station.pfWidth, 0).toFixed(3)} m</span></div>
                <div class="google-earth-popup-row"><span class="label">Plan Attached</span><span class="value">${planAttached ? "Yes" : "No"}</span></div>
            `,
        );
        bounds.extend(point);
    });

    if (!bounds.isEmpty()) {
        googleEarthMap.fitBounds(bounds, 60);
    }
    if (getStoredGoogleMapsMapId()) {
        if (typeof googleEarthMap.setTilt === "function") {
            googleEarthMap.setTilt(55);
        }
        if (typeof googleEarthMap.setHeading === "function") {
            googleEarthMap.setHeading(getCorridorHeadingDegrees());
        }
    } else if (typeof googleEarthMap.setTilt === "function") {
        googleEarthMap.setTilt(0);
    }
    updateGoogleEarthViewHud();
    syncGoogleEarthControls();
}

function addGoogleEarthMarker(point, title, color, contentHtml) {
    if (!googleEarthMap || !window.google?.maps) return null;
    const marker = new google.maps.Marker({
        map: googleEarthMap,
        position: { lat: point.lat, lng: point.lng },
        title,
        icon: {
            path: google.maps.SymbolPath.CIRCLE,
            fillColor: color,
            fillOpacity: 0.95,
            strokeColor: "#ffffff",
            strokeWeight: 1.5,
            scale: 6.5,
        },
    });

    marker.addListener("click", () => {
        if (!googleEarthInfoWindow) {
            googleEarthInfoWindow = new google.maps.InfoWindow();
        }
        googleEarthInfoWindow.setContent(`<div class="google-earth-popup">${contentHtml}</div>`);
        googleEarthInfoWindow.open({ anchor: marker, map: googleEarthMap });
    });

    googleEarthMarkers.push(marker);
    return marker;
}

function getStationGroupsForMap() {
    const stationGroups = new Map();
    (state.loopPlatformRows || []).forEach((lp) => {
        const stationName = String(lp.station || "").trim() || "Station";
        const key = stationName.toLowerCase();
        const group = stationGroups.get(key) || {
            station: stationName,
            csb: Number.isFinite(_safeNum(lp.csb, NaN)) ? _safeNum(lp.csb, NaN) : NaN,
            tc: 0,
            pfWidth: 0,
            loopStartCh: Number.POSITIVE_INFINITY,
            loopEndCh: Number.NEGATIVE_INFINITY,
            remarks: [],
        };

        const loopStartCh = _safeNum(lp.loopStartCh, NaN);
        const loopEndCh = _safeNum(lp.loopEndCh, NaN);
        if (Number.isFinite(loopStartCh)) group.loopStartCh = Math.min(group.loopStartCh, loopStartCh);
        if (Number.isFinite(loopEndCh)) group.loopEndCh = Math.max(group.loopEndCh, loopEndCh);
        group.tc = Math.max(group.tc, _safeNum(lp.tc, 0));
        group.pfWidth = Math.max(group.pfWidth, _safeNum(lp.pfWidth, 0));
        if (lp.remarks) group.remarks.push(String(lp.remarks).trim());

        stationGroups.set(key, group);
    });
    return Array.from(stationGroups.values());
}

function getStationMidChainage(station) {
    const csbCh = Number.isFinite(station?.csb) ? station.csb : NaN;
    if (Number.isFinite(csbCh)) return csbCh;
    if (Number.isFinite(station?.loopStartCh) && Number.isFinite(station?.loopEndCh)) {
        return (station.loopStartCh + station.loopEndCh) / 2;
    }
    return NaN;
}

function interpolateAlignmentMetricsAtChainage(chainageAbs) {
    const rows = Array.isArray(state.calcRows) ? state.calcRows : [];
    const defaultFormationW = _safeNum(state.settings?.visual?.formationWidthFill, _safeNum(state.settings?.formationWidthFill, 7.85));
    const defaultMetrics = {
        proposedLevel: 0,
        groundLevel: 0,
        formationWidth: defaultFormationW,
        centerOffset: 0,
        toeWidth: defaultFormationW,
        topWidth: defaultFormationW,
        bank: 0,
        cut: 0,
    };
    if (!rows.length || !Number.isFinite(chainageAbs)) return defaultMetrics;

    const standardTc = _safeNum(state.settings?.visual?.minTc, 5.3);
    const metricForRow = (row) => {
        const envelope = typeof compute3DEnvelopeForRow === "function"
            ? compute3DEnvelopeForRow(row, standardTc, defaultFormationW)
            : { currentFormationW: defaultFormationW };
        return {
            proposedLevel: _safeNum(row?.proposedLevel, 0),
            groundLevel: _safeNum(row?.groundLevel, 0),
            formationWidth: _safeNum(envelope?.currentFormationW, defaultFormationW),
            centerOffset: _safeNum(envelope?.centerOffset, 0),
            toeWidth: row?.bank > 0.001
                ? _safeNum(row?.fillBottom, _safeNum(envelope?.currentFormationW, defaultFormationW))
                : row?.cut > 0.001
                    ? _safeNum(row?.cutBottom, _safeNum(envelope?.currentFormationW, defaultFormationW))
                    : _safeNum(row?.topWidth, _safeNum(envelope?.currentFormationW, defaultFormationW)),
            topWidth: _safeNum(row?.topWidth, _safeNum(envelope?.currentFormationW, defaultFormationW)),
            bank: _safeNum(row?.bank, 0),
            cut: _safeNum(row?.cut, 0),
        };
    };

    const firstCh = _safeNum(rows[0]?.chainage, NaN);
    const lastCh = _safeNum(rows[rows.length - 1]?.chainage, NaN);
    if (!Number.isFinite(firstCh) || !Number.isFinite(lastCh)) return defaultMetrics;
    if (chainageAbs <= firstCh) return metricForRow(rows[0]);
    if (chainageAbs >= lastCh) return metricForRow(rows[rows.length - 1]);

    let low = 0;
    let high = rows.length - 1;
    while (low <= high) {
        const mid = Math.floor((low + high) / 2);
        const midCh = _safeNum(rows[mid]?.chainage, NaN);
        if (midCh < chainageAbs) low = mid + 1;
        else if (midCh > chainageAbs) high = mid - 1;
        else return metricForRow(rows[mid]);
    }

    const leftRow = rows[Math.max(0, high)];
    const rightRow = rows[Math.min(rows.length - 1, low)];
    const leftCh = _safeNum(leftRow?.chainage, chainageAbs);
    const rightCh = _safeNum(rightRow?.chainage, chainageAbs);
    if (!Number.isFinite(leftCh) || !Number.isFinite(rightCh) || Math.abs(rightCh - leftCh) < 0.001) {
        return metricForRow(leftRow || rightRow);
    }

    const leftMetrics = metricForRow(leftRow);
    const rightMetrics = metricForRow(rightRow);
    const ratio = (chainageAbs - leftCh) / (rightCh - leftCh);
    return {
        proposedLevel: leftMetrics.proposedLevel + ((rightMetrics.proposedLevel - leftMetrics.proposedLevel) * ratio),
        groundLevel: leftMetrics.groundLevel + ((rightMetrics.groundLevel - leftMetrics.groundLevel) * ratio),
        formationWidth: leftMetrics.formationWidth + ((rightMetrics.formationWidth - leftMetrics.formationWidth) * ratio),
        centerOffset: leftMetrics.centerOffset + ((rightMetrics.centerOffset - leftMetrics.centerOffset) * ratio),
        toeWidth: leftMetrics.toeWidth + ((rightMetrics.toeWidth - leftMetrics.toeWidth) * ratio),
        topWidth: leftMetrics.topWidth + ((rightMetrics.topWidth - leftMetrics.topWidth) * ratio),
        bank: leftMetrics.bank + ((rightMetrics.bank - leftMetrics.bank) * ratio),
        cut: leftMetrics.cut + ((rightMetrics.cut - leftMetrics.cut) * ratio),
    };
}

function build3DAlignmentSamples(points, startChOffset) {
    if (!Array.isArray(points) || !points.length) return [];
    const sampleWindow = getTypicalChainageStep(state.calcRows || []);
    return points.map((point) => {
        const alignmentCh = _safeNum(point?.ch, NaN);
        const chainageAbs = alignmentCh + startChOffset;
        const metrics = interpolateAlignmentMetricsAtChainage(chainageAbs);
        const tangent = getAlignmentTangentAtChainage(alignmentCh, points, sampleWindow);
        return {
            lat: point.lat,
            lng: point.lng,
            ch: alignmentCh,
            chainageAbs,
            proposedLevel: metrics.proposedLevel,
            groundLevel: metrics.groundLevel,
            formationWidth: metrics.formationWidth,
            centerOffset: metrics.centerOffset,
            toeWidth: metrics.toeWidth,
            topWidth: metrics.topWidth,
            bank: metrics.bank,
            cut: metrics.cut,
            tangent,
        };
    }).filter((sample) => Number.isFinite(sample.lat) && Number.isFinite(sample.lng));
}

function build3DRibbonPolygonCoordinates(samples, xMinResolver, xMaxResolver, altitudeResolver) {
    const left = [];
    const right = [];
    samples.forEach((sample) => {
        if (!sample?.tangent) return;
        const xMin = xMinResolver(sample);
        const xMax = xMaxResolver(sample);
        const altitude = altitudeResolver(sample);
        if (!Number.isFinite(xMin) || !Number.isFinite(xMax) || !Number.isFinite(altitude)) return;
        const leftPoint = offsetLatLng(sample, sample.tangent, xMin);
        const rightPoint = offsetLatLng(sample, sample.tangent, xMax);
        left.push(`${leftPoint.lng},${leftPoint.lat},${altitude.toFixed(3)}`);
        right.push(`${rightPoint.lng},${rightPoint.lat},${altitude.toFixed(3)}`);
    });
    if (left.length < 2 || right.length < 2) return "";
    const ring = [...left, ...right.reverse()];
    ring.push(left[0]);
    return ring.join(" ");
}

function build3DSideFacePolygonCoordinates(samples, topXResolver, topAltitudeResolver, bottomXResolver, bottomAltitudeResolver) {
    const top = [];
    const bottom = [];
    samples.forEach((sample) => {
        if (!sample?.tangent) return;
        const topX = topXResolver(sample);
        const topAltitude = topAltitudeResolver(sample);
        const bottomX = bottomXResolver(sample);
        const bottomAltitude = bottomAltitudeResolver(sample);
        if (!Number.isFinite(topX) || !Number.isFinite(topAltitude) || !Number.isFinite(bottomX) || !Number.isFinite(bottomAltitude)) return;
        const topPoint = offsetLatLng(sample, sample.tangent, topX);
        const bottomPoint = offsetLatLng(sample, sample.tangent, bottomX);
        top.push(`${topPoint.lng},${topPoint.lat},${topAltitude.toFixed(3)}`);
        bottom.push(`${bottomPoint.lng},${bottomPoint.lat},${bottomAltitude.toFixed(3)}`);
    });
    if (top.length < 2 || bottom.length < 2) return "";
    const ring = [...top, ...bottom.reverse()];
    ring.push(top[0]);
    return ring.join(" ");
}

function build3DCorridorPolygonCoordinates(samples) {
    return build3DRibbonPolygonCoordinates(
        samples,
        (sample) => _safeNum(sample.centerOffset, 0) + (Math.max(_safeNum(sample.formationWidth, 7.85), 2) / 2),
        (sample) => _safeNum(sample.centerOffset, 0) - (Math.max(_safeNum(sample.formationWidth, 7.85), 2) / 2),
        (sample) => _safeNum(sample.proposedLevel, 0),
    );
}

function build3DRangeSamples(points, startChOffset, startCh, endCh, stepMeters = 20) {
    if (!Array.isArray(points) || points.length < 2) return [];
    const fromCh = Math.min(startCh, endCh);
    const toCh = Math.max(startCh, endCh);
    if (!Number.isFinite(fromCh) || !Number.isFinite(toCh)) return [];

    const length = Math.max(toCh - fromCh, 0);
    const steps = Math.max(2, Math.min(80, Math.ceil(length / Math.max(stepMeters, 5))));
    const sampleWindow = getTypicalChainageStep(state.calcRows || []);
    const samples = [];

    for (let i = 0; i <= steps; i += 1) {
        const chainageAbs = fromCh + ((length * i) / steps);
        const alignmentCh = chainageAbs - startChOffset;
        const point = getLatLngFromChainage(alignmentCh, points);
        const tangent = getAlignmentTangentAtChainage(alignmentCh, points, sampleWindow);
        const metrics = interpolateAlignmentMetricsAtChainage(chainageAbs);
        if (!point || !tangent) continue;
        samples.push({
            lat: point.lat,
            lng: point.lng,
            ch: alignmentCh,
            chainageAbs,
            proposedLevel: metrics.proposedLevel,
            groundLevel: metrics.groundLevel,
            formationWidth: metrics.formationWidth,
            centerOffset: metrics.centerOffset,
            toeWidth: metrics.toeWidth,
            topWidth: metrics.topWidth,
            bank: metrics.bank,
            cut: metrics.cut,
            tangent,
        });
    }

    return samples;
}

function build3DSpanLineCoordinates(points, startChOffset, startCh, endCh, xResolver, altitudeResolver, stepMeters = 20) {
    const samples = build3DRangeSamples(points, startChOffset, startCh, endCh, stepMeters);
    if (samples.length < 2) return "";
    return samples.map((sample) => {
        const x = xResolver(sample);
        const altitude = altitudeResolver(sample);
        if (!Number.isFinite(x) || !Number.isFinite(altitude)) return "";
        const point = offsetLatLng(sample, sample.tangent, x);
        return `${point.lng},${point.lat},${altitude.toFixed(3)}`;
    }).filter(Boolean).join(" ");
}

function build3DSpanPolygonCoordinates(points, startChOffset, startCh, endCh, xMinResolver, xMaxResolver, altitudeResolver, stepMeters = 20) {
    const samples = build3DRangeSamples(points, startChOffset, startCh, endCh, stepMeters);
    return build3DRibbonPolygonCoordinates(samples, xMinResolver, xMaxResolver, altitudeResolver);
}

function buildKmlPlacemark(name, styleUrl, geometry, description = "") {
    if (!geometry) return "";
    return `
        <Placemark>
            <name>${escapeHtml(name)}</name>
            ${description ? `<description><![CDATA[${description}]]></description>` : ""}
            ${styleUrl ? `<styleUrl>#${styleUrl}</styleUrl>` : ""}
            ${geometry}
        </Placemark>
    `;
}

function buildAlignmentKmlDocument() {
    const points = state.kmlData?.points || [];
    if (!points.length) return "";

    const startChOffset = (state.calcRows && state.calcRows.length) ? _safeNum(state.calcRows[0].chainage) : 0;
    const projectName = escapeHtml(state.project?.name || "Earthsoft Alignment");
    const samples = build3DAlignmentSamples(points, startChOffset);
    const has3DLevels = samples.some((sample) => Math.abs(_safeNum(sample.proposedLevel, 0)) > 0.001);
    const lineCoordinates = samples.map((sample) => `${sample.lng},${sample.lat},${_safeNum(sample.proposedLevel, 0).toFixed(3)}`).join(" ");
    const terrainCoordinates = samples.map((sample) => `${sample.lng},${sample.lat},${_safeNum(sample.groundLevel, 0).toFixed(3)}`).join(" ");
    const corridorPolygonCoordinates = build3DCorridorPolygonCoordinates(samples);

    const stationPlacemarks = getStationGroupsForMap().map((station) => {
        const midCh = getStationMidChainage(station);
        const point = Number.isFinite(midCh) ? getLatLngFromChainage(midCh - startChOffset, points) : null;
        if (!point) return "";
        const metrics = interpolateAlignmentMetricsAtChainage(midCh);
        return `
            <Placemark>
                <name>${escapeHtml(station.station || "Station")}</name>
                <description><![CDATA[CSB: ${Number.isFinite(station.csb) ? station.csb.toFixed(3) : "-"} | Formation RL: ${_safeNum(metrics.proposedLevel, 0).toFixed(3)}]]></description>
                <Point><coordinates>${point.lng},${point.lat},${_safeNum(metrics.proposedLevel, 0).toFixed(3)}</coordinates></Point>
            </Placemark>
        `;
    }).join("");

    const bridgePlacemarks = (state.bridgeRows || []).map((bridge) => {
        const startCh = _safeNum(bridge?.startChainage, NaN);
        const endCh = _safeNum(bridge?.endChainage, NaN);
        if (!Number.isFinite(startCh) || !Number.isFinite(endCh)) return "";
        const midCh = (startCh + endCh) / 2;
        const point = getLatLngFromChainage((midCh - startChOffset), points);
        if (!point) return "";
        const metrics = interpolateAlignmentMetricsAtChainage(midCh);
        return `
            <Placemark>
                <name>Bridge ${escapeHtml(bridge.bridgeNo || "-")}</name>
                <description><![CDATA[${escapeHtml(bridge.bridgeCategory || "Bridge")} | Formation RL: ${_safeNum(metrics.proposedLevel, 0).toFixed(3)}]]></description>
                <Point><coordinates>${point.lng},${point.lat},${_safeNum(metrics.proposedLevel, 0).toFixed(3)}</coordinates></Point>
            </Placemark>
        `;
    }).join("");

    const terrainRibbonCoordinates = build3DRibbonPolygonCoordinates(
        samples,
        (sample) => _safeNum(sample.centerOffset, 0) + (Math.max(_safeNum(sample.toeWidth, sample.formationWidth), _safeNum(sample.formationWidth, 7.85)) / 2) + 14,
        (sample) => _safeNum(sample.centerOffset, 0) - (Math.max(_safeNum(sample.toeWidth, sample.formationWidth), _safeNum(sample.formationWidth, 7.85)) / 2) - 14,
        (sample) => _safeNum(sample.groundLevel, 0),
    );

    const embankmentPlacemarkGroup = buildContiguousSampleGroups(samples, (sample) => _safeNum(sample.bank, 0) > 0.1)
        .map((group, index) => {
            const footprintCoordinates = build3DRibbonPolygonCoordinates(
                group,
                (sample) => _safeNum(sample.centerOffset, 0) + (Math.max(_safeNum(sample.toeWidth, sample.formationWidth), 2) / 2),
                (sample) => _safeNum(sample.centerOffset, 0) - (Math.max(_safeNum(sample.toeWidth, sample.formationWidth), 2) / 2),
                (sample) => _safeNum(sample.groundLevel, 0),
            );
            const leftFaceCoordinates = build3DSideFacePolygonCoordinates(
                group,
                (sample) => _safeNum(sample.centerOffset, 0) + (Math.max(_safeNum(sample.formationWidth, 7.85), 2) / 2),
                (sample) => _safeNum(sample.proposedLevel, 0) - 0.25,
                (sample) => _safeNum(sample.centerOffset, 0) + (Math.max(_safeNum(sample.toeWidth, sample.formationWidth), 2) / 2),
                (sample) => Math.min(_safeNum(sample.proposedLevel, 0) - 0.25, _safeNum(sample.groundLevel, 0)),
            );
            const rightFaceCoordinates = build3DSideFacePolygonCoordinates(
                group,
                (sample) => _safeNum(sample.centerOffset, 0) - (Math.max(_safeNum(sample.formationWidth, 7.85), 2) / 2),
                (sample) => _safeNum(sample.proposedLevel, 0) - 0.25,
                (sample) => _safeNum(sample.centerOffset, 0) - (Math.max(_safeNum(sample.toeWidth, sample.formationWidth), 2) / 2),
                (sample) => Math.min(_safeNum(sample.proposedLevel, 0) - 0.25, _safeNum(sample.groundLevel, 0)),
            );
            return [
                buildKmlPlacemark(
                    `Embankment Berm ${index + 1}`,
                    "embankmentSurface",
                    footprintCoordinates ? `<Polygon>
                        <extrude>0</extrude>
                        <altitudeMode>absolute</altitudeMode>
                        <outerBoundaryIs>
                            <LinearRing>
                                <coordinates>${footprintCoordinates}</coordinates>
                            </LinearRing>
                        </outerBoundaryIs>
                    </Polygon>` : "",
                    "Ground footprint of the embankment / berm envelope.",
                ),
                buildKmlPlacemark(
                    `Embankment Left Face ${index + 1}`,
                    "embankmentSurface",
                    leftFaceCoordinates ? `<Polygon>
                        <extrude>0</extrude>
                        <altitudeMode>absolute</altitudeMode>
                        <outerBoundaryIs>
                            <LinearRing>
                                <coordinates>${leftFaceCoordinates}</coordinates>
                            </LinearRing>
                        </outerBoundaryIs>
                    </Polygon>` : "",
                    "Left embankment face between the formation shoulder and berm / toe.",
                ),
                buildKmlPlacemark(
                    `Embankment Right Face ${index + 1}`,
                    "embankmentSurface",
                    rightFaceCoordinates ? `<Polygon>
                        <extrude>0</extrude>
                        <altitudeMode>absolute</altitudeMode>
                        <outerBoundaryIs>
                            <LinearRing>
                                <coordinates>${rightFaceCoordinates}</coordinates>
                            </LinearRing>
                        </outerBoundaryIs>
                    </Polygon>` : "",
                    "Right embankment face between the formation shoulder and berm / toe.",
                ),
            ].join("");
        })
        .join("");

    const cuttingPlacemarkGroup = buildContiguousSampleGroups(samples, (sample) => _safeNum(sample.cut, 0) > 0.1)
        .map((group, index) => {
            const footprintCoordinates = build3DRibbonPolygonCoordinates(
                group,
                (sample) => _safeNum(sample.centerOffset, 0) + (Math.max(_safeNum(sample.toeWidth, sample.formationWidth), 2) / 2),
                (sample) => _safeNum(sample.centerOffset, 0) - (Math.max(_safeNum(sample.toeWidth, sample.formationWidth), 2) / 2),
                (sample) => _safeNum(sample.groundLevel, 0),
            );
            const leftFaceCoordinates = build3DSideFacePolygonCoordinates(
                group,
                (sample) => _safeNum(sample.centerOffset, 0) + ((Math.max(_safeNum(sample.formationWidth, 7.85), 2) / 2) + 2),
                (sample) => _safeNum(sample.proposedLevel, 0) + 0.25,
                (sample) => _safeNum(sample.centerOffset, 0) + Math.max((Math.max(_safeNum(sample.toeWidth, sample.formationWidth), 2) / 2), ((Math.max(_safeNum(sample.formationWidth, 7.85), 2) / 2) + 2)),
                (sample) => Math.max(_safeNum(sample.proposedLevel, 0) + 0.25, _safeNum(sample.groundLevel, 0)),
            );
            const rightFaceCoordinates = build3DSideFacePolygonCoordinates(
                group,
                (sample) => _safeNum(sample.centerOffset, 0) - ((Math.max(_safeNum(sample.formationWidth, 7.85), 2) / 2) + 2),
                (sample) => _safeNum(sample.proposedLevel, 0) + 0.25,
                (sample) => _safeNum(sample.centerOffset, 0) - Math.max((Math.max(_safeNum(sample.toeWidth, sample.formationWidth), 2) / 2), ((Math.max(_safeNum(sample.formationWidth, 7.85), 2) / 2) + 2)),
                (sample) => Math.max(_safeNum(sample.proposedLevel, 0) + 0.25, _safeNum(sample.groundLevel, 0)),
            );
            return [
                buildKmlPlacemark(
                    `Cutting Surface ${index + 1}`,
                    "cuttingSurface",
                    footprintCoordinates ? `<Polygon>
                        <extrude>0</extrude>
                        <altitudeMode>absolute</altitudeMode>
                        <outerBoundaryIs>
                            <LinearRing>
                                <coordinates>${footprintCoordinates}</coordinates>
                            </LinearRing>
                        </outerBoundaryIs>
                    </Polygon>` : "",
                    "Ground footprint of the cutting envelope.",
                ),
                buildKmlPlacemark(
                    `Cutting Left Face ${index + 1}`,
                    "cuttingSurface",
                    leftFaceCoordinates ? `<Polygon>
                        <extrude>0</extrude>
                        <altitudeMode>absolute</altitudeMode>
                        <outerBoundaryIs>
                            <LinearRing>
                                <coordinates>${leftFaceCoordinates}</coordinates>
                            </LinearRing>
                        </outerBoundaryIs>
                    </Polygon>` : "",
                    "Left cutting face between the corridor shoulder and the top of cut.",
                ),
                buildKmlPlacemark(
                    `Cutting Right Face ${index + 1}`,
                    "cuttingSurface",
                    rightFaceCoordinates ? `<Polygon>
                        <extrude>0</extrude>
                        <altitudeMode>absolute</altitudeMode>
                        <outerBoundaryIs>
                            <LinearRing>
                                <coordinates>${rightFaceCoordinates}</coordinates>
                            </LinearRing>
                        </outerBoundaryIs>
                    </Polygon>` : "",
                    "Right cutting face between the corridor shoulder and the top of cut.",
                ),
            ].join("");
        })
        .join("");

    const standardTc = _safeNum(state.settings?.visual?.minTc, 5.3);
    const loopTrackPlacemarks = (state.loopPlatformRows || []).map((row, index) => {
        const startCh = _safeNum(row?.loopStartCh, NaN);
        const endCh = _safeNum(row?.loopEndCh, NaN);
        if (!Number.isFinite(startCh) || !Number.isFinite(endCh) || Math.abs(endCh - startCh) < 0.1) return "";
        const midCh = (startCh + endCh) / 2;
        const layout = typeof buildStationSequenceLayout === "function"
            ? buildStationSequenceLayout(midCh, row?.station || "", standardTc, { useRanges: true })
            : null;
        const trackItem = layout?.trackItems?.find((item) => item.row === row)
            || layout?.trackItems?.find((item) =>
                String(item?.row?.station || "").trim() === String(row?.station || "").trim() &&
                String(item?.row?.lineType || "").trim() === String(row?.lineType || "").trim() &&
                String(item?.row?.lineName || "").trim() === String(row?.lineName || "").trim() &&
                String(item?.row?.side || "").trim() === String(row?.side || "").trim(),
            )
            || null;
        const xOffset = Number.isFinite(layout?.offsetByItem?.get(trackItem))
            ? layout.offsetByItem.get(trackItem)
            : (String(row?.side || "").toLowerCase() === "right" ? -_safeNum(row?.tc, 0) : _safeNum(row?.tc, 0));
        const coordinates = build3DSpanLineCoordinates(
            points,
            startChOffset,
            startCh,
            endCh,
            () => xOffset,
            (sample) => _safeNum(sample.proposedLevel, 0) + 0.42,
            18,
        );
        return buildKmlPlacemark(
            `${row?.station || "Station"} ${row?.lineType || "Track"} ${row?.lineName || `Line ${index + 1}`}`.trim(),
            "loopTrackLine",
            `<LineString>
                <extrude>0</extrude>
                <altitudeMode>absolute</altitudeMode>
                <coordinates>${coordinates}</coordinates>
            </LineString>`,
            `${row?.lineType || "Track"} line exported from the station / loop layout.`,
        );
    }).join("");

    const platformPlacemarks = (state.loopPlatformRows || []).map((row, index) => {
        const pfStartCh = _safeNum(row?.pfStartCh, NaN);
        const pfEndCh = _safeNum(row?.pfEndCh, NaN);
        const pfWidth = _safeNum(row?.pfWidth, NaN);
        if (!Number.isFinite(pfStartCh) || !Number.isFinite(pfEndCh) || !(pfWidth > 0)) return "";

        const midCh = (pfStartCh + pfEndCh) / 2;
        const layout = typeof buildStationSequenceLayout === "function"
            ? buildStationSequenceLayout(midCh, row?.station || "", standardTc, { useRanges: true })
            : null;
        const refTrack = layout?.trackItems?.find((item) => item.row === row)
            || layout?.trackItems?.find((item) => String(item?.side || "").trim() === String(row?.side || "").trim())
            || layout?.trackItems?.find((item) => item.isMain)
            || layout?.trackItems?.[0]
            || null;
        const refOffset = Number.isFinite(layout?.offsetByItem?.get(refTrack))
            ? layout.offsetByItem.get(refTrack)
            : (String(row?.side || "").toLowerCase() === "right" ? -_safeNum(row?.tc, 0) : _safeNum(row?.tc, 0));
        const isLeft = String(row?.side || "").toLowerCase() === "left";
        const cx = refOffset + (isLeft ? -(2.8 + (pfWidth / 2)) : (2.8 + (pfWidth / 2)));
        const coordinates = build3DSpanPolygonCoordinates(
            points,
            startChOffset,
            pfStartCh,
            pfEndCh,
            () => cx + (pfWidth / 2),
            () => cx - (pfWidth / 2),
            (sample) => _safeNum(sample.proposedLevel, 0) + 0.4,
            14,
        );
        return buildKmlPlacemark(
            `${row?.station || "Station"} Platform ${index + 1}`,
            "platformDeck",
            `<Polygon>
                <extrude>0</extrude>
                <altitudeMode>absolute</altitudeMode>
                <outerBoundaryIs>
                    <LinearRing>
                        <coordinates>${coordinates}</coordinates>
                    </LinearRing>
                </outerBoundaryIs>
            </Polygon>`,
            `Platform deck exported from Earthsoft (${_safeNum(pfWidth, 0).toFixed(3)} m width).`,
        );
    }).join("");

    const bridgeSurfacePlacemarks = (state.bridgeRows || []).map((bridge) => {
        const startCh = _safeNum(bridge?.startChainage, NaN);
        const endCh = _safeNum(bridge?.endChainage, NaN);
        if (!Number.isFinite(startCh) || !Number.isFinite(endCh) || Math.abs(endCh - startCh) < 0.1) return "";
        const metrics = interpolateAlignmentMetricsAtChainage((startCh + endCh) / 2);
        const width = Math.max(_safeNum(metrics.formationWidth, 7.85) + 2, 4);
        const isTunnel = /tunnel/i.test(String(bridge?.bridgeCategory || ""));
        const coordinates = build3DSpanPolygonCoordinates(
            points,
            startChOffset,
            startCh,
            endCh,
            (sample) => _safeNum(sample.centerOffset, metrics.centerOffset) + (width / 2),
            (sample) => _safeNum(sample.centerOffset, metrics.centerOffset) - (width / 2),
            (sample) => _safeNum(sample.proposedLevel, 0) + (isTunnel ? 0.2 : 0),
            18,
        );
        return buildKmlPlacemark(
            `${bridge?.bridgeCategory || "Structure"} ${bridge?.bridgeNo || ""}`.trim(),
            isTunnel ? "tunnelSurface" : "bridgeSurface",
            `<Polygon>
                <extrude>${isTunnel ? 0 : 1}</extrude>
                <altitudeMode>absolute</altitudeMode>
                <outerBoundaryIs>
                    <LinearRing>
                        <coordinates>${coordinates}</coordinates>
                    </LinearRing>
                </outerBoundaryIs>
            </Polygon>`,
            `${bridge?.bridgeCategory || "Structure"} span exported from Earthsoft.`,
        );
    }).join("");

    const retainingWallPlacemarks = (state.earthworkStructures?.retainingWalls || []).map((wall, index) => {
        const startCh = _safeNum(wall?.fromCh, NaN);
        const endCh = _safeNum(wall?.toCh, NaN);
        if (!Number.isFinite(startCh) || !Number.isFinite(endCh) || Math.abs(endCh - startCh) < 0.1) return "";
        const metrics = interpolateAlignmentMetricsAtChainage((startCh + endCh) / 2);
        const offsets = String(wall?.side || "").toLowerCase() === "both"
            ? [
                _safeNum(metrics.centerOffset, 0) + (_safeNum(metrics.formationWidth, 7.85) / 2) + 1.2,
                _safeNum(metrics.centerOffset, 0) - (_safeNum(metrics.formationWidth, 7.85) / 2) - 1.2,
            ]
            : [
                _safeNum(metrics.centerOffset, 0) + ((String(wall?.side || "").toLowerCase() === "left" ? 1 : -1) * ((_safeNum(metrics.formationWidth, 7.85) / 2) + 1.2)),
            ];
        return offsets.map((xOffset, sideIndex) => buildKmlPlacemark(
            `Retaining Wall ${index + 1}${offsets.length > 1 ? `-${sideIndex + 1}` : ""}`,
            "retainingWallLine",
            `<LineString>
                <extrude>1</extrude>
                <altitudeMode>absolute</altitudeMode>
                <coordinates>${build3DSpanLineCoordinates(
                    points,
                    startChOffset,
                    startCh,
                    endCh,
                    () => xOffset,
                    (sample) => _safeNum(sample.proposedLevel, 0) + (String(wall?.wallType || "").includes("4") ? 4 : 3),
                    16,
                )}</coordinates>
            </LineString>`,
            `Retaining wall exported from the earthwork structures sheet.`,
        )).join("");
    }).join("");

    const drainPlacemarks = (state.earthworkStructures?.drains || []).map((drain, index) => {
        const startCh = _safeNum(drain?.fromCh, NaN);
        const endCh = _safeNum(drain?.toCh, NaN);
        if (!Number.isFinite(startCh) || !Number.isFinite(endCh) || Math.abs(endCh - startCh) < 0.1) return "";
        const metrics = interpolateAlignmentMetricsAtChainage((startCh + endCh) / 2);
        const drainType = typeof getDrainTypeMeta === "function" ? getDrainTypeMeta(drain?.drainType).key : String(drain?.drainType || "side_drain");
        let sideOffset = 1.6;
        if (drainType === "catchwater_drain") sideOffset = 4.5;
        if (drainType === "berm_drain") sideOffset = 3.2;
        if (drainType === "yard_drain") sideOffset = 2.2;
        const offsets = String(drain?.side || "").toLowerCase() === "both"
            ? [
                _safeNum(metrics.centerOffset, 0) + (_safeNum(metrics.formationWidth, 7.85) / 2) + sideOffset,
                _safeNum(metrics.centerOffset, 0) - (_safeNum(metrics.formationWidth, 7.85) / 2) - sideOffset,
            ]
            : [
                _safeNum(metrics.centerOffset, 0) + ((String(drain?.side || "").toLowerCase() === "left" ? 1 : -1) * ((_safeNum(metrics.formationWidth, 7.85) / 2) + sideOffset)),
            ];
        return offsets.map((xOffset, sideIndex) => buildKmlPlacemark(
            `Drain ${index + 1}${offsets.length > 1 ? `-${sideIndex + 1}` : ""}`,
            drainType === "catchwater_drain"
                ? "drainCatchwater"
                : drainType === "berm_drain"
                    ? "drainBerm"
                    : drainType === "yard_drain"
                        ? "drainYard"
                        : "drainSide",
            `<LineString>
                <extrude>0</extrude>
                <altitudeMode>absolute</altitudeMode>
                <coordinates>${build3DSpanLineCoordinates(
                    points,
                    startChOffset,
                    startCh,
                    endCh,
                    () => xOffset,
                    (sample) => _safeNum(sample.groundLevel, 0) + 0.12,
                    16,
                )}</coordinates>
            </LineString>`,
            `${drainType.replace(/_/g, " ")} exported from the earthwork structures sheet.`,
        )).join("");
    }).join("");

    return `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
    <Document>
        <name>${projectName}</name>
        <Style id="alignmentLine3d">
            <LineStyle>
                <color>ff00ffff</color>
                <width>5</width>
            </LineStyle>
        </Style>
        <Style id="alignmentLineGround">
            <LineStyle>
                <color>ff9dfcff</color>
                <width>3</width>
            </LineStyle>
        </Style>
        <Style id="terrainLine">
            <LineStyle>
                <color>ff5f708a</color>
                <width>2</width>
            </LineStyle>
        </Style>
        <Style id="corridorDeck">
            <LineStyle>
                <color>ccffffff</color>
                <width>1.5</width>
            </LineStyle>
            <PolyStyle>
                <color>664fc3f7</color>
            </PolyStyle>
        </Style>
        <Style id="terrainSurface">
            <LineStyle>
                <color>99698f45</color>
                <width>1</width>
            </LineStyle>
            <PolyStyle>
                <color>44587d39</color>
            </PolyStyle>
        </Style>
        <Style id="embankmentSurface">
            <LineStyle>
                <color>9922c55e</color>
                <width>1</width>
            </LineStyle>
            <PolyStyle>
                <color>5522c55e</color>
            </PolyStyle>
        </Style>
        <Style id="cuttingSurface">
            <LineStyle>
                <color>99f43f5e</color>
                <width>1</width>
            </LineStyle>
            <PolyStyle>
                <color>55f43f5e</color>
            </PolyStyle>
        </Style>
        <Style id="loopTrackLine">
            <LineStyle>
                <color>ffe2e8f0</color>
                <width>2.2</width>
            </LineStyle>
        </Style>
        <Style id="platformDeck">
            <LineStyle>
                <color>ccb0c4de</color>
                <width>1.2</width>
            </LineStyle>
            <PolyStyle>
                <color>6694a3b8</color>
            </PolyStyle>
        </Style>
        <Style id="bridgeSurface">
            <LineStyle>
                <color>cc94a3b8</color>
                <width>1.2</width>
            </LineStyle>
            <PolyStyle>
                <color>66475569</color>
            </PolyStyle>
        </Style>
        <Style id="tunnelSurface">
            <LineStyle>
                <color>cccbd5e1</color>
                <width>1.2</width>
            </LineStyle>
            <PolyStyle>
                <color>6694a3b8</color>
            </PolyStyle>
        </Style>
        <Style id="retainingWallLine">
            <LineStyle>
                <color>ffd97706</color>
                <width>3</width>
            </LineStyle>
        </Style>
        <Style id="drainSide">
            <LineStyle>
                <color>ff60a5fa</color>
                <width>2</width>
            </LineStyle>
        </Style>
        <Style id="drainYard">
            <LineStyle>
                <color>ff22c55e</color>
                <width>2</width>
            </LineStyle>
        </Style>
        <Style id="drainCatchwater">
            <LineStyle>
                <color>ff06b6d4</color>
                <width>2.2</width>
            </LineStyle>
        </Style>
        <Style id="drainBerm">
            <LineStyle>
                <color>ffa855f7</color>
                <width>2</width>
            </LineStyle>
        </Style>
        <Folder>
            <name>Alignment</name>
            <Placemark>
                <name>${projectName} 3D Alignment</name>
                <description><![CDATA[Exported from Earthsoft with projected formation RL along the imported alignment.]]></description>
                <styleUrl>#alignmentLine3d</styleUrl>
                <LineString>
                    <extrude>${has3DLevels ? 1 : 0}</extrude>
                    <tessellate>0</tessellate>
                    <altitudeMode>${has3DLevels ? "absolute" : "clampToGround"}</altitudeMode>
                    <coordinates>${lineCoordinates}</coordinates>
                </LineString>
            </Placemark>
            <Placemark>
                <name>${projectName} Surface Trace</name>
                <styleUrl>#alignmentLineGround</styleUrl>
                <LineString>
                    <extrude>0</extrude>
                    <tessellate>1</tessellate>
                    <altitudeMode>clampToGround</altitudeMode>
                    <coordinates>${samples.map((sample) => `${sample.lng},${sample.lat},0`).join(" ")}</coordinates>
                </LineString>
            </Placemark>
            <Placemark>
                <name>${projectName} Ground Profile</name>
                <styleUrl>#terrainLine</styleUrl>
                <LineString>
                    <extrude>0</extrude>
                    <tessellate>0</tessellate>
                    <altitudeMode>${has3DLevels ? "absolute" : "clampToGround"}</altitudeMode>
                    <coordinates>${terrainCoordinates}</coordinates>
                </LineString>
            </Placemark>
            ${corridorPolygonCoordinates ? `
            <Placemark>
                <name>${projectName} Corridor Deck</name>
                <description><![CDATA[Formation-width deck derived from the Earthsoft 3D corridor envelope.]]></description>
                <styleUrl>#corridorDeck</styleUrl>
                <Polygon>
                    <extrude>${has3DLevels ? 1 : 0}</extrude>
                    <altitudeMode>${has3DLevels ? "absolute" : "clampToGround"}</altitudeMode>
                    <outerBoundaryIs>
                        <LinearRing>
                            <coordinates>${corridorPolygonCoordinates}</coordinates>
                    </LinearRing>
                </outerBoundaryIs>
            </Polygon>
        </Placemark>` : ""}
        </Folder>
        <Folder>
            <name>${escapeHtml("Terrain & Earthwork")}</name>
            ${terrainRibbonCoordinates ? buildKmlPlacemark(
                `${projectName} Terrain Ribbon`,
                "terrainSurface",
                `<Polygon>
                    <extrude>0</extrude>
                    <altitudeMode>absolute</altitudeMode>
                    <outerBoundaryIs>
                        <LinearRing>
                            <coordinates>${terrainRibbonCoordinates}</coordinates>
                        </LinearRing>
                    </outerBoundaryIs>
                </Polygon>`,
                "Ground ribbon exported along the modeled corridor.",
            ) : ""}
            ${embankmentPlacemarkGroup}
            ${cuttingPlacemarkGroup}
        </Folder>
        <Folder>
            <name>${escapeHtml("Track & Platforms")}</name>
            ${loopTrackPlacemarks}
            ${platformPlacemarks}
        </Folder>
        <Folder>
            <name>Stations</name>
            ${stationPlacemarks}
        </Folder>
        <Folder>
            <name>Structures</name>
            ${bridgePlacemarks}
            ${bridgeSurfacePlacemarks}
            ${retainingWallPlacemarks}
            ${drainPlacemarks}
        </Folder>
    </Document>
</kml>`;
}

function buildAlignmentKmzReadme() {
    return [
        "Earthsoft 3D Alignment Export",
        "",
        "This KMZ contains:",
        "- doc.kml: 3D alignment, surface trace, ground profile, corridor deck, terrain ribbon, embankment and cutting faces, loop lines, platform decks, and structures",
        "",
        "Recommended use:",
        "1. Open Google Earth Web or Google Earth Pro.",
        "2. Import this KMZ file.",
        "3. Enable the Alignment, Terrain & Earthwork, Track & Platforms, and Structures folders.",
        "4. Tilt the camera for the full 3D corridor view.",
        "",
        "Notes:",
        "- 3D heights are based on Earthsoft proposed formation levels.",
        "- Surface trace is included as a clamp-to-ground fallback for review.",
    ].join("\n");
}

async function downloadAlignmentKml() {
    const kmlText = buildAlignmentKmlDocument();
    if (!kmlText) {
        alert("Import a KML alignment first.");
        return;
    }

    const fileBase = String(state.project?.name || "earthsoft-alignment")
        .trim()
        .replace(/[^a-z0-9]+/gi, "-")
        .replace(/^-+|-+$/g, "")
        .toLowerCase() || "earthsoft-alignment";

    let blob;
    let extension = "kmz";

    if (typeof JSZip !== "undefined") {
        const zip = new JSZip();
        zip.file("doc.kml", kmlText);
        zip.file("README.txt", buildAlignmentKmzReadme());
        blob = await zip.generateAsync({
            type: "blob",
            compression: "DEFLATE",
            compressionOptions: { level: 6 },
        });
    } else {
        blob = new Blob([kmlText], { type: "application/vnd.google-earth.kml+xml;charset=utf-8" });
        extension = "kml";
        alert("JSZip is unavailable, so Earthsoft exported a .kml instead of a .kmz.");
    }

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${fileBase}.${extension}`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
}

function drawPseudo3DOverlay(points, startChOffset, settings) {
    const samples = buildPseudo3DSamples(points, startChOffset);
    if (!samples.length) return;

    if (settings.terrain) {
        drawPseudo3DRibbon(samples, {
            pane: "mapPseudo3dTerrainPane",
            fillColor: "#587d39",
            fillOpacity: 0.12,
            color: "#587d39",
            weight: 1,
            opacity: 0.18,
            xMin: (sample) => sample.centerOffset - (sample.currentFormationW / 2) - 18,
            xMax: (sample) => sample.centerOffset + (sample.currentFormationW / 2) + 18,
        });
    }

    if (settings.earthwork) {
        drawPseudo3DEarthwork(samples, "fill");
        drawPseudo3DEarthwork(samples, "cut");
    }

    if (settings.corridor) {
        drawPseudo3DRibbon(samples, {
            pane: "mapPseudo3dCorridorPane",
            fillColor: "#475569",
            fillOpacity: 0.42,
            color: "#94a3b8",
            weight: 1.5,
            opacity: 0.58,
            xMin: (sample) => sample.centerOffset - (sample.currentFormationW / 2) - 0.7,
            xMax: (sample) => sample.centerOffset + (sample.currentFormationW / 2) + 0.7,
        });
        drawPseudo3DTrackRails(samples);
        drawPseudo3DPlatforms(points, startChOffset);
    }

    if (settings.structures) {
        drawPseudo3DBridgeAndTunnelOverlays(points, startChOffset);
    }

    if (settings.ewStructures) {
        drawPseudo3DEarthworkStructures(points, startChOffset);
    }

    if (settings.labels) {
        drawPseudo3DLabels(points, startChOffset);
    }
}

function buildPseudo3DSamples(points, startChOffset) {
    const rows = Array.isArray(state.calcRows) ? state.calcRows : [];
    if (!rows.length || !points?.length) return [];

    const standardTc = _safeNum(state.settings?.visual?.minTc, 5.3);
    const formationW = _safeNum(state.settings?.visual?.formationWidthFill, 7.85);
    const sampleWindow = getTypicalChainageStep(rows);
    const stride = Math.max(1, Math.ceil(rows.length / 850));
    const sampleIndexes = new Set([0, rows.length - 1]);

    for (let i = 0; i < rows.length; i += stride) sampleIndexes.add(i);

    const addChainageIndex = (value) => {
        const parsed = _safeNum(value, NaN);
        if (!Number.isFinite(parsed) || typeof findNearestCalcRowIndexByChainage !== "function") return;
        sampleIndexes.add(findNearestCalcRowIndexByChainage(parsed));
    };

    (state.bridgeRows || []).forEach((row) => {
        addChainageIndex(row?.startChainage);
        addChainageIndex(row?.endChainage);
    });
    (state.earthworkStructures?.retainingWalls || []).forEach((row) => {
        addChainageIndex(row?.fromCh);
        addChainageIndex(row?.toCh);
    });
    (state.earthworkStructures?.drains || []).forEach((row) => {
        addChainageIndex(row?.fromCh);
        addChainageIndex(row?.toCh);
    });
    (state.loopPlatformRows || []).forEach((row) => {
        addChainageIndex(row?.loopStartCh);
        addChainageIndex(row?.loopEndCh);
        addChainageIndex(row?.pfStartCh);
        addChainageIndex(row?.pfEndCh);
    });

    return Array.from(sampleIndexes)
        .sort((a, b) => a - b)
        .map((index) => {
            const row = rows[index];
            if (!row) return null;
            const chainageAbs = _safeNum(row.chainage, NaN);
            if (!Number.isFinite(chainageAbs)) return null;
            const alignmentCh = chainageAbs - startChOffset;
            const center = getLatLngFromChainage(alignmentCh, points);
            const tangent = getAlignmentTangentAtChainage(alignmentCh, points, sampleWindow);
            if (!center || !tangent) return null;
            const envelope = typeof compute3DEnvelopeForRow === "function"
                ? compute3DEnvelopeForRow(row, standardTc, formationW)
                : {
                    trackItems: [{ order: 0, offset: 0, side: "", isMain: true, row: {} }],
                    platformItems: [],
                    minX: -(formationW / 2),
                    maxX: formationW / 2,
                    currentFormationW: formationW,
                    centerOffset: 0,
                };
            return { index, row, chainageAbs, center, tangent, envelope, centerOffset: envelope.centerOffset, currentFormationW: envelope.currentFormationW };
        })
        .filter(Boolean);
}

function projectPseudo3DPoint(sample, x) {
    return offsetLatLng(sample.center, sample.tangent, -x);
}

function drawPseudo3DRibbon(samples, options) {
    const left = [];
    const right = [];
    samples.forEach((sample) => {
        const xMin = options.xMin(sample);
        const xMax = options.xMax(sample);
        if (!Number.isFinite(xMin) || !Number.isFinite(xMax)) return;
        left.push(projectPseudo3DPoint(sample, xMin));
        right.push(projectPseudo3DPoint(sample, xMax));
    });
    if (left.length < 2 || right.length < 2) return;

    L.polygon(
        [...left, ...right.slice().reverse()].map((pt) => [pt.lat, pt.lng]),
        {
            pane: options.pane,
            color: options.color,
            weight: options.weight,
            opacity: options.opacity,
            fillColor: options.fillColor,
            fillOpacity: options.fillOpacity,
            interactive: false,
        },
    ).addTo(mapItems);
}

function buildContiguousSampleGroups(samples, predicate) {
    const groups = [];
    let current = [];
    samples.forEach((sample) => {
        if (predicate(sample)) {
            current.push(sample);
            return;
        }
        if (current.length > 1) groups.push(current);
        current = [];
    });
    if (current.length > 1) groups.push(current);
    return groups;
}

function drawPseudo3DEarthwork(samples, kind) {
    const isFill = kind === "fill";
    const groups = buildContiguousSampleGroups(samples, (sample) => isFill ? _safeNum(sample.row?.bank, 0) > 0.1 : _safeNum(sample.row?.cut, 0) > 0.1);
    groups.forEach((group) => {
        drawPseudo3DRibbon(group, {
            pane: "mapPseudo3dEarthworkPane",
            fillColor: isFill ? "#22c55e" : "#f43f5e",
            fillOpacity: isFill ? 0.18 : 0.14,
            color: isFill ? "#22c55e" : "#f43f5e",
            weight: 1,
            opacity: 0.28,
            xMin: (sample) => sample.centerOffset - (_safeNum(isFill ? sample.row?.fillBottom : sample.row?.cutBottom, sample.currentFormationW) / 2),
            xMax: (sample) => sample.centerOffset + (_safeNum(isFill ? sample.row?.fillBottom : sample.row?.cutBottom, sample.currentFormationW) / 2),
        });
    });
}

function drawPseudo3DTrackRails(samples) {
    const railSegments = new Map();

    samples.forEach((sample) => {
        (sample.envelope?.trackItems || []).forEach((track) => {
            const key = `${track.order}|${track.side || "center"}|${track.isMain ? "main" : "aux"}`;
            const entry = railSegments.get(key) || { left: [], right: [] };
            entry.left.push(projectPseudo3DPoint(sample, _safeNum(track.offset, 0) - 0.835));
            entry.right.push(projectPseudo3DPoint(sample, _safeNum(track.offset, 0) + 0.835));
            railSegments.set(key, entry);
        });
    });

    railSegments.forEach((entry) => {
        if (entry.left.length > 1) {
            L.polyline(entry.left.map((pt) => [pt.lat, pt.lng]), {
                pane: "mapPseudo3dRailPane",
                color: "#e2e8f0",
                weight: 1.35,
                opacity: 0.95,
                interactive: false,
            }).addTo(mapItems);
        }
        if (entry.right.length > 1) {
            L.polyline(entry.right.map((pt) => [pt.lat, pt.lng]), {
                pane: "mapPseudo3dRailPane",
                color: "#e2e8f0",
                weight: 1.35,
                opacity: 0.95,
                interactive: false,
            }).addTo(mapItems);
        }
    });
}

function drawPseudo3DPlatforms(points, startChOffset) {
    const standardTc = _safeNum(state.settings?.visual?.minTc, 5.3);
    (state.loopPlatformRows || []).forEach((platformRow) => {
        const pfStartCh = _safeNum(platformRow?.pfStartCh, NaN);
        const pfEndCh = _safeNum(platformRow?.pfEndCh, NaN);
        const pfWidth = _safeNum(platformRow?.pfWidth, NaN);
        if (!Number.isFinite(pfStartCh) || !Number.isFinite(pfEndCh) || !(pfWidth > 0)) return;

        const midCh = (pfStartCh + pfEndCh) / 2;
        const layout = typeof buildStationSequenceLayout === "function"
            ? buildStationSequenceLayout(midCh, platformRow?.station || "", standardTc, { useRanges: true })
            : null;
        const trackItems = layout
            ? layout.trackItems.map((item) => ({ ...item, offset: layout.offsetByItem.get(item) }))
            : [{ offset: 0, side: "", isMain: true }];
        const isLeft = String(platformRow?.side || "").toLowerCase() === "left";
        const refTrack = trackItems.find((track) => track.side === platformRow?.side)
            || trackItems.find((track) => track.isMain)
            || trackItems[0];
        const cx = _safeNum(refTrack?.offset, 0) + (isLeft ? -(2.8 + (pfWidth / 2)) : (2.8 + (pfWidth / 2)));
        const xMin = cx - (pfWidth / 2);
        const xMax = cx + (pfWidth / 2);
        const polygon = buildPseudo3DSpanPolygon(points, startChOffset, pfStartCh, pfEndCh, xMin, xMax);
        if (!polygon) return;
        L.polygon(polygon, {
            pane: "mapPseudo3dCorridorPane",
            color: "#cbd5e1",
            weight: 1,
            opacity: 0.55,
            fillColor: "#94a3b8",
            fillOpacity: 0.32,
            interactive: false,
        }).addTo(mapItems);
    });
}

function buildPseudo3DSpanPolygon(points, startChOffset, startCh, endCh, xMin, xMax) {
    const chStart = Math.min(startCh, endCh);
    const chEnd = Math.max(startCh, endCh);
    const startSample = buildPseudo3DChainageSample(chStart, points, startChOffset);
    const endSample = buildPseudo3DChainageSample(chEnd, points, startChOffset);
    if (!startSample || !endSample) return null;
    return [
        [projectPseudo3DPoint(startSample, xMin).lat, projectPseudo3DPoint(startSample, xMin).lng],
        [projectPseudo3DPoint(endSample, xMin).lat, projectPseudo3DPoint(endSample, xMin).lng],
        [projectPseudo3DPoint(endSample, xMax).lat, projectPseudo3DPoint(endSample, xMax).lng],
        [projectPseudo3DPoint(startSample, xMax).lat, projectPseudo3DPoint(startSample, xMax).lng],
    ];
}

function buildPseudo3DChainageSample(chainageAbs, points, startChOffset) {
    const alignmentCh = chainageAbs - startChOffset;
    const center = getLatLngFromChainage(alignmentCh, points);
    const tangent = getAlignmentTangentAtChainage(alignmentCh, points, getTypicalChainageStep(state.calcRows || []));
    if (!center || !tangent) return null;
    return { center, tangent };
}

function drawPseudo3DBridgeAndTunnelOverlays(points, startChOffset) {
    const standardTc = _safeNum(state.settings?.visual?.minTc, 5.3);
    const formationW = _safeNum(state.settings?.visual?.formationWidthFill, 7.85);
    (state.bridgeRows || []).forEach((bridge) => {
        const startCh = _safeNum(bridge?.startChainage, NaN);
        const endCh = _safeNum(bridge?.endChainage, NaN);
        if (!Number.isFinite(startCh) || !Number.isFinite(endCh)) return;
        const midCh = (startCh + endCh) / 2;
        const rowIndex = typeof findNearestCalcRowIndexByChainage === "function" ? findNearestCalcRowIndexByChainage(midCh) : 0;
        const row = state.calcRows?.[rowIndex];
        const envelope = row && typeof compute3DEnvelopeForRow === "function"
            ? compute3DEnvelopeForRow(row, standardTc, formationW)
            : { currentFormationW: formationW, centerOffset: 0 };
        const width = _safeNum(envelope.currentFormationW, formationW) + 2;
        const polygon = buildPseudo3DSpanPolygon(
            points,
            startChOffset,
            startCh,
            endCh,
            _safeNum(envelope.centerOffset, 0) - (width / 2),
            _safeNum(envelope.centerOffset, 0) + (width / 2),
        );
        if (!polygon) return;

        const isTunnel = /tunnel/i.test(String(bridge?.bridgeCategory || ""));
        L.polygon(polygon, {
            pane: "mapPseudo3dStructurePane",
            color: isTunnel ? "#cbd5e1" : "#475569",
            weight: 1.5,
            opacity: 0.85,
            fillColor: isTunnel ? "#94a3b8" : "#64748b",
            fillOpacity: isTunnel ? 0.34 : 0.28,
            dashArray: isTunnel ? "8 6" : null,
            interactive: false,
        }).addTo(mapItems);
    });
}

function drawPseudo3DEarthworkStructures(points, startChOffset) {
    const standardTc = _safeNum(state.settings?.visual?.minTc, 5.3);
    const formationW = _safeNum(state.settings?.visual?.formationWidthFill, 7.85);

    (state.earthworkStructures?.retainingWalls || []).forEach((wall) => {
        const fromCh = _safeNum(wall?.fromCh, NaN);
        const toCh = _safeNum(wall?.toCh, NaN);
        if (!Number.isFinite(fromCh) || !Number.isFinite(toCh)) return;
        const midCh = (fromCh + toCh) / 2;
        const rowIndex = typeof findNearestCalcRowIndexByChainage === "function" ? findNearestCalcRowIndexByChainage(midCh) : 0;
        const row = state.calcRows?.[rowIndex];
        const envelope = row && typeof compute3DEnvelopeForRow === "function"
            ? compute3DEnvelopeForRow(row, standardTc, formationW)
            : { currentFormationW: formationW, centerOffset: 0 };
        const xBase = _safeNum(envelope.centerOffset, 0);
        const halfWidth = _safeNum(envelope.currentFormationW, formationW) / 2;
        const offsets = String(wall?.side || "").toLowerCase() === "both"
            ? [xBase - (halfWidth + 1.2), xBase + (halfWidth + 1.2)]
            : [xBase + ((String(wall?.side || "").toLowerCase() === "left" ? -1 : 1) * (halfWidth + 1.2))];
        offsets.forEach((x) => {
            const line = buildPseudo3DSpanLine(points, startChOffset, fromCh, toCh, x);
            if (!line) return;
            L.polyline(line, {
                pane: "mapPseudo3dStructurePane",
                color: "#d97706",
                weight: String(wall?.wallType || "").includes("4") ? 5 : 4,
                opacity: 0.92,
                interactive: false,
            }).addTo(mapItems);
        });
    });

    (state.earthworkStructures?.drains || []).forEach((drain) => {
        const fromCh = _safeNum(drain?.fromCh, NaN);
        const toCh = _safeNum(drain?.toCh, NaN);
        if (!Number.isFinite(fromCh) || !Number.isFinite(toCh)) return;
        const midCh = (fromCh + toCh) / 2;
        const rowIndex = typeof findNearestCalcRowIndexByChainage === "function" ? findNearestCalcRowIndexByChainage(midCh) : 0;
        const row = state.calcRows?.[rowIndex];
        const envelope = row && typeof compute3DEnvelopeForRow === "function"
            ? compute3DEnvelopeForRow(row, standardTc, formationW)
            : { currentFormationW: formationW, centerOffset: 0 };
        const xBase = _safeNum(envelope.centerOffset, 0);
        const halfWidth = _safeNum(envelope.currentFormationW, formationW) / 2;
        const drainType = typeof getDrainTypeMeta === "function" ? getDrainTypeMeta(drain?.drainType).key : String(drain?.drainType || "side_drain");
        let sideOffset = 1.6;
        if (drainType === "catchwater_drain") sideOffset = 4.5;
        if (drainType === "berm_drain") sideOffset = 3.2;
        if (drainType === "yard_drain") sideOffset = 2.2;
        const drainColor = drainType === "yard_drain"
            ? "#22c55e"
            : drainType === "catchwater_drain"
                ? "#06b6d4"
                : drainType === "berm_drain"
                    ? "#a855f7"
                    : "#60a5fa";
        const offsets = String(drain?.side || "").toLowerCase() === "both"
            ? [xBase - (halfWidth + sideOffset), xBase + (halfWidth + sideOffset)]
            : [xBase + ((String(drain?.side || "").toLowerCase() === "left" ? -1 : 1) * (halfWidth + sideOffset))];
        offsets.forEach((x) => {
            const line = buildPseudo3DSpanLine(points, startChOffset, fromCh, toCh, x);
            if (!line) return;
            L.polyline(line, {
                pane: "mapPseudo3dStructurePane",
                color: drainColor,
                weight: drainType === "catchwater_drain" ? 4 : 3,
                opacity: 0.9,
                dashArray: drainType === "berm_drain" ? "8 6" : null,
                interactive: false,
            }).addTo(mapItems);
        });
    });
}

function buildPseudo3DSpanLine(points, startChOffset, startCh, endCh, x) {
    const chStart = Math.min(startCh, endCh);
    const chEnd = Math.max(startCh, endCh);
    const midCh = (chStart + chEnd) / 2;
    const samples = [chStart, midCh, chEnd]
        .map((chainage) => buildPseudo3DChainageSample(chainage, points, startChOffset))
        .filter(Boolean);
    if (samples.length < 2) return null;
    return samples.map((sample) => {
        const pt = projectPseudo3DPoint(sample, x);
        return [pt.lat, pt.lng];
    });
}

function drawPseudo3DLabels(points, startChOffset) {
    const rows = Array.isArray(state.calcRows) ? state.calcRows : [];
    if (!rows.length) return;
    const startCh = _safeNum(rows[0]?.chainage, 0);
    const endCh = _safeNum(rows[rows.length - 1]?.chainage, startCh);
    const startKm = Math.ceil(startCh / 1000);
    const endKm = Math.floor(endCh / 1000);

    for (let km = startKm; km <= endKm; km += 1) {
        const chainageAbs = km * 1000;
        const pt = getLatLngFromChainage(chainageAbs - startChOffset, points);
        if (!pt) continue;
        L.marker([pt.lat, pt.lng], {
            pane: "mapPseudo3dLabelPane",
            interactive: false,
            icon: L.divIcon({
                className: "map-km-label-wrap",
                html: `<div class="map-km-chip">KM ${km}</div>`,
                iconSize: null,
            }),
        }).addTo(mapItems);
    }
}

// ── Utilities ─────────────────────────────────────────────────────────────

function getDistanceInMeters(p1, p2) {
    const R = 6371e3;
    const dLat = (p2.lat - p1.lat) * Math.PI / 180;
    const dLng = (p2.lng - p1.lng) * Math.PI / 180;
    const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
        Math.cos(p1.lat * Math.PI / 180) * Math.cos(p2.lat * Math.PI / 180) *
        Math.sin(dLng / 2) * Math.sin(dLng / 2);
    return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function getLatLngFromChainage(targetCh, points) {
    if (!points || points.length < 2) return null;
    if (targetCh <= points[0].ch) return points[0];
    if (targetCh >= points[points.length - 1].ch) return points[points.length - 1];

    let low = 0, high = points.length - 1;
    while (low <= high) {
        const mid = Math.floor((low + high) / 2);
        if (points[mid].ch < targetCh) low = mid + 1;
        else if (points[mid].ch > targetCh) high = mid - 1;
        else return points[mid];
    }
    const p1 = points[high], p2 = points[low];
    if (!p1 || !p2) return null;
    const ratio = (targetCh - p1.ch) / (p2.ch - p1.ch);
    return { lat: p1.lat + ratio * (p2.lat - p1.lat), lng: p1.lng + ratio * (p2.lng - p1.lng) };
}

function drawLandBoundaryOverlay(points, startChOffset) {
    const segments = buildLandBoundarySegments(points, startChOffset);
    if (!segments.length) return false;

    segments.forEach((segment) => {
        const left = segment.left.map((pt) => [pt.lat, pt.lng]);
        const right = segment.right.map((pt) => [pt.lat, pt.lng]);

        if (left.length > 2 && right.length > 2) {
            L.polygon([...left, ...right.slice().reverse()], {
                color: '#f97316',
                weight: 1,
                opacity: 0.18,
                fillColor: '#f97316',
                fillOpacity: 0.08,
                interactive: false,
            }).addTo(mapItems);
        }

        if (left.length > 1) {
            L.polyline(left, {
                color: '#fb923c',
                weight: 2,
                opacity: 0.78,
                dashArray: '8 6',
                interactive: false,
            }).addTo(mapItems);
        }

        if (right.length > 1) {
            L.polyline(right, {
                color: '#fb923c',
                weight: 2,
                opacity: 0.78,
                dashArray: '8 6',
                interactive: false,
            }).addTo(mapItems);
        }
    });

    return true;
}

function buildLandBoundarySegments(points, startChOffset) {
    const rows = Array.isArray(state.calcRows) ? state.calcRows : [];
    if (rows.length < 2 || !points || points.length < 2) return [];

    const bridges = Array.isArray(state.bridgeRows) ? state.bridgeRows : [];
    const laSettings = getMapLASettings();
    const sampleWindow = getTypicalChainageStep(rows);
    const segments = [];
    let currentSegment = null;

    const flushSegment = () => {
        if (!currentSegment) return;
        if (currentSegment.left.length > 1 && currentSegment.right.length > 1) {
            segments.push(currentSegment);
        }
        currentSegment = null;
    };

    rows.forEach((row) => {
        const chainageAbs = _safeNum(row?.chainage, NaN);
        if (!Number.isFinite(chainageAbs)) {
            flushSegment();
            return;
        }

        const alignmentCh = chainageAbs - startChOffset;
        const center = getLatLngFromChainage(alignmentCh, points);
        const tangent = getAlignmentTangentAtChainage(alignmentCh, points, sampleWindow);
        const halfWidth = getLandBoundaryHalfWidth(row, chainageAbs, bridges, laSettings);

        if (!center || !tangent || !(halfWidth > 0)) {
            flushSegment();
            return;
        }

        if (!currentSegment) {
            currentSegment = { left: [], right: [] };
        }

        currentSegment.left.push(offsetLatLng(center, tangent, halfWidth));
        currentSegment.right.push(offsetLatLng(center, tangent, -halfWidth));
    });

    flushSegment();
    return segments;
}

function getMapLASettings() {
    if (typeof getLASettings === "function") {
        return getLASettings();
    }
    return {
        offsetFromToe: _safeNum(state.settings?.laOffsetFromToe, 10),
        bridgeExtraWidth: _safeNum(state.settings?.laBridgeExtraWidth, 5),
    };
}

function getLandBoundaryHalfWidth(row, chainageAbs, bridges, laSettings) {
    const bridge = (bridges || []).find((item) => chainageAbs >= _safeNum(item?.startChainage, NaN) && chainageAbs <= _safeNum(item?.endChainage, NaN));
    const bridgeCategories = typeof LA_BRIDGE_LAND_CATEGORIES !== "undefined"
        ? LA_BRIDGE_LAND_CATEGORIES
        : ["Major", "ROB", "RUB", "Viaduct"];

    if (bridge) {
        const category = String(bridge.bridgeCategory || "").trim();
        if (/tunnel/i.test(category)) return 0;
        if (bridgeCategories.some((item) => item.toLowerCase() === category.toLowerCase()) ||
            /\brob\b/i.test(String(bridge.bridgeNo || "")) ||
            /\brub\b/i.test(String(bridge.bridgeNo || ""))) {
            const bridgeWidth = _safeNum(state.settings?.formationWidthFill, 7.85);
            return (bridgeWidth + (2 * _safeNum(laSettings?.bridgeExtraWidth, 5))) / 2;
        }
    }

    const topWidth = _safeNum(row?.topWidth, _safeNum(state.settings?.formationWidthFill, 7.85));
    const fillBottom = _safeNum(row?.fillBottom, topWidth);
    const cutBottom = _safeNum(row?.cutBottom, _safeNum(topWidth, _safeNum(state.settings?.cuttingWidth, topWidth)));
    const toeWidth = row?.bank > 0.001
        ? fillBottom
        : row?.cut > 0.001
            ? cutBottom
            : topWidth;

    return (toeWidth / 2) + _safeNum(laSettings?.offsetFromToe, 10);
}

function getTypicalChainageStep(rows) {
    if (!Array.isArray(rows) || rows.length < 2) return 5;
    for (let i = 1; i < rows.length; i++) {
        const prev = _safeNum(rows[i - 1]?.chainage, NaN);
        const next = _safeNum(rows[i]?.chainage, NaN);
        const delta = Math.abs(next - prev);
        if (delta > 0.001) {
            return Math.max(2, Math.min(25, delta / 2));
        }
    }
    return 5;
}

function getAlignmentTangentAtChainage(targetCh, points, sampleWindow = 5) {
    if (!points || points.length < 2) return null;
    const minCh = points[0].ch;
    const maxCh = points[points.length - 1].ch;
    const prev = getLatLngFromChainage(Math.max(minCh, targetCh - sampleWindow), points);
    const next = getLatLngFromChainage(Math.min(maxCh, targetCh + sampleWindow), points);
    if (!prev || !next) return null;

    const metersPerDegLat = 111320;
    const metersPerDegLng = Math.max(Math.cos(((prev.lat + next.lat) / 2) * Math.PI / 180) * metersPerDegLat, 1e-6);
    const east = (next.lng - prev.lng) * metersPerDegLng;
    const north = (next.lat - prev.lat) * metersPerDegLat;
    const length = Math.hypot(east, north);
    if (!length) return null;

    return { east: east / length, north: north / length };
}

function offsetLatLng(point, tangent, offsetMeters) {
    const metersPerDegLat = 111320;
    const metersPerDegLng = Math.max(Math.cos(point.lat * Math.PI / 180) * metersPerDegLat, 1e-6);
    const normalEast = -tangent.north;
    const normalNorth = tangent.east;

    return {
        lat: point.lat + ((normalNorth * offsetMeters) / metersPerDegLat),
        lng: point.lng + ((normalEast * offsetMeters) / metersPerDegLng),
    };
}

function escapeHtml(value) {
    return String(value || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

function escapeHtmlAttr(value) {
    return escapeHtml(value);
}

function updateMapLegend(stationPlanCount, hasLandBoundary, pseudo3DEnabled = false) {
    const noteEl = document.getElementById("alignmentMapLegendNote");
    if (!noteEl) return;
    const boundaryNote = hasLandBoundary ? " Orange dashed lines show the land acquisition boundary." : "";
    const overlayNote = pseudo3DEnabled ? " Pseudo-3D overlay is active for corridor, structures, and chainage labels." : "";
    if (stationPlanCount > 0) {
        noteEl.textContent = `${stationPlanCount} station plan${stationPlanCount === 1 ? "" : "s"} mapped. Station labels are shown once per station at CSB.${boundaryNote}${overlayNote}`;
    } else {
        noteEl.textContent = `Station locations are shown after KML import. Station labels are shown once per station at CSB.${boundaryNote}${overlayNote}`;
    }
}

function openStationPlanDialog(stationChoices) {
    const modal = document.getElementById("stationPlanModal");
    const select = document.getElementById("stationPlanStationSelect");
    const text = document.getElementById("stationPlanModalText");
    if (!modal || !select) {
        const fallbackStation = pendingStationPlanTarget || stationChoices[0] || "";
        if (fallbackStation) attachPendingStationPlan(fallbackStation);
        return;
    }

    select.innerHTML = stationChoices.map((station) => {
        const safeLabel = escapeHtml(station);
        const selected = pendingStationPlanTarget && station.toLowerCase().trim() === pendingStationPlanTarget.toLowerCase().trim()
            ? " selected"
            : "";
        return `<option value="${safeLabel}"${selected}>${safeLabel}</option>`;
    }).join("");

    if (text) {
        text.textContent = pendingStationPlanFile
            ? `Selected file: ${pendingStationPlanFile.name}. Choose the station this conceptual plan belongs to.`
            : "Select the station this conceptual plan belongs to.";
    }

    modal.showModal();
}

function closeStationPlanDialog() {
    const modal = document.getElementById("stationPlanModal");
    if (!modal) {
        clearPendingStationPlan();
        return;
    }
    modal.close();
}

function clearPendingStationPlan() {
    pendingStationPlanTarget = "";
    pendingStationPlanFile = null;
}

function attachPendingStationPlan(stationName) {
    const file = pendingStationPlanFile;
    if (!file) {
        alert("No station plan file is pending.");
        closeStationPlanDialog();
        return;
    }

    const normalizedStationName = String(stationName || "").toLowerCase().trim();
    const stationChoices = [...new Set(state.loopPlatformRows.map((r) => String(r.station || "").trim()).filter(Boolean))];
    const matchedStation = stationChoices.find((name) => name.toLowerCase().trim() === normalizedStationName);
    if (!matchedStation) {
        alert("That station name does not match the imported loop/station list.");
        return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
        state.stationPlans[normalizedStationName] = evt.target.result;
        saveState();
        if (state.kmlData) drawAlignmentMap();
        renderGoogleEarthPage();

        const modal = document.getElementById("stationPlanModal");
        if (modal) {
            modal.returnValue = "confirmed";
            modal.close();
        }
        alert(`Attached station plan to ${matchedStation}. Click its station label or cyan marker on the map to view it.`);
        clearPendingStationPlan();
    };
    reader.readAsDataURL(file);
}
