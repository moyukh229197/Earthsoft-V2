/**
 * alignment-map.js
 * Handles the Leaflet.js (Free Map) integration for geographic alignment and earthwork visual overlays.
 */

let alignmentMap = null;
let baseLayers = {};
let mapItems = null; // We'll use a FeatureGroup for easy clearing
let pendingStationPlanTarget = "";
let pendingStationPlanFile = null;

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

    if (importKmlBtn && kmlImportInput) {
        importKmlBtn.addEventListener("click", () => kmlImportInput.click());
        kmlImportInput.addEventListener("change", handleKmlImport);
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
        });
    });

    document.addEventListener("click", (event) => {
        const uploadBtn = event.target.closest("[data-upload-station-plan]");
        if (!uploadBtn || !stationPlanImportInput) return;
        pendingStationPlanTarget = String(uploadBtn.dataset.uploadStationPlan || "").trim();
        if (!pendingStationPlanTarget) return;
        stationPlanImportInput.click();
    });
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

    // Base Alignment Polyline
    L.polyline(path, {
        color: '#ffffff',
        weight: 3,
        opacity: 0.8
    }).addTo(mapItems);

    const hasLandBoundary = drawLandBoundaryOverlay(points, startChOffset);
    const showMapStations = state.settings?.showMapStations !== false;
    const showMapBridges = state.settings?.showMapBridges !== false;
    const showMapCurves = state.settings?.showMapCurves !== false;

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

    updateMapLegend(Object.keys(state.stationPlans || {}).length, hasLandBoundary);
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

function updateMapLegend(stationPlanCount, hasLandBoundary) {
    const noteEl = document.getElementById("alignmentMapLegendNote");
    if (!noteEl) return;
    const boundaryNote = hasLandBoundary ? " Orange dashed lines show the land acquisition boundary." : "";
    if (stationPlanCount > 0) {
        noteEl.textContent = `${stationPlanCount} station plan${stationPlanCount === 1 ? "" : "s"} mapped. Station labels are shown once per station at CSB.${boundaryNote}`;
    } else {
        noteEl.textContent = `Station locations are shown after KML import. Station labels are shown once per station at CSB.${boundaryNote}`;
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
