/**
 * alignment-map.js
 * Handles the Leaflet map integration for geographic alignment and earthwork visual overlays.
 */

let alignmentMap = null;
let alignmentPolyline = null;
let alignmentLayerGroup = null;

const MAP_TILES = {
    streets: 'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
    satellite: 'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',
    terrain: 'https://server.arcgisonline.com/ArcGIS/rest/services/World_Topo_Map/MapServer/tile/{z}/{y}/{x}'
};
let currentTileLayer = null;

// Initialize Map System once DOM is loaded
document.addEventListener("DOMContentLoaded", () => {
    // Wait for the app to init and Leaflet to load
    setTimeout(initMapSystem, 500);
});

function initMapSystem() {
    if (typeof L === 'undefined') return;

    // Setup DOM Event Listeners
    if (els.importKmlBtn) els.importKmlBtn.addEventListener("click", () => els.kmlImportInput.click());
    if (els.kmlImportInput) els.kmlImportInput.addEventListener("change", handleKmlImport);
    if (els.importStationPlanBtn) els.importStationPlanBtn.addEventListener("click", () => els.stationPlanImportInput.click());
    if (els.stationPlanImportInput) els.stationPlanImportInput.addEventListener("change", handleStationPlanImport);
    if (els.mapTileSelect) els.mapTileSelect.addEventListener("change", (e) => setMapTile(e.target.value));

    // Initialize Map
    const mapContainer = document.getElementById("alignmentMapContainer");
    const emptyState = document.getElementById("alignmentMapEmpty");

    if (!mapContainer || !emptyState) return;

    alignmentMap = L.map(mapContainer, { zoomControl: false }).setView([20.5937, 78.9629], 5); // Default to India center
    L.control.zoom({ position: 'bottomright' }).addTo(alignmentMap);

    currentTileLayer = L.tileLayer(MAP_TILES.streets, {
        attribution: 'Map data &copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
        maxZoom: 19,
    }).addTo(alignmentMap);

    alignmentLayerGroup = L.layerGroup().addTo(alignmentMap);

    // Re-draw map if data exists (on load)
    if (state.kmlData) {
        drawAlignmentMap();
    }

    // Handle map resize on page reveal
    const workNavBtns = document.querySelectorAll(".work-nav-btn");
    workNavBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            if (btn.dataset.workPageBtn === "alignment-map") {
                setTimeout(() => {
                    alignmentMap.invalidateSize();
                    if (state.kmlData) drawAlignmentMap();
                }, 300);
            }
        });
    });
}

function setMapTile(type) {
    if (!alignmentMap || !currentTileLayer) return;
    alignmentMap.removeLayer(currentTileLayer);

    let attribution = '';
    if (type === 'streets') attribution = 'Map data &copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>';
    else attribution = 'Tiles &copy; Esri &mdash; Source: Esri';

    currentTileLayer = L.tileLayer(MAP_TILES[type] || MAP_TILES.streets, {
        attribution: attribution,
        maxZoom: 19
    }).addTo(alignmentMap);
}

// ── KML Import ────────────────────────────────────────────────────────────

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

                // Find a .kml file inside the zip (usually doc.kml)
                let kmlFile = null;
                zip.forEach((relativePath, zipEntry) => {
                    if (relativePath.toLowerCase().endsWith('.kml')) {
                        kmlFile = zipEntry;
                    }
                });

                if (!kmlFile) {
                    alert("Could not find a .kml file inside the KMZ archive.");
                    return;
                }

                const kmlText = await kmlFile.async("text");
                parseKMLData(kmlText);
            } catch (err) {
                console.error("KMZ Extraction Error:", err);
                alert("Failed to read KMZ file. Is it corrupted?");
            } finally {
                e.target.value = ""; // Reset
            }
        };
        reader.readAsArrayBuffer(file);
    } else {
        const reader = new FileReader();
        reader.onload = (evt) => {
            const kmlText = evt.target.result;
            parseKMLData(kmlText);
            e.target.value = ""; // Reset
        };
        reader.readAsText(file);
    }
}

function parseKMLData(kmlText) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(kmlText, "text/xml");

    // Find all LineString coordinates
    const lineStrings = xmlDoc.getElementsByTagName("LineString");
    if (!lineStrings || lineStrings.length === 0) {
        alert("No LineString path found in KML file.");
        return;
    }

    // Merge multiple LineStrings if present, or just use the first one
    let points = [];
    for (let i = 0; i < lineStrings.length; i++) {
        const coordsNode = lineStrings[i].getElementsByTagName("coordinates")[0];
        if (coordsNode) {
            const coordStr = coordsNode.textContent.trim();
            const coordPairs = coordStr.split(/\s+/);
            coordPairs.forEach(pair => {
                const [lng, lat, elev] = pair.split(',');
                if (lat && lng) {
                    points.push({ lat: parseFloat(lat), lng: parseFloat(lng) });
                }
            });
        }
    }

    if (points.length === 0) {
        alert("Could not parse coordinates from the KML file.");
        return;
    }

    // Calculate cumulative distances (chainage representation for the KML line)
    // We use Haversine to get meters distance between points
    let cumDist = 0;
    points[0].ch = 0;
    for (let i = 1; i < points.length; i++) {
        const d = getDistanceInMeters(points[i - 1], points[i]);
        cumDist += d;
        points[i].ch = cumDist;
    }

    state.kmlData = {
        points: points,
        totalDistance: cumDist
    };

    saveState();
    if (document.querySelector('.work-page[data-work-page="alignment-map"]').style.display !== 'none') {
        drawAlignmentMap();
    }
}

// ── Image Import ──────────────────────────────────────────────────────────

function handleStationPlanImport(e) {
    const file = e.target.files[0];
    if (!file) return;

    // Ask user which station loop/platform this belongs to
    if (!state.loopPlatformRows || state.loopPlatformRows.length === 0) {
        alert("Please add Loops & Station data first so you can attach this plan to a specific station element.");
        e.target.value = "";
        return;
    }

    const stationName = prompt("Enter the Station Name or Loop/Platform Description this plan belongs to (e.g. 'Station A Loop'):\n\nAvailable:\n" + state.loopPlatformRows.map(r => r.desc).join('\n'));

    if (!stationName) {
        e.target.value = "";
        return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
        // Store image base64
        state.stationPlans[stationName.toLowerCase().trim()] = evt.target.result;
        saveState();
        alert(`Conceptual plan saved for ${stationName}. It will display on the map when you click the station marker.`);
        if (state.kmlData) drawAlignmentMap();
        e.target.value = "";
    };
    reader.readAsDataURL(file);
}

// ── Map Rendering Core ───────────────────────────────────────────────────

function drawAlignmentMap() {
    if (!alignmentMap || !state.kmlData || !state.kmlData.points.length) {
        // Show empty state
        document.getElementById("alignmentMapContainer").style.display = "none";
        document.getElementById("alignmentMapEmpty").style.display = "flex";
        return;
    }

    // Hide empty state, show map
    document.getElementById("alignmentMapContainer").style.display = "block";
    document.getElementById("alignmentMapEmpty").style.display = "none";
    alignmentMap.invalidateSize();

    const points = state.kmlData.points;
    const latlngs = points.map(p => [p.lat, p.lng]);

    alignmentLayerGroup.clearLayers();

    // Draw Base Polyline representing KML Alignment
    const baseAlign = L.polyline(latlngs, {
        color: '#ffffff',
        weight: 3,
        opacity: 0.3,
        dashArray: '6, 6'
    }).addTo(alignmentLayerGroup);

    alignmentMap.fitBounds(baseAlign.getBounds(), { padding: [50, 50] });

    // If we have calculated earthwork data, draw overlay segments
    if (state.calcRows && state.calcRows.length > 1) {
        const kmlStartCh = 0; // Assuming KML starts at CH 0 for now

        // Draw continuous coloured segments for earthworks
        for (let i = 1; i < state.calcRows.length; i++) {
            const r0 = state.calcRows[i - 1];
            const r1 = state.calcRows[i];

            const isFill = r1.bank > 0.001;
            const isCut = r1.cut > 0.001;
            if (!isFill && !isCut) continue;

            const p0 = getLatLngFromChainage(r0.chainage - kmlStartCh, points);
            const p1 = getLatLngFromChainage(r1.chainage - kmlStartCh, points);

            if (p0 && p1) {
                L.polyline([[p0.lat, p0.lng], [p1.lat, p1.lng]], {
                    color: isFill ? '#22c55e' : '#f43f5e', // Green for Fill, Red for Cut
                    weight: 6,
                    opacity: 0.75
                }).addTo(alignmentLayerGroup);
            }
        }
    }

    // Draw Bridges
    if (state.bridgeRows && state.bridgeRows.length) {
        state.bridgeRows.forEach(br => {
            const midCh = (safeNum(br.startChainage) + safeNum(br.endChainage)) / 2;
            const pt = getLatLngFromChainage(midCh, points);
            if (pt) {
                const marker = L.circleMarker([pt.lat, pt.lng], {
                    radius: 8,
                    fillColor: '#3b82f6', // blue
                    color: '#60a5fa',
                    weight: 2,
                    fillOpacity: 0.8
                }).addTo(alignmentLayerGroup);

                marker.bindPopup(`
          <div class="map-popup-title">Bridge ${br.bridgeNo}</div>
          <div class="map-popup-row"><span class="label">Type:</span> <span class="value">${br.bridgeType}</span></div>
          <div class="map-popup-row"><span class="label">Category:</span> <span class="value">${br.category}</span></div>
          <div class="map-popup-row"><span class="label">Span:</span> <span class="value">${br.spanConfig}</span></div>
          <div class="map-popup-row"><span class="label">Start CH:</span> <span class="value">${safeNum(br.startChainage).toFixed(3)}</span></div>
          <div class="map-popup-row"><span class="label">End CH:</span> <span class="value">${safeNum(br.endChainage).toFixed(3)}</span></div>
        `);
            }
        });
    }

    // Draw Curves
    if (state.curveRows && state.curveRows.length) {
        state.curveRows.forEach(cr => {
            const midCh = (safeNum(cr.startChainage) + safeNum(cr.endChainage)) / 2;
            const pt = getLatLngFromChainage(midCh, points);
            if (pt) {
                const marker = L.circleMarker([pt.lat, pt.lng], {
                    radius: 7,
                    fillColor: '#eab308', // yellow/gold
                    color: '#fde047',
                    weight: 2,
                    fillOpacity: 0.8
                }).addTo(alignmentLayerGroup);

                marker.bindPopup(`
          <div class="map-popup-title">Curve</div>
          <div class="map-popup-row"><span class="label">Radius:</span> <span class="value">${cr.radius} m</span></div>
          <div class="map-popup-row"><span class="label">Direction:</span> <span class="value">${cr.direction || 'N/A'}</span></div>
          <div class="map-popup-row"><span class="label">Start CH:</span> <span class="value">${safeNum(cr.startChainage).toFixed(3)}</span></div>
          <div class="map-popup-row"><span class="label">End CH:</span> <span class="value">${safeNum(cr.endChainage).toFixed(3)}</span></div>
        `);
            }
        });
    }

    // Draw Stations & Loops
    if (state.loopPlatformRows && state.loopPlatformRows.length) {
        state.loopPlatformRows.forEach(lp => {
            const midCh = (safeNum(lp.startChainage) + safeNum(lp.endChainage)) / 2;
            const pt = getLatLngFromChainage(midCh, points);
            if (pt) {
                const marker = L.circleMarker([pt.lat, pt.lng], {
                    radius: 10,
                    fillColor: '#06b6d4', // cyan
                    color: '#22d3ee',
                    weight: 2,
                    fillOpacity: 0.9
                }).addTo(alignmentLayerGroup);

                let imgHtml = '';
                const lowerDesc = String(lp.desc).toLowerCase().trim();
                if (state.stationPlans[lowerDesc]) {
                    const dataUrl = state.stationPlans[lowerDesc];
                    if (dataUrl.startsWith('data:application/pdf')) {
                        imgHtml = `<iframe src="${dataUrl}#toolbar=0&navpanes=0&scrollbar=0" class="station-plan-pdf" frameborder="0"></iframe>`;
                    } else {
                        imgHtml = `<img src="${dataUrl}" class="station-plan-img" alt="Station Plan"/>`;
                    }
                }

                marker.bindPopup(`
          <div class="map-popup-title">Station Yard / Loop</div>
          <div class="map-popup-row"><span class="label">Desc:</span> <span class="value">${lp.desc}</span></div>
          <div class="map-popup-row"><span class="label">Extra Width:</span> <span class="value">${lp.formationWidthAddition} m</span></div>
          <div class="map-popup-row"><span class="label">Center Dist:</span> <span class="value">${lp.trackCenterDistance} m</span></div>
          <div class="map-popup-row"><span class="label">Start CH:</span> <span class="value">${safeNum(lp.startChainage).toFixed(3)}</span></div>
          <div class="map-popup-row"><span class="label">End CH:</span> <span class="value">${safeNum(lp.endChainage).toFixed(3)}</span></div>
          ${imgHtml}
        `, { maxWidth: 350 });
            }
        });
    }
}


// ── Utilities ─────────────────────────────────────────────────────────────

function getDistanceInMeters(p1, p2) {
    const R = 6371e3; // Earth radius in meters
    const rad = Math.PI / 180;
    const lat1 = p1.lat * rad;
    const lat2 = p2.lat * rad;
    const dLat = (p2.lat - p1.lat) * rad;
    const dLng = (p2.lng - p1.lng) * rad;

    const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
        Math.cos(lat1) * Math.cos(lat2) *
        Math.sin(dLng / 2) * Math.sin(dLng / 2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

    return R * c;
}

function getLatLngFromChainage(targetCh, pointsArr) {
    if (!pointsArr || pointsArr.length < 2) return null;
    if (targetCh <= pointsArr[0].ch) return pointsArr[0];
    if (targetCh >= pointsArr[pointsArr.length - 1].ch) return pointsArr[pointsArr.length - 1];

    // Binary search for closest points
    let low = 0, high = pointsArr.length - 1;
    while (low <= high) {
        const mid = Math.floor((low + high) / 2);
        if (pointsArr[mid].ch < targetCh) low = mid + 1;
        else if (pointsArr[mid].ch > targetCh) high = mid - 1;
        else return pointsArr[mid];
    }

    const p1 = pointsArr[high];
    const p2 = pointsArr[low];

    if (!p1 || !p2) return null;

    // Linear interpolation
    const ratio = (targetCh - p1.ch) / (p2.ch - p1.ch);
    const lat = p1.lat + ratio * (p2.lat - p1.lat);
    const lng = p1.lng + ratio * (p2.lng - p1.lng);

    return { lat, lng };
}
