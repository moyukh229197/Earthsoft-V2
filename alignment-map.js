/**
 * alignment-map.js
 * Handles the Google Maps integration for geographic alignment and earthwork visual overlays.
 */

let alignmentMap = null;
let googleInfoWindow = null;
let mapItems = []; // Store markers and polylines for easy clearing

// Initialize UI and Event Listeners when DOM is ready
document.addEventListener("DOMContentLoaded", () => {
    // Map control buttons
    const importKmlBtn = document.getElementById("importKmlBtn");
    const kmlImportInput = document.getElementById("kmlImportInput");
    const importStationPlanBtn = document.getElementById("importStationPlanBtn");
    const stationPlanImportInput = document.getElementById("stationPlanImportInput");
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
    if (clearMapBtn) {
        clearMapBtn.addEventListener("click", clearMapData);
    }
    if (mapTypeSelect) {
        mapTypeSelect.addEventListener("change", (e) => {
            if (alignmentMap) {
                alignmentMap.setMapTypeId(e.target.value);
            }
        });
    }

    // Handle map resize on page reveal
    const workNavBtns = document.querySelectorAll(".work-nav-btn");
    workNavBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            if (btn.dataset.workPageBtn === "alignment-map") {
                setTimeout(() => {
                    if (alignmentMap) {
                        google.maps.event.trigger(alignmentMap, "resize");
                        if (state.kmlData) drawAlignmentMap();
                    }
                }, 300);
            }
        });
    });
});

function clearMapData() {
    if (!confirm("Are you sure you want to clear all map data? This will remove the alignment and all station plans.")) return;

    state.kmlData = null;
    state.stationPlans = {};
    saveState();

    // Clear visualization
    if (mapItems && mapItems.length) {
        mapItems.forEach(item => { if (item && item.setMap) item.setMap(null); });
    }
    mapItems = [];

    // Reset view
    document.getElementById("alignmentMapContainer").style.display = "none";
    document.getElementById("alignmentMapEmpty").style.display = "flex";

    console.log("Map data cleared.");
}

// Initialize Map System (called by Google Maps API loader callback)
window.initMapSystem = function () {
    console.log("initMapSystem / initMap called!");
    if (typeof google === 'undefined') {
        console.warn("Google objects not found in initMapSystem.");
        return;
    }

    // Initialize Map
    const mapContainer = document.getElementById("alignmentMapContainer");
    const emptyState = document.getElementById("alignmentMapEmpty");

    if (!mapContainer || !emptyState) {
        console.error("Map container or empty state element not found.");
        return;
    }

    if (alignmentMap) {
        console.log("Map already initialized.");
        return;
    }

    alignmentMap = new google.maps.Map(mapContainer, {
        center: { lat: 20.5937, lng: 78.9629 }, // Default to India center
        zoom: 5,
        mapTypeId: google.maps.MapTypeId.HYBRID,
        tilt: 0,
        mapTypeControl: true,
        streetViewControl: false,
        fullscreenControl: true,
        backgroundColor: '#0d1117',
        styles: [
            { elementType: "geometry", stylers: [{ color: "#242f3e" }] },
            { elementType: "labels.text.stroke", stylers: [{ color: "#242f3e" }] },
            { elementType: "labels.text.fill", stylers: [{ color: "#746855" }] },
            {
                featureType: "administrative.locality",
                elementType: "labels.text.fill",
                stylers: [{ color: "#d59563" }],
            }
        ]
    });

    googleInfoWindow = new google.maps.InfoWindow();

    // Re-draw map if data exists (on load)
    if (state.kmlData) {
        drawAlignmentMap();
    }
};

// Alias for keyless script compatibility (it hardcodes initMap)
window.initMap = window.initMapSystem;

// Second Fallback: If the callback doesn't fire, try to init after some time
setTimeout(() => {
    if (!alignmentMap && typeof google !== 'undefined' && google.maps) {
        console.log("Fallback: Initializing map manually as callback didn't fire.");
        window.initMapSystem();
    }
}, 3000);

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
    console.log("Parsing KML data (Regex Mode)...");

    // Use regex to find coordinate blocks - much more robust against namespaces
    const coordRegex = /<coordinates>([\s\S]*?)<\/coordinates>/g;
    const points_extracted = [];
    let match;

    while ((match = coordRegex.exec(kmlText)) !== null) {
        const text = match[1].trim();
        // Split by whitespace/newlines
        const tuples = text.split(/[\s\n\r]+/);
        tuples.forEach(tuple => {
            const parts = tuple.split(',').map(s => s.trim());
            if (parts.length >= 2) {
                const lng = parseFloat(parts[0]);
                const lat = parseFloat(parts[1]);
                if (!isNaN(lat) && !isNaN(lng)) {
                    points_extracted.push({ lat, lng });
                }
            }
        });
    }

    if (points_extracted.length === 0) {
        console.error("No coordinates found using regex. Falling back to DOM parser...");
        // Fallback to DOM parser if regex failed (unlikely but safe)
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(kmlText, "text/xml");
        const coordsNodes = xmlDoc.getElementsByTagName("coordinates");
        for (let i = 0; i < coordsNodes.length; i++) {
            const text = coordsNodes[i].textContent || "";
            const tuples = text.trim().split(/[\s\n\r]+/);
            tuples.forEach(tuple => {
                const parts = tuple.split(',').map(s => s.trim());
                if (parts.length >= 2) {
                    const lng = parseFloat(parts[0]);
                    const lat = parseFloat(parts[1]);
                    if (!isNaN(lat) && !isNaN(lng)) points_extracted.push({ lat, lng });
                }
            });
        }
    }

    if (points_extracted.length === 0) {
        console.error("Critical: No valid coordinates found in KML.");
        alert("Could not parse coordinates from the KML file. Please ensure it contains standard Google Earth / KML path data.");
        return;
    }

    console.log(`Successfully extracted ${points_extracted.length} points.`);

    // Remove duplicates
    const points = [];
    points_extracted.forEach((p, idx) => {
        if (idx === 0) points.push(p);
        else {
            const last = points[points.length - 1];
            if (Math.abs(last.lat - p.lat) > 1e-9 || Math.abs(last.lng - p.lng) > 1e-9) {
                points.push(p);
            }
        }
    });

    let cumDist = 0;
    points[0].ch = 0;
    for (let i = 1; i < points.length; i++) {
        cumDist += getDistanceInMeters(points[i - 1], points[i]);
        points[i].ch = cumDist;
    }

    state.kmlData = { points, totalDistance: cumDist };
    saveState();

    alert(`Successfully imported ${points.length} coordinates and mapped ${cumDist.toFixed(0)}m of alignment.`);

    const container = document.querySelector('.work-page[data-work-page="alignment-map"]');
    if (container && container.style.display !== 'none') {
        console.log("On map page, triggering draw...");
        drawAlignmentMap();
    } else {
        console.log("Not on map page, state saved. Please go to Alignment Map to see results.");
    }
}

// ── Station Plan Import ────────────────────────────────────────────────────

function handleStationPlanImport(e) {
    const file = e.target.files[0];
    if (!file) return;

    if (!state.loopPlatformRows || state.loopPlatformRows.length === 0) {
        alert("Please add Loops & Station data first.");
        return;
    }

    const stationName = prompt("Enter Station Name to attach this plan to:\n" + state.loopPlatformRows.map(r => r.desc).join('\n'));
    if (!stationName) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
        state.stationPlans[stationName.toLowerCase().trim()] = evt.target.result;
        saveState();
        if (state.kmlData) drawAlignmentMap();
        e.target.value = "";
    };
    reader.readAsDataURL(file);
}

// ── Map Rendering Core ───────────────────────────────────────────────────

function drawAlignmentMap() {
    console.log("Drawing Alignment Map...");
    if (!alignmentMap) {
        console.warn("Map not initialized yet. Skipping draw.");
        return;
    }

    if (!state.kmlData || !state.kmlData.points || !state.kmlData.points.length) {
        console.log("No KML data to draw.");
        document.getElementById("alignmentMapContainer").style.display = "none";
        document.getElementById("alignmentMapEmpty").style.display = "flex";
        return;
    }

    // Hide empty state, show map
    document.getElementById("alignmentMapContainer").style.display = "block";
    document.getElementById("alignmentMapEmpty").style.display = "none";

    // Trigger resize is important if the container was hidden
    google.maps.event.trigger(alignmentMap, "resize");

    const points = state.kmlData.points;
    const path = points.map(p => ({ lat: p.lat, lng: p.lng }));

    // Define bounds for the alignment
    const bounds = new google.maps.LatLngBounds();
    path.forEach(p => bounds.extend(p));

    // Clear previous items
    if (mapItems && mapItems.length) {
        mapItems.forEach(item => { if (item && item.setMap) item.setMap(null); });
    }
    mapItems = [];

    // Base Alignment Polyline
    const basePoly = new google.maps.Polyline({
        path: path,
        geodesic: true,
        strokeColor: "#ffffff",
        strokeOpacity: 0.8, // Increased for visibility
        strokeWeight: 3,    // Thicker line
        map: alignmentMap
    });
    mapItems.push(basePoly);

    if (!bounds.isEmpty()) {
        alignmentMap.fitBounds(bounds);
        console.log("Map bounds fitted to alignment.");
    }

    // Offset Mapping: Most projects don't start at CH 0. 
    // We assume the first point of the KML matches the first point of our calculations.
    const startChOffset = (state.calcRows && state.calcRows.length) ? _safeNum(state.calcRows[0].chainage) : 0;
    console.log(`Using Chainage Offset: ${startChOffset}`);

    // Earthwork Overlays
    if (state.calcRows && state.calcRows.length > 1) {
        for (let i = 1; i < state.calcRows.length; i++) {
            const r0 = state.calcRows[i - 1], r1 = state.calcRows[i];
            const isFill = r1.bank > 0.001, isCut = r1.cut > 0.001;
            if (!isFill && !isCut) continue;

            // Map project chainage to KML-relative distance
            const p0 = getLatLngFromChainage(r0.chainage - startChOffset, points);
            const p1 = getLatLngFromChainage(r1.chainage - startChOffset, points);

            if (p0 && p1) {
                const ewPoly = new google.maps.Polyline({
                    path: [p0, p1],
                    strokeColor: isFill ? "#22c55e" : "#f43f5e",
                    strokeOpacity: 0.9,
                    strokeWeight: 6,
                    map: alignmentMap
                });
                mapItems.push(ewPoly);
            }
        }
    }

    // Helper to add titled markers
    const addMarker = (midCh, title, content, color) => {
        // Map project chainage to KML-relative distance
        const pt = getLatLngFromChainage(midCh - startChOffset, points);
        if (!pt) return;

        const marker = new google.maps.Marker({
            position: pt,
            map: alignmentMap,
            title: title,
            icon: {
                path: google.maps.SymbolPath.CIRCLE,
                fillColor: color,
                fillOpacity: 0.9,
                strokeColor: '#ffffff',
                strokeWeight: 1,
                scale: 8
            }
        });

        marker.addListener("click", () => {
            googleInfoWindow.setContent(content);
            googleInfoWindow.open(alignmentMap, marker);
        });
        mapItems.push(marker);
    };

    // Bridges
    state.bridgeRows.forEach(br => {
        const mid = (_safeNum(br.startChainage) + _safeNum(br.endChainage)) / 2;
        const content = `
      <div class="map-popup-title">Bridge ${br.bridgeNo}</div>
      <div class="map-popup-row"><span class="label">Type:</span> <span class="value">${br.bridgeType}</span></div>
      <div class="map-popup-row"><span class="label">Start:</span> <span class="value">${_safeNum(br.startChainage).toFixed(3)}</span></div>
      <div class="map-popup-row"><span class="label">End:</span> <span class="value">${_safeNum(br.endChainage).toFixed(3)}</span></div>
    `;
        addMarker(mid, `Bridge ${br.bridgeNo}`, content, "#3b82f6");
    });

    // Curves
    state.curveRows.forEach(cr => {
        const mid = (_safeNum(cr.startChainage) + _safeNum(cr.endChainage)) / 2;
        const content = `
      <div class="map-popup-title">Curve</div>
      <div class="map-popup-row"><span class="label">Radius:</span> <span class="value">${cr.radius} m</span></div>
    `;
        addMarker(mid, "Curve", content, "#eab308");
    });

    // Stations
    state.loopPlatformRows.forEach(lp => {
        const mid = (_safeNum(lp.startChainage) + _safeNum(lp.endChainage)) / 2;
        let planHtml = '';
        const plan = state.stationPlans[String(lp.desc).toLowerCase().trim()];
        if (plan) {
            if (plan.startsWith('data:application/pdf')) {
                planHtml = `<iframe src="${plan}#toolbar=0" class="station-plan-pdf" style="width:100%;height:200px;border:none;"></iframe>`;
            } else {
                planHtml = `<img src="${plan}" class="station-plan-img" style="width:100%;max-height:150px;object-fit:contain;"/>`;
            }
        }

        const content = `
      <div class="map-popup-title">${lp.desc}</div>
      <div class="map-popup-row"><span class="label">Start CH:</span> <span class="value">${_safeNum(lp.startChainage).toFixed(3)}</span></div>
      <div class="map-popup-row"><span class="label">End CH:</span> <span class="value">${_safeNum(lp.endChainage).toFixed(3)}</span></div>
      ${planHtml}
    `;
        addMarker(mid, lp.desc, content, "#06b6d4");
    });
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
