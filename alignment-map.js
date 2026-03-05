/**
 * alignment-map.js
 * Handles the Leaflet.js (Free Map) integration for geographic alignment and earthwork visual overlays.
 */

let alignmentMap = null;
let baseLayers = {};
let mapItems = null; // We'll use a FeatureGroup for easy clearing

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
    saveState();

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

    const coordRegex = /<coordinates>([\s\S]*?)<\/coordinates>/g;
    const points_extracted = [];
    let match;

    while ((match = coordRegex.exec(kmlText)) !== null) {
        const text = match[1].trim();
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
        alert("Could not parse coordinates from the KML file.");
        return;
    }

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
        drawAlignmentMap();
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
    if (!alignmentMap || !mapItems) return;

    if (!state.kmlData || !state.kmlData.points || !state.kmlData.points.length) {
        document.getElementById("alignmentMapContainer").style.display = "none";
        document.getElementById("alignmentMapEmpty").style.display = "flex";
        return;
    }

    document.getElementById("alignmentMapContainer").style.display = "block";
    document.getElementById("alignmentMapEmpty").style.display = "none";
    alignmentMap.invalidateSize();

    mapItems.clearLayers();

    const points = state.kmlData.points;
    const path = points.map(p => [p.lat, p.lng]);

    // Base Alignment Polyline
    L.polyline(path, {
        color: '#ffffff',
        weight: 3,
        opacity: 0.8
    }).addTo(mapItems);

    const bounds = L.latLngBounds(path);
    if (bounds.isValid()) {
        alignmentMap.fitBounds(bounds, { padding: [50, 50] });
    }

    const startChOffset = (state.calcRows && state.calcRows.length) ? _safeNum(state.calcRows[0].chainage) : 0;

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
                    weight: 6,
                    opacity: 0.9
                }).addTo(mapItems);
            }
        }
    }

    // Helper to add titled markers
    const addMarker = (midCh, title, content, color) => {
        const pt = getLatLngFromChainage(midCh - startChOffset, points);
        if (!pt) return;

        L.circleMarker([pt.lat, pt.lng], {
            radius: 8,
            fillColor: color,
            color: '#ffffff',
            weight: 1,
            opacity: 1,
            fillOpacity: 0.9
        }).bindPopup(content, { maxWidth: 350 }).addTo(mapItems);
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
        const planKey = String(lp.desc).toLowerCase().trim();
        const plan = state.stationPlans[planKey];

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
