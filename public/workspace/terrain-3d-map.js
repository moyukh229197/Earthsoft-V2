/**
 * terrain-3d-map.js
 * Cesium-based 3D terrain tab for rendering the alignment over OSM + terrain.
 */

let terrain3dViewer = null;
let terrain3dCesiumLoadPromise = null;
let terrain3dInitialized = false;
let terrain3dBuildingsTileset = null;
let terrain3dBuildingEntities = [];
let terrain3dBuildingsEnabled = false;
let terrain3dLastDataKey = "";
let terrain3dMapType = "osm";
let terrain3dTerrainSource = "arcgis";
let terrain3dFlightPath = [];
let terrain3dCurrentImageryLayer = null;
const TERRAIN3D_MAP_TYPE_STORAGE_KEY = "earthsoft_terrain3d_map_type";
const TERRAIN3D_TERRAIN_SOURCE_STORAGE_KEY = "earthsoft_terrain3d_terrain_source";
const TERRAIN3D_BUILDINGS_MAX_FEATURES = 1000;
const terrain3dFlightState = {
  running: false,
  rafId: null,
  lastTs: 0,
  speedMps: 70,
  progressMeters: 0,
  cumulativeDistances: [],
  totalDistance: 0,
};

const TERRAIN3D_CESIUM_VERSION = "1.124.0";
const TERRAIN3D_CESIUM_BASE_URL = `https://cdn.jsdelivr.net/npm/cesium@${TERRAIN3D_CESIUM_VERSION}/Build/Cesium/`;
const TERRAIN3D_CESIUM_JS_URL = `${TERRAIN3D_CESIUM_BASE_URL}Cesium.js`;
const TERRAIN3D_CESIUM_CSS_URL = `${TERRAIN3D_CESIUM_BASE_URL}Widgets/widgets.css`;

function terrain3dSafeNum(v, fallback = 0) {
  const n = Number(v);
  return Number.isFinite(n) ? n : fallback;
}

function terrain3dNormalizeElevationMeters(rawValue) {
  const n = terrain3dSafeNum(rawValue, Number.NaN);
  if (!Number.isFinite(n)) return Number.NaN;
  const abs = Math.abs(n);
  if (abs > 10000) return n / 100;
  if (abs > 1200) return n / 10;
  return n;
}

function terrain3dClampReasonableHeight(heightMeters) {
  const h = terrain3dSafeNum(heightMeters, Number.NaN);
  if (!Number.isFinite(h)) return Number.NaN;
  return Math.max(-500, Math.min(9000, h));
}

function terrain3dResolveDesignHeight(proposedRaw, groundRaw) {
  const proposed = terrain3dNormalizeElevationMeters(proposedRaw);
  const ground = terrain3dNormalizeElevationMeters(groundRaw);

  if (Number.isFinite(proposed) && Number.isFinite(ground)) {
    const delta = proposed - ground;
    if (Math.abs(delta) > 300) {
      const scaled10 = ground + (delta / 10);
      if (Math.abs(scaled10 - ground) <= 300) {
        return terrain3dClampReasonableHeight(scaled10);
      }
      const scaled100 = ground + (delta / 100);
      if (Math.abs(scaled100 - ground) <= 300) {
        return terrain3dClampReasonableHeight(scaled100);
      }
      return terrain3dClampReasonableHeight(ground);
    }
    return terrain3dClampReasonableHeight(proposed);
  }

  if (Number.isFinite(proposed)) return terrain3dClampReasonableHeight(proposed);
  if (Number.isFinite(ground)) return terrain3dClampReasonableHeight(ground);
  return 0;
}

function terrain3dParseChainage(value) {
  if (typeof parseChainage === "function") {
    const parsed = parseChainage(value);
    if (Number.isFinite(parsed)) return parsed;
  }
  if (Number.isFinite(Number(value))) return Number(value);
  const text = String(value || "").trim();
  const kmMatch = text.match(/^(\d+)\+(\d+(?:\.\d+)?)$/);
  if (kmMatch) {
    return (Number(kmMatch[1]) * 1000) + Number(kmMatch[2]);
  }
  return Number.NaN;
}

function terrain3dSetStatus(message, tone = "default") {
  const statusEl = document.getElementById("terrain3dStatus");
  if (!statusEl) return;
  statusEl.textContent = message;
  statusEl.classList.remove("is-ready", "is-warning");
  if (tone === "ready") statusEl.classList.add("is-ready");
  if (tone === "warning") statusEl.classList.add("is-warning");
}

function terrain3dSetVisibility(hasMap) {
  const emptyEl = document.getElementById("terrain3dEmpty");
  const mapEl = document.getElementById("terrain3dMapContainer");
  if (emptyEl) emptyEl.style.display = hasMap ? "none" : "flex";
  if (mapEl) mapEl.style.display = hasMap ? "block" : "none";
}

function terrain3dInterpolatePointByChainage(points, targetCh) {
  if (!Array.isArray(points) || points.length < 2 || !Number.isFinite(targetCh)) return null;
  if (targetCh <= terrain3dSafeNum(points[0]?.ch, 0)) return points[0];
  const lastCh = terrain3dSafeNum(points[points.length - 1]?.ch, 0);
  if (targetCh >= lastCh) return points[points.length - 1];

  let low = 0;
  let high = points.length - 1;
  while (low <= high) {
    const mid = Math.floor((low + high) / 2);
    const midCh = terrain3dSafeNum(points[mid]?.ch, Number.NaN);
    if (midCh < targetCh) low = mid + 1;
    else if (midCh > targetCh) high = mid - 1;
    else return points[mid];
  }
  const p1 = points[high];
  const p2 = points[low];
  if (!p1 || !p2) return null;
  const p1Ch = terrain3dSafeNum(p1.ch, 0);
  const p2Ch = terrain3dSafeNum(p2.ch, 0);
  const span = Math.max(p2Ch - p1Ch, 1e-6);
  const t = (targetCh - p1Ch) / span;
  return {
    lat: terrain3dSafeNum(p1.lat, 0) + (terrain3dSafeNum(p2.lat, 0) - terrain3dSafeNum(p1.lat, 0)) * t,
    lng: terrain3dSafeNum(p1.lng, 0) + (terrain3dSafeNum(p2.lng, 0) - terrain3dSafeNum(p1.lng, 0)) * t,
  };
}

function terrain3dDecimatePoints(points, maxPoints = 1800) {
  if (!Array.isArray(points) || !points.length) return [];
  if (points.length <= maxPoints) return [...points];
  const reduced = [];
  const step = Math.max(1, Math.ceil(points.length / maxPoints));
  for (let i = 0; i < points.length; i += step) {
    reduced.push(points[i]);
  }
  const last = points[points.length - 1];
  if (reduced[reduced.length - 1] !== last) reduced.push(last);
  return reduced;
}

function terrain3dBuildGroundPath(kmlPoints) {
  return terrain3dDecimatePoints(kmlPoints, 2200)
    .map((pt) => ({
      lat: terrain3dSafeNum(pt?.lat, Number.NaN),
      lng: terrain3dSafeNum(pt?.lng, Number.NaN),
    }))
    .filter((pt) => Number.isFinite(pt.lat) && Number.isFinite(pt.lng));
}

function terrain3dBuildDesignPath(kmlPoints, calcRows) {
  if (!Array.isArray(calcRows) || calcRows.length < 2) return [];
  const startAbs = terrain3dParseChainage(calcRows[0]?.chainage);
  if (!Number.isFinite(startAbs)) return [];

  const samples = [];
  const step = Math.max(1, Math.ceil(calcRows.length / 1600));
  for (let i = 0; i < calcRows.length; i += step) {
    const row = calcRows[i];
    const chAbs = terrain3dParseChainage(row?.chainage);
    if (!Number.isFinite(chAbs)) continue;
    const point = terrain3dInterpolatePointByChainage(kmlPoints, chAbs - startAbs);
    if (!point) continue;
    const height = terrain3dResolveDesignHeight(row?.proposedLevel, row?.groundLevel);
    samples.push({
      lat: terrain3dSafeNum(point.lat, Number.NaN),
      lng: terrain3dSafeNum(point.lng, Number.NaN),
      height: terrain3dSafeNum(height, 0),
    });
  }

  const lastRow = calcRows[calcRows.length - 1];
  const lastAbs = terrain3dParseChainage(lastRow?.chainage);
  const lastPoint = Number.isFinite(lastAbs) ? terrain3dInterpolatePointByChainage(kmlPoints, lastAbs - startAbs) : null;
  if (lastPoint) {
    const height = terrain3dResolveDesignHeight(lastRow?.proposedLevel, lastRow?.groundLevel);
    samples.push({
      lat: terrain3dSafeNum(lastPoint.lat, Number.NaN),
      lng: terrain3dSafeNum(lastPoint.lng, Number.NaN),
      height: terrain3dSafeNum(height, 0),
    });
  }

  const filtered = samples.filter((pt) => Number.isFinite(pt.lat) && Number.isFinite(pt.lng) && Number.isFinite(pt.height));
  const deduped = [];
  filtered.forEach((pt) => {
    const prev = deduped[deduped.length - 1];
    if (!prev || Math.abs(prev.lat - pt.lat) > 1e-9 || Math.abs(prev.lng - pt.lng) > 1e-9 || Math.abs(prev.height - pt.height) > 0.001) {
      deduped.push(pt);
    }
  });
  return deduped;
}

function terrain3dBuildDataKey(kmlPoints, calcRows) {
  return [
    kmlPoints?.length || 0,
    calcRows?.length || 0,
    terrain3dParseChainage(calcRows?.[0]?.chainage),
    terrain3dParseChainage(calcRows?.[calcRows?.length - 1]?.chainage),
  ].join("|");
}

function terrain3dGetMapTilerKey() {
  return String(window.MAPTILER_API_KEY || "").trim();
}

function terrain3dLoadStoredMapType() {
  try {
    const saved = String(localStorage.getItem(TERRAIN3D_MAP_TYPE_STORAGE_KEY) || "").trim().toLowerCase();
    if (saved === "osm" || saved === "dark" || saved === "topo" || saved === "satellite") {
      terrain3dMapType = saved;
    }
  } catch (_) {}
}

function terrain3dSaveMapType(mapType) {
  try {
    localStorage.setItem(TERRAIN3D_MAP_TYPE_STORAGE_KEY, String(mapType || "osm").toLowerCase());
  } catch (_) {}
}

function terrain3dLoadStoredTerrainSource() {
  try {
    const saved = String(localStorage.getItem(TERRAIN3D_TERRAIN_SOURCE_STORAGE_KEY) || "").trim().toLowerCase();
    if (saved === "arcgis" || saved === "maptiler") {
      terrain3dTerrainSource = saved;
    }
  } catch (_) {}
}

function terrain3dSaveTerrainSource(source) {
  try {
    localStorage.setItem(TERRAIN3D_TERRAIN_SOURCE_STORAGE_KEY, String(source || "arcgis").toLowerCase());
  } catch (_) {}
}

function terrain3dCreatePublicImageryProvider(Cesium, mapType) {
  const type = String(mapType || "osm").toLowerCase();
  if (type === "satellite") {
    return new Cesium.UrlTemplateImageryProvider({
      url: "https://services.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
      maximumLevel: 19,
      credit: "Esri, Maxar, Earthstar Geographics",
    });
  }
  if (type === "dark") {
    return new Cesium.UrlTemplateImageryProvider({
      url: "https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}.png",
      subdomains: ["a", "b", "c", "d"],
      maximumLevel: 20,
      credit: "OpenStreetMap, CARTO",
    });
  }
  if (type === "topo") {
    return new Cesium.UrlTemplateImageryProvider({
      url: "https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png",
      subdomains: ["a", "b", "c"],
      maximumLevel: 17,
      credit: "OpenTopoMap, OpenStreetMap contributors",
    });
  }
  return new Cesium.OpenStreetMapImageryProvider({
    url: "https://tile.openstreetmap.org/",
    maximumLevel: 19,
  });
}

function terrain3dCreateImageryProvider(Cesium, mapType, options = {}) {
  const type = String(mapType || "osm").toLowerCase();
  const useMapTiler = options.useMapTiler !== false;
  const mapTilerKey = terrain3dGetMapTilerKey();
  if (useMapTiler && mapTilerKey && type !== "osm") {
    const mapId = type === "satellite"
      ? "satellite"
      : type === "dark"
        ? "backdrop"
        : type === "topo"
          ? "topo-v2"
          : "streets-v2";
    const ext = type === "satellite" ? "jpg" : "png";
    return new Cesium.UrlTemplateImageryProvider({
      url: `https://api.maptiler.com/maps/${mapId}/{z}/{x}/{y}.${ext}?key=${encodeURIComponent(mapTilerKey)}`,
      maximumLevel: 20,
      credit: "MapTiler, OpenStreetMap contributors",
    });
  }
  return terrain3dCreatePublicImageryProvider(Cesium, type);
}

function terrain3dMapTypeLabel(mapType) {
  const type = String(mapType || "osm").toLowerCase();
  if (type === "dark") return "Dark";
  if (type === "topo") return "Topo";
  if (type === "satellite") return "Satellite";
  return "OSM Street";
}

function terrain3dApplyMapType(mapType, announce = true) {
  const Cesium = window.Cesium;
  if (!terrain3dViewer || !Cesium) return;
  const type = String(mapType || "osm").toLowerCase();
  const hasMapTilerKey = Boolean(terrain3dGetMapTilerKey());
  const prefersMapTiler = hasMapTilerKey && type !== "osm";
  try {
    const imagery = terrain3dCreateImageryProvider(Cesium, type, { useMapTiler: prefersMapTiler });
    terrain3dViewer.imageryLayers.removeAll();
    terrain3dCurrentImageryLayer = terrain3dViewer.imageryLayers.addImageryProvider(imagery);
    const fallbackToPublic = () => {
      const fallbackToOsm = () => {
        try {
          const osmImagery = terrain3dCreatePublicImageryProvider(Cesium, "osm");
          terrain3dViewer.imageryLayers.removeAll();
          terrain3dCurrentImageryLayer = terrain3dViewer.imageryLayers.addImageryProvider(osmImagery);
          terrain3dMapType = "osm";
          terrain3dSaveMapType("osm");
          terrain3dSetStatus("Selected map failed to load. Reverted to OSM Street.", "warning");
        } catch (innerError) {
          console.warn("OSM fallback imagery failed:", innerError);
        }
      };

      try {
        const fallbackImagery = prefersMapTiler
          ? terrain3dCreatePublicImageryProvider(Cesium, type)
          : terrain3dCreatePublicImageryProvider(Cesium, "osm");
        terrain3dViewer.imageryLayers.removeAll();
        terrain3dCurrentImageryLayer = terrain3dViewer.imageryLayers.addImageryProvider(fallbackImagery);
        if (!prefersMapTiler) {
          fallbackToOsm();
          return;
        }
        terrain3dSetStatus(`MapTiler tiles unavailable on this origin. Switched ${terrain3dMapTypeLabel(type)} to public fallback.`, "warning");
      } catch (error) {
        console.warn("Public imagery fallback failed:", error);
        fallbackToOsm();
      }
    };
    if (terrain3dCurrentImageryLayer?.imageryProvider?.errorEvent?.addEventListener) {
      terrain3dCurrentImageryLayer.imageryProvider.errorEvent.addEventListener(() => {
        fallbackToPublic();
      });
    }
    setTimeout(() => {
      if (terrain3dViewer && terrain3dViewer.imageryLayers.length === 0) {
        fallbackToPublic();
      }
    }, 2500);
    terrain3dMapType = type;
    terrain3dSaveMapType(type);
    if (announce) {
      terrain3dSetStatus(`Map switched to ${terrain3dMapTypeLabel(type)}.`, "ready");
    }
  } catch (error) {
    console.warn("Failed to switch terrain 3D map type:", error);
    terrain3dSetStatus("Could not switch map type. Using existing base layer.", "warning");
  }
}

async function terrain3dApplyTerrainProvider() {
  const Cesium = window.Cesium;
  if (!terrain3dViewer || !Cesium) return;
  const mapTilerKey = terrain3dGetMapTilerKey();
  const preferMapTilerTerrain = window.MAPTILER_FORCE_TERRAIN === true || terrain3dTerrainSource === "maptiler";

  const applyNonMapTilerTerrainFallback = async (reasonText = "") => {
    const reason = String(reasonText || "").trim();
    const tone = reason.toLowerCase().includes("selected arcgis") ? "ready" : "warning";
    const prefix = tone === "ready" ? "Terrain source updated" : "MapTiler terrain blocked";

    try {
      if (typeof Cesium.createWorldTerrainAsync === "function") {
        terrain3dViewer.terrainProvider = await Cesium.createWorldTerrainAsync();
        terrain3dSetStatus(
          reason
            ? `${prefix} (${reason}). Switched to Cesium World Terrain.`
            : "Switched to Cesium World Terrain.",
          tone
        );
        return true;
      }
    } catch (error) {
      console.warn("Cesium World Terrain (async) fallback failed:", error);
    }

    try {
      if (typeof Cesium.createWorldTerrain === "function") {
        terrain3dViewer.terrainProvider = Cesium.createWorldTerrain();
        terrain3dSetStatus(
          reason
            ? `${prefix} (${reason}). Switched to Cesium World Terrain.`
            : "Switched to Cesium World Terrain.",
          tone
        );
        return true;
      }
    } catch (error) {
      console.warn("Cesium World Terrain fallback failed:", error);
    }

    try {
      if (Cesium.ArcGISTiledElevationTerrainProvider?.fromUrl) {
        terrain3dViewer.terrainProvider = await Cesium.ArcGISTiledElevationTerrainProvider.fromUrl(
          "https://elevation3d.arcgis.com/arcgis/rest/services/WorldElevation3D/Terrain3D/ImageServer"
        );
        terrain3dSetStatus(
          reason
            ? `${prefix} (${reason}). Switched to ArcGIS elevation terrain.`
            : "Switched to ArcGIS elevation terrain.",
          tone
        );
        return true;
      }
    } catch (error) {
      console.warn("ArcGIS elevation fallback failed:", error);
    }

    terrain3dViewer.terrainProvider = new Cesium.EllipsoidTerrainProvider();
    terrain3dSetStatus(
      reason
        ? `${prefix} (${reason}). No alternate terrain source available, using flat fallback.`
        : "No terrain provider available. Showing flat fallback.",
      "warning"
    );
    return false;
  };

  try {
    if (preferMapTilerTerrain && mapTilerKey) {
      const terrainUrl = `https://api.maptiler.com/tiles/terrain-quantized-mesh-v2/tiles.json?key=${encodeURIComponent(mapTilerKey)}`;
      const probe = await fetch(terrainUrl, { method: "GET", mode: "cors", credentials: "omit" });
      if (!probe.ok) {
        throw new Error(`MapTiler terrain probe failed with HTTP ${probe.status}`);
      }
      let mapTilerProvider;
      if (Cesium.CesiumTerrainProvider?.fromUrl) {
        mapTilerProvider = await Cesium.CesiumTerrainProvider.fromUrl(terrainUrl);
      } else {
        mapTilerProvider = new Cesium.CesiumTerrainProvider({ url: terrainUrl });
      }
      if (mapTilerProvider?.errorEvent?.addEventListener) {
        let failedOver = false;
        mapTilerProvider.errorEvent.addEventListener(() => {
          if (failedOver || !terrain3dViewer) return;
          failedOver = true;
          void applyNonMapTilerTerrainFallback("check MapTiler Allowed Origins");
        });
      }
      terrain3dViewer.terrainProvider = mapTilerProvider;
      terrain3dSetStatus("3D terrain is active via MapTiler terrain mesh.", "ready");
      return;
    }

    if (preferMapTilerTerrain && !mapTilerKey) {
      await applyNonMapTilerTerrainFallback("MapTiler key missing");
      return;
    }
    await applyNonMapTilerTerrainFallback("selected ArcGIS");
  } catch (error) {
    await applyNonMapTilerTerrainFallback("check MapTiler key and origin");
    console.warn("Terrain load failed:", error);
  }
}

async function terrain3dApplyTerrainSource(source, announce = true) {
  const normalized = String(source || "arcgis").toLowerCase() === "maptiler" ? "maptiler" : "arcgis";
  terrain3dTerrainSource = normalized;
  terrain3dSaveTerrainSource(normalized);
  const selectEl = document.getElementById("terrain3dTerrainSourceSelect");
  if (selectEl && selectEl.value !== normalized) {
    selectEl.value = normalized;
  }

  if (!terrain3dViewer || !window.Cesium) return;
  await terrain3dApplyTerrainProvider();
  if (announce && normalized === "maptiler" && !terrain3dGetMapTilerKey()) {
    terrain3dSetStatus("MapTiler terrain selected but key is missing. Using ArcGIS elevation terrain.", "warning");
  }
}

function terrain3dResetFlightState() {
  terrain3dFlightPath = [];
  terrain3dFlightState.running = false;
  terrain3dFlightState.lastTs = 0;
  terrain3dFlightState.progressMeters = 0;
  terrain3dFlightState.cumulativeDistances = [];
  terrain3dFlightState.totalDistance = 0;
  if (terrain3dFlightState.rafId) {
    cancelAnimationFrame(terrain3dFlightState.rafId);
    terrain3dFlightState.rafId = null;
  }
  terrain3dSyncFlightControls();
}

function terrain3dSyncFlightControls() {
  const startBtn = document.getElementById("terrain3dStartFlyBtn");
  const stopBtn = document.getElementById("terrain3dStopFlyBtn");
  if (startBtn) {
    startBtn.disabled = terrain3dFlightPath.length < 2;
    startBtn.innerHTML = terrain3dFlightState.running
      ? '<i class="ri-pause-circle-line" style="margin-right:4px;"></i> Pause Fly-Through'
      : '<i class="ri-play-circle-line" style="margin-right:4px;"></i> Start Fly-Through';
  }
  if (stopBtn) {
    stopBtn.disabled = !terrain3dFlightState.running && terrain3dFlightState.progressMeters <= 0;
  }
}

function terrain3dComputeBearingDeg(lat1, lng1, lat2, lng2) {
  const phi1 = (lat1 * Math.PI) / 180;
  const phi2 = (lat2 * Math.PI) / 180;
  const dLon = ((lng2 - lng1) * Math.PI) / 180;
  const y = Math.sin(dLon) * Math.cos(phi2);
  const x = (Math.cos(phi1) * Math.sin(phi2)) - (Math.sin(phi1) * Math.cos(phi2) * Math.cos(dLon));
  const brng = (Math.atan2(y, x) * 180) / Math.PI;
  return (brng + 360) % 360;
}

function terrain3dBuildFlightPath(designPath, groundPath, Cesium) {
  const src = designPath.length >= 2
    ? designPath
    : groundPath.map((pt) => ({ ...pt, height: 24 }));
  if (src.length < 2) return [];

  const prepared = src.map((pt) => ({
    lat: terrain3dSafeNum(pt.lat, Number.NaN),
    lng: terrain3dSafeNum(pt.lng, Number.NaN),
    height: terrain3dSafeNum(pt.height, 18),
  })).filter((pt) => Number.isFinite(pt.lat) && Number.isFinite(pt.lng));

  return prepared.map((pt) => ({
    ...pt,
    cartesian: Cesium.Cartesian3.fromDegrees(pt.lng, pt.lat, pt.height + 8),
  }));
}

function terrain3dUpdateFlightDistances(Cesium) {
  const cumulative = [0];
  for (let i = 1; i < terrain3dFlightPath.length; i += 1) {
    const prev = terrain3dFlightPath[i - 1].cartesian;
    const curr = terrain3dFlightPath[i].cartesian;
    const segment = Cesium.Cartesian3.distance(prev, curr);
    cumulative.push(cumulative[i - 1] + Math.max(segment, 0));
  }
  terrain3dFlightState.cumulativeDistances = cumulative;
  terrain3dFlightState.totalDistance = cumulative[cumulative.length - 1] || 0;
}

function terrain3dSampleFlightAtDistance(distanceMeters, Cesium) {
  if (!terrain3dFlightPath.length) return null;
  if (distanceMeters <= 0) return terrain3dFlightPath[0];
  if (distanceMeters >= terrain3dFlightState.totalDistance) {
    return terrain3dFlightPath[terrain3dFlightPath.length - 1];
  }

  const cumulative = terrain3dFlightState.cumulativeDistances;
  let low = 0;
  let high = cumulative.length - 1;
  while (low <= high) {
    const mid = Math.floor((low + high) / 2);
    if (cumulative[mid] < distanceMeters) low = mid + 1;
    else if (cumulative[mid] > distanceMeters) high = mid - 1;
    else return terrain3dFlightPath[mid];
  }

  const idx = Math.max(1, low);
  const prevD = cumulative[idx - 1];
  const nextD = cumulative[idx];
  const t = (distanceMeters - prevD) / Math.max(nextD - prevD, 1e-6);
  const p1 = terrain3dFlightPath[idx - 1];
  const p2 = terrain3dFlightPath[idx];

  const out = new Cesium.Cartesian3();
  Cesium.Cartesian3.lerp(p1.cartesian, p2.cartesian, t, out);
  return {
    lat: p1.lat + ((p2.lat - p1.lat) * t),
    lng: p1.lng + ((p2.lng - p1.lng) * t),
    height: p1.height + ((p2.height - p1.height) * t),
    cartesian: out,
  };
}

function terrain3dStopFlyThrough(resetToStart = false) {
  terrain3dFlightState.running = false;
  if (terrain3dFlightState.rafId) {
    cancelAnimationFrame(terrain3dFlightState.rafId);
    terrain3dFlightState.rafId = null;
  }
  terrain3dFlightState.lastTs = 0;
  if (resetToStart) {
    terrain3dFlightState.progressMeters = 0;
  }
  terrain3dSyncFlightControls();
}

function terrain3dFlyThroughTick(ts) {
  if (!terrain3dFlightState.running || !terrain3dViewer || !window.Cesium) return;
  if (state?.activeWorkPage !== "terrain-3d-map") {
    terrain3dStopFlyThrough(false);
    return;
  }

  const Cesium = window.Cesium;
  if (!terrain3dFlightState.lastTs) terrain3dFlightState.lastTs = ts;
  const dtSec = Math.min((ts - terrain3dFlightState.lastTs) / 1000, 0.12);
  terrain3dFlightState.lastTs = ts;
  terrain3dFlightState.progressMeters += terrain3dFlightState.speedMps * dtSec;
  terrain3dFlightState.progressMeters = Math.min(terrain3dFlightState.progressMeters, terrain3dFlightState.totalDistance);

  const current = terrain3dSampleFlightAtDistance(terrain3dFlightState.progressMeters, Cesium);
  const ahead = terrain3dSampleFlightAtDistance(
    Math.min(terrain3dFlightState.progressMeters + 70, terrain3dFlightState.totalDistance),
    Cesium
  ) || current;
  if (!current || !ahead) {
    terrain3dStopFlyThrough(false);
    return;
  }

  const headingDeg = terrain3dComputeBearingDeg(current.lat, current.lng, ahead.lat, ahead.lng);
  const destination = Cesium.Cartesian3.fromDegrees(current.lng, current.lat, current.height + 9);
  terrain3dViewer.camera.setView({
    destination,
    orientation: {
      heading: Cesium.Math.toRadians(headingDeg),
      pitch: Cesium.Math.toRadians(-12),
      roll: 0,
    },
  });

  if (terrain3dFlightState.progressMeters >= terrain3dFlightState.totalDistance - 0.5) {
    terrain3dSetStatus("Fly-through completed.", "ready");
    terrain3dStopFlyThrough(false);
    return;
  }

  terrain3dFlightState.rafId = requestAnimationFrame(terrain3dFlyThroughTick);
}

function terrain3dStartOrPauseFlyThrough() {
  if (!terrain3dViewer || !window.Cesium) return;
  if (terrain3dFlightPath.length < 2 || terrain3dFlightState.totalDistance <= 1) {
    terrain3dSetStatus("No valid alignment route to start fly-through.", "warning");
    return;
  }

  if (terrain3dFlightState.running) {
    terrain3dSetStatus("Fly-through paused.", "warning");
    terrain3dStopFlyThrough(false);
    return;
  }

  terrain3dFlightState.running = true;
  terrain3dFlightState.lastTs = 0;
  terrain3dSyncFlightControls();
  terrain3dSetStatus(`Fly-through running at ${Math.round(terrain3dFlightState.speedMps)} m/s.`, "ready");
  terrain3dFlightState.rafId = requestAnimationFrame(terrain3dFlyThroughTick);
}

function terrain3dEnsureCesiumLoaded() {
  if (window.Cesium) return Promise.resolve(window.Cesium);
  if (terrain3dCesiumLoadPromise) return terrain3dCesiumLoadPromise;

  terrain3dCesiumLoadPromise = new Promise((resolve, reject) => {
    if (!document.getElementById("terrain3dCesiumCss")) {
      const link = document.createElement("link");
      link.id = "terrain3dCesiumCss";
      link.rel = "stylesheet";
      link.href = TERRAIN3D_CESIUM_CSS_URL;
      document.head.appendChild(link);
    }

    if (window.CESIUM_BASE_URL == null) {
      window.CESIUM_BASE_URL = TERRAIN3D_CESIUM_BASE_URL;
    }

    const existingScript = document.getElementById("terrain3dCesiumJs");
    if (existingScript) {
      existingScript.addEventListener("load", () => resolve(window.Cesium));
      existingScript.addEventListener("error", () => reject(new Error("Failed to load Cesium script.")));
      return;
    }

    const script = document.createElement("script");
    script.id = "terrain3dCesiumJs";
    script.src = TERRAIN3D_CESIUM_JS_URL;
    script.async = true;
    script.onload = () => {
      if (window.Cesium) resolve(window.Cesium);
      else reject(new Error("Cesium loaded but global object is unavailable."));
    };
    script.onerror = () => reject(new Error("Failed to load Cesium from CDN."));
    document.head.appendChild(script);
  }).catch((error) => {
    terrain3dCesiumLoadPromise = null;
    throw error;
  });

  return terrain3dCesiumLoadPromise;
}

async function terrain3dEnsureViewer() {
  const mapEl = document.getElementById("terrain3dMapContainer");
  if (!mapEl) return null;
  const Cesium = await terrain3dEnsureCesiumLoaded();

  if (terrain3dViewer) return terrain3dViewer;

  terrain3dViewer = new Cesium.Viewer(mapEl, {
    animation: false,
    timeline: false,
    geocoder: false,
    homeButton: true,
    fullscreenButton: false,
    sceneModePicker: true,
    navigationHelpButton: false,
    baseLayerPicker: false,
    selectionIndicator: false,
    infoBox: false,
    terrainProvider: new Cesium.EllipsoidTerrainProvider(),
  });

  terrain3dApplyMapType(terrain3dMapType, false);

  terrain3dViewer.scene.globe.depthTestAgainstTerrain = false;
  terrain3dViewer.scene.globe.enableLighting = false;
  terrain3dViewer.scene.globe.showGroundAtmosphere = false;
  terrain3dViewer.scene.globe.baseColor = Cesium.Color.fromCssColorString("#1e293b");
  if (terrain3dViewer.scene.skyAtmosphere) terrain3dViewer.scene.skyAtmosphere.show = false;
  if (terrain3dViewer.scene.skyBox) terrain3dViewer.scene.skyBox.show = false;
  if (terrain3dViewer.scene.sun) terrain3dViewer.scene.sun.show = false;
  if (terrain3dViewer.scene.moon) terrain3dViewer.scene.moon.show = false;
  terrain3dViewer.scene.backgroundColor = Cesium.Color.fromCssColorString("#0b1220");
  await terrain3dApplyTerrainProvider();

  terrain3dInitialized = true;
  return terrain3dViewer;
}

function terrain3dAddAlignmentEntities(Cesium, groundPath, designPath) {
  if (!terrain3dViewer) return;
  terrain3dViewer.entities.removeAll();
  terrain3dBuildingEntities = [];
  if (!terrain3dBuildingsTileset) terrain3dBuildingsEnabled = false;

  const groundDegrees = [];
  groundPath.forEach((pt) => {
    groundDegrees.push(pt.lng, pt.lat);
  });

  if (groundDegrees.length >= 4) {
    terrain3dViewer.entities.add({
      name: "Ground Alignment Trace",
      polyline: {
        positions: Cesium.Cartesian3.fromDegreesArray(groundDegrees),
        width: 2,
        clampToGround: true,
        material: Cesium.Color.fromCssColorString("#f97316").withAlpha(0.95),
      },
    });
  }

  if (designPath.length >= 2) {
    const designGroundDegrees = [];
    designPath.forEach((pt) => {
      designGroundDegrees.push(pt.lng, pt.lat);
    });
    terrain3dViewer.entities.add({
      name: "Design Alignment (Ground Projection)",
      polyline: {
        positions: Cesium.Cartesian3.fromDegreesArray(designGroundDegrees),
        width: 3,
        clampToGround: true,
        material: Cesium.Color.fromCssColorString("#38bdf8").withAlpha(0.78),
      },
    });

    const designDegreesHeights = [];
    designPath.forEach((pt) => {
      designDegreesHeights.push(pt.lng, pt.lat, pt.height);
    });

    terrain3dViewer.entities.add({
      name: "Design 3D Alignment",
      polyline: {
        positions: Cesium.Cartesian3.fromDegreesArrayHeights(designDegreesHeights),
        width: 5,
        material: new Cesium.PolylineGlowMaterialProperty({
          glowPower: 0.24,
          color: Cesium.Color.fromCssColorString("#22d3ee"),
        }),
      },
    });

    const start = designPath[0];
    const end = designPath[designPath.length - 1];
    terrain3dViewer.entities.add({
      name: "Start",
      position: Cesium.Cartesian3.fromDegrees(start.lng, start.lat, start.height),
      point: {
        pixelSize: 10,
        color: Cesium.Color.LIME,
        outlineColor: Cesium.Color.BLACK,
        outlineWidth: 1.5,
      },
      label: {
        text: "START",
        font: "12px Outfit, sans-serif",
        fillColor: Cesium.Color.WHITE,
        outlineColor: Cesium.Color.BLACK,
        outlineWidth: 2,
        style: Cesium.LabelStyle.FILL_AND_OUTLINE,
        pixelOffset: new Cesium.Cartesian2(0, -18),
      },
    });
    terrain3dViewer.entities.add({
      name: "End",
      position: Cesium.Cartesian3.fromDegrees(end.lng, end.lat, end.height),
      point: {
        pixelSize: 10,
        color: Cesium.Color.RED,
        outlineColor: Cesium.Color.BLACK,
        outlineWidth: 1.5,
      },
      label: {
        text: "END",
        font: "12px Outfit, sans-serif",
        fillColor: Cesium.Color.WHITE,
        outlineColor: Cesium.Color.BLACK,
        outlineWidth: 2,
        style: Cesium.LabelStyle.FILL_AND_OUTLINE,
        pixelOffset: new Cesium.Cartesian2(0, -18),
      },
    });
  }
}

async function terrain3dFlyToAlignment() {
  if (!terrain3dViewer || !window.Cesium) return;
  const Cesium = window.Cesium;
  try {
    const bbox = terrain3dGetAlignmentBbox();
    if (bbox) {
      const rectangle = Cesium.Rectangle.fromDegrees(bbox.west, bbox.south, bbox.east, bbox.north);
      await terrain3dViewer.camera.flyTo({
        destination: rectangle,
        duration: 1.2,
      });
    }

    const entities = terrain3dViewer.entities.values;
    if (!entities.length) return;
    await terrain3dViewer.flyTo(terrain3dViewer.entities, {
      duration: 1.6,
    });
  } catch (error) {
    console.warn("Cesium flyTo failed:", error);
  }
}

function terrain3dClearBuildingEntities() {
  if (!terrain3dViewer) return;
  if (terrain3dBuildingsTileset) {
    terrain3dViewer.scene.primitives.remove(terrain3dBuildingsTileset);
    terrain3dBuildingsTileset = null;
  }
  if (terrain3dBuildingEntities.length) {
    terrain3dBuildingEntities.forEach((entity) => terrain3dViewer.entities.remove(entity));
    terrain3dBuildingEntities = [];
  }
}

function terrain3dGetAlignmentBbox() {
  const points = Array.isArray(state?.kmlData?.points) ? state.kmlData.points : [];
  if (!points.length) return null;
  let minLat = Infinity, maxLat = -Infinity, minLng = Infinity, maxLng = -Infinity;
  points.forEach((pt) => {
    const lat = terrain3dSafeNum(pt?.lat, Number.NaN);
    const lng = terrain3dSafeNum(pt?.lng, Number.NaN);
    if (!Number.isFinite(lat) || !Number.isFinite(lng)) return;
    minLat = Math.min(minLat, lat);
    maxLat = Math.max(maxLat, lat);
    minLng = Math.min(minLng, lng);
    maxLng = Math.max(maxLng, lng);
  });
  if (!Number.isFinite(minLat) || !Number.isFinite(minLng)) return null;
  const padLat = Math.max((maxLat - minLat) * 0.2, 0.01);
  const padLng = Math.max((maxLng - minLng) * 0.2, 0.01);
  return {
    south: minLat - padLat,
    west: minLng - padLng,
    north: maxLat + padLat,
    east: maxLng + padLng,
  };
}

function terrain3dParseBuildingHeight(tags) {
  const heightRaw = String(tags?.height || "").replace(/m/gi, "").trim();
  const direct = terrain3dSafeNum(heightRaw, Number.NaN);
  if (Number.isFinite(direct) && direct > 1) return direct;
  const levels = terrain3dSafeNum(tags?.["building:levels"], Number.NaN);
  if (Number.isFinite(levels) && levels > 0) return Math.max(6, levels * 3.2);
  return 12;
}

async function terrain3dLoadOverpassBuildings(Cesium) {
  const bbox = terrain3dGetAlignmentBbox();
  if (!bbox) throw new Error("No alignment bbox available for building fetch.");
  const query = `[out:json][timeout:25];(way["building"](${bbox.south},${bbox.west},${bbox.north},${bbox.east});relation["building"](${bbox.south},${bbox.west},${bbox.north},${bbox.east}););out geom tags;`;
  const url = `https://overpass-api.de/api/interpreter?data=${encodeURIComponent(query)}`;
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Overpass HTTP ${response.status}`);
  const data = await response.json();
  const elements = Array.isArray(data?.elements) ? data.elements : [];
  const buildingCandidates = elements.filter((el) => Array.isArray(el?.geometry) && el.geometry.length >= 3);
  const selected = buildingCandidates.slice(0, TERRAIN3D_BUILDINGS_MAX_FEATURES);

  selected.forEach((feature) => {
    const coords = [];
    feature.geometry.forEach((p) => {
      const lat = terrain3dSafeNum(p?.lat, Number.NaN);
      const lon = terrain3dSafeNum(p?.lon, Number.NaN);
      if (!Number.isFinite(lat) || !Number.isFinite(lon)) return;
      coords.push(lon, lat);
    });
    if (coords.length < 6) return;
    const entity = terrain3dViewer.entities.add({
      polygon: {
        hierarchy: Cesium.Cartesian3.fromDegreesArray(coords),
        height: 0,
        extrudedHeight: terrain3dParseBuildingHeight(feature?.tags),
        material: Cesium.Color.fromCssColorString("#94a3b8").withAlpha(0.38),
        outline: true,
        outlineColor: Cesium.Color.fromCssColorString("#334155").withAlpha(0.75),
      },
    });
    terrain3dBuildingEntities.push(entity);
  });
  return terrain3dBuildingEntities.length;
}

async function terrain3dToggleBuildings() {
  const btn = document.getElementById("terrain3dBuildingsBtn");
  const Cesium = window.Cesium;
  if (!terrain3dViewer || !Cesium) return;
  if (terrain3dBuildingsEnabled) {
    terrain3dBuildingsEnabled = false;
    if (terrain3dBuildingsTileset) terrain3dBuildingsTileset.show = false;
    terrain3dBuildingEntities.forEach((entity) => { entity.show = false; });
    if (btn) btn.innerHTML = '<i class="ri-community-line" style="margin-right:4px;"></i> Show OSM Buildings';
    terrain3dSetStatus("Buildings hidden.", "warning");
    return;
  }

  if (terrain3dBuildingsTileset) {
    terrain3dBuildingsEnabled = true;
    terrain3dBuildingsTileset.show = true;
    if (btn) btn.innerHTML = '<i class="ri-community-line" style="margin-right:4px;"></i> Hide OSM Buildings';
    terrain3dSetStatus("OSM buildings visible.", "ready");
    return;
  }

  if (terrain3dBuildingEntities.length) {
    terrain3dBuildingsEnabled = true;
    terrain3dBuildingEntities.forEach((entity) => { entity.show = true; });
    if (btn) btn.innerHTML = '<i class="ri-community-line" style="margin-right:4px;"></i> Hide OSM Buildings';
    terrain3dSetStatus("Building overlay visible.", "ready");
    return;
  }

  try {
    terrain3dBuildingsTileset = await Cesium.createOsmBuildingsAsync();
    terrain3dViewer.scene.primitives.add(terrain3dBuildingsTileset);
    terrain3dBuildingsEnabled = true;
    if (btn) btn.innerHTML = '<i class="ri-community-line" style="margin-right:4px;"></i> Hide OSM Buildings';
    terrain3dSetStatus("OSM buildings loaded.", "ready");
    return;
  } catch (error) {
    console.warn("Cesium OSM buildings unavailable, using Overpass fallback:", error);
  }

  try {
    const loaded = await terrain3dLoadOverpassBuildings(Cesium);
    terrain3dBuildingsEnabled = loaded > 0;
    if (terrain3dBuildingsEnabled) {
      if (btn) btn.innerHTML = '<i class="ri-community-line" style="margin-right:4px;"></i> Hide OSM Buildings';
      terrain3dSetStatus(`Loaded ${loaded} building footprints (fallback mode).`, "ready");
    } else {
      terrain3dSetStatus("No nearby buildings found for this alignment.", "warning");
    }
  } catch (error) {
    console.warn("Building fallback failed:", error);
    terrain3dSetStatus("Buildings could not be loaded (both primary and fallback failed).", "warning");
  }

  if (btn) {
    btn.innerHTML = terrain3dBuildingsEnabled
      ? '<i class="ri-community-line" style="margin-right:4px;"></i> Hide OSM Buildings'
      : '<i class="ri-community-line" style="margin-right:4px;"></i> Show OSM Buildings';
  }
}

async function renderTerrain3DMapPage(force = false) {
  const kmlPoints = Array.isArray(state?.kmlData?.points) ? state.kmlData.points : [];
  const calcRows = Array.isArray(state?.calcRows) ? state.calcRows : [];
  const dataKey = terrain3dBuildDataKey(kmlPoints, calcRows);

  if (!kmlPoints.length) {
    terrain3dResetFlightState();
    terrain3dSetVisibility(false);
    terrain3dSetStatus("Import KML/KMZ in Alignment Map first. 3D terrain rendering starts after alignment data is available.", "warning");
    return;
  }

  terrain3dSetVisibility(true);

  try {
    await terrain3dEnsureViewer();
  } catch (error) {
    terrain3dSetVisibility(false);
    terrain3dSetStatus("Cesium failed to load. Check internet connection and try Refresh.", "warning");
    console.error(error);
    return;
  }

  if (!force && terrain3dLastDataKey === dataKey && terrain3dInitialized) {
    terrain3dViewer.resize();
    return;
  }
  terrain3dStopFlyThrough(false);
  terrain3dFlightState.progressMeters = 0;

  const groundPath = terrain3dBuildGroundPath(kmlPoints);
  const designPath = terrain3dBuildDesignPath(kmlPoints, calcRows);
  terrain3dAddAlignmentEntities(window.Cesium, groundPath, designPath);
  terrain3dFlightPath = terrain3dBuildFlightPath(designPath, groundPath, window.Cesium);
  terrain3dUpdateFlightDistances(window.Cesium);
  terrain3dSyncFlightControls();

  if (designPath.length >= 2) {
    terrain3dSetStatus(`Rendered ${designPath.length} design samples. Fly-through route: ${Math.round(terrain3dFlightState.totalDistance)} m.`, "ready");
  } else {
    terrain3dSetStatus("KML alignment rendered. Run calculation/verify for full design elevations. Fly-through uses a fallback route for now.", "warning");
  }

  terrain3dLastDataKey = dataKey;
  terrain3dViewer.resize();
  setTimeout(() => terrain3dViewer && terrain3dViewer.resize(), 100);
  await terrain3dFlyToAlignment();
}

window.renderTerrain3DMapPage = renderTerrain3DMapPage;

document.addEventListener("DOMContentLoaded", () => {
  const refreshBtn = document.getElementById("terrain3dRefreshBtn");
  const flyBtn = document.getElementById("terrain3dFlyBtn");
  const startFlyBtn = document.getElementById("terrain3dStartFlyBtn");
  const stopFlyBtn = document.getElementById("terrain3dStopFlyBtn");
  const mapTypeSelect = document.getElementById("terrain3dMapTypeSelect");
  const terrainSourceSelect = document.getElementById("terrain3dTerrainSourceSelect");
  const speedSelect = document.getElementById("terrain3dSpeedSelect");
  const buildingsBtn = document.getElementById("terrain3dBuildingsBtn");

  if (refreshBtn) {
    refreshBtn.addEventListener("click", () => {
      renderTerrain3DMapPage(true);
    });
  }
  if (flyBtn) {
    flyBtn.addEventListener("click", () => {
      terrain3dFlyToAlignment();
    });
  }
  if (startFlyBtn) {
    startFlyBtn.addEventListener("click", () => {
      terrain3dStartOrPauseFlyThrough();
    });
  }
  if (stopFlyBtn) {
    stopFlyBtn.addEventListener("click", () => {
      terrain3dSetStatus("Fly-through stopped.", "warning");
      terrain3dStopFlyThrough(true);
      terrain3dFlyToAlignment();
    });
  }
  if (mapTypeSelect) {
    terrain3dLoadStoredMapType();
    mapTypeSelect.value = terrain3dMapType;
    mapTypeSelect.addEventListener("change", () => {
      terrain3dApplyMapType(mapTypeSelect.value, true);
    });
  }
  terrain3dLoadStoredTerrainSource();
  if (terrainSourceSelect) {
    terrainSourceSelect.value = terrain3dTerrainSource;
    terrainSourceSelect.addEventListener("change", () => {
      terrain3dApplyTerrainSource(terrainSourceSelect.value, true);
    });
  }
  if (speedSelect) {
    terrain3dFlightState.speedMps = terrain3dSafeNum(speedSelect.value, 70);
    speedSelect.addEventListener("change", () => {
      terrain3dFlightState.speedMps = Math.max(10, terrain3dSafeNum(speedSelect.value, 70));
      if (terrain3dFlightState.running) {
        terrain3dSetStatus(`Fly-through speed updated to ${Math.round(terrain3dFlightState.speedMps)} m/s.`, "ready");
      }
    });
  }
  if (buildingsBtn) {
    buildingsBtn.addEventListener("click", () => {
      terrain3dToggleBuildings();
    });
  }
  terrain3dSyncFlightControls();
});
