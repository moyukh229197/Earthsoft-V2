export default async function handler(req, res) {
  const maptilerKey = String(
    process.env.MAPTILER_API_KEY
    || process.env.EARTHSOFT_MAPTILER_API_KEY
    || process.env.VITE_MAPTILER_API_KEY
    || ""
  ).trim();

  const googleMapsKey = String(
    process.env.EARTHSOFT_GOOGLE_MAPS_API_KEY
    || process.env.GOOGLE_MAPS_API_KEY
    || process.env.VITE_GOOGLE_MAPS_API_KEY
    || ""
  ).trim();

  const googleMapId = String(
    process.env.EARTHSOFT_GOOGLE_MAP_ID
    || process.env.GOOGLE_MAP_ID
    || process.env.VITE_GOOGLE_MAP_ID
    || ""
  ).trim();

  res.setHeader("Content-Type", "application/javascript; charset=utf-8");
  res.setHeader("Cache-Control", "no-store, max-age=0");

  res.status(200).send(`
window.MAPTILER_API_KEY = ${JSON.stringify(maptilerKey)};
window.EARTHSOFT_GOOGLE_MAPS_API_KEY = ${JSON.stringify(googleMapsKey)};
window.EARTHSOFT_GOOGLE_MAP_ID = ${JSON.stringify(googleMapId)};
`);
}

