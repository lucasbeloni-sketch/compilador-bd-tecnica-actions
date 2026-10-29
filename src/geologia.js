// ---- Classificacao geologica (dureza de perfuracao) — port do Apps Script classificarSoloRochoso. ----
// Funcoes puras. AZ/BA trazem UTM SIRGAS 2000 fuso 24S (ou graus decimais); converte pra
// lat/lon e testa em qual poligono do Mapa Geologico do Ceara (SGB/CPRM 1:500.000, 2021) cai.
// Dados em data/ceara_geo.json: [categoria, minX, minY, maxX, maxY, [aneis...]], graus x 10000.
const CEARA_GEO = require("../data/ceara_geo.json");

const ESCALA = 10000;

const ROTULOS = {
  HARD: "DIFÍCIL — rocha dura",
  MISTA: "MÉDIA — pode ter rocha dura",
  SED: "FÁCIL — terreno mole (sedimentar)",
  SURF: "FÁCIL — areia/aluvião (sem rocha)",
  FORA: "CONFERIR — sem dado no mapa",
};

function parseCoord(v) {
  if (v === "" || v === null || v === undefined) return null;
  if (typeof v === "number") return v;
  const f = parseFloat(String(v).trim().replace(/\s/g, "").replace(",", "."));
  return isNaN(f) ? null : f;
}

// Decide qual valor e Este (<1.000.000) e qual e Norte (>=1.000.000).
function identificarEN(v1, v2) {
  let E, N;
  if (Math.abs(v1) >= 1000000 && Math.abs(v2) < 1000000) { N = v1; E = v2; }
  else if (Math.abs(v2) >= 1000000 && Math.abs(v1) < 1000000) { N = v2; E = v1; }
  else return null; // ex.: par "1 / 1"
  if (E < 100000 || E > 900000 || N < 1000000 || N > 10000000) return null;
  return { E, N };
}

// UTM -> lat/lon (SIRGAS 2000 / GRS80), hemisferio sul.
function utmParaLatLon(E, N, fuso) {
  const a = 6378137.0, f = 1 / 298.257222101, k0 = 0.9996;
  const e2 = f * (2 - f);
  const e1 = (1 - Math.sqrt(1 - e2)) / (1 + Math.sqrt(1 - e2));
  const ep2 = e2 / (1 - e2);
  const lon0 = ((fuso * 6 - 183) * Math.PI) / 180;
  const x = E - 500000.0;
  const y = N - 10000000.0;
  const M = y / k0;
  const mu = M / (a * (1 - e2 / 4 - (3 * e2 * e2) / 64 - (5 * e2 * e2 * e2) / 256));
  const phi1 = mu
    + ((3 * e1) / 2 - (27 * Math.pow(e1, 3)) / 32) * Math.sin(2 * mu)
    + ((21 * e1 * e1) / 16 - (55 * Math.pow(e1, 4)) / 32) * Math.sin(4 * mu)
    + ((151 * Math.pow(e1, 3)) / 96) * Math.sin(6 * mu)
    + ((1097 * Math.pow(e1, 4)) / 512) * Math.sin(8 * mu);
  const sp = Math.sin(phi1), cp = Math.cos(phi1), tp = Math.tan(phi1);
  const C1 = ep2 * cp * cp, T1 = tp * tp;
  const N1 = a / Math.sqrt(1 - e2 * sp * sp);
  const R1 = (a * (1 - e2)) / Math.pow(1 - e2 * sp * sp, 1.5);
  const D = x / (N1 * k0);
  const lat = phi1 - ((N1 * tp) / R1) * ((D * D) / 2
    - ((5 + 3 * T1 + 10 * C1 - 4 * C1 * C1 - 9 * ep2) * Math.pow(D, 4)) / 24
    + ((61 + 90 * T1 + 298 * C1 + 45 * T1 * T1 - 252 * ep2 - 3 * C1 * C1) * Math.pow(D, 6)) / 720);
  const lon = lon0 + (D - ((1 + 2 * T1 + C1) * Math.pow(D, 3)) / 6
    + ((5 - 2 * C1 + 28 * T1 - 3 * C1 * C1 + 8 * ep2 + 24 * T1 * T1) * Math.pow(D, 5)) / 120) / cp;
  return [(lat * 180) / Math.PI, (lon * 180) / Math.PI];
}

// Aceita UTM (metros) ou graus decimais. null pra par vazio/invalido.
function obterLatLon(v1, v2, fuso) {
  if (v1 === null || v2 === null) return null;
  const a1 = Math.abs(v1), a2 = Math.abs(v2);
  if (a1 >= 1000000 || a2 >= 1000000) {
    const en = identificarEN(v1, v2);
    return en ? utmParaLatLon(en.E, en.N, fuso) : null;
  }
  if (a1 <= 90 && a2 <= 90) {
    let lat, lon;
    if (a1 <= 20 && a2 >= 20) { lat = v1; lon = v2; }
    else if (a2 <= 20 && a1 >= 20) { lat = v2; lon = v1; }
    else { lat = v1; lon = v2; }
    // Sem sinal: Nordeste e sempre Sul/Oeste.
    if (lat > 0) lat = -lat;
    if (lon > 0) lon = -lon;
    if (lat < -34 || lat > 6 || lon < -74 || lon > -34) return null;
    return [lat, lon];
  }
  return null;
}

function pontoNoAnel(r, x, y) {
  let dentro = false;
  const n = r.length / 2;
  for (let i = 0, j = n - 1; i < n; j = i++) {
    const xi = r[2 * i], yi = r[2 * i + 1], xj = r[2 * j], yj = r[2 * j + 1];
    if ((yi > y) !== (yj > y) && x < ((xj - xi) * (y - yi)) / (yj - yi) + xi) dentro = !dentro;
  }
  return dentro;
}

function pontoNaFeicao(aneis, x, y) {
  if (!pontoNoAnel(aneis[0], x, y)) return false;
  for (let k = 1; k < aneis.length; k++) if (pontoNoAnel(aneis[k], x, y)) return false;
  return true;
}

function classificarLatLon(lat, lon) {
  const x = Math.round(lon * ESCALA);
  const y = Math.round(lat * ESCALA);
  for (const f of CEARA_GEO) {
    if (x < f[1] || x > f[3] || y < f[2] || y > f[4]) continue;
    if (pontoNaFeicao(f[5], x, y)) return f[0];
  }
  return "FORA";
}

// Rotulo pra um par (AZ, BA). "" quando sem coordenada valida.
function classificarPar(a, b, fuso = 24) {
  const ll = obterLatLon(parseCoord(a), parseCoord(b), fuso);
  if (!ll) return "";
  const cat = classificarLatLon(ll[0], ll[1]);
  return ROTULOS[cat] || cat;
}

module.exports = { ROTULOS, parseCoord, obterLatLon, utmParaLatLon, classificarLatLon, classificarPar };
