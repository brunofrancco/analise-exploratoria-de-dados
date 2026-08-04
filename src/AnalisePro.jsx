import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import _ from "lodash";
import {
  BarChart, Bar, AreaChart, Area, ScatterChart, Scatter,
  PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer, ComposedChart, Line, Treemap, ReferenceLine, ZAxis,
  RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar as RadarSeries,
} from "recharts";
import {
  Upload, LayoutDashboard, BarChart3, Sigma, FlaskConical, FolderKanban,
  FileText, AlertTriangle, Info, Trash2, Download,
  Plus, Activity, Gauge, Sparkles,
  ChevronDown, Sun, Moon, Clock, Layers, Lightbulb, CheckCircle2, XCircle,
  TrendingUp, ClipboardList, FileSpreadsheet, FileCode, Printer,
  RefreshCw, Share2, Filter, Calendar, ShieldCheck, ShieldAlert, TrendingDown,
  ArrowUpRight, ArrowDownRight, Zap, Target, Rocket, Flame, Radar as RadarIcon,
  Maximize2, Compass,
  FileBarChart2, GitCompare, ExternalLink, Beaker, Wand2, Table2, ListChecks,
} from "lucide-react";

/* =========================================================================
   BACKEND API (Python) — usado pelas abas de bibliotecas de análise
   automática (ydata-profiling, Sweetviz). Ajuste esta URL após o deploy do
   serviço no Render.
   ========================================================================= */
const API_BASE = "https://analise-exploratoria-backend.onrender.com";

/* =========================================================================
   DESIGN TOKENS
   ========================================================================= */
const FONT_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');`;

const T = {
  bg: "#F6F7F9",
  panel: "#FFFFFF",
  border: "#E3E6EB",
  ink: "#12141C",
  sub: "#5B6270",
  faint: "#9AA1AE",
  teal: "#0B6E6E",
  tealDark: "#075454",
  tealSoft: "#E4F2F1",
  amber: "#9A5B12",
  amberSoft: "#FBEEDD",
  red: "#B42318",
  redSoft: "#FBEAE8",
  green: "#166A47",
  greenSoft: "#E6F3EC",
  blue: "#2B5C8A",
};

const HEAT_SCALE = (v) => {
  // v in [-1,1] -> diverging teal (neg) - white - amber (pos)
  const t = Math.max(-1, Math.min(1, v));
  if (t >= 0) {
    const k = t;
    const r = Math.round(255 - k * (255 - 154));
    const g = Math.round(255 - k * (255 - 91));
    const b = Math.round(255 - k * (255 - 18));
    return `rgb(${r},${g},${b})`;
  } else {
    const k = -t;
    const r = Math.round(255 - k * (255 - 11));
    const g = Math.round(255 - k * (255 - 110));
    const b = Math.round(255 - k * (255 - 110));
    return `rgb(${r},${g},${b})`;
  }
};

/* =========================================================================
   NUMERIC / STATISTICS ENGINE
   ========================================================================= */
const isNum = (v) => typeof v === "number" && !Number.isNaN(v);

function toNumberOrNull(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return Number.isNaN(v) ? null : v;
  const s = String(v).trim().replace(/\./g, "").replace(",", ".");
  const n1 = Number(String(v).trim());
  if (!Number.isNaN(n1) && String(v).trim() !== "") return n1;
  const n2 = Number(s);
  if (!Number.isNaN(n2) && s !== "") return n2;
  return null;
}

function mean(arr) { return arr.length ? _.sum(arr) / arr.length : NaN; }
function median(arr) {
  if (!arr.length) return NaN;
  const s = [...arr].sort((a, b) => a - b);
  const m = s.length;
  return m % 2 ? s[(m - 1) / 2] : (s[m / 2 - 1] + s[m / 2]) / 2;
}
function mode(arr) {
  if (!arr.length) return [];
  const counts = new Map();
  arr.forEach((v) => counts.set(v, (counts.get(v) || 0) + 1));
  const max = Math.max(...counts.values());
  if (max === 1) return [];
  return [...counts.entries()].filter(([, c]) => c === max).map(([v]) => v).slice(0, 3);
}
function variance(arr, sample = true) {
  if (arr.length < 2) return NaN;
  const m = mean(arr);
  const ss = _.sum(arr.map((v) => (v - m) ** 2));
  return ss / (arr.length - (sample ? 1 : 0));
}
function std(arr, sample = true) { return Math.sqrt(variance(arr, sample)); }
function cv(arr) { const m = mean(arr); return m !== 0 ? (std(arr) / Math.abs(m)) * 100 : NaN; }
function percentile(arr, p) {
  if (!arr.length) return NaN;
  const s = [...arr].sort((a, b) => a - b);
  const idx = (p / 100) * (s.length - 1);
  const lo = Math.floor(idx), hi = Math.ceil(idx);
  if (lo === hi) return s[lo];
  return s[lo] + (s[hi] - s[lo]) * (idx - lo);
}
function quartiles(arr) { return { q1: percentile(arr, 25), q2: percentile(arr, 50), q3: percentile(arr, 75) }; }
function iqr(arr) { const { q1, q3 } = quartiles(arr); return q3 - q1; }
function range(arr) { return arr.length ? Math.max(...arr) - Math.min(...arr) : NaN; }
function skewness(arr) {
  const n = arr.length; if (n < 3) return NaN;
  const m = mean(arr), s = std(arr, true);
  const g1 = (n / ((n - 1) * (n - 2))) * _.sum(arr.map((v) => ((v - m) / s) ** 3));
  return g1;
}
function kurtosisExcess(arr) {
  const n = arr.length; if (n < 4) return NaN;
  const m = mean(arr), s = std(arr, true);
  const term = (n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3));
  const sum4 = _.sum(arr.map((v) => ((v - m) / s) ** 4));
  const corr = (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
  return term * sum4 - corr;
}
function covariance(x, y) {
  const n = Math.min(x.length, y.length); if (n < 2) return NaN;
  const mx = mean(x), my = mean(y);
  let s = 0; for (let i = 0; i < n; i++) s += (x[i] - mx) * (y[i] - my);
  return s / (n - 1);
}
function pearson(x, y) {
  const n = Math.min(x.length, y.length); if (n < 2) return NaN;
  const cov = covariance(x, y), sx = std(x, true), sy = std(y, true);
  if (sx === 0 || sy === 0) return NaN;
  return cov / (sx * sy);
}
function rankArr(arr) {
  const idx = arr.map((v, i) => [v, i]).sort((a, b) => a[0] - b[0]);
  const ranks = new Array(arr.length);
  let i = 0;
  while (i < idx.length) {
    let j = i;
    while (j + 1 < idx.length && idx[j + 1][0] === idx[i][0]) j++;
    const avgRank = (i + j) / 2 + 1;
    for (let k = i; k <= j; k++) ranks[idx[k][1]] = avgRank;
    i = j + 1;
  }
  return ranks;
}
function spearman(x, y) { return pearson(rankArr(x), rankArr(y)); }

function outliersIQR(arr) {
  const { q1, q3 } = quartiles(arr);
  const box = q3 - q1;
  const lo = q1 - 1.5 * box, hi = q3 + 1.5 * box;
  return arr.filter((v) => v < lo || v > hi);
}

/* --- special functions (Numerical-Recipes-style approximations) --- */
function gammaln(x) {
  const cof = [76.18009172947146, -86.50532032941677, 24.01409824083091,
    -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5];
  let y = x, tmp = x + 5.5;
  tmp -= (x + 0.5) * Math.log(tmp);
  let ser = 1.000000000190015;
  for (let j = 0; j < 6; j++) { y += 1; ser += cof[j] / y; }
  return -tmp + Math.log((2.5066282746310005 * ser) / x);
}
function betacf(x, a, b) {
  const MAXIT = 200, EPS = 3e-7, FPMIN = 1e-30;
  const qab = a + b, qap = a + 1, qam = a - 1;
  let c = 1, d = 1 - (qab * x) / qap;
  if (Math.abs(d) < FPMIN) d = FPMIN;
  d = 1 / d; let h = d;
  for (let m = 1; m <= MAXIT; m++) {
    const m2 = 2 * m;
    let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2));
    d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d; h *= d * c;
    aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2));
    d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d; const del = d * c; h *= del;
    if (Math.abs(del - 1) < EPS) break;
  }
  return h;
}
function betai(a, b, x) {
  if (x <= 0) return 0; if (x >= 1) return 1;
  const bt = Math.exp(gammaln(a + b) - gammaln(a) - gammaln(b) + a * Math.log(x) + b * Math.log(1 - x));
  if (x < (a + 1) / (a + b + 2)) return (bt * betacf(x, a, b)) / a;
  return 1 - (bt * betacf(1 - x, b, a)) / b;
}
function tTwoTailP(t, df) { return betai(df / 2, 0.5, df / (df + t * t)); }
// Finds the critical t-value for a given two-tailed significance level (bisection over tTwoTailP).
function tInvTwoTail(pTarget, df) {
  let lo = 0, hi = 1000;
  for (let i = 0; i < 80; i++) {
    const mid = (lo + hi) / 2;
    if (tTwoTailP(mid, df) > pTarget) lo = mid; else hi = mid;
  }
  return (lo + hi) / 2;
}

/* --- incomplete gamma function (Numerical-Recipes-style) — powers the chi-square test --- */
function gser(a, x) {
  const ITMAX = 200, EPS = 3e-9;
  const gln = gammaln(a);
  if (x <= 0) return { gamser: 0, gln };
  let ap = a, sum = 1 / a, del = sum;
  for (let n = 1; n <= ITMAX; n++) {
    ap += 1; del *= x / ap; sum += del;
    if (Math.abs(del) < Math.abs(sum) * EPS) break;
  }
  return { gamser: sum * Math.exp(-x + a * Math.log(x) - gln), gln };
}
function gcf(a, x) {
  const ITMAX = 200, EPS = 3e-9, FPMIN = 1e-30;
  const gln = gammaln(a);
  let b = x + 1 - a, c = 1 / FPMIN, d = 1 / b, h = d;
  for (let i = 1; i <= ITMAX; i++) {
    const an = -i * (i - a);
    b += 2;
    d = an * d + b; if (Math.abs(d) < FPMIN) d = FPMIN;
    c = b + an / c; if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d; const del = d * c; h *= del;
    if (Math.abs(del - 1) < EPS) break;
  }
  return { gammcf: Math.exp(-x + a * Math.log(x) - gln) * h, gln };
}
function gammaP(a, x) { if (x < 0 || a <= 0) return NaN; return x < a + 1 ? gser(a, x).gamser : 1 - gcf(a, x).gammcf; }
function gammaQ(a, x) { if (x < 0 || a <= 0) return NaN; return x < a + 1 ? 1 - gser(a, x).gamser : gcf(a, x).gammcf; }
// Upper-tail p-value of the chi-square distribution — used by the chi-square test of independence.
function chiSquareUpperP(x2, df) { return df > 0 ? Math.max(0, Math.min(1, gammaQ(df / 2, x2 / 2))) : NaN; }
// Upper-tail p-value of the F distribution — used by ANOVA.
function fUpperP(F, d1, d2) { if (F <= 0 || d1 <= 0 || d2 <= 0) return 1; return betai(d2 / 2, d1 / 2, d2 / (d2 + d1 * F)); }

/* --- hypothesis tests & association measures (client-side, no backend needed) --- */
function twoSampleTTest(a, b) {
  const na = a.length, nb = b.length;
  if (na < 2 || nb < 2) return null;
  const ma = mean(a), mb = mean(b);
  const va = variance(a, true), vb = variance(b, true);
  const se = Math.sqrt(va / na + vb / nb);
  const t = se !== 0 ? (ma - mb) / se : 0;
  const df = (va / na + vb / nb) ** 2 / ((va / na) ** 2 / (na - 1) + (vb / nb) ** 2 / (nb - 1));
  const p = Number.isFinite(df) && df > 0 ? tTwoTailP(t, df) : NaN;
  const pooledSd = Math.sqrt(((na - 1) * va + (nb - 1) * vb) / (na + nb - 2));
  const cohensD = pooledSd !== 0 ? (ma - mb) / pooledSd : NaN;
  const tCrit = Number.isFinite(df) && df > 0 ? tInvTwoTail(0.05, df) : NaN;
  const ciLow = (ma - mb) - tCrit * se, ciHigh = (ma - mb) + tCrit * se;
  return { t, df, p, meanA: ma, meanB: mb, diff: ma - mb, ciLow, ciHigh, cohensD, na, nb };
}
function oneWayANOVA(groups) {
  const clean = groups.filter((g) => g.length >= 2);
  if (clean.length < 2) return null;
  const allVals = clean.flat();
  const grandMean = mean(allVals);
  const k = clean.length, n = allVals.length;
  const ssBetween = _.sum(clean.map((g) => g.length * (mean(g) - grandMean) ** 2));
  const ssWithin = _.sum(clean.map((g) => _.sum(g.map((v) => (v - mean(g)) ** 2))));
  const dfBetween = k - 1, dfWithin = n - k;
  const msBetween = ssBetween / dfBetween, msWithin = dfWithin > 0 ? ssWithin / dfWithin : NaN;
  const F = msWithin > 0 ? msBetween / msWithin : NaN;
  const p = Number.isFinite(F) ? fUpperP(F, dfBetween, dfWithin) : NaN;
  const ssTotal = ssBetween + ssWithin;
  const etaSquared = ssTotal !== 0 ? ssBetween / ssTotal : NaN;
  return { F, dfBetween, dfWithin, p, etaSquared, k, n, means: clean.map((g) => mean(g)) };
}
function chiSquareTest(catA, catB) {
  const n = Math.min(catA.length, catB.length);
  const levelsA = _.uniq(catA.slice(0, n)), levelsB = _.uniq(catB.slice(0, n));
  const table = levelsA.map(() => levelsB.map(() => 0));
  for (let i = 0; i < n; i++) {
    const ai = levelsA.indexOf(catA[i]), bi = levelsB.indexOf(catB[i]);
    table[ai][bi]++;
  }
  const rowSums = table.map((row) => _.sum(row));
  const colSums = levelsB.map((_b, j) => _.sum(table.map((row) => row[j])));
  let chi2 = 0;
  for (let i = 0; i < levelsA.length; i++) for (let j = 0; j < levelsB.length; j++) {
    const expected = (rowSums[i] * colSums[j]) / n;
    if (expected > 0) chi2 += (table[i][j] - expected) ** 2 / expected;
  }
  const df = (levelsA.length - 1) * (levelsB.length - 1);
  const p = df > 0 ? chiSquareUpperP(chi2, df) : NaN;
  const minDim = Math.min(levelsA.length, levelsB.length);
  const cramersV = df > 0 && minDim > 1 ? Math.sqrt(chi2 / (n * (minDim - 1))) : NaN;
  return { chi2, df, p, cramersV, table, levelsA, levelsB, rowSums, colSums, n };
}
function pointBiserial(binaryVals, numericVals) { return pearson(binaryVals, numericVals); }
function benjaminiHochberg(pvalues) {
  const m = pvalues.length;
  const idx = pvalues.map((p, i) => [Number.isNaN(p) ? 1 : p, i]).sort((a, b) => a[0] - b[0]);
  const adjusted = new Array(m);
  let prevMin = 1;
  for (let rank = m; rank >= 1; rank--) {
    const [p, origIdx] = idx[rank - 1];
    const val = Math.min(prevMin, (p * m) / rank);
    prevMin = val;
    adjusted[origIdx] = Math.min(1, val);
  }
  return adjusted;
}
function bonferroniCorrect(pvalues) { const m = pvalues.length; return pvalues.map((p) => Math.min(1, p * m)); }
function cohensDLabel(d) { const a = Math.abs(d); return a < 0.2 ? "desprezível" : a < 0.5 ? "pequeno" : a < 0.8 ? "médio" : "grande"; }
function cramersVLabel(v) { if (Number.isNaN(v)) return "—"; return v < 0.1 ? "desprezível" : v < 0.3 ? "moderada" : v < 0.5 ? "forte" : "muito forte"; }
function etaSquaredLabel(e) { if (Number.isNaN(e)) return "—"; return e < 0.01 ? "desprezível" : e < 0.06 ? "pequeno" : e < 0.14 ? "médio" : "grande"; }

// Discretizes a numeric array into quantile bins (used by mutual information for continuous variables).
function quantileBin(vals, nbins = 8) {
  const sorted = [...vals].sort((a, b) => a - b);
  const edges = [];
  for (let i = 1; i < nbins; i++) edges.push(percentile(sorted, (i / nbins) * 100));
  return vals.map((v) => { let b = 0; while (b < edges.length && v > edges[b]) b++; return b; });
}
// Normalized mutual information in [0,1] — captures non-linear associations Pearson/Spearman miss.
function mutualInformation(xVals, yVals, xNumeric, yNumeric) {
  const n = Math.min(xVals.length, yVals.length);
  if (n < 4) return { mi: NaN, normalized: NaN };
  const xb = (xNumeric ? quantileBin(xVals.slice(0, n)) : xVals.slice(0, n)).map(String);
  const yb = (yNumeric ? quantileBin(yVals.slice(0, n)) : yVals.slice(0, n)).map(String);
  const xCounts = {}, yCounts = {}, jointCounts = {};
  for (let i = 0; i < n; i++) {
    const jk = xb[i] + "" + yb[i];
    xCounts[xb[i]] = (xCounts[xb[i]] || 0) + 1;
    yCounts[yb[i]] = (yCounts[yb[i]] || 0) + 1;
    jointCounts[jk] = (jointCounts[jk] || 0) + 1;
  }
  let mi = 0;
  for (const jk in jointCounts) {
    const sepIdx = jk.indexOf("");
    const xk = jk.slice(0, sepIdx), yk = jk.slice(sepIdx + 1);
    const pxy = jointCounts[jk] / n, px = xCounts[xk] / n, py = yCounts[yk] / n;
    mi += pxy * Math.log2(pxy / (px * py));
  }
  const entropyOf = (counts) => -Object.values(counts).reduce((s, c) => { const p = c / n; return s + p * Math.log2(p); }, 0);
  const hx = entropyOf(xCounts), hy = entropyOf(yCounts);
  const norm = Math.min(hx, hy) > 0 ? mi / Math.min(hx, hy) : 0;
  return { mi, normalized: Math.max(0, Math.min(1, norm)) };
}

/* --- time-series diagnostics: (augmented) Dickey-Fuller stationarity test + ACF/PACF --- */
function adfTest(seriesVals) {
  const n = seriesVals.length;
  if (n < 12) return null;
  const y = seriesVals.slice(1);
  const yLag = seriesVals.slice(0, -1);
  const dy = y.map((v, i) => v - yLag[i]);
  const reg = multipleRegression(yLag.map((v) => [v]), dy);
  if (!reg) return null;
  const tStat = reg.coefStats[1]?.t;
  const critical = { "1%": -3.43, "5%": -2.86, "10%": -2.57 };
  return { tStat, critical, stationary: tStat < critical["5%"], n: dy.length };
}
function acfSeries(seriesVals, maxLag) {
  const n = seriesVals.length;
  const m = mean(seriesVals);
  const c0 = _.sum(seriesVals.map((v) => (v - m) ** 2));
  const out = [];
  for (let k = 0; k <= maxLag; k++) {
    let ck = 0;
    for (let t = 0; t < n - k; t++) ck += (seriesVals[t] - m) * (seriesVals[t + k] - m);
    out.push(c0 !== 0 ? ck / c0 : 0);
  }
  return out;
}
function pacfSeries(seriesVals, maxLag) {
  const rho = acfSeries(seriesVals, maxLag);
  const out = [1];
  let prevPhi = [];
  for (let k = 1; k <= maxLag; k++) {
    if (k === 1) { out.push(rho[1]); prevPhi = [rho[1]]; continue; }
    let num = rho[k];
    for (let j = 0; j < k - 1; j++) num -= prevPhi[j] * rho[k - 1 - j];
    let den = 1;
    for (let j = 0; j < k - 1; j++) den -= prevPhi[j] * rho[j + 1];
    const phik = den !== 0 ? num / den : 0;
    const newPhi = prevPhi.map((pj, j) => pj - phik * prevPhi[k - 2 - j]);
    newPhi.push(phik);
    out.push(phik);
    prevPhi = newPhi;
  }
  return out;
}

/* =========================================================================
   DATA LOADING / SCHEMA INFERENCE  (robust header + type detection)
   ========================================================================= */
function cleanHeaderName(raw, idx) {
  let s = String(raw ?? "").replace(/^\uFEFF/, "").trim().replace(/\s+/g, " ");
  if (!s || /^__EMPTY/.test(s)) s = `coluna_${idx + 1}`;
  return s;
}

// Normalizes raw parsed rows: trims/cleans header names, dedupes them,
// and builds the column list from the UNION of keys across a sample of
// rows (not just the first row) so sparse/irregular records don't lose columns.
function normalizeParsedRows(rawRows) {
  if (!rawRows.length) return { rows: [], fieldNames: [] };
  const sampleSize = Math.min(rawRows.length, 2000);
  const rawKeySet = [];
  const seenRaw = new Set();
  for (let i = 0; i < sampleSize; i++) {
    Object.keys(rawRows[i] || {}).forEach((k) => {
      if (!seenRaw.has(k)) { seenRaw.add(k); rawKeySet.push(k); }
    });
  }
  const usedClean = new Map();
  const keyMap = {}; // rawKey -> finalCleanName
  rawKeySet.forEach((rawKey, idx) => {
    let clean = cleanHeaderName(rawKey, idx);
    const count = usedClean.get(clean) || 0;
    usedClean.set(clean, count + 1);
    if (count > 0) clean = `${clean}_${count + 1}`;
    keyMap[rawKey] = clean;
  });
  const fieldNames = rawKeySet.map((k) => keyMap[k]);
  const rows = rawRows.map((r) => {
    const out = {};
    rawKeySet.forEach((rawKey) => { out[keyMap[rawKey]] = r ? r[rawKey] : undefined; });
    return out;
  });
  return { rows, fieldNames };
}

// Detects whether a numeric-looking column uses BR decimal style (1.234,56)
// or US/international style (1,234.56 or 12.5), based on a value sample.
function detectDecimalStyle(sampleValues) {
  let brScore = 0, usScore = 0;
  sampleValues.forEach((v) => {
    if (typeof v === "number") return;
    const s = String(v).trim();
    if (/^-?\d{1,3}(\.\d{3})+,\d+$/.test(s)) brScore += 3;
    else if (/^-?\d{1,3}(,\d{3})+\.\d+$/.test(s)) usScore += 3;
    else if (/^-?\d+,\d{1,2}$/.test(s)) brScore += 1;
    else if (/^-?\d+\.\d{1,}$/.test(s)) usScore += 1;
    else if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) brScore += 1; // 1.234 (thousands, no decimals)
  });
  return brScore > usScore ? "br" : "us";
}

function parseNumberWithStyle(v, style) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return Number.isNaN(v) ? null : v;
  let s = String(v).trim();
  if (!s) return null;
  s = s.replace(/^R\$\s*/i, "").replace(/%$/, "").trim();
  if (style === "br") s = s.replace(/\./g, "").replace(",", ".");
  else s = s.replace(/,/g, "");
  const n = Number(s);
  return Number.isNaN(n) ? null : n;
}

function tryParseDate(v) {
  if (v instanceof Date) return Number.isNaN(+v) ? null : v;
  if (typeof v === "number") return null;
  if (typeof v !== "string") return null;
  const s = v.trim();
  let m;
  if (/^\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2}(:\d{2})?)?/.test(s)) {
    const d = new Date(s); return Number.isNaN(+d) ? null : d;
  }
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))) {
    const d = new Date(+m[3], +m[2] - 1, +m[1]); return Number.isNaN(+d) ? null : d;
  }
  if ((m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/))) {
    const d = new Date(+m[3], +m[2] - 1, +m[1]); return Number.isNaN(+d) ? null : d;
  }
  return null;
}

function inferColumns(rows) {
  if (!rows.length) return [];
  const keys = Object.keys(rows[0]);
  const catThreshold = Math.max(15, Math.min(50, Math.round(rows.length * 0.05)));
  return keys.map((name) => {
    const raw = rows.map((r) => r[name]);
    const nonNull = raw.filter((v) => v !== null && v !== undefined && String(v).trim() !== "");
    const missing = raw.length - nonNull.length;
    const numberStyle = detectDecimalStyle(nonNull.slice(0, 300));
    const asNum = nonNull.map((v) => parseNumberWithStyle(v, numberStyle));
    const numericRatio = nonNull.length ? asNum.filter((v) => v !== null).length / nonNull.length : 0;
    const uniqueVals = _.uniq(nonNull.map((v) => (v instanceof Date ? +v : String(v))));
    const asDate = nonNull.map(tryParseDate);
    const dateRatio = nonNull.length ? asDate.filter((d) => d !== null).length / nonNull.length : 0;
    let type = "text";
    if (dateRatio > 0.85 && numericRatio < 0.85) type = "date";
    else if (numericRatio > 0.85) type = "numeric";
    else if (uniqueVals.length <= catThreshold) type = "categorical";
    return {
      name, type, missing, missingPct: raw.length ? (missing / raw.length) * 100 : 0,
      unique: uniqueVals.length, numberStyle,
    };
  });
}

function buildTypedRows(rows, columns) {
  return rows.map((r) => {
    const out = {};
    columns.forEach((c) => {
      const v = r[c.name];
      if (v === null || v === undefined || String(v).trim() === "") { out[c.name] = null; return; }
      if (c.type === "numeric") out[c.name] = parseNumberWithStyle(v, c.numberStyle);
      else if (c.type === "date") out[c.name] = tryParseDate(v) || new Date(v);
      else out[c.name] = String(v).trim();
    });
    return out;
  });
}

/* --- simple linear regression + time-series helpers (Aulas 37-41) --- */
function linearRegression(x, y) {
  const n = Math.min(x.length, y.length);
  const mx = mean(x), my = mean(y);
  const b = covariance(x, y) / variance(x, true);
  const a = my - b * mx;
  return { a, b };
}
function movingAverage(arr, window) {
  const out = new Array(arr.length).fill(null);
  const half = Math.floor(window / 2);
  for (let i = 0; i < arr.length; i++) {
    const lo = Math.max(0, i - half), hi = Math.min(arr.length - 1, i + half);
    const slice = arr.slice(lo, hi + 1);
    out[i] = mean(slice);
  }
  return out;
}
function mape(actual, forecast) {
  const pairs = actual.map((a, i) => [a, forecast[i]]).filter(([a, f]) => isNum(a) && isNum(f) && a !== 0);
  if (!pairs.length) return NaN;
  return (100 * _.sum(pairs.map(([a, f]) => Math.abs((a - f) / a)))) / pairs.length;
}
function rmse(actual, forecast) {
  const pairs = actual.map((a, i) => [a, forecast[i]]).filter(([a, f]) => isNum(a) && isNum(f));
  if (!pairs.length) return NaN;
  return Math.sqrt(_.sum(pairs.map(([a, f]) => (a - f) ** 2)) / pairs.length);
}

/* --- Estratégico: statistics-only predictive helpers (no Machine Learning) --- */
function exponentialSmoothing(y, alpha = 0.3) {
  if (!y.length) return [];
  const out = [y[0]];
  for (let i = 1; i < y.length; i++) out.push(alpha * y[i] + (1 - alpha) * out[i - 1]);
  return out;
}

// Small ordinary least-squares solver via Gaussian elimination (no external deps, no ML).
function solveLinearSystem(A, b) {
  const n = A.length;
  const M = A.map((row, i) => [...row, b[i]]);
  for (let col = 0; col < n; col++) {
    let pivot = col;
    for (let r = col + 1; r < n; r++) if (Math.abs(M[r][col]) > Math.abs(M[pivot][col])) pivot = r;
    if (Math.abs(M[pivot][col]) < 1e-10) return null;
    [M[col], M[pivot]] = [M[pivot], M[col]];
    for (let r = 0; r < n; r++) {
      if (r === col) continue;
      const factor = M[r][col] / M[col][col];
      for (let c = col; c <= n; c++) M[r][c] -= factor * M[col][c];
    }
  }
  return M.map((row, i) => row[n] / row[i]);
}

// Multiple linear regression Y = b0 + b1*X1 + ... + bk*Xk via normal equations.
// Returns coefficients, R², and per-coefficient significance (t-test), same rigor as the rest of the app.
function multipleRegression(predictorMatrix, y) {
  const n = y.length, k = predictorMatrix[0]?.length || 0;
  if (!n || !k) return null;
  const X = predictorMatrix.map((row) => [1, ...row]);
  const p = k + 1;
  const XtX = Array.from({ length: p }, (_r, i) => Array.from({ length: p }, (_c, j) => _.sum(X.map((row) => row[i] * row[j]))));
  const Xty = Array.from({ length: p }, (_r, i) => _.sum(X.map((row, idx) => row[i] * y[idx])));
  const coefs = solveLinearSystem(XtX, Xty);
  if (!coefs) return null;
  const fitted = X.map((row) => _.sum(row.map((v, i) => v * coefs[i])));
  const residuals = y.map((v, i) => v - fitted[i]);
  const yMean = mean(y);
  const ssTot = _.sum(y.map((v) => (v - yMean) ** 2));
  const ssRes = _.sum(residuals.map((r) => r ** 2));
  const r2 = ssTot !== 0 ? 1 - ssRes / ssTot : 0;
  const df = n - p;
  const sigma2 = df > 0 ? ssRes / df : NaN;
  let XtXInv = null;
  try {
    XtXInv = Array.from({ length: p }, (_r, i) => {
      const e = Array.from({ length: p }, (_c, j) => (i === j ? 1 : 0));
      return solveLinearSystem(XtX, e);
    });
  } catch { XtXInv = null; }
  const coefStats = coefs.map((c, i) => {
    const se = XtXInv && XtXInv[i] ? Math.sqrt(Math.max(0, sigma2 * XtXInv[i][i])) : NaN;
    const t = se ? c / se : NaN;
    const p_ = df > 0 && !Number.isNaN(t) ? tTwoTailP(t, df) : NaN;
    return { coef: c, se, t, p: p_ };
  });
  return { coefs, coefStats, r2, df, sigma2, fitted, residuals };
}

// Projects a linear trend to a target date (e.g. end of month/year) — descriptive extrapolation, not ML.
function projectTrendToDate(series, dateColName, valueColName, targetDate) {
  const withDate = series.filter((r) => r[dateColName] instanceof Date && !Number.isNaN(+r[dateColName]) && isNum(r[valueColName]))
    .sort((a, b) => a[dateColName] - b[dateColName]);
  if (withDate.length < 4) return null;
  const t0 = +withDate[0][dateColName];
  const dayMs = 86400000;
  const t = withDate.map((r) => (+r[dateColName] - t0) / dayMs);
  const y = withDate.map((r) => r[valueColName]);
  const { a, b } = linearRegression(t, y);
  const fitted = t.map((ti) => a + b * ti);
  const residuals = y.map((v, i) => v - fitted[i]);
  const residualSd = std(residuals, true) || 0;
  const targetT = (+targetDate - t0) / dayMs;
  const expected = a + b * targetT;
  return { slope: b, intercept: a, expected, optimistic: expected + residualSd, pessimistic: expected - residualSd, residualSd, n: withDate.length };
}

/* =========================================================================
   SMALL UI PRIMITIVES
   ========================================================================= */
const Card = ({ children, style, className = "" }) => (
  <div className={className} style={{ background: T.panel, border: `1px solid ${T.border}`, borderRadius: 12, ...style }}>
    {children}
  </div>
);

const SectionTitle = ({ eyebrow, title, right }) => (
  <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "space-between", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
    <div>
      {eyebrow && <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 11, letterSpacing: "0.12em", color: T.teal, textTransform: "uppercase", marginBottom: 4 }}>{eyebrow}</div>}
      <h2 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, fontWeight: 700, color: T.ink, margin: 0 }}>{title}</h2>
    </div>
    {right}
  </div>
);

const Pill = ({ children, tone = "neutral" }) => {
  const map = {
    neutral: { bg: "#F1F2F5", fg: T.sub },
    teal: { bg: T.tealSoft, fg: T.tealDark },
    amber: { bg: T.amberSoft, fg: T.amber },
    red: { bg: T.redSoft, fg: T.red },
    green: { bg: T.greenSoft, fg: T.green },
  };
  const c = map[tone];
  return (
    <span style={{ background: c.bg, color: c.fg, fontSize: 12, fontWeight: 600, padding: "3px 9px", borderRadius: 999, fontFamily: "'JetBrains Mono', monospace" }}>
      {children}
    </span>
  );
};

const Btn = ({ children, onClick, variant = "primary", disabled, style, icon: Icon }) => {
  const base = { border: "none", borderRadius: 9, padding: "9px 16px", fontSize: 13.5, fontWeight: 600, cursor: disabled ? "not-allowed" : "pointer", display: "inline-flex", alignItems: "center", gap: 7, fontFamily: "'Inter', sans-serif", opacity: disabled ? 0.5 : 1, transition: "all .15s" };
  const variants = {
    primary: { background: T.teal, color: "#fff" },
    ghost: { background: "transparent", color: T.ink, border: `1px solid ${T.border}` },
    danger: { background: T.redSoft, color: T.red },
    subtle: { background: "#F1F2F5", color: T.ink },
  };
  return (
    <button onClick={disabled ? undefined : onClick} style={{ ...base, ...variants[variant], ...style }}>
      {Icon && <Icon size={15} />}
      {children}
    </button>
  );
};

const Select = ({ value, onChange, options, placeholder }) => (
  <select value={value ?? ""} onChange={(e) => onChange(e.target.value)}
    style={{ padding: "8px 10px", borderRadius: 8, border: `1px solid ${T.border}`, fontSize: 13.5, fontFamily: "'Inter', sans-serif", color: T.ink, background: "#fff", minWidth: 160 }}>
    {placeholder && <option value="">{placeholder}</option>}
    {options.map((o) => <option key={o.value ?? o} value={o.value ?? o}>{o.label ?? o}</option>)}
  </select>
);

const StatTile = ({ label, value, tone = "neutral", sub }) => (
  <Card style={{ padding: "16px 18px", flex: "1 1 160px" }}>
    <div style={{ fontSize: 11.5, color: T.sub, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</div>
    <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 26, fontWeight: 700, color: tone === "neutral" ? T.ink : T[tone], marginTop: 4 }}>{value}</div>
    {sub && <div style={{ fontSize: 12, color: T.faint, marginTop: 2 }}>{sub}</div>}
  </Card>
);

function downloadBlob(content, filename, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}
function toCSV(rows) {
  if (!rows.length) return "";
  const keys = Object.keys(rows[0]);
  const lines = [keys.join(",")];
  rows.forEach((r) => lines.push(keys.map((k) => JSON.stringify(r[k] ?? "")).join(",")));
  return lines.join("\n");
}
function fmt(n, d = 3) {
  if (n === null || n === undefined || Number.isNaN(n)) return "—";
  if (typeof n !== "number") return String(n);
  if (Math.abs(n) >= 1000) return n.toLocaleString("pt-BR", { maximumFractionDigits: 1 });
  return n.toLocaleString("pt-BR", { maximumFractionDigits: d });
}
function fmtP(p) {
  if (p === null || p === undefined || Number.isNaN(p)) return "—";
  return p < 0.0001 ? "< 0.0001" : p.toFixed(4);
}

/* =========================================================================
   UPLOAD VIEW
   ========================================================================= */
function UploadView({ onLoaded }) {
  const [dragOver, setDragOver] = useState(false);
  const [error, setError] = useState(null);
  const [busy, setBusy] = useState(false);
  const inputRef = useRef(null);

  const handleFile = useCallback((file) => {
    if (!file) return;
    setBusy(true); setError(null);
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "csv" || ext === "tsv" || ext === "txt") {
      Papa.parse(file, {
        header: true, skipEmptyLines: true, dynamicTyping: false,
        transformHeader: (h) => h.replace(/^\uFEFF/, "").trim(),
        complete: (res) => {
          if (!res.data.length) { setError("O arquivo parece vazio."); setBusy(false); return; }
          const { rows } = normalizeParsedRows(res.data);
          onLoaded(rows, file.name, file);
          setBusy(false);
        },
        error: (err) => { setError("Falha ao ler CSV: " + err.message); setBusy(false); },
      });
    } else if (ext === "xlsx" || ext === "xls") {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const json = XLSX.utils.sheet_to_json(sheet, { defval: null });
          if (!json.length) { setError("A planilha parece vazia."); setBusy(false); return; }
          const { rows } = normalizeParsedRows(json);
          onLoaded(rows, file.name, file);
        } catch (err) { setError("Falha ao ler Excel: " + err.message); }
        setBusy(false);
      };
      reader.readAsArrayBuffer(file);
    } else if (ext === "json") {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const parsed = JSON.parse(e.target.result);
          const arr = Array.isArray(parsed) ? parsed : parsed.data || Object.values(parsed);
          if (!Array.isArray(arr) || !arr.length) { setError("JSON precisa ser uma lista de registros."); setBusy(false); return; }
          const { rows } = normalizeParsedRows(arr);
          onLoaded(rows, file.name, file);
        } catch (err) { setError("JSON inválido: " + err.message); }
        setBusy(false);
      };
      reader.readAsText(file);
    } else {
      setError("Formato não suportado. Use CSV, TSV, XLSX, XLS ou JSON.");
      setBusy(false);
    }
  }, [onLoaded]);

  return (
    <div style={{ maxWidth: 720, margin: "40px auto" }}>
      <div style={{ textAlign: "center", marginBottom: 28 }}>
        <div style={{ display: "inline-flex", alignItems: "center", gap: 8, background: T.tealSoft, color: T.tealDark, padding: "5px 12px", borderRadius: 999, fontSize: 12, fontWeight: 700, fontFamily: "'JetBrains Mono', monospace", letterSpacing: "0.04em" }}>
          <Sparkles size={13} /> ANÁLISE DE DADOS
        </div>
        <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 34, fontWeight: 700, color: T.ink, margin: "14px 0 8px" }}>
          Carregue sua base para começar
        </h1>
        <p style={{ color: T.sub, fontSize: 15, maxWidth: 520, margin: "0 auto", lineHeight: 1.5 }}>
          CSV, TSV, Excel (.xlsx/.xls) ou JSON. Tudo é processado localmente no seu navegador — nenhum dado sai do seu computador.
        </p>
      </div>

      <div
        onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
        onDragLeave={() => setDragOver(false)}
        onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]); }}
        onClick={() => inputRef.current?.click()}
        style={{
          border: `2px dashed ${dragOver ? T.teal : T.border}`, borderRadius: 16, padding: "48px 24px",
          textAlign: "center", cursor: "pointer", background: dragOver ? T.tealSoft : T.panel, transition: "all .15s",
        }}>
        <input ref={inputRef} type="file" accept=".csv,.tsv,.txt,.xlsx,.xls,.json" style={{ display: "none" }}
          onChange={(e) => handleFile(e.target.files[0])} />
        <Upload size={34} color={T.teal} style={{ marginBottom: 12 }} />
        <div style={{ fontWeight: 700, color: T.ink, fontSize: 16 }}>
          {busy ? "Processando arquivo…" : "Arraste um arquivo aqui ou clique para selecionar"}
        </div>
        <div style={{ color: T.faint, fontSize: 13, marginTop: 6 }}>.csv · .tsv · .xlsx · .xls · .json</div>
      </div>

      {error && (
        <div style={{ marginTop: 16, background: T.redSoft, color: T.red, padding: "12px 14px", borderRadius: 10, fontSize: 13.5, display: "flex", gap: 8, alignItems: "flex-start" }}>
          <AlertTriangle size={16} style={{ flexShrink: 0, marginTop: 1 }} /> {error}
        </div>
      )}

      <div style={{ marginTop: 28, display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12 }}>
        {[
          { icon: Gauge, t: "Diagnóstico automático", d: "Tipos de variável, ausentes, duplicados e outliers em segundos." },
          { icon: BarChart3, t: "EDA instantânea", d: "Histogramas, boxplots, dispersão e correlação sem configurar nada." },
          { icon: FlaskConical, t: "Relatórios automáticos", d: "ydata-profiling e Sweetviz gerados a partir da mesma base." },
        ].map((f, i) => (
          <Card key={i} style={{ padding: 16 }}>
            <f.icon size={18} color={T.teal} />
            <div style={{ fontWeight: 700, fontSize: 13.5, marginTop: 8, color: T.ink }}>{f.t}</div>
            <div style={{ fontSize: 12.5, color: T.sub, marginTop: 4, lineHeight: 1.4 }}>{f.d}</div>
          </Card>
        ))}
      </div>
    </div>
  );
}

/* =========================================================================
   DASHBOARD TAB
   ========================================================================= */
function QualityRing({ score }) {
  const R = 54, C = 2 * Math.PI * R;
  const color = score >= 80 ? T.green : score >= 55 ? T.amber : T.red;
  return (
    <svg width={140} height={140} viewBox="0 0 140 140">
      <circle cx="70" cy="70" r={R} fill="none" stroke="#EEF0F3" strokeWidth="14" />
      <circle cx="70" cy="70" r={R} fill="none" stroke={color} strokeWidth="14" strokeLinecap="round"
        strokeDasharray={`${(score / 100) * C} ${C}`} transform="rotate(-90 70 70)"
        style={{ transition: "stroke-dasharray 0.8s ease" }} />
      <text x="70" y="66" textAnchor="middle" fontSize="30" fontWeight="700" fontFamily="'Space Grotesk', sans-serif" fill={T.ink}>{Math.round(score)}</text>
      <text x="70" y="86" textAnchor="middle" fontSize="11" fontFamily="'JetBrains Mono', monospace" fill={T.sub}>de 100</text>
    </svg>
  );
}

function computeDashboard(rows, columns) {
  const rowCount = rows.length, colCount = columns.length;
  const missingTotal = _.sum(columns.map((c) => c.missing));
  const missingPct = rowCount * colCount ? (missingTotal / (rowCount * colCount)) * 100 : 0;
  const seen = new Set(); let duplicates = 0;
  rows.forEach((r) => { const k = JSON.stringify(r); if (seen.has(k)) duplicates++; else seen.add(k); });
  const memBytes = JSON.stringify(rows).length;
  const numericCols = columns.filter((c) => c.type === "numeric");
  const catCols = columns.filter((c) => c.type === "categorical");
  let corrs = [];
  let corrPairs = [];
  for (let i = 0; i < numericCols.length; i++) {
    for (let j = i + 1; j < numericCols.length; j++) {
      const x = rows.map((r) => r[numericCols[i].name]).filter(isNum);
      const y = rows.map((r) => r[numericCols[j].name]).filter(isNum);
      const n = Math.min(x.length, y.length);
      const r = pearson(x.slice(0, n), y.slice(0, n));
      if (!Number.isNaN(r)) { corrs.push(Math.abs(r)); corrPairs.push({ a: numericCols[i].name, b: numericCols[j].name, r }); }
    }
  }
  corrPairs.sort((p1, p2) => Math.abs(p2.r) - Math.abs(p1.r));
  const topCorrelations = corrPairs.slice(0, 5);
  const avgCorr = corrs.length ? mean(corrs) : NaN;
  let outlierTotal = 0;
  const outliersByCol = [];
  numericCols.forEach((c) => {
    const vals = rows.map((r) => r[c.name]).filter(isNum);
    const n = outliersIQR(vals).length;
    outlierTotal += n;
    if (n > 0) outliersByCol.push({ name: c.name, count: n, pct: vals.length ? (n / vals.length) * 100 : 0 });
  });
  outliersByCol.sort((a, b) => b.count - a.count);
  const dtypeDist = _.countBy(columns, "type");
  const outlierPct = rowCount * numericCols.length ? (outlierTotal / (rowCount * Math.max(numericCols.length, 1))) * 100 : 0;
  const duplicatePct = rowCount ? (duplicates / rowCount) * 100 : 0;
  const quality = Math.max(0, Math.min(100, 100 - missingPct * 0.6 - duplicatePct * 0.8 - outlierPct * 0.4));

  // Low-variance / constant numeric columns (little analytical value as-is)
  const lowVarianceCols = numericCols.map((c) => {
    const vals = rows.map((r) => r[c.name]).filter(isNum);
    return { name: c.name, cv: cv(vals), constant: _.uniq(vals).length <= 1 };
  }).filter((c) => c.constant || (!Number.isNaN(c.cv) && Math.abs(c.cv) < 1));

  // High-cardinality text columns (likely IDs / free text — low direct BI value)
  const highCardinalityText = columns.filter((c) => c.type === "text" && rowCount && c.unique / rowCount > 0.9)
    .map((c) => c.name);

  // Pareto (80/20) on the categorical column with the most reasonable cardinality (mirrors Aula 24)
  let pareto = null;
  const paretoCandidate = catCols.filter((c) => c.unique >= 3).sort((a, b) => b.unique - a.unique)[0];
  if (paretoCandidate) {
    const counts = _.countBy(rows.map((r) => r[paretoCandidate.name]).filter((v) => v !== null));
    const sorted = Object.entries(counts).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
    const total = _.sum(sorted.map((s) => s.value));
    let cum = 0, idx80 = sorted.length;
    for (let i = 0; i < sorted.length; i++) { cum += sorted[i].value; if (cum / total >= 0.8) { idx80 = i + 1; break; } }
    pareto = { column: paretoCandidate.name, totalCategories: sorted.length, categoriesFor80: idx80, top: sorted.slice(0, 5), total };
  }

  return {
    rowCount, colCount, missingTotal, missingPct, duplicates, duplicatePct, memBytes, avgCorr,
    outlierTotal, outlierPct, dtypeDist, quality, topCorrelations, outliersByCol,
    lowVarianceCols, highCardinalityText, pareto, allCorrPairs: corrPairs,
  };
}

function buildExecutiveInsights(d, columns) {
  const insights = [];
  if (d.missingPct > 5) insights.push({ tone: "amber", text: `${fmt(d.missingPct, 1)}% das células estão ausentes — considere tratar (imputação ou exclusão) antes de análises mais profundas.` });
  else insights.push({ tone: "green", text: "Baixo percentual de valores ausentes — base pronta para análise sem tratamento prévio." });
  if (d.duplicates > 0) insights.push({ tone: "amber", text: `${d.duplicates} linha(s) duplicada(s) (${fmt(d.duplicatePct, 1)}%) — verifique se são erros de coleta/carga.` });
  if (d.outlierTotal > 0) insights.push({ tone: "amber", text: `${d.outlierTotal} outlier(s) identificados pelo método IQR, concentrados principalmente em: ${d.outliersByCol.slice(0, 3).map((o) => o.name).join(", ")}.` });
  if (d.topCorrelations.length) {
    const top = d.topCorrelations[0];
    insights.push({ tone: "teal", text: `Correlação mais forte da base: ${top.a} × ${top.b} (r = ${fmt(top.r, 2)}) — ${Math.abs(top.r) > 0.7 ? "relação forte" : Math.abs(top.r) > 0.4 ? "relação moderada" : "relação fraca"}, ${top.r >= 0 ? "positiva" : "negativa"}.` });
  }
  if (d.pareto) insights.push({ tone: "teal", text: `Princípio de Pareto em "${d.pareto.column}": ${d.pareto.categoriesFor80} de ${d.pareto.totalCategories} categorias concentram 80% dos registros.` });
  if (d.lowVarianceCols.length) insights.push({ tone: "neutral", text: `Colunas com variação quase nula: ${d.lowVarianceCols.slice(0, 4).map((c) => c.name).join(", ")} — pouco poder discriminante para modelos ou segmentações.` });
  if (d.highCardinalityText.length) insights.push({ tone: "neutral", text: `Colunas de texto com altíssima cardinalidade (prováveis IDs/textos livres): ${d.highCardinalityText.slice(0, 4).join(", ")}.` });
  return insights;
}

/* --- Dashboard: sparkline + adaptive KPI series (reuses existing stat fns only) --- */
function Sparkline({ data, width = 110, height = 32, color = T.teal }) {
  if (!data || data.length < 2) return <svg width={width} height={height} />;
  const min = Math.min(...data), max = Math.max(...data);
  const range = max - min || 1;
  const stepX = width / (data.length - 1);
  const points = data.map((v, i) => `${i * stepX},${height - ((v - min) / range) * (height - 4) - 2}`).join(" ");
  const lastY = height - ((data[data.length - 1] - min) / range) * (height - 4) - 2;
  return (
    <svg width={width} height={height} viewBox={`0 0 ${width} ${height}`} preserveAspectRatio="none">
      <polyline points={points} fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
      <circle cx={width} cy={lastY} r="2.5" fill={color} />
    </svg>
  );
}

function buildKpiSeries(filteredRows, numericCols, dateColName) {
  if (!numericCols.length || !filteredRows.length) return [];
  const candidates = numericCols.filter((c) => c.unique < Math.max(5, filteredRows.length * 0.98));
  const pool = candidates.length ? candidates : numericCols;
  const scored = pool.map((c) => {
    const vals = filteredRows.map((r) => r[c.name]).filter(isNum);
    return { name: c.name, vals, cvAbs: Number.isNaN(cv(vals)) ? 0 : Math.abs(cv(vals)) };
  }).filter((s) => s.vals.length > 0);
  scored.sort((a, b) => b.cvAbs - a.cvAbs);
  return scored.slice(0, 4).map(({ name, vals }) => {
    let sparkData = null, growthPct = null, hasTrend = false;
    if (dateColName) {
      const withDate = filteredRows.filter((r) => r[dateColName] instanceof Date && !Number.isNaN(+r[dateColName]) && isNum(r[name]))
        .sort((a, b) => a[dateColName] - b[dateColName]);
      if (withDate.length >= 4) {
        hasTrend = true;
        const buckets = Math.min(14, withDate.length);
        const chunkSize = Math.max(1, Math.ceil(withDate.length / buckets));
        sparkData = [];
        for (let i = 0; i < withDate.length; i += chunkSize) sparkData.push(mean(withDate.slice(i, i + chunkSize).map((r) => r[name])));
        const half = Math.floor(withDate.length / 2);
        const m1 = mean(withDate.slice(0, half).map((r) => r[name]));
        const m2 = mean(withDate.slice(half).map((r) => r[name]));
        growthPct = m1 !== 0 ? ((m2 - m1) / Math.abs(m1)) * 100 : null;
      }
    }
    if (!sparkData) {
      const sorted = [...vals].sort((a, b) => a - b);
      const buckets = Math.min(14, sorted.length);
      const chunkSize = Math.max(1, Math.ceil(sorted.length / buckets));
      sparkData = [];
      for (let i = 0; i < sorted.length; i += chunkSize) sparkData.push(mean(sorted.slice(i, i + chunkSize)));
    }
    return { name, total: _.sum(vals), avg: mean(vals), n: vals.length, growthPct, hasTrend, sparkData };
  });
}

function buildRichInsights(d) {
  const items = [];
  if (d.topCorrelations.length) {
    const top = d.topCorrelations[0];
    const strength = Math.abs(top.r) > 0.7 ? "alto" : Math.abs(top.r) > 0.4 ? "médio" : "baixo";
    items.push({
      icon: Activity, tone: Math.abs(top.r) > 0.4 ? "teal" : "neutral",
      title: `Correlação ${top.r >= 0 ? "positiva" : "negativa"} entre ${top.a} e ${top.b}`,
      simple: `As duas variáveis tendem a se mover ${top.r >= 0 ? "na mesma direção" : "em direções opostas"}.`,
      technical: `Coeficiente de Pearson r = ${fmt(top.r, 2)}.`,
      impact: strength, confidence: Math.round(Math.abs(top.r) * 100),
    });
  }
  if (d.pareto) {
    items.push({
      icon: Target, tone: "teal",
      title: `Concentração em "${d.pareto.column}" (Pareto 80/20)`,
      simple: "Poucas categorias respondem pela maior parte dos registros.",
      technical: `${d.pareto.categoriesFor80} de ${d.pareto.totalCategories} categorias concentram 80% do total (n=${d.pareto.total}).`,
      impact: "médio", confidence: 85,
    });
  }
  if (d.outlierTotal > 0) {
    items.push({
      icon: AlertTriangle, tone: "amber",
      title: "Outliers identificados",
      simple: `${d.outlierTotal} valor(es) fora do padrão esperado, concentrados em ${d.outliersByCol.slice(0, 2).map((o) => o.name).join(", ")}.`,
      technical: `Método IQR (1.5×IQR). Maior concentração: ${d.outliersByCol[0]?.name} (${d.outliersByCol[0]?.count} outliers, ${fmt(d.outliersByCol[0]?.pct, 1)}%).`,
      impact: d.outlierTotal > d.rowCount * 0.05 ? "alto" : "médio", confidence: 90,
    });
  }
  if (d.missingPct > 5) {
    items.push({
      icon: Info, tone: "amber", title: "Volume relevante de dados ausentes",
      simple: `${fmt(d.missingPct, 1)}% das células da base estão vazias.`,
      technical: `${d.missingTotal} células ausentes no total.`,
      impact: d.missingPct > 20 ? "alto" : "médio", confidence: 95,
    });
  } else {
    items.push({
      icon: CheckCircle2, tone: "green", title: "Baixa taxa de valores ausentes",
      simple: "A base está bem preenchida, sem necessidade de tratamento prévio significativo.",
      technical: `Apenas ${fmt(d.missingPct, 1)}% de ausência nas células.`,
      impact: "baixo", confidence: 95,
    });
  }
  if (d.duplicates > 0) {
    items.push({
      icon: Layers, tone: "amber", title: "Registros duplicados detectados",
      simple: `${d.duplicates} linha(s) repetida(s) — pode indicar erro de carga ou integração.`,
      technical: `${fmt(d.duplicatePct, 1)}% das linhas são duplicatas exatas.`,
      impact: d.duplicatePct > 5 ? "alto" : "baixo", confidence: 99,
    });
  }
  return items;
}

function buildAlerts(d, columns) {
  const alerts = [];
  columns.forEach((c) => {
    if (c.missingPct > 50) alerts.push({ level: "critico", text: `Coluna "${c.name}" com ${fmt(c.missingPct, 1)}% de valores ausentes — considere excluir ou investigar a fonte.` });
    else if (c.missingPct > 20) alerts.push({ level: "atencao", text: `Coluna "${c.name}" com ${fmt(c.missingPct, 1)}% de ausentes.` });
  });
  if (d.duplicatePct > 10) alerts.push({ level: "critico", text: `${fmt(d.duplicatePct, 1)}% das linhas são duplicadas — risco de dupla contagem em agregações.` });
  else if (d.duplicates > 0) alerts.push({ level: "atencao", text: `${d.duplicates} linha(s) duplicada(s) identificada(s).` });
  d.outliersByCol.slice(0, 3).forEach((o) => {
    if (o.pct > 10) alerts.push({ level: "atencao", text: `"${o.name}" tem ${fmt(o.pct, 1)}% de outliers — revisar coleta ou considerar tratamento.` });
  });
  d.lowVarianceCols.forEach((c) => {
    if (c.constant) alerts.push({ level: "info", text: `Coluna "${c.name}" é constante (valor único) — sem poder analítico.` });
  });
  if (!alerts.length) alerts.push({ level: "ok", text: "Nenhum alerta crítico identificado — base em boas condições gerais." });
  return alerts;
}

/* --- Dashboard: reusable BI panels (Top/Bottom ranking, benchmark, waterfall, radar, bubble, stats) --- */
function TopBottomCard({ C, rows, catCols, catCol, setCatCol }) {
  const sorted = useMemo(() => {
    if (!catCol) return [];
    const counts = _.countBy(rows.map((r) => r[catCol]).filter((v) => v !== null && v !== undefined));
    return Object.entries(counts).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [rows, catCol]);
  const top10 = sorted.slice(0, 10);
  const bottom10 = sorted.slice(-10).reverse();
  return (
    <Card style={{ padding: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontWeight: 700, fontSize: 14, color: C.ink }}>Ranking — Top 10 vs. Bottom 10</div>
        <Select value={catCol} onChange={setCatCol} options={catCols.map((c) => c.name)} />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 16 }}>
        <div>
          <div style={{ fontSize: 11.5, fontWeight: 700, color: C.green, marginBottom: 6, textTransform: "uppercase" }}>Top 10</div>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={top10} layout="vertical" margin={{ left: 8 }}>
              <XAxis type="number" tick={{ fontSize: 10 }} />
              <YAxis type="category" dataKey="name" width={100} tick={{ fontSize: 9.5 }} />
              <Tooltip /><Bar dataKey="value" fill={C.teal} radius={[0, 3, 3, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div>
          <div style={{ fontSize: 11.5, fontWeight: 700, color: C.amber, marginBottom: 6, textTransform: "uppercase" }}>Bottom 10</div>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={bottom10} layout="vertical" margin={{ left: 8 }}>
              <XAxis type="number" tick={{ fontSize: 10 }} />
              <YAxis type="category" dataKey="name" width={100} tick={{ fontSize: 9.5 }} />
              <Tooltip /><Bar dataKey="value" fill={C.amber} radius={[0, 3, 3, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </Card>
  );
}

function BenchmarkCard({ C, rows, catCols, numericCols, catCol, setCatCol, numCol, setNumCol }) {
  const data = useMemo(() => {
    if (!catCol || !numCol) return { bars: [], overall: 0 };
    const groups = _.groupBy(rows.filter((r) => r[catCol] !== null && r[catCol] !== undefined && isNum(r[numCol])), catCol);
    const bars = Object.entries(groups).map(([name, items]) => ({ name, avg: mean(items.map((it) => it[numCol])) }))
      .sort((a, b) => b.avg - a.avg).slice(0, 12);
    const overall = mean(rows.map((r) => r[numCol]).filter(isNum));
    return { bars, overall };
  }, [rows, catCol, numCol]);
  return (
    <Card style={{ padding: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontWeight: 700, fontSize: 14, color: C.ink }}>Benchmark por categoria</div>
        <div style={{ display: "flex", gap: 8 }}>
          <Select value={catCol} onChange={setCatCol} options={catCols.map((c) => c.name)} />
          <Select value={numCol} onChange={setNumCol} options={numericCols.map((c) => c.name)} />
        </div>
      </div>
      <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 10 }}>Média por categoria comparada à média geral (linha de referência).</div>
      <ResponsiveContainer width="100%" height={240}>
        <BarChart data={data.bars}>
          <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
          <XAxis dataKey="name" tick={{ fontSize: 9 }} interval={0} angle={-30} textAnchor="end" height={60} />
          <YAxis tick={{ fontSize: 10 }} />
          <Tooltip />
          <ReferenceLine y={data.overall} stroke={C.red} strokeDasharray="4 3" label={{ value: "Média geral", fontSize: 10, fill: C.red, position: "insideTopRight" }} />
          <Bar dataKey="avg" radius={[3, 3, 0, 0]}>
            {data.bars.map((b, i) => <Cell key={i} fill={b.avg >= data.overall ? C.teal : C.amber} />)}
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </Card>
  );
}

function WaterfallCard({ C, rows, catCols, numericCols, catCol, setCatCol, numCol, setNumCol }) {
  const data = useMemo(() => {
    if (!catCol || !numCol) return [];
    const groups = _.groupBy(rows.filter((r) => r[catCol] !== null && r[catCol] !== undefined && isNum(r[numCol])), catCol);
    const sums = Object.entries(groups).map(([name, items]) => ({ name, value: _.sum(items.map((it) => it[numCol])) }))
      .sort((a, b) => b.value - a.value).slice(0, 8);
    let cum = 0;
    return sums.map((s) => { const base = cum; cum += s.value; return { name: s.name, base, value: s.value }; });
  }, [rows, catCol, numCol]);
  return (
    <Card style={{ padding: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontWeight: 700, fontSize: 14, color: C.ink }}>Waterfall — contribuição acumulada</div>
        <div style={{ display: "flex", gap: 8 }}>
          <Select value={catCol} onChange={setCatCol} options={catCols.map((c) => c.name)} />
          <Select value={numCol} onChange={setNumCol} options={numericCols.map((c) => c.name)} />
        </div>
      </div>
      <ResponsiveContainer width="100%" height={240}>
        <BarChart data={data}>
          <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
          <XAxis dataKey="name" tick={{ fontSize: 9 }} interval={0} angle={-30} textAnchor="end" height={60} />
          <YAxis tick={{ fontSize: 10 }} />
          <Tooltip />
          <Bar dataKey="base" stackId="a" fill="transparent" />
          <Bar dataKey="value" stackId="a" fill={C.teal} radius={[3, 3, 0, 0]} />
        </BarChart>
      </ResponsiveContainer>
    </Card>
  );
}

function RadarProfileCard({ C, rows, numericCols }) {
  const data = useMemo(() => numericCols.slice(0, 8).map((c) => {
    const vals = rows.map((r) => r[c.name]).filter(isNum);
    if (!vals.length) return { metric: c.name, value: 0 };
    const mn = Math.min(...vals), mx = Math.max(...vals);
    const avg = mean(vals);
    return { metric: c.name, value: Math.round(mx > mn ? ((avg - mn) / (mx - mn)) * 100 : 50) };
  }), [rows, numericCols]);
  if (data.length < 3) return null;
  return (
    <Card style={{ padding: 20 }}>
      <div style={{ fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 4 }}>Perfil das variáveis (Radar)</div>
      <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 6 }}>Média de cada variável normalizada (0–100) entre seu próprio mínimo e máximo.</div>
      <ResponsiveContainer width="100%" height={260}>
        <RadarChart data={data} outerRadius={90}>
          <PolarGrid stroke={C.border} />
          <PolarAngleAxis dataKey="metric" tick={{ fontSize: 10, fill: C.sub }} />
          <PolarRadiusAxis domain={[0, 100]} tick={{ fontSize: 9 }} />
          <RadarSeries dataKey="value" stroke={C.teal} fill={C.teal} fillOpacity={0.35} />
          <Tooltip />
        </RadarChart>
      </ResponsiveContainer>
    </Card>
  );
}

function BubbleChartCard({ C, rows, numericCols }) {
  const [xCol, setXCol] = useState(numericCols[0]?.name || "");
  const [yCol, setYCol] = useState(numericCols[1]?.name || numericCols[0]?.name || "");
  const [sizeCol, setSizeCol] = useState(numericCols[2]?.name || numericCols[0]?.name || "");
  const data = useMemo(() => rows.filter((r) => isNum(r[xCol]) && isNum(r[yCol]) && isNum(r[sizeCol]))
    .map((r) => ({ x: r[xCol], y: r[yCol], z: r[sizeCol] })), [rows, xCol, yCol, sizeCol]);
  if (numericCols.length < 2) return null;
  return (
    <Card style={{ padding: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontWeight: 700, fontSize: 14, color: C.ink }}>Bubble Chart</div>
        <div style={{ display: "flex", gap: 8 }}>
          <Select value={xCol} onChange={setXCol} options={numericCols.map((c) => c.name)} />
          <Select value={yCol} onChange={setYCol} options={numericCols.map((c) => c.name)} />
          <Select value={sizeCol} onChange={setSizeCol} options={numericCols.map((c) => c.name)} />
        </div>
      </div>
      <ResponsiveContainer width="100%" height={260}>
        <ScatterChart>
          <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
          <XAxis dataKey="x" name={xCol} tick={{ fontSize: 10 }} />
          <YAxis dataKey="y" name={yCol} tick={{ fontSize: 10 }} />
          <ZAxis dataKey="z" range={[40, 400]} name={sizeCol} />
          <Tooltip cursor={{ strokeDasharray: "3 3" }} />
          <Scatter data={data} fill={C.teal} fillOpacity={0.5} />
        </ScatterChart>
      </ResponsiveContainer>
    </Card>
  );
}

function StatSummaryCard({ C, rows, numericCols }) {
  const [col, setCol] = useState(numericCols[0]?.name || "");
  const vals = useMemo(() => rows.map((r) => r[col]).filter(isNum), [rows, col]);
  if (!numericCols.length) return null;
  const n = vals.length;
  const m = mean(vals), md = median(vals), sd = std(vals);
  const se = n ? sd / Math.sqrt(n) : NaN;
  const modeVals = mode(vals);
  return (
    <Card style={{ padding: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontWeight: 700, fontSize: 14, color: C.ink }}>Análise estatística resumida</div>
        <Select value={col} onChange={setCol} options={numericCols.map((c) => c.name)} />
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
        <KpiCard C={C} label="N" value={n} />
        <KpiCard C={C} label="Média" value={fmt(m, 2)} />
        <KpiCard C={C} label="Mediana" value={fmt(md, 2)} />
        <KpiCard C={C} label="Moda" value={modeVals.length ? modeVals.map((v) => fmt(v, 1)).join(", ") : "—"} />
        <KpiCard C={C} label="Desvio padrão" value={fmt(sd, 2)} />
        <KpiCard C={C} label="Assimetria" value={fmt(skewness(vals), 2)} />
        <KpiCard C={C} label="Curtose" value={fmt(kurtosisExcess(vals), 2)} />
        <KpiCard C={C} label="IC 95% da média" value={Number.isNaN(se) ? "—" : `${fmt(m - 1.96 * se, 1)} — ${fmt(m + 1.96 * se, 1)}`} />
      </div>
    </Card>
  );
}

function FilterBar({ C, rows, dateCols, catCols, dateStart, setDateStart, dateEnd, setDateEnd, filterCol1, setFilterCol1, filterVal1, setFilterVal1, filterCol2, setFilterCol2, filterVal2, setFilterVal2, onReset, activeCount }) {
  const options1 = useMemo(() => filterCol1 ? _.uniq(rows.map((r) => r[filterCol1]).filter((v) => v !== null && v !== undefined)).map(String).sort() : [], [rows, filterCol1]);
  const options2 = useMemo(() => filterCol2 ? _.uniq(rows.map((r) => r[filterCol2]).filter((v) => v !== null && v !== undefined)).map(String).sort() : [], [rows, filterCol2]);
  return (
    <Card style={{ padding: "14px 18px", marginBottom: 16, display: "flex", flexWrap: "wrap", gap: 12, alignItems: "center" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 700, color: T.sub, textTransform: "uppercase" }}>
        <Filter size={13} /> Filtros
      </div>
      {dateCols.length > 0 && (
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <Calendar size={13} color={T.faint} />
          <input type="date" value={dateStart} onChange={(e) => setDateStart(e.target.value)} style={{ padding: "6px 8px", borderRadius: 7, border: `1px solid ${T.border}`, fontSize: 12.5 }} />
          <span style={{ color: T.faint, fontSize: 12 }}>até</span>
          <input type="date" value={dateEnd} onChange={(e) => setDateEnd(e.target.value)} style={{ padding: "6px 8px", borderRadius: 7, border: `1px solid ${T.border}`, fontSize: 12.5 }} />
        </div>
      )}
      <div style={{ display: "flex", gap: 6 }}>
        <Select value={filterCol1} onChange={(v) => { setFilterCol1(v); setFilterVal1(""); }} options={catCols.map((c) => c.name)} placeholder="Coluna 1" />
        <Select value={filterVal1} onChange={setFilterVal1} options={options1} placeholder="Todos" />
      </div>
      <div style={{ display: "flex", gap: 6 }}>
        <Select value={filterCol2} onChange={(v) => { setFilterCol2(v); setFilterVal2(""); }} options={catCols.map((c) => c.name)} placeholder="Coluna 2" />
        <Select value={filterVal2} onChange={setFilterVal2} options={options2} placeholder="Todos" />
      </div>
      {activeCount > 0 && <Btn variant="subtle" onClick={onReset}>Limpar filtros ({activeCount})</Btn>}
    </Card>
  );
}
/* --- Reusable fullscreen chart wrapper (Maximize2 pattern applied to key charts) --- */
function ChartFrame({ C, title, subtitle, children, height = 260 }) {
  const [full, setFull] = useState(false);
  return (
    <>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 }}>
        <div>
          <div style={{ fontWeight: 700, fontSize: 14, color: C.ink }}>{title}</div>
          {subtitle && <div style={{ fontSize: 11.5, color: C.faint, marginTop: 2 }}>{subtitle}</div>}
        </div>
        <button onClick={() => setFull(true)} title="Tela cheia" style={{ background: "transparent", border: `1px solid ${C.border}`, borderRadius: 7, width: 28, height: 28, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", flexShrink: 0 }}>
          <Maximize2 size={13} color={C.faint} />
        </button>
      </div>
      <div style={{ marginTop: 8 }}>{children(height)}</div>
      {full && (
        <div onClick={() => setFull(false)} style={{ position: "fixed", inset: 0, background: "rgba(10,12,16,0.55)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: C.panel, borderRadius: 16, padding: 24, width: "min(1100px, 100%)", maxHeight: "88vh", overflow: "auto", boxShadow: "0 20px 60px rgba(0,0,0,0.35)" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
              <div style={{ fontWeight: 700, fontSize: 16, color: C.ink }}>{title}</div>
              <button onClick={() => setFull(false)} style={{ background: "transparent", border: `1px solid ${C.border}`, borderRadius: 7, width: 30, height: 30, cursor: "pointer" }}>✕</button>
            </div>
            {children(520)}
          </div>
        </div>
      )}
    </>
  );
}

function ImpactEffortMatrix({ C, items }) {
  if (!items.length) return null;
  return (
    <div>
      <ResponsiveContainer width="100%" height={260}>
        <ScatterChart margin={{ top: 10, right: 20, bottom: 10, left: 0 }}>
          <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
          <XAxis type="number" dataKey="effort" name="Esforço" domain={[0, 100]} tick={{ fontSize: 10 }} label={{ value: "Esforço →", position: "insideBottomRight", fontSize: 10, fill: C.faint }} />
          <YAxis type="number" dataKey="impact" name="Impacto" domain={[0, 100]} tick={{ fontSize: 10 }} label={{ value: "Impacto →", angle: -90, position: "insideLeft", fontSize: 10, fill: C.faint }} />
          <ReferenceLine x={50} stroke={C.border} /><ReferenceLine y={50} stroke={C.border} />
          <Tooltip cursor={{ strokeDasharray: "3 3" }} content={({ active, payload }) => active && payload?.length ? (
            <div style={{ background: C.panel, border: `1px solid ${C.border}`, borderRadius: 8, padding: 8, fontSize: 11.5, color: C.ink }}>{payload[0].payload.label}</div>
          ) : null} />
          <Scatter data={items} fill={C.teal}>
            {items.map((it, i) => <Cell key={i} fill={it.impact >= 50 && it.effort < 50 ? C.green : it.impact >= 50 ? C.teal : C.amber} />)}
          </Scatter>
        </ScatterChart>
      </ResponsiveContainer>
      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 10, color: C.faint, marginTop: 4, fontFamily: "'JetBrains Mono', monospace" }}>
        <span>alto impacto · baixo esforço = prioridade</span><span>baixo impacto · alto esforço = evitar</span>
      </div>
    </div>
  );
}

function DashboardTab({ rows, columns, fileName }) {
  const [theme, setTheme] = useState("light");
  const C = theme === "dark" ? REPORT_DARK : T;
  const [subTab, setSubTab] = useState("executivo");

  const dateCols = columns.filter((c) => c.type === "date");
  const catCols = columns.filter((c) => c.type === "categorical");
  const numericCols = columns.filter((c) => c.type === "numeric");
  const dateColName = dateCols[0]?.name || "";

  const [dateStart, setDateStart] = useState("");
  const [dateEnd, setDateEnd] = useState("");
  const [filterCol1, setFilterCol1] = useState("");
  const [filterVal1, setFilterVal1] = useState("");
  const [filterCol2, setFilterCol2] = useState("");
  const [filterVal2, setFilterVal2] = useState("");
  const [refreshedAt, setRefreshedAt] = useState(() => new Date());
  const [shareState, setShareState] = useState("idle");

  const [rankCatCol, setRankCatCol] = useState(catCols[0]?.name || "");
  const [benchCatCol, setBenchCatCol] = useState(catCols[0]?.name || "");
  const [benchNumCol, setBenchNumCol] = useState(numericCols[0]?.name || "");
  const [wfCatCol, setWfCatCol] = useState(catCols[0]?.name || "");
  const [wfNumCol, setWfNumCol] = useState(numericCols[0]?.name || "");
  const [corrMethod, setCorrMethod] = useState("pearson");
  const [sensitivityPct, setSensitivityPct] = useState(0);

  const activeFilterCount = [dateStart, dateEnd, filterVal1, filterVal2].filter(Boolean).length;

  const filteredRows = useMemo(() => {
    let out = rows;
    if (dateColName && (dateStart || dateEnd)) {
      out = out.filter((r) => {
        if (!(r[dateColName] instanceof Date) || Number.isNaN(+r[dateColName])) return false;
        if (dateStart && r[dateColName] < new Date(dateStart)) return false;
        if (dateEnd && r[dateColName] > new Date(dateEnd + "T23:59:59")) return false;
        return true;
      });
    }
    if (filterCol1 && filterVal1) out = out.filter((r) => String(r[filterCol1]) === filterVal1);
    if (filterCol2 && filterVal2) out = out.filter((r) => String(r[filterCol2]) === filterVal2);
    return out;
  }, [rows, dateColName, dateStart, dateEnd, filterCol1, filterVal1, filterCol2, filterVal2]);

  const d = useMemo(() => computeDashboard(filteredRows, columns), [filteredRows, columns]);
  const insights = useMemo(() => buildExecutiveInsights(d, columns), [d, columns]);
  const richInsights = useMemo(() => buildRichInsights(d), [d]);
  const alerts = useMemo(() => buildAlerts(d, columns), [d, columns]);
  const kpiSeries = useMemo(() => buildKpiSeries(filteredRows, numericCols, dateColName), [filteredRows, numericCols, dateColName]);
  const relevantCorrelations = d.allCorrPairs.filter((p) => Math.abs(p.r) > 0.5);

  // Descriptive-only insights for the Executivo sub-tab (same computed data, filtered for presentation)
  const descriptiveInsights = insights.filter((i) => !/correla|pareto/i.test(i.text));

  const periodLabel = useMemo(() => {
    if (!dateColName) return null;
    const dates = filteredRows.map((r) => r[dateColName]).filter((v) => v instanceof Date && !Number.isNaN(+v));
    if (!dates.length) return null;
    const min = new Date(Math.min(...dates)), max = new Date(Math.max(...dates));
    return `${min.toLocaleDateString("pt-BR")} — ${max.toLocaleDateString("pt-BR")}`;
  }, [filteredRows, dateColName]);

  const trendData = useMemo(() => {
    if (!dateColName || !numericCols.length) return null;
    const nc = numericCols[0].name;
    const series = filteredRows.filter((r) => r[dateColName] instanceof Date && !Number.isNaN(+r[dateColName]) && isNum(r[nc]))
      .sort((a, b) => a[dateColName] - b[dateColName]);
    if (series.length < 6) return null;
    const y = series.map((r) => r[nc]);
    const t = y.map((_v, i) => i);
    const { a, b } = linearRegression(t, y);
    const ma = movingAverage(y, Math.min(7, Math.max(3, Math.floor(y.length / 6))));
    let cum = 0;
    const chart = series.map((r, i) => { cum += y[i]; return { date: r[dateColName].toLocaleDateString("pt-BR"), valor: y[i], media_movel: ma[i], tendencia: a + b * t[i], acumulado: cum }; });
    return { chart, nc, slope: b, lastCum: cum };
  }, [filteredRows, dateColName, numericCols]);

  // Comparação entre períodos (1ª metade vs 2ª metade) — puramente descritivo
  const periodComparison = useMemo(() => kpiSeries.filter((k) => k.hasTrend).map((k) => {
    const half = k.sparkData.length ? Math.floor(k.sparkData.length / 2) : 0;
    const first = mean(k.sparkData.slice(0, half || 1));
    const second = mean(k.sparkData.slice(half || 1));
    return { name: k.name, anterior: first, atual: second };
  }), [kpiSeries]);

  // ---- ESTRATÉGICO: previsão estatística (sem Machine Learning) ----
  const stratMetric = numericCols[0]?.name || "";
  const forecast = useMemo(() => {
    if (!dateColName || !stratMetric) return null;
    const withDate = filteredRows.filter((r) => r[dateColName] instanceof Date && !Number.isNaN(+r[dateColName]) && isNum(r[stratMetric]))
      .sort((a, b) => a[dateColName] - b[dateColName]);
    if (withDate.length < 6) return null;
    const lastDate = withDate[withDate.length - 1][dateColName];
    const endOfMonth = new Date(lastDate.getFullYear(), lastDate.getMonth() + 1, 0);
    const endOfYear = new Date(lastDate.getFullYear(), 11, 31);
    const pMonth = projectTrendToDate(filteredRows, dateColName, stratMetric, endOfMonth);
    const pYear = projectTrendToDate(filteredRows, dateColName, stratMetric, endOfYear);
    if (!pMonth) return null;
    const y = withDate.map((r) => r[stratMetric]);
    const ses = exponentialSmoothing(y, 0.3);
    const chart = withDate.map((r, i) => ({ date: r[dateColName].toLocaleDateString("pt-BR"), real: y[i], suavizado: ses[i] }));
    const t0 = +withDate[0][dateColName];
    const dayMs = 86400000;
    const remainingDays = Math.max(1, Math.round((+endOfMonth - +lastDate) / dayMs));
    const step = Math.max(1, Math.round(remainingDays / 4));
    for (let i = 1; i <= 4; i++) {
      const futureDate = new Date(+lastDate + step * i * dayMs);
      const ft = (+futureDate - t0) / dayMs;
      chart.push({
        date: futureDate.toLocaleDateString("pt-BR"),
        esperado: pMonth.intercept + pMonth.slope * ft,
        otimista: pMonth.intercept + pMonth.slope * ft + pMonth.residualSd,
        pessimista: pMonth.intercept + pMonth.slope * ft - pMonth.residualSd,
      });
    }
    return { pMonth, pYear, chart, metric: stratMetric, lastDate, endOfMonth, endOfYear };
  }, [filteredRows, dateColName, stratMetric]);

  const regressionResult = useMemo(() => {
    if (numericCols.length < 2) return null;
    const response = numericCols[0].name;
    const predictors = numericCols.slice(1, 3).map((c) => c.name);
    const validRows = filteredRows.filter((r) => isNum(r[response]) && predictors.every((p) => isNum(r[p])));
    if (validRows.length < predictors.length + 3) return null;
    const X = validRows.map((r) => predictors.map((p) => r[p]));
    const y = validRows.map((r) => r[response]);
    const res = multipleRegression(X, y);
    if (!res) return null;
    return { ...res, response, predictors, n: validRows.length };
  }, [filteredRows, numericCols]);

  const impactEffortItems = useMemo(() => {
    const items = [];
    if (d.missingPct > 1) items.push({ label: "Tratar dados ausentes", impact: Math.min(100, d.missingPct * 3), effort: 55 });
    if (d.duplicates > 0) items.push({ label: "Remover duplicados", impact: Math.min(100, Math.max(20, d.duplicatePct * 4)), effort: 20 });
    if (d.outlierTotal > 0) items.push({ label: "Tratar outliers", impact: Math.min(100, Math.max(20, d.outlierPct * 3)), effort: 50 });
    if (d.pareto) items.push({ label: `Focar top categorias de "${d.pareto.column}"`, impact: 80, effort: 25 });
    if (d.topCorrelations[0]) items.push({ label: `Investigar relação ${d.topCorrelations[0].a} × ${d.topCorrelations[0].b}`, impact: Math.round(Math.abs(d.topCorrelations[0].r) * 100), effort: 40 });
    if (!items.length) items.push({ label: "Nenhuma ação prioritária identificada", impact: 20, effort: 20 });
    return items;
  }, [d]);

  const shareSummary = () => `Dashboard — ${fileName}\n${d.rowCount} linhas · ${d.colCount} colunas\nQualidade: ${Math.round(d.quality)}/100\nOutliers: ${d.outlierTotal}\nGerado em ${new Date().toLocaleString("pt-BR")}`;
  const handleShare = async () => {
    try {
      if (navigator.clipboard?.writeText) { await navigator.clipboard.writeText(shareSummary()); setShareState("copied"); }
      else throw new Error("clipboard indisponível");
    } catch { downloadBlob(shareSummary(), "resumo_dashboard.txt", "text/plain"); setShareState("downloaded"); }
    setTimeout(() => setShareState("idle"), 2000);
  };

  const dashboardActions = [];
  if (d.missingPct > 5) dashboardActions.push("Tratar valores ausentes antes de análises mais profundas.");
  if (d.duplicates > 0) dashboardActions.push("Investigar e remover registros duplicados.");
  if (d.outlierTotal > 0) dashboardActions.push("Revisar outliers — podem ser erros de coleta ou eventos legítimos raros.");
  if (d.pareto) dashboardActions.push(`Concentrar esforços nas ${d.pareto.categoriesFor80} categorias de "${d.pareto.column}" responsáveis por 80% da base.`);
  if (!dashboardActions.length) dashboardActions.push("Base em boas condições — nenhuma ação corretiva crítica identificada.");

  const alertLevelMap = { critico: "warning", atencao: "warning", info: "info", ok: "success" };
  const alertIconMap = { critico: Flame, atencao: AlertTriangle, info: Info, ok: CheckCircle2 };

  const SUB_TABS = [
    { id: "executivo", label: "Executivo", icon: LayoutDashboard, hint: "O que aconteceu?" },
    { id: "analitico", label: "Analítico", icon: Compass, hint: "Por que aconteceu?" },
    { id: "estrategico", label: "Estratégico", icon: Rocket, hint: "O que fazer?" },
  ];

  return (
    <div style={{ background: theme === "dark" ? C.bg : "transparent", margin: theme === "dark" ? -28 : 0, padding: theme === "dark" ? 28 : 0, borderRadius: theme === "dark" ? 14 : 0, transition: "background 0.25s ease" }}>
      <style>{`.dash-fade { animation: dashFadeIn 0.35s ease; } @keyframes dashFadeIn { from { opacity:0; transform:translateY(4px);} to { opacity:1; transform:translateY(0);} }`}</style>

      {/* CABEÇALHO */}
      <div className="dash-fade" style={{ background: C.panel, border: `1px solid ${C.border}`, borderRadius: 16, padding: "20px 24px", marginBottom: 16, display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 16 }}>
        <div style={{ display: "flex", gap: 14 }}>
          <div style={{ width: 44, height: 44, borderRadius: 12, background: C.tealSoft, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
            <LayoutDashboard size={21} color={C.teal} />
          </div>
          <div>
            <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 11, letterSpacing: "0.1em", color: C.teal, textTransform: "uppercase", fontWeight: 700 }}>Dashboard</div>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 23, fontWeight: 700, color: C.ink, marginTop: 2 }}>{fileName}</div>
            <div style={{ fontSize: 12.5, color: C.sub, marginTop: 4 }}>
              {d.rowCount.toLocaleString("pt-BR")} registros · atualizado às {refreshedAt.toLocaleTimeString("pt-BR")}{periodLabel && ` · período: ${periodLabel}`}
            </div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <button onClick={() => setTheme((t) => (t === "light" ? "dark" : "light"))} title="Alternar tema"
            style={{ width: 36, height: 36, borderRadius: 9, border: `1px solid ${C.border}`, background: "transparent", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>
            {theme === "light" ? <Moon size={15} color={C.ink} /> : <Sun size={15} color={C.ink} />}
          </button>
          <Btn variant="ghost" icon={Share2} style={{ background: "transparent", color: C.ink, borderColor: C.border }} onClick={handleShare}>
            {shareState === "idle" ? "Compartilhar" : shareState === "copied" ? "Copiado!" : "Baixado"}
          </Btn>
          <Btn variant="ghost" icon={RefreshCw} style={{ background: "transparent", color: C.ink, borderColor: C.border }} onClick={() => setRefreshedAt(new Date())}>Atualizar dados</Btn>
        </div>
      </div>

      {/* SUB-NAVEGAÇÃO: Executivo / Analítico / Estratégico */}
      <div className="dash-fade" style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
        {SUB_TABS.map((s) => {
          const active = subTab === s.id;
          return (
            <button key={s.id} onClick={() => setSubTab(s.id)}
              style={{
                display: "flex", alignItems: "center", gap: 8, padding: "10px 16px", borderRadius: 11,
                border: `1px solid ${active ? C.teal : C.border}`, background: active ? C.tealSoft : C.panel,
                cursor: "pointer", flex: "1 1 200px", textAlign: "left",
              }}>
              <s.icon size={16} color={active ? C.teal : C.faint} />
              <div>
                <div style={{ fontWeight: 700, fontSize: 13, color: active ? (C.tealDark || C.teal) : C.ink }}>{s.label}</div>
                <div style={{ fontSize: 10.5, color: C.faint }}>{s.hint}</div>
              </div>
            </button>
          );
        })}
      </div>

      {/* FILTROS GLOBAIS (compartilhados entre as três subabas) */}
      {catCols.length > 0 && (
        <FilterBar C={C} rows={rows} dateCols={dateCols} catCols={catCols}
          dateStart={dateStart} setDateStart={setDateStart} dateEnd={dateEnd} setDateEnd={setDateEnd}
          filterCol1={filterCol1} setFilterCol1={setFilterCol1} filterVal1={filterVal1} setFilterVal1={setFilterVal1}
          filterCol2={filterCol2} setFilterCol2={setFilterCol2} filterVal2={filterVal2} setFilterVal2={setFilterVal2}
          onReset={() => { setDateStart(""); setDateEnd(""); setFilterCol1(""); setFilterVal1(""); setFilterCol2(""); setFilterVal2(""); }}
          activeCount={activeFilterCount} />
      )}

      {/* =========================== EXECUTIVO — Análise Descritiva =========================== */}
      {subTab === "executivo" && (
        <div className="dash-fade">
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12, marginBottom: 8 }}>
            <KpiCard C={C} icon={Layers} label="Registros" value={d.rowCount.toLocaleString("pt-BR")} sub={`${d.colCount} variáveis`} />
            <KpiCard C={C} icon={Gauge} label="Qualidade dos dados" value={`${Math.round(d.quality)}/100`} tone={d.quality >= 80 ? "green" : d.quality >= 55 ? "amber" : "red"} progress={d.quality} />
            <KpiCard C={C} icon={Info} label="Valores ausentes" value={`${fmt(d.missingPct, 1)}%`} tone={d.missingPct > 5 ? "amber" : "green"} />
            <KpiCard C={C} icon={AlertTriangle} label="Outliers (IQR)" value={d.outlierTotal} tone={d.outlierTotal > 0 ? "amber" : "green"} />
            <KpiCard C={C} icon={trendData?.slope > 0 ? TrendingUp : TrendingDown} label="Tendência geral"
              value={trendData ? (trendData.slope > 0 ? "Alta" : trendData.slope < 0 ? "Baixa" : "Estável") : "—"}
              sub={trendData ? trendData.nc : "sem série temporal"} tone={trendData?.slope > 0 ? "green" : trendData?.slope < 0 ? "red" : "neutral"} />
            <KpiCard C={C} icon={ShieldCheck} label="Status geral" value={d.quality >= 80 ? "Saudável" : d.quality >= 55 ? "Atenção" : "Crítico"} tone={d.quality >= 80 ? "green" : d.quality >= 55 ? "amber" : "red"} />
          </div>

          {kpiSeries.length > 0 && (
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 12, marginBottom: 16, marginTop: 12 }}>
              {kpiSeries.map((k, i) => (
                <div key={i} style={{ background: C.panel, border: `1px solid ${C.border}`, borderRadius: 14, padding: "16px 18px" }}>
                  <div style={{ fontSize: 11, color: C.sub, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em" }}>{k.name}</div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", marginTop: 6 }}>
                    <div>
                      <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 21, fontWeight: 700, color: C.ink }}>{fmt(k.avg, 2)}</div>
                      {k.growthPct !== null && (
                        <div style={{ display: "flex", alignItems: "center", gap: 3, fontSize: 11.5, color: k.growthPct >= 0 ? C.green : C.red, marginTop: 2 }}>
                          {k.growthPct >= 0 ? <ArrowUpRight size={12} /> : <ArrowDownRight size={12} />} {fmt(Math.abs(k.growthPct), 1)}% {k.hasTrend ? "vs. 1ª metade do período" : ""}
                        </div>
                      )}
                    </div>
                    <Sparkline data={k.sparkData} color={k.growthPct === null || k.growthPct >= 0 ? C.teal : C.red} />
                  </div>
                  <div style={{ fontSize: 10.5, color: C.faint, marginTop: 4 }}>{k.hasTrend ? "média por intervalo de tempo" : "forma da distribuição"} · n={k.n}</div>
                </div>
              ))}
            </div>
          )}

          <div style={{ background: C.panel, border: `1px solid ${C.border}`, borderRadius: 16, padding: 20, marginBottom: 16 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
              <Sparkles size={15} color={C.teal} /> <div style={{ fontWeight: 700, fontSize: 14, color: C.ink }}>Resumo executivo automático</div>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {descriptiveInsights.map((ins, i) => (
                <div key={i} style={{ display: "flex", gap: 8, alignItems: "flex-start", fontSize: 13, color: C.ink, lineHeight: 1.5 }}>
                  <span style={{ marginTop: 5, width: 6, height: 6, borderRadius: 999, background: C[ins.tone] || C.faint, flexShrink: 0 }} />
                  {ins.text}
                </div>
              ))}
            </div>
          </div>

          <ReportDivider C={C} label="Indicadores de qualidade dos dados" />
          <div style={{ display: "grid", gridTemplateColumns: "1.1fr 1.4fr", gap: 16, marginBottom: 16 }}>
            <Card style={{ padding: 20, display: "flex", alignItems: "center", gap: 20, background: C.panel, border: `1px solid ${C.border}` }}>
              <QualityRing score={d.quality} />
              <div>
                <div style={{ fontSize: 12, fontWeight: 700, color: C.sub, textTransform: "uppercase", letterSpacing: "0.06em" }}>Score de confiabilidade</div>
                <div style={{ fontSize: 13, color: C.ink, marginTop: 6, lineHeight: 1.5 }}>Combina ausentes, duplicados e outliers (IQR) em um único indicador 0–100.</div>
                <div style={{ marginTop: 10 }}><Pill tone={d.quality >= 80 ? "green" : d.quality >= 55 ? "amber" : "red"}>{d.quality >= 80 ? "Confiável" : d.quality >= 55 ? "Atenção" : "Crítico"}</Pill></div>
              </div>
            </Card>
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              <div>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11.5, color: C.sub, marginBottom: 4 }}><span>Completude</span><span>{fmt(100 - d.missingPct, 1)}%</span></div>
                <div style={{ background: C.border, borderRadius: 999, height: 7 }}><div style={{ width: `${100 - d.missingPct}%`, height: "100%", borderRadius: 999, background: C.teal }} /></div>
              </div>
              <div>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11.5, color: C.sub, marginBottom: 4 }}><span>Consistência (sem duplicados)</span><span>{fmt(100 - d.duplicatePct, 1)}%</span></div>
                <div style={{ background: C.border, borderRadius: 999, height: 7 }}><div style={{ width: `${100 - d.duplicatePct}%`, height: "100%", borderRadius: 999, background: C.blue }} /></div>
              </div>
              <div>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11.5, color: C.sub, marginBottom: 4 }}><span>Ausência de outliers</span><span>{fmt(100 - Math.min(100, d.outlierPct), 1)}%</span></div>
                <div style={{ background: C.border, borderRadius: 999, height: 7 }}><div style={{ width: `${100 - Math.min(100, d.outlierPct)}%`, height: "100%", borderRadius: 999, background: C.amber }} /></div>
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 4 }}>
                {Object.entries(d.dtypeDist).map(([k, v]) => <Pill key={k} tone={k === "numeric" ? "teal" : k === "date" ? "amber" : "neutral"}>{k}: {v}</Pill>)}
              </div>
            </div>
          </div>

          {trendData && (
            <>
              <ReportDivider C={C} label="Evolução temporal & tendências" />
              <Card style={{ padding: 20, marginBottom: 16, background: C.panel, border: `1px solid ${C.border}` }}>
                <ChartFrame C={C} title={`Evolução temporal — ${trendData.nc}`} subtitle={`Real, média móvel e tendência. Inclinação ≈ ${fmt(trendData.slope, 3)} (${trendData.slope > 0 ? "crescente" : trendData.slope < 0 ? "decrescente" : "estável"}).`}>
                  {(h) => (
                    <ResponsiveContainer width="100%" height={h}>
                      <ComposedChart data={trendData.chart}>
                        <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                        <XAxis dataKey="date" tick={{ fontSize: 9 }} minTickGap={30} />
                        <YAxis tick={{ fontSize: 10 }} />
                        <Tooltip /><Legend wrapperStyle={{ fontSize: 11 }} />
                        <Area type="monotone" dataKey="valor" name="Real" stroke={C.teal} fill={C.tealSoft} />
                        <Line type="monotone" dataKey="media_movel" name="Média móvel" stroke={C.blue} strokeWidth={2} dot={false} />
                        <Line type="monotone" dataKey="tendencia" name="Tendência" stroke={C.red} strokeWidth={2} strokeDasharray="4 3" dot={false} />
                      </ComposedChart>
                    </ResponsiveContainer>
                  )}
                </ChartFrame>
              </Card>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
                <Card style={{ padding: 20, background: C.panel, border: `1px solid ${C.border}` }}>
                  <div style={{ fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 10 }}>Acumulado</div>
                  <ResponsiveContainer width="100%" height={190}>
                    <AreaChart data={trendData.chart}>
                      <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                      <XAxis dataKey="date" tick={{ fontSize: 9 }} minTickGap={30} /><YAxis tick={{ fontSize: 10 }} />
                      <Tooltip /><Area type="monotone" dataKey="acumulado" stroke={C.blue} fill={C.tealSoft} />
                    </AreaChart>
                  </ResponsiveContainer>
                </Card>
                <Card style={{ padding: 20, background: C.panel, border: `1px solid ${C.border}` }}>
                  <div style={{ fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 10 }}>Comparação entre períodos (1ª vs. 2ª metade)</div>
                  {periodComparison.length ? (
                    <ResponsiveContainer width="100%" height={190}>
                      <BarChart data={periodComparison}>
                        <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                        <XAxis dataKey="name" tick={{ fontSize: 9 }} /><YAxis tick={{ fontSize: 10 }} />
                        <Tooltip /><Legend wrapperStyle={{ fontSize: 11 }} />
                        <Bar dataKey="anterior" name="Período anterior" fill={C.border} radius={[3, 3, 0, 0]} />
                        <Bar dataKey="atual" name="Período atual" fill={C.teal} radius={[3, 3, 0, 0]} />
                      </BarChart>
                    </ResponsiveContainer>
                  ) : <div style={{ color: C.sub, fontSize: 12.5 }}>Necessita de uma coluna de data para comparar períodos.</div>}
                </Card>
              </div>
            </>
          )}

          {catCols.length > 0 && (
            <>
              <ReportDivider C={C} label="Rankings" />
              <div style={{ marginBottom: 16 }}><TopBottomCard C={C} rows={filteredRows} catCols={catCols} catCol={rankCatCol} setCatCol={setRankCatCol} /></div>
            </>
          )}

          <ReportDivider C={C} label="Alertas" />
          <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 16 }}>
            {alerts.map((a, i) => <Callout key={i} C={C} tone={alertLevelMap[a.level] || "info"} icon={alertIconMap[a.level] || Info}>{a.text}</Callout>)}
          </div>

          <Callout C={C} tone={d.quality >= 80 ? "success" : d.quality >= 55 ? "warning" : "warning"} icon={ShieldCheck}>
            Status geral do dataset: base "{fileName}" com {d.rowCount.toLocaleString("pt-BR")} registros, qualidade {Math.round(d.quality)}/100 — {d.quality >= 80 ? "pronta para uso analítico." : d.quality >= 55 ? "recomenda-se atenção antes de decisões críticas." : "requer tratamento antes de qualquer decisão."}
          </Callout>
        </div>
      )}

      {/* =========================== ANALÍTICO — Análise Diagnóstica =========================== */}
      {subTab === "analitico" && (
        <div className="dash-fade">
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12, marginBottom: 16 }}>
            <KpiCard C={C} icon={Activity} label="Correlações relevantes" value={relevantCorrelations.length} sub="|r| > 0.5" tone={relevantCorrelations.length ? "teal" : "neutral"} />
            <KpiCard C={C} icon={AlertTriangle} label="Outliers (IQR)" value={d.outlierTotal} tone={d.outlierTotal > 0 ? "amber" : "green"} />
            <KpiCard C={C} icon={Info} label="Duplicados" value={d.duplicates} sub={`${fmt(d.duplicatePct, 1)}%`} tone={d.duplicates > 0 ? "amber" : "green"} />
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 14, marginBottom: 16 }}>
            {richInsights.map((ins, i) => (
              <div key={i} style={{ background: C.panel, border: `1px solid ${C.border}`, borderRadius: 14, padding: 16 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                  <div style={{ width: 28, height: 28, borderRadius: 8, background: C[`${ins.tone}Soft`] || C.tealSoft, display: "flex", alignItems: "center", justifyContent: "center" }}>
                    <ins.icon size={14} color={C[ins.tone] || C.teal} />
                  </div>
                  <div style={{ fontWeight: 700, fontSize: 12.5, color: C.ink }}>{ins.title}</div>
                </div>
                <div style={{ fontSize: 12.5, color: C.ink, marginBottom: 6, lineHeight: 1.5 }}>{ins.simple}</div>
                <div style={{ fontSize: 11, color: C.faint, marginBottom: 10, fontFamily: "'JetBrains Mono', monospace" }}>{ins.technical}</div>
                <div style={{ display: "flex", gap: 6 }}>
                  <Pill tone={ins.impact === "alto" ? "red" : ins.impact === "médio" ? "amber" : "neutral"}>impacto {ins.impact}</Pill>
                  <Pill tone="teal">{ins.confidence}% confiança</Pill>
                </div>
              </div>
            ))}
          </div>

          <ReportDivider C={C} label="Diagnóstico — outliers, ausentes e duplicidades" />
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(340px, 1fr))", gap: 16, marginBottom: 16 }}>
            <Card style={{ padding: 20, background: C.panel, border: `1px solid ${C.border}` }}>
              <div style={{ fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 12 }}>Outliers por variável (IQR)</div>
              {d.outliersByCol.length ? (
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {d.outliersByCol.slice(0, 8).map((o, i) => (
                    <div key={i} style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, color: C.ink, borderBottom: `1px solid ${C.border}`, paddingBottom: 6 }}>
                      <span>{o.name}</span><span style={{ fontFamily: "'JetBrains Mono', monospace", color: C.amber }}>{o.count} ({fmt(o.pct, 1)}%)</span>
                    </div>
                  ))}
                </div>
              ) : <div style={{ color: C.sub, fontSize: 12.5 }}>Nenhum outlier identificado.</div>}
            </Card>
            <Card style={{ padding: 20, background: C.panel, border: `1px solid ${C.border}` }}>
              <div style={{ fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 12 }}>Dados ausentes por coluna</div>
              <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: 220, overflowY: "auto" }}>
                {[...columns].sort((a, b) => b.missingPct - a.missingPct).slice(0, 8).map((c, i) => (
                  <div key={i} style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, color: C.ink, borderBottom: `1px solid ${C.border}`, paddingBottom: 6 }}>
                    <span>{c.name}</span><span style={{ fontFamily: "'JetBrains Mono', monospace", color: c.missingPct > 5 ? C.amber : C.sub }}>{fmt(c.missingPct, 1)}%</span>
                  </div>
                ))}
              </div>
            </Card>
            <Card style={{ padding: 20, background: C.panel, border: `1px solid ${C.border}` }}>
              <div style={{ fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 12 }}>Duplicidades</div>
              <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 28, fontWeight: 700, color: d.duplicates > 0 ? C.amber : C.green }}>{d.duplicates}</div>
              <div style={{ fontSize: 12, color: C.sub, marginTop: 4 }}>{fmt(d.duplicatePct, 1)}% das linhas são duplicatas exatas.</div>
            </Card>
          </div>

          {catCols.length > 0 && numericCols.length > 0 && (
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(360px, 1fr))", gap: 16, marginBottom: 16 }}>
              <BenchmarkCard C={C} rows={filteredRows} catCols={catCols} numericCols={numericCols} catCol={benchCatCol} setCatCol={setBenchCatCol} numCol={benchNumCol} setNumCol={setBenchNumCol} />
              <WaterfallCard C={C} rows={filteredRows} catCols={catCols} numericCols={numericCols} catCol={wfCatCol} setCatCol={setWfCatCol} numCol={wfNumCol} setNumCol={setWfNumCol} />
            </div>
          )}

          {numericCols.length >= 2 && (
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(340px, 1fr))", gap: 16, marginBottom: 16 }}>
              <Card style={{ padding: 20, background: C.panel, border: `1px solid ${C.border}` }}>
                <div style={{ fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 12 }}>Força das relações entre variáveis</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: 260, overflowY: "auto" }}>
                  {d.allCorrPairs.slice(0, 12).map((p, i) => {
                    const strength = Math.abs(p.r) > 0.7 ? "forte" : Math.abs(p.r) > 0.4 ? "moderada" : "fraca";
                    const tone = Math.abs(p.r) > 0.7 ? "green" : Math.abs(p.r) > 0.4 ? "amber" : "neutral";
                    return (
                      <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 12, borderBottom: `1px solid ${C.border}`, paddingBottom: 6 }}>
                        <span style={{ color: C.ink }}>{p.a} × {p.b}</span>
                        <div style={{ display: "flex", gap: 6, alignItems: "center" }}><Pill tone={tone}>{strength}</Pill><span style={{ fontFamily: "'JetBrains Mono', monospace", color: C.sub }}>{fmt(p.r, 2)}</span></div>
                      </div>
                    );
                  })}
                  {!d.allCorrPairs.length && <div style={{ color: C.sub, fontSize: 12.5 }}>Sem pares numéricos suficientes.</div>}
                </div>
              </Card>
              {numericCols.length >= 3 && <RadarProfileCard C={C} rows={filteredRows} numericCols={numericCols} />}
            </div>
          )}

          <ReportDivider C={C} label="Exploração de dados (EDA) — todas as funcionalidades existentes" />
          <EDATab rows={filteredRows} columns={columns} />
        </div>
      )}

      {/* =========================== ESTRATÉGICO — Análise Preditiva e Prescritiva =========================== */}
      {subTab === "estrategico" && (
        <div className="dash-fade">
          <Callout C={C} tone="info" icon={Info}>
            Esta aba usa exclusivamente técnicas estatísticas (regressão, suavização exponencial, intervalos de confiança) — nenhum modelo de Machine Learning é utilizado.
          </Callout>

          <ReportDivider C={C} label="Preditiva — forecast estatístico" />
          {forecast ? (
            <>
              <Card style={{ padding: 20, marginBottom: 16, background: C.panel, border: `1px solid ${C.border}` }}>
                <ChartFrame C={C} title={`Projeção — ${forecast.metric}`} subtitle={`Real, suavização exponencial (α=0.3) e cenários até o fechamento do mês (${forecast.endOfMonth.toLocaleDateString("pt-BR")}).`}>
                  {(h) => (
                    <ResponsiveContainer width="100%" height={h}>
                      <ComposedChart data={forecast.chart}>
                        <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
                        <XAxis dataKey="date" tick={{ fontSize: 9 }} minTickGap={30} />
                        <YAxis tick={{ fontSize: 10 }} />
                        <Tooltip /><Legend wrapperStyle={{ fontSize: 11 }} />
                        <Area type="monotone" dataKey="real" name="Real" stroke={C.teal} fill={C.tealSoft} />
                        <Line type="monotone" dataKey="suavizado" name="Suavização exponencial" stroke={C.blue} strokeWidth={2} dot={false} />
                        <Line type="monotone" dataKey="otimista" name="Cenário otimista" stroke={C.green} strokeWidth={1.5} strokeDasharray="2 2" dot={false} />
                        <Line type="monotone" dataKey="esperado" name="Cenário esperado" stroke={C.red} strokeWidth={2} strokeDasharray="4 3" dot={false} />
                        <Line type="monotone" dataKey="pessimista" name="Cenário pessimista" stroke={C.amber} strokeWidth={1.5} strokeDasharray="2 2" dot={false} />
                      </ComposedChart>
                    </ResponsiveContainer>
                  )}
                </ChartFrame>
              </Card>

              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 12, marginBottom: 16 }}>
                <KpiCard C={C} icon={TrendingDown} label="Cenário pessimista (fim do mês)" value={fmt(forecast.pMonth.pessimistic, 2)} tone="amber" />
                <KpiCard C={C} icon={Target} label="Cenário esperado (fim do mês)" value={fmt(forecast.pMonth.expected, 2)} tone="teal" />
                <KpiCard C={C} icon={TrendingUp} label="Cenário otimista (fim do mês)" value={fmt(forecast.pMonth.optimistic, 2)} tone="green" />
                <KpiCard C={C} icon={Calendar} label="Projeção fim de ano" value={fmt(forecast.pYear?.expected, 2)} sub={forecast.endOfYear.toLocaleDateString("pt-BR")} />
              </div>

              <Card style={{ padding: 20, marginBottom: 16, background: C.panel, border: `1px solid ${C.border}` }}>
                <div style={{ fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 6 }}>Análise de sensibilidade / simulação de metas</div>
                <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 14 }}>Ajuste o percentual para simular um cenário diferente sobre a projeção esperada (what-if).</div>
                <input type="range" min={-50} max={50} value={sensitivityPct} onChange={(e) => setSensitivityPct(Number(e.target.value))} style={{ width: "100%" }} />
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: C.sub, marginTop: 6 }}>
                  <span>Ajuste: <b style={{ color: C.ink }}>{sensitivityPct > 0 ? "+" : ""}{sensitivityPct}%</b></span>
                  <span>Resultado simulado: <b style={{ color: C.teal }}>{fmt(forecast.pMonth.expected * (1 + sensitivityPct / 100), 2)}</b></span>
                </div>
              </Card>
            </>
          ) : (
            <Callout C={C} tone="warning" icon={Info}>Projeção indisponível — é necessária uma coluna de data e ao menos 6 observações numéricas válidas.</Callout>
          )}

          {regressionResult && (
            <Card style={{ padding: 20, marginBottom: 16, background: C.panel, border: `1px solid ${C.border}` }}>
              <div style={{ fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 4 }}>
                {regressionResult.predictors.length > 1 ? "Regressão múltipla" : "Regressão linear"} — {regressionResult.response} ~ {regressionResult.predictors.join(" + ")}
              </div>
              <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 12 }}>R² = {fmt(regressionResult.r2, 3)} · n = {regressionResult.n} · graus de liberdade = {regressionResult.df}</div>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead><tr style={{ textAlign: "left", color: C.sub, borderBottom: `1px solid ${C.border}` }}>
                    <th style={{ padding: "6px 8px" }}>Termo</th><th style={{ padding: "6px 8px" }}>Coeficiente</th><th style={{ padding: "6px 8px" }}>t</th><th style={{ padding: "6px 8px" }}>p-valor</th><th style={{ padding: "6px 8px" }}>Significância</th>
                  </tr></thead>
                  <tbody>
                    {["Intercepto", ...regressionResult.predictors].map((name, i) => (
                      <tr key={i} style={{ borderBottom: `1px solid ${C.border}` }}>
                        <td style={{ padding: "6px 8px", fontWeight: 600, color: C.ink }}>{name}</td>
                        <td style={{ padding: "6px 8px", color: C.ink, fontFamily: "'JetBrains Mono', monospace" }}>{fmt(regressionResult.coefStats[i].coef, 4)}</td>
                        <td style={{ padding: "6px 8px", color: C.ink, fontFamily: "'JetBrains Mono', monospace" }}>{fmt(regressionResult.coefStats[i].t, 2)}</td>
                        <td style={{ padding: "6px 8px", color: C.ink, fontFamily: "'JetBrains Mono', monospace" }}>{fmtP(regressionResult.coefStats[i].p)}</td>
                        <td style={{ padding: "6px 8px" }}><Pill tone={regressionResult.coefStats[i].p < 0.05 ? "green" : "amber"}>{regressionResult.coefStats[i].p < 0.05 ? "Significativo" : "Não significativo"}</Pill></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{ fontSize: 11, color: C.faint, marginTop: 10 }}>Coeficientes com p &lt; 0.05 indicam efeito estatisticamente significativo do preditor sobre {regressionResult.response}, mantendo as demais variáveis constantes.</div>
            </Card>
          )}

          <ReportDivider C={C} label="Prescritiva — plano de ação e priorização" />
          <div style={{ display: "grid", gridTemplateColumns: "1.1fr 1fr", gap: 16, marginBottom: 16 }}>
            <Card style={{ padding: 20, background: C.panel, border: `1px solid ${C.border}` }}>
              <div style={{ fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 4 }}>Matriz de priorização (Impacto × Esforço)</div>
              <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 8 }}>Priorize o quadrante superior-esquerdo: alto impacto, baixo esforço.</div>
              <ImpactEffortMatrix C={C} items={impactEffortItems} />
            </Card>
            <Card style={{ padding: 20, background: C.panel, border: `1px solid ${C.border}` }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}><ClipboardList size={15} color={C.teal} /><div style={{ fontWeight: 700, fontSize: 14, color: C.ink }}>Plano de ação recomendado</div></div>
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {dashboardActions.map((a, i) => <div key={i} style={{ fontSize: 12.5, color: C.ink, lineHeight: 1.5 }}>{i + 1}. {a}</div>)}
              </div>
            </Card>
          </div>

          <div style={{ background: C.panel, border: `1px solid ${C.border}`, borderRadius: 16, padding: 22, marginBottom: 8 }}>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 20 }}>
              <div>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}><ShieldCheck size={15} color={C.teal} /><div style={{ fontWeight: 700, fontSize: 13, color: C.ink }}>Situação geral</div></div>
                <div style={{ fontSize: 12.5, color: C.ink, lineHeight: 1.6 }}>
                  A base "{fileName}" tem qualidade <b>{Math.round(d.quality)}/100</b> ({d.quality >= 80 ? "confiável" : d.quality >= 55 ? "requer atenção" : "crítica"}).
                  {forecast && ` Projeção esperada para o fechamento do mês: ${fmt(forecast.pMonth.expected, 2)}.`}
                </div>
              </div>
              <div>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}><ShieldAlert size={15} color={C.amber} /><div style={{ fontWeight: 700, fontSize: 13, color: C.ink }}>Maiores riscos</div></div>
                <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                  {alerts.filter((a) => a.level === "critico" || a.level === "atencao").slice(0, 3).map((a, i) => <div key={i} style={{ fontSize: 12, color: C.ink, lineHeight: 1.5 }}>• {a.text}</div>)}
                  {!alerts.some((a) => a.level === "critico" || a.level === "atencao") && <div style={{ fontSize: 12, color: C.sub }}>Nenhum risco relevante identificado.</div>}
                </div>
              </div>
              <div>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}><Target size={15} color={C.teal} /><div style={{ fontWeight: 700, fontSize: 13, color: C.ink }}>Maiores oportunidades</div></div>
                <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                  {richInsights.filter((r) => r.tone === "teal" || r.tone === "green").slice(0, 3).map((r, i) => <div key={i} style={{ fontSize: 12, color: C.ink, lineHeight: 1.5 }}>• {r.title}</div>)}
                </div>
              </div>
              <div>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}><Rocket size={15} color={C.red} /><div style={{ fontWeight: 700, fontSize: 13, color: C.ink }}>Próximas ações</div></div>
                <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                  {dashboardActions.map((a, i) => <div key={i} style={{ fontSize: 12, color: C.ink, lineHeight: 1.5 }}>{i + 1}. {a}</div>)}
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      <div style={{ marginTop: 12, fontSize: 11, color: C.faint, textAlign: "center" }}>
        Dashboard gerado automaticamente pela AnálisePro · {refreshedAt.toLocaleString("pt-BR")}
      </div>
    </div>
  );
}

/* =========================================================================
   EDA TAB
   ========================================================================= */
function histogramBins(vals, binCount = 12) {
  if (!vals.length) return [];
  const min = Math.min(...vals), max = Math.max(...vals);
  if (min === max) return [{ label: fmt(min), count: vals.length }];
  const width = (max - min) / binCount;
  const bins = Array.from({ length: binCount }, (_, i) => ({
    x0: min + i * width, x1: min + (i + 1) * width, count: 0,
  }));
  vals.forEach((v) => {
    let idx = Math.floor((v - min) / width);
    if (idx >= binCount) idx = binCount - 1;
    if (idx < 0) idx = 0;
    bins[idx].count++;
  });
  return bins.map((b) => ({ label: `${fmt(b.x0, 1)}–${fmt(b.x1, 1)}`, count: b.count }));
}

function BoxPlotSVG({ vals, width = 560, height = 190, color = T.teal }) {
  if (!vals.length) return null;
  const { q1, q2, q3 } = quartiles(vals);
  const box = q3 - q1;
  const lo = Math.max(Math.min(...vals), q1 - 1.5 * box);
  const hi = Math.min(Math.max(...vals), q3 + 1.5 * box);
  const outliers = vals.filter((v) => v < lo || v > hi);
  const min = Math.min(...vals), max = Math.max(...vals);
  const pad = 40;
  const scale = (v) => pad + ((v - min) / (max - min || 1)) * (width - pad * 2);
  const cy = height / 2 + 24;
  // deterministic pseudo-random jitter for a strip/violin-style rug of points
  const seeded = (i) => { const x = Math.sin(i * 12.9898) * 43758.5453; return x - Math.floor(x); };
  const stripY = cy - 56;
  return (
    <svg width="100%" viewBox={`0 0 ${width} ${height}`} preserveAspectRatio="xMidYMid meet">
      {vals.slice(0, 400).map((v, i) => (
        <circle key={i} cx={scale(v)} cy={stripY + (seeded(i) - 0.5) * 26} r="2.2" fill={color} fillOpacity="0.28" />
      ))}
      <line x1={scale(lo)} y1={cy} x2={scale(q1)} y2={cy} stroke={T.faint} strokeWidth="1.5" />
      <line x1={scale(q3)} y1={cy} x2={scale(hi)} y2={cy} stroke={T.faint} strokeWidth="1.5" />
      <line x1={scale(lo)} y1={cy - 14} x2={scale(lo)} y2={cy + 14} stroke={T.faint} strokeWidth="1.5" />
      <line x1={scale(hi)} y1={cy - 14} x2={scale(hi)} y2={cy + 14} stroke={T.faint} strokeWidth="1.5" />
      <rect x={scale(q1)} y={cy - 26} width={Math.max(scale(q3) - scale(q1), 1)} height="52" fill={color} fillOpacity="0.18" stroke={color} strokeWidth="1.5" rx="4" />
      <line x1={scale(q2)} y1={cy - 26} x2={scale(q2)} y2={cy + 26} stroke={color} strokeWidth="2.5" />
      {outliers.map((o, i) => <circle key={i} cx={scale(o)} cy={cy} r="3.5" fill={T.red} fillOpacity="0.65" />)}
      {[min, q1, q2, q3, max].map((v, i) => (
        <text key={i} x={scale(v)} y={height - 8} textAnchor="middle" fontSize="10.5" fontFamily="'JetBrains Mono', monospace" fill={T.sub}>{fmt(v, 1)}</text>
      ))}
      <text x={pad} y={stripY - 34} fontSize="10" fontFamily="'JetBrains Mono', monospace" fill={T.faint}>dispersão dos pontos (rug)</text>
    </svg>
  );
}

function ParetoCard({ rows, catCols, catCol, setCatCol }) {
  const data = useMemo(() => {
    if (!catCol) return { bars: [], categoriesFor80: 0, total: 0 };
    const counts = _.countBy(rows.map((r) => r[catCol]).filter((v) => v !== null && v !== undefined));
    const sorted = Object.entries(counts).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
    const total = _.sum(sorted.map((s) => s.value));
    let cum = 0, categoriesFor80 = sorted.length;
    const bars = sorted.slice(0, 15).map((s, i) => {
      cum += s.value;
      if (cum / total >= 0.8 && categoriesFor80 === sorted.length && i < sorted.length) categoriesFor80 = i + 1;
      return { name: s.name, value: s.value, cumPct: Math.round((cum / total) * 1000) / 10 };
    });
    return { bars, categoriesFor80, total: sorted.length };
  }, [rows, catCol]);

  return (
    <Card style={{ padding: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontWeight: 700, fontSize: 14, color: T.ink }}>Gráfico de Pareto (80/20)</div>
        <Select value={catCol} onChange={setCatCol} options={catCols.map((c) => c.name)} />
      </div>
      <div style={{ fontSize: 11.5, color: T.faint, marginBottom: 10 }}>
        {data.categoriesFor80} de {data.total} categorias concentram 80% dos registros.
      </div>
      <ResponsiveContainer width="100%" height={240}>
        <ComposedChart data={data.bars}>
          <CartesianGrid strokeDasharray="3 3" stroke={T.border} />
          <XAxis dataKey="name" tick={{ fontSize: 9 }} interval={0} angle={-30} textAnchor="end" height={60} />
          <YAxis yAxisId="left" tick={{ fontSize: 10 }} />
          <YAxis yAxisId="right" orientation="right" domain={[0, 100]} tick={{ fontSize: 10 }} />
          <Tooltip />
          <Bar yAxisId="left" dataKey="value" fill={T.teal} radius={[3, 3, 0, 0]} />
          <Line yAxisId="right" type="monotone" dataKey="cumPct" stroke={T.amber} strokeWidth={2.5} dot={{ r: 2.5 }} />
        </ComposedChart>
      </ResponsiveContainer>
    </Card>
  );
}

function TreemapCard({ rows, catCols, catCol, setCatCol }) {
  const TREEMAP_COLORS = [T.teal, T.blue, T.amber, T.green, "#7C6FDB", "#E07A5F", "#3D5A80", "#C77DFF", T.red, "#4C956C"];
  const data = useMemo(() => {
    if (!catCol) return [];
    const counts = _.countBy(rows.map((r) => r[catCol]).filter((v) => v !== null && v !== undefined));
    return Object.entries(counts).map(([name, size]) => ({ name, size })).sort((a, b) => b.size - a.size).slice(0, 20);
  }, [rows, catCol]);
  return (
    <Card style={{ padding: 20 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontWeight: 700, fontSize: 14, color: T.ink }}>Treemap de categorias</div>
        <Select value={catCol} onChange={setCatCol} options={catCols.map((c) => c.name)} />
      </div>
      <ResponsiveContainer width="100%" height={240}>
        <Treemap data={data} dataKey="size" nameKey="name" stroke="#fff" fill={T.teal} aspectRatio={4 / 3}>
          {data.map((_e, i) => <Cell key={i} fill={TREEMAP_COLORS[i % TREEMAP_COLORS.length]} />)}
          <Tooltip />
        </Treemap>
      </ResponsiveContainer>
    </Card>
  );
}

function ScatterMatrix({ rows, numericCols }) {
  const cols = numericCols.slice(0, 4).map((c) => c.name);
  if (cols.length < 2) return null;
  const cell = 120;
  return (
    <Card style={{ padding: 20 }}>
      <div style={{ fontWeight: 700, fontSize: 14, color: T.ink, marginBottom: 4 }}>Scatter Matrix (Pair Plot)</div>
      <div style={{ fontSize: 11.5, color: T.faint, marginBottom: 12 }}>Combinações par a par entre as {cols.length} primeiras variáveis numéricas.</div>
      <div style={{ overflowX: "auto" }}>
        <div style={{ display: "grid", gridTemplateColumns: `80px repeat(${cols.length}, ${cell}px)`, gap: 4 }}>
          <div />
          {cols.map((c) => <div key={c} style={{ fontSize: 10.5, fontWeight: 700, color: T.ink, textAlign: "center", alignSelf: "end", paddingBottom: 4 }}>{c}</div>)}
          {cols.map((rowName, i) => (
            <React.Fragment key={rowName}>
              <div style={{ fontSize: 10.5, fontWeight: 700, color: T.ink, display: "flex", alignItems: "center", justifyContent: "flex-end", paddingRight: 6, writingMode: "vertical-rl", transform: "rotate(180deg)" }}>{rowName}</div>
              {cols.map((colName, j) => {
                if (i === j) {
                  const vals = rows.map((r) => r[rowName]).filter(isNum);
                  return (
                    <div key={j} style={{ width: cell, height: cell, background: T.tealSoft, borderRadius: 4 }}>
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={histogramBins(vals, 8)}>
                          <Bar dataKey="count" fill={T.teal} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  );
                }
                const pts = rows.filter((r) => isNum(r[colName]) && isNum(r[rowName])).map((r) => ({ x: r[colName], y: r[rowName] }));
                return (
                  <div key={j} style={{ width: cell, height: cell, border: `1px solid ${T.border}`, borderRadius: 4 }}>
                    <ResponsiveContainer width="100%" height="100%">
                      <ScatterChart margin={{ top: 4, right: 4, bottom: 4, left: 4 }}>
                        <XAxis dataKey="x" hide />
                        <YAxis dataKey="y" hide />
                        <Scatter data={pts} fill={T.teal} fillOpacity={0.5} />
                      </ScatterChart>
                    </ResponsiveContainer>
                  </div>
                );
              })}
            </React.Fragment>
          ))}
        </div>
      </div>
    </Card>
  );
}

function TimeSeriesDecomposition({ rows, dateCols, numericCols }) {
  const [dateColName, setDateColName] = useState(dateCols[0]?.name || "");
  const [numColName, setNumColName] = useState(numericCols[0]?.name || "");
  const [period, setPeriod] = useState("weekday");

  const series = useMemo(() => {
    if (!dateColName || !numColName) return [];
    return rows.filter((r) => r[dateColName] instanceof Date && !Number.isNaN(+r[dateColName]) && isNum(r[numColName]))
      .sort((a, b) => a[dateColName] - b[dateColName])
      .map((r) => ({ date: r[dateColName], value: r[numColName] }));
  }, [rows, dateColName, numColName]);

  const result = useMemo(() => {
    if (series.length < 8) return null;
    const t = series.map((_s, i) => i);
    const y = series.map((s) => s.value);
    const { a, b } = linearRegression(t, y);
    const trend = t.map((ti) => a + b * ti);
    const periodKey = (d) => period === "weekday" ? d.getDay() : d.getMonth();
    const nPeriods = period === "weekday" ? 7 : 12;
    const ratios = series.map((s, i) => (trend[i] !== 0 ? s.value / trend[i] : 1));
    const bucket = Array.from({ length: nPeriods }, () => []);
    series.forEach((s, i) => bucket[periodKey(s.date)].push(ratios[i]));
    const seasonalIndex = bucket.map((b2) => (b2.length ? mean(b2) : 1));
    const fitted = series.map((s, i) => trend[i] * seasonalIndex[periodKey(s.date)]);
    const chartData = series.map((s, i) => ({
      date: s.date.toLocaleDateString("pt-BR"), real: s.value, tendencia: trend[i], ajustado: fitted[i],
    }));
    const labels = period === "weekday"
      ? ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"]
      : ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
    const seasonalData = seasonalIndex.map((v, i) => ({ label: labels[i], index: Math.round(v * 100) / 100 }));
    return {
      chartData, seasonalData, trendDirection: b > 0 ? "crescente" : b < 0 ? "decrescente" : "estável",
      mapeVal: mape(y, fitted), rmseVal: rmse(y, fitted), slope: b,
    };
  }, [series, period]);

  if (!dateCols.length || !numericCols.length) return null;

  return (
    <Card style={{ padding: 20, marginBottom: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontWeight: 700, fontSize: 14, color: T.ink }}>Decomposição de série temporal (tendência + sazonalidade)</div>
        <div style={{ display: "flex", gap: 8 }}>
          <Select value={dateColName} onChange={setDateColName} options={dateCols.map((c) => c.name)} />
          <Select value={numColName} onChange={setNumColName} options={numericCols.map((c) => c.name)} />
          <Select value={period} onChange={setPeriod} options={[{ value: "weekday", label: "Ciclo semanal" }, { value: "month", label: "Ciclo mensal" }]} />
        </div>
      </div>
      {!result ? (
        <div style={{ color: T.sub, fontSize: 13, marginTop: 10 }}>Dados insuficientes para decompor a série (mínimo 8 pontos válidos).</div>
      ) : (
        <>
          <div style={{ fontSize: 11.5, color: T.faint, margin: "6px 0 14px" }}>
            Modelo multiplicativo (Y = Tendência × Sazonalidade). Tendência {result.trendDirection}, inclinação ≈ {fmt(result.slope, 3)} por observação.
          </div>
          <ResponsiveContainer width="100%" height={220}>
            <ComposedChart data={result.chartData}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.border} />
              <XAxis dataKey="date" tick={{ fontSize: 9 }} minTickGap={30} />
              <YAxis tick={{ fontSize: 10 }} />
              <Tooltip />
              <Legend wrapperStyle={{ fontSize: 11 }} />
              <Area type="monotone" dataKey="real" name="Real" stroke={T.teal} fill={T.tealSoft} />
              <Line type="monotone" dataKey="tendencia" name="Tendência" stroke={T.red} strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="ajustado" name="Ajustado (T×S)" stroke={T.amber} strokeWidth={2} dot={false} strokeDasharray="4 3" />
            </ComposedChart>
          </ResponsiveContainer>
          <div style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr", gap: 16, marginTop: 16 }}>
            <div>
              <div style={{ fontSize: 12, color: T.sub, marginBottom: 8, fontWeight: 600 }}>Índice sazonal por {period === "weekday" ? "dia da semana" : "mês"}</div>
              <ResponsiveContainer width="100%" height={160}>
                <BarChart data={result.seasonalData}>
                  <CartesianGrid strokeDasharray="3 3" stroke={T.border} />
                  <XAxis dataKey="label" tick={{ fontSize: 10 }} />
                  <YAxis tick={{ fontSize: 10 }} />
                  <ReferenceLine y={1} stroke={T.faint} strokeDasharray="3 3" />
                  <Tooltip />
                  <Bar dataKey="index" fill={T.blue} radius={[3, 3, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 10, justifyContent: "center" }}>
              <StatTile label="MAPE (erro %)" value={Number.isNaN(result.mapeVal) ? "—" : `${fmt(result.mapeVal, 1)}%`} tone={result.mapeVal > 20 ? "amber" : "green"} />
              <StatTile label="RMSE" value={Number.isNaN(result.rmseVal) ? "—" : fmt(result.rmseVal, 2)} />
            </div>
          </div>
        </>
      )}
    </Card>
  );
}

function CorrelationHeatmap({ columns, rows, method = "pearson" }) {
  const cols = columns.filter((c) => c.type === "numeric").map((c) => c.name);
  const data = useMemo(() => {
    const series = cols.map((name) => rows.map((r) => r[name]).filter(isNum));
    return cols.map((_r, i) => cols.map((_c, j) => {
      const x = rows.map((r) => r[cols[i]]).filter(isNum);
      const y = rows.map((r) => r[cols[j]]).filter(isNum);
      const n = Math.min(x.length, y.length);
      return method === "spearman" ? spearman(x.slice(0, n), y.slice(0, n)) : pearson(x.slice(0, n), y.slice(0, n));
    }));
  }, [cols.join(","), rows, method]);

  if (cols.length < 2) return <div style={{ color: T.sub, fontSize: 13 }}>É preciso de ao menos 2 colunas numéricas.</div>;
  const cell = Math.min(64, Math.floor(560 / cols.length));
  return (
    <div style={{ overflowX: "auto" }}>
      <div style={{ display: "grid", gridTemplateColumns: `120px repeat(${cols.length}, ${cell}px)`, gap: 2 }}>
        <div />
        {cols.map((c) => <div key={c} style={{ fontSize: 10, fontFamily: "'JetBrains Mono', monospace", color: T.sub, textAlign: "center", writingMode: "vertical-rl", height: 90, padding: 4 }}>{c}</div>)}
        {cols.map((rowName, i) => (
          <React.Fragment key={rowName}>
            <div style={{ fontSize: 11, fontFamily: "'JetBrains Mono', monospace", color: T.sub, display: "flex", alignItems: "center", justifyContent: "flex-end", paddingRight: 8 }}>{rowName}</div>
            {cols.map((_c, j) => (
              <div key={j} title={fmt(data[i][j], 3)} style={{ width: cell, height: cell, background: HEAT_SCALE(data[i][j] || 0), display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10, fontFamily: "'JetBrains Mono', monospace", color: Math.abs(data[i][j] || 0) > 0.55 ? "#fff" : T.ink, borderRadius: 3 }}>
                {fmt(data[i][j], 2)}
              </div>
            ))}
          </React.Fragment>
        ))}
      </div>
    </div>
  );
}

/* =========================================================================
   ASSOCIATION MATRIX — mede força de associação entre QUALQUER par de
   variáveis (num-num = Pearson |r|, cat-cat = Cramér's V, num-cat = √η²
   via ANOVA). Complementa a matriz de correlação, que só cobre numéricas.
   ========================================================================= */
function AssociationMatrix({ columns, rows }) {
  const cols = columns.filter((c) => c.type === "numeric" || c.type === "categorical").slice(0, 12);
  const data = useMemo(() => cols.map((ca) => cols.map((cb) => {
    if (ca.name === cb.name) return { value: 1, method: "—" };
    if (ca.type === "numeric" && cb.type === "numeric") {
      const x = rows.map((r) => r[ca.name]).filter(isNum), y = rows.map((r) => r[cb.name]).filter(isNum);
      const n = Math.min(x.length, y.length);
      return { value: Math.abs(pearson(x.slice(0, n), y.slice(0, n))), method: "r" };
    }
    if (ca.type === "categorical" && cb.type === "categorical") {
      const a = rows.map((r) => r[ca.name]).filter((v) => v !== null && v !== undefined);
      const b = rows.map((r) => r[cb.name]).filter((v) => v !== null && v !== undefined);
      const n = Math.min(a.length, b.length);
      const res = chiSquareTest(a.slice(0, n), b.slice(0, n));
      return { value: res.cramersV, method: "V" };
    }
    const numCol = ca.type === "numeric" ? ca.name : cb.name;
    const catCol = ca.type === "numeric" ? cb.name : ca.name;
    const catLevels = _.uniq(rows.map((r) => r[catCol]).filter((v) => v !== null && v !== undefined)).slice(0, 30);
    const groups = catLevels.map((lv) => rows.filter((r) => r[catCol] === lv).map((r) => r[numCol]).filter(isNum));
    const res = oneWayANOVA(groups);
    return { value: res ? Math.sqrt(Math.max(0, res.etaSquared)) : NaN, method: "√η²" };
  })), [cols.map((c) => c.name).join(","), rows]);

  if (cols.length < 2) return <div style={{ color: T.sub, fontSize: 13 }}>É preciso de ao menos 2 colunas numéricas ou categóricas.</div>;
  const cell = Math.min(58, Math.floor(560 / cols.length));
  return (
    <div>
      <div style={{ overflowX: "auto" }}>
        <div style={{ display: "grid", gridTemplateColumns: `130px repeat(${cols.length}, ${cell}px)`, gap: 2 }}>
          <div />
          {cols.map((c) => <div key={c.name} style={{ fontSize: 10, fontFamily: "'JetBrains Mono', monospace", color: T.sub, textAlign: "center", writingMode: "vertical-rl", height: 90, padding: 4 }}>{c.name}</div>)}
          {cols.map((rowCol, i) => (
            <React.Fragment key={rowCol.name}>
              <div style={{ fontSize: 11, fontFamily: "'JetBrains Mono', monospace", color: T.sub, display: "flex", alignItems: "center", justifyContent: "flex-end", paddingRight: 8 }}>{rowCol.name}</div>
              {cols.map((_c, j) => {
                const cellData = data[i][j];
                const v = Number.isNaN(cellData.value) ? 0 : cellData.value;
                return (
                  <div key={j} title={`${cellData.method}: ${fmt(cellData.value, 3)}`}
                    style={{ width: cell, height: cell, background: HEAT_SCALE(v), display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", fontSize: 9, fontFamily: "'JetBrains Mono', monospace", color: v > 0.55 ? "#fff" : T.ink, borderRadius: 3 }}>
                    <div>{fmt(cellData.value, 2)}</div>
                    <div style={{ fontSize: 7.5, opacity: 0.8 }}>{cellData.method}</div>
                  </div>
                );
              })}
            </React.Fragment>
          ))}
        </div>
      </div>
      <div style={{ marginTop: 10, fontSize: 11.5, color: T.faint }}>
        <b>r</b> = Pearson (numérica×numérica) · <b>V</b> = Cramér's V (categórica×categórica) · <b>√η²</b> = raiz do eta-quadrado da ANOVA (numérica×categórica). Todos em escala 0–1.
      </div>
    </div>
  );
}

/* =========================================================================
   STATIONARITY PANEL — teste de Dickey-Fuller (simplificado) + ACF/PACF
   ========================================================================= */
function StationarityPanel({ rows, dateCols, numericCols }) {
  const [dateColName, setDateColName] = useState(dateCols[0]?.name || "");
  const [numColName, setNumColName] = useState(numericCols[0]?.name || "");

  const series = useMemo(() => {
    if (!dateColName || !numColName) return [];
    return rows.filter((r) => r[dateColName] instanceof Date && !Number.isNaN(+r[dateColName]) && isNum(r[numColName]))
      .sort((a, b) => a[dateColName] - b[dateColName])
      .map((r) => r[numColName]);
  }, [rows, dateColName, numColName]);

  const adf = useMemo(() => (series.length >= 12 ? adfTest(series) : null), [series]);
  const maxLag = Math.min(20, Math.floor(series.length / 3));
  const acfVals = useMemo(() => (series.length > maxLag ? acfSeries(series, maxLag) : []), [series, maxLag]);
  const pacfVals = useMemo(() => (series.length > maxLag ? pacfSeries(series, maxLag) : []), [series, maxLag]);
  const bound = series.length ? 1.96 / Math.sqrt(series.length) : 0;
  const acfData = acfVals.map((v, k) => ({ lag: k, value: v }));
  const pacfData = pacfVals.map((v, k) => ({ lag: k, value: v }));

  if (!dateCols.length || !numericCols.length) return null;

  return (
    <Card style={{ padding: 20, marginBottom: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontWeight: 700, fontSize: 14, color: T.ink }}>Diagnósticos de série temporal (estacionariedade e autocorrelação)</div>
        <div style={{ display: "flex", gap: 8 }}>
          <Select value={dateColName} onChange={setDateColName} options={dateCols.map((c) => c.name)} />
          <Select value={numColName} onChange={setNumColName} options={numericCols.map((c) => c.name)} />
        </div>
      </div>
      {series.length < 12 ? (
        <div style={{ color: T.sub, fontSize: 13, marginTop: 10 }}>Dados insuficientes (mínimo 12 pontos válidos).</div>
      ) : (
        <>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 12, margin: "14px 0" }}>
            <StatTile label="Estatística ADF (t)" value={adf ? fmt(adf.tStat, 3) : "—"} />
            <StatTile label="Valor crítico 5%" value={adf ? fmt(adf.critical["5%"], 2) : "—"} />
            <StatTile label="Série estacionária?" value={adf ? (adf.stationary ? "Sim" : "Não") : "—"} tone={adf ? (adf.stationary ? "green" : "amber") : "neutral"} />
          </div>
          <div style={{ fontSize: 11.5, color: T.faint, marginBottom: 14 }}>
            Teste de Dickey-Fuller (regressão Δy = α + γ·y₋₁ + ε; estatística t de γ comparada a valores críticos assintóticos de MacKinnon — aproximação, não um p-valor exato). t abaixo do valor crítico de 5% sugere estacionariedade.
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
            <div>
              <div style={{ fontSize: 12, color: T.sub, marginBottom: 8, fontWeight: 600 }}>ACF — autocorrelação</div>
              <ResponsiveContainer width="100%" height={180}>
                <BarChart data={acfData}>
                  <CartesianGrid strokeDasharray="3 3" stroke={T.border} />
                  <XAxis dataKey="lag" tick={{ fontSize: 9 }} />
                  <YAxis domain={[-1, 1]} tick={{ fontSize: 10 }} />
                  <ReferenceLine y={bound} stroke={T.faint} strokeDasharray="3 3" />
                  <ReferenceLine y={-bound} stroke={T.faint} strokeDasharray="3 3" />
                  <Tooltip />
                  <Bar dataKey="value" fill={T.teal} radius={[2, 2, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div>
              <div style={{ fontSize: 12, color: T.sub, marginBottom: 8, fontWeight: 600 }}>PACF — autocorrelação parcial</div>
              <ResponsiveContainer width="100%" height={180}>
                <BarChart data={pacfData}>
                  <CartesianGrid strokeDasharray="3 3" stroke={T.border} />
                  <XAxis dataKey="lag" tick={{ fontSize: 9 }} />
                  <YAxis domain={[-1, 1]} tick={{ fontSize: 10 }} />
                  <ReferenceLine y={bound} stroke={T.faint} strokeDasharray="3 3" />
                  <ReferenceLine y={-bound} stroke={T.faint} strokeDasharray="3 3" />
                  <Tooltip />
                  <Bar dataKey="value" fill={T.blue} radius={[2, 2, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
          <div style={{ fontSize: 11.5, color: T.faint, marginTop: 10 }}>Linhas tracejadas = banda de significância aproximada (±1.96/√n).</div>
        </>
      )}
    </Card>
  );
}

/* =========================================================================
   EDA — DATA PROFILING (Dicionário de Dados) — reuses existing stat fns only
   ========================================================================= */
function detectBooleanLike(uniqueStrings) {
  if (uniqueStrings.length !== 2) return false;
  const set = new Set(uniqueStrings.map((v) => String(v).trim().toLowerCase()));
  const pairs = [["true", "false"], ["sim", "não"], ["sim", "nao"], ["yes", "no"], ["1", "0"], ["verdadeiro", "falso"], ["s", "n"]];
  return pairs.some((pair) => pair.every((p) => set.has(p)));
}

function buildDataDictionary(rows, columns) {
  const total = rows.length;
  return columns.map((c) => {
    const raw = rows.map((r) => r[c.name]);
    const nonNullVals = raw.filter((v) => v !== null && v !== undefined);
    const nonNull = nonNullVals.length;
    const nullCount = total - nonNull;
    const nullPct = total ? (nullCount / total) * 100 : 0;
    const uniqueStrings = _.uniq(nonNullVals.map((v) => (v instanceof Date ? v.toISOString() : String(v))));
    const uniqueCount = uniqueStrings.length;
    const uniquePct = nonNull ? (uniqueCount / nonNull) * 100 : 0;
    const duplicateValues = Math.max(0, nonNull - uniqueCount);
    const isBoolLike = c.type === "categorical" && detectBooleanLike(uniqueStrings);
    const detectedType = isBoolLike ? "Booleano" : c.type === "numeric" ? "Numérico" : c.type === "date" ? "Data/Hora" : c.type === "categorical" ? "Categórico" : "Texto";
    let cardinality = "Baixa";
    if (uniquePct > 90) cardinality = "Alta"; else if (uniquePct > 10) cardinality = "Média";

    let stats = {}, exampleValues = uniqueStrings.slice(0, 3), classification = "Texto Livre", outlierCount = 0, outlierPct = 0;

    if (c.type === "numeric") {
      const numVals = nonNullVals.filter(isNum);
      const q = quartiles(numVals);
      const out = outliersIQR(numVals);
      outlierCount = out.length;
      outlierPct = numVals.length ? (outlierCount / numVals.length) * 100 : 0;
      const md = mode(numVals);
      stats = {
        min: numVals.length ? Math.min(...numVals) : NaN, max: numVals.length ? Math.max(...numVals) : NaN,
        mean: mean(numVals), median: median(numVals), modeVal: md.length ? md[0] : NaN,
        std: std(numVals), variance: variance(numVals), q1: q.q1, q3: q.q3, iqr: q.q3 - q.q1,
        skew: skewness(numVals), kurt: kurtosisExcess(numVals),
      };
      classification = uniquePct > 95 && total > 20 ? "Identificador" : "Numérica";
      exampleValues = numVals.slice(0, 3).map((v) => fmt(v, 2));
    } else if (c.type === "date") {
      const dates = nonNullVals.filter((v) => v instanceof Date && !Number.isNaN(+v));
      const minD = dates.length ? new Date(Math.min(...dates)) : null;
      const maxD = dates.length ? new Date(Math.max(...dates)) : null;
      stats = { minDate: minD, maxDate: maxD, rangeDays: minD && maxD ? Math.round((maxD - minD) / 86400000) : NaN };
      classification = "Temporal";
      exampleValues = dates.slice(0, 3).map((v) => v.toLocaleDateString("pt-BR"));
    } else {
      const counts = _.countBy(nonNullVals.map(String));
      const sortedCounts = Object.entries(counts).sort((a, b) => b[1] - a[1]);
      stats = { modeVal: sortedCounts[0]?.[0], modeCount: sortedCounts[0]?.[1], topCategories: sortedCounts.slice(0, 10) };
      classification = isBoolLike ? "Booleana" : uniquePct > 95 && total > 20 ? "Identificador" : "Categórica";
    }

    let quality = "Excelente";
    if (nullPct > 20 || outlierPct > 15) quality = "Crítica";
    else if (nullPct > 5 || outlierPct > 5) quality = "Atenção";
    else if (nullPct > 0.5 || outlierPct > 0.5) quality = "Boa";

    return {
      name: c.name, detectedType, suggestedType: detectedType, classification, rawType: c.type, isBoolLike,
      totalRecords: total, nonNull, nullCount, nullPct, uniqueCount, uniquePct, duplicateValues, cardinality,
      exampleValues, quality, outlierCount, outlierPct, ...stats,
    };
  });
}

function qualityTone(q) { return q === "Excelente" ? "green" : q === "Boa" ? "teal" : q === "Atenção" ? "amber" : "red"; }

function DataDictionaryTable({ C, dict, onSelect }) {
  const [search, setSearch] = useState("");
  const [typeFilter, setTypeFilter] = useState("");
  const [qualityFilter, setQualityFilter] = useState("");
  const [nullFilter, setNullFilter] = useState("");
  const [sortKey, setSortKey] = useState("name");
  const [sortDir, setSortDir] = useState("asc");

  const filtered = useMemo(() => {
    let out = dict;
    if (search) out = out.filter((c) => c.name.toLowerCase().includes(search.toLowerCase()));
    if (typeFilter) out = out.filter((c) => c.rawType === typeFilter);
    if (qualityFilter) out = out.filter((c) => c.quality === qualityFilter);
    if (nullFilter) out = out.filter((c) => c.nullPct >= Number(nullFilter));
    return [...out].sort((a, b) => {
      const av = a[sortKey], bv = b[sortKey];
      const aBad = av === undefined || av === null || (typeof av === "number" && Number.isNaN(av));
      const bBad = bv === undefined || bv === null || (typeof bv === "number" && Number.isNaN(bv));
      if (aBad && bBad) return 0;
      if (aBad) return 1; if (bBad) return -1;
      if (typeof av === "string") return sortDir === "asc" ? av.localeCompare(bv) : bv.localeCompare(av);
      return sortDir === "asc" ? av - bv : bv - av;
    });
  }, [dict, search, typeFilter, qualityFilter, nullFilter, sortKey, sortDir]);

  const toggleSort = (key) => {
    if (sortKey === key) setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    else { setSortKey(key); setSortDir("asc"); }
  };

  const rowForExport = (c) => ({
    Coluna: c.name, Tipo: c.detectedType, Classificacao: c.classification, Total: c.totalRecords,
    NaoNulos: c.nonNull, Nulos: c.nullCount, PctNulos: fmt(c.nullPct, 2), Unicos: c.uniqueCount, PctUnicos: fmt(c.uniquePct, 2),
    Cardinalidade: c.cardinality, Min: c.min ?? "", Max: c.max ?? "", Media: c.mean ?? "", Mediana: c.median ?? "", DesvioPadrao: c.std ?? "",
    Outliers: c.outlierCount, PctOutliers: fmt(c.outlierPct, 2), Qualidade: c.quality,
  });
  const exportCsv = () => downloadBlob(toCSV(filtered.map(rowForExport)), "dicionario_dados.csv", "text/csv");
  const exportXlsx = () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filtered.map(rowForExport)), "Dicionario");
    XLSX.writeFile(wb, "dicionario_dados.xlsx");
  };

  const COLS = [
    ["name", "Coluna"], ["detectedType", "Tipo"], ["classification", "Classificação"],
    ["nonNull", "Não nulos"], ["nullCount", "Nulos"], ["nullPct", "% Nulos"],
    ["uniqueCount", "Únicos"], ["uniquePct", "% Únicos"], ["cardinality", "Cardinalidade"],
    ["min", "Mín"], ["max", "Máx"], ["mean", "Média"], ["std", "Desvio"],
    ["outlierCount", "Outliers"], ["quality", "Qualidade"],
  ];

  return (
    <Card style={{ padding: 20, background: C.panel, border: `1px solid ${C.border}` }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 10 }}>
        <div style={{ fontWeight: 700, fontSize: 14, color: C.ink }}>Dicionário de dados ({filtered.length} de {dict.length} colunas)</div>
        <div style={{ display: "flex", gap: 6 }}>
          <Btn variant="ghost" icon={Download} onClick={exportCsv}>CSV</Btn>
          <Btn variant="ghost" icon={FileSpreadsheet} onClick={exportXlsx}>Excel</Btn>
        </div>
      </div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 14 }}>
        <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Buscar coluna..."
          style={{ padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 12.5, minWidth: 160 }} />
        <Select value={typeFilter} onChange={setTypeFilter} placeholder="Todos os tipos" options={[{ value: "numeric", label: "Numérico" }, { value: "categorical", label: "Categórico" }, { value: "date", label: "Data/Hora" }, { value: "text", label: "Texto" }]} />
        <Select value={qualityFilter} onChange={setQualityFilter} placeholder="Toda qualidade" options={["Excelente", "Boa", "Atenção", "Crítica"]} />
        <Select value={nullFilter} onChange={setNullFilter} placeholder="Qualquer % nulos" options={[{ value: "5", label: "> 5% nulos" }, { value: "20", label: "> 20% nulos" }, { value: "50", label: "> 50% nulos" }]} />
      </div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, minWidth: 1000 }}>
          <thead>
            <tr style={{ textAlign: "left", color: C.sub, borderBottom: `1px solid ${C.border}` }}>
              {COLS.map(([key, label]) => (
                <th key={key} onClick={() => toggleSort(key)} style={{ padding: "8px 10px", cursor: "pointer", whiteSpace: "nowrap", userSelect: "none" }}>
                  {label} {sortKey === key ? (sortDir === "asc" ? "▲" : "▼") : ""}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((c) => (
              <tr key={c.name} onClick={() => onSelect(c)} style={{ borderBottom: `1px solid ${C.border}`, cursor: "pointer" }}
                onMouseEnter={(e) => (e.currentTarget.style.background = C.tealSoft)} onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}>
                <td style={{ padding: "8px 10px", fontWeight: 700, color: C.ink, fontFamily: "'JetBrains Mono', monospace" }}>{c.name}</td>
                <td style={{ padding: "8px 10px" }}><Pill tone={c.rawType === "numeric" ? "teal" : c.rawType === "date" ? "amber" : "neutral"}>{c.detectedType}</Pill></td>
                <td style={{ padding: "8px 10px", color: C.sub }}>{c.classification}</td>
                <td style={{ padding: "8px 10px", color: C.ink }}>{c.nonNull}</td>
                <td style={{ padding: "8px 10px", color: C.ink }}>{c.nullCount}</td>
                <td style={{ padding: "8px 10px", color: c.nullPct > 5 ? C.amber : C.sub }}>{fmt(c.nullPct, 1)}%</td>
                <td style={{ padding: "8px 10px", color: C.ink }}>{c.uniqueCount}</td>
                <td style={{ padding: "8px 10px", color: C.sub }}>{fmt(c.uniquePct, 1)}%</td>
                <td style={{ padding: "8px 10px", color: C.sub }}>{c.cardinality}</td>
                <td style={{ padding: "8px 10px", color: C.ink }}>{isNum(c.min) ? fmt(c.min, 1) : "—"}</td>
                <td style={{ padding: "8px 10px", color: C.ink }}>{isNum(c.max) ? fmt(c.max, 1) : "—"}</td>
                <td style={{ padding: "8px 10px", color: C.ink }}>{isNum(c.mean) ? fmt(c.mean, 1) : "—"}</td>
                <td style={{ padding: "8px 10px", color: C.ink }}>{isNum(c.std) ? fmt(c.std, 1) : "—"}</td>
                <td style={{ padding: "8px 10px", color: c.outlierCount > 0 ? C.amber : C.sub }}>{c.outlierCount || "—"}</td>
                <td style={{ padding: "8px 10px" }}><Pill tone={qualityTone(c.quality)}>{c.quality}</Pill></td>
              </tr>
            ))}
          </tbody>
        </table>
        {!filtered.length && <div style={{ color: C.sub, fontSize: 12.5, padding: 16, textAlign: "center" }}>Nenhuma coluna corresponde aos filtros.</div>}
      </div>
      <div style={{ fontSize: 11, color: C.faint, marginTop: 10 }}>Clique em uma linha para abrir o perfil detalhado da variável.</div>
    </Card>
  );
}

function VariableProfileModal({ C, profile, rows, onClose, notes, onAddNote, onRemoveNote }) {
  const [noteDraft, setNoteDraft] = useState("");
  const myNotes = useMemo(() => (notes || []).filter((n) => n.variable === profile?.name), [notes, profile?.name]);
  if (!profile) return null;
  const vals = profile.rawType === "numeric" ? rows.map((r) => r[profile.name]).filter(isNum) : [];
  const monthlyDist = useMemo(() => {
    if (profile.rawType !== "date") return [];
    const counts = {};
    rows.forEach((r) => {
      const v = r[profile.name];
      if (v instanceof Date && !Number.isNaN(+v)) {
        const key = `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}`;
        counts[key] = (counts[key] || 0) + 1;
      }
    });
    return Object.entries(counts).sort((a, b) => a[0].localeCompare(b[0])).map(([label, count]) => ({ label, count }));
  }, [profile, rows]);

  return (
    <div onClick={onClose} style={{ position: "fixed", inset: 0, background: "rgba(10,12,16,0.55)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: C.panel, borderRadius: 16, padding: 24, width: "min(760px, 100%)", maxHeight: "88vh", overflow: "auto", boxShadow: "0 20px 60px rgba(0,0,0,0.35)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16 }}>
          <div>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 19, fontWeight: 700, color: C.ink }}>{profile.name}</div>
            <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
              <Pill tone={profile.rawType === "numeric" ? "teal" : profile.rawType === "date" ? "amber" : "neutral"}>{profile.detectedType}</Pill>
              <Pill tone="neutral">{profile.classification}</Pill>
              <Pill tone={qualityTone(profile.quality)}>{profile.quality}</Pill>
            </div>
          </div>
          <button onClick={onClose} style={{ background: "transparent", border: `1px solid ${C.border}`, borderRadius: 7, width: 30, height: 30, cursor: "pointer" }}>✕</button>
        </div>

        <div style={{ fontSize: 11.5, fontWeight: 700, color: C.sub, textTransform: "uppercase", marginBottom: 8 }}>Informações gerais</div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 18 }}>
          <KpiCard C={C} label="Total de registros" value={profile.totalRecords} />
          <KpiCard C={C} label="Valores únicos" value={profile.uniqueCount} sub={`${fmt(profile.uniquePct, 1)}%`} />
          <KpiCard C={C} label="Ausentes" value={profile.nullCount} tone={profile.nullPct > 5 ? "amber" : "green"} sub={`${fmt(profile.nullPct, 1)}%`} />
          <KpiCard C={C} label="% preenchimento" value={`${fmt(100 - profile.nullPct, 1)}%`} />
          <KpiCard C={C} label="Cardinalidade" value={profile.cardinality} />
        </div>

        {profile.rawType === "numeric" && (
          <>
            <div style={{ fontSize: 11.5, fontWeight: 700, color: C.sub, textTransform: "uppercase", marginBottom: 8 }}>Estatísticas</div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 16 }}>
              <KpiCard C={C} label="Média" value={fmt(profile.mean, 2)} /><KpiCard C={C} label="Mediana" value={fmt(profile.median, 2)} />
              <KpiCard C={C} label="Moda" value={Number.isNaN(profile.modeVal) ? "—" : fmt(profile.modeVal, 2)} />
              <KpiCard C={C} label="Mínimo" value={fmt(profile.min, 2)} /><KpiCard C={C} label="Máximo" value={fmt(profile.max, 2)} />
              <KpiCard C={C} label="Desvio padrão" value={fmt(profile.std, 2)} /><KpiCard C={C} label="Variância" value={fmt(profile.variance, 2)} />
              <KpiCard C={C} label="Q1 / Q3" value={`${fmt(profile.q1, 1)} / ${fmt(profile.q3, 1)}`} /><KpiCard C={C} label="IQR" value={fmt(profile.iqr, 2)} />
              <KpiCard C={C} label="Assimetria" value={fmt(profile.skew, 2)} /><KpiCard C={C} label="Curtose" value={fmt(profile.kurt, 2)} />
              <KpiCard C={C} label="Outliers" value={profile.outlierCount} sub={`${fmt(profile.outlierPct, 1)}%`} tone={profile.outlierCount ? "amber" : "green"} />
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr", gap: 16, marginBottom: 8 }}>
              <div><div style={{ fontSize: 11.5, color: C.sub, marginBottom: 6, fontWeight: 600 }}>Histograma</div>
                <ResponsiveContainer width="100%" height={160}><BarChart data={histogramBins(vals, 10)}><Bar dataKey="count" fill={C.teal} radius={[3, 3, 0, 0]} /></BarChart></ResponsiveContainer>
              </div>
              <div><div style={{ fontSize: 11.5, color: C.sub, marginBottom: 6, fontWeight: 600 }}>Boxplot</div><BoxPlotSVG vals={vals} height={160} /></div>
            </div>
          </>
        )}

        {(profile.rawType === "categorical" || profile.rawType === "text") && profile.topCategories && (
          <>
            <div style={{ fontSize: 11.5, fontWeight: 700, color: C.sub, textTransform: "uppercase", marginBottom: 8 }}>Top 10 categorias</div>
            <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 10 }}>
              {profile.topCategories.map(([val, count], i) => (
                <div key={i} style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <div style={{ width: 130, fontSize: 11.5, color: C.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{val}</div>
                  <div style={{ flex: 1, background: C.border, borderRadius: 4, height: 8 }}><div style={{ width: `${(count / profile.topCategories[0][1]) * 100}%`, background: C.teal, height: 8, borderRadius: 4 }} /></div>
                  <div style={{ width: 90, fontSize: 11, color: C.sub, fontFamily: "'JetBrains Mono', monospace", textAlign: "right" }}>{count} ({fmt((count / profile.nonNull) * 100, 1)}%)</div>
                </div>
              ))}
            </div>
          </>
        )}

        {profile.rawType === "date" && (
          <>
            <div style={{ fontSize: 11.5, fontWeight: 700, color: C.sub, textTransform: "uppercase", marginBottom: 8 }}>Cobertura temporal</div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 16 }}>
              <KpiCard C={C} label="Data mínima" value={profile.minDate ? profile.minDate.toLocaleDateString("pt-BR") : "—"} />
              <KpiCard C={C} label="Data máxima" value={profile.maxDate ? profile.maxDate.toLocaleDateString("pt-BR") : "—"} />
              <KpiCard C={C} label="Período coberto" value={Number.isNaN(profile.rangeDays) ? "—" : `${profile.rangeDays} dias`} />
            </div>
            {monthlyDist.length > 0 && (
              <ResponsiveContainer width="100%" height={160}>
                <BarChart data={monthlyDist}><XAxis dataKey="label" tick={{ fontSize: 9 }} /><YAxis tick={{ fontSize: 10 }} /><Tooltip /><Bar dataKey="count" fill={C.amber} radius={[3, 3, 0, 0]} /></BarChart>
              </ResponsiveContainer>
            )}
          </>
        )}

        <div style={{ fontSize: 11.5, fontWeight: 700, color: C.sub, textTransform: "uppercase", marginTop: 16, marginBottom: 6 }}>Exemplos de valores</div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {profile.exampleValues.map((v, i) => <Pill key={i}>{String(v)}</Pill>)}
        </div>

        {onAddNote && (
          <>
            <div style={{ fontSize: 11.5, fontWeight: 700, color: C.sub, textTransform: "uppercase", marginTop: 18, marginBottom: 8 }}>Anotações desta variável</div>
            {myNotes.length > 0 && (
              <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 10 }}>
                {myNotes.map((n) => (
                  <div key={n.id} style={{ display: "flex", justifyContent: "space-between", gap: 8, fontSize: 12.5, color: C.ink, background: C.bg, borderRadius: 8, padding: "8px 10px" }}>
                    <div>{n.text}<div style={{ fontSize: 10.5, color: C.faint, marginTop: 2 }}>{new Date(n.createdAt).toLocaleString("pt-BR")}</div></div>
                    <button onClick={() => onRemoveNote(n.id)} style={{ border: "none", background: "transparent", cursor: "pointer", color: C.red, flexShrink: 0 }}><Trash2 size={13} /></button>
                  </div>
                ))}
              </div>
            )}
            <div style={{ display: "flex", gap: 8 }}>
              <input value={noteDraft} onChange={(e) => setNoteDraft(e.target.value)} placeholder="Por que este outlier foi mantido? Alguma decisão a documentar?"
                style={{ flex: 1, padding: "8px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 12.5 }} />
              <Btn icon={Plus} onClick={() => { if (noteDraft.trim()) { onAddNote(profile.name, noteDraft.trim()); setNoteDraft(""); } }}>Anotar</Btn>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

function MissingDataPanel({ C, dict }) {
  const withMissing = [...dict].filter((c) => c.nullCount > 0).sort((a, b) => b.nullPct - a.nullPct);
  return (
    <Card style={{ padding: 20, background: C.panel, border: `1px solid ${C.border}` }}>
      <div style={{ fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 12 }}>Valores ausentes por coluna</div>
      {withMissing.length ? (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {withMissing.map((c, i) => (
            <div key={i} style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <div style={{ width: 140, fontSize: 11.5, color: C.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{c.name}</div>
              <div style={{ flex: 1, background: C.border, borderRadius: 4, height: 9 }}>
                <div style={{ width: `${Math.min(100, c.nullPct)}%`, height: "100%", borderRadius: 4, background: c.nullPct > 20 ? C.red : c.nullPct > 5 ? C.amber : C.teal }} />
              </div>
              <div style={{ width: 90, fontSize: 11, color: C.sub, fontFamily: "'JetBrains Mono', monospace", textAlign: "right" }}>{c.nullCount} ({fmt(c.nullPct, 1)}%)</div>
            </div>
          ))}
        </div>
      ) : <Callout C={C} tone="success" icon={CheckCircle2}>Nenhum valor ausente identificado na base.</Callout>}
      {withMissing.some((c) => c.nullPct > 30) && (
        <div style={{ marginTop: 14 }}>
          <Callout C={C} tone="warning">Colunas com mais de 30% de ausência ({withMissing.filter((c) => c.nullPct > 30).map((c) => c.name).join(", ")}) são candidatas a exclusão ou exigem estratégia de imputação específica.</Callout>
        </div>
      )}
    </Card>
  );
}

function OutliersPanel({ C, dict, rows }) {
  const withOutliers = dict.filter((c) => c.rawType === "numeric" && c.outlierCount > 0).sort((a, b) => b.outlierCount - a.outlierCount);
  return (
    <Card style={{ padding: 20, background: C.panel, border: `1px solid ${C.border}` }}>
      <div style={{ fontWeight: 700, fontSize: 14, color: C.ink, marginBottom: 12 }}>Outliers por variável (método IQR, 1.5×)</div>
      {withOutliers.length ? (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 14 }}>
          {withOutliers.slice(0, 6).map((c, i) => (
            <div key={i} style={{ border: `1px solid ${C.border}`, borderRadius: 10, padding: 12 }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                <span style={{ fontWeight: 700, fontSize: 12.5, color: C.ink }}>{c.name}</span>
                <Pill tone="amber">{c.outlierCount} ({fmt(c.outlierPct, 1)}%)</Pill>
              </div>
              <BoxPlotSVG vals={rows.map((r) => r[c.name]).filter(isNum)} height={140} />
            </div>
          ))}
        </div>
      ) : <Callout C={C} tone="success" icon={CheckCircle2}>Nenhum outlier relevante identificado.</Callout>}
    </Card>
  );
}

function ProfilingInsightsPanel({ C, dict, d }) {
  const insights = [];
  const pk = dict.find((c) => c.nullCount === 0 && c.uniquePct >= 99.5 && c.totalRecords > 5);
  if (pk) insights.push({ tone: "teal", text: `"${pk.name}" tem 100% de preenchimento e ${fmt(pk.uniquePct, 1)}% de unicidade — possível chave primária.` });
  const fkCandidates = dict.filter((c) => c !== pk && /(^id$|_id$|^id_|codigo|code)/i.test(c.name) && c.cardinality !== "Alta");
  if (fkCandidates.length) insights.push({ tone: "neutral", text: `Possíveis chaves estrangeiras / códigos de referência: ${fkCandidates.slice(0, 4).map((c) => c.name).join(", ")}.` });
  const constants = dict.filter((c) => c.uniqueCount <= 1);
  if (constants.length) insights.push({ tone: "amber", text: `Campos constantes (podem ser removidos): ${constants.map((c) => c.name).join(", ")}.` });
  const highCardText = dict.filter((c) => c.rawType === "text" && c.uniquePct > 90);
  if (highCardText.length) insights.push({ tone: "neutral", text: `Alta cardinalidade em texto livre (prováveis identificadores): ${highCardText.slice(0, 4).map((c) => c.name).join(", ")}.` });
  const highMissing = dict.filter((c) => c.nullPct > 30);
  if (highMissing.length) insights.push({ tone: "amber", text: `Colunas com mais de 30% de ausência: ${highMissing.map((c) => c.name).join(", ")} — considere remover ou tratar.` });
  const strongCorr = d.allCorrPairs.filter((p) => Math.abs(p.r) > 0.9);
  if (strongCorr.length) insights.push({ tone: "amber", text: `Colunas altamente correlacionadas (redundância possível): ${strongCorr.slice(0, 3).map((p) => `${p.a}×${p.b} (r=${fmt(p.r, 2)})`).join("; ")}.` });
  if (d.duplicates > 0) insights.push({ tone: "amber", text: `${d.duplicates} linha(s) duplicada(s) exatas na base — considere deduplicar antes de agregações.` });
  const criticalCols = dict.filter((c) => c.quality === "Crítica");
  if (criticalCols.length) insights.push({ tone: "amber", text: `Colunas com qualidade crítica: ${criticalCols.map((c) => c.name).join(", ")}.` });
  if (!insights.length) insights.push({ tone: "green", text: "Nenhum problema estrutural relevante identificado — base limpa e pronta para análise." });

  return (
    <Card style={{ padding: 20, background: C.panel, border: `1px solid ${C.border}` }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}><Sparkles size={15} color={C.teal} /><div style={{ fontWeight: 700, fontSize: 14, color: C.ink }}>Insights automáticos de Data Profiling</div></div>
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {insights.map((ins, i) => <Callout key={i} C={C} tone={ins.tone === "amber" ? "warning" : ins.tone === "green" ? "success" : "info"}>{ins.text}</Callout>)}
      </div>
    </Card>
  );
}

function EDATab({ rows, columns, notes, onAddNote, onRemoveNote }) {
  const numericCols = columns.filter((c) => c.type === "numeric");
  const catCols = columns.filter((c) => c.type === "categorical");
  const dateCols = columns.filter((c) => c.type === "date");
  const [col1, setCol1] = useState(numericCols[0]?.name || "");
  const [col2, setCol2] = useState(numericCols[1]?.name || numericCols[0]?.name || "");
  const [catCol, setCatCol] = useState(catCols[0]?.name || "");
  const [corrMethod, setCorrMethod] = useState("pearson");
  const [edaView, setEdaView] = useState("perfil");
  const [selectedVar, setSelectedVar] = useState(null);

  const dict = useMemo(() => buildDataDictionary(rows, columns), [rows, columns]);
  const dashD = useMemo(() => computeDashboard(rows, columns), [rows, columns]);

  const vals1 = useMemo(() => rows.map((r) => r[col1]).filter(isNum), [rows, col1]);
  const catCounts = useMemo(() => {
    if (!catCol) return [];
    const counts = _.countBy(rows.map((r) => r[catCol]).filter((v) => v !== null && v !== undefined));
    return Object.entries(counts).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value).slice(0, 12);
  }, [rows, catCol]);

  const scatterData = useMemo(() => {
    if (!col1 || !col2) return [];
    return rows.filter((r) => isNum(r[col1]) && isNum(r[col2])).map((r) => ({ x: r[col1], y: r[col2] }));
  }, [rows, col1, col2]);

  const timeSeries = useMemo(() => {
    if (!dateCols.length || !numericCols.length) return null;
    const dc = dateCols[0].name, nc = numericCols[0].name;
    return rows.filter((r) => r[dc] instanceof Date && !Number.isNaN(+r[dc]) && isNum(r[nc]))
      .sort((a, b) => a[dc] - b[dc])
      .map((r) => ({ date: r[dc].toLocaleDateString("pt-BR"), value: r[nc] }));
  }, [rows, dateCols, numericCols]);

  const PIE_COLORS = [T.teal, T.amber, T.blue, T.green, T.red, "#7C6FDB", "#C77DFF", "#4C956C", "#E07A5F", "#3D5A80", "#F2CC8F", "#8D99AE"];

  return (
    <div>
      <SectionTitle eyebrow="Exploração de dados" title="EDA — Análise Exploratória e Data Profiling" />

      <div style={{ display: "flex", gap: 8, marginBottom: 18, flexWrap: "wrap" }}>
        {[{ id: "perfil", label: "Perfil de Dados", icon: Layers }, { id: "visualizacoes", label: "Visualizações", icon: BarChart3 }].map((s) => {
          const active = edaView === s.id;
          return (
            <button key={s.id} onClick={() => setEdaView(s.id)}
              style={{ display: "flex", alignItems: "center", gap: 8, padding: "9px 15px", borderRadius: 10, border: `1px solid ${active ? T.teal : T.border}`, background: active ? T.tealSoft : "#fff", cursor: "pointer" }}>
              <s.icon size={15} color={active ? T.teal : T.faint} />
              <span style={{ fontWeight: 700, fontSize: 12.5, color: active ? T.tealDark : T.ink }}>{s.label}</span>
            </button>
          );
        })}
      </div>

      {edaView === "perfil" && (
        <div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 16 }}>
            <StatTile label="Registros" value={rows.length.toLocaleString("pt-BR")} />
            <StatTile label="Colunas" value={columns.length} />
            <StatTile label="Numéricas" value={numericCols.length} tone="teal" />
            <StatTile label="Categóricas / Booleanas" value={dict.filter((c) => c.rawType === "categorical").length} />
            <StatTile label="Data/Hora" value={dateCols.length} tone="amber" />
            <StatTile label="Texto" value={dict.filter((c) => c.rawType === "text").length} />
            <StatTile label="% Ausentes" value={`${fmt(dashD.missingPct, 1)}%`} tone={dashD.missingPct > 5 ? "amber" : "neutral"} />
            <StatTile label="Duplicados" value={dashD.duplicates} tone={dashD.duplicates > 0 ? "amber" : "green"} />
            <StatTile label="Memória (aprox.)" value={`${(dashD.memBytes / 1024).toFixed(0)} KB`} />
            <StatTile label="Score de qualidade" value={`${Math.round(dashD.quality)}/100`} tone={dashD.quality >= 80 ? "green" : dashD.quality >= 55 ? "amber" : "red"} />
          </div>

          <ReportDivider C={T} label="Dicionário de dados" />
          <div style={{ marginBottom: 16 }}>
            <DataDictionaryTable C={T} dict={dict} onSelect={setSelectedVar} />
          </div>

          <ReportDivider C={T} label="Qualidade dos dados" />
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(340px, 1fr))", gap: 16, marginBottom: 16 }}>
            <MissingDataPanel C={T} dict={dict} />
            <OutliersPanel C={T} dict={dict} rows={rows} />
          </div>

          <div style={{ marginBottom: 16 }}>
            <ProfilingInsightsPanel C={T} dict={dict} d={dashD} />
          </div>
        </div>
      )}

      {edaView === "visualizacoes" && (
        <>
      {numericCols.length > 0 && (
        <Card style={{ padding: 20, marginBottom: 16 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 10 }}>
            <div style={{ fontWeight: 700, fontSize: 14, color: T.ink }}>Distribuição de variável numérica</div>
            <Select value={col1} onChange={setCol1} options={numericCols.map((c) => c.name)} />
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr", gap: 20 }}>
            <div>
              <div style={{ fontSize: 12, color: T.sub, marginBottom: 8, fontWeight: 600 }}>Histograma</div>
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={histogramBins(vals1)}>
                  <CartesianGrid strokeDasharray="3 3" stroke={T.border} />
                  <XAxis dataKey="label" tick={{ fontSize: 9.5 }} interval={1} />
                  <YAxis tick={{ fontSize: 10 }} />
                  <Tooltip />
                  <Bar dataKey="count" fill={T.teal} radius={[3, 3, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div>
              <div style={{ fontSize: 12, color: T.sub, marginBottom: 8, fontWeight: 600 }}>Boxplot (regra 1.5×IQR)</div>
              <BoxPlotSVG vals={vals1} />
            </div>
          </div>
        </Card>
      )}

      <div style={{ display: "grid", gridTemplateColumns: catCols.length ? "1fr 1fr" : "1fr", gap: 16, marginBottom: 16 }}>
        {numericCols.length > 1 && (
          <Card style={{ padding: 20 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
              <div style={{ fontWeight: 700, fontSize: 14, color: T.ink }}>Dispersão (Scatter)</div>
              <div style={{ display: "flex", gap: 8 }}>
                <Select value={col1} onChange={setCol1} options={numericCols.map((c) => c.name)} />
                <Select value={col2} onChange={setCol2} options={numericCols.map((c) => c.name)} />
              </div>
            </div>
            <ResponsiveContainer width="100%" height={240}>
              <ScatterChart>
                <CartesianGrid strokeDasharray="3 3" stroke={T.border} />
                <XAxis dataKey="x" name={col1} tick={{ fontSize: 10 }} />
                <YAxis dataKey="y" name={col2} tick={{ fontSize: 10 }} />
                <Tooltip cursor={{ strokeDasharray: "3 3" }} />
                <Scatter data={scatterData} fill={T.teal} fillOpacity={0.6} />
              </ScatterChart>
            </ResponsiveContainer>
          </Card>
        )}
        {catCols.length > 0 && (
          <Card style={{ padding: 20 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
              <div style={{ fontWeight: 700, fontSize: 14, color: T.ink }}>Categorias mais frequentes</div>
              <Select value={catCol} onChange={setCatCol} options={catCols.map((c) => c.name)} />
            </div>
            <ResponsiveContainer width="100%" height={240}>
              <PieChart>
                <Pie data={catCounts} dataKey="value" nameKey="name" outerRadius={85} label={{ fontSize: 10 }}>
                  {catCounts.map((_e, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                </Pie>
                <Tooltip />
              </PieChart>
            </ResponsiveContainer>
          </Card>
        )}
      </div>

      {catCols.length > 0 && (
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
          <ParetoCard rows={rows} catCols={catCols} catCol={catCol} setCatCol={setCatCol} />
          <TreemapCard rows={rows} catCols={catCols} catCol={catCol} setCatCol={setCatCol} />
        </div>
      )}

      <TimeSeriesDecomposition rows={rows} dateCols={dateCols} numericCols={numericCols} />

      <StationarityPanel rows={rows} dateCols={dateCols} numericCols={numericCols} />

      {timeSeries && timeSeries.length > 1 && (
        <Card style={{ padding: 20, marginBottom: 16 }}>
          <div style={{ fontWeight: 700, fontSize: 14, color: T.ink, marginBottom: 12 }}>Série temporal — {numericCols[0].name} por {dateCols[0].name}</div>
          <ResponsiveContainer width="100%" height={220}>
            <AreaChart data={timeSeries}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.border} />
              <XAxis dataKey="date" tick={{ fontSize: 9.5 }} minTickGap={30} />
              <YAxis tick={{ fontSize: 10 }} />
              <Tooltip />
              <Area type="monotone" dataKey="value" stroke={T.teal} fill={T.tealSoft} />
            </AreaChart>
          </ResponsiveContainer>
        </Card>
      )}

      {numericCols.length > 1 && (
        <div style={{ marginBottom: 16 }}>
          <ScatterMatrix rows={rows} numericCols={numericCols} />
        </div>
      )}

      <Card style={{ padding: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
          <div style={{ fontWeight: 700, fontSize: 14, color: T.ink }}>Matriz de correlação</div>
          <Select value={corrMethod} onChange={setCorrMethod} options={[{ value: "pearson", label: "Pearson" }, { value: "spearman", label: "Spearman" }]} />
        </div>
        <CorrelationHeatmap columns={columns} rows={rows} method={corrMethod} />
      </Card>

      <div style={{ marginTop: 16, marginBottom: 16 }}>
        <Card style={{ padding: 20 }}>
          <div style={{ fontWeight: 700, fontSize: 14, color: T.ink, marginBottom: 4 }}>Matriz de associação (todas as variáveis)</div>
          <div style={{ fontSize: 12, color: T.sub, marginBottom: 14 }}>Combina numéricas e categóricas num único mapa de força de associação — cobre relações que a correlação de Pearson/Spearman não captura.</div>
          <AssociationMatrix columns={columns} rows={rows} />
        </Card>
      </div>

      <div style={{ marginTop: 14, fontSize: 12, color: T.faint, display: "flex", gap: 6, alignItems: "center" }}>
        <Info size={13} /> Sunburst e mapas geográficos (lat/long) seguem no roadmap da próxima fase.
      </div>
        </>
      )}

      {selectedVar && <VariableProfileModal C={T} profile={selectedVar} rows={rows} onClose={() => setSelectedVar(null)} notes={notes} onAddNote={onAddNote} onRemoveNote={onRemoveNote} />}
    </div>
  );
}

/* =========================================================================
   DESCRIPTIVE STATISTICS TAB
   ========================================================================= */
function DescriptiveTab({ rows, columns }) {
  const numericCols = columns.filter((c) => c.type === "numeric");
  const stats = useMemo(() => numericCols.map((c) => {
    const vals = rows.map((r) => r[c.name]).filter(isNum);
    const m = mode(vals);
    return {
      name: c.name, n: vals.length, mean: mean(vals), median: median(vals),
      mode: m.length ? m.map((v) => fmt(v, 2)).join(", ") : "—",
      variance: variance(vals), std: std(vals), cv: cv(vals),
      skew: skewness(vals), kurt: kurtosisExcess(vals),
      min: Math.min(...vals), max: Math.max(...vals),
      ...quartiles(vals), iqr: iqr(vals), range: range(vals),
    };
  }), [rows, numericCols.map((c) => c.name).join(",")]);

  const cols = [
    ["n", "N"], ["mean", "Média"], ["median", "Mediana"], ["mode", "Moda"],
    ["std", "Desvio padrão"], ["variance", "Variância"], ["cv", "CV (%)"],
    ["skew", "Assimetria"], ["kurt", "Curtose (exc.)"], ["min", "Mín"],
    ["q1", "Q1"], ["q2", "Q2"], ["q3", "Q3"], ["max", "Máx"], ["iqr", "IQR"], ["range", "Amplitude"],
  ];

  return (
    <div>
      <SectionTitle eyebrow="Estatística" title="Estatística Descritiva"
        right={<Btn variant="ghost" icon={Download} onClick={() => downloadBlob(toCSV(stats), "estatistica_descritiva.csv", "text/csv")}>Exportar CSV</Btn>} />

      <Card style={{ padding: 20, marginBottom: 16, overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, minWidth: 900 }}>
          <thead>
            <tr style={{ borderBottom: `1px solid ${T.border}`, textAlign: "left" }}>
              <th style={{ padding: "8px 10px", color: T.sub, position: "sticky", left: 0, background: T.panel }}>Variável</th>
              {cols.map(([k, label]) => <th key={k} style={{ padding: "8px 10px", color: T.sub, whiteSpace: "nowrap" }}>{label}</th>)}
            </tr>
          </thead>
          <tbody>
            {stats.map((s) => (
              <tr key={s.name} style={{ borderBottom: `1px solid ${T.border}` }}>
                <td style={{ padding: "8px 10px", fontWeight: 700, color: T.ink, fontFamily: "'JetBrains Mono', monospace", position: "sticky", left: 0, background: T.panel }}>{s.name}</td>
                {cols.map(([k]) => (
                  <td key={k} style={{ padding: "8px 10px", color: T.ink, fontFamily: "'JetBrains Mono', monospace", whiteSpace: "nowrap" }}>
                    {typeof s[k] === "string" ? s[k] : fmt(s[k], 2)}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
        {!stats.length && <div style={{ color: T.sub, fontSize: 13 }}>Nenhuma coluna numérica encontrada.</div>}
      </Card>

      <Card style={{ padding: 20 }}>
        <div style={{ fontWeight: 700, fontSize: 14, color: T.ink, marginBottom: 6 }}>Como interpretar</div>
        <ul style={{ fontSize: 12.5, color: T.sub, lineHeight: 1.8, margin: 0, paddingLeft: 18 }}>
          <li><b>CV (coeficiente de variação):</b> dispersão relativa — acima de ~30% indica alta variabilidade.</li>
          <li><b>Assimetria (skewness):</b> perto de 0 é simétrica; positiva = cauda à direita; negativa = cauda à esquerda.</li>
          <li><b>Curtose (excesso):</b> perto de 0 é semelhante à normal; positiva = caudas mais pesadas (mais outliers prováveis).</li>
          <li><b>IQR:</b> Q3 − Q1, usado para detectar outliers pela regra 1.5×IQR.</li>
        </ul>
      </Card>
    </div>
  );
}

/* =========================================================================
   HYPOTHESIS TESTS TAB — t-test, ANOVA, qui-quadrado, correlação com p-valor
   ========================================================================= */
function conclusionSentence(p, alpha = 0.05) {
  if (Number.isNaN(p)) return "Dados insuficientes para concluir.";
  return p < alpha
    ? `p ${fmtP(p)} < ${alpha} → há evidência estatística de diferença/associação (rejeita H₀).`
    : `p ${fmtP(p)} ≥ ${alpha} → não há evidência estatística suficiente (não rejeita H₀).`;
}

function HypothesisTestsTab({ rows, columns, onLogAction }) {
  const numericCols = columns.filter((c) => c.type === "numeric");
  const catCols = columns.filter((c) => c.type === "categorical");
  const MODES = [
    { id: "ttest", label: "2 grupos (t-test)" },
    { id: "anova", label: "Vários grupos (ANOVA)" },
    { id: "chi2", label: "Categórica × categórica (χ²)" },
    { id: "corr", label: "Correlação com significância" },
  ];
  const [mode, setMode] = useState("ttest");
  const [catCol, setCatCol] = useState(catCols[0]?.name || "");
  const [numCol, setNumCol] = useState(numericCols[0]?.name || "");
  const [catCol2, setCatCol2] = useState(catCols[1]?.name || catCols[0]?.name || "");
  const [numCol2, setNumCol2] = useState(numericCols[1]?.name || numericCols[0]?.name || "");
  const [corrMethod, setCorrMethod] = useState("pearson");
  const [levelA, setLevelA] = useState("");
  const [levelB, setLevelB] = useState("");
  const [correction, setCorrection] = useState("fdr");
  const [batchCat, setBatchCat] = useState(catCols[0]?.name || "");

  const levels = useMemo(() => {
    if (!catCol) return [];
    return _.uniq(rows.map((r) => r[catCol]).filter((v) => v !== null && v !== undefined)).slice(0, 40);
  }, [rows, catCol]);

  useEffect(() => {
    if (levels.length) { setLevelA(levels[0]); setLevelB(levels.length > 1 ? levels[1] : levels[0]); }
  }, [catCol, levels.join("|")]);

  const ttestResult = useMemo(() => {
    if (mode !== "ttest" || !catCol || !numCol || !levelA || !levelB || levelA === levelB) return null;
    const a = rows.filter((r) => r[catCol] === levelA).map((r) => r[numCol]).filter(isNum);
    const b = rows.filter((r) => r[catCol] === levelB).map((r) => r[numCol]).filter(isNum);
    return twoSampleTTest(a, b);
  }, [mode, rows, catCol, numCol, levelA, levelB]);

  const anovaResult = useMemo(() => {
    if (mode !== "anova" || !catCol || !numCol) return null;
    const groups = levels.map((lv) => rows.filter((r) => r[catCol] === lv).map((r) => r[numCol]).filter(isNum));
    return oneWayANOVA(groups);
  }, [mode, rows, catCol, numCol, levels.join("|")]);

  const chi2Result = useMemo(() => {
    if (mode !== "chi2" || !catCol || !catCol2) return null;
    const a = rows.map((r) => r[catCol]).filter((v) => v !== null && v !== undefined);
    const b = rows.map((r) => r[catCol2]).filter((v) => v !== null && v !== undefined);
    const n = Math.min(a.length, b.length);
    return chiSquareTest(a.slice(0, n), b.slice(0, n));
  }, [mode, rows, catCol, catCol2]);

  const corrResult = useMemo(() => {
    if (mode !== "corr" || !numCol || !numCol2) return null;
    const x = rows.map((r) => r[numCol]).filter(isNum);
    const y = rows.map((r) => r[numCol2]).filter(isNum);
    const n = Math.min(x.length, y.length);
    const r = corrMethod === "spearman" ? spearman(x.slice(0, n), y.slice(0, n)) : pearson(x.slice(0, n), y.slice(0, n));
    const df = n - 2;
    const t = df > 0 && Math.abs(r) < 1 ? (r * Math.sqrt(df / (1 - r * r))) : NaN;
    const p = df > 0 ? tTwoTailP(t, df) : NaN;
    return { r, n, df, p, method: corrMethod };
  }, [mode, rows, numCol, numCol2, corrMethod]);

  const batchResults = useMemo(() => {
    if (!batchCat) return [];
    const batchLevels = _.uniq(rows.map((r) => r[batchCat]).filter((v) => v !== null && v !== undefined)).slice(0, 40);
    if (batchLevels.length < 2) return [];
    return numericCols.map((c) => {
      const groups = batchLevels.map((lv) => rows.filter((r) => r[batchCat] === lv).map((r) => r[c.name]).filter(isNum));
      if (batchLevels.length === 2) {
        const res = twoSampleTTest(groups[0], groups[1]);
        return res ? { variable: c.name, test: "t-test", stat: res.t, p: res.p, effect: res.cohensD, effectLabel: cohensDLabel(res.cohensD) } : null;
      }
      const res = oneWayANOVA(groups);
      return res ? { variable: c.name, test: "ANOVA", stat: res.F, p: res.p, effect: res.etaSquared, effectLabel: etaSquaredLabel(res.etaSquared) } : null;
    }).filter(Boolean);
  }, [rows, numericCols.map((c) => c.name).join(","), batchCat]);

  const correctedP = useMemo(() => {
    const pvals = batchResults.map((r) => r.p);
    if (correction === "bonferroni") return bonferroniCorrect(pvals);
    if (correction === "fdr") return benjaminiHochberg(pvals);
    return pvals;
  }, [batchResults, correction]);

  const logTTest = () => {
    if (!ttestResult) return;
    onLogAction?.({
      label: `t-test: ${numCol} entre "${levelA}" e "${levelB}" de ${catCol}`,
      code: `# Teste t (Welch) — ${numCol} por ${catCol}\ngroup_a = df.loc[df['${catCol}'] == '${levelA}', '${numCol}'].dropna()\ngroup_b = df.loc[df['${catCol}'] == '${levelB}', '${numCol}'].dropna()\nt_stat, p_value = stats.ttest_ind(group_a, group_b, equal_var=False)\nprint(f"t={t_stat:.4f}, p={p_value:.4f}")`,
    });
  };
  const logAnova = () => {
    if (!anovaResult) return;
    onLogAction?.({
      label: `ANOVA: ${numCol} por ${catCol}`,
      code: `# ANOVA de 1 fator — ${numCol} por ${catCol}\ngroups = [g['${numCol}'].dropna().values for _, g in df.groupby('${catCol}')]\nf_stat, p_value = stats.f_oneway(*groups)\nprint(f"F={f_stat:.4f}, p={p_value:.4f}")`,
    });
  };
  const logChi2 = () => {
    if (!chi2Result) return;
    onLogAction?.({
      label: `Qui-quadrado: ${catCol} × ${catCol2}`,
      code: `# Qui-quadrado de independência — ${catCol} x ${catCol2}\ncontingency = pd.crosstab(df['${catCol}'], df['${catCol2}'])\nchi2, p_value, dof, expected = stats.chi2_contingency(contingency)\nprint(f"chi2={chi2:.4f}, p={p_value:.4f}")`,
    });
  };
  const logCorr = () => {
    if (!corrResult) return;
    const fn = corrMethod === "spearman" ? "spearmanr" : "pearsonr";
    onLogAction?.({
      label: `Correlação (${corrMethod}): ${numCol} × ${numCol2}`,
      code: `# Correlação ${corrMethod} com significância — ${numCol} x ${numCol2}\nr, p_value = stats.${fn}(df['${numCol}'].dropna(), df['${numCol2}'].dropna())\nprint(f"r={r:.4f}, p={p_value:.4f}")`,
    });
  };
  const logBatch = () => {
    if (!batchResults.length) return;
    onLogAction?.({
      label: `Testes em lote: todas as numéricas vs. "${batchCat}" (correção ${correction})`,
      code: `# Testes em lote — cada variável numérica vs. ${batchCat}, com correção ${correction}\nfrom statsmodels.stats.multitest import multipletests\nresults = []\nfor col in ${JSON.stringify(numericCols.map((c) => c.name))}:\n    groups = [g[col].dropna().values for _, g in df.groupby('${batchCat}')]\n    groups = [g for g in groups if len(g) > 1]\n    if len(groups) == 2:\n        stat, p = stats.ttest_ind(*groups, equal_var=False)\n    else:\n        stat, p = stats.f_oneway(*groups)\n    results.append({"variavel": col, "estatistica": stat, "p": p})\nresults_df = pd.DataFrame(results)\n_, results_df["p_ajustado"], _, _ = multipletests(results_df["p"], method="${correction === "bonferroni" ? "bonferroni" : "fdr_bh"}")`,
    });
  };

  return (
    <div>
      <SectionTitle eyebrow="Inferência estatística" title="Testes de Hipótese"
        right={<div style={{ fontSize: 11.5, color: T.faint, maxWidth: 280, textAlign: "right" }}>Nível de significância α = 0.05 em todos os testes.</div>} />

      <div style={{ display: "flex", gap: 8, marginBottom: 18, flexWrap: "wrap" }}>
        {MODES.map((m) => {
          const active = mode === m.id;
          return (
            <button key={m.id} onClick={() => setMode(m.id)}
              style={{ padding: "9px 15px", borderRadius: 10, border: `1px solid ${active ? T.teal : T.border}`, background: active ? T.tealSoft : "#fff", cursor: "pointer", fontWeight: 700, fontSize: 12.5, color: active ? T.tealDark : T.ink }}>
              {m.label}
            </button>
          );
        })}
      </div>

      {mode === "ttest" && (
        <Card style={{ padding: 20, marginBottom: 16 }}>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 16 }}>
            <Select value={catCol} onChange={setCatCol} options={catCols.map((c) => c.name)} placeholder="Variável categórica (grupo)" />
            <Select value={levelA} onChange={setLevelA} options={levels} placeholder="Grupo A" />
            <Select value={levelB} onChange={setLevelB} options={levels} placeholder="Grupo B" />
            <Select value={numCol} onChange={setNumCol} options={numericCols.map((c) => c.name)} placeholder="Variável numérica" />
          </div>
          {!catCols.length && <div style={{ color: T.sub, fontSize: 13 }}>É preciso de ao menos uma coluna categórica.</div>}
          {ttestResult && (
            <>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 14 }}>
                <StatTile label="t" value={fmt(ttestResult.t, 3)} />
                <StatTile label="Graus de liberdade" value={fmt(ttestResult.df, 1)} />
                <StatTile label="p-valor" value={fmtP(ttestResult.p)} tone={ttestResult.p < 0.05 ? "green" : "amber"} />
                <StatTile label="Diferença de médias" value={fmt(ttestResult.diff, 3)} sub={`IC95%: [${fmt(ttestResult.ciLow, 2)}, ${fmt(ttestResult.ciHigh, 2)}]`} />
                <StatTile label="Cohen's d" value={fmt(ttestResult.cohensD, 3)} sub={`efeito ${cohensDLabel(ttestResult.cohensD)}`} />
              </div>
              <div style={{ fontSize: 13, color: T.ink, marginBottom: 12 }}>{conclusionSentence(ttestResult.p)}</div>
              <Btn variant="ghost" icon={FileCode} onClick={logTTest}>Adicionar ao script Python</Btn>
            </>
          )}
        </Card>
      )}

      {mode === "anova" && (
        <Card style={{ padding: 20, marginBottom: 16 }}>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 16 }}>
            <Select value={catCol} onChange={setCatCol} options={catCols.map((c) => c.name)} placeholder="Variável categórica (grupo)" />
            <Select value={numCol} onChange={setNumCol} options={numericCols.map((c) => c.name)} placeholder="Variável numérica" />
          </div>
          {anovaResult && (
            <>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 14 }}>
                <StatTile label="F" value={fmt(anovaResult.F, 3)} />
                <StatTile label="df (entre / dentro)" value={`${anovaResult.dfBetween} / ${anovaResult.dfWithin}`} />
                <StatTile label="p-valor" value={fmtP(anovaResult.p)} tone={anovaResult.p < 0.05 ? "green" : "amber"} />
                <StatTile label="η² (eta-quadrado)" value={fmt(anovaResult.etaSquared, 3)} sub={`efeito ${etaSquaredLabel(anovaResult.etaSquared)}`} />
                <StatTile label="Grupos (k)" value={anovaResult.k} />
              </div>
              <div style={{ fontSize: 13, color: T.ink, marginBottom: 12 }}>{conclusionSentence(anovaResult.p)}</div>
              <Btn variant="ghost" icon={FileCode} onClick={logAnova}>Adicionar ao script Python</Btn>
            </>
          )}
        </Card>
      )}

      {mode === "chi2" && (
        <Card style={{ padding: 20, marginBottom: 16 }}>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 16 }}>
            <Select value={catCol} onChange={setCatCol} options={catCols.map((c) => c.name)} placeholder="Categórica A" />
            <Select value={catCol2} onChange={setCatCol2} options={catCols.map((c) => c.name)} placeholder="Categórica B" />
          </div>
          {catCols.length < 2 && <div style={{ color: T.sub, fontSize: 13 }}>É preciso de ao menos 2 colunas categóricas.</div>}
          {chi2Result && (
            <>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 14 }}>
                <StatTile label="χ²" value={fmt(chi2Result.chi2, 3)} />
                <StatTile label="Graus de liberdade" value={chi2Result.df} />
                <StatTile label="p-valor" value={fmtP(chi2Result.p)} tone={chi2Result.p < 0.05 ? "green" : "amber"} />
                <StatTile label="Cramér's V" value={fmt(chi2Result.cramersV, 3)} sub={`associação ${cramersVLabel(chi2Result.cramersV)}`} />
                <StatTile label="N" value={chi2Result.n} />
              </div>
              <div style={{ fontSize: 13, color: T.ink, marginBottom: 12 }}>{conclusionSentence(chi2Result.p)}</div>
              <Btn variant="ghost" icon={FileCode} onClick={logChi2}>Adicionar ao script Python</Btn>
            </>
          )}
        </Card>
      )}

      {mode === "corr" && (
        <Card style={{ padding: 20, marginBottom: 16 }}>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 16 }}>
            <Select value={numCol} onChange={setNumCol} options={numericCols.map((c) => c.name)} placeholder="Variável A" />
            <Select value={numCol2} onChange={setNumCol2} options={numericCols.map((c) => c.name)} placeholder="Variável B" />
            <Select value={corrMethod} onChange={setCorrMethod} options={[{ value: "pearson", label: "Pearson" }, { value: "spearman", label: "Spearman" }]} />
          </div>
          {corrResult && (
            <>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 14 }}>
                <StatTile label="r" value={fmt(corrResult.r, 3)} />
                <StatTile label="N" value={corrResult.n} />
                <StatTile label="Graus de liberdade" value={corrResult.df} />
                <StatTile label="p-valor" value={fmtP(corrResult.p)} tone={corrResult.p < 0.05 ? "green" : "amber"} />
              </div>
              <div style={{ fontSize: 13, color: T.ink, marginBottom: 12 }}>{conclusionSentence(corrResult.p)}</div>
              <Btn variant="ghost" icon={FileCode} onClick={logCorr}>Adicionar ao script Python</Btn>
            </>
          )}
        </Card>
      )}

      <ReportDivider C={T} label="Testes em lote com correção para múltiplas comparações" />
      <Card style={{ padding: 20 }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 16, alignItems: "center" }}>
          <div style={{ fontSize: 12.5, color: T.sub }}>Testar todas as variáveis numéricas contra:</div>
          <Select value={batchCat} onChange={setBatchCat} options={catCols.map((c) => c.name)} placeholder="Variável categórica alvo" />
          <Select value={correction} onChange={setCorrection} options={[{ value: "fdr", label: "Correção FDR (Benjamini-Hochberg)" }, { value: "bonferroni", label: "Correção Bonferroni" }, { value: "none", label: "Sem correção" }]} />
        </div>
        {!batchResults.length && <div style={{ color: T.sub, fontSize: 13 }}>Selecione uma variável categórica com pelo menos 2 grupos.</div>}
        {batchResults.length > 0 && (
          <>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                <thead>
                  <tr style={{ borderBottom: `1px solid ${T.border}`, textAlign: "left" }}>
                    <th style={{ padding: "8px 10px", color: T.sub }}>Variável</th>
                    <th style={{ padding: "8px 10px", color: T.sub }}>Teste</th>
                    <th style={{ padding: "8px 10px", color: T.sub }}>Estatística</th>
                    <th style={{ padding: "8px 10px", color: T.sub }}>p</th>
                    <th style={{ padding: "8px 10px", color: T.sub }}>p ajustado</th>
                    <th style={{ padding: "8px 10px", color: T.sub }}>Efeito</th>
                    <th style={{ padding: "8px 10px", color: T.sub }}>Significativo?</th>
                  </tr>
                </thead>
                <tbody>
                  {batchResults.map((r, i) => (
                    <tr key={r.variable} style={{ borderBottom: `1px solid ${T.border}` }}>
                      <td style={{ padding: "8px 10px", fontWeight: 700, color: T.ink, fontFamily: "'JetBrains Mono', monospace" }}>{r.variable}</td>
                      <td style={{ padding: "8px 10px", color: T.sub }}>{r.test}</td>
                      <td style={{ padding: "8px 10px", fontFamily: "'JetBrains Mono', monospace" }}>{fmt(r.stat, 3)}</td>
                      <td style={{ padding: "8px 10px", fontFamily: "'JetBrains Mono', monospace" }}>{fmtP(r.p)}</td>
                      <td style={{ padding: "8px 10px", fontFamily: "'JetBrains Mono', monospace" }}>{fmtP(correctedP[i])}</td>
                      <td style={{ padding: "8px 10px" }}>{r.effectLabel}</td>
                      <td style={{ padding: "8px 10px" }}>
                        {correctedP[i] < 0.05 ? <Pill tone="green">Sim</Pill> : <Pill tone="neutral">Não</Pill>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{ marginTop: 14 }}>
              <Btn variant="ghost" icon={FileCode} onClick={logBatch}>Adicionar lote ao script Python</Btn>
            </div>
          </>
        )}
      </Card>
    </div>
  );
}

/* =========================================================================
   FEATURE ENGINEERING TAB — colunas por fórmula, binning, transformação, one-hot
   ========================================================================= */
// Tiny safe arithmetic parser: columns are referenced as [Nome da Coluna]. No eval().
function tokenizeFormula(expr) { return expr.match(/\[[^\]]+\]|\d+\.?\d*|\.\d+|[+\-*/()]/g) || []; }
function parseFormulaTokens(tokens) {
  if (!tokens.length) throw new Error("Fórmula vazia.");
  let pos = 0;
  const peek = () => tokens[pos];
  const next = () => tokens[pos++];
  function parseExpr() {
    let node = parseTerm();
    while (peek() === "+" || peek() === "-") { const op = next(); node = { type: "bin", op, left: node, right: parseTerm() }; }
    return node;
  }
  function parseTerm() {
    let node = parseFactor();
    while (peek() === "*" || peek() === "/") { const op = next(); node = { type: "bin", op, left: node, right: parseFactor() }; }
    return node;
  }
  function parseFactor() {
    const tok = peek();
    if (tok === undefined) throw new Error("Fórmula incompleta.");
    if (tok === "(") { next(); const node = parseExpr(); if (peek() !== ")") throw new Error("Parêntese não fechado."); next(); return node; }
    if (tok === "-") { next(); return { type: "neg", value: parseFactor() }; }
    if (tok.startsWith("[")) { next(); return { type: "col", name: tok.slice(1, -1) }; }
    if (/^[\d.]/.test(tok)) { next(); return { type: "num", value: Number(tok) }; }
    throw new Error(`Token inesperado: "${tok}".`);
  }
  const root = parseExpr();
  if (pos !== tokens.length) throw new Error(`Sintaxe inválida perto de "${tokens[pos]}" — verifique se falta um operador (+ − × ÷) entre os termos.`);
  return root;
}
function evalFormulaNode(node, row) {
  if (!node) return NaN;
  if (node.type === "num") return node.value;
  if (node.type === "col") { const v = row[node.name]; return isNum(v) ? v : NaN; }
  if (node.type === "neg") return -evalFormulaNode(node.value, row);
  if (node.type === "bin") {
    const l = evalFormulaNode(node.left, row), r = evalFormulaNode(node.right, row);
    if (node.op === "+") return l + r;
    if (node.op === "-") return l - r;
    if (node.op === "*") return l * r;
    if (node.op === "/") return r !== 0 ? l / r : NaN;
  }
  return NaN;
}
function buildColumnMeta(name, values, type) {
  const nonNull = values.filter((v) => v !== null && v !== undefined && !(typeof v === "number" && Number.isNaN(v)));
  const missing = values.length - nonNull.length;
  const uniqueVals = _.uniq(nonNull.map((v) => String(v)));
  return { name, type, missing, missingPct: values.length ? (missing / values.length) * 100 : 0, unique: uniqueVals.length, numberStyle: "us", derived: true };
}

function FeatureEngineeringTab({ rows, columns, setRows, setColumns, onLogAction }) {
  const numericCols = columns.filter((c) => c.type === "numeric");
  const catCols = columns.filter((c) => c.type === "categorical");
  const derivedCols = columns.filter((c) => c.derived);
  const [subMode, setSubMode] = useState("formula");
  const [msg, setMsg] = useState(null);

  // formula
  const [formulaName, setFormulaName] = useState("");
  const [formulaExpr, setFormulaExpr] = useState("");
  // binning
  const [binCol, setBinCol] = useState(numericCols[0]?.name || "");
  const [binMethod, setBinMethod] = useState("quantile");
  const [binCount, setBinCount] = useState(4);
  const [binName, setBinName] = useState("");
  // transform
  const [transCol, setTransCol] = useState(numericCols[0]?.name || "");
  const [transType, setTransType] = useState("log1p");
  const [transName, setTransName] = useState("");
  // one-hot
  const [oheCol, setOheCol] = useState(catCols[0]?.name || "");
  const [oheTopN, setOheTopN] = useState(5);

  const showMsg = (tone, text) => setMsg({ tone, text });

  const addColumn = (name, values, type, pythonCode, label) => {
    if (!name.trim()) { showMsg("amber", "Dê um nome à nova coluna."); return; }
    if (columns.some((c) => c.name === name)) { showMsg("amber", "Já existe uma coluna com esse nome."); return; }
    const meta = buildColumnMeta(name, values, type);
    setColumns((prev) => [...prev, meta]);
    setRows((prev) => prev.map((r, i) => ({ ...r, [name]: values[i] })));
    onLogAction?.({ label, code: pythonCode });
    showMsg("green", `Coluna "${name}" criada com sucesso.`);
  };

  const removeColumn = (name) => {
    setColumns((prev) => prev.filter((c) => c.name !== name));
    setRows((prev) => prev.map((r) => { const { [name]: _drop, ...rest } = r; return rest; }));
  };

  const applyFormula = () => {
    if (!formulaExpr.trim()) { showMsg("amber", "Escreva uma fórmula usando [Nome da Coluna]."); return; }
    try {
      const ast = parseFormulaTokens(tokenizeFormula(formulaExpr));
      const values = rows.map((r) => { const v = evalFormulaNode(ast, r); return Number.isFinite(v) ? v : null; });
      const pyExpr = formulaExpr.replace(/\[([^\]]+)\]/g, "df['$1']");
      addColumn(formulaName || "nova_coluna", values, "numeric", `# Coluna derivada por fórmula\ndf['${formulaName || "nova_coluna"}'] = ${pyExpr}`, `Fórmula: ${formulaName} = ${formulaExpr}`);
    } catch (err) { showMsg("red", err?.message || "Fórmula inválida."); }
  };

  const applyBinning = () => {
    const vals = rows.map((r) => r[binCol]);
    const name = binName || `${binCol}_faixa`;
    let edges = [];
    const numericVals = vals.filter(isNum);
    if (!numericVals.length) { showMsg("amber", "Coluna sem valores numéricos suficientes."); return; }
    if (binMethod === "quantile") {
      for (let i = 1; i < binCount; i++) edges.push(percentile(numericVals, (i / binCount) * 100));
    } else {
      const lo = Math.min(...numericVals), hi = Math.max(...numericVals);
      for (let i = 1; i < binCount; i++) edges.push(lo + ((hi - lo) * i) / binCount);
    }
    edges = _.uniq(edges);
    const values = vals.map((v) => {
      if (!isNum(v)) return null;
      let b = 0; while (b < edges.length && v > edges[b]) b++;
      return `Faixa ${b + 1}`;
    });
    const pyMethod = binMethod === "quantile" ? `pd.qcut(df['${binCol}'], q=${binCount}, labels=[f"Faixa {i+1}" for i in range(${binCount})], duplicates='drop')` : `pd.cut(df['${binCol}'], bins=${binCount}, labels=[f"Faixa {i+1}" for i in range(${binCount})])`;
    addColumn(name, values, "categorical", `# Binning (${binMethod}) de ${binCol}\ndf['${name}'] = ${pyMethod}`, `Binning: ${name} (${binMethod}, ${binCount} faixas)`);
  };

  const applyTransform = () => {
    const vals = rows.map((r) => r[transCol]);
    const name = transName || `${transCol}_${transType}`;
    let values, pyCode;
    if (transType === "log1p") { values = vals.map((v) => isNum(v) && v > -1 ? Math.log1p(v) : null); pyCode = `np.log1p(df['${transCol}'])`; }
    else if (transType === "sqrt") { values = vals.map((v) => isNum(v) && v >= 0 ? Math.sqrt(v) : null); pyCode = `np.sqrt(df['${transCol}'])`; }
    else if (transType === "zscore") {
      const numericVals = vals.filter(isNum);
      const m = mean(numericVals), s = std(numericVals, true);
      values = vals.map((v) => isNum(v) && s ? (v - m) / s : null);
      pyCode = `(df['${transCol}'] - df['${transCol}'].mean()) / df['${transCol}'].std()`;
    } else {
      const numericVals = vals.filter(isNum);
      const lo = Math.min(...numericVals), hi = Math.max(...numericVals);
      values = vals.map((v) => isNum(v) && hi !== lo ? (v - lo) / (hi - lo) : null);
      pyCode = `(df['${transCol}'] - df['${transCol}'].min()) / (df['${transCol}'].max() - df['${transCol}'].min())`;
    }
    addColumn(name, values, "numeric", `# Transformação (${transType}) de ${transCol}\ndf['${name}'] = ${pyCode}`, `Transformação: ${name} (${transType})`);
  };

  const applyOneHot = () => {
    if (!oheCol) return;
    const counts = _.countBy(rows.map((r) => r[oheCol]).filter((v) => v !== null && v !== undefined));
    const topCats = Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, oheTopN).map(([k]) => k);
    setColumns((prev) => {
      const newCols = topCats.map((cat) => buildColumnMeta(`${oheCol}_${cat}`, rows.map((r) => (String(r[oheCol]) === cat ? 1 : 0)), "numeric"));
      return [...prev, ...newCols];
    });
    setRows((prev) => prev.map((r) => {
      const extra = {};
      topCats.forEach((cat) => { extra[`${oheCol}_${cat}`] = String(r[oheCol]) === cat ? 1 : 0; });
      return { ...r, ...extra };
    }));
    onLogAction?.({ label: `One-hot encoding: ${oheCol} (top ${oheTopN})`, code: `# One-hot encoding — ${oheCol} (top ${oheTopN} categorias)\ntop_cats = df['${oheCol}'].value_counts().head(${oheTopN}).index\nfor cat in top_cats:\n    df[f"${oheCol}_{cat}"] = (df['${oheCol}'] == cat).astype(int)` });
    showMsg("green", `${topCats.length} colunas one-hot criadas a partir de "${oheCol}".`);
  };

  const SUBMODES = [
    { id: "formula", label: "Fórmula" },
    { id: "binning", label: "Binning" },
    { id: "transform", label: "Transformação" },
    { id: "onehot", label: "One-hot encoding" },
  ];

  return (
    <div>
      <SectionTitle eyebrow="Preparação de dados" title="Engenharia de Features" />
      <div style={{ display: "flex", gap: 8, marginBottom: 18, flexWrap: "wrap" }}>
        {SUBMODES.map((m) => {
          const active = subMode === m.id;
          return (
            <button key={m.id} onClick={() => { setSubMode(m.id); setMsg(null); }}
              style={{ padding: "9px 15px", borderRadius: 10, border: `1px solid ${active ? T.teal : T.border}`, background: active ? T.tealSoft : "#fff", cursor: "pointer", fontWeight: 700, fontSize: 12.5, color: active ? T.tealDark : T.ink }}>
              {m.label}
            </button>
          );
        })}
      </div>

      {subMode === "formula" && (
        <Card style={{ padding: 20, marginBottom: 16 }}>
          <div style={{ fontSize: 12.5, color: T.sub, marginBottom: 10 }}>
            Use <code>[Nome da Coluna]</code> para referenciar colunas numéricas. Operadores: + − × ÷ e parênteses.
          </div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 10 }}>
            {numericCols.map((c) => (
              <button key={c.name} onClick={() => setFormulaExpr((e) => e + `[${c.name}]`)}
                style={{ fontSize: 11, padding: "4px 8px", borderRadius: 6, border: `1px solid ${T.border}`, background: "#fff", cursor: "pointer", fontFamily: "'JetBrains Mono', monospace" }}>
                {c.name}
              </button>
            ))}
          </div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <input placeholder="Nome da nova coluna" value={formulaName} onChange={(e) => setFormulaName(e.target.value)}
              style={{ padding: "9px 11px", borderRadius: 8, border: `1px solid ${T.border}`, fontSize: 13, minWidth: 200 }} />
            <input placeholder="Ex.: [Receita] / [Unidades]" value={formulaExpr} onChange={(e) => setFormulaExpr(e.target.value)}
              style={{ padding: "9px 11px", borderRadius: 8, border: `1px solid ${T.border}`, fontSize: 13, flex: 1, minWidth: 240, fontFamily: "'JetBrains Mono', monospace" }} />
            <Btn icon={Plus} onClick={applyFormula}>Criar coluna</Btn>
          </div>
        </Card>
      )}

      {subMode === "binning" && (
        <Card style={{ padding: 20, marginBottom: 16 }}>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <Select value={binCol} onChange={setBinCol} options={numericCols.map((c) => c.name)} />
            <Select value={binMethod} onChange={setBinMethod} options={[{ value: "quantile", label: "Quantis (mesmo N por faixa)" }, { value: "equalwidth", label: "Largura igual" }]} />
            <input type="number" min={2} max={20} value={binCount} onChange={(e) => setBinCount(Math.max(2, Math.min(20, +e.target.value || 2)))}
              style={{ padding: "9px 11px", borderRadius: 8, border: `1px solid ${T.border}`, fontSize: 13, width: 90 }} />
            <input placeholder="Nome da nova coluna (opcional)" value={binName} onChange={(e) => setBinName(e.target.value)}
              style={{ padding: "9px 11px", borderRadius: 8, border: `1px solid ${T.border}`, fontSize: 13, minWidth: 200 }} />
            <Btn icon={Plus} onClick={applyBinning}>Criar coluna</Btn>
          </div>
        </Card>
      )}

      {subMode === "transform" && (
        <Card style={{ padding: 20, marginBottom: 16 }}>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <Select value={transCol} onChange={setTransCol} options={numericCols.map((c) => c.name)} />
            <Select value={transType} onChange={setTransType} options={[{ value: "log1p", label: "log(1+x)" }, { value: "sqrt", label: "Raiz quadrada" }, { value: "zscore", label: "Padronização (z-score)" }, { value: "minmax", label: "Normalização (min-max)" }]} />
            <input placeholder="Nome da nova coluna (opcional)" value={transName} onChange={(e) => setTransName(e.target.value)}
              style={{ padding: "9px 11px", borderRadius: 8, border: `1px solid ${T.border}`, fontSize: 13, minWidth: 200 }} />
            <Btn icon={Plus} onClick={applyTransform}>Criar coluna</Btn>
          </div>
        </Card>
      )}

      {subMode === "onehot" && (
        <Card style={{ padding: 20, marginBottom: 16 }}>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
            <Select value={oheCol} onChange={setOheCol} options={catCols.map((c) => c.name)} />
            <div style={{ fontSize: 12.5, color: T.sub }}>Top</div>
            <input type="number" min={2} max={20} value={oheTopN} onChange={(e) => setOheTopN(Math.max(2, Math.min(20, +e.target.value || 2)))}
              style={{ padding: "9px 11px", borderRadius: 8, border: `1px solid ${T.border}`, fontSize: 13, width: 80 }} />
            <div style={{ fontSize: 12.5, color: T.sub }}>categorias mais frequentes</div>
            <Btn icon={Plus} onClick={applyOneHot}>Criar colunas</Btn>
          </div>
        </Card>
      )}

      {msg && <div style={{ marginBottom: 16, fontSize: 12.5, color: T[msg.tone] || T.sub, background: msg.tone === "green" ? T.greenSoft : msg.tone === "red" ? T.redSoft : T.amberSoft, padding: 10, borderRadius: 8 }}>{msg.text}</div>}

      <ReportDivider C={T} label="Colunas derivadas nesta sessão" />
      <Card style={{ padding: 20 }}>
        {!derivedCols.length && <div style={{ color: T.sub, fontSize: 13 }}>Nenhuma coluna derivada ainda. Use as ferramentas acima para criar features.</div>}
        {derivedCols.length > 0 && (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {derivedCols.map((c) => (
              <div key={c.name} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 10px", background: T.bg, borderRadius: 8 }}>
                <div style={{ fontSize: 12.5, fontFamily: "'JetBrains Mono', monospace", color: T.ink }}>{c.name} <span style={{ color: T.faint }}>({c.type})</span></div>
                <button onClick={() => removeColumn(c.name)} style={{ border: "none", background: "transparent", cursor: "pointer", color: T.red }}><Trash2 size={14} /></button>
              </div>
            ))}
          </div>
        )}
      </Card>
    </div>
  );
}

/* =========================================================================
   GROUP COMPARISON TAB — group-by / pivot com teste estatístico integrado
   ========================================================================= */
function aggregateValues(vals, method) {
  if (method === "count") return vals.length;
  const nums = vals.filter(isNum);
  if (!nums.length) return NaN;
  if (method === "mean") return mean(nums);
  if (method === "median") return median(nums);
  if (method === "sum") return _.sum(nums);
  if (method === "std") return std(nums, true);
  if (method === "min") return Math.min(...nums);
  if (method === "max") return Math.max(...nums);
  return NaN;
}
const AGG_LABELS = { mean: "Média", median: "Mediana", sum: "Soma", std: "Desvio padrão", min: "Mínimo", max: "Máximo", count: "Contagem" };
const AGG_PY = { mean: "mean", median: "median", sum: "sum", std: "std", min: "min", max: "max", count: "count" };

function GroupComparisonTab({ rows, columns, onLogAction }) {
  const numericCols = columns.filter((c) => c.type === "numeric");
  const catCols = columns.filter((c) => c.type === "categorical");
  const [groupCol, setGroupCol] = useState(catCols[0]?.name || "");
  const [breakCol, setBreakCol] = useState("");
  const [metricCol, setMetricCol] = useState(numericCols[0]?.name || "");
  const [agg, setAgg] = useState("mean");

  const levels = useMemo(() => (groupCol ? _.uniq(rows.map((r) => r[groupCol]).filter((v) => v !== null && v !== undefined)).slice(0, 30) : []), [rows, groupCol]);
  const breakLevels = useMemo(() => (breakCol ? _.uniq(rows.map((r) => r[breakCol]).filter((v) => v !== null && v !== undefined)).slice(0, 12) : []), [rows, breakCol]);

  const groupTable = useMemo(() => {
    if (!groupCol || !metricCol) return [];
    return levels.map((lv) => {
      const subset = rows.filter((r) => r[groupCol] === lv).map((r) => r[metricCol]);
      return { group: String(lv), n: subset.filter((v) => v !== null && v !== undefined).length, value: aggregateValues(subset, agg) };
    });
  }, [rows, groupCol, metricCol, agg, levels.join("|")]);

  const pivotTable = useMemo(() => {
    if (!groupCol || !breakCol || !metricCol) return null;
    return levels.map((lv) => ({
      group: String(lv),
      cells: breakLevels.map((bl) => {
        const subset = rows.filter((r) => r[groupCol] === lv && r[breakCol] === bl).map((r) => r[metricCol]);
        return aggregateValues(subset, agg);
      }),
    }));
  }, [rows, groupCol, breakCol, metricCol, agg, levels.join("|"), breakLevels.join("|")]);

  const testResult = useMemo(() => {
    if (!groupCol || !metricCol || levels.length < 2) return null;
    const groups = levels.map((lv) => rows.filter((r) => r[groupCol] === lv).map((r) => r[metricCol]).filter(isNum));
    if (levels.length === 2) { const res = twoSampleTTest(groups[0], groups[1]); return res ? { kind: "t-test", ...res } : null; }
    const res = oneWayANOVA(groups);
    return res ? { kind: "ANOVA", ...res } : null;
  }, [rows, groupCol, metricCol, levels.join("|")]);

  const logGroupBy = () => {
    onLogAction?.({
      label: `Group-by: ${metricCol} agregado (${agg}) por ${groupCol}`,
      code: `# Comparação de grupos\ndf.groupby('${groupCol}')['${metricCol}'].${AGG_PY[agg]}()${breakCol ? `\n# Pivot com quebra adicional\npd.pivot_table(df, values='${metricCol}', index='${groupCol}', columns='${breakCol}', aggfunc='${AGG_PY[agg]}')` : ""}`,
    });
  };

  return (
    <div>
      <SectionTitle eyebrow="Análise por grupo" title="Comparação de Grupos (Pivot)" />
      <Card style={{ padding: 20, marginBottom: 16 }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <Select value={groupCol} onChange={setGroupCol} options={catCols.map((c) => c.name)} placeholder="Agrupar por" />
          <Select value={breakCol} onChange={setBreakCol} options={catCols.filter((c) => c.name !== groupCol).map((c) => c.name)} placeholder="Quebrar por (opcional)" />
          <Select value={metricCol} onChange={setMetricCol} options={numericCols.map((c) => c.name)} placeholder="Métrica" />
          <Select value={agg} onChange={setAgg} options={Object.entries(AGG_LABELS).map(([value, label]) => ({ value, label }))} />
        </div>
      </Card>

      {!pivotTable && groupTable.length > 0 && (
        <Card style={{ padding: 20, marginBottom: 16 }}>
          <div style={{ fontWeight: 700, fontSize: 14, color: T.ink, marginBottom: 12 }}>{AGG_LABELS[agg]} de {metricCol} por {groupCol}</div>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={groupTable}>
              <CartesianGrid strokeDasharray="3 3" stroke={T.border} />
              <XAxis dataKey="group" tick={{ fontSize: 10 }} />
              <YAxis tick={{ fontSize: 10 }} />
              <Tooltip />
              <Bar dataKey="value" fill={T.teal} radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
          <div style={{ overflowX: "auto", marginTop: 16 }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
              <thead>
                <tr style={{ borderBottom: `1px solid ${T.border}`, textAlign: "left" }}>
                  <th style={{ padding: "8px 10px", color: T.sub }}>{groupCol}</th>
                  <th style={{ padding: "8px 10px", color: T.sub }}>N</th>
                  <th style={{ padding: "8px 10px", color: T.sub }}>{AGG_LABELS[agg]}</th>
                </tr>
              </thead>
              <tbody>
                {groupTable.map((r) => (
                  <tr key={r.group} style={{ borderBottom: `1px solid ${T.border}` }}>
                    <td style={{ padding: "8px 10px", fontWeight: 700, color: T.ink }}>{r.group}</td>
                    <td style={{ padding: "8px 10px", fontFamily: "'JetBrains Mono', monospace" }}>{r.n}</td>
                    <td style={{ padding: "8px 10px", fontFamily: "'JetBrains Mono', monospace" }}>{fmt(r.value, 3)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      )}

      {pivotTable && (
        <Card style={{ padding: 20, marginBottom: 16, overflowX: "auto" }}>
          <div style={{ fontWeight: 700, fontSize: 14, color: T.ink, marginBottom: 12 }}>Pivot: {AGG_LABELS[agg]} de {metricCol} — {groupCol} × {breakCol}</div>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, minWidth: 500 }}>
            <thead>
              <tr style={{ borderBottom: `1px solid ${T.border}`, textAlign: "left" }}>
                <th style={{ padding: "8px 10px", color: T.sub }}>{groupCol} \ {breakCol}</th>
                {breakLevels.map((bl) => <th key={bl} style={{ padding: "8px 10px", color: T.sub, whiteSpace: "nowrap" }}>{String(bl)}</th>)}
              </tr>
            </thead>
            <tbody>
              {pivotTable.map((row) => (
                <tr key={row.group} style={{ borderBottom: `1px solid ${T.border}` }}>
                  <td style={{ padding: "8px 10px", fontWeight: 700, color: T.ink }}>{row.group}</td>
                  {row.cells.map((v, j) => <td key={j} style={{ padding: "8px 10px", fontFamily: "'JetBrains Mono', monospace" }}>{fmt(v, 2)}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </Card>
      )}

      <Btn variant="ghost" icon={FileCode} onClick={logGroupBy} style={{ marginBottom: 16 }}>Adicionar ao script Python</Btn>

      {testResult && (
        <>
          <ReportDivider C={T} label="Diferença entre grupos é estatisticamente significativa?" />
          <Card style={{ padding: 20 }}>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 14 }}>
              <StatTile label={testResult.kind} value={fmt(testResult.kind === "t-test" ? testResult.t : testResult.F, 3)} />
              <StatTile label="p-valor" value={fmtP(testResult.p)} tone={testResult.p < 0.05 ? "green" : "amber"} />
              <StatTile label="Tamanho de efeito" value={fmt(testResult.kind === "t-test" ? testResult.cohensD : testResult.etaSquared, 3)}
                sub={testResult.kind === "t-test" ? cohensDLabel(testResult.cohensD) : etaSquaredLabel(testResult.etaSquared)} />
            </div>
            <div style={{ fontSize: 13, color: T.ink }}>{conclusionSentence(testResult.p)}</div>
          </Card>
        </>
      )}
    </div>
  );
}

/* =========================================================================
   DATA QUALITY RULES TAB — data-contracts leve, validado contra a base ativa
   ========================================================================= */
const RULE_TYPE_LABELS = { notnull: "Completude mínima", range: "Intervalo numérico", allowed: "Valores permitidos", regex: "Padrão (regex)", unique: "Unicidade mínima" };
const QUALITY_RULES_KEY = "analisepro_quality_rules_v1";

function loadQualityRules() {
  try { return JSON.parse(localStorage.getItem(QUALITY_RULES_KEY) || "[]"); } catch { return []; }
}
function saveQualityRules(rules) {
  try { localStorage.setItem(QUALITY_RULES_KEY, JSON.stringify(rules)); } catch { /* noop */ }
}

function validateQualityRule(rule, rows) {
  const vals = rows.map((r) => r[rule.column]);
  const n = vals.length;
  if (rule.type === "notnull") {
    const nonNull = vals.filter((v) => v !== null && v !== undefined && String(v).trim() !== "");
    const pct = n ? (nonNull.length / n) * 100 : 0;
    return { passed: pct >= rule.minPct, metricLabel: `${fmt(pct, 1)}% preenchido (mín. ${rule.minPct}%)`, violations: n - nonNull.length };
  }
  if (rule.type === "range") {
    const violators = vals.filter((v) => isNum(v) && (v < rule.min || v > rule.max));
    return { passed: violators.length === 0, metricLabel: `${violators.length} fora de [${rule.min}, ${rule.max}]`, violations: violators.length, examples: violators.slice(0, 5) };
  }
  if (rule.type === "allowed") {
    const allowSet = new Set(String(rule.values || "").split(",").map((s) => s.trim()).filter(Boolean));
    const violators = vals.filter((v) => v !== null && v !== undefined && !allowSet.has(String(v)));
    return { passed: violators.length === 0, metricLabel: `${violators.length} valores não permitidos`, violations: violators.length, examples: _.uniq(violators).slice(0, 5) };
  }
  if (rule.type === "regex") {
    let re; try { re = new RegExp(rule.pattern); } catch { return { passed: false, metricLabel: "Padrão regex inválido", violations: n }; }
    const violators = vals.filter((v) => v !== null && v !== undefined && !re.test(String(v)));
    return { passed: violators.length === 0, metricLabel: `${violators.length} não correspondem ao padrão`, violations: violators.length, examples: _.uniq(violators).slice(0, 5) };
  }
  if (rule.type === "unique") {
    const nonNull = vals.filter((v) => v !== null && v !== undefined);
    const uniqueCount = _.uniq(nonNull.map(String)).length;
    const pct = nonNull.length ? (uniqueCount / nonNull.length) * 100 : 0;
    return { passed: pct >= rule.minPct, metricLabel: `${fmt(pct, 1)}% únicos (mín. ${rule.minPct}%)`, violations: nonNull.length - uniqueCount };
  }
  return { passed: true, metricLabel: "—", violations: 0 };
}

function DataQualityTab({ rows, columns, onLogAction }) {
  const [rules, setRules] = useState(loadQualityRules);
  const [ruleType, setRuleType] = useState("notnull");
  const [ruleCol, setRuleCol] = useState(columns[0]?.name || "");
  const [params, setParams] = useState({ minPct: 95, min: 0, max: 100, values: "", pattern: "" });

  useEffect(() => { saveQualityRules(rules); }, [rules]);

  const addRule = () => {
    if (!ruleCol) return;
    const rule = { id: uid(), column: ruleCol, type: ruleType, ...params };
    setRules((prev) => [...prev, rule]);
  };
  const removeRule = (id) => setRules((prev) => prev.filter((r) => r.id !== id));

  const results = useMemo(() => rules.map((rule) => {
    const colExists = columns.some((c) => c.name === rule.column);
    if (!colExists) return { rule, colExists: false };
    return { rule, colExists: true, ...validateQualityRule(rule, rows) };
  }), [rules, rows, columns.map((c) => c.name).join(",")]);

  const passCount = results.filter((r) => r.passed).length;

  const logRules = () => {
    const lines = rules.map((r) => {
      if (r.type === "notnull") return `assert (df['${r.column}'].notna().mean() * 100) >= ${r.minPct}, "Falha: completude de ${r.column}"`;
      if (r.type === "range") return `assert df['${r.column}'].between(${r.min}, ${r.max}).all(), "Falha: intervalo de ${r.column}"`;
      if (r.type === "allowed") return `assert df['${r.column}'].isin(${JSON.stringify(String(r.values || "").split(",").map((s) => s.trim()).filter(Boolean))}).all(), "Falha: valores permitidos de ${r.column}"`;
      if (r.type === "regex") return `assert df['${r.column}'].astype(str).str.match(r"${r.pattern}").all(), "Falha: padrão de ${r.column}"`;
      if (r.type === "unique") return `assert (df['${r.column}'].nunique() / df['${r.column}'].count() * 100) >= ${r.minPct}, "Falha: unicidade de ${r.column}"`;
      return "";
    });
    onLogAction?.({ label: `Regras de qualidade de dados (${rules.length})`, code: `# Validação de regras de qualidade (data contract leve)\n${lines.join("\n")}` });
  };

  return (
    <div>
      <SectionTitle eyebrow="Data contracts" title="Regras de Qualidade"
        right={rules.length > 0 && <div style={{ fontSize: 12.5, color: T.sub }}>{passCount}/{rules.length} regras aprovadas</div>} />

      <Card style={{ padding: 20, marginBottom: 16 }}>
        <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 12, color: T.ink }}>Nova regra</div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
          <Select value={ruleCol} onChange={setRuleCol} options={columns.map((c) => c.name)} placeholder="Coluna" />
          <Select value={ruleType} onChange={setRuleType} options={Object.entries(RULE_TYPE_LABELS).map(([value, label]) => ({ value, label }))} />
          {ruleType === "notnull" && (
            <>
              <div style={{ fontSize: 12.5, color: T.sub }}>Mín. % preenchido</div>
              <input type="number" min={0} max={100} value={params.minPct} onChange={(e) => setParams((p) => ({ ...p, minPct: +e.target.value }))}
                style={{ padding: "9px 11px", borderRadius: 8, border: `1px solid ${T.border}`, fontSize: 13, width: 90 }} />
            </>
          )}
          {ruleType === "range" && (
            <>
              <input type="number" placeholder="Mín" value={params.min} onChange={(e) => setParams((p) => ({ ...p, min: +e.target.value }))}
                style={{ padding: "9px 11px", borderRadius: 8, border: `1px solid ${T.border}`, fontSize: 13, width: 100 }} />
              <input type="number" placeholder="Máx" value={params.max} onChange={(e) => setParams((p) => ({ ...p, max: +e.target.value }))}
                style={{ padding: "9px 11px", borderRadius: 8, border: `1px solid ${T.border}`, fontSize: 13, width: 100 }} />
            </>
          )}
          {ruleType === "allowed" && (
            <input placeholder="valores separados por vírgula" value={params.values} onChange={(e) => setParams((p) => ({ ...p, values: e.target.value }))}
              style={{ padding: "9px 11px", borderRadius: 8, border: `1px solid ${T.border}`, fontSize: 13, minWidth: 240 }} />
          )}
          {ruleType === "regex" && (
            <input placeholder="expressão regular" value={params.pattern} onChange={(e) => setParams((p) => ({ ...p, pattern: e.target.value }))}
              style={{ padding: "9px 11px", borderRadius: 8, border: `1px solid ${T.border}`, fontSize: 13, minWidth: 200, fontFamily: "'JetBrains Mono', monospace" }} />
          )}
          {ruleType === "unique" && (
            <>
              <div style={{ fontSize: 12.5, color: T.sub }}>Mín. % únicos</div>
              <input type="number" min={0} max={100} value={params.minPct} onChange={(e) => setParams((p) => ({ ...p, minPct: +e.target.value }))}
                style={{ padding: "9px 11px", borderRadius: 8, border: `1px solid ${T.border}`, fontSize: 13, width: 90 }} />
            </>
          )}
          <Btn icon={Plus} onClick={addRule}>Adicionar regra</Btn>
        </div>
        <div style={{ fontSize: 11.5, color: T.faint, marginTop: 10 }}>
          <Info size={12} style={{ verticalAlign: -1 }} /> As regras ficam salvas neste navegador e são revalidadas automaticamente ao trocar de base.
        </div>
      </Card>

      {!results.length && <div style={{ color: T.sub, fontSize: 13 }}>Nenhuma regra cadastrada ainda.</div>}

      {results.length > 0 && (
        <>
          <Card style={{ padding: 20, marginBottom: 16, overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
              <thead>
                <tr style={{ borderBottom: `1px solid ${T.border}`, textAlign: "left" }}>
                  <th style={{ padding: "8px 10px", color: T.sub }}>Status</th>
                  <th style={{ padding: "8px 10px", color: T.sub }}>Coluna</th>
                  <th style={{ padding: "8px 10px", color: T.sub }}>Regra</th>
                  <th style={{ padding: "8px 10px", color: T.sub }}>Resultado</th>
                  <th style={{ padding: "8px 10px", color: T.sub }} />
                </tr>
              </thead>
              <tbody>
                {results.map(({ rule, colExists, passed, metricLabel }) => (
                  <tr key={rule.id} style={{ borderBottom: `1px solid ${T.border}` }}>
                    <td style={{ padding: "8px 10px" }}>
                      {!colExists ? <Pill tone="neutral">N/D</Pill> : passed ? <CheckCircle2 size={16} color={T.green} /> : <XCircle size={16} color={T.red} />}
                    </td>
                    <td style={{ padding: "8px 10px", fontWeight: 700, color: T.ink, fontFamily: "'JetBrains Mono', monospace" }}>{rule.column}</td>
                    <td style={{ padding: "8px 10px", color: T.sub }}>{RULE_TYPE_LABELS[rule.type]}</td>
                    <td style={{ padding: "8px 10px", color: T.ink }}>{colExists ? metricLabel : "Coluna ausente na base atual"}</td>
                    <td style={{ padding: "8px 10px" }}>
                      <button onClick={() => removeRule(rule.id)} style={{ border: "none", background: "transparent", cursor: "pointer", color: T.red }}><Trash2 size={14} /></button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </Card>
          <Btn variant="ghost" icon={FileCode} onClick={logRules}>Adicionar validações ao script Python</Btn>
        </>
      )}
    </div>
  );
}

/* =========================================================================
   PROJECTS TAB  (persisted via window.storage — personal, not shared)
   ========================================================================= */
function uid() { return "p_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8); }

function ProjectsTab({ rows, columns, fileName, onOpenProject }) {
  const [projects, setProjects] = useState([]);
  const [loading, setLoading] = useState(true);
  const [name, setName] = useState("");
  const [desc, setDesc] = useState("");
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState(null);

  const loadList = useCallback(async () => {
    setLoading(true);
    try {
      const listing = await window.storage.list("project:");
      const keys = listing?.keys || [];
      const items = [];
      for (const k of keys) {
        try {
          const r = await window.storage.get(k, false);
          if (r?.value) items.push(JSON.parse(r.value));
        } catch { /* skip */ }
      }
      items.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
      setProjects(items);
    } catch { setProjects([]); }
    setLoading(false);
  }, []);

  useEffect(() => { loadList(); }, [loadList]);

  const save = async () => {
    if (!name.trim()) { setMsg({ tone: "amber", text: "Dê um nome ao projeto." }); return; }
    if (!rows.length) { setMsg({ tone: "amber", text: "Carregue uma base de dados primeiro." }); return; }
    setSaving(true);
    const d = computeDashboard(rows, columns);
    const project = {
      id: uid(), name: name.trim(), description: desc.trim(), fileName,
      createdAt: new Date().toISOString(),
      columns, summary: d,
      sampleRows: rows.slice(0, 300),
    };
    try {
      const payload = JSON.stringify(project);
      if (payload.length > 4.5 * 1024 * 1024) {
        project.sampleRows = rows.slice(0, 60);
      }
      const ok = await window.storage.set(`project:${project.id}`, JSON.stringify(project), false);
      if (!ok) throw new Error("Falha ao salvar.");
      setMsg({ tone: "green", text: "Projeto salvo com sucesso." });
      setName(""); setDesc("");
      loadList();
    } catch (e) { setMsg({ tone: "red", text: "Erro ao salvar: " + e.message }); }
    setSaving(false);
  };

  const remove = async (id) => {
    try { await window.storage.delete(`project:${id}`, false); loadList(); } catch { /* noop */ }
  };

  return (
    <div>
      <SectionTitle eyebrow="Gestão" title="Projetos de Análise" />
      <div style={{ display: "grid", gridTemplateColumns: "360px 1fr", gap: 16 }}>
        <Card style={{ padding: 20 }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 10, color: T.ink }}>Salvar análise atual</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <input placeholder="Nome do projeto" value={name} onChange={(e) => setName(e.target.value)}
              style={{ padding: "9px 11px", borderRadius: 8, border: `1px solid ${T.border}`, fontSize: 13.5 }} />
            <textarea placeholder="Descrição (opcional)" value={desc} onChange={(e) => setDesc(e.target.value)} rows={3}
              style={{ padding: "9px 11px", borderRadius: 8, border: `1px solid ${T.border}`, fontSize: 13.5, resize: "vertical", fontFamily: "'Inter', sans-serif" }} />
            <Btn onClick={save} disabled={saving} icon={Plus}>{saving ? "Salvando…" : "Salvar projeto"}</Btn>
            {msg && <div style={{ fontSize: 12.5, color: T[msg.tone] || T.sub, background: msg.tone === "green" ? T.greenSoft : msg.tone === "red" ? T.redSoft : T.amberSoft, padding: 8, borderRadius: 8 }}>{msg.text}</div>}
            <div style={{ fontSize: 11.5, color: T.faint, lineHeight: 1.5, marginTop: 4 }}>
              <Info size={12} style={{ verticalAlign: -1 }} /> O projeto guarda a base de dados (até 300 linhas de amostra), colunas e indicadores do dashboard. Fica salvo apenas no seu navegador/conta — nada é compartilhado.
            </div>
          </div>
        </Card>

        <Card style={{ padding: 20 }}>
          <div style={{ fontWeight: 700, fontSize: 13.5, marginBottom: 12, color: T.ink }}>Projetos salvos ({projects.length})</div>
          {loading && <div style={{ color: T.sub, fontSize: 13 }}>Carregando…</div>}
          {!loading && !projects.length && <div style={{ color: T.sub, fontSize: 13 }}>Nenhum projeto salvo ainda.</div>}
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {projects.map((p) => (
              <div key={p.id} style={{ border: `1px solid ${T.border}`, borderRadius: 10, padding: 14, display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10, flexWrap: "wrap" }}>
                <div style={{ flex: 1, minWidth: 200 }}>
                  <div style={{ fontWeight: 700, fontSize: 14, color: T.ink }}>{p.name}</div>
                  {p.description && <div style={{ fontSize: 12.5, color: T.sub, marginTop: 2 }}>{p.description}</div>}
                  <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap" }}>
                    <Pill>{p.fileName}</Pill>
                    <Pill tone="teal">{p.summary?.rowCount} linhas</Pill>
                    <Pill tone="teal">{p.columns?.length} colunas</Pill>
                    <Pill tone={p.summary?.quality >= 80 ? "green" : "amber"}>Qualidade {Math.round(p.summary?.quality || 0)}</Pill>
                  </div>
                  <div style={{ fontSize: 11, color: T.faint, marginTop: 6 }}>Salvo em {new Date(p.createdAt).toLocaleString("pt-BR")}</div>
                </div>
                <div style={{ display: "flex", gap: 8 }}>
                  <Btn variant="subtle" onClick={() => onOpenProject(p)}>Abrir</Btn>
                  <Btn variant="danger" icon={Trash2} onClick={() => remove(p.id)} />
                </div>
              </div>
            ))}
          </div>
        </Card>
      </div>
    </div>
  );
}

/* =========================================================================
   REPORTS TAB
   ========================================================================= */
function buildMarkdownReport({ fileName, d, columns, stats }) {
  const lines = [];
  lines.push(`# Relatório de Análise de Dados — ${fileName}`);
  lines.push(`_Gerado em ${new Date().toLocaleString("pt-BR")}_\n`);
  lines.push(`## Resumo Executivo\n`);
  lines.push(`A base contém **${d.rowCount.toLocaleString("pt-BR")} linhas** e **${d.colCount} colunas**, com um score de qualidade de dados de **${Math.round(d.quality)}/100**. `
    + `Foram identificados **${d.missingTotal} valores ausentes** (${fmt(d.missingPct, 1)}%), **${d.duplicates} registros duplicados** e **${d.outlierTotal} outliers** pelo método IQR.\n`);
  lines.push(`## Relatório Técnico\n`);
  lines.push(`### Colunas e tipos\n`);
  lines.push(`| Coluna | Tipo | % Ausente | Únicos |`);
  lines.push(`|---|---|---|---|`);
  columns.forEach((c) => lines.push(`| ${c.name} | ${c.type} | ${fmt(c.missingPct, 1)}% | ${c.unique} |`));
  if (stats?.length) {
    lines.push(`\n### Estatística descritiva\n`);
    lines.push(`| Variável | N | Média | Mediana | Desvio padrão | CV% | Assimetria | Curtose | Q1 | Q3 |`);
    lines.push(`|---|---|---|---|---|---|---|---|---|---|`);
    stats.forEach((s) => lines.push(`| ${s.name} | ${s.n} | ${fmt(s.mean)} | ${fmt(s.median)} | ${fmt(s.std)} | ${fmt(s.cv, 1)} | ${fmt(s.skew)} | ${fmt(s.kurt)} | ${fmt(s.q1)} | ${fmt(s.q3)} |`));
  }
  lines.push(`## Plano de ação sugerido\n`);
  const actions = [];
  if (d.missingPct > 5) actions.push("Tratar valores ausentes (imputação ou exclusão) antes de modelagens futuras.");
  if (d.duplicates > 0) actions.push("Remover ou investigar registros duplicados.");
  if (d.outlierTotal > 0) actions.push("Investigar outliers identificados — podem ser erros de coleta ou eventos legítimos raros.");
  if (!actions.length) actions.push("Base de dados em boas condições — nenhuma ação de limpeza crítica identificada.");
  actions.forEach((a) => lines.push(`- ${a}`));
  return lines.join("\n");
}

// Gera um script Python/pandas reproduzível a partir das ações registradas
// (testes de hipótese, features criadas, comparações de grupo, regras de qualidade).
function buildPythonScript({ fileName, columns, pipelineLog }) {
  const lines = [];
  lines.push(`"""`);
  lines.push(`Script gerado automaticamente pelo AnálisePro — reproduz a análise feita na interface.`);
  lines.push(`Base: ${fileName}`);
  lines.push(`Gerado em ${new Date().toLocaleString("pt-BR")}`);
  lines.push(`"""`);
  lines.push(`import pandas as pd`);
  lines.push(`import numpy as np`);
  lines.push(`from scipy import stats`);
  lines.push(``);
  lines.push(`# Ajuste o caminho/separador conforme o arquivo original`);
  lines.push(`df = pd.read_csv("${fileName || "dados.csv"}", sep=None, engine="python")`);
  lines.push(``);
  lines.push(`# Tipos inferidos pelo AnálisePro (ajuste se necessário)`);
  columns.filter((c) => !c.derived).forEach((c) => {
    if (c.type === "numeric") lines.push(`df['${c.name}'] = pd.to_numeric(df['${c.name}'], errors='coerce')`);
    else if (c.type === "date") lines.push(`df['${c.name}'] = pd.to_datetime(df['${c.name}'], errors='coerce')`);
  });
  lines.push(``);
  lines.push(`print(df.describe(include='all'))`);
  if (!pipelineLog?.length) {
    lines.push(``);
    lines.push(`# Nenhuma ação adicional registrada ainda — use os botões "Adicionar ao script Python"`);
    lines.push(`# nas abas de Testes Estatísticos, Engenharia de Features, Comparação de Grupos e Regras de Qualidade.`);
  } else {
    pipelineLog.forEach((entry, i) => {
      lines.push(``);
      lines.push(`# ---- ${i + 1}. ${entry.label} ----`);
      lines.push(entry.code);
    });
  }
  return lines.join("\n");
}

/* --- Reports: BI-grade visual primitives (presentation only, no calc changes) --- */
const REPORT_DARK = {
  bg: "#12141A", panel: "#1B1E27", border: "#2A2E3A", ink: "#F1F3F7", sub: "#A6ADBB", faint: "#6E7686",
  teal: "#2DD4C6", tealDark: "#8FEDE3", tealSoft: "rgba(45,212,198,0.14)",
  amber: "#F0B056", amberSoft: "rgba(240,176,86,0.14)",
  red: "#F2726B", redSoft: "rgba(242,114,107,0.14)",
  green: "#57C98B", greenSoft: "rgba(87,201,139,0.14)",
  blue: "#6FA8E0",
};

function KpiCard({ C, icon: Icon, label, value, sub, tone = "neutral", progress }) {
  return (
    <div style={{ background: C.panel, border: `1px solid ${C.border}`, borderRadius: 14, padding: "16px 18px", flex: "1 1 160px", minWidth: 150, boxShadow: "0 1px 2px rgba(16,24,40,0.04)" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ fontSize: 11, color: C.sub, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</div>
        {Icon && <div style={{ width: 26, height: 26, borderRadius: 8, background: tone === "neutral" ? C.tealSoft : (C[`${tone}Soft`] || C.tealSoft), display: "flex", alignItems: "center", justifyContent: "center" }}>
          <Icon size={13} color={tone === "neutral" ? C.teal : (C[tone] || C.teal)} />
        </div>}
      </div>
      <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 25, fontWeight: 700, color: tone === "neutral" ? C.ink : (C[tone] || C.ink), marginTop: 8 }}>{value}</div>
      {sub && <div style={{ fontSize: 11.5, color: C.faint, marginTop: 3 }}>{sub}</div>}
      {progress !== undefined && (
        <div style={{ marginTop: 10, background: C.border, borderRadius: 999, height: 5, overflow: "hidden" }}>
          <div style={{ width: `${Math.max(0, Math.min(100, progress))}%`, height: "100%", background: tone === "neutral" ? C.teal : (C[tone] || C.teal), borderRadius: 999, transition: "width 0.6s ease" }} />
        </div>
      )}
    </div>
  );
}

function Callout({ C, tone = "info", icon: Icon, children }) {
  const map = {
    success: { bg: C.greenSoft, fg: C.green, Ic: CheckCircle2 },
    warning: { bg: C.amberSoft, fg: C.amber, Ic: AlertTriangle },
    info: { bg: C.tealSoft, fg: C.tealDark || C.teal, Ic: Info },
  };
  const cfg = map[tone] || map.info;
  const IconC = Icon || cfg.Ic;
  return (
    <div style={{ display: "flex", gap: 10, alignItems: "flex-start", background: cfg.bg, borderRadius: 10, padding: "12px 14px" }}>
      <IconC size={16} color={cfg.fg} style={{ flexShrink: 0, marginTop: 1 }} />
      <div style={{ fontSize: 13, color: C.ink, lineHeight: 1.55 }}>{children}</div>
    </div>
  );
}

function ReportDivider({ C, label }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 12, margin: "26px 0 18px" }}>
      <div style={{ flex: 1, height: 1, background: C.border }} />
      {label && <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: C.faint, fontFamily: "'JetBrains Mono', monospace" }}>{label}</div>}
      <div style={{ flex: 1, height: 1, background: C.border }} />
    </div>
  );
}

function ReportAccordion({ C, icon: Icon, title, subtitle, defaultOpen = false, children }) {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <div className="rpt-accordion" style={{ background: C.panel, border: `1px solid ${C.border}`, borderRadius: 14, marginBottom: 14, overflow: "hidden", boxShadow: "0 1px 2px rgba(16,24,40,0.04)" }}>
      <button onClick={() => setOpen((o) => !o)} className="no-print" style={{ width: "100%", display: "flex", alignItems: "center", justifyContent: "space-between", padding: "16px 20px", background: "transparent", border: "none", cursor: "pointer", textAlign: "left" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          {Icon && <div style={{ width: 34, height: 34, borderRadius: 9, background: C.tealSoft, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}><Icon size={16} color={C.teal} /></div>}
          <div>
            <div style={{ fontWeight: 700, fontSize: 14, color: C.ink }}>{title}</div>
            {subtitle && <div style={{ fontSize: 11.5, color: C.faint, marginTop: 2 }}>{subtitle}</div>}
          </div>
        </div>
        <ChevronDown size={17} color={C.faint} style={{ transform: open ? "rotate(180deg)" : "rotate(0deg)", transition: "transform 0.25s ease" }} />
      </button>
      <div className="rpt-accordion-body" style={{ display: "grid", gridTemplateRows: open ? "1fr" : "0fr", transition: "grid-template-rows 0.3s ease" }}>
        <div style={{ overflow: "hidden" }}>
          <div style={{ padding: "0 20px 20px", borderTop: `1px solid ${C.border}`, paddingTop: 16 }}>{children}</div>
        </div>
      </div>
    </div>
  );
}

function TimelineStep({ C, icon: Icon, title, description, status }) {
  const tone = status === "done" ? "green" : status === "warn" ? "amber" : "neutral";
  const dotColor = tone === "neutral" ? C.faint : C[tone];
  return (
    <div style={{ display: "flex", gap: 14 }}>
      <div style={{ display: "flex", flexDirection: "column", alignItems: "center" }}>
        <div style={{ width: 30, height: 30, borderRadius: 999, background: tone === "neutral" ? C.border : (C[`${tone}Soft`] || C.tealSoft), display: "flex", alignItems: "center", justifyContent: "center", border: `2px solid ${dotColor}`, flexShrink: 0 }}>
          <Icon size={13} color={dotColor} />
        </div>
        <div style={{ flex: 1, width: 2, background: C.border, marginTop: 2, marginBottom: 2 }} />
      </div>
      <div style={{ paddingBottom: 22 }}>
        <div style={{ fontWeight: 700, fontSize: 13.5, color: C.ink }}>{title}</div>
        <div style={{ fontSize: 12, color: C.sub, marginTop: 2, lineHeight: 1.5 }}>{description}</div>
      </div>
    </div>
  );
}

/* --- Reports: export helpers (formatting only — same source data/calcs) --- */
function exportExcelReport({ fileName, d, columns, stats }) {
  const wb = XLSX.utils.book_new();
  const resumo = [
    ["Indicador", "Valor"],
    ["Arquivo", fileName],
    ["Linhas", d.rowCount],
    ["Colunas", d.colCount],
    ["Qualidade dos dados (0-100)", Math.round(d.quality)],
    ["Valores ausentes (%)", fmt(d.missingPct, 1)],
    ["Duplicados", d.duplicates],
    ["Outliers (IQR)", d.outlierTotal],
    ["Gerado em", new Date().toLocaleString("pt-BR")],
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumo), "Resumo");
  const colRows = [["Coluna", "Tipo", "% Ausente", "Únicos"], ...columns.map((c) => [c.name, c.type, fmt(c.missingPct, 1), c.unique])];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(colRows), "Colunas");
  if (stats?.length) {
    const statRows = [["Variável", "N", "Média", "Mediana", "Desvio padrão", "CV%", "Assimetria", "Curtose", "Q1", "Q3"],
      ...stats.map((s) => [s.name, s.n, fmt(s.mean), fmt(s.median), fmt(s.std), fmt(s.cv, 1), fmt(s.skew), fmt(s.kurt), fmt(s.q1), fmt(s.q3)])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(statRows), "Estatistica");
  }
  XLSX.writeFile(wb, `relatorio_${(fileName || "analise").replace(/\.[^.]+$/, "").replace(/[^\w\-]+/g, "_")}.xlsx`);
}

function buildHtmlReport(markdown, fileName) {
  const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const lines = markdown.split("\n");
  let html = "", inTable = false, inList = false;
  lines.forEach((line) => {
    const bolded = esc(line).replace(/\*\*(.+?)\*\*/g, "<b>$1</b>");
    if (line.startsWith("# ")) html += `<h1>${bolded.slice(2)}</h1>`;
    else if (line.startsWith("## ")) html += `<h2>${bolded.slice(3)}</h2>`;
    else if (line.startsWith("### ")) html += `<h3>${bolded.slice(4)}</h3>`;
    else if (line.startsWith("|")) {
      if (!inTable) { html += "<table>"; inTable = true; }
      const cells = line.split("|").slice(1, -1);
      if (!cells.every((c) => /^-+$/.test(c.trim()))) html += "<tr>" + cells.map((c) => `<td>${esc(c.trim())}</td>`).join("") + "</tr>";
    } else {
      if (inTable) { html += "</table>"; inTable = false; }
      if (line.startsWith("- ")) { if (!inList) { html += "<ul>"; inList = true; } html += `<li>${bolded.slice(2)}</li>`; }
      else {
        if (inList) { html += "</ul>"; inList = false; }
        if (line.startsWith("_")) html += `<p class="muted">${bolded.replaceAll("_", "")}</p>`;
        else if (!line.trim()) html += "<br/>";
        else html += `<p>${bolded}</p>`;
      }
    }
  });
  if (inTable) html += "</table>";
  if (inList) html += "</ul>";
  return `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Relatório — ${fileName}</title>
<style>
body{font-family:Inter,Arial,sans-serif;color:#12141C;max-width:900px;margin:40px auto;padding:0 20px;line-height:1.6;background:#F6F7F9;}
h1{font-size:28px;margin-bottom:4px;} h2{color:#0B6E6E;margin-top:32px;} h3{margin-top:20px;}
table{border-collapse:collapse;width:100%;margin:12px 0;font-size:13px;background:#fff;}
td{border:1px solid #E3E6EB;padding:6px 10px;} .muted{color:#9AA1AE;font-size:13px;}
</style></head><body>${html}</body></html>`;
}

/* =========================================================================
   REPORTS TAB — BI-grade visual redesign (Power BI / Tableau style)
   Calculations, stats and test logic are 100% unchanged from before.
   ========================================================================= */
function ReportsTab({ rows, columns, fileName, pipelineLog, notes }) {
  const [theme, setTheme] = useState("light");
  const C = theme === "dark" ? REPORT_DARK : T;
  const d = useMemo(() => computeDashboard(rows, columns), [rows, columns]);
  const insights = useMemo(() => buildExecutiveInsights(d, columns), [d, columns]);
  const numericCols = columns.filter((c) => c.type === "numeric");
  const stats = useMemo(() => numericCols.map((c) => {
    const vals = rows.map((r) => r[c.name]).filter(isNum);
    return { name: c.name, n: vals.length, mean: mean(vals), median: median(vals), std: std(vals), cv: cv(vals), skew: skewness(vals), kurt: kurtosisExcess(vals), ...quartiles(vals) };
  }), [rows, numericCols.map((c) => c.name).join(",")]);

  const markdown = useMemo(() => buildMarkdownReport({ fileName, d, columns, stats }), [fileName, d, columns, stats]);
  const pythonScript = useMemo(() => buildPythonScript({ fileName, columns, pipelineLog }), [fileName, columns, pipelineLog]);

  const insightToneMap = { amber: "warning", green: "success", teal: "info", neutral: "info", red: "warning" };
  const relevantCorrelations = d.topCorrelations.filter((p) => Math.abs(p.r) > 0.5);

  const actions = [];
  if (d.missingPct > 5) actions.push("Tratar valores ausentes (imputação ou exclusão) antes de modelagens futuras.");
  if (d.duplicates > 0) actions.push("Remover ou investigar registros duplicados.");
  if (d.outlierTotal > 0) actions.push("Investigar outliers identificados — podem ser erros de coleta ou eventos legítimos raros.");
  if (d.pareto) actions.push(`Priorizar as ${d.pareto.categoriesFor80} categorias de "${d.pareto.column}" que concentram 80% dos registros.`);
  if (!actions.length) actions.push("Base de dados em boas condições — nenhuma ação de limpeza crítica identificada.");

  return (
    <div style={{ background: theme === "dark" ? C.bg : "transparent", margin: theme === "dark" ? -28 : 0, padding: theme === "dark" ? 28 : 0, borderRadius: theme === "dark" ? 14 : 0, transition: "background 0.25s ease" }}>
      <style>{`
        @media print {
          .rpt-accordion-body { grid-template-rows: 1fr !important; }
        }
        .rpt-fade { animation: rptFadeIn 0.4s ease; }
        @keyframes rptFadeIn { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; transform: translateY(0); } }
      `}</style>

      {/* HEADER */}
      <div className="rpt-fade" style={{ background: C.panel, border: `1px solid ${C.border}`, borderRadius: 16, padding: "22px 24px", marginBottom: 18, display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 16 }}>
        <div style={{ display: "flex", gap: 14, alignItems: "flex-start" }}>
          <div style={{ width: 44, height: 44, borderRadius: 12, background: C.tealSoft, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
            <FileText size={21} color={C.teal} />
          </div>
          <div>
            <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 11, letterSpacing: "0.1em", color: C.teal, textTransform: "uppercase", fontWeight: 700 }}>Relatório executivo</div>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 24, fontWeight: 700, color: C.ink, marginTop: 2 }}>Análise de Dados — {fileName}</div>
            <div style={{ fontSize: 12.5, color: C.sub, marginTop: 4 }}>
              {d.rowCount.toLocaleString("pt-BR")} linhas · {d.colCount} colunas · gerado em {new Date().toLocaleString("pt-BR")}
            </div>
          </div>
        </div>
        <div className="no-print" style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
          <button onClick={() => setTheme((t) => (t === "light" ? "dark" : "light"))}
            title="Alternar tema do relatório"
            style={{ width: 36, height: 36, borderRadius: 9, border: `1px solid ${C.border}`, background: "transparent", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>
            {theme === "light" ? <Moon size={15} color={C.ink} /> : <Sun size={15} color={C.ink} />}
          </button>
          <Btn variant="ghost" icon={Download} style={{ background: "transparent", color: C.ink, borderColor: C.border }} onClick={() => downloadBlob(markdown, "relatorio_analise.md", "text/markdown")}>Markdown</Btn>
          <Btn variant="ghost" icon={FileSpreadsheet} style={{ background: "transparent", color: C.ink, borderColor: C.border }} onClick={() => exportExcelReport({ fileName, d, columns, stats })}>Excel</Btn>
          <Btn variant="ghost" icon={FileCode} style={{ background: "transparent", color: C.ink, borderColor: C.border }} onClick={() => downloadBlob(buildHtmlReport(markdown, fileName), "relatorio_analise.html", "text/html")}>HTML</Btn>
          <Btn variant="ghost" icon={FileCode} style={{ background: "transparent", color: C.ink, borderColor: C.border }} onClick={() => downloadBlob(pythonScript, "analise_reproduzivel.py", "text/x-python")}>Python</Btn>
          <Btn icon={Printer} onClick={() => window.print()}>PDF</Btn>
        </div>
      </div>

      {/* KPIs */}
      <div className="rpt-fade" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 12, marginBottom: 18 }}>
        <KpiCard C={C} icon={Layers} label="Registros" value={d.rowCount.toLocaleString("pt-BR")} sub={`${d.colCount} variáveis`} />
        <KpiCard C={C} icon={Gauge} label="Qualidade dos dados" value={`${Math.round(d.quality)}/100`} tone={d.quality >= 80 ? "green" : d.quality >= 55 ? "amber" : "red"} progress={d.quality} />
        <KpiCard C={C} icon={TrendingUp} label="Variáveis numéricas" value={numericCols.length} sub={`de ${d.colCount} colunas`} />
        <KpiCard C={C} icon={AlertTriangle} label="Outliers (IQR)" value={d.outlierTotal} tone={d.outlierTotal > 0 ? "amber" : "green"} />
        <KpiCard C={C} icon={Activity} label="Correlações relevantes" value={relevantCorrelations.length} sub="|r| > 0.5" tone={relevantCorrelations.length ? "teal" : "neutral"} />
      </div>

      {/* RESUMO EXECUTIVO */}
      <div className="rpt-fade" style={{ background: C.panel, border: `1px solid ${C.border}`, borderRadius: 16, padding: 22, marginBottom: 4 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 9, marginBottom: 14 }}>
          <Sparkles size={16} color={C.teal} />
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 16, color: C.ink }}>Resumo executivo</div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {insights.map((ins, i) => (
            <Callout key={i} C={C} tone={insightToneMap[ins.tone] || "info"}>{ins.text}</Callout>
          ))}
        </div>
      </div>

      <ReportDivider C={C} label="Linha do tempo da análise" />

      <div className="rpt-fade" style={{ background: C.panel, border: `1px solid ${C.border}`, borderRadius: 16, padding: "22px 24px 4px", marginBottom: 4 }}>
        <TimelineStep C={C} icon={Upload} status="done" title="Importação dos dados"
          description={`Arquivo "${fileName}" carregado com ${d.rowCount.toLocaleString("pt-BR")} linhas e ${d.colCount} colunas.`} />
        <TimelineStep C={C} icon={Gauge} status={d.quality >= 70 ? "done" : "warn"} title="Diagnóstico de qualidade"
          description={`Score de qualidade ${Math.round(d.quality)}/100 — ${fmt(d.missingPct, 1)}% de ausentes, ${d.duplicates} duplicado(s), ${d.outlierTotal} outlier(s).`} />
        <TimelineStep C={C} icon={BarChart3} status="done" title="Exploração de dados (EDA)"
          description={`Distribuições, correlações e séries temporais analisadas. ${d.topCorrelations.length ? `Correlação mais forte: ${d.topCorrelations[0].a} × ${d.topCorrelations[0].b} (r=${fmt(d.topCorrelations[0].r, 2)}).` : "Sem colunas numéricas suficientes para correlação."}`} />
        <TimelineStep C={C} icon={Lightbulb} status="done" title="Conclusões e recomendações"
          description="Plano de ação sugerido gerado automaticamente a partir dos indicadores acima." />
      </div>

      <ReportDivider C={C} label="Detalhamento técnico" />

      <ReportAccordion C={C} icon={Layers} title="Colunas & tipos de dados" subtitle={`${d.colCount} colunas identificadas`}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
            <thead>
              <tr style={{ textAlign: "left", color: C.sub, borderBottom: `1px solid ${C.border}` }}>
                <th style={{ padding: "8px 10px" }}>Coluna</th><th style={{ padding: "8px 10px" }}>Tipo</th>
                <th style={{ padding: "8px 10px" }}>% Ausente</th><th style={{ padding: "8px 10px" }}>Únicos</th>
              </tr>
            </thead>
            <tbody>
              {columns.map((c) => (
                <tr key={c.name} style={{ borderBottom: `1px solid ${C.border}` }}>
                  <td style={{ padding: "8px 10px", fontWeight: 600, color: C.ink, fontFamily: "'JetBrains Mono', monospace" }}>{c.name}</td>
                  <td style={{ padding: "8px 10px" }}><Pill tone={c.type === "numeric" ? "teal" : c.type === "date" ? "amber" : "neutral"}>{c.type}</Pill></td>
                  <td style={{ padding: "8px 10px", color: c.missingPct > 5 ? C.amber : C.sub }}>{fmt(c.missingPct, 1)}%</td>
                  <td style={{ padding: "8px 10px", color: C.sub }}>{c.unique}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </ReportAccordion>

      <ReportAccordion C={C} icon={Sigma} title="Estatística descritiva" subtitle={`${stats.length} variável(is) numérica(s)`}>
        {stats.length ? (
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, minWidth: 780 }}>
              <thead>
                <tr style={{ textAlign: "left", color: C.sub, borderBottom: `1px solid ${C.border}` }}>
                  {["Variável", "N", "Média", "Mediana", "Desvio padrão", "CV%", "Assimetria", "Curtose", "Q1", "Q3"].map((h) => <th key={h} style={{ padding: "8px 10px", whiteSpace: "nowrap" }}>{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {stats.map((s) => (
                  <tr key={s.name} style={{ borderBottom: `1px solid ${C.border}` }}>
                    <td style={{ padding: "8px 10px", fontWeight: 700, color: C.ink, fontFamily: "'JetBrains Mono', monospace" }}>{s.name}</td>
                    <td style={{ padding: "8px 10px", color: C.ink }}>{s.n}</td>
                    <td style={{ padding: "8px 10px", color: C.ink }}>{fmt(s.mean)}</td>
                    <td style={{ padding: "8px 10px", color: C.ink }}>{fmt(s.median)}</td>
                    <td style={{ padding: "8px 10px", color: C.ink }}>{fmt(s.std)}</td>
                    <td style={{ padding: "8px 10px", color: C.ink }}>{fmt(s.cv, 1)}</td>
                    <td style={{ padding: "8px 10px", color: C.ink }}>{fmt(s.skew)}</td>
                    <td style={{ padding: "8px 10px", color: C.ink }}>{fmt(s.kurt)}</td>
                    <td style={{ padding: "8px 10px", color: C.ink }}>{fmt(s.q1)}</td>
                    <td style={{ padding: "8px 10px", color: C.ink }}>{fmt(s.q3)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : <div style={{ color: C.sub, fontSize: 13 }}>Nenhuma coluna numérica encontrada.</div>}
      </ReportAccordion>

      <ReportDivider C={C} label="Conclusões & recomendações" />

      <div className="rpt-fade" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))", gap: 16, marginBottom: 8 }}>
        <div style={{ background: C.panel, border: `1px solid ${C.border}`, borderRadius: 16, padding: 20 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
            <Lightbulb size={16} color={C.amber} />
            <div style={{ fontWeight: 700, fontSize: 14, color: C.ink }}>Recomendações de negócio</div>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {actions.map((a, i) => (
              <div key={i} style={{ display: "flex", gap: 8, fontSize: 12.5, color: C.ink, lineHeight: 1.5 }}>
                <span style={{ color: C.amber, fontWeight: 700 }}>{i + 1}.</span>{a}
              </div>
            ))}
          </div>
        </div>
      </div>

      {pipelineLog?.length > 0 && (
        <>
          <ReportDivider C={C} label="Código Python reprodutível" />
          <div className="rpt-fade no-print" style={{ background: C.panel, border: `1px solid ${C.border}`, borderRadius: 16, padding: 20, marginBottom: 8 }}>
            <div style={{ fontSize: 12.5, color: C.sub, marginBottom: 12 }}>
              {pipelineLog.length} ação(ões) registrada(s) nas abas de Testes Estatísticos, Engenharia de Features, Comparação de Grupos e Regras de Qualidade. Baixe o script completo pelo botão "Python" no topo da página.
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {pipelineLog.map((entry, i) => (
                <div key={entry.id || i} style={{ fontSize: 12, fontFamily: "'JetBrains Mono', monospace", color: C.ink, background: C.bg, padding: "6px 10px", borderRadius: 6 }}>
                  {i + 1}. {entry.label}
                </div>
              ))}
            </div>
          </div>
        </>
      )}

      {notes?.length > 0 && (
        <>
          <ReportDivider C={C} label="Notas de análise" />
          <div className="rpt-fade" style={{ background: C.panel, border: `1px solid ${C.border}`, borderRadius: 16, padding: 20, marginBottom: 8 }}>
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {notes.map((n) => (
                <div key={n.id} style={{ fontSize: 12.5, color: C.ink, borderLeft: `3px solid ${C.teal}`, paddingLeft: 10 }}>
                  <b>{n.variable}</b>: {n.text}
                  <div style={{ fontSize: 11, color: C.faint, marginTop: 2 }}>{new Date(n.createdAt).toLocaleString("pt-BR")}</div>
                </div>
              ))}
            </div>
          </div>
        </>
      )}

      <div style={{ marginTop: 18, fontSize: 11.5, color: C.faint, textAlign: "center" }}>
        Relatório gerado automaticamente pela AnálisePro · {new Date().toLocaleString("pt-BR")}
      </div>
    </div>
  );
}

/* =========================================================================
   BIBLIOTECAS DE ANÁLISE AUTOMÁTICA (backend Python)
   Cada aba consome a mesma base carregada uma única vez (backendSessionId),
   via a API Python hospedada separadamente (ver /backend no repositório).
   ========================================================================= */
function LibraryStatusNotice({ backendStatus }) {
  if (backendStatus === "uploading") {
    return (
      <Card style={{ padding: 24, textAlign: "center", color: T.sub, fontSize: 13.5 }}>
        Enviando a base para o servidor de análise (Python)…
      </Card>
    );
  }
  if (backendStatus === "error" || backendStatus === "idle") {
    return (
      <div style={{ background: T.amberSoft, color: T.amber, padding: "14px 16px", borderRadius: 10, fontSize: 13.5, display: "flex", gap: 8, alignItems: "flex-start" }}>
        <AlertTriangle size={16} style={{ flexShrink: 0, marginTop: 1 }} />
        Não foi possível conectar ao servidor de análise em Python para esta biblioteca. As demais abas do AnálisePro continuam funcionando normalmente — tente novamente carregando a base outra vez.
      </div>
    );
  }
  return null;
}

// Abas cujo backend devolve HTML pronto (o navegador carrega direto via <iframe src>).
function IframeReportTab({ eyebrow, title, subtitle, endpoint, sessionId, backendStatus }) {
  const [loaded, setLoaded] = useState(false);
  useEffect(() => { setLoaded(false); }, [sessionId]);
  return (
    <div>
      <SectionTitle eyebrow={eyebrow} title={title} right={sessionId && (
        <Btn variant="ghost" icon={ExternalLink} onClick={() => window.open(`${API_BASE}/api/${endpoint}/${sessionId}`, "_blank")}>Abrir em nova aba</Btn>
      )} />
      {subtitle && <div style={{ fontSize: 12.5, color: T.sub, marginTop: -10, marginBottom: 16 }}>{subtitle}</div>}
      {!sessionId ? <LibraryStatusNotice backendStatus={backendStatus} /> : (
        <Card style={{ padding: 0, overflow: "hidden", minHeight: 560 }}>
          {!loaded && <div style={{ padding: 24, textAlign: "center", color: T.sub, fontSize: 13.5 }}>Gerando relatório… isso pode levar alguns segundos na primeira vez.</div>}
          <iframe title={title} src={`${API_BASE}/api/${endpoint}/${sessionId}`} onLoad={() => setLoaded(true)}
            style={{ width: "100%", height: 720, border: "none", display: loaded ? "block" : "none" }} />
        </Card>
      )}
    </div>
  );
}

function YdataProfilingTab({ sessionId, backendStatus }) {
  return <IframeReportTab eyebrow="ydata-profiling" title="Relatório completo de EDA" subtitle="Perfil estatístico completo da base — tipos, distribuições, correlações, valores ausentes e alertas de qualidade." endpoint="ydata-profiling" sessionId={sessionId} backendStatus={backendStatus} />;
}
function SweetvizTab({ sessionId, backendStatus }) {
  return <IframeReportTab eyebrow="Sweetviz" title="Relatório visual e comparação de variáveis" subtitle="Visão visual das variáveis da base, com histogramas, associações e comparação entre colunas." endpoint="sweetviz" sessionId={sessionId} backendStatus={backendStatus} />;
}
/* =========================================================================
   MAIN APP
   ========================================================================= */
const NAV = [
  { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { id: "eda", label: "Exploração (EDA)", icon: BarChart3 },
  { id: "stats", label: "Estatística Descritiva", icon: Sigma },
  { id: "hypothesis", label: "Testes Estatísticos", icon: Beaker },
  { id: "features", label: "Engenharia de Features", icon: Wand2 },
  { id: "groups", label: "Comparação de Grupos", icon: Table2 },
  { id: "quality", label: "Regras de Qualidade", icon: ListChecks },
  { id: "ydata", label: "ydata-profiling", icon: FileBarChart2 },
  { id: "sweetviz", label: "Sweetviz", icon: GitCompare },
  { id: "projects", label: "Projetos", icon: FolderKanban },
  { id: "reports", label: "Relatórios", icon: FileText },
];

const NOTES_KEY = "analisepro_notes_v1";
function loadNotes() { try { return JSON.parse(localStorage.getItem(NOTES_KEY) || "[]"); } catch { return []; } }
function saveNotes(notes) { try { localStorage.setItem(NOTES_KEY, JSON.stringify(notes)); } catch { /* noop */ } }

function AnalisePro() {
  const [columns, setColumns] = useState([]);
  const [rows, setRows] = useState([]);
  const [fileName, setFileName] = useState("");
  const [activeTab, setActiveTab] = useState("dashboard");

  const [backendSessionId, setBackendSessionId] = useState(null);
  const [backendStatus, setBackendStatus] = useState("idle"); // idle | uploading | ready | error

  const [pipelineLog, setPipelineLog] = useState([]);
  const logAction = useCallback((entry) => {
    setPipelineLog((prev) => [...prev, { id: uid(), at: new Date().toISOString(), ...entry }]);
  }, []);

  const [notes, setNotes] = useState(loadNotes);
  useEffect(() => { saveNotes(notes); }, [notes]);
  const addNote = useCallback((variable, text) => {
    setNotes((prev) => [...prev, { id: uid(), variable, text, createdAt: new Date().toISOString() }]);
  }, []);
  const removeNote = useCallback((id) => setNotes((prev) => prev.filter((n) => n.id !== id)), []);

  const handleLoaded = useCallback((data, name, file) => {
    const cols = inferColumns(data);
    const typed = buildTypedRows(data, cols);
    setColumns(cols); setRows(typed); setFileName(name);
    setActiveTab("dashboard");

    // Envia o arquivo original ao backend Python para alimentar as abas de
    // bibliotecas (ydata-profiling, Sweetviz). Isso é independente da
    // análise client-side acima — se o backend falhar ou estiver
    // indisponível, o restante do app continua funcionando normalmente.
    setBackendSessionId(null);
    if (file) {
      setBackendStatus("uploading");
      const form = new FormData();
      form.append("file", file);
      fetch(`${API_BASE}/api/upload`, { method: "POST", body: form })
        .then((res) => { if (!res.ok) throw new Error("Upload falhou (" + res.status + ")"); return res.json(); })
        .then((json) => { setBackendSessionId(json.session_id); setBackendStatus("ready"); })
        .catch(() => setBackendStatus("error"));
    } else {
      setBackendStatus("error");
    }
  }, []);

  const handleOpenProject = useCallback((p) => {
    const cols = p.columns || [];
    const typed = (p.sampleRows || []).map((r) => {
      const out = {};
      cols.forEach((c) => {
        const v = r[c.name];
        out[c.name] = c.type === "date" && v ? new Date(v) : v;
      });
      return out;
    });
    setColumns(cols); setRows(typed); setFileName(p.fileName + " (projeto: " + p.name + ")");
    setActiveTab("dashboard");
  }, []);

  const hasData = rows.length > 0;

  return (
    <div style={{ fontFamily: "'Inter', sans-serif", background: T.bg, minHeight: 640, color: T.ink, display: "flex", borderRadius: 14, overflow: "hidden", border: `1px solid ${T.border}` }}>
      <style>{FONT_IMPORT}{`
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { height: 8px; width: 8px; }
        ::-webkit-scrollbar-thumb { background: #D7DBE2; border-radius: 8px; }
        @media print { .no-print { display: none !important; } }
      `}</style>

      {/* SIDEBAR */}
      <div className="no-print" style={{ width: 220, background: "#FFFFFF", borderRight: `1px solid ${T.border}`, padding: "20px 14px", display: "flex", flexDirection: "column", flexShrink: 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "0 8px", marginBottom: 24 }}>
          <div style={{ width: 30, height: 30, borderRadius: 8, background: T.teal, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Activity size={16} color="#fff" />
          </div>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 15, color: T.ink }}>Análise<span style={{ color: T.teal }}>Pro</span></div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
          {NAV.map((n) => {
            const active = activeTab === n.id;
            const disabled = !hasData && n.id !== "projects";
            return (
              <button key={n.id} disabled={disabled} onClick={() => setActiveTab(n.id)}
                style={{
                  display: "flex", alignItems: "center", gap: 10, padding: "9px 12px", borderRadius: 9, border: "none",
                  background: active ? T.tealSoft : "transparent", color: active ? T.tealDark : disabled ? T.faint : T.ink,
                  fontWeight: active ? 700 : 500, fontSize: 13.5, cursor: disabled ? "not-allowed" : "pointer", textAlign: "left", opacity: disabled ? 0.5 : 1,
                }}>
                <n.icon size={16} /> {n.label}
              </button>
            );
          })}
        </div>
        {hasData && (
          <div style={{ marginTop: "auto", paddingTop: 16, borderTop: `1px solid ${T.border}` }}>
            <div style={{ fontSize: 11, color: T.faint, fontWeight: 600, textTransform: "uppercase" }}>Base ativa</div>
            <div style={{ fontSize: 12.5, color: T.ink, fontWeight: 600, marginTop: 4, wordBreak: "break-word" }}>{fileName}</div>
            <div style={{ fontSize: 11.5, color: T.sub, marginTop: 2 }}>{rows.length.toLocaleString("pt-BR")} linhas · {columns.length} colunas</div>
            <Btn variant="ghost" style={{ marginTop: 10, width: "100%", justifyContent: "center" }} onClick={() => { setRows([]); setColumns([]); setFileName(""); setActiveTab("dashboard"); setBackendSessionId(null); setBackendStatus("idle"); }}>Trocar base</Btn>
          </div>
        )}
      </div>

      {/* MAIN CONTENT */}
      <div style={{ flex: 1, padding: 28, overflow: "auto", maxHeight: "90vh" }}>
        {!hasData && activeTab !== "projects" && <UploadView onLoaded={handleLoaded} />}
        {!hasData && activeTab === "projects" && (
          <ProjectsTab rows={[]} columns={[]} fileName="" onOpenProject={handleOpenProject} />
        )}
        {hasData && activeTab === "dashboard" && <DashboardTab rows={rows} columns={columns} fileName={fileName} />}
        {hasData && activeTab === "eda" && <EDATab rows={rows} columns={columns} notes={notes} onAddNote={addNote} onRemoveNote={removeNote} />}
        {hasData && activeTab === "stats" && <DescriptiveTab rows={rows} columns={columns} />}
        {hasData && activeTab === "hypothesis" && <HypothesisTestsTab rows={rows} columns={columns} onLogAction={logAction} />}
        {hasData && activeTab === "features" && <FeatureEngineeringTab rows={rows} columns={columns} setRows={setRows} setColumns={setColumns} onLogAction={logAction} />}
        {hasData && activeTab === "groups" && <GroupComparisonTab rows={rows} columns={columns} onLogAction={logAction} />}
        {hasData && activeTab === "quality" && <DataQualityTab rows={rows} columns={columns} onLogAction={logAction} />}
        {hasData && activeTab === "ydata" && <YdataProfilingTab sessionId={backendSessionId} backendStatus={backendStatus} />}
        {hasData && activeTab === "sweetviz" && <SweetvizTab sessionId={backendSessionId} backendStatus={backendStatus} />}
        {hasData && activeTab === "projects" && <ProjectsTab rows={rows} columns={columns} fileName={fileName} onOpenProject={handleOpenProject} />}
        {hasData && activeTab === "reports" && <ReportsTab rows={rows} columns={columns} fileName={fileName} pipelineLog={pipelineLog} notes={notes} />}
      </div>
    </div>
  );
}

export default AnalisePro;
