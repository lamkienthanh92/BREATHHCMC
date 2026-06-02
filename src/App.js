import { useState, useCallback, useMemo } from "react";
import * as XLSX from "xlsx";
import { ROADS, ROAD_NAMES } from "./roads.js";
import {
  ROAD_PM25,
  lookupPM25,
  getPM25Risk,
  calcExposureScore,
  getScoreRisk,
} from "./air.js";

// ── Geometry ──────────────────────────────────────────────────────────
function haversine(a, b, c, d) {
  const R = 6371000,
    dL = ((c - a) * Math.PI) / 180,
    dG = ((d - b) * Math.PI) / 180;
  const x =
    Math.sin(dL / 2) ** 2 +
    Math.cos((a * Math.PI) / 180) *
      Math.cos((c * Math.PI) / 180) *
      Math.sin(dG / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(x), Math.sqrt(1 - x));
}
function p2s(pA, pB, aA, aB, bA, bB) {
  const dx = bB - aB,
    dy = bA - aA,
    ls = dx * dx + dy * dy;
  let t = ls === 0 ? 0 : ((pB - aB) * dx + (pA - aA) * dy) / ls;
  t = Math.max(0, Math.min(1, t));
  return haversine(pA, pB, aA + t * dy, aB + t * dx);
}
function distToRoad(pA, pB, road) {
  const { segs, halfWidth: h } = road;
  if (!segs || !segs.length) return Infinity;
  let m = Infinity;
  for (const seg of segs) {
    for (let i = 0; i < seg.length - 1; i++) {
      const d = p2s(pA, pB, seg[i][0], seg[i][1], seg[i + 1][0], seg[i + 1][1]);
      if (d < m) m = d;
    }
  }
  return Math.max(0, Math.round(m - h));
}

// ── Helpers ──────────────────────────────────────────────────────────
function getRoadRisk(d) {
  if (d < 50)
    return { label: "High", color: "#9B1C1C", bg: "#FEF2F2", bar: "#EF4444" };
  if (d < 200)
    return { label: "Med", color: "#92400E", bg: "#FFFBEB", bar: "#F59E0B" };
  return { label: "Low", color: "#166534", bg: "#F0FDF4", bar: "#22C55E" };
}
function parseCoord(s) {
  if (!s?.trim()) return null;
  const m = s.trim().match(/(-?\d+\.?\d*)[,\s]+(-?\d+\.?\d*)/);
  if (!m) return null;
  const a = parseFloat(m[1]),
    b = parseFloat(m[2]);
  if (isNaN(a) || isNaN(b) || a < 10 || a > 11 || b < 106 || b > 107)
    return null;
  return { lat: a, lng: b };
}

// ── Excel export ─────────────────────────────────────────────────────
function exportExcel(sessions) {
  const wb = XLSX.utils.book_new();

  // Sheet 1 — patient results
  const h1 = [
    "#",
    "Patient / Sample ID",
    "Latitude",
    "Longitude",
    "Nearest Road",
    "Min Distance (m)",
    "Home PM2.5 (µg/m³)",
    "PM2.5 Home Category",
    "Exposure Score",
    "Overall Risk Level",
    "Measured At",
    ...ROAD_NAMES.map((n) => `Dist_${n} (m)`),
    ...ROAD_NAMES.map((n) => `PM2.5_${n} (µg/m³)`),
  ];
  const r1 = sessions.map((s, i) => {
    const dm = Object.fromEntries(s.dists.map((d) => [d.name, d.dist]));
    return [
      i + 1,
      s.patientName || "",
      s.geo.lat,
      s.geo.lng,
      s.closest.name,
      s.minDist,
      s.pm25 ?? "N/A",
      getPM25Risk(s.pm25).label,
      s.score,
      getScoreRisk(s.score).label,
      s.timestamp,
      ...ROAD_NAMES.map((n) => dm[n] ?? ""),
      ...ROAD_NAMES.map((n) => ROAD_PM25[n] ?? ""),
    ];
  });
  const ws1 = XLSX.utils.aoa_to_sheet([h1, ...r1]);
  ws1["!cols"] = [
    { wch: 4 },
    { wch: 22 },
    { wch: 12 },
    { wch: 12 },
    { wch: 26 },
    { wch: 14 },
    { wch: 20 },
    { wch: 28 },
    { wch: 16 },
    { wch: 26 },
    { wch: 22 },
    ...ROAD_NAMES.map(() => ({ wch: 16 })),
    ...ROAD_NAMES.map(() => ({ wch: 18 })),
  ];
  XLSX.utils.book_append_sheet(wb, ws1, "Patient Results");

  // Sheet 2 — PM2.5 per road
  const h2 = ["Road Name", "PM2.5 centroid (µg/m³)", "WHO 2021 Category"];
  const r2 = ROAD_NAMES.map((n) => {
    const v = ROAD_PM25[n];
    return [n, v ?? "N/A", getPM25Risk(v).label];
  });
  const ws2 = XLSX.utils.aoa_to_sheet([h2, ...r2]);
  ws2["!cols"] = [{ wch: 32 }, { wch: 22 }, { wch: 28 }];
  XLSX.utils.book_append_sheet(wb, ws2, "PM2.5 by Road");

  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const url = URL.createObjectURL(
    new Blob([out], { type: "application/octet-stream" })
  );
  const a = document.createElement("a");
  a.href = url;
  a.download = `copd_hcmc_${new Date().toISOString().slice(0, 10)}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

// ── Styles ────────────────────────────────────────────────────────────
const card = {
  background: "#fff",
  border: "1px solid #E5E7EB",
  borderRadius: 14,
  padding: "1rem 1.2rem",
};
const inp = {
  width: "100%",
  padding: "10px 13px",
  fontSize: 14,
  border: "1px solid #E5E7EB",
  borderRadius: 9,
  background: "#FAFAFA",
  color: "#111",
  outline: "none",
  fontFamily: "inherit",
  boxSizing: "border-box",
};

// ── App ───────────────────────────────────────────────────────────────
export default function App() {
  const [coord, setCoord] = useState("");
  const [name, setName] = useState("");
  const [result, setResult] = useState(null);
  const [err, setErr] = useState("");
  const [asc, setAsc] = useState(true);
  const [tab, setTab] = useState("traffic"); // "traffic" | "pm25"
  const [sessions, setSessions] = useState([]);

  const doSearch = useCallback(() => {
    const p = parseCoord(coord);
    if (!p) {
      setErr("Invalid coordinates — expected format: lat, lng");
      return;
    }
    setErr("");
    const dists = ROAD_NAMES.map((n) => ({
      name: n,
      dist: distToRoad(p.lat, p.lng, ROADS[n]),
    }));
    const minDist = Math.min(...dists.map((d) => d.dist));
    const closest = dists.reduce((a, b) => (a.dist < b.dist ? a : b));
    const pm25 = lookupPM25(p.lat, p.lng);
    const score = calcExposureScore(minDist, pm25);
    const r = {
      geo: p,
      dists,
      minDist,
      closest,
      pm25,
      score,
      patientName: name.trim(),
      timestamp: new Date().toLocaleString("en-US"),
    };
    setResult(r);
    setSessions((prev) => [...prev, r]);
  }, [coord, name]);

  const sorted = useMemo(() => {
    if (!result) return [];
    const arr = [...result.dists];
    if (tab === "traffic")
      return arr.sort((a, b) => (asc ? a.dist - b.dist : b.dist - a.dist));
    return arr.sort((a, b) => {
      const pa = ROAD_PM25[a.name] ?? 0,
        pb = ROAD_PM25[b.name] ?? 0;
      return asc ? pa - pb : pb - pa;
    });
  }, [result, asc, tab]);

  const maxDist = result ? Math.max(...result.dists.map((d) => d.dist)) : 1;
  const maxPM25 = Math.max(...ROAD_NAMES.map((n) => ROAD_PM25[n] ?? 0));
  const ok = parseCoord(coord);
  const roadRk = result ? getRoadRisk(result.minDist) : null;
  const pm25Rk = result ? getPM25Risk(result.pm25) : null;
  const scoreRk = result ? getScoreRisk(result.score) : null;

  return (
    <div
      style={{
        padding: "1.25rem 0.75rem",
        fontFamily: "'DM Sans','Segoe UI',sans-serif",
        maxWidth: 760,
        margin: "0 auto",
        fontSize: 13,
        color: "#111",
      }}
    >
      {/* ── Header ── */}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 11,
          marginBottom: 14,
        }}
      >
        <div
          style={{
            width: 40,
            height: 40,
            borderRadius: 11,
            background: "#FEF2F2",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            flexShrink: 0,
          }}
        >
          <svg width="20" height="20" viewBox="0 0 16 16" fill="none">
            <path
              d="M8 1.5L2 13.5h12L8 1.5z"
              stroke="#9B1C1C"
              strokeWidth="1.4"
              fill="none"
              strokeLinejoin="round"
            />
            <line
              x1="8"
              y1="6"
              x2="8"
              y2="9.5"
              stroke="#9B1C1C"
              strokeWidth="1.4"
              strokeLinecap="round"
            />
            <circle cx="8" cy="11.5" r="0.8" fill="#9B1C1C" />
          </svg>
        </div>
        <div>
          <h2 style={{ fontSize: 16, fontWeight: 700, margin: "0 0 2px" }}>
            Traffic & PM2.5 Exposure — COPD HCMC
          </h2>
          <p style={{ fontSize: 11, color: "#9CA3AF", margin: 0 }}>
            {ROAD_NAMES.length} road corridors · PM2.5 NASA ACAG 2023 (0.05°) ·
            OSM GeoJSON
          </p>
        </div>
      </div>

      {/* ── Input ── */}
      <div style={{ ...card, marginBottom: 12 }}>
        <label
          style={{
            fontSize: 11,
            fontWeight: 700,
            color: "#3B82F6",
            textTransform: "uppercase",
            letterSpacing: "0.06em",
            display: "block",
            marginBottom: 8,
          }}
        >
          📍 Google Maps Coordinates
          <span
            style={{
              fontWeight: 400,
              color: "#9CA3AF",
              fontSize: 10,
              marginLeft: 8,
            }}
          >
            right-click → click the coordinate line → paste
          </span>
        </label>
        <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
          <div style={{ flex: 2 }}>
            <input
              style={{
                ...inp,
                borderColor: coord ? "#3B82F6" : "#E5E7EB",
                background: coord ? "#EFF6FF" : "#FAFAFA",
                fontFamily: "monospace",
                fontSize: 14,
              }}
              placeholder="10.758773, 106.649111"
              value={coord}
              onChange={(e) => setCoord(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && doSearch()}
              autoFocus
            />
            {!coord ? (
              <div style={{ fontSize: 10, color: "#9CA3AF", marginTop: 5 }}>
                Paste coordinates from Google Maps
              </div>
            ) : !ok ? (
              <div style={{ fontSize: 10, color: "#EF4444", marginTop: 5 }}>
                ⚠ Not recognized — expected "lat, lng"
              </div>
            ) : (
              <div style={{ fontSize: 10, color: "#166534", marginTop: 5 }}>
                ✓ {ok.lat.toFixed(6)}, {ok.lng.toFixed(6)}
              </div>
            )}
          </div>
          <div style={{ flex: 1, minWidth: 0 }}>
            <input
              style={{
                ...inp,
                borderColor: name ? "#8B5CF6" : "#E5E7EB",
                background: name ? "#F5F3FF" : "#FAFAFA",
              }}
              placeholder="Patient name / Sample ID"
              value={name}
              onChange={(e) => setName(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && doSearch()}
            />
            <div style={{ fontSize: 10, color: "#9CA3AF", marginTop: 5 }}>
              {name ? (
                <span style={{ color: "#7C3AED" }}>✓ {name}</span>
              ) : (
                "Sample label"
              )}
            </div>
          </div>
          <button
            onClick={doSearch}
            disabled={!ok}
            style={{
              height: 42,
              padding: "0 22px",
              fontSize: 13,
              fontWeight: 700,
              border: "none",
              borderRadius: 9,
              flexShrink: 0,
              fontFamily: "inherit",
              whiteSpace: "nowrap",
              background: ok ? "#111" : "#D1D5DB",
              color: "#fff",
              cursor: ok ? "pointer" : "not-allowed",
            }}
          >
            Measure
          </button>
        </div>
        {err && (
          <div
            style={{
              marginTop: 8,
              fontSize: 12,
              color: "#9B1C1C",
              background: "#FEF2F2",
              borderRadius: 7,
              padding: "7px 12px",
            }}
          >
            {err}
          </div>
        )}
      </div>

      {/* ── Session bar ── */}
      {sessions.length > 0 && (
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            marginBottom: 10,
            padding: "8px 14px",
            background: "#F8FAFC",
            borderRadius: 9,
            border: "1px solid #E5E7EB",
          }}
        >
          <span style={{ fontSize: 12, color: "#6B7280" }}>
            💾 <strong style={{ color: "#111" }}>{sessions.length}</strong>{" "}
            measurement{sessions.length !== 1 ? "s" : ""}
            {sessions.some((s) => s.patientName) && (
              <span style={{ color: "#7C3AED", marginLeft: 8 }}>
                ·{" "}
                {[
                  ...new Set(
                    sessions
                      .filter((s) => s.patientName)
                      .map((s) => s.patientName)
                  ),
                ]
                  .slice(0, 3)
                  .join(", ")}
                {sessions.filter((s) => s.patientName).length > 3 ? "…" : ""}
              </span>
            )}
          </span>
          <button
            onClick={() => exportExcel(sessions)}
            style={{
              padding: "6px 16px",
              fontSize: 12,
              fontWeight: 700,
              border: "none",
              borderRadius: 7,
              background: "#166534",
              color: "#fff",
              cursor: "pointer",
              fontFamily: "inherit",
            }}
          >
            ⬇ Export Excel ({sessions.length} case
            {sessions.length !== 1 ? "s" : ""})
          </button>
        </div>
      )}

      {/* ── Result ── */}
      {result && (
        <>
          {/* 4 metric cards */}
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(4,1fr)",
              gap: 8,
              marginBottom: 12,
            }}
          >
            {/* Coordinates */}
            <div
              style={{
                ...card,
                display: "flex",
                flexDirection: "column",
                gap: 4,
              }}
            >
              <div
                style={{
                  fontSize: 10,
                  fontWeight: 700,
                  color: "#9CA3AF",
                  textTransform: "uppercase",
                }}
              >
                Coordinates
              </div>
              <div
                style={{
                  fontSize: 11,
                  fontWeight: 700,
                  lineHeight: 1.6,
                  fontFamily: "monospace",
                }}
              >
                {result.geo.lat.toFixed(5)}
                <br />
                {result.geo.lng.toFixed(5)}
              </div>
              {result.patientName && (
                <div
                  style={{ fontSize: 10, color: "#7C3AED", fontWeight: 600 }}
                >
                  👤 {result.patientName}
                </div>
              )}
            </div>

            {/* Nearest road */}
            <div
              style={{
                ...card,
                display: "flex",
                flexDirection: "column",
                gap: 3,
              }}
            >
              <div
                style={{
                  fontSize: 10,
                  fontWeight: 700,
                  color: "#9CA3AF",
                  textTransform: "uppercase",
                }}
              >
                Nearest Road
              </div>
              <div
                style={{
                  fontSize: 26,
                  fontWeight: 800,
                  lineHeight: 1,
                  fontVariantNumeric: "tabular-nums",
                }}
              >
                {result.minDist.toLocaleString("en-US")}
                <span style={{ fontSize: 12, fontWeight: 400, marginLeft: 3 }}>
                  m
                </span>
              </div>
              <div
                style={{
                  fontSize: 10,
                  color: "#6B7280",
                  overflow: "hidden",
                  textOverflow: "ellipsis",
                  whiteSpace: "nowrap",
                }}
              >
                {result.closest.name}
              </div>
              <span
                style={{
                  display: "inline-block",
                  padding: "2px 8px",
                  borderRadius: 20,
                  alignSelf: "flex-start",
                  background: roadRk.bg,
                  color: roadRk.color,
                  fontSize: 10,
                  fontWeight: 700,
                }}
              >
                {roadRk.label}
              </span>
            </div>

            {/* Home PM2.5 */}
            <div
              style={{
                ...card,
                background: pm25Rk.bg,
                border: `1px solid ${pm25Rk.bar}33`,
                display: "flex",
                flexDirection: "column",
                gap: 3,
              }}
            >
              <div
                style={{
                  fontSize: 10,
                  fontWeight: 700,
                  color: pm25Rk.color,
                  textTransform: "uppercase",
                  opacity: 0.8,
                }}
              >
                Home PM2.5
              </div>
              <div
                style={{
                  fontSize: 26,
                  fontWeight: 800,
                  color: pm25Rk.color,
                  lineHeight: 1,
                  fontVariantNumeric: "tabular-nums",
                }}
              >
                {result.pm25 ?? "–"}
                <span style={{ fontSize: 12, fontWeight: 400, marginLeft: 3 }}>
                  µg/m³
                </span>
              </div>
              <div style={{ fontSize: 10, color: pm25Rk.color, opacity: 0.7 }}>
                NASA ACAG 2023
              </div>
              <span
                style={{
                  display: "inline-block",
                  padding: "2px 8px",
                  borderRadius: 20,
                  alignSelf: "flex-start",
                  background: "rgba(255,255,255,0.6)",
                  color: pm25Rk.color,
                  border: `1px solid ${pm25Rk.bar}`,
                  fontSize: 10,
                  fontWeight: 700,
                }}
              >
                {pm25Rk.label}
              </span>
            </div>

            {/* Exposure Score */}
            <div
              style={{
                ...card,
                background: scoreRk.bg,
                border: `1px solid ${scoreRk.bar}33`,
                display: "flex",
                flexDirection: "column",
                gap: 3,
              }}
            >
              <div
                style={{
                  fontSize: 10,
                  fontWeight: 700,
                  color: scoreRk.color,
                  textTransform: "uppercase",
                  opacity: 0.8,
                }}
              >
                Exposure Score
              </div>
              <div
                style={{
                  fontSize: 26,
                  fontWeight: 800,
                  color: scoreRk.color,
                  lineHeight: 1,
                }}
              >
                {result.score}
                <span style={{ fontSize: 12, fontWeight: 400 }}>/100</span>
              </div>
              <div style={{ fontSize: 10, color: scoreRk.color, opacity: 0.7 }}>
                Traffic 60% · PM2.5 40%
              </div>
              <span
                style={{
                  display: "inline-block",
                  padding: "2px 8px",
                  borderRadius: 20,
                  alignSelf: "flex-start",
                  background: "rgba(255,255,255,0.6)",
                  color: scoreRk.color,
                  border: `1px solid ${scoreRk.bar}`,
                  fontSize: 10,
                  fontWeight: 700,
                }}
              >
                {scoreRk.label}
              </span>
            </div>
          </div>

          {/* Tab table */}
          <div style={{ ...card, padding: 0, overflow: "hidden" }}>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                padding: "10px 16px",
                borderBottom: "1px solid #F3F4F6",
                gap: 8,
                flexWrap: "wrap",
              }}
            >
              <div style={{ display: "flex", gap: 6 }}>
                {[
                  ["traffic", "🚗 Distance"],
                  ["pm25", "🌫 Road PM2.5"],
                ].map(([k, lbl]) => (
                  <button
                    key={k}
                    onClick={() => setTab(k)}
                    style={{
                      padding: "5px 12px",
                      fontSize: 11,
                      fontWeight: 700,
                      border: "none",
                      borderRadius: 7,
                      cursor: "pointer",
                      fontFamily: "inherit",
                      background: tab === k ? "#111" : "#F3F4F6",
                      color: tab === k ? "#fff" : "#6B7280",
                    }}
                  >
                    {lbl}
                  </button>
                ))}
              </div>
              <button
                onClick={() => setAsc((s) => !s)}
                style={{
                  fontSize: 11,
                  color: "#6B7280",
                  border: "1px solid #E5E7EB",
                  background: "#FAFAFA",
                  cursor: "pointer",
                  padding: "4px 10px",
                  borderRadius: 6,
                  fontFamily: "inherit",
                }}
              >
                {asc ? "↑ Ascending" : "↓ Descending"}
              </button>
            </div>

            <table
              style={{
                width: "100%",
                borderCollapse: "collapse",
                fontSize: 13,
              }}
            >
              <thead>
                <tr>
                  {(tab === "traffic"
                    ? [
                        "#",
                        "Road Name",
                        "Distance to Edge",
                        "Road PM2.5",
                        "Risk Level",
                      ]
                    : [
                        "#",
                        "Road Name",
                        "PM2.5 centroid (µg/m³)",
                        "Distance",
                        "WHO Category",
                      ]
                  ).map((h) => (
                    <th
                      key={h}
                      style={{
                        padding: "8px 14px",
                        textAlign: "left",
                        fontSize: 11,
                        fontWeight: 700,
                        color: "#9CA3AF",
                        background: "#FAFAFA",
                        borderBottom: "1px solid #F3F4F6",
                        whiteSpace: "nowrap",
                      }}
                    >
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sorted.map((r, i) => {
                  const isC = r.name === result.closest.name;
                  const rPM25 = ROAD_PM25[r.name];
                  const rowBg = isC ? "#EFF6FF" : "transparent";
                  const border = "1px solid #F9FAFB";
                  const tdBase = { padding: "7px 14px", borderBottom: border };

                  if (tab === "traffic") {
                    const rk = getRoadRisk(r.dist);
                    const prk = getPM25Risk(rPM25);
                    const pct = Math.min(
                      100,
                      Math.round((r.dist / maxDist) * 100)
                    );
                    return (
                      <tr key={r.name} style={{ background: rowBg }}>
                        <td
                          style={{
                            ...tdBase,
                            color: "#D1D5DB",
                            fontSize: 11,
                            width: 28,
                          }}
                        >
                          {i + 1}
                        </td>
                        <td
                          style={{
                            ...tdBase,
                            fontWeight: isC ? 700 : 500,
                            whiteSpace: "nowrap",
                          }}
                        >
                          {isC && (
                            <span style={{ color: "#2563EB", marginRight: 5 }}>
                              ★
                            </span>
                          )}
                          {r.name}
                        </td>
                        <td style={tdBase}>
                          <div
                            style={{
                              display: "flex",
                              alignItems: "center",
                              gap: 8,
                            }}
                          >
                            <div
                              style={{
                                width: 60,
                                height: 4,
                                background: "#F3F4F6",
                                borderRadius: 2,
                                flexShrink: 0,
                              }}
                            >
                              <div
                                style={{
                                  width: `${pct}%`,
                                  height: 4,
                                  borderRadius: 2,
                                  background: rk.bar,
                                }}
                              />
                            </div>
                            <span
                              style={{
                                fontWeight: 700,
                                fontVariantNumeric: "tabular-nums",
                              }}
                            >
                              {r.dist.toLocaleString("en-US")} m
                            </span>
                          </div>
                        </td>
                        <td style={tdBase}>
                          <span
                            style={{
                              fontWeight: 600,
                              color: prk.color,
                              fontSize: 12,
                            }}
                          >
                            {rPM25 ?? "–"} µg/m³
                          </span>
                        </td>
                        <td style={tdBase}>
                          <span
                            style={{
                              display: "inline-block",
                              padding: "2px 8px",
                              borderRadius: 20,
                              background: rk.bg,
                              color: rk.color,
                              fontSize: 10,
                              fontWeight: 700,
                            }}
                          >
                            {rk.label}
                          </span>
                        </td>
                      </tr>
                    );
                  } else {
                    const prk = getPM25Risk(rPM25);
                    const pct = Math.min(
                      100,
                      Math.round(((rPM25 ?? 0) / maxPM25) * 100)
                    );
                    return (
                      <tr key={r.name} style={{ background: rowBg }}>
                        <td
                          style={{
                            ...tdBase,
                            color: "#D1D5DB",
                            fontSize: 11,
                            width: 28,
                          }}
                        >
                          {i + 1}
                        </td>
                        <td
                          style={{
                            ...tdBase,
                            fontWeight: isC ? 700 : 500,
                            whiteSpace: "nowrap",
                          }}
                        >
                          {isC && (
                            <span style={{ color: "#2563EB", marginRight: 5 }}>
                              ★
                            </span>
                          )}
                          {r.name}
                        </td>
                        <td style={tdBase}>
                          <div
                            style={{
                              display: "flex",
                              alignItems: "center",
                              gap: 8,
                            }}
                          >
                            <div
                              style={{
                                width: 60,
                                height: 4,
                                background: "#F3F4F6",
                                borderRadius: 2,
                                flexShrink: 0,
                              }}
                            >
                              <div
                                style={{
                                  width: `${pct}%`,
                                  height: 4,
                                  borderRadius: 2,
                                  background: prk.bar,
                                }}
                              />
                            </div>
                            <span
                              style={{
                                fontWeight: 700,
                                fontVariantNumeric: "tabular-nums",
                                color: prk.color,
                              }}
                            >
                              {rPM25 ?? "N/A"} µg/m³
                            </span>
                          </div>
                        </td>
                        <td
                          style={{ ...tdBase, color: "#9CA3AF", fontSize: 12 }}
                        >
                          {r.dist.toLocaleString("en-US")} m
                        </td>
                        <td style={tdBase}>
                          <span
                            style={{
                              display: "inline-block",
                              padding: "2px 8px",
                              borderRadius: 20,
                              background: prk.bg,
                              color: prk.color,
                              fontSize: 10,
                              fontWeight: 700,
                            }}
                          >
                            {prk.short}
                          </span>
                        </td>
                      </tr>
                    );
                  }
                })}
              </tbody>
            </table>
          </div>

          <p
            style={{
              fontSize: 10,
              color: "#D1D5DB",
              marginTop: 8,
              lineHeight: 1.8,
            }}
          >
            Distance: Haversine point-to-segment to road edge · PM2.5: NASA ACAG
            V6GL03 (2023) nearest-neighbor 0.05°×0.05° · Exposure Score =
            Traffic 60% + PM2.5 40% · WHO AQG 2021 target ≤5 µg/m³
          </p>
        </>
      )}
    </div>
  );
}
