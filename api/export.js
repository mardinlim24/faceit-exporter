// /pages/api/export.js  (Next.js - Vercel)
// npm i exceljs
import ExcelJS from "exceljs";

export default async function handler(req, res) {
  try {
    if (req.method !== "GET") {
      res.setHeader("Allow", "GET");
      return res.status(405).send("Method Not Allowed");
    }

    const { url, format } = req.query;

    if (!url || typeof url !== "string") {
      return res.status(400).json({ error: "Missing ?url=" });
    }

    // لو بدك JSON للتجربة: /api/export?url=...&format=json
    const wantJson = String(format || "").toLowerCase() === "json";

    const origin =
      req.headers["x-forwarded-proto"] && req.headers.host
        ? `${req.headers["x-forwarded-proto"]}://${req.headers.host}`
        : `https://${req.headers.host}`;

    // هذا لازم يكون موجود: /api/stats?url=... (يرجع JSON)
    const statsUrl = `${origin}/api/stats?url=${encodeURIComponent(url)}`;

    const r = await fetch(statsUrl, {
      headers: { "User-Agent": "faceit-exporter-xlsx" },
    });

    if (!r.ok) {
      const txt = await r.text().catch(() => "");
      return res
        .status(502)
        .json({ error: "Failed to fetch stats JSON", status: r.status, body: txt });
    }

    const data = await r.json();

    if (wantJson) {
      return res.status(200).json(data);
    }

    // ====== توقع شكل البيانات ======
    // data.summary = { match_id, team1, team2, score1, score2, winner, best_of, region, competition, status }
    // data.players = [ { team, nickname, rank, rws, kills, deaths, assists, adr, kd, kr, hs, hs_pct, k5, k4, k3, k2, mvps } ]

    const summary = data?.summary || {};
    const players = Array.isArray(data?.players) ? data.players : [];

    // ====== إنشاء ملف Excel ======
    const wb = new ExcelJS.Workbook();
    wb.creator = "faceit-exporter";
    wb.created = new Date();

    // Sheet 1: MatchSummary
    const ws1 = wb.addWorksheet("MatchSummary", {
      views: [{ state: "frozen", ySplit: 1 }],
    });

    ws1.columns = [
      { header: "field", key: "field", width: 22 },
      { header: "value", key: "value", width: 60 },
    ];

    const summaryRows = [
      ["match_id", summary.match_id],
      ["team1", summary.team1],
      ["team2", summary.team2],
      ["score1", summary.score1],
      ["score2", summary.score2],
      ["winner", summary.winner],
      ["best_of", summary.best_of],
      ["region", summary.region],
      ["competition", summary.competition],
      ["status", summary.status],
      ["source_url", url],
    ];

    for (const [field, value] of summaryRows) {
      ws1.addRow({ field, value: value ?? "" });
    }

    ws1.getRow(1).font = { bold: true };
    ws1.getRow(1).alignment = { vertical: "middle" };

    // Sheet 2: PlayerStats
    const ws2 = wb.addWorksheet("PlayerStats", {
      views: [{ state: "frozen", ySplit: 1 }],
    });

    const cols = [
      "team",
      "nickname",
      "rank",
      "rws",
      "kills",
      "deaths",
      "assists",
      "adr",
      "kd",
      "kr",
      "hs",
      "hs_pct",
      "k5",
      "k4",
      "k3",
      "k2",
      "mvps",
    ];

    ws2.columns = cols.map((c) => ({ header: c, key: c, width: c.length < 6 ? 10 : 14 }));
    ws2.getRow(1).font = { bold: true };

    // املأ الداتا (مع معالجة null)
    for (const p of players) {
      const row = {};
      for (const c of cols) row[c] = p?.[c] ?? "";
      ws2.addRow(row);
    }

    // تنسيق بسيط للأرقام
    const numericCols = new Set([
      "rws",
      "kills",
      "deaths",
      "assists",
      "adr",
      "kd",
      "kr",
      "hs",
      "hs_pct",
      "k5",
      "k4",
      "k3",
      "k2",
      "mvps",
    ]);

    ws2.columns.forEach((col) => {
      if (numericCols.has(col.key)) {
        col.numFmt = "0.00";
      }
      if (["kills", "deaths", "assists", "hs", "k5", "k4", "k3", "k2", "mvps"].includes(col.key)) {
        col.numFmt = "0";
      }
    });

    // ====== رجّع ملف XLSX صحيح ======
    const filename = `faceit_match_${summary.match_id || "export"}.xlsx`;

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.setHeader("Cache-Control", "no-store");

    await wb.xlsx.write(res);
    res.end();
  } catch (e) {
    return res.status(500).json({
      error: "Export failed",
      message: e?.message || String(e),
    });
  }
}
