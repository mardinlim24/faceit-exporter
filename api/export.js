import ExcelJS from "exceljs";

function extractMatchId(input) {
  const url = String(input || "").trim();
  const m = url.match(/\/room\/([^/]+)\/scoreboard/i);
  if (m) return m[1];
  if (url.startsWith("1-") && url.length > 10) return url;
  throw new Error("Invalid FACEIT link. Paste the scoreboard URL or match id.");
}

function pick(obj, keys, fallback = "") {
  for (const k of keys) {
    if (obj && obj[k] !== undefined && obj[k] !== null && obj[k] !== "") return obj[k];
  }
  return fallback;
}

async function faceitFetchJson(url, apiKey) {
  const r = await fetch(url, {
    headers: { Authorization: `Bearer ${apiKey}` },
  });
  const text = await r.text();
  if (!r.ok) {
    throw new Error(`FACEIT API ${r.status}: ${text}`);
  }
  return JSON.parse(text);
}

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") return res.status(405).send("Method not allowed");

    const { url } = req.body || {};
    const matchId = extractMatchId(url);

    // Support BOTH env var names just in case you used a different one in Vercel
    const apiKey = process.env.FACEIT_API_KEY || process.env.FACEIT_API_KEY;
    if (!apiKey) return res.status(500).send("Missing API key. Set FACEIT_API_KEY in Vercel Environment Variables.");

    // Fetch match details + stats
    const match = await faceitFetchJson(`https://open.faceit.com/data/v4/matches/${matchId}`, apiKey);
    const stats = await faceitFetchJson(`https://open.faceit.com/data/v4/matches/${matchId}/stats`, apiKey);

    const wb = new ExcelJS.Workbook();

    // ======================
    // Sheet: Match Summary
    // ======================
    const summary = wb.addWorksheet("Match Summary");
    const team1 = match?.teams?.faction1?.name || "faction1";
    const team2 = match?.teams?.faction2?.name || "faction2";

    const score1 = match?.results?.score?.faction1 ?? "";
    const score2 = match?.results?.score?.faction2 ?? "";
    const winner = match?.results?.winner ?? "";
    const region = match?.region ?? "";
    const competition = match?.competition?.name ?? "";
    const bestOf = match?.best_of ?? "";
    const status = match?.status ?? "";

    summary.getColumn(1).width = 22;
    summary.getColumn(2).width = 46;

    summary.addRow(["Match ID", matchId]);
    summary.addRow(["Team 1", team1]);
    summary.addRow(["Team 2", team2]);
    summary.addRow(["Score", `${score1} - ${score2}`]);
    summary.addRow(["Winner", winner]);
    summary.addRow(["Best Of", bestOf]);
    summary.addRow(["Status", status]);
    summary.addRow(["Region", region]);
    summary.addRow(["Competition", competition]);

    summary.getRow(1).font = { bold: true };

    // ======================
    // Sheet: Player Stats
    // ======================
    const sh = wb.addWorksheet("Player Stats");
    sh.columns = [
      { header: "Team", key: "team", width: 20 },
      { header: "Nickname", key: "nickname", width: 20 },
      { header: "Rank", key: "rank", width: 10 },
      { header: "RWS", key: "rws", width: 10 },
      { header: "K", key: "kills", width: 8 },
      { header: "D", key: "deaths", width: 8 },
      { header: "A", key: "assists", width: 8 },
      { header: "ADR", key: "adr", width: 10 },
      { header: "K/D", key: "kd", width: 10 },
      { header: "K/R", key: "kr", width: 10 },
      { header: "HS", key: "hs", width: 8 },
      { header: "HS %", key: "hs_pct", width: 10 },
      { header: "5k", key: "k5", width: 8 },
      { header: "4k", key: "k4", width: 8 },
      { header: "3k", key: "k3", width: 8 },
      { header: "2k", key: "k2", width: 8 },
      { header: "MVPs", key: "mvps", width: 10 },
    ];

    // Header styling
    sh.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    sh.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF111111" } };
    sh.views = [{ state: "frozen", ySplit: 1 }];
    sh.autoFilter = { from: "A1", to: "Q1" };

    // Parse stats safely: rounds[0] usually exists
    const round0 = (stats?.rounds && stats.rounds[0]) ? stats.rounds[0] : null;
    if (!round0 || !round0.teams) {
      return res.status(400).send("No rounds/teams found in FACEIT stats for this match.");
    }

    for (const team of round0.teams) {
      const teamName =
        pick(team?.team_stats, ["Team", "team", "Name", "name"], "") ||
        team?.team_id ||
        "";

      for (const player of (team.players || [])) {
        const ps = player.player_stats || {};

        sh.addRow({
          team: teamName,
          nickname: player.nickname || "",
          rank: pick(player, ["game_player_id", "player_id"], ""),
          rws: pick(ps, ["RWS", "rws"], ""),
          kills: pick(ps, ["Kills", "kills"], ""),
          deaths: pick(ps, ["Deaths", "deaths"], ""),
          assists: pick(ps, ["Assists", "assists"], ""),
          adr: pick(ps, ["ADR", "Average Damage per Round", "avg_damage_per_round"], ""),
          kd: pick(ps, ["K/D Ratio", "K/D", "kd_ratio"], ""),
          kr: pick(ps, ["K/R Ratio", "K/R", "kr_ratio"], ""),
          hs: pick(ps, ["Headshots", "HS", "headshots"], ""),
          hs_pct: pick(ps, ["Headshots %", "HS %", "hs_percentage"], ""),
          k5: pick(ps, ["Penta Kills", "5K", "penta_kills"], 0),
          k4: pick(ps, ["Quadro Kills", "4K", "quadro_kills"], 0),
          k3: pick(ps, ["Triple Kills", "3K", "triple_kills"], 0),
          k2: pick(ps, ["Double Kills", "2K", "double_kills"], 0),
          mvps: pick(ps, ["MVPs", "MVP", "mvps"], 0),
        });
      }
    }

    // Optional: sort by RWS desc if numeric
    // (Excel will still open fine even if strings)
    // You can ignore this if you don't care about sorting.

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="faceit_full_stats.xlsx"'
    );

    await wb.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error(error); // so you see details in Vercel Logs
    res.status(400).send(error?.message || "Error");
  }
}
