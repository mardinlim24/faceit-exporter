import ExcelJS from "exceljs";

function extractMatchId(url) {
  const match = url.match(/\/room\/([^/]+)\/scoreboard/i);
  if (match) return match[1];
  if (url.startsWith("1-")) return url;
  throw new Error("Invalid FACEIT link");
}

async function getMatchData(matchId, apiKey) {
  const matchRes = await fetch(
    `https://open.faceit.com/data/v4/matches/${matchId}`,
    { headers: { Authorization: `Bearer ${apiKey}` } }
  );

  const statsRes = await fetch(
    `https://open.faceit.com/data/v4/matches/${matchId}/stats`,
    { headers: { Authorization: `Bearer ${apiKey}` } }
  );

  if (!matchRes.ok || !statsRes.ok)
    throw new Error("FACEIT API error");

  return {
    match: await matchRes.json(),
    stats: await statsRes.json()
  };
}

export default async function handler(req, res) {
  try {
    if (req.method !== "POST")
      return res.status(405).send("Method not allowed");

    const { url } = req.body;
    const apiKey = process.env.FACEIT_API_KEY;

    const matchId = extractMatchId(url);
    const { match, stats } = await getMatchData(matchId, apiKey);

    const workbook = new ExcelJS.Workbook();

    // =====================
    // Match Summary Sheet
    // =====================

    const summary = workbook.addWorksheet("Match Summary");

    summary.addRow(["Team 1", match.teams.faction1.name]);
    summary.addRow(["Team 2", match.teams.faction2.name]);
    summary.addRow(["Score", match.results.score.faction1 + " - " + match.results.score.faction2]);
    summary.addRow(["Winner", match.results.winner]);
    summary.addRow(["Region", match.region]);
    summary.addRow(["Competition", match.competition.name]);

    summary.getColumn(1).width = 20;
    summary.getColumn(2).width = 30;

    // =====================
    // Player Stats Sheet
    // =====================

    const sheet = workbook.addWorksheet("Player Stats");

    sheet.columns = [
      { header: "Team", key: "team", width: 20 },
      { header: "Nickname", key: "nickname", width: 20 },
      { header: "Rank", key: "rank", width: 10 },
      { header: "RWS", key: "rws", width: 10 },
      { header: "Kills", key: "kills", width: 10 },
      { header: "Deaths", key: "deaths", width: 10 },
      { header: "Assists", key: "assists", width: 10 },
      { header: "ADR", key: "adr", width: 10 },
      { header: "K/D", key: "kd", width: 10 },
      { header: "K/R", key: "kr", width: 10 },
      { header: "HS", key: "hs", width: 10 },
      { header: "HS%", key: "hspercent", width: 10 },
      { header: "5k", key: "k5", width: 10 },
      { header: "4k", key: "k4", width: 10 },
      { header: "3k", key: "k3", width: 10 },
      { header: "2k", key: "k2", width: 10 },
      { header: "MVPs", key: "mvps", width: 10 }
    ];

    stats.rounds[0].teams.forEach(team => {
      team.players.forEach(player => {
        const s = player.player_stats;

        sheet.addRow({
          team: team.team_stats.Team,
          nickname: player.nickname,
          rank: player.game_player_id,
          rws: s.RWS,
          kills: s.Kills,
          deaths: s.Deaths,
          assists: s.Assists,
          adr: s.ADR,
          kd: s["K/D Ratio"],
          kr: s["K/R Ratio"],
          hs: s.Headshots,
          hspercent: s["Headshots %"],
          k5: s["Penta Kills"],
          k4: s["Quadro Kills"],
          k3: s["Triple Kills"],
          k2: s["Double Kills"],
          mvps: s.MVPs
        });
      });
    });

    sheet.getRow(1).font = { bold: true };

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="faceit_full_stats.xlsx"`
    );

    await workbook.xlsx.write(res);
    res.end();

  } catch (error) {
    res.status(400).send(error.message);
  }
}
