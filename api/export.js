import ExcelJS from "exceljs";

function extractMatchId(url) {
  const match = url.match(/\/room\/([^/]+)\/scoreboard/i);
  if (match) return match[1];

  if (url.startsWith("1-")) return url;

  throw new Error("Invalid FACEIT link");
}

async function getMatchStats(matchId, apiKey) {
  const response = await fetch(
    `https://open.faceit.com/data/v4/matches/${matchId}/stats`,
    {
      headers: {
        Authorization: `Bearer ${apiKey}`,
      },
    }
  );

  if (!response.ok) {
    throw new Error("FACEIT API error");
  }

  return await response.json();
}

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).send("Method not allowed");
    }

    const { url } = req.body;
    const apiKey = process.env.FACEIT_API_KEY;

    if (!apiKey) {
      return res.status(500).send("Missing FACEIT_API_KEY");
    }

    const matchId = extractMatchId(url);
    const data = await getMatchStats(matchId, apiKey);

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Player Stats");

    sheet.columns = [
      { header: "Nickname", key: "nickname", width: 20 },
      { header: "Kills", key: "kills", width: 10 },
      { header: "Deaths", key: "deaths", width: 10 },
      { header: "Assists", key: "assists", width: 10 },
    ];

    data.rounds[0].teams.forEach(team => {
      team.players.forEach(player => {
        sheet.addRow({
          nickname: player.nickname,
          kills: player.player_stats.Kills,
          deaths: player.player_stats.Deaths,
          assists: player.player_stats.Assists,
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
      `attachment; filename="faceit_stats.xlsx"`
    );

    await workbook.xlsx.write(res);
    res.end();

  } catch (error) {
    res.status(400).send(error.message);
  }
}
