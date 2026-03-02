function extractMatchId(input) {
  const s = String(input || "").trim();
  const m = s.match(/\/room\/([^/]+)\/scoreboard/i);
  if (m) return m[1];
  if (s.startsWith("1-") && s.length > 10) return s;
  throw new Error("Invalid URL. Paste FACEIT scoreboard link.");
}

function num(v) {
  if (v === null || v === undefined || v === "") return "";
  const n = Number(v);
  return Number.isFinite(n) ? n : v;
}

function pick(obj, keys, fallback = "") {
  for (const k of keys) {
    const v = obj?.[k];
    if (v !== undefined && v !== null && v !== "") return v;
  }
  return fallback;
}

async function faceitJson(url, apiKey) {
  const r = await fetch(url, { headers: { Authorization: `Bearer ${apiKey}` } });
  const text = await r.text();
  if (!r.ok) throw new Error(`FACEIT API ${r.status}: ${text}`);
  return JSON.parse(text);
}

export default async function handler(req, res) {
  try {
    // Allow GET + POST
    if (req.method !== "GET" && req.method !== "POST") {
      res.setHeader("Allow", ["GET", "POST"]);
      return res.status(405).send("Method not allowed");
    }

    const apiKey = process.env.FACEIT_API_KEY;
    if (!apiKey) return res.status(500).json({ error: "Missing FACEIT_API_KEY" });

    // url can come from query (GET) or body (POST)
    const inputUrl = req.method === "GET" ? req.query.url : req.body?.url;
    const matchId = extractMatchId(inputUrl);

    const match = await faceitJson(`https://open.faceit.com/data/v4/matches/${matchId}`, apiKey);
    const stats = await faceitJson(`https://open.faceit.com/data/v4/matches/${matchId}/stats`, apiKey);

    const round0 = stats?.rounds?.[0];
    if (!round0?.teams?.length) {
      return res.status(400).json({ error: "No stats rounds found (match may be ongoing/private)" });
    }

    const players = [];
    for (const team of round0.teams) {
      const teamName =
        pick(team?.team_stats, ["Team", "team", "Name", "name"], "") ||
        team?.team_id ||
        "";

      for (const p of team.players || []) {
        const ps = p.player_stats || {};
        players.push({
          team: teamName,
          nickname: p.nickname || "",
          rank: p.game_player_id || p.player_id || "",
          rws: num(pick(ps, ["RWS", "rws"], "")),
          kills: num(pick(ps, ["Kills", "kills"], "")),
          deaths: num(pick(ps, ["Deaths", "deaths"], "")),
          assists: num(pick(ps, ["Assists", "assists"], "")),
          adr: num(pick(ps, ["ADR", "Average Damage per Round", "avg_damage_per_round"], "")),
          kd: num(pick(ps, ["K/D Ratio", "K/D", "kd_ratio"], "")),
          kr: num(pick(ps, ["K/R Ratio", "K/R", "kr_ratio"], "")),
          hs: num(pick(ps, ["Headshots", "HS", "headshots"], "")),
          hs_pct: num(pick(ps, ["Headshots %", "HS %", "hs_percentage"], "")),
          k5: num(pick(ps, ["Penta Kills", "5K", "penta_kills"], 0)),
          k4: num(pick(ps, ["Quadro Kills", "4K", "quadro_kills"], 0)),
          k3: num(pick(ps, ["Triple Kills", "3K", "triple_kills"], 0)),
          k2: num(pick(ps, ["Double Kills", "2K", "double_kills"], 0)),
          mvps: num(pick(ps, ["MVPs", "MVP", "mvps"], 0)),
        });
      }
    }

    // IMPORTANT: Excel Power Query likes plain JSON
    return res.status(200).json({
      match_id: matchId,
      summary: {
        team1: match?.teams?.faction1?.name || "",
        team2: match?.teams?.faction2?.name || "",
        score1: match?.results?.score?.faction1 ?? "",
        score2: match?.results?.score?.faction2 ?? "",
        winner: match?.results?.winner ?? "",
        best_of: match?.best_of ?? "",
        region: match?.region ?? "",
        competition: match?.competition?.name ?? "",
        status: match?.status ?? "",
      },
      players,
    });
  } catch (e) {
    console.error(e);
    return res.status(400).json({ error: e?.message || "Error" });
  }
}
