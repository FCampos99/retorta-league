export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).end();

  const { password, season_id, date, winners, losers } = req.body;

  // Verificar palavra-passe
  if (!password || password !== process.env.REPORT_PASSWORD) {
    return res.status(401).json({ error: 'Palavra-passe incorrecta' });
  }

  const token = process.env.REPORT_TOKEN;
  if (!token) return res.status(500).json({ error: 'Token não configurado no servidor' });

  try {
    const ghRes = await fetch(
      'https://api.github.com/repos/FCampos99/retorta-league/actions/workflows/reportar.yml/dispatches',
      {
        method: 'POST',
        headers: {
          Authorization: `token ${token}`,
          Accept: 'application/vnd.github.v3+json',
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          ref: 'main',
          inputs: {
            season_id,
            date,
            winners: JSON.stringify(winners),
            losers:  JSON.stringify(losers),
          },
        }),
      }
    );

    if (ghRes.status === 204 || ghRes.ok) {
      res.status(200).json({ ok: true });
    } else {
      const err = await ghRes.json();
      res.status(500).json({ error: err.message || 'Erro desconhecido' });
    }
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}
