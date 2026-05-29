export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).end();

  const { filename } = req.body;
  const token = process.env.CAPAS_TOKEN;
  if (!token) return res.status(500).json({ error: 'Token não configurado no servidor' });

  const api = 'https://api.github.com/repos/FCampos99/retorta-league/contents/capas';
  const headers = {
    Authorization: `token ${token}`,
    Accept: 'application/vnd.github.v3+json',
    'Content-Type': 'application/json',
  };

  try {
    // 1. Apagar imagem
    const imgRes = await fetch(`${api}/${filename}`, { headers });
    if (imgRes.ok) {
      const imgData = await imgRes.json();
      await fetch(`${api}/${filename}`, {
        method: 'DELETE',
        headers,
        body: JSON.stringify({ message: `Remover: ${filename}`, sha: imgData.sha }),
      });
    }

    // 2. Actualizar index.json
    const idxRes = await fetch(`${api}/index.json`, { headers });
    if (idxRes.ok) {
      const d = await idxRes.json();
      const index = JSON.parse(Buffer.from(d.content, 'base64').toString()).filter(c => c.filename !== filename);
      await fetch(`${api}/index.json`, {
        method: 'PUT',
        headers,
        body: JSON.stringify({ message: 'Índice actualizado', sha: d.sha, content: Buffer.from(JSON.stringify(index, null, 2)).toString('base64') }),
      });
    }

    res.status(200).json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}
