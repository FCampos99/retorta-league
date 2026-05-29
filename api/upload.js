export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).end();

  const { filename, base64, date } = req.body;
  const token = process.env.CAPAS_TOKEN;
  if (!token) return res.status(500).json({ error: 'Token não configurado no servidor' });

  const api = 'https://api.github.com/repos/FCampos99/retorta-league/contents/capas';
  const headers = {
    Authorization: `token ${token}`,
    Accept: 'application/vnd.github.v3+json',
    'Content-Type': 'application/json',
  };

  try {
    // 1. Upload da imagem
    const imgRes = await fetch(`${api}/${filename}`, {
      method: 'PUT',
      headers,
      body: JSON.stringify({ message: `Capa: ${filename}`, content: base64 }),
    });
    if (!imgRes.ok) throw new Error((await imgRes.json()).message);

    // 2. Actualizar index.json
    let index = [], sha = null;
    const idxRes = await fetch(`${api}/index.json`, { headers });
    if (idxRes.ok) {
      const d = await idxRes.json();
      sha = d.sha;
      index = JSON.parse(Buffer.from(d.content, 'base64').toString());
    }
    index.push({ filename, date, caption: '' });
    const body = { message: 'Índice actualizado', content: Buffer.from(JSON.stringify(index, null, 2)).toString('base64') };
    if (sha) body.sha = sha;
    await fetch(`${api}/index.json`, { method: 'PUT', headers, body: JSON.stringify(body) });

    res.status(200).json({ ok: true, filename });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}
