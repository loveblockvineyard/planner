/* Vercel serverless function — POST /api/smartsheet
   Proxies add-rows calls to Smartsheet so the API token
   never ships to the browser.

   Setup:
     Vercel Dashboard → Project → Settings → Environment Variables
     Add:  SMARTSHEET_TOKEN = <your Smartsheet API token>
     (Smartsheet → Account → Personal Settings → API Access → Generate Token)
*/
export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const token = process.env.SMARTSHEET_TOKEN;
  if (!token) return res.status(500).json({ error: 'SMARTSHEET_TOKEN env var not set' });

  const { sheetId, rows } = req.body || {};
  if (!sheetId || !Array.isArray(rows) || !rows.length)
    return res.status(400).json({ error: 'Missing sheetId or rows' });

  try {
    const ssRes = await fetch(
      `https://api.smartsheet.com/2.0/sheets/${sheetId}/rows`,
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(rows),
      }
    );
    const data = await ssRes.json();
    if (!ssRes.ok) return res.status(ssRes.status).json(data);
    res.status(200).json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
