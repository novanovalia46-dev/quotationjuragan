/**
 * Proxy Vercel -> Google Apps Script.
 * Hindari masalah CORS saat frontend memanggil script.google.com.
 *
 * Set environment variable di Vercel:
 *   GAS_WEB_APP_URL = https://script.google.com/macros/s/XXXX/exec
 */
module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');

  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    return res.status(204).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, message: 'Gunakan metode POST.' });
  }

  var gasUrl = process.env.GAS_WEB_APP_URL;
  if (!gasUrl) {
    return res.status(500).json({
      success: false,
      message: 'GAS_WEB_APP_URL belum di-set. Isi di Vercel Project Settings > Environment Variables.'
    });
  }

  try {
    var payload = typeof req.body === 'string' ? req.body : JSON.stringify(req.body || {});
    var response = await fetch(gasUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: payload,
      redirect: 'follow'
    });
    var text = await response.text();
    try {
      var json = JSON.parse(text);
      return res.status(200).json(json);
    } catch (err) {
      return res.status(502).json({
        success: false,
        message: 'Respons Apps Script bukan JSON. Deploy ulang web app: Execute as Me, Who has access = Anyone, lalu salin URL /exec terbaru.'
      });
    }
  } catch (err) {
    return res.status(502).json({
      success: false,
      message: 'Gagal menghubungi Apps Script: ' + (err && err.message ? err.message : 'unknown')
    });
  }
};
