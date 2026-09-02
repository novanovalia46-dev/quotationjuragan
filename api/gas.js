/**
 * Proxy Vercel -> Google Apps Script.
 * Set sekali di Vercel: Environment Variable
 *   Key:   GAS_WEB_APP_URL
 *   Value: https://script.google.com/macros/s/XXXX/exec
 * lalu Redeploy. Semua akun memakai URL itu, tanpa kotak kuning.
 */
var GAS_WEB_APP_URL = ''; // opsional hardcode: 'https://script.google.com/macros/s/XXXX/exec'

function isValidGasUrl(url) {
  url = String(url || '').trim();
  return /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec\/?$/.test(url)
    || /^https:\/\/script\.google\.com\/a\/macros\/[^/]+\/s\/[A-Za-z0-9_-]+\/exec\/?$/.test(url);
}

function serverGasUrl() {
  return String(process.env.GAS_WEB_APP_URL || GAS_WEB_APP_URL || '').trim();
}

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');

  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    return res.status(204).end();
  }

  if (req.method === 'GET') {
    return res.status(200).json({
      success: true,
      configured: isValidGasUrl(serverGasUrl())
    });
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, message: 'Gunakan metode POST.' });
  }

  var body = req.body;
  if (typeof body === 'string') {
    try { body = JSON.parse(body); } catch (e) { body = {}; }
  }
  body = body || {};

  var gasUrl = serverGasUrl() || body.gasUrl || '';
  gasUrl = String(gasUrl).trim();

  if (!gasUrl) {
    return res.status(500).json({
      success: false,
      code: 'NO_GAS_URL',
      message: 'URL Apps Script belum di-set di Vercel (Environment Variable GAS_WEB_APP_URL).'
    });
  }
  if (!isValidGasUrl(gasUrl)) {
    return res.status(400).json({
      success: false,
      code: 'BAD_GAS_URL',
      message: 'URL tidak valid. Harus https://script.google.com/macros/s/.../exec'
    });
  }

  try {
    var response = await fetch(gasUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body),
      redirect: 'follow'
    });
    var text = await response.text();
    try {
      return res.status(200).json(JSON.parse(text));
    } catch (err) {
      return res.status(502).json({
        success: false,
        message: 'Respons Apps Script bukan JSON. Deploy Web app: Execute as Me, Who has access = Anyone.'
      });
    }
  } catch (err) {
    return res.status(502).json({
      success: false,
      message: 'Gagal menghubungi Apps Script: ' + (err && err.message ? err.message : 'unknown')
    });
  }
};
