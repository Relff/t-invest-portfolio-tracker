/**
 * proxy.js — личный прокси для подключения к T-Invest API в обход
 * проблемы с сертификатами НУЦ Минцифры (см. PROXY-SETUP.md рядом).
 *
 * ВАЖНО, чем это отличается от "быстрого" решения через
 * rejectUnauthorized: false:
 *
 * Здесь проверка SSL-сертификата остаётся ПОЛНОСТЬЮ ВКЛЮЧЕННОЙ.
 * Мы не отключаем защиту — мы просто РАСШИРЯЕМ список удостоверяющих
 * центров, которым доверяем, добавив в него ещё и Минцифру, рядом со
 * всеми обычными мировыми CA.
 *
 * v3: roots.pem может быть склеен неидеально (например, через Windows
 * `copy /b`, где между файлами не всегда остаётся перенос строки) —
 * из-за этого Node.js может неправильно разобрать несколько
 * сертификатов в одном файле. Поэтому сами извлекаем КАЖДЫЙ отдельный
 * сертификат по маркерам BEGIN/END CERTIFICATE регуляркой, и передаём
 * их в https.request МАССИВОМ отдельных сертификатов, а не одним
 * куском текста — так не важно, есть ли между ними лишний перенос
 * строки или нет.
 *
 * Это ЛИЧНЫЙ прокси — разворачивается на СВОЁМ аккаунте Vercel, не на
 * чужом.
 */

const https = require('https');
const fs = require('fs');
const path = require('path');

function extractCertificates(pemContent) {
  const text = pemContent.toString('utf8');
  const certRegex = /-----BEGIN CERTIFICATE-----[\s\S]+?-----END CERTIFICATE-----/g;
  return text.match(certRegex) || [];
}

let certs = [];
let loadInfo = '';
const candidatePaths = [
  path.join(__dirname, 'roots.pem'),
  path.join(__dirname, '..', 'roots.pem'),
  path.join(process.cwd(), 'roots.pem'),
  '/var/task/roots.pem',
  '/var/task/api/roots.pem',
];
for (const p of candidatePaths) {
  try {
    const raw = fs.readFileSync(p);
    certs = extractCertificates(raw);
    loadInfo = 'Найден файл: ' + p + '. Извлечено сертификатов: ' + certs.length + '. Размер файла: ' + raw.length + ' байт.';
    break;
  } catch (e) {
    loadInfo = 'Не нашёл файл ни по одному из путей: ' + candidatePaths.join(', ');
  }
}

module.exports = function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Authorization, Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  // Диагностический режим — открой прокси в браузере с ?debug=1,
  // чтобы увидеть, сколько сертификатов реально загрузилось, не делая
  // запрос к самому Т-Банку.
  if (req.query.debug) {
    return res.status(200).json({
      loadInfo: loadInfo,
      certCount: certs.length,
      certPreviews: certs.map(function(c) { return c.substring(0, 60) + '...'; }),
    });
  }

  if (!certs.length) {
    return res.status(500).json({ error: 'Сертификаты не загружены. ' + loadInfo });
  }

  const { path: apiPath } = req.query;
  if (!apiPath) return res.status(400).json({ error: 'Missing path parameter' });

  const payload = req.method !== 'GET' && req.method !== 'HEAD' && req.body
    ? JSON.stringify(req.body)
    : '';

  const options = {
    hostname: 'invest-public-api.tbank.ru',
    port: 443,
    path: '/rest/' + apiPath.replace(/^\/+/, ''), // на случай, если apiPath уже начинается со слэша
    method: req.method,
    ca: certs, // массив отдельных сертификатов, а не один склеенный блок
    headers: {
      'Authorization': req.headers['authorization'] || '',
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'Content-Length': Buffer.byteLength(payload),
    },
  };

  const proxyReq = https.request(options, (proxyRes) => {
    let data = '';
    proxyRes.on('data', (chunk) => { data += chunk; });
    proxyRes.on('end', () => {
      res.setHeader('Content-Type', 'application/json');
      res.status(proxyRes.statusCode).send(data);
    });
  });

  proxyReq.on('error', (err) => {
    res.status(500).json({ error: err.message, loadInfo: loadInfo });
  });

  if (payload) proxyReq.write(payload);
  proxyReq.end();
}
