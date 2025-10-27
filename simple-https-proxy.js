const https = require('https');
const http = require('http');
const fs = require('fs');

const HTTPS_PORT = 8000;
const HTTP_TARGET_PORT = 8001;

// Load SSL certificates
const options = {
  key: fs.readFileSync('/Users/ali/.office-addin-dev-certs/localhost.key'),
  cert: fs.readFileSync('/Users/ali/.office-addin-dev-certs/localhost.crt'),
  // Add TLS options for better compatibility
  secureProtocol: 'TLSv1_2_method',
  ciphers: 'ECDHE-RSA-AES128-GCM-SHA256:ECDHE-RSA-AES256-GCM-SHA384:ECDHE-RSA-AES128-SHA256:ECDHE-RSA-AES256-SHA384'
};

const server = https.createServer(options, (req, res) => {
  console.log(`HTTPS Proxy: ${req.method} ${req.url}`);

  // Add CORS headers
  res.setHeader('Access-Control-Allow-Origin', 'https://localhost:3000');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  res.setHeader('Access-Control-Allow-Credentials', 'true');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  // Create proxy request to Django
  const proxyReq = http.request({
    hostname: 'localhost',
    port: HTTP_TARGET_PORT,
    path: req.url,
    method: req.method,
    headers: {
      ...req.headers,
      'Host': 'localhost:8001',
      'X-Forwarded-Proto': 'https'
    }
  }, (proxyRes) => {
    console.log(`Proxy Response: ${proxyRes.statusCode}`);
    
    // Copy response headers
    Object.keys(proxyRes.headers).forEach(key => {
      res.setHeader(key, proxyRes.headers[key]);
    });
    
    res.writeHead(proxyRes.statusCode);
    proxyRes.pipe(res);
  });

  // Handle proxy errors
  proxyReq.on('error', (err) => {
    console.error('Proxy request error:', err);
    res.writeHead(500, { 'Content-Type': 'text/plain' });
    res.end('Proxy error: ' + err.message);
  });

  // Pipe request body to proxy
  req.pipe(proxyReq);
});

server.listen(HTTPS_PORT, () => {
  console.log(`Simple HTTPS proxy listening on https://localhost:${HTTPS_PORT}`);
  console.log(`Forwarding requests to http://localhost:${HTTP_TARGET_PORT}`);
});

server.on('error', (err) => {
  console.error('HTTPS server error:', err);
});
