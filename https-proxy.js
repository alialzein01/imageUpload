const https = require('https');
const http = require('http');
const fs = require('fs');
const { createProxyMiddleware } = require('http-proxy-middleware');

// Increase max listeners to avoid warning
process.setMaxListeners(15);

// Load SSL certificates
const options = {
  key: fs.readFileSync('/Users/ali/.office-addin-dev-certs/localhost.key'),
  cert: fs.readFileSync('/Users/ali/.office-addin-dev-certs/localhost.crt')
};

// Create HTTPS server
const server = https.createServer(options, (req, res) => {
  // Set CORS headers
  res.setHeader('Access-Control-Allow-Origin', 'https://localhost:3000');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  res.setHeader('Access-Control-Allow-Credentials', 'true');

  // Handle preflight requests
  if (req.method === 'OPTIONS') {
    res.writeHead(200);
    res.end();
    return;
  }

  // Proxy to Django backend
  const proxy = createProxyMiddleware({
    target: 'http://localhost:8001',
    changeOrigin: true,
    pathRewrite: {
      '^/api': '/api'
    }
  });

  proxy(req, res);
});

const PORT = 8000;
server.listen(PORT, '0.0.0.0', () => {
  console.log(`HTTPS proxy server running on https://localhost:${PORT}`);
  console.log('Proxying requests to Django backend on http://localhost:8001');
});
