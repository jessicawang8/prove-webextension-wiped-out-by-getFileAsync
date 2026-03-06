const https = require('https');
const fs = require('fs');
const path = require('path');
const os = require('os');

const certDir = path.join(os.homedir(), '.office-addin-dev-certs');
const options = {
	key: fs.readFileSync(path.join(certDir, 'localhost.key')),
	cert: fs.readFileSync(path.join(certDir, 'localhost.crt'))
};

let pendingDownload = null;

const MIME_TYPES = {
	'.html': 'text/html',
	'.js': 'application/javascript',
	'.css': 'text/css',
	'.png': 'image/png',
	'.json': 'application/json'
};

const server = https.createServer(options, (req, res) => {
	res.setHeader('Access-Control-Allow-Origin', '*');
	res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
	res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

	if (req.method === 'OPTIONS') {
		res.writeHead(200);
		res.end();
		return;
	}

	if (req.method === 'POST' && req.url === '/upload') {
		let body = '';
		req.on('data', chunk => { body += chunk; });
		req.on('end', () => {
			pendingDownload = body;
			res.writeHead(200, { 'Content-Type': 'application/json' });
			res.end(JSON.stringify({ ok: true }));
		});
		return;
	}

	if (req.method === 'GET' && req.url === '/download') {
		if (!pendingDownload) {
			res.writeHead(404, { 'Content-Type': 'text/plain' });
			res.end('No file available');
			return;
		}
		const buffer = Buffer.from(pendingDownload, 'base64');
		res.writeHead(200, {
			'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			'Content-Disposition': 'attachment; filename="test.docx"',
			'Content-Length': buffer.length
		});
		res.end(buffer);
		pendingDownload = null;
		return;
	}

	// Serve static files
	let filePath = req.url === '/' ? '/index.html' : req.url.split('?')[0];
	const fullPath = path.join(__dirname, filePath);

	fs.readFile(fullPath, (err, data) => {
		if (err) {
			res.writeHead(404, { 'Content-Type': 'text/plain' });
			res.end('Not found');
			return;
		}
		const ext = path.extname(filePath);
		res.writeHead(200, {
			'Content-Type': MIME_TYPES[ext] || 'application/octet-stream',
			'Cache-Control': 'no-cache'
		});
		res.end(data);
	});
});

server.listen(3008, () => {
	console.log('Server running at https://localhost:3008');
});
