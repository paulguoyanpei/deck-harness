// proxy-logger.js —— 记录 Claude Code <-> DeepSeek 之间所有 raw request/response
//
// 用法:
//   npm install
//   node logger.js
//   set -x ANTHROPIC_BASE_URL "http://localhost:8888"   # 注意：不要带 /v1
//   set -x ANTHROPIC_AUTH_TOKEN "<DeepSeek API Key>"
//   claude
const http = require('http');
const httpProxy = require('http-proxy');
const zlib = require('zlib');
const fs = require('fs');
const path = require('path');
const { PassThrough } = require('stream');

// DeepSeek 的 Anthropic 兼容端点。少了 /anthropic 会打到 OpenAI 格式的路径上，404。
const TARGET = process.env.UPSTREAM_BASE_URL || 'https://api.deepseek.com/anthropic';
const PORT = Number(process.env.PORT || 8888);

// 默认落在脚本同目录，不受启动时 cwd 影响；要改位置就设 LOG_DIR
const LOG_DIR = process.env.LOG_DIR || __dirname;
const REQ_LOG_PATH = path.join(LOG_DIR, 'api-requests.jsonl');
const RES_LOG_PATH = path.join(LOG_DIR, 'api-responses.jsonl');
fs.mkdirSync(LOG_DIR, { recursive: true });

const proxy = httpProxy.createProxyServer({
  target: TARGET,
  changeOrigin: true,
  secure: true, // 不要关：API key 就是从这条连接发出去的
});

// 用 WriteStream 而不是 appendFileSync，避免阻塞事件循环
const reqLog = fs.createWriteStream(REQ_LOG_PATH, { flags: 'a' });
const resLog = fs.createWriteStream(RES_LOG_PATH, { flags: 'a' });

// 别把凭证写进日志文件
const SECRETS = ['x-api-key', 'authorization', 'proxy-authorization'];
function redact(headers) {
  const out = { ...headers };
  for (const h of SECRETS) if (h in out) out[h] = '<redacted>';
  return out;
}

// 上游默认会 gzip，直接 toString() 得到的是二进制乱码
function decode(buf, encoding) {
  try {
    switch ((encoding || '').toLowerCase()) {
      case 'gzip': return zlib.gunzipSync(buf).toString('utf8');
      case 'br': return zlib.brotliDecompressSync(buf).toString('utf8');
      case 'deflate': return zlib.inflateSync(buf).toString('utf8');
      default: return buf.toString('utf8');
    }
  } catch (err) {
    return `<decode failed: ${err.message}; base64=${buf.toString('base64')}>`;
  }
}

let seq = 0;

// 记录响应。proxyRes 事件在 proxyRes.pipe(res) 之前同步触发，所以这里挂
// data 监听是安全的，不会漏包。
proxy.on('proxyRes', (proxyRes, req) => {
  const chunks = [];
  proxyRes.on('data', (chunk) => chunks.push(chunk));
  proxyRes.on('end', () => {
    const body = decode(Buffer.concat(chunks), proxyRes.headers['content-encoding']);
    resLog.write(JSON.stringify({
      id: req.__id,
      timestamp: new Date().toISOString(),
      durationMs: Date.now() - req.__start,
      statusCode: proxyRes.statusCode,
      headers: proxyRes.headers,
      body,
    }) + '\n');
    console.log(`📥 #${req.__id} ${proxyRes.statusCode} (${Date.now() - req.__start}ms, ${body.length}B)`);
  });
});

// 上游连不上时如果没有这个 handler，http-proxy 会 throw，整个进程直接挂掉
proxy.on('error', (err, req, res) => {
  console.error(`❌ #${req.__id} proxy error:`, err.message);
  if (res && !res.headersSent) {
    res.writeHead(502, { 'content-type': 'application/json' });
  }
  if (res && !res.writableEnded) {
    res.end(JSON.stringify({
      type: 'error',
      error: { type: 'api_error', message: `proxy: ${err.message}` },
    }));
  }
});

const server = http.createServer((req, res) => {
  req.__id = ++seq;
  req.__start = Date.now();

  // 自己读 body 并 tee 一份给 http-proxy（options.buffer），这样读取顺序完全可控，
  // 不用赌 proxyReq 事件和 req.pipe() 谁先跑
  const chunks = [];
  const buffered = new PassThrough();
  req.on('data', (chunk) => {
    chunks.push(chunk);
    buffered.write(chunk);
  });
  req.on('end', () => {
    buffered.end();
    const body = Buffer.concat(chunks).toString('utf8');
    reqLog.write(JSON.stringify({
      id: req.__id,
      timestamp: new Date().toISOString(),
      method: req.method,
      url: req.url,
      headers: redact(req.headers),
      body,
    }) + '\n');
    console.log(`📤 #${req.__id} ${req.method} ${req.url} (${body.length}B)`);
  });
  req.on('error', (err) => {
    console.error(`❌ #${req.__id} client error:`, err.message);
    buffered.destroy(err);
  });

  proxy.web(req, res, { buffer: buffered });
});

// 流式响应可以跑很久，Node 18 默认 requestTimeout 是 300s，会把长请求掐断
server.requestTimeout = 0;
server.headersTimeout = 0;

server.listen(PORT, () => {
  console.log(`🔍 proxy listening on http://localhost:${PORT}`);
  console.log(`   -> ${TARGET}`);
  console.log(`   req log: ${REQ_LOG_PATH}`);
  console.log(`   res log: ${RES_LOG_PATH}`);
});
