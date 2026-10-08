import http from 'node:http';
import fs from 'node:fs/promises';
import path from 'node:path';

const root = process.cwd();
const port = 4173;
const server=http.createServer(async (req,res)=>{try{const pathname=new URL(req.url,'http://127.0.0.1').pathname;const files={'/':'index.html','/app_local.html':'app_local.html','/pid-logo.png':'pid-logo.png'};if(files[pathname]){const ext=path.extname(files[pathname]);res.writeHead(200,{'Content-Type':ext==='.png'?'image/png':'text/html; charset=utf-8','Cache-Control':'no-cache'});return res.end(await fs.readFile(path.join(root,files[pathname])))}res.writeHead(404);res.end('Not found')}catch(e){res.writeHead(500);res.end(e.message)}});
server.listen(port,'127.0.0.1',()=>console.log(`PID local app: http://127.0.0.1:${port}/app_local.html`));






