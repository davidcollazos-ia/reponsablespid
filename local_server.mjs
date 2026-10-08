import http from 'node:http';
import fs from 'node:fs/promises';
import path from 'node:path';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const root = process.cwd();
const workbookPath = path.join(root, 'outputs', 'personas_servicios_pid.xlsx');
const port = 4173;
const json = (res, status, body) => { res.writeHead(status, {'Content-Type':'application/json; charset=utf-8','Access-Control-Allow-Origin':'*'}); res.end(JSON.stringify(body)); };
async function readState(){
  const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(workbookPath));
  const peopleSheet = wb.worksheets.getItem('Personas');
  const serviceSheet = wb.worksheets.getItem('Servicios');
  const assignmentSheet = wb.worksheets.getItem('Asignaciones');
  const p = peopleSheet.getUsedRange().values.slice(1).filter(r=>r[0]).map(r=>({id:r[0],name:r[1]||'',last:r[2]||'',email:r[3]||'',type:r[4]||'SEGITTUR',area:r[5]||'',active:r[6]||'Sí',notes:r[7]||''}));
  const s = serviceSheet.getUsedRange().values.slice(1).filter(r=>r[0]).map(r=>({id:r[0],code:r[1]||'',name:r[2]||'',desc:r[3]||'',urlpre:r[4]||'',urlpro:'',active:r[5]||'Activo'}));
  const a = assignmentSheet.getUsedRange().values.slice(1).filter(r=>r[0]&&(r[1]||r[2])).map(r=>({id:r[0],personId:r[1]||'',serviceId:r[2]||'',role:r[3]||'Principal',active:r[4]||'Validado',notes:r[5]||''}));
  return {people:p.filter(x=>x.name||x.last||x.email),services:s,assignments:a};
}
async function saveState(body){
  const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(workbookPath));
  const ps = wb.worksheets.getItem('Personas'); const ss = wb.worksheets.getItem('Servicios'); const as = wb.worksheets.getItem('Asignaciones');
  const people = body.people||[]; const services = body.services||[]; const assignments = body.assignments||[];
  ps.getRange('A2:H500').clear({applyTo:'contents'}); ss.getRange('A2:F500').clear({applyTo:'contents'}); as.getRange('A2:F500').clear({applyTo:'contents'});
  if(people.length) ps.getRange(`A2:H${people.length+1}`).values=people.map(p=>[p.id,p.name,p.last,p.email,p.type,p.area,p.active,p.notes]);
  if(services.length) ss.getRange(`A2:E${services.length+1}`).values=services.map(s=>[s.id,s.name,s.desc,s.urlpre||'',s.active]);
  if(assignments.length) as.getRange(`A2:H${assignments.length+1}`).values=assignments.map(a=>[a.id,a.personId,a.serviceId,a.role,a.start||'',a.end||'',a.active,a.notes||'']);
  wb.recalculate(); const out=await SpreadsheetFile.exportXlsx(wb); await out.save(workbookPath); return {saved:true,path:workbookPath,people:people.length,services:services.length,assignments:assignments.length};
}
const server=http.createServer(async (req,res)=>{try{if(req.method==='GET'&&req.url==='/pid-logo.png'){res.writeHead(200,{'Content-Type':'image/png','Cache-Control':'no-cache'});return res.end(await fs.readFile(path.join(root,'pid-logo.png')));}if(req.method==='GET'&&req.url==='/api/state')return json(res,200,await readState());if(req.method==='POST'&&req.url==='/api/save'){let raw='';for await(const c of req)raw+=c;return json(res,200,await saveState(JSON.parse(raw)));}if(req.method==='GET'&&(req.url==='/'||req.url==='/app_local.html')){res.writeHead(200,{'Content-Type':'text/html; charset=utf-8'});return res.end(await fs.readFile(path.join(root,'app_local.html')));}res.writeHead(404);res.end('Not found');}catch(e){json(res,500,{error:e.message});}});
server.listen(port,'127.0.0.1',()=>console.log(`PID local app: http://127.0.0.1:${port}/app_local.html`));






