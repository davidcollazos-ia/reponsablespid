import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const file='outputs/personas_servicios_pid.xlsx'; const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(file)); const sh=wb.worksheets.getItem('Asignaciones');
const old=sh.getUsedRange().values; const rows=old.slice(1).filter(r=>r[0]&&(r[1]||r[2])).map(r=>[r[0],r[1]||'',r[2]||'',r[3]||'Principal',r[6]||r[4]||'Sí',r[7]||r[5]||'']);
sh.getRange('A1:H500').clear({applyTo:'contents'}); sh.getRange(`A1:F${Math.max(rows.length+1,2)}`).values=[['AsignacionID','PersonaID','ServicioID','Rol','Activo','Notas'],...(rows.length?rows:[['ASI-0001','','','Principal','Sí','']])]; wb.recalculate(); const out=await SpreadsheetFile.exportXlsx(wb); await out.save(file); console.log('Asignaciones simplificadas');
