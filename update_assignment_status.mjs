import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const file='outputs/personas_servicios_pid.xlsx'; const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(file)); const sh=wb.worksheets.getItem('Asignaciones');
sh.getRange('E2:E500').dataValidation={rule:{type:'list',values:['Validado','No Validado']}};
const vals=sh.getUsedRange().values; for(let i=1;i<vals.length;i++){if(vals[i][0]&&!['Validado','No Validado'].includes(vals[i][4])) sh.getCell(i,4).values=[['Validado']];}
wb.recalculate(); const out=await SpreadsheetFile.exportXlsx(wb); await out.save(file); console.log('Estado actualizado');
