import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const file='outputs/personas_servicios_pid.xlsx';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(file));
const people=wb.worksheets.getItem('Personas'); const cat=wb.worksheets.getItem('Catalogos');
cat.getRange('B2:B5').values=[['RRII'],['Desarrollo de Negocio'],['IDI'],['Comunicación']];
people.getRange('F2:F500').dataValidation={rule:{type:'list',values:['RRII','Desarrollo de Negocio','IDI','Comunicación']}};
wb.recalculate(); const out=await SpreadsheetFile.exportXlsx(wb); await out.save(file); console.log('Áreas actualizadas');
